import {
  getGraphToken,
  queryListItems,
  createListItem,
  updateListItemFields,
  uploadFileToDrive,
  createDocLibrary,
  listExistsGraph,
} from "./_utils/graphClient.js";

interface UploadedFile {
  name: string;
  content: string;
  contentType: string;
}

interface JobApplyBody {
  jobListingId: string;
  jobTitle: string;
  applicantName: string;
  applicantEmail: string;
  applicantPhone: string;
  currentPosition?: string;
  currentDepartment?: string;
  coverLetter?: string;
  files?: UploadedFile[];
  customAnswers?: Record<string, unknown>;
  submittedByEmail?: string;
  forceApply?: boolean;
}

interface ApiRequest {
  body: Record<string, unknown>;
  method: string;
}

interface ApiResponse {
  status(code: number): ApiResponse;
  json(data: Record<string, unknown>): void;
  setHeader(name: string, value: string): void;
  end(): void;
}

/** Simple HTML entity encoder — prevents XSS in email HTML */
function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function generateSubmissionRef(): string {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `JOB-${y}${m}${d}-${rand}`;
}

function decodeBase64(content: string): Uint8Array {
  let b64 = content;
  if (b64.startsWith("data:")) {
    const commaIdx = b64.indexOf(",");
    b64 = commaIdx >= 0 ? b64.substring(commaIdx + 1) : b64;
  }
  return new Uint8Array(Buffer.from(b64, "base64"));
}

export default async function handler(req: ApiRequest, res: ApiResponse) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const body = req.body as unknown as JobApplyBody & { accessToken?: string };
  const {
    jobListingId,
    jobTitle,
    applicantName,
    applicantEmail,
    applicantPhone,
    coverLetter,
    files,
    customAnswers,
  } = body;
  const userToken = body.accessToken || "";

  if (!jobListingId || !jobTitle || !applicantName || !applicantEmail || !applicantPhone) {
    return res.status(400).json({
      error: "Missing required fields: jobListingId, jobTitle, applicantName, applicantEmail, applicantPhone",
    });
  }

  try {
    const sysToken = await getGraphToken();

    // Duplicate check: block if same email already applied for same job,
    // unless the submitter is someone else or forceApply is true (admin test)
    if (!(body as Record<string, unknown>).forceApply) {
      try {
        const existing = await queryListItems(sysToken, "Job Applications", {
          filter: `fields/ApplicantEmail eq '${encodeURIComponent(applicantEmail)}' and fields/JobListingID eq ${Number(jobListingId)}`,
          top: 1,
        });
        if (existing.length > 0) {
          const submitterEmail = (body as Record<string, unknown>).submittedByEmail as string || "";
          if (submitterEmail.toLowerCase() === applicantEmail.toLowerCase()) {
            return res.status(409).json({
              error: "You have already applied for this position. Multiple applications are not allowed.",
            });
          }
        }
      } catch {
        // If duplicate check fails, proceed anyway
      }
    }

    const uploadToken = userToken || sysToken; // prefer user token for file ops
    const submissionRef = generateSubmissionRef();
    const submittedAt = new Date().toISOString();

    // Upload all documents to document library
    const docLibName = "Job Applications Files";
    let docLibReady = false;
    const allDocs: { name: string; url: string; isCoverLetter: boolean }[] = [];
    let resumeUrl = "";

    async function ensureDocLib(): Promise<void> {
      if (!docLibReady) {
        if (!(await listExistsGraph(uploadToken, docLibName))) {
          await createDocLibrary(uploadToken, docLibName);
        }
        docLibReady = true;
      }
    }

    async function uploadDoc(name: string, content: string, isCoverLetter: boolean): Promise<string | null> {
      try {
        await ensureDocLib();
        const safeName = name.replace(/[^a-zA-Z0-9._-]/g, "_");
        const uniqueName = `${submissionRef}_${safeName}`;
        const binary = decodeBase64(content);
        const fileUrl = await uploadFileToDrive(uploadToken, docLibName, uniqueName, binary);
        allDocs.push({ name, url: fileUrl, isCoverLetter });
        return fileUrl;
      } catch (e) {
        console.error("[API job-apply] Upload failed:", (e as Error).message);
        return null;
      }
    }

    // Cover letter is stored directly as text column — no file upload needed

    // Upload resume and additional files
    if (files && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        if (!file.name || !file.content) continue;
        const url = await uploadDoc(file.name, file.content, false);
        if (i === 0 && url) resumeUrl = url;
      }
    }

    // Build application fields — using exact SP column names
    const applicationFields: Record<string, unknown> = {
      Title: `${jobTitle} - ${applicantName}`,
      JobListingID: Number(jobListingId),
      ApplicantName: applicantName,
      ApplicantEmail: applicantEmail,
      ApplicantPhone: applicantPhone,
      Status: "New",
      SubmittedAt: submittedAt,
      SubmissionRef: submissionRef,
      ResumeUrl: resumeUrl,
    };

    // Reasoning stored as long text column (no file upload)
    if (coverLetter && coverLetter.trim()) {
      applicationFields.Reasoning = coverLetter;
    }

    // Only include CustomAnswers if there's data — column may not exist on the list
    const hasCustomAnswers = customAnswers && typeof customAnswers === "object" && Object.keys(customAnswers).length > 0;
    if (hasCustomAnswers) {
      applicationFields.CustomAnswers = JSON.stringify(customAnswers);
    }

    // Create the item with Title plus essential text fields in one call,
    // then attempt optional fields individually so missing columns don't block
    const coreFields: Record<string, unknown> = {
      Title: applicationFields.Title,
      Status: "New",
      SubmissionRef: applicationFields.SubmissionRef,
    };
    // Add applicant info fields individually (may not exist on list)
    for (const k of ["ApplicantName", "ApplicantEmail", "ApplicantPhone", "CurrentPosition", "CurrentDepartment", "JobListingID"]) {
      if (applicationFields[k] != null && applicationFields[k] !== "") {
        coreFields[k] = applicationFields[k];
      }
    }

    const created = await createListItem(sysToken, "Job Applications", coreFields);
    const itemId = created.id;

    // Attempt optional fields — silently skip any that don't exist or have type issues
    const optionalFields = ["SubmittedAt", "ResumeUrl", "Reasoning", "CustomAnswers"];
    for (const key of optionalFields) {
      const value = applicationFields[key];
      if (value === "" || value == null) continue;
      try {
        await updateListItemFields(sysToken, "Job Applications", itemId, { [key]: value });
      } catch {
        // Column may not exist — application still goes through with core fields
      }
    }

    // Increment Application_x0020_Count on Internal Job Listing
    try {
      const jobItems = await queryListItems(sysToken, "Internal Job Listing", { top: 1000 });
      const jobItem = jobItems.find((item) => String(item.id) === String(jobListingId));
      const currentCount = Number(jobItem?.fields?.Application_x0020_Count) || 0;
      await updateListItemFields(sysToken, "Internal Job Listing", String(jobListingId), {
        Application_x0020_Count: currentCount + 1,
      });
    } catch (e) {
      console.error("[API job-apply] Failed to increment count:", (e as Error).message);
    }

    // Send email to HR recruitment
    const hrEmail = process.env.HR_RECRUITMENT_EMAIL || "";
    const fromAddress = process.env.EMAIL_FROM_ADDRESS || "";

    if (hrEmail && fromAddress) {
      try {
        const docListHtml = allDocs.length > 0
          ? `<p><strong>Documents:</strong></p><ul>${allDocs.map((d) =>
              `<li><a href="${d.url}">${escapeHtml(d.name)}${d.isCoverLetter ? " (Cover Letter)" : ""}</a></li>`
            ).join("")}</ul>`
          : "";

        const customHtml = customAnswers && Object.keys(customAnswers).length > 0
          ? `<p><strong>Additional Responses:</strong></p><table style="border-collapse:collapse;width:100%;max-width:600px;margin-bottom:16px">${
              Object.entries(customAnswers).map(([k, v]) =>
                `<tr style="border:1px solid #d1d5db"><td style="padding:8px;border:1px solid #d1d5db;background:#f3f4f6;font-weight:600;width:30%">${escapeHtml(k)}</td><td style="padding:8px;border:1px solid #d1d5db">${escapeHtml(String(v ?? ""))}</td></tr>`
              ).join("")
            }</table>`
          : "";

        const submittedAtFormatted = new Date(submittedAt).toLocaleString("en-MY", {
          year: "numeric", month: "long", day: "numeric", hour: "2-digit", minute: "2-digit", timeZoneName: "short",
        });

        const reasoningPreview = coverLetter && coverLetter.trim()
          ? `<p><strong>Reasoning:</strong></p><blockquote style="background:#f5f5f5;padding:12px;border-left:4px solid #0078D4;margin:0 0 12px 0;font-size:13px;line-height:1.6;">${escapeHtml(coverLetter).replace(/\n/g, "<br>")}</blockquote>`
          : "";

        const eh = (s: string) => escapeHtml(s);
        const htmlBody = `
          <!DOCTYPE html>
          <html>
          <head>
            <meta charset="UTF-8">
            <style>
              body { font-family: 'Segoe UI', Arial, sans-serif; color: #1a1a1a; font-size: 13px; line-height: 1.5; padding: 24px; }
              h2 { color: #0078D4; font-size: 20px; font-weight: 600; margin: 0 0 16px 0; padding-bottom: 8px; border-bottom: 2px solid #0078D4; }
              table { border-collapse: collapse; width: 100%; max-width: 600px; margin-bottom: 16px; }
              td { padding: 8px 12px; border: 1px solid #d1d5db; font-size: 13px; vertical-align: top; }
              td:first-child { background-color: #f3f4f6; font-weight: 600; width: 30%; white-space: nowrap; }
              a { color: #0078D4; text-decoration: none; }
              .section { margin-top: 16px; }
              .section-title { font-weight: 600; font-size: 14px; color: #0078D4; margin: 0 0 8px 0; padding-bottom: 4px; border-bottom: 1px solid #e5e7eb; }
              @media print {
                body { padding: 0; font-size: 11px; }
                td { border-color: #000; }
                h2 { border-bottom-color: #000; color: #000; }
                a { color: #000; text-decoration: underline; }
              }
            </style>
          </head>
          <body>
            <h2>New Job Application</h2>
            <table>
              <tr><td>Position</td><td>${eh(jobTitle)}</td></tr>
              <tr><td>Applicant</td><td>${eh(applicantName)}</td></tr>
              <tr><td>Email</td><td><a href="mailto:${eh(applicantEmail)}">${eh(applicantEmail)}</a></td></tr>
              <tr><td>Phone</td><td>${eh(applicantPhone)}</td></tr>
              <tr><td>Reference</td><td style="font-family:monospace;letter-spacing:0.05em;">${eh(submissionRef)}</td></tr>
              <tr><td>Date Submitted</td><td>${submittedAtFormatted}</td></tr>
            </table>
            ${customHtml ? `<div class="section"><p class="section-title">Additional Responses</p>${customHtml}</div>` : ""}
            ${docListHtml ? `<div class="section">${docListHtml}</div>` : ""}
            ${reasoningPreview ? `<div class="section">${reasoningPreview}</div>` : ""}
          </body>
          </html>
        `;

        const attachments: Array<Record<string, unknown>> = [];
        if (files && files.length > 0) {
          for (const file of files) {
            if (!file.name || !file.content) continue;
            let b64 = file.content;
            if (b64.startsWith("data:")) {
              const commaIdx = b64.indexOf(",");
              b64 = commaIdx >= 0 ? b64.substring(commaIdx + 1) : b64;
            }
            attachments.push({
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: file.name,
              contentType: file.contentType || "application/octet-stream",
              contentBytes: b64,
            });
          }
        }

        const emailPayload: Record<string, unknown> = {
          message: {
            subject: `Job Application: ${jobTitle} - ${applicantName}`,
            body: { contentType: "HTML", content: htmlBody },
            toRecipients: [{ emailAddress: { address: hrEmail } }],
            attachments,
          },
          saveToSentItems: false,
        };

        const graphRes = await fetch(
          `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromAddress)}/sendMail`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${sysToken}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(emailPayload),
          }
        );

        if (!graphRes.ok) {
          const errText = await graphRes.text();
          console.error("[API job-apply] sendMail failed:", graphRes.status, errText);
        }
      } catch (e) {
        console.error("[API job-apply] Email failed:", (e as Error).message);
      }
    }

    return res.status(200).json({
      success: true,
      applicationId: created.id,
      submissionRef,
    });
  } catch (e) {
    console.error("[API job-apply]", e);
    return res.status(500).json({ error: "Internal server error. Please try again." });
  }
}
