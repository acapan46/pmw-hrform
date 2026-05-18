import { useState, useEffect, useRef } from "react";
import { useParams, useNavigate } from "react-router-dom";
import {
  Box,
  Typography,
  TextField,
  Button,
  Paper,
  Container,
  Chip,
  CircularProgress,
  Alert,
  Divider,
  IconButton,
  Grid,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  FormHelperText,
} from "@mui/material";
import {
  ArrowBack,
  UploadFile,
  CheckCircle,
  Description,
  Close,
  LocationOn,
  Work,
} from "@mui/icons-material";
import { useReactiveForm, required, email, phone, requiredFile } from "../hooks/useReactiveForm";
import { useUserProfile } from "../hooks/useUserProfile";
import { useMsal } from "@azure/msal-react";
import { pdf } from "@react-pdf/renderer";
import { fetchJobs, submitApplication } from "../utils/careersService";
import JobApplyPdfDocument from "../utils/JobApplyPdfDocument";
import type { JobListing, CustomFieldDefinition } from "../types";

// eslint-disable-next-line @typescript-eslint/consistent-indexed-object-style
interface FormValues extends Record<string, unknown> {
  name: string;
  email: string;
  phone: string;
  currentPosition: string;
  currentDepartment: string;
  coverLetter: string;
  files: FileEntry[];
}

interface FileEntry {
  name: string;
  content: string;
  contentType: string;
}

const MAX_FILES = 5;
const ACCEPTED_TYPES = [
  "application/pdf",
  "application/msword",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "image/jpeg",
  "image/png",
];

function SuccessView({
  submissionRef,
  onBrowseMore,
}: {
  submissionRef: string;
  onBrowseMore: () => void;
}) {
  return (
    <Box sx={{ textAlign: "center", py: 6 }}>
      <Box
        sx={{
          width: 72,
          height: 72,
          borderRadius: "50%",
          backgroundColor: "#E6F4EA",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          mx: "auto",
          mb: 3,
        }}
      >
        <CheckCircle sx={{ fontSize: 40, color: "#34A853" }} />
      </Box>
      <Typography variant="h5" sx={{ fontWeight: 700, color: "#111827", mb: 1 }}>
        Application Submitted!
      </Typography>
      <Typography variant="body1" sx={{ color: "#6B7280", mb: 3 }}>
        Your application has been received successfully.
      </Typography>
      <Paper
        variant="outlined"
        sx={{
          display: "inline-flex",
          alignItems: "center",
          gap: 1.5,
          px: 3,
          py: 2,
          borderRadius: "12px",
          borderColor: "#0078D4",
          backgroundColor: "#F0F7FF",
          mb: 4,
        }}
      >
        <Typography variant="body2" sx={{ color: "#6B7280" }}>
          Reference No.
        </Typography>
        <Typography
          variant="h6"
          sx={{ fontWeight: 700, color: "#0078D4", letterSpacing: "0.05em", fontFamily: "monospace" }}
        >
          {submissionRef}
        </Typography>
      </Paper>
      <Typography variant="body2" sx={{ color: "#9CA3AF", mb: 4 }}>
        We will review your application and get back to you via email.
      </Typography>
      <Button
        variant="outlined"
        onClick={onBrowseMore}
        sx={{
          borderRadius: "12px",
          textTransform: "none",
          fontWeight: 600,
          borderColor: "#0078D4",
          color: "#0078D4",
          px: 4,
          py: 1.2,
        }}
      >
        Browse More Jobs
      </Button>
    </Box>
  );
}

function FileUploadArea({
  files,
  onAdd,
  onRemove,
}: {
  files: FileEntry[];
  onAdd: (entries: FileEntry[]) => void;
  onRemove: (index: number) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [reading, setReading] = useState(false);

  const handleSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;

    setReading(true);
    const newEntries: FileEntry[] = [];
    const remaining = MAX_FILES - files.length;

    for (let i = 0; i < Math.min(fileList.length, remaining); i++) {
      const file = fileList[i];
      try {
        const content = await readFileAsBase64(file);
        newEntries.push({ name: file.name, content, contentType: file.type });
      } catch {
        // skip files that fail to read
      }
    }

    onAdd(newEntries);
    setReading(false);
    // Reset input so same file can be selected again
    if (inputRef.current) inputRef.current.value = "";
  };

  return (
    <Box>
      <Typography variant="body2" sx={{ fontWeight: 600, color: "#374151", mb: 1 }}>
        Resume & Supporting Documents
      </Typography>
      <Typography variant="caption" sx={{ color: "#9CA3AF", display: "block", mb: 1.5 }}>
        Accepted: PDF, DOC, DOCX, JPEG, PNG (Max {MAX_FILES} files)
      </Typography>

      {/* Drop zone */}
      <Paper
        variant="outlined"
        onClick={() => inputRef.current?.click()}
        sx={{
          borderStyle: "dashed",
          borderColor: "#D1D5DB",
          borderRadius: "12px",
          p: 3,
          textAlign: "center",
          cursor: "pointer",
          transition: "all 0.2s",
          backgroundColor: "#FAFBFC",
          "&:hover": {
            borderColor: "#0078D4",
            backgroundColor: "#F0F7FF",
          },
          opacity: reading ? 0.6 : 1,
        }}
      >
        <input
          ref={inputRef}
          type="file"
          multiple
          accept={ACCEPTED_TYPES.join(",")}
          onChange={handleSelect}
          style={{ display: "none" }}
        />
        {reading ? (
          <CircularProgress size={24} sx={{ color: "#0078D4" }} />
        ) : (
          <>
            <UploadFile sx={{ fontSize: 32, color: "#9CA3AF", mb: 1 }} />
            <Typography variant="body2" sx={{ color: "#6B7280" }}>
              Click to upload or drag and drop
            </Typography>
          </>
        )}
      </Paper>

      {/* File list */}
      {files.length > 0 && (
        <Box sx={{ mt: 1.5, display: "flex", flexDirection: "column", gap: 1 }}>
          {files.map((file, i) => (
            <Paper
              key={i}
              variant="outlined"
              sx={{
                display: "flex",
                alignItems: "center",
                gap: 1,
                px: 1.5,
                py: 1,
                borderRadius: "8px",
                borderColor: "#E5E7EB",
              }}
            >
              <Description sx={{ fontSize: 18, color: "#6B7280", flexShrink: 0 }} />
              <Typography variant="body2" sx={{ color: "#374151", flex: 1, fontSize: "0.8rem" }} noWrap>
                {file.name}
              </Typography>
              <IconButton size="small" onClick={() => onRemove(i)} sx={{ color: "#9CA3AF" }}>
                <Close sx={{ fontSize: 16 }} />
              </IconButton>
            </Paper>
          ))}
        </Box>
      )}
    </Box>
  );
}

function readFileAsBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      // Strip the data:...;base64, prefix if present
      const base64 = result.includes("base64,") ? result.split("base64,")[1] : result;
      resolve(base64);
    };
    reader.onerror = () => reject(new Error(`Failed to read ${file.name}`));
    reader.readAsDataURL(file);
  });
}

export default function JobApplyPage() {
  const { jobId } = useParams<{ jobId: string }>();
  const navigate = useNavigate();
  const profile = useUserProfile();
  const { instance, accounts } = useMsal();

  const [job, setJob] = useState<JobListing | null>(null);
  const [jobLoading, setJobLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | null>(null);
  const [submitted, setSubmitted] = useState(false);
  const [submissionRef, setSubmissionRef] = useState("");
  const [customAnswers, setCustomAnswers] = useState<Record<string, unknown>>({});
  const [isAdmin, setIsAdmin] = useState(false);
  const [duplicateBlocked, setDuplicateBlocked] = useState(false);

  // Check if user is admin (group membership)
  useEffect(() => {
    let cancelled = false;
    async function check() {
      try {
        const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");
        const resp = await instance.acquireTokenSilent({
          scopes: [`${new URL(SP_SITE_URL).origin}/AllSites.Manage`],
          account: accounts[0],
        });
        const token = resp.accessToken;
        const groupResp = await fetch(
          `${SP_SITE_URL}/_api/web/sitegroups/getByName('_HR_ Forms Owners')/users?$select=Email`,
          { headers: { Accept: "application/json;odata=nometadata", Authorization: `Bearer ${token}` } },
        );
        if (groupResp.ok) {
          const data = await groupResp.json() as { value?: { Email?: string }[] };
          const userEmail = accounts[0]?.username?.toLowerCase() || "";
          if (!cancelled) {
            setIsAdmin((data.value || []).some((u) => (u.Email || "").toLowerCase() === userEmail));
          }
        }
      } catch {
        // Not admin — proceed as regular user
      }
    }
    void check();
    return () => { cancelled = true; };
  }, [instance, accounts]);

  // Fetch job details for summary
  useEffect(() => {
    let cancelled = false;
    async function load() {
      try {
        const jobs = await fetchJobs();
        const found = jobs.find((j) => j.id === jobId);
        if (!cancelled) setJob(found || null);
      } catch {
        if (!cancelled) setJob(null);
      } finally {
        if (!cancelled) setJobLoading(false);
      }
    }
    void load();
    return () => { cancelled = true; };
  }, [jobId]);

  const form = useReactiveForm<FormValues>({
    name: { value: "", validators: [required] },
    email: { value: "", validators: [required, email] },
    phone: { value: "", validators: [required, phone] },
    currentPosition: { value: "" },
    currentDepartment: { value: "" },
    coverLetter: { value: "" },
    files: { value: [], validators: [requiredFile] },
  });

  // Pre-fill from profile once loaded
  useEffect(() => {
    if (!profile.loading && !profile.error) {
      form.setValue({
        name: profile.displayName || form.value.name,
        email: profile.email || form.value.email,
        phone: profile.phone || form.value.phone,
        currentPosition: profile.jobTitle || form.value.currentPosition,
        currentDepartment: profile.department || form.value.currentDepartment,
      });
    }
    // Only run when profile loads
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [profile.loading, profile.error]);

  const setCustomAnswer = (name: string, value: unknown) => {
    setCustomAnswers((prev) => ({ ...prev, [name]: value }));
  };

  const doSubmit = async (values: FormValues, forceApply = false) => {
    if (!jobId || !job) return;
    setSubmitting(true);
    setSubmitError(null);
    setDuplicateBlocked(false);

    try {
      // Acquire user token with Graph scope for file uploads
      let accessToken = "";
      try {
        const resp = await instance.acquireTokenSilent({
          scopes: ["https://graph.microsoft.com/.default"],
          account: accounts[0],
        });
        accessToken = resp.accessToken;
      } catch {
        // User token not available — API will fall back to system credentials
      }

      // Generate submission ref client-side (mirrors generateSubmissionRef in API)
      const pdfSubmissionRef = `JOB-${new Date().getFullYear()}${String(new Date().getMonth() + 1).padStart(2, "0")}${String(new Date().getDate()).padStart(2, "0")}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`;

      // Generate PDF
      let pdfBase64 = "";
      try {
        const pdfBlob = await pdf(
          JobApplyPdfDocument({
            data: {
              jobTitle: job.title,
              applicantName: values.name,
              applicantEmail: values.email,
              applicantPhone: values.phone,
              currentPosition: values.currentPosition,
              currentDepartment: values.currentDepartment,
              submissionRef: pdfSubmissionRef,
              submittedAt: new Date().toISOString(),
              reasoning: values.coverLetter,
              customAnswers,
            },
          }),
        ).toBlob();
        pdfBase64 = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => {
            const result = reader.result as string;
            resolve(result.includes("base64,") ? result.split("base64,")[1] : result);
          };
          reader.onerror = reject;
          reader.readAsDataURL(pdfBlob);
        });
      } catch {
        // PDF generation failed — submit without it
      }

      const files = pdfBase64
        ? [...values.files, { name: "JobApplication.pdf", content: pdfBase64, contentType: "application/pdf" }]
        : values.files;

      const result = await submitApplication({
        jobListingId: jobId,
        jobTitle: job.title,
        applicantName: values.name,
        applicantEmail: values.email,
        applicantPhone: values.phone,
        currentPosition: values.currentPosition,
        currentDepartment: values.currentDepartment,
        coverLetter: values.coverLetter,
        files,
        customAnswers,
        accessToken,
        submittedByEmail: accounts[0]?.username || "",
        forceApply,
      });
      setSubmissionRef(result.submissionRef);
      setSubmitted(true);
    } catch (err) {
      const msg = err instanceof Error ? err.message : "";
      // Detect duplicate rejection (409)
      if (msg.includes("already applied")) {
        setDuplicateBlocked(true);
      }
      setSubmitError(msg || "Submission failed. Please try again.");
    } finally {
      setSubmitting(false);
    }
  };

  const handleSubmit = form.submit((values) => doSubmit(values, false));

  if (submitted) {
    return (
      <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
        <Container maxWidth="sm" sx={{ py: 8 }}>
          <SuccessView submissionRef={submissionRef} onBrowseMore={() => navigate("/careers")} />
        </Container>
      </Box>
    );
  }

  return (
    <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
      {/* Header */}
      <Paper
        sx={{
          borderRadius: 0,
          boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
          backgroundColor: "#ffffff",
          position: "sticky",
          top: 0,
          zIndex: 10,
        }}
      >
        <Container maxWidth="md">
          <Box sx={{ display: "flex", alignItems: "center", gap: 1.5, py: 2 }}>
            <IconButton onClick={() => navigate("/careers")} sx={{ color: "#6B7280" }}>
              <ArrowBack />
            </IconButton>
            <Box>
              <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>
                Back to jobs
              </Typography>
              <Typography variant="h6" sx={{ fontWeight: 700, color: "#111827", fontSize: "1.1rem" }}>
                Submit Application
              </Typography>
            </Box>
          </Box>
        </Container>
      </Paper>

      <Container maxWidth="md" sx={{ py: 4 }}>
        <Grid container spacing={3}>
          {/* Job Summary */}
          <Grid size={{ xs: 12, md: 4 }}>
            <Paper
              sx={{
                p: 2.5,
                borderRadius: "16px",
                boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
                position: "sticky",
                top: 88,
              }}
            >
              {jobLoading ? (
                <Box sx={{ display: "flex", justifyContent: "center", py: 3 }}>
                  <CircularProgress size={24} sx={{ color: "#0078D4" }} />
                </Box>
              ) : job ? (
                <>
                  <Box
                    sx={{
                      width: 48,
                      height: 48,
                      borderRadius: "12px",
                      backgroundColor: "#F0F7FF",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      mb: 1.5,
                    }}
                  >
                    <Work sx={{ fontSize: 24, color: "#0078D4" }} />
                  </Box>
                  <Typography variant="h6" sx={{ fontWeight: 700, color: "#111827", fontSize: "1rem", mb: 0.5 }}>
                    {job.title}
                  </Typography>
                  <Chip
                    label={job.department}
                    size="small"
                    sx={{ backgroundColor: "#6264A7", color: "#ffffff", fontWeight: 500, fontSize: "0.7rem", borderRadius: "8px", mb: 1.5 }}
                  />
                  {job.location && (
                    <Box sx={{ display: "flex", alignItems: "center", gap: 0.5, mb: 0.5 }}>
                      <LocationOn sx={{ fontSize: 14, color: "#6B7280" }} />
                      <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>
                        {job.location}
                      </Typography>
                    </Box>
                  )}
                  <Chip
                    label={job.employmentType}
                    size="small"
                    variant="outlined"
                    sx={{ borderRadius: "8px", fontSize: "0.7rem", borderColor: "#D1D5DB", color: "#6B7280" }}
                  />
                </>
              ) : (
                <Typography variant="body2" sx={{ color: "#9CA3AF" }}>
                  Job not found.
                </Typography>
              )}
            </Paper>
          </Grid>

          {/* Application Form */}
          <Grid size={{ xs: 12, md: 8 }}>
            <Paper sx={{ p: 3, borderRadius: "16px", boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
              {/* Profile loading indicator */}
              {profile.loading && (
                <Box sx={{ display: "flex", alignItems: "center", gap: 1, mb: 2 }}>
                  <CircularProgress size={14} sx={{ color: "#0078D4" }} />
                  <Typography variant="caption" sx={{ color: "#6B7280" }}>
                    Loading your profile...
                  </Typography>
                </Box>
              )}
              {profile.error && (
                <Alert severity="warning" sx={{ mb: 2, borderRadius: "8px" }}>
                  Could not load profile. Please fill in your details manually.
                </Alert>
              )}

              <form onSubmit={handleSubmit}>
                <Box sx={{ display: "flex", flexDirection: "column", gap: 2.5 }}>
                  {/* Name */}
                  <TextField
                    label="Full Name"
                    value={form.controls.name.value}
                    onChange={(e) => form.controls.name.setValue(e.target.value)}
                    onBlur={form.controls.name.onBlur}
                    error={form.controls.name.touched && !!form.controls.name.errors.required}
                    helperText={form.controls.name.touched && form.controls.name.errors.required ? "Name is required" : ""}
                    fullWidth
                    required
                    variant="outlined"
                    slotProps={{
                      input: { sx: { borderRadius: "10px" } },
                    }}
                  />

                  {/* Email */}
                  <TextField
                    label="Email Address"
                    type="email"
                    value={form.controls.email.value}
                    onChange={(e) => form.controls.email.setValue(e.target.value)}
                    onBlur={form.controls.email.onBlur}
                    error={form.controls.email.touched && !!(form.controls.email.errors.required || form.controls.email.errors.email)}
                    helperText={
                      form.controls.email.touched && form.controls.email.errors.required
                        ? "Email is required"
                        : form.controls.email.touched && form.controls.email.errors.email
                          ? "Please enter a valid email"
                          : ""
                    }
                    fullWidth
                    required
                    variant="outlined"
                    slotProps={{
                      input: { sx: { borderRadius: "10px" } },
                    }}
                  />

                  {/* Phone */}
                  <TextField
                    label="Phone Number"
                    value={form.controls.phone.value}
                    onChange={(e) => form.controls.phone.setValue(e.target.value)}
                    onBlur={form.controls.phone.onBlur}
                    error={form.controls.phone.touched && !!(form.controls.phone.errors.required || form.controls.phone.errors.phone)}
                    helperText={
                      form.controls.phone.touched && form.controls.phone.errors.required
                        ? "Phone number is required"
                        : form.controls.phone.touched && form.controls.phone.errors.phone
                          ? "Please enter a valid Malaysian phone number"
                          : ""
                    }
                    fullWidth
                    required
                    variant="outlined"
                    placeholder="e.g. +60 12-345 6789"
                    slotProps={{
                      input: { sx: { borderRadius: "10px" } },
                    }}
                  />

                  {/* Current Position */}
                  <TextField
                    label="Current Position"
                    value={form.controls.currentPosition.value}
                    onChange={(e) => form.controls.currentPosition.setValue(e.target.value)}
                    fullWidth
                    variant="outlined"
                    placeholder="e.g. Senior Engineer"
                    slotProps={{
                      input: { sx: { borderRadius: "10px" } },
                    }}
                  />

                  {/* Current Department */}
                  <TextField
                    label="Current Department"
                    value={form.controls.currentDepartment.value}
                    onChange={(e) => form.controls.currentDepartment.setValue(e.target.value)}
                    fullWidth
                    variant="outlined"
                    placeholder="e.g. Information Technology"
                    slotProps={{
                      input: { sx: { borderRadius: "10px" } },
                    }}
                  />

                  {/* Reasoning */}
                  <TextField
                    label="Reasoning (Optional)"
                    value={form.controls.coverLetter.value}
                    onChange={(e) => form.controls.coverLetter.setValue(e.target.value)}
                    fullWidth
                    multiline
                    rows={5}
                    variant="outlined"
                    placeholder="Explain your interest in this position and why you'd be a great fit..."
                    slotProps={{
                      input: { sx: { borderRadius: "10px" } },
                    }}
                  />

                  <Divider />

                  {/* File Upload */}
                  <FileUploadArea
                    files={form.controls.files.value}
                    onAdd={(newFiles) => {
                      const current = form.controls.files.value;
                      form.controls.files.setValue([...current, ...newFiles].slice(0, MAX_FILES));
                    }}
                    onRemove={(index) => {
                      const current = form.controls.files.value;
                      form.controls.files.setValue(current.filter((_, i) => i !== index));
                    }}
                  />
                  {form.controls.files.touched && form.controls.files.errors.requiredFile && (
                    <Typography variant="caption" sx={{ color: "#DC2626", fontWeight: 500, display: "flex", alignItems: "center", gap: 0.5 }}>
                      At least one supporting document is required (resume or CV).
                    </Typography>
                  )}
                  {/* Dynamic Custom Fields */}
                  {job?.customFields && job.customFields.length > 0 && (
                    <>
                      <Divider />
                      <Typography variant="subtitle2" sx={{ fontWeight: 600, color: "#374151" }}>
                        Additional Questions
                      </Typography>
                      {job.customFields.map((field: CustomFieldDefinition) => (
                        <Box key={field.name}>
                          {field.type === "choice" ? (
                            <FormControl fullWidth>
                              <InputLabel>{field.label}{field.required ? " *" : ""}</InputLabel>
                              <Select
                                value={String(customAnswers[field.name] ?? "")}
                                label={`${field.label}${field.required ? " *" : ""}`}
                                onChange={(e) => setCustomAnswer(field.name, e.target.value)}
                                sx={{ borderRadius: "10px" }}
                              >
                                {(field.choices || []).map((opt) => (
                                  <MenuItem key={opt} value={opt}>{opt}</MenuItem>
                                ))}
                              </Select>
                              {field.required && !customAnswers[field.name] && (
                                <FormHelperText error>This field is required</FormHelperText>
                              )}
                            </FormControl>
                          ) : field.type === "textarea" ? (
                            <TextField
                              label={`${field.label}${field.required ? " *" : ""}`}
                              value={String(customAnswers[field.name] ?? "")}
                              onChange={(e) => setCustomAnswer(field.name, e.target.value)}
                              fullWidth
                              multiline
                              rows={3}
                              variant="outlined"
                              slotProps={{ input: { sx: { borderRadius: "10px" } } }}
                            />
                          ) : field.type === "number" ? (
                            <TextField
                              label={`${field.label}${field.required ? " *" : ""}`}
                              type="number"
                              value={String(customAnswers[field.name] ?? "")}
                              onChange={(e) => setCustomAnswer(field.name, e.target.value)}
                              fullWidth
                              variant="outlined"
                              slotProps={{ input: { sx: { borderRadius: "10px" } } }}
                            />
                          ) : field.type === "date" ? (
                            <TextField
                              label={`${field.label}${field.required ? " *" : ""}`}
                              type="date"
                              value={String(customAnswers[field.name] ?? "")}
                              onChange={(e) => setCustomAnswer(field.name, e.target.value)}
                              fullWidth
                              variant="outlined"
                              slotProps={{
                                input: { sx: { borderRadius: "10px" } },
                                inputLabel: { shrink: true },
                              }}
                            />
                          ) : (
                            <TextField
                              label={`${field.label}${field.required ? " *" : ""}`}
                              value={String(customAnswers[field.name] ?? "")}
                              onChange={(e) => setCustomAnswer(field.name, e.target.value)}
                              fullWidth
                              variant="outlined"
                              slotProps={{ input: { sx: { borderRadius: "10px" } } }}
                            />
                          )}
                        </Box>
                      ))}
                    </>
                  )}

                  {/* Error */}
                  {submitError && (
                    <Alert
                      severity={duplicateBlocked ? "warning" : "error"}
                      sx={{
                        borderRadius: "8px",
                        fontWeight: 500,
                        fontSize: "0.85rem",
                        ...(duplicateBlocked ? {} : {
                          backgroundColor: "#FEF2F2",
                          color: "#991B1B",
                          "& .MuiAlert-icon": { color: "#DC2626" },
                        }),
                      }}
                    >
                      {submitError}
                    </Alert>
                  )}

                  {/* Test Submit (admin only — bypasses duplicate check) */}
                  {duplicateBlocked && isAdmin && (
                    <Button
                      variant="outlined"
                      fullWidth
                      disabled={submitting}
                      onClick={() => doSubmit(form.value, true)}
                      sx={{
                        borderRadius: "12px",
                        textTransform: "none",
                        fontWeight: 600,
                        fontSize: "0.9rem",
                        py: 1.3,
                        borderColor: "#E67635",
                        color: "#E67635",
                        "&:hover": { borderColor: "#D4621A", backgroundColor: "rgba(230, 118, 53, 0.06)" },
                      }}
                    >
                      Test Submit (bypass duplicate check)
                    </Button>
                  )}

                  {/* Submit */}
                  <Button
                    type="submit"
                    variant="contained"
                    fullWidth
                    disabled={submitting || !form.valid}
                    sx={{
                      borderRadius: "12px",
                      textTransform: "none",
                      backgroundColor: "#0078D4",
                      fontWeight: 600,
                      fontSize: "0.95rem",
                      py: 1.5,
                      boxShadow: "0 2px 8px rgba(0, 120, 212, 0.25)",
                      transition: "all 0.25s",
                      "&:hover": {
                        backgroundColor: "#106EBE",
                        boxShadow: "0 4px 14px rgba(0, 120, 212, 0.35)",
                      },
                      "&:disabled": {
                        backgroundColor: "#93C5FD",
                      },
                    }}
                  >
                    {submitting ? (
                      <Box sx={{ display: "flex", alignItems: "center", gap: 1 }}>
                        <CircularProgress size={18} sx={{ color: "#ffffff" }} />
                        <span>Submitting...</span>
                      </Box>
                    ) : (
                      "Submit Application"
                    )}
                  </Button>

                  <Typography variant="caption" sx={{ color: "#9CA3AF", textAlign: "center" }}>
                    By submitting, you agree to the processing of your personal data for recruitment purposes.
                  </Typography>
                </Box>
              </form>
            </Paper>
          </Grid>
        </Grid>
      </Container>
    </Box>
  );
}
