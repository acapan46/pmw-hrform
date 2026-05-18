/**
 * EvaluationPage.tsx — Layer evaluation/approval interface.
 * Route: /eval/:token (public) or /eval/:formSlug/:responseId/:layerNumber (365)
 */
import { useEffect, useState, useCallback } from "react";
import { useParams } from "react-router-dom";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";

import { getLayerResponseData, updateLayerStatus, submitEvaluationData, getFormConfigByTitle, spGet, readMatrixChildItems } from "../utils/formBuilderSP";
import type { MatrixColumnDef } from "../utils/formBuilderSP";
import { SP_LAYER_STATUS } from "../utils/statusConstants";
import type { LayerConfigItem, EvaluationDataEntry } from "../types";
import DOMPurify from "dompurify";
import EvaluationSummary from "../components/builder/EvaluationSummary";
import { loginRequest } from "../auth/msalConfig";
import { generateAndStorePdf, buildPdfLayerResults } from "../utils/generateFormPdf";
import type { PdfFormData } from "../utils/FormPdfDocument";
import { rowsToHtml, getDynamicMatrixFields } from "../utils/DynamicMatrix";
import LockIcon from "@mui/icons-material/Lock";
import WarningIcon from "@mui/icons-material/Warning";

const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");

// ── PDF Helper ─────────────────────────────────────────────────────────────
async function loadPdfAndGenerate(token: string, listTitle: string, responseItemId: number, formTitle: string, formStatus: string): Promise<void> {
  try {
    const cfg = await getFormConfigByTitle(token, formTitle);
    if (!cfg) return;

    const formVersion = (cfg as unknown as Record<string, unknown>).CurrentVersion as string || "1.0";

    const versionData = await spGet(
      token,
      `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(cfg.Title)}' and FormVersion eq '${encodeURIComponent(formVersion)}'&$select=SurveyJSON&$top=1`
    ) as { value?: { SurveyJSON?: string }[] };

    const rawSurvey = versionData.value?.[0]?.SurveyJSON;
    if (!rawSurvey) return;

    const parsed = JSON.parse(rawSurvey);
    const surveyContent = parsed.surveyJson || parsed;

    const respItem = await spGet(
      token,
      `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${responseItemId})`
    ) as Record<string, unknown>;

    const SYSTEM_FIELDS = new Set([
      'Id','Title','SubmittedBy','SubmittedAt','Status','CurrentApprovalLayer',
      'FormVersion','FormID','RawJSON','CurrentLayer','FormStatus','EvaluationData',
      'Author','Editor','Created','Modified','ContentType','PermMask',
      'L1_Status','L1_Email','L1_SignedAt','L1_Rejection','L1_Signature',
      'L2_Status','L2_Email','L2_SignedAt','L2_Rejection','L2_Signature',
      'L3_Status','L3_Email','L3_SignedAt','L3_Rejection','L3_Signature',
    ]);

    const data: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(respItem)) {
      if (!SYSTEM_FIELDS.has(k) && v !== null && v !== undefined) {
        data[k] = v;
      }
    }

    await generateAndStorePdf(token, listTitle, responseItemId, {
      surveyJson: surveyContent as PdfFormData["surveyJson"],
      responseData: data,
      layerResults: buildPdfLayerResults(respItem),
      meta: {
        submittedBy: (respItem.SubmittedBy as string) || "",
        submittedAt: (respItem.SubmittedAt as string) || "",
        formTitle,
        formVersion,
        formStatus,
      },
      logoUrl: "/logo-128.png",
    });
  } catch (e) {
    console.warn("[loadPdfAndGenerate] failed:", e);
  }
}

type AuthState = "checking" | "authorized" | "unauthorized" | "error";
type ActionState = "idle" | "submitting" | "success" | "error";

// ── Styling ──
const COLORS = {
  purple: "#5B21B6", purpleLight: "#7C3AED", purplePale: "#EDE9FE",
  bg: "#F8F7FF", cardBg: "#FFFFFF", border: "#E5E3F0",
  textPrimary: "#1E1B4B", textSecond: "#6B7280", textMuted: "#9CA3AF",
  green: "#059669", greenPale: "#D1FAE5",
  red: "#DC2626", redPale: "#FEE2E2",
  shadow: "0 1px 3px rgba(91,33,182,0.08),0 4px 16px rgba(91,33,182,0.06)",
};

const sectionCard: React.CSSProperties = {
  background: COLORS.cardBg,
  border: `1px solid ${COLORS.border}`,
  borderRadius: 16,
  padding: 24,
  marginBottom: 20,
  boxShadow: COLORS.shadow,
};

const btnPrimary: React.CSSProperties = {
  padding: "12px 32px",
  borderRadius: 10,
  border: "none",
  background: `linear-gradient(135deg, ${COLORS.purple}, ${COLORS.purpleLight})`,
  color: "#fff",
  fontSize: 14,
  fontWeight: 600,
  cursor: "pointer",
  fontFamily: "'Segoe UI', system-ui, sans-serif",
};

const btnOutline: React.CSSProperties = {
  ...btnPrimary,
  background: "transparent",
  border: `1.5px solid ${COLORS.red}`,
  color: COLORS.red,
};

// ── Component ──
export default function EvaluationPage() {
  const { token: routeToken, formSlug, responseId, layerNumber } = useParams<{
    token: string;
    formSlug: string;
    responseId: string;
    layerNumber: string;
  }>();
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [authState, setAuthState] = useState<AuthState>("checking");
  const [userEmail, setUserEmail] = useState<string | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const [responseData, setResponseData] = useState<Record<string, unknown> | null>(null);
  const [currentLayer, setCurrentLayer] = useState<LayerConfigItem | null>(null);
  const [previousResults, setPreviousResults] = useState<Record<string, unknown>[]>([]);
  const [formTitle, setFormTitle] = useState("");

  const [actionState, setActionState] = useState<ActionState>("idle");
  const [evaluationFields] = useState<Record<string, unknown>>({});
  const [rejectionReason, setRejectionReason] = useState("");
  const [signatureData] = useState<string | null>(null);
  const [checkboxApproved, setCheckboxApproved] = useState(false);
  const [matrixTables, setMatrixTables] = useState<Record<string, { columns: MatrixColumnDef[]; rows: Record<string, unknown>[]; html: string }>>({});

  const isPublic = !!routeToken;
  const displayLayerNumber = isPublic
    ? 1  // Will be resolved from token
    : parseInt(layerNumber || "0", 10);

  // ── Auth ──
  useEffect(() => {
    if (isPublic) {
      // Public mode — no auth needed, but need SP token for potential writes
      setAuthState("authorized");
      setUserEmail("SYSTEM");
      return;
    }
    if (inProgress !== InteractionStatus.None) return;
    if (!isAuthenticated) {
      setAuthState("unauthorized");
      setLoading(false);
      return;
    }
    const email = accounts[0]?.username || null;
    setUserEmail(email);
    const origin = new URL(SP_SITE_URL).origin;
    instance.acquireTokenSilent({ scopes: [`${origin}/AllSites.Manage`], account: accounts[0] })
      .then(r => { setToken(r.accessToken); setAuthState("authorized"); })
      .catch(() => { setAuthState("error"); setError("Failed to acquire token."); });
  }, [isPublic, isAuthenticated, inProgress, instance, accounts]);

  // ── Load data ──
  useEffect(() => {
    if (authState !== "authorized") return;
    if (isPublic) {
      // Public: fetch filtered data from API
      const loadPublic = async () => {
        try {
          const params = new URLSearchParams(window.location.search);
          const itemId = params.get("item");
          if (!itemId) { setError("Missing response item ID."); setLoading(false); return; }

          const res = await fetch(`/api/evaluate?token=${encodeURIComponent(routeToken || "")}&responseItemId=${itemId}`);
          const json = await res.json();
          if (!json.success) { setError(json.error || "Failed to load data."); setLoading(false); return; }

          setFormTitle(json.data.formTitle);
          setResponseData(json.data.fields);
          setCurrentLayer({
            layerNumber: json.data.layerNumber,
            type: json.data.layerType,
            authMode: "public" as const,
            assignee: { type: "user" as const, value: "" },
            title: json.data.layerTitle,
          } as LayerConfigItem);

          // Build previous results from the filtered fields
          const prev: Record<string, unknown>[] = [];
          if (json.data.totalLayers > 0) {
            for (let n = 1; n < json.data.layerNumber; n++) {
              prev.push({
                layerNumber: n,
                status: json.data.fields[`L${n}_Status`] || null,
                email: json.data.fields[`L${n}_Email`] || null,
                signedAt: json.data.fields[`L${n}_SignedAt`] || null,
              });
            }
          }
          setPreviousResults(prev);
          setLoading(false);
        } catch (e) {
          setError(e instanceof Error ? e.message : "Failed to load evaluation data.");
          setLoading(false);
        }
      };
      loadPublic();
      return; // Skip the 365 load path
    }
    if (!formSlug || !responseId || !displayLayerNumber) {
      setError("Invalid URL parameters.");
      setLoading(false);
      return;
    }

    const load = async () => {
      if (!token) return;
      try {
        // Resolve formTitle from slug
        const slugData = await fetch(`${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$filter=Slug eq '${encodeURIComponent(formSlug)}'&$select=Title,LayerConfig&$top=1`, {
          headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
        });
        const slugJson = await slugData.json();
        const resolvedTitle = slugJson.value?.[0]?.Title;
        if (!resolvedTitle) { setError("Form not found."); setLoading(false); return; }
        setFormTitle(resolvedTitle);

        const data = await getLayerResponseData(token, resolvedTitle, parseInt(responseId, 10), displayLayerNumber);
        if (!data) { setError("Could not load evaluation data."); setLoading(false); return; }
        setResponseData(data.responseFields);
        setCurrentLayer(data.currentLayer || null);
        setPreviousResults(data.previousResults);

        // Load matrix child list data for dynamicmatrix fields
        const itemFormVersion = data.responseFields.FormVersion as string | undefined;
        if (itemFormVersion) {
          loadMatrixChildData(token, resolvedTitle, parseInt(responseId, 10), itemFormVersion);
        }
      } catch (e) {
        setError(e instanceof Error ? e.message : "Failed to load data.");
      }
      setLoading(false);
    };
    load();
  }, [authState, isPublic, formSlug, responseId, displayLayerNumber, token]);

  // ── Submit action (365 mode) ──
  const handleSubmit = useCallback(async (action: "approve" | "reject" | "confirm") => {
    if (!token || !userEmail) return;
    setActionState("submitting");
    try {
      const listTitle = formTitle; // list is named after form title
      const respId = parseInt(responseId || "0", 10);
      const now = new Date().toISOString();

      if (action === "reject") {
        await updateLayerStatus(token, listTitle, respId, displayLayerNumber, {
          status: SP_LAYER_STATUS.REJECTED,
          signedAt: now,
          rejection: rejectionReason,
        });
        // Update FormStatus
        const formStatusUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${respId})`;
        await fetch(formStatusUrl, {
          method: "PATCH",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": await (await fetch(`${SP_SITE_URL}/_api/contextinfo`, {
              method: "POST", headers: { Authorization: `Bearer ${token}` },
            })).json().then((d: { FormDigestValue: string }) => d.FormDigestValue),
          },
          body: JSON.stringify({ FormStatus: "Rejected" }),
        });
        // Generate PDF on rejection
        await loadPdfAndGenerate(token, listTitle, respId, formTitle, "rejected").catch(e => console.warn("[EvalPage] PDF failed:", e));
      } else if (action === "confirm" && currentLayer?.type === "evaluation") {
        await submitEvaluationData(token, listTitle, respId, displayLayerNumber, {
          confirmerEmail: userEmail,
          confirmerName: accounts[0]?.name ?? undefined,
          fields: evaluationFields,
          signatureUrl: signatureData,
        });
        await updateLayerStatus(token, listTitle, respId, displayLayerNumber, {
          status: SP_LAYER_STATUS.CONFIRMED,
          signedAt: now,
          signature: signatureData || undefined,
        });
        // Generate PDF on evaluation confirmed
        await loadPdfAndGenerate(token, listTitle, respId, formTitle, "confirmed").catch(e => console.warn("[EvalPage] PDF failed:", e));
      } else if (action === "approve") {
        await updateLayerStatus(token, listTitle, respId, displayLayerNumber, {
          status: SP_LAYER_STATUS.APPROVED,
          signedAt: now,
          signature: signatureData || undefined,
        });
        // Generate PDF on approval
        await loadPdfAndGenerate(token, listTitle, respId, formTitle, "approved").catch(e => console.warn("[EvalPage] PDF failed:", e));
      }

      setActionState("success");
    } catch (e) {
      console.error("[EvalPage] submit error:", e);
      setActionState("error");
    }
  }, [token, userEmail, formTitle, responseId, displayLayerNumber, rejectionReason, evaluationFields, signatureData, currentLayer, accounts]);

  /** Load matrix child list data for dynamicmatrix fields and enrich responseData */
  const loadMatrixChildData = async (
    tkn: string,
    resolvedTitle: string,
    respId: number,
    formVersion: string
  ) => {
    try {
      // Load the version's SurveyJSON to detect dynamicmatrix fields
      const versionData = await spGet(
        tkn,
        `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(resolvedTitle)}' and FormVersion eq '${encodeURIComponent(formVersion)}'&$select=SurveyJSON&$top=1`
      ) as { value?: { SurveyJSON?: string }[] };

      const rawSurvey = versionData.value?.[0]?.SurveyJSON;
      if (!rawSurvey) return;

      const parsed = JSON.parse(rawSurvey);
      const surveyDef = parsed.surveyJson || parsed;
      const matrixFields = getDynamicMatrixFields(surveyDef);

      if (matrixFields.length === 0) return;

      const tables: Record<string, { columns: MatrixColumnDef[]; rows: Record<string, unknown>[]; html: string }> = {};
      for (const mf of matrixFields) {
        const safeName = mf.name.replace(/[^a-zA-Z0-9_ -]/g, "").trim();
        const childListName = `${resolvedTitle} Matrix ${safeName}`;

        try {
          const rows = await readMatrixChildItems(tkn, childListName, respId);
          if (rows.length > 0) {
            const cols = mf.columns as MatrixColumnDef[];
            tables[mf.name] = {
              columns: cols,
              rows,
              html: rowsToHtml(mf.columns, rows),
            };
          }
        } catch {
          // Child list not found — skip this field
        }
      }

      setMatrixTables(tables);

      // Enrich responseData with matrix data in SurveyJS-compatible format
      if (Object.keys(tables).length > 0) {
        setResponseData((prev) => {
          if (!prev) return prev;
          const enriched = { ...prev };
          for (const [fieldName, entry] of Object.entries(tables)) {
            enriched[fieldName] = {
              rows: entry.rows,
              html: entry.html,
              json: JSON.stringify(entry.rows),
            };
          }
          return enriched;
        });
      }
    } catch {
      // Silently fail — matrix data is non-critical
    }
  };

  // ── Render ──
  if (authState === "checking" || loading) {
    return (
      <div style={{ minHeight: "100vh", background: COLORS.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ color: COLORS.textMuted, fontSize: 14 }}>Loading...</div>
      </div>
    );
  }

  if (authState === "unauthorized") {
    return (
      <div style={{ minHeight: "100vh", background: COLORS.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
        <div style={{ background: COLORS.cardBg, borderRadius: 22, padding: "56px 44px", maxWidth: 420, width: "100%", textAlign: "center", border: `1px solid ${COLORS.border}`, boxShadow: COLORS.shadow }}>
          <div style={{ fontSize: 32, marginBottom: 16, display: 'flex', justifyContent: 'center' }}><LockIcon style={{ fontSize: 40 }} /></div>
          <div style={{ fontSize: 20, fontWeight: 700, color: COLORS.textPrimary, marginBottom: 8 }}>Sign in required</div>
          <p style={{ color: COLORS.textSecond, fontSize: 13, marginBottom: 24 }}>You need to sign in with your Microsoft 365 account to access this evaluation.</p>
          <button onClick={() => instance.loginRedirect({ ...loginRequest })} style={btnPrimary}>Sign in with Microsoft 365</button>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div style={{ minHeight: "100vh", background: COLORS.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
        <div style={{ background: COLORS.cardBg, borderRadius: 22, padding: "56px 44px", maxWidth: 420, textAlign: "center", border: `1px solid ${COLORS.border}` }}>
          <div style={{ fontSize: 32, marginBottom: 16, display: 'flex', justifyContent: 'center' }}><WarningIcon style={{ fontSize: 40 }} /></div>
          <div style={{ fontSize: 20, fontWeight: 700, color: COLORS.red, marginBottom: 8 }}>Error</div>
          <p style={{ color: COLORS.textSecond, fontSize: 13 }}>{error}</p>
        </div>
      </div>
    );
  }

  if (actionState === "success") {
    return (
      <div style={{ minHeight: "100vh", background: COLORS.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
        <div style={{ background: COLORS.cardBg, borderRadius: 22, padding: "56px 44px", maxWidth: 420, textAlign: "center", border: `1px solid ${COLORS.border}`, boxShadow: COLORS.shadow }}>
          <div style={{ fontSize: 48, marginBottom: 16 }}>✓</div>
          <div style={{ fontSize: 22, fontWeight: 700, color: COLORS.green, marginBottom: 8 }}>Submitted Successfully</div>
          <p style={{ color: COLORS.textSecond, fontSize: 13, marginBottom: 24 }}>
            Your response has been recorded. You may close this page.
          </p>
        </div>
      </div>
    );
  }

  const isEvaluation = currentLayer?.type === "evaluation";
  const isSignatureRequired = currentLayer?.type === "approval" && (currentLayer as unknown as Record<string, unknown>).confirmationType === "signature";
  const isCheckboxMode = currentLayer?.type === "approval" && (currentLayer as unknown as Record<string, unknown>).confirmationType === "checkbox";

  return (
    <div style={{ minHeight: "100vh", background: COLORS.bg, padding: "32px 16px" }}>
      <div style={{ maxWidth: 720, margin: "0 auto" }}>

        {/* Header */}
        <div style={{ fontSize: 24, fontWeight: 700, color: COLORS.textPrimary, marginBottom: 4 }}>
          {currentLayer?.title || (isEvaluation ? "Evaluation" : "Approval")}
        </div>
        <div style={{ fontSize: 13, color: COLORS.textSecond, marginBottom: 24 }}>
          {formTitle || "Form"} — Layer {displayLayerNumber}
          {currentLayer?.description && <div style={{ marginTop: 4 }}>{currentLayer.description}</div>}
        </div>

        {/* Previous Layer Results */}
        {previousResults.length > 0 && (
          <div style={sectionCard}>
            <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.textSecond, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 12 }}>
              Previous Layers
            </div>
            {previousResults.map((pr, i) => {
              const evalData = pr.evaluationData as EvaluationDataEntry | undefined;
              if (evalData?.status === "confirmed") {
                return (
                  <EvaluationSummary
                    key={i}
                    result={{
                      layerNumber: pr.layerNumber as number,
                      type: "evaluation",
                      status: "confirmed",
                      email: evalData.confirmerEmail || null,
                      confirmedAt: evalData.confirmedAt || null,
                      fields: evalData.fields || {},
                      notes: evalData.notes,
                    }}
                    layerTitle={`Layer ${pr.layerNumber}`}
                  />
                );
              }
              return (
                <div key={i} style={{ background: COLORS.purplePale, borderRadius: 10, padding: "12px 16px", marginBottom: 10, fontSize: 13, color: COLORS.textPrimary }}>
                  Layer {Number(pr.layerNumber)}: <strong>{String(pr.status || "Completed")}</strong>
                  {pr.signedAt ? <span style={{ color: COLORS.textMuted, marginLeft: 8 }}>— {new Date(pr.signedAt as string).toLocaleDateString()}</span> : null}
                </div>
              );
            })}
          </div>
        )}

        {/* Submission Data Preview */}
        {responseData && (
          <div style={sectionCard}>
            <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.textSecond, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 12 }}>
              Submission Data
            </div>
            <div style={{ fontSize: 13, color: COLORS.textSecond }}>
              <div>Form ID: {String(responseData.FormID || responseData.formId || "—")}</div>
              <div>Submitted: {responseData.SubmittedAt ? new Date(String(responseData.SubmittedAt)).toLocaleDateString() : "—"}</div>
              <div>Version: {String(responseData.FormVersion || responseData.formVersion || "—")}</div>
            </div>

            {/* Matrix Tables — from child lists */}
            {Object.keys(matrixTables).length > 0 && (
              <div style={{ marginTop: 16, borderTop: `1px solid ${COLORS.border}`, paddingTop: 16 }}>
                <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.purple, marginBottom: 12 }}>
                  Matrix Tables
                </div>
                {Object.entries(matrixTables).map(([fieldName, entry]) => (
                  <div key={fieldName} style={{ marginBottom: 16 }}>
                    <div style={{ fontSize: 11, fontWeight: 600, color: COLORS.textMuted, marginBottom: 4 }}>
                      {entry.columns[0]?.title || fieldName}
                    </div>
                    <div
                      style={{ overflow: "auto", border: `1px solid ${COLORS.border}`, borderRadius: 8 }}
                      dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(entry.html) }}
                    />
                    <div style={{ fontSize: 10, color: COLORS.textMuted, marginTop: 4 }}>
                      {entry.rows.length} row{entry.rows.length !== 1 ? "s" : ""}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Current Layer Action */}
        <div style={sectionCard}>
          <div style={{ fontSize: 13, fontWeight: 700, color: COLORS.textPrimary, marginBottom: 16 }}>
            {isEvaluation ? "Your Evaluation" : "Your Decision"}
          </div>

          {isEvaluation && (
            /* Evaluation form — SurveyJS would go here, simplified for now */
            <div style={{ fontSize: 13, color: COLORS.textSecond, marginBottom: 16, fontStyle: "italic" }}>
              Evaluation fields from layer configuration will render here with SurveyJS.
            </div>
          )}

          {isSignatureRequired && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.textMuted, marginBottom: 6 }}>
                Signature
              </div>
              <div style={{ fontSize: 13, color: COLORS.textSecond, fontStyle: "italic" }}>
                Signature pad will be rendered here.
              </div>
            </div>
          )}

          {isCheckboxMode && (
            <label style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, cursor: "pointer" }}>
              <input
                type="checkbox"
                checked={checkboxApproved}
                onChange={(e) => setCheckboxApproved(e.target.checked)}
                style={{ width: 18, height: 18, accentColor: COLORS.purple }}
              />
              <span style={{ fontSize: 14, color: COLORS.textPrimary }}>I approve this submission</span>
            </label>
          )}

          {/* Rejection reason (always available for approval layers) */}
          {!isEvaluation && (
            <div style={{ marginBottom: 16 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.textMuted, marginBottom: 6 }}>
                Rejection Reason <span style={{ fontWeight: 400, color: COLORS.textMuted }}>(optional)</span>
              </div>
              <textarea
                value={rejectionReason}
                onChange={(e) => setRejectionReason(e.target.value)}
                placeholder="Enter reason if rejecting..."
                style={{
                  width: "100%",
                  minHeight: 60,
                  padding: 10,
                  borderRadius: 8,
                  border: `1px solid ${COLORS.border}`,
                  fontSize: 13,
                  fontFamily: "inherit",
                  resize: "vertical",
                  outline: "none",
                }}
              />
            </div>
          )}

          {/* Action buttons */}
          <div style={{ display: "flex", gap: 12 }}>
            {isEvaluation ? (
              <button onClick={() => handleSubmit("confirm")} style={btnPrimary} disabled={actionState === "submitting"}>
                {actionState === "submitting" ? "Submitting..." : "Confirm Evaluation"}
              </button>
            ) : (
              <>
                <button
                  onClick={() => handleSubmit(isCheckboxMode && !checkboxApproved ? "reject" : "approve")}
                  style={isCheckboxMode && !checkboxApproved ? btnOutline : btnPrimary}
                  disabled={actionState === "submitting" || (isCheckboxMode && !checkboxApproved)}
                >
                  {actionState === "submitting" ? "Submitting..." : isCheckboxMode && !checkboxApproved ? "Reject" : "Approve"}
                </button>
                <button onClick={() => handleSubmit("reject")} style={btnOutline} disabled={actionState === "submitting"}>
                  Reject
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}