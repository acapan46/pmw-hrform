/**
 * DynamicFormPage.tsx - Public form renderer
 * Route: /form/:formId
 */
import { useEffect, useState, useMemo, useCallback, useRef } from "react";
import { useParams, useSearchParams } from "react-router-dom";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { Model, Serializer } from "survey-core";
import { Survey } from "survey-react-ui";
import { LayeredDarkPanelless, LayeredLightPanelless } from "survey-core/themes";
import "survey-core/survey-core.min.css";

import { getLatestFormBySlug, getFormVersion, spGet, spPost, triggerApprovalNotification, getSharePointChoices, getFilteredListChoices, uploadSignatureImage, getFormConfigByTitle } from "../utils/formBuilderSP";
import type { LayerConfig } from "../types";
import { SP_LAYER_STATUS, SP_FORM_STATUS } from "../utils/statusConstants";
import { registerSignaturePad } from "../utils/SignaturePad";
import { loginRequest } from "../auth/msalConfig";
import { clearStoredAuthDecision } from "../utils/authDecision";
import Logo from "../components/Logo";
import { generateAndStorePdf, buildPdfLayerResults } from "../utils/generateFormPdf";
import type { PdfFormData } from "../utils/FormPdfDocument";

const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");

// ── Register custom SurveyJS widgets and properties ────────────────────
registerSignaturePad();

if (!Serializer.findProperty("text", "autocapitalize")) {
  Serializer.addProperty("text", {
    name: "autocapitalize",
    category: "general",
    choices: ["none", "sentences", "words", "characters"],
    default: "none",
  });
}

// Theme tokens
const LIGHT = {
  purple: "#5B21B6", purpleLight: "#7C3AED", purplePale: "#EDE9FE", purpleMid: "#DDD6FE",
  purpleDark: "#3B0764", bg: "#F8F7FF", cardBg: "#FFFFFF", offWhite: "#F8F7FF", border: "#E5E3F0",
  textPrimary: "#1E1B4B", textSecond: "#6B7280", textMuted: "#9CA3AF",
  green: "#059669", greenPale: "#D1FAE5", greenBorder: "#6EE7B7",
  red: "#DC2626", redPale: "#FEE2E2", amber: "#D97706", amberPale: "#FEF3C7",
  shadow: "0 1px 3px rgba(91,33,182,0.08),0 4px 16px rgba(91,33,182,0.06)",
  shadowLg: "0 8px 40px rgba(91,33,182,0.16)", shadowFab: "0 4px 20px rgba(91,33,182,0.18)",
};

const DARK = {
  ...LIGHT, bg: "#0F0B1E", cardBg: "#1A1330", offWhite: "#160E28", border: "#2D2456",
  textPrimary: "#EDE9FE", textSecond: "#A78BFA", textMuted: "#6D5FA6",
  greenPale: "#052e16", greenBorder: "#166534", redPale: "#3b0707", amberPale: "#2d1b00",
  shadow: "0 1px 3px rgba(0,0,0,.4),0 4px 16px rgba(0,0,0,.3)",
  shadowLg: "0 8px 40px rgba(0,0,0,.5)", shadowFab: "0 4px 20px rgba(0,0,0,.4)",
};

const globalCss = (t: typeof LIGHT) => `
  @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'DM Sans',sans-serif;background:${t.bg};color:${t.textPrimary};transition:background .3s,color .3s}
  @keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
  @keyframes spin{to{transform:rotate(360deg)}}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
   .dfp-survey-wrap .sd-root-modern{background:transparent!important}
.dfp-survey-wrap .sd-container-modern>.sd-title{text-align:center!important}
.dfp-survey-wrap .sd-row{display:flex!important;flex-wrap:wrap!important}
  ::-webkit-scrollbar{width:5px}
  ::-webkit-scrollbar-thumb{background:${t.purpleMid};border-radius:10px}
`;

const Spinner = ({ size = 30, t }: { size?: number; t: typeof LIGHT }) => (
  <div style={{ width: size, height: size, border: `2.5px solid ${t.purpleMid}`, borderTop: `2.5px solid ${t.purple}`, borderRadius: "50%", animation: "spin .85s linear infinite", flexShrink: 0 }} />
);

const MsIcon = () => (
  <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
    <rect x="1" y="1" width="6.5" height="6.5" fill="#F25022" />
    <rect x="8.5" y="1" width="6.5" height="6.5" fill="#7FBA00" />
    <rect x="1" y="8.5" width="6.5" height="6.5" fill="#00A4EF" />
    <rect x="8.5" y="8.5" width="6.5" height="6.5" fill="#FFB900" />
  </svg>
);

const ScrollProgress = ({ t }: { t: typeof LIGHT }) => {
  const [pct, setPct] = useState(0);
  useEffect(() => {
    const fn = () => {
      const el = document.documentElement;
      const total = el.scrollHeight - el.clientHeight;
      setPct(total > 0 ? Math.min(100, (el.scrollTop / total) * 100) : 0);
    };
    window.addEventListener("scroll", fn, { passive: true });
    return () => window.removeEventListener("scroll", fn);
  }, []);
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, height: 3, zIndex: 9999, pointerEvents: "none" }}>
      <div style={{ height: "100%", width: `${pct}%`, background: `linear-gradient(90deg,${t.purple},${t.purpleLight})`, transition: "width .1s linear", borderRadius: "0 2px 2px 0" }} />
    </div>
  );
};

const SuccessScreen = ({ formTitle, onReset, t }: { formTitle: string; onReset: () => void; t: typeof LIGHT }) => (
  <div style={{ textAlign: "center", padding: "60px 20px", animation: "fadeUp .3s ease" }}>
    <div style={{ width: 72, height: 72, borderRadius: "50%", background: t.greenPale, border: `2px solid ${t.greenBorder}`, display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 20px", fontSize: 32 }}>OK</div>
    <div style={{ fontFamily: "'DM Serif Display',serif", fontSize: 26, color: t.textPrimary, marginBottom: 10 }}>Submitted Successfully</div>
    <p style={{ color: t.textSecond, fontSize: 14, lineHeight: 1.8, maxWidth: 420, margin: "0 auto 10px" }}>Your response for <strong>{formTitle}</strong> has been recorded.</p>
    <button onClick={onReset} style={{ padding: "11px 30px", borderRadius: 10, border: `1px solid ${t.border}`, background: t.cardBg, color: t.textSecond, fontSize: 13, cursor: "pointer", fontFamily: "'DM Sans'" }}>Submit another response</button>
  </div>
);

const PrivateGate = ({ formTitle, onSignIn, t }: { formTitle: string; onSignIn: () => void; t: typeof LIGHT }) => (
  <div style={{ minHeight: "100vh", background: t.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
    <div style={{ background: t.cardBg, borderRadius: 22, padding: "56px 44px", maxWidth: 420, width: "100%", textAlign: "center", boxShadow: t.shadowLg, border: `1px solid ${t.border}`, animation: "fadeUp .3s ease" }}>
      <div style={{ width: 66, height: 66, borderRadius: 18, margin: "0 auto 22px", background: t.purplePale, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 30 }}>LOCK</div>
      <div style={{ fontFamily: "'DM Serif Display',serif", fontSize: 24, color: t.textPrimary, marginBottom: 10 }}>Sign in required</div>
      <p style={{ color: t.textSecond, fontSize: 13, lineHeight: 1.7, marginBottom: 32 }}><strong>{formTitle || "This form"}</strong> is restricted.</p>
      <button onClick={onSignIn} style={{ width: "100%", padding: "14px", borderRadius: 11, border: "none", background: `linear-gradient(135deg,${t.purple},${t.purpleLight})`, color: "#fff", fontSize: 14, fontWeight: 600, cursor: "pointer", fontFamily: "'DM Sans'", display: "flex", alignItems: "center", justifyContent: "center", gap: 10 }}>
        <MsIcon /> Sign in with Microsoft 365
      </button>
    </div>
  </div>
);

export default function DynamicFormPage() {
  const { formId } = useParams<{ formId: string }>();
  const [searchParams] = useSearchParams();
  const pinVersion = searchParams.get("version");
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [dark, _setDark] = useState(() => { try { return localStorage.getItem("dfp_dark") === "1"; } catch { return false; } });
  const t = dark ? DARK : LIGHT;

  useEffect(() => { document.body.style.background = t.bg; document.body.style.color = t.textPrimary; return () => { document.body.style.background = ""; document.body.style.color = ""; }; }, [t]);

  const [loading, setLoading] = useState(true);
  const [formData, setFormData] = useState<{ formConfig: Record<string, unknown>; surveyJson: Record<string, unknown>; meta: Record<string, unknown> } | null>(null);
  const [enrichedSurveyJson, setEnrichedSurveyJson] = useState<Record<string, unknown> | null>(null);
  const [error, setError] = useState("");
  const [submitStatus, setSubmitStatus] = useState<string | null>(null);
  const [resetKey, setResetKey] = useState(0);
  const tokenRef = useRef<string | null>(null);
  const userEmail = accounts[0]?.username || null;
  const lastDataRef = useRef<Record<string, unknown> | null>(null);

  useEffect(() => {
    if (inProgress !== InteractionStatus.None) return;
    if (!isAuthenticated) return;
    const origin = new URL(import.meta.env.VITE_SP_SITE_URL || "https://placeholder.sharepoint.com").origin;
    instance.acquireTokenSilent({ scopes: [`${origin}/AllSites.Manage`], account: accounts[0] }).then(r => { tokenRef.current = r.accessToken; }).catch(() => {});
  }, [isAuthenticated, inProgress, instance, accounts]);

  useEffect(() => {
    if (!formId) { setError("No form slug provided."); setLoading(false); return; }

    const load = async () => {
      try {
        const origin = new URL(import.meta.env.VITE_SP_SITE_URL || "https://placeholder.sharepoint.com").origin;
        let token = tokenRef.current;

        // Try to acquire token if authenticated
        if (!token && isAuthenticated && accounts[0]) {
          try {
            const r = await instance.acquireTokenSilent({ scopes: [`${origin}/AllSites.Manage`], account: accounts[0] });
            token = r.accessToken;
            tokenRef.current = token;
          } catch (silentErr) {
            console.error("[DFP] acquireTokenSilent failed:", silentErr);
          }
        }

        if (token) {
          // Authenticated path — load directly from SharePoint
          let cfgRaw: Record<string, unknown>;
          let ver: { surveyJson: unknown; meta: unknown } | null;
          if (pinVersion) {
            const cfgRes = await fetch(`${SP_SITE_URL}/_api/web/lists/getbytitle('Master%20Form')/items?$filter=Slug eq '${encodeURIComponent(formId)}'&$select=Title,CurrentVersion,FormID,NumberOfApprovalLayer,Slug,IsPublic,ApprovalRules,ConditionField&$top=1`, { headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" } });
            if (!cfgRes.ok) throw new Error(`Failed to load form config: ${cfgRes.status} ${cfgRes.statusText}`);
            cfgRaw = (await cfgRes.json()).value?.[0];
            if (!cfgRaw) throw new Error(`Form "${formId}" not found.`);
            ver = await getFormVersion(token, cfgRaw.Title as string, pinVersion);
            if (!ver) throw new Error(`Version ${pinVersion} not found.`);
          } else {
            const latest = await getLatestFormBySlug(token, formId);
            if (!latest) throw new Error(`Form "${formId}" not found.`);
            cfgRaw = latest.formConfig as unknown as Record<string, unknown>;
            ver = { surveyJson: latest.surveyJson, meta: latest.meta };
          }
          setFormData({
            formConfig: cfgRaw,
            surveyJson: (ver?.surveyJson || null) as Record<string, unknown>,
            meta: (ver?.meta || {}) as Record<string, unknown>,
          });
        } else if (!isAuthenticated) {
          // Unauthenticated path — try public API fallback
          const res = await fetch(`/api/form-config?slug=${encodeURIComponent(formId)}${pinVersion ? `&version=${pinVersion}` : ""}`);
          const contentType = res.headers.get("content-type") || "";
          const responseText = await res.text();

          if (contentType.includes("text/html") || responseText.trim().startsWith("<")) {
            throw new Error("API endpoint not available (returned HTML). Are you running 'vercel dev'?");
          }

          // Detect if Vite served the raw TypeScript source instead of executing the API
          if (responseText.includes("export default async function") || responseText.includes('from "/api/_utils/')) {
            throw new Error("API route is returning source code instead of executing. Make sure you're running 'vercel dev' (not 'npm run dev').");
          }

          let parsed: { error?: string; formConfig?: Record<string, unknown>; surveyJson?: Record<string, unknown>; meta?: Record<string, unknown> };
          try {
            parsed = JSON.parse(responseText);
          } catch {
            throw new Error(`Server returned non-JSON: ${responseText.substring(0, 200)}`);
          }

          if (!res.ok) {
            throw new Error(parsed.error || `Server error: ${res.status}`);
          }

          if (!parsed.formConfig) {
            throw new Error("Invalid API response: missing formConfig.");
          }

          setFormData({
            formConfig: parsed.formConfig,
            surveyJson: (parsed.surveyJson || {}) as Record<string, unknown>,
            meta: (parsed.meta || {}) as Record<string, unknown>,
          });
        } else {
          // Authenticated but could not acquire token
          throw new Error("Unable to get authentication token. Please sign in again.");
        }
      } catch (e) {
        setError(e instanceof Error ? e.message : String(e));
      } finally {
        setLoading(false);
      }
    };

    load();
  }, [formId, pinVersion, isAuthenticated, instance, accounts]);

  // Enrich survey JSON with SharePoint-sourced choices
  useEffect(() => {
    const baseJson = formData?.surveyJson;
    if (!baseJson) { setEnrichedSurveyJson(null); return; }

    const tokenRaw = tokenRef.current;
    if (!tokenRaw) { setEnrichedSurveyJson(baseJson); return; }
    const token = tokenRaw; // narrowed to string

    const clone = JSON.parse(JSON.stringify(baseJson)) as Record<string, unknown>;

    async function enrich(): Promise<void> {
      const pages = (clone.pages || []) as { elements?: Record<string, unknown>[] }[];
      const pending: Promise<void>[] = [];

      function walk(elements: Record<string, unknown>[]) {
        for (const el of elements) {
          if (el.type === "panel" && Array.isArray(el.elements)) {
            walk(el.elements as Record<string, unknown>[]);
            continue;
          }

          // Main field spChoicesSource
          const src = el.spChoicesSource as { list?: string; column?: string } | undefined;
          if (src?.list && src?.column) {
            pending.push(
              getSharePointChoices(src.list, src.column, token)
                .then((choices) => {
                  if (choices.length > 0) el.choices = choices;
                })
                .catch(() => {})
            );
          }

          // Main field spFilteredListSource
          const fls = el.spFilteredListSource as { list?: string; valueColumn?: string; filterColumn?: string; filterValue?: string } | undefined;
          if (fls?.list && fls?.valueColumn) {
            pending.push(
              getFilteredListChoices(fls.list, fls.valueColumn, token, fls.filterColumn, fls.filterValue)
                .then((choices) => {
                  if (choices.length > 0) el.choices = choices;
                })
                .catch(() => {})
            );
          }

          // Matrix column choicesSource / filteredListSource
          if ((el.type === "matrixdynamic" || el.type === "dynamicmatrix") && Array.isArray(el.columns)) {
            const cols = el.columns as Record<string, unknown>[];
            for (const col of cols) {
              const colSrc = col.choicesSource as { list?: string; column?: string } | undefined;
              if (colSrc?.list && colSrc?.column) {
                pending.push(
                  getSharePointChoices(colSrc.list, colSrc.column, token)
                    .then((choices) => {
                      if (choices.length > 0) col.choices = choices;
                    })
                    .catch(() => {})
                );
              }
              const colFls = col.filteredListSource as { list?: string; valueColumn?: string; filterColumn?: string; filterValue?: string } | undefined;
              if (colFls?.list && colFls?.valueColumn) {
                pending.push(
                  getFilteredListChoices(colFls.list, colFls.valueColumn, token, colFls.filterColumn, colFls.filterValue)
                    .then((choices) => {
                      if (choices.length > 0) col.choices = choices;
                    })
                    .catch(() => {})
                );
              }
            }
          }
        }
      }

      for (const page of pages) {
        if (Array.isArray(page.elements)) walk(page.elements);
      }

      await Promise.all(pending);
      setEnrichedSurveyJson(clone);
    }

    enrich().catch(() => setEnrichedSurveyJson(baseJson));
  }, [formData]);

  const survey = useMemo(() => {
    const json = enrichedSurveyJson;
    if (!json) return null;
    try {
      const m = new Model(json);
      m.applyTheme(dark ? LayeredDarkPanelless : LayeredLightPanelless);
      m.showCompletedPage = false;
      m.onValueChanged.add((_, options) => {
        const q = m.getQuestionByName(options.name);
        if (!q || q.getType() !== "text") return;
        const mode = (q as Record<string, unknown>).autocapitalize as string | undefined;
        if (!mode || mode === "none") return;
        const val = options.value;
        if (typeof val !== "string") return;
        const transform = (v: string) => {
          switch (mode) {
            case "words": return v.replace(/\b\w/g, c => c.toUpperCase());
            case "sentences": return v.replace(/(^\w|[.!?]\s+\w)/g, c => c.toUpperCase());
            case "characters": return v.toUpperCase();
            default: return v;
          }
        };
        const next = transform(val);
        if (next !== val) {
          q.value = next;
        }
      });
      return m;
    } catch (e) { console.error("[DFP] Model error:", e); return null; }
  }, [enrichedSurveyJson, resetKey]);

  useEffect(() => { survey?.applyTheme(dark ? LayeredDarkPanelless : LayeredLightPanelless); }, [dark, survey]);

  const onCompleting = useCallback((sender: { data: Record<string, unknown> }) => { lastDataRef.current = { ...sender.data }; }, []);
  const onComplete = useCallback(async () => {
    setSubmitStatus("loading");
    const raw = lastDataRef.current ?? {};
    const cfg = formData?.formConfig;
    if (!cfg) { setSubmitStatus("error"); return; }
    try {
      let activeLayers: { email: string; name: string }[] = [];
      let resolvedLayerCount = 0;
      const token = tokenRef.current;
      const formId = String(cfg.FormID || "");

      // Upload signature images (keep existing code)
      if (token) {
        for (const [k, v] of Object.entries(raw)) {
          if (typeof v === "string" && v.startsWith("data:image/")) {
            try {
              const imageUrl = await uploadSignatureImage(token, formId, "submission", v);
              raw[k] = { Url: imageUrl, Description: "Signature" };
            } catch (e) {
              console.warn("[DFP] signature upload failed for", k, (e as Error).message);
            }
          }
        }
      }

      // Step 2: Resolve layers — try LayerConfig first, fall back to old rules
      let layerConfigParsed: LayerConfig | null = null;
      const rawLayerConfig = cfg.LayerConfig as string | undefined;
      if (rawLayerConfig && rawLayerConfig.trim()) {
        try { layerConfigParsed = JSON.parse(rawLayerConfig); } catch {}
      }

      if (layerConfigParsed?.layers?.length) {
        resolvedLayerCount = layerConfigParsed.layers.length;
        for (const layer of layerConfigParsed.layers) {
          if (layer.assignee.type === "user") {
            activeLayers.push({ email: layer.assignee.value, name: "" });
          } else {
            const fieldRef = layer.assignee.value.replace("${", "").replace("}", "");
            const fieldVal = String(raw[fieldRef] ?? "");
            activeLayers.push({ email: fieldVal, name: "" });
          }
        }
      } else {
        // Old approval rules / approvers list fallback (keep existing logic)
        let approvalRules = null;
        try { approvalRules = cfg.ApprovalRules ? JSON.parse(cfg.ApprovalRules as string) : null; } catch {}
        if (approvalRules?.conditionField && approvalRules?.rules?.length) {
          const condVal = String(raw[approvalRules.conditionField] ?? "").toLowerCase();
          const matched = approvalRules.rules.find((r: Record<string, unknown>) => (r.when as string).toLowerCase() === condVal);
          if (matched) {
            activeLayers = matched.layers;
            resolvedLayerCount = matched.layers.length;
          }
        } else if (token) {
          const apData = await spGet(token, `${SP_SITE_URL}/_api/web/lists/getbytitle('Approvers')/items?$filter=FormTitle eq '${encodeURIComponent(cfg.Title as string)}'&$select=LayerNumber,ApproverEmail,ApproverName&$orderby=LayerNumber asc&$top=10`).catch(() => ({ value: [] })) as { value: Record<string, string>[] };
          activeLayers = (apData.value ?? []).map((a) => ({ email: a.ApproverEmail, name: a.ApproverName }));
          resolvedLayerCount = activeLayers.length;
        }
      }

      // Step 3: Build body (keep existing logic)
      const body: Record<string, unknown> = {};
      for (const [k, v] of Object.entries(raw)) {
        if (v && typeof v === "object" && (v as Record<string, unknown>).html && (v as Record<string, unknown>).json) {
          body[`${k}_Response`] = (v as Record<string, unknown>).html;
          body[`${k}_Json`] = typeof (v as Record<string, unknown>).json === "string" ? (v as Record<string, unknown>).json : JSON.stringify((v as Record<string, unknown>).json);
        } else if (Array.isArray(v)) { body[k] = JSON.stringify(v); }
        else if (v && typeof v === "object") {
          if ("Url" in (v as Record<string, unknown>)) {
            body[k] = v;
          } else {
            body[k] = JSON.stringify(v);
          }
        }
        else if (typeof v === "number" || typeof v === "boolean") { body[k] = String(v); }
        else { body[k] = v; }
      }
      body.SubmittedAt = new Date().toISOString();
      body.FormVersion = cfg.CurrentVersion;
      body.FormID = cfg.FormID;

      // Step 4: Write layer status columns
      if (layerConfigParsed?.layers?.length) {
        // Enhanced path — use new constants
        for (let n = 1; n <= resolvedLayerCount; n++) {
          body[`L${n}_Status`] = SP_LAYER_STATUS.PENDING;
          body[`L${n}_Email`] = activeLayers[n - 1]?.email ?? "";
        }
        body.FormStatus = SP_FORM_STATUS.SUBMITTED;
        body.CurrentLayer = resolvedLayerCount > 0 ? 1 : 0;
      } else {
        // Legacy path — keep old behavior
        for (let n = 1; n <= resolvedLayerCount; n++) {
          body[`L${n}_Status`] = n === 1 ? "Pending" : "Waiting";
          body[`L${n}_Email`] = activeLayers[n - 1]?.email ?? "";
        }
      }

      // Step 5: Submit
      let submittedByEmail = "";
      if (token) {
        submittedByEmail = userEmail || accounts[0]?.username || "authenticated-user";
        body.SubmittedBy = submittedByEmail;
        const listUrl = `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(cfg.Title as string)}')/items`;
        let result: { Id?: number } | undefined;
        try {
          result = await spPost(token, listUrl, body) as { Id?: number };
        } catch (submitErr) {
          const msg = submitErr instanceof Error ? submitErr.message : String(submitErr);
          // If the response list is missing enhanced layer columns (pre-provisioning),
          // retry without FormStatus / CurrentLayer
          if ((msg.includes('FormStatus') || msg.includes('CurrentLayer')) && body.FormStatus !== undefined) {
            console.warn("[DFP] retrying without FormStatus/CurrentLayer (missing columns)");
            delete body.FormStatus;
            delete body.CurrentLayer;
            result = await spPost(token, listUrl, body) as { Id?: number };
          } else {
            throw submitErr;
          }
        }

        // Step 6: Trigger notification
        if (resolvedLayerCount > 0 && result?.Id) {
          const layer1Email = activeLayers[0]?.email;
          const formSlug = (cfg.Slug as string) || (cfg.slug as string) || "";
          const baseUrl = window.location.origin;

          if (layerConfigParsed?.layers?.[0]?.type === "evaluation" && layerConfigParsed.layers[0].authMode === "365" && layer1Email) {
            const reviewLink = formSlug
              ? `${baseUrl}/eval/${encodeURIComponent(formSlug)}/${result.Id}/1`
              : undefined;
            await triggerApprovalNotification(token, {
              formTitle: cfg.Title as string,
              submittedBy: submittedByEmail,
              responseItemId: result.Id,
              layer: 1,
              totalLayers: resolvedLayerCount,
              action: "submit",
              nextApproverEmail: layer1Email,
              reviewLink,
            }).catch((e: unknown) => console.warn("[DFP] notification skipped:", e instanceof Error ? e.message : String(e)));
          } else if (resolvedLayerCount > 0) {
            await triggerApprovalNotification(token, {
              formTitle: cfg.Title as string,
              submittedBy: submittedByEmail,
              responseItemId: result.Id,
              layer: 1,
              totalLayers: resolvedLayerCount,
              action: "submit",
              ...(layer1Email ? { nextApproverEmail: layer1Email } : {}),
            }).catch((e: unknown) => console.warn("[DFP] notification skipped:", e instanceof Error ? e.message : String(e)));
          }
        }

        // Step 7: Generate PDF for no-layers submission (immediate terminal state)
        if (resolvedLayerCount === 0 && result?.Id && token) {
          try {
            const cfgData = await getFormConfigByTitle(token, cfg.Title as string);
            const formVer = cfgData ? (cfgData as unknown as Record<string, unknown>).CurrentVersion as string || "1.0" : "1.0";
            const verData = await spGet(
              token,
              `${SP_SITE_URL}/_api/web/lists/getbytitle('Web%20Form%20Versions')/items?$filter=FormTitle eq '${encodeURIComponent(cfg.Title as string)}' and FormVersion eq '${encodeURIComponent(formVer)}'&$select=SurveyJSON&$top=1`
            ) as { value?: { SurveyJSON?: string }[] };
            const rawSurvey = verData.value?.[0]?.SurveyJSON;
            if (rawSurvey) {
              const parsed = JSON.parse(rawSurvey);
              const surveyContent = parsed.surveyJson || parsed;
              const respItem = await spGet(
                token,
                `${SP_SITE_URL}/_api/web/lists/getbytitle('${encodeURIComponent(cfg.Title as string)}')/items(${result.Id})`
              ) as Record<string, unknown>;
              const SYSTEM_FIELDS = new Set(['Id','Title','SubmittedBy','SubmittedAt','Status','CurrentApprovalLayer','FormVersion','FormID','RawJSON','CurrentLayer','FormStatus','EvaluationData','Author','Editor','Created','Modified','ContentType','PermMask','PdfUrl','L1_Status','L1_Email','L1_SignedAt','L1_Rejection','L1_Signature','L2_Status','L2_Email','L2_SignedAt','L2_Rejection','L2_Signature','L3_Status','L3_Email','L3_SignedAt','L3_Rejection','L3_Signature']);
              const pdfData: Record<string, unknown> = {};
              for (const [k, v] of Object.entries(respItem)) { if (!SYSTEM_FIELDS.has(k) && v !== null && v !== undefined) pdfData[k] = v; }
              await generateAndStorePdf(token, cfg.Title as string, result.Id, {
                surveyJson: surveyContent as PdfFormData["surveyJson"],
                responseData: pdfData,
                layerResults: buildPdfLayerResults(respItem),
                meta: { submittedBy: submittedByEmail, submittedAt: new Date().toISOString(), formTitle: cfg.Title as string, formVersion: formVer, formStatus: "submitted" },
                logoUrl: "/logo-128.png",
              });
            }
          } catch (pdfErr) { console.warn("[DFP] PDF generation failed:", pdfErr); }
        }
      } else {
        submittedByEmail = "GUEST";
        body.SubmittedBy = submittedByEmail;
        const res = await fetch("/api/submit-form", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ listTitle: cfg.Title, body }) });
        if (!res.ok) { const err = await res.json().catch(() => ({})) as { error?: string }; throw new Error(err.error || `Submit failed: ${res.status}`); }
      }
      setSubmitStatus("success");
    } catch (e) { console.error("[DFP] submit error:", e); setSubmitStatus("error"); }
  }, [formData, userEmail, accounts]);

  useEffect(() => {
    if (!survey) return;
    survey.onCompleting.add(onCompleting);
    survey.onComplete.add(onComplete);
    return () => { survey.onCompleting.remove(onCompleting); survey.onComplete.remove(onComplete); };
  }, [survey, onCompleting, onComplete]);

  const handleSignIn = useCallback(() => {
    try {
      sessionStorage.setItem("pmw_post_login_redirect", window.location.pathname + window.location.search);
    } catch {
      // May fail if storage is inaccessible
    }
    instance.loginRedirect({ ...loginRequest, redirectStartPage: window.location.href });
  }, [instance]);
  const handleSignOut = useCallback(() => {
    clearStoredAuthDecision();
    instance.logoutRedirect({ postLogoutRedirectUri: window.location.href });
  }, [instance]);
  const handleReset = useCallback(() => { setSubmitStatus(null); lastDataRef.current = null; setResetKey(k => k + 1); }, []);

  const isPublicForm = formData?.formConfig?.IsPublic !== false;
  const formTitle = String(formData?.formConfig?.Title || formData?.surveyJson?.title || "Form");

  useEffect(() => { document.title = formTitle ? `Form: ${formTitle}` : "Form — PMW HR Form"; }, [formTitle]);
  const formVersion = String(formData?.formConfig?.CurrentVersion || "1.0");
  const showBanner = (formData?.meta?.showBanner as boolean) !== false;
  const isoStandardsText = (formData?.meta?.isoStandards as string) || "ISO 9001 · ISO 14001 · ISO 45001";
  const companiesText = (formData?.meta?.companies as string) || "";
  const companyLines = companiesText.split("\n").filter(Boolean);
  const logoUrl = (formData?.meta?.logoUrl as string) || "";

  if (loading || (formData && !formData.surveyJson && !error)) return (
    <div style={{ minHeight: "100vh", background: t.bg, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 16 }}>
      <style>{globalCss(t)}</style>
      <Spinner t={t} />
      <div style={{ fontSize: 13, color: t.textMuted, animation: "pulse 1.5s infinite" }}>Loading form...</div>
    </div>
  );

  if (error) return (
    <div style={{ minHeight: "100vh", background: t.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <style>{globalCss(t)}</style>
      <div style={{ background: t.cardBg, borderRadius: 20, padding: "56px 44px", maxWidth: 420, textAlign: "center", boxShadow: t.shadowLg, border: `1px solid ${t.border}` }}>
        <div style={{ fontSize: 44, marginBottom: 18 }}>ERR</div>
        <div style={{ fontFamily: "'DM Serif Display',serif", fontSize: 22, color: t.red, marginBottom: 10 }}>Form not found</div>
        <p style={{ color: t.textSecond, fontSize: 13, lineHeight: 1.7 }}>{error}</p>
      </div>
    </div>
  );

  if (!isPublicForm && !isAuthenticated) return (<><style>{globalCss(t)}</style><PrivateGate formTitle={formTitle} onSignIn={handleSignIn} t={t} /></>);

  return (
    <div style={{ minHeight: "100vh", background: t.bg }}>
      <style>{globalCss(t)}</style>
      <ScrollProgress t={t} />
      <header style={{ background: t.cardBg, borderBottom: `1px solid ${t.border}`, height: 54, display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 24px", position: "sticky", top: 0, zIndex: 50 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <a href="/" title="Back to Dashboard" style={{ display: "flex", alignItems: "center", gap: 4, height: 28, padding: "0 10px", border: `1px solid ${t.border}`, borderRadius: 6, background: "none", color: t.textSecond, fontSize: 11, textDecoration: "none", cursor: "pointer", fontFamily: "'DM Sans'" }}>← Dashboard</a>
          <Logo size={28} />
          <span style={{ fontFamily: "'DM Serif Display',serif", fontSize: 15, color: t.textPrimary }}>{formTitle}</span>
          {pinVersion && <span style={{ fontSize: 10, fontWeight: 700, color: t.amber, background: t.amberPale, borderRadius: 20, padding: "2px 10px" }}>v{pinVersion}</span>}
          {!isPublicForm && <span style={{ fontSize: 10, fontWeight: 700, color: t.purple, background: t.purplePale, borderRadius: 20, padding: "2px 10px" }}>Private</span>}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {isAuthenticated ? (
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 11, color: t.green }}>
              <div style={{ width: 7, height: 7, borderRadius: "50%", background: t.green }} />
              {userEmail?.split("@")[0]}
            </div>
          ) : (<button onClick={handleSignIn} style={{ height: 28, padding: "0 13px", border: `1px solid ${t.purpleMid}`, borderRadius: 8, background: "none", color: t.purple, fontSize: 11, fontWeight: 600, cursor: "pointer", fontFamily: "'DM Sans'", display: "flex", alignItems: "center", gap: 6 }}><MsIcon /> Sign in</button>)}
          <span style={{ fontSize: 10, color: t.textMuted }}>v{formVersion}</span>
        </div>
      </header>

      {showBanner && (
        <div style={{ borderBottom: `1px solid ${t.border}`, background: t.cardBg }}>
          <div style={{ background: `linear-gradient(135deg,${t.purpleDark},${t.purple})`, padding: "16px 24px" }}>
            <div style={{ fontSize: 9, color: "rgba(255,255,255,0.5)", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 3 }}>{isoStandardsText}</div>
            <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 17, color: "#fff" }}>{formTitle}</div>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <tbody>
              <tr style={{ borderBottom: `1px solid ${t.border}` }}>
                <td style={{ width: 140, borderRight: `1px solid ${t.border}`, background: t.offWhite, padding: "9px 14px", verticalAlign: "middle", textAlign: "center" }}>
                  <img src={logoUrl || "/logo-48.png"} alt="Company Logo" style={{ maxWidth: "100%", maxHeight: 42, objectFit: "contain" }} />
                </td>
                <td style={{ padding: "12px 16px", fontWeight: 700, fontSize: 13, color: t.textPrimary }}>
                  {companyLines.length > 0
                    ? companyLines.map((line, i) => <div key={i} style={i > 0 ? { marginTop: 4 } : undefined}>{line}</div>)
                    : "PMW INTERNATIONAL BERHAD"}
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      )}

      <div style={{ maxWidth: 860, margin: "0 auto", padding: "28px 24px 88px", animation: "fadeUp .3s ease" }}>
        {submitStatus === "success" ? (
          <SuccessScreen formTitle={formTitle} onReset={handleReset} t={t} />
        ) : (
          <div>
            {!isPublicForm && isAuthenticated && (
              <div style={{ background: t.greenPale, border: `1px solid ${t.greenBorder}`, borderRadius: 12, padding: "12px 16px", marginBottom: 18, display: "flex", alignItems: "center", gap: 12 }}>
                <div style={{ width: 36, height: 36, borderRadius: "50%", background: `linear-gradient(135deg,${t.green},#34D399)`, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, fontWeight: 700 }}>{(userEmail?.[0] || "?").toUpperCase()}</div>
                <div style={{ flex: 1 }}><div style={{ fontSize: 12, fontWeight: 700, color: t.green }}>Submitting as yourself</div><div style={{ fontSize: 11, color: t.textSecond }}>{userEmail}</div></div>
                <button onClick={handleSignOut} style={{ fontSize: 11, color: t.textSecond, background: "none", border: `1px solid ${t.border}`, borderRadius: 7, padding: "5px 11px", cursor: "pointer", fontFamily: "'DM Sans'" }}>Sign out</button>
              </div>
            )}
            {survey ? <div className="dfp-survey-wrap"><Survey model={survey} /></div> : <div style={{ textAlign: "center", padding: 40, color: t.textMuted }}>Unable to render form.</div>}
            {submitStatus === "loading" && <div style={{ marginTop: 16, padding: "13px 16px", background: t.purplePale, border: `1px solid ${t.purpleMid}`, borderRadius: 10, color: t.purple, fontSize: 13 }}><Spinner size={14} t={t} /> Submitting...</div>}
            {submitStatus === "error" && <div style={{ marginTop: 16, padding: "13px 16px", background: t.redPale, border: "1px solid #FCA5A5", borderRadius: 10, color: t.red, fontSize: 13 }}>X Submission failed. Please try again.</div>}
          </div>
        )}
        <div style={{ marginTop: 32, textAlign: "center", fontSize: 11, color: t.textMuted }}>PMW International Berhad HR Forms</div>
      </div>
    </div>
  );
}