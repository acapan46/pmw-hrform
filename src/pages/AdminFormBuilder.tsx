/**
 * AdminFormBuilder.tsx — Full admin form builder with sidebar
 * Integrates with custom FormBuilder component
 */
import { useState, useEffect, useCallback, useRef, useMemo } from "react";
import type { ReactNode } from "react";
import { useNavigate, useParams } from "react-router-dom";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import FormBuilder from "../components/builder/FormBuilder";
import FormLibrary from "../components/builder/FormLibrary";
import VersionHistory from "../components/builder/VersionHistory";
import AuditLog from "../components/builder/AuditLog";
import ProvisionOverlay from "../components/builder/ProvisionOverlay";
import LayerConfigPanel from "../components/builder/LayerConfigPanel";
import { C } from "../components/builder/constants";
import { flattenQuestions, getSpColumnKind } from "../utils/FormBuilderEngine";
import { createSpClient } from "../utils/sharepointClient";
import { SP_STATIC } from "../utils/spConfig";
import type { SurveyJson, LayerConfig, LayerConfigItem } from "../types";

// MUI Icons
import ArrowBackIcon from "@mui/icons-material/ArrowBack";
import FolderIcon from "@mui/icons-material/Folder";
import SettingsIcon from "@mui/icons-material/Settings";
import RocketLaunchIcon from "@mui/icons-material/RocketLaunch";
import DescriptionIcon from "@mui/icons-material/Description";
import HistoryIcon from "@mui/icons-material/History";
import ReceiptLongIcon from "@mui/icons-material/ReceiptLong";
import ChevronLeftIcon from "@mui/icons-material/ChevronLeft";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";
import LayersIcon from "@mui/icons-material/Layers";
import SaveIcon from "@mui/icons-material/Save";
import EditNoteIcon from "@mui/icons-material/EditNote";
import CheckCircleIcon from "@mui/icons-material/CheckCircle";
import WarningIcon from "@mui/icons-material/Warning";
import CloseIcon from "@mui/icons-material/Close";
import PublicIcon from "@mui/icons-material/Public";
import BlockIcon from "@mui/icons-material/Block";
import LockIcon from "@mui/icons-material/Lock";

import {
  slugify,
  getAllFormConfigs,
  getFormConfig,
  upsertFormConfig,
  upsertApprovers,
  saveFormVersion,
  getFormVersionHistory,
  getFormVersion,
  logEvent,
  getFormLog,
  diffSurveyJson,
  isVersionGreater,
  incrementMinor,
  incrementMajor,
  bootstrapSystemLists,
  addColumn,
  listExists,
  createSpList,
  deleteForm,
  hardDeleteForm,
  getSharePointChoices,
} from "../utils/formBuilderSP";

const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");

// ── Design tokens ─────────────────────────────────────────────────────────────
const G = `*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;background:${C.offWhite};color:${C.textPrimary}}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
@keyframes spin{to{transform:rotate(360deg)}}@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
::-webkit-scrollbar{width:5px}::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:10px}`;
const inp = {
  width: "100%",
  height: 34,
  border: `1px solid ${C.border}`,
  borderRadius: 8,
  padding: "0 11px",
  fontSize: 13,
  fontFamily: "-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif",
  color: C.textPrimary,
  background: C.white,
  outline: "none",
};

// ── Simple Spinner component (inline) ─────────────────────────────────────
const Spinner = ({ size = 18 }: { size?: number }) => (
  <div style={{
    width: size,
    height: size,
    border: `2px solid #D1D5DB`,
    borderTop: `2px solid ${C.purple}`,
    borderRadius: "50%",
    animation: "spin 0.9s linear infinite",
    flexShrink: 0,
  }} />
);

// ── Inline helper components ──────────────────────────────────────────────────
const Tag = ({ children, color = C.purple, bg = C.purplePale }: { children: ReactNode; color?: string; bg?: string }) => (
  <span style={{ fontSize: 10, fontWeight: 700, color, background: bg, borderRadius: 20, padding: "2px 9px", textTransform: "uppercase", letterSpacing: ".04em" }}>{children}</span>
);

function TextInput({ value, onChange, placeholder, error, disabled, ...rest }: { value: string; onChange: (v: string) => void; placeholder?: string; error?: string; disabled?: boolean; [k: string]: unknown }) {
  const [f, setF] = useState(false);
  return (
    <>
      <input
        value={value ?? ""}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        disabled={disabled}
        style={{
          ...inp,
          borderColor: error ? C.red : f ? C.purple : C.border,
          boxShadow: f ? `0 0 0 3px ${error ? C.redPale : C.purplePale}` : "none",
          transition: "all .15s",
          opacity: disabled ? 0.6 : 1,
          cursor: disabled ? "not-allowed" : "text",
        }}
        onFocus={() => setF(true)}
        onBlur={() => setF(false)}
        {...rest}
      />
      {error && <div style={{ fontSize: 10, color: C.red, marginTop: 3 }}>{error}</div>}
    </>
  );
}

function FB({ label, hint, children, required }: { label: string; hint?: string; children: ReactNode; required?: boolean }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 11, fontWeight: 600, color: C.textPrimary, marginBottom: 4 }}>
        {label}{required && <span style={{ color: C.red, marginLeft: 3 }}>*</span>}
      </label>
      {hint && <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4, lineHeight: 1.5 }}>{hint}</div>}
      {children}
    </div>
  );
}

function ToggleSwitch({ checked, onChange, label }: { checked: boolean; onChange: (v: boolean) => void; label?: string }) {
  return (
    <label style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", userSelect: "none" }}>
      <div
        onClick={() => onChange(!checked)}
        style={{
          width: 36, height: 20, borderRadius: 10, flexShrink: 0,
          background: checked ? C.purple : "#D1D5DB", position: "relative",
          transition: "background 0.2s", cursor: "pointer",
        }}
      >
        <div style={{
          position: "absolute", top: 3, left: checked ? 19 : 3,
          width: 14, height: 14, borderRadius: "50%", background: C.white,
          transition: "left 0.2s", boxShadow: "0 1px 3px rgba(0,0,0,0.2)",
        }} />
      </div>
      {label && <span style={{ fontSize: 12, color: C.textSecond }}>{label}</span>}
    </label>
  );
}

// ── Main component ────────────────────────────────────────────────────────────
export default function AdminFormBuilder() {
  const navigate = useNavigate();
  const { formTitle: paramTitle } = useParams<{ formTitle: string }>();
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [authChecked, setAuthChecked] = useState(false);
  const [siteUsers, setSiteUsers] = useState<{ email: string; name: string }[]>([]);
  const tokenRef = useRef<string | null>(null);
  const [allForms, setAllForms] = useState<{ Id?: string; Title: string; FormID?: string; CurrentVersion?: string; Slug?: string; NumberOfApprovalLayer?: number; IsPublic?: boolean; IsPublished?: boolean; ApprovalRules?: string }[]>([]);
  const [libraryOpen, setLibraryOpen] = useState(true);
  const [isEditing, setIsEditing] = useState(false);
  const [originalVersion, setOriginalVersion] = useState<string | null>(null);
  const [meta, setMeta] = useState({
    formTitle: "",
    formId: "",
    formVersion: "1.0",
    slug: "",
    isoStandards: "ISO 9001 · ISO 14001 · ISO 45001",
    companies: "PMW INDUSTRIES SDN BHD\nPMW CONCRETE INDUSTRIES SDN BHD\nPMW LIGHTING INDUSTRIES SDN BHD\nPMW WINABUMI SDN BHD",
    logoUrl: "",
  });
  const [showBanner, setShowBanner] = useState(true);
  const [isPublic, setIsPublic] = useState(true);
  const setM = useCallback((k: string, v: string) => setMeta(m => ({ ...m, [k]: v })), []);
  const [slugError, setSlugError] = useState("");
  const [slugChecking, setSlugChecking] = useState(false);
  const [slugLocked, setSlugLocked] = useState(false);
  const [slugManual, setSlugManual] = useState(false);
  const [isDraft, setIsDraft] = useState(false);
  const [numLayers, setNumLayers] = useState(0);
  const [layers, setLayers] = useState<{ email: string; name: string }[]>(Array.from({ length: 5 }, () => ({ email: "", name: "" })));
  const [surveyJson, setSurveyJson] = useState<SurveyJson | null>(null);
  const [initialJson, setInitialJson] = useState<SurveyJson | null>(null);
  const prevSurveyRef = useRef<SurveyJson | null>(null);
  const [sidebarTab, setSidebarTab] = useState("meta");
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const sidebarTabsRef = useRef<HTMLDivElement>(null);
  const [sidebarWidth, setSidebarWidth] = useState(300);
  const resizingRef = useRef(false);
  const [canSidebarScrollLeft, setCanSidebarScrollLeft] = useState(false);
  const [canSidebarScrollRight, setCanSidebarScrollRight] = useState(false);
  const [versionHistory, setVersionHistory] = useState<{ FormVersion: string; PublishedBy?: string; PublishedAt?: string }[]>([]);
  const [viewingOld, setViewingOld] = useState<{ version: string; json: SurveyJson } | null>(null);
  const [newVersionMode, setNewVersionMode] = useState<"minor" | "major">("minor");
  const [auditLog, setAuditLog] = useState<{ EventType: string; EventSummary?: string; BeforeJSON?: string; AfterJSON?: string; EventAt?: string }[]>([]);
  const [logLoading, setLogLoading] = useState(false);
  const [provLogs, setProvLogs] = useState<{ m: string; t: string }[]>([]);
  const [provisioning, setProvisioning] = useState(false);
  const [provOk, setProvOk] = useState(false);
  const [provErr, setProvErr] = useState(false);
  const [toast, setToast] = useState<{ msg: string; type: string } | null>(null);
  const [accessDenied, setAccessDenied] = useState(false);
  const [deleteConfirm, setDeleteConfirm] = useState<{ Id?: string; Title: string } | null>(null);
  const [hardDeleteConfirm, setHardDeleteConfirm] = useState<{ Id?: string; Title: string } | null>(null);
  const [deleting, setDeleting] = useState(false);
  const [approvalRules, setApprovalRules] = useState<{ conditionField: string; rules: { when: string; layers: { email: string; name: string; role: string }[] }[] } | null>(null);
  const [layerConfig, setLayerConfig] = useState<LayerConfig | null>(null);

  const pLog = useCallback((m: string, t: string = "info") => setProvLogs(l => [...l, { m, t }]), []);
  const showToast = useCallback((msg: string, type: string = "info") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 4000);
  }, []);

  useEffect(() => {
    if (accessDenied) {
      showToast("Access denied — redirecting to dashboard. You need HR Form Owner permissions.", "err");
      setAccessDenied(false);
    }
  }, [accessDenied, showToast]);

  const checkSidebarScrollState = () => {
    const el = sidebarTabsRef.current;
    if (el) {
      setCanSidebarScrollLeft(el.scrollLeft > 0);
      setCanSidebarScrollRight(el.scrollLeft < el.scrollWidth - el.clientWidth - 1);
    }
  };

  const sidebarScrollLeft = () => {
    const el = sidebarTabsRef.current;
    if (el) el.scrollBy({ left: -200, behavior: 'smooth' });
  };

  const sidebarScrollRight = () => {
    const el = sidebarTabsRef.current;
    if (el) el.scrollBy({ left: 200, behavior: 'smooth' });
  };

  useEffect(() => {
    checkSidebarScrollState();
    window.addEventListener('resize', checkSidebarScrollState);
    return () => window.removeEventListener('resize', checkSidebarScrollState);
  }, []);

  // ── Sidebar resize via drag ────────────────────────────────────────────
  useEffect(() => {
    const onMouseMove = (e: MouseEvent) => {
      if (!resizingRef.current) return;
      // side = right sidebar, so dragging from main content area
      // width = distance from right edge of viewport to cursor
      const newWidth = Math.max(220, Math.min(800, window.innerWidth - e.clientX));
      setSidebarWidth(newWidth);
    };
    const onMouseUp = () => {
      resizingRef.current = false;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);
    return () => {
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
    };
  }, []);

  const proposedVersion = useMemo(() => {
    if (!isEditing || !originalVersion) return meta.formVersion;
    return newVersionMode === "minor" ? incrementMinor(originalVersion) : incrementMajor(originalVersion);
  }, [isEditing, originalVersion, newVersionMode, meta.formVersion]);

  useEffect(() => {
    if (slugManual || slugLocked) return;
    setM("slug", slugify(meta.formTitle));
  }, [meta.formTitle, slugManual, slugLocked, setM]);

  useEffect(() => {
    if (!meta.slug) {
      setSlugError("");
      return;
    }
    let cancelled = false;
    setSlugChecking(true);
    const t = setTimeout(() => {
      const slugToCheck = slugify(meta.slug);
      if (!slugToCheck) {
        if (!cancelled) { setSlugError(""); setSlugChecking(false); }
        return;
      }
      const others = allForms.filter(
        f => f.Slug && slugify(f.Slug) === slugToCheck && f.Title !== (isEditing ? meta.formTitle : null)
      );
      const conflict = others.length > 0 ? others[0].Title : null;
      if (!cancelled) {
        setSlugError(conflict ? `Used by: "${conflict}"` : "");
        setSlugChecking(false);
      }
    }, 600);
    return () => {
      cancelled = true;
      clearTimeout(t);
    };
  }, [meta.slug, isEditing, meta.formTitle, allForms]);

  useEffect(() => {
    if (inProgress !== InteractionStatus.None) return;
    if (!isAuthenticated) {
      navigate("/user/dashboard");
      return;
    }
    createSpClient(instance, accounts).isGroupMember(SP_STATIC.adminGroup).then(async admin => {
      if (!admin) {
        setAccessDenied(true);
        setTimeout(() => navigate("/user/dashboard"), 200);
        return;
      }
      setAuthChecked(true);
      const origin = new URL(import.meta.env.VITE_SP_SITE_URL || "https://placeholder.sharepoint.com").origin;
      try {
        const r = await instance.acquireTokenSilent({ scopes: [`${origin}/AllSites.Manage`], account: accounts[0] });
        tokenRef.current = r.accessToken;
        bootstrapSystemLists(r.accessToken, () => { }).catch(e => console.warn("[AFB] bootstrap:", e.message));
        try {
          const ud = await fetch(`${SP_SITE_URL}/_api/web/siteusers?$select=Email,Title&$filter=PrincipalType eq 1`, {
            headers: { Authorization: `Bearer ${r.accessToken}`, Accept: "application/json;odata=nometadata" },
          }).then(res => res.json());
          setSiteUsers((ud.value || []).filter((u: { Email: string }) => u.Email).map((u: { Email: string; Title: string }) => ({ email: u.Email, name: u.Title })));
        } catch { /* ignore */ }
        getAllFormConfigs(r.accessToken).then(setAllForms).catch(e => showToast(`Could not load forms: ${e.message}`, "err"));
      } catch (e) {
        console.error("[AFB] token:", e);
        showToast("Authentication error — please refresh.", "err");
      }
    }).catch(e => {
      console.warn("[AFB] init effect error:", e);
    });
  }, [isAuthenticated, inProgress, navigate, instance, accounts, showToast]);

  const refreshLib = useCallback(() => {
    setTimeout(() => {
      getAllFormConfigs(tokenRef.current!).then(setAllForms).catch(() => { /* ignore */ });
    }, 800);
  }, []);

  const loadForEdit = useCallback(async (cfg: Record<string, unknown> | string) => {
    try {
      const token = tokenRef.current;
      if (!token) return;
      const c = typeof cfg === "object" ? cfg : await getFormConfig(token, cfg);
      if (!c) {
        showToast(`Form not found.`, "err");
        return;
      }
      const data = await getFormVersion(token, c.Title as string, c.CurrentVersion as string || "1.0");
      if (!data) {
        showToast(`Version data not found.`, "err");
        return;
      }
      const loaded = (data.surveyJson || data) as SurveyJson;
      setInitialJson(loaded);
      prevSurveyRef.current = loaded;
      setViewingOld(null);
      setMeta({
        formTitle: c.Title as string,
        formId: (c.FormID as string) || "",
        formVersion: (c.CurrentVersion as string) || "1.0",
        slug: (c.Slug as string) || slugify(c.Title as string),
        isoStandards: (data.meta as Record<string, unknown>)?.isoStandards as string || "ISO 9001 · ISO 14001 · ISO 45001",
        companies: (data.meta as Record<string, unknown>)?.companies as string || "PMW INDUSTRIES SDN BHD\nPMW CONCRETE INDUSTRIES SDN BHD\nPMW LIGHTING INDUSTRIES SDN BHD\nPMW WINABUMI SDN BHD",
        logoUrl: ((data.meta as Record<string, unknown>)?.logoUrl as string) || "",
      });
      setShowBanner((data.meta as Record<string, unknown>)?.showBanner !== false);
      setOriginalVersion(c.CurrentVersion as string);
      setNumLayers((c.NumberOfApprovalLayer as number) || 0);
      setSlugLocked(true);
      setIsEditing(true);
      setIsDraft(c.IsPublished === false);
      setIsPublic(c.IsPublic !== false);
      if (c.ApprovalRules) {
        try {
          setApprovalRules(JSON.parse(c.ApprovalRules as string));
        } catch {
          setApprovalRules(null);
        }
      } else {
        setApprovalRules(null);
      }
      // Load enhanced LayerConfig if present, otherwise derive from legacy fields
      if (c.LayerConfig) {
        try {
          const parsed = JSON.parse(c.LayerConfig as string) as LayerConfig;
          setLayerConfig(parsed);
          // Derive numLayers from layerConfig for backward compat
          setNumLayers(parsed.layers.length);
        } catch {
          setLayerConfig(null);
        }
      } else {
        setLayerConfig(null);
      }
      getFormVersionHistory(token, c.Title as string).then(setVersionHistory).catch(() => {});
      setLogLoading(true);
      getFormLog(token, c.Title as string).then(l => {
        setAuditLog(l);
        setLogLoading(false);
      }).catch(() => { setLogLoading(false); });
    } catch (e) {
      console.warn("[AFB] loadForEdit error:", e);
      showToast(`Failed to load form: ${(e as Error).message}`, "err");
    }
  }, [showToast]);

  useEffect(() => {
    if (!paramTitle || !authChecked || !tokenRef.current) return;
    loadForEdit(decodeURIComponent(paramTitle)).catch(e => {
      console.warn("[AFB] URL param load error:", e);
    });
  }, [paramTitle, authChecked, loadForEdit]);

  const handleNew = () => {
    setIsEditing(false);
    setOriginalVersion(null);
    setInitialJson(null);
    prevSurveyRef.current = null;
    setSlugLocked(false);
    setSlugManual(false);
    setVersionHistory([]);
    setAuditLog([]);
    setViewingOld(null);
    setShowBanner(true);
    setMeta({
      formTitle: "",
      formId: "",
      formVersion: "1.0",
      slug: "",
      isoStandards: "ISO 9001 · ISO 14001 · ISO 45001",
      companies: "PMW INDUSTRIES SDN BHD\nPMW CONCRETE INDUSTRIES SDN BHD\nPMW LIGHTING INDUSTRIES SDN BHD\nPMW WINABUMI SDN BHD",
      logoUrl: "",
    });
    setNumLayers(0);
    setLayers(Array.from({ length: 5 }, () => ({ email: "", name: "" })));
    setLayerConfig(null);
    setIsDraft(false);
    setIsPublic(true);
    navigate("/admin/builder");
  };

  const handleSaveDraft = useCallback(async () => {
    if (!meta.formTitle.trim()) {
      showToast("Form title required.", "err");
      setSidebarTab("meta");
      return;
    }
    const token = tokenRef.current;
    if (!token) { showToast("Auth unavailable.", "err"); return; }
    const usedJson = surveyJson;
    if (!usedJson) { showToast("No fields yet.", "err"); return; }
    const version = "1.0";
    const userEmail = accounts[0]?.username || "admin";
    try {
      const layerConfigToSave = layerConfig || (numLayers > 0 ? {
        version: "1.0" as const,
        layers: layers.slice(0, numLayers).map((l, i): import("../types").LayerConfigItem => ({
          layerNumber: i + 1,
          type: "approval" as const,
          authMode: "365" as const,
          assignee: { type: "user" as const, value: l.email },
          title: `Layer ${i + 1}`,
          confirmationType: "signature" as const,
          allowRejectionReason: true,
        })),
      } : null);
      await upsertFormConfig(token, meta.formTitle.trim(), {
        formId: meta.formId.trim() || undefined,
        numLayers: layerConfig ? layerConfig.layers.length : numLayers,
        slug: meta.slug,
        version,
        isPublished: false,
        isPublic,
        conditionField: approvalRules?.conditionField || layerConfig?.routing?.[0]?.conditionField || "",
        approvalRules: approvalRules || null,
        layerConfig: layerConfigToSave ? JSON.stringify(layerConfigToSave) : "",
      });
      await saveFormVersion(token, {
        listTitle: meta.formTitle.trim(),
        slug: meta.slug,
        version,
        surveyJson: usedJson,
        meta: { isoStandards: meta.isoStandards, companies: meta.companies, formId: meta.formId, formVersion: version, showBanner, logoUrl: meta.logoUrl },
        changedBy: userEmail,
        layerConfig: layerConfig,
      });
      setIsDraft(true);
      setIsEditing(true);
      setOriginalVersion(version);
      setSlugLocked(true);
      setMeta(m => ({ ...m, formVersion: version }));
      showToast(`Draft saved for "${meta.formTitle}".`, "ok");
      refreshLib();
    } catch (e) {
      showToast(`Save draft failed: ${(e as Error).message}`, "err");
    }
  }, [meta, surveyJson, numLayers, layers, slugError, isPublic, showBanner, approvalRules, layerConfig, accounts, showToast, refreshLib]);

  const handleDelete = (f: { Id?: string; Title: string }) => {
    setDeleteConfirm({ Id: f.Id, Title: f.Title });
  };

  const handleHardDelete = (f: { Id?: string; Title: string }) => {
    setHardDeleteConfirm({ Id: f.Id, Title: f.Title });
  };

  const handleDeleteConfirm = async () => {
    if (!deleteConfirm || !tokenRef.current) return;
    setDeleting(true);
    try {
      const result = await deleteForm(
        tokenRef.current,
        deleteConfirm.Title,
        deleteConfirm.Id || "",
      );
      showToast(
        `Deleted "${deleteConfirm.Title}" — ${result.versionsDeleted} versions, ${result.logEntriesDeleted} log entries, ${result.approversDeleted} approvers removed.`,
        "ok"
      );
      if (meta.formTitle === deleteConfirm.Title) {
        handleNew();
      }
      refreshLib();
    } catch (e) {
      showToast(`Delete failed: ${(e as Error).message}`, "err");
    } finally {
      setDeleting(false);
      setDeleteConfirm(null);
    }
  };

  const handleHardDeleteConfirm = async () => {
    if (!hardDeleteConfirm || !tokenRef.current) return;
    setDeleting(true);
    try {
      const result = await hardDeleteForm(
        tokenRef.current,
        hardDeleteConfirm.Title,
        hardDeleteConfirm.Id || "",
      );
      const parts: string[] = [];
      if (result.responseListDeleted) {
        parts.push("response list deleted entirely");
      }
      parts.push(`${result.versionsDeleted} versions, ${result.logEntriesDeleted} log entries, ${result.approversDeleted} approvers removed`);
      showToast(
        `Hard-deleted "${hardDeleteConfirm.Title}" — ${parts.join("; ")}.`,
        "ok"
      );
      if (meta.formTitle === hardDeleteConfirm.Title) {
        handleNew();
      }
      refreshLib();
    } catch (e) {
      showToast(`Hard delete failed: ${(e as Error).message}`, "err");
    } finally {
      setDeleting(false);
      setHardDeleteConfirm(null);
    }
  };

  const handleViewVersion = async (ver: string) => {
    try {
      const data = await getFormVersion(tokenRef.current!, meta.formTitle, ver);
      if (!data) {
        showToast(`v${ver} not found.`, "err");
        return;
      }
      setViewingOld({ version: ver, json: (data.surveyJson || data) as SurveyJson });
      setSidebarTab("version");
    } catch (e) {
      showToast(`Failed: ${(e as Error).message}`, "err");
    }
  };

  useEffect(() => {
    if (sidebarTab !== "log" || !isEditing || !tokenRef.current) return;
    setLogLoading(true);
    getFormLog(tokenRef.current, meta.formTitle).then(l => {
      setAuditLog(l);
      setLogLoading(false);
    });
  }, [sidebarTab, isEditing, meta.formTitle]);

  const handlePublish = useCallback(async (jsonArg?: SurveyJson) => {
    if (!meta.formTitle.trim()) {
      showToast("Form title required.", "err");
      setSidebarTab("meta");
      return;
    }
    if (!meta.formId.trim()) {
      showToast("Form ID required.", "err");
      setSidebarTab("meta");
      return;
    }
    if (slugError) {
      showToast(`Slug conflict: ${slugError}`, "err");
      return;
    }
    const usedJson = jsonArg || surveyJson as SurveyJson;
    if (!usedJson) {
      showToast("No fields yet.", "err");
      return;
    }
    const token = tokenRef.current;
    if (!token) {
      showToast("Auth unavailable.", "err");
      return;
    }
    const version = isEditing && !isDraft ? proposedVersion : meta.formVersion;
    const title = meta.formTitle.trim();
    const userEmail = accounts[0]?.username || "admin";
    // Derive effective numLayers from layerConfig if present
    const effectiveNumLayers = layerConfig ? layerConfig.layers.length : numLayers;
    const activeLayers = layers.slice(0, effectiveNumLayers);
    setProvLogs([]);
    setProvOk(false);
    setProvErr(false);
    setProvisioning(true);
    try {
      const diffs = diffSurveyJson(prevSurveyRef.current, usedJson) as { type: string; summary: string; before: unknown; after: unknown }[];
      const slugToCheck = slugify(meta.slug);
      const others = allForms.filter(
        f => f.Slug && slugify(f.Slug) === slugToCheck && f.Title !== (isEditing ? title : null)
      );
      const conflict = others.length > 0 ? others[0].Title : null;
      if (conflict) throw new Error(`Slug "${meta.slug}" used by "${conflict}".`);
      if (isEditing && originalVersion && !isDraft && !isVersionGreater(version, originalVersion)) throw new Error(`v${version} must be > v${originalVersion}.`);

      pLog(`Checking list "${title}"…`);
      if (!(await listExists(token, title))) {
        pLog(`Creating list…`);
        await createSpList(token, title);
        pLog(`Created`, "ok");
      } else pLog(`List exists — checking columns`, "ok");

      pLog(`Provisioning columns…`);
      const questions = flattenQuestions(usedJson);
      for (const q of questions) {
        const sp = getSpColumnKind(q);
        if (q.type === "matrixdynamic" || q.type === "tableinput" || q.type === "dynamicmatrix") {
          await addColumn(token, title, `${q.name}_Response`, 3, true, true);
          pLog(`     ✓ ${q.name}_Response (Enhanced Rich Text)`);
          await addColumn(token, title, `${q.name}_Json`, 3, true, false);
          pLog(`     ✓ ${q.name}_Json (JSON backup)`);
        } else if (sp) {
          // Formula fields (_expression or expression type) get a Number column
          // so the calculated value is stored in SharePoint.
          const isFormula = !!(q as unknown as Record<string, unknown>)._expression || q.type === "expression";
          if (isFormula) {
            await addColumn(token, title, q.name, 9, false, false); // Number (kind 9)
            pLog(`     ✓ ${q.name} (Formula → Number)`);
            continue; // skip the normal column creation below
          }
          // Extract choices for Choice (6) and MultiChoice (15) columns
          let choiceValues: string[] | undefined;
          if (sp.FieldTypeKind === 6 || sp.FieldTypeKind === 15) {
            const src = (q as { spChoicesSource?: { list?: string; column?: string } }).spChoicesSource;
            if (src?.list && src?.column) {
              try {
                choiceValues = await getSharePointChoices(src.list, src.column, token);
                pLog(`     ↳ Fetched ${choiceValues.length} choices from "${src.list}.${src.column}"`);
              } catch { choiceValues = []; }
            }
            if (!choiceValues || choiceValues.length === 0) {
              const rawChoices = (q as { choices?: (string | { value: string; text: string })[] }).choices;
              if (Array.isArray(rawChoices) && rawChoices.length > 0) {
                choiceValues = rawChoices.map((c) => (typeof c === 'string' ? c : c.value || c.text || '')).filter(Boolean);
              }
            }
          }
          await addColumn(token, title, q.name, sp.FieldTypeKind, sp.FieldTypeKind === 3, false, choiceValues);
          pLog(`     ✓ ${q.name} (${sp.label})`);
        }
      }
      // Always-present system columns (legacy + enhanced)
      for (const [n, k] of [["SubmittedAt", 4], ["FormVersion", 2], ["FormID", 2], ["SubmittedBy", 2]] as [string, number][]) {
        await addColumn(token, title, n, k);
      }
      await addColumn(token, title, 'Status', 2); // Legacy status (compat with ApprovalDashboard)
      await addColumn(token, title, 'CurrentApprovalLayer', 9); // Legacy layer counter
      await addColumn(token, title, 'RawJSON', 3, true); // Full survey JSON backup
      await addColumn(token, title, 'PdfUrl', 2); // Generated PDF link
      // Per-layer columns (provision at least 3 layers for ApprovalDashboard compat)
      const provisionLayerCount = Math.max(effectiveNumLayers, 3);
      for (let n = 1; n <= provisionLayerCount; n++) {
        for (const [col, k] of [["Status", 2], ["Email", 2], ["SignedAt", 4], ["Rejection", 3], ["Signature", 3]] as [string, number][]) {
          await addColumn(token, title, `L${n}_${col}`, k, k === 3);
        }
      }
      // Enhanced layer system columns (added when layers are present)
      if (effectiveNumLayers > 0) {
        await addColumn(token, title, 'EvaluationData', 3, true);
        await addColumn(token, title, 'CurrentLayer', 9);
        await addColumn(token, title, 'FormStatus', 2);
      }
      pLog(`Columns done`, "ok");

      pLog(`Updating Form Config…`);
      // Build LayerConfig JSON from UI state if not already set
      const layerConfigToSave = layerConfig || (effectiveNumLayers > 0 ? {
        version: "1.0" as const,
        layers: activeLayers.map((l, i): LayerConfigItem => ({
          layerNumber: i + 1,
          type: "approval" as const,
          authMode: "365" as const,
          assignee: { type: "user" as const, value: l.email },
          title: `Layer ${i + 1}`,
          confirmationType: "signature" as const,
          allowRejectionReason: true,
        })),
      } : null);
      await upsertFormConfig(token, title, {
        formId: meta.formId.trim(),
        numLayers: effectiveNumLayers,
        slug: meta.slug,
        version,
        isPublished: true,
        isPublic,
        conditionField: approvalRules?.conditionField || layerConfig?.routing?.[0]?.conditionField || "",
        approvalRules: approvalRules || null,
        layerConfig: layerConfigToSave ? JSON.stringify(layerConfigToSave) : "",
      });
      pLog(`Form Config saved`, "ok");

      if (effectiveNumLayers > 0) {
        pLog(`Writing approvers…`);
        // When using the new LayerConfigPanel, layer assignee emails are stored
        // in layerConfig.layers[].assignee.value (type: "user") or as field references.
        // The old `activeLayers` (from `layers` state) is empty in that case,
        // so we must extract from layerConfig instead.
        const approversToWrite = layerConfig
          ? layerConfig.layers.map((l) => ({
              email: l.assignee.type === "user" ? l.assignee.value : "",
              name: l.title ?? "",
            }))
          : activeLayers;
        await upsertApprovers(token, title, approversToWrite);
        pLog(`Approvers saved`, "ok");
      }

      pLog(`Saving version v${version}…`);
      await saveFormVersion(token, {
        listTitle: title,
        slug: meta.slug,
        version,
        surveyJson: usedJson,
        meta: { isoStandards: meta.isoStandards, companies: meta.companies, formId: meta.formId, formVersion: version, showBanner, logoUrl: meta.logoUrl },
        changedBy: userEmail,
        layerConfig: layerConfig,
      });
      pLog(`Version saved`, "ok");

      if (!isEditing) {
        await logEvent(token, {
          formTitle: title,
          eventType: "FORM_CREATED",
          changedBy: userEmail,
          summary: `Created. Route: /form/${meta.slug}`,
          before: null,
          after: { slug: meta.slug, version },
        });
      } else {
        for (const d of diffs) {
          await logEvent(token, {
            formTitle: title,
            eventType: d.type,
            changedBy: userEmail,
            summary: d.summary,
            before: d.before,
            after: d.after,
          });
        }
        await logEvent(token, {
          formTitle: title,
          eventType: "VERSION_BUMPED",
          changedBy: userEmail,
          summary: `v${originalVersion} → v${version}`,
          before: { version: originalVersion },
          after: { version },
        });
        await logEvent(token, {
          formTitle: title,
          eventType: "PUBLISHED",
          changedBy: userEmail,
          summary: `Published v${version}`,
          before: null,
          after: { version, slug: meta.slug },
        });
      }
        pLog(`✓ "${title}" v${version} live at /form/${meta.slug}`, "ok");
      setProvOk(true);
      prevSurveyRef.current = usedJson;
      setOriginalVersion(version);
      setMeta(m => ({ ...m, formVersion: version }));
      setIsEditing(true);
      setIsDraft(false);
      setSlugLocked(true);
      refreshLib();
      getFormVersionHistory(token, title).then(setVersionHistory);
    } catch (e) {
      pLog(`Error: ${(e as Error).message}`, "err");
      setProvErr(true);
    }
  }, [meta, surveyJson, numLayers, layers, isEditing, originalVersion, proposedVersion, slugError, isPublic, showBanner, pLog, refreshLib, approvalRules, layerConfig, accounts, showToast]);

  if (!authChecked) {
    return (
      <div style={{ minHeight: "100vh", background: C.offWhite, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <style>{G}</style>
        <Spinner size={36} />
      </div>
    );
  }

  const sidebarTabs = [
    { id: "meta", label: "Meta", icon: <DescriptionIcon style={{ fontSize: 14 }} /> },
    { id: "layers", label: "Layers", icon: <LayersIcon style={{ fontSize: 14 }} /> },
    { id: "version", label: "Versions", icon: <HistoryIcon style={{ fontSize: 14 }} /> },
    { id: "log", label: "Log", icon: <ReceiptLongIcon style={{ fontSize: 14 }} /> },
    { id: "publish", label: "Publish", icon: <RocketLaunchIcon style={{ fontSize: 14 }} /> },
  ];

  const formBuilderKey = viewingOld
    ? `view_${meta.formTitle}_v${viewingOld.version}`
    : initialJson
      ? `edit_${meta.formTitle}_${JSON.stringify(initialJson).slice(0, 60)}`
      : "new";

  return (
    <div style={{ minHeight: "100vh", background: C.offWhite, display: "flex", flexDirection: "column" }}>
      <style>{G}</style>
      {toast && (
        <div style={{
          position: "fixed",
          top: 16,
          right: 16,
          zIndex: 10000,
          background: toast.type === "err" ? C.red : toast.type === "ok" ? C.green : C.purple,
          color: C.white,
          padding: "10px 18px",
          borderRadius: 10,
          fontSize: 12,
          boxShadow: C.shadowMd,
          animation: "fadeUp .2s ease",
          maxWidth: 360,
        }}>
          {toast.msg}
        </div>
      )}

      <header style={{
        background: C.white,
        borderBottom: `1px solid ${C.border}`,
        height: 52,
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        padding: "0 18px",
        position: "sticky",
        top: 0,
        zIndex: 100,
        boxShadow: "0 1px 0 rgba(91,33,182,.06)",
        flexShrink: 0,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <button
            type="button"
            onClick={() => { window.location.assign("/"); }}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 4,
              height: 30,
              border: `1px solid ${C.border}`,
              borderRadius: 7,
              background: C.white,
              color: C.textSecond,
              fontSize: 12,
              cursor: "pointer",
              padding: "0 12px",
              fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
            }}
            onMouseEnter={(e) => { e.currentTarget.style.background = "#F3F4F6"; }}
            onMouseLeave={(e) => { e.currentTarget.style.background = C.white; }}
          >
            <ArrowBackIcon style={{ fontSize: 14 }} /> Dashboard
          </button>
          <div style={{ width: 1, height: 17, background: C.border }} />
          <span style={{ fontSize: 18, color: '#6264A7', display: 'inline-flex' }}><DescriptionIcon style={{ fontSize: 18 }} /></span>
          <span style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 16, color: C.textPrimary }}>
            {isEditing ? `Editing: ${meta.formTitle}` : "New Form"}
          </span>
          <Tag color={C.amber} bg={C.amberPale}>⚙ Admin</Tag>
          {isEditing && <Tag>v{meta.formVersion}</Tag>}
          {isDraft && <Tag color={C.amber} bg={C.amberPale}>Draft</Tag>}
          {meta.slug && (
            <a
                  href={`/form/${meta.slug}`}
              target="_blank"
              rel="noreferrer"
              style={{
                fontSize: 10,
                color: C.purple,
                textDecoration: "none",
                background: C.purplePale,
                borderRadius: 20,
                padding: "2px 9px",
                border: `1px solid ${C.purpleMid}`,
              }}
            >
                    /form/{meta.slug} ↗
            </a>
          )}
        </div>
        <div style={{ display: "flex", gap: 7 }}>
          <button
            onClick={() => setLibraryOpen(o => !o)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              height: 30,
              padding: "0 12px",
              border: `1px solid ${C.border}`,
              borderRadius: 7,
              background: libraryOpen ? C.purplePale : C.white,
              color: libraryOpen ? C.purple : C.textSecond,
              fontSize: 12,
              cursor: "pointer",
              fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
            }}
            onMouseEnter={(e) => { if (!libraryOpen) { e.currentTarget.style.background = "#F9FAFB"; e.currentTarget.style.borderColor = C.purple; e.currentTarget.style.color = C.purple; } }}
            onMouseLeave={(e) => { if (!libraryOpen) { e.currentTarget.style.background = C.white; e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.textSecond; } }}
          >
            <FolderIcon style={{ fontSize: 14 }} /> Forms
          </button>
          <button
            onClick={() => setSidebarOpen(o => !o)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              height: 30,
              padding: "0 12px",
              border: `1px solid ${C.border}`,
              borderRadius: 7,
              background: sidebarOpen ? C.purplePale : C.white,
              color: sidebarOpen ? C.purple : C.textSecond,
              fontSize: 12,
              cursor: "pointer",
              fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
            }}
            onMouseEnter={(e) => { if (!sidebarOpen) { e.currentTarget.style.background = "#F9FAFB"; e.currentTarget.style.borderColor = C.purple; e.currentTarget.style.color = C.purple; } }}
            onMouseLeave={(e) => { if (!sidebarOpen) { e.currentTarget.style.background = C.white; e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.textSecond; } }}
          >
            <SettingsIcon style={{ fontSize: 14 }} /> Settings
          </button>
          {(!isEditing || isDraft) && (
            <button
              onClick={() => handleSaveDraft()}
              disabled={!meta.formTitle.trim() || !!viewingOld}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                height: 30,
                padding: "0 14px",
                border: `1px solid ${C.border}`,
                borderRadius: 7,
                background: C.white,
                color: C.textSecond,
                fontSize: 12,
                fontWeight: 600,
                cursor: !meta.formTitle.trim() || viewingOld ? "not-allowed" : "pointer",
                fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                opacity: !meta.formTitle.trim() || viewingOld ? 0.5 : 1,
              }}
              onMouseEnter={(e) => { if (!viewingOld) e.currentTarget.style.background = C.offWhite; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = C.white; }}
            >
              <SaveIcon style={{ fontSize: 14, marginRight: 4 }} /> Save Draft
            </button>
          )}
          <button
            onClick={() => handlePublish(surveyJson as SurveyJson)}
            disabled={!!viewingOld}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              height: 30,
              padding: "0 16px",
              border: "none",
              borderRadius: 7,
              background: viewingOld ? C.border : `linear-gradient(135deg,${C.purple},${C.purpleLight})`,
              color: C.white,
              fontSize: 12,
              fontWeight: 600,
              cursor: viewingOld ? "not-allowed" : "pointer",
              fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
              boxShadow: viewingOld ? "none" : "0 2px 8px rgba(91,33,182,.25)",
              opacity: viewingOld ? 0.5 : 1,
            }}
            onMouseEnter={(e) => { if (!viewingOld) e.currentTarget.style.boxShadow = "0 4px 12px rgba(91,33,182,.35)"; }}
            onMouseLeave={(e) => { if (!viewingOld) e.currentTarget.style.boxShadow = "0 2px 8px rgba(91,33,182,.25)"; }}
          >
            <RocketLaunchIcon style={{ fontSize: 14 }} /> Publish
          </button>
        </div>
      </header>

      {viewingOld && (
        <div style={{ background: C.amberPale, borderBottom: "1px solid #FDE68A", padding: "7px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", fontSize: 12, color: C.amber }}>
          <span>👁 Viewing archived <strong>v{viewingOld.version}</strong> — read only</span>
          <button
            onClick={() => setViewingOld(null)}
            style={{ background: "none", border: "none", color: C.amber, cursor: "pointer", fontWeight: 600, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}
          >
            <CloseIcon style={{ fontSize: 14, marginRight: 4 }} /> Back to current
          </button>
        </div>
      )}

      <div style={{ flex: 1, display: "flex", overflow: "hidden", height: "calc(100vh - 52px)" }}>
        {libraryOpen && (
          <div style={{ width: 215, flexShrink: 0, borderRight: `1px solid ${C.border}`, background: C.white, overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <FormLibrary forms={allForms} onEdit={loadForEdit} onNew={handleNew} onDelete={handleDelete} onHardDelete={handleHardDelete} current={meta.formTitle} />
          </div>
        )}

        <div style={{ flex: 1, overflow: "hidden" }}>
          <FormBuilder
            key={formBuilderKey}
            initialJson={viewingOld?.json || initialJson}
            onChange={json => {
              if (!viewingOld) setSurveyJson(json);
            }}

            height="100%"
            readOnly={!!viewingOld}
            token={tokenRef.current || undefined}
            showBanner={showBanner}
            meta={{ isoStandards: meta.isoStandards, companies: meta.companies, formTitle: meta.formTitle, logoUrl: meta.logoUrl }}
          />
        </div>

        {sidebarOpen && (
          <div style={{ width: sidebarWidth, flexShrink: 0, borderLeft: `1px solid ${C.border}`, background: C.white, display: "flex", flexDirection: "column", overflow: "hidden", position: "relative" }}>
            {/* Drag handle for resizing */}
            <div
              onMouseDown={() => {
                resizingRef.current = true;
                document.body.style.cursor = "ew-resize";
                document.body.style.userSelect = "none";
              }}
              style={{
                position: "absolute",
                left: 0,
                top: 0,
                bottom: 0,
                width: 4,
                cursor: "ew-resize",
                zIndex: 10,
              }}
            />
            <div style={{ display: "flex", alignItems: "center", borderBottom: `1px solid ${C.border}` }}>
              <button
                onClick={sidebarScrollLeft}
                style={{
                  visibility: canSidebarScrollLeft ? 'visible' : 'hidden',
                  flexShrink: 0,
                  width: 24,
                  height: 24,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: C.white,
                  border: `1px solid ${C.border}`,
                  borderRadius: 4,
                  cursor: "pointer",
                  marginLeft: 4,
                  color: C.textSecond,
                }}
              >
                <ChevronLeftIcon style={{ fontSize: 16 }} />
              </button>
              <div
                ref={sidebarTabsRef}
                onScroll={checkSidebarScrollState}
                style={{ display: "flex", overflowX: "auto", gap: 6, padding: "8px 6px", scrollbarWidth: "none", flex: 1 }}
              >
                {sidebarTabs.map(t => (
                  <button
                    key={t.id}
                    onClick={() => setSidebarTab(t.id)}
                    style={{
                      flex: "0 0 auto",
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                      height: 30,
                      border: `1px solid ${sidebarTab === t.id ? C.purple : C.border}`,
                      borderRadius: 7,
                      background: sidebarTab === t.id ? C.purplePale : C.white,
                      color: sidebarTab === t.id ? C.purple : C.textSecond,
                      fontSize: 12,
                      fontWeight: sidebarTab === t.id ? 600 : 400,
                      cursor: "pointer",
                      padding: "0 12px",
                      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                      whiteSpace: "nowrap",
                    }}
                    onMouseEnter={(e) => { if (sidebarTab !== t.id) { e.currentTarget.style.background = "#F9FAFB"; e.currentTarget.style.borderColor = C.purple; e.currentTarget.style.color = C.purple; } }}
                    onMouseLeave={(e) => { if (sidebarTab !== t.id) { e.currentTarget.style.background = C.white; e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.textSecond; } }}
                  >
                    {t.icon} {t.label}
                  </button>
                ))}
              </div>
              <button
                onClick={sidebarScrollRight}
                style={{
                  visibility: canSidebarScrollRight ? 'visible' : 'hidden',
                  flexShrink: 0,
                  width: 24,
                  height: 24,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  background: C.white,
                  border: `1px solid ${C.border}`,
                  borderRadius: 4,
                  cursor: "pointer",
                  marginRight: 4,
                  color: C.textSecond,
                }}
              >
                <ChevronRightIcon style={{ fontSize: 16 }} />
              </button>
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: "13px 13px 24px" }}>
              {sidebarTab === "meta" && (
                <div style={{ animation: "fadeUp .15s ease" }}>
                  <FB label="Form Title" hint="Becomes the SP list name. Locked after first publish." required>
                    <TextInput
                      value={meta.formTitle}
                      onChange={v => setM("formTitle", v)}
                      placeholder="Training Application Form"
                      disabled={isEditing}
                    />
                    {isEditing && <div style={{ fontSize: 10, color: C.textMuted, marginTop: 3 }}>Locked after first publish</div>}
                  </FB>
                  <FB label="Form ID / Doc No." required>
                    <TextInput value={meta.formId} onChange={v => setM("formId", v)} placeholder="PMW-HR-001" />
                  </FB>
                  <FB label="Version" hint={isEditing && !isDraft ? `Current: v${originalVersion} → New: v${proposedVersion}` : isEditing ? `v${meta.formVersion} (draft)` : undefined}>
                    {isEditing ? (
                      <div style={{ display: "flex", gap: 7 }}>
                        {(["minor", "major"] as const).map(m => (
                          <button
                            key={m}
                            onClick={() => setNewVersionMode(m)}
                            style={{
                              flex: 1,
                              height: 32,
                              borderRadius: 7,
                              border: `1px solid ${newVersionMode === m ? C.purple : C.border}`,
                              background: newVersionMode === m ? C.purplePale : C.white,
                              color: newVersionMode === m ? C.purple : C.textSecond,
                              fontSize: 12,
                              cursor: "pointer",
                              fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                              fontWeight: newVersionMode === m ? 600 : 400,
                            }}
                          >
                            {m === "minor" ? `v${incrementMinor(originalVersion!)}` : `v${incrementMajor(originalVersion!)}`}
                            <div style={{ fontSize: 9, color: C.textMuted, marginTop: 1 }}>{m} bump</div>
                          </button>
                        ))}
                      </div>
                    ) : (
                      <TextInput value={meta.formVersion} onChange={v => setM("formVersion", v)} placeholder="1.0" />
                    )}
                  </FB>
                  <FB label="Route slug" hint="URL: /form/{slug}. Locked after first publish.">
                    <div style={{ position: "relative" }}>
                      <TextInput
                        value={meta.slug}
                        onChange={v => {
                          setM("slug", slugify(v));
                          setSlugManual(true);
                        }}
                        placeholder="training-application"
                        disabled={slugLocked}
                        error={slugError}
                      />
                      {slugChecking && (
                        <div style={{ position: "absolute", right: 9, top: "50%", transform: "translateY(-50%)" }}>
                          <Spinner size={12} />
                        </div>
                      )}
                      {!slugError && !slugChecking && meta.slug && !slugLocked && (
                        <div style={{ fontSize: 10, color: C.green, marginTop: 3 }}>✓ Slug available</div>
                      )}
                      {slugLocked && <div style={{ fontSize: 10, color: C.textMuted, marginTop: 3 }}>Locked — cannot change after publish</div>}
                    </div>
                  </FB>
                  <FB label="ISO Standards">
                    <TextInput value={meta.isoStandards} onChange={v => setM("isoStandards", v)} />
                  </FB>
                  <FB label="Companies" hint="One per line">
                    <textarea
                      value={meta.companies}
                      onChange={e => setM("companies", e.target.value)}
                      rows={4}
                      style={{ ...inp, height: "auto", padding: "7px 10px", resize: "vertical", lineHeight: 1.7 }}
                    />
                  </FB>
                  <FB label="Logo URL" hint="Custom logo URL for the banner (defaults to /logo-128.png)">
                    <TextInput value={meta.logoUrl} onChange={v => setM("logoUrl", v)} placeholder="https://example.com/logo.png" />
                  </FB>
                  <FB label="Header / Companies banner" hint="Show or hide the ISO + company banner at the top of the form.">
                    <div style={{
                      background: showBanner ? C.greenPale : C.offWhite,
                      border: `1px solid ${showBanner ? "#6EE7B7" : C.border}`,
                      borderRadius: 8,
                      padding: "10px 12px",
                      transition: "all .2s",
                    }}>
                      <ToggleSwitch checked={showBanner} onChange={setShowBanner} label={showBanner ? "Banner visible" : "Banner hidden"} />
                      <div style={{ fontSize: 10, color: C.textMuted, marginTop: 6, lineHeight: 1.5 }}>
                        {showBanner ? "ISO standards and company names appear at the top of the form." : "Form opens without the header banner."}
                      </div>
                    </div>
                  </FB>
                  <FB label="Form Access">
                    <div style={{ display: "flex", gap: 7 }}>
                      {[
                        { v: true, label: <><PublicIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 2 }} /> Public</>, hint: "Any M365 user" },
                        { v: false, label: <><LockIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 2 }} /> Private</>, hint: "Explicit sign-in gate" },
                      ].map(opt => (
                        <button
                          key={String(opt.v)}
                          onClick={() => setIsPublic(opt.v)}
                          style={{
                            flex: 1,
                            padding: "8px 4px",
                            borderRadius: 8,
                            cursor: "pointer",
                            border: `1.5px solid ${isPublic === opt.v ? C.purple : C.border}`,
                            background: isPublic === opt.v ? C.purplePale : C.white,
                            color: isPublic === opt.v ? C.purple : C.textSecond,
                            fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                            transition: "all .13s",
                          }}
                        >
                          <div style={{ fontSize: 12, fontWeight: 600 }}>{opt.label}</div>
                          <div style={{ fontSize: 9, color: C.textMuted, marginTop: 2 }}>{opt.hint}</div>
                        </button>
                      ))}
                    </div>
                  </FB>
                </div>
              )}

              {sidebarTab === "layers" && (
                <LayerConfigPanel
                  value={layerConfig}
                  onChange={setLayerConfig}
                  siteUsers={siteUsers}
                  formFieldNames={surveyJson ? flattenQuestions(surveyJson).map(q => q.name) : []}
                  slug={meta.slug}
                />
              )}

              {sidebarTab === "version" && (
                <div style={{ animation: "fadeUp .15s ease" }}>
                  {!isEditing ? (
                    <div style={{ fontSize: 11, color: C.textMuted, fontStyle: "italic" }}>
                      Publish a form first to see version history.
                    </div>
                  ) : (
                    <>
                      {viewingOld && (
                        <div style={{
                          background: C.amberPale,
                          border: "1px solid #FDE68A",
                          borderRadius: 8,
                          padding: "8px 11px",
                          fontSize: 11,
                          color: C.amber,
                          marginBottom: 12,
                          display: "flex",
                          justifyContent: "space-between",
                          alignItems: "center",
                        }}>
                          <span>Viewing v{viewingOld.version}</span>
                          <button
                            onClick={() => setViewingOld(null)}
                            style={{ background: "none", border: "none", color: C.amber, cursor: "pointer", fontWeight: 600, fontSize: 11 }}
                          >
                            ← Back to current
                          </button>
                        </div>
                      )}
                      <VersionHistory history={versionHistory} current={originalVersion || ""} onView={handleViewVersion} />
                    </>
                  )}
                </div>
              )}

              {sidebarTab === "log" && (
                <div style={{ animation: "fadeUp .15s ease" }}>
                  {logLoading ? (
                    <div style={{ display: "flex", justifyContent: "center", padding: 20 }}>
                      <Spinner />
                    </div>
                  ) : (
                    <AuditLog logs={auditLog} />
                  )}
                </div>
              )}

              {sidebarTab === "publish" && (
                <div style={{ animation: "fadeUp .15s ease" }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 9, marginBottom: 16 }}>
                    {[
                      ["Form", meta.formTitle || <em style={{ color: C.red }}>Missing <WarningIcon style={{ fontSize: 12, verticalAlign: 'middle' }} /></em>],
                      ["Form ID", meta.formId || <em style={{ color: C.red }}>Missing <WarningIcon style={{ fontSize: 12, verticalAlign: 'middle' }} /></em>],
                      ["Version", isEditing && !isDraft ? `${originalVersion} → ${proposedVersion}` : `v${meta.formVersion}${isDraft ? " (draft)" : ""}`],
                      ["Status", isDraft ? <span style={{ color: C.amber }}><EditNoteIcon style={{ fontSize: 14, verticalAlign: 'middle', marginRight: 2 }} /> Draft</span> : isEditing ? <span style={{ color: C.green }}><CheckCircleIcon style={{ fontSize: 14, verticalAlign: 'middle', marginRight: 2 }} /> Published</span> : "—"],
                      ["Route", meta.slug ? `/form/${meta.slug}` : <em style={{ color: C.amber }}>No slug</em>],
                      ["Layers", numLayers || "None"],
                      ["Banner", showBanner ? <><CheckCircleIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 2 }} /> Visible</> : <><BlockIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 2 }} /> Hidden</>],
                      ["Access", isPublic ? <><PublicIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 2 }} /> Public</> : <><LockIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 2 }} /> Private</>],
                    ].map(([k, v]) => (
                      <div key={k as string} style={{ display: "flex", gap: 10, fontSize: 12 }}>
                        <span style={{ color: C.textMuted, minWidth: 70 }}>{k as string}:</span>
                        <span style={{ color: C.textPrimary, fontWeight: 500 }}>{v as ReactNode}</span>
                      </div>
                    ))}
                  </div>
                  {slugError && (
                    <div style={{ background: C.redPale, border: "1px solid #FCA5A5", borderRadius: 8, padding: "7px 10px", fontSize: 11, color: C.red, marginBottom: 10 }}>
                      <WarningIcon style={{ fontSize: 12, verticalAlign: 'middle', marginRight: 4 }} /> {slugError}
                    </div>
                  )}
                  {(!isEditing || isDraft) && (
                    <button
                      onClick={() => handleSaveDraft()}
                      disabled={!meta.formTitle.trim() || !!viewingOld}
                      style={{
                        width: "100%",
                        padding: "10px 0",
                        borderRadius: 9,
                        border: `1px solid ${C.border}`,
                        background: C.white,
                        color: !meta.formTitle.trim() || viewingOld ? C.textMuted : C.textSecond,
                        fontSize: 13,
                        fontWeight: 600,
                        cursor: !meta.formTitle.trim() || viewingOld ? "not-allowed" : "pointer",
                        fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                        marginBottom: 8,
                        opacity: !meta.formTitle.trim() || viewingOld ? 0.5 : 1,
                      }}
                    >
<SaveIcon style={{ fontSize: 14, marginRight: 4 }} /> Save Draft
                    </button>
                  )}
                  <button
                    onClick={() => handlePublish(surveyJson as SurveyJson)}
                    disabled={!meta.formTitle || !meta.formId || !!slugError || !!viewingOld}
                    style={{
                      width: "100%",
                      padding: "11px 0",
                      borderRadius: 9,
                      border: "none",
                      background: !meta.formTitle || !meta.formId || slugError || viewingOld ? C.border : `linear-gradient(135deg,${C.purple},${C.purpleLight})`,
                      color: !meta.formTitle || !meta.formId || slugError || viewingOld ? C.textMuted : C.white,
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: !meta.formTitle || !meta.formId || slugError || viewingOld ? "not-allowed" : "pointer",
                      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                    }}
                  >
                    {viewingOld ? <><WarningIcon style={{ fontSize: 14, verticalAlign: 'middle', marginRight: 4 }} /> Close version preview to publish</> : isDraft ? <><RocketLaunchIcon style={{ fontSize: 14, verticalAlign: 'middle', marginRight: 4 }} /> Publish (make live)</> : <><RocketLaunchIcon style={{ fontSize: 14, verticalAlign: 'middle', marginRight: 4 }} /> Publish to SharePoint</>}
                  </button>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      {deleteConfirm && (
        <div style={{
          position: "fixed",
          inset: 0,
          zIndex: 10001,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "rgba(30,27,75,0.45)",
          animation: "fadeUp .15s ease",
        }}>
          <div style={{
            background: C.white,
            borderRadius: 16,
            padding: "24px 28px",
            maxWidth: 400,
            width: "90%",
            boxShadow: C.shadowMd,
            textAlign: "center",
          }}>
            <div style={{ fontSize: 32, marginBottom: 8, display: 'flex', justifyContent: 'center' }}><WarningIcon style={{ fontSize: 40 }} /></div>
            <div style={{ fontSize: 15, fontWeight: 600, color: C.textPrimary, marginBottom: 6 }}>
              Delete &ldquo;{deleteConfirm.Title}&rdquo;?
            </div>
            <div style={{ fontSize: 12, color: C.textMuted, lineHeight: 1.6, marginBottom: 20 }}>
              This will permanently remove this form and all related data — versions, audit logs, and approver records.
              <br /><br />
              <span style={{ color: C.amber }}>Submission data in the form&rsquo;s list will NOT be deleted.</span>
            </div>
            <div style={{ display: "flex", gap: 9, justifyContent: "center" }}>
              <button
                onClick={() => setDeleteConfirm(null)}
                disabled={deleting}
                style={{
                  height: 36,
                  padding: "0 20px",
                  borderRadius: 9,
                  border: `1px solid ${C.border}`,
                  background: C.white,
                  color: C.textSecond,
                  fontSize: 13,
                  cursor: deleting ? "not-allowed" : "pointer",
                  fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                  opacity: deleting ? 0.6 : 1,
                }}
              >
                Cancel
              </button>
              <button
                onClick={handleDeleteConfirm}
                disabled={deleting}
                style={{
                  height: 36,
                  padding: "0 20px",
                  borderRadius: 9,
                  border: "none",
                  background: `linear-gradient(135deg,${C.red},#B91C1C)`,
                  color: C.white,
                  fontSize: 13,
                  fontWeight: 600,
                  cursor: deleting ? "not-allowed" : "pointer",
                  fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                  opacity: deleting ? 0.6 : 1,
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                {deleting && <Spinner size={14} />}
                {deleting ? "Deleting…" : "Delete Forever"}
              </button>
            </div>
          </div>
        </div>
      )}

      {hardDeleteConfirm && (
        <div style={{
          position: "fixed",
          inset: 0,
          zIndex: 10001,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "rgba(127,29,29,0.45)",
          animation: "fadeUp .15s ease",
        }}>
          <div style={{
            background: C.white,
            borderRadius: 16,
            padding: "24px 28px",
            maxWidth: 420,
            width: "90%",
            boxShadow: C.shadowMd,
            textAlign: "center",
          }}>
            <div style={{ fontSize: 36, marginBottom: 8 }}>💀</div>
            <div style={{ fontSize: 15, fontWeight: 700, color: "#991B1B", marginBottom: 6 }}>
              Permanently delete ALL data for &ldquo;{hardDeleteConfirm.Title}&rdquo;?
            </div>
            <div style={{ fontSize: 12, color: C.textMuted, lineHeight: 1.6, marginBottom: 6 }}>
              This will completely destroy everything related to this form:
            </div>
            <div style={{
              fontSize: 12,
              color: "#991B1B",
              lineHeight: 1.7,
              marginBottom: 16,
              textAlign: "left",
              background: "#FEF2F2",
              borderRadius: 8,
              padding: "10px 14px",
              border: "1px solid #FECACA",
            }}>
              <div>✦ Form configuration (Master Form)</div>
              <div>✦ All version history (Web Form Versions)</div>
              <div>✦ Audit log entries (Form Builder Log)</div>
              <div>✦ Approver records (Approvers)</div>
              <div style={{ fontWeight: 700 }}>✦ ALL submissions in &ldquo;{hardDeleteConfirm.Title} Responses&rdquo; list</div>
            </div>
            <div style={{ fontSize: 11, color: C.red, fontWeight: 600, marginBottom: 18 }}>
              <WarningIcon style={{ fontSize: 14, verticalAlign: 'middle', marginRight: 4 }} /> This action is irreversible. All submission data will be permanently lost.
            </div>
            <div style={{ display: "flex", gap: 9, justifyContent: "center" }}>
              <button
                onClick={() => setHardDeleteConfirm(null)}
                disabled={deleting}
                style={{
                  height: 36,
                  padding: "0 20px",
                  borderRadius: 9,
                  border: `1px solid ${C.border}`,
                  background: C.white,
                  color: C.textSecond,
                  fontSize: 13,
                  cursor: deleting ? "not-allowed" : "pointer",
                  fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                  opacity: deleting ? 0.6 : 1,
                }}
              >
                Cancel
              </button>
              <button
                onClick={handleHardDeleteConfirm}
                disabled={deleting}
                style={{
                  height: 36,
                  padding: "0 24px",
                  borderRadius: 9,
                  border: "none",
                  background: `linear-gradient(135deg,#DC2626,#991B1B)`,
                  color: C.white,
                  fontSize: 13,
                  fontWeight: 700,
                  cursor: deleting ? "not-allowed" : "pointer",
                  fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif",
                  opacity: deleting ? 0.6 : 1,
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                {deleting && <Spinner size={14} />}
                {deleting ? "Deleting…" : "Delete Everything"}
              </button>
            </div>
          </div>
        </div>
      )}

      {provisioning && (
        <ProvisionOverlay
          logs={provLogs}
          success={provOk}
          error={provErr}
          onDone={() => {
            setProvisioning(false);
            if (provOk) navigate(`/admin/builder/${encodeURIComponent(meta.formTitle)}`);
          }}
        />
      )}
    </div>
  );
}