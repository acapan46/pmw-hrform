import { useState, useEffect, useCallback } from "react";
import { Routes, Route, useNavigate, useLocation, Navigate } from "react-router-dom";
import {
  useMsal,
  useIsAuthenticated,
} from "@azure/msal-react";
import { ThemeProvider, CssBaseline, Box } from "@mui/material";
import theme from "./theme";
import { loginRequest } from "./auth/msalConfig";
import { createSpClient } from "./utils/sharepointClient";
import { SP_STATIC, loadConfig, filterVisibleLists, getMissingConfigs, generateMeta } from "./utils/spConfig";
import { getStoredAuthDecision, setStoredAuthDecision, clearStoredAuthDecision } from "./utils/authDecision";
import type { PageState, Submission, ApprovalLayer, DiscoveredList, ListMetaEntry, LoadedConfig, LayerConfig, ApprovalLayerConfig, ApprovalLayerResult, EvaluationLayerResult, EvaluationDataEntry } from "./types";
import { normalizeLayerStatus } from "./utils/statusConstants";

// Auth screens
import ChoiceScreen from "./components/auth/ChoiceScreen";
import GuestLanding from "./components/auth/GuestLanding";
import WrongTenantScreen from "./components/auth/WrongTenantScreen";
import LoadingScreen from "./components/auth/LoadingScreen";
import ErrorScreen from "./components/auth/ErrorScreen";
import AdminGuard from "./components/auth/AdminGuard";
import ErrorBoundary from "./components/ErrorBoundary";

// Form builder pages
import DynamicFormPage from "./pages/DynamicFormPage";
import ApprovalDashboard from "./components/builder/ApprovalDashboard";
import ResponseViewer from "./components/builder/ResponseViewer";
import AdminFormBuilder from "./pages/AdminFormBuilder";
import AdminHomePage from "./pages/AdminHomePage";
import EvaluationPage from "./pages/EvaluationPage";
import { DashboardProvider } from "./contexts/DashboardContext";



const ALLOWED_TENANT_ID = import.meta.env.VITE_AZURE_TENANT_ID || "";

function normalizeStatus(status: string | null): string {
  if (!status) return "pending";
  const normalized = status.toLowerCase().replace(/[\s_-]/g, "");
  if (normalized === "fullyapproved") return "fullyapproved";
  if (normalized === "approved") return "approved";
  if (normalized.includes("reject")) return "rejected";
  if (normalized.includes("progress") || normalized.includes("review")) return "inprogress";
  return "pending";
}

function mapSubmission(
  raw: Record<string, unknown>,
  listTitle: string,
  listMetaMap: Record<string, ListMetaEntry>,
  layerConfigs?: Record<string, LayerConfig | null>
): Submission {
  const id = String(raw.Id || "");
  const title = String(raw.Title || "Untitled");
  const formId = String(raw.FormId || "");
  const formVersion = String(raw.FormVersion || "1");
  const formStatus = raw.FormStatus ? String(raw.FormStatus) : null;
  const submittedByEmail = String(
    raw._authorEmail ||
      (raw.Author as Record<string, unknown> | undefined)?.Email ||
      raw.submittedByEmail ||
      ""
  );
  const submittedAt = raw.SubmittedAt ? String(raw.SubmittedAt) : null;
  const currentLayer = raw.CurrentLayer ? Number(raw.CurrentLayer) : 0;

  const cfg = layerConfigs?.[listTitle] ?? null;
  const layersConfig = cfg?.layers ?? [];

  let totalLayers = layersConfig.length;
  if (!totalLayers) {
    totalLayers = 1;
    if (raw.L2_Email) totalLayers = 2;
    if (raw.L3_Email) totalLayers = 3;
  }

  const layers: (ApprovalLayer | null)[] = [];
  const enhancedLayers: (ApprovalLayerResult | EvaluationLayerResult | null)[] = [];

  if (layersConfig.length > 0) {
    for (let i = 0; i < layersConfig.length; i++) {
      const lc = layersConfig[i];
      const n = lc.layerNumber;
      const statusVal = raw[`L${n}_Status`] ? String(raw[`L${n}_Status`]) : null;
      const emailVal = raw[`L${n}_Email`] ? String(raw[`L${n}_Email`]) : null;
      const signedAtVal = raw[`L${n}_SignedAt`] ? String(raw[`L${n}_SignedAt`]) : null;
      const rejectionVal = raw[`L${n}_Rejection`] ? String(raw[`L${n}_Rejection`]) : null;
      const signatureVal = raw[`L${n}_Signature`] ? String(raw[`L${n}_Signature`]) : null;
      const canonicalStatus = normalizeLayerStatus(statusVal);

      layers.push({
        status: statusVal || "pending",
        outcome: canonicalStatus === "approved" ? "approved" : canonicalStatus === "rejected" ? "rejected" : undefined,
        email: emailVal,
        signedAt: signedAtVal,
        rejectionReason: rejectionVal,
        signature: signatureVal,
      });

      if (lc.type === "evaluation") {
        let evalData: EvaluationDataEntry | null = null;
        const rawEvalData = raw.EvaluationData as string | undefined;
        if (rawEvalData) {
          try {
            const allEvalData = JSON.parse(rawEvalData) as Record<number, EvaluationDataEntry>;
            evalData = allEvalData[n] ?? null;
          } catch {}
        }
        enhancedLayers.push({
          layerNumber: n,
          type: "evaluation",
          status: canonicalStatus,
          email: emailVal,
          confirmedAt: evalData?.confirmedAt ?? null,
          fields: evalData?.fields ?? {},
          notes: evalData?.notes,
        });
      } else {
        enhancedLayers.push({
          layerNumber: n,
          type: "approval",
          status: canonicalStatus,
          outcome: canonicalStatus === "approved" ? "approved" : canonicalStatus === "rejected" ? "rejected" : undefined,
          email: emailVal,
          signedAt: signedAtVal,
          rejectionReason: rejectionVal,
          signature: signatureVal,
          confirmedVia: (lc as ApprovalLayerConfig).confirmationType ?? "signature",
        });
      }
    }
  } else {
    // Legacy path — old L1-L3 loop
    for (let i = 1; i <= 3; i++) {
      const statusVal = raw[`L${i}_Status`] ? String(raw[`L${i}_Status`]) : null;
      const emailVal = raw[`L${i}_Email`] ? String(raw[`L${i}_Email`]) : null;
      const signedAtVal = raw[`L${i}_SignedAt`] ? String(raw[`L${i}_SignedAt`]) : null;
      const rejectionVal = raw[`L${i}_Rejection`] ? String(raw[`L${i}_Rejection`]) : null;
      const signatureVal = raw[`L${i}_Signature`] ? String(raw[`L${i}_Signature`]) : null;
      if (statusVal || emailVal) {
        layers.push({
          status: statusVal || "pending",
          outcome: statusVal === "approved" ? "approved" : statusVal === "rejected" ? "rejected" : undefined,
          email: emailVal,
          signedAt: signedAtVal,
          rejectionReason: rejectionVal,
          signature: signatureVal,
        });
      }
    }
  }

  // Filter internal fields
  const submissionData: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(raw)) {
    if (
      key.startsWith("odata.") ||
      /^L[1-9]_/.test(key) ||
      key === "FormStatus" ||
      key === "CurrentLayer" ||
      key === "EvaluationData" ||
      key === "FormId" ||
      key === "FormVersion" ||
      key === "Title" ||
      key === "Id" ||
      key === "_authorEmail" ||
      key === "SubmittedAt" ||
      key === "AuthorId"
    ) {
      continue;
    }
    submissionData[key] = value;
  }

  return {
    id,
    submissionId: id,
    listTitle,
    formId,
    formVersion,
    title,
    submittedByEmail,
    submittedAt,
    formStatus,
    totalLayers,
    layers: layers.filter(Boolean) as ApprovalLayer[],
    meta: listMetaMap[listTitle] ?? generateMeta(listTitle),
    submissionData,
    currentLayer,
    enhancedLayers: enhancedLayers.length > 0 ? enhancedLayers : undefined,
    layerConfig: cfg,
  };
}

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [pageState, setPageState] = useState<PageState>("checking");
  const [errorMsg, setErrorMsg] = useState("");
  const userEmail = accounts[0]?.username || "";
  const [isAdmin, setIsAdmin] = useState(false);

  // Dashboard data
  const [submissions, setSubmissions] = useState<Submission[]>([]);
  const [visibleLists, setVisibleLists] = useState<DiscoveredList[]>([]);
  const [loadedConfig, setLoadedConfig] = useState<LoadedConfig | null>(null);
  const [missingConfigs, setMissingConfigs] = useState<string[]>([]);
  const [detailItem, setDetailItem] = useState<Submission | null>(null);
  const [loading, setLoading] = useState(false);
  const [loadProgress, setLoadProgress] = useState(0);
  const [loadStatus, setLoadStatus] = useState("Initializing...");

  // Filters
  const [search, setSearch] = useState("");
  const [listFilter, setListFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [sortBy, setSortBy] = useState("newest");
  const [submitterFilter, setSubmitterFilter] = useState("");

  const navigate = useNavigate();

  // Auth state machine
  useEffect(() => {
    if (inProgress !== "none") return;

    if (isAuthenticated) {
      setPageState("loading");
      return;
    }

    // Check for redirect result first before deciding page state
const decision = getStoredAuthDecision();
if (decision === "guest") {
  setPageState("guest");
} else {
  setPageState("choice");
}
  }, [isAuthenticated, inProgress, accounts, instance]);

  
  useEffect(() => {
    if (pageState !== "loading" || !isAuthenticated) return;

    const email = accounts[0]?.username || "";
    const tenantId = accounts[0]?.tenantId || "";

    if (ALLOWED_TENANT_ID && tenantId !== ALLOWED_TENANT_ID) {
      setPageState("wrong_tenant");
      return;
    }

    setLoading(true);
    setLoadProgress(0);
    setLoadStatus("Initializing...");
    const spClient = createSpClient(instance, accounts);

    async function fetchData() {
      try {
        // Step 1: Check admin status
        setLoadStatus("Checking permissions...");
        setLoadProgress(10);
        const adminResult = await spClient.isGroupMember(SP_STATIC.adminGroup);
        setIsAdmin(adminResult);
        setLoadProgress(20);

        // Step 2: Discover lists
        setLoadStatus("Discovering lists...");
        const allLists = await spClient.discoverLists();
        setLoadProgress(35);

        // Step 3: Load config from Master Form
        setLoadStatus("Loading configuration...");
        const config = await loadConfig(spClient);
        setLoadedConfig(config);
        setLoadProgress(50);

        // Build map of list → set of emails that should see submissions (including layer assignees)
        const assigneeVisibilityMap: Record<string, Set<string>> = {};
        for (const [title, cfg] of Object.entries(config.layerConfigs || {})) {
          if (!cfg?.layers) continue;
          for (const layer of cfg.layers) {
            if (layer.assignee.type === "user" && layer.assignee.value) {
              if (!assigneeVisibilityMap[title]) assigneeVisibilityMap[title] = new Set();
              assigneeVisibilityMap[title].add(layer.assignee.value.toLowerCase());
            }
          }
        }

        // Step 4: Filter visible lists
        const visible = filterVisibleLists(allLists, adminResult, config.allowedTitles);
        setVisibleLists(visible);

        const listMetaMap: Record<string, ListMetaEntry> = { ...config.listMetaMap };
        for (const list of visible) {
          if (!listMetaMap[list.title]) {
            listMetaMap[list.title] = generateMeta(list.title);
          }
        }

        // Step 5: Fetch submissions
        const totalLists = visible.length;
        setLoadStatus(
          totalLists > 0
            ? `Fetching submissions from ${totalLists} list${totalLists !== 1 ? "s" : ""}...`
            : "No lists to fetch from."
        );

        const allSubmissions: Submission[] = [];
        for (let i = 0; i < visible.length; i++) {
          const list = visible[i];
          setLoadStatus(`Fetching submissions from "${list.title}" (${i + 1}/${totalLists})...`);
          setLoadProgress(50 + Math.round(((i + 1) / totalLists) * 45));

          try {
            const items = await spClient.queryList(list.title, {
              select: "*",
              orderby: "Created desc",
              top: adminResult ? 5000 : 1000,
            });
            for (const item of items) {
              allSubmissions.push(mapSubmission(item, list.title, listMetaMap, config.layerConfigs));
            }
          } catch {
            // Skip lists that fail to query
          }
        }

        // Step 6: Finalize
        setLoadStatus("Finalizing...");
        setLoadProgress(98);

        const visibleTitles = new Set(visible.map((l) => l.title));
        let finalSubmissions = allSubmissions.filter((s) => visibleTitles.has(s.listTitle));
        if (!adminResult && email) {
          const lowerEmail = email.toLowerCase();
          finalSubmissions = finalSubmissions.filter((s) => {
            // User's own submissions
            if (s.submittedByEmail.toLowerCase() === lowerEmail) return true;
            // Submissions where user is a layer assignee
            const assignees = assigneeVisibilityMap[s.listTitle];
            if (assignees?.has(lowerEmail)) return true;
            return false;
          });
        }

        setSubmissions(finalSubmissions);
        setMissingConfigs(getMissingConfigs(visible, config.layerConfig));
        setLoadProgress(100);
        setLoadStatus("Ready.");
        setPageState("ready");
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : "Unknown error occurred";
        setErrorMsg(message);
        setPageState("error");
      } finally {
        setLoading(false);
      }
    }

    fetchData();
  }, [pageState, isAuthenticated]); // eslint-disable-line react-hooks/exhaustive-deps

  // Navigate to preserved route after successful login
  useEffect(() => {
    if (pageState === "ready" && isAuthenticated) {
      try {
        const redirectPath = sessionStorage.getItem("pmw_post_login_redirect");
        if (redirectPath) {
          sessionStorage.removeItem("pmw_post_login_redirect");
          // Root or legacy adminhomepage → role-specific dashboard
          if (redirectPath === "/" || redirectPath === "/adminhomepage") {
            navigate(isAdmin ? "/admin/dashboard" : "/user/dashboard", { replace: true });
          } else {
            navigate(redirectPath);
          }
        } else {
          // No stored redirect — go to role-appropriate dashboard
          navigate(isAdmin ? "/admin/dashboard" : "/user/dashboard", { replace: true });
        }
      } catch {
        // Ignore storage errors
      }
    }
  }, [pageState, isAuthenticated, navigate, isAdmin]);

  const handleLogin = () => {
    console.log("handleLogin called, inProgress:", inProgress, "accounts:", instance.getAllAccounts());
    
    // Check if login already in progress
    if (inProgress !== "none") {
      console.warn("Login already in progress, waiting...");
      return;
    }
    
    setStoredAuthDecision("msal");

    // Preserve current route for post-login redirect
    try {
      sessionStorage.setItem("pmw_post_login_redirect", window.location.pathname + window.location.search);
    } catch {
      // May fail if storage is inaccessible
    }

    // Clear MSAL sessionStorage cache to remove stale interaction state
    // This is the key fix for interaction_in_progress error
    try {
      sessionStorage.removeItem("msal.interaction.status");
      sessionStorage.removeItem("msal.login.error");
    } catch {
      // May fail if storage is inaccessible
    }
    
    console.log("Calling instance.loginRedirect...");
    instance.loginRedirect(loginRequest);
  };

  const handleGuest = () => {
    setStoredAuthDecision("guest");
    setPageState("guest");
  };

  const handleSwitchAccount = useCallback(() => {
    instance.logoutPopup().catch(() => {
      instance.logoutRedirect();
    });
    clearStoredAuthDecision();
    setTimeout(() => {
      instance.loginRedirect(loginRequest);
    }, 100);
  }, [instance]);

  const handleSignOut = useCallback(() => {
    instance.logoutRedirect();
    clearStoredAuthDecision();
  }, [instance]);

  const handleForgetChoice = () => {
    clearStoredAuthDecision();
    setPageState("choice");
  };

  // Filter + sort logic
  const filteredSubmissions = submissions.filter((item) => {
    if (search) {
      const searchLower = search.toLowerCase();
      if (
        !item.title.toLowerCase().includes(searchLower) &&
        !item.formId.toLowerCase().includes(searchLower) &&
        !item.submissionId.toLowerCase().includes(searchLower)
      ) {
        return false;
      }
    }
    if (listFilter && item.listTitle !== listFilter) return false;
    if (statusFilter !== "all" && normalizeStatus(item.formStatus) !== statusFilter.toLowerCase()) return false;
    if (submitterFilter && !item.submittedByEmail.toLowerCase().includes(submitterFilter.toLowerCase()))
      return false;
    return true;
  });

  const sortedSubmissions = [...filteredSubmissions].sort((a, b) => {
    switch (sortBy) {
      case "oldest":
        return (a.submittedAt || "").localeCompare(b.submittedAt || "");
      case "status":
        return normalizeStatus(a.formStatus).localeCompare(normalizeStatus(b.formStatus));
      case "list":
        return a.listTitle.localeCompare(b.listTitle);
      default: // newest
        return (b.submittedAt || "").localeCompare(a.submittedAt || "");
    }
  });

  const listMetaMap = { ...loadedConfig?.listMetaMap };
  for (const list of visibleLists) {
    if (!listMetaMap[list.title]) {
      listMetaMap[list.title] = generateMeta(list.title);
    }
  }

  const hasFilters = !!(search || listFilter || statusFilter !== "all" || submitterFilter);

  const location = useLocation();
  const currentPath = location.pathname;
  const isFormRoute = currentPath.startsWith("/form/");
  const isAdminRoute = currentPath.startsWith("/admin/");

  // ---- Render ----

  if (pageState === "checking" || (pageState === "loading" && loading && !isFormRoute && !isAdminRoute)) {
    return (
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <LoadingScreen userEmail={userEmail || undefined} progress={loadProgress} status={loadStatus} />
      </ThemeProvider>
    );
  }

  const showAuthGate = !isAuthenticated && !isFormRoute;

  if (showAuthGate && pageState === "choice") {
    return (
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <ChoiceScreen onLogin={handleLogin} onGuest={handleGuest} />
      </ThemeProvider>
    );
  }

  if (showAuthGate && pageState === "guest") {
    return (
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <GuestLanding onLogin={handleLogin} onForgetChoice={handleForgetChoice} />
      </ThemeProvider>
    );
  }

  if (showAuthGate && pageState === "wrong_tenant") {
    return (
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <WrongTenantScreen userEmail={userEmail} onLogout={handleSignOut} onSwitch={handleSwitchAccount} />
      </ThemeProvider>
    );
  }

  if (showAuthGate && pageState === "error") {
    return (
      <ThemeProvider theme={theme}>
        <CssBaseline />
        <ErrorScreen errorMsg={errorMsg} onRetry={() => setPageState("loading")} onSignOut={handleSignOut} />
      </ThemeProvider>
    );
  }

  // ---- Dashboard (ready state) ----
  const adminDashboardInner = (
    <ErrorBoundary>
      <DashboardProvider
        userEmail={userEmail}
        isAdmin={isAdmin}
        submissions={submissions}
        visibleLists={visibleLists}
        listMetaMap={listMetaMap}
        missingConfigs={missingConfigs}
        hasFilters={hasFilters}
        detailItem={detailItem}
        setDetailItem={setDetailItem}
        search={search}
        setSearch={setSearch}
        listFilter={listFilter}
        setListFilter={setListFilter}
        statusFilter={statusFilter}
        setStatusFilter={setStatusFilter}
        sortBy={sortBy}
        setSortBy={setSortBy}
        submitterFilter={submitterFilter}
        setSubmitterFilter={setSubmitterFilter}
        sortedSubmissions={sortedSubmissions}
        onSignOut={handleSignOut}
        onSwitchAccount={handleSwitchAccount}
        onOpenBuilder={() => navigate("/admin/builder")}
        onEditForm={(listTitle: string) => navigate(`/admin/builder/${encodeURIComponent(listTitle)}`)}
      >
        <AdminHomePage />
      </DashboardProvider>
    </ErrorBoundary>
  );

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <ErrorBoundary>
        <Routes>
          <Route
            path="/form/:formId"
            element={
              <ErrorBoundary>
                <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
                  <DynamicFormPage />
                </Box>
              </ErrorBoundary>
            }
          />
          <Route
            path="/admin/approvals"
            element={
              <AdminGuard isAdmin={isAdmin}>
                <ErrorBoundary>
                  <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
                    <ApprovalDashboard />
                  </Box>
                </ErrorBoundary>
              </AdminGuard>
            }
          />
          <Route
            path="/admin/responses/:formTitle"
            element={
              <AdminGuard isAdmin={isAdmin}>
                <ErrorBoundary>
                  <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
                    <ResponseViewer />
                  </Box>
                </ErrorBoundary>
              </AdminGuard>
            }
          />
          <Route
            path="/admin/builder"
            element={
              <AdminGuard isAdmin={isAdmin}>
                <ErrorBoundary>
                  <Box sx={{ minHeight: "100vh" }}>
                    <AdminFormBuilder />
                  </Box>
                </ErrorBoundary>
              </AdminGuard>
            }
          />
          <Route
            path="/admin/builder/:formTitle"
            element={
              <AdminGuard isAdmin={isAdmin}>
                <ErrorBoundary>
                  <Box sx={{ minHeight: "100vh" }}>
                    <AdminFormBuilder />
                  </Box>
                </ErrorBoundary>
              </AdminGuard>
            }
          />
          <Route
            path="/admin/dashboard"
            element={
              <AdminGuard isAdmin={isAdmin}>
                {adminDashboardInner}
              </AdminGuard>
            }
          />
          <Route
            path="/user/dashboard"
            element={
              <ErrorBoundary>
                {adminDashboardInner}
              </ErrorBoundary>
            }
          />
          <Route
            path="/eval/:token"
            element={
              <ErrorBoundary>
                <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
                  <EvaluationPage />
                </Box>
              </ErrorBoundary>
            }
          />
          <Route
            path="/eval/:formSlug/:responseId/:layerNumber"
            element={
              <ErrorBoundary>
                <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
                  <EvaluationPage />
                </Box>
              </ErrorBoundary>
            }
          />
          <Route
            path="*"
            element={
              pageState === "ready" ? (
                <Navigate to={isAdmin ? "/admin/dashboard" : "/user/dashboard"} replace />
              ) : (
                adminDashboardInner
              )
            }
          />
        </Routes>

      </ErrorBoundary>
    </ThemeProvider>
  );
}
