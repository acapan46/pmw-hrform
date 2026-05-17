# COMPREHENSIVE AUDIT: pmw-hrform

**Date**: 2026-05-16
**Audit Scope**: SurveyJS v2.5 integration, SharePoint REST/Graph integration, TypeScript quality, architecture & feature completeness
**Sources**: 4 parallel agent audits + 25+ direct tool searches across 65+ source files (~15,000 lines)

---

## 🔴 CRITICAL (Immediate Fix Required — Runtime Failure / Data Loss / Security)

### [C1] `registerDynamicMatrix()` is NEVER called — DynamicMatrix widget is dead code
- **File**: `src/utils/DynamicMatrix.tsx:405` (definition) → never imported anywhere
- **Issue**: `registerDynamicMatrix()` defined at module level but never imported or called. `registerQuestionData()` (populates data registry) also never called.
- **Fix**: Import and call in `DynamicFormPage.tsx` and `FormBuilder.tsx` before Survey mount:
  ```typescript
  import { registerDynamicMatrix, registerQuestionData } from "../utils/DynamicMatrix";
  registerDynamicMatrix();
  registerQuestionData(enrichedSurveyJson);
  ```
- **Impact**: Forms with `dynamicmatrix` fields render as blank/unrecognized elements

### [C2] SurveyJS `Model` instances NEVER disposed — memory leak on every render
- **6 creation sites, 0 `.dispose()` calls across entire project**:
  1. `src/pages/DynamicFormPage.tsx:370` — `new Model(json)` in `useMemo`, no dispose
  2. `src/components/builder/ApprovalDashboard.tsx:722` — new eval model per item
  3. `src/components/builder/ApprovalDashboard.tsx:866` — another eval model
  4. `src/components/builder/ApprovalDashboard.tsx:1030` — preview model per render
  5. `src/components/builder/ResponseViewer.tsx:249` — Model inline during render
  6. `src/components/builder/FormBuilder.tsx:1546` — LivePreviewModal Model
- **Fix pattern**:
  ```typescript
  const modelRef = useRef<Model | null>(null);
  useEffect(() => {
    return () => { modelRef.current?.dispose(); };
  }, []);
  // Before creating new:
  modelRef.current?.dispose();
  modelRef.current = new Model(json);
  ```
- **Impact**: Each navigation/rerender leaks event handlers, DOM observers, timers

### [C3] ResponseViewer creates Survey Model inline during every render
- **File**: `src/components/builder/ResponseViewer.tsx:246-266`
- **Issue**: `const previewSurvey = surveyJson ? (() => { ... new Model(...) })() : null` — runs on every render, not memoized
- **Fix**: Wrap in `useMemo` with `[surveyJson, selectedSubmission?.RawJSON]` deps
- **Impact**: Performance bomb + memory leak on every state change

### [C4] OData injection — single quotes unescaped in ALL `$filter` values
- **File**: `src/utils/formBuilderSP.ts` — 40+ sites use `encodeURIComponent()` in `$filter` values
- **Issue**: `encodeURIComponent()` does NOT escape `'`. Form title like `O'Brien` breaks SP REST queries.
- **Only correct site**: `getFormSubmissions` at line ~130 uses `.replace(/'/g, "''")`
- **Fix**: Create shared utility:
  ```typescript
  export function sanitizeOData(val: string): string {
    return val.replace(/'/g, "''");
  }
  ```
  Apply to ALL `$filter` string interpolation values.

### [C5] XSS via `dangerouslySetInnerHTML` with unsanitized matrix data
- **Files**:
  - `src/utils/formBuilderSP.ts:1288` — `dynamicMatrixToHtml()` builds `<table>` with `v ?? ''` inline
  - `src/pages/EvaluationPage.tsx:515` — renders `entry.html` via `dangerouslySetInnerHTML`
  - `src/components/dashboard/DetailModal.tsx:508` (uses DOMPurify — safe)
- **Fix**: Entity-encode `<>&"'` in matrix cell values, or use DOMPurify for all HTML rendering

### [C6] `listExists()` has inverted catch logic
- **File**: `src/utils/formBuilderSP.ts:516-518`
- **Code**:
  ```typescript
  catch (e: any) { return !e?.response?.ok; }
  ```
- **Issue**: On network errors, `e.response` is `undefined`, `!undefined` = `true` → returns "list exists" when it doesn't
- **Fix**: `catch { return false; }`

### [C7] `spDelete()` silently swallows ALL errors
- **File**: `src/utils/formBuilderSP.ts:574-585`
- **Issue**: No `response.ok` check. Callers increment counters regardless of success.
- **Fix**: Add `if (!response.ok) throw new Error("spDelete failed: ${response.status}")`

### [C8] 2740-line `FormBuilder.tsx` — single component doing everything
- **File**: `src/components/builder/FormBuilder.tsx`
- **Contains**: Palette, Canvas, FieldCard, PropertyPanel (4 tabs), ConditionEditor, ValidationEditor, LogicRulesEditor, ValueMappingSection, ChoicesEditor, MatrixColumnsEditor, DefaultValueEditor, JsonPreview, LivePreviewModal, PixelPerfectPreview, 10+ inline atom components
- **Fix**: Decompose into:
  ```
  src/components/builder/
    atoms/        (Pill, IconBtn, Toggle, Input, Select, PropLabel, PropRow)
    palette/      (Palette, type icons)
    canvas/       (Canvas, FieldCard, drag-drop reorder)
    properties/   (PropertyPanel, GeneralTab, OptionsTab, LogicTab, ValidationTab)
    editors/      (ChoicesEditor, MatrixColumnsEditor, SpChoicesSourceEditor)
    preview/      (JsonPreview, LivePreviewModal, PixelPerfectPreview)
    settings/     (Form Settings panel)
  ```

### [C9] 706-line `App.tsx` — monolithic orchestrator
- **File**: `src/App.tsx`
- **Contains**: Auth state machine (6 states), dashboard data loading, filter/sort logic, 12 Routes, 27 `useState` calls
- **Fix**: Extract `useDashboardData()` hook + `AuthContext`

### [C10] No integration/E2E tests + CI doesn't run tests
- **Only tests**: 77 tests for `FormBuilderEngine.ts` only
- **CI**: `.github/workflows/ci.yml` — only `npm ci && npm run build`, no test step
- **Fix**: Add `- run: npx vitest run` to CI workflow

---

## 🟠 IMPORTANT (Fix Soon — Data Integrity / UX / Maintainability)

### SurveyJS Issues
| # | Issue | Location |
|---|---|---|
| I1 | Formula evaluation **copy-pasted** in 2 files | `DynamicFormPage.tsx:374-410` + `FormBuilder.tsx:1557-1597` |
| I2 | `onValueChanged` handlers accumulate on undisposed models | All SurveyJS model creation sites |
| I3 | `onCompleting` + `useEffect` — race condition on rapid resubmit | `DynamicFormPage.tsx:441-445` |
| I4 | `autocapitalize` Serializer property registered **twice** | `DynamicFormPage.tsx:33-40` + `FormBuilder.tsx:106-113` |
| I5 | Inconsistent themes: `FlatLightPanelless` vs `LayeredLightPanelless` | Across components |
| I6 | Visual logic rules builder generates wrong SurveyJS operators | `FormBuilder.tsx:500-523` |
| I7 | Formula evaluator can't handle SurveyJS functions (`iif`, `concat`, etc.) | `FormBuilderEngine.ts:1068-1123` |
| I8 | SurveyJS CSS imported 5 times | `main.tsx`, `DynamicFormPage.tsx`, `FormBuilder.tsx`, `ApprovalDashboard.tsx`, `ResponseViewer.tsx` |
| I9 | DynamicMatrix module-level registry unsafe for concurrent surveys | `DynamicMatrix.tsx:38` |

### SharePoint Issues
| # | Issue | Location |
|---|---|---|
| I10 | `resolveUserEmails` — N+1 pattern (1 call per AuthorId) | `sharepointClient.ts:103-129` |
| I11 | `deleteFormVersions`/`deleteFormLogEntries` — N+1 deletion | `formBuilderSP.ts:712-749` |
| I12 | No pagination — `$top=500` hardcoded, no `$skiptoken` | Both SP clients |
| I13 | `queryListByEmail()` fetches ALL items, filters client-side | `sharepointClient.ts:259-306` |
| I14 | `createSpList()` has 1.5s hardcoded delay | `formBuilderSP.ts:498` |
| I15 | `catch (e: any)` in ~20 locations | `formBuilderSP.ts` (multiple) |
| I16 | 30-min digest cache never invalidated on 401 | Both SP clients |
| I17 | Missing `$select` on many SP queries | `getFormSubmissions`, `getLayerResponseData` |
| I18 | No request timeout on any `fetch()` call | Both clients + Graph client |
| I19 | `saveFormVersion` double-stringifies with pretty-print | `formBuilderSP.ts:94` |
| I20 | `diffSurveyJson` only checks first page's elements | `formBuilderSP.ts:853-854` |

### Architecture & Code Quality
| # | Issue | Location |
|---|---|---|
| I21 | No route-level code splitting — all pages eagerly imported | `App.tsx:29-31` |
| I22 | 27 props tunneled through `DashboardProvider` | `App.tsx:561-585`, `DashboardContext.tsx` |
| I23 | `handlePublish` — 210-line callback, 23 deps | `AdminFormBuilder.tsx:600-810` |
| I24 | `useMemo`/`useCallback` proliferation (React 19 anti-pattern) | `FormBuilder.tsx` (18+), `AdminFormBuilder.tsx` |
| I25 | Form builder uses `C` color object, not MUI theme | `constants.ts` vs `theme/index.ts` |
| I26 | Only one ErrorBoundary wrapping ALL routes | `App.tsx:596-700` |
| I27 | `api/send-email.ts` has no auth/access control | Public endpoint, no rate limiting |

---

## 🟡 MINOR (Clean Up When Touching)

| # | Issue | Location |
|---|---|---|
| M1 | `import React` unnecessary in React 19 | `FormBuilder.tsx:5` |
| M2 | `Record<string, unknown>` used 250× across 27 files | Project-wide |
| M3 | 15 empty `catch {}` blocks | 5 files |
| M4 | 46 `console.log/warn/error` calls in production | 15 files |
| M5 | `api/_utils/sharepoint.ts` — confirmed dead code (92 lines) | DELETE IT |
| M6 | Form builder has `OK` text + `LOCK` text as icons | `DynamicFormPage.tsx:129,139` |
| M7 | No `AbortController` timeout on any `fetch()` | All SP/Graph clients |
| M8 | WrongTenantScreen exposes `tenantId` to user | `WrongTenantScreen.tsx` |
| M9 | `npm run build` uses `tsc -b` (slow full rebuild) | `package.json` |
| M10 | Duplicate digest cache in both SP clients | `sharepointClient.ts` + `formBuilderSP.ts` |
| M11 | `SP_SITE_URL` recomputed in function bodies | `sharepointClient.ts` |
| M12 | Missing `$select` on `getLayerResponseData` — fetches ALL columns | `formBuilderSP.ts:1519` |
| M13 | `FormBuilder.tsx` has `eslint-disable` and `any[]` | Various lines |
| M14 | `AdminGuard` has 4-second redirect delay | `AdminGuard.tsx:45` |

---

## ⚫ MISSING (Features Expected in Production Form Builder)

| # | Feature | Notes |
|---|---|---|
| F1 | **Analytics / Reporting Dashboard** | No submission counts, trends, response distributions |
| F2 | **Webhooks** | `WebhookConfig` type exists at `types/index.ts:724` — ZERO implementation |
| F3 | **Email Notification Config UI** | `EmailTemplate` type exists — hardcoded in `triggerApprovalNotification()` |
| F4 | **Form Templates** | No template library |
| F5 | **Pre-fill from URL** | No `?field=value` query param support |
| F6 | **Autosave / Draft Recovery** | No localStorage backup during form fill |
| F7 | **CAPTCHA / Bot Protection** | Public forms have no protection |
| F8 | **Form Scheduling / Expiration** | No start/end date |
| F9 | **File Management UI** | No browse/manage/delete for uploaded files |
| F10 | **GDPR / Data Retention** | No auto-purge, data export, or delete policies |
| F11 | **i18n / Multi-language** | `translations` field in types — no implementation |
| F12 | **Accessibility (WCAG)** | No ARIA audit, keyboard nav, screen reader |
| F13 | **Offline Support** | No Service Worker |
| F14 | **Form Copy/Duplicate** | No "Duplicate" in Form Library |
| F15 | **Transaction rollback on provisioning failure** | Partial column creation on failure — no cleanup |
| F16 | **Health check endpoint** | No `/api/health` |

---

## ✅ WHAT'S DONE WELL

| Strength | Details |
|---|---|
| **57 field types** with per-type editors | Matrix, signature, file upload, ranking, formula, hierarchy, NPS, etc. |
| **Multi-layer approval chains** | Conditional routing, manual branches, field-reference assignees, public tokens |
| **Custom SurveyJS builder** without Creator | Correct SurveyJSON generation via bespoke react-dnd canvas |
| **SharePoint-native storage** | Master Form → Versions → Log → Approvers → child lists for matrix data |
| **PDF generation** | Server-side `@react-pdf/renderer` with embedded layer results/signatures |
| **Dual auth mode** | M365 authenticated + public guest with separate Graph API endpoints |
| **CSP-compliant formulas** | `safeEvalArithmetic()` avoids `new Function()` block |
| **Type system** | 872-line types with discriminated unions, branded types |
| **Status management** | `SP_LAYER_STATUS`/`SP_FORM_STATUS` with legacy migration |
| **Versioning + Audit** | Web Form Versions + Form Builder Log with diff view |

---

## PRIORITY ACTION PLAN

### Week 1: Critical Fixes
1. **C1** — Call `registerDynamicMatrix()` (5 min)
2. **C2-C3** — Add `.dispose()` at 6 Model sites + memoize ResponseViewer (1 hr)
3. **C4** — Create `sanitizeODataValue()` + fix all 40+ `$filter` sites (30 min)
4. **C5** — Sanitize `dynamicMatrixToHtml()` output (15 min)
5. **C6** — Fix `listExists()` catch (5 min)
6. **C7** — Add `response.ok` check to `spDelete()` (5 min)
7. **C10** — Add `npx vitest run` to CI (2 min)
8. **M5** — Delete `api/_utils/sharepoint.ts` (1 min)

### Week 2: Important Fixes
9. **I1** — Extract shared formula evaluator (30 min)
10. **I6** — Fix RulesSection operator mapping (20 min)
11. **I21** — Add `React.lazy()` for admin pages (15 min)
12. **C8** — Begin decomposing FormBuilder.tsx (2-3 days)
13. **I10** — Fix `resolveUserEmails` N+1 with `$expand=Author` (30 min)
14. **I16** — Add 401 digest cache invalidation (20 min)

### Week 3-4: Architecture
15. **C9** — Extract `AuthContext` + `useDashboardData` hook (1 day)
16. **I4** — Move `autocapitalize` to shared module (10 min)
17. **I23** — Decompose `handlePublish` (1 day)
18. **I24** — Remove unnecessary `useMemo`/`useCallback` (1 day)
19. **I26** — Wrap each Route in individual ErrorBoundary (30 min)

### Future: Feature Gaps
20. **F1** — Analytics dashboard
21. **F2** — Webhook executor
22. **F7** — CAPTCHA for public forms
23. **F15** — SurveyJS CSP expression engine migration (`settings.expressionCspEnabler = true`)

---

## KEY METRICS

| Metric | Value |
|--------|-------|
| Source files (tsx/ts) | 56 |
| Largest file | `FormBuilder.tsx` — 2740 lines |
| Files >500 lines | 12 |
| SurveyJS Model instances | 6 — **0 disposed** |
| Empty `catch {}` blocks | 15 |
| `catch (e: any)` | ~20 |
| `console.*` in production | 46 calls across 15 files |
| `Record<string, unknown>` | 250 across 27 files |
| Test coverage | 77 tests — **1 module only** |
| API routes | 4 (form-config, submit-form, evaluate, send-email) |
| Dead code | `api/_utils/sharepoint.ts` (92 lines) |
