# AGENTS.md — pmw-hrform

## Structure
- **Single app at root**. All commands run from root.
- Entry: `src/main.tsx` → `msalInstance.initialize()` → `BrowserRouter` → `AuthProvider` → `App.tsx`
- Theme: `src/theme/index.ts` (MUI custom, `#0078D4` primary / `#6264A7` secondary). MUI v9 — `Grid` (not Grid2), `slotProps` replaces `PaperProps`.
- **Sub-instructions**: `src/utils/AGENTS.md`, `src/components/builder/AGENTS.md`, `src/pages/AGENTS.md`, `api/AGENTS.md`, `src/components/auth/AGENTS.md`, `src/components/dashboard/AGENTS.md`.

## Commands
```bash
npm run dev        # Vite dev server (port 3000) — frontend only; /api/* will 404
npm run dev:api    # vercel dev — runs BOTH Vite frontend + API routes locally
npm run build      # tsc -b && vite build — fails on new TS errors
npm run lint       # ESLint flat config (many pre-existing warnings)
npx vitest run     # 77 pure-logic unit tests, ~200ms
```
- `npm run build` is the **only** reliable check. `lsp_diagnostics` catches TS errors too.
- `build_errors.txt` / `build_status.txt` are stale — ignore them.

## Stack
- React 19 + TypeScript ~6.0.2 (`verbatimModuleSyntax`, `erasableSyntaxOnly: true` — no runtime `enum`/`namespace`, `noUnusedLocals`, `noUnusedParameters`)
- Vite 8 with `@vitejs/plugin-react` (Oxc-based, React Compiler NOT enabled). `define: { global: 'globalThis' }`.
- `buffer` polyfill: `globalThis.Buffer = Buffer` in `main.tsx` (needed by some SP responses)
- MUI v9, `@azure/msal-react`/`@azure/msal-browser`, `react-router-dom` v7
- **SurveyJS v2.5** (`survey-core`, `survey-react-ui`) — Custom form builder (NOT SurveyJS Creator). CSS imported in `main.tsx`.
- `@react-pdf/renderer` — server-side PDF. `generateAndStorePdf` returns URL only.
- `react-dnd` v16 (HTML5 backend) — drag-drop canvas in form builder.
- API: Vercel serverless functions in `api/` — **not Express**. Graph API client (`graphClient.ts`) for list operations.

## CI
- `.github/workflows/ci.yml`: `npm ci` → `npm run build` → **`npx vitest run`** (build AND tests both gate the pipeline).
- Runs on pull_request + push to main/master branches (ubuntu-latest, Node 20).

## Gotchas

### Formula / Calculated Fields
- Formulas convert to `type: "text", readOnly: true` with a custom `_expression` property in SurveyJSON (SurveyJS native `expression` conflicts with manual evaluation).
- **CSP blocks `new Function()`**: The published form runs under SharePoint's CSP which blocks `unsafe-eval`. Use `safeEvalArithmetic()` (recursive descent parser in `FormBuilderEngine.ts`) instead of `new Function()` or `eval()`. If a formula shows 0 and the console shows an `EvalError`, this is the cause.
- **Expression blank on load**: `buildQuestionTree` reads `_expression` (new format) first, falls back to `expression` (old/native format). If both are missing, expression shows `""` in the builder.

### DynamicMatrix — Separate SharePoint List
- `dynamicmatrix` fields now provision **child lists** named `{FormTitle} Matrix {FieldName}` (NOT just `_Html`/`_Json` columns on the main list).
- Child list columns: `ParentResponseId` (Number — stores parent item ID), `RowIndex` (Number), plus per-matrix-column fields mapped by cellType (Text→2, Dropdown→6, Date→4, Number→9, Checkbox→15, Boolean→8).
- `_Html`/`_Json` columns are still created for backward compatibility but child list is the primary storage.
- Auth submission: `writeMatrixChildItems()` after parent created → `spPatch` parent with `{fieldName}_RowIds` (JSON array of child item IDs).
- Guest submission: `/api/submit-form` accepts `matrixData` parameter, creates child items via Graph API.
- Viewing: `readMatrixChildItems()` reconstructs rows; `rowsToHtml()` generates display table.

### SurveyJS Custom Widget Registration Pattern
```
Serializer.addClass("widgetname", [...props], () => new QuestionModel(""), "questionparent")
  → ElementFactory.Instance.registerElement("widgetname", ...)
  → ReactQuestionFactory.Instance.registerQuestion("widgetname", ...)
```
Used by: `SignaturePad` (`src/utils/SignaturePad.tsx`) and `DynamicMatrix` (`src/utils/DynamicMatrix.tsx`).

### Auth & MSAL
- Auth state machine in `App.tsx`: `checking → loading → ready/wrong_tenant/error` or `guest/choice`.
- Auth decision persisted in `localStorage` (`pmw_hr_auth_decision`). Post-login redirect in `sessionStorage` (`pmw_post_login_redirect`).
- Admin detection via SharePoint group `_HR_ Forms Owners`.
- `handleRedirectPromise()` uses **3s timeout** (`Promise.race`) — fix for hung redirects in private/incognito.
- Clears `sessionStorage` keys `msal.interaction.status` + `msal.login.error` before `loginRedirect()` — DO NOT remove.
- Required env vars validated at startup in `main.tsx`: `VITE_AZURE_CLIENT_ID`, `VITE_AZURE_TENANT_ID`, `VITE_SP_SITE_URL`.
- `src/auth/` contains `msalConfig.ts` (MSAL config + loginRequest) and `AuthProvider.tsx`.

### SharePoint REST
- **OData**: `odata=nometadata` — responses use `data.value` not `data.d.results`.
- **Dual client**: Dashboard uses `createSpClient(msalInstance, accounts)` → `sharepointClient.ts`. Builder uses raw token → `formBuilderSP.ts` (independent digest cache, 30min expiry).
- **Tier query pattern**: Separate lightweight SP queries for optional columns (CurrentLayer, CurrentApprovalLayer, EvaluationData, PdfUrl). Main query only has guaranteed columns. 400 errors on optional queries are caught silently.
- **PATCH**: Use `spPatch()` which sends `X-HTTP-Method: MERGE` with `IF-MATCH: *`.

### DynamicFormPage Submission Flow
- `onCompleting` handler prevents SurveyJS auto-complete (`options.allowComplete = false`), captures data, sets `submitStatus: "loading"`.
- A separate `useEffect` on `submitStatus` triggers `doSubmitForm()`. This split prevents the async submission from blocking the survey's complete cycle.
- `onComplete` is intentionally NOT registered.

## Key Types (`src/types/index.ts`)
- `FormBuilderField` — all field properties including `format?: string`, `expression`, `defaultValue`, `decimalPlaces`, `currencySymbol`, etc.
- `LayerConfig`, `LayerConfigItem` — approval/evaluation layer sequence
- `SurveyJson` — the SurveyJS JSON shape
- `MatrixColumnDef` (in `formBuilderSP.ts`) — dynamicmatrix column definitions

## Anti-Patterns
- **NO `console.log/warn/error`** in production (many exist — remove when touching)
- **NO `any`** — many files have them; fix types when touching
- **NO `@ts-ignore` / `@ts-expect-error`**
- **NO runtime `enum`/`namespace`** — `erasableSyntaxOnly: true`
- **NO `forwardRef`** or manual memoization — React 19 makes these unnecessary
- **NO `useMemo`/`useCallback`** — redundant in React 19; remove when refactoring
- **No path aliases** — all imports relative (`../../utils/...`). No barrel exports except `src/components/builder/index.ts`.
- `FormBuilder.tsx` has `eslint-disable` and `any[]` usage
- `DetailModal.tsx` uses `dangerouslySetInnerHTML` — audit XSS if user input reaches `value`
- `api/_utils/sharepoint.ts` is dead code — safe to delete
- `src/components/session/` is an empty directory
- **Build**: Run `npm run build` after all changes. Do NOT add new TS errors.

## Conventions
- **PowerShell**: use `workdir` parameter with `bash` tool
- **Prefer `import type`** for type-only imports (`verbatimModuleSyntax`)
- **Styling**: Form builder uses inline styles via `C` color object (`src/components/builder/constants.ts`). Published form (`DynamicFormPage.tsx`) uses CSS-in-JS with theme tokens. Dashboard uses MUI components with theme overrides.
- **State**: Local `useState` only — no context stores except `DashboardContext` in `AdminHomePage`.
- **All pages eagerly imported** in `App.tsx` — no `React.lazy()`.

## Testing
- 77 unit tests in `src/utils/__tests__/FormBuilderEngine.test.ts` (pure logic, no network/SP).
- Run: `npx vitest run`. Watch: `npx vitest`.
- No integration/E2E tests exist. No MSW mock handlers.

## Env Vars
| Var | Controls |
|---|---|
| `VITE_SP_SITE_URL` | SharePoint site all SP calls target |
| `VITE_AZURE_CLIENT_ID` | Azure AD app for MSAL auth |
| `VITE_AZURE_TENANT_ID` | Tenant for auth + API |
| `SYSTEM_CLIENT_ID` / `SYSTEM_CLIENT_SECRET` | API server-side auth (Vercel) — NOT `VITE_` prefixed |

## Routing
| Route | Component |
|---|---|
| `/form/:formId` | `DynamicFormPage` (`src/pages/DynamicFormPage.tsx`) |
| `/admin/builder[/:formTitle]` | `AdminFormBuilder` (`src/pages/AdminFormBuilder.tsx`) |
| `/admin/approvals` | `ApprovalDashboard` (`src/components/builder/ApprovalDashboard.tsx`) |
| `/admin/responses/:formTitle` | `ResponseViewer` (`src/components/builder/ResponseViewer.tsx`) |
| `/adminhomepage` | `AdminHomePage` (`src/pages/AdminHomePage.tsx`) |
| `/eval/:token` / `/eval/:formSlug/:responseId/:layerNumber` | `EvaluationPage` (`src/pages/EvaluationPage.tsx`) |

## Builder Architecture (summary)
```
AdminFormBuilder.tsx (page — /admin/builder)
  ├── FormLibrary (sidebar)
  ├── FormBuilder.tsx (canvas — react-dnd drag-drop)
  │     ├── Palette (57 question types)
  │     ├── Canvas (FieldCard reorder + panel nesting)
  │     ├── PropertyPanel (per-field OR Form Settings when deselected)
  │     ├── JsonPreview (collapsed raw JSON)
  │     └── LivePreviewModal (survey-react-ui)
  ├── LayerConfigPanel (approval/evaluation layer sequence editor)
  │     ├── LayerCard[], EvalElementPicker, PublicLinkDisplay
  ├── VersionHistory / AuditLog / ProvisionOverlay
```
