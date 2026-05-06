# AGENTS.md — pmw-hrform

## Structure
- **Single app at root** (NOT `pmw-hrform-app/`). All commands run from root.
- Entry: `src/main.tsx` → `BrowserRouter` → `AuthProvider` → `App.tsx`
- Theme: `src/theme/index.ts` (MUI custom, #0078D4 primary / #6264A7 secondary)
- Assets: `public/` (favicon, icons.svg), `src/assets/` (hero.png)
- **Sub-instructions**: `src/utils/AGENTS.md`, `src/components/builder/AGENTS.md`, `src/components/dashboard/AGENTS.md`, `src/pages/AGENTS.md`, `api/AGENTS.md` have additional hard-earned context.

## Commands (run from root)
```bash
npm run dev        # Vite dev server (port 3000)
npm run dev:api    # Vercel dev server (runs api/ routes locally)
npm run build      # tsc -b && vite build
npm run lint       # ESLint flat config
npm run preview    # Preview production build
```

## Stack
- React 19 + TypeScript ~6.0.2 (ES2022 target, bundler moduleResolution, `verbatimModuleSyntax`)
- Vite 8 with `@vitejs/plugin-react` (Oxc-based, React Compiler NOT enabled)
- MUI v9 (`@mui/material`, `@mui/icons-material`) — `Grid` (not Grid2) in v9.0.0
- `@azure/msal-react` + `@azure/msal-browser` — Azure AD authentication
- **react-router-dom v7** — Routing now active: `BrowserRouter` in `main.tsx`, `<Routes>` in `App.tsx`
- **SurveyJS v2.5** — `survey-core`, `survey-react-ui`; **Custom form builder** (NOT SurveyJS Creator)
- ESLint: flat config with `tseslint.configs.recommended` (not type-checked), `react-hooks`, `react-refresh`

## Auth State Machine (in `App.tsx`)
```
checking → (isAuthenticated?) → loading → (tenant check) → ready / wrong_tenant / error
         → (not auth, stored decision?) → guest / choice
         → (not auth, no decision) → choice → MSAL login or guest
```
States: `checking` | `choice` | `guest` | `loading` | `ready` | `wrong_tenant` | `error`
- Auth decision persisted in `localStorage` (`pmw_hr_auth_decision`)
- Post-login redirect stored in `sessionStorage` (`pmw_post_login_redirect`), read after MSAL callback to restore the intended route
- Admin detection via SharePoint group `_HR_ Forms Owners`
- Tenant validation via `VITE_AZURE_TENANT_ID` env var

## Routing (react-router-dom v7)
- `"/form/:formId"` → `DynamicFormPage` (public/private form rendering)
- `"/admin/builder"` → `AdminFormBuilder` (full builder with sidebar, library, publish)
- `"/admin/builder/:formTitle"` → `AdminFormBuilder` (loads specific form for editing)
- `"/admin/approvals"` → `ApprovalDashboard` (approval workflow)
- `"/admin/responses/:formTitle"` → `ResponseViewer` (submission responses)
- `"/adminhomepage"` → `AdminHomePage` (explicit dashboard route)
- `"*"` → `AdminHomePage` (catch-all: Header, StatsRow, ListSummaryCards, Toolbar, SubmissionRow)
- Auth screens (`ChoiceScreen`, `GuestLanding`, `WrongTenantScreen`, etc.) are rendered directly in `App.tsx` before `<Routes>` — NOT via `HomePage.tsx`
- `HomePage.tsx` is **dead code** — not imported by `App.tsx`; the real choice screen is `ChoiceScreen.tsx`
- Header "Form Builder" button navigates to `/admin/builder` (not a modal)
- The old `<Dialog>` FormBuilder in `AdminHomePage` is dead code — `builderOpen` is never set to `true`

## API Routes (Vercel-style serverless)
- `api/form-config.ts` — Public form config fetcher (used by unauthenticated users)
- `api/submit-form.ts` — Public form submission handler (guest submissions)
- `api/_utils/graphClient.ts` — **Active** server-side client; uses Microsoft Graph API (`graph.microsoft.com/v1.0`) with client-credentials token (`https://graph.microsoft.com/.default`). Exports: `getGraphToken`, `graphGet`, `graphPost`, `queryListItems`, `createListItem`
- `api/_utils/sharepoint.ts` — **Dead code** — exists but is not imported by any API route
- **Import paths**: API routes import from `./_utils/graphClient.ts`
- `vercel.json` configures SPA rewrite rules and CORS headers for `/api/*`

## SharePoint Integration
- **odata format**: `odata=nometadata` — responses use `data.value` not `data.d.results`
- **User resolution**: `resolveUserEmails()` in `sharepointClient.ts` maps `AuthorId` → email via `_api/web/getUserById()`
- **`sharepointClient.ts`** exports `createSpClient(instance, accounts)` — passes MSAL instance
- **`formBuilderSP.ts`** — STANDALONE file; uses raw `token: string` param, NOT `createSpClient`
- Config loaded from `Master Form` list (FormId, TotalLayers mapping)
- Lists discovered via `/_api/web/lists` (system lists filtered by BaseTemplate + properties)
- Submissions queried per visible list with `$orderby: Created desc`
- Digest cached with 30min expiry, `X-HTTP-Method` headers for writes

## SP Column Type Mapping
| SurveyJS Type | SP FieldTypeKind | SharePoint Type |
|---|---|---|
| `text`, `email`, `url`, `tel`, `password`, `masked`, `autocomplete`, `taginput`, `time`, `colorpicker`, `nric`, `otp`, `hierarchy` | 2 | Text |
| `comment`, `richtext`, `addressblock`, `jsoneditor`, `locationpicker` | 3 | Note (Multi-line) |
| `date`, `datetime` | 4 | DateTime |
| `dropdown`, `radiogroup`, `buttongroup` (single) | 6 | **Choice** |
| `checkbox`, `buttongroup` (multi) | 15 | **MultiChoice** |
| `boolean`, `toggleswitch`, `consent` | 8 | Boolean/YesNo |
| `number`, `currency`, `slider`, `starrating`, `nps`, `duration`, `formula`, `unitconverter`, `counter`, `scorecard`, `rating` | 9 | Number |
| `dynamicmatrix`, `tableinput` | — | `_Html` (RichText) + `_Json` (Note) |
| `ranking`, `budgetallocator`, `rangeslider`, `daterange` | 3 | Note (stores JSON) |
| `signaturepad` | 11 | Image (URL/Hyperlink, DisplayFormat:2) |
| Layout/display types, `file`, `imageupload`, `audiorecorder` | null | No column |

### SharePoint Choice Source (`spChoicesSource`)
- Choice fields can pull values from existing SP list columns via `spChoicesSource = { list, column, multiSelect }`
- Builder UI has "Manual" / "SharePoint List" toggle in Options tab
- At runtime, choices are fetched live and injected into the SurveyJS model (`DynamicFormPage.tsx`)
- At publish time, `provisionResponseList()` fetches latest SP choices and creates the SP column with those values

## System Lists Filtering
- Both user and admin contexts see system lists filtered in `filterVisibleLists()`
- `SYSTEM_BASE_TEMPLATES` set: 109, 111, 112, 113, 114, 116, 119, 130, 140, 212, 300, 850
- Property-based exclusions: `isCatalog`, `isSiteAssetsLibrary`, `isApplicationList`, `isSystemList`, `noCrawl`
- `DiscoveredList` type includes all these fields

## Form Builder System (admin-only)
- **Entry**: Header "Form Builder" button (visible only when `isAdmin=true`) → navigates to `/admin/builder`
- **Pages**: `AdminFormBuilder.tsx` (full builder), `DynamicFormPage.tsx` (end-user rendering)
- **Components** (`src/components/builder/`): `FormBuilder.tsx` (react-dnd canvas), `FormLibrary.tsx` (sidebar), `VersionHistory.tsx`, `AuditLog.tsx`, `ApproverRow.tsx`, `ProvisionOverlay.tsx`
- **Styling**: Inline styles via `C` color object (`constants.ts`), NOT MUI components
- **SharePoint lists**: `Master Form`, `Web Form Versions`, `Form Builder Log`, `Approvers`
- **Approval layers**: 1–3 configurable, saved as `L1_Approvers`, `L2_Approvers`, etc.
- **Versioning**: `Web Form Versions` list, auto-incrementing

## Dashboard
- `Header.tsx` — Sticky top bar with user menu, role badge, admin tools, **Form Builder button**
- `ListSummaryCards.tsx` — Grid of form list cards; `onEditForm` navigates to `/admin/builder/:formTitle`
- `Toolbar.tsx` — Search + filter dropdowns (list, status, sort, submitter)
- `DetailModal.tsx` — Full detail dialog with fields, signatures, approval chain
- `ConfigWarningBanner.tsx` — Amber warning for unconfigured lists
- `mapSubmission()` in `App.tsx` filters internal fields using `/^L[1-3]_/` regex (not broad `startsWith("L")`)

## Data Views (DetailModal)
- `DetailModal.tsx` renders `submissionData` fields using `formatFieldValue()`:
  - Dates: auto-detected and locale-formatted
  - User objects: `{ Email, Title }` → displays email or title
  - Lookup fields: `{ Value }` → displays value
  - Booleans: "Yes"/"No"
  - Arrays: comma-joined
- Fields like "Location", "LeaveType" etc. are preserved (not incorrectly filtered)

## Loading Screen
- `LoadingScreen.tsx` accepts `progress` (0-100) and `status` (string) props
- `App.tsx` `fetchData()` reports 6 steps with percentage:
  - 10%: Checking permissions
  - 20%: Discovering lists
  - 35%: Loading configuration
  - 50-95%: Fetching submissions per list with `(n/total)` status
  - 98%: Finalizing
  - 100%: Ready

## Auth Components (`src/components/auth/`)
- **Shared visual pattern**: White base, 2-3 radial gradient blobs, subtle SVG accent lines, `position: absolute` with `pointerEvents: none`, `zIndex: 0`
- **Card pattern**: `elevation={0}` with custom border + shadow, top accent bar gradient, `borderRadius: 24`
- **Button patterns**: Primary `contained` (#0078D4, borderRadius 12px), Secondary `outlined`
- `ChoiceScreen.tsx` — MSAL login vs guest choice with Microsoft logo (base64 SVG)
- `GuestLanding.tsx` — Guest-only landing page with sign-in prompt
- `WrongTenantScreen.tsx` — Access denied for non-PMW tenants
- `LoadingScreen.tsx` — Progress bar with % and status subtitle during data loading
- `ErrorScreen.tsx` — Error state with retry and sign-out

## Core Modules
- `src/auth/AuthProvider.tsx` — `MsalProvider` wrapper
- `src/auth/msalConfig.ts` — MSAL configuration with dynamic `AllSites.Manage` SP scope
- `src/types/index.ts` — Shared types: `PageState`, `Submission`, `ApprovalLayer`, `ListMetaEntry`, `DiscoveredList`, `LoadedConfig`, `SharePointClient`, **`FormConfig`, `SurveyJson`, `FormVersionData`**
- `src/utils/sharepointClient.ts` — SP REST client (CRUD, digest cache, `isGroupMember`, list discovery, `resolveUserEmails`, `queryListByEmail`, `queryListByGuid`, `getSiteUsers`)
- `src/utils/spConfig.ts` — Config loader (`loadConfig` from Master Form, `filterVisibleLists`, `generateMeta`, `getMissingConfigs`)
- `src/utils/authDecision.ts` — `localStorage` auth persistence helpers (`pmw_hr_auth_decision`)
- `src/utils/formBuilderSP.ts` — **Standalone** SP REST for form builder (uses raw `token: string`, NOT `createSpClient`). Exports: `getFormConfig`, `saveFormConfig`, `getAllFormConfigs`, `saveFormVersion`, `getFormVersionHistory`, `getFormSubmissions`, `submitFormResponse`, `getSharePointChoices`, `provisionResponseList`, `bootstrapSystemLists`, `upsertApprovers`, `deleteForm`, `sendSpEmail`, `logEvent`
- `src/utils/FormBuilderEngine.ts` — Pure data logic for form building (validate, versioning, approval layers)
- `api/_utils/graphClient.ts` — Microsoft Graph API client for serverless functions (`getGraphToken`, `graphGet`, `graphPost`, `queryListItems`, `createListItem`)

## Anti-Patterns (This Project)
- **NO `useMemo`/`useCallback`** — React 19 makes these unnecessary; remove when touching code
- **NO `console.log/warn/error` in production** — replace with proper logging or remove when touching code
- **NO `dangerouslySetInnerHTML` without audit** — `DetailModal.tsx` uses it; verify XSS safety if user input reaches it
- **NO dead code** — remove unused `.backup`, `.txt`, dead pages when found. `references/` directory at root is stale JS copies — don't import from it.
- **NO `@ts-ignore` / `@ts-expect-error`** — never suppress type errors

## Conventions
- **PowerShell**: use `workdir` parameter with `bash` tool; PowerShell does NOT support `&&` or native `grep`/`ls` commands
- **File paths**: use full Windows paths or forward slashes (C:/Users/user/pmw-hrform/src/...)
- **TypeScript**: tsconfig uses project references (`tsconfig.json` → `tsconfig.app.json` + `tsconfig.node.json`). Run `tsc -b` for type-checking. `"erasableSyntaxOnly": true` — no runtime `enum`/`namespace`; use `const` objects or string unions.
- **ESLint**: NOT type-checked. Many pre-existing errors exist—focus on new errors you introduce.
- **React 19**: no `forwardRef` needed, no manual memoization (`useMemo`/`useCallback`)
- **MUI v9.0.0**: uses `Grid` (not `Grid2`); `slotProps` replaces `PaperProps` on Dialog
- **Prefer `import type`**: for type-only imports (`verbatimModuleSyntax`); many files still use plain `import` — fix when touching
- **`.env.local`** at app root contains Azure AD credentials (`VITE_AZURE_*`) and SP URL (`VITE_SP_SITE_URL`). **Never commit or expose these values.**
  - `VITE_AZURE_CLIENT_ID`, `VITE_AZURE_TENANT_ID`, `VITE_AZURE_AUTHORITY`, `VITE_SP_SITE_URL`
  - `VITE_AZURE_TENANT_ID` is used for tenant validation in auth flow
  - See `.env.example` for the required variable names (no values).
- **Verifying changes**: Run `npm run build` before claiming work is done. `npm run lint` has many pre-existing warnings—check only your changed files with `lsp_diagnostics`.

## Notes
- **No CI/CD** — zero GitHub Actions, Docker, or deployment configs
- **No tests** — zero unit, integration, or E2E tests; no vitest/jest/playwright in dependencies
- **No path aliases** — all imports use relative paths (`../../utils/...`); vite.config.ts has no `resolve.alias`

## Signature Upload Flow
Signatures are uploaded as PNG files to a `Signature Images` SharePoint document library during form submission. The flow:
1. User signs in the custom `SignaturePad.tsx` widget → base64 data URI stored in `question.value`
2. On submit (`DynamicFormPage.tsx` `onComplete`), signature fields (values starting with `data:image/`) are detected
3. `uploadSignatureImage()` in `formBuilderSP.ts` converts base64 → binary, determines daily counter, and calls `spUploadFile()` to upload to the library
4. File naming: `{action}-{formId}-{yymmdd}{xxx}.png` (e.g., `submission-HR001-260507001.png`)
5. The returned `ServerRelativeUrl` replaces the base64 value in the submission body
6. SP image column (kind 11, `DisplayFormat: 2`) renders the URL as an image

**Future actions** (`approval`, `reject`) will use the same `uploadSignatureImage()` with different action prefixes. The `action` parameter in `ApprovalLayer` and `triggerApprovalNotification` already supports these values.
