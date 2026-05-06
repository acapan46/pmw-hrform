# AGENTS.md — src/pages/

**Scope:** Top-level route components. Each maps 1:1 to a route defined in `App.tsx`.

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| Admin dashboard | `AdminHomePage.tsx` | Route `/adminhomepage` and catch-all. Props: ~25 from `App.tsx` (prop-drilling). Contains dead `<Dialog>` FormBuilder (builderOpen never true). |
| Form builder page | `AdminFormBuilder.tsx` | Routes `/admin/builder[/:formTitle]`. Hosts `FormBuilder` + `FormLibrary` + sidebar with meta/form settings. Manages `showBanner`, `meta` (isoStandards, companies, logoUrl), publish flow. |
| Public form renderer | `DynamicFormPage.tsx` | Route `/form/:formId`. Auth gate bypassed for public forms. SurveyJS model + theme + submission handler. Signature upload hooked in `onComplete`. |
| Dead landing page | `HomePage.tsx` | Not imported anywhere. `ChoiceScreen.tsx` in `src/components/auth/` is the real auth landing. Safe to delete. |

## Conventions
- **Prop-drilling**: `AdminHomePage` receives massive props from `App.tsx` — no context abstraction yet.
- **Eager imports**: All pages imported statically in `App.tsx` — no `React.lazy()` or route-level code splitting.
- **No barrel export**: Import each page directly by path, e.g. `import AdminHomePage from "../pages/AdminHomePage"`.
- **Each page is self-contained**: Pages don't import from other pages.

## Anti-Patterns
- `HomePage.tsx` — dead code, 274 lines, never imported.
- `AdminHomePage.tsx` — contains dead `<Dialog>` rendering FormBuilder with `builderOpen` that is never set to true.
- `DynamicFormPage.tsx` — has `console.error`/`console.warn` calls (remove or replace with proper logging).
- `AdminFormBuilder.tsx` — has `console.error`/`console.warn` calls.
