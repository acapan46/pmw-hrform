# AGENTS.md — api/

**Scope:** Vercel serverless functions. Run locally via `npm run dev:api`. Deployed to Vercel alongside the SPA.

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| Public form config | `form-config.ts` | `GET /api/form-config?slug=X[&version=Y]`. Reads Master Form + Web Form Versions via Graph API. |
| Public form submit | `submit-form.ts` | `POST /api/submit-form`. Verifies form is public, creates list item via Graph API. |
| Graph client (active) | `_utils/graphClient.ts` | Client-credentials token for `graph.microsoft.com/v1.0`. Exports: `getGraphToken`, `graphGet`, `graphPost`, `queryListItems`, `createListItem`. |
| SP REST client (dead) | `_utils/sharepoint.ts` | Dead code — not imported by any API route. All routes use `graphClient.ts` instead. |

## Conventions
- **Import paths**: API routes import from `./_utils/graphClient.ts` (relative, `_` prefix convention).
- **OData**: Uses `odata=nometadata` — responses use `data.value` not `data.d.results`.
- **CORS**: `vercel.json` adds CORS headers (`Access-Control-Allow-Origin: *`) for all `/api/*` routes.
- **Environment**: API routes run server-side in Vercel (Node.js runtime). They use `process.env` for secrets, not `import.meta.env.VITE_*`.

## Anti-Patterns
- `_utils/sharepoint.ts` — dead code (superseded by `graphClient.ts`). Safe to delete.
- `console.error` in both routes — replace with proper logging.
