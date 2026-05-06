# AGENTS.md — src/components/builder/

**Scope:** Custom drag-drop form builder (admin-only). NOT SurveyJS Creator — bespoke UI over `survey-react-ui`.

## WHERE TO LOOK
| Task | File | Notes |
|------|------|-------|
| Main builder UI | `FormBuilder.tsx` | react-dnd canvas, `survey-react-ui` renderer, keyboard shortcuts |
| Form list sidebar | `FormLibrary.tsx` | Available forms, create/load/delete |
| Version history | `VersionHistory.tsx` | Web Form Versions list, restore prior version |
| Audit log | `AuditLog.tsx` | Form Builder Log entries with diff view |
| Approver input | `ApproverRow.tsx` | User search + L1/L2/L3 approver assignment |
| Publish overlay | `ProvisionOverlay.tsx` | SharePoint list provisioning status spinner |
| Barrel exports | `index.ts` | Only barrel export in the entire app |
| Colors | `constants.ts` | `C` color object — inline styles, NOT MUI theme |

## Builder Architecture
```
AdminFormBuilder.tsx (page — route: /admin/builder)
  ├── FormLibrary (sidebar — fetches from Master Form SP list)
  ├── FormBuilder.tsx (canvas — react-dnd drag-drop)
  │     ├── Palette (question types sidebar)
  │     ├── Canvas (field cards with reorder)
  │     ├── PropertyPanel (per-field settings)
  │     ├── JsonPreview (collapsed JSON output)
  │     └── LivePreviewModal (survey-react-ui renderer)
  ├── VersionHistory (side panel — Web Form Versions SP list)
  ├── AuditLog (side panel — Form Builder Log SP list)
  ├── ApproverRow[] (L1-L3 assignment)
  └── ProvisionOverlay (publish status spinner)
```

## Conventions
- **Styling**: Inline styles via `C` object — no MUI components, no CSS modules
- **State**: Local `useState` only — no context or external store
- **Barrel exports**: `index.ts` re-exports all builder components; import from `components/builder`

## Form Builder Features
- **SP Choice Source**: Choice fields (dropdown, radiogroup, checkbox, buttongroup) can pull values from existing SharePoint list columns via `spChoicesSource` property. Toggle "Manual" / "SharePoint List" in Options tab.
- **Matrix Column Editor**: `dynamicmatrix` fields have a per-column editor in Options tab — set cell type (text/dropdown/date/number/checkbox/boolean), manual choices, or SP choice source per column.
- **Signature Widget**: `signaturepad` uses a custom modal-based widget (`src/utils/SignaturePad.tsx`) registered via `ReactQuestionFactory`. Click to open modal → sign → save/lock. Image stored as base64, uploaded to `Signature Images` doc library on submit.
- **Logo Setter**: Banner logo URL configurable in the form meta sidebar; defaults to `/logo-128.png`.
- **Form Title Toggle**: "Show form title" toggle in Form Settings controls SurveyJS title visibility (sets `titleLocation: "hidden"`). Title is centered via CSS when visible.

## Anti-Patterns
- `FormBuilder.tsx` has `eslint-disable` and `any[]` usage — fix when touching
- `useMemo`/`useCallback` used extensively — unnecessary in React 19; remove when refactoring
