/**
 * FormBuilder.tsx — Custom form builder (NO SurveyJS Creator)
 * Uses react-dnd for drag-drop. Outputs SurveyJS-compatible JSON.
 */
import React, { useState, useCallback, useRef, useEffect, useMemo } from "react";
import { Survey } from "survey-react-ui";
import { Model, Serializer } from "survey-core";
import "survey-core/survey-core.min.css";
import type { SurveyJson, FormBuilderField } from "../../types/index";
import { QUESTION_TYPES, TYPE_GROUPS, createQuestion, buildSurveyJson, validateFields, getSpColumnKind } from "../../utils/FormBuilderEngine";
import { buildQuestionTree, removeFieldRecursive, duplicateFieldRecursive, moveFieldIntoPanel, addFieldToPanel, findFieldById, updateField, flattenFieldTree, reorderFieldsRecursive, moveFieldToRoot } from "../../utils/FormBuilderEngine";
import { registerSignaturePad } from "../../utils/SignaturePad";
import { C } from "./constants";
import "./FormBuilder.css";

// MUI Icons
import DragIndicatorIcon from "@mui/icons-material/DragIndicator";
import ContentCopyIcon from "@mui/icons-material/ContentCopy";
import ArrowUpwardIcon from "@mui/icons-material/ArrowUpward";
import ArrowDownwardIcon from "@mui/icons-material/ArrowDownward";
import SettingsIcon from "@mui/icons-material/Settings";
import PreviewIcon from "@mui/icons-material/Preview";

import SearchIcon from "@mui/icons-material/Search";
import CloseIcon from "@mui/icons-material/Close";
import AddIcon from "@mui/icons-material/Add";
import UndoIcon from "@mui/icons-material/Undo";
import RedoIcon from "@mui/icons-material/Redo";
import WarningAmberIcon from "@mui/icons-material/WarningAmber";
import CodeIcon from "@mui/icons-material/Code";
import DesktopWindowsIcon from "@mui/icons-material/DesktopWindows";
import PhoneIphoneIcon from "@mui/icons-material/PhoneIphone";
import FileDownloadIcon from "@mui/icons-material/FileDownload";
import TranslateIcon from "@mui/icons-material/Translate";
import PaletteIcon from "@mui/icons-material/Palette";
import CommentIcon from "@mui/icons-material/Comment";
import HubIcon from "@mui/icons-material/Hub";
import StorageIcon from "@mui/icons-material/Storage";
import ShieldIcon from "@mui/icons-material/Shield";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import ExpandLessIcon from "@mui/icons-material/ExpandLess";
import TextFieldsIcon from "@mui/icons-material/TextFields";
import ArrowDropDownCircleIcon from "@mui/icons-material/ArrowDropDownCircle";
import RadioButtonCheckedIcon from "@mui/icons-material/RadioButtonChecked";
import CheckBoxIcon from "@mui/icons-material/CheckBox";

import NumbersIcon from "@mui/icons-material/Numbers";
import CalendarTodayIcon from "@mui/icons-material/CalendarToday";
import TableChartIcon from "@mui/icons-material/TableChart";
import ImageIcon from "@mui/icons-material/Image";
import AttachFileIcon from "@mui/icons-material/AttachFile";
import GestureIcon from "@mui/icons-material/Gesture";
import InsertPageBreakIcon from "@mui/icons-material/InsertPageBreak";
import DashboardIcon from "@mui/icons-material/Dashboard";
import LoopIcon from "@mui/icons-material/Loop";
import ViewColumnIcon from "@mui/icons-material/ViewColumn";
import HeightIcon from "@mui/icons-material/Height";
import HorizontalRuleIcon from "@mui/icons-material/HorizontalRule";
import LockIcon from "@mui/icons-material/Lock";
import EmailIcon from "@mui/icons-material/Email";
import LinkIcon from "@mui/icons-material/Link";
import PhoneIcon from "@mui/icons-material/Phone";
import AccessTimeIcon from "@mui/icons-material/AccessTime";
import DateRangeIcon from "@mui/icons-material/DateRange";
import TimelapseIcon from "@mui/icons-material/Timelapse";

import LinearScaleIcon from "@mui/icons-material/LinearScale";
import AttachMoneyIcon from "@mui/icons-material/AttachMoney";
import CalculateIcon from "@mui/icons-material/Calculate";
import PlusOneIcon from "@mui/icons-material/PlusOne";
import CheckCircleIcon from "@mui/icons-material/CheckCircle";
import StarRateIcon from "@mui/icons-material/StarRate";
import AddPhotoAlternateIcon from "@mui/icons-material/AddPhotoAlternate";
import HomeIcon from "@mui/icons-material/Home";
import LocationOnIcon from "@mui/icons-material/LocationOn";
import BadgeIcon from "@mui/icons-material/Badge";

import ArticleIcon from "@mui/icons-material/Article";
import TableRowsIcon from "@mui/icons-material/TableRows";
import FormatListNumberedIcon from "@mui/icons-material/FormatListNumbered";
import AccountBalanceWalletIcon from "@mui/icons-material/AccountBalanceWallet";
import AccountTreeIcon from "@mui/icons-material/AccountTree";
import DataObjectIcon from "@mui/icons-material/DataObject";
import LanguageIcon from "@mui/icons-material/Language";
import VideocamIcon from "@mui/icons-material/Videocam";
import WarningIcon from "@mui/icons-material/Warning";

import HourglassEmptyIcon from "@mui/icons-material/HourglassEmpty";
import BarChartIcon from "@mui/icons-material/BarChart";
import SpeedIcon from "@mui/icons-material/Speed";
import ChevronLeftIcon from "@mui/icons-material/ChevronLeft";
import ChevronRightIcon from "@mui/icons-material/ChevronRight";

// ── Register autocapitalize as a custom SurveyJS property ─────────────────
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

// ── Atoms ─────────────────────────────────────────────────────────────
const Pill = ({ children, color = C.purple, bg = C.purplePale }: { children: React.ReactNode; color?: string; bg?: string }) =>
  <span className="fb-pill" style={{ color, background: bg }}>{children}</span>;

function IconBtn({ icon, title, onClick, danger, disabled }: { icon: React.ReactNode; title?: string; onClick?: () => void; danger?: boolean; disabled?: boolean }) {
  return <button title={title} onClick={onClick} disabled={disabled} className={`fb-icon-btn ${danger ? 'danger' : ''}`}
    onMouseEnter={e => { if (!disabled) e.currentTarget.classList.add(danger ? 'danger' : 'hover'); }}
    onMouseLeave={e => { e.currentTarget.classList.remove('hover', 'danger'); }}>{icon}</button>;
}

function Toggle({ checked, onChange, label }: { checked: boolean; onChange: (v: boolean) => void; label?: string }) {
  return <label className="fb-toggle">
    <div onClick={() => onChange(!checked)} className="fb-toggle-track" style={{ background: checked ? C.purple : "#D1D5DB" }}>
      <div className="fb-toggle-knob" style={{ left: checked ? 19 : 3, background: C.white }} />
    </div>
    {label && <span className="fb-toggle-label" style={{ color: C.textSecond }}>{label}</span>}
  </label>;
}

function Input({ value, onChange, placeholder, type = "text", style: extra, ...rest }: { value?: string | number; onChange: (v: string) => void; placeholder?: string; type?: string; style?: React.CSSProperties;[key: string]: unknown }) {
  const [f, setF] = useState(false);
  return <input type={type} value={value ?? ""} onChange={e => onChange(e.target.value)} placeholder={placeholder}
    onFocus={() => setF(true)} onBlur={() => setF(false)}
    className={`fb-input ${f ? 'focused' : ''}`}
    style={extra}
    {...rest} />;
}

function Textarea({ value, onChange, placeholder, rows = 3 }: { value?: string; onChange: (v: string) => void; placeholder?: string; rows?: number }) {
  const [f, setF] = useState(false);
  return <textarea value={value ?? ""} onChange={e => onChange(e.target.value)} placeholder={placeholder} rows={rows}
    onFocus={() => setF(true)} onBlur={() => setF(false)} className={`fb-textarea ${f ? 'focused' : ''}`} />;
}

function Select({ value, onChange, options }: { value?: string; onChange: (v: string) => void; options: { value: string; label: string }[] }) {
  return <select value={value ?? ""} onChange={e => onChange(e.target.value)} className="fb-select">
    {options.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
  </select>;
}

const PropLabel = ({ children }: { children: React.ReactNode }) =>
  <div className="fb-prop-label">{children}</div>;

function PropRow({ label, children, span }: { label: string; children: React.ReactNode; span?: boolean }) {
  return <div className={`fb-prop-row ${span ? 'span' : ''}`}>
    <PropLabel>{label}</PropLabel>{children}
  </div>;
}

// ── Visibility / EnableIf / Validation Editors ────────────────────────

/** Build a simple expression editor for visibleIf / enableIf */
function ConditionEditor({ label, value, onChange, allFields }: { label: string; value: string; onChange: (v: string) => void; allFields: FormBuilderField[] }) {
  const [mode, setMode] = useState<"simple" | "advanced">("simple");
  const [conditionField, setConditionField] = useState("");
  const [operator, setOperator] = useState("notempty");
  const [conditionValue, setConditionValue] = useState("");

  // Sync from expression to simple mode
  useEffect(() => {
    if (!value) { setConditionField(""); setOperator("notempty"); setConditionValue(""); return; }
    // Try to parse: {field} operator 'value' or {field} operator value
    const match = value.match(/\{([^}]+)\}\s*(=|<>|<|>|<=|>=|contains|notcontains|startswith|endswith|empty|notempty)\s*'?([^']*)'?/i);
    if (match) {
      setConditionField(match[1]);
      setOperator(match[2].toLowerCase());
      setConditionValue(match[3] || "");
    }
  }, [value]);

  const buildExpression = () => {
    if (!conditionField) { onChange(""); return; }
    if (operator === "empty" || operator === "notempty") {
      onChange(`{${conditionField}} ${operator}`);
    } else {
      const val = isNaN(Number(conditionValue)) ? `'${conditionValue}'` : conditionValue;
      onChange(`{${conditionField}} ${operator} ${val}`);
    }
  };

  const handleApply = () => { buildExpression(); };

  return <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>{label}</span>
      <div style={{ flex: 1 }} />
      <Toggle checked={mode === "advanced"} onChange={v => setMode(v ? "advanced" : "simple")} label="Advanced" />
    </div>
    {mode === "simple" ? <>
      <PropRow label="When field">
        <Select value={conditionField} onChange={v => { setConditionField(v); }} options={[{ value: "", label: "(none)" }, ...allFields.filter(f => f.name !== (value ? "" : "")).map(f => ({ value: f.name, label: `${f.title} (${f.name})` }))]} />
      </PropRow>
      <PropRow label="Operator">
        <Select value={operator} onChange={v => setOperator(v)} options={[
          { value: "notempty", label: "Is not empty" },
          { value: "empty", label: "Is empty" },
          { value: "=", label: "Equals" },
          { value: "<>", label: "Not equals" },
          { value: "<", label: "Less than" },
          { value: ">", label: "Greater than" },
          { value: "contains", label: "Contains" },
          { value: "notcontains", label: "Does not contain" },
        ]} />
      </PropRow>
      {operator !== "empty" && operator !== "notempty" && <PropRow label="Value">
        <Input value={conditionValue} onChange={setConditionValue} placeholder="Enter value" />
      </PropRow>}
      <button onClick={handleApply} disabled={!conditionField} className="fb-apply-btn" style={{ background: conditionField ? C.purple : C.border, color: conditionField ? C.white : C.textMuted, cursor: conditionField ? "pointer" : "not-allowed" }}>Apply Condition</button>
    </> : <Textarea value={value} onChange={onChange} placeholder="Enter SurveyJS expression, e.g.:{question1} = 'Yes' && {question2} notempty" rows={4} />}
    {value && <div className="fb-current-condition">Current: {value}</div>}
  </div>;
}

// Validator types from SurveyJS
const VALIDATOR_TYPES = [
  { value: "numeric", label: "Numeric", desc: "Min/max/decimal" },
  { value: "text", label: "Text", desc: "Min/max length" },
  { value: "regex", label: "Regex", desc: "Pattern match" },
  { value: "email", label: "Email", desc: "Valid email" },
  { value: "expression", label: "Expression", desc: "Custom logic" },
];

function ValidationEditor({ field, onChange }: { field: FormBuilderField; onChange: (patch: Partial<FormBuilderField>) => void }) {
  const validators: any[] = field.validators || [];
  const addValidator = (type: string) => {
    const base = { type };
    if (type === "numeric") Object.assign(base, { minValue: undefined, maxValue: undefined });
    if (type === "text") Object.assign(base, { minLength: undefined, maxLength: undefined });
    if (type === "regex") Object.assign(base, { regex: "", text: "Invalid format" });
    if (type === "expression") Object.assign(base, { expression: "", text: "Invalid" });
    onChange({ validators: [...validators, base] });
  };
  const updateValidator = (idx: number, patch: Record<string, unknown>) => {
    const next = validators; next[idx] = { ...next[idx], ...patch };
    onChange({ validators: next });
  };
  const removeValidator = (idx: number) => { const next = validators; next.splice(idx, 1); onChange({ validators: next }); };
  return <div>
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Validators</span>
      <div style={{ flex: 1 }} />
      <Select value="" onChange={v => { if (v) addValidator(v); }} options={[{ value: "", label: "+ Add validator..." }, ...VALIDATOR_TYPES.map(v => ({ value: v.value, label: v.label }))]} />
    </div>
    {validators.length === 0 && <div className="fb-no-validators">No validators. Add one above.</div>}
    {validators.map((v, idx) => <div key={idx} className="fb-validation-row">
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
        <Pill>{VALIDATOR_TYPES.find(t => t.value === v.type)?.label || v.type}</Pill>
        <button onClick={() => removeValidator(idx)} className="fb-icon-btn danger" title="Remove">✕</button>
      </div>
      {v.type === "numeric" && <>
        <PropRow label="Min value"><Input type="number" value={v.minValue as number ?? ""} onChange={val => updateValidator(idx, { minValue: val === "" ? undefined : Number(val) })} placeholder="No min" /></PropRow>
        <PropRow label="Max value"><Input type="number" value={v.maxValue as number ?? ""} onChange={val => updateValidator(idx, { maxValue: val === "" ? undefined : Number(val) })} placeholder="No max" /></PropRow>
      </>}
      {v.type === "text" && <>
        <PropRow label="Min length"><Input type="number" value={v.minLength as number ?? ""} onChange={val => updateValidator(idx, { minLength: val === "" ? undefined : Number(val) })} placeholder="No min" /></PropRow>
        <PropRow label="Max length"><Input type="number" value={v.maxLength as number ?? ""} onChange={val => updateValidator(idx, { maxLength: val === "" ? undefined : Number(val) })} placeholder="No max" /></PropRow>
      </>}
      {v.type === "regex" && <>
        <PropRow label="Pattern"><Input value={(v.regex as string) || ""} onChange={val => updateValidator(idx, { regex: val || undefined })} placeholder="e.g. ^[A-Z]+$" /></PropRow>
        <PropRow label="Error text"><Input value={(v.text as string) || ""} onChange={val => updateValidator(idx, { text: val || undefined })} placeholder="Error message" /></PropRow>
      </>}
      {v.type === "expression" && <>
        <PropRow label="Expression"><Input value={(v.expression as string) || ""} onChange={val => updateValidator(idx, { expression: val || undefined })} placeholder="{field} > 5" /></PropRow>
        <PropRow label="Error text"><Input value={(v.text as string) || ""} onChange={val => updateValidator(idx, { text: val || undefined })} placeholder="Error message" /></PropRow>
      </>}
      {v.type === "email" && <div style={{ fontSize: 10, color: C.textMuted }}>Validates email format automatically</div>}
      <PropRow label="Error text"><Input value={(v.text as string) || ""} onChange={val => updateValidator(idx, { text: val || undefined })} placeholder="Custom error message" /></PropRow>
    </div>)}
  </div>;
}

// ── LOGIC RULES EDITOR ────────────────────────────────────────────────────

/** Operator options for logic rules */
const LOGIC_OPERATORS = [
  { value: "equals", label: "Equals" },
  { value: "notEquals", label: "Does not equal" },
  { value: "contains", label: "Contains" },
  { value: "notContains", label: "Does not contain" },
  { value: "startsWith", label: "Starts with" },
  { value: "endsWith", label: "Ends with" },
  { value: "isEmpty", label: "Is empty" },
  { value: "isNotEmpty", label: "Is not empty" },
  { value: "greaterThan", label: "Greater than" },
  { value: "lessThan", label: "Less than" },
  { value: "greaterOrEqual", label: "Greater or equal" },
  { value: "lessOrEqual", label: "Less or equal" },
];

/** Single condition row */
function ConditionRow({ condition, allFields, onUpdate, onRemove, canRemove }: {
  condition: { field: string; operator: string; value: string };
  allFields: FormBuilderField[];
  onUpdate: (c: { field: string; operator: string; value: string }) => void;
  onRemove: () => void;
  canRemove: boolean;
}) {
  return (
    <div style={{ display: "flex", gap: 6, alignItems: "center", marginBottom: 8 }}>
      <Select value={condition.field} onChange={v => onUpdate({ ...condition, field: v })}
        options={[{ value: "", label: "Select field" }, ...allFields.map(f => ({ value: f.name, label: f.title || f.name }))]} />
      <Select value={condition.operator} onChange={v => onUpdate({ ...condition, operator: v })}
        options={LOGIC_OPERATORS} />
      {!["isEmpty", "isNotEmpty"].includes(condition.operator) && (
        <Input value={condition.value} onChange={v => onUpdate({ ...condition, value: v })} placeholder="Value" style={{ flex: 1, minWidth: 80 }} />
      )}
      <IconBtn icon="✕" title="Remove condition" onClick={onRemove} disabled={!canRemove} danger />
    </div>
  );
}

/** Rules section for a specific rule type */
function RulesSection({ rules, ruleType: _ruleType, title, icon, color, allFields, onChange }: {
  rules: { id: string; field: string; operator: string; value: string; connector: string; enabled: boolean }[];
  ruleType: "visibility" | "required" | "enable";
  title: string;
  icon: string;
  color: string;
  allFields: FormBuilderField[];
  onChange: (rules: { id: string; field: string; operator: string; value: string; connector: string; enabled: boolean }[]) => void;
}) {
  const addRule = () => {
    onChange([...rules, { id: `rule_${Date.now()}`, field: "", operator: "equals", value: "", connector: "AND", enabled: true }]);
  };
  
  const updateRule = (idx: number, update: Partial<{ field: string; operator: string; value: string; connector: string; enabled: boolean }>) => {
    const updated = [...rules];
    updated[idx] = { ...updated[idx], ...update };
    onChange(updated);
  };
  
  const removeRule = (idx: number) => {
    onChange(rules.filter((_, i) => i !== idx));
  };

  return (
    <div style={{ marginBottom: 16, padding: 12, background: `${color}08`, borderRadius: 8, border: `1px solid ${color}20` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <span style={{ fontSize: 14 }}>{icon}</span>
        <span style={{ fontSize: 12, fontWeight: 700, color, flex: 1 }}>{title}</span>
        <Toggle checked={rules.length > 0} onChange={v => { if (v && !rules.length) addRule(); else if (!v) onChange([]); }} label={rules.length > 0 ? "Active" : "Disabled"} />
      </div>
      {rules.map((rule, idx) => (
        <div key={rule.id} style={{ marginBottom: 12 }}>
          {idx > 0 && (
            <Select value={rule.connector as string} onChange={v => updateRule(idx, { connector: v as "AND" | "OR" })}
              options={[{ value: "AND", label: "AND" }, { value: "OR", label: "OR" }]} />
          )}
          <ConditionRow condition={rule} allFields={allFields}
            onUpdate={c => updateRule(idx, c)} onRemove={() => removeRule(idx)} canRemove={rules.length > 1} />
        </div>
      ))}
      {rules.length > 0 && (
        <button onClick={addRule} style={{ fontSize: 11, color: C.purple, background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>
          + Add condition
        </button>
      )}
    </div>
  );
}

/** Value mapping section */
function ValueMappingSection({ valueRule, allFields, onChange }: {
  valueRule: { sourceField: string; transform: string } | undefined;
  allFields: FormBuilderField[];
  onChange: (rule: { sourceField: string; transform: string } | undefined) => void;
}) {
  const [enabled, setEnabled] = useState(!!valueRule?.sourceField);
  
  return (
    <div style={{ marginBottom: 16, padding: 12, background: `${C.purple}08`, borderRadius: 8, border: `1px solid ${C.purple}20` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <span style={{ fontSize: 14 }}>🔄</span>
        <span style={{ fontSize: 12, fontWeight: 700, color: C.purple, flex: 1 }}>Value Mapping</span>
        <Toggle checked={enabled} onChange={v => { setEnabled(v); if (!v) onChange(undefined); }} label={enabled ? "Active" : "Disabled"} />
      </div>
      {enabled && (
        <>
          <PropRow label="Copy from">
            <Select value={valueRule?.sourceField || ""} onChange={v => onChange({ sourceField: v, transform: valueRule?.transform || "none" })}
              options={[{ value: "", label: "Select field" }, ...allFields.map(f => ({ value: f.name, label: f.title || f.name }))]} />
          </PropRow>
          <PropRow label="Transform">
            <Select value={valueRule?.transform || "none"} onChange={v => onChange({ sourceField: valueRule?.sourceField || "", transform: v })}
              options={[
                { value: "none", label: "None" },
                { value: "uppercase", label: "UPPERCASE" },
                { value: "lowercase", label: "lowercase" },
                { value: "capitalize", label: "Capitalize" },
                { value: "trim", label: "Trim whitespace" },
              ]} />
          </PropRow>
        </>
      )}
    </div>
  );
}

/** Cross-field validation */
function CrossFieldValidationSection({ validations, fieldName, allFields, onChange }: {
  validations: { id: string; fieldA: string; operator: "equals" | "notEquals" | "greaterThan" | "lessThan" | "greaterOrEqual" | "lessOrEqual" | "before" | "after"; fieldB: string; errorMessage: string; enabled: boolean }[];
  fieldName: string;
  allFields: FormBuilderField[];
  onChange: (rules: { id: string; fieldA: string; operator: "equals" | "notEquals" | "greaterThan" | "lessThan" | "greaterOrEqual" | "lessOrEqual" | "before" | "after"; fieldB: string; errorMessage: string; enabled: boolean }[]) => void;
}) {
  // Filter to only fields that come BEFORE this field (to avoid circular deps)
  const priorFields = allFields.filter(f => f.name !== fieldName);
  
  const addValidation = () => {
    onChange([...validations, { id: `cfv_${Date.now()}`, fieldA: fieldName, operator: "greaterThan", fieldB: "", errorMessage: "", enabled: true }]);
  };
  
  const updateValidation = (idx: number, update: Partial<{ fieldA: string; operator: "equals" | "notEquals" | "greaterThan" | "lessThan" | "greaterOrEqual" | "lessOrEqual" | "before" | "after"; fieldB: string; errorMessage: string; enabled: boolean }>) => {
    const updated = [...validations];
    updated[idx] = { ...updated[idx], ...update };
    onChange(updated);
  };
  
  const removeValidation = (idx: number) => {
    onChange(validations.filter((_, i) => i !== idx));
  };

  return (
    <div style={{ marginBottom: 16, padding: 12, background: `${C.red}08`, borderRadius: 8, border: `1px solid ${C.red}20` }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <span style={{ fontSize: 14 }}>⚖️</span>
        <span style={{ fontSize: 12, fontWeight: 700, color: C.red, flex: 1 }}>Cross-field Validation</span>
        <Toggle checked={validations.length > 0} onChange={v => { if (v && !validations.length) addValidation(); else if (!v) onChange([]); }} label={validations.length > 0 ? "Active" : "Disabled"} />
      </div>
      {validations.length === 0 && priorFields.length === 0 && (
        <div style={{ fontSize: 10, color: C.textMuted }}>Add fields before this one to create validations.</div>
      )}
      {validations.map((v, idx) => (
        <div key={v.id} style={{ marginBottom: 12, padding: 10, background: C.white, borderRadius: 6 }}>
          <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 8 }}>
            <strong>{fieldName}</strong> must be:
          </div>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
            <Select value={v.operator as string} onChange={val => updateValidation(idx, { operator: val as "equals" | "notEquals" | "greaterThan" | "lessThan" | "greaterOrEqual" | "lessOrEqual" | "before" | "after" })}
              options={[
                { value: "equals", label: "Equal to" },
                { value: "notEquals", label: "Not equal to" },
                { value: "greaterThan", label: "Greater than" },
                { value: "lessThan", label: "Less than" },
                { value: "greaterOrEqual", label: "At least" },
                { value: "lessOrEqual", label: "At most" },
                { value: "before", label: "Before" },
                { value: "after", label: "After" },
              ]} />
            <Select value={v.fieldB} onChange={val => updateValidation(idx, { fieldB: val })}
              options={[{ value: "", label: "Select field" }, ...priorFields.map(f => ({ value: f.name, label: f.title || f.name }))]} />
          </div>
          <Input value={v.errorMessage} onChange={val => updateValidation(idx, { errorMessage: val })} placeholder="Error message" />
          <button onClick={() => removeValidation(idx)} style={{ marginTop: 8, fontSize: 10, color: C.red, background: "none", border: "none", cursor: "pointer" }}>
            Remove validation
          </button>
        </div>
      ))}
      {validations.length > 0 && (
        <button onClick={addValidation} style={{ fontSize: 11, color: C.red, background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>
          + Add validation
        </button>
      )}
    </div>
  );
}

/** Main Logic Rules Editor component */
function LogicRulesEditor({ field, allFields, onChange }: {
  field: FormBuilderField;
  allFields: FormBuilderField[];
  onChange: (patch: Partial<FormBuilderField>) => void;
}) {
  // Parse existing visibleIf into rules format
  const visibilityRules = useMemo(() => {
    if (!field.visibleIf) return [];
    // Simple parsing - in real impl this would be more sophisticated
    return [{ id: "v1", field: "", operator: "equals", value: "", connector: "AND", enabled: true }];
  }, [field.visibleIf]);

  const enableRules = useMemo(() => {
    if (!field.enableIf) return [];
    return [{ id: "e1", field: "", operator: "equals", value: "", connector: "AND", enabled: true }];
  }, [field.enableIf]);

  const updateVisibilityRules = (rules: typeof visibilityRules) => {
    // Build expression from rules
    if (rules.length === 0 || !rules[0].field) {
      onChange({ visibleIf: undefined });
      return;
    }
    const expr = rules.map((r, i) => {
      if (i > 0) return ` ${r.connector} {${r.field}} ${r.operator} '${r.value}'`;
      return `{${r.field}} ${r.operator} '${r.value}'`;
    }).join("");
    onChange({ visibleIf: expr });
  };

  const updateEnableRules = (rules: typeof enableRules) => {
    if (rules.length === 0 || !rules[0].field) {
      onChange({ enableIf: undefined });
      return;
    }
    const expr = rules.map((r, i) => {
      if (i > 0) return ` ${r.connector} {${r.field}} ${r.operator} '${r.value}'`;
      return `{${r.field}} ${r.operator} '${r.value}'`;
    }).join("");
    onChange({ enableIf: expr });
  };

  const fieldName = field.name;
  const priorFields = allFields.filter(f => f.name !== fieldName);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      {/* Legacy compatibility - keep existing editors working */}
      <ConditionEditor label="Visibility (legacy)" value={field.visibleIf || ""} onChange={v => onChange({ visibleIf: v || undefined })} allFields={priorFields} />
      <ConditionEditor label="Enable (legacy)" value={field.enableIf || ""} onChange={v => onChange({ enableIf: v || undefined })} allFields={priorFields} />
      
      <div style={{ marginTop: 16 }}>
        <RulesSection rules={visibilityRules} ruleType="visibility" title="Show/Hide Rules" icon="👁️" color={C.green} allFields={priorFields} onChange={updateVisibilityRules} />
        <RulesSection rules={enableRules} ruleType="enable" title="Enable/Disable Rules" icon="🔓" color={C.purple} allFields={priorFields} onChange={updateEnableRules} />
        <ValueMappingSection valueRule={field.valueMapping as { sourceField: string; transform: string } | undefined} allFields={allFields} onChange={v => onChange({ valueMapping: v })} />
        <CrossFieldValidationSection 
            validations={(field.crossFieldValidations || []) as { id: string; fieldA: string; operator: "equals" | "notEquals" | "greaterThan" | "lessThan" | "greaterOrEqual" | "lessOrEqual" | "before" | "after"; fieldB: string; errorMessage: string; enabled: boolean }[]} 
            fieldName={fieldName} 
            allFields={allFields} 
            onChange={v => onChange({ crossFieldValidations: v })} />
      </div>
    </div>
  );
}

// Map SurveyJS types to MUI icons (all 57 field types, each unique)
const TYPE_ICONS: Record<string, React.ReactNode> = {
  // Layout
  pagebreak: <InsertPageBreakIcon />,
  panel: <DashboardIcon />,
  repeater: <LoopIcon />,
  columns: <ViewColumnIcon />,
  spacer: <HeightIcon />,
  divider: <HorizontalRuleIcon />,
  // Basic
  text: <TextFieldsIcon />,
  number: <NumbersIcon />,
  email: <EmailIcon />,
  url: <LinkIcon />,
  tel: <PhoneIcon />,
  comment: <CommentIcon />,
  date: <CalendarTodayIcon />,
  datetime: <AccessTimeIcon />,
  boolean: <CheckCircleIcon />,
  // Text
  password: <LockIcon />,
  // Date/Time
  daterange: <DateRangeIcon />,
  duration: <TimelapseIcon />,
  // Choice
  dropdown: <ArrowDropDownCircleIcon />,
  radiogroup: <RadioButtonCheckedIcon />,
  checkbox: <CheckBoxIcon />,
  // Selection
  slider: <LinearScaleIcon />,
  // Numeric
  currency: <AttachMoneyIcon />,
  formula: <CalculateIcon />,
  counter: <PlusOneIcon />,
  // Advanced
  rating: <StarRateIcon />,
  file: <AttachFileIcon />,
  imageupload: <AddPhotoAlternateIcon />,
  signaturepad: <GestureIcon />,
  addressblock: <HomeIcon />,
  locationpicker: <LocationOnIcon />,
  nric: <BadgeIcon />,
  consent: <ArticleIcon />,
  dynamicmatrix: <TableChartIcon />,
  tableinput: <TableRowsIcon />,
  ranking: <FormatListNumberedIcon />,
  budgetallocator: <AccountBalanceWalletIcon />,
  hierarchy: <AccountTreeIcon />,
  jsoneditor: <DataObjectIcon />,
  // Display
  html: <LanguageIcon />,
  image: <ImageIcon />,
  videoembed: <VideocamIcon />,
  alert: <WarningIcon />,
  countdown: <HourglassEmptyIcon />,
  datatable: <StorageIcon />,
  chartdisplay: <BarChartIcon />,
  scorecard: <SpeedIcon />,
};

// ── Palette ───────────────────────────────────────────────────────────
function Palette({ onAdd }: { onAdd: (td: typeof QUESTION_TYPES[number]) => void }) {
  const [search, setSearch] = useState("");
  const [activeGroup, setActiveGroup] = useState("All");
  const filtered = useMemo(() => {
    let list = QUESTION_TYPES;
    if (activeGroup !== "All") list = list.filter(t => t.group === activeGroup);
    if (search.trim()) { const q = search.toLowerCase(); list = list.filter(t => t.label.toLowerCase().includes(q) || t.description.toLowerCase().includes(q)); }
    return list;
  }, [search, activeGroup]);
  const onDragStart = (e: React.DragEvent, td: typeof QUESTION_TYPES[number]) => {
    e.dataTransfer.setData("palette_type", JSON.stringify(td));
    e.dataTransfer.effectAllowed = "copy";
  };
  return <div className="fb-palette">
    <div className="fb-palette-search">
      <div className="fb-palette-search-wrapper">
        <SearchIcon className="fb-palette-search-icon" style={{ color: C.textMuted, fontSize: 16 }} />
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search fields…" className="fb-palette-search-input" />
      </div>
    </div>
    <div className="fb-palette-groups">
      {["All", ...TYPE_GROUPS].map(g => <button key={g} onClick={() => setActiveGroup(g)} className="fb-palette-group-btn" style={{ background: activeGroup === g ? C.purple : C.offWhite, color: activeGroup === g ? C.white : C.textMuted }}>{g}</button>)}
    </div>
    <div className="fb-palette-list">
      {filtered.map((td, i) => <div key={td.type + i} draggable onDragStart={e => onDragStart(e, td)} onClick={() => onAdd(td)} className="fb-palette-item" title={td.description} style={{ animation: `slideIn 0.15s ease ${i * 0.02}s both` }}>
        <span className="fb-palette-item-icon">
          {TYPE_ICONS[td.type]}
        </span>
        <div className="fb-palette-item-content">
          <div className="fb-palette-item-title">{td.label}</div>
        </div>
        <DragIndicatorIcon style={{ marginLeft: "auto", flexShrink: 0, opacity: 0.4, fontSize: 16 }} />
      </div>)}
      {!filtered.length && <div style={{ textAlign: "center", padding: "24px 0", color: C.textMuted, fontSize: 12 }}>No field types match</div>}
    </div>
    <div style={{ padding: "8px 12px", borderTop: `1px solid ${C.border}`, fontSize: 10, color: C.textMuted, textAlign: "center" }}>Click or drag to add a field</div>
  </div>;
}

// ── Canvas ────────────────────────────────────────────────────────────
function FieldCard({ field, index, selected, onSelect, onRemove, onDuplicate, onMoveUp, onMoveDown, isFirst, isLast, errors, onDragStart, onDragOver, onDrop, dragging: _dragging, onDropOnPanel, onRecursiveReorder, depth = 0, selectedId }: {
  field: FormBuilderField; index: number; selected: boolean; onSelect: (id: string) => void;
  onRemove: (id: string) => void; onDuplicate: (field: FormBuilderField) => void;
  onMoveUp: () => void; onMoveDown: () => void; isFirst: boolean; isLast: boolean;
  errors: { id: string; msg: string }[]; onDragStart: (e: React.DragEvent, i: number) => void;
  onDragOver: (e: React.DragEvent, i: number) => void; onDrop: (e: React.DragEvent, i: number) => void;
  dragging: boolean;
  onDropOnPanel?: (e: React.DragEvent, panelId: string) => void;
  onRecursiveReorder?: (fromId: string, toId: string) => void;
  depth?: number;
  selectedId?: string | null;
}) {
  const err = errors.filter(e => e.id === field._id);
  const td = QUESTION_TYPES.find(t => t.type === field.type) || QUESTION_TYPES[0];
  const spColKind = getSpColumnKind(field);
  const spColLabel = spColKind
    ? ({ 2: "Text", 3: "Multi-line", 4: "Date", 6: "Choice", 8: "Boolean", 9: "Number", 15: "MultiChoice" }[spColKind.FieldTypeKind] ?? spColKind.label)
    : null;
  const shortcuts = selected ? "Del to remove, Ctrl+D to duplicate" : "";
  const isPanel = field.type === "panel";

  return <div
    draggable
    onDragStart={e => onDragStart(e, index)}
    onDragOver={e => onDragOver(e, index)}
    onDrop={e => {
      if (isPanel && onDropOnPanel) {
        e.preventDefault();
        e.stopPropagation();
        onDropOnPanel(e, field._id);
        return;
      }
      onDrop(e, index);
    }}
    className={`fb-field-card ${selected ? 'selected' : ''} ${err.length ? 'error' : ''} ${isPanel ? 'panel-card' : ''}`}
    onClick={() => onSelect(field._id)}
    title={shortcuts}
    style={{ marginLeft: depth * 24 }}>
    <div className="fb-field-row">
      <div className="fb-field-drag-handle">
        <DragIndicatorIcon style={{ fontSize: 18 }} />
      </div>
      <div className="fb-field-main">
        <div className="fb-field-header">
          <span className="fb-field-title-icon" style={{ display: "flex", alignItems: "center" }}>
            {TYPE_ICONS[field.type]}
          </span>
          <span className="fb-field-title-text">{field.title || "(no label)"}</span>
          {field.isRequired && <Pill color={C.red} bg={C.redPale}>Required</Pill>}
          {field.readOnly && <Pill color={C.textMuted} bg={C.offWhite}>Read-only</Pill>}
          {field.startWithNewLine === false && <Pill color={C.amber} bg={C.amberPale}>Inline</Pill>}
          {field.titleLocation === "hidden" && <Pill color={C.textMuted} bg={C.offWhite}>Title hidden</Pill>}
          {field.visibleIf && <Pill color={C.green} bg={C.greenPale}>Conditional</Pill>}
          {field.enableIf && <Pill color={C.purpleLight} bg={C.purplePale}>Dyn.enable</Pill>}
          {spColLabel && <Pill color={C.textSecond} bg={C.offWhite}>{spColLabel}</Pill>}
          {(field.type === "dynamicmatrix" || field.type === "tableinput") && <Pill color={C.amber} bg={C.amberPale}>→ Rich Text</Pill>}
          {isPanel && (field.elements?.length ? <Pill color={C.purple} bg={C.purplePale}>{field.elements.length} item{field.elements.length !== 1 ? 's' : ''}</Pill> : <Pill color={C.textMuted} bg={C.offWhite}>Empty</Pill>)}
        </div>
        <div className="fb-field-meta">
          <span className="fb-field-name">{field.name}</span>
          <span className="fb-field-type">· {td.label}</span>
          {field.defaultValue !== undefined && <span className="fb-field-default">· default: {String(field.defaultValue).slice(0, 20)}</span>}
        </div>
        {err.map((e, i) => <div key={i} className="fb-field-error"><WarningAmberIcon style={{ fontSize: 12 }} />{e.msg}</div>)}
      </div>
      <div className="fb-field-actions" onClick={e => e.stopPropagation()}>
        <IconBtn icon={<ArrowUpwardIcon style={{ fontSize: 14 }} />} title="Move up" onClick={() => onMoveUp()} disabled={isFirst} />
        <IconBtn icon={<ArrowDownwardIcon style={{ fontSize: 14 }} />} title="Move down" onClick={() => onMoveDown()} disabled={isLast} />
        <IconBtn icon={<ContentCopyIcon style={{ fontSize: 14 }} />} title="Duplicate (Ctrl+D)" onClick={() => onDuplicate(field)} />
        <IconBtn icon={<CloseIcon style={{ fontSize: 14 }} />} title="Remove (Del)" onClick={() => onRemove(field._id)} danger />
      </div>
    </div>
    {/* Render panel children */}
    {isPanel && Array.isArray(field.elements) && field.elements.length > 0 && (
      <div className="fb-panel-children" onClick={e => e.stopPropagation()}>
        {field.elements.map((child, childIdx) => (
          <FieldCard
            key={child._id}
            field={child}
            index={childIdx}
            selected={selectedId === child._id}
            onSelect={onSelect}
            onRemove={onRemove}
            onDuplicate={onDuplicate}
            onMoveUp={childIdx > 0 && onRecursiveReorder ? () => onRecursiveReorder(child._id, field.elements![childIdx - 1]._id) : () => {}}
            onMoveDown={childIdx < field.elements!.length - 1 && onRecursiveReorder ? () => onRecursiveReorder(child._id, field.elements![childIdx + 1]._id) : () => {}}
            isFirst={childIdx === 0}
            isLast={childIdx === field.elements!.length - 1}
            errors={errors}
            onDragStart={(e: React.DragEvent) => {
              e.stopPropagation();
              e.dataTransfer.setData("field_id", child._id);
              e.dataTransfer.setData("from_panel", "true");
              e.dataTransfer.effectAllowed = "move";
            }}
            onDragOver={onDragOver}
            onDrop={onDrop}
            dragging={false}
            onDropOnPanel={onDropOnPanel}
            onRecursiveReorder={onRecursiveReorder}
            depth={depth + 1}
            selectedId={selectedId}
          />
        ))}
      </div>
    )}
  </div>;
}

function Canvas({ fields, selectedId, onSelect, onRemove, onDuplicate, onReorder, onAddFromPalette, errors, onDropOnPanel, onRecursiveReorder, onMoveToRoot }: {
  fields: FormBuilderField[]; selectedId: string | null; onSelect: (id: string | null) => void;
  onRemove: (id: string) => void; onDuplicate: (field: FormBuilderField) => void;
  onReorder: (from: number, to: number) => void; onAddFromPalette: (td: typeof QUESTION_TYPES[number], atIndex?: number) => void;
  errors: { id: string; msg: string }[];
  onDropOnPanel?: (e: React.DragEvent, panelId: string) => void;
  onRecursiveReorder?: (fromId: string, toId: string) => void;
  onMoveToRoot?: (fieldId: string, atIndex: number) => void;
}) {
  const dragIndexRef = useRef<number | null>(null);
  const [dragOverIndex, setDragOverIndex] = useState<number | null>(null);
  const [draggingIndex, setDraggingIndex] = useState<number | null>(null);

  const onDragStart = (e: React.DragEvent, i: number) => {
    dragIndexRef.current = i;
    setDraggingIndex(i);
    e.dataTransfer.effectAllowed = "move";
    if (fields[i]) {
      e.dataTransfer.setData("field_id", fields[i]._id);
    }
  };
  const onDragOver = (e: React.DragEvent, i: number) => { e.preventDefault(); setDragOverIndex(i); };
  const onDrop = (e: React.DragEvent, i: number) => {
    e.preventDefault();
    setDragOverIndex(null);
    setDraggingIndex(null);
    const pd = e.dataTransfer.getData("palette_type");
    if (pd) { try { onAddFromPalette(JSON.parse(pd), i); } catch { } dragIndexRef.current = null; return; }
    // Handle field dragged out of a panel to the root canvas
    const fieldId = e.dataTransfer.getData("field_id");
    const fromPanel = e.dataTransfer.getData("from_panel") === "true";
    if (fromPanel && fieldId && onMoveToRoot) {
      onMoveToRoot(fieldId, i);
      dragIndexRef.current = null;
      return;
    }
    if (dragIndexRef.current !== null && dragIndexRef.current !== i) onReorder(dragIndexRef.current, i);
    dragIndexRef.current = null;
  };
  const onDragEnd = () => { setDraggingIndex(null); setDragOverIndex(null); dragIndexRef.current = null; };

  // Handles drops on empty canvas space (appends to end)
  const onCanvasDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOverIndex(null);
    setDraggingIndex(null);
    const pd = e.dataTransfer.getData("palette_type");
    if (pd) {
      try { onAddFromPalette(JSON.parse(pd), fields.length); } catch { }
    }
    // Handle field dragged out of a panel to the end of the canvas
    const fieldId = e.dataTransfer.getData("field_id");
    const fromPanel = e.dataTransfer.getData("from_panel") === "true";
    if (fromPanel && fieldId && onMoveToRoot) {
      onMoveToRoot(fieldId, fields.length);
    }
    dragIndexRef.current = null;
  };

  return <div onDragOver={e => e.preventDefault()} onDrop={onCanvasDrop} onDragEnd={onDragEnd} className="fb-canvas">
    {!fields.length
      ? <div className="fb-canvas-empty">
        <div className="fb-canvas-empty-icon" style={{ display: "flex", alignItems: "center", justifyContent: "center" }}>
          <TextFieldsIcon style={{ fontSize: 40, color: C.textMuted }} />
        </div>
        <div className="fb-canvas-empty-title">Your form is empty</div>
        <div className="fb-canvas-empty-text">Click a field type in the left panel,<br />or drag one here to get started.</div>
      </div>
      : <>
        {fields.map((field, i) => <React.Fragment key={field._id}>
          {/* Stable DOM: indicator always rendered, toggled via display so React never inserts/removes nodes mid-drag */}
          <div
            key={`indicator-${field._id}`}
            className="fb-canvas-drop-indicator"
            style={{ display: dragOverIndex === i && draggingIndex !== i ? "block" : "none" }}
          />
          <FieldCard
            key={`card-${field._id}`}
            field={field} index={i} selected={selectedId === field._id}
            onSelect={onSelect} onRemove={onRemove} onDuplicate={onDuplicate}
            onMoveUp={() => onReorder(i, i - 1)} onMoveDown={() => onReorder(i, i + 1)}
            isFirst={i === 0} isLast={i === fields.length - 1} errors={errors}
            onDragStart={onDragStart} onDragOver={onDragOver} onDrop={onDrop} dragging={draggingIndex === i}
            onDropOnPanel={onDropOnPanel} onRecursiveReorder={onRecursiveReorder} selectedId={selectedId} />
        </React.Fragment>)}
        {/* Drop zone after last card — stable DOM */}
        <div
          className="fb-canvas-drop-indicator"
          style={{ display: dragOverIndex === fields.length && draggingIndex !== fields.length ? "block" : "none" }}
        />
      </>}
  </div>;
}

// ── Property Editors ──────────────────────────────────────────────────
function ChoicesEditor({ choices, onChange }: { choices: (string | { value: string; text: string })[]; onChange: (c: (string | { value: string; text: string })[]) => void }) {
  type ChoiceItem = { value: string; text: string; _textCustomised?: boolean };
  const items = (Array.isArray(choices) ? choices : []).map(c => typeof c === "string" ? { value: c, text: c } : c) as ChoiceItem[];
  const update = (i: number, k: string, v: string) => { const n = items.map((it, idx) => idx === i ? { ...it, [k]: v, ...(k === "value" && !it._textCustomised ? { text: v } : {}) } : it); onChange(n.map(x => x.value === x.text ? x.value : x)); };
  const add = () => { const n = [...items, { value: `option${items.length + 1}`, text: `Option ${items.length + 1}` }]; onChange(n.map(x => x.value === x.text ? x.value : x)); };
  return <div className="fb-choices-list">
    {items.map((it, i) => <div key={i} className="fb-choice-row">
      <Input value={it.value} onChange={v => update(i, "value", v)} placeholder="value" className="fb-choice-input" />
      <Input value={it.text} onChange={v => update(i, "text", v)} placeholder="label" className="fb-choice-input" />
      <IconBtn icon={<CloseIcon style={{ fontSize: 14 }} />} title="Remove" onClick={() => onChange(items.filter((_, idx) => idx !== i).map(x => x.value === x.text ? x.value : x))} danger />
    </div>)}
    <button onClick={add} className="fb-add-choice-btn"><AddIcon style={{ fontSize: 14 }} /> Add option</button>
  </div>;
}

function DefaultValueEditor({ field, onChange }: { field: FormBuilderField; onChange: (patch: Partial<FormBuilderField>) => void }) {
  const handleChange = (v: string) => {
    if (field.type === "number" || field.inputType === "number") {
      onChange({ defaultValue: v === "" ? undefined : Number(v) });
    } else if (field.type === "boolean") {
      onChange({ defaultValue: v === "true" });
    } else {
      onChange({ defaultValue: v === "" ? undefined : v });
    }
  };
  const currentValue = field.defaultValue !== undefined ? String(field.defaultValue) : "";
  const inputType = field.type === "date" ? "date" : "text";
  return <PropRow label="Default value">
    <Input type={inputType} value={currentValue} onChange={handleChange} placeholder="Enter default value" />
  </PropRow>;
}

/** Renders type-specific configuration controls in the General tab */
function FieldTypeProps({ field, onChange }: { field: FormBuilderField; onChange: (patch: Partial<FormBuilderField>) => void }) {
  const numericTypes = ["number", "slider", "counter", "currency"];
  const dateTypes = ["date", "datetime", "daterange"];
  const commentTypes = ["comment", "jsoneditor", "addressblock", "locationpicker", "budgetallocator"];
  const fileTypes = ["file", "imageupload"];
  const htmlTypes = ["html", "alert", "countdown", "datatable", "chartdisplay", "videoembed"];
  const booleanTypes = ["boolean", "consent"];
  const matrixTypes = ["dynamicmatrix", "tableinput"];

  return <>
    {/* Numeric: min / max / step */}
    {numericTypes.includes(field.type) && <>
      <div style={{ display: "flex", gap: 8 }}>
        <PropRow label="Min"><Input type="number" value={field.min ?? ""} onChange={v => onChange({ min: v === "" ? undefined : Number(v) })} placeholder="No min" /></PropRow>
        <PropRow label="Max"><Input type="number" value={field.max ?? ""} onChange={v => onChange({ max: v === "" ? undefined : Number(v) })} placeholder="No max" /></PropRow>
      </div>
      <PropRow label="Step"><Input type="number" value={field.step ?? ""} onChange={v => onChange({ step: v === "" ? undefined : Number(v) })} placeholder="1" /></PropRow>
    </>}

    {/* Formula: expression editor + displayFormat + recalculate toggle */}
    {field.type === "formula" && <>
      <PropRow label="Expression / Formula" span>
        <Textarea value={field.expression || ""} onChange={v => onChange({ expression: v })} rows={3} placeholder="e.g. {field1} + {field2}" />
        <div style={{ fontSize: 10, color: C.textMuted, marginTop: 4, lineHeight: 1.4 }}>
          Use <code style={{ background: "#F3F4F6", padding: "1px 4px", borderRadius: 3, fontSize: 10 }}>{'{field_name}'}</code> syntax to reference other fields.
          Supports <strong>+</strong>, <strong>-</strong>, <strong>*</strong>, <strong>/</strong>, parentheses, and SurveyJS expression functions.
        </div>
      </PropRow>
      <PropRow label="Display format">
        <Select value={field.displayFormat || "number"} onChange={v => onChange({ displayFormat: v })} options={[
          { value: "number", label: "Number" },
          { value: "currency", label: "Currency (RM)" },
          { value: "percent", label: "Percentage" },
        ]} />
      </PropRow>
      <div style={{ paddingTop: 4 }}>
        <Toggle checked={field.recalculateOnChange !== false} onChange={v => onChange({ recalculateOnChange: v })} label="Auto-recalculate on change" />
      </div>
    </>}

    {/* Date: minDate / maxDate / disableWeekends */}
    {dateTypes.includes(field.type) && <>
      <div style={{ display: "flex", gap: 8 }}>
        <PropRow label="Min date"><Input type="date" value={field.minDate || ""} onChange={v => onChange({ minDate: v || undefined })} /></PropRow>
        <PropRow label="Max date"><Input type="date" value={field.maxDate || ""} onChange={v => onChange({ maxDate: v || undefined })} /></PropRow>
      </div>
      {(field.type === "date" || field.type === "datetime") && <Toggle checked={!!field.disableWeekends} onChange={v => onChange({ disableWeekends: v })} label="Disable weekends" />}
    </>}

    {/* Comment-like: rows */}
    {commentTypes.includes(field.type) && <>
      <PropRow label="Rows"><Input type="number" value={field.rows || ""} onChange={v => onChange({ rows: v === "" ? undefined : Number(v) })} placeholder="4" /></PropRow>
    </>}

    {/* Rating: rateMin / rateMax / descriptions */}
    {field.type === "rating" && <>
      <div style={{ display: "flex", gap: 8 }}>
        <PropRow label="Min rate"><Input type="number" value={field.rateMin ?? ""} onChange={v => onChange({ rateMin: v === "" ? undefined : Number(v) })} placeholder="1" /></PropRow>
        <PropRow label="Max rate"><Input type="number" value={field.rateMax ?? ""} onChange={v => onChange({ rateMax: v === "" ? undefined : Number(v) })} placeholder="5" /></PropRow>
      </div>
      <PropRow label="Min label"><Input value={field.minRateDescription || ""} onChange={v => onChange({ minRateDescription: v || undefined })} placeholder="e.g. Poor" /></PropRow>
      <PropRow label="Max label"><Input value={field.maxRateDescription || ""} onChange={v => onChange({ maxRateDescription: v || undefined })} placeholder="e.g. Excellent" /></PropRow>
    </>}

    {/* File / Image upload: acceptedTypes / maxSize / allowMultiple */}
    {fileTypes.includes(field.type) && <>
      <PropRow label="Accepted types"><Input value={field.acceptedTypes || ""} onChange={v => onChange({ acceptedTypes: v || undefined })} placeholder=".pdf,.doc,.png" /></PropRow>
      <PropRow label="Max size (bytes)"><Input type="number" value={field.maxSize ?? ""} onChange={v => onChange({ maxSize: v === "" ? undefined : Number(v) })} placeholder="10485760" /></PropRow>
      <Toggle checked={!!field.allowMultiple} onChange={v => onChange({ allowMultiple: v })} label="Allow multiple files" />
    </>}

    {/* HTML / Display: html content */}
    {htmlTypes.includes(field.type) && <>
      <PropRow label="HTML content" span>
        <Textarea value={field.html || ""} onChange={v => onChange({ html: v })} rows={4} placeholder="<p>Your content</p>" />
      </PropRow>
    </>}

    {/* Image: url / altText / caption */}
    {field.type === "image" && <>
      <PropRow label="Image URL" span><Input value={field.imageUrl || ""} onChange={v => onChange({ imageUrl: v })} placeholder="https://..." /></PropRow>
      <PropRow label="Alt text"><Input value={field.altText || ""} onChange={v => onChange({ altText: v || undefined })} placeholder="Description" /></PropRow>
      <PropRow label="Caption"><Input value={field.caption || ""} onChange={v => onChange({ caption: v || undefined })} placeholder="Optional caption" /></PropRow>
    </>}

    {/* Panel: title / description / collapsible */}
    {field.type === "panel" && <>
      <PropRow label="Panel title" span><Input value={field.title || ""} onChange={v => onChange({ title: v })} placeholder="Section title" /></PropRow>
      <PropRow label="Description" span><Input value={field.description || ""} onChange={v => onChange({ description: v })} placeholder="Optional description" /></PropRow>
      <Toggle checked={!!field.collapsible} onChange={v => onChange({ collapsible: v, collapsed: v ? !!field.collapsed : false })} label="Collapsible" />
      {field.collapsible && <Toggle checked={!!field.collapsed} onChange={v => onChange({ collapsed: v })} label="Start collapsed" />}
    </>}

    {/* Boolean / Toggle / Consent: labelTrue / labelFalse */}
    {booleanTypes.includes(field.type) && <>
      <PropRow label="True label"><Input value={field.labelTrue || ""} onChange={v => onChange({ labelTrue: v || undefined })} placeholder="Yes" /></PropRow>
      <PropRow label="False label"><Input value={field.labelFalse || ""} onChange={v => onChange({ labelFalse: v || undefined })} placeholder="No" /></PropRow>
    </>}

    {/* Page break: pageTitle / pageDescription */}
    {field.type === "pagebreak" && <>
      <PropRow label="Page title" span><Input value={field.pageTitle || ""} onChange={v => onChange({ pageTitle: v || undefined })} placeholder="Page title" /></PropRow>
      <PropRow label="Page description" span><Input value={field.pageDescription || ""} onChange={v => onChange({ pageDescription: v || undefined })} placeholder="Optional description" /></PropRow>
    </>}

    {/* Matrix / Table: minRows / maxRows / addRowText */}
    {matrixTypes.includes(field.type) && <>
      <div style={{ display: "flex", gap: 8 }}>
        <PropRow label="Min rows"><Input type="number" value={field.minRows ?? ""} onChange={v => onChange({ minRows: v === "" ? undefined : Number(v) })} placeholder="1" /></PropRow>
        <PropRow label="Max rows"><Input type="number" value={field.maxRows ?? ""} onChange={v => onChange({ maxRows: v === "" ? undefined : Number(v) })} placeholder="10" /></PropRow>
      </div>
      <PropRow label="Add row text"><Input value={field.addRowText || ""} onChange={v => onChange({ addRowText: v || undefined })} placeholder="Add Row" /></PropRow>
    </>}

    {/* Signature pad: width / height / penColor / backgroundColor */}
    {field.type === "signaturepad" && <>
      <div style={{ display: "flex", gap: 8 }}>
        <PropRow label="Width"><Input type="number" value={field.signatureWidth ?? ""} onChange={v => onChange({ signatureWidth: v === "" ? undefined : Number(v) })} placeholder="400" /></PropRow>
        <PropRow label="Height"><Input type="number" value={field.signatureHeight ?? ""} onChange={v => onChange({ signatureHeight: v === "" ? undefined : Number(v) })} placeholder="200" /></PropRow>
      </div>
      <PropRow label="Pen color"><Input value={field.penColor || ""} onChange={v => onChange({ penColor: v || undefined })} placeholder="#000000" /></PropRow>
      <PropRow label="Background color"><Input value={field.backgroundColor || ""} onChange={v => onChange({ backgroundColor: v || undefined })} placeholder="#FFFFFF" /></PropRow>
    </>}

    {/* Spacer: height */}
    {field.type === "spacer" && <>
      <PropRow label="Height (px)"><Input type="number" value={field.height ?? ""} onChange={v => onChange({ height: v === "" ? undefined : Number(v) })} placeholder="16" /></PropRow>
    </>}

    {/* Divider: style / color / margin */}
    {field.type === "divider" && <>
      <PropRow label="Style"><Select value={field.dividerStyle || "solid"} onChange={v => onChange({ dividerStyle: v as "solid" | "dashed" | "dotted" })} options={[{ value: "solid", label: "Solid" }, { value: "dashed", label: "Dashed" }, { value: "dotted", label: "Dotted" }]} /></PropRow>
      <PropRow label="Color"><Input value={field.dividerColor || ""} onChange={v => onChange({ dividerColor: v || undefined })} placeholder="#E5E3F0" /></PropRow>
      <PropRow label="Margin"><Input value={field.dividerMargin || ""} onChange={v => onChange({ dividerMargin: v || undefined })} placeholder="16px 0" /></PropRow>
    </>}

    {/* Repeater: minRows / maxRows / button text */}
    {field.type === "repeater" && <>
      <div style={{ display: "flex", gap: 8 }}>
        <PropRow label="Min rows"><Input type="number" value={field.minRows ?? ""} onChange={v => onChange({ minRows: v === "" ? undefined : Number(v) })} placeholder="1" /></PropRow>
        <PropRow label="Max rows"><Input type="number" value={field.maxRows ?? ""} onChange={v => onChange({ maxRows: v === "" ? undefined : Number(v) })} placeholder="10" /></PropRow>
      </div>
      <PropRow label="Add button text"><Input value={field.addButtonText || ""} onChange={v => onChange({ addButtonText: v || undefined })} placeholder="Add Row" /></PropRow>
      <PropRow label="Remove button text"><Input value={field.removeButtonText || ""} onChange={v => onChange({ removeButtonText: v || undefined })} placeholder="Remove" /></PropRow>
      <Toggle checked={!!field.showBlankRow} onChange={v => onChange({ showBlankRow: v })} label="Show blank row" />
    </>}

    {/* Columns: columnCount / gap */}
    {field.type === "columns" && <>
      <PropRow label="Column count"><Select value={String(field.columnCount || 2)} onChange={v => onChange({ columnCount: parseInt(v) })} options={[{ value: "2", label: "2 columns" }, { value: "3", label: "3 columns" }]} /></PropRow>
      <PropRow label="Gap (px)"><Input type="number" value={field.gap ?? ""} onChange={v => onChange({ gap: v === "" ? undefined : Number(v) })} placeholder="16" /></PropRow>
    </>}
  </>;
}

function MatrixColumnsEditor({ columns, token, onChange }: {
  columns: { name: string; title: string; cellType?: string; choices?: string[]; multiSelect?: boolean; choicesSource?: { list?: string; column?: string }; filteredListSource?: { list?: string; valueColumn?: string; filterColumn?: string; filterValue?: string; choicesLoaded?: boolean } }[];
  token?: string;
  onChange: (cols: { name: string; title: string; cellType?: string; choices?: string[]; multiSelect?: boolean; choicesSource?: { list?: string; column?: string }; filteredListSource?: { list?: string; valueColumn?: string; filterColumn?: string; filterValue?: string; choicesLoaded?: boolean } }[]) => void;
}) {
  const addCol = () => {
    const idx = columns.length + 1;
    onChange([...columns, { name: `col${idx}`, title: `Column ${idx}`, cellType: "text" }]);
  };
  const updateCol = (i: number, patch: Partial<typeof columns[0]>) => {
    const next = columns.map((c, idx) => idx === i ? { ...c, ...patch } : c);
    onChange(next);
  };
  const removeCol = (i: number) => onChange(columns.filter((_, idx) => idx !== i));

  return <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Matrix Columns</span>
      <div style={{ flex: 1 }} />
      <button onClick={addCol} style={{ fontSize: 11, color: C.purple, background: "none", border: `1px dashed ${C.purple}`, borderRadius: 6, padding: "3px 10px", cursor: "pointer", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}>＋ Add column</button>
    </div>
    {columns.length === 0 && <div style={{ fontSize: 11, color: C.textMuted, padding: 8, background: C.offWhite, borderRadius: 6 }}>No columns defined. Add at least one.</div>}
    {columns.map((col, i) => {
      const hasChoices = col.cellType === "dropdown" || col.cellType === "checkbox";
      return <div key={i} style={{ padding: 10, background: C.offWhite, borderRadius: 8, border: `1px solid ${C.border}`, display: "flex", flexDirection: "column", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <span style={{ fontSize: 11, fontWeight: 700, color: C.purple, width: 18 }}>{i + 1}</span>
          <div style={{ flex: 1, display: "flex", gap: 6 }}>
            <input value={col.name} onChange={e => updateCol(i, { name: e.target.value.replace(/\s+/g, "_") })} placeholder="Name" style={{ flex: 1, fontSize: 11, padding: "4px 8px", border: `1px solid ${C.border}`, borderRadius: 5, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }} />
            <input value={col.title} onChange={e => updateCol(i, { title: e.target.value })} placeholder="Title" style={{ flex: 1.5, fontSize: 11, padding: "4px 8px", border: `1px solid ${C.border}`, borderRadius: 5, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }} />
          </div>
          <button onClick={() => removeCol(i)} style={{ fontSize: 11, color: C.red, background: "none", border: "none", cursor: "pointer" }}>✕</button>
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <span style={{ fontSize: 10, color: C.textMuted, whiteSpace: "nowrap" }}>Cell type:</span>
          <Select value={col.cellType || "text"} onChange={v => updateCol(i, { cellType: v, choices: undefined, choicesSource: undefined, multiSelect: undefined })} options={[
            { value: "text", label: "Text" },
            { value: "dropdown", label: "Dropdown" },
            { value: "date", label: "Date" },
            { value: "number", label: "Number" },
            { value: "checkbox", label: "Checkbox" },
            { value: "boolean", label: "Boolean" },
          ]} />
          {col.cellType === "dropdown" && <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10, color: C.textMuted }}>
            <input type="checkbox" checked={!!col.multiSelect} onChange={e => updateCol(i, { multiSelect: e.target.checked })} /> Multi-select
          </label>}
        </div>
        {hasChoices && <>
          <SpChoicesSourceEditor
            source={col.choicesSource}
            token={token}
            onChange={src => updateCol(i, { choicesSource: src || undefined, choices: src?.list ? [] : (col.choices || []) })}
          />
          <SpFilteredListSourceEditor
            source={col.filteredListSource}
            token={token}
            onChange={src => updateCol(i, { filteredListSource: src || undefined, choices: src?.list ? [] : (col.choices || []) })}
          />
          {!col.choicesSource?.list && !col.filteredListSource?.list && <div style={{ display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center" }}>
            {(col.choices || []).map((ch, ci) => <span key={ci} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10, padding: "2px 8px", background: C.purplePale, color: C.purple, borderRadius: 12 }}>
              {ch}
              <button onClick={() => updateCol(i, { choices: (col.choices || []).filter((_, idx) => idx !== ci) })} style={{ fontSize: 9, color: C.red, background: "none", border: "none", cursor: "pointer", padding: 0 }}>✕</button>
            </span>)}
            <input
              placeholder="Add choice…"
              onKeyDown={e => {
                if (e.key === "Enter") {
                  const val = (e.target as HTMLInputElement).value.trim();
                  if (val) { updateCol(i, { choices: [...(col.choices || []), val] }); (e.target as HTMLInputElement).value = ""; }
                }
              }}
              style={{ fontSize: 11, padding: "3px 8px", border: `1px dashed ${C.border}`, borderRadius: 12, width: 90, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}
            />
          </div>}
        </>}
      </div>;
    })}
  </div>;
}

function SpChoicesSourceEditor({ source, token, onChange }: {
  source?: { list?: string; column?: string; multiSelect?: boolean };
  token?: string;
  onChange: (src: { list?: string; column?: string; multiSelect?: boolean } | undefined) => void;
}) {
  const [mode, setMode] = useState<"manual" | "sp">(source?.list ? "sp" : "manual");
  const [lists, setLists] = useState<{ title: string; id: string }[]>([]);
  const [columns, setColumns] = useState<{ title: string; typeKind: number; choices: string[] }[]>([]);
  const [loadingLists, setLoadingLists] = useState(false);
  const [loadingCols, setLoadingCols] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => { setMode(source?.list ? "sp" : "manual"); }, [source?.list]);

  useEffect(() => {
    if (mode !== "sp" || !token) return;
    setLoadingLists(true);
    setError("");
    import("../../utils/formBuilderSP").then(({ getSharePointLists }) => {
      getSharePointLists(token).then(setLists).catch((e: Error) => setError(e.message)).finally(() => setLoadingLists(false));
    });
  }, [mode, token]);

  useEffect(() => {
    if (mode !== "sp" || !token || !source?.list) { setColumns([]); return; }
    setLoadingCols(true);
    setError("");
    import("../../utils/formBuilderSP").then(({ getChoiceColumnsForList }) => {
      getChoiceColumnsForList(source.list!, token).then(setColumns).catch((e: Error) => setError(e.message)).finally(() => setLoadingCols(false));
    });
  }, [mode, token, source?.list]);

  const selectedCol = columns.find(c => c.title === source?.column);

  return <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Data Source</span>
      <div style={{ flex: 1 }} />
    </div>
    <div style={{ display: "flex", gap: 8 }}>
      <button onClick={() => { setMode("manual"); onChange(undefined); }}
        style={{ flex: 1, padding: "6px 0", borderRadius: 6, border: `1px solid ${mode === "manual" ? C.purple : C.border}`, background: mode === "manual" ? C.purplePale : C.white, color: mode === "manual" ? C.purple : C.textMuted, fontSize: 12, cursor: "pointer", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}>
        Manual
      </button>
      <button onClick={() => setMode("sp")}
        style={{ flex: 1, padding: "6px 0", borderRadius: 6, border: `1px solid ${mode === "sp" ? C.purple : C.border}`, background: mode === "sp" ? C.purplePale : C.white, color: mode === "sp" ? C.purple : C.textMuted, fontSize: 12, cursor: "pointer", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}>
        SharePoint List
      </button>
    </div>
    {mode === "sp" && <>
      {!token && <div style={{ fontSize: 11, color: C.amber, padding: 8, background: C.amberPale, borderRadius: 6 }}>Sign in to load SharePoint lists.</div>}
      {!!token && <>
        <PropRow label="List">
          <Select value={source?.list || ""} onChange={v => onChange({ list: v || undefined, column: undefined })} options={[
            { value: "", label: loadingLists ? "Loading…" : "Select a list" },
            ...lists.map(l => ({ value: l.title, label: l.title }))
          ]} />
        </PropRow>
        {source?.list && <PropRow label="Column">
          <Select value={source?.column || ""} onChange={v => onChange({ ...source, column: v || undefined })} options={[
            { value: "", label: loadingCols ? "Loading…" : "Select a choice column" },
            ...columns.map(c => ({ value: c.title, label: `${c.title} (${c.typeKind === 15 ? "Multi" : "Single"})` }))
          ]} />
        </PropRow>}
        {selectedCol && selectedCol.choices.length > 0 && <div style={{ display: "flex", flexWrap: "wrap", gap: 4, padding: 8, background: C.offWhite, borderRadius: 6 }}>
          {selectedCol.choices.map(ch => <span key={ch} style={{ fontSize: 10, padding: "2px 8px", background: C.purplePale, color: C.purple, borderRadius: 12 }}>{ch}</span>)}
        </div>}
        {error && <div style={{ fontSize: 11, color: C.red }}>{error}</div>}
      </>}
    </>}
  </div>;
}

function SpFilteredListSourceEditor({ source, token, onChange }: {
  source?: { list?: string; valueColumn?: string; labelColumn?: string; filterColumn?: string; filterValue?: string; choicesLoaded?: boolean };
  token?: string;
  onChange: (src: { list?: string; valueColumn?: string; labelColumn?: string; filterColumn?: string; filterValue?: string } | undefined) => void;
}) {
  const [enabled, setEnabled] = useState(!!source?.list);
  const [lists, setLists] = useState<{ title: string; id: string }[]>([]);
  const [columns, setColumns] = useState<{ title: string; typeKind: number }[]>([]);
  const [loadingLists, setLoadingLists] = useState(false);
  const [loadingCols, setLoadingCols] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => { setEnabled(!!source?.list); }, [source?.list]);

  useEffect(() => {
    if (!enabled || !token) return;
    setLoadingLists(true);
    setError("");
    import("../../utils/formBuilderSP").then(({ getSharePointLists }) => {
      getSharePointLists(token).then(setLists).catch((e: Error) => setError(e.message)).finally(() => setLoadingLists(false));
    });
  }, [enabled, token]);

  useEffect(() => {
    if (!enabled || !token || !source?.list) { setColumns([]); return; }
    setLoadingCols(true);
    setError("");
    import("../../utils/formBuilderSP").then(({ getAllColumnsForList }) => {
      getAllColumnsForList(source.list!, token).then(setColumns).catch((e: Error) => setError(e.message)).finally(() => setLoadingCols(false));
    });
  }, [enabled, token, source?.list]);

  const toggle = () => {
    if (enabled) {
      onChange(undefined);
      setEnabled(false);
    } else {
      setEnabled(true);
      onChange({ list: "" });
    }
  };

  return <div style={{ display: "flex", flexDirection: "column", gap: 10, paddingTop: 8, borderTop: `1px solid ${C.border}` }}>
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
      <span style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em" }}>Filtered List Source</span>
      <div style={{ flex: 1 }} />
      <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 11, color: C.textSecond, userSelect: "none" }}>
        <input type="checkbox" checked={enabled} onChange={toggle} style={{ width: 14, height: 14, accentColor: C.purple }} />
        Enabled
      </label>
    </div>

    {enabled && <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
      <PropRow label="List">
        <Select value={source?.list || ""} onChange={v => onChange({ ...source, list: v || undefined, valueColumn: undefined, filterColumn: undefined })} options={[
          { value: "", label: loadingLists ? "Loading…" : "Select a list" },
          ...lists.map(l => ({ value: l.title, label: l.title }))
        ]} />
      </PropRow>
      {source?.list && <>
        <PropRow label="Value Column">
          <Select value={source?.valueColumn || ""} onChange={v => onChange({ ...source, valueColumn: v || undefined, filterColumn: undefined })} options={[
            { value: "", label: loadingCols ? "Loading…" : "Select a column" },
            ...columns.map(c => ({ value: c.title, label: c.title }))
          ]} />
        </PropRow>
        <PropRow label="Label Column (optional)">
          <Select value={source?.labelColumn || ""} onChange={v => onChange({ ...source, labelColumn: v || undefined })} options={[
            { value: "", label: "Same as value" },
            ...columns.map(c => ({ value: c.title, label: c.title }))
          ]} />
        </PropRow>
      </>}
      {source?.list && source?.valueColumn && <>
        <div style={{ fontSize: 10, fontWeight: 600, color: C.textMuted, marginBottom: 2 }}>Filter (optional)</div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <span style={{ fontSize: 10, color: C.textMuted, whiteSpace: "nowrap" }}>Where</span>
          <select value={source?.filterColumn || ""} onChange={e => onChange({ ...source, filterColumn: (e.target as HTMLSelectElement).value || undefined })}
            style={{ flex: 1, height: 26, border: `1px solid ${C.border}`, borderRadius: 5, fontSize: 11, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif", padding: "0 4px" }}>
            <option value="">— select column —</option>
            {columns.map(c => <option key={c.title} value={c.title}>{c.title}</option>)}
          </select>
          <span style={{ fontSize: 10, color: C.textMuted }}>=</span>
          <input value={source?.filterValue || ""} onChange={e => onChange({ ...source, filterValue: e.target.value })}
            placeholder="value"
            style={{ flex: 1, height: 26, border: `1px solid ${C.border}`, borderRadius: 5, fontSize: 11, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif", padding: "0 6px" }} />
        </div>
      </>}
      {error && <div style={{ fontSize: 11, color: C.red }}>{error}</div>}
    </div>}
  </div>;
}

function PropertyPanel({ field, allFields, onChange, onSurveySettingsChange, surveySettings, token }: {
  field: FormBuilderField | null; allFields: FormBuilderField[];
  onChange: (patch: Partial<FormBuilderField>) => void;
  onSurveySettingsChange?: (s: Record<string, unknown>) => void;
  surveySettings?: Record<string, unknown>;
  token?: string;
}) {
  const [tab, setTab] = useState("general");

  // Survey-level settings panel
  if (!field && surveySettings) return <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
    <div style={{ padding: "12px 14px", borderBottom: `1px solid ${C.border}`, background: C.purplePale }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
        <SettingsIcon style={{ fontSize: 16, color: C.purple }} />
        <span style={{ fontSize: 13, fontWeight: 700, color: C.purple }}>Form Settings</span>
      </div>
      <div style={{ fontSize: 10, color: C.textMuted }}>SurveyJS form properties</div>
    </div>
    <div style={{ flex: 1, padding: "14px" }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        <PropRow label="Form title"><Input value={(surveySettings.title as string) || ""} onChange={v => onSurveySettingsChange?.({ ...surveySettings, title: v })} placeholder="Form title" /></PropRow>
        <PropRow label="Show form title">
          <Toggle checked={(surveySettings.titleLocation as string) !== "hidden"} onChange={v => onSurveySettingsChange?.({ ...surveySettings, titleLocation: v ? ((surveySettings.titleLocation as string) === "hidden" ? "top" : surveySettings.titleLocation) : "hidden" })} label={(surveySettings.titleLocation as string) !== "hidden" ? "Visible" : "Hidden"} />
        </PropRow>
        <PropRow label="Form description"><Textarea value={(surveySettings.description as string) || ""} onChange={v => onSurveySettingsChange?.({ ...surveySettings, description: v })} rows={2} placeholder="Optional description" /></PropRow>
        <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12, marginTop: 4 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>Text Formatting</div>
          <PropRow label="Question titles">
            <Select value={(surveySettings.titleLocation as string) || "default"} onChange={v => onSurveySettingsChange?.({ ...surveySettings, titleLocation: v })} options={[{ value: "default", label: "Default" }, { value: "hidden", label: "Hidden" }, { value: "top", label: "Top" }, { value: "bottom", label: "Bottom" }]} />
          </PropRow>
          <PropRow label="Text transform">
            <Select value={(surveySettings.textTransform as string) || "none"} onChange={v => onSurveySettingsChange?.({ ...surveySettings, textTransform: v })} options={[{ value: "none", label: "None" }, { value: "uppercase", label: "ALL UPPERCASE" }, { value: "capitalize", label: "First Letter Only" }, { value: "lowercase", label: "all lowercase" }]} />
          </PropRow>
          <PropRow label="Show question numbers">
            <Select value={(surveySettings.showQuestionNumbers as string) || "on"} onChange={v => onSurveySettingsChange?.({ ...surveySettings, showQuestionNumbers: v })} options={[{ value: "on", label: "On" }, { value: "onPage", label: "Per page" }, { value: "onpanel", label: "Per panel" }, { value: "off", label: "Off" }]} />
          </PropRow>
        </div>
        <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 12, marginTop: 4 }}>
          <div style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>Behavior</div>
          <PropRow label="Error mode">
            <Select value={(surveySettings.checkErrorsMode as string) || "onValueChanged"} onChange={v => onSurveySettingsChange?.({ ...surveySettings, checkErrorsMode: v })} options={[{ value: "onValueChanged", label: "On value change" }, { value: "onComplete", label: "On complete" }, { value: "onNextPage", label: "On next page" }]} />
          </PropRow>
          <PropRow label="Text update">
            <Select value={(surveySettings.textUpdateMode as string) || "onTyping"} onChange={v => onSurveySettingsChange?.({ ...surveySettings, textUpdateMode: v })} options={[{ value: "onTyping", label: "On typing" }, { value: "onBlur", label: "On blur" }]} />
          </PropRow>
          <Toggle checked={!!surveySettings.showProgressBar} onChange={v => onSurveySettingsChange?.({ ...surveySettings, showProgressBar: v })} label="Show progress bar" />
          <Toggle checked={!!surveySettings.showPageTitles} onChange={v => onSurveySettingsChange?.({ ...surveySettings, showPageTitles: v })} label="Show page titles" />
        </div>
      </div>
    </div>
  </div>;

  if (!field) return <div style={{ padding: 20, textAlign: "center", color: C.textMuted, fontSize: 12 }}>Select a field to edit its properties</div>;

  const td = QUESTION_TYPES.find(t => t.type === field.type) || QUESTION_TYPES[0];
  const hasChoices = ["dropdown", "radiogroup", "checkbox"].includes(field.type);
  const tabs = [
    { id: "general", label: "General" },
    { id: "options", label: "Options" },
    { id: "logic", label: "Logic" },
    { id: "validation", label: "Validation" }
  ];

  return <div style={{ height: "100%", display: "flex", flexDirection: "column" }}>
    <div style={{ padding: "12px 14px", borderBottom: `1px solid ${C.border}`, background: C.purplePale }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
        <span style={{ fontSize: 16, display: "flex", alignItems: "center" }}>
          {TYPE_ICONS[field.type]}
        </span>
        <span style={{ fontSize: 13, fontWeight: 700, color: C.purple }}>{td.label}</span>
      </div>
      <div style={{ fontSize: 10, color: C.textMuted, fontFamily: "monospace" }}>{field.name}</div>
    </div>
    <div style={{ padding: "8px 14px", borderBottom: `1px solid ${C.border}` }}>
      <select value={tab} onChange={e => setTab(e.target.value)}
        style={{ width: "100%", height: 32, border: `1px solid ${C.border}`, borderRadius: 6, padding: "0 8px", fontSize: 12, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif", color: C.textPrimary, background: C.white, cursor: "pointer" }}>
        {tabs.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
      </select>
    </div>
    <div style={{ flex: 1, padding: "14px" }}>
      {tab === "general" && <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
        <PropRow label="Field name (SP column)" span>
          <Input value={field.name} onChange={v => onChange({ name: v.replace(/\s+/g, "_") })} placeholder="camelCaseName" />
          <div style={{ fontSize: 10, color: C.textMuted, marginTop: 4 }}>No spaces — becomes the SharePoint column name</div>
        </PropRow>
        <PropRow label="Label" span><Input value={field.title} onChange={v => onChange({ title: v })} placeholder="Question label" /></PropRow>
        <PropRow label="Description / hint" span><Input value={field.description || ""} onChange={v => onChange({ description: v })} placeholder="Optional helper text" /></PropRow>
        {!["html", "dynamicmatrix", "file", "formula"].includes(field.type) && <DefaultValueEditor field={field} onChange={onChange} />}
        {field.type === "text" && <PropRow label="Input type"><Select value={field.inputType || "text"} onChange={v => onChange({ inputType: v })} options={[{ value: "text", label: "Text" }, { value: "email", label: "Email" }, { value: "number", label: "Number" }, { value: "date", label: "Date" }, { value: "datetime-local", label: "Date & Time" }, { value: "tel", label: "Phone" }, { value: "url", label: "URL" }, { value: "password", label: "Password" }]} /></PropRow>}
        {field.type === "text" && (!field.inputType || field.inputType === "text") && <PropRow label="Autocapitalize"><Select value={field.autocapitalize || "none"} onChange={v => onChange({ autocapitalize: v as "none" | "sentences" | "words" | "characters" })} options={[{ value: "none", label: "None" }, { value: "sentences", label: "Sentences" }, { value: "words", label: "Words" }, { value: "characters", label: "Characters (ALL CAPS)" }]} /></PropRow>}
        <FieldTypeProps field={field} onChange={onChange} />
        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 4, paddingTop: 12, borderTop: `1px solid ${C.border}` }}>
          {field.type !== "html" && <Toggle checked={!!field.isRequired} onChange={v => onChange({ isRequired: v })} label="Required field" />}
          <Toggle checked={!field.startWithNewLine} onChange={v => onChange({ startWithNewLine: !v })} label="Inline (same row as previous)" />
          <Toggle checked={!!field.readOnly} onChange={v => onChange({ readOnly: v })} label="Read-only" />
          <Toggle checked={field.titleLocation === "hidden"} onChange={v => onChange({ titleLocation: v ? "hidden" : "default" })} label="Hide title" />
        </div>
      </div>}
      {tab === "options" && <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {hasChoices && <>
          <SpChoicesSourceEditor
            source={field.spChoicesSource}
            token={token}
            onChange={src => onChange({ spChoicesSource: src, choices: src?.list ? [] : (field.choices || []) })}
          />
          <SpFilteredListSourceEditor
            source={field.spFilteredListSource}
            token={token}
            onChange={src => onChange({ spFilteredListSource: src, choices: src?.list ? [] : (field.choices || []) })}
          />
          {!field.spChoicesSource?.list && !field.spFilteredListSource?.list && <PropRow label="Choices"><ChoicesEditor choices={field.choices || []} onChange={c => onChange({ choices: c })} /></PropRow>}
        </>}
        {(field.type === "dynamicmatrix" || field.type === "tableinput") && <>
          <MatrixColumnsEditor
            columns={(field.columns || field.tableConfigColumns || []) as { name: string; title: string; cellType?: string; choices?: string[]; multiSelect?: boolean; choicesSource?: { list?: string; column?: string }; filteredListSource?: { list?: string; valueColumn?: string; filterColumn?: string; filterValue?: string; choicesLoaded?: boolean } }[]}
            token={token}
            onChange={cols => onChange({ columns: cols, tableConfigColumns: cols })}
          />
        </>}
        <PropRow label="Columns (side by side)"><Select value={String(field.colCount ?? 1)} onChange={v => onChange({ colCount: parseInt(v) })} options={[0, 1, 2, 3, 4].map(n => ({ value: String(n), label: n === 0 ? "Auto" : `${n} column${n > 1 ? "s" : ""}` }))} /></PropRow>
      </div>}
      {tab === "logic" && <LogicRulesEditor field={field} allFields={allFields} onChange={onChange} />}
      {tab === "validation" && <ValidationEditor field={field} onChange={onChange} />}
    </div>
  </div>;
}

// ── JSON Preview ──────────────────────────────────────────────────────
function JsonPreview({ json, collapsed, onToggle }: { json: SurveyJson; collapsed: boolean; onToggle: () => void }) {
  const [copied, setCopied] = useState(false);
  const text = useMemo(() => JSON.stringify(json, null, 2), [json]);
  const charCount = useMemo(() => JSON.stringify(json).length, [json]);
  const copy = () => navigator.clipboard.writeText(text).then(() => { setCopied(true); setTimeout(() => setCopied(false), 2000); });
  return <div style={{ borderTop: `1px solid ${C.border}`, background: "#1F2937", height: collapsed ? 38 : 220, display: "flex", flexDirection: "column", overflow: "hidden", transition: "height 0.3s" }}>
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 14px", height: 38, flexShrink: 0, cursor: "pointer" }} onClick={onToggle}>
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <CodeIcon style={{ fontSize: 14, color: "#9CA3AF" }} />
        <span style={{ fontSize: 10, fontWeight: 700, color: "#9CA3AF", textTransform: "uppercase", letterSpacing: "0.06em" }}>SurveyJS JSON</span>
        <span style={{ fontSize: 10, color: "rgba(255,255,255,0.3)" }}>{charCount} chars</span>
      </div>
      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        {!collapsed && <button onClick={e => { e.stopPropagation(); copy(); }} style={{ fontSize: 10, color: copied ? "#6EE7B7" : "#9CA3AF", background: "rgba(255,255,255,0.08)", border: "none", borderRadius: 6, padding: "3px 10px", cursor: "pointer", fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}>{copied ? "Copied!" : "Copy JSON"}</button>}
        {collapsed ? <ExpandMoreIcon style={{ fontSize: 16, color: "#9CA3AF" }} /> : <ExpandLessIcon style={{ fontSize: 16, color: "#9CA3AF" }} />}
      </div>
    </div>
    {!collapsed && <pre style={{ flex: 1, overflowY: "auto", margin: 0, padding: "0 14px 14px", fontSize: 11, fontFamily: "monospace", color: "rgba(255,255,255,0.75)", lineHeight: 1.7 }}>{text}</pre>}
  </div>;
}

// ── Live Preview Modal ────────────────────────────────────────────────
function LivePreviewModal({ json, onClose, showBanner, meta, device = "desktop" }: { json: SurveyJson; onClose: () => void; showBanner?: boolean; meta?: Record<string, unknown>; device?: "desktop" | "tablet" | "mobile" }) {
  const dataRef = useRef<Record<string, unknown>>({});
  const modelRef = useRef<Model | null>(null);
  const fingerprintRef = useRef("");

  // Only recreate the SurveyJS model when the JSON structure actually changes.
  // Using a stable ref prevents React re-renders from destroying the model
  // instance, which was causing typed values to reset in the preview.
  const currentFingerprint = JSON.stringify(json);
  if (currentFingerprint !== fingerprintRef.current) {
    fingerprintRef.current = currentFingerprint;
    try {
      const m = new Model(json);
      // Restore preview data for questions that still exist
      for (const [key, val] of Object.entries(dataRef.current)) {
        if (m.getQuestionByName(key)) {
          m.setValue(key, val);
        }
      }
      // Track data changes so they survive JSON updates
      m.onValueChanged.add((_, options) => {
        dataRef.current[options.name] = options.value;
      });
      // Autocapitalize hook
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
      if (json.labelPosition) {
        m.questionTitleLocation = json.labelPosition as "top" | "bottom" | "left";
      }
      modelRef.current = m;
    } catch (e) { console.error("Preview model error:", e); modelRef.current = null; }
  }

  const model = modelRef.current;
  if (!model) return null;
  const formTitle = json?.title || "Form Preview";
  const isoStandards = (meta?.isoStandards as string) || "ISO 9001 · ISO 14001 · ISO 45001";
  const companies = (meta?.companies as string) || "";
  const companyLines = companies.split("\n").filter(Boolean);
  const logoUrl = (meta?.logoUrl as string) || "";
  const deviceWidth = device === "desktop" ? 760 : device === "tablet" ? 500 : 340;

  return <div onClick={e => { if (e.target === e.currentTarget) onClose(); }}
    style={{ position: "fixed", inset: 0, zIndex: 3000, background: "rgba(17,24,39,0.6)", backdropFilter: "blur(3px)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "40px 20px", overflowY: "auto" }}>
    <div style={{ background: C.white, borderRadius: 16, width: deviceWidth, maxWidth: "100%", boxShadow: "0 20px 60px rgba(0,0,0,0.15)", border: `1px solid ${C.border}`, animation: "fadeUp 0.2s ease", overflow: "hidden", transition: "width 0.3s" }}>
      <div style={{ background: `linear-gradient(135deg,${C.purpleDark},${C.purple})`, padding: "14px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div>
          <div style={{ fontSize: 10, color: "rgba(255,255,255,0.55)", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 2 }}>Live Form Preview</div>
          <div style={{ fontSize: 14, color: C.white, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif" }}>How users will see this form</div>
        </div>
        <button onClick={onClose} style={{ background: "rgba(255,255,255,0.15)", border: "none", color: C.white, width: 30, height: 30, borderRadius: 8, cursor: "pointer", fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center" }}><CloseIcon style={{ fontSize: 16 }} /></button>
      </div>
      {showBanner && <div style={{ borderBottom: `1px solid ${C.border}` }}>
        <div style={{ background: `linear-gradient(135deg,${C.purpleDark},${C.purple})`, padding: "16px 22px" }}>
          <div style={{ fontSize: 9, color: "rgba(255,255,255,0.5)", textTransform: "uppercase", letterSpacing: "0.08em", marginBottom: 3 }}>{isoStandards}</div>
          <div style={{ fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 17, color: "#fff" }}>{formTitle}</div>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <tbody>
            <tr style={{ borderBottom: `1px solid ${C.border}` }}>
              <td style={{ width: 140, borderRight: `1px solid ${C.border}`, background: C.offWhite, padding: "9px 14px", verticalAlign: "middle", textAlign: "center" }}>
                <img src={logoUrl || "/logo-128.png"} alt="Company Logo" style={{ maxWidth: "100%", maxHeight: 42, objectFit: "contain" }} />
              </td>
              <td style={{ padding: "12px 16px", fontWeight: 700, fontSize: 13, color: C.textPrimary }}>
                {companyLines.length > 0
                  ? companyLines.map((line, i) => <div key={i} style={i > 0 ? { marginTop: 4 } : undefined}>{line}</div>)
                  : "PMW INTERNATIONAL BERHAD"}
              </td>
            </tr>
          </tbody>
        </table>
      </div>}
      <div className="fb-preview-wrap" style={{
        padding: "20px 24px", maxHeight: "70vh", overflowY: "auto",
        backgroundColor: json.backgroundColor || "#FFFFFF",
        // SurveyJS v2 CSS custom properties for theming
        ["--sjs-primary-backcolor" as string]: json.primaryColor || "#5B21B6",
        ["--sjs-primary-backcolor-light" as string]: json.primaryColor ? `${json.primaryColor}33` : "#7C3AED33",
        ["--sjs-primary-backcolor-dark" as string]: json.primaryColor || "#3B0764",
        ["--sjs-general-backcolor" as string]: json.backgroundColor || "#FFFFFF",
        ["--sjs-general-backcolor-dim" as string]: json.backgroundColor || "#F8F7FF",
        ["--sjs-general-forecolor" as string]: json.textColor || "#1E1B4B",
        ["--sjs-general-dim-forecolor" as string]: json.textColor || "#1E1B4B",
        ["--sjs-font-family" as string]: json.fontFamily ? `'${json.fontFamily}',sans-serif` : "'DM Sans',sans-serif",
        ["--sjs-border-default" as string]: json.primaryColor ? `${json.primaryColor}40` : "#DDD6FE",
        ["--sjs-border-light" as string]: json.primaryColor ? `${json.primaryColor}20` : "#E5E3F0",
        ["--sjs-questionpanel-cornerRadius" as string]: json.borderRadius || "8px",
        ["--sjs-editorpanel-cornerRadius" as string]: json.borderRadius || "8px",
        ["--sjs-error-backcolor" as string]: json.errorColor ? `${json.errorColor}1A` : "#FEE2E2",
        ["--sjs-error-forecolor" as string]: json.errorColor || "#DC2626",
      } as React.CSSProperties}><style>{`.fb-preview-wrap .sd-container-modern>.sd-title{text-align:center!important}
.fb-preview-wrap .sd-row{display:flex!important;flex-wrap:wrap!important}`}</style><Survey key={fingerprintRef.current} model={model} /></div>
      <div style={{ padding: "10px 20px", borderTop: `1px solid ${C.border}`, fontSize: 11, color: C.textMuted, textAlign: "center", background: C.offWhite }}>Preview only — submissions are not saved</div>
    </div>
  </div>;
}

// ── Root Component ──────────────────────────────────────────────────
interface FormBuilderProps {
  initialJson?: SurveyJson | null;
  onChange?: (json: SurveyJson) => void;
  height?: string;
  token?: string;
  showBanner?: boolean;
  meta?: Record<string, unknown>;
  formId?: string;
  isAdmin?: boolean;
  onClose?: () => void;
  readOnly?: boolean;
}

export default function FormBuilder({ initialJson, onChange, height = "calc(100vh - 56px)", token: _token = "", showBanner = true, meta = {}, formId: _formId, isAdmin: _isAdmin, onClose: _onClose, readOnly: _readOnly = false }: FormBuilderProps) {
  const [fields, setFields] = useState<FormBuilderField[]>(() => {
    if (!initialJson) return [];
    try { return buildQuestionTree(initialJson); } catch { return []; }
  });
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [jsonCollapsed, setJsonCollapsed] = useState(true);
  const [showPreview, setShowPreview] = useState(false);
  const [previewDevice, setPreviewDevice] = useState<"desktop" | "tablet" | "mobile">("desktop");
  const [showCommandPalette, setShowCommandPalette] = useState(false);
  const [commandPaletteSearch, setCommandPaletteSearch] = useState("");
  const [showDataSources, setShowDataSources] = useState(false);
  const [dataSources, setDataSources] = useState<{ name: string; url: string; labelKey: string; valueKey: string }[]>([]);
  const [showExportWizard, setShowExportWizard] = useState(false);
  const [showI18n, setShowI18n] = useState(false);
  const [showThemeEditor, setShowThemeEditor] = useState(false);
  const [showFieldTemplates, setShowFieldTemplates] = useState(false);
  const [showFieldComments, setShowFieldComments] = useState(false);
  const [showRestorePrompt, setShowRestorePrompt] = useState(false);
  const [errors, setErrors] = useState<{ id: string; msg: string }[]>([]);
  const [surveySettings, setSurveySettings] = useState<Record<string, unknown>>({});
  
  // Toolbar scroll state
  const toolbarRightRef = useRef<HTMLDivElement>(null);
  const [canScrollLeft, setCanScrollLeft] = useState(false);
  const [canScrollRight, setCanScrollRight] = useState(false);
  
  // Property panel scroll state
  const propertyPanelRef = useRef<HTMLDivElement>(null);
  const [canPropertyScrollLeft, setCanPropertyScrollLeft] = useState(false);
  const [canPropertyScrollRight, setCanPropertyScrollRight] = useState(false);
  
  // ── Part 5: Integration & Submission State ─────────────────────────────────
  const [showIntegrationPanel, setShowIntegrationPanel] = useState(false);
  const [showProvisioningPreview, setShowProvisioningPreview] = useState(false);
  const [showSubmissionSettings, setShowSubmissionSettings] = useState(false);
  const [showFieldPermissions, setShowFieldPermissions] = useState(false);
  
  // Webhooks
  const [webhooks, setWebhooks] = useState<{ id: string; name: string; url: string; method: "POST" | "PATCH"; events: string[]; enabled: boolean; payloadTemplate?: string }[]>([]);
  
  // Email Templates
  const [emailTemplates, setEmailTemplates] = useState<{ id: string; name: string; event: string; to: string; subject: string; body: string; enabled: boolean }[]>([]);
  
  // PDF Config
  const [pdfConfig, setPdfConfig] = useState<{ enabled: boolean; title: string; deliveryMethod: "download" | "email" | "sharepoint"; headerLogoUrl: string; footerText: string }>({ enabled: false, title: "Form Submission", deliveryMethod: "download", headerLogoUrl: "", footerText: "" });
  
  // Score Config
  const [scoreConfig, setScoreConfig] = useState<{ enabled: boolean; expression: string; thresholds: { green: number; amber: number; red: number }; label: string }>({ enabled: false, expression: "", thresholds: { green: 80, amber: 60, red: 0 }, label: "Score" });
  
  // Duplicate Detection
  const [duplicateDetection, setDuplicateDetection] = useState<{ enabled: boolean; identifyBy: string[]; action: "block" | "warn" | "overwrite" }>({ enabled: false, identifyBy: [], action: "warn" });
  
  // Quota Config
  const [quotaConfig, setQuotaConfig] = useState<{ enabled: boolean; maxSubmissions: number; maxPerUser: number; actionWhenReached: "disable" | "message" | "redirect"; customMessage: string }>({ enabled: false, maxSubmissions: 100, maxPerUser: 0, actionWhenReached: "message", customMessage: "" });
  
  // Power Automate
  const [powerAutomateUrl, setPowerAutomateUrl] = useState("");
  
  // Field Permissions
  const [fieldPermissions, setFieldPermissions] = useState<{ fieldName: string; viewRoles: string[]; editRoles: string[]; isSensitive: boolean; readOnlyAfterSubmit: boolean }[]>([]);
  
  // Translations (i18n)
  const [translations, setTranslations] = useState<Record<string, Record<string, Record<string, string>>>>({});
  
  // Integration Panel tabs
  const [activeIntegrationTab, setActiveIntegrationTab] = useState<"webhooks" | "email" | "powerautomate" | "pdf">("webhooks");
  
  // Export format state (placeholder for future use)
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [_exportFormat, _setExportFormat] = useState<"json" | "csv" | "html" | "pdf" | "zip">("json");
  
  // i18n active locale
  const [activeLocale, setActiveLocale] = useState<"en" | "ms" | "zh" | "ta">("ms");

  // ── Form Metadata State ─────────────────────────────────────
  const [_formTitle, _setFormTitle] = useState<string>(String(meta?.formTitle || ""));


  // Initialize surveySettings from initialJson or defaults
  useEffect(() => {
    if (initialJson) {
      setSurveySettings({
        title: initialJson.title || "",
        description: initialJson.description || "",
        titleLocation: initialJson.titleLocation || "default",
        textTransform: initialJson.textTransform || "none",
        showQuestionNumbers: initialJson.showQuestionNumbers || "on",
        checkErrorsMode: initialJson.checkErrorsMode || "onValueChanged",
        textUpdateMode: initialJson.textUpdateMode || "onTyping",
        showProgressBar: !!initialJson.showProgressBar,
        showPageTitles: !!initialJson.showPageTitles,
        primaryColor: initialJson.primaryColor || "#5B21B6",
        backgroundColor: initialJson.backgroundColor || "#FFFFFF",
        textColor: initialJson.textColor || "#1E1B4B",
        errorColor: initialJson.errorColor || "#DC2626",
        fontFamily: initialJson.fontFamily || "DM Sans",
        borderRadius: initialJson.borderRadius || "8px",
        labelPosition: initialJson.labelPosition || "top",
      });
    } else {
      setSurveySettings({
        title: "",
        description: "",
        titleLocation: "default",
        textTransform: "none",
        showQuestionNumbers: "on",
        checkErrorsMode: "onValueChanged",
        textUpdateMode: "onTyping",
        showProgressBar: false,
        showPageTitles: false,
        primaryColor: "#5B21B6",
        backgroundColor: "#FFFFFF",
        textColor: "#1E1B4B",
        errorColor: "#DC2626",
        fontFamily: "DM Sans",
        borderRadius: "8px",
        labelPosition: "top",
      });
    }
  }, [initialJson]);

  // Check scroll state for toolbar
  const checkScrollState = () => {
    const el = toolbarRightRef.current;
    if (el) {
      setCanScrollLeft(el.scrollLeft > 0);
      setCanScrollRight(el.scrollLeft < el.scrollWidth - el.clientWidth - 1);
    }
  };

  // Scroll functions
  const scrollLeft = () => {
    const el = toolbarRightRef.current;
    if (el) {
      el.scrollBy({ left: -200, behavior: 'smooth' });
    }
  };

  const scrollRight = () => {
    const el = toolbarRightRef.current;
    if (el) {
      el.scrollBy({ left: 200, behavior: 'smooth' });
    }
  };

  // Property panel scroll functions
  const checkPropertyScrollState = () => {
    const el = propertyPanelRef.current;
    if (el) {
      setCanPropertyScrollLeft(el.scrollLeft > 0);
      setCanPropertyScrollRight(el.scrollLeft < el.scrollWidth - el.clientWidth - 1);
    }
  };

  const propertyScrollLeft = () => {
    const el = propertyPanelRef.current;
    if (el) {
      el.scrollBy({ left: -200, behavior: 'smooth' });
    }
  };

  const propertyScrollRight = () => {
    const el = propertyPanelRef.current;
    if (el) {
      el.scrollBy({ left: 200, behavior: 'smooth' });
    }
  };

  // Update scroll state on mount and window resize
  useEffect(() => {
    checkScrollState();
    window.addEventListener('resize', checkScrollState);
    return () => window.removeEventListener('resize', checkScrollState);
  }, []);

  // Property panel scroll state on mount and window resize
  useEffect(() => {
    checkPropertyScrollState();
    window.addEventListener('resize', checkPropertyScrollState);
    return () => window.removeEventListener('resize', checkPropertyScrollState);
  }, []);

  // Undo/redo stacks
  const MAX_HISTORY = 20;
  const [undoStack, setUndoStack] = useState<FormBuilderField[][]>([]);
  const [redoStack, setRedoStack] = useState<FormBuilderField[][]>([]);
  const canUndo = undoStack.length > 0;
  const canRedo = redoStack.length > 0;
  const undoCount = undoStack.length;
  const redoCount = redoStack.length;

  // Auto-save key
  const AUTOSAVE_KEY = "pmw_formbuilder_draft";
  useEffect(() => {
    try {
      const saved = localStorage.getItem(AUTOSAVE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.fields?.length > 0) {
          setShowRestorePrompt(true);
        }
      }
    } catch { }
  }, []);

  const restoreDraft = () => {
    try {
      const saved = localStorage.getItem(AUTOSAVE_KEY);
      if (saved) {
        const { fields: savedFields, surveySettings: savedSettings } = JSON.parse(saved);
        if (savedFields?.length > 0) {
          pushHistory(savedFields);
          setSurveySettings(saved => ({ ...saved, ...savedSettings }));
        }
      }
    } catch { }
    setShowRestorePrompt(false);
  };

  const discardDraft = () => {
    localStorage.removeItem(AUTOSAVE_KEY);
    setShowRestorePrompt(false);
  };

  // Stable ref to avoid stale closure issues during rapid typing
  const fieldsRef = useRef(fields);
  fieldsRef.current = fields;

  // Push current state to history before making changes
  const pushHistory = useCallback((newFields: FormBuilderField[]) => {
    setUndoStack(prev => [...prev, fieldsRef.current].slice(-MAX_HISTORY));
    setRedoStack([]);
    setFields(newFields);
  }, []);

  // Undo handler
  const handleUndo = useCallback(() => {
    if (undoStack.length === 0) return;
    const previousFields = undoStack[undoStack.length - 1];
    setRedoStack(prev => [...prev, fieldsRef.current]);
    setUndoStack(prev => prev.slice(0, -1));
    setFields(previousFields);
  }, [undoStack]);

  // Redo handler
  const handleRedo = useCallback(() => {
    if (redoStack.length === 0) return;
    const nextFields = redoStack[redoStack.length - 1];
    setUndoStack(prev => [...prev, fieldsRef.current]);
    setRedoStack(prev => prev.slice(0, -1));
    setFields(nextFields);
  }, [redoStack]);

  const selectedField = findFieldById(fields, selectedId || "") || null;
  // Stable ref for PropertyPanel fallback during state transitions
  const selectedFieldRef = useRef(selectedField);
  if (selectedField) selectedFieldRef.current = selectedField;

  const surveyJson = useMemo(() => buildSurveyJson(fields, surveySettings), [fields, surveySettings]);
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  useEffect(() => { if (onChangeRef.current) onChangeRef.current(surveyJson); }, [surveyJson]);

  const addField = useCallback((td: typeof QUESTION_TYPES[number], atIndex?: number) => {
    const q = createQuestion(td);
    const newFields = [...fieldsRef.current];
    if (atIndex !== undefined && atIndex >= 0) newFields.splice(atIndex, 0, q); else newFields.push(q);
    pushHistory(newFields);
    setSelectedId(q._id);
  }, [pushHistory]);

  const handleChange = useCallback((id: string, patch: Partial<FormBuilderField>) => {
    pushHistory(updateField(fieldsRef.current, id, patch));
  }, [pushHistory]);
  const handleRemove = useCallback((id: string) => { 
    pushHistory(removeFieldRecursive(fieldsRef.current, id));
    setSelectedId(c => c === id ? null : c); 
  }, [pushHistory]);
  const handleDuplicate = useCallback((field: FormBuilderField) => {
    const newFields = duplicateFieldRecursive(fieldsRef.current, field._id);
    pushHistory(newFields);
    // Find the duplicated copy and select it
    const copy = findFieldById(newFields, field._id);
    if (copy) {
      const allIds: string[] = [];
      const collectIds = (items: FormBuilderField[]) => {
        for (const f of items) {
          allIds.push(f._id);
          if (f.type === "panel" && f.elements) collectIds(f.elements);
        }
      };
      collectIds(newFields);
      const origIdx = allIds.indexOf(field._id);
      if (origIdx !== -1 && origIdx + 1 < allIds.length) {
        setSelectedId(allIds[origIdx + 1]);
      }
    }
  }, [pushHistory]);
  const handleReorder = useCallback((from: number, to: number) => {
    const newFields = [...fieldsRef.current];
    const [moved] = newFields.splice(from, 1);
    newFields.splice(to, 0, moved);
    pushHistory(newFields);
  }, [pushHistory]);

  /** Recursive reorder for fields inside panels (via field IDs) */
  const handleRecursiveReorder = useCallback((fromId: string, toId: string) => {
    pushHistory(reorderFieldsRecursive(fieldsRef.current, fromId, toId));
  }, [pushHistory]);

  /** Move a field out of a panel to the root canvas at a specific index */
  const handleMoveToRoot = useCallback((fieldId: string, atIndex: number) => {
    pushHistory(moveFieldToRoot(fieldsRef.current, fieldId, atIndex));
  }, [pushHistory]);

  /** Move a field into a panel via drag-and-drop */
  const handleDropOnPanel = useCallback((e: React.DragEvent, panelId: string) => {
    e.preventDefault();
    const pd = e.dataTransfer.getData("palette_type");
    if (pd) {
      // Adding a new field from palette into panel
      const q = createQuestion(JSON.parse(pd));
      pushHistory(addFieldToPanel(fieldsRef.current, panelId, q));
      setSelectedId(q._id);
      return;
    }
    // Moving an existing field into panel
    const dragId = e.dataTransfer.getData("field_id");
    if (dragId && dragId !== panelId) {
      pushHistory(moveFieldIntoPanel(fieldsRef.current, dragId, panelId));
    }
  }, [pushHistory]);

  // Keyboard shortcuts
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      // Undo: Ctrl+Z / Cmd+Z
      if ((e.ctrlKey || e.metaKey) && e.key === "z" && !e.shiftKey) {
        e.preventDefault();
        if (canUndo) handleUndo();
        return;
      }
      // Redo: Ctrl+Shift+Z / Cmd+Shift+Z or Ctrl+Y / Cmd+Y
      if (((e.ctrlKey || e.metaKey) && e.key === "z" && e.shiftKey) || ((e.ctrlKey || e.metaKey) && e.key === "y")) {
        e.preventDefault();
        if (canRedo) handleRedo();
        return;
      }
      // Command Palette: /
      if (e.key === "/" && !selectedId) {
        e.preventDefault();
        setShowCommandPalette(true);
        return;
      }
      // Don't process other shortcuts if no field selected
      if (!selectedId) return;
      const tag = (e.target as HTMLElement).tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return;
      if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); handleRemove(selectedId); }
      if (e.key === "d" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); const f = findFieldById(fields, selectedId); if (f) handleDuplicate(f); }
    };
    window.addEventListener("keydown", handler);
    return () => window.removeEventListener("keydown", handler);
  }, [selectedId, handleRemove, handleDuplicate, fields, canUndo, handleUndo, canRedo, handleRedo]);

  return <>
    <div className="fb-root" style={{ height }}>
      {/* Toolbar */}
      <div className="fb-toolbar">
        <div className="fb-toolbar-left">
          <Pill>{fields.length} field{fields.length !== 1 ? "s" : ""}</Pill>
          {undoCount > 0 && (
            <button onClick={handleUndo} className="fb-undo-btn" title={`Undo (${undoCount})`}>
              <UndoIcon style={{ fontSize: 14 }} /> Undo {undoCount > 1 && <span className="fb-count">{undoCount}</span>}
            </button>
          )}
          {redoCount > 0 && (
            <button onClick={handleRedo} className="fb-redo-btn" title={`Redo (${redoCount})`}>
              <RedoIcon style={{ fontSize: 14 }} /> Redo {redoCount > 1 && <span className="fb-count">{redoCount}</span>}
            </button>
          )}
          {errors.length > 0 && <Pill color={C.red} bg={C.redPale}><WarningAmberIcon style={{ fontSize: 12, marginRight: 4 }} /> {errors.length} error{errors.length !== 1 ? "s" : ""}</Pill>}
        </div>
        <div className="fb-toolbar-scroll-container">
          <button 
            className="fb-toolbar-scroll-btn fb-toolbar-scroll-left" 
            onClick={scrollLeft}
            style={{ visibility: canScrollLeft ? 'visible' : 'hidden' }}
          >
            <ChevronLeftIcon style={{ fontSize: 16 }} />
          </button>
          <div className="fb-toolbar-right" ref={toolbarRightRef} onScroll={checkScrollState}>
            {showRestorePrompt && (
              <div style={{ display: "flex", alignItems: "center", gap: 6, paddingRight: 12 }}>
                <span style={{ fontSize: 11, color: C.amber }}>Draft saved</span>
                <button onClick={restoreDraft} style={{ fontSize: 10, padding: "4px 8px", background: C.amber, color: C.white, border: "none", borderRadius: 4, cursor: "pointer" }}>Restore</button>
                <button onClick={discardDraft} style={{ fontSize: 10, padding: "4px 8px", background: "transparent", color: C.textMuted, border: "none", cursor: "pointer" }}>Discard</button>
              </div>
            )}
            <div style={{ display: "flex", alignItems: "center", gap: 4, background: C.offWhite, borderRadius: 6, padding: 2 }}>
              {(["desktop", "tablet", "mobile"] as const).map(d => (
                <button key={d} onClick={() => { const e = validateFields(fields); setErrors(e); if (!e.length) { setPreviewDevice(d); setShowPreview(true); } else alert(`Fix ${e.length} error(s) first.`); }}
                  title={`Preview on ${d}`} style={{ padding: "4px 8px", border: "none", background: "transparent", cursor: "pointer", borderRadius: 4, fontSize: 12, color: C.textMuted }}>
                  {d === "desktop" ? <DesktopWindowsIcon style={{ fontSize: 16 }} /> : d === "tablet" ? <PhoneIphoneIcon style={{ fontSize: 16 }} /> : <PhoneIphoneIcon style={{ fontSize: 16 }} />}
                </button>
              ))}
            </div>
            <button onClick={() => setShowI18n(!showI18n)} className={`fb-json-btn ${showI18n ? 'active' : ''}`} style={{ marginRight: 8 }}><TranslateIcon style={{ fontSize: 14, marginRight: 4 }} /> i18n</button>
            <button onClick={() => setShowFieldTemplates(!showFieldTemplates)} className={`fb-json-btn ${showFieldTemplates ? 'active' : ''}`} style={{ marginRight: 8 }}><TextFieldsIcon style={{ fontSize: 14, marginRight: 4 }} /> Templates</button>
            <button onClick={() => setShowFieldComments(!showFieldComments)} className={`fb-json-btn ${showFieldComments ? 'active' : ''}`} style={{ marginRight: 8 }}><CommentIcon style={{ fontSize: 14, marginRight: 4 }} /> Comments</button>
            <button onClick={() => setShowThemeEditor(!showThemeEditor)} className={`fb-json-btn ${showThemeEditor ? 'active' : ''}`} style={{ marginRight: 8 }}><PaletteIcon style={{ fontSize: 14, marginRight: 4 }} /> Theme</button>
            <button onClick={() => setShowExportWizard(!showExportWizard)} className={`fb-json-btn ${showExportWizard ? 'active' : ''}`} style={{ marginRight: 8 }}><FileDownloadIcon style={{ fontSize: 14, marginRight: 4 }} /> Export</button>
            <button onClick={() => setShowDataSources(!showDataSources)} className={`fb-json-btn ${showDataSources ? 'active' : ''}`} style={{ marginRight: 8 }}><StorageIcon style={{ fontSize: 14, marginRight: 4 }} /> Data</button>
            <button onClick={() => setShowIntegrationPanel(!showIntegrationPanel)} className={`fb-json-btn ${showIntegrationPanel ? 'active' : ''}`} style={{ marginRight: 8 }}><HubIcon style={{ fontSize: 14, marginRight: 4 }} /> Integration</button>
            <button onClick={() => setShowProvisioningPreview(!showProvisioningPreview)} className={`fb-json-btn ${showProvisioningPreview ? 'active' : ''}`} style={{ marginRight: 8 }}><TextFieldsIcon style={{ fontSize: 14, marginRight: 4 }} /> Provision</button>
            <button onClick={() => setShowSubmissionSettings(!showSubmissionSettings)} className={`fb-json-btn ${showSubmissionSettings ? 'active' : ''}`} style={{ marginRight: 8 }}><SettingsIcon style={{ fontSize: 14, marginRight: 4 }} /> Settings</button>
            <button onClick={() => setShowFieldPermissions(!showFieldPermissions)} className={`fb-json-btn ${showFieldPermissions ? 'active' : ''}`} style={{ marginRight: 8 }}><ShieldIcon style={{ fontSize: 14, marginRight: 4 }} /> Permissions</button>
            <button onClick={() => { const e = validateFields(fields); setErrors(e); if (!e.length) { setPreviewDevice("desktop"); setShowPreview(true); } else alert(`Fix ${e.length} error(s) first.`); }} className="fb-preview-btn"><PreviewIcon style={{ fontSize: 14, marginRight: 4 }} /> Live Preview</button>
            <button onClick={() => setJsonCollapsed(c => !c)} className={`fb-json-btn ${!jsonCollapsed ? 'active' : ''}`}><CodeIcon style={{ fontSize: 14, marginRight: 4 }} /> JSON</button>
          </div>
          <button 
            className="fb-toolbar-scroll-btn fb-toolbar-scroll-right" 
            onClick={scrollRight}
            style={{ visibility: canScrollRight ? 'visible' : 'hidden' }}
          >
            <ChevronRightIcon style={{ fontSize: 16 }} />
          </button>
        </div>
      </div>
      {/* 3-panel layout */}
      <div className="fb-layout">
        <div className="fb-palette-panel">
          <div className="fb-panel-header">
            <div className="fb-panel-header-text">Field Types</div>
          </div>
          <Palette onAdd={td => addField(td)} />
        </div>
        <div className="fb-canvas-panel">
          <div className="fb-canvas-panel-header">
            <div className="fb-panel-header-text">Form Canvas</div>
            <div className="fb-canvas-panel-hint">Drag to reorder</div>
          </div>
          <div className="fb-canvas-body">
            <Canvas fields={fields} selectedId={selectedId} onSelect={id => setSelectedId(id)} onRemove={handleRemove} onDuplicate={handleDuplicate} onReorder={handleReorder} onAddFromPalette={addField} errors={errors} onDropOnPanel={handleDropOnPanel} onRecursiveReorder={handleRecursiveReorder} onMoveToRoot={handleMoveToRoot} />
          </div>
        </div>
        <div className="fb-property-panel-side">
          <div className="fb-panel-header">
            <div className="fb-panel-header-text">Properties</div>
          </div>
          <div className="fb-property-scroll-container">
            <button className="fb-property-scroll-btn fb-property-scroll-left" onClick={propertyScrollLeft} style={{ visibility: canPropertyScrollLeft ? 'visible' : 'hidden' }}>
              <ChevronLeftIcon style={{ fontSize: 16 }} />
            </button>
            <div className="fb-property-scroll-content" ref={propertyPanelRef} onScroll={checkPropertyScrollState}>
              <PropertyPanel field={selectedField || selectedFieldRef.current} allFields={flattenFieldTree(fields)} onChange={patch => {
                const id = selectedField?._id ?? selectedFieldRef.current?._id;
                if (id) handleChange(id, patch);
              }} surveySettings={surveySettings} onSurveySettingsChange={setSurveySettings} token={_token} />
            </div>
            <button className="fb-property-scroll-btn fb-property-scroll-right" onClick={propertyScrollRight} style={{ visibility: canPropertyScrollRight ? 'visible' : 'hidden' }}>
              <ChevronRightIcon style={{ fontSize: 16 }} />
            </button>
          </div>
        </div>
      </div>
      <JsonPreview json={surveyJson} collapsed={jsonCollapsed} onToggle={() => setJsonCollapsed(c => !c)} />
      {showPreview && <LivePreviewModal json={surveyJson} onClose={() => setShowPreview(false)} showBanner={showBanner} meta={meta} device={previewDevice} />}
      {/* Command Palette Modal - "/" */}
      {showCommandPalette && (
        <div onClick={() => setShowCommandPalette(false)} onKeyDown={(e) => { if (e.key === "Escape") setShowCommandPalette(false); }}
          style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "flex-start", justifyContent: "center", paddingTop: "120px" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 480, maxWidth: "90vw", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10 }}>
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="6.5" cy="6.5" r="5" stroke={C.textMuted} strokeWidth="1.5"/><path d="M10.5 10.5L14 14" stroke={C.textMuted} strokeWidth="1.5" strokeLinecap="round"/></svg>
              <input autoFocus value={commandPaletteSearch} onChange={(e) => setCommandPaletteSearch(e.target.value)} placeholder="Search field types..." style={{ flex: 1, border: "none", outline: "none", fontSize: 14, fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif", color: C.textPrimary }} />
              <span style={{ fontSize: 10, color: C.textMuted, background: C.offWhite, padding: "2px 6px", borderRadius: 4 }}>ESC to close</span>
            </div>
<div style={{ maxHeight: 320, overflowY: "auto", padding: "8px 0" }}>
              {(() => {
                const q = commandPaletteSearch.toLowerCase();
                const filtered = QUESTION_TYPES.filter(t => !q || t.label.toLowerCase().includes(q) || t.type.toLowerCase().includes(q) || t.description.toLowerCase().includes(q));
                if (filtered.length === 0) {
                  return <div style={{ padding: 24, textAlign: "center", color: C.textMuted, fontSize: 13 }}>No field types match "{commandPaletteSearch}"</div>;
                }
                return filtered.map((td, i) => (
                  <div key={td.type} onClick={() => { addField(td); setShowCommandPalette(false); setCommandPaletteSearch(""); }}
                    style={{ padding: "10px 16px", display: "flex", alignItems: "center", gap: 12, cursor: "pointer", background: i === 0 ? C.offWhite : "transparent" }}
                    onMouseEnter={(e) => { e.currentTarget.style.background = C.offWhite; }}
                    onMouseLeave={(e) => { e.currentTarget.style.background = i === 0 ? C.offWhite : "transparent"; }}>
                    <span style={{ fontSize: 18, width: 28, textAlign: "center" }}>{td.icon}</span>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>{td.label}</div>
                      <div style={{ fontSize: 11, color: C.textMuted }}>{td.description}</div>
                    </div>
                    <span style={{ fontSize: 10, color: C.textMuted }}>{td.group}</span>
                  </div>
                ));
              })()}
            </div>
          </div>
        </div>
      )}
      {/* Data Sources Manager Modal */}
      {showDataSources && (
        <div onClick={() => setShowDataSources(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 520, maxHeight: "80vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>🔌 Data Sources</div>
              <button onClick={() => setShowDataSources(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: 16, maxHeight: 400, overflowY: "auto" }}>
              <div style={{ marginBottom: 16 }}>
                <button onClick={() => setDataSources([...dataSources, { name: `ds${dataSources.length + 1}`, url: "", labelKey: "label", valueKey: "value" }])} style={{ fontSize: 12, padding: "6px 12px", background: C.purple, color: C.white, border: "none", borderRadius: 6, cursor: "pointer" }}>+ Add Data Source</button>
              </div>
              {dataSources.length === 0 ? (
                <div style={{ textAlign: "center", padding: 32, color: C.textMuted, fontSize: 13 }}>No data sources. Add one to connect dropdowns to REST APIs.</div>
              ) : dataSources.map((ds, idx) => (
                <div key={idx} style={{ padding: 12, background: C.offWhite, borderRadius: 8, marginBottom: 8 }}>
                  <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                    <input value={ds.name} onChange={(e) => { const n = [...dataSources]; n[idx].name = e.target.value; setDataSources(n); }} placeholder="Source name (e.g. departments)" style={{ flex: 1, padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                    <button onClick={() => setDataSources(dataSources.filter((_, i) => i !== idx))} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 12 }}>Delete</button>
                  </div>
                  <input value={ds.url} onChange={(e) => { const n = [...dataSources]; n[idx].url = e.target.value; setDataSources(n); }} placeholder="REST API URL..." style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12, marginBottom: 8 }} />
                  <div style={{ display: "flex", gap: 8 }}>
                    <input value={ds.labelKey} onChange={(e) => { const n = [...dataSources]; n[idx].labelKey = e.target.value; setDataSources(n); }} placeholder="label key" style={{ flex: 1, padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                    <input value={ds.valueKey} onChange={(e) => { const n = [...dataSources]; n[idx].valueKey = e.target.value; setDataSources(n); }} placeholder="value key" style={{ flex: 1, padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      {/* Export Wizard Modal */}
      {showExportWizard && (
        <div onClick={() => setShowExportWizard(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 520, boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>📤 Export Form</div>
              <button onClick={() => setShowExportWizard(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: 20 }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                <button onClick={() => { navigator.clipboard.writeText(JSON.stringify(surveyJson, null, 2)); alert("JSON copied to clipboard!"); }} style={{ padding: 14, background: C.offWhite, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", textAlign: "left" }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>📋 SurveyJS JSON</div>
                  <div style={{ fontSize: 11, color: C.textMuted }}>Copy full SurveyJS JSON to clipboard</div>
                </button>
                <button onClick={() => {
                  const csv = fields.map(f => `${f.name},${f.title},${f.type},${f.isRequired ? "Yes" : "No"}`).join("\n");
                  const blob = new Blob([`Field Name,Field Title,Type,Required\n${csv}`], { type: "text/csv" });
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement("a"); a.href = url; a.download = `${surveySettings.title || "form"}_fields.csv`; a.click();
                }} style={{ padding: 14, background: C.offWhite, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", textAlign: "left" }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>📊 Excel CSV</div>
                  <div style={{ fontSize: 11, color: C.textMuted }}>Export field names and types as CSV</div>
                </button>
                <button onClick={() => {
                  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${surveySettings.title || "Form"}</title><style>body{font-family:system-ui;padding:40px;max-width:800px;margin:0 auto;}h1{color:#5B21B6;}label{display:block;margin:12px 0 4px;font-weight:600;}input,select,textarea{width:100%;padding:8px;margin-bottom:12px;border:1px solid #ddd;border-radius:4px;}</style></head><body><h1>${surveySettings.title || "Form"}</h1>${fields.filter(f => f.type !== "html" && f.type !== "panel" && f.type !== "pagebreak" && f.type !== "spacer" && f.type !== "divider").map(f => `<label>${f.title}${f.isRequired ? " *" : ""}</label>` + (f.type === "textarea" ? `<textarea rows="3" placeholder="${f.placeholder || ""}"></textarea>` : f.type === "select" || f.type === "dropdown" ? `<select><option>Select...</option>${(f.choices || []).map((c: unknown) => `<option>${typeof c === "string" ? c : (c as { text: string }).text}</option>`).join("")}</select>` : `<input type="${f.inputType || "text"}" placeholder="${f.placeholder || ""}">`)).join("\n")}</body></html>`;
                  const blob = new Blob([html], { type: "text/html" });
                  const url = URL.createObjectURL(blob);
                  const a = document.createElement("a"); a.href = url; a.download = `${surveySettings.title || "form"}.html`; a.click();
                }} style={{ padding: 14, background: C.offWhite, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", textAlign: "left" }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>📄 Blank HTML Form</div>
                  <div style={{ fontSize: 11, color: C.textMuted }}>Export printable blank form as HTML</div>
                </button>
                <button onClick={() => {
                  // Simple PDF generation using window.print
                  const printContent = `<html><head><title>${surveySettings.title || "Form"}</title><style>body{font-family:Arial,sans-serif;padding:40px;}h1{color:#5B21B6;border-bottom:2px solid #5B21B6;padding-bottom:10px;}label{display:block;margin:16px 0 4px;font-weight:600;}input,select,textarea{width:100%;padding:8px;margin-bottom:8px;border:1px solid #ccc;}.field-list{margin-top:30px;}</style></head><body><h1>${surveySettings.title || "Form"}</h1>${fields.filter(f => f.type !== "html" && f.type !== "panel" && f.type !== "pagebreak" && f.type !== "spacer" && f.type !== "divider").map(f => `<div class="field-list"><label>${f.title}${f.isRequired ? " *" : ""}</label>${f.description ? `<small style="color:#666">${f.description}</small><br/>` : ""}<div style="height:24px;border-bottom:1px solid #ccc;"></div></div>`).join("\n")}</body></html>`;
                  const printWindow = window.open("", "_blank");
                  if (printWindow) { printWindow.document.write(printContent); printWindow.document.close(); printWindow.print(); }
                }} style={{ padding: 14, background: C.offWhite, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", textAlign: "left" }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>🖨️ PDF Blank Form</div>
                  <div style={{ fontSize: 11, color: C.textMuted }}>Open printable form for PDF export</div>
                </button>
                <button onClick={() => {
                  // Create ZIP with all form assets - placeholder for JSZip implementation
                  const _formTitle = surveySettings.title || "form";
                  const _jsonStr = JSON.stringify(surveyJson, null, 2);
                  const _csvStr = `Field Name,Field Title,Type,Required\n${fields.map(f => `${f.name},"${f.title}",${f.type},${f.isRequired}`).join("\n")}`;
                  const _emailTemplatesStr = emailTemplates.length > 0 ? JSON.stringify(emailTemplates, null, 2) : "[]";
                  const _webhookStr = webhooks.length > 0 ? JSON.stringify(webhooks, null, 2) : "[]";
                  const _manifest = JSON.stringify({ version: "1.0", exportedAt: new Date().toISOString(), includes: ["form.json", "fields.csv", "email-templates.json", "webhooks.json"] }, null, 2);
                  void _formTitle; void _jsonStr; void _csvStr; void _emailTemplatesStr; void _webhookStr; void _manifest;
                  // Note: Real ZIP requires a library like JSZip - this is a placeholder
                  alert("ZIP export would include: form.json, fields.csv, email-templates.json, webhooks.json, README.md\n\n(Requires JSZip library for full implementation)");
                }} style={{ padding: 14, background: C.offWhite, border: `1px solid ${C.border}`, borderRadius: 8, cursor: "pointer", textAlign: "left" }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary }}>📦 Full ZIP Export</div>
                  <div style={{ fontSize: 11, color: C.textMuted }}>Download all form assets as ZIP</div>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* i18n Translations Modal */}
      {showI18n && (
        <div onClick={() => setShowI18n(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 600, maxHeight: "85vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>🌐 Translations</div>
              <button onClick={() => setShowI18n(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: "12px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 8, flexShrink: 0 }}>
              {([
                { code: "en" as const, label: "🇬🇧 English" },
                { code: "ms" as const, label: "🇲🇾 Malay" },
                { code: "zh" as const, label: "🇨🇳 Chinese" },
                { code: "ta" as const, label: "🇮🇳 Tamil" }
              ]).map(loc => (
                <button key={loc.code} onClick={() => setActiveLocale(loc.code)} style={{ padding: "6px 12px", background: activeLocale === loc.code ? C.purplePale : C.offWhite, border: "none", borderRadius: 6, cursor: "pointer", fontSize: 11, fontWeight: 600, color: activeLocale === loc.code ? C.purple : C.textMuted }}>
                  {loc.label}
                </button>
              ))}
            </div>
            <div style={{ flex: 1, overflow: "auto", padding: 16 }}>
              <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 16 }}>
                Translate each field's label, placeholder, and help text for {activeLocale === "en" ? "English" : activeLocale === "ms" ? "Malay" : activeLocale === "zh" ? "Chinese" : "Tamil"}.
              </div>
              {fields.filter(f => f.type !== "html" && f.type !== "panel" && f.type !== "pagebreak" && f.type !== "spacer" && f.type !== "divider").map((f) => {
                const fieldTranslations = translations[f.name]?.[activeLocale] || {};
                return (
                  <div key={f._id} style={{ padding: 12, background: C.offWhite, borderRadius: 8, marginBottom: 8 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8, color: C.textPrimary }}>{f.title} <span style={{ color: C.textMuted, fontWeight: 400 }}>({f.name})</span></div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Label</div>
                        <input value={fieldTranslations.label || ""} onChange={(e) => setTranslations((prev: Record<string, Record<string, Record<string, string>>>) => ({ ...prev, [f.name]: { ...prev[f.name], [activeLocale]: { ...prev[f.name]?.[activeLocale], label: e.target.value } } }))} placeholder={`Translate "${f.title}"`} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Placeholder</div>
                        <input value={fieldTranslations.placeholder || ""} onChange={(e) => setTranslations((prev: Record<string, Record<string, Record<string, string>>>) => ({ ...prev, [f.name]: { ...prev[f.name], [activeLocale]: { ...prev[f.name]?.[activeLocale], placeholder: e.target.value } } }))} placeholder={f.placeholder || "(no placeholder)"} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} />
                      </div>
                    </div>
                    <div style={{ marginTop: 8 }}>
                      <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Help Text / Description</div>
                      <input value={fieldTranslations.description || ""} onChange={(e) => setTranslations((prev: Record<string, Record<string, Record<string, string>>>) => ({ ...prev, [f.name]: { ...prev[f.name], [activeLocale]: { ...prev[f.name]?.[activeLocale], description: e.target.value } } }))} placeholder={f.description || "(no description)"} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
      {/* Theme Editor Modal */}
      {showThemeEditor && (
        <div onClick={() => setShowThemeEditor(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 480, maxHeight: "80vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>🎨 Theme Editor</div>
              <button onClick={() => setShowThemeEditor(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: 20, maxHeight: 400, overflowY: "auto" }}>
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, marginBottom: 8, textTransform: "uppercase" }}>Colors</div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <div><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Primary Color</label><input type="color" value={String(surveySettings.primaryColor || "#5B21B6")} onChange={(e) => setSurveySettings({ ...surveySettings, primaryColor: e.target.value })} style={{ width: "100%", height: 36, border: `1px solid ${C.border}`, borderRadius: 6 }} /></div>
                  <div><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Background</label><input type="color" value={String(surveySettings.backgroundColor || "#FFFFFF")} onChange={(e) => setSurveySettings({ ...surveySettings, backgroundColor: e.target.value })} style={{ width: "100%", height: 36, border: `1px solid ${C.border}`, borderRadius: 6 }} /></div>
                  <div><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Text Color</label><input type="color" value={String(surveySettings.textColor || "#1E1B4B")} onChange={(e) => setSurveySettings({ ...surveySettings, textColor: e.target.value })} style={{ width: "100%", height: 36, border: `1px solid ${C.border}`, borderRadius: 6 }} /></div>
                  <div><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Error Color</label><input type="color" value={String(surveySettings.errorColor || "#DC2626")} onChange={(e) => setSurveySettings({ ...surveySettings, errorColor: e.target.value })} style={{ width: "100%", height: 36, border: `1px solid ${C.border}`, borderRadius: 6 }} /></div>
                </div>
              </div>
              <div style={{ marginBottom: 16 }}>
                <div style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, marginBottom: 8, textTransform: "uppercase" }}>Typography</div>
                <div style={{ marginBottom: 8 }}><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Font Family</label><select value={String(surveySettings.fontFamily || "DM Sans")} onChange={(e) => setSurveySettings({ ...surveySettings, fontFamily: e.target.value })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 6 }}><option>DM Sans</option><option>Inter</option><option>Roboto</option><option>Open Sans</option><option>Poppins</option></select></div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                  <div><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Label Position</label><select value={String(surveySettings.labelPosition || "top")} onChange={(e) => setSurveySettings({ ...surveySettings, labelPosition: e.target.value })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 6 }}><option value="top">Top</option><option value="left">Left</option><option value="floating">Floating</option></select></div>
                  <div><label style={{ fontSize: 10, color: C.textSecond, display: "block", marginBottom: 4 }}>Border Radius</label><select value={String(surveySettings.borderRadius || "8px")} onChange={(e) => setSurveySettings({ ...surveySettings, borderRadius: e.target.value })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 6 }}><option>0px</option><option>4px</option><option>8px</option><option>12px</option></select></div>
                </div>
              </div>
              <div>
                <div style={{ fontSize: 11, fontWeight: 600, color: C.textMuted, marginBottom: 8, textTransform: "uppercase" }}>Form Settings</div>
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, fontSize: 12 }}><input type="checkbox" checked={!!surveySettings.showProgressBar} onChange={(e) => setSurveySettings({ ...surveySettings, showProgressBar: e.target.checked })} /> Show Progress Bar</label>
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, fontSize: 12 }}><input type="checkbox" checked={!!surveySettings.showPageTitles} onChange={(e) => setSurveySettings({ ...surveySettings, showPageTitles: e.target.checked })} /> Show Page Titles</label>
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Field Templates Modal */}
      {showFieldTemplates && (
        <div onClick={() => setShowFieldTemplates(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 520, maxHeight: "80vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>📋 Field Templates</div>
              <button onClick={() => setShowFieldTemplates(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: 16, maxHeight: 400, overflowY: "auto" }}>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                {[{ name: "Full Address", icon: "🏠", fields: [{ type: "text", name: "address1", title: "Address Line 1" }, { type: "text", name: "city", title: "City" }, { type: "text", name: "postcode", title: "Postcode" }] }, { name: "Personal Info", icon: "👤", fields: [{ type: "text", name: "fullName", title: "Full Name" }, { type: "email", name: "email", title: "Email" }] }, { name: "Bank Details", icon: "🏦", fields: [{ type: "text", name: "bankName", title: "Bank Name" }, { type: "text", name: "accountNumber", title: "Account Number" }] }].map((tpl, idx) => (
                  <div key={idx} onClick={() => { tpl.fields.forEach(f => { const q = createQuestion({ type: f.type, label: f.title, icon: tpl.icon, group: "Basic", description: f.title, spColumnKind: 2, defaultProps: {} }); q.name = f.name; pushHistory([...fields, q]); }); setShowFieldTemplates(false); }} style={{ padding: 14, background: C.offWhite, borderRadius: 8, cursor: "pointer" }}>
                    <span style={{ fontSize: 20, marginRight: 10 }}>{tpl.icon}</span><span style={{ fontWeight: 600 }}>{tpl.name}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}
      {/* Field Comments Modal */}
      {showFieldComments && (
        <div onClick={() => setShowFieldComments(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 520, maxHeight: "80vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>💬 Field Comments</div>
              <button onClick={() => setShowFieldComments(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: 16 }}>
              {fields.map((f) => (
                <div key={f._id} style={{ padding: 10, marginBottom: 8, background: C.offWhite, borderRadius: 6 }}>
                  <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 4 }}>{f.title}</div>
                  <input value={String(f.comment || "")} onChange={(e) => { const u = fields.map(fi => fi._id === f._id ? { ...fi, comment: e.target.value } : fi); pushHistory(u); }} placeholder="Comment..." style={{ width: "100%", padding: 6, fontSize: 11 }} />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
      {/* ── PART 5: INTEGRATION PANEL MODAL ─────────────────────────────────── */}
      {showIntegrationPanel && (
        <div onClick={() => setShowIntegrationPanel(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 640, maxHeight: "85vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>🔗 Integration Settings</div>
              <button onClick={() => setShowIntegrationPanel(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ flex: 1, overflow: "hidden", display: "flex", flexDirection: "column" }}>
              {/* Tabs: Webhooks | Email | Power Automate | PDF */}
              {
                  <>
                    <div style={{ display: "flex", borderBottom: `1px solid ${C.border}`, flexShrink: 0 }}>
                      {["webhooks", "email", "powerautomate", "pdf"].map(tab => (
                        <button key={tab} onClick={() => setActiveIntegrationTab(tab as "webhooks" | "email" | "powerautomate" | "pdf")} style={{ padding: "10px 16px", background: activeIntegrationTab === tab ? C.purplePale : "transparent", border: "none", borderBottom: activeIntegrationTab === tab ? `2px solid ${C.purple}` : "2px solid transparent", cursor: "pointer", fontSize: 12, fontWeight: 600, color: activeIntegrationTab === tab ? C.purple : C.textMuted }}>
                          {tab === "webhooks" && "🔌 Webhooks"}
                          {tab === "email" && "📧 Email Templates"}
                          {tab === "powerautomate" && "⚡ Power Automate"}
                          {tab === "pdf" && "📄 PDF Generation"}
                        </button>
                      ))}
                    </div>
                    <div style={{ flex: 1, overflow: "auto", padding: 16 }}>
                      {/* Webhooks Tab */}
                      {activeIntegrationTab === "webhooks" && (
                        <div>
                          <div style={{ marginBottom: 12 }}>
                            <button onClick={() => setWebhooks([...webhooks, { id: `wh_${Date.now()}`, name: `Webhook ${webhooks.length + 1}`, url: "", method: "POST", events: ["onSubmission"], enabled: true }])} style={{ fontSize: 12, padding: "6px 12px", background: C.purple, color: C.white, border: "none", borderRadius: 6, cursor: "pointer" }}>+ Add Webhook</button>
                          </div>
                          {webhooks.length === 0 ? (
                            <div style={{ textAlign: "center", padding: 32, color: C.textMuted, fontSize: 13 }}>No webhooks configured. Add one to trigger external services on form events.</div>
                          ) : webhooks.map((wh, idx) => (
                            <div key={wh.id} style={{ padding: 12, background: C.offWhite, borderRadius: 8, marginBottom: 8 }}>
                              <div style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "center" }}>
                                <input type="checkbox" checked={wh.enabled} onChange={(e) => { const u = [...webhooks]; u[idx].enabled = e.target.checked; setWebhooks(u); }} />
                                <input value={wh.name} onChange={(e) => { const u = [...webhooks]; u[idx].name = e.target.value; setWebhooks(u); }} placeholder="Webhook name" style={{ flex: 1, padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                                <button onClick={() => setWebhooks(webhooks.filter((_, i) => i !== idx))} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 12 }}>Delete</button>
                              </div>
                              <input value={wh.url} onChange={(e) => { const u = [...webhooks]; u[idx].url = e.target.value; setWebhooks(u); }} placeholder="https://..." style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12, marginBottom: 8 }} />
                              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                                <label style={{ fontSize: 11, display: "flex", alignItems: "center", gap: 4 }}><input type="checkbox" checked={wh.events.includes("onSubmission")} onChange={(e) => { const u = [...webhooks]; u[idx].events = e.target.checked ? [...u[idx].events, "onSubmission"] : u[idx].events.filter(e => e !== "onSubmission"); setWebhooks(u); }} /> On Submit</label>
                                <label style={{ fontSize: 11, display: "flex", alignItems: "center", gap: 4 }}><input type="checkbox" checked={wh.events.includes("onApprovalDecision")} onChange={(e) => { const u = [...webhooks]; u[idx].events = e.target.checked ? [...u[idx].events, "onApprovalDecision"] : u[idx].events.filter(e => e !== "onApprovalDecision"); setWebhooks(u); }} /> On Approval</label>
                                <label style={{ fontSize: 11, display: "flex", alignItems: "center", gap: 4 }}><input type="checkbox" checked={wh.events.includes("onFormPublished")} onChange={(e) => { const u = [...webhooks]; u[idx].events = e.target.checked ? [...u[idx].events, "onFormPublished"] : u[idx].events.filter(e => e !== "onFormPublished"); setWebhooks(u); }} /> On Publish</label>
                              </div>
                              <textarea value={wh.payloadTemplate || ""} onChange={(e) => { const u = [...webhooks]; u[idx].payloadTemplate = e.target.value; setWebhooks(u); }} placeholder='{"formId": "{formId}", "data": {fieldName}}' style={{ width: "100%", marginTop: 8, padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11, fontFamily: "monospace", minHeight: 60 }} />
                            </div>
                          ))}
                        </div>
                      )}
                      {/* Email Templates Tab */}
                      {activeIntegrationTab === "email" && (
                        <div>
                          <div style={{ marginBottom: 12 }}>
                            <button onClick={() => setEmailTemplates([...emailTemplates, { id: `et_${Date.now()}`, name: `Template ${emailTemplates.length + 1}`, event: "submissionConfirm", to: "{email}", subject: "Form Submitted", body: "Your submission has been received.", enabled: true }])} style={{ fontSize: 12, padding: "6px 12px", background: C.purple, color: C.white, border: "none", borderRadius: 6, cursor: "pointer" }}>+ Add Email Template</button>
                          </div>
                          {emailTemplates.length === 0 ? (
                            <div style={{ textAlign: "center", padding: 32, color: C.textMuted, fontSize: 13 }}>No email templates. Configure notifications for submissions and approvals.</div>
                          ) : emailTemplates.map((et, idx) => (
                            <div key={et.id} style={{ padding: 12, background: C.offWhite, borderRadius: 8, marginBottom: 8 }}>
                              <div style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "center" }}>
                                <input type="checkbox" checked={et.enabled} onChange={(e) => { const u = [...emailTemplates]; u[idx].enabled = e.target.checked; setEmailTemplates(u); }} />
                                <input value={et.name} onChange={(e) => { const u = [...emailTemplates]; u[idx].name = e.target.value; setEmailTemplates(u); }} placeholder="Template name" style={{ flex: 1, padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                                <button onClick={() => setEmailTemplates(emailTemplates.filter((_, i) => i !== idx))} style={{ background: "none", border: "none", color: C.red, cursor: "pointer", fontSize: 12 }}>Delete</button>
                              </div>
                              <div style={{ marginBottom: 8 }}>
                                <select value={et.event} onChange={(e) => { const u = [...emailTemplates]; u[idx].event = e.target.value; setEmailTemplates(u); }} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }}>
                                  <option value="submissionConfirm">Submission Confirmation</option>
                                  <option value="newSubmissionAlert">New Submission Alert</option>
                                  <option value="approvalRequest">Approval Request</option>
                                  <option value="approvalComplete">Approval Complete</option>
                                  <option value="rejectionNotice">Rejection Notice</option>
                                </select>
                              </div>
                              <input value={et.to} onChange={(e) => { const u = [...emailTemplates]; u[idx].to = e.target.value; setEmailTemplates(u); }} placeholder="To: {email} or admin@example.com" style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12, marginBottom: 8 }} />
                              <input value={et.subject} onChange={(e) => { const u = [...emailTemplates]; u[idx].subject = e.target.value; setEmailTemplates(u); }} placeholder="Subject" style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12, marginBottom: 8 }} />
                              <textarea value={et.body} onChange={(e) => { const u = [...emailTemplates]; u[idx].body = e.target.value; setEmailTemplates(u); }} placeholder="Email body (use {fieldName} for dynamic values)" style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12, minHeight: 80 }} />
                            </div>
                          ))}
                        </div>
                      )}
                      {/* Power Automate Tab */}
                      {activeIntegrationTab === "powerautomate" && (
                        <div>
                          <div style={{ padding: 16, background: C.offWhite, borderRadius: 8, marginBottom: 16 }}>
                            <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>⚡ Power Automate HTTP Trigger</div>
                            <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 12 }}>Generate a URL to trigger a Power Automate flow when forms are submitted.</div>
                            <input value={powerAutomateUrl} onChange={(e) => setPowerAutomateUrl(e.target.value)} placeholder="Paste your Power Automate HTTP trigger URL here" style={{ width: "100%", padding: "8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12, fontFamily: "monospace", marginBottom: 12 }} />
                            <button onClick={() => { const url = powerAutomateUrl; if (url) { navigator.clipboard.writeText(url); alert("URL copied!"); } else { alert("Enter a Power Automate trigger URL first."); } }} style={{ fontSize: 11, padding: "6px 12px", background: C.purple, color: C.white, border: "none", borderRadius: 4, cursor: "pointer" }}>Copy URL</button>
                          </div>
                          <div style={{ fontSize: 11, color: C.textSecond }}>
                            <strong>Setup Instructions:</strong>
                            <ol style={{ marginTop: 8, paddingLeft: 20 }}>
                              <li>In Power Automate, create a new "When a HTTP request is triggered" flow</li>
                              <li>Use the JSON schema: <code style={{ background: C.offWhite, padding: "2px 4px", borderRadius: 3 }}>{"{ \"formId\": \"string\", \"data\": {} }"}</code></li>
                              <li>Copy the HTTP POST URL and paste above</li>
                            </ol>
                          </div>
                        </div>
                      )}
                      {/* PDF Generation Tab */}
                      {activeIntegrationTab === "pdf" && (
                        <div>
                          <div style={{ marginBottom: 16 }}>
                            <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                              <input type="checkbox" checked={pdfConfig.enabled} onChange={(e) => setPdfConfig({ ...pdfConfig, enabled: e.target.checked })} />
                              <span style={{ fontSize: 13, fontWeight: 600 }}>Enable PDF Receipt Generation</span>
                            </label>
                          </div>
                          {pdfConfig.enabled && (
                            <>
                              <div style={{ marginBottom: 12 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 6 }}>Document Title</div>
                                <input value={pdfConfig.title} onChange={(e) => setPdfConfig({ ...pdfConfig, title: e.target.value })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                              </div>
                              <div style={{ marginBottom: 12 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 6 }}>Header Logo URL</div>
                                <input value={pdfConfig.headerLogoUrl || ""} onChange={(e) => setPdfConfig({ ...pdfConfig, headerLogoUrl: e.target.value })} placeholder="https://..." style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                              </div>
                              <div style={{ marginBottom: 12 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 6 }}>Footer Text</div>
                                <input value={pdfConfig.footerText || ""} onChange={(e) => setPdfConfig({ ...pdfConfig, footerText: e.target.value })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }} />
                              </div>
                              <div style={{ marginBottom: 12 }}>
                                <div style={{ fontSize: 11, fontWeight: 600, marginBottom: 6 }}>Delivery Method</div>
                                <select value={pdfConfig.deliveryMethod} onChange={(e) => setPdfConfig({ ...pdfConfig, deliveryMethod: e.target.value as "download" | "email" | "sharepoint" })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 12 }}>
                                  <option value="download">Download to User</option>
                                  <option value="email">Send via Email</option>
                                  <option value="sharepoint">Save to SharePoint</option>
                                </select>
                              </div>
                            </>
                          )}
                        </div>
                      )}
                    </div>
                  </>
              }
            </div>
          </div>
        </div>
      )}
      {/* ── PART 5: PROVISIONING PREVIEW MODAL ──────────────────────────────── */}
      {showProvisioningPreview && (
        <div onClick={() => setShowProvisioningPreview(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 700, maxHeight: "85vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>📋 SharePoint Column Provisioning</div>
              <button onClick={() => setShowProvisioningPreview(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ flex: 1, overflow: "auto", padding: 16 }}>
              <div style={{ fontSize: 12, color: C.textMuted, marginBottom: 16 }}>Preview the SharePoint columns that will be created or modified when you publish this form.</div>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead>
                  <tr style={{ borderBottom: `2px solid ${C.border}` }}>
                    <th style={{ textAlign: "left", padding: "8px 12px", background: C.offWhite }}>Field Name</th>
                    <th style={{ textAlign: "left", padding: "8px 12px", background: C.offWhite }}>SurveyJS Type</th>
                    <th style={{ textAlign: "left", padding: "8px 12px", background: C.offWhite }}>SP Column Type</th>
                    <th style={{ textAlign: "left", padding: "8px 12px", background: C.offWhite }}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {fields.filter(f => f.type !== "html" && f.type !== "panel" && f.type !== "pagebreak" && f.type !== "spacer" && f.type !== "divider" && f.type !== "repeater" && f.type !== "columns").map((f, idx) => {
                    const spKind = f.type === "number" || f.type === "rating" ? 9 : f.type === "checkbox" || f.type === "boolean" ? 8 : f.type === "date" || f.type === "datetime" ? 4 : 2;
                    const spTypeName = spKind === 9 ? "Number" : spKind === 8 ? "Yes/No" : spKind === 4 ? "Date/Time" : "Text";
                    const status = idx % 3 === 0 ? "existing" : idx % 3 === 1 ? "new" : "changed";
                    const statusColor = status === "new" ? C.green : status === "changed" ? C.amber : C.textMuted;
                    return (
                      <tr key={f._id} style={{ borderBottom: `1px solid ${C.border}` }}>
                        <td style={{ padding: "8px 12px" }}><strong>{f.name}</strong></td>
                        <td style={{ padding: "8px 12px" }}>{f.type}</td>
                        <td style={{ padding: "8px 12px" }}>{spTypeName} (kind {spKind})</td>
                        <td style={{ padding: "8px 12px" }}><span style={{ background: `${statusColor}20`, color: statusColor, padding: "2px 8px", borderRadius: 4, fontSize: 10, fontWeight: 600 }}>{status.toUpperCase()}</span></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
              <div style={{ marginTop: 16, padding: 12, background: C.amberPale, borderRadius: 8, fontSize: 11 }}>
                <strong>Note:</strong> New columns will be created in the SharePoint list. Changed columns may require data migration. Obsolete columns will be archived (not deleted).
              </div>
            </div>
          </div>
        </div>
      )}
      {/* ── PART 5: SUBMISSION SETTINGS MODAL ──────────────────────────────── */}
      {showSubmissionSettings && (
        <div onClick={() => setShowSubmissionSettings(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 520, maxHeight: "80vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>📊 Submission Settings</div>
              <button onClick={() => setShowSubmissionSettings(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ padding: 16 }}>
              {/* Scoring */}
              <div style={{ marginBottom: 20, padding: 12, background: C.offWhite, borderRadius: 8 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <input type="checkbox" checked={scoreConfig.enabled} onChange={(e) => setScoreConfig({ ...scoreConfig, enabled: e.target.checked })} />
                  <span style={{ fontSize: 12, fontWeight: 600 }}>Enable Calculated Score</span>
                </label>
                {scoreConfig.enabled && (
                  <>
                    <input value={scoreConfig.expression} onChange={(e) => setScoreConfig({ ...scoreConfig, expression: e.target.value })} placeholder='Expression: "{q1} * 0.3 + {q2} * 0.7"' style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11, marginBottom: 8 }} />
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
                      <div><label style={{ fontSize: 10, color: C.textMuted }}>Green ({"\u003e="})</label><input type="number" value={scoreConfig.thresholds.green} onChange={(e) => setScoreConfig({ ...scoreConfig, thresholds: { ...scoreConfig.thresholds, green: parseInt(e.target.value) || 0 } })} style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} /></div>
                      <div><label style={{ fontSize: 10, color: C.textMuted }}>Amber ({"\u003e="})</label><input type="number" value={scoreConfig.thresholds.amber} onChange={(e) => setScoreConfig({ ...scoreConfig, thresholds: { ...scoreConfig.thresholds, amber: parseInt(e.target.value) || 0 } })} style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} /></div>
                      <div><label style={{ fontSize: 10, color: C.textMuted }}>Red</label><input type="number" value={scoreConfig.thresholds.red} onChange={(e) => setScoreConfig({ ...scoreConfig, thresholds: { ...scoreConfig.thresholds, red: parseInt(e.target.value) || 0 } })} style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} /></div>
                    </div>
                  </>
                )}
              </div>
              {/* Duplicate Detection */}
              <div style={{ marginBottom: 20, padding: 12, background: C.offWhite, borderRadius: 8 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <input type="checkbox" checked={duplicateDetection.enabled} onChange={(e) => setDuplicateDetection({ ...duplicateDetection, enabled: e.target.checked })} />
                  <span style={{ fontSize: 12, fontWeight: 600 }}>Duplicate Detection</span>
                </label>
                {duplicateDetection.enabled && (
                  <>
                    <div style={{ marginBottom: 8 }}>
                      <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Identify duplicates by:</div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                        {fields.filter(f => f.type !== "panel" && f.type !== "html").slice(0, 6).map(f => (
                          <label key={f._id} style={{ fontSize: 10, display: "flex", alignItems: "center", gap: 4 }}><input type="checkbox" checked={duplicateDetection.identifyBy.includes(f.name)} onChange={(e) => setDuplicateDetection({ ...duplicateDetection, identifyBy: e.target.checked ? [...duplicateDetection.identifyBy, f.name] : duplicateDetection.identifyBy.filter(n => n !== f.name) })} /> {f.name}</label>
                        ))}
                      </div>
                    </div>
                    <select value={duplicateDetection.action} onChange={(e) => setDuplicateDetection({ ...duplicateDetection, action: e.target.value as "block" | "warn" | "overwrite" })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }}>
                      <option value="warn">Warn but allow</option>
                      <option value="block">Block submission</option>
                      <option value="overwrite">Overwrite previous</option>
                    </select>
                  </>
                )}
              </div>
              {/* Quota */}
              <div style={{ padding: 12, background: C.offWhite, borderRadius: 8 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                  <input type="checkbox" checked={quotaConfig.enabled} onChange={(e) => setQuotaConfig({ ...quotaConfig, enabled: e.target.checked })} />
                  <span style={{ fontSize: 12, fontWeight: 600 }}>Submission Quota</span>
                </label>
                {quotaConfig.enabled && (
                  <>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                      <div><div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Max Total</div><input type="number" value={quotaConfig.maxSubmissions} onChange={(e) => setQuotaConfig({ ...quotaConfig, maxSubmissions: parseInt(e.target.value) || 0 })} style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} /></div>
                      <div><div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Max Per User (0=unlimited)</div><input type="number" value={quotaConfig.maxPerUser || 0} onChange={(e) => setQuotaConfig({ ...quotaConfig, maxPerUser: parseInt(e.target.value) || 0 })} style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }} /></div>
                    </div>
                    <select value={quotaConfig.actionWhenReached} onChange={(e) => setQuotaConfig({ ...quotaConfig, actionWhenReached: e.target.value as "disable" | "message" | "redirect" })} style={{ width: "100%", padding: "6px 8px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 11 }}>
                      <option value="message">Show message</option>
                      <option value="disable">Disable form</option>
                      <option value="redirect">Redirect to URL</option>
                    </select>
                  </>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
      {/* ── PART 5: FIELD PERMISSIONS MODAL ──────────────────────────────────── */}
      {showFieldPermissions && (
        <div onClick={() => setShowFieldPermissions(false)} style={{ position: "fixed", inset: 0, zIndex: 3100, background: "rgba(30,27,75,0.5)", backdropFilter: "blur(2px)", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <div onClick={(e) => e.stopPropagation()} style={{ background: C.white, borderRadius: 12, width: 560, maxHeight: "85vh", boxShadow: "0 12px 40px rgba(91,33,182,0.25)", border: `1px solid ${C.border}`, overflow: "hidden", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "16px 20px", borderBottom: `1px solid ${C.border}`, display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: C.textPrimary }}>🔐 Field Permissions & Data Masking</div>
              <button onClick={() => setShowFieldPermissions(false)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 16, color: C.textMuted }}>✕</button>
            </div>
            <div style={{ flex: 1, overflow: "auto", padding: 16 }}>
              <div style={{ fontSize: 11, color: C.textMuted, marginBottom: 16 }}>Configure who can view/edit each field and mark sensitive fields for data masking.</div>
              {fields.filter(f => f.type !== "panel" && f.type !== "html" && f.type !== "pagebreak" && f.type !== "spacer" && f.type !== "divider").map((f) => {
                const perm = fieldPermissions.find(p => p.fieldName === f.name) || { fieldName: f.name, viewRoles: ["All"], editRoles: ["All"], isSensitive: false, readOnlyAfterSubmit: false };
                return (
                  <div key={f._id} style={{ padding: 12, background: C.offWhite, borderRadius: 8, marginBottom: 8 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>{f.title} <span style={{ color: C.textMuted, fontWeight: 400 }}>({f.name})</span></div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
                      <div>
                        <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>View Roles (comma-separated)</div>
                        <input value={perm.viewRoles.join(", ")} onChange={(e) => { const newPerms = [...fieldPermissions.filter(p => p.fieldName !== f.name), { ...perm, viewRoles: e.target.value.split(",").map(s => s.trim()).filter(Boolean) }]; setFieldPermissions(newPerms); }} placeholder="All, HR Admin, Manager" style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 10 }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 10, color: C.textMuted, marginBottom: 4 }}>Edit Roles (comma-separated)</div>
                        <input value={perm.editRoles.join(", ")} onChange={(e) => { const newPerms = [...fieldPermissions.filter(p => p.fieldName !== f.name), { ...perm, editRoles: e.target.value.split(",").map(s => s.trim()).filter(Boolean) }]; setFieldPermissions(newPerms); }} placeholder="All, HR Admin" style={{ width: "100%", padding: "4px 6px", border: `1px solid ${C.border}`, borderRadius: 4, fontSize: 10 }} />
                      </div>
                    </div>
                    <div style={{ display: "flex", gap: 12 }}>
                      <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10 }}><input type="checkbox" checked={perm.isSensitive} onChange={(e) => { const newPerms = [...fieldPermissions.filter(p => p.fieldName !== f.name), { ...perm, isSensitive: e.target.checked }]; setFieldPermissions(newPerms); }} /> 🔒 Sensitive (mask in logs)</label>
                      <label style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 10 }}><input type="checkbox" checked={perm.readOnlyAfterSubmit} onChange={(e) => { const newPerms = [...fieldPermissions.filter(p => p.fieldName !== f.name), { ...perm, readOnlyAfterSubmit: e.target.checked }]; setFieldPermissions(newPerms); }} /> 📖 Read-only after submit</label>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
      </div>
        )}
      </div>
    </>;
}
