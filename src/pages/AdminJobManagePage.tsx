import { useState, useEffect, useCallback, useRef } from "react";
import {
  Box,
  Typography,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Button,
  Chip,
  Skeleton,
  Alert,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Checkbox,
  FormControlLabel,
  IconButton,
  ToggleButton,
  Snackbar,
  Grid,
  Card,
  CardContent,
  Divider,
  Stack,
  Tooltip,
} from "@mui/material";
import {
  Add,
  Edit,
  Delete,
  Close,
  Refresh,
  Work,
  FormatBold,
  FormatItalic,
  FormatListBulleted,
  FormatListNumbered,
} from "@mui/icons-material";
import DOMPurify from "dompurify";
import { useMsal } from "@azure/msal-react";
import { fetchAdminJobs, createJobListing, updateJobListing, fetchColumnChoices } from "../utils/careersService";
import type { JobListing, CustomFieldDefinition } from "../types";

const EMPLOYMENT_TYPES = ["Full-Time", "Part-Time", "Contract", "Internship"];
const FIELD_TYPES: CustomFieldDefinition["type"][] = ["text", "textarea", "number", "choice", "date"];

const EMPTY_JOB = {
  title: "",
  jobDescription: "",
  department: "",
  location: "",
  employmentType: "",
  salaryMin: null as number | null,
  salaryMax: null as number | null,
  closingDate: "",
  status: "New",
};

function MiniFormBuilder({
  fields,
  onChange,
}: {
  fields: CustomFieldDefinition[];
  onChange: (fields: CustomFieldDefinition[]) => void;
}) {
  const [editIndex, setEditIndex] = useState<number | null>(null);
  const [editField, setEditField] = useState<CustomFieldDefinition>({
    name: "",
    label: "",
    type: "text",
    required: false,
  });
  const [choicesInput, setChoicesInput] = useState("");
  const [showForm, setShowForm] = useState(false);

  const resetForm = () => {
    setEditField({ name: "", label: "", type: "text", required: false });
    setChoicesInput("");
    setEditIndex(null);
    setShowForm(false);
  };

  const handleSave = () => {
    if (!editField.label.trim()) return;
    const name = editField.name || editField.label.toLowerCase().replace(/\s+/g, "_").replace(/[^a-z0-9_]/g, "");
    const choices = choicesInput.split(",").map((s) => s.trim()).filter(Boolean);
    const saved = { ...editField, name, choices: choices.length > 0 ? choices : undefined };
    if (editIndex !== null) {
      const next = [...fields];
      next[editIndex] = saved;
      onChange(next);
    } else {
      onChange([...fields, saved]);
    }
    resetForm();
  };

  const handleDelete = (index: number) => {
    onChange(fields.filter((_, i) => i !== index));
  };

  const handleEdit = (index: number) => {
    setEditIndex(index);
    setEditField({ ...fields[index] });
    setChoicesInput((fields[index].choices || []).join(", "));
    setShowForm(true);
  };

  return (
    <Box>
      <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between", mb: 1 }}>
        <Typography variant="subtitle2" sx={{ fontWeight: 600, color: "#374151" }}>
          Custom Application Questions
        </Typography>
        <Button
          size="small"
          startIcon={<Add />}
          onClick={() => { resetForm(); setShowForm(true); }}
          sx={{ borderRadius: "8px", textTransform: "none", fontWeight: 600 }}
        >
          Add Question
        </Button>
      </Box>

      {/* Question list */}
      {fields.length === 0 && !showForm && (
        <Typography variant="body2" sx={{ color: "#9CA3AF", py: 2, textAlign: "center" }}>
          No custom questions added yet.
        </Typography>
      )}

      <Stack spacing={1} sx={{ mb: 2 }}>
        {fields.map((field, i) => (
          <Paper
            key={i}
            variant="outlined"
            sx={{ display: "flex", alignItems: "center", gap: 1, px: 1.5, py: 1, borderRadius: "8px", borderColor: "#E5E7EB" }}
          >
            <Box sx={{ flex: 1 }}>
              <Typography variant="body2" sx={{ fontWeight: 600, color: "#111827", fontSize: "0.85rem" }}>
                {field.label}
              </Typography>
              <Box sx={{ display: "flex", gap: 0.5, mt: 0.25 }}>
                <Chip label={field.type} size="small" sx={{ height: 20, fontSize: "0.65rem", borderRadius: "4px", backgroundColor: "#F3F4F6", color: "#6B7280" }} />
                {field.required && <Chip label="Required" size="small" sx={{ height: 20, fontSize: "0.65rem", borderRadius: "4px", backgroundColor: "#FEF3C7", color: "#92400E" }} />}
              </Box>
            </Box>
            <IconButton size="small" onClick={() => handleEdit(i)} sx={{ color: "#6B7280" }}><Edit sx={{ fontSize: 16 }} /></IconButton>
            <IconButton size="small" onClick={() => handleDelete(i)} sx={{ color: "#DC2626" }}><Delete sx={{ fontSize: 16 }} /></IconButton>
          </Paper>
        ))}
      </Stack>

      {/* Add/Edit form */}
      <Dialog open={showForm} onClose={resetForm} maxWidth="xs" fullWidth slotProps={{ paper: { sx: { borderRadius: "16px" } } }}>
        <DialogTitle sx={{ pb: 1 }}>
          <Typography variant="h6" component="div" sx={{ fontWeight: 700, color: "#111827" }}>
            {editIndex !== null ? "Edit Question" : "Add Question"}
          </Typography>
        </DialogTitle>
        <DialogContent>
          <Stack spacing={2} sx={{ mt: 1 }}>
            <TextField
              label="Question Label"
              value={editField.label}
              onChange={(e) => setEditField({ ...editField, label: e.target.value, name: e.target.value.toLowerCase().replace(/\s+/g, "_").replace(/[^a-z0-9_]/g, "") })}
              fullWidth
              size="small"
              slotProps={{ input: { sx: { borderRadius: "8px" } } }}
            />
            <FormControl fullWidth size="small">
              <InputLabel>Type</InputLabel>
              <Select
                value={editField.type}
                label="Type"
                onChange={(e) => setEditField({ ...editField, type: e.target.value as CustomFieldDefinition["type"] })}
                sx={{ borderRadius: "8px" }}
              >
                {FIELD_TYPES.map((t) => <MenuItem key={t} value={t}>{t}</MenuItem>)}
              </Select>
            </FormControl>
            <FormControlLabel
              control={<Checkbox checked={editField.required} onChange={(e) => setEditField({ ...editField, required: e.target.checked })} />}
              label="Required"
            />
            {editField.type === "choice" && (
              <TextField
                label="Choices (separate by comma)"
                value={choicesInput}
                onChange={(e) => setChoicesInput(e.target.value)}
                fullWidth
                size="small"
                slotProps={{ input: { sx: { borderRadius: "8px" } } }}
              />
            )}
          </Stack>
        </DialogContent>
        <DialogActions sx={{ px: 3, pb: 2 }}>
          <Button onClick={resetForm} sx={{ borderRadius: "8px", textTransform: "none", color: "#6B7280" }}>Cancel</Button>
          <Button variant="contained" onClick={handleSave} disabled={!editField.label.trim()} sx={{ borderRadius: "8px", textTransform: "none", backgroundColor: "#0078D4" }}>
            {editIndex !== null ? "Update" : "Add"}
          </Button>
        </DialogActions>
      </Dialog>
    </Box>
  );
}

// ── Lightweight Rich Text Editor ──────────────────────────────────────────

function RichTextEditor({
  value,
  onChange,
  minHeight = 150,
}: {
  value: string;
  onChange: (html: string) => void;
  minHeight?: number;
}) {
  const editorRef = useRef<HTMLDivElement>(null);
  const [format, setFormat] = useState<string[]>([]);

  // Sync innerHTML when value changes externally
  useEffect(() => {
    if (editorRef.current && editorRef.current.innerHTML !== value) {
      editorRef.current.innerHTML = DOMPurify.sanitize(value);
    }
  }, [value]);

  const exec = (cmd: string, val?: string) => {
    document.execCommand(cmd, false, val);
    if (editorRef.current) {
      onChange(editorRef.current.innerHTML);
    }
    editorRef.current?.focus();
  };

  const handleMouseUp = () => {
    if (!editorRef.current) return;
    // Detect active formats
    const active: string[] = [];
    if (document.queryCommandState("bold")) active.push("bold");
    if (document.queryCommandState("italic")) active.push("italic");
    if (document.queryCommandState("insertUnorderedList")) active.push("ul");
    if (document.queryCommandState("insertOrderedList")) active.push("ol");
    setFormat(active);
  };

  const handleInput = () => {
    if (editorRef.current) {
      onChange(editorRef.current.innerHTML);
    }
  };

  const handlePaste = (e: React.ClipboardEvent) => {
    // Strip rich formatting on paste, keep only basic structure
    e.preventDefault();
    const text = e.clipboardData.getData("text/plain");
    document.execCommand("insertText", false, text);
  };

  return (
    <Box
      sx={{
        border: "1px solid #D1D5DB",
        borderRadius: "10px",
        overflow: "hidden",
        "&:focus-within": {
          borderColor: "#0078D4",
          boxShadow: "0 0 0 2px rgba(0, 120, 212, 0.15)",
        },
      }}
    >
      {/* Toolbar */}
      <Box
        sx={{
          display: "flex",
          gap: 0.5,
          p: 0.5,
          borderBottom: "1px solid #E5E7EB",
          backgroundColor: "#F9FAFB",
        }}
      >
        <Tooltip title="Bold">
          <ToggleButton
            value="bold"
            selected={format.includes("bold")}
            onMouseDown={(e) => { e.preventDefault(); exec("bold"); }}
            size="small"
            sx={{ border: "none", borderRadius: "6px", p: "4px 8px", minWidth: 32 }}
          >
            <FormatBold sx={{ fontSize: 18 }} />
          </ToggleButton>
        </Tooltip>
        <Tooltip title="Italic">
          <ToggleButton
            value="italic"
            selected={format.includes("italic")}
            onMouseDown={(e) => { e.preventDefault(); exec("italic"); }}
            size="small"
            sx={{ border: "none", borderRadius: "6px", p: "4px 8px", minWidth: 32 }}
          >
            <FormatItalic sx={{ fontSize: 18 }} />
          </ToggleButton>
        </Tooltip>
        <Box sx={{ width: 1, backgroundColor: "#E5E7EB", mx: 0.5 }} />
        <Tooltip title="Bullet List">
          <ToggleButton
            value="ul"
            selected={format.includes("ul")}
            onMouseDown={(e) => { e.preventDefault(); exec("insertUnorderedList"); }}
            size="small"
            sx={{ border: "none", borderRadius: "6px", p: "4px 8px", minWidth: 32 }}
          >
            <FormatListBulleted sx={{ fontSize: 18 }} />
          </ToggleButton>
        </Tooltip>
        <Tooltip title="Numbered List">
          <ToggleButton
            value="ol"
            selected={format.includes("ol")}
            onMouseDown={(e) => { e.preventDefault(); exec("insertOrderedList"); }}
            size="small"
            sx={{ border: "none", borderRadius: "6px", p: "4px 8px", minWidth: 32 }}
          >
            <FormatListNumbered sx={{ fontSize: 18 }} />
          </ToggleButton>
        </Tooltip>
      </Box>

      {/* Editor area */}
      <Box
        ref={editorRef}
        contentEditable
        suppressContentEditableWarning
        onInput={handleInput}
        onMouseUp={handleMouseUp}
        onKeyUp={handleMouseUp}
        onPaste={handlePaste}
        sx={{
          minHeight,
          p: 2,
          outline: "none",
          fontSize: "0.9rem",
          lineHeight: 1.7,
          color: "#374151",
          "&:empty:before": {
            content: '"Describe the job role, responsibilities, and requirements..."',
            color: "#9CA3AF",
            pointerEvents: "none",
          },
          "& ul, & ol": { pl: 3, mb: 1 },
          "& li": { mb: 0.5 },
          "& strong": { fontWeight: 600 },
          "& p": { mb: 1 },
        }}
      />
    </Box>
  );
}

function JobFormDialog({
  open,
  onClose,
  onSave,
  initial,
  departmentChoices,
  employmentTypeChoices,
}: {
  open: boolean;
  onClose: () => void;
  onSave: (data: Record<string, unknown>, customFields: CustomFieldDefinition[]) => Promise<void>;
  initial: (typeof EMPTY_JOB) & { customFields: CustomFieldDefinition[] } | null;
  departmentChoices: string[];
  employmentTypeChoices: string[];
}) {
  const isEdit = !!initial;
  const [title, setTitle] = useState(initial?.title ?? "");
  const [jobDescription, setJobDescription] = useState(initial?.jobDescription ?? "");
  const [department, setDepartment] = useState(initial?.department ?? "");
  const [location, setLocation] = useState(initial?.location ?? "");
  const [employmentType, setEmploymentType] = useState(initial?.employmentType ?? "");
  const [salaryMin, setSalaryMin] = useState(initial?.salaryMin ?? null);
  const [salaryMax, setSalaryMax] = useState(initial?.salaryMax ?? null);
  const [closingDate, setClosingDate] = useState(initial?.closingDate ?? "");
  const [status, setStatus] = useState(initial?.status ?? "New");
  const [customFields, setCustomFields] = useState<CustomFieldDefinition[]>(initial?.customFields ?? []);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    if (open && initial) {
      setTitle(initial.title);
      setJobDescription(initial.jobDescription);
      setDepartment(initial.department);
      setLocation(initial.location);
      setEmploymentType(initial.employmentType);
      setSalaryMin(initial.salaryMin);
      setSalaryMax(initial.salaryMax);
      setClosingDate(initial.closingDate);
      setStatus(initial.status);
      setCustomFields(initial.customFields);
    } else if (open) {
      setTitle(""); setJobDescription(""); setDepartment(""); setLocation("");
      setEmploymentType(""); setSalaryMin(null); setSalaryMax(null); setClosingDate(""); setStatus("New");
      setCustomFields([]);
    }
  }, [open, initial]);

  const handleSave = async () => {
    if (!title.trim()) return;
    setSaving(true);
    try {
      await onSave(
        {
          title: title.trim(),
          jobDescription,
          department: department.trim(),
          location: location.trim(),
          employmentType,
          salaryMin: salaryMin ?? 0,
          salaryMax: salaryMax ?? 0,
          closingDate: closingDate || null,
          status,
        },
        customFields,
      );
    } finally {
      setSaving(false);
    }
  };

  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth slotProps={{ paper: { sx: { borderRadius: "16px" } } }}>
      <DialogTitle sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", pb: 1 }}>
        <Typography variant="h6" component="div" sx={{ fontWeight: 700, color: "#111827" }}>
          {isEdit ? "Edit Job Listing" : "Create New Job"}
        </Typography>
        <IconButton onClick={onClose} size="small"><Close /></IconButton>
      </DialogTitle>
      <DialogContent>
        <Grid container spacing={2} sx={{ mt: 0.5 }}>
          <Grid size={{ xs: 12, md: 6 }}>
            <TextField label="Job Title" value={title} onChange={(e) => setTitle(e.target.value)} fullWidth required size="small" slotProps={{ input: { sx: { borderRadius: "8px" } } }} />
          </Grid>
          <Grid size={{ xs: 6, md: 3 }}>
            <FormControl fullWidth size="small">
              <InputLabel>Department</InputLabel>
              <Select value={department} label="Department" onChange={(e) => setDepartment(e.target.value)} sx={{ borderRadius: "8px" }}>
                {departmentChoices.length > 0
                  ? departmentChoices.map((d) => <MenuItem key={d} value={d}>{d}</MenuItem>)
                  : <MenuItem value="">No choices loaded</MenuItem>
                }
              </Select>
            </FormControl>
          </Grid>
          <Grid size={{ xs: 6, md: 3 }}>
            <TextField label="Location" value={location} onChange={(e) => setLocation(e.target.value)} fullWidth size="small" slotProps={{ input: { sx: { borderRadius: "8px" } } }} />
          </Grid>
          <Grid size={{ xs: 12 }}>
            <Typography variant="body2" sx={{ fontWeight: 600, color: "#374151", mb: 0.5 }}>
              Job Description
            </Typography>
            <RichTextEditor value={jobDescription} onChange={setJobDescription} minHeight={180} />
          </Grid>
          <Grid size={{ xs: 6, md: 3 }}>
            <FormControl fullWidth size="small">
              <InputLabel>Employment Type</InputLabel>
              <Select value={employmentType} label="Employment Type" onChange={(e) => setEmploymentType(e.target.value)} sx={{ borderRadius: "8px" }}>
                {(employmentTypeChoices.length > 0 ? employmentTypeChoices : EMPLOYMENT_TYPES).map((t) => (
                  <MenuItem key={t} value={t}>{t}</MenuItem>
                ))}
              </Select>
            </FormControl>
          </Grid>
          <Grid size={{ xs: 3, md: 2 }}>
            <TextField label="Salary Min" type="number" value={salaryMin ?? ""} onChange={(e) => setSalaryMin(e.target.value ? Number(e.target.value) : null)} fullWidth size="small" slotProps={{ input: { sx: { borderRadius: "8px" } } }} />
          </Grid>
          <Grid size={{ xs: 3, md: 2 }}>
            <TextField label="Salary Max" type="number" value={salaryMax ?? ""} onChange={(e) => setSalaryMax(e.target.value ? Number(e.target.value) : null)} fullWidth size="small" slotProps={{ input: { sx: { borderRadius: "8px" } } }} />
          </Grid>
          <Grid size={{ xs: 6, md: 2 }}>
            <TextField label="Closing Date" type="date" value={closingDate} onChange={(e) => setClosingDate(e.target.value)} fullWidth size="small" slotProps={{ inputLabel: { shrink: true }, input: { sx: { borderRadius: "8px" } } }} />
          </Grid>
          {isEdit && (
            <Grid size={{ xs: 6, md: 3 }}>
              <FormControl fullWidth size="small">
                <InputLabel>Status</InputLabel>
                <Select value={status} label="Status" onChange={(e) => setStatus(e.target.value)} sx={{ borderRadius: "8px" }}>
                  <MenuItem value="New">New (Active)</MenuItem>
                  <MenuItem value="Closed">Closed</MenuItem>
                </Select>
              </FormControl>
            </Grid>
          )}
        </Grid>

        <Divider sx={{ my: 2.5 }} />

        <MiniFormBuilder fields={customFields} onChange={setCustomFields} />
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 2, gap: 1 }}>
        <Button onClick={onClose} sx={{ borderRadius: "8px", textTransform: "none", color: "#6B7280" }}>Cancel</Button>
        <Button variant="contained" onClick={handleSave} disabled={!title.trim() || saving} sx={{ borderRadius: "8px", textTransform: "none", backgroundColor: "#0078D4" }}>
          {saving ? "Saving..." : isEdit ? "Update Job" : "Create Job"}
        </Button>
      </DialogActions>
    </Dialog>
  );
}

export default function AdminJobManagePage() {
  const { instance, accounts } = useMsal();
  const [jobs, setJobs] = useState<JobListing[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editJob, setEditJob] = useState<(typeof EMPTY_JOB) & { customFields: CustomFieldDefinition[] } | null>(null);
  const [snackbar, setSnackbar] = useState<{ message: string; severity: "success" | "error" | "warning" } | null>(null);
  const [departmentChoices, setDepartmentChoices] = useState<string[]>([]);
  const [employmentTypeChoices, setEmploymentTypeChoices] = useState<string[]>([]);

  /** Create the CustomFields column on the job list via SharePoint REST (client-side token). */
  async function ensureCustomFieldsColumn(): Promise<boolean> {
    try {
      const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");
      const tokenRes = await instance.acquireTokenSilent({
        scopes: [`${new URL(SP_SITE_URL).origin}/AllSites.Manage`],
        account: accounts[0],
      });
      const token = tokenRes.accessToken;

      // Get request digest
      const digestResp = await fetch(`${SP_SITE_URL}/_api/contextinfo`, {
        method: "POST",
        headers: { Accept: "application/json;odata=nometadata", Authorization: `Bearer ${token}` },
      });
      if (!digestResp.ok) return false;
      const digestData = await digestResp.json();
      const digest: string = digestData.FormDigestValue;

      // Try to add the CustomFields Note column — 400 likely means column already exists
      const resp = await fetch(`${SP_SITE_URL}/_api/web/lists/getbytitle('Internal Job Listing')/fields`, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=nometadata",
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": digest,
        },
        body: JSON.stringify({
          __metadata: { type: "SP.FieldMultiLineText" },
          FieldTypeKind: 3,
          Title: "CustomFields",
          StaticName: "CustomFields",
        }),
      });
      if (!resp.ok) {
        const text = await resp.text();
        // Column already exists — not an error
        if (text.toLowerCase().includes("duplicate") || text.toLowerCase().includes("already exists")) return true;
        console.warn("[ensureCustomFieldsColumn]", resp.status, text);
      }
      return true;
    } catch {
      return false;
    }
  }

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await fetchAdminJobs();
      setJobs(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load jobs");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { void load(); }, [load]);

  useEffect(() => {
    async function loadChoices() {
      const [dept, emp] = await Promise.all([
        fetchColumnChoices("Internal Job Listing", "Department"),
        fetchColumnChoices("Internal Job Listing", "Employment Type"),
      ]);
      setDepartmentChoices(dept);
      setEmploymentTypeChoices(emp);
    }
    void loadChoices();
  }, []);

  const handleCreate = () => {
    setEditJob(null);
    setDialogOpen(true);
  };

  const handleEdit = (job: JobListing) => {
    setEditJob({
      title: job.title,
      jobDescription: job.jobDescription,
      department: job.department,
      location: job.location,
      employmentType: job.employmentType,
      salaryMin: job.salaryMin,
      salaryMax: job.salaryMax,
      closingDate: job.closingDate ?? "",
      status: job.status,
      customFields: job.customFields ?? [],
    });
    setDialogOpen(true);
  };

  const handleSave = async (
    data: Record<string, unknown>,
    customFields: CustomFieldDefinition[],
  ) => {
    try {
      // If custom fields are present, ensure the column exists first
      const hasCustom = Array.isArray(customFields) && customFields.length > 0;
      if (hasCustom) {
        await ensureCustomFieldsColumn();
      }

      if (editJob) {
        // Update existing
        const result = await updateJobListing(
          jobs.find((j) => j.title === editJob.title)?.id || "",
          { ...data, customFields },
        );
        if (result.success) {
          setSnackbar({
            message: result.warning || "Job updated",
            severity: result.warning ? "warning" : "success",
          });
        }
      } else {
        const result = await createJobListing({ ...data, customFields });
        if (result.success) {
          setSnackbar({
            message: (result as { warning?: string }).warning || "Job created",
            severity: (result as { warning?: string }).warning ? "warning" : "success",
          });
        }
      }
      setDialogOpen(false);
      void load();
    } catch (err) {
      setSnackbar({
        message: err instanceof Error ? err.message : "Operation failed",
        severity: "error",
      });
    }
  };

  const handleClose = async (job: JobListing) => {
    try {
      const result = await updateJobListing(job.id, { status: "Closed" });
      if (result.success) {
        setSnackbar({ message: "Job closed", severity: "success" });
        void load();
      }
    } catch (err) {
      setSnackbar({
        message: err instanceof Error ? err.message : "Failed to close job",
        severity: "error",
      });
    }
  };

  return (
    <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
      {/* Header */}
      <Paper sx={{ borderRadius: 0, boxShadow: "0 1px 3px rgba(0,0,0,0.06)", backgroundColor: "#ffffff", position: "sticky", top: 0, zIndex: 10 }}>
        <Box sx={{ maxWidth: 1280, mx: "auto", px: { xs: 2, sm: 3, md: 4 }, py: 2.5 }}>
          <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <Box>
              <Typography variant="h5" sx={{ fontWeight: 700, color: "#111827", fontSize: "1.3rem" }}>
                Manage Job Listings
              </Typography>
              <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.85rem" }}>
                Create and manage internal job postings
              </Typography>
            </Box>
            <Box sx={{ display: "flex", gap: 1 }}>
              <Button variant="outlined" startIcon={<Refresh />} onClick={load} disabled={loading} sx={{ borderRadius: "10px", textTransform: "none", fontWeight: 600, borderColor: "#D1D5DB", color: "#6B7280" }}>
                Refresh
              </Button>
              <Button variant="contained" startIcon={<Add />} onClick={handleCreate} sx={{ borderRadius: "10px", textTransform: "none", fontWeight: 600, backgroundColor: "#0078D4" }}>
                Create New Job
              </Button>
            </Box>
          </Box>
        </Box>
      </Paper>

      <Box sx={{ maxWidth: 1280, mx: "auto", px: { xs: 2, sm: 3, md: 4 }, py: 3 }}>
        {/* Stats */}
        <Grid container spacing={2} sx={{ mb: 3 }}>
          {[
            { label: "Total Jobs", value: jobs.length, color: "#0078D4", icon: <Work /> },
            { label: "Active", value: jobs.filter((j) => j.status === "New").length, color: "#34A853", icon: <Work /> },
            { label: "Closed", value: jobs.filter((j) => j.status !== "New").length, color: "#9CA3AF", icon: <Work /> },
          ].map((stat) => (
            <Grid size={{ xs: 4 }} key={stat.label}>
              <Card sx={{ borderRadius: "16px", boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
                <CardContent sx={{ p: 2.5, textAlign: "center" }}>
                  <Typography variant="h4" sx={{ fontWeight: 700, color: stat.color }}>{stat.value}</Typography>
                  <Typography variant="caption" sx={{ color: "#6B7280", fontWeight: 500 }}>{stat.label}</Typography>
                </CardContent>
              </Card>
            </Grid>
          ))}
        </Grid>

        {/* Loading */}
        {loading && (
          <TableContainer component={Paper} sx={{ borderRadius: "16px", boxShadow: "0 1px 3px rgba(0,0,0,0.06)" }}>
            <Table>
              <TableHead>
                <TableRow sx={{ backgroundColor: "#F9FAFB" }}>
                  {["Title", "Department", "Type", "Salary", "Status", "Applicants", "Actions"].map((h) => (
                    <TableCell key={h} sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>{h}</TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {[1, 2, 3, 4].map((i) => (
                  <TableRow key={i}>
                    <TableCell><Skeleton variant="text" width={140} /></TableCell>
                    <TableCell><Skeleton variant="rounded" width={80} height={24} sx={{ borderRadius: "8px" }} /></TableCell>
                    <TableCell><Skeleton variant="text" width={90} /></TableCell>
                    <TableCell><Skeleton variant="text" width={70} /></TableCell>
                    <TableCell><Skeleton variant="rounded" width={60} height={24} sx={{ borderRadius: "8px" }} /></TableCell>
                    <TableCell><Skeleton variant="text" width={30} /></TableCell>
                    <TableCell><Skeleton variant="text" width={60} /></TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        )}

        {/* Error */}
        {!loading && error && <Alert severity="error" sx={{ borderRadius: "12px", mb: 3, fontWeight: 500, backgroundColor: "#FEF2F2", color: "#991B1B", "& .MuiAlert-icon": { color: "#DC2626" } }} action={<Button size="small" onClick={load} sx={{ textTransform: "none" }}>Retry</Button>}>{error}</Alert>}

        {/* Empty */}
        {!loading && !error && jobs.length === 0 && (
          <Box sx={{ textAlign: "center", py: 8 }}>
            <Work sx={{ fontSize: 48, color: "#D1D5DB", mb: 2 }} />
            <Typography variant="h6" sx={{ color: "#6B7280", fontWeight: 600 }}>No Job Listings</Typography>
            <Typography variant="body2" sx={{ color: "#9CA3AF", mb: 2 }}>Create your first job posting.</Typography>
            <Button variant="contained" startIcon={<Add />} onClick={handleCreate} sx={{ borderRadius: "10px", textTransform: "none", backgroundColor: "#0078D4" }}>Create New Job</Button>
          </Box>
        )}

        {/* Table */}
        {!loading && !error && jobs.length > 0 && (
          <TableContainer component={Paper} sx={{ borderRadius: "16px", boxShadow: "0 1px 3px rgba(0,0,0,0.06)", overflow: "hidden" }}>
            <Table>
              <TableHead>
                <TableRow sx={{ backgroundColor: "#F9FAFB" }}>
                  {["Title", "Department", "Type", "Salary", "Status", "Applicants", "Actions"].map((h) => (
                    <TableCell key={h} sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>{h}</TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {jobs.map((job) => (
                  <TableRow key={job.id} hover sx={{ "&:hover": { backgroundColor: "#FAFBFC" } }}>
                    <TableCell>
                      <Typography variant="body2" sx={{ fontWeight: 600, color: "#111827", fontSize: "0.85rem" }}>{job.title}</Typography>
                      {job.location && <Typography variant="caption" sx={{ color: "#9CA3AF" }}>{job.location}</Typography>}
                    </TableCell>
                    <TableCell><Chip label={job.department} size="small" sx={{ borderRadius: "8px", fontSize: "0.7rem", backgroundColor: "#6264A7", color: "#fff" }} /></TableCell>
                    <TableCell><Typography variant="body2" sx={{ color: "#374151", fontSize: "0.8rem" }}>{job.employmentType}</Typography></TableCell>
                    <TableCell><Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>{job.salaryMin ? `RM ${job.salaryMin}` : "-"}</Typography></TableCell>
                    <TableCell>
                      <Chip label={job.status === "New" ? "Active" : "Closed"} size="small" sx={{ borderRadius: "8px", fontSize: "0.7rem", backgroundColor: job.status === "New" ? "#E6F4EA" : "#F3F4F6", color: job.status === "New" ? "#34A853" : "#6B7280", fontWeight: 600 }} />
                    </TableCell>
                    <TableCell><Typography variant="body2" sx={{ fontWeight: 600, color: "#0078D4", fontSize: "0.85rem" }}>{job.applicationCount}</Typography></TableCell>
                    <TableCell>
                      <Box sx={{ display: "flex", gap: 0.5 }}>
                        <IconButton size="small" onClick={() => handleEdit(job)} sx={{ color: "#6B7280" }}><Edit sx={{ fontSize: 18 }} /></IconButton>
                        {job.status === "New" && (
                          <IconButton size="small" onClick={() => handleClose(job)} sx={{ color: "#DC2626" }}><Delete sx={{ fontSize: 18 }} /></IconButton>
                        )}
                      </Box>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        )}
      </Box>

      {/* Create/Edit Dialog */}
      <JobFormDialog
        open={dialogOpen}
        onClose={() => setDialogOpen(false)}
        onSave={handleSave}
        initial={editJob}
        departmentChoices={departmentChoices}
        employmentTypeChoices={employmentTypeChoices}
      />

      {/* Snackbar */}
      <Snackbar open={!!snackbar} autoHideDuration={4000} onClose={() => setSnackbar(null)} anchorOrigin={{ vertical: "bottom", horizontal: "center" }}>
        {snackbar ? (
          <Alert
            severity={snackbar.severity}
            onClose={() => setSnackbar(null)}
            sx={{
              borderRadius: "10px",
              fontWeight: 600,
              fontSize: "0.9rem",
              boxShadow: "0 6px 20px rgba(0,0,0,0.15)",
              "& .MuiAlert-icon": { fontSize: 22, alignSelf: "center" },
            }}
          >
            {snackbar.message}
          </Alert>
        ) : undefined}
      </Snackbar>
    </Box>
  );
}
