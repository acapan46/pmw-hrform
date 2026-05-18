import { useState, useEffect, useCallback } from "react";
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
  Chip,
  Select,
  MenuItem,
  Skeleton,
  Alert,
  Button,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Snackbar,
  Grid,
  Card,
  CardContent,
  IconButton,
  Checkbox,
  LinearProgress,
} from "@mui/material";
import {
  Close,
  Refresh,
  People,
  NewReleases,
  CheckCircle,
  Delete as DeleteIcon,
} from "@mui/icons-material";
import { fetchApplications, updateApplicationStatus, deleteApplications } from "../utils/careersService";
import type { JobAdminApplication } from "../types";

const STATUS_OPTIONS = ["New", "Reviewed"] as const;

const STATUS_COLORS: Record<string, string> = {
  New: "#0078D4",
  Reviewed: "#34A853",
};

function StatusChip({ status }: { status: string }) {
  const color = STATUS_COLORS[status] || "#6B7280";
  return (
    <Chip
      label={status}
      size="small"
      sx={{
        backgroundColor: `${color}18`,
        color,
        fontWeight: 600,
        fontSize: "0.75rem",
        borderRadius: "8px",
      }}
    />
  );
}

function formatDate(dateStr: string): string {
  try {
    const d = new Date(dateStr);
    return d.toLocaleDateString("en-MY", {
      day: "numeric",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  } catch {
    return dateStr;
  }
}

export default function AdminJobsPage() {
  const [applications, setApplications] = useState<JobAdminApplication[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [selectedApp, setSelectedApp] = useState<JobAdminApplication | null>(null);
  const [snackbar, setSnackbar] = useState<{ message: string; severity: "success" | "error" } | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [confirmDeleteOpen, setConfirmDeleteOpen] = useState(false);
  const [deleting, setDeleting] = useState(false);
  const [deleteResult, setDeleteResult] = useState<string | null>(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await fetchApplications();
      setApplications(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load applications");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void load();
  }, [load]);

  const handleStatusChange = useCallback(
    async (applicationId: string, newStatus: string) => {
      try {
        const success = await updateApplicationStatus(applicationId, newStatus);
        if (success) {
          setApplications((prev) =>
            prev.map((app) => (app.id === applicationId ? { ...app, status: newStatus } : app)),
          );
          setSnackbar({ message: "Status updated successfully", severity: "success" });
        }
      } catch (err) {
        setSnackbar({
          message: err instanceof Error ? err.message : "Failed to update status",
          severity: "error",
        });
      }
    },
    [],
  );

  const allSelected = applications.length > 0 && selectedIds.size === applications.length;

  const toggleSelect = (id: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (allSelected) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(applications.map((a) => a.id)));
    }
  };

  const handleDelete = async () => {
    setDeleting(true);
    setDeleteResult(null);
    try {
      const result = await deleteApplications([...selectedIds]);
      const msg = `Deleted ${result.deleted} application${result.deleted !== 1 ? "s" : ""}`;
      if (result.errors && result.errors.length > 0) {
        setSnackbar({ message: `${msg}. Errors: ${result.errors.join("; ")}`, severity: "error" });
      } else {
        setSnackbar({ message: msg, severity: "success" });
      }
      setSelectedIds(new Set());
      setConfirmDeleteOpen(false);
      void load();
    } catch (err) {
      setSnackbar({
        message: err instanceof Error ? err.message : "Failed to delete applications",
        severity: "error",
      });
    } finally {
      setDeleting(false);
    }
  };

  const stats = {
    total: applications.length,
    new: applications.filter((a) => a.status === "New").length,
    reviewed: applications.filter((a) => a.status === "Reviewed").length,
  };

  return (
    <Box sx={{ minHeight: "100vh", backgroundColor: "#F8F9FC" }}>
      {/* Header */}
      <Paper
        sx={{
          borderRadius: 0,
          boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
          backgroundColor: "#ffffff",
          position: "sticky",
          top: 0,
          zIndex: 10,
        }}
      >
        <Box sx={{ maxWidth: 1280, mx: "auto", px: { xs: 2, sm: 3, md: 4 }, py: 2.5 }}>
          <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <Box>
              <Typography variant="h5" sx={{ fontWeight: 700, color: "#111827", fontSize: "1.3rem" }}>
                Job Applications
              </Typography>
              <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.85rem" }}>
                Manage incoming applications
              </Typography>
            </Box>
            <Button
              variant="outlined"
              startIcon={<Refresh />}
              onClick={load}
              disabled={loading}
              sx={{
                borderRadius: "10px",
                textTransform: "none",
                fontWeight: 600,
                borderColor: "#D1D5DB",
                color: "#6B7280",
              }}
            >
              Refresh
            </Button>
          </Box>
        </Box>
      </Paper>

      <Box sx={{ maxWidth: 1280, mx: "auto", px: { xs: 2, sm: 3, md: 4 }, py: 3 }}>
        {/* Stats Row */}
        <Grid container spacing={2} sx={{ mb: 3 }}>
          {[
            { label: "Total Applications", value: stats.total, icon: <People />, color: "#0078D4" },
            { label: "New", value: stats.new, icon: <NewReleases />, color: "#F59E0B" },
            { label: "Reviewed", value: stats.reviewed, icon: <CheckCircle />, color: "#34A853" },
          ].map((stat) => (
            <Grid size={{ xs: 6, sm: 4 }} key={stat.label}>
              <Card
                sx={{
                  borderRadius: "16px",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
                  transition: "transform 0.2s",
                  "&:hover": { transform: "translateY(-2px)" },
                }}
              >
                <CardContent sx={{ p: 2.5 }}>
                  <Box sx={{ display: "flex", alignItems: "center", gap: 1.5, mb: 0.5 }}>
                    <Box sx={{ color: stat.color, display: "flex" }}>{stat.icon}</Box>
                    <Typography variant="h5" sx={{ fontWeight: 700, color: "#111827" }}>
                      {stat.value}
                    </Typography>
                  </Box>
                  <Typography variant="caption" sx={{ color: "#6B7280", fontWeight: 500 }}>
                    {stat.label}
                  </Typography>
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
                  {["Reference", "Applicant", "Job Title", "Status", "Submitted", "Actions"].map((h) => (
                    <TableCell key={h} sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>{h}</TableCell>
                  ))}
                </TableRow>
              </TableHead>
              <TableBody>
                {[1, 2, 3, 4].map((i) => (
                  <TableRow key={i}>
                    <TableCell><Skeleton variant="text" width={110} /></TableCell>
                    <TableCell>
                      <Skeleton variant="text" width={130} />
                      <Skeleton variant="text" width={160} height={14} />
                    </TableCell>
                    <TableCell><Skeleton variant="text" width={120} /></TableCell>
                    <TableCell><Skeleton variant="rounded" width={70} height={24} sx={{ borderRadius: "8px" }} /></TableCell>
                    <TableCell><Skeleton variant="text" width={100} /></TableCell>
                    <TableCell><Skeleton variant="text" width={50} /></TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        )}

        {/* Error */}
        {!loading && error && (
          <Alert
            severity="error"
            sx={{ borderRadius: "12px", mb: 3, fontWeight: 500, backgroundColor: "#FEF2F2", color: "#991B1B", "& .MuiAlert-icon": { color: "#DC2626" } }}
            action={
              <Button size="small" onClick={load} sx={{ textTransform: "none" }}>
                Retry
              </Button>
            }
          >
            {error}
          </Alert>
        )}

        {/* Empty */}
        {!loading && !error && applications.length === 0 && (
          <Box sx={{ textAlign: "center", py: 8 }}>
            <People sx={{ fontSize: 48, color: "#D1D5DB", mb: 2 }} />
            <Typography variant="h6" sx={{ color: "#6B7280", fontWeight: 600 }}>
              No Applications Yet
            </Typography>
            <Typography variant="body2" sx={{ color: "#9CA3AF" }}>
              Applications from job postings will appear here.
            </Typography>
          </Box>
        )}

        {/* Delete bar */}
        {selectedIds.size > 0 && (
          <Paper
            sx={{
              mb: 2,
              p: 1.5,
              borderRadius: "12px",
              display: "flex",
              alignItems: "center",
              gap: 2,
              backgroundColor: "#FEF2F2",
              border: "1px solid #FECACA",
            }}
          >
            <Typography variant="body2" sx={{ color: "#991B1B", fontWeight: 600, flex: 1 }}>
              {selectedIds.size} application{selectedIds.size !== 1 ? "s" : ""} selected
            </Typography>
            <Button
              variant="contained"
              color="error"
              size="small"
              startIcon={<DeleteIcon />}
              onClick={() => setConfirmDeleteOpen(true)}
              sx={{ borderRadius: "8px", textTransform: "none", fontWeight: 600 }}
            >
              Delete
            </Button>
            <Button
              size="small"
              onClick={() => setSelectedIds(new Set())}
              sx={{ borderRadius: "8px", textTransform: "none", color: "#6B7280", fontWeight: 500 }}
            >
              Clear
            </Button>
          </Paper>
        )}

        {/* Table */}
        {!loading && !error && applications.length > 0 && (
          <TableContainer
            component={Paper}
            sx={{
              borderRadius: "16px",
              boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
              overflow: "hidden",
            }}
          >
            <Table>
              <TableHead>
                <TableRow sx={{ backgroundColor: "#F9FAFB" }}>
                  <TableCell padding="checkbox">
                    <Checkbox
                      checked={allSelected}
                      indeterminate={selectedIds.size > 0 && !allSelected}
                      onChange={toggleSelectAll}
                      sx={{ color: "#D1D5DB", "&.Mui-checked": { color: "#0078D4" } }}
                    />
                  </TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>
                    Reference
                  </TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>
                    Applicant
                  </TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>
                    Job Title
                  </TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>
                    Status
                  </TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>
                    Submitted
                  </TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>
                    Actions
                  </TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {applications.map((app) => (
                  <TableRow
                    key={app.id}
                    hover
                    selected={selectedIds.has(app.id)}
                    sx={{ cursor: "pointer", "&:hover": { backgroundColor: "#FAFBFC" } }}
                    onClick={() => setSelectedApp(app)}
                  >
                    <TableCell padding="checkbox" onClick={(e) => e.stopPropagation()}>
                      <Checkbox
                        checked={selectedIds.has(app.id)}
                        onChange={() => toggleSelect(app.id)}
                        sx={{ color: "#D1D5DB", "&.Mui-checked": { color: "#0078D4" } }}
                      />
                    </TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ fontFamily: "monospace", fontWeight: 600, color: "#0078D4", fontSize: "0.8rem" }}>
                        {app.submissionRef}
                      </Typography>
                    </TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ fontWeight: 600, color: "#111827", fontSize: "0.85rem" }}>
                        {app.applicantName}
                      </Typography>
                      <Typography variant="caption" sx={{ color: "#9CA3AF" }}>
                        {app.applicantEmail}
                      </Typography>
                    </TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ color: "#374151", fontSize: "0.85rem" }}>
                        {app.jobTitle}
                      </Typography>
                    </TableCell>
                    <TableCell>
                      <StatusChip status={app.status} />
                    </TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>
                        {formatDate(app.submittedAt)}
                      </Typography>
                    </TableCell>
                    <TableCell onClick={(e) => e.stopPropagation()}>
                      <Select
                        value={app.status}
                        size="small"
                        onChange={(e) => handleStatusChange(app.id, e.target.value)}
                        sx={{
                          borderRadius: "8px",
                          fontSize: "0.8rem",
                          minWidth: 120,
                          "& .MuiOutlinedInput-notchedOutline": { borderColor: "#E5E7EB" },
                        }}
                      >
                        {STATUS_OPTIONS.map((opt) => (
                          <MenuItem key={opt} value={opt} sx={{ fontSize: "0.85rem" }}>
                            {opt}
                          </MenuItem>
                        ))}
                      </Select>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </TableContainer>
        )}

        {/* Detail Dialog */}
        <Dialog
          open={!!selectedApp}
          onClose={() => setSelectedApp(null)}
          maxWidth="sm"
          fullWidth
          slotProps={{
            paper: {
              sx: { borderRadius: "16px", p: 1 },
            },
          }}
        >
          {selectedApp && (
            <>
              <DialogTitle sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", pb: 1 }}>
                <Typography variant="h6" component="div" sx={{ fontWeight: 700, color: "#111827" }}>
                  Application Details
                </Typography>
                <IconButton onClick={() => setSelectedApp(null)} size="small">
                  <Close />
                </IconButton>
              </DialogTitle>
              <DialogContent>
                <Box sx={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <Box>
                    <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                      Reference
                    </Typography>
                    <Typography variant="body2" sx={{ fontFamily: "monospace", fontWeight: 600, color: "#0078D4" }}>
                      {selectedApp.submissionRef}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                      Applicant
                    </Typography>
                    <Typography variant="body1" sx={{ fontWeight: 600, color: "#111827" }}>
                      {selectedApp.applicantName}
                    </Typography>
                    <Typography variant="body2" sx={{ color: "#6B7280" }}>
                      {selectedApp.applicantEmail}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                      Position
                    </Typography>
                    <Typography variant="body1" sx={{ color: "#374151" }}>
                      {selectedApp.jobTitle}
                    </Typography>
                  </Box>
                  <Box>
                    <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                      Status
                    </Typography>
                    <Box sx={{ mt: 0.5 }}>
                      <StatusChip status={selectedApp.status} />
                    </Box>
                  </Box>
                  <Box>
                    <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                      Submitted
                    </Typography>
                    <Typography variant="body2" sx={{ color: "#6B7280" }}>
                      {formatDate(selectedApp.submittedAt)}
                    </Typography>
                  </Box>

                  {(selectedApp.resumeUrl || selectedApp.coverLetterUrl) && (
                    <Box>
                      <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                        Documents
                      </Typography>
                      <Box sx={{ display: "flex", flexDirection: "column", gap: 0.5, mt: 0.5 }}>
                        {selectedApp.resumeUrl && (
                          <Box
                            component="a"
                            href={selectedApp.resumeUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            sx={{
                              display: "inline-flex", alignItems: "center", gap: 1,
                              px: 1.5, py: 0.75, borderRadius: "8px",
                              color: "#0078D4", fontWeight: 600, fontSize: "0.85rem",
                              backgroundColor: "#F0F7FF", border: "1px solid rgba(0,120,212,0.15)",
                              textDecoration: "none", width: "fit-content",
                              "&:hover": { backgroundColor: "#DBEAFE" },
                              "&::before": { content: "'📄 '", fontSize: "14px" },
                            }}
                          >
                            View Resume
                          </Box>
                        )}
                        {selectedApp.coverLetterUrl && (
                          <Box
                            component="a"
                            href={selectedApp.coverLetterUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            sx={{
                              display: "inline-flex", alignItems: "center", gap: 1,
                              px: 1.5, py: 0.75, borderRadius: "8px",
                              color: "#0078D4", fontWeight: 600, fontSize: "0.85rem",
                              backgroundColor: "#F0F7FF", border: "1px solid rgba(0,120,212,0.15)",
                              textDecoration: "none", width: "fit-content",
                              "&:hover": { backgroundColor: "#DBEAFE" },
                              "&::before": { content: "'📝 '", fontSize: "14px" },
                            }}
                          >
                            View Cover Letter
                          </Box>
                        )}
                      </Box>
                    </Box>
                  )}

                  {selectedApp.customAnswers && Object.keys(selectedApp.customAnswers).length > 0 && (
                    <Box>
                      <Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>
                        Additional Responses
                      </Typography>
                      <Box sx={{ display: "flex", flexDirection: "column", gap: 1, mt: 0.5 }}>
                        {Object.entries(selectedApp.customAnswers).map(([key, value]) => (
                          <Box key={key}>
                            <Typography variant="caption" sx={{ color: "#6B7280", fontWeight: 600, display: "block" }}>
                              {key}
                            </Typography>
                            <Typography variant="body2" sx={{ color: "#374151" }}>
                              {String(value ?? "")}
                            </Typography>
                          </Box>
                        ))}
                      </Box>
                    </Box>
                  )}
                </Box>
              </DialogContent>
              <DialogActions sx={{ px: 3, pb: 2 }}>
                <Button
                  variant="outlined"
                  onClick={() => setSelectedApp(null)}
                  sx={{ borderRadius: "10px", textTransform: "none", borderColor: "#D1D5DB", color: "#6B7280" }}
                >
                  Close
                </Button>
              </DialogActions>
            </>
          )}
        </Dialog>

        {/* Delete confirmation dialog */}
        <Dialog open={confirmDeleteOpen} onClose={() => !deleting && setConfirmDeleteOpen(false)} maxWidth="xs" fullWidth slotProps={{ paper: { sx: { borderRadius: "16px" } } }}>
          <DialogTitle sx={{ pb: 1 }}>
            <Typography variant="h6" component="div" sx={{ fontWeight: 700, color: "#111827" }}>
              Delete Applications
            </Typography>
          </DialogTitle>
          <DialogContent>
            <Typography variant="body2" sx={{ color: "#6B7280" }}>
              Are you sure you want to delete <strong>{selectedIds.size}</strong> application{selectedIds.size !== 1 ? "s" : ""}? This action cannot be undone.
            </Typography>
            {deleting && <LinearProgress sx={{ mt: 2, borderRadius: "4px" }} />}
            {deleteResult && (
              <Alert severity="info" sx={{ mt: 2, borderRadius: "8px" }}>{deleteResult}</Alert>
            )}
          </DialogContent>
          <DialogActions sx={{ px: 3, pb: 2, gap: 1 }}>
            <Button
              onClick={() => setConfirmDeleteOpen(false)}
              disabled={deleting}
              sx={{ borderRadius: "8px", textTransform: "none", color: "#6B7280" }}
            >
              Cancel
            </Button>
            <Button
              variant="contained"
              color="error"
              onClick={handleDelete}
              disabled={deleting}
              sx={{ borderRadius: "8px", textTransform: "none", fontWeight: 600 }}
            >
              {deleting ? "Deleting..." : `Delete ${selectedIds.size}`}
            </Button>
          </DialogActions>
        </Dialog>

        {/* Snackbar */}
        <Snackbar
          open={!!snackbar}
          autoHideDuration={4000}
          onClose={() => setSnackbar(null)}
          anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
        >
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
    </Box>
  );
}
