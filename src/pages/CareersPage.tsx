import { useState, useEffect, useMemo } from "react";
import { useNavigate } from "react-router-dom";
import {
  Box,
  Typography,
  Card,
  CardContent,
  Button,
  Chip,
  Grid,
  Alert,
  Container,
  Paper,
  TextField,
  InputAdornment,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Skeleton,
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableRow,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  IconButton,
} from "@mui/material";
import {
  Work,
  LocationOn,
  CalendarToday,
  People,
  AttachMoney,
  AccessTime,
  Search as SearchIcon,
  Close,
} from "@mui/icons-material";
import DOMPurify from "dompurify";
import { useMsal } from "@azure/msal-react";
import { fetchJobs, fetchMyApplications } from "../utils/careersService";
import type { JobListing, JobAdminApplication } from "../types";

function formatSalary(min: number | null, max: number | null): string {
  if (min == null && max == null) return "";
  const fmt = (n: number) => `RM ${n.toLocaleString()}`;
  if (min != null && max != null) return `${fmt(min)} - ${fmt(max)}`;
  if (min != null) return `From ${fmt(min)}`;
  return max != null ? `Up to ${fmt(max)}` : "";
}

function formatDate(dateStr: string | null): string {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  return d.toLocaleDateString("en-MY", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function getEmploymentTypeColor(type: string): string {
  const map: Record<string, string> = {
    "Full-time": "#0078D4",
    "Part-time": "#6264A7",
    Contract: "#E67635",
    Internship: "#498205",
  };
  return map[type] || "#6B7280";
}

function JobCard({ job, onSelect, isApplied }: { job: JobListing; onSelect: (job: JobListing) => void; isApplied: boolean }) {
  return (
    <Card
      sx={{
        borderRadius: "16px",
        boxShadow: "0 1px 3px rgba(0,0,0,0.06), 0 1px 2px rgba(0,0,0,0.04)",
        transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
        cursor: "pointer",
        opacity: isApplied ? 0.75 : 1,
        "&:hover": {
          boxShadow: "0 8px 25px rgba(0,0,0,0.08), 0 2px 6px rgba(0,0,0,0.04)",
          transform: isApplied ? "none" : "translateY(-2px)",
        },
      }}
      onClick={() => onSelect(job)}
    >
      <CardContent sx={{ p: 3 }}>
        <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", mb: 1.5 }}>
          <Typography variant="h6" sx={{ fontWeight: 700, color: "#111827", fontSize: "1.1rem", lineHeight: 1.3 }}>
            {job.title}
          </Typography>
          <Chip
            label={job.employmentType}
            size="small"
            sx={{
              backgroundColor: `${getEmploymentTypeColor(job.employmentType)}14`,
              color: getEmploymentTypeColor(job.employmentType),
              fontWeight: 600,
              fontSize: "0.7rem",
              borderRadius: "8px",
              ml: 1,
              flexShrink: 0,
            }}
          />
        </Box>

        <Box sx={{ display: "flex", gap: 0.5, flexWrap: "wrap", mb: 2 }}>
          <Chip
            label={job.department}
            size="small"
            sx={{
              backgroundColor: "#6264A7",
              color: "#ffffff",
              fontWeight: 500,
              fontSize: "0.7rem",
              borderRadius: "8px",
            }}
          />
          {isApplied && (
            <Chip
              label="Already Submitted"
              size="small"
              sx={{
                backgroundColor: "#E6F4EA",
                color: "#34A853",
                fontWeight: 600,
                fontSize: "0.65rem",
                borderRadius: "8px",
              }}
            />
          )}
        </Box>

        <Grid container spacing={1} sx={{ mb: 1 }}>
          {job.location && (
            <Grid size={{ xs: 6 }}>
              <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
                <LocationOn sx={{ fontSize: 14, color: "#6B7280" }} />
                <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>
                  {job.location}
                </Typography>
              </Box>
            </Grid>
          )}
          {formatSalary(job.salaryMin, job.salaryMax) && (
            <Grid size={{ xs: 6 }}>
              <Box sx={{ display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 0.5 }}>
                <AttachMoney sx={{ fontSize: 14, color: "#6B7280" }} />
                <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>
                  {formatSalary(job.salaryMin, job.salaryMax)}
                </Typography>
              </Box>
            </Grid>
          )}
        </Grid>

        <Box sx={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
          {job.closingDate && (
            <Box sx={{ display: "flex", alignItems: "center", gap: 0.5 }}>
              <CalendarToday sx={{ fontSize: 12, color: "#9CA3AF" }} />
              <Typography variant="caption" sx={{ color: "#9CA3AF" }}>
                Closing {formatDate(job.closingDate)}
              </Typography>
            </Box>
          )}
          <Box sx={{ display: "flex", alignItems: "center", gap: 0.5, ml: "auto" }}>
            <People sx={{ fontSize: 12, color: "#9CA3AF" }} />
            <Typography variant="caption" sx={{ color: "#9CA3AF" }}>
              {job.applicationCount} {job.applicationCount === 1 ? "applicant" : "applicants"}
            </Typography>
          </Box>
        </Box>
      </CardContent>
    </Card>
  );
}

function JobDetailDialog({
  job,
  open,
  onClose,
  isApplied,
  isAdmin,
  onTestSubmit,
}: {
  job: JobListing | null;
  open: boolean;
  onClose: () => void;
  isApplied: boolean;
  isAdmin: boolean;
  onTestSubmit: (jobId: string) => void;
}) {
  const navigate = useNavigate();

  if (!job) return null;

  return (
    <Dialog
      open={open}
      onClose={onClose}
      maxWidth="md"
      fullWidth
      slotProps={{
        paper: {
          sx: { borderRadius: "16px", maxHeight: "90vh" },
        },
      }}
    >
      <DialogTitle sx={{ pb: 1, pr: 8 }}>
        <Box sx={{ display: "flex", alignItems: "center", gap: 1, flexWrap: "wrap", mb: 1 }}>
          <Typography variant="h5" component="div" sx={{ fontWeight: 700, color: "#111827", fontSize: "1.25rem" }}>
            {job.title}
          </Typography>
          {isApplied && (
            <Chip label="Already Submitted" size="small" sx={{ backgroundColor: "#E6F4EA", color: "#34A853", fontWeight: 600, borderRadius: "8px", fontSize: "0.7rem" }} />
          )}
        </Box>
        <Box sx={{ display: "flex", flexWrap: "wrap", gap: 1, alignItems: "center" }}>
          <Chip
            label={job.department}
            size="small"
            sx={{ backgroundColor: "#6264A7", color: "#fff", fontWeight: 500, borderRadius: "8px" }}
          />
          <Chip
            label={job.employmentType}
            size="small"
            sx={{
              backgroundColor: `${getEmploymentTypeColor(job.employmentType)}14`,
              color: getEmploymentTypeColor(job.employmentType),
              fontWeight: 600,
              borderRadius: "8px",
            }}
          />
          {job.location && (
            <Typography variant="caption" sx={{ color: "#6B7280", display: "flex", alignItems: "center", gap: 0.3 }}>
              <LocationOn sx={{ fontSize: 14 }} /> {job.location}
            </Typography>
          )}
          {formatSalary(job.salaryMin, job.salaryMax) && (
            <Typography variant="caption" sx={{ color: "#6B7280", display: "flex", alignItems: "center", gap: 0.3 }}>
              <AttachMoney sx={{ fontSize: 14 }} /> {formatSalary(job.salaryMin, job.salaryMax)}
            </Typography>
          )}
        </Box>
        <Box sx={{ display: "flex", gap: 2, mt: 1 }}>
          {job.closingDate && (
            <Typography variant="caption" sx={{ color: "#9CA3AF", display: "flex", alignItems: "center", gap: 0.3 }}>
              <CalendarToday sx={{ fontSize: 12 }} /> Closing {formatDate(job.closingDate)}
            </Typography>
          )}
          <Typography variant="caption" sx={{ color: "#9CA3AF", display: "flex", alignItems: "center", gap: 0.3 }}>
            <People sx={{ fontSize: 12 }} /> {job.applicationCount} {job.applicationCount === 1 ? "applicant" : "applicants"}
          </Typography>
        </Box>
        <IconButton
          onClick={onClose}
          size="small"
          sx={{ position: "absolute", right: 12, top: 12, color: "#6B7280" }}
        >
          <Close />
        </IconButton>
      </DialogTitle>

      <DialogContent dividers sx={{ py: 3 }}>
        {job.jobDescription ? (
          <Box
            sx={{
              "& p": { mb: 1.5, lineHeight: 1.7, color: "#374151", fontSize: "0.9rem" },
              "& ul, & ol": { pl: 3, mb: 1.5 },
              "& li": { mb: 0.5, lineHeight: 1.7, color: "#374151", fontSize: "0.9rem" },
              "& h1, & h2, & h3, & h4": { mt: 2, mb: 1, fontWeight: 600, color: "#111827" },
              "& strong": { fontWeight: 600 },
              "& a": { color: "#0078D4", textDecoration: "none", fontWeight: 500, display: "inline-flex", alignItems: "center", gap: "4px", padding: "1px 6px", borderRadius: "6px", backgroundColor: "#F0F7FF", border: "1px solid rgba(0,120,212,0.15)", "&:hover": { backgroundColor: "#DBEAFE", textDecoration: "underline" } },
              "& a[href$='.jpg'], & a[href$='.jpeg'], & a[href$='.png'], & a[href$='.gif'], & a[href$='.svg'], & a[href$='.webp']": { "&::before": { content: "'🖼 '", fontSize: "12px" } },
              "& a[href$='.pdf']": { "&::before": { content: "'📄 '", fontSize: "12px" } },
              "& a[href$='.doc'], & a[href$='.docx']": { "&::before": { content: "'📝 '", fontSize: "12px" } },
              "& a[href$='.xls'], & a[href$='.xlsx']": { "&::before": { content: "'📊 '", fontSize: "12px" } },
              "& br": { display: "block", content: '""', mb: 0.5 },
            }}
            dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(job.jobDescription) }}
          />
        ) : (
          <Typography variant="body2" sx={{ color: "#9CA3AF" }}>
            No description provided.
          </Typography>
        )}
      </DialogContent>

      <DialogActions sx={{ px: 3, py: 2, gap: 1 }}>
        <Button
          onClick={onClose}
          sx={{ borderRadius: "10px", textTransform: "none", color: "#6B7280" }}
        >
          Close
        </Button>
        {isApplied && isAdmin ? (
          <Button
            variant="outlined"
            onClick={() => onTestSubmit(job.id)}
            sx={{
              borderRadius: "10px",
              textTransform: "none",
              fontWeight: 600,
              borderColor: "#E67635",
              color: "#E67635",
              px: 3,
              "&:hover": { borderColor: "#D4621A", backgroundColor: "rgba(230, 118, 53, 0.06)" },
            }}
          >
            Test Submit
          </Button>
        ) : isApplied ? (
          <Button
            variant="contained"
            disabled
            sx={{
              borderRadius: "10px",
              textTransform: "none",
              fontWeight: 600,
              px: 4,
              backgroundColor: "#9CA3AF",
            }}
          >
            Already Submitted
          </Button>
        ) : (
          <Button
            variant="contained"
            onClick={() => navigate(`/careers/${job.id}/apply`)}
            sx={{
              borderRadius: "10px",
              textTransform: "none",
              backgroundColor: "#0078D4",
              fontWeight: 600,
              px: 4,
            }}
          >
            Apply Now
          </Button>
        )}
      </DialogActions>
    </Dialog>
  );
}

export default function CareersPage() {
  const { instance, accounts } = useMsal();
  const navigate = useNavigate();
  const userEmail = accounts[0]?.username?.toLowerCase() || "";
  const [jobs, setJobs] = useState<JobListing[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchText, setSearchText] = useState("");
  const [deptFilter, setDeptFilter] = useState("");
  const [typeFilter, setTypeFilter] = useState("");
  const [salaryMinFilter, setSalaryMinFilter] = useState("");
  const [salaryMaxFilter, setSalaryMaxFilter] = useState("");
  const [sortBy, setSortBy] = useState("newest");
  const [selectedJob, setSelectedJob] = useState<JobListing | null>(null);
  const [selectedApp, setSelectedApp] = useState<JobAdminApplication | null>(null);
  const [myApps, setMyApps] = useState<JobAdminApplication[]>([]);
  const [isAdmin, setIsAdmin] = useState(false);
  const [appliedFilter, setAppliedFilter] = useState("all"); // "all" | "applied" | "unapplied"

  // Jobs that the current user has applied to → set of job listing IDs
  const appliedJobIds = useMemo(() => new Set(myApps.map((a) => a.jobListingId).filter(Boolean)), [myApps]);

  const isJobApplied = (jobId: string) => appliedJobIds.has(jobId);

  useEffect(() => {
    let cancelled = false;
    async function load() {
      try {
        const [jobData, appData] = await Promise.all([
          fetchJobs(),
          userEmail ? fetchMyApplications(userEmail).catch(() => []) : Promise.resolve([] as JobAdminApplication[]),
        ]);
        if (!cancelled) {
          setJobs(jobData);
          setMyApps(appData);
        }
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : "Failed to load jobs");
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    void load();
    return () => { cancelled = true; };
  }, [userEmail]);

  // Check admin status
  useEffect(() => {
    let cancelled = false;
    async function check() {
      try {
        const SP_SITE_URL = (import.meta.env.VITE_SP_SITE_URL || "").replace(/\/$/, "");
        const resp = await instance.acquireTokenSilent({
          scopes: [`${new URL(SP_SITE_URL).origin}/AllSites.Manage`],
          account: accounts[0],
        });
        const token = resp.accessToken;
        const groupResp = await fetch(
          `${SP_SITE_URL}/_api/web/sitegroups/getByName('_HR_ Forms Owners')/users?$select=Email`,
          { headers: { Accept: "application/json;odata=nometadata", Authorization: `Bearer ${token}` } },
        );
        if (groupResp.ok) {
          const data = await groupResp.json() as { value?: { Email?: string }[] };
          if (!cancelled) setIsAdmin((data.value || []).some((u) => (u.Email || "").toLowerCase() === userEmail));
        }
      } catch { /* not admin */ }
    }
    if (userEmail) void check();
    return () => { cancelled = true; };
  }, [instance, accounts, userEmail]);

  const departments = useMemo(() => {
    const set = new Set(jobs.map((j) => j.department).filter(Boolean));
    return [...set].sort();
  }, [jobs]);

  const employmentTypes = useMemo(() => {
    const set = new Set(jobs.map((j) => j.employmentType).filter(Boolean));
    return [...set].sort();
  }, [jobs]);

  const filteredJobs = useMemo(() => {
    const result = jobs.filter((job) => {
      if (searchText) {
        const q = searchText.toLowerCase();
        const matchesSearch =
          job.title.toLowerCase().includes(q) ||
          job.department.toLowerCase().includes(q) ||
          (job.location || "").toLowerCase().includes(q) ||
          job.employmentType.toLowerCase().includes(q);
        if (!matchesSearch) return false;
      }
      if (deptFilter && job.department !== deptFilter) return false;
      if (typeFilter && job.employmentType !== typeFilter) return false;
      if (salaryMinFilter) {
        const min = Number(salaryMinFilter);
        if (!isNaN(min) && (job.salaryMax == null || job.salaryMax < min)) return false;
      }
      if (salaryMaxFilter) {
        const max = Number(salaryMaxFilter);
        if (!isNaN(max) && (job.salaryMin == null || job.salaryMin > max)) return false;
      }
      if (appliedFilter === "applied" && !isJobApplied(job.id)) return false;
      if (appliedFilter === "unapplied" && isJobApplied(job.id)) return false;
      return true;
    });

    if (sortBy === "name") {
      result.sort((a, b) => a.title.localeCompare(b.title));
    } else {
      // newest first
      result.sort((a, b) => new Date(b.created).getTime() - new Date(a.created).getTime());
    }

    return result;
  }, [jobs, searchText, deptFilter, typeFilter, salaryMinFilter, salaryMaxFilter, sortBy, appliedFilter, appliedJobIds]);

  const hasFilters = searchText || deptFilter || typeFilter || salaryMinFilter || salaryMaxFilter || appliedFilter !== "all";

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
        <Container maxWidth="lg">
          <Box sx={{ display: "flex", alignItems: "center", gap: 2, py: 2.5 }}>
            <Box
              sx={{
                width: 44,
                height: 44,
                borderRadius: "12px",
                backgroundColor: "#0078D4",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <Work sx={{ fontSize: 24, color: "#ffffff" }} />
            </Box>
            <Box>
              <Typography variant="h5" sx={{ fontWeight: 700, color: "#111827", fontSize: "1.3rem", lineHeight: 1.2 }}>
                Career Opportunities
              </Typography>
              <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.85rem" }}>
                Explore open positions and join our team
              </Typography>
            </Box>
            <Box sx={{ flexGrow: 1 }} />
            {myApps.length > 0 && (
              <Button
                variant="outlined"
                size="small"
                onClick={() => setAppliedFilter(appliedFilter === "applied" ? "all" : "applied")}
                sx={{
                  borderRadius: "10px",
                  textTransform: "none",
                  fontWeight: 600,
                  fontSize: "0.8rem",
                  borderColor: "#D1D5DB",
                  color: appliedFilter === "applied" ? "#0078D4" : "#6B7280",
                  backgroundColor: appliedFilter === "applied" ? "#F0F7FF" : "transparent",
                }}
              >
                My Applications ({myApps.length})
              </Button>
            )}
          </Box>
        </Container>
      </Paper>

      <Container maxWidth="lg" sx={{ py: 4 }}>
        {/* Filters (hidden when viewing My Applications) */}
        {!loading && !error && jobs.length > 0 && appliedFilter !== "applied" && (
          <Paper
            sx={{
              p: 2,
              mb: 3,
              borderRadius: "16px",
              boxShadow: "0 1px 3px rgba(0,0,0,0.06)",
              display: "flex",
              flexWrap: "wrap",
              gap: 2,
              alignItems: "center",
            }}
          >
            <TextField
              placeholder="Search jobs..."
              value={searchText}
              onChange={(e) => setSearchText(e.target.value)}
              size="small"
              sx={{
                flex: "1 1 260px",
                minWidth: 200,
                "& .MuiOutlinedInput-root": {
                  borderRadius: "10px",
                  backgroundColor: "#F8F9FC",
                },
              }}
              slotProps={{
                input: {
                  startAdornment: (
                    <InputAdornment position="start">
                      <SearchIcon sx={{ color: "#6B7280", fontSize: 20 }} />
                    </InputAdornment>
                  ),
                },
              }}
            />
            <FormControl size="small" sx={{ minWidth: 160 }}>
              <InputLabel>Department</InputLabel>
              <Select
                value={deptFilter}
                label="Department"
                onChange={(e) => setDeptFilter(e.target.value)}
                sx={{ borderRadius: "10px", backgroundColor: "#F8F9FC" }}
              >
                <MenuItem value="">All departments</MenuItem>
                {departments.map((d) => (
                  <MenuItem key={d} value={d}>{d}</MenuItem>
                ))}
              </Select>
            </FormControl>
            <FormControl size="small" sx={{ minWidth: 150 }}>
              <InputLabel>Type</InputLabel>
              <Select
                value={typeFilter}
                label="Type"
                onChange={(e) => setTypeFilter(e.target.value)}
                sx={{ borderRadius: "10px", backgroundColor: "#F8F9FC" }}
              >
                <MenuItem value="">All types</MenuItem>
                {employmentTypes.map((t) => (
                  <MenuItem key={t} value={t}>{t}</MenuItem>
                ))}
              </Select>
            </FormControl>
            <TextField
              placeholder="Min"
              type="number"
              value={salaryMinFilter}
              onChange={(e) => setSalaryMinFilter(e.target.value)}
              size="small"
              sx={{
                width: 100,
                "& .MuiOutlinedInput-root": {
                  borderRadius: "10px",
                  backgroundColor: "#F8F9FC",
                },
              }}
              slotProps={{
                input: {
                  startAdornment: (
                    <InputAdornment position="start">
                      <Typography variant="caption" sx={{ color: "#6B7280", fontWeight: 600 }}>RM</Typography>
                    </InputAdornment>
                  ),
                },
                htmlInput: { min: 0 },
              }}
            />
            <TextField
              placeholder="Max"
              type="number"
              value={salaryMaxFilter}
              onChange={(e) => setSalaryMaxFilter(e.target.value)}
              size="small"
              sx={{
                width: 100,
                "& .MuiOutlinedInput-root": {
                  borderRadius: "10px",
                  backgroundColor: "#F8F9FC",
                },
              }}
              slotProps={{
                input: {
                  startAdornment: (
                    <InputAdornment position="start">
                      <Typography variant="caption" sx={{ color: "#6B7280", fontWeight: 600 }}>RM</Typography>
                    </InputAdornment>
                  ),
                },
                htmlInput: { min: 0 },
              }}
            />
            <FormControl size="small" sx={{ minWidth: 130 }}>
              <InputLabel>Applied</InputLabel>
              <Select
                value={appliedFilter}
                label="Applied"
                onChange={(e) => setAppliedFilter(e.target.value)}
                sx={{ borderRadius: "10px", backgroundColor: "#F8F9FC" }}
              >
                <MenuItem value="all">All jobs</MenuItem>
                <MenuItem value="applied">Applied</MenuItem>
                <MenuItem value="unapplied">Unapplied</MenuItem>
              </Select>
            </FormControl>
            <FormControl size="small" sx={{ minWidth: 130 }}>
              <InputLabel>Sort</InputLabel>
              <Select
                value={sortBy}
                label="Sort"
                onChange={(e) => setSortBy(e.target.value)}
                sx={{ borderRadius: "10px", backgroundColor: "#F8F9FC" }}
              >
                <MenuItem value="newest">Newest</MenuItem>
                <MenuItem value="name">Name</MenuItem>
              </Select>
            </FormControl>
            {hasFilters && (
              <Chip
                label={`${filteredJobs.length} of ${jobs.length} positions`}
                size="small"
                sx={{
                  backgroundColor: "#F0F7FF",
                  color: "#0078D4",
                  fontWeight: 600,
                  fontSize: "0.75rem",
                  height: 32,
                }}
              />
            )}
          </Paper>
        )}

        {/* Loading */}
        {loading && (
          <Grid container spacing={2.5}>
            {[1, 2, 3].map((i) => (
              <Grid size={{ xs: 12, sm: 6, lg: 4 }} key={i}>
                <Paper sx={{ p: 3, borderRadius: "16px" }}>
                  <Skeleton variant="text" width="75%" height={28} sx={{ mb: 1 }} />
                  <Skeleton variant="rounded" width={100} height={24} sx={{ borderRadius: "8px", mb: 2 }} />
                  <Box sx={{ display: "flex", gap: 2, mb: 1 }}>
                    <Skeleton variant="text" width="40%" height={20} />
                    <Skeleton variant="text" width="40%" height={20} />
                  </Box>
                  <Box sx={{ display: "flex", gap: 2 }}>
                    <Skeleton variant="text" width="35%" height={16} />
                    <Skeleton variant="text" width="30%" height={16} />
                  </Box>
                </Paper>
              </Grid>
            ))}
          </Grid>
        )}

        {/* Error */}
        {!loading && error && (
          <Alert
            severity="error"
            sx={{ borderRadius: "12px", mb: 3, fontWeight: 500, backgroundColor: "#FEF2F2", color: "#991B1B", "& .MuiAlert-icon": { color: "#DC2626" } }}
            action={
              <Button size="small" onClick={() => window.location.reload()} sx={{ textTransform: "none" }}>
                Retry
              </Button>
            }
          >
            {error}
          </Alert>
        )}

        {/* Empty */}
        {!loading && !error && jobs.length === 0 && (
          <Box sx={{ textAlign: "center", py: 8 }}>
            <AccessTime sx={{ fontSize: 48, color: "#D1D5DB", mb: 2 }} />
            <Typography variant="h6" sx={{ color: "#6B7280", fontWeight: 600, mb: 0.5 }}>
              No Open Positions
            </Typography>
            <Typography variant="body2" sx={{ color: "#9CA3AF" }}>
              There are no job openings at the moment. Please check back later.
            </Typography>
          </Box>
        )}
        {!loading && !error && jobs.length > 0 && filteredJobs.length === 0 && (
          <Box sx={{ textAlign: "center", py: 8 }}>
            <SearchIcon sx={{ fontSize: 48, color: "#D1D5DB", mb: 2 }} />
            <Typography variant="h6" sx={{ color: "#6B7280", fontWeight: 600, mb: 0.5 }}>
              No Positions Match
            </Typography>
            <Typography variant="body2" sx={{ color: "#9CA3AF" }}>
              Try adjusting your search or filters.
            </Typography>
          </Box>
        )}

        {/* My Applications list */}
        {!loading && !error && appliedFilter === "applied" && myApps.length > 0 && (
          <Paper sx={{ borderRadius: "16px", boxShadow: "0 1px 3px rgba(0,0,0,0.06)", overflow: "hidden" }}>
            <Table>
              <TableHead>
                <TableRow sx={{ backgroundColor: "#F9FAFB" }}>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>Reference</TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>Job Title</TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>Status</TableCell>
                  <TableCell sx={{ fontWeight: 600, color: "#6B7280", fontSize: "0.75rem", textTransform: "uppercase" }}>Submitted</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {myApps.map((app) => (
                  <TableRow key={app.id} hover sx={{ cursor: "pointer", "&:hover": { backgroundColor: "#FAFBFC" } }} onClick={() => setSelectedApp(app)}>
                    <TableCell>
                      <Typography variant="body2" sx={{ fontFamily: "monospace", fontWeight: 600, color: "#0078D4", fontSize: "0.8rem" }}>
                        {app.submissionRef}
                      </Typography>
                    </TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ fontWeight: 600, color: "#111827", fontSize: "0.85rem" }}>
                        {app.jobTitle}
                      </Typography>
                    </TableCell>
                    <TableCell>
                      <Chip
                        label={app.status || "New"}
                        size="small"
                        sx={{
                          borderRadius: "8px",
                          fontSize: "0.7rem",
                          fontWeight: 600,
                          backgroundColor: app.status === "Reviewed" ? "#E6F4EA" : "#F0F7FF",
                          color: app.status === "Reviewed" ? "#34A853" : "#0078D4",
                        }}
                      />
                    </TableCell>
                    <TableCell>
                      <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.8rem" }}>
                        {app.submittedAt ? formatDate(app.submittedAt) : "—"}
                      </Typography>
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </Paper>
        )}

        {/* Job Cards Grid (hidden when viewing My Applications) */}
        {!loading && !error && appliedFilter !== "applied" && filteredJobs.length > 0 && (
          <Grid container spacing={2.5}>
            {filteredJobs.map((job) => (
              <Grid size={{ xs: 12, sm: 6, lg: 4 }} key={job.id}>
                <JobCard job={job} onSelect={setSelectedJob} isApplied={isJobApplied(job.id)} />
              </Grid>
            ))}
          </Grid>
        )}

        {/* Application detail dialog */}
        <Dialog open={!!selectedApp} onClose={() => setSelectedApp(null)} maxWidth="sm" fullWidth slotProps={{ paper: { sx: { borderRadius: "16px" } } }}>
          {selectedApp && (
            <>
              <DialogTitle sx={{ pb: 1 }}>
                <Typography variant="h6" component="div" sx={{ fontWeight: 700, color: "#111827" }}>
                  Application Details
                </Typography>
                <IconButton onClick={() => setSelectedApp(null)} size="small" sx={{ position: "absolute", right: 12, top: 12, color: "#6B7280" }}><Close /></IconButton>
              </DialogTitle>
              <DialogContent>
                <Box sx={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  <Box><Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>Reference</Typography><Typography variant="body2" sx={{ fontFamily: "monospace", fontWeight: 600, color: "#0078D4" }}>{selectedApp.submissionRef}</Typography></Box>
                  <Box><Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>Job Title</Typography><Typography variant="body1" sx={{ fontWeight: 600, color: "#111827" }}>{selectedApp.jobTitle}</Typography></Box>
                  <Box><Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>Applicant</Typography><Typography variant="body1" sx={{ fontWeight: 600, color: "#111827" }}>{selectedApp.applicantName}</Typography><Typography variant="body2" sx={{ color: "#6B7280" }}>{selectedApp.applicantEmail}</Typography></Box>
                  <Box><Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>Status</Typography><Chip label={selectedApp.status || "New"} size="small" sx={{ borderRadius: "8px", fontWeight: 600, backgroundColor: selectedApp.status === "Reviewed" ? "#E6F4EA" : "#F0F7FF", color: selectedApp.status === "Reviewed" ? "#34A853" : "#0078D4" }} /></Box>
                  <Box><Typography variant="caption" sx={{ color: "#9CA3AF", fontWeight: 500 }}>Submitted</Typography><Typography variant="body2" sx={{ color: "#6B7280" }}>{selectedApp.submittedAt ? formatDate(selectedApp.submittedAt) : "—"}</Typography></Box>
                </Box>
              </DialogContent>
              <DialogActions sx={{ px: 3, pb: 2 }}>
                <Button onClick={() => setSelectedApp(null)} sx={{ borderRadius: "8px", textTransform: "none", color: "#6B7280" }}>Close</Button>
              </DialogActions>
            </>
          )}
        </Dialog>

        {/* Job detail dialog */}
        <JobDetailDialog
          job={selectedJob}
          open={!!selectedJob}
          onClose={() => setSelectedJob(null)}
          isApplied={!!(selectedJob && isJobApplied(selectedJob.id))}
          isAdmin={isAdmin}
          onTestSubmit={(jobId) => {
            setSelectedJob(null);
            navigate(`/careers/${jobId}/apply`);
          }}
        />
      </Container>
    </Box>
  );
}
