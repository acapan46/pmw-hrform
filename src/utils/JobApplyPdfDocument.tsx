/**
 * JobApplyPdfDocument.tsx — Corporate-style PDF for job applications.
 * Printer-friendly, B&W clear, suitable for HR records.
 */
import { Document, Page, View, Text, StyleSheet } from "@react-pdf/renderer";

const C = {
  primary: "#1a1a1a",
  border: "#999",
  borderLight: "#ccc",
  text: "#000",
  muted: "#333",
  headerBg: "#eee",
};

const S = StyleSheet.create({
  page: { padding: 36, fontFamily: "Helvetica", fontSize: 10, color: C.text },
  // Header
  header: { marginBottom: 20, paddingBottom: 12, borderBottomWidth: 2, borderBottomColor: C.text },
  headerTop: { flexDirection: "row", justifyContent: "space-between", alignItems: "flex-start" },
  docTitle: { fontSize: 18, fontWeight: "heavy", color: C.text },
  docSub: { fontSize: 7, color: C.muted, marginTop: 2 },
  headerRight: { alignItems: "flex-end" },
  refText: { fontSize: 8, fontFamily: "Courier", color: C.muted },
  dateText: { fontSize: 7, color: C.muted, marginTop: 2 },
  // Section
  sectionTitle: { fontSize: 10, fontWeight: "heavy", textTransform: "uppercase", letterSpacing: 1, marginTop: 16, marginBottom: 6, paddingBottom: 2, borderBottomWidth: 1, borderBottomColor: C.text },
  // Info table
  infoRow: { flexDirection: "row", borderBottomWidth: 0.5, borderBottomColor: C.borderLight, paddingVertical: 4 },
  infoLabel: { width: "30%", fontSize: 8, color: C.muted, fontWeight: "bold" },
  infoValue: { width: "70%", fontSize: 8, color: C.text },
  // Field rows
  fieldRow: { flexDirection: "row", paddingVertical: 3, paddingHorizontal: 2, borderBottomWidth: 0.3, borderBottomColor: C.borderLight },
  fieldRowAlt: { backgroundColor: "#f9f9f9" },
  fieldLabel: { width: "35%", fontSize: 7.5, color: C.muted, paddingRight: 4 },
  fieldValue: { width: "65%", fontSize: 7.5, color: C.text },
  // Reasoning box
  reasoningBox: { marginTop: 8, padding: 10, borderWidth: 0.5, borderColor: C.borderLight, backgroundColor: "#fafafa" },
  reasoningText: { fontSize: 7.5, color: C.text, lineHeight: 1.6 },
  // Footer
  footer: { position: "absolute", bottom: 20, left: 36, right: 36, flexDirection: "row", justifyContent: "space-between", paddingTop: 4, borderTopWidth: 0.5, borderTopColor: C.borderLight, fontSize: 6, color: C.muted },
});

interface JobApplyPdfData {
  jobTitle: string;
  applicantName: string;
  applicantEmail: string;
  applicantPhone: string;
  currentPosition?: string;
  currentDepartment?: string;
  submissionRef: string;
  submittedAt: string;
  reasoning?: string;
  customAnswers?: Record<string, unknown>;
}

export default function JobApplyPdfDocument({ data }: { data: JobApplyPdfData }) {
  const submittedAtFormatted = new Date(data.submittedAt).toLocaleString("en-MY", {
    year: "numeric", month: "long", day: "numeric", hour: "2-digit", minute: "2-digit",
  });

  const hasCustom = data.customAnswers && Object.keys(data.customAnswers).length > 0;

  return (
    <Document>
      <Page size="A4" style={S.page}>
        {/* ═══ HEADER ═══ */}
        <View style={S.header}>
          <View style={S.headerTop}>
            <View>
              <Text style={S.docTitle}>Job Application</Text>
              <Text style={S.docSub}>Internal Application Form — PMW Group</Text>
            </View>
            <View style={S.headerRight}>
              <Text style={S.refText}>Ref: {data.submissionRef}</Text>
              <Text style={S.dateText}>Submitted: {submittedAtFormatted}</Text>
            </View>
          </View>
        </View>

        {/* ═══ APPLICANT INFO ═══ */}
        <Text style={S.sectionTitle}>Applicant Information</Text>
        <View style={S.infoRow}><Text style={S.infoLabel}>Position</Text><Text style={S.infoValue}>{data.jobTitle}</Text></View>
        <View style={S.infoRow}><Text style={S.infoLabel}>Full Name</Text><Text style={S.infoValue}>{data.applicantName}</Text></View>
        <View style={S.infoRow}><Text style={S.infoLabel}>Email</Text><Text style={S.infoValue}>{data.applicantEmail}</Text></View>
        <View style={S.infoRow}><Text style={S.infoLabel}>Phone</Text><Text style={S.infoValue}>{data.applicantPhone}</Text></View>
        {data.currentPosition && <View style={S.infoRow}><Text style={S.infoLabel}>Current Position</Text><Text style={S.infoValue}>{data.currentPosition}</Text></View>}
        {data.currentDepartment && <View style={S.infoRow}><Text style={S.infoLabel}>Current Department</Text><Text style={S.infoValue}>{data.currentDepartment}</Text></View>}

        {/* ═══ REASONING ═══ */}
        {data.reasoning && (
          <>
            <Text style={S.sectionTitle}>Reasoning</Text>
            <View style={S.reasoningBox}>
              <Text style={S.reasoningText}>{data.reasoning}</Text>
            </View>
          </>
        )}

        {/* ═══ CUSTOM ANSWERS ═══ */}
        {hasCustom && (
          <>
            <Text style={S.sectionTitle}>Additional Responses</Text>
            {Object.entries(data.customAnswers!).map(([key, value], i) => (
              <View key={i} style={[S.fieldRow, i % 2 === 1 ? S.fieldRowAlt : {}]}>
                <Text style={S.fieldLabel}>{key}</Text>
                <Text style={S.fieldValue}>{String(value ?? "")}</Text>
              </View>
            ))}
          </>
        )}

        {/* ═══ FOOTER ═══ */}
        <View style={S.footer} fixed>
          <Text>PMW HR Forms — Job Application</Text>
          <Text render={({ pageNumber, totalPages }) => `Page ${pageNumber} of ${totalPages}`} />
        </View>
      </Page>
    </Document>
  );
}
