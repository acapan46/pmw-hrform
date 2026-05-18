import { getGraphToken, queryListItems } from "./_utils/graphClient.js";

interface ApiRequest {
  body: Record<string, unknown>;
  method: string;
}

interface ApiResponse {
  status(code: number): ApiResponse;
  json(data: Record<string, unknown>): void;
  setHeader(name: string, value: string): void;
  end(): void;
}

export default async function handler(req: ApiRequest, res: ApiResponse) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  try {
    const token = await getGraphToken();
    const items = await queryListItems(token, "Internal Job Listing", { top: 500 });

    const jobs = items
      .filter((item) => {
        const status = String(item.fields.Status || "").toLowerCase();
        return status === "new";
      })
      .map((item) => {
        let customFields: Record<string, unknown>[] | undefined;
        const raw = item.fields.CustomFields;
        if (raw && typeof raw === "string") {
          try { customFields = JSON.parse(raw) as Record<string, unknown>[]; } catch { /* ignore */ }
        }

        return {
          id: String(item.id || item.fields.Id || ""),
          title: String(item.fields.Title || ""),
          jobDescription: String(item.fields.Job_x0020_Description || ""),
          department: String(item.fields.Department || ""),
          location: String(item.fields.Location || ""),
          employmentType: String(item.fields.Employment_x0020_Type || ""),
          salaryMin: item.fields.Salary_x0020_Min != null ? Number(item.fields.Salary_x0020_Min) : null,
          salaryMax: item.fields.Salary_x0020_Max != null ? Number(item.fields.Salary_x0020_Max) : null,
          closingDate: item.fields.Closing_x0020_Date ? String(item.fields.Closing_x0020_Date) : null,
          status: String(item.fields.Status || "New"),
          applicationCount: Number(item.fields.Application_x0020_Count) || 0,
          created: String(item.fields.Created || ""),
          customFields,
        };
      });

    return res.status(200).json({ jobs } as unknown as Record<string, unknown>);
  } catch (e) {
    console.error("[API jobs-list]", e);
    return res.status(500).json({ error: "Internal server error. Please try again." });
  }
}
