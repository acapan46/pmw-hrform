import { getGraphToken } from "./_utils/graphClient.js";

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
  // Only allow requests from the app's own origin
  const hdrs = (req as Record<string, unknown>).headers as Record<string, string | undefined> || {};
  const origin = hdrs['origin'] || hdrs['referer'] || '';
  const allowedOrigins = [process.env.VITE_SP_SITE_URL || '', process.env.APP_ORIGIN || ''].filter(Boolean);
  if (allowedOrigins.length > 0 && !allowedOrigins.some(o => origin.startsWith(o))) {
    return res.status(403).json({ error: 'Forbidden' });
  }

  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { to, subject, body } = req.body as Record<string, unknown>;

  if (!to || !subject || !body) {
    return res.status(400).json({ error: "Missing required fields: to, subject, body" });
  }

  const recipients = Array.isArray(to) ? (to as string[]) : [to as string];
  // Sender: use EMAIL_FROM env var, or fall back to a placeholder that will fail gracefully
  const fromAddress = process.env.EMAIL_FROM_ADDRESS || process.env.VITE_EMAIL_FROM_ADDRESS || "";
  if (!fromAddress) {
    return res.status(500).json({
      error: "EMAIL_FROM_ADDRESS not configured. Set this env var to a mail-enabled user (e.g. admin@tenant.onmicrosoft.com). " +
        "The Azure AD app registration also needs the 'Mail.Send' application permission (granted by admin).",
    });
  }

  try {
    const token = await getGraphToken();

    const graphRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromAddress)}/sendMail`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          message: {
            subject,
            body: {
              contentType: "HTML",
              content: body,
            },
            toRecipients: recipients.map((r: string) => ({
              emailAddress: { address: r },
            })),
          },
          saveToSentItems: false,
        }),
      }
    );

    if (!graphRes.ok) {
      const errText = await graphRes.text();
      throw new Error(`Graph sendMail failed ${graphRes.status}: ${errText}`);
    }

    return res.status(200).json({ ok: true });
  } catch (e) {
    return res.status(500).json({ error: "Internal server error. Please try again." });
  }
}
