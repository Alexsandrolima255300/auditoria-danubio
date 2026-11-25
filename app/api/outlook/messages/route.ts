import { NextRequest, NextResponse } from "next/server";
import { openToken, refreshMicrosoftToken, sealToken } from "@/lib/microsoft";

export const dynamic = "force-dynamic";

async function graphSearch(accessToken: string, term: string) {
  const url = new URL("https://graph.microsoft.com/v1.0/me/messages");
  url.searchParams.set("$search", `"${term}"`);
  url.searchParams.set("$top", "50");
  url.searchParams.set("$select", "id,subject,from,receivedDateTime,hasAttachments,bodyPreview,webLink,internetMessageId");
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}`, ConsistencyLevel: "eventual" },
    cache: "no-store",
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error?.message || `Microsoft Graph retornou HTTP ${response.status}.`);
  return (data.value || []) as Array<Record<string, any>>;
}

export async function GET(request: NextRequest) {
  const encryptedRefresh = request.cookies.get("outlook_refresh")?.value;
  if (!encryptedRefresh) return NextResponse.json({ connected: false, error: "Outlook ainda não conectado." }, { status: 401 });
  try {
    const refreshToken = openToken(encryptedRefresh);
    const tokens = await refreshMicrosoftToken(refreshToken);
    const [danubio, siimp] = await Promise.all([
      graphSearch(tokens.access_token, "Danubio"),
      graphSearch(tokens.access_token, "Siimp Sistemas"),
    ]);
    const unique = new Map<string, Record<string, any>>();
    for (const mail of [...danubio, ...siimp]) unique.set(mail.id, mail);
    const emails = [...unique.values()]
      .filter((mail) => {
        const sender = mail.from?.emailAddress?.name || "";
        const address = mail.from?.emailAddress?.address || "";
        const text = `${mail.subject || ""} ${mail.bodyPreview || ""}`;
        return /siimp/i.test(sender + " " + address) || /danubio/i.test(text);
      })
      .sort((a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime())
      .slice(0, 50)
      .map((mail) => ({
        id: mail.id,
        subject: mail.subject || "(sem assunto)",
        sender: mail.from?.emailAddress?.name || mail.from?.emailAddress?.address || "Remetente desconhecido",
        senderEmail: mail.from?.emailAddress?.address || "",
        receivedAt: mail.receivedDateTime,
        hasAttachments: Boolean(mail.hasAttachments),
        preview: mail.bodyPreview || "",
        webLink: mail.webLink || "",
      }));
    const response = NextResponse.json({ connected: true, count: emails.length, emails });
    if (tokens.refresh_token) {
      response.cookies.set("outlook_refresh", sealToken(tokens.refresh_token), {
        httpOnly: true, secure: process.env.NODE_ENV === "production", sameSite: "lax", maxAge: 60 * 60 * 24 * 30, path: "/",
      });
    }
    return response;
  } catch (error) {
    const response = NextResponse.json({ connected: false, error: error instanceof Error ? error.message : "Falha ao consultar Outlook." }, { status: 400 });
    if (/invalid_grant|revoked|Sessão do Outlook inválida/i.test(error instanceof Error ? error.message : "")) {
      response.cookies.set("outlook_refresh", "", { httpOnly: true, maxAge: 0, path: "/" });
    }
    return response;
  }
}
