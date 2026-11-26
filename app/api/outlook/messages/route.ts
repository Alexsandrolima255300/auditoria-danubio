import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";

const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";

async function db(path: string) {
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error("Falta cadastrar SUPABASE_SERVICE_ROLE_KEY na Vercel.");
  const response = await fetch(PROJECT_URL + "/rest/v1/" + path, {
    headers: { apikey: key, Authorization: "Bearer " + key }, cache: "no-store",
  });
  const body = await response.text();
  if (!response.ok) throw new Error("Supabase: " + body.slice(0, 400));
  return body ? JSON.parse(body) : [];
}

export async function GET(request: NextRequest) {
  // A tela consulta o banco local; não dispara uma nova busca no Outlook a cada minuto.
  if (!request.cookies.get("outlook_refresh")?.value) {
    return NextResponse.json({ connected: false, error: "Outlook ainda não conectado." }, { status: 401 });
  }
  try {
    const connections = await db("outlook_connections?select=id,mailbox_email&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if (!connections?.length) return NextResponse.json({ connected: false, error: "Outlook ainda não conectado." }, { status: 401 });
    const connection = connections[0];
    const brazilDate = new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/Sao_Paulo", year: "numeric", month: "2-digit", day: "2-digit",
    }).format(new Date());
    // O Brasil está em UTC-3 (sem horário de verão); o intervalo é calculado no fuso de Brasília.
    const dayStart = new Date(brazilDate + "T00:00:00-03:00");
    const nextDayStart = new Date(dayStart.getTime() + 24 * 60 * 60 * 1000);
    const select = "select=graph_message_id,subject,sender_name,sender_email,received_at,has_attachments";
    const [rows, todayRows] = await Promise.all([
      db("email_messages?" + select + "&connection_id=eq." + encodeURIComponent(connection.id) + "&order=received_at.desc&limit=1000"),
      db("email_messages?" + select + "&connection_id=eq." + encodeURIComponent(connection.id) +
        "&received_at=gte." + encodeURIComponent(dayStart.toISOString()) +
        "&received_at=lt." + encodeURIComponent(nextDayStart.toISOString()) +
        "&order=received_at.desc&limit=1000"),
    ]);
    const mapMail = (mail: any) => ({
      id: mail.graph_message_id,
      subject: mail.subject || "(sem assunto)",
      sender: mail.sender_name || mail.sender_email || "Remetente desconhecido",
      senderEmail: mail.sender_email || "",
      receivedAt: mail.received_at,
      hasAttachments: Boolean(mail.has_attachments),
      preview: "Mensagem registrada na última sincronização.",
      webLink: "",
    });
    const emails = (rows || []).map(mapMail);
    const todayEmails = (todayRows || []).map(mapMail);
    return NextResponse.json({
      connected: true, mailbox: connection.mailbox_email, count: emails.length, emails,
      brazilDate, todayCount: todayEmails.length, todayEmails,
    });
  } catch (error) {
    return NextResponse.json({ connected: false, error: error instanceof Error ? error.message : "Falha ao consultar mensagens salvas." }, { status: 400 });
  }
}
