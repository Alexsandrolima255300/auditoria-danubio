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
    const rows = await db("email_messages?select=graph_message_id,subject,sender_name,sender_email,received_at,has_attachments&connection_id=eq." +
      encodeURIComponent(connection.id) + "&order=received_at.desc&limit=50");
    const emails = (rows || []).map((mail: any) => ({
      id: mail.graph_message_id,
      subject: mail.subject || "(sem assunto)",
      sender: mail.sender_name || mail.sender_email || "Remetente desconhecido",
      senderEmail: mail.sender_email || "",
      receivedAt: mail.received_at,
      hasAttachments: Boolean(mail.has_attachments),
      preview: "Mensagem registrada na última sincronização.",
      webLink: "",
    }));
    return NextResponse.json({ connected: true, mailbox: connection.mailbox_email, count: emails.length, emails });
  } catch (error) {
    return NextResponse.json({ connected: false, error: error instanceof Error ? error.message : "Falha ao consultar mensagens salvas." }, { status: 400 });
  }
}
