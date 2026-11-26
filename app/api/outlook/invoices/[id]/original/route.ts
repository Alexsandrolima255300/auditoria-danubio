import { NextRequest, NextResponse } from "next/server";
import { openToken, refreshMicrosoftToken } from "../../../../../lib/microsoft";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";
const KEY = () => process.env.SUPABASE_SERVICE_ROLE_KEY || "";

async function db(path: string) {
  const key = KEY();
  if (!key) throw new Error("SUPABASE_SERVICE_ROLE_KEY não configurada.");
  const response = await fetch(PROJECT_URL + "/rest/v1/" + path, {
    headers: { apikey: key, Authorization: "Bearer " + key, "Content-Type": "application/json" },
    cache: "no-store",
  });
  const body = await response.text();
  if (!response.ok) throw new Error("Supabase: " + body.slice(0, 400));
  return body ? JSON.parse(body) : [];
}

export async function GET(req: NextRequest) {
  const refreshCookie = req.cookies.get("outlook_refresh")?.value;
  if (!refreshCookie) return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });

  try {
    const segments = req.nextUrl.pathname.split("/").filter(Boolean);
    const originalIndex = segments.lastIndexOf("original");
    const invoiceId = originalIndex > 0 ? segments[originalIndex - 1] : "";
    if (!invoiceId) return NextResponse.json({ error: "Fatura não informada." }, { status: 400 });

    const connections = await db("outlook_connections?select=id,revoked_at&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if (!connections.length) return NextResponse.json({ error: "Conexão Outlook não encontrada." }, { status: 404 });
    const connection = connections[0];

    const invoices = await db("danubio_invoices?select=id,connection_id,email_message_id,graph_attachment_id,file_name&id=eq." + encodeURIComponent(invoiceId) + "&connection_id=eq." + encodeURIComponent(connection.id) + "&limit=1");
    if (!invoices.length) return NextResponse.json({ error: "Fatura não encontrada ou fora da sua conexão." }, { status: 404 });
    const invoice = invoices[0];
    if (!invoice.email_message_id || !invoice.graph_attachment_id) {
      return NextResponse.json({ error: "O vínculo com o anexo original não está disponível para esta fatura." }, { status: 404 });
    }

    const messages = await db("email_messages?select=id,graph_message_id&id=eq." + encodeURIComponent(invoice.email_message_id) + "&connection_id=eq." + encodeURIComponent(connection.id) + "&limit=1");
    if (!messages.length || !messages[0].graph_message_id) {
      return NextResponse.json({ error: "O e-mail original não foi encontrado." }, { status: 404 });
    }

    const tokens = await refreshMicrosoftToken(openToken(refreshCookie));
    const graphUrl = "https://graph.microsoft.com/v1.0/me/messages/" +
      encodeURIComponent(messages[0].graph_message_id) + "/attachments/" +
      encodeURIComponent(invoice.graph_attachment_id);
    const graphResponse = await fetch(graphUrl, {
      headers: { Authorization: "Bearer " + tokens.access_token },
      cache: "no-store",
    });
    const attachment = await graphResponse.json();
    if (!graphResponse.ok) {
      return NextResponse.json({ error: attachment.error?.message || "Não foi possível recuperar o PDF original do Outlook." }, { status: graphResponse.status });
    }
    if (!attachment.contentBytes) {
      return NextResponse.json({ error: "O Outlook não disponibilizou o conteúdo deste anexo." }, { status: 404 });
    }

    const pdf = Buffer.from(attachment.contentBytes, "base64");
    const filename = String(attachment.name || invoice.file_name || "fatura-original.pdf").replace(/[\\/"\r\n]/g, "_");
    return new NextResponse(new Uint8Array(pdf), {
      status: 200,
      headers: {
        "Content-Type": attachment.contentType || "application/pdf",
        "Content-Disposition": 'inline; filename="' + filename + '"',
        "Cache-Control": "private, no-store, max-age=0",
        "X-Content-Type-Options": "nosniff",
      },
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Falha ao abrir a fatura original." }, { status: 500 });
  }
}
