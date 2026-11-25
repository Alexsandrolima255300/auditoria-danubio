import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";

export async function GET(request: NextRequest) {
  const secret = process.env.CRON_SECRET;
  if (!secret || request.headers.get("authorization") !== "Bearer " + secret) {
    return NextResponse.json({ error: "Não autorizado." }, { status: 401 });
  }

  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) {
    return NextResponse.json({ error: "Falta cadastrar SUPABASE_SERVICE_ROLE_KEY na Vercel." }, { status: 500 });
  }

  try {
    const cutoff = new Date(Date.now() - 45 * 24 * 60 * 60 * 1000).toISOString();
    const invoiceResponse = await fetch(
      PROJECT_URL + "/rest/v1/danubio_invoices?created_at=lt." + encodeURIComponent(cutoff) + "&select=id",
      { method: "DELETE", headers: { apikey: key, Authorization: "Bearer " + key, Prefer: "return=representation" }, cache: "no-store" }
    );
    const invoiceBody = await invoiceResponse.text();
    if (!invoiceResponse.ok) throw new Error("Supabase faturas: " + invoiceBody.slice(0, 500));
    const deletedInvoices = invoiceBody ? JSON.parse(invoiceBody) : [];

    const response = await fetch(
      PROJECT_URL + "/rest/v1/email_messages?received_at=lt." + encodeURIComponent(cutoff) + "&select=id",
      {
        method: "DELETE",
        headers: {
          apikey: key,
          Authorization: "Bearer " + key,
          Prefer: "return=representation",
        },
        cache: "no-store",
      }
    );
    const body = await response.text();
    if (!response.ok) throw new Error("Supabase: " + body.slice(0, 500));
    const deletedMessages = body ? JSON.parse(body) : [];
    return NextResponse.json({
      ok: true,
      retentionDays: 45,
      cutoff,
      deletedEmails: Array.isArray(deletedMessages) ? deletedMessages.length : 0,
      deletedInvoices: Array.isArray(deletedInvoices) ? deletedInvoices.length : 0,
      note: "Anexos, auditorias e revisões vinculados são removidos automaticamente pelo banco em cascata. Faturas e itens vinculados a e-mails antigos também são removidos; checkpoints e resumos diários são preservados.",
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Falha na limpeza de retenção." }, { status: 500 });
  }
}
