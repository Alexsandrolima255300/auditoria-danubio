import { NextRequest, NextResponse } from "next/server";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";
const KEY = () => process.env.SUPABASE_SERVICE_ROLE_KEY || "";

async function db(path: string, init: RequestInit = {}) {
  const key = KEY();
  if (!key) throw new Error("SUPABASE_SERVICE_ROLE_KEY não configurada.");
  const response = await fetch(PROJECT_URL + "/rest/v1/" + path, {
    ...init,
    headers: { apikey: key, Authorization: "Bearer " + key, "Content-Type": "application/json", ...(init.headers || {}) },
    cache: "no-store",
  });
  const body = await response.text();
  if (!response.ok) throw new Error("Supabase: " + body.slice(0, 500));
  return body ? JSON.parse(body) : [];
}

const normCte = (v: string) => String(v || "").replace(/\s/g, "").toUpperCase();

export async function POST(req: NextRequest) {
  if (!req.cookies.get("outlook_refresh")?.value) {
    return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });
  }
  try {
    const invoiceId = req.nextUrl.pathname.split("/").filter(Boolean).slice(-2)[0];
    if (!invoiceId) return NextResponse.json({ error: "Fatura não informada." }, { status: 400 });

    const connections = await db("outlook_connections?select=id&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if (!connections.length) throw new Error("Conexão Outlook não encontrada.");
    const connectionId = connections[0].id;
    const invoices = await db("danubio_invoices?select=id,connection_id,period_start,period_end,extraction_data,invoice_number&id=eq." + encodeURIComponent(invoiceId) + "&connection_id=eq." + encodeURIComponent(connectionId) + "&limit=1");
    if (!invoices.length) return NextResponse.json({ error: "Fatura não encontrada nesta conexão." }, { status: 404 });
    const invoice = invoices[0];

    const items = await db("danubio_invoice_items?select=id,row_number,cte_number,issued_at,weight_kg,charged_total&invoice_id=eq." + encodeURIComponent(invoiceId) + "&order=row_number.asc&limit=1000");
    if (!items.length) return NextResponse.json({ error: "A fatura não possui CT-es para comparar." }, { status: 400 });

    const messages = await db("email_messages?select=id&connection_id=eq." + encodeURIComponent(connectionId) + "&limit=5000");
    const messageIds = messages.map((m: any) => m.id).filter(Boolean);
    const attachments = messageIds.length
      ? await db("freight_attachments?select=id,email_message_id&email_message_id=in.(" + messageIds.join(",") + ")&limit=5000")
      : [];
    const attachmentIds = attachments.map((a: any) => a.id).filter(Boolean);
    const audits = attachmentIds.length
      ? await db("freight_audits?select=id,attachment_id,weight_kg,extracted_fields,freight_charge&attachment_id=in.(" + attachmentIds.join(",") + ")&limit=5000")
      : [];
    const attachmentMap = new Map(attachments.map((a: any) => [a.id, a]));
    const source: any[] = [];
    for (const audit of audits) {
      const raw = String(audit.extracted_fields?.extractedTextSample || "");
      const header = raw.match(/CTE\s*N[º°]?\s*DOCUMENTO:\s*(\d+)\s+S[ÉE]RIE:\s*(\d+)/i);
      const cte = header ? header[2] + "/" + header[1] : (raw.match(/CTE\s*(?:N[º°]|NUMERO)?\s*[:.]?\s*(\d+\/\d+)/i)?.[1] || null);
      const issuedMatch = raw.match(/DATA\/?HORA\s+EMISS[AÃ]O\s+(\d{2})\/(\d{2})\/(\d{4})/i);
      const sourceIssued = issuedMatch ? issuedMatch[3] + "-" + issuedMatch[2] + "-" + issuedMatch[1] : null;
      const attachment: any = attachmentMap.get(audit.attachment_id);
      if (!cte || !sourceIssued || !attachment) continue;
      if (sourceIssued < (invoice.period_start || "0000-01-01") || sourceIssued > (invoice.period_end || "9999-12-31")) continue;
      source.push({ ...audit, attachment_id: attachment.id, cte, sourceIssued });
    }

    const seen = new Set<string>();
    const updates = items.map((item: any) => {
      const matches = source.filter((x: any) =>
        normCte(x.cte) === normCte(item.cte_number) &&
        (!item.issued_at || x.sourceIssued === item.issued_at)
      );
      const original = matches[0] || null;
      const weightDifference = original && original.weight_kg != null && item.weight_kg != null
        ? Math.round((Number(item.weight_kg) - Number(original.weight_kg)) * 1000) / 1000 : null;
      const originalTotal = original ? (original.extracted_fields?.grossFreight ?? original.freight_charge) : null;
      const valueDifference = originalTotal != null && item.charged_total != null
        ? Math.round((Number(item.charged_total) - Number(originalTotal)) * 100) / 100 : null;
      const key = normCte(item.cte_number);
      const duplicate = seen.has(key); seen.add(key);
      let status = "not_found";
      if (duplicate) status = "duplicate";
      else if (original) {
        const wd = weightDifference !== null && Math.abs(weightDifference) > 0.001;
        const vd = valueDifference !== null && Math.abs(valueDifference) > 0.01;
        status = wd && vd ? "both_different" : wd ? "weight_difference" : vd ? "value_difference" : "matched";
      }
      return {
        id: item.id, original_attachment_id: original?.attachment_id || null,
        original_weight_kg: original?.weight_kg ?? null, original_freight_total: originalTotal,
        weight_difference: weightDifference, value_difference: valueDifference, comparison_status: status,
      };
    });

    for (let i = 0; i < updates.length; i += 10) {
      await Promise.all(updates.slice(i, i + 10).map((item: any) => {
        const { id, ...patch } = item;
        return db("danubio_invoice_items?id=eq." + encodeURIComponent(id), {
          method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify(patch),
        });
      }));
    }
    const statuses = updates.reduce((out: any, item: any) => {
      out[item.comparison_status] = (out[item.comparison_status] || 0) + 1; return out;
    }, {});
    const extractionData = { ...(invoice.extraction_data || {}), statuses, comparisonRefreshedAt: new Date().toISOString(), comparisonSource: "historical_backfill" };
    await db("danubio_invoices?id=eq." + encodeURIComponent(invoiceId), {
      method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify({ extraction_data: extractionData }),
    });
    return NextResponse.json({ ok: true, invoiceNumber: invoice.invoice_number, compared: updates.length, statuses, recoveredOriginals: updates.filter((x: any) => x.original_attachment_id).length });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Falha ao refazer a comparação da fatura." }, { status: 500 });
  }
}
