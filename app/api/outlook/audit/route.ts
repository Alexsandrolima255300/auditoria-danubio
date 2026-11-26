import { NextRequest, NextResponse } from "next/server";
import { openToken, refreshMicrosoftToken, sealToken } from "../../../../lib/microsoft";
import { evaluateFreight, extractCte } from "../../../../lib/cte-audit";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;

const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";

function supabaseHeaders() {
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error("Falta cadastrar SUPABASE_SERVICE_ROLE_KEY na Vercel para salvar auditorias.");
  return { apikey: key, Authorization: "Bearer " + key, "Content-Type": "application/json" };
}
async function db(path: string, init: RequestInit = {}) {
  const response = await fetch(PROJECT_URL + "/rest/v1/" + path, {
    ...init, headers: { ...supabaseHeaders(), ...(init.headers || {}) }, cache: "no-store",
  });
  const body = await response.text();
  if (!response.ok) throw new Error("Supabase: " + body.slice(0, 500));
  return body ? JSON.parse(body) : null;
}
async function graph(accessToken: string, path: string) {
  const response = await fetch("https://graph.microsoft.com/v1.0" + path, {
    headers: { Authorization: "Bearer " + accessToken, ConsistencyLevel: "eventual" }, cache: "no-store",
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error?.message || "Microsoft Graph HTTP " + response.status);
  return data;
}
async function searchAll(accessToken: string, term: string, maxMessages: number, startingUrl?: string | null) {
  const first = new URL("https://graph.microsoft.com/v1.0/me/messages");
  first.searchParams.set("$search", '"' + term + '"');
  first.searchParams.set("$top", String(maxMessages));
  first.searchParams.set("$select", "id,subject,from,receivedDateTime,hasAttachments,internetMessageId,conversationId");
  let next: string | null = startingUrl || first.toString();
  const items: any[] = [];
  let pages = 0;
  while (next && items.length < maxMessages && pages < 1) {
    const response = await fetch(next, { headers: { Authorization: "Bearer " + accessToken, ConsistencyLevel: "eventual" }, cache: "no-store" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error?.message || "Falha na busca de e-mails.");
    items.push(...(data.value || []));
    next = data["@odata.nextLink"] || null;
    pages++;
  }
  return { items: items.slice(0, maxMessages), hasMore: Boolean(next), nextLink: next };
}
async function listIncremental(accessToken: string, startingUrl: string | null, since: string) {
  const first = new URL("https://graph.microsoft.com/v1.0/me/messages");
  first.searchParams.set("$filter", "receivedDateTime ge " + since);
  first.searchParams.set("$top", "2");
  first.searchParams.set("$select", "id,subject,from,receivedDateTime,hasAttachments,internetMessageId,conversationId,bodyPreview");
  const url = startingUrl || first.toString();
  const response = await fetch(url, { headers: { Authorization: "Bearer " + accessToken }, cache: "no-store" });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error?.message || "Falha na sincronização incremental do Outlook.");
  const candidates = (data.value || []).filter((m: any) =>
    /danubio|danúbio/i.test((m.subject || "") + " " + (m.bodyPreview || "")) ||
    /siimp/i.test((m.from?.emailAddress?.name || "") + " " + (m.from?.emailAddress?.address || "") + " " + (m.subject || "") + " " + (m.bodyPreview || ""))
  );
  return { items: candidates, nextLink: data["@odata.nextLink"] || null };
}
async function listDateMessages(accessToken: string, startingUrl: string | null, dayStart: string, nextDayStart: string) {
  const first = new URL("https://graph.microsoft.com/v1.0/me/messages");
  first.searchParams.set("$filter", "receivedDateTime ge " + dayStart + " and receivedDateTime lt " + nextDayStart);
  first.searchParams.set("$top", "25");
  first.searchParams.set("$select", "id,subject,from,receivedDateTime,hasAttachments,internetMessageId,conversationId,bodyPreview");
  const url = startingUrl || first.toString();
  const response = await fetch(url, { headers: { Authorization: "Bearer " + accessToken }, cache: "no-store" });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error?.message || "Falha ao buscar e-mails da data escolhida.");
  const candidates = (data.value || []).filter((m: any) =>
    /danubio|danúbio/i.test((m.subject || "") + " " + (m.bodyPreview || "")) ||
    /siimp/i.test((m.from?.emailAddress?.name || "") + " " + (m.from?.emailAddress?.address || "") + " " + (m.subject || "") + " " + (m.bodyPreview || ""))
  );
  return { items: candidates, nextLink: data["@odata.nextLink"] || null };
}
function isPdf(a: any) {
  return (a.contentType || "").toLowerCase().includes("pdf") || (a.name || "").toLowerCase().endsWith(".pdf");
}
function isCteCandidate(a: any) {
  return isPdf(a) && /cte|ct-e|conhecimento|frete|danubio|danúbio/i.test(a.name || "");
}

export async function GET(request: NextRequest) {
  const secret = process.env.CRON_SECRET;
  if (!secret || request.headers.get("authorization") !== "Bearer " + secret) {
    return NextResponse.json({ error: "Não autorizado." }, { status: 401 });
  }
  const internalRequest = new NextRequest(request.url, {
    method: "POST",
    headers: request.headers,
    body: "{}",
  });
  return POST(internalRequest);
}

export async function POST(request: NextRequest) {
  const payload = await request.json().catch(() => ({} as any));
  const encryptedRefresh = request.cookies.get("outlook_refresh")?.value;
  const cronSecret = process.env.CRON_SECRET;
  const cronAuthorized = Boolean(cronSecret && request.headers.get("authorization") === "Bearer " + cronSecret);
  if (!encryptedRefresh && !cronAuthorized) return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });
  try {
    let refreshToken: string;
    if (encryptedRefresh) {
      refreshToken = openToken(encryptedRefresh);
    } else {
      const connections = await db("outlook_connections?select=encrypted_token_cache&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
      if (!connections?.length || !connections[0].encrypted_token_cache) {
        throw new Error("Outlook ainda não conectado. Faça uma primeira auditoria manual para registrar a conexão.");
      }
      refreshToken = openToken(connections[0].encrypted_token_cache);
    }
    const tokens = await refreshMicrosoftToken(refreshToken);
    const profile = await graph(tokens.access_token, "/me?$select=displayName,mail,userPrincipalName");
    const mailbox = profile.mail || profile.userPrincipalName;
    if (!mailbox) throw new Error("Não foi possível identificar a caixa de e-mail conectada.");

    const existing = await db("outlook_connections?select=id,audit_checkpoint&mailbox_email=eq." + encodeURIComponent(mailbox) + "&revoked_at=is.null&limit=1");
    let connectionId: string;
    const connectionPayload = {
      provider: "microsoft_graph", mailbox_email: mailbox, tenant_id: process.env.MICROSOFT_TENANT_ID,
      encrypted_token_cache: sealToken(tokens.refresh_token || refreshToken),
      token_expires_at: new Date(Date.now() + tokens.expires_in * 1000).toISOString(),
      granted_scopes: ["openid", "profile", "offline_access", "User.Read", "Mail.Read"],
      last_refreshed_at: new Date().toISOString(), revoked_at: null,
    };
    if (existing?.length) {
      connectionId = existing[0].id;
      await db("outlook_connections?id=eq." + connectionId, { method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify(connectionPayload) });
    } else {
      const created = await db("outlook_connections", { method: "POST", headers: { Prefer: "return=representation" }, body: JSON.stringify(connectionPayload) });
      connectionId = created[0].id;
    }

    // O checkpoint fica no Supabase, não no navegador. Assim, uma página recarregada
    // ou execução interrompida continua do último nextLink confirmado.
    const checkpoint = existing?.[0]?.audit_checkpoint || {};
    let messages: any[] = [];
    let hasMore = false;
    let nextLink: string | null = null;
    let syncStartedAt: string | null = null;
    let legacyDanubio: any = { hasMore: false, nextLink: null };
    let legacySiimp: any = { hasMore: false, nextLink: null };

    let dateNextLink: string | null = null;
    if (typeof payload.receivedDate === "string" && /^\\d{4}-\\d{2}-\\d{2}$/.test(payload.receivedDate)) {
      const dayStart = new Date(payload.receivedDate + "T00:00:00-03:00").toISOString();
      const nextDayStart = new Date(new Date(dayStart).getTime() + 24 * 60 * 60 * 1000).toISOString();
      const batch = await listDateMessages(tokens.access_token, typeof payload.dateNextLink === "string" ? payload.dateNextLink : null, dayStart, nextDayStart);
      messages = batch.items;
      dateNextLink = batch.nextLink;
      hasMore = Boolean(dateNextLink);
    } else if (payload.reprocessProcessed === true) {
      const [danubio, siimp] = await Promise.all([
        payload.doneDanubio === true ? Promise.resolve({ items: [], hasMore: false, nextLink: null }) : searchAll(tokens.access_token, "Danubio", 2, payload.cursorDanubio),
        payload.doneSiimp === true ? Promise.resolve({ items: [], hasMore: false, nextLink: null }) : searchAll(tokens.access_token, "Siimp Sistemas", 2, payload.cursorSiimp),
      ]);
      const unique = new Map<string, any>();
      for (const m of [...danubio.items, ...siimp.items]) unique.set(m.id, m);
      messages = [...unique.values()].sort((a,b) => new Date(a.receivedDateTime).getTime() - new Date(b.receivedDateTime).getTime());
      legacyDanubio = danubio; legacySiimp = siimp;
      hasMore = danubio.hasMore || siimp.hasMore;
      nextLink = null;
    } else {
      syncStartedAt = checkpoint.syncStartedAt || new Date().toISOString();
      let lastSyncAt = checkpoint.lastSyncAt || null;
      if (!lastSyncAt) {
        const latestSaved = await db("email_messages?select=received_at&connection_id=eq." + connectionId + "&order=received_at.desc&limit=1");
        if (latestSaved?.[0]?.received_at) lastSyncAt = latestSaved[0].received_at;
      }
      const since = lastSyncAt
        ? new Date(new Date(lastSyncAt).getTime() - 10 * 60 * 1000).toISOString()
        : new Date(Date.now() - 10 * 60 * 1000).toISOString();
      const batch = await listIncremental(tokens.access_token, checkpoint.nextLink || null, since);
      messages = batch.items;
      nextLink = batch.nextLink;
      hasMore = Boolean(nextLink);
    }
    messages = messages.filter((m: any) => !m["@removed"]);
    let emailsSaved = 0, attachmentsSaved = 0, pdfsRead = 0, auditsCreated = 0, ignored = 0, alreadyProcessedSkipped = 0, errors = 0;
    const results: any[] = [];
    for (const mail of messages) {
      try {
        const senderName = mail.from?.emailAddress?.name || "";
        const senderEmail = mail.from?.emailAddress?.address || "";
        const subject = mail.subject || "(sem assunto)";
        const matchedTerms = [];
        if (/danubio/i.test(subject)) matchedTerms.push("Danubio");
        if (/siimp/i.test(senderName + " " + senderEmail + " " + subject)) matchedTerms.push("Siimp Sistemas");
        const savedMessages = await db("email_messages?on_conflict=connection_id,graph_message_id", {
          method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=representation" },
          body: JSON.stringify({
            connection_id: connectionId, graph_message_id: mail.id, internet_message_id: mail.internetMessageId || null,
            conversation_id: mail.conversationId || null, sender_name: senderName, sender_email: senderEmail,
            subject, received_at: mail.receivedDateTime, has_attachments: Boolean(mail.hasAttachments),
            matched_terms: matchedTerms.length ? matchedTerms : ["Danubio/Siimp search"],
            processing_status: mail.hasAttachments ? "discovered" : "attachments_checked",
          }),
        });
        const messageId = savedMessages?.[0]?.id;
        emailsSaved++;
        if (!mail.hasAttachments || !messageId) continue;

        const attachmentData = await graph(tokens.access_token, "/me/messages/" + encodeURIComponent(mail.id) + "/attachments?$top=50");
        for (const attachment of attachmentData.value || []) {
          if (attachment["@odata.type"] && !attachment["@odata.type"].includes("fileAttachment")) continue;
          if (attachment.isInline || !isPdf(attachment)) { ignored++; continue; }
          const attachmentId = attachment.id;
          // Arquivos já extraídos ficam registrados e não são baixados/lidos novamente
          // em varreduras automáticas. Só uma solicitação humana explícita força a releitura.
          const previousAttachments = await db(
            "freight_attachments?select=id,processing_status&email_message_id=eq." +
            encodeURIComponent(messageId) + "&graph_attachment_id=eq." + encodeURIComponent(attachmentId) + "&limit=1"
          );
          const previousAttachment = previousAttachments?.[0];
          if (previousAttachment?.processing_status === "extracted" && payload.reprocessProcessed !== true) {
            ignored++;
            alreadyProcessedSkipped++;
            continue;
          }
          const savedAttachments = await db("freight_attachments?on_conflict=email_message_id,graph_attachment_id", {
            method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=representation" },
            body: JSON.stringify({
              email_message_id: messageId, graph_attachment_id: attachmentId, file_name: attachment.name || "anexo.pdf",
              content_type: attachment.contentType || "application/pdf", size_bytes: attachment.size || null,
              document_type: isCteCandidate(attachment) ? "cte" : "unknown", processing_status: "pending",
            }),
          });
          const attachmentRowId = savedAttachments?.[0]?.id;
          attachmentsSaved++;
          if (!attachmentRowId) continue;
          try {
            const file = await graph(tokens.access_token, "/me/messages/" + encodeURIComponent(mail.id) + "/attachments/" + encodeURIComponent(attachmentId));
            if (!file.contentBytes) { ignored++; continue; }
            const pdfBuffer = Buffer.from(file.contentBytes, "base64");
            if (pdfBuffer.length > 12 * 1024 * 1024) { ignored++; continue; }
            const extracted = await extractCte(pdfBuffer);
            const audit = evaluateFreight(extracted);
            pdfsRead++;
            await db("freight_attachments?id=eq." + attachmentRowId, {
              method: "PATCH", headers: { Prefer: "return=minimal" },
              body: JSON.stringify({ processing_status: "extracted", document_type: "cte" }),
            });
            await db("freight_audits?on_conflict=attachment_id", {
              method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=minimal" },
              body: JSON.stringify({
                attachment_id: attachmentRowId, tomador_name: extracted.tomadorName, tomador_cnpj: extracted.tomadorCnpj,
                invoice_value: extracted.invoiceValue, weight_kg: extracted.weightKg, freight_charge: extracted.freightNet,
                reference_minimum: audit.reference, discrepancy_amount: audit.discrepancy, audit_status: audit.status,
                discrepancy_reason: audit.reason, extracted_fields: { ...extracted.fields, grossFreight: extracted.freightGross, icmsValue: extracted.icmsValue, extractionErrors: extracted.errors },
              }),
            });
            auditsCreated++;
            results.push({ file: attachment.name, email: subject, status: audit.status, reason: audit.reason, tomadorCnpj: extracted.tomadorCnpj, invoiceValue: extracted.invoiceValue, weightKg: extracted.weightKg, freight: extracted.freightNet, freightGross: extracted.freightGross, icms: extracted.icmsValue, reference: audit.reference, discrepancy: audit.discrepancy });
          } catch (e) {
            errors++;
            await db("freight_attachments?id=eq." + attachmentRowId, { method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify({ processing_status: "error" }) });
            results.push({ file: attachment.name, status: "error", reason: e instanceof Error ? e.message : "Falha ao ler PDF." });
          }
        }
        await db("email_messages?id=eq." + messageId, { method: "PATCH", headers: { Prefer: "return=minimal" }, body: JSON.stringify({ processing_status: "attachments_checked" }) });
      } catch (e) {
        errors++;
        results.push({ email: mail.subject, status: "error", reason: e instanceof Error ? e.message : "Falha ao processar e-mail." });
      }
    }

    if (payload.reprocessProcessed !== true && !payload.receivedDate) {
      const nextCheckpoint = hasMore
        ? { mode: "incremental", lastSyncAt: checkpoint.lastSyncAt || null, syncStartedAt, nextLink }
        : { mode: "incremental", lastSyncAt: syncStartedAt, syncStartedAt: null, nextLink: null };
      await db("outlook_connections?id=eq." + connectionId, {
        method: "PATCH", headers: { Prefer: "return=minimal" },
        body: JSON.stringify({ audit_checkpoint: nextCheckpoint }),
      });
    }
    const response = NextResponse.json({
      ok: true, mailbox, searched: ["Danubio", "Siimp Sistemas"], emailsFound: messages.length,
      emailsSaved, attachmentsSaved, pdfsRead, auditsCreated, ignored, errors,
      alreadyProcessedSkipped,
      hasMore,
      nextDanubio: legacyDanubio.nextLink, nextSiimp: legacySiimp.nextLink, dateNextLink,
      doneDanubio: !legacyDanubio.hasMore, doneSiimp: !legacySiimp.hasMore, results,
      syncMode: payload.reprocessProcessed === true ? "manual-reprocess" : "incremental",
      note: "A sincronização incremental consulta apenas mensagens recebidas desde o último checkpoint. Mensagens e anexos são processados em modo somente leitura.",
    });
    if (tokens.refresh_token) response.cookies.set("outlook_refresh", sealToken(tokens.refresh_token), {
      httpOnly: true, secure: process.env.NODE_ENV === "production", sameSite: "lax", maxAge: 60 * 60 * 24 * 30, path: "/",
    });
    return response;
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Falha na auditoria dos e-mails." }, { status: 500 });
  }
}
