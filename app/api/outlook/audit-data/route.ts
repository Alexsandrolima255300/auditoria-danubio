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
  if (!request.cookies.get("outlook_refresh")?.value) {
    return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });
  }

  try {
    const connections = await db("outlook_connections?select=id,mailbox_email&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if (!connections?.length) return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });
    const connection = connections[0];

    const messages = await db(
      "email_messages?select=id,graph_message_id,subject,sender_name,sender_email,received_at,has_attachments,processing_status&connection_id=eq." +
      encodeURIComponent(connection.id) + "&order=received_at.desc&limit=2000"
    );

    if (!messages.length) {
      return NextResponse.json({ mailbox: connection.mailbox_email, documents: [], emails: [], totals: { documents: 0, divergences: 0, errors: 0 } });
    }

    const messageIds = messages.map((m: any) => m.id).filter(Boolean);
    const attachments = await db(
      "freight_attachments?select=*&email_message_id=in.(" + messageIds.join(",") + ")&order=created_at.desc&limit=1000"
    );
    const attachmentIds = attachments.map((a: any) => a.id).filter(Boolean);
    const audits = attachmentIds.length
      ? await db("freight_audits?select=*&attachment_id=in.(" + attachmentIds.join(",") + ")&order=created_at.desc&limit=1000")
      : [];

    const messageMap = new Map(messages.map((m: any) => [m.id, m]));
    const auditMap = new Map(audits.map((a: any) => [a.attachment_id, a]));
    const reviews = attachmentIds.length
      ? await db("freight_divergence_reviews?select=*&attachment_id=in.(" + attachmentIds.join(",") + ")")
      : [];
    const reviewMap = new Map(reviews.map((r: any) => [r.attachment_id, r]));
    const documents = attachments.map((attachment: any) => ({
      ...attachment,
      email: messageMap.get(attachment.email_message_id) || null,
      audit: auditMap.get(attachment.id) || null,
      review: reviewMap.get(attachment.id) || null,
    }));

    const divergences = documents.filter((doc: any) => {
      const needsAttention = ["divergent", "needs_review"].includes(doc.audit?.audit_status) || doc.processing_status === "error";
      return needsAttention && doc.review?.review_status !== "completed";
    });

    return NextResponse.json({
      mailbox: connection.mailbox_email,
      emails: messages,
      documents,
      divergences,
      totals: {
        documents: documents.length,
        divergences: divergences.length,
        errors: documents.filter((doc: any) => doc.processing_status === "error").length,
        audited: audits.length,
      },
    });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Falha ao carregar os dados da auditoria." }, { status: 500 });
  }
}
