import { NextRequest, NextResponse } from "next/server";

export const dynamic = "force-dynamic";
const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";

async function db(path: string, method = "GET", payload?: unknown) {
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error("Falta cadastrar SUPABASE_SERVICE_ROLE_KEY na Vercel.");
  const response = await fetch(PROJECT_URL + "/rest/v1/" + path, {
    method,
    headers: {
      apikey: key,
      Authorization: "Bearer " + key,
      "Content-Type": "application/json",
      ...(method !== "GET" ? { Prefer: "resolution=merge-duplicates,return=representation" } : {}),
    },
    ...(payload ? { body: JSON.stringify(payload) } : {}),
    cache: "no-store",
  });
  const body = await response.text();
  if (!response.ok) throw new Error("Supabase: " + body.slice(0, 400));
  return body ? JSON.parse(body) : [];
}

export async function POST(request: NextRequest) {
  if (!request.cookies.get("outlook_refresh")?.value) {
    return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });
  }
  try {
    const body = await request.json();
    const attachmentId = String(body.attachmentId || "");
    const reviewStatus = String(body.reviewStatus || "pending");
    const assessment = body.assessment ? String(body.assessment) : null;
    const reason = String(body.reason || "").trim();
    const reviewedBy = String(body.reviewedBy || "").trim();
    if (!/^[0-9a-f-]{36}$/i.test(attachmentId)) return NextResponse.json({ error: "Ocorrência inválida." }, { status: 400 });
    if (!["pending", "in_progress", "completed"].includes(reviewStatus)) return NextResponse.json({ error: "Situação de revisão inválida." }, { status: 400 });
    if (assessment && !["correct_charge", "incorrect_charge", "unconfirmed"].includes(assessment)) return NextResponse.json({ error: "Avaliação inválida." }, { status: 400 });
    if (reviewedBy.length > 100 || reason.length > 3000) return NextResponse.json({ error: "O responsável ou a justificativa excede o limite permitido." }, { status: 400 });

    const connections = await db("outlook_connections?select=id&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if (!connections.length) return NextResponse.json({ error: "Outlook ainda não conectado." }, { status: 401 });
    const attachments = await db("freight_attachments?select=id,email_message_id&id=eq." + attachmentId + "&limit=1");
    if (!attachments.length) return NextResponse.json({ error: "Ocorrência não encontrada." }, { status: 404 });
    const messages = await db("email_messages?select=id,connection_id&id=eq." + attachments[0].email_message_id + "&limit=1");
    if (!messages.length || messages[0].connection_id !== connections[0].id) {
      return NextResponse.json({ error: "Você não tem acesso a esta ocorrência." }, { status: 403 });
    }

    const rows = await db("freight_divergence_reviews?on_conflict=attachment_id", "POST", {
      attachment_id: attachmentId,
      review_status: reviewStatus,
      assessment,
      reason: reason || null,
      reviewed_by: reviewedBy || null,
      reviewed_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    });
    return NextResponse.json({ review: rows[0] || null });
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Não foi possível salvar a revisão." }, { status: 500 });
  }
}
