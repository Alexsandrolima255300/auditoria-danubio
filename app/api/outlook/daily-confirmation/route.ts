import { NextRequest, NextResponse } from "next/server";
import { openToken, refreshMicrosoftToken } from "../../../../lib/microsoft";

export const dynamic = "force-dynamic";
const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";
function headers() {
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error("Configuração segura do banco ausente.");
  return { apikey: key, Authorization: "Bearer " + key, "Content-Type": "application/json" };
}
async function db(path: string, init: RequestInit = {}) {
  const r = await fetch(PROJECT_URL + "/rest/v1/" + path, { ...init, headers: { ...headers(), ...(init.headers || {}) }, cache: "no-store" });
  const body = await r.text();
  if (!r.ok) throw new Error("Banco de dados: " + body.slice(0, 300));
  return body ? JSON.parse(body) : null;
}
function todayBrasilia() {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "America/Sao_Paulo", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
}
async function authorized(request: NextRequest) {
  const cookie = request.cookies.get("outlook_refresh")?.value;
  if (!cookie) return false;
  const token = openToken(cookie);
  await refreshMicrosoftToken(token);
  return true;
}
export async function GET(request: NextRequest) {
  try {
    if (!(await authorized(request))) return NextResponse.json({ error: "Conecte o Outlook para consultar o fechamento diário." }, { status: 401 });
    const date = todayBrasilia();
    const rows = await db("daily_audit_confirmations?select=audit_date,status,confirmed_by,confirmed_at,email_count,audit_count& audit_date=eq." + date);
    return NextResponse.json({ date, confirmation: rows?.[0] || { audit_date: date, status: "pending", confirmed_by: null, confirmed_at: null } });
  } catch (e) { return NextResponse.json({ error: e instanceof Error ? e.message : "Falha ao consultar confirmação." }, { status: 500 }); }
}
export async function POST(request: NextRequest) {
  try {
    if (!(await authorized(request))) return NextResponse.json({ error: "Conecte o Outlook antes de confirmar." }, { status: 401 });
    const body = await request.json().catch(() => ({}));
    const name = String(body.name || "").trim().replace(/\\s+/g, " ");
    if (name.length < 3 || name.length > 100) return NextResponse.json({ error: "Informe o nome completo do responsável (3 a 100 caracteres)." }, { status: 400 });
    const date = todayBrasilia();
    const [emails, audits] = await Promise.all([
      db("email_messages?select=id&received_at=gte." + date + "T00:00:00-03:00&received_at=lt." + date + "T00:00:00-03:00"),
      db("freight_audits?select=id&created_at=gte." + date + "T00:00:00-03:00&created_at=lt." + date + "T00:00:00-03:00")
    ]);
    const saved = await db("daily_audit_confirmations?on_conflict=audit_date", {
      method: "POST", headers: { Prefer: "resolution=merge-duplicates,return=representation" },
      body: JSON.stringify({ audit_date: date, status: "confirmed", confirmed_by: name, confirmed_at: new Date().toISOString(), email_count: emails.length, audit_count: audits.length, updated_at: new Date().toISOString() })
    });
    return NextResponse.json({ ok: true, confirmation: saved?.[0] });
  } catch (e) { return NextResponse.json({ error: e instanceof Error ? e.message : "Falha ao registrar confirmação." }, { status: 500 }); }
}
