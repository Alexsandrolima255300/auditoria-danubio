import { NextRequest, NextResponse } from "next/server";
import { microsoftConfig, sealToken } from "../../../../../lib/microsoft";

export const runtime = "nodejs";

const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";

async function saveConnection(mailbox: string, refreshToken: string, expiresIn: number) {
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
  if (!key) throw new Error("Falta configurar SUPABASE_SERVICE_ROLE_KEY na Vercel.");
  const headers = { apikey: key, Authorization: "Bearer " + key, "Content-Type": "application/json" };
  const existingResponse = await fetch(
    PROJECT_URL + "/rest/v1/outlook_connections?select=id&mailbox_email=eq." + encodeURIComponent(mailbox) + "&limit=1",
    { headers, cache: "no-store" }
  );
  const existingText = await existingResponse.text();
  if (!existingResponse.ok) throw new Error("Não foi possível atualizar a conexão no banco: " + existingText.slice(0, 250));
  const existing = existingText ? JSON.parse(existingText) : [];
  const payload = {
    provider: "microsoft_graph",
    mailbox_email: mailbox,
    tenant_id: process.env.MICROSOFT_TENANT_ID,
    encrypted_token_cache: sealToken(refreshToken),
    token_expires_at: new Date(Date.now() + expiresIn * 1000).toISOString(),
    granted_scopes: ["openid", "profile", "offline_access", "User.Read", "Mail.Read"],
    last_refreshed_at: new Date().toISOString(),
    revoked_at: null,
  };
  const url = PROJECT_URL + "/rest/v1/outlook_connections" + (existing?.[0]?.id ? "?id=eq." + encodeURIComponent(existing[0].id) : "");
  const saved = await fetch(url, {
    method: existing?.[0]?.id ? "PATCH" : "POST",
    headers: { ...headers, Prefer: "return=minimal" },
    body: JSON.stringify(payload),
    cache: "no-store",
  });
  if (!saved.ok) throw new Error("Não foi possível salvar a nova autorização do Outlook: " + (await saved.text()).slice(0, 250));
}

export async function GET(request: NextRequest) {
  const { searchParams, origin } = new URL(request.url);
  const code = searchParams.get("code");
  const state = searchParams.get("state");
  const expectedState = request.cookies.get("outlook_oauth_state")?.value;
  const microsoftError = searchParams.get("error_description") || searchParams.get("error");
  if (microsoftError) return NextResponse.redirect(new URL(`/?outlook_error=${encodeURIComponent(microsoftError)}`, origin));
  if (!code || !state || !expectedState || state !== expectedState) {
    return NextResponse.redirect(new URL("/?outlook_error=Falha+na+validacao+de+seguranca.+Tente+conectar+novamente.", origin));
  }

  try {
    const { clientId, clientSecret, tenantId, redirectUri } = microsoftConfig();
    const body = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri,
      scope: "openid profile offline_access User.Read Mail.Read",
    });
    const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body, cache: "no-store",
    });
    const tokens = await tokenResponse.json();
    if (!tokenResponse.ok || !tokens.refresh_token || !tokens.access_token) {
      throw new Error(tokens.error_description || "A Microsoft não retornou uma nova autorização para o Outlook.");
    }

    const profileResponse = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName", {
      headers: { Authorization: "Bearer " + tokens.access_token }, cache: "no-store",
    });
    const profile = await profileResponse.json();
    if (!profileResponse.ok) throw new Error(profile.error?.message || "Não foi possível validar a caixa corporativa.");
    const mailbox = profile.mail || profile.userPrincipalName;
    if (!mailbox) throw new Error("Não foi possível identificar a caixa de e-mail conectada.");

    // Persistir a nova credencial no Supabase aqui é essencial: o cron não possui
    // os cookies do navegador e precisa usar o refresh token atualizado do banco.
    await saveConnection(mailbox, tokens.refresh_token, Number(tokens.expires_in || 3600));

    const response = NextResponse.redirect(new URL("/?outlook_connected=1", origin));
    response.cookies.set("outlook_refresh", sealToken(tokens.refresh_token), {
      httpOnly: true, secure: process.env.NODE_ENV === "production", sameSite: "lax", maxAge: 60 * 60 * 24 * 30, path: "/",
    });
    response.cookies.set("outlook_oauth_state", "", { httpOnly: true, maxAge: 0, path: "/" });
    return response;
  } catch (error) {
    return NextResponse.redirect(new URL(`/?outlook_error=${encodeURIComponent(error instanceof Error ? error.message : "Falha na conexão com Outlook.")}`, origin));
  }
}
