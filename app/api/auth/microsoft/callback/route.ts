import { NextRequest, NextResponse } from "next/server";
import { microsoftConfig, sealToken } from "../../../../../lib/microsoft";

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
    if (!tokenResponse.ok || !tokens.refresh_token) {
      throw new Error(tokens.error_description || "A Microsoft não retornou autorização para acesso contínuo ao Outlook.");
    }
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
