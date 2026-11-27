import { NextRequest, NextResponse } from "next/server";
import { randomBytes } from "crypto";
import { microsoftConfig } from "../../../../lib/microsoft";

export async function GET(request: NextRequest) {
  try {
    const { clientId, tenantId, redirectUri } = microsoftConfig();
    const state = randomBytes(24).toString("hex");
    const url = new URL(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize`);
    url.search = new URLSearchParams({
      client_id: clientId,
      response_type: "code",
      redirect_uri: redirectUri,
      response_mode: "query",
      scope: "openid profile offline_access User.Read Mail.Read",
      state,
      prompt: "consent",
    }).toString();
    const response = NextResponse.redirect(url);
    response.cookies.set("outlook_oauth_state", state, {
      httpOnly: true, secure: process.env.NODE_ENV === "production", sameSite: "lax", maxAge: 600, path: "/",
    });
    return response;
  } catch (error) {
    return NextResponse.json({ error: error instanceof Error ? error.message : "Falha ao iniciar autenticação." }, { status: 500 });
  }
}
