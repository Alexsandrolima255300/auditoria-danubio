import { createCipheriv, createDecipheriv, createHash, randomBytes } from "crypto";

export function microsoftConfig() {
  const clientId = process.env.MICROSOFT_CLIENT_ID;
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET;
  const tenantId = process.env.MICROSOFT_TENANT_ID;
  const redirectUri = process.env.MICROSOFT_REDIRECT_URI;
  if (!clientId || !clientSecret || !tenantId || !redirectUri) {
    throw new Error("Configuração Microsoft incompleta. Confira MICROSOFT_CLIENT_ID, MICROSOFT_CLIENT_SECRET, MICROSOFT_TENANT_ID e MICROSOFT_REDIRECT_URI na Vercel.");
  }
  return { clientId, clientSecret, tenantId, redirectUri };
}

function encryptionKey() {
  const { clientSecret, tenantId } = microsoftConfig();
  return createHash("sha256").update(clientSecret + ":" + tenantId + ":ALEX_FREIGHT_AUDIT").digest();
}

export function sealToken(token: string) {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", encryptionKey(), iv);
  const encrypted = Buffer.concat([cipher.update(token, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, encrypted]).toString("base64url");
}

export function openToken(value: string) {
  const data = Buffer.from(value, "base64url");
  if (data.length < 29) throw new Error("Sessão do Outlook inválida.");
  const decipher = createDecipheriv("aes-256-gcm", encryptionKey(), data.subarray(0, 12));
  decipher.setAuthTag(data.subarray(12, 28));
  return Buffer.concat([decipher.update(data.subarray(28)), decipher.final()]).toString("utf8");
}

export async function refreshMicrosoftToken(refreshToken: string) {
  const { clientId, clientSecret, tenantId } = microsoftConfig();
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: "openid profile offline_access User.Read Mail.Read",
  });
  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
    cache: "no-store",
  });
  const data = await response.json();
  if (!response.ok || !data.access_token) {
    throw new Error(data.error_description || "Não foi possível renovar a conexão com a Microsoft.");
  }
  return data as { access_token: string; refresh_token?: string; expires_in: number };
}
