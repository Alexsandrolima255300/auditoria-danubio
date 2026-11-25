import { createServerClient } from "@supabase/ssr";
import { NextRequest, NextResponse } from "next/server";

function isAuthorizedCron(request: NextRequest) {
  const secret = process.env.CRON_SECRET;
  if (!secret || request.headers.get("authorization") !== "Bearer " + secret) return false;
  return (
    (request.nextUrl.pathname === "/api/outlook/audit" && request.method === "GET") ||
    (request.nextUrl.pathname === "/api/retention/cleanup" && request.method === "GET")
  );
}

export async function middleware(request: NextRequest) {
  const pathname = request.nextUrl.pathname;
  const isPublicPage = pathname === "/login";
  const isMicrosoftCallback = pathname === "/api/auth/microsoft/callback";
  if (isMicrosoftCallback || isAuthorizedCron(request)) return NextResponse.next();

  const url = process.env.NEXT_PUBLIC_SUPABASE_URL || process.env.SUPABASE_URL;
  const key = process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY;
  if (!url || !key) {
    if (pathname.startsWith("/api/")) {
      return NextResponse.json({ error: "Autenticação indisponível: configure as variáveis públicas do Supabase." }, { status: 503 });
    }
    return new NextResponse("Autenticação indisponível: configure as variáveis públicas do Supabase.", { status: 503 });
  }

  let response = NextResponse.next({ request });
  const supabase = createServerClient(url, key, {
    cookies: {
      getAll() {
        return request.cookies.getAll();
      },
      setAll(cookiesToSet, headers) {
        cookiesToSet.forEach(({ name, value }) => request.cookies.set(name, value));
        response = NextResponse.next({ request });
        cookiesToSet.forEach(({ name, value, options }) => response.cookies.set(name, value, options));
        Object.entries(headers).forEach(([name, value]) => response.headers.set(name, value));
      },
    },
  });

  const { data, error } = await supabase.auth.getClaims();
  const isAuthenticated = !error && Boolean(data?.claims);

  if (!isAuthenticated) {
    if (pathname.startsWith("/api/")) {
      return NextResponse.json({ error: "Sessão não autenticada. Entre novamente." }, { status: 401 });
    }
    if (!isPublicPage) {
      const loginUrl = request.nextUrl.clone();
      loginUrl.pathname = "/login";
      loginUrl.searchParams.set("next", pathname);
      return NextResponse.redirect(loginUrl);
    }
    return response;
  }

  if (isPublicPage) {
    const destination = request.nextUrl.clone();
    const next = destination.searchParams.get("next");
    destination.pathname = next && next.startsWith("/") && !next.startsWith("//") ? next : "/";
    destination.search = "";
    return NextResponse.redirect(destination);
  }

  return response;
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico|.*\\.(?:svg|png|jpg|jpeg|gif|webp)$).*)"],
};
