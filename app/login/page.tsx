"use client";

import { FormEvent, useState } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import { createClient } from "../../lib/supabase/client";

export default function LoginPage() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const [email, setEmail] = useState("expedicao@brasilengrenagens.com.br");
  const [password, setPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setLoading(true);
    setError("");
    try {
      const supabase = createClient();
      const { error: signInError } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password,
      });
      if (signInError) throw signInError;
      const next = searchParams.get("next");
      router.replace(next && next.startsWith("/") && !next.startsWith("//") ? next : "/");
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Não foi possível entrar. Confira seus dados.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="login-shell">
      <section className="login-visual">
        <div className="login-brand"><span className="login-mark">AF</span><span>ALEX FREIGHT <b>AUDIT</b></span></div>
        <div className="login-message">
          <span className="login-kicker">CONTROLE • CONFERÊNCIA • CONFIANÇA</span>
          <h1>Auditoria de fretes com clareza em cada cobrança.</h1>
          <p>Um ambiente seguro para acompanhar documentos, identificar divergências e registrar revisões.</p>
        </div>
        <div className="login-visual-foot"><span>●</span> Ambiente corporativo protegido</div>
      </section>
      <section className="login-panel-wrap">
        <div className="login-panel">
          <div className="login-mobile-brand"><span className="login-mark">AF</span><span>ALEX FREIGHT <b>AUDIT</b></span></div>
          <span className="login-kicker">BEM-VINDO DE VOLTA</span>
          <h2>Acesse sua conta</h2>
          <p className="login-intro">Entre com o e-mail e a senha cadastrados no sistema.</p>
          <form onSubmit={handleSubmit}>
            <label htmlFor="login-email">E-mail corporativo</label>
            <input id="login-email" type="email" autoComplete="username" value={email} onChange={event => setEmail(event.target.value)} required />
            <label htmlFor="login-password">Senha</label>
            <div className="login-password-wrap">
              <input id="login-password" type={showPassword ? "text" : "password"} autoComplete="current-password" value={password} onChange={event => setPassword(event.target.value)} required />
              <button type="button" className="login-show-password" onClick={() => setShowPassword(value => !value)}>{showPassword ? "Ocultar" : "Mostrar"}</button>
            </div>
            {error && <div className="login-error" role="alert">{error}</div>}
            <button className="login-submit" type="submit" disabled={loading}>{loading ? "Verificando acesso..." : "Entrar no painel"} <span>→</span></button>
          </form>
          <div className="login-security"><span>▣</span><p>Seu acesso é protegido. Não compartilhe sua senha com outras pessoas.</p></div>
          <div className="login-copy">ALEX FREIGHT AUDIT · AUDITORIA DANÚBIO</div>
        </div>
      </section>
    </main>
  );
}
