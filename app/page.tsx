'use client';

import { useEffect, useState } from 'react';

const brl = (n: number | null | undefined) => typeof n === 'number' && Number.isFinite(n) ? n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : 'Não informado';
const onlyDigits = (s: string) => s.replace(/\D/g, '');
const cents = (n: number) => Math.round((n + Number.EPSILON) * 100) / 100;

type Mail = { id: string; subject: string; sender: string; senderEmail: string; receivedAt: string; hasAttachments: boolean; preview: string; webLink: string };

export default function Home() {
  const [peso, setPeso] = useState('2,88');
  const [nota, setNota] = useState('697,99');
  const [cobrado, setCobrado] = useState('88,00');
  const [cnpj, setCnpj] = useState('39.860.057/0001-04');
  const [resultado, setResultado] = useState(false);
  const [emails, setEmails] = useState<Mail[]>([]);
  const [mailLoading, setMailLoading] = useState(false);
  const [mailError, setMailError] = useState('');
  const [connected, setConnected] = useState(false);
  const [mailCount, setMailCount] = useState<number | null>(null);
  const [auditLoading, setAuditLoading] = useState(false);
  const [auditResult, setAuditResult] = useState<any>(null);
  const [auditCursor, setAuditCursor] = useState<any>({ cursorDanubio: null, cursorSiimp: null, doneDanubio: false, doneSiimp: false });
  const [auditProgress, setAuditProgress] = useState({ processed: 0, percent: 0, running: false });
  const [responsibleName, setResponsibleName] = useState('');
  const [dailyConfirmation, setDailyConfirmation] = useState<any>(null);
  const [confirmationError, setConfirmationError] = useState('');
  const [confirmationLoading, setConfirmationLoading] = useState(false);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get('outlook_connected') === '1') {
      setConnected(true);
      window.history.replaceState({}, '', '/');
      void syncEmails();
    }
    if (params.get('outlook_error')) {
      setMailError(params.get('outlook_error') || 'Falha ao conectar.');
      window.history.replaceState({}, '', '/');
    }
  }, []);

  async function loadDailyConfirmation() {
    try { const r = await fetch('/api/outlook/daily-confirmation', { cache: 'no-store' }); const d = await r.json(); if (r.ok) setDailyConfirmation(d.confirmation); } catch {}
  }

  async function confirmDailyAudit() {
    setConfirmationLoading(true); setConfirmationError('');
    try { const r = await fetch('/api/outlook/daily-confirmation', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: responsibleName }) }); const d = await r.json(); if (!r.ok) throw new Error(d.error || 'Não foi possível confirmar.'); setDailyConfirmation(d.confirmation); setResponsibleName(''); }
    catch (e) { setConfirmationError(e instanceof Error ? e.message : 'Falha ao confirmar.'); } finally { setConfirmationLoading(false); }
  }

  async function syncEmails() {
    setMailLoading(true);
    setMailError('');
    try {
      const response = await fetch('/api/outlook/messages', { cache: 'no-store' });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Não foi possível consultar os e-mails.');
      setConnected(true);
      setEmails(data.emails || []);
      setMailCount(data.count ?? 0);
      await loadDailyConfirmation();
    } catch (error) {
      setMailError(error instanceof Error ? error.message : 'Erro ao consultar Outlook.');
      if ((error instanceof Error ? error.message : '').includes('não conectado')) setConnected(false);
    } finally {
      setMailLoading(false);
    }
  }


  useEffect(() => { if (!connected) return; const timer = window.setInterval(() => { void syncEmails(); }, 60000); return () => window.clearInterval(timer); }, [connected]);

  async function runAudit(reprocessProcessed = false) {
    setAuditLoading(true); setMailError(''); setAuditResult(null);
    setAuditProgress({ processed: 0, percent: 0, running: true });
    // A varredura normal retoma pelo checkpoint salvo no Supabase.
    // O cursor local é usado somente na releitura manual de PDFs antigos.
    let cursor = { cursorDanubio: null, cursorSiimp: null, doneDanubio: false, doneSiimp: false };
    let totals = { emailsFound: 0, pdfsRead: 0, auditsCreated: 0, errors: 0, ignored: 0 };
    let allResults: any[] = [];
    try {
      // Cada chamada processa um lote pequeno. A API persiste o próximo link no Supabase,
      // então a execução pode continuar mesmo depois de recarregar a página.
      for (let batch = 0; batch < 500; batch++) {
        const response = await fetch('/api/outlook/audit', {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ ...cursor, reprocessProcessed }), cache: 'no-store'
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Falha ao auditar e-mails.');
        totals.emailsFound += data.emailsFound || 0;
        totals.pdfsRead += data.pdfsRead || 0;
        totals.auditsCreated += data.auditsCreated || 0;
        totals.errors += data.errors || 0;
        totals.ignored += data.ignored || 0;
        allResults = [...allResults, ...(data.results || [])];
        cursor = {
          cursorDanubio: data.nextDanubio || null,
          cursorSiimp: data.nextSiimp || null,
          doneDanubio: data.doneDanubio === true,
          doneSiimp: data.doneSiimp === true,
        };
        setAuditCursor(cursor);
        const percent = data.hasMore ? Math.min(95, Math.max(5, Math.round((totals.emailsFound / Math.max(mailCount || 1, totals.emailsFound + 1)) * 100))) : 100;
        setAuditProgress({ processed: totals.emailsFound, percent, running: Boolean(data.hasMore) });
        setAuditResult({ ...totals, results: allResults, hasMore: Boolean(data.hasMore) });
        setConnected(true);
        if (!data.hasMore) break;
        if (batch === 499) throw new Error('A busca atingiu o limite de segurança de lotes. Execute novamente para continuar de onde parou.');
      }
      setAuditResult({ ...totals, results: allResults, hasMore: false });
      setAuditProgress({ processed: totals.emailsFound, percent: 100, running: false });
    } catch (error) {
      setMailError(error instanceof Error ? error.message : 'Erro na auditoria.');
      setAuditResult({ ...totals, results: allResults, hasMore: true });
      setAuditProgress(p => ({ ...p, running: false }));
    } finally { setAuditLoading(false); }
  }

  const kg = Number(peso.replace(',', '.')) || 0;
  const nf = Number(nota.replace(',', '.')) || 0;
  const valorCobrado = cents(Number(cobrado.replace(',', '.')) || 0);
  const pesoCalc = cents(kg * 0.7);
  const notaCalc = cents(nf * 0.015);
  const esperado = Math.max(100, pesoCalc, notaCalc);
  const criterio = esperado === 100 ? 'Referência contratual de R$ 100,00' : pesoCalc >= notaCalc ? 'R$ 0,70 por kg' : '1,5% sobre o valor da nota';
  const tomadorValido = onlyDigits(cnpj) === '39860057000104';
  const abaixoDoMinimoAceito = valorCobrado <= 100;
  const divergente = tomadorValido && !abaixoDoMinimoAceito && Math.abs(valorCobrado - esperado) >= 0.01;
  const diferenca = cents(Math.abs(valorCobrado - esperado));
  const situacaoValor = valorCobrado > esperado ? 'acima do valor de referência' : 'abaixo do valor de referência';

  return <main className="page">
    <aside className="sidebar"><div className="logo">AF</div><div className="brand">ALEX FREIGHT <b>AUDIT</b></div><p className="menu-label">MENU</p><div className="menu-item selected">Visão geral</div><div className="menu-item">Documentos</div><div className="menu-item">Divergências</div><div className="menu-item">E-mails</div><div className="side-note">Auditoria Danúbio<br/>Revisão humana antes de contestar</div></aside>
    <section className="content"><header><div><span className="overline">PAINEL DE CONTROLE</span><h1>Auditoria de fretes</h1><p>Confira as cobranças da transportadora Danúbio.</p></div><span className="status">● Transportadora: Danúbio</span></header>
      <div className="cards"><article><small>Transportadoras monitoradas</small><strong>01</strong><span>Danúbio</span></article><article><small>Regra de cálculo</small><strong>R$ 0,70</strong><span>Por kg ou 1,5% da nota</span></article><article><small>Referência contratual</small><strong>R$ 100,00</strong><span>Cobranças até esse valor são aceitas</span></article></div>
      <section className="mail-panel mail-live"><div className="mail-icon">✉</div><div className="mail-main"><h3>Conexão com Microsoft Outlook</h3><p>Conecte a caixa corporativa para localizar mensagens do Siimp Sistemas e e-mails que mencionem “Danubio”. A consulta é somente leitura.</p><span className="pill">Siimp Sistemas</span> <span className="pill">Palavra-chave: Danubio</span> <span className="pill">Somente leitura</span>
        {mailError && <div className="mail-error">{mailError}</div>}
        {connected && <div className="mail-success">Conta conectada. {mailCount === null ? '' : `${mailCount} mensagem(ns) localizada(s).`}</div>}
        <div className="daily-review">
          <div className="daily-review-head"><div><h3>Conferência humana do dia</h3><p>Informe o nome do responsável somente depois de revisar as auditorias e pendências.</p></div><span className={dailyConfirmation?.status === 'confirmed' ? 'daily-state done' : 'daily-state'}>{dailyConfirmation?.status === 'confirmed' ? 'Confirmado' : 'Pendente'}</span></div>
          {dailyConfirmation?.status === 'confirmed' ? <div className="mail-success">Conferido por <b>{dailyConfirmation.confirmed_by}</b> em {new Date(dailyConfirmation.confirmed_at).toLocaleString('pt-BR')}.</div> : <><label className="daily-label">Nome do responsável</label><input className="daily-input" value={responsibleName} onChange={e=>setResponsibleName(e.target.value)} placeholder="Digite o nome completo" maxLength={100}/><button className="daily-confirm" onClick={()=>void confirmDailyAudit()} disabled={!connected || confirmationLoading || responsibleName.trim().length < 3}>{confirmationLoading ? 'Registrando confirmação...' : 'Confirmar auditoria do dia'}</button>{confirmationError && <div className="mail-error">{confirmationError}</div>}</>}
          <small className="daily-foot">A confirmação registra o responsável e o horário. Novos e-mails exigem nova conferência.</small>
        </div>
        <div className="mail-actions">{!connected && <button onClick={()=>window.location.href='/api/outlook/connect'}>Conectar Microsoft Outlook</button>}<button className="light" onClick={()=>void syncEmails()} disabled={mailLoading}>{mailLoading ? 'Consultando mensagens...' : connected ? 'Atualizar e-mails' : 'Testar conexão'}</button><button onClick={()=>void runAudit()} disabled={!connected || auditLoading}>{auditLoading ? 'Auditando e-mails...' : auditResult?.hasMore ? 'Continuar varredura' : 'Auditar e-mails e CT-e'}</button><button className="light" onClick={()=>{ if(window.confirm('Isso vai baixar e ler novamente os PDFs já processados. Deseja continuar?')) void runAudit(true); }} disabled={!connected || auditLoading}>Reprocessar PDFs já lidos</button></div>{(auditLoading || auditResult) && <div className="mail-list" style={{marginTop:16}}>
          <h4>{auditLoading ? 'Auditoria em andamento' : 'Progresso da auditoria'}</h4>
          <div style={{height:10,background:'#e5e7eb',borderRadius:20,overflow:'hidden',margin:'10px 0'}}>
            <div style={{height:'100%',width:auditProgress.percent+'%',background:'#2563eb',transition:'width .25s ease'}} />
          </div>
          <small>{auditProgress.percent}% · {auditProgress.processed} e-mail(s) processado(s){auditLoading ? ' · lendo próximos lotes automaticamente...' : ' · processamento pausado/concluído'}</small>
        </div>}
        {auditResult && <div className="mail-success">Varredura: {auditResult.emailsFound} e-mails, {auditResult.pdfsRead} PDFs lidos, {auditResult.auditsCreated} auditorias salvas, {auditResult.errors} erros. {auditResult.hasMore ? 'Ainda há mensagens pendentes; clique em Auditar para retomar.' : 'Busca concluída.'}</div>}
        {auditResult?.results?.some((item:any)=>['divergent','error','needs_review'].includes(item.status)) && <div className="mail-error" role="alert"><b>Notificações de auditoria</b>{auditResult.results.filter((item:any)=>['divergent','error','needs_review'].includes(item.status)).map((item:any,i:number)=><div key={i} style={{marginTop:10,paddingTop:8,borderTop:'1px solid #ddd'}}>
          <b>{item.status === 'divergent' ? 'DIVERGÊNCIA' : item.status === 'error' ? 'ERRO NA LEITURA' : 'REVISÃO NECESSÁRIA'} — {item.file || item.email}</b>
          <div>{item.reason}</div>
          {item.invoiceValue != null && <div>Valor da NF: {brl(item.invoiceValue)}</div>}
          {item.weightKg != null && <div>Peso: {item.weightKg} kg</div>}
          {item.tomadorCnpj && <div>CNPJ do tomador: {item.tomadorCnpj}</div>}
          {item.freightGross != null && <div>Frete total: {brl(item.freightGross)} · ICMS: {brl(item.icms ?? 0)} · Frete sem ICMS: {brl(item.freight ?? 0)}</div>}
          {item.reference != null && <div>Referência: {brl(item.reference)} · Diferença: {brl(item.discrepancy ?? 0)}</div>}
        </div>)}</div>}
        {auditResult?.results?.length > 0 && <div className="mail-list"><h4>CT-e processados</h4>{auditResult.results.map((item:any,i:number)=><article className="mail-item" key={i}><div><b>{item.file || item.email}</b><small>{item.reason}</small>{item.invoiceValue != null && <small>NF: {brl(item.invoiceValue)} · Peso: {item.weightKg ?? 'Não identificado'} kg · Tomador: {item.tomadorCnpj || 'Não identificado'}</small>}</div><div className="mail-meta"><span className="pill">{item.status}</span>{item.freightGross !== undefined && <small>Frete total: {brl(item.freightGross)} · ICMS: {brl(item.icms ?? 0)} · Frete sem ICMS: {brl(item.freight)}</small>}</div></article>)}</div>}
        {emails.length > 0 && <div className="mail-list"><h4>E-mails já sincronizados · lista atualizada a cada minuto</h4>{emails.map(mail=><article key={mail.id} className="mail-item"><div><b>{mail.subject}</b><small>{mail.sender} · {new Date(mail.receivedAt).toLocaleString('pt-BR')}</small><p>{mail.preview}</p></div><div className="mail-meta">{mail.hasAttachments && <span className="pill">Com anexo</span>}{mail.webLink && <a href={mail.webLink} target="_blank" rel="noreferrer">Abrir e-mail</a>}</div></article>)}</div>}
      </div><span className={connected ? 'not-connected connected' : 'not-connected'}>{connected ? 'Conectado' : 'Aguardando conexão'}</span></section>
      <footer>ALEX FREIGHT AUDIT · Relatórios de divergência sujeitos à revisão humana.</footer>
    </section>
  </main>;
}
