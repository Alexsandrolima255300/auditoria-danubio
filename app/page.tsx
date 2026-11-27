'use client';

import { useCallback, useEffect, useState } from 'react';
import LogoutButton from '../components/logout-button';

const brl = (n: number | string | null | undefined) => {
  const value = typeof n === 'string' ? Number(n) : n;
  return typeof value === 'number' && Number.isFinite(value) ? value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : 'Não informado';
};
const dateTime = (value?: string | null) => value ? new Date(value).toLocaleString('pt-BR') : 'Data não informada';
const textField = (text: string, pattern: RegExp) => {
  const match = text.match(pattern);
  return match?.[1]?.trim() || 'Não identificado no PDF';
};
const statusLabel = (status?: string) => (({ accepted: 'Aceito', divergent: 'Divergência', needs_review: 'Revisão necessária', out_of_scope: 'Fora do escopo', pending: 'Pendente', extracted: 'Extraído', error: 'Erro na leitura', discovered: 'Localizado', attachments_checked: 'Anexos conferidos' } as Record<string, string>)[status || ''] || status || 'Não informado');

type Section = 'overview' | 'documents' | 'divergences' | 'newEmails' | 'emails' | 'search' | 'invoices';
type Mail = { id: string; subject: string; sender: string; senderEmail: string; receivedAt: string; hasAttachments: boolean; preview: string; webLink: string };
type AuditData = { mailbox?: string; emails: any[]; documents: any[]; completedDocuments?: any[]; divergences: any[]; totals: { documents: number; divergences: number; errors: number; audited: number } };
const navigation: { id: Section; label: string; icon: string }[] = [
  { id: 'overview', label: 'Visão geral', icon: '▦' },
  { id: 'documents', label: 'Documentos', icon: '▤' },
  { id: 'divergences', label: 'Divergências', icon: '⚑' },
  { id: 'newEmails', label: 'Novos e-mails', icon: '✉' },
  { id: 'emails', label: 'E-mails', icon: '▣' },
  { id: 'search', label: 'Pesquisa', icon: '⌕' },
  { id: 'invoices', label: 'Faturas', icon: '▧' },
];

export default function Home() {
  const [activeSection, setActiveSection] = useState<Section>('overview');
  const [emails, setEmails] = useState<Mail[]>([]);
  const [todayEmails, setTodayEmails] = useState<Mail[]>([]);
  const [todayEmailCount, setTodayEmailCount] = useState(0);
  const [brazilDate, setBrazilDate] = useState('');
  const [mailLoading, setMailLoading] = useState(false);
  const [mailError, setMailError] = useState('');
  const [connected, setConnected] = useState(false);
  const [mailCount, setMailCount] = useState<number | null>(null);
  const [emailFolderState, setEmailFolderState] = useState<Record<string, boolean>>({});
  const [searchDate, setSearchDate] = useState('');
  const [searchCte, setSearchCte] = useState('');
  const [auditLoading, setAuditLoading] = useState(false);
  const [auditResult, setAuditResult] = useState<any>(null);
  const [auditProgress, setAuditProgress] = useState({ processed: 0, percent: 0, running: false });
  const [responsibleName, setResponsibleName] = useState('');
  const [dailyConfirmation, setDailyConfirmation] = useState<any>(null);
  const [confirmationError, setConfirmationError] = useState('');
  const [confirmationLoading, setConfirmationLoading] = useState(false);
  const [savedData, setSavedData] = useState<AuditData | null>(null);
  const [dataLoading, setDataLoading] = useState(false);
  const [initialLoading, setInitialLoading] = useState(true);
  const [dataError, setDataError] = useState('');
  const [selectedReviewId, setSelectedReviewId] = useState<string | null>(null);
  const [reviewDraft, setReviewDraft] = useState({ reviewStatus: 'pending', assessment: '', reason: '', reviewedBy: '' });
  const [reviewSaving, setReviewSaving] = useState(false);
  const [reviewError, setReviewError] = useState('');
  const [reviewSuccess, setReviewSuccess] = useState('');
  const [invoiceData, setInvoiceData] = useState<any[]>([]);
  const [invoiceLoading, setInvoiceLoading] = useState(false);
  const [invoiceProgress, setInvoiceProgress] = useState({ processed: 0, total: 0, percent: 0 });
  const [invoiceError, setInvoiceError] = useState('');
  const [invoiceRecoveryId, setInvoiceRecoveryId] = useState<string | null>(null);
  const [invoiceRecoveryStatusId, setInvoiceRecoveryStatusId] = useState<string | null>(null);
  const [invoiceRecoveryStatus, setInvoiceRecoveryStatus] = useState('');
    const [invoiceFoldersOpen, setInvoiceFoldersOpen] = useState<Record<string, boolean>>({});
const [selectedInvoiceId, setSelectedInvoiceId] = useState<string | null>(null);
  const [selectedInvoiceDiscrepancy, setSelectedInvoiceDiscrepancy] = useState<any>(null);

  const loadAuditData = useCallback(async () => {
    setDataLoading(true);
    setDataError('');
    try {
      const response = await fetch('/api/outlook/audit-data', { cache: 'no-store' });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Não foi possível carregar os registros.');
      setSavedData(data);
    } catch (error) {
      setDataError(error instanceof Error ? error.message : 'Falha ao carregar documentos e divergências.');
    } finally {
      setDataLoading(false);
    }
  }, []);

  async function loadInvoices() {
    setInvoiceLoading(true); setInvoiceError('');
    try {
      const response = await fetch('/api/outlook/invoices', { method: 'GET', cache: 'no-store' });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Não foi possível conferir as faturas.');
      setInvoiceData(data.invoices || []);
    } catch (error) { setInvoiceError(error instanceof Error ? error.message : 'Falha ao carregar faturas.'); }
    finally { setInvoiceLoading(false); }
  }

  async function scanInvoices() {
    setInvoiceLoading(true); setInvoiceError('');
    setInvoiceProgress({ processed: 0, total: 0, percent: 0 });
    try {
      let processed = 0;
      let total = 0;
      let keepGoing = true;
      while (keepGoing) {
        const response = await fetch('/api/outlook/invoices?batch=5', { method: 'POST', cache: 'no-store' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Não foi possível processar as faturas.');
        processed += Number(data.scannedEmails || 0);
        total = processed + Number(data.remainingEmails || 0);
        setInvoiceProgress({ processed, total, percent: total ? Math.round(processed / total * 100) : 100 });
        setInvoiceData(data.invoices || []);
        keepGoing = Boolean(data.hasMore);
      }
    } catch (error) { setInvoiceError(error instanceof Error ? error.message : 'Falha ao processar faturas.'); }
    finally { setInvoiceLoading(false); }
  }

  async function recoverInvoiceCtes(invoice: any) {
    setInvoiceRecoveryId(invoice.id);
    setInvoiceRecoveryStatusId(invoice.id);
    setInvoiceRecoveryStatus('Iniciando busca histórica no Outlook...');
    setInvoiceError('');
    try {
      let daysDone = 0;
      let emailsFound = 0;
      let pdfsRead = 0;
      let auditsCreated = 0;
      let errors = 0;
      const start = new Date('2026-10-16T12:00:00Z');
      const end = new Date('2026-11-16T12:00:00Z');
      for (let day = new Date(start); day <= end; day.setUTCDate(day.getUTCDate() + 1)) {
        const receivedDate = day.toISOString().slice(0, 10);
        let dateNextLink: string | null = null;
        let page = 0;
        do {
          const response = await fetch('/api/outlook/audit', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ receivedDate, dateNextLink }),
            cache: 'no-store',
          });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error || 'Falha ao buscar e-mails históricos.');
          emailsFound += Number(data.emailsFound || 0);
          pdfsRead += Number(data.pdfsRead || 0);
          auditsCreated += Number(data.auditsCreated || 0);
          errors += Number(data.errors || 0);
          dateNextLink = data.dateNextLink || null;
          page++;
          if (page > 100) throw new Error('A busca histórica retornou páginas demais para um único dia.');
        } while (dateNextLink);
        daysDone++;
        setInvoiceRecoveryStatus(`Buscando e-mails: ${daysDone}/32 dias · ${pdfsRead} PDFs lidos · ${auditsCreated} auditorias registradas...`);
      }

      setInvoiceRecoveryStatus('Documentos recuperados. Refazendo a comparação dos 162 CT-es...');
      const reconcileResponse = await fetch(`/api/outlook/invoices/${encodeURIComponent(invoice.id)}/reconcile`, {
        method: 'POST', cache: 'no-store',
      });
      const reconcile = await reconcileResponse.json();
      if (!reconcileResponse.ok) throw new Error(reconcile.error || 'Falha ao refazer a comparação da fatura.');
      const refreshed = await fetch('/api/outlook/invoices', { cache: 'no-store' });
      const refreshedData = await refreshed.json();
      if (refreshed.ok) setInvoiceData(refreshedData.invoices || []);
      setInvoiceRecoveryStatus(`Busca concluída: ${pdfsRead} PDFs lidos, ${auditsCreated} auditorias criadas, ${reconcile.recoveredOriginals} CT-es originais vinculados e ${reconcile.statuses?.not_found || 0} ainda não localizados. ${errors ? errors + ' erro(s) de leitura.' : ''}`);
    } catch (error) {
      setInvoiceError(error instanceof Error ? error.message : 'Falha na recuperação histórica.');
      setInvoiceRecoveryStatus('');
    } finally {
      setInvoiceRecoveryId(null);
    }
  }

  useEffect(() => {
    if (activeSection === 'invoices') void loadInvoices();
  }, [activeSection]);

  async function loadDailyConfirmation() {
    try {
      const response = await fetch('/api/outlook/daily-confirmation', { cache: 'no-store' });
      const data = await response.json();
      if (response.ok) setDailyConfirmation(data.confirmation);
    } catch {}
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
      setTodayEmails(data.todayEmails || []);
      setTodayEmailCount(data.todayCount ?? 0);
      setBrazilDate(data.brazilDate || '');
      setMailCount(data.count ?? 0);
      await Promise.all([loadDailyConfirmation(), loadAuditData()]);
    } catch (error) {
      setMailError(error instanceof Error ? error.message : 'Erro ao consultar Outlook.');
      if ((error instanceof Error ? error.message : '').toLowerCase().includes('não conectado')) setConnected(false);
    } finally {
      setMailLoading(false);
      setInitialLoading(false);
    }
  }

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get('outlook_error')) setMailError(params.get('outlook_error') || 'Falha ao conectar.');
    if (params.get('outlook_connected') === '1' || params.get('outlook_error')) window.history.replaceState({}, '', '/');
    void syncEmails();
  }, []);

  useEffect(() => {
    if (!connected) return;
    const timer = window.setInterval(() => { void syncEmails(); }, 60000);
    return () => window.clearInterval(timer);
  }, [connected]);

  async function saveDivergenceReview(attachmentId: string) {
    setReviewSaving(true);
    setReviewError('');
    setReviewSuccess('');
    try {
      const response = await fetch('/api/outlook/divergence-review', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ attachmentId, ...reviewDraft }),
      });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Não foi possível salvar a revisão.');
      setReviewSuccess('Revisão salva com sucesso.');
      await loadAuditData();
    } catch (error) {
      setReviewError(error instanceof Error ? error.message : 'Falha ao salvar a revisão.');
    } finally {
      setReviewSaving(false);
    }
  }

  async function confirmDailyAudit() {
    setConfirmationLoading(true);
    setConfirmationError('');
    try {
      const response = await fetch('/api/outlook/daily-confirmation', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name: responsibleName }) });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Não foi possível confirmar.');
      setDailyConfirmation(data.confirmation);
      setResponsibleName('');
    } catch (error) {
      setConfirmationError(error instanceof Error ? error.message : 'Falha ao confirmar.');
    } finally {
      setConfirmationLoading(false);
    }
  }

  async function runAuditForDate(receivedDate: string) {
    setAuditLoading(true);
    setMailError('');
    setAuditResult(null);
    setAuditProgress({ processed: 0, percent: 0, running: true });
    let dateNextLink: string | null = null;
    const totals = { emailsFound: 0, pdfsRead: 0, auditsCreated: 0, errors: 0, ignored: 0 };
    const allResults: any[] = [];
    try {
      for (let batch = 0; batch < 500; batch++) {
        const response = await fetch('/api/outlook/audit', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ receivedDate, dateNextLink }),
          cache: 'no-store',
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Falha ao conferir os fretes deste dia.');
        totals.emailsFound += data.emailsFound || 0;
        totals.pdfsRead += data.pdfsRead || 0;
        totals.auditsCreated += data.auditsCreated || 0;
        totals.errors += data.errors || 0;
        totals.ignored += data.ignored || 0;
        allResults.push(...(data.results || []));
        dateNextLink = data.dateNextLink || null;
        const percent = data.hasMore ? Math.min(95, Math.max(5, Math.round((batch + 1) / (batch + 2) * 100))) : 100;
        setAuditProgress({ processed: totals.emailsFound, percent, running: Boolean(data.hasMore) });
        setAuditResult({ ...totals, results: [...allResults], hasMore: Boolean(data.hasMore), receivedDate });
        if (!data.hasMore) break;
        if (batch === 499) throw new Error('A conferência deste dia atingiu o limite de segurança. Execute novamente para continuar.');
      }
      setAuditResult({ ...totals, results: allResults, hasMore: false, receivedDate });
      setAuditProgress({ processed: totals.emailsFound, percent: 100, running: false });
      await loadAuditData();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Erro na conferência deste dia.';
      setMailError(message);
      if (/reconectar microsoft outlook|aadsts50173|invalid_grant/i.test(message)) setConnected(false);
      setAuditResult({ ...totals, results: allResults, hasMore: Boolean(dateNextLink), receivedDate });
      setAuditProgress(progress => ({ ...progress, running: false }));
    } finally {
      setAuditLoading(false);
    }
  }

  async function runAudit(reprocessProcessed = false) {
    setAuditLoading(true);
    setMailError('');
    setAuditResult(null);
    setAuditProgress({ processed: 0, percent: 0, running: true });
    let cursor: any = { cursorDanubio: null, cursorSiimp: null, doneDanubio: false, doneSiimp: false };
    const totals = { emailsFound: 0, pdfsRead: 0, auditsCreated: 0, errors: 0, ignored: 0 };
    const allResults: any[] = [];
    try {
      for (let batch = 0; batch < 500; batch++) {
        const response = await fetch('/api/outlook/audit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ...cursor, reprocessProcessed }), cache: 'no-store' });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Falha ao auditar e-mails.');
        totals.emailsFound += data.emailsFound || 0;
        totals.pdfsRead += data.pdfsRead || 0;
        totals.auditsCreated += data.auditsCreated || 0;
        totals.errors += data.errors || 0;
        totals.ignored += data.ignored || 0;
        allResults.push(...(data.results || []));
        cursor = { cursorDanubio: data.nextDanubio || null, cursorSiimp: data.nextSiimp || null, doneDanubio: data.doneDanubio === true, doneSiimp: data.doneSiimp === true };
        const percent = data.hasMore ? Math.min(95, Math.max(5, Math.round((totals.emailsFound / Math.max(mailCount || 1, totals.emailsFound + 1)) * 100))) : 100;
        setAuditProgress({ processed: totals.emailsFound, percent, running: Boolean(data.hasMore) });
        setAuditResult({ ...totals, results: [...allResults], hasMore: Boolean(data.hasMore) });
        setConnected(true);
        if (!data.hasMore) break;
        if (batch === 499) throw new Error('A busca atingiu o limite de segurança de lotes. Execute novamente para continuar de onde parou.');
      }
      setAuditResult({ ...totals, results: allResults, hasMore: false });
      setAuditProgress({ processed: totals.emailsFound, percent: 100, running: false });
      await Promise.all([loadAuditData(), loadDailyConfirmation()]);
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Erro na auditoria.';
      setMailError(message);
      if (/reconectar microsoft outlook|aadsts50173|invalid_grant/i.test(message)) setConnected(false);
      setAuditResult({ ...totals, results: allResults, hasMore: true });
      setAuditProgress(progress => ({ ...progress, running: false }));
    } finally {
      setAuditLoading(false);
    }
  }

  const documents = savedData?.documents || [];
  // A aba Documentos é um arquivo de ocorrências resolvidas, nunca uma listagem geral de CT-es.
  const completedDocuments = documents.filter((doc: any) =>
    ["divergent", "needs_review"].includes(doc.audit?.audit_status) &&
    doc.review?.review_status === "completed"
  );
  const divergences = savedData?.divergences || [];
  const totals = savedData?.totals || { documents: 0, divergences: 0, errors: 0, audited: 0 };
  // Registros persistidos no Supabase são associados ao e-mail original pelo ID do Microsoft Graph.
  const getEmailAudit = (mail: Mail) => {
    const docs = (savedData?.documents || []).filter((doc: any) => doc.email?.graph_message_id === mail.id);
    const audited = docs.filter((doc: any) => Boolean(doc.audit));
    const divergences = docs.filter((doc: any) =>
      ["divergent", "needs_review"].includes(doc.audit?.audit_status) || doc.processing_status === "error"
    );
    return { docs, audited, divergences };
  };
  const sectionInfo: Record<Section, { title: string; subtitle: string }> = {
    overview: { title: 'Auditoria de fretes', subtitle: 'Acompanhe a operação e a situação das cobranças da Danúbio.' },
    documents: { title: 'Documentos', subtitle: 'Consulte somente as divergências que já tiveram a revisão concluída.' },
    divergences: { title: 'Divergências', subtitle: 'Revise cobranças divergentes, falhas de leitura e casos pendentes.' },
    newEmails: { title: 'Novos e-mails', subtitle: 'Mensagens recebidas hoje, considerando a virada do dia à meia-noite no horário de Brasília.' },
    emails: { title: 'E-mails', subtitle: 'Gerencie a conexão com Outlook e acompanhe as mensagens sincronizadas.' },
    search: { title: 'Pesquisa', subtitle: 'Encontre documentos por dia de recebimento ou número do CT-e.' },
    invoices: { title: 'Conferência de faturas', subtitle: 'Compare cada CT-e cobrado pela Danúbio com o documento original recebido por e-mail.' },
  };

  function renderDocument(doc: any) {
    const audit = doc.audit || {};
    const email = doc.email || {};
    const freight = audit.extracted_fields?.grossFreight ?? audit.freight_charge;
    const rawText = String(audit.extracted_fields?.extractedTextSample || '');
    const cteNumber = textField(rawText, /CTE Nº DOCUMENTO:\s*([^\s]+)/i);
    return <article className="record-card" key={doc.id}>
      <div className="record-heading">
        <div><b>{doc.file_name || 'Anexo sem nome'}</b><small>{email.subject || 'E-mail não identificado'} · {dateTime(email.received_at)}</small></div>
        <span className={'record-status ' + (audit.audit_status || doc.processing_status || 'pending')}>{statusLabel(audit.audit_status || doc.processing_status)}</span>
      </div>
      <div className="record-grid">
        <span>Número do CT-e<strong>{cteNumber}</strong></span>
        <span>Tomador<strong>{audit.tomador_cnpj || 'Não identificado'}</strong></span>
        <span>Valor da NF<strong>{brl(audit.invoice_value)}</strong></span>
        <span>Peso<strong>{audit.weight_kg != null ? String(audit.weight_kg) + ' kg' : 'Não identificado'}</strong></span>
        <span>Frete total<strong>{brl(freight)}</strong></span>
        <span>Referência<strong>{brl(audit.reference_minimum)}</strong></span>
        <span>Diferença<strong>{brl(audit.discrepancy_amount)}</strong></span>
      </div>
      {doc.review?.review_status === 'completed' && <p className="record-reason">Revisão concluída{doc.review.reviewed_by ? ' por ' + doc.review.reviewed_by : ''}{doc.review.reviewed_at ? ' em ' + dateTime(doc.review.reviewed_at) : ''}.{doc.review.assessment ? ' Avaliação: ' + ({ correct_charge: 'Cobrança correta', incorrect_charge: 'Cobrança incorreta', unconfirmed: 'Não confirmado' } as Record<string, string>)[doc.review.assessment] : ''}{doc.review.reason ? ' Observações: ' + doc.review.reason : ''}</p>}
      {(audit.discrepancy_reason || doc.processing_status === 'error') && <p className="record-reason">{audit.discrepancy_reason || 'O PDF apresentou erro durante a leitura.'}</p>}
    </article>;
  }

  return <main className="page">
    {initialLoading && <div className="initial-loading" role="status" aria-live="polite"><div className="loading-brand"><span>AF</span><div>ALEX FREIGHT <b>AUDIT</b></div></div><div className="loading-spinner" /><h2>Preparando sua auditoria</h2><p>Estamos carregando seus e-mails, documentos e informações.</p><div className="loading-dots"><i /><i /><i /></div></div>}
    <aside className="sidebar">
      <div className="logo">AF</div>
      <div className="brand">ALEX FREIGHT <b>AUDIT</b></div>
      <p className="menu-label">MENU</p>
      <nav className="main-nav" aria-label="Menu principal">
        {navigation.map(item => <button key={item.id} className={'menu-item ' + (activeSection === item.id ? 'selected' : '')} onClick={() => setActiveSection(item.id)}>
          <span className="menu-icon">{item.icon}</span>{item.label}
          {item.id === 'divergences' && totals.divergences > 0 && <span className="menu-count">{totals.divergences}</span>}
          {item.id === 'newEmails' && todayEmailCount > 0 && <span className="menu-count">{todayEmailCount}</span>}
        </button>)}
      </nav>
      <div className="side-note">Auditoria Danúbio<br />Revisão humana antes de contestar</div>
    </aside>

    <section className="content">
      <header>
        <div><span className="overline">PAINEL DE CONTROLE</span><h1>{sectionInfo[activeSection].title}</h1><p>{sectionInfo[activeSection].subtitle}</p></div>
        <div className="header-actions"><span className="status">● Transportadora: Danúbio</span><LogoutButton /></div>
      </header>

      {activeSection === 'overview' && <>
        <div className="cards">
          <article><small>Documentos localizados</small><strong>{totals.documents}</strong><span>Arquivos registrados no sistema</span></article>
          <article><small>Auditorias realizadas</small><strong>{totals.audited}</strong><span>Documentos com dados extraídos</span></article>
          <article><small>Divergências / revisão</small><strong>{totals.divergences}</strong><span>{totals.errors} erro(s) de leitura</span></article>
        </div>
        <div className="overview-grid">
          <article className="overview-panel"><span className="overview-symbol">▤</span><h3>Documentos</h3><strong>{totals.documents}</strong><p>CT-es e anexos registrados no banco de dados.</p><button className="secondary-action" onClick={() => setActiveSection('documents')}>Ver documentos</button></article>
          <article className="overview-panel"><span className="overview-symbol warning-symbol">⚑</span><h3>Divergências</h3><strong>{totals.divergences}</strong><p>Itens que precisam de atenção ou revisão humana.</p><button className="secondary-action" onClick={() => setActiveSection('divergences')}>Revisar divergências</button></article>
          <article className="overview-panel"><span className="overview-symbol">✉</span><h3>E-mails</h3><strong>{mailCount ?? emails.length}</strong><p>Mensagens sincronizadas com a caixa corporativa.</p><button className="secondary-action" onClick={() => setActiveSection('emails')}>Abrir e-mails</button></article>
        </div>
        <div className="overview-footer">
          <div><b>Regra de auditoria</b><p>Referência mínima de R$ 100,00, R$ 0,70 por kg ou 1,5% do valor da nota, conforme o maior critério.</p></div>
          <div><b>Conexão Outlook</b><p>{connected ? 'Caixa corporativa conectada. A consulta é somente leitura.' : 'Conecte o Outlook na seção E-mails para iniciar a sincronização.'}</p></div>
        </div>
        {dataLoading && <p className="subtle-message">Atualizando indicadores...</p>}
        {dataError && <div className="mail-error">{dataError}</div>}
      </>}

      {activeSection === 'invoices' && <section className="data-section">
        <div className="section-toolbar"><div><h2>Faturas da Danúbio</h2><p>Leitura dos PDFs recebidos por e-mail e comparação com os CT-es originais. Histórico respeita a retenção de 45 dias.</p></div><div className="invoice-toolbar"><button className="refresh-button" onClick={() => void loadInvoices()} disabled={invoiceLoading}>Atualizar</button><button onClick={() => void scanInvoices()} disabled={invoiceLoading || !connected}>{invoiceLoading ? 'Conferindo PDFs...' : 'Buscar e conferir faturas no Outlook'}</button></div></div>
        {invoiceLoading && <div className="invoice-progress" role="status" aria-live="polite">
          <div className="invoice-progress-heading"><b>Conferência em andamento</b><strong>{invoiceProgress.percent}%</strong></div>
          <div className="invoice-progress-track"><div className="invoice-progress-fill" style={{ width: invoiceProgress.percent + '%' }} /></div>
          <div className="invoice-progress-detail"><span>{invoiceProgress.processed} de {invoiceProgress.total} e-mails com anexos verificados</span><span>Faltam {Math.max(0, invoiceProgress.total - invoiceProgress.processed)}</span></div>
        </div>}
        {invoiceError && <div className="mail-error">{invoiceError}</div>}
        {!connected && <div className="alert warning">Conecte o Outlook na aba E-mails para localizar as faturas.</div>}
        {invoiceLoading && <div className="empty-state">Lendo PDFs das faturas e comparando os documentos...</div>}
        {!invoiceLoading && invoiceData.length === 0 && <div className="empty-state"><span>▧</span><b>Nenhuma fatura processada</b><p>Use “Buscar e conferir faturas no Outlook”. O sistema procura PDFs de fatura da Danúbio nos e-mails ainda dentro dos 45 dias.</p></div>}
        {(() => {
          const grouped = new Map<string, {key:string; title:string; invoices:any[]}>();
          for (const invoice of invoiceData) {
            const start = invoice.period_start || null;
            const end = invoice.period_end || null;
            const key = start && end ? start + '__' + end : 'period_unknown';
            const title = start && end
              ? new Date(start+'T12:00:00').toLocaleDateString('pt-BR') + ' até ' + new Date(end+'T12:00:00').toLocaleDateString('pt-BR')
              : 'Período não identificado';
            if (!grouped.has(key)) grouped.set(key,{key,title,invoices:[]});
            grouped.get(key)!.invoices.push(invoice);
          }
          const folders = [...grouped.values()].sort((a,b) => {
            if (a.key === 'period_unknown') return 1;
            if (b.key === 'period_unknown') return -1;
            return b.key.localeCompare(a.key);
          });
          return folders.length ? <div className="invoice-folders-list">{folders.map(folder => {
            const open = invoiceFoldersOpen[folder.key] !== false;
            return <section className="invoice-folder" key={folder.key}>
              <button className="invoice-folder-toggle" onClick={()=>setInvoiceFoldersOpen(current=>({...current,[folder.key]:!open}))} aria-expanded={open}>
                <span className="invoice-folder-icon">{open?'📂':'📁'}</span>
                <span className="invoice-folder-name"><b>Faturas · {folder.title}</b><small>{folder.invoices.length} fatura(s) já varrida(s) e salva(s)</small></span>
                <span className="invoice-folder-chevron">{open?'Recolher ▲':'Abrir ▼'}</span>
              </button>
              {open && <div className="record-list invoice-folder-records">{folder.invoices.map((invoice:any) => {
          const items=invoice.items||[]; const meta=invoice.extraction_data||{}; const counts=meta.statuses||{}; const expanded=selectedInvoiceId===invoice.id;
          const reportItems=items.filter((item:any)=>item.comparison_status!=='matched');
          const matchedItems=items.filter((item:any)=>item.comparison_status==='matched');
          const sumField=(rows:any[],field:string)=>rows.reduce((sum:number,item:any)=>sum+(Number(item[field])||0),0);
          const reportCharged=sumField(items,'charged_total');
          const comparableItems=items.filter((item:any)=>item.original_freight_total!=null&&item.comparison_status!=='duplicate');
          const reportOriginal=sumField(comparableItems,'original_freight_total');
          const reportAbove=sumField(comparableItems.filter((item:any)=>Number(item.value_difference)>0),'value_difference');
          const reportBelow=Math.abs(sumField(comparableItems.filter((item:any)=>Number(item.value_difference)<0),'value_difference'));
          const reportWeight=sumField(comparableItems,'weight_difference');
          const reportValueDifference=sumField(comparableItems.filter((item:any)=>item.value_difference!=null),'value_difference');
          const reportNotFound=items.filter((item:any)=>item.comparison_status==='not_found').length;
          const reportDuplicate=items.filter((item:any)=>item.comparison_status==='duplicate').length;
          const reportWeightCount=items.filter((item:any)=>['weight_difference','both_different'].includes(item.comparison_status)).length;
          const reportValueCount=items.filter((item:any)=>['value_difference','both_different'].includes(item.comparison_status)).length;
          const label=(s:string)=>({matched:'Confere',weight_difference:'Peso divergente',value_difference:'Valor divergente',both_different:'Peso e valor divergentes',not_found:'CT-e não localizado',duplicate:'CT-e duplicado',needs_review:'Revisar'} as Record<string,string>)[s]||s;
          return <article className="record-card" key={invoice.id}>
            <button className="invoice-card-toggle" onClick={()=>setSelectedInvoiceId(expanded?null:invoice.id)}>
              <div className="record-heading"><div><b>Fatura nº {invoice.invoice_number||'Não identificada'} · {invoice.file_name}</b><small>Emitida: {invoice.issued_at||'Não identificada'} · Período da fatura: {invoice.period_start&&invoice.period_end?new Date(invoice.period_start+'T12:00:00').toLocaleDateString('pt-BR')+' até '+new Date(invoice.period_end+'T12:00:00').toLocaleDateString('pt-BR'):'Período não identificado'} · Vencimento: {invoice.due_at||'Não identificado'}</small></div><span className={'record-status '+(invoice.processing_status==='processed'?'accepted':'needs_review')}>{invoice.processing_status==='processed'?'Extraída':'Revisar extração'}</span></div>
              <div className="record-grid"><span>CT-es na fatura<strong>{invoice.document_count??items.length}</strong></span><span>Peso total<strong>{invoice.total_weight!=null?Number(invoice.total_weight).toLocaleString('pt-BR')+' kg':'Não identificado'}</strong></span><span>Total da fatura<strong>{brl(invoice.total_invoice)}</strong></span><span>Confere<strong>{counts.matched||0}</strong></span><span>Divergências<strong>{(counts.weight_difference||0)+(counts.value_difference||0)+(counts.both_different||0)}</strong></span><span>Não localizados<strong>{counts.not_found||0}</strong></span></div>
              <div className="invoice-expand-label">{expanded?'Ocultar documentos ▲':'Ver conferência dos CT-es ▼'}</div>
            </button>
            <div style={{padding:'0 14px 14px',display:'grid',gap:8}}>
              <a href={`/api/outlook/invoices/${encodeURIComponent(invoice.id)}/original`} target="_blank" rel="noreferrer" aria-label={`Abrir PDF original da fatura ${invoice.invoice_number||''}`} style={{boxSizing:'border-box',display:'flex',width:'100%',minHeight:42,alignItems:'center',justifyContent:'center',gap:9,padding:'11px 16px',borderRadius:9,background:'#1d4ed8',color:'#ffffff',fontSize:13,fontWeight:700,textDecoration:'none',border:'1px solid #1d4ed8',cursor:'pointer'}}>
                📄 Abrir fatura original recebida por e-mail
              </a>
              {invoice.invoice_number === '57197' && <button type="button" onClick={()=>void recoverInvoiceCtes(invoice)} disabled={invoiceRecoveryId!==null} style={{width:'100%',minHeight:42,borderRadius:9,border:'1px solid #b7791f',background:'#fff7e6',color:'#8a4b08',fontSize:13,fontWeight:700,cursor:invoiceRecoveryId?'wait':'pointer',opacity:invoiceRecoveryId?0.7:1}}>
                {invoiceRecoveryId===invoice.id ? '⏳ Recuperando CT-es históricos...' : '↻ Recuperar 146 CT-es não localizados'}
              </button>}
              {invoiceRecoveryStatusId===invoice.id && <small style={{color:'#a16207'}}>{invoiceRecoveryStatus}</small>}
            </div>
            {expanded && <div className="invoice-items-wrap">
              <section className="invoice-report">
                <div className="invoice-report-heading">
                  <div><span className="invoice-report-kicker">RELATÓRIO AUTOMÁTICO DE AUDITORIA</span><h3>Relatório de divergências · Fatura {invoice.invoice_number||'não identificada'}</h3><p>Gerado após a leitura do PDF e a comparação com os CT-es originais.</p></div>
                  <button className="invoice-report-print" onClick={()=>window.print()}>Imprimir / salvar PDF</button>
                </div>
                <div className="invoice-report-stats">
                  <div><small>CT-es conferidos</small><strong>{items.length}</strong></div>
                  <div><small>Sem divergência</small><strong>{matchedItems.length}</strong></div>
                  <div><small>Com divergência</small><strong>{reportItems.length}</strong></div>
                  <div><small>Não localizados</small><strong>{reportNotFound}</strong></div>
                </div>
                <div className="invoice-report-financial">
                  <div><span>Total cobrado na fatura (itens lidos)</span><b>{brl(reportCharged)}</b></div>
                  <div><span>Total original dos CT-es localizados</span><b>{brl(reportOriginal)}</b></div>
                  <div><span>Cobrança acima do CT-e original</span><b>{brl(reportAbove)}</b></div>
                  <div><span>Cobrança abaixo do CT-e original</span><b>{brl(reportBelow)}</b></div>
                  <div><span>Diferença líquida dos valores comparáveis</span><b>{brl(reportValueDifference)}</b></div>
                  <div><span>Diferença líquida de peso</span><b>{reportWeight.toLocaleString('pt-BR',{maximumFractionDigits:3})} kg</b></div>
                </div>
                <div className="invoice-report-notes">
                  <span>{reportWeightCount} CT-e(s) com divergência de peso</span>
                  <span>{reportValueCount} CT-e(s) com divergência de valor</span>
                  <span>{reportDuplicate} CT-e(s) duplicado(s)</span>
                </div>
                <h4>Ocorrências que exigem análise</h4>
                {reportItems.length ? <div className="invoice-table-scroll"><table className="invoice-table"><thead><tr><th>CT-e</th><th>NF</th><th>Situação</th><th>Peso fatura</th><th>Peso CT-e</th><th>Dif. peso</th><th>Valor fatura</th><th>Valor CT-e</th><th>Dif. valor</th></tr></thead><tbody>{reportItems.map((item:any)=><tr key={'report-'+item.id} className="invoice-discrepancy-row" tabIndex={0} role="button" aria-label={'Ver explicação da divergência do CT-e '+(item.cte_number||'não identificado')} onClick={()=>setSelectedInvoiceDiscrepancy({invoice,item})} onKeyDown={event=>{if(event.key==='Enter'||event.key===' '){event.preventDefault();setSelectedInvoiceDiscrepancy({invoice,item});}}}><td>{item.cte_number||'—'}</td><td>{item.invoice_number||'—'}</td><td>{label(item.comparison_status)}</td><td>{item.weight_kg!=null?Number(item.weight_kg).toLocaleString('pt-BR')+' kg':'—'}</td><td>{item.original_weight_kg!=null?Number(item.original_weight_kg).toLocaleString('pt-BR')+' kg':'Não localizado'}</td><td>{item.weight_difference!=null?Number(item.weight_difference).toLocaleString('pt-BR')+' kg':'—'}</td><td>{brl(item.charged_total)}</td><td>{brl(item.original_freight_total)}</td><td>{item.value_difference!=null?brl(item.value_difference):'—'}</td></tr>)}</tbody></table></div> : <div className="invoice-report-clear">Nenhuma divergência encontrada nos CT-es comparados.</div>}
                <p className="invoice-report-footnote">Critério: valores comparados pelo frete total incluindo ICMS; tolerância de R$ 0,01. Pesos são comparados com o peso extraído do CT-e original. Itens não localizados ou duplicados precisam de conferência humana. Totais originais consideram somente CT-es localizados; diferenças financeiras consideram somente itens com comparação disponível.</p>
              </section>
              <div className="invoice-totals"><b>Fechamento dos itens extraídos</b><span>Quantidade: {meta.totalsCheck?.countItems??items.length} / {invoice.document_count??'—'}</span><span>Peso somado: {meta.totalsCheck?.weightItems!=null?Number(meta.totalsCheck.weightItems).toLocaleString('pt-BR')+' kg':'—'}</span><span>Soma dos fretes: {brl(meta.totalsCheck?.sumItems)}</span></div>
              <div className="invoice-table-scroll"><table className="invoice-table"><thead><tr><th>CT-e</th><th>NF</th><th>Peso fatura</th><th>Peso CT-e</th><th>Dif. peso</th><th>Valor fatura</th><th>Valor CT-e</th><th>Situação</th></tr></thead><tbody>{items.map((item:any)=><tr key={item.id}><td>{item.cte_number||'—'}</td><td>{item.invoice_number||'—'}</td><td>{item.weight_kg!=null?Number(item.weight_kg).toLocaleString('pt-BR')+' kg':'—'}</td><td>{item.original_weight_kg!=null?Number(item.original_weight_kg).toLocaleString('pt-BR')+' kg':'Não localizado'}</td><td>{item.weight_difference!=null?Number(item.weight_difference).toLocaleString('pt-BR')+' kg':'—'}</td><td>{brl(item.charged_total)}</td><td>{brl(item.original_freight_total)}</td><td><span className={'record-status '+(item.comparison_status==='matched'?'accepted':'needs_review')}>{label(item.comparison_status)}</span></td></tr>)}</tbody></table></div>
            </div>}
          </article>;
        })}</div>}
            </section>;
          })}</div> : <div className="empty-state"><span>▧</span><b>Nenhuma fatura processada</b><p>Use “Buscar e conferir faturas no Outlook”. As faturas lidas serão salvas automaticamente em pastas pelo período que consta no PDF, respeitando a retenção de 45 dias.</p></div>;
        })()}
        {selectedInvoiceDiscrepancy && (() => {
          const { invoice, item } = selectedInvoiceDiscrepancy;
          const statusText = ({matched:'Confere',weight_difference:'Peso divergente',value_difference:'Valor divergente',both_different:'Peso e valor divergentes',not_found:'CT-e não localizado',duplicate:'CT-e duplicado',needs_review:'Revisar'} as Record<string,string>)[item.comparison_status] || item.comparison_status || 'Revisar';
          const charged = Number(item.charged_total);
          const original = item.original_freight_total == null ? null : Number(item.original_freight_total);
          const chargedWeight = item.weight_kg == null ? null : Number(item.weight_kg);
          const originalWeight = item.original_weight_kg == null ? null : Number(item.original_weight_kg);
          const valueDiff = item.value_difference == null ? null : Number(item.value_difference);
          const weightDiff = item.weight_difference == null ? null : Number(item.weight_difference);
          const explanation = item.comparison_status === 'not_found'
            ? 'O CT-e informado na fatura não foi localizado entre os documentos originais disponíveis no período. É necessário conferir o número do documento e localizar o CT-e antes de validar a cobrança.'
            : item.comparison_status === 'duplicate'
              ? 'O mesmo CT-e apareceu mais de uma vez na fatura. Confira se houve duplicidade de lançamento antes de aprovar o pagamento.'
              : item.comparison_status === 'weight_difference'
                ? 'O peso informado na fatura é diferente do peso registrado no CT-e original. A diferença precisa ser conferida com o documento de origem.'
                : item.comparison_status === 'value_difference'
                  ? 'O valor cobrado na fatura é diferente do valor total registrado no CT-e original. A comparação considera o frete total, incluindo ICMS, com tolerância de R$ 0,01.'
                  : item.comparison_status === 'both_different'
                    ? 'Foram encontradas duas diferenças: o peso da fatura não coincide com o CT-e original e o valor cobrado também está diferente. Confira os dois campos antes de contestar.'
                    : 'Este item foi sinalizado para revisão. Confira os dados extraídos da fatura e o CT-e original antes de concluir a análise.';
          return <div className="invoice-discrepancy-backdrop" role="presentation" onMouseDown={event=>{if(event.target===event.currentTarget)setSelectedInvoiceDiscrepancy(null);}}>
            <section className="invoice-discrepancy-modal" role="dialog" aria-modal="true" aria-labelledby="invoice-discrepancy-title">
              <div className="invoice-discrepancy-heading">
                <div><span className="invoice-report-kicker">ANÁLISE DA COBRANÇA</span><h2 id="invoice-discrepancy-title">Divergência · CT-e {item.cte_number||'não identificado'}</h2><p>Fatura nº {invoice.invoice_number||'não identificada'} · {invoice.file_name||'Arquivo não identificado'}</p></div>
                <button className="invoice-discrepancy-close no-print" onClick={()=>setSelectedInvoiceDiscrepancy(null)} aria-label="Fechar">×</button>
              </div>
              <div className="invoice-discrepancy-status"><b>{statusText}</b></div>
              <p className="invoice-discrepancy-explanation">{explanation}</p>
              <div className="invoice-discrepancy-values">
                <div><small>Valor cobrado na fatura</small><strong>{brl(item.charged_total)}</strong></div>
                <div><small>Valor no CT-e original</small><strong>{brl(item.original_freight_total)}</strong></div>
                <div><small>Peso na fatura</small><strong>{chargedWeight==null?'Não informado':chargedWeight.toLocaleString('pt-BR')+' kg'}</strong></div>
                <div><small>Peso no CT-e original</small><strong>{originalWeight==null?'Não localizado':originalWeight.toLocaleString('pt-BR')+' kg'}</strong></div>
                <div><small>Diferença de valor</small><strong>{valueDiff==null?'Não calculada':brl(valueDiff)}</strong></div>
                <div><small>Diferença de peso</small><strong>{weightDiff==null?'Não calculada':weightDiff.toLocaleString('pt-BR',{maximumFractionDigits:3})+' kg'}</strong></div>
              </div>
              <div className="invoice-discrepancy-guidance">
                <b>O que precisa ser conferido</b>
                <p>{original==null?'Localize o CT-e original e confirme se o documento pertence ao período desta fatura.':'Compare os valores e pesos acima com o CT-e original. Se a diferença for confirmada, registre a ocorrência com a Danúbio antes de aprovar o pagamento.'}</p>
                <small>Período da fatura: {invoice.period_start&&invoice.period_end?new Date(invoice.period_start+'T12:00:00').toLocaleDateString('pt-BR')+' até '+new Date(invoice.period_end+'T12:00:00').toLocaleDateString('pt-BR'):'Não identificado'}.</small>
              </div>
              <div className="invoice-discrepancy-actions no-print">
                <button className="refresh-button" onClick={()=>setSelectedInvoiceDiscrepancy(null)}>Fechar</button>
                <button onClick={()=>{document.body.classList.add('printing-invoice-discrepancy');const cleanup=()=>{document.body.classList.remove('printing-invoice-discrepancy');window.removeEventListener('afterprint',cleanup);};window.addEventListener('afterprint',cleanup);window.print();}}>Imprimir divergência / salvar PDF</button>
              </div>
              <p className="invoice-discrepancy-foot">Relatório gerado automaticamente a partir dos dados extraídos. A ocorrência deve passar por conferência humana antes de qualquer contestação.</p>
            </section>
          </div>;
        })()}
      </section>}

      {activeSection === 'search' && (() => {
        const normalizedCte = searchCte.trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        const hasSearchCriteria = Boolean(searchDate || searchCte.trim());
        const results = hasSearchCriteria ? documents.filter((doc: any) => {
          const receivedAt = doc.email?.received_at;
          const day = receivedAt ? new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(receivedAt)) : '';
          const rawText = String(doc.audit?.extracted_fields?.extractedTextSample || '');
          const cte = textField(rawText, /CTE Nº DOCUMENTO:\s*([^\s]+)/i);
          const normalizedDocCte = cte.toLowerCase().replace(/[^a-z0-9]/g, '');
          const matchesDate = !searchDate || day === searchDate;
          const matchesCte = !normalizedCte || normalizedDocCte.includes(normalizedCte) || String(doc.file_name || '').toLowerCase().includes(searchCte.trim().toLowerCase());
          return matchesDate && matchesCte;
        }) : [];
        return <section className="data-section">
          <div className="section-toolbar"><div><h2>Localizar CT-e</h2><p>Pesquise nos documentos salvos nos últimos 45 dias.</p></div><button className="refresh-button" onClick={() => void loadAuditData()} disabled={dataLoading}>{dataLoading ? 'Atualizando...' : 'Atualizar dados'}</button></div>
          <div className="audit-search-fields">
            <label>Pesquisar por dia de recebimento<input type="date" value={searchDate} onChange={e => setSearchDate(e.target.value)} /></label>
            <label>Pesquisar por número do CT-e<input value={searchCte} onChange={e => setSearchCte(e.target.value)} placeholder="Digite o número do CT-e" /></label>
            <button className="search-clear" onClick={() => { setSearchDate(''); setSearchCte(''); }}>Limpar pesquisa</button>
          </div>
          <div className="search-results-heading"><b>{hasSearchCriteria ? results.length + ' documento(s) encontrado(s)' : 'Aguardando pesquisa'}</b><span>Os CT-es só aparecem depois que você informar uma data ou número para pesquisar.</span></div>
          {dataError && <div className="mail-error">{dataError}</div>}
          {dataLoading && !savedData ? <div className="empty-state">Carregando documentos...</div> : !hasSearchCriteria ? <div className="empty-state"><span>⌕</span><b>Pesquise um CT-e</b><p>Digite o número do CT-e ou selecione o dia de recebimento. Nenhum documento será listado antes da pesquisa.</p></div> : results.length ? <div className="record-list">{results.map(renderDocument)}</div> : <div className="empty-state"><span>⌕</span><b>Nenhum documento encontrado</b><p>Altere a data ou confira o número do CT-e. A pesquisa considera os registros ainda dentro do prazo de retenção.</p></div>}
        </section>;
      })()}

      {activeSection === 'documents' && <section className="data-section">
        <div className="section-toolbar"><div><h2>Divergências concluídas</h2><p>{completedDocuments.length} CT-e(s) com divergência e revisão concluída · dados salvos no Supabase</p></div><button className="refresh-button" onClick={() => void loadAuditData()} disabled={dataLoading}>{dataLoading ? 'Atualizando...' : 'Atualizar lista'}</button></div>
        {dataError && <div className="mail-error">{dataError}</div>}
        {dataLoading && !savedData ? <div className="empty-state">Carregando documentos...</div> : completedDocuments.length ? <div className="record-list">{completedDocuments.map(renderDocument)}</div> : <div className="empty-state"><span>▤</span><b>Nenhuma divergência concluída</b><p>Somente CT-es identificados com divergência e cuja revisão foi concluída aparecerão aqui. Os demais CT-es continuam organizados na aba E-mails por dia.</p><button onClick={() => setActiveSection('divergences')}>Ir para Divergências</button></div>}
      </section>}

      {activeSection === 'divergences' && <section className="data-section">
        <div className="section-toolbar"><div><h2>Itens que exigem atenção</h2><p>{divergences.length} ocorrência(s) · clique em uma ocorrência para registrar a revisão</p></div><button className="refresh-button" onClick={() => void loadAuditData()} disabled={dataLoading}>{dataLoading ? 'Atualizando...' : 'Atualizar lista'}</button></div>
        {dataError && <div className="mail-error">{dataError}</div>}
        {divergences.length ? <div className="record-list">{divergences.map((doc: any) => {
          const audit = doc.audit || {};
          const email = doc.email || {};
          const review = doc.review || {};
          const isSelected = selectedReviewId === doc.id;
          const freight = audit.extracted_fields?.grossFreight ?? audit.freight_charge;
          const reviewStatusLabel = ({ pending: 'Pendente', in_progress: 'Em andamento', completed: 'Concluída' } as Record<string, string>)[review.review_status || 'pending'];
          const assessmentLabel = ({ correct_charge: 'Cobrança correta', incorrect_charge: 'Cobrança incorreta', unconfirmed: 'Não confirmado' } as Record<string, string>)[review.assessment || ''] || 'Não avaliada';
          return <article className={'record-card divergence-card ' + (isSelected ? 'expanded' : '')} key={doc.id}>
            <button className="divergence-open" onClick={() => {
              setSelectedReviewId(isSelected ? null : doc.id);
              setReviewDraft({
                reviewStatus: review.review_status || 'pending',
                assessment: review.assessment || '',
                reason: review.reason || '',
                reviewedBy: review.reviewed_by || '',
              });
              setReviewError('');
              setReviewSuccess('');
            }}>
              <div className="record-heading"><div><b>{doc.file_name || 'Anexo sem nome'}</b><small>{email.subject || 'E-mail não identificado'} · {dateTime(email.received_at)}</small></div><span className={'record-status ' + (audit.audit_status || doc.processing_status || 'pending')}>{statusLabel(audit.audit_status || doc.processing_status)}</span></div>
              <div className="record-grid">
                <span>Frete total<strong>{brl(freight)}</strong></span>
                <span>Referência<strong>{brl(audit.reference_minimum)}</strong></span>
                <span>Diferença<strong>{brl(audit.discrepancy_amount)}</strong></span>
              </div>
              <div className="divergence-footer"><span className={'review-chip review-' + (review.review_status || 'pending')}>Revisão: {reviewStatusLabel}</span><span>{assessmentLabel}</span><b>{isSelected ? 'Fechar ▲' : 'Clique para revisar ▼'}</b></div>
            </button>
            {(audit.discrepancy_reason || doc.processing_status === 'error') && <p className="record-reason">{audit.discrepancy_reason || 'O PDF apresentou erro durante a leitura.'}</p>}
            {isSelected && <div className="occurrence-details">
              <h3>Detalhes da cobrança · CT-e</h3>
              {(() => {
                const raw = String(audit.extracted_fields?.extractedTextSample || '');
                const icms = audit.extracted_fields?.icmsValue;
                const baseFreight = audit.freight_charge;
                const gross = audit.extracted_fields?.grossFreight ?? freight;
                const grossNumber = Number(gross);
                const invoiceNumber = Number(audit.invoice_value);
                const weightNumber = Number(audit.weight_kg);
                const impliedPercent = Number.isFinite(grossNumber) && grossNumber >= 0 && Number.isFinite(invoiceNumber) && invoiceNumber > 0
                  ? (grossNumber / invoiceNumber) * 100 : null;
                const impliedPerKg = Number.isFinite(grossNumber) && grossNumber >= 0 && Number.isFinite(weightNumber) && weightNumber > 0
                  ? grossNumber / weightNumber : null;
                const cteNumber = textField(raw, /CTE Nº DOCUMENTO:\s*([^\s]+)/i);
                const product = textField(raw, /PRODUTO PREDOMINANTE\s+(.+?)\s+OUTRAS CARACTERÍSTICAS/i);
                const volumes = textField(raw, /QUANTIDADE DE VOLUMES \(UND\)\s*([\d.,]+)/i);
                const cubage = textField(raw, /CUBAGEM \(M3\)\s*([\d.,]+)/i);
                const origin = textField(raw, /INICIO DA PRESTAÇÃO\s+(.+?)\s+TERMINO DA PRESTAÇÃO/i);
                const destination = textField(raw, /TERMINO DA PRESTAÇÃO\s+(.+?)\s+REMETENTE:/i);
                const receiver = textField(raw, /VALOR A RECEBER\s*([\d.,]+)/i);
                const icmsRate = textField(raw, /ALÍQUOTA DO ICMS\s*([\d.,]+)\s*/i);
                return <>
                  <div className="record-grid">
                    <span>Número do CT-e<strong>{cteNumber}</strong></span>
                    <span>Valor da nota fiscal<strong>{brl(audit.invoice_value)}</strong></span>
                    <span>Peso bruto<strong>{audit.weight_kg != null ? String(audit.weight_kg) + ' kg' : 'Não identificado'}</strong></span>
                    <span>Quantidade de volumes<strong>{volumes === 'Não identificado no PDF' ? volumes : volumes + ' un.'}</strong></span>
                    <span>Cubagem<strong>{cubage === 'Não identificado no PDF' ? cubage : cubage + ' m³'}</strong></span>
                    <span>Produto predominante<strong>{product}</strong></span>
                    <span>Origem<strong>{origin}</strong></span>
                    <span>Destino<strong>{destination}</strong></span>
                  </div>
                  <h4>Composição dos valores cobrados</h4>
                  <div className="record-grid">
                    <span>Frete informado (sem ICMS)<strong>{brl(baseFreight)}</strong></span>
                    <span>ICMS<strong>{brl(icms)}</strong></span>
                    <span>Alíquota do ICMS<strong>{icmsRate === 'Não identificado no PDF' ? icmsRate : icmsRate + '%'}</strong></span>
                    <span>Frete total com ICMS<strong>{brl(gross)}</strong></span>
                    <span>Valor a receber<strong>{receiver === 'Não identificado no PDF' ? brl(gross) : brl(receiver.replace(/\\./g, '').replace(',', '.'))}</strong></span>
                    <span>Valor de referência<strong>{brl(audit.reference_minimum)}</strong></span>
                    <span>Diferença apurada<strong>{brl(audit.discrepancy_amount)}</strong></span>
                  </div>
                  <h4>Equivalência da cobrança</h4>
                  <div className="record-grid">
                    <span>Se fosse cobrado por percentual da NF<strong>{impliedPercent !== null ? impliedPercent.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%' : 'Não calculável — valor da NF ausente'}</strong></span>
                    <span>Se fosse cobrado por peso<strong>{impliedPerKg !== null ? brl(impliedPerKg) + ' por kg' : 'Não calculável — peso ausente'}</strong></span>
                  </div>
                  <p className="detail-note">Equivalência calculada usando o frete total com ICMS: percentual = frete total ÷ valor da NF × 100; valor por kg = frete total ÷ peso bruto. São taxas equivalentes da cobrança, não tarifas contratuais confirmadas.</p>
                  <p className="detail-note">Os valores são apresentados conforme os campos extraídos do CT-e. O frete total considera o ICMS, conforme a regra de auditoria.</p>
                </>;
              })()}
            </div>}
            {isSelected && <div className="review-editor">
              <h3>Revisão da ocorrência</h3>
              <label>Situação da ocorrência</label>
              <select value={reviewDraft.reviewStatus} onChange={e => setReviewDraft(v => ({ ...v, reviewStatus: e.target.value }))}>
                <option value="pending">Pendente</option><option value="in_progress">Em andamento</option><option value="completed">Concluída</option>
              </select>
              <label>A cobrança está correta?</label>
              <select value={reviewDraft.assessment} onChange={e => setReviewDraft(v => ({ ...v, assessment: e.target.value }))}>
                <option value="">Selecione uma avaliação</option><option value="correct_charge">Sim, está correta</option><option value="incorrect_charge">Não, está incorreta</option><option value="unconfirmed">Ainda não foi possível confirmar</option>
              </select>
              <label>Motivo da divergência / observações</label>
              <textarea rows={4} maxLength={3000} value={reviewDraft.reason} onChange={e => setReviewDraft(v => ({ ...v, reason: e.target.value }))} placeholder="Descreva por que a cobrança diverge, o que foi verificado e qual providência foi tomada..." />
              <label>Responsável pela revisão</label>
              <input value={reviewDraft.reviewedBy} maxLength={100} onChange={e => setReviewDraft(v => ({ ...v, reviewedBy: e.target.value }))} placeholder="Nome de quem analisou" />
              <div className="review-actions"><button onClick={() => void saveDivergenceReview(doc.id)} disabled={reviewSaving}>{reviewSaving ? 'Salvando...' : 'Salvar revisão'}</button></div>
              {reviewError && <div className="mail-error">{reviewError}</div>}
              {reviewSuccess && <div className="mail-success">{reviewSuccess}</div>}
              {review.reviewed_at && <small className="review-last">Última revisão salva em {dateTime(review.reviewed_at)}{review.reviewed_by ? ' por ' + review.reviewed_by : ''}.</small>}
            </div>}
          </article>;
        })}</div> : <div className="empty-state success-empty"><span>✓</span><b>Nenhuma divergência registrada</b><p>Não há divergências, erros de leitura ou revisões pendentes nos documentos salvos.</p></div>}
      </section>}

      {activeSection === 'newEmails' && <section className="data-section">
        <div className="section-toolbar"><div><h2>Recebidos hoje</h2><p>{brazilDate ? new Date(brazilDate + 'T12:00:00-03:00').toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' }) : 'Data de Brasília'} · {todayEmailCount} novo(s) e-mail(s)</p></div><button className="refresh-button" onClick={() => void syncEmails()} disabled={mailLoading}>{mailLoading ? 'Atualizando...' : 'Atualizar lista'}</button></div>
        <div className="mail-success">A contagem considera o horário de Brasília (UTC−3). À meia-noite, a lista começa um novo dia automaticamente. O histórico de e-mails permanece salvo.</div>
        {!connected ? <div className="empty-state">Conecte o Outlook para consultar os novos e-mails.</div> : todayEmails.length ? <div className="mail-list">{todayEmails.map(mail => <article key={mail.id} className="mail-item"><div><b>{mail.subject}</b><small>{mail.sender} · {dateTime(mail.receivedAt)}</small><p>{mail.preview}</p></div><div className="mail-meta">{mail.hasAttachments && <span className="pill">Com anexo</span>}</div></article>)}</div> : <div className="empty-state"><span>✉</span><b>Nenhum novo e-mail hoje</b><p>Quando uma mensagem chegar, ela aparecerá aqui. A contagem será reiniciada na próxima meia-noite de Brasília.</p></div>}
      </section>}

      {activeSection === 'emails' && <section className="mail-panel mail-live">
        <div className="mail-icon">✉</div>
        <div className="mail-main">
          <h3>Conexão com Microsoft Outlook</h3>
          <p>O sistema localiza mensagens do Siimp Sistemas e e-mails que mencionem “Danubio”. A consulta é somente leitura: as mensagens não são marcadas como lidas nem alteradas.</p>
          <span className="pill">Siimp Sistemas</span> <span className="pill">Palavra-chave: Danubio</span> <span className="pill">Somente leitura</span>
          {mailError && <div className="mail-error">{mailError}</div>}
          {connected && <div className="mail-success">Conta conectada. {mailCount === null ? '' : String(mailCount) + ' mensagem(ns) localizada(s).'}</div>}
          <div className="daily-review">
            <div className="daily-review-head"><div><h3>Conferência humana do dia</h3><p>Informe o nome do responsável somente depois de revisar as auditorias e pendências.</p></div><span className={dailyConfirmation?.status === 'confirmed' ? 'daily-state done' : 'daily-state'}>{dailyConfirmation?.status === 'confirmed' ? 'Confirmado' : 'Pendente'}</span></div>
            {dailyConfirmation?.status === 'confirmed'
              ? <div className="mail-success">Conferido por <b>{dailyConfirmation.confirmed_by}</b> em {dateTime(dailyConfirmation.confirmed_at)}.</div>
              : <><label className="daily-label">Nome do responsável</label><input className="daily-input" value={responsibleName} onChange={event => setResponsibleName(event.target.value)} placeholder="Digite o nome completo" maxLength={100} /><button className="daily-confirm" onClick={() => void confirmDailyAudit()} disabled={!connected || confirmationLoading || responsibleName.trim().length < 3}>{confirmationLoading ? 'Registrando confirmação...' : 'Confirmar auditoria do dia'}</button>{confirmationError && <div className="mail-error">{confirmationError}</div>}</>}
            <small className="daily-foot">A confirmação registra o responsável e o horário. Novos e-mails exigem nova conferência.</small>
          </div>
          <div className="mail-actions">
            <button onClick={() => { window.location.href = '/api/outlook/connect'; }}>Reconectar Microsoft Outlook</button>
            <button className="light" onClick={() => void syncEmails()} disabled={mailLoading}>{mailLoading ? 'Consultando mensagens...' : connected ? 'Atualizar e-mails' : 'Testar conexão'}</button>
            <button onClick={() => void runAudit()} disabled={!connected || auditLoading}>{auditLoading ? 'Auditando e-mails...' : auditResult?.hasMore ? 'Continuar varredura' : 'Auditar e-mails e CT-e'}</button>
            <button className="light" onClick={() => { if (window.confirm('Isso vai baixar e ler novamente os PDFs já processados. Deseja continuar?')) void runAudit(true); }} disabled={!connected || auditLoading}>Reprocessar PDFs já lidos</button>
          </div>
          {(auditLoading || auditResult) && <div className="mail-list progress-panel"><h4>{auditResult?.receivedDate ? 'Conferência de ' + new Date(auditResult.receivedDate + 'T12:00:00-03:00').toLocaleDateString('pt-BR') : auditLoading ? 'Auditoria em andamento' : 'Progresso da auditoria'}</h4><div className="progress-track"><div className="progress-value" style={{ width: auditProgress.percent + '%' }} /></div><small>{auditProgress.percent}% · {auditProgress.processed} e-mail(s) processado(s){auditLoading ? ' · lendo próximos lotes automaticamente...' : ' · processamento pausado/concluído'}</small></div>}
          {auditResult && <div className="mail-success">{auditResult.receivedDate ? 'Conferência do dia ' + new Date(auditResult.receivedDate + 'T12:00:00-03:00').toLocaleDateString('pt-BR') + ': ' : 'Varredura: '}{auditResult.emailsFound} e-mails, {auditResult.pdfsRead} PDFs lidos, {auditResult.auditsCreated} auditorias salvas, {auditResult.errors} erros. {auditResult.hasMore ? 'Ainda há mensagens pendentes; execute novamente para retomar.' : 'Busca concluída.'}</div>}
          <div className="mail-list">
            <div className="email-folders-heading"><div><h4>E-mails organizados por dia</h4><p>{emails.length} mensagem(ns) carregada(s) · histórico disponível por até 45 dias</p></div><span className="email-folder-count">{Object.keys(emails.reduce((groups: Record<string, Mail[]>, mail) => { const key = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(mail.receivedAt)); (groups[key] ||= []).push(mail); return groups; }, {})).length} dia(s)</span></div>
            {emails.length ? Object.entries(emails.reduce((groups: Record<string, Mail[]>, mail) => { const key = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Sao_Paulo', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(mail.receivedAt)); (groups[key] ||= []).push(mail); return groups; }, {})).sort(([a], [b]) => b.localeCompare(a)).map(([day, dayEmails]) => {
              const expanded = emailFolderState[day] ?? day === brazilDate;
              const dayLabel = new Date(day + 'T12:00:00-03:00').toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo', weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' });
              const dayAudit = dayEmails.map(mail => getEmailAudit(mail));
              const dayDocs = dayAudit.reduce((sum, item) => sum + item.docs.length, 0);
              const dayAudited = dayAudit.reduce((sum, item) => sum + item.audited.length, 0);
              const dayDivergences = dayAudit.reduce((sum, item) => sum + item.divergences.length, 0);
              const dayComplete = dayDocs > 0 && dayAudit.every((item, index) => !dayEmails[index].hasAttachments || item.audited.length > 0);
              return <section className="email-day-folder" key={day}>
                <button className="email-folder-toggle" aria-expanded={expanded} onClick={() => setEmailFolderState(current => ({ ...current, [day]: !expanded }))}>
                  <span className="email-folder-icon">{expanded ? '▾' : '▸'}</span><span className="email-folder-title"><b>{dayLabel}</b><small>{dayEmails.length} e-mail(s) · {dayAudited}/{dayDocs} documento(s) auditado(s) · {dayDivergences} divergência(s)</small></span><span className="email-folder-state">{dayComplete ? 'Conferido' : dayAudited ? 'Parcial' : 'Pendente'} · {expanded ? 'Recolher' : 'Abrir pasta'}</span>
                </button>
                <button className="secondary-action" onClick={() => void runAuditForDate(day)} disabled={!connected || auditLoading}>
                  {auditLoading && auditResult?.receivedDate === day ? 'Conferindo este dia...' : dayComplete ? 'Conferir novamente este dia' : 'Conferir fretes deste dia'}
                </button>
                {expanded && <div className="email-folder-content">{dayEmails.map(mail => {
                  const auditInfo = getEmailAudit(mail);
                  const hasAudits = auditInfo.audited.length > 0;
                  const hasDivergences = auditInfo.divergences.length > 0;
                  const stateLabel = !mail.hasAttachments ? 'Sem anexo' : hasDivergences ? 'Com divergência' : hasAudits ? 'Conferido · sem divergência' : 'Ainda não conferido';
                  return <article key={mail.id} className="mail-item"><div><b>{mail.subject}</b><small>{mail.sender} · {dateTime(mail.receivedAt)}</small><p>{mail.preview}</p>{auditInfo.docs.length > 0 && <small>{auditInfo.docs.map((doc:any) => doc.file_name + ': ' + statusLabel(doc.audit?.audit_status || doc.processing_status)).join(' · ')}</small>}</div><div className="mail-meta">{mail.hasAttachments && <span className="pill">Com anexo</span>}<span className={'record-status ' + (hasDivergences ? 'divergent' : hasAudits ? 'accepted' : 'pending')}>{stateLabel}</span>{mail.webLink && <a href={mail.webLink} target="_blank" rel="noreferrer">Abrir e-mail</a>}</div></article>;
                })}</div>}
              </section>;
            }) : <p className="subtle-message">Nenhuma mensagem sincronizada até o momento.</p>}
          </div>
        </div>
        <span className={connected ? 'not-connected connected' : 'not-connected'}>{connected ? 'Conectado' : 'Aguardando conexão'}</span>
      </section>}

      <footer>ALEX FREIGHT AUDIT · Relatórios de divergência sujeitos à revisão humana.</footer>
    </section>
  </main>;
}
