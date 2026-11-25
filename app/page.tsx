'use client';

import { useCallback, useEffect, useState } from 'react';

const brl = (n: number | string | null | undefined) => {
  const value = typeof n === 'string' ? Number(n) : n;
  return typeof value === 'number' && Number.isFinite(value) ? value.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' }) : 'Não informado';
};
const dateTime = (value?: string | null) => value ? new Date(value).toLocaleString('pt-BR') : 'Data não informada';
const statusLabel = (status?: string) => (({ accepted: 'Aceito', divergent: 'Divergência', needs_review: 'Revisão necessária', out_of_scope: 'Fora do escopo', pending: 'Pendente', extracted: 'Extraído', error: 'Erro na leitura', discovered: 'Localizado', attachments_checked: 'Anexos conferidos' } as Record<string, string>)[status || ''] || status || 'Não informado');

type Section = 'overview' | 'documents' | 'divergences' | 'newEmails' | 'emails';
type Mail = { id: string; subject: string; sender: string; senderEmail: string; receivedAt: string; hasAttachments: boolean; preview: string; webLink: string };
type AuditData = { mailbox?: string; emails: any[]; documents: any[]; divergences: any[]; totals: { documents: number; divergences: number; errors: number; audited: number } };
const navigation: { id: Section; label: string; icon: string }[] = [
  { id: 'overview', label: 'Visão geral', icon: '▦' },
  { id: 'documents', label: 'Documentos', icon: '▤' },
  { id: 'divergences', label: 'Divergências', icon: '⚑' },
  { id: 'newEmails', label: 'Novos e-mails', icon: '✉' },
  { id: 'emails', label: 'E-mails', icon: '▣' },
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
  const [auditLoading, setAuditLoading] = useState(false);
  const [auditResult, setAuditResult] = useState<any>(null);
  const [auditProgress, setAuditProgress] = useState({ processed: 0, percent: 0, running: false });
  const [responsibleName, setResponsibleName] = useState('');
  const [dailyConfirmation, setDailyConfirmation] = useState<any>(null);
  const [confirmationError, setConfirmationError] = useState('');
  const [confirmationLoading, setConfirmationLoading] = useState(false);
  const [savedData, setSavedData] = useState<AuditData | null>(null);
  const [dataLoading, setDataLoading] = useState(false);
  const [dataError, setDataError] = useState('');
  const [selectedReviewId, setSelectedReviewId] = useState<string | null>(null);
  const [reviewDraft, setReviewDraft] = useState({ reviewStatus: 'pending', assessment: '', reason: '', reviewedBy: '' });
  const [reviewSaving, setReviewSaving] = useState(false);
  const [reviewError, setReviewError] = useState('');
  const [reviewSuccess, setReviewSuccess] = useState('');

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
      setMailError(error instanceof Error ? error.message : 'Erro na auditoria.');
      setAuditResult({ ...totals, results: allResults, hasMore: true });
      setAuditProgress(progress => ({ ...progress, running: false }));
    } finally {
      setAuditLoading(false);
    }
  }

  const documents = savedData?.documents || [];
  const divergences = savedData?.divergences || [];
  const totals = savedData?.totals || { documents: 0, divergences: 0, errors: 0, audited: 0 };
  const sectionInfo: Record<Section, { title: string; subtitle: string }> = {
    overview: { title: 'Auditoria de fretes', subtitle: 'Acompanhe a operação e a situação das cobranças da Danúbio.' },
    documents: { title: 'Documentos', subtitle: 'Consulte os CT-es e arquivos localizados nas mensagens sincronizadas.' },
    divergences: { title: 'Divergências', subtitle: 'Revise cobranças divergentes, falhas de leitura e casos pendentes.' },
    newEmails: { title: 'Novos e-mails', subtitle: 'Mensagens recebidas hoje, considerando a virada do dia à meia-noite no horário de Brasília.' },
    emails: { title: 'E-mails', subtitle: 'Gerencie a conexão com Outlook e acompanhe as mensagens sincronizadas.' },
  };

  function renderDocument(doc: any) {
    const audit = doc.audit || {};
    const email = doc.email || {};
    const freight = audit.extracted_fields?.grossFreight ?? audit.freight_charge;
    return <article className="record-card" key={doc.id}>
      <div className="record-heading">
        <div><b>{doc.file_name || 'Anexo sem nome'}</b><small>{email.subject || 'E-mail não identificado'} · {dateTime(email.received_at)}</small></div>
        <span className={'record-status ' + (audit.audit_status || doc.processing_status || 'pending')}>{statusLabel(audit.audit_status || doc.processing_status)}</span>
      </div>
      <div className="record-grid">
        <span>Tomador<strong>{audit.tomador_cnpj || 'Não identificado'}</strong></span>
        <span>Valor da NF<strong>{brl(audit.invoice_value)}</strong></span>
        <span>Peso<strong>{audit.weight_kg != null ? String(audit.weight_kg) + ' kg' : 'Não identificado'}</strong></span>
        <span>Frete total<strong>{brl(freight)}</strong></span>
        <span>Referência<strong>{brl(audit.reference_minimum)}</strong></span>
        <span>Diferença<strong>{brl(audit.discrepancy_amount)}</strong></span>
      </div>
      {(audit.discrepancy_reason || doc.processing_status === 'error') && <p className="record-reason">{audit.discrepancy_reason || 'O PDF apresentou erro durante a leitura.'}</p>}
    </article>;
  }

  return <main className="page">
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
        <span className="status">● Transportadora: Danúbio</span>
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

      {activeSection === 'documents' && <section className="data-section">
        <div className="section-toolbar"><div><h2>Todos os documentos</h2><p>{totals.documents} arquivo(s) registrado(s) · dados salvos no Supabase</p></div><button className="refresh-button" onClick={() => void loadAuditData()} disabled={dataLoading}>{dataLoading ? 'Atualizando...' : 'Atualizar lista'}</button></div>
        {dataError && <div className="mail-error">{dataError}</div>}
        {dataLoading && !savedData ? <div className="empty-state">Carregando documentos...</div> : documents.length ? <div className="record-list">{documents.map(renderDocument)}</div> : <div className="empty-state"><span>▤</span><b>Nenhum documento registrado</b><p>Quando a auditoria encontrar anexos, eles aparecerão aqui, mesmo depois de sair e voltar ao painel.</p><button onClick={() => setActiveSection('emails')}>Ir para E-mails</button></div>}
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
            {!connected && <button onClick={() => { window.location.href = '/api/outlook/connect'; }}>Conectar Microsoft Outlook</button>}
            <button className="light" onClick={() => void syncEmails()} disabled={mailLoading}>{mailLoading ? 'Consultando mensagens...' : connected ? 'Atualizar e-mails' : 'Testar conexão'}</button>
            <button onClick={() => void runAudit()} disabled={!connected || auditLoading}>{auditLoading ? 'Auditando e-mails...' : auditResult?.hasMore ? 'Continuar varredura' : 'Auditar e-mails e CT-e'}</button>
            <button className="light" onClick={() => { if (window.confirm('Isso vai baixar e ler novamente os PDFs já processados. Deseja continuar?')) void runAudit(true); }} disabled={!connected || auditLoading}>Reprocessar PDFs já lidos</button>
          </div>
          {(auditLoading || auditResult) && <div className="mail-list progress-panel"><h4>{auditLoading ? 'Auditoria em andamento' : 'Progresso da auditoria'}</h4><div className="progress-track"><div className="progress-value" style={{ width: auditProgress.percent + '%' }} /></div><small>{auditProgress.percent}% · {auditProgress.processed} e-mail(s) processado(s){auditLoading ? ' · lendo próximos lotes automaticamente...' : ' · processamento pausado/concluído'}</small></div>}
          {auditResult && <div className="mail-success">Varredura: {auditResult.emailsFound} e-mails, {auditResult.pdfsRead} PDFs lidos, {auditResult.auditsCreated} auditorias salvas, {auditResult.errors} erros. {auditResult.hasMore ? 'Ainda há mensagens pendentes; execute novamente para retomar.' : 'Busca concluída.'}</div>}
          <div className="mail-list"><h4>E-mails sincronizados · atualização automática a cada minuto</h4>
            {emails.length ? emails.map(mail => <article key={mail.id} className="mail-item"><div><b>{mail.subject}</b><small>{mail.sender} · {dateTime(mail.receivedAt)}</small><p>{mail.preview}</p></div><div className="mail-meta">{mail.hasAttachments && <span className="pill">Com anexo</span>}{mail.webLink && <a href={mail.webLink} target="_blank" rel="noreferrer">Abrir e-mail</a>}</div></article>) : <p className="subtle-message">Nenhuma mensagem sincronizada até o momento.</p>}
          </div>
        </div>
        <span className={connected ? 'not-connected connected' : 'not-connected'}>{connected ? 'Conectado' : 'Aguardando conexão'}</span>
      </section>}

      <footer>ALEX FREIGHT AUDIT · Relatórios de divergência sujeitos à revisão humana.</footer>
    </section>
  </main>;
}
