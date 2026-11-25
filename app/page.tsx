'use client';

import { useEffect, useState } from 'react';

const brl = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
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
    } catch (error) {
      setMailError(error instanceof Error ? error.message : 'Erro ao consultar Outlook.');
      if ((error instanceof Error ? error.message : '').includes('não conectado')) setConnected(false);
    } finally {
      setMailLoading(false);
    }
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
      <div className="section-title"><div><h2>Conferir um CT-e</h2><p>Simule a conferência usando os dados do documento.</p></div><span className="tag">REGRA DANÚBIO</span></div>
      <div className="columns"><section className="panel"><h3>Dados do documento</h3><label>CNPJ do tomador<input value={cnpj} onChange={e=>setCnpj(e.target.value)}/></label><div className="two"><label>Peso (kg)<input value={peso} onChange={e=>setPeso(e.target.value)}/></label><label>Valor da nota (R$)<input value={nota} onChange={e=>setNota(e.target.value)}/></label></div><label>Frete cobrado (R$)<input value={cobrado} onChange={e=>setCobrado(e.target.value)}/></label><button onClick={()=>setResultado(true)}>Conferir cobrança</button><small className="hint">Somente o tomador Brasil Engrenagens, CNPJ 39.860.057/0001-04, pode gerar alerta. ICMS não é considerado.</small></section>
      <section className="panel"><h3>Resultado da auditoria</h3>{!resultado ? <div className="empty"><div className="empty-icon">✓</div><b>Aguardando conferência</b><p>Preencha os dados ao lado para visualizar o cálculo.</p></div> : <><div className={divergente ? 'alert warning' : 'alert good'}>{!tomadorValido ? 'Fora do escopo: tomador diferente da Brasil Engrenagens.' : abaixoDoMinimoAceito ? 'Cobrança aceita — valor de até R$ 100,00, sem divergência.' : divergente ? 'Divergência identificada nesta cobrança.' : 'Cobrança correta. Nenhuma divergência.'}</div><div className="line"><span>Cálculo por peso (R$ 0,70/kg)</span><b>{brl(pesoCalc)}</b></div><div className="line"><span>Cálculo sobre a nota (1,5%)</span><b>{brl(notaCalc)}</b></div><div className="line"><span>Critério de referência</span><b>{criterio}</b></div><div className="line total"><span>Referência contratual</span><b>{brl(esperado)}</b></div><div className="line"><span>Frete cobrado</span><b>{brl(valorCobrado)}</b></div>{divergente && <div className="difference">Cobrança {situacaoValor}: <b>{brl(diferenca)}</b></div>}<button className="light" onClick={()=>setResultado(false)}>Nova conferência</button></>}</section></div>
      <section className="mail-panel mail-live"><div className="mail-icon">✉</div><div className="mail-main"><h3>Conexão com Microsoft Outlook</h3><p>Conecte a caixa corporativa para localizar mensagens do Siimp Sistemas e e-mails que mencionem “Danubio”. A consulta é somente leitura.</p><span className="pill">Siimp Sistemas</span> <span className="pill">Palavra-chave: Danubio</span> <span className="pill">Somente leitura</span>
        {mailError && <div className="mail-error">{mailError}</div>}
        {connected && <div className="mail-success">Conta conectada. {mailCount === null ? '' : `${mailCount} mensagem(ns) localizada(s).`}</div>}
        <div className="mail-actions">{!connected && <button onClick={()=>window.location.href='/api/outlook/connect'}>Conectar Microsoft Outlook</button>}<button className="light" onClick={()=>void syncEmails()} disabled={mailLoading}>{mailLoading ? 'Consultando mensagens...' : connected ? 'Atualizar e-mails' : 'Testar conexão'}</button></div>
        {emails.length > 0 && <div className="mail-list"><h4>Mensagens encontradas</h4>{emails.map(mail=><article key={mail.id} className="mail-item"><div><b>{mail.subject}</b><small>{mail.sender} · {new Date(mail.receivedAt).toLocaleString('pt-BR')}</small><p>{mail.preview}</p></div><div className="mail-meta">{mail.hasAttachments && <span className="pill">Com anexo</span>}{mail.webLink && <a href={mail.webLink} target="_blank" rel="noreferrer">Abrir e-mail</a>}</div></article>)}</div>}
      </div><span className={connected ? 'not-connected connected' : 'not-connected'}>{connected ? 'Conectado' : 'Aguardando conexão'}</span></section>
      <footer>ALEX FREIGHT AUDIT · Relatórios de divergência sujeitos à revisão humana.</footer>
    </section>
  </main>;
}
