'use client';

import { useState } from 'react';

const brl = (n: number) => n.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
const onlyDigits = (s: string) => s.replace(/\D/g, '');

export default function Home() {
  const [peso, setPeso] = useState('2,88');
  const [nota, setNota] = useState('697,99');
  const [cobrado, setCobrado] = useState('88,00');
  const [cnpj, setCnpj] = useState('39.860.057/0001-04');
  const [resultado, setResultado] = useState(false);
  const kg = Number(peso.replace(',', '.')) || 0;
  const nf = Number(nota.replace(',', '.')) || 0;
  const valorCobrado = Number(cobrado.replace(',', '.')) || 0;
  const pesoCalc = kg * 0.7;
  const notaCalc = nf * 0.015;
  const esperado = Math.max(100, pesoCalc, notaCalc);
  const tomadorValido = onlyDigits(cnpj) === '39860057000104';
  const divergente = Math.abs(valorCobrado - esperado) >= 0.01;

  return <main className="page">
    <aside className="sidebar"><div className="logo">AF</div><div className="brand">ALEX FREIGHT <b>AUDIT</b></div><p className="menu-label">MENU</p><div className="menu-item selected">Visão geral</div><div className="menu-item">Documentos</div><div className="menu-item">Divergências</div><div className="menu-item">E-mails</div><div className="side-note">Auditoria Danúbio<br/>Revisão humana antes de contestar</div></aside>
    <section className="content"><header><div><span className="overline">PAINEL DE CONTROLE</span><h1>Auditoria de fretes</h1><p>Confira as cobranças da transportadora Danúbio.</p></div><span className="status">● Transportadora: Danúbio</span></header>
      <div className="cards"><article><small>Transportadoras monitoradas</small><strong>01</strong><span>Danúbio</span></article><article><small>Regra de cálculo</small><strong>R$ 0,70</strong><span>Por kg ou 1,5% da nota</span></article><article><small>Frete mínimo</small><strong>R$ 100,00</strong><span>Valor mínimo contratado</span></article></div>
      <div className="section-title"><div><h2>Conferir um CT-e</h2><p>Simule a conferência usando os dados do documento.</p></div><span className="tag">REGRA DANÚBIO</span></div>
      <div className="columns"><section className="panel"><h3>Dados do documento</h3><label>CNPJ do tomador<input value={cnpj} onChange={e=>setCnpj(e.target.value)}/></label><div className="two"><label>Peso (kg)<input value={peso} onChange={e=>setPeso(e.target.value)}/></label><label>Valor da nota (R$)<input value={nota} onChange={e=>setNota(e.target.value)}/></label></div><label>Frete cobrado (sem ICMS) (R$)<input value={cobrado} onChange={e=>setCobrado(e.target.value)}/></label><button onClick={()=>setResultado(true)}>Conferir cobrança</button><small className="hint">Somente o tomador Brasil Engrenagens, CNPJ 39.860.057/0001-04, pode gerar alerta.</small></section>
      <section className="panel"><h3>Resultado da auditoria</h3>{!resultado ? <div className="empty"><div className="empty-icon">✓</div><b>Aguardando conferência</b><p>Preencha os dados ao lado para visualizar o cálculo.</p></div> : <><div className={tomadorValido && divergente ? 'alert warning' : 'alert good'}>{!tomadorValido ? 'Fora do escopo: tomador diferente da Brasil Engrenagens.' : divergente ? 'Divergência identificada nesta cobrança.' : 'Cobrança correta. Nenhuma divergência.'}</div><div className="line"><span>Frete por peso (R$ 0,70/kg)</span><b>{brl(pesoCalc)}</b></div><div className="line"><span>Percentual da nota (1,5%)</span><b>{brl(notaCalc)}</b></div><div className="line"><span>Maior dos cálculos</span><b>{brl(Math.max(pesoCalc, notaCalc))}</b></div><div className="line"><span>Mínimo aplicado</span><b>{brl(100)}</b></div><div className="line total"><span>Valor esperado</span><b>{brl(esperado)}</b></div><div className="line"><span>Frete cobrado (sem ICMS)</span><b>{brl(valorCobrado)}</b></div>{tomadorValido && <div className="difference">Diferença: <b>{brl(Math.abs(valorCobrado-esperado))}</b></div>}<button className="light" onClick={()=>setResultado(false)}>Nova conferência</button></>}</section></div>
      <section className="mail-panel"><div className="mail-icon">✉</div><div><h3>Monitoramento do Outlook</h3><p>O sistema será configurado para localizar e-mails do Siimp Sistemas e pesquisar também a palavra “Danubio” nas mensagens. A conexão ainda precisa ser ativada.</p><span className="pill">Siimp Sistemas</span> <span className="pill">Palavra-chave: Danubio</span> <span className="pill">Somente Danúbio</span></div><span className="not-connected">Não conectado</span></section>
      <footer>ALEX FREIGHT AUDIT · Relatórios de divergência sujeitos à revisão humana.</footer>
    </section>
  </main>;
}