import { NextRequest, NextResponse } from "next/server";
import { openToken, refreshMicrosoftToken } from "../../../../lib/microsoft";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist/legacy/build/pdf.mjs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 300;
GlobalWorkerOptions.workerSrc = pathToFileURL(join(process.cwd(), "node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs")).href;
const PROJECT_URL = process.env.SUPABASE_URL || "https://shzyfxkjsssjpmesvhxu.supabase.co";
const KEY = () => process.env.SUPABASE_SERVICE_ROLE_KEY || "";
async function db(path: string, init: RequestInit = {}) {
  const key = KEY(); if (!key) throw new Error("SUPABASE_SERVICE_ROLE_KEY não configurada.");
  const r = await fetch(PROJECT_URL + "/rest/v1/" + path, { ...init, headers: { apikey:key, Authorization:"Bearer "+key, "Content-Type":"application/json", ...(init.headers||{}) }, cache:"no-store" });
  const t = await r.text(); if (!r.ok) throw new Error("Supabase: "+t.slice(0,500)); return t ? JSON.parse(t) : [];
}
async function graph(token:string,path:string) {
  const r=await fetch("https://graph.microsoft.com/v1.0"+path,{headers:{Authorization:"Bearer "+token},cache:"no-store"});
  const d=await r.json(); if(!r.ok) throw new Error(d.error?.message||"Erro Microsoft Graph"); return d;
}
const num=(s:string)=>{const n=Number(s.replace(/\./g,"").replace(",", "."));return Number.isFinite(n)?n:null};
const money=(s:string)=>num(s.replace(/^R\$\s*/,""));
const date=(s:string)=>{const m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);return m?m[3]+"-"+m[2]+"-"+m[1]:null};
async function pdfText(buffer:Buffer) {
  const pdf=await getDocument({data:new Uint8Array(buffer),useSystemFonts:true,isEvalSupported:false}).promise;
  const pages:string[]=[];
  for(let p=1;p<=pdf.numPages;p++){
    const page=await pdf.getPage(p);const c=await page.getTextContent();
    const rows=new Map<number,{x:number,s:string}[]>();
    for(const item of c.items as any[]){if(!("str" in item)||!item.str.trim())continue;const y=Math.round(item.transform?.[5]||0);const a=rows.get(y)||[];a.push({x:item.transform?.[4]||0,s:item.str});rows.set(y,a);}
    pages.push([...rows.entries()].sort((a,b)=>b[0]-a[0]).map(([,a])=>a.sort((x,y)=>x.x-y.x).map(x=>x.s).join(" ").replace(/\s+/g," ").trim()).join("\n"));
  }
  return {text:pages.join("\n"),pages:pdf.numPages};
}
function parseInvoice(text:string) {
  const normalized=text.replace(/\u00a0/g," ").replace(/[ \t]+/g," ");
  // A Danúbio posiciona o número da fatura em uma coluna separada do título.
  // Capturamos o número junto ao status EMITIDA ou na linha DATA DE EMISSÃO.
  const headerEnd=normalized.search(/\bDOCUMENTOS\b/i);
  const invoiceHeader=headerEnd>=0?normalized.slice(0,headerEnd):normalized.slice(0,2500);
  const headerDateAndNumber=invoiceHeader.match(/\b(\d{2}\/\d{2}\/\d{4})\s+(\d{4,})\b/);
  const invoiceNo=invoiceHeader.match(/\b(\d{4,})\s+EMITIDA\b/i)?.[1] ||
    headerDateAndNumber?.[2] ||
    invoiceHeader.match(/N[º°]?\s*FATURA[\s\S]{0,250}?(\d{4,})\b/i)?.[1] || null;
  const issued=headerDateAndNumber?.[1] ||
    invoiceHeader.match(/DATA DE\s+EMISS[ÃA]O[\s\S]{0,500}?(\d{2}\/\d{2}\/\d{4})/i)?.[1] || null;
  const due=normalized.match(/VENCIMENTO[\s\S]{0,80}?(\d{2}\/\d{2}\/\d{4})/i)?.[1] || null;
  const total=normalized.match(/TOTAL\s+VENCIMENTO\s+R\$\s*([\d.]+,\d{2})/i)?.[1] || normalized.match(/VALOR DA FATURA[\s\S]{0,100}?R\$\s*([\d.]+,\d{2})/i)?.[1] || null;
  const totalWeight=normalized.match(/([\d.]+,\d{3})\s*KG/i)?.[1] || null;
  const totalServices=normalized.match(/TOTAL SERVI[ÇC]O[\s\S]{0,120}?R\$\s*([\d.]+,\d{2})/i)?.[1] || null;
  const totalIcms=normalized.match(/TOTALIZADORES[\s\S]{0,500}?ICMS[\s\S]{0,120}?R\$\s*([\d.]+,\d{2})/i)?.[1] || null;
  const docCount=normalized.match(/QTD\.\s*DOC\.\s+(\d+)/i)?.[1] || null;
  const rows=normalized.split("\n");
  const items:any[]=[]; let current:any=null;
  const full=/^(\d{2}\/\d{2}\/\d{4})\s+CT-?E\s+(\d+\/\d+)\s+(\d+\/\d+)\s+([\d.,]+)\s+R\$\s*([\d.,]+)\s+R\$\s*([\d.,]+)\s+R\$\s*([\d.,]+)\s+R\$\s*([\d.,]+)\s+R\$\s*([\d.,]+)\s+(\d{2}\/\d{2}\/\d{4})/i;
  const continuation=/^(\d+\/\d+)\s+([\d.,]+)\s+(\d{2}\/\d{2}\/\d{4})$/;
  for(const line0 of rows){
    const line=line0.replace(/\s+/g," ").trim(); const m=line.match(full);
    if(m){
      current={row_number:items.length+1,issued_at:date(m[1]),cte_number:m[2],invoice_number:m[3],weight_kg:num(m[4]),gris:money(m[5]),insurance:money(m[6]),toll:money(m[7]),icms:money(m[8]),charged_total:money(m[9]),raw_row:line};
      items.push(current);continue;
    }
    const c=line.match(continuation);
    if(c&&current&&line.length<80){
      current.weight_kg=Math.round(((current.weight_kg||0)+(num(c[2])||0))*1000)/1000;
      current.raw_row+=" | "+line;
    }
  }
  return {invoiceNo,issued:date(issued||""),due:date(due||""),total:money(total||""),totalWeight:num(totalWeight||""),totalServices:money(totalServices||""),totalIcms:money(totalIcms||""),docCount:docCount?Number(docCount):null,items,pages:0,textLength:text.length};
}
function normCte(v:string){return String(v||"").replace(/\s/g,"").toUpperCase();}
function normCompany(v:string){return String(v||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/[^A-Z0-9]+/g," ").replace(/\s+/g," ").trim();}
const DANUBIO_COMPANY=normCompany("DANUBIO AZUL TRANSP. DE CARGAS E ENCOMENDAS LTDA (MTZ)");
export async function GET(req:NextRequest){
  const cookie=req.cookies.get("outlook_refresh")?.value;
  if(!cookie)return NextResponse.json({error:"Outlook ainda não conectado."},{status:401});
  try{
    const connections=await db("outlook_connections?select=id&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if(!connections.length)throw new Error("Conexão Outlook não encontrada.");
    const cutoff=new Date(Date.now()-45*86400000).toISOString();
    const saved=await db("danubio_invoices?select=*,items:danubio_invoice_items(*)&connection_id=eq."+connections[0].id+"&created_at=gte."+encodeURIComponent(cutoff)+"&order=created_at.desc&limit=200");
    return NextResponse.json({ok:true,invoices:saved});
  }catch(e){return NextResponse.json({error:e instanceof Error?e.message:"Falha ao carregar faturas."},{status:500});}
}
export async function POST(req:NextRequest){
  const cookie=req.cookies.get("outlook_refresh")?.value;
  if(!cookie)return NextResponse.json({error:"Outlook ainda não conectado."},{status:401});
  try{
    const connections=await db("outlook_connections?select=id,encrypted_token_cache&revoked_at=is.null&order=last_refreshed_at.desc&limit=1");
    if(!connections.length)throw new Error("Conexão Outlook não encontrada.");
    const connection=connections[0];const refresh=openToken(cookie);const tokens=await refreshMicrosoftToken(refresh);
    const cutoff=new Date(Date.now()-45*86400000).toISOString();
    const mails=await db("email_messages?select=id,graph_message_id,subject,received_at,has_attachments&has_attachments=eq.true&connection_id=eq."+connection.id+"&received_at=gte."+encodeURIComponent(cutoff)+"&order=received_at.desc&limit=1000");
    const offset = Math.max(0, Number(req.nextUrl.searchParams.get("offset") || 0));
    const batchSize = Math.min(10, Math.max(1, Number(req.nextUrl.searchParams.get("batch") || 5)));
    const batchMails = mails.slice(offset, offset + batchSize);
    const invoices:any[]=[];
    let completedInvoicesSkipped=0;
    for(const mail of batchMails){
      // Se este e-mail já gerou uma ou mais faturas totalmente processadas
      // pela versão atual do extrator, não precisamos nem consultar os anexos novamente.
      const completedForEmail=await db(
        "danubio_invoices?select=id,processing_status,extraction_data&connection_id=eq."+
        connection.id+"&email_message_id=eq."+encodeURIComponent(mail.id)+"&limit=50"
      );
      const emailAlreadyComplete=completedForEmail.length>0&&completedForEmail.every((invoice:any)=>
        invoice.processing_status==="processed"&&Number(invoice.extraction_data?.parserVersion||0)>=5
      );
      if(emailAlreadyComplete){
        completedInvoicesSkipped++;
        invoices.push({email:mail.subject,alreadyProcessed:true,reason:"Todas as faturas deste e-mail já estão concluídas."});
        continue;
      }
      const atts=await graph(tokens.access_token,"/me/messages/"+encodeURIComponent(mail.graph_message_id)+"/attachments?$top=50");
      for(const a of atts.value||[]){
        if(a.isInline||!((a.contentType||"").toLowerCase().includes("pdf")||(a.name||"").toLowerCase().endsWith(".pdf")))continue;
        const exists=await db("danubio_invoices?select=id,processing_status&connection_id=eq."+connection.id+"&graph_attachment_id=eq."+encodeURIComponent(a.id)+"&limit=1");
        if(exists.length&&exists[0].processing_status==="processed"){const prior=await db("danubio_invoices?select=extraction_data&id=eq."+exists[0].id+"&limit=1");if(Number(prior[0]?.extraction_data?.parserVersion||0)>=5){invoices.push({id:exists[0].id,alreadyProcessed:true});continue;}}
        const file=await graph(tokens.access_token,"/me/messages/"+encodeURIComponent(mail.graph_message_id)+"/attachments/"+encodeURIComponent(a.id));
        if(!file.contentBytes)continue;
        const buffer=Buffer.from(file.contentBytes,"base64");if(buffer.length>20*1024*1024)continue;
        const extracted=await pdfText(buffer);
        if(!normCompany(extracted.text).includes(DANUBIO_COMPANY))continue;
        const parsed=parseInvoice(extracted.text);parsed.pages=extracted.pages;
        if(!parsed.invoiceNo&&!parsed.items.length)continue;
        const invoiceRows=await db("danubio_invoices?on_conflict=connection_id,graph_attachment_id",{
          method:"POST",headers:{Prefer:"resolution=merge-duplicates,return=representation"},
          body:JSON.stringify({connection_id:connection.id,email_message_id:mail.id,graph_attachment_id:a.id,file_name:a.name||"fatura.pdf",invoice_number:parsed.invoiceNo,issued_at:parsed.issued,due_at:parsed.due,document_count:parsed.docCount,total_weight:parsed.totalWeight,total_services:parsed.totalServices,total_icms:parsed.totalIcms,total_invoice:parsed.total,processing_status:"needs_review",extraction_data:{parserVersion:5,pages:parsed.pages,textLength:parsed.textLength}})
        });
        const invoiceId=invoiceRows[0].id;
        const itemDates=parsed.items.map((x:any)=>x.issued_at).filter(Boolean).sort();
        const periodStart=itemDates[0]||null;
        const periodEnd=itemDates[itemDates.length-1]||null;
        await db("danubio_invoices?id=eq."+invoiceId,{method:"PATCH",headers:{Prefer:"return=minimal"},body:JSON.stringify({period_start:periodStart,period_end:periodEnd})});
        await db("danubio_invoice_items?invoice_id=eq."+invoiceId,{method:"DELETE"});
        const cteDocs=await db("freight_audits?select=id,attachment_id,weight_kg,extracted_fields,freight_charge&limit=5000");
        const attIds=[...new Set(cteDocs.map((x:any)=>x.attachment_id))];
        const attRows=attIds.length?await db("freight_attachments?select=id,email_message_id&id=in.("+attIds.join(",")+")"): [];
        const emailIds=[...new Set(attRows.map((x:any)=>x.email_message_id))];
        const emailRows=emailIds.length?await db("email_messages?select=id,connection_id&id=in.("+emailIds.join(",")+")"): [];
        const validEmails=new Set(emailRows.filter((x:any)=>x.connection_id===connection.id).map((x:any)=>x.id));
        const validAtts=new Set(attRows.filter((x:any)=>validEmails.has(x.email_message_id)).map((x:any)=>x.id));
        const auditCandidates=cteDocs.filter((x:any)=>validAtts.has(x.attachment_id));
        const source:any[]=[];
        for(const x of auditCandidates){
          const raw=String(x.extracted_fields?.extractedTextSample||"");
          const header=raw.match(/CTE\s*N[º°]?\s*DOCUMENTO:\s*(\d+)\s+S[ÉE]RIE:\s*(\d+)/i);
          const cte=header?header[2]+"/"+header[1]:(raw.match(/CTE\s*(?:N[º°]|NUMERO)?\s*[:.]?\s*(\d+\/\d+)/i)?.[1]||null);
          const issuedMatch=raw.match(/DATA\/?HORA\s+EMISS[AÃ]O\s+(\d{2}\/\d{2}\/\d{4})/i);
          const sourceIssued=issuedMatch?date(issuedMatch[1]):null;
          if(cte&&sourceIssued)source.push({...x,cte,sourceIssued});
        }
        const payload:any[]=[];
        const seen=new Set<string>();
        for(const item of parsed.items){
          const matches=source.filter((x:any)=>normCte(x.cte)===normCte(item.cte_number)&&(!item.issued_at||x.sourceIssued===item.issued_at)&&(!periodStart||x.sourceIssued>=periodStart)&&(!periodEnd||x.sourceIssued<=periodEnd));
          const original=matches[0]||null;
          const weightDiff=original&&original.weight_kg!=null&&item.weight_kg!=null?Math.round((item.weight_kg-Number(original.weight_kg))*1000)/1000:null;
          const originalTotal=original?(original.extracted_fields?.grossFreight??original.freight_charge):null;
          const valueDiff=originalTotal!=null&&item.charged_total!=null?Math.round((item.charged_total-Number(originalTotal))*100)/100:null;
          const duplicate=seen.has(normCte(item.cte_number));seen.add(normCte(item.cte_number));
          let status="not_found";
          if(duplicate)status="duplicate";
          else if(original){
            const wd=weightDiff!==null&&Math.abs(weightDiff)>0.001;const vd=valueDiff!==null&&Math.abs(valueDiff)>0.01;
            status=wd&&vd?"both_different":wd?"weight_difference":vd?"value_difference":"matched";
          }
          payload.push({...item,invoice_id:invoiceId,original_attachment_id:original?.attachment_id||null,original_weight_kg:original?.weight_kg??null,original_freight_total:originalTotal,weight_difference:weightDiff,value_difference:valueDiff,comparison_status:status});
        }
        for(let i=0;i<payload.length;i+=100){
          await db("danubio_invoice_items",{method:"POST",headers:{Prefer:"return=minimal"},body:JSON.stringify(payload.slice(i,i+100))});
        }
        const statuses=payload.reduce((o:any,x:any)=>(o[x.comparison_status]=(o[x.comparison_status]||0)+1,o),{});
        const totalsCheck={sumItems:Math.round(payload.reduce((s:number,x:any)=>s+(x.charged_total||0),0)*100)/100,weightItems:Math.round(payload.reduce((s:number,x:any)=>s+(x.weight_kg||0),0)*1000)/1000,countItems:payload.length};
        const processed=parsed.items.length>0&&parsed.items.every((x:any)=>x.cte_number&&x.weight_kg!=null&&x.charged_total!=null);
        await db("danubio_invoices?id=eq."+invoiceId,{method:"PATCH",headers:{Prefer:"return=minimal"},body:JSON.stringify({processing_status:processed?"processed":"needs_review",extraction_data:{parserVersion:5,pages:parsed.pages,textLength:parsed.textLength,statuses,totalsCheck,parsedRows:payload.length,periodStart,periodEnd,periodSource:itemDates.length?"invoice_item_issue_dates":"not_identified"}})});
        invoices.push({id:invoiceId,invoiceNumber:parsed.invoiceNo,fileName:a.name,documentCount:parsed.docCount,parsedRows:payload.length,total:parsed.total,totalWeight:parsed.totalWeight,statuses,totalsCheck,processingStatus:processed?"processed":"needs_review"});
      }
    }
    const saved=await db("danubio_invoices?select=*,items:danubio_invoice_items(*)&connection_id=eq."+connection.id+"&created_at=gte."+encodeURIComponent(cutoff)+"&order=created_at.desc&limit=200");
    return NextResponse.json({ok:true,found:invoices.length,invoices:saved,processedNow:invoices,scannedEmails:batchMails.length,totalEmails:mails.length,completedInvoicesSkipped,nextOffset:offset+batchMails.length,hasMore:offset+batchMails.length<mails.length});
  }catch(e){return NextResponse.json({error:e instanceof Error?e.message:"Falha ao conferir faturas."},{status:500});}
}
