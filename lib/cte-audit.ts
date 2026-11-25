import { getDocument, GlobalWorkerOptions } from "pdfjs-dist/legacy/build/pdf.mjs";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

// Em produção na Vercel, o worker precisa ser resolvido como arquivo local.
GlobalWorkerOptions.workerSrc = pathToFileURL(
  join(process.cwd(), "node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs")
).href;

export type ExtractedCte = {
  tomadorName: string | null;
  tomadorCnpj: string | null;
  invoiceValue: number | null;
  weightKg: number | null;
  freightGross: number | null;
  icmsValue: number | null;
  freightNet: number | null;
  fields: Record<string, unknown>;
  errors: string[];
};

function money(value: string): number | null {
  const cleaned = value.replace(/R\$\s*/gi, "").replace(/\s/g, "").replace(/\.(?=\d{3}(?:,|$))/g, "").replace(",", ".");
  const n = Number(cleaned);
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : null;
}
function first(text: string, patterns: RegExp[]): string | null {
  for (const re of patterns) {
    const match = text.match(re);
    if (match?.[1]) return match[1].trim();
  }
  return null;
}
function parseMoney(text: string, labels: string[]): number | null {
  for (const label of labels) {
    const re = new RegExp(label + "[^\\d]{0,35}(?:R\\$\\s*)?([\\d.]+,\\d{2}|\\d+,\\d{2})", "i");
    const match = text.match(re);
    if (match?.[1]) return money(match[1]);
  }
  return null;
}

export async function extractCte(buffer: Buffer): Promise<ExtractedCte> {
  const pdf = await getDocument({ data: new Uint8Array(buffer), useSystemFonts: true, isEvalSupported: false }).promise;
  const pages: string[] = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pages.push(content.items.map((item: any) => "str" in item ? item.str : "").join(" "));
  }
  const text = pages.join("\n").replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ");
  const name = first(text, [
    /TOMADOR(?: DO SERVICO| DO SERVIÇO)?\s*[:\n-]+\s*([^\n]{4,100})/i,
    /NOME\s*\/\s*RAZAO SOCIAL[^\n]*\n\s*([^\n]{4,100})/i,
  ]);
  // Nunca use o primeiro CNPJ do documento: normalmente ele pertence ao emitente.
  // Para a empresa auditada, confirme o CNPJ exato no texto integral do DACTE.
  const targetCnpj = "39860057000104";
  const digitsOnly = text.replace(/\D/g, "");
  const isTargetCompany = /BRASIL\s+ENGRENAGENS\s+E\s+CORRENTES/i.test(text);
  let cnpj: string | null = null;
  if (isTargetCompany && digitsOnly.includes(targetCnpj)) {
    cnpj = targetCnpj;
  } else if (!isTargetCompany) {
    // Para outros tomadores, só aceite CNPJ próximo ao rótulo TOMADOR.
    const tomadorIndex = text.search(/TOMADOR(?:\s+DO\s+SERVI[ÇC]O)?/i);
    if (tomadorIndex >= 0) {
      const tomadorBlock = text.slice(tomadorIndex, tomadorIndex + 900);
      const contextualCnpj = first(tomadorBlock, [
        /CNPJ\s*[:\n-]*\s*([\d./-]{14,18})/i,
        /\b(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})\b/,
      ]);
      cnpj = contextualCnpj ? contextualCnpj.replace(/\D/g, "") : null;
    }
  }
  const invoiceValue = parseMoney(text, [
    "VALOR TOTAL DAS MERCADORIAS", "VALOR TOTAL DA MERCADORIA",
    "VALOR DA MERCADORIA", "VALOR DA CARGA", "VALOR TOTAL DA CARGA",
    "VALOR TOTAL DA NOTA", "VALOR DA NOTA",
  ]);
  const weightRaw = first(text, [
    // DACTE da Danubio costuma extrair como: "PESO BRUTO (KG) 87,4800".
    // Aceita valor após espaço, dois-pontos, hífen ou quebra de linha.
    /PESO(?:\s+BRUTO)?(?:\s+TOTAL)?(?:\s+DA\s+CARGA)?(?:\s*\(\s*KG\s*\))?\s*(?:[:\n-]\s*|\s+)([\d.,]+)\s*(?:KG|KGS)?/i,
    /PESO\s+REAL(?:\s*\(\s*KG\s*\))?\s*(?:[:\n-]\s*|\s+)([\d.,]+)\s*(?:KG|KGS)?/i,
    /PESO\s+BRUTO\s+KG\s+([\d.,]+)/i,
  ]);
  const weightKg = weightRaw ? Number(weightRaw.replace(/\./g, "").replace(",", ".")) : null;
  const freightGross = parseMoney(text, ["VALOR A RECEBER", "VALOR TOTAL DA PRESTACAO DO SERVICO", "VALOR TOTAL DA PRESTAÇÃO DO SERVIÇO", "VALOR TOTAL DO FRETE"]);
  const icmsValue = parseMoney(text, ["VALOR DO ICMS", "VALOR ICMS"]);
  const freightNet = freightGross !== null && icmsValue !== null ? Math.round((freightGross - icmsValue) * 100) / 100 : freightGross;
  const errors: string[] = [];
  if (!cnpj) errors.push("CNPJ do tomador não identificado");
  if (invoiceValue === null) errors.push("Valor da carga/nota não identificado");
  if (weightKg === null || !Number.isFinite(weightKg)) errors.push("Peso em kg não identificado");
  if (freightNet === null) errors.push("Valor do frete não identificado");
  if (freightGross !== null && icmsValue === null) errors.push("ICMS não identificado; confirmar se o frete extraído está sem ICMS");
  return {
    tomadorName: name, tomadorCnpj: cnpj, invoiceValue,
    weightKg: Number.isFinite(weightKg as number) ? weightKg : null,
    freightGross, icmsValue, freightNet,
    fields: { pdfPages: pdf.numPages, extractedTextLength: text.length, extractedTextSample: text.slice(0, 5000), targetCompanyNameFound: isTargetCompany, targetCnpjFound: digitsOnly.includes(targetCnpj) },
    errors,
  };
}

export function evaluateFreight(data: ExtractedCte) {
  if (!data.tomadorCnpj) {
    return { status: "needs_review", reference: null, discrepancy: null, reason: "CNPJ do tomador não identificado com segurança." };
  }
  if (data.tomadorCnpj !== "39860057000104") {
    return { status: "out_of_scope", reference: null, discrepancy: null, reason: "Tomador diferente da Brasil Engrenagens." };
  }
  if (data.invoiceValue === null || data.weightKg === null || data.freightNet === null || data.errors.length) {
    return { status: "needs_review", reference: null, discrepancy: null, reason: data.errors.join("; ") || "Dados insuficientes para auditoria." };
  }
  const reference = Math.round(Math.max(100, data.weightKg * 0.7, data.invoiceValue * 0.015) * 100) / 100;
  if (data.freightNet <= 100) return { status: "accepted", reference, discrepancy: 0, reason: "Frete de até R$ 100,00 aceito pela regra contratual." };
  const discrepancy = Math.round(Math.abs(data.freightNet - reference) * 100) / 100;
  if (discrepancy >= 0.01) return { status: "divergent", reference, discrepancy, reason: data.freightNet > reference ? "Frete acima da referência calculada." : "Frete abaixo da referência calculada." };
  return { status: "accepted", reference, discrepancy: 0, reason: "Frete confere com a referência calculada." };
}
