import pdfParse from "pdf-parse";

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
  const parsed = await pdfParse(buffer);
  const text = (parsed.text || "").replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ");
  const name = first(text, [
    /TOMADOR(?: DO SERVICO| DO SERVIÇO)?\s*[:\n-]+\s*([^\n]{4,100})/i,
    /NOME\s*\/\s*RAZAO SOCIAL[^\n]*\n\s*([^\n]{4,100})/i,
  ]);
  const cnpjRaw = first(text, [
    /CNPJ\s*[:\n-]*\s*([\d./-]{14,18})/i,
    /\b(\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2})\b/,
  ]);
  const cnpj = cnpjRaw ? cnpjRaw.replace(/\D/g, "") : null;
  const invoiceValue = parseMoney(text, ["VALOR DA CARGA", "VALOR DA MERCADORIA", "VALOR TOTAL DAS MERCADORIAS", "VALOR DA NOTA"]);
  const weightRaw = first(text, [
    /PESO(?: BRUTO| TOTAL)?(?: DA CARGA)?\s*[:\n-]+\s*([\d.,]+)\s*(?:KG|KGS)?/i,
    /PESO REAL\s*[:\n-]+\s*([\d.,]+)/i,
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
    fields: { pdfPages: parsed.numpages, extractedTextLength: text.length, extractedTextSample: text.slice(0, 1000) },
    errors,
  };
}

export function evaluateFreight(data: ExtractedCte) {
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
