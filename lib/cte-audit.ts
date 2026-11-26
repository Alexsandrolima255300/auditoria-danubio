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
  // O DACTE pode trazer primeiro o rótulo-resumo "TOMADOR DO SERVIÇO DESTINATÁRIO".
  // Só o campo explícito "TOMADOR DO SERVIÇO:" identifica o tomador real.
  // O CNPJ deve ser extraído do mesmo bloco, nunca de qualquer ocorrência no documento
  // (o CNPJ da Brasil pode aparecer como remetente sem ser o tomador).
  const tomadorMatch = text.match(/TOMADOR\\s+DO\\s+SERVI[ÇC]O\\s*:\\s*([\\s\\S]{0,1200})/i);
  const tomadorBlock = tomadorMatch?.[1] || "";
  const name = first(tomadorBlock, [
    /^(.+?)(?=\\s+ENDEREÇO\\s*:|\\s+CNPJ(?:\\/CPF)?\\s*:|\\s+IE\\s*:)/i,
  ]);
  const contextualCnpj = first(tomadorBlock, [
    /CNPJ(?:\\/CPF)?\\s*[:\\n-]*\\s*([\\d./-]{14,18})/i,
  ]);
  const cnpjDigits = contextualCnpj ? contextualCnpj.replace(/\\D/g, "") : "";
  const cnpj = cnpjDigits.length === 14 ? cnpjDigits : null;
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
    fields: { pdfPages: pdf.numPages, extractedTextLength: text.length, extractedTextSample: text.slice(0, 5000), tomadorBlock: tomadorBlock.slice(0, 1200), targetCompanyIsTomador: cnpj === "39860057000104" },
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
  if (data.invoiceValue === null || data.weightKg === null || data.freightGross === null || data.errors.length) {
    return { status: "needs_review", reference: null, discrepancy: null, reason: data.errors.join("; ") || "Dados insuficientes para auditoria." };
  }
  const reference = Math.round(Math.max(100, data.weightKg * 0.7, data.invoiceValue * 0.015) * 100) / 100;
  // A referência comercial corresponde ao frete total do CT-e. O ICMS é
  // exibido separadamente, mas não deve ser subtraído antes da comparação.
  if (data.freightGross <= 100) return { status: "accepted", reference, discrepancy: 0, reason: "Frete total de até R$ 100,00 aceito pela regra contratual." };
  const discrepancy = Math.round(Math.abs(data.freightGross - reference) * 100) / 100;
  // Tolerância de R$ 0,01 para diferenças de arredondamento.
  if (discrepancy > 0.01) return { status: "divergent", reference, discrepancy, reason: data.freightGross > reference ? "Frete total acima da referência calculada." : "Frete total abaixo da referência calculada." };
  return { status: "accepted", reference, discrepancy: 0, reason: discrepancy > 0 ? "Frete total confere dentro da tolerância de arredondamento de R$ 0,01." : "Frete total confere com a referência calculada." };
}
