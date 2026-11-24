// Pure pricing and parsing rules for the principal negotiation (no server state).
// Everything the Nov 2026 review asked for in "Angka dan harga" lives here so it can be tested alone
// (npm run test:pricing).

export const DEFAULT_PPN = 0.11;

// ---- text helpers ------------------------------------------------------------------------
// Pieces per PO unit read from text such as "BOX@100", "PACK/50", "10 STRIP @ 10 TAB"; null when none is found.
export function matchUoMQty(text?: string | null): number | null {
  if (!text) return null;
  const s = String(text).trim();
  const multi = s.match(/(\d+)\s*(?:strip|blister|alu|tube|botol|vial|amp\w*|sachet|bag)\s*(?:@|x|isi|\*)\s*(\d+)/i);
  if (multi) return Math.max(1, parseInt(multi[1], 10) * parseInt(multi[2], 10));
  const m =
    s.match(/(?:@|\/|isi\s*|x\s*|:\s*|=\s*|of\s*)(\d+)/i) ||
    s.match(/\b(?:box|pack|pak|dus|btl|bag|strip)\s*(\d+)\b/i) ||
    s.match(/\b(\d+)\s*(?:pcs|pc|tab\w*|cap\w*|amp\w*|vial\w*|strip\w*|roll\w*|test\w*|pair\w*|bag\w*|btl|botol|sachet|syringe|unit|lembar|set)/i);
  return m ? Math.max(1, parseInt(m[1], 10)) : null;
}
export function parseUoMQty(text?: string | null): number {
  return matchUoMQty(text) ?? 1;
}

// "BOX@12", "Box 12", "box@12", "Box isi 12" -> "BOX@12". One spelling per pack so the same item is never
// counted as two units.
export function normalizeUom(text?: string | null): string {
  const s = String(text ?? "").trim();
  if (!s) return "";
  const up = s.toUpperCase();
  const qty = matchUoMQty(up);
  const words = (up.match(/[A-Z]+/g) || []).filter((w) => !["ISI", "X", "OF", "PCS", "PC"].includes(w));
  const unit = words[0] || "";
  if (unit && qty) return `${unit}@${qty}`;
  if (qty && !unit) return `@${qty}`;
  return up.replace(/\s+/g, " ");
}

// ERP codes: trimmed and upper case (" dum000006 " -> "DUM000006").
export const normalizeCode = (v: any): string => String(v ?? "").trim().toUpperCase();

// Principal names are stored in upper case ("pt dummy uji farma" -> "PT DUMMY UJI FARMA").
export const normalizePrincipalName = (v: any): string => String(v ?? "").replace(/\s+/g, " ").trim().toUpperCase();

// "Rp 112.500" -> 112500; "1.234,5" -> 1234.5; "15%" -> 15. null when it is not a number.
export function parseMoney(v: any): number | null {
  if (v == null) return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;
  let s = String(v).trim().replace(/rp\.?/i, "").replace(/%/g, "").replace(/\s/g, "");
  if (!s) return null;
  const lastDot = s.lastIndexOf(".");
  const lastComma = s.lastIndexOf(",");
  if (lastDot >= 0 && lastComma >= 0) s = lastComma > lastDot ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  else if (lastComma >= 0) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, "") : s.replace(",", ".");
  else if (lastDot >= 0 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, "");
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

// ---- discount: one format everywhere -----------------------------------------------------
// The principal types a number from 0 to 100 (the % sign is optional and added on screen). Up to two
// decimals ("27,5" -> 27.5%). Returns the discount as a fraction (0.275). Blank means "not filled".
export interface DiscountResult { value: number | null; error?: string }
export function parseDiscount(raw: any): DiscountResult {
  if (raw == null) return { value: null };
  if (typeof raw === "number") {
    if (!Number.isFinite(raw)) return { value: null, error: "Diskon bukan angka." };
    return checkDisc(raw);
  }
  const s = String(raw).trim().replace(/\s/g, "");
  if (s === "") return { value: null };
  const body = s.replace(/%/g, "");
  if (!/^\d+([.,]\d+)?$/.test(body)) return { value: null, error: `Diskon "${String(raw).trim()}" bukan angka. Isi angka 0 sampai 100, contoh 15 atau 27,5.` };
  return checkDisc(parseFloat(body.replace(",", ".")));
}
function checkDisc(percent: number): DiscountResult {
  if (percent < 0) return { value: null, error: "Diskon tidak boleh negatif." };
  if (percent >= 100) return { value: null, error: `Diskon ${percent}% tidak wajar. Isi angka antara 0 dan 100 (contoh 15 untuk 15%).` };
  const rounded = Math.round(percent * 100) / 100;
  return { value: Math.round(rounded * 100) / 10000 };
}
// Excel percent-formatted cells hold 0.275 for 27.5%; a number typed into a General cell is already a percent.
export function discountFromCell(value: number, percentFormat: boolean): DiscountResult {
  return checkDisc(percentFormat ? value * 100 : value);
}
export function fmtPct(frac: number | null | undefined): string {
  if (frac == null) return "";
  const p = Math.round(frac * 10000) / 100;
  return `${String(p).replace(".", ",")}%`;
}

// ---- cost and margin ---------------------------------------------------------------------
export type Binding = "Nett" | "Disc" | "Nett to Disc";

// HNA per piece (before PPN) the price is built from.
// Nett and Disc: the principal's RFQ HNA. Nett to Disc: the formulary HNA (A4), else the average 2026 PO price.
export function priceHnaPc(it: any, binding: string, formHnaPc: number | null, poAvgPc: number | null): number | null {
  const rfqPc = it.rfq_hna != null && it.rfq_qty > 0 ? it.rfq_hna / it.rfq_qty : null;
  if (binding === "Nett to Disc") return formHnaPc ?? poAvgPc ?? rfqPc;
  return rfqPc;
}

export const STEP_KEYS = ["rfq", "co", "fb1", "on"] as const;
const STEP_DISC: Record<string, string> = { rfq: "rfq_disc", co: "co_disc", fb1: "fb1_disc", on: "on_disc" };

export function enrichCycleItem(item: any, ppn = DEFAULT_PPN, cycle?: any) {
  const it: any = { ...item };
  const binding: string = cycle?.binding || it.binding || "Nett";
  const negoType: string = cycle?.nego_type || it.nego_type || "Renewal";
  const monthsElapsed: number = cycle?.months_elapsed || 12;
  const vat = 1 + ppn;

  // Pack sizes. PO quantities count PO units; the formulary has its own conversion.
  const poPack = matchUoMQty(it.po_unit_text) || it.form_conv || it.mou_qty || it.rfq_qty || 1;
  const formConv = it.form_conv || poPack;
  it.po_unit_pcs = poPack;

  // ---- formulary: Base Price (no discount) and Nett Price (after discount), per piece incl. PPN
  const formHnaPc: number | null = it.formulary_hna ? it.formulary_hna / formConv : null;
  const formDisc: number = it.form_disc ?? 0;
  it.form_hna_pc = formHnaPc;
  it.form_base_price = formHnaPc != null ? formHnaPc * vat : (it.form_base_price_given ?? null);
  it.form_nett_price = formHnaPc != null ? formHnaPc * (1 - formDisc) * vat : (it.form_price_given ?? null);
  it.form_margin_unit = it.form_base_price != null && it.form_nett_price != null ? it.form_base_price - it.form_nett_price : null;

  // ---- PO: one price everywhere, the average of the PO year (incl. PPN)
  const poAvgPc: number | null = it.po_avg_price ? it.po_avg_price / poPack : null;
  it.po_avg_pc = poAvgPc;
  it.po_unit_price = poAvgPc != null ? poAvgPc * vat : null;
  it.annual_pcs = it.po_qty_12m ? (it.po_qty_12m * poPack * 12) / monthsElapsed : null;
  it.po_annualized_value = ((it.po_value_12m || 0) * 12) / monthsElapsed;          // before PPN
  it.po_annualized_value_ppn = it.po_annualized_value * vat;                      // "Nilai PO setahun (+PPN)"

  // ---- MOU: only when the SKU really is in the MOU
  it.in_mou = it.mou_hna != null && it.mou_qty > 0;
  it.mou_unit_price = it.in_mou ? (it.mou_hna / it.mou_qty) * (1 - (it.mou_disc || 0)) * vat : null;

  // ---- step prices
  const hnaPc = priceHnaPc(it, binding, formHnaPc, poAvgPc);
  const hasRfq = it.rfq_hna != null && it.rfq_qty > 0;
  const priceAt = (disc: number | null | undefined, schemeBinding = binding): number | null => {
    if (!hasRfq || disc == null) return null;
    const base = priceHnaPc(it, schemeBinding, formHnaPc, poAvgPc);
    return base != null ? base * (1 - disc) * vat : null;
  };
  it.rfq_unit_price = hasRfq ? priceAt(it.rfq_disc ?? 0) : null;
  it.co_unit_price = priceAt(it.co_disc);
  it.fb1_unit_price = priceAt(it.fb1_disc);
  it.on_unit_price = priceAt(it.on_disc);
  // Online Nego can hold one discount per scheme: Nett (price from the RFQ HNA) and Disc (price from the formulary HNA).
  it.on_nett_unit_price = priceAt(it.on_disc_nett, "Nett");
  it.on_disc_unit_price = priceAt(it.on_disc_disc, "Disc");

  // ---- comparison (acuan): Nett Price formulary, else MOU, else average PO price
  const acuan = it.form_nett_price ?? it.mou_unit_price ?? it.po_unit_price ?? null;
  const acuanTag = it.form_nett_price != null ? "Formularium" : it.mou_unit_price != null ? "MOU" : it.po_unit_price != null ? "PO" : "tanpa acuan";
  it.primary_bench = acuan;
  it.bench_tag = acuanTag;
  it.base_from = acuanTag;
  it.base_price_cmp = acuan;

  // ---- margin base price per piece incl. PPN
  // Nett: RFQ HNA. Disc and Nett to Disc: formulary HNA (empty: average PO price).
  let baseHnaPc: number | null;
  if (binding === "Nett") baseHnaPc = it.rfq_hna != null && it.rfq_qty > 0 ? it.rfq_hna / it.rfq_qty : it.in_mou ? it.mou_hna / it.mou_qty : formHnaPc ?? poAvgPc;
  else baseHnaPc = formHnaPc ?? poAvgPc;
  const baseMargin = baseHnaPc != null ? baseHnaPc * vat : null;
  it.hna_per_pc = baseHnaPc;
  it.base_margin_price = baseMargin;
  it.base_price = baseMargin;

  // ---- per step: cost impact and margin
  const poValue = it.po_annualized_value_ppn;
  const formMarginUnit = (): number | null => {
    if (it.form_margin_unit != null) return it.form_margin_unit;
    return baseMargin != null && acuan != null ? baseMargin - acuan : null;
  };
  const metrics = (price: number | null) => {
    if (price == null) return null;
    const annual = it.annual_pcs;
    const totalCost = annual ? price * annual : null;
    let costImpact: number | null = null;
    if (annual) {
      if (negoType === "Force Majeure" && it.mou_unit_price != null) costImpact = (price - it.mou_unit_price) * annual;
      else if (it.po_annualized_value) costImpact = price * annual - poValue;
    }
    const marginPc = baseMargin != null ? baseMargin - price : null;
    const totalMargin = marginPc != null && annual ? marginPc * annual : null;
    const fm = formMarginUnit();
    const poMarginUnit = baseMargin != null && it.po_unit_price != null ? baseMargin - it.po_unit_price : null;
    return {
      price, total_cost: totalCost, cost_impact: costImpact,
      margin_pc: marginPc, total_margin: totalMargin,
      margin_impact_form: marginPc != null && fm != null && annual ? (marginPc - fm) * annual : null,
      margin_impact_po: marginPc != null && poMarginUnit != null && annual ? (marginPc - poMarginUnit) * annual : null,
    };
  };
  const byStep: Record<string, any> = {};
  for (const k of STEP_KEYS) byStep[k] = metrics(it[`${k}_unit_price`]);
  // Pre-Nego = result up to Feedback I; Post-Nego = the agreed (FIN) result
  const preKey = (["fb1", "co", "rfq"] as const).find((k) => byStep[k]);
  byStep.pre = preKey ? { ...byStep[preKey], step: preKey } : null;
  byStep.post = byStep.on ? { ...byStep.on, step: "on" } : null;
  const latestKey = (["on", "fb1", "co", "rfq"] as const).find((k) => byStep[k]);
  byStep.latest = latestKey ? { ...byStep[latestKey], step: latestKey } : null;
  it.m = byStep;

  for (const k of STEP_KEYS) {
    const m = byStep[k];
    it[`impact_${k}`] = m?.cost_impact ?? null;
    it[`margin_pc_${k}`] = m?.margin_pc ?? null;
    it[`margin_annual_${k}`] = m?.total_margin ?? null;
    it[`margin_impact_${k}_vs_formulary`] = m?.margin_impact_form ?? null;
    it[`margin_impact_${k}_vs_po`] = m?.margin_impact_po ?? null;
  }
  const L = byStep.latest;
  it.latest_stage = L?.step ?? null;
  it.latest_price = L?.price ?? null;
  it.change_pct = L && acuan ? L.price / acuan - 1 : null;
  it.impact = L?.cost_impact ?? null;
  it.margin_pc = L?.margin_pc ?? null;
  it.margin_pct = L && baseMargin ? L.margin_pc / baseMargin : null;
  it.margin_annual = L?.total_margin ?? null;
  it.margin_vs_formulary = L?.margin_impact_form ?? null;
  it.margin_vs_po = L?.margin_impact_po ?? null;
  it.margin_impact_vs_formulary = it.margin_vs_formulary;
  it.margin_impact_vs_po = it.margin_vs_po;
  it.formulary_unit_price = it.form_nett_price;
  it.margin_pc_formulary = formMarginUnit();
  it.margin_pc_po = baseMargin != null && it.po_unit_price != null ? baseMargin - it.po_unit_price : null;

  // ---- A7 columns
  it.variance_hna_pc = it.rfq_hna != null && it.rfq_qty > 0 && formHnaPc != null ? it.rfq_hna / it.rfq_qty - formHnaPc : null;
  it.variance_disc = it.rfq_disc != null && it.form_disc != null ? it.rfq_disc - it.form_disc : null;
  void hnaPc;
  return it;
}

const sum = (xs: (number | null | undefined)[]) => {
  const v = xs.filter((x): x is number => x != null);
  return v.length ? v.reduce((a, b) => a + b, 0) : null;
};

export function summarizeCycleItems(enriched: any[]) {
  const active = enriched.filter((i) => i.item_status !== "Discontinue");
  const stages = (["rfq", "co", "fb1", "on"] as const).map((key) => {
    const rows = active.filter((i) => i[`${key}_unit_price`] != null);
    const cmp = (i: any) => i.primary_bench;
    return {
      key,
      label: { rfq: "RFQ", co: "Counter offer", fb1: "Feedback I", on: "Online nego" }[key],
      items: rows.length,
      increases: rows.filter((i) => cmp(i) && i[`${key}_unit_price`] > cmp(i) * 1.0005).length,
      decreases: rows.filter((i) => cmp(i) && i[`${key}_unit_price`] < cmp(i) * 0.9995).length,
      impact: sum(rows.map((i) => i[`impact_${key}`])),
      margin_impact_vs_formulary: sum(rows.map((i) => i[`margin_impact_${key}_vs_formulary`])),
      margin_impact_vs_po: sum(rows.map((i) => i[`margin_impact_${key}_vs_po`])),
    };
  });
  const rfqFilled = active.filter((i) => i.rfq_hna != null);
  const poValue = active.reduce((s, i) => s + (i.po_annualized_value_ppn || 0), 0);
  return {
    items: enriched.length,
    active: active.length,
    discontinued: enriched.length - active.length,
    with_mou: enriched.filter((i) => i.in_mou).length,
    with_po: enriched.filter((i) => i.po_qty_12m != null && i.po_qty_12m !== 0).length,
    po_value_12m: poValue,
    baseline_spend: poValue,
    stages,
    rfq_filled: rfqFilled.length,
    rfq_active_pct: active.length ? rfqFilled.length / active.length : 0,
    latest_impact: sum(active.map((i) => i.impact)),
    margin_vs_formulary: sum(active.map((i) => i.margin_vs_formulary)),
    margin_vs_po: sum(active.map((i) => i.margin_vs_po)),
    margin_impact_vs_formulary: sum(active.map((i) => i.margin_vs_formulary)),
    margin_impact_vs_po: sum(active.map((i) => i.margin_vs_po)),
  };
}

// ---- escalation (D2): over the % of the principal's annual PO value OR over the Rp limit ----
export function escalationFor(costImpactFin: number | null, poValueAnnualPpn: number, pctLimit: number, rpLimit: number) {
  if (costImpactFin == null || costImpactFin <= 0) return { needed: false, reasons: [] as string[], pct: costImpactFin != null && poValueAnnualPpn ? costImpactFin / poValueAnnualPpn : null };
  const pct = poValueAnnualPpn > 0 ? costImpactFin / poValueAnnualPpn : null;
  const reasons: string[] = [];
  if (pctLimit > 0 && pct != null && pct > pctLimit) reasons.push("pct");
  if (rpLimit > 0 && costImpactFin > rpLimit) reasons.push("rp");
  return { needed: reasons.length > 0, reasons, pct };
}

// Master data files (formulary, MOU) are looser than the principal form: "10%" is 10%, 0,1 is 10%, 20 is 20%.
export function parseDiscLenient(raw: any): number | null {
  if (raw == null || raw === "") return null;
  const s = String(raw).trim();
  if (!s) return null;
  const n = parseMoney(s);
  if (n == null || n < 0) return null;
  const frac = s.includes("%") ? n / 100 : n < 1 ? n : n / 100;
  return frac >= 1 ? null : Math.round(frac * 10000) / 10000;
}

// Principal names compare without case, dots and extra spaces ("PT. Dummy  Farma" = "pt dummy farma").
export const normName = (v: any): string => String(v ?? "").toUpperCase().replace(/[.,]/g, "").replace(/\s+/g, " ").trim();
