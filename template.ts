// Template_Nego: the workbook the principal knows. It is generated here (no shipped file), so every hint
// names the right column letter, validation runs to row 1000 and Excel opens it without a repair message.
// Staff version: with the current MOU and every step. Vendor version: without any Siloam reference values.
import ExcelJS from "exceljs";
import { parseDiscount, discountFromCell, parseMoney } from "./pricing.js";

export const FIRST_ROW = 8;
export const LAST_ROW = 1000;
const HEADER_ROW = 7;

export interface TemplateOptions {
  forPrincipal: boolean;       // vendor copy: no MOU values, counter offer only once sent
  showCounterOffer: boolean;
  principal: string;
  distributor: string;
  contractStart: string;
  contractEnd: string;
  binding: string;
  deliveryFee: string;
  category: string;            // Drugs: Active Ingredient instead of Catalog No. (REF)
  ppn: number;
}

export const REASONS: { key: string; label: string }[] = [
  { key: "raw_material", label: "Kenaikan harga bahan baku" },
  { key: "fx", label: "Perubahan nilai tukar (kurs)" },
  { key: "production", label: "Kenaikan biaya produksi / logistik" },
  { key: "other", label: "Lainnya" },
];

const fmtLongDate = (iso: string) => {
  const m = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const [y, mo, d] = String(iso || "").split("-").map(Number);
  return y ? `${String(d).padStart(2, "0")} ${m[mo - 1]} ${y}` : "-";
};

interface ColSpec { key: string; header: string; width: number; kind: "text" | "qty" | "money" | "disc" | "price" | "status" | "reason"; staffOnly?: boolean; counter?: boolean; input?: boolean; section: string }

function columnsFor(o: TemplateOptions): ColSpec[] {
  const drugs = o.category === "Drugs";
  const cols: ColSpec[] = [
    { key: "erp", header: "ERP Code", width: 14, kind: "text", section: "Item" },
    { key: "name", header: "Item Name", width: 42, kind: "text", section: "Item" },
    { key: "brand", header: "Brand / Manufacturer", width: 20, kind: "text", input: true, section: "Item" },
    drugs
      ? { key: "ai", header: "Active Ingredient", width: 22, kind: "text", input: true, section: "Item" }
      : { key: "ref", header: "Catalog No. (REF)", width: 18, kind: "text", input: true, section: "Item" },
    { key: "status", header: "Item Status", width: 14, kind: "status", input: true, section: "Item" },
    { key: "remarks", header: "Remarks", width: 24, kind: "text", input: true, section: "Item" },
    { key: "mouQty", header: "MOU_Qty/PO Unit", width: 12, kind: "qty", staffOnly: true, section: "Current MOU (reference)" },
    { key: "mouHna", header: "MOU_HNA/PO Unit (excl. PPN)", width: 18, kind: "money", staffOnly: true, section: "Current MOU (reference)" },
    { key: "mouDisc", header: "MOU_Disc%", width: 11, kind: "disc", staffOnly: true, section: "Current MOU (reference)" },
    { key: "mouPrice", header: "MOU_Unit Price/pcs (incl. PPN)", width: 18, kind: "price", staffOnly: true, section: "Current MOU (reference)" },
    { key: "rfqQty", header: "RFQ_Qty/PO Unit", width: 12, kind: "qty", input: true, section: "RFQ" },
    { key: "rfqHna", header: "RFQ_HNA/PO Unit (excl. PPN)", width: 18, kind: "money", input: true, section: "RFQ" },
    { key: "rfqDisc", header: "RFQ_Disc%", width: 11, kind: "disc", input: true, section: "RFQ" },
    { key: "rfqPrice", header: "RFQ_Unit Price/PCS (incl. PPN)", width: 18, kind: "price", section: "RFQ" },
    { key: "reason", header: "Alasan Kenaikan Harga", width: 30, kind: "reason", input: true, section: "RFQ" },
    { key: "coDisc", header: "CO_Disc%", width: 11, kind: "disc", counter: true, section: "Counter offer" },
    { key: "coPrice", header: "CO_Unit Price/PCS (incl. PPN)", width: 18, kind: "price", counter: true, section: "Counter offer" },
    { key: "coNote", header: "CO_Note (Siloam)", width: 28, kind: "text", staffOnly: true, section: "Counter offer" },
    { key: "fb1Disc", header: "FB1_Disc%", width: 11, kind: "disc", counter: true, input: true, section: "Feedback I" },
    { key: "fb1Price", header: "FB1_Unit Price/PCS (incl. PPN)", width: 18, kind: "price", counter: true, section: "Feedback I" },
    { key: "onDisc", header: "ON_Disc%", width: 11, kind: "disc", staffOnly: true, section: "Online Nego" },
    { key: "onPrice", header: "ON_Unit Price/PCS (incl. PPN)", width: 18, kind: "price", staffOnly: true, section: "Online Nego" },
    { key: "onNett", header: "ON_Disc% (skema Nett)", width: 14, kind: "disc", staffOnly: true, section: "Online Nego" },
    { key: "onDiscScheme", header: "ON_Disc% (skema Disc)", width: 14, kind: "disc", staffOnly: true, section: "Online Nego" },
  ];
  return cols.filter((c) => (o.forPrincipal ? !c.staffOnly : true) && (!c.counter || !o.forPrincipal || o.showCounterOffer));
}

export async function buildTemplate(_unused: string | null, items: any[], o: TemplateOptions): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Siloam Negotiation";
  const ws = wb.addWorksheet("Form");
  const cols = columnsFor(o);
  const L: Record<string, string> = {};
  cols.forEach((c, i) => { L[c.key] = ws.getColumn(i + 1).letter; });
  const at = (k: string) => cols.findIndex((c) => c.key === k) + 1;

  const bindingText = o.binding === "Disc" ? "Discount (% off HNA)" : o.binding === "Nett to Disc" ? "Nett to Disc (formulary HNA × (1 − disc))" : "Nett Price (HNA after discount)";
  ws.getCell("A1").value = `Contract Period: ${fmtLongDate(o.contractStart)}  -  ${fmtLongDate(o.contractEnd)}`;
  ws.getCell("A2").value = `Binding: ${bindingText}`;
  ws.getCell("A3").value = `Delivery Fee: ${o.deliveryFee}`;
  ws.getCell("A4").value = `Principal: ${o.principal}`;
  ws.getCell("A5").value = `Distributor: ${o.distributor || "-"}`;
  for (let r = 1; r <= 5; r++) ws.getCell(r, 1).font = { bold: r === 4 };

  // Hints that name the right columns (taken from the column map, never typed by hand).
  const hint = (k: string, text: string) => { if (at(k)) { const c = ws.getCell(6, at(k)); c.value = text; c.font = { italic: true, size: 8, color: { argb: "FF555555" } }; c.alignment = { wrapText: true, vertical: "top" }; } };
  hint("status", "Active atau Discontinue. Discontinue wajib dengan surat DC (lampirkan di portal).");
  hint("rfqQty", `Jumlah pcs per kemasan. Basis HNA kolom ${L.rfqHna}.`);
  hint("rfqHna", `HNA per satuan pembelian (kolom ${L.rfqQty}), sebelum PPN.`);
  hint("rfqDisc", "Angka 0 sampai 100 (15 = 15%). Maksimal 2 desimal.");
  hint("reason", "Wajib bila harga per pcs naik dari MOU. Pilih dari daftar.");
  if (at("fb1Disc")) hint("fb1Disc", "Angka 0 sampai 100 (15 = 15%).");
  ws.getRow(6).height = 44;

  // header rows
  let prev = ""; let start = 1;
  const sectionRow = 5;  // section titles share row 5 with column A5 only for column >= first section; keep them on row 6 hints? use row 5 right of A
  void sectionRow; void prev; void start;
  cols.forEach((c, i) => {
    const cell = ws.getCell(HEADER_ROW, i + 1);
    cell.value = c.header;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.alignment = { wrapText: true, vertical: "middle" };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: c.input ? "FF1F6F5C" : c.staffOnly ? "FF5B6770" : "FF3A4A5A" } };
    ws.getColumn(i + 1).width = c.width;
  });
  ws.getRow(HEADER_ROW).height = 34;

  const vat = 1 + o.ppn;
  const put = (r: number, k: string, v: any) => { if (at(k)) ws.getCell(r, at(k)).value = v == null || v === "" ? null : v; };
  const noPriceFormula = o.binding === "Nett to Disc";
  const priceFormula = (r: number, discKey: string, qtyKey: string, hnaKey: string) =>
    ({ formula: `IFERROR(IF(${L[hnaKey]}${r}="","",${L[hnaKey]}${r}/${L[qtyKey]}${r}*(1-N(${L[discKey]}${r})/100)*${vat}),"")` });

  for (let r = FIRST_ROW; r <= LAST_ROW; r++) {
    const it = items[r - FIRST_ROW];
    if (it) {
      put(r, "erp", it.erp_code); put(r, "name", it.item_name); put(r, "brand", it.brand);
      put(r, "ref", it.catalog_no); put(r, "ai", it.active_ingredient); put(r, "status", it.item_status); put(r, "remarks", it.remarks);
      put(r, "rfqQty", it.rfq_qty); put(r, "rfqHna", it.rfq_hna);
      // Discounts are written as numbers 0-100 (the cell shows the % sign), so 15 can never be read as 1,500%.
      const pctOf = (v: any) => (v == null ? null : Math.round(v * 10000) / 100);
      put(r, "rfqDisc", pctOf(it.rfq_disc));
      put(r, "reason", it.price_reason ? (REASONS.find((x) => x.key === it.price_reason)?.label || it.price_reason) : null);
      if (!o.forPrincipal) {
        put(r, "mouQty", it.mou_qty); put(r, "mouHna", it.mou_hna); put(r, "mouDisc", pctOf(it.mou_disc));
        put(r, "coNote", it.co_note); put(r, "onDisc", pctOf(it.on_disc)); put(r, "onNett", pctOf(it.on_disc_nett)); put(r, "onDiscScheme", pctOf(it.on_disc_disc));
      }
      if (!o.forPrincipal || o.showCounterOffer) put(r, "coDisc", pctOf(it.co_disc));
      put(r, "fb1Disc", pctOf(it.fb1_disc));
    }
    if (!noPriceFormula) {
      if (!o.forPrincipal) put(r, "mouPrice", priceFormula(r, "mouDisc", "mouQty", "mouHna"));
      put(r, "rfqPrice", priceFormula(r, "rfqDisc", "rfqQty", "rfqHna"));
      const hnaK = "rfqHna", qtyK = "rfqQty";
      const abs = (c: string) => `$${L[c]}${r}`;
      const af = (d: string) => ({ formula: `IFERROR(IF(OR(${abs(hnaK)}="",${L[d]}${r}=""),"",${abs(hnaK)}/${abs(qtyK)}*(1-${L[d]}${r}/100)*${vat}),"")` });
      put(r, "coPrice", af("coDisc")); put(r, "fb1Price", af("fb1Disc"));
      if (!o.forPrincipal) put(r, "onPrice", af("onDisc"));
    } else if (it && !o.forPrincipal) {
      put(r, "rfqPrice", it.rfq_unit_price); put(r, "coPrice", it.co_unit_price); put(r, "fb1Price", it.fb1_unit_price); put(r, "onPrice", it.on_unit_price); put(r, "mouPrice", it.mou_unit_price);
    }
  }

  // Formats and validation to row 1000, so a 200-item template is fully covered.
  const rng = (k: string) => `${L[k]}${FIRST_ROW}:${L[k]}${LAST_ROW}`;
  for (const c of cols) {
    const colIdx = at(c.key);
    for (let r = FIRST_ROW; r <= LAST_ROW; r++) {
      const cell = ws.getCell(r, colIdx);
      if (c.kind === "disc") cell.numFmt = '0.##"%"';
      else if (c.kind === "money") cell.numFmt = "#,##0.##";
      else if (c.kind === "price") cell.numFmt = "#,##0.00";
      else if (c.kind === "qty") cell.numFmt = "0";
      if (c.input) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4FBF8" } };
      if (c.kind === "price") cell.font = { color: { argb: "FF777777" } };
    }
    if (c.kind === "price" || !c.input && !c.staffOnly) continue;
  }
  const dv = (k: string, def: any) => {
    if (!at(k)) return;
    for (let r = FIRST_ROW; r <= LAST_ROW; r++) ws.getCell(r, at(k)).dataValidation = def;
  };
  dv("status", { type: "list", allowBlank: true, formulae: ['"Active,Discontinue"'], showErrorMessage: true, errorTitle: "Item Status", error: "Pilih Active atau Discontinue." });
  dv("rfqQty", { type: "whole", operator: "greaterThanOrEqual", allowBlank: true, formulae: [1], showErrorMessage: true, errorTitle: "Qty/PO Unit", error: "Isi bilangan bulat minimal 1." });
  dv("rfqHna", { type: "decimal", operator: "greaterThan", allowBlank: true, formulae: [0], showErrorMessage: true, errorTitle: "HNA", error: "HNA harus lebih dari 0." });
  for (const k of ["rfqDisc", "fb1Disc"]) dv(k, { type: "decimal", operator: "between", allowBlank: true, formulae: [0, 99.99], showErrorMessage: true, errorTitle: "Diskon", error: "Isi angka 0 sampai 99,99 (15 = 15%)." });
  dv("reason", { type: "list", allowBlank: true, formulae: [`"${REASONS.map((x) => x.label).join(",")}"`], showErrorMessage: true, errorTitle: "Alasan", error: "Pilih salah satu alasan dari daftar." });
  void rng;

  ws.views = [{ state: "frozen", xSplit: 2, ySplit: HEADER_ROW, topLeftCell: `C${FIRST_ROW}`, activeCell: `C${FIRST_ROW}` }];
  ws.autoFilter = { from: { row: HEADER_ROW, column: 1 }, to: { row: HEADER_ROW, column: cols.length } };
  if (o.forPrincipal) ws.getCell("A6").value = "Isi kolom hijau. Kolom abu-abu dihitung otomatis.";
  ws.getCell("A6").font = { italic: true, size: 9 };
  return wb;
}

export interface TemplateRow { row: number; erp_code: string; item_name: string; cells: Record<string, string> }

const percentScaled = (fmt?: string) => !!fmt && /%/.test(fmt.replace(/"[^"]*"/g, "").replace(/\\./g, ""));

function text(v: ExcelJS.CellValue, fmt?: string, isDisc = false): string {
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") return isDisc && percentScaled(fmt) ? `${Math.round(v * 10000) / 100}%` : String(v);
  if (typeof v === "object") {
    const o = v as any;
    if (o.result !== undefined) return text(o.result, fmt, isDisc);
    if (o.richText) return o.richText.map((t: any) => t.text).join("");
    if (o.text !== undefined) return String(o.text);
    return "";
  }
  return String(v).trim();
}

// Reads a filled template. Headers are found by name (the row that contains "ERP Code"), so a few
// inserted title rows or reordered columns still work. A discount cell formatted as a percent (0.275 shown as
// 27.5%) is read as 27.5%, a plain number is already a percent (15 = 15%).
export async function readTemplate(buffer: Buffer): Promise<TemplateRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.getWorksheet("Form") || wb.worksheets[0];
  if (!ws) throw new Error("The workbook has no sheet");
  let headerRow = 0;
  const headers: Record<string, number> = {};
  for (let r = 1; r <= 20 && !headerRow; r++) {
    ws.getRow(r).eachCell((c) => { if (text(c.value).toLowerCase() === "erp code") headerRow = r; });
  }
  if (!headerRow) throw new Error('This is not Template_Nego: no "ERP Code" column found');
  ws.getRow(headerRow).eachCell((c, n) => { const t = text(c.value); if (t) headers[t.toLowerCase()] = n; });
  const at = (names: string[]) => { for (const n of names) if (headers[n.toLowerCase()]) return headers[n.toLowerCase()]; return 0; };
  const cols: Record<string, number> = {
    erp: at(["ERP Code"]), name: at(["Item Name"]), brand: at(["Brand / Manufacturer", "Brand"]), ref: at(["Catalog No. (REF)", "Catalog No"]),
    ai: at(["Active Ingredient"]), status: at(["Item Status", "Status"]), remarks: at(["Remarks"]),
    rfqQty: at(["RFQ_Qty/PO Unit"]), rfqHna: at(["RFQ_HNA/PO Unit (excl. PPN)"]), rfqDisc: at(["RFQ_Disc%"]), reason: at(["Alasan Kenaikan Harga"]),
    coDisc: at(["CO_Disc%"]), coNote: at(["CO_Note (Siloam)", "CO_Note"]), fb1Disc: at(["FB1_Disc%"]), onDisc: at(["ON_Disc%"]),
    onNett: at(["ON_Disc% (skema Nett)"]), onDiscScheme: at(["ON_Disc% (skema Disc)"]),
  };
  const discKeys = new Set(["rfqDisc", "coDisc", "fb1Disc", "onDisc", "onNett", "onDiscScheme"]);
  const rows: TemplateRow[] = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const get = (k: string) => { if (!cols[k]) return ""; const c = row.getCell(cols[k]); return text(c.value, c.numFmt, discKeys.has(k)); };
    const erp = get("erp");
    if (!erp && !get("name")) continue;
    const cells: Record<string, string> = {};
    for (const k of Object.keys(cols)) if (k !== "erp" && k !== "name") cells[k] = get(k);
    rows.push({ row: r, erp_code: erp.trim().toUpperCase(), item_name: get("name"), cells });
  }
  return rows;
}

const STATUS_WORDS: Record<string, string> = { active: "Active", aktif: "Active", discontinue: "Discontinue", discontinued: "Discontinue", "tidak dijual lagi": "Discontinue", inactive: "Discontinue" };

export interface CellChange { field: string; header: string; value: any }
export interface RowCheck { changes: CellChange[]; errors: { column: string; value: string; message: string }[] }

const FIELD_META: Record<string, { header: string; field: string; kind: "text" | "status" | "qty" | "money" | "disc" | "reason" }> = {
  brand: { header: "Brand / Manufacturer", field: "brand", kind: "text" },
  ref: { header: "Catalog No. (REF)", field: "catalog_no", kind: "text" },
  ai: { header: "Active Ingredient", field: "active_ingredient", kind: "text" },
  status: { header: "Item Status", field: "item_status", kind: "status" },
  remarks: { header: "Remarks", field: "remarks", kind: "text" },
  rfqQty: { header: "RFQ_Qty/PO Unit", field: "rfq_qty", kind: "qty" },
  rfqHna: { header: "RFQ_HNA/PO Unit", field: "rfq_hna", kind: "money" },
  rfqDisc: { header: "RFQ_Disc%", field: "rfq_disc", kind: "disc" },
  reason: { header: "Alasan Kenaikan Harga", field: "price_reason", kind: "reason" },
  coDisc: { header: "CO_Disc%", field: "co_disc", kind: "disc" },
  coNote: { header: "CO_Note", field: "co_note", kind: "text" },
  fb1Disc: { header: "FB1_Disc%", field: "fb1_disc", kind: "disc" },
  onDisc: { header: "ON_Disc%", field: "on_disc", kind: "disc" },
  onNett: { header: "ON_Disc% (skema Nett)", field: "on_disc_nett", kind: "disc" },
  onDiscScheme: { header: "ON_Disc% (skema Disc)", field: "on_disc_disc", kind: "disc" },
};

// Turns the cells of one row into validated changes. Blank cells never erase a value. A wrong cell is
// reported with its column and value; the other cells of the row still count.
// `allowed` limits which columns this reader may change (a vendor may not touch the counter offer).
export function checkRow(cells: Record<string, string>, allowed: string[]): RowCheck {
  const changes: CellChange[] = [];
  const errors: RowCheck["errors"] = [];
  for (const key of allowed) {
    const raw = (cells[key] ?? "").trim();
    if (raw === "") continue;
    const m = FIELD_META[key];
    if (!m) continue;
    const bad = (message: string) => errors.push({ column: m.header, value: raw, message });
    if (m.kind === "text") { changes.push({ field: m.field, header: m.header, value: raw.slice(0, 300) }); continue; }
    if (m.kind === "status") {
      const v = STATUS_WORDS[raw.toLowerCase()];
      if (!v) bad("pakai Active atau Discontinue"); else changes.push({ field: m.field, header: m.header, value: v });
      continue;
    }
    if (m.kind === "reason") {
      const hit = REASONS.find((x) => x.label.toLowerCase() === raw.toLowerCase() || x.key === raw.toLowerCase());
      changes.push({ field: m.field, header: m.header, value: hit ? hit.key : raw.slice(0, 300) });
      continue;
    }
    if (m.kind === "disc") {
      const d = parseDiscount(raw);
      if (d.error) bad(d.error); else if (d.value != null) changes.push({ field: m.field, header: m.header, value: d.value });
      continue;
    }
    const n = parseMoney(raw);
    if (n == null) { bad("bukan angka"); continue; }
    if (m.kind === "qty") {
      if (n < 1 || !Number.isFinite(n)) bad("minimal 1"); else changes.push({ field: m.field, header: m.header, value: Math.round(n) });
    } else if (n <= 0) bad("harus lebih dari 0");
    else changes.push({ field: m.field, header: m.header, value: n });
  }
  return { changes, errors };
}

// kept for callers that read Excel percent cells directly
export { discountFromCell };
