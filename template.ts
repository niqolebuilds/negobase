// Template_Nego: the workbook the principal knows. Builds it for a negotiation (staff version with the
// current MOU, vendor version without any Siloam reference values) and reads a filled one back.
import ExcelJS from "exceljs";
import { toNumber } from "./tabular.js";

// Column positions in the "Form" sheet; data starts at row 8, headers are on row 7.
export const COL = {
  erp: 1, name: 2, brand: 3, ref: 4, status: 5, remarks: 6,
  mouQty: 7, mouHna: 8, mouDisc: 9, mouPrice: 10,
  rfqQty: 11, rfqHna: 12, rfqDisc: 13, rfqPrice: 14,
  coDisc: 15, coPrice: 16, fb1Disc: 17, fb1Price: 18, onDisc: 19, onPrice: 20,
} as const;
const FIRST_ROW = 8;
const FORMULA_ROWS = 500; // the shipped template carries formulas to row 500

export interface TemplateOptions {
  forPrincipal: boolean;       // vendor copy: no MOU values, counter offer only once sent
  showCounterOffer: boolean;
  principal: string;
  distributor: string;
  contractStart: string;
  contractEnd: string;
  binding: string;
  deliveryFee: string;
}

const fmtLongDate = (iso: string) => {
  const m = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const [y, mo, d] = iso.split("-").map(Number);
  return `${String(d).padStart(2, "0")} ${m[mo - 1]} ${y}`;
};

export async function buildTemplate(templatePath: string, items: any[], o: TemplateOptions): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(templatePath);
  const ws = wb.getWorksheet("Form") || wb.worksheets[0];
  ws.getCell("A1").value = `Contract Period: ${fmtLongDate(o.contractStart)}  -  ${fmtLongDate(o.contractEnd)}`;
  const bindingText = o.binding === "Disc" ? "Discount (% off HNA)" : o.binding === "Nett to Disc" ? "Nett to Disc (formulary HNA × (1 − disc))" : "Nett Price (HNA after discount)";
  ws.getCell("A2").value = `Binding: ${bindingText}`;
  ws.getCell("A3").value = `Delivery Fee: ${o.deliveryFee}`;
  ws.getCell("A4").value = `Principal: ${o.principal}`;
  ws.getCell("A5").value = `Distributor: ${o.distributor || "-"}`;

  const put = (r: number, c: number, v: any) => { ws.getCell(r, c).value = v == null || v === "" ? null : v; };
  items.forEach((it, i) => {
    const r = FIRST_ROW + i;
    put(r, COL.erp, it.erp_code);
    put(r, COL.name, it.item_name);
    put(r, COL.brand, it.brand);
    put(r, COL.ref, it.catalog_no);
    put(r, COL.status, it.item_status);
    put(r, COL.remarks, it.remarks);
    if (!o.forPrincipal) {
      put(r, COL.mouQty, it.mou_qty);
      put(r, COL.mouHna, it.mou_hna);
      put(r, COL.mouDisc, it.mou_disc);
    }
    put(r, COL.rfqQty, it.rfq_qty);
    put(r, COL.rfqHna, it.rfq_hna);
    put(r, COL.rfqDisc, it.rfq_disc);
    if (!o.forPrincipal || o.showCounterOffer) put(r, COL.coDisc, it.co_disc);
    put(r, COL.fb1Disc, it.fb1_disc);
    if (!o.forPrincipal) put(r, COL.onDisc, it.on_disc);
    // formulas past the shipped 500 rows
    if (r > FORMULA_ROWS) {
      const L = (c: number) => ws.getColumn(c).letter;
      const price = (q: number, h: number, d: number, abs = false) => {
        const qa = abs ? `$${L(h)}${r}` : `${L(h)}${r}`, qq = abs ? `$${L(q)}${r}` : `${L(q)}${r}`;
        return { formula: `IFERROR(IF(${qa}="","",${qa}/${qq}*(1-${L(d)}${r})*1.11),"")` };
      };
      if (!o.forPrincipal) ws.getCell(r, COL.mouPrice).value = price(COL.mouQty, COL.mouHna, COL.mouDisc);
      ws.getCell(r, COL.rfqPrice).value = price(COL.rfqQty, COL.rfqHna, COL.rfqDisc);
      ws.getCell(r, COL.coPrice).value = price(COL.rfqQty, COL.rfqHna, COL.coDisc, true);
      ws.getCell(r, COL.fb1Price).value = price(COL.rfqQty, COL.rfqHna, COL.fb1Disc, true);
      ws.getCell(r, COL.onPrice).value = price(COL.rfqQty, COL.rfqHna, COL.onDisc, true);
    }
  });
  if (o.forPrincipal) {
    // The vendor copy never carries current MOU, counter-offer analysis or online-nego values.
    ws.getCell("G6").value = "";
    for (const c of [COL.mouQty, COL.mouHna, COL.mouDisc, COL.mouPrice]) { ws.getColumn(c).hidden = true; }
    if (!o.showCounterOffer) for (const c of [COL.coDisc, COL.coPrice]) ws.getColumn(c).hidden = true;
    for (const c of [COL.onDisc, COL.onPrice]) ws.getColumn(c).hidden = true;
  }
  return wb;
}

export interface TemplateRow { row: number; erp_code: string; item_name: string; cells: Record<string, string> }

function text(v: ExcelJS.CellValue): string {
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    const o = v as any;
    if (o.result !== undefined) return text(o.result);
    if (o.richText) return o.richText.map((t: any) => t.text).join("");
    if (o.text !== undefined) return String(o.text);
    return "";
  }
  return String(v).trim();
}

// Reads a filled template. Headers are found by name (the row that contains "ERP Code"), so a few
// inserted title rows or reordered columns still work.
export async function readTemplate(buffer: Buffer): Promise<TemplateRow[]> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer as any);
  const ws = wb.getWorksheet("Form") || wb.worksheets[0];
  if (!ws) throw new Error("The workbook has no sheet");
  let headerRow = 0;
  const headers: Record<string, number> = {};
  for (let r = 1; r <= 20 && !headerRow; r++) {
    ws.getRow(r).eachCell((c, n) => { if (text(c.value).toLowerCase() === "erp code") headerRow = r; });
  }
  if (!headerRow) throw new Error('This is not Template_Nego: no "ERP Code" column found');
  ws.getRow(headerRow).eachCell((c, n) => { const t = text(c.value); if (t) headers[t.toLowerCase()] = n; });
  const at = (names: string[]) => { for (const n of names) if (headers[n.toLowerCase()]) return headers[n.toLowerCase()]; return 0; };
  const cols: Record<string, number> = {
    erp: at(["ERP Code"]), name: at(["Item Name"]), brand: at(["Brand / Manufacturer", "Brand"]), ref: at(["Catalog No. (REF)", "Catalog No"]),
    status: at(["Item Status", "Status"]), remarks: at(["Remarks"]),
    rfqQty: at(["RFQ_Qty/PO Unit"]), rfqHna: at(["RFQ_HNA/PO Unit (excl. PPN)"]), rfqDisc: at(["RFQ_Disc%"]),
    coDisc: at(["CO_Disc%"]), fb1Disc: at(["FB1_Disc%"]), onDisc: at(["ON_Disc%"]),
  };
  const rows: TemplateRow[] = [];
  for (let r = headerRow + 1; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const get = (k: string) => (cols[k] ? text(row.getCell(cols[k]).value) : "");
    const erp = get("erp");
    if (!erp && !get("name")) continue;
    const cells: Record<string, string> = {};
    for (const k of Object.keys(cols)) if (k !== "erp" && k !== "name") cells[k] = get(k);
    rows.push({ row: r, erp_code: erp.toUpperCase(), item_name: get("name"), cells });
  }
  return rows;
}

const STATUS_WORDS: Record<string, string> = { active: "Active", aktif: "Active", discontinue: "Discontinue", discontinued: "Discontinue", "tidak dijual lagi": "Discontinue", inactive: "Discontinue" };

export interface CellChange { field: string; header: string; value: any }
export interface RowCheck { changes: CellChange[]; errors: string[] }

const FIELD_META: Record<string, { header: string; field: string; kind: "text" | "status" | "qty" | "money" | "disc" }> = {
  brand: { header: "Brand / Manufacturer", field: "brand", kind: "text" },
  ref: { header: "Catalog No. (REF)", field: "catalog_no", kind: "text" },
  status: { header: "Item Status", field: "item_status", kind: "status" },
  remarks: { header: "Remarks", field: "remarks", kind: "text" },
  rfqQty: { header: "RFQ_Qty/PO Unit", field: "rfq_qty", kind: "qty" },
  rfqHna: { header: "RFQ_HNA/PO Unit", field: "rfq_hna", kind: "money" },
  rfqDisc: { header: "RFQ_Disc%", field: "rfq_disc", kind: "disc" },
  coDisc: { header: "CO_Disc%", field: "co_disc", kind: "disc" },
  fb1Disc: { header: "FB1_Disc%", field: "fb1_disc", kind: "disc" },
  onDisc: { header: "ON_Disc%", field: "on_disc", kind: "disc" },
};

// Turns the cells of one row into validated changes. Blank cells never erase a value.
// `allowed` limits which columns this reader may change (a vendor may not touch the counter offer).
export function checkRow(cells: Record<string, string>, allowed: string[]): RowCheck {
  const changes: CellChange[] = [];
  const errors: string[] = [];
  for (const key of allowed) {
    const raw = (cells[key] ?? "").trim();
    if (raw === "") continue;
    const m = FIELD_META[key];
    if (!m) continue;
    if (m.kind === "text") { changes.push({ field: m.field, header: m.header, value: raw.slice(0, 300) }); continue; }
    if (m.kind === "status") {
      const v = STATUS_WORDS[raw.toLowerCase()];
      if (!v) errors.push(`${m.header}: use Active or Discontinue`); else changes.push({ field: m.field, header: m.header, value: v });
      continue;
    }
    const n = toNumber(raw);
    if (n == null) { errors.push(`${m.header}: "${raw}" is not a number`); continue; }
    if (m.kind === "qty") {
      if (n < 1 || !Number.isFinite(n)) errors.push(`${m.header}: must be at least 1`); else changes.push({ field: m.field, header: m.header, value: Math.round(n) });
    } else if (m.kind === "money") {
      if (n <= 0) errors.push(`${m.header}: must be more than 0`); else changes.push({ field: m.field, header: m.header, value: n });
    } else {
      // A discount is 0 to 1 (a percent-formatted cell) or written with a % sign; a bare 15 is refused.
      const frac = raw.includes("%") ? n / 100 : n;
      if (frac < 0 || frac >= 1) errors.push(`${m.header}: a discount must be between 0% and 100% (a cell showing 15 means 1500%)`);
      else changes.push({ field: m.field, header: m.header, value: frac });
    }
  }
  return { changes, errors };
}
