// Reads uploaded CSV / Excel files into plain rows, and looks columns up by loose header names,
// so exports from different systems ("Item Number", "Kode Barang", "ERP Code"...) all work.
import ExcelJS from "exceljs";

export type Row = Record<string, string>;

// RFC 4180 style: quoted fields may contain commas, quotes ("") and line breaks.
export function parseCsvText(text: string): Row[] {
  const src = text.replace(/^﻿/, "");
  const sep = detectSeparator(src);
  const records: string[][] = [];
  let field = "";
  let rec: string[] = [];
  let quoted = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') { field += '"'; i++; } else quoted = false;
      } else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === sep) { rec.push(field); field = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && src[i + 1] === "\n") i++;
      rec.push(field); field = "";
      if (rec.some((c) => c.trim() !== "")) records.push(rec);
      rec = [];
    } else field += ch;
  }
  rec.push(field);
  if (rec.some((c) => c.trim() !== "")) records.push(rec);
  return recordsToRows(records);
}

// Indonesian spreadsheets often export with ";" instead of ",".
function detectSeparator(src: string): string {
  const first = src.split(/\r?\n/, 1)[0] || "";
  const commas = (first.match(/,/g) || []).length;
  const semis = (first.match(/;/g) || []).length;
  return semis > commas ? ";" : ",";
}

function recordsToRows(records: string[][]): Row[] {
  if (!records.length) return [];
  const headers = records[0].map((h) => h.trim());
  return records.slice(1).map((r) => {
    const row: Row = {};
    headers.forEach((h, i) => { if (h) row[h] = (r[i] ?? "").trim(); });
    return row;
  });
}

function cellText(v: ExcelJS.CellValue): string {
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "object") {
    const o = v as any;
    if (o.result !== undefined) return cellText(o.result);   // formula
    if (o.richText) return o.richText.map((t: any) => t.text).join("");
    if (o.text !== undefined) return String(o.text);          // hyperlink
    if (o.error) return "";
  }
  return String(v).trim();
}

export async function readTable(file: { originalname: string; buffer: Buffer }): Promise<Row[]> {
  const name = file.originalname.toLowerCase();
  if (name.endsWith(".xlsx")) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(file.buffer as any);
    const ws = wb.worksheets.find((w) => w.rowCount > 0);
    if (!ws) return [];
    const records: string[][] = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
      const cells: string[] = [];
      for (let c = 1; c <= row.cellCount; c++) cells.push(cellText(row.getCell(c).value));
      if (cells.some((c) => c !== "")) records.push(cells);
    });
    return recordsToRows(records);
  }
  if (name.endsWith(".csv") || name.endsWith(".txt")) return parseCsvText(file.buffer.toString("utf-8"));
  throw new Error("Unsupported file type. Upload .csv or .xlsx");
}

const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");

// First column whose header matches any alias (ignoring case, spaces and punctuation).
export function pick(row: Row, ...aliases: string[]): string {
  const keys = Object.keys(row);
  for (const a of aliases) {
    const na = norm(a);
    const k = keys.find((x) => norm(x) === na);
    if (k !== undefined && row[k] !== "") return row[k];
  }
  return "";
}

export function hasColumn(rows: Row[], ...aliases: string[]): boolean {
  if (!rows.length) return false;
  const keys = Object.keys(rows[0]).map(norm);
  return aliases.some((a) => keys.includes(norm(a)));
}

// "Rp 1.234.567,50", "1,234.50", "1234.5", "15%" -> number (percent signs are dropped, not divided).
export function toNumber(v: string | undefined | null): number | null {
  if (v == null) return null;
  let s = String(v).trim().replace(/rp/i, "").replace(/%/g, "").replace(/\s/g, "");
  if (!s) return null;
  const lastDot = s.lastIndexOf(".");
  const lastComma = s.lastIndexOf(",");
  if (lastDot >= 0 && lastComma >= 0) {
    // the later separator is the decimal mark
    s = lastComma > lastDot ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  } else if (lastComma >= 0) {
    s = /^\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, "") : s.replace(",", ".");
  } else if (lastDot >= 0 && /^\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, "");
  }
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

// "2027-01-31", "31/01/2027", "31-Jan-2027" -> "2027-01-31"; null when not a date.
export function toIsoDate(v: string | undefined | null): string | null {
  if (!v) return null;
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
  const d = new Date(s);
  return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
}
