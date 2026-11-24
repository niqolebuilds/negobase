// Excel Confirmation, draft BAK (Berita Acara Kesepakatan) and draft PKS as Word files.
// If Siloam uploaded its own .docx template for the binding, its {{placeholders}} are filled; otherwise a built-in
// draft is generated. Both carry the same agreed prices as the Excel Confirmation.
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, HeadingLevel, BorderStyle } from "docx";

export interface ConfRow { no: number; name: string; code: string; uom: number; base_unit: string; hna: number; disc: number; unit_price: number; unit_price_vat: number }
export interface DocMeta {
  principal: string; distributor: string; category: string; binding: string; scheme: "Nett" | "Disc";
  contract_start: string; contract_end: string; delivery_fee: string; ppn: number;
  meeting_at?: string | null; attendees_siloam?: string; attendees_principal?: string; notes?: string;
  followups?: { text: string; due?: string | null; owner?: string }[];
  sku_count: number; po_value_year?: number | null; cost_impact?: number | null;
}

const r2 = (n: number) => Math.round(n * 100) / 100;
export const fmtIdr = (n: number | null | undefined, d = 2) => (n == null ? "-" : n.toLocaleString("id-ID", { minimumFractionDigits: d, maximumFractionDigits: d }));
export const fmtPctText = (frac: number) => `${(Math.round(frac * 10000) / 100).toLocaleString("id-ID", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
const fmtLong = (iso?: string | null) => {
  if (!iso) return "-";
  const m = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const [y, mo, d] = iso.slice(0, 10).split("-").map(Number);
  return y ? `${d} ${m[mo - 1]} ${y}` : "-";
};

// Unit Price = HNA / UOM x (1 - Disc). Unit Price (+VAT) = Unit Price x (1 + PPN). Rounded to 2 decimals only for display.
export function confRow(no: number, name: string, code: string, uom: number, baseUnit: string, hna: number, disc: number, ppn: number): ConfRow {
  const up = (hna / uom) * (1 - disc);
  return { no, name, code, uom, base_unit: baseUnit, hna, disc, unit_price: r2(up), unit_price_vat: r2(up * (1 + ppn)) };
}

export async function buildConfirmationWorkbook(meta: DocMeta, rows: ConfRow[]): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Confirmation");
  ws.addRow([`Agreed prices: ${meta.principal}`]).font = { bold: true, size: 13 };
  ws.addRow([`Contract ${meta.contract_start} to ${meta.contract_end} · Binding ${meta.binding} · Scheme ${meta.scheme}`]);
  ws.addRow([]);
  const head = ws.addRow(["No", "Item Name", "ERP Code", "UOM", "Base Unit", "HNA", "Disc", "Unit Price", "Unit Price (+VAT)"]);
  head.font = { bold: true, color: { argb: "FFFFFFFF" } };
  head.eachCell((c) => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3A4A5A" } }; });
  for (const r of rows) ws.addRow([r.no, r.name, r.code, r.uom, r.base_unit, r2(r.hna), r.disc, r.unit_price, r.unit_price_vat]);
  const first = 5;
  for (let i = first; i < first + rows.length; i++) {
    ws.getCell(i, 6).numFmt = "#,##0.00";
    ws.getCell(i, 7).numFmt = "0.00%";
    ws.getCell(i, 8).numFmt = "#,##0.00";
    ws.getCell(i, 9).numFmt = "#,##0.00";
  }
  [6, 44, 14, 8, 10, 16, 10, 16, 18].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  ws.views = [{ state: "frozen", ySplit: 4 }];
  return wb;
}

// ---- Word ----------------------------------------------------------------------------------
const cell = (text: string, bold = false, align: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT) =>
  new TableCell({ children: [new Paragraph({ alignment: align, children: [new TextRun({ text, bold, size: 18 })] })], margins: { top: 40, bottom: 40, left: 60, right: 60 } });
const p = (text: string, opts: { bold?: boolean; heading?: boolean; center?: boolean; size?: number } = {}) =>
  new Paragraph({ heading: opts.heading ? HeadingLevel.HEADING_1 : undefined, alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text, bold: opts.bold, size: opts.size })] });

function priceTable(rows: ConfRow[]) {
  const head = ["No", "Nama Barang", "Kode ERP", "UOM", "HNA", "Disc", "Harga Satuan", "Harga Satuan (+PPN)"];
  const body = rows.map((r) => new TableRow({ children: [cell(String(r.no)), cell(r.name), cell(r.code), cell(`${r.uom} ${r.base_unit}`), cell(fmtIdr(r.hna), false, AlignmentType.RIGHT), cell(fmtPctText(r.disc), false, AlignmentType.RIGHT), cell(fmtIdr(r.unit_price), false, AlignmentType.RIGHT), cell(fmtIdr(r.unit_price_vat), false, AlignmentType.RIGHT)] }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "999999" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "999999" }, left: { style: BorderStyle.SINGLE, size: 4, color: "999999" }, right: { style: BorderStyle.SINGLE, size: 4, color: "999999" }, insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" }, insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" } },
    rows: [new TableRow({ tableHeader: true, children: head.map((h) => cell(h, true)) }), ...body],
  });
}

function schemeText(meta: DocMeta): string {
  return meta.scheme === "Nett"
    ? "Skema Nett: harga satuan dihitung dari HNA yang disepakati dikurangi diskon, ditambah PPN."
    : "Skema Disc: harga satuan dihitung dari HNA formularium Siloam dikurangi diskon yang disepakati, ditambah PPN.";
}

export async function buildBakDocx(meta: DocMeta, rows: ConfRow[]): Promise<Buffer> {
  const today = new Date().toISOString().slice(0, 10);
  const doc = new Document({
    sections: [{
      children: [
        p("DRAFT - BERITA ACARA KESEPAKATAN HARGA (BAK)", { bold: true, center: true, size: 26 }),
        p(`Skema ${meta.scheme}`, { center: true, bold: true }),
        p(`Tanggal: ${fmtLong(today)}`),
        p(`Pada hari ini telah dilaksanakan negosiasi harga antara Siloam Hospitals dan ${meta.principal}${meta.distributor ? ` (distributor: ${meta.distributor})` : ""} untuk kategori ${meta.category}.`),
        p(`Periode kontrak: ${fmtLong(meta.contract_start)} s/d ${fmtLong(meta.contract_end)}. Biaya pengiriman: ${meta.delivery_fee}.`),
        p(schemeText(meta)),
        p(`Negosiasi online: ${meta.meeting_at ? fmtLong(meta.meeting_at) : "-"}. Peserta Siloam: ${meta.attendees_siloam || "-"}. Peserta principal: ${meta.attendees_principal || "-"}.`),
        p("Kesepakatan harga untuk SKU berikut:", { bold: true }),
        priceTable(rows),
        p(""),
        p(`Jumlah SKU: ${meta.sku_count}. Harga sudah termasuk PPN ${Math.round(meta.ppn * 100)}% pada kolom Harga Satuan (+PPN). Lampiran: Excel Confirmation.`),
        ...(meta.notes ? [p("Catatan negosiasi:", { bold: true }), p(meta.notes)] : []),
        ...(meta.followups?.length ? [p("Tindak lanjut:", { bold: true }), ...meta.followups.map((f) => p(`- ${f.text}${f.owner ? ` (${f.owner})` : ""}${f.due ? `, paling lambat ${fmtLong(f.due)}` : ""}`))] : []),
        p(""),
        p("Demikian berita acara ini dibuat untuk dijadikan dasar pembuatan Perjanjian Kerja Sama."),
        p(""),
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: [cell("Siloam Hospitals", true, AlignmentType.CENTER), cell(meta.principal, true, AlignmentType.CENTER)] }), new TableRow({ children: [cell("\n\n\n(.........................)", false, AlignmentType.CENTER), cell("\n\n\n(.........................)", false, AlignmentType.CENTER)] })] }),
      ],
    }],
  });
  return Buffer.from(await Packer.toBuffer(doc));
}

export async function buildPksDocx(meta: DocMeta, rows: ConfRow[]): Promise<Buffer> {
  const clauses = [
    ["Pasal 1 - Ruang Lingkup", `${meta.principal} menyediakan barang kategori ${meta.category} kepada Siloam Hospitals sesuai daftar harga pada Lampiran (hasil BAK), sebanyak ${meta.sku_count} SKU.`],
    ["Pasal 2 - Jangka Waktu", `Perjanjian berlaku ${fmtLong(meta.contract_start)} s/d ${fmtLong(meta.contract_end)}.`],
    ["Pasal 3 - Harga", `${schemeText(meta)} Harga yang berlaku adalah harga pada Lampiran. Perubahan harga hanya sah dengan addendum tertulis.`],
    ["Pasal 4 - Pengiriman", `Biaya pengiriman: ${meta.delivery_fee}.`],
    ["Pasal 5 - Pembayaran", "Ketentuan pembayaran mengikuti kebijakan pembayaran Siloam Hospitals. (isi sesuai standar Legal)"],
    ["Pasal 6 - Penyelesaian Perselisihan", "Perselisihan diselesaikan secara musyawarah. (isi sesuai standar Legal)"],
  ];
  const doc = new Document({
    sections: [{
      children: [
        p("DRAFT - PERJANJIAN KERJA SAMA (PKS)", { bold: true, center: true, size: 26 }),
        p(`Antara Siloam Hospitals dan ${meta.principal}`, { center: true }),
        p("Draft ini dibuat otomatis dari hasil BAK dan perlu ditinjau oleh Legal sebelum ditandatangani.", { size: 18 }),
        ...clauses.flatMap(([h, t]) => [p(h, { bold: true }), p(t)]),
        p("Lampiran - Daftar Harga", { bold: true }),
        priceTable(rows),
      ],
    }],
  });
  return Buffer.from(await Packer.toBuffer(doc));
}

// Fills {{placeholders}} of a Siloam .docx template. Word often splits typed text into runs, so each paragraph's
// runs are merged before replacing; the price table lines go where {{tabel_harga}} stands.
export async function fillDocxTemplate(buf: Buffer, meta: DocMeta, rows: ConfRow[]): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buf);
  const f = zip.file("word/document.xml");
  if (!f) throw new Error("Template Word tidak valid (tidak ada word/document.xml).");
  let xml = await f.async("string");
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const today = new Date().toISOString().slice(0, 10);
  const vals: Record<string, string> = {
    principal: meta.principal, distributor: meta.distributor || "-", kategori: meta.category, binding: meta.binding, skema: meta.scheme,
    tanggal: fmtLong(today), kontrak_mulai: fmtLong(meta.contract_start), kontrak_selesai: fmtLong(meta.contract_end), biaya_kirim: meta.delivery_fee,
    jumlah_sku: String(meta.sku_count), nilai_po_setahun: fmtIdr(meta.po_value_year, 0), cost_impact: fmtIdr(meta.cost_impact, 0),
    tanggal_meeting: fmtLong(meta.meeting_at), peserta_siloam: meta.attendees_siloam || "-", peserta_principal: meta.attendees_principal || "-", catatan: meta.notes || "-",
    ppn: `${Math.round(meta.ppn * 100)}%`,
  };
  const table = rows.map((r) => `${r.no}. ${r.name} (${r.code}) | ${r.uom} ${r.base_unit} | HNA ${fmtIdr(r.hna)} | Disc ${fmtPctText(r.disc)} | ${fmtIdr(r.unit_price)} | +PPN ${fmtIdr(r.unit_price_vat)}`);
  xml = xml.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, (para) => {
    const plain = [...para.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join("");
    if (!plain.includes("{{")) return para;
    const open = para.match(/^<w:p[^>]*>/)![0];
    const pPr = (para.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [""])[0];
    const rPr = (para.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [""])[0];
    const mk = (t: string) => `${open}${pPr}<w:r>${rPr}<w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
    if (plain.includes("{{tabel_harga}}")) return table.map((l) => mk(esc(l))).join("");
    return mk(plain.replace(/\{\{\s*(\w+)\s*\}\}/g, (_m, k) => esc(vals[k] ?? `{{${k}}}`)));
  });
  zip.file("word/document.xml", xml);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}
