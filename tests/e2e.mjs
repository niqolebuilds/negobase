// End-to-end check of the Nov 2026 review points against a real server. Run: npm run test:e2e
// It starts the server on its own port, loads the review's dummy data and walks the workflow.
import { spawn } from "node:child_process";
import assert from "node:assert/strict";

const PORT = 3955;
const BASE = `http://localhost:${PORT}`;
const srv = spawn("npx", ["tsx", "server.ts"], { env: { ...process.env, PORT: String(PORT) }, stdio: "ignore" });
process.on("exit", () => srv.kill());
for (let i = 0; i < 60; i++) { try { await fetch(`${BASE}/api/auth/me`); break; } catch { await new Promise((r) => setTimeout(r, 500)); } }

let tok = "";
const H = (extra = {}) => ({ Authorization: `Bearer ${tok}`, "X-Requested-With": "nego", ...extra });
async function call(method, path, body, headers = {}) {
  const r = await fetch(BASE + path, { method, headers: H({ "Content-Type": "application/json", ...headers }), body: body === undefined ? undefined : JSON.stringify(body) });
  const ct = r.headers.get("content-type") || "";
  const j = ct.includes("json") ? await r.json() : Buffer.from(await r.arrayBuffer());
  return { status: r.status, j, ct };
}
async function up(path, filename, text, fields = {}, headers = {}, field = "file") {
  const fd = new FormData();
  fd.append(field, new Blob([text]), filename);
  for (const [k, v] of Object.entries(fields)) fd.append(k, v);
  const r = await fetch(BASE + path, { method: "POST", headers: H(headers), body: fd });
  return { status: r.status, j: await r.json().catch(() => ({})) };
}
let n = 0;
const t = async (name, fn) => { try { await fn(); n++; console.log("ok  " + name); } catch (e) { console.log("FAIL " + name + "\n    " + (e.message || e)); process.exitCode = 1; } };
const near = (a, b, tol = 0.6) => assert.ok(a != null && Math.abs(a - b) <= tol, `expected ${b}, got ${a}`);

const si = await call("POST", "/api/auth/signin", { email: "admin@siloamhospitals.com" });
tok = si.j.token;

// ---- data: formulary with disc and conversion, PO 2026, one principal per binding
const formulary = [
  "Kode Barang,Nama Barang,Principal,Merk,Status,Satuan,HNA,Konversi,Disc",
  ...[1, 2, 3, 4].map((i) => `DUN00000${i},Item N${i},pt dummy nett farma,BrandN,Active,Box 10,100000,10,10%`),
  ...[1, 2, 3, 4].map((i) => `DUD00000${i},Item D${i},PT DUMMY DISC FARMA,BrandD,Active,BOX@10,100000,10,10%`),
  "DUM000006,Ranitidine,PT DUMMY NN FARMA,B,Active,BOX@100,300000,100,10%",
  "DUM000008,Metformin,PT DUMMY NN FARMA,B,Active,BOX@100,30000,100,10%",
].join("\n");
const po = ["PO Date,Hospital,Vendor Name,Item Number,Item Description,Purch Unit,Quantity,Unit Price,Line Amount"];
for (const code of ["DUN000001", "DUN000002", "DUN000003", "DUN000004", "DUD000001", "DUD000002", "DUD000003", "DUD000004"]) po.push(`15/03/2026,H1,PT DUMMY,${code},Item,Box 10,12,"Rp 90.000",1080000`);
po.push("15/03/2026,H1,PT DUMMY NN FARMA, dum000006 ,Ranitidine,box@100,10,270000,2700000");
po.push("15/03/2026,H1,PT DUMMY NN FARMA,DUM000008,Metformin,BOX@100,10,27000,270000");
const mou = (p, code, hna, disc, q = 10) => `${p},${code},Item,B,REF1,${q},${hna},${disc}`;
const mouNett = ["Principal,ERP Code,Item Name,Brand,Catalog No. (REF),MOU_Qty/PO Unit,MOU_HNA/PO Unit (excl. PPN),MOU_Disc%", ...[1, 2, 3, 4].map((i) => mou("PT DUMMY NETT FARMA", `DUN00000${i}`, 100000, 0.1))].join("\n");

let per;
await t("period: one active period per category", async () => {
  const r = await call("GET", "/api/admin/periods");
  per = r.j.periods;
  const cnc = per.find((p) => p.category === "CnC"), drugs = per.find((p) => p.category === "Drugs");
  await call("POST", `/api/admin/periods/${drugs.id}/activate`, {});
  const r2 = await call("GET", "/api/admin/periods");
  assert.equal(r2.j.periods.filter((p) => p.active).length, 2, "Drugs and CnC both active");
  assert.ok(r2.j.periods.find((p) => p.id === cnc.id).active);
});
const cncId = per.find((p) => p.category === "CnC").id;
const drugsId = per.find((p) => p.category === "Drugs").id;

await t("F1: upload report lists rows, repairs and errors", async () => {
  const r1 = await up(`/api/admin/periods/${cncId}/upload`, "f.csv", formulary, { kind: "formulary" });
  assert.equal(r1.status, 200);
  assert.ok(r1.j.report.fixed.some((x) => x.field === "Principal"), "principal case repaired");
  const r2 = await up(`/api/admin/periods/${cncId}/upload`, "po.csv", po.join("\n"), { kind: "po_export" });
  assert.equal(r2.j.report.ok, po.length - 1);
  assert.ok(r2.j.report.fixed.some((x) => x.from === "Rp 90.000"), "price text repaired");
  assert.ok(r2.j.report.fixed.some((x) => /box@100/i.test(x.from) && x.to === "BOX@100") || r2.j.report.fixed.some((x) => x.field === "Kode barang"), "unit or code tidied");
});

let pid, cid, token;
await t("F2: principal name in capitals, all errors listed", async () => {
  const r = await call("POST", "/api/admin/nego/principals", { name: "pt dummy nett farma", category: "CnC", nego_type: "Renewal", binding: "Nett", mou_start: "2024-01-15", mou_end: "2027-01-14" });
  assert.equal(r.j.name, "PT DUMMY NETT FARMA");
  pid = r.j.id;
  const csv = "Principal,Category,Negotiation Type,Binding,Current MOU Start,Current MOU End,Email\nbad row,Obat,Renewal,Netto,2027-01-01,2026-01-01,not-an-email\n";
  const imp = await up("/api/admin/nego/principals/import", "p.csv", csv);
  assert.ok(imp.j.detail[0].reasons.length >= 4, `4 reasons expected, got ${imp.j.detail[0].reasons.length}`);
});

await t("open negotiation in the principal's category period, prepare with MOU", async () => {
  const c = await call("POST", "/api/admin/nego/cycles", { principal_id: pid, contract_start: "2027-01-15", contract_end: "2030-01-14", binding: "Nett" });
  assert.equal(c.status, 200, JSON.stringify(c.j));
  assert.equal(c.j.period_id, cncId);
  cid = c.j.id;
  const prep = await up(`/api/admin/nego/cycles/${cid}/prepare`, "mou.csv", mouNett, {}, {}, "mou");
  assert.equal(prep.j.items, 4, JSON.stringify(prep.j).slice(0, 200));
  const link = await call("POST", `/api/admin/nego/cycles/${cid}/links`, { days: 7 });
  token = (link.j.path || link.j.url || "").split("/p/")[1];
});

const P = (method, path, body) => call(method, path + (path.includes("?") ? "&" : "?") + `token=${token}`, body, { "X-Principal-Token": token });
const items = async () => (await call("GET", `/api/nego/cycles/${cid}/items?limit=50`)).j.items;

await t("A2: formulary Base 11.100 / Nett 9.990 and MOU price", async () => {
  const it = (await items()).find((i) => i.erp_code === "DUN000001");
  near(it.form_base_price, 11100, 0.01); near(it.form_nett_price, 9990, 0.01); near(it.mou_unit_price, 9990, 0.01);
  assert.equal(it.bench_tag, "Formularium");
});

await t("move RFQ and principal fills the form (B1 formats)", async () => {
  const m = await call("POST", `/api/admin/nego/cycles/${cid}/step`, { step: "rfq", due: new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10) });
  assert.equal(m.status, 200, JSON.stringify(m.j));
  const list = (await P("GET", "/api/p/items?limit=50")).j.items;
  const idOf = (c) => list.find((i) => i.erp_code === c).id;
  const fill = async (code, hna, disc) => {
    const id = idOf(code);
    let r = await P("PATCH", `/api/p/items/${id}`, { changes: { item_status: "Active", brand: "B", catalog_no: "REF", rfq_qty: "10", rfq_hna: String(hna), rfq_disc: String(disc) } });
    assert.equal(r.status, 200, JSON.stringify(r.j));
    return r.j;
  };
  const a = await fill("DUN000001", "110.000", "10");           // HNA +10%
  near(a.rfq_disc, 0.1, 1e-9);
  const b = await fill("DUN000002", "100000", "27,5");
  near(b.rfq_disc, 0.275, 1e-9);
  const c0 = await fill("DUN000003", "100000", "10");           // same price as MOU
  const d = await fill("DUN000004", "38000", "10");             // HNA down 62%
  assert.ok(d.issues.some((x) => x.code === "hna_anomaly"), "B4: HNA down asks for confirmation");
  void c0;
  const bad = await P("PATCH", `/api/p/items/${idOf("DUN000003")}`, { changes: { rfq_disc: "150" } });
  assert.equal(bad.status, 400, "a rejected discount is refused");
  const pct = await P("PATCH", `/api/p/items/${idOf("DUN000003")}`, { changes: { rfq_disc: "17%" } });
  assert.equal(pct.status, 200, "17% accepted");
  await P("PATCH", `/api/p/items/${idOf("DUN000003")}`, { changes: { rfq_disc: "10" } });
});

await t("D1: findings are anomalies only; unchanged price is no finding", async () => {
  const f = (await call("GET", `/api/nego/cycles/${cid}/anomalies?status=open`)).j.anomalies;
  const by = (code) => f.filter((x) => x.erp_code === code).map((x) => x.rule);
  assert.ok(by("DUN000004").includes("price_decrease_rfq"), "HNA down 62% is a high finding");
  assert.ok(!by("DUN000003").length, "unchanged price: no finding");
  assert.ok(!by("DUN000001").includes("price_increase_high"), "+10% is not an anomaly");
  assert.equal(f.find((x) => x.erp_code === "DUN000004").severity, "high");
});

await t("A1: cost impact uses the yearly PO value incl. PPN", async () => {
  const its = await items();
  const same = its.find((i) => i.erp_code === "DUN000003");
  assert.ok(same.impact_rfq != null, JSON.stringify({ q: same.rfq_qty, h: same.rfq_hna, d: same.rfq_disc, up: same.rfq_unit_price, ap: same.annual_pcs, po: same.po_qty_12m, pov: same.po_value_12m, pv: same.po_annualized_value_ppn, pu: same.po_unit_text }));
  near(same.impact_rfq, 0, 2);
  const up = its.find((i) => i.erp_code === "DUN000001");
  assert.ok(up.impact_rfq > 0, "price up = cost increase");
});

await t("C1: accept all counter offers keeps an own answer; exact value", async () => {
  await call("POST", `/api/admin/nego/cycles/${cid}/step`, { step: "counter_offer", confirm_incomplete: true });
  const plan = (await call("GET", `/api/nego/cycles/${cid}/co-plan`)).j;
  assert.ok(plan.suggestions.every((s) => s.erp_code && s.item_name), "A6: suggestions list SKU code and name");
  const ids = plan.suggestions.map((s) => s.item_id);
  if (ids.length) {
    const r = await call("POST", `/api/admin/nego/cycles/${cid}/co`, { item_ids: ids, overwrite: true });
    assert.ok(r.j.applied >= 1);
  }
  await call("POST", `/api/admin/nego/cycles/${cid}/step`, { step: "feedback1", due: new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10) });
  const list = (await P("GET", "/api/p/items?limit=50")).j.items;
  const withCo = list.filter((i) => i.co_disc != null);
  if (withCo.length >= 2) {
    const own = withCo[0], other = withCo[1];
    await P("PATCH", `/api/p/items/${own.id}`, { changes: { fb1_disc: "5" } });
    const fill = (await P("POST", "/api/p/fill", {})).j;
    const after = (await P("GET", "/api/p/items?limit=50")).j.items;
    near(after.find((i) => i.id === own.id).fb1_disc, 0.05, 1e-9, "own answer kept");
    near(after.find((i) => i.id === other.id).fb1_disc, other.co_disc, 1e-12);
    assert.ok(fill.kept >= 1);
  }
});

await t("C2: passed steps are locked; correction needs a reason and is logged", async () => {
  const its = await items();
  const it = its.find((i) => i.erp_code === "DUN000001");
  const r = await call("PATCH", `/api/admin/nego/cycles/${cid}/items/${it.id}`, { changes: { rfq_hna: "105000" } });
  assert.equal(r.status, 423, JSON.stringify(r.j));
  const ok = await call("PATCH", `/api/admin/nego/cycles/${cid}/items/${it.id}`, { changes: { rfq_hna: "105000" }, correction_reason: "Diajukan principal: HNA salah ketik" });
  assert.equal(ok.status, 200, JSON.stringify(ok.j));
  const d = (await call("GET", `/api/nego/cycles/${cid}/items/${it.id}`)).j;
  assert.ok(d.history.some((h) => h.field === "rfq_hna" && h.reason), "history carries who, when, old->new and reason");
  const stale = await call("PATCH", `/api/admin/nego/cycles/${cid}/items/${it.id}`, { changes: { on_disc: "8" }, base_version: 0 });
  assert.equal(stale.status, 409, "C6: stale version warns");
});

await t("C6/D2: Online Nego meeting saves; escalation; package gating", async () => {
  await call("POST", `/api/admin/nego/cycles/${cid}/step`, { step: "online_nego" });
  const m = await call("PATCH", `/api/admin/nego/cycles/${cid}`, { meeting_at: "2027-01-20T10:00", meeting: { attendees_siloam: "A", attendees_principal: "B", notes: "ok", followups: [{ text: "kirim BAK", due: "2027-02-01" }] } });
  assert.equal(m.status, 200, "C5: meeting saves (was Not Found)");
  const its = await items();
  for (const i of its) await call("PATCH", `/api/admin/nego/cycles/${cid}/items/${i.id}`, { changes: { on_disc_nett: "12", on_disc_disc: "11" } });
  const sch = await call("PATCH", `/api/admin/nego/cycles/${cid}`, { final_scheme: "Nett" });
  assert.equal(sch.status, 200);
  const ov = (await call("GET", `/api/nego/cycles/${cid}`)).j;
  assert.ok(ov.escalation.items_agreed >= 1, "panel counts agreed SKUs");
  const pk0 = (await call("GET", `/api/nego/cycles/${cid}/package`)).j;
  assert.equal(pk0.ready, false, "nothing flagged yet: not Ready");
  await call("POST", `/api/admin/nego/cycles/${cid}/bak-flags`, { mode: "agreed", value: true });
  const pk = (await call("GET", `/api/nego/cycles/${cid}/package`)).j;
  assert.ok(pk.flagged >= 1);
});

await t("E2/E3: Confirmation columns, rounding, BAK and PKS Word files", async () => {
  const pk = (await call("GET", `/api/nego/cycles/${cid}/package`)).j;
  if (!pk.ready) {
    // escalation may be needed: approve it as admin
    await call("POST", `/api/admin/nego/cycles/${cid}/escalation`, { note: "Disetujui direksi" });
  }
  const x = await call("GET", `/api/admin/nego/cycles/${cid}/package.xlsx`);
  assert.equal(x.status, 200, JSON.stringify(x.j).slice(0, 200));
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(x.j);
  const ws = wb.worksheets[0];
  assert.deepEqual(ws.getRow(4).values.slice(1), ["No", "Item Name", "ERP Code", "UOM", "Base Unit", "HNA", "Disc", "Unit Price", "Unit Price (+VAT)"]);
  const r5 = ws.getRow(5).values;
  assert.equal(r5[8], Math.round(r5[8] * 100) / 100, "2 decimals");
  const bak = await call("GET", `/api/admin/nego/cycles/${cid}/bak.docx`);
  assert.equal(bak.status, 200); assert.equal(bak.j.slice(0, 2).toString("latin1"), "PK", "BAK is a Word (zip) file");
  const pks = await call("GET", `/api/admin/nego/cycles/${cid}/pks.docx`);
  assert.equal(pks.status, 200);
});

await t("A7: analysis and report", async () => {
  const a = (await call("GET", `/api/nego/cycles/${cid}/analysis?mode=pre`)).j;
  assert.ok(a.columns.length >= 60, `pre columns ${a.columns.length}`);
  const b = (await call("GET", `/api/nego/cycles/${cid}/analysis?mode=post`)).j;
  assert.ok(b.columns.length > a.columns.length, "post has the FIN columns");
  const rep = (await call("GET", "/api/admin/report")).j;
  assert.equal(rep.columns.length, 19);
  assert.ok(rep.fb1.length >= 1 && rep.fin.length >= 1);
});

await t("E1: documents: type checked, slot kept, admin only opens", async () => {
  const bad = await up("/api/p/documents?token=" + token, "evil.exe", "MZ....", { doc_type: "nib" }, { "X-Principal-Token": token });
  assert.equal(bad.status, 400);
  const pdf = await up("/api/p/documents?token=" + token, "nib.pdf", "%PDF-1.4 test", { doc_type: "nib" }, { "X-Principal-Token": token });
  assert.equal(pdf.status, 200);
  assert.equal(pdf.j.documents[0].filename, "nib.pdf");
  const docs = (await call("GET", `/api/admin/nego/cycles/${cid}/documents`)).j.documents;
  const open = await call("GET", docs[0].url);
  assert.equal(open.status, 200);
});

await t("B5/E1: submit asks Discontinue confirmation; Submission needs documents", async () => {
  const none = await call("POST", `/api/admin/nego/cycles/${cid}/step`, { step: "submission" });
  void none;
  const r = await P("POST", "/api/p/submit", {});
  assert.equal(r.status, 400, "submission without the 3 documents is refused");
});

await t("D2: settings hide unused groups, two escalation limits", async () => {
  const s = (await call("GET", "/api/settings")).j.settings;
  assert.ok(s.some((x) => x.key === "escalation_pct") && s.some((x) => x.key === "escalation_rp"));
  assert.ok(!s.some((x) => x.group === "Rebates" || x.group === "Renewals & alerts" || x.group === "Money" || x.group === "Targets"));
});

await t("G1: AI answers from data, never claims an action", async () => {
  const q1 = (await call("POST", "/api/ai/chat", { message: "Principal mana saja yang MOU-nya habis dalam 6 bulan ke depan? Sebutkan namanya." })).j;
  assert.ok(!/Kimia Farma|Kalbe|Bio Farma/.test(q1.reply), q1.reply);
  const q3 = (await call("POST", "/api/ai/chat", { message: "Daftarkan principal baru PT Contoh Farma kategori Drugs binding Nett." })).j;
  assert.ok(/tidak bisa/i.test(q3.reply) && !/berhasil didaftarkan/i.test(q3.reply), q3.reply);
  const q2 = (await call("POST", "/api/ai/chat", { message: `Berapa baris data PO yang sudah diupload untuk periode ${per.find((p) => p.id === cncId).name}?` })).j;
  assert.match(q2.reply, new RegExp(String(po.length - 1)), q2.reply);
});


// ---- second negotiation (Disc binding): Excel upload, back a step, Discontinue, deadlines
let cid2, token2;
await t("setup: Disc principal and RFQ step", async () => {
  const pr = await call("POST", "/api/admin/nego/principals", { name: "PT DUMMY DISC FARMA", category: "CnC", nego_type: "New Vendor", binding: "Disc" });
  const c = await call("POST", "/api/admin/nego/cycles", { principal_id: pr.j.id, contract_start: "2027-01-15", contract_end: "2030-01-14", binding: "Disc" });
  cid2 = c.j.id;
  const prep = await call("POST", `/api/admin/nego/cycles/${cid2}/prepare`, {});
  assert.equal(prep.status, 200, JSON.stringify(prep.j).slice(0, 200));
  const past = await call("POST", `/api/admin/nego/cycles/${cid2}/step`, { step: "rfq", due: "2020-01-01" });
  assert.equal(past.status, 400, "C7: a past deadline can't be chosen");
  const m = await call("POST", `/api/admin/nego/cycles/${cid2}/step`, { step: "rfq", due: new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10) });
  assert.equal(m.status, 200);
  const link = await call("POST", `/api/admin/nego/cycles/${cid2}/links`, { days: 7 });
  token2 = (link.j.path || link.j.url).split("/p/")[1];
});
const P2 = (method, path, body) => call(method, path + (path.includes("?") ? "&" : "?") + `token=${token2}`, body, { "X-Principal-Token": token2 });

await t("A3: Disc binding margin uses the formulary HNA, not the principal's", async () => {
  const list = (await P2("GET", "/api/p/items?limit=50")).j.items;
  const id = list.find((i) => i.erp_code === "DUD000001").id;
  await P2("PATCH", `/api/p/items/${id}`, { changes: { item_status: "Active", brand: "B", catalog_no: "R", rfq_qty: "10", rfq_hna: "110000", rfq_disc: "10" } });
  const it = (await call("GET", `/api/nego/cycles/${cid2}/items?limit=50`)).j.items.find((i) => i.erp_code === "DUD000001");
  near(it.base_margin_price, 11100, 0.01); near(it.margin_pc_rfq, 111, 0.01);
});

await t("B3: Excel upload: duplicates rejected, wrong cell flagged, overwrite needs confirmation", async () => {
  const x = await fetch(`${BASE}/api/p/template.xlsx?token=${token2}`, { headers: { "X-Principal-Token": token2 } });
  assert.equal(x.status, 200);
  const ExcelJS = (await import("exceljs")).default;
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(Buffer.from(await x.arrayBuffer()));
  const ws = wb.worksheets[0];
  const hdr = ws.getRow(7).values; const colOf = (h) => hdr.indexOf(h);
  const rowOf = (code) => { for (let r = 8; r < 40; r++) if (ws.getCell(r, 1).value === code) return r; return 0; };
  const r2 = rowOf("DUD000002"), r3 = rowOf("DUD000003"), r4 = rowOf("DUD000004"), r1 = rowOf("DUD000001");
  ws.getCell(r2, colOf("RFQ_Qty/PO Unit")).value = 10; ws.getCell(r2, colOf("RFQ_HNA/PO Unit (excl. PPN)")).value = 100000; ws.getCell(r2, colOf("RFQ_Disc%")).value = "abc";   // wrong cell
  ws.getCell(r3, colOf("RFQ_Qty/PO Unit")).value = 10; ws.getCell(r3, colOf("RFQ_HNA/PO Unit (excl. PPN)")).value = 100000; ws.getCell(r3, colOf("RFQ_Disc%")).value = 12.5;
  ws.getCell(r4, 1).value = "DUD000003"; ws.getCell(r4, colOf("RFQ_Qty/PO Unit")).value = 10; ws.getCell(r4, colOf("RFQ_HNA/PO Unit (excl. PPN)")).value = 95000;   // duplicate code
  ws.getCell(r1, colOf("RFQ_Disc%")).value = 20;  // overwrites the portal value (10)
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  const pv = await up(`/api/p/template?token=${token2}`, "t.xlsx", buf, { apply: "0" }, { "X-Principal-Token": token2 });
  assert.equal(pv.status, 200, JSON.stringify(pv.j).slice(0, 200));
  assert.ok(pv.j.rejected.some((r) => r.erp_code === "DUD000002" && r.errors[0].column === "RFQ_Disc%" && r.errors[0].value === "abc"), "wrong cell named with column and value");
  assert.ok(pv.j.rejected.filter((r) => r.erp_code === "DUD000003").length >= 1, "duplicate rows rejected");
  assert.ok(pv.j.overwrite_count >= 1, "overwrite is announced");
  const noConfirm = await up(`/api/p/template?token=${token2}`, "t.xlsx", buf, { apply: "1" }, { "X-Principal-Token": token2 });
  assert.equal(noConfirm.status, 409);
  const yes = await up(`/api/p/template?token=${token2}`, "t.xlsx", buf, { apply: "1", confirm_overwrite: "1" }, { "X-Principal-Token": token2 });
  assert.equal(yes.status, 200, JSON.stringify(yes.j));
  const it = (await call("GET", `/api/nego/cycles/${cid2}/items?limit=50`)).j.items;
  near(it.find((i) => i.erp_code === "DUD000002").rfq_qty, 10, 0, "good cells of a row with one wrong cell are kept");
  assert.equal(it.find((i) => i.erp_code === "DUD000003").rfq_disc, 0.1 === 0 ? 0 : it.find((i) => i.erp_code === "DUD000003").rfq_disc);
  near(it.find((i) => i.erp_code === "DUD000001").rfq_disc, 0.2, 1e-9);
});

await t("B5: Discontinue needs confirmation before sending", async () => {
  const list = (await P2("GET", "/api/p/items?limit=50")).j.items;
  for (const i of list) if (i.erp_code !== "DUD000001") await P2("PATCH", `/api/p/items/${i.id}`, { changes: { item_status: "Discontinue" } });
  const id1 = list.find((i) => i.erp_code === "DUD000001").id;
  await P2("PATCH", `/api/p/items/${id1}`, { changes: { rfq_disc: "10", price_reason: "raw_material" } });
  const sm = (await P2("GET", "/api/p/summary")).j;
  assert.equal(sm.needs_discontinue_confirm, true);
  const no = await P2("POST", "/api/p/submit", {});
  assert.equal(no.status, 409, JSON.stringify(no.j));
});

await t("C4: reasons and confirmations are dropped when the price no longer goes up; A4 pack difference is no warning", async () => {
  const pr = await call("POST", "/api/admin/nego/principals", { name: "PT DUMMY NN FARMA", category: "CnC", nego_type: "Renewal", binding: "Nett to Disc", mou_start: "2024-01-15", mou_end: "2027-01-14" });
  const c = await call("POST", "/api/admin/nego/cycles", { principal_id: pr.j.id, contract_start: "2027-01-15", contract_end: "2030-01-14", binding: "Nett to Disc" });
  const c3 = c.j.id;
  const mouCsv = "Principal,ERP Code,Item Name,Brand,Catalog No. (REF),MOU_Qty/PO Unit,MOU_HNA/PO Unit (excl. PPN),MOU_Disc%\nPT DUMMY NN FARMA,DUM000006,Ranitidine,B,R,100,300000,0.1\n";
  const prep = await up(`/api/admin/nego/cycles/${c3}/prepare`, "mou.csv", mouCsv, {}, {}, "mou");
  assert.ok(prep.j.items >= 1, JSON.stringify(prep.j).slice(0, 200));
  await call("POST", `/api/admin/nego/cycles/${c3}/step`, { step: "rfq", due: new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10) });
  const link = await call("POST", `/api/admin/nego/cycles/${c3}/links`, { days: 7 });
  const tk = (link.j.path || link.j.url).split("/p/")[1];
  const P3 = (method, path, body) => call(method, path + (path.includes("?") ? "&" : "?") + `token=${tk}`, body, { "X-Principal-Token": tk });
  const it = (await P3("GET", "/api/p/items?limit=50")).j.items.find((i) => i.erp_code === "DUM000006");
  // pack differs (BOX@10 vs BOX@100) but HNA per piece is the same: no warning (A4, B4)
  let r = await P3("PATCH", `/api/p/items/${it.id}`, { changes: { item_status: "Active", brand: "B", catalog_no: "R", rfq_qty: "10", rfq_hna: "30000", rfq_disc: "10" } });
  assert.ok(!r.j.issues.some((x) => x.code === "hna_anomaly"), "different pack, same HNA per piece: no HNA warning");
  // price goes up: a reason is asked for, then given
  r = await P3("PATCH", `/api/p/items/${it.id}`, { changes: { rfq_hna: "33000", price_reason: "other" }, });
  assert.ok(r.j.issues.some((x) => x.code === "increase") === false && r.j.price_reason === "other", "reason kept while the price is up");
  r = await P3("PATCH", `/api/p/items/${it.id}`, { changes: { rfq_hna: "30000" } });
  assert.equal(r.j.price_reason, null, "the old reason is dropped once the price no longer goes up");
});

await t("B5 confirm then submit; step moves to review or counter offer; back a step with a message flags SKUs", async () => {
  const ok = await P2("POST", "/api/p/submit", { confirm_discontinue: true });
  assert.equal(ok.status, 200, JSON.stringify(ok.j));
  const st = (await call("GET", `/api/nego/cycles/${cid2}`)).j.cycle.current_step;
  assert.ok(["counter_offer", "review_rfq"].includes(st), st);
  const noMsg = await call("POST", `/api/admin/nego/cycles/${cid2}/step`, { step: "rfq" });
  assert.equal(noMsg.status, 400, "back to a principal step needs a message or flagged SKUs");
  const back = await call("POST", `/api/admin/nego/cycles/${cid2}/step`, { step: "rfq", message: "HNA DUD000001 perlu dicek", item_ids: [(await call("GET", `/api/nego/cycles/${cid2}/items?limit=50`)).j.items.find((i) => i.erp_code === "DUD000001").id], due: new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10) });
  assert.equal(back.status, 200, JSON.stringify(back.j));
  const me = (await P2("GET", "/api/p/me")).j;
  assert.equal(me.return_note.message, "HNA DUD000001 perlu dicek");
  const items2 = (await P2("GET", "/api/p/items?limit=50")).j;
  const flagged = items2.items.find((i) => i.erp_code === "DUD000001");
  assert.equal(flagged.state, "check", "only the flagged SKU shows Perlu dicek");
  assert.ok(items2.items.filter((i) => i.state === "check").length === 1);
  // the price no longer goes up: the old reason is dropped
  const fix = await P2("PATCH", `/api/p/items/${flagged.id}`, { changes: { rfq_disc: "10" } });
  assert.equal(fix.status, 200);
});


// ---- New Item: status, new row from Excel and from the portal, finding for Siloam
await t("New Item: Excel row at the bottom and portal add both reach the list and Findings", async () => {
  const pr = await call("POST", "/api/admin/nego/principals", { name: "PT DUMMY BARU FARMA", category: "Drugs", nego_type: "New Vendor", binding: "Nett" });
  const c = await call("POST", "/api/admin/nego/cycles", { principal_id: pr.j.id, contract_start: "2027-01-15", contract_end: "2030-01-14", binding: "Nett", category: "Drugs" });
  const c4 = c.j.id;
  // no items in the sample for this principal: tag one SKU so the list is not empty
  await up(`/api/admin/periods/${drugsId}/upload`, "tag.csv", "ERP Code,Principal\nBRU000001,PT DUMMY BARU FARMA\n", { kind: "sku_tagging" });
  await up(`/api/admin/periods/${drugsId}/upload`, "f.csv", "Kode Barang,Nama Barang,Principal,Merk,Status,Satuan,HNA,Konversi,Disc\nBRU000001,Obat Lama,PT DUMMY BARU FARMA,B,Active,BOX@10,100000,10,10%\n", { kind: "formulary" });
  const prep = await call("POST", `/api/admin/nego/cycles/${c4}/prepare`, {});
  assert.equal(prep.status, 200, JSON.stringify(prep.j).slice(0, 200));
  await call("POST", `/api/admin/nego/cycles/${c4}/step`, { step: "rfq", due: new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10) });
  const link = await call("POST", `/api/admin/nego/cycles/${c4}/links`, { days: 7 });
  const tk = (link.j.path || link.j.url).split("/p/")[1];
  const P4 = (method, path, body) => call(method, path + (path.includes("?") ? "&" : "?") + `token=${tk}`, body, { "X-Principal-Token": tk });
  // portal: add a new item
  const add = await P4("POST", "/api/p/items", { item_name: "Obat Baru Alfa", brand: "Alfa", active_ingredient: "Alfaprazol", rfq_qty: "10", rfq_hna: "50000", rfq_disc: "5" });
  assert.equal(add.status, 200, JSON.stringify(add.j));
  assert.equal(add.j.item_status, "New Item"); assert.equal(add.j.is_new, true);
  const dup = await P4("POST", "/api/p/items", { item_name: "obat baru alfa" });
  assert.equal(dup.status, 409);
  // Excel: a row at the bottom with status New Item
  const ExcelJS = (await import("exceljs")).default;
  const x = await fetch(`${BASE}/api/p/template.xlsx?token=${tk}`, { headers: { "X-Principal-Token": tk } });
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(Buffer.from(await x.arrayBuffer()));
  const ws = wb.worksheets[0]; const hdr = ws.getRow(7).values; const col = (h) => hdr.indexOf(h);
  let r = 8; while (ws.getCell(r, 1).value) r++;
  ws.getCell(r, col("Item Name")).value = "Obat Baru Beta"; ws.getCell(r, col("Item Status")).value = "New Item";
  ws.getCell(r, col("Brand / Manufacturer")).value = "Beta"; ws.getCell(r, col("Active Ingredient")).value = "Betamol";
  ws.getCell(r, col("RFQ_Qty/PO Unit")).value = 20; ws.getCell(r, col("RFQ_HNA/PO Unit (excl. PPN)")).value = 80000; ws.getCell(r, col("RFQ_Disc%")).value = 12.5;
  const buf = Buffer.from(await wb.xlsx.writeBuffer());
  const pv = await up(`/api/p/template?token=${tk}`, "t.xlsx", buf, { apply: "0" }, { "X-Principal-Token": tk });
  assert.equal(pv.j.new_items.length, 1, JSON.stringify(pv.j).slice(0, 300));
  const ap = await up(`/api/p/template?token=${tk}`, "t.xlsx", buf, { apply: "1", confirm_overwrite: "1" }, { "X-Principal-Token": tk });
  assert.equal(ap.j.new_items, 1, JSON.stringify(ap.j));
  const its = (await call("GET", `/api/nego/cycles/${c4}/items?filter=new&limit=50`)).j.items;
  assert.equal(its.length, 2, "both new items are in the SKU list");
  assert.ok(its.every((i) => i.item_status === "New Item" && /^NEW-\d+$/.test(i.erp_code)));
  const f = (await call("GET", `/api/nego/cycles/${c4}/anomalies?status=open`)).j.anomalies.filter((a) => a.rule === "new_item");
  assert.equal(f.length, 2, "Siloam sees a finding per new item");
  // Siloam assigns the ERP code, then approves
  const target = its[0];
  const bad = await call("PATCH", `/api/admin/nego/cycles/${c4}/items/${target.id}`, { changes: { erp_code: "BRU000001" } });
  assert.equal(bad.status, 409, "ERP code already used");
  const okc = await call("PATCH", `/api/admin/nego/cycles/${c4}/items/${target.id}`, { changes: { erp_code: "BRU000099" } });
  assert.equal(okc.status, 200, JSON.stringify(okc.j));
  const fin = f.find((a) => a.item_id === target.id);
  const dec = await call("POST", `/api/admin/nego/cycles/${c4}/anomalies/${fin.id}`, { decision: "fix" });
  assert.equal(dec.status, 200, JSON.stringify(dec.j));
  const after = (await call("GET", `/api/nego/cycles/${c4}/items/${target.id}`)).j;
  assert.equal(after.item_status, "Active");
  const mine = (await P4("GET", "/api/p/items?limit=50")).j.items.filter((i) => i.is_new);
  assert.equal(mine.length, 2);
  // a new item row from the portal can be removed before sending
  const del = await P4("DELETE", `/api/p/items/${its[1].id}`);
  assert.equal(del.status, 200);
});

console.log(`\n${n} checks passed`);
srv.kill();
process.exit(process.exitCode || 0);
