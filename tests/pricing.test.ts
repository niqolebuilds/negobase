// Run: npm run test:pricing. The numbers come straight from the Nov 2026 review.
import assert from "node:assert/strict";
import { enrichCycleItem, parseDiscount, normalizeUom, escalationFor, discountFromCell, summarizeCycleItems } from "../pricing.js";

const close = (a: number | null, b: number, tol = 0.6, msg = "") => assert.ok(a != null && Math.abs(a - b) <= tol, `${msg} expected ${b}, got ${a}`);
let n = 0;
const t = (name: string, fn: () => void) => { fn(); n++; console.log("ok  " + name); };

// A1 - DUX000001: price RFQ = MOU = PO = 12.021/pcs, 60 pcs a year -> cost impact 0
t("A1 unchanged price = cost impact 0", () => {
  const it = enrichCycleItem({
    item_status: "Active", mou_qty: 1, mou_hna: 10830, mou_disc: 0, rfq_qty: 1, rfq_hna: 10830, rfq_disc: 0,
    po_unit_text: "PCS", po_qty_12m: 60, po_value_12m: 649800, po_avg_price: 10830,
  }, 0.11, { binding: "Nett", months_elapsed: 12 });
  close(it.rfq_unit_price, 12021.3);
  close(it.po_annualized_value_ppn, 721278, 1);
  close(it.impact_rfq, 0, 1);
});

// A2 - DUN000001: HNA 10.000/pcs, formulary disc 10%: Base 11.100, Nett 9.990
t("A2 formulary Base and Nett price", () => {
  const it = enrichCycleItem({
    item_status: "Active", formulary_hna: 100000, form_conv: 10, form_disc: 0.1, mou_qty: 10, mou_hna: 100000, mou_disc: 0.1,
    rfq_qty: 10, rfq_hna: 100000, rfq_disc: 0.1, po_unit_text: "BOX@10", po_qty_12m: 12, po_value_12m: 1080000, po_avg_price: 90000,
  }, 0.11, { binding: "Nett", months_elapsed: 12 });
  close(it.form_base_price, 11100, 0.01);
  close(it.form_nett_price, 9990, 0.01);
  assert.equal(it.bench_tag, "Formularium");
  close(it.change_pct as number, 0, 1e-9, "same price is 0%");
  close(it.margin_vs_formulary, 0, 1, "unchanged margin impact");
});

// A3 - Disc binding takes the base from the formulary HNA, not the principal's HNA
t("A3 Disc base price from formulary", () => {
  const it = enrichCycleItem({
    item_status: "Active", formulary_hna: 100000, form_conv: 10, form_disc: 0.1, rfq_qty: 10, rfq_hna: 110000, rfq_disc: 0.1,
    po_unit_text: "BOX@10", po_qty_12m: 12, po_value_12m: 1080000, po_avg_price: 90000,
  }, 0.11, { binding: "Disc", months_elapsed: 12 });
  close(it.base_margin_price, 11100, 0.01);
  close(it.rfq_unit_price, 10989, 0.01);
  close(it.margin_pc_rfq, 111, 0.01);
});

// A4 - Nett to Disc: formulary HNA / formulary pack, whatever pack the principal typed
t("A4 Nett to Disc uses formulary pack", () => {
  const it = enrichCycleItem({
    item_status: "Active", formulary_hna: 300000, form_conv: 100, form_disc: 0.1, rfq_qty: 10, rfq_hna: 30000, rfq_disc: 0.1,
    po_unit_text: "BOX@100", po_qty_12m: 10, po_value_12m: 2700000, po_avg_price: 270000,
  }, 0.11, { binding: "Nett to Disc", months_elapsed: 12 });
  close(it.rfq_unit_price, 2997, 0.01);
  close(it.impact_rfq, 0, 1);
  assert.ok(Math.abs(it.change_pct) < 1e-9);
});

// A5 - not in MOU: blank MOU price even though the formulary has a pack
t("A5 item outside the MOU has no MOU price", () => {
  const it = enrichCycleItem({ item_status: "Active", formulary_hna: 30000, form_conv: 100, mou_qty: null, mou_hna: null }, 0.11, { binding: "Nett" });
  assert.equal(it.mou_unit_price, null);
});

// B1 - one discount format
t("B1 discount format", () => {
  assert.equal(parseDiscount("15").value, 0.15);
  assert.equal(parseDiscount("17%").value, 0.17);
  assert.equal(parseDiscount("27,5").value, 0.275);
  assert.equal(parseDiscount("27.5%").value, 0.275);
  assert.equal(parseDiscount("").value, null);
  assert.ok(parseDiscount("abc").error);
  assert.ok(parseDiscount("150").error);
  assert.equal(discountFromCell(0.275, true).value, 0.275);
  assert.ok(discountFromCell(15, true).error, "15 in a percent cell is 1500%");
  assert.equal(discountFromCell(15, false).value, 0.15);
});

t("F1 unit spelling", () => {
  assert.equal(normalizeUom("BOX@12"), "BOX@12");
  assert.equal(normalizeUom("Box 12"), "BOX@12");
  assert.equal(normalizeUom("box@12"), "BOX@12");
});

// D2 - escalation: 5% of the annual PO value or Rp 50 Jt
t("D2 escalation examples", () => {
  const e = (po: number, ci: number) => escalationFor(ci, po, 0.05, 50e6).needed;
  assert.equal(e(2e9, 60e6), true);
  assert.equal(e(400e6, 30e6), true);
  assert.equal(e(1e9, 20e6), false);
  assert.equal(e(1e9, -50e6), false);
});

t("summary totals", () => {
  const items = [
    enrichCycleItem({ item_status: "Active", mou_qty: 1, mou_hna: 100, mou_disc: 0, rfq_qty: 1, rfq_hna: 100, rfq_disc: 0, po_unit_text: "PCS", po_qty_12m: 10, po_value_12m: 1000, po_avg_price: 100 }, 0.11, { binding: "Nett" }),
  ];
  const s = summarizeCycleItems(items);
  close(s.latest_impact, 0, 1e-6);
});
console.log(`${n} tests passed`);
