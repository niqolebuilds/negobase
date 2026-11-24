// Template_Nego round trip (B1, B2): opens clean, hints name the right columns, validation to row 1000,
// a percent-formatted Excel cell (27.5%) is read as 27.5%.
import assert from "node:assert/strict";
import JSZip from "jszip";
import { buildTemplate, readTemplate, checkRow } from "../template.js";

const o = { forPrincipal: true, showCounterOffer: false, principal: "PT X", distributor: "", contractStart: "2027-01-01", contractEnd: "2030-01-01", binding: "Nett", deliveryFee: "Free", category: "CnC", ppn: 0.11 };
const items = [{ erp_code: "A1", item_name: "Item 1", brand: "B", catalog_no: "R", item_status: "Active", rfq_qty: 10, rfq_hna: 1000, rfq_disc: 0.275 }];
const wb = await buildTemplate(null, items, o);
const buf = Buffer.from(await wb.xlsx.writeBuffer());
const zip = await JSZip.loadAsync(buf);
assert.ok(!Object.keys(zip.files).some((f) => f.includes("/tables/")), "no table parts (they made Excel repair the file)");
const ws = wb.worksheets[0];
const headers = ws.getRow(7).values.slice(1) as string[];
const col = (h: string) => headers.indexOf(h) + 1;
const letter = (h: string) => ws.getColumn(col(h)).letter;
assert.match(String(ws.getCell(6, col("RFQ_Qty/PO Unit")).value), new RegExp(`kolom ${letter("RFQ_HNA/PO Unit (excl. PPN)")}`), "Qty hint names the HNA column");
assert.match(String(ws.getCell(6, col("RFQ_HNA/PO Unit (excl. PPN)")).value), new RegExp(`kolom ${letter("RFQ_Qty/PO Unit")}`), "HNA hint names the Qty column");
assert.ok(ws.getCell(1000, col("Item Status")).dataValidation, "validation reaches row 1000");
assert.ok(!ws.getCell(1001, col("Item Status")).dataValidation);
assert.equal((ws.views[0] as any).topLeftCell, "C8", "opens on the first item row");

const rows = await readTemplate(buf);
assert.equal(rows[0].erp_code, "A1");
const chk = checkRow(rows[0].cells, ["brand", "rfqQty", "rfqHna", "rfqDisc"]);
assert.equal(chk.errors.length, 0, JSON.stringify(chk.errors));
assert.equal(chk.changes.find((c) => c.field === "rfq_disc")?.value, 0.275);

// a percent-formatted cell
const wb2 = await buildTemplate(null, items, o);
const c = wb2.worksheets[0].getCell(8, col("RFQ_Disc%")); c.value = 0.275; c.numFmt = "0.0%";
const rr = await readTemplate(Buffer.from(await wb2.xlsx.writeBuffer()));
assert.equal(checkRow(rr[0].cells, ["rfqDisc"]).changes[0].value, 0.275, "27.5% cell is accepted");
// a wrong cell does not drop the row
rr[0].cells.rfqDisc = "abc";
const mixed = checkRow({ ...rr[0].cells, rfqQty: "10", rfqHna: "5000" }, ["rfqQty", "rfqHna", "rfqDisc"]);
assert.equal(mixed.changes.length, 2); assert.equal(mixed.errors.length, 1);
assert.deepEqual(Object.keys(mixed.errors[0]).sort(), ["column", "message", "value"]);
console.log("template checks passed");
