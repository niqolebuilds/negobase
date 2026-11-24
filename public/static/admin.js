"use strict";
// Admin console. The server rejects every /api/admin call from non-admins; this page
// is only the interface. Sections: price data (upload, preview, apply, versions),
// documents, engine settings, users, activity.

const adminState = { sub: "periods", staged: null };

function adminEnter(sub) {
  const nego = shell.user && shell.user.role === "negotiator";  // read-only master data, nothing else
  if (nego) sub = "master";
  $("admin-subtabs").hidden = nego;
  adminState.sub = ["periods", "principals", "report", "master", "documents", "settings", "users", "activity"].includes(sub) ? sub : "periods";
  document.querySelectorAll("#admin-subtabs a").forEach((a) => a.classList.toggle("on", a.dataset.sub === adminState.sub));
  const body = clear($("admin-body"));
  return ({ periods: adminPeriods, principals: adminPrincipals, report: adminReport, master: adminMaster, documents: adminDocs, settings: adminSettings, users: adminUsers, activity: adminActivity })[adminState.sub](body);
}

function card(title, sub, ...kids) {
  return el("section", { class: "card admin-card" }, el("h2", { text: title }), sub ? el("p", { class: "sub", text: sub }) : null, ...kids);
}

function adminField(label, input, hint) {
  return el("label", {}, label, input, hint ? el("span", { class: "hint", text: hint }) : null);
}

// ---------- negotiation periods & period data hub ----------
async function adminPeriods(body) {
  const statusBox = el("div");
  const periodListBox = el("div");
  const newPeriodFormBox = el("div");
  const uploadGridBox = el("div");

  body.append(
    card("Negotiation Periods & Central Data Hub",
      "Manage negotiation periods (e.g. CnC 2026 End-Year, Drugs 2026 Mid-Year). One period is active per category, so Drugs and CnC run at the same time. Upload the PO export, formulary, SKU tagging and market benchmark once per period. Master SKU lives in Master data.",
      statusBox),
    card("Active Period Data Hubs", "Centralized upload of period data. After each upload you see how many rows were saved, which were rejected and why, and what was repaired.", uploadGridBox),
    card("All Negotiation Periods", "Activating a period only changes the active period of its own category.", periodListBox),
    card("Create New Negotiation Period", "Start a new negotiation cycle with custom category and running months.", newPeriodFormBox)
  );

  let pData;
  try {
    pData = await api("/api/admin/periods");
  } catch (err) {
    statusBox.append(el("p", { class: "error", text: err.message }));
    return;
  }

  const periods = pData.periods || [];
  // One active period per category: Drugs and CnC run side by side, each with its own data hub.
  const actives = periods.filter((p) => p.active);
  const UPLOAD_CARDS = [
    { kind: "po_export", label: "PO Export (raw PO lines)", desc: "Raw PO lines. Qty, value, average price and yearly figures per item are calculated by the app. Units and codes are tidied automatically.", cols: "Vendor Name, Item Number, Quantity, Unit Price, Line Amount, Purch Unit, PO Date, Hospital", count: (p) => p.po_rows_count },
    { kind: "formulary", label: "Formulary (Siloam file)", desc: "The Siloam formulary: HNA, conversion, discount, Price, Base Price, HNA for SBP and type (CnC Reguler, CnC BPJS, Drugs). RFQ prices are compared with the Nett Price.", cols: "Kode Barang, Nama Barang, Principal, Merk, Status, Satuan, HNA, Konversi, Disc, Price, Base Price, HNA for SBP, Jenis", count: (p) => p.formulary_count },
    { kind: "sku_tagging", label: "SKU - Principal Tagging", desc: "Which principal supplies each SKU. This is the only place items are tagged to a principal.", cols: "ERP Code, Principal", count: (p) => p.sku_tagging_count },
    { kind: "benchmark", label: "Market benchmark (once per period)", desc: "Competitor prices from Siloam's own price file, with the source link. Used in the item table and the counter-offer plan.", cols: "ERP Code, Source, Price/PCS (incl. PPN), Link", count: (p) => p.benchmark_count },
  ];
  clear(statusBox);
  clear(uploadGridBox);
  if (!actives.length) clear(statusBox).append(el("p", { class: "muted", text: "No active period set. Select or create one below." }));
  for (const activePeriod of actives) {
    statusBox.append(
      el("div", { class: "preview ok", style: "display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 8px;" },
        el("div", {},
          el("h3", { text: `Active ${activePeriod.category} period: ${activePeriod.name}` }),
          el("p", { class: "small muted", text: `Running months: ${activePeriod.months_elapsed} · Created: ${fmtDate(String(activePeriod.created_at).slice(0, 10))} · Principals of this category read this period` })),
        el("a", { class: "act secondary", href: "#admin/principals", text: "Principals: add or import →" })));
    uploadGridBox.append(el("h3", { style: "margin: 18px 0 4px", text: `${activePeriod.category} · ${activePeriod.name}` }));
    const grid = el("div", { style: "display: grid; grid-template-columns: repeat(auto-fit, minmax(270px, 1fr)); gap: 18px; margin-top: 8px; align-items: stretch;" });
    UPLOAD_CARDS.forEach((c) => {
      const finput = el("input", { type: "file", accept: ".csv,.xlsx", class: "sr-only" });
      const info = (activePeriod.files || {})[c.kind];
      const n = c.count(activePeriod);
      const statusSpan = el("div", { class: "small", style: "min-height: 52px; display: flex; flex-direction: column; justify-content: flex-start; gap: 4px; margin-top: 6px;" },
        n ? el("span", { class: "zone good", text: `${num(n)} rows loaded ✓` }) : el("span", { class: "zone warn", text: "Not uploaded yet" }),
        info ? el("div", { class: "muted", style: "font-size: 11px; line-height: 1.35; word-break: break-word;", text: `${info.name} · ${new Date(info.at).toLocaleString("en-GB")} · ${info.by}` }) : null,
        info ? el("button", { type: "button", class: "act ghost small-btn", text: "Lihat hasil upload", onclick: async () => { try { showReport(c.label, await api(`/api/admin/periods/${activePeriod.id}/report/${c.kind}`)); } catch (e) { toast(e.message); } } }) : null);
      const zone = dropZone(finput, `Upload ${c.label}`);
      zone.style.cssText += "min-height:76px;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;padding:16px 12px;box-sizing:border-box;";
      finput.addEventListener("change", async () => {
        if (!finput.files.length) return;
        const fd = new FormData();
        fd.append("file", finput.files[0]);
        fd.append("kind", c.kind);
        statusSpan.textContent = "Uploading & reading…";
        try {
          const res = await upload(`/api/admin/periods/${activePeriod.id}/upload`, fd);
          toast(`${c.label}: ${res.report.ok} of ${res.report.total} rows saved`);
          await adminEnter("periods");
          showReport(`${c.label} · ${activePeriod.name}`, res.report);
        } catch (e) {
          statusSpan.textContent = `Error: ${e.message}`;
          statusSpan.className = "small error";
        }
      });
      grid.append(el("div", { class: "card period-hub-card", style: "padding: 20px; border: 1px solid var(--border); border-radius: 12px; display: flex; flex-direction: column; justify-content: space-between; height: 100%; box-sizing: border-box;" },
        el("div", { style: "flex: 1 1 auto; display: flex; flex-direction: column; min-height: 110px; margin-bottom: 24px;" },
          el("b", { style: "font-size: 14px; font-weight: 700; margin-bottom: 6px; line-height: 1.3;", text: c.label }),
          el("p", { class: "small muted", style: "margin: 0 0 8px 0; line-height: 1.45;", text: c.desc }),
          el("p", { class: "small muted", style: "margin: 0; line-height: 1.4; font-size: 11px;", text: `Columns: ${c.cols}` })),
        el("div", { style: "margin-top: auto; display: flex; flex-direction: column; gap: 8px;" }, zone, statusSpan)));
    });
    uploadGridBox.append(grid);
  }

  // Periods table
  const tbl = el("table");
  clear(periodListBox).append(el("div", { class: "table-wrap" }, tbl));
  table(tbl, [
    { label: "Period Name", get: (p) => el("div", {}, el("b", { text: p.name }), p.active ? el("span", { class: "badge ok", style: "margin-left: 8px;", text: "Active" }) : null) },
    { label: "Category", get: (p) => p.category },
    { label: "Months Elapsed", num: true, get: (p) => `${p.months_elapsed} mo` },
    { label: "PO Rows", num: true, get: (p) => num(p.po_rows_count || 0) },
    { label: "Formulary", num: true, get: (p) => num(p.formulary_count || 0) },
    { label: "Benchmark", num: true, get: (p) => num(p.benchmark_count || 0) },
    { label: "Status", get: (p) => p.active ? el("span", { class: "pill good", text: `Active ${p.category} hub` }) : el("button", {
      type: "button", class: "act ghost small-btn", text: "Activate", onclick: async () => {
        try {
          await send(`/api/admin/periods/${p.id}/activate`, "POST", {});
          toast(`Period ${p.name} activated for ${p.category}. Other categories are not affected.`);
          adminEnter("periods");
        } catch (e) { toast(e.message); }
      }
    }) }
  ], periods);

  // New Period Form: category + Mid-/End-Year + year, e.g. "CnC 2026 End-Year"
  const pCat = el("select", {}, el("option", { value: "CnC", text: "CnC" }), el("option", { value: "Drugs", text: "Drugs" }));
  const pHalf = el("select", {}, el("option", { value: "End-Year", text: "End-Year" }), el("option", { value: "Mid-Year", text: "Mid-Year" }));
  const pYear = el("input", { type: "number", min: "2020", max: "2100", value: String(new Date().getFullYear()) });
  const pMonths = el("input", { type: "number", min: "1", max: "12", value: "12" });
  const preview = el("p", { class: "small" });
  const syncName = () => { preview.textContent = `Period name: ${pCat.value} ${pYear.value} ${pHalf.value}`; };
  [pCat, pHalf, pYear].forEach((x) => x.addEventListener("input", syncName));
  syncName();
  const formMsg = el("p", { class: "error", hidden: "" });
  const pForm = el("form", { class: "progress-form" },
    el("div", { class: "row2" }, adminField("Category", pCat), adminField("Cycle", pHalf)),
    el("div", { class: "row2" }, adminField("Year", pYear), adminField("Running months (bulan berjalan)", pMonths, "Annualized = YTD × 12 ÷ running months")),
    preview,
    formMsg,
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Create period" }))
  );
  pForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    try {
      const made = await send("/api/admin/periods", "POST", { category: pCat.value, half: pHalf.value, year: parseInt(pYear.value, 10), months_elapsed: parseInt(pMonths.value, 10) || 12 });
      toast(`Period "${made.name}" created`);
      adminEnter("periods");
    } catch (err) {
      formMsg.textContent = err.message;
      formMsg.hidden = false;
    }
  });
  clear(newPeriodFormBox).append(pForm);
}

// ---------- principals (add, import, edit) ----------
async function adminPrincipals(body) {
  const listBox = el("div");
  body.append(
    card("Principals",
      "Every principal Siloam negotiates with: Drugs and CnC, renewals, new vendors, addenda and force majeure. Add them one by one or import the template. The negotiation team opens each negotiation from the Negotiation tab.",
      el("div", { class: "actions" },
        act("Add principal", () => principalForm(), "primary"),
        act("Import principals", () => importPrincipals(), "secondary"),
        el("a", { class: "act ghost small-btn", href: "/api/admin/nego/principals-template.xlsx", text: "Excel template ⤓" }),
        el("a", { class: "act ghost small-btn", href: "/api/admin/nego/principals-template.csv", text: "CSV template ⤓" }))),
    card("All principals", null, listBox));
  let d;
  try { d = await api("/api/nego/principals"); } catch (err) { listBox.append(el("p", { class: "error", text: err.message })); return; }
  const tbl = el("table");
  clear(listBox).append(d.principals.length ? el("div", { class: "table-wrap" }, tbl) : el("p", { class: "muted", text: "No principals yet. Add one or import the template." }));
  if (!d.principals.length) return;
  table(tbl, [
    { label: "Principal", get: (p) => el("div", {}, el("b", { text: p.name }), el("div", { class: "small muted", text: p.distributor || "" })) },
    { label: "Category", get: (p) => p.category || "—" },
    { label: "Type", get: (p) => p.nego_type || "Renewal" },
    { label: "Binding", get: (p) => p.binding || "Nett" },
    { label: "Current MOU", get: (p) => (p.mou_end ? `${p.mou_start ? fmtDate(p.mou_start) + " – " : "until "}${fmtDate(p.mou_end)}` : el("span", { class: "muted", text: p.nego_type === "New Vendor" ? "New vendor" : "Not set" })) },
    { label: "Negotiation", get: (p) => (p.cycle ? p.cycle.step_label : el("span", { class: "muted", text: "Not started" })) },
    { label: "", get: (p) => act("Edit", () => principalForm(p), "ghost") },
  ], d.principals);
}

function dropZone(input, label) {
  const zone = el("label", { class: "dropzone" }, input, el("b", { text: label }), el("span", { class: "muted small", text: "Drag a file here or click to choose" }));
  const name = el("span", { class: "filename" });
  zone.append(name);
  input.addEventListener("change", () => { name.textContent = input.files.length ? input.files[0].name : ""; });
  zone.addEventListener("dragover", (e) => { e.preventDefault(); zone.classList.add("over"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("over"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault(); zone.classList.remove("over");
    if (e.dataTransfer.files.length) { input.files = e.dataTransfer.files; input.dispatchEvent(new Event("change")); }
  });
  return zone;
}

// ---------- price data ----------
async function adminData(body) {
  const statusBox = el("div");
  const previewBox = el("div");
  const versionsBox = el("div");
  const historyBox = el("div");
  const kind = el("select", { id: "up-kind" },
    ["price_history", "sku_master", "price_index", "rebate_programs"].map((k) => el("option", { value: k })));
  const file = el("input", { type: "file", accept: ".csv,.xlsx", id: "up-file", class: "sr-only" });
  const msg = el("p", { class: "error", role: "alert", hidden: "" });
  const form = el("form", { class: "upload-form" },
    el("label", { for: "up-kind" }, "What are you uploading?", kind),
    dropZone(file, "Choose a CSV or Excel (.xlsx) file"),
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Check file" }),
      el("span", { class: "muted small", text: "Nothing changes until you apply it." })),
    msg);
  const templates = el("p", { class: "small" }, "Templates: ");
  body.append(
    card("Current data", "What the engine is using right now.", statusBox),
    card("Upload price lists and records", "Purchase orders, contracts, quotes, the SKU master, the price index or rebate programs. The file is checked row by row first; you see a preview and choose to add or replace.",
      form, templates, previewBox),
    card("Versions", "Every apply creates a version. Activate an older one to roll back.", versionsBox),
    card("Recent uploads", null, historyBox));

  let kinds = {};
  try { const u = await api("/api/admin/uploads"); kinds = u.kinds; renderUploads(historyBox, u.uploads); } catch (e) { historyBox.append(el("p", { class: "error", text: e.message })); }
  [...kind.options].forEach((o) => { o.textContent = kinds[o.value] || o.value; });
  Object.entries(kinds).forEach(([k, label], i) => {
    templates.append(i ? ", " : "", el("a", { href: `/api/admin/templates/${k}`, text: label.split(" (")[0] }));
  });
  renderVersions(statusBox, versionsBox);

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    msg.hidden = true;
    if (!file.files.length) { msg.textContent = "Choose a file first."; msg.hidden = false; return; }
    const fd = new FormData();
    fd.append("file", file.files[0]);
    fd.append("kind", kind.value);
    clear(previewBox).append(el("p", { class: "muted", text: "Checking…" }));
    try {
      adminState.staged = await upload("/api/admin/uploads", fd);
      renderPreview(previewBox, adminState.staged, () => adminEnter("data"));
    } catch (ex) { clear(previewBox); msg.textContent = ex.message; msg.hidden = false; }
  });
}

function renderPreview(box, up, done) {
  clear(box);
  const ok = up.status === "staged";
  const facts = [
    ["Rows", num(up.rows)],
    up.date_from ? ["Dates", `${fmtDate(up.date_from)} – ${fmtDate(up.date_to)}`] : null,
    up.vendors != null ? ["Vendors", num(up.vendors)] : null,
    up.skus != null ? ["Products", num(up.skus)] : null,
    up.hospitals != null ? ["Hospitals", num(up.hospitals)] : null,
  ].filter(Boolean);
  const sample = el("table");
  box.append(el("div", { class: `preview ${ok ? "ok" : "bad"}` },
    el("h3", { text: ok ? `Ready to apply: ${up.filename}` : `Fix these before applying: ${up.filename}` }),
    el("div", { class: "facts" }, facts.map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: v })))),
    up.new_vendors && up.new_vendors.length ? el("p", { class: "small", text: `New vendors: ${up.new_vendors.join(", ")}` }) : null,
    up.new_skus && up.new_skus.length ? el("p", { class: "small", text: `New products: ${up.new_skus.join(", ")}` }) : null,
    up.unknown_columns && up.unknown_columns.length ? el("p", { class: "small muted", text: `Ignored columns: ${up.unknown_columns.join(", ")}` }) : null,
    up.errors.length ? el("ul", { class: "errors" }, up.errors.map((x) => el("li", { text: x }))) : null,
    el("details", {}, el("summary", { text: "First rows" }), el("div", { class: "table-wrap" }, sample))));
  const cols = Object.keys(up.sample[0] || {}).slice(0, 9);
  table(sample, cols.map((c) => ({ label: c, get: (r) => r[c] })), up.sample);
  const mode = el("div", { class: "modes", role: "radiogroup", "aria-label": "How to apply" },
    el("label", {}, el("input", { type: "radio", name: "mode", value: "add", checked: "" }), el("b", { text: "Add" }), " to the current data (duplicates are skipped)"),
    el("label", {}, el("input", { type: "radio", name: "mode", value: "replace" }), el("b", { text: "Replace" }), " this file in the current data"));
  const out = el("p", { class: "small", role: "status" });
  if (ok) box.lastChild.append(mode);
  box.lastChild.append(el("div", { class: "actions" },
    ok ? act("Apply", async () => {
      const m = box.querySelector("input[name=mode]:checked").value;
      if (m === "replace" && !confirm("Replace this file in the current data? You can roll back from Versions.")) return;
      out.textContent = "Applying…";
      try {
        const res = await send(`/api/admin/uploads/${up.id}/apply`, "POST", { mode: m });
        toast(`Applied: ${num(res.rows)} rows now in use`);
        await loadCatalog();
        done();
      } catch (ex) { out.textContent = ex.message; }
    }) : null,
    act("Discard", async () => { try { await send(`/api/admin/uploads/${up.id}/discard`, "POST", {}); } catch (_) { /* already gone */ } done(); }, "ghost"),
    out));
}

async function renderVersions(statusBox, versionsBox) {
  let v;
  try { v = await api("/api/admin/versions"); } catch (e) { versionsBox.append(el("p", { class: "error", text: e.message })); return; }
  const s = v.status || {};
  clear(statusBox).append(...[el("div", { class: "facts" },
    ...[["Source", v.active === "sample" ? "Bundled sample" : `Version ${v.active}`],
      ["Rows", num(s.rows)], ["Products", num(s.skus)], ["Vendors", num(s.vendors)], ["Hospitals", num(s.hospitals)],
      ["Dates", s.date_from ? `${fmtDate(s.date_from)} – ${fmtDate(s.date_to)}` : "—"]].map(([l, x]) => el("div", {}, el("span", { text: l }), el("b", { text: x })))),
    s.is_sample ? el("p", { class: "small warn-text", text: "This is still sample data. Replace the price list with your real export to switch it off." }) : null,
    v.env_override ? el("p", { class: "small warn-text", text: "NEGOTIATION_DATA_DIR is set on the server, so it overrides the versions below." }) : null].filter(Boolean));
  const tbl = el("table");
  const rows = [{ id: "sample", created_at: null, user_email: "", note: "Bundled synthetic sample", rows: null, active: v.active === "sample" }, ...v.versions];
  clear(versionsBox).append(el("div", { class: "table-wrap" }, tbl));
  table(tbl, [
    { label: "Version", get: (r) => r.id },
    { label: "When", get: (r) => (r.created_at ? stamp.format(new Date(r.created_at)) : "—") },
    { label: "By", get: (r) => r.user_email || "—" },
    { label: "What", get: (r) => r.note || "" },
    { label: "Rows", num: true, get: (r) => (r.rows == null ? "—" : num(r.rows)) },
    { label: "", get: (r) => r.active ? el("span", { class: "zone good", text: "In use" })
      : el("button", { type: "button", class: "act ghost", text: "Activate", onclick: async (e) => {
        e.stopPropagation();
        if (!confirm(`Switch the engine to ${r.id === "sample" ? "the sample data" : `version ${r.id}`}?`)) return;
        try { await send(`/api/admin/versions/${encodeURIComponent(r.id)}/activate`, "POST", {}); await loadCatalog(); toast("Version activated"); adminEnter("data"); }
        catch (ex) { toast(ex.message); }
      } }) },
  ], rows);
}

function renderUploads(box, uploads) {
  if (!uploads.length) { box.append(el("p", { class: "muted small", text: "No uploads yet." })); return; }
  const tbl = el("table");
  box.append(el("div", { class: "table-wrap" }, tbl));
  table(tbl, [
    { label: "When", get: (r) => stamp.format(new Date(r.at)) },
    { label: "File", get: (r) => r.filename },
    { label: "Type", get: (r) => r.kind.replace("_", " ") },
    { label: "Rows", num: true, get: (r) => num(r.rows) },
    { label: "Status", get: (r) => el("span", { class: `zone ${r.status === "applied" ? "good" : r.status === "invalid" ? "bad" : "warn"}`, text: r.status }) },
    { label: "By", get: (r) => r.user_email },
  ], uploads);
}

// ---------- documents ----------
// Siloam's Word templates for the draft BAK (per scheme) and the PKS. Placeholders such as {{principal}} are filled in.
async function docTemplatesCard() {
  const box = el("div");
  const wrap = card("Template BAK dan PKS (Word)", "Upload template .docx Siloam. Tanpa upload, app memakai draft bawaan. Nett memakai template Nett; Disc dan Nett to Disc memakai template Disc.", box);
  let d;
  try { d = await api("/api/admin/doc-templates"); } catch (e) { box.append(el("p", { class: "error", text: e.message })); return wrap; }
  for (const t of d.templates) {
    const file = el("input", { type: "file", accept: ".docx", class: "sr-only", id: `tpl-${t.key}` });
    file.addEventListener("change", async () => {
      if (!file.files.length) return;
      const fd = new FormData(); fd.append("file", file.files[0]);
      try { await upload(`/api/admin/doc-templates/${t.key}`, fd); toast(`${t.label} tersimpan`); adminEnter("documents"); } catch (e) { toast(e.message); }
    });
    box.append(el("div", { class: "actions", style: "align-items:center" }, el("b", { text: t.label }),
      t.file ? el("span", { class: "small muted", text: `${t.file.name} · ${new Date(t.file.at).toLocaleString("en-GB")}` }) : el("span", { class: "small muted", text: "pakai draft bawaan" }),
      el("label", { class: "act secondary small-btn", for: `tpl-${t.key}`, text: t.file ? "Ganti" : "Upload .docx" }), file,
      t.file ? act("Hapus", async () => { await send(`/api/admin/doc-templates/${t.key}`, "DELETE"); adminEnter("documents"); }, "ghost") : null));
  }
  box.append(el("p", { class: "small muted", text: `Placeholder: ${d.placeholders.map((x) => `{{${x}}}`).join(" ")}. {{tabel_harga}} diganti daftar harga per SKU.` }));
  return wrap;
}

async function adminDocs(body) {
  body.append(await docTemplatesCard());
  const c = shell.catalog;
  const file = el("input", { type: "file", id: "doc-file", class: "sr-only", accept: ".pdf,.docx,.xlsx,.csv,.png,.jpg,.jpeg" });
  const vendor = el("select", { id: "doc-vendor" }, el("option", { value: "", text: "No vendor" }), c.vendors.map((v) => el("option", { value: v, text: v })));
  const sku = el("select", { id: "doc-sku" }, el("option", { value: "", text: "No product" }), c.skus.map((s) => el("option", { value: s.sku, text: s.sku_name })));
  const note = el("input", { id: "doc-note", placeholder: "e.g. Signed 2025 contract", maxlength: 300 });
  const msg = el("p", { class: "small", role: "status" });
  const listBox = el("div");
  const search = el("input", { placeholder: "Search documents", "aria-label": "Search documents" });
  const form = el("form", { class: "upload-form" },
    dropZone(file, "Choose a document (PDF, Word, Excel, CSV or image, up to 20 MB)"),
    el("div", { class: "row3" },
      el("label", { for: "doc-vendor" }, "Vendor", vendor), el("label", { for: "doc-sku" }, "Product", sku),
      el("label", { for: "doc-note" }, "Note", note)),
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Upload" }), msg));
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!file.files.length) { msg.textContent = "Choose a file first."; return; }
    const fd = new FormData();
    fd.append("file", file.files[0]); fd.append("vendor", vendor.value); fd.append("sku", sku.value); fd.append("note", note.value);
    msg.textContent = "Uploading…";
    try { await upload("/api/admin/documents", fd); toast("Document uploaded"); adminEnter("documents"); }
    catch (ex) { msg.textContent = ex.message; }
  });
  body.append(
    card("Upload a document", "Contracts, quotes, price letters. Link it to a vendor or product so negotiators find it in the renewal and vendor views. Everyone signed in can open documents.", form),
    card("Documents", null, search, listBox));
  const load = async () => {
    let docs = [];
    try { docs = await api(`/api/documents${search.value ? `?text=${encodeURIComponent(search.value)}` : ""}`); } catch (e) { clear(listBox).append(el("p", { class: "error", text: e.message })); return; }
    const tbl = el("table");
    clear(listBox).append(docs.length ? el("div", { class: "table-wrap" }, tbl) : el("p", { class: "muted small", text: "No documents yet." }));
    if (!docs.length) return;
    table(tbl, [
      { label: "File", get: (d) => el("a", { href: `/api/documents/${encodeURIComponent(d.id)}/download`, text: d.filename }) },
      { label: "Vendor", get: (d) => d.vendor || "—" },
      { label: "Product", get: (d) => d.sku || "—" },
      { label: "Note", get: (d) => d.note || "" },
      { label: "Size", num: true, get: (d) => `${num(d.size / 1024)} KB` },
      { label: "Added", get: (d) => `${stamp.format(new Date(d.at))} · ${d.user_email.split("@")[0]}` },
      { label: "", get: (d) => el("button", { type: "button", class: "act ghost", text: "Delete", onclick: async () => {
        if (!confirm(`Delete ${d.filename}?`)) return;
        try { await send(`/api/admin/documents/${encodeURIComponent(d.id)}`, "DELETE"); toast("Deleted"); load(); } catch (ex) { toast(ex.message); }
      } }) },
    ], docs);
  };
  let t;
  search.addEventListener("input", () => { clearTimeout(t); t = setTimeout(load, 250); });
  load();
}

// ---------- settings ----------
async function adminSettings(body) {
  let d;
  try { d = await api("/api/settings"); } catch (e) { body.append(el("p", { class: "error", text: e.message })); return; }
  const groups = {};
  d.settings.forEach((s) => { (groups[s.group] = groups[s.group] || []).push(s); });
  const inputs = {};
  const form = el("form", { class: "settings-form" });
  const show = (s, v) => (s.kind === "fraction" && v != null ? String(+(v * 100).toFixed(4)) : v == null ? "" : String(v));
  Object.entries(groups).forEach(([g, items]) => {
    const fs = el("fieldset", {}, el("legend", { text: g }));
    items.forEach((s) => {
      let input;
      if (s.kind === "bool") {
        input = el("input", { type: "checkbox", id: `set-${s.key}` });
        input.checked = !!s.value;
      } else {
        input = el("input", { id: `set-${s.key}`, inputmode: "decimal", value: show(s, s.value) || null,
          placeholder: s.optional ? "not set" : show(s, s.default) });
      }
      inputs[s.key] = { s, input };
      const reset = el("button", { type: "button", class: "act ghost", text: "Default", title: `Reset to ${show(s, s.default) || "not set"}`,
        onclick: () => { if (s.kind === "bool") input.checked = !!s.default; else input.value = show(s, s.default); } });
      fs.append(el("div", { class: "setting" },
        el("label", { for: input.id }, el("b", { text: s.label }), s.help ? el("span", { class: "muted small", text: s.help }) : null),
        el("div", { class: "inline" }, input, s.kind === "fraction" ? el("span", { class: "unit", text: "%" }) : null, reset)));
    });
    form.append(fs);
  });
  const msg = el("p", { class: "small", role: "status" });
  form.append(el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Save settings" }), msg));
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const changes = {};
    for (const [k, { s, input }] of Object.entries(inputs)) {
      let v;
      if (s.kind === "bool") v = input.checked;
      else if (input.value.trim() === "") v = s.optional ? null : s.default;
      else {
        v = parseFloat(input.value.replace(",", "."));
        if (!Number.isFinite(v)) { msg.textContent = `"${s.label}" must be a number`; input.focus(); return; }
        if (s.kind === "fraction") v = v / 100;
      }
      if (v !== s.value) changes[k] = v;
    }
    if (!Object.keys(changes).length) { msg.textContent = "Nothing changed."; return; }
    msg.textContent = "Saving…";
    try {
      const res = await send("/api/admin/settings", "PUT", { changes });
      toast(`Saved ${res.changed.length} setting${res.changed.length === 1 ? "" : "s"}`);
      await loadCatalog();
      adminEnter("settings");
    } catch (ex) { msg.textContent = ex.message; }
  });
  body.append(card("Engine settings", "These drive every target, plan and alert, in the app and for Claude via the MCP server. Changes apply immediately and are logged.", form),
    card("About the assistant", null,
      el("p", { class: "small", text: "The chat understands typed questions with a built-in parser that runs on this server; nothing is sent outside. The engine computes every number." }),
      el("p", { class: "small muted", text: "Sending unclear questions to Claude is planned. It needs an Anthropic API key on the server, and the question text and engine results would be sent to Anthropic. The switch above records your choice; it has no effect until that's built." })));
}

// ---------- users ----------
async function adminUsers(body) {
  const email = el("input", { id: "u-email", type: "email", placeholder: "name@siloamhospitals.com", required: "" });
  const name = el("input", { id: "u-name", placeholder: "Full name", required: "" });
  const role = el("select", { id: "u-role" }, el("option", { value: "viewer", text: "Viewer: sees everything, updates renewals" }), el("option", { value: "negotiator", text: "Negotiator: runs negotiations (prices, counter offers, anomalies, principal links)" }),
    el("option", { value: "admin", text: "Admin: also uploads data, tunes the engine, manages users" }));
  const msg = el("p", { class: "small", role: "status" });
  const form = el("form", { class: "row3" },
    el("label", { for: "u-email" }, "Email", email), el("label", { for: "u-name" }, "Name", name), el("label", { for: "u-role" }, "Role", role),
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Invite" }), msg));
  const listBox = el("div");
  const linkBox = el("div", { id: "setup-link-box" });
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    try {
      const u = await send("/api/admin/users", "POST", { email: email.value, name: name.value, role: role.value });
      await adminEnter("users");
      if (u.setup_link) showSetupLink(u.email, u.setup_link);
    } catch (ex) { msg.textContent = ex.message; }
  });
  body.append(card("Invite someone", "Only invited people can sign in. Inviting gives you a one-time link: send it to them (email or WhatsApp) so they set their own password. You never see their password.", form),
    linkBox, card("Users", null, listBox));
  let users;
  try { users = (await api("/api/admin/users")).users; } catch (e) { listBox.append(el("p", { class: "error", text: e.message })); return; }
  const tbl = el("table");
  listBox.append(el("div", { class: "table-wrap" }, tbl));
  const patch = async (u, change, label) => {
    try { await send(`/api/admin/users/${u.id}`, "PATCH", change); toast(label); adminEnter("users"); } catch (ex) { toast(ex.message); adminEnter("users"); }
  };
  table(tbl, [
    { label: "Name", get: (u) => u.name },
    { label: "Email", get: (u) => u.email },
    { label: "Role", get: (u) => {
      const s = el("select", { "aria-label": `Role for ${u.name}` }, el("option", { value: "viewer", text: "Viewer" }), el("option", { value: "negotiator", text: "Negotiator" }), el("option", { value: "admin", text: "Admin" }));
      s.value = u.role;
      s.addEventListener("change", () => patch(u, { role: s.value }, `${u.name} is now ${s.value}`));
      return s;
    } },
    { label: "Status", get: (u) => el("span", { class: `zone ${u.active ? "good" : "bad"}`, text: u.active ? "Active" : "Disabled" }) },
    { label: "Password", get: (u) => el("span", { class: `zone ${u.has_password ? "good" : "warn"}`, text: u.has_password ? "Set" : "Not set yet" }) },
    { label: "Last sign-in", get: (u) => (u.last_login ? stamp.format(new Date(u.last_login)) : "never") },
    { label: "", get: (u) => el("div", { class: "actions" },
      u.active ? el("button", { type: "button", class: "act ghost", text: u.has_password ? "Reset link" : "Setup link",
        onclick: async () => {
          try { const r = await send(`/api/admin/users/${u.id}/setup-link`, "POST", {}); showSetupLink(r.email, r.setup_link); }
          catch (ex) { toast(ex.message); }
        } }) : null,
      el("button", { type: "button", class: "act ghost", text: u.active ? "Disable" : "Enable",
        onclick: () => patch(u, { active: !u.active }, `${u.name} ${u.active ? "disabled" : "enabled"}`) })) },
  ], users);
}

function showSetupLink(email, link) {
  const box = $("setup-link-box");
  if (!box) return;
  const input = el("input", { value: link, readonly: "", "aria-label": "Setup link", class: "linkbox" });
  const copy = el("button", { type: "button", class: "act", text: "Copy link",
    onclick: async () => { try { await navigator.clipboard.writeText(link); toast("Link copied"); } catch (_) { input.select(); } } });
  clear(box).append(card(`Password link for ${email}`,
    "Send this to them. It works once and expires in 7 days. Opening it lets them choose their password and signs them in. Making a new link cancels the old one.",
    el("div", { class: "row-link" }, input, copy)));
  box.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ---------- activity ----------
function describeEvent(e) {
  const d = e.detail || {};
  switch (e.action) {
    case "signin": return "signed in";
    case "signin.failed": return "failed to sign in";
    case "user.password_set": return `set a password for ${d.email}`;
    case "user.setup_link": return `made a password link for ${d.email}`;
    case "user.invite": return `invited ${d.email} as ${d.role}`;
    case "user.update": return `updated ${d.email}${d.role ? ` (role: ${d.role})` : ""}${d.active === 0 ? " (disabled)" : d.active === 1 ? " (enabled)" : ""}`;
    case "upload.stage": return `checked ${d.file} (${num(d.rows)} rows, ${d.status})`;
    case "upload.apply": return `applied an upload (${d.mode}); ${num(d.rows_now)} rows now in use`;
    case "upload.discard": return "discarded an upload";
    case "version.activate": return `switched the data to ${d.version === "sample" ? "the sample" : `version ${d.version}`}`;
    case "document.upload": return `uploaded ${d.file}${d.vendor ? ` for ${d.vendor}` : ""}`;
    case "document.delete": return `deleted ${d.file}`;
    case "settings.update": return `changed settings: ${Object.entries(d).map(([k, v]) => `${k} ${v.from ?? "unset"} → ${v.to ?? "unset"}`).join(", ")}`;
    default: return e.action;
  }
}

async function adminActivity(body) {
  let d;
  try { d = await api("/api/admin/activity"); } catch (e) { body.append(el("p", { class: "error", text: e.message })); return; }
  body.append(card("Activity", "Who did what, newest first.",
    d.events.length ? el("ol", { class: "timeline" }, d.events.map((e) => el("li", {},
      el("b", { text: e.user_email.split("@")[0] }), ` ${describeEvent(e)}`,
      el("span", { class: "muted small", text: ` · ${stamp.format(new Date(e.at))}` }))))
      : el("p", { class: "muted", text: "Nothing yet." })));
}


// ---------- report: one row per principal, results at Feedback I and at FIN ----------
async function adminReport(body) {
  const status = el("select", {}, [["all", "Semua status"], ["open", "Open"], ["closed", "Selesai"]].map(([v, l]) => el("option", { value: v, text: l })));
  const cat = el("select", {}, [["all", "Semua kategori"], ["Drugs", "Drugs"], ["CnC", "CnC"]].map(([v, l]) => el("option", { value: v, text: l })));
  const per = el("select", {}, el("option", { value: "", text: "Semua periode" }));
  const out = el("div");
  const xls = el("a", { class: "act secondary", href: "/api/admin/report.xlsx", text: "Export Excel ⤓" });
  body.append(card("Report: ringkasan semua principal", "Satu baris per negosiasi. Dua tabel: hasil sampai Feedback I dan hasil FIN. Angka mengikuti Pre-Nego dan Post-Nego di tiap negosiasi.",
    el("div", { class: "filters board-filters" }, el("label", {}, "Status", status), el("label", {}, "Kategori", cat), el("label", {}, "Periode", per), xls), out));
  const qs = () => { const p = new URLSearchParams({ status: status.value, category: cat.value }); if (per.value) p.set("period", per.value); return p.toString(); };
  let first = true;
  async function load() {
    xls.href = `/api/admin/report.xlsx?${qs()}`;
    clear(out).append(el("p", { class: "muted", text: "Loading…" }));
    let d;
    try { d = await api(`/api/admin/report?${qs()}`); } catch (e) { clear(out).append(el("p", { class: "error", text: e.message })); return; }
    if (first) { d.periods.forEach((p) => per.append(el("option", { value: p.id, text: p.name }))); first = false; }
    const fmt = (c, v) => (v == null || v === "" ? "—" : c.kind === "pct" ? pct(v, 2) : c.kind === "money" ? (v > 0 && /impact|cost_impact/.test(c.key) ? "+" : "") + money(v) : c.kind === "num" ? num(v) : String(v));
    const tableFor = (title, rows) => {
      const t = el("table", { class: "items" });
      table(t, d.columns.map((c) => ({ label: c.label, num: c.kind !== "text", get: (r) => fmt(c, r[c.key]) })), rows, (r) => { location.hash = `#nego/${r.cycle_id}/analysis`; });
      return el("div", {}, el("h3", { text: `${title} (${rows.length} principal)` }), rows.length ? el("div", { class: "table-wrap", style: "overflow:auto" }, t) : el("p", { class: "muted", text: "Tidak ada data." }));
    };
    clear(out).append(tableFor("Hasil FB1", d.fb1), tableFor("Hasil FIN", d.fin));
  }
  [status, cat, per].forEach((x) => x.addEventListener("change", load));
  load();
}
