"use strict";
// Principal negotiations: the principals list (MOU alerts, which step each cycle is at)
// and one cycle's workspace (step tracker, items table laid out like Template_Nego,
// the anomaly queue, files and activity). Everyone signed in can look; only admins see
// the buttons that change things, and the server enforces that on every call.

const ng = { list: null, cid: null, ov: null, tab: "items", page: { q: "", filter: "all", status: "all", brand: "all", sort: "sort", sort_col: "", sort_dir: "asc", offset: 0, limit: 50 }, findStatus: "open", rule: "" };
const isAdmin = () => shell.user && shell.user.role === "admin";
const canWork = () => shell.user && (shell.user.role === "admin" || shell.user.role === "negotiator");
const WHO = { admin: "Siloam", principal: "Principal", reference: "Reference" };
const STEP_DESC = {
  prepare: "Siloam builds the item list from the PO year, the formulary and the current MOU.",
  review_rfq: "Optional: Siloam marks the anomalous SKUs and sends them back to the principal to check. With no anomaly, go straight to the counter offer.",
  identification: "The principal confirms brand, catalogue no. and status for every item.",
  current_mou: "Current MOU prices, shown to the principal for reference only.",
  rfq: "The principal quotes Qty/PO unit, HNA and discount for every active item.",
  counter_offer: "Siloam sends a counter-offer discount per item. The engine suggests where to hold the MOU price.",
  feedback1: "The principal accepts the counter offer or proposes a Feedback I discount.",
  online_nego: "Siloam records the discounts agreed in the online meeting.",
  submission: "The principal submits company documents; Siloam checks and files the BAK data.",
  closed: "Selesai. Harga sepakat masuk ke BAK dan PKS.",
};
const FILTERS = [
  ["all", "All items"],
  ["active", "Active items only"],
  ["discontinued", "Discontinued only"],
  ["new", "New items (cek Siloam)"],
  ["anomalies", "With open findings"],
  ["increase", "Price up vs acuan (+%)"],
  ["decrease", "Price down vs acuan (-%)"],
  ["has_rfq", "With RFQ price quoted"],
  ["no_rfq", "No RFQ price yet"],
  ["no_mou", "Not in current MOU"],
  ["has_co", "With Counter Offer"],
  ["has_fb1", "With Feedback I"],
  ["has_on", "With Online Nego"],
];
const SORTS = [
  ["sort", "Template order (Default)"],
  ["name_asc", "Item Name (A → Z)"],
  ["name_desc", "Item Name (Z → A)"],
  ["erp_asc", "ERP Code (Ascending)"],
  ["erp_desc", "ERP Code (Descending)"],
  ["impact", "Biggest cost impact (Impact/yr)"],
  ["change", "Biggest price increase (+%)"],
  ["change_down", "Biggest price decrease (-%)"],
  ["spend", "Biggest PO spend (+PPN)"],
  ["rfq_price_desc", "Highest RFQ price/pc"],
  ["rfq_price_asc", "Lowest RFQ price/pc"],
  ["mou_price_desc", "Highest MOU price/pc"],
  ["mou_price_asc", "Lowest MOU price/pc"],
  ["co_disc_desc", "Highest Counter Offer %"],
  ["fb1_disc_desc", "Highest Feedback I %"],
  ["on_disc_desc", "Highest Online Nego %"],
  ["margin_pc_desc", "Highest Margin/pc"],
  ["margin_pc_asc", "Lowest Margin/pc"],
  ["margin_form_desc", "Biggest Margin Impact (vs Formularium)"],
  ["margin_po_desc", "Biggest Margin Impact (vs PO)"],
  ["findings", "Most open findings"],
];

function ngEnter(rest) {
  wireBoard();
  const cid = rest[0] ? parseInt(rest[0], 10) : null;
  if (cid) { ngCycle(cid, rest[1] || "items"); return; }
  ngList();
}

// ---------- principals list ----------
async function ngList() {
  ng.cid = null;
  ng.viewMode = ng.viewMode || "list";
  const page = clear($("nego"));
  const tiles = el("div", { class: "kpis5" });
  const tools = el("div", { class: "ng-tools" });
  const search = el("input", { id: "ng-search", placeholder: "Principal or distributor", type: "search" });
  const attention = el("input", { type: "checkbox", id: "ng-attn" });
  const box = el("div", { class: ng.viewMode === "board" ? "board-wrap card ng-board" : "table-wrap card ng-table" });

  const viewSwitcher = el("div", { class: "view-switcher" },
    el("button", {
      type: "button",
      class: `view-btn ${ng.viewMode === "list" ? "active" : ""}`,
      text: "☷ List View",
      onclick: () => { ng.viewMode = "list"; ngList(); }
    }),
    el("button", {
      type: "button",
      class: `view-btn ${ng.viewMode === "board" ? "active" : ""}`,
      text: "☶ Board View",
      onclick: () => { ng.viewMode = "board"; ngList(); }
    })
  );

  page.append(
    el("div", { class: "page-head ng-head" },
      el("div", {}, el("h1", { text: "Principal negotiations" }),
        el("p", { class: "muted", text: "One negotiation per principal with all of its items, step by step as in Template_Nego." })),
      tools),
    tiles,
    el("div", { class: "filters board-filters" },
      el("label", {}, "Search", search),
      el("label", { class: "check" }, attention, "Needs attention"),
      ng.filter6m ? el("span", { class: "pill alert", style: "cursor: pointer;", text: "Filtered: MOU ends in 6m ✕", onclick: () => { ng.filter6m = false; ngList(); } }) : null,
      ng.filterFindings ? el("span", { class: "pill medium", style: "cursor: pointer;", text: "Filtered: Has open findings ✕", onclick: () => { ng.filterFindings = false; ngList(); } }) : null,
      ng.filterOverdue ? el("span", { class: "pill alert", style: "cursor: pointer;", text: "Filtered: lewat batas waktu ✕", onclick: () => { ng.filterOverdue = false; ngList(); } }) : null),
    box);

  tools.append(viewSwitcher);
  // Adding and importing principals lives in Admin → Principals; starting a negotiation stays here.
  if (isAdmin()) tools.append(el("a", { class: "act ghost", href: "#admin/principals", text: "Manage principals in Admin →" }));
  box.append(el("p", { class: "muted pad", text: "Loading…" }));
  let d;
  try { d = ng.list = await api("/api/nego/principals"); } catch (e) { clear(box).append(el("p", { class: "error pad", text: e.message })); return; }
  const ps = d.principals;
  const alerts = ps.filter((p) => p.mou_alert);
  const open = ps.filter((p) => p.cycle && p.cycle.status === "open");

  const m6Total = d.kpi_mou_6m ? d.kpi_mou_6m.total : alerts.length;
  const m6Open = d.kpi_mou_6m ? d.kpi_mou_6m.open : open.filter(p => p.mou_alert).length;
  const m6Unopened = d.kpi_mou_6m ? d.kpi_mou_6m.unopened : (m6Total - m6Open);
  const t2 = tile(`MOU ends within ${d.alert_months || 6} months`, num(m6Total), `${m6Open} with negotiation open · ${m6Unopened} unopened`);
  t2.classList.add("clickable-tile");
  if (ng.filter6m) t2.classList.add("tile-active");
  t2.title = "Click to filter principals whose MOU ends within 6 months";
  t2.addEventListener("click", () => {
    ng.filter6m = !ng.filter6m;
    ng.filterFindings = false;
    ng.filterOverdue = false;
    ngList();
  });

  const findingsPrincipals = d.principals_with_findings ?? ps.filter(p => p.cycle && p.cycle.open_anomalies > 0).length;
  const t4 = tile("Open findings", num(findingsPrincipals), "principals with open findings to review");
  t4.classList.add("clickable-tile");
  if (ng.filterFindings) t4.classList.add("tile-active");
  t4.title = "Click to filter principals with open findings";
  t4.addEventListener("click", () => {
    ng.filterFindings = !ng.filterFindings;
    ng.filter6m = false;
    ng.filterOverdue = false;
    ngList();
  });

  const overdueN = d.overdue_count ?? ps.filter((p) => p.overdue_days > 0).length;
  const t5 = tile("Lewat batas waktu", num(overdueN), overdueN ? "principal belum mengirim, klik untuk filter" : "tidak ada yang terlambat");
  t5.classList.add("clickable-tile");
  if (ng.filterOverdue) t5.classList.add("tile-active");
  if (overdueN) t5.classList.add("alert");
  t5.title = "Klik untuk menampilkan negosiasi yang lewat batas waktu";
  t5.addEventListener("click", () => { ng.filterOverdue = !ng.filterOverdue; ng.filter6m = false; ng.filterFindings = false; ngList(); });

  clear(tiles).append(
    tile("Principals", num(new Set(ps.map((p) => p.id)).size)),
    t2,
    tile("Negotiations open", num(open.length)),
    t4,
    t5);
  if (m6Total > 0) tiles.children[1].classList.add("alert");

  if (!ps.length) {
    clear(box).append(el("div", { class: "empty-state" },
      el("h3", { text: "No principals yet" }),
      el("p", { class: "muted", text: isAdmin() ? "Add or import principals in Admin → Principals, or load the fictional sample to try the workflow." : "An admin adds principals in Admin. You open the negotiations here." }),
      isAdmin() ? el("div", { class: "actions" }, act("Load sample principals", loadSample), el("a", { class: "act secondary", href: "#admin/principals", text: "Go to Admin → Principals" })) : null));
    return;
  }

  let pSortCol = "name", pSortDir = "asc";
  const render = () => {
    const q = search.value.trim().toLowerCase();
    let rows = ps.filter((p) => {
      if (q && !`${p.name} ${p.distributor || ""}`.toLowerCase().includes(q)) return false;
      if (attention.checked && !p.mou_alert && !(p.cycle && p.cycle.status === "open" && p.cycle.open_anomalies)) return false;
      if (ng.filter6m && !(p.mou_days_left != null && p.mou_days_left <= 183 && p.mou_days_left >= 0) && !p.mou_alert) return false;
      if (ng.filterFindings && !(p.cycle && p.cycle.open_anomalies > 0)) return false;
      if (ng.filterOverdue && !(p.overdue_days > 0)) return false;
      return true;
    });

    if (ng.viewMode === "board") {
      renderNegoBoard(box, rows, d);
      return;
    }

    const dir = pSortDir === "desc" ? -1 : 1;
    rows.sort((a, b) => {
      let va = "", vb = "";
      if (pSortCol === "name") { va = a.name || ""; vb = b.name || ""; }
      else if (pSortCol === "category") { va = a.category || ""; vb = b.category || ""; }
      else if (pSortCol === "mou") { const da = a.mou_days_left ?? 999999; const db = b.mou_days_left ?? 999999; return (da - db) * dir; }
      else if (pSortCol === "items") { const ia = a.cycle ? a.cycle.items : -1; const ib = b.cycle ? b.cycle.items : -1; return (ia - ib) * dir; }
      else if (pSortCol === "anomalies") { const oa = a.cycle ? (a.cycle.open_anomalies || 0) : -1; const ob = b.cycle ? (b.cycle.open_anomalies || 0) : -1; return (oa - ob) * dir; }
      else if (pSortCol === "stage") { va = a.cycle ? a.cycle.current_step : ""; vb = b.cycle ? b.cycle.current_step : ""; }
      return String(va).localeCompare(String(vb)) * dir;
    });

    const t = el("table", { class: "ng-principals" });
    const colsDef = [
      { id: "name", label: "Principal", get: (p) => el("div", {}, el("b", { text: p.name }), p.category && p.category.includes("&") ? el("span", { class: "pill", style: "margin-left:6px", text: p.row_category }) : null, el("div", { class: "small muted", text: p.distributor || "" })) },
      { id: "category", label: "Category", get: (p) => p.category || "—" },
      { id: "mou", label: "MOU ends", get: (p) => mouCell(p) },
      { id: "stage", label: "Negotiation", get: (p) => cycleCell(p) },
      { id: "items", label: "Items", num: true, get: (p) => (p.cycle ? num(p.cycle.items) : "—") },
      { id: "anomalies", label: "Open findings", num: true, get: (p) => (p.cycle && p.cycle.open_anomalies ? el("span", { class: "pill medium", text: num(p.cycle.open_anomalies) }) : "—") },
      { id: "contact", label: "Contact", get: (p) => el("div", { class: "small" }, p.contact_name || "", el("div", { class: "muted", text: [p.contact_email, p.contact_phone].filter(Boolean).join(" · ") })) },
    ];
    table(t, colsDef.map((c) => ({
      ...c,
      label: `${c.label} ${pSortCol === c.id ? (pSortDir === "asc" ? "▲" : "▼") : ""}`,
    })), rows, (p) => { if (p.cycle) location.hash = `#nego/${p.cycle.id}`; else if (canWork()) startCycleForm(p); else principalInfo(p); });

    const ths = t.querySelectorAll("th");
    colsDef.forEach((c, idx) => {
      const th = ths[idx];
      if (th && c.id !== "contact") {
        th.style.cursor = "pointer";
        th.title = `Click to sort by ${c.label}`;
        th.addEventListener("click", () => {
          if (pSortCol === c.id) pSortDir = pSortDir === "asc" ? "desc" : "asc";
          else { pSortCol = c.id; pSortDir = "asc"; }
          render();
        });
      }
    });

    clear(box).append(rows.length ? t : el("p", { class: "muted pad", text: "No principals match." }));
  };
  search.addEventListener("input", render);
  attention.addEventListener("change", render);
  render();
}

function renderNegoBoard(container, ps, d) {
  clear(container);
  const stages = [
    { key: "not_started", label: "Not Started" },
    { key: "preparing", label: "Preparing" },
    { key: "with_principal", label: "With Principal" },
    { key: "negotiating", label: "In Negotiation" },
    { key: "renewed", label: "Agreed / Renewed" },
  ];

  const stageOf = (p) => {
    if (!p.cycle) return "not_started";
    if (p.cycle.status === "closed" || p.cycle.current_step === "closed") return "renewed";
    if (["rfq", "feedback1", "submission"].includes(p.cycle.current_step)) return "with_principal";
    if (p.cycle.current_step === "prepare") return "preparing";
    return "negotiating";
  };

  const boardGrid = el("div", { class: "board", style: "display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 14px; width: 100%;" });

  stages.forEach(st => {
    const inCol = ps.filter(p => stageOf(p) === st.key);
    const colEl = el("section", { class: `col s-${st.key}` },
      el("header", {}, el("b", { text: st.label }), el("span", { class: "count", text: String(inCol.length) }))
    );
    const cardList = el("ul", { class: "cards" });
    if (!inCol.length) {
      cardList.append(el("li", { class: "empty-col", text: "None" }));
    } else {
      inCol.forEach(p => {
        const card = el("li", { class: `bcard ${p.mou_alert ? "urgent" : ""}`, tabindex: 0 });
        // Element.append prints null as the word "null", so empty slots are dropped.
        card.append(...[
          el("div", { class: "b-top" },
            el("span", { class: `days ${p.mou_days_left != null && p.mou_days_left <= 90 ? "soon" : ""}` },
              p.mou_end ? `${fmtDate(p.mou_end)} (${p.mou_days_left != null ? (p.mou_days_left < 0 ? `ended ${-p.mou_days_left}d ago` : `${p.mou_days_left}d left`) : ""})` : "MOU not set"
            )
          ),
          el("b", { class: "b-name", text: p.name }),
          el("span", { class: "meta", text: [p.row_category || p.category, p.distributor ? `via ${p.distributor}` : ""].filter(Boolean).join(" · ") }),
          p.cycle ? el("span", { class: "next", text: `Step: ${p.cycle.step_label}` }) : null,
          p.cycle && p.cycle.open_anomalies ? el("span", { class: "pill medium", style: "margin-top: 6px; width: fit-content;", text: `${p.cycle.open_anomalies} open findings` }) : null
        ].filter(Boolean));
        card.addEventListener("click", () => {
          if (p.cycle) location.hash = `#nego/${p.cycle.id}`;
          else if (canWork()) startCycleForm(p);
          else principalInfo(p);
        });
        cardList.append(card);
      });
    }
    colEl.append(cardList);
    boardGrid.append(colEl);
  });
  container.append(boardGrid);
}

function mouCell(p) {
  if (!p.mou_end) return el("span", { class: "muted", text: "Not set" });
  const d = p.mou_days_left;
  const when = d < 0 ? `ended ${num(-d)} days ago` : `in ${num(d)} days`;
  return el("div", { class: p.mou_alert ? "mou alert" : "mou" },
    p.mou_alert ? el("span", { class: "dot", "aria-hidden": "true" }) : null,
    el("span", { text: fmtDate(p.mou_end) }), el("div", { class: "small muted", text: when }),
    p.mou_alert ? el("div", { class: "small warn-text", text: "Start the negotiation" }) : null);
}

function cycleCell(p) {
  const c = p.cycle;
  if (!c) return canWork() ? el("button", { type: "button", class: "act small-btn", text: "Start negotiation", onclick: (e) => { e.stopPropagation(); startCycleForm(p); } }) : el("span", { class: "muted", text: "Not started" });
  const keys = ng.list.steps.map((s) => s.key);
  const at = keys.indexOf(c.current_step);
  return el("div", { class: "cyc" },
    el("span", { class: `stage ${c.status === "closed" ? "s-agreed" : "s-negotiating"}`, text: c.step_label }),
    c.submitted_steps && c.submitted_steps[c.current_step] ? el("span", { class: "sent-pill", text: "Principal sent ✓" }) : null,
    p.overdue_days ? el("span", { class: "late-pill", text: `Overdue ${p.overdue_days} day${p.overdue_days === 1 ? "" : "s"}` }) :
      c.step_due && !(c.submitted_steps || {})[c.current_step] ? el("span", { class: "small muted", text: `Due ${fmtDate(c.step_due)}` }) : null,
    el("div", { class: "mini-steps", "aria-label": `Step ${at + 1} of ${keys.length}` },
      keys.slice(0, -1).map((k, i) => el("i", { class: i < at ? "done" : i === at ? "now" : "" }))),
    el("div", { class: "small muted", text: `${fmtDate(c.contract_start)} – ${fmtDate(c.contract_end)}` }));
}

function principalInfo(p) {
  openPanel(el("p", { class: "eyebrow", text: "Principal" }), el("h2", { id: "drawer-title", text: p.name }),
    el("p", { class: "muted", text: "No negotiation yet. The Siloam negotiation team starts one from this list." }));
}

function openPanel(...kids) {
  closeDrawer.back = document.activeElement;
  const body = clear($("drawer-body"));
  body.append(...kids);
  $("drawer").hidden = false;
  $("drawer-scrim").hidden = false;
  $("drawer-close").focus();
  return body;
}

function field(label, input, hint) {
  return el("label", {}, label, input, hint ? el("span", { class: "hint", text: hint }) : null);
}

// After a principal changes, refresh whichever page the form was opened from.
function afterPrincipalChange() {
  if (location.hash.startsWith("#admin")) { if (typeof adminEnter === "function") adminEnter(adminState.sub); return; }
  if (ng.cid) ngCycle(ng.cid, ng.tab); else ngList();
}

function principalForm(p = null) {
  const v = p || {};
  const inp = (name, attrs = {}) => el("input", { name, value: v[name] || null, ...attrs });

  const binding = el("select", { name: "binding" },
    [["Nett", "Nett price (HNA after discount)"], ["Disc", "Discount (% off HNA)"], ["Nett to Disc", "Nett to Disc (formulary HNA × (1 − disc))"]].map(([b, l]) => el("option", { value: b, text: l })));
  binding.value = v.binding || "Nett";

  // Category: Drugs, CnC, or both.
  const has = (c) => String(v.category || "").includes(c);
  const catBoxes = ["Drugs", "CnC"].map((c) => {
    const box = el("input", { type: "checkbox", name: "cat", value: c });
    box.checked = p ? has(c) : c === "Drugs";
    return box;
  });
  const catField = el("fieldset", { class: "cat-field" }, el("legend", { text: "Category" }),
    el("div", { class: "check-row" }, catBoxes.map((b) => el("label", { class: "check" }, b, b.value))),
    el("span", { class: "hint", text: "Pick Drugs, CnC, or both." }));

  const negoType = el("select", { name: "nego_type" }, ["Renewal", "New Vendor", "Addendum", "Force Majeure"].map((t) => el("option", { value: t, text: t })));
  negoType.value = v.nego_type || "Renewal";

  const mouStart = inp("mou_start", { type: "date" });
  const mouEnd = inp("mou_end", { type: "date" });
  const mouRow = el("div", { class: "row2" }, field("Current MOU start", mouStart), field("Current MOU end", mouEnd, "Drives the 6-month alert"));
  const syncType = () => {
    const nv = negoType.value === "New Vendor";
    mouRow.hidden = nv; // a new vendor has no current MOU
    mouEnd.required = !nv;
  };
  negoType.addEventListener("change", syncType);

  const msg = el("p", { class: "error", role: "alert", hidden: "" });
  const form = el("form", { class: "progress-form" },
    field("Principal name", inp("name", { required: "" }), "Saved in capitals (PT CONTOH FARMA)"),
    field("Distributor (optional)", inp("distributor", { placeholder: "Optional. It can change and there can be several" })),
    el("div", { class: "row2" }, catField, field("Negotiation type", negoType)),
    field("Binding", binding, "Nett to Disc: final price/pc = formulary HNA ÷ conversion × (1 − final disc) × 1,11 (PO HNA when the formulary has none)"),
    mouRow,
    field("Contact person (PIC)", inp("contact_name")),
    el("div", { class: "row2" }, field("Contact email", inp("contact_email", { type: "email" })), field("WhatsApp", inp("contact_phone", { placeholder: "+62 …" }))),
    field("Notes (optional)", inp("notes", { placeholder: "Specific conditions, distributor terms, etc." })),
    msg,
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: p ? "Save" : "Add principal" })));
  syncType();
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const body = Object.fromEntries(new FormData(form).entries());
    body.category = catBoxes.filter((b) => b.checked).map((b) => b.value);
    delete body.cat;
    if (!body.category.length) { msg.textContent = "Pick Drugs, CnC, or both."; msg.hidden = false; return; }
    if (body.nego_type === "New Vendor") { body.mou_start = ""; body.mou_end = ""; }
    try {
      const saved = p ? await send(`/api/admin/nego/principals/${p.id}`, "PATCH", body) : await send("/api/admin/nego/principals", "POST", body);
      closeDrawer();
      toast(p ? "Saved" : `${saved.name} added`);
      afterPrincipalChange();
    } catch (ex) { msg.textContent = ex.message; msg.hidden = false; }
  });
  openPanel(el("p", { class: "eyebrow", text: p ? "Edit principal" : "New principal" }), el("h2", { id: "drawer-title", text: p ? p.name : "Add a principal" }), form);
}

function importPrincipals() {
  const file = el("input", { type: "file", accept: ".csv,.xlsx", class: "sr-only", name: "file" });
  const out = el("div", { role: "status" });
  const form = el("form", { class: "upload-form" },
    el("p", { class: "small" }, "Use the template. Same fields as Add principal: principal, distributor (optional), category, negotiation type, binding, items, current MOU period and contact."),
    el("div", { class: "actions" },
      el("a", { class: "act ghost small-btn", href: "/api/admin/nego/principals-template.xlsx", text: "Download Excel template ⤓" }),
      el("a", { class: "act ghost small-btn", href: "/api/admin/nego/principals-template.csv", text: "Download CSV template ⤓" })),
    el("p", { class: "muted small", text: "Principals with the same name are updated. Rows with a problem are skipped and listed below; the rest are saved." }),
    dropZone(file, "Choose the filled template"),
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Import" })), out);
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!file.files.length) { clear(out).append(el("p", { class: "error", text: "Choose a file first." })); return; }
    const fd = new FormData(); fd.append("file", file.files[0]);
    clear(out).append(el("p", { class: "muted", text: "Importing…" }));
    try {
      const r = await upload("/api/admin/nego/principals/import", fd);
      clear(out).append(el("p", { class: r.errors.length ? "warn-text" : "good-text", text: `${r.added} added, ${r.updated} updated${r.errors.length ? `, ${r.errors.length} skipped` : ""}.` }),
        r.errors.length ? el("ul", { class: "errors" }, r.errors.map((x) => el("li", { text: x }))) : null);
      afterPrincipalChange();
    } catch (ex) { clear(out).append(el("p", { class: "error", text: ex.message })); }
  });
  openPanel(el("p", { class: "eyebrow", text: "Principals" }), el("h2", { id: "drawer-title", text: "Import principals" }), form);
}

async function loadSample() {
  try {
    toast("Loading the sample…");
    const r = await send("/api/admin/nego/sample", "POST", {});
    toast(`${r.principals.added} principals, ${r.opened.length} negotiations opened`);
    ngList();
  } catch (e) { toast(e.message); }
}

// Date-only arithmetic in UTC, so the browser's time zone never shifts a day.
const utc = (iso) => { const [y, m, d] = iso.split("-").map(Number); return new Date(Date.UTC(y, m - 1, d)); };
const isoOf = (d) => d.toISOString().slice(0, 10);
function addYears(iso, n) { const d = utc(iso); d.setUTCFullYear(d.getUTCFullYear() + n); d.setUTCDate(d.getUTCDate() - 1); return isoOf(d); }
function nextDay(iso) { const d = utc(iso); d.setUTCDate(d.getUTCDate() + 1); return isoOf(d); }

function startCycleForm(p) {
  // A renewal starts the day after the current MOU ends. With no MOU end (a new vendor) nothing is
  // filled in: the negotiator enters both dates.
  const hasMou = !!p.mou_end && p.nego_type !== "New Vendor";
  const start0 = hasMou ? nextDay(p.mou_end) : "";
  const start = el("input", { type: "date", name: "contract_start", value: start0 || null, required: "" });
  const end = el("input", { type: "date", name: "contract_end", value: start0 ? addYears(start0, 3) : null, required: "" });
  start.addEventListener("change", () => { if (start.value && hasMou) end.value = addYears(start.value, 3); });
  const binding = el("select", { name: "binding" }, [["Nett", "Nett price (HNA after discount)"], ["Disc", "Discount (% off HNA)"], ["Nett to Disc", "Nett to Disc (formulary HNA × (1 − disc))"]].map(([b, l]) => el("option", { value: b, text: l })));
  binding.value = p.binding || "Nett";
  const fee = el("input", { name: "delivery_fee", value: "Free for all Siloam Hospitals units" });
  const msg = el("p", { class: "error", role: "alert", hidden: "" });
  const form = el("form", { class: "progress-form" },
    el("div", { class: "row2" }, field("New contract starts", start, hasMou ? null : "Required: enter the date"), field("New contract ends", end, hasMou ? "Three years from the start; change if needed" : "Required: enter the date")),
    field("Binding", binding), field("Delivery fee", fee), msg,
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Open negotiation" }), isAdmin() ? act("Edit principal", () => principalForm(p), "ghost") : null));
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!start.value || !end.value) { msg.textContent = "Enter the new contract start and end dates."; msg.hidden = false; return; }
    try {
      const c = await send("/api/admin/nego/cycles", "POST", { principal_id: p.id, category: p.row_category, ...Object.fromEntries(new FormData(form).entries()) });
      closeDrawer();
      toast("Negotiation opened. Next: build the item list.");
      location.hash = `#nego/${c.id}/files`;
    } catch (ex) { msg.textContent = ex.message; msg.hidden = false; }
  });
  openPanel(el("p", { class: "eyebrow", text: `Start negotiation · ${p.nego_type || "Renewal"}${p.row_category ? ` · ${p.row_category}` : ""}` }), el("h2", { id: "drawer-title", text: p.name }),
    el("p", { class: "muted", text: hasMou ? `Current MOU ends ${fmtDate(p.mou_end)}.` : p.nego_type === "New Vendor" ? "New vendor: there is no current MOU." : "No MOU end date on file." }), form);
}

// ---------- one cycle ----------
async function ngCycle(cid, tab) {
  const page = $("nego");
  if (ng.cid !== cid || !ng.ov) {
    clear(page).append(el("p", { class: "muted pad", text: "Loading…" }));
    ng.page = { q: "", filter: "all", status: "all", brand: "all", sort: "sort", sort_col: "", sort_dir: "asc", offset: 0, limit: 50 };
  }
  ng.cid = cid;
  ng.tab = ["items", "findings", "analysis", "brands", "negotiate", "files", "activity"].includes(tab) ? tab : "items";
  try { ng.ov = await api(`/api/nego/cycles/${cid}`); } catch (e) { clear(page).append(el("p", { class: "error pad", text: e.message }), el("a", { href: "#nego", text: "← All principals" })); return; }
  renderCycle();
}

async function refreshOverview() {
  ng.ov = await api(`/api/nego/cycles/${ng.cid}`);
  renderHeaderParts();
}

function renderCycle() {
  const page = clear($("nego"));
  const c = ng.ov.cycle;
  const tabs = el("nav", { class: "subtabs", "aria-label": "Negotiation sections" },
    [["items", "Items"], ["findings", "Findings"], ["analysis", "Analysis"], ["brands", "Brands"], ["negotiate", "Negotiate"], ["files", "Files & prepare"], ["activity", "Activity"]].map(([k, label]) =>
      el("a", { href: `#nego/${c.id}/${k}`, class: ng.tab === k ? "on" : null, "data-tab": k }, label, k === "findings" ? el("span", { class: "count", id: "ng-find-count" }) : null)));
  page.append(
    el("a", { class: "back", href: "#nego", text: "← All principals" }),
    el("div", { class: "page-head ng-head" },
      el("div", {},
        el("p", { class: "eyebrow", text: c.principal.distributor ? `Distributor: ${c.principal.distributor}` : "Principal" }),
        el("h1", { text: c.principal.name }),
        el("p", { class: "muted", text: `Contract ${fmtDate(c.contract_start)} – ${fmtDate(c.contract_end)} · Binding ${({ Disc: "discount", "Nett to Disc": "nett to disc" })[c.binding] || "nett price"} · PPN ${pct(ng.ov.ppn, 0)}` })),
      el("div", { class: "ng-tools" },
        el("a", { class: "act secondary", href: `/api/nego/cycles/${c.id}/export.xlsx`, text: "Download Template_Nego" }),
        canWork() && c.status === "open" ? act("Principal link", () => linkPanel(c), "secondary") : null,
        canWork() && c.status === "open" ? el("a", { class: "act ghost", href: `/p?cycle_id=${c.id}`, target: "_blank", rel: "noopener", text: "Preview Portal ↗", title: "Open the vendor's page display in a new tab" }) : null,
        isAdmin() ? act("Edit principal", () => principalForm(c.principal), "ghost") : null)),
    el("div", { id: "ng-steps" }),
    el("div", { class: "kpis5", id: "ng-kpis" }),
    el("div", { id: "ng-stages" }),
    tabs,
    el("div", { id: "ng-body" }));
  renderHeaderParts();
  ({ items: renderItems, findings: renderFindings, analysis: renderAnalysis, brands: renderBrands, negotiate: renderNegotiate, files: renderFiles, activity: renderActivity })[ng.tab]($("ng-body"));
}

// The step after `at`: Review RFQ is optional, so it is only offered when there are anomalies to send back.
function nextStepIndex(keys, at, a) {
  let n = at + 1;
  if (keys[n] === "review_rfq" && !(a && a.open_high > 0)) n++;
  return n;
}

function renderHeaderParts() {
  const { cycle: c, steps, kpis: k, anomalies: a, escalation: esc } = ng.ov;
  const keys = steps.map((s) => s.key);
  const at = keys.indexOf(c.current_step);
  const done = c.status === "closed" || c.current_step === "closed";
  const stepsBox = clear($("ng-steps"));
  const shown = steps.filter((s) => s.key !== "closed" && (s.key !== "review_rfq" || c.current_step === "review_rfq" || (a && a.open_high > 0)));
  const ol = el("ol", { class: "stepper" }, [...shown.map((s) => {
    const i = keys.indexOf(s.key);
    const passed = done || i < at;
    return el("li", { class: passed ? "done" : i === at ? "now" : "" },
      el("span", { class: "n", text: passed ? "✓" : String(i) }),
      el("span", { class: "t" }, el("b", { text: s.label }), el("small", { class: `who ${s.who}`, text: s.optional ? `${WHO[s.who]} · opsional` : WHO[s.who] })));
  }), done ? el("li", { class: "done final" }, el("span", { class: "n", text: "✓" }), el("span", { class: "t" }, el("b", { text: "Selesai" }), el("small", { class: "who", text: c.completed_at ? fmtDate(c.completed_at.slice(0, 10)) : "" }))) : null].filter(Boolean));
  const nextIdx = nextStepIndex(keys, at, a);
  const nextKey = keys[nextIdx];
  const noItems = !k.items;
  const move = canWork() && !done ? el("div", { class: "step-actions" },
    el("p", { class: "small", text: STEP_DESC[c.current_step] }),
    nextKey ? (noItems
      ? el("span", { class: "step-blocked" }, el("button", { type: "button", class: "act", disabled: "", text: `Move to ${steps[nextIdx].label}` }),
          el("span", { class: "small warn-text", text: "Build the item list first (Files & prepare). The principal can't start without items." }))
      : act(`Move to ${steps[nextIdx].label}`, () => moveStep(nextKey))) : null,
    at > 0 ? act("Back a step", () => backDrawer(keys[at - 1]), "ghost") : null) :
    canWork() && done ? el("div", { class: "step-actions" }, el("p", { class: "small", text: STEP_DESC.closed }), act("Reopen (back to Submission)", () => backDrawer("submission"), "ghost")) :
    el("p", { class: "small muted step-actions", text: STEP_DESC[c.current_step] });
  const sent = (c.submitted_steps || {})[c.current_step];
  const turn = steps[at] && steps[at].who === "principal";
  const late = c.overdue_days > 0;
  const banner = done ? el("p", { class: "sent-banner", text: `Negosiasi selesai${c.completed_at ? ` pada ${new Date(c.completed_at).toLocaleString("en-GB")}` : ""}. Harga sepakat siap untuk BAK dan PKS.` })
    : sent ? el("p", { class: "sent-banner", text: `The principal sent ${steps[at].label} on ${new Date(sent.at).toLocaleString("en-GB")}. Review the findings, then move to the next step.` })
    : turn ? el("p", { class: late ? "turn-banner late" : "turn-banner", text: `It's the principal's turn${c.step_due ? `, batas waktu ${fmtDate(c.step_due)}` : ""}${late ? `. Terlambat ${c.overdue_days} hari.` : ". Share the Principal link to let them in."}` }) : null;
  const returned = ng.ov.return_note ? el("p", { class: "turn-banner", text: `Dikembalikan ke principal: ${ng.ov.return_note.message || "(tanpa pesan)"}${ng.ov.return_note.flagged ? ` · ${ng.ov.return_note.flagged} SKU ditandai` : ""}` }) : null;
  const escBanner = esc && (esc.status === "needed" || esc.status === "approved") ? el("p", { class: esc.status === "needed" ? "turn-banner late" : "sent-banner", text: esc.status === "needed" ? `Eskalasi: cost impact FIN ${money(esc.impact)} (${pct(esc.impact_pct, 1)} dari nilai PO setahun) melewati batas. Perlu persetujuan sebelum Submission.` : `Eskalasi disetujui oleh ${esc.approval.by}: ${esc.approval.note}` }) : null;
  stepsBox.append(el("div", { class: "card stepcard" }, ol, banner, returned, escBanner, move));

  const latest = k.latest_impact;
  const mForm = k.margin_vs_formulary;
  const mPo = k.margin_vs_po;
  clear($("ng-kpis")).append(
    tile("Items", num(k.items), `${num(k.active)} active · ${num(k.discontinued)} discontinued`),
    tile("RFQ received", `${num(k.rfq_filled)} / ${num(k.active)}`, k.rfq_filled ? `${pct(k.rfq_filled / Math.max(k.active, 1), 0)} of active items` : "waiting for the principal"),
    tile("Open findings", num(a.open), a.open_high ? `${num(a.open_high)} high · ${num(a.kept)} kept · ${num(a.fixed)} fixed` : `${num(a.kept)} kept · ${num(a.fixed)} fixed`),
    tile("Cost impact a year", latest == null ? "—" : `${latest > 0 ? "+" : ""}${money(latest)}`, latest == null ? "once prices come in" : `latest step vs nilai PO setahun (+PPN) ${money(k.baseline_spend)}`),
    tile("Margin impact (Formularium)", mForm == null ? "—" : `${mForm >= 0 ? "+" : ""}${money(mForm)}`, mPo == null ? "vs Formularium" : `vs Form · vs PO: ${mPo >= 0 ? "+" : ""}${money(mPo)} a year`)
  );
  const tiles = $("ng-kpis").children;
  if (a.open_high) tiles[2].classList.add("alert");
  if (latest != null) tiles[3].classList.add(latest > 0 ? "alert" : "good");
  if (mForm != null && tiles[4]) tiles[4].classList.add(mForm >= 0 ? "good" : "alert");
  const fc = $("ng-find-count");
  if (fc) fc.textContent = a.open ? String(a.open) : "";

  const stages = k.stages.filter((s) => s.items);
  const box = clear($("ng-stages"));
  if (stages.length) {
    box.append(el("div", { class: "stage-strip" }, stages.map((s) =>
      el("div", { class: "st" },
        el("span", { class: "small muted", text: s.label }),
        el("b", { class: s.impact > 0 ? "up" : "down", text: s.impact == null ? "—" : `${s.impact > 0 ? "+" : ""}${money(s.impact)} (Cost)` }),
        el("div", { class: "small", style: "font-size: 11px; margin-top: 2px;" },
          el("span", { class: (s.margin_impact_vs_formulary || 0) >= 0 ? "good-text" : "warn-text", text: `Margin vs Form: ${s.margin_impact_vs_formulary != null ? money(s.margin_impact_vs_formulary) : "—"}` }),
          " · ",
          el("span", { class: (s.margin_impact_vs_po || 0) >= 0 ? "good-text" : "warn-text", text: `vs PO: ${s.margin_impact_vs_po != null ? money(s.margin_impact_vs_po) : "—"}` })
        ),
        el("span", { class: "small", text: `${num(s.increases)} up · ${num(s.decreases)} down of ${num(s.items)}` })))));
  }
}

const PRINCIPAL_STEPS = new Set(["rfq", "feedback1", "submission"]);

function showFindingsWarningModal(openFindings, nextStepLabel) {
  return new Promise((resolve) => {
    const body = openPanel(
      el("p", { class: "eyebrow warn-text", text: "Open Findings Warning" }),
      el("h2", { id: "drawer-title", text: `Proceed to ${nextStepLabel}?` }),
      el("div", { class: "preview bad" },
        el("h3", { text: `⚠️ ${openFindings} Open Finding${openFindings === 1 ? "" : "s"} Detected` }),
        el("p", { class: "small", text: `There are currently ${openFindings} unresolved anomaly or pricing rule findings. We recommend reviewing or resolving them before advancing the negotiation step.` })
      ),
      el("div", { class: "actions", style: "margin-top: 18px;" },
        act("Review Findings First", () => { closeDrawer(); location.hash = `#nego/${ng.cid}/findings`; resolve(false); }, "primary"),
        act(`Proceed to ${nextStepLabel} Anyway`, () => { closeDrawer(); resolve(true); }, "ghost"),
        act("Cancel", () => { closeDrawer(); resolve(false); }, "ghost")
      )
    );
  });
}

// Warning before the step moves on while active SKUs have no complete RFQ.
function incompleteModal(r, nextLabel) {
  return new Promise((resolve) => {
    const list = el("ul", { class: "errors", style: "max-height:240px;overflow:auto" }, (r.incomplete || []).map((x) => el("li", { text: `${x.erp_code} ${x.item_name}: ${x.missing.join(", ")} belum diisi` })));
    openPanel(
      el("p", { class: "eyebrow warn-text", text: "RFQ belum lengkap" }),
      el("h2", { id: "drawer-title", text: `Pindah ke ${nextLabel}?` }),
      el("div", { class: "preview bad" }, el("h3", { text: `${num(r.incomplete_count)} SKU aktif belum punya RFQ lengkap` }), list),
      el("div", { class: "actions", style: "margin-top: 18px;" },
        act("Kembali dulu", () => { closeDrawer(); resolve(false); }, "primary"),
        act(`Tetap pindah ke ${nextLabel}`, () => { closeDrawer(); resolve(true); }, "ghost")));
  });
}

async function moveStep(step) {
  const c = ng.ov.cycle;
  const label = ng.ov.steps.find((s) => s.key === step).label;
  const keys = ng.ov.steps.map((s) => s.key);
  const curIdx = keys.indexOf(c.current_step);
  const targetIdx = keys.indexOf(step);
  const isAdvancing = targetIdx > curIdx;

  if (isAdvancing && (!c.items || c.items === 0) && (!ng.ov.kpis || ng.ov.kpis.items === 0)) {
    toast("Cannot advance: negotiation has 0 items. Build or import item list first.");
    return;
  }

  const openFindings = ng.ov.anomalies?.open || 0;
  if (isAdvancing && openFindings > 0 && c.current_step !== "review_rfq") {
    const proceed = await showFindingsWarningModal(openFindings, label);
    if (!proceed) return;
  }

  if (!PRINCIPAL_STEPS.has(step)) {
    if (!confirm(`Move this negotiation to "${label}"?`)) return;
    return doMove(step, label, false);
  }
  // A principal step: set the deadline (never a day that has passed) and optionally send the link.
  let auto = { enabled: false, missing: [] };
  try {
    const notifyRes = await api("/api/admin/nego/notify");
    if (notifyRes && typeof notifyRes === "object") auto = { enabled: Boolean(notifyRes.enabled ?? notifyRes.configured), missing: Array.isArray(notifyRes.missing) ? notifyRes.missing : [], ...notifyRes };
  } catch (_) { /* shown as off */ }
  const p = ng.ov?.cycle?.principal || {};
  const to = [p.contact_email, p.contact_phone].filter(Boolean).join(" · ");
  const sendBox = el("input", { type: "checkbox", id: "send-link" });
  sendBox.checked = auto.enabled && !!to;
  if (!auto.enabled || !to) sendBox.disabled = true;
  const missingVars = (Array.isArray(auto.missing) && auto.missing.length > 0) ? auto.missing.join(" and ") : "POWER_AUTOMATE_EMAIL_URL";
  const why = !auto.enabled ? `Automatic sending is off: set ${missingVars} on the server (see deploy/POWER_AUTOMATE.md). You can still share the link by hand.`
    : !to ? "The principal has no email or WhatsApp number. Add one with Edit principal, or share the link by hand." : `To: ${to}`;
  const days = (await api("/api/settings").catch(() => ({ settings: [] }))).settings.find((x) => x.key === "step_days");
  const d0 = new Date(); d0.setDate(d0.getDate() + (days ? days.value : 7));
  const today = new Date().toISOString().slice(0, 10);
  const dueInput = el("input", { type: "date", id: "step-due", value: d0.toISOString().slice(0, 10), min: today });
  const body = openPanel(el("p", { class: "eyebrow", text: "Move step" }), el("h2", { id: "drawer-title", text: `Move to ${label}` }),
    el("p", { class: "muted", text: "It's the principal's turn at this step. They fill it online or in Excel from their link." }),
    el("form", { class: "progress-form", onsubmit: (e) => e.preventDefault() }, field("Batas waktu untuk principal", dueInput, "Tanggal yang sudah lewat tidak bisa dipilih.")),
    el("label", { class: "check big-check" }, sendBox, el("span", {}, el("b", { text: "Send the link to the principal automatically" }), el("span", { class: "small muted block", text: "Email and WhatsApp, through Power Automate." }))),
    el("p", { class: `small ${auto.enabled && to ? "" : "warn-text"}`, text: why }),
    el("div", { class: "actions" }, act(`Move to ${label}`, async () => {
      if (!dueInput.value || dueInput.value < today) { toast("Batas waktu tidak boleh tanggal yang sudah lewat."); return; }
      closeDrawer(); await doMove(step, label, sendBox.checked, dueInput.value);
    }), act("Cancel", closeDrawer, "ghost")));
  return body;
}

// Back a step: Siloam writes why, marks the SKUs to fix (they show "Perlu dicek" in the portal) and sets a new deadline.
function backDrawer(step) {
  const label = ng.ov.steps.find((s) => s.key === step).label;
  const principalStep = PRINCIPAL_STEPS.has(step);
  const msg = el("textarea", { rows: "4", placeholder: "Alasan dikembalikan, tampil di portal principal" });
  const flag = el("input", { type: "checkbox" });
  flag.checked = (ng.ov.anomalies?.open_high || 0) > 0;
  const today = new Date().toISOString().slice(0, 10);
  const d0 = new Date(); d0.setDate(d0.getDate() + 7);
  const due = el("input", { type: "date", min: today, value: d0.toISOString().slice(0, 10) });
  openPanel(el("p", { class: "eyebrow", text: "Back a step" }), el("h2", { id: "drawer-title", text: `Kembali ke ${label}` }),
    el("p", { class: "muted", text: principalStep ? "Negosiasi dikembalikan ke principal. Tulis alasannya dan tandai SKU yang harus diperbaiki." : "Negosiasi dikembalikan ke step Siloam." }),
    el("form", { class: "progress-form", onsubmit: (e) => e.preventDefault() },
      field("Pesan untuk principal", msg),
      principalStep ? el("label", { class: "check big-check" }, flag, el("span", {}, el("b", { text: `Tandai SKU dengan temuan high untuk dicek (${ng.ov.anomalies?.open_high || 0})` }), el("span", { class: "small muted block", text: "Di portal SKU ini tampil Perlu dicek. SKU lain tetap Selesai." }))) : null,
      principalStep ? field("Batas waktu baru", due, "Tidak bisa tanggal yang sudah lewat.") : null),
    el("div", { class: "actions" }, act(`Kembali ke ${label}`, async () => {
      if (principalStep && due.value < today) { toast("Batas waktu tidak boleh tanggal yang sudah lewat."); return; }
      closeDrawer();
      await doMove(step, label, false, principalStep ? due.value : null, { message: msg.value.trim(), flag_findings: flag.checked });
    }), act("Cancel", closeDrawer, "ghost")));
}

async function doMove(step, label, sendLink, due, extra = {}, confirmIncomplete = false) {
  try {
    const r = await send(`/api/admin/nego/cycles/${ng.cid}/step`, "POST", { step, send_link: sendLink, due: due || null, confirm_incomplete: confirmIncomplete, ...extra });
    const m = r.message;
    toast(m ? (m.status === "queued" ? `Now at ${label}. Link sent to ${m.recipient}` : `Now at ${label}. Link not sent: ${m.reason === "no_contact" ? "no contact details" : "automatic sending is off"}`) : `Now at ${label}`);
    await refreshOverview();
    if (ng.tab === "findings") renderFindings($("ng-body"));
    if (ng.tab === "activity") renderActivity($("ng-body"));
    if (ng.tab === "negotiate") renderNegotiate($("ng-body"));
  } catch (e) {
    if (e.data && e.data.needs_confirm) { if (await incompleteModal(e.data, label)) return doMove(step, label, sendLink, due, extra, true); return; }
    toast(e.message);
  }
}

// ---------- items ----------
const SECTION_CLASS = { rfq: "sec-rfq", current_mou: "sec-mou", counter_offer: "sec-co", feedback1: "sec-fb", online_nego: "sec-on", analysis: "sec-eng" };
const ITEM_COLS = [
  ["erp_code", "ERP Code", "rfq", "text"],
  ["item_name", "Item Name", "rfq", "text"],
  ["brand", "Brand", "rfq", "text"],
  ["catalog_no", "REF", "rfq", "text"],
  ["item_status", "Status", "rfq", "text"],
  ["mou_qty", "Qty/PO unit", "current_mou", "qty"],
  ["mou_hna", "HNA/PO unit", "current_mou", "money"],
  ["mou_disc", "Disc", "current_mou", "pct"],
  ["mou_unit_price", "Price/pc", "current_mou", "price"],
  ["rfq_qty", "Qty/PO unit", "rfq", "qty"],
  ["rfq_hna", "HNA/PO unit", "rfq", "money"],
  ["rfq_disc", "Disc", "rfq", "pct"],
  ["rfq_unit_price", "Price/pc", "rfq", "price"],
  ["co_disc", "Disc", "counter_offer", "pct"],
  ["co_unit_price", "Price/pc", "counter_offer", "price"],
  ["fb1_disc", "Disc", "feedback1", "pct"],
  ["fb1_unit_price", "Price/pc", "feedback1", "price"],
  ["on_disc", "Disc", "online_nego", "pct"],
  ["on_unit_price", "Price/pc", "online_nego", "price"],
];
const SECTION_STEP = { rfq: "rfq", feedback1: "feedback1" };
const secOf = (fld) => (ITEM_COLS.find((c) => c[0] === fld) || [])[2];
function stepPassed(step) {
  if (!step) return false;
  const keys = ng.ov.steps.map((s) => s.key);
  return keys.indexOf(ng.ov.cycle.current_step) > keys.indexOf(step);
}
const SECTION_SHORT = { rfq: "RFQ & Identification", current_mou: "Current MOU", counter_offer: "Counter offer", feedback1: "Feedback I", online_nego: "Online nego", analysis: "Analisis & catatan" };

function cellValue(kind, v) {
  if (v == null || v === "") return "";
  if (kind === "pct") return typeof v === "number" ? pct(v, Math.abs(v * 100 - Math.round(v * 100)) > 1e-6 ? 1 : 0) : String(v);
  if (kind === "money") return num(v, 0);
  if (kind === "price") return num(v, v < 100 ? 2 : 0);
  if (kind === "qty") return num(v, v % 1 ? 2 : 0);
  return String(v);
}

const ALL_ITEM_COLS = [
  ...ITEM_COLS.map(([fld, label, sec, kind]) => ({ fld, label, sec, kind })),
  { fld: "bench_pp", label: "Market/pc", sec: "engine", kind: "price" },
  { fld: "change_pct", label: "vs MOU", sec: "engine", kind: "pct" },
  { fld: "impact", label: "Impact/yr", sec: "engine", kind: "money" },
  { fld: "flags", label: "Findings", sec: "engine", kind: "num" },
];

async function renderItems(body) {
  clear(body);
  const p = ng.page;

  // Filter toolbar inputs
  const q = el("input", { type: "search", id: "ng-q", value: p.q || null, placeholder: "Search ERP code, name, brand, REF…" });
  
  const statusSelect = el("select", { id: "ng-status" },
    [["all", "All statuses"], ["Active", "Active only"], ["Discontinue", "Discontinued only"], ["New Item", "New items"]].map(([k, l]) =>
      el("option", { value: k, text: l })
    )
  );
  statusSelect.value = p.status || "all";

  const brandSelect = el("select", { id: "ng-brand" },
    [el("option", { value: "all", text: "All brands" }),
     ...(p.brands || []).map((b) => el("option", { value: b, text: b }))]
  );
  brandSelect.value = p.brand || "all";

  const filterSelect = el("select", { id: "ng-filter" }, FILTERS.map(([k, l]) => el("option", { value: k, text: l })));
  filterSelect.value = p.filter || "all";

  const sortSelect = el("select", { id: "ng-sort" }, [
    ...SORTS.map(([k, l]) => el("option", { value: k, text: l })),
    ...(p.sort_col ? [el("option", { value: "custom", text: `Custom: ${p.sort_col} (${p.sort_dir})` })] : [])
  ]);
  sortSelect.value = p.sort_col ? "custom" : (p.sort || "sort");

  const countBadge = el("span", { class: "filter-badge", id: "ng-items-badge", text: "Loading items…" });
  const metaSummary = el("div", { class: "items-status-summary" }, countBadge);

  const legend = el("div", { class: "legend-cells" },
    el("span", {}, el("i", { class: "c-req" }), "Principal fills"),
    el("span", {}, el("i", { class: "c-opt" }), "Optional"),
    el("span", {}, el("i", { class: "c-ref" }), "Reference / calculated"),
    el("span", {}, el("i", { class: "c-adm" }), "Siloam fills"));

  const tagsRow = el("div", { class: "active-tags-row", id: "ng-active-tags", style: "display: none;" });

  const filterBar = el("div", { class: "items-filter-bar" },
    el("label", { class: "filter-field grow-search" }, "Search items", q),
    el("label", { class: "filter-field" }, "Status", statusSelect),
    el("label", { class: "filter-field" }, "Brand", brandSelect),
    el("label", { class: "filter-field" }, "Condition", filterSelect),
    el("label", { class: "filter-field" }, "Sort by", sortSelect),
    metaSummary,
    legend,
    tagsRow
  );

  const wrap = el("div", { class: "table-wrap card items-wrap" });
  const pager = el("div", { class: "pager" });

  body.append(filterBar, wrap, pager);

  let timer;
  q.addEventListener("input", () => {
    clearTimeout(timer);
    timer = setTimeout(() => {
      p.q = q.value.trim();
      p.offset = 0;
      loadItems(wrap, pager);
    }, 250);
  });

  statusSelect.addEventListener("change", () => {
    p.status = statusSelect.value;
    p.offset = 0;
    loadItems(wrap, pager);
  });

  brandSelect.addEventListener("change", () => {
    p.brand = brandSelect.value;
    p.offset = 0;
    loadItems(wrap, pager);
  });

  filterSelect.addEventListener("change", () => {
    p.filter = filterSelect.value;
    p.offset = 0;
    loadItems(wrap, pager);
  });

  sortSelect.addEventListener("change", () => {
    if (sortSelect.value !== "custom") {
      p.sort = sortSelect.value;
      p.sort_col = "";
      p.sort_dir = "asc";
      p.offset = 0;
      loadItems(wrap, pager);
    }
  });

  loadItems(wrap, pager);
}

async function loadItems(wrap, pager) {
  const p = ng.page;
  const qs = new URLSearchParams({
    q: p.q,
    filter: p.filter,
    status: p.status,
    brand: p.brand,
    sort: p.sort_col ? "" : p.sort,
    sort_col: p.sort_col || "",
    sort_dir: p.sort_dir || "asc",
    offset: p.offset,
    limit: p.limit || 50,
  });
  wrap.classList.add("loading");
  let d;
  try {
    d = await api(`/api/nego/cycles/${ng.cid}/items?${qs}`);
  } catch (e) {
    clear(wrap).append(el("p", { class: "error pad", text: e.message }));
    return;
  } finally {
    wrap.classList.remove("loading");
  }

  p.brands = d.brands || [];
  p.statuses = d.statuses || [];

  // Populate brand select options dynamically from cycle data
  const brandSelect = $("ng-brand");
  if (brandSelect && p.brands.length) {
    const cur = p.brand;
    clear(brandSelect).append(
      el("option", { value: "all", text: "All brands" }),
      ...p.brands.map((b) => el("option", { value: b, text: b }))
    );
    brandSelect.value = cur && p.brands.includes(cur) ? cur : "all";
  }

  // Update status select if needed
  const statusSelect = $("ng-status");
  if (statusSelect) statusSelect.value = p.status || "all";

  // Update sort select value
  const sortSelect = $("ng-sort");
  if (sortSelect) {
    if (p.sort_col) {
      let opt = sortSelect.querySelector("option[value='custom']");
      if (!opt) {
        opt = el("option", { value: "custom" });
        sortSelect.append(opt);
      }
      const colObj = ALL_ITEM_COLS.find((c) => c.fld === p.sort_col);
      opt.textContent = `Column: ${colObj ? colObj.label : p.sort_col} (${p.sort_dir === "asc" ? "▲ Asc" : "▼ Desc"})`;
      sortSelect.value = "custom";
    } else {
      const opt = sortSelect.querySelector("option[value='custom']");
      if (opt) opt.remove();
      sortSelect.value = p.sort || "sort";
    }
  }

  // Update active tags and count badge
  updateActiveFilterTags(wrap, pager, d);

  if (!d.total) {
    const totalInCycle = (d.counts && d.counts.total) != null ? d.counts.total : (ng.ov && ng.ov.kpis ? ng.ov.kpis.items : 0);
    const hasFilter = !!(p.q || p.filter !== "all" || p.status !== "all" || p.brand !== "all" || p.sort_col || p.sort !== "sort");
    clear(wrap).append(el("div", { class: "empty-state" },
      el("h3", { text: !totalInCycle ? "No items yet" : "No items match your filter" }),
      el("p", { class: "muted", text: !totalInCycle ? "Build the item list from the PO export, formulary and current MOU." : "Try adjusting your search keywords, status, brand, or condition filters." }),
      hasFilter ? el("div", { class: "actions" }, act("Reset all filters", () => resetItemFilters(wrap, pager))) : null,
      !totalInCycle && isAdmin() ? act("Go to Files & prepare", () => { location.hash = `#nego/${ng.cid}/files`; }) : null));
    clear(pager);
    return;
  }

  const t = el("table", { class: "items" });
  const sections = [];
  ITEM_COLS.forEach(([, , sec]) => {
    const last = sections[sections.length - 1];
    if (last && last.sec === sec) last.n++;
    else sections.push({ sec, n: 1 });
  });
  const cols = ng.ov.columns.reduce((m, c) => { m[c.field] = c; return m; }, {});

  // Complete column definitions (current_mou, rfq, counter_offer, feedback1, online_nego, and analysis columns)
  const allCols = [
    ...ITEM_COLS.map(([fld, label, sec, kind]) => ({ fld, label, sec, kind, title: (cols[fld] || {}).header || label })),
    { fld: "margin_pc", label: "Margin/pc", sec: "analysis", kind: "money", title: "Margin per piece (Base price HNA/pc * 1.11 - step price)" },
    { fld: "margin_vs_formulary", label: "vs Form", sec: "analysis", kind: "money", title: "Annualized margin impact vs formularium" },
    { fld: "margin_vs_po", label: "vs PO", sec: "analysis", kind: "money", title: "Annualized margin impact vs last PO" },
    { fld: "bench_pp", label: "Market/pc", sec: "analysis", kind: "price", title: "Harga kompetitor terendah per pcs dari file benchmark periode" },
    { fld: "change_pct", label: "vs Acuan", sec: "analysis", kind: "pct", title: "Perubahan harga vs acuan: Nett Price formularium, jika kosong MOU, jika kosong rata-rata PO (lihat kolom Acuan di detail item)" },
    { fld: "impact", label: "Impact/yr", sec: "analysis", kind: "money", title: "Cost impact setahun = Unit Price (+PPN) × qty setahun − nilai PO setahun (+PPN)" },
    { fld: "flags", label: "Findings", sec: "analysis", kind: "num", title: "Open rule findings and anomalies" },
  ];

  const thead = el("thead", {},
    el("tr", { class: "sec" },
      sections.map((x) => el("th", { colspan: x.n, class: SECTION_CLASS[x.sec], text: SECTION_SHORT[x.sec] })),
      el("th", { colspan: 7, class: "sec-eng", text: "Analysis & Margins" })),
    el("tr", {}, allCols.map((col) => {
      const isSorted = p.sort_col === col.fld;
      const isAsc = isSorted && p.sort_dir === "asc";
      const isDesc = isSorted && p.sort_dir === "desc";
      const th = el("th", {
        class: [
          col.kind !== "text" ? "num" : "",
          col.fld === "item_name" ? "namecol" : "",
          col.fld === "erp_code" ? "codecol" : "",
          "sortable",
          isSorted ? `sorted sorted-${p.sort_dir}` : ""
        ].filter(Boolean).join(" "),
        title: isSorted
          ? `${col.title || col.label} (Sorted ${p.sort_dir === "asc" ? "Ascending" : "Descending"}. Click to ${p.sort_dir === "asc" ? "sort descending" : "clear sort"})`
          : `${col.title || col.label} (Click to sort)`
      });

      const thWrap = el("div", { class: "th-wrap" });
      const thText = el("span", { class: "th-text", text: col.label });
      thWrap.append(thText);

      // Header filter button for Brand and Status
      if (col.fld === "brand") {
        const isFiltered = p.brand && p.brand !== "all";
        const filterBtn = el("button", {
          type: "button",
          class: `th-filter-trigger ${isFiltered ? "filtered" : ""}`,
          title: isFiltered ? `Brand filtered: ${p.brand} (Click to change)` : "Filter by Brand",
          text: "⚲",
          onclick: (e) => {
            e.stopPropagation();
            toggleColFilterMenu(th, "brand", p.brands, p.brand, (val) => {
              p.brand = val;
              p.offset = 0;
              loadItems(wrap, pager);
            });
          }
        });
        thWrap.append(filterBtn);
      } else if (col.fld === "item_status") {
        const isFiltered = p.status && p.status !== "all";
        const filterBtn = el("button", {
          type: "button",
          class: `th-filter-trigger ${isFiltered ? "filtered" : ""}`,
          title: isFiltered ? `Status filtered: ${p.status} (Click to change)` : "Filter by Status",
          text: "⚲",
          onclick: (e) => {
            e.stopPropagation();
            toggleColFilterMenu(th, "status", ["Active", "Discontinue", "New Item"], p.status, (val) => {
              p.status = val;
              p.offset = 0;
              loadItems(wrap, pager);
            });
          }
        });
        thWrap.append(filterBtn);
      }

      // Sort indicator icon
      const sortIcon = el("span", {
        class: "th-sort-icon",
        text: isAsc ? "▲" : isDesc ? "▼" : "↕"
      });
      thWrap.append(sortIcon);
      th.append(thWrap);

      // Click header to sort
      th.addEventListener("click", () => {
        if (p.sort_col !== col.fld) {
          p.sort_col = col.fld;
          p.sort_dir = "asc";
          p.sort = "";
        } else if (p.sort_dir === "asc") {
          p.sort_dir = "desc";
        } else {
          p.sort_col = "";
          p.sort_dir = "asc";
          p.sort = "sort";
        }
        p.offset = 0;
        loadItems(wrap, pager);
      });

      return th;
    }))
  );
  t.append(thead);

  // Table rows
  const tb = el("tbody");
  for (const it of d.items) {
    const tr = el("tr", { class: `clickable ${it.item_status === "Discontinue" ? "disc" : ""}`, tabindex: 0, onclick: () => openItem(it.id) });
    tr.addEventListener("keydown", (e) => { if (e.key === "Enter") openItem(it.id); });
    for (const [fld, , , kind] of ITEM_COLS) {
      const who = (cols[fld] || {}).who;
      const cls = [kind !== "text" ? "num" : "", who === "principal" ? "c-req" : who === "optional" ? "c-opt" : who === "admin" ? "c-adm" : "c-ref",
        fld === "item_name" ? "namecol" : "", fld === "erp_code" ? "codecol" : "",
        who === "principal" && (it[fld] == null || it[fld] === "") && it.item_status !== "Discontinue" && stepPassed(SECTION_STEP[secOf(fld)]) ? "empty" : ""].join(" ");
      tr.append(el("td", { class: cls, text: cellValue(kind, it[fld]), title: fld === "item_name" ? it.item_name : null }));
    }
    const ch = it.change_pct;
    tr.append(el("td", { class: `num ${it.margin_pc > 0 ? "good-text" : it.margin_pc < 0 ? "warn-text" : ""}`, text: it.margin_pc == null ? "" : money(it.margin_pc) }));
    tr.append(el("td", { class: `num ${it.margin_vs_formulary > 0 ? "good-text" : it.margin_vs_formulary < 0 ? "warn-text" : ""}`, text: it.margin_vs_formulary == null ? "" : money(it.margin_vs_formulary) }));
    tr.append(el("td", { class: `num ${it.margin_vs_po > 0 ? "good-text" : it.margin_vs_po < 0 ? "warn-text" : ""}`, text: it.margin_vs_po == null ? "" : money(it.margin_vs_po) }));
    tr.append(el("td", { class: "num", text: it.bench_pp == null ? "" : num(it.bench_pp, it.bench_pp < 100 ? 2 : 0), title: it.bench_source || null }));
    tr.append(el("td", { class: `num ${ch > 0.0005 ? "up" : ch < -0.0005 ? "down" : ""}`, title: it.bench_tag ? `vs ${it.bench_tag}` : null, text: ch == null ? "" : `${ch > 0 ? "+" : ""}${pct(ch)}${it.bench_tag && it.bench_tag !== "tanpa acuan" ? ` ${({ Formularium: "Form", MOU: "MOU", PO: "PO" })[it.bench_tag] || ""}` : ""}` }));
    tr.append(el("td", { class: `num ${it.impact > 0 ? "up" : it.impact < 0 ? "down" : ""}`, text: it.impact == null ? "" : `${it.impact > 0 ? "+" : ""}${money(it.impact)}` }));
    tr.append(el("td", { class: "flags" }, it.flags.slice(0, 2).map((x) => el("span", { class: `sev ${x.severity}`, text: x.label })),
      it.flags.length > 2 ? el("span", { class: "small muted", text: ` +${it.flags.length - 2}` }) : null));
    tb.append(tr);
  }
  t.append(tb);
  clear(wrap).append(t);

  // Pager with page size selector
  const limit = p.limit || 50;
  const p0 = d.offset + 1, p1 = Math.min(d.offset + d.items.length, d.total);
  const pageSizeSelect = el("select", { class: "small-btn", title: "Items per page" },
    [25, 50, 100, 200, 500].map((sz) => el("option", { value: sz, text: `${sz} / page` }))
  );
  pageSizeSelect.value = limit;
  pageSizeSelect.addEventListener("change", () => {
    p.limit = Number(pageSizeSelect.value);
    p.offset = 0;
    loadItems(wrap, pager);
  });

  clear(pager).append(
    el("span", { class: "small muted", text: `Showing ${num(p0)}–${num(p1)} of ${num(d.total)} items` }),
    pageSizeSelect,
    el("button", { type: "button", class: "act ghost", text: "← Previous", disabled: d.offset ? null : "", onclick: () => { p.offset = Math.max(0, d.offset - limit); loadItems(wrap, pager); } }),
    el("button", { type: "button", class: "act ghost", text: "Next →", disabled: p1 < d.total ? null : "", onclick: () => { p.offset = d.offset + limit; loadItems(wrap, pager); } }));
}

function updateActiveFilterTags(wrap, pager, d) {
  const p = ng.page;
  const tagsRow = $("ng-active-tags");
  const badge = $("ng-items-badge");

  const totalCycleItems = (d.counts && d.counts.total) != null ? d.counts.total : (ng.ov && ng.ov.kpis ? ng.ov.kpis.items : d.total);

  if (badge) {
    if (d.total < totalCycleItems) {
      badge.textContent = `Showing ${num(d.total)} of ${num(totalCycleItems)} items (${num(totalCycleItems - d.total)} filtered out)`;
      badge.className = "filter-badge active-filter";
    } else {
      badge.textContent = `${num(d.total)} items total`;
      badge.className = "filter-badge";
    }
  }

  if (!tagsRow) return;
  clear(tagsRow);

  const chips = [];

  if (p.q) {
    chips.push(createChip(`Search: "${p.q}"`, () => {
      p.q = "";
      const qInput = $("ng-q");
      if (qInput) qInput.value = "";
      p.offset = 0;
      loadItems(wrap, pager);
    }));
  }

  if (p.status && p.status !== "all") {
    chips.push(createChip(`Status: ${p.status}`, () => {
      p.status = "all";
      const sel = $("ng-status");
      if (sel) sel.value = "all";
      p.offset = 0;
      loadItems(wrap, pager);
    }));
  }

  if (p.brand && p.brand !== "all") {
    chips.push(createChip(`Brand: ${p.brand}`, () => {
      p.brand = "all";
      const sel = $("ng-brand");
      if (sel) sel.value = "all";
      p.offset = 0;
      loadItems(wrap, pager);
    }));
  }

  if (p.filter && p.filter !== "all") {
    const fLabel = (FILTERS.find((x) => x[0] === p.filter) || [])[1] || p.filter;
    chips.push(createChip(`Filter: ${fLabel}`, () => {
      p.filter = "all";
      const sel = $("ng-filter");
      if (sel) sel.value = "all";
      p.offset = 0;
      loadItems(wrap, pager);
    }));
  }

  if (p.sort_col) {
    const colObj = ALL_ITEM_COLS.find((c) => c.fld === p.sort_col);
    const label = colObj ? colObj.label : p.sort_col;
    chips.push(createChip(`Sorted: ${label} (${p.sort_dir === "asc" ? "▲ Asc" : "▼ Desc"})`, () => {
      p.sort_col = "";
      p.sort_dir = "asc";
      p.sort = "sort";
      p.offset = 0;
      loadItems(wrap, pager);
    }, "chip-sort"));
  } else if (p.sort && p.sort !== "sort") {
    const sLabel = (SORTS.find((x) => x[0] === p.sort) || [])[1] || p.sort;
    chips.push(createChip(`Sorted: ${sLabel}`, () => {
      p.sort = "sort";
      p.sort_col = "";
      p.sort_dir = "asc";
      p.offset = 0;
      loadItems(wrap, pager);
    }, "chip-sort"));
  }

  if (chips.length) {
    tagsRow.style.display = "flex";
    tagsRow.append(
      el("span", { class: "active-tags-label", text: "Active filters:" }),
      ...chips,
      el("button", {
        type: "button",
        class: "clear-filters-btn",
        text: "Clear all filters",
        onclick: () => resetItemFilters(wrap, pager)
      })
    );
  } else {
    tagsRow.style.display = "none";
  }
}

function createChip(text, onRemove, extraCls = "") {
  const chip = el("span", { class: `filter-chip ${extraCls}` },
    el("span", { text }),
    el("button", { type: "button", class: "filter-chip-remove", text: "×", title: "Remove filter", onclick: onRemove })
  );
  return chip;
}

function resetItemFilters(wrap, pager) {
  const p = ng.page;
  p.q = "";
  p.status = "all";
  p.brand = "all";
  p.filter = "all";
  p.sort = "sort";
  p.sort_col = "";
  p.sort_dir = "asc";
  p.offset = 0;
  const qInput = $("ng-q");
  if (qInput) qInput.value = "";
  const statusSel = $("ng-status");
  if (statusSel) statusSel.value = "all";
  const brandSel = $("ng-brand");
  if (brandSel) brandSel.value = "all";
  const filterSel = $("ng-filter");
  if (filterSel) filterSel.value = "all";
  const sortSel = $("ng-sort");
  if (sortSel) sortSel.value = "sort";
  loadItems(wrap, pager);
}

function toggleColFilterMenu(th, type, options, currentValue, onSelect) {
  const existing = document.querySelector(".col-filter-menu");
  if (existing) {
    const isSame = existing._th === th;
    existing.remove();
    if (isSame) return;
  }
  const menu = el("div", { class: "col-filter-menu" });
  menu._th = th;
  const searchInput = el("input", { type: "search", placeholder: `Search ${type}…` });
  const list = el("div", { class: "col-filter-list" });
  
  const allLabel = el("label", { class: "col-filter-item" });
  const allRadio = el("input", { type: "radio", name: `col-${type}`, value: "all", checked: !currentValue || currentValue === "all" });
  allLabel.append(allRadio, el("span", { text: `All ${type === "brand" ? "brands" : "statuses"}` }));
  allRadio.addEventListener("change", () => { onSelect("all"); menu.remove(); });
  list.append(allLabel);

  const opts = (options || []).filter(Boolean);
  const items = [];
  opts.forEach((opt) => {
    const lbl = el("label", { class: "col-filter-item" });
    const rad = el("input", { type: "radio", name: `col-${type}`, value: opt, checked: currentValue === opt });
    lbl.append(rad, el("span", { text: opt }));
    rad.addEventListener("change", () => { onSelect(opt); menu.remove(); });
    list.append(lbl);
    items.push({ opt: opt.toLowerCase(), el: lbl });
  });

  searchInput.addEventListener("input", () => {
    const sq = searchInput.value.toLowerCase().trim();
    items.forEach((it) => {
      it.el.style.display = !sq || it.opt.includes(sq) ? "" : "none";
    });
  });

  const foot = el("div", { class: "col-filter-foot" });
  foot.append(
    el("button", { type: "button", class: "act ghost small-btn", text: "Reset", onclick: () => { onSelect("all"); menu.remove(); } }),
    el("button", { type: "button", class: "act small-btn", text: "Close", onclick: () => menu.remove() })
  );

  menu.append(searchInput, list, foot);
  document.body.appendChild(menu);

  const rect = th.getBoundingClientRect();
  menu.style.position = "fixed";
  menu.style.top = `${rect.bottom + 2}px`;
  menu.style.left = `${Math.max(10, Math.min(rect.left, window.innerWidth - 270))}px`;

  const closeHandler = (e) => {
    if (!menu.contains(e.target) && !th.contains(e.target)) {
      menu.remove();
      document.removeEventListener("pointerdown", closeHandler);
    }
  };
  setTimeout(() => document.addEventListener("pointerdown", closeHandler), 10);
  setTimeout(() => searchInput.focus(), 20);
}

// ---------- one item ----------
const EDIT_GROUPS = [
  ["rfq", ["brand", "catalog_no", "active_ingredient", "item_status", "remarks", "rfq_qty", "rfq_hna", "rfq_disc"]],
  ["current_mou", ["mou_qty", "mou_hna", "mou_disc"]],
  ["counter_offer", ["co_disc"]],
  ["feedback1", ["fb1_disc"]],
  ["online_nego", ["on_disc", "on_disc_nett", "on_disc_disc"]],
  ["analysis", ["proceed_bak", "bucket", "sbp", "need_sbp", "notes"]],
];
// The step that owns each field. Fields of a step the negotiation has passed are locked (a correction needs a reason).
const FIELD_OWNER = { mou_qty: "prepare", mou_hna: "prepare", mou_disc: "prepare", brand: "rfq", catalog_no: "rfq", active_ingredient: "rfq", item_status: "rfq", remarks: "rfq", rfq_qty: "rfq", rfq_hna: "rfq", rfq_disc: "rfq", co_disc: "counter_offer", co_note: "counter_offer", fb1_disc: "feedback1", on_disc: "online_nego", on_disc_nett: "online_nego", on_disc_disc: "online_nego" };
const FIELD_HEAD = { proceed_bak: "Lanjut ke BAK? (Ya/Tidak)", active_ingredient: "Active Ingredient", on_disc_nett: "ON Disc% (skema Nett)", on_disc_disc: "ON Disc% (skema Disc)", bucket: "Bucket (Consump/Charge)", sbp: "SBP? (Ya/Tidak)", need_sbp: "Perlu SBP setelah nego? (Ya/Tidak)", notes: "Notes", catalog_no: "Catalog No. (REF)" };
function fieldLocked(f) {
  const owner = FIELD_OWNER[f];
  if (!owner) return false;
  const keys = ng.ov.steps.map((s) => s.key);
  return keys.indexOf(ng.ov.cycle.current_step) > keys.indexOf(owner);
}

function renderItemMarginCard(it) {
  const card = el("div", { class: "card item-margin-card", style: "border: 1.5px solid #bfdbfe; background: #f8faff; padding: 16px; border-radius: 12px; margin: 16px 0;" });

  const header = el("div", { style: "display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 12px;" },
    el("div", {},
      el("h3", { style: "margin: 0; font-size: 15px; color: #0c21a4; display: flex; align-items: center; gap: 6px;", text: "📊 Calculated Margin Analysis" }),
      el("p", { class: "small muted", style: "margin: 3px 0 0 0;", text: "Margin per piece & annualized margin impact across negotiation steps" })),
    el("div", { style: "display: flex; gap: 6px; flex-wrap: wrap;" },
      el("span", { class: "badge", style: "background: #e0e7ff; color: #0c21a4; font-family: monospace; font-size: 11px;", text: "Base Price = HNA/pc × 1,11" }),
      el("span", { class: "badge", style: "background: #e0e7ff; color: #0c21a4; font-family: monospace; font-size: 11px;", text: "Margin/pc = Base Price - Harga/pc" }))
  );

  const baseMarginPrice = it.base_margin_price || (it.base_price ? it.base_price : null);
  const tiles = el("div", { class: "row3", style: "display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin-bottom: 14px;" },
    el("div", { class: "tile", style: "background: #fff; border: 1px solid #dbeafe; padding: 10px 14px; border-radius: 8px;" },
      el("div", { class: "label", style: "font-size: 11.5px; color: #475569;", text: "Base Price" }),
      el("div", { class: "value", style: "font-size: 18px; font-weight: 700; color: #0c21a4; margin-top: 2px;", text: unit(baseMarginPrice) }),
      el("div", { class: "note", style: "font-size: 11px; color: #64748b;", text: `HNA/pc ${unit(it.hna_per_pc)} + PPN · ${ng.ov.cycle.binding === "Nett" ? "dari HNA RFQ" : "dari HNA formularium (kosong: rata-rata PO)"}` })),
    el("div", { class: "tile", style: "background: #fff; border: 1px solid #dbeafe; padding: 10px 14px; border-radius: 8px;" },
      el("div", { class: "label", style: "font-size: 11.5px; color: #475569;", text: "Margin Baseline (Formularium)" }),
      el("div", { class: "value", style: "font-size: 18px; font-weight: 700; color: #0f172a; margin-top: 2px;", text: unit(it.margin_pc_formulary) }),
      el("div", { class: "note", style: "font-size: 11px; color: #64748b;", text: `Nett Price formularium: ${unit(it.form_nett_price)}/pc · acuan ${it.bench_tag}: ${unit(it.primary_bench)}/pc` })),
    el("div", { class: "tile", style: "background: #fff; border: 1px solid #dbeafe; padding: 10px 14px; border-radius: 8px;" },
      el("div", { class: "label", style: "font-size: 11.5px; color: #475569;", text: "Margin Baseline (rata-rata PO)" }),
      el("div", { class: "value", style: "font-size: 18px; font-weight: 700; color: #0f172a; margin-top: 2px;", text: unit(it.margin_pc_po) }),
      el("div", { class: "note", style: "font-size: 11px; color: #64748b;", text: `Rata-rata PO (+PPN): ${unit(it.po_unit_price)}/pc` }))
  );

  const stepRows = [
    { step: "RFQ", label: "RFQ (Penawaran Awal)", price: it.rfq_unit_price, margin: it.margin_pc_rfq, impactForm: it.margin_impact_rfq_vs_formulary, impactPo: it.margin_impact_rfq_vs_po },
    { step: "Counter Offer", label: "Counter Offer (Target RS)", price: it.co_unit_price, margin: it.margin_pc_co, impactForm: it.margin_impact_co_vs_formulary, impactPo: it.margin_impact_co_vs_po },
    { step: "Feedback I", label: "Feedback I (Revisi Prinsipal)", price: it.fb1_unit_price, margin: it.margin_pc_fb1, impactForm: it.margin_impact_fb1_vs_formulary, impactPo: it.margin_impact_fb1_vs_po },
    { step: "Online Nego", label: "Online Nego (Sepakat)", price: it.on_unit_price, margin: it.margin_pc_on, impactForm: it.margin_impact_on_vs_formulary, impactPo: it.margin_impact_on_vs_po },
  ];

  const stepTable = el("div", { class: "table-wrap", style: "margin-bottom: 12px;" },
    el("table", { style: "width: 100%; border-collapse: collapse; background: #fff; border-radius: 8px; overflow: hidden; border: 1px solid #e2e8f0; font-size: 12.5px;" },
      el("thead", {},
        el("tr", { style: "background: #f8fafc; border-bottom: 1px solid #e2e8f0;" },
          el("th", { style: "padding: 8px 10px; text-align: left; color: #0f172a;", text: "Step Negosiasi" }),
          el("th", { style: "padding: 8px 10px; text-align: right; color: #0f172a;", text: "Harga/pc" }),
          el("th", { style: "padding: 8px 10px; text-align: right; color: #0f172a;", text: "Margin/pc" }),
          el("th", { style: "padding: 8px 10px; text-align: right; color: #0f172a;", text: "Margin %" }),
          el("th", { style: "padding: 8px 10px; text-align: right; color: #0f172a;", text: "Margin Impact (vs Form)" }),
          el("th", { style: "padding: 8px 10px; text-align: right; color: #0f172a;", text: "Margin Impact (vs PO)" })
        )
      ),
      el("tbody", {},
        stepRows.map((r) => {
          const hasPrice = r.price != null;
          const marginPct = (hasPrice && baseMarginPrice && r.margin != null) ? (r.margin / baseMarginPrice) : null;
          return el("tr", { style: `border-bottom: 1px solid #f1f5f9; ${hasPrice ? "" : "opacity: 0.5;"}` },
            el("td", { style: "padding: 8px 10px; font-weight: 600; color: #0f172a;", text: r.label }),
            el("td", { style: "padding: 8px 10px; text-align: right; font-family: monospace; color: #0f172a;", text: unit(r.price) }),
            el("td", { style: `padding: 8px 10px; text-align: right; font-weight: 700; ${r.margin > 0 ? "color: #047857;" : r.margin < 0 ? "color: #b91c1c;" : ""}`, text: unit(r.margin) }),
            el("td", { style: "padding: 8px 10px; text-align: right; color: #1e293b;", text: pct(marginPct) }),
            el("td", { style: `padding: 8px 10px; text-align: right; font-weight: 600; ${r.impactForm > 0 ? "color: #047857;" : r.impactForm < 0 ? "color: #b91c1c;" : ""}`, text: r.impactForm == null ? "—" : `${r.impactForm > 0 ? "+" : ""}${money(r.impactForm)}` }),
            el("td", { style: `padding: 8px 10px; text-align: right; font-weight: 600; ${r.impactPo > 0 ? "color: #047857;" : r.impactPo < 0 ? "color: #b91c1c;" : ""}`, text: r.impactPo == null ? "—" : `${r.impactPo > 0 ? "+" : ""}${money(r.impactPo)}` })
          );
        })
      )
    )
  );

  const latestImpactForm = it.margin_impact_vs_formulary;
  const latestImpactPo = it.margin_impact_vs_po;
  const footerNote = el("div", { style: "display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 10px;" },
    el("div", { class: `preview ${latestImpactForm >= 0 ? "ok" : "bad"}`, style: "margin: 0; padding: 10px 12px; border-radius: 8px;" },
      el("b", { style: "display: block; font-size: 12px; margin-bottom: 2px;", text: "Versi 1: Margin Impact vs Formularium" }),
      el("div", { style: "font-size: 16px; font-weight: 700; color: " + (latestImpactForm >= 0 ? "#047857" : "#b91c1c") + ";", text: latestImpactForm == null ? "—" : `${latestImpactForm >= 0 ? "+" : ""}${money(latestImpactForm)}` }),
      el("p", { class: "small muted", style: "margin: 3px 0 0 0; font-size: 11px;", text: "Formula: (Qty Tahunan × Margin Step) - Margin Formularium saat ini" })),
    el("div", { class: `preview ${latestImpactPo >= 0 ? "ok" : "bad"}`, style: "margin: 0; padding: 10px 12px; border-radius: 8px;" },
      el("b", { style: "display: block; font-size: 12px; margin-bottom: 2px;", text: "Versi 2: Margin Impact vs PO" }),
      el("div", { style: "font-size: 16px; font-weight: 700; color: " + (latestImpactPo >= 0 ? "#047857" : "#b91c1c") + ";", text: latestImpactPo == null ? "—" : `${latestImpactPo >= 0 ? "+" : ""}${money(latestImpactPo)}` }),
      el("p", { class: "small muted", style: "margin: 3px 0 0 0; font-size: 11px;", text: "Formula: (Qty Tahunan × Margin Step) - Margin PO saat ini" }))
  );

  card.append(header, tiles, stepTable, footerNote);
  return card;
}

const BOOL_FIELDS = new Set(["sbp", "need_sbp", "proceed_bak"]);
const SELECT_FIELDS = { item_status: ["", "Active", "Discontinue", "New Item"], bucket: ["", "Consump", "Charge"] };

async function openItem(id) {
  const body = openPanel(el("p", { class: "muted", text: "Loading…" }));
  let it;
  try { it = await api(`/api/nego/cycles/${ng.cid}/items/${id}`); } catch (e) { clear(body).append(el("p", { class: "error", text: e.message })); return; }
  const cols = ng.ov.columns.reduce((m, c) => { m[c.field] = c; return m; }, {});
  const acuan = it.primary_bench;
  const tag = it.bench_tag || "acuan";
  const prices = el("div", { class: "pricegrid" },
    [["MOU", it.mou_unit_price], ["RFQ", it.rfq_unit_price], ["Counter offer", it.co_unit_price], ["Feedback I", it.fb1_unit_price], ["Online nego", it.on_unit_price], ["Rata-rata PO", it.po_unit_price]]
      .filter(([, v]) => v != null).map(([l, v]) => el("div", { class: l === "Online nego" ? "t" : null }, el("span", { text: `${l} per piece` }), el("b", { text: unit(v) }),
        acuan && !["MOU", "Rata-rata PO"].includes(l) ? el("small", { text: `${v / acuan - 1 > 0 ? "+" : ""}${pct(v / acuan - 1)} vs ${tag}` }) : null)));
  const facts = el("div", { class: "facts" },
    [["Satuan PO", it.po_unit_text || "—"], ["Pieces / PO unit", it.po_unit_pcs ? num(it.po_unit_pcs) : "—"], ["Qty PO tahun ini", it.po_qty_12m ? num(it.po_qty_12m) : "—"],
      ["Nilai PO setahun (+PPN)", it.po_annualized_value_ppn ? money(it.po_annualized_value_ppn) : "—"], ["Pieces a year", it.annual_pcs ? num(it.annual_pcs) : "—"],
      ["Impact a year", it.impact == null ? "—" : `${it.impact > 0 ? "+" : ""}${money(it.impact)}`],
      ["Acuan harga", acuan ? `${tag}: ${unit(acuan)}` : "—"],
      ["Formularium", it.form_base_price != null ? `Base ${unit(it.form_base_price)} · Nett ${unit(it.form_nett_price)}` : "—"],
      ["Source", { po: "PO history", formulary: "Formulary (not bought)", mou: "MOU only", template: "Template" }[it.source] || it.source || "—"]]
      .map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: v }))));
  const marginBox = renderItemMarginCard(it);
  const findings = el("div", { class: "findings" });
  renderFindingList(findings, it.anomalies.filter((a) => a.status !== "cleared"), () => openItem(id), true);
  const reason = el("div", {},
    it.price_reason ? el("p", { class: "reason-note" }, el("b", { text: "Alasan principal: " }), ({ raw_material: 'Kenaikan harga bahan baku', fx: 'Perubahan nilai tukar (kurs)', production: 'Kenaikan biaya produksi / logistik', other: 'Lainnya' })[it.price_reason] || it.price_reason) : null,
    it.co_note ? el("p", { class: "reason-note co" }, el("b", { text: "Catatan counter offer: " }), it.co_note) : null,
    it.needs_check ? el("p", { class: "reason-note co" }, el("b", { text: "Dikembalikan ke principal untuk dicek: " }), it.check_note || "") : null,
    (it.comments || []).length ? el("div", { class: "reason-note" }, el("b", { text: "Komentar / pengajuan koreksi principal" }), el("ul", { class: "timeline" }, it.comments.map((m) => el("li", {}, m.text, el("div", { class: "small muted", text: `${m.by} · ${new Date(m.at).toLocaleString("en-GB")}${m.resolved ? " · selesai" : ""}` }))))) : null,
    it.benchmarks && it.benchmarks.length ? el("div", { class: "bm-list" }, el("b", { text: "Harga kompetitor (file benchmark periode)" }),
      it.benchmarks.map((m) => el("p", { class: "small" }, `${unit(m.price_pp)}/pcs · ${m.source}`, m.link ? el("a", { href: m.link, target: "_blank", rel: "noopener", text: " ↗ sumber" }) : null))) : null);

  // The form. Fields of a step already passed are locked: a correction needs a reason and is logged.
  const form = el("form", { class: "progress-form item-form" });
  const inputs = {};
  const anyLocked = Object.keys(FIELD_OWNER).some((f) => fieldLocked(f));
  const rsn = el("textarea", { rows: "2", placeholder: "Alasan koreksi (atas pengajuan principal)", name: "correction_reason" });
  const rsnBox = field("Alasan koreksi (wajib untuk data terkunci)", rsn, "Tercatat di riwayat item dan Activity.");
  rsnBox.hidden = true;
  const drugs = (ng.ov.cycle.category || "") === "Drugs";
  for (const [sec, fields0] of EDIT_GROUPS) {
    const fields = [...(sec === "rfq" && it.is_new ? ["erp_code", "item_name"] : []), ...fields0.filter((f) => !(f === "catalog_no" && drugs) && !(f === "active_ingredient" && !drugs))];
    const fs = el("fieldset", { class: SECTION_CLASS[sec] }, el("legend", { text: SECTION_SHORT[sec] }));
    const lockedHere = fields.some((f) => fieldLocked(f));
    if (lockedHere) fs.append(el("p", { class: "small muted", style: "margin-top: 0;", text: "🔒 Data step ini sudah terkunci karena negosiasi sudah melewatinya. Mengubahnya = koreksi: isi alasan di bawah." }));
    const grid = el("div", { class: "row3" });
    for (const f of fields) {
      const c = cols[f] || { header: FIELD_HEAD[f] || f, kind: "text" };
      const isPct = c.kind === "pct" || ["on_disc_nett", "on_disc_disc"].includes(f);
      let input;
      if (SELECT_FIELDS[f]) {
        input = el("select", { name: f }, SELECT_FIELDS[f].map((v) => el("option", { value: v, text: v || "—" })));
        input.value = it[f] || "";
      } else if (BOOL_FIELDS.has(f)) {
        input = el("select", { name: f }, [["", "—"], ["true", "Ya"], ["false", "Tidak"]].map(([v, l]) => el("option", { value: v, text: l })));
        input.value = it[f] == null ? "" : String(!!it[f]);
      } else {
        const v = it[f];
        const shown = v == null ? "" : isPct ? String(+(v * 100).toFixed(2)).replace(".", ",") : String(v);
        input = el("input", { name: f, value: shown, inputmode: c.kind === "text" ? null : "decimal", placeholder: isPct ? "0–100" : null });
      }
      if (!canWork()) input.setAttribute("disabled", "");
      inputs[f] = { input, pct: isPct, locked: fieldLocked(f) };
      input.addEventListener("input", () => { rsnBox.hidden = !Object.entries(inputs).some(([ff, x]) => x.locked && x.input.value.trim() !== (x.start ?? x.input.value.trim())); });
      inputs[f].start = input.value.trim();
      grid.append(field(isPct ? `${(c.header || FIELD_HEAD[f] || f).replace("%", "")} (%)` : (c.header || FIELD_HEAD[f] || f) + (inputs[f].locked ? " 🔒" : ""), input));
    }
    fs.append(grid);
    form.append(fs);
  }
  form.append(rsnBox);
  const msg = el("p", { class: "error", role: "alert", hidden: "" });
  if (canWork()) form.append(msg, el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Save changes" }),
    el("span", { class: "small muted", text: "Setiap perubahan dicatat: siapa, kapan, nilai lama → baru." })));
  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const changes = {};
    for (const [f, x] of Object.entries(inputs)) {
      const raw = x.input.value.trim();
      if (raw === x.start) continue;
      changes[f] = raw === "" ? null : raw;            // the server reads the number as typed (0 to 100 for a discount)
    }
    if (!Object.keys(changes).length) { msg.textContent = "Nothing changed."; msg.hidden = false; return; }
    const lockedHit = Object.keys(changes).some((f) => inputs[f].locked);
    if (lockedHit && !rsn.value.trim()) { rsnBox.hidden = false; msg.textContent = "Isi alasan koreksi untuk data yang terkunci."; msg.hidden = false; return; }
    try {
      const r = await send(`/api/admin/nego/cycles/${ng.cid}/items/${id}`, "PATCH", { changes, base_version: it.version, correction_reason: rsn.value.trim() || undefined });
      toast(`Saved ${r.changed.length} change${r.changed.length === 1 ? "" : "s"}`);
      await refreshOverview();
      if (ng.tab === "items") loadItems(document.querySelector(".items-wrap"), document.querySelector(".pager"));
      openItem(id);
    } catch (ex) {
      msg.textContent = ex.message; msg.hidden = false;
      if (ex.status === 409) msg.append(" ", el("button", { type: "button", class: "act ghost small-btn", text: "Muat ulang nilai terbaru", onclick: () => openItem(id) }));
    }
  });
  const fmtH = (v) => (v == null ? "—" : typeof v === "number" && Math.abs(v) < 1 && v !== 0 ? `${+(v * 100).toFixed(2)}%` : String(v));
  const hist = it.history.length ? el("ul", { class: "timeline" }, it.history.slice(0, 40).map((h) =>
    el("li", {}, el("b", { text: h.label || (cols[h.field] || {}).header || h.field }), ` ${fmtH(h.from)} → ${fmtH(h.to)}`,
      el("div", { class: "small muted", text: `${h.by} · ${h.source}${h.reason ? ` · alasan: ${h.reason}` : ""} · ${new Date(h.at).toLocaleString("en-GB")}` })))) : el("p", { class: "muted small", text: "No changes yet." });
  clear(body).append(
    el("p", { class: "eyebrow", text: it.erp_code || "No ERP code" }),
    el("h2", { id: "drawer-title", text: it.item_name || "Unnamed item" }),
    el("p", { class: "muted small", text: [it.brand, drugs ? it.active_ingredient : it.catalog_no, it.item_status, it.bucket].filter(Boolean).join(" · ") }),
    it.is_new && it.item_status === "New Item" ? el("p", { class: "reason-note co" }, el("b", { text: "Item baru dari principal. " }), "Isi ERP code resmi di bawah (sekarang kode sementara), lalu setujui lewat Findings atau ubah status ke Active.") : null,
    prices, facts, marginBox, reason,
    el("h3", { text: "Findings" }), findings,
    el("h3", { text: canWork() ? "Edit" : "Values" }), anyLocked ? el("p", { class: "small muted", text: "Data step yang sudah lewat terkunci. Koreksi hanya atas pengajuan principal." }) : null, form,
    el("h3", { text: "History" }), hist);
}

function drawUomResult(box, res, it, after) {
  clear(box);
  const sci = res.science_profile || {};
  const links = (res.grounding_links || []).map((l) => el("a", { class: "small", style: "display: inline-block; margin-right: 8px;", href: l.uri, target: "_blank", rel: "noopener", text: l.title || l.uri }));
  const structImg = sci.structure_png ? el("img", { src: sci.structure_png, alt: sci.generic_name, style: "max-width: 130px; border-radius: 6px; border: 1px solid #cbd5e1; background: #fff; padding: 4px;" }) : null;
  const applyBtn = act("Apply UOM & Generic to Item", async () => {
    try {
      await send("/api/science/uom-standardize/apply", "POST", {
        cycle_id: ng.cid,
        item_id: it.id,
        erp_code: it.erp_code,
        pcs_per_po_unit: res.pcs_per_po_unit,
        standardized_po_unit_text: res.standardized_po_unit_text,
        generic_name: res.generic_name,
        group_key: res.group_key,
        clinical_unit: res.clinical_unit,
      });
      toast(`Applied: ${res.pcs_per_po_unit} ${res.clinical_unit}/PO unit · ${res.generic_name}`);
      await refreshOverview();
      after();
    } catch (e) { toast(e.message); }
  });

  put(box,
    el("div", { class: "preview ok", style: "margin-top: 10px; padding: 12px;" },
      el("div", { style: "display: flex; gap: 12px; align-items: flex-start;" },
        structImg,
        el("div", { style: "flex: 1;" },
          el("b", { text: res.standardized_uom_label }),
          el("div", { class: "small muted", style: "margin-top: 3px;", text: `Standardized generic: ${res.generic_name} · Group: ${res.group_key}` }),
          el("p", { class: "small", style: "margin: 6px 0;", text: res.rationale }),
          sci.pubchem_cid ? el("div", { class: "small muted", text: `PubChem CID ${sci.pubchem_cid} · Formula ${sci.molecular_formula || "—"} · Weight ${sci.molecular_weight || "—"}` }) : null,
          sci.chembl_id ? el("div", { class: "small muted", text: `ChEMBL ${sci.chembl_id} · ${sci.molecule_type || ""} · Max Phase ${sci.max_phase ?? "N/A"}` }) : null,
          sci.fda_dosage_forms && sci.fda_dosage_forms.length ? el("div", { class: "small muted", text: `openFDA Dosage Forms: ${sci.fda_dosage_forms.join(", ")}` }) : null,
          sci.fda_sample_packages && sci.fda_sample_packages.length ? el("div", { class: "small muted", text: `openFDA Packaging: ${sci.fda_sample_packages.slice(0, 2).join(" · ")}` }) : null)),
      links.length ? el("div", { style: "margin-top: 8px;" }, el("b", { class: "small", text: "Google Search Grounding Sources: " }), ...links) : null,
      canWork() ? el("div", { class: "actions", style: "margin-top: 10px;" }, applyBtn) : null));
}

function drawBenchmarkResult(box, res, it, after) {
  clear(box);
  const cands = res.benchmark_candidates || [];
  const sci = res.science_profile || {};
  const links = (res.grounding_links || []).map((l) => el("a", { class: "small", style: "display: inline-block; margin-right: 8px;", href: l.uri, target: "_blank", rel: "noopener", text: l.title || l.uri }));

  const t = el("table", { class: "bm-table", style: "margin-top: 8px; width: 100%;" });
  table(t, [
    { label: "Source", get: (c) => el("b", { text: c.source }) },
    { label: "Specification", get: (c) => el("div", {}, c.description, el("div", { class: "small muted", text: `Pack: ${c.pack_uom} (${c.pcs_per_pack} pcs)` })) },
    { label: "Pack Price", num: true, get: (c) => unit(c.pack_price_idr) },
    { label: "Price/pc", num: true, get: (c) => el("b", { class: "good", text: unit(c.price_pp) }) },
    { label: "", get: (c) => canWork() ? act("Confirm benchmark", async () => {
      try {
        await send("/api/science/sku-benchmark/apply", "POST", {
          cycle_id: ng.cid,
          item_id: it.id,
          source: c.source,
          bm_name: c.description,
          bm_unit: c.pack_uom,
          price_pp: c.price_pp,
          confidence: c.confidence,
          pcs_per_po_unit: c.pcs_per_pack,
          generic_name: res.science_profile?.generic_name,
        });
        toast(`Benchmark confirmed: ${unit(c.price_pp)}/pc (${c.source})`);
        await refreshOverview();
        after();
      } catch (e) { toast(e.message); }
    }, "ghost") : "" },
  ], cands);

  put(box,
    el("div", { class: "preview ok", style: "margin-top: 10px; padding: 12px;" },
      el("h4", { style: "margin: 0 0 6px 0; color: #047857;", text: `Recommended Benchmark: ${unit(res.recommended_bench_pp)} per clinical ${res.uom_standardization?.clinical_unit || "piece"}` }),
      res.saving_vs_current_pct ? el("p", { class: "small good", text: `Saves ${pct(res.saving_vs_current_pct, 1)} vs current Siloam price (${unit(res.current_price_pp)}/pc)` }) : null,
      el("p", { class: "small", text: res.negotiation_insight }),
      sci.fda_manufacturers_count ? el("p", { class: "small muted", text: `Bioequivalence leverage: ${sci.fda_manufacturers_count} FDA-registered generic manufacturers for ${sci.generic_name}. Strong multi-source competitive pressure.` }) : null,
      el("div", { class: "table-wrap" }, t),
      links.length ? el("div", { style: "margin-top: 8px;" }, el("b", { class: "small", text: "Google Search Grounding Sources: " }), ...links) : null));
}

// ---------- findings ----------
async function renderFindings(body) {
  clear(body);
  const status = el("select", { id: "ng-fstatus" }, [["open", "Open"], ["kept", "Kept as is"], ["fixed", "Fixed"], ["cleared", "Cleared by a rescan"], ["", "All"]].map(([k, l]) => el("option", { value: k, text: l })));
  status.value = ng.findStatus;
  const chips = el("div", { class: "chips" });
  const list = el("div", { class: "findings" });
  body.append(el("div", { class: "filters board-filters" }, el("label", {}, "Status", status),
    canWork() ? act("Rescan now", async () => {
      try { const r = await send(`/api/admin/nego/cycles/${ng.cid}/scan`, "POST", {}); toast(`${r.found} findings · ${r.new} new · ${r.cleared} cleared`); await refreshOverview(); renderFindings(body); } catch (e) { toast(e.message); }
    }, "secondary") : null,
    el("p", { class: "small muted grow", text: "Temuan hanya anomali: harga naik >50%, harga RFQ turun >0,5%, atau format diskon salah, dibanding Nett Price formularium (kosong: MOU, lalu rata-rata PO). Fix memakai nilai saran, Keep as is perlu alasan, Kembalikan ke principal menandai SKU untuk dicek." })),
  chips, list);
  status.addEventListener("change", () => { ng.findStatus = status.value; ng.rule = ""; renderFindings(body); });
  let d;
  try { d = await api(`/api/nego/cycles/${ng.cid}/anomalies?status=${encodeURIComponent(ng.findStatus)}`); } catch (e) { list.append(el("p", { class: "error", text: e.message })); return; }
  const draw = () => {
    clear(chips).append(el("button", { type: "button", class: `chipbtn ${ng.rule ? "" : "on"}`, text: `All (${d.anomalies.length})`, onclick: () => { ng.rule = ""; draw(); } }),
      ...Object.entries(d.by_rule).sort((a, b) => b[1] - a[1]).map(([label, n]) =>
        el("button", { type: "button", class: `chipbtn ${ng.rule === label ? "on" : ""}`, text: `${label} (${n})`, onclick: () => { ng.rule = label; draw(); } })));
    const rows = ng.rule ? d.anomalies.filter((a) => a.label === ng.rule) : d.anomalies;
    renderFindingList(list, rows, () => renderFindings(body));
  };
  draw();
}

function renderFindingList(box, rows, after, inDrawer = false) {
  clear(box);
  if (!rows.length) { box.append(el("p", { class: "muted small", text: inDrawer ? "Nothing flagged for this item." : "Nothing here." })); return; }
  const shown = rows.slice(0, 200);
  for (const a of shown) {
    const reason = el("input", { placeholder: "Why keep it? e.g. confirmed with principal", "aria-label": "Reason" });
    const keepBox = el("div", { class: "keep", hidden: "" }, reason,
      act("Keep as is", async () => { await decideFinding(a, "keep", reason.value, after); }),
      act("Cancel", () => { keepBox.hidden = true; }, "ghost"));
    const actions = canWork() ? el("div", { class: "actions" },
      a.status === "open" && a.suggestion != null ? act(a.suggestion_text || "Apply fix", () => decideFinding(a, "fix", "", after)) : null,
      a.status === "open" ? act(a.suggestion != null ? "Mark fixed" : "Fixed it", () => decideFinding(a, "fix", "", after), a.suggestion != null ? "ghost" : "secondary") : null,
      a.status === "open" ? act("Keep as is…", () => { keepBox.hidden = false; reason.focus(); }, "ghost") : null,
      a.status === "open" && a.item_id && a.severity === "high" && !["principal_comment"].includes(a.rule) ? act("Kembalikan ke principal", () => decideFinding(a, "return", a.suggestion_text || "", after), "ghost") : null,
      a.status === "kept" || a.status === "fixed" ? act("Reopen", () => decideFinding(a, "reopen", "", after), "ghost") : null,
      !inDrawer && a.item_id ? act("Open item", () => openItem(a.item_id), "ghost") : null) :
      (!inDrawer && a.item_id ? el("div", { class: "actions" }, act("Open item", () => openItem(a.item_id), "ghost")) : null);
    box.append(el("article", { class: `finding ${a.severity} st-${a.status}` },
      el("div", { class: "f-top" }, el("span", { class: `sev ${a.severity}`, text: a.severity }), el("b", { text: a.label }),
        !inDrawer ? el("span", { class: "small muted", text: `${a.erp_code || ""} ${a.item_name || ""}` }) : null,
        a.status !== "open" ? el("span", { class: "small status", text: { kept: "Kept as is", fixed: "Fixed", cleared: "Cleared" }[a.status] }) : null),
      el("p", { text: a.message }),
      a.principal_reason || (a.principal_confirmed && a.principal_confirmed.length) || a.principal_comment ? el("p", { class: "small reason-note" }, el("b", { text: "Jawaban principal: " }), [a.principal_reason ? `alasan ${({ raw_material: "bahan baku", fx: "kurs", production: "biaya produksi/logistik", other: "lainnya" })[a.principal_reason] || a.principal_reason}` : null, (a.principal_confirmed || []).length ? `konfirmasi "Ya, sudah benar" (${a.principal_confirmed.join(", ")})` : null, a.principal_comment ? `komentar: ${a.principal_comment}` : null].filter(Boolean).join(" · ")) : null,
      a.reason ? el("p", { class: "small muted", text: `Reason: ${a.reason} — ${a.decided_by}` }) : null,
      actions, keepBox));
  }
  if (rows.length > shown.length) box.append(el("p", { class: "muted small", text: `Showing the first ${shown.length} of ${rows.length}. Filter by rule to see the rest.` }));
}

async function decideFinding(a, decision, reason, after) {
  try {
    await send(`/api/admin/nego/cycles/${ng.cid}/anomalies/${a.id}`, "POST", { decision, reason });
    toast(decision === "fix" ? "Fixed" : decision === "keep" ? "Kept, reason logged" : decision === "return" ? "SKU ditandai untuk dicek principal" : "Reopened");
    await refreshOverview();
    after();
  } catch (e) { toast(e.message); }
}

// ---------- files ----------
function renderFiles(body) {
  clear(body);
  const c = ng.ov.cycle;
  const ex = card("Template_Nego", "The workbook the principal knows: same layout, colours, drop-downs and formulas, filled with this negotiation. Rows grow past 150 when needed.",
    el("div", { class: "actions" }, el("a", { class: "act", href: `/api/nego/cycles/${c.id}/export.xlsx`, text: "Download Template_Nego" })));
  body.append(ex);
  if (!canWork()) return;

  // PO lines, formulary, SKU master and tagging come from the period's data hub (Admin, uploaded once per
  // period). Only the principal's current MOU / last Template_Nego is uploaded here.
  const per = ng.ov.period;
  const hub = per ? el("div", { class: `preview ${per.po_rows || per.formulary ? "ok" : "bad"}` },
    el("b", { text: `Data for ${per.name}` }),
    el("div", { class: "facts" }, [["PO lines", per.po_rows], ["Formulary", per.formulary], ["SKU tagging", per.sku_tagging]].map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: num(v) })))),
    el("p", { class: "small muted", text: per.po_rows || per.formulary ? "Loaded once in Admin → Negotiation Periods; every principal of this category reads from it." : "Nothing uploaded for this period yet." }),
    isAdmin() ? el("a", { class: "act ghost small-btn", href: "#admin/periods", text: "Open the period data hub →" }) : null) : null;
  const files = { mou: el("input", { type: "file", accept: ".csv,.xlsx", class: "sr-only", name: "mou" }) };
  const prepOut = el("div", { role: "status" });
  const prep = el("form", { class: "upload-form wide" },
    hub,
    el("div", { class: "drop3" }, dropZone(files.mou, ng.ov.has_mou_file ? "Current MOU / last Template_Nego (on file, drop a new one to replace)" : "Current MOU / last Template_Nego")),
    el("p", { class: "small muted", text: "CSV or Excel. Columns: ERP Code, Brand, Catalog No. (REF), MOU_Qty/PO Unit, MOU_HNA/PO Unit (excl. PPN), MOU_Disc%. Building again replaces the item list and its findings." }),
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Build item list" })), prepOut);
  prep.addEventListener("submit", async (e) => {
    e.preventDefault();
    const fd = new FormData();
    Object.entries(files).forEach(([k, f]) => { if (f.files.length) fd.append(k, f.files[0]); });
    if (ng.ov.kpis.items && !confirm("Replace the current item list and its findings?")) return;
    clear(prepOut).append(el("p", { class: "muted", text: "Building…" }));
    try {
      const r = await upload(`/api/admin/nego/cycles/${c.id}/prepare`, fd);
      clear(prepOut).append(prepSummary(r));
      await refreshOverview();
      toast(`${r.items} items · ${r.scan.found} findings`);
    } catch (ex2) { clear(prepOut).append(el("p", { class: "error", text: ex2.message })); }
  });
  const prevSummary = c.prepare_summary ? prepSummary(c.prepare_summary) : null;
  body.append(card("Prepare: build the item list", "Unique items bought from this principal in the last 12 months, plus formulary items not bought, with the current MOU prices.", prep, prevSummary ? el("details", {}, el("summary", { text: `Last build: ${new Date(c.prepared_at).toLocaleString("en-GB")}` }), prevSummary) : null));

  const tfile = el("input", { type: "file", accept: ".xlsx", class: "sr-only", name: "file" });
  const impOut = el("div", { role: "status" });
  const imp = el("form", { class: "upload-form wide" }, dropZone(tfile, "Filled Template_Nego (.xlsx)"),
    el("p", { class: "small muted", text: "Matched by ERP code, then item name. Blank cells never erase a value. You see every change before it's applied." }),
    el("div", { class: "actions" }, el("button", { type: "submit", class: "act", text: "Check file" })), impOut);
  imp.addEventListener("submit", async (e) => {
    e.preventDefault();
    if (!tfile.files.length) { clear(impOut).append(el("p", { class: "error", text: "Choose the file first." })); return; }
    const fd = new FormData(); fd.append("file", tfile.files[0]); fd.append("apply", "0");
    clear(impOut).append(el("p", { class: "muted", text: "Checking…" }));
    try { const plan = await upload(`/api/admin/nego/cycles/${c.id}/import`, fd); clear(impOut).append(importPreview(plan, tfile.files[0])); }
    catch (ex2) { clear(impOut).append(el("p", { class: "error", text: ex2.message })); }
  });
  body.append(card("Import a filled template", "The principal's reply by email, or values typed offline. Identification, RFQ and discount columns are read; calculated columns are recomputed.", imp));
}

function prepSummary(r) {
  const po = r.po || {};
  return el("div", { class: "preview ok" },
    el("div", { class: "facts" },
      [["Items", num(r.items)], ["From POs", num((r.by_source || {}).po || 0)], ["Formulary only", num((r.by_source || {}).formulary || 0)], ["MOU only", num((r.by_source || {}).mou || 0)],
        po.period_from ? ["PO period", `${fmtDate(po.period_from)} – ${fmtDate(po.period_to)}`] : null,
        po.po_lines_used != null ? ["PO lines used", `${num(po.po_lines_used)} of ${num(po.po_lines)}`] : null,
        r.scan ? ["Findings", num(r.scan.found)] : null].filter(Boolean).map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: v })))),
    r.several_po_units && r.several_po_units.length ? el("p", { class: "small warn-text", text: `Bought in more than one PO unit: ${r.several_po_units.slice(0, 8).join(", ")}${r.several_po_units.length > 8 ? "…" : ""}` }) : null,
    r.mou_errors && r.mou_errors.length ? el("div", { class: "preview bad" }, el("b", { text: `${r.mou_errors.length} baris MOU ditolak` }), el("ul", { class: "errors" }, r.mou_errors.slice(0, 30).map((x) => el("li", { text: `Baris ${x.row} ${x.code}: ${x.message}` })))) : null,
    r.columns_used ? el("details", {}, el("summary", { text: "Columns used" }), el("ul", { class: "small" },
      Object.entries(r.columns_used).map(([k, m]) => el("li", { text: `${{ po: "PO export", formulary: "Formulary", mou: "MOU" }[k] || k}: ${Object.entries(m).map(([a, b]) => `${a} ← ${b}`).join(", ")}` })))) : null);
}

function importPreview(plan, file) {
  const bad = plan.warning || plan.errors.length;
  const box = el("div", { class: `preview ${bad ? "bad" : "ok"}` },
    el("h3", { text: plan.changes.length ? `${num(plan.items_changed)} items would change` : "No changes found" }),
    el("div", { class: "facts" }, [["Rows in file", plan.rows_in_file], ["Berhasil dibaca", plan.ok_rows ?? plan.matched], ["Ditolak / ada error", plan.rejected_rows ?? plan.errors.length], ["Not in this negotiation", plan.unmatched_count], ["Items not in the file", plan.not_in_file]]
      .map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: num(v) })))),
    plan.warning ? el("p", { class: "error", text: plan.warning }) : null,
    Object.keys(plan.changes_by_column).length ? el("p", { class: "small", text: `Changes by column: ${Object.entries(plan.changes_by_column).map(([k, v]) => `${k} ${v}`).join(" · ")}` }) : null,
    plan.errors.length ? el("details", { open: "" }, el("summary", { text: `${plan.errors.length} baris dengan error (sel yang benar tetap dihitung)` }),
      el("ul", { class: "errors" }, plan.errors.slice(0, 60).map((x) => el("li", { text: `Baris ${x.row} ${x.erp_code || ""}: ${x.message}` })))) : null,
    plan.new_items && plan.new_items.length ? el("div", { class: "preview ok" }, el("b", { text: `${plan.new_items.length} item baru akan ditambahkan (status New Item, masuk Findings untuk dicek)` }), el("ul", { class: "small" }, plan.new_items.map((n) => el("li", { text: `Baris ${n.row}: ${n.item_name}` })))) : null,
    plan.unmatched.length ? el("p", { class: "small muted", text: `Not matched: ${plan.unmatched.slice(0, 6).map((u) => u.erp_code || u.item_name).join(", ")}${plan.unmatched_count > 6 ? "…" : ""}` }) : null);
  const reason = el("textarea", { rows: "2", placeholder: "Alasan koreksi (atas pengajuan principal)" });
  if (plan.locked_count) box.append(el("div", { class: "preview bad" },
    el("p", { class: "error", text: `${plan.locked_count} perubahan menyentuh data step yang sudah terkunci (${(plan.locked_columns || []).join(", ")}). Ini koreksi: hanya atas pengajuan principal, tercatat di riwayat item.` }),
    field("Alasan koreksi (wajib)", reason)));
  if (plan.changes.length || (plan.new_items && plan.new_items.length)) {
    const t = el("table");
    table(t, [{ label: "Baris", get: (c) => c.row }, { label: "ERP", get: (c) => c.erp_code }, { label: "Item", get: (c) => c.item_name }, { label: "Column", get: (c) => c.header },
      { label: "Now", get: (c) => fmtAny(c.field, c.old) }, { label: "File", get: (c) => fmtAny(c.field, c.new) }], plan.changes.slice(0, 100));
    box.append(el("details", { open: "" }, el("summary", { text: `Changes${plan.changes.length > 100 ? " (first 100)" : ""}` }), el("div", { class: "table-wrap" }, t)));
    if (!plan.warning) box.append(el("div", { class: "actions" }, act("Apply changes", async () => {
      if (plan.locked_count && !reason.value.trim()) { toast("Isi alasan koreksi untuk data yang terkunci."); return; }
      const fd = new FormData(); fd.append("file", file); fd.append("apply", "1"); if (reason.value.trim()) fd.append("correction_reason", reason.value.trim());
      try {
        const r = await upload(`/api/admin/nego/cycles/${ng.cid}/import`, fd);
        toast(`${r.applied_items} items updated · ${r.applied_changes} values${r.rejected_rows ? ` · ${r.rejected_rows} baris ditolak` : ""}`);
        await refreshOverview();
        renderFiles($("ng-body"));
      } catch (e) { toast(e.message); }
    })));
  }
  return box;
}

function fmtAny(f, v) {
  if (v == null) return "—";
  if (/_disc/.test(f)) return pct(v, 2);
  if (typeof v === "number") return num(v, v % 1 ? 2 : 0);
  return String(v);
}

// ---------- activity ----------
const EVENT_TEXT = {
  "cycle.create": "Negotiation opened", "cycle.prepare": "Item list built", "cycle.update": "Details changed", "cycle.step": "Step changed", "cycle.due": "Batas waktu diubah",
  "cycle.import": "Template Excel diimpor", "item.edit": "Item diubah", "item.correction": "Koreksi data terkunci", "co.apply": "Counter offer diterapkan", "on.fill": "Online Nego diisi dari FB1/CO",
  "finding.decide": "Keputusan temuan", "meeting.save": "Meeting disimpan", "scheme.final": "Skema final dipilih", "escalation.approve": "Eskalasi disetujui", "escalation.revoke": "Persetujuan eskalasi dicabut",
  "bak.flags": "Tanda lanjut ke BAK", "package.excel": "Excel Confirmation diunduh", "package.bak": "Draft BAK diunduh", "package.pks": "Draft PKS diunduh",
  "principal.submit": "Principal sent the step", "principal.excel": "Principal uploaded Excel", "principal.comment": "Principal mengajukan koreksi", "principal.document": "Principal mengunggah dokumen",
  "principal.accept_all": "Principal menerima counter offer", "principal.discontinue_confirmed": "Principal mengonfirmasi SKU Discontinue",
  "link.create": "Principal link created", "link.revoke": "Principal link revoked", "link.open": "Principal opened the link",
  "message.queued": "Link queued for the principal", "message.sent": "Message delivered to Power Automate", "message.failed": "Message failed", "message.skipped": "Message not sent",
};
function eventExtra(e) {
  const d = e.detail || {};
  const stepLabel = (k) => (ng.ov.steps.find((s) => s.key === k) || {}).label || k;
  switch (e.action) {
    case "cycle.step": return `${stepLabel(d.from)} → ${stepLabel(d.to)}${d.message ? ` · ${d.message}` : ""}${d.flagged ? ` · ${d.flagged} SKU ditandai` : ""}`;
    case "cycle.prepare": return `${d.items} items`;
    case "cycle.import": return `${d.items} items, ${d.changes} values${d.reason ? ` · alasan: ${d.reason}` : ""}`;
    case "item.edit": case "item.correction": return `${d.erp_code}: ${(d.fields || []).join(", ")}${d.reason ? ` · alasan: ${d.reason}` : ""}`;
    case "co.apply": return `${d.applied} SKU`;
    case "finding.decide": return `${d.erp_code} · ${d.decision}`;
    case "principal.submit": return `${stepLabel(d.step)} → ${stepLabel(d.next_step)}`;
    case "principal.comment": return `${d.erp_code}: ${d.text}`;
    case "meeting.save": return d.at || "";
    case "scheme.final": return d.scheme || "";
    case "escalation.approve": return d.note || "";
    default: return d.to || d.error || d.reason || (d.rows != null ? `${d.rows} baris` : "");
  }
}
async function renderMessages(box) {
  let d;
  try { d = await api(`/api/admin/nego/cycles/${ng.cid}/messages`); } catch (e) { box.append(el("p", { class: "error", text: e.message })); return; }
  const a = d.automation;
  clear(box).append(el("div", { class: "card msgs" },
    el("div", { class: "msg-head" }, el("h3", { text: "Messages to the principal" }),
      el("span", { class: `pill ${a.enabled ? "good" : ""}`, text: a.enabled ? `Automatic sending on · ${a.webhook_host}` : "Automatic sending off" }),
      canWork() && ng.ov.cycle.status === "open" && PRINCIPAL_STEPS.has(ng.ov.cycle.current_step) ? act("Send link now", async () => {
        try { const r = await send(`/api/admin/nego/cycles/${ng.cid}/links/send`, "POST", {}); toast(r.status === "queued" ? `Link sent to ${r.recipient}` : "Not sent: see the message list"); await refreshOverview(); renderActivity($("ng-body")); } catch (e) { toast(e.message); }
      }, "secondary") : null),
    d.messages.length ? el("table", { class: "msg-table" }, el("thead", {}, el("tr", {}, ["When", "Message", "To", "Status", ""].map((h) => el("th", { text: h })))),
      el("tbody", {}, d.messages.map((m) => el("tr", {},
        el("td", { text: new Date(m.created_at).toLocaleString("en-GB") }),
        el("td", { text: { "principal.step_opened": "Link for the step", "siloam.step_submitted": "Notice to Siloam: principal sent", "siloam.mou_alert": "MOU alert" }[m.event] || m.event }),
        el("td", { text: m.recipient || "—" }),
        el("td", {}, el("span", { class: `mstat s-${m.status}`, text: { queued: "Queued", sending: "Sending", sent: "Sent", failed: `Failed (${m.attempts} tries)`, skipped: "Not sent" }[m.status] || m.status }),
          m.last_error ? el("div", { class: "small muted", text: m.last_error }) : null),
        el("td", {}, canWork() && m.status === "failed" ? act("Retry", async () => {
          try { await send(`/api/admin/nego/cycles/${ng.cid}/messages/${m.id}/retry`, "POST", {}); toast("Queued again"); setTimeout(() => renderActivity($("ng-body")), 800); } catch (e) { toast(e.message); }
        }, "ghost") : null))))) : el("p", { class: "muted small", text: "No messages yet. They're sent when the negotiation moves to a principal step." })));
}

function renderActivity(body) {
  clear(body);
  const msgBox = el("div");
  body.append(msgBox);
  renderMessages(msgBox);
  const ev = ng.ov.events;
  if (!ev.length) { body.append(el("p", { class: "muted", text: "Nothing yet." })); return; }
  body.append(el("h3", { class: "act-title", text: "Activity (Siloam dan principal)" }));
  body.append(el("div", { class: "card" }, el("ul", { class: "timeline" }, ev.map((e) => {
    const extra = eventExtra(e);
    return el("li", {}, el("b", { text: EVENT_TEXT[e.action] || e.action }), extra ? ` · ${extra}` : "",
      el("div", { class: "small muted", text: `${e.user_email} · ${new Date(e.at).toLocaleString("en-GB")}` }));
  }))));
}

// ---------- principal link ----------
async function linkPanel(c) {
  const body = openPanel(el("p", { class: "eyebrow", text: "Principal link" }), el("h2", { id: "drawer-title", text: c.principal.name }));
  const days = el("input", { type: "number", min: 1, max: 90, placeholder: "ikut Engine settings", id: "link-days" });
  const out = el("div", { role: "status" });
  const list = el("div");
  const draw = (links) => {
    clear(list).append(links.length ? el("ul", { class: "timeline" }, links.map((l) => el("li", {},
      el("b", { text: l.revoked_at ? "Revoked" : l.expires_at < new Date().toISOString() ? "Expired" : "Active" }),
      ` · created ${new Date(l.created_at).toLocaleString("en-GB")} by ${l.created_by}`,
      el("div", { class: "small muted", text: `Valid until ${new Date(l.expires_at).toLocaleString("en-GB")}${l.last_used_at ? ` · last opened ${new Date(l.last_used_at).toLocaleString("en-GB")}` : " · not opened yet"}` }),
      el("div", { style: "margin-top: 6px; display: flex; gap: 8px; align-items: center;" },
        !l.revoked_at ? el("a", { class: "act small secondary", href: l.path || l.url, target: "_blank", rel: "noopener", text: "Open & Test Display ↗" }) : null,
        !l.revoked_at ? act("Revoke", async () => { try { draw((await send(`/api/admin/nego/cycles/${c.id}/links/${l.id}/revoke`, "POST", {})).links); } catch (e) { toast(e.message); } }, "ghost") : null)))) :
      el("p", { class: "muted small", text: "No links yet." }));
  };
  const create = act("Create link", async () => {
    try {
      const l = await send(`/api/admin/nego/cycles/${c.id}/links`, "POST", { days: +days.value || undefined });
      const url = location.origin + (l.path || l.url);
      const field = el("input", { value: url, readonly: "", id: "link-url" });
      clear(out).append(el("div", { class: "preview ok" },
        el("p", { class: "small", text: "Copy this link now and send it to the principal by WhatsApp or email. You can also test what the principal sees using the Test Display button." }),
        field, el("div", { class: "actions" },
          act("Copy", async () => { field.select(); try { await navigator.clipboard.writeText(url); toast("Copied"); } catch (_) { document.execCommand("copy"); toast("Copied"); } }),
          el("a", { class: "act primary", href: url, target: "_blank", rel: "noopener", text: "Open & Test Display ↗" }),
          el("a", { class: "act secondary", href: `https://wa.me/?text=${encodeURIComponent(`Siloam Hospitals: silakan isi ${c.principal.name} di tautan berikut: ${url}`)}`, target: "_blank", rel: "noopener", text: "Share on WhatsApp" }))));
      draw((await api(`/api/admin/nego/cycles/${c.id}/links`)).links);
    } catch (e) { clear(out).append(el("p", { class: "error", text: e.message })); }
  });
  body.append(
    el("p", { class: "muted", text: "The principal opens this link to fill their current step online or in Excel. They only see their own items and their own columns: no PO volumes, no Siloam findings, no other principals." }),
    el("p", { class: "small", text: "Links are sent automatically when you move the negotiation to a principal step (if Power Automate is set up). To send by hand, create one here." }),
    el("form", { class: "progress-form", onsubmit: (e) => e.preventDefault() }, field("Valid for (days)", days), el("div", { class: "actions" }, create)),
    out, el("h3", { text: "Links" }), list);
  try { draw((await api(`/api/admin/nego/cycles/${c.id}/links`)).links); } catch (e) { list.append(el("p", { class: "error", text: e.message })); }
}

// ---------- brands: the same generic from different brands, side by side ----------
async function renderBrands(body) {
  clear(body);
  put(body, el("p", { class: "muted pad", text: "Comparing brands…" }));
  let d;
  try { d = await api(`/api/nego/cycles/${ng.cid}/groups`); } catch (e) { put(clear(body), el("p", { class: "error pad", text: e.message })); return; }
  const stage = { rfq: "RFQ", co: "Counter offer", fb1: "Feedback I", on: "Online nego", mou: "MOU", po: "Last PO" };
  const groupCard = (g) => card(`${g.label}`, `${g.members.length} items · cheapest ${unit(g.cheapest.price)}/pc · highest is ${pct(g.spread, 0)} above${g.potential_saving ? ` · ${money(g.potential_saving)} a year if the dearer ones matched the cheapest` : ""}`,
    el("div", { class: "table-wrap" }, el("table", {},
      el("thead", {}, el("tr", {}, ["Item", "Brand", "Price/pc incl. PPN", "Basis", "Above cheapest", "Annual pieces", "Extra cost a year"].map((h, i) => el("th", { class: i >= 2 && i !== 3 ? "num" : null, text: h })))),
      el("tbody", {}, g.members.map((m, i) => el("tr", { class: i === 0 ? "best" : null },
        el("td", {}, el("b", { text: m.item_name || "—" }), el("div", { class: "small muted", text: m.erp_code || "" })),
        el("td", { text: m.brand || "—" }),
        el("td", { class: "num", text: unit(m.price) }),
        el("td", { text: stage[m.stage] || "—" }),
        el("td", { class: "num", text: i === 0 ? "Cheapest" : `+${pct(m.gap, 1)}` }),
        el("td", { class: "num", text: m.annual_pcs ? num(m.annual_pcs) : "—" }),
        el("td", { class: "num", text: i === 0 || m.extra_cost == null ? "—" : money(m.extra_cost) })))))));
  const flagged = d.groups.filter((g) => g.flagged), rest = d.groups.filter((g) => !g.flagged);
  put(clear(body),
    card("Same medicine, different brands", `Items with the same generic name or group (set in Admin > Master data) compared per piece including PPN. A brand is flagged when it costs ${pct(d.spread_threshold, 0)} or more above the cheapest. Principals never see this.`,
      el("div", { class: "facts" }, [["Groups compared", d.groups.length], ["Flagged", d.flagged], ["Possible saving a year", d.potential_saving ? money(d.potential_saving) : "—"],
        ["Items in a group", d.grouped_items], ["Items with no generic name yet", d.ungrouped_items]].map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: typeof v === "number" ? num(v) : v })))),
      d.ungrouped_items ? el("p", { class: "small muted", text: isAdmin() ? "Label more items in Admin > Master data to compare more." : "An admin labels items in Master data." }) : null,
      canWork() && d.ungrouped_items ? act("Open master data", () => { location.hash = "#admin/master"; }, "secondary") : null),
    d.groups.length ? null : el("p", { class: "muted pad", text: "No group has two or more priced items yet." }),
    flagged.map(groupCard),
    rest.length ? el("details", {}, el("summary", { text: `${rest.length} groups within ${pct(d.spread_threshold, 0)}` }), rest.map(groupCard)) : null);
}

// ---------- negotiate: counter offer, benchmarks, online nego, package ----------
// Append children, skipping null/false (Element.append would print them as text).
const put = (n, ...kids) => { n.append(...kids.flat(2).filter((k) => k != null && k !== false)); return n; };

function benchRow(m, after) {
  return el("div", { class: `bm st-${m.status}` },
    el("div", {}, el("b", { text: `${m.source}: ${unit(m.price_pp)}/pc` }), el("span", { class: "small muted", text: ` · ${m.bm_name}${m.bm_unit ? ` (${m.bm_unit})` : ""}` }),
      el("div", { class: "small muted", text: `Match ${pct(m.confidence, 0)} by ${m.method === "ref" ? "catalogue no." : m.method === "remembered" ? "earlier confirmation" : "name"} · ${m.status === "confirmed" ? "confirmed" : m.status === "rejected" ? "rejected" : "to review"}` })),
    canWork() ? el("div", { class: "actions" },
      m.status !== "confirmed" ? act("Confirm", () => decideBench(m.id, "confirm", after), "ghost") : null,
      m.status !== "rejected" ? act("Not the same", () => decideBench(m.id, "reject", after), "ghost") : null) : null);
}

async function decideBench(mid, decision, after) {
  try { await send(`/api/admin/nego/cycles/${ng.cid}/benchmarks/${mid}`, "POST", { decision }); await refreshOverview(); after(); }
  catch (e) { toast(e.message); }
}

async function renderNegotiate(body) {
  clear(body);
  const c = ng.ov.cycle;
  const coBox = el("div"), onBox = el("div"), pkBox = el("div");
  put(body,
    card("4 · Counter offer", "Saran diskon counter offer per SKU, dihitung dari HNA RFQ sampai harga pembanding terbaik (Nett Price formularium, MOU, rata-rata PO, benchmark kompetitor). Centang SKU yang mau dipakai, lalu Apply.", coBox),
    card("6 · Online Nego", "Catat hasil meeting: peserta, kesepakatan per SKU, tindak lanjut, dan pilih skema final (Nett atau Disc). Disc sepakat per SKU diisi di tab Items atau di sini.", onBox),
    card("Submission package", "SKU yang ditandai lanjut ke BAK, lengkap (Qty, HNA, harga FIN). Excel Confirmation, draft BAK, lalu draft PKS. Dokumen yang dikirim principal hanya bisa dibuka admin.", pkBox));
  drawCO(coBox); drawOn(onBox, c); drawPackage(pkBox);
  if (canWork()) { const aiBox = el("div"); put(body, card("Draft CO, notulen, dan BAK", "Draft memakai data nego per SKU (disc dan harga FIN), bukan angka karangan. Periksa sebelum dikirim.", aiBox)); drawAssist(aiBox); }
}

async function drawAssist(box) {
  let st;
  try { st = await api("/api/admin/nego/assist"); } catch (e) { put(clear(box), el("p", { class: "error", text: e.message })); return; }
  const out = el("div", { "aria-live": "polite" });
  const go = (kind, label) => act(label, async () => {
    put(clear(out), el("p", { class: "muted small", text: "Drafting…" }));
    try {
      const r = await send(`/api/admin/nego/cycles/${ng.cid}/assist`, "POST", { kind });
      const ta = el("textarea", { rows: 14, class: "draft" });
      ta.value = r.text;
      put(clear(out), ta, el("div", { class: "actions" },
        act("Copy", async () => { ta.select(); try { await navigator.clipboard.writeText(ta.value); } catch (_) { document.execCommand("copy"); } toast("Copied"); }),
        el("span", { class: "small muted", text: `Disusun dari data nego: ${Object.entries(r.facts || {}).map(([k, v]) => `${k} ${v}`).join(", ")}. Periksa sebelum dikirim.` })));
    } catch (e) { put(clear(out), el("p", { class: "error", text: e.message })); }
  }, "secondary");
  void st;
  put(clear(box), el("div", { class: "actions" },
    go("co_cover", "Pesan counter offer (Bahasa)"), go("meeting_summary", "Notulen online nego (Bahasa)"), go("bak_points", "Poin BAK"), go("escalation_note", "Catatan eskalasi")), out);
}

// Counter-offer plan: a list per SKU (code, name, RFQ disc, suggested disc, comparison price and source, remarks).
// Siloam ticks the SKUs to use, then applies.
async function drawCO(box) {
  put(clear(box), el("p", { class: "muted small", text: "Working out suggestions…" }));
  let plan;
  try { plan = await api(`/api/nego/cycles/${ng.cid}/co-plan`); } catch (e) { put(clear(box), el("p", { class: "error", text: e.message })); return; }
  const k = ng.ov.kpis.stages;
  const rfq = k.find((s) => s.key === "rfq"), co = k.find((s) => s.key === "co");
  const facts = el("div", { class: "facts" },
    [["Items with an RFQ", num(rfq.items)], ["RFQ impact a year", rfq.impact == null ? "—" : `${rfq.impact > 0 ? "+" : ""}${money(rfq.impact)}`],
      ["Counter offers set", num(co.items)], ["Counter offer impact", co.impact == null ? "—" : `${co.impact > 0 ? "+" : ""}${money(co.impact)}`],
      ["Saran baru", num(plan.count)], ["Sama dengan RFQ (tanpa counter)", num(plan.same_as_rfq || 0)],
      ["Impact jika diterapkan vs RFQ", (plan.co_impact_vs_rfq ?? plan.co_impact) == null ? "—" : `${(plan.co_impact_vs_rfq ?? plan.co_impact) > 0 ? "+" : ""}${money(plan.co_impact_vs_rfq ?? plan.co_impact)}`]]
      .map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: v }))));
  const chosen = new Set(plan.suggestions.filter((s) => s.before == null).map((s) => s.item_id));
  const t = el("table", { class: "bm-table" });
  const all = el("input", { type: "checkbox", checked: "" });
  const applyBtn = act("Apply", async () => {
    if (!chosen.size) { toast("Centang SKU yang mau dipakai."); return; }
    try { const r = await send(`/api/admin/nego/cycles/${ng.cid}/co`, "POST", { item_ids: [...chosen], overwrite: true }); toast(`${r.applied} counter offers set`); await refreshOverview(); drawCO(box); } catch (e) { toast(e.message); }
  });
  const syncBtn = () => { applyBtn.textContent = `Apply ${chosen.size} SKU terpilih`; applyBtn.disabled = !chosen.size; };
  all.checked = chosen.size === plan.suggestions.length;
  all.addEventListener("change", () => { if (all.checked) plan.suggestions.forEach((s) => chosen.add(s.item_id)); else chosen.clear(); box.querySelectorAll("input.pick").forEach((i) => { i.checked = all.checked; }); syncBtn(); });
  if (plan.suggestions.length) {
    table(t, [
      { label: "", get: (s) => { const i = el("input", { type: "checkbox", class: "pick" }); i.checked = chosen.has(s.item_id); i.addEventListener("change", () => { if (i.checked) chosen.add(s.item_id); else chosen.delete(s.item_id); syncBtn(); }); return i; } },
      { label: "Kode", get: (s) => s.erp_code },
      { label: "Nama SKU", get: (s) => el("a", { href: "#", text: s.item_name, onclick: (e) => { e.preventDefault(); openItem(s.item_id); } }) },
      { label: "Disc RFQ", num: true, get: (s) => pct(s.rfq_disc, 2) },
      { label: "Disc saran", num: true, get: (s) => pct(s.co_disc, 2) },
      { label: "Harga pembanding", num: true, get: (s) => s.compare_price == null ? "—" : `${unit(s.compare_price)} (${s.compare_source})` },
      { label: "Harga CO", num: true, get: (s) => unit(s.co_unit_price) },
      { label: "Remarks", get: (s) => s.remarks ? el("span", { class: "warn-text small", text: s.remarks }) : (s.before != null ? el("span", { class: "small muted", text: `Sekarang ${pct(s.before, 2)}` }) : "") },
    ], plan.suggestions.slice(0, 300));
  }
  put(clear(box), facts,
    plan.suggestions.length ? el("div", {}, el("label", { class: "check" }, all, " Pilih semua"), el("div", { class: "table-wrap", style: "max-height:420px;overflow:auto" }, t)) : el("p", { class: "small muted", text: "Tidak ada saran baru: semua SKU dengan RFQ sudah sama dengan pembanding atau sudah punya counter offer." }),
    canWork() && plan.suggestions.length ? el("div", { class: "actions" }, applyBtn) : null);
  syncBtn();
}

function drawOn(box, c) {
  const esc = ng.ov.escalation;
  const m = c.meeting || {};
  const at = el("input", { type: "datetime-local", value: (m.at || c.meeting_at || "").slice(0, 16) || null });
  const sil = el("input", { value: m.attendees_siloam || "", placeholder: "Nama peserta Siloam" });
  const prin = el("input", { value: m.attendees_principal || "", placeholder: "Nama peserta principal" });
  const notes = el("textarea", { rows: 3, placeholder: "Catatan jalannya meeting" });
  notes.value = m.notes || c.meeting_notes || "";
  const follow = el("div");
  const rows = (m.followups || []).map((f) => ({ ...f }));
  const drawFollow = () => {
    clear(follow);
    rows.forEach((f, i) => {
      const tx = el("input", { value: f.text || "", placeholder: "Tindak lanjut" });
      const ow = el("input", { value: f.owner || "", placeholder: "PIC" });
      const du = el("input", { type: "date", value: f.due || null });
      tx.addEventListener("input", () => { f.text = tx.value; }); ow.addEventListener("input", () => { f.owner = ow.value; }); du.addEventListener("input", () => { f.due = du.value; });
      follow.append(el("div", { class: "row3" }, tx, ow, du, act("✕", () => { rows.splice(i, 1); drawFollow(); }, "ghost")));
    });
    follow.append(act("+ Tindak lanjut", () => { rows.push({ text: "", owner: "", due: "" }); drawFollow(); }, "ghost"));
  };
  drawFollow();
  const agreed = el("div", { class: "small", id: "on-agreed" });
  const escCls = esc.status === "needed" ? "need" : esc.status === "approved" || esc.status === "ok" ? "ok" : "";
  const escTitle = { none: "Belum ada harga sepakat", ok: "Dalam batas: tidak perlu eskalasi", needed: "Perlu eskalasi: persetujuan atasan sebelum Submission", approved: "Eskalasi disetujui" }[esc.status];
  const scheme = el("select", { disabled: canWork() ? null : "" }, [["", "Belum dipilih"], ["Nett", "Nett (template BAK Nett)"], ["Disc", "Disc (template BAK Disc)"]].map(([v, l]) => el("option", { value: v, text: l })));
  scheme.value = ng.ov.final_scheme || "";
  put(clear(box),
    el("div", { class: `esc ${escCls}` },
      el("b", { text: escTitle }),
      esc.items_agreed ? el("div", { class: "facts" },
        [["SKU sepakat", `${num(esc.items_agreed)} dari ${num(esc.active)} aktif`], ["Cost impact FIN", `${esc.impact > 0 ? "+" : ""}${money(esc.impact)}`], ["Dari nilai PO setahun (+PPN)", esc.impact_pct == null ? "—" : pct(esc.impact_pct, 2)], ["Batas", `${pct(esc.limit_pct, 1)} atau ${money(esc.limit_rp)}`]]
          .map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: v })))) : null,
      esc.status === "approved" ? el("p", { class: "small", text: `Disetujui ${esc.approval.by}: ${esc.approval.note}` }) : null,
      esc.status === "needed" && isAdmin() ? (() => { const n = el("input", { placeholder: "Catatan persetujuan atasan" }); return el("div", { class: "actions" }, n, act("Setujui eskalasi", async () => { try { await send(`/api/admin/nego/cycles/${ng.cid}/escalation`, "POST", { note: n.value }); toast("Eskalasi disetujui"); await refreshOverview(); drawOn(box, ng.ov.cycle); } catch (e) { toast(e.message); } })); })() : null),
    canWork() ? el("form", { class: "progress-form", onsubmit: async (e) => {
      e.preventDefault();
      try { await send(`/api/admin/nego/cycles/${ng.cid}`, "PATCH", { meeting_at: at.value, meeting: { attendees_siloam: sil.value, attendees_principal: prin.value, notes: notes.value, followups: rows } }); toast("Meeting saved"); await refreshOverview(); drawOn(box, ng.ov.cycle); } catch (ex) { toast(ex.message); }
    } },
      el("div", { class: "row3" }, field("Tanggal dan jam meeting", at), field("Peserta Siloam", sil), field("Peserta principal", prin)),
      field("Catatan meeting", notes),
      el("h4", { text: "Tindak lanjut dan tenggat" }), follow,
      el("div", { class: "actions" }, el("button", { type: "submit", class: "act secondary", text: "Save meeting" }),
        act("Isi disc sepakat dari FB1 / counter offer", async () => {
          try { const r = await send(`/api/admin/nego/cycles/${ng.cid}/on-fill`, "POST", { overwrite: false }); toast(`${r.filled} items filled`); await refreshOverview(); renderNegotiate($("ng-body")); } catch (e) { toast(e.message); }
        }, "ghost"))) : el("p", { class: "small", text: m.notes || "No meeting recorded yet." }),
    el("div", {}, el("h4", { text: "Skema final" }), el("p", { class: "small muted", text: "Isi disc untuk skema Nett dan skema Disc per SKU di detail item (kolom ON Disc skema). Pilih satu skema final: disc sepakat dan template BAK mengikuti." }),
      el("div", { class: "actions" }, scheme, canWork() ? act("Simpan skema", async () => { try { const r = await send(`/api/admin/nego/cycles/${ng.cid}`, "PATCH", { final_scheme: scheme.value || null }); toast(scheme.value ? `Skema ${scheme.value} dipilih` : "Skema dikosongkan"); await refreshOverview(); drawOn(box, ng.ov.cycle); void r; } catch (e) { toast(e.message); } }, "secondary") : null)),
    agreed);
  // per-SKU agreements, straight from the items
  (async () => {
    try {
      const d = await api(`/api/nego/cycles/${ng.cid}/items?filter=has_on&limit=60`);
      if (!d.items.length) { agreed.textContent = "Belum ada disc sepakat per SKU."; return; }
      const t = el("table");
      table(t, [{ label: "SKU", get: (i) => `${i.erp_code} ${i.item_name}` }, { label: "Disc FIN", num: true, get: (i) => pct(i.on_disc, 2) }, { label: "Harga FIN", num: true, get: (i) => unit(i.on_unit_price) }, { label: "Skema Nett", num: true, get: (i) => i.on_disc_nett == null ? "—" : pct(i.on_disc_nett, 2) }, { label: "Skema Disc", num: true, get: (i) => i.on_disc_disc == null ? "—" : pct(i.on_disc_disc, 2) }], d.items);
      clear(agreed).append(el("h4", { text: `Kesepakatan per SKU (${d.total})` }), el("div", { class: "table-wrap", style: "max-height:260px;overflow:auto" }, t));
    } catch (_) { /* shown empty */ }
  })();
}

async function drawPackage(box) {
  let chk;
  try { chk = await api(`/api/nego/cycles/${ng.cid}/package`); } catch (e) { put(clear(box), el("p", { class: "error", text: e.message })); return; }
  const dl = (href, label, kind = "") => el("a", { class: `act ${kind}`.trim(), href, text: label });
  put(clear(box), el("div", { class: `preview ${chk.ready ? "ok" : "bad"}` },
    el("h3", { text: chk.ready ? `Package siap: ${num(chk.agreed)} SKU lanjut ke BAK` : `Package belum siap (${num(chk.flagged)} SKU ditandai)` }),
    el("div", { class: "facts" }, [["SKU aktif", chk.active], ["Lanjut ke BAK", chk.flagged], ["Lengkap", chk.agreed], ["Tidak ikut BAK", chk.not_flagged], ["Discontinue", chk.discontinued]].map(([l, v]) => el("div", {}, el("span", { text: l }), el("b", { text: num(v) })))),
    chk.problems.length ? el("ul", { class: "errors" }, chk.problems.map((p) => el("li", { text: p }))) : null,
    chk.incomplete && chk.incomplete.length ? el("details", {}, el("summary", { text: `${chk.incomplete.length} SKU belum lengkap` }), el("ul", { class: "errors" }, chk.incomplete.slice(0, 50).map((x) => el("li", { text: `${x.erp_code} ${x.item_name}: ${x.missing.join(", ")}` })))) : null,
    chk.templates ? el("p", { class: "small" }, el("b", { text: `Template BAK: ${chk.templates.bak.label.replace("BAK template - ", "")}` }), ` · ${chk.templates.bak.using}. Binding ${chk.templates.binding}${chk.templates.from_final_scheme ? `, skema final ${chk.templates.scheme}` : ""} → template ${chk.templates.scheme}. PKS: ${chk.templates.pks.using}. `, isAdmin() ? el("a", { href: "#admin/documents", text: "Ganti template →" }) : null) : null,
    canWork() ? el("div", { class: "actions" },
      act("Tandai semua SKU dengan harga FIN", async () => { try { const r = await send(`/api/admin/nego/cycles/${ng.cid}/bak-flags`, "POST", { mode: "agreed", value: true }); toast(`${r.changed} SKU ditandai`); drawPackage(box); } catch (e) { toast(e.message); } }, "secondary"),
      act("Hapus semua tanda", async () => { try { const r = await send(`/api/admin/nego/cycles/${ng.cid}/bak-flags`, "POST", { value: false }); toast(`${r.changed} tanda dihapus`); drawPackage(box); } catch (e) { toast(e.message); } }, "ghost")) : null,
    canWork() && chk.ready ? el("div", { class: "actions" },
      dl(`/api/admin/nego/cycles/${ng.cid}/package.xlsx`, "1 · Excel Confirmation"),
      dl(`/api/admin/nego/cycles/${ng.cid}/bak.docx`, "2 · Draft BAK (Word)", "secondary"),
      dl(`/api/admin/nego/cycles/${ng.cid}/pks.docx`, "3 · Draft PKS (Word)", "secondary")) : null));
  if (!canWork()) return;
  let docs;
  try { docs = await api(`/api/admin/nego/cycles/${ng.cid}/documents`); } catch (_) { return; }
  put(box, el("h3", { text: "Dokumen dari principal" }),
    docs.documents.length ? el("ul", { class: "doclist" }, docs.documents.map((d) => el("li", {},
      isAdmin() ? el("a", { href: d.url, target: "_blank", rel: "noopener", text: `${d.label}: ${d.filename}` }) : el("span", { text: `${d.label}: ${d.filename}` }),
      el("span", { class: "small muted", text: ` · ${num(d.size / 1024, 0)} KB · ${new Date(d.uploaded_at).toLocaleString("en-GB")}` })))) :
      el("p", { class: "small muted", text: "Belum ada. Principal mengunggah di step Submission." }),
    !isAdmin() && docs.documents.length ? el("p", { class: "small muted", text: "Hanya admin yang bisa membuka dokumen." }) : null);
}

// ---------- analysis: Pre-Nego (results up to Feedback I), Post-Nego (FIN) ----------
const AN = { mode: "pre", step: "", q: "", offset: 0, limit: 100 };
async function renderAnalysis(body) {
  clear(body);
  const mode = el("select", {}, [["pre", "Pre-Nego (hasil sampai FB1)"], ["post", "Post-Nego (hasil FIN)"]].map(([v, l]) => el("option", { value: v, text: l })));
  mode.value = AN.mode;
  const step = el("select", {}, [["", "Step terakhir"], ["rfq", "RFQ"], ["co", "Counter offer"], ["fb1", "Feedback I"], ["on", "FIN (Online Nego)"]].map(([v, l]) => el("option", { value: v, text: l })));
  step.value = AN.step;
  const q = el("input", { type: "search", placeholder: "Cari ERP, nama, brand…", value: AN.q || null });
  const sum = el("div"), table0 = el("div", { class: "table-wrap", style: "max-height:560px;overflow:auto" }), pager = el("div", { class: "pager" });
  const qs = () => `mode=${AN.mode}${AN.step ? `&step=${AN.step}` : ""}`;
  body.append(el("div", { class: "filters board-filters" },
    el("label", {}, "Tampilan", mode), el("label", {}, "Step", step), el("label", {}, "Cari", q),
    el("a", { class: "act secondary", href: `/api/nego/cycles/${ng.cid}/analysis.xlsx?${qs()}`, id: "an-xls", text: "Export Excel ⤓" })), sum, table0, pager);
  const reload = () => { AN.mode = mode.value; AN.step = step.value; AN.q = q.value; AN.offset = 0; $("an-xls").href = `/api/nego/cycles/${ng.cid}/analysis.xlsx?${qs()}`; load(); };
  mode.addEventListener("change", () => { if (mode.value === "post" && !step.value) step.value = ""; reload(); });
  step.addEventListener("change", reload);
  let timer; q.addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(reload, 300); });
  async function load() {
    clear(table0).append(el("p", { class: "muted pad", text: "Loading…" }));
    let d;
    try { d = await api(`/api/nego/cycles/${ng.cid}/analysis?${qs()}&q=${encodeURIComponent(AN.q)}&offset=${AN.offset}&limit=${AN.limit}`); } catch (e) { clear(table0).append(el("p", { class: "error pad", text: e.message })); return; }
    const s = d.summary, y = s.po_year;
    const sg = (v) => (v == null ? "—" : `${v > 0 ? "+" : ""}${money(v)}`);
    const tilesRow = el("div", { class: "kpis5" },
      tile(`Nilai PO (+PPN)`, money(s.po_values[y[2]]), `${y[0]}: ${money(s.po_values[y[0]])} · ${y[1]}: ${money(s.po_values[y[1]])} · ${y[2]} setahun`),
      tile("Growth", s.growth_26_vs_25 == null ? "—" : pct(s.growth_26_vs_25, 1), `${y[2]} vs ${y[1]} · vs ${y[0]}: ${s.growth_26_vs_24 == null ? "—" : pct(s.growth_26_vs_24, 1)}`),
      tile("Cost impact", sg(s.cost_impact), s.cost_impact_pct == null ? "—" : `${pct(s.cost_impact_pct, 2)} dari nilai PO setahun`),
      tile("Margin impact vs PO", sg(s.margin_impact_po), s.margin_impact_po_pct == null ? "—" : pct(s.margin_impact_po_pct, 2)),
      tile("Margin impact vs Formularium", sg(s.margin_impact_form), s.margin_impact_form_pct == null ? "—" : pct(s.margin_impact_form_pct, 2)));
    const b = s.bucket;
    clear(sum).append(el("p", { class: "small muted", text: `${num(s.sku)} SKU dengan harga di step ini · Bucket: Consump ${num(b.Consump)}, Charge ${num(b.Charge)}, tanpa bucket ${num(b.other)}` }), tilesRow);
    const groups = [];
    d.columns.forEach((c) => { const last = groups[groups.length - 1]; if (last && last.g === c.group) last.n++; else groups.push({ g: c.group, n: 1 }); });
    const fmt = (c, v) => (v == null || v === "" ? "" : c.kind === "pct" ? pct(v, 2) : c.kind === "price" ? num(v, 2) : c.kind === "money" ? num(v, 0) : c.kind === "num" ? num(v, v % 1 ? 2 : 0) : String(v));
    const t = el("table", { class: "items" },
      el("thead", {}, el("tr", { class: "sec" }, groups.map((x) => el("th", { colspan: x.n, text: x.g }))), el("tr", {}, d.columns.map((c) => el("th", { class: c.kind !== "text" ? "num" : "", text: c.label })))),
      el("tbody", {}, d.rows.map((r) => el("tr", { class: "clickable", onclick: () => openItem(r.id ?? null) }, d.columns.map((c) => el("td", { class: c.kind !== "text" ? "num" : "", text: fmt(c, r[c.key]) }))))));
    clear(table0).append(d.rows.length ? t : el("p", { class: "muted pad", text: "Belum ada data." }));
    const p0 = d.offset + 1, p1 = Math.min(d.offset + d.rows.length, d.total);
    clear(pager).append(el("span", { class: "small muted", text: `${d.total ? `${num(p0)}–${num(p1)}` : 0} dari ${num(d.total)} SKU · ${d.columns.length} kolom` }),
      el("button", { type: "button", class: "act ghost", text: "← Previous", disabled: d.offset ? null : "", onclick: () => { AN.offset = Math.max(0, d.offset - AN.limit); load(); } }),
      el("button", { type: "button", class: "act ghost", text: "Next →", disabled: p1 < d.total ? null : "", onclick: () => { AN.offset = d.offset + AN.limit; load(); } }));
  }
  load();
}
