"use strict";
// Principal portal: one step at a time, one SKU at a time if they like, with checks as they
// type. Every value goes through the server (portal.py), which decides what this principal
// may see and change; this page only makes entry easy. Data goes into the DOM through
// textContent, never innerHTML.

const P = { me: null, lang: "id", filter: "todo", q: "", items: [], total: 0, counts: {}, view: "list", one: 0, loading: false };
try { P.lang = localStorage.getItem("nego-p-lang") || "id"; } catch (_) { /* private mode */ }
const NARROW = window.innerWidth < 700;
try { P.view = localStorage.getItem("nego-p-view") || (NARROW ? "one" : "table"); } catch (_) { P.view = NARROW ? "one" : "table"; }
if (!["table", "list", "one"].includes(P.view) || (NARROW && P.view === "table")) P.view = NARROW ? "one" : "table";

const T = {
  id: {
    portal: "Portal Principal", loading: "Memuat…",
    invalid_t: "Tautan tidak berlaku", invalid_p: "Tautan ini sudah kedaluwarsa atau dicabut. Minta tautan baru ke tim pengadaan.",
    step_of: (n) => `Langkah ${n} dari 6`, your_task: "Tugas Anda",
    steps: { rfq: "Isi data item & harga penawaran (RFQ)", feedback1: "Tanggapi counter offer", submission: "Kirim dokumen perusahaan" },
    last_step: "Langkah terakhir", doc_required: "Wajib", doc_optional: "Opsional", doc_upload: "Unggah", doc_uploaded: "Terunggah", doc_delete: "Hapus",
    doc_note: "Dokumen hanya bisa dibuka admin Siloam. Hanya PDF, JPG, atau PNG, maksimal 10 MB per file.",
    howto: {
      rfq: ["Untuk setiap SKU: pilih status (Aktif atau Discontinue), isi brand dan nomor katalog (REF), atau zat aktif untuk obat.", "Isi HNA per kemasan (sebelum PPN), isi per kemasan, dan diskon sebagai angka 0 sampai 100 (contoh: 15). Tanda % muncul otomatis.", "Jika ada peringatan, periksa kembali isian Anda. Jika harga naik, mohon isi alasannya. Semua tersimpan otomatis."],
      feedback1: ["Tim pengadaan mengirim counter offer berupa diskon per SKU.", "Klik Terima, atau isi diskon Anda sendiri dan alasannya.", "Klik Kirim jika sudah selesai."],
      submission: ["Harga sudah disepakati. Langkah terakhir: unggah dokumen perusahaan.", "Dokumen bertanda Wajib harus diunggah sebelum mengirim.", "Klik Kirim jika sudah lengkap."],
    },
    help: "Bantuan", help_t: "Butuh bantuan?", help_p: "Hubungi tim pengadaan. Kami akan membantu Anda mengisi.", help_wa: "Chat WhatsApp", help_mail: "Kirim email", help_none: "Balas pesan atau email kami, tim kami akan membantu.",
    excel_or: "Lebih suka Excel? Unduh file, isi, lalu unggah kembali.", view_table: "Tabel", view_list: "Kartu", view_one: "Satu per satu",
    h_item: "Barang", h_status: "Status", h_brand: "Brand", h_ref: "REF", h_remarks: "Keterangan", h_hna: "HNA / kemasan", h_qty: "Isi (pcs)", h_disc: "Diskon % (0–100)", h_unit: "Harga/pcs", h_rfq: "Harga Anda", h_co: "Counter Offer", h_own: "Diskon Anda",
    tip_paste: "Tips: salin sel dari Excel lalu tempel di kolom pertama yang mau diisi.",
    contract: "Kontrak", due: "Batas waktu", progress: (d, a) => `${d} dari ${a} SKU selesai`,
    download: "Unduh Excel", upload: "Unggah Excel", send: "Kirim Tanggapan",
    f: { todo: "Belum diisi", check: "Perlu dicek", up: "Harga naik", all: "Semua", done: "Selesai" },
    search: "Cari nama, kode, brand, REF",
    fill_fb1: "Terima semua counter offer", next_todo: "Ke SKU berikutnya yang belum diisi",
    fill_confirm_fb1: () => "Terima counter offer untuk semua SKU yang belum ditanggapi?",
    filled: (n) => `${n} SKU terisi`, none: "Tidak ada SKU di sini.", more: "Tampilkan lebih banyak", prev: "← Sebelumnya", next: "Berikutnya →",
    state: { missing: "Belum diisi", check: "Perlu dicek", done: "Selesai" },
    status: "Status", active: "Aktif", disc_btn: "Discontinue",
    qty: (pack) => `Isi per ${pack} (pcs)`, hna: (pack) => `HNA per ${pack}, sebelum PPN`, disc: "Diskon",
    per_pcs: "Harga per pcs termasuk PPN",
    brand: "Brand / pabrikan", ref: "Nomor katalog (REF)", ai: "Zat aktif (Active Ingredient)", remarks: "Keterangan (opsional)",
    rfq_price: "Harga RFQ Anda", co: "Counter Offer", accept: "Terima counter offer", own: "Atau isi diskon Anda",
    reason: "Alasan", reason_pick: "Pilih alasan…", reason_detail: "Keterangan singkat (opsional)", yes: "Ya, sudah benar", confirmed: "Sudah dikonfirmasi",
    saved: "Tersimpan ✓", saving: "Menyimpan…", unsaved: "Belum tersimpan. Coba lagi", retry: "Coba lagi",
    waiting_t: "Menunggu Verifikasi", waiting_p: "Tidak ada yang perlu diisi saat ini. Tim pengadaan akan menghubungi Anda untuk langkah berikutnya.",
    sent_t: "Terima kasih, sudah terkirim", sent_p: (at) => `Dikirim pada ${at}. Tim pengadaan sedang meninjau. Anda akan dihubungi untuk langkah berikutnya.`,
    sum_t: "Periksa sebelum mengirim", sum_items: "SKU", sum_done: "Selesai", sum_missing: "Belum diisi", sum_check: "Perlu dicek", sum_disc: "Discontinue",
    sum_left: "Perlu dilengkapi dulu", sum_fix: "Perbaiki", sum_up: "Harga naik (dengan alasan)", sum_ok: "Semua SKU sudah lengkap.",
    sum_agree: "Saya sudah memeriksa dan data ini benar.", sum_send: "Kirim sekarang", cancel: "Batal", close: "Tutup",
    up_t: "Unggah Excel", up_p: "Gunakan file Excel yang diunduh dari halaman ini. Kolom yang bisa diisi berwarna oranye. Baris yang salah tidak disimpan dan ditampilkan di bawah.",
    up_check: "Periksa file", up_res: (a, r) => `${a} baris siap disimpan, ${r} baris perlu diperbaiki.`, up_save: (n) => `Simpan ${n} baris`,
    up_done: (n) => `${n} baris tersimpan`, row: "Baris", not_found: (n) => `${n} baris tidak dikenali (kode ERP tidak ada di daftar).`,
    sent_ok: "Terkirim. Terima kasih!", err: "Terjadi kesalahan",
    locked_t: "Sudah terkirim dan dikunci", locked_p: (at) => `Isian ini dikirim${at ? ` pada ${at}` : ""} dan sekarang dikunci supaya angkanya tidak berubah saat counter offer disiapkan. Jika ada yang perlu diubah, tulis pengajuan koreksi pada SKU yang bersangkutan.`,
    not_open: "Langkah ini belum dibuka. Tim pengadaan akan menghubungi Anda.",
    wait: { review_rfq: "Siloam sedang meninjau RFQ Anda", counter_offer: "Menunggu counter offer dari Siloam", online_nego: "Menunggu jadwal online nego dengan Siloam", closed: "Selesai. Terima kasih", prepare: "Siloam sedang menyiapkan negosiasi" },
    overdue: (n) => `Terlambat ${n} hari`, returned: "Siloam mengembalikan isian Anda", flagged: "Siloam meminta Anda memeriksa SKU ini",
    comment_btn: "Ajukan koreksi", comment_ph: "Tulis koreksi yang Anda minta (mis. HNA salah ketik)", comment_send: "Kirim ke Siloam", comment_sent: "Pengajuan terkirim ke Siloam", comment_open: "Menunggu Siloam", comment_done: "Sudah ditangani",
    dc_t: "Konfirmasi SKU Discontinue", dc_p: (n) => `${n} SKU dinyatakan Discontinue. Surat DC diurus bersama Siloam; Anda bisa melampirkan surat DC atau dokumen lain (tidak wajib).`, dc_confirm: "Saya konfirmasi SKU di atas Discontinue.", dc_attach: "Lampirkan dokumen", dc_attached: (n) => `${n} dokumen terlampir`,
    up_changes: "Isian yang akan berubah", up_overwrite: (n) => `${n} isian di portal akan ditimpa oleh file ini.`, up_overwrite_ok: "Saya setuju menimpa isian portal dengan isi file ini.", up_old: "Sekarang", up_new: "File",
    disc_bad: "Isi diskon 0 sampai 100",
  },
  en: {
    portal: "Principal Portal", loading: "Loading…",
    invalid_t: "Link not valid", invalid_p: "This link has expired or was revoked. Ask the procurement team for a new one.",
    step_of: (n) => `Step ${n} of 6`, your_task: "Your task",
    steps: { rfq: "Item details & price quote (RFQ)", feedback1: "Respond to counter offer", submission: "Send company documents" },
    last_step: "Last step", doc_required: "Required", doc_optional: "Optional", doc_upload: "Upload", doc_uploaded: "Uploaded", doc_delete: "Delete",
    doc_note: "Only Siloam admins can open the documents. PDF, JPG or PNG only, up to 10 MB each.",
    howto: {
      rfq: ["For every SKU: choose the status (Active or Discontinue), then fill in the brand and catalogue number (REF).", "Enter the HNA per pack (before PPN), the pieces per pack and the discount as a number from 0 to 100 (e.g. 15). The % sign appears by itself.", "If a warning shows, please re-check your entry. If the price goes up, please give the reason. Everything saves automatically."],
      feedback1: ["Procurement sent a counter-offer discount for each SKU.", "Click Accept, or enter your own discount and the reason.", "Click Send when you're done."],
      submission: ["Prices are agreed. Last step: upload your company documents.", "Documents marked Required must be uploaded before sending.", "Click Send when everything is there."],
    },
    help: "Help", help_t: "Need help?", help_p: "Contact the procurement team. We'll help you fill it in.", help_wa: "Chat on WhatsApp", help_mail: "Send an email", help_none: "Reply to our message or email and our team will help.",
    excel_or: "Prefer Excel? Download the file, fill it in, then upload it again.", view_table: "Table", view_list: "Cards", view_one: "One by one",
    h_item: "Item", h_status: "Status", h_brand: "Brand", h_ref: "REF", h_remarks: "Remarks", h_hna: "HNA / pack", h_qty: "Pieces", h_disc: "Discount % (0–100)", h_unit: "Price/pc", h_rfq: "Your price", h_co: "Counter Offer", h_own: "Your discount",
    tip_paste: "Tip: copy cells from Excel and paste them into the first column you want to fill.",
    contract: "Contract", due: "Due", progress: (d, a) => `${d} of ${a} SKUs done`,
    download: "Download Excel", upload: "Upload Excel", send: "Send Response",
    f: { todo: "To fill", check: "To check", up: "Price up", all: "All", done: "Done" },
    search: "Search name, code, brand, REF",
    fill_fb1: "Accept every counter offer", next_todo: "Next SKU to fill",
    fill_confirm_fb1: () => "Accept counter offer for every SKU you haven't answered?",
    filled: (n) => `${n} SKUs filled`, none: "No SKUs here.", more: "Show more", prev: "← Previous", next: "Next →",
    state: { missing: "To fill", check: "To check", done: "Done" },
    status: "Status", active: "Active", disc_btn: "Discontinue",
    qty: (pack) => `Pieces per ${pack}`, hna: (pack) => `HNA per ${pack}, before PPN`, disc: "Discount",
    per_pcs: "Price per piece incl. PPN",
    brand: "Brand / manufacturer", ref: "Catalogue no. (REF)", ai: "Active ingredient", remarks: "Remarks (optional)",
    rfq_price: "Your RFQ price", co: "Counter Offer", accept: "Accept counter offer", own: "Or enter your discount",
    reason: "Reason", reason_pick: "Choose a reason…", reason_detail: "Short detail (optional)", yes: "Yes, that's right", confirmed: "Confirmed",
    saved: "Saved ✓", saving: "Saving…", unsaved: "Not saved. Try again", retry: "Try again",
    waiting_t: "Waiting for Verification", waiting_p: "There's nothing to fill right now. The procurement team will contact you for the next step.",
    sent_t: "Thank you, it's sent", sent_p: (at) => `Sent on ${at}. The procurement team is reviewing it and will contact you for the next step.`,
    sum_t: "Check before sending", sum_items: "SKUs", sum_done: "Done", sum_missing: "To fill", sum_check: "To check", sum_disc: "Discontinue",
    sum_left: "Complete these first", sum_fix: "Fix", sum_up: "Price increases (with reasons)", sum_ok: "Every SKU is complete.",
    sum_agree: "I've checked this and it's correct.", sum_send: "Send now", cancel: "Cancel", close: "Close",
    up_t: "Upload Excel", up_p: "Use the Excel file downloaded from this page. The cells you can fill are orange. Rows with mistakes aren't saved and are listed below.",
    up_check: "Check file", up_res: (a, r) => `${a} rows ready to save, ${r} rows need fixing.`, up_save: (n) => `Save ${n} rows`,
    up_done: (n) => `${n} rows saved`, row: "Row", not_found: (n) => `${n} rows not recognised (ERP code not in the list).`,
    sent_ok: "Sent. Thank you!", err: "Something went wrong",
    locked_t: "Sent and locked", locked_p: (at) => `This form was sent${at ? ` on ${at}` : ""} and is now locked so the numbers stay put while the counter offer is prepared. If something must change, leave a correction request on that SKU.`,
    not_open: "This step is not open yet. The procurement team will contact you.",
    wait: { review_rfq: "Siloam is reviewing your RFQ", counter_offer: "Waiting for Siloam's counter offer", online_nego: "Waiting for the online nego with Siloam", closed: "Done. Thank you", prepare: "Siloam is preparing the negotiation" },
    overdue: (n) => `${n} days late`, returned: "Siloam sent your form back", flagged: "Siloam asks you to check this SKU",
    comment_btn: "Request a correction", comment_ph: "Describe the correction you need (e.g. HNA typo)", comment_send: "Send to Siloam", comment_sent: "Request sent to Siloam", comment_open: "Waiting for Siloam", comment_done: "Handled",
    dc_t: "Confirm Discontinue SKUs", dc_p: (n) => `${n} SKUs are marked Discontinue. The DC letter is arranged with Siloam; you may attach it or another document (optional).`, dc_confirm: "I confirm the SKUs above are Discontinue.", dc_attach: "Attach a document", dc_attached: (n) => `${n} document(s) attached`,
    up_changes: "Entries that will change", up_overwrite: (n) => `${n} entries already in the portal will be overwritten by this file.`, up_overwrite_ok: "I agree to overwrite the portal entries with this file.", up_old: "Now", up_new: "File",
    disc_bad: "Enter a discount from 0 to 100",
  },
};
const t = (k, ...a) => { const v = k.split(".").reduce((o, p) => (o == null ? o : o[p]), T[P.lang]); return typeof v === "function" ? v(...a) : v ?? k; };
const msg = (o) => (o ? o[P.lang] || o.id : "");

// ---------- helpers ----------
const $ = (id) => document.getElementById(id);
function el(tag, attrs = {}, ...kids) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v == null || v === false) continue;
    if (k === "class") n.className = v;
    else if (k === "text") n.textContent = v;
    else if (k.startsWith("on")) n.addEventListener(k.slice(2), v);
    else n.setAttribute(k, v === true ? "" : v);
  }
  for (const c of kids.flat(3)) if (c != null && c !== false) n.append(c instanceof Node ? c : document.createTextNode(String(c)));
  return n;
}
const clear = (n) => { while (n.firstChild) n.removeChild(n.firstChild); return n; };
// Append children, skipping null/false (Element.append would print them as text).
const put = (n, ...kids) => { n.append(...kids.flat(3).filter((k) => k != null && k !== false)); return n; };
const nf = (d) => new Intl.NumberFormat("id-ID", { minimumFractionDigits: d, maximumFractionDigits: d });
const num = (v, d = 0) => (v == null ? "—" : nf(d).format(v));
const rp = (v) => (v == null ? "—" : `Rp ${num(v, v < 100 && v % 1 ? 2 : 0)}`);
const pctTxt = (f) => (f == null ? "" : `${num(f * 100, Math.abs(f * 100 - Math.round(f * 100)) > 1e-6 ? 2 : 0)}%`);
const fmtDate = (iso) => new Intl.DateTimeFormat(P.lang === "id" ? "id-ID" : "en-GB", { day: "numeric", month: "long", year: "numeric" }).format(new Date(iso.length <= 10 ? iso + "T00:00:00" : iso));
const unitPrice = (hna, qty, disc) => (hna == null || !qty || Number.isNaN(disc) ? null : (hna / qty) * (1 - (disc || 0)) * (1 + (P.me?.ppn ?? 0.11)));
function toast(text) {
  const n = $("ptoast"); n.textContent = text; n.hidden = false;
  clearTimeout(toast.t); toast.t = setTimeout(() => { n.hidden = true; }, 3000);
}
async function call(path, { method = "GET", body, form } = {}) {
  const sp = new URLSearchParams(location.search);
  const token = sp.get("token") || sp.get("t");
  const cycleId = sp.get("cycle_id") || sp.get("cid");
  if (cycleId && !path.includes("cycle_id=") && !path.includes("cid=")) {
    path += (path.includes("?") ? "&" : "?") + "cycle_id=" + encodeURIComponent(cycleId);
  }
  if (token && !path.includes("token=") && !path.includes("t=")) {
    path += (path.includes("?") ? "&" : "?") + "token=" + encodeURIComponent(token);
  }
  const opts = { method, headers: { "X-Requested-With": "nego" }, credentials: "same-origin" };
  if (token) opts.headers["X-Principal-Token"] = token;
  const staffToken = localStorage.getItem("nego_token") || sessionStorage.getItem("nego_token");
  if (staffToken) {
    opts.headers["Authorization"] = `Bearer ${staffToken}`;
    opts.headers["X-Nego-Token"] = staffToken;
  }
  if (form) opts.body = form;
  else if (body !== undefined) { opts.body = JSON.stringify(body); opts.headers["Content-Type"] = "application/json"; }
  const r = await fetch(path, opts);
  let j = {};
  try { j = await r.json(); } catch (_) { /* not json */ }
  if (r.status === 401) { showInvalid(); const e401 = new Error(j.error || "401"); e401.status = 401; throw e401; }
  if (!r.ok) {
    const e = String(j.error || r.statusText);
    const err = new Error(e.includes(" / ") ? e.split(" / ")[P.lang === "id" ? 0 : 1] : e);
    err.status = r.status;
    throw err;
  }
  return j;
}
// Typed numbers: HNA and Qty keep digits (and a decimal comma for HNA); discounts are percent.
const digits = (s) => String(s || "").replace(/[^\d]/g, "");
// Values go to the server with a decimal comma and no thousands dots, so "1,25" can never be read as 1.250.
function hnaValue(s) { const [a, b] = String(s || "").replace(/\./g, "").split(","); const d = digits(a); return d ? (b !== undefined ? `${d},${digits(b)}` : d) : ""; }
function hnaShow(v) {
  if (v == null || v === "") return "";
  const [a, b] = typeof v === "number" ? String(v).split(".") : String(v).split(",");
  return nf(0).format(+a) + (b !== undefined ? "," + b : "");
}
// Discount: the principal types a number from 0 to 100 (15 or 27,5); the % sign is added by the page.
// A value outside 0-100 gives NaN: no price is shown and no reason is asked for.
function discValue(s) { return String(s || "").trim().replace(/%/g, ""); }
const discShow = (f) => (f == null ? "" : `${String(+(f * 100).toFixed(2)).replace(".", ",")}%`);
const discFraction = (s) => {
  const v = String(s || "").trim().replace(/%/g, "").replace(",", ".");
  if (!v) return 0;
  const n = +v;
  return Number.isNaN(n) || n < 0 || n >= 100 ? NaN : n / 100;
};

// ---------- boot ----------
function showInvalid() {
  put(clear($("app")), el("section", { class: "pcard center" }, el("h1", { text: t("invalid_t") }), el("p", { text: t("invalid_p") })));
}

// ---------- zoom controller ----------
function initZoomController() {
  const ZOOM_LEVELS = [0.75, 0.8, 0.85, 0.9, 0.95, 1.0, 1.05, 1.1, 1.15, 1.25, 1.35];
  const STORAGE_KEY = "nego_web_zoom_level";
  const dec = $("zoom-dec");
  const val = $("zoom-val");
  const inc = $("zoom-inc");
  if (!dec || !inc) return;

  let current = 1.0;
  try {
    const saved = parseFloat(localStorage.getItem(STORAGE_KEY));
    if (!isNaN(saved) && saved >= 0.7 && saved <= 1.4) current = saved;
  } catch (_) {}

  const apply = (scale, save = true) => {
    scale = Math.min(1.4, Math.max(0.7, Math.round(scale * 100) / 100));
    const pct = Math.round(scale * 100);
    document.documentElement.style.zoom = String(scale);
    document.documentElement.style.setProperty("--app-zoom", String(scale));
    document.documentElement.style.setProperty("--base-font-size", `${14 * scale}px`);
    if (val) {
      val.textContent = `${pct}%`;
      val.title = scale === 1.0 
        ? (P.lang === "id" ? "Ukuran standar (100%). Klik untuk reset" : "Default size (100%). Click to reset")
        : (P.lang === "id" ? `Ukuran: ${pct}%. Klik untuk reset 100%` : `Size: ${pct}%. Click to reset 100%`);
    }
    if (save) {
      try { localStorage.setItem(STORAGE_KEY, String(scale)); } catch (_) {}
    }
  };

  apply(current, false);

  dec.addEventListener("click", () => {
    const next = [...ZOOM_LEVELS].reverse().find((z) => z < current - 0.01) || ZOOM_LEVELS[0];
    current = next;
    apply(current);
  });

  inc.addEventListener("click", () => {
    const next = ZOOM_LEVELS.find((z) => z > current + 0.01) || ZOOM_LEVELS[ZOOM_LEVELS.length - 1];
    current = next;
    apply(current);
  });

  if (val) {
    val.addEventListener("click", () => {
      current = 1.0;
      apply(current);
    });
  }

  window.addEventListener("storage", (e) => {
    if (e.key === STORAGE_KEY && e.newValue) {
      const v = parseFloat(e.newValue);
      if (!isNaN(v) && v >= 0.7 && v <= 1.4) {
        current = v;
        apply(current, false);
      }
    }
  });
}

async function boot() {
  document.querySelectorAll(".lang button").forEach((b) => b.addEventListener("click", () => setLang(b.dataset.lang)));
  $("helpbtn").addEventListener("click", openHelp);
  initZoomController();
  setLang(P.lang, false);
  aiInit();
  if (new URLSearchParams(location.search).get("invalid")) { showInvalid(); return; }
  try {
    P.me = await call("/api/p/me");
  } catch (err) {
    if (err.status !== 401 && !/401/.test(err.message)) {
      console.error("Principal portal me error:", err);
      put(clear($("app")), el("section", { class: "pcard center" },
        el("h2", { text: P.lang === "id" ? "Gagal memuat portal principal" : "Failed to load principal portal" }),
        el("p", { class: "muted", text: err.message || (P.lang === "id" ? "Tautan atau sesi tidak valid" : "Invalid link or session") }),
        el("button", { type: "button", class: "pbtn primary", onclick: () => location.reload(), text: t("retry") })
      ));
    }
    return;
  }
  $("who").textContent = P.me.principal.name;
  if (P.me.locked) P.filter = "all";
  render();
}

// ---------- AI helper ----------
const AI = { history: [], busy: false };
const AIT = {
  id: { title: "Asisten pengisian RFQ", ph: "Tanya cara mengisi…", send: "Kirim", hi: "Halo! Saya bisa membantu mengisi RFQ: format diskon, HNA, isi per kemasan, status item. Anda tetap menekan Kirim sendiri.", err: "Maaf, belum bisa menjawab saat ini." },
  en: { title: "RFQ filling assistant", ph: "Ask how to fill in…", send: "Send", hi: "Hi! I can help you fill in the RFQ: discount format, HNA, pieces per pack, item status. You always press Send yourself.", err: "Sorry, I can't answer right now." },
};
function aiAdd(who, text) {
  const log = $("ai-log");
  log.append(el("div", { class: `aimsg ${who}`, text }));
  log.scrollTop = log.scrollHeight;
}
function aiLang() {
  const x = AIT[P.lang];
  $("ai-title").textContent = x.title; $("ai-input").placeholder = x.ph; $("ai-send").textContent = x.send;
  if (!$("ai-log").children.length) aiAdd("bot", x.hi);
}
function aiInit() {
  // Staff previewing the portal already have the app-wide AI button.
  if (window.self !== window.top) { $("aibtn").hidden = true; return; }
  $("aibtn").addEventListener("click", () => {
    const open = $("aipanel").hidden;
    $("aipanel").hidden = !open; $("aibtn").setAttribute("aria-expanded", String(open));
    if (open) { aiLang(); $("ai-input").focus(); }
  });
  $("ai-close").addEventListener("click", () => { $("aipanel").hidden = true; $("aibtn").setAttribute("aria-expanded", "false"); });
  $("ai-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const text = $("ai-input").value.trim();
    if (!text || AI.busy) return;
    $("ai-input").value = ""; aiAdd("me", text); AI.busy = true;
    try {
      const r = await call("/api/p/chat", { method: "POST", body: { message: text, lang: P.lang, history: AI.history } });
      aiAdd("bot", r.reply);
      AI.history.push({ role: "user", text }, { role: "model", text: r.reply });
    } catch (err) { aiAdd("bot", err?.message || AIT[P.lang].err); }
    AI.busy = false;
  });
}

function setLang(l, redraw = true) {
  P.lang = l === "en" ? "en" : "id";
  try { localStorage.setItem("nego-p-lang", P.lang); } catch (_) { /* ignore */ }
  document.documentElement.lang = P.lang;
  document.querySelectorAll(".lang button").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.lang === P.lang)));
  $("t-portal").textContent = t("portal");
  $("helpbtn").textContent = t("help");
  document.title = t("portal");
  if ($("ai-title")) aiLang();
  if (redraw && P.me) render();
}

// ---------- page ----------
function render() {
  const app = clear($("app"));
  const me = P.me;
  if (!me.step) {
    const sent = me.submitted;
    put(app, el("section", { class: "pcard center" },
      el("h1", { text: sent ? t("sent_t") : t("waiting_t") }),
      el("p", { text: sent ? t("sent_p", fmtDate(sent.at)) : t("waiting_p") })));
    return;
  }
  const howto = el("ol", { class: "howto" }, t(`howto.${me.step}`).map((s) => el("li", { text: s })));
  const prog = el("div", { class: "prog" }, el("div", { class: "bar" }, el("i", { id: "prog-fill" })), el("span", { id: "prog-text" }));
  const sendbar = me.locked ? null : el("div", { class: "sendbar" }, prog,
    el("button", { type: "button", class: "pbtn primary big", text: t("send"), onclick: openSummary }));
  const lockBanner = me.locked ? el("section", { class: "pcard lockcard", role: "status" },
    el("h2", { text: `🔒 ${t(me.lock_reason === "sent" ? "locked_t" : "not_open")}` }),
    me.lock_reason === "sent" ? el("p", { text: t("locked_p", me.submitted ? fmtDate(me.submitted.at) : "") }) : null) : null;
  const waiting = me.turn === "siloam";
  const returned = me.return_note ? el("p", { class: "warnline", role: "status" }, `${t("returned")}${me.return_note.message ? `: ${me.return_note.message}` : ""}`) : null;
  put(app,
    el("section", { class: "pcard stepcard" },
      el("p", { class: "eyebrow", text: waiting ? "" : t("your_task") }),
      el("h1", { text: waiting ? (t(`wait.${me.current_step}`) !== `wait.${me.current_step}` ? t(`wait.${me.current_step}`) : t(`steps.${me.step}`)) : t(`steps.${me.step}`) }),
      el("p", { class: "muted", text: `${t("contract")} ${fmtDate(me.contract.start)} – ${fmtDate(me.contract.end)}${me.due && !waiting ? ` · ${t("due")} ${fmtDate(me.due)}` : ""}` }),
      me.overdue_days > 0 && !waiting ? el("p", { class: "warnline", role: "status", text: `⚠ ${t("overdue", me.overdue_days)}` }) : null,
      returned,
      me.locked ? null : howto,
      me.step === "submission" || me.locked ? null : el("div", { class: "excelrow" },
        el("span", { class: "muted", text: me.excel_upload ? t("excel_or") : "" }),
        el("a", { class: "pbtn", href: "/api/p/template.xlsx", text: t("download") }),
        me.excel_upload ? el("button", { type: "button", class: "pbtn", text: t("upload"), onclick: openUpload }) : null)),
    lockBanner);
  if (me.step === "submission") { if (prog) prog.hidden = true; put(app, el("section", { class: "pcard", id: "docs" }), sendbar); loadDocs(); return; }
  put(app, toolbar(), el("div", { id: "list", class: `skus view-${P.view}` }), el("div", { id: "more" }), sendbar);
  load(true);
}

const VIEWS = () => (window.innerWidth < 700 ? ["list", "one"] : ["table", "list", "one"]);

function toolbar() {
  const chips = el("div", { class: "chips", role: "tablist", id: "chips" });
  const search = el("input", { type: "search", class: "psearch", placeholder: t("search"), value: P.q || null, "aria-label": t("search") });
  let timer;
  search.addEventListener("input", () => { clearTimeout(timer); timer = setTimeout(() => { P.q = search.value.trim(); load(true); }, 250); });
  const views = el("div", { class: "seg", role: "group" },
    VIEWS().map((v) => el("button", { type: "button", "aria-pressed": String(P.view === v), text: t(`view_${v}`),
      onclick: () => { P.view = v; try { localStorage.setItem("nego-p-view", v); } catch (_) { /* ignore */ } P.one = 0; renderList(); document.querySelectorAll(".seg button").forEach((b, i) => b.setAttribute("aria-pressed", String(VIEWS()[i] === v))); } })));
  const bulk = P.me.step !== "feedback1" || P.me.locked ? null :
    el("button", { type: "button", class: "pbtn ghost", text: t("fill_fb1"), onclick: bulkFill });
  return el("section", { class: "toolbar" }, chips,
    el("div", { class: "tools" }, search, views, bulk,
      P.me.locked ? null : el("button", { type: "button", class: "pbtn ghost", text: t("next_todo"), onclick: nextTodo })));
}

function renderChips() {
  const box = $("chips");
  if (!box) return;
  put(clear(box), ...["todo", "check", "up", "all", "done"].filter((k) => (k !== "up" || P.me.step === "rfq") && (!P.me.locked || k === "all")).map((k) =>
    el("button", { type: "button", role: "tab", class: `chip ${P.filter === k ? "on" : ""} c-${k}`, "aria-selected": String(P.filter === k),
      onclick: () => { P.filter = k; load(true); } }, t(`f.${k}`), el("b", { text: num(P.counts[k] ?? 0) }))));
  const c = P.counts;
  const done = c.done ?? 0, all = c.all ?? 0;
  if ($("prog-fill")) {
    $("prog-fill").style.width = `${all ? (done / all) * 100 : 0}%`;
    $("prog-text").textContent = t("progress", num(done), num(all));
  }
}

async function load(reset) {
  if (reset) { P.items = []; P.one = 0; }
  P.loading = true;
  const qs = new URLSearchParams({ filter: P.filter, q: P.q, offset: P.items.length, limit: 50 });
  let d;
  try { d = await call(`/api/p/items?${qs}`); } catch (e) { toast(e.message); return; } finally { P.loading = false; }
  if (reset && P.filter === "todo" && !d.total && d.counts.all && !load.switched) {
    load.switched = true; // nothing left to fill: show everything instead of an empty list
    P.filter = "all";
    return load(true);
  }
  P.items = P.items.concat(d.items);
  P.total = d.total;
  P.counts = d.counts;
  renderChips();
  renderList();
}

async function refreshCounts() {
  clearTimeout(refreshCounts.t);
  refreshCounts.t = setTimeout(async () => {
    try { const d = await call(`/api/p/items?${new URLSearchParams({ filter: "all", limit: 1 })}`); P.counts = d.counts; renderChips(); } catch (_) { /* next save retries */ }
  }, 500);
}

function renderList() {
  const list = clear($("list"));
  list.className = `skus view-${P.view}`;
  const more = clear($("more"));
  if (!P.items.length) { put(list, el("p", { class: "muted center pad", text: t("none") })); return; }
  if (P.view === "one") {
    P.one = Math.max(0, Math.min(P.one, P.items.length - 1));
    put(list, card(P.items[P.one]));
    put(more, el("div", { class: "onenav" },
      el("button", { type: "button", class: "pbtn", text: t("prev"), disabled: P.one === 0, onclick: () => { P.one--; renderList(); focusFirst(); } }),
      el("span", { class: "muted", text: `${P.one + 1} / ${num(P.total)}` }),
      el("button", { type: "button", class: "pbtn primary", text: t("next"), disabled: P.one >= P.total - 1, onclick: goNext })));
    return;
  }
  if (P.view === "table") {
    put(list, theader(), ...P.items.map(trow), el("p", { class: "muted small tip", text: t("tip_paste") }));
    list.addEventListener("keydown", tableKeys);
    list.addEventListener("paste", tablePaste);
  } else P.items.forEach((it) => put(list, card(it)));
  if (P.items.length < P.total) put(more, el("button", { type: "button", class: "pbtn wide", text: t("more"), onclick: () => load(false) }));
}

async function goNext() {
  if (P.one >= P.items.length - 1 && P.items.length < P.total) await load(false);
  P.one = Math.min(P.one + 1, P.items.length - 1);
  renderList();
  focusFirst();
}
function focusFirst() { const i = $("list").querySelector("input:not([disabled]), button.sbtn"); if (i) i.focus(); }

function nextTodo() {
  if (P.view === "one") {
    const i = P.items.findIndex((x, k) => k > P.one && x.state === "missing");
    if (i >= 0) { P.one = i; renderList(); focusFirst(); return; }
  } else {
    const a = $("list").querySelector("article.st-missing, .trow.st-missing");
    if (a) { a.scrollIntoView({ behavior: "smooth", block: "center" }); const i = a.querySelector("input, button.sbtn"); if (i) setTimeout(() => i.focus(), 300); return; }
  }
  P.filter = "todo"; load(true);
}

async function bulkFill() {
  if (!confirm(t("fill_confirm_fb1"))) return;
  try { const r = await call("/api/p/fill", { method: "POST", body: {} }); toast(t("filled", r.filled)); load(true); } catch (e) { toast(e.message); }
}

// ---------- table view (looks and moves like an Excel sheet) ----------
const TCOLS = {
  rfq: ["item", "status", "brand", "ref", "hna", "qty", "disc", "unit"],
  feedback1: ["item", "rfq", "co", "accept", "own", "unit"],
};
const HEAD = { item: "h_item", status: "h_status", brand: "h_brand", ref: "h_ref", remarks: "h_remarks", hna: "h_hna", qty: "h_qty", disc: "h_disc", unit: "h_unit", rfq: "h_rfq", co: "h_co", accept: "", own: "h_own" };

function theader() {
  return el("div", { class: `thead tc-${P.me.step}`, role: "row" },
    TCOLS[P.me.step].map((c) => el("span", { role: "columnheader", text: c === "ref" && P.me.category === "Drugs" ? t("ai") : HEAD[c] ? t(HEAD[c]) : "" })));
}

function statusSelect(a, it) {
  const sel = el("select", { "aria-label": t("status"), "data-col": "status" },
    el("option", { value: "Active", text: t("active"), selected: it.item_status === "Active" }),
    el("option", { value: "Discontinue", text: t("disc_btn"), selected: it.item_status === "Discontinue" }));
  if (!it.item_status) sel.insertBefore(el("option", { value: "", text: "—", selected: true }), sel.firstChild);
  sel.addEventListener("change", () => { if (sel.value) save(a, it, { item_status: sel.value }); });
  sel.disabled = !!P.me.locked;
  return sel;
}

function tinput(a, it, col, field, attrs, toServer, show) {
  const i = el("input", { "data-col": col, value: show(it[field]) || null, ...attrs });
  i.addEventListener("change", () => save(a, it, { [field]: toServer(i.value) || null }));
  if (P.me.locked) i.disabled = true;
  return i;
}

function trow(it) {
  const step = P.me.step;
  const a = el("div", { class: `trow st-${it.state}`, "data-id": it.id, role: "row" });
  const off = it.item_status === "Discontinue";
  const cell = {};
  cell.item = el("div", { class: "tc-item" },
    el("b", { text: it.item_name || "—" }),
    el("span", { class: "meta", text: [it.erp_code, it.brand, it.catalog_no || it.active_ingredient].filter(Boolean).join(" · ") }),
    el("span", { class: `state s-${it.state}`, text: t(`state.${it.state}`) }),
    el("span", { class: "savestate", "aria-live": "polite" }));
  cell.status = statusSelect(a, it);
  if (step === "rfq") {
    const plain = (field, col) => tinput(a, it, col, field, { "aria-label": t(HEAD[col]) }, (v) => v.trim(), (v) => v || "");
    cell.brand = plain("brand", "brand"); cell.ref = plain(P.me.category === "Drugs" ? "active_ingredient" : "catalog_no", "ref");
    const hna = tinput(a, it, "hna", "rfq_hna", { inputmode: "decimal", "aria-label": t("h_hna") }, hnaValue, hnaShow);
    const qty = tinput(a, it, "qty", "rfq_qty", { inputmode: "numeric", "aria-label": t("h_qty") }, digits, (v) => (v != null ? String(v) : ""));
    const disc = tinput(a, it, "disc", "rfq_disc", { inputmode: "decimal", "aria-label": t("h_disc") }, discValue, discShow);
    for (const i of [hna, qty, disc, cell.brand, cell.ref]) i.disabled = off || !!P.me.locked;
    const unit = el("b", { class: "tc-unit" });
    const live = () => {
      const n = (x) => +String(x).replace(",", ".");
      const pr = unitPrice(n(hnaValue(hna.value)) || null, +digits(qty.value) || null, discFraction(disc.value));
      const bad = Number.isNaN(discFraction(disc.value));
      unit.textContent = off || bad ? "—" : pr == null ? "—" : rp(pr);
    };
    hna.addEventListener("input", () => { const v = hnaValue(hna.value); hna.value = v ? hnaShow(v) : ""; live(); });
    qty.addEventListener("input", () => { qty.value = digits(qty.value); live(); });
    disc.addEventListener("input", () => { disc.value = disc.value.replace(/[^\d.,%]/g, ""); live(); });
    disc.addEventListener("blur", () => { const f = discFraction(disc.value); if (disc.value && !Number.isNaN(f)) disc.value = discShow(Math.round(f * 10000) / 10000); });
    cell.hna = el("div", { class: "affix" }, el("span", { text: "Rp" }), hna);
    cell.qty = qty; cell.disc = disc; disc.placeholder = "15";
    cell.unit = unit;
    live();
  } else {
    const has = it.co_disc != null;
    cell.rfq = el("span", { text: rp(it.rfq_unit_price) });
    cell.co = has ? el("span", {}, `${pctTxt(it.co_disc)} → `, el("b", { text: rp(it.co_unit_price) })) : el("span", { class: "muted", text: "—" });
    const accepted = has && it.fb1_disc != null && Math.abs(it.fb1_disc - it.co_disc) < 1e-9;
    cell.accept = has ? el("button", { type: "button", class: `pbtn small ${accepted ? "on" : "primary"}`, text: accepted ? `✓ ${t("accept")}` : t("accept"),
      onclick: () => save(a, it, { fb1_disc: "__co__" }, true), disabled: !!P.me.locked }) : el("span");
    const own = tinput(a, it, "own", "fb1_disc", { inputmode: "decimal", "aria-label": t("h_own") }, discValue, discShow);
    own.disabled = !has || !!P.me.locked;
    cell.unit = el("b", { class: "tc-unit", text: has ? rp(it.fb1_unit_price) : "—" });
    own.addEventListener("input", () => {
      own.value = own.value.replace(/[^\d.,%]/g, "");
      const pr = unitPrice(it.rfq_hna, it.rfq_qty, discFraction(own.value));
      cell.unit.textContent = pr == null ? "—" : rp(pr);
    });
    cell.own = own; own.placeholder = "15";
  }
  put(a, el("div", { class: `tcells tc-${step}` }, TCOLS[step].map((c) => el("div", { class: `td td-${c}`, role: "cell" }, cell[c]))), el("div", { class: "issues" }));
  paintIssues(a, it);
  return a;
}

// Enter moves down the same column, like a sheet.
function tableKeys(e) {
  if (e.key !== "Enter" || e.target.tagName !== "INPUT") return;
  e.preventDefault();
  e.target.dispatchEvent(new Event("change"));
  const col = e.target.dataset.col, row = e.target.closest(".trow");
  for (let r = row && row.nextElementSibling; r; r = r.nextElementSibling) {
    const n = r.classList.contains("trow") && r.querySelector(`input[data-col="${col}"]:not([disabled])`);
    if (n) { n.focus(); n.select(); return; }
  }
}

// Pasting several cells copied from Excel fills across and down from the focused cell.
function tablePaste(e) {
  const target = e.target;
  if (target.tagName !== "INPUT" || !target.dataset.col) return;
  const text = (e.clipboardData || window.clipboardData).getData("text");
  if (!/[\t\n]/.test(text.replace(/\r?\n$/, ""))) return;
  e.preventDefault();
  const grid = text.replace(/\r?\n$/, "").split(/\r?\n/).map((r) => r.split("\t"));
  let row = target.closest(".trow");
  const cols = [...row.querySelectorAll("input[data-col]")].map((i) => i.dataset.col);
  const c0 = cols.indexOf(target.dataset.col);
  for (const cells of grid) {
    if (!row) break;
    cells.forEach((v, j) => {
      const i = row.querySelector(`input[data-col="${cols[c0 + j]}"]:not([disabled])`);
      if (!i) return;
      i.value = v.trim();
      i.dispatchEvent(new Event("input"));
      i.dispatchEvent(new Event("change"));
    });
    do { row = row.nextElementSibling; } while (row && !row.classList.contains("trow"));
  }
}

// ---------- help ----------
function openHelp() {
  const h = (P.me && P.me.help) || {};
  const steps = P.me && P.me.step ? t(`howto.${P.me.step}`) : [];
  const wa = String(h.whatsapp || "").replace(/[^\d]/g, "");
  const links = [wa ? el("a", { class: "pbtn primary", href: `https://wa.me/${wa}`, target: "_blank", rel: "noopener", text: t("help_wa") }) : null,
    h.email ? el("a", { class: "pbtn", href: `mailto:${h.email}`, text: t("help_mail") }) : null].filter(Boolean);
  put(clear($("dlg-body")), el("h2", { text: t("help_t") }),
    steps.length ? el("ol", { class: "howto" }, steps.map((x) => el("li", { text: x }))) : null,
    el("p", { text: links.length ? t("help_p") : t("help_none") }),
    el("div", { class: "pactions" }, ...links, el("button", { type: "button", class: "pbtn", text: t("close"), onclick: () => $("dlg").close() })));
  $("dlg").showModal();
}

// ---------- one SKU ----------
function card(it) {
  const a = el("article", { class: `sku st-${it.state}`, "data-id": it.id });
  const head = el("header", { class: "sku-head" },
    el("div", {}, el("h2", { text: it.item_name || "—" }),
      el("p", { class: "meta", text: [it.erp_code, it.brand, it.catalog_no || it.active_ingredient].filter(Boolean).join(" · ") })),
    el("span", { class: `state s-${it.state}`, text: t(`state.${it.state}`) }));
  put(a, head);
  const body = el("div", { class: "sku-body" });
  put(a, body);
  ({ rfq: rfqBody, feedback1: fbBody })[P.me.step](body, it, a);
  put(a, el("div", { class: "issues" }), el("div", { class: "savestate", "aria-live": "polite" }));
  paintIssues(a, it);
  a.addEventListener("keydown", (e) => {
    if (e.key !== "Enter" || e.target.tagName !== "INPUT") return;
    e.preventDefault();
    e.target.dispatchEvent(new Event("change"));
    const inputs = [...$("list").querySelectorAll("input:not([disabled]):not([type=search])")];
    const i = inputs.indexOf(e.target);
    if (i >= 0 && i < inputs.length - 1) inputs[i + 1].focus();
    else if (P.view === "one") goNext();
  });
  return a;
}

function statusButtons(a, it) {
  const mk = (val, label) => el("button", { type: "button", class: `sbtn ${it.item_status === val ? "on" : ""}`, "aria-pressed": String(it.item_status === val), text: label,
    disabled: !!P.me.locked, onclick: () => save(a, it, { item_status: val }) });
  return el("div", { class: "field status" }, el("span", { class: "lbl", text: t("status") }),
    el("div", { class: "seg big" }, mk("Active", t("active")), mk("Discontinue", t("disc_btn"))));
}

function rfqBody(body, it, a) {
  const pack = P.lang === "id" ? "kemasan" : "pack";
  const input = (field, label, attrs = {}) => {
    const i = el("input", { value: it[field] || null, disabled: !!P.me.locked || it.item_status === "Discontinue", ...attrs });
    i.addEventListener("change", () => { if ((i.value.trim() || null) !== (it[field] || null)) save(a, it, { [field]: i.value.trim() }); });
    return el("label", { class: "field" }, el("span", { class: "lbl", text: label }), i);
  };
  put(body, statusButtons(a, it));
  if (it.item_status === "Discontinue") return;
  put(body, el("div", { class: "grid3" }, input("brand", t("brand")), P.me.category === "Drugs" ? input("active_ingredient", t("ai")) : input("catalog_no", t("ref")), input("remarks", t("remarks"))));
  const lock = !!P.me.locked;
  const qty = el("input", { inputmode: "numeric", value: it.rfq_qty != null ? String(it.rfq_qty) : null, "aria-label": t("qty", pack), disabled: lock });
  const hna = el("input", { inputmode: "decimal", value: hnaShow(it.rfq_hna) || null, "aria-label": t("hna", pack), disabled: lock });
  const disc = el("input", { inputmode: "decimal", value: it.rfq_disc != null ? discShow(it.rfq_disc) : null, "aria-label": t("disc"), placeholder: "15", disabled: lock });
  const result = el("div", { class: "result" });
  const live = () => {
    const n = (x) => +String(x).replace(",", ".");
    const p = unitPrice(n(hnaValue(hna.value)) || null, +digits(qty.value) || null, discFraction(disc.value));
    put(clear(result), el("span", { class: "lbl", text: t("per_pcs") }), el("b", { text: p == null ? "—" : rp(p) }));
  };
  qty.addEventListener("input", () => { qty.value = digits(qty.value); live(); });
  hna.addEventListener("input", () => { const v = hnaValue(hna.value); hna.value = v ? hnaShow(v) : ""; live(); });
  disc.addEventListener("input", () => { disc.value = disc.value.replace(/[^\d.,%]/g, ""); live(); });
  disc.addEventListener("blur", () => { const f = discFraction(disc.value); if (disc.value && !Number.isNaN(f)) disc.value = discShow(Math.round(f * 10000) / 10000); });
  qty.addEventListener("change", () => save(a, it, { rfq_qty: digits(qty.value) || null }));
  hna.addEventListener("change", () => save(a, it, { rfq_hna: hnaValue(hna.value) || null }));
  disc.addEventListener("change", () => save(a, it, { rfq_disc: discValue(disc.value) || null }));
  put(body, el("div", { class: "grid3" },
    el("label", { class: "field" }, el("span", { class: "lbl", text: t("hna", pack) }), el("div", { class: "affix" }, el("span", { text: "Rp" }), hna)),
    el("label", { class: "field" }, el("span", { class: "lbl", text: t("qty", pack) }), el("div", { class: "affix" }, qty, el("span", { text: "pcs" }))),
    el("label", { class: "field" }, el("span", { class: "lbl", text: t("disc") }), disc)),
  result);
  live();
}

function fbBody(body, it, a) {
  if (it.co_disc == null) { put(body, el("p", { class: "muted", text: "—" })); return; }
  const disc = el("input", { inputmode: "decimal", value: it.fb1_disc != null ? discShow(it.fb1_disc) : null, "aria-label": t("own"), disabled: !!P.me.locked });
  const fbPrice = el("b", { text: rp(it.fb1_unit_price) });
  disc.addEventListener("input", () => {
    disc.value = disc.value.replace(/[^\d.,%]/g, "");
    const pr = unitPrice(it.rfq_hna, it.rfq_qty, discFraction(disc.value));
    fbPrice.textContent = pr == null ? "—" : rp(pr);
  });
  disc.addEventListener("blur", () => { const f = discFraction(disc.value); if (disc.value && !Number.isNaN(f)) disc.value = discShow(Math.round(f * 10000) / 10000); });
  disc.addEventListener("change", () => save(a, it, { fb1_disc: discValue(disc.value) || null }));
  const accepted = it.fb1_disc != null && Math.abs(it.fb1_disc - it.co_disc) < 1e-9;
  put(body, 
    el("div", { class: "ref" }, el("span", { class: "lbl", text: t("rfq_price") }), el("span", {}, `HNA ${rp(it.rfq_hna)} · ${t("disc")} ${pctTxt(it.rfq_disc || 0)} → `, el("b", { text: `${rp(it.rfq_unit_price)}/pcs` }))),
    el("div", { class: "ref co" }, el("span", { class: "lbl", text: t("co") }), el("span", {}, `${t("disc")} ${pctTxt(it.co_disc)} → `, el("b", { text: `${rp(it.co_unit_price)}/pcs` })),
      el("button", { type: "button", class: `pbtn small ${accepted ? "on" : "primary"}`, text: accepted ? `✓ ${t("accept")}` : t("accept"),
        disabled: !!P.me.locked, onclick: () => save(a, it, { fb1_disc: "__co__" }, true) })),
    el("label", { class: "field narrow" }, el("span", { class: "lbl", text: t("own") }), disc),
    el("div", { class: "result" }, el("span", { class: "lbl", text: t("per_pcs") }), fbPrice));
}

// A sent form is read-only: a change is requested with a comment on the SKU, which Siloam sees on the item and in Findings.
function paintComments(a, it, box) {
  const list = it.comments || [];
  for (const m of list) put(box, el("p", { class: "small muted", text: `💬 ${m.text} · ${m.resolved ? t("comment_done") : t("comment_open")}` }));
  const ta = el("textarea", { rows: 2, placeholder: t("comment_ph"), maxlength: 500, hidden: true });
  const send = el("button", { type: "button", class: "pbtn small primary", text: t("comment_send"), hidden: true });
  const open = el("button", { type: "button", class: "pbtn small ghost", text: t("comment_btn"), onclick: () => { ta.hidden = false; send.hidden = false; open.hidden = true; ta.focus(); } });
  send.addEventListener("click", async () => {
    if (!ta.value.trim()) return;
    try { const r = await call(`/api/p/items/${it.id}/comment`, { method: "POST", body: { text: ta.value.trim() } }); it.comments = r.comments; toast(t("comment_sent")); clear(box); paintComments(a, it, box); } catch (e) { toast(e.message); }
  });
  put(box, open, ta, send);
}

function paintIssues(a, it) {
  const box = clear(a.querySelector(".issues"));
  if (P.me.locked) { if (P.me.lock_reason === "sent") paintComments(a, it, box); return; } // sent: read-only, comments only
  const open = new Set(it.outstanding);
  const confirmed = new Set((it.principal_confirmed || "").split(",").filter(Boolean));
  for (const i of it.issues) {
    if (i.level === "missing") continue; // the state badge says it
    const row = el("div", { class: `issue lv-${i.level} ${open.has(i.code) ? "" : "resolved"}` }, el("p", { text: msg(i) }));
    if (i.level === "reason") put(row, reasonPicker(a, it));
    if (i.level === "confirm") {
      put(row, confirmed.has(i.code) ? el("span", { class: "ok", text: `✓ ${t("confirmed")}` }) :
        el("button", { type: "button", class: "pbtn small", text: t("yes"), disabled: !!P.me.locked, onclick: () => save(a, it, {}, true, i.code) }));
    }
    put(box, row);
  }
}

function reasonPicker(a, it) {
  const [pick0, ...rest] = (it.price_reason || "").split(" — ");
  const known = P.me.reasons.find((r) => r.key === pick0 || r.id === pick0 || r.en === pick0);
  const sel = el("select", { "aria-label": t("reason"), disabled: !!P.me.locked }, el("option", { value: "", text: t("reason_pick") }),
    P.me.reasons.map((r) => el("option", { value: r.key, text: r[P.lang], selected: known && known.key === r.key })));
  const detail = el("input", { placeholder: t("reason_detail"), value: known ? rest.join(" — ") || null : it.price_reason || null, maxlength: 200, disabled: !!P.me.locked });
  const push = () => {
    const v = [sel.value, detail.value.trim()].filter(Boolean).join(" — ");
    if (v !== (it.price_reason || "")) save(a, it, { price_reason: v || null });
  };
  sel.addEventListener("change", push);
  detail.addEventListener("change", push);
  return el("div", { class: "reason" }, sel, detail);
}

// Saves one SKU. Unsaved changes are kept in this browser and retried, so nothing typed is lost.
const PENDING = "nego-p-pending";
function pending() { try { return JSON.parse(localStorage.getItem(PENDING) || "{}"); } catch (_) { return {}; } }
function setPending(p) { try { localStorage.setItem(PENDING, JSON.stringify(p)); } catch (_) { /* ignore */ } }

async function save(a, it, changes, rerender = false, confirmCode = null) {
  const st = a.querySelector(".savestate");
  st.className = "savestate saving"; st.textContent = t("saving");
  const p = pending(); p[it.id] = { ...(p[it.id] || {}), ...changes }; setPending(p);
  try {
    const fresh = await call(`/api/p/items/${it.id}`, { method: "PATCH", body: { changes: p[it.id], confirm: confirmCode } });
    const q = pending(); delete q[it.id]; setPending(q);
    Object.assign(it, fresh);
    const statusChanged = "item_status" in changes;
    if (rerender || statusChanged) {
      const next = a.classList.contains("trow") ? trow(it) : card(it);
      a.replaceWith(next);
      a = next;
    } else {
      a.className = a.className.replace(/\bst-\w+/, `st-${it.state}`);
      const s = a.querySelector(".state"); s.className = `state s-${it.state}`; s.textContent = t(`state.${it.state}`);
      paintIssues(a, it);
    }
    const st2 = a.querySelector(".savestate"); st2.className = "savestate ok"; st2.textContent = t("saved");
    refreshCounts();
  } catch (e) {
    // A refused value (validation, locked) is not queued for retry; only network failures are.
    if (e.status >= 400 && e.status < 500) { const q = pending(); delete q[it.id]; setPending(q); }
    const st2 = a.isConnected ? a.querySelector(".savestate") : st;
    st2.className = "savestate bad"; put(clear(st2), `${t("unsaved")}: ${e.message} `,
      el("button", { type: "button", class: "pbtn small", text: t("retry"), onclick: () => save(a, it, {}, rerender, confirmCode) }));
  }
}

// ---------- documents (last step) ----------
async function loadDocs() {
  let d;
  try { d = await call("/api/p/documents"); } catch (e) { toast(e.message); return; }
  drawDocs(d);
}

function drawDocs(d) {
  const box = clear($("docs"));
  put(box, el("p", { class: "muted", text: t("doc_note") }));
  const list = el("div", { class: "doclist-p" });
  for (const ty of d.types) {
    const mine = d.documents.filter((x) => x.doc_type === ty.key);
    const file = el("input", { type: "file", accept: ".pdf,.jpg,.jpeg,.png", class: "sr-only", id: `f-${ty.key}` });
    file.addEventListener("change", async () => {
      if (!file.files.length) return;
      const fd = new FormData(); fd.append("file", file.files[0]); fd.append("doc_type", ty.key);
      try { drawDocs(await call("/api/p/documents", { method: "POST", form: fd })); toast(t("doc_uploaded")); } catch (e) { toast(e.message); }
    });
    put(list, el("div", { class: `docrow ${mine.length ? "have" : ty.required ? "need" : ""}` },
      el("div", {}, el("b", { text: ty[P.lang] }), el("span", { class: `tag ${ty.required ? "req" : ""}`, text: ty.required ? t("doc_required") : t("doc_optional") }),
        mine.map((x) => el("div", { class: "small" }, `✓ ${x.filename} `,
          el("button", { type: "button", class: "pbtn small ghost", text: t("doc_delete"), onclick: async () => {
            try { drawDocs(await call(`/api/p/documents/${x.id}`, { method: "DELETE" })); } catch (e) { toast(e.message); } } })))),
      el("label", { class: "pbtn small", for: `f-${ty.key}` }, t("doc_upload")), file));
  }
  put(box, list);
}

// Optional attachment for Discontinue SKUs (DC letter or another document), uploaded from the send dialog.
function dcAttach(s) {
  const file = el("input", { type: "file", accept: ".pdf,.jpg,.jpeg,.png", class: "sr-only", id: "dcfile" });
  const note = el("span", { class: "small muted", text: s.dc_documents ? t("dc_attached", s.dc_documents) : "" });
  file.addEventListener("change", async () => {
    if (!file.files.length) return;
    const fd = new FormData(); fd.append("file", file.files[0]); fd.append("doc_type", "dc_letter");
    try { const d = await call("/api/p/documents", { method: "POST", form: fd }); note.textContent = t("dc_attached", d.documents.filter((x) => x.doc_type === "dc_letter").length); } catch (e) { toast(e.message); }
  });
  return el("div", { class: "pactions" }, el("label", { class: "pbtn small", for: "dcfile" }, t("dc_attach")), file, note);
}

// ---------- send ----------
async function openSummary() {
  let s;
  try { s = await call("/api/p/summary"); } catch (e) { toast(e.message); return; }
  const agree = el("input", { type: "checkbox", id: "agree" });
  const dcCheck = el("input", { type: "checkbox", id: "dcok" });
  const sendBtn = el("button", { type: "button", class: "pbtn primary", text: t("sum_send"), disabled: true });
  const ready = () => { sendBtn.disabled = !(agree.checked && s.can_send && (!s.needs_discontinue_confirm || dcCheck.checked)); };
  agree.addEventListener("change", ready);
  dcCheck.addEventListener("change", ready);
  sendBtn.addEventListener("click", async () => {
    sendBtn.disabled = true;
    try {
      await call("/api/p/submit", { method: "POST", body: { confirm_discontinue: dcCheck.checked } });
      $("dlg").close();
      toast(t("sent_ok"));
      P.me = await call("/api/p/me");
      render();
    } catch (e) { toast(e.message); sendBtn.disabled = false; }
  });
  const facts = s.step === "submission" ? [["doc_uploaded", s.documents], ["sum_missing", s.missing]]
    : [["sum_items", s.items], ["sum_done", s.done], ["sum_missing", s.missing], ["sum_check", s.to_check], ["sum_disc", s.discontinued]];
  const body = clear($("dlg-body"));
  put(body, el("h2", { text: t("sum_t") }),
    el("div", { class: "facts" }, facts.map(([k, v]) => el("div", { class: k === "sum_missing" && v ? "bad" : k === "sum_check" && v ? "warn" : "" }, el("span", { text: t(k) }), el("b", { text: num(v) })))),
    s.outstanding_count ? el("div", {}, el("h3", { text: `${t("sum_left")} (${num(s.outstanding_count)})` }),
      el("ul", { class: "left" }, s.outstanding.map((o) => el("li", {},
        el("div", {}, el("b", { text: o.item_name }), el("span", { class: "muted", text: ` ${o.erp_code || ""}` }),
          el("div", { class: "small", text: o.issues.map((i) => msg(i)).join(" ") })),
        el("button", { type: "button", class: "pbtn small", text: t("sum_fix"), onclick: () => { $("dlg").close(); if (s.step !== "submission") { P.filter = "all"; P.q = o.erp_code || o.item_name; render(); } } }))))) :
      el("p", { class: "okline", text: `✓ ${t("sum_ok")}` }),
    s.increase_count ? el("details", {}, el("summary", { text: `${t("sum_up")} (${num(s.increase_count)})` }),
      el("ul", { class: "left" }, s.increases.map((x) => el("li", {}, el("div", {}, el("b", { text: x.item_name }),
        el("div", { class: "small", text: x.reason || "—" })))))) : null,
    s.needs_discontinue_confirm ? el("div", { class: "dcbox" }, el("h3", { text: t("dc_t") }), el("p", { class: "small", text: t("dc_p", s.discontinued) }),
      el("ul", { class: "left" }, s.discontinued_list.slice(0, 30).map((x) => el("li", {}, el("b", { text: x.item_name }), el("span", { class: "muted", text: ` ${x.erp_code}` })))),
      dcAttach(s), el("label", { class: "agree" }, dcCheck, t("dc_confirm"))) : null,
    el("label", { class: "agree" }, agree, t("sum_agree")),
    el("div", { class: "pactions" }, sendBtn, el("button", { type: "button", class: "pbtn", text: t("cancel"), onclick: () => $("dlg").close() })));
  if (!s.can_send) agree.disabled = true;
  $("dlg").showModal();
}

// ---------- Excel ----------
function openUpload() {
  const file = el("input", { type: "file", accept: ".xlsx", id: "upfile" });
  const out = el("div", { "aria-live": "polite" });
  const check = el("button", { type: "button", class: "pbtn primary", text: t("up_check") });
  check.addEventListener("click", async () => {
    if (!file.files.length) { file.focus(); return; }
    const fd = new FormData(); fd.append("file", file.files[0]); fd.append("apply", "0");
    put(clear(out), el("p", { class: "muted", text: t("loading") }));
    try { showUpload(out, await call("/api/p/template", { method: "POST", form: fd }), file.files[0]); }
    catch (e) { put(clear(out), el("p", { class: "error", text: e.message })); }
  });
  put(clear($("dlg-body")), el("h2", { text: t("up_t") }), el("p", { class: "muted", text: t("up_p") }),
    el("div", { class: "pactions" }, file, check), out,
    el("div", { class: "pactions" }, el("button", { type: "button", class: "pbtn", text: t("close"), onclick: () => $("dlg").close() })));
  $("dlg").showModal();
}

function showUpload(out, r, f) {
  const bad = r.rejected.length + r.read_errors.length;
  const agree = el("input", { type: "checkbox", id: "ow" });
  const save = el("button", { type: "button", class: "pbtn primary", text: t("up_save", r.saved_rows ?? r.accepted), disabled: r.overwrite_count > 0, onclick: async () => {
    const fd = new FormData(); fd.append("file", f); fd.append("apply", "1"); if (r.overwrite_count) fd.append("confirm_overwrite", "1");
    try { const res = await call("/api/p/template", { method: "POST", form: fd }); toast(t("up_done", res.applied)); $("dlg").close(); load(true); }
    catch (e) { toast(e.message); }
  } });
  agree.addEventListener("change", () => { save.disabled = !agree.checked; });
  const fmtV = (v) => (v == null ? "—" : typeof v === "number" && v > 0 && v < 1 ? discShow(v) : String(v));
  put(clear(out), el("p", { class: bad ? "warnline" : "okline", text: t("up_res", r.saved_rows ?? r.accepted, bad) }),
    r.unmatched_count ? el("p", { class: "small muted", text: t("not_found", r.unmatched_count) }) : null,
    bad ? el("ul", { class: "left errs" }, [...r.read_errors, ...r.rejected].map((x) => el("li", {},
      el("b", { text: `${t("row")} ${x.row}` }), x.item_name ? el("span", { text: ` ${x.item_name}` }) : null,
      el("div", { class: "small", text: x.messages.map((m) => msg(m)).join(" ") })))) : null,
    r.changes && r.changes.length ? el("details", { open: r.overwrite_count > 0 }, el("summary", { text: `${t("up_changes")} (${num(r.changes.length)})` }),
      el("ul", { class: "small left" }, r.changes.slice(0, 100).map((c) => el("li", { text: `${c.erp_code} ${c.item_name} · ${c.header}: ${fmtV(c.old)} → ${fmtV(c.new)}` })))) : null,
    r.overwrite_count ? el("div", { class: "warnline" }, el("p", { text: t("up_overwrite", r.overwrite_count) }), el("label", { class: "agree" }, agree, t("up_overwrite_ok"))) : null,
    r.notes.length ? el("details", {}, el("summary", { text: `${r.notes.length} catatan / notes` }),
      el("ul", { class: "small" }, r.notes.map((n) => el("li", { text: `${t("row")} ${n.row}: ${msg(n)}` })))) : null,
    r.accepted || (r.changes && r.changes.length) ? save : null);
}

boot();
