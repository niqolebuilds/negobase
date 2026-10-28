// App shell: who is signed in, which page shows, the header.
// 3 Main Sections: Admin, Negotiation, Principal.
// Floating AI Assistant present across all views.

const shell = { user: null, catalog: null, intended: null, provider: null, aiHistory: [] };
const PAGES = ["welcome", "signin", "setpw", "admin", "nego", "principal", "chat", "renewals"];
const NEEDS_AUTH = new Set(["admin", "nego", "principal"]);

function toast(text) {
  const t = $("toast");
  t.textContent = text;
  t.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => { t.hidden = true; }, 3200);
}

function initials(name) {
  return (name || "?").split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join("");
}

function renderHeader() {
  const u = shell.user;
  $("apptabs").hidden = !u;
  $("who").hidden = !u;
  $("admin-tab").hidden = !(u && (u.role === "admin" || u.role === "negotiator"));
  $("admin-tab").textContent = u && u.role === "negotiator" ? "Master data" : "Admin";
  if (u) {
    $("who-name").textContent = u.name;
    $("who-role").textContent = { admin: "Administrator", negotiator: "Negotiator" }[u.role] || "Viewer";
    $("who-initials").textContent = initials(u.name);
  }
}

async function loadCatalog() {
  shell.catalog = await api("/api/catalog");
  state.currency = shell.catalog.currency || "IDR";
  if (shell.catalog.banner) { $("banner").textContent = shell.catalog.banner; $("banner").hidden = false; }
  else $("banner").hidden = true;
}

function currentRoute() {
  const [page, ...rest] = (location.hash.slice(1) || "welcome").split("/").map(decodeURIComponent);
  return { page: PAGES.includes(page) ? page : "welcome", rest };
}

async function route() {
  const { page, rest } = currentRoute();
  // A side panel never survives a change of page.
  if ($("drawer") && !$("drawer").hidden && typeof closeDrawer === "function") closeDrawer();
  if (NEEDS_AUTH.has(page) && !shell.user) {
    shell.intended = location.hash;
    location.hash = "#signin";
    return;
  }
  if (page === "chat" || page === "renewals") {
    location.hash = "#nego";
    return;
  }
  if (page === "admin" && shell.user && shell.user.role === "negotiator" && rest[0] !== "master") {
    location.hash = "#admin/master";
    return;
  }
  if (page === "signin" && shell.user) {
    location.hash = "#nego";
    return;
  }
  if (page === "setpw") setpwEnter(rest[0] || "");
  PAGES.forEach((p) => {
    const elPage = $(p);
    if (elPage) elPage.hidden = p !== page;
  });
  document.querySelectorAll(".apptabs a").forEach((a) => a.classList.toggle("on", a.dataset.page === page));
  document.body.dataset.page = page;
  hideTip();
  updateAiRoleContext(page);

  if (page === "signin") {
    const err = $("signin-error");
    if (err) err.hidden = true;
    setTimeout(() => $("signin-email").focus(), 30);
  }
  if (page === "nego") ngEnter(rest);
  if (page === "admin") adminEnter(rest[0] || "periods");
  if (page === "principal") principalEnter(rest[0] || "");
  window.scrollTo(0, 0);
}

function updateAiRoleContext(page) {
  const tag = $("ai-role-tag");
  if (!tag) return;
  if (page === "admin") tag.textContent = "Role: Administrator";
  else if (page === "principal") tag.textContent = "Role: Principal / Vendor";
  else tag.textContent = "Role: Negotiator";
}

// Beta: only accounts listed by the server (admins, when NEGO_ADMIN_PASSWORD is set) need a password.
function needsPassword(email) {
  const p = shell.provider || {};
  return Boolean(p.password_required) || (p.password_required_for || []).includes(String(email).toLowerCase());
}

async function signIn(e) {
  e.preventDefault();
  const email = $("signin-email").value.trim();
  const password = $("signin-password").value;
  const err = $("signin-error");
  err.hidden = true;
  if (!email) { $("signin-email").focus(); return; }
  if (!password && needsPassword(email)) { $("signin-password").focus(); return; }
  try {
    const d = await send("/api/auth/signin", "POST", { email, password });
    $("signin-password").value = "";
    await afterSignIn(d);
  } catch (ex) {
    err.textContent = ex.message;
    err.hidden = false;
  }
}

async function afterSignIn(d) {
  if (d.token) {
    localStorage.setItem("nego_token", d.token);
    sessionStorage.setItem("nego_token", d.token);
  }
  shell.user = d.user;
  await loadCatalog();
  renderHeader();
  toast(`Welcome, ${d.user.name.split(" ")[0]}`);
  const go = shell.intended && shell.intended !== "#signin" ? shell.intended : "#nego";
  shell.intended = null;
  location.hash = go;
}

// One-time link from an admin: #setpw/<token>.
async function setpwEnter(token) {
  const who = $("setpw-who"), err = $("setpw-error"), btn = $("setpw-submit");
  err.hidden = true;
  btn.disabled = true;
  setpwEnter.token = token;
  try {
    const d = await api(`/api/auth/setup?token=${encodeURIComponent(token)}`);
    who.textContent = `For ${d.name} (${d.email}). Use at least ${d.min_length} characters.`;
    btn.disabled = false;
    setTimeout(() => $("setpw-password").focus(), 30);
  } catch (ex) {
    who.textContent = "";
    err.textContent = ex.message;
    err.hidden = false;
  }
}

async function setPassword(e) {
  e.preventDefault();
  const pw = $("setpw-password").value, again = $("setpw-confirm").value, err = $("setpw-error");
  err.hidden = true;
  if (pw !== again) { err.textContent = "The two passwords don't match."; err.hidden = false; return; }
  try {
    const d = await send("/api/auth/setup", "POST", { token: setpwEnter.token, password: pw });
    $("setpw-password").value = ""; $("setpw-confirm").value = "";
    setpwEnter.token = null;
    history.replaceState(null, "", "#nego");
    await afterSignIn(d);
  } catch (ex) { err.textContent = ex.message; err.hidden = false; }
}

async function signOut() {
  localStorage.removeItem("nego_token");
  sessionStorage.removeItem("nego_token");
  try { await send("/api/auth/signout", "POST", {}); } catch (_) { /* already signed out */ }
  shell.user = null;
  renderHeader();
  location.hash = "#welcome";
}

// Floating AI Chat Assistant Wiring
function initFloatingAiAssistant() {
  const toggleBtn = $("ai-toggle-btn");
  const chatWindow = $("ai-chat-window");
  const closeBtn = $("ai-close-btn");
  const form = $("ai-composer-form");
  const input = $("ai-input");
  const body = $("ai-chat-body");
  const chips = $("ai-quick-chips");

  if (!toggleBtn || !chatWindow) return;

  toggleBtn.addEventListener("click", () => {
    chatWindow.hidden = !chatWindow.hidden;
    if (!chatWindow.hidden) {
      setTimeout(() => input.focus(), 50);
      body.scrollTop = body.scrollHeight;
    }
  });

  closeBtn.addEventListener("click", () => {
    chatWindow.hidden = true;
  });

  if (chips) {
    chips.querySelectorAll(".ai-chip").forEach((btn) => {
      btn.addEventListener("click", () => {
        const prompt = btn.dataset.prompt;
        if (prompt) {
          input.value = prompt;
          form.dispatchEvent(new Event("submit"));
        }
      });
    });
  }

  function formatAiText(raw) {
    const wrapper = el("div", { class: "ai-bubble" });
    const text = String(raw || "").trim();
    const paragraphs = text.split(/\n\s*\n/);
    paragraphs.forEach((para) => {
      const p = el("p");
      const lines = para.split("\n");
      lines.forEach((line, idx) => {
        if (idx > 0) p.appendChild(el("br"));
        const tokens = line.split(/(\*\*[^*]+\*\*|`[^`]+`)/g);
        tokens.forEach((tok) => {
          if (!tok) return;
          if (tok.startsWith("**") && tok.endsWith("**")) {
            p.appendChild(el("strong", { text: tok.slice(2, -2) }));
          } else if (tok.startsWith("`") && tok.endsWith("`")) {
            p.appendChild(el("code", { text: tok.slice(1, -1) }));
          } else {
            p.appendChild(document.createTextNode(tok));
          }
        });
      });
      wrapper.appendChild(p);
    });
    return wrapper;
  }

  form.addEventListener("submit", async (e) => {
    e.preventDefault();
    const text = input.value.trim();
    if (!text) return;
    input.value = "";

    // Append user message with speech bubble
    const userMsg = el("div", { class: "ai-msg ai-user" },
      el("div", { class: "ai-bubble" }, el("p", { text }))
    );
    body.append(userMsg);
    body.scrollTop = body.scrollHeight;

    // Append typing indicator with animated 3-dot pulse
    const typing = el("div", { class: "ai-msg ai-bot ai-typing" },
      el("div", { class: "ai-msg-avatar" }, "✨"),
      el("div", { class: "ai-typing-indicator" },
        el("span", { class: "ai-dot" }),
        el("span", { class: "ai-dot" }),
        el("span", { class: "ai-dot" })
      )
    );
    body.append(typing);
    body.scrollTop = body.scrollHeight;

    const { page } = currentRoute();
    const role = page === "admin" ? "admin" : page === "principal" ? "principal" : "negotiator";
    const cycleId = ng.cid || null;

    try {
      shell.aiHistory.push({ role: "user", text });
      const res = await send("/api/ai/chat", "POST", {
        message: text,
        role,
        cycleId,
        history: shell.aiHistory.slice(-6),
      });
      typing.remove();
      shell.aiHistory.push({ role: "model", text: res.reply });

      const botMsg = el("div", { class: "ai-msg ai-bot" },
        el("div", { class: "ai-msg-avatar" }, "✨"),
        formatAiText(res.reply)
      );
      body.append(botMsg);
      body.scrollTop = body.scrollHeight;
    } catch (err) {
      typing.remove();
      const errMsg = el("div", { class: "ai-msg ai-bot error" },
        el("div", { class: "ai-msg-avatar" }, "⚠️"),
        el("div", { class: "ai-bubble" },
          el("p", { text: `Maaf, terjadi gangguan: ${err.message}` })
        )
      );
      body.append(errMsg);
      body.scrollTop = body.scrollHeight;
    }
  });
}

async function principalEnter(cidParam) {
  const container = clear($("principal"));
  container.className = "principal-embed-view page-v";
  
  const navBar = el("div", { class: "principal-nav-bar" });
  const selectGroup = el("div", { class: "principal-selector-group" });
  const label = el("label", { for: "principal-cycle-select", text: "Viewing Principal:" });
  const select = el("select", { id: "principal-cycle-select" });
  selectGroup.append(label, select);

  const toolsGroup = el("div", { class: "principal-tools-group", style: "display: flex; gap: 10px; align-items: center;" });
  const openLink = el("a", { class: "act ghost small-btn", target: "_blank", rel: "noopener", text: "Open Standalone Portal ↗" });
  const roleNote = el("span", { class: "badge", text: "Isolated Principal Session (Zero internal benchmark leakage)" });
  toolsGroup.append(roleNote, openLink);

  navBar.append(selectGroup, toolsGroup);

  const iframe = el("iframe", { class: "principal-iframe-frame", src: "about:blank" });
  container.append(navBar, iframe);

  try {
    const d = await api("/api/nego/principals");
    const openCycles = (d.principals || []).filter(p => p.cycle && p.cycle.status === "open").map(p => ({
      cycle_id: p.cycle.id,
      name: `${p.name} (${p.category || "Principal"})`,
      step: p.cycle.step_label,
      token: p.cycle.active_token || ""
    }));

    if (!openCycles.length) {
      clear(container).append(
        navBar,
        el("div", { class: "empty-state pad" },
          el("h3", { text: "No active negotiation cycles" }),
          el("p", { class: "muted", text: "Open a negotiation in the Negotiations tab to preview or test the principal portal." }))
      );
      return;
    }

    clear(select).append(...openCycles.map(c => el("option", { value: c.cycle_id, text: `${c.name} — Step: ${c.step}` })));

    let activeCid = cidParam ? parseInt(cidParam, 10) : null;
    if (!activeCid || !openCycles.some(c => c.cycle_id === activeCid)) {
      activeCid = openCycles[0].cycle_id;
    }

    select.value = String(activeCid);
    const updateSrc = (cid) => {
      const cyc = openCycles.find(c => c.cycle_id === cid);
      const tkn = cyc?.token || "";
      const query = tkn ? `cycle_id=${cid}&token=${encodeURIComponent(tkn)}` : `cycle_id=${cid}`;
      iframe.src = `/p?${query}`;
      openLink.href = `/p?${query}`;
    };
    updateSrc(activeCid);

    select.addEventListener("change", () => {
      activeCid = parseInt(select.value, 10);
      location.hash = `#principal/${activeCid}`;
      updateSrc(activeCid);
    });
  } catch (err) {
    container.append(el("p", { class: "error pad", text: err.message }));
  }
}

async function shellInit() {
  try {
    const d = await api("/api/auth/me");
    shell.user = d.user;
    shell.provider = d.provider;
    $("signin-notice").textContent = d.provider ? d.provider.notice : "";
    $("signin-password-row").hidden = !(d.provider && (d.provider.password_required || (d.provider.password_required_for || []).length));
    if (shell.user) await loadCatalog();
  } catch (e) {
    $("welcome").prepend(el("p", { class: "error", text: `Can't reach the server: ${e.message}` }));
  }
  renderHeader();
  $("start").addEventListener("click", () => { location.hash = shell.user ? "#nego" : "#signin"; });
  $("signin-form").addEventListener("submit", signIn);
  $("setpw-form").addEventListener("submit", setPassword);
  $("signout").addEventListener("click", signOut);
  document.querySelectorAll(".demo-login-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const email = btn.dataset.email;
      if (email) {
        $("signin-email").value = email;
        $("signin-password").value = "";
        $("signin-error").hidden = true;
        const fakeEvt = { preventDefault: () => {} };
        signIn(fakeEvt);
      }
    });
  });
  window.addEventListener("nego:signed-out", () => {
    if (!shell.user) return;
    shell.user = null;
    renderHeader();
    shell.intended = location.hash;
    toast("Your session ended. Please sign in again.");
    location.hash = "#signin";
  });
  window.addEventListener("hashchange", route);
  initFloatingAiAssistant();
  route();
}
shellInit();
