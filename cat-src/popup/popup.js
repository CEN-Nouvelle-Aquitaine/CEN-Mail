/**
 * Mail-CAT CEN — popup.js v1.0.0
 */
"use strict";

// ─── Onglets ─────────────────────────────────────────────────
const tabs   = [...document.querySelectorAll(".tab")];
const panels = [...document.querySelectorAll(".panel")];

tabs.forEach(tab => {
  tab.addEventListener("click", () => {
    tabs.forEach(t   => t.classList.remove("active"));
    panels.forEach(p => p.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById(`panel-${tab.dataset.tab}`).classList.add("active");
    switch (tab.dataset.tab) {
      case "graph":  loadGraphState(); break;
      case "labels": loadLabels();     break;
      case "sync":   loadSyncAccounts(); break;
      case "tags":   loadTags();       break;
    }
  });
});

// ─── Helpers UI ──────────────────────────────────────────────
const send = msg => messenger.runtime.sendMessage(msg);

function setStatus(el, html, type = "info") { el.innerHTML = html; el.className = `status show ${type}`; }
function hideStatus(el) { el.className = "status"; }

function esc(s) {
  return String(s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}

// ─── Modale ──────────────────────────────────────────────────
const modalOverlay = document.getElementById("modal-overlay");
const modalTitle   = document.getElementById("modal-title");
const modalIcon    = document.getElementById("modal-icon");
const modalBody    = document.getElementById("modal-body");
const modalConfirm = document.getElementById("modal-confirm");
const modalCancel  = document.getElementById("modal-cancel");

function confirm(opts) {
  return new Promise(resolve => {
    modalIcon.textContent    = opts.icon ?? "⚠️";
    modalTitle.textContent   = opts.title ?? "Confirmer";
    modalBody.innerHTML      = opts.html ?? "";
    modalConfirm.textContent = opts.confirmLabel ?? "✓ Confirmer";
    modalConfirm.className   = `btn ${opts.confirmClass ?? "btn-primary"}`;
    modalOverlay.classList.add("show");
    const onConfirm = () => { modalOverlay.classList.remove("show"); cleanup(); resolve(true); };
    const onCancel  = () => { modalOverlay.classList.remove("show"); cleanup(); resolve(false); };
    const cleanup   = () => {
      modalConfirm.removeEventListener("click", onConfirm);
      modalCancel.removeEventListener("click",  onCancel);
    };
    modalConfirm.addEventListener("click", onConfirm);
    modalCancel.addEventListener("click",  onCancel);
  });
}

// ─── Listener broadcasts background ─────────────────────────
messenger.runtime.onMessage.addListener(msg => {
  switch (msg.type) {
    case "GRAPH_AUTH_OK":
      updateGraphAuthStatus(true);
      graphAuth.disabled = false;
      graphAuth.innerHTML = "🔑 Se connecter avec Microsoft";
      setStatus(graphStatus, `✅ Connecté ! Token valide ${Math.round(msg.expires_in/60)} minutes.`, "success");
      break;
    case "GRAPH_AUTH_ERROR":
      updateGraphAuthStatus(false);
      graphAuth.disabled = false;
      graphAuth.innerHTML = "🔑 Se connecter avec Microsoft";
      setStatus(graphStatus, "❌ " + msg.error, "error");
      break;

    case "SYNC_PROGRESS": {
      syncPhase.textContent = msg.label ?? "";
      if      (msg.phase === "scan_src")       syncBar.style.width = "20%";
      else if (msg.phase === "scan_src_done") { syncBar.style.width = "40%"; syncCount.textContent = msg.label; }
      else if (msg.phase === "scan_dst")       syncBar.style.width = "60%";
      else if (msg.phase === "scan_dst_done") { syncBar.style.width = "80%"; syncCount.textContent = msg.label; }
      else if (msg.phase === "analyse_done")   syncBar.style.width = "100%";
      break;
    }
    case "SYNC_ANALYSE_DONE":
      _syncResult = msg;
      renderSyncResults(msg);
      showSyncStep(3);
      break;
    case "SYNC_APPLY_PROGRESS": {
      const pct = msg.total > 0 ? Math.round((msg.done/msg.total)*100) : 0;
      syncApplyBar.style.width = pct + "%";
      syncApplyCount.textContent = `${msg.done} / ${msg.total}`;
      syncApplyPct.textContent   = pct + " %";
      break;
    }
    case "SYNC_APPLY_DONE":
      showSyncStep(3);
      setStatus(syncStatus,
        `✅ ${msg.done} message(s) mis à jour${msg.errors?.length ? ` · ${msg.errors.length} erreur(s)` : ""}`,
        msg.errors?.length ? "warning" : "success");
      break;
    case "SYNC_ERROR":
      showSyncStep(1);
      setStatus(syncStatus, "❌ " + msg.error, "error");
      break;

    case "GRAPH_APPLY_PROGRESS": {
      const gpct = msg.total > 0 ? Math.round((msg.done/msg.total)*100) : 0;
      syncApplyBar.style.width = gpct + "%";
      syncApplyCount.textContent = `${msg.done} / ${msg.total}`;
      syncApplyPct.textContent   = gpct + " %";
      break;
    }
    case "GRAPH_APPLY_DONE": {
      showSyncStep(3);
      const skip = msg.skipped ? ` · ${msg.skipped} non trouvés dans Graph` : "";
      const errs = msg.errors?.length ? ` · ${msg.errors.length} erreur(s)` : "";
      setStatus(syncStatus,
        `✅ ${msg.done}/${msg.total} catégories appliquées${skip}${errs}`,
        (!msg.errors?.length && !msg.skipped) ? "success" : "warning");
      break;
    }
    case "GRAPH_ERROR":
      showSyncStep(3);
      setStatus(syncStatus, "❌ Graph : " + msg.error, "error");
      break;
  }
});

// ═══════════════════════════════════════════════════════════════
// ONGLET GRAPH
// ═══════════════════════════════════════════════════════════════
const graphAuth       = document.getElementById("graph-auth");
const graphDisconnect = document.getElementById("graph-disconnect");
const graphAuthIcon   = document.getElementById("graph-auth-icon");
const graphAuthLabel  = document.getElementById("graph-auth-label");
const graphAuthSub    = document.getElementById("graph-auth-sub");
const graphStatus     = document.getElementById("graph-status");

async function loadGraphState() {
  const r = await send({ action:"graphIsAuthenticated" });
  updateGraphAuthStatus(r.authenticated);
}

function updateGraphAuthStatus(authenticated) {
  if (authenticated) {
    graphAuthIcon.textContent  = "🔓";
    graphAuthLabel.textContent = "Connecté à Microsoft 365";
    graphAuthLabel.style.color = "var(--success)";
    graphAuthSub.textContent   = "Token actif — les catégories seront appliquées via Graph.";
    graphAuth.style.display       = "none";
    graphDisconnect.style.display = "block";
  } else {
    graphAuthIcon.textContent  = "🔒";
    graphAuthLabel.textContent = "Non connecté";
    graphAuthLabel.style.color = "var(--text)";
    graphAuthSub.textContent   = "Cliquez 'Se connecter'. Sans connexion, les catégories s'appliquent via IMAP.";
    graphAuth.style.display       = "block";
    graphDisconnect.style.display = "none";
  }
  if (syncApplySmart && syncApplySmart.style.display !== "none") _updateApplyMode();
}

graphAuth.addEventListener("click", async () => {
  graphAuth.disabled = true;
  graphAuth.innerHTML = '<span class="spin"></span> Connexion…';
  setStatus(graphStatus, "La fenêtre Microsoft va s'ouvrir. Connectez-vous puis revenez ici.", "info");
  try {
    await send({ action:"graphAuthenticate" });
  } catch(e) {
    // broadcast gère GRAPH_AUTH_OK / GRAPH_AUTH_ERROR
  }
});

graphDisconnect.addEventListener("click", () => {
  updateGraphAuthStatus(false);
  hideStatus(graphStatus);
});

// ═══════════════════════════════════════════════════════════════
// ONGLET ÉTIQUETTES
// ═══════════════════════════════════════════════════════════════
const lblList   = document.getElementById("lbl-list");
const lblSave   = document.getElementById("lbl-save");
const lblStatus = document.getElementById("lbl-status");
let _labelsData = [];

async function loadLabels() {
  lblList.innerHTML = '<div style="padding:14px;text-align:center;color:var(--text-3)"><span class="spin"></span> Chargement…</div>';
  hideStatus(lblStatus);
  try {
    _labelsData = await send({ action:"getTbTagsWithMapping" });
    renderLabels();
  } catch(e) {
    lblList.innerHTML = `<div style="padding:12px;color:var(--error)">${esc(e.message)}</div>`;
  }
}

function renderLabels() {
  if (!_labelsData.length) {
    lblList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--text-3)">Aucune étiquette Thunderbird configurée.</div>';
    return;
  }
  lblList.innerHTML = _labelsData.map(t => `
    <div class="lbl-item">
      <div class="lbl-swatch" style="background:${esc(t.color||'#888')}"></div>
      <span class="lbl-name">${esc(t.tag)}</span>
      <span class="lbl-arrow">→</span>
      <input type="text" class="ol-input" data-key="${esc(t.key)}"
        value="${esc(t.olCategory && t.olCategory !== '__skip__' ? t.olCategory : '')}"
        placeholder="Nom exact dans Outlook…"
        style="flex:1;min-width:0;padding:4px 8px;font-size:12px;border:1px solid var(--border);border-radius:var(--radius-sm);background:var(--surface)">
    </div>
  `).join("");
}

lblSave.addEventListener("click", async () => {
  const mapping = {};
  lblList.querySelectorAll(".ol-input").forEach(inp => {
    const key = inp.dataset.key;
    const val = inp.value.trim();
    if (key) mapping[key] = val || "__skip__";
  });
  const mapped = Object.entries(mapping).filter(([,v]) => v !== "__skip__");
  const ok = await confirm({
    icon: "💾", title: "Sauvegarder le mapping",
    html: `
      <div class="highlight">${mapped.length} étiquette(s) mappée(s) vers des catégories Outlook.</div>
      <ul>
        ${_labelsData.map(t => {
          const cat = mapping[t.key];
          const label = cat && cat !== "__skip__"
            ? `→ <strong>${esc(cat)}</strong>`
            : `<span style="color:var(--text-3)">— ignorée</span>`;
          return `<li>${esc(t.tag)} ${label}</li>`;
        }).join("")}
      </ul>`,
    confirmLabel: "💾 Sauvegarder",
  });
  if (!ok) return;
  try {
    await send({ action:"saveMapping", mapping });
    setStatus(lblStatus, "✅ Mapping sauvegardé.", "success");
    setTimeout(() => hideStatus(lblStatus), 2500);
    _labelsData = await send({ action:"getTbTagsWithMapping" });
  } catch(e) {
    setStatus(lblStatus, "❌ " + e.message, "error");
  }
});

// Auto-création catégories Outlook
const lblAutoCreate = document.getElementById("lbl-auto-create");
lblAutoCreate.addEventListener("click", async () => {
  const authState = await send({ action:"graphIsAuthenticated" });
  if (!authState.authenticated) {
    const goGraph = await confirm({
      icon: "🔗", title: "Connexion Graph requise",
      html: `<div class="highlight">Pour créer automatiquement les catégories dans Outlook, connectez-vous d'abord dans l'onglet <strong>Connexion M365</strong>.</div>`,
      confirmLabel: "Aller à Connexion M365",
    });
    if (goGraph) {
      tabs.forEach(t => t.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));
      document.querySelector("[data-tab='graph']").classList.add("active");
      document.getElementById("panel-graph").classList.add("active");
      loadGraphState();
    }
    return;
  }

  const tags = await send({ action:"listTags" });
  if (!tags?.length) { setStatus(lblStatus, "Aucune étiquette à exporter.", "warning"); return; }

  const ok = await confirm({
    icon: "🔗", title: "Créer les catégories Outlook",
    html: `
      <div class="highlight"><strong>${tags.length} étiquette(s)</strong> seront créées comme catégories dans Outlook (celles qui existent déjà seront ignorées).</div>
      <ul>${tags.map(t => `<li><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${esc(t.color||'#888')};margin-right:5px"></span>${esc(t.tag)}</li>`).join("")}</ul>`,
    confirmLabel: "🔗 Créer dans Outlook",
    confirmClass: "btn-orange",
  });
  if (!ok) return;

  lblAutoCreate.disabled = true;
  lblAutoCreate.innerHTML = '<span class="spin"></span> Création…';
  try {
    const r = await send({ action:"autoCreateCategories" });
    if (r?.error) throw new Error(r.error);
    const createdList = r.created.length ? `${r.created.length} créée(s)` : "Aucune nouvelle catégorie";
    const skippedList = r.skipped.length ? ` · ${r.skipped.length} existante(s)` : "";
    setStatus(lblStatus, `✅ ${createdList}${skippedList}. Mapping sauvegardé.`, "success");
    _labelsData = await send({ action:"getTbTagsWithMapping" });
    renderLabels();
  } catch(e) {
    setStatus(lblStatus, "❌ " + e.message, "error");
  }
  lblAutoCreate.disabled = false;
  lblAutoCreate.innerHTML = "🔗 Créer automatiquement dans Outlook";
});

// ═══════════════════════════════════════════════════════════════
// ONGLET SYNCHRONISATION
// ═══════════════════════════════════════════════════════════════
const syncSrc          = document.getElementById("sync-src");
const syncDst          = document.getElementById("sync-dst");
const syncAnalyse      = document.getElementById("sync-analyse");
const syncStep1        = document.getElementById("sync-step1");
const syncStep2        = document.getElementById("sync-step2");
const syncStep3        = document.getElementById("sync-step3");
const syncStep4        = document.getElementById("sync-step4");
const syncBar          = document.getElementById("sync-bar");
const syncCount        = document.getElementById("sync-count");
const syncPct          = document.getElementById("sync-pct");
const syncPhase        = document.getElementById("sync-phase");
const syncSummary      = document.getElementById("sync-summary");
const syncCatList      = document.getElementById("sync-cat-list");
const syncUnmapped     = document.getElementById("sync-unmapped");
const syncUnmappedMsg  = document.getElementById("sync-unmapped-msg");
const syncApplySmart   = document.getElementById("sync-apply-smart");
const syncApplyModeLabel = document.getElementById("sync-apply-mode-label");
const syncApplyBar     = document.getElementById("sync-apply-bar");
const syncApplyCount   = document.getElementById("sync-apply-count");
const syncApplyPct     = document.getElementById("sync-apply-pct");
const syncReset        = document.getElementById("sync-reset");
const syncStatus       = document.getElementById("sync-status");

let _syncResult     = null;
let _syncCategories = [];
let _accOpen        = {};
let _syncFilter     = { text:"", from:null, to:null, sort:"date-desc" };

function showSyncStep(n) {
  [syncStep1, syncStep2, syncStep3, syncStep4].forEach((el, i) => {
    el.style.display = (i+1 === n) ? "block" : "none";
  });
}

async function loadSyncAccounts() {
  try {
    const accounts = await send({ action:"getAccounts" });
    [syncSrc, syncDst].forEach(sel => {
      sel.innerHTML = "<option value=''>— Sélectionner un compte —</option>";
      for (const acc of accounts) {
        if (acc.type === "none") continue;
        const opt = document.createElement("option");
        opt.value = acc.id;
        const icon = /outlook|microsoft|office365|hotmail|live/i.test(acc.name) ? "🏢" : "📬";
        opt.textContent = `${icon} ${acc.name}`;
        sel.appendChild(opt);
      }
    });
  } catch(e) { setStatus(syncStatus, "❌ " + e.message, "error"); }
}

syncAnalyse.addEventListener("click", async () => {
  const src = syncSrc.value, dst = syncDst.value;
  if (!src) { setStatus(syncStatus,"⚠️ Sélectionnez un compte source.","warning"); return; }
  if (!dst) { setStatus(syncStatus,"⚠️ Sélectionnez un compte destination.","warning"); return; }
  if (src === dst) { setStatus(syncStatus,"⚠️ Source et destination identiques.","warning"); return; }

  const srcLabel = syncSrc.options[syncSrc.selectedIndex].textContent.trim();
  const dstLabel = syncDst.options[syncDst.selectedIndex].textContent.trim();
  const mapping  = await send({ action:"loadMapping" });
  const mappedCount = Object.values(mapping||{}).filter(v => v && v !== "__skip__").length;

  const ok = await confirm({
    icon: "🔍", title: "Analyser les deux boîtes",
    html: `
      <div class="highlight"><strong>Aucune modification ne sera faite à cette étape.</strong></div>
      <ul>
        <li>📬 Source : <strong>${esc(srcLabel)}</strong></li>
        <li>📭 Destination : <strong>${esc(dstLabel)}</strong></li>
        <li>${mappedCount > 0
          ? `✅ ${mappedCount} étiquette(s) mappée(s)`
          : `⚠️ Aucun mapping — configurez l'onglet <strong>Correspondances</strong> d'abord`}</li>
      </ul>`,
    confirmLabel: "🔍 Lancer l'analyse",
  });
  if (!ok) return;

  hideStatus(syncStatus);
  showSyncStep(2);
  syncBar.style.width = "10%";
  syncCount.textContent = "Initialisation…";
  const r = await send({ action:"analyseBoxes", srcAccountId:src, dstAccountId:dst });
  if (r?.error) { setStatus(syncStatus, "❌ " + r.error, "error"); showSyncStep(1); }
});

let _notFoundByCategory = [];

function showNotFoundModal() {
  if (!_notFoundByCategory.length) return;
  const total = _notFoundByCategory.reduce((s,c) => s + c.notFoundMessages.length, 0);
  const rows  = _notFoundByCategory.map(cat => {
    const items = cat.notFoundMessages.map(m => {
      const d = m.date ? new Date(m.date).toLocaleDateString("fr-FR") : "—";
      return `<div style="padding:4px 0;border-bottom:1px solid var(--border);font-size:11.5px">
        <div style="font-weight:500">${esc(m.subject || "(sans sujet)")}</div>
        <div style="color:var(--text-3);font-size:10.5px">👤 ${esc(m.sender||"—")} · 📅 ${d}</div>
      </div>`;
    }).join("");
    return `<div style="margin-bottom:10px">
      <div style="font-weight:700;color:var(--accent);margin-bottom:4px">${esc(cat.olCategory)} (${cat.notFoundMessages.length})</div>
      ${items}
    </div>`;
  }).join("");
  confirm({
    icon: "⚠️", title: `${total} message(s) non trouvés côté Outlook`,
    html: `<div class="highlight" style="margin-bottom:10px">Ces messages ont une étiquette côté source mais n'ont pas été trouvés dans la boîte Outlook. Ils n'ont probablement pas encore été migrés.</div>${rows}`,
    confirmLabel: "Fermer", confirmClass: "btn-ghost",
  });
}

function renderSyncResults(r) {
  _notFoundByCategory = r.notFoundTotal ? r.categories.filter(c => c.notFoundMessages?.length) : [];
  const toApply = r.categories.reduce((s,c) => s + c.messages.length, 0);

  syncSummary.innerHTML = `
    <div class="sync-sum-card"><div class="sync-sum-num">${r.srcTotal}</div><div class="sync-sum-label">avec étiquettes (source)</div></div>
    <div class="sync-sum-card"><div class="sync-sum-num">${r.dstTotal}</div><div class="sync-sum-label">indexés (destination)</div></div>
    <div class="sync-sum-card"><div class="sync-sum-num" style="color:var(--accent)">${toApply}</div><div class="sync-sum-label">catégories à appliquer</div></div>
    ${r.notFoundTotal ? `
    <div class="sync-sum-card" id="sync-nf-card" style="cursor:pointer;border:1px solid var(--warning);border-radius:var(--radius)">
      <div class="sync-sum-num" style="color:var(--warning)">${r.notFoundTotal}</div>
      <div class="sync-sum-label" style="color:var(--warning)">non trouvés ℹ️</div>
    </div>` : ""}
  `;
  if (r.notFoundTotal) {
    const nfCard = document.getElementById("sync-nf-card");
    if (nfCard) nfCard.addEventListener("click", showNotFoundModal);
  }

  if (r.noMapping.length) {
    syncUnmapped.style.display = "block";
    syncUnmappedMsg.innerHTML = `⚠️ ${r.noMapping.length} étiquette(s) sans mapping — configurez l'onglet <strong>Correspondances</strong>.`;
  } else {
    syncUnmapped.style.display = "none";
  }

  if (!r.categories.length) {
    syncCatList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--text-3)">Aucune catégorie à appliquer — tout est à jour.</div>';
    syncApplySmart.style.display = "none";
    syncApplyModeLabel.style.display = "none";
    return;
  }

  _syncCategories = r.categories.map(cat => ({
    ...cat,
    messages: cat.messages.map(m => ({ ...m, selected: true }))
  }));

  renderDetailedList();
  updateSelCount();
  syncApplySmart.style.display = "";
  syncApplyModeLabel.style.display = "";
  _updateApplyMode();
}

async function _updateApplyMode() {
  try {
    const auth = await send({ action:"graphIsAuthenticated" });
    if (auth.authenticated) {
      syncApplySmart.className = "btn btn-orange";
      syncApplySmart.innerHTML = "🔗 Appliquer";
      syncApplyModeLabel.textContent = "via Graph (Outlook Online)";
      syncApplySmart.dataset.mode = "graph";
    } else {
      syncApplySmart.className = "btn btn-primary";
      syncApplySmart.innerHTML = "✓ Appliquer";
      syncApplyModeLabel.textContent = "via IMAP (Thunderbird)";
      syncApplySmart.dataset.mode = "imap";
    }
  } catch {
    syncApplySmart.className = "btn btn-primary";
    syncApplySmart.innerHTML = "✓ Appliquer";
    syncApplyModeLabel.textContent = "via IMAP";
    syncApplySmart.dataset.mode = "imap";
  }
}

function applyFilters(messages) {
  return messages.filter(m => {
    if (_syncFilter.text) {
      const q = _syncFilter.text.toLowerCase();
      if (!((m.subject||"").toLowerCase().includes(q) || (m.sender||"").toLowerCase().includes(q))) return false;
    }
    if (_syncFilter.from && m.date && new Date(m.date) < new Date(_syncFilter.from)) return false;
    if (_syncFilter.to   && m.date && new Date(m.date) > new Date(_syncFilter.to + "T23:59:59")) return false;
    return true;
  }).sort((a, b) => {
    switch (_syncFilter.sort) {
      case "date-asc":  return new Date(a.date||0) - new Date(b.date||0);
      case "date-desc": return new Date(b.date||0) - new Date(a.date||0);
      case "subject":   return (a.subject||"").localeCompare(b.subject||"");
      case "sender":    return (a.sender||"").localeCompare(b.sender||"");
      default: return 0;
    }
  });
}

function renderDetailedList() {
  if (!_syncCategories.length) return;
  let html = "";
  _syncCategories.forEach((cat, ci) => {
    const filtered = applyFilters(cat.messages);
    const selCount = filtered.filter(m => m.selected).length;
    const isOpen   = _accOpen[ci] !== false;
    html += `
      <div class="sync-acc-header" data-ci="${ci}">
        <input type="checkbox" class="cat-master-check" data-ci="${ci}"
          ${selCount === filtered.length ? "checked" : ""}
          ${selCount > 0 && selCount < filtered.length ? "indeterminate-js" : ""}
          onclick="event.stopPropagation(); toggleCatAll(${ci}, this.checked)">
        <span class="sync-acc-toggle ${isOpen ? "open" : ""}">▶</span>
        <span class="sync-acc-name">${esc(cat.olCategory)}</span>
        <span class="sync-acc-cnt">${filtered.length}</span>
        <span class="sync-acc-sel">${selCount} sélectionné(s)</span>
        ${cat.notFound ? `<span style="font-size:10px;color:var(--warning)">⚠️ ${cat.notFound} non migrés</span>` : ""}
      </div>
      <div class="sync-acc-subbar">
        <button class="btn btn-ghost btn-sm" style="padding:2px 7px;font-size:10.5px" onclick="setCatSel(${ci},true)">Tout ☑</button>
        <button class="btn btn-ghost btn-sm" style="padding:2px 7px;font-size:10.5px" onclick="setCatSel(${ci},false)">Tout ☐</button>
        <button class="btn btn-ghost btn-sm" style="padding:2px 7px;font-size:10.5px" onclick="invertCatSel(${ci})">Inverser</button>
      </div>
      <div class="sync-msg-list ${isOpen ? "open" : ""}" id="sync-msg-list-${ci}">
        ${filtered.length === 0
          ? `<div style="padding:10px 28px;font-size:11.5px;color:var(--text-3)">Aucun message ne correspond aux filtres.</div>`
          : filtered.map(m => {
              const origIdx = cat.messages.indexOf(m);
              const dateStr = m.date ? new Date(m.date).toLocaleDateString("fr-FR") : "—";
              return `
                <div class="sync-msg-item" data-ci="${ci}" data-mi="${origIdx}">
                  <input type="checkbox" ${m.selected ? "checked" : ""}>
                  <div class="sync-msg-body">
                    <div class="sync-msg-subject">${esc(m.subject || "(sans sujet)")}</div>
                    <div class="sync-msg-meta">
                      <span class="sync-msg-sender">👤 ${esc(m.sender||"—")}</span>
                      <span style="flex-shrink:0">📅 ${dateStr}</span>
                    </div>
                  </div>
                </div>`;
            }).join("")}
      </div>`;
  });
  syncCatList.innerHTML = html;

  syncCatList.querySelectorAll(".cat-master-check[indeterminate-js]").forEach(cb => { cb.indeterminate = true; });

  syncCatList.querySelectorAll(".sync-acc-header").forEach(hdr => {
    hdr.addEventListener("click", e => {
      if (e.target.type === "checkbox") return;
      const ci = +hdr.dataset.ci;
      _accOpen[ci] = !(_accOpen[ci] !== false);
      hdr.querySelector(".sync-acc-toggle").classList.toggle("open", _accOpen[ci]);
      document.getElementById(`sync-msg-list-${ci}`).classList.toggle("open", _accOpen[ci]);
    });
  });

  syncCatList.querySelectorAll(".sync-msg-item").forEach(item => {
    item.addEventListener("click", e => {
      const ci = +item.dataset.ci, mi = +item.dataset.mi;
      const cb = item.querySelector("input[type=checkbox]");
      const newVal = e.target === cb ? cb.checked : !_syncCategories[ci].messages[mi].selected;
      _syncCategories[ci].messages[mi].selected = newVal;
      if (cb) cb.checked = newVal;
      _updateCatHeader(ci);
      updateSelCount();
    });
  });
}

function _updateCatHeader(ci) {
  const cat = _syncCategories[ci]; if (!cat) return;
  const filtered = applyFilters(cat.messages);
  const selCount = filtered.filter(m => m.selected).length;
  const hdr = syncCatList.querySelector(`.sync-acc-header[data-ci="${ci}"]`); if (!hdr) return;
  const badge    = hdr.querySelector(".sync-acc-sel");
  const masterCb = hdr.querySelector(".cat-master-check");
  if (badge)    badge.textContent = `${selCount} sélectionné(s)`;
  if (masterCb) { masterCb.checked = selCount === filtered.length && filtered.length > 0; masterCb.indeterminate = selCount > 0 && selCount < filtered.length; }
}

window.toggleCatAll = function(ci, checked) {
  applyFilters(_syncCategories[ci].messages).forEach(m => {
    m.selected = checked;
    const item = syncCatList.querySelector(`.sync-msg-item[data-ci="${ci}"][data-mi="${_syncCategories[ci].messages.indexOf(m)}"]`);
    if (item) { const cb = item.querySelector("input"); if (cb) cb.checked = checked; }
  });
  _updateCatHeader(ci); updateSelCount();
};
window.setCatSel = function(ci, val) {
  applyFilters(_syncCategories[ci].messages).forEach(m => {
    m.selected = val;
    const item = syncCatList.querySelector(`.sync-msg-item[data-ci="${ci}"][data-mi="${_syncCategories[ci].messages.indexOf(m)}"]`);
    if (item) { const cb = item.querySelector("input"); if (cb) cb.checked = val; }
  });
  _updateCatHeader(ci); updateSelCount();
};
window.invertCatSel = function(ci) {
  applyFilters(_syncCategories[ci].messages).forEach(m => {
    m.selected = !m.selected;
    const item = syncCatList.querySelector(`.sync-msg-item[data-ci="${ci}"][data-mi="${_syncCategories[ci].messages.indexOf(m)}"]`);
    if (item) { const cb = item.querySelector("input"); if (cb) cb.checked = m.selected; }
  });
  _updateCatHeader(ci); updateSelCount();
};

function updateSelCount() {
  const total = _syncCategories.reduce((s,c) => s + applyFilters(c.messages).length, 0);
  const sel   = _syncCategories.reduce((s,c) => s + applyFilters(c.messages).filter(m => m.selected).length, 0);
  document.getElementById("sync-sel-count").textContent = `${sel} sélectionné(s) sur ${total}`;
}

function getSelectedForApply() {
  return _syncCategories
    .map(cat => ({ ...cat, messages: cat.messages.filter(m => m.selected) }))
    .filter(cat => cat.messages.length > 0);
}

// Filtres
document.getElementById("sync-filter-text").addEventListener("input", e => {
  _syncFilter.text = e.target.value; renderDetailedList(); updateSelCount();
});
document.getElementById("sync-filter-from").addEventListener("change", e => {
  _syncFilter.from = e.target.value || null; renderDetailedList(); updateSelCount();
});
document.getElementById("sync-filter-to").addEventListener("change", e => {
  _syncFilter.to = e.target.value || null; renderDetailedList(); updateSelCount();
});
document.getElementById("sync-filter-clear").addEventListener("click", () => {
  _syncFilter = { text:"", from:null, to:null, sort:_syncFilter.sort };
  document.getElementById("sync-filter-text").value = "";
  document.getElementById("sync-filter-from").value = "";
  document.getElementById("sync-filter-to").value   = "";
  renderDetailedList(); updateSelCount();
});
document.getElementById("sync-sort").addEventListener("change", e => {
  _syncFilter.sort = e.target.value; renderDetailedList(); updateSelCount();
});

document.getElementById("sync-sel-all").addEventListener("click", () => {
  _syncCategories.forEach((cat, ci) => { applyFilters(cat.messages).forEach(m => { m.selected = true; }); _updateCatHeader(ci); });
  syncCatList.querySelectorAll(".sync-msg-item input[type=checkbox]").forEach(cb => cb.checked = true);
  updateSelCount();
});
document.getElementById("sync-sel-none").addEventListener("click", () => {
  _syncCategories.forEach((cat, ci) => { applyFilters(cat.messages).forEach(m => { m.selected = false; }); _updateCatHeader(ci); });
  syncCatList.querySelectorAll(".sync-msg-item input[type=checkbox]").forEach(cb => cb.checked = false);
  updateSelCount();
});
document.getElementById("sync-sel-invert").addEventListener("click", () => {
  _syncCategories.forEach((cat, ci) => {
    applyFilters(cat.messages).forEach(m => {
      m.selected = !m.selected;
      const item = syncCatList.querySelector(`.sync-msg-item[data-ci="${ci}"][data-mi="${cat.messages.indexOf(m)}"]`);
      if (item) { const cb = item.querySelector("input"); if (cb) cb.checked = m.selected; }
    });
    _updateCatHeader(ci);
  });
  updateSelCount();
});

syncApplySmart.addEventListener("click", async () => {
  const selected = getSelectedForApply();
  if (!selected.length) { setStatus(syncStatus,"⚠️ Sélectionnez au moins un message.","warning"); return; }
  const total = selected.reduce((s,c) => s + c.messages.length, 0);
  const mode  = syncApplySmart.dataset.mode || "imap";

  if (mode === "graph") {
    const ok = await confirm({
      icon: "🔗", title: "Appliquer via Microsoft Graph",
      html: `
        <div class="highlight"><strong>${total} message(s)</strong> vont recevoir leur catégorie directement dans Outlook via l'API Graph.</div>
        <ul>${selected.map(c => `<li><strong>${c.messages.length} msg</strong> → <em>${esc(c.olCategory)}</em></li>`).join("")}</ul>`,
      confirmLabel: `🔗 Appliquer via Graph (${total})`,
      confirmClass: "btn-orange",
    });
    if (!ok) return;
    showSyncStep(4);
    syncApplyBar.style.width = "0%";
    syncApplyCount.textContent = "Application via Graph en cours…";
    await send({ action:"applyCategoriesViaGraph", categories: selected });
  } else {
    const ok = await confirm({
      icon: "✓", title: "Appliquer les étiquettes via IMAP",
      html: `
        <div class="highlight"><strong>${total} message(s)</strong> vont recevoir leur étiquette via IMAP.</div>
        <ul>${selected.map(c => `<li><strong>${c.messages.length} msg</strong> → <em>${esc(c.olCategory)}</em></li>`).join("")}</ul>`,
      confirmLabel: `✓ Appliquer IMAP (${total})`,
    });
    if (!ok) return;
    showSyncStep(4);
    await send({ action:"applyCategories", categories: selected });
  }
});

syncReset.addEventListener("click", () => {
  _syncResult = null; hideStatus(syncStatus); showSyncStep(1);
});

document.getElementById("sync-apply-cancel").addEventListener("click", async () => {
  await send({ action:"cancelOp" });
  showSyncStep(3);
  setStatus(syncStatus, "⚠️ Annulation demandée — résultat partiel ci-dessus.", "warning");
});

// ═══════════════════════════════════════════════════════════════
// ONGLET TAGS
// ═══════════════════════════════════════════════════════════════
const tagList   = document.getElementById("tag-list");
const tagAddBtn = document.getElementById("tag-add-btn");
const tagRefresh = document.getElementById("tag-refresh");
const tagForm   = document.getElementById("tag-form");
const tagFTitle = document.getElementById("tag-form-title");
const tagName   = document.getElementById("tag-name");
const tagColor  = document.getElementById("tag-color");
const tagSave   = document.getElementById("tag-save");
const tagCancel = document.getElementById("tag-cancel");
const tagStatus = document.getElementById("tag-status");
let _editKey = null;

async function loadTags() {
  tagList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--text-3)"><span class="spin"></span> Chargement…</div>';
  hideStatus(tagStatus);
  try {
    const tags = await send({ action:"listTags" });
    if (!tags?.length) {
      tagList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--text-3)">Aucune étiquette définie.</div>';
      return;
    }
    tagList.innerHTML = tags.map(t => `
      <div class="tag-item">
        <div class="tag-swatch" style="background:${esc(t.color||'#888')}"></div>
        <span class="tag-name">${esc(t.tag)}</span>
        <span class="tag-key">${esc(t.key)}</span>
        <div class="tag-acts">
          <button class="btn btn-ghost btn-sm" onclick="openTagEdit('${encodeURIComponent(t.key)}','${encodeURIComponent(t.tag)}','${esc(t.color||'#4caf50')}')" title="Modifier">✏️</button>
          <button class="btn btn-danger btn-sm" onclick="deleteTag('${encodeURIComponent(t.key)}','${encodeURIComponent(t.tag)}')" title="Supprimer">🗑</button>
        </div>
      </div>`).join("");
  } catch(e) {
    tagList.innerHTML = `<div style="padding:12px;color:var(--error)">${esc(e.message)}</div>`;
  }
}

tagRefresh.addEventListener("click", loadTags);
tagAddBtn.addEventListener("click", () => {
  _editKey = null; tagFTitle.textContent = "Nouvelle étiquette";
  tagName.value = ""; tagColor.value = "#4caf50";
  tagForm.classList.add("show"); tagName.focus();
});
tagCancel.addEventListener("click", () => { tagForm.classList.remove("show"); _editKey = null; });
tagSave.addEventListener("click", async () => {
  const name = tagName.value.trim(), color = tagColor.value;
  if (!name) { setStatus(tagStatus,"⚠️ Nom requis.","warning"); return; }
  tagSave.disabled = true;
  try {
    if (_editKey) {
      const r = await send({ action:"renameTag", key:_editKey, name, color });
      if (r?.error) throw new Error(r.error);
      setStatus(tagStatus, "✅ Étiquette renommée.", "success");
    } else {
      const r = await send({ action:"createTag", name, color });
      if (r?.error) throw new Error(r.error);
      setStatus(tagStatus, `✅ Étiquette "${name}" créée.`, "success");
    }
    tagForm.classList.remove("show"); _editKey = null;
    await loadTags();
  } catch(e) { setStatus(tagStatus, "❌ " + e.message, "error"); }
  tagSave.disabled = false;
});

window.openTagEdit = function(encKey, encName, color) {
  _editKey = decodeURIComponent(encKey);
  const name = decodeURIComponent(encName);
  tagFTitle.textContent = `Modifier « ${name} »`;
  tagName.value = name; tagColor.value = color;
  tagForm.classList.add("show"); tagName.focus();
};

window.deleteTag = async function(encKey, encName) {
  const key = decodeURIComponent(encKey), name = decodeURIComponent(encName);
  const ok = await confirm({
    icon: "🗑️", title: "Supprimer l'étiquette",
    html: `<div class="highlight">L'étiquette <strong>${esc(name)}</strong> sera supprimée de Thunderbird.</div>`,
    confirmLabel: "🗑️ Supprimer", confirmClass: "btn-danger",
  });
  if (!ok) return;
  try {
    const r = await send({ action:"deleteTag", key });
    if (r?.error) throw new Error(r.error);
    setStatus(tagStatus, `✅ « ${name} » supprimée.`, "success");
    await loadTags();
  } catch(e) { setStatus(tagStatus, "❌ " + e.message, "error"); }
};

// ─── Restauration état ────────────────────────────────────────
async function restoreState() {
  const state = await send({ action:"getState" });
  if (!state) return;

  const SYNC_TYPES = ["SYNC_APPLY_DONE","GRAPH_APPLY_DONE","SYNC_ANALYSE_DONE",
                      "SYNC_APPLY_PROGRESS","GRAPH_APPLY_PROGRESS","SYNC_ERROR","GRAPH_ERROR"];
  if (!SYNC_TYPES.includes(state.type)) return;

  tabs.forEach(t   => t.classList.remove("active"));
  panels.forEach(p => p.classList.remove("active"));
  document.querySelector("[data-tab='sync']").classList.add("active");
  document.getElementById("panel-sync").classList.add("active");
  loadSyncAccounts();

  if (state.type === "SYNC_APPLY_DONE" || state.type === "GRAPH_APPLY_DONE") {
    showSyncStep(3);
    const skip = state.skipped ? ` · ${state.skipped} non trouvés` : "";
    const errs = state.errors?.length ? ` · ${state.errors.length} erreur(s)` : "";
    setStatus(syncStatus, `✅ ${state.done||0}/${state.total||0} appliqués${skip}${errs}`,
      state.errors?.length ? "warning" : "success");
    await send({ action:"clearState" });
  } else if (state.type === "SYNC_ANALYSE_DONE") {
    _syncResult = state; renderSyncResults(state); showSyncStep(3);
  } else if (state.type === "SYNC_APPLY_PROGRESS" || state.type === "GRAPH_APPLY_PROGRESS") {
    showSyncStep(4);
    const pct = state.total > 0 ? Math.round((state.done / state.total) * 100) : 0;
    syncApplyBar.style.width = pct + "%";
    syncApplyCount.textContent = `${state.done} / ${state.total} (reprise)`;
  } else if (state.type === "SYNC_ERROR" || state.type === "GRAPH_ERROR") {
    setStatus(syncStatus, "❌ " + state.error, "error");
    showSyncStep(1);
    await send({ action:"clearState" });
  }
}

restoreState();
loadGraphState();
