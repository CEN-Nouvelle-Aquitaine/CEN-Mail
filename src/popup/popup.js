/**
 * Mail-CEN popup.js v5.3
 */
"use strict";

// ─── Onglets ───────────────────────────────────────────────────
const tabs   = [...document.querySelectorAll(".tab")];
const panels = [...document.querySelectorAll(".panel")];

tabs.forEach(tab => {
  tab.addEventListener("click", () => {
    tabs.forEach(t   => t.classList.remove("active"));
    panels.forEach(p => p.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById(`panel-${tab.dataset.tab}`).classList.add("active");
    switch (tab.dataset.tab) {
      case "m365":      break;
      case "labels":    loadLabels(); break;
      case "migration": if (!_foldersLoaded) loadFolders(); break;
      case "sync":      loadSyncAccounts(); break;
      case "export":    loadExport(); break;
      case "tags":      loadTags(); break;
    }
  });
});

// ─── Helpers UI ────────────────────────────────────────────────
const send = msg => messenger.runtime.sendMessage(msg);

function setStatus(el, html, type="info") {
  el.innerHTML = html; el.className = `status show ${type}`;
}
function hideStatus(el) { el.className = "status"; }

function setProgress(bar, countEl, pctEl, wrap, done, total) {
  wrap.classList.add("show");
  const pct = total > 0 ? Math.min(100, Math.round((done/total)*100)) : 0;
  bar.style.width = pct + "%";
  countEl.textContent = `${done} / ${total}`;
  pctEl.textContent   = pct + " %";
}

function esc(s) {
  return String(s)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;")
    .replace(/'/g,"&#39;");
}

function triggerDownload(filename, content, mime="text/plain") {
  const blob = new Blob([content], { type: mime });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ─── Listener unique broadcasts background ────────────────────
messenger.runtime.onMessage.addListener(msg => {
  switch (msg.type) {
    // Migration
    case "MIG_PROGRESS":
      setProgress(migBar, migCount, migPct, migProg, msg.done, msg.total);
      break;
    case "MIG_DONE":    onMigDone(msg);        break;
    case "MIG_ERROR":   onMigError(msg.error); break;
    case "M365_ACCOUNT_DETECTED": break; // page statique, pas de handler

    // Graph auth
    case "GRAPH_AUTH_OK":
      updateGraphAuthStatus(true);
      graphAuth.disabled = false;
      graphAuth.innerHTML = "🔑 Se connecter avec Microsoft";
      setStatus(graphStatus,
        `✅ Connecte ! Token valide ${Math.round(msg.expires_in/60)} minutes.`,
        "success");
      break;
    case "GRAPH_AUTH_ERROR":
      updateGraphAuthStatus(false);
      graphAuth.disabled = false;
      graphAuth.innerHTML = "🔑 Se connecter avec Microsoft";
      setStatus(graphStatus, "❌ " + msg.error, "error");
      break;

    // Sync analyse
    case "SYNC_PROGRESS": {
      syncPhase.textContent = msg.label ?? "";
      if (msg.phase === "scan_src")       syncBar.style.width = "20%";
      else if (msg.phase === "scan_src_done") { syncBar.style.width = "40%"; syncCount.textContent = msg.label; }
      else if (msg.phase === "scan_dst")  syncBar.style.width = "60%";
      else if (msg.phase === "scan_dst_done") { syncBar.style.width = "80%"; syncCount.textContent = msg.label; }
      else if (msg.phase === "analyse_done")  syncBar.style.width = "100%";
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

    // Graph
    case "GRAPH_APPLY_PROGRESS": {
      const gpct = msg.total > 0 ? Math.round((msg.done/msg.total)*100) : 0;
      syncApplyBar.style.width = gpct + "%";
      syncApplyCount.textContent = `${msg.done} / ${msg.total}`;
      syncApplyPct.textContent   = gpct + " %";
      break;
    }
    case "GRAPH_APPLY_DONE": {
      showSyncStep(3);
      const skip = msg.skipped ? ` · ${msg.skipped} non trouves` : "";
      const errs = msg.errors?.length ? ` · ${msg.errors.length} erreur(s)` : "";
      setStatus(syncStatus,
        `✅ ${msg.done}/${msg.total} categories appliquees dans Outlook${skip}${errs}`,
        msg.errors?.length ? "warning" : "success");
      break;
    }
    case "GRAPH_ERROR":
      showSyncStep(3);
      setStatus(syncStatus, "❌ Graph : " + msg.error, "error");
      break;
  }
});

// ═══════════════════════════════════════════════════════════════
// ONGLET M365 — page statique, plus de JS nécessaire
// ═══════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════
// MODALE DE CONFIRMATION
// ═══════════════════════════════════════════════════════════════
const modalOverlay = document.getElementById("modal-overlay");
const modalTitle   = document.getElementById("modal-title");
const modalIcon    = document.getElementById("modal-icon");
const modalBody    = document.getElementById("modal-body");
const modalConfirm = document.getElementById("modal-confirm");
const modalCancel  = document.getElementById("modal-cancel");

function confirm(opts) {
  // opts = { icon, title, html, confirmLabel, confirmClass, onConfirm }
  return new Promise(resolve => {
    modalIcon.textContent    = opts.icon ?? "⚠️";
    modalTitle.textContent   = opts.title ?? "Confirmer";
    modalBody.innerHTML      = opts.html ?? "";
    modalConfirm.textContent = opts.confirmLabel ?? "✓ Confirmer";
    modalConfirm.className   = `btn ${opts.confirmClass ?? "btn-primary"}`;

    modalOverlay.classList.add("show");

    const onConfirm = () => {
      modalOverlay.classList.remove("show");
      cleanup();
      resolve(true);
    };
    const onCancel = () => {
      modalOverlay.classList.remove("show");
      cleanup();
      resolve(false);
    };
    const cleanup = () => {
      modalConfirm.removeEventListener("click", onConfirm);
      modalCancel.removeEventListener("click",  onCancel);
    };
    modalConfirm.addEventListener("click", onConfirm);
    modalCancel.addEventListener("click",  onCancel);
  });
}

// ═══════════════════════════════════════════════════════════════
// ONGLET 2 : ÉTIQUETTES / MAPPING MANUEL
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
      <input type="text"
        class="ol-input"
        data-key="${esc(t.key)}"
        value="${esc(t.olCategory && t.olCategory !== '__skip__' ? t.olCategory : '')}"
        placeholder="Nom exact dans Outlook…"
        style="flex:1;min-width:0;padding:4px 8px;font-size:12px">
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
    icon         : "💾",
    title        : "Sauvegarder le mapping",
    html         : `
      <div class="highlight">
        ${mapped.length} étiquette(s) mappée(s) vers des catégories Outlook.
      </div>
      <ul>
        ${_labelsData.map(t => {
          const cat = mapping[t.key];
          const label = cat && cat !== "__skip__"
            ? `→ <strong>${esc(cat)}</strong>`
            : `<span style="color:var(--text-3)">— ignorée</span>`;
          return `<li>${esc(t.tag)} ${label}</li>`;
        }).join("")}
      </ul>
      <p style="margin-top:10px;font-size:11.5px;color:var(--text-3)">
        Ce mapping sera appliqué automatiquement lors des migrations suivantes.
      </p>`,
    confirmLabel : "💾 Sauvegarder",
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

// ── Auto-création catégories Outlook via Graph ──────────────
const lblAutoCreate = document.getElementById("lbl-auto-create");
if (lblAutoCreate) {
  lblAutoCreate.addEventListener("click", async () => {
    // Vérifier authentification Graph
    const authState = await send({ action:"graphIsAuthenticated" });
    if (!authState.authenticated) {
      const goGraph = await confirm({
        icon : "🔗",
        title: "Connexion Graph requise",
        html : `<div class="highlight">
          Pour créer automatiquement les catégories dans Outlook, vous devez
          d'abord vous connecter dans l'onglet <strong>Graph</strong>.
        </div>`,
        confirmLabel: "Aller dans l'onglet Graph",
      });
      if (goGraph) {
        tabs.forEach(t => t.classList.remove("active"));
        panels.forEach(p => p.classList.remove("active"));
        document.querySelector("[data-tab='graph']").classList.add("active");
        document.getElementById("panel-graph").classList.add("active");
      }
      return;
    }

    // Charger les tags TB pour le récap
    const tags = await send({ action:"listTags" });
    if (!tags?.length) {
      setStatus(lblStatus, "Aucune etiquette Thunderbird a exporter.", "warning");
      return;
    }

    const ok = await confirm({
      icon : "🔗",
      title: "Creer les categories Outlook",
      html : `
        <div class="highlight">
          <strong>${tags.length} etiquette(s)</strong> Thunderbird seront creees
          comme categories dans Outlook (celles qui existent deja seront ignorees).
        </div>
        <ul>
          ${tags.map(t => `<li><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:${esc(t.color||'#888')};margin-right:5px"></span>${esc(t.tag)}</li>`).join("")}
        </ul>
        <p style="margin-top:10px;font-size:11.5px;color:var(--text-3)">
          Le mapping sera automatiquement rempli et sauvegarde.
        </p>`,
      confirmLabel: "🔗 Creer dans Outlook",
      confirmClass: "btn-orange",
    });
    if (!ok) return;

    lblAutoCreate.disabled = true;
    lblAutoCreate.innerHTML = '<span class="spin"></span> Creation en cours…';

    try {
      const r = await send({ action:"autoCreateCategories" });
      if (r?.error) throw new Error(r.error);

      const createdList = r.created.length
        ? `${r.created.length} categorie(s) creee(s)`
        : "Aucune nouvelle categorie";
      const skippedList = r.skipped.length
        ? ` · ${r.skipped.length} existante(s)`
        : "";
      setStatus(lblStatus, `✅ ${createdList}${skippedList}. Mapping sauvegarde.`, "success");

      // Recharger les labels avec le nouveau mapping
      _labelsData = await send({ action:"getTbTagsWithMapping" });
      renderLabels();
    } catch(e) {
      setStatus(lblStatus, "❌ " + e.message, "error");
    }

    lblAutoCreate.disabled = false;
    lblAutoCreate.innerHTML = "🔗 Creer automatiquement dans Outlook";
  });
}

// ═══════════════════════════════════════════════════════════════
// ONGLET 3 : MIGRATION
// ═══════════════════════════════════════════════════════════════
const migSrc     = document.getElementById("mig-src");
const migDst     = document.getElementById("mig-dst");
const migStart   = document.getElementById("mig-start");
const migCancel  = document.getElementById("mig-cancel");
const migRefresh = document.getElementById("mig-refresh");
const migProg    = document.getElementById("mig-prog");
const migBar     = document.getElementById("mig-bar");
const migCount   = document.getElementById("mig-count");
const migPct     = document.getElementById("mig-pct");
const migStatus  = document.getElementById("mig-status");
const migErrLog  = document.getElementById("mig-errlog");

let _foldersLoaded = false;

function folderIcon(type) {
  const m = { inbox:"📥", trash:"🗑️", drafts:"📝", sent:"📤", archives:"📦", junk:"🚫" };
  return m[type] ?? "📁";
}

async function loadFolders() {
  migSrc.innerHTML = "<option value=''>Chargement…</option>";
  migDst.innerHTML = "<option value=''>Chargement…</option>";
  try {
    const tree = await send({ action:"getFolderTree" });
    _foldersLoaded = true;
    const byAccount = {};
    for (const f of tree) {
      if (!byAccount[f.accountId])
        byAccount[f.accountId] = { name: f.accountName, folders: [] };
      byAccount[f.accountId].folders.push(f);
    }
    const build = sel => {
      sel.innerHTML = "<option value=''>— Sélectionner un dossier —</option>";
      for (const [, acc] of Object.entries(byAccount)) {
        const grp = document.createElement("optgroup");
        grp.label = "📭 " + acc.name;
        for (const f of acc.folders) {
          const opt = document.createElement("option");
          opt.value = f.id;
          opt.textContent = "\u00a0".repeat(f.depth*3) + folderIcon(f.type) + " " + f.name;
          grp.appendChild(opt);
        }
        sel.appendChild(grp);
      }
    };
    build(migSrc); build(migDst);
    hideStatus(migStatus);
  } catch(e) {
    migSrc.innerHTML = "<option value=''>Erreur</option>";
    setStatus(migStatus, "❌ " + e.message, "error");
  }
}

document.querySelector("[data-tab='migration']").addEventListener("click", () => {
  if (!_foldersLoaded) loadFolders();
});

migRefresh.addEventListener("click", () => {
  _foldersLoaded = false; loadFolders();
});

migStart.addEventListener("click", async () => {
  const src  = migSrc.value;
  const dst  = migDst.value;
  const mode = document.querySelector("input[name='mig-mode']:checked").value;
  const recursive = document.getElementById("mig-recursive").checked;

  if (!src) { setStatus(migStatus,"⚠️ Sélectionnez un dossier source.","warning"); return; }
  if (!dst) { setStatus(migStatus,"⚠️ Sélectionnez un dossier destination.","warning"); return; }
  if (src===dst) { setStatus(migStatus,"⚠️ Source et destination identiques.","warning"); return; }

  const srcLabel = migSrc.options[migSrc.selectedIndex].textContent.trim();
  const dstLabel = migDst.options[migDst.selectedIndex].textContent.trim();
  const modeLabel = mode === "move" ? "déplacés (supprimés de la source)" : "copiés (source conservée)";
  const recLabel  = recursive ? "Tous les sous-dossiers seront inclus." : "Dossier seul, sans sous-dossiers.";

  const ok = await confirm({
    icon : "📦",
    title: "Confirmer la migration",
    html : `
      <div class="highlight">
        Les messages seront <strong>${modeLabel}</strong>.
      </div>
      <ul>
        <li>📂 Source : <strong>${esc(srcLabel)}</strong></li>
        <li>📁 Destination : <strong>${esc(dstLabel)}</strong></li>
        <li>${recLabel}</li>
      </ul>
      <p style="margin-top:10px;font-size:11.5px;color:var(--error)">
        ⚠️ ${mode === "move" ? "Les emails seront supprimés de la source après transfert." : "Les emails resteront dans la source."}
      </p>`,
    confirmLabel : mode === "move" ? "📦 Démarrer la migration" : "📋 Démarrer la copie",
    confirmClass : "btn-primary",
  });

  if (!ok) return;

  migStart.disabled = true; migCancel.disabled = false;
  migErrLog.classList.remove("show");
  setProgress(migBar, migCount, migPct, migProg, 0, 0);
  setStatus(migStatus, '<span class="spin"></span> Migration en cours…', "info");

  const force = document.getElementById("mig-force").checked;
  const r = await send({ action:"startMigrationTree", source:src, dest:dst, mode, force });
  if (r?.error) {
    setStatus(migStatus, "❌ " + r.error, "error");
    migStart.disabled = false; migCancel.disabled = true;
  }
});

migCancel.addEventListener("click", async () => {
  migCancel.disabled = true;
  await send({ action:"cancelMigration" });
  setStatus(migStatus, "⏹ Annulation…", "warning");
});

function onMigDone(msg) {
  migStart.disabled = false; migCancel.disabled = true;
  setProgress(migBar, migCount, migPct, migProg, msg.done, msg.total);
  const icon = msg.status==="cancelled" ? "⏹" : "✅";
  const type = msg.status==="cancelled" ? "warning" : "success";
  const errs = msg.errors?.length ? ` · ⚠️ ${msg.errors.length} erreur(s)` : "";
  const dups = msg.skipped ? ` · 🔄 ${msg.skipped} doublon(s) ignoré(s)` : "";
  setStatus(migStatus, `${icon} ${msg.done}/${msg.total} messages traités${dups}${errs}`, type);
  if (msg.errors?.length) {
    migErrLog.classList.add("show");
    migErrLog.innerHTML = msg.errors.slice(0,30)
      .map(e => `• ${esc(e.subject||"?")} — ${esc(e.reason)}`).join("<br>");
    if (msg.errors.length>30) migErrLog.innerHTML += `<br><em>… et ${msg.errors.length-30} autre(s)</em>`;
  }
  send({ action:"clearMigState" });
}

function onMigError(err) {
  migStart.disabled = false; migCancel.disabled = true;
  setStatus(migStatus, "❌ " + err, "error");
}

// ═══════════════════════════════════════════════════════════════
// ONGLET 4 : SYNCHRONISATION — flux en 4 étapes
// ═══════════════════════════════════════════════════════════════
const syncSrc      = document.getElementById("sync-src");
const syncDst      = document.getElementById("sync-dst");
const syncAnalyse  = document.getElementById("sync-analyse");
const syncStep1    = document.getElementById("sync-step1");
const syncStep2    = document.getElementById("sync-step2");
const syncStep3    = document.getElementById("sync-step3");
const syncStep4    = document.getElementById("sync-step4");
const syncBar      = document.getElementById("sync-bar");
const syncCount    = document.getElementById("sync-count");
const syncPct      = document.getElementById("sync-pct");
const syncPhase    = document.getElementById("sync-phase");
const syncSummary  = document.getElementById("sync-summary");
const syncCatList  = document.getElementById("sync-cat-list");
const syncCheckAll = document.getElementById("sync-check-all");
const syncUnmapped = document.getElementById("sync-unmapped");
const syncUnmappedMsg = document.getElementById("sync-unmapped-msg");
const syncApply    = document.getElementById("sync-apply");
const syncApplyBar = document.getElementById("sync-apply-bar");
const syncApplyCount = document.getElementById("sync-apply-count");
const syncApplyPct = document.getElementById("sync-apply-pct");
const syncReset    = document.getElementById("sync-reset");
const syncStatus   = document.getElementById("sync-status");

let _syncResult = null; // résultat de l'analyse

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

// Étape 1 → 2 : Analyser
syncAnalyse.addEventListener("click", async () => {
  const src = syncSrc.value;
  const dst = syncDst.value;
  if (!src) { setStatus(syncStatus,"⚠️ Sélectionnez un compte source.","warning"); return; }
  if (!dst) { setStatus(syncStatus,"⚠️ Sélectionnez un compte destination.","warning"); return; }
  if (src===dst) { setStatus(syncStatus,"⚠️ Source et destination identiques.","warning"); return; }

  const srcLabel = syncSrc.options[syncSrc.selectedIndex].textContent.trim();
  const dstLabel = syncDst.options[syncDst.selectedIndex].textContent.trim();
  const mapping  = await send({ action:"loadMapping" });
  const mappedCount = Object.values(mapping||{}).filter(v => v && v !== "__skip__").length;

  const ok = await confirm({
    icon : "🔍",
    title: "Analyser les deux boîtes",
    html : `
      <div class="highlight">
        L'analyse va comparer les deux boîtes et identifier les emails
        dont les catégories Outlook sont à appliquer.<br>
        <strong>Aucune modification ne sera faite à cette étape.</strong>
      </div>
      <ul>
        <li>📬 Source : <strong>${esc(srcLabel)}</strong></li>
        <li>📭 Destination : <strong>${esc(dstLabel)}</strong></li>
        <li>${mappedCount > 0
          ? `✅ ${mappedCount} étiquette(s) mappée(s) vers des catégories Outlook`
          : `⚠️ Aucun mapping configuré — configurez l'onglet Étiquettes d'abord`}</li>
      </ul>`,
    confirmLabel: "🔍 Lancer l'analyse",
  });
  if (!ok) return;

  hideStatus(syncStatus);
  showSyncStep(2);
  syncBar.style.width = "10%";
  syncCount.textContent = "Initialisation…";

  const r = await send({ action:"analyseBoxes", srcAccountId:src, dstAccountId:dst });
  if (r?.error) {
    setStatus(syncStatus, "❌ " + r.error, "error");
    showSyncStep(1);
  }
});

// (Sync broadcasts handled in unified listener above)

function renderSyncResults(r) {
  // Résumé global
  const toApply = r.categories.reduce((s,c) => s + c.messages.length, 0);
  syncSummary.innerHTML = `
    <div class="sync-sum-card">
      <div class="sync-sum-num">${r.srcTotal}</div>
      <div class="sync-sum-label">avec étiquettes (source)</div>
    </div>
    <div class="sync-sum-card">
      <div class="sync-sum-num">${r.dstTotal}</div>
      <div class="sync-sum-label">indexés (destination)</div>
    </div>
    <div class="sync-sum-card">
      <div class="sync-sum-num" style="color:var(--accent)">${toApply}</div>
      <div class="sync-sum-label">catégories à appliquer</div>
    </div>
    ${r.notFoundTotal ? `
    <div class="sync-sum-card">
      <div class="sync-sum-num" style="color:var(--warning)">${r.notFoundTotal}</div>
      <div class="sync-sum-label">non trouvés côté Outlook</div>
    </div>` : ""}
  `;

  if (r.noMapping.length) {
    syncUnmapped.style.display = "block";
    syncUnmappedMsg.innerHTML = `⚠️ ${r.noMapping.length} étiquette(s) sans mapping — configurez l'onglet <strong>Étiquettes</strong>.`;
  } else {
    syncUnmapped.style.display = "none";
  }

  if (!r.categories.length) {
    syncCatList.innerHTML = '<div style="padding:16px;text-align:center;color:var(--text-3)">Aucune catégorie à appliquer — tout est déjà à jour.</div>';
    syncApply.style.display = "none";
    document.getElementById("sync-apply-graph").style.display = "none";
    return;
  }

  // Stocker les données enrichies pour le filtrage/tri
  _syncCategories = r.categories.map(cat => ({
    ...cat,
    // Enrichir chaque message avec un flag de sélection
    messages: cat.messages.map(m => ({ ...m, selected: true }))
  }));

  renderDetailedList();
  updateSelCount();

  syncApply.style.display = "";
  document.getElementById("sync-apply-graph").style.display = "";
}

// État des filtres
let _syncFilter = { text: "", from: null, to: null, sort: "date-desc" };
let _syncCategories = [];
let _accOpen = {}; // catIndex → bool (accordéon ouvert)

function applyFilters(messages) {
  return messages.filter(m => {
    if (_syncFilter.text) {
      const q = _syncFilter.text.toLowerCase();
      if (!((m.subject||"").toLowerCase().includes(q) ||
            (m.sender||"").toLowerCase().includes(q))) return false;
    }
    if (_syncFilter.from && m.date && new Date(m.date) < new Date(_syncFilter.from)) return false;
    if (_syncFilter.to   && m.date && new Date(m.date) > new Date(_syncFilter.to + "T23:59:59")) return false;
    return true;
  }).sort((a, b) => {
    switch (_syncFilter.sort) {
      case "date-asc":   return new Date(a.date||0) - new Date(b.date||0);
      case "date-desc":  return new Date(b.date||0) - new Date(a.date||0);
      case "subject":    return (a.subject||"").localeCompare(b.subject||"");
      case "sender":     return (a.sender||"").localeCompare(b.sender||"");
      default:           return 0;
    }
  });
}

function renderDetailedList() {
  if (!_syncCategories.length) return;

  let html = "";
  _syncCategories.forEach((cat, ci) => {
    const filtered   = applyFilters(cat.messages);
    const selCount   = filtered.filter(m => m.selected).length;
    const isOpen     = _accOpen[ci] !== false; // ouvert par défaut

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
        <button class="btn btn-ghost btn-sm" style="padding:2px 7px;font-size:10.5px"
          onclick="setCatSel(${ci},true)">Tout ☑</button>
        <button class="btn btn-ghost btn-sm" style="padding:2px 7px;font-size:10.5px"
          onclick="setCatSel(${ci},false)">Tout ☐</button>
        <button class="btn btn-ghost btn-sm" style="padding:2px 7px;font-size:10.5px"
          onclick="invertCatSel(${ci})">Inverser</button>
      </div>
      <div class="sync-msg-list ${isOpen ? "open" : ""}" id="sync-msg-list-${ci}">
        ${filtered.length === 0
          ? `<div style="padding:10px 28px;font-size:11.5px;color:var(--text-3)">Aucun message ne correspond aux filtres.</div>`
          : filtered.map((m, mi) => {
              const origIdx = cat.messages.indexOf(m);
              const dateStr = m.date ? new Date(m.date).toLocaleDateString("fr-FR") : "—";
              return `
                <div class="sync-msg-item" onclick="toggleMsg(${ci},${origIdx})">
                  <input type="checkbox" ${m.selected ? "checked" : ""}
                    onclick="event.stopPropagation();toggleMsg(${ci},${origIdx})">
                  <div class="sync-msg-body">
                    <div class="sync-msg-subject">${esc(m.subject || "(sans sujet)")}</div>
                    <div class="sync-msg-meta">
                      <span class="sync-msg-sender">👤 ${esc(m.sender || "—")}</span>
                      <span class="sync-msg-date">📅 ${dateStr}</span>
                    </div>
                  </div>
                </div>`;
            }).join("")}
      </div>`;
  });

  syncCatList.innerHTML = html;

  // Mettre à jour l'état indeterminate des checkboxes
  syncCatList.querySelectorAll(".cat-master-check[indeterminate-js]").forEach(cb => {
    cb.indeterminate = true;
  });

  // Clic sur header → toggle accordéon
  syncCatList.querySelectorAll(".sync-acc-header").forEach(hdr => {
    hdr.addEventListener("click", (e) => {
      if (e.target.type === "checkbox") return;
      const ci = +hdr.dataset.ci;
      _accOpen[ci] = !(_accOpen[ci] !== false);
      const toggle = hdr.querySelector(".sync-acc-toggle");
      const list   = document.getElementById(`sync-msg-list-${ci}`);
      toggle.classList.toggle("open", _accOpen[ci]);
      list.classList.toggle("open", _accOpen[ci]);
    });
  });
}

// Fonctions de sélection exposées globalement
window.toggleMsg = function(ci, mi) {
  _syncCategories[ci].messages[mi].selected = !_syncCategories[ci].messages[mi].selected;
  renderDetailedList();
  updateSelCount();
};

window.toggleCatAll = function(ci, checked) {
  applyFilters(_syncCategories[ci].messages).forEach(m => m.selected = checked);
  renderDetailedList();
  updateSelCount();
};

window.setCatSel = function(ci, val) {
  applyFilters(_syncCategories[ci].messages).forEach(m => m.selected = val);
  renderDetailedList();
  updateSelCount();
};

window.invertCatSel = function(ci) {
  applyFilters(_syncCategories[ci].messages).forEach(m => m.selected = !m.selected);
  renderDetailedList();
  updateSelCount();
};

function updateSelCount() {
  const total = _syncCategories.reduce((s,c) => s + applyFilters(c.messages).length, 0);
  const sel   = _syncCategories.reduce((s,c) => s + applyFilters(c.messages).filter(m => m.selected).length, 0);
  document.getElementById("sync-sel-count").textContent = `${sel} sélectionné(s) sur ${total} message(s)`;
}

function getSelectedForApply() {
  return _syncCategories
    .map(cat => ({
      ...cat,
      messages: cat.messages.filter(m => m.selected)
    }))
    .filter(cat => cat.messages.length > 0);
}

// Filtres
document.getElementById("sync-filter-text").addEventListener("input", e => {
  _syncFilter.text = e.target.value;
  renderDetailedList();
  updateSelCount();
});
document.getElementById("sync-filter-from").addEventListener("change", e => {
  _syncFilter.from = e.target.value || null;
  renderDetailedList(); updateSelCount();
});
document.getElementById("sync-filter-to").addEventListener("change", e => {
  _syncFilter.to = e.target.value || null;
  renderDetailedList(); updateSelCount();
});
document.getElementById("sync-filter-clear").addEventListener("click", () => {
  _syncFilter = { text:"", from:null, to:null, sort:_syncFilter.sort };
  document.getElementById("sync-filter-text").value = "";
  document.getElementById("sync-filter-from").value = "";
  document.getElementById("sync-filter-to").value   = "";
  renderDetailedList(); updateSelCount();
});
document.getElementById("sync-sort").addEventListener("change", e => {
  _syncFilter.sort = e.target.value;
  renderDetailedList(); updateSelCount();
});

// Sélection globale
document.getElementById("sync-sel-all").addEventListener("click", () => {
  _syncCategories.forEach(cat => applyFilters(cat.messages).forEach(m => m.selected = true));
  renderDetailedList(); updateSelCount();
});
document.getElementById("sync-sel-none").addEventListener("click", () => {
  _syncCategories.forEach(cat => applyFilters(cat.messages).forEach(m => m.selected = false));
  renderDetailedList(); updateSelCount();
});
document.getElementById("sync-sel-invert").addEventListener("click", () => {
  _syncCategories.forEach(cat => applyFilters(cat.messages).forEach(m => m.selected = !m.selected));
  renderDetailedList(); updateSelCount();
});

if (syncCheckAll) {
  syncCheckAll.addEventListener("change", () => {
    syncCatList.querySelectorAll(".cat-check").forEach(cb => {
      cb.checked = syncCheckAll.checked;
    });
  });
}

syncApply.addEventListener("click", async () => {
  const selected = getSelectedForApply();
  if (!selected.length) {
    setStatus(syncStatus, "⚠️ Sélectionnez au moins un message.", "warning");
    return;
  }
  const total = selected.reduce((s,c) => s + c.messages.length, 0);

  const ok = await confirm({
    icon : "✓",
    title: "Appliquer les étiquettes via IMAP",
    html : `
      <div class="highlight">
        <strong>${total} message(s)</strong> vont recevoir leur étiquette via IMAP.
      </div>
      <ul>
        ${selected.map(c => `<li><strong>${c.messages.length} msg</strong> → <em>${esc(c.olCategory)}</em></li>`).join("")}
      </ul>
      <p style="margin-top:10px;font-size:11.5px;color:var(--text-3)">
        Note : visible dans Thunderbird, mais peut ne pas apparaître dans Outlook Online.
        Utilisez "Appliquer via Graph" pour Outlook Online.
      </p>`,
    confirmLabel: `✓ Appliquer IMAP (${total})`,
  });
  if (!ok) return;

  showSyncStep(4);
  await send({ action:"applyCategories", categories: selected });
});

// Recommencer
syncReset.addEventListener("click", () => {
  _syncResult = null;
  hideStatus(syncStatus);
  showSyncStep(1);
});

// ═══════════════════════════════════════════════════════════════
// ONGLET 5 : EXPORT
// ═══════════════════════════════════════════════════════════════
const expBooks     = document.getElementById("exp-books");
const expContactsBtn = document.getElementById("exp-contacts-btn");
const expCalsAuto  = document.getElementById("exp-cals-auto");
const expCalsGuide = document.getElementById("exp-cals-guide");
const expStatus    = document.getElementById("exp-status");

let _addressBooks = [];

async function loadExport() {
  expBooks.innerHTML = '<div style="padding:14px;text-align:center;color:var(--text-3)"><span class="spin"></span> Chargement…</div>';
  hideStatus(expStatus);

  try {
    _addressBooks = await send({ action:"listAddressBooks" });
    if (!_addressBooks.length) {
      expBooks.innerHTML = '<div style="padding:14px;text-align:center;color:var(--text-3)">Aucun carnet d\'adresses trouvé.</div>';
    } else {
      expBooks.innerHTML = _addressBooks.map((b, i) => `
        <div class="exp-item">
          <input type="checkbox" class="book-check" data-i="${i}" checked>
          <span class="exp-icon">${b.type==="carddav" ? "☁️" : "📒"}</span>
          <span class="exp-name">${esc(b.name)}</span>
          <span class="exp-count">${b.count} contact(s)</span>
        </div>
      `).join("");
    }
  } catch(e) {
    expBooks.innerHTML = `<div style="padding:12px;color:var(--error)">Erreur : ${esc(e.message)}</div>`;
  }

  // Calendriers
  try {
    const cals = await send({ action:"detectCalendars" });
    if (cals.available && cals.calendars.length) {
      expCalsAuto.style.display  = "block";
      expCalsGuide.style.display = "none";
      document.getElementById("exp-cal-list").innerHTML = cals.calendars.map((c,i) => `
        <div class="exp-item">
          <input type="checkbox" class="cal-check" data-i="${i}" checked>
          <span class="exp-icon">📅</span>
          <span class="exp-name">${esc(c.name)}</span>
        </div>
      `).join("");
    }
  } catch {}
}

expContactsBtn.addEventListener("click", async () => {
  const selected = [...expBooks.querySelectorAll(".book-check:checked")]
    .map(cb => _addressBooks[+cb.dataset.i])
    .filter(Boolean);

  if (!selected.length) {
    setStatus(expStatus, "⚠️ Sélectionnez au moins un carnet.", "warning");
    return;
  }

  const total = selected.reduce((s, b) => s + b.count, 0);

  const ok = await confirm({
    icon : "📇",
    title: "Exporter les contacts",
    html : `
      <div class="highlight">
        ${total} contact(s) seront exportés dans un fichier <strong>.vcf</strong>.
      </div>
      <ul>
        ${selected.map(b => `<li>${b.type==="carddav"?"☁️":"📒"} ${esc(b.name)} — ${b.count} contact(s)</li>`).join("")}
      </ul>
      <p style="margin-top:10px;font-size:11.5px;color:var(--text-3)">
        Le fichier .vcf peut être importé directement dans Outlook :<br>
        Fichier → Ouvrir et exporter → Importer/Exporter → Importer un fichier .vcf
      </p>`,
    confirmLabel: "💾 Exporter",
  });

  if (!ok) return;
  expContactsBtn.disabled = true;
  expContactsBtn.innerHTML = '<span class="spin"></span> Export en cours…';
  try {
    const r = await send({ action:"exportContacts", bookIds: selected.map(b => b.id) });
    if (!r.vcf || !r.vcf.trim()) {
      setStatus(expStatus, "ℹ️ Aucun contact à exporter.", "info");
    } else {
      const date = new Date().toISOString().slice(0,10);
      triggerDownload(`contacts-cen-${date}.vcf`, r.vcf, "text/vcard");
      setStatus(expStatus, `✅ Export téléchargé.`, "success");
    }
  } catch(e) {
    setStatus(expStatus, "❌ " + e.message, "error");
  }
  expContactsBtn.disabled = false;
  expContactsBtn.innerHTML = "💾 Exporter les contacts sélectionnés (.vcf)";
});

// ═══════════════════════════════════════════════════════════════
// ONGLET 6 : TAGS
// ═══════════════════════════════════════════════════════════════
const tagList    = document.getElementById("tag-list");
const tagAddBtn  = document.getElementById("tag-add-btn");
const tagRefresh = document.getElementById("tag-refresh");
const tagForm    = document.getElementById("tag-form");
const tagFTitle  = document.getElementById("tag-form-title");
const tagName    = document.getElementById("tag-name");
const tagColor   = document.getElementById("tag-color");
const tagSave    = document.getElementById("tag-save");
const tagCancel  = document.getElementById("tag-cancel");
const tagStatus  = document.getElementById("tag-status");

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
          <button class="btn btn-ghost btn-sm"
            onclick="openTagEdit('${encodeURIComponent(t.key)}','${encodeURIComponent(t.tag)}','${esc(t.color||'#4caf50')}')"
            title="Modifier">✏️</button>
          <button class="btn btn-danger btn-sm"
            onclick="deleteTag('${encodeURIComponent(t.key)}','${encodeURIComponent(t.tag)}')"
            title="Supprimer">🗑</button>
        </div>
      </div>`).join("");
  } catch(e) {
    tagList.innerHTML = `<div style="padding:12px;color:var(--error)">${esc(e.message)}</div>`;
  }
}

tagRefresh.addEventListener("click", loadTags);

tagAddBtn.addEventListener("click", () => {
  _editKey = null;
  tagFTitle.textContent = "Nouvelle étiquette";
  tagName.value  = "";
  tagColor.value = "#4caf50";
  tagForm.classList.add("show");
  tagName.focus();
});

tagCancel.addEventListener("click", () => {
  tagForm.classList.remove("show"); _editKey = null;
});

tagSave.addEventListener("click", async () => {
  const name  = tagName.value.trim();
  const color = tagColor.value;
  if (!name) { setStatus(tagStatus,"⚠️ Nom requis.","warning"); return; }
  tagSave.disabled = true;
  try {
    if (_editKey) {
      const r = await send({ action:"renameTag", key:_editKey, name, color });
      if (r?.error) throw new Error(r.error);
      setStatus(tagStatus, `✅ Étiquette renommée.`, "success");
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
  tagName.value  = name;
  tagColor.value = color;
  tagForm.classList.add("show");
  tagName.focus();
};

window.deleteTag = async function(encKey, encName) {
  const key  = decodeURIComponent(encKey);
  const name = decodeURIComponent(encName);

  const ok = await confirm({
    icon : "🗑️",
    title: `Supprimer l'étiquette`,
    html : `
      <div class="highlight">
        L'étiquette <strong>${esc(name)}</strong> sera supprimée de Thunderbird.
      </div>
      <p style="font-size:11.5px;color:var(--text-3)">
        Les emails qui portaient cette étiquette la conserveront dans leur
        métadonnées IMAP, mais elle n'apparaîtra plus dans l'interface TB.
      </p>`,
    confirmLabel : "🗑️ Supprimer",
    confirmClass : "btn-danger",
  });

  if (!ok) return;
  try {
    const r = await send({ action:"deleteTag", key });
    if (r?.error) throw new Error(r.error);
    setStatus(tagStatus, `✅ « ${name} » supprimée.`, "success");
    await loadTags();
  } catch(e) { setStatus(tagStatus, "❌ " + e.message, "error"); }
};

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
    graphAuthSub.textContent   = "Token actif — vous pouvez utiliser 'Appliquer via Graph' dans l'onglet Synchro.";
    graphAuth.style.display       = "none";
    graphDisconnect.style.display = "block";
  } else {
    graphAuthIcon.textContent  = "🔒";
    graphAuthLabel.textContent = "Non connecté";
    graphAuthLabel.style.color = "var(--text)";
    graphAuthSub.textContent   = "Cliquez 'Se connecter' pour vous authentifier avec votre compte Microsoft CEN.";
    graphAuth.style.display       = "block";
    graphDisconnect.style.display = "none";
  }
}

document.querySelector("[data-tab='graph']").addEventListener("click", loadGraphState);

graphAuth.addEventListener("click", async () => {
  graphAuth.disabled = true;
  graphAuth.innerHTML = '<span class="spin"></span> Connexion…';
  setStatus(graphStatus, "La fenetre Microsoft va s'ouvrir. Connectez-vous puis revenez ici.", "info");

  try {
    const r = await send({ action:"graphAuthenticate" });
    if (r?.error) throw new Error(r.error);
    // Le background gère via broadcast GRAPH_AUTH_OK / GRAPH_AUTH_ERROR
  } catch(e) {
    // Le popup peut se fermer pendant l'auth — normal
  }
});

graphDisconnect.addEventListener("click", () => {
  updateGraphAuthStatus(false);
  hideStatus(graphStatus);
});

// ── Bouton "Appliquer via Graph" dans la synchro ──────────────
document.getElementById("sync-apply-graph").addEventListener("click", async () => {
  // Vérifier authentification Graph
  const authState = await send({ action:"graphIsAuthenticated" });
  if (!authState.authenticated) {
    const ok = await confirm({
      icon : "🔗",
      title: "Authentification Graph requise",
      html : `
        <div class="highlight">
          Pour appliquer les catégories via Microsoft Graph, vous devez
          d'abord vous authentifier dans l'onglet <strong>🔗 Graph</strong>.
        </div>
        <p style="font-size:11.5px;color:var(--text-3);margin-top:8px">
          Allez dans l'onglet Graph, saisissez votre secret et authentifiez-vous.
        </p>`,
      confirmLabel: "Aller dans l'onglet Graph",
    });
    if (ok) {
      tabs.forEach(t => t.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));
      document.querySelector("[data-tab='graph']").classList.add("active");
      document.getElementById("panel-graph").classList.add("active");
      loadGraphState();
    }
    return;
  }

  const selected = getSelectedForApply();

  if (!selected.length) {
    setStatus(syncStatus, "⚠️ Sélectionnez au moins un message.", "warning");
    return;
  }

  const total = selected.reduce((s,c) => s + c.messages.length, 0);

  const ok = await confirm({
    icon : "🔗",
    title: "Appliquer les catégories via Microsoft Graph",
    html : `
      <div class="highlight">
        <strong>${total} message(s)</strong> vont recevoir leur catégorie
        directement dans Outlook via l'API Microsoft Graph.
      </div>
      <ul>
        ${selected.map(c => `
          <li><strong>${c.messages.length} msg</strong> → <em>${esc(c.olCategory)}</em></li>
        `).join("")}
      </ul>
      <p style="margin-top:10px;font-size:11.5px;color:var(--text-3)">
        Les catégories seront visibles immédiatement dans Outlook Online.
      </p>`,
    confirmLabel: `🔗 Appliquer via Graph (${total} messages)`,
    confirmClass: "btn-orange",
  });
  if (!ok) return;

  showSyncStep(4);
  syncApplyBar.style.width = "0%";
  syncApplyCount.textContent = "Application via Graph en cours…";

  await send({ action:"applyCategoriesViaGraph", categories: selected });
});

// (Graph broadcasts handled in unified listener above)
function onSyncDone(msg) {
  showSyncStep(1);
  const errs = msg.errors?.length ? ` · ${msg.errors.length} erreur(s)` : "";
  setStatus(syncStatus, `✅ Synchronisation terminée — ${msg.done} message(s)${errs}`, msg.errors?.length ? "warning" : "success");
}

function onSyncError(err) {
  showSyncStep(1);
  setStatus(syncStatus, "❌ " + err, "error");
}

async function restoreState() {
  const state = await send({ action:"getMigState" });
  if (!state) return;

  // Trouver l'onglet concerné
  const isMig  = ["MIG_PROGRESS","MIG_DONE","MIG_ERROR"].includes(state.type);
  const isSync = ["SYNC_PROGRESS","SYNC_DONE","SYNC_ERROR"].includes(state.type);
  const tabKey = isMig ? "migration" : isSync ? "sync" : null;
  if (!tabKey) return;

  // Basculer sur le bon onglet
  tabs.forEach(t   => t.classList.remove("active"));
  panels.forEach(p => p.classList.remove("active"));
  document.querySelector(`[data-tab='${tabKey}']`).classList.add("active");
  document.getElementById(`panel-${tabKey}`).classList.add("active");

  if (isMig) {
    if (!_foldersLoaded) loadFolders();
    if (state.type === "MIG_PROGRESS") {
      setProgress(migBar, migCount, migPct, migProg, state.done, state.total);
      setStatus(migStatus, '<span class="spin"></span> Migration en cours (reprise)…', "info");
      migStart.disabled = true; migCancel.disabled = false;
    } else if (state.type === "MIG_DONE") {
      onMigDone(state); await send({ action:"clearMigState" });
    } else if (state.type === "MIG_ERROR") {
      onMigError(state.error); await send({ action:"clearMigState" });
    }
  } else if (isSync) {
    loadSyncAccounts();
    if (state.type === "SYNC_DONE") {
      onSyncDone(state); await send({ action:"clearMigState" });
    } else if (state.type === "SYNC_ERROR") {
      onSyncError(state.error); await send({ action:"clearMigState" });
    }
  }
}

restoreState();
loadGraphState();
