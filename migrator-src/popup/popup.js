/**
 * Mail-Migrator CEN — popup.js v1.0.0
 * Logique UI : arbre de dossiers checkboxes, progression, gestion doublons.
 */
"use strict";

// ─────────────────────────────────────────────────────────────
// HELPERS UI
// ─────────────────────────────────────────────────────────────

const $      = id => document.getElementById(id);
const escHtml = s => (s || "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

function setStatus(id, type, html) {
  const el = $(id);
  el.className = type ? `status show ${type}` : "status";
  el.innerHTML = html || "";
}

function goStep(step) {
  ["step-setup","step-progress","step-results"].forEach(id => {
    $(id).style.display = id === step ? "" : "none";
  });
  const badge = $("hdr-running");
  if (badge) badge.style.display = step === "step-progress" ? "" : "none";
}

// ─────────────────────────────────────────────────────────────
// ÉTAT
// ─────────────────────────────────────────────────────────────

let _duplicates    = [];   // liste complète des doublons reçus à la fin
let _dstFolderId   = null; // dossier destination choisi pour la migration

// ─────────────────────────────────────────────────────────────
// ARBRE DE DOSSIERS LOCAUX
// ─────────────────────────────────────────────────────────────

/**
 * Construit un nœud HTML pour un dossier local.
 * @param {Object} folder  - dossier TB (id, name, subFolders)
 * @param {number} depth   - niveau d'indentation
 */
function buildFolderNode(folder, depth) {
  const hasSubs = Array.isArray(folder.subFolders) && folder.subFolders.length > 0;
  const indent  = depth * 18;

  const node = document.createElement("div");
  node.className = "f-node";
  node.dataset.folderId = folder.id;

  // ── En-tête du nœud ──
  const header = document.createElement("div");
  header.className = "f-header";
  header.style.paddingLeft = `${8 + indent}px`;

  // Flèche expand/collapse
  const toggle = document.createElement("span");
  toggle.className = "f-toggle" + (hasSubs ? " clickable" : "");
  toggle.textContent = hasSubs ? "▶" : "";

  // Checkbox
  const cb = document.createElement("input");
  cb.type = "checkbox";
  cb.className = "f-cb";
  cb.dataset.id   = folder.id;
  cb.dataset.name = folder.name;

  // Icône + nom
  const icon = document.createElement("span");
  icon.className = "f-icon";
  icon.textContent = hasSubs ? "📂" : "📁";

  const nameEl = document.createElement("span");
  nameEl.className = "f-name";
  nameEl.textContent = folder.name;

  header.appendChild(toggle);
  header.appendChild(cb);
  header.appendChild(icon);
  header.appendChild(nameEl);
  node.appendChild(header);

  // ── Sous-dossiers ──
  if (hasSubs) {
    const children = document.createElement("div");
    children.className = "f-children";

    for (const sub of folder.subFolders) {
      children.appendChild(buildFolderNode(sub, depth + 1));
    }
    node.appendChild(children);

    // Toggle expand/collapse
    toggle.addEventListener("click", e => {
      e.stopPropagation();
      const open = children.classList.toggle("open");
      toggle.textContent = open ? "▼" : "▶";
      icon.textContent   = open ? "📂" : "📁";
    });

    // Clic sur l'en-tête (hors checkbox/toggle) → expand
    header.addEventListener("click", e => {
      if (e.target === cb || e.target === toggle) return;
      const open = children.classList.toggle("open");
      toggle.textContent = open ? "▼" : "▶";
      icon.textContent   = open ? "📂" : "📁";
    });
  }

  // ── Cascade check vers les enfants ──
  cb.addEventListener("change", () => {
    node.querySelectorAll(".f-cb").forEach(c => { c.checked = cb.checked; });
  });

  return node;
}

function renderLocalTree(localAccounts) {
  const container = $("local-tree");
  container.innerHTML = "";

  if (!localAccounts || localAccounts.length === 0) {
    container.innerHTML = `<div class="tree-empty">Aucun dossier local trouvé.</div>`;
    return;
  }

  for (const acc of localAccounts) {
    // En-tête de compte
    const accHdr = document.createElement("div");
    accHdr.className = "acc-hdr";
    accHdr.innerHTML = `<span>🖥</span> ${escHtml(acc.name)}`;
    container.appendChild(accHdr);

    if (!acc.folders || acc.folders.length === 0) {
      const empty = document.createElement("div");
      empty.className = "tree-empty";
      empty.textContent = "Aucun sous-dossier";
      container.appendChild(empty);
      continue;
    }

    for (const folder of acc.folders) {
      container.appendChild(buildFolderNode(folder, 0));
    }
  }
}

// ─────────────────────────────────────────────────────────────
// LISTE DESTINATION IMAP (OUTLOOK)
// ─────────────────────────────────────────────────────────────

function renderImapFolders(imapAccounts) {
  const select = $("dst-select");
  select.innerHTML = `<option value="">— Choisir un dossier de destination Outlook —</option>`;

  if (!imapAccounts || imapAccounts.length === 0) {
    select.innerHTML = `<option value="">Aucun compte Outlook trouvé</option>`;
    return;
  }

  function addOptions(parent, folders, depth) {
    for (const f of folders) {
      const opt = document.createElement("option");
      opt.value       = f.id;
      opt.textContent = "  ".repeat(depth) + f.name;
      parent.appendChild(opt);
      if (f.subFolders?.length) addOptions(parent, f.subFolders, depth + 1);
    }
  }

  for (const acc of imapAccounts) {
    const grp = document.createElement("optgroup");
    grp.label = acc.name;
    addOptions(grp, acc.folders ?? [], 0);
    select.appendChild(grp);
  }
}

// ─────────────────────────────────────────────────────────────
// SÉLECTION DES DOSSIERS
// ─────────────────────────────────────────────────────────────

/**
 * Retourne les IDs des dossiers cochés, en excluant ceux dont un ancêtre
 * est aussi coché (pour éviter de traiter le même contenu deux fois).
 */
function getSelectedFolderIds() {
  const all = [...document.querySelectorAll(".f-cb:checked")].map(cb => cb.dataset.id);
  const allSet = new Set(all);

  // Pour chaque checkbox cochée, vérifier si son nœud parent est aussi coché
  return [...document.querySelectorAll(".f-cb:checked")]
    .filter(cb => {
      const node   = cb.closest(".f-node");
      const parent = node?.parentElement?.closest(".f-node");
      const parentCb = parent?.querySelector(":scope > .f-header > .f-cb");
      // On garde ce dossier si son parent direct n'est pas coché
      return !parentCb || !parentCb.checked;
    })
    .map(cb => cb.dataset.id);
}

// ─────────────────────────────────────────────────────────────
// PROGRESSION
// ─────────────────────────────────────────────────────────────

function updateProgress(msg) {
  const done  = msg.done  || 0;
  const total = msg.total || 0;
  const pct   = total > 0 ? Math.round(done / total * 100) : 0;

  $("prog-fill").style.width = `${pct}%`;
  $("prog-count").textContent = `${done} / ${total}`;
  $("prog-pct").textContent   = `${pct}%`;

  const meta = [];
  if (msg.dupes)   meta.push(`${msg.dupes} doublon(s) détecté(s)`);
  if (msg.errors)  meta.push(`${msg.errors} erreur(s)`);
  $("prog-meta").textContent = meta.join(" · ");

  if (msg.currentFolder) {
    $("prog-folder").textContent = `📁 ${msg.currentFolder}`;
  }

  // Rappel du profil de vitesse actif
  const profileLabel = { rapide: "⚡ Rapide", normal: "▶ Normal", prudent: "🐢 Prudent", lent: "🐌 Lent" };
  const speedEl = $("prog-speed");
  if (speedEl) {
    const sel = $("speed-select");
    speedEl.textContent = sel ? (profileLabel[sel.value] || "") : "";
  }
}

// ─────────────────────────────────────────────────────────────
// RÉSULTATS
// ─────────────────────────────────────────────────────────────

function showResults(msg) {
  _duplicates = msg.duplicates || [];

  goStep("step-results");

  const cancelled = msg.status === "cancelled";
  $("res-icon").textContent  = cancelled ? "⚠️" : "✅";
  $("res-title").textContent = cancelled ? "Migration annulée" : "Migration terminée";
  $("res-title").style.color = cancelled ? "var(--warning)" : "var(--success)";

  const subparts = [];
  if (cancelled) subparts.push("Opération interrompue par l'utilisateur.");
  if (msg.done)  subparts.push(`${msg.done} message(s) copié(s) avec succès.`);
  $("res-subtitle").textContent = subparts.join(" ");

  $("res-done").textContent   = msg.done   || 0;
  $("res-dupes").textContent  = _duplicates.length;
  $("res-errors").textContent = msg.errors?.length || 0;

  // Log erreurs
  const errLog = $("res-errlog");
  if (msg.errors?.length) {
    errLog.style.display = "";
    errLog.innerHTML = msg.errors.slice(0, 30).map(e =>
      `<div>⚠ ${escHtml(e.subject)} — ${escHtml(e.reason)}</div>`
    ).join("");
    if (msg.errors.length > 30) {
      errLog.innerHTML += `<div style="color:var(--text-3)">…et ${msg.errors.length - 30} autres erreurs</div>`;
    }
  } else {
    errLog.style.display = "none";
    errLog.innerHTML = "";
  }

  // Section doublons
  if (_duplicates.length > 0) {
    $("res-dupes-section").style.display = "";
    $("dupes-count-label").textContent = `${_duplicates.length} doublon(s)`;
    renderDuplicates(_duplicates);
  } else {
    $("res-dupes-section").style.display = "none";
  }

  setStatus("res-status", "", "");
  $("btn-force").disabled = false;
}

function renderDuplicates(dupes) {
  const container = $("dupes-list");
  container.innerHTML = "";

  dupes.forEach((dup, idx) => {
    const row = document.createElement("div");
    row.className = "dup-row";

    const cb = document.createElement("input");
    cb.type      = "checkbox";
    cb.className = "dup-cb";
    cb.dataset.idx = idx;

    const info = document.createElement("div");
    info.className = "dup-info";

    const subj = document.createElement("div");
    subj.className   = "dup-subject";
    subj.textContent = dup.subject || "(sans objet)";

    const meta = document.createElement("div");
    meta.className = "dup-meta";
    const dateStr = dup.date
      ? new Date(dup.date).toLocaleDateString("fr-FR", { day:"2-digit", month:"2-digit", year:"numeric" })
      : "?";
    meta.textContent = `📁 ${dup.srcFolder}  ·  ${dateStr}`;

    info.appendChild(subj);
    info.appendChild(meta);
    row.appendChild(cb);
    row.appendChild(info);
    container.appendChild(row);
  });
}

// ─────────────────────────────────────────────────────────────
// FORCE COPY DES DOUBLONS SÉLECTIONNÉS
// ─────────────────────────────────────────────────────────────

async function forceCopySelected() {
  const selectedIdxs = [...document.querySelectorAll(".dup-cb:checked")]
    .map(cb => parseInt(cb.dataset.idx));

  if (!selectedIdxs.length) {
    setStatus("res-status", "warning", "Aucun doublon sélectionné.");
    return;
  }

  const selected = selectedIdxs.map(i => _duplicates[i]);
  $("btn-force").disabled = true;
  setStatus("res-status", "info",
    `<span class="spin"></span> Copie de ${selected.length} message(s) en cours…`);

  const r = await messenger.runtime.sendMessage({
    action    : "forceCopy",
    duplicates: selected,
  });

  if (r?.error) {
    $("btn-force").disabled = false;
    setStatus("res-status", "error", `Erreur : ${escHtml(r.error)}`);
  }
  // Le résultat final arrive via onMessage (FORCE_DONE)
}

// ─────────────────────────────────────────────────────────────
// CHARGEMENT INITIAL DE L'ARBRE
// ─────────────────────────────────────────────────────────────

async function loadFolderTree() {
  $("local-tree").innerHTML = `<div class="tree-empty"><span class="spin"></span> Chargement…</div>`;
  $("dst-select").innerHTML = `<option value="">— Chargement… —</option>`;

  const r = await messenger.runtime.sendMessage({ action: "buildFolderTree" });

  if (r?.error) {
    $("local-tree").innerHTML =
      `<div class="tree-empty" style="color:var(--error)">Erreur : ${escHtml(r.error)}</div>`;
    $("dst-select").innerHTML = `<option value="">Erreur de chargement</option>`;
    return;
  }

  renderLocalTree(r.local);
  renderImapFolders(r.imap);
}

// ─────────────────────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────
// RESTAURATION D'ÉTAT AU RÉOUVERTURE DU POPUP
// ─────────────────────────────────────────────────────────────

async function restoreState() {
  // 1. Demander l'état persisté
  const saved = await messenger.runtime.sendMessage({ action: "getMigState" });
  if (!saved) return false; // Rien à restaurer

  // 2. Vérifier si une migration est encore en cours
  const { running } = await messenger.runtime.sendMessage({ action: "isRunning" });

  switch (saved.type) {

    case "COPY_PROGRESS":
      if (running) {
        // Migration toujours en cours → afficher l'écran de progression
        goStep("step-progress");
        updateProgress(saved);
        $("btn-cancel").disabled = false;
        return true;
      }
      // Le background est mort entre temps (TB redémarré) → cleanup
      await messenger.runtime.sendMessage({ action: "clearMigState" });
      return false;

    case "COPY_DONE":
      // Résultats disponibles → restaurer l'écran de résultats
      showResults(saved);
      return true;

    case "COPY_ERROR":
      // Erreur précédente → l'afficher sur l'écran de setup
      setStatus("setup-status", "error", `Dernière erreur : ${escHtml(saved.error)}`);
      return false; // On reste sur setup pour permettre de relancer

    case "FORCE_DONE":
      // Force-copy terminé mais popup fermé avant → juste afficher setup
      await messenger.runtime.sendMessage({ action: "clearMigState" });
      return false;
  }

  return false;
}

// ─────────────────────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", async () => {

  // Tenter de restaurer un état existant avant de charger l'arbre
  goStep("step-setup");
  const restored = await restoreState();

  // Charger l'arbre seulement si on est (ou reste) sur le setup
  const onSetup = $("step-setup").style.display !== "none";
  if (onSetup) await loadFolderTree();

  // ── Sélection rapide ──
  $("sel-all").addEventListener("click", () => {
    document.querySelectorAll(".f-cb").forEach(cb => { cb.checked = true; });
  });
  $("sel-none").addEventListener("click", () => {
    document.querySelectorAll(".f-cb").forEach(cb => { cb.checked = false; });
  });

  // ── Actualiser l'arbre ──
  $("btn-refresh").addEventListener("click", loadFolderTree);

  // ── Démarrer la migration ──
  $("btn-start").addEventListener("click", async () => {
    const srcFolderIds = getSelectedFolderIds();
    _dstFolderId = $("dst-select").value;

    if (!srcFolderIds.length) {
      setStatus("setup-status", "warning", "Sélectionnez au moins un dossier source.");
      return;
    }
    if (!_dstFolderId) {
      setStatus("setup-status", "warning", "Choisissez un dossier de destination Outlook.");
      return;
    }

    setStatus("setup-status", "", "");
    goStep("step-progress");
    $("prog-folder").textContent = "Initialisation de la copie…";
    $("btn-cancel").disabled = false;

    const r = await messenger.runtime.sendMessage({
      action      : "startCopy",
      srcFolderIds,
      dstFolderId : _dstFolderId,
      speedProfile: $("speed-select").value,
    });

    if (r?.error) {
      goStep("step-setup");
      setStatus("setup-status", "error", `Erreur : ${escHtml(r.error)}`);
    }
  });

  // ── Annuler ──
  $("btn-cancel").addEventListener("click", async () => {
    $("btn-cancel").disabled = true;
    $("btn-cancel").textContent = "Annulation…";
    await messenger.runtime.sendMessage({ action: "cancel" });
  });

  // ── Doublons : sélection rapide ──
  $("sel-dupes-all").addEventListener("click", () => {
    document.querySelectorAll(".dup-cb").forEach(cb => { cb.checked = true; });
  });
  $("sel-dupes-none").addEventListener("click", () => {
    document.querySelectorAll(".dup-cb").forEach(cb => { cb.checked = false; });
  });

  // ── Force copy ──
  $("btn-force").addEventListener("click", forceCopySelected);

  // ── Recommencer ──
  $("btn-restart").addEventListener("click", async () => {
    // Effacer l'état persisté
    await messenger.runtime.sendMessage({ action: "clearMigState" });

    // Reset progress display
    $("prog-fill").style.width  = "0%";
    $("prog-count").textContent = "0 / 0";
    $("prog-pct").textContent   = "0%";
    $("prog-meta").textContent  = "";
    $("prog-folder").textContent = "";
    $("btn-cancel").disabled    = false;
    $("btn-cancel").textContent = "✕ Annuler la copie";

    // Reset results
    $("res-errlog").style.display = "none";
    $("res-errlog").innerHTML     = "";
    setStatus("res-status", "", "");
    $("btn-force").disabled = false;
    _duplicates = [];

    goStep("step-setup");
    await loadFolderTree();
  });

  // ── Écoute des messages background ──
  messenger.runtime.onMessage.addListener(msg => {
    if (!msg?.type) return;

    switch (msg.type) {

      case "COPY_PROGRESS":
        updateProgress(msg);
        break;

      case "COPY_DONE":
        showResults(msg);
        break;

      case "COPY_ERROR":
        goStep("step-setup");
        setStatus("setup-status", "error", `Erreur : ${escHtml(msg.error)}`);
        break;

      case "FORCE_PROGRESS":
        setStatus("res-status", "info",
          `<span class="spin"></span> ${msg.done} / ${msg.total} message(s) copiés…`);
        break;

      case "FORCE_DONE": {
        $("btn-force").disabled = false;
        const errInfo = msg.errors?.length
          ? ` (${msg.errors.length} erreur(s) : ${msg.errors.map(e => escHtml(e.subject)).slice(0,3).join(", ")})`
          : "";
        setStatus("res-status", "success",
          `✓ ${msg.done} message(s) copiés avec succès.${errInfo}`);
        break;
      }
    }
  });
});
