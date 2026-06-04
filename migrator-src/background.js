/**
 * Mail-Migrator CEN — background.js v1.6.7
 *
 * Stratégie : messenger.messages.copy() uniquement.
 * C'est l'équivalent API du "Copier vers" natif de Thunderbird (même code path
 * que le glisser-déposer), ce qui garantit la préservation de la date
 * originale (INTERNALDATE IMAP) — contrairement à messages.import() qui passe
 * par un APPEND brut sans INTERNALDATE.
 *
 * Persistance : chaque état important est sauvegardé dans storage.local
 * sous la clé STATE_KEY. Le popup le lit à l'ouverture pour se restaurer.
 */
"use strict";

console.log("[Mail-Migrator CEN] Chargé v1.6.7");

const STATE_KEY = "mig_state";

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────

// Profils de temporisation — l'UI transmet le nom du profil choisi
// Calcul débit : (batchSize * msgDelay + batchDelay) / batchSize = ms/msg
// M365 limite à ~3 600 msgs/heure = 1 msg/sec maximum
// copyTimeout : délai max avant d'abandonner une copie bloquée (ms)
const SPEED_PROFILES = {
  rapide  : { batchSize: 5, batchDelay: 3000, msgDelay:  500, copyTimeout:  45000 }, // ~3 200/h
  normal  : { batchSize: 3, batchDelay: 3000, msgDelay:  500, copyTimeout:  60000 }, // ~2 400/h
  prudent : { batchSize: 2, batchDelay: 4000, msgDelay:  800, copyTimeout:  90000 }, // ~1 300/h
  lent    : { batchSize: 1, batchDelay: 6000, msgDelay: 1000, copyTimeout: 120000 }, // ~500/h
  ultra   : { batchSize: 1, batchDelay:10000, msgDelay: 2000, copyTimeout: 180000 }, // ~300/h
};

// Profil actif (modifié au démarrage de chaque copie)
let CFG = {
  BATCH_SIZE   : 3,
  BATCH_DELAY  : 3000,
  MSG_DELAY    : 500,
  COPY_TIMEOUT : 60000,
  RETRY_MAX    : 3,
  RETRY_BACKOFF: 2000,
};

function applySpeedProfile(profileName) {
  const p = SPEED_PROFILES[profileName] ?? SPEED_PROFILES.normal;
  CFG.BATCH_SIZE   = p.batchSize;
  CFG.BATCH_DELAY  = p.batchDelay;
  CFG.MSG_DELAY    = p.msgDelay;
  CFG.COPY_TIMEOUT = p.copyTimeout;
  console.log(`[Migrator] Profil de vitesse : ${profileName} — batch=${p.batchSize}, batchDelay=${p.batchDelay}ms, timeout=${p.copyTimeout/1000}s`);
}

// ─────────────────────────────────────────────────────────────
// ÉTAT
// ─────────────────────────────────────────────────────────────
const state = { running: false, cancel: false };

// Journal en mémoire (inclus dans chaque snapshot persisté)
const LOG_MAX  = 150;
const logLines = [];

// ─────────────────────────────────────────────────────────────
// UTILITAIRES
// ─────────────────────────────────────────────────────────────

const sleep = ms => new Promise(r => setTimeout(r, ms));

/**
 * Enveloppe une promesse avec un timeout.
 * Si la promesse ne se résout pas dans `ms` millisecondes, rejette avec
 * une erreur explicite — évite les blocages silencieux sur IMAP APPEND.
 */
function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Délai dépassé (${ms / 1000} s) — opération bloquée`)), ms)
    ),
  ]);
}

/**
 * Ajoute une ligne au journal et la diffuse immédiatement au popup.
 * level : "info" | "ok" | "warn" | "error"
 */
function log(level, text) {
  const line = { level, text, ts: Date.now() };
  logLines.push(line);
  if (logLines.length > LOG_MAX) logLines.shift();
  messenger.runtime.sendMessage({ type: "LOG", ...line }).catch(() => {});
}

// Types à persister dans storage (pour restauration au réouverture du popup)
const PERSIST_TYPES = new Set(["COPY_PROGRESS","COPY_DONE","COPY_ERROR","FORCE_DONE"]);

function broadcast(msg) {
  // Inclure le journal dans chaque snapshot pour pouvoir le restaurer
  const payload = PERSIST_TYPES.has(msg.type) ? { ...msg, log: [...logLines] } : msg;
  messenger.runtime.sendMessage(payload).catch(() => {});
  if (PERSIST_TYPES.has(msg.type)) {
    browser.storage.local.set({ [STATE_KEY]: { ...payload, ts: Date.now() } });
  }
}

function isPermErr(e) {
  const m = (e?.message || "").toLowerCase();
  // "délai dépassé" = timeout — ne pas réessayer, l'opération est déjà abandonnée
  return ["permission denied", "quota", "no such folder", "already contains", "délai dépassé"].some(p => m.includes(p));
}

function isM365ThrottleErr(e) {
  const m = e?.message || "";
  return m.includes("2153054241") || m.includes("0x80550021");
}

async function withRetry(fn, label = "op") {
  let last;
  for (let i = 1; i <= CFG.RETRY_MAX; i++) {
    if (state.cancel) throw new Error("Annulé par l'utilisateur");
    try { return await fn(); }
    catch(e) {
      last = e;
      if (isPermErr(e)) throw e;
      if (i < CFG.RETRY_MAX) {
        console.warn(`[Migrator] ${label} échec (essai ${i}/${CFG.RETRY_MAX}) : ${e.message}`);
        await sleep(CFG.RETRY_BACKOFF * i);
      }
    }
  }
  throw last;
}

/**
 * Collecte tous les messages d'un dossier.
 * Gère l'async iterator (TB 128+) et l'ancienne API paginée.
 */
async function getAllMessages(folderId) {
  const msgs = [];
  const result = await withTimeout(messenger.messages.list(folderId), 60000);
  if (result && result[Symbol.asyncIterator]) {
    for await (const m of result) msgs.push(m);
  } else {
    let page = result;
    do {
      msgs.push(...(page?.messages ?? []));
      page = page?.id ? await messenger.messages.continueList(page.id) : null;
    } while (page?.messages?.length);
  }
  return msgs;
}

// ─────────────────────────────────────────────────────────────
// GESTION DES DOSSIERS
// ─────────────────────────────────────────────────────────────

async function ensureSubFolder(parentFolderId, name) {
  let subs = [];
  try { subs = await messenger.folders.getSubFolders(parentFolderId, false); } catch {}
  const existing = subs.find(f => f.name === name);
  if (existing) return existing;

  log("info", `📂 Création : ${name}`);
  try {
    const created = await withRetry(
      () => messenger.folders.create(parentFolderId, name),
      `create-folder-${name}`
    );
    return created;
  } catch(e) {
    // Outlook IMAP : parfois la création réussit mais lève quand même une erreur
    await sleep(1000);
    try {
      const subs2 = await messenger.folders.getSubFolders(parentFolderId, false);
      const found = subs2.find(f => f.name === name);
      if (found) { log("warn", `⚠ ${name} : erreur création mais dossier trouvé`); return found; }
    } catch {}
    log("error", `✗ Impossible de créer ${name} : ${e.message}`);
    throw e;
  }
}

// ─────────────────────────────────────────────────────────────
// COPIE RÉCURSIVE D'UN DOSSIER
// ─────────────────────────────────────────────────────────────

function snap(p) {
  return { done: p.done, total: p.total, dupes: p.duplicates.length, errors: p.errors.length };
}

async function copyFolderRecursive(srcFolder, dstFolder, progress, dateFilter, folderMap, selectedSet) {
  if (state.cancel) return;

  const srcId = srcFolder.id ?? srcFolder;
  const dstId = dstFolder.id ?? dstFolder;

  // ── Indexer les Message-ID déjà présents en destination (détection doublons)
  const dstIndex = new Set();
  try {
    const dstMsgs = await getAllMessages(dstId);
    for (const m of dstMsgs) if (m.headerMessageId) dstIndex.add(m.headerMessageId);
    if (dstMsgs.length) log("info", `🔍 "${dstFolder.name}" : ${dstMsgs.length} message(s) déjà présent(s)`);
  } catch(e) {
    log("warn", `⚠ Impossible d'indexer "${dstFolder.name}" : ${e.message}`);
  }

  // ── Lister les messages sources
  let srcMsgs = [];
  try {
    srcMsgs = await getAllMessages(srcId);
  } catch(e) {
    log("error", `✗ Lecture source "${srcFolder.name}" : ${e.message}`);
  }

  // ── Filtre par plage de dates (optionnel)
  if (dateFilter && (dateFilter.from || dateFilter.to)) {
    const before = srcMsgs.length;
    srcMsgs = srcMsgs.filter(m => {
      const ts = m.date ? new Date(m.date).getTime() : null;
      if (ts === null) return true;
      if (dateFilter.from && ts < dateFilter.from) return false;
      if (dateFilter.to   && ts > dateFilter.to)   return false;
      return true;
    });
    const excluded = before - srcMsgs.length;
    if (excluded > 0) {
      log("info", `🗓 "${srcFolder.name}" : ${excluded} hors plage ignoré(s) (${srcMsgs.length} dans la plage)`);
    }
  }

  // ── Partitionner : à copier vs doublons connus
  const toCopy = [];
  for (const m of srcMsgs) {
    if (m.headerMessageId && dstIndex.has(m.headerMessageId)) {
      progress.duplicates.push({
        id: m.id, subject: m.subject || "(sans objet)",
        date: m.date, headerMessageId: m.headerMessageId,
        srcFolder: srcFolder.name, dstFolderId: dstId,
      });
    } else {
      toCopy.push(m);
    }
  }

  const dupeCount = srcMsgs.length - toCopy.length;
  log("info", `📁 "${srcFolder.name}" — ${toCopy.length} à copier${dupeCount ? `, ${dupeCount} doublon(s) ignoré(s)` : ""}`);

  progress.total += toCopy.length;
  broadcast({ type: "COPY_PROGRESS", ...snap(progress), currentFolder: srcFolder.name });

  // ── Copie par batchs
  for (let i = 0; i < toCopy.length; i += CFG.BATCH_SIZE) {
    if (state.cancel) return;
    const batch = toCopy.slice(i, i + CFG.BATCH_SIZE);

    for (const m of batch) {
      if (state.cancel) return;
      const subj = (m.subject || "(sans objet)").substring(0, 60);
      try {
        await withRetry(
          () => withTimeout(messenger.messages.copy([m.id], dstId), CFG.COPY_TIMEOUT),
          `copy-${m.id}`
        );
        progress.done++;
        log("ok", `✓ [${progress.done}/${progress.total}] ${subj}`);
      } catch(e) {
        const errMsg = (e.message || "").toLowerCase();
        if (errMsg.includes("already contains")) {
          progress.duplicates.push({
            id: m.id, subject: m.subject || "(sans objet)",
            date: m.date, headerMessageId: m.headerMessageId,
            srcFolder: srcFolder.name, dstFolderId: dstId,
          });
          log("warn", `↩ Doublon : ${subj}`);
        } else if (isM365ThrottleErr(e)) {
          // Throttling M365 : backoff progressif 30 s → 60 s → 120 s
          let copied = false;
          for (const wait of [30000, 60000, 120000]) {
            log("warn", `⏳ Throttling M365 sur "${subj}", attente ${wait / 1000} s…`);
            await sleep(wait);
            if (state.cancel) break;
            try {
              await withTimeout(messenger.messages.copy([m.id], dstId), CFG.COPY_TIMEOUT);
              progress.done++;
              log("ok", `✓ [${progress.done}/${progress.total}] ${subj}`);
              copied = true;
              break;
            } catch(e2) {
              if (!isM365ThrottleErr(e2)) {
                progress.errors.push({ subject: m.subject || "(sans objet)", reason: e2.message });
                log("error", `✗ Erreur : ${subj} — ${e2.message}`);
                copied = true; // stopper la boucle
                break;
              }
              // encore du throttling → tentative suivante
            }
          }
          if (!copied && !state.cancel) {
            progress.errors.push({ subject: m.subject || "(sans objet)", reason: "Throttling M365 persistant (3 tentatives épuisées)" });
            log("error", `✗ Throttling persistant après 3 tentatives : ${subj}`);
          }
        } else {
          progress.errors.push({ subject: m.subject || "(sans objet)", reason: e.message });
          log("error", `✗ Erreur : ${subj} — ${e.message}`);
        }
      }
      await sleep(CFG.MSG_DELAY);
    }

    broadcast({ type: "COPY_PROGRESS", ...snap(progress), currentFolder: srcFolder.name });
    await sleep(CFG.BATCH_DELAY);
  }

  if (toCopy.length > 0) {
    log("ok", `✅ "${srcFolder.name}" terminé : ${toCopy.length} traité(s)`);
  }

  // ── Récursion sous-dossiers (via la map pré-construite)
  let subs = srcFolder.subFolders ?? [];
  if (!subs.length) {
    try { subs = await messenger.folders.getSubFolders(srcId, false); } catch {}
  }

  for (const sub of subs) {
    if (state.cancel) return;
    // Respecter la sélection de l'utilisateur
    if (!selectedSet.has(sub.id)) continue;
    const dstSub = folderMap.get(sub.id);
    if (!dstSub) {
      log("warn", `⚠ Dossier destination manquant pour "${sub.name}", ignoré`);
      continue;
    }
    try {
      await copyFolderRecursive(sub, dstSub, progress, dateFilter, folderMap, selectedSet);
    } catch(e) {
      log("error", `✗ Sous-dossier "${sub.name}" : ${e.message}`);
      progress.errors.push({ subject: `[Dossier] ${sub.name}`, reason: e.message });
    }
  }
}

// ─────────────────────────────────────────────────────────────
// PHASE 1 : PRÉ-CRÉATION DE TOUS LES DOSSIERS DESTINATION
// ─────────────────────────────────────────────────────────────

/**
 * Parcourt récursivement les dossiers source et crée tous les dossiers
 * destination en une seule passe, AVANT de copier quoi que ce soit.
 * Retourne une Map : srcFolderId → dstFolderObject
 */
async function buildDestFolderMap(srcRootFolders, dstFolderId, selectedSet) {
  const map = new Map();
  let created = 0;

  async function recurse(srcFolder, dstParentId) {
    if (state.cancel) return;
    let dstFolder;
    try {
      dstFolder = await ensureSubFolder(dstParentId, srcFolder.name);
    } catch(e) {
      log("error", `✗ Impossible de créer "${srcFolder.name}" : ${e.message}`);
      return;
    }
    map.set(srcFolder.id, dstFolder);
    created++;

    let subs = srcFolder.subFolders ?? [];
    if (!subs.length) {
      try { subs = await messenger.folders.getSubFolders(srcFolder.id, false); } catch {}
    }
    for (const sub of subs) {
      if (state.cancel) return;
      // Ne créer que les sous-dossiers explicitement cochés
      if (!selectedSet.has(sub.id)) continue;
      await recurse(sub, dstFolder.id);
    }
  }

  for (const srcFolder of srcRootFolders) {
    await recurse(srcFolder, dstFolderId);
  }

  return { map, created };
}

// ─────────────────────────────────────────────────────────────
// POINT D'ENTRÉE : DÉMARRER LA COPIE
// ─────────────────────────────────────────────────────────────

async function startCopy(srcFolderIds, dstFolderId, speedProfile = "normal", dateFrom = null, dateTo = null) {
  state.running = true;
  state.cancel  = false;
  logLines.length = 0; // Réinitialiser le journal
  applySpeedProfile(speedProfile);

  const p = SPEED_PROFILES[speedProfile] ?? SPEED_PROFILES.normal;
  log("info", `🚀 Migration démarrée — profil : ${speedProfile} (batch ${p.batchSize} msgs, ${p.batchDelay}ms entre batchs)`);

  const dateFilter = (dateFrom || dateTo) ? { from: dateFrom, to: dateTo } : null;
  if (dateFilter) {
    const fmt = ts => ts ? new Date(ts).toLocaleDateString("fr-FR", { day:"2-digit", month:"2-digit", year:"numeric" }) : "∞";
    log("info", `🗓 Filtre de date actif : du ${fmt(dateFrom)} au ${fmt(dateTo)}`);
  }

  const progress = { done: 0, total: 0, duplicates: [], errors: [] };
  broadcast({ type: "COPY_PROGRESS", ...snap(progress), currentFolder: "" });

  // ── Construire le cache de dossiers
  const accounts = await messenger.accounts.list();
  const folderCache = {};

  function cacheFolder(f) {
    if (!f?.id) return;
    folderCache[f.id] = f;
    for (const sub of (f.subFolders ?? [])) cacheFolder(sub);
  }

  for (const acc of accounts) {
    let folders = acc.folders ?? [];
    if (!folders.length) {
      try {
        folders = await messenger.folders.getSubFolders(
          acc.rootFolder?.id ?? acc.rootFolder, true
        );
      } catch {}
    }
    for (const f of folders) cacheFolder(f);
  }

  // ── Vérifier la destination
  const dstFolder = folderCache[dstFolderId];
  if (!dstFolder) {
    broadcast({ type: "COPY_ERROR", error: "Dossier destination introuvable. Actualisez et réessayez." });
    state.running = false;
    return;
  }

  // ── Dédupliquer : exclure les IDs dont un ancêtre est aussi sélectionné
  function isDescendantOf(childId, parentId) {
    const parent = folderCache[parentId];
    if (!parent?.subFolders?.length) return false;
    for (const sub of parent.subFolders) {
      if (sub.id === childId) return true;
      if (isDescendantOf(childId, sub.id)) return true;
    }
    return false;
  }

  const rootSrcIds = srcFolderIds.filter(id =>
    !srcFolderIds.some(otherId => otherId !== id && isDescendantOf(id, otherId))
  );

  // selectedSet contient TOUS les IDs cochés dans l'UI (y compris sous-dossiers)
  const selectedSet   = new Set(srcFolderIds);
  const rootSrcFolders = rootSrcIds.map(id => folderCache[id]).filter(Boolean);

  // ── Phase 1 : créer tous les dossiers destination en une seule passe
  log("info", `📂 Phase 1 — création de l'arborescence destination (${rootSrcFolders.length} dossier(s) racine)…`);
  const { map: folderMap, created: foldersCreated } = await buildDestFolderMap(rootSrcFolders, dstFolderId, selectedSet);

  if (state.cancel) {
    state.running = false;
    log("warn", "⛔ Annulé pendant la création des dossiers.");
    broadcast({ type: "COPY_DONE", done: 0, total: 0, duplicates: [], errors: [], status: "cancelled" });
    return;
  }

  log("ok", `✓ ${foldersCreated} dossier(s) prêt(s). Attente 15 s pour que M365 les enregistre…`);
  await sleep(5000); log("info", "⏳ 10 s…");
  await sleep(5000); log("info", "⏳ 5 s…");
  await sleep(5000);
  log("ok", "✅ Dossiers enregistrés — démarrage de la copie des messages.");

  // ── Phase 2 : copier les messages dossier par dossier
  log("info", "📨 Phase 2 — copie des messages…");
  for (const srcFolder of rootSrcFolders) {
    if (state.cancel) break;
    const dstSub = folderMap.get(srcFolder.id);
    if (!dstSub) {
      progress.errors.push({ subject: `[Dossier] ${srcFolder.name}`, reason: "Dossier destination non créé" });
      continue;
    }
    await copyFolderRecursive(srcFolder, dstSub, progress, dateFilter, folderMap, selectedSet);
  }

  state.running = false;

  if (state.cancel) {
    log("warn", `⛔ Migration annulée par l'utilisateur — ${progress.done} message(s) déjà copiés`);
  } else {
    log("ok", `🏁 Migration terminée — ${progress.done} copiés · ${progress.duplicates.length} doublons · ${progress.errors.length} erreurs`);
  }

  // ── Compte-rendu des messages non migrés
  if (progress.errors.length > 0) {
    log("warn", `─────────────────────────────────────────`);
    log("warn", `📋 COMPTE-RENDU — ${progress.errors.length} message(s) non migré(s) :`);
    for (const err of progress.errors) {
      log("error", `  • ${err.subject} — ${err.reason}`);
    }
    log("warn", `─────────────────────────────────────────`);
  }

  broadcast({
    type      : "COPY_DONE",
    done      : progress.done,
    total     : progress.total,
    duplicates: progress.duplicates,
    errors    : progress.errors,
    status    : state.cancel ? "cancelled" : "done",
  });
}

// ─────────────────────────────────────────────────────────────
// FORCE-COPY DES DOUBLONS SÉLECTIONNÉS
// ─────────────────────────────────────────────────────────────

async function forceCopyDuplicates(duplicates) {
  state.running = true;
  let done = 0;
  const errors = [];

  for (const dup of duplicates) {
    if (state.cancel) break;
    try {
      // messenger.messages.copy() — même mécanisme natif, préserve la date
      await withRetry(
        () => withTimeout(messenger.messages.copy([dup.id], dup.dstFolderId), CFG.COPY_TIMEOUT),
        "force-copy"
      );
      done++;
    } catch(e) {
      errors.push({ subject: dup.subject, reason: e.message });
    }
    broadcast({ type: "FORCE_PROGRESS", done, total: duplicates.length });
    await sleep(CFG.MSG_DELAY);
  }

  state.running = false;
  broadcast({ type: "FORCE_DONE", done, total: duplicates.length, errors });
}

// ─────────────────────────────────────────────────────────────
// ARBRE DES DOSSIERS POUR L'UI
// ─────────────────────────────────────────────────────────────

async function buildFolderTree() {
  const accounts = await messenger.accounts.list();
  const source = [];
  const imap   = [];

  for (const acc of accounts) {
    let folders = acc.folders ?? [];
    if (!folders.length) {
      try {
        folders = await messenger.folders.getSubFolders(
          acc.rootFolder?.id ?? acc.rootFolder, true
        );
      } catch {}
    }
    const entry = { id: acc.id, name: acc.name, type: acc.type, folders };
    source.push(entry);
    if (acc.type === "imap" || acc.type === "pop3") {
      imap.push(entry);
    }
  }

  return { source, imap };
}

// ─────────────────────────────────────────────────────────────
// BUS DE MESSAGES
// ─────────────────────────────────────────────────────────────

messenger.runtime.onMessage.addListener(async (req) => {
  try {
    switch (req.action) {

      case "buildFolderTree":
        return buildFolderTree();

      case "startCopy": {
        if (state.running) return { error: "Une migration est déjà en cours." };
        startCopy(req.srcFolderIds, req.dstFolderId, req.speedProfile, req.dateFrom ?? null, req.dateTo ?? null).catch(e => {
          state.running = false;
          broadcast({ type: "COPY_ERROR", error: e.message });
        });
        return { started: true };
      }

      case "forceCopy": {
        if (state.running) return { error: "Une opération est déjà en cours." };
        forceCopyDuplicates(req.duplicates).catch(e => {
          state.running = false;
          broadcast({ type: "COPY_ERROR", error: e.message });
        });
        return { started: true };
      }

      case "cancel":
        state.cancel = true;
        return { ok: true };

      case "isRunning":
        return { running: state.running };

      case "getMigState": {
        const s = await browser.storage.local.get(STATE_KEY);
        return s[STATE_KEY] ?? null;
      }

      case "clearMigState":
        logLines.length = 0;
        await browser.storage.local.remove(STATE_KEY);
        return { ok: true };

      default:
        return { error: `Action inconnue : ${req.action}` };
    }
  } catch(e) {
    console.error("[Migrator] Erreur bus :", e);
    return { error: e.message };
  }
});
