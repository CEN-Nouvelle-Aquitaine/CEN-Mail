/**
 * Mail-Migrator CEN — background.js v1.5.0
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

console.log("[Mail-Migrator CEN] Chargé v1.5.0");

const STATE_KEY = "mig_state";

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────

// Profils de temporisation — l'UI transmet le nom du profil choisi
const SPEED_PROFILES = {
  rapide  : { batchSize: 10, batchDelay:  600, msgDelay:  80 },
  normal  : { batchSize:  5, batchDelay: 1500, msgDelay: 200 },
  prudent : { batchSize:  3, batchDelay: 3000, msgDelay: 500 },
  lent    : { batchSize:  1, batchDelay: 5000, msgDelay: 800 },
};

// Profil actif (modifié au démarrage de chaque copie)
let CFG = {
  BATCH_SIZE   : 5,
  BATCH_DELAY  : 1500,
  MSG_DELAY    : 200,
  RETRY_MAX    : 3,
  RETRY_BACKOFF: 2000,
};

function applySpeedProfile(profileName) {
  const p = SPEED_PROFILES[profileName] ?? SPEED_PROFILES.normal;
  CFG.BATCH_SIZE  = p.batchSize;
  CFG.BATCH_DELAY = p.batchDelay;
  CFG.MSG_DELAY   = p.msgDelay;
  console.log(`[Migrator] Profil de vitesse : ${profileName} — batch=${p.batchSize}, batchDelay=${p.batchDelay}ms, msgDelay=${p.msgDelay}ms`);
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
    messenger.storage.local.set({ [STATE_KEY]: { ...payload, ts: Date.now() } });
  }
}

function isPermErr(e) {
  const m = (e?.message || "").toLowerCase();
  return ["permission denied", "quota", "no such folder", "already contains"].some(p => m.includes(p));
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
  const result = await messenger.messages.list(folderId);
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

  log("info", `📂 Création dossier : ${name}`);
  try {
    const created = await withRetry(
      () => messenger.folders.create(parentFolderId, name),
      `create-folder-${name}`
    );
    log("ok", `✓ Dossier créé : ${name}`);
    return created;
  } catch(e) {
    // Outlook IMAP : parfois la création réussit mais lève quand même une erreur
    await sleep(1000);
    try {
      const subs2 = await messenger.folders.getSubFolders(parentFolderId, false);
      const found = subs2.find(f => f.name === name);
      if (found) { log("warn", `⚠ Dossier ${name} : erreur à la création mais dossier existant trouvé`); return found; }
    } catch {}
    log("error", `✗ Impossible de créer le dossier ${name} : ${e.message}`);
    throw e;
  }
}

// ─────────────────────────────────────────────────────────────
// COPIE RÉCURSIVE D'UN DOSSIER
// ─────────────────────────────────────────────────────────────

function snap(p) {
  return { done: p.done, total: p.total, dupes: p.duplicates.length, errors: p.errors.length };
}

async function copyFolderRecursive(srcFolder, dstFolder, progress, dateFilter = null) {
  if (state.cancel) return;

  const srcId = srcFolder.id ?? srcFolder;
  const dstId = dstFolder.id ?? dstFolder;

  // ── Indexer les Message-ID déjà présents en destination (détection doublons)
  const dstIndex = new Set();
  try {
    const dstMsgs = await getAllMessages(dstId);
    for (const m of dstMsgs) if (m.headerMessageId) dstIndex.add(m.headerMessageId);
    if (dstMsgs.length) log("info", `🔍 Index destination "${dstFolder.name}" : ${dstMsgs.length} message(s) existant(s)`);
  } catch(e) {
    log("warn", `⚠ Impossible d'indexer "${dstFolder.name}" : ${e.message}`);
  }

  // ── Lister les messages sources
  let srcMsgs = [];
  try {
    srcMsgs = await getAllMessages(srcId);
  } catch(e) {
    log("error", `✗ Lecture dossier source "${srcFolder.name}" : ${e.message}`);
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
      log("info", `🗓 "${srcFolder.name}" : ${excluded} message(s) hors plage ignoré(s) (${srcMsgs.length} dans la plage)`);
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

  // ── Copie par batchs via l'API native Thunderbird (même mécanisme que "Copier vers")
  for (let i = 0; i < toCopy.length; i += CFG.BATCH_SIZE) {
    if (state.cancel) return;
    const batch = toCopy.slice(i, i + CFG.BATCH_SIZE);

    for (const m of batch) {
      if (state.cancel) return;
      const subj = (m.subject || "(sans objet)").substring(0, 60);
      try {
        await withRetry(
          () => messenger.messages.copy([m.id], dstId),
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
    log("ok", `✅ "${srcFolder.name}" terminé : ${toCopy.length} message(s) traité(s)`);
  }

  // ── Récursion sous-dossiers
  let subs = srcFolder.subFolders ?? [];
  if (!subs.length) {
    try { subs = await messenger.folders.getSubFolders(srcId, false); } catch {}
  }

  for (const sub of subs) {
    if (state.cancel) return;
    try {
      const dstSub = await ensureSubFolder(dstId, sub.name);
      await copyFolderRecursive(sub, dstSub, progress, dateFilter);
    } catch(e) {
      log("error", `✗ Sous-dossier "${sub.name}" : ${e.message}`);
      progress.errors.push({ subject: `[Dossier] ${sub.name}`, reason: e.message });
    }
  }
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

  // ── Copier chaque dossier racine sélectionné
  for (const srcId of rootSrcIds) {
    if (state.cancel) break;
    const srcFolder = folderCache[srcId];
    if (!srcFolder) continue;

    let dstSub;
    try {
      dstSub = await ensureSubFolder(dstFolderId, srcFolder.name);
    } catch(e) {
      progress.errors.push({ subject: `[Dossier] ${srcFolder.name}`, reason: e.message });
      continue;
    }

    await copyFolderRecursive(srcFolder, dstSub, progress, dateFilter);
  }

  state.running = false;

  if (state.cancel) {
    log("warn", `⛔ Migration annulée par l'utilisateur — ${progress.done} message(s) déjà copiés`);
  } else {
    log("ok", `🏁 Migration terminée — ${progress.done} copiés · ${progress.duplicates.length} doublons · ${progress.errors.length} erreurs`);
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
        () => messenger.messages.copy([dup.id], dup.dstFolderId),
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
  const local = [];
  const imap  = [];

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
    if (acc.type === "none" || acc.type === "local") {
      local.push(entry);
    } else if (acc.type === "imap" || acc.type === "pop3") {
      imap.push(entry);
    }
  }

  return { local, imap };
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
        const s = await messenger.storage.local.get(STATE_KEY);
        return s[STATE_KEY] ?? null;
      }

      case "clearMigState":
        logLines.length = 0;
        await messenger.storage.local.remove(STATE_KEY);
        return { ok: true };

      default:
        return { error: `Action inconnue : ${req.action}` };
    }
  } catch(e) {
    console.error("[Migrator] Erreur bus :", e);
    return { error: e.message };
  }
});
