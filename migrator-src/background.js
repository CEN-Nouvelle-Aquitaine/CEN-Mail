/**
 * Mail-Migrator CEN — background.js v1.0.0
 * Copie dossiers locaux → Outlook en préservant les dates (INTERNALDATE IMAP).
 *
 * Correction clé par rapport à l'ancienne version :
 *   messages.import(file, dstId, { date: message.date, … })
 *   Le champ `date` est transmis dans la commande IMAP APPEND comme INTERNALDATE,
 *   ce qui évite qu'Outlook affiche la date du transfert au lieu de la date d'envoi.
 */
"use strict";

console.log("[Mail-Migrator CEN] Chargé v1.0.0");

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────
const CFG = {
  BATCH_SIZE   : 5,     // messages par batch (throttle IMAP Outlook)
  BATCH_DELAY  : 1200,  // ms entre batchs
  MSG_DELAY    : 150,   // ms entre messages au sein d'un batch
  RETRY_MAX    : 3,
  RETRY_BACKOFF: 2000,  // ms, multiplié par le numéro d'essai
  TEMP_FOLDER  : "MigTemp-CEN",
};

// ─────────────────────────────────────────────────────────────
// ÉTAT
// ─────────────────────────────────────────────────────────────
const state = { running: false, cancel: false };
let _tempFolder = null;

// ─────────────────────────────────────────────────────────────
// UTILITAIRES
// ─────────────────────────────────────────────────────────────

const sleep = ms => new Promise(r => setTimeout(r, ms));

function broadcast(msg) {
  messenger.runtime.sendMessage(msg).catch(() => {});
}

function isPermErr(e) {
  const m = (e?.message || "").toLowerCase();
  return ["already contains", "permission denied", "quota", "no such folder"].some(p => m.includes(p));
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
 * Collecte tous les messages d'un dossier (gère la pagination TB 128+ et l'ancienne API).
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

/**
 * Récupère le message en tant que File (octets bruts) sans décodage UTF-8.
 * Ne JAMAIS décoder en string puis ré-encoder — cela corrompt les caractères multi-octets.
 */
async function getRawFile(messageId) {
  const raw = await messenger.messages.getRaw(messageId);
  if (raw instanceof Blob) {
    return new File([raw], `${messageId}.eml`, { type: "message/rfc822" });
  }
  // Anciens TB : BinaryString (1 char = 1 octet)
  const bytes = new Uint8Array(raw.length);
  for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i) & 0xff;
  return new File([bytes], `${messageId}.eml`, { type: "message/rfc822" });
}

// ─────────────────────────────────────────────────────────────
// DOSSIER TEMPORAIRE LOCAL
// ─────────────────────────────────────────────────────────────

async function getTempFolder() {
  if (_tempFolder) {
    // Vérifie que le dossier existe encore
    try {
      await messenger.folders.getSubFolders(_tempFolder.id, false);
      return _tempFolder;
    } catch { _tempFolder = null; }
  }
  const accounts = await messenger.accounts.list();
  const local =
    accounts.find(a => a.type === "none") ??
    accounts.find(a => a.type === "local") ??
    accounts[0];
  if (!local) throw new Error("Aucun compte local disponible pour le dossier temporaire.");
  const rootId = local.rootFolder?.id ?? local.rootFolder;
  const subs = await messenger.folders.getSubFolders(rootId, false);
  _tempFolder = subs.find(f => f.name === CFG.TEMP_FOLDER)
    ?? await messenger.folders.create(rootId, CFG.TEMP_FOLDER);
  return _tempFolder;
}

async function cleanTempFolder() {
  if (!_tempFolder) return;
  try {
    const msgs = await getAllMessages(_tempFolder.id ?? _tempFolder);
    if (!msgs.length) {
      await messenger.folders.delete(_tempFolder.id ?? _tempFolder).catch(() => {});
      _tempFolder = null;
    }
  } catch {}
}

// ─────────────────────────────────────────────────────────────
// COPIE D'UN MESSAGE AVEC PRÉSERVATION DE DATE
// ─────────────────────────────────────────────────────────────

/**
 * Copie un message vers dstFolderId en préservant la date originale.
 *
 * POURQUOI ça marchait pas avant :
 *   L'ancienne version appelait messages.import() sans le champ `date`.
 *   Résultat : l'IMAP APPEND ne transmettait pas d'INTERNALDATE → Outlook
 *   utilisait l'horodatage de réception comme date, pas la date du message.
 *
 * SOLUTION :
 *   Passer `date: message.date` dans les props d'import. TB 128+ inclut
 *   ce champ dans la commande IMAP APPEND → INTERNALDATE = date originale.
 */
async function copyMessagePreservingDate(message, dstFolderId) {
  const props = {
    date   : message.date,        // ← LE FIX : INTERNALDATE IMAP = date originale
    flagged: message.flagged,
    read   : message.read,
    tags   : message.tags ?? [],
  };

  let file;
  try {
    await withRetry(async () => { file = await getRawFile(message.id); }, "getRaw");
  } catch(e) {
    return { error: `Lecture du message impossible : ${e.message}` };
  }

  // ── Stratégie 1 : import direct vers IMAP (APPEND avec INTERNALDATE)
  try {
    await withRetry(async () => {
      const msg = await messenger.messages.import(file, dstFolderId, props);
      if (!msg) throw new Error("Import retourné null");
    }, "import-direct");
    return { ok: true, method: "direct" };
  } catch(e) {
    if ((e.message || "").toLowerCase().includes("already contains")) {
      return { skipped: true, reason: "doublon" };
    }
    console.warn(`[Migrator] Import direct échoué (${e.message}), essai fallback temp…`);
  }

  // ── Stratégie 2 : import vers dossier local temp, puis déplacement IMAP
  // Le move() TB préserve la date mieux que certains serveurs IMAP qui ignorent
  // l'INTERNALDATE dans l'APPEND.
  try {
    const temp = await getTempFolder();
    let localMsg;
    await withRetry(async () => {
      localMsg = await messenger.messages.import(file, temp.id ?? temp, props);
      if (!localMsg) throw new Error("Import temp null");
    }, "import-temp");

    await withRetry(async () => {
      await messenger.messages.move([localMsg.id], dstFolderId);
    }, "move-from-temp");

    return { ok: true, method: "temp+move" };
  } catch(e) {
    if ((e.message || "").toLowerCase().includes("already contains")) {
      return { skipped: true, reason: "doublon" };
    }
    return { error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
// GESTION DES DOSSIERS
// ─────────────────────────────────────────────────────────────

async function ensureSubFolder(parentFolderId, name) {
  let subs = [];
  try { subs = await messenger.folders.getSubFolders(parentFolderId, false); } catch {}
  const existing = subs.find(f => f.name === name);
  if (existing) return existing;

  try {
    return await withRetry(
      () => messenger.folders.create(parentFolderId, name),
      `create-folder-${name}`
    );
  } catch(e) {
    // Outlook IMAP : parfois la création réussit mais lève quand même une erreur
    await sleep(1000);
    try {
      const subs2 = await messenger.folders.getSubFolders(parentFolderId, false);
      const found = subs2.find(f => f.name === name);
      if (found) return found;
    } catch {}
    throw e;
  }
}

// ─────────────────────────────────────────────────────────────
// COPIE RÉCURSIVE D'UN DOSSIER
// ─────────────────────────────────────────────────────────────

function snap(p) {
  return {
    done  : p.done,
    total : p.total,
    dupes : p.duplicates.length,
    errors: p.errors.length,
  };
}

async function copyFolderRecursive(srcFolder, dstFolder, progress) {
  if (state.cancel) return;

  const srcId = srcFolder.id ?? srcFolder;
  const dstId = dstFolder.id ?? dstFolder;

  // ── Index des Message-ID déjà présents en destination (détection doublons)
  const dstIndex = new Set();
  try {
    const dstMsgs = await getAllMessages(dstId);
    for (const m of dstMsgs) if (m.headerMessageId) dstIndex.add(m.headerMessageId);
  } catch(e) {
    console.warn(`[Migrator] Impossible d'indexer destination "${dstFolder.name}" :`, e.message);
  }

  // ── Lister les messages sources
  let srcMsgs = [];
  try {
    srcMsgs = await getAllMessages(srcId);
  } catch(e) {
    console.warn(`[Migrator] list "${srcFolder.name}" :`, e.message);
  }

  // ── Partitionner : à copier vs doublons connus
  const toCopy = [];
  for (const m of srcMsgs) {
    if (m.headerMessageId && dstIndex.has(m.headerMessageId)) {
      progress.duplicates.push({
        id             : m.id,
        subject        : m.subject || "(sans objet)",
        date           : m.date,
        headerMessageId: m.headerMessageId,
        srcFolder      : srcFolder.name,
        dstFolderId    : dstId,
      });
    } else {
      toCopy.push(m);
    }
  }

  progress.total += toCopy.length;
  broadcast({ type: "COPY_PROGRESS", ...snap(progress), currentFolder: srcFolder.name });

  // ── Copie par batchs
  for (let i = 0; i < toCopy.length; i += CFG.BATCH_SIZE) {
    if (state.cancel) return;
    const batch = toCopy.slice(i, i + CFG.BATCH_SIZE);

    for (const m of batch) {
      if (state.cancel) return;
      const result = await copyMessagePreservingDate(m, dstId);
      if (result.ok)      progress.done++;
      else if (result.skipped) {
        progress.duplicates.push({
          id: m.id, subject: m.subject || "(sans objet)",
          date: m.date, headerMessageId: m.headerMessageId,
          srcFolder: srcFolder.name, dstFolderId: dstId,
        });
      } else {
        progress.errors.push({ subject: m.subject || "(sans objet)", reason: result.error });
      }
      await sleep(CFG.MSG_DELAY);
    }

    broadcast({ type: "COPY_PROGRESS", ...snap(progress), currentFolder: srcFolder.name });
    await sleep(CFG.BATCH_DELAY);
  }

  // ── Récursion sous-dossiers
  let subs = srcFolder.subFolders ?? [];
  if (!subs.length) {
    try { subs = await messenger.folders.getSubFolders(srcId, false); } catch {}
  }

  for (const sub of subs) {
    if (state.cancel) return;
    if (sub.name === CFG.TEMP_FOLDER) continue;
    try {
      const dstSub = await ensureSubFolder(dstId, sub.name);
      await copyFolderRecursive(sub, dstSub, progress);
    } catch(e) {
      progress.errors.push({ subject: `[Dossier] ${sub.name}`, reason: e.message });
    }
  }
}

// ─────────────────────────────────────────────────────────────
// POINT D'ENTRÉE : DÉMARRER LA COPIE
// ─────────────────────────────────────────────────────────────

async function startCopy(srcFolderIds, dstFolderId) {
  state.running = true;
  state.cancel  = false;
  _tempFolder   = null;

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

  // ── Dédupliquer srcFolderIds : retirer tout ID dont un ancêtre est aussi sélectionné
  // Pour cela on vérifie si le parent immédiat du dossier figure dans la liste
  const srcSet = new Set(srcFolderIds);

  function hasSelectedAncestor(folderId) {
    const f = folderCache[folderId];
    if (!f) return false;
    // Cherche si un dossier de la liste est parent de folderId
    for (const candidateId of srcSet) {
      if (candidateId === folderId) continue;
      const candidate = folderCache[candidateId];
      if (!candidate) continue;
      // Vérifie si folderId est dans les sous-dossiers de candidate
      if (isDescendantOf(folderId, candidateId, folderCache)) return true;
    }
    return false;
  }

  function isDescendantOf(childId, parentId, cache) {
    const parent = cache[parentId];
    if (!parent?.subFolders?.length) return false;
    for (const sub of parent.subFolders) {
      if (sub.id === childId) return true;
      if (isDescendantOf(childId, sub.id, cache)) return true;
    }
    return false;
  }

  const rootSrcIds = srcFolderIds.filter(id => !hasSelectedAncestor(id));

  // ── Copier chaque dossier racine sélectionné vers la destination
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

    await copyFolderRecursive(srcFolder, dstSub, progress);
  }

  state.running = false;
  await cleanTempFolder();

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
      const message = await messenger.messages.get(dup.id);
      if (!message) throw new Error("Message introuvable (a-t-il été supprimé ?)");

      const file = await getRawFile(dup.id);
      await withRetry(() => messenger.messages.import(file, dup.dstFolderId, {
        date   : message.date,   // préservation de date ici aussi
        flagged: message.flagged,
        read   : message.read,
        tags   : message.tags ?? [],
      }), "force-import");
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
        startCopy(req.srcFolderIds, req.dstFolderId).catch(e => {
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

      default:
        return { error: `Action inconnue : ${req.action}` };
    }
  } catch(e) {
    console.error("[Migrator] Erreur bus :", e);
    return { error: e.message };
  }
});
