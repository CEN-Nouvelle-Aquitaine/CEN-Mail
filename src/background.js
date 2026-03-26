/**
 * Mail-CEN background.js v5.3
 * Modules : M365 · Étiquettes · Migration · Synchronisation · Export · Tags
 */
"use strict";
console.log("[Mail-CEN] Chargement v5.3");

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────
const CFG = {
  BATCH_SIZE    : 20,
  BATCH_DELAY   : 600,
  POLL_RETRIES  : 30,
  POLL_INTERVAL : 500,
  TEMP_FOLDER   : "Mail-CEN-Temp",
  CHECKPOINT_KEY: "cen_checkpoint",
  MAPPING_KEY   : "cen_label_mapping",
  MIG_STATE_KEY : "cen_mig_state",
};

// Outlook catégories standard (clé IMAP → couleur hex)
const OL_CATEGORIES = {
  "Red Category"    : "#ef4444",
  "Orange Category" : "#f97316",
  "Yellow Category" : "#eab308",
  "Green Category"  : "#22c55e",
  "Blue Category"   : "#3b82f6",
  "Purple Category" : "#a855f7",
};

const mig = { running: false, cancel: false };
let _folderCache    = null;
let _accountIdCache = {};

// ─────────────────────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));

function broadcast(msg) {
  messenger.runtime.sendMessage(msg).catch(() => {});
  if (["MIG_PROGRESS","MIG_DONE","MIG_ERROR","SYNC_PROGRESS","SYNC_DONE"].includes(msg.type))
    messenger.storage.local.set({ [CFG.MIG_STATE_KEY]: { ...msg, ts: Date.now() } });
}

function encodeHeader(name, value) {
  const S = " =?utf-8?q?", NL = "=?=", E = "?=";
  let lines = [], cur = `${name}:${S}`;
  for (const b of new TextEncoder().encode(value)) {
    const c = String.fromCharCode(b);
    const enc = /[A-Za-z0-9!*+\-\/]/.test(c) && c !== " "
      ? c : "=" + b.toString(16).toUpperCase().padStart(2,"0");
    if (cur.length + enc.length + E.length > 78) { lines.push(cur + NL); cur = S; }
    cur += enc;
  }
  lines.push(cur + E);
  return lines.join("\r\n");
}

// ─────────────────────────────────────────────────────────────
// DOSSIER TEMPORAIRE
// ─────────────────────────────────────────────────────────────
async function getTempFolder() {
  const accounts = await messenger.accounts.list(false);
  const local = accounts.find(a => a.type === "none");
  if (!local) throw new Error("Aucun compte local trouvé.");
  const subs = await messenger.folders.getSubFolders(local.rootFolder.id, false);
  let temp = subs.find(f => f.name === CFG.TEMP_FOLDER);
  if (!temp) temp = await messenger.folders.create(local.rootFolder.id, CFG.TEMP_FOLDER);
  return temp;
}

// ─────────────────────────────────────────────────────────────
// MODULE DOSSIERS
// ─────────────────────────────────────────────────────────────
async function buildFolderTree() {
  const accounts = await messenger.accounts.list();
  const flat = [];
  _folderCache    = {};
  _accountIdCache = {};

  function walk(folders, depth, accId, accName, accType) {
    for (const f of folders) {
      flat.push({ id: f.id, name: f.name, path: f.path, type: f.type,
                  accountId: accId, accountName: accName, accountType: accType, depth });
      _folderCache[f.id]    = f;
      _accountIdCache[f.id] = accId;
      if (f.subFolders?.length) walk(f.subFolders, depth+1, accId, accName, accType);
    }
  }
  for (const acc of accounts) {
    const rootFolders = acc.folders ?? await messenger.folders.getSubFolders(acc.rootFolder.id, false);
    walk(rootFolders, 0, acc.id, acc.name, acc.type);
  }
  return flat;
}

async function ensureFolder(parentFolder, name) {
  const pid = parentFolder.id ?? parentFolder;
  const subs = await messenger.folders.getSubFolders(pid, false);
  const ex = subs.find(f => f.name === name);
  if (ex) return ex;
  return await messenger.folders.create(pid, name);
}

// ─────────────────────────────────────────────────────────────
// MODULE ÉTIQUETTES / MAPPING
// ─────────────────────────────────────────────────────────────
async function getAllTbTags() {
  return await messenger.messages.tags.list();
}

async function loadMapping() {
  const d = await messenger.storage.local.get(CFG.MAPPING_KEY);
  return d[CFG.MAPPING_KEY] ?? {};
}

async function saveMapping(mapping) {
  await messenger.storage.local.set({ [CFG.MAPPING_KEY]: mapping });
}

/**
 * Résout les tags TB d'un message en mots-clés Outlook selon le mapping.
 * Retourne un tableau de noms de catégories Outlook.
 */
async function resolveMappedCategories(tbTagKeys) {
  if (!tbTagKeys?.length) return [];
  const mapping = await loadMapping();
  return tbTagKeys
    .map(k => mapping[k])
    .filter(v => v && v !== "__skip__");
}

// ─────────────────────────────────────────────────────────────
// MODULE SUBJECT-TAG
// ─────────────────────────────────────────────────────────────
async function getTagNamesForMessage(message) {
  if (!message.tags?.length) return [];
  const all = await getAllTbTags();
  return message.tags.map(k => all.find(t => t.key === k)?.tag).filter(Boolean);
}

async function applyTagsToSubject(message) {
  const tagNames = await getTagNamesForMessage(message);
  if (!tagNames.length) return null;

  const full = await messenger.messages.getFull(message.id);
  const currentSubject =
    (Array.isArray(full.headers.subject) && full.headers.subject[0]) ||
    message.subject || "";
  const prefix = tagNames.map(t => `{${t}}`).join("");
  if (currentSubject.startsWith(prefix)) return null;
  const newSubject = prefix + currentSubject;

  let raw = (await messenger.messages.getRaw(message.id))
    .replace(/\r/g,"").replace(/\n/g,"\r\n");
  const hdrEnd = raw.search(/\r\n\r\n/);
  if (hdrEnd === -1) throw new Error("Message malformé.");
  let hdr  = "\r\n" + raw.substring(0, hdrEnd+2).replace(/\r\r/,"\r");
  const body = raw.substring(hdrEnd+2);

  while (/\r\nSubject: .*\r\n\s+/.test(hdr))
    hdr = hdr.replace(/(Subject: .*)(\r\n\s+)/,"$1 ");
  if (hdr.includes("\nSubject: "))
    hdr = hdr.replace(/\nSubject: .*\r\n/, "\n" + encodeHeader("Subject", newSubject) + "\r\n");
  else
    hdr += encodeHeader("Subject", newSubject) + "\r\n";

  const server = message.headerMessageId.split("@").pop();
  const uid = crypto.randomUUID();
  hdr = hdr.replace(/\nMessage-ID: *.*\r\n/i, `\nMessage-ID: <${uid}@${server}>\r\n`);
  if (!/\nDate: /i.test(hdr) && message.date)
    hdr += `Date: ${message.date.toUTCString()}\r\n`;

  const ts = new Date().toString().replace(/\(.+\)/,"").substring(0,60);
  const xhdr = `X-Subject-Tag: ${ts}`;
  if (!hdr.includes("\nX-Subject-Tag: "))
    hdr += xhdr + "\r\n" + encodeHeader("X-Subject-Tag-OriginalSubject", currentSubject) + "\r\n";
  else
    hdr = hdr.replace(/\nX-Subject-Tag: .+\r\n/, `\n${xhdr}\r\n`);
  hdr = hdr.substring(2);

  const content = hdr + body;
  const bytes = new Uint8Array(content.length);
  for (let i = 0; i < content.length; i++) bytes[i] = content.charCodeAt(i) & 0xff;
  const file = new File([bytes], `${uid}.eml`, { type:"message/rfc822" });

  const temp = await getTempFolder();
  const localMsg = await messenger.messages.import(file, temp.id ?? temp, {
    flagged: message.flagged, read: message.read, tags: message.tags,
  });
  if (!localMsg) throw new Error("Import échoué.");

  const moved = await new Promise((resolve, reject) => {
    let tries = 0;
    const poll = async () => {
      const pg = await messenger.messages.query({
        folderId: message.folder.id ?? message.folder, headerMessageId: localMsg.headerMessageId });
      let page = pg;
      do {
        const found = page.messages.find(m => m.headerMessageId === localMsg.headerMessageId);
        if (found) { resolve(found); return; }
        page = page.id ? await messenger.messages.continueList(page.id) : null;
      } while (page?.messages.length);
      if (++tries > CFG.POLL_RETRIES) { reject(new Error("Message non retrouvé.")); return; }
      setTimeout(poll, CFG.POLL_INTERVAL);
    };
    messenger.messages.move([localMsg.id], message.folder.id ?? message.folder);
    setTimeout(poll, CFG.POLL_INTERVAL);
  });
  await messenger.messages.move([message.id], temp.id ?? temp);
  return moved;
}

async function runSubjectTagOnIds(ids) {
  let ok=0, skip=0, err=0, errors=[];
  for (const id of ids) {
    try {
      const msg = await messenger.messages.get(id);
      if (!msg) { skip++; continue; }
      const r = await applyTagsToSubject(msg);
      r ? ok++ : skip++;
    } catch(e) { err++; errors.push({ id, reason: e.message }); }
  }
  return { ok, skip, err, errors };
}

async function runSubjectTagAll() {
  const accounts = await messenger.accounts.list();
  const tagged = [];
  async function scan(folders) {
    for (const folder of folders) {
      if (folder.name === CFG.TEMP_FOLDER) continue;
      try {
        let page = await messenger.messages.list(folder.id ?? folder);
        do {
          tagged.push(...page.messages.filter(m => m.tags?.length));
          page = page.id ? await messenger.messages.continueList(page.id) : null;
        } while (page);
      } catch {}
      if (folder.subFolders?.length) await scan(folder.subFolders);
    }
  }
  for (const acc of accounts) {
    const rootFolders = acc.folders ?? await messenger.folders.getSubFolders(acc.rootFolder.id, false);
    await scan(rootFolders);
  }
  if (!tagged.length) return { ok:0, skip:0, err:0, total:0 };
  let ok=0, skip=0, err=0;
  for (const msg of tagged) {
    try { const r = await applyTagsToSubject(msg); r ? ok++ : skip++; } catch { err++; }
    broadcast({ type:"ST_PROGRESS", done: ok+skip+err, total: tagged.length });
  }
  return { ok, skip, err, total: tagged.length };
}

// ─────────────────────────────────────────────────────────────
// MODULE MIGRATION
// ─────────────────────────────────────────────────────────────

/** Import local temp → move vers IMAP (préserve la date via INTERNALDATE) */
async function importViaLocalTemp(file, dstFolder, props) {
  const temp = await getTempFolder();
  const localMsg = await messenger.messages.import(file, temp.id ?? temp, props);
  if (!localMsg) throw new Error("Import local échoué.");
  await messenger.messages.move([localMsg.id], dstFolder.id ?? dstFolder);
  return localMsg;
}

async function migrateFolderRecursive(srcFolder, dstFolder, srcAccId, dstAccId, mode, progress) {
  if (mig.cancel) return;
  const crossAccount = srcAccId !== dstAccId;

  // Index des Message-ID déjà présents dans la destination (anti-doublons)
  const dstMsgIds = new Set();
  try {
    let dstPage = await messenger.messages.list(dstFolder.id ?? dstFolder);
    do {
      for (const m of dstPage.messages) {
        if (m.headerMessageId) dstMsgIds.add(m.headerMessageId);
      }
      dstPage = dstPage.id ? await messenger.messages.continueList(dstPage.id) : null;
    } while (dstPage);
  } catch {}

  let all = [];
  try {
    let page = await messenger.messages.list(srcFolder.id ?? srcFolder);
    do {
      all.push(...page.messages);
      page = page.id ? await messenger.messages.continueList(page.id) : null;
    } while (page);
  } catch(e) { console.warn(`[Migration] Dossier ${srcFolder.name}:`, e); }

  // Filtrer les doublons
  const before = all.length;
  all = all.filter(m => !m.headerMessageId || !dstMsgIds.has(m.headerMessageId));
  const skipped = before - all.length;
  if (skipped > 0) {
    progress.skipped = (progress.skipped || 0) + skipped;
    console.log(`[Migration] ${srcFolder.name}: ${skipped} doublon(s) ignoré(s)`);
  }

  progress.total += all.length;
  broadcast({ type:"MIG_PROGRESS", done: progress.done, total: progress.total, skipped: progress.skipped || 0, errors: progress.errors });

  for (let i = 0; i < all.length; i += CFG.BATCH_SIZE) {
    if (mig.cancel) return;
    const batch = all.slice(i, i + CFG.BATCH_SIZE);

    if (!crossAccount) {
      const ids = batch.map(m => m.id);
      const dstId = dstFolder.id ?? dstFolder;
      try {
        mode === "move"
          ? await messenger.messages.move(ids, dstId)
          : await messenger.messages.copy(ids, dstId);
        progress.done += ids.length;
      } catch {
        for (const m of batch) {
          if (mig.cancel) return;
          try {
            mode === "move"
              ? await messenger.messages.move([m.id], dstId)
              : await messenger.messages.copy([m.id], dstId);
            progress.done++;
          } catch(e) {
            progress.errors.push({ id: m.id, subject: m.subject, reason: e.message });
          }
        }
      }
    } else {
      for (const m of batch) {
        if (mig.cancel) return;
        try {
          const raw = await messenger.messages.getRaw(m.id);
          const content = raw.replace(/\r/g,"").replace(/\n/g,"\r\n");
          const bytes = new Uint8Array(content.length);
          for (let k=0; k<content.length; k++) bytes[k] = content.charCodeAt(k) & 0xff;
          const file = new File([bytes], `${m.id}.eml`, { type:"message/rfc822" });

          await importViaLocalTemp(file, dstFolder, {
            flagged: m.flagged, read: m.read, tags: m.tags ?? [],
          });
          if (mode === "move") await messenger.messages.delete([m.id], true);
          progress.done++;
        } catch(e) {
          progress.errors.push({ id: m.id, subject: m.subject, reason: e.message });
        }
        if (progress.done % 5 === 0) await sleep(50);
      }
    }

    broadcast({ type:"MIG_PROGRESS", done: progress.done, total: progress.total, errors: progress.errors });
    await sleep(CFG.BATCH_DELAY);
  }

  // Récursion sous-dossiers
  for (const sub of (srcFolder.subFolders ?? [])) {
    if (mig.cancel) return;
    if (sub.name === CFG.TEMP_FOLDER) continue;
    let targetSub;
    try { targetSub = await ensureFolder(dstFolder, sub.name); }
    catch(e) {
      progress.errors.push({ id:null, subject:`[Dossier] ${sub.name}`, reason: e.message });
      continue;
    }
    await migrateFolderRecursive(sub, targetSub, srcAccId, dstAccId, mode, progress);
  }
}

async function migrateFolderTree(sourceFolderId, destFolderId, mode="move") {
  mig.running = true; mig.cancel = false;
  if (!_folderCache) await buildFolderTree();
  const srcFolder = _folderCache[sourceFolderId];
  const dstFolder = _folderCache[destFolderId];
  if (!srcFolder) throw new Error("Dossier source introuvable.");
  if (!dstFolder) throw new Error("Dossier destination introuvable.");
  const srcAccId = _accountIdCache[sourceFolderId];
  const dstAccId = _accountIdCache[destFolderId];
  const progress = { done:0, total:0, skipped:0, errors:[] };
  broadcast({ type:"MIG_PROGRESS", done:0, total:0, skipped:0, errors:[] });
  await migrateFolderRecursive(srcFolder, dstFolder, srcAccId, dstAccId, mode, progress);
  mig.running = false;
  await messenger.storage.local.remove(CFG.MIG_STATE_KEY);
  return { done: progress.done, total: progress.total, skipped: progress.skipped,
           errors: progress.errors, status: mig.cancel ? "cancelled" : "done" };
}

// ─────────────────────────────────────────────────────────────
// MODULE SYNCHRONISATION (compare deux boîtes IMAP)
// ─────────────────────────────────────────────────────────────

/**
 * ÉTAPE 1 — Analyse comparative des deux boîtes.
 * Ne modifie rien. Retourne un état des lieux par catégorie.
 */
async function analyseBoxes(srcAccountId, dstAccountId) {
  const accounts = await messenger.accounts.list();
  const srcAcc = accounts.find(a => a.id === srcAccountId);
  const dstAcc = accounts.find(a => a.id === dstAccountId);
  if (!srcAcc) throw new Error("Compte source introuvable.");
  if (!dstAcc) throw new Error("Compte destination introuvable.");

  const mapping = await loadMapping();

  // Scan source — messages avec étiquettes
  broadcast({ type:"SYNC_PROGRESS", phase:"scan_src", label:"Analyse de la boîte source…" });
  const srcTagged = new Map(); // headerMessageId → { tags }

  async function scanSrc(folders) {
    for (const folder of folders) {
      try {
        let page = await messenger.messages.list(folder.id ?? folder);
        do {
          for (const m of page.messages) {
            if (m.tags?.length && m.headerMessageId)
              srcTagged.set(m.headerMessageId, {
                tags   : m.tags,
                subject: m.subject,
                sender : m.author,
                date   : m.date,
              });
          }
          page = page.id ? await messenger.messages.continueList(page.id) : null;
        } while (page);
      } catch {}
      if (folder.subFolders?.length) await scanSrc(folder.subFolders);
    }
  }
  const srcFolders = srcAcc.folders ?? await messenger.folders.getSubFolders(srcAcc.rootFolder.id, false);
  await scanSrc(srcFolders);
  broadcast({ type:"SYNC_PROGRESS", phase:"scan_src_done",
    label: `Source : ${srcTagged.size} message(s) avec étiquettes trouvés` });

  // Scan destination — index par Message-ID
  broadcast({ type:"SYNC_PROGRESS", phase:"scan_dst", label:"Analyse de la boîte destination…" });
  const dstIndex = new Map(); // headerMessageId → { id, tags }

  async function scanDst(folders) {
    for (const folder of folders) {
      try {
        let page = await messenger.messages.list(folder.id ?? folder);
        do {
          for (const m of page.messages) {
            if (m.headerMessageId)
              dstIndex.set(m.headerMessageId, { id: m.id, tags: m.tags ?? [] });
          }
          page = page.id ? await messenger.messages.continueList(page.id) : null;
        } while (page);
      } catch {}
      if (folder.subFolders?.length) await scanDst(folder.subFolders);
    }
  }
  const dstFolders = dstAcc.folders ?? await messenger.folders.getSubFolders(dstAcc.rootFolder.id, false);
  await scanDst(dstFolders);
  broadcast({ type:"SYNC_PROGRESS", phase:"scan_dst_done",
    label: `Destination : ${dstIndex.size} messages indexés` });

  // Index tags TB pour retrouver les couleurs
  const tbTags = await messenger.messages.tags.list();
  const tbTagIndex = {}; // key → { tag, color }
  for (const t of tbTags) tbTagIndex[t.key] = t;

  // Construire l'état des lieux par catégorie Outlook
  const byCategory  = {};
  const noMapping   = [];
  let notFoundTotal = 0;

  for (const [msgId, srcInfo] of srcTagged) {
    const dstMsg = dstIndex.get(msgId);

    for (const tagKey of srcInfo.tags) {
      const olCat = mapping[tagKey];

      if (!olCat || olCat === "__skip__") {
        if (!noMapping.find(n => n.key === tagKey))
          noMapping.push({ key: tagKey, name: tbTagIndex[tagKey]?.tag ?? tagKey });
        continue;
      }

      if (!byCategory[olCat]) byCategory[olCat] = {
        olCategory : olCat,
        tbTagKey   : tagKey,
        color      : tbTagIndex[tagKey]?.color ?? "#4caf50",
        messages   : [],
        notFound   : 0,
      };

      if (!dstMsg) {
        byCategory[olCat].notFound++;
        notFoundTotal++;
      } else {
        const alreadyHas = dstMsg.tags.some(t => t.toLowerCase() === olCat.toLowerCase());
        if (!alreadyHas)
          byCategory[olCat].messages.push({
            dstId  : dstMsg.id,
            subject: srcInfo.subject,
            sender : srcInfo.sender,
            date   : srcInfo.date,
          });
      }
    }
  }

  broadcast({ type:"SYNC_PROGRESS", phase:"analyse_done", label:"Analyse terminée" });

  return {
    categories: Object.values(byCategory),
    noMapping,
    notFoundTotal,
    srcTotal: srcTagged.size,
    dstTotal: dstIndex.size,
  };
}

/**
 * ÉTAPE 2 — Application des catégories sélectionnées.
 * Crée le tag TB si nécessaire, puis l'applique sur chaque message.
 */
async function applyCategories(selectedCategories) {
  let done = 0;
  const total  = selectedCategories.reduce((s, c) => s + c.messages.length, 0);
  const errors = [];

  // Index des tags TB existants : tag name (lowercase) → clé
  const existingTags = await messenger.messages.tags.list();
  const tagByName = {};
  const tagByKey  = new Set();
  for (const t of existingTags) {
    tagByName[t.tag.toLowerCase()] = t.key;
    tagByKey.add(t.key.toLowerCase());
  }

  // Résoudre la clé TB pour chaque catégorie
  for (const cat of selectedCategories) {
    const nameLow = cat.olCategory.toLowerCase();

    // 1. Le tag existe déjà par son nom exact → utiliser sa clé
    if (tagByName[nameLow]) {
      cat._resolvedKey = tagByName[nameLow];
      continue;
    }

    // 2. Créer le tag avec une clé normalisée (alphanumérique uniquement)
    const safeKey = nameLow
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // retirer accents
      .replace(/[^a-z0-9]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_|_$/g, "")
      .substring(0, 20) + "_" + Date.now().toString(36);

    try {
      await messenger.messages.tags.create(safeKey, cat.olCategory, cat.color ?? "#4caf50");
      cat._resolvedKey = safeKey;
      tagByName[nameLow] = safeKey;
      console.log(`[Sync] Tag créé : "${cat.olCategory}" → clé "${safeKey}"`);
    } catch(e) {
      console.error(`[Sync] Impossible de créer le tag "${cat.olCategory}" :`, e.message);
      errors.push({ reason: `Tag "${cat.olCategory}" non créé : ${e.message}` });
      cat._resolvedKey = null;
    }
  }

  // Appliquer les tags sur chaque message
  for (const cat of selectedCategories) {
    if (!cat._resolvedKey) continue; // tag non créé → skip

    for (const msg of cat.messages) {
      try {
        const current  = await messenger.messages.get(msg.dstId);
        const currTags = current.tags ?? [];
        const alreadyHas = currTags.some(t => t.toLowerCase() === cat._resolvedKey.toLowerCase());
        if (!alreadyHas) {
          await messenger.messages.update(msg.dstId, { tags: [...currTags, cat._resolvedKey] });
        }
        done++;
      } catch(e) {
        errors.push({ id: msg.dstId, reason: e.message });
        console.error(`[Sync] Erreur msg ${msg.dstId} :`, e.message);
      }
      broadcast({ type:"SYNC_APPLY_PROGRESS", done, total });
      if (done % 10 === 0) await sleep(20);
    }
  }

  await messenger.storage.local.remove(CFG.MIG_STATE_KEY);
  broadcast({ type:"SYNC_APPLY_DONE", done, total, errors });
  return { done, total, errors };
}


// ─────────────────────────────────────────────────────────────
// MODULE EXPORT (contacts + calendriers)
// ─────────────────────────────────────────────────────────────
async function listAddressBooks() {
  try {
    const books = await messenger.addressBooks.list(true);
    const result = [];
    for (const book of books) {
      // Ignorer les carnets collectés (bruit)
      if (book.name === "Adresses collectées" || book.name === "Collected Addresses") continue;
      result.push({
        id   : book.id,
        name : book.name,
        type : book.type ?? "local",
        count: (book.contacts ?? []).length,
      });
    }
    return result;
  } catch(e) {
    console.error("[Export] Erreur listAddressBooks:", e);
    return [];
  }
}

async function exportContactsVcf(bookIds) {
  let vcf = "";
  for (const bookId of bookIds) {
    try {
      const book = await messenger.addressBooks.get(bookId, true);
      for (const contact of (book.contacts ?? [])) {
        if (contact.vCard) vcf += contact.vCard.trim() + "\r\n\r\n";
      }
    } catch(e) {
      console.warn(`[Export] Carnet ${bookId}:`, e);
    }
  }
  return vcf;
}

// Calendriers : pas d'API WebExtension standard → guide + détection TbSync
async function detectCalendars() {
  // Tenter de détecter des calendriers via l'API expérimentale si disponible
  if (typeof messenger.calendar !== "undefined") {
    try {
      const cals = await messenger.calendar.calendars.query({});
      return { available: true, calendars: cals };
    } catch {}
  }
  // Pas d'API → retourner instructions manuelles
  return { available: false, calendars: [] };
}

/**
 * Scanne un compte IMAP pour détecter les catégories Outlook.
 *
 * Stratégie en deux passes :
 *  1. Liste rapide de tous les messages via messages.list() → collecte message.tags
 *     (tags TB enregistrés — peut être vide pour les catégories Outlook non mappées)
 *  2. Sur un échantillon de messages, lit le header "Keywords" via getFull()
 *     pour attraper les mots-clés IMAP bruts qu'Outlook pose
 *  3. Fusionne les deux résultats
 */
async function scanOutlookCategories(accountId) {
  const accounts = await messenger.accounts.list();
  const account  = accounts.find(a => a.id === accountId);
  if (!account) throw new Error("Compte introuvable.");

  // Flags système à ignorer
  const SYSTEM = new Set([
    "\\seen","\\answered","\\flagged","\\deleted","\\draft","\\recent",
    "$mdnsent","$forwarded","forwarded","junk","nonjunk","notjunk",
    "x-subject-tag",
  ]);
  const isTbBuiltin = k => /^\$label\d+$/.test(k);
  const isSystem    = k => SYSTEM.has(k) || isTbBuiltin(k) || k.startsWith("\\");

  const counter = {}; // clé lowercase → { count, displayName }

  function addKey(rawKey) {
    if (!rawKey?.trim()) return;
    const k = rawKey.trim().toLowerCase();
    if (isSystem(k)) return;
    if (!counter[k]) counter[k] = { count:0, displayName: rawKey.trim() };
    counter[k].count++;
  }

  // Récupérer l'arborescence complète via l'API folders
  async function getFolders(parent) {
    try {
      return await messenger.folders.getSubFolders(parent.id ?? parent, false);
    } catch { return []; }
  }

  // Passe 1 : collecte via message.tags (tags TB enregistrés)
  const allMessages = []; // garder les ids pour la passe 2
  const sampled     = new Set();

  async function scanFolder(folder) {
    try {
      let page = await messenger.messages.list(folder.id ?? folder);
      do {
        for (const m of page.messages) {
          // Passe 1 : tags TB
          for (const tag of (m.tags ?? [])) addKey(tag);
          // Garder pour échantillon passe 2
          if (allMessages.length < 500) allMessages.push(m.id);
        }
        page = page.id ? await messenger.messages.continueList(page.id) : null;
      } while (page);
    } catch(e) {
      console.warn(`[Scan] Dossier ${folder.name}:`, e.message);
    }
    // Récursion sous-dossiers
    const subs = await getFolders(folder);
    for (const sub of subs) await scanFolder(sub);
  }

  // Traverser tous les dossiers du compte
  const accFolders = account.folders ?? await messenger.folders.getSubFolders(account.rootFolder.id, false);
  for (const folder of accFolders) {
    await scanFolder(folder);
  }

  // Passe 2 : lire les headers Keywords/X-Keywords sur un échantillon
  // pour attraper les mots-clés IMAP bruts (catégories Outlook non enregistrées dans TB)
  const SAMPLE_SIZE = Math.min(200, allMessages.length);
  // Prendre des messages régulièrement espacés pour couvrir toute la boîte
  const step    = Math.max(1, Math.floor(allMessages.length / SAMPLE_SIZE));
  const sample  = allMessages.filter((_, i) => i % step === 0).slice(0, SAMPLE_SIZE);

  for (const msgId of sample) {
    try {
      const full = await messenger.messages.getFull(msgId);
      // Le header "keywords" contient les mots-clés IMAP space-séparés
      const kwHeader = full.headers?.keywords?.[0] ?? full.headers?.["x-keywords"]?.[0] ?? "";
      if (kwHeader) {
        for (const kw of kwHeader.split(/\s+/)) addKey(kw);
      }
      // Certains clients Outlook écrivent aussi dans x-microsoft-antispam ou x-ms-exchange-calendar-*
      // On les ignore volontairement
    } catch {}
  }

  const results = Object.entries(counter)
    .map(([, v]) => ({ key: v.displayName, displayName: v.displayName, count: v.count }))
    .sort((a, b) => b.count - a.count);

  console.log(`[Scan Outlook] ${results.length} catégorie(s) détectée(s) sur ${allMessages.length} messages (échantillon: ${sample.length})`);
  return results;
}


// ─────────────────────────────────────────────────────────────
// MODULE MICROSOFT GRAPH — Catégories Outlook
// ─────────────────────────────────────────────────────────────

// Client ID et Tenant ID intégrés — pas de secret (flux PKCE public)
const GRAPH_CLIENT_ID = "bcfabced-c4a2-4425-bb5d-46ef4d8c547c";
const GRAPH_TENANT_ID = "898a7ac2-f878-44ab-80f0-1e1852b7bebd";
const GRAPH_SCOPE     = "Mail.ReadWrite MailboxSettings.ReadWrite";

// Token en mémoire (session uniquement)
let _graphToken    = null;
let _graphTokenExp = 0;
let _deviceCodeCancel = false;

function isTokenValid() {
  return _graphToken && Date.now() < _graphTokenExp - 60000;
}

/**
 * Extraction du token depuis une URL de redirect (fragment #access_token=...).
 */
function parseTokenFromUrl(responseUrl) {
  const hash = responseUrl.split("#")[1] || "";
  const params = new URLSearchParams(hash);
  const accessToken = params.get("access_token");
  const expiresIn = params.get("expires_in");
  const error = params.get("error");
  const errorDesc = params.get("error_description");

  if (error) throw new Error(errorDesc || error);
  if (!accessToken) throw new Error("Token absent de la reponse");
  return { access_token: accessToken, expires_in: parseInt(expiresIn) || 3600 };
}

/**
 * VERSION A : launchWebAuthFlow + implicit flow.
 * Methode recommandee par Mozilla pour les extensions.
 * Le token revient directement dans l'URL — zero fetch, zero CORS.
 */
async function graphAuthenticateA() {
  const redirectUri = messenger.identity.getRedirectURL();
  const nonce = crypto.randomUUID();

  const authUrl =
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${GRAPH_CLIENT_ID}` +
    `&response_type=token` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&scope=${encodeURIComponent(GRAPH_SCOPE)}` +
    `&nonce=${nonce}` +
    `&prompt=select_account`;

  console.log("[Graph A] redirect_uri:", redirectUri);
  console.log("[Graph A] Lancement launchWebAuthFlow (implicit)...");

  const responseUrl = await messenger.identity.launchWebAuthFlow({
    url: authUrl,
    interactive: true,
  });

  console.log("[Graph A] Callback recu");
  return parseTokenFromUrl(responseUrl);
}

/**
 * VERSION B : Onglet Thunderbird + implicit flow (fallback).
 * Ouvre l'auth dans un onglet TB, capture le redirect avec tabs.onUpdated.
 */
async function graphAuthenticateB() {
  const redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient";
  const nonce = crypto.randomUUID();

  const authUrl =
    `https://login.microsoftonline.com/${GRAPH_TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${GRAPH_CLIENT_ID}` +
    `&response_type=token` +
    `&redirect_uri=${encodeURIComponent(redirectUri)}` +
    `&scope=${encodeURIComponent(GRAPH_SCOPE)}` +
    `&nonce=${nonce}` +
    `&prompt=select_account`;

  console.log("[Graph B] Ouverture onglet auth (implicit)...");

  const responseUrl = await new Promise((resolve, reject) => {
    let authTabId = null;

    const onUpdated = (tabId, changeInfo, tab) => {
      if (tabId !== authTabId) return;
      const tabUrl = changeInfo.url || tab.url || "";
      if (tabUrl.startsWith(redirectUri)) {
        messenger.tabs.onUpdated.removeListener(onUpdated);
        messenger.tabs.remove(tabId).catch(() => {});
        resolve(tabUrl);
      }
    };

    messenger.tabs.onUpdated.addListener(onUpdated);
    messenger.tabs.create({ url: authUrl, active: true }).then(tab => {
      authTabId = tab.id;
    }).catch(e => {
      messenger.tabs.onUpdated.removeListener(onUpdated);
      reject(e);
    });

    setTimeout(() => {
      messenger.tabs.onUpdated.removeListener(onUpdated);
      if (authTabId) messenger.tabs.remove(authTabId).catch(() => {});
      reject(new Error("Timeout — connexion non completee en 5 minutes."));
    }, 300000);
  });

  console.log("[Graph B] Callback recu");
  return parseTokenFromUrl(responseUrl);
}

/**
 * Authentification Graph : essaie A puis B en fallback.
 */
async function graphAuthenticate() {
  let tokenData;
  try {
    tokenData = await graphAuthenticateA();
  } catch(errA) {
    console.warn("[Graph] Methode A echouee:", errA.message, "— tentative methode B...");
    tokenData = await graphAuthenticateB();
  }

  _graphToken    = tokenData.access_token;
  _graphTokenExp = Date.now() + (tokenData.expires_in * 1000);
  console.log("[Graph] Authentifie, token valide", tokenData.expires_in, "s");
  return { ok: true, expires_in: tokenData.expires_in };
}



/**
 * Recherche un message dans Graph par son Internet Message-ID (RFC 2822).
 * Retourne l'ID Graph interne du message.
 */
async function findGraphMessageId(internetMessageId) {
  if (!isTokenValid()) throw new Error("Token Graph expiré — reconnectez-vous.");

  // Nettoyer le Message-ID (retirer les <>)
  const cleanId = internetMessageId.replace(/^<|>$/g, "");

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages` +
    `?$filter=internetMessageId eq '${encodeURIComponent("<" + cleanId + ">")}'` +
    `&$select=id,subject,internetMessageId,categories` +
    `&$top=1`,
    { headers: { "Authorization": `Bearer ${_graphToken}` } }
  );

  if (!resp.ok) {
    const err = await resp.json();
    throw new Error("Graph query failed: " + (err.error?.message ?? resp.status));
  }

  const data = await resp.json();
  return data.value?.[0] ?? null;
}

/**
 * Applique des catégories Outlook sur un message via Graph.
 * @param {string} graphMsgId - ID Graph interne du message
 * @param {string[]} categories - noms des catégories à ajouter
 */
async function applyGraphCategories(graphMsgId, categories) {
  if (!isTokenValid()) throw new Error("Token Graph expiré — reconnectez-vous.");

  const resp = await fetch(
    `https://graph.microsoft.com/v1.0/me/messages/${graphMsgId}`,
    {
      method : "PATCH",
      headers: {
        "Authorization": `Bearer ${_graphToken}`,
        "Content-Type" : "application/json",
      },
      body: JSON.stringify({ categories }),
    }
  );

  if (!resp.ok) {
    const err = await resp.json();
    throw new Error("Graph PATCH failed: " + (err.error?.message ?? resp.status));
  }
  return true;
}

/**
 * Synchronisation complète via Graph :
 * Pour chaque message sélectionné, trouve son ID Graph et applique les catégories.
 */
async function applyCategoriesViaGraph(selectedCategories) {
  let done = 0, skipped = 0;
  const total  = selectedCategories.reduce((s, c) => s + c.messages.length, 0);
  const errors = [];

  for (const cat of selectedCategories) {
    for (const msg of cat.messages) {
      try {
        // Retrouver le message TB pour obtenir son Internet Message-ID
        const tbMsg = await messenger.messages.get(msg.dstId);
        if (!tbMsg?.headerMessageId) { skipped++; continue; }

        // Chercher dans Graph par Internet Message-ID
        const graphMsg = await findGraphMessageId(tbMsg.headerMessageId);
        if (!graphMsg) {
          skipped++;
          console.warn(`[Graph] Message non trouvé: ${tbMsg.headerMessageId}`);
          continue;
        }

        // Fusionner avec les catégories existantes
        const existing   = graphMsg.categories ?? [];
        const newCats    = [...new Set([...existing, cat.olCategory])];
        await applyGraphCategories(graphMsg.id, newCats);
        done++;

      } catch(e) {
        errors.push({ subject: msg.subject, reason: e.message });
        console.error(`[Graph] Erreur:`, e.message);
      }
      broadcast({ type:"GRAPH_APPLY_PROGRESS", done, total, skipped });
      if (done % 5 === 0) await sleep(100); // respecter le throttling Graph
    }
  }

  return { done, total, skipped, errors };
}

// ─────────────────────────────────────────────────────────────
// MODULE GRAPH — Catégories Outlook (masterCategories)
// ─────────────────────────────────────────────────────────────

// Couleurs Outlook : preset0..preset24 — mapping hex approché
const OL_PRESET_COLORS = [
  { preset: "preset0",  name: "Red",        hex: "#e7514c" },
  { preset: "preset1",  name: "Orange",     hex: "#f5a623" },
  { preset: "preset2",  name: "Brown",      hex: "#a0522d" },
  { preset: "preset3",  name: "Yellow",     hex: "#f7d64e" },
  { preset: "preset4",  name: "Green",      hex: "#4caf50" },
  { preset: "preset5",  name: "Teal",       hex: "#009688" },
  { preset: "preset6",  name: "Olive",      hex: "#808000" },
  { preset: "preset7",  name: "Blue",       hex: "#2196f3" },
  { preset: "preset8",  name: "Purple",     hex: "#9c27b0" },
  { preset: "preset9",  name: "Cranberry",  hex: "#c62828" },
  { preset: "preset10", name: "Steel",      hex: "#607d8b" },
  { preset: "preset11", name: "DarkSteel",  hex: "#37474f" },
  { preset: "preset12", name: "Gray",       hex: "#9e9e9e" },
  { preset: "preset13", name: "DarkGray",   hex: "#616161" },
  { preset: "preset14", name: "Black",      hex: "#212121" },
  { preset: "preset15", name: "DarkRed",    hex: "#b71c1c" },
  { preset: "preset16", name: "DarkOrange", hex: "#e65100" },
  { preset: "preset17", name: "DarkBrown",  hex: "#5d4037" },
  { preset: "preset18", name: "DarkYellow", hex: "#f9a825" },
  { preset: "preset19", name: "DarkGreen",  hex: "#2e7d32" },
  { preset: "preset20", name: "DarkTeal",   hex: "#00695c" },
  { preset: "preset21", name: "DarkOlive",  hex: "#556b2f" },
  { preset: "preset22", name: "DarkBlue",   hex: "#1565c0" },
  { preset: "preset23", name: "DarkPurple", hex: "#6a1b9a" },
  { preset: "preset24", name: "DarkCranberry", hex: "#880e4f" },
];

function hexToRgb(hex) {
  const m = hex.replace("#","").match(/.{2}/g);
  return m ? m.map(c => parseInt(c, 16)) : [128,128,128];
}

function closestPreset(hex) {
  const [r,g,b] = hexToRgb(hex);
  let best = "preset4", bestDist = Infinity;
  for (const p of OL_PRESET_COLORS) {
    const [pr,pg,pb] = hexToRgb(p.hex);
    const d = (r-pr)**2 + (g-pg)**2 + (b-pb)**2;
    if (d < bestDist) { bestDist = d; best = p.preset; }
  }
  return best;
}

async function listOutlookCategories() {
  if (!isTokenValid()) throw new Error("Token Graph expire — reconnectez-vous.");
  const resp = await fetch(
    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    { headers: { "Authorization": `Bearer ${_graphToken}` } }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Graph: " + (err.error?.message ?? resp.status));
  }
  const data = await resp.json();
  return data.value ?? [];
}

async function createOutlookCategory(displayName, color) {
  if (!isTokenValid()) throw new Error("Token Graph expire — reconnectez-vous.");
  const resp = await fetch(
    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${_graphToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ displayName, color }),
    }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Graph: " + (err.error?.message ?? resp.status));
  }
  return await resp.json();
}

async function autoCreateCategoriesFromTags() {
  // 1. Lire les tags TB
  const tbTags = await messenger.messages.tags.list();
  if (!tbTags.length) return { created: [], skipped: [], mapping: {} };

  // 2. Lire les catégories Outlook existantes
  const olCats = await listOutlookCategories();
  const olNames = new Set(olCats.map(c => c.displayName.toLowerCase()));

  // 3. Créer les manquantes + construire le mapping
  const created = [], skipped = [], mapping = {};
  const existingMapping = await loadMapping();

  for (const tag of tbTags) {
    const catName = tag.tag; // utiliser le nom du tag TB comme nom de catégorie
    mapping[tag.key] = catName;

    if (olNames.has(catName.toLowerCase())) {
      skipped.push({ key: tag.key, name: catName, reason: "existe deja" });
      continue;
    }

    try {
      const preset = closestPreset(tag.color || "#4caf50");
      await createOutlookCategory(catName, preset);
      created.push({ key: tag.key, name: catName, color: preset });
      olNames.add(catName.toLowerCase());
    } catch(e) {
      skipped.push({ key: tag.key, name: catName, reason: e.message });
    }
    await sleep(100);
  }

  // 4. Sauvegarder le mapping
  await saveMapping(mapping);

  return { created, skipped, mapping };
}

async function listTagsFn() { return await messenger.messages.tags.list(); }
async function renameTagFn(key,name,col) { await messenger.messages.tags.update(key, name, col); }
async function deleteTagFn(key)          { await messenger.messages.tags.delete(key); }
async function createTagFn(name, color) {
  const key = name.toLowerCase().replace(/[^a-z0-9]/g,"_").substring(0,20) + "_" + Date.now().toString(36);
  await messenger.messages.tags.create(key, name, color || "#4caf50");
  return key;
}

// ─────────────────────────────────────────────────────────────
// MODULE M365
// ─────────────────────────────────────────────────────────────
async function probeM365Domain(email) {
  const domain = email.split("@")[1]?.toLowerCase();
  if (!domain) throw new Error("Email invalide.");

  let isMicrosoft = false;
  try {
    const resp = await fetch(
      `https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=MX`,
      { headers: { "Accept":"application/json" } }
    );
    if (resp.ok) {
      const dns = await resp.json();
      isMicrosoft = (dns.Answer ?? []).some(r =>
        (r.data||"").toLowerCase().includes("mail.protection.outlook.com") ||
        (r.data||"").toLowerCase().includes("office365.com")
      );
    }
  } catch {}

  // Vérifier si déjà configuré
  const accounts = await messenger.accounts.list();
  const existing = accounts.find(acc =>
    acc.identities?.some(id => id.email?.toLowerCase() === email.toLowerCase())
  );

  return {
    email, domain, isMicrosoft,
    isPersonal: /^(outlook|hotmail|live|msn)\.(com|fr)/.test(domain),
    alreadyConfigured: !!existing,
    existingAccountName: existing?.name ?? null,
    imap: { server:"outlook.office365.com", port:993, security:"SSL/TLS", auth:"OAuth2" },
    smtp: { server:"smtp.office365.com",    port:587, security:"STARTTLS", auth:"OAuth2" },
  };
}

async function waitForNewAccount(email, timeoutMs=120000) {
  const start    = Date.now();
  const emailLow = email.toLowerCase();
  const before   = new Set((await messenger.accounts.list()).map(a => a.id));
  return new Promise(resolve => {
    const poll = async () => {
      if (Date.now() - start > timeoutMs) { resolve(null); return; }
      const accounts = await messenger.accounts.list();
      const newAcc = accounts.find(acc =>
        !before.has(acc.id) &&
        acc.identities?.some(id => id.email?.toLowerCase() === emailLow)
      );
      if (newAcc) { resolve(newAcc); return; }
      setTimeout(poll, 2000);
    };
    setTimeout(poll, 2000);
  });
}

// ─────────────────────────────────────────────────────────────
// MENU CONTEXTUEL
// ─────────────────────────────────────────────────────────────
function createContextMenu() {
  messenger.menus.remove("cen-subject-tag").catch(() => {});
  messenger.menus.create({
    id: "cen-subject-tag",
    title: "🌿 Mail-CEN — Ajouter les tags au sujet {Tag}",
    contexts: ["message_list"],
  }, () => {
    if (!messenger.runtime.lastError) messenger.menus.refresh().catch(() => {});
  });
}
messenger.runtime.onInstalled.addListener(createContextMenu);
messenger.runtime.onStartup.addListener(createContextMenu);
createContextMenu();

messenger.menus.onClicked.addListener(async (info) => {
  if (info.menuItemId !== "cen-subject-tag") return;
  const messages = info.selectedMessages?.messages;
  if (!messages?.length) return;
  await messenger.notifications.create("cen-proc", {
    type:"basic", title:"Mail-CEN",
    message: `Traitement de ${messages.length} message(s)…`
  });
  const result = await runSubjectTagOnIds(messages.map(m => m.id));
  await messenger.notifications.create("cen-done", {
    type:"basic", title:"Mail-CEN",
    message: `✓ ${result.ok} traité(s), ${result.skip} ignoré(s)${result.err ? `, ${result.err} erreur(s)` : ""}`
  });
});

// ─────────────────────────────────────────────────────────────
// BUS DE MESSAGES
// ─────────────────────────────────────────────────────────────
messenger.runtime.onMessage.addListener(async (req) => {
  try {
    switch (req.action) {

      // ── M365
      case "probeM365":        return probeM365Domain(req.email);
      case "openAccountSetup":
        // TB 140+ ne permet plus d'ouvrir about:accountsetup depuis une extension
        // On retourne les instructions pour l'utilisateur
        return { ok:true, manual: true };
      case "waitForNewAccount":
        waitForNewAccount(req.email, 120000).then(acc =>
          broadcast({ type:"M365_ACCOUNT_DETECTED", account:acc }));
        return { started:true };

      // ── Étiquettes / Mapping
      case "getTbTagsWithMapping": {
        const tags = await getAllTbTags();
        const mapping = await loadMapping();
        return tags.map(t => ({ ...t, olCategory: mapping[t.key] ?? "" }));
      }
      case "saveMapping":  await saveMapping(req.mapping); return { ok:true };
      case "loadMapping":  return loadMapping();

      // ── Migration
      case "getFolderTree":
        _folderCache = null; _accountIdCache = {};
        return buildFolderTree();
      case "startMigrationTree": {
        if (mig.running) return { error:"Une migration est déjà en cours." };
        migrateFolderTree(req.source, req.dest, req.mode ?? "move")
          .then(r  => broadcast({ type:"MIG_DONE", ...r }))
          .catch(e => { broadcast({ type:"MIG_ERROR", error:e.message }); mig.running=false; });
        return { started:true };
      }
      case "cancelMigration": mig.cancel=true; return { ok:true };
      case "getMigState": {
        const s = await messenger.storage.local.get(CFG.MIG_STATE_KEY);
        return s[CFG.MIG_STATE_KEY] ?? null;
      }
      case "clearMigState":
        await messenger.storage.local.remove(CFG.MIG_STATE_KEY); return { ok:true };

      // ── Synchronisation
      case "analyseBoxes": {
        if (mig.running) return { error:"Une opération est déjà en cours." };
        mig.running = true;
        analyseBoxes(req.srcAccountId, req.dstAccountId)
          .then(r  => { mig.running=false; broadcast({ type:"SYNC_ANALYSE_DONE", ...r }); })
          .catch(e => { mig.running=false; broadcast({ type:"SYNC_ERROR", error:e.message }); });
        return { started:true };
      }

      case "applyCategories": {
        if (mig.running) return { error:"Une opération est déjà en cours." };
        mig.running = true;
        applyCategories(req.categories)
          .then(() => { mig.running = false; })
          .catch(e  => {
            mig.running = false;
            broadcast({ type:"SYNC_ERROR", error: e.message });
          });
        return { started: true };
      }

      // ── Export
      case "listAddressBooks":  return listAddressBooks();
      case "exportContacts":    return { vcf: await exportContactsVcf(req.bookIds) };
      case "detectCalendars":   return detectCalendars();

      case "scanOutlookCategories":
        return scanOutlookCategories(req.accountId);

      // ── Microsoft Graph ────────────────────────────────────
      case "graphAuthenticate":
        // Lancer en async pour survivre à la fermeture du popup
        graphAuthenticate()
          .then(r  => broadcast({ type:"GRAPH_AUTH_OK", expires_in: r.expires_in }))
          .catch(e => broadcast({ type:"GRAPH_AUTH_ERROR", error: e.message }));
        return { started: true };

      case "graphIsAuthenticated":
        return { authenticated: isTokenValid() };

      case "graphCancelAuth":
        _deviceCodeCancel = true;
        return { ok: true };

      case "listOutlookCategories":
        if (!isTokenValid()) return { error: "Non authentifie — connectez-vous d'abord." };
        return { categories: await listOutlookCategories() };

      case "autoCreateCategories":
        if (!isTokenValid()) return { error: "Non authentifie — connectez-vous d'abord." };
        return autoCreateCategoriesFromTags();

      case "applyCategoriesViaGraph": {
        if (!isTokenValid()) return { error: "Non authentifié — connectez-vous d'abord." };
        if (mig.running) return { error: "Une opération est déjà en cours." };
        mig.running = true;
        applyCategoriesViaGraph(req.categories)
          .then(r  => { mig.running=false; broadcast({ type:"GRAPH_APPLY_DONE", ...r }); })
          .catch(e => { mig.running=false; broadcast({ type:"GRAPH_ERROR", error:e.message }); });
        return { started: true };
      }

      // ── Tags TB
      case "getAccounts": {
        const accs = await messenger.accounts.list();
        return accs.map(a => ({ id:a.id, name:a.name, type:a.type }));
      }
      case "listTags":   return listTagsFn();
      case "renameTag":  await renameTagFn(req.key, req.name, req.color); return { ok:true };
      case "deleteTag":  await deleteTagFn(req.key); return { ok:true };
      case "createTag":  return { ok:true, key: await createTagFn(req.name, req.color) };

      // ── Subject-Tag
      case "processAll":      return runSubjectTagAll();
      case "processSelected": return runSubjectTagOnIds(req.ids);

      default: return false;
    }
  } catch(e) {
    console.error("[Mail-CEN] Erreur:", req.action, e);
    return { error: e.message };
  }
});

console.log("[Mail-CEN] Prêt v5.3");
