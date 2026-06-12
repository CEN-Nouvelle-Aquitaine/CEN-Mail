/**
 * Mail-CAT CEN — background.js v1.0.0
 * Modules : Mapping étiquettes TB ↔ catégories Outlook · Microsoft Graph · Synchronisation · Tags TB
 */
"use strict";
console.log("[Mail-CAT CEN] Chargé v1.0.0");

// Keepalive — empêche TB (MV3 Limited Event Page) de suspendre le background
messenger.alarms.onAlarm.addListener(alarm => {
  if (alarm.name === "mailcat-keepalive") console.log("[Mail-CAT] keepalive");
});
function startKeepalive() { messenger.alarms.create("mailcat-keepalive", { periodInMinutes: 0.4 }); }
function stopKeepalive()  { messenger.alarms.clear("mailcat-keepalive"); }

// ─────────────────────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────────────────────
const CFG = {
  BATCH_SIZE   : 5,
  BATCH_DELAY  : 1500,
  MSG_DELAY    : 200,
  RETRY_MAX    : 4,
  RETRY_BACKOFF: 2000,
  HEALTH_THRESHOLD : 5,
  HEALTH_COOLDOWN  : 30000,
  HEALTH_DELAY_MULT: 3,
  MAPPING_KEY  : "cen_label_mapping",
  STATE_KEY    : "cat_state",
};

const TRANSIENT_PATTERNS = [
  "aborted","timeout","connection","network","interrupt","offline","busy","locked",
  "imap is busy","status: 2153054241","status: 0x80550021",
  "status: 2153054209","status: 2147500037",
];
const PERMANENT_PATTERNS = [
  "already contains","no such folder","permission denied","quota","message-id is null",
];

function classifyError(e) {
  const msg = (e?.message || "").toLowerCase();
  if (PERMANENT_PATTERNS.some(p => msg.includes(p))) return "permanent";
  if (TRANSIENT_PATTERNS.some(p => msg.includes(p))) return "transient";
  return "unknown";
}

const health = { consecutiveErrors: 0, degraded: false, totalErrors: 0, lastErrorAt: 0 };
function noteSuccess() { health.consecutiveErrors = 0; }

async function noteError(e) {
  health.consecutiveErrors++;
  health.totalErrors++;
  health.lastErrorAt = Date.now();
  if (health.consecutiveErrors >= CFG.HEALTH_THRESHOLD && !health.degraded) {
    health.degraded = true;
    broadcast({ type: "CAT_HEALTH", degraded: true, consecutive: health.consecutiveErrors });
    await cancellableSleep(CFG.HEALTH_COOLDOWN);
  }
}

function getDelayMult() { return health.degraded ? CFG.HEALTH_DELAY_MULT : 1; }

// ─────────────────────────────────────────────────────────────
// UTILS
// ─────────────────────────────────────────────────────────────
const sleep = ms => new Promise(r => setTimeout(r, ms));

const op = { running: false, cancel: false };

async function cancellableSleep(ms) {
  const step = 200; let elapsed = 0;
  while (elapsed < ms) {
    if (op.cancel) return;
    const wait = Math.min(step, ms - elapsed);
    await new Promise(r => setTimeout(r, wait));
    elapsed += wait;
  }
}

async function withRetry(fn, label = "op") {
  let lastErr;
  for (let attempt = 1; attempt <= CFG.RETRY_MAX; attempt++) {
    if (op.cancel) throw new Error("Annulé par l'utilisateur");
    try {
      const r = await fn();
      if (attempt > 1) console.log(`[Mail-CAT] ${label} OK après ${attempt} essais`);
      return r;
    } catch(e) {
      lastErr = e;
      const cls = classifyError(e);
      if (cls === "permanent") throw e;
      if (attempt === CFG.RETRY_MAX) throw e;
      const delay = CFG.RETRY_BACKOFF * attempt * getDelayMult();
      console.warn(`[Mail-CAT] ${label} échec ${cls} (essai ${attempt}/${CFG.RETRY_MAX}), retry dans ${delay}ms`);
      await cancellableSleep(delay);
      if (op.cancel) throw new Error("Annulé par l'utilisateur");
    }
  }
  throw lastErr;
}

async function collectMessages(asyncList) {
  const messages = [];
  if (asyncList[Symbol.asyncIterator]) {
    for await (const msg of asyncList) messages.push(msg);
    return messages;
  }
  let page = asyncList;
  do {
    messages.push(...page.messages);
    page = page.id ? await messenger.messages.continueList(page.id) : null;
  } while (page?.messages?.length);
  return messages;
}

const PERSISTED_TYPES = new Set([
  "SYNC_PROGRESS","SYNC_ANALYSE_DONE","SYNC_APPLY_PROGRESS","SYNC_APPLY_DONE","SYNC_ERROR",
  "GRAPH_APPLY_PROGRESS","GRAPH_APPLY_DONE","GRAPH_ERROR",
]);

function broadcast(msg) {
  messenger.runtime.sendMessage(msg).catch(() => {});
  if (PERSISTED_TYPES.has(msg.type))
    messenger.storage.local.set({ [CFG.STATE_KEY]: { ...msg, ts: Date.now() } });
}

// ─────────────────────────────────────────────────────────────
// MODULE MAPPING
// ─────────────────────────────────────────────────────────────
async function getAllTbTags() { return messenger.messages.tags.list(); }

async function loadMapping() {
  const d = await messenger.storage.local.get(CFG.MAPPING_KEY);
  return d[CFG.MAPPING_KEY] ?? {};
}

async function saveMapping(mapping) {
  await messenger.storage.local.set({ [CFG.MAPPING_KEY]: mapping });
}

// ─────────────────────────────────────────────────────────────
// MODULE TAGS TB
// ─────────────────────────────────────────────────────────────
async function listTagsFn()           { return messenger.messages.tags.list(); }
async function renameTagFn(key,name,col) { await messenger.messages.tags.update(key, name, col); }
async function deleteTagFn(key)          { await messenger.messages.tags.delete(key); }
async function createTagFn(name, color) {
  const key = name.toLowerCase().replace(/[^a-z0-9]/g,"_").substring(0,20) + "_" + Date.now().toString(36);
  await messenger.messages.tags.create(key, name, color || "#4caf50");
  return key;
}

// ─────────────────────────────────────────────────────────────
// MODULE MICROSOFT GRAPH — Authentification
// ─────────────────────────────────────────────────────────────
const GRAPH_CLIENT_ID = "bcfabced-c4a2-4425-bb5d-46ef4d8c547c";
const GRAPH_TENANT_ID = "898a7ac2-f878-44ab-80f0-1e1852b7bebd";
const GRAPH_SCOPE     = "Mail.ReadWrite MailboxSettings.ReadWrite";

let _graphToken    = null;
let _graphTokenExp = 0;

function isTokenValid() { return _graphToken && Date.now() < _graphTokenExp - 60000; }

function parseTokenFromUrl(responseUrl) {
  const hash   = responseUrl.split("#")[1] || "";
  const params = new URLSearchParams(hash);
  const accessToken = params.get("access_token");
  const expiresIn   = params.get("expires_in");
  const error       = params.get("error");
  const errorDesc   = params.get("error_description");
  if (error) throw new Error(errorDesc || error);
  if (!accessToken) throw new Error("Token absent de la réponse");
  return { access_token: accessToken, expires_in: parseInt(expiresIn) || 3600 };
}

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
  console.log("[Graph A] Lancement launchWebAuthFlow...");
  const responseUrl = await messenger.identity.launchWebAuthFlow({ url: authUrl, interactive: true });
  return parseTokenFromUrl(responseUrl);
}

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
  console.log("[Graph B] Ouverture onglet auth...");
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
      reject(new Error("Timeout — connexion non complétée en 5 minutes."));
    }, 300000);
  });
  return parseTokenFromUrl(responseUrl);
}

async function graphAuthenticate() {
  let tokenData;
  try {
    tokenData = await graphAuthenticateA();
  } catch(errA) {
    console.warn("[Graph] Méthode A échouée:", errA.message, "— tentative méthode B...");
    tokenData = await graphAuthenticateB();
  }
  _graphToken    = tokenData.access_token;
  _graphTokenExp = Date.now() + (tokenData.expires_in * 1000);
  console.log("[Graph] Authentifié, token valide", tokenData.expires_in, "s");
  return { ok: true, expires_in: tokenData.expires_in };
}

function graphFetch(url, opts = {}, timeoutMs = 15000) {
  const ctrl = new AbortController();
  const tid  = setTimeout(() => ctrl.abort(), timeoutMs);
  return fetch(url, { ...opts, signal: ctrl.signal }).finally(() => clearTimeout(tid));
}

// ─────────────────────────────────────────────────────────────
// MODULE GRAPH — Catégories Outlook (masterCategories)
// ─────────────────────────────────────────────────────────────
const OL_PRESET_COLORS = [
  { preset:"preset0",  hex:"#e7514c" },{ preset:"preset1",  hex:"#f5a623" },
  { preset:"preset2",  hex:"#a0522d" },{ preset:"preset3",  hex:"#f7d64e" },
  { preset:"preset4",  hex:"#4caf50" },{ preset:"preset5",  hex:"#009688" },
  { preset:"preset6",  hex:"#808000" },{ preset:"preset7",  hex:"#2196f3" },
  { preset:"preset8",  hex:"#9c27b0" },{ preset:"preset9",  hex:"#c62828" },
  { preset:"preset10", hex:"#607d8b" },{ preset:"preset11", hex:"#37474f" },
  { preset:"preset12", hex:"#9e9e9e" },{ preset:"preset13", hex:"#616161" },
  { preset:"preset14", hex:"#212121" },{ preset:"preset15", hex:"#b71c1c" },
  { preset:"preset16", hex:"#e65100" },{ preset:"preset17", hex:"#5d4037" },
  { preset:"preset18", hex:"#f9a825" },{ preset:"preset19", hex:"#2e7d32" },
  { preset:"preset20", hex:"#00695c" },{ preset:"preset21", hex:"#556b2f" },
  { preset:"preset22", hex:"#1565c0" },{ preset:"preset23", hex:"#6a1b9a" },
  { preset:"preset24", hex:"#880e4f" },
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
  if (!isTokenValid()) throw new Error("Token Graph expiré — reconnectez-vous.");
  const resp = await graphFetch(
    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    { headers: { "Authorization": `Bearer ${_graphToken}` } }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Graph: " + (err.error?.message ?? resp.status));
  }
  return (await resp.json()).value ?? [];
}

async function createOutlookCategory(displayName, color) {
  if (!isTokenValid()) throw new Error("Token Graph expiré — reconnectez-vous.");
  const resp = await graphFetch(
    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    {
      method: "POST",
      headers: { "Authorization": `Bearer ${_graphToken}`, "Content-Type": "application/json" },
      body: JSON.stringify({ displayName, color }),
    }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Graph: " + (err.error?.message ?? resp.status));
  }
  return resp.json();
}

async function autoCreateCategoriesFromTags() {
  const tbTags  = await messenger.messages.tags.list();
  if (!tbTags.length) return { created: [], skipped: [], mapping: {} };

  const olCats = await listOutlookCategories();
  const olNames = new Set(olCats.map(c => c.displayName.toLowerCase()));

  const created = [], skipped = [], mapping = {};

  for (const tag of tbTags) {
    const catName = tag.tag;
    mapping[tag.key] = catName;

    if (olNames.has(catName.toLowerCase())) {
      skipped.push({ key: tag.key, name: catName, reason: "existe déjà" });
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

  await saveMapping(mapping);
  return { created, skipped, mapping };
}

// ─────────────────────────────────────────────────────────────
// MODULE GRAPH — Application des catégories sur les messages
// ─────────────────────────────────────────────────────────────
async function findGraphMessageId(internetMessageId) {
  if (!isTokenValid()) throw new Error("Token Graph expiré — reconnectez-vous.");
  const raw          = internetMessageId.trim();
  const withBrackets = raw.startsWith("<") ? raw : `<${raw}>`;
  const params = new URLSearchParams({
    "$filter": `internetMessageId eq '${withBrackets}'`,
    "$select": "id,subject,internetMessageId,categories",
    "$top"   : "1",
  });
  const resp = await graphFetch(
    `https://graph.microsoft.com/v1.0/me/messages?${params}`,
    { headers: { "Authorization": `Bearer ${_graphToken}` } }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Graph query failed: " + (err.error?.message ?? resp.status));
  }
  return (await resp.json()).value?.[0] ?? null;
}

async function applyGraphCategories(graphMsgId, categories) {
  if (!isTokenValid()) throw new Error("Token Graph expiré — reconnectez-vous.");
  const resp = await graphFetch(
    `https://graph.microsoft.com/v1.0/me/messages/${graphMsgId}`,
    {
      method : "PATCH",
      headers: { "Authorization": `Bearer ${_graphToken}`, "Content-Type": "application/json" },
      body   : JSON.stringify({ categories }),
    }
  );
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error("Graph PATCH failed: " + (err.error?.message ?? resp.status));
  }
  return true;
}

async function applyCategoriesViaGraph(selectedCategories) {
  let done = 0, skipped = 0;
  const total  = selectedCategories.reduce((s, c) => s + c.messages.length, 0);
  const errors = [];

  for (const cat of selectedCategories) {
    if (op.cancel) break;
    for (const msg of cat.messages) {
      if (op.cancel) break;
      try {
        const tbMsg = await messenger.messages.get(msg.dstId);
        if (!tbMsg?.headerMessageId) { skipped++; continue; }

        const graphMsg = await findGraphMessageId(tbMsg.headerMessageId);
        if (!graphMsg) { skipped++; continue; }

        const existing = graphMsg.categories ?? [];
        await applyGraphCategories(graphMsg.id, [...new Set([...existing, cat.olCategory])]);
        done++;
      } catch(e) {
        errors.push({ subject: msg.subject, reason: e.message });
        console.error(`[Graph] Erreur:`, e.message);
      }
      broadcast({ type: "GRAPH_APPLY_PROGRESS", done, total, skipped });
      if (done % 5 === 0) await cancellableSleep(100);
    }
  }
  return { done, total, skipped, errors };
}

// ─────────────────────────────────────────────────────────────
// MODULE SCAN IMAP — Catégories Outlook détectées via IMAP
// ─────────────────────────────────────────────────────────────
async function scanOutlookCategories(accountId) {
  const accounts = await messenger.accounts.list();
  const account  = accounts.find(a => a.id === accountId);
  if (!account) throw new Error("Compte introuvable.");

  const SYSTEM    = new Set(["\\seen","\\answered","\\flagged","\\deleted","\\draft","\\recent",
    "$mdnsent","$forwarded","forwarded","junk","nonjunk","notjunk","x-subject-tag"]);
  const isTbBuiltin = k => /^\$label\d+$/.test(k);
  const isSystem    = k => SYSTEM.has(k) || isTbBuiltin(k) || k.startsWith("\\");

  const counter = {};
  function addKey(rawKey) {
    if (!rawKey?.trim()) return;
    const k = rawKey.trim().toLowerCase();
    if (isSystem(k)) return;
    if (!counter[k]) counter[k] = { count:0, displayName: rawKey.trim() };
    counter[k].count++;
  }

  const allMessages = [];

  async function scanFolder(folder) {
    try {
      const msgs = await collectMessages(await messenger.messages.list(folder.id ?? folder));
      for (const m of msgs) {
        for (const tag of (m.tags ?? [])) addKey(tag);
        if (allMessages.length < 500) allMessages.push(m.id);
      }
    } catch(e) { console.warn(`[Scan] ${folder.name}:`, e.message); }
    let subs = folder.subFolders ?? [];
    if (!subs.length) {
      try { subs = await messenger.folders.getSubFolders(folder.id ?? folder, false); } catch {}
    }
    for (const sub of subs) await scanFolder(sub);
  }

  const accFolders = account.folders ?? await messenger.folders.getSubFolders(account.rootFolder.id, false);
  for (const folder of accFolders) await scanFolder(folder);

  const SAMPLE_SIZE = Math.min(200, allMessages.length);
  const step   = Math.max(1, Math.floor(allMessages.length / SAMPLE_SIZE));
  const sample = allMessages.filter((_, i) => i % step === 0).slice(0, SAMPLE_SIZE);

  for (const msgId of sample) {
    try {
      const full = await messenger.messages.getFull(msgId);
      const kwHeader = full.headers?.keywords?.[0] ?? full.headers?.["x-keywords"]?.[0] ?? "";
      if (kwHeader) for (const kw of kwHeader.split(/\s+/)) addKey(kw);
    } catch {}
  }

  return Object.entries(counter)
    .map(([, v]) => ({ key: v.displayName, displayName: v.displayName, count: v.count }))
    .sort((a, b) => b.count - a.count);
}

// ─────────────────────────────────────────────────────────────
// MODULE SYNCHRONISATION — Analyse + Application IMAP
// ─────────────────────────────────────────────────────────────
async function analyseBoxes(srcAccountId, dstAccountId) {
  const accounts = await messenger.accounts.list();
  const srcAcc = accounts.find(a => a.id === srcAccountId);
  const dstAcc = accounts.find(a => a.id === dstAccountId);
  if (!srcAcc) throw new Error("Compte source introuvable.");
  if (!dstAcc) throw new Error("Compte destination introuvable.");

  const mapping = await loadMapping();

  broadcast({ type:"SYNC_PROGRESS", phase:"scan_src", label:"Analyse de la boîte source…" });
  const srcTagged = new Map();

  async function scanSrc(folders) {
    for (const folder of folders) {
      try {
        const msgs = await collectMessages(await messenger.messages.list(folder.id ?? folder));
        for (const m of msgs) {
          if (m.tags?.length && m.headerMessageId)
            srcTagged.set(m.headerMessageId, { tags:m.tags, subject:m.subject, sender:m.author, date:m.date });
        }
      } catch(e) { console.debug(`[Sync.src] ${folder.name}:`, e.message); }
      let subs = folder.subFolders ?? [];
      if (!subs.length) { try { subs = await messenger.folders.getSubFolders(folder.id ?? folder, false); } catch {} }
      if (subs.length) await scanSrc(subs);
    }
  }
  await scanSrc(srcAcc.folders ?? await messenger.folders.getSubFolders(srcAcc.rootFolder.id, false));
  broadcast({ type:"SYNC_PROGRESS", phase:"scan_src_done", label:`Source : ${srcTagged.size} message(s) avec étiquettes` });

  broadcast({ type:"SYNC_PROGRESS", phase:"scan_dst", label:"Analyse de la boîte destination…" });
  const dstIndex = new Map();

  async function scanDst(folders) {
    for (const folder of folders) {
      try {
        const msgs = await collectMessages(await messenger.messages.list(folder.id ?? folder));
        for (const m of msgs) {
          if (m.headerMessageId) dstIndex.set(m.headerMessageId, { id:m.id, tags:m.tags ?? [] });
        }
      } catch(e) { console.debug(`[Sync.dst] ${folder.name}:`, e.message); }
      let subs = folder.subFolders ?? [];
      if (!subs.length) { try { subs = await messenger.folders.getSubFolders(folder.id ?? folder, false); } catch {} }
      if (subs.length) await scanDst(subs);
    }
  }
  await scanDst(dstAcc.folders ?? await messenger.folders.getSubFolders(dstAcc.rootFolder.id, false));
  broadcast({ type:"SYNC_PROGRESS", phase:"scan_dst_done", label:`Destination : ${dstIndex.size} messages indexés` });

  const tbTags = await messenger.messages.tags.list();
  const tbTagIndex = {};
  for (const t of tbTags) tbTagIndex[t.key] = t;

  const byCategory = {};
  const noMapping  = [];
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
        olCategory: olCat, tbTagKey: tagKey,
        color: tbTagIndex[tagKey]?.color ?? "#4caf50",
        messages: [], notFound: 0, notFoundMessages: [],
      };
      if (!dstMsg) {
        byCategory[olCat].notFound++;
        byCategory[olCat].notFoundMessages.push({ subject:srcInfo.subject, sender:srcInfo.sender, date:srcInfo.date });
        notFoundTotal++;
      } else {
        const alreadyHas = dstMsg.tags.some(t => t.toLowerCase() === olCat.toLowerCase());
        if (!alreadyHas)
          byCategory[olCat].messages.push({ dstId:dstMsg.id, subject:srcInfo.subject, sender:srcInfo.sender, date:srcInfo.date });
      }
    }
  }

  broadcast({ type:"SYNC_PROGRESS", phase:"analyse_done", label:"Analyse terminée" });
  return {
    categories: Object.values(byCategory),
    noMapping, notFoundTotal,
    srcTotal: srcTagged.size, dstTotal: dstIndex.size,
  };
}

async function applyCategories(selectedCategories) {
  let done = 0;
  const total  = selectedCategories.reduce((s, c) => s + c.messages.length, 0);
  const errors = [];

  const existingTags = await messenger.messages.tags.list();
  const tagByName = {};
  for (const t of existingTags) tagByName[t.tag.toLowerCase()] = t.key;

  for (const cat of selectedCategories) {
    const nameLow = cat.olCategory.toLowerCase();
    if (tagByName[nameLow]) {
      cat._resolvedKey = tagByName[nameLow];
      continue;
    }
    const safeKey = nameLow
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/[^a-z0-9]/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "")
      .substring(0, 20) + "_" + Date.now().toString(36);
    try {
      await messenger.messages.tags.create(safeKey, cat.olCategory, cat.color ?? "#4caf50");
      cat._resolvedKey = safeKey;
      tagByName[nameLow] = safeKey;
    } catch(e) {
      errors.push({ reason: `Tag "${cat.olCategory}" non créé : ${e.message}` });
      cat._resolvedKey = null;
    }
  }

  for (const cat of selectedCategories) {
    if (op.cancel) break;
    if (!cat._resolvedKey) continue;
    for (const msg of cat.messages) {
      if (op.cancel) break;
      try {
        const current  = await messenger.messages.get(msg.dstId);
        const currTags = current.tags ?? [];
        const alreadyHas = currTags.some(t => t.toLowerCase() === cat._resolvedKey.toLowerCase());
        if (!alreadyHas)
          await messenger.messages.update(msg.dstId, { tags: [...currTags, cat._resolvedKey] });
        done++;
      } catch(e) {
        errors.push({ id: msg.dstId, reason: e.message });
      }
      broadcast({ type:"SYNC_APPLY_PROGRESS", done, total });
      if (done % 10 === 0) await cancellableSleep(20);
    }
  }

  await messenger.storage.local.remove(CFG.STATE_KEY);
  broadcast({ type:"SYNC_APPLY_DONE", done, total, errors });
  return { done, total, errors };
}

// ─────────────────────────────────────────────────────────────
// BUS DE MESSAGES
// ─────────────────────────────────────────────────────────────
messenger.runtime.onMessage.addListener(async (req) => {
  try {
    switch (req.action) {

      // ── Tags TB
      case "listTags":   return listTagsFn();
      case "renameTag":  await renameTagFn(req.key, req.name, req.color); return { ok:true };
      case "deleteTag":  await deleteTagFn(req.key); return { ok:true };
      case "createTag":  return { ok:true, key: await createTagFn(req.name, req.color) };

      // ── Mapping
      case "getTbTagsWithMapping": {
        const tags    = await getAllTbTags();
        const mapping = await loadMapping();
        return tags.map(t => ({ ...t, olCategory: mapping[t.key] ?? "" }));
      }
      case "saveMapping":  await saveMapping(req.mapping); return { ok:true };
      case "loadMapping":  return loadMapping();

      // ── Comptes
      case "getAccounts": {
        const accs = await messenger.accounts.list();
        return accs.map(a => ({ id:a.id, name:a.name, type:a.type }));
      }

      // ── Microsoft Graph
      case "graphAuthenticate":
        graphAuthenticate()
          .then(r  => broadcast({ type:"GRAPH_AUTH_OK", expires_in: r.expires_in }))
          .catch(e => broadcast({ type:"GRAPH_AUTH_ERROR", error: e.message }));
        return { started: true };

      case "graphIsAuthenticated":
        return { authenticated: isTokenValid() };

      case "listOutlookCategories":
        if (!isTokenValid()) return { error: "Non authentifié — connectez-vous d'abord." };
        return { categories: await listOutlookCategories() };

      case "autoCreateCategories":
        if (!isTokenValid()) return { error: "Non authentifié — connectez-vous d'abord." };
        return autoCreateCategoriesFromTags();

      case "scanOutlookCategories":
        return scanOutlookCategories(req.accountId);

      // ── Synchronisation
      case "analyseBoxes": {
        if (op.running) return { error: "Une opération est déjà en cours." };
        op.running = true; op.cancel = false;
        startKeepalive();
        analyseBoxes(req.srcAccountId, req.dstAccountId)
          .then(r  => { op.running=false; stopKeepalive(); broadcast({ type:"SYNC_ANALYSE_DONE", ...r }); })
          .catch(e => { op.running=false; stopKeepalive(); broadcast({ type:"SYNC_ERROR", error:e.message }); });
        return { started: true };
      }

      case "applyCategories": {
        if (op.running) return { error: "Une opération est déjà en cours." };
        op.running = true; op.cancel = false;
        startKeepalive();
        applyCategories(req.categories)
          .then(r => {
            op.running=false; stopKeepalive();
            messenger.notifications.create("cat-sync-done", {
              type:"basic", title:"Mail-CAT CEN",
              message: `✅ ${r?.done ?? "?"} message(s) traités via IMAP.`,
            });
          })
          .catch(e => {
            op.running=false; stopKeepalive();
            broadcast({ type:"SYNC_ERROR", error: e.message });
          });
        return { started: true };
      }

      case "applyCategoriesViaGraph": {
        if (!isTokenValid()) return { error: "Non authentifié — connectez-vous d'abord." };
        if (op.running) return { error: "Une opération est déjà en cours." };
        op.running = true; op.cancel = false;
        startKeepalive();
        applyCategoriesViaGraph(req.categories)
          .then(r => {
            op.running=false; stopKeepalive();
            broadcast({ type:"GRAPH_APPLY_DONE", ...r });
            const skip = r.skipped ? ` · ${r.skipped} non trouvés` : "";
            messenger.notifications.create("cat-graph-done", {
              type:"basic", title:"Mail-CAT CEN",
              message: `✅ ${r.done}/${r.total} catégories appliquées${skip}.`,
            });
          })
          .catch(e => {
            op.running=false; stopKeepalive();
            broadcast({ type:"GRAPH_ERROR", error: e.message });
          });
        return { started: true };
      }

      case "cancelOp": op.cancel = true; return { ok: true };

      case "getState": {
        const s = await messenger.storage.local.get(CFG.STATE_KEY);
        return s[CFG.STATE_KEY] ?? null;
      }
      case "clearState":
        await messenger.storage.local.remove(CFG.STATE_KEY); return { ok:true };

      default: return { error: `Action inconnue : ${req.action}` };
    }
  } catch(e) {
    console.error("[Mail-CAT] Erreur bus:", req.action, e);
    return { error: e.message };
  }
});

console.log("[Mail-CAT CEN] Prêt v1.0.0");
