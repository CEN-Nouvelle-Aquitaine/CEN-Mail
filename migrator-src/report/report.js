"use strict";

const esc = s => (s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function fmt(n) { return (n ?? 0).toLocaleString("fr-FR"); }

function durStr(ms) {
  if (!ms) return "—";
  const s = Math.round(ms / 1000);
  if (s < 60) return `${s} s`;
  return `${Math.floor(s / 60)} min ${s % 60} s`;
}

async function load() {
  const data = await browser.storage.local.get("mig_report");
  const r = data.mig_report;

  if (!r) {
    document.getElementById("meta").textContent = "Aucun rapport disponible — lancez d'abord une migration.";
    return;
  }

  // ── En-tête
  const date = new Date(r.ts).toLocaleString("fr-FR", {
    day: "2-digit", month: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit",
  });
  const statusLabel = r.status === "cancelled" ? "⚠ Annulée" : "✅ Terminée";
  document.getElementById("meta").innerHTML =
    `${statusLabel} &nbsp;·&nbsp; ${date} &nbsp;·&nbsp; Profil : <b>${esc(r.profile)}</b> &nbsp;·&nbsp; Durée : <b>${durStr(r.duration)}</b>`;

  // ── Cartes résumé
  const t = r.totals;
  document.getElementById("cards").innerHTML = `
    <div class="card c-ok">  <div class="val">${fmt(t.copied)}</div>    <div class="lbl">Copiés</div></div>
    <div class="card c-dupe"><div class="val">${fmt(t.duplicates)}</div> <div class="lbl">Doublons ignorés</div></div>
    <div class="card c-err"> <div class="val">${fmt(t.errors)}</div>     <div class="lbl">Erreurs</div></div>
    <div class="card c-info"><div class="val">${fmt(r.folders.length)}</div><div class="lbl">Dossiers</div></div>
  `;

  // ── Tableau dossiers
  const tbody = document.getElementById("tbody");
  if (!r.folders.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="empty">Aucun dossier traité.</td></tr>`;
  } else {
    for (const f of r.folders) {
      const allOk   = f.errors === 0 && f.copied === f.toCopy;
      const hasErr  = f.errors > 0;
      const allFail = f.errors > 0 && f.copied === 0 && f.toCopy > 0;
      const stClass = allFail ? "st-err" : hasErr ? "st-warn" : "st-ok";
      const stBadge = allFail
        ? `<span class="badge-err">❌ Échec</span>`
        : hasErr
          ? `<span class="badge-warn">⚠ Partiel</span>`
          : `<span class="badge-ok">✅ OK</span>`;
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${esc(f.name)}</td>
        <td class="num">${fmt(f.srcCount)}</td>
        <td class="num">${fmt(f.toCopy + f.duplicatesKnown)}</td>
        <td class="num">${fmt(f.copied)}</td>
        <td class="num">${fmt(f.duplicates)}</td>
        <td class="num">${fmt(f.errors)}</td>
        <td>${stBadge}</td>
      `;
      tbody.appendChild(tr);
    }
  }

  // ── Liste erreurs
  const errs = r.errors || [];
  if (errs.length > 0) {
    document.getElementById("err-section").style.display = "";
    const block = document.getElementById("err-block");
    for (const e of errs) {
      const div = document.createElement("div");
      div.className = "err-row";
      div.innerHTML = `<span class="err-subj" title="${esc(e.subject)}">${esc(e.subject)}</span><span class="err-reason" title="${esc(e.reason)}">— ${esc(e.reason)}</span>`;
      block.appendChild(div);
    }
  }
}

document.addEventListener("DOMContentLoaded", load);
