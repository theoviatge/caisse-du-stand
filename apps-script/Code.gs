/* Caisse du stand : tableau Google Sheets partagé entre les téléphones.
 *
 * À coller dans Extensions > Apps Script du tableau, puis Déployer > Nouveau déploiement >
 * Application Web, « Exécuter en tant que : moi », « Qui a accès : tout le monde ».
 *
 * Change le code secret ci-dessous : chaque téléphone doit avoir le même.
 */
const CODE = "a-changer";

const SYNC = "_sync";               // feuille technique : ne pas modifier à la main
const HEAD = ["seq", "key", "col", "id", "updatedAt", "deleted", "data", "device"];

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function syncSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SYNC);
  if (!sh) {
    sh = ss.insertSheet(SYNC);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function rows_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEAD.length).getValues() : [];
}

/* GET ?code=…&since=N : tout ce qui a changé depuis N */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.code !== CODE) return out_({ ok: false, error: "code secret incorrect" });
  const since = Number(p.since) || 0;
  const rows = rows_(syncSheet_());
  let max = since;
  const records = [];
  rows.forEach(r => {
    const seq = Number(r[0]) || 0;
    if (seq > max) max = seq;
    if (seq <= since) return;
    let data = null;
    try { data = r[6] ? JSON.parse(r[6]) : null; } catch (err) {}
    records.push({ key: r[1], col: r[2], id: String(r[3]), updatedAt: Number(r[4]) || 0, deleted: r[5] === true || r[5] === "TRUE", data: data });
  });
  return out_({ ok: true, records: records, max: max });
}

/* POST {code, device, records:[{key,col,id,updatedAt,deleted,data}]} */
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "requête illisible" }); }
  if (body.code !== CODE) return out_({ ok: false, error: "code secret incorrect" });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = syncSheet_();
    const rows = rows_(sh);
    const at = {};
    let seq = 0;
    rows.forEach((r, i) => { at[r[1]] = i; seq = Math.max(seq, Number(r[0]) || 0); });
    const add = [];
    (body.records || []).forEach(x => {
      if (!x || !x.key) return;
      const line = [0, x.key, x.col, String(x.id), Number(x.updatedAt) || 0, !!x.deleted, x.deleted ? "" : JSON.stringify(x.data), body.device || ""];
      const i = at[x.key];
      if (i !== undefined) {
        if ((Number(rows[i][4]) || 0) > line[4]) return;   // le tableau a déjà plus récent
        line[0] = ++seq; rows[i] = line;
      } else {
        line[0] = ++seq; at[x.key] = rows.length + add.length; add.push(line);
      }
    });
    const all = rows.concat(add);
    if (all.length) sh.getRange(2, 1, all.length, HEAD.length).setValues(all);
    SpreadsheetApp.flush();
    try { rebuild_(all); } catch (err) { console.error(err); }
    return out_({ ok: true, max: seq });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- feuilles lisibles, refaites à chaque envoi ---------- */
function eur_(c) { return (Number(c) || 0) / 100; }
function kg_(g) { return g ? (Number(g) || 0) / 1000 : ""; }

function write_(name, head, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  if (data.length) sh.getRange(2, 1, data.length, head.length).setValues(data);
  sh.setFrozenRows(1);
}

function rebuild_(all) {
  const by = { products: [], tickets: [], retours: [] };
  all.forEach(r => {
    if (r[5] === true || r[5] === "TRUE" || !by[r[2]]) return;
    try { by[r[2]].push(Object.assign({ _id: String(r[3]) }, JSON.parse(r[6]))); } catch (err) {}
  });

  const ventes = [];
  by.tickets.sort((a, b) => String(a.at).localeCompare(String(b.at))).forEach(t => {
    (t.items || []).forEach(it => {
      ventes.push([t.day || "", t.at ? new Date(t.at) : "", t.vente || "", t.caisse || "", it.base || it.name || "", it.qty || 1, kg_(it.weight),
        eur_(it.price) * (it.qty || 1), eur_(it.disc), t.pay || "", t.status === "annule" ? "annulé" : "", t._id]);
    });
  });
  write_("Ventes", ["Jour", "Heure", "Vente", "Caisse", "Produit", "Qté", "Poids (kg)", "Prix (€)", "Remise ligne (€)", "Paiement", "Annulé", "Ticket"], ventes);

  const tickets = by.tickets.map(t => [t.day || "", t.at ? new Date(t.at) : "", t.vente || "", t.caisse || "", (t.items || []).length,
    eur_(t.subtotal || t.total), eur_(t.remise), eur_(t.total), t.pay || "", eur_(t.given), eur_(t.change), t.status === "annule" ? "annulé" : "", t._id]);
  tickets.sort((a, b) => String(a[1]).localeCompare(String(b[1])));
  write_("Tickets", ["Jour", "Heure", "Vente", "Caisse", "Lignes", "Sous-total (€)", "Remise (€)", "Total (€)", "Paiement", "Donné (€)", "Rendu (€)", "Annulé", "Ticket"], tickets);

  const retours = by.retours.sort((a, b) => String(a.at).localeCompare(String(b.at))).map(r => [r.day || "", r.at ? new Date(r.at) : "", r.client || r.vente || "",
    r.caisse || "", r.name || "", r.lot || "", r.qty || 1, kg_(r.weight), eur_(r.price), r.code || "", r._id]);
  write_("Retours", ["Jour", "Heure", "Client / vente", "Caisse", "Produit", "Lot", "Qté", "Poids (kg)", "Prix (€)", "Code-barres", "Id"], retours);

  const produits = by.products.sort((a, b) => String(a.cat).localeCompare(String(b.cat), "fr") || String(a.name).localeCompare(String(b.name), "fr"))
    .map(p => [p.cat || "", p.name || "", eur_(p.price), p.unit === "kg" ? "€/kg" : "pièce", p.ean || "", p._id]);
  write_("Produits", ["Catégorie", "Produit", "Prix (€)", "Unité", "Code-barres", "Id"], produits);
}
