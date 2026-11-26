/* Caisse mobile : sauvegardes automatiques (comme pour les canards).
 *
 * - Chaque nuit (vers 2 h) : une copie complète (on garde les 30 dernières) du tableau « CAISSE MOBILE - SYNCHRO »
 *   dans le dossier « Sauvegardes caisse », nommée « Caisse mobile – sauvegarde AAAA-MM-JJ ».
 * - Chaque lundi (vers 6 h) : un mail avec le résumé, le lien de la dernière copie
 *   et un fichier .json à réimporter dans l'appli (Synchronisation › Importer une sauvegarde).
 *
 * À coller dans un NOUVEAU fichier du même projet Apps Script que Code.gs
 * (Extensions > Apps Script du tableau, « + » > Script, nom : Sauvegarde),
 * puis lancer une fois la fonction « installerSauvegardes ».
 * Pas besoin de refaire le déploiement de l'application Web.
 */

const SV_DOSSIER = "Sauvegardes caisse";             // créé à côté du tableau s'il n'existe pas
const SV_NOM = "Caisse mobile – sauvegarde ";         // + date
const SV_HEURE_NUIT = 2;                              // copie chaque nuit vers 2 h
const SV_HEURE_LUNDI = 6;                             // mail le lundi vers 6 h
const SV_GARDER = 30;                                 // garder les 30 copies les plus récentes (0 = tout garder)
const SV_MAIL = "";                                   // vide = ton adresse Google ; sinon "adresse1,adresse2"

/* À lancer UNE fois : crée (ou recrée) les deux déclencheurs. */
function installerSauvegardes() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === "sauvegardeNuit" || f === "mailLundi") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sauvegardeNuit").timeBased().everyDays(1).atHour(SV_HEURE_NUIT).create();
  ScriptApp.newTrigger("mailLundi").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(SV_HEURE_LUNDI).create();
  sauvegardeNuit();   // une première copie tout de suite, pour vérifier
  mailLundi();        // et un premier mail tout de suite, pour vérifier
}

function svDate_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Europe/Paris", "yyyy-MM-dd");
}

function svDossier_() {
  const fichier = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = fichier.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName(SV_DOSSIER);
  return it.hasNext() ? it.next() : parent.createFolder(SV_DOSSIER);
}

/* Copie du tableau, chaque nuit. */
function sauvegardeNuit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dossier = svDossier_();
  const nom = SV_NOM + svDate_();
  const deja = dossier.getFilesByName(nom);
  while (deja.hasNext()) deja.next().setTrashed(true);    // une seule copie par jour
  DriveApp.getFileById(ss.getId()).makeCopy(nom, dossier);

  if (SV_GARDER > 0) {
    const copies = [], tous = dossier.getFiles();
    while (tous.hasNext()) { const f = tous.next(); if (f.getName().indexOf(SV_NOM) === 0) copies.push(f); }
    copies.sort((a, b) => b.getDateCreated() - a.getDateCreated());
    copies.slice(SV_GARDER).forEach(f => f.setTrashed(true));
  }
}

/* Contenu de l'appli, au format de « Télécharger une sauvegarde ». */
function svDonnees_() {
  const out = { app: "caisse-du-stand", at: new Date().toISOString(), products: {}, tickets: {}, retours: {}, retoursFeuilles: {} };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_sync");
  const n = sh ? sh.getLastRow() - 1 : 0;
  if (n > 0) {
    sh.getRange(2, 1, n, 7).getValues().forEach(r => {
      const col = r[2], id = String(r[3]), supprime = r[5] === true || r[5] === "TRUE";
      if (supprime || !out[col] || !r[6]) return;
      try { out[col][id] = JSON.parse(r[6]); } catch (err) {}
    });
  }
  return out;
}

/* Mail du lundi. */
function mailLundi() {
  const date = svDate_();
  const d = svDonnees_();
  const nb = c => Object.keys(d[c]).length;
  const valides = Object.keys(d.tickets).map(k => d.tickets[k]).filter(t => t.status !== "annule");
  const depuis = Date.now() - 7 * 86400000;
  const semaine = valides.filter(t => t.at && new Date(t.at).getTime() >= depuis);
  const ca = semaine.reduce((s, t) => s + (Number(t.total) || 0), 0) / 100;

  const copies = svDossier_().getFiles();
  let derniere = null;
  while (copies.hasNext()) {
    const f = copies.next();
    if (f.getName().indexOf(SV_NOM) === 0 && (!derniere || f.getDateCreated() > derniere.getDateCreated())) derniere = f;
  }

  const corps = [
    "Sauvegarde automatique : " + nb("products") + " produits, " + valides.length + " tickets, " + nb("retours") + " retours.",
    "7 derniers jours : " + semaine.length + " tickets, " + ca.toFixed(2).replace(".", ",") + " €.",
    derniere ? "Dernière copie du tableau : " + derniere.getName() + "\n" + derniere.getUrl() : "Pas encore de copie du tableau.",
    "",
    "Pour la remettre dans l'appli : Synchronisation › Importer une sauvegarde, et",
    "choisir ce fichier .json."
  ].join("\n");

  const fichier = Utilities.newBlob(JSON.stringify(d), "application/json", "caisse-sauvegarde-" + date + ".json");
  MailApp.sendEmail({
    to: SV_MAIL || Session.getEffectiveUser().getEmail(),
    subject: "Sauvegarde caisse mobile – " + date,
    body: corps,
    attachments: [fichier]
  });
}
