/**
 * Pige de Noël — moteur (API) en Google Apps Script
 * Le site lui-même est hébergé sur GitHub Pages. Installation : voir README.md
 *
 * Règles de la pige :
 *  - on ne se pige pas soi-même;
 *  - deux conjoint(e)s ne peuvent pas se piger l'un l'autre;
 *  - on ne peut pas piger la même personne que l'an dernier (onglet « Historique »).
 */

// ===== Paramètres à ajuster =====
const CONFIG = {
  titre: 'Pige de Noël',
  budget: '',       // ex. '50 $' (laisser vide pour ne pas l'afficher)
  dateEchange: '',  // ex. '24 décembre, chez grand-maman'
  urlSite: '',      // adresse GitHub Pages, ex. 'https://moncompte.github.io/pige-noel/'
};

// ===== Constantes internes =====
const ONGLET_PART = 'Participants';
const ONGLET_HIST = 'Historique';
const ENTETES_PART = ['Code', 'Nom', 'Courriel', 'Conjoint(e)', 'Choix 1', 'Choix 2', 'Choix 3', 'Mis à jour'];
const ENTETES_HIST = ['Année', 'Donneur', 'Receveur'];
const COL_CHOIX = 5;              // colonne E
const CLE_PIGE = 'PIGE';          // la pige est gardée hors de la feuille, pour que l'organisateur ne la voie pas
const ALPHABET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // sans 0/O, 1/I/L

// ======================================================================
// Menu dans Google Sheets
// ======================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎁 Pige de Noël')
    .addItem('1. Préparer les onglets', 'preparer')
    .addItem('2. Générer les codes', 'genererCodes')
    .addItem('3. Envoyer les invitations (code + lien)', 'envoyerInvitations')
    .addSeparator()
    .addItem('Vérifier qu’une pige est possible', 'testerPige')
    .addItem('Forcer la pige maintenant', 'forcerPige')
    .addItem('Renvoyer les courriels de la pige', 'renvoyerCourriels')
    .addSeparator()
    .addItem('Préparer une nouvelle année', 'nouvelleAnnee')
    .addToUi();
}

function preparer() {
  const ss = SpreadsheetApp.getActive();
  [[ONGLET_PART, ENTETES_PART], [ONGLET_HIST, ENTETES_HIST]].forEach(([nom, entetes]) => {
    const f = ss.getSheetByName(nom) || ss.insertSheet(nom);
    f.getRange(1, 1, 1, entetes.length).setValues([entetes]).setFontWeight('bold');
    f.setFrozenRows(1);
  });
  alerte('Onglets prêts.\n\nRemplis Nom, Courriel et Conjoint(e) dans « Participants » ' +
         '(le nom du conjoint doit être écrit exactement comme dans la colonne Nom), ' +
         'puis lance « Générer les codes ».');
}

function genererCodes() {
  const f = feuille(ONGLET_PART);
  const n = f.getLastRow() - 1;
  if (n < 1) return alerte('Aucun participant dans l’onglet « Participants ».');
  const val = f.getRange(2, 1, n, 2).getValues();
  const existants = new Set(val.map(r => String(r[0]).trim().toUpperCase()).filter(Boolean));
  let crees = 0;
  val.forEach(r => {
    if (String(r[1]).trim() && !String(r[0]).trim()) {
      let c;
      do { c = nouveauCode(); } while (existants.has(c));
      existants.add(c);
      r[0] = c;
      crees++;
    }
  });
  f.getRange(2, 1, n, 1).setValues(val.map(r => [r[0]]));
  const err = validerParticipants(lireParticipants());
  alerte(crees + ' code(s) généré(s).' + (err.length ? '\n\nÀ corriger :\n- ' + err.join('\n- ') : ''));
}

function envoyerInvitations() {
  const url = CONFIG.urlSite.trim();
  if (!url) return alerte('Inscris d’abord l’adresse du site GitHub Pages dans CONFIG.urlSite (en haut du script).');
  const parts = lireParticipants();
  const err = validerParticipants(parts);
  if (err.length) return alerte('À corriger avant l’envoi :\n- ' + err.join('\n- '));
  if (!confirmer('Envoyer une invitation à ' + parts.length + ' participant(s)?')) return;
  parts.forEach(p => {
    const lien = url + '?code=' + encodeURIComponent(p.code);
    MailApp.sendEmail({
      to: p.courriel,
      name: CONFIG.titre,
      subject: '🎄 ' + CONFIG.titre + ' : inscris tes idées de cadeaux',
      htmlBody: gabarit(
        '<p>Bonjour ' + esc(p.nom) + ',</p>' +
        '<p>La pige de Noël est ouverte! Inscris tes trois idées de cadeaux ici :</p>' +
        '<p><a href="' + esc(lien) + '" style="display:inline-block;background:#1b4332;color:#fff;' +
        'padding:12px 20px;border-radius:8px;text-decoration:none">Inscrire mes choix</a></p>' +
        '<p>Ton code personnel : <strong style="letter-spacing:2px">' + esc(p.code) + '</strong></p>' +
        '<p>Dès que tout le monde aura inscrit ses choix, la pige se fera automatiquement ' +
        'et tu recevras un courriel avec la personne que tu as pigée.</p>'),
    });
  });
  alerte('Invitations envoyées.');
}

function testerPige() {
  const parts = lireParticipants();
  const err = validerParticipants(parts);
  if (err.length) return alerte('À corriger :\n- ' + err.join('\n- '));
  try {
    trouverAppariement(parts);
    alerte('✔ Une pige valide est possible avec les règles actuelles. (Rien n’a été enregistré.)');
  } catch (e) {
    alerte('✘ ' + e.message);
  }
}

function forcerPige() {
  if (lirePige()) return alerte('La pige est déjà faite.');
  const parts = lireParticipants();
  const manquants = parts.filter(p => !estInscrit(p)).map(p => p.nom);
  const avis = manquants.length ? '\n\nSans choix complets : ' + manquants.join(', ') : '';
  if (!confirmer('Lancer la pige maintenant et envoyer les courriels?' + avis)) return;
  verrouiller(() => effectuerPige(parts));
  alerte('Pige effectuée et courriels envoyés.');
}

function renvoyerCourriels() {
  const pige = lirePige();
  if (!pige) return alerte('Aucune pige n’a encore été faite.');
  if (!confirmer('Renvoyer à chaque participant le courriel de sa pige?')) return;
  const echecs = envoyerCourrielsPige(pige, lireParticipants(), true);
  alerte(echecs.length ? 'Échecs : ' + echecs.join(', ') : 'Courriels renvoyés.');
}

function nouvelleAnnee() {
  const pige = lirePige();
  if (!confirmer('Archiver la pige actuelle dans « Historique » et effacer les choix de tout le monde ' +
                 'pour une nouvelle année? (Les codes sont conservés.)')) return;
  if (pige) {
    const h = feuille(ONGLET_HIST);
    const lignes = pige.paires.map(pa => [pige.annee, pa.donneur, pa.receveur]);
    h.getRange(h.getLastRow() + 1, 1, lignes.length, 3).setValues(lignes);
    PropertiesService.getScriptProperties().deleteProperty(CLE_PIGE);
  }
  const f = feuille(ONGLET_PART);
  if (f.getLastRow() > 1) f.getRange(2, COL_CHOIX, f.getLastRow() - 1, 4).clearContent();
  alerte('Nouvelle année prête.' + (pige ? ' La pige ' + pige.annee + ' a été archivée.' : ''));
}

// ======================================================================
// Application web (ce que voit la famille)
// ======================================================================
// Le site GitHub envoie des requêtes POST { action, code, choix } et reçoit du JSON.
function doPost(e) {
  let reponse;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let data;
    if (req.action === 'connexion') data = connexion(req.code);
    else if (req.action === 'enregistrer') data = enregistrerChoix(req.code, req.choix);
    else throw new Error('Action inconnue.');
    reponse = { ok: true, data: data };
  } catch (err) {
    reponse = { ok: false, erreur: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(reponse))
    .setMimeType(ContentService.MimeType.JSON);
}

// Permet de vérifier dans un navigateur que l'API répond.
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, titre: CONFIG.titre }))
    .setMimeType(ContentService.MimeType.JSON);
}

function connexion(code) {
  const parts = lireParticipants();
  return etat(trouver(parts, code), parts);
}

function enregistrerChoix(code, choix) {
  if (!Array.isArray(choix) || choix.length !== 3) throw new Error('Trois choix sont requis.');
  choix = choix.map(c => String(c || '').trim().slice(0, 300));
  if (!choix.every(Boolean)) throw new Error('Les trois choix doivent être remplis.');

  return verrouiller(() => {
    if (lirePige()) throw new Error('La pige est déjà faite : les choix ne peuvent plus être modifiés.');
    const parts = lireParticipants();
    const p = trouver(parts, code);
    // L'apostrophe empêche qu'un texte commençant par « = » soit lu comme une formule
    const securise = choix.map(c => /^[=+\-@]/.test(c) ? "'" + c : c);
    feuille(ONGLET_PART).getRange(p.ligne, COL_CHOIX, 1, 4).setValues([[...securise, new Date()]]);
    p.choix = choix;

    if (parts.every(estInscrit)) {
      try { effectuerPige(parts); }
      catch (e) { avertirAdmin('La pige automatique a échoué : ' + e.message); }
    }
    return etat(p, parts);
  });
}

function etat(p, parts) {
  const pige = lirePige();
  let pigee = null;
  if (pige) {
    const pa = pige.paires.find(x => x.donneur === p.nom);
    const r = pa && parts.find(x => x.nom === pa.receveur);
    if (r) pigee = { nom: r.nom, choix: r.choix.filter(Boolean), avatar: avatarDe(r.nom) };
  }
  return {
    nom: p.nom,
    avatar: avatarDe(p.nom),
    choix: p.choix,
    inscrits: parts.filter(estInscrit).length,
    total: parts.length,
    enAttente: parts.filter(x => !estInscrit(x)).map(x => x.nom),
    // la famille, pour la guirlande de visages (nom, a répondu ou non, personnage)
    famille: parts.map(x => ({ nom: x.nom, inscrit: estInscrit(x), avatar: avatarDe(x.nom) })),
    pigeFaite: !!pige,
    pigee: pigee,
    budget: CONFIG.budget,
    dateEchange: CONFIG.dateEchange,
    titre: CONFIG.titre,
  };
}

// ======================================================================
// La pige
// ======================================================================
function effectuerPige(parts) {
  const err = validerParticipants(parts);
  if (err.length) throw new Error('Pige impossible :\n- ' + err.join('\n- '));
  const app = trouverAppariement(parts);
  const pige = {
    annee: new Date().getFullYear(),
    date: new Date().toISOString(),
    paires: parts.map(p => ({ donneur: p.nom, receveur: app[p.nom], envoye: false })),
  };
  ecrirePige(pige);
  envoyerCourrielsPige(pige, parts, false);
  return pige;
}

/** Retourne { donneur: receveur } respectant toutes les règles. */
function trouverAppariement(parts) {
  const noms = parts.map(p => p.nom);
  const conj = {};
  parts.forEach(p => { if (p.conjoint) conj[p.nom] = p.conjoint; });
  const avant = dernierePige();
  const interdit = (d, r) => d === r || conj[d] === r || conj[r] === d || avant[d] === r;

  // 1) Tirage parfaitement équitable : permutations au hasard jusqu'à en trouver une valide
  for (let essai = 0; essai < 20000; essai++) {
    const recv = melanger(noms.slice());
    if (noms.every((d, i) => !interdit(d, recv[i]))) {
      const res = {};
      noms.forEach((d, i) => { res[d] = recv[i]; });
      return res;
    }
  }

  // 2) Si les contraintes sont très serrées : recherche exhaustive
  const permis = {};
  noms.forEach(d => { permis[d] = noms.filter(r => !interdit(d, r)); });
  const ordre = melanger(noms.slice()).sort((a, b) => permis[a].length - permis[b].length);
  const res = {}, pris = new Set();
  const chercher = i => {
    if (i === ordre.length) return true;
    const d = ordre[i];
    for (const r of melanger(permis[d].slice())) {
      if (pris.has(r)) continue;
      res[d] = r; pris.add(r);
      if (chercher(i + 1)) return true;
      pris.delete(r); delete res[d];
    }
    return false;
  };
  if (chercher(0)) return res;
  throw new Error('Aucune pige possible avec ces règles (trop peu de participants ou trop de contraintes).');
}

/** Jumelages de l'année la plus récente de l'onglet « Historique » : { donneur: receveur } */
function dernierePige() {
  const f = SpreadsheetApp.getActive().getSheetByName(ONGLET_HIST);
  const res = {};
  if (!f || f.getLastRow() < 2) return res;
  const val = f.getRange(2, 1, f.getLastRow() - 1, 3).getValues()
    .filter(r => r[0] !== '' && String(r[1]).trim());
  if (!val.length) return res;
  const max = Math.max(...val.map(r => Number(r[0])));
  val.filter(r => Number(r[0]) === max)
     .forEach(r => { res[String(r[1]).trim()] = String(r[2]).trim(); });
  return res;
}

function envoyerCourrielsPige(pige, parts, tous) {
  const parNom = {};
  parts.forEach(p => { parNom[p.nom] = p; });
  const echecs = [];
  pige.paires.forEach(pa => {
    if (pa.envoye && !tous) return;
    try {
      const d = parNom[pa.donneur], r = parNom[pa.receveur];
      MailApp.sendEmail({
        to: d.courriel,
        name: CONFIG.titre,
        subject: '🎁 ' + CONFIG.titre + ' : ta personne pigée',
        htmlBody: courrielPige(d, r),
      });
      pa.envoye = true;
    } catch (e) {
      echecs.push(pa.donneur);
    }
  });
  ecrirePige(pige);
  avertirAdmin('La pige a été faite pour ' + pige.paires.length + ' participants et les courriels ont été envoyés.' +
    (echecs.length ? '\n\nÉchecs d’envoi : ' + echecs.join(', ') +
      '\nUtilise le menu « Renvoyer les courriels de la pige ».' : '') +
    '\n\n(Les jumelages ne sont pas indiqués ici, pour que tu gardes la surprise.)');
  return echecs;
}

function courrielPige(d, r) {
  const idees = r.choix.filter(Boolean).map(c => '<li>' + esc(c) + '</li>').join('') ||
                '<li><em>Aucune idée inscrite</em></li>';
  const details = [
    CONFIG.budget && 'Budget : ' + esc(CONFIG.budget),
    CONFIG.dateEchange && 'Échange : ' + esc(CONFIG.dateEchange),
  ].filter(Boolean).join(' · ');
  return gabarit(
    '<p>Bonjour ' + esc(d.nom) + ',</p>' +
    '<p>Cette année, tu offres un cadeau à :</p>' +
    '<p style="font-size:26px;font-weight:bold;color:#a4161a;margin:6px 0 18px">' + esc(r.nom) + '</p>' +
    '<p>Ses idées de cadeaux :</p><ul>' + idees + '</ul>' +
    (details ? '<p>' + details + '</p>' : '') +
    '<p>Chut… c’est un secret! 🤫</p>');
}

// ======================================================================
// Utilitaires
// ======================================================================
function lireParticipants() {
  const f = feuille(ONGLET_PART);
  const n = f.getLastRow() - 1;
  if (n < 1) return [];
  return f.getRange(2, 1, n, ENTETES_PART.length).getValues()
    .map((r, i) => ({
      ligne: i + 2,
      code: String(r[0]).trim().toUpperCase(),
      nom: String(r[1]).trim(),
      courriel: String(r[2]).trim(),
      conjoint: String(r[3]).trim(),
      choix: [r[4], r[5], r[6]].map(x => String(x).trim()),
    }))
    .filter(p => p.nom);
}

function estInscrit(p) { return p.choix.every(Boolean); }

/**
 * Personnage dessiné de la personne (fichier Avatars.gs, facultatif).
 * Les personnages restent dans Apps Script : le site ne les reçoit qu'après un code valide.
 */
function avatarDe(nom) {
  if (typeof AVATARS === 'undefined') return '';
  const cle = n => String(n).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z]/g, '');
  const k = cle(nom);
  for (const n in AVATARS) if (cle(n) === k) return AVATARS[n];
  return '';
}

function validerParticipants(parts) {
  const err = [];
  const parNom = {};
  parts.forEach(p => {
    if (parNom[p.nom]) err.push('Nom en double : ' + p.nom);
    parNom[p.nom] = p;
  });
  parts.forEach(p => {
    if (!p.code) err.push(p.nom + ' : aucun code (lance « Générer les codes »)');
    if (!/^\S+@\S+\.\S+$/.test(p.courriel)) err.push(p.nom + ' : courriel manquant ou invalide');
    if (p.conjoint) {
      const c = parNom[p.conjoint];
      if (!c) err.push(p.nom + ' : conjoint(e) « ' + p.conjoint + ' » introuvable dans la colonne Nom');
      else if (c.conjoint && c.conjoint !== p.nom) err.push(p.nom + ' et ' + c.nom + ' : conjoints incohérents');
    }
  });
  if (parts.length < 3) err.push('Il faut au moins 3 participants.');
  return err;
}

function trouver(parts, code) {
  code = String(code || '').trim().toUpperCase();
  const p = code && parts.find(x => x.code === code);
  if (!p) {
    Utilities.sleep(800); // ralentit les essais au hasard
    throw new Error('Code invalide.');
  }
  return p;
}

function lirePige() {
  const v = PropertiesService.getScriptProperties().getProperty(CLE_PIGE);
  return v ? JSON.parse(v) : null;
}

function ecrirePige(pige) {
  PropertiesService.getScriptProperties().setProperty(CLE_PIGE, JSON.stringify(pige));
}

function verrouiller(fn) {
  const verrou = LockService.getScriptLock();
  verrou.waitLock(30000);
  try { return fn(); } finally { verrou.releaseLock(); }
}

function avertirAdmin(message) {
  const admin = Session.getEffectiveUser().getEmail();
  if (admin) MailApp.sendEmail(admin, CONFIG.titre + ' — avis', message);
}

function feuille(nom) {
  const f = SpreadsheetApp.getActive().getSheetByName(nom);
  if (!f) throw new Error('Onglet « ' + nom + ' » introuvable : lance d’abord « Préparer les onglets ».');
  return f;
}

function nouveauCode() {
  let c = '';
  for (let i = 0; i < 6; i++) c += ALPHABET[Math.floor(Math.random() * ALPHABET.length)];
  return c;
}

function melanger(a) {
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function esc(s) {
  return String(s).replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

function gabarit(corps) {
  const pied = CONFIG.urlSite
    ? '<p style="margin:24px 0 0;font-size:13px;color:#8a7d66">Revoir sur le site : <a href="' +
      esc(CONFIG.urlSite) + '" style="color:#9e2a2b">' + esc(CONFIG.urlSite) + '</a></p>' : '';
  return '<div style="background:#123524;padding:28px 12px;font-family:Georgia,serif">' +
    '<div style="max-width:520px;margin:auto;background:#fbf6ec;border-radius:14px;overflow:hidden;' +
    'border:1px solid #d9b45b">' +
    '<div style="background:#9e2a2b;color:#fbf6ec;text-align:center;padding:18px 16px;' +
    'font-size:24px;letter-spacing:1px">❄ ' + esc(CONFIG.titre) + ' ❄</div>' +
    '<div style="padding:24px 26px;color:#2a2620;line-height:1.55;font-size:16px">' + corps + pied +
    '</div></div></div>';
}

function alerte(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

function confirmer(msg) {
  const ui = SpreadsheetApp.getUi();
  return ui.alert(msg, ui.ButtonSet.YES_NO) === ui.Button.YES;
}
