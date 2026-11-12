/**
 * CRM Eco-Roy — serveur Google Apps Script
 * ------------------------------------------------------------
 * Rôle : base de données (Google Sheet), stockage photos (Drive),
 *        API de synchro pour l'appli, et bots automatiques.
 *
 * Installation : voir README (étapes 1 à 5). En résumé :
 *   1. Coller ce fichier dans un nouveau projet Apps Script
 *   2. Exécuter setup()  → crée la base, le dossier photos, les codes, les bots
 *   3. Déployer en « Application Web » (exécuter en tant que : moi ; accès : tout le monde)
 *   4. Exécuter migrationExistant() une fois pour importer tes Sheets actuels
 */

// ==== Sources existantes (tes Sheets actuels) ====
var SOURCES = {
  CALL_MAROC: ['1_87KvRa1jfvdxp3xAdoRvRD6K2N3_p3qf1ulpAcgzF8', '171BIwY0zo_qeAmrtnLqYz7iJfPIcrC_cr__JStXBt2Y'], // call maroc mabrour pac / iso
  SUIVI_CLIENTS: '1A3fueyXLGOSvygNXJN-OQkiM_4SLNs-B6JR27qig0t4',
  COMPTE_PAC: '1jWVXUI-Aj0OsqmdteFk6qSGp9V1brrD-ldkyykVvkrE',
  IGOR: '185ltgGF_DRZ-z_kQjhpEaw_bD3nzEFrTvm7YazIkObw'
};

var COLS = ['id', 'serverTs', 'updatedAt', 'nom', 'prenom', 'tel', 'cp', 'ville', 'source', 'etape', 'poseur', 'datePose', 'couleur', 'margePrev', 'resteAEncaisser', 'json'];
// Pipeline v2 (organigramme Sassi 10/11/2026)
var ETAPES = {1: 'Lead reçu', 2: 'Téléprospection', 3: 'Confirmation', 4: 'Devis envoyé', 5: 'Dossier · pièces & comptes', 6: 'À planifier', 7: 'Pose programmée', 8: 'Chantier réalisé', 9: 'Facture & dépôts', 10: 'Paiement en cours', 11: 'Soldé · clôturé', 12: 'Suivi long terme', 0: 'Perdu'};
var PV = 2;
function migrerV2(d) {
  if (Number(d.pv) === PV) return false;
  var solde = ['Paiement demandé', 'Payé'].indexOf(d.statutMPR) >= 0 || ['Paiement demandé', 'Payé'].indexOf(d.statutOPX) >= 0;
  var map = function (x) { x = Number(x); if (x === 8) return solde ? 10 : 9; var m = {1: 1, 2: 3, 3: 5, 4: 5, 5: 6, 6: 7, 7: 8, 9: 11, 0: 0}; return m[x] === undefined ? x : m[x]; };
  d.etape = map(d.etape);
  (d.historique || []).forEach(function (h) { if (h.pv !== PV) { h.etape = map(h.etape); h.pv = PV; } });
  if (Number(d.etape) === 4 && !d.statutDevis) d.statutDevis = 'Envoyé · en attente';
  d.pv = PV; return true;
}
/** À lancer UNE fois : passe tous les dossiers au pipeline v2 (rien n'est supprimé). */
function migrerPipelineV2() {
  var sh = dbSheet(); var index = indexRows(sh); var n = 0;
  readAll().forEach(function (r) { var d = r.data; if (migrerV2(d)) { d.updatedAt = Date.now(); writeRecord(sh, index, d); n++; } });
  Logger.log(n + ' dossiers passés au pipeline v2');
}
var POSEUR_FIELDS = ['id', 'nom', 'prenom', 'tel', 'adresse', 'cp', 'ville', 'produit', 'materiel', 'fournisseur', 'datePose', 'poseur', 'etape', 'habitat',
  'surface', 'chauffage', 'emetteurs', 'compteur', 'notesTech', 'photos', 'satisfaction', 'notesPoseur', 'updatedAt', 'createdAt', 'historique'];
var POSEUR_WRITABLE = ['photos', 'satisfaction', 'notesPoseur', 'etape', 'historique'];

/* ============================================================
   INSTALLATION
   ============================================================ */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var ss;
  if (props.getProperty('DB_ID')) { ss = SpreadsheetApp.openById(props.getProperty('DB_ID')); }
  else { ss = SpreadsheetApp.create('CRM Eco-Roy — base'); props.setProperty('DB_ID', ss.getId()); }

  var sh = ss.getSheetByName('Dossiers') || ss.insertSheet('Dossiers');
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.hideColumns(COLS.indexOf('json') + 1);

  var cfg = ss.getSheetByName('Config') || ss.insertSheet('Config');
  if (cfg.getLastRow() < 2) {
    cfg.getRange(1, 1, 1, 3).setValues([['Clé', 'Valeur', 'Explication']]).setFontWeight('bold');
    var rows = [
      ['ADMIN_TOKEN', randomCode(24), 'Code gérant à saisir dans l\'appli (ne le partage pas)'],
      ['POSEUR_Igor', randomPin(), 'PIN de Igor'],
      ['POSEUR_Kevin', randomPin(), 'PIN de Kevin'],
      ['EMAIL_ALERTES', Session.getEffectiveUser().getEmail(), 'Adresse qui reçoit les alertes et le récap du matin'],
      ['MARGE_MINI', '1000', 'Alerte si marge prévisionnelle en dessous (€)'],
      ['RELANCE_CLIENTS', 'non', 'oui = mail auto au client quand il manque des pièces (tous les 3 jours)'],
      ['SIGNATURE', 'L\'équipe Eco-Roy', 'Signature des mails clients'],
      ['FORM_ORIGINES', '*', 'Domaines autorisés pour le formulaire du site (* = tous)']
    ];
    cfg.getRange(2, 1, rows.length, 3).setValues(rows);
    cfg.autoResizeColumns(1, 3);
  }
  var def = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);

  if (!props.getProperty('PHOTOS_ID')) {
    var folder = DriveApp.createFolder('CRM Eco-Roy — Photos chantiers');
    props.setProperty('PHOTOS_ID', folder.getId());
  }

  // Bots
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('botImportLeads').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('botRecapMatin').timeBased().atHour(7).nearMinute(45).everyDays(1).inTimezone('Europe/Paris').create();
  ScriptApp.newTrigger('botRelancesClients').timeBased().atHour(10).everyDays(1).inTimezone('Europe/Paris').create();

  Logger.log('✅ Base : ' + ss.getUrl());
  Logger.log('✅ Codes : onglet « Config » de la base');
  Logger.log('➡️  Étape suivante : Déployer > Nouveau déploiement > Application Web');
}

/* ============================================================
   API (appelée par l'appli et par le formulaire du site)
   ============================================================ */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'actus') return json(actionActus());
  return json({ok: true, app: 'CRM Eco-Roy', time: Date.now()});
}
function actusSheet() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')).getSheetByName('Actualites'); }
function actionActusAdmin() {
  var sh = actusSheet(); if (!sh) return {articles: []};
  var v = sh.getDataRange().getValues(); var out = [];
  v.slice(1).forEach(function (r) { if (!r[0]) return; out.push({slug: String(r[0]), date: r[1] instanceof Date ? Utilities.formatDate(r[1], 'Europe/Paris', 'yyyy-MM-dd') : String(r[1] || ''), titre: String(r[2]), resume: String(r[3] || ''), contenu: String(r[4] || ''), categorie: String(r[5] || ''), sources: String(r[6] || ''), statut: String(r[7] || '')}); });
  out.reverse(); return {articles: out};
}
function actionActuStatut(slug, statut) {
  if (['publie', 'brouillon', 'refuse'].indexOf(statut) < 0) throw new Error('Statut inconnu');
  var sh = actusSheet(); var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(slug)) { sh.getRange(i + 1, 8).setValue(statut); CacheService.getScriptCache().remove('actus'); return {ok: true}; }
  throw new Error('Article introuvable');
}
/* ----- Actualités du site : seules les lignes au statut « publie » sont servies ----- */
function actionActus() {
  var cache = CacheService.getScriptCache(); var c = cache.get('actus'); if (c) return JSON.parse(c);
  var sh = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')).getSheetByName('Actualites');
  var out = [];
  if (sh) {
    var v = sh.getDataRange().getValues();
    v.slice(1).forEach(function (r) {
      var statut = String(r[7] || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
      if (statut !== 'publie' || !r[0] || !r[2]) return;
      var d = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Europe/Paris', 'yyyy-MM-dd') : String(r[1] || '');
      out.push({slug: String(r[0]), date: d, titre: String(r[2]), resume: String(r[3] || ''), contenu: String(r[4] || ''), categorie: String(r[5] || ''), sources: String(r[6] || '')});
    });
  }
  out.sort(function (a, b) { return a.date < b.date ? 1 : -1; });
  var res = {articles: out.slice(0, 60)};
  try { cache.put('actus', JSON.stringify(res), 600); } catch (err) {}
  return res;
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    if (req.action === 'lead') return json(actionLeadSite(req));
    if (req.action === 'leadPhotos') return json(actionLeadPhotos(req));
    if (req.action === 'leadDimensionnement') return json(actionLeadDim(req));
    var who = authenticate(req.auth || {});
    switch (req.action) {
      case 'login': return json({role: who.role, poseur: who.poseur, regie: who.regie});
      case 'pull': return json(actionPull(who, Number(req.since) || 0));
      case 'push': return json(actionPush(who, req.records || []));
      case 'upload': return json(actionUpload(who, req));
      case 'file': return json(actionFile(who, req.id));
      case 'actusAdmin': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json(actionActusAdmin());
      case 'actuStatut': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json(actionActuStatut(req.slug, req.statut));
      case 'regies': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json(actionRegies());
      case 'regieAdd': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json(actionRegieAdd(req.nom));
      case 'regieDel': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json(actionRegieDel(req.nom));
      case 'catalogue': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json({items: catRead()});
      case 'catalogueSave': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json(catSave(req));
      case 'getSettings': if (who.role !== 'admin') throw new Error('Réservé au gérant'); return json({settings: JSON.parse(PropertiesService.getScriptProperties().getProperty('SETTINGS') || 'null')});
      case 'setSettings': if (who.role !== 'admin') throw new Error('Réservé au gérant'); var js = JSON.stringify(req.settings || {}); if (js.length > 8500) throw new Error('Réglages trop volumineux'); PropertiesService.getScriptProperties().setProperty('SETTINGS', js); return json({ok: true});
      default: return json({error: 'Action inconnue'});
    }
  } catch (err) {
    return json({error: String(err.message || err)});
  }
}

function authenticate(a) {
  var c = getConfig();
  if (a.token && a.token === c.ADMIN_TOKEN) return {role: 'admin'};
  if (a.regie && a.pin) {
    var rk = Object.keys(c).filter(function (k) { return k.indexOf('REGIE_') === 0 && keyNom(k.slice(6)) === keyNom(a.regie); })[0];
    if (rk && String(c[rk]).trim() === String(a.pin).trim()) return {role: 'regie', regie: rk.slice(6)};
    Utilities.sleep(800);
    throw new Error('Nom de régie ou PIN incorrect');
  }
  if (a.poseur && a.pin) {
    var key = Object.keys(c).filter(function (k) { return k.indexOf('POSEUR_') === 0 && k.slice(7).toLowerCase() === String(a.poseur).trim().toLowerCase(); })[0];
    if (key && String(c[key]).trim() === String(a.pin).trim()) return {role: 'poseur', poseur: key.slice(7)};
    Utilities.sleep(800);
    throw new Error('Prénom ou PIN incorrect');
  }
  Utilities.sleep(800);
  throw new Error('Code incorrect');
}

function actionPull(who, since) {
  var all = readAll();
  var out = [], ids = [];
  all.forEach(function (r) {
    if (who.role === 'regie') {
      if (!visibleForRegie(r.data, who.regie)) return;
      ids.push(r.data.id);
      if (r.serverTs > since) out.push(stripForRegie(r.data));
    } else if (who.role === 'poseur') {
      if (!visibleForPoseur(r.data, who.poseur)) return;
      ids.push(r.data.id);
      if (r.serverTs > since) out.push(stripForPoseur(r.data));
    } else if (r.serverTs > since) {
      out.push(r.data);
    }
  });
  var res = {now: Date.now(), records: out};
  if (who.role === 'poseur' || who.role === 'regie') { res.full = true; res.ids = ids; }
  return res;
}

function actionPush(who, records) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = dbSheet(); var index = indexRows(sh); var returned = [];
    records.forEach(function (inc) {
      if (!inc || !inc.id) return;
      if (Number(inc.pv) !== PV) migrerV2(inc); // appli pas encore à jour : on traduit les étapes
      var row = index[inc.id]; var stored = row ? JSON.parse(sh.getRange(row, COLS.length).getValue()) : null;
      var merged;
      if (who.role === 'regie') {
        if (stored) {
          if (!visibleForRegie(stored, who.regie)) return;
          merged = JSON.parse(JSON.stringify(stored));
          merged.commentairesRegie = mergeComments(stored.commentairesRegie, inc.commentairesRegie);
          if (JSON.stringify(merged.commentairesRegie) === JSON.stringify(stored.commentairesRegie || [])) { returned.push(stripForRegie(stored)); return; }
          merged.updatedAt = Date.now();
          var lastC = merged.commentairesRegie[merged.commentairesRegie.length - 1];
          notifyAdmin('💬 Message de ' + who.regie + ' : ' + nomComplet(merged), lastC ? lastC.txt : '');
        } else {
          var tel9 = normTel(inc.tel).slice(-9);
          var dup = tel9 && readAll().filter(function (r) { return !r.data._deleted && normTel(r.data.tel).slice(-9) === tel9; })[0];
          merged = newDossier();
          REGIE_CREATE.forEach(function (k) { if (inc[k] !== undefined) merged[k] = clean(inc[k]); });
          merged.id = inc.id; merged.etape = 1; merged.source = 'Régie'; merged.regie = who.regie; merged.pv = PV;
          merged.createdAt = Date.now(); merged.updatedAt = Date.now(); merged.historique = [{t: Date.now(), etape: 1, pv: PV}];
          merged.commentairesRegie = mergeComments([], inc.commentairesRegie);
          merged.demandeRappel = 'Envoyé par la régie ' + who.regie + ' le ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
          if (dup) { merged.aVerifier = true; merged.notes = '⚠️ Doublon possible : même téléphone que « ' + nomComplet(dup.data) + ' » (déjà dans le CRM, source ' + (dup.data.source || '?') + ').'; }
          notifyAdmin('📥 Nouveau client de la régie ' + who.regie + ' : ' + nomComplet(merged), 'Tél : ' + (merged.tel || '') + ' · ' + (merged.cp || '') + ' ' + (merged.ville || '') + ' · ' + (merged.produit || '') + (dup ? '\n⚠️ Doublon possible avec ' + nomComplet(dup.data) : ''));
        }
      } else if (who.role === 'poseur') {
        if (!stored || !visibleForPoseur(stored, who.poseur)) return;
        merged = JSON.parse(JSON.stringify(stored));
        POSEUR_WRITABLE.forEach(function (k) { if (inc[k] !== undefined) merged[k] = inc[k]; });
        // un poseur ne peut que passer le chantier à « Installé »
        if (Number(inc.etape) !== Number(stored.etape)) merged.etape = (Number(inc.etape) === 8 && Number(stored.etape) <= 8) ? 8 : stored.etape;
        merged.photos = mergePhotos(stored.photos, inc.photos);
        merged.updatedAt = Math.max(Number(inc.updatedAt) || 0, Number(stored.updatedAt) || 0);
      } else {
        if (stored && Number(stored.updatedAt) > Number(inc.updatedAt)) {
          merged = stored; merged.photos = mergePhotos(stored.photos, inc.photos); merged.commentairesRegie = mergeComments(stored.commentairesRegie, inc.commentairesRegie);
        } else {
          merged = inc;
          if (stored) {
            merged.photos = mergePhotos(stored.photos, inc.photos);
            if (stored.satisfaction && !inc.satisfaction) merged.satisfaction = stored.satisfaction;
            if (stored.notesPoseur && !inc.notesPoseur) merged.notesPoseur = stored.notesPoseur;
            if (stored.commentairesRegie) merged.commentairesRegie = mergeComments(stored.commentairesRegie, inc.commentairesRegie);
          }
        }
      }
      writeRecord(sh, index, merged);
      returned.push(who.role === 'poseur' ? stripForPoseur(merged) : who.role === 'regie' ? stripForRegie(merged) : merged);
      if (who.role === 'poseur' && stored && Number(merged.etape) === 8 && Number(stored.etape) < 8) notifyAdmin('✅ Chantier terminé : ' + nomComplet(merged), who.poseur + ' a déclaré le chantier de ' + nomComplet(merged) + ' terminé.\nPhotos après : ' + countPhotos(merged, 'apres') + '\nSatisfaction : ' + (merged.satisfaction && merged.satisfaction.note ? merged.satisfaction.note + '/5' : 'non remplie'));
    });
    return {records: returned};
  } finally { lock.releaseLock(); }
}

function actionUpload(who, req) {
  var folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PHOTOS_ID'));
  if (who.role === 'regie') throw new Error('Non autorisé');
  var d = findRecord(req.dossierId);
  if (who.role === 'poseur' && (!d || !visibleForPoseur(d, who.poseur))) throw new Error('Chantier non autorisé');
  var sub = clientFolder(folder, d ? nomComplet(d) + ' — ' + (d.ville || d.cp || '') : 'Divers');
  var full = dataUrlToBlob(req.full, (req.kind || 'photo') + '_' + Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyyMMdd_HHmmss') + extOf(req.full));
  var f = sub.createFile(full);
  var thumbId = f.getId();
  if (req.thumb && req.thumb !== req.full) {
    var t = sub.createFile(dataUrlToBlob(req.thumb, 'mini_' + f.getName()));
    thumbId = t.getId();
  }
  return {fileId: f.getId(), thumbId: thumbId, url: f.getUrl()};
}

function actionFile(who, id) {
  if (who.role === 'regie') throw new Error('Non autorisé');
  var f = DriveApp.getFileById(id);
  var root = PropertiesService.getScriptProperties().getProperty('PHOTOS_ID');
  var p = f.getParents(); var ok = false;
  while (p.hasNext()) { var par = p.next(); var gp = par.getParents(); if (par.getId() === root) ok = true; while (gp.hasNext()) if (gp.next().getId() === root) ok = true; }
  if (!ok) throw new Error('Fichier non autorisé');
  var b = f.getBlob();
  return {dataUrl: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes())};
}

/* ----- Formulaire d'éligibilité du site (public) ----- */
function actionLeadSite(req) {
  var c = getConfig();
  if (req.website) return {ok: true}; // champ piège anti-robots
  var tel = normTel(req.tel);
  if (!req.nom || tel.length < 10) throw new Error('Nom et téléphone obligatoires');
  var cache = CacheService.getScriptCache();
  if (cache.get('lead_' + tel)) { var dj = readAll().filter(function (r) { return normTel(r.data.tel).slice(-9) === tel.slice(-9); })[0]; return {ok: true, deja: true, token: dj ? dj.data.leadToken : undefined}; }
  cache.put('lead_' + tel, '1', 600);

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = dbSheet(); var index = indexRows(sh);
    var exist = readAll().filter(function (r) { return normTel(r.data.tel).slice(-9) === tel.slice(-9); })[0];
    var couleur = couleurCalc(req.rfr, req.nbPersonnes, req.cp);
    var score = scoreLead(req, couleur);
    var d = exist ? exist.data : newDossier();
    var add = {
      nom: clean(req.nom), prenom: clean(req.prenom), tel: tel, email: clean(req.email), adresse: clean(req.adresse), cp: clean(req.cp), ville: clean(req.ville),
      source: req.source || 'Site internet', produit: req.produit || 'PAC air/eau', habitat: req.habitat || '', proprietaire: req.proprietaire || '',
      surface: req.surface || '', chauffage: req.chauffage || '', anneeConstruction: req.anneeConstruction || '', nbPersonnes: req.nbPersonnes || '', rfr: req.rfr || '',
      utm: req.utm || '', scoreLead: score.score, temperature: score.label
    };
    Object.keys(add).forEach(function (k) { if (add[k] !== '' && (d[k] === undefined || d[k] === '' || d[k] === null)) d[k] = add[k]; });
    d.notes = ((d.notes || '') + '\n[Site ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM HH:mm') + '] ' + (req.message || '') + ' — score ' + score.score + '/100 (' + score.label + ')' + (req.utm ? ' — ' + req.utm : '')).trim();
    if (req.rappelLe && /^\d{4}-\d\d-\d\dT\d\d:\d\d$/.test(String(req.rappelLe))) { d.rappelLe = String(req.rappelLe); d.statutAppel = 'À rappeler'; d.demandeRappel = 'Créneau choisi sur le site : ' + clean(req.creneau); d.notes += '\n📅 Créneau de rappel choisi : ' + clean(req.creneau); }
    if (!d.leadToken) d.leadToken = randomCode(20);
    d.updatedAt = Date.now();
    writeRecord(sh, index, d);
    notifyAdmin((score.label === 'chaud' ? '🔥 ' : '📥 ') + 'Lead site : ' + nomComplet(d) + ' — ' + (couleur || '?') + ' — ' + score.score + '/100',
      'Nouveau lead depuis le site.\n\nNom : ' + nomComplet(d) + '\nTél : ' + d.tel + '\nCP : ' + d.cp + ' ' + (d.ville || '') + '\nCouleur MPR : ' + (couleur || 'non calculable') +
      (d.rappelLe ? '\n📅 À rappeler : ' + (d.demandeRappel || d.rappelLe) : '') + '\nChauffage : ' + (d.chauffage || '?') + ' · Surface : ' + (d.surface || '?') + ' m²\nScore : ' + score.score + '/100 (' + score.label + ')\n' + score.detail.join('\n') + (exist ? '\n\n⚠️ Ce numéro existait déjà dans le CRM (fiche mise à jour).' : ''));
    return {ok: true, token: d.leadToken, couleur: couleur, aideMPR: couleur ? ({bleu: 5000, jaune: 4000, violet: 3000, rose: 0})[couleur] : null, score: score.score};
  } finally { lock.releaseLock(); }
}

function scoreLead(r, couleur) {
  var s = 0, detail = [];
  function add(p, why) { s += p; detail.push((p >= 0 ? '+' : '') + p + ' ' + why); }
  if (/propri/i.test(r.proprietaire || '') && !/bailleur/i.test(r.proprietaire || '')) add(20, 'propriétaire occupant');
  if (/maison/i.test(r.habitat || '')) add(15, 'maison individuelle');
  if (/gaz|fioul/i.test(r.chauffage || '')) add(20, 'chauffage gaz/fioul');
  else if (/lectri/i.test(r.chauffage || '')) add(5, 'chauffage électrique');
  if (Number(r.surface) >= 70) add(10, 'surface ≥ 70 m²');
  if (couleur === 'bleu') add(25, 'profil bleu'); else if (couleur === 'jaune') add(20, 'profil jaune'); else if (couleur === 'violet') add(10, 'profil violet'); else if (couleur === 'rose') add(0, 'profil rose (CEE seul)');
  if (r.anneeConstruction && Number(r.anneeConstruction) <= new Date().getFullYear() - 15) add(10, 'logement > 15 ans');
  if (r.email) add(0, 'e-mail fourni');
  s = Math.max(0, Math.min(100, s));
  return {score: s, label: s >= 70 ? 'chaud' : s >= 45 ? 'tiède' : 'froid', detail: detail};
}

/* ============================================================
   BOTS
   ============================================================ */

/** Toutes les 15 min : importe les nouvelles lignes des Sheets du call Maroc. */
function botImportLeads() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(10000)) return;
  try {
    var sh = dbSheet(); var index = indexRows(sh); var all = readAll(); var created = [];
    var known = {}; all.forEach(function (r) { known[keyNom(r.data.nom) + '|' + String(r.data.cp || '')] = true; var t = normTel(r.data.tel).slice(-9); if (t) known['t' + t] = true; });
    SOURCES.CALL_MAROC.forEach(function (id) {
      var rows; try { rows = readTable(id); } catch (e) { return; }
      rows.forEach(function (r) {
        var nom = r.get('nom'); if (!nom || /^dupont$/i.test(nom.trim())) return;
        var cp = r.get('cp'); var t = normTel(r.get('tel')).slice(-9);
        if (known[keyNom(nom) + '|' + cp] || (t && known['t' + t])) return;
        var d = leadFromCallRow(r);
        writeRecord(sh, index, d); created.push(d);
        known[keyNom(nom) + '|' + cp] = true; if (t) known['t' + t] = true;
      });
    });
    if (created.length) notifyAdmin('📥 ' + created.length + ' nouveau(x) lead(s) du call Maroc', created.map(function (d) { return '• ' + nomComplet(d) + ' — ' + (d.cp || '') + ' ' + (d.ville || '') + ' — ' + (d.tel || ''); }).join('\n'));
  } finally { lock.releaseLock(); }
}

/** Chaque matin : récap des actions du jour par mail. */
function botRecapMatin() {
  var c = getConfig(); var mini = Number(c.MARGE_MINI) || 1000;
  var all = readAll().map(function (r) { return r.data; }).filter(function (d) { return !d._deleted && Number(d.etape) !== 0; });
  var sec = {leads: [], pieces: [], opx: [], poses: [], terrain: [], mpr: [], soldes: [], marge: []};
  var totalAEnc = 0;
  var CRM_START = new Date('2026-11-06T00:00:00').getTime(); var nbVerif = 0;
  all.forEach(function (d) {
    var e = Number(d.etape);
    var resteV = resteAEncaisser(d); if (e >= 8 && e <= 11) totalAEnc += resteV;
    if (d.aVerifier) { nbVerif++; if (e === 10 && resteV > 0 && joursDepuis(d.dateDemandeSolde) >= 30) sec.soldes.push(nomComplet(d) + ' — ' + euros(resteV) + ' depuis ' + joursDepuis(d.dateDemandeSolde) + ' j (' + (d.statutOPX || '') + ' / ' + (d.statutMPR || '') + ')'); return; }
    var poseCRM = new Date(d.dateInstallation || d.datePose || 0).getTime() >= CRM_START;
    if (e === 1 && joursDepuis(d.createdAt) >= 1 && joursDepuis(d.createdAt) <= 30) sec.leads.push(nomComplet(d) + ' (' + joursDepuis(d.createdAt) + ' j) — ' + (d.tel || ''));
    if (e >= 4 && e <= 5 && !(d.pAvis && d.pCni && d.pDevis)) sec.pieces.push(nomComplet(d) + ' — manque ' + [!d.pAvis && 'avis', !d.pCni && 'CNI', !d.pDevis && 'devis'].filter(Boolean).join(', '));
    if (e >= 5 && e <= 6 && d.statutOPX === 'Déposé' && joursDepuis(d.dateDepotOPX) >= 10) sec.opx.push(nomComplet(d) + ' — déposé il y a ' + joursDepuis(d.dateDepotOPX) + ' j');
    if (e >= 6 && e <= 7 && d.datePose) { var j = -joursDepuis(d.datePose); if (j >= 0 && j <= 3) sec.poses.push((j === 0 ? "AUJOURD'HUI" : 'J-' + j) + ' · ' + nomComplet(d) + ' — ' + (d.poseur || '⚠ sans poseur') + (d.dateCommande ? '' : ' — ⚠ matériel non commandé')); }
    if (e >= 8 && e <= 10 && poseCRM && (countPhotos(d, 'apres') === 0 || !(d.satisfaction && d.satisfaction.note))) sec.terrain.push(nomComplet(d) + ' (' + (d.poseur || '?') + ') — manque ' + [countPhotos(d, 'apres') === 0 && 'photos après', !(d.satisfaction && d.satisfaction.note) && 'satisfaction'].filter(Boolean).join(', '));
    if (e >= 8 && e <= 10 && d.statutMPR === 'À déposer' && couleurCalc(d.rfr, d.nbPersonnes, d.cp) !== 'rose') sec.mpr.push(nomComplet(d));
    var reste = resteV;
    if (e === 10 && reste > 0 && joursDepuis(d.dateDemandeSolde) >= 30) sec.soldes.push(nomComplet(d) + ' — ' + euros(reste) + ' depuis ' + joursDepuis(d.dateDemandeSolde) + ' j (' + (d.statutOPX || '') + ' / ' + (d.statutMPR || '') + ')');
    if (e >= 4 && e <= 10 && couts(d) > 0 && recettes(d) > 0 && margePrev(d) < mini) sec.marge.push(nomComplet(d) + ' — ' + euros(margePrev(d)));
  });
  var titres = {poses: '🔧 Poses des 3 prochains jours', leads: '📞 Leads à rappeler', soldes: '💰 Soldes en attente > 30 j', pieces: '📄 Pièces manquantes', opx: '⏳ OPX sans retour (> 10 j)',
    mpr: '🏛️ Demandes MPR à déposer', terrain: '📸 Chantiers incomplets (photos / satisfaction)', marge: '📉 Marges sous ' + mini + ' €'};
  // 🔴 = à traiter aujourd'hui (argent ou chantier en jeu) · 🟠 = cette semaine
  var rouge = {poses: true, leads: true, soldes: true};
  var CRM_URL = 'https://sassi-14.github.io/crm-eco-roy/';
  var esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;'); };
  var body = ['Bonjour,', '', 'Ouvrir le CRM : ' + CRM_URL, '', 'Reste à encaisser total : ' + euros(totalAEnc), ''];
  var html = '<div style="font-family:Arial,sans-serif;font-size:15px;color:#1c2620;max-width:640px">' +
    '<p>Bonjour Sassi,</p>' +
    '<p><a href="' + CRM_URL + '" style="display:inline-block;background:#14532d;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:8px">Ouvrir le CRM</a></p>' +
    '<p style="font-size:17px"><b>💰 Reste à encaisser : ' + esc(euros(totalAEnc)) + '</b></p>';
  var n = 0, nRouge = 0;
  ['rouge', 'orange'].forEach(function (niv) {
    var cles = Object.keys(titres).filter(function (k) { return sec[k].length && (niv === 'rouge') === !!rouge[k]; });
    if (!cles.length) return;
    html += niv === 'rouge'
      ? '<div style="border-left:6px solid #d92d20;background:#fdecea;padding:10px 14px;margin:14px 0;border-radius:6px"><p style="margin:0 0 6px;color:#b42318;font-weight:bold;font-size:16px">🔴 ATTENTION — à faire aujourd\'hui</p>'
      : '<div style="border-left:6px solid #f79009;background:#fff6e5;padding:10px 14px;margin:14px 0;border-radius:6px"><p style="margin:0 0 6px;color:#b54708;font-weight:bold;font-size:16px">🟠 Cette semaine</p>';
    body.push(niv === 'rouge' ? '🔴 ATTENTION — À FAIRE AUJOURD\'HUI' : '🟠 CETTE SEMAINE');
    cles.forEach(function (k) {
      n += sec[k].length; if (niv === 'rouge') nRouge += sec[k].length;
      body.push(titres[k] + ' (' + sec[k].length + ')');
      html += '<p style="margin:8px 0 4px;font-weight:bold">' + esc(titres[k]) + ' (' + sec[k].length + ')</p><ul style="margin:0;padding-left:20px">';
      sec[k].slice(0, 10).forEach(function (l) { body.push('  • ' + l); html += '<li>' + esc(l) + '</li>'; });
      if (sec[k].length > 10) { body.push('  … + ' + (sec[k].length - 10) + ' autres dans le CRM'); html += '<li><i>… + ' + (sec[k].length - 10) + ' autres dans le CRM</i></li>'; }
      html += '</ul>'; body.push('');
    });
    html += '</div>';
  });
  if (!n) { body.push('Rien d\'urgent aujourd\'hui. 👌'); html += '<p>✅ Rien d\'urgent aujourd\'hui.</p>'; }
  if (nbVerif) { body.push('ℹ️ ' + nbVerif + ' dossiers importés restent « à vérifier ».'); html += '<p style="color:#667085;font-size:13px">ℹ️ ' + nbVerif + ' dossiers importés restent « à vérifier » (seuls leurs soldes en retard apparaissent ici).</p>'; }
  html += '<p style="color:#667085;font-size:13px">Aurore, la Matinale — équipe Eco Roy</p></div>';
  var to = c.EMAIL_ALERTES || Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({to: to, subject: (nRouge ? '🔴 ' + nRouge + ' urgence(s) · ' : '☀️ ') + 'Récap Eco Roy — ' + n + ' action(s)', body: body.join('\n'), htmlBody: html});
}

/** Chaque jour (si activé) : relance mail des clients à qui il manque des pièces, max 1 fois / 3 jours, 4 relances. */
function botRelancesClients() {
  var c = getConfig(); if (String(c.RELANCE_CLIENTS).toLowerCase() !== 'oui') return;
  var sh = dbSheet(); var index = indexRows(sh);
  readAll().forEach(function (r) {
    var d = r.data; var e = Number(d.etape);
    if (d._deleted || e < 4 || e > 5 || !d.email || (d.pAvis && d.pCni) || d.oppositionContact) return;
    d.relances = d.relances || [];
    if (d.relances.length >= 4) return;
    var last = d.relances[d.relances.length - 1]; if (last && joursDepuis(last) < 3) return;
    var manque = [!d.pAvis && 'votre dernier avis d\'imposition (toutes les pages)', !d.pCni && 'une copie de votre pièce d\'identité (recto-verso)'].filter(Boolean);
    MailApp.sendEmail({to: d.email, subject: 'Votre dossier pompe à chaleur — pièces à nous transmettre',
      body: 'Bonjour ' + (d.prenom || d.nom || '') + ',\n\nPour finaliser votre dossier d\'aides (MaPrimeRénov\' et CEE), il nous manque encore :\n- ' + manque.join('\n- ') +
        '\n\nVous pouvez simplement répondre à ce mail avec une photo lisible de chaque document.\n\nMerci et à très vite,\n' + (c.SIGNATURE || 'Eco-Roy')});
    d.relances.push(Date.now()); d.updatedAt = Date.now(); writeRecord(sh, index, d);
  });
}

/* ============================================================
   MIGRATION : import de tes Sheets actuels (à lancer UNE fois)
   ============================================================ */
function migrationExistant() {
  var sh = dbSheet(); var index = indexRows(sh);
  if (Object.keys(index).length > 0) throw new Error('La base contient déjà des dossiers : migration annulée pour éviter les doublons.');
  var list = [];
  function match(nom) { var k = keyNom(nom); if (k.length < 3) return null; return list.filter(function (d) { var x = keyNom(d.nom + ' ' + (d.prenom || '')); var y = keyNom(d.nom); return x.indexOf(k) === 0 || k.indexOf(y) === 0 && y.length >= 4; })[0] || null; }

  // 1) compte pac 1€ : finances (1 ligne = 1 dossier)
  var cp = SpreadsheetApp.openById(SOURCES.COMPTE_PAC).getSheets()[0].getDataRange().getDisplayValues();
  var h = cp.findIndex ? cp.findIndex(function (r) { return r.join('|').toUpperCase().indexOf('VALO CEE') >= 0; }) : 0;
  for (var i = h + 1; i < cp.length; i++) {
    var r = cp[i]; var nomBrut = String(r[0] || '').trim(); if (!nomBrut || /^clients$/i.test(nomBrut)) continue;
    var parts = nomBrut.split(/\s+/); var d = newDossier();
    d.nom = cap(nomBrut); d.prenom = ''; d.source = mapSource(r[7]);
    var dateTxt = String(r[1] || '').trim(); var dt = parseDateFr(dateTxt);
    if (dt) d.datePose = dt; else if (dateTxt) d.notes = 'Colonne date : ' + dateTxt;
    d.valoCEE = firstNum(r[2]); if (/off\s*watt/i.test(r[2])) d.notes = ((d.notes || '') + ' · CEE via Off Watt').trim();
    d.valoMPR = firstNum(r[3]);
    var st = splitNomPrix(r[4]); d.poseur = st.nom; d.prixPoseur = st.prix;
    d.produit = mapProduit(r[5]); var fo = splitNomPrix(r[6]); d.prixMateriel = fo.prix; d.fournisseur = mapFournisseur(r[6]);
    d.regie = String(r[7] || '').trim(); d.commission = firstNum(r[8]);
    if (/manque/i.test(r[8])) d.notes = ((d.notes || '') + ' · Commission : ' + r[8]).trim();
    var statut = String(r[10] || '') + ' ' + String(r[11] || '');
    applyStatutTexte(d, statut);
    d.etape = /instal/i.test(r[10]) && !/prevue|prévue/i.test(r[10]) ? 9 : (/previsite/i.test(statut) ? 2 : (d.datePose ? 7 : 3));
    if (/prevue|prévue/i.test(r[10])) d.etape = 6;
    d.aVerifier = true; d.historique = [{t: Date.now(), etape: d.etape}];
    list.push(d);
  }

  // 2) Suivi clients : étapes, statuts OPX / MPR, soldes
  var su = SpreadsheetApp.openById(SOURCES.SUIVI_CLIENTS).getSheets()[0].getDataRange().getDisplayValues();
  for (var j = 1; j < su.length; j++) {
    var s = su[j]; if (!s[0]) continue;
    var dd = match(s[0]); if (!dd) { dd = newDossier(); dd.nom = cap(s[0]); dd.aVerifier = true; list.push(dd); }
    var et = parseInt(String(s[1]).replace(/\D.*$/, ''), 10);
    var mapE = {2: 2, 5: 5, 7: 7, 9: 8}; if (mapE[et] !== undefined) dd.etape = mapE[et];
    if (s[2]) dd.tel = normTel(s[2]);
    if (s[3]) { var em = String(s[3]).split(/\s+/)[0]; if (/@/.test(em)) dd.email = em; } // on ne reprend JAMAIS un mot de passe
    if (/^\d{5}$/.test(String(s[4]).trim())) dd.cp = String(s[4]).trim();
    if (s[5] && !dd.source) dd.source = mapSource(s[5]);
    if (s[6]) dd.statutMPR = ({'Déposé': 'Déposé', 'Validé': 'Octroyé', 'À déposer': 'À déposer'})[s[6]] || dd.statutMPR;
    if (s[7]) dd.statutOPX = ({'Déposé': 'Déposé', 'Validé': 'Validé', 'À déposer': 'À déposer'})[s[7]] || dd.statutOPX;
    if (s[8] && !dd.poseur) dd.poseur = cap(s[8]);
    if (s[9] && !dd.fournisseur) dd.fournisseur = mapFournisseur(s[9]);
    if (s[11] === 'Reçu') { dd.statutMPR = 'Payé'; dd.encaisseMPR = dd.valoMPR || ''; } else if (s[11] === 'Demandé') dd.statutMPR = 'Paiement demandé'; else if (s[11] === 'Non concerné') dd.statutMPR = 'Non concerné';
    if (s[12] === 'Reçu') { dd.statutOPX = 'Payé'; dd.encaisseCEE = dd.valoCEE || ''; } else if (s[12] === 'Demandé') dd.statutOPX = 'Paiement demandé'; else if (s[12] === 'Pièces demandées') dd.notes = ((dd.notes || '') + ' · OPX : pièces demandées').trim();
    if (s[13]) dd.notes = ((dd.notes || '') + ' · ' + s[13]).trim();
    if (dd.etape === 8 && !dd.dateDemandeSolde) dd.dateDemandeSolde = dd.datePose || '';
    if (dd.etape >= 8 && resteAEncaisser(dd) === 0 && (dd.valoCEE || dd.valoMPR)) dd.etape = 9;
    dd.historique = [{t: Date.now(), etape: dd.etape}];
  }

  // 3) igor install + 4) call Maroc : coordonnées et infos techniques
  var techSources = [SOURCES.IGOR].concat(SOURCES.CALL_MAROC);
  techSources.forEach(function (id, srcIdx) {
    var rows; try { rows = readTable(id); } catch (e) { return; }
    rows.forEach(function (rr) {
      var nom = rr.get('nom'); if (!nom || /^dupont$/i.test(nom.trim())) return;
      var x = match(nom) || match(nom + ' ' + rr.get('prenom'));
      if (!x) { x = leadFromCallRow(rr); if (srcIdx === 0) x.source = 'Igor'; list.push(x); return; }
      fillBlanks(x, leadFromCallRow(rr));
    });
  });

  list.forEach(function (d) {
    // anciens chantiers « INSTALLER » absents du Suivi clients : considérés soldés (à vérifier)
    if (Number(d.etape) === 9) {
      if (d.valoCEE && !d.encaisseCEE) { d.encaisseCEE = d.valoCEE; d.statutOPX = 'Payé'; }
      if (d.valoMPR && !d.encaisseMPR) { d.encaisseMPR = d.valoMPR; d.statutMPR = 'Payé'; }
    }
    // « MARTINO Carla » + prénom « Carla » → nom « Martino »
    var p = keyNom(d.prenom);
    if (p && keyNom(d.nom).length > p.length && keyNom(d.nom).slice(-p.length) === p) {
      d.nom = d.nom.split(' ').filter(function (w) { return keyNom(w) !== p; }).join(' ') || d.nom;
    }
    writeRecord(sh, index, d);
  });
  Logger.log('✅ ' + list.length + ' dossiers importés. Ils sont marqués « à vérifier » dans l\'appli.');
}

/* ============================================================
   Outils
   ============================================================ */
function leadFromCallRow(r) {
  var d = newDossier();
  d.source = 'Call Maroc'; d.nom = cap(r.get('nom')); d.prenom = cap(r.get('prenom'));
  d.adresse = r.get('adresse').trim(); d.cp = r.get('cp').trim(); if (/^\d{4}$/.test(d.cp)) d.cp = '0' + d.cp; d.ville = cap(r.get('ville'));
  d.tel = normTel(r.get('tel')); var em = r.get('mail'); if (/@/.test(em)) d.email = em.split(/[\s\/]+/).filter(function (x) { return /@/.test(x); })[0];
  d.compteur = /tri|380/i.test(r.get('compteur')) ? 'Tri' : (/mono/i.test(r.get('compteur')) ? 'Mono' : '');
  d.habitat = /maison/i.test(r.get('habitat')) ? 'Maison individuelle' : (/appart/i.test(r.get('habitat')) ? 'Appartement' : '');
  var ch = r.get('chauffage') + ' ' + r.get('_A'); d.chauffage = /fioul/i.test(ch) ? 'Chaudière fioul' : (/gaz/i.test(ch) ? 'Chaudière gaz' : (/bois/i.test(ch) ? 'Bois' : ''));
  var surf = (r.get('situation') + ' ' + r.get('_A')).match(/(\d{2,3})\s*m/i); if (surf) d.surface = surf[1];
  var rad = r.get('_A').match(/(\d{1,2})\s*rad/i); if (rad) d.emetteurs = rad[1] + ' radiateurs';
  var coul = r.get('categorie').toLowerCase(); if (/bleu|jaune|violet|rose/.test(coul)) d.couleurForcee = coul.match(/bleu|jaune|violet|rose/)[0];
  d.produit = mapProduit(r.get('produit') || 'PAC');
  d.notes = [r.get('_A'), r.get('installe')].filter(function (x) { return x && x.trim(); }).join(' · ');
  var inst = r.get('installe');
  if (/infesab|infaisab/i.test(inst)) d.etape = 0;
  return d;
}

function readTable(id) {
  var v = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getDisplayValues();
  var hRow = -1;
  for (var i = 0; i < Math.min(6, v.length); i++) { var j = v[i].join('|').toLowerCase(); if (j.indexOf('nom') >= 0 && j.indexOf('cp') >= 0) { hRow = i; break; } }
  if (hRow < 0) return [];
  var H = v[hRow].map(function (x) { return String(x).toLowerCase().trim(); });
  function col(re) { for (var k = 0; k < H.length; k++) if (re.test(H[k])) return k; return -1; }
  var map = {nom: col(/^nom$/), prenom: col(/pr[ée]nom/), adresse: col(/adresse/), cp: col(/^cp$|code postal/), ville: col(/ville/), tel: col(/^t[ée]l/),
    compteur: col(/compteur/), mail: col(/mail/), situation: col(/situation/), habitat: col(/habitat/), chauffage: col(/chauffage/), categorie: col(/cat[ée]gorie/),
    produit: col(/produit/), installe: col(/install/), _A: 0};
  return v.slice(hRow + 1).map(function (row) { return {get: function (k) { var c = map[k]; return c >= 0 ? String(row[c] || '') : ''; }}; });
}

function applyStatutTexte(d, t) {
  t = String(t).toUpperCase();
  if (/DEMANDE MPR PAIEMENT/.test(t)) d.statutMPR = 'Paiement demandé';
  else if (/DEMANDE MPR FAIT/.test(t)) d.statutMPR = 'Déposé';
  else if (/ATTENTE OCTROI/.test(t)) d.statutMPR = 'Déposé';
  else if (/ATTENTE CODE/.test(t)) d.statutMPR = 'Attente code';
  if (/OPX FAIT/.test(t)) d.statutOPX = 'Déposé';
}
function mapSource(s) { s = String(s || '').toLowerCase(); if (/call|maroc/.test(s)) return 'Call Maroc'; if (/hello|ha\b/.test(s)) return 'Hello Artisan'; if (/effy/.test(s)) return 'Effy'; if (/igor/.test(s)) return 'Igor'; if (/parrain/.test(s)) return 'Parrainage'; if (/interne|direct/.test(s)) return 'Direct'; if (/fred/.test(s)) return 'Apporteur'; return s ? 'Autre' : ''; }
function mapProduit(s) { s = String(s || '').toLowerCase(); if (/ite/.test(s) && !/pac/.test(s)) return 'ITE'; if (/ssc/.test(s)) return 'SSC'; if (/btd|thermo/.test(s)) return 'PAC + ballon thermo'; if (/ballon\s*el|be\b/.test(s)) return 'PAC + ballon élec'; if (/pac/.test(s)) return 'PAC air/eau'; return 'Autre'; }
function mapFournisseur(s) { s = String(s || '').toLowerCase(); if (/brother/.test(s)) return 'Brother Négoce'; if (/mick|micha|bengo|axdis/.test(s)) return 'Axdis'; if (/elle\s*m/.test(s)) return 'LM Plastic'; if (/\blm\b|plastic/.test(s)) return 'LM Plastic'; if (/fourni\s*pose/.test(s)) return 'Fourni-posé'; return s.trim() ? 'Autre' : ''; }
function splitNomPrix(s) { s = String(s || ''); var p = firstNum(s); var nom = s.replace(/-?\d[\d\s.,]*/g, ' ').replace(/\b(ttc|ht|paye|payé|poseur|€)\b/ig, ' ').replace(/[\/?]/g, ' ').replace(/\s+/g, ' ').trim(); return {prix: p, nom: cap(nom)}; }
function firstNum(s) { var m = String(s || '').match(/-?\d[\d\s]*(?:[.,]\d+)?/); if (!m) return ''; return String(Number(m[0].replace(/\s/g, '').replace(',', '.'))); }
function parseDateFr(s) {
  s = String(s || '').toLowerCase().trim(); if (!s) return '';
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); if (m) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + pad(m[2]) + '-' + pad(m[1]); }
  var mois = ['janv', 'f[ée]v', 'mars', 'avr', 'mai', 'juin', 'juil', 'ao[uû]', 'sept', 'oct', 'nov', 'd[ée]c'];
  var dm = s.match(/(\d{1,2})\s*([a-zéû]+)/); if (!dm) return '';
  for (var i = 0; i < 12; i++) if (new RegExp('^' + mois[i]).test(dm[2])) return new Date().getFullYear() + '-' + pad(i + 1) + '-' + pad(dm[1]);
  return '';
}
function pad(n) { return ('0' + n).slice(-2); }
function cap(s) { s = String(s || '').trim().replace(/\s+/g, ' '); return s.toLowerCase().replace(/(^|[\s\-'])([a-zà-ÿ])/g, function (a, b, c) { return b + c.toUpperCase(); }); }
function clean(s) { return String(s == null ? '' : s).trim().slice(0, 300); }
function keyNom(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]/g, ''); }
function normTel(s) { var d = String(s || '').split(/\/\/|\/| ou /)[0].replace(/\D/g, ''); if (d.indexOf('33') === 0 && d.length === 11) d = '0' + d.slice(2); if (d.length === 9) d = '0' + d; return d; }
function fillBlanks(a, b) { Object.keys(b).forEach(function (k) { if (['id', 'createdAt', 'updatedAt', 'historique', 'etape', 'source', 'photos'].indexOf(k) >= 0) return; if ((a[k] === undefined || a[k] === '') && b[k] !== '' && b[k] !== undefined) a[k] = b[k]; }); if (b.notes && a.notes !== b.notes && String(a.notes || '').indexOf(b.notes) < 0) a.notes = ((a.notes || '') + ' · ' + b.notes).trim(); }

function newDossier() {
  var now = Date.now();
  return {id: now.toString(36) + Math.random().toString(36).slice(2, 8), createdAt: now, updatedAt: now, etape: 1, source: '', nom: '', prenom: '', tel: '', email: '', adresse: '', cp: '', ville: '',
    produit: 'PAC air/eau', habitat: '', proprietaire: '', nbPersonnes: '', rfr: '', statutOPX: 'À déposer', statutMPR: 'À déposer', statutAppel: 'À contacter', pv: PV, photos: [], historique: [{t: now, etape: 1, pv: PV}]};
}
function nomComplet(d) { return [d.nom, d.prenom].filter(Boolean).join(' ') || 'Sans nom'; }
function countPhotos(d, k) { return (d.photos || []).filter(function (p) { return p.kind === k && !p.del; }).length; }
function n(v) { var x = Number(String(v == null ? '' : v).replace(/\s/g, '').replace(',', '.')); return isFinite(x) ? x : 0; }
function couts(d) { return n(d.prixMateriel) + n(d.prixPoseur) + n(d.commission); }
function recettes(d) { return n(d.valoCEE) + n(d.valoMPR) + n(d.rac); }
function margePrev(d) { return recettes(d) - couts(d); }
function resteAEncaisser(d) { return Math.max(0, n(d.valoCEE) - n(d.encaisseCEE)) + Math.max(0, n(d.valoMPR) - n(d.encaisseMPR)); }
function joursDepuis(ts) { if (!ts) return null; var t = new Date(ts).getTime(); if (isNaN(t)) return null; return Math.floor((Date.now() - t) / 864e5); }
function euros(v) { return Math.round(v).toLocaleString('fr-FR') + ' €'; }

var PLAFONDS = {idf: {bleu: [24031, 35270, 42357, 49455, 56580, 7116], jaune: [29253, 42933, 51564, 60208, 68877, 8663], violet: [40851, 60051, 71846, 84562, 96817, 12257]},
  hors: {bleu: [17363, 25393, 30540, 35676, 40835, 5151], jaune: [22259, 32553, 39148, 45735, 52348, 6598], violet: [31185, 45842, 55196, 64550, 73907, 9357]}};
function couleurCalc(rfr, nb, cp) {
  if (rfr === '' || rfr == null || isNaN(Number(rfr)) || !nb) return '';
  var d = String(cp || '').replace(/\s/g, ''); if (d.length === 4) d = '0' + d; d = d.slice(0, 2);
  var t = ['75', '77', '78', '91', '92', '93', '94', '95'].indexOf(d) >= 0 ? PLAFONDS.idf : PLAFONDS.hors;
  function pl(a) { var k = Math.max(1, Number(nb)); return k <= 5 ? a[k - 1] : a[4] + (k - 5) * a[5]; }
  rfr = Number(rfr); if (rfr <= pl(t.bleu)) return 'bleu'; if (rfr <= pl(t.jaune)) return 'jaune'; if (rfr <= pl(t.violet)) return 'violet'; return 'rose';
}

function visibleForPoseur(d, poseur) {
  if (!d || d._deleted) return false;
  var e = Number(d.etape); var p = String(d.poseur || '').toLowerCase();
  if (p.indexOf(String(poseur).toLowerCase()) < 0) return false;
  if (e >= 6 && e <= 7) return true;
  if (e >= 8 && e <= 11) { var since = joursDepuis(d.dateInstallation || d.datePose); return since === null || since <= 30; }
  return false;
}
/* ----- Régies : chacune ne voit que ses clients, sans les montants (sauf sa commission) ----- */
var REGIE_FIELDS = ['id', 'nom', 'prenom', 'tel', 'email', 'adresse', 'cp', 'ville', 'produit', 'etape', 'createdAt', 'updatedAt', 'source', 'regie', 'pv', 'historique',
  'statutDevis', 'dateEnvoiDevis', 'dateSignatureDevis', 'datePose', 'dateInstallation', 'motifKO', 'commission', 'commissionPayeeLe', 'commentairesRegie', 'chauffage', 'surface', 'proprietaire', '_deleted'];
var REGIE_CREATE = ['nom', 'prenom', 'tel', 'email', 'adresse', 'cp', 'ville', 'produit', 'chauffage', 'surface', 'proprietaire'];
function visibleForRegie(d, regie) {
  if (!d || d._deleted) return false;
  var k = keyNom(regie); if (!k) return false;
  return keyNom(d.regie) === k || keyNom(d.source) === k;
}
function configSheet() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')).getSheetByName('Config'); }
function actionRegies() {
  var v = configSheet().getDataRange().getValues(); var all = readAll(); var out = [];
  v.forEach(function (r) { var k = String(r[0] || ''); if (k.indexOf('REGIE_') !== 0) return; var nom = k.slice(6);
    var n = all.filter(function (x) { return visibleForRegie(x.data, nom); }).length; out.push({nom: nom, pin: String(r[1]), clients: n}); });
  return {regies: out};
}
function actionRegieAdd(nom) {
  nom = String(nom || '').replace(/[^\wÀ-ÿ' -]/g, '').trim().slice(0, 40); if (!nom) throw new Error('Nom de régie vide');
  var sh = configSheet(); var v = sh.getDataRange().getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).indexOf('REGIE_') === 0 && keyNom(String(v[i][0]).slice(6)) === keyNom(nom)) return {nom: String(v[i][0]).slice(6), pin: String(v[i][1]), existait: true};
  var pin = randomPin(); sh.appendRow(['REGIE_' + nom, pin, 'PIN de la régie ' + nom]);
  CacheService.getScriptCache().remove('cfg'); return {nom: nom, pin: pin};
}
function actionRegieDel(nom) {
  var sh = configSheet(); var v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]).indexOf('REGIE_') === 0 && keyNom(String(v[i][0]).slice(6)) === keyNom(nom)) sh.deleteRow(i + 1);
  CacheService.getScriptCache().remove('cfg'); return {ok: true};
}
function stripForRegie(d) { var o = {}; REGIE_FIELDS.forEach(function (k) { if (d[k] !== undefined) o[k] = d[k]; }); return o; }
function mergeComments(a, b) {
  var m = {}; (a || []).concat(b || []).forEach(function (c) { if (!c || !c.t || !c.txt) return; m[c.t + '|' + String(c.txt).slice(0, 40)] = {t: Number(c.t), by: clean(c.by), txt: String(c.txt).slice(0, 2000)}; });
  return Object.keys(m).map(function (k) { return m[k]; }).sort(function (x, y) { return x.t - y.t; });
}
function stripForPoseur(d) { var o = {}; POSEUR_FIELDS.forEach(function (k) { if (d[k] !== undefined) o[k] = d[k]; }); return o; }
function mergePhotos(a, b) {
  var m = {}; (a || []).concat(b || []).forEach(function (p) { if (!p || !p.id) return; var o = m[p.id] || {}; Object.keys(p).forEach(function (k) { if (p[k] !== undefined && p[k] !== '' && !(k === 'pending' && o.fileId)) o[k] = p[k]; }); if (o.fileId) o.pending = false; m[p.id] = o; });
  return Object.keys(m).map(function (k) { return m[k]; });
}

function dbSheet() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')).getSheetByName('Dossiers'); }
function indexRows(sh) { var last = sh.getLastRow(); var idx = {}; if (last < 2) return idx; sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (r[0]) idx[r[0]] = i + 2; }); return idx; }
function readAll() {
  var sh = dbSheet(); var last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, COLS.length).getValues().filter(function (r) { return r[0]; }).map(function (r) { return {serverTs: Number(r[1]) || 0, data: JSON.parse(r[COLS.length - 1])}; });
}
function findRecord(id) { var r = readAll().filter(function (x) { return x.data.id === id; })[0]; return r ? r.data : null; }
function writeRecord(sh, index, d) {
  var row = [d.id, Date.now(), d.updatedAt || Date.now(), d.nom || '', d.prenom || '', d.tel || '', d.cp || '', d.ville || '', d.source || '',
    (d._deleted ? 'SUPPRIMÉ' : (ETAPES[Number(d.etape)] || d.etape)), d.poseur || '', d.datePose || '', d.couleurForcee || couleurCalc(d.rfr, d.nbPersonnes, d.cp),
    Math.round(margePrev(d)), Math.round(resteAEncaisser(d)), JSON.stringify(d)];
  if (row[COLS.length - 1].length > 49000) throw new Error('Dossier trop volumineux : ' + nomComplet(d));
  if (index[d.id]) sh.getRange(index[d.id], 1, 1, COLS.length).setValues([row]);
  else { sh.appendRow(row); index[d.id] = sh.getLastRow(); }
}
function getConfig() {
  var c = CacheService.getScriptCache().get('cfg'); if (c) return JSON.parse(c);
  var v = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')).getSheetByName('Config').getDataRange().getValues();
  var o = {}; v.slice(1).forEach(function (r) { if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim(); });
  CacheService.getScriptCache().put('cfg', JSON.stringify(o), 120); return o;
}
function clientFolder(root, name) { name = name.replace(/[\\\/:*?"<>|]/g, ' ').trim().slice(0, 80); var it = root.getFoldersByName(name); return it.hasNext() ? it.next() : root.createFolder(name); }
function dataUrlToBlob(dataUrl, name) { var m = String(dataUrl).match(/^data:([^;]+);base64,(.*)$/); if (!m) throw new Error('Image invalide'); return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name); }
function extOf(dataUrl) { return /image\/png/.test(dataUrl) ? '.png' : '.jpg'; }
function notifyAdmin(subject, body) { try { var to = getConfig().EMAIL_ALERTES; if (to) MailApp.sendEmail(to, subject, body); } catch (e) { Logger.log(e); } }
function randomCode(len) { var c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789', s = ''; for (var i = 0; i < len; i++) s += c.charAt(Math.floor(Math.random() * c.length)); return s; }
function randomPin() { return String(Math.floor(100000 + Math.random() * 900000)); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ============================================================
   CATALOGUE PRODUITS (onglet « Catalogue » de la base)
   ============================================================ */
var CAT_COLS = ['id', 'categorie', 'marque', 'reference', 'operation', 'prixTTC', 'tva', 'libelle', 'description', 'prixAchat', 'actif', 'maj'];
function catSheet() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID'));
  var sh = ss.getSheetByName('Catalogue');
  if (!sh) { sh = ss.insertSheet('Catalogue'); sh.getRange(1, 1, 1, CAT_COLS.length).setValues([CAT_COLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function catRead() {
  var sh = catSheet(); var last = sh.getLastRow(); if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, CAT_COLS.length).getValues().filter(function (r) { return r[0]; }).map(function (r) { var o = {}; CAT_COLS.forEach(function (k, i) { o[k] = r[i]; }); return o; });
}
function catSave(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = catSheet(); var items = req.items || []; var del = req.deleteIds || [];
    if (req.replaceAll) { if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, CAT_COLS.length).clearContent(); }
    var last = sh.getLastRow(); var idx = {};
    if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (r[0]) idx[r[0]] = i + 2; });
    var rowsNew = [];
    items.forEach(function (it) {
      if (!it.id) it.id = 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      it.maj = Date.now();
      var row = CAT_COLS.map(function (k) { var v = it[k] == null ? '' : it[k]; return typeof v === 'string' ? v.slice(0, 5000) : v; });
      if (idx[it.id]) sh.getRange(idx[it.id], 1, 1, CAT_COLS.length).setValues([row]); else rowsNew.push(row);
    });
    if (rowsNew.length) sh.getRange(sh.getLastRow() + 1, 1, rowsNew.length, CAT_COLS.length).setValues(rowsNew);
    del.forEach(function (id) { if (idx[id]) sh.getRange(idx[id], 1, 1, CAT_COLS.length).clearContent(); });
    return {ok: true, saved: items.length, deleted: del.length};
  } finally { lock.releaseLock(); }
}

/* ============================================================
   BOT RELANCES PAIEMENTS (CEE → délégataire, MaPrimeRénov' → Anah)
   Chaque lundi 9h. Désactivé tant que RELANCE_CEE / RELANCE_MPR ≠ oui.
   ============================================================ */
function ensureConfig() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID'));
  var cfg = ss.getSheetByName('Config'); var have = {};
  cfg.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) have[String(r[0]).trim()] = true; });
  var add = [
    ['EMAIL_DRAPO', '', 'E-mail de ton contact chez le délégataire CEE (relances des primes CEE)'],
    ['RELANCE_CEE', 'non', 'oui = relance automatique du délégataire chaque lundi pour les primes CEE en attente > 21 j'],
    ['EMAIL_ANAH', '', 'E-mail de la délégation Anah qui instruit tes dossiers (facultatif)'],
    ['RELANCE_MPR', 'non', 'oui = relance automatique de l\'Anah chaque lundi pour les soldes MPR en attente > 30 j'],
    ['DELAI_RELANCE_CEE', '21', 'Jours d\'attente avant relance CEE'],
    ['DELAI_RELANCE_MPR', '30', 'Jours d\'attente avant relance MPR']
  ].filter(function (r) { return !have[r[0]]; });
  if (add.length) cfg.getRange(cfg.getLastRow() + 1, 1, add.length, 3).setValues(add);
  CacheService.getScriptCache().remove('cfg');
}
function botRelancesPaiements() {
  ensureConfig();
  var c = getConfig(); var sh = dbSheet(); var index = indexRows(sh);
  var dCEE = Number(c.DELAI_RELANCE_CEE) || 21, dMPR = Number(c.DELAI_RELANCE_MPR) || 30;
  var cee = [], mpr = [];
  readAll().forEach(function (r) {
    var d = r.data; var e = Number(d.etape); if (d._deleted || e < 8 || e > 10) return;
    var ref = d.dateDemandeSolde || d.dateInstallation || d.datePose; var j = joursDepuis(ref); if (j === null) return;
    var resteCEE = Math.max(0, n(d.valoCEE) - n(d.encaisseCEE)), resteMPR = Math.max(0, n(d.valoMPR) - n(d.encaisseMPR));
    if (resteCEE > 0 && j >= dCEE && d.statutOPX !== 'Payé') cee.push({d: d, reste: resteCEE, j: j});
    if (resteMPR > 0 && j >= dMPR && d.statutMPR !== 'Payé' && d.statutMPR !== 'Non concerné') mpr.push({d: d, reste: resteMPR, j: j});
  });
  var sig = c.SIGNATURE || 'Eco-Roy'; var resume = [];
  function ligne(x, refKey) { return '- ' + nomComplet(x.d) + (x.d.cp ? ' (' + x.d.cp + ')' : '') + (x.d[refKey] ? ' — dossier n° ' + x.d[refKey] : '') + ' — ' + euros(x.reste) + ' — en attente depuis ' + x.j + ' jours'; }
  // CEE → délégataire
  if (cee.length) {
    var corpsCEE = 'Bonjour,\n\nNous revenons vers vous concernant le versement des primes CEE pour les dossiers suivants, travaux réalisés et dossiers complets de notre côté :\n\n' +
      cee.map(function (x) { return ligne(x, 'refOPX'); }).join('\n') + '\n\nTotal en attente : ' + euros(cee.reduce(function (s, x) { return s + x.reste; }, 0)) +
      '\n\nPouvez-vous nous indiquer pour chacun le statut et la date de paiement prévue, ou les pièces éventuellement manquantes ?\n\nMerci d\'avance,\n' + sig;
    if (String(c.RELANCE_CEE).toLowerCase() === 'oui' && c.EMAIL_DRAPO) {
      MailApp.sendEmail({to: c.EMAIL_DRAPO, cc: c.EMAIL_ALERTES, subject: 'Relance paiement primes CEE — ' + cee.length + ' dossier(s) — ' + sig, body: corpsCEE});
      cee.forEach(function (x) { (x.d.relancesPaiement = x.d.relancesPaiement || []).push({t: Date.now(), type: 'CEE'}); x.d.updatedAt = Date.now(); writeRecord(sh, index, x.d); });
      resume.push('✅ Relance CEE envoyée à ' + c.EMAIL_DRAPO + ' (' + cee.length + ' dossiers).');
    } else resume.push('⏸ Relance CEE prête mais NON envoyée (RELANCE_CEE=non ou EMAIL_DRAPO vide). Texte :\n\n' + corpsCEE);
  }
  // MPR → Anah
  if (mpr.length) {
    var corpsMPR = 'Bonjour,\n\nEn qualité de mandataire MaPrimeRénov\' (ECO ROY), nous sollicitons le versement des soldes pour les dossiers suivants, dont les demandes de paiement ont été déposées :\n\n' +
      mpr.map(function (x) { return ligne(x, 'refMPR'); }).join('\n') + '\n\nPouvez-vous nous indiquer l\'état d\'instruction de ces demandes et, le cas échéant, les pièces complémentaires attendues ?\n\nCordialement,\n' + sig;
    if (String(c.RELANCE_MPR).toLowerCase() === 'oui' && c.EMAIL_ANAH) {
      MailApp.sendEmail({to: c.EMAIL_ANAH, cc: c.EMAIL_ALERTES, subject: 'Mandataire ECO ROY — relance paiement soldes MaPrimeRénov\' (' + mpr.length + ' dossiers)', body: corpsMPR});
      mpr.forEach(function (x) { (x.d.relancesPaiement = x.d.relancesPaiement || []).push({t: Date.now(), type: 'MPR'}); x.d.updatedAt = Date.now(); writeRecord(sh, index, x.d); });
      resume.push('✅ Relance MPR envoyée à ' + c.EMAIL_ANAH + ' (' + mpr.length + ' dossiers).');
    } else resume.push('⏸ Relance MPR prête mais NON envoyée (RELANCE_MPR=non ou EMAIL_ANAH vide). À coller aussi dans la messagerie de chaque dossier sur maprimerenov.gouv.fr :\n\n' + corpsMPR);
  }
  if (resume.length) notifyAdmin('💰 Relances paiements — CEE ' + euros(cee.reduce(function (s, x) { return s + x.reste; }, 0)) + ' · MPR ' + euros(mpr.reduce(function (s, x) { return s + x.reste; }, 0)), resume.join('\n\n————————\n\n'));
}
/** À lancer une fois après mise à jour : ajoute les nouveaux réglages et le déclencheur du lundi sans toucher aux autres. */
function installerNouveauxBots() {
  ensureConfig();
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'botRelancesPaiements'; });
  if (!has) ScriptApp.newTrigger('botRelancesPaiements').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).inTimezone('Europe/Paris').create();
  Logger.log('✅ Bot relances paiements installé (lundi 9h). Réglages ajoutés dans l\'onglet Config.');
}

/* ----- Tunnel du site : photos avant chantier + note de dimensionnement ----- */
function leadFind(req) {
  var tel = normTel(req.tel); if (tel.length < 10) throw new Error('Téléphone manquant');
  var r = readAll().filter(function (x) { return !x.data._deleted && normTel(x.data.tel).slice(-9) === tel.slice(-9); })
    .sort(function (a, b) { return (b.data.updatedAt || 0) - (a.data.updatedAt || 0); })[0];
  if (!r) throw new Error('Dossier introuvable : refaites la simulation');
  if (r.data.leadToken && req.token !== r.data.leadToken) throw new Error('Lien expiré : refaites la simulation');
  return r.data;
}
var PHOTO_LABELS = {'chaudiere': 'Chaudière actuelle', 'unite-exterieure': 'Emplacement unité extérieure', 'radiateur': 'Radiateur', 'tableau-electrique': 'Tableau / compteur', 'plaque-chaudiere': 'Plaque chaudière'};
function actionLeadPhotos(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var d = leadFind(req); var ph = (req.photos || [])[0]; if (!ph || !ph.dataUrl) throw new Error('Photo manquante');
    if (String(ph.dataUrl).length > 4000000) throw new Error('Photo trop lourde');
    if ((d.photos || []).filter(function (p) { return p.source === 'site'; }).length >= 12) throw new Error('Nombre maximum de photos atteint');
    var folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PHOTOS_ID'));
    var sub = clientFolder(folder, nomComplet(d) + ' — ' + (d.ville || d.cp || ''));
    var kind = String(ph.kind || 'photo').replace(/[^a-z\-]/g, '').slice(0, 30);
    var f = sub.createFile(dataUrlToBlob(ph.dataUrl, 'site_' + kind + '_' + Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyyMMdd_HHmmss') + extOf(ph.dataUrl)));
    (d.photos = d.photos || []).push({id: 'ws' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), kind: 'avant', label: PHOTO_LABELS[kind] || kind, source: 'site', fileId: f.getId(), thumbId: f.getId(), url: f.getUrl(), t: Date.now(), by: 'client', pending: false});
    d.updatedAt = Date.now(); var sh = dbSheet(); writeRecord(sh, indexRows(sh), d);
    return {ok: true};
  } finally { lock.releaseLock(); }
}
function actionLeadDim(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var d = leadFind(req); var dim = req.dimensionnement || {}; var pd = req.preDevis || {};
    var keep = {}; Object.keys(dim).slice(0, 60).forEach(function (k) { var v = dim[k]; keep[k] = typeof v === 'string' ? v.slice(0, 200) : v; });
    d.dimensionnement = keep; d.preDevis = {prixTTC: pd.prixTTC, cee: pd.cee, mpr: pd.mpr, racSignature: pd.racSignature, racFinal: pd.racFinal, puissance: pd.puissance, produit: String(pd.produit || '').slice(0, 200), t: Date.now()};
    function fill(k, v) { if (v !== undefined && v !== null && v !== '' && (d[k] === undefined || d[k] === '' || d[k] === null)) d[k] = v; }
    fill('surface', dim.surface); fill('materiel', pd.produit);
    fill('compteur', dim.compteur === 'mono' ? 'Mono' : dim.compteur === 'tri' ? 'Tri' : '');
    fill('emetteurs', ({fonte: 'Radiateurs fonte', 'acier-alu': 'Radiateurs acier/alu', plancher: 'Plancher chauffant', mixte: 'Radiateurs + plancher'})[dim.emetteurs]);
    if (pd.cee) fill('valoCEE', String(pd.cee)); if (pd.mpr) fill('valoMPR', String(pd.mpr));
    d.notes = ((d.notes || '') + '\n[Note de dimensionnement site ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM HH:mm') + '] déperditions ' + (dim.deperditionsKw || '?') + ' kW → PAC ' + (pd.puissance || '?') + ' kW · ' + (pd.produit || '') +
      ' · pré-devis ' + (pd.prixTTC || '?') + ' € TTC, RAC signature ' + (pd.racSignature == null ? '?' : pd.racSignature) + ' €, RAC final ' + (pd.racFinal == null ? '?' : pd.racFinal) + ' €' + ((dim.alertes || []).length ? ' · ALERTES : ' + dim.alertes.join(' ; ') : '')).trim();
    if (Number(d.etape) <= 2) { d.etape = 3; (d.historique = d.historique || []).push({t: Date.now(), etape: 3, pv: PV}); }
    d.dossierSiteComplet = true; d.updatedAt = Date.now(); var sh = dbSheet(); writeRecord(sh, indexRows(sh), d);
    var nbPh = (d.photos || []).filter(function (p) { return p.source === 'site'; }).length;
    notifyAdmin('📐 Dossier complet depuis le site : ' + nomComplet(d) + ' — PAC ' + (pd.puissance || '?') + ' kW',
      nomComplet(d) + ' (' + (d.tel || '') + ', ' + (d.cp || '') + ' ' + (d.ville || '') + ') a terminé le tunnel.\n\nDéperditions : ' + (dim.deperditionsKw || '?') + ' kW\nPuissance conseillée : ' + (pd.puissance || '?') + ' kW — ' + (pd.produit || '') +
      '\nPré-devis : ' + (pd.prixTTC || '?') + ' € TTC · CEE ' + (pd.cee == null ? '?' : pd.cee) + ' € · MPR ' + (pd.mpr == null ? '?' : pd.mpr) + ' €\nReste à charge : ' + (pd.racSignature == null ? '?' : pd.racSignature) + ' € à la signature, ' + (pd.racFinal == null ? '?' : pd.racFinal) + ' € final\nPhotos reçues : ' + nbPh +
      ((dim.alertes || []).length ? '\n⚠️ ' + dim.alertes.join('\n⚠️ ') : '') + '\n\n→ Valide la note dans le CRM et envoie le devis définitif.');
    return {ok: true};
  } finally { lock.releaseLock(); }
}

/* ============================================================
   MARCO — relance des devis non signés (envoi manuel, après « oui » de Sassi)
   apercuRelanceDevis() : envoie UN exemple à Sassi + la liste des destinataires.
   envoyerRelanceDevis() : envoie à tous (1 seule fois par client, 80 max par lancement).
   ============================================================ */
function relanceDevisCibles() {
  return readAll().map(function (r) { return r.data; }).filter(function (d) {
    var e = Number(d.etape);
    return !d._deleted && e >= 1 && e <= 4 && d.email && /@/.test(d.email) && !d.relanceDevisLe && !d.stopMails &&
      ((d.devis && d.devis.length) || d.devisTTC);
  });
}
function relanceDevisMail(d) {
  var prenom = d.prenom ? d.prenom.charAt(0).toUpperCase() + d.prenom.slice(1).toLowerCase() : '';
  var civ = prenom ? 'Bonjour ' + prenom : 'Bonjour';
  var sujet = 'Votre projet de pompe à chaleur — on fait le point ?';
  var texte = civ + ',\n\nVous nous aviez demandé un devis pour une pompe à chaleur' + (d.ville ? ' à ' + d.ville : '') + '. Où en est votre projet ?\n\n' +
    'Les aides 2026 (prime CEE et MaPrimeRénov\') sont toujours là, et nous les déduisons directement de votre facture : vous ne payez que le reste.\n\n' +
    'Pour remettre votre devis à jour, il suffit de :\n- répondre à ce mail avec un créneau pour vous rappeler,\n- ou nous appeler au 09 72 14 91 62,\n- ou refaire la simulation en 2 minutes sur https://eco-roy.fr/eligibilite.html\n\n' +
    'Pose en 1 à 2 jours par nos équipes RGE, dossiers d\'aides gérés de A à Z.\n\nBien cordialement,\nL\'équipe Eco Roy\nROY CONSULTING — 37 av. de Saint-Mandé, 75012 Paris — 09 72 14 91 62\n\n' +
    'Vous ne souhaitez plus recevoir nos messages ? Répondez simplement « STOP ».';
  var html = '<div style="font-family:Arial,sans-serif;font-size:15px;color:#1c2620;max-width:600px;line-height:1.5">' +
    '<p>' + civ + ',</p><p>Vous nous aviez demandé un devis pour une pompe à chaleur' + (d.ville ? ' à ' + d.ville : '') + '. <b>Où en est votre projet ?</b></p>' +
    '<p>Les aides 2026 (prime CEE et MaPrimeRénov\') sont toujours là, et nous les <b>déduisons directement de votre facture</b> : vous ne payez que le reste.</p>' +
    '<p style="margin:22px 0"><a href="https://eco-roy.fr/eligibilite.html" style="background:#e9a23b;color:#14532d;font-weight:bold;text-decoration:none;padding:12px 22px;border-radius:24px">Mettre à jour mes aides</a></p>' +
    '<p>Ou répondez simplement à ce mail avec un créneau pour vous rappeler, ou appelez-nous au <b>09 72 14 91 62</b>.</p>' +
    '<p>Pose en 1 à 2 jours par nos équipes RGE, dossiers d\'aides gérés de A à Z.</p>' +
    '<p>Bien cordialement,<br><b>L\'équipe Eco Roy</b><br><span style="color:#667085;font-size:13px">ROY CONSULTING — 37 av. de Saint-Mandé, 75012 Paris — RGE QualiPAC</span></p>' +
    '<p style="color:#98a2b3;font-size:12px">Vous ne souhaitez plus recevoir nos messages ? Répondez simplement « STOP ».</p></div>';
  return {sujet: sujet, texte: texte, html: html};
}
function apercuRelanceDevis() {
  var c = getConfig(); var list = relanceDevisCibles();
  var ex = relanceDevisMail(list[0] || {prenom: 'Jean', ville: 'Évreux'});
  MailApp.sendEmail({to: c.EMAIL_ALERTES, subject: '[APERÇU — rien n\'est parti] ' + ex.sujet, body: ex.texte, htmlBody: ex.html +
    '<hr><p><b>' + list.length + ' clients recevraient ce mail :</b><br>' + list.map(function (d) { return nomComplet(d) + ' — ' + (d.ville || '') + ' (' + (d.source || '') + ')'; }).join('<br>') + '</p>'});
  return list.length;
}
function envoyerRelanceDevis() {
  var sh = dbSheet(); var index = indexRows(sh); var n = 0;
  relanceDevisCibles().slice(0, 80).forEach(function (d) {
    var m = relanceDevisMail(d);
    try {
      MailApp.sendEmail({to: d.email, subject: m.sujet, body: m.texte, htmlBody: m.html, name: 'Eco Roy', replyTo: getConfig().EMAIL_ALERTES});
      d.relanceDevisLe = Date.now(); (d.historique = d.historique || []).push({t: Date.now(), note: 'Relance devis par mail (Marco)'});
      d.updatedAt = Date.now(); writeRecord(sh, index, d); n++;
    } catch (e) { Logger.log(nomComplet(d) + ' : ' + e); }
  });
  notifyAdmin('📨 Marco : ' + n + ' relances devis envoyées', n + ' clients ont reçu la relance de devis. Les réponses arrivent dans ta boîte contact@roy-consulting.com.');
  return n;
}

/* Mise à jour du 9/11/2026 : relevé Qhare / portail mandataire MPR + réponses de Sassi. Lancer une fois. */
function majQhare20261109() {
  var A = '⚠️ ', N = '[Relevé Qhare 09/11] ';
  var U = {
    'mw54y4x2xwkpmg': {etape: 8, refMPR: 'MPR-2026-139541', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-07-29', note: 'Posé. Subvention MPR 5 000 € accordée, demande de solde PAS déposée. ' + A + 'Facture FA855260727-1 datée du 27/07 = avant le devis (29/07) et avant le dépôt MPR (29/07) : dates confirmées réelles par Sassi → ne pas déposer le solde tel quel, avis à prendre (risque de refus / contrôle).'},
    'mw54y4x2c3fupd': {etape: 8, nom: 'Karbal', prenom: 'Hicham', refMPR: 'MPR-2026-158514', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-09-14', note: 'Posé. Subvention MPR 5 000 € accordée, solde à déposer. Facture FA855260911-2 (devis 04/09, facture 11/09) OK. ' + A + 'Reste à charge 1 600 € non payé : offert contre encart pub (Partenaire Vitrine) → la remise doit figurer sur la facture et les aides sont plafonnées (écrêtement) : à régler avant le dépôt du solde.'},
    'mw54y4x2etxe46': {etape: 8, refMPR: 'MPR-2026-163369', valoMPR: '4500', statutMPR: 'Octroyé', dateDepotMPR: '2026-09-23', note: 'Posé. Subvention MPR 4 500 € accordée, demande de solde PAS déposée. ' + A + 'Facture FA855260914-1 du 14/09 = avant le devis (23/09) et avant le dépôt MPR (23/09) : dates réelles → ne pas déposer tel quel, avis à prendre.'},
    'mw54y6zhac8vn7': {etape: 8, refMPR: 'MPR-2026-165821', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-09-28', note: 'Posé. Subvention MPR 5 000 € accordée, solde à déposer. Facture PAC FA855260918-1 (devis 15/09, facture 18/09) OK + facture 500 € FA855261007-1 à clarifier. ' + A + 'Reste à charge 1 600 € offert contre encart pub (Partenaire Vitrine) : remise à faire figurer sur la facture, aides plafonnées.'},
    'mw54y4x2sg5zo3': {etape: 8, note: 'Doublon de « Wasseh Habibi Adam » (dossier principal).'},
    'mw57ma753swnfa': {etape: 8, refMPR: 'MPR-2026-168463', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-10-03', note: 'Posé. Subvention MPR 5 000 € accordée (Botros / Drevenakova), solde à déposer. Facture FA855260923-1 (devis 17/09, facture 23/09) OK. ' + A + 'Reste à charge 1 600 € offert contre encart pub (Partenaire Vitrine) : remise à faire figurer sur la facture, aides plafonnées.'},
    'mw54y4x2jlvac4': {etape: 8, note: 'Doublon de « BOTROS MAGDY » (dossier principal avec devis et MPR).'},
    'mw54y5ofu5660i': {note: 'Co-titulaire du dossier BOTROS MAGDY (MPR-2026-168463).'},
    'mw54y4x2vbt0fw': {etape: 8, refMPR: 'MPR-2026-170254', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-10-07', note: 'Posé. Subvention MPR 5 000 € accordée, demande de solde PAS déposée. ' + A + 'Facture FA855260925-1 du 25/09 = avant le devis (07/10) et avant le dépôt MPR (07/10) : dates réelles → ne pas déposer tel quel, avis à prendre.'},
    'mw54y6zhf0ow0w': {etape: 8, note: 'Doublon probable de « Danelutti » (Christian Gérard, Val-de-Virieu) : dossier principal = Danelutti.'},
    'mw54y4x2rq9061': {etape: 8, refMPR: 'MPR-2026-170303', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-10-07', note: 'Posé. Subvention MPR 5 000 € accordée, demande de solde PAS déposée. ' + A + 'Facture FA855260722-2 du 22/07 = avant le devis (07/10) et avant le dépôt MPR (07/10) : dates réelles → ne pas déposer tel quel, avis à prendre.'},
    'mw54y4x3qr20c9': {etape: 5, refMPR: 'MPR-2026-171868', valoMPR: '3000', statutMPR: 'Octroyé', dateDepotMPR: '2026-10-09', note: 'PAS ENCORE POSÉ (normal, confirmé par Sassi). Subvention MPR 3 000 € accordée. ' + A + 'Qhare : date d\'émission de facture manquante sur le devis → à corriger avant la demande de solde.'},
    'mw54y4x3pvg8ot': {etape: 8, refMPR: 'MPR-2026-175994', valoMPR: '5000', statutMPR: 'Octroyé', dateDepotMPR: '2026-10-16', note: 'Posé (confirmé par Sassi). Subvention MPR 5 000 € accordée, solde à déposer. ' + A + 'Facture FA855261016-1 datée du 16/10 = même jour que le dépôt MPR : vérifier que la pose a bien eu lieu APRÈS le dépôt. Encore en « Lead » dans Qhare : mettre à jour.'},
    'mw57ma798vpkp9': {etape: 8, refMPR: 'MPR-2026-178368', valoMPR: '4000', statutMPR: 'Octroyé', dateDepotMPR: '2026-10-20', note: 'Posé (confirmé par Sassi). Subvention MPR 4 000 € accordée, solde à déposer. ' + A + 'Pas de facture dans Qhare : facture à établir (date postérieure au dépôt MPR du 20/10 et à la pose).'},
    'mw54y4x28s30u3': {etape: 8, refMPR: 'MPR-2026-121370', valoMPR: '4800', statutMPR: 'Paiement demandé', note: 'Solde MPR confirmé par l\'Anah : 4 800 € à payer, rien à faire.'},
    'mw54y4x2ybkcez': {etape: 8, refMPR: 'MPR-2026-87704', valoMPR: '6200', statutMPR: 'Paiement demandé', note: 'Solde MPR confirmé par l\'Anah : 6 200 € à payer, rien à faire.'},
    'mw54y4x1xco7cy': {etape: 8, refMPR: 'MPR-2026-113564', valoMPR: '5000', statutMPR: 'Paiement demandé', note: 'Solde MPR confirmé par l\'Anah : 5 000 € à payer, rien à faire.'},
    'mw54y4x21gkoje': {etape: 8, refMPR: 'MPR-2026-127536', valoMPR: '4000', statutMPR: 'Paiement demandé', note: 'Solde MPR confirmé par l\'Anah : 4 000 € à payer, rien à faire.'},
    'mw54y4x1c2mole': {etape: 8, refMPR: 'MPR-2026-132088', valoMPR: '5000', statutMPR: 'Paiement demandé', note: 'Solde MPR confirmé par l\'Anah : 5 000 € à payer, rien à faire.'},
    'mw54y4x2kdaomh': {etape: 8, refMPR: 'MPR-2026-139210', valoMPR: '6200', statutMPR: 'Paiement demandé', note: 'Solde MPR confirmé par l\'Anah : 6 200 € à payer, rien à faire.'},
    'mw54y5oen2wuh8': {etape: 8, refMPR: 'MPR-2025-401909', valoMPR: '12500', statutMPR: 'Paiement demandé', note: 'Solde MPR 12 500 € déposé (dossier MEBAREK / BOUKHALFA), en instruction à l\'Anah : à surveiller.'}
  };
  var sh = dbSheet(); var index = indexRows(sh); var n = 0, manquants = [];
  var byId = {}; readAll().forEach(function (r) { byId[r.data.id] = r.data; });
  Object.keys(U).forEach(function (id) {
    var d = byId[id]; if (!d) { manquants.push(id); return; }
    var u = U[id];
    ['etape', 'nom', 'prenom', 'refMPR', 'valoMPR', 'statutMPR', 'dateDepotMPR'].forEach(function (k) { if (u[k] !== undefined) d[k] = u[k]; });
    if (u.etape !== undefined) (d.historique = d.historique || []).push({t: Date.now(), etape: u.etape});
    d.notes = ((d.notes || '') + '\n' + N + u.note).trim();
    if (/⚠️/.test(u.note)) d.alerteConformite = true;
    d.updatedAt = Date.now(); writeRecord(sh, index, d); n++;
  });
  Logger.log(n + ' dossiers mis à jour' + (manquants.length ? ' — introuvables : ' + manquants.join(', ') : ''));
  return n;
}

function majQhare20261109b() {
  var sh = dbSheet(); var index = indexRows(sh); var byId = {}; readAll().forEach(function (r) { byId[r.data.id] = r.data; });
  var b = byId['mw57ma798vpkp9']; if (b) { b.etape = 6; b.datePose = '2026-11-14'; (b.historique = b.historique || []).push({t: Date.now(), etape: 6}); b.notes = (b.notes || '') + '\n[09/11] Correction Sassi : PAS ENCORE POSÉ. Pose reportée au samedi 14/11 faute de matériel. Facture à établir après la pose.'; b.updatedAt = Date.now(); writeRecord(sh, index, b); }
  var m = byId['mw54y4x3pvg8ot']; if (m) { m.notes = (m.notes || '') + '\n[09/11] Sassi confirme : pose APRÈS le dépôt MPR du 16/10. Facture à dater au jour de fin de pose (pas le 16/10) avant dépôt du solde.'; m.updatedAt = Date.now(); writeRecord(sh, index, m); }
  Logger.log('ok ' + !!b + ' ' + !!m);
}

/* Tri des Google Sheets (call maroc mabrour pac, igor install, Suivi clients) vs CRM — 9/11/2026.
   Règle : doublon = même personne + même adresse + même chantier. Rien n'est supprimé ni déplacé :
   chaque cas passe dans la file « À vérifier » avec une note. Sassi valide, puis on applique. */
function triDrive20261109() {
  var T = '[Tri Drive 09/11] ';
  var U = {
    // Doublons proposés
    'mw54y4x2pnal4c': 'DOUBLON PROPOSÉ de « Boukhalfa » (fiche MPR 12 500 €). Même chantier ITE 18 rue du Havre Béthune. Si OK → fusion (date de pose 25/07 recopiée) puis Perdu.',
    'mw54y4x21odhvp': 'DOUBLON PROPOSÉ de « Boukhalfa ». Attention : marge -2 500 € (poseur Stella) = coût à reporter sur la fiche principale avant fusion.',
    'mw54y5oen2wuh8': 'FICHE PRINCIPALE Boukhalfa (MPR-2025-401909). ⚠️ « Capelle Valerie » = même adresse/tél mais AUTRE chantier (PAC CEE) → on garde les 2. « Kostov » = tél du pro mis par le call → sans lien.',
    'mw54y5oexc74qx': 'DOUBLON PROPOSÉ de « Mohamed Malki » (9 rue des Pâquerettes, Le Blanc-Mesnil). Si OK → Perdu.',
    'mw54y4x2goddpn': 'FICHE PRINCIPALE Malki. NB : nom/prénom inversés (nom = Malki, prénom = Mohamed).',
    'mw54y5ofvavsb6': 'DOUBLON PROPOSÉ de « Ayoub Marchiche » (Clermont-Ferrand). Si OK → Perdu.',
    'mw54y4x2cclnuy': 'FICHE PRINCIPALE Marchiche. NB : nom/prénom inversés (nom = Marchiche, prénom = Ayoub).',
    'mw54y5of9rtngf': 'DOUBLON PROPOSÉ de « De Oliveira Amina » (Vernouillet). Si OK → Perdu.',
    'mw57ma72wyfmys': 'DOUBLON PROPOSÉ de « Dejeddi Ourabah » (17 rue Charles Derenne, Villeneuve-sur-Lot). Tél différent (770… vs 767…) : vérifier le bon numéro. Si OK → Perdu.',
    'mw54y4x2sg5zo3': 'DOUBLON PROPOSÉ de « Wasseh Habibi Adam » (Morangis). Si OK → Perdu.',
    'mw54y4x2jlvac4': 'DOUBLON PROPOSÉ de « BOTROS MAGDY » (Créteil). Drevencova = co-titulaire, même chantier. Si OK → Perdu.',
    'mw54y6zhf0ow0w': 'DOUBLON PROPOSÉ de « Danelutti » (Val-de-Virieu) — relevé Qhare. À confirmer : Christian Gérard = Danelutti ?',
    // Cas à trancher par Sassi
    'mw54y4x04szr1d': 'À TRANCHER : 2 fiches « Santos ». Celle-ci a une marge -2 998 € (poseur Stella) = coût du même chantier ou autre chantier ?',
    'mw54y4x0afioa2': 'À TRANCHER : 2 fiches « Santos » (voir l\'autre, marge -2 998 €).',
    'mw54y4x18999ye': 'À TRANCHER : « De Brosse » (soldé) = même client que « TANGUY DE BROSSE » (Saint-Cyr-l\'École, installé) ? Même chantier ou 2 chantiers ?',
    'mw54y4x18uicfn': 'À TRANCHER : « Didier » (soldé) = « DIDIER Marcel » (Viry-Châtillon, installé) ? « DIDIER Cristophe » (même ville, parrainage) est une autre personne.',
    'mw54y4x32h8tv6': 'À TRANCHER : Farid Bennadi a 2 chantiers (fichier Igor) : Vouziers 15 rue de Tabure (SCI) et Buzancy 10 rue Professeur Gilbert (locataire). Cette fiche = lequel ?',
    'mw54y6f48255p6': 'Chantier Farid Bennadi — VOUZIERS (15 rue de Tabure, SCI). L\'autre chantier = Buzancy (locataire).',
    'mw54y5ofkbite5': 'DOUBLON PROPOSÉ : « Farid » sans infos (Suivi clients) = Farid Bennadi. Si OK → Perdu.',
    'mw54y4x1kf3067': 'PAS UN DOUBLON (proposé) : « Cailleux/ssc » = 2e chantier (SSC) du même client que « Cailleux ». On garde les 2.',
    'mw54y6zdcqfgrm': 'PAS UN DOUBLON : même tél que Diouane mais autre personne/adresse (Noyelles-sous-Lens). On garde.',
    // Fiches test / incomplètes
    'mw77mpcc2kvi3f': 'TEST du formulaire du site (sassi ilan). Si OK → Perdu.',
    'mw89a3tg8ik8wf': 'TEST du formulaire du site (Mmm Margot). Si OK → Perdu.',
    'mw9732ipwuiy0y': 'TEST du formulaire du site (CONSULTING ROY). Si OK → Perdu.',
    'mw57ma5hsksgl1': 'INCOMPLET : « Koreak grill bar » (pro, Levallois) — téléphone faux (101010102). À compléter ou Perdu.',
    'mw57ma6f2wbp8d': 'INCOMPLET : adresse saisie à la place du nom (101 av. Saint-Mandé, Paris 12e). Nom du client à compléter.'
  };
  var sh = dbSheet(); var index = indexRows(sh); var byId = {}; readAll().forEach(function (r) { byId[r.data.id] = r.data; });
  var n = 0, manquants = [];
  Object.keys(U).forEach(function (id) {
    var d = byId[id]; if (!d) { manquants.push(id); return; }
    d.notes = ((d.notes || '') + '\n' + T + U[id]).trim(); d.aVerifier = true; d.updatedAt = Date.now(); writeRecord(sh, index, d); n++;
  });
  // Chantier manquant : Diouane n°2 (Harnes), présent dans « call maroc mabrour pac », absent du CRM
  var deja = Object.keys(byId).some(function (k) { var d = byId[k]; return /diouane/i.test(d.nom || '') && /harnes|62440/i.test((d.ville || '') + ' ' + (d.cp || '') + ' ' + (d.adresse || '')); });
  if (!deja) {
    var x = newDossier(); x.source = 'Call Maroc'; x.nom = 'Diouane'; x.prenom = 'Mohamed'; x.tel = '627664541';
    x.adresse = '7B rue du Petit Moulin'; x.cp = '62440'; x.ville = 'Harnes'; x.couleur = 'bleu'; x.habitat = 'Maison individuelle';
    x.notes = T + '2e CHANTIER de M. Diouane (le 1er = 102 rue Victor Mathé, Hénin-Beaumont). 80 m², 5 radiateurs, Wolf 8-10. Source : fichier call maroc. À valider.';
    x.aVerifier = true; writeRecord(sh, index, x); n++;
  }
  Logger.log(n + ' fiches dans « À vérifier »' + (manquants.length ? ' — introuvables : ' + manquants.join(', ') : ''));
  return n;
}

/* Réponses de Sassi au tri Drive — 9/11/2026 22:56 */
function decisionsTri20261109() {
  var T = '[Décision Sassi 09/11] ';
  var U = {
    'mw54y4x0afioa2': {note: '2 CHANTIERS « Santos » confirmés : pas un doublon. On garde les 2.', ok: true},
    'mw54y4x04szr1d': {note: '2 CHANTIERS « Santos » confirmés : pas un doublon. On garde les 2.', ok: true},
    'mw54y4x18999ye': {note: 'MÊME CLIENT que « TANGUY DE BROSSE » (Saint-Cyr-l\'École) → à fusionner dans cette fiche-là.', ok: false},
    'mw57ma4w6c3s31': {note: 'FICHE PRINCIPALE De Brosse (même client que la fiche « De Brosse » soldée → fusion à faire).', ok: false},
    'mw54y4x18uicfn': {note: 'MÊME CLIENT que « DIDIER Marcel » (Viry-Châtillon) → à fusionner dans cette fiche-là.', ok: false},
    'mw57ma55i619dt': {note: 'FICHE PRINCIPALE Didier Marcel (même client que la fiche « Didier » soldée → fusion à faire).', ok: false},
    'mw54y4x32h8tv6': {note: 'Farid Bennadi a PLUSIEURS MAISONS : fiche gardée comme chantier séparé (Buzancy ou autre adresse à compléter).', ok: true},
    'mw54y6f48255p6': {note: 'Farid Bennadi — plusieurs maisons confirmées. Chantier Vouziers gardé.', ok: true}
  };
  var sh = dbSheet(); var index = indexRows(sh); var byId = {}; readAll().forEach(function (r) { byId[r.data.id] = r.data; });
  var n = 0, manquants = [];
  Object.keys(U).forEach(function (id) {
    var d = byId[id]; if (!d) { manquants.push(id); return; }
    d.notes = ((d.notes || '') + '\n' + T + U[id].note).trim(); d.aVerifier = !U[id].ok; d.updatedAt = Date.now(); writeRecord(sh, index, d); n++;
  });
  Logger.log(n + ' fiches mises à jour' + (manquants.length ? ' — introuvables : ' + manquants.join(', ') : ''));
}

/* Fusion des doublons validée par Sassi — 9/11/2026 23:05. Rien n'est supprimé : le doublon passe en « Perdu » (étape 0) avec la note « fusionné dans… ». */
function fusionDoublons20261109() {
  var T = '[Fusion 09/11] ';
  var P = [ // [doublon, principale]
    ['mw54y4x2pnal4c', 'mw54y5oen2wuh8'], ['mw54y4x21odhvp', 'mw54y5oen2wuh8'], ['mw54y5oexc74qx', 'mw54y4x2goddpn'],
    ['mw54y5ofvavsb6', 'mw54y4x2cclnuy'], ['mw54y5of9rtngf', 'mw54y4x2vh3sl4'], ['mw57ma72wyfmys', 'mw54y4x20ngilg'],
    ['mw54y4x2sg5zo3', 'mw54y6zhac8vn7'], ['mw54y4x2jlvac4', 'mw57ma753swnfa'], ['mw54y6zhf0ow0w', 'mw54y4x2vbt0fw'],
    ['mw54y5ofkbite5', 'mw54y4x32h8tv6'], ['mw54y4x18999ye', 'mw57ma4w6c3s31'], ['mw54y4x18uicfn', 'mw57ma55i619dt']
  ];
  var CHAMPS = ['prenom', 'tel', 'email', 'adresse', 'cp', 'ville', 'couleur', 'habitat', 'datePose', 'poseur', 'fournisseur', 'materiel', 'refMPR', 'refOPX'];
  var TESTS = ['mw77mpcc2kvi3f', 'mw89a3tg8ik8wf', 'mw9732ipwuiy0y'];
  var sh = dbSheet(); var index = indexRows(sh); var byId = {}; readAll().forEach(function (r) { byId[r.data.id] = r.data; });
  var touches = {}, n = 0, manquants = [];
  P.forEach(function (p) {
    var dup = byId[p[0]], main = byId[p[1]]; if (!dup || !main) { manquants.push(p.join('>')); return; }
    var copies = [];
    CHAMPS.forEach(function (k) { if ((main[k] == null || main[k] === '') && dup[k] != null && dup[k] !== '') { main[k] = dup[k]; copies.push(k); } });
    main.notes = ((main.notes || '') + '\n' + T + 'Fiche « ' + nomComplet(dup) + ' » fusionnée ici' + (copies.length ? ' (repris : ' + copies.join(', ') + ')' : '') + '. Marge de l\'ancienne fiche : ' + (dup.margePrev != null ? dup.margePrev : '?') + ' €.').trim();
    dup.notes = ((dup.notes || '') + '\n' + T + 'DOUBLON — fusionné dans « ' + nomComplet(main) + ' ». Fiche gardée pour l\'historique.').trim();
    dup.etape = 0; (dup.historique = dup.historique || []).push({t: Date.now(), etape: 0}); dup.aVerifier = false; dup.updatedAt = Date.now();
    writeRecord(sh, index, dup); touches[main.id] = main; n++;
  });
  var mk = byId['mw54y4x2goddpn']; if (mk) { mk.nom = 'Malki'; mk.prenom = 'Mohamed'; touches[mk.id] = mk; }
  var mc = byId['mw54y4x2cclnuy']; if (mc) { mc.nom = 'Marchiche'; mc.prenom = 'Ayoub'; touches[mc.id] = mc; }
  Object.keys(touches).forEach(function (id) {
    var d = touches[id];
    // Boukhalfa : coût -2 500 € (poseur Stella) à reporter à la main → reste « à vérifier »
    d.aVerifier = (id === 'mw54y5oen2wuh8');
    if (id === 'mw54y5oen2wuh8') d.notes += '\n' + T + 'À FAIRE : reporter le coût poseur Stella (-2 500 €) dans l\'onglet Chantier.';
    d.updatedAt = Date.now(); writeRecord(sh, index, d);
  });
  TESTS.forEach(function (id) { var d = byId[id]; if (!d) return; d.etape = 0; (d.historique = d.historique || []).push({t: Date.now(), etape: 0}); d.aVerifier = false; d.notes = ((d.notes || '') + '\n' + T + 'Fiche TEST du formulaire → Perdu.').trim(); d.updatedAt = Date.now(); writeRecord(sh, index, d); n++; });
  Logger.log(n + ' fiches passées en Perdu, ' + Object.keys(touches).length + ' fiches principales complétées' + (manquants.length ? ' — introuvables : ' + manquants.join(', ') : ''));
}

/* Sassi 09/11 23:47 : Koreak grill bar et '101 av. Saint-Mandé' jamais installés */
function jamaisInstalles20261109() {
  var sh = dbSheet(); var index = indexRows(sh); var byId = {}; readAll().forEach(function (r) { byId[r.data.id] = r.data; });
  ['mw57ma5hsksgl1', 'mw57ma6f2wbp8d'].forEach(function (id) { var d = byId[id]; if (!d) return; d.notes = ((d.notes || '') + '\n[Décision Sassi 09/11] Jamais installé.').trim(); d.aVerifier = false; d.updatedAt = Date.now(); writeRecord(sh, index, d); });
  Logger.log('ok');
}
/* ============================================================
   PIPELINE v2 — AUTOMATISATIONS (phase 2) · 10/11/2026
   Marco (accusé, NRP, devis), Paula (rappels de pose), Stella (avis), Stan (entretien) : e-mails + SMS aux clients, 1 seule fois par message.
   Sécurité : rien ne part tant que AUTO_MAILS ≠ oui (onglet Config).
   En attendant, apercuAutomatisations() envoie à Sassi la liste + les textes.
   SMS : Brevo, actif seulement si SMS_ACTIF = oui ET BREVO_API_KEY rempli (par Sassi).
   ============================================================ */
var AUTO_START = new Date('2026-11-10T00:00:00').getTime();
function ensureConfigV2() {
  ensureConfig();
  var cfg = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('DB_ID')).getSheetByName('Config'); var have = {};
  cfg.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0]) have[String(r[0]).trim()] = true; });
  var add = [
    ['AUTO_MAILS', 'non', 'oui = les e-mails automatiques partent aux clients (accusé, NRP, relances devis, rappels de pose, avis Google, entretien)'],
    ['SMS_ACTIF', 'non', 'oui = les SMS partent aussi (il faut BREVO_API_KEY)'],
    ['BREVO_API_KEY', '', 'Clé API Brevo pour les SMS — à coller par Sassi lui-même'],
    ['SMS_EXPEDITEUR', 'EcoRoy', 'Nom affiché sur les SMS (11 caractères max, sans espace)'],
    ['TEL_SOCIETE', '09 72 14 91 62', 'Téléphone affiché dans les messages'],
    ['AVIS_GOOGLE_URL', '', 'Lien direct « laisser un avis » de la fiche Google (après validation de la fiche)'],
    ['LIEN_AGENDA', '', 'Lien de prise de rendez-vous (créneau de rappel / date de pose), facultatif'],
    ['EMAIL_COMPTABLE', '', 'E-mail du comptable pour l\'export mensuel (facultatif)'],
    ['PILOTAGE', 'oui', 'oui = rapport de pilotage (Victor) chaque matin + bilan le lundi']
  ].filter(function (r) { return !have[r[0]]; });
  if (add.length) cfg.getRange(cfg.getLastRow() + 1, 1, add.length, 3).setValues(add);
  CacheService.getScriptCache().remove('cfg');
}
function telInter(t) { var s = String(t || '').replace(/[^\d+]/g, ''); if (/^\+/.test(s)) return s.slice(1); if (/^0\d{9}$/.test(s)) return '33' + s.slice(1); if (/^[67]\d{8}$/.test(s)) return '33' + s; return ''; }
function envoyerSMS(d, txt) {
  var c = getConfig(); if (String(c.SMS_ACTIF).toLowerCase() !== 'oui' || !c.BREVO_API_KEY) return false;
  var tel = telInter(d.tel); if (!tel || !/^33[67]/.test(tel)) return false;
  var r = UrlFetchApp.fetch('https://api.brevo.com/v3/transactionalSMS/sms', {method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {'api-key': c.BREVO_API_KEY, accept: 'application/json'},
    payload: JSON.stringify({sender: (c.SMS_EXPEDITEUR || 'EcoRoy').slice(0, 11), recipient: tel, content: txt, type: 'transactional'})});
  var ok = r.getResponseCode() < 300; (d.comms = d.comms || []).push({t: Date.now(), canal: 'sms', modele: 'auto', txt: (ok ? '' : '⚠ échec SMS : ') + txt});
  return ok;
}
function prenomDe(d) { var p = String(d.prenom || '').trim(); return p ? p.charAt(0).toUpperCase() + p.slice(1).toLowerCase() : ''; }
/** Liste des messages dus maintenant : [{d, cle, sujet, texte, sms}] */
function messagesDus(c) {
  var tel = c.TEL_SOCIETE || '09 72 14 91 62', agenda = c.LIEN_AGENDA || '';
  var sig = '\n\nBien cordialement,\nL\'équipe Eco Roy\nROY CONSULTING — 37 av. de Saint-Mandé, 75012 Paris — ' + tel + '\n\nVous ne souhaitez plus recevoir nos messages ? Répondez « STOP ».';
  var out = [];
  readAll().forEach(function (r) {
    var d = r.data; if (d._deleted || d.oppositionContact || d.stopMails) return;
    var e = Number(d.etape), env = d.envois || {}, bj = prenomDe(d) ? 'Bonjour ' + prenomDe(d) : 'Bonjour';
    function add(cle, sujet, texte, sms) { if (!env[cle]) out.push({d: d, cle: cle, sujet: sujet, texte: bj + ',\n\n' + texte + sig, sms: sms}); }
    // 01 · accusé de réception (leads arrivés depuis la mise en route)
    if (e >= 1 && e <= 2 && Number(d.createdAt) >= AUTO_START && joursDepuis(d.createdAt) <= 2)
      add('accuse', 'Eco Roy — nous avons bien reçu votre demande', 'Merci pour votre demande concernant votre projet de pompe à chaleur. Un conseiller Eco Roy vous rappelle très vite (en général dans l\'heure, du lundi au samedi).' +
        (agenda ? '\n\nVous pouvez aussi choisir vous-même votre créneau de rappel ici : ' + agenda : '') + '\n\nPour gagner du temps, préparez si possible votre dernier avis d\'imposition : il sert à calculer vos aides.',
        'Eco Roy : nous avons bien recu votre demande, un conseiller vous rappelle tres vite au ' + tel + '.' + (agenda ? ' Votre creneau : ' + agenda : ''));
    // 02 · après chaque appel sans réponse
    if (e === 2 && d.statutAppel === 'NRP' && Number(d.nbNRP) > 0)
      add('nrp' + d.nbNRP, 'Eco Roy — nous avons essayé de vous joindre', 'Nous avons essayé de vous joindre au sujet de votre demande de pompe à chaleur, sans succès.\n\nRappelez-nous au ' + tel + (agenda ? ', ou choisissez le créneau qui vous arrange : ' + agenda : ', ou répondez à ce mail avec vos disponibilités') + '.',
        'Eco Roy a essaye de vous joindre pour votre projet de pompe a chaleur. Rappelez-nous au ' + tel + (agenda ? ' ou choisissez un creneau : ' + agenda : '') + '.');
    // 03 · relances devis J+3 / J+7 / J+14
    if (e === 4 && d.statutDevis === 'Envoyé · en attente' && d.dateEnvoiDevis) {
      var jd = joursDepuis(d.dateEnvoiDevis);
      [[3, 'devisJ3'], [7, 'devisJ7'], [14, 'devisJ14']].forEach(function (x) { if (jd >= x[0] && (x[0] === 14 || jd < x[0] + 4))
        add(x[1], 'Votre devis Eco Roy — avez-vous des questions ?', 'Nous vous avons envoyé votre devis il y a ' + jd + ' jours. Avez-vous des questions sur le montant des aides, le matériel ou les délais ?\n\nUn simple appel au ' + tel + ' ou une réponse à ce mail suffit : nous reprenons tout avec vous.',
          x[0] === 7 ? 'Eco Roy : avez-vous pu consulter votre devis pompe a chaleur ? Une question ? ' + tel : null); });
    }
    // 05 · rappels de pose J-3 et J-1
    if (e === 7 && d.datePose) {
      var j = -joursDepuis(d.datePose);
      if (j <= 3 && j >= 2) add('poseJ3', 'Votre installation Eco Roy approche', 'Votre installation de pompe à chaleur est prévue le ' + Utilities.formatDate(new Date(d.datePose), 'Europe/Paris', 'dd/MM/yyyy') + (d.poseur ? ' avec notre équipe (' + d.poseur + ')' : '') + '.\n\nPensez à dégager l\'accès à la chaudière et à l\'emplacement de l\'unité extérieure. En cas d\'empêchement, prévenez-nous au plus vite au ' + tel + '.', null);
      if (j === 1) add('poseJ1', 'Rappel : installation Eco Roy demain', 'Petit rappel : notre équipe vient demain pour installer votre pompe à chaleur. Une personne majeure doit être présente. Pour toute question : ' + tel + '.',
        'Eco Roy : rappel, installation de votre pompe a chaleur demain. Une question ? ' + tel);
    }
    // 06 · après chantier : avis Google (tous les clients, sans filtre) + entretien si satisfait
    var inst = d.dateInstallation || (e >= 8 ? d.datePose : '');
    var posePost = inst && new Date(inst).getTime() >= AUTO_START;
    if (e >= 8 && e <= 11 && posePost && joursDepuis(inst) >= 1 && c.AVIS_GOOGLE_URL)
      add('avis', 'Votre avis compte pour Eco Roy', 'Merci de nous avoir fait confiance pour l\'installation de votre pompe à chaleur.\n\nVotre retour, quel qu\'il soit, nous aide à nous améliorer et aide d\'autres familles à se décider. Il vous suffit d\'1 minute : ' + c.AVIS_GOOGLE_URL, null);
    if (e >= 8 && e <= 11 && posePost && d.satisfaction && Number(d.satisfaction.note) >= 4 && (!d.offreEntretien || d.offreEntretien === 'Non proposé'))
      add('entretien', 'Votre pompe à chaleur : on s\'occupe de l\'entretien ?', 'Votre installation est terminée et nous sommes ravis qu\'elle vous donne satisfaction.\n\nL\'entretien régulier d\'une pompe à chaleur est obligatoire pour garder la garantie du fabricant. Eco Roy vous propose 3 formules, sans engagement de durée :\n- Essentiel : 14,90 €/mois\n- Sérénité : 26,90 €/mois\n- Premium : 39,90 €/mois\n(2 mois offerts si vous payez à l\'année)\n\nTous les détails : https://eco-roy.fr/entretien-pompe-a-chaleur.html\nRépondez simplement à ce mail avec la formule choisie : nous vous envoyons le contrat et le mandat de prélèvement.', null);
    // 09 · suivi long terme : rappel à la date choisie (si consentement)
    if (e === 12 && d.consentementContact && d.rappelLTLe && String(d.rappelLTLe) <= Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd'))
      add('lt' + d.rappelLTLe, 'Eco Roy — où en est votre projet ?', 'Vous nous aviez demandé de revenir vers vous à cette période pour votre projet de pompe à chaleur. Les aides 2026 sont toujours disponibles.\n\nVoulez-vous que nous refassions le point ? Répondez à ce mail ou appelez le ' + tel + '.', null);
  });
  return out;
}
/** Toutes les heures : envoie les messages dus (si AUTO_MAILS = oui). */
function botAutomatisations() {
  var c = getConfig(); if (String(c.AUTO_MAILS).toLowerCase() !== 'oui') return;
  var h = Number(Utilities.formatDate(new Date(), 'Europe/Paris', 'H')); if (h < 9 || h >= 20) return; // pas de message la nuit
  var sh = dbSheet(); var index = indexRows(sh); var n = 0, touches = {};
  messagesDus(c).slice(0, 60).forEach(function (m) {
    var d = touches[m.d.id] || m.d; var ok = false;
    if (d.email && /@/.test(d.email)) { try { MailApp.sendEmail({to: d.email, subject: m.sujet, body: m.texte, name: 'Eco Roy', replyTo: c.EMAIL_ALERTES || ''}); ok = true; (d.comms = d.comms || []).push({t: Date.now(), canal: 'mail', modele: m.cle, txt: m.sujet}); } catch (e) { Logger.log(e); } }
    if (m.sms && envoyerSMS(d, m.sms)) ok = true;
    if (ok || !(d.email && /@/.test(d.email))) { d.envois = d.envois || {}; d.envois[m.cle] = Date.now(); if (m.cle === 'entretien') d.offreEntretien = 'Proposé'; if (m.cle === 'avis') d.avisGoogleDemande = true; d.updatedAt = Date.now(); touches[d.id] = d; n++; }
  });
  Object.keys(touches).forEach(function (id) { writeRecord(sh, index, touches[id]); });
  if (n) Logger.log(n + ' message(s) envoyés');
}
/** À lancer à la main : envoie à Sassi la liste et les textes de ce qui partirait. Rien n'est envoyé aux clients. */
function apercuAutomatisations() {
  ensureConfigV2(); var c = getConfig(); var l = messagesDus(c);
  var ex = {}; l.forEach(function (m) { if (!ex[m.cle.replace(/\d+$/, '')]) ex[m.cle.replace(/\d+$/, '')] = m; });
  var corps = 'Aperçu des messages automatiques (rien n\'a été envoyé aux clients).\n\nAUTO_MAILS = ' + (c.AUTO_MAILS || 'non') + ' · SMS_ACTIF = ' + (c.SMS_ACTIF || 'non') +
    '\n\n' + l.length + ' message(s) partiraient maintenant :\n' + l.map(function (m) { return '- ' + nomComplet(m.d) + ' — ' + m.sujet + (m.d.email ? '' : ' (pas d\'e-mail)') + (m.sms ? ' + SMS' : ''); }).join('\n') +
    '\n\n———— EXEMPLES DE TEXTES ————\n\n' + Object.keys(ex).map(function (k) { return '【' + k + '】 ' + ex[k].sujet + '\n\n' + ex[k].texte + (ex[k].sms ? '\n\nSMS : ' + ex[k].sms : ''); }).join('\n\n————————\n\n');
  notifyAdmin('👀 Aperçu des messages automatiques — ' + l.length + ' message(s)', corps);
  Logger.log(l.length + ' messages en attente');
}

/* ============================================================
   VICTOR — PILOTAGE DE LA SOCIÉTÉ (phase 3)
   Chaque matin 7h30 : argent, échéances, chantiers, leads, risques.
   Le lundi : bilan de la semaine + chiffres par source.
   ============================================================ */
function botPilotage() {
  var c = getConfig(); if (String(c.PILOTAGE || 'oui').toLowerCase() !== 'oui') return;
  var all = readAll().map(function (r) { return r.data; }).filter(function (d) { return !d._deleted; });
  var auj = new Date(); var lundi = auj.getDay() === 1;
  var e = function (d) { return Number(d.etape); };
  var resteCEE = 0, resteMPR = 0, aEmettre = [], echeF = [], echeC = [], risques = [], bloquees = [], attenteLong = [];
  all.forEach(function (d) {
    var et = e(d);
    if (et >= 8 && et <= 11) { resteCEE += Math.max(0, n(d.valoCEE) - n(d.encaisseCEE)); resteMPR += Math.max(0, n(d.valoMPR) - n(d.encaisseMPR)); }
    if (et >= 8 && et <= 9 && !d.dateFacture && !(d.devis || []).some(function (v) { return v.fac; })) aEmettre.push(nomComplet(d));
    if (d.dateFactureFournisseur && !d.fournisseurPayeLe) { var jf = joursDepuis(d.dateFactureFournisseur); if (jf >= 30) echeF.push(nomComplet(d) + ' — ' + (d.fournisseur || '?') + ' ' + euros(n(d.montantFactureFournisseur)) + ' — échéance dans ' + (45 - jf) + ' j'); }
    var ip = d.dateInstallation || d.datePose; if (et >= 8 && et <= 11 && n(d.commission) > 0 && !d.commissionPayeeLe && ip) { var jc = joursDepuis(ip); if (jc >= 14 && jc <= 60) echeC.push(nomComplet(d) + ' — ' + (d.regie || d.source || '?') + ' ' + euros(n(d.commission)) + ' — pose il y a ' + jc + ' j (max 30)'); }
    var seq = [d.dateSignatureDevis, d.dateDepotMPR, d.datePose, d.dateFacture].filter(Boolean);
    for (var i = 1; i < seq.length; i++) if (String(seq[i]) < String(seq[i - 1])) { risques.push(nomComplet(d) + ' — dates dans le désordre (devis/dépôt MPR/pose/facture)'); break; }
    if (et === 7) { var mprOk = ['Déposé', 'Octroyé', 'Paiement demandé', 'Payé', 'Non concerné'].indexOf(d.statutMPR) >= 0 || couleurCalc(d.rfr, d.nbPersonnes, d.cp) === 'rose'; if (!mprOk) bloquees.push(nomComplet(d) + ' — pose programmée le ' + (d.datePose || '?') + ' sans dépôt MPR'); }
    if (et === 10 && joursDepuis(d.dateDemandeSolde) >= 45) attenteLong.push(nomComplet(d) + ' — ' + euros(resteAEncaisser(d)) + ' attendu depuis ' + joursDepuis(d.dateDemandeSolde) + ' j');
  });
  var S = function (t, l) { return l.length ? '\n\n' + t + ' (' + l.length + ')\n- ' + l.slice(0, 15).join('\n- ') + (l.length > 15 ? '\n- … +' + (l.length - 15) : '') : ''; };
  var nouv7 = all.filter(function (d) { return joursDepuis(d.createdAt) <= 7; });
  var body = 'Bonjour Sassi,\n\nPILOTAGE ECO ROY — ' + Utilities.formatDate(auj, 'Europe/Paris', 'EEEE dd/MM/yyyy') +
    '\n\n💰 ARGENT À ENCAISSER : ' + euros(resteCEE + resteMPR) + ' (CEE ' + euros(resteCEE) + ' · MPR ' + euros(resteMPR) + ')' +
    S('🧾 Factures à émettre', aEmettre) + S('🏭 Fournisseurs à payer (45 j)', echeF) + S('🤝 Commissions à payer', echeC) + S('⏳ Paiements en retard (> 45 j)', attenteLong) +
    S('🔴 RISQUES conformité', risques) + S('🔒 Poses à risque', bloquees) +
    '\n\n📈 ACTIVITÉ 7 DERNIERS JOURS : ' + nouv7.length + ' nouveau(x) lead(s), ' + all.filter(function (d) { return e(d) >= 8 && e(d) <= 11 && joursDepuis(d.dateInstallation || d.datePose) <= 7; }).length + ' chantier(s) posé(s).' +
    '\nPipeline : ' + [[1, 3, 'leads'], [4, 4, 'devis en attente'], [5, 6, 'dossiers'], [7, 7, 'poses programmées'], [8, 10, 'chantiers à encaisser']].map(function (x) { return all.filter(function (d) { return e(d) >= x[0] && e(d) <= x[1]; }).length + ' ' + x[2]; }).join(' · ');
  if (lundi) {
    var src = {}; all.forEach(function (d) { var s = d.source || 'Autre'; src[s] = src[s] || {l: 0, p: 0, m: 0}; src[s].l++; if (e(d) >= 8 && e(d) <= 11) { src[s].p++; src[s].m += margePrev(d); } });
    body += '\n\n📊 BILAN PAR SOURCE (depuis le début)\n' + Object.keys(src).sort(function (a, b) { return src[b].l - src[a].l; }).map(function (s) { var x = src[s]; return '- ' + s + ' : ' + x.l + ' dossiers → ' + x.p + ' posés (' + (x.l ? Math.round(100 * x.p / x.l) : 0) + ' %) · marge ' + euros(x.m); }).join('\n');
    var vieux = all.filter(function (d) { return e(d) === 12 && joursDepuis(d.updatedAt) >= 1095; });
    if (vieux.length) body += '\n\n🛡️ RGPD : ' + vieux.length + ' lead(s) sans contact depuis 3 ans → à supprimer après ton accord.';
  }
  body += '\n\nOuvrir le CRM : https://sassi-14.github.io/crm-eco-roy/\n\nVictor, pilotage — équipe Eco Roy';
  MailApp.sendEmail({to: c.EMAIL_ALERTES || Session.getEffectiveUser().getEmail(), subject: (risques.length || bloquees.length ? '🔴 ' : '📊 ') + 'Pilotage Eco Roy — ' + euros(resteCEE + resteMPR) + ' à encaisser' + (lundi ? ' · bilan semaine' : ''), body: body});
}
/** Le 1er du mois : export comptable du mois précédent (feuille Google dans le Drive) + envoi au comptable si renseigné. */
function botExportCompta() {
  var c = getConfig(); var now = new Date(); var deb = new Date(now.getFullYear(), now.getMonth() - 1, 1), fin = new Date(now.getFullYear(), now.getMonth(), 1);
  var dans = function (s) { if (!s) return false; var t = new Date(s).getTime(); return t >= deb.getTime() && t < fin.getTime(); };
  var rows = [['Client', 'Ville', 'Date pose', 'N° facture', 'Date facture', 'Valo CEE', 'CEE encaissé', 'Valo MPR', 'MPR encaissé', 'Reste à charge', 'Matériel', 'Fournisseur', 'Poseur', 'Commission', 'Fournisseur payé le', 'Commission payée le', 'Marge prév.']];
  readAll().forEach(function (r) { var d = r.data; if (d._deleted) return; if (!(dans(d.dateFacture) || dans(d.dateInstallation || d.datePose) || dans(d.fournisseurPayeLe) || dans(d.commissionPayeeLe))) return;
    rows.push([nomComplet(d), d.ville || '', d.dateInstallation || d.datePose || '', d.numFacture || '', d.dateFacture || '', n(d.valoCEE), n(d.encaisseCEE), n(d.valoMPR), n(d.encaisseMPR), n(d.rac), n(d.prixMateriel), d.fournisseur || '', n(d.prixPoseur), n(d.commission), d.fournisseurPayeLe || '', d.commissionPayeeLe || '', Math.round(margePrev(d))]); });
  var nom = 'Export compta Eco Roy — ' + Utilities.formatDate(deb, 'Europe/Paris', 'MM-yyyy');
  var ss = SpreadsheetApp.create(nom); ss.getSheets()[0].getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  notifyAdmin('📒 ' + nom + ' prêt', (rows.length - 1) + ' ligne(s). Fichier : ' + ss.getUrl() + (c.EMAIL_COMPTABLE ? '\n\nÀ transmettre au comptable (' + c.EMAIL_COMPTABLE + ') après vérification.' : ''));
}
/** À lancer une fois : réglages + déclencheurs des nouveaux bots (sans toucher aux autres). */
function installerBotsV2() {
  ensureConfigV2();
  var have = {}; ScriptApp.getProjectTriggers().forEach(function (t) { have[t.getHandlerFunction()] = true; });
  if (!have.botAutomatisations) ScriptApp.newTrigger('botAutomatisations').timeBased().everyHours(1).create();
  if (!have.botPilotage) ScriptApp.newTrigger('botPilotage').timeBased().atHour(7).nearMinute(30).everyDays(1).inTimezone('Europe/Paris').create();
  if (!have.botExportCompta) ScriptApp.newTrigger('botExportCompta').timeBased().onMonthDay(1).atHour(8).inTimezone('Europe/Paris').create();
  Logger.log('✅ Bots v2 installés : automatisations (toutes les heures, inactives tant que AUTO_MAILS=non), pilotage (7h30), export compta (le 1er du mois).');
}
