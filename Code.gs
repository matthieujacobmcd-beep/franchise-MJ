/**
 * ============================================================================
 *  FRANCHISE PY — Saisie KPI Lourdes & Pouzac
 *  Backend Google Apps Script (à coller dans l'éditeur Apps Script du Google
 *  Sheet dédié). Voir GUIDE_MISE_EN_PLACE.md pour les étapes de déploiement.
 * ============================================================================
 *
 *  Ce script fait deux choses :
 *   - doPost(e)  : reçoit une saisie depuis le formulaire HTML et l'écrit
 *                  (ou la met à jour) dans le bon onglet du Google Sheet.
 *   - doGet(e)   : renvoie toutes les données en JSON pour alimenter le
 *                  dashboard en temps réel.
 *
 *  IMPORTANT : si vous ajoutez/renommez un KPI, faites le changement à 3
 *  endroits qui doivent rester alignés :
 *    1. La liste "columns" ci-dessous (les identifiants, pas les libellés)
 *    2. KPI_CONFIG dans kpi_config.js (fichier partagé par la saisie et le dashboard)
 *    3. La ligne d'en-tête du bon onglet dans le Google Sheet (créée
 *       automatiquement au premier envoi si l'onglet n'existe pas encore —
 *       si l'onglet existe déjà, corrigez la ligne d'en-tête à la main)
 * ============================================================================
 */

// ---- 1. Collez ici l'ID de votre Google Sheet -----------------------------
// (dans l'URL du Sheet : https://docs.google.com/spreadsheets/d/CET_ID_LA/edit)
const SHEET_ID = 'COLLER_ICI_ID_DU_GOOGLE_SHEET';

// ---- 1bis. Dossier Google Drive pour l'onglet Marketing --------------------
// Tout fichier déposé dans ce dossier Drive (affiche, flyer, PDF, visuel...)
// apparaît automatiquement dans l'espace "Marketing" du dashboard — rien
// d'autre à faire. Laissez tel quel (COLLER_ICI...) pour désactiver.
// (dans l'URL du dossier : https://drive.google.com/drive/folders/CET_ID_LA)
const MARKETING_FOLDER_ID = '11US2U2Jc3_d75gwFh_xlggrGNWTBAAmn';

// ---- 2. Référentiel des onglets / colonnes ---------------------------------
// L'ordre des colonnes ici = l'ordre des colonnes créées dans le Sheet.
const SHEETS = {
  quotidien: {
    name: 'Quotidien',
    dateField: 'date',
    columns: [
      'date', 'restaurant',
      'ca_net_total', 'ca_net_alim', 'transactions', 'panier_moyen', 'pertes',
      'horodatage'
    ]
  },
  hebdomadaire: {
    name: 'Hebdomadaire',
    dateField: 'semaine',
    columns: [
      'semaine', 'restaurant',
      'oepe', 'r2p', 'oph', 'hello_mcdo', 'reseaux_sociaux',
      'analyse_silliker', 'audit_merieux', 'vphe', 'vphg', 'uvhe', 'tche',
      'horodatage'
    ]
  },
  mensuel: {
    name: 'Mensuel',
    dateField: 'mois',
    columns: [
      'mois', 'restaurant',
      'effectif', 'etp', 'heures_comp', 'heures_sup',
      'turnover_12m', 'turnover_90j', 'absent_imprevu', 'absent_injust',
      'marge', 'pac', 'cout_mo', 'cash_flow', 'resultat_net',
      'remboursement', 'ecart_rendement',
      'horodatage'
    ]
  },
  ponctuel: {
    name: 'Ponctuel',
    dateField: 'date',
    columns: [
      'date', 'restaurant', 'type', 'score', 'commentaire',
      'horodatage'
    ]
  }
};

const RESTAURANTS = ['LOURDES', 'POUZAC'];

// ---- Bannières d'actualité & demandes d'intervention maintenance ----------
const BANNIERES_SHEET = 'Bannieres';
const BANNIERES_COLUMNS = ['id', 'type', 'titre', 'message', 'restaurant', 'date', 'horodatage'];

const INTERVENTIONS_SHEET = 'Interventions';
const INTERVENTIONS_COLUMNS = [
  'id', 'date', 'restaurant', 'demandeur', 'description', 'urgence',
  'techniciens', 'superviseur', 'statut', 'horodatage'
];

// ============================================================================
//  RÉCEPTION D'UNE SAISIE / ACTION
// ============================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.type === 'banniere') return handleBanniere_(data);
    if (data.type === 'intervention') return handleIntervention_(data);

    const freq = data.frequence;
    const cfg = SHEETS[freq];

    if (!cfg) {
      return jsonOut({ ok: false, error: 'Fréquence inconnue : ' + freq });
    }
    if (RESTAURANTS.indexOf(data.restaurant) === -1) {
      return jsonOut({ ok: false, error: 'Restaurant inconnu : ' + data.restaurant });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
      sheet.appendRow(cfg.columns);
      sheet.setFrozenRows(1);
    }

    data.horodatage = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const row = cfg.columns.map(function (c) {
      return data[c] !== undefined && data[c] !== null ? data[c] : '';
    });

    const rowIndex = findExistingRow_(sheet, cfg, data);
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return jsonOut({ ok: true, updated: rowIndex > 0 });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/**
 * Cherche une ligne déjà existante pour la même période + le même
 * restaurant (+ le même type pour les KPI ponctuels), afin de la mettre
 * à jour plutôt que de créer un doublon si on corrige une saisie.
 */
function findExistingRow_(sheet, cfg, data) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return -1;

  const headers = values[0];
  const dateColIdx = headers.indexOf(cfg.dateField);
  const restoColIdx = headers.indexOf('restaurant');
  const typeColIdx = headers.indexOf('type'); // -1 si la colonne n'existe pas

  for (let i = 1; i < values.length; i++) {
    const sameDate = String(values[i][dateColIdx]) === String(data[cfg.dateField]);
    const sameResto = values[i][restoColIdx] === data.restaurant;
    const sameType = typeColIdx === -1 || values[i][typeColIdx] === data.type;
    if (sameDate && sameResto && sameType) return i + 1; // +1 : index Sheet 1-based
  }
  return -1;
}

// ============================================================================
//  BANNIÈRES D'ACTUALITÉ
// ============================================================================
function handleBanniere_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(BANNIERES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(BANNIERES_SHEET);
    sheet.appendRow(BANNIERES_COLUMNS);
    sheet.setFrozenRows(1);
  }

  if (data.action === 'delete') {
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('id');
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === String(data.id)) {
        sheet.deleteRow(i + 1);
        return jsonOut({ ok: true, deleted: true });
      }
    }
    return jsonOut({ ok: false, error: 'Bannière introuvable : ' + data.id });
  }

  // action par défaut : ajout
  const row = {
    id: data.id || String(Date.now()),
    type: data.banner_type || data.type_banniere || 'info',
    titre: data.titre || '',
    message: data.message || '',
    restaurant: data.restaurant || '', // vide = toutes les vues
    date: data.date || '',
    horodatage: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
  sheet.appendRow(BANNIERES_COLUMNS.map(function (c) { return row[c]; }));
  return jsonOut({ ok: true, id: row.id });
}

// ============================================================================
//  DEMANDES D'INTERVENTION MAINTENANCE
// ============================================================================
function handleIntervention_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(INTERVENTIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(INTERVENTIONS_SHEET);
    sheet.appendRow(INTERVENTIONS_COLUMNS);
    sheet.setFrozenRows(1);
  }

  const row = {
    id: data.id || String(Date.now()),
    date: data.date || '',
    restaurant: data.restaurant || '',
    demandeur: data.demandeur || '',
    description: data.description || '',
    urgence: data.urgence || '',
    techniciens: data.techniciens || '',
    superviseur: data.superviseur || '',
    statut: data.statut || 'Ouverte',
    horodatage: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };
  sheet.appendRow(INTERVENTIONS_COLUMNS.map(function (c) { return row[c]; }));
  return jsonOut({ ok: true, id: row.id });
}

// ============================================================================
//  LECTURE POUR LE DASHBOARD
// ============================================================================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const result = {};

    Object.keys(SHEETS).forEach(function (freq) {
      const cfg = SHEETS[freq];
      const sheet = ss.getSheetByName(cfg.name);
      if (!sheet) { result[freq] = []; return; }

      const values = sheet.getDataRange().getValues();
      if (values.length < 2) { result[freq] = []; return; }

      const headers = values[0];
      result[freq] = values.slice(1)
        .filter(function (r) { return r[0] !== ''; }) // ignore lignes vides
        .map(function (r) {
          const obj = {};
          headers.forEach(function (h, i) { obj[h] = r[i]; });
          return obj;
        });
    });

    result.bannieres = readSheetAsObjects_(ss, BANNIERES_SHEET);
    result.interventions = readSheetAsObjects_(ss, INTERVENTIONS_SHEET);
    result.marketing = getMarketingFiles_();
    // Diagnostic temporaire : à retirer une fois le problème résolu.
    result.marketing_debug = {
      folder_id_configure: MARKETING_FOLDER_ID,
      nb_fichiers_trouves: result.marketing.length,
      derniere_erreur: MARKETING_LAST_ERROR
    };

    return jsonOut({ ok: true, data: result });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ============================================================================
//  MARKETING — liste en direct des fichiers d'un dossier Google Drive
// ============================================================================
function getMarketingFiles_() {
  if (!MARKETING_FOLDER_ID || MARKETING_FOLDER_ID.indexOf('COLLER_ICI') === 0) return [];
  try {
    const folder = DriveApp.getFolderById(MARKETING_FOLDER_ID);
    const files = folder.getFiles();
    const out = [];
    while (files.hasNext()) {
      const f = files.next();
      // Rend le fichier visible à quiconque a le lien, pour que l'aperçu
      // fonctionne pour toute personne qui ouvre le dashboard (pas seulement
      // vous) — fait automatiquement, rien à régler sur chaque fichier.
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {}
      const updated = f.getLastUpdated();
      out.push({
        id: f.getId(),
        name: f.getName(),
        mimeType: f.getMimeType(),
        url: f.getUrl(),
        updated: updated ? Utilities.formatDate(updated, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
      });
    }
    out.sort(function (a, b) { return String(b.updated).localeCompare(String(a.updated)); });
    return out;
  } catch (err) {
    // NE PAS avaler l'erreur silencieusement : on la remonte dans le JSON
    // (champ marketing_error) pour pouvoir diagnostiquer depuis le
    // navigateur en ouvrant simplement l'URL du Web App. À retirer une
    // fois que le dossier Marketing fonctionne normalement.
    MARKETING_LAST_ERROR = String(err);
    return [];
  }
}
var MARKETING_LAST_ERROR = '';

function readSheetAsObjects_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
