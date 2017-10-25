var NOM_FEUILLES_MENSUELLES = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Juin', 'Juil', 'Aou', 'Sep', 'Oct', 'Nov', 'Déc'];

var ALPHABET = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];

// Index et lettres des colonnes dans les feuilles mensuelles
var idx = 0;
var LETTRE_COL_DATE = ALPHABET[idx];
var IDX_COL_DATE = ++idx;
var LETTRE_COL_INTITULE = ALPHABET[idx];
var IDX_COL_INTITULE = ++idx;
var LETTRE_COL_PRECISION = ALPHABET[idx];
var IDX_COL_PRECISION = ++idx;
var LETTRE_COL_CREDIT = ALPHABET[idx];
var IDX_COL_CREDIT = ++idx;
var LETTRE_COL_DEBIT = ALPHABET[idx];
var IDX_COL_DEBIT = ++idx;
var LETTRE_COL_RESTE_REEL = ALPHABET[idx];
var IDX_COL_RESTE_REEL = ++idx;
var LETTRE_COL_RESTE_BANQUE = ALPHABET[idx];
var IDX_COL_RESTE_BANQUE = ++idx;

var NOM_FEUILLE_OPERATIONS = 'Opérations';
// Index dans la feuille opérations
idx = 1;
var FEUILLE_OPERATIONS_COL_JOUR = idx++;
var FEUILLE_OPERATIONS_COL_MOIS_DEBUT = idx++;
var FEUILLE_OPERATIONS_COL_MOIS_FIN = idx++;
var FEUILLE_OPERATIONS_COL_FREQUENCE = idx++;
var FEUILLE_OPERATIONS_COL_LIBELLE = idx++;
var FEUILLE_OPERATIONS_COL_CREDIT = idx++;
var FEUILLE_OPERATIONS_COL_DEBIT = idx++;
var FEUILLE_OPERATIONS_COL_REMARQUE = idx++;
var FEUILLE_OPERATIONS_COL_COMPTE_BNC_EXPRESS = idx++; // Pour compta Ju only


function initNewMonth() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var feuilleAGenerer;
  for (var i = 0; i < NOM_FEUILLES_MENSUELLES.length; i++) {
    var feuilleAGenerer =  NOM_FEUILLES_MENSUELLES[i];
    
    var sheet = spreadsheet.getSheetByName(feuilleAGenerer);
    if (!sheet) {
      break;
    } else {
      feuilleAGenerer = ''; 
    }
  }
  
  if (feuilleAGenerer) {
    var templateSheet = spreadsheet.getSheetByName('template');
    var monthSheet = templateSheet.copyTo(spreadsheet);
    monthSheet.setName(feuilleAGenerer);
    spreadsheet.setActiveSheet(monthSheet);
    spreadsheet.moveActiveSheet(1);
    
    var idxLastRow = monthSheet.getLastRow();
    
    var idxMoisAGenerer = NOM_FEUILLES_MENSUELLES.indexOf(feuilleAGenerer);
    
    // Init restes / mois précédent
    if (idxMoisAGenerer == 0) {
      
    } else {
      // On initialise le montant avec la dernière valeur du mois précédent
      var lastMonthSheetName = NOM_FEUILLES_MENSUELLES[idxMoisAGenerer-1];
      var lastMonthSheet = spreadsheet.getSheetByName(lastMonthSheetName);
      
      var lastRowLastMonth = lastMonthSheet.getLastRow();
      for (var i = 6; i <= lastRowLastMonth; i++) {
        var cellResteReel = lastMonthSheet.getRange(i, IDX_COL_RESTE_REEL);
        if (cellResteReel.getValue() == '' && cellResteReel.getFormula() == '') {
          Logger.log(i + ' : ' + IDX_COL_RESTE_REEL + ' est vide');
          lastRowLastMonth = i-1;
          break;
        }
      }

      monthSheet.getRange(3, 2).setFormula('='+lastMonthSheetName+'!'+LETTRE_COL_RESTE_REEL+lastRowLastMonth);
      monthSheet.getRange(3, 3).setFormula('='+lastMonthSheetName+'!'+LETTRE_COL_RESTE_BANQUE+lastRowLastMonth);
    }
    
    // TODO Init opérations non pointées du mois n-1
    
    // Init opérations
    var sheetOperations = spreadsheet.getSheetByName(NOM_FEUILLE_OPERATIONS);
    var idxLastOperation = sheetOperations.getLastRow();
    var idxInsertionInMonthSheet = idxLastRow + 2;
    for (var i = 2; i <= idxLastOperation; i++) {
      var operation = sheetOperations.getRange(i, 1, 1, 9).getValues()[0];
      if (isOperationInMonth(operation, idxMoisAGenerer+1)) {
          monthSheet.getRange(idxInsertionInMonthSheet, IDX_COL_DATE).setValue(operation[FEUILLE_OPERATIONS_COL_JOUR - 1]);
          monthSheet.getRange(idxInsertionInMonthSheet, IDX_COL_INTITULE).setValue(operation[FEUILLE_OPERATIONS_COL_LIBELLE - 1]);
          monthSheet.getRange(idxInsertionInMonthSheet, IDX_COL_CREDIT).setValue(operation[FEUILLE_OPERATIONS_COL_CREDIT - 1]);
          monthSheet.getRange(idxInsertionInMonthSheet, IDX_COL_DEBIT).setValue(operation[FEUILLE_OPERATIONS_COL_DEBIT - 1]);
          monthSheet.getRange(idxInsertionInMonthSheet, IDX_COL_PRECISION).setValue(operation[FEUILLE_OPERATIONS_COL_COMPTE_BNC_EXPRESS - 1]);
          monthSheet.getRange(idxInsertionInMonthSheet, IDX_COL_RESTE_REEL).setValue(operation[FEUILLE_OPERATIONS_COL_REMARQUE - 1]);
      
          idxInsertionInMonthSheet++;
      }
    }
  }
}

function isOperationInMonth(operation, idxMoisAGenerer) {
  var moisDebut = operation[FEUILLE_OPERATIONS_COL_MOIS_DEBUT - 1];
  var moisFin = operation[FEUILLE_OPERATIONS_COL_MOIS_FIN - 1];
  var frequence = operation[FEUILLE_OPERATIONS_COL_FREQUENCE - 1];
  
  if (idxMoisAGenerer >= moisDebut && idxMoisAGenerer <= moisFin) {
    var isOkFrequence = (idxMoisAGenerer - moisDebut)%frequence == 0
    return isOkFrequence;    
  }
}

// Pour initialiser la feuille d'une nouvelle année
function supprimerTousLesMois() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  for (var i = 0; i < NOM_FEUILLES_MENSUELLES.length; i++) {
    var feuilleASuppr =  NOM_FEUILLES_MENSUELLES[i];
    
    var sheet = spreadsheet.getSheetByName(feuilleASuppr);
    if (!sheet) {
      break;
    } else {
      spreadsheet.deleteSheet(sheet);
    }
  }
}
