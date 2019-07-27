// Id fichier listing facture 2016-2017
var ID_FEUILLE_LISTING_FACTURE = '15b7mrxzgZlBQtgfegTFLOGX58gHsI7G6KTtL-H5dDhU';

function onOpen() {
  // Création du menu perso
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "01 - Réinitialiser le fichier",
      functionName : "reinitialiser"
    }
    ,
    {
      name : "02 - Ajouter colonnes",
      functionName : "enrichirFichier"
    }
    ,
    {
      name : "03 - Supprimer factures annulées",
      functionName : "supprFacturesAnnulees"
    }
    ,
    {
      name : "04 - Clean feuille Données",
      functionName : "cleanFeuilleDonnees"
    }
    ,
    {
      name : "05 - Export Maison Retraite",
      functionName : "exportMaisonRetraite"
    }
  ];
  spreadsheet.addMenu("Menu Perso", entries);
}

function reinitialiser() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.getSheetByName('Données').clear();
  
  // Suppr onglet Factures
  var sheetFactures = spreadsheet.getSheetByName('Factures');
  spreadsheet.deleteSheet(sheetFactures);
  
  // Suppr onglet Données v1 si existant
  var sheetDonneesV1 = spreadsheet.getSheetByName('Données v1');
  if (sheetDonneesV1) spreadsheet.deleteSheet(sheetDonneesV1);
  
  // Suppr onglet Export MR si existant
  var sheetExportMR = spreadsheet.getSheetByName('ExportMR');
  if (sheetExportMR) spreadsheet.deleteSheet(sheetExportMR);
}

function test() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Données v1');
  
  var lastRow = sheet.getLastRow();
  
  var formulas = sheet.getRange(846,13,1,5).getFormulasR1C1();
  for (var i = 0; i < formulas.length; i++) {
   Logger.log(formulas[i]); 
  }
}

/**
 * A partir de Données v1, filtrer sur Libellé = Part patient - Virement.
 * Filtrer sur les montants positifs
 *
 * Trier sur la date d’écriture (pour que ça soit trié dans BNC Express)
 * Dans la colonne moyen paiement (M), mettre “Virement MR”.
 * Dans la colonne Libellé (G), mettre ="Acte du "&JOUR(N1)&"/"&SI(MOIS(N1)>9;MOIS(N1);"0"&MOIS(N1)) et tirer jusqu’en bas
 *
 * Dans Données v1, supprimer les lignes exportées
 */
function exportMaisonRetraite() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetV1 = spreadsheet.getSheetByName('Données v1');
  
  var lastRow = sheetV1.getLastRow();
  var allData = sheetV1.getRange(2, 1, lastRow - 1, 15).getDisplayValues();
  
  var dataToExport = allData.filter(function(a) { 
    
    // Libellé = Part patient - Virement
    if (a[6] == 'Part patient - Virement') { 
      // montants positifs (On utilise parseInt pour convertir 7.5 en un int (7). Sinon, '7.5' !> 0
      return (parseInt(a[11])) > 0;
    }
    return false
  });
  
  Logger.log("Maison retraite : " + dataToExport.length + " lignes")
  
  // Trier sur la date d’écriture (pour que ça soit trié dans BNC Express)
  dataToExport.sort(function(a, b) {
    return b[1] - a[1];
  });
  
  var sheetExportMR = spreadsheet.insertSheet("ExportMR")
  var export = dataToExport.map(function(a, idx) {
    var date = a[13].split('/')
    var jour = date[0]
    var mois = date[1]
    if (mois.length == 1) mois = '0'+mois
    
    // Dans la colonne Libellé (G), mettre ="Acte du "&JOUR(N1)&"/"&SI(MOIS(N1)>9;MOIS(N1);"0"&MOIS(N1)) et tirer jusqu’en bas
    a[6] = 'Acte du ' + jour + '/' + mois
    
    // Dans la colonne moyen paiement (M), mettre “Virement MR”.
    a[12] = "Virement MR"
    
    return a;
  })
  sheetExportMR.getRange(1,1, export.length, 15).setValues(export)
  
  // Supprimer les lignes dans Donnees V1
  allData = sheetV1.getRange(2, 1, lastRow - 1, 15).getDisplayValues();
  var dataToSuppr = allData.map(function(a, idx) {
    var obj = {};
    obj.idx = idx + 2;
    obj.data = a;
    return obj;
  }).filter(function(a) { 
    // Libellé = Part patient - Virement (positif et négatif)
    return a.data[6] == 'Part patient - Virement'
  });
  Logger.log("Maison retraite : " + dataToSuppr.length + " lignes à supprimer")
  dataToSuppr.sort(function(a, b) {
    return b.idx - a.idx;
  });
  
  for (i in dataToSuppr) {
    sheetV1.deleteRow(dataToSuppr[i].idx);
  }
}

function enrichirFichier() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Créer un nouvel onglet “Factures”. Y Coller le contenu du fichier Sheet “Listing Factures”
  var ssListingFacture = SpreadsheetApp.openById(ID_FEUILLE_LISTING_FACTURE);
  var sheetListing = ssListingFacture.getSheetByName('Listing');
  var sheetFactures = sheetListing.copyTo(spreadsheet);
  sheetFactures.setName('Factures');
  
  var sheet = spreadsheet.getSheetByName('Données');
  
  var lastRow = sheet.getLastRow();
  
  sheet.getRange(1, 9, 1, 7).setValues([['Num Facture', 'Nom Patient', 'NumFact + Patient', 'Montant', 'Moyen paiement', 'Date facture', 'Payé le jour même']])
  
  // En colonne I, calculer le numéro de facture : =0+STXT(D2;3;6)
  sheet.getRange(2, 9, lastRow - 1).setFormulaR1C1("=0+RIGHT(R[0]C[-5];LEN(R[0]C[-5])-2)");
  
  // Sur l’onglet principal, en colonne J, on va récupérer le nom/prénom du patient. 
  sheet.getRange(2, 10, lastRow - 1).setFormulaR1C1("=VLOOKUP(R[0]C[-1];Factures!R2C[-9]:C[-4];6;FALSE)");
  
  // Dans BNC Express, on souhaite que le numéro de facture et le nom/prénom soient stockés au même endroit
  sheet.getRange(2, 11, lastRow - 1).setFormulaR1C1('=R[0]C[-2]&" - "&R[0]C[-1]');
  
  // Exceptionnellement, on peut avoir des écritures qui sont passées à l’envers, pour en annuler une autre. On va donc cumuler débit et crédit, tout en reformatant correctement
  sheet.getRange(2, 12, lastRow - 1).setFormulaR1C1('=0+SUBSTITUTE(R[0]C[-7];".";",";1)-SUBSTITUTE(R[0]C[-6];".";",";1)');
  
  // Récupération du type de paiement
  sheet.getRange(2, 13, lastRow - 1).setFormulaR1C1('=VLOOKUP(R[0]C[-10];MoyensPaiement!R1C[-12]:R12C[-11];2;FALSE)');
  
  // Récupération date de facture
  sheet.getRange(2, 14, lastRow - 1).setFormulaR1C1('=VLOOKUP(R[0]C[-5];Factures!R2C[-13]:C[-8];3;FALSE)');
  
  // Vérif si dateFacture = datePaiement
  sheet.getRange(2, 15, lastRow - 1).setFormulaR1C1('=RIGHT(R[0]C[-1];4)&"-"&LEFT(RIGHT(R[0]C[-1];7);2)&"-"&LEFT(R[0]C[-1];2)=LEFT(R[0]C[-13];10)');
  
  // Figer la première ligne
  sheet.setFrozenRows(1);
}

function cleanFeuilleDonnees() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Données');
  
  var lastRow = sheet.getLastRow();
  var allData = sheet.getRange(1, 1, lastRow - 1, 15).getDisplayValues();
  
  var dataToKeep = allData.filter(function(a) { 
    var noCompte = a[2]
    return noCompte == 'No de compte' || (noCompte != 999902 && noCompte > 998800);
  });
  
  sheet.clear()
  
  sheet.getRange(1, 1, dataToKeep.length, 15).setValues(dataToKeep)
}

function supprFacturesAnnulees() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Données');
  
  var lastRow = sheet.getLastRow();
  
  var data = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  
  var facturesImpayees = data.map(function(a, idx) {
    var obj = {};
    obj.idx = idx;
    obj.data = a;
    return obj;
  }).filter(function(a) { 
    return a.data[2] == 998807;
  });
  
  var indexsASuppr = [];
  var lignesParNum = {}; // lignesParNum[numFacture] = {idx: index
  for (i in facturesImpayees) {
    var obj = facturesImpayees[i];
    if (!lignesParNum[obj.data[8]]) {
      lignesParNum[obj.data[8]] = [];
    }
    
    lignesParNum[obj.data[8]].push(obj);
  }
  
  for (i in lignesParNum) {
    var obj = lignesParNum[i];
    if (obj.length > 1) {
      Logger.log("La facture " + i + " apparait plusieurs fois");
      
      var factAmount = obj.reduce(function(a, b) {
        // Pb les nb à virgules avec un point sont considérés comme des dates ...
        return +a.data[11] + +b.data[11];
      });
      Logger.log('Montant : ' + factAmount);
      
      if (!factAmount) {
        Logger.log("La facture " + i + " DOIT ETRE SUPPRIMEE");
        indexsASuppr = indexsASuppr.concat(
          obj.map(function(a) {
            return a.idx + 2; 
          })
        );
      }
    }
  }
  
  indexsASuppr.sort(function(a, b) {
    return b - a;
  });
  
  for (i in indexsASuppr) {
    sheet.deleteRow(indexsASuppr[i]);
  }
  
  creerCopieOngletV1();
}

function creerCopieOngletV1() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Données');
  
  var copy = sheet.copyTo(spreadsheet);
  copy.setName('Données v1');
}
