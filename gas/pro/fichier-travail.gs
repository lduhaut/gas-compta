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

function enrichirFichier() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Données');
  
  var lastRow = sheet.getLastRow();
  
  // En colonne I, calculer le numéro de facture : =0+STXT(D2;3;6)
  sheet.getRange(1, 9).setValue('Num Facture');
  sheet.getRange(2, 9, lastRow - 1).setFormulaR1C1("=0+RIGHT(R[0]C[-5];5)");
  
  // Créer un nouvel onglet “Factures”. Y Coller le contenu du fichier Sheet “Listing Factures”
  var ssListingFacture = SpreadsheetApp.openById(ID_FEUILLE_LISTING_FACTURE);
  var sheetListing = ssListingFacture.getSheetByName('Listing');
  var sheetFactures = sheetListing.copyTo(spreadsheet);
  sheetFactures.setName('Factures');
  
  // Sur l’onglet principal, en colonne J, on va récupérer le nom/prénom du patient. 
  sheet.getRange(1, 10).setValue('Nom Patient');
  sheet.getRange(2, 10, lastRow - 1).setFormulaR1C1("=VLOOKUP(R[0]C[-1];Factures!R2C[-9]:C[-4];6;FALSE)");
  
  // Dans BNC Express, on souhaite que le numéro de facture et le nom/prénom soient stockés au même endroit
  sheet.getRange(1, 11).setValue('NumFact + Patient');
  sheet.getRange(2, 11, lastRow - 1).setFormulaR1C1('=R[0]C[-2]&" - "&R[0]C[-1]');
  
  // Exceptionnellement, on peut avoir des écritures qui sont passées à l’envers, pour en annuler une autre. On va donc cumuler débit et crédit, tout en reformatant correctement
  sheet.getRange(1, 12).setValue('Montant');
  sheet.getRange(2, 12, lastRow - 1).setFormulaR1C1('=0+SUBSTITUTE(R[0]C[-7];".";",";1)-SUBSTITUTE(R[0]C[-6];".";",";1)');
  
  // Récupération du type de paiement
  sheet.getRange(1, 13).setValue('Moyen paiement');
  sheet.getRange(2, 13, lastRow - 1).setFormulaR1C1('=VLOOKUP(R[0]C[-10];MoyensPaiement!R1C[-12]:R12C[-11];2;FALSE)');
  
  // Récupération date de facture
  sheet.getRange(1, 14).setValue('Date facture');
  sheet.getRange(2, 14, lastRow - 1).setFormulaR1C1('=VLOOKUP(R[0]C[-5];Factures!R2C[-13]:C[-8];3;FALSE)');
  
  // Vérif si dateFacture = datePaiement
  sheet.getRange(1, 15).setValue('Payé le jour même');
  sheet.getRange(2, 15, lastRow - 1).setFormulaR1C1('=RIGHT(R[0]C[-1];4)&"-"&LEFT(RIGHT(R[0]C[-1];7);2)&"-"&LEFT(R[0]C[-1];2)=LEFT(R[0]C[-13];10)');
  
  
  
  // Figer la première ligne
  sheet.setFrozenRows(1);
}
