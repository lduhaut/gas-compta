function onOpen() {
  
  // Création du menu perso
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
    name : "Nouveau Mois",
    functionName : "LDUComptes.initNewMonth"
    }
                ];
  spreadsheet.addMenu("Menu Perso", entries);
  
  
};

function reinit() {
  LDUComptes.supprimerTousLesMois(); 
}
