function onOpen() {
  // Création du menu perso
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "Nouveau Mois",
      functionName : "LDUComptes.initNewMonth"
    },
    {
      name : "Export pour BNC Express",
      functionName : "exportBNCExpress"
    }
  ];
  spreadsheet.addMenu("Menu Perso", entries);
};

function reinit() {
  LDUComptes.supprimerTousLesMois(); 
}

/**
 * Export pour BNC Express :
 * Exporte les écritures non présentes dans BNC Express (colonne I) de la feuille active, dans une feuille 'Export BNCExpress' 
 * N'exporte que les écritures de la catégorie 'Autres', les autres étant gérées normalement par import d'ICT.
 */
function exportBNCExpress() {
  var firstDataLineNum = 5;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = SpreadsheetApp.getActiveSheet();
  
  var exportSheet = ss.getSheetByName("Export BNCExpress");
  if (!exportSheet) {
    exportSheet = ss.insertSheet("Export BNCExpress");
  } else {
    exportSheet.clear(); 
  }
  
  var isExportOn = false;
  var data = currentSheet.getRange(5, 1, 200, 9).getValues();
  var exportedData = [];
  
  var lastDate = '';
  var libelleMonth = currentSheet.getName();
  var numMonth = LDUComptes.NOM_FEUILLES_MENSUELLES.indexOf(libelleMonth);
  
  // Index des lignes exportées
  var indexes = [];
  
  for (var i = 0; i<data.length; i++) {
    var row = data[i];
    
    var libelle = row[1];
    if (libelle == 'Total dép prévu') {
      // On ne traite pas les lignes après
      Logger.log('BREAK on Total dép prévu');
      break; 
    }
    
    var resteReel = row[5];
    var resteBanque = row[6];
    if (resteReel == '' && resteBanque == '') {
      // On ne traite pas les lignes après
      Logger.log('BREAK on Restes');
      break; 
    }
    
    if (isExportOn) {
      if (libelle == 'Tiers payant, ALD & co') {
        isExportOn = false;
        // On stoppe le traitement
        Logger.log('Tiers payant, ALD & co');
        break;
      } else {
        var checked = row[7];
        var inBNC = row[8];
        if (checked == 'ok' && !inBNC) {
          // On traite les lignes validées en banque, mais pas encore dans BNC Express
          Logger.log('i : ' + i + ' : ' + row[1]); 
          
          var date = row[0];
          if (date) {
            // Midi, pour ne pas se prendre la tête avec le décalage horaire
            lastDate = Utilities.formatDate(new Date(2017, numMonth, date, 12), "GMT", "dd/MM/yyyy");
          }
          date = lastDate; 
          
          var credit = row[3];
          credit = (typeof credit == 'number') ? credit.toFixed(2).toString().replace('.',',') : '';
          var debit = row[4];
          debit = (typeof debit == 'number') ? debit.toFixed(2).toString().replace('.',',') : '';
          
          var rowExp = [];
          exportedData.push(rowExp);
          rowExp.push(date);
          rowExp.push(libelle);
          
          rowExp.push(credit);
          rowExp.push(debit);
          rowExp.push(credit ? credit : debit);
          
          var numCompte = row[2];
          if (!numCompte) {
            numCompte = '490100';
          }
          rowExp.push(numCompte);
          indexes.push(i + firstDataLineNum);
        }
      }
    } else {
      if (libelle == 'Autres' || libelle == 'Rétrocessions') {
        isExportOn = true;
      }
    }
  }
  
  exportSheet.getRange(1, 1, exportedData.length, exportedData[0].length).setValues(exportedData);
  
  // update champ BNC Express dans le sheet
  var now = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy");
  for (var t in indexes) {
    var numRow = indexes[t];
    currentSheet.getRange(numRow, 9).setValue(now);
  }
}
