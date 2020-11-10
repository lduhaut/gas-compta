var ANNEE = 2019;
const LIB_COMMISSION_CB = "Comission Paiements CB - Crédit Mutuel"

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
    },
    {
      name : "Import CCM",
      functionName : "importReleveBanque"
    }
  ];
  spreadsheet.addMenu("Menu Perso", entries);
};

function reinit() {
  LDUComptes.supprimerTousLesMois(); 
}

function importReleveBanque() {
  const firstDataLineNum = 3
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName("ImportBanque");
  
  let dataRange = importSheet.getRange(firstDataLineNum, 1, importSheet.getLastRow() + 1 - firstDataLineNum, 4)
  let dataFormulas = dataRange.getFormulas()
  let data = dataRange.getValues().map(function(line, index) {
    let res = {}
    res.date = getNumDayOfDate_(line[0])
    res.libelle = line[1]
    res.debit = readAmount_(line[2])
    if (line[3] == "#ERROR!") {
      // Quand on colle "+ 25,00 EUR" ça devient une formule ...
      res.credit = readAmount_(dataFormulas[index][3])
    } else {
      res.credit = readAmount_(line[3])
    }
    
    addData_(res)
    
    return res
  })
  
  // Regroupement des lignes ComCB par date
  for (let i in data) {
    let ope = data[i]
    if (ope.libelle == LIB_COMMISSION_CB) {
      let otherCommCBSameDay = data.filter((op, j) => op.libelle == LIB_COMMISSION_CB && j < i && op.debit > 0 && op.date == ope.date)
      if (otherCommCBSameDay.length > 0) {
        let cumul = parseFloat(otherCommCBSameDay[0].debit) + parseFloat(ope.debit)
        /* Logger.log('\n\nCOM CB A TRAITER : ')
        Logger.log(JSON.stringify(otherCommCBSameDay[0]))
        Logger.log(JSON.stringify(ope))
        Logger.log('+ ' + otherCommCBSameDay[0].debit + ' + ' + ope.debit + ' = ' + cumul) */
        otherCommCBSameDay[0].debit = cumul
        ope.debit = 0
      }
    }
  }
  data = data.filter(op => op.libelle != LIB_COMMISSION_CB || op.debit > 0)
  
  /*for (let i in data) {
    Logger.log('data['+i+'] = ' + JSON.stringify(data[i]))
  }*/
  
  data.sort((a, b) => a.date - b.date)
  data.sort((a, b) => ('' + a.categorie).localeCompare(b.categorie))
  
  let dataFinales = data.map(obj => {
                             let d = []
                             d[0] = obj.categorie || 'Manuel'
                             d[1] = obj.date
                             d[2] = obj.libelle
                             d[3] = obj.compte
                             d[4] = obj.credit
                             d[5] = obj.debit
                             return d
                             })
  
  for (let i in dataFinales) {
    Logger.log('data['+i+'] = ' + JSON.stringify(dataFinales[i]))
  }
  
  importSheet.getRange(firstDataLineNum, 6, dataFinales.length, dataFinales[0].length).setValues(dataFinales)
}

function getNumDayOfDate_(dte) {
  Logger.log('getNumDayOfDate_ ' + dte)
  return dte.getDate()
}

function readAmount_(libelle) {
  const res = libelle.replace(',', '.').replace(/[^\d.]/g,'')
  // Logger.log('readAmount : ' + libelle + ' -> ' + res)
  return res;
}

function addData_(operation) {
  const CATEGORY_AUTRES = "03_Autres"
  const CATEGORY_PAIEMENTS_CB = "01_Paiements CB"
  const CATEGORY_NOEMIE = "04_Tiers payant, ALD & Maison retraite & co"
  
  if (operation.libelle.indexOf("COMCB") == 0) {
    operation.categorie = CATEGORY_AUTRES
    operation.compte = 627000
    operation.libelle = LIB_COMMISSION_CB
  } else if (operation.libelle.indexOf("REMCB") == 0) {
    operation.categorie = CATEGORY_PAIEMENTS_CB
    operation.libelle = "Paiements CB"
  } else if (operation.libelle.indexOf("VIR ") > 0 && operation.libelle.indexOf("Détail \n      Cliquer pour déplier" == 0)) {
    operation.categorie = CATEGORY_NOEMIE
    
    let ligneLibelle = operation.libelle.split('\n').find(a => a.indexOf("VIR ") >= 0)
    
    operation.libelle = ligneLibelle.substring(ligneLibelle.indexOf("VIR ") + 4)
  } else if (operation.libelle.indexOf("Détail \n      Cliquer pour déplier") == 0) {
    operation.libelle = operation.libelle.substring(operation.libelle.indexOf("opération ") + 10)
  }
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
      if (libelle.indexOf('Tiers payant, ALD') == 0) {
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
            lastDate = Utilities.formatDate(new Date(ANNEE, numMonth, date, 12), "GMT", "dd/MM/yyyy");
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
