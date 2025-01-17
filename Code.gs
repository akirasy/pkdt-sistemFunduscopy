/**
 * Load keyword arguments needed for this project.
 */
function loadProjectKwargs() {
  let mainSpreadsheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/14Kd4o-ziuHQ7ef80_kBIp3o8DHMw7cKuey7kPzasCQs/');
  let output = new Object();
  mainSpreadsheet.getSheetByName('kwargs')
    .getDataRange()
    .getValues()
    .map(item => { output[item[0]] = item[1] });
  return output
}

/**
 * Set formula to column till the end of rows.
 * @param {SpreadsheetApp.Sheet} sheet Sheet object to perform action
 * @param {String} columnAlphabet Column alphabet
 * @param {Number} startRow Start row number
 * @param {Number} formula Formula to set in cell
 */
function setFormulaToWholeColumn(sheet, columnAlphabet, startRow, formula) {
  let templateA1Notation = columnAlphabet + startRow.toString();
  let template = sheet.getRange(templateA1Notation);
  template.setFormula(formula);
  template.copyTo(sheet.getRange(templateA1Notation + ':' + columnAlphabet), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
}

/**
 * Add 2-dimension values to sheet.
 * @param {SpreadsheetApp.Sheet} sheet Sheet object to perform action
 * @param {Array^2} values 2-dimension values
 * @param {Number} startRow Start row number
 * @param {Number} startColumn Start column number
 */
function addValuesToSheet(sheet, values, startRow, startColumn) {
  if (values.length != 0) {
    sheet
      .getRange(startRow, startColumn, values.length, values[0].length)
      .setValues(values);
  };
}

/**
 * Get completed fundus cases from Fundus Operator and copy to respective KK.
 * @param {Object} projectKwargs Common project variables from loadProjectKwargs()
 * @param {Spreadsheet} operatorFundusSpreadsheet Spreadsheet object of Operator Fundus
 * @param {String} operation Use either `selesai` or `defaulter`
 */
function copyOperatorCaseToKk(projectKwargs, operatorFundusSpreadsheet, operation) {
  let intent = {
    selesai   : { operatorSheetName : 'selesaiFundus', bukuDaftarSheetName: 'toReview' },
    defaulter : { operatorSheetName : 'defaulter', bukuDaftarSheetName: 'defaulter' }
  };

  // Get data from operator
  let operatorFundusSheet = operatorFundusSpreadsheet.getSheetByName(intent[operation].operatorSheetName);
  let dataRange = operatorFundusSheet.getRange(2, 1, operatorFundusSheet.getLastRow(), operatorFundusSheet.getMaxColumns());
  let dataRangeValues = dataRange.getValues();
  if (operation == 'selesai') { dataRange.clearContent() };

  // Sort cases into clinics
  let sortedCasesByClinic = {
    kk_temerloh        : dataRangeValues.filter(item => { return item[0] == 'KKT'  }), 
    kk_tanjung_lalang  : dataRangeValues.filter(item => { return item[0] == 'KKTL' }), 
    kk_bandar_mentakab : dataRangeValues.filter(item => { return item[0] == 'KKBM' }), 
    kk_lanchang        : dataRangeValues.filter(item => { return item[0] == 'KKL'  }),
    kk_sanggang        : dataRangeValues.filter(item => { return item[0] == 'KKS'  }), 
    kk_kerdau          : dataRangeValues.filter(item => { return item[0] == 'KKK'  }), 
    kk_kuala_tekal     : dataRangeValues.filter(item => { return item[0] == 'KKKT' }), 
    kk_kuala_krau      : dataRangeValues.filter(item => { return item[0] == 'KKKK' })
  };
  
  // Loop through clinics and put data into corresponding clinics
  Object.entries(sortedCasesByClinic).forEach(clinic => {
    let clinicKey = clinic[0];
    let clinicValue = clinic[1];
    Logger.log('Processing clinic: ' + clinicKey + '. Found ' + clinicValue.length + ' entry.');

    if (clinicValue.length != 0) {
      let bukuDaftarSheet = SpreadsheetApp.openByUrl(projectKwargs[clinicKey]).getSheetByName(intent[operation].bukuDaftarSheetName);
      addValuesToSheet(bukuDaftarSheet, clinicValue, bukuDaftarSheet.getLastRow()+1, 1);
      setFormulaToWholeColumn(bukuDaftarSheet, 'P', 3, '=IFERROR(VLOOKUP(F3,rawNDR!$A:$E,5,0),"")');
    };
  });
}

/**
 * Move reviewed cases to doneReview sheet.
 * @param {Spreadsheet} bukuDaftarSpreadsheet Spreadsheet object of respective clinic Buku Daftar
 */
function caseReviewCleanup(bukuDaftarSpreadsheet) {
  let toReviewSheet = bukuDaftarSpreadsheet.getSheetByName('toReview');
  let doneReviewSheet = bukuDaftarSpreadsheet.getSheetByName('doneReview');

  // Select doneReview case
  let doneValues = new Array();
  let notDoneValues = new Array();

  let dataRange;
  let toReviewSheetLastRow = toReviewSheet.getLastRow();
  if (toReviewSheetLastRow > 3) {
    dataRange = toReviewSheet.getRange(3, 1, toReviewSheetLastRow-2, toReviewSheet.getMaxColumns());
  } else {
    dataRange = toReviewSheet.getRange('3:3');
  }
  let dataRangeValues = dataRange.getValues();
  dataRangeValues.forEach(item => {
    if (item[29] == 'Selesai') {
      doneValues.push(item);
    } else if (item[29] != 'Selesai' && item[0] != '') {
      notDoneValues.push(item);
    };
  });
  
  // Copy values to doneReviewSheet
  addValuesToSheet(doneReviewSheet, doneValues, (doneReviewSheet.getLastRow()+1), 1);

  // Restructure values in toReviewSheet
  toReviewSheet.getRange('3:3').clearContent();
  if (toReviewSheet.getLastRow() > 3) {
    toReviewSheet.deleteRows(4, toReviewSheet.getMaxRows()-3);
  };
  addValuesToSheet(toReviewSheet, notDoneValues, 3, 1);
  setFormulaToWholeColumn(toReviewSheet, 'P', 3, '=IFERROR(VLOOKUP(F3,rawNDR!$A:$E,5,0),"")');
}

/**
 * Run trigger operator cleanup.
 */
function triggerOperatorToClinic() {
  Logger.log('Running trigger...');
  let projectKwargs = loadProjectKwargs();

  let operatorFundusTemerloh = SpreadsheetApp.openByUrl(projectKwargs['operator_fundus_temerloh']);
  // Do note that it is important to run for sheet defaulter first only then proceed to sheet selesai
  copyOperatorCaseToKk(projectKwargs, operatorFundusTemerloh, 'defaulter');
  copyOperatorCaseToKk(projectKwargs, operatorFundusTemerloh, 'selesai');

  let operatorFundusMentakab = SpreadsheetApp.openByUrl(projectKwargs['operator_fundus_mentakab']);
  // Do note that it is important to run for sheet defaulter first only then proceed to sheet selesai
  copyOperatorCaseToKk(projectKwargs, operatorFundusMentakab, 'defaulter');
  copyOperatorCaseToKk(projectKwargs, operatorFundusMentakab, 'selesai');
}

/**
 * Run trigger clinic cleanup.
 */
function triggerClinicCleanup() {
  Logger.log('Running trigger...');
  let projectKwargs = loadProjectKwargs();
  let listedClinic = [
    'kk_temerloh',
    'kk_tanjung_lalang',
    'kk_bandar_mentakab',
    'kk_lanchang',
    'kk_sanggang',
    'kk_kerdau',
    'kk_kuala_tekal',
    'kk_kuala_krau'
  ];
  listedClinic.forEach(clinic => {
    let bukuDaftarSpreadsheet = SpreadsheetApp.openByUrl(projectKwargs[clinic]);
    caseReviewCleanup(bukuDaftarSpreadsheet);
  });
}
