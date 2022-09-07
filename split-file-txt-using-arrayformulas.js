function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('PROCESADOR')
    .addItem('LIMPAR', 'clearSheets')
    .addItem('EXECUTAR', 'executeScript')
    .addToUi();
}

const myApp = SpreadsheetApp.getActiveSpreadsheet();
const mySheetMain = myApp.getSheetByName("PAINEL");
const mySheetWork = myApp.getSheetByName("TEST");
const myFileName = mySheetMain.getRange('B6').getValue();
//const myFile = DriveApp.getFilesByName(myFileName).next().getBlob().getDataAsString();
const myFolder = DriveApp.getFolderById('1qE66odB0YyUkUouMW_gcRj7PaQ9jiMqt');

function executeScript() {
  resetFilter
  //clearSheets();
  importFile();
  splitFileTxt();
  //setFormulasToValues();
}

function resetFilter() {
  var activeRange = mySheetWork.getRange(1, 1, mySheetWork.getMaxRows(), mySheetWork.getMaxColumns()).activate();
  var filter = mySheetWork.getFilter();

  if (filter === null) {
    activeRange.createFilter();
  }
    activeRange.getFilter().remove();
}

function clearSheets(){
  resetFilter();

  mySheetWork.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  mySheetWork.getRange('A1').setValue('Fonte de Dados 👇');
  mySheetWork.getRange('A2').activate();
  SpreadsheetApp.getUi().alert('Ok! Base limpa. Retornar ao PAINEL e continuar.'); //exibe msg alert

  mySheetMain.getRange('A2').activate();
}

function importFile(){

 //EM DESENVOLVIMENTO

}

function splitFileTxt() {
  //extrai data, primeira coluna do arquivo
  mySheetWork.getRange('B1').setValue('Referência');
  //mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B2; LAYOUT!D2)))))');
  mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA( TRIM( MID( $A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) ) & "/" & TRIM( MID($A$2:$A; (LAYOUT!B2+4); (LAYOUT!D2-6) ) ) & "/" & TRIM(MID($A$2:$A; LAYOUT!B2; (LAYOUT!D2-4) )))))');

  //extrai unidade consumidora
  mySheetWork.getRange('C1').setValue('Unidade Consumidora');
  mySheetWork.getRange('C2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B3; LAYOUT!D3)))))');
  
 //extrai 
  mySheetWork.getRange('D1').setValue('KW');
  mySheetWork.getRange('D2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B4; LAYOUT!D4)))))');
 
 //extrai 
  mySheetWork.getRange('E1').setValue('COSIP');
  mySheetWork.getRange('E2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B5; LAYOUT!D5)))))');

 //extrai 
  mySheetWork.getRange('F1').setValue('Valor Total');
  mySheetWork.getRange('F2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B6; LAYOUT!D6)))))');

 //extrai 
  mySheetWork.getRange('G1').setValue('Ano');
  mySheetWork.getRange('G2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B7; LAYOUT!D7)))))');

 //extrai 
  mySheetWork.getRange('H1').setValue('M1');
  mySheetWork.getRange('H2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B8; LAYOUT!D8)))))');

 //extrai 
  mySheetWork.getRange('I1').setValue('Mês');
  mySheetWork.getRange('I2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(switch(H2:H;"01";"JAN";"02";"FEV";"03";"MAR";"04";"ABR";"05";"MAI";"06";"JUN";"07";"JUL";"08";"AGO";"09";"SET";"10";"OUT";"11";"NOV";"12";"DEZ";"NONE"))))');

 //extrai 
  mySheetWork.getRange('J1').setValue('Sequência');
  mySheetWork.getRange('J2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B10; LAYOUT!D10)))))');

 //extrai 
  mySheetWork.getRange('K1').setValue('Celesc');
  mySheetWork.getRange('K2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B11; LAYOUT!D11)))))');

 //extrai 
  mySheetWork.getRange('L1').setValue('Grupo Tarifário');
  mySheetWork.getRange('L2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12)))))');

 //extrai 
  mySheetWork.getRange('M1').setValue('Grupo Tarifário');
  //mySheetWork.getRange('M2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);"";ARRAYFORMULA(IF($L$2:$L = "4a"; "B4a"; $L$2:$L))))');
  mySheetWork.getRange('M2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(IF(TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12)) = "4a"; "B4a"; TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12))))))');

  mySheetWork.getRange('A1').activate();
  var ultimaLinha = mySheetWork.getLastRow();

  mySheetWork.getRange('A1').activate();

  SpreadsheetApp.getUi().alert('Arquivo processado com sucesso!'); //exibe msg alert
};

function setFormulasToValues() {
  var range = mySheetWork.getRange('A:Z').activate();
  range.copyValuesToRange(mySheetWork, 1, range.getLastColumn(), 1, range.getLastRow());
};
