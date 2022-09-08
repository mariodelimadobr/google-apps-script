const myApp = SpreadsheetApp.getActiveSpreadsheet();
const mySheetMain = myApp.getSheetByName("PAINEL");
const mySheetDados = myApp.getSheetByName("DADOS");
const mySheetWork = myApp.getSheetByName("TEST");
const myFileName = mySheetMain.getRange('B6').getValue();

const myFolder = DriveApp.getFolderById('1qE66odB0YyUkUouMW_gcRj7PaQ9jiMqt');

//DEFINE LOOP DA BUSCA PELO ARQUIVO NO DRIVE
//const myFile = DriveApp.getFilesByName(myFileName).next().getBlob().getDataAsString();

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('PROCESADOR')
    .addItem('LIMPAR', 'clearSheets')
    .addItem('EXECUTAR', 'executeScript')
    .addToUi();

  mySheetMain.getRange('A3').activate();
}

function executeScript() {
  
  resetFilter();

  //clearSheets();

  importFile();

  splitFileTxt();

  copiarParaDados();

  mySheetWork.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  //IMPORT
  SpreadsheetApp.getUi().alert('Arquivo processado com sucesso!'); //exibe msg alert
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

  mySheetWork.getRange('A2').activate();
  mySheetMain.getRange('A2').activate();
}

function importFile(){

 //EM DESENVOLVIMENTO

}

function splitFileTxt() {
  //extrai data, primeira coluna do arquivo
  mySheetWork.getRange('B1').setValue('Referência');
  //mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B2; LAYOUT!D2)))))'); // OLD COMAND
  //mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA( TRIM( MID( $A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) ) & "/" & TRIM( MID($A$2:$A; (LAYOUT!B2+4); (LAYOUT!D2-6) ) ) & "/" & TRIM(MID($A$2:$A; LAYOUT!B2; (LAYOUT!D2-4) )))))');

  //DEV CONVERT PARA 1º DIA DO MÊS
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
  mySheetWork.getRange('K1').setValue('Endereço Celesc');
  mySheetWork.getRange('K2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B11; LAYOUT!D11)))))');

 //extrai 
  mySheetWork.getRange('L1').setValue('Grupo Tarifário');
  //mySheetWork.getRange('L2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12)))))'); //OLD
  mySheetWork.getRange('L2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(IF(TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12)) = "4a"; "B4a"; TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12))))))');

 //busca na aba Localidades 
  //mySheetWork.getRange('M1').setValue('Sigla');
  //mySheetWork.getRange('M2').setFormula('=QUERY(LocalidadesPMJ!$A$2:$Z; "SELECT K, L WHERE A like '%"&C2&"%' "; 0)');

 //mySheetWork.getRange('N1').setValue('Secretaria');



  mySheetWork.getRange('A1').activate();
  var ultimaLinha = mySheetWork.getLastRow();

  mySheetWork.getRange('A1').activate();

};

function setFormulasToValues() {
  var range = mySheetWork.getRange('A:Z').activate();
  range.copyValuesToRange(mySheetWork, 1, range.getLastColumn(), 1, range.getLastRow());
};

function copiarParaDados() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:C1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('TEST'), true);
  spreadsheet.getRange('B2').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DADOS'), true);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('A2').activate();
  spreadsheet.getRange('TEST!B2:L600').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('PAINEL'), true);
};



function DEV_copiarParaDados() {

  mySheetWork.getRange('B2').activate();
  var currentCell = mySheetWork.getCurrentCell();

  mySheetWork.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = mySheetWork.getCurrentCell();
  mySheetWork.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();

  mySheetDados.getRange('A1').activate();
  mySheetDados.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();

  mySheetDados.getRange('A' + mySheetDados.getLastRow() + 1).activate();

  mySheetDados.getRange('A' + mySheetDados.getLastRow() + 1).copyTo(mySheetDados.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  
  mySheetDados.getRange('A1').activate();

  mySheetMain.getRange('A1').activate();
};
