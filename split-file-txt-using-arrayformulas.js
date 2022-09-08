/*
https://spreadsheet.dev/
*/

const myApp = SpreadsheetApp.getActiveSpreadsheet();
const mySheetMain = myApp.getSheetByName("PAINEL");
const mySheetDados = myApp.getSheetByName("DADOS");
const mySheetWork = myApp.getSheetByName("PROCESSADOR");

const myFileName = mySheetMain.getRange('B7').getValue();
const myFolder = DriveApp.getFolderById('1qE66odB0YyUkUouMW_gcRj7PaQ9jiMqt');

let xk_col_fonte_de_dados = mySheetWork.getRange("A2").getValue();

//const lastRowSheetWork = mySheetWork.getLastRow() +1;
//const lastRowSheetDados = mySheetDados.getLastRow() +1;

//BUSCA ARQUIVO NO DRIVE PARA IMPORTAR
//const myFile = DriveApp.getFilesByName(myFileName).next().getBlob().getDataAsString();

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ PROCESADOR')
    .addItem('LIMPAR', 'clearSheets')
    .addItem('EXECUTAR', 'executeScript')
    .addToUi();

  mySheetMain.getRange('A3').activate();
}

function executeScript() {
  
    if(xk_col_fonte_de_dados != "")
    {
      resetFilter();

      //clearSheets();

      importFile();

      splitFileTxt();

      copiarParaAbaDados();

      mySheetWork.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

      SpreadsheetApp.flush();
    }
    else
    {
      SpreadsheetApp.getUi().alert("Ops!", 'Aba PROCESSADOR sem dados. \n\n Veja orientações na aba PAINEL.', SpreadsheetApp.getUi().ButtonSet.OK); //msg alert
      SpreadsheetApp.flush();
    }

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
  
  SpreadsheetApp.flush();
 
  mySheetWork.getRange('A1').setValue('Fonte de Dados 👇');
  mySheetWork.getRange('A2').activate();

  SpreadsheetApp.getUi().alert("Ok!", 'Base limpa. \n\n Retorne à aba PAINEL. \n\n Selecione Mês e Ano. \n\n Siga as orientações.', SpreadsheetApp.getUi().ButtonSet.OK); //msg alert
  
  SpreadsheetApp.flush();

  mySheetWork.getRange('A2').activate();

  mySheetMain.getRange('C4').activate();
  
  SpreadsheetApp.flush();
}

function importFile(){
 //EM DESENVOLVIMENTO

  SpreadsheetApp.flush();
}

function splitFileTxt() {
  //extrai data, primeira coluna do arquivo
  mySheetWork.getRange('B1').setValue('Referência');
  //mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B2; LAYOUT!D2)))))'); // OLD COMAND
  //mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA( TRIM(IF(B2>1;"01")MID($A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) ) & "/" & TRIM( MID($A$2:$A; (LAYOUT!B2+4); (LAYOUT!D2-6) ) ) & "/" & TRIM(MID($A$2:$A; LAYOUT!B2; (LAYOUT!D2-4) )))))');

  //DEV CONVERT PARA 1º DIA DO MÊS
  mySheetWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA( TRIM( IF( MID( $A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) >1; "01"; MID( $A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) ) ) & "/" & TRIM( MID($A$2:$A; (LAYOUT!B2+4); (LAYOUT!D2-6) ) ) & "/" & TRIM(MID($A$2:$A; LAYOUT!B2; (LAYOUT!D2-4) )))))');

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
  mySheetWork.getRange('L2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(IF(TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12)) = "4a"; "B4a"; TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12))))))');

  mySheetWork.getRange('Z1').setValue('=CONCATENATE("Duplicadas" & CHAR(10) & SUM(L2:L))');
  mySheetWork.getRange('Z2').setFormula('=ARRAYFORMULA(IF(ISBLANK(B2:B);""; COUNTIFS(DADOS!A2:A; B2:B; DADOS!B2:B; C2:C)))');

 //busca na aba Localidades 
  //mySheetWork.getRange('M1').setValue('Sigla');
  //mySheetWork.getRange('M2').setFormula('=QUERY(LocalidadesPMJ!$A$2:$Z; "SELECT K, L WHERE A like '%"&C2&"%' "; 0)');

  SpreadsheetApp.flush();
};


function sumColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s1 = ss.getSheetByName("PROCESSADOR");
  var dataRange = s1.getDataRange();
  var lastrow = dataRange.getLastRow();
  var values = s1.getRange(1, 26, lastrow, 1).getValues();
  var result = 0;
  for (var i = 0; i < values.length; i++) {
    result += typeof values[i][0] == 'number' ? values[i][0] : 0;
    //Logger.log(result);
  }
  //s1.getRange(i + 1, 26).setValue(result);
}


//COPIAR PARA DA ABA PROCESSADOR PARA ABA DADOS
function copiarParaAbaDados(){

  if( sumColumn() < 1 ) 
  {
      let lastRowSheetWork = mySheetWork.getLastRow();

      let validacao = mySheetDados.getRange("A1").getValue();

      if(validacao == ""){

        var rowInicio = 1;
        var lastRow = 1;

      }else{

        var rowInicio = 2;
        var lastRow = mySheetDados.getLastRow() + 1;

      }
      
      mySheetDados.insertRowsAfter(mySheetDados.getMaxRows(), lastRowSheetWork -1);

      let area = mySheetWork.getRange( "B2" +  ":L" + lastRowSheetWork ).getValues();

      mySheetDados.getRange("A" + lastRow + ":K" + (lastRow + area.length - 1)).setValues(area);

      mySheetDados.getRange("A1").activate();

      SpreadsheetApp.getUi().alert('Ok! \n\n Arquivo processado com sucesso!'); //exibe msg alert

      SpreadsheetApp.flush();
  }
  else
  {
    SpreadsheetApp.getUi().alert('Ops!', 'Duplicidade detectada! \n\n Os registros do arquivo processado já encontram-se na base de dados. \n\n Veja orientações na aba PAINEL. \n\n Dúvidas: entrar em contato com SAP.UAO', SpreadsheetApp.getUi().ButtonSet.OK); //msg alert

     SpreadsheetApp.flush();
  }

}

//COPIAR PARA OUTRA PLANILHA
//const ssId = "1UqnWpZgf1DhkuNe-LxeFpzFhTs7o2oUCoqIyF8Wutnw"
//const ssBase = SpreadsheetApp.openById(ssId);
//const sheetBase = ssBase.getSheetByName("BASE");
