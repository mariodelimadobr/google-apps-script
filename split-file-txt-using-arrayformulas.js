const shApp = SpreadsheetApp.getActiveSpreadsheet();
const abaMain = shApp.getSheetByName("PAINEL");
const abaWork = shApp.getSheetByName("PROCESSADOR");
const abaData = shApp.getSheetByName("DADOS");
const abaLayout = shApp.getSheetByName("LAYOUT");
const abaLocais = shApp.getSheetByName("LocalidadesPMJ");
const abaMenuDep = shApp.getSheetByName("Menu_Dependente");

const abaHistorico = shApp.getSheetByName('HISTORICO');

//GET DADOS DO USUÁRIO
const email_user = Session.getActiveUser().getEmail();

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ PROCESADOR')
    //.addItem('LIMPAR', 'clearSheets')
    //.addItem('Importar Arquivo', 'importFile')
    .addItem('EXECUTAR', 'executeScript')
    .addToUi();

    abaLayout.hideSheet();
    abaLocais.hideSheet();
    abaMenuDep.hideSheet();

  abaMain.getRange('A3').activate();
}

function executeScript() {

      clearFilter();

      clearSheets();

      importFile();

      getDataFile(); //pega as colunas necessárias, conforme aba LAYOUT

      copyInsertData(); //xk duplicidade e, se não existir, copia para aba DADOS

      clearSheetWork();

      //SpreadsheetApp.flush();

      shApp.setActiveSheet(abaMain);
}

function clearSheetWork(){
  shApp.setActiveSheet(abaWork);

  abaWork.getRange(2, 1, abaWork.getMaxRows(), abaWork.getMaxColumns()).clearContent();
 
  abaWork.getRange(3,1).activate();

  var currentCell = abaWork.getCurrentCell();

  abaWork.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();

  currentCell.activateAsCurrentCell();
  
  abaWork.deleteRows(abaWork.getActiveRange().getRow(), abaWork.getActiveRange().getNumRows());

}

function clearFilter() {
  shApp.setActiveSheet(abaWork);

  var activeRange = abaWork.getRange(1, 1, abaWork.getMaxRows(), abaWork.getMaxColumns()).activate();
  var filter = abaWork.getFilter();

  if (filter === null) {
    activeRange.createFilter();
  }
    activeRange.getFilter().remove();
}

function clearSheets(){
  shApp.setActiveSheet(abaWork);

  abaWork.getRange(2, 1, abaWork.getMaxRows(), abaWork.getMaxColumns()).clear({contentsOnly: true, skipFilteredRows: true});
  
  abaWork.getRange('A2').activate();

  abaMain.getRange('C4').activate();

  //SpreadsheetApp.getUi().alert("👍 Sucesso!", 'Base limpa. \n\n Retorne à aba PAINEL. \n\n Selecione Mês e Ano. \n\n Siga as orientações.', SpreadsheetApp.getUi().ButtonSet.OK);
  
  SpreadsheetApp.getActive().toast('Limpeza','Executada');
  
  ////SpreadsheetApp.flush();
}

function importFile() {

  var arquivo = abaMain.getRange('B7').getValue();
  var pasta = DriveApp.getFolderById('1qE66odB0YyUkUouMW_gcRj7PaQ9jiMqt'); // OFICIAL
  //var pasta = DriveApp.getFolderById('1joIDUK87MC5coulc75wS6rDfoEXjK1o2'); // DEV

  shApp.setActiveSheet(abaWork);

  if (pasta.getFilesByName(arquivo).hasNext())
  {
    var arquivo = pasta.getFilesByName(arquivo).next();
    var csvData = Utilities.parseCsv(arquivo.getBlob().getDataAsString('ISO-8859-1'), ';');

    var ultLinhaHistorico = abaHistorico.getLastRow();
    var verarquivo = Utilities.formatDate(arquivo.getLastUpdated(), 'GMT-03:00', 'dd/MM/yyyy HH:mm:ss');
    var ultVerHistorico;

    if (ultLinhaHistorico > 1) 
    {
      ultVerHistorico = Utilities.formatDate(abaHistorico.getRange("D" + ultLinhaHistorico).getValue(), 'GMT-03:00', 'dd/MM/yyyy HH:mm:ss');
    }
    else
    {
      ultVerHistorico = '00/00/00 00:00:0000';
    }

    if (ultVerHistorico != verarquivo)
    {
      //DEPEJA OS DADOS DO ARQUIVO
      abaWork.getRange('A2').activate();
      abaWork.getRange(2, 1, csvData.length, csvData[0].length).setValues(csvData)

      //AJUSTA LARGURA DAS COLUNAS
      //abaWork.autoResizeColumns(2, 20);

      var linhasArquivo = abaWork.getLastRow() - 1;

      ultLinhaHistorico = ultLinhaHistorico + 1;
      abaHistorico.getRange('A' + ultLinhaHistorico).setValue(arquivo);
      abaHistorico.getRange('B' + ultLinhaHistorico).setValue(linhasArquivo);
      abaHistorico.getRange('C' + ultLinhaHistorico).setValue('Processado');
      abaHistorico.getRange('D' + ultLinhaHistorico).setValue(arquivo.getLastUpdated());
      abaHistorico.getRange('E' + ultLinhaHistorico).setValue(email_user);

      //SpreadsheetApp.getUi().alert('👍 Sucesso!','Importação concluída.', SpreadsheetApp.getUi().ButtonSet.OK);
      SpreadsheetApp.getActive().toast('Importação','Executada');
    } 
    else 
    {
      shApp.setActiveSheet(abaHistorico);
      
      //SpreadsheetApp.flush();

      SpreadsheetApp.getUi().alert('⚠️ Atenção!', 'Um arquivo com mesmo nome já foi processado anteriormente. \n\n Verique a versão do arquivo. \n\n Veja os registros na aba HISTORICO.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } 
  else 
  {
    SpreadsheetApp.getUi().alert('⚠️ Ops!', 'Arquivo não encontrado na pasta', SpreadsheetApp.getUi().ButtonSet.OK);
    
    shApp.setActiveSheet(abaMain);
  }
}

function getDataFile() {
  shApp.setActiveSheet(abaWork);

  abaWork.getRange('A1').setValue('Fonte de Dados 👇');

  //extrai data, primeira coluna do arquivo
  abaWork.getRange('B1').setValue('Referência');
  abaWork.getRange('B2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA( TRIM( IF( MID( $A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) >1; "01"; MID( $A$2:$A; (LAYOUT!B2+6); LAYOUT!D2 ) ) ) & "/" & TRIM( MID($A$2:$A; (LAYOUT!B2+4); (LAYOUT!D2-6) ) ) & "/" & TRIM(MID($A$2:$A; LAYOUT!B2; (LAYOUT!D2-4) )))))');

  //extrai unidade consumidora
  abaWork.getRange('C1').setValue('Unidade Consumidora');
  abaWork.getRange('C2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B3; LAYOUT!D3)))))');
  
 //extrai 
  abaWork.getRange('D1').setValue('KW');
  abaWork.getRange('D2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B4; LAYOUT!D4)))))');
 
 //extrai 
  abaWork.getRange('E1').setValue('COSIP');
  abaWork.getRange('E2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B5; LAYOUT!D5)))))');

 //extrai 
  abaWork.getRange('F1').setValue('Valor Total');
  abaWork.getRange('F2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B6; LAYOUT!D6)))))');

 //extrai 
  abaWork.getRange('G1').setValue('Ano');
  abaWork.getRange('G2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B7; LAYOUT!D7)))))');

 //extrai 
  abaWork.getRange('H1').setValue('M1');
  abaWork.getRange('H2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B8; LAYOUT!D8)))))');

 //extrai 
  abaWork.getRange('I1').setValue('Mês');
  abaWork.getRange('I2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(switch(H2:H;"01";"JAN";"02";"FEV";"03";"MAR";"04";"ABR";"05";"MAI";"06";"JUN";"07";"JUL";"08";"AGO";"09";"SET";"10";"OUT";"11";"NOV";"12";"DEZ";"NONE"))))');

 //extrai 
  abaWork.getRange('J1').setValue('Sequência');
  abaWork.getRange('J2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B10; LAYOUT!D10)))))');

 //extrai 
  abaWork.getRange('K1').setValue('Endereço Celesc');
  abaWork.getRange('K2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(TRIM(MID($A$2:$A; LAYOUT!B11; LAYOUT!D11)))))');

 //extrai 
  abaWork.getRange('L1').setValue('Grupo Tarifário');
  abaWork.getRange('L2').setFormula('=ARRAYFORMULA(IF(ISBLANK($A$2:$A);""; ARRAYFORMULA(IF(TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12)) = "4a"; "B4a"; TRIM(MID($A$2:$A; LAYOUT!B12; LAYOUT!D12))))))');

  abaWork.getRange('Z1').setValue('=CONCATENATE("Duplicadas" & CHAR(10) & SUM(Z2:Z))');
  abaWork.getRange('Z2').setFormula('=ARRAYFORMULA(IF(ISBLANK(B2:B);""; COUNTIFS(DADOS!A2:A; B2:B; DADOS!B2:B; C2:C)))');

 //query na aba Localidades 
  //abaWork.getRange('M1').setValue('Sigla');
  //abaWork.getRange('M2').setFormula('=QUERY(LocalidadesPMJ!$A$2:$Z; "SELECT K, L WHERE A like '%"&C2&"%' "; 0)');

  //SpreadsheetApp.flush();
};


function somaCol() {

  var spreadsheet = SpreadsheetApp.getActive();

  var guiadados = spreadsheet.getSheetByName("PROCESSADOR");

  var area = guiadados.getRange("Z2:Z" + guiadados.getLastRow() );

  var dados = area.getValues();

  var total = "0";

  // length ate o final da lista   linha++ aumentando posiçao por posição
  for (var linha = 0; linha < dados.length; linha++) {

    var total = Number(dados[linha][0]);

    total = Number(total) + Number(total);

  }
  
  //Logger.log(total);
  
}


function copyInsertData() {

  shApp.setActiveSheet(abaWork);
  
  if(somaCol() >= 1) {
    SpreadsheetApp.getUi().alert('⏰ Atenção!', 'Os dados do arquivo processado já encontram-se na base. \n\n Verifique na aba DADOS se há datas e unidades iguais às da aba PROCESSADOR. \n\n Dúvidas: entrar em contato com SAP.UAO', SpreadsheetApp.getUi().ButtonSet.OK);
  }
  else
  {
    copiarParaAbaData();
  }

  //SpreadsheetApp.flush();

}

function copiarParaAbaData(){
  
  shApp.setActiveSheet(abaWork);
  let lastRowSheetWork = abaWork.getLastRow();

  shApp.setActiveSheet(abaData);

  abaData.getRange('A1').setValue('Referência');
  abaData.getRange('B1').setValue('Unidade Consumidora');
  abaData.getRange('C1').setValue('KW');
  abaData.getRange('D1').setValue('COSIP');
  abaData.getRange('E1').setValue('Valor Total');
  abaData.getRange('F1').setValue('Ano');
  abaData.getRange('G1').setValue('M1');
  abaData.getRange('H1').setValue('Mês');
  abaData.getRange('I1').setValue('Sequência');
  abaData.getRange('J1').setValue('Endereço Celesc');
  abaData.getRange('K1').setValue('Grupo Tarifário');

  let validacao = abaData.getRange("A1").getValue();

  if(validacao == ""){

    var rowInicio = 1;
    var lastRow = 1;

  }else{

    var rowInicio = 2;
    var lastRow = abaData.getLastRow() + 1;

  }
  
  abaData.insertRowsAfter(abaData.getMaxRows(), lastRowSheetWork -1);

  let area = abaWork.getRange( "B2" +  ":L" + lastRowSheetWork ).getValues();

  abaData.getRange("A" + lastRow + ":K" + (lastRow + area.length - 1)).setValues(area);

  shApp.setActiveSheet(abaWork);
  abaWork.getRange(2, 1, abaWork.getMaxRows(), abaWork.getMaxColumns()).clear({contentsOnly: true, skipFilteredRows: true});
  
  abaWork.getRange('A2').activate();
  
  SpreadsheetApp.getUi().alert('👍 Sucesso!', 'Arquivo processado!.', SpreadsheetApp.getUi().ButtonSet.OK);
  
  //SpreadsheetApp.flush();

}

//COPIAR PARA OUTRA PLANILHA
//const ssId = "1UqnWpZgf1DhkuNe-LxeFpzFhTs7o2oUCoqIyF8Wutnw"
//const ssBase = SpreadsheetApp.openById(ssId);
//const sheetBase = ssBase.getSheetByName("BASE");



