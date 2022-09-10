var dado = abaMain.getRange('A1').getValue();

if(dado != '')
{
  abaMain.getRange('A1').setBackground("red");
  abaMain.getRange('A1').setFontColor('white');

  ui.alert('⚠️ Title!', 'Célula cheia! \n\n O que você vai fazer agora?.', ui.ButtonSet.OK_CANCEL);
}
