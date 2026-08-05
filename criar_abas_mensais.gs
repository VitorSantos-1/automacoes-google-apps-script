function criarAbasMensais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var template = ss.getSheetByName("PADRÃO");
  var ano = new Date().getFullYear(); // Pega o ano atual
  var meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
              "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

  // Verifica se a aba padrão existe
  if (!template) {
    SpreadsheetApp.getUi().alert("Erro: A aba PADRÂO não foi encontrada!");
    return;
  }

  // Cria cada aba mensal
  meses.forEach(function(mes) {
    var nomeAba = mes + " de " + ano;

    // Verifica se a aba já existe
    if (!ss.getSheetByName(nomeAba)) {
      var novaAba = template.copyTo(ss);
      novaAba.setName(nomeAba);

      // Atualiza os cabeçalhos
      novaAba.getRange('A1').setValue(mes);
      novaAba.getRange('D1').setValue(ano);

      // Posiciona a nova aba após a última existente
      ss.setActiveSheet(novaAba);
      ss.moveActiveSheet(ss.getSheets().length);
    }
  });

  // Mantém a aba PADRÂO oculta
  template.hideSheet();
  ss.toast('Abas mensais criadas com sucesso para ' + ano + '!');
}

// Adiciona menu personalizado
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 Ferramentas')
    .addItem('Criar Abas Mensais', 'criarAbasMensais')
    .addToUi();
}