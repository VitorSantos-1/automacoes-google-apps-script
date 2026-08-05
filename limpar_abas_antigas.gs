function apagarAbasOcultasComMaisDe30Dias() {
  var LIMITE_DIAS = 30;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var hoje = new Date();
  var mesAtual = hoje.getMonth(); // 0 a 11
  var anoAtual = hoje.getFullYear();
  var msPorDia = 1000 * 60 * 60 * 24;

  // Regex para capturar DD/MM ou DD/MM/AAAA
  var regexData = /(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?/;

  sheets.forEach(function(sheet) {
    // Só queremos abas ocultas
    if (!sheet.isSheetHidden()) return;

    var nome = sheet.getName();
    var match = nome.match(regexData);

    if (!match) {
      // Se não tiver data no nome, ignora
      return;
    }

    var dia = parseInt(match[1], 10);
    var mes = parseInt(match[2], 10) - 1; // meses em JS vão de 0 a 11
    var ano;

    if (match[3]) {
      // Se o ano vier explícito no nome, usa ele
      ano = parseInt(match[3], 10);
    } else {
      // Lógica inteligente: se o mês da aba for MAIOR que o mês atual,
      // significa que é do ano passado
      if (mes > mesAtual) {
        ano = anoAtual - 1;
      } else {
        ano = anoAtual;
      }
    }

    var dataAba = new Date(ano, mes, dia);
    var diffDias = Math.floor((hoje - dataAba) / msPorDia);

    if (diffDias > LIMITE_DIAS) {
      Logger.log('Apagando aba oculta: "' + nome + '" (data interpretada: ' + dia + '/' + (mes+1) + '/' + ano + ', há ' + diffDias + ' dias)');
      ss.deleteSheet(sheet);
    }
  });
}