// =====================================================================
// CONFIGURAÇÃO GLOBAL
// =====================================================================
var ABAS_IGNORADAS  = ['Dados - Consolidado', 'Banco de dados', 'modelo_para_copiar'];
var NOME_ABA_DADOS  = 'Banco de dados'; // fonte única — evita typo

var COLUNAS_OBRIGATORIAS = [
  { col: 13, nome: "M" },
  { col: 16, nome: "P" },
  { col: 17, nome: "Q" },
  { col: 18, nome: "R" },
  { col: 19, nome: "S" },
  { col: 20, nome: "T" },
  { col: 21, nome: "U" }
  // { col: x, nome: "y" }, // para adicionar mais coluna nessa tupla
];

var COLUNAS_LIVRES = COLUNAS_OBRIGATORIAS.map(function(c) { return c.col; });

// =====================================================================
// HELPERS GLOBAIS
// =====================================================================
function isAbaValidada(nomeAba) {
  return !ABAS_IGNORADAS.includes(nomeAba);
}

function isValorPermitido(valor) {
  return valor === "" || valor === null || valor === undefined ||
         valor === 1  || valor === "1"  ||
         valor === 9  || valor === "9";
}

// Remove acentos, troca espaços por _ e deixa minúsculo
// Era inline no Código 2 — agora reutilizável globalmente
function limparTexto(texto) {
  return String(texto)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_")
    .toLowerCase()
    .trim();
}

// =====================================================================
// onOpen
// =====================================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚫 Controle Rígido')
    .addItem('🔍 Verificar Pendências',      'verificarPendencias')
    .addItem('🎯 Ir para Primeira Pendente', 'irParaPrimeiraPendente')
    .addItem('📊 Relatório Simples',         'relatorioSimples')
    .addSeparator()
    .addItem('⚙️ Configurar Abas',           'configurarAbas')
    .addToUi();
}

// =====================================================================
// onEdit — UNIFICADO
// Ordem de execução:
//   1. Guarda de entrada (aba ignorada, linha de cabeçalho)
//   2. Data automática (Coluna A)
//   3. Lógica de bloqueio — apaga e redireciona se necessário
//   4. Lógica de dropdown — só roda se não houve bloqueio
//   5. Auto-status RECEBEU (Horários)
//   6. Auto-status RECEBEU (Coluna M = 1)
// =====================================================================
function onEdit(e) {
  try {
    if (!e || !e.source || !e.range) return;

    var sheet   = e.source.getActiveSheet();
    var nomeAba = sheet.getName();

    // Abas ignoradas saem aqui 
    if (!isAbaValidada(nomeAba)) return;

    var range = e.range;
    var row   = range.getRow();
    var col   = range.getColumn();

    if (row === 1) return;

    // -----------------------------------------------------------------
    // NOVO BLOCO — Data automática na Coluna A (1)
    // Preenche com a data de hoje se a linha for >= 7 e a célula A estiver vazia.
    // -----------------------------------------------------------------
    if (row >= 7) {
      var celulaData = sheet.getRange(row, 1);
      if (celulaData.getValue() === "") {
        var dataAtual = new Date();
        // Zera as horas para manter os dados analíticos limpos (00:00:00)
        dataAtual.setHours(0, 0, 0, 0); 
        celulaData.setValue(dataAtual);
      }
    }

    // -----------------------------------------------------------------
    // BLOCO 1 — Validação de pendências 
    // -----------------------------------------------------------------
    var bloqueado = false;

    if (!COLUNAS_LIVRES.includes(col)) {
      var pendencia = encontrarPrimeiraPendencia(sheet, row);

      if (pendencia !== null) {
        bloqueado = true;

        var valorDigitado = range.getValue();
        if (valorDigitado !== "" && valorDigitado !== null && valorDigitado !== undefined) {
          range.setValue("");
        }

        var detalhesPendencia = montarDetalhesPendencia(pendencia.linha, pendencia.colunasFaltando);

        try {
          var ui = SpreadsheetApp.getUi();
          ui.alert(
            '🚫 BLOQUEIO ATIVO - CONTEÚDO APAGADO',
            '❌ Sua digitação foi APAGADA automaticamente!\n\n' +
            '🎯 PENDÊNCIA DETECTADA na linha: ' + pendencia.linha + '\n\n' +
            detalhesPendencia + '\n\n' +
            '✅ Complete os campos acima primeiro!\n' +
            '🔄 Você será redirecionado agora.',
            ui.ButtonSet.OK
          );
        } catch (alertError) {}

        try {
          sheet.getRange(pendencia.linha, pendencia.colunasFaltando[0]).activate();
        } catch (redirectError) {}
      }
    }

    // -----------------------------------------------------------------
    // BLOCO 2 — Dropdown de Resolução 
    // -----------------------------------------------------------------
    if (!bloqueado && row >= 6 && (col === 15 || col === 17)) {
      _atualizarDropdownResolucao(e, sheet, row);
    }

    // -----------------------------------------------------------------
    // BLOCO 3 — Auto-status RECEBEU (colunas F, G, H, I → P)
    // -----------------------------------------------------------------
    var COLUNAS_HORARIO = [6, 7, 8, 9]; // F, G, H, I

    if (!bloqueado && row >= 7 && COLUNAS_HORARIO.includes(col)) {
      _atualizarStatusRecebeu(sheet, row);
    }

    // -----------------------------------------------------------------
    // NOVO BLOCO — Auto-status RECEBEU (coluna M = 1)
    // -----------------------------------------------------------------
    if (!bloqueado && row >= 7 && col === 13) {
      var valorM = range.getValue();
      // Verifica se o valor digitado é numérico 1 ou string "1"
      if (valorM === 1 || valorM === "1") {
        var celulaStatus = sheet.getRange(row, 16); // Coluna P
        if (celulaStatus.getValue() !== "RECEBEU") {
          celulaStatus.setValue("RECEBEU");
        }
      }
    }

  } catch (error) {
    console.error('Erro no onEdit unificado:', error);
  }
}

// =====================================================================
// Atualiza o dropdown de Resolução (col R = 18)
// Acionado quando O (15) = Ocorrência ou Q (17) = Setor mudam
// =====================================================================
function _atualizarDropdownResolucao(e, sheet, row) {
  if (sheet.getMaxColumns() < 18) return;

  var valorO  = sheet.getRange(row, 15).getValue(); // Ocorrência
  var valorQ  = sheet.getRange(row, 17).getValue(); // Setor
  var celulaR = sheet.getRange(row, 18);            // Resolução

  celulaR.clearDataValidations();
  celulaR.clearContent();

  if (!valorO || !valorQ) return;

  // Chave: setor_ocorrencia (ex: "comercial_divergencia_de_custo")
  var chave = limparTexto(valorQ) + "_" + limparTexto(valorO);

  var abaDados = e.source.getSheetByName(NOME_ABA_DADOS);
  if (!abaDados) return;

  var cabecalhos = abaDados.getRange("F1:AH1").getValues()[0].map(limparTexto);
  var indice     = cabecalhos.indexOf(chave);

  if (indice === -1) return;

  var opcoes = abaDados.getRange(2, 6 + indice, 10, 1).getValues()
    .map(function(r) { return r[0]; })
    .filter(function(texto) { return texto !== ""; });

  if (opcoes.length > 0) {
    var regra = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcoes, true)
      .setAllowInvalid(false)
      .build();
    celulaR.setDataValidation(regra);
  }
}

// =====================================================================
// Verifica se F, G, H, I estão todos preenchidos e atualiza STATUS (P)
// =====================================================================
function _atualizarStatusRecebeu(sheet, row) {
  var COLUNAS_HORARIO = [6, 7, 8, 9];
  var COL_STATUS      = 16;

  var todoPreenchido = COLUNAS_HORARIO.every(function(colNum) {
    var val = sheet.getRange(row, colNum).getValue();
    return val !== "" && val !== null && val !== undefined;
  });

  if (todoPreenchido) {
    var celulaStatus = sheet.getRange(row, COL_STATUS);
    // Só escreve se ainda não estiver como RECEBEU, evita loop desnecessário
    if (celulaStatus.getValue() !== "RECEBEU") {
      celulaStatus.setValue("RECEBEU");
    }
  }
}

// =====================================================================
// Monta texto descritivo das colunas pendentes
// =====================================================================
function montarDetalhesPendencia(linha, colunasFaltando) {
  var texto = '📝 Colunas pendentes:\n';
  for (var i = 0; i < colunasFaltando.length; i++) {
    var colNum  = colunasFaltando[i];
    var nomeCol = COLUNAS_OBRIGATORIAS.find(function(c) { return c.col === colNum; }).nome;
    texto += '   • ' + nomeCol + linha + ' deve ser preenchida\n';
  }
  return texto;
}

// =====================================================================
// Encontra a primeira linha anterior com pendência
// =====================================================================
function encontrarPrimeiraPendencia(sheet, linhaAtual) {
  try {
    var linhaAnterior = linhaAtual - 1;
    if (linhaAnterior < 7) return null;

    var startRow = Math.max(7, linhaAnterior - 299);
    var numRows  = linhaAnterior - startRow + 1;
    if (numRows <= 0) return null;

    var dadosM       = sheet.getRange(startRow, 13, numRows, 1).getValues();
    var dadosColunas = {};

    for (var c = 0; c < COLUNAS_OBRIGATORIAS.length; c++) {
      var colNum = COLUNAS_OBRIGATORIAS[c].col;
      dadosColunas[colNum] = sheet.getRange(startRow, colNum, numRows, 1).getValues();
    }

    for (var i = 0; i < numRows; i++) {
      var valorM = dadosM[i][0];
      if (!isValorPermitido(valorM)) {
        var colunasFaltando = [];
        for (var c = 0; c < COLUNAS_OBRIGATORIAS.length; c++) {
          var colNum   = COLUNAS_OBRIGATORIAS[c].col;
          var valorCol = dadosColunas[colNum][i][0];
          if (valorCol === "" || valorCol === null || valorCol === undefined) {
            colunasFaltando.push(colNum);
          }
        }
        if (colunasFaltando.length > 0) {
          return { linha: startRow + i, colunasFaltando: colunasFaltando };
        }
      }
    }

    return null;

  } catch (error) {
    console.error('Erro ao procurar pendências:', error);
    return null;
  }
}

// =====================================================================
// MENU: Verificar pendências
// =====================================================================
function verificarPendencias() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui    = SpreadsheetApp.getUi();

  if (!isAbaValidada(sheet.getName())) {
    ui.alert('Esta aba está excluída da validação');
    return;
  }

  var lastRow = Math.min(sheet.getLastRow(), 500);
  if (lastRow < 7) { ui.alert('Não há linhas de dados para verificar'); return; }

  var numRows      = lastRow - 7 + 1;
  var dadosM       = sheet.getRange(7, 13, numRows, 1).getValues();
  var dadosColunas = {};

  for (var c = 0; c < COLUNAS_OBRIGATORIAS.length; c++) {
    var colNum = COLUNAS_OBRIGATORIAS[c].col;
    dadosColunas[colNum] = sheet.getRange(7, colNum, numRows, 1).getValues();
  }

  var pendencias = [];

  for (var i = 0; i < numRows && pendencias.length < 20; i++) {
    var valorM = dadosM[i][0];
    if (!isValorPermitido(valorM)) {
      var faltando = [];
      for (var c = 0; c < COLUNAS_OBRIGATORIAS.length; c++) {
        var colNum   = COLUNAS_OBRIGATORIAS[c].col;
        var nomeCol  = COLUNAS_OBRIGATORIAS[c].nome;
        var valorCol = dadosColunas[colNum][i][0];
        if (valorCol === "" || valorCol === null || valorCol === undefined) {
          faltando.push(nomeCol);
        }
      }
      if (faltando.length > 0) {
        pendencias.push({ linha: 7 + i, valorM: valorM, faltando: faltando });
      }
    }
  }

  if (pendencias.length === 0) {
    ui.alert('✅ Sem pendências', 'Nenhuma pendência encontrada!', ui.ButtonSet.OK);
  } else {
    var mensagem = '⚠️ PENDÊNCIAS ENCONTRADAS: ' + pendencias.length + '\n\n';
    for (var j = 0; j < Math.min(pendencias.length, 10); j++) {
      mensagem += '• Linha ' + pendencias[j].linha +
                  ' (M=' + pendencias[j].valorM + ') → faltam: ' +
                  pendencias[j].faltando.join(', ') + '\n';
    }
    if (pendencias.length > 10) mensagem += '... e mais ' + (pendencias.length - 10) + ' pendências\n';
    mensagem += '\n🎯 Ir para a primeira pendência?';

    var resposta = ui.alert('📋 Relatório de Pendências', mensagem, ui.ButtonSet.YES_NO);
    if (resposta == ui.Button.YES) {
      var primeiraLetra = pendencias[0].faltando[0];
      var colDestino = COLUNAS_OBRIGATORIAS.find(function(c) {
        return c.nome === primeiraLetra;
      }).col;
      sheet.getRange(pendencias[0].linha, colDestino).activate();
    }
  }
}

// =====================================================================
// MENU: Ir para primeira pendente
// =====================================================================
function irParaPrimeiraPendente() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui    = SpreadsheetApp.getUi();

  if (!isAbaValidada(sheet.getName())) {
    ui.alert('Esta aba está excluída da validação');
    return;
  }

  var pendencia = encontrarPrimeiraPendencia(sheet, sheet.getLastRow() + 1);

  if (pendencia === null) {
    ui.alert('✅ Sem pendências', 'Nenhuma pendência encontrada!', ui.ButtonSet.OK);
  } else {
    sheet.getRange(pendencia.linha, pendencia.colunasFaltando[0]).activate();
    var nomesCol = pendencia.colunasFaltando.map(function(n) {
      return COLUNAS_OBRIGATORIAS.find(function(c) { return c.col === n; }).nome;
    });
    ui.alert(
      '🎯 Direcionado para pendência',
      'Linha: ' + pendencia.linha + '\n' +
      'Colunas pendentes: ' + nomesCol.join(', '),
      ui.ButtonSet.OK
    );
  }
}

// =====================================================================
// MENU: Relatório simples
// =====================================================================
function relatorioSimples() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui    = SpreadsheetApp.getUi();

  if (!isAbaValidada(sheet.getName())) {
    ui.alert('Esta aba está excluída da validação');
    return;
  }

  var amostra = Math.min(100, sheet.getLastRow() - 6);
  if (amostra <= 0) { ui.alert('Não há dados para analisar'); return; }

  var dadosM       = sheet.getRange(7, 13, amostra, 1).getValues();
  var dadosColunas = {};
  for (var c = 0; c < COLUNAS_OBRIGATORIAS.length; c++) {
    var colNum = COLUNAS_OBRIGATORIAS[c].col;
    dadosColunas[colNum] = sheet.getRange(7, colNum, amostra, 1).getValues();
  }

  var pendencias = 0, completas = 0, permitidas = 0;

  for (var i = 0; i < amostra; i++) {
    var valorM = dadosM[i][0];
    if (isValorPermitido(valorM)) {
      permitidas++;
    } else {
      var temFaltando = COLUNAS_OBRIGATORIAS.some(function(cfg) {
        var v = dadosColunas[cfg.col][i][0];
        return v === "" || v === null || v === undefined;
      });
      temFaltando ? pendencias++ : completas++;
    }
  }

  ui.alert(
    '📈 Análise Rápida',
    '📊 RELATÓRIO (amostra de ' + amostra + ' linhas)\n\n' +
    '✅ Completas: '               + completas  + '\n' +
    '⚪ Permitidas (M válido): '   + permitidas + '\n' +
    '❌ Pendências: '               + pendencias,
    ui.ButtonSet.OK
  );
}

// =====================================================================
// MENU: Configurar abas
// =====================================================================
function configurarAbas() {
  var ui         = SpreadsheetApp.getUi();
  var nomeAba    = SpreadsheetApp.getActiveSheet().getName();
  var isExcluida = ABAS_IGNORADAS.includes(nomeAba);

  var mensagem = 'Aba atual: "' + nomeAba + '"\n';
  mensagem += 'Status: ' + (isExcluida ? '⛔ EXCLUÍDA (sem validação)' : '✅ ATIVA (com validação)') + '\n\n';
  mensagem += 'Abas excluídas:\n' + (ABAS_IGNORADAS.length ? ABAS_IGNORADAS.join(', ') : 'nenhuma') + '\n\n';
  mensagem += isExcluida ? 'Deseja REATIVAR a validação nesta aba?' : 'Deseja EXCLUIR esta aba da validação?';

  var resposta = ui.alert('⚙️ Configurar Validação', mensagem, ui.ButtonSet.YES_NO);
  if (resposta == ui.Button.YES) {
    if (isExcluida) {
      ABAS_IGNORADAS.splice(ABAS_IGNORADAS.indexOf(nomeAba), 1);
      ui.alert('✅ Validação reativada nesta aba (temporariamente)', '', ui.ButtonSet.OK);
    } else {
      ABAS_IGNORADAS.push(nomeAba);
      ui.alert('⛔ Aba excluída da validação (temporariamente)', '', ui.ButtonSet.OK);
    }
  }
}