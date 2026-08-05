/** * Sistema de oferta - V9.2 (Otimizado)
 * Vitor
 */

/* CONFIGURAÇÕES GERAIS
   ======================= */
const START_ROW = 2;
const END_ROW = 2000;
const COL_CODIGO = 4; // D
const COL_DESC = 5; // E
const COL_PRECO = 7; // G

const MIN_ROWS = 320;
const PASSWORD = "mds";
const PROTECTION_DESC = "Protegido por script";

const PASTE_PROTECTION_SHEETS = [
  "ALERTA OPÇÃO",
  "ENCARTE OPÇÃO",
  "OFERTA FDS OPÇÃO",
  "OFERTA VIRTUAL",
];

/* ZONAS (Unitário / Família)
   ============================ */
const ZONES = [
  {
    name: "Tabela 1",
    sheets: [], // vazio = qualquer aba
    areaA1: "D7:M320",
    checkCol: "M", // coluna onde fica o valor 1 ou 2
    productNameCol: "E", // coluna do nome do produto
    completedCol: "M", // coluna onde o cursor volta se cancelar
  },
];
const FIRST_DATA_OFFSET = 1;

/* CAPA / APP somente no encarte
   =================================== */
const ENCARTE_OPCAO_SHEET = "ENCARTE OPÇÃO";
const COL_H_VISIBILITY_SHEETS = [
  "ALERTA OPÇÃO",
  "ENCARTE OPÇÃO",
  "OFERTA VIRTUAL"
];

const MKT_COL = 10; // coluna J — onde se digita "Capa" ou "App"
const PRECO_APP_COL = 8; // coluna H — PREÇO APP (mostrar/ocultar)
const DATA_COLS_MOVE = [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]; // D até N
const ANCHOR_COL_MOVE = 5; // coluna E (Descrição) — verifica se slot está livre
const MAIN_DATA_START = 31; // primeira linha de produto após seções reservadas

const CAPA_APP_SECTIONS = {
  App: { firstRow: 7, lastRow: 11, maxItems: 5, label: "APP" },
  Capa: { firstRow: 13, lastRow: 28, maxItems: 16, label: "CAPA" },
};

/* GATILHO ÚNICO DE EDIÇÃO (INSTALÁVEL)
   ====================================== */
function onEditInfo(e) {
  try {
    // Evita conflitos de concorrência se a lixeira/limpeza estiver executando no mesmo momento
    const cache = CacheService.getScriptCache();
    if (cache.get("SCRIPT_RUNNING") === "true") return;

    if (!e || !e.range) return;

    const ss = e.source;
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const editedCell = e.range;
    const editedRow = editedCell.getRow();
    const editedCol = editedCell.getColumn();

    /* ── BLOCO 1: BLOQUEIO DE COLAR / ARRASTAR ──────────────────── */
    if (PASTE_PROTECTION_SHEETS.includes(sheetName)) {
      if (editedCell.getNumRows() > 1 || editedCell.getNumColumns() > 1) {
        editedCell.clearContent();
        SpreadsheetApp.getUi().alert(
          "🚫 AÇÃO BLOQUEADA",
          "Não é permitido colar intervalos ou arrastar células nesta aba.\n" +
          "Por favor, faça a entrada MANUALMENTE, item por item.",
          SpreadsheetApp.getUi().ButtonSet.OK,
        );
        return;
      }
    }

    /* ── BLOCO 2: CÓDIGO (D) — DUPLICIDADE + BUSCA EM OUTRAS ABAS ─ */
    if (
      editedRow >= START_ROW &&
      editedRow <= END_ROW &&
      editedCol === COL_CODIGO
    ) {
      const val = normalizeCode_(editedCell.getValue());
      if (!val) return;

      // 2.1 – Duplicidade na própria aba
      const codesHere = flatValues_(
        sheet.getRange(START_ROW, COL_CODIGO, END_ROW - START_ROW + 1, 1),
      );
      const dupCount = codesHere.reduce(
        (acc, v) => acc + (normalizeCode_(v) === val ? 1 : 0),
        0,
      );
      if (dupCount > 1) {
        editedCell.clearContent();
        SpreadsheetApp.getUi().alert(
          "⛔ Duplicidade Detectada",
          `O código "${val}" já existe nesta tabela ("${sheetName}").\n` +
          "O valor foi apagado. Digite um código único.",
          SpreadsheetApp.getUi().ButtonSet.OK,
        );
        return;
      }

      // 2.2 – Busca SOMENTE nas outras abas (exclui a aba atual)
      const encontrados = [];
      for (const sh of ss.getSheets()) {
        if (sh.isSheetHidden()) continue;
        if (sh.getName() === sheetName) continue; // pula a aba atual
        const targetCodes = flatValues_(
          sh.getRange(START_ROW, COL_CODIGO, END_ROW - START_ROW + 1, 1),
        );
        const idx = targetCodes.findIndex((v) => normalizeCode_(v) === val);
        if (idx !== -1) {
          const rowFound = idx + START_ROW;
          const desc = sh.getRange(rowFound, COL_DESC).getDisplayValue();
          const price = sh.getRange(rowFound, COL_PRECO).getValue();
          const priceFmt = isFiniteNumber_(price)
            ? `R$ ${Number(price).toFixed(2)}`
            : "(sem preço)";
          encontrados.push({
            sheet: sh.getName(),
            row: rowFound,
            desc,
            priceFmt,
          });
        }
      }

      if (encontrados.length > 0) {
        SpreadsheetApp.getUi().alert(
          "ℹ️ Produto encontrado em outras abas",
          formatInfoMessage_(val, encontrados),
          SpreadsheetApp.getUi().ButtonSet.OK,
        );
      }

      return;
    }

    /* CAPA / APP — MOVE / RETORNA LINHA ───────────────────── */
    if (COL_H_VISIBILITY_SHEETS.includes(sheetName) && editedCol === MKT_COL) {
      const tipoRaw = String(editedCell.getValue()).trim();

      // 4b — slot reservado foi apagado?
      const currentSection = Object.values(CAPA_APP_SECTIONS).find(
        (s) => editedRow >= s.firstRow && editedRow <= s.lastRow,
      );
      if (currentSection) {
        if (tipoRaw === "")
          returnFromSection_(sheet, currentSection, editedRow, ss, true);
        return;
      }

      // 4a — linha normal, valor é Capa/App?
      const tipoKey = Object.keys(CAPA_APP_SECTIONS).find(
        (k) => k.toLowerCase() === tipoRaw.toLowerCase(),
      );
      if (tipoKey && sheetName === ENCARTE_OPCAO_SHEET) {
        moveToSection_(
          sheet,
          CAPA_APP_SECTIONS[tipoKey],
          editedRow,
          tipoKey,
          ss,
        );
      }

      // 4c — atualiza coluna H
      atualizarVisibilidadeColunaH_(sheet);
      return;
    }

    /* UNITÁRIO / FAMÍLIA ────────────────────────────────── */
    const zone = detectZoneForRow_(sheet, editedRow);
    if (!zone) return;

    const area = safeGetRange_(sheet, zone.areaA1);
    if (!area) return;

    const areaStartRow = area.getRow();
    const areaEndRow = areaStartRow + area.getNumRows() - 1;
    const areaStartCol = area.getColumn();
    const areaEndCol = areaStartCol + area.getNumColumns() - 1;

    if (
      editedRow < areaStartRow ||
      editedRow > areaEndRow ||
      editedCol < areaStartCol ||
      editedCol > areaEndCol
    )
      return;

    const firstDataRow = areaStartRow + FIRST_DATA_OFFSET;
    if (editedRow < firstDataRow) return;

    const prevRow = editedRow - 1;
    if (prevRow < firstDataRow) return;

    const checkIdx = colToIndex_(zone.checkCol);
    const focusIdx = colToIndex_(zone.completedCol);
    const nameIdx = zone.productNameCol
      ? colToIndex_(zone.productNameCol)
      : null;

    if (
      isMergedCell_(sheet, prevRow, checkIdx) ||
      (nameIdx && isMergedCell_(sheet, prevRow, nameIdx)) ||
      isMergedCell_(sheet, prevRow, focusIdx)
    )
      return;

    const productPrev = nameIdx
      ? String(sheet.getRange(prevRow, nameIdx).getDisplayValue()).trim()
      : "";
    if (!productPrev) return;

    const checkValPrev = checkIdx
      ? String(sheet.getRange(prevRow, checkIdx).getDisplayValue()).trim()
      : "";

    const isValidUF = (v) => ["1", "2"].includes(String(v).trim());
    if (isValidUF(checkValPrev)) return;

    const ui = SpreadsheetApp.getUi();
    const hint = productPrev ? `Produto: ${productPrev}\n` : "";
    const resp = ui.prompt(
      'Definir "Unitário / Família"',
      `${hint}Linha ${prevRow}: Digite 1 para Família ou 2 para Unitário.`,
      ui.ButtonSet.OK_CANCEL,
    );

    if (resp.getSelectedButton() !== ui.Button.OK) {
      if (focusIdx && !isMergedCell_(sheet, prevRow, focusIdx)) {
        ss.setActiveSheet(sheet);
        sheet.setActiveSelection(sheet.getRange(prevRow, focusIdx));
      }
      ss.toast(
        "Preenchimento obrigatório pendente na linha " + prevRow + ".",
        "Atenção",
        5,
      );
      return;
    }

    const valUF = String(resp.getResponseText()).trim();
    if (!isValidUF(valUF)) {
      ui.alert("Valor inválido. Use apenas 1 (Família) ou 2 (Unitário)!");
      if (focusIdx && !isMergedCell_(sheet, prevRow, focusIdx)) {
        ss.setActiveSheet(sheet);
        sheet.setActiveSelection(sheet.getRange(prevRow, focusIdx));
      }
      return;
    }

    if (checkIdx && !isMergedCell_(sheet, prevRow, checkIdx)) {
      sheet.getRange(prevRow, checkIdx).setValue(Number(valUF));
    }

    ss.toast(
      `Linha ${prevRow} — ${valUF === "1" ? "Família" : "Unitário"} definido.`,
      "OK",
      3,
    );
  } catch (err) {
    console.error("Erro em onEditInfo:", err);
    try {
      SpreadsheetApp.getActive().toast(
        "Erro no onEditInfo: " + err.message,
        "Erro",
        5,
      );
    } catch (_) { }
  }
}

/* PROTEÇÃO / DESPROTEÇÃO DE COLUNAS
   ==================================== */
const SHEETS_CONFIG = {
  "ALERTA OPÇÃO": ["E", "K", "L", "N"],
  "ENCARTE OPÇÃO": ["E", "K", "L", "N"],
  "OFERTA FDS OPÇÃO": ["E", "K", "L", "N"],
  "OFERTA VIRTUAL": ["E", "K", "L", "N"],
};

function ProtegerColunasMultiplas() {
  if (!passwordOk_()) return;
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const res = { protegidas: [], faltantes: [] };

  for (const sheetName in SHEETS_CONFIG) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      res.faltantes.push(sheetName);
      continue;
    }
    const lastRow = Math.max(getLastRow_(sheet), MIN_ROWS);
    removeOldProtectionsInSheet_(sheet);
    SHEETS_CONFIG[sheetName].forEach((col) => {
      const prot = sheet
        .getRange(`${col}1:${col}${lastRow}`)
        .protect()
        .setDescription(PROTECTION_DESC);
      try {
        const editors = prot.getEditors();
        if (editors && editors.length) prot.removeEditors(editors);
        if (prot.canDomainEdit && prot.canDomainEdit())
          prot.setDomainEdit(false);
      } catch (e) {
        Logger.log("Erro editores: " + e);
      }
    });
    res.protegidas.push(sheetName);
  }

  let msg = "";
  if (res.protegidas.length)
    msg += "Abas protegidas: " + res.protegidas.join(", ") + "\n";
  if (res.faltantes.length)
    msg += "Abas não encontradas: " + res.faltantes.join(", ") + "\n";
  ui.alert("Proteção", msg || "Nenhuma aba configurada.", ui.ButtonSet.OK);
}

function DesprotegerPlanilhasMultiplas() {
  if (!passwordOk_()) return;
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const res = { desprotegidas: [], faltantes: [] };

  for (const sheetName in SHEETS_CONFIG) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      res.faltantes.push(sheetName);
      continue;
    }
    if (removeOldProtectionsInSheet_(sheet)) res.desprotegidas.push(sheetName);
  }

  let msg = "";
  if (res.desprotegidas.length)
    msg += "Abas desprotegidas: " + res.desprotegidas.join(", ") + "\n";
  if (res.faltantes.length)
    msg += "Abas não encontradas: " + res.faltantes.join(", ") + "\n";
  ui.alert(
    "Desproteção",
    msg || "Nenhuma proteção encontrada.",
    ui.ButtonSet.OK,
  );
}

/* PROTEÇÃO DE ABA INTEIRA — BLOQUEIA UNHIDE DE LINHAS OCULTAS
   =============================================================
   Protege a aba ENCARTE OPÇÃO como um todo. Apenas as linhas
   VISÍVEIS (não ocultas) são adicionadas como faixas liberadas,
   de modo que linhas ocultas ficam completamente blindadas:
   ninguém consegue reexibi-las, editá-las ou preenchê-las.
   A proteção é recalculada automaticamente ao mover / retornar.
   ==============================================================*/
const SHEET_PROT_DESC = "AbaEncarte - Linhas ocultas bloqueadas";

function aplicarProtecaoAbaEncarte_(sheet) {
  // 1. Remove qualquer proteção de aba já existente para o encarte
  removerProtecaoAbaEncarte_(sheet);

  // 2. Protege a aba inteira
  const prot = sheet.protect().setDescription(SHEET_PROT_DESC);

  // 3. Descobre quais linhas estão VISÍVEIS (podem ser editadas)
  const lastRow = Math.max(sheet.getLastRow(), MAIN_DATA_START + 10);
  const unprotectedRanges = [];

  for (let r = 1; r <= lastRow; r++) {
    try {
      // isRowHiddenByUser_ só retorna true se a linha foi oculta manualmente/por script
      if (!sheet.isRowHiddenByUser(r)) {
        unprotectedRanges.push(sheet.getRange(r, 1, 1, sheet.getMaxColumns()));
      }
    } catch (e) {
      // Se não conseguir verificar, libera por segurança
      unprotectedRanges.push(sheet.getRange(r, 1, 1, sheet.getMaxColumns()));
    }
  }

  // 4. Libera apenas as linhas visíveis para edição
  if (unprotectedRanges.length) {
    prot.setUnprotectedRanges(unprotectedRanges);
  }

  // 5. Garante que ninguém (exceto o dono do script) passe por cima
  try {
    const editors = prot.getEditors();
    if (editors && editors.length) prot.removeEditors(editors);
    if (prot.canDomainEdit && prot.canDomainEdit()) prot.setDomainEdit(false);
  } catch (e) {
    Logger.log("aplicarProtecaoAbaEncarte_: " + e);
  }
}

function removerProtecaoAbaEncarte_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach((p) => {
    if (p.getDescription() === SHEET_PROT_DESC) {
      try { p.remove(); } catch (e) { Logger.log(e); }
    }
  });
}

/* ATIVAR / DESATIVAR PROTEÇÃO MANUALMENTE (menu)
   ================================================ */
function AtivarBloqueioLinhasOcultas() {
  if (!passwordOk_()) return;
  const sheet = SpreadsheetApp.getActive().getSheetByName(ENCARTE_OPCAO_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Aba "' + ENCARTE_OPCAO_SHEET + '" não encontrada.');
    return;
  }
  aplicarProtecaoAbaEncarte_(sheet);
  SpreadsheetApp.getUi().alert(
    "🔒 Bloqueio Ativado",
    "As linhas ocultas do ENCARTE OPÇÃO estão protegidas.\n" +
    "Ninguém conseguirá reexibi-las ou editá-las.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function DesativarBloqueioLinhasOcultas() {
  if (!passwordOk_()) return;
  const sheet = SpreadsheetApp.getActive().getSheetByName(ENCARTE_OPCAO_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Aba "' + ENCARTE_OPCAO_SHEET + '" não encontrada.');
    return;
  }
  removerProtecaoAbaEncarte_(sheet);
  SpreadsheetApp.getUi().alert(
    "🔓 Bloqueio Desativado",
    "A proteção de linhas ocultas foi removida do ENCARTE OPÇÃO.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* BACKUP POR ABA
   ================= */
function backupAlertaOpcao() {
  fazerBackupPorAba_("ALERTA OPÇÃO");
}
function backupEncarteOpcao() {
  fazerBackupPorAba_("ENCARTE OPÇÃO");
}
function backupOfertaFds() {
  fazerBackupPorAba_("OFERTA FDS OPÇÃO");
}
function backupOfertaVirtual() {
  fazerBackupPorAba_("OFERTA VIRTUAL");
}

function fazerBackupPorAba_(nomeAba) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const aba = ss.getSheetByName(nomeAba);

  if (!aba) {
    ui.alert(
      "Backup",
      `A aba "${nomeAba}" não foi encontrada.`,
      ui.ButtonSet.OK,
    );
    return;
  }

  const dataHoje = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
  const nomeCopia = `${nomeAba} - ${dataHoje}`;

  if (ss.getSheetByName(nomeCopia)) {
    ui.alert(
      "Backup",
      `Já existe um backup de "${nomeAba}" para hoje.`,
      ui.ButtonSet.OK,
    );
    return;
  }

  const nova = aba.copyTo(ss);
  nova.setName(nomeCopia);
  ss.setActiveSheet(nova);
  ss.moveActiveSheet(ss.getSheets().length);
  nova.hideSheet();
  ui.alert(
    "Backup",
    `Backup de "${nomeAba}" criado com sucesso!`,
    ui.ButtonSet.OK,
  );
}

/* LIMPEZA POR TABELA
   ===================== */
const CLEAR_CONFIG = {
  "Tabela 1": {
    cols: ["D", "F", "G", "H", "I", "J", "M"],
    anchors: ["D", "F"],
  },
};

function limparTabela1() {
  limparPorZonaDinamica_("Tabela 1");
}

/* DEVOLVE TODOS OS PRODUTOS DAS SEÇÕES CAPA/APP AO CORPO PRINCIPAL
   ================================================================== */
function devolverTodosCapaApp_(sheet, ss) {
  let devolvidos = 0;

  for (const secKey in CAPA_APP_SECTIONS) {
    const sec = CAPA_APP_SECTIONS[secKey];
    const numRows = sec.lastRow - sec.firstRow + 1;

    // OTIMIZAÇÃO: Lê todas as âncoras da seção de uma vez só (Batch Read)
    let anchorValues = [];
    try {
      anchorValues = sheet.getRange(sec.firstRow, ANCHOR_COL_MOVE, numRows, 1).getDisplayValues();
    } catch (e) {
      Logger.log(`Erro ao ler âncoras da seção ${secKey}: ${e}`);
      continue;
    }

    // Percorre de baixo para cima
    for (let i = numRows - 1; i >= 0; i--) {
      const r = sec.firstRow + i;
      const anchorVal = String(anchorValues[i][0]).trim();

      if (anchorVal !== "") {
        try {
          // Devolve o produto para a origem sem exclui-lo ainda (skipDelete = true, skipProtectionReapply = true)
          returnFromSection_(sheet, sec, r, ss, true, true);
          devolvidos++;
        } catch (e) {
          Logger.log(`Erro ao devolver linha ${r} da seção ${secKey}: ${e}`);
        }
      }
    }
  }

  if (devolvidos > 0) {
    // Reaplica a proteção apenas uma vez ao finalizar todas as devoluções (Performance)
    aplicarProtecaoAbaEncarte_(sheet);

    ss.toast(
      `↩️ ${devolvidos} produto(s) devolvido(s) para suas linhas originais com sucesso!`,
      "Restaurando",
      4
    );
    // Pausa curta para o toast ser visível
    Utilities.sleep(1500);
  }

  return devolvidos;
}

function limparPorZonaDinamica_(zoneName) {
  const cache = CacheService.getScriptCache();
  cache.put("SCRIPT_RUNNING", "true", 60);

  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const ss = SpreadsheetApp.getActive();

    /* ── DEVOLVE PRODUTOS CAPA/APP ANTES DE LIMPAR ─────────────────────── */
    if (sheet.getName() === ENCARTE_OPCAO_SHEET) {
      const temProdutoCapaApp = Object.values(CAPA_APP_SECTIONS).some((sec) => {
        for (let r = sec.firstRow; r <= sec.lastRow; r++) {
          try {
            const val = String(
              sheet.getRange(r, ANCHOR_COL_MOVE).getDisplayValue()
            ).trim();
            if (val !== "") return true;
          } catch (_) { }
        }
        return false;
      });

      if (temProdutoCapaApp) {
        const aviso = ui.alert(
          "⚠️ Produtos em Capa/App detectados",
          "Existem produtos nas seções Capa e/ou App.\n\n" +
          "Eles serão DEVOLVIDOS às suas linhas originais antes da limpeza.\n\n" +
          "Deseja continuar?",
          ui.ButtonSet.YES_NO
        );
        if (aviso !== ui.Button.YES) return;

        devolverTodosCapaApp_(sheet, ss);
      }
    }
    /* ────────────────────────────────────────────────────────────────────── */
    const zone = getZoneByName_(zoneName);

    if (!zone) {
      ui.alert("Limpar", `Zona "${zoneName}" não encontrada.`, ui.ButtonSet.OK);
      return;
    }
    if (
      zone.sheets &&
      zone.sheets.length &&
      !zone.sheets.includes(sheet.getName())
    ) {
      ui.alert(
        "Limpar",
        `A aba "${sheet.getName()}" não pertence à zona "${zoneName}".`,
        ui.ButtonSet.OK,
      );
      return;
    }

    const bounds = getZoneBounds_(sheet, zone);
    if (!bounds) {
      ui.alert("Limpar", `Área inválida para "${zoneName}".`, ui.ButtonSet.OK);
      return;
    }

    const cfg = CLEAR_CONFIG[zoneName];
    if (!cfg) {
      ui.alert(
        "Limpar",
        `Configuração não encontrada para "${zoneName}".`,
        ui.ButtonSet.OK,
      );
      return;
    }

    const targetIdxs = cfg.cols.map(colToIndex_);
    const anchorIdxs = cfg.anchors.map(colToIndex_);
    const firstRowArea = bounds.startRow;
    const lastRowArea = bounds.endRow;
    const numRowsArea = lastRowArea - firstRowArea + 1;

    const anchorValues = anchorIdxs.map((idx) =>
      sheet.getRange(firstRowArea, idx, numRowsArea, 1).getDisplayValues(),
    );

    let firstDataRow = null;
    for (let rOff = 0; rOff < numRowsArea; rOff++) {
      const r = firstRowArea + rOff;
      for (let a = 0; a < anchorIdxs.length; a++) {
        if (isMergedCell_(sheet, r, anchorIdxs[a])) continue;
        if (String(anchorValues[a][rOff][0]).trim() !== "") {
          firstDataRow = r;
          break;
        }
      }
      if (firstDataRow !== null) break;
    }
    if (firstDataRow === null) firstDataRow = firstRowArea + FIRST_DATA_OFFSET;

    const targetValues = targetIdxs.map((idx) =>
      sheet.getRange(firstRowArea, idx, numRowsArea, 1).getDisplayValues(),
    );

    let lastDataRow = null;
    for (let rOff = numRowsArea - 1; rOff >= 0; rOff--) {
      const r = firstRowArea + rOff;
      for (let t = 0; t < targetIdxs.length; t++) {
        if (String(targetValues[t][rOff][0]).trim() !== "") {
          lastDataRow = r;
          break;
        }
      }
      if (lastDataRow !== null) break;
    }
    if (lastDataRow === null || lastDataRow < firstDataRow) {
      ui.alert(
        "Limpar",
        `Não encontrei dados em "${zoneName}" para limpar.`,
        ui.ButtonSet.OK,
      );
      return;
    }

    const allCols = Array.from(new Set([...targetIdxs, ...anchorIdxs])).sort(
      (a, b) => a - b,
    );
    const minCol = allCols[0];
    const width = allCols[allCols.length - 1] - minCol + 1;
    const mergedSet = buildMergedSet_(
      sheet,
      firstDataRow,
      lastDataRow,
      minCol,
      width,
    );

    const rangesA1 = [];
    for (const colIdx of targetIdxs) {
      let runStart = null;
      for (let r = firstDataRow; r <= lastDataRow; r++) {
        const merged = mergedSet.has(r + ":" + colIdx);
        if (!merged && runStart === null) runStart = r;
        if ((merged || r === lastDataRow) && runStart !== null) {
          const endR = merged ? r - 1 : r;
          rangesA1.push(
            `${indexToCol_(colIdx)}${runStart}:${indexToCol_(colIdx)}${endR}`,
          );
          runStart = null;
        }
      }
    }

    if (!rangesA1.length) {
      SpreadsheetApp.getActive().toast(
        `${zoneName}: nenhuma célula não-mesclada para limpar.`,
        "Info",
        4,
      );
      return;
    }

    const confirm = ui.alert(
      `Limpar ${zoneName}`,
      `Apagará colunas ${cfg.cols.join(", ")} (linhas ${firstDataRow}–${lastDataRow}) ` +
      `na aba "${sheet.getName()}".\nDeseja continuar?`,
      ui.ButtonSet.YES_NO,
    );
    if (confirm !== ui.Button.YES) return;

    try {
      sheet.getRangeList(rangesA1).clearContent();
      SpreadsheetApp.getActive().toast(
        `${zoneName}: ${rangesA1.length} bloco(s) limpo(s).`,
        "Pronto",
        5,
      );
    } catch (err) {
      ui.alert("Erro ao limpar", err.message, ui.ButtonSet.OK);
    }
  } finally {
    cache.remove("SCRIPT_RUNNING");
  }
}

/* MOVER LINHA PARA SEÇÃO CAPA / APP (Com proteção de Fórmulas e Super Rápido)
   ============================================================================= */
function moveToSection_(sheet, sec, srcRow, tipoKey, ss) {
  const ui = SpreadsheetApp.getUi();

  const numRows = sec.lastRow - sec.firstRow + 1;
  const anchorValues = sheet
    .getRange(sec.firstRow, ANCHOR_COL_MOVE, numRows, 1)
    .getDisplayValues();

  let targetRow = null;

  for (let i = 0; i < anchorValues.length; i++) {
    if (String(anchorValues[i][0]).trim() === "") {
      targetRow = sec.firstRow + i;
      break;
    }
  }

  if (targetRow === null) {
    ui.alert(
      `⚠️ Limite Máximo Atingido`,
      `Você já tem exatamente as ${sec.maxItems} linhas ocupadas na seção ${sec.label}!\n\nNão é possível adicionar mais. Remova um item antes de adicionar outro.`,
      ui.ButtonSet.OK,
    );
    sheet.getRange(srcRow, MKT_COL).clearContent();
    return;
  }

  // Exibe o slot de destino na seção Capa/App
  sheet.showRows(targetRow, 1);

  const minCol = Math.min(...DATA_COLS_MOVE);
  const maxCol = Math.max(...DATA_COLS_MOVE);
  const colsWidth = maxCol - minCol + 1;

  const srcRange = sheet.getRange(srcRow, minCol, 1, colsWidth);
  const dstRange = sheet.getRange(targetRow, minCol, 1, colsWidth);

  const srcValues = srcRange.getValues()[0];
  const srcFormulas = srcRange.getFormulas()[0];

  // Salva valores E fórmulas da linha inteira na nota do slot
  const fullRowRange = sheet.getRange(srcRow, 1, 1, sheet.getMaxColumns());
  const fullRowValues = fullRowRange.getValues()[0];
  const fullRowFormulas = fullRowRange.getFormulas()[0];

  const noteData = JSON.stringify({
    origRow: srcRow,
    values: fullRowValues,
    formulas: fullRowFormulas
  });
  sheet.getRange(targetRow, ANCHOR_COL_MOVE).setNote(noteData);

  // Copia os dados para o slot da seção Capa/App,
  // adaptando as referências de linha das fórmulas para a linha de destino.
  // Sempre sobrescreve a célula para garantir que as fórmulas fiquem corretas e atualizadas.
  DATA_COLS_MOVE.forEach((col) => {
    const arrayIndex = col - minCol;
    const cellDestino = sheet.getRange(targetRow, col);

    if (srcFormulas[arrayIndex]) {
      const formulaAdaptada = adaptFormulasToRow_(srcFormulas[arrayIndex], srcRow, targetRow);
      cellDestino.setFormula(formulaAdaptada);
    } else {
      cellDestino.setValue(srcValues[arrayIndex]);
    }
  });

  // Limpa as colunas de dados da linha original e a OCULTA (não deleta)
  // A linha permanece na planilha, vazia e oculta, pronta para receber os dados de volta
  DATA_COLS_MOVE.forEach((col) => {
    const cell = sheet.getRange(srcRow, col);
    if (!cell.isPartOfMerge()) cell.clearContent();
  });
  sheet.hideRows(srcRow, 1);

  // Reaplica a proteção de aba para excluir a linha agora oculta das faixas editáveis.
  // Isso impede que usuários reexibam (unhide) ou editem a linha enquanto ela estiver reservada.
  aplicarProtecaoAbaEncarte_(sheet);

  const slotNum = targetRow - sec.firstRow + 1;
  ss.toast(
    `✅ Movido para ${sec.label} — vaga ${slotNum}/${sec.maxItems}.`,
    "Sucesso",
    4,
  );
}

/* RETORNAR LINHA DA SEÇÃO PARA O CORPO PRINCIPAL
   =============================================================== */
function returnFromSection_(sheet, sec, srcRow, ss, skipDelete = false, skipProtectionReapply = false) {
  const cellAncoraOrigem = sheet.getRange(srcRow, ANCHOR_COL_MOVE);
  const nota = cellAncoraOrigem.getNote();
  let targetRow = null;
  let linhaSalva = null;
  let formulasSalvas = null;

  try {
    if (nota) {
      const dadosNota = JSON.parse(nota);
      if (dadosNota.origRow) {
        targetRow = dadosNota.origRow;
        linhaSalva = dadosNota.values;
        formulasSalvas = dadosNota.formulas;
      }
    }
  } catch (e) { }

  if (!targetRow) {
    SpreadsheetApp.getUi().alert(
      "⚠️ Erro ao retornar",
      "Não foi possível encontrar a linha de origem deste produto.\n" +
      "A nota de rastreamento está ausente ou corrompida.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // ⚠️ CRÍTICO: Limpa coluna J (MKT_COL) do SLOT antes de qualquer outra coisa.
  // Se não limpar antes, ao ocultar a linha o onEdit pode re-disparar sobre a célula J
  // ainda contendo o valor antigo, criando o loop de exclusão.
  sheet.getRange(srcRow, MKT_COL).clearContent();

  // A linha original foi apenas OCULTADA (não deletada) na ida.
  // Desoculta para que ela volte a ser exibida no corpo principal.
  sheet.showRows(targetRow, 1);

  // Restauração Otimizada (Batch Write): Restaura os dados e fórmulas originais salvos na linha de origem (targetRow)
  if (linhaSalva && formulasSalvas) {
    const colsCount = sheet.getMaxColumns();
    const rowValues = [];

    for (let c = 1; c <= colsCount; c++) {
      const idx = c - 1;
      const formula = formulasSalvas[idx];
      const value = linhaSalva[idx];

      // NÃO restaura a coluna J (MKT_COL, que é a coluna 10).
      // Ela deve permanecer vazia ao voltar para o corpo principal da planilha.
      if (c === MKT_COL) {
        rowValues.push("");
      } else {
        if (formula && String(formula).startsWith("=")) {
          rowValues.push(formula);
        } else {
          rowValues.push(value);
        }
      }
    }

    try {
      const targetRange = sheet.getRange(targetRow, 1, 1, colsCount);
      targetRange.setValues([rowValues]);
    } catch (e) {
      // Fallback célula por célula caso ocorra erro de mesclagem
      for (let c = 1; c <= colsCount; c++) {
        if (c === MKT_COL) continue;
        try {
          const cell = sheet.getRange(targetRow, c);
          if (!cell.isPartOfMerge()) {
            const idx = c - 1;
            const formula = formulasSalvas[idx];
            if (formula && String(formula).startsWith("=")) {
              cell.setFormula(formula);
            } else {
              cell.setValue(linhaSalva[idx]);
            }
          }
        } catch (_) { }
      }
    }
  }

  // ── SE NÃO FOR PULAR EXCLUSÃO, REALIZA O DELETE DOS DADOS ESPECÍFICOS NA ORIGEM ──────
  if (!skipDelete) {
    const colLetter = (c) => {
      let temp, letter = "";
      while (c > 0) {
        temp = (c - 1) % 26;
        letter = String.fromCharCode(65 + temp) + letter;
        c = (c - temp - 1) / 26;
      }
      return letter;
    };

    // Apaga apenas as colunas de dados D, F, G, H, I, J, M da linha de origem, mantendo outras (como E, K, L, N com fórmulas)
    // Mesmas colunas do CLEAR_CONFIG "Tabela 1" — exclui definitivamente mas mantém visível e limpa.
    const DELETE_COLS = [4, 6, 7, 8, 9, 10, 13]; // D=4, F=6, G=7, H=8, I=9, J=10, M=13
    const a1Ranges = DELETE_COLS.map(col => `${colLetter(col)}${targetRow}`);

    try {
      sheet.getRangeList(a1Ranges).clearContent();
    } catch (e) {
      // Fallback em caso de células mescladas
      DELETE_COLS.forEach((col) => {
        try {
          const cell = sheet.getRange(targetRow, col);
          if (!cell.isPartOfMerge()) cell.clearContent();
        } catch (_) { }
      });
    }
  }
  // ───────────────────────────────────────────────────────────────────────────

  // Limpa os dados do slot Capa/App de forma otimizada (Batch Write), remove a nota e oculta o slot
  const minCol = Math.min(...DATA_COLS_MOVE);
  const maxCol = Math.max(...DATA_COLS_MOVE);
  const colsWidth = maxCol - minCol + 1;
  sheet.getRange(srcRow, minCol, 1, colsWidth).clearContent();

  cellAncoraOrigem.clearNote();
  sheet.hideRows(srcRow, 1);

  // Reaplica a proteção de aba para refletir o novo estado se não for pulada (Performance)
  if (!skipProtectionReapply) {
    aplicarProtecaoAbaEncarte_(sheet);
  }

  if (!skipDelete) {
    ss.toast(
      `🗑️ Produto excluído da seção ${sec.label} (linha original ${targetRow} limpa).`,
      "Excluído",
      4
    );
  } else {
    ss.toast(
      `↩️ Produto devolvido para a origem ${targetRow} (dados restaurados).`,
      "Devolvido",
      4
    );
  }
}

/* EXCLUIR PRODUTO DEFINITIVAMENTE VIA LIXEIRA (BOTÃO)
   ==================================================== */
function excluirViaLixeira() {
  const cache = CacheService.getScriptCache();
  cache.put("SCRIPT_RUNNING", "true", 15);

  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== ENCARTE_OPCAO_SHEET) {
      ss.toast("⚠️ Esta função só pode ser usada na aba " + ENCARTE_OPCAO_SHEET + ".", "Aviso", 4);
      return;
    }

    const activeRow = sheet.getActiveCell().getRow();

    // 1. Verifica se a linha ativa pertence a uma das seções Capa/App
    const currentSection = Object.values(CAPA_APP_SECTIONS).find(
      (s) => activeRow >= s.firstRow && activeRow <= s.lastRow
    );

    if (currentSection) {
      const anchorVal = String(sheet.getRange(activeRow, ANCHOR_COL_MOVE).getDisplayValue()).trim();
      if (anchorVal === "") {
        ss.toast("⚠️ O slot selecionado já está vazio.", "Aviso", 4);
        return;
      }

      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert(
        "🗑️ Excluir Produto",
        `Deseja realmente EXCLUIR definitivamente o produto "${anchorVal}" da planilha?`,
        ui.ButtonSet.YES_NO
      );

      if (confirm === ui.Button.YES) {
        returnFromSection_(sheet, currentSection, activeRow, ss, false);
      }
      return;
    }

    // 2. Se for uma linha normal do corpo principal
    if (activeRow >= MAIN_DATA_START) {
      const anchorVal = String(sheet.getRange(activeRow, ANCHOR_COL_MOVE).getDisplayValue()).trim();
      if (anchorVal === "") {
        ss.toast("⚠️ A linha selecionada já está vazia.", "Aviso", 4);
        return;
      }

      const ui = SpreadsheetApp.getUi();
      const confirm = ui.alert(
        "🗑️ Excluir Produto",
        `Deseja realmente EXCLUIR definitivamente o produto "${anchorVal}" da planilha?`,
        ui.ButtonSet.YES_NO
      );

      if (confirm === ui.Button.YES) {
        const DELETE_COLS = [4, 6, 7, 8, 9, 10, 13]; // D, F, G, H, I, J, M
        DELETE_COLS.forEach((col) => {
          try {
            const cell = sheet.getRange(activeRow, col);
            if (!cell.isPartOfMerge()) cell.clearContent();
          } catch (e) { }
        });
        ss.toast(`🗑️ Produto "${anchorVal}" excluído com sucesso.`, "Excluído", 4);
      }
      return;
    }

    ss.toast("⚠️ Selecione uma linha de produto válida nas seções Capa, App ou no corpo principal para usar a lixeira.", "Aviso", 4);
  } finally {
    cache.remove("SCRIPT_RUNNING");
  }
}

/* VISIBILIDADE COLUNA H (PREÇO APP)
   ================================= */
function atualizarVisibilidadeColunaH_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const mktValues = sheet.getRange(2, MKT_COL, lastRow - 1, 1).getValues();
  const temApp = mktValues.some(
    (row) => String(row[0]).trim().toLowerCase() === "app",
  );

  if (temApp) {
    sheet.showColumns(PRECO_APP_COL);
  } else {
    sheet.hideColumns(PRECO_APP_COL);
  }
}

/* INICIALIZAR SLOTS (ocultar linhas vazias de Capa/App)
   ======================================================= */
function inicializarSlotsCapaApp() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(ENCARTE_OPCAO_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Aba "' + ENCARTE_OPCAO_SHEET + '" não encontrada.',
    );
    return;
  }

  let ocultadas = 0;
  let exibidas = 0;

  for (const secKey in CAPA_APP_SECTIONS) {
    const sec = CAPA_APP_SECTIONS[secKey];
    for (let r = sec.firstRow; r <= sec.lastRow; r++) {
      const anchorVal = String(
        sheet.getRange(r, ANCHOR_COL_MOVE).getDisplayValue(),
      ).trim();
      if (anchorVal === "") {
        sheet.hideRows(r, 1);
        ocultadas++;
      } else {
        sheet.showRows(r, 1);
        exibidas++;
      }
    }
  }

  ss.toast(
    `Slots vazios ocultos: ${ocultadas} | Slots com produto visíveis: ${exibidas}`,
    "Seções Capa/App",
    5,
  );
}

function buildMergedSet_(sheet, startRow, endRow, startCol, numCols) {
  const merged = new Set();
  try {
    const merges = sheet
      .getRange(startRow, startCol, endRow - startRow + 1, numCols)
      .getMergedRanges();
    for (const r of merges) {
      const r0 = r.getRow(),
        c0 = r.getColumn();
      for (let rr = r0; rr < r0 + r.getNumRows(); rr++)
        for (let cc = c0; cc < c0 + r.getNumColumns(); cc++)
          merged.add(rr + ":" + cc);
    }
  } catch (_) { }
  return merged;
}

/* MENUS E INICIALIZAÇÃO (onOpen)
   ============================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  COL_H_VISIBILITY_SHEETS.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) atualizarVisibilidadeColunaH_(sheet);
  });

  const menuProtecao = ui
    .createMenu("Proteção")
    .addItem("Proteger Colunas (senha)", "ProtegerColunasMultiplas")
    .addItem("Desproteger Abas (senha)", "DesprotegerPlanilhasMultiplas")
    .addSeparator()
    .addItem("🔒 Ativar Bloqueio Linhas Ocultas (senha)", "AtivarBloqueioLinhasOcultas")
    .addItem("🔓 Desativar Bloqueio Linhas Ocultas (senha)", "DesativarBloqueioLinhasOcultas");

  const menuUtils = ui
    .createMenu("Utilitários")
    .addItem("🗑️ Excluir Produto Selecionado (Lixeira)", "excluirViaLixeira")
    .addItem("Limpar Tabela 1 — rápido", "limparTabela1")
    .addItem("Ocultar slots vazios Capa/App", "inicializarSlotsCapaApp")
    .addSeparator()
    .addItem("🔧 Instalar TODOS os Gatilhos", "instalarTodosGatilhos");

  ui.createMenu("Automação")
    .addSubMenu(menuProtecao)
    .addSubMenu(menuUtils)
    .addToUi();
}

/* GATILHOS INSTALÁVEIS UNIFICADO
   ============================== */
function instalarTodosGatilhos() {
  const ssId = SpreadsheetApp.getActive().getId();
  let installedMsg = "";

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onEditInfo").forSpreadsheet(ssId).onEdit().create();
  installedMsg += "✅ onEditInfo (Edição)\n";

  SpreadsheetApp.getUi().alert(
    "Gatilhos Instalados!",
    "Os seguintes gatilhos estão ativos:\n\n" +
    installedMsg +
    "\nLembre-se: O backup automático foi removido conforme solicitado.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/* BACKUP AUTOMÁTICO
   ==================== */
const DATE_MAP = {
  "ALERTA OPÇÃO": { from: "G4", to: "H4" },
  "ENCARTE OPÇÃO": { from: "G4", to: "H4" },
  "OFERTA FDS OPÇÃO": { from: "G4", to: "H4" },
  "OFERTA VIRTUAL": { from: "G4", to: "H4" },
};

function backupAutomaticoPorData() {
  const ss = SpreadsheetApp.getActive();
  const hoje = new Date();
  for (const aba in DATE_MAP) {
    const sheet = ss.getSheetByName(aba);
    if (!sheet) continue;
    const cfg = DATE_MAP[aba];
    const dataDe = parseDateSafely_(sheet.getRange(cfg.from).getDisplayValue());
    const dataAte = parseDateSafely_(sheet.getRange(cfg.to).getDisplayValue());
    if (!dataAte || hoje < dataAte) continue;
    const nomeBackup = `${aba} - de ${formatDate_(dataDe)} até ${formatDate_(dataAte)}`;
    if (ss.getSheetByName(nomeBackup)) continue;
    const nova = sheet.copyTo(ss);
    nova.setName(nomeBackup);
    ss.setActiveSheet(nova);
    ss.moveActiveSheet(ss.getSheets().length);
    nova.hideSheet();
  }
}

function instalarGatilhoBackupAutomatico() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "backupAutomaticoPorData")
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("backupAutomaticoPorData")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  SpreadsheetApp.getUi().alert(
    "Backup automático instalado — roda todo dia às 08:00.",
  );
}

function removerGatilhoBackupAutomatico() {
  let removed = 0;
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (t.getHandlerFunction() === "backupAutomaticoPorData") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  SpreadsheetApp.getUi().alert(
    removed
      ? `${removed} gatilho(s) de backup automático removido(s).`
      : "Nenhum gatilho de backup automático encontrado.",
  );
}

/**
 * Adapta as referências de linha relativas de uma fórmula.
 * Ex: =SEERRO(PROCV(D35;'Banco'!$A:$B;2);"")  →  =SEERRO(PROCV(D7;'Banco'!$A:$B;2);"")
 * Apenas referências relativas de linha são substituídas (ex: D35, F35).
 * Referências absolutas de linha ($A$35, A$35) NÃO são alteradas.
 */
function adaptFormulasToRow_(formula, oldRow, newRow) {
  if (!formula) return formula;
  // Captura: (caractere anterior)(coluna)(número)
  // Se o caractere anterior for '$', é linha absoluta — não substitui.
  return formula.replace(
    /(^|[^$\w])([A-Za-z]+)(\d+)/g,
    (match, antes, col, num) => {
      if (Number(num) === Number(oldRow)) {
        return antes + col + newRow;
      }
      return match;
    }
  );
}

/* AUXILIARES
   ============= */
function normalizeCode_(v) {
  return v == null ? "" : String(v).trim();
}
function isFiniteNumber_(v) {
  return typeof v === "number" && isFinite(v);
}
function flatValues_(range) {
  return [].concat(...range.getDisplayValues());
}
function passwordOk_() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Proteção", "Digite a senha:", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return false;
  if (resp.getResponseText() !== PASSWORD) {
    ui.alert("Senha incorreta.");
    return false;
  }
  return true;
}
function getLastRow_(sheet) {
  return sheet.getLastRow() || 1;
}
function removeOldProtectionsInSheet_(sheet) {
  let removed = false;
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach((p) => {
    if (p.getDescription && p.getDescription() === PROTECTION_DESC) {
      try {
        p.remove();
        removed = true;
      } catch (e) {
        Logger.log(e);
      }
    }
  });
  return removed;
}
function detectZoneForRow_(sheet, row) {
  return (
    ZONES.find((z) => {
      if (z.sheets.length && !z.sheets.includes(sheet.getName())) return false;
      const b = getZoneBounds_(sheet, z);
      return b && row >= b.startRow && row <= b.endRow;
    }) || null
  );
}
function getZoneBounds_(sheet, zone) {
  const rng = safeGetRange_(sheet, zone.areaA1);
  return rng
    ? { startRow: rng.getRow(), endRow: rng.getRow() + rng.getNumRows() - 1 }
    : null;
}
function safeGetRange_(sheet, a1) {
  try {
    return sheet.getRange(a1);
  } catch (_) {
    return null;
  }
}
function isMergedCell_(sheet, row, colIdx) {
  if (!colIdx) return false;
  try {
    return sheet.getRange(row, colIdx).isPartOfMerge();
  } catch (_) {
    return false;
  }
}
function colToIndex_(letters) {
  if (!letters) return null;
  return letters
    .split("")
    .reduce((n, c) => n * 26 + c.toUpperCase().charCodeAt(0) - 64, 0);
}
function indexToCol_(index) {
  let s = "",
    n = index;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function getZoneByName_(name) {
  return ZONES.find((z) => z.name === name) || null;
}
function parseDateSafely_(v) {
  if (!v) return null;
  const p = v.split("/");
  if (p.length !== 2) return null;
  const d = new Date(new Date().getFullYear(), Number(p[1]) - 1, Number(p[0]));
  return isNaN(d.getTime()) ? null : d;
}
function formatDate_(date) {
  if (!date) return "";
  return (
    String(date.getDate()).padStart(2, "0") +
    "/" +
    String(date.getMonth() + 1).padStart(2, "0")
  );
}
function formatInfoMessage_(codigo, itens) {
  const header = [
    "🔎 Produto encontrado em outras abas",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    `Código: ${codigo}`,
    `Ocorrências em outras abas: ${itens.length}`,
    "────────────────────────────────────",
  ].join("\n");
  const body = itens
    .map(
      (it, i) =>
        `#${i + 1} • ${it.sheet} (linha ${it.row})\n` +
        `  ├─ Descrição: ${it.desc || "(sem descrição)"}\n` +
        `  └─ Preço: ${it.priceFmt}`,
    )
    .join("\n");
  return `${header}\n${body}\n\n💡 Este código já está em uso em outra(s) aba(s).`;
}