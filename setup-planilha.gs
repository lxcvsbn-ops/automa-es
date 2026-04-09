/**
 * FAZ Capital — Setup da Planilha CASOS_CRITICOS
 *
 * Como usar:
 *  1. Abra a planilha no Google Sheets
 *  2. Menu: Extensões > Apps Script
 *  3. Cole este código, salve (Ctrl+S)
 *  4. Clique em "Executar" (função setupPlanilha)
 *  5. Autorize as permissões solicitadas
 *
 * ATENÇÃO: Isso vai recriar os cabeçalhos e a formatação.
 * Os dados existentes NÃO serão apagados (apenas a linha 1 de cabeçalho).
 */

function setupPlanilha() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Página1');

  if (!sheet) {
    sheet = ss.insertSheet('Página1');
  }

  // ── Colunas ──────────────────────────────────────────────────────────────
  // "auto" = preenchido pelo workflow N8N (não editar manualmente)
  const COLUNAS = [
    // [nome,            largura, grupo]
    ['ID Caso',          100,    'manual'],
    ['Data Abertura',    110,    'manual'],
    ['Hora Abertura',     90,    'manual'],
    ['Código Cliente',   130,    'manual'],
    ['Assunto',          220,    'manual'],
    ['Descrição',        300,    'manual'],
    ['Protocolo N1',     120,    'manual'],
    ['Protocolo BV',     120,    'manual'],
    ['Responsável',      220,    'manual'],
    ['Status',           380,    'manual'],
    ['Ultimo Status Enviado', 380, 'auto'],
    ['Thread ID',        280,    'auto'],
    ['Protocolo BV Enviado', 150, 'auto'],
    ['Invite Enviado',   120,    'auto'],
    ['Escalado',         100,    'auto'],
  ];

  const NUM_COLS = COLUNAS.length;
  const NUM_ROWS = 1000; // linhas com formatação

  // ── Cabeçalhos ────────────────────────────────────────────────────────────
  const headers = COLUNAS.map(c => c[0]);
  sheet.getRange(1, 1, 1, NUM_COLS).setValues([headers]);

  // ── Larguras das colunas ──────────────────────────────────────────────────
  COLUNAS.forEach((col, i) => sheet.setColumnWidth(i + 1, col[1]));
  sheet.setRowHeight(1, 38);

  // ── Estilo do cabeçalho ───────────────────────────────────────────────────
  const COR_HEADER_MANUAL = '#1565C0';  // Azul escuro — colunas manuais
  const COR_HEADER_AUTO   = '#37474F';  // Cinza escuro — colunas automáticas

  COLUNAS.forEach((col, i) => {
    const cell = sheet.getRange(1, i + 1);
    cell.setBackground(col[2] === 'auto' ? COR_HEADER_AUTO : COR_HEADER_MANUAL);
    cell.setFontColor('#FFFFFF');
    cell.setFontWeight('bold');
    cell.setFontSize(10);
    cell.setFontFamily('Arial');
    cell.setHorizontalAlignment('center');
    cell.setVerticalAlignment('middle');
    cell.setWrap(false);

    if (col[2] === 'auto') {
      cell.setNote('⚙️ Preenchido automaticamente pelo workflow N8N.\nNão edite manualmente.');
    }
  });

  // ── Congelar linha de cabeçalho ───────────────────────────────────────────
  sheet.setFrozenRows(1);

  // ── Banding (linhas alternadas) ───────────────────────────────────────────
  // Remove bandings existentes antes de criar novo
  const existingBandings = sheet.getBandings();
  existingBandings.forEach(b => b.remove());

  const dataRange = sheet.getRange(1, 1, NUM_ROWS, NUM_COLS);
  const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor(COR_HEADER_MANUAL);
  banding.setFirstRowColor('#FFFFFF');
  banding.setSecondRowColor('#E8EAF6');
  banding.setFooterRowColor(null);

  // ── Validação de dados (dropdowns) ───────────────────────────────────────
  const simNaoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SIM', ''], true)
    .setAllowInvalid(false)
    .build();

  ['Protocolo BV Enviado', 'Invite Enviado', 'Escalado'].forEach(colName => {
    const colIdx = headers.indexOf(colName) + 1;
    if (colIdx > 0) {
      sheet.getRange(2, colIdx, NUM_ROWS - 1, 1).setDataValidation(simNaoRule);
    }
  });

  // ── Formatação condicional ────────────────────────────────────────────────
  const rules = [];

  // Protocolo BV Enviado = SIM → verde
  const colProtoBVEnv = headers.indexOf('Protocolo BV Enviado') + 1;
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SIM')
      .setBackground('#43A047')
      .setFontColor('#FFFFFF')
      .setRanges([sheet.getRange(2, colProtoBVEnv, NUM_ROWS - 1, 1)])
      .build()
  );

  // Invite Enviado = SIM → verde
  const colInviteEnv = headers.indexOf('Invite Enviado') + 1;
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SIM')
      .setBackground('#43A047')
      .setFontColor('#FFFFFF')
      .setRanges([sheet.getRange(2, colInviteEnv, NUM_ROWS - 1, 1)])
      .build()
  );

  // Escalado = SIM → vermelho
  const colEscalado = headers.indexOf('Escalado') + 1;
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('SIM')
      .setBackground('#E53935')
      .setFontColor('#FFFFFF')
      .setRanges([sheet.getRange(2, colEscalado, NUM_ROWS - 1, 1)])
      .build()
  );

  // Status preenchido e diferente de Ultimo Status Enviado → amarelo (pendente de envio)
  const colStatus     = headers.indexOf('Status') + 1;
  const colUltStatus  = headers.indexOf('Ultimo Status Enviado') + 1;
  const colThread     = headers.indexOf('Thread ID') + 1;

  // Thread ID preenchido → fundo cinza claro (caso ativo com e-mail aberto)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextIsNotEmpty()
      .setBackground('#E0F2F1')
      .setRanges([sheet.getRange(2, colThread, NUM_ROWS - 1, 1)])
      .build()
  );

  sheet.setConditionalFormatRules(rules);

  // ── Formatação das colunas de data e hora ─────────────────────────────────
  const colData = headers.indexOf('Data Abertura') + 1;
  const colHora = headers.indexOf('Hora Abertura') + 1;
  sheet.getRange(2, colData, NUM_ROWS - 1, 1).setNumberFormat('DD/MM/YYYY');
  sheet.getRange(2, colHora, NUM_ROWS - 1, 1).setNumberFormat('HH:MM');

  // ── Wrap nas colunas de texto longo ──────────────────────────────────────
  ['Status', 'Ultimo Status Enviado', 'Descrição'].forEach(colName => {
    const colIdx = headers.indexOf(colName) + 1;
    if (colIdx > 0) {
      sheet.getRange(2, colIdx, NUM_ROWS - 1, 1).setWrap(true);
    }
  });

  // ── Linha de exemplo (row 2) ──────────────────────────────────────────────
  const exemplo = [
    'CRIT-001',
    new Date(),
    '09:30',
    'CLI-12345',
    'Cobrança de tarifa indevida',
    'Cliente questiona cobrança da tarifa de custódia do mês de março.',
    '19829360',
    '',
    'lucas.boeira@fazcapital.com.br Lucas Boeira',
    '',
    '',
    '',
    '',
    '',
    '',
  ];
  sheet.getRange(2, 1, 1, NUM_COLS).setValues([exemplo]);
  sheet.getRange(2, 1, 1, NUM_COLS).setFontStyle('italic');
  sheet.getRange(2, 1, 1, NUM_COLS).setFontColor('#9E9E9E');

  // Nota na linha de exemplo
  sheet.getRange(2, 1).setNote('💡 Linha de exemplo — pode apagar ou substituir pelos dados reais.');

  // ── Legenda no rodapé ─────────────────────────────────────────────────────
  criarLegenda_(sheet, NUM_COLS);

  // ── Resultado ─────────────────────────────────────────────────────────────
  SpreadsheetApp.getUi().alert(
    '✅ Planilha configurada!\n\n' +
    'Colunas em AZUL → preencher manualmente\n' +
    'Colunas em CINZA → preenchidas pelo workflow N8N\n\n' +
    'A linha 2 é um exemplo — pode apagar.\n' +
    'Legenda adicionada em uma aba separada.'
  );
}

// ── Legenda em aba separada ───────────────────────────────────────────────────
function criarLegenda_(ss_or_sheet, numColsPrincipal) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let leg = ss.getSheetByName('Legenda');
  if (!leg) leg = ss.insertSheet('Legenda');
  leg.clear();
  leg.setTabColor('#FFC107');

  const linhas = [
    ['LEGENDA DA PLANILHA CASOS_CRITICOS', '', ''],
    ['', '', ''],
    ['COLUNA', 'PREENCHIMENTO', 'DESCRIÇÃO'],
    ['ID Caso',               'Manual',     'Identificador único do chamado (ex: CRIT-001)'],
    ['Data Abertura',         'Manual',     'Data em que o chamado foi criado'],
    ['Hora Abertura',         'Manual',     'Hora de abertura do chamado'],
    ['Código Cliente',        'Manual',     'Código ou nome do cliente'],
    ['Assunto',               'Manual',     'Título resumido do chamado'],
    ['Descrição',             'Manual',     'Descrição completa do problema'],
    ['Protocolo N1',          'Manual',     'Número do protocolo de nível 1'],
    ['Protocolo BV',          'Manual',     'Número do protocolo BV (quando recebido do banco)'],
    ['Responsável',           'Manual',     'E-mail e nome do responsável (ex: email@faz.com.br Nome Sobrenome)'],
    ['Status',                'Manual',     'Status atual / nota de atualização para o follow-up. Altere aqui para disparar novo e-mail.'],
    ['Ultimo Status Enviado', 'Automático', 'Último status processado pelo workflow. NÃO editar — controla o loop.'],
    ['Thread ID',             'Automático', 'ID da conversa de e-mail. Preenchido ao criar o chamado no Telegram.'],
    ['Protocolo BV Enviado',  'Automático', 'SIM quando o workflow enviou e-mail informando o Protocolo BV.'],
    ['Invite Enviado',        'Automático', 'SIM quando o convite de calendário D+2 foi criado.'],
    ['Escalado',              'Automático', 'SIM quando o caso foi escalado para BDO.'],
    ['', '', ''],
    ['COMO USAR', '', ''],
    ['1.', 'Preencha as colunas manuais ao abrir um chamado.', ''],
    ['2.', 'O workflow do Telegram preenche Thread ID automaticamente.', ''],
    ['3.', 'Quando houver atualização, altere a coluna "Status".', ''],
    ['4.', 'O workflow de follow-up detecta a mudança e cria o rascunho de e-mail.', ''],
    ['5.', 'Abra o rascunho no Outlook, anexe documentos se necessário, envie.', ''],
    ['6.', 'Quando receber o Protocolo BV, preencha a coluna "Protocolo BV".', ''],
    ['7.', 'O workflow envia automaticamente o e-mail com o número do protocolo.', ''],
    ['', '', ''],
    ['LEGENDA DE CORES', '', ''],
    ['Verde (SIM)',     'Invite Enviado / Protocolo BV Enviado', 'Ação já realizada pelo workflow'],
    ['Vermelho (SIM)', 'Escalado',                              'Caso escalado para BDO'],
    ['Azul claro',     'Thread ID',                            'Caso com thread de e-mail ativa'],
  ];

  leg.getRange(1, 1, linhas.length, 3).setValues(linhas);

  // Título
  const titulo = leg.getRange(1, 1, 1, 3);
  titulo.merge();
  titulo.setValue('LEGENDA DA PLANILHA CASOS_CRITICOS');
  titulo.setBackground('#1565C0');
  titulo.setFontColor('#FFFFFF');
  titulo.setFontWeight('bold');
  titulo.setFontSize(13);
  titulo.setHorizontalAlignment('center');
  titulo.setVerticalAlignment('middle');
  leg.setRowHeight(1, 40);

  // Cabeçalho da tabela (linha 3)
  const cabecalho = leg.getRange(3, 1, 1, 3);
  cabecalho.setBackground('#37474F');
  cabecalho.setFontColor('#FFFFFF');
  cabecalho.setFontWeight('bold');
  cabecalho.setHorizontalAlignment('center');

  // Seções em negrito
  [20, 29].forEach(row => {
    const sec = leg.getRange(row, 1, 1, 3);
    sec.setBackground('#E3F2FD');
    sec.setFontWeight('bold');
    leg.setRowHeight(row, 30);
  });

  // Destaque linhas manuais vs automáticas
  for (let r = 4; r <= 18; r++) {
    const tipo = leg.getRange(r, 2).getValue();
    if (tipo === 'Automático') {
      leg.getRange(r, 1, 1, 3).setBackground('#ECEFF1');
      leg.getRange(r, 2).setFontColor('#546E7A').setFontStyle('italic');
    }
  }

  // Exemplos de cores na legenda
  leg.getRange(31, 1).setBackground('#43A047').setFontColor('#FFFFFF');
  leg.getRange(32, 1).setBackground('#E53935').setFontColor('#FFFFFF');
  leg.getRange(33, 1).setBackground('#E0F2F1');

  // Larguras
  leg.setColumnWidth(1, 180);
  leg.setColumnWidth(2, 280);
  leg.setColumnWidth(3, 380);

  // Banding na tabela de colunas
  leg.getRange(3, 1, 16, 3)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    .setHeaderRowColor('#37474F')
    .setFirstRowColor('#FFFFFF')
    .setSecondRowColor('#E8EAF6');
}
