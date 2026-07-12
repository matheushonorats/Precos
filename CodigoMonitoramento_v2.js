/**
 * ╔══════════════════════════════════════════════════════════════╗
 * ║   SISTEMA DE MONITORAMENTO DE PREÇOS TURÍSTICOS — SETUR     ║
 * ║   Versão 2.0 — Edição Completa e Auto-Configurável          ║
 * ╚══════════════════════════════════════════════════════════════╝
 *
 * INSTALAÇÃO:
 *   1. No Google Sheets, vá em Extensões → Apps Script
 *   2. Apague todo o conteúdo existente
 *   3. Cole este script completo
 *   4. Salve (Ctrl+S)
 *   5. Clique em "Executar" a função: setupCompleto
 *   6. Autorize as permissões quando solicitado
 *   7. Insira sua chave da API Gemini na aba "Painel de Controle", célula B3
 *
 * Obtenha a chave gratuita em: https://aistudio.google.com
 */

// ============================================================
// CONFIGURAÇÕES GLOBAIS
// ============================================================
const CFG = {
  SHEET_MONITOR : 'Monitoramento',
  SHEET_PAINEL  : 'Painel de Controle',
  SHEET_HISTORICO: 'Histórico de Preços',
  SHEET_LOG     : 'Log de Erros',
  API_CELL      : 'B3',
  BATCH_SIZE_CELL: 'B17',   // Célula onde o usuário pode ajustar o lote
  DELAY_CELL    : 'B18',    // Célula onde o usuário pode ajustar o delay
  MODEL_CELL    : 'B19',    // Célula do modelo Gemini

  BATCH_SIZE    : 25,        // Itens por execução (padrão — evita timeout de 6min)
  DELAY_MS      : 4500,      // Delay entre requisições (ms)
  GEMINI_MODEL  : 'gemini-3.1-flash-lite',
  TEXT_MAX_CHARS: 15000,
  TEXT_MIN_CHARS: 50,
  MAX_RETRIES   : 3,

  // Nome da função usada pelo trigger de continuação
  TRIGGER_FN    : 'continuarProcessamento',

  // Domínios que bloqueiam scraping ou exigem JavaScript para renderizar
  PLATAFORMAS_JS: [
    'instagram.com', 'facebook.com', 'fb.com',
    'ifood.com.br',  'goomer.app',   'ola.click',
    'menudino.com',  'negocio.site', 'business.site',
    'blogspot.com',  'blogspot.com.br', 'wordpress.com',
    'placeweb.site', 'webnode.com',  'wixsite.com',
    'reservecom',    'stays.net',    'hubt.com.br'
  ],

  COLS: {
    ID: 1, NOME: 2, ATIVIDADE: 3, SITE: 4,
    PRECO_ANT: 5, PRECO_ATUAL: 6, VARIACAO: 7,
    STATUS: 8, DETALHES: 9, ATUALIZACAO: 10
  }
};

// ============================================================
// MENU — Executado automaticamente ao abrir a planilha
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔍 Monitor de Preços')
    .addItem('🚀 Configuração Inicial (Executar 1x)', 'setupCompleto')
    .addSeparator()
    .addItem('⚡ Atualizar Todos os Preços', 'atualizarTodosPrecos')
    .addItem('🎯 Atualizar Apenas Linhas Selecionadas', 'atualizarPrecosSelecionados')
    .addItem('▶️ Continuar Processamento (Retomar)', 'continuarProcessamento')
    .addSeparator()
    .addItem('📅 Ativar Atualização Diária Automática', 'configurarGatilhoDiario')
    .addItem('❌ Desativar Atualização Automática', 'removerGatilhosDiarios')
    .addSeparator()
    .addItem('🔄 Resetar Status para "Aguardando"', 'resetarStatus')
    .addItem('📊 Ver Resumo de Resultados', 'mostrarResumo')
    .addToUi();
}


// ============================================================
// SETUP COMPLETO — Cria e formata toda a planilha
// ============================================================
function setupCompleto() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    _criarAbas(ss);
    _setupPainelDeControle(ss);
    _setupAbaMonitoramento(ss);
    _setupAbaHistorico(ss);
    _setupAbaLog(ss);
    ss.setActiveSheet(ss.getSheetByName(CFG.SHEET_PAINEL));

    ui.alert(
      '✅ Setup Concluído com Sucesso!',
      'Planilha configurada e dados carregados!\n\n' +
      'PRÓXIMOS PASSOS:\n' +
      '1. Insira sua chave Gemini na célula B3 desta aba\n' +
      '2. Use o menu "Monitor de Preços" → "Atualizar Todos os Preços"\n\n' +
      'IMPORTANTE: Para listas com mais de 25 itens, o sistema processa\n' +
      'em lotes automáticos. Se parar, clique em "Continuar Processamento".\n\n' +
      'Para automação total, use "Ativar Atualização Diária Automática".',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('❌ Erro no Setup', 'Detalhe: ' + e.message, ui.ButtonSet.OK);
    _logErro('setupCompleto', 'N/A', e.toString());
  }
}

function _criarAbas(ss) {
  [CFG.SHEET_PAINEL, CFG.SHEET_MONITOR, CFG.SHEET_HISTORICO, CFG.SHEET_LOG].forEach((nome, idx) => {
    if (!ss.getSheetByName(nome)) ss.insertSheet(nome);
    try {
      ss.setActiveSheet(ss.getSheetByName(nome));
      ss.moveActiveSheet(idx + 1);
    } catch (e) {}
  });
}

function _setupPainelDeControle(ss) {
  const s = ss.getSheetByName(CFG.SHEET_PAINEL);
  s.clear();
  s.setColumnWidth(1, 260);
  s.setColumnWidth(2, 420);

  // ── Título ───────────────────────────────────────
  s.getRange('A1:B1').merge()
    .setValue('⚙️  PAINEL DE CONTROLE — MONITOR DE PREÇOS SETUR')
    .setBackground('#0f3460').setFontColor('#e0e0e0')
    .setFontWeight('bold').setFontSize(13)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(1, 38);

  // ── Seção API ────────────────────────────────────
  s.getRange('A2:B2').merge()
    .setValue('🔑  CONFIGURAÇÃO DA API')
    .setBackground('#16213e').setFontColor('#e0e0e0').setFontWeight('bold');

  s.getRange('A3').setValue('Chave da API Gemini (aistudio.google.com)').setFontStyle('italic');
  s.getRange('B3')
    .setValue('INSIRA_SUA_CHAVE_AQUI')
    .setBackground('#fff3cd').setFontWeight('bold').setFontColor('#856404')
    .setBorder(true, true, true, true, false, false, '#ffc107', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  s.getRange('A4').setValue('Obter chave gratuita:').setFontColor('#555');
  s.getRange('B4').setValue('→ aistudio.google.com  →  "Get API Key"').setFontColor('#0066cc');

  // ── Seção Status ─────────────────────────────────
  s.getRange('A6:B6').merge()
    .setValue('📊  STATUS DO MONITORAMENTO')
    .setBackground('#16213e').setFontColor('#e0e0e0').setFontWeight('bold');

  const statusRows = [
    ['Total de Estabelecimentos:', `=COUNTA(${CFG.SHEET_MONITOR}!A2:A)`, '#000'],
    ['✅ Atualizados com Sucesso:', `=COUNTIF(${CFG.SHEET_MONITOR}!H:H,"Sucesso*")`, '#155724'],
    ['🔍 Preço Não Encontrado:', `=COUNTIF(${CFG.SHEET_MONITOR}!H:H,"Preço não encontrado")`, '#856404'],
    ['❌ Erros de Acesso:', `=COUNTIFS(${CFG.SHEET_MONITOR}!H:H,"Erro*")`, '#721c24'],
    ['🚫 Plataforma JS (não suportado):', `=COUNTIF(${CFG.SHEET_MONITOR}!H:H,"Plataforma JS*")`, '#383d41'],
    ['⚠️ URL Inválida:', `=COUNTIF(${CFG.SHEET_MONITOR}!H:H,"URL Inválida*")`, '#856404'],
    ['⏳ Aguardando / Processando:', `=COUNTIF(${CFG.SHEET_MONITOR}!H:H,"Aguardando")+COUNTIF(${CFG.SHEET_MONITOR}!H:H,"Processando...")`, '#0c5460'],
    ['📈 Taxa de Cobertura:', `=IFERROR(TEXT(COUNTIF(${CFG.SHEET_MONITOR}!H:H,"Sucesso*")/COUNTA(${CFG.SHEET_MONITOR}!A2:A),"0.0%"),"0.0%")`, '#000'],
  ];

  statusRows.forEach((row, i) => {
    s.getRange(7 + i, 1).setValue(row[0]);
    if (row[1].startsWith('=')) {
      s.getRange(7 + i, 2).setFormula(row[1]).setFontColor(row[2]).setFontWeight('bold');
    } else {
      s.getRange(7 + i, 2).setValue(row[1]).setFontColor(row[2]).setFontWeight('bold');
    }
  });

  // ── Seção Config Avançada ────────────────────────
  s.getRange('A16:B16').merge()
    .setValue('⚙️  CONFIGURAÇÕES AVANÇADAS')
    .setBackground('#16213e').setFontColor('#e0e0e0').setFontWeight('bold');

  const configs = [
    ['Itens por execução (lote):', CFG.BATCH_SIZE, 'Recomendado: 20-30. Lotes maiores causam timeout (limite 6 min)'],
    ['Delay entre requisições (ms):', CFG.DELAY_MS,  'Mínimo: 1500. Aumente se houver erros 429 (rate limit)'],
    ['Modelo Gemini:', CFG.GEMINI_MODEL, 'gemini-3.1-flash-lite (alto limite de cota/rápido) | gemini-3.1-pro (mais preciso)'],
  ];

  configs.forEach((c, i) => {
    s.getRange(17 + i, 1).setValue(c[0]);
    s.getRange(17 + i, 2).setValue(c[1])
      .setBackground('#f8f9fa')
      .setNote(c[2]);
  });

  // ── Seção Última Execução ────────────────────────
  s.getRange('A21:B21').merge()
    .setValue('🕐  ÚLTIMA EXECUÇÃO')
    .setBackground('#16213e').setFontColor('#e0e0e0').setFontWeight('bold');

  [
    ['Início da execução:', '-'],
    ['Último lote processado:', '-'],
    ['Situação atual:', 'Pronto'],
  ].forEach((row, i) => {
    s.getRange(22 + i, 1).setValue(row[0]);
    s.getRange(22 + i, 2).setValue(row[1]);
  });

  s.setFrozenRows(1);
}

function _setupAbaMonitoramento(ss) {
  const s = ss.getSheetByName(CFG.SHEET_MONITOR);
  s.clear();

  const headers = [
    'ID', 'Nome Fantasia', 'Atividade', 'Site',
    'Preço Anterior (R$)', 'Preço Atual (R$)', 'Variação %',
    'Status', 'Detalhes / Observação', 'Última Atualização'
  ];

  s.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#0f3460').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setWrap(false);

  // Larguras otimizadas
  [80, 260, 190, 240, 135, 135, 100, 170, 300, 160].forEach((w, i) => s.setColumnWidth(i + 1, w));

  s.setFrozenRows(1);
  s.setFrozenColumns(2);

  // Popular dados do CSV
  const dados = _getDadosCSV();
  if (dados.length > 0) {
    s.getRange(2, 1, dados.length, 10).setValues(dados);
  }

  // Formatos de número
  const lr = dados.length + 1;
  s.getRange(2, CFG.COLS.PRECO_ANT,  lr, 1).setNumberFormat('R$ #,##0.00');
  s.getRange(2, CFG.COLS.PRECO_ATUAL, lr, 1).setNumberFormat('R$ #,##0.00');
  s.getRange(2, CFG.COLS.VARIACAO,    lr, 1).setNumberFormat('0.00%');
  s.getRange(2, CFG.COLS.ATUALIZACAO, lr, 1).setNumberFormat('dd/mm/yyyy HH:mm');

  // Formatação condicional de status
  const sr = s.getRange(2, CFG.COLS.STATUS, lr, 1);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Sucesso')
      .setBackground('#d4edda').setFontColor('#155724').setRanges([sr]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Erro')
      .setBackground('#f8d7da').setFontColor('#721c24').setRanges([sr]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Preço não encontrado')
      .setBackground('#fff3cd').setFontColor('#856404').setRanges([sr]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('Plataforma JS')
      .setBackground('#e2e3e5').setFontColor('#383d41').setRanges([sr]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('URL Inválida')
      .setBackground('#ffeeba').setFontColor('#856404').setRanges([sr]).build(),
  ]);
}

function _setupAbaLog(ss) {
  const s = ss.getSheetByName(CFG.SHEET_LOG);
  s.clear();
  s.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Função', 'Estabelecimento / ID', 'Detalhe do Erro']])
    .setBackground('#721c24').setFontColor('#ffffff').setFontWeight('bold');
  [160, 150, 260, 500].forEach((w, i) => s.setColumnWidth(i + 1, w));
  s.setFrozenRows(1);
}

function _setupAbaHistorico(ss) {
  const s = ss.getSheetByName(CFG.SHEET_HISTORICO);
  if (s.getLastRow() === 0) { // Só configura se estiver vazia para não sobrescrever histórico existente
    const headers = ['Timestamp', 'ID', 'Nome Fantasia', 'Atividade', 'Preço (R$)'];
    s.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#0f3460').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
    [160, 80, 260, 190, 120].forEach((w, i) => s.setColumnWidth(i + 1, w));
    s.setFrozenRows(1);
    s.getRange('A:A').setNumberFormat('dd/mm/yyyy HH:mm');
    s.getRange('E:E').setNumberFormat('R$ #,##0.00');
  }
}


// ============================================================
// PROCESSAMENTO PRINCIPAL
// ============================================================
function atualizarTodosPrecos() {
  _iniciarProcessamento(false);
}

function atualizarPrecosSelecionados() {
  _iniciarProcessamento(true);
}

/**
 * Inicia ou reinicia o processamento a partir do índice 0 (ou seleção)
 */
function _iniciarProcessamento(apenasSelecionados) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = _lerConfigs(ss);
  if (!cfg) return;

  const { apiKey, sheet } = cfg;
  const { rowsToProcess, aviso } = _coletarLinhas(sheet, apenasSelecionados);
  if (aviso) {
    _mostrarAlerta('Aviso', aviso);
    return;
  }

  if (!apenasSelecionados && rowsToProcess.length > 50 && _podeMostrarUI()) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      '⚠️ Confirmação',
      `Você vai atualizar ${rowsToProcess.length} estabelecimentos.\n\n` +
      `O sistema processará em lotes de ${_getBatchSize(ss)} itens por execução em segundo plano de forma contínua.\n\nDeseja continuar?`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  // Salvar fila no estado do script
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'fila'       : JSON.stringify(rowsToProcess),
    'indice'     : '0',
    'apiKey'     : apiKey,
    'iniciado'   : new Date().toISOString(),
  });

  _atualizarPainelExecucao(ss, new Date(), 0, rowsToProcess.length, 'Processando...');
  _executarLote(ss, sheet, apiKey, rowsToProcess, 0);
}

/**
 * Continua a partir do ponto onde parou (chamado pelo trigger ou pelo usuário)
 */
function continuarProcessamento() {
  // O primeiro passo do gatilho é se auto-remover para não acumular
  _removerTriggerContinuacao();

  const props = PropertiesService.getScriptProperties();
  const filaJson = props.getProperty('fila');

  if (!filaJson) {
    _mostrarAlerta('Aviso', 'Nenhum processamento em andamento para continuar.\nInicie com "Atualizar Todos os Preços".');
    return;
  }

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(CFG.SHEET_MONITOR);
  const apiKey  = props.getProperty('apiKey');
  const indice  = parseInt(props.getProperty('indice') || '0', 10);
  const fila    = JSON.parse(filaJson);

  if (indice >= fila.length) {
    _finalizarProcessamento(ss, fila.length);
    return;
  }

  _executarLote(ss, sheet, apiKey, fila, indice);
}

/**
 * Executa um lote de itens a partir do índice dado
 */
function _executarLote(ss, sheet, apiKey, fila, indiceInicial) {
  const batchSize = _getBatchSize(ss);
  const delayMs   = _getDelay(ss);
  const modelo    = _getModelo(ss);

  const indiceFlnal = Math.min(indiceInicial + batchSize, fila.length);
  const lote = fila.slice(indiceInicial, indiceFlnal);

  // Marcar "Processando..." em lote
  const updates = lote.map(item => [item.rowNum, CFG.COLS.STATUS, 'Processando...']);
  _batchSetStatus(sheet, updates);
  SpreadsheetApp.flush();

  const resultados = [];

  for (let i = 0; i < lote.length; i++) {
    const item = lote[i];
    const resultado = _processarItem(apiKey, modelo, item);
    resultados.push({ rowNum: item.rowNum, ...resultado });
    Utilities.sleep(delayMs);
  }

  // Escrever resultados em batch
  _batchEscreverResultados(sheet, resultados);
  SpreadsheetApp.flush();

  const novoIndice = indiceFlnal;
  PropertiesService.getScriptProperties().setProperty('indice', String(novoIndice));
  _atualizarPainelExecucao(ss, null, novoIndice, fila.length, novoIndice < fila.length ? 'Processando...' : 'Concluído');

  if (novoIndice >= fila.length) {
    _finalizarProcessamento(ss, fila.length);
  } else {
    // Agendar próximo lote automaticamente via gatilho
    _removerTriggerContinuacao();
    
    try {
      ScriptApp.newTrigger(CFG.TRIGGER_FN)
        .timeBased()
        .after(5000) // Executa após 5 segundos
        .create();
      
      if (_podeMostrarUI()) {
        ss.toast(
          `Lote concluído. Processados ${novoIndice} de ${fila.length} estabelecimentos. Continuando em segundo plano...`,
          '🔍 Monitor de Preços',
          6
        );
      } else {
        console.log(`[Auto-Trigger] Lote concluído (${novoIndice}/${fila.length}). Próximo lote agendado.`);
      }
    } catch (e) {
      _logErro('criar_gatilho_continuacao', 'N/A', e.toString());
      _mostrarAlerta('Erro de Agendamento', 'Falha ao agendamento automático do próximo lote: ' + e.message);
    }
  }
}

/**
 * Processa um único estabelecimento: valida URL → fetch → Gemini
 */
function _processarItem(apiKey, modelo, item) {
  const url = item.site;
  let fallbackNecessario = false;
  let motivoOriginal = "";

  // 1. Validar URL
  const validacao = _validarURL(url);
  if (!validacao.ok) {
    fallbackNecessario = true;
    motivoOriginal = 'URL Inválida: ' + validacao.motivo;
  }

  let html = null;
  let fetchStatus = null;

  // 2. Verificar plataforma bloqueada (só se URL for válida)
  if (!fallbackNecessario) {
    const plataforma = _detectarPlataformaJS(url);
    if (plataforma) {
      fallbackNecessario = true;
      motivoOriginal = `Plataforma JS — Sem acesso estático (${plataforma})`;
    }
  }

  // 3. Fazer fetch do site com retry (só se não marcou fallback)
  if (!fallbackNecessario) {
    for (let tentativa = 1; tentativa <= CFG.MAX_RETRIES; tentativa++) {
      try {
        const resp = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true,
          followRedirects: true,
          validateHttpsCertificates: false,
          headers: {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
            'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
          }
        });
        fetchStatus = resp.getResponseCode();
        if (fetchStatus === 200) {
          html = resp.getContentText('UTF-8');
          break;
        }
        if (fetchStatus === 429) {
          if (tentativa === CFG.MAX_RETRIES) {
            fallbackNecessario = true;
            motivoOriginal = 'Erro: HTTP 429 (Muitas requisições)';
          } else {
            _esperarComBackoff(tentativa, true);
          }
        } else {
          if (tentativa === CFG.MAX_RETRIES) {
            fallbackNecessario = true;
            motivoOriginal = `Erro: HTTP ${fetchStatus}`;
          } else {
            _esperarComBackoff(tentativa, false);
          }
        }
      } catch (e) {
        if (tentativa === CFG.MAX_RETRIES) {
          _logErro('fetch', item.id + ' | ' + item.nome, e.toString());
          fallbackNecessario = true;
          motivoOriginal = 'Erro: Falha de conexão (' + e.message + ')';
        } else {
          const is429 = e.toString().includes('429');
          _esperarComBackoff(tentativa, is429);
        }
      }
    }

    if (!fallbackNecessario && !html) {
      fallbackNecessario = true;
      motivoOriginal = `Erro: HTTP ${fetchStatus}`;
    }
  }

  // 4. Limpar HTML e chamar Gemini API (só se fetch deu certo)
  let geminiResult = null;
  if (!fallbackNecessario && html) {
    const textoLimpo = _extrairTextoLimpo(html);
    if (textoLimpo.length < CFG.TEXT_MIN_CHARS) {
      fallbackNecessario = true;
      motivoOriginal = 'Erro: Site vazio ou bloqueou acesso';
    } else {
      for (let tentativa = 1; tentativa <= CFG.MAX_RETRIES; tentativa++) {
        try {
          geminiResult = _chamarGeminiAPI(apiKey, modelo, item.nome, item.atividade, textoLimpo.substring(0, CFG.TEXT_MAX_CHARS));
          break;
        } catch (e) {
          if (tentativa === CFG.MAX_RETRIES) {
            _logErro('gemini', item.id + ' | ' + item.nome, e.toString());
            fallbackNecessario = true;
            motivoOriginal = 'Erro: API Gemini falhou';
          } else {
            const is429 = e.toString().includes('429');
            _esperarComBackoff(tentativa, is429);
          }
        }
      }
    }
  }

  // Se funcionou até aqui e achou o preço, retorna sucesso!
  if (!fallbackNecessario && geminiResult && geminiResult.price > 0) {
    return {
      precoAnterior: item.precoAtual,
      novoPreco: geminiResult.price,
      status: 'Sucesso',
      detalhes: geminiResult.details || '',
      timestamp: new Date()
    };
  }

  // Caso contrário, tenta o fallback com Google Search Grounding!
  motivoOriginal = motivoOriginal || (geminiResult ? (geminiResult.details || 'Preço não encontrado no site oficial.') : 'Preço não encontrado.');
  
  for (let tentativa = 1; tentativa <= CFG.MAX_RETRIES; tentativa++) {
    try {
      const searchResult = _chamarGeminiAPIGoogleSearch(apiKey, modelo, item.nome, item.atividade, url);
      if (searchResult && searchResult.price > 0) {
        return {
          precoAnterior: item.precoAtual,
          novoPreco: searchResult.price,
          status: 'Sucesso (via Busca Google)',
          detalhes: (searchResult.details || '') + ` (Originalmente: ${motivoOriginal})`,
          timestamp: new Date()
        };
      }
      break; // Se rodou mas retornou price <= 0, não adianta tentar de novo
    } catch (err) {
      if (tentativa === CFG.MAX_RETRIES) {
        _logErro('google_search_fallback', item.id + ' | ' + item.nome, err.toString());
      } else {
        const is429 = err.toString().includes('429');
        _esperarComBackoff(tentativa, is429);
      }
    }
  }

  // Se o fallback falhar ou também não achar preço, retorna o erro/início original
  return {
    precoAnterior: item.precoAtual,
    novoPreco: null,
    status: 'Preço não encontrado',
    detalhes: motivoOriginal,
    timestamp: new Date()
  };
}


// ============================================================
// ESCRITA EM BATCH — Evita milhares de chamadas individuais
// ============================================================
function _batchSetStatus(sheet, updates) {
  updates.forEach(([rowNum, col, valor]) => {
    sheet.getRange(rowNum, col).setValue(valor);
  });
}

function _batchEscreverResultados(sheet, resultados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const histSheet = ss.getSheetByName(CFG.SHEET_HISTORICO);
  const rowsToAppend = [];

  resultados.forEach(r => {
    // Preço anterior
    sheet.getRange(r.rowNum, CFG.COLS.PRECO_ANT).setValue(r.precoAnterior || '');
    // Preço atual
    if (r.novoPreco !== null) {
      sheet.getRange(r.rowNum, CFG.COLS.PRECO_ATUAL).setValue(r.novoPreco);
      // Variação
      if (r.precoAnterior && r.precoAnterior > 0) {
        sheet.getRange(r.rowNum, CFG.COLS.VARIACAO).setValue((r.novoPreco - r.precoAnterior) / r.precoAnterior);
      } else {
        sheet.getRange(r.rowNum, CFG.COLS.VARIACAO).setValue('');
      }

      // Se for sucesso, prepara linha para histórico
      if (r.status.startsWith('Sucesso')) {
        const rowValues = sheet.getRange(r.rowNum, 1, 1, 3).getValues()[0]; // ID, Nome, Atividade
        rowsToAppend.push([r.timestamp, rowValues[0], rowValues[1], rowValues[2], r.novoPreco]);
      }
    } else {
      sheet.getRange(r.rowNum, CFG.COLS.VARIACAO).setValue('');
    }
    sheet.getRange(r.rowNum, CFG.COLS.STATUS).setValue(r.status);
    sheet.getRange(r.rowNum, CFG.COLS.DETALHES).setValue(r.detalhes);
    sheet.getRange(r.rowNum, CFG.COLS.ATUALIZACAO).setValue(r.timestamp);
  });

  // Salvar no histórico de uma só vez
  if (histSheet && rowsToAppend.length > 0) {
    try {
      histSheet.getRange(histSheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
    } catch (e) {
      _logErro('salvar_historico', 'N/A', e.toString());
    }
  }
}

function _finalizarProcessamento(ss, total) {
  _removerTriggerContinuacao();
  PropertiesService.getScriptProperties().deleteProperty('fila');
  PropertiesService.getScriptProperties().deleteProperty('indice');
  PropertiesService.getScriptProperties().deleteProperty('apiKey');
  PropertiesService.getScriptProperties().deleteProperty('iniciado');
  _atualizarPainelExecucao(ss, null, total, total, 'Concluído ✅');
  _mostrarAlerta(
    '🎉 Processamento Concluído!',
    `Todos os ${total} estabelecimentos foram processados.\n\n` +
    `Veja o resumo no "Painel de Controle" ou use "Ver Resumo de Resultados".`
  );
}


// ============================================================
// CHAMADA DA API GEMINI
// ============================================================
function _chamarGeminiAPI(apiKey, modelo, nome, atividade, textoSite) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent?key=${apiKey}`;

  const orientacao = _orientacaoPreco(atividade);

  const prompt =
    `Você é um monitor de preços automatizado especializado em turismo brasileiro.\n` +
    `Estabelecimento: "${nome}"\nCategoria: ${atividade}\n\n` +
    `INSTRUÇÃO ESPECÍFICA:\n${orientacao}\n\n` +
    `REGRAS OBRIGATÓRIAS:\n` +
    `- Use APENAS preços em Reais (BRL) explicitamente mencionados no texto.\n` +
    `- Se não houver preço claro e numérico, retorne price: -1.\n` +
    `- Não invente, estime ou infira preços não listados.\n` +
    `- Se houver faixa de preços, retorne o menor valor.\n\n` +
    `TEXTO EXTRAÍDO DO SITE:\n"""\n${textoSite}\n"""\n\n` +
    `Responda SOMENTE com o JSON conforme o schema abaixo.`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          price:    { type: 'NUMBER'  },
          currency: { type: 'STRING'  },
          details:  { type: 'STRING'  },
          status:   { type: 'STRING'  }
        },
        required: ['price', 'currency', 'details', 'status']
      }
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error(`Gemini HTTP ${resp.getResponseCode()}: ${resp.getContentText().substring(0, 200)}`);
  }

  const json = JSON.parse(resp.getContentText());
  if (json.candidates?.[0]?.content?.parts?.[0]) {
    return JSON.parse(json.candidates[0].content.parts[0].text);
  }

  throw new Error('Resposta Gemini em formato inesperado.');
}

function _chamarGeminiAPIGoogleSearch(apiKey, modelo, nome, atividade, site) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelo}:generateContent?key=${apiKey}`;

  const orientacao = _orientacaoPreco(atividade);

  const prompt =
    `Você é um monitor de preços automatizado especializado em turismo brasileiro.\n` +
    `Sua tarefa é pesquisar na internet (usando o Google Search) o preço atual do serviço oferecido por este estabelecimento:\n\n` +
    `Estabelecimento: "${nome}"\n` +
    `Categoria: ${atividade}\n` +
    `Site oficial (caso esteja funcionando): "${site}"\n\n` +
    `INSTRUÇÃO ESPECÍFICA DE BUSCA:\n${orientacao}\n\n` +
    `INSTRUÇÕES DE PESQUISA:\n` +
    `1. Use o Google Search para procurar tarifas, diárias ou preços para o estabelecimento "${nome}" em São Sebastião, SP.\n` +
    `2. Se o site oficial estiver fora do ar ou não contiver preços, pesquise em sites de reservas/hospedagem (como Booking.com, TripAdvisor, Hoteis.com, etc.) ou notícias/páginas locais.\n` +
    `3. Use APENAS preços em Reais (BRL) explicitamente mencionados nos resultados de pesquisa atuais.\n` +
    `4. Se não encontrar um preço numérico claro, retorne price: -1.\n` +
    `5. Se houver faixa de preços, retorne o menor valor da diária/tarifa padrão.\n\n` +
    `Responda SOMENTE com o JSON conforme o schema abaixo.`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    tools: [{ google_search: {} }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          price:    { type: 'NUMBER'  },
          currency: { type: 'STRING'  },
          details:  { type: 'STRING'  },
          status:   { type: 'STRING'  }
        },
        required: ['price', 'currency', 'details', 'status']
      }
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error(`Gemini HTTP ${resp.getResponseCode()}: ${resp.getContentText().substring(0, 200)}`);
  }

  const json = JSON.parse(resp.getContentText());
  if (json.candidates?.[0]?.content?.parts?.[0]) {
    return JSON.parse(json.candidates[0].content.parts[0].text);
  }

  throw new Error('Resposta Gemini em formato inesperado.');
}

/**
 * Retorna a instrução de busca de preço baseada na atividade
 */
function _orientacaoPreco(atividade) {
  const a = atividade.toLowerCase();
  if (a.includes('hospedagem') || a.includes('hotel') || a.includes('pousada') || a.includes('hostel')) {
    return 'Busque a diária padrão, diária mínima, tarifa "a partir de" ou diária de quarto duplo/casal. Extraia apenas o valor numérico mais representativo.';
  }
  if (a.includes('acampamento') || a.includes('camping')) {
    return 'Busque o valor da diária de camping, pernoite por pessoa ou tarifa de barraca. Retorne o valor mais básico encontrado.';
  }
  if (a.includes('restaurante') || a.includes('alimentação') || a.includes('bar') || a.includes('cafeteria') || a.includes('pizzaria') || a.includes('sorveteria')) {
    return 'Busque o preço médio do prato principal, prato feito (PF), buffet, prato do dia, pizza individual/média ou principal mais popular. Ignore preços de bebidas isoladas.';
  }
  if (a.includes('locadora') || a.includes('veículo') || a.includes('carro')) {
    return 'Busque a diária de locação de veículo econômico (menor categoria disponível). Retorne o menor valor de diária encontrado.';
  }
  if (a.includes('agência') || a.includes('turismo')) {
    return 'Busque o preço do passeio ou pacote mais básico/econômico oferecido. Retorne o menor valor de passeio encontrado.';
  }
  if (a.includes('náutico') || a.includes('pesca') || a.includes('mergulho') || a.includes('barco') || a.includes('lancha')) {
    return 'Busque o preço do passeio náutico mais básico, hora de aluguel de embarcação, ou taxa de mergulho. Retorne o menor valor por pessoa.';
  }
  if (a.includes('evento') || a.includes('casamento') || a.includes('convenção') || a.includes('infraestrutura')) {
    return 'Busque o valor de locação do espaço, taxa de entrada, ou pacote de evento mais básico. Retorne o menor valor encontrado.';
  }
  if (a.includes('parque') || a.includes('lazer') || a.includes('aquático')) {
    return 'Busque o valor do ingresso individual, diária ou entrada. Retorne o preço mais básico para adulto.';
  }
  if (a.includes('transporte') || a.includes('transfer')) {
    return 'Busque o preço da passagem, translado ou frete mais básico por pessoa. Retorne o menor valor encontrado.';
  }
  if (a.includes('prestador') || a.includes('especializado')) {
    return 'Busque o preço do serviço principal ou atividade mais básica oferecida. Retorne o menor valor claro encontrado.';
  }
  return 'Busque o preço do produto ou serviço principal oferecido. Retorne o valor numérico mais representativo encontrado na página.';
}


// ============================================================
// LIMPEZA DE HTML
// ============================================================
function _extrairTextoLimpo(html) {
  if (!html) return '';
  let text = html;

  // Remove doctype e tags de metadados
  text = text.replace(/<!DOCTYPE[^>]*>/gi, '');
  text = text.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  text = text.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
  text = text.replace(/<head\b[^<]*(?:(?!<\/head>)<[^<]*)*<\/head>/gi, '');
  text = text.replace(/<svg\b[^<]*(?:(?!<\/svg>)<[^<]*)*<\/svg>/gi, '');
  text = text.replace(/<!--[\s\S]*?-->/g, '');
  text = text.replace(/<noscript\b[^<]*(?:(?!<\/noscript>)<[^<]*)*<\/noscript>/gi, '');
  
  // Remove elementos estruturais e midia irrelevantes para preco (economiza tokens)
  text = text.replace(/<nav\b[^<]*(?:(?!<\/nav>)<[^<]*)*<\/nav>/gi, ' ');
  text = text.replace(/<footer\b[^<]*(?:(?!<\/footer>)<[^<]*)*<\/footer>/gi, ' ');
  text = text.replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, ' ');
  text = text.replace(/<video\b[^<]*(?:(?!<\/video>)<[^<]*)*<\/video>/gi, ' ');
  text = text.replace(/<audio\b[^<]*(?:(?!<\/audio>)<[^<]*)*<\/audio>/gi, ' ');
  text = text.replace(/<canvas\b[^<]*(?:(?!<\/canvas>)<[^<]*)*<\/canvas>/gi, ' ');
  text = text.replace(/<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>/gi, ' ');
  text = text.replace(/<embed\b[^<]*(?:(?!<\/embed>)<[^<]*)*<\/embed>/gi, ' ');
  text = text.replace(/<map\b[^<]*(?:(?!<\/map>)<[^<]*)*<\/map>/gi, ' ');
  text = text.replace(/<area\b[^<]*(?:(?!<\/area>)<[^<]*)*<\/area>/gi, ' ');

  // Remove todos os atributos de tags (mantém só o texto)
  text = text.replace(/<\/?[a-z][a-z0-9]*[^>]*>/gi, ' ');

  // Entidades HTML comuns
  text = text
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)))
    .replace(/&[a-z]+;/gi, ' ');

  // Colapsar espaços múltiplos
  text = text.replace(/\s+/g, ' ').trim();

  return text;
}


// ============================================================
// VALIDAÇÃO DE URL
// ============================================================
function _validarURL(url) {
  if (!url || url.toString().trim() === '' || url === '-') {
    return { ok: false, motivo: 'URL vazia ou ausente' };
  }

  const u = url.toString().trim();

  // Typo: ww. em vez de www.
  if (/^https?:\/\/ww\./.test(u)) {
    return { ok: false, motivo: 'URL com typo (ww. em vez de www.)' };
  }

  // Contém @ fora do protocolo (típico de Instagram handle)
  if (u.replace(/^https?:\/\//, '').includes('@')) {
    return { ok: false, motivo: 'URL inválida (contém @, provável handle de rede social)' };
  }

  // Não começa com http
  if (!/^https?:\/\//i.test(u)) {
    return { ok: false, motivo: 'URL sem protocolo http/https' };
  }

  return { ok: true };
}

function _detectarPlataformaJS(url) {
  const u = url.toLowerCase();
  for (const plat of CFG.PLATAFORMAS_JS) {
    if (u.includes(plat)) return plat;
  }
  return null;
}


// ============================================================
// COLETA DE LINHAS PARA PROCESSAR
// ============================================================
function _coletarLinhas(sheet, apenasSelecionados) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rowsToProcess: [], aviso: 'Nenhum dado encontrado na aba "Monitoramento".' };

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  let selectedRows = [];

  if (apenasSelecionados) {
    const active = sheet.getActiveRange();
    if (!active) return { rowsToProcess: [], aviso: 'Selecione as linhas que deseja atualizar antes de clicar.' };
    const start = active.getRow(), num = active.getNumRows();
    for (let r = start; r < start + num; r++) {
      if (r >= 2 && r <= lastRow) selectedRows.push(r);
    }
    if (selectedRows.length === 0) return { rowsToProcess: [], aviso: 'Nenhuma linha selecionada.' };
  }

  const rows = [];
  data.forEach((row, i) => {
    const rowNum = i + 2;
    if (apenasSelecionados && !selectedRows.includes(rowNum)) return;
    const site = (row[3] || '').toString().trim();
    if (site && site !== '-' && site !== '') {
      rows.push({
        index    : i,
        rowNum,
        id       : row[0],
        nome     : row[1],
        atividade: row[2],
        site,
        precoAtual: row[5]
      });
    }
  });

  if (rows.length === 0) return { rowsToProcess: [], aviso: 'Nenhum estabelecimento com URL válida encontrado.' };
  return { rowsToProcess: rows, aviso: null };
}


// ============================================================
// GATILHOS (TRIGGERS)
// ============================================================
function configurarGatilhoDiario() {
  _removerGatilhosPorFuncao('atualizarTodosPrecos');
  ScriptApp.newTrigger('atualizarTodosPrecos')
    .timeBased().everyDays(1).atHour(2).create();
  SpreadsheetApp.getUi().alert(
    '✅ Gatilho Configurado',
    'O monitoramento automático foi configurado para rodar diariamente às 02:00.\n\n' +
    'Nota: A execução ocorre em segundo plano, sem abrir a planilha.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function removerGatilhosDiarios() {
  const removidos = _removerGatilhosPorFuncao('atualizarTodosPrecos');
  SpreadsheetApp.getUi().alert(
    'Gatilhos Removidos',
    `${removidos} gatilho(s) de atualização automática foram removidos.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _removerTriggerContinuacao() {
  _removerGatilhosPorFuncao(CFG.TRIGGER_FN);
}

function _removerGatilhosPorFuncao(nomeFuncao) {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === nomeFuncao) {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  return count;
}


// ============================================================
// UTILITÁRIOS
// ============================================================
function resetarStatus() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Confirmação', 'Isso vai resetar o status de TODOS os itens para "Aguardando".\nPreços já coletados serão preservados. Continuar?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_MONITOR);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const range = sheet.getRange(2, CFG.COLS.STATUS, lastRow - 1, 1);
  const values = range.getValues().map(() => ['Aguardando']);
  range.setValues(values);
  ui.alert('✅ Status Resetado', 'Todos os itens voltaram para "Aguardando".', ui.ButtonSet.OK);
}

function mostrarResumo() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_MONITOR);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Sem dados.'); return; }

  const statuses = sheet.getRange(2, CFG.COLS.STATUS, lastRow - 1, 1).getValues().flat();
  const count = (match) => statuses.filter(s => s.toString().startsWith(match)).length;

  const total       = statuses.length;
  const sucesso     = count('Sucesso');
  const naoEnc      = count('Preço não encontrado');
  const erro        = count('Erro');
  const platJS      = count('Plataforma JS');
  const urlInv      = count('URL Inválida');
  const aguardando  = count('Aguardando') + count('Processando');
  const cobertura   = total > 0 ? ((sucesso / total) * 100).toFixed(1) : 0;

  SpreadsheetApp.getUi().alert(
    '📊 Resumo do Monitoramento',
    `Total de estabelecimentos : ${total}\n` +
    `✅ Preço encontrado        : ${sucesso} (${cobertura}% de cobertura)\n` +
    `🔍 Preço não encontrado    : ${naoEnc}\n` +
    `❌ Erros de acesso         : ${erro}\n` +
    `🚫 Plataforma JS           : ${platJS}\n` +
    `⚠️ URL Inválida            : ${urlInv}\n` +
    `⏳ Aguardando              : ${aguardando}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function _atualizarPainelExecucao(ss, inicio, atual, total, situacao) {
  const s = ss.getSheetByName(CFG.SHEET_PAINEL);
  if (!s) return;
  if (inicio) s.getRange('B22').setValue(Utilities.formatDate(inicio, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'));
  s.getRange('B23').setValue(`Lote ${atual}/${total}`);
  s.getRange('B24').setValue(situacao);
}

function _logErro(funcao, nome, erro) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CFG.SHEET_LOG);
    if (!sheet) return;
    sheet.appendRow([new Date(), funcao, nome, erro]);
  } catch (e) {}
}

function _podeMostrarUI() {
  try {
    return SpreadsheetApp.getUi() !== null;
  } catch (e) {
    return false;
  }
}

function _mostrarAlerta(titulo, mensagem) {
  if (_podeMostrarUI()) {
    SpreadsheetApp.getUi().alert(titulo, mensagem, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    console.log(`[ALERTA] ${titulo}: ${mensagem}`);
  }
}

function _esperarComBackoff(tentativa, isRateLimit) {
  const baseMs = isRateLimit ? 10000 : 2000; // 10s para rate limit, 2s para outro erro
  const sleepMs = baseMs * Math.pow(2, tentativa - 1);
  const maxSleepMs = 120000; // Limite máximo de 2 minutos
  const finalSleepMs = Math.min(sleepMs, maxSleepMs);
  
  console.log(`[BACKOFF] Tentativa ${tentativa} (${isRateLimit ? '429 Rate Limit' : 'Erro'}). Aguardando ${finalSleepMs / 1000}s...`);
  Utilities.sleep(finalSleepMs);
}

function _lerConfigs(ss) {
  const sheetConfig = ss.getSheetByName(CFG.SHEET_PAINEL);
  if (!sheetConfig) {
    _mostrarAlerta('Erro', 'Aba "Painel de Controle" não encontrada. Execute "Configuração Inicial" primeiro.');
    return null;
  }
  const apiKey = sheetConfig.getRange(CFG.API_CELL).getValue().toString().trim();
  if (!apiKey || apiKey.startsWith('INSIRA')) {
    _mostrarAlerta('⚠️ Configuração Necessária', 'Insira sua chave da API Gemini na célula B3 da aba "Painel de Controle".\n\nObtenha a chave gratuita em: aistudio.google.com');
    return null;
  }
  const sheetMonitor = ss.getSheetByName(CFG.SHEET_MONITOR);
  if (!sheetMonitor) {
    _mostrarAlerta('Erro', 'Aba "Monitoramento" não encontrada. Execute "Configuração Inicial" primeiro.');
    return null;
  }
  return { apiKey, sheet: sheetMonitor };
}

function _getBatchSize(ss) {
  try { return parseInt(ss.getSheetByName(CFG.SHEET_PAINEL).getRange(CFG.BATCH_SIZE_CELL).getValue(), 10) || CFG.BATCH_SIZE; } catch(e) { return CFG.BATCH_SIZE; }
}
function _getDelay(ss) {
  try { return parseInt(ss.getSheetByName(CFG.SHEET_PAINEL).getRange(CFG.DELAY_CELL).getValue(), 10) || CFG.DELAY_MS; } catch(e) { return CFG.DELAY_MS; }
}
function _getModelo(ss) {
  try { return ss.getSheetByName(CFG.SHEET_PAINEL).getRange(CFG.MODEL_CELL).getValue().toString().trim() || CFG.GEMINI_MODEL; } catch(e) { return CFG.GEMINI_MODEL; }
}


// ============================================================
// DADOS — CSV Completo dos 351 Estabelecimentos
// Erros corrigidos: typos de URL, &amp; → &, handles de redes sociais pré-marcados
// ============================================================
function _getDadosCSV() {
  return [
    ["SS0479","Locadora Global","13 - Locadora de veículos para turistas","https://www.locadoraglobal.com.br","","","","Aguardando","",""],
    ["SS0083","Hospedaria Sol do Araçá","03 - Meio de Hospedagem","https://www.soldoaraca.com.br","","","","Aguardando","",""],
    ["SS0100","Hotel Jota","03 - Meio de Hospedagem","https://www.hoteljota.com","","","","Aguardando","",""],
    ["SS0144","Miradouro de São Sebastião","03 - Meio de Hospedagem","https://www.miradourohotel.com.br","","","","Aguardando","",""],
    ["SS0411","Scuba do Dive Adventure Turismo","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.scubadodive.com.br","","","","Aguardando","",""],
    ["SS0540","Espaço Don Laurindo","12 - Prestador de Infraestrutura de apoio para eventos","https://www.donlaurindo.com","","","","Aguardando","",""],
    ["SS0107","Hotel Pousada Toque Toque Pequeno (Pousada Aparas)","03 - Meio de Hospedagem","https://www.pousadaaparas.com.br","","","","Aguardando","",""],
    ["SS0157","Pousada Alapage","03 - Meio de Hospedagem","https://alapa.stays.net/pt/","","","","Plataforma JS — Sem acesso estático","Site usa plataforma stays.net que requer JavaScript para exibir preços.",""],
    ["SS0203","Pousada Hybisco","03 - Meio de Hospedagem","https://www.pousadahybisco.com.br","","","","Aguardando","",""],
    ["SS0234","Pousada Recanto Caiçara Toque-toque pequeno","03 - Meio de Hospedagem","https://recantocaicara.com.br/","","","","Aguardando","",""],
    ["SS0371","Toque Toque Tur","02 - Agência de Turismo","https://www.toquetoquetur.com","","","","Aguardando","",""],
    ["SS0374","União Turismo Ecológico (Mako Dive Center)","02 - Agência de Turismo","https://www.makobr.com","","","","Aguardando","",""],
    ["SS0116","Ilha de Toque Toque Eco Boutique Hotel","03 - Meio de Hospedagem","https://www.ilhadetoquetoque.com.br","","","","Aguardando","",""],
    ["SS0168","Pousada Bougainville","03 - Meio de Hospedagem","https://www.bougainvilletoquetoque.com.br/","","","","Aguardando","",""],
    ["SS0595","Alcatrazes Restaurante e Pizzaria","15 - Restaurante cafeteria bar e similares","https://www.restaurantealcatrazes.com.br","","","","Aguardando","",""],
    ["SS0600","Azzura Pizzas","15 - Restaurante cafeteria bar e similares","https://azzurapizzas.com.br","","","","Aguardando","",""],
    ["SS0601","Badauê Juquehy","15 - Restaurante cafeteria bar e similares","http://www.restaurantebadaue.com.br/juquehy/","","","","Aguardando","",""],
    ["SS0485","Alegria no Mar","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.alegrianomar.com.br","","","","Aguardando","",""],
    ["SS0482","CuatroCinco Guest House","03 - Meio de Hospedagem","https://www.cuatrocincobr.com","","","","Aguardando","",""],
    ["SS0602","Badauê Maresias","15 - Restaurante cafeteria bar e similares","http://www.restaurantebadaue.com.br/maresias/","","","","Aguardando","",""],
    ["SS0383","Passeios Náuticos Litoral","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.passeiolitoral.com.br","","","","Aguardando","",""],
    ["SS0521","Conhecendo Ilhabela","02 - Agência de Turismo","https://www.instagram.com/conhecendo_ilhabela","","","","Plataforma JS — Sem acesso estático","Site usa instagram.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0490","Capitão Ximango","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.barcocapitaoximango.com.br","","","","Aguardando","",""],
    ["SS0514","JD Locadora","13 - Locadora de veículos para turistas","https://www.jdturismo.com.br","","","","Aguardando","",""],
    ["SS0358","Nacional Ecotur","02 - Agência de Turismo","https://www.nacionalecotur.com.br","","","","Aguardando","",""],
    ["SS0376","Universo Marinho","02 - Agência de Turismo","https://www.universomarinho.com.br","","","","Aguardando","",""],
    ["SS0690","Nai Santiago","03 - Meio de Hospedagem","https://www.naisantiago.com.br","","","","Aguardando","",""],
    ["SS0013","Abricó Beach Hotel","03 - Meio de Hospedagem","https://www.abricobeachhotel.com.br","","","","Aguardando","",""],
    ["SS0034","Boulevard Avenida São Sebastião","03 - Meio de Hospedagem","http://boulevardsaosebastiao.com.br/","","","","Aguardando","",""],
    ["SS0570","Castelo das Artes Comunicação e Eventos","04 - Organizadora de Eventos","https://www.ocastelodasartes.com","","","","Aguardando","",""],
    ["SS0099","Hotel Guarda Mor","03 - Meio de Hospedagem","https://www.hotelguardamor.com.br","","","","Aguardando","",""],
    ["SS0108","Hotel Recanto dos Pássaros","03 - Meio de Hospedagem","https://www.recantodospassaros.tur.br","","","","Aguardando","",""],
    ["SS0480","Localiza Rent a Car","13 - Locadora de veículos para turistas","https://www.localiza.com","","","","Aguardando","",""],
    ["SS0946","Pararanga Náutica","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.pararanga.com.br","","","","Aguardando","",""],
    ["SS0155","Porto Grande Hotel & Convention","03 - Meio de Hospedagem","https://www.portograndehotel.com.br","","","","Aguardando","",""],
    ["SS0363","Proa Turismo","02 - Agência de Turismo","https://www.proaturismo.com.br","","","","Aguardando","",""],
    ["SS0380","Azimuth Escola Náutica","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.azimuthnautica.com.br","","","","Aguardando","",""],
    ["SS0334","Conheça Ilhabela","02 - Agência de Turismo","https://www.conhecailhabela.com.br","","","","Aguardando","",""],
    ["SS0944","Espaço Arandela","09 - Casa de espetáculos e equipamento de animação turística","https://www.espacoarandela.com.br","","","","Aguardando","",""],
    ["SS0386","Fridsland & Co. LTDA","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.fridsland.com.br","","","","Aguardando","",""],
    ["SS0095","Hotel do Sol Executive","03 - Meio de Hospedagem","https://www.hoteldosolexecutive.com.br","","","","Aguardando","",""],
    ["SS0114","Hotel Veleiro","03 - Meio de Hospedagem","https://www.hotelveleiro.com","","","","Aguardando","",""],
    ["SS0392","Key Marine","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.lojakeymarine.com.br","","","","Aguardando","",""],
    ["SS0393","Lemar Náutica","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.lemarnautica.com.br","","","","Aguardando","",""],
    ["SS0406","MB Lanchas","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.mblanchas.com.br","","","","Aguardando","",""],
    ["SS0377","Voosh Viagens e Turismo","02 - Agência de Turismo","https://www.vooshviagens.com.br","","","","Aguardando","",""],
    ["SS0004","Camping do Mazinho","01 - Acampamento Turístico","https://www.campingdomazinho.com.br","","","","Aguardando","",""],
    ["SS0534","Casa Mira","12 - Prestador de Infraestrutura de apoio para eventos","https://www.mira.casa","","","","Aguardando","",""],
    ["SS0055","Ciribaí Praia Hotel","03 - Meio de Hospedagem","https://www.ciribaipraiahotel.com.br","","","","Aguardando","",""],
    ["SS0617","Brasileira Gourmet","15 - Restaurante cafeteria bar e similares","https://www.brasileiragourmet.com","","","","Aguardando","",""],
    ["SS0150","Paúba Beach Hotel","03 - Meio de Hospedagem","https://www.paubabeach.com.br","","","","Aguardando","",""],
    ["SS0151","Paúba Praia Hotel","03 - Meio de Hospedagem","https://www.paubapraiahotel.com.br","","","","Aguardando","",""],
    ["SS0156","Porto Paúba","03 - Meio de Hospedagem","https://www.portopauba.com.br","","","","Aguardando","",""],
    ["SS0260","Pousada Vila Ipuan","03 - Meio de Hospedagem","https://www.pousadavilaipuan.com.br","","","","Aguardando","",""],
    ["SS0466","Residencial Marina del Sol","03 - Meio de Hospedagem","https://www.marinadelsol.com.br","","","","Aguardando","",""],
    ["SS0297","Tartarugas de Paúba","03 - Meio de Hospedagem","https://www.pousadatartarugasdepauba.com.br","","","","Aguardando","",""],
    ["SS0016","Amora Hotel","03 - Meio de Hospedagem","https://www.amorahotel.com.br","","","","Aguardando","",""],
    ["SS0017","Amoreiras Hotel Pousada","03 - Meio de Hospedagem","https://www.amoreiras.com.br","","","","Aguardando","",""],
    ["SS0020","Arco Iris I","03 - Meio de Hospedagem","https://www.arcoirischales.com.br","","","","Aguardando","",""],
    ["SS0022","Arco-Íris II","03 - Meio de Hospedagem","https://www.arcoirischales.com.br","","","","Aguardando","",""],
    ["SS0026","Azul Banana","03 - Meio de Hospedagem","https://www.azulbanana.com.br","","","","Aguardando","",""],
    ["SS0164","Beach Camp Maresias","03 - Meio de Hospedagem","https://www.beachcamp.com.br","","","","Aguardando","",""],
    ["SS0031","Beach Hotel Maresias","03 - Meio de Hospedagem","https://www.maresiasbeachhotel.com.br","","","","Aguardando","",""],
    ["SS0049","Chalés Água Grande","03 - Meio de Hospedagem","https://www.chalesaguagrande.com","","","","Aguardando","",""],
    ["SS0036","Chalés Cabelo Maresias","03 - Meio de Hospedagem","https://www.cabelomaresias.com/","","","","Aguardando","",""],
    ["SS0051","Chalés do Paulo","03 - Meio de Hospedagem","https://www.chalesdopaulomaresias.com","","","","Aguardando","",""],
    ["SS0054","Chili-In Suites Maresias","03 - Meio de Hospedagem","https://www.chiliinmaresias.com.br","","","","Aguardando","",""],
    ["SS0069","Coconut's Maresias Hotel","03 - Meio de Hospedagem","https://www.coconutshotel.com.br","","","","Aguardando","",""],
    ["SS0058","Condomínio Verdes Mares","03 - Meio de Hospedagem","https://www.pousadaverdesmares.com.br","","","","Aguardando","",""],
    ["SS0059","Costa Verde Road Trips","03 - Meio de Hospedagem","https://www.costaverderoadtrips.com.br","","","","Aguardando","",""],
    ["SS0061","Delta Maresias Club","03 - Meio de Hospedagem","https://www.deltamaresiasclub.com.br","","","","Aguardando","",""],
    ["SS0062","Descanso do Cansado","03 - Meio de Hospedagem","https://www.pousadadescansodocansado.com.br","","","","Aguardando","",""],
    ["SS0065","Duke Beach Hotel Maresias","03 - Meio de Hospedagem","https://www.dukebeach.com.br","","","","Aguardando","",""],
    ["SS0186","Villa Garden Maresias","03 - Meio de Hospedagem","https://www.villagardenmaresias.com.br/br/","","","","Aguardando","",""],
    ["SS0637","Chocomel Doces","15 - Restaurante cafeteria bar e similares","https://chocomel-doces.goomer.app/menu","","","","Plataforma JS — Sem acesso estático","Site usa goomer.app, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0082","Hospedaria Recanto da Barra","03 - Meio de Hospedagem","https://www.facebook.com/recantodabarramaresias","","","","Plataforma JS — Sem acesso estático","Site usa facebook.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0084","Hostel Caiçara Maresias","03 - Meio de Hospedagem","https://www.hostelcaicara.com.br/","","","","Aguardando","",""],
    ["SS0096","Hotel e Pousada Canto Mágico","03 - Meio de Hospedagem","https://www.cantomagico.com.br","","","","Aguardando","",""],
    ["SS0097","Hotel e Pousada Katmandu","03 - Meio de Hospedagem","https://www.katmandumaresias.com.br","","","","Aguardando","",""],
    ["SS0101","Hotel Kiribati","03 - Meio de Hospedagem","https://www.hotelkiribati.com.br","","","","Aguardando","",""],
    ["SS0115","Hotel Pousada dos Condes","03 - Meio de Hospedagem","https://www.pousadadoscondes.com.br","","","","Aguardando","",""],
    ["SS0094","Hotel Canto do Rio Maresias","03 - Meio de Hospedagem","https://www.hotelcantodorio.com.br","","","","Aguardando","",""],
    ["SS0118","Jungle Beach","03 - Meio de Hospedagem","https://www.junglebeach.tur.br","","","","Aguardando","",""],
    ["SS0125","Maui Maresias","03 - Meio de Hospedagem","https://www.mauimaresias.com.br","","","","Aguardando","",""],
    ["SS0126","Kyrios Pousada","03 - Meio de Hospedagem","https://www.kyriospousada.com.br","","","","Aguardando","",""],
    ["SS0130","Lua Chalés","03 - Meio de Hospedagem","https://www.luachalesmaresias.com.br","","","","Aguardando","",""],
    ["SS0518","Luai Cabanas","10 - Centro de convenções","https://www.luaicabanas.com.br","","","","Aguardando","",""],
    ["SS0958","Daluka Sorvetes e Cia","15 - Restaurante cafeteria bar e similares","https://daluka-sorvetes-e-cia.ola.click/","","","","Plataforma JS — Sem acesso estático","Site usa ola.click, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0136","Marerê","03 - Meio de Hospedagem","https://www.pousadamarere.com","","","","Aguardando","",""],
    ["SS0348","Maresias de Itu - Residencial","03 - Meio de Hospedagem","https://maresiasdeitu.negocio.site/","","","","Plataforma JS — Sem acesso estático","Site usa negocio.site, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0137","Maresias Hostel","03 - Meio de Hospedagem","https://www.maresiashostel.com.br","","","","Aguardando","",""],
    ["SS0139","Maresias Praia Hotel","03 - Meio de Hospedagem","https://www.maresiaspraiahotel.com.br","","","","Aguardando","",""],
    ["SS0447","Maresias Tur","11 - Prestador especializado em segmentos turístico","https://www.maresiastur.com.br","","","","Aguardando","",""],
    ["SS0645","Deck Bar Canoa Barra do Una","15 - Restaurante cafeteria bar e similares","https://bistrocanoa.business.site","","","","Plataforma JS — Sem acesso estático","Site usa business.site, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0945","Olisa Hotel Boutique","03 - Meio de Hospedagem","https://www.olisa.com.br","","","","Aguardando","",""],
    ["SS0159","Pousada Aysu Maresias","03 - Meio de Hospedagem","https://aysumaresias.com.br/contato/","","","","Aguardando","",""],
    ["SS0028","Pousada Azul da Cor do Mar","03 - Meio de Hospedagem","https://www.azuldacordomar.com.br","","","","Aguardando","",""],
    ["SS0169","Pousada Brig","03 - Meio de Hospedagem","https://www.pousadabrigmaresias.com.br","","","","Aguardando","",""],
    ["SS0176","Pousada Cantos das Estrelas","03 - Meio de Hospedagem","https://www.pousadamardeny.com.br","","","","Aguardando","",""],
    ["SS0180","Pousada Catavento","03 - Meio de Hospedagem","https://www.chalescataventomaresias.com.br/","","","","Aguardando","",""],
    ["SS0191","Pousada Delta Mar Maresias","03 - Meio de Hospedagem","https://www.deltamarpousada.com.br","","","","Aguardando","",""],
    ["SS0202","Pousada Frente Mar","03 - Meio de Hospedagem","https://www.frentemar.com.br","","","","Aguardando","",""],
    ["SS0205","Pousada In Bali Maresias","03 - Meio de Hospedagem","https://www.pousadainbali.com.br/","","","","Aguardando","",""],
    ["SS0212","Pousada Mandala Maresias","03 - Meio de Hospedagem","https://www.pousadamandalamaresias.com.br/contact.html","","","","Aguardando","",""],
    ["SS0214","Pousada Maréatoa","03 - Meio de Hospedagem","https://www.mareatoa.com.br","","","","Aguardando","",""],
    ["SS0215","Pousada Maresias Beira-Mar","03 - Meio de Hospedagem","https://www.maresiasbeiramar.com.br","","","","Aguardando","",""],
    ["SS0217","Pousada Mosaico Brasil em Maresias","03 - Meio de Hospedagem","https://www.pousadamosaico.com.br","","","","Aguardando","",""],
    ["SS0218","Pousada Mosaico Maresias","03 - Meio de Hospedagem","https://www.pousadamosaico.com.br","","","","Aguardando","",""],
    ["SS0219","Pousada Mundial","03 - Meio de Hospedagem","https://www.pousadamundial.com.br","","","","Aguardando","",""],
    ["SS0222","Pousada Nusa Dua Maresias","03 - Meio de Hospedagem","https://www.pousadanusadua.com","","","","Aguardando","",""],
    ["SS0223","Pousada Pantai Maresias","03 - Meio de Hospedagem","https://www.pantaimaresias.com.br","","","","Aguardando","",""],
    ["SS0225","Pousada Pé da Mata","03 - Meio de Hospedagem","https://www.pedamata.com","","","","Aguardando","",""],
    ["SS0228","Pousada Porto Mare","03 - Meio de Hospedagem","https://www.pousadaportomare.com.br","","","","Aguardando","",""],
    ["SS0230","Pousada Pura Vida","03 - Meio de Hospedagem","https://www.puravidamaresias.com.br","","","","Aguardando","",""],
    ["SS0235","Pousada Refúgio de Maresias","03 - Meio de Hospedagem","https://www.refugiomaresias.com.br","","","","Aguardando","",""],
    ["SS0247","Pousada Toca da Anita","03 - Meio de Hospedagem","https://tocadaanitamaresias.com.br/","","","","Aguardando","",""],
    ["SS0248","Pousada Toca da Praia","03 - Meio de Hospedagem","https://www.tocadapraia.com.br","","","","Aguardando","",""],
    ["SS0254","Pousada Vida Mansa","03 - Meio de Hospedagem","https://www.vidamansamaresias.com.br/","","","","Aguardando","",""],
    ["SS0263","Pousada Vila Maresias","03 - Meio de Hospedagem","https://www.vilamaresias.com.br","","","","Aguardando","",""],
    ["SS0265","Pousada Villa Blu","03 - Meio de Hospedagem","https://www.villablumaresias.com.br","","","","Aguardando","",""],
    ["SS0275","Recanto da Nina","03 - Meio de Hospedagem","https://pousadamaresiasrecantonina.com.br/","","","","Aguardando","",""],
    ["SS0659","Familia Restaurante","15 - Restaurante cafeteria bar e similares","https://www.familiarestaurante.com.br","","","","Aguardando","",""],
    ["SS0667","Guato Gastronomia","15 - Restaurante cafeteria bar e similares","https://www.guato.com.br","","","","Aguardando","",""],
    ["SS0478","Sirena Maresias","09 - Casa de espetáculos e equipamento de animação turística","https://www.sirena.com.br/","","","","Aguardando","",""],
    ["SS0367","Maresias Escola Náutica","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.ecoturmaresias.com","","","","Aguardando","",""],
    ["SS0296","Tambayba Hotel Pousada","03 - Meio de Hospedagem","https://www.tambayba.com.br","","","","Aguardando","",""],
    ["SS0854","Jota R Pizzaria & Restaurante","15 - Restaurante cafeteria bar e similares","https://www.ifood.com.br/delivery/sao-sebastiao-sp/pizzaria-jota-r-barra-do-una/4fbafe93-df7c-40aa-8d05-66a474d91b97","","","","Plataforma JS — Sem acesso estático","Site usa ifood.com.br, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0302","Tubes Maresias","03 - Meio de Hospedagem","https://www.tubesmaresias.com.br/","","","","Aguardando","",""],
    ["SS0305","Verdes Mares","03 - Meio de Hospedagem","https://www.pousadaverdesmares.com.br","","","","Aguardando","",""],
    ["SS0307","Vila Maremota","03 - Meio de Hospedagem","https://www.villamaremotta.com.br","","","","Aguardando","",""],
    ["SS0308","Vilarejo Chalé","03 - Meio de Hospedagem","https://www.vilarejomaresias.com.br","","","","Aguardando","",""],
    ["SS0314","Villa Miracá","03 - Meio de Hospedagem","https://www.villamiraca.com.br","","","","Aguardando","",""],
    ["SS0317","Villa'I Mare","03 - Meio de Hospedagem","https://www.villamare.com.br","","","","Aguardando","",""],
    ["SS0110","Hotel Sambaqui","03 - Meio de Hospedagem","https://www.hotelsambaqui.com.br","","","","Aguardando","",""],
    ["SS0298","Tiê Cama e Café","03 - Meio de Hospedagem","https://www.tiecamaecafe.com.br","","","","Aguardando","",""],
    ["SS0322","Winner Beach Hotel","03 - Meio de Hospedagem","https://hotelwinnerbeach.com/","","","","Aguardando","",""],
    ["SS0042","Casa D'Aroeira","03 - Meio de Hospedagem","https://www.casadaroeira.com","","","","Aguardando","",""],
    ["SS0572","Espaço Amar Casamentos","04 - Organizadora de Eventos","https://www.amarcasamentos.com.br","","","","Aguardando","",""],
    ["SS0073","Flow Hostel","03 - Meio de Hospedagem","https://www.flowhostel.com.br","","","","Aguardando","",""],
    ["SS0441","Green Way Brasil","11 - Prestador especializado em segmentos turístico","https://www.greenway.com.br","","","","Aguardando","",""],
    ["SS0078","Hippocampus Juquehy Pousada","03 - Meio de Hospedagem","https://www.hippocampus.tur.br","","","","Aguardando","",""],
    ["SS0079","Hiu Hotel","03 - Meio de Hospedagem","https://www.hiuhotel.com.br","","","","Aguardando","",""],
    ["SS0122","Hotel Juquei Frente ao Mar","03 - Meio de Hospedagem","https://www.frenteaomar.com.br","","","","Aguardando","",""],
    ["SS0104","Hotel Pousada Baobá","03 - Meio de Hospedagem","https://www.pousadabaoba.com.br","","","","Aguardando","",""],
    ["SS0120","Juquehy La Plage Hotel","03 - Meio de Hospedagem","https://www.juquehylaplagehotel.com.br","","","","Aguardando","",""],
    ["SS0121","Juquehy Praia Hotel","03 - Meio de Hospedagem","https://www.juquehypraia.com.br","","","","Aguardando","",""],
    ["SS0119","Juqueí Beach Hotel","03 - Meio de Hospedagem","https://www.juqueihotel.com.br","","","","Aguardando","",""],
    ["SS0943","Juqy Beach House","04 - Organizadora de Eventos","https://www.juqybeachhouse.com.br","","","","Aguardando","",""],
    ["SS0132","Maison Chez Louise at Louis","03 - Meio de Hospedagem","https://www.louiselouis.com.br","","","","Aguardando","",""],
    ["SS0141","Mata Atlântica Pousada","03 - Meio de Hospedagem","https://www.pousadamataatlantica.com.br","","","","Aguardando","",""],
    ["SS0698","Mozebas Outside Bar","15 - Restaurante cafeteria bar e similares","https://mozebas.blogspot.com.br/","","","","Plataforma JS — Sem acesso estático","Site usa blogspot.com.br, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0145","Montão do Trigo Hotel e Restaurante","03 - Meio de Hospedagem","https://www.pousadamontaodotrigo.com.br","","","","Aguardando","",""],
    ["SS0161","Pousada Atobá Juquehy","03 - Meio de Hospedagem","https://www.pousadaatobajuquehy.com.br","","","","Aguardando","",""],
    ["SS0166","Pousada Bico Verde","03 - Meio de Hospedagem","https://www.pousadabicoverde.com.br","","","","Aguardando","",""],
    ["SS0182","Pousada Chalés da Lua Juquehy","03 - Meio de Hospedagem","https://www.chalesdaluajuquehy.com","","","","Aguardando","",""],
    ["SS0192","Pousada do Almirante","03 - Meio de Hospedagem","https://www.pousadadoalmirante.com.br","","","","Aguardando","",""],
    ["SS0201","Pousada Etoile","03 - Meio de Hospedagem","https://www.pousadaetoile.com.br","","","","Aguardando","",""],
    ["SS0207","Pousada Irashai","03 - Meio de Hospedagem","https://www.pousadairashai.com.br","","","","Aguardando","",""],
    ["SS0208","Pousada Jabuti","03 - Meio de Hospedagem","https://pousadajabutijuquehy.com.br/","","","","Aguardando","",""],
    ["SS0216","Pousada Moryba","03 - Meio de Hospedagem","https://www.moryba.com.br","","","","Aguardando","",""],
    ["SS0531","Pousada Pono Village","03 - Meio de Hospedagem","https://www.ponovillage.com","","","","Aguardando","",""],
    ["SS0242","Pousada Sol e Mar","03 - Meio de Hospedagem","https://www.pousadasolemaremjuquehy.com.br","","","","Aguardando","",""],
    ["SS0245","Pousada Terra","03 - Meio de Hospedagem","https://www.pousadaterra.com.br","","","","Aguardando","",""],
    ["SS0250","Pousada Tupinambá","03 - Meio de Hospedagem","https://pousadatupinamba.com.br","","","","Aguardando","",""],
    ["SS0258","Pousada Vila do Sol","03 - Meio de Hospedagem","https://pousadaviladosol.com.br/","","","","Aguardando","",""],
    ["SS0261","Pousada Vila Juquehy","03 - Meio de Hospedagem","https://www.pousadavilajuquehy.com.br/","","","","Aguardando","",""],
    ["SS0264","Pousada Vila Real","03 - Meio de Hospedagem","https://www.pousadavilareal.com.br","","","","Aguardando","",""],
    ["SS0267","Pousada Villa Encanto","03 - Meio de Hospedagem","https://www.villaencanto.com.br","","","","Aguardando","",""],
    ["SS0268","Pousada Villa Marítima","03 - Meio de Hospedagem","https://www.villamaritima.com.br","","","","Aguardando","",""],
    ["SS0276","Recanto dos Tangarás","03 - Meio de Hospedagem","https://recantodostangaras.com.br","","","","Aguardando","",""],
    ["SS0280","Recanto Verde Praia Hotel","03 - Meio de Hospedagem","https://www.recantoverdepraiahotel.com.br","","","","Aguardando","",""],
    ["SS0966","Nippoke Maresias","15 - Restaurante cafeteria bar e similares","https://nippokemaresias.menudino.com","","","","Plataforma JS — Sem acesso estático","Site usa menudino.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0238","Rosa dos Ventos Juquehy","03 - Meio de Hospedagem","https://www.rosadosventosjuquehy.com.br","","","","Aguardando","",""],
    ["SS0293","Solar do Parcel","03 - Meio de Hospedagem","https://www.solardoparcel.com.br","","","","Aguardando","",""],
    ["SS0464","Surf's Up Club","11 - Prestador especializado em segmentos turístico","https://www.surfupclub.com","","","","Aguardando","",""],
    ["SS0562","TF Eventos","12 - Prestador de Infraestrutura de apoio para eventos","https://www.tfeventos.com.br","","","","Aguardando","",""],
    ["SS0299","Tô na Praia Juquehy","03 - Meio de Hospedagem","https://www.pousadatonapraiajuquehy.com.br","","","","Aguardando","",""],
    ["SS0504","Marea Turismo Nautico","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.mareaturismonautico.com.br","","","","Aguardando","",""],
    ["SS0574","Flow Desenvolvimento Sustentável","04 - Organizadora de Eventos","https://www.flowsustentavel.com.br","","","","Aguardando","",""],
    ["SS0571","Costa Norte Aventuras","04 - Organizadora de Eventos","https://www.costanorteaventuras.com.br","","","","Aguardando","",""],
    ["SS0349","Gonçalves Turismo","02 - Agência de Turismo","https://www.goncalvesturismo.tur.br","","","","Aguardando","",""],
    ["SS0018","Ana Doce Pousada","03 - Meio de Hospedagem","https://www.pousadaanadoce.com.br","","","","Aguardando","",""],
    ["SS0347","Fragatas Viagens e Turismo","02 - Agência de Turismo","https://www.fragatas.com.br","","","","Aguardando","",""],
    ["SS0085","Hostel Central","03 - Meio de Hospedagem","https://www.hostelcentralsaosebastiao.com.br","","","","Aguardando","",""],
    ["SS0106","Hotel Pousada Garoupas","03 - Meio de Hospedagem","https://www.pousadagaroupas.com.br","","","","Aguardando","",""],
    ["SS0109","Hotel Roma","03 - Meio de Hospedagem","https://www.hotelroma.tur.br","","","","Aguardando","",""],
    ["SS0734","Pizzaria da Vila","15 - Restaurante cafeteria bar e similares","https://www.davilapizzaria.com.br","","","","Aguardando","",""],
    ["SS0183","Pousada da Banda","03 - Meio de Hospedagem","https://www.pousadadabanda.com.br","","","","Aguardando","",""],
    ["SS0188","Pousada da Sesmaria","03 - Meio de Hospedagem","https://www.pousadadasesmaria.com.br","","","","Aguardando","",""],
    ["SS0195","Pousada do Ipê","03 - Meio de Hospedagem","https://www.pousadaipesaosebastiao.com.br","","","","Aguardando","",""],
    ["SS0959","Proa Restaurante","15 - Restaurante cafeteria bar e similares","https://www.hubt.com.br/proarestaurante/","","","","Plataforma JS — Sem acesso estático","Site usa hubt.com.br, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0754","Quiosque Lima Limão","15 - Restaurante cafeteria bar e similares","https://quiosquelimalimao.wordpress.com","","","","Plataforma JS — Sem acesso estático","Site usa wordpress.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0583","RS Náutica","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.rsnautica.com.br","","","","Aguardando","",""],
    ["SS0763","Restaurante Amigos TPC","15 - Restaurante cafeteria bar e similares","https://tebarpraiaclube.com.br","","","","Aguardando","",""],
    ["SS0835","Stylo Lava Rápido e Estacionamento","16 - Outros","https://www.stylolavarapido.com.br","","","","Aguardando","",""],
    ["SS0764","Restaurante Antigas","15 - Restaurante cafeteria bar e similares","https://www.restauranteantigas.com.br","","","","Aguardando","",""],
    ["SS0515","Tebar Praia Clube","14 - Parque aquático e empreendimentos de lazer","https://www.tebarpraiaclube.com.br","","","","Aguardando","",""],
    ["SS0533","Ciranda de Dois","12 - Prestador de Infraestrutura de apoio para eventos","https://www.cirandadedois.com.br","","","","Aguardando","",""],
    ["SS0771","Restaurante Canoa Barra do Una","15 - Restaurante cafeteria bar e similares","https://canoa.com.br","","","","Aguardando","",""],
    ["SS0023","Armação de Camburi","03 - Meio de Hospedagem","https://www.pousadaarmacao.com.br","","","","Aguardando","",""],
    ["SS0027","Azul Banana - Camburi","03 - Meio de Hospedagem","https://www.azulbanana.com.br","","","","Aguardando","",""],
    ["SS0030","Beach Hotel Cambury","03 - Meio de Hospedagem","https://www.Camburihotel.com.br","","","","Aguardando","",""],
    ["SS0032","Beach Hotel Sunset","03 - Meio de Hospedagem","https://www.sunsethotel.com.br/","","","","Aguardando","",""],
    ["SS0941","Beach Hotel Villas","03 - Meio de Hospedagem","https://www.beachhotelvillas.com.br","","","","Aguardando","",""],
    ["SS0002","Camping Camburi","01 - Acampamento Turístico","https://www.campingcamburi.com.br","","","","Aguardando","",""],
    ["SS0041","Casa Bacarirá","03 - Meio de Hospedagem","https://www.casabacarira.com.br","","","","Aguardando","",""],
    ["SS0045","Casa do Pescador","03 - Meio de Hospedagem","https://casadopescador2.webnode.com/","","","","Plataforma JS — Sem acesso estático","Site usa webnode.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0050","Chalés do Camburi","03 - Meio de Hospedagem","https://www.chalesdocamburi.com.br","","","","Aguardando","",""],
    ["SS0056","Coco Bambu","03 - Meio de Hospedagem","https://www.pousadacocobambu.com.br","","","","Aguardando","",""],
    ["SS0338","Dreamsurf Travel","02 - Agência de Turismo","https://www.dreamsurftur.com.br","","","","Aguardando","",""],
    ["SS0339","Eco Dynamic","02 - Agência de Turismo","https://www.ecodynamic.com.br","","","","Aguardando","",""],
    ["SS0069","Estalagem Camburi","03 - Meio de Hospedagem","https://www.estalagemcamburi.com.br","","","","Aguardando","",""],
    ["SS0071","Flat Camburi","03 - Meio de Hospedagem","https://www.flatcamburi.com.br","","","","Aguardando","",""],
    ["SS0072","Flat Coqueiros","03 - Meio de Hospedagem","https://www.flatcoqueiros.com.br","","","","Aguardando","",""],
    ["SS0093","Hotel Camburi Praia","03 - Meio de Hospedagem","https://www.hotelcamburipraia.com.br","","","","Aguardando","",""],
    ["SS0113","Hotel Spa Nau Royal","03 - Meio de Hospedagem","https://www.nauroyal.com.br","","","","Aguardando","",""],
    ["SS0127","Laika Hostel","03 - Meio de Hospedagem","https://www.laikahostel.com.br/","","","","Aguardando","",""],
    ["SS0671","Paradiso Al Mare Pousada","03 - Meio de Hospedagem","https://www.paradisoalmare.com.br","","","","Aguardando","",""],
    ["SS0171","Pousada Camburi","03 - Meio de Hospedagem","https://www.pousadacamburi.com.br","","","","Aguardando","",""],
    ["SS0172","Pousada Camburizinho","03 - Meio de Hospedagem","https://www.pousadacamburizinho.com.br","","","","Aguardando","",""],
    ["SS0174","Pousada Canto do Camburi","03 - Meio de Hospedagem","https://www.pousadacantodocamburi.com.br","","","","Aguardando","",""],
    ["SS0187","Pousada da Morena","03 - Meio de Hospedagem","https://www.pousadadamorena.com.br","","","","Aguardando","",""],
    ["SS0190","Pousada das Praias","03 - Meio de Hospedagem","https://www.pousadadaspraias.com.br","","","","Aguardando","",""],
    ["SS0220","Pousada Náutica Camburioca","03 - Meio de Hospedagem","https://www.camburioca.com.br","","","","Aguardando","",""],
    ["SS0226","Pousada Pinheiro","03 - Meio de Hospedagem","https://www.pousadadopinheirocamburi.com.br","","","","Aguardando","",""],
    ["SS0227","Pousada Portal do Cacau","03 - Meio de Hospedagem","https://www.portaldocacau.com.br","","","","Aguardando","",""],
    ["SS0255","Pousada Vila Atlântica Inn","03 - Meio de Hospedagem","https://www.vilaatlantica.com.br","","","","Aguardando","",""],
    ["SS0257","Pousada Vila Camburi","03 - Meio de Hospedagem","https://www.vilaCamburi.com.br","","","","Aguardando","",""],
    ["SS0266","Pousada Villa dos Manacás","03 - Meio de Hospedagem","https://www.villamanakas.com.br","","","","Aguardando","",""],
    ["SS0197","Pousada do Rosa","03 - Meio de Hospedagem","https://www.pousadadorosa.com.br","","","","Aguardando","",""],
    ["SS0282","Refúgio de Camburi","03 - Meio de Hospedagem","https://www.refugiodecamburi.com.br","","","","Aguardando","",""],
    ["SS0790","Restaurante Ébano","15 - Restaurante cafeteria bar e similares","https://www.restauranteebano.com.br","","","","Aguardando","",""],
    ["SS0508","Transline Tour Litoral","13 - Locadora de veículos para turistas","https://www.translinetourlitoral.wixsite.com/vans","","","","Plataforma JS — Sem acesso estático","Site usa wixsite.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0301","Tree Park Camburi","03 - Meio de Hospedagem","https://www.treepark.com.br/bra/","","","","Aguardando","",""],
    ["SS0304","Alua Camburi","03 - Meio de Hospedagem","https://www.ventosdocamburi.com.br","","","","Aguardando","",""],
    ["SS0309","Villa Bebek Hotel","03 - Meio de Hospedagem","https://www.villabebek.com.br","","","","Aguardando","",""],
    ["SS0310","Villa Camboa","03 - Meio de Hospedagem","https://www.villacamboa.com.br","","","","Aguardando","",""],
    ["SS0315","Villa Paradiso","03 - Meio de Hospedagem","https://www.villaparadisopousada.com.br","","","","Aguardando","",""],
    ["SS0019","Chalés Ancoradouro","03 - Meio de Hospedagem","https://www.chalesancoradouro.com.br","","","","Aguardando","",""],
    ["SS0074","Ganesh Chalés","03 - Meio de Hospedagem","https://ganeshchales.com/","","","","Aguardando","",""],
    ["SS0075","Guaru-bora Pousada","03 - Meio de Hospedagem","https://www.pousadaguarubora.com.br","","","","Aguardando","",""],
    ["SS0077","HB Point Pousada","03 - Meio de Hospedagem","https://www.pousadahbpoint.com.br","","","","Aguardando","",""],
    ["SS0796","Restaurante La Trattoria","15 - Restaurante cafeteria bar e similares","https://www.villamare.com.br","","","","Aguardando","",""],
    ["SS0167","Pousada Boracéia Beach","03 - Meio de Hospedagem","https://www.pousadaboraceiabeach.com.br","","","","Aguardando","",""],
    ["SS0365","Raízes Aventura Ecoturismo","02 - Agência de Turismo","https://www.raizesaventura.com.br","","","","Aguardando","",""],
    ["SS0289","Salvetti Praia Hotel","03 - Meio de Hospedagem","https://www.salvettipraiahotel.com.br","","","","Aguardando","",""],
    ["SS0503","7Stay","11 - Prestador especializado em segmentos turístico","https://www.7stay.com.br","","","","Aguardando","",""],
    ["SS0799","Restaurante Ogan","15 - Restaurante cafeteria bar e similares","https://www.restauranteogan.com.br/","","","","Aguardando","",""],
    ["SS0015","Alma Surf Tour & House","03 - Meio de Hospedagem","https://www.almasurfhouse.com.br","","","","Aguardando","",""],
    ["SS0024","Associação Pousada Recanto do Barão","03 - Meio de Hospedagem","https://associacaopousadarecantodobarao.placeweb.site/","","","","Plataforma JS — Sem acesso estático","Site usa placeweb.site, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
    ["SS0025","Atena Praia Hotel","03 - Meio de Hospedagem","https://www.atenapraiahotel.com.br","","","","Aguardando","",""],
    ["SS0381","Basílio Despachante Náutico","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.escolabasilio.com.br","","","","Aguardando","",""],
    ["SS0007","Camping Porongaba","01 - Acampamento Turístico","https://www.porongaba.com.br","","","","Aguardando","",""],
    ["SS0043","Casa da Pedra","03 - Meio de Hospedagem","https://www.casadepedrapousada.com.br","","","","Aguardando","",""],
    ["SS0384","Marina Vitória","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.marinavitoria.com.br","","","","Aguardando","",""],
    ["SS0048","Chalés Água de Coco","03 - Meio de Hospedagem","https://www.pousadachalesaguadecoco.com.br","","","","Aguardando","",""],
    ["SS0385","Club Porto das Ilhas - Garagem Nautica","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.portodasilhas.com","","","","Aguardando","",""],
    ["SS0060","Dani Hotel","03 - Meio de Hospedagem","https://www.danihotel.com.br","","","","Aguardando","",""],
    ["SS0336","Defender Tour","02 - Agência de Turismo","https://www.defendertour.com.br","","","","Aguardando","",""],
    ["SS0340","Eco Experience","02 - Agência de Turismo","https://www.ecoexperience.com.br","","","","Aguardando","",""],
    ["SS0805","Restaurante Ravenala","15 - Restaurante cafeteria bar e similares","https://restauranteravenala.com.br","","","","Aguardando","",""],
    ["SS0088","Pousada Di Mari","03 - Meio de Hospedagem","https://www.dimari.com.br","","","","Aguardando","",""],
    ["SS0105","Hotel Pousada Cavalo Marinho","03 - Meio de Hospedagem","https://www.cavalomarinho.com.br","","","","Aguardando","",""],
    ["SS0131","Lugar Comum","03 - Meio de Hospedagem","https://www.pousadalugarcomum.com.br","","","","Aguardando","",""],
    ["SS0398","Marina Canto do Rio","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.marinacantodorio.com.br","","","","Aguardando","",""],
    ["SS0140","Mariuá","03 - Meio de Hospedagem","https://www.pousadamariua.com.br","","","","Aguardando","",""],
    ["SS0146","Morada Das Ilhas","03 - Meio de Hospedagem","https://www.moradadasilhas.com.br","","","","Aguardando","",""],
    ["SS0147","Morada dos Colibris","03 - Meio de Hospedagem","https://www.moradadoscolibris.com","","","","Aguardando","",""],
    ["SS0152","Pé na Areia","03 - Meio de Hospedagem","https://pousadapenaareia.net","","","","Aguardando","",""],
    ["SS0153","Pérola da Mata","03 - Meio de Hospedagem","https://www.peroladamata.com.br","","","","Aguardando","",""],
    ["SS0014","Pousada Águas do Mar","03 - Meio de Hospedagem","https://aguadomar.com.br","","","","Aguardando","",""],
    ["SS0158","Pousada Apoema","03 - Meio de Hospedagem","https://www.pousadaapoema.com.br/","","","","Aguardando","",""],
    ["SS0402","Pousada Boiçucanga","03 - Meio de Hospedagem","https://www.pousadaboicucanga.com.br","","","","Aguardando","",""],
    ["SS0175","Pousada Canto Verde","03 - Meio de Hospedagem","https://www.pousadacantoverde.com.br","","","","Aguardando","",""],
    ["SS0177","Pousada Casarão","03 - Meio de Hospedagem","https://www.casaraopousada.com.br","","","","Aguardando","",""],
    ["SS0181","Pousada Céu Azul","03 - Meio de Hospedagem","https://www.pousadaceuazul.com.br","","","","Aguardando","",""],
    ["SS0184","Pousada da Barra","03 - Meio de Hospedagem","https://www.pousadadabarraboicucanga.com","","","","Aguardando","",""],
    ["SS0213","Pousada Maori","03 - Meio de Hospedagem","https://www.pousadamaori.com.br","","","","Aguardando","",""],
    ["SS0233","Pousada Raios de Luar","03 - Meio de Hospedagem","https://www.pousadaraiosdeluar.com.br","","","","Aguardando","",""],
    ["SS0244","Pousada Tempo Rei","03 - Meio de Hospedagem","https://www.pousadatemporei.com.br","","","","Aguardando","",""],
    ["SS0811","Restaurante Tiê Sahy","15 - Restaurante cafeteria bar e similares","https://www.tiesahy.com.br","","","","Aguardando","",""],
    ["SS0948","Recanto das Margaridas","03 - Meio de Hospedagem","https://www.recantodasmargaridas.com","","","","Aguardando","",""],
    ["SS0284","Residencial Dolce Villa","03 - Meio de Hospedagem","https://www.pousadadolcevilla.com.br","","","","Aguardando","",""],
    ["SS0814","Rochinha Centro","15 - Restaurante cafeteria bar e similares","https://www.sorvetesrochinha.com.br","","","","Aguardando","",""],
    ["SS0818","Santo Gole Maresias","15 - Restaurante cafeteria bar e similares","https://www.santogolemaresias.com.br","","","","Aguardando","",""],
    ["SS0294","Summit Beach Hotel","03 - Meio de Hospedagem","https://www.summithotels.com.br","","","","Aguardando","",""],
    ["SS0300","Toca do Capitão","03 - Meio de Hospedagem","https://www.tocadocapitao.com.br","","","","Aguardando","",""],
    ["SS0372","Trilhando Litoral","02 - Agência de Turismo","https://www.trilhandolitoral.com.br","","","","Aguardando","",""],
    ["SS0471","VBoat","11 - Prestador especializado em segmentos turístico","https://www.vboat.com.br","","","","Aguardando","",""],
    ["SS0318","Villafranca Residencial","03 - Meio de Hospedagem","https://www.villafrancaresidencial.com.br/","","","","Aguardando","",""],
    ["SS0070","Estalagem do Píer","03 - Meio de Hospedagem","https://www.estalagemdopier.com.br","","","","Aguardando","",""],
    ["SS0098","Hotel e Restaurante Canoa","03 - Meio de Hospedagem","https://www.canoa.com.br","","","","Aguardando","",""],
    ["SS0396","Marina Boreste","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.boreste.com.br","","","","Aguardando","",""],
    ["SS0397","Agência Marina Canoa Barra do Una","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.canoa.com.br","","","","Aguardando","",""],
    ["SS0400","Marina Marinella","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.nauticamarinella.com.br","","","","Aguardando","",""],
    ["SS0825","Sino Pizzaria","15 - Restaurante cafeteria bar e similares","https://sinopizzaria.com.br","","","","Aguardando","",""],
    ["SS0240","Pousada Samburá","03 - Meio de Hospedagem","https://www.pousadasambura.com.br/","","","","Aguardando","",""],
    ["SS0251","Pousada Una","03 - Meio de Hospedagem","https://www.pousadauna.com.br","","","","Aguardando","",""],
    ["SS0827","Sorveteria Rocha","15 - Restaurante cafeteria bar e similares","https://www.sorveteriarocha.com.br","","","","Aguardando","",""],
    ["SS0828","Sorveteria Rocha II","15 - Restaurante cafeteria bar e similares","https://www.sorvetesrocha.com.br","","","","Aguardando","",""],
    ["SS0831","Sqina Bar Maresias","15 - Restaurante cafeteria bar e similares","https://squina.com.br","","","","Aguardando","",""],
    ["SS0064","Drifter Hostel","03 - Meio de Hospedagem","https://www.drifter.com.br","","","","Aguardando","",""],
    ["SS0090","Hotel Aldeia do Sahy","03 - Meio de Hospedagem","https://www.sahy.tur.br","","","","Aguardando","",""],
    ["SS0836","Subway São Sebastião","15 - Restaurante cafeteria bar e similares","https://www.subway.com","","","","Aguardando","",""],
    ["SS0160","Pousada Aroeira","03 - Meio de Hospedagem","https://www.pousadaaroeira.com.br","","","","Aguardando","",""],
    ["SS0170","Pousada Brigitte Bardot","03 - Meio de Hospedagem","https://www.brigittebardot.com.br","","","","Aguardando","",""],
    ["SS0224","Pousada Parcel das Ilhas","03 - Meio de Hospedagem","https://www.parceldasilhas.com.br","","","","Aguardando","",""],
    ["SS0239","Pousada Sahy da Terra","03 - Meio de Hospedagem","https://www.sahydaterra.com/","","","","Aguardando","",""],
    ["SS0246","Pousada Tie Sahy","03 - Meio de Hospedagem","https://www.tiesahy.com.br","","","","Aguardando","",""],
    ["SS0283","Refúgio Porto Sahy","03 - Meio de Hospedagem","https://www.pousadanosahy.com.br","","","","Aguardando","",""],
    ["SS0840","Taioba Gastronomia","15 - Restaurante cafeteria bar e similares","https://www.taiobagastronomia.com.br","","","","Aguardando","",""],
    ["SS0288","Sahydaterra Pousada","03 - Meio de Hospedagem","https://www.sahydaterra.com","","","","Aguardando","",""],
    ["SS0029","Barequeçaba Praia Hotel","03 - Meio de Hospedagem","https://www.barehotel.com.br","","","","Aguardando","",""],
    ["SS0035","Brisa do Mar Praia Hotel Barequeçaba","03 - Meio de Hospedagem","https://www.brisadomarbare.com.br","","","","Aguardando","",""],
    ["SS0102","Hotel Portal da Praia","03 - Meio de Hospedagem","https://www.hotelportaldapraia.com.br","","","","Aguardando","",""],
    ["SS0103","Hotel Portal de Barequeçaba","03 - Meio de Hospedagem","https://www.hotelportalbare.com.br","","","","Aguardando","",""],
    ["SS0476","Ilha Sol e Mar","06 - Transportadora Turística","https://ilhasolemarturismo.com.br/","","","","Aguardando","",""],
    ["SS0124","Kauano Pousada","03 - Meio de Hospedagem","https://www.kauano.com.br","","","","Aguardando","",""],
    ["SS0148","Nayane Pousada","03 - Meio de Hospedagem","https://www.pousadanayane.com.br","","","","Aguardando","",""],
    ["SS0955","Portugal na Mira","02 - Agência de Turismo","https://www.portugalnamira.com","","","","Aguardando","",""],
    ["SS0179","Pousada Castelinho","03 - Meio de Hospedagem","https://www.pousadacastelinho.org","","","","Aguardando","",""],
    ["SS0200","Pousada Estrela Mare","03 - Meio de Hospedagem","https://www.pousadaestrelamare.com","","","","Aguardando","",""],
    ["SS0236","Pousada Residencial Águas de Bare","03 - Meio de Hospedagem","https://www.aguadebare.com.br","","","","Aguardando","",""],
    ["SS0253","Pousada Victoria Reggia","03 - Meio de Hospedagem","https://www.pousadavictoriareggia.com","","","","Aguardando","",""],
    ["SS0256","Pousada Vila Barequeçaba","03 - Meio de Hospedagem","https://pousadavilabarequecaba.com.br/","","","","Aguardando","",""],
    ["SS0269","Pousada Viva Barê","03 - Meio de Hospedagem","https://pousadavivabare.com.br/","","","","Aguardando","",""],
    ["SS0273","Pris Hotel","03 - Meio de Hospedagem","https://www.prishotel.com.br","","","","Aguardando","",""],
    ["SS0303","Valentina Praia Hotel","03 - Meio de Hospedagem","https://www.valentinapraiahotel.com.br","","","","Aguardando","",""],
    ["SS0313","Villa Mare Residence","03 - Meio de Hospedagem","https://www.villamareresidence.com.br/","","","","Aguardando","",""],
    ["SS0319","Villagio Valentina","03 - Meio de Hospedagem","https://www.villagiovalentina.com.br","","","","Aguardando","",""],
    ["SS0320","Vista bela Resort","03 - Meio de Hospedagem","https://www.vistabela.com.br","","","","Aguardando","",""],
    ["SS0566","Vistabella Resort","12 - Prestador de Infraestrutura de apoio para eventos","https://www.vistabela.com.br","","","","Aguardando","",""],
    ["SS0163","Pousada Azul Maria","03 - Meio de Hospedagem","https://www.azulmaria.com.br","","","","Aguardando","",""],
    ["SS0229","Pousada Praia da Baleia","03 - Meio de Hospedagem","https://www.pousadapraiadabaleia.com.br/","","","","Aguardando","",""],
    ["SS0092","Hotel Arrastão","03 - Meio de Hospedagem","https://www.hotelarrastao.tur.br","","","","Aguardando","",""],
    ["SS0399","Marina Igararecê","08 - Empreendimento de Apoio ao Turismo Náutico ou à Pesca Desportiva","https://www.marinaigararece.com","","","","Aguardando","",""],
    ["SS0272","Praia e Montanha","03 - Meio de Hospedagem","https://www.praiaemontanha.com","","","","Aguardando","",""],
    ["SS0940","Adilson Imóveis","16 - Outros","https://www.adimov.com.br","","","","Aguardando","",""],
    ["SS0923","Décio Imóveis LTDA","16 - Outros","https://www.decioimoveis.com.br","","","","Aguardando","",""],
    ["SS0930","Docellar Imóveis","16 - Outros","https://www.docellarimoveis.com.br","","","","Aguardando","",""],
    ["SS0932","DPassos Imóveis Imobiliária","16 - Outros","https://www.dpassos.com.br","","","","Aguardando","",""],
    ["SS0935","Engimoveis Negócios Imobiliários","16 - Outros","https://www.engimoveis.com.br","","","","Aguardando","",""],
    ["SS0920","HTS Imóveis","16 - Outros","https://www.htsimoveis.com.br","","","","Aguardando","",""],
    ["SS0918","Imobiliária Padrão","16 - Outros","https://www.imobiliariapadrao.com.br","","","","Aguardando","",""],
    ["SS0929","Imóveis D'Paula","16 - Outros","https://www.imoveisdpaula.com.br","","","","Aguardando","",""],
    ["SS0438","Hotel La Serena","03 - Meio de Hospedagem","https://www.laserenahotal.com.br","","","","Aguardando","",""],
    ["SS0335","Pousada Aconchego Recanto Canto do mar","03 - Meio de Hospedagem","https://www.pousadarecantocantodomar.com.br","","","","Aguardando","",""],
    ["SS0970","Girlstrips Litoral","02 - Agência de Turismo","https://www.girltrips.com.br","","","","Aguardando","",""],
    ["SS0976","Pousada Kaimana","03 - Meio de Hospedagem","https://www.pousadakaimana.com.br","","","","Aguardando","",""],
    ["SS0982","BR Juréia Cama e Café","03 - Meio de Hospedagem","https://www.brjureiacamaecafe.com.br","","","","Aguardando","",""],
    ["SS0996","Suítes Recanto Sabiá Boiçucanga","03 - Meio de Hospedagem","https://www.instagram.com/suites.recantosabia.boicucanga","","","","Plataforma JS — Sem acesso estático","Site usa instagram.com, que bloqueia scraping ou requer JavaScript para renderizar preços.",""],
  ];
}
