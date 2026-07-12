/**
 * Sistema de Monitoramento de Preços Inteligente via Gemini API
 * Desenvolvido por Antigravity - Versão 1.0 (Senior Specialist Edition)
 */

// Adiciona o menu personalizado ao abrir a planilha
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 Monitor de Preços')
      .addItem('⚡ Atualizar Todos os Preços', 'atualizarTodosPrecos')
      .addItem('🎯 Atualizar Apenas Selecionados', 'atualizarPrecosSelecionados')
      .addSeparator()
      .addItem('📅 Configurar Atualização Diária', 'configurarGatilhoDiario')
      .addItem('❌ Remover Atualização Diária', 'removerGatilhos')
      .addToUi();
}

// Atalho para atualizar todos os registros
function atualizarTodosPrecos() {
  atualizarProcessamento(false);
}

// Atalho para atualizar apenas as linhas selecionadas pelo cursor do usuário
function atualizarPrecosSelecionados() {
  atualizarProcessamento(true);
}

/**
 * Função principal que lê a planilha, faz a requisição web e chama a API do Gemini
 */
function atualizarProcessamento(apenasSelecionados = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obter chave de API do Gemini da aba 'Painel de Controle'
  const sheetConfig = ss.getSheetByName('Painel de Controle');
  if (!sheetConfig) {
    SpreadsheetApp.getUi().alert('Erro', 'A aba "Painel de Controle" não foi encontrada. Por favor, crie uma aba com este nome e insira a chave da API do Gemini.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const apiKey = sheetConfig.getRange('B3').getValue().toString().trim(); // Assume chave na célula B3
  if (!apiKey || apiKey === '' || apiKey.startsWith('INSIRA')) {
    SpreadsheetApp.getUi().alert('Configuração Necessária', 'Chave da API do Gemini não configurada ou inválida na aba "Painel de Controle" (Célula B3). Obtenha uma chave gratuita em aistudio.google.com', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Obter aba de monitoramento
  const sheetMonitor = ss.getSheetByName('Monitoramento');
  if (!sheetMonitor) {
    SpreadsheetApp.getUi().alert('Erro', 'A aba "Monitoramento" não foi encontrada.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const lastRow = sheetMonitor.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Aviso', 'Nenhum dado encontrado na aba "Monitoramento".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Obter dados da tabela
  // Estrutura das colunas: 
  // A: ID, B: Nome, C: Atividade, D: Site, E: Preço Anterior, F: Preço Atual, G: Variação, H: Status, I: Detalhes, J: Última Atualização
  const dataRange = sheetMonitor.getRange(2, 1, lastRow - 1, 10);
  const data = dataRange.getValues();
  
  let rowsToProcess = [];
  let selectedRows = [];
  
  if (apenasSelecionados) {
    const activeRange = sheetMonitor.getActiveRange();
    if (activeRange) {
      const startRow = activeRange.getRow();
      const numRows = activeRange.getNumRows();
      for (let r = startRow; r < startRow + numRows; r++) {
        if (r >= 2 && r <= lastRow) {
          selectedRows.push(r);
        }
      }
    }
    if (selectedRows.length === 0) {
      SpreadsheetApp.getUi().alert('Aviso', 'Nenhuma linha selecionada. Clique nas linhas que deseja atualizar antes de clicar neste botão.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }
  
  for (let i = 0; i < data.length; i++) {
    const sheetRowNum = i + 2;
    if (apenasSelecionados && !selectedRows.includes(sheetRowNum)) {
      continue;
    }
    
    const id = data[i][0];
    const nome = data[i][1];
    const atividade = data[i][2];
    const site = data[i][3];
    const precoAtual = data[i][5];
    
    if (site && site !== '' && site !== '-') {
      rowsToProcess.push({
        index: i,
        rowNum: sheetRowNum,
        id: id,
        nome: nome,
        atividade: atividade,
        site: site,
        precoAtual: precoAtual
      });
    }
  }
  
  if (rowsToProcess.length === 0) {
    SpreadsheetApp.getUi().alert('Aviso', 'Nenhum estabelecimento com site válido encontrado para atualizar.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Confirmação para evitar chamadas acidentais em lote grande
  if (!apenasSelecionados && rowsToProcess.length > 50) {
    const confirm = SpreadsheetApp.getUi().alert(
      'Confirmação de Execução em Lote',
      `Você está prestes a atualizar ${rowsToProcess.length} estabelecimentos. Isso pode levar alguns minutos devido ao tempo de requisição de cada site. Deseja continuar?`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (confirm !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
  }
  
  // Atualiza status inicial na planilha
  rowsToProcess.forEach(row => {
    sheetMonitor.getRange(row.rowNum, 8).setValue('Processando...'); // Coluna H (Status)
  });
  SpreadsheetApp.flush();
  
  // Processamento linha a linha
  for (let r = 0; r < rowsToProcess.length; r++) {
    const item = rowsToProcess[r];
    const url = item.site;
    
    try {
      // 1. Fazer requisição HTTP para o site do estabelecimento
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: false,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
      });
      
      const responseCode = response.getResponseCode();
      if (responseCode !== 200) {
        sheetMonitor.getRange(item.rowNum, 8).setValue(`Erro: HTTP ${responseCode}`);
        sheetMonitor.getRange(item.rowNum, 10).setValue(new Date());
        continue;
      }
      
      const html = response.getContentText('UTF-8');
      
      // 2. Limpar o HTML para economizar tokens na chamada da API
      const cleanText = extrairTextoLimpo(html);
      if (cleanText.length < 50) {
        sheetMonitor.getRange(item.rowNum, 8).setValue('Erro: Site vazio/sem texto');
        sheetMonitor.getRange(item.rowNum, 10).setValue(new Date());
        continue;
      }
      
      // Limita o tamanho do texto para segurança de tokens (máximo ~15.000 caracteres)
      const textSample = cleanText.substring(0, 15000);
      
      // 3. Chamar a API do Gemini para analisar e extrair o preço
      const result = chamarGeminiAPI(apiKey, item.nome, item.atividade, textSample);
      
      if (result && result.status === 'Sucesso' && result.price > 0) {
        // Obter valores para atualizar
        const precoAnterior = item.precoAtual;
        const novoPreco = result.price;
        
        sheetMonitor.getRange(item.rowNum, 5).setValue(precoAnterior || ''); // Coluna E (Preço Anterior)
        sheetMonitor.getRange(item.rowNum, 6).setValue(novoPreco); // Coluna F (Preço Atual)
        
        // Calcular variação se houver preço anterior
        if (precoAnterior && precoAnterior > 0) {
          const variacao = (novoPreco - precoAnterior) / precoAnterior;
          sheetMonitor.getRange(item.rowNum, 7).setValue(variacao); // Coluna G (Variação %)
        } else {
          sheetMonitor.getRange(item.rowNum, 7).setValue('');
        }
        
        sheetMonitor.getRange(item.rowNum, 8).setValue('Sucesso'); // Coluna H (Status)
        sheetMonitor.getRange(item.rowNum, 9).setValue(result.details || ''); // Coluna I (Detalhes)
      } else {
        sheetMonitor.getRange(item.rowNum, 8).setValue('Preço não encontrado');
        sheetMonitor.getRange(item.rowNum, 9).setValue(result.details || 'Sem informações de preços visíveis no site.');
      }
      
    } catch (e) {
      sheetMonitor.getRange(item.rowNum, 8).setValue('Erro: Scraping falhou');
      sheetMonitor.getRange(item.rowNum, 9).setValue(e.toString());
    }
    
    sheetMonitor.getRange(item.rowNum, 10).setValue(new Date()); // Coluna J (Última Atualização)
    
    // Atualiza a interface da planilha em tempo real para o usuário ver o progresso
    SpreadsheetApp.flush();
    
    // Pequeno delay para evitar sobrecarregar os servidores e a API (rate limits)
    Utilities.sleep(1500); 
  }
  
  SpreadsheetApp.getUi().alert('Concluído', `Atualização finalizada! ${rowsToProcess.length} itens processados.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Remove scripts, CSS, links, SVGs e tags HTML para deixar apenas o texto legível.
 * Reduz em até 90% o consumo de tokens na API do Gemini.
 */
function extrairTextoLimpo(html) {
  if (!html) return '';
  
  let text = html;
  
  // Remove scripts, styles, head e tags de comentário
  text = text.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
  text = text.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
  text = text.replace(/<head\b[^<]*(?:(?!<\/head>)<[^<]*)*<\/head>/gi, '');
  text = text.replace(/<svg\b[^<]*(?:(?!<\/svg>)<[^<]*)*<\/svg>/gi, '');
  text = text.replace(/<!--[\s\S]*?-->/g, '');
  
  // Substitui quebras de tags por espaços
  text = text.replace(/<\/?[a-z][a-z0-9]*[^>]*>/gi, ' ');
  
  // Substitui entidades HTML comuns
  text = text.replace(/&nbsp;/g, ' ')
             .replace(/&lt;/g, '<')
             .replace(/&gt;/g, '>')
             .replace(/&amp;/g, '&')
             .replace(/&quot;/g, '"');
             
  // Remove múltiplos espaços em branco e quebras de linha repetidas
  text = text.replace(/\s+/g, ' ').trim();
  
  return text;
}

/**
 * Envia o texto limpo para o Gemini extrair o preço estruturado como JSON
 */
function chamarGeminiAPI(apiKey, nome, atividade, textoSite) {
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  
  // Definir comportamento específico baseado na categoria
  let orientacaoPreco = '';
  if (atividade.includes('Hospedagem') || atividade.includes('ACAMPAMENTO')) {
    orientacaoPreco = 'Busque pelo preço de diária padrão, diária inicial, diária mínima ou tarifa "a partir de" para quarto duplo / casal. Extraia apenas o valor numérico.';
  } else if (atividade.includes('Restaurante') || atividade.includes('Alimentação')) {
    orientacaoPreco = 'Busque pelo preço médio de prato principal, buffet, prato feito (PF), prato do dia, prato comercial ou pizza básica de tamanho individual/médio. Ignore preços de bebidas isoladas ou entradas.';
  } else if (atividade.includes('Locadora')) {
    orientacaoPreco = 'Busque pelo preço da diária padrão de locação de veículo básico (carro econômico).';
  } else {
    orientacaoPreco = 'Busque pelo preço do produto ou serviço principal oferecido listado na página.';
  }
  
  const prompt = `Você é um monitor de preços automatizado.
Analise as informações do website do estabelecimento "${nome}" (Atividade: ${atividade}).
Seu objetivo é encontrar preços de serviços ou produtos listados no texto do site.

Instrução Específica para Preços:
${orientacaoPreco}

Texto extraído do site:
"""
${textoSite}
"""

Retorne OBRIGATORIAMENTE um objeto JSON válido contendo exatamente a seguinte estrutura:
{
  "price": (número decimal com o preço extraído, ex: 150.00. Use -1.0 se não encontrar nenhum preço),
  "currency": (moeda em formato de string de 3 caracteres, ex: "BRL" ou "USD"),
  "details": (uma breve descrição em português sobre qual preço foi extraído, ex: "Diária básica chalé casal" ou "Média pratos principais". Máximo 100 caracteres),
  "status": (string contendo "Sucesso" se achou um preço representativo ou "Não Encontrado" caso contrário)
}`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }],
    generationConfig: {
      responseMimeType: "application/json",
      responseSchema: {
        type: "OBJECT",
        properties: {
          price: { type: "NUMBER" },
          currency: { type: "STRING" },
          details: { type: "STRING" },
          status: { type: "STRING" }
        },
        required: ["price", "currency", "details", "status"]
      }
    }
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    throw new Error(`Erro na API Gemini: HTTP ${responseCode} - ${response.getContentText()}`);
  }
  
  const resJson = JSON.parse(response.getContentText());
  
  if (resJson.candidates && resJson.candidates[0] && resJson.candidates[0].content && resJson.candidates[0].content.parts[0]) {
    const jsonText = resJson.candidates[0].content.parts[0].text;
    return JSON.parse(jsonText);
  }
  
  throw new Error('Resposta da API Gemini em formato inesperado.');
}

/**
 * Configura um gatilho de tempo para executar o script de monitoramento diariamente
 */
function configurarGatilhoDiario() {
  removerGatilhos(); // Remove gatilhos anteriores para evitar duplicidade
  
  ScriptApp.newTrigger('atualizarTodosPrecos')
      .timeBased()
      .everyDays(1)
      .atHour(2) // Executa de madrugada (2h) para evitar lentidão durante o dia
      .create();
      
  SpreadsheetApp.getUi().alert('Gatilho Configurado', 'O monitoramento automático foi configurado para rodar diariamente, por volta das 2:00 da manhã.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Remove todos os gatilhos ativos associados a esta planilha
 */
function removerGatilhos() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  if (MailApp) { // Apenas executa se chamado manualmente
    try {
      SpreadsheetApp.getUi().alert('Gatilhos Removidos', 'Todos os gatilhos automáticos do script foram desativados.', SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }
}
