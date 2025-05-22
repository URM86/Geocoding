/**
 * Google Sheets Geocoding Script com Processamento em Lotes Automatizado
 * -----------------------------------------------------
 * Este script permite geocodificar grandes quantidades de endereços ou coordenadas
 * em uma planilha Google Sheets, superando as limitações de tempo de execução
 * através de um sistema de processamento em lotes totalmente automatizado.
 * 
 * Autor original: nuket (https://github.com/nuket/google-sheets-geocoding-macro)
 * Modificações para processamento em lotes: Ulises Rodrigo Magdalena
 * Otimizações e automação completa: Manus AI
 * Data: 22/05/2025
 * 
 * COMO USAR:
 * 1. Selecione 3 colunas na planilha (Endereço, Latitude, Longitude)
 * 2. Use o menu "Geocode" para escolher a função desejada
 * 3. O processamento continuará automaticamente até o final, sem necessidade de intervenção
 */

/**
 * CONFIGURAÇÕES REGIONAIS
 * ----------------------
 * Define a região prioritária para a geocodificação (ex: 'us', 'br', 'uk')
 * Isso ajuda a melhorar a precisão das geocodificações em cada país
 */
function getGeocodingRegion() {
  return PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'br';
}

/**
 * CONFIGURAÇÕES DO PROCESSAMENTO EM LOTES
 * --------------------------------------
 * Estas constantes controlam o comportamento do processamento em lotes
 */

// Número de linhas processadas em cada execução
// Mantido em 50 para respeitar os limites da API do Google
const BATCH_SIZE = 50; 

// Intervalo em milissegundos entre o processamento de lotes consecutivos
// Aumentado para reduzir o risco de bloqueio pela API do Google
const PAUSE_BETWEEN_BATCH = 2000; 

// Intervalo em milissegundos entre solicitações individuais de geocodificação
// Ajustado para equilibrar velocidade e segurança
const PAUSE_BETWEEN_REQUESTS = 250;

// Número máximo de tentativas para uma geocodificação
const MAX_RETRIES = 3;

// Intervalo em milissegundos para esperar após um erro OVER_QUERY_LIMIT
const OVER_LIMIT_PAUSE = 5000;

// Chave para armazenar o estado do processamento
const STATE_KEY = 'GEOCODING_STATE';

/**
 * ESTRUTURA DE ESTADO DO PROCESSAMENTO
 * ----------------------------------
 * Armazena informações sobre o estado atual do processamento em lotes
 * para permitir retomada automática e rastreamento de progresso
 */
function getProcessingState() {
  const stateJson = PropertiesService.getDocumentProperties().getProperty(STATE_KEY);
  if (!stateJson) {
    return {
      currentRow: 1,
      totalRows: 0,
      sheetId: '',
      rangeA1: '',
      mode: '',
      errors: 0,
      processed: 0,
      startTime: new Date().getTime()
    };
  }
  return JSON.parse(stateJson);
}

function saveProcessingState(state) {
  PropertiesService.getDocumentProperties().setProperty(STATE_KEY, JSON.stringify(state));
}

function clearProcessingState() {
  PropertiesService.getDocumentProperties().deleteProperty(STATE_KEY);
  // Limpa também quaisquer triggers pendentes para evitar execuções duplicadas
  clearAllTriggers();
}

/**
 * GERENCIAMENTO DE TRIGGERS
 * -----------------------
 * Funções para gerenciar os triggers que permitem a execução automática em lotes
 */
function clearAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'addressToPositionBatch' || 
        triggers[i].getHandlerFunction() === 'positionToAddressBatch' ||
        triggers[i].getHandlerFunction() === 'continueBatchProcessing') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function scheduleContinuation(functionName) {
  // Limpa triggers existentes para evitar execuções duplicadas
  clearAllTriggers();
  
  // Agenda a próxima execução
  ScriptApp.newTrigger('continueBatchProcessing')
    .timeBased()
    .after(PAUSE_BETWEEN_BATCH)
    .create();
}

/**
 * FUNÇÃO DE CONTINUAÇÃO UNIFICADA
 * -----------------------------
 * Esta função centraliza a lógica de continuação do processamento em lotes,
 * independente do tipo de geocodificação (endereço->posição ou posição->endereço)
 */
function continueBatchProcessing() {
  const state = getProcessingState();
  
  // Verifica se há um processamento em andamento
  if (!state.mode) {
    Logger.log('Nenhum processamento em andamento.');
    clearProcessingState();
    return;
  }
  
  // Continua o processamento de acordo com o modo
  if (state.mode === 'addressToPosition') {
    processAddressToPositionBatch();
  } else if (state.mode === 'positionToAddress') {
    processPositionToAddressBatch();
  }
}

/**
 * GEOCODIFICAÇÃO DE ENDEREÇOS PARA COORDENADAS (EM LOTES AUTOMATIZADOS)
 * ------------------------------------------------------------------
 * Função de entrada para iniciar o processamento de endereços para coordenadas
 */
function addressToPositionBatch() {
  // Limpa qualquer estado anterior e triggers pendentes
  clearProcessingState();
  
  // Obtém a planilha ativa e o intervalo selecionado
  const sheet = SpreadsheetApp.getActiveSheet();
  const cells = sheet.getActiveRange();
  
  // Verifica se a seleção tem pelo menos 3 colunas (Endereço, Lat, Lng)
  if (cells.getNumColumns() != 3) {
    SpreadsheetApp.getUi().alert("Selecione 3 colunas: Endereço, Latitude, Longitude");
    return;
  }
  
  // Inicializa o estado do processamento
  const state = getProcessingState();
  state.currentRow = 1;
  state.totalRows = cells.getNumRows();
  state.sheetId = sheet.getSheetId();
  state.rangeA1 = cells.getA1Notation();
  state.mode = 'addressToPosition';
  state.startTime = new Date().getTime();
  saveProcessingState(state);
  
  // Cria uma célula de status para o usuário acompanhar o progresso
  const statusCell = sheet.getRange(1, cells.getNumColumns() + 4);
  statusCell.setValue("Iniciando processamento automatizado de " + state.totalRows + " linhas...");
  
  // Inicia o processamento do primeiro lote
  processAddressToPositionBatch();
}

/**
 * Processa um lote de endereços para coordenadas
 * Esta função é chamada repetidamente até que todos os lotes sejam processados
 */
function processAddressToPositionBatch() {
  // Recupera o estado atual do processamento
  const state = getProcessingState();
  if (!state.mode || state.mode !== 'addressToPosition') {
    Logger.log('Nenhum processamento de endereço para posição em andamento.');
    return;
  }
  
  // Obtém a planilha e o intervalo com base no estado salvo
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(s => s.getSheetId() === state.sheetId);
  if (!sheet) {
    Logger.log('Planilha não encontrada.');
    clearProcessingState();
    return;
  }
  
  const cells = sheet.getRange(state.rangeA1);
  
  // Define as colunas para cada tipo de dado
  const addressColumn = 1;  // Primeira coluna da seleção: endereço
  const latColumn = addressColumn + 1;  // Segunda coluna: latitude
  const lngColumn = addressColumn + 2;  // Terceira coluna: longitude
  
  // Calcula o fim do lote atual (não excede o total de linhas)
  const startRow = state.currentRow;
  const endRow = Math.min(startRow + BATCH_SIZE - 1, state.totalRows);
  
  // Cria uma célula de status para o usuário acompanhar o progresso
  const statusCell = sheet.getRange(1, cells.getNumColumns() + 4);
  statusCell.setValue("Processando lote: linhas " + startRow + " a " + endRow + 
                     " de " + state.totalRows + 
                     " (" + Math.round((startRow-1)/state.totalRows*100) + "% concluído)");
  
  // Inicializa o geocodificador com a região configurada
  const geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  
  // Processa cada linha do lote atual
  for (let row = startRow; row <= endRow; row++) {
    // Obtém o endereço da célula atual
    const address = cells.getCell(row, addressColumn).getValue();
    
    // Pula linhas com endereço vazio
    if (!address) {
      state.processed++;
      continue;
    }
    
    // Tenta geocodificar com sistema de retry
    let success = false;
    let retries = 0;
    let location;
    
    while (!success && retries < MAX_RETRIES) {
      try {
        // Tenta geocodificar o endereço
        location = geocoder.geocode(address);
        
        // Verifica se atingiu o limite de consultas
        if (location.status === 'OVER_QUERY_LIMIT') {
          Logger.log('Limite de consultas atingido. Aguardando...');
          Utilities.sleep(OVER_LIMIT_PAUSE);
          retries++;
          continue;
        }
        
        // Processa o resultado se bem-sucedido
        if (location.status === 'OK') {
          // Extrai as coordenadas do primeiro resultado
          const lat = location["results"][0]["geometry"]["location"]["lat"];
          const lng = location["results"][0]["geometry"]["location"]["lng"];
          
          // Preenche as células com as coordenadas
          cells.getCell(row, latColumn).setValue(lat);
          cells.getCell(row, lngColumn).setValue(lng);
          success = true;
        } else {
          // Registra o erro quando a geocodificação falha
          cells.getCell(row, latColumn).setValue("ERRO");
          cells.getCell(row, lngColumn).setValue(location.status);
          success = true; // Considera processado mesmo com erro
          state.errors++;
        }
      } catch (e) {
        // Captura quaisquer erros inesperados
        retries++;
        if (retries >= MAX_RETRIES) {
          cells.getCell(row, latColumn).setValue("ERRO");
          cells.getCell(row, lngColumn).setValue(e.toString());
          state.errors++;
        }
        Utilities.sleep(PAUSE_BETWEEN_REQUESTS * 2); // Pausa maior após erro
      }
    }
    
    state.processed++;
    
    // Adiciona uma pausa entre solicitações para evitar limites de taxa da API
    // Usa uma pausa variável para parecer mais com comportamento humano
    const randomPause = PAUSE_BETWEEN_REQUESTS * (0.8 + Math.random() * 0.4);
    Utilities.sleep(randomPause);
  }
  
  // Atualiza o estado para o próximo lote
  state.currentRow = endRow + 1;
  saveProcessingState(state);
  
  // Verifica se ainda há mais linhas para processar
  if (endRow < state.totalRows) {
    // Agenda automaticamente a próxima execução
    scheduleContinuation('addressToPositionBatch');
    
    // Atualiza a célula de status com informações sobre o próximo lote
    const percentComplete = Math.round(endRow/state.totalRows*100);
    statusCell.setValue("Processado lote: linhas " + startRow + " a " + endRow + 
                       " de " + state.totalRows + 
                       " (" + percentComplete + "% concluído). Próximo lote em breve.");
  } else {
    // Quando todas as linhas forem processadas, limpa o estado e finaliza
    const elapsedTime = Math.round((new Date().getTime() - state.startTime) / 1000);
    statusCell.setValue("Processamento concluído: " + state.totalRows + 
                       " linhas processadas em " + elapsedTime + 
                       " segundos. Erros: " + state.errors);
    clearProcessingState();
  }
}

/**
 * GEOCODIFICAÇÃO DE COORDENADAS PARA ENDEREÇOS (EM LOTES AUTOMATIZADOS)
 * ------------------------------------------------------------------
 * Função de entrada para iniciar o processamento de coordenadas para endereços
 */
function positionToAddressBatch() {
  // Limpa qualquer estado anterior e triggers pendentes
  clearProcessingState();
  
  // Obtém a planilha ativa e o intervalo selecionado
  const sheet = SpreadsheetApp.getActiveSheet();
  const cells = sheet.getActiveRange();
  
  // Verifica se a seleção tem pelo menos 3 colunas (Endereço, Lat, Lng)
  if (cells.getNumColumns() != 3) {
    SpreadsheetApp.getUi().alert("Selecione 3 colunas: Endereço, Latitude, Longitude");
    return;
  }
  
  // Inicializa o estado do processamento
  const state = getProcessingState();
  state.currentRow = 1;
  state.totalRows = cells.getNumRows();
  state.sheetId = sheet.getSheetId();
  state.rangeA1 = cells.getA1Notation();
  state.mode = 'positionToAddress';
  state.startTime = new Date().getTime();
  saveProcessingState(state);
  
  // Cria uma célula de status para o usuário acompanhar o progresso
  const statusCell = sheet.getRange(1, cells.getNumColumns() + 4);
  statusCell.setValue("Iniciando processamento automatizado de " + state.totalRows + " linhas...");
  
  // Inicia o processamento do primeiro lote
  processPositionToAddressBatch();
}

/**
 * Processa um lote de coordenadas para endereços
 * Esta função é chamada repetidamente até que todos os lotes sejam processados
 */
function processPositionToAddressBatch() {
  // Recupera o estado atual do processamento
  const state = getProcessingState();
  if (!state.mode || state.mode !== 'positionToAddress') {
    Logger.log('Nenhum processamento de posição para endereço em andamento.');
    return;
  }
  
  // Obtém a planilha e o intervalo com base no estado salvo
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(s => s.getSheetId() === state.sheetId);
  if (!sheet) {
    Logger.log('Planilha não encontrada.');
    clearProcessingState();
    return;
  }
  
  const cells = sheet.getRange(state.rangeA1);
  
  // Define as colunas para cada tipo de dado
  const addressColumn = 1;  // Primeira coluna da seleção: endereço (será preenchida)
  const latColumn = addressColumn + 1;  // Segunda coluna: latitude (entrada)
  const lngColumn = addressColumn + 2;  // Terceira coluna: longitude (entrada)
  
  // Calcula o fim do lote atual (não excede o total de linhas)
  const startRow = state.currentRow;
  const endRow = Math.min(startRow + BATCH_SIZE - 1, state.totalRows);
  
  // Cria uma célula de status para o usuário acompanhar o progresso
  const statusCell = sheet.getRange(1, cells.getNumColumns() + 4);
  statusCell.setValue("Processando lote: linhas " + startRow + " a " + endRow + 
                     " de " + state.totalRows + 
                     " (" + Math.round((startRow-1)/state.totalRows*100) + "% concluído)");
  
  // Inicializa o geocodificador com a região configurada
  const geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  
  // Processa cada linha do lote atual
  for (let row = startRow; row <= endRow; row++) {
    // Obtém as coordenadas da linha atual
    const lat = cells.getCell(row, latColumn).getValue();
    const lng = cells.getCell(row, lngColumn).getValue();
    
    // Pula linhas com coordenadas vazias
    if (!lat || !lng) {
      state.processed++;
      continue;
    }
    
    // Tenta geocodificar com sistema de retry
    let success = false;
    let retries = 0;
    let location;
    
    while (!success && retries < MAX_RETRIES) {
      try {
        // Realiza a geocodificação reversa (coordenadas para endereço)
        location = geocoder.reverseGeocode(lat, lng);
        
        // Verifica se atingiu o limite de consultas
        if (location.status === 'OVER_QUERY_LIMIT') {
          Logger.log('Limite de consultas atingido. Aguardando...');
          Utilities.sleep(OVER_LIMIT_PAUSE);
          retries++;
          continue;
        }
        
        // Processa o resultado se bem-sucedido
        if (location.status === 'OK') {
          // Obtém o endereço formatado do primeiro resultado
          const address = location["results"][0]["formatted_address"];
          cells.getCell(row, addressColumn).setValue(address);
          success = true;
        } else {
          // Registra o erro quando a geocodificação falha
          cells.getCell(row, addressColumn).setValue("ERRO: " + location.status);
          success = true; // Considera processado mesmo com erro
          state.errors++;
        }
      } catch (e) {
        // Captura quaisquer erros inesperados
        retries++;
        if (retries >= MAX_RETRIES) {
          cells.getCell(row, addressColumn).setValue("ERRO: " + e.toString());
          state.errors++;
        }
        Utilities.sleep(PAUSE_BETWEEN_REQUESTS * 2); // Pausa maior após erro
      }
    }
    
    state.processed++;
    
    // Adiciona uma pausa entre solicitações para evitar limites de taxa da API
    // Usa uma pausa variável para parecer mais com comportamento humano
    const randomPause = PAUSE_BETWEEN_REQUESTS * (0.8 + Math.random() * 0.4);
    Utilities.sleep(randomPause);
  }
  
  // Atualiza o estado para o próximo lote
  state.currentRow = endRow + 1;
  saveProcessingState(state);
  
  // Verifica se ainda há mais linhas para processar
  if (endRow < state.totalRows) {
    // Agenda automaticamente a próxima execução
    scheduleContinuation('positionToAddressBatch');
    
    // Atualiza a célula de status com informações sobre o próximo lote
    const percentComplete = Math.round(endRow/state.totalRows*100);
    statusCell.setValue("Processado lote: linhas " + startRow + " a " + endRow + 
                       " de " + state.totalRows + 
                       " (" + percentComplete + "% concluído). Próximo lote em breve.");
  } else {
    // Quando todas as linhas forem processadas, limpa o estado e finaliza
    const elapsedTime = Math.round((new Date().getTime() - state.startTime) / 1000);
    statusCell.setValue("Processamento concluído: " + state.totalRows + 
                       " linhas processadas em " + elapsedTime + 
                       " segundos. Erros: " + state.errors);
    clearProcessingState();
  }
}

/**
 * FUNCIONALIDADE DE RESET
 * ---------------------
 * Esta função permite reiniciar o estado do processamento em lotes
 * caso ocorra algum problema ou o usuário deseje recomeçar.
 */
function resetBatchProcessing() {
  clearProcessingState();
  SpreadsheetApp.getUi().alert("Status de processamento em lotes redefinido e triggers cancelados.");
}

/**
 * CONFIGURAÇÃO DE REGIÃO
 * --------------------
 * Esta função permite ao usuário configurar a região de geocodificação
 */
function setGeocodingRegion() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Configurar Região de Geocodificação',
    'Digite o código de país de 2 letras (ex: br, us, uk):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const region = response.getResponseText().trim().toLowerCase();
    if (region && region.length === 2) {
      PropertiesService.getDocumentProperties().setProperty('GEOCODING_REGION', region);
      ui.alert('Região configurada para: ' + region);
    } else {
      ui.alert('Código de região inválido. Use um código de 2 letras (ex: br, us, uk).');
    }
  }
}

/**
 * FUNÇÕES ORIGINAIS (MANTIDAS PARA COMPATIBILIDADE)
 * -----------------------------------------------
 * Estas são as funções da implementação original, mantidas para
 * compatibilidade e para processar pequenas quantidades de dados.
 */

/**
 * Geocodifica endereços para coordenadas (versão original)
 * Adequada apenas para pequenas quantidades de dados.
 */
function addressToPosition() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  // Deve ter selecionado 3 colunas (Endereço, Lat, Lng).
  if (cells.getNumColumns() != 3) {
    SpreadsheetApp.getUi().alert("Selecione 3 colunas: Endereço, Latitude, Longitude");
    return;
  }
  
  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var address = cells.getCell(addressRow, addressColumn).getValue();
    
    // Geocodifica o endereço e insere o par lat, lng nos
    // 2º e 3º elementos da linha atual.
    location = geocoder.geocode(address);
   
    // Só altera as células se o geocodificador parecer ter obtido uma
    // resposta válida.
    if (location.status == 'OK') {
      lat = location["results"][0]["geometry"]["location"]["lat"];
      lng = location["results"][0]["geometry"]["location"]["lng"];
      
      cells.getCell(addressRow, latColumn).setValue(lat);
      cells.getCell(addressRow, lngColumn).setValue(lng);
    }
  }
}

/**
 * Geocodifica coordenadas para endereços (versão original)
 * Adequada apenas para pequenas quantidades de dados.
 */
function positionToAddress() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  // Deve ter selecionado 3 colunas (Endereço, Lat, Lng).
  if (cells.getNumColumns() != 3) {
    SpreadsheetApp.getUi().alert("Selecione 3 colunas: Endereço, Latitude, Longitude");
    return;
  }

  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var lat = cells.getCell(addressRow, latColumn).getValue();
    var lng = cells.getCell(addressRow, lngColumn).getValue();
    
    // Geocodifica o par lat, lng para um endereço.
    location = geocoder.reverseGeocode(lat, lng);
   
    // Só altera as células se o geocodificador parecer ter obtido uma
    // resposta válida.
    if (location.status == 'OK') {
      var address = location["results"][0]["formatted_address"];
      cells.getCell(addressRow, addressColumn).setValue(address);
    }
  }  
}

/**
 * GERAÇÃO DE MENU E INICIALIZAÇÃO
 * -----------------------------
 * Estas funções criam o menu na interface do usuário e
 * configuram a aplicação quando a planilha é aberta.
 */

/**
 * Gera os itens do menu
 * A versão ampliada inclui as novas funções de processamento em lotes
 */
function generateMenu() {
  var entries = [
    {
      name: "⚡ Geocodificar Endereços -> Lat, Long (Automático)",
      functionName: "addressToPositionBatch"
    },
    {
      name: "⚡ Geocodificar Lat, Long -> Endereços (Automático)",
      functionName: "positionToAddressBatch"
    },
    {
      name: "Configurar Região de Geocodificação",
      functionName: "setGeocodingRegion"
    },
    {
      name: "Redefinir Status do Processamento",
      functionName: "resetBatchProcessing"
    },
    {
      name: "Geocodificar Endereços -> Lat, Long (Original)",
      functionName: "addressToPosition"
    },
    {
      name: "Geocodificar Lat, Long -> Endereços (Original)",
      functionName: "positionToAddress"
    }
  ];
  
  return entries;
}

/**
 * Função executada automaticamente quando a planilha é aberta
 * Adiciona o menu customizado à interface
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Geocode', generateMenu());
}
