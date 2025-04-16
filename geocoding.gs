/**
 * Google Sheets Geocoding Script com Processamento em Lotes
 * -----------------------------------------------------
 * Este script permite geocodificar grandes quantidades de endereços ou coordenadas
 * em uma planilha Google Sheets, superando as limitações de tempo de execução
 * através de um sistema de processamento em lotes.
 * 
 * Autor original: nuket (https://github.com/nuket/google-sheets-geocoding-macro)
 * Modificações para processamento em lotes: Ulises Rodrigo Magdalena
 * Data: 16/04/2025
 * 
 * COMO USAR:
 * 1. Selecione 3 colunas na planilha (Endereço, Latitude, Longitude)
 * 2. Use o menu "Geocode" para escolher a função desejada
 * 3. Para grandes volumes, use as opções "Em Lotes"
 */

/**
 * CONFIGURAÇÕES REGIONAIS
 * ----------------------
 * Define a região prioritária para a geocodificação (ex: 'us', 'br', 'uk')
 * Isso ajuda a melhorar a precisão das geocodificações em cada país
 */
function getGeocodingRegion() {
  // Retorna a região configurada ou 'us' como padrão
  return PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'us';
}

/**
 * CONFIGURAÇÕES DO PROCESSAMENTO EM LOTES
 * --------------------------------------
 * Estas constantes controlam o comportamento do processamento em lotes
 */

// Número de linhas processadas em cada execução
// NOTA: Um valor maior processa mais linhas por vez, mas aumenta o risco de atingir
// o limite de tempo de execução do Google Apps Script (6 minutos)
const BATCH_SIZE = 50; 

// Intervalo em milissegundos entre o processamento de lotes consecutivos
// NOTA: Este valor pode ser ajustado para controlar a velocidade do processamento
// Um valor maior reduz o risco de atingir limites de taxa de API
const PAUSE_BETWEEN_BATCH = 1000; 

// Intervalo em milissegundos entre solicitações individuais de geocodificação
// NOTA: Este valor ajuda a evitar erros de OVER_QUERY_LIMIT da API do Google Maps
const PAUSE_BETWEEN_REQUESTS = 200;

/**
 * GEOCODIFICAÇÃO DE ENDEREÇOS PARA COORDENADAS (EM LOTES)
 * ------------------------------------------------------
 * Esta função processa endereços e obtém suas coordenadas geográficas (lat/lng)
 * de forma escalonável, dividindo o trabalho em lotes menores.
 * 
 * MELHORIAS EM RELAÇÃO À VERSÃO ORIGINAL:
 * 1. Processamento em lotes para superar o limite de tempo de execução de 6 minutos
 * 2. Sistema de retomada que permite continuar de onde parou
 * 3. Feedback visual do progresso na planilha
 * 4. Tratamento de erros mais robusto
 * 5. Pausas controladas para evitar limites de taxa da API
 */
function addressToPositionBatch() {
  // Obtém a planilha ativa e o intervalo selecionado
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  // Verifica se a seleção tem pelo menos 3 colunas (Endereço, Lat, Lng)
  if (cells.getNumColumns() != 3) {
    SpreadsheetApp.getUi().alert("Selecione 3 colunas: Endereço, Latitude, Longitude");
    return;
  }
  
  // Define as colunas para cada tipo de dado
  var addressColumn = 1;  // Primeira coluna da seleção: endereço
  var latColumn = addressColumn + 1;  // Segunda coluna: latitude
  var lngColumn = addressColumn + 2;  // Terceira coluna: longitude
  var totalRows = cells.getNumRows();  // Total de linhas a processar
  
  // Recupera o ponto de retomada se houver um processamento anterior
  // Isso permite continuar de onde parou caso o script seja interrompido
  var startRow = PropertiesService.getDocumentProperties().getProperty('BATCH_CURRENT_ROW') || 1;
  startRow = parseInt(startRow);
  
  var ui = SpreadsheetApp.getUi();
  
  // Se existir um processamento anterior, pergunta ao usuário se deseja continuar
  if (startRow > 1) {
    var response = ui.alert(
      'Continuar processamento',
      'Foi encontrado um processamento anterior na linha ' + startRow + '. Deseja continuar de onde parou?',
      ui.ButtonSet.YES_NO
    );
    
    // Se o usuário optar por não continuar, reinicia do começo
    if (response == ui.Button.NO) {
      startRow = 1;
    }
  }
  
  // Calcula o fim do lote atual (não excede o total de linhas)
  var endRow = Math.min(startRow + BATCH_SIZE - 1, totalRows);
  
  // Cria uma célula de status para o usuário acompanhar o progresso
  // Esta célula é posicionada à direita da área selecionada
  var statusCell = sheet.getRange(1, cells.getNumColumns() + 4);
  statusCell.setValue("Processando lote: linhas " + startRow + " a " + endRow + " de " + totalRows);
  
  // Inicializa o geocodificador com a região configurada
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  
  // Processa cada linha do lote atual
  for (var row = startRow; row <= endRow; row++) {
    // Obtém o endereço da célula atual
    var address = cells.getCell(row, addressColumn).getValue();
    
    // Pula linhas com endereço vazio
    if (!address) continue;
    
    try {
      // Tenta geocodificar o endereço
      var location = geocoder.geocode(address);
      
      // Processa o resultado se bem-sucedido
      if (location.status == 'OK') {
        // Extrai as coordenadas do primeiro resultado
        var lat = location["results"][0]["geometry"]["location"]["lat"];
        var lng = location["results"][0]["geometry"]["location"]["lng"];
        
        // Preenche as células com as coordenadas
        cells.getCell(row, latColumn).setValue(lat);
        cells.getCell(row, lngColumn).setValue(lng);
      } else {
        // Registra o erro quando a geocodificação falha
        cells.getCell(row, latColumn).setValue("ERRO");
        cells.getCell(row, lngColumn).setValue(location.status);
      }
      
      // Adiciona uma pausa entre solicitações para evitar limites de taxa da API
      Utilities.sleep(PAUSE_BETWEEN_REQUESTS);
    } catch (e) {
      // Captura quaisquer erros inesperados
      cells.getCell(row, latColumn).setValue("ERRO");
      cells.getCell(row, lngColumn).setValue(e.toString());
    }
  }
  
  // Verifica se ainda há mais linhas para processar
  if (endRow < totalRows) {
    // Salva a próxima linha a ser processada para uso futuro
    PropertiesService.getDocumentProperties().setProperty('BATCH_CURRENT_ROW', endRow + 1);
    
    // Agenda automaticamente a próxima execução após o intervalo definido
    // Esta é a parte crucial que permite processar grandes volumes sem atingir limites de tempo
    ScriptApp.newTrigger('addressToPositionBatch')
      .timeBased()
      .after(PAUSE_BETWEEN_BATCH)
      .create();
      
    // Atualiza a célula de status com informações sobre o próximo lote
    statusCell.setValue("Processado lote: linhas " + startRow + " a " + endRow + " de " + totalRows + ". Próximo lote em breve.");
  } else {
    // Quando todas as linhas forem processadas, limpa o estado e finaliza
    PropertiesService.getDocumentProperties().deleteProperty('BATCH_CURRENT_ROW');
    statusCell.setValue("Processamento concluído: " + totalRows + " linhas processadas.");
  }
}

/**
 * GEOCODIFICAÇÃO DE COORDENADAS PARA ENDEREÇOS (EM LOTES)
 * ------------------------------------------------------
 * Esta função realiza o processo inverso: a partir de coordenadas (lat/lng),
 * obtém os endereços correspondentes, também utilizando o processamento em lotes.
 * 
 * A estrutura e a lógica são semelhantes à função addressToPositionBatch,
 * mas com o fluxo de dados invertido.
 */
function positionToAddressBatch() {
  // Obtém a planilha ativa e o intervalo selecionado
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  // Verifica se a seleção tem pelo menos 3 colunas (Endereço, Lat, Lng)
  if (cells.getNumColumns() != 3) {
    SpreadsheetApp.getUi().alert("Selecione 3 colunas: Endereço, Latitude, Longitude");
    return;
  }
  
  // Define as colunas para cada tipo de dado
  var addressColumn = 1;  // Primeira coluna da seleção: endereço (será preenchida)
  var latColumn = addressColumn + 1;  // Segunda coluna: latitude (entrada)
  var lngColumn = addressColumn + 2;  // Terceira coluna: longitude (entrada)
  var totalRows = cells.getNumRows();  // Total de linhas a processar
  
  // Recupera o ponto de retomada se houver um processamento anterior
  var startRow = PropertiesService.getDocumentProperties().getProperty('BATCH_CURRENT_ROW') || 1;
  startRow = parseInt(startRow);
  
  var ui = SpreadsheetApp.getUi();
  
  // Se existir um processamento anterior, pergunta ao usuário se deseja continuar
  if (startRow > 1) {
    var response = ui.alert(
      'Continuar processamento',
      'Foi encontrado um processamento anterior na linha ' + startRow + '. Deseja continuar de onde parou?',
      ui.ButtonSet.YES_NO
    );
    
    // Se o usuário optar por não continuar, reinicia do começo
    if (response == ui.Button.NO) {
      startRow = 1;
    }
  }
  
  // Calcula o fim do lote atual (não excede o total de linhas)
  var endRow = Math.min(startRow + BATCH_SIZE - 1, totalRows);
  
  // Cria uma célula de status para o usuário acompanhar o progresso
  var statusCell = sheet.getRange(1, cells.getNumColumns() + 4);
  statusCell.setValue("Processando lote: linhas " + startRow + " a " + endRow + " de " + totalRows);
  
  // Inicializa o geocodificador com a região configurada
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  
  // Processa cada linha do lote atual
  for (var row = startRow; row <= endRow; row++) {
    // Obtém as coordenadas da linha atual
    var lat = cells.getCell(row, latColumn).getValue();
    var lng = cells.getCell(row, lngColumn).getValue();
    
    // Pula linhas com coordenadas vazias
    if (!lat || !lng) continue;
    
    try {
      // Realiza a geocodificação reversa (coordenadas para endereço)
      var location = geocoder.reverseGeocode(lat, lng);
      
      // Processa o resultado se bem-sucedido
      if (location.status == 'OK') {
        // Obtém o endereço formatado do primeiro resultado
        var address = location["results"][0]["formatted_address"];
        cells.getCell(row, addressColumn).setValue(address);
      } else {
        // Registra o erro quando a geocodificação falha
        cells.getCell(row, addressColumn).setValue("ERRO: " + location.status);
      }
      
      // Adiciona uma pausa entre solicitações para evitar limites de taxa da API
      Utilities.sleep(PAUSE_BETWEEN_REQUESTS);
    } catch (e) {
      // Captura quaisquer erros inesperados
      cells.getCell(row, addressColumn).setValue("ERRO: " + e.toString());
    }
  }
  
  // Verifica se ainda há mais linhas para processar
  if (endRow < totalRows) {
    // Salva a próxima linha a ser processada para uso futuro
    PropertiesService.getDocumentProperties().setProperty('BATCH_CURRENT_ROW', endRow + 1);
    
    // Agenda automaticamente a próxima execução após o intervalo definido
    ScriptApp.newTrigger('positionToAddressBatch')
      .timeBased()
      .after(PAUSE_BETWEEN_BATCH)
      .create();
      
    // Atualiza a célula de status com informações sobre o próximo lote
    statusCell.setValue("Processado lote: linhas " + startRow + " a " + endRow + " de " + totalRows + ". Próximo lote em breve.");
  } else {
    // Quando todas as linhas forem processadas, limpa o estado e finaliza
    PropertiesService.getDocumentProperties().deleteProperty('BATCH_CURRENT_ROW');
    statusCell.setValue("Processamento concluído: " + totalRows + " linhas processadas.");
  }
}

/**
 * FUNCIONALIDADE DE RESET
 * ---------------------
 * Esta função permite reiniciar o estado do processamento em lotes
 * caso ocorra algum problema ou o usuário deseje recomeçar.
 */
function resetBatchProcessing() {
  PropertiesService.getDocumentProperties().deleteProperty('BATCH_CURRENT_ROW');
  SpreadsheetApp.getUi().alert("Status de processamento em lotes redefinido.");
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
      name: "Geocodificar Endereços -> Lat, Long (Original)",
      functionName: "addressToPosition"
    },
    {
      name: "Geocodificar Lat, Long -> Endereços (Original)",
      functionName: "positionToAddress"
    },
    {
      name: "⚡ Geocodificar Endereços -> Lat, Long (Em Lotes)",
      functionName: "addressToPositionBatch"
    },
    {
      name: "⚡ Geocodificar Lat, Long -> Endereços (Em Lotes)",
      functionName: "positionToAddressBatch"
    },
    {
      name: "Redefinir Status do Processamento em Lotes",
      functionName: "resetBatchProcessing"
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