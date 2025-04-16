# 📍 Google Sheets Geocoding em Lotes

Este script para Google Sheets permite geocodificar grandes quantidades de endereços ou coordenadas (até milhares de linhas), superando as limitações de tempo de execução padrão do Google Apps Script através de um sistema de processamento em lotes inteligente.

## 🌟 Características

- **Processamento em Lotes**: Divide grandes conjuntos de dados em lotes menores e gerenciáveis  
- **Geocodificação Bidirecional**: Converte endereços em coordenadas (lat/lng) e vice-versa  
- **Retomada Automática**: Continua de onde parou caso o processamento seja interrompido  
- **Interface Amigável**: Menu integrado na planilha para fácil acesso  
- **Feedback Visual**: Exibe o progresso em tempo real na planilha  
- **Tratamento de Erros**: Identifica e registra problemas durante a geocodificação  
- **Baixo Consumo de API**: Evita limites de taxa com pausas estratégicas  

## ⚙️ Funcionamento

O script utiliza o serviço Maps do Google Apps Script para realizar geocodificações, mas supera as limitações padrão:

| Problema Original              | Solução Implementada                                |
|-------------------------------|------------------------------------------------------|
| Limite de 6 minutos de execução | Processamento em lotes com gatilhos programados     |
| Limite de taxa da API          | Pausas controladas entre solicitações               |
| Falhas silenciosas             | Sistema robusto de tratamento e registro de erros   |
| Sem feedback de progresso      | Célula de status mostrando etapas do processamento  |
| Reinício do zero em caso de erro | Sistema de retomada para continuar de onde parou  |

## 🚀 Como Usar

### Instalação

1. Abra sua planilha no Google Sheets  
2. Vá para **Extensões > Apps Script**  
3. Cole o código do arquivo [`geocoding_script.gs`](geocoding_script.gs) no editor  
4. Salve e volte para sua planilha  
5. Atualize a página para visualizar o novo menu **Geocode**  

### Utilização Básica

1. Organize seus dados em três colunas:
   - Coluna 1: Endereços (ou vazia para geocodificação reversa)
   - Coluna 2: Latitude (ou vazia para geocodificação direta)
   - Coluna 3: Longitude (ou vazia para geocodificação direta)

2. Selecione as três colunas, incluindo todas as linhas a processar  

3. No menu **Geocode**, escolha:
   - **Geocodificar Endereços → Lat, Long (Em Lotes)** para converter endereços em coordenadas
   - **Geocodificar Lat, Long → Endereços (Em Lotes)** para conversão reversa  

4. O processamento será iniciado em lotes. Uma célula de status será exibida à direita da seleção, mostrando o progresso.

### Para Grandes Volumes de Dados

1. Utilize as opções "Em Lotes" no menu  
2. O processamento ocorrerá em lotes de 50 linhas (valor padrão, configurável)  
3. Caso o navegador seja fechado, o processo pode ser retomado ao reabrir a planilha  
4. Para reiniciar completamente, utilize **"Redefinir Status do Processamento em Lotes"**

## ⚠️ Considerações Importantes

- **Limites da API**: A API do Google Maps impõe um limite de 2.500 solicitações por dia para contas gratuitas  
- **Tempo de Processamento**: A depender do volume de dados, o processo pode durar várias horas  
- **Região de Geocodificação**: O script utiliza 'us' como região padrão. Modifique a função `getGeocodingRegion()` para outro país (ex: 'br', 'uk')  
- **Formato dos Endereços**: Recomenda-se o uso de endereços completos (logradouro, cidade, estado, país)

## 🔧 Configurações Avançadas

Você pode personalizar os parâmetros de desempenho no início do script:

```javascript
const BATCH_SIZE = 50;              // Número de linhas por lote
const PAUSE_BETWEEN_BATCH = 1000;   // Pausa entre lotes (em milissegundos)
const PAUSE_BETWEEN_REQUESTS = 200; // Pausa entre cada requisição (em milissegundos)
