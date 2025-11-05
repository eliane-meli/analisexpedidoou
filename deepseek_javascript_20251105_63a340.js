function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    // Configurar CORS headers
    const headers = {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    };
    
    // Conectar à planilha específica
    const spreadsheetId = '1iLftWtEUacg6iYCzZKhChR6b5nqrNQxTz6R3lulBCro';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('Página1') || spreadsheet.getSheets()[0];
    
    // Obter todos os dados
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    // Verificar se há dados
    if (data.length <= 1) {
      return createResponse({
        success: false,
        error: 'Nenhum dado encontrado na planilha',
        count: 0
      }, headers);
    }
    
    // Obter cabeçalhos
    const headersData = data[0];
    
    // Converter para JSON
    const jsonData = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowData = {};
      
      for (let j = 0; j < headersData.length; j++) {
        const header = headersData[j];
        const value = row[j];
        rowData[header] = value;
      }
      
      // Verificar se a linha tem dados válidos
      if (rowData['Data'] && (rowData['QNTD Expedida'] || rowData['Quantidade Expedida'])) {
        jsonData.push(rowData);
      }
    }
    
    console.log(`Processados ${jsonData.length} registros válidos`);
    
    return createResponse({
      success: true,
      data: jsonData,
      count: jsonData.length,
      message: 'Dados carregados com sucesso',
      lastUpdate: new Date().toISOString()
    }, headers);
    
  } catch (error) {
    console.error('Erro no Apps Script:', error.toString());
    
    return createResponse({
      success: false,
      error: error.toString(),
      message: 'Erro ao acessar a planilha'
    }, headers);
  }
}

function createResponse(data, headers) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  
  // Aplicar headers CORS
  Object.keys(headers).forEach(key => {
    output.setHeader(key, headers[key]);
  });
  
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}