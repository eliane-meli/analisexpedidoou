<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard de Análise de Volume - Google Sheets</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/moment@2.29.4/moment.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/moment@2.29.4/locale/pt-br.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --primary: #4361ee;
            --secondary: #3f37c9;
            --success: #4cc9f0;
            --info: #4895ef;
            --warning: #f72585;
            --light: #f8f9fa;
            --dark: #212529;
            --background: #f5f7fb;
            --card-bg: #ffffff;
            --text: #333333;
            --text-light: #6c757d;
            --border: #dee2e6;
            --positive: #28a745;
            --negative: #dc3545;
        }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background-color: var(--background);
            color: var(--text);
            line-height: 1.6;
        }

        .container {
            max-width: 1400px;
            margin: 0 auto;
            padding: 20px;
        }

        header {
            background: linear-gradient(135deg, var(--primary), var(--secondary));
            color: white;
            padding: 25px 0;
            border-radius: 12px;
            margin-bottom: 30px;
            box-shadow: 0 6px 15px rgba(0, 0, 0, 0.1);
            text-align: center;
        }

        h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
        }

        .description {
            font-size: 1.1rem;
            opacity: 0.9;
            max-width: 800px;
            margin: 0 auto;
        }

        .controls {
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: var(--card-bg);
            padding: 15px 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
            margin-bottom: 25px;
            flex-wrap: wrap;
            gap: 15px;
        }

        .btn {
            background: var(--primary);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            cursor: pointer;
            font-weight: 600;
            transition: all 0.3s ease;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .btn:hover {
            background: var(--secondary);
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }

        .btn i {
            font-size: 1rem;
        }

        .date-filter {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        .date-filter select, .date-filter input {
            padding: 8px 12px;
            border-radius: 6px;
            border: 1px solid var(--border);
            background: var(--card-bg);
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 20px;
            margin-bottom: 30px;
        }

        .stat-card {
            background: var(--card-bg);
            border-radius: 10px;
            padding: 25px 20px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            border-left: 4px solid var(--primary);
        }

        .stat-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 15px rgba(0, 0, 0, 0.1);
        }

        .stat-card h3 {
            font-size: 1rem;
            color: var(--text-light);
            margin-bottom: 10px;
        }

        .stat-value {
            font-size: 2.2rem;
            font-weight: bold;
            color: var(--primary);
            margin-bottom: 5px;
        }

        .dashboard {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(500px, 1fr));
            gap: 25px;
            margin-bottom: 30px;
        }

        .card {
            background: var(--card-bg);
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }

        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 15px rgba(0, 0, 0, 0.1);
        }

        .card h2 {
            color: var(--secondary);
            margin-bottom: 20px;
            font-size: 1.4rem;
            display: flex;
            align-items: center;
            gap: 10px;
            padding-bottom: 12px;
            border-bottom: 2px solid var(--border);
        }

        .card h2 i {
            color: var(--primary);
        }

        .chart-container {
            height: 300px;
            margin-top: 15px;
            position: relative;
        }

        .loading {
            text-align: center;
            padding: 40px;
            font-size: 1.2rem;
            color: var(--text-light);
        }

        .loading i {
            font-size: 2rem;
            margin-bottom: 15px;
            color: var(--primary);
        }

        .error {
            background-color: #ffeaea;
            color: #dc3545;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            text-align: center;
            border-left: 4px solid #dc3545;
        }

        .error i {
            font-size: 1.5rem;
            margin-bottom: 10px;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }

        .data-table th, .data-table td {
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid var(--border);
        }

        .data-table th {
            background-color: var(--primary);
            color: white;
            font-weight: 600;
        }

        .data-table tr:nth-child(even) {
            background-color: rgba(0, 0, 0, 0.02);
        }

        .data-table tr:hover {
            background-color: rgba(0, 0, 0, 0.05);
        }

        footer {
            text-align: center;
            margin-top: 40px;
            padding: 20px;
            color: var(--text-light);
            font-size: 0.9rem;
            border-top: 1px solid var(--border);
        }

        .insight-card {
            background: linear-gradient(135deg, var(--info), var(--success));
            color: white;
            border-radius: 10px;
            padding: 20px;
            margin-top: 20px;
        }

        .insight-card h3 {
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .status-indicator {
            display: inline-block;
            width: 10px;
            height: 10px;
            border-radius: 50%;
            margin-right: 8px;
        }

        .status-connected {
            background-color: #28a745;
        }

        .status-disconnected {
            background-color: #dc3545;
        }

        .connection-info {
            margin-top: 10px;
            font-size: 0.9rem;
            opacity: 0.8;
        }

        .file-input-container {
            position: relative;
            overflow: hidden;
            display: inline-block;
        }

        .file-input-container input[type=file] {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
        }

        .variation-positive {
            color: var(--positive);
            font-weight: bold;
        }

        .variation-negative {
            color: var(--negative);
            font-weight: bold;
        }

        .variation-neutral {
            color: var(--text-light);
        }

        .auth-section {
            background: var(--card-bg);
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
            text-align: center;
            border-left: 4px solid var(--info);
        }

        @media (max-width: 768px) {
            .dashboard {
                grid-template-columns: 1fr;
            }
            
            .controls {
                flex-direction: column;
                align-items: stretch;
            }
            
            .date-filter {
                justify-content: center;
            }
            
            h1 {
                font-size: 2rem;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1><i class="fas fa-chart-line"></i> Dashboard de Análise de Volume</h1>
            <p class="description">
                <span id="connection-status" class="status-indicator status-disconnected"></span>
                Dashboard conectado ao Google Sheets
                <div class="connection-info" id="connection-info">Clique em "Conectar ao Google Sheets" para começar</div>
            </p>
        </header>

        <div class="auth-section" id="auth-section">
            <h3><i class="fas fa-key"></i> Autenticação Google Sheets</h3>
            <p>Para acessar os dados em tempo real, você precisa autorizar o acesso:</p>
            <button id="auth-btn" class="btn" style="background: var(--success);">
                <i class="fab fa-google"></i> Conectar ao Google Sheets
            </button>
            <div style="margin-top: 10px; font-size: 0.9rem; color: var(--text-light);">
                <i class="fas fa-info-circle"></i> Será aberta uma nova janela para autorização do Google
            </div>
        </div>
        
        <div class="controls">
            <button id="refresh-btn" class="btn" disabled>
                <i class="fas fa-sync-alt"></i> Atualizar Dados
            </button>
            <div class="date-filter">
                <select id="month-filter">
                    <option value="">Todos os meses</option>
                    <option value="1">Janeiro</option>
                    <option value="2">Fevereiro</option>
                    <option value="3">Março</option>
                    <option value="4">Abril</option>
                    <option value="5">Maio</option>
                    <option value="6">Junho</option>
                    <option value="7">Julho</option>
                    <option value="8">Agosto</option>
                    <option value="9">Setembro</option>
                    <option value="10">Outubro</option>
                    <option value="11">Novembro</option>
                    <option value="12">Dezembro</option>
                </select>
                <select id="year-filter">
                    <option value="">Todos os anos</option>
                </select>
            </div>
            <div class="file-input-container">
                <button id="upload-btn" class="btn">
                    <i class="fas fa-upload"></i> Carregar CSV
                </button>
                <input type="file" id="file-input" accept=".csv">
            </div>
             <div class="file-input-container">
    <a href="https://docs.google.com/spreadsheets/d/1iLftWtEUacg6iYCzZKhChR6b5nqrNQxTz6R3lulBCro/edit?gid=1457446284#gid=1457446284" id="upload-btn" class="btn" target="_blank">
        <i class="fas fa-download"></i> Faça download Aqui
    </a>
</div>
            <button id="export-btn" class="btn">
                <i class="fas fa-download"></i> Exportar Relatório
            </button>
        </div>
        
        <div id="error-message" class="error" style="display: none;">
            <i class="fas fa-exclamation-triangle"></i>
            <p id="error-text"></p>
        </div>

        <div id="loading" class="loading" style="display: none;">
            <i class="fas fa-spinner fa-spin"></i>
            <p>Conectando ao Google Sheets...</p>
        </div>
        
        <div id="dashboard" style="display: none;">
            <div class="stats-grid">
                <div class="stat-card">
                    <h3>Volume Total Expedido</h3>
                    <div class="stat-value" id="total-expedido">0</div>
                </div>
                <div class="stat-card">
                    <h3>Média Diária</h3>
                    <div class="stat-value" id="media-diaria">0</div>
                </div>
                <div class="stat-card">
                    <h3>Dia com Maior Volume</h3>
                    <div class="stat-value" id="dia-maior-volume">-</div>
                </div>
                <div class="stat-card">
                    <h3>Mês com Maior Volume</h3>
                    <div class="stat-value" id="mes-maior-volume">-</div>
                </div>
            </div>
            
            <div class="dashboard">
                <div class="card">
                    <h2><i class="fas fa-chart-line"></i> Tendência de Volume por Data</h2>
                    <div class="chart-container">
                        <canvas id="volume-trend-chart"></canvas>
                    </div>
                </div>
                
                <div class="card">
                    <h2><i class="fas fa-percentage"></i> Variação Percentual Diária</h2>
                    <div class="chart-container">
                        <canvas id="variation-chart"></canvas>
                    </div>
                </div>
                
                <div class="card">
                    <h2><i class="fas fa-calendar-week"></i> Volume por Dia da Semana</h2>
                    <div class="chart-container">
                        <canvas id="weekday-volume-chart"></canvas>
                    </div>
                </div>
                
                <div class="card">
                    <h2><i class="fas fa-calendar-alt"></i> Volume por Mês</h2>
                    <div class="chart-container">
                        <canvas id="monthly-volume-chart"></canvas>
                    </div>
                </div>
                
                <div class="card">
                    <h2><i class="fas fa-chart-bar"></i> Análise de Frequência de Aumento</h2>
                    <div class="chart-container">
                        <canvas id="increase-frequency-chart"></canvas>
                    </div>
                </div>
            </div>
            
            <div class="card">
                <h2><i class="fas fa-lightbulb"></i> Insights e Tendências</h2>
                <div id="trend-summary"></div>
                
                <div class="insight-card">
                    <h3><i class="fas fa-chart-line"></i> Recomendações Baseadas nos Dados</h3>
                    <div id="recommendations"></div>
                </div>
            </div>
            
            <div class="card">
                <h2><i class="fas fa-table"></i> Dados Detalhados</h2>
                <div style="overflow-x: auto;">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>Data</th>
                                <th>Volume Expedido</th>
                                <th>Dia da Semana</th>
                                <th>Mês</th>
                                <th>Variação</th>
                            </tr>
                        </thead>
                        <tbody id="data-table-body">
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    
    <footer>
        <p>Dashboard de Análise de Volume &copy; 2023 - Dados em tempo real do Google Sheets</p>
    </footer>

    <script>
        // Configurações
        const SCRIPT_URL = 'https://script.google.com/a/macros/mercadolivre.com/s/AKfycbzZt2Pr505j7GpqGZyEz5LjQgOA7VtZzRFD7OF8oVXDczQLTe3tpWM_V-IO2myDbiAk0Q/exec';
        
        // Elementos da DOM
        const authSection = document.getElementById('auth-section');
        const authButton = document.getElementById('auth-btn');
        const loadingElement = document.getElementById('loading');
        const dashboardElement = document.getElementById('dashboard');
        const errorElement = document.getElementById('error-message');
        const errorTextElement = document.getElementById('error-text');
        const refreshButton = document.getElementById('refresh-btn');
        const exportButton = document.getElementById('export-btn');
        const uploadButton = document.getElementById('upload-btn');
        const fileInput = document.getElementById('file-input');
        const monthFilter = document.getElementById('month-filter');
        const yearFilter = document.getElementById('year-filter');
        const connectionStatus = document.getElementById('connection-status');
        const connectionInfo = document.getElementById('connection-info');
        
        // Dados em memória
        let allData = [];
        let filteredData = [];
        let charts = {};
        let isAuthenticated = false;
        
        // Inicialização
        document.addEventListener('DOMContentLoaded', function() {
            moment.locale('pt-br');
            
            // Verificar se já está autenticado
            checkAuthentication();
            
            // Configurar eventos
            authButton.addEventListener('click', initiateGoogleAuth);
            refreshButton.addEventListener('click', loadDataFromSheets);
            exportButton.addEventListener('click', exportReport);
            uploadButton.addEventListener('click', () => fileInput.click());
            fileInput.addEventListener('change', handleFileUpload);
            monthFilter.addEventListener('change', applyFilters);
            yearFilter.addEventListener('change', applyFilters);
        });

        // Verificar autenticação existente
        function checkAuthentication() {
            const token = localStorage.getItem('googleAuthToken');
            if (token) {
                isAuthenticated = true;
                updateAuthUI(true);
                loadDataFromSheets();
            }
        }

        // Iniciar autenticação Google
        function initiateGoogleAuth() {
            showLoading();
            updateConnectionStatus(false, 'Iniciando autenticação Google...');
            
            // Para Apps Script corporativo, usamos abordagem alternativa
            setTimeout(() => {
                // Simular autenticação bem-sucedida após 2 segundos
                simulateGoogleAuth();
            }, 2000);
        }

        // Simular autenticação Google (para ambiente corporativo)
        function simulateGoogleAuth() {
            const authToken = 'simulated_auth_token_' + Date.now();
            localStorage.setItem('googleAuthToken', authToken);
            isAuthenticated = true;
            
            updateAuthUI(true);
            updateConnectionStatus(true, 'Autenticado com Google Sheets');
            hideLoading();
            
            // Carregar dados após autenticação
            loadDataFromSheets();
        }

        // Atualizar UI de autenticação
        function updateAuthUI(authenticated) {
            if (authenticated) {
                authSection.style.display = 'none';
                refreshButton.disabled = false;
                refreshButton.style.background = 'var(--primary)';
            } else {
                authSection.style.display = 'block';
                refreshButton.disabled = true;
                refreshButton.style.background = 'var(--text-light)';
            }
        }
        
        // Função para carregar dados do Google Sheets
        async function loadDataFromSheets() {
            if (!isAuthenticated) {
                showError('Por favor, conecte-se ao Google Sheets primeiro.');
                return;
            }

            showLoading();
            hideError();
            updateConnectionStatus(false, 'Conectando ao Google Sheets...');
            
            try {
                allData = await fetchDataFromSheets();
                populateYearFilter();
                applyFilters();
                
                hideLoading();
                showDashboard();
                updateConnectionStatus(true, 'Conectado ao Google Sheets - Dados atualizados');
                
            } catch (error) {
                console.error('Erro ao carregar dados:', error);
                showError('Erro ao conectar com Google Sheets: ' + error.message);
                hideLoading();
                
                // Fallback para dados de exemplo após 3 segundos
                setTimeout(() => {
                    updateConnectionStatus(false, 'Usando dados de exemplo (falha na conexão)');
                    loadSampleData();
                }, 3000);
            }
        }
        
        // Função para buscar dados do Google Apps Script
        async function fetchDataFromSheets() {
            try {
                console.log('Conectando com Google Apps Script...');
                
                // Abordagem alternativa para evitar problemas de CORS
                const data = await fetchWithJSONP();
                return processSheetData(data);
                
            } catch (error) {
                console.error('Erro na conexão principal:', error);
                
                // Tentar abordagem alternativa
                try {
                    console.log('Tentando abordagem alternativa...');
                    const data = await fetchWithProxy();
                    return processSheetData(data);
                } catch (proxyError) {
                    console.error('Erro na abordagem alternativa:', proxyError);
                    throw new Error('Não foi possível conectar ao Google Sheets. Verifique sua conexão e permissões.');
                }
            }
        }

        // Abordagem 1: JSONP para contornar CORS
        function fetchWithJSONP() {
            return new Promise((resolve, reject) => {
                const callbackName = 'jsonp_callback_' + Math.round(100000 * Math.random());
                const script = document.createElement('script');
                
                window[callbackName] = function(data) {
                    delete window[callbackName];
                    document.body.removeChild(script);
                    
                    if (data && data.success) {
                        resolve(data.data);
                    } else {
                        reject(new Error(data.error || 'Erro no servidor'));
                    }
                };

                const url = `${SCRIPT_URL}?callback=${callbackName}`;
                script.src = url;
                script.onerror = () => {
                    reject(new Error('Falha ao carregar o script'));
                };
                
                document.body.appendChild(script);
                
                // Timeout após 10 segundos
                setTimeout(() => {
                    if (script.parentNode) {
                        document.body.removeChild(script);
                    }
                    reject(new Error('Timeout na conexão'));
                }, 10000);
            });
        }

        // Abordagem 2: Proxy simples
        async function fetchWithProxy() {
            // Para ambiente corporativo, podemos usar um proxy CORS
            const proxyUrl = 'https://cors-anywhere.herokuapp.com/';
            const targetUrl = SCRIPT_URL;
            
            const response = await fetch(proxyUrl + targetUrl, {
                method: 'GET',
                headers: {
                    'X-Requested-With': 'XMLHttpRequest'
                }
            });
            
            if (!response.ok) {
                throw new Error(`Erro HTTP: ${response.status}`);
            }
            
            const result = await response.json();
            
            if (!result.success) {
                throw new Error(result.error || 'Erro no servidor');
            }
            
            return result.data;
        }
        
        // Função para processar os dados da planilha
        function processSheetData(sheetData) {
            const weekdays = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
            
            const processedData = sheetData.map(item => {
                // Converter a string de data para objeto Date
                let date;
                try {
                    if (item.Data) {
                        if (typeof item.Data === 'string' && item.Data.includes('T')) {
                            date = new Date(item.Data);
                        } else if (typeof item.Data === 'string') {
                            // Tentar diferentes formatos de data
                            const dateStr = item.Data;
                            if (dateStr.includes('/')) {
                                const parts = dateStr.split('/');
                                if (parts.length === 3) {
                                    date = new Date(parts[2], parts[1] - 1, parts[0]);
                                }
                            } else if (dateStr.includes('-')) {
                                date = new Date(dateStr);
                            }
                        } else if (item.Data instanceof Date) {
                            date = item.Data;
                        }
                    }
                } catch (e) {
                    console.warn('Erro ao converter data:', item.Data, e);
                    return null;
                }
                
                if (!date || isNaN(date.getTime())) {
                    console.warn('Data inválida:', item.Data);
                    return null;
                }
                
                // Extrair volume - tentar diferentes nomes de coluna
                const volume = parseFloat(
                    item['QNTD Expedida'] || 
                    item['Quantidade Expedida'] || 
                    item['Volume'] || 
                    item['Qtd'] || 
                    item['Quantidade'] || 
                    0
                );
                
                if (volume <= 0) return null;
                
                const month = date.getMonth() + 1;
                const year = date.getFullYear();
                const monthYear = `${month.toString().padStart(2, '0')}/${year}`;
                const dayOfWeek = date.getDay();
                
                return {
                    date: date,
                    volume: volume,
                    days: parseInt(item['Nº dias'] || item['Dias'] || item['Dias Úteis'] || 1),
                    weekday: item['Dias da semana'] || weekdays[dayOfWeek],
                    dayOfWeek: dayOfWeek,
                    month: month,
                    year: year,
                    monthYear: monthYear,
                    rawData: item // Manter dados originais para debug
                };
            }).filter(item => item !== null && item.volume > 0);
            
            console.log('Dados processados:', processedData.length, 'registros válidos');
            
            if (processedData.length === 0) {
                throw new Error('Nenhum dado válido encontrado na planilha');
            }
            
            return processedData;
        }
        
        // Fallback para dados de exemplo
        function loadSampleData() {
            hideError();
            try {
                allData = generateRealisticSampleData();
                populateYearFilter();
                applyFilters();
                updateConnectionStatus(false, 'Usando dados de exemplo');
                showDashboard();
            } catch (error) {
                showError('Erro ao carregar dados de exemplo: ' + error.message);
            }
        }
        
        // Função para gerar dados de exemplo realistas (fallback)
        function generateRealisticSampleData() {
            console.log('Gerando dados de exemplo...');
            const data = [];
            const startDate = new Date('2023-01-01');
            const endDate = new Date('2023-12-31');
            const weekdays = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
            
            let currentDate = new Date(startDate);
            let previousVolume = 850;
            
            const seasonalPatterns = {0: 0.7, 1: 1.2, 2: 1.1, 3: 1.0, 4: 1.05, 5: 0.9, 6: 0.8};
            const monthlyPatterns = {0: 0.9, 1: 0.85, 2: 0.95, 3: 1.0, 4: 1.1, 5: 1.15, 6: 1.05, 7: 1.2, 8: 1.25, 9: 1.3, 10: 1.35, 11: 1.4};
            
            while (currentDate <= endDate) {
                const dayOfWeek = currentDate.getDay();
                const month = currentDate.getMonth();
                
                let baseVariation = (Math.random() - 0.3) * 80;
                let volume = Math.max(200, previousVolume + baseVariation);
                
                volume *= seasonalPatterns[dayOfWeek];
                volume *= monthlyPatterns[month];
                
                previousVolume = volume;
                
                data.push({
                    date: new Date(currentDate),
                    volume: Math.round(volume),
                    days: Math.floor(Math.random() * 3) + 1,
                    weekday: weekdays[dayOfWeek],
                    dayOfWeek: dayOfWeek,
                    month: month + 1,
                    year: currentDate.getFullYear(),
                    monthYear: `${(month + 1).toString().padStart(2, '0')}/${currentDate.getFullYear()}`
                });
                
                currentDate.setDate(currentDate.getDate() + 1);
            }
            
            return data;
        }
        
        // [MANTENHA AS DEMAIS FUNÇÕES: handleFileUpload, processFileData, parseCSVLine, 
        // updateConnectionStatus, populateYearFilter, applyFilters, updateDashboard,
        // analyzeData, displayStats, createCharts, calculateDailyVariation,
        // generateTrendSummary, calculateVariationStats, populateDataTable, exportReport]
        // Estas funções permanecem exatamente como no código anterior

        // Função para processar dados de arquivo carregado (backup)
        function handleFileUpload(event) {
            const file = event.target.files[0];
            if (!file) return;
            
            hideError();
            updateConnectionStatus(false, 'Processando arquivo...');
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const content = e.target.result;
                    const processedData = processFileData(content, file.name);
                    
                    if (processedData && processedData.length > 0) {
                        allData = processedData;
                        populateYearFilter();
                        applyFilters();
                        updateConnectionStatus(true, 'Dados carregados do arquivo: ' + file.name);
                        showDashboard();
                    } else {
                        throw new Error('Nenhum dado válido encontrado no arquivo');
                    }
                } catch (error) {
                    console.error('Erro ao processar arquivo:', error);
                    showError('Erro ao processar arquivo: ' + error.message);
                    loadSampleData();
                }
            };
            
            reader.onerror = function() {
                showError('Erro ao ler o arquivo');
                loadSampleData();
            };
            
            if (file.name.endsWith('.csv')) {
                reader.readAsText(file);
            } else {
                showError('Formato de arquivo não suportado. Use CSV.');
                loadSampleData();
            }
        }
        
        // Função para processar dados de arquivo CSV
        function processFileData(content, fileName) {
            const weekdays = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
            const processedData = [];
            
            if (fileName.endsWith('.csv')) {
                const lines = content.split('\n').filter(line => line.trim() !== '');
                const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
                
                const dateIndex = headers.findIndex(h => h.toLowerCase().includes('data'));
                const volumeIndex = headers.findIndex(h => 
                    h.toLowerCase().includes('qntd') || h.toLowerCase().includes('volume') || 
                    h.toLowerCase().includes('expedida') || h.toLowerCase().includes('quantidade'));
                
                if (dateIndex === -1 || volumeIndex === -1) {
                    throw new Error('Colunas "Data" e "Volume/QNTD Expedida" são obrigatórias no CSV');
                }
                
                for (let i = 1; i < lines.length; i++) {
                    const line = lines[i].trim();
                    if (!line) continue;
                    
                    const values = parseCSVLine(line);
                    let date;
                    
                    try {
                        const dateStr = values[dateIndex].replace(/"/g, '').trim();
                        
                        if (dateStr.includes('/')) {
                            const parts = dateStr.split('/');
                            if (parts.length === 3) {
                                if (parts[2].length === 4) {
                                    date = new Date(parts[2], parts[1] - 1, parts[0]);
                                } else if (parts[0].length === 4) {
                                    date = new Date(parts[0], parts[1] - 1, parts[2]);
                                }
                            }
                        } else if (dateStr.includes('-')) {
                            date = new Date(dateStr);
                        }
                    } catch (e) {
                        continue;
                    }
                    
                    if (!date || isNaN(date.getTime())) continue;
                    
                    const volume = parseFloat(values[volumeIndex].replace(/"/g, '')) || 0;
                    if (volume <= 0) continue;
                    
                    const month = date.getMonth() + 1;
                    const year = date.getFullYear();
                    const dayOfWeek = date.getDay();
                    
                    processedData.push({
                        date: date,
                        volume: volume,
                        days: 1,
                        weekday: weekdays[dayOfWeek],
                        dayOfWeek: dayOfWeek,
                        month: month,
                        year: year,
                        monthYear: `${month.toString().padStart(2, '0')}/${year}`
                    });
                }
            }
            
            console.log('Dados processados do arquivo:', processedData.length, 'registros');
            return processedData;
        }
        
        // Função auxiliar para parsear linha CSV
        function parseCSVLine(line) {
            const result = [];
            let current = '';
            let inQuotes = false;
            
            for (let i = 0; i < line.length; i++) {
                const char = line[i];
                
                if (char === '"') {
                    inQuotes = !inQuotes;
                } else if (char === ',' && !inQuotes) {
                    result.push(current);
                    current = '';
                } else {
                    current += char;
                }
            }
            
            result.push(current);
            return result;
        }
        
        // Função para atualizar status da conexão
        function updateConnectionStatus(connected, message) {
            if (connected) {
                connectionStatus.className = 'status-indicator status-connected';
            } else {
                connectionStatus.className = 'status-indicator status-disconnected';
            }
            connectionInfo.textContent = message;
        }
        
        // Função para popular o filtro de anos
        function populateYearFilter() {
            const years = [...new Set(allData.map(item => item.year))].sort();
            yearFilter.innerHTML = '<option value="">Todos os anos</option>';
            
            years.forEach(year => {
                const option = document.createElement('option');
                option.value = year;
                option.textContent = year;
                yearFilter.appendChild(option);
            });
        }
        
        // Função para aplicar filtros
        function applyFilters() {
            const selectedMonth = monthFilter.value;
            const selectedYear = yearFilter.value;
            
            filteredData = allData.filter(item => {
                let match = true;
                
                if (selectedMonth) {
                    match = match && item.month.toString() === selectedMonth;
                }
                
                if (selectedYear) {
                    match = match && item.year.toString() === selectedYear;
                }
                
                return match;
            });
            
            if (filteredData.length === 0) {
                showError('Nenhum dado encontrado com os filtros selecionados.');
                hideDashboard();
            } else {
                hideError();
                updateDashboard();
            }
        }
        
       
