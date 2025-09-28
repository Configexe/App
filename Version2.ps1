# -----------------------------------------------------------------------------
# PowerChart Designer: Editor de Dados
# Versão: 15.0 - Edição de Renderização Estável
# Autor: Seu Nome/Empresa
# Descrição: Corrige o bug crítico de crescimento infinito do gráfico ao
#            atualizar, recriando o canvas de forma robusta a cada renderização.
#            Garante uma experiência de visualização estável e confiável.
# -----------------------------------------------------------------------------

# --- 1. Carregar Assemblies Necessárias ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "Não foi possível carregar as assemblies necessárias."
    exit 1
}

# --- 2. Funções Principais ---

Function Load-CSVData {
    param(
        [Parameter(Mandatory=$true)]$DataGridView,
        [Parameter(Mandatory=$true)]$StatusLabel,
        [Parameter(Mandatory=$true)]$GenerateButton
    )

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Selecione o arquivo CSV"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        $StatusLabel.Text = "Analisando: $(Split-Path $FilePath -Leaf)..."
        $StatusLabel.Refresh()

        $Data = $null
        try {
            $firstLine = Get-Content -Path $FilePath -TotalCount 1
            $bestDelimiter = if (($firstLine -split ';').Count -gt ($firstLine -split ',').Count) { ';' } else { ',' }
            $Data = Import-Csv -Path $FilePath -Delimiter $bestDelimiter
        }
        catch {
            # O erro será tratado abaixo
        }

        if ($null -ne $Data -and $Data.Count -gt 0) {
            $DataGridView.DataSource = [System.Collections.ArrayList]$Data
            $DataGridView.AutoSizeColumnsMode = 'AllCells'
            $StatusLabel.Text = "Arquivo carregado: $(Split-Path $FilePath -Leaf)"
            $GenerateButton.Enabled = $true
        } else {
            $DataGridView.DataSource = $null
            [System.Windows.Forms.MessageBox]::Show("Não foi possível ler os dados do arquivo CSV.", "Erro de Leitura", "OK", "Error")
            $StatusLabel.Text = "Falha ao carregar arquivo."
            $GenerateButton.Enabled = $false
        }
    }
}

Function Generate-HtmlReport {
    param(
        [Parameter(Mandatory=$true)]$DataGridView,
        [Parameter(Mandatory=$true)]$StatusLabel
    )

    if ($null -eq $DataGridView.DataSource -or $DataGridView.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Não há dados carregados para gerar o relatório.", "Aviso", "OK", "Warning")
        return
    }

    $StatusLabel.Text = "Gerando relatório HTML..."
    $StatusLabel.Refresh()

    $DataForJson = $DataGridView.DataSource | ForEach-Object {
        $properties = @{}
        foreach ($prop in $_.PSObject.Properties) {
            $properties[$prop.Name] = $prop.Value
        }
        New-Object -TypeName PSObject -Property $properties
    }

    $JsonData = $DataForJson | ConvertTo-Json -Compress -Depth 5
    $JsonColumnNames = $DataGridView.Columns.DataPropertyName | ConvertTo-Json -Compress

    $OutputPath = Join-Path $env:TEMP "PowerChart_Relatorio.html"
    $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -JsonColumnNames $JsonColumnNames
    
    try {
        $HtmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
        Start-Process $OutputPath
        $StatusLabel.Text = "Relatório gerado e aberto com sucesso!"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao gerar ou abrir o arquivo HTML: $($_.Exception.Message)", "Erro", "OK", "Error")
        $StatusLabel.Text = "Falha ao gerar o relatório."
    }
}

Function Get-HtmlTemplate {
    param($JsonData, $JsonColumnNames)

    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PowerChart - Relatório Dinâmico</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        .card { background-color: white; border-radius: 0.75rem; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1); padding: 1.5rem; transition: all 0.3s ease-in-out; }
        .kpi-value { font-size: 2rem; font-weight: 900; color: #1e293b; }
        .kpi-label { font-size: 0.875rem; color: #64748b; margin-top: 0.25rem; }
        .chart-container { position: relative; width: 100%; height: 550px; } /* Altura fixa para estabilidade */
        .chart-selector label { border: 2px solid #e5e7eb; border-radius: 0.5rem; padding: 0.75rem; cursor: pointer; transition: all 0.2s ease-in-out; text-align: center; }
        .chart-selector label:hover { border-color: #9ca3af; background-color: #f9fafb; }
        .chart-selector input:checked + label { border-color: #3b82f6; background-color: #eff6ff; box-shadow: 0 0 0 2px #3b82f6; }
        .chart-selector input { display: none; }
        .control-hidden { display: none !important; }
        .format-panel input[type='number'] { -moz-appearance: textfield; }
        .format-panel input[type='number']::-webkit-inner-spin-button, 
        .format-panel input[type='number']::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        .divider { border-top: 1px solid #e5e7eb; }
    </style>
</head>
<body class="text-gray-900">
    <header class="bg-[#0f172a] text-white text-center py-12 px-4">
        <h1 class="text-4xl md:text-5xl font-black tracking-tight">Relatório Dinâmico Interativo</h1>
        <p class="mt-4 text-lg text-blue-200 max-w-3xl mx-auto">Dados processados via PowerChart Editor.</p>
    </header>
    <main class="container mx-auto p-4 md:p-8 -mt-10">
        <section id="controls" class="card mb-6">
             <div class="grid grid-cols-1 md:grid-cols-2 gap-8">
                <div>
                    <h2 class="text-xl font-bold text-[#1e293b] mb-4">1. Seleção de Dados</h2>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-4 items-end">
                        <div><label for="x-axis">Eixo X:</label><select id="x-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                        <div><label for="y1-axis">Série Y1:</label><select id="y1-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                        <div id="y2-axis-control"><label for="y2-axis">Série Y2:</label><select id="y2-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                        <div><label for="y1-color">Cor Y1:</label><input type="color" id="y1-color" value="#3b82f6" class="w-full h-10 mt-1"></div>
                        <div id="y2-color-control"><label for="y2-color">Cor Y2:</label><input type="color" id="y2-color" value="#ef4444" class="w-full h-10 mt-1"></div>
                    </div>
                </div>
                <div class="flex flex-col justify-between">
                    <div>
                        <h2 class="text-xl font-bold text-[#1e293b] mb-4">2. Escolha o Tipo de Gráfico</h2>
                         <div class="chart-selector grid grid-cols-4 gap-4">
                            <div><input type="radio" name="chart-type" value="combo" id="type-combo" checked><label for="type-combo" class="flex flex-col items-center justify-center h-full"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/><path d="M3 12l5-4 5 6 5-4"/></svg><span class="text-xs font-semibold">Combo</span></label></div>
                            <div><input type="radio" name="chart-type" value="bar" id="type-bar"><label for="type-bar" class="flex flex-col items-center justify-center h-full"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg><span class="text-xs font-semibold">Barras</span></label></div>
                            <div><input type="radio" name="chart-type" value="line" id="type-line"><label for="type-line" class="flex flex-col items-center justify-center h-full"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M3 17l5-4 5 6 5-4 4 2"/></svg><span class="text-xs font-semibold">Linha</span></label></div>
                            <div><input type="radio" name="chart-type" value="stacked" id="type-stacked"><label for="type-stacked" class="flex flex-col items-center justify-center h-full"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="#3b82f6" stroke="#3b82f6" stroke-width="1"><rect x="5" y="12" width="4" height="6"/><rect x="10" y="8" width="4" height="10"/><rect x="15" y="4" width="4" height="14"/><path d="M5 12V9h4v3m1-4V4h4v4m1-4V2h4v2" fill="#ef4444"/></svg><span class="text-xs font-semibold">Empilhado</span></label></div>
                         </div>
                    </div>
                     <div class="mt-4"><button id="update-charts-btn" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg transition hover:bg-blue-700 flex items-center justify-center">
                        <span id="btn-text">Gerar / Atualizar Gráfico</span>
                        <svg id="btn-spinner" class="animate-spin ml-3 h-5 w-5 text-white hidden" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                            <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                            <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                        </svg>
                     </button></div>
                </div>
            </div>
        </section>

        <section id="kpis" class="mb-6"><div id="kpi-grid" class="grid grid-cols-1 md:grid-cols-3 gap-6"></div></section>
        
        <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">
            <div class="lg:col-span-3 card">
                 <div class="flex justify-between items-center mb-4">
                    <h3 id="chart-title" class="text-xl font-bold text-[#1e293b]"></h3>
                    <button id="download-btn" class="bg-gray-100 text-gray-700 hover:bg-gray-200 font-bold py-2 px-4 rounded-lg transition text-sm flex items-center">
                        <svg class="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                        Baixar (PNG)
                    </button>
                </div>
                <div class="chart-container"><canvas id="mainChart"></canvas></div>
            </div>
            <div id="format-panel-container" class="lg:col-span-1">
                <div class="card format-panel">
                    <h3 class="text-xl font-bold text-[#1e293b] mb-4">Formatar Visual</h3>
                    <div class="space-y-4">
                        <div>
                            <span class="font-semibold text-gray-700 text-sm">Aparência</span>
                            <div class="flex items-center mt-2"><input id="dark-mode" type="checkbox" class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="dark-mode" class="ml-2 block text-sm text-gray-900">Modo Escuro</label></div>
                        </div>
                        <div class="divider"></div>
                        <div>
                            <span class="font-semibold text-gray-700 text-sm">Rótulos de Dados</span>
                            <div class="flex items-center mt-2"><input id="show-labels" type="checkbox" class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="show-labels" class="ml-2 block text-sm text-gray-900">Exibir rótulos</label></div>
                        </div>
                        <div class="divider"></div>
                        <div id="show-grid-control">
                            <span class="font-semibold text-gray-700 text-sm">Linhas de Grade</span>
                            <div class="flex items-center mt-2"><input id="show-grid" type="checkbox" checked class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="show-grid" class="ml-2 block text-sm text-gray-900">Exibir grades</label></div>
                        </div>
                        <div id="y-axis-max-control" class="divider">
                             <span class="font-semibold text-gray-700 text-sm">Eixo Y (Primário)</span>
                             <div class="flex items-center mt-2"><input id="y-axis-auto" type="checkbox" checked class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="y-axis-auto" class="ml-2 block text-sm text-gray-900">Automático</label></div>
                             <input type="number" id="y-axis-max" placeholder="Ex: 100" disabled class="mt-2 block w-full rounded-md border-gray-300 shadow-sm text-sm disabled:bg-gray-200 disabled:cursor-not-allowed">
                        </div>
                         <div id="y2-axis-max-control" class="divider">
                             <span class="font-semibold text-gray-700 text-sm">Eixo Y (Secundário)</span>
                             <div class="flex items-center mt-2"><input id="y2-axis-auto" type="checkbox" checked class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="y2-axis-auto" class="ml-2 block text-sm text-gray-900">Automático</label></div>
                             <input type="number" id="y2-axis-max" placeholder="Ex: 100" disabled class="mt-2 block w-full rounded-md border-gray-300 shadow-sm text-sm disabled:bg-gray-200 disabled:cursor-not-allowed">
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </main>
    <script>
        // --- INÍCIO DO JAVASCRIPT REESCRITO (VERSÃO 15.0) ---
        var RAW_DATA = $JsonData;
        var COLUMN_NAMES = $JsonColumnNames;
        var chartInstance;

        Chart.register(ChartDataLabels);

        function parseNumericValue(value) {
            if (typeof value === 'number') return value;
            if (typeof value !== 'string') return 0;
            var cleanValue = value.replace(/[^0-9,-]/g, '').replace(',', '.');
            var parsed = parseFloat(cleanValue);
            return isNaN(parsed) ? 0 : parsed;
        }

        function findDefaultAxes() {
            var defaults = { xAxis: null, y1Axis: null, y2Axis: null };
            var numericCols = [], textCols = [];
            if (!RAW_DATA || RAW_DATA.length === 0) return defaults;
            var firstRow = RAW_DATA[0];
            for (var i = 0; i < COLUMN_NAMES.length; i++) {
                var colName = COLUMN_NAMES[i];
                var value = firstRow[colName];
                if (value === null || typeof value === 'undefined') continue;
                var parsedValue = parseFloat(String(value).replace(',', '.'));
                if (!isNaN(parsedValue) && String(value).trim() !== '') { numericCols.push(colName); } else { textCols.push(colName); }
            }
            defaults.xAxis = textCols[0] || COLUMN_NAMES[0] || null;
            defaults.y1Axis = numericCols[0] || (COLUMN_NAMES.length > 1 ? COLUMN_NAMES[1] : null);
            defaults.y2Axis = numericCols[1] || (COLUMN_NAMES.length > 2 ? COLUMN_NAMES[2] : null);
            return defaults;
        }

        function updateKPIs(dataY1, dataY2, y1Name, y2Name, labels, isY2Enabled) {
            var totalY1 = dataY1.reduce(function(a, b) { return a + b; }, 0);
            var totalY2 = isY2Enabled ? dataY2.reduce(function(a, b) { return a + b; }, 0) : 0;
            var maxY1 = -Infinity, bestIndex = -1;
            for (var i = 0; i < dataY1.length; i++) { if (dataY1[i] > maxY1) { maxY1 = dataY1[i]; bestIndex = i; } }
            var bestXLabel = labels[bestIndex];
            var kpiGrid = document.getElementById('kpi-grid');
            kpiGrid.innerHTML = '<div class="card"><div class="kpi-value">' + totalY1.toLocaleString('pt-BR') + '</div><div class="kpi-label">Total de ' + y1Name + '</div></div>' + '<div class="card"><div class="kpi-value">' + (isY2Enabled ? totalY2.toLocaleString('pt-BR') : 'N/A') + '</div><div class="kpi-label">Total de ' + (isY2Enabled ? y2Name : '-') + '</div></div>' + '<div class="card"><div class="kpi-value">' + bestXLabel + '</div><div class="kpi-label">Ponto de Maior ' + y1Name + '</div></div>';
        }
        
        function populateControls(defaultSelections) {
            var axes = ['x-axis', 'y1-axis', 'y2-axis'];
            axes.forEach(function(id) {
                var select = document.getElementById(id);
                select.innerHTML = (id === 'y2-axis') ? '<option value="Nenhum">Nenhum</option>' : '';
                COLUMN_NAMES.forEach(function(name) { select.innerHTML += '<option value="' + name + '">' + name + '</option>'; });
            });
            document.getElementById('x-axis').value = defaultSelections.xAxis || '';
            document.getElementById('y1-axis').value = defaultSelections.y1Axis || '';
            document.getElementById('y2-axis').value = defaultSelections.y2Axis || 'Nenhum';
        }

        function updateControlVisibility() {
            var chartType = document.querySelector('input[name="chart-type"]:checked').value;
            var controlsToReset = ["y2-axis-control", "y2-color-control", "show-grid-control", "y-axis-max-control", "y2-axis-max-control"];
            controlsToReset.forEach(function(id) { document.getElementById(id).classList.remove('control-hidden'); });
            var hideY2 = ['pie', 'doughnut', 'radar', 'scatter'];
            if (hideY2.indexOf(chartType) > -1) {
                document.getElementById('y2-axis-control').classList.add('control-hidden');
                document.getElementById('y2-color-control').classList.add('control-hidden');
                document.getElementById('y2-axis-max-control').classList.add('control-hidden');
            }
            var hideGrid = ['pie', 'doughnut'];
            if (hideGrid.indexOf(chartType) > -1) { document.getElementById('show-grid-control').classList.add('control-hidden'); }
        }

        function renderChart() {
            var btn = document.getElementById('update-charts-btn'), btnText = document.getElementById('btn-text'), btnSpinner = document.getElementById('btn-spinner');
            btn.disabled = true;
            btnText.textContent = 'Gerando...';
            btnSpinner.classList.remove('hidden');

            setTimeout(function() {
                if (chartInstance) { chartInstance.destroy(); }
                
                var container = document.querySelector('.chart-container');
                container.innerHTML = '<canvas id="mainChart"></canvas>';
                var ctx = document.getElementById('mainChart').getContext('2d');

                var chartType = document.querySelector('input[name="chart-type"]:checked').value;
                var uiConfig = {
                    showLabels: document.getElementById('show-labels').checked,
                    showGrid: document.getElementById('show-grid').checked,
                    isDarkMode: document.getElementById('dark-mode').checked,
                    xCol: document.getElementById('x-axis').value,
                    y1Col: document.getElementById('y1-axis').value,
                    y2Col: document.getElementById('y2-axis').value,
                    y1Color: document.getElementById('y1-color').value,
                    y2Color: document.getElementById('y2-color').value,
                    isY2Enabled: document.getElementById('y2-axis').value !== 'Nenhum',
                    yAxisAuto: document.getElementById('y-axis-auto').checked,
                    yAxisMax: document.getElementById('y-axis-max').value,
                    y2AxisAuto: document.getElementById('y2-axis-auto').checked,
                    y2AxisMax: document.getElementById('y2-axis-max').value
                };

                if (!uiConfig.xCol || !uiConfig.y1Col) {
                    btn.disabled = false;
                    btnText.textContent = 'Gerar / Atualizar Gráfico';
                    btnSpinner.classList.add('hidden');
                    return;
                }

                var labels = RAW_DATA.map(function(d) { return d[uiConfig.xCol]; });
                var dataY1 = RAW_DATA.map(function(d) { return parseNumericValue(d[uiConfig.y1Col]); });
                var dataY2 = uiConfig.isY2Enabled ? RAW_DATA.map(function(d) { return parseNumericValue(d[uiConfig.y2Col]); }) : [];

                updateKPIs(dataY1, dataY2, uiConfig.y1Col, uiConfig.y2Col, labels, uiConfig.isY2Enabled);
                updateControlVisibility();
                
                var chartTitle = document.querySelector('label[for="type-' + chartType + '"]').textContent;
                document.getElementById('chart-title').textContent = chartTitle + ": " + uiConfig.y1Col + " por " + uiConfig.xCol;

                var chartData = { labels: labels, datasets: [] };
                var chartOptions = buildChartOptions(chartType, uiConfig);
                
                var datasetsBuilder = {
                    'combo': function() { var ds = [{ type: 'bar', label: uiConfig.y1Col, data: dataY1, backgroundColor: uiConfig.y1Color + 'B3', yAxisID: 'y' }]; if (uiConfig.isY2Enabled) ds.push({ type: 'line', label: uiConfig.y2Col, data: dataY2, borderColor: uiConfig.y2Color, tension: 0.4, yAxisID: 'y1' }); return ds; },
                    'bar': function() { return [{ label: uiConfig.y1Col, data: dataY1, backgroundColor: uiConfig.y1Color }]; },
                    'line': function() { return [{ label: uiConfig.y1Col, data: dataY1, borderColor: uiConfig.y1Color, backgroundColor: uiConfig.y1Color + '33', fill: true, tension: 0.4 }]; },
                    'stacked': function() { var ds = [{ label: uiConfig.y1Col, data: dataY1, backgroundColor: uiConfig.y1Color }]; if (uiConfig.isY2Enabled) ds.push({ label: uiConfig.y2Col, data: dataY2, backgroundColor: uiConfig.y2Color }); return ds; },
                    'pie': function() { var colors = labels.map(function(_, i) { return 'hsl(' + (360 * i / labels.length) + ', 70%, 60%)'; }); return [{ label: uiConfig.y1Col, data: dataY1, backgroundColor: colors }]; },
                    'radar': function() { var ds = [{ label: uiConfig.y1Col, data: dataY1, borderColor: uiConfig.y1Color, backgroundColor: uiConfig.y1Color + '4D' }]; if (uiConfig.isY2Enabled) ds.push({ label: uiConfig.y2Col, data: dataY2, borderColor: uiConfig.y2Color, backgroundColor: uiConfig.y2Color + '4D' }); return ds; },
                    'scatter': function() { return [{ label: uiConfig.y1Col + ' vs ' + uiConfig.xCol, data: dataY1.map(function(val, idx) { return { x: labels[idx], y: val }; }), backgroundColor: uiConfig.y1Color }]; }
                };
                chartData.datasets = (datasetsBuilder[chartType] || datasetsBuilder['bar'])();
                if (chartType === 'doughnut') chartData.datasets = datasetsBuilder['pie']();

                var chartRealType = ['combo', 'stacked'].indexOf(chartType) > -1 ? 'bar' : chartType;
                
                chartInstance = new Chart(ctx, { type: chartRealType, data: chartData, options: chartOptions });

                btn.disabled = false;
                btnText.textContent = 'Gerar / Atualizar Gráfico';
                btnSpinner.classList.add('hidden');
            }, 100);
        }

        function buildChartOptions(chartType, config) {
            var fontColor = config.isDarkMode ? '#E2E8F0' : '#64748B';
            var gridColor = config.isDarkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';
            var options = { responsive: true, maintainAspectRatio: false, animation: { duration: 500 }, plugins: { legend: { position: 'bottom', labels: { color: fontColor } }, datalabels: { display: config.showLabels, color: config.isDarkMode ? '#FFFFFF' : '#334155', anchor: 'end', align: 'top', formatter: function(value) { return typeof value === 'object' ? value.y.toLocaleString('pt-BR') : value.toLocaleString('pt-BR'); }, font: { weight: 'bold' } } }, scales: {} };
            if (['pie', 'doughnut'].indexOf(chartType) > -1) { options.plugins.datalabels.align = 'center'; options.plugins.datalabels.color = 'white'; return options; }
            options.scales = { x: { grid: { display: config.showGrid, color: gridColor }, ticks: { color: fontColor } }, y: { grid: { display: config.showGrid, color: gridColor }, ticks: { color: fontColor }, beginAtZero: true, position: 'left' } };
            if (!config.yAxisAuto && config.yAxisMax) { options.scales.y.max = parseFloat(config.yAxisMax); }
            if (chartType === 'combo' && config.isY2Enabled) {
                options.scales.y1 = { display: true, position: 'right', grid: { drawOnChartArea: false }, beginAtZero: true, ticks: { color: fontColor } };
                if (!config.y2AxisAuto && config.y2AxisMax) { options.scales.y1.max = parseFloat(config.y2AxisMax); }
            }
            if (chartType === 'stacked') { options.scales.x.stacked = true; options.scales.y.stacked = true; }
            if (chartType === 'radar') { options.scales = { r: { grid: { display: config.showGrid, color: gridColor }, pointLabels: { color: fontColor }, angleLines: { color: gridColor } } }; }
            if (chartType === 'scatter') {
                 options.scales.x.type = 'category';
            }
            return options;
        }

        function downloadChart() { if (chartInstance) { var a = document.createElement('a'); a.href = chartInstance.toBase64Image(); a.download = 'PowerChart_Grafico.png'; a.click(); } }
        
        document.addEventListener('DOMContentLoaded', function() {
            try {
                var defaultSelections = findDefaultAxes();
                populateControls(defaultSelections);
                renderChart();

                document.getElementById('update-charts-btn').addEventListener('click', renderChart);
                document.getElementById('download-btn').addEventListener('click', downloadChart);
                
                document.querySelectorAll('input[name="chart-type"]').forEach(function(el) { el.addEventListener('change', renderChart); });

                function setupAxisControls(checkboxId, inputId) {
                    var checkbox = document.getElementById(checkboxId);
                    var input = document.getElementById(inputId);
                    checkbox.addEventListener('change', function(e) {
                        input.disabled = e.target.checked;
                        if (!e.target.checked) { if(input.value === '') input.value = 100; input.focus(); } 
                        else { input.value = ''; }
                    });
                }
                setupAxisControls('y-axis-auto', 'y-axis-max');
                setupAxisControls('y2-axis-auto', 'y2-axis-max');
            } catch (e) { }
        });
    </script>
</body>
</html>
"@
}

# --- 3. Construção da Interface Gráfica (Windows Forms) ---
# O código do editor PowerShell continua o mesmo.
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerChart Editor 15.0"
$Form.Width = 1200
$Form.Height = 800
$Form.StartPosition = "CenterScreen"

$MainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$MainLayout.Dock = "Fill"
$MainLayout.ColumnCount = 1
$MainLayout.RowCount = 2
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$Form.Controls.Add($MainLayout)

$ControlPanel = New-Object System.Windows.Forms.Panel
$ControlPanel.Dock = "Fill"
$ControlPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)
$ControlPanel.Padding = New-Object System.Windows.Forms.Padding(5)
$MainLayout.Controls.Add($ControlPanel, 0, 0)

$ButtonLoadCsv = New-Object System.Windows.Forms.Button
$ButtonLoadCsv.Text = "Carregar CSV"
$ButtonLoadCsv.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$ButtonLoadCsv.Location = New-Object System.Drawing.Point(10, 10)
$ButtonLoadCsv.Size = New-Object System.Drawing.Size(120, 30)
$ControlPanel.Controls.Add($ButtonLoadCsv)

$ButtonGenerateHtml = New-Object System.Windows.Forms.Button
$ButtonGenerateHtml.Text = "Gerar e Visualizar Relatório HTML"
$ButtonGenerateHtml.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ButtonGenerateHtml.Location = New-Object System.Drawing.Point(140, 10)
$ButtonGenerateHtml.Size = New-Object System.Drawing.Size(220, 30)
$ButtonGenerateHtml.Enabled = $false
$ControlPanel.Controls.Add($ButtonGenerateHtml)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Aguardando arquivo CSV..."
$StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$StatusLabel.Location = New-Object System.Drawing.Point(370, 15)
$StatusLabel.AutoSize = $true
$ControlPanel.Controls.Add($StatusLabel)

$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Dock = "Fill"
$DataGridView.BackgroundColor = [System.Drawing.Color]::White
$DataGridView.BorderStyle = "None"
$DataGridView.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$MainLayout.Controls.Add($DataGridView, 0, 1)

# --- 4. Eventos ---
$ButtonLoadCsv.Add_Click({
    Load-CSVData -DataGridView $DataGridView -StatusLabel $StatusLabel -GenerateButton $ButtonGenerateHtml
})

$ButtonGenerateHtml.Add_Click({
    Generate-HtmlReport -DataGridView $DataGridView -StatusLabel $StatusLabel
})

# --- 5. Exibir a Janela ---
$Form.ShowDialog() | Out-Null

