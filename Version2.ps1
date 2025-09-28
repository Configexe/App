# -----------------------------------------------------------------------------
# Power-Graphx Editor: Análise e Visualização de Dados
# Versão: 17.4 - Edição Final com Formatação Completa
# Autor: Seu Nome/Empresa
# Descrição: Adiciona controles avançados para rótulos de dados (posição,
#            tamanho), controle manual da escala do Eixo Y e reintroduz o
#            gráfico de Barras Empilhadas.
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

# --- 2. Funções Auxiliares ---

Function Show-InputBox {
    param(
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultText = ""
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Font = 'Segoe UI, 9'
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object System.Drawing.Size(350, 120)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(10, 15)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $DefaultText
    $textBox.Location = New-Object System.Drawing.Point(12, 40)
    $textBox.Size = New-Object System.Drawing.Size(326, 23)
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(182, 75)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(263, 75)
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    if ($form.ShowDialog() -eq 'OK') {
        return $textBox.Text
    }
    return $null
}


# --- 3. Funções Principais ---

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

    $ColumnStructure = $DataGridView.Columns | ForEach-Object {
        [PSCustomObject]@{
            OriginalName = $_.DataPropertyName
            DisplayName  = $_.HeaderText
        }
    }

    $JsonData = $DataForJson | ConvertTo-Json -Compress -Depth 5
    $JsonColumnStructure = $ColumnStructure | ConvertTo-Json -Compress

    $OutputPath = Join-Path $env:TEMP "PowerGraphx_Relatorio.html"
    $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -JsonColumnStructure $JsonColumnStructure
    
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
    param($JsonData, $JsonColumnStructure)

    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power-Graphx - Relatório Dinâmico</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        .card { background-color: white; border-radius: 0.75rem; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1); padding: 1.5rem; transition: all 0.3s ease-in-out; }
        .chart-container { position: relative; width: 100%; height: 650px; }
        .chart-selector label { border: 2px solid #e5e7eb; border-radius: 0.5rem; padding: 0.5rem; cursor: pointer; transition: all 0.2s ease-in-out; text-align: center; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; }
        .chart-selector label:hover { border-color: #9ca3af; background-color: #f9fafb; }
        .chart-selector input:checked + label { border-color: #3b82f6; background-color: #eff6ff; box-shadow: 0 0 0 2px #3b82f6; }
        .chart-selector input { display: none; }
        .divider { border-top: 1px solid #e5e7eb; }
    </style>
</head>
<body class="text-gray-900">
    <header class="bg-[#0f172a] text-white text-center py-12 px-4">
        <h1 class="text-4xl md:text-5xl font-black tracking-tight">Relatório Dinâmico Interativo</h1>
        <p class="mt-4 text-lg text-blue-200 max-w-3xl mx-auto">Dados processados via Power-Graphx Editor.</p>
    </header>
    <main class="container mx-auto p-4 md:p-8 -mt-10">
        <section id="controls" class="card mb-6">
             <div class="grid grid-cols-1 lg:grid-cols-2 gap-x-8 gap-y-6">
                <div>
                    <h2 class="text-xl font-bold text-[#1e293b] mb-4">1. Seleção de Dados</h2>
                    <div class="mt-4 pt-4 border-t">
                        <div class="flex justify-between items-center mb-2">
                             <h3 class="text-lg font-bold text-[#1e293b]">Séries de Dados (Eixos)</h3>
                             <button id="add-series-btn" class="bg-blue-500 text-white text-xs font-bold py-1 px-3 rounded-full hover:bg-blue-600 transition">+ Adicionar Série</button>
                        </div>
                        <div id="series-container" class="space-y-3"></div>
                    </div>
                </div>
                 <div class="flex flex-col justify-between">
                     <div>
                         <h2 class="text-xl font-bold text-[#1e293b] mb-4">2. Escolha o Tipo de Gráfico</h2>
                         <div class="chart-selector grid grid-cols-3 sm:grid-cols-6 gap-2">
                            <div><input type="radio" name="chart-type" value="combo" id="type-combo" checked><label for="type-combo"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/><path d="M3 12l5-4 5 6 5-4"/></svg><span class="text-xs font-semibold">Combo</span></label></div>
                            <div><input type="radio" name="chart-type" value="stacked" id="type-stacked"><label for="type-stacked"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="#3b82f6" stroke="#3b82f6" stroke-width="1"><rect x="5" y="12" width="4" height="6"/><rect x="10" y="8" width="4" height="10"/><rect x="15" y="4" width="4" height="14"/><path d="M5 12V9h4v3m1-4V4h4v4m1-4V2h4v2" fill="#ef4444"/></svg><span class="text-xs font-semibold">Empilhado</span></label></div>
                            <div><input type="radio" name="chart-type" value="line" id="type-line"><label for="type-line"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M3 17l5-4 5 6 5-4 4 2"/></svg><span class="text-xs font-semibold">Linha</span></label></div>
                            <div><input type="radio" name="chart-type" value="scatter" id="type-scatter"><label for="type-scatter"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="7" cy="7" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="17" cy="17" r="1"/><circle cx="7" cy="17" r="1"/><circle cx="17" cy="7" r="1"/></svg><span class="text-xs font-semibold">Dispersão</span></label></div>
                            <div><input type="radio" name="chart-type" value="bubble" id="type-bubble"><label for="type-bubble"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="8" r="4"/><circle cx="16" cy="16" r="6"/><circle cx="17" cy="7" r="2"/></svg><span class="text-xs font-semibold">Bolhas</span></label></div>
                            <div><input type="radio" name="chart-type" value="pie" id="type-pie"><label for="type-pie"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/></svg><span class="text-xs font-semibold">Pizza</span></label></div>
                         </div>
                     </div>
                      <div class="mt-6"><button id="update-charts-btn" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg transition hover:bg-blue-700 flex items-center justify-center text-lg">
                          Gerar / Atualizar Gráfico
                      </button></div>
                 </div>
             </div>
        </section>
        
        <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">
            <div id="chart-card" class="lg:col-span-3 card transition-all duration-300">
                 <div class="flex justify-between items-center mb-4">
                     <h3 id="chart-title" class="text-xl font-bold text-[#1e293b]"></h3>
                 </div>
                 <div class="chart-container"><canvas id="mainChart"></canvas></div>
            </div>
            <div id="format-panel" class="lg:col-span-1 card">
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
                        <div id="label-options" class="mt-2 space-y-2 hidden">
                            <div>
                                <label for="label-position" class="text-xs text-gray-600">Posição:</label>
                                <select id="label-position" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm">
                                    <option value="top">Topo</option>
                                    <option value="center">Centro</option>
                                    <option value="bottom">Base</option>
                                </select>
                            </div>
                            <div>
                                <label for="label-size" class="text-xs text-gray-600">Tamanho Fonte:</label>
                                <input type="number" id="label-size" value="12" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm">
                            </div>
                        </div>
                    </div>
                    <div class="divider"></div>
                    <div>
                        <span class="font-semibold text-gray-700 text-sm">Linhas de Grade</span>
                        <div class="flex items-center mt-2"><input id="show-grid" type="checkbox" checked class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="show-grid" class="ml-2 block text-sm text-gray-900">Exibir grades</label></div>
                    </div>
                    <div class="divider"></div>
                    <div id="y-axis-max-control">
                        <span class="font-semibold text-gray-700 text-sm">Eixo Y (Primário)</span>
                        <div class="flex items-center mt-2"><input id="y-axis-auto" type="checkbox" checked class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="y-axis-auto" class="ml-2 block text-sm text-gray-900">Automático</label></div>
                        <input type="number" id="y-axis-max" placeholder="Ex: 100" disabled class="mt-2 block w-full rounded-md border-gray-300 shadow-sm text-sm disabled:bg-gray-100">
                    </div>
                </div>
            </div>
        </div>
    </main>
    <script>
        // --- INÍCIO DO JAVASCRIPT CORRIGIDO (VERSÃO 17.4) ---
        let chartInstance;
        let seriesCounter = 0;
        const seriesColors = ["#3b82f6", "#ef4444", "#22c55e", "#f97316", "#8b5cf6", "#14b8a6"];
        
        function initializeApp(rawData, columnStructure) {
            if (!rawData || !columnStructure) {
                console.error("Dados ou estrutura de colunas não fornecidos.");
                return;
            }
            Chart.register(ChartDataLabels);

            function parseValue(value) {
                if (typeof value === 'number') return value;
                if (typeof value !== 'string') return value;
                const cleanValue = value.replace(/[^0-9,-]/g, '').replace(',', '.');
                const parsed = parseFloat(cleanValue);
                return isNaN(parsed) ? value : parsed;
            }
            
            function isNumeric(colName) {
                if (rawData.length === 0) return false;
                const sampleValue = rawData[0][colName];
                if (sampleValue === null || typeof sampleValue === 'undefined') return false;
                const parsed = parseFloat(String(sampleValue).replace(',', '.'));
                return !isNaN(parsed) && String(sampleValue).trim() !== '';
            }

            function populateSelect(selectElement, type = 'all') {
                if (!selectElement) return;
                selectElement.innerHTML = '';
                const options = columnStructure.filter(col => {
                    if (type === 'numeric') return isNumeric(col.OriginalName);
                    if (type === 'text') return !isNumeric(col.OriginalName);
                    return true;
                });
                
                options.forEach(col => {
                    const option = document.createElement('option');
                    option.value = col.OriginalName;
                    option.textContent = col.DisplayName;
                    selectElement.appendChild(option);
                });
            }

            function addSeriesControl(isFirst = false) {
                const seriesContainer = document.getElementById('series-container');
                const seriesId = ++seriesCounter;
                const defaultColor = seriesColors[(seriesId - 1) % seriesColors.length];
                const seriesDiv = document.createElement('div');
                seriesDiv.id = 'series-' + seriesId;
                seriesDiv.className = 'p-3 border rounded-lg bg-gray-50 grid grid-cols-1 sm:grid-cols-2 gap-3 items-end';
                
                let content = '';
                const eixoXLabel = isFirst ? 'Eixo X:' : 'Eixo X' + seriesId + ':';
                const eixoYLabel = isFirst ? 'Eixo Y:' : 'Eixo Y' + seriesId + ':';
                
                content += '<div class="x-axis-control"><label class="text-xs font-semibold">' + eixoXLabel + '</label><select name="x-axis" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm"></select></div>';
                content += '<div><label class="text-xs font-semibold">' + eixoYLabel + '</label><select name="y-axis" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm y-axis-select"></select></div>';
                content += '<div class="size-axis-control" style="display: none;"><label class="text-xs font-semibold">Tamanho (Bolha):</label><select name="size-axis" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm"></select></div>';
                content += '<div class="combo-type-control"><label class="text-xs font-semibold">Tipo:</label><select name="series-type" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm"><option value="bar">Barra</option><option value="line">Linha</option></select></div>';
                content += '<div class="flex items-end space-x-2"><div class="w-full"><label class="text-xs font-semibold">Cor:</label><input type="color" value="' + defaultColor + '" name="color" class="axis-select mt-1 w-full h-9"></div>';
                if (!isFirst) {
                    content += '<button type="button" onclick="this.parentElement.parentElement.remove(); renderChart();" class="h-9 px-3 bg-red-500 text-white rounded-md hover:bg-red-600 transition text-sm font-bold">&times;</button>';
                }
                content += '</div>';

                seriesDiv.innerHTML = content;
                seriesContainer.appendChild(seriesDiv);
                
                populateSelect(seriesDiv.querySelector('[name="x-axis"]'), 'all');
                populateSelect(seriesDiv.querySelector('[name="y-axis"]'), 'numeric');
                populateSelect(seriesDiv.querySelector('[name="size-axis"]'), 'numeric');
            }

            window.renderChart = function() {
                if (chartInstance) { chartInstance.destroy(); }
                const container = document.querySelector('.chart-container');
                container.innerHTML = '<canvas id="mainChart"></canvas>';
                const ctx = document.getElementById('mainChart').getContext('2d');
                
                const chartType = document.querySelector('input[name="chart-type"]:checked').value;
                const seriesControls = document.querySelectorAll('#series-container > div');
                if (seriesControls.length === 0) return;

                const chartData = { datasets: [] };
                let isCategorical = !['scatter', 'bubble'].includes(chartType);

                if (isCategorical) {
                    const firstXAxisSelect = seriesControls[0] ? seriesControls[0].querySelector('[name="x-axis"]') : null;
                    if (firstXAxisSelect && firstXAxisSelect.value) {
                         chartData.labels = rawData.map(d => d[firstXAxisSelect.value]);
                    }
                }
                
                seriesControls.forEach((control) => {
                    const yColSelect = control.querySelector('[name="y-axis"]');
                    const xColSelect = control.querySelector('[name="x-axis"]');
                    const colorInput = control.querySelector('[name="color"]');
                    const typeSelect = control.querySelector('[name="series-type"]');
                    
                    if (!yColSelect || !yColSelect.value || !xColSelect || !xColSelect.value || !colorInput) return;
                    
                    const yCol = yColSelect.value;
                    const xCol = xColSelect.value;
                    const colInfo = columnStructure.find(c => c.OriginalName === yCol);
                    if (!colInfo) return;

                    const dataset = {
                        label: colInfo.DisplayName,
                        borderColor: colorInput.value,
                        backgroundColor: colorInput.value + 'B3',
                        type: (chartType === 'combo' && typeSelect) ? typeSelect.value : (chartType === 'stacked' ? 'bar' : undefined),
                    };
                    
                    if (chartType === 'line') dataset.type = 'line';

                    if (chartType === 'pie') {
                        dataset.data = rawData.map(d => parseValue(d[yCol]));
                        if (chartData.labels) {
                            dataset.backgroundColor = chartData.labels.map((_, i) => seriesColors[i % seriesColors.length]);
                        }
                    } else if (isCategorical) {
                        dataset.data = rawData.map(d => parseValue(d[yCol]));
                    } else {
                        const sizeColSelect = control.querySelector('[name="size-axis"]');
                        dataset.data = rawData.map(d => ({
                            x: parseValue(d[xCol]),
                            y: parseValue(d[yCol]),
                            r: (chartType === 'bubble' && sizeColSelect && sizeColSelect.value) ? parseValue(d[sizeColSelect.value]) : undefined
                        }));
                    }
                    chartData.datasets.push(dataset);
                });
                
                if (chartData.datasets.length === 0) return;
                
                const finalChartType = ['combo', 'stacked'].includes(chartType) ? 'bar' : chartType;
                const options = buildChartOptions(chartType, isCategorical);
                chartInstance = new Chart(ctx, { type: finalChartType, data: chartData, options: options });
            }

            function buildChartOptions(chartType, isCategorical) {
                const darkMode = document.getElementById('dark-mode').checked;
                const showGrid = document.getElementById('show-grid').checked;
                const showLabels = document.getElementById('show-labels').checked;
                const labelPos = document.getElementById('label-position').value;
                const labelSize = document.getElementById('label-size').value;
                const yAxisAuto = document.getElementById('y-axis-auto').checked;
                const yAxisMax = parseFloat(document.getElementById('y-axis-max').value);

                const fontColor = darkMode ? '#E2E8F0' : '#64748B';
                const gridColor = darkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';

                const options = {
                    responsive: true, maintainAspectRatio: false,
                    plugins: { 
                        legend: { labels: { color: fontColor } },
                        datalabels: { 
                            display: showLabels, 
                            color: fontColor,
                            font: { size: labelSize || 12 },
                            align: labelPos === 'bottom' ? 'bottom' : (labelPos === 'center' ? 'center' : 'top'),
                            anchor: labelPos === 'bottom' ? 'start' : (labelPos === 'center' ? 'center' : 'end'),
                            formatter: (value) => typeof value === 'object' ? value.y : value
                        }
                    },
                    scales: {}
                };

                if (chartType === 'pie') return options;
                
                const axisOptions = { grid: { display: showGrid, color: gridColor }, ticks: { color: fontColor } };
                options.scales.x = { ...axisOptions, type: isCategorical ? 'category' : 'linear' };
                options.scales.y = { ...axisOptions, beginAtZero: true };

                if (!yAxisAuto && !isNaN(yAxisMax)) {
                    options.scales.y.max = yAxisMax;
                }

                if (chartType === 'stacked') {
                    options.scales.x.stacked = true;
                    options.scales.y.stacked = true;
                }

                return options;
            }

            function updateUI() {
                const chartType = document.querySelector('input[name="chart-type"]:checked').value;
                const isScatterOrBubble = ['scatter', 'bubble'].includes(chartType);
                
                document.querySelectorAll('#series-container > div').forEach((sc, i) => {
                    if (!sc) return;
                    const xAxisControl = sc.querySelector('.x-axis-control');
                    const sizeAxisControl = sc.querySelector('.size-axis-control');
                    const comboTypeControl = sc.querySelector('.combo-type-control');

                    if (xAxisControl) xAxisControl.style.display = (isScatterOrBubble || i === 0) ? 'block' : 'none';
                    if (sizeAxisControl) sizeAxisControl.style.display = chartType === 'bubble' ? 'block' : 'none';
                    if (comboTypeControl) comboTypeControl.style.display = chartType === 'combo' ? 'block' : 'none';
                });
                
                document.getElementById('y-axis-max-control').style.display = chartType === 'pie' ? 'none' : 'block';
                
                renderChart();
            }
            
            // --- Inicialização e Eventos ---
            addSeriesControl(true);
            document.getElementById('controls').addEventListener('change', updateUI);
            document.getElementById('format-panel').addEventListener('change', renderChart);
            document.getElementById('add-series-btn').addEventListener('click', () => { addSeriesControl(false); updateUI(); });
            document.getElementById('update-charts-btn').addEventListener('click', renderChart);
            
            document.getElementById('show-labels').addEventListener('change', (e) => {
                document.getElementById('label-options').style.display = e.target.checked ? 'block' : 'none';
            });
            document.getElementById('y-axis-auto').addEventListener('change', (e) => {
                document.getElementById('y-axis-max').disabled = e.target.checked;
            });
            
            updateUI();
        }

        document.addEventListener('DOMContentLoaded', function() {
            try {
                initializeApp($JsonData, $JsonColumnStructure);
            } catch (e) {
                console.error("Erro fatal ao inicializar o Power-Graphx:", e);
                document.body.innerHTML = '<div class="text-center p-8 bg-red-100 text-red-700"><h1>Ocorreu um erro crítico</h1><p>Não foi possível renderizar o relatório. Verifique o console para mais detalhes.</p></div>';
            }
        });
    </script>
</body>
</html>
"@
}

# --- 4. Construção da Interface Gráfica (Windows Forms) ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Power-Graphx Editor 17.4"
$Form.Width = 1200
$Form.Height = 800
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = 'Sizable'
$Form.WindowState = 'Maximized'
try { $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\imageres.dll,25") } catch {}


$MainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$MainLayout.Dock = "Fill"
$MainLayout.ColumnCount = 1
$MainLayout.RowCount = 3
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$Form.Controls.Add($MainLayout)

# --- Painel de Controles Superior ---
$ControlPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$ControlPanel.Dock = "Fill"
$ControlPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)
$ControlPanel.Padding = New-Object System.Windows.Forms.Padding(5)
$MainLayout.Controls.Add($ControlPanel, 0, 0)

$ButtonLoadCsv = New-Object System.Windows.Forms.Button
$ButtonLoadCsv.Text = "Carregar CSV"
$ButtonLoadCsv.Font = "Segoe UI, 9"
$ButtonLoadCsv.Size = New-Object System.Drawing.Size(120, 30)
$ControlPanel.Controls.Add($ButtonLoadCsv)

$ButtonGenerateHtml = New-Object System.Windows.Forms.Button
$ButtonGenerateHtml.Text = "Gerar e Visualizar Relatório"
$ButtonGenerateHtml.Font = "Segoe UI, 9, Bold"
$ButtonGenerateHtml.Size = New-Object System.Drawing.Size(200, 30)
$ButtonGenerateHtml.Enabled = $false
$ControlPanel.Controls.Add($ButtonGenerateHtml)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Aguardando arquivo CSV..."
$StatusLabel.Font = "Segoe UI, 9"
$StatusLabel.Margin = "5,5,0,0"
$StatusLabel.TextAlign = "MiddleLeft"
$StatusLabel.AutoSize = $true
$ControlPanel.Controls.Add($StatusLabel)

# --- Painel de Busca ---
$SearchPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$SearchPanel.Dock = "Fill"
$SearchPanel.BackColor = [System.Drawing.Color]::FromArgb(220, 225, 230)
$SearchPanel.Visible = $false
$SearchPanel.Padding = 5
$MainLayout.Controls.Add($SearchPanel, 0, 1)

$SearchLabel = New-Object System.Windows.Forms.Label
$SearchLabel.Text = "Buscar:"
$SearchLabel.Font = "Segoe UI, 9"
$SearchLabel.Margin = "0,3,0,0"
$SearchPanel.Controls.Add($SearchLabel)

$SearchTextBox = New-Object System.Windows.Forms.TextBox
$SearchTextBox.Size = New-Object System.Drawing.Size(250, 23)
$SearchPanel.Controls.Add($SearchTextBox)

$SearchButton = New-Object System.Windows.Forms.Button
$SearchButton.Text = "Buscar"
$SearchButton.Size = New-Object System.Drawing.Size(75, 25)
$SearchPanel.Controls.Add($SearchButton)

$CloseSearchButton = New-Object System.Windows.Forms.Button
$CloseSearchButton.Text = "Fechar"
$CloseSearchButton.Size = New-Object System.Drawing.Size(75, 25)
$SearchPanel.Controls.Add($CloseSearchButton)

$SearchResultLabel = New-Object System.Windows.Forms.Label
$SearchResultLabel.Text = ""
$SearchResultLabel.Font = "Segoe UI, 9"
$SearchResultLabel.Margin = "10,3,0,0"
$SearchPanel.Controls.Add($SearchResultLabel)

# --- Data Grid View ---
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Dock = "Fill"
$DataGridView.BackgroundColor = [System.Drawing.Color]::White
$DataGridView.BorderStyle = "None"
$DataGridView.ColumnHeadersDefaultCellStyle.Font = "Segoe UI, 9, Bold"
$DataGridView.ReadOnly = $true
$DataGridView.AllowUserToAddRows = $false
$MainLayout.Controls.Add($DataGridView, 0, 2)

# --- Menu de Contexto para Renomear Coluna ---
$ContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$RenameMenuItem = $ContextMenu.Items.Add("Renomear Coluna...")
$Global:ColumnToRenameIndex = -1

# --- 5. Eventos ---
$ButtonLoadCsv.Add_Click({
    Load-CSVData -DataGridView $DataGridView -StatusLabel $StatusLabel -GenerateButton $ButtonGenerateHtml
})

$ButtonGenerateHtml.Add_Click({
    Generate-HtmlReport -DataGridView $DataGridView -StatusLabel $StatusLabel
})

$DataGridView.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $Global:ColumnToRenameIndex = $e.ColumnIndex
        $ContextMenu.Show($DataGridView, $DataGridView.PointToClient([System.Windows.Forms.Cursor]::Position))
    }
})

$RenameMenuItem.Add_Click({
    if ($Global:ColumnToRenameIndex -ge 0) {
        $column = $DataGridView.Columns[$Global:ColumnToRenameIndex]
        $newName = Show-InputBox -Title "Renomear Coluna" -Prompt "Digite o novo nome para a coluna '$($column.HeaderText)':'" -DefaultText $column.HeaderText
        if ($newName) {
            $column.HeaderText = $newName
        }
    }
})

$Form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq 'F') {
        $SearchPanel.Visible = !$SearchPanel.Visible
        if ($SearchPanel.Visible) { $SearchTextBox.Focus() }
    }
})

$CloseSearchButton.Add_Click({ $SearchPanel.Visible = $false })

$SearchButton.Add_Click({
    $searchTerm = $SearchTextBox.Text.ToLower()
    if ([string]::IsNullOrWhiteSpace($searchTerm)) { return }

    $DataGridView.ClearSelection()
    $defaultCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $highlightCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $highlightCellStyle.BackColor = [System.Drawing.Color]::Yellow

    $foundCount = 0
    foreach ($row in $DataGridView.Rows) {
        foreach ($cell in $row.Cells) {
            $cell.Style = $defaultCellStyle
            if ($cell.Value -and $cell.Value.ToString().ToLower().Contains($searchTerm)) {
                $cell.Style = $highlightCellStyle
                $foundCount++
            }
        }
    }
    $SearchResultLabel.Text = "$foundCount ocorrência(s) encontrada(s)."
})


# --- 6. Exibir a Janela ---
$Form.ShowDialog() | Out-Null

