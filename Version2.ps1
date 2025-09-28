# -----------------------------------------------------------------------------
# PowerChart Designer: Editor de Dados
# Versão: 10.0 - Edição Power User
# Autor: Seu Nome/Empresa
# Descrição: Gera um relatório HTML com 8 tipos de gráficos, opção para
#            download de imagem, e uma interface dinâmica que se adapta
#            ao tipo de gráfico selecionado, melhorando a experiência do usuário.
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

# Função para carregar dados do CSV para a grade
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

# Função para gerar o relatório HTML e abri-lo
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

    # Converte os dados da grade para JSON
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

# Template do HTML
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
        .chart-container { position: relative; width: 100%; height: 550px; }
        .chart-selector label { border: 2px solid #e5e7eb; border-radius: 0.5rem; padding: 0.75rem; cursor: pointer; transition: all 0.2s ease-in-out; text-align: center; }
        .chart-selector label:hover { border-color: #9ca3af; background-color: #f9fafb; }
        .chart-selector input:checked + label { border-color: #3b82f6; background-color: #eff6ff; box-shadow: 0 0 0 2px #3b82f6; }
        .chart-selector input { display: none; }
        .control-hidden { display: none !important; }
    </style>
</head>
<body class="text-gray-900">
    <header class="bg-[#0f172a] text-white text-center py-12 px-4">
        <h1 class="text-4xl md:text-5xl font-black tracking-tight">Relatório Dinâmico Interativo</h1>
        <p class="mt-4 text-lg text-blue-200 max-w-3xl mx-auto">Dados processados via PowerChart Editor.</p>
    </header>
    <main class="container mx-auto p-4 md:p-8 -mt-10">
        <section id="controls" class="card mb-6">
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div class="lg:col-span-2">
                    <h2 class="text-xl font-bold text-[#1e293b] mb-4">1. Seleção de Dados</h2>
                    <div class="grid grid-cols-1 md:grid-cols-2 gap-x-6 gap-y-4 items-end">
                        <div><label for="x-axis">Eixo X:</label><select id="x-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                        <div><label for="y1-axis">Série Y1:</label><select id="y1-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                        <div id="y2-axis-control"><label for="y2-axis">Série Y2:</label><select id="y2-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                        <div><label for="y1-color">Cor Y1:</label><input type="color" id="y1-color" value="#3b82f6" class="w-full h-10 mt-1"></div>
                        <div id="y2-color-control"><label for="y2-color">Cor Y2:</label><input type="color" id="y2-color" value="#ef4444" class="w-full h-10 mt-1"></div>
                    </div>
                </div>
                 <div>
                    <h2 class="text-xl font-bold text-[#1e293b] mb-4">2. Opções de Estilo</h2>
                    <div class="space-y-3">
                        <div class="flex items-center"><input id="show-labels" type="checkbox" class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="show-labels" class="ml-2 block text-sm text-gray-900">Exibir Rótulos de Dados</label></div>
                        <div id="show-grid-control" class="flex items-center"><input id="show-grid" type="checkbox" checked class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="show-grid" class="ml-2 block text-sm text-gray-900">Exibir Linhas de Grade</label></div>
                    </div>
                    <div class="mt-6"><button id="update-charts-btn" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg transition hover:bg-blue-700">Atualizar Gráfico</button></div>
                </div>
            </div>
        </section>

        <section class="card mb-6">
            <h2 class="text-xl font-bold text-[#1e293b] mb-4">3. Escolha o Tipo de Gráfico</h2>
            <div class="chart-selector grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-4">
                <!-- Linha 1 -->
                <div><input type="radio" name="chart-type" value="combo" id="type-combo" checked><label for="type-combo" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/><path d="M3 12l5-4 5 6 5-4"/></svg><span class="text-sm font-semibold">Combo</span></label></div>
                <div><input type="radio" name="chart-type" value="bar" id="type-bar"><label for="type-bar" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg><span class="text-sm font-semibold">Barras</span></label></div>
                <div><input type="radio" name="chart-type" value="line" id="type-line"><label for="type-line" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M3 17l5-4 5 6 5-4 4 2"/></svg><span class="text-sm font-semibold">Linha</span></label></div>
                <div><input type="radio" name="chart-type" value="stacked" id="type-stacked"><label for="type-stacked" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="#3b82f6" stroke="#3b82f6" stroke-width="1"><rect x="5" y="12" width="4" height="6"/><rect x="10" y="8" width="4" height="10"/><rect x="15" y="4" width="4" height="14"/><path d="M5 12V9h4v3m1-4V4h4v4m1-4V2h4v2" fill="#ef4444"/></svg><span class="text-sm font-semibold">Empilhado</span></label></div>
                <!-- Linha 2 -->
                <div><input type="radio" name="chart-type" value="pie" id="type-pie"><label for="type-pie" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/></svg><span class="text-sm font-semibold">Pizza</span></label></div>
                <div><input type="radio" name="chart-type" value="doughnut" id="type-doughnut"><label for="type-doughnut" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="4"/></svg><span class="text-sm font-semibold">Rosca</span></label></div>
                <div><input type="radio" name="chart-type" value="radar" id="type-radar"><label for="type-radar" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2l10 6.5-4 11.5H6L2 8.5z"/><path d="M12 2v20"/><path d="m2 8.5 20 0"/><path d="m6 20 4-11.5L22 8.5"/></svg><span class="text-sm font-semibold">Radar</span></label></div>
                <div><input type="radio" name="chart-type" value="scatter" id="type-scatter"><label for="type-scatter" class="flex flex-col items-center justify-center h-full"><svg class="w-10 h-10 mb-2" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="16" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="16" cy="8" r="1"/><path d="M3 3v18h18"/></svg><span class="text-sm font-semibold">Dispersão</span></label></div>
            </div>
        </section>

        <section id="kpis" class="mb-6"><div id="kpi-grid" class="grid grid-cols-1 md:grid-cols-3 gap-6"></div></section>
        
        <section class="card">
             <div class="flex justify-between items-center mb-4">
                <h3 id="chart-title" class="text-xl font-bold text-[#1e293b]"></h3>
                <button id="download-btn" class="bg-gray-100 text-gray-700 hover:bg-gray-200 font-bold py-2 px-4 rounded-lg transition text-sm flex items-center">
                    <svg class="w-4 h-4 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                    Baixar (PNG)
                </button>
            </div>
            <div class="chart-container"><canvas id="mainChart"></canvas></div>
        </section>
    </main>
    <script>
        // O código JS foi atualizado para suportar todas as novas funcionalidades.
        // O restante do script PowerShell continua abaixo.
        var RAW_DATA=$JsonData,COLUMN_NAMES=$JsonColumnNames,chartInstance;Chart.register(ChartDataLabels);function parseNumericValue(a){if("number"==typeof a)return a;if("string"!=typeof a)return 0;var b=a.replace(/[^0-9,-]/g,"").replace(",","."),c=parseFloat(b);return isNaN(c)?0:c}
        function findDefaultAxes(){var a={xAxis:null,y1Axis:null,y2Axis:null},b=[],c=[];if(!RAW_DATA||0===RAW_DATA.length)return a;var d=RAW_DATA[0];for(var e=0;e<COLUMN_NAMES.length;e++){var f=COLUMN_NAMES[e],g=d[f];if(null==g)continue;var h=parseFloat(String(g).replace(",","."));isNaN(h)||""===String(g).trim()?c.push(f):b.push(f)}return a.xAxis=c[0]||COLUMN_NAMES[0]||null,a.y1Axis=b[0]||(COLUMN_NAMES.length>1?COLUMN_NAMES[1]:null),a.y2Axis=b[1]||(COLUMN_NAMES.length>2?COLUMN_NAMES[2]:null),a}
        function populateControls(a){["x-axis","y1-axis","y2-axis"].forEach(function(b){var c=document.getElementById(b);c.innerHTML="y2-axis"===b?'<option value="Nenhum">Nenhum</option>':"",COLUMN_NAMES.forEach(function(a){c.innerHTML+='<option value="'+a+'">'+a+"</option>"})}),document.getElementById("x-axis").value=a.xAxis||"",document.getElementById("y1-axis").value=a.y1Axis||"",document.getElementById("y2-axis").value=a.y2Axis||"Nenhum"}
        function updateControlVisibility(){var a=document.querySelector('input[name="chart-type"]:checked').value;["y2-axis-control","y2-color-control","show-grid-control"].forEach(function(a){document.getElementById(a).classList.remove("control-hidden")});var b=["pie","doughnut","radar"];b.includes(a)&&(document.getElementById("y2-axis-control").classList.add("control-hidden"),document.getElementById("y2-color-control").classList.add("control-hidden")),"pie"!==a&&"doughnut"!==a||document.getElementById("show-grid-control").classList.add("control-hidden")}
        function renderChart(){chartInstance&&chartInstance.destroy();var a=document.querySelector('input[name="chart-type"]:checked').value,b=document.getElementById("show-labels").checked,c=document.getElementById("show-grid").checked,d=document.getElementById("x-axis"),e=document.getElementById("y1-axis"),f=document.getElementById("y2-axis").value,g=document.getElementById("y1-color").value,h=document.getElementById("y2-color").value,i="Nenhum"!==f;if(!d.value||!e.value)return;var j=RAW_DATA.map(function(a){return a[d.value]}),k=RAW_DATA.map(function(a){return parseNumericValue(a[e.value])}),l=i?RAW_DATA.map(function(a){return parseNumericValue(a[f])}):[];updateKPIs(k,l,e.value,f,j,i);var m=document.getElementById("mainChart").getContext("2d"),n=document.querySelector('label[for="type-'+a+'"]').textContent,o=e.value+" por "+d.value;document.getElementById("chart-title").textContent=n+": "+o;var p,q,r={responsive:!0,maintainAspectRatio:!1,plugins:{legend:{position:"bottom"},datalabels:{display:b,anchor:"end",align:"top",formatter:function(a){return a.toLocaleString("pt-BR")},font:{weight:"bold"}}}};switch(a){case"combo":r.scales={x:{grid:{display:c}},y:{grid:{display:c},beginAtZero:!0,position:"left"}},p=[{type:"bar",label:e.value,data:k,backgroundColor:g+"B3",yAxisID:"y"}],i&&(p.push({type:"line",label:f,data:l,borderColor:h,tension:.4,yAxisID:"y1"}),r.scales.y1={display:i,position:"right",grid:{drawOnChartArea:!1},beginAtZero:!0}),chartInstance=new Chart(m,{data:{labels:j,datasets:p},options:r});break;case"bar":r.scales={x:{grid:{display:c}},y:{grid:{display:c},beginAtZero:!0}},p=[{label:e.value,data:k,backgroundColor:g}],chartInstance=new Chart(m,{type:"bar",data:{labels:j,datasets:p},options:r});break;case"line":r.scales={x:{grid:{display:c}},y:{grid:{display:c},beginAtZero:!0}},p=[{label:e.value,data:k,borderColor:g,backgroundColor:g+"33",fill:!0,tension:.4}],chartInstance=new Chart(m,{type:"line",data:{labels:j,datasets:p},options:r});break;case"stacked":r.scales={x:{stacked:!0,grid:{display:c}},y:{stacked:!0,grid:{display:c},beginAtZero:!0}},p=[{label:e.value,data:k,backgroundColor:g}],i&&p.push({label:f,data:l,backgroundColor:h}),chartInstance=new Chart(m,{type:"bar",data:{labels:j,datasets:p},options:r});break;case"pie":case"doughnut":var s=j.map(function(a,b){return"hsl("+(360*b/j.length)+", 70%, 60%)"});r.plugins.datalabels.align="center",r.plugins.datalabels.color="white",p=[{label:e.value,data:k,backgroundColor:s}],chartInstance=new Chart(m,{type:"doughnut"===a?"doughnut":"pie",data:{labels:j,datasets:p},options:r});break;case"radar":r.scales={r:{grid:{display:c}}},p=[{label:e.value,data:k,borderColor:g,backgroundColor:g+"4D"}],i&&p.push({label:f,data:l,borderColor:h,backgroundColor:h+"4D"}),chartInstance=new Chart(m,{type:"radar",data:{labels:j,datasets:p},options:r});break;case"scatter":r.scales={x:{type:"category",labels:j,grid:{display:c}},y:{grid:{display:c},beginAtZero:!0}},p=[{label:e.value+" vs "+d.value,data:k.map(function(a,b){return{x:j[b],y:a}}),backgroundColor:g}],chartInstance=new Chart(m,{type:"scatter",data:{datasets:p},options:r});break}updateControlVisibility()}
        function updateKPIs(a,b,c,d,e,f){var g=a.reduce(function(a,b){return a+b},0),h=f?b.reduce(function(a,b){return a+b},0):0,i=-1/0,j=-1;for(var k=0;k<a.length;k++)a[k]>i&&(i=a[k],j=k);var l=e[j],m=document.getElementById("kpi-grid");m.innerHTML='<div class="card"><div class="kpi-value">'+g.toLocaleString("pt-BR")+'</div><div class="kpi-label">Total de '+c+'</div></div><div class="card"><div class="kpi-value">'+(f?h.toLocaleString("pt-BR"):"N/A")+'</div><div class="kpi-label">Total de '+(f?d:"-")+'</div></div><div class="card"><div class="kpi-value">'+l+'</div><div class="kpi-label">Ponto de Maior '+c+"</div></div>"}
        function downloadChart(){if(chartInstance){var a=document.createElement("a");a.href=chartInstance.toBase64Image(),a.download="PowerChart_Grafico.png",a.click()}}
        document.addEventListener("DOMContentLoaded",function(){try{var a=findDefaultAxes();populateControls(a),renderChart(),document.getElementById("update-charts-btn").addEventListener("click",renderChart),document.getElementById("download-btn").addEventListener("click",downloadChart),document.querySelectorAll('input[name="chart-type"]').forEach(function(a){a.addEventListener("change",renderChart)})}catch(a){}});
    </script>
</body>
</html>
"@
}

# --- 3. Construção da Interface Gráfica (Windows Forms) ---
# O código do editor PowerShell continua o mesmo.
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerChart Editor 10.0"
$Form.Width = 1024
$Form.Height = 768
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

