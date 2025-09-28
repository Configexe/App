# -----------------------------------------------------------------------------
# PowerChart Designer: Editor de Dados
# Versão: 8.0 - Arquitetura Desacoplada (Editor + Relatório HTML)
# Autor: Seu Nome/Empresa
# Descrição: Esta ferramenta é responsável por carregar, validar e editar
#            dados de arquivos CSV. Ela gera um arquivo HTML separado para
#            a visualização dos gráficos no navegador padrão do usuário.
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

    # Converte os dados da grade (que podem ter sido editados) para JSON
    $DataForJson = $DataGridView.DataSource | ForEach-Object {
        $properties = @{}
        foreach ($prop in $_.PSObject.Properties) {
            $properties[$prop.Name] = $prop.Value
        }
        New-Object -TypeName PSObject -Property $properties
    }

    $JsonData = $DataForJson | ConvertTo-Json -Compress -Depth 5
    $JsonColumnNames = $DataGridView.Columns.DataPropertyName | ConvertTo-Json -Compress

    # Caminho do arquivo de saída
    $OutputPath = Join-Path $env:TEMP "PowerChart_Relatorio.html"

    # Gera o conteúdo do HTML
    $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -JsonColumnNames $JsonColumnNames
    
    try {
        $HtmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
        # Abre o arquivo no navegador padrão
        Start-Process $OutputPath
        $StatusLabel.Text = "Relatório gerado e aberto com sucesso!"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao gerar ou abrir o arquivo HTML: $($_.Exception.Message)", "Erro", "OK", "Error")
        $StatusLabel.Text = "Falha ao gerar o relatório."
    }
}

# Template do HTML (separado para clareza)
Function Get-HtmlTemplate {
    param($JsonData, $JsonColumnNames)

    # Este é o conteúdo do arquivo PowerChart_Relatorio.html
    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PowerChart - Relatório Dinâmico</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Inter', sans-serif; background-color: #f8fafc; }
        .card { background-color: white; border-radius: 0.75rem; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1); padding: 1.5rem; }
        .kpi-value { font-size: 2rem; font-weight: 900; color: #1e293b; }
        .kpi-label { font-size: 0.875rem; color: #64748b; margin-top: 0.25rem; }
        .chart-container { position: relative; width: 100%; height: 380px; }
    </style>
</head>
<body class="text-gray-900">
    <header class="bg-[#0f172a] text-white text-center py-12 px-4">
        <h1 class="text-4xl md:text-5xl font-black tracking-tight">Relatório Dinâmico Interativo</h1>
        <p class="mt-4 text-lg text-blue-200 max-w-3xl mx-auto">Dados gerados e processados via PowerChart Editor.</p>
    </header>
    <main class="container mx-auto p-4 md:p-8 -mt-10">
        <section id="controls" class="card mb-6">
            <h2 class="text-xl font-bold text-[#1e293b] mb-4">Configurações dos Gráficos</h2>
            <div class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-4 items-end">
                <div><label for="x-axis">Eixo X:</label><select id="x-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                <div><label for="y1-axis">Série Y1:</label><select id="y1-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                <div><label for="y1-color">Cor Y1:</label><input type="color" id="y1-color" value="#3b82f6" class="w-full h-10 mt-1"></div>
                <div><label for="y2-axis">Série Y2:</label><select id="y2-axis" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm"></select></div>
                <div><label for="y2-color">Cor Y2:</label><input type="color" id="y2-color" value="#ef4444" class="w-full h-10 mt-1"></div>
            </div>
            <div class="mt-4"><button id="update-charts-btn" class="w-full bg-blue-600 text-white font-bold py-2 px-4 rounded-lg">Atualizar Gráficos</button></div>
        </section>
        <section id="kpis" class="mb-6"><div id="kpi-grid" class="grid grid-cols-1 md:grid-cols-3 gap-6"></div></section>
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6 mt-6">
            <section class="card lg:col-span-2"><h3 class="text-lg font-bold">Gráfico Combo (Barra + Linha)</h3><div class="chart-container"><canvas id="comboChart"></canvas></div></section>
            <section class="card"><h3 class="text-lg font-bold">Comparativo de Barras</h3><div class="chart-container"><canvas id="barChart"></canvas></div></section>
            <section class="card"><h3 class="text-lg font-bold">Evolução em Linha</h3><div class="chart-container"><canvas id="lineChart"></canvas></div></section>
            <section class="card lg:col-span-2"><h3 class="text-lg font-bold">Gráfico de Barras Empilhadas</h3><div class="chart-container"><canvas id="stackedBarChart"></canvas></div></section>
        </div>
    </main>
    <script>
        var RAW_DATA = $JsonData;
        var COLUMN_NAMES = $JsonColumnNames;
        // Restante do JS é idêntico à versão 7.0 e funcionará perfeitamente no navegador.
        var myCharts = [];
        function parseNumericValue(v){if(typeof v==='number'){return v}if(typeof v!=='string'){return 0}var c=v.replace(/[^0-9,-]/g,'').replace(',','.');var p=parseFloat(c);return isNaN(p)?0:p}
        function findDefaultAxes(){var d={xAxis:null,y1Axis:null,y2Axis:null},n=[],t=[];if(!RAW_DATA||RAW_DATA.length===0){return d}var f=RAW_DATA[0];for(var i=0;i<COLUMN_NAMES.length;i++){var o=COLUMN_NAMES[i],a=f[o];if(a===null||a===undefined)continue;var l=parseFloat(String(a).replace(',','.'));!isNaN(l)&&String(a).trim()!==''?n.push(o):t.push(o)}d.xAxis=t[0]||COLUMN_NAMES[0]||null;d.y1Axis=n[0]||(COLUMN_NAMES.length>1?COLUMN_NAMES[1]:null);d.y2Axis=n[1]||(COLUMN_NAMES.length>2?COLUMN_NAMES[2]:null);return d}
        function populateControls(d){var n=['x-axis','y1-axis','y2-axis'];n.forEach(function(o){var a=document.getElementById(o);a.innerHTML=o==='y2-axis'?'<option value="Nenhum">Nenhum</option>':'';COLUMN_NAMES.forEach(function(n){a.innerHTML+='<option value="'+n+'">'+n+'</option>'})});document.getElementById('x-axis').value=d.xAxis||'';document.getElementById('y1-axis').value=d.y1Axis||'';document.getElementById('y2-axis').value=d.y2Axis||'Nenhum'}
        function renderCharts(){myCharts.forEach(function(d){d.destroy()});myCharts=[];var n=document.getElementById('x-axis').value,o=document.getElementById('y1-axis').value,a=document.getElementById('y2-axis').value,l=document.getElementById('y1-color').value,r=document.getElementById('y2-color').value,t=a!=='Nenhum';if(!n||!o){return}var i=RAW_DATA.map(function(d){return d[n]}),e=RAW_DATA.map(function(d){return parseNumericValue(d[o])}),s=t?RAW_DATA.map(function(d){return parseNumericValue(d[a])}):[];updateKPIs(e,s,o,a,i,t);myCharts.push(createComboChart(i,e,s,o,a,l,r,t));myCharts.push(createBarChart(i,e,o,l));myCharts.push(createLineChart(i,e,o,l));myCharts.push(createStackedBarChart(i,e,s,o,a,l,r,t))}
        function updateKPIs(d,n,o,a,l,r){var t=d.reduce(function(d,n){return d+n},0),i=r?n.reduce(function(d,n){return d+n},0):0,e=-Infinity,s=-1;for(var u=0;u<d.length;u++){if(d[u]>e){e=d[u];s=u}}var c=l[s],h=document.getElementById('kpi-grid');h.innerHTML='<div class="card"><div class="kpi-value">'+t.toLocaleString('pt-BR')+'</div><div class="kpi-label">Total de '+o+'</div></div>'+'<div class="card"><div class="kpi-value">'+(r?i.toLocaleString('pt-BR'):'N/A')+'</div><div class="kpi-label">Total de '+(r?a:'-')+'</div></div>'+'<div class="card"><div class="kpi-value">'+c+'</div><div class="kpi-label">Ponto de Maior '+o+'</div></div>'}
        var defaultChartOptions={responsive:!0,maintainAspectRatio:!1,plugins:{legend:{position:'bottom'}}};
        function createComboChart(d,n,o,a,l,r,t,i){var e=document.getElementById('comboChart').getContext('2d'),s=[{type:'bar',label:a,data:n,backgroundColor:r+'B3',yAxisID:'y'}];if(i){s.push({type:'line',label:l,data:o,borderColor:t,tension:.4,yAxisID:'y1'})}var u={responsive:!0,maintainAspectRatio:!1,plugins:{legend:{position:'bottom'}},scales:{y:{position:'left'},y1:{display:i,position:'right',grid:{drawOnChartArea:!1}}}};return new Chart(e,{data:{labels:d,datasets:s},options:u})}
        function createBarChart(d,n,o,a){return new Chart(document.getElementById('barChart').getContext('2d'),{type:'bar',data:{labels:d,datasets:[{label:o,data:n,backgroundColor:a}]},options:defaultChartOptions})}
        function createLineChart(d,n,o,a){return new Chart(document.getElementById('lineChart').getContext('2d'),{type:'line',data:{labels:d,datasets:[{label:o,data:n,borderColor:a,backgroundColor:a+'33',fill:!0,tension:.4}]},options:defaultChartOptions})}
        function createStackedBarChart(d,n,o,a,l,r,t,i){var e=document.getElementById('stackedBarChart').getContext('2d'),s=[{label:a,data:n,backgroundColor:r}];if(i){s.push({label:l,data:o,backgroundColor:t})}var u={responsive:!0,maintainAspectRatio:!1,plugins:{legend:{position:'bottom'}},scales:{x:{stacked:!0},y:{stacked:!0}}};return new Chart(e,{type:'bar',data:{labels:d,datasets:s},options:u})}
        document.addEventListener('DOMContentLoaded',function(){try{var d=findDefaultAxes();populateControls(d);renderCharts();document.getElementById('update-charts-btn').addEventListener('click',renderCharts)}catch(n){}});
    </script>
</body>
</html>
"@
}


# --- 3. Construção da Interface Gráfica (Windows Forms) ---

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerChart Editor 8.0"
$Form.Width = 1024
$Form.Height = 768
$Form.StartPosition = "CenterScreen"

# Layout Principal
$MainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$MainLayout.Dock = "Fill"
$MainLayout.ColumnCount = 1
$MainLayout.RowCount = 2
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$Form.Controls.Add($MainLayout)

# Painel de Controles (Topo)
$ControlPanel = New-Object System.Windows.Forms.Panel
$ControlPanel.Dock = "Fill"
$ControlPanel.BackColor = [System.Drawing.Color]::FromArgb(240, 242, 245)
$ControlPanel.Padding = New-Object System.Windows.Forms.Padding(5)
$MainLayout.Controls.Add($ControlPanel, 0, 0)

# Botões
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
$ButtonGenerateHtml.Enabled = $false # Desabilitado até carregar dados
$ControlPanel.Controls.Add($ButtonGenerateHtml)

# Label de Status
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Text = "Aguardando arquivo CSV..."
$StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$StatusLabel.Location = New-Object System.Drawing.Point(370, 15)
$StatusLabel.AutoSize = $true
$ControlPanel.Controls.Add($StatusLabel)

# Grade de Dados
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
