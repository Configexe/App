# -----------------------------------------------------------------------------
# Power-Graphx Editor: Análise e Visualização de Dados
# Versão: 1.7.9 - Edição com Documentação Completa
# Autor: jefferson/configexe

# -----------------------------------------------------------------------------

# --- 1. Carregar Assemblies Necessárias ---
# Adiciona as bibliotecas .NET necessárias para a criação de interfaces gráficas
# com Windows Forms e para manipulação de elementos gráficos como ícones e cores.
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "Não foi possível carregar as assemblies necessárias."
    exit 1
}

# --- 2. Variáveis Globais de Estado ---
# Mantêm o estado da aplicação entre diferentes funções.
$Global:OriginalData = $null      # Armazena a lista de dados original carregada do CSV.
$Global:IsDataFiltered = $false     # Flag para indicar se um filtro está atualmente aplicado.
$Global:ColumnToModifyIndex = -1    # Armazena o índice da coluna clicada com o botão direito.

# --- 3. Funções Auxiliares de UI ---

# Função para criar uma caixa de diálogo e solicitar entrada do usuário.
# O PowerShell não possui um "InputBox" nativo, então criamos um formulário simples para isso.
Function Show-InputBox {
    param(
        [string]$Title,
        [string]$Prompt,
        [string]$DefaultText = ""
    )
    # Criação e configuração do formulário.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Font = 'Segoe UI, 9'
    $form.StartPosition = 'CenterScreen'
    $form.ClientSize = New-Object System.Drawing.Size(350, 120)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Adiciona um rótulo com a pergunta para o usuário.
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(10, 15)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    # Adiciona a caixa de texto para entrada.
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $DefaultText
    $textBox.Location = New-Object System.Drawing.Point(12, 40)
    $textBox.Size = New-Object System.Drawing.Size(326, 23)
    $form.Controls.Add($textBox)

    # Adiciona o botão "OK".
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(182, 75)
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton # Permite pressionar Enter para confirmar.

    # Adiciona o botão "Cancelar".
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancelar"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(263, 75)
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton # Permite pressionar Esc para cancelar.

    # Exibe o formulário e aguarda a interação do usuário.
    if ($form.ShowDialog() -eq 'OK') {
        return $textBox.Text # Retorna o texto se OK for pressionado.
    }
    return $null # Retorna nulo se for cancelado.
}

# Função auxiliar para obter os nomes de propriedade de todas as colunas no DataGridView.
Function Get-ColumnNames {
    param([Parameter(Mandatory=$true)]$DataGridView)
    $columnNames = @()
    foreach ($column in $DataGridView.Columns) {
        $columnNames += $column.DataPropertyName
    }
    return $columnNames
}

# --- 4. Funções de Manipulação de Dados ---

# Abre um diálogo para carregar um arquivo CSV e exibi-lo na grade.
Function Load-CSVData {
    param(
        [Parameter(Mandatory=$true)]$DataGridView,
        [Parameter(Mandatory=$true)]$StatusLabel,
        [Parameter(Mandatory=$true)]$MenuItemsToEnable
    )
    # Configura e exibe a janela de seleção de arquivo.
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Selecione o arquivo CSV"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        $StatusLabel.Text = "Analisando: $(Split-Path $FilePath -Leaf)..."
        $StatusLabel.Owner.Refresh()

        $Data = $null
        try {
            # Detecta automaticamente o delimitador (vírgula ou ponto e vírgula) com base na primeira linha.
            $firstLine = Get-Content -Path $FilePath -TotalCount 1
            $bestDelimiter = if (($firstLine -split ';').Count -gt ($firstLine -split ',').Count) { ';' } else { ',' }
            $Data = Import-Csv -Path $FilePath -Delimiter $bestDelimiter
        }
        catch {
            # O erro será tratado na verificação de $Data nulo.
        }

        if ($null -ne $Data -and $Data.Count -gt 0) {
            # Se os dados forem carregados com sucesso:
            $Global:OriginalData = [System.Collections.ArrayList]$Data # Salva uma cópia original dos dados.
            $DataGridView.DataSource = $Global:OriginalData # Vincula os dados à grade.
            $DataGridView.AutoSizeColumnsMode = 'AllCells'
            $StatusLabel.Text = "Arquivo carregado: $(Split-Path $FilePath -Leaf) ($($Data.Count) linhas)"
            $MenuItemsToEnable | ForEach-Object { $_.Enabled = $true } # Ativa os menus de dados e relatório.
            $Global:IsDataFiltered = $false
        } else {
            # Se houver falha na leitura:
            $DataGridView.DataSource = $null
            $Global:OriginalData = $null
            [System.Windows.Forms.MessageBox]::Show("Não foi possível ler os dados do arquivo CSV.", "Erro de Leitura", "OK", "Error")
            $StatusLabel.Text = "Falha ao carregar arquivo."
            $MenuItemsToEnable | ForEach-Object { $_.Enabled = $false } # Desativa os menus.
        }
    }
}

# Exibe um diálogo para ordenar os dados na grade por uma coluna específica.
Function Sort-Data {
    param($DataGridView, $StatusLabel)
    
    $columnNames = Get-ColumnNames -DataGridView $DataGridView
    if ($columnNames.Count -eq 0) { return }

    # Cria o formulário de diálogo para escolher as opções de ordenação.
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Ordenar Dados"; $form.ClientSize = "300,150"; $form.StartPosition = 'CenterParent'
    
    $columnLabel = New-Object System.Windows.Forms.Label; $columnLabel.Text = "Ordenar por coluna:"; $columnLabel.Location = "10,10"; $form.Controls.Add($columnLabel)
    $columnComboBox = New-Object System.Windows.Forms.ComboBox; $columnComboBox.Location = "10,30"; $columnComboBox.Size = "280,20"; $columnComboBox.Items.AddRange($columnNames); $columnComboBox.SelectedIndex = 0; $form.Controls.Add($columnComboBox)
    
    $orderLabel = New-Object System.Windows.Forms.Label; $orderLabel.Text = "Ordem:"; $orderLabel.Location = "10,60"; $form.Controls.Add($orderLabel)
    $ascRadio = New-Object System.Windows.Forms.RadioButton; $ascRadio.Text = "Crescente"; $ascRadio.Location = "10,80"; $ascRadio.Checked = $true; $form.Controls.Add($ascRadio)
    $descRadio = New-Object System.Windows.Forms.RadioButton; $descRadio.Text = "Decrescente"; $descRadio.Location = "120,80"; $form.Controls.Add($descRadio)
    
    $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "OK"; $okButton.Location = "130,110"; $okButton.DialogResult = 'OK'; $form.Controls.Add($okButton)
    $cancelButton = New-Object System.Windows.Forms.Button; $cancelButton.Text = "Cancelar"; $cancelButton.Location = "210,110"; $cancelButton.DialogResult = 'Cancel'; $form.Controls.Add($cancelButton)

    if ($form.ShowDialog() -eq 'OK') {
        $columnToSort = $columnComboBox.SelectedItem
        $descending = $descRadio.Checked
        $currentData = $DataGridView.DataSource
        
        # Mapa para ordenar nomes de meses corretamente (ex: Jan, Fev, Mar...).
        $monthMap = @{
            'janeiro' = 1; 'jan' = 1; 'fevereiro' = 2; 'fev' = 2; 'março' = 3; 'mar' = 3; 'abril' = 4; 'abr' = 4;
            'maio' = 5; 'mai' = 5; 'junho' = 6; 'jun' = 6; 'julho' = 7; 'jul' = 7; 'agosto' = 8; 'ago' = 8;
            'setembro' = 9; 'set' = 9; 'outubro' = 10; 'out' = 10; 'novembro' = 11; 'nov' = 11; 'dezembro' = 12; 'dez' = 12
        }

        # Bloco de script para "ordenação inteligente".
        # Ele tenta converter o valor para o tipo mais apropriado antes de comparar.
        $sortExpression = {
            param($row)
            $value = $row."$columnToSort"
            if ($null -eq $value) { return $null }

            # 1. Tenta converter para número de mês.
            $valueStr = $value.ToString().ToLower().Trim()
            if ($monthMap.ContainsKey($valueStr)) { return $monthMap[$valueStr] }

            # 2. Tenta converter para Data/Hora.
            try { return [datetime]::Parse($value, [System.Globalization.CultureInfo]::CurrentCulture) } catch {}
            
            # 3. Tenta converter para número decimal (considerando cultura pt-BR para vírgulas).
            $decimalValue = 0
            $ci = [System.Globalization.CultureInfo]::GetCultureInfo('pt-BR')
            if ([decimal]::TryParse($value.ToString(), [System.Globalization.NumberStyles]::Any, $ci, [ref]$decimalValue)) {
                return $decimalValue
            }

            # 4. Se tudo falhar, ordena como texto.
            return $value.ToString()
        }

        # Aplica a ordenação usando a expressão customizada.
        $StatusLabel.Text = "Ordenando..."; $StatusLabel.Owner.Refresh()
        $sortedData = $currentData | Sort-Object -Property @{Expression = $sortExpression} -Descending:$descending
        $DataGridView.DataSource = [System.Collections.ArrayList]$sortedData
        $StatusLabel.Text = "Dados ordenados por '$columnToSort'."
    }
}

# Adiciona uma nova coluna à grade com base em uma fórmula PowerShell fornecida pelo usuário.
Function Add-CalculatedColumn {
    param($DataGridView, $StatusLabel)
    
    $newColumnName = Show-InputBox -Title "Adicionar Coluna Calculada" -Prompt "Nome da nova coluna:"
    if ([string]::IsNullOrWhiteSpace($newColumnName)) { return }

    $columnNames = (Get-ColumnNames -DataGridView $DataGridView) -join "', '"
    $formula = Show-InputBox -Title "Fórmula da Coluna" -Prompt "Digite a fórmula (ex: `$_.Valor * 1.1`). Colunas: '$columnNames'"
    if ([string]::IsNullOrWhiteSpace($formula)) { return }

    try {
        # Converte a string da fórmula em um bloco de script executável.
        $scriptBlock = [scriptblock]::Create($formula)
        $currentData = $DataGridView.DataSource
        
        $StatusLabel.Text = "Calculando nova coluna..."; $StatusLabel.Owner.Refresh()
        # Itera sobre cada linha dos dados e aplica a fórmula.
        foreach ($row in $currentData) {
            $result = & $scriptBlock -InputObject $row
            # Adiciona o resultado como uma nova propriedade na linha.
            Add-Member -InputObject $row -MemberType NoteProperty -Name $newColumnName -Value $result
        }

        # Força a atualização do DataGridView para mostrar a nova coluna.
        $DataGridView.DataSource = $null
        $DataGridView.DataSource = $currentData
        $DataGridView.AutoSizeColumnsMode = 'AllCells'
        $StatusLabel.Text = "Coluna '$newColumnName' adicionada."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erro na fórmula: $($_.Exception.Message)", "Erro de Sintaxe", "OK", "Error")
        $StatusLabel.Text = "Falha ao adicionar coluna."
    }
}

# Restaura os dados originais, removendo qualquer filtro aplicado.
Function Remove-DataFilter {
    param($DataGridView, $StatusLabel)

    if (!$Global:IsDataFiltered) { 
        $StatusLabel.Text = "Nenhum filtro ativo para remover."
        return 
    }
    
    $StatusLabel.Text = "Removendo filtro..."; $StatusLabel.Owner.Refresh()
    $DataGridView.DataSource = $Global:OriginalData
    $Global:IsDataFiltered = $false
    $StatusLabel.Text = "Filtro removido. Exibindo todos os $($Global:OriginalData.Count) registros."
}

# --- 5. Funções do Relatório HTML ---

# Gera o arquivo HTML com os dados atuais e o abre no navegador padrão.
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
    $StatusLabel.Owner.Refresh()

    # Converte os dados do PowerShell para um formato genérico (PSObject).
    $DataForJson = $DataGridView.DataSource | ForEach-Object {
        $properties = @{}
        foreach ($prop in $_.PSObject.Properties) {
            $properties[$prop.Name] = $prop.Value
        }
        New-Object -TypeName PSObject -Property $properties
    }

    # Obtém a estrutura das colunas, incluindo nomes renomeados.
    $ColumnStructure = $DataGridView.Columns | ForEach-Object {
        [PSCustomObject]@{
            OriginalName = $_.DataPropertyName
            DisplayName  = $_.HeaderText
        }
    }

    # Serializa os dados e a estrutura de colunas para JSON.
    $JsonData = $DataForJson | ConvertTo-Json -Compress -Depth 5
    $JsonColumnStructure = $ColumnStructure | ConvertTo-Json -Compress

    # Gera o conteúdo HTML completo, injetando os dados JSON.
    $OutputPath = Join-Path $env:TEMP "PowerGraphx_Relatorio.html"
    $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -JsonColumnStructure $JsonColumnStructure
    
    try {
        # Salva o HTML em um arquivo temporário e o abre.
        $HtmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
        Start-Process $OutputPath
        $StatusLabel.Text = "Relatório gerado e aberto com sucesso!"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao gerar ou abrir o arquivo HTML: $($_.Exception.Message)", "Erro", "OK", "Error")
        $StatusLabel.Text = "Falha ao gerar o relatório."
    }
}

# Função que contém o template HTML, CSS e JavaScript do relatório.
Function Get-HtmlTemplate {
    param($JsonData, $JsonColumnStructure)

    # A 'here-string' (@"..."@) permite escrever um bloco de texto multilinha.
    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power-Graphx - By jeff</title>
    <!-- Inclusão de bibliotecas externas via CDN -->
    <script src="https://cdn.tailwindcss.com"></script> <!-- Framework CSS para estilização rápida -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script> <!-- Biblioteca de gráficos -->
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script> <!-- Plugin para rótulos nos gráficos -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap" rel="stylesheet">
    <!-- Estilos CSS customizados -->
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
    <!-- Cabeçalho da página -->
    <header class="bg-[#0f172a] text-white text-center py-12 px-4">
        <h1 class="text-4xl md:text-5xl font-black tracking-tight">Relatório Dinâmico Interativo - By jeff</h1>
        <p class="mt-4 text-lg text-blue-200 max-w-3xl mx-auto">Dados processados via Power-Graphx Editor - By jeff.</p>
    </header>
    <!-- Conteúdo principal -->
    <main class="container mx-auto p-4 md:p-8 -mt-10">
        <!-- Seção de Controles (Seleção de dados e tipo de gráfico) -->
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
                              <div><input type="radio" name="chart-type" value="bar" id="type-bar" checked><label for="type-bar"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg><span class="text-xs font-semibold">Barra</span></label></div>
                              <div><input type="radio" name="chart-type" value="combo" id="type-combo"><label for="type-combo"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/><path d="M3 12l5-4 5 6 5-4"/></svg><span class="text-xs font-semibold">Combo</span></label></div>
                              <div><input type="radio" name="chart-type" value="stacked" id="type-stacked"><label for="type-stacked"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="#3b82f6" stroke="#fff" stroke-width="1"><rect x="5" y="12" width="4" height="6"/><rect x="10" y="8" width="4" height="10"/><rect x="15" y="4" width="4" height="14"/><rect x="5" y="9" width="4" height="3" fill="#ef4444"/><rect x="10" y="4" width="4" height="4" fill="#ef4444"/><rect x="15" y="2" width="4" height="2" fill="#ef4444"/></svg><span class="text-xs font-semibold">Empilhado</span></label></div>
                              <div><input type="radio" name="chart-type" value="line" id="type-line"><label for="type-line"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3v18h18"/><path d="M3 17l5-4 5 6 5-4 4 2"/></svg><span class="text-xs font-semibold">Linha</span></label></div>
                              <div><input type="radio" name="chart-type" value="horizontalBar" id="type-horizontalBar"><label for="type-horizontalBar"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" transform="rotate(90) scale(1,-1)"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg><span class="text-xs font-semibold">Horizontal</span></label></div>
                              <div><input type="radio" name="chart-type" value="stackedLine" id="type-stackedLine"><label for="type-stackedLine"><svg class="w-8 h-8 mb-1" viewBox="0 0 24 24" fill="#3b82f6" fill-opacity="0.5" stroke="#3b82f6" stroke-width="2"><path d="M3 17l5-4 5 6 5-4 4 2V21H3z"/><path d="M3 12l5-3 5 5 5-3 4 2v5l-4-2-5 3-5-5-5 3z" fill="#ef4444" fill-opacity="0.5" stroke="#ef4444"/></svg><span class="text-xs font-semibold">Linha Emp.</span></label></div>
                          </div>
                      </div>
                      <div class="mt-6"><button id="update-charts-btn" class="w-full bg-blue-600 text-white font-bold py-3 px-4 rounded-lg transition hover:bg-blue-700 flex items-center justify-center text-lg">
                          Atualizar Gráfico
                      </button></div>
                  </div>
              </div>
        </section>
        
        <!-- Grid para o gráfico e o painel de formatação -->
        <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">
            <!-- Card do Gráfico -->
            <div id="chart-card" class="lg:col-span-3 card transition-all duration-300">
                  <div class="flex justify-between items-center mb-4">
                      <h3 id="chart-title" class="text-xl font-bold text-[#1e293b]"></h3>
                  </div>
                  <div class="chart-container"><canvas id="mainChart"></canvas></div>
            </div>
            <!-- Painel de Formatação -->
            <div id="format-panel" class="lg:col-span-1 card">
                <h3 class="text-xl font-bold text-[#1e293b] mb-4">Formatar Visual</h3>
                <div class="space-y-4">
                    <!-- Opções de Aparência, Rótulos, Barras, Linhas, etc. -->
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
                    <div id="bar-options" class="hidden">
                        <span class="font-semibold text-gray-700 text-sm">Opções de Barra</span>
                        <div class="mt-2">
                             <label for="bar-border-radius" class="text-xs text-gray-600">Arredondamento da Borda:</label>
                             <input type="number" id="bar-border-radius" value="0" min="0" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm">
                        </div>
                        <div class="divider mt-4"></div>
                    </div>
                    <div id="line-options" class="hidden">
                        <span class="font-semibold text-gray-700 text-sm">Opções de Linha</span>
                        <div class="mt-2">
                             <label for="line-interpolation" class="text-xs text-gray-600">Interpolação:</label>
                             <select id="line-interpolation" class="mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm">
                                 <option value="0.0">Linear</option>
                                 <option value="0.4" selected>Suave (Padrão)</option>
                                 <option value="1.0">Curva Máxima</option>
                             </select>
                        </div>
                        <div class="divider mt-4"></div>
                    </div>
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
                    <div class="divider"></div>
                     <div>
                        <span class="font-semibold text-gray-700 text-sm">Exportar</span>
                        <button id="download-chart-btn" class="mt-2 w-full bg-gray-600 text-white font-bold py-2 px-3 rounded-lg text-sm transition hover:bg-gray-700">
                            Baixar Gráfico (PNG)
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </main>
    <script>
        // --- INÍCIO DO JAVASCRIPT ---
        // Variáveis globais para o script da página.
        let chartInstance; // Armazena a instância do gráfico atual para poder atualizá-la ou destruí-la.
        let seriesCounter = 0; // Contador para dar IDs únicos a cada série de dados adicionada.
        const seriesColors = ["#3b82f6", "#ef4444", "#22c55e", "#f97316", "#8b5cf6", "#14b8a6"]; // Paleta de cores padrão.
        
        // Função principal, executada após o carregamento da página.
        function initializeApp(rawData, columnStructure) {
            if (!rawData || !columnStructure) {
                console.error("Dados ou estrutura de colunas não fornecidos.");
                return;
            }
            // Registra o plugin de rótulos de dados para que o Chart.js possa usá-lo.
            Chart.register(ChartDataLabels);

            // Tenta converter um valor de texto (do CSV) em um número. Remove caracteres não numéricos.
            function parseValue(value) {
                if (typeof value === 'number') return value;
                if (typeof value !== 'string') return value;
                const cleanValue = value.replace(/[^0-9,-]/g, '').replace(',', '.');
                const parsed = parseFloat(cleanValue);
                return isNaN(parsed) ? value : parsed;
            }
            
            // Verifica se uma coluna contém dados numéricos, analisando uma amostra.
            function isNumeric(colName) {
                if (rawData.length === 0) return false;
                const sampleValue = rawData.find(d => d[colName] !== null && typeof d[colName] !== 'undefined');
                if (!sampleValue) return false;
                const parsed = parseFloat(String(sampleValue[colName]).replace(',', '.'));
                return !isNaN(parsed) && String(sampleValue[colName]).trim() !== '';
            }

            // Preenche um elemento <select> (dropdown) com os nomes das colunas.
            function populateSelect(selectElement, type = 'all') {
                if (!selectElement) return;
                selectElement.innerHTML = '';
                const options = columnStructure.filter(col => {
                    if (type === 'numeric') return isNumeric(col.OriginalName); // Filtra apenas colunas numéricas
                    if (type === 'text') return !isNumeric(col.OriginalName); // Filtra apenas colunas de texto
                    return true; // Pega todas as colunas
                });
                
                options.forEach(col => {
                    const option = document.createElement('option');
                    option.value = col.OriginalName;
                    option.textContent = col.DisplayName;
                    selectElement.appendChild(option);
                });
            }

            // Adiciona um novo bloco de controles para uma série de dados na UI.
            function addSeriesControl(isFirst = false) {
                const seriesContainer = document.getElementById('series-container');
                const seriesId = ++seriesCounter;
                const defaultColor = seriesColors[(seriesId - 1) % seriesColors.length];
                const seriesDiv = document.createElement('div');
                seriesDiv.id = 'series-' + seriesId;
                seriesDiv.className = 'p-3 border rounded-lg bg-gray-50 grid grid-cols-1 sm:grid-cols-2 gap-3 items-end';
                
                // Cria o HTML para os controles da nova série (eixos, tipo, cor).
                let content = '';
                const eixoXLabel = isFirst ? 'Eixo X / Grupo:' : 'Eixo X / Grupo ' + seriesId + ':';
                const eixoYLabel = isFirst ? 'Eixo Y / Valor:' : 'Eixo Y / Valor ' + seriesId + ':';
                
                content += '<div class="x-axis-control"><label class="text-xs font-semibold">' + eixoXLabel + '</label><select name="x-axis" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm"></select></div>';
                content += '<div><label class="text-xs font-semibold">' + eixoYLabel + '</label><select name="y-axis" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm y-axis-select"></select></div>';
                content += '<div class="combo-type-control"><label class="text-xs font-semibold">Tipo:</label><select name="series-type" class="axis-select mt-1 block w-full rounded-md border-gray-300 shadow-sm text-sm"><option value="bar">Barra</option><option value="line">Linha</option></select></div>';
                content += '<div class="flex items-end space-x-2"><div class="w-full"><label class="text-xs font-semibold">Cor:</label><input type="color" value="' + defaultColor + '" name="color" class="axis-select mt-1 w-full h-9"></div>';
                if (!isFirst) {
                    content += '<button type="button" onclick="this.parentElement.parentElement.remove(); renderChart();" class="h-9 px-3 bg-red-500 text-white rounded-md hover:bg-red-600 transition text-sm font-bold">&times;</button>';
                }
                content += '</div>';

                seriesDiv.innerHTML = content;
                seriesContainer.appendChild(seriesDiv);
                
                // Preenche os novos dropdowns com as colunas disponíveis.
                populateSelect(seriesDiv.querySelector('[name="x-axis"]'), 'all');
                populateSelect(seriesDiv.querySelector('[name="y-axis"]'), 'numeric');
            }

            // Função principal de renderização do gráfico. É chamada sempre que algo muda.
            window.renderChart = function() {
                // Destrói o gráfico anterior para evitar sobreposição e problemas de memória.
                if (chartInstance) { chartInstance.destroy(); }
                const container = document.querySelector('.chart-container');
                container.innerHTML = '<canvas id="mainChart"></canvas>';
                const ctx = document.getElementById('mainChart').getContext('2d');
                
                const chartType = document.querySelector('input[name="chart-type"]:checked').value;
                const seriesControls = document.querySelectorAll('#series-container > div');
                if (seriesControls.length === 0) return;

                const chartData = { datasets: [] };
                
                // Define os rótulos do eixo X (categorias).
                const firstXAxisSelect = seriesControls[0] ? seriesControls[0].querySelector('[name="x-axis"]') : null;
                if (firstXAxisSelect && firstXAxisSelect.value) {
                     // Pega os valores únicos da coluna do eixo X NA ORDEM EM QUE APARECEM.
                     // A remoção do .sort() aqui é CRUCIAL para respeitar a ordenação feita no PowerShell.
                     chartData.labels = [...new Set(rawData.map(d => d[firstXAxisSelect.value]))];
                }
                
                // Itera sobre cada bloco de série de dados para construir os datasets do gráfico.
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
                    const lineTension = document.getElementById('line-interpolation').value;
                    const borderRadius = document.getElementById('bar-border-radius').value;

                    // Monta o objeto de 'dataset' para o Chart.js.
                    const dataset = {
                        label: colInfo.DisplayName,
                        borderColor: colorInput.value,
                        backgroundColor: colorInput.value + 'B3',
                    };
                    
                    // Define o tipo de gráfico para esta série específica.
                    let seriesType = (chartType === 'combo' && typeSelect) ? typeSelect.value : (['bar', 'stacked', 'horizontalBar'].includes(chartType) ? 'bar' : chartType);
                    if(chartType === 'stackedLine') seriesType = 'line';
                    
                    if (seriesType === 'line') dataset.tension = parseFloat(lineTension);
                    if (seriesType === 'bar') dataset.borderRadius = parseInt(borderRadius) || 0;
                    dataset.type = seriesType;

                    // Processa e agrega os dados. Para cada rótulo no eixo X, soma os valores correspondentes.
                    dataset.data = chartData.labels.map(label => {
                        const relevantRows = rawData.filter(d => d[xCol] === label);
                        if (relevantRows.length === 0) return 0;
                        const sum = relevantRows.reduce((acc, curr) => acc + (parseValue(curr[yCol]) || 0), 0);
                        return sum;
                    });
                    chartData.datasets.push(dataset);
                });
                
                if (chartData.datasets.length === 0) return;
                
                // Cria a nova instância do gráfico com os dados e opções.
                const finalChartType = ['bar', 'combo', 'stacked', 'horizontalBar'].includes(chartType) ? 'bar' : chartType;
                const options = buildChartOptions(chartType);
                chartInstance = new Chart(ctx, { type: finalChartType, data: chartData, options: options });
            }

            // Constrói o objeto de opções de formatação para o Chart.js.
            function buildChartOptions(chartType) {
                // Lê todas as opções do painel de formatação.
                const darkMode = document.getElementById('dark-mode').checked;
                const showGrid = document.getElementById('show-grid').checked;
                const showLabels = document.getElementById('show-labels').checked;
                const labelPos = document.getElementById('label-position').value;
                const labelSize = document.getElementById('label-size').value;
                const yAxisAuto = document.getElementById('y-axis-auto').checked;
                const yAxisMax = parseFloat(document.getElementById('y-axis-max').value);

                const fontColor = darkMode ? '#E2E8F0' : '#64748B';
                const gridColor = darkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';

                // Objeto base de opções.
                const options = {
                    responsive: true, maintainAspectRatio: false,
                    animation: {
                        delay: (context) => { // Adiciona uma animação de delay para um efeito de carregamento.
                            let delay = 0;
                            if (context.type === 'data' && context.mode === 'default') {
                                delay = context.dataIndex * 30 + context.datasetIndex * 100;
                            }
                            return delay;
                        }
                    },
                    plugins: { 
                        legend: { labels: { color: fontColor } },
                        datalabels: { 
                            display: showLabels, 
                            color: fontColor,
                            font: { size: labelSize || 12 },
                            align: labelPos === 'bottom' ? 'bottom' : (labelPos === 'center' ? 'center' : 'top'),
                            anchor: labelPos === 'bottom' ? 'start' : (labelPos === 'center' ? 'center' : 'end'),
                            formatter: (value, context) => {
                                return typeof value === 'object' ? (value.y || value.x) : value;
                            }
                        }
                    },
                    scales: {}
                };
                
                // Configurações específicas para os eixos X e Y.
                const axisOptions = { grid: { display: showGrid, color: gridColor }, ticks: { color: fontColor } };
                options.scales.x = { ...axisOptions };
                options.scales.y = { ...axisOptions, beginAtZero: true };

                if (chartType === 'horizontalBar') options.indexAxis = 'y'; // Inverte os eixos para barra horizontal.
                if (!yAxisAuto && !isNaN(yAxisMax)) options.scales.y.max = yAxisMax; // Define um máximo manual para o eixo Y.
                if (chartType === 'stacked' || chartType === 'stackedLine') { // Habilita o empilhamento.
                    options.scales.x.stacked = true;
                    options.scales.y.stacked = true;
                }

                return options;
            }

            // Atualiza a UI, mostrando/ocultando painéis de opções conforme o tipo de gráfico.
            function updateUI() {
                const chartType = document.querySelector('input[name="chart-type"]:checked').value;
                const isLine = ['line', 'combo', 'stackedLine'].includes(chartType);
                const isBar = ['bar', 'combo', 'stacked', 'horizontalBar'].includes(chartType);

                document.getElementById('line-options').style.display = isLine ? 'block' : 'none';
                document.getElementById('bar-options').style.display = isBar ? 'block' : 'none';
                
                renderChart();
            }

            // Gera uma imagem PNG do gráfico e a oferece para download.
            function downloadChart() {
                if (!chartInstance) {
                    alert('Gere um gráfico antes de tentar fazer o download.');
                    return;
                }
                const link = document.createElement('a');
                link.href = chartInstance.toBase64Image('image/png', 1.0); // Converte o canvas para imagem base64.
                link.download = 'power-graphx-chart.png';
                link.click();
            }
            
            // --- Inicialização e Eventos ---
            // Adiciona o primeiro controle de série e define todos os listeners de eventos.
            addSeriesControl(true);
            document.getElementById('controls').addEventListener('change', updateUI);
            document.getElementById('format-panel').addEventListener('change', renderChart);
            document.getElementById('add-series-btn').addEventListener('click', () => { addSeriesControl(false); updateUI(); });
            document.getElementById('update-charts-btn').addEventListener('click', renderChart);
            document.getElementById('download-chart-btn').addEventListener('click', downloadChart);
            
            document.getElementById('show-labels').addEventListener('change', (e) => {
                document.getElementById('label-options').style.display = e.target.checked ? 'block' : 'none';
            });
            document.getElementById('y-axis-auto').addEventListener('change', (e) => {
                document.getElementById('y-axis-max').disabled = e.target.checked;
            });
            
            updateUI(); // Chama a primeira renderização.
        }

        // Ponto de entrada do JavaScript: quando o HTML está pronto, a aplicação é inicializada.
        document.addEventListener('DOMContentLoaded', function() {
            try {
                // Os dados injetados pelo PowerShell são usados para iniciar o app.
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


# --- 6. Construção da Interface Gráfica (Windows Forms) ---
# Criação do formulário principal da aplicação.
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Power-Graphx Editor 1.7.9"
$Form.Width = 1200
$Form.Height = 800
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = 'Sizable'
$Form.WindowState = 'Maximized'
try { $Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\imageres.dll,25") } catch {}

# --- Menu Principal ---
$MenuStrip = New-Object System.Windows.Forms.MenuStrip
$MenuStrip.Dock = "Top"

# Menu Arquivo
$FileMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Arquivo")
$MenuLoadCsv = New-Object System.Windows.Forms.ToolStripMenuItem("Carregar CSV...")
$MenuGenerateHtml = New-Object System.Windows.Forms.ToolStripMenuItem("Gerar e Visualizar Relatório")
$MenuGenerateHtml.Enabled = $false
$MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem("Sair")
$FileMenu.DropDownItems.AddRange(@($MenuLoadCsv, $MenuGenerateHtml, (New-Object System.Windows.Forms.ToolStripSeparator), $MenuExit))

# Menu Editar
$EditMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Editar")
$MenuFind = New-Object System.Windows.Forms.ToolStripMenuItem("Localizar... (Ctrl+F)")
$EditMenu.DropDownItems.Add($MenuFind)

# Menu Dados
$DataMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Dados")
$DataMenu.Enabled = $false
$MenuSort = New-Object System.Windows.Forms.ToolStripMenuItem("Ordenar Coluna...")
$MenuFilter = New-Object System.Windows.Forms.ToolStripMenuItem("Filtrar Dados...")
$MenuRemoveFilter = New-Object System.Windows.Forms.ToolStripMenuItem("Remover Filtro")
$MenuCalculatedColumn = New-Object System.Windows.Forms.ToolStripMenuItem("Adicionar Coluna Calculada...")
$MenuRemoveColumn = New-Object System.Windows.Forms.ToolStripMenuItem("Remover Coluna...")
$DataMenu.DropDownItems.AddRange(@($MenuSort, $MenuFilter, $MenuRemoveFilter, (New-Object System.Windows.Forms.ToolStripSeparator), $MenuCalculatedColumn, $MenuRemoveColumn))

$MenuStrip.Items.AddRange(@($FileMenu, $EditMenu, $DataMenu))
$Form.Controls.Add($MenuStrip)


# --- Painel de Status ---
$StatusStrip = New-Object System.Windows.Forms.StatusStrip
$StatusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel("Aguardando arquivo CSV...")
$StatusStrip.Items.Add($StatusLabel)
$Form.Controls.Add($StatusStrip)


# --- Painel Principal ---
$MainPanel = New-Object System.Windows.Forms.Panel
$MainPanel.Dock = "Fill"
$Form.Controls.Add($MainPanel)


# --- Layout Principal dentro do Painel ---
$MainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$MainLayout.Dock = "Fill"
$MainLayout.ColumnCount = 1
$MainLayout.RowCount = 2
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 35)))
$MainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$MainPanel.Controls.Add($MainLayout)

# --- Painel de Busca ---
$SearchPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$SearchPanel.Dock = "Fill"
$SearchPanel.BackColor = [System.Drawing.Color]::FromArgb(220, 225, 230)
$SearchPanel.Visible = $false
$SearchPanel.Padding = 5
$MainLayout.Controls.Add($SearchPanel, 0, 0)

$SearchLabel = New-Object System.Windows.Forms.Label; $SearchLabel.Text = "Buscar:"; $SearchLabel.Font = "Segoe UI, 9"; $SearchLabel.Margin = "0,3,0,0"; $SearchPanel.Controls.Add($SearchLabel)
$SearchTextBox = New-Object System.Windows.Forms.TextBox; $SearchTextBox.Size = New-Object System.Drawing.Size(250, 23); $SearchPanel.Controls.Add($SearchTextBox)
$SearchButton = New-Object System.Windows.Forms.Button; $SearchButton.Text = "Buscar"; $SearchButton.Size = New-Object System.Drawing.Size(75, 25); $SearchPanel.Controls.Add($SearchButton)
$CloseSearchButton = New-Object System.Windows.Forms.Button; $CloseSearchButton.Text = "Fechar"; $CloseSearchButton.Size = New-Object System.Drawing.Size(75, 25); $SearchPanel.Controls.Add($CloseSearchButton)
$SearchResultLabel = New-Object System.Windows.Forms.Label; $SearchResultLabel.Text = ""; $SearchResultLabel.Font = "Segoe UI, 9"; $SearchResultLabel.Margin = "10,3,0,0"; $SearchPanel.Controls.Add($SearchResultLabel)

# --- Data Grid View (Grade de Dados) ---
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Dock = "Fill"
$DataGridView.BackgroundColor = [System.Drawing.Color]::White
$DataGridView.BorderStyle = "None"
$DataGridView.ColumnHeadersDefaultCellStyle.Font = "Segoe UI, 9, Bold"
$DataGridView.ReadOnly = $true
$DataGridView.AllowUserToAddRows = $false
$MainLayout.Controls.Add($DataGridView, 0, 1)

# --- Menu de Contexto (Clique com o botão direito na coluna) ---
$ContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$RenameMenuItem = $ContextMenu.Items.Add("Renomear Coluna...")
$RemoveContextMenuItem = $ContextMenu.Items.Add("Remover Coluna")


# --- 7. Eventos ---
# Define o que acontece quando cada item de menu ou botão é clicado.

$MenuLoadCsv.Add_Click({
    Load-CSVData -DataGridView $DataGridView -StatusLabel $StatusLabel -MenuItemsToEnable @($MenuGenerateHtml, $DataMenu)
})

$MenuGenerateHtml.Add_Click({
    Generate-HtmlReport -DataGridView $DataGridView -StatusLabel $StatusLabel
})

$MenuExit.Add_Click({ $Form.Close() })

$MenuFind.Add_Click({
    $SearchPanel.Visible = !$SearchPanel.Visible
    if ($SearchPanel.Visible) { $SearchTextBox.Focus() }
})

$MenuSort.Add_Click({ Sort-Data -DataGridView $DataGridView -StatusLabel $StatusLabel })

$MenuCalculatedColumn.Add_Click({ Add-CalculatedColumn -DataGridView $DataGridView -StatusLabel $StatusLabel })

$MenuRemoveFilter.Add_Click({ Remove-DataFilter -DataGridView $DataGridView -StatusLabel $StatusLabel })

# Evento para o clique com o botão direito no cabeçalho de uma coluna.
$DataGridView.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $Global:ColumnToModifyIndex = $e.ColumnIndex # Guarda o índice da coluna clicada.
        $ContextMenu.Show($DataGridView, $DataGridView.PointToClient([System.Windows.Forms.Cursor]::Position)) # Mostra o menu de contexto.
    }
})

# Evento para renomear a coluna selecionada.
$RenameMenuItem.Add_Click({
    if ($Global:ColumnToModifyIndex -ge 0) {
        $column = $DataGridView.Columns[$Global:ColumnToModifyIndex]
        $newName = Show-InputBox -Title "Renomear Coluna" -Prompt "Digite o novo nome para a coluna '$($column.HeaderText)':'" -DefaultText $column.HeaderText
        if ($newName) {
            $column.HeaderText = $newName # Altera o texto do cabeçalho.
        }
    }
})

# Evento para remover a coluna selecionada.
$RemoveContextMenuItem.Add_Click({
    if ($Global:ColumnToModifyIndex -ge 0) {
        $columnToRemove = $DataGridView.Columns[$Global:ColumnToModifyIndex].DataPropertyName
        $currentData = [System.Collections.ArrayList]$DataGridView.DataSource
        
        $StatusLabel.Text = "Removendo coluna..."; $StatusLabel.Owner.Refresh()

        # Usa Select-Object para criar uma nova lista de objetos sem a propriedade da coluna a ser removida.
        $newData = $currentData | Select-Object * -ExcludeProperty $columnToRemove
        
        # Re-vincula os dados à grade para refletir a remoção.
        $DataGridView.DataSource = $null
        $DataGridView.DataSource = [System.Collections.ArrayList]$newData
        
        # Atualiza também a cópia original dos dados se necessário.
        if ($Global:IsDataFiltered) {
             $Global:OriginalData = [System.Collections.ArrayList]($Global:OriginalData | Select-Object * -ExcludeProperty $columnToRemove)
        } else {
             $Global:OriginalData = $DataGridView.DataSource
        }

        $StatusLabel.Text = "Coluna '$columnToRemove' removida."
    }
})

# Evento global de pressionamento de tecla no formulário (para atalhos).
$Form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq 'F') { # Atalho Ctrl+F
        $e.SuppressKeyPress = $true # Impede o som de "bip" do Windows.
        $SearchPanel.Visible = !$SearchPanel.Visible
        if ($SearchPanel.Visible) { $SearchTextBox.Focus() }
    }
})

$CloseSearchButton.Add_Click({ $SearchPanel.Visible = $false })

# Evento do botão de busca.
$SearchButton.Add_Click({
    $searchTerm = $SearchTextBox.Text.ToLower()
    if ([string]::IsNullOrWhiteSpace($searchTerm)) { return }

    $DataGridView.ClearSelection()
    # Define estilos para células normais e células destacadas.
    $defaultCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $highlightCellStyle = New-Object System.Windows.Forms.DataGridViewCellStyle
    $highlightCellStyle.BackColor = [System.Drawing.Color]::Yellow

    $foundCount = 0
    # Itera sobre todas as células da grade.
    foreach ($row in $DataGridView.Rows) {
        foreach ($cell in $row.Cells) {
            $cell.Style = $defaultCellStyle # Reseta o estilo da célula.
            # Se a célula contiver o termo de busca, aplica o estilo de destaque.
            if ($cell.Value -and $cell.Value.ToString().ToLower().Contains($searchTerm)) {
                $cell.Style = $highlightCellStyle
                $foundCount++
            }
        }
    }
    $SearchResultLabel.Text = "$foundCount ocorrência(s) encontrada(s)."
})


# --- 8. Exibir a Janela ---
# Inicia a aplicação, exibindo o formulário e entrando no loop de eventos.
$Form.ShowDialog() | Out-Null

