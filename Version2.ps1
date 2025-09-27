# -----------------------------------------------------------------------------
# PowerChart Designer: Gerador de Gráfico Combo Interativo (Puro .NET)
# Funções: Leitura CSV robusta, Edição de Dados em tempo real, Gráfico Combo Dinâmico,
# Seleção de Tipo/Cor, Eixo Secundário Opcional e Salvamento de Imagem/CSV.
# -----------------------------------------------------------------------------

# --- 1. Carregar Assemblies Necessárias ---
try {
    # Assemblies para o gráfico, janela e cores
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "Não foi possível carregar as assemblies necessárias. Verifique a instalação do .NET Framework/Core."
    exit 1
}

# --- 2. Preparação dos Dados (Global State) ---
$script:LoadedData = @()       # Armazena os dados importados/editados
$script:ColumnNames = @()     # Armazena os nomes das colunas disponíveis
$script:ChartInstance = $null  # Armazena o objeto Chart para salvamento

# Tipos de gráfico disponíveis no ComboBox
$ChartTypes = [System.Enum]::GetNames([System.Windows.Forms.DataVisualization.Charting.SeriesChartType]) | Where-Object {
    $_ -notin @("PointAndFigure", "Stock", "Candlestick", "ErrorBar")
}

# --- 3. Funções de Utilitário e Sincronização ---

# Função para Carregar CSV (com identificação automática de delimitador)
Function Load-CSVData {
    param(
        [Parameter(Mandatory=$true)]$TextBoxFilePath,
        [Parameter(Mandatory=$true)]$DataGridView,
        [Parameter(Mandatory=$true)]$ComboXAxis,
        [Parameter(Mandatory=$true)]$ComboY1Data,
        [Parameter(Mandatory=$true)]$ComboY2Data
    )

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Selecione o arquivo CSV"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        $TextBoxFilePath.Text = $FilePath
        
        $Data = @()
        $DelimiterFound = ";" # Tenta ponto e vírgula primeiro

        # Tentativa 1: Ponto e Vírgula (Padrão PT-BR)
        try { $Data = Import-Csv -Path $FilePath -Delimiter ";" -ErrorAction Stop } catch { $Data = @() }

        # Tentativa 2: Vírgula (Padrão US/Internacional)
        if ($Data.Count -eq 0 -or $Data[0].PSObject.Properties.Name.Count -le 1) {
            try { 
                $Data = Import-Csv -Path $FilePath -Delimiter "," -ErrorAction Stop
                $DelimiterFound = ","
            } catch { $Data = @() }
        }

        if ($Data.Count -gt 0) {
            $script:LoadedData = $Data
            $script:ColumnNames = $script:LoadedData[0].PSObject.Properties.Name
            
            # Limpar e popular DataGridView
            $DataGridView.DataSource = $null
            $DataGridView.DataSource = $script:LoadedData
            $DataGridView.Refresh()

            # Configurar DataGridView para permitir edição
            $DataGridView.AllowUserToAddRows = $true
            $DataGridView.AllowUserToDeleteRows = $true

            # Atualizar ComboBoxes
            $ComboXAxis.Items.Clear(); $ComboY1Data.Items.Clear(); $ComboY2Data.Items.Clear()
            $ComboXAxis.Items.AddRange($script:ColumnNames)
            $ComboY1Data.Items.AddRange($script:ColumnNames)
            $ComboY2Data.Items.Add("Nenhum") # Opção para desabilitar
            $ComboY2Data.Items.AddRange($script:ColumnNames)
            
            # Tenta pré-selecionar colunas
            $ComboXAxis.SelectedIndex = 0
            if ($script:ColumnNames.Count -gt 1) { $ComboY1Data.SelectedIndex = 1 }
            $ComboY2Data.SelectedItem = "Nenhum" 
            
            # Formata colunas numéricas no grid para visualização
            foreach ($colName in $script:ColumnNames) {
                if ($DataGridView.Columns[$colName].ValueType -eq [System.Double] -or $DataGridView.Columns[$colName].ValueType -eq [System.Int32]) {
                    $DataGridView.Columns[$colName].DefaultCellStyle.Format = "N2"
                }
            }

        } else {
            [System.Windows.Forms.MessageBox]::Show("Erro: Arquivo CSV vazio ou formato de delimitador inválido.", "Erro de Leitura", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $script:LoadedData = @()
            $script:ColumnNames = @()
        }
    }
}

# Sincroniza as edições do DataGridView com a variável de dados do gráfico
Function Sync-DataGridToChart {
    param(
        [Parameter(Mandatory=$true)]$DataGridView
    )

    if ($DataGridView.DataSource) {
        # Exporta o DataSource (BindingList) de volta para um Array de PSCustomObject
        # Isso garante que as edições na grade sejam refletidas no $script:LoadedData
        $NewData = @()
        foreach ($item in $DataGridView.DataSource) {
            $NewData += $item
        }
        $script:LoadedData = $NewData
        
        # Atualiza a lista de nomes de colunas caso o usuário tenha adicionado/removido
        if ($script:LoadedData.Count -gt 0) {
            $script:ColumnNames = $script:LoadedData[0].PSObject.Properties.Name
        }
    }
}

# Função para Salvar o Gráfico
Function Save-ChartImage {
    param(
        [Parameter(Mandatory=$true)]$Chart
    )
    # [Restante da função Save-ChartImage...]
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg|Bitmap (*.bmp)|*.bmp"
    $SaveFileDialog.Title = "Salvar Gráfico como Imagem"

    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $SaveFileDialog.FileName
        $Format = [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png

        if ($FilePath.EndsWith(".jpg", [System.StringComparison]::OrdinalIgnoreCase) -or $FilePath.EndsWith(".jpeg", [System.StringComparison]::OrdinalIgnoreCase)) {
            $Format = [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Jpeg
        } elseif ($FilePath.EndsWith(".bmp", [System.StringComparison]::OrdinalIgnoreCase)) {
            $Format = [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Bmp
        }

        try {
            $Chart.SaveImage($FilePath, $Format)
            [System.Windows.Forms.MessageBox]::Show("Gráfico salvo com sucesso em: $FilePath", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao salvar o arquivo: $($_.Exception.Message)", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
}

# Função para Salvar CSV Editado
Function Save-EditedCSV {
    param(
        [Parameter(Mandatory=$true)]$DataGridView
    )
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV (Ponto e Vírgula) (*.csv)|*.csv"
    $SaveFileDialog.Title = "Salvar Edições do CSV"
    
    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $SaveFileDialog.FileName
        
        # Sincroniza antes de salvar para garantir que as últimas edições estejam no $script:LoadedData
        Sync-DataGridToChart -DataGridView $DataGridView

        try {
            # Exporta usando ponto e vírgula, facilitando a leitura em Excel
            $script:LoadedData | Export-Csv -Path $FilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("CSV editado salvo com sucesso em: $FilePath", "Sucesso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao salvar o CSV: $($_.Exception.Message)", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
}

# --- 4. Inicialização da Janela e Controles (Design Moderno) ---

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerChart Designer (CSV Dynamic)"
$Form.Width = 1400
$Form.Height = 850
$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 245) 

# Componente SplitContainer para dividir a tela: Controles à esquerda, Visualização à direita
$SplitContainer = New-Object System.Windows.Forms.SplitContainer
$SplitContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$SplitContainer.SplitterDistance = 320 # Largura do painel de controle
$SplitContainer.Orientation = [System.Windows.Forms.Orientation]::Vertical
$SplitContainer.IsSplitterFixed = $true
$SplitContainer.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$Form.Controls.Add($SplitContainer)

# Painel de Controles (Sidebar - Left Panel)
$PanelControls = $SplitContainer.Panel1
$PanelControls.BackColor = [System.Drawing.Color]::FromArgb(40, 50, 60) # Azul escuro/quase preto
$PanelControls.Padding = New-Object System.Windows.Forms.Padding(10)
$PanelControls.AutoScroll = $true

# Painel de Visualização (Right Panel)
$PanelVisualization = $SplitContainer.Panel2
$PanelVisualization.Padding = New-Object System.Windows.Forms.Padding(10)

# --- Controles dentro do Painel Lateral (Sidebar) ---
$FlowLayoutPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$FlowLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$FlowLayoutPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
$FlowLayoutPanel.WrapContents = $false
$FlowLayoutPanel.AutoScroll = $true
$PanelControls.Controls.Add($FlowLayoutPanel)

# Função auxiliar para criar labels estilizados
function New-StyledLabel {
    param($Text, $Bold = $false)
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Text
    $Label.AutoSize = $true
    $Label.ForeColor = [System.Drawing.Color]::White
    $Label.Font = New-Object System.Drawing.Font("Segoe UI", 10, @($Bold, 0)[$Bold -eq $false])
    $Label.Margin = New-Object System.Windows.Forms.Padding(5, 10, 5, 5) # Top margin for spacing
    $FlowLayoutPanel.Controls.Add($Label)
    return $Label
}

# --- Título do Painel ---
$Title = New-StyledLabel -Text "PowerChart Designer" -Bold $true
$Title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$Title.Margin = New-Object System.Windows.Forms.Padding(5, 15, 5, 25)

# ----------------- Importação CSV -----------------
New-StyledLabel -Text "1. Fonte de Dados CSV:" -Bold $true

$TextBoxFilePath = New-Object System.Windows.Forms.TextBox
$TextBoxFilePath.Width = 290
$TextBoxFilePath.ReadOnly = $true
$TextBoxFilePath.BackColor = [System.Drawing.Color]::Gainsboro
$FlowLayoutPanel.Controls.Add($TextBoxFilePath)

$ButtonSelectFile = New-Object System.Windows.Forms.Button
$ButtonSelectFile.Text = "Abrir CSV"
$ButtonSelectFile.Width = 140
$ButtonSelectFile.Height = 30
$ButtonSelectFile.BackColor = [System.Drawing.Color]::FromArgb(100, 150, 255) # Azul vibrante
$ButtonSelectFile.ForeColor = [System.Drawing.Color]::White
$ButtonSelectFile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$FlowLayoutPanel.Controls.Add($ButtonSelectFile)

# ----------------- Seleção de Eixos (Dados) -----------------
New-StyledLabel -Text "2. Seleção de Eixos (Dados):" -Bold $true

# Eixo X
New-StyledLabel -Text "Eixo X (Rótulos):" 
$ComboXAxis = New-Object System.Windows.Forms.ComboBox
$ComboXAxis.Width = 290
$FlowLayoutPanel.Controls.Add($ComboXAxis)

# Série 1 (Y1)
New-StyledLabel -Text "Série 1 (Eixo Primário Y1):" 
$ComboY1Data = New-Object System.Windows.Forms.ComboBox
$ComboY1Data.Width = 290
$FlowLayoutPanel.Controls.Add($ComboY1Data)

# Série 2 (Y2)
New-StyledLabel -Text "Série 2 (Eixo Secundário Y2):" 
$ComboY2Data = New-Object System.Windows.Forms.ComboBox
$ComboY2Data.Width = 290
$FlowLayoutPanel.Controls.Add($ComboY2Data)

# ----------------- Configurações de Aparência (Y1) -----------------
New-StyledLabel -Text "3. Aparência Série 1 (Y1):" -Bold $true

# Tipo de Gráfico - Y1
New-StyledLabel -Text "Tipo Y1:" 
$ComboY1Type = New-Object System.Windows.Forms.ComboBox
$ComboY1Type.Width = 290
$ComboY1Type.Items.AddRange($ChartTypes)
$ComboY1Type.SelectedItem = "Column" 
$FlowLayoutPanel.Controls.Add($ComboY1Type)

# Cor - Y1
$ButtonY1Color = New-Object System.Windows.Forms.Button
$ButtonY1Color.Text = "Cor Y1: DeepSkyBlue"
$ButtonY1Color.Width = 290
$ButtonY1Color.Tag = [System.Drawing.Color]::DeepSkyBlue
$ButtonY1Color.BackColor = [System.Drawing.Color]::DeepSkyBlue
$ButtonY1Color.ForeColor = [System.Drawing.Color]::White
$ButtonY1Color.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ButtonY1Color.Add_Click({
    $ColorDialog = New-Object System.Windows.Forms.ColorDialog
    if ($ColorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ButtonY1Color.Tag = $ColorDialog.Color
        $ButtonY1Color.BackColor = $ColorDialog.Color
        $ButtonY1Color.Text = "Cor Y1: $($ColorDialog.Color.Name)"
    }
})
$FlowLayoutPanel.Controls.Add($ButtonY1Color)

# ----------------- Configurações de Aparência (Y2) -----------------
New-StyledLabel -Text "4. Aparência Série 2 (Y2):" -Bold $true

# Tipo de Gráfico - Y2
New-StyledLabel -Text "Tipo Y2:" 
$ComboY2Type = New-Object System.Windows.Forms.ComboBox
$ComboY2Type.Width = 290
$ComboY2Type.Items.AddRange($ChartTypes)
$ComboY2Type.SelectedItem = "Line" 
$FlowLayoutPanel.Controls.Add($ComboY2Type)

# Cor - Y2
$ButtonY2Color = New-Object System.Windows.Forms.Button
$ButtonY2Color.Text = "Cor Y2: OrangeRed"
$ButtonY2Color.Width = 290
$ButtonY2Color.Tag = [System.Drawing.Color]::OrangeRed
$ButtonY2Color.BackColor = [System.Drawing.Color]::OrangeRed
$ButtonY2Color.ForeColor = [System.Drawing.Color]::White
$ButtonY2Color.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$ButtonY2Color.Add_Click({
    $ColorDialog = New-Object System.Windows.Forms.ColorDialog
    if ($ColorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ButtonY2Color.Tag = $ColorDialog.Color
        $ButtonY2Color.BackColor = $ColorDialog.Color
        $ButtonY2Color.Text = "Cor Y2: $($ColorDialog.Color.Name)"
    }
})
$FlowLayoutPanel.Controls.Add($ButtonY2Color)

# --- Botões de Ação (Bottom Section) ---
$PanelAction = New-Object System.Windows.Forms.Panel
$PanelAction.Height = 150 # Altura ajustada para 3 botões
$PanelAction.Width = 300
$FlowLayoutPanel.Controls.Add($PanelAction)

$ButtonUpdate = New-Object System.Windows.Forms.Button
$ButtonUpdate.Name = "ButtonUpdate"
$ButtonUpdate.Text = "ATUALIZAR GRÁFICO"
$ButtonUpdate.Location = New-Object System.Drawing.Point(5, 5)
$ButtonUpdate.Width = 290
$ButtonUpdate.Height = 40
$ButtonUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$ButtonUpdate.BackColor = [System.Drawing.Color]::FromArgb(0, 170, 255) # Azul mais claro
$ButtonUpdate.ForeColor = [System.Drawing.Color]::White
$ButtonUpdate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PanelAction.Controls.Add($ButtonUpdate)

$ButtonSaveImage = New-Object System.Windows.Forms.Button
$ButtonSaveImage.Text = "SALVAR GRÁFICO (Imagem)"
$ButtonSaveImage.Location = New-Object System.Drawing.Point(5, 55)
$ButtonSaveImage.Width = 290
$ButtonSaveImage.Height = 30
$ButtonSaveImage.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$ButtonSaveImage.BackColor = [System.Drawing.Color]::LightGray
$ButtonSaveImage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PanelAction.Controls.Add($ButtonSaveImage)

$ButtonSaveCSV = New-Object System.Windows.Forms.Button
$ButtonSaveCSV.Text = "SALVAR CSV (Editado)"
$ButtonSaveCSV.Location = New-Object System.Drawing.Point(5, 95)
$ButtonSaveCSV.Width = 290
$ButtonSaveCSV.Height = 30
$ButtonSaveCSV.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$ButtonSaveCSV.BackColor = [System.Drawing.Color]::LightGray
$ButtonSaveCSV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PanelAction.Controls.Add($ButtonSaveCSV)


# --- 5. Objeto de Gráfico e Editor de Dados (Painel de Visualização) ---

# SplitContainer Vertical para Gráfico (Top) e DataGrid (Bottom)
$VizSplitter = New-Object System.Windows.Forms.SplitContainer
$VizSplitter.Dock = [System.Windows.Forms.DockStyle]::Fill
$VizSplitter.Orientation = [System.Windows.Forms.Orientation]::Horizontal
$VizSplitter.SplitterDistance = 450 # Altura inicial do gráfico
$VizSplitter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$PanelVisualization.Controls.Add($VizSplitter)

# 5.1. Objeto de Gráfico (Top Panel)
$Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$Chart.Dock = [System.Windows.Forms.DockStyle]::Fill
$Chart.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 255)
$Chart.BorderSkin = New-Object System.Windows.Forms.DataVisualization.Charting.BorderSkin
$Chart.BorderSkin.SkinStyle = [System.Windows.Forms.DataVisualization.Charting.BorderSkinStyle]::Emboss
$VizSplitter.Panel1.Controls.Add($Chart)
$script:ChartInstance = $Chart # Salva a referência

# 5.2. Data Grid (Bottom Panel)
$DataGridView = New-Object System.Windows.Forms.DataGridView
$DataGridView.Dock = [System.Windows.Forms.DockStyle]::Fill
$DataGridView.BackgroundColor = [System.Drawing.Color]::White
$DataGridView.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(240, 245, 255)
$DataGridView.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
$DataGridView.AutoGenerateColumns = $true
$VizSplitter.Panel2.Controls.Add($DataGridView)

# --- 6. Função de Atualização do Gráfico ---
Function Update-Chart {
    param(
        [Parameter(Mandatory=$true)]$Chart,
        [Parameter(Mandatory=$true)]$ComboY1Type,
        [Parameter(Mandatory=$true)]$ButtonY1Color,
        [Parameter(Mandatory=$true)]$ComboY2Type,
        [Parameter(Mandatory=$true)]$ButtonY2Color,
        [Parameter(Mandatory=$true)]$ComboXAxis,
        [Parameter(Mandatory=$true)]$ComboY1Data,
        [Parameter(Mandatory=$true)]$ComboY2Data
    )

    # 6.0. Sincroniza dados do grid antes de plotar
    Sync-DataGridToChart -DataGridView $DataGridView

    $Y2ColumnName = $ComboY2Data.SelectedItem

    if (-not $script:LoadedData -or $script:LoadedData.Count -eq 0) {
        # ... Mensagem de instrução ...
        $Chart.Titles.Clear()
        $Chart.Series.Clear()
        $Chart.ChartAreas.Clear()
        $Chart.Titles.Add("Carregue um Arquivo CSV e Selecione as Colunas")
        $Chart.Titles[0].Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
        $Chart.Titles[0].ForeColor = [System.Drawing.Color]::Gray
        return
    }

    if (-not $ComboXAxis.SelectedItem -or -not $ComboY1Data.SelectedItem) {
        [System.Windows.Forms.MessageBox]::Show("Selecione as colunas para os Eixos X e Y1.", "Seleção Incompleta", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    # 6.1. Limpar e Reconfigurar
    $Chart.Titles.Clear()
    $Chart.Series.Clear()
    $Chart.ChartAreas.Clear()
    $Chart.Legends.Clear()

    # Nomes das colunas selecionadas
    $XColumnName = $ComboXAxis.SelectedItem
    $Y1ColumnName = $ComboY1Data.SelectedItem
    
    # Título do Gráfico
    $TitleText = if ($Y2ColumnName -and $Y2ColumnName -ne "Nenhum") {
        "Gráfico Combinado: $Y1ColumnName (Y1) e $Y2ColumnName (Y2) por $XColumnName"
    } else {
        "Gráfico Simples: $Y1ColumnName por $XColumnName"
    }
    $Chart.Titles.Add($TitleText)
    $Chart.Titles[0].Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)

    # Configurar a Área de Plotagem
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea("MainArea")
    $ChartArea.BackColor = [System.Drawing.Color]::White
    $ChartArea.BorderColor = [System.Drawing.Color]::LightGray
    $ChartArea.BorderWidth = 1
    
    # Configurar Eixo X (Rótulos)
    $ChartArea.AxisX.Title = $XColumnName
    $ChartArea.AxisX.MajorGrid.Enabled = $false
    $ChartArea.AxisX.Interval = 1
    $ChartArea.AxisX.LabelStyle.Angle = -45

    # Configurar Eixo Y Primário (Y1)
    $ChartArea.AxisY.Title = "$Y1ColumnName (Eixo Primário)"
    $ChartArea.AxisY.MajorGrid.Enabled = $false # Remove a grade do Y1
    
    # Configurar Eixo Y Secundário (Y2) - HABILITADO APENAS SE HOUVER SELEÇÃO
    $IsY2Enabled = ($Y2ColumnName -and $Y2ColumnName -ne "Nenhum")
    $ChartArea.AxisY2.Enabled = [System.Windows.Forms.DataVisualization.Charting.AxisEnabled]::False
    if ($IsY2Enabled) {
        $ChartArea.AxisY2.Enabled = [System.Windows.Forms.DataVisualization.Charting.AxisEnabled]::True
        $ChartArea.AxisY2.Title = "$Y2ColumnName (Eixo Secundário)"
        $ChartArea.AxisY2.MajorGrid.Enabled = $false # Remove a grade do Y2
    }

    $Chart.ChartAreas.Add($ChartArea)

    # 6.2. Série 1: Dados (Y1)
    $SeriesY1 = New-Object System.Windows.Forms.DataVisualization.Charting.Series($Y1ColumnName)
    $SeriesY1.ChartType = $ComboY1Type.SelectedItem
    $SeriesY1.Color = $ButtonY1Color.Tag
    $SeriesY1.IsValueShownAsLabel = $true
    
    # 6.3. Adicionar Pontos Dinamicamente
    $i = 0
    foreach ($dataRow in $script:LoadedData) {
        $xLabel = $dataRow.$XColumnName
        
        # PARSING ROBUSTO PARA Y1: Troca vírgula por ponto para conversão Double
        $Y1String = $dataRow.$Y1ColumnName -replace ",", "."
        try { $y1Value = [double]::Parse($Y1String, [System.Globalization.CultureInfo]::InvariantCulture) } catch { $y1Value = 0 }
        
        # Ponto Série 1 (Y1)
        $point1 = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint($i, $y1Value)
        $point1.AxisLabel = $xLabel
        $SeriesY1.Points.Add($point1)

        # Se Y2 estiver habilitado, adicione a Série 2
        if ($IsY2Enabled) {
            # PARSING ROBUSTO PARA Y2
            $Y2String = $dataRow.$Y2ColumnName -replace ",", "."
            try { $y2Value = [double]::Parse($Y2String, [System.Globalization.CultureInfo]::InvariantCulture) } catch { $y2Value = 0 }

            # Ponto Série 2 (Y2)
            $point2 = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint($i, $y2Value)
            $point2.Label = "$y2Value" 
            $SeriesY2.Points.Add($point2)
        }
        
        $i++
    }
    
    $Chart.Series.Add($SeriesY1)

    # 6.4. Série 2: Dados (Y2) - Só adiciona se estiver habilitado
    if ($IsY2Enabled) {
        $SeriesY2 = New-Object System.Windows.Forms.DataVisualization.Charting.Series($Y2ColumnName)
        $SeriesY2.ChartType = $ComboY2Type.SelectedItem
        $SeriesY2.YAxisType = [System.Windows.Forms.DataVisualization.Charting.AxisType]::Secondary 
        $SeriesY2.Color = $ButtonY2Color.Tag
        $SeriesY2.BorderWidth = 3
        $SeriesY2.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle
        $SeriesY2.MarkerSize = 7
        $SeriesY2.IsValueShownAsLabel = $true
        $Chart.Series.Add($SeriesY2)
    }


    # 6.5. Configurar Legenda
    $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.Docking = [System.Windows.Forms.DataVisualization.Charting.Docking]::Bottom
    $Legend.Alignment = [System.Windows.Forms.DataVisualization.Charting.StringAlignment]::Center
    $Legend.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $Chart.Legends.Add($Legend)
}

# --- 7. Eventos dos Botões ---

$ButtonSelectFile.Add_Click({
    Load-CSVData `
        -TextBoxFilePath $TextBoxFilePath `
        -DataGridView $DataGridView `
        -ComboXAxis $ComboXAxis `
        -ComboY1Data $ComboY1Data `
        -ComboY2Data $ComboY2Data
    $ButtonUpdate.PerformClick() # Atualiza o gráfico automaticamente
})

$ButtonUpdate.Add_Click({
    Update-Chart `
        -Chart $Chart `
        -ComboY1Type $ComboY1Type `
        -ButtonY1Color $ButtonY1Color `
        -ComboY2Type $ComboY2Type `
        -ButtonY2Color $ButtonY2Color `
        -ComboXAxis $ComboXAxis `
        -ComboY1Data $ComboY1Data `
        -ComboY2Data $ComboY2Data
})

$ButtonSaveImage.Add_Click({
    if ($script:ChartInstance) {
        Save-ChartImage -Chart $script:ChartInstance
    } else {
        [System.Windows.Forms.MessageBox]::Show("Nenhum gráfico para salvar.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
})

$ButtonSaveCSV.Add_Click({
    if ($DataGridView.DataSource) {
        Save-EditedCSV -DataGridView $DataGridView
    } else {
        [System.Windows.Forms.MessageBox]::Show("Nenhum dado para salvar.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
})

# Desenhar a instrução inicial
$Form.Add_Load({
    Update-Chart `
        -Chart $Chart `
        -ComboY1Type $ComboY1Type `
        -ButtonY1Color $ButtonY1Color `
        -ComboY2Type $ComboY2Type `
        -ButtonY2Color $ButtonY2Color `
        -ComboXAxis $ComboXAxis `
        -ComboY1Data $ComboY1Data `
        -ComboY2Data $ComboY2Data
})


# Exibe a janela
$Form.ShowDialog() | Out-Null
