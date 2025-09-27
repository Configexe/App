# -----------------------------------------------------------------------------
# Gráfico Combinado Interativo (CSV Dynamic) - Com Salvar Imagem e Design Moderno
# Permite ao usuário escolher dados, tipo, cor e se a série Y2 será exibida.
# -----------------------------------------------------------------------------

# --- 1. Carregar Assemblies Necessárias ---
try {
    # Assembly para o gráfico
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    # Assembly para a janela (Forms) e cores
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Error "Não foi possível carregar as assemblies necessárias. Verifique a instalação do .NET Framework/Core."
    exit 1
}

# --- 2. Preparação dos Dados (Global State) ---
$script:LoadedData = @()     # Armazena os dados importados do CSV
$script:ColumnNames = @()    # Armazena os nomes das colunas disponíveis
$ChartTypes = [System.Enum]::GetNames([System.Windows.Forms.DataVisualization.Charting.SeriesChartType]) | Where-Object {
    $_ -notin @("PointAndFigure", "Stock", "Candlestick", "ErrorBar") # Remove tipos complexos
}

# --- 3. Funções de Utilitário ---

# Função para Carregar CSV
Function Load-CSVData {
    param(
        [Parameter(Mandatory=$true)]$Form,
        [Parameter(Mandatory=$true)]$TextBoxFilePath,
        [Parameter(Mandatory=$true)]$ComboXAxis,
        [Parameter(Mandatory=$true)]$ComboY1Data,
        [Parameter(Mandatory=$true)]$ComboY2Data,
        [Parameter(Mandatory=$true)]$ButtonUpdate
    )

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Selecione o arquivo CSV"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        $TextBoxFilePath.Text = $FilePath
        
        # Tenta carregar com o delimitador ";" e depois ","
        try {
            $Data = Import-Csv -Path $FilePath -Delimiter ";"
            if ($Data.Count -eq 0 -or $Data[0].PSObject.Properties.Name.Count -le 1) {
                $Data = Import-Csv -Path $FilePath -Delimiter ","
            }
            
            $script:LoadedData = $Data
            
            if ($script:LoadedData.Count -gt 0) {
                $script:ColumnNames = $script:LoadedData[0].PSObject.Properties.Name
                
                # Atualizar ComboBoxes
                $ComboXAxis.Items.Clear(); $ComboY1Data.Items.Clear(); $ComboY2Data.Items.Clear()
                $ComboXAxis.Items.AddRange($script:ColumnNames)
                $ComboY1Data.Items.AddRange($script:ColumnNames)
                # Inclui uma opção para desabilitar a Série Y2
                $ComboY2Data.Items.Add("Nenhum") 
                $ComboY2Data.Items.AddRange($script:ColumnNames)
                
                # Tenta pré-selecionar colunas
                $ComboXAxis.SelectedIndex = 0
                if ($script:ColumnNames.Count -gt 1) { $ComboY1Data.SelectedIndex = 1 }
                $ComboY2Data.SelectedItem = "Nenhum" # Começa desabilitado

                $ButtonUpdate.PerformClick()
                
            } else {
                [System.Windows.Forms.MessageBox]::Show("O arquivo CSV está vazio ou ilegível.", "Erro de Dados", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                $script:LoadedData = @()
                $script:ColumnNames = @()
            }

        } catch {
            [System.Windows.Forms.MessageBox]::Show("Erro ao ler o arquivo CSV: $($_.Exception.Message)", "Erro de Leitura", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    }
}

# Função para Salvar o Gráfico
Function Save-ChartImage {
    param(
        [Parameter(Mandatory=$true)]$Chart
    )

    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "PNG Image (*.png)|*.png|JPEG Image (*.jpg)|*.jpg|Bitmap (*.bmp)|*.bmp"
    $SaveFileDialog.Title = "Salvar Gráfico como Imagem"

    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $SaveFileDialog.FileName
        $Format = [System.Windows.Forms.DataVisualization.Charting.ChartImageFormat]::Png

        # Determina o formato de imagem com base na extensão
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

# --- 4. Inicialização da Janela e Controles (Design Moderno) ---

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerChart Designer (CSV Dynamic)"
$Form.Width = 1300
$Form.Height = 780
$Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$Form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 245) # Cinza Claro suave

# --- Painel de Controles (Sidebar) ---
$PanelControls = New-Object System.Windows.Forms.Panel
$PanelControls.Dock = [System.Windows.Forms.DockStyle]::Left
$PanelControls.Width = 320
$PanelControls.BackColor = [System.Drawing.Color]::FromArgb(40, 50, 60) # Azul escuro/quase preto
$PanelControls.Padding = New-Object System.Windows.Forms.Padding(10)
$Form.Controls.Add($PanelControls)

# Função auxiliar para criar labels estilizados
function New-StyledLabel {
    param($Text, $Y, $Bold = $false)
    $Label = New-Object System.Windows.Forms.Label
    $Label.Text = $Text
    $Label.Location = New-Object System.Drawing.Point(10, $Y)
    $Label.AutoSize = $true
    $Label.ForeColor = [System.Drawing.Color]::White
    $Label.Font = New-Object System.Drawing.Font("Segoe UI", 10, @($Bold, 0)[$Bold -eq $false])
    $PanelControls.Controls.Add($Label)
    return $Label
}

# Variável de posicionamento vertical
$YPosition = 15

# --- Título do Painel ---
$Title = New-StyledLabel -Text "PowerChart Designer" -Y $YPosition -Bold $true
$Title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$YPosition += 50

# ----------------- Importação CSV -----------------
$LabelCSV = New-StyledLabel -Text "1. Fonte de Dados CSV:" -Y $YPosition -Bold $true
$YPosition += 25

$TextBoxFilePath = New-Object System.Windows.Forms.TextBox
$TextBoxFilePath.Location = New-Object System.Drawing.Point(10, $YPosition)
$TextBoxFilePath.Width = 200
$TextBoxFilePath.ReadOnly = $true
$TextBoxFilePath.BackColor = [System.Drawing.Color]::Gainsboro
$PanelControls.Controls.Add($TextBoxFilePath)

$ButtonSelectFile = New-Object System.Windows.Forms.Button
$ButtonSelectFile.Text = "Abrir CSV"
$ButtonSelectFile.Location = New-Object System.Drawing.Point(215, $YPosition)
$ButtonSelectFile.Width = 85
$ButtonSelectFile.BackColor = [System.Drawing.Color]::FromArgb(100, 150, 255) # Azul vibrante
$ButtonSelectFile.ForeColor = [System.Drawing.Color]::White
$ButtonSelectFile.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PanelControls.Controls.Add($ButtonSelectFile)
$YPosition += 45

# ----------------- Seleção de Eixos (Dados) -----------------
$LabelAxes = New-StyledLabel -Text "2. Seleção de Eixos:" -Y $YPosition -Bold $true
$YPosition += 25

# Eixo X
New-StyledLabel -Text "Eixo X (Rótulos):" -Y $YPosition
$YPosition += 20
$ComboXAxis = New-Object System.Windows.Forms.ComboBox
$ComboXAxis.Location = New-Object System.Drawing.Point(10, $YPosition)
$ComboXAxis.Width = 290
$PanelControls.Controls.Add($ComboXAxis)
$YPosition += 40

# Série 1 (Y1)
New-StyledLabel -Text "Série 1 (Eixo Primário Y1):" -Y $YPosition
$YPosition += 20
$ComboY1Data = New-Object System.Windows.Forms.ComboBox
$ComboY1Data.Location = New-Object System.Drawing.Point(10, $YPosition)
$ComboY1Data.Width = 290
$PanelControls.Controls.Add($ComboY1Data)
$YPosition += 45

# Série 2 (Y2)
New-StyledLabel -Text "Série 2 (Eixo Secundário Y2):" -Y $YPosition
$YPosition += 20
$ComboY2Data = New-Object System.Windows.Forms.ComboBox
$ComboY2Data.Location = New-Object System.Drawing.Point(10, $YPosition)
$ComboY2Data.Width = 290
$PanelControls.Controls.Add($ComboY2Data)
$YPosition += 45

# ----------------- Configurações de Aparência (Y1) -----------------
$LabelAesthetics = New-StyledLabel -Text "3. Aparência Y1:" -Y $YPosition -Bold $true
$YPosition += 25

# Tipo de Gráfico - Y1
New-StyledLabel -Text "Tipo Y1:" -Y $YPosition
$YPosition += 20
$ComboY1Type = New-Object System.Windows.Forms.ComboBox
$ComboY1Type.Location = New-Object System.Drawing.Point(10, $YPosition)
$ComboY1Type.Width = 290
$ComboY1Type.Items.AddRange($ChartTypes)
$ComboY1Type.SelectedItem = "Column" 
$PanelControls.Controls.Add($ComboY1Type)
$YPosition += 40

# Cor - Y1
$ButtonY1Color = New-Object System.Windows.Forms.Button
$ButtonY1Color.Text = "Cor Y1 (Coluna): DeepSkyBlue"
$ButtonY1Color.Location = New-Object System.Drawing.Point(10, $YPosition)
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
        $ButtonY1Color.Text = "Cor Y1 (Coluna): $($ColorDialog.Color.Name)"
    }
})
$PanelControls.Controls.Add($ButtonY1Color)
$YPosition += 45

# ----------------- Configurações de Aparência (Y2) -----------------
$LabelAesthetics2 = New-StyledLabel -Text "4. Aparência Y2:" -Y $YPosition -Bold $true
$YPosition += 25

# Tipo de Gráfico - Y2
New-StyledLabel -Text "Tipo Y2:" -Y $YPosition
$YPosition += 20
$ComboY2Type = New-Object System.Windows.Forms.ComboBox
$ComboY2Type.Location = New-Object System.Drawing.Point(10, $YPosition)
$ComboY2Type.Width = 290
$ComboY2Type.Items.AddRange($ChartTypes)
$ComboY2Type.SelectedItem = "Line" 
$PanelControls.Controls.Add($ComboY2Type)
$YPosition += 40

# Cor - Y2
$ButtonY2Color = New-Object System.Windows.Forms.Button
$ButtonY2Color.Text = "Cor Y2 (Linha): OrangeRed"
$ButtonY2Color.Location = New-Object System.Drawing.Point(10, $YPosition)
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
        $ButtonY2Color.Text = "Cor Y2 (Linha): $($ColorDialog.Color.Name)"
    }
})
$PanelControls.Controls.Add($ButtonY2Color)
$YPosition += 60


# --- Botões de Ação (Bottom) ---
$PanelAction = New-Object System.Windows.Forms.Panel
$PanelAction.Dock = [System.Windows.Forms.DockStyle]::Bottom
$PanelAction.Height = 50
$PanelAction.BackColor = [System.Drawing.Color]::FromArgb(40, 50, 60)
$Form.Controls.Add($PanelAction)

$ButtonUpdate = New-Object System.Windows.Forms.Button
$ButtonUpdate.Name = "ButtonUpdate"
$ButtonUpdate.Text = "ATUALIZAR GRÁFICO"
$ButtonUpdate.Location = New-Object System.Drawing.Point(10, 10)
$ButtonUpdate.Width = 190
$ButtonUpdate.Height = 30
$ButtonUpdate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$ButtonUpdate.BackColor = [System.Drawing.Color]::FromArgb(0, 170, 255) # Azul mais claro
$ButtonUpdate.ForeColor = [System.Drawing.Color]::White
$ButtonUpdate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PanelAction.Controls.Add($ButtonUpdate)

$ButtonSaveImage = New-Object System.Windows.Forms.Button
$ButtonSaveImage.Text = "SALVAR GRÁFICO"
$ButtonSaveImage.Location = New-Object System.Drawing.Point(210, 10)
$ButtonSaveImage.Width = 100
$ButtonSaveImage.Height = 30
$ButtonSaveImage.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$ButtonSaveImage.BackColor = [System.Drawing.Color]::LightGray
$ButtonSaveImage.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$PanelAction.Controls.Add($ButtonSaveImage)


# --- 5. Objeto de Gráfico (Central) ---
$Chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart
$Chart.Dock = [System.Windows.Forms.DockStyle]::Fill
$Chart.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 255) # Fundo levemente azulado
$Chart.BorderSkin = New-Object System.Windows.Forms.DataVisualization.Charting.BorderSkin
$Chart.BorderSkin.SkinStyle = [System.Windows.Forms.DataVisualization.Charting.BorderSkinStyle]::Emboss # Efeito 3D leve
$Form.Controls.Add($Chart)


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

    $Y2ColumnName = $ComboY2Data.SelectedItem

    if (-not $script:LoadedData -or $script:LoadedData.Count -eq 0) {
        # ... Mensagem de instrução ...
        $Chart.Titles.Clear()
        $Chart.Series.Clear()
        $Chart.ChartAreas.Clear()
        $Chart.Legends.Clear()
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
    $ChartArea.AxisY.MajorGrid.Enabled = $false
    
    # Configurar Eixo Y Secundário (Y2) - HABILITADO APENAS SE HOUVER SELEÇÃO
    $IsY2Enabled = ($Y2ColumnName -and $Y2ColumnName -ne "Nenhum")
    $ChartArea.AxisY2.Enabled = [System.Windows.Forms.DataVisualization.Charting.AxisEnabled]::False
    if ($IsY2Enabled) {
        $ChartArea.AxisY2.Enabled = [System.Windows.Forms.DataVisualization.Charting.AxisEnabled]::True
        $ChartArea.AxisY2.Title = "$Y2ColumnName (Eixo Secundário)"
        $ChartArea.AxisY2.MajorGrid.Enabled = $false
        # Você pode adicionar formatação de % aqui se souber que os dados são porcentagens:
        # $ChartArea.AxisY2.LabelStyle.Format = "P0" 
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
        
        try { $y1Value = [double]::Parse($dataRow.$Y1ColumnName) } catch { $y1Value = 0 }
        
        # Ponto Série 1 (Y1)
        $point1 = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint($i, $y1Value)
        $point1.AxisLabel = $xLabel
        $SeriesY1.Points.Add($point1)

        # Se Y2 estiver habilitado, adicione a Série 2
        if ($IsY2Enabled) {
            try { $y2Value = [double]::Parse($dataRow.$Y2ColumnName) } catch { $y2Value = 0 }

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
        -Form $Form `
        -TextBoxFilePath $TextBoxFilePath `
        -ComboXAxis $ComboXAxis `
        -ComboY1Data $ComboY1Data `
        -ComboY2Data $ComboY2Data `
        -ButtonUpdate $ButtonUpdate
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
    if (-not $script:LoadedData -or -not $ComboXAxis.SelectedItem) {
         [System.Windows.Forms.MessageBox]::Show("Carregue dados e gere o gráfico antes de salvar.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    } else {
        Save-ChartImage -Chart $Chart
    }
})

# Desenhar a instrução inicial
Update-Chart `
    -Chart $Chart `
    -ComboY1Type $ComboY1Type `
    -ButtonY1Color $ButtonY1Color `
    -ComboY2Type $ComboY2Type `
    -ButtonY2Color $ButtonY2Color `
    -ComboXAxis $ComboXAxis `
    -ComboY1Data $ComboY1Data `
    -ComboY2Data $ComboY2Data

# Exibe a janela
$Form.ShowDialog() | Out-Null
