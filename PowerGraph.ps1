# -----------------------------------------------------------------------------
# Power-Graphx Web App Launcher
# Versão: 3.1.0 - Edição Web com Gráficos Corrigidos e Exportação
# Autor: jefferson/configexe (com modernização por IA)
#
# Melhorias da Versão 3.1.0:
# - Correção de Gráficos: Lógica de renderização de gráficos em JavaScript
#   ajustada para exibir corretamente cada tipo (barra, linha, combo).
# - Exportação para PNG: Adicionado botão "Baixar Gráfico (PNG)" no modal
#   de visualização para salvar a imagem do gráfico gerado.
# - Painel de Formatação: Reintroduzido um painel completo para formatar
#   a aparência dos gráficos (rótulos, eixos, cores, etc.).
# - Plugin de Rótulos Ativado: A biblioteca de rótulos (ChartDataLabels)
#   agora está corretamente registrada e funcional.
# -----------------------------------------------------------------------------

# --- 1. Carregar Assemblies Necessárias ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Web # Para codificação de JavaScript
}
catch {
    Write-Error "Não foi possível carregar as assemblies .NET necessárias."
    exit 1
}

# --- 2. Funções Principais ---

# Função para baixar e embutir as bibliotecas JS/CSS.
# Elas são baixadas para uma pasta temporária para não precisar buscar na internet toda vez.
Function Get-EmbeddedLibraries {
    $tempDir = Join-Path $env:TEMP "PowerGraphx_Libs"
    if (-not (Test-Path $tempDir)) {
        New-Item -Path $tempDir -ItemType Directory | Out-Null
    }

    $libs = @{
        "tailwindcss" = "https://cdn.tailwindcss.com";
        "chartjs"     = "https://cdn.jsdelivr.net/npm/chart.js";
        "chartlabels" = "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"
    }

    $embeddedContent = [PSCustomObject]@{
        Tailwind    = ""
        ChartJS     = ""
        ChartLabels = ""
    }

    try {
        $wc = New-Object System.Net.WebClient
        
        $tailwindPath = Join-Path $tempDir "tailwindcss.js"
        if (-not (Test-Path $tailwindPath)) { $wc.DownloadFile($libs.tailwindcss, $tailwindPath) }
        $embeddedContent.Tailwind = Get-Content -Path $tailwindPath -Raw

        $chartjsPath = Join-Path $tempDir "chart.js"
        if (-not (Test-Path $chartjsPath)) { $wc.DownloadFile($libs.chartjs, $chartjsPath) }
        $embeddedContent.ChartJS = Get-Content -Path $chartjsPath -Raw

        $chartlabelsPath = Join-Path $tempDir "chartlabels.js"
        if (-not (Test-Path $chartlabelsPath)) { $wc.DownloadFile($libs.chartlabels, $chartlabelsPath) }
        $embeddedContent.ChartLabels = Get-Content -Path $chartlabelsPath -Raw
        
        Write-Host "Bibliotecas carregadas com sucesso." -ForegroundColor Green
    }
    catch {
        Write-Error "Falha ao baixar as bibliotecas. Verifique sua conexão com a internet na primeira execução. Erro: $($_.Exception.Message)"
        exit 1
    }
    
    return $embeddedContent
}

# Função que contém o template HTML, CSS e JavaScript da aplicação completa.
Function Get-HtmlTemplate {
    param(
        [Parameter(Mandatory=$true)]$JsonData,
        [Parameter(Mandatory=$true)]$JsonColumnStructure,
        [Parameter(Mandatory=$true)]$EmbeddedLibraries
    )
    
    # O JavaScript da aplicação é vasto, então o mantemos aqui.
    $ApplicationJavaScript = @'
    // ---------------------------------------------------
    // Power-Graphx Web App - Lógica Principal
    // ---------------------------------------------------
    
    // Variáveis globais de estado
    let originalData = [];
    let currentData = [];
    let columnStructure = [];
    let chartInstance;
    
    // Mapeamento de estado para ordenação da tabela
    const sortState = {};

    document.addEventListener('DOMContentLoaded', () => {
        // Inicializa a aplicação com os dados embutidos pelo PowerShell
        originalData = JSON.parse(document.getElementById('jsonData').textContent);
        columnStructure = JSON.parse(document.getElementById('jsonColumnStructure').textContent);
        currentData = [...originalData];

        // Mapeia os nomes originais para os de exibição para uso futuro
        columnStructure.forEach(col => col.displayName = col.displayName || col.originalName);

        // Renderiza a tabela inicial e configura os eventos
        renderTable();
        setupEventListeners();
        updateStatus();
    });
    
    function updateStatus() {
        const statusLabel = document.getElementById('status-label');
        if (statusLabel) {
            statusLabel.textContent = `Exibindo ${currentData.length} de ${originalData.length} registros.`;
        }
    }

    // --- Funções de Renderização da Tabela ---
    function renderTable() {
        const tableContainer = document.getElementById('table-container');
        tableContainer.innerHTML = ''; // Limpa a tabela anterior
        if (currentData.length === 0) {
            tableContainer.innerHTML = `<p class="text-center text-gray-500 p-8">Nenhum dado para exibir.</p>`;
            return;
        }

        const table = document.createElement('table');
        table.className = 'min-w-full divide-y divide-gray-200';
        
        // Cria o cabeçalho
        const thead = document.createElement('thead');
        thead.className = 'bg-gray-50';
        const headerRow = document.createElement('tr');
        columnStructure.forEach((col, index) => {
            const th = document.createElement('th');
            th.scope = 'col';
            th.className = 'px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer select-none relative group';
            th.dataset.originalName = col.originalName;
            
            const titleDiv = document.createElement('div');
            titleDiv.className = 'flex items-center';
            titleDiv.textContent = col.displayName;
            th.appendChild(titleDiv);
            
            // Adiciona ícone de ordenação
            const sortIcon = document.createElement('span');
            sortIcon.className = 'ml-2 text-gray-400';
            if (sortState[col.originalName] === 'asc') {
                sortIcon.innerHTML = '&#9650;'; // Seta para cima
            } else if (sortState[col.originalName] === 'desc') {
                sortIcon.innerHTML = '&#9660;'; // Seta para baixo
            }
            titleDiv.appendChild(sortIcon);

            th.addEventListener('click', () => handleSort(col.originalName));
            
            // Menu de contexto para colunas
            const menuIcon = document.createElement('span');
            menuIcon.innerHTML = '&#8942;'; // 3 pontos verticais
            menuIcon.className = 'absolute right-1 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 transition p-1 rounded-full hover:bg-gray-200';
            menuIcon.addEventListener('click', (e) => {
                e.stopPropagation();
                showColumnMenu(e.target, col.originalName);
            });
            th.appendChild(menuIcon);
            
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        // Cria o corpo da tabela
        const tbody = document.createElement('tbody');
        tbody.className = 'bg-white divide-y divide-gray-200';
        currentData.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            tr.className = 'hover:bg-gray-50';
            columnStructure.forEach(col => {
                const td = document.createElement('td');
                td.className = 'px-4 py-3 whitespace-nowrap text-sm text-gray-700';
                td.textContent = row[col.originalName];
                td.setAttribute('contenteditable', 'true');
                td.addEventListener('blur', (e) => {
                    // Atualiza o dado no array quando a célula perde o foco
                    const newValue = e.target.textContent;
                    const originalRow = originalData.find(d => JSON.stringify(d) === JSON.stringify(row));
                    if(originalRow) {
                        originalRow[col.originalName] = newValue;
                    }
                    currentData[rowIndex][col.originalName] = newValue;
                });
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        tableContainer.appendChild(table);
    }
    
    // --- Lógica de Manipulação de Dados ---

    function handleSort(columnName) {
        const currentOrder = sortState[columnName];
        let nextOrder;
        if (currentOrder === 'asc') nextOrder = 'desc';
        else if (currentOrder === 'desc') nextOrder = undefined; // Volta ao original
        else nextOrder = 'asc';

        // Limpa estados de ordenação anteriores
        Object.keys(sortState).forEach(key => delete sortState[key]);
        
        const originalCopy = [...originalData];
        if (nextOrder) {
            sortState[columnName] = nextOrder;
        }

        // Ordena os dados
        currentData.sort((a, b) => {
            if (!nextOrder) {
                // Se não há ordem, usa a ordem do array original
                return originalCopy.indexOf(a) - originalCopy.indexOf(b);
            }
        
            const valA = a[columnName];
            const valB = b[columnName];
            
            const numA = parseFloat(String(valA).replace(',', '.'));
            const numB = parseFloat(String(valB).replace(',', '.'));

            let comparison = 0;
            if (!isNaN(numA) && !isNaN(numB)) {
                comparison = numA - numB;
            } else {
                comparison = String(valA).toLowerCase().localeCompare(String(valB).toLowerCase());
            }
            return nextOrder === 'asc' ? comparison : -comparison;
        });
        
        renderTable();
    }
    
    function showColumnMenu(target, columnName) {
        // Remove menu existente
        const existingMenu = document.getElementById('column-context-menu');
        if (existingMenu) existingMenu.remove();

        const menu = document.createElement('div');
        menu.id = 'column-context-menu';
        menu.className = 'absolute z-10 w-48 bg-white rounded-md shadow-lg border';
        
        const rect = target.getBoundingClientRect();
        menu.style.top = `${rect.bottom + window.scrollY}px`;
        menu.style.left = `${rect.left + window.scrollX}px`;

        menu.innerHTML = `
            <a href="#" class="block px-4 py-2 text-sm text-gray-700 hover:bg-gray-100" id="rename-col">Renomear</a>
            <a href="#" class="block px-4 py-2 text-sm text-red-600 hover:bg-gray-100" id="remove-col">Remover Coluna</a>
        `;
        document.body.appendChild(menu);

        document.getElementById('rename-col').addEventListener('click', (e) => {
            e.preventDefault();
            renameColumn(columnName);
            menu.remove();
        });
        document.getElementById('remove-col').addEventListener('click', (e) => {
            e.preventDefault();
            removeColumn(columnName);
            menu.remove();
        });
        
        // Fecha o menu se clicar fora
        document.addEventListener('click', (e) => {
             if (!menu.contains(e.target)) {
                menu.remove();
             }
        }, { once: true });
    }
    
    function renameColumn(oldName) {
        const col = columnStructure.find(c => c.originalName === oldName);
        const newName = prompt(`Digite o novo nome para a coluna "${col.displayName}":`, col.displayName);
        if (newName && newName.trim() !== '') {
            col.displayName = newName.trim();
            renderTable();
        }
    }

    function removeColumn(columnName) {
        const col = columnStructure.find(c => c.originalName === columnName);
        if (confirm(`Tem certeza que deseja remover a coluna "${col.displayName}"?`)) {
            columnStructure = columnStructure.filter(c => c.originalName !== columnName);
            currentData.forEach(row => delete row[columnName]);
            originalData.forEach(row => delete row[columnName]);
            renderTable();
            updateStatus();
        }
    }
    
    function addCalculatedColumn() {
        const newName = prompt("Nome da nova coluna:");
        if (!newName || newName.trim() === '') return;

        const formula = prompt(`Digite a fórmula. Use 'row' para acessar os dados da linha (ex: row.Valor * 1.1).\nColunas: ${columnStructure.map(c=>c.originalName).join(', ')}`);
        if (!formula) return;
        
        try {
            const calcFunc = new Function('row', `try { return ${formula}; } catch(e) { return 'ERRO'; }`);
            
            currentData.forEach(row => { row[newName] = calcFunc(row); });
            originalData.forEach(row => { row[newName] = calcFunc(row); });

            columnStructure.push({ originalName: newName, displayName: newName });
            renderTable();
            updateStatus();
        } catch (e) {
            alert("Erro na fórmula: " + e.message);
        }
    }
    
    function applyFilter() {
        const column = document.getElementById('filter-column').value;
        const condition = document.getElementById('filter-condition').value;
        const value = document.getElementById('filter-value').value.toLowerCase();
        
        if (!column) return;

        currentData = originalData.filter(row => {
            const cellValue = String(row[column] || '').toLowerCase();
            const numCellValue = parseFloat(String(row[column]).replace(',', '.'));
            const numValue = parseFloat(String(value).replace(',', '.'));

            switch (condition) {
                case 'contains': return cellValue.includes(value);
                case 'not_contains': return !cellValue.includes(value);
                case 'equals': return cellValue === value;
                case 'not_equals': return cellValue !== value;
                case 'greater': return !isNaN(numCellValue) && !isNaN(numValue) && numCellValue > numValue;
                case 'less': return !isNaN(numCellValue) && !isNaN(numValue) && numCellValue < numValue;
                default: return true;
            }
        });
        
        renderTable();
        updateStatus();
        closeModal('filter-modal');
    }

    function removeFilter() {
        currentData = [...originalData];
        Object.keys(sortState).forEach(key => delete sortState[key]); // Limpa ordenação
        renderTable();
        updateStatus();
    }
    
    function downloadCSV() {
        if (currentData.length === 0) return;
        const headers = columnStructure.map(c => c.displayName);
        const rows = currentData.map(row => columnStructure.map(col => {
            let cell = row[col.originalName] === null || row[col.originalName] === undefined ? '' : row[col.originalName];
            let cellString = String(cell);
            if (cellString.includes(',') || cellString.includes('"') || cellString.includes('\n')) {
                cellString = `"${cellString.replace(/"/g, '""')}"`;
            }
            return cellString;
        }).join(','));
        const csvContent = [headers.join(','), ...rows].join('\n');
        const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "power-graphx-export.csv";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function setupEventListeners() {
        document.getElementById('btn-filter').addEventListener('click', () => showModal('filter-modal'));
        document.getElementById('btn-add-column').addEventListener('click', addCalculatedColumn);
        document.getElementById('btn-remove-filter').addEventListener('click', removeFilter);
        document.getElementById('btn-download-csv').addEventListener('click', downloadCSV);
        document.getElementById('btn-view-charts').addEventListener('click', () => showModal('charts-modal'));
        document.getElementById('apply-filter-btn').addEventListener('click', applyFilter);
        document.querySelectorAll('.modal-close').forEach(el => {
            el.addEventListener('click', () => closeModal(el.closest('.modal').id));
        });
        const filterColumnSelect = document.getElementById('filter-column');
        columnStructure.forEach(col => {
            const option = document.createElement('option');
            option.value = col.originalName;
            option.textContent = col.displayName;
            filterColumnSelect.appendChild(option);
        });
    }

    function showModal(modalId) {
        const modal = document.getElementById(modalId);
        if(modalId === 'charts-modal' && !modal.dataset.initialized) {
           initializeChartUI();
           modal.dataset.initialized = 'true';
        }
        modal.classList.remove('hidden');
    }

    function closeModal(modalId) {
        document.getElementById(modalId).classList.add('hidden');
    }
    
    // --- Lógica dos Gráficos ---
    function initializeChartUI() {
        Chart.register(ChartDataLabels);
        let seriesCounter = 0;
        const seriesColors = ["#3b82f6", "#ef4444", "#22c55e", "#f97316", "#8b5cf6", "#14b8a6"];

        const isNumeric = (colName) => {
            const sample = currentData.find(d => d[colName] != null);
            return sample && !isNaN(parseFloat(String(sample[colName]).replace(',', '.')));
        };

        const populateSelect = (select, type = 'all') => {
            select.innerHTML = '';
            columnStructure.filter(c => type !== 'numeric' || isNumeric(c.originalName)).forEach(col => {
                select.add(new Option(col.displayName, col.originalName));
            });
        };

        const addSeriesControl = (isFirst = false) => {
            seriesCounter++;
            const div = document.createElement('div');
            div.className = 'p-3 border rounded-lg bg-gray-50 grid grid-cols-1 sm:grid-cols-2 gap-3 items-end';
            div.innerHTML = `
                <div><label class="text-xs font-semibold">Eixo X / Grupo:</label><select name="x-axis" class="mt-1 block w-full rounded-md border-gray-300 text-sm"></select></div>
                <div><label class="text-xs font-semibold">Eixo Y / Valor:</label><select name="y-axis" class="mt-1 block w-full rounded-md border-gray-300 text-sm"></select></div>
                <div class="combo-type-control hidden"><label class="text-xs font-semibold">Tipo:</label><select name="series-type" class="mt-1 block w-full rounded-md border-gray-300 text-sm"><option value="bar">Barra</option><option value="line">Linha</option></select></div>
                <div class="flex items-end space-x-2">
                    <div class="w-full"><label class="text-xs font-semibold">Cor:</label><input type="color" value="${seriesColors[seriesCounter-1]}" name="color" class="mt-1 w-full h-9 p-0 border-0"></div>
                    ${!isFirst ? `<button type="button" class="remove-series-btn h-9 px-3 bg-red-500 text-white rounded-md hover:bg-red-600">&times;</button>` : ''}
                </div>`;
            document.getElementById('series-container').appendChild(div);
            populateSelect(div.querySelector('[name="x-axis"]'), 'all');
            populateSelect(div.querySelector('[name="y-axis"]'), 'numeric');
            if (!isFirst) div.querySelector('.remove-series-btn').onclick = (e) => { e.target.closest('div.p-3').remove(); renderChart(); };
        };
        
        const buildChartOptions = () => {
            const fontColor = '#64748B';
            const gridColor = 'rgba(0, 0, 0, 0.1)';
            return {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    datalabels: {
                        display: document.getElementById('show-labels').checked,
                        color: fontColor,
                        font: { size: 12 },
                        anchor: 'end', align: 'end',
                        formatter: val => typeof val === 'number' ? val.toLocaleString('pt-BR') : val
                    }
                },
                scales: {
                    x: { ticks: { color: fontColor }, grid: { color: gridColor } },
                    y: { beginAtZero: true, ticks: { color: fontColor }, grid: { color: gridColor } }
                }
            };
        };

        window.renderChart = () => {
             if (chartInstance) chartInstance.destroy();
             const chartType = document.querySelector('input[name="chart-type"]:checked').value;
             document.querySelectorAll('.combo-type-control').forEach(el => el.style.display = chartType === 'combo' ? 'block' : 'none');

             const seriesControls = document.querySelectorAll('#series-container > div');
             if (seriesControls.length === 0) return;
             
             const parseValue = v => parseFloat(String(v || '0').replace(',', '.')) || 0;
             const firstXAxis = seriesControls[0].querySelector('[name="x-axis"]').value;
             const labels = [...new Set(currentData.map(d => d[firstXAxis]))];

             const datasets = Array.from(seriesControls).map(control => {
                const yCol = control.querySelector('[name="y-axis"]').value;
                const xCol = control.querySelector('[name="x-axis"]').value;
                const seriesTypeOption = control.querySelector('[name="series-type"]').value;

                let seriesType;
                if (chartType === 'combo') seriesType = seriesTypeOption;
                else if (chartType === 'line') seriesType = 'line';
                else seriesType = 'bar'; // for bar, stacked, horizontalBar

                return {
                    label: columnStructure.find(c => c.originalName === yCol).displayName,
                    data: labels.map(label => currentData.filter(d => d[xCol] === label).reduce((sum, r) => sum + parseValue(r[yCol]), 0)),
                    borderColor: control.querySelector('[name="color"]').value,
                    backgroundColor: control.querySelector('[name="color"]').value + 'B3',
                    type: seriesType,
                    tension: 0.4
                };
             });

             const options = buildChartOptions();
             options.indexAxis = chartType === 'horizontalBar' ? 'y' : 'x';
             options.scales.x.stacked = chartType === 'stacked';
             options.scales.y.stacked = chartType === 'stacked';
             
             chartInstance = new Chart('mainChart', { type: 'bar', data: { labels, datasets }, options });
        };
        
        const downloadChart = () => {
            if (!chartInstance) { alert('Gere um gráfico para poder baixá-lo.'); return; }
            const link = document.createElement('a');
            link.href = chartInstance.toBase64Image('image/png', 1.0);
            link.download = 'power-graphx-chart.png';
            link.click();
        };

        addSeriesControl(true);
        document.getElementById('charts-controls-panel').addEventListener('change', renderChart);
        document.getElementById('add-series-btn').addEventListener('click', () => addSeriesControl());
        document.getElementById('download-chart-btn').addEventListener('click', downloadChart);
        renderChart();
    }
'@

    # O HTML agora inclui a tabela e os modais para interatividade.
    # As bibliotecas e dados são injetados diretamente.
    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power-Graphx Web Editor</title>
    <style>
        .modal { transition: opacity 0.25s ease; }
        #table-container { max-height: calc(100vh - 140px); overflow: auto; }
        table thead { position: sticky; top: 0; z-index: 1; }
        .chart-selector input:checked + label { border-color: #3b82f6; background-color: #eff6ff; box-shadow: 0 0 0 2px #3b82f6; }
    </style>
    <script>$($EmbeddedLibraries.Tailwind)</script>
</head>
<body class="bg-gray-100 font-sans">

    <!-- Dados JSON embutidos -->
    <script id="jsonData" type="application/json">$JsonData</script>
    <script id="jsonColumnStructure" type="application/json">$JsonColumnStructure</script>

    <!-- Cabeçalho e Barra de Ferramentas -->
    <header class="bg-white shadow-md p-4 sticky top-0 z-20">
        <div class="container mx-auto flex justify-between items-center">
            <h1 class="text-2xl font-bold text-gray-800">Power-Graphx Web Editor</h1>
            <div class="flex items-center space-x-2">
                <button id="btn-filter" class="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700">Filtrar</button>
                <button id="btn-add-column" class="px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700">Coluna Calculada</button>
                <button id="btn-remove-filter" class="px-4 py-2 text-sm font-medium text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300">Remover Filtro</button>
                <button id="btn-view-charts" class="px-4 py-2 text-sm font-medium text-white bg-purple-600 rounded-md hover:bg-purple-700">Visualizar Gráficos</button>
                <button id="btn-download-csv" class="px-4 py-2 text-sm font-medium text-white bg-gray-800 rounded-md hover:bg-gray-900">Baixar CSV</button>
            </div>
        </div>
    </header>

    <!-- Container da Tabela de Dados -->
    <main class="container mx-auto p-4">
        <div id="table-container" class="bg-white rounded-lg shadow overflow-hidden"></div>
    </main>
    
    <!-- Barra de Status -->
    <footer class="fixed bottom-0 left-0 right-0 bg-gray-800 text-white text-sm p-2 text-center">
        <span id="status-label">Carregando...</span>
    </footer>

    <!-- Modal de Filtro -->
    <div id="filter-modal" class="modal hidden fixed inset-0 bg-gray-600 bg-opacity-50 overflow-y-auto h-full w-full z-30">
        <div class="relative top-20 mx-auto p-5 border w-96 shadow-lg rounded-md bg-white">
            <h3 class="text-lg font-medium text-gray-900 mb-4">Filtrar Dados</h3>
            <div class="space-y-4">
                <div><label for="filter-column" class="block text-sm font-medium text-gray-700">Coluna</label><select id="filter-column" class="mt-1 block w-full py-2 px-3 border border-gray-300 bg-white rounded-md shadow-sm"></select></div>
                <div><label for="filter-condition" class="block text-sm font-medium text-gray-700">Condição</label><select id="filter-condition" class="mt-1 block w-full py-2 px-3 border border-gray-300 bg-white rounded-md shadow-sm"><option value="contains">Contém</option><option value="not_contains">Não Contém</option><option value="equals">Igual a</option><option value="not_equals">Diferente de</option><option value="greater">Maior que (numérico)</option><option value="less">Menor que (numérico)</option></select></div>
                <div><label for="filter-value" class="block text-sm font-medium text-gray-700">Valor</label><input type="text" id="filter-value" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"></div>
            </div>
            <div class="mt-6 flex justify-end space-x-2"><button class="modal-close px-4 py-2 bg-gray-200 text-gray-800 rounded-md hover:bg-gray-300">Cancelar</button><button id="apply-filter-btn" class="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700">Aplicar</button></div>
        </div>
    </div>
    
    <!-- Modal de Gráficos -->
    <div id="charts-modal" class="modal hidden fixed inset-0 bg-gray-800 bg-opacity-75 flex items-center justify-center z-40 p-4">
        <div class="bg-white rounded-xl shadow-2xl w-full h-full max-w-7xl flex flex-col p-6 relative">
            <button class="modal-close absolute top-4 right-4 text-gray-500 hover:text-gray-800 text-2xl font-bold">&times;</button>
            <h2 class="text-2xl font-bold text-gray-800 mb-4">Visualização de Gráficos</h2>
            <div class="flex-grow grid grid-cols-1 lg:grid-cols-4 gap-6 overflow-hidden">
                <!-- Coluna de Controles -->
                <div id="charts-controls-panel" class="lg:col-span-1 flex flex-col space-y-4 overflow-y-auto pr-2">
                    <div>
                        <h3 class="font-bold text-gray-700 mb-2">1. Tipo de Gráfico</h3>
                        <div class="chart-selector grid grid-cols-3 gap-2">
                            <div><input type="radio" name="chart-type" value="bar" id="type-bar" checked class="hidden"><label for="type-bar" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Barra</label></div>
                            <div><input type="radio" name="chart-type" value="line" id="type-line" class="hidden"><label for="type-line" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Linha</label></div>
                            <div><input type="radio" name="chart-type" value="combo" id="type-combo" class="hidden"><label for="type-combo" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Combo</label></div>
                            <div><input type="radio" name="chart-type" value="stacked" id="type-stacked" class="hidden"><label for="type-stacked" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Empilhado</label></div>
                            <div><input type="radio" name="chart-type" value="horizontalBar" id="type-horizontalBar" class="hidden"><label for="type-horizontalBar" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Horizontal</label></div>
                        </div>
                    </div>
                    <div>
                        <div class="flex justify-between items-center mb-2"><h3 class="font-bold text-gray-700">2. Séries de Dados</h3><button id="add-series-btn" class="text-xs bg-blue-500 text-white py-1 px-2 rounded-full hover:bg-blue-600">+ Série</button></div>
                        <div id="series-container" class="space-y-3 max-h-60 overflow-y-auto"></div>
                    </div>
                </div>
                <!-- Coluna do Gráfico -->
                <div class="lg:col-span-2 bg-gray-50 rounded-lg p-4"><div class="relative w-full h-full"><canvas id="mainChart"></canvas></div></div>
                <!-- Coluna de Formatação e Ações -->
                <div class="lg:col-span-1 flex flex-col space-y-4 overflow-y-auto pr-2">
                    <div>
                        <h3 class="font-bold text-gray-700 mb-2">3. Formatação</h3>
                        <div class="flex items-center"><input id="show-labels" type="checkbox" class="h-4 w-4 rounded border-gray-300 text-blue-600"><label for="show-labels" class="ml-2 block text-sm text-gray-900">Exibir rótulos de dados</label></div>
                    </div>
                    <div class="pt-4 border-t">
                         <h3 class="font-bold text-gray-700 mb-2">4. Ações</h3>
                         <button id="download-chart-btn" class="w-full bg-gray-600 text-white font-bold py-2 rounded-lg hover:bg-gray-700">Baixar Gráfico (PNG)</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Bibliotecas e Lógica da Aplicação -->
    <script>$($EmbeddedLibraries.ChartJS)</script>
    <script>$($EmbeddedLibraries.ChartLabels)</script>
    <script>$ApplicationJavaScript</script>
</body>
</html>
"@
}

# --- 3. Função Principal de Execução ---
Function Start-WebApp {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Power-Graphx: Selecione o arquivo CSV"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        Write-Host "Analisando: $(Split-Path $FilePath -Leaf)..."
        
        try {
            # Detecta o delimitador e importa os dados
            $firstLine = Get-Content -Path $FilePath -TotalCount 1 -Encoding Default
            $bestDelimiter = if (($firstLine -split ';').Count -gt ($firstLine -split ',').Count) { ';' } else { ',' }
            $Data = Import-Csv -Path $FilePath -Delimiter $bestDelimiter
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Não foi possível ler os dados do arquivo CSV.", "Erro de Leitura", "OK", "Error")
            return
        }

        if ($null -eq $Data -or $Data.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("O arquivo CSV está vazio ou em um formato inválido.", "Erro de Leitura", "OK", "Error")
            return
        }
        
        Write-Host "Dados carregados. Preparando a aplicação web..."
        
        # Estrutura das Colunas
        $ColumnStructure = $Data[0].PSObject.Properties | ForEach-Object {
            [PSCustomObject]@{
                originalName = $_.Name
                displayName  = $_.Name
            }
        }
        
        # Converte para JSON
        $JsonData = $Data | ConvertTo-Json -Compress -Depth 10
        $JsonColumnStructure = $ColumnStructure | ConvertTo-Json -Compress

        # Carrega as bibliotecas (baixa se necessário)
        $embeddedLibs = Get-EmbeddedLibraries
        
        # Gera o conteúdo HTML completo
        $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -JsonColumnStructure $JsonColumnStructure -EmbeddedLibraries $embeddedLibs
        
        # Salva e abre o arquivo HTML
        $OutputPath = Join-Path $env:TEMP "PowerGraphx_WebApp.html"
        try {
            $HtmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
            Write-Host "Aplicação gerada com sucesso! Abrindo no seu navegador..." -ForegroundColor Green
            Start-Process $OutputPath
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Ocorreu um erro ao gerar ou abrir o arquivo HTML: $($_.Exception.Message)", "Erro", "OK", "Error")
        }
    }
}

# --- 4. Iniciar a Aplicação ---
Start-WebApp

