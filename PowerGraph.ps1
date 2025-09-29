# -----------------------------------------------------------------------------
# Power-Graphx Web App Launcher
# Versão: 4.1.2 - Correção Crítica na Criação de Gráficos
# Autor: jefferson/configexe (com modernização por IA)
#
# Melhorias da Versão 4.1.2:
# - CORREÇÃO DE BUG: Resolvido o problema em que os seletores de Eixo X e
#   Eixo Y não eram populados ao adicionar um novo painel de gráfico.
# - ESTABILIDADE: Corrigido um erro relacionado aos seletores de tipo de
#   gráfico que poderia ocorrer ao usar múltiplos painéis.
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

# Função para fornecer URLs de CDN.
Function Get-CdnLibraryTags {
    $libs = @{
        "tailwindcss" = "<script src=`"https://cdn.tailwindcss.com`"></script>";
        "chartjs"     = "<script src=`"https://cdn.jsdelivr.net/npm/chart.js`"></script>";
        "chartlabels" = "<script src=`"https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0`"></script>";
        "alasql"      = "<script src=`"https://cdn.jsdelivr.net/npm/alasql@4`"></script>"
    }
    return $libs.Values -join "`n    "
}


# Função que contém o template HTML, CSS e JavaScript da aplicação completa.
Function Get-HtmlTemplate {
    param(
        [Parameter(Mandatory=$true)]$JsonData,
        [Parameter(Mandatory=$true)]$JsonColumnStructure,
        [Parameter(Mandatory=$true)]$CdnLibraryTags
    )
    
    $ApplicationJavaScript = @'
    // ---------------------------------------------------
    // Power-Graphx Web App - Lógica Principal (v4.1.2 com AlaSQL)
    // ---------------------------------------------------
    
    // Variáveis globais
    let originalData = [];
    let currentData = [];
    let columnStructure = [];
    let chartInstances = {};
    let chartAnalysisCounter = 0;
    const initialTableName = 'source_data';
    const sortState = {};

    // Ponto de entrada mais seguro
    window.onload = () => {
        try {
            initializeApp();
        } catch (e) {
            console.error("Falha crítica na inicialização:", e);
            const statusLabel = document.getElementById('status-label');
            if (statusLabel) statusLabel.textContent = `Erro ao carregar a aplicação. Pressione F12 para ver detalhes.`;
            alert(`Ocorreu um erro ao carregar a aplicação. Pressione F12 para ver o console de erros.`);
        }
    };

    function initializeApp() {
        originalData = JSON.parse(document.getElementById('jsonData').textContent);
        let initialColumnStructure = JSON.parse(document.getElementById('jsonColumnStructure').textContent);
        currentData = JSON.parse(JSON.stringify(originalData)); 
        
        updateColumnStructure(initialColumnStructure);
        setupEventListeners();
        
        initializeDB();
        renderTable();
        updateStatus();
    }
    
    function initializeDB() {
        alasql.tables[initialTableName] = { data: originalData };
        updateTableListUI();
        document.getElementById('sql-status').textContent = 'Motor SQL (AlaSQL) pronto.';
    }
    
    function updateTableListUI() {
        const tables = Object.keys(alasql.tables);
        const listEl = document.getElementById('table-list');
        listEl.innerHTML = '';
        if (tables.length > 0) {
            tables.forEach(tableName => {
                const li = document.createElement('li');
                li.textContent = tableName;
                listEl.appendChild(li);
            });
        } else {
            listEl.innerHTML = '<li>Nenhuma tabela carregada.</li>';
        }

        document.querySelectorAll('.chart-data-source').forEach(select => {
            const currentVal = select.value;
            select.innerHTML = '';
            tables.forEach(t => select.add(new Option(t, t)));
            select.value = tables.includes(currentVal) ? currentVal : tables[0];
            select.dispatchEvent(new Event('change'));
        });
    }

    function handleFileUploads(event) {
        const files = event.target.files;
        if (!files.length) return;

        const statusEl = document.getElementById('sql-status');
        
        Array.from(files).forEach(file => {
            let tableName = prompt(`Digite o nome para a tabela do arquivo "${file.name}":`, file.name.split('.')[0].replace(/[^a-zA-Z0-9_]/g, '_'));
            if (!tableName) return;
            tableName = tableName.replace(/[^a-zA-Z0-9_]/g, '_');

            statusEl.textContent = `Lendo arquivo ${file.name}...`;
            const reader = new FileReader();
            reader.onload = function(e) {
                const fileContent = e.target.result;
                try {
                    statusEl.textContent = `Processando "${tableName}" com AlaSQL...`;
                    let data = alasql('SELECT * FROM CSV(?, {headers:true, separator:";"})', [fileContent]);
                    if (data.length > 0 && Object.keys(data[0]).length <= 1) {
                        const dataComma = alasql('SELECT * FROM CSV(?, {headers:true, separator:","})', [fileContent]);
                        if (dataComma.length > 0 && Object.keys(dataComma[0]).length > 1) {
                            data = dataComma;
                        }
                    }
                    alasql.tables[tableName] = { data: data };
                    
                    statusEl.textContent = `Tabela "${tableName}" criada com sucesso.`;
                    updateTableListUI();
                } catch(err) {
                    statusEl.textContent = `Erro ao carregar a tabela "${tableName}": ${err.message}`;
                    console.error(err);
                }
            };
            reader.readAsText(file);
        });
        
        event.target.value = '';
    }

    function runQueryAndUpdateUI() {
        const query = document.getElementById('sql-editor').value;
        if (!query.trim()) return;

        const statusEl = document.getElementById('sql-status');
        try {
            statusEl.textContent = 'Executando consulta...';
            const resultData = alasql(query);

            if (resultData.length > 0) {
                const newColumns = Object.keys(resultData[0]).map(name => ({
                    originalName: name,
                    displayName: name
                }));
                updateColumnStructure(newColumns);
                currentData = resultData;
                statusEl.textContent = `Consulta executada com sucesso. ${resultData.length} linhas retornadas.`;
            } else {
                currentData = [];
                updateColumnStructure([]);
                statusEl.textContent = 'Consulta executada, mas não retornou linhas.';
            }
            
            Object.keys(sortState).forEach(key => delete sortState[key]);
            renderTable();
            updateStatus();

            document.querySelectorAll('.chart-analysis-section').forEach(section => {
                const id = section.dataset.id;
                renderChart(id);
            });
        } catch (e) {
            statusEl.textContent = `Erro na consulta SQL: ${e.message}`;
            console.error(e);
        }
    }

    function addChartAnalysis() {
        chartAnalysisCounter++;
        const template = document.getElementById('chart-analysis-template').innerHTML;
        const newChartHtml = template.replace(/__ID__/g, chartAnalysisCounter);
        
        const container = document.getElementById('charts-container');
        const div = document.createElement('div');
        div.innerHTML = newChartHtml;
        container.appendChild(div.firstElementChild);

        initializeChartUI(chartAnalysisCounter);
    }

    function removeChartAnalysis(id) {
        const section = document.getElementById(`chart-section-${id}`);
        if (section) section.remove();
        if (chartInstances[id]) {
            chartInstances[id].destroy();
            delete chartInstances[id];
        }
    }

    function showUnpivotHelper() {
        const tables = Object.keys(alasql.tables);
        if (tables.length === 0) { alert("Carregue uma tabela primeiro."); return; }
        const tableName = prompt("Qual tabela você quer transformar (unpivot)?\n\Tabelas disponíveis: " + tables.join(', '), tables[0]);
        if (!tableName || !alasql.tables[tableName]) { alert("Tabela inválida."); return; }
        
        const tableData = alasql.tables[tableName].data;
        if(tableData.length === 0) { alert("A tabela selecionada está vazia."); return; }

        const columns = Object.keys(tableData[0]);
        const idColumnsStr = prompt("Digite as colunas que devem ser MANTIDAS (separadas por vírgula).\n\nEx: ID, Produto\n\nColunas disponíveis: " + columns.join(', '));
        if (idColumnsStr === null) return;
        
        const idColumns = idColumnsStr.split(',').map(c => c.trim()).filter(c => columns.includes(c) && c);
        const pivotColumns = columns.filter(c => !idColumns.includes(c));
        const newTableName = prompt("Qual o nome da nova tabela a ser criada?", `${tableName}_unpivoted`);
        if (!newTableName) return;
        const categoryColName = prompt("Qual o nome da nova coluna para as categorias (ex: Mes)?", "Categoria");
        if (!categoryColName) return;
        const valueColName = prompt("Qual o nome da nova coluna para os valores (ex: Vendas)?", "Valor");
        if (!valueColName) return;
        
        const idColumnsQuoted = idColumns.map(c => `\`${c}\``);
        const idColumnSelector = idColumns.length > 0 ? idColumnsQuoted.join(', ') + ',' : '';

        const selectClauses = pivotColumns.map(pCol => 
            `SELECT ${idColumnSelector} '${pCol}' AS \`${categoryColName}\`, \`${pCol}\` AS \`${valueColName}\` FROM \`${tableName}\``
        );

        const fullQuery = `SELECT * INTO \`${newTableName}\` FROM (${selectClauses.join(' UNION ALL ')})`;
        
        document.getElementById('sql-editor').value = fullQuery;
        runQueryAndUpdateUI();
    }

    function setupEventListeners() {
        document.getElementById('btn-filter').addEventListener('click', () => document.getElementById('filter-modal').classList.remove('hidden'));
        document.getElementById('btn-download-csv').addEventListener('click', downloadCSV);
        document.getElementById('btn-toggle-sql').addEventListener('click', () => document.getElementById('sql-section').classList.toggle('hidden'));
        document.getElementById('btn-run-sql').addEventListener('click', runQueryAndUpdateUI);
        document.getElementById('btn-reset-data').addEventListener('click', () => {
            Object.keys(alasql.tables).forEach(key => {
                if(key !== initialTableName) delete alasql.tables[key];
            });
            updateTableListUI();
            document.getElementById('sql-editor').value = `SELECT * FROM ${initialTableName};`;
            runQueryAndUpdateUI();
        });
        document.getElementById('apply-filter-btn').addEventListener('click', () => applyFilter(true));
        document.querySelectorAll('.modal-close').forEach(el => el.addEventListener('click', () => el.closest('.modal').classList.add('hidden')));
        document.getElementById('btn-add-csv').addEventListener('click', () => document.getElementById('csv-upload-input').click());
        document.getElementById('csv-upload-input').addEventListener('change', handleFileUploads);
        document.getElementById('btn-add-chart').addEventListener('click', addChartAnalysis);
        document.getElementById('btn-sql-unpivot').addEventListener('click', showUnpivotHelper);
    }

    function updateColumnStructure(newStructure) {
        columnStructure = newStructure.map(col => ({
             ...col,
             displayName: col.displayName || col.originalName
        }));
        const filterColumnSelect = document.getElementById('filter-column');
        filterColumnSelect.innerHTML = '<option value="">-- Selecione --</option>';
        newStructure.forEach(col => filterColumnSelect.add(new Option(col.displayName, col.originalName)));
    }

    function updateStatus() {
        const total = (originalData && originalData.length) ? originalData.length : 0;
        document.getElementById('status-label').textContent = `Exibindo ${currentData.length} registros. (Original: ${total})`;
    }

    function renderTable() {
        const tableContainer = document.getElementById('table-container');
        tableContainer.innerHTML = ''; 
        if (currentData.length === 0) {
            tableContainer.innerHTML = `<p class="text-center text-gray-500 p-8">Nenhum dado para exibir.</p>`;
            return;
        }

        const table = document.createElement('table');
        table.className = 'min-w-full divide-y divide-gray-200';
        
        const thead = document.createElement('thead');
        thead.className = 'bg-gray-50';
        const headerRow = document.createElement('tr');
        columnStructure.forEach((col) => {
            const th = document.createElement('th');
            th.className = 'px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer select-none relative group';
            th.dataset.originalName = col.originalName;
            
            const titleDiv = document.createElement('div');
            titleDiv.className = 'flex items-center';
            titleDiv.textContent = col.displayName;
            
            const sortIcon = document.createElement('span');
            sortIcon.className = 'ml-2 text-gray-400';
            if (sortState[col.originalName] === 'asc') { sortIcon.innerHTML = '&#9650;'; } 
            else if (sortState[col.originalName] === 'desc') { sortIcon.innerHTML = '&#9660;'; }
            titleDiv.appendChild(sortIcon);
            th.appendChild(titleDiv);
            th.addEventListener('click', () => handleSort(col.originalName));
            
            const menuIcon = document.createElement('span');
            menuIcon.innerHTML = '&#8942;';
            menuIcon.className = 'absolute right-1 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 transition p-1 rounded-full hover:bg-gray-200';
            menuIcon.addEventListener('click', (e) => { e.stopPropagation(); showColumnMenu(e.target, col.originalName); });
            th.appendChild(menuIcon);
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        tbody.className = 'bg-white divide-y divide-gray-200';
        currentData.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            tr.className = 'hover:bg-gray-50';
            tr.dataset.rowId = rowIndex;
            
            columnStructure.forEach(col => {
                const td = document.createElement('td');
                td.className = 'px-4 py-3 whitespace-nowrap text-sm text-gray-700';
                td.textContent = row[col.originalName];
                td.setAttribute('contenteditable', 'false');
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        tableContainer.appendChild(table);
    }
    
    function handleSort(columnName) {
        const currentOrder = sortState[columnName];
        let nextOrder;
        if (currentOrder === 'asc') nextOrder = 'desc';
        else if (currentOrder === 'desc') nextOrder = undefined;
        else nextOrder = 'asc';

        Object.keys(sortState).forEach(key => delete sortState[key]);
        
        if (nextOrder) {
            sortState[columnName] = nextOrder;
            currentData.sort((a, b) => {
                const valA = a[columnName];
                const valB = b[columnName];
                const numA = (typeof valA === 'number') ? valA : parseFloat(String(valA).replace(',', '.'));
                const numB = (typeof valB === 'number') ? valB : parseFloat(String(valB).replace(',', '.'));

                let comparison = 0;
                if (!isNaN(numA) && !isNaN(numB)) {
                    comparison = numA - numB;
                } else {
                    comparison = String(valA || '').toLowerCase().localeCompare(String(valB || '').toLowerCase());
                }
                return nextOrder === 'asc' ? comparison : -comparison;
            });
        }
        renderTable();
    }
    
    function showColumnMenu(target, columnName) {
        const existingMenu = document.getElementById('column-context-menu');
        if (existingMenu) existingMenu.remove();
        const menu = document.createElement('div');
        menu.id = 'column-context-menu';
        menu.className = 'absolute z-50 w-48 bg-white rounded-md shadow-lg border';
        const rect = target.getBoundingClientRect();
        menu.style.top = `${rect.bottom + window.scrollY}px`;
        menu.style.left = `${rect.left + window.scrollX}px`;
        menu.innerHTML = `<a href="#" class="block px-4 py-2 text-sm text-gray-700 hover:bg-gray-100" id="rename-col">Renomear</a>`;
        document.body.appendChild(menu);
        document.getElementById('rename-col').onclick = (e) => { e.preventDefault(); renameColumn(columnName); menu.remove(); };
        document.addEventListener('click', (e) => { if (!menu.contains(e.target)) menu.remove(); }, { once: true });
    }
    
    function renameColumn(oldName) {
        const col = columnStructure.find(c => c.originalName === oldName);
        const newName = prompt(`Digite o novo nome para a coluna "${col.displayName}":`, col.displayName);
        if (newName && newName.trim()) { col.displayName = newName.trim(); renderTable(); }
    }

    function applyFilter(close = true) {
        const column = document.getElementById('filter-column').value;
        const condition = document.getElementById('filter-condition').value;
        const value = document.getElementById('filter-value').value.toLowerCase();
        
        const filteredData = currentData.filter(row => {
            if (!column) return true;
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
        
        const originalVisibility = currentData;
        currentData = filteredData;
        renderTable();
        currentData = originalVisibility; 
        
        if(close) document.getElementById('filter-modal').classList.add('hidden');
    }

    function downloadCSV() {
        if (currentData.length === 0) return;
        const headers = columnStructure.map(c => c.displayName);
        const rows = currentData.map(row => columnStructure.map(col => {
            let cell = row[col.originalName] ?? '';
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

    function initializeChartUI(id) {
        const section = document.getElementById(`chart-section-${id}`);
        if (!section) return;

        section.addEventListener('change', () => renderChart(id));
        section.addEventListener('input', () => renderChart(id));
        section.querySelector('.add-series-btn').addEventListener('click', () => addSeriesControl(id));
        section.querySelector('.download-chart-btn').addEventListener('click', () => downloadChart(id));
        section.querySelector('.remove-chart-btn').addEventListener('click', () => removeChartAnalysis(id));
        section.querySelector('.y-axis-auto').onchange = (e) => {
            section.querySelector('.y-axis-max').disabled = e.target.checked;
        };

        section.querySelector('.chart-data-source').addEventListener('change', (e) => {
            const selectedTable = e.target.value;
            const tableData = alasql.tables[selectedTable]?.data;
            let cols = [];
            if(tableData && tableData.length > 0) {
                cols = Object.keys(tableData[0]).map(name => ({ originalName: name, displayName: name }));
            }
            
            section.querySelectorAll('.series-control').forEach(series => {
                const xAxisSelect = series.querySelector('select[name="x-axis"]');
                const yAxisSelect = series.querySelector('select[name="y-axis"]');
                
                const currentX = xAxisSelect.value;
                xAxisSelect.innerHTML = '';
                cols.forEach(c => xAxisSelect.add(new Option(c.displayName, c.originalName)));
                xAxisSelect.value = currentX;

                const currentY = yAxisSelect.value;
                yAxisSelect.innerHTML = '';
                const numericCols = cols.filter(c => tableData && tableData.length > 0 && tableData[0][c.originalName] && !isNaN(parseFloat(String(tableData[0][c.originalName]).replace(',','.'))));
                numericCols.forEach(c => yAxisSelect.add(new Option(c.displayName, c.originalName)));
                yAxisSelect.value = currentY;
            });
            renderChart(id);
        });

        addSeriesControl(id, true);
        updateTableListUI();
    }

    function addSeriesControl(chartId, isFirst = false) {
        const seriesContainer = document.getElementById(`series-container-${chartId}`);
        const newSeries = document.createElement('div');
        newSeries.className = 'p-3 border rounded-lg bg-gray-50 grid grid-cols-1 sm:grid-cols-2 gap-3 items-end series-control';
        newSeries.innerHTML = `
            <div><label class="text-xs font-semibold">Eixo X / Grupo:</label><select name="x-axis" class="mt-1 block w-full rounded-md border-gray-300 text-sm"></select></div>
            <div><label class="text-xs font-semibold">Eixo Y / Valor:</label><div class="flex space-x-1"><select name="y-axis" class="mt-1 block w-2/3 rounded-md border-gray-300 text-sm"></select><select name="aggregation" class="mt-1 block w-1/3 rounded-md border-gray-300 text-sm"><option value="sum">Soma</option><option value="avg">Média</option><option value="count">Contagem</option><option value="min">Mínimo</option><option value="max">Máximo</option></select></div></div>
            <div class="combo-type-control" style="display: none;"><label class="text-xs font-semibold">Tipo:</label><select name="series-type" class="mt-1 block w-full rounded-md border-gray-300 text-sm"><option value="bar">Barra</option><option value="line">Linha</option></select></div>
            <div class="flex items-end space-x-2"><div class="w-full"><label class="text-xs font-semibold">Cor:</label><input type="color" value="#3b82f6" name="color" class="mt-1 w-full h-9 p-0 border-0 bg-white rounded-md"></div>
                ${!isFirst ? `<button type="button" class="remove-series-btn h-9 px-3 bg-red-500 text-white rounded-md hover:bg-red-600">&times;</button>` : ''}</div>`;
        
        if (!isFirst) {
            newSeries.querySelector('.remove-series-btn').onclick = () => {
                newSeries.remove();
                renderChart(chartId);
            };
        }
        seriesContainer.appendChild(newSeries);
        
        const dataSourceSelect = document.getElementById(`chart-data-source-${chartId}`);
        if(dataSourceSelect.value) {
           dataSourceSelect.dispatchEvent(new Event('change'));
        }
    }

    function renderChart(id) {
        if (chartInstances[id]) chartInstances[id].destroy();

        const section = document.getElementById(`chart-section-${id}`);
        if (!section) return;

        const dataSource = section.querySelector('.chart-data-source').value;
        const chartData = alasql.tables[dataSource]?.data;
        const canvas = document.getElementById(`mainChart-${id}`);
        const ctx = canvas.getContext('2d');

        if (!chartData || chartData.length === 0) {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            return;
        }
        
        const chartType = section.querySelector(`input[name="chart-type-${id}"]:checked`).value;
        section.querySelectorAll('.combo-type-control').forEach(el => el.style.display = chartType === 'combo' ? 'block' : 'none');
        
        const seriesControls = section.querySelectorAll('.series-control');
        if (seriesControls.length === 0) return;
        const firstXAxis = seriesControls[0].querySelector('[name="x-axis"]').value;
        if (!firstXAxis) {
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            return;
        }

        const labels = [...new Set(chartData.map(d => d[firstXAxis]))].sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric: true}));

        const datasets = Array.from(seriesControls).map(control => {
            const yCol = control.querySelector('[name="y-axis"]').value;
            const xCol = control.querySelector('[name="x-axis"]').value;
            const agg = control.querySelector('[name="aggregation"]').value;
            const seriesTypeOption = control.querySelector('[name="series-type"]').value;
            let seriesType = chartType === 'combo' ? seriesTypeOption : (chartType === 'line' ? 'line' : 'bar');
            
            const data = labels.map(label => {
                const group = chartData.filter(d => d[xCol] == label).map(r => parseFloat(String(r[yCol] || '0').replace(',', '.')) || 0);
                if (group.length === 0) return 0;
                switch(agg) {
                    case 'sum': return group.reduce((a, b) => a + b, 0);
                    case 'avg': return group.reduce((a, b) => a + b, 0) / group.length;
                    case 'count': return group.length;
                    case 'min': return Math.min(...group);
                    case 'max': return Math.max(...group);
                    default: return 0;
                }
            });

            return {
                label: `${yCol} (${agg})`,
                data: data, borderColor: control.querySelector('[name="color"]').value,
                backgroundColor: control.querySelector('[name="color"]').value + 'B3', type: seriesType,
                tension: parseFloat(section.querySelector('.line-interpolation').value) || 0.4,
                borderRadius: parseInt(section.querySelector('.bar-border-radius').value) || 0
            };
        });

        const fontColor = '#64748B';
        const gridColor = section.querySelector('.show-grid').checked ? 'rgba(0, 0, 0, 0.1)' : 'transparent';
        const yAxisAuto = section.querySelector('.y-axis-auto').checked;
        const yAxisMax = parseFloat(section.querySelector('.y-axis-max').value);
        const labelPos = section.querySelector('.label-position').value;
        const chartTitle = section.querySelector('.chart-title-input').value;
        const chartSubtitle = section.querySelector('.chart-subtitle-input').value;

        const options = {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                title: { display: !!chartTitle, text: chartTitle, font: { size: 18 } },
                subtitle: { display: !!chartSubtitle, text: chartSubtitle, padding: { bottom: 10 } },
                datalabels: {
                    display: section.querySelector('.show-labels').checked,
                    color: fontColor,
                    font: { size: parseInt(section.querySelector('.label-size').value) || 12 },
                    align: labelPos, anchor: labelPos === 'center' ? 'center' : (labelPos === 'start' ? 'start' : 'end'),
                    formatter: val => typeof val === 'number' ? val.toLocaleString('pt-BR') : val
                }
            },
            scales: {
                x: { ticks: { color: fontColor }, grid: { color: gridColor }, stacked: chartType === 'stacked' },
                y: { beginAtZero: true, ticks: { color: fontColor }, grid: { color: gridColor }, max: yAxisAuto ? undefined : yAxisMax, stacked: chartType === 'stacked' }
            },
            indexAxis: chartType === 'horizontalBar' ? 'y' : 'x'
        };

        chartInstances[id] = new Chart(canvas, { type: 'bar', data: { labels, datasets }, options });
    }

    function downloadChart(id) {
        const chart = chartInstances[id];
        if (!chart) { alert('Gere um gráfico para poder baixá-lo.'); return; }
        const link = document.createElement('a');
        link.href = chart.toBase64Image('image/png', 1.0);
        link.download = `power-graphx-chart-${id}.png`;
        link.click();
    }
'@

    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power-Graphx SQL Editor (AlaSQL)</title>
    <style>
        .modal { transition: opacity 0.25s ease; }
        #table-container { max-height: calc(100vh - 200px); overflow: auto; }
        table thead { position: sticky; top: 0; z-index: 10; }
        .chart-selector input:checked + label { border-color: #3b82f6; background-color: #eff6ff; box-shadow: 0 0 0 2px #3b82f6; }
        .divider { border-top: 1px solid #e5e7eb; margin-top: 1rem; margin-bottom: 1rem; }
        #charts-container { scroll-margin-top: 80px; }
    </style>
    $($CdnLibraryTags)
</head>
<body class="bg-gray-100 font-sans">
    <script id="jsonData" type="application/json">$JsonData</script>
    <script id="jsonColumnStructure" type="application/json">$JsonColumnStructure</script>
    <input type="file" id="csv-upload-input" multiple accept=".csv" class="hidden">

    <header class="bg-white shadow-md p-4 sticky top-0 z-20">
        <div class="container mx-auto">
            <div class="flex justify-between items-center">
                <h1 class="text-2xl font-bold text-gray-800">Power-Graphx SQL Editor</h1>
                <div class="flex items-center space-x-2">
                    <button id="btn-add-csv" class="px-4 py-2 text-sm font-medium text-white bg-orange-600 rounded-md hover:bg-orange-700">Adicionar CSV</button>
                    <button id="btn-filter" class="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700">Filtrar</button>
                    <button id="btn-toggle-sql" class="px-4 py-2 text-sm font-medium text-white bg-teal-600 rounded-md hover:bg-teal-700">Console SQL</button>
                    <button id="btn-add-chart" class="px-4 py-2 text-sm font-medium text-white bg-purple-600 rounded-md hover:bg-purple-700">Adicionar Gráfico</button>
                    <button id="btn-download-csv" class="px-4 py-2 text-sm font-medium text-white bg-gray-800 rounded-md hover:bg-gray-900">Baixar CSV</button>
                </div>
            </div>
            <div class="text-xs text-gray-500 mt-1" id="status-label">Carregando...</div>
        </div>
    </header>

    <main class="container mx-auto p-4">
        <section id="sql-section" class="hidden mb-6 bg-white rounded-lg shadow">
             <div class="p-6 grid grid-cols-1 md:grid-cols-4 gap-6">
                <div class="md:col-span-3">
                    <h2 class="text-2xl font-bold text-gray-800 mb-2">Console SQL (AlaSQL)</h2>
                    <div class="flex items-center space-x-2 mb-4">
                       <p class="text-sm text-gray-600">Exemplo: <code>SELECT * FROM source_data LIMIT 10;</code></p>
                       <button id="btn-sql-unpivot" class="px-3 py-1 text-xs font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700">Ajudante Unpivot</button>
                    </div>
                    <textarea id="sql-editor" class="w-full h-32 p-2 font-mono text-sm border border-gray-300 rounded-md" placeholder="SELECT * FROM source_data;">SELECT * FROM source_data;</textarea>
                </div>
                <div class="md:col-span-1">
                    <h3 class="text-lg font-bold text-gray-700 mb-2">Tabelas Carregadas</h3>
                    <div class="bg-gray-50 p-3 rounded-md h-32 overflow-y-auto">
                        <ul id="table-list" class="list-disc list-inside text-sm font-mono text-gray-800">
                        </ul>
                    </div>
                </div>
                <div class="md:col-span-4 mt-2 flex justify-between items-center">
                    <div id="sql-status" class="text-sm text-gray-500 italic">Aguardando inicialização...</div>
                    <div class="flex-shrink-0">
                        <button id="btn-reset-data" class="px-4 py-2 text-sm font-medium text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300">Resetar Dados</button>
                        <button id="btn-run-sql" class="px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700">Executar Consulta</button>
                    </div>
                </div>
             </div>
        </section>
        
        <div id="table-container" class="bg-white rounded-lg shadow overflow-hidden mb-6"></div>
        
        <div id="charts-container" class="space-y-6"></div>
    </main>

    <template id="chart-analysis-template">
        <section id="chart-section-__ID__" class="chart-analysis-section bg-white rounded-lg shadow" data-id="__ID__">
              <div class="p-6">
                  <div class="flex justify-between items-center mb-4">
                    <h2 class="text-2xl font-bold text-gray-800">Análise Gráfica __ID__</h2>
                    <button class="remove-chart-btn text-red-500 hover:text-red-700 font-bold text-xl">&times;</button>
                  </div>
                  <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">
                      <div class="lg:col-span-1 flex flex-col space-y-4">
                          <div>
                            <h3 class="font-bold text-gray-700 mb-2">1. Fonte de Dados</h3>
                            <select id="chart-data-source-__ID__" class="chart-data-source mt-1 block w-full rounded-md border-gray-300 text-sm"></select>
                          </div>
                          <div>
                              <h3 class="font-bold text-gray-700 mb-2">2. Tipo de Gráfico</h3>
                              <div class="chart-selector grid grid-cols-3 gap-2">
                                  <div><input type="radio" name="chart-type-__ID__" value="bar" id="type-bar-__ID__" checked class="hidden"><label for="type-bar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Barra</label></div>
                                  <div><input type="radio" name="chart-type-__ID__" value="line" id="type-line-__ID__" class="hidden"><label for="type-line-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Linha</label></div>
                                  <div><input type="radio" name="chart-type-__ID__" value="combo" id="type-combo-__ID__" class="hidden"><label for="type-combo-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Combo</label></div>
                                  <div><input type="radio" name="chart-type-__ID__" value="stacked" id="type-stacked-__ID__" class="hidden"><label for="type-stacked-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Empilhado</label></div>
                                  <div><input type="radio" name="chart-type-__ID__" value="horizontalBar" id="type-horizontalBar-__ID__" class="hidden"><label for="type-horizontalBar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Horizontal</label></div>
                              </div>
                          </div>
                          <div>
                              <div class="flex justify-between items-center mb-2"><h3 class="font-bold text-gray-700">3. Séries de Dados</h3><button class="add-series-btn text-xs bg-blue-500 text-white py-1 px-2 rounded-full hover:bg-blue-600">+ Série</button></div>
                              <div id="series-container-__ID__" class="space-y-3 max-h-60 overflow-y-auto"></div>
                          </div>
                      </div>
                      <div class="lg:col-span-2 bg-gray-50 rounded-lg p-4"><div class="relative w-full h-full min-h-[400px]"><canvas id="mainChart-__ID__"></canvas></div></div>
                      <div class="lg:col-span-1 flex flex-col space-y-2 text-sm">
                          <h3 class="font-bold text-gray-700 mb-2">4. Formatar Visual</h3>
                          <div><span class="font-semibold text-gray-700">Títulos</span>
                               <div class="mt-2 space-y-2">
                                  <div><label class="text-xs text-gray-600">Título:</label><input type="text" placeholder="Título do Gráfico" class="chart-title-input mt-1 block w-full rounded-md border-gray-300 text-xs"></div>
                                  <div><label class="text-xs text-gray-600">Subtítulo:</label><input type="text" placeholder="Subtítulo do Gráfico" class="chart-subtitle-input mt-1 block w-full rounded-md border-gray-300 text-xs"></div>
                              </div>
                          </div>
                          <div class="divider"></div>
                          <div><span class="font-semibold text-gray-700">Rótulos de Dados</span>
                              <div class="flex items-center mt-2"><input id="show-labels-__ID__" type="checkbox" class="show-labels h-4 w-4 rounded border-gray-300"><label for="show-labels-__ID__" class="ml-2 text-gray-900">Exibir rótulos</label></div>
                              <div class="mt-2 space-y-2">
                                  <div><label class="text-xs text-gray-600">Posição:</label><select class="label-position mt-1 block w-full rounded-md border-gray-300 text-xs"><option value="end">Topo/Direita</option><option value="center">Centro</option><option value="start">Base/Esquerda</option></select></div>
                                  <div><label class="text-xs text-gray-600">Tamanho Fonte:</label><input type="number" value="12" class="label-size mt-1 block w-full rounded-md border-gray-300 text-xs"></div>
                              </div>
                          </div>
                          <div class="divider"></div>
                          <div><span class="font-semibold text-gray-700">Opções de Barra/Linha</span>
                              <div class="mt-2"><label class="text-xs text-gray-600">Arredondamento da Borda:</label><input type="number" value="0" min="0" class="bar-border-radius mt-1 block w-full rounded-md border-gray-300 text-xs"></div>
                              <div class="mt-2"><label class="text-xs text-gray-600">Interpolação da Linha:</label><select class="line-interpolation mt-1 block w-full rounded-md border-gray-300 text-xs"><option value="0.0">Linear</option><option value="0.4" selected>Suave</option><option value="1.0">Curva Máxima</option></select></div>
                          </div>
                          <div class="divider"></div>
                          <div><span class="font-semibold text-gray-700">Eixos e Grade</span>
                              <div class="flex items-center mt-2"><input id="show-grid-__ID__" type="checkbox" checked class="show-grid h-4 w-4 rounded border-gray-300"><label for="show-grid-__ID__" class="ml-2 text-gray-900">Exibir grades</label></div>
                              <div class="flex items-center mt-2"><input id="y-axis-auto-__ID__" type="checkbox" checked class="y-axis-auto h-4 w-4 rounded border-gray-300"><label for="y-axis-auto-__ID__" class="ml-2 text-gray-900">Eixo Y Automático</label></div>
                              <input type="number" placeholder="Ex: 100" class="y-axis-max mt-1 block w-full rounded-md border-gray-300 text-xs" disabled>
                          </div>
                          <div class="divider"></div>
                          <div>
                              <h3 class="font-bold text-gray-700 mb-2">5. Ações</h3>
                              <button class="download-chart-btn w-full bg-gray-600 text-white font-bold py-2 rounded-lg hover:bg-gray-700 text-sm">Baixar Gráfico (PNG)</button>
                          </div>
                      </div>
                  </div>
              </div>
        </section>
    </template>

    <div id="filter-modal" class="modal hidden fixed inset-0 bg-gray-600 bg-opacity-50 overflow-y-auto h-full w-full z-30">
        <div class="relative top-20 mx-auto p-5 border w-96 shadow-lg rounded-md bg-white">
            <div class="flex justify-between items-center"><h3 class="text-lg font-medium text-gray-900">Filtrar Dados (Vista Atual)</h3><button class="modal-close font-bold text-xl">&times;</button></div>
            <div class="mt-4 space-y-4">
                <div><label for="filter-column" class="block text-sm font-medium text-gray-700">Coluna</label><select id="filter-column" class="mt-1 block w-full py-2 px-3 border border-gray-300 bg-white rounded-md shadow-sm"></select></div>
                <div><label for="filter-condition" class="block text-sm font-medium text-gray-700">Condição</label><select id="filter-condition" class="mt-1 block w-full py-2 px-3 border border-gray-300 bg-white rounded-md shadow-sm"><option value="contains">Contém</option><option value="not_contains">Não Contém</option><option value="equals">Igual a</option><option value="not_equals">Diferente de</option><option value="greater">Maior que (numérico)</option><option value="less">Menor que (numérico)</option></select></div>
                <div><label for="filter-value" class="block text-sm font-medium text-gray-700">Valor</label><input type="text" id="filter-value" class="mt-1 block w-full px-3 py-2 border border-gray-300 rounded-md shadow-sm"></div>
            </div>
            <div class="mt-6 flex justify-end space-x-2"><button class="modal-close px-4 py-2 bg-gray-200 text-gray-800 rounded-md hover:bg-gray-300">Cancelar</button><button id="apply-filter-btn" class="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700">Aplicar</button></div>
        </div>
    </div>

    <script>
        $ApplicationJavaScript
    </script>
</body>
</html>
"@
}

# --- 3. Função Principal de Execução ---
Function Start-WebApp {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $OpenFileDialog.Title = "Power-Graphx: Selecione o arquivo CSV inicial"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        Write-Host "Analisando: $(Split-Path $FilePath -Leaf)..."
        
        try {
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
        
        $ColumnStructure = $Data[0].PSObject.Properties | ForEach-Object {
            [PSCustomObject]@{
                originalName = $_.Name
                displayName  = $_.Name
            }
        }
        
        $JsonData = $Data | ConvertTo-Json -Compress -Depth 10
        $JsonColumnStructure = $ColumnStructure | ConvertTo-Json -Compress
        $cdnTags = Get-CdnLibraryTags
        $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -JsonColumnStructure $JsonColumnStructure -CdnLibraryTags $cdnTags
        
        $OutputPath = Join-Path $env:TEMP "PowerGraphx_AlaSQL_WebApp.html"
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
