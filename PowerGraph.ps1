# -----------------------------------------------------------------------------
# Power-Graphx Web App Launcher
# Versão: 6.6.2 - Lógica de Delimitador Aprimorada
# Autor: jefferson/configexe
#
# Melhorias da Versão 6.6.2:
# - INTEGRAÇÃO: Incorporada a lógica de detecção de delimitador com if/elseif
#   sugerida pelo usuário, que é mais clara, robusta e corrige o bug de
#   "chaves duplicadas". Excelente trabalho!
# - COMPATIBILIDADE: Adicionado -Encoding Default ao Import-Csv para melhor
#   compatibilidade com arquivos gerados por diferentes programas.
# -----------------------------------------------------------------------------

# --- 1. Configurações Iniciais e de Encoding ---
# Garante que todo o script opere em UTF-8 para compatibilidade de caracteres
$PSDefaultParameterValues['*:Encoding'] = 'utf8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# --- 2. Carregar Assemblies Necessárias ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Web
}
catch {
    Write-Error "Não foi possível carregar as assemblies .NET necessárias."
    exit 1
}

# --- 3. Funções Principais ---

# Função para fornecer URLs de CDN.
Function Get-CdnLibraryTags {
    $libs = @(
        '<script src="https://cdn.tailwindcss.com"></script>',
        '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js" integrity="sha384-e6cc9LaIG7xZ3XD5B+jtr1NhTWPQGQdRCh6xiZ+ZFUtWCpg4ycv3Sh+SkZoopvUY" crossorigin="anonymous"></script>',
        '<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js" integrity="sha384-y49Zu59jZHJL/PLKgZPv3k2WI9c0Yp3pWB76V8OBVCb0QBKS8l4Ff3YslzHVX76Y" crossorigin="anonymous"></script>',
        '<script src="https://cdn.jsdelivr.net/npm/alasql@4.1.2/dist/alasql.min.js" integrity="sha384-jJv67p3ipYhUXBEyC6HHwcdBifwMunNP2pOiuY2/6Hme7elFehskJ7cT2tfsKhJC" crossorigin="anonymous"></script>'
    )
    return $libs -join "`n    "
}


# Função que contém o template HTML, CSS e JavaScript da aplicação completa.
Function Get-HtmlTemplate {
    param(
        [Parameter(Mandatory=$true)]$JsonData,
        [Parameter(Mandatory=$true)]$CdnLibraryTags
    )
    
    $ApplicationJavaScript = @'
    // ---------------------------------------------------
    // Power-Graphx Web App - Lógica Principal (v6.6.0)
    // ---------------------------------------------------
    
    // Variáveis globais
    let originalData = [];
    let currentData = [];
    let columnStructure = [];
    let chartInstances = {};
    let chartAnalysisCounter = 0;
    const initialTableName = 'source_data';
    const sortState = {};
    let isEditMode = false;
    let conditionalFormattingRules = [];
    let chartObservers = {}; // Para gerenciar os observers de cada gráfico

    // Função Debounce para performance
    function debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }

    // Ponto de entrada
    window.onload = () => {
        try {
            Chart.register(ChartDataLabels);
            initializeApp();
            loadStateAndRestoreUI();
        } catch (e) {
            console.error("Falha crítica na inicialização:", e);
            document.getElementById('status-label').textContent = `Erro ao carregar a aplicação. Pressione F12 para ver detalhes.`;
            alert(`Ocorreu um erro ao carregar a aplicação. Pressione F12 para ver o console de erros.`);
        }
    };

    function initializeApp() {
        const jsonDataEl = document.getElementById('jsonData');
        try {
            if (jsonDataEl && jsonDataEl.textContent.trim()) {
                originalData = JSON.parse(jsonDataEl.textContent);
            } else {
                originalData = []; 
                throw new Error("Dados JSON não encontrados ou vazios no HTML.");
            }
        } catch(e) {
            console.error("Erro ao carregar dados JSON:", e);
            alert("Erro ao carregar dados iniciais. A aplicação pode não funcionar corretamente.");
            originalData = [];
        }

        currentData = JSON.parse(JSON.stringify(originalData)); 
        
        initializeDB();
        
        let initialColumnStructure = Object.keys(currentData[0] || {}).map(name => ({ originalName: name, displayName: name }));
        updateColumnStructure(initialColumnStructure);
        
        setupEventListeners();
        renderTable();
        updateStatus();
    }
    
    function initializeDB() {
        alasql.tables[initialTableName] = { data: JSON.parse(JSON.stringify(originalData)) };
        updateTableListUI();
        document.getElementById('sql-status').textContent = 'Motor SQL (AlaSQL) pronto.';
    }
    
    function updateTableListUI() {
        const tables = Object.keys(alasql.tables);
        const listEl = document.getElementById('table-list');
        listEl.innerHTML = '';
        tables.forEach(tableName => {
            const li = document.createElement('li');
            li.textContent = tableName;
            listEl.appendChild(li);
        });

        document.querySelectorAll('.chart-data-source, #calc-column-table').forEach(select => {
            const currentVal = select.value;
            select.innerHTML = '';
            tables.forEach(t => select.add(new Option(t, t)));
            select.value = tables.includes(currentVal) ? currentVal : tables[0];
            if (select.id !== 'calc-column-table') {
                 select.dispatchEvent(new Event('change'));
            }
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
            statusEl.innerHTML = `<span class="animate-pulse">⏳ Processando ${file.name}...</span>`;
            const reader = new FileReader();
            reader.onload = function(e) {
                const fileContent = e.target.result;
                try {
                    let data = alasql('SELECT * FROM CSV(?, {headers:true, separator:";"})', [fileContent]);
                    if (data.length > 0 && Object.keys(data[0]).length <= 1) {
                        const dataComma = alasql('SELECT * FROM CSV(?, {headers:true, separator:","})', [fileContent]);
                        if (dataComma.length > 0 && Object.keys(dataComma[0]).length > 1) { data = dataComma; }
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
            
            if (Array.isArray(resultData)) {
                currentData = resultData;
                let newColumns = resultData.length > 0 ? Object.keys(resultData[0]).map(name => ({ originalName: name, displayName: name })) : [];
                updateColumnStructure(newColumns);
                statusEl.textContent = `Consulta executada. ${resultData.length} linhas retornadas.`;
            } else {
                updateTableListUI();
                statusEl.textContent = `Comando executado. A lista de tabelas foi atualizada.`;
            }
            
            Object.keys(sortState).forEach(key => delete sortState[key]);
            renderTable();
            updateStatus();
            document.querySelectorAll('.chart-analysis-section').forEach(section => renderChart(section.dataset.id));
        } catch (e) {
            statusEl.textContent = `Erro na consulta SQL: ${e.message}`;
            console.error(e);
            alert(`Erro na consulta SQL: ${e.message}\n\nDica: Verifique se os nomes das colunas estão entre crases (\`Nome da Coluna\`). Use o botão "Formatar SQL" para ajudar.`);
        }
    }

    function formatSql() {
        const editor = document.getElementById('sql-editor');
        let query = editor.value;
        
        query = query.replace(/;(\s*DROP|\s*DELETE|\s*UPDATE)/gi, '; /* COMANDO BLOQUEADO */ $1');

        const keywords = new Set(['SELECT', 'FROM', 'WHERE', 'GROUP', 'BY', 'ORDER', 'AS', 'ON', 'LEFT', 'RIGHT', 'INNER', 'OUTER', 'JOIN', 'LIMIT', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'AND', 'OR', 'NOT', 'IN', 'LIKE', 'IS', 'NULL', 'COUNT', 'SUM', 'AVG', 'MIN', 'MAX', 'DISTINCT']);
        
        query = query.replace(/(?<![`'"])\b([a-zA-Z_][\w]*)\b(?![`'"])/g, (match, word) => {
            if (keywords.has(word.toUpperCase()) || !isNaN(word)) {
                return word;
            }
            return `\`${word}\``;
        });
        
        editor.value = query;
    }
    
    function setupEventListeners() {
        document.getElementById('btn-filter').addEventListener('click', () => {
            updateModalColumnSelector('filter-column');
            document.getElementById('filter-modal').classList.remove('hidden');
        });
        document.getElementById('btn-cond-format').addEventListener('click', () => {
            updateModalColumnSelector('cond-format-column');
            renderCondFormatRules();
            document.getElementById('cond-format-modal').classList.remove('hidden');
        });
        document.getElementById('btn-edit-data').addEventListener('click', toggleEditMode);
        document.getElementById('add-cond-format-rule-btn').addEventListener('click', addCondFormatRule);
        document.getElementById('btn-download-csv').addEventListener('click', downloadCSV);
        document.getElementById('btn-save-state').addEventListener('click', saveState);
        document.getElementById('btn-toggle-sql').addEventListener('click', () => document.getElementById('sql-section').classList.toggle('hidden'));
        document.getElementById('btn-run-sql').addEventListener('click', runQueryAndUpdateUI);
        document.getElementById('btn-format-sql').addEventListener('click', formatSql);
        document.getElementById('btn-new-column').addEventListener('click', () => document.getElementById('calc-column-modal').classList.remove('hidden'));
        document.getElementById('apply-calc-column-btn').addEventListener('click', applyCalculatedColumn);
        document.getElementById('btn-reset-data').addEventListener('click', () => {
            const freshData = JSON.parse(JSON.stringify(originalData));
            Object.keys(alasql.tables).forEach(key => { if(key !== initialTableName) delete alasql.tables[key]; });
            alasql(`DROP TABLE IF EXISTS ${initialTableName}`);
            alasql.tables[initialTableName] = { data: freshData };
            currentData = freshData;
            
            updateTableListUI();
            let initialCols = currentData.length > 0 ? Object.keys(currentData[0]).map(name => ({ originalName: name, displayName: name })) : [];
            updateColumnStructure(initialCols);
            document.getElementById('sql-editor').value = `SELECT * FROM ${initialTableName};`;
            runQueryAndUpdateUI();
        });
        document.getElementById('apply-filter-btn').addEventListener('click', () => applyFilter(true));
        document.querySelectorAll('.modal-close').forEach(el => el.addEventListener('click', () => el.closest('.modal').classList.add('hidden')));
        document.getElementById('btn-add-csv').addEventListener('click', () => document.getElementById('csv-upload-input').click());
        document.getElementById('csv-upload-input').addEventListener('change', handleFileUploads);
        document.getElementById('btn-update-source').addEventListener('click', () => document.getElementById('csv-update-input').click());
        document.getElementById('csv-update-input').addEventListener('change', handleDataSourceUpdate);
        document.getElementById('btn-add-chart').addEventListener('click', addChartAnalysis);
    }
    
    function handleDataSourceUpdate(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const fileContent = e.target.result;
            try {
                let newData = alasql('SELECT * FROM CSV(?, {headers:true, separator:";"})', [fileContent]);
                if (newData.length > 0 && Object.keys(newData[0]).length <= 1) {
                    const dataComma = alasql('SELECT * FROM CSV(?, {headers:true, separator:","})', [fileContent]);
                    if (dataComma.length > 0 && Object.keys(dataComma[0]).length > 1) { newData = dataComma; }
                }

                if (newData.length === 0) {
                    alert("Erro: O novo arquivo CSV está vazio ou em um formato inválido.");
                    return;
                }
                
                const originalHeaders = Object.keys(originalData[0] || {}).sort();
                const newHeaders = Object.keys(newData[0]).sort();

                if (JSON.stringify(originalHeaders) !== JSON.stringify(newHeaders)) {
                    alert("Erro: As colunas do novo arquivo não correspondem à fonte de dados original.\n\nOriginal: " + originalHeaders.join(', ') + "\nNovo: " + newHeaders.join(', ') + "\n\nA atualização foi cancelada.");
                    return;
                }

                alasql.tables[initialTableName].data = newData;
                originalData = JSON.parse(JSON.stringify(newData));

                alert("Fonte de dados atualizada com sucesso! Recalculando visualizações...");
                document.getElementById('sql-editor').value = `SELECT * FROM ${initialTableName};`;
                runQueryAndUpdateUI();
                
            } catch(err) {
                alert(`Ocorreu um erro ao processar o novo arquivo CSV: ${err.message}`);
                console.error(err);
            }
        };
        reader.readAsText(file);
        event.target.value = '';
    }

    function updateColumnStructure(newStructure) {
        columnStructure = newStructure.map(col => ({ ...col, displayName: col.displayName || col.originalName }));
        updateModalColumnSelector('filter-column');
        updateModalColumnSelector('cond-format-column');
    }

    function updateModalColumnSelector(selectId) {
        const selectElement = document.getElementById(selectId);
        if(!selectElement) return;
        const currentVal = selectElement.value;
        selectElement.innerHTML = '<option value="">-- Selecione a Coluna --</option>';
        columnStructure.forEach(col => selectElement.add(new Option(col.displayName, col.originalName)));
        selectElement.value = currentVal;
    }

    function updateStatus() {
        document.getElementById('status-label').textContent = `Exibindo ${currentData.length} registros.`;
    }

    function toggleEditMode() {
        isEditMode = !isEditMode;
        const btn = document.getElementById('btn-edit-data');
        btn.classList.toggle('bg-red-600', isEditMode);
        btn.classList.toggle('hover:bg-red-700', isEditMode);
        btn.classList.toggle('bg-yellow-600', !isEditMode);
        btn.classList.toggle('hover:bg-yellow-700', !isEditMode);
        btn.textContent = isEditMode ? 'Sair do Modo de Edição' : 'Editar Dados';
        renderTable();
    }

    function handleCellEdit(e) {
        const td = e.target;
        const rowIndex = parseInt(td.parentElement.dataset.rowId);
        const colName = td.dataset.columnName;
        
        if(rowIndex >= 0 && colName && currentData[rowIndex]) {
            currentData[rowIndex][colName] = td.textContent;
            document.querySelectorAll('.chart-analysis-section').forEach(section => renderChart(section.dataset.id));
        }
    }
    
    function renderTable() {
        const tableContainer = document.getElementById('table-container');
        tableContainer.innerHTML = ''; 
        if (!currentData || currentData.length === 0) {
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
            th.className = 'px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider cursor-pointer select-none';
            th.textContent = col.displayName;
            th.addEventListener('click', () => handleSort(col.originalName));
            if(sortState[col.originalName]) {
                th.innerHTML += sortState[col.originalName] === 'asc' ? ' &#9650;' : ' &#9660;';
            }
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
                td.dataset.columnName = col.originalName;
                if(isEditMode) {
                    td.setAttribute('contenteditable', 'true');
                    td.classList.add('bg-yellow-100', 'outline-blue-400');
                    td.addEventListener('blur', handleCellEdit);
                }
                tr.appendChild(td);
            });
            applyRowFormatting(tr, row);
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        tableContainer.appendChild(table);
    }

    function handleSort(columnName) {
        const currentOrder = sortState[columnName];
        let nextOrder = (currentOrder === 'asc') ? 'desc' : (currentOrder === 'desc' ? undefined : 'asc');
        Object.keys(sortState).forEach(key => delete sortState[key]);
        if (nextOrder) {
            sortState[columnName] = nextOrder;
            currentData.sort((a, b) => {
                const valA = a[columnName], valB = b[columnName];
                const numA = parseFloat(String(valA).replace(',', '.')), numB = parseFloat(String(valB).replace(',', '.'));
                let comparison = (!isNaN(numA) && !isNaN(numB)) ? numA - numB : String(valA || '').toLowerCase().localeCompare(String(valB || '').toLowerCase());
                return nextOrder === 'asc' ? comparison : -comparison;
            });
        }
        renderTable();
    }
    
    function addCondFormatRule() {
        const column = document.getElementById('cond-format-column').value;
        const condition = document.getElementById('cond-format-condition').value;
        const value = document.getElementById('cond-format-value').value;
        const color = document.getElementById('cond-format-color').value;
        const applyTo = document.querySelector('input[name="cond-format-apply-to"]:checked').value;
        if(!column || !value) { alert("Selecione a coluna e o valor."); return; }
        conditionalFormattingRules.push({ id: Date.now(), column, condition, value, color, applyTo });
        renderCondFormatRules();
        renderTable();
    }

    function removeCondFormatRule(id) {
        conditionalFormattingRules = conditionalFormattingRules.filter(rule => rule.id !== id);
        renderCondFormatRules();
        renderTable();
    }

    function renderCondFormatRules() {
        const container = document.getElementById('cond-format-rules-list');
        container.innerHTML = '';
        if (conditionalFormattingRules.length === 0) {
            container.innerHTML = '<p class="text-xs text-gray-500">Nenhuma regra aplicada.</p>';
            return;
        }
        conditionalFormattingRules.forEach(rule => {
            const ruleDiv = document.createElement('div');
            ruleDiv.className = 'flex items-center justify-between p-2 bg-gray-50 rounded-md';
            const applyToText = rule.applyTo === 'row' ? 'na linha inteira' : 'na célula';
            ruleDiv.innerHTML = `<div class="flex items-center text-sm">
                <div class="w-4 h-4 rounded-full mr-2 border border-gray-300" style="background-color: ${rule.color};"></div>
                <span>Se <b>${rule.column}</b> ${rule.condition.replace(/_/g, ' ')} <b>${rule.value}</b> (aplicar ${applyToText})</span>
            </div>
            <button class="text-red-500 font-bold hover:text-red-700" data-rule-id="${rule.id}">&times;</button>`;
            ruleDiv.querySelector('button').addEventListener('click', () => removeCondFormatRule(rule.id));
            container.appendChild(ruleDiv);
        });
    }

    function applyRowFormatting(tr, rowData) {
        tr.querySelectorAll('td').forEach(td => {
             td.style.backgroundColor = '';
             td.style.color = '';
        });

        for (const rule of conditionalFormattingRules) {
            const cellValue = rowData[rule.column];
            const ruleVal = rule.value;
            const cellValStr = String(cellValue || '').toLowerCase();
            const ruleValStr = String(ruleVal).toLowerCase();
            const cellValNum = parseFloat(String(cellValue).replace(',', '.'));
            const ruleValNum = parseFloat(String(ruleVal).replace(',', '.'));
            let match = false;

            switch(rule.condition) {
                case 'greater': match = !isNaN(cellValNum) && !isNaN(ruleValNum) && cellValNum > ruleValNum; break;
                case 'less': match = !isNaN(cellValNum) && !isNaN(ruleValNum) && cellValNum < ruleValNum; break;
                case 'equals': match = cellValStr === ruleValStr; break;
                case 'not_equals': match = cellValStr !== ruleValStr; break;
                case 'contains': match = cellValStr.includes(ruleValStr); break;
                case 'not_contains': match = !cellValStr.includes(ruleValStr); break;
            }

            if (match) {
                const hex = rule.color.replace('#', '');
                const r = parseInt(hex.substring(0, 2), 16);
                const g = parseInt(hex.substring(2, 4), 16);
                const b = parseInt(hex.substring(4, 6), 16);
                const brightness = ((r * 299) + (g * 587) + (b * 114)) / 1000;
                const textColor = brightness > 125 ? 'black' : 'white';

                const apply = (cell) => {
                    cell.style.backgroundColor = rule.color;
                    cell.style.color = textColor;
                };

                if (rule.applyTo === 'row') {
                    tr.querySelectorAll('td').forEach(apply);
                } else {
                    const targetCell = tr.querySelector(`td[data-column-name="${rule.column}"]`);
                    if (targetCell) apply(targetCell);
                }
            }
        }
    }
    
    function addChartAnalysis() {
        chartAnalysisCounter++;
        const template = document.getElementById('chart-analysis-template').innerHTML;
        const newChartHtml = template.replace(/__ID__/g, chartAnalysisCounter);
        const container = document.getElementById('charts-container');
        const div = document.createElement('div');
        div.innerHTML = newChartHtml;
        const newSection = div.firstElementChild;
        container.appendChild(newSection);
        
        const observer = new IntersectionObserver((entries) => {
            entries.forEach(entry => {
                const id = entry.target.dataset.id;
                const navLink = document.querySelector(`#chart-nav a[data-nav-id="${id}"]`);
                if (!navLink) return;
                
                if (entry.isIntersecting) {
                    document.querySelectorAll('#chart-nav a.font-bold').forEach(a => a.classList.remove('bg-blue-100', 'text-blue-700', 'font-bold'));
                    navLink.classList.add('bg-blue-100', 'text-blue-700', 'font-bold');
                }
            });
        }, { rootMargin: '-40% 0px -40% 0px', threshold: 0.1 });
        observer.observe(newSection);
        chartObservers[chartAnalysisCounter] = observer;

        initializeChartUI(chartAnalysisCounter);
        updateChartNav();
    }

    function removeChartAnalysis(id) {
        const section = document.getElementById(`chart-section-${id}`);
        if (section) {
            if (chartObservers[id]) {
                chartObservers[id].unobserve(section);
                delete chartObservers[id];
            }
            section.remove();
        }
        if (chartInstances[id]) {
            chartInstances[id].destroy();
            delete chartInstances[id];
        }
        updateChartNav();
    }
    
    function initializeChartUI(id) {
        const section = document.getElementById(`chart-section-${id}`);
        if (!section) return;

        const debouncedRenderChart = debounce((chartId) => renderChart(chartId), 300);

        const titleInput = section.querySelector('.chart-title-input-main');
        titleInput.addEventListener('input', () => {
             debouncedRenderChart(id);
             updateChartNav();
        });

        section.querySelectorAll(`input[name="chart-type-${id}"]`).forEach(radio => {
            radio.addEventListener('change', (e) => {
                const chartType = e.target.value;
                const isCombo = chartType === 'combo';
                const isFloating = chartType === 'floatingBar';
                const isLineArea = ['line', 'area'].includes(chartType);

                section.querySelector('.point-styling-options').style.display = isLineArea ? 'block' : 'none';

                section.querySelectorAll('.series-control').forEach(control => {
                    control.querySelector('.combo-type-control').style.display = isCombo ? 'block' : 'none';
                    control.querySelector('.secondary-axis-control').style.display = isCombo ? 'block' : 'none';
                    control.querySelector('.y-axis-label').textContent = isFloating ? 'Valor Final:' : 'Eixo Y / Valor:';
                    control.querySelector('.y-axis-start-control').style.display = isFloating ? 'block' : 'none';
                    control.querySelector('select[name="aggregation"]').disabled = isFloating;
                });
                renderChart(id); 
            });
        });

        section.addEventListener('change', (e) => {
            if (!e.target.classList.contains('chart-title-input-main')) {
                renderChart(id);
            }
        });
        section.addEventListener('input', (e) => {
             if (!e.target.classList.contains('chart-title-input-main')) {
                debouncedRenderChart(id);
            }
        });

        section.querySelector('.add-series-btn').addEventListener('click', () => addSeriesControl(id));
        section.querySelector('.download-chart-btn').addEventListener('click', () => downloadChart(id));
        section.querySelector('.remove-chart-btn').addEventListener('click', () => removeChartAnalysis(id));
        section.querySelector('.y-axis-auto').onchange = (e) => { section.querySelector('.y-axis-max').disabled = e.target.checked; };
        
        const dataSourceSelect = section.querySelector('.chart-data-source');
        dataSourceSelect.addEventListener('change', (e) => {
            const selectedTable = e.target.value;
            const tableData = alasql.tables[selectedTable]?.data;
            updateChartAxisSelectors(id, tableData);
            renderChart(id);
        });

        const groupBySelect = section.querySelector('.group-by-select');
        groupBySelect.addEventListener('change', (e) => {
            const hasGroupBy = !!e.target.value;
            const seriesContainer = section.querySelector(`#series-container-${id}`);
            const allSeries = seriesContainer.querySelectorAll('.series-control');
            
            section.querySelector('.add-series-btn').style.display = hasGroupBy ? 'none' : 'inline-block';

            if (hasGroupBy && allSeries.length > 1) {
                for (let i = allSeries.length - 1; i > 0; i--) {
                    allSeries[i].remove();
                }
            }
        });

        addSeriesControl(id, true);
        updateTableListUI();
        const selectedTable = dataSourceSelect.value;
        if(selectedTable && alasql.tables[selectedTable]?.data) {
             updateChartAxisSelectors(id, alasql.tables[selectedTable].data);
             renderChart(id);
        }
    }
    
    function addSeriesControl(chartId, isFirst = false) {
        const seriesContainer = document.getElementById(`series-container-${chartId}`);
        const newSeries = document.createElement('div');
        newSeries.className = 'p-3 border rounded-lg bg-gray-50 grid grid-cols-1 sm:grid-cols-2 gap-3 items-end series-control';
        newSeries.innerHTML = `
            <div><label class="text-xs font-semibold">Eixo X / Categoria:</label><select name="x-axis" class="mt-1 block w-full rounded-md border-gray-300 text-sm"></select></div>
            <div class="y-axis-start-control" style="display: none;"><label class="text-xs font-semibold">Valor Inicial:</label><select name="y-axis-start" class="mt-1 block w-full rounded-md border-gray-300 text-sm"></select></div>
            <div><label class="text-xs font-semibold y-axis-label">Eixo Y / Valor:</label><div class="flex space-x-1"><select name="y-axis" class="mt-1 block w-2/3 rounded-md border-gray-300 text-sm"></select><select name="aggregation" class="mt-1 block w-1/3 rounded-md border-gray-300 text-sm"><option value="sum">Soma</option><option value="avg">Média</option><option value="count">Contagem</option><option value="min">Mínimo</option><option value="max">Máximo</option><option value="percentage_total">% do Total</option><option value="none">Nenhum</option></select></div></div>
            <div class="sm:col-span-2"><label class="text-xs font-semibold">Nome da Série (Legenda):</label><input type="text" name="series-label" class="mt-1 block w-full rounded-md border-gray-300 text-sm" placeholder="Opcional"></div>
            <div class="combo-type-control" style="display: none;"><label class="text-xs font-semibold">Tipo:</label><select name="series-type" class="mt-1 block w-full rounded-md border-gray-300 text-sm"><option value="bar">Barra</option><option value="line">Linha</option></select></div>
            <div class="secondary-axis-control" style="display: none;"><label class="flex items-center text-xs font-semibold"><input type="checkbox" name="secondary-axis" class="h-4 w-4 mr-2 rounded border-gray-300">Usar Eixo Secundário</label></div>
            <div class="flex items-end space-x-2"><div class="w-full"><label class="text-xs font-semibold">Cor:</label><input type="color" value="#${(0x1000000+Math.random()*0xffffff).toString(16).substr(1,6)}" name="color" class="mt-1 w-full h-9 p-0 border-0 bg-white rounded-md"></div>
                ${!isFirst ? `<button type="button" class="remove-series-btn h-9 px-3 bg-red-500 text-white rounded-md hover:bg-red-600">&times;</button>` : ''}</div>`;
        
        if (!isFirst) {
            newSeries.querySelector('.remove-series-btn').onclick = () => { newSeries.remove(); renderChart(chartId); };
        }

        const section = document.getElementById(`chart-section-${chartId}`);
        const chartType = section.querySelector(`input[name="chart-type-${chartId}"]:checked`).value;
        if (chartType === 'combo') {
            newSeries.querySelector('.combo-type-control').style.display = 'block';
            newSeries.querySelector('.secondary-axis-control').style.display = 'block';
        }
        if (chartType === 'floatingBar') {
            newSeries.querySelector('.y-axis-start-control').style.display = 'block';
            newSeries.querySelector('.y-axis-label').textContent = 'Valor Final:';
            newSeries.querySelector('select[name="aggregation"]').disabled = true;
        }

        seriesContainer.appendChild(newSeries);
        const dataSourceSelect = document.getElementById(`chart-data-source-${chartId}`);
        if(dataSourceSelect.value) {
           updateChartAxisSelectors(chartId, alasql.tables[dataSourceSelect.value]?.data);
        }
    }
    
    function updateChartAxisSelectors(chartId, data) {
        const section = document.getElementById(`chart-section-${chartId}`);
        if (!section || !data || data.length === 0) return;
        const cols = Object.keys(data[0]).map(name => ({ originalName: name, displayName: name }));
        
        const groupBySelect = section.querySelector('select[name="group-by"]');
        const currentGroupBy = groupBySelect.value;
        groupBySelect.innerHTML = '<option value="">-- Nenhum --</option>';
        cols.forEach(c => groupBySelect.add(new Option(c.displayName, c.originalName)));
        if(cols.find(c => c.originalName === currentGroupBy)) groupBySelect.value = currentGroupBy;

        section.querySelectorAll('.series-control').forEach(series => {
            const xAxisSelect = series.querySelector('select[name="x-axis"]');
            const yAxisSelect = series.querySelector('select[name="y-axis"]');
            const yAxisStartSelect = series.querySelector('select[name="y-axis-start"]');

            const currentX = xAxisSelect.value, currentY = yAxisSelect.value, currentYStart = yAxisStartSelect.value;
            
            [xAxisSelect, yAxisSelect, yAxisStartSelect].forEach(sel => { sel.innerHTML = ''; });
            
            cols.forEach(c => {
                xAxisSelect.add(new Option(c.displayName, c.originalName));
                yAxisSelect.add(new Option(c.displayName, c.originalName));
                yAxisStartSelect.add(new Option(c.displayName, c.originalName));
            });
            
            if (cols.find(c => c.originalName === currentX)) xAxisSelect.value = currentX;
            if (cols.find(c => c.originalName === currentY)) yAxisSelect.value = currentY;
            if (cols.find(c => c.originalName === currentYStart)) yAxisStartSelect.value = currentYStart;
        });
    }

    function renderChart(id) {
        if (chartInstances[id]) chartInstances[id].destroy();
        const section = document.getElementById(`chart-section-${id}`);
        if (!section) return;

        const chartType = section.querySelector(`input[name="chart-type-${id}"]:checked`).value;
        const chartData = alasql.tables[section.querySelector('.chart-data-source').value]?.data;
        const canvas = section.querySelector('canvas');
        
        if (!chartData || chartData.length === 0) {
            if (canvas) { canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height); }
            return;
        }
        
        const seriesControls = section.querySelectorAll('.series-control');
        if (seriesControls.length === 0) return;

        const firstXAxis = seriesControls[0].querySelector('[name="x-axis"]').value;
        if (!firstXAxis) return;
        
        const labels = [...new Set(chartData.map(d => d[firstXAxis]))].sort((a,b) => String(a).localeCompare(String(b), undefined, {numeric: true}));
        let datasets = [];
        const groupingColumn = section.querySelector('select[name="group-by"]').value;
        const colorPalette = ['#3b82f6', '#ef4444', '#22c55e', '#f97316', '#8b5cf6', '#ec4899', '#14b8a6', '#eab308'];

        if (groupingColumn && chartType !== 'floatingBar') {
             const uniqueGroups = [...new Set(chartData.map(d => d[groupingColumn]))].sort();
             const control = seriesControls[0];
             const yCol = control.querySelector('[name="y-axis"]').value, xCol = control.querySelector('[name="x-axis"]').value, agg = control.querySelector('[name="aggregation"]').value;
             const seriesTypeOption = control.querySelector('[name="series-type"]').value;
             const useSecondaryAxis = control.querySelector('[name="secondary-axis"]').checked;

             let totalForPercent = 1;
             if (agg === 'percentage_total') {
                 totalForPercent = chartData.map(r => parseFloat(String(r[yCol] || '0').replace(',', '.')) || 0).reduce((a, b) => a + b, 0);
             }

             uniqueGroups.forEach((groupValue, index) => {
                 const groupFilteredData = chartData.filter(d => d[groupingColumn] == groupValue);
                 const data = labels.map(label => {
                     const groupDataForLabel = groupFilteredData.filter(d => d[xCol] == label);
                     return calculateAggregation(groupDataForLabel, yCol, agg, totalForPercent);
                 });
                 const color = colorPalette[index % colorPalette.length];
                 datasets.push({ label: groupValue, data, borderColor: color, backgroundColor: color + 'B3', type: seriesTypeOption, yAxisID: useSecondaryAxis ? 'y1' : 'y', stack: chartType === 'groupedStackedBar' ? xCol : undefined });
             });
        } else {
             datasets = Array.from(seriesControls).map(control => {
                 const xCol = control.querySelector('[name="x-axis"]').value;
                 const yCol = control.querySelector('[name="y-axis"]').value;
                 const yColStart = control.querySelector('[name="y-axis-start"]').value;
                 const agg = control.querySelector('[name="aggregation"]').value;
                 const customLabel = control.querySelector('input[name="series-label"]').value;
                 const seriesTypeOption = control.querySelector('[name="series-type"]').value;
                 const useSecondaryAxis = control.querySelector('[name="secondary-axis"]').checked;

                 let totalForPercent = 1;
                 if (agg === 'percentage_total') {
                     totalForPercent = chartData.map(r => parseFloat(String(r[yCol] || '0').replace(',', '.')) || 0).reduce((a, b) => a + b, 0);
                 }
                 
                 const data = labels.map(label => {
                     const groupData = chartData.filter(d => d[xCol] == label);
                     if (chartType === 'floatingBar') {
                         if (groupData.length > 0) {
                             const startVal = parseFloat(String(groupData[0][yColStart] || '0').replace(',', '.'));
                             const endVal = parseFloat(String(groupData[0][yCol] || '0').replace(',', '.'));
                             return [startVal, endVal];
                         }
                         return [0,0];
                     }
                     return calculateAggregation(groupData, yCol, agg, totalForPercent);
                 });

                 let seriesLabel = customLabel.trim() || `${yCol}`;
                 if (agg !== 'none' && chartType !== 'floatingBar') {
                    const aggText = control.querySelector('[name="aggregation"] option:checked').textContent;
                    seriesLabel += ` (${aggText})`;
                 }
                 
                 return { label: seriesLabel, data, borderColor: control.querySelector('[name="color"]').value, backgroundColor: control.querySelector('[name="color"]').value + 'B3', type: seriesTypeOption, yAxisID: useSecondaryAxis ? 'y1' : 'y' };
             });
        }
        
        const options = buildChartOptions(section, datasets);
        let finalChartType = 'bar';
        if(chartType === 'combo') {
            finalChartType = 'bar';
        } else if (chartType === 'area') {
            finalChartType = 'line';
            datasets.forEach(ds => ds.fill = true);
        } else if (chartType === 'stackedBar' || chartType === 'groupedStackedBar') {
            finalChartType = 'bar';
            options.scales.x.stacked = true;
            options.scales.y.stacked = true;
        } else if (chartType === 'horizontalBar' || chartType === 'floatingBar') {
            finalChartType = 'bar';
            options.indexAxis = 'y';
        } else {
            finalChartType = chartType;
        }
        
        datasets.forEach(ds => {
            ds.tension = parseFloat(section.querySelector('.line-interpolation').value) || 0.4;
            ds.borderRadius = parseInt(section.querySelector('.bar-border-radius').value) || 0;
            if(chartType !== 'combo') ds.type = finalChartType;
        });

        chartInstances[id] = new Chart(canvas, { type: 'bar', data: { labels, datasets }, options });
    }

    function calculateAggregation(data, column, aggType, totalForPercent = 1) {
        if (data.length === 0) return 0;
        const numericValues = data.map(r => parseFloat(String(r[column] || '0').replace(',', '.')) || 0).filter(v => !isNaN(v));

        switch(aggType) {
            case 'count': return data.length;
            case 'none': return numericValues.length > 0 ? numericValues[0] : 0;
            case 'sum': return numericValues.reduce((a, b) => a + b, 0);
            case 'avg': return numericValues.length > 0 ? numericValues.reduce((a, b) => a + b, 0) / numericValues.length : 0;
            case 'min': return numericValues.length > 0 ? Math.min(...numericValues) : 0;
            case 'max': return numericValues.length > 0 ? Math.max(...numericValues) : 0;
            case 'percentage_total':
                const sum = numericValues.reduce((a, b) => a + b, 0);
                return totalForPercent > 0 ? (sum / totalForPercent) * 100 : 0;
            default: return 0;
        }
    }

    function buildChartOptions(section, datasets) {
        const chartSubtitle = section.querySelector('.chart-subtitle-input').value;
        const yAxisAuto = section.querySelector('.y-axis-auto').checked, yAxisMax = parseFloat(section.querySelector('.y-axis-max').value);
        const labelPos = section.querySelector('.label-position').value;
        const gridColor = section.querySelector('.show-grid').checked ? 'rgba(0, 0, 0, 0.1)' : 'transparent';
        const fontColor = '#64748B';

        datasets.forEach(ds => {
            ds.pointStyle = section.querySelector('.point-style').value;
            ds.radius = parseInt(section.querySelector('.point-size').value);
            ds.hoverRadius = parseInt(section.querySelector('.point-size').value) + 2;
        });

        const options = {
            responsive: true, maintainAspectRatio: false,
            plugins: {
                title: { display: false },
                subtitle: { display: !!chartSubtitle, text: chartSubtitle, padding: { bottom: 10 } },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) { label += ': '; }
                            const isPercent = label.includes('% do Total');
                            let value;
                            if (context.dataset.data[context.dataIndex] && Array.isArray(context.dataset.data[context.dataIndex])) {
                                value = `[${context.dataset.data[context.dataIndex].join(', ')}]`;
                            } else {
                                value = context.parsed.y;
                                if (context.chart.options.indexAxis === 'y') {
                                    value = context.parsed.x;
                                }
                            }

                            if (value !== null && typeof value !== 'string') {
                                label += value.toLocaleString('pt-BR', { maximumFractionDigits: 2 });
                                if (isPercent) label += '%';
                            } else {
                                label += value;
                            }
                            return label;
                        }
                    }
                },
                datalabels: {
                    display: section.querySelector('.show-labels').checked, color: fontColor,
                    font: { size: parseInt(section.querySelector('.label-size').value) || 12 },
                    align: labelPos, anchor: labelPos === 'center' ? 'center' : (labelPos === 'start' ? 'start' : 'end'),
                    formatter: (value, ctx) => {
                        if (Array.isArray(value)) return null;
                        const isPercent = ctx.dataset.label.includes('% do Total');
                        let formattedVal = typeof value === 'number' ? value.toLocaleString('pt-BR', { maximumFractionDigits: 2 }) : value;
                        return isPercent ? formattedVal + '%' : formattedVal;
                    }
                }
            },
            scales: {
                y: { beginAtZero: true, type: 'linear', position: 'left', max: yAxisAuto ? undefined : yAxisMax, grid: { color: gridColor }, ticks: { color: fontColor } },
                x: { grid: { color: gridColor }, ticks: { color: fontColor } }
            }
        };

        if (datasets.some(ds => ds.yAxisID === 'y1')) {
            options.scales.y1 = { type: 'linear', position: 'right', beginAtZero: true, grid: { drawOnChartArea: false }, ticks: { color: fontColor } };
        }
        return options;
    }
    
    function applyCalculatedColumn() {
        const tableName = document.getElementById('calc-column-table').value;
        const newColumnName = document.getElementById('calc-column-name').value;
        let formula = document.getElementById('calc-column-formula').value;
        if (!tableName || !newColumnName.trim() || !formula.trim()) { alert("Por favor, preencha todos os campos."); return; }
        formula = formula.replace(/\[([^\]]+)\]/g, '`$1`');

        try {
            const tableData = alasql.tables[tableName].data;
            const query = `SELECT *, ${formula} AS \`${newColumnName}\` FROM ?`;
            const newData = alasql(query, [tableData]);
            alasql.tables[tableName].data = newData;

            alert(`Coluna "${newColumnName}" adicionada à tabela "${tableName}" com sucesso.`);
            document.getElementById('calc-column-modal').classList.add('hidden');
            
            runQueryAndUpdateUI();

        } catch(e) {
            alert(`Erro ao aplicar a fórmula: ${e.message}`);
            console.error(e);
        }
    }

    function downloadChart(id) {
        const chart = chartInstances[id];
        if (!chart) { alert('Gere um gráfico para poder baixá-lo.'); return; }
        const link = document.createElement('a');
        link.href = chart.toBase64Image('image/png', 1.0);
        link.download = `power-graphx-chart-${id}.png`;
        link.click();
    }

    function applyFilter(closeModal = true) {
        const column = document.getElementById('filter-column').value;
        const condition = document.getElementById('filter-condition').value;
        const value = document.getElementById('filter-value').value;
        const query = document.getElementById('sql-editor').value;

        if (!column) {
            runQueryAndUpdateUI();
        } else {
             let filterClause = ` WHERE \`${column}\` `;
             switch (condition) {
                 case 'contains': filterClause += `LIKE '%${value}%'`; break;
                 case 'not_contains': filterClause += `NOT LIKE '%${value}%'`; break;
                 case 'equals': filterClause += `= '${value}'`; break;
                 case 'not_equals': filterClause += `!= '${value}'`; break;
                 default:
                     if (isNaN(value) || value.trim() === '') {
                         filterClause += `${condition} '${value}'`;
                     } else {
                         filterClause += `${condition} ${value}`;
                     }
             }
             try {
                 const filteredQuery = `SELECT * FROM (${query}) AS subquery${filterClause}`;
                 currentData = alasql(filteredQuery);
             } catch(e) {
                 console.error("Não foi possível aplicar filtro na query atual, aplicando na fonte original.", e);
                 currentData = alasql(`SELECT * FROM ${initialTableName}${filterClause}`);
             }
        }
        
        renderTable();
        updateStatus();
        if(closeModal) document.getElementById('filter-modal').classList.add('hidden');
    }

    function downloadCSV() {
        if (currentData.length === 0) return;
        const headers = columnStructure.map(c => c.displayName);
        const rows = currentData.map(row => columnStructure.map(col => {
            let cellString = String(row[col.originalName] ?? '');
            cellString = cellString.replace(/\r?\n/g, ' '); 
            if (cellString.includes(',') || cellString.includes('"')) {
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
    
    function saveState() {
        const state = {
            version: '6.6.0',
            alasqlTables: {},
            conditionalFormatting: conditionalFormattingRules,
            sqlEditorContent: document.getElementById('sql-editor').value,
            charts: []
        };
        
        Object.keys(alasql.tables).forEach(name => {
            state.alasqlTables[name] = alasql.tables[name].data;
        });
        
        document.querySelectorAll('.chart-analysis-section').forEach(section => {
            const id = section.dataset.id;
            const chartConfig = {
                id: id,
                mainTitle: section.querySelector('.chart-title-input-main').value,
                dataSource: section.querySelector('.chart-data-source').value,
                chartType: section.querySelector(`input[name="chart-type-${id}"]:checked`).value,
                groupBy: section.querySelector('select[name="group-by"]').value,
                subtitle: section.querySelector('.chart-subtitle-input').value,
                showLabels: section.querySelector('.show-labels').checked,
                labelPosition: section.querySelector('.label-position').value,
                labelSize: section.querySelector('.label-size').value,
                showGrid: section.querySelector('.show-grid').checked,
                yAxisAuto: section.querySelector('.y-axis-auto').checked,
                yAxisMax: section.querySelector('.y-axis-max').value,
                barBorderRadius: section.querySelector('.bar-border-radius').value,
                lineInterpolation: section.querySelector('.line-interpolation').value,
                pointStyle: section.querySelector('.point-style').value,
                pointSize: section.querySelector('.point-size').value,
                series: []
            };

            section.querySelectorAll('.series-control').forEach(sc => {
                chartConfig.series.push({
                    xAxis: sc.querySelector('[name="x-axis"]').value,
                    yAxis: sc.querySelector('[name="y-axis"]').value,
                    yAxisStart: sc.querySelector('[name="y-axis-start"]').value,
                    aggregation: sc.querySelector('[name="aggregation"]').value,
                    label: sc.querySelector('[name="series-label"]').value,
                    color: sc.querySelector('[name="color"]').value,
                    seriesType: sc.querySelector('[name="series-type"]').value,
                    secondaryAxis: sc.querySelector('[name="secondary-axis"]').checked
                });
            });
            state.charts.push(chartConfig);
        });
        
        const stateJson = JSON.stringify(state);
        const newDoc = document.cloneNode(true);
        const stateScript = newDoc.createElement('script');
        stateScript.id = 'savedState';
        stateScript.type = 'application/json';
        stateScript.textContent = stateJson;
        newDoc.head.appendChild(stateScript);
        
        const oldJsonData = newDoc.getElementById('jsonData');
        if (oldJsonData) oldJsonData.remove();

        const blob = new Blob([newDoc.documentElement.outerHTML], { type: 'text/html' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'PowerGraphx_BI_Sessao.html';
        link.click();
    }

    function loadStateAndRestoreUI() {
        const stateScript = document.getElementById('savedState');
        if (!stateScript) return;
        
        const state = JSON.parse(stateScript.textContent);
        
        Object.keys(alasql.tables).forEach(name => delete alasql.tables[name]);
        Object.keys(state.alasqlTables || {}).forEach(name => {
            alasql.tables[name] = { data: state.alasqlTables[name] };
        });
        
        originalData = alasql.tables[initialTableName]?.data || [];
        document.getElementById('sql-editor').value = state.sqlEditorContent;
        conditionalFormattingRules = state.conditionalFormatting || [];
        renderCondFormatRules();
        
        document.getElementById('charts-container').innerHTML = '';
        chartInstances = {};
        chartAnalysisCounter = 0;
        (state.charts || []).forEach(chartConfig => {
            addChartAnalysis();
            restoreChartUI(chartAnalysisCounter, chartConfig);
        });

        runQueryAndUpdateUI();
        updateChartNav();
    }
    
    function restoreChartUI(id, config) {
        const section = document.getElementById(`chart-section-${id}`);
        if (!section) return;

        section.querySelector('.chart-title-input-main').value = config.mainTitle || `Análise Gráfica ${id}`;
        section.querySelector('.chart-data-source').value = config.dataSource;
        section.querySelector(`input[name="chart-type-${id}"][value="${config.chartType}"]`).checked = true;
        section.querySelector('select[name="group-by"]').value = config.groupBy;
        section.querySelector('.chart-subtitle-input').value = config.subtitle;
        section.querySelector('.show-labels').checked = config.showLabels;
        section.querySelector('.label-position').value = config.labelPosition;
        section.querySelector('.label-size').value = config.labelSize;
        section.querySelector('.show-grid').checked = config.showGrid;
        section.querySelector('.y-axis-auto').checked = config.yAxisAuto;
        section.querySelector('.y-axis-max').value = config.yAxisMax;
        section.querySelector('.y-axis-max').disabled = config.yAxisAuto;
        section.querySelector('.bar-border-radius').value = config.barBorderRadius;
        section.querySelector('.line-interpolation').value = config.lineInterpolation;
        section.querySelector('.point-style').value = config.pointStyle;
        section.querySelector('.point-size').value = config.pointSize;

        const seriesContainer = section.querySelector(`#series-container-${id}`);
        seriesContainer.innerHTML = '';
        
        (config.series || []).forEach((s, index) => {
            addSeriesControl(id, index === 0);
            const newSeriesControl = seriesContainer.lastElementChild;
            newSeriesControl.querySelector('[name="x-axis"]').value = s.xAxis;
            newSeriesControl.querySelector('[name="y-axis"]').value = s.yAxis;
            newSeriesControl.querySelector('[name="y-axis-start"]').value = s.yAxisStart;
            newSeriesControl.querySelector('[name="aggregation"]').value = s.aggregation;
            newSeriesControl.querySelector('[name="series-label"]').value = s.label;
            newSeriesControl.querySelector('[name="color"]').value = s.color;
            newSeriesControl.querySelector('[name="series-type"]').value = s.seriesType;
            newSeriesControl.querySelector('[name="secondary-axis"]').checked = s.secondaryAxis;
        });
        
        section.querySelector(`input[name="chart-type-${id}"]:checked`).dispatchEvent(new Event('change', { bubbles: true }));
        section.querySelector('select[name="group-by"]').dispatchEvent(new Event('change', { bubbles: true }));
    }

    function updateChartNav() {
        const nav = document.getElementById('chart-nav');
        const container = nav.querySelector('ul');
        container.innerHTML = '';
        const sections = document.querySelectorAll('.chart-analysis-section');
        
        if (sections.length < 2) {
            nav.classList.add('hidden');
            return;
        }
        
        nav.classList.remove('hidden');
        sections.forEach(section => {
            const id = section.dataset.id;
            const title = section.querySelector('.chart-title-input-main').value || `Análise ${id}`;
            const li = document.createElement('li');
            li.innerHTML = `<a href="#chart-section-${id}" data-nav-id="${id}" class="block p-2 text-sm text-gray-700 hover:bg-gray-200 rounded-md truncate transition-colors">${title}</a>`;
            li.querySelector('a').addEventListener('click', (e) => {
                e.preventDefault();
                section.scrollIntoView({ behavior: 'smooth' });
            });
            container.appendChild(li);
        });
    }
'@

    $template = @'
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power-Graphx BI</title>
    <style>
        html { scroll-behavior: smooth; }
        .modal { transition: opacity 0.25s ease; }
        #table-container { max-height: calc(100vh - 250px); overflow: auto; }
        table thead { position: sticky; top: 0; z-index: 10; }
        .chart-selector input:checked + label { border-color: #3b82f6; background-color: #eff6ff; box-shadow: 0 0 0 2px #3b82f6; }
        .divider { border-top: 1px solid #e5e7eb; margin: 1rem 0; }
        .chart-title-input-main { background: transparent; border: 1px solid transparent; font-size: 1.5rem; font-weight: bold; color: #1f2937; width: 100%; padding: 2px 5px; border-radius: 4px; transition: all 0.2s ease-in-out; }
        .chart-title-input-main:hover { border: 1px solid #d1d5db; }
        .chart-title-input-main:focus { outline: none; box-shadow: 0 0 0 2px #3b82f6; border-color: #3b82f6; background: white; }
    </style>
    {{CDN_TAGS}}
</head>
<body class="bg-gray-100 font-sans">
    <script id="jsonData" type="application/json">{{JSON_DATA}}</script>
    <input type="file" id="csv-upload-input" multiple accept=".csv" class="hidden">
    <input type="file" id="csv-update-input" accept=".csv" class="hidden">

    <nav id="chart-nav" class="hidden fixed top-1/4 left-4 bg-white/80 backdrop-blur-sm shadow-lg rounded-lg p-2 w-48 z-20">
        <h4 class="font-bold text-sm text-gray-800 mb-2 px-2">Gráficos</h4>
        <ul class="space-y-1"></ul>
    </nav>

    <header class="bg-white shadow-md p-4 sticky top-0 z-20">
        <div class="container mx-auto">
            <div class="flex flex-wrap justify-between items-center gap-4">
                <h1 class="text-2xl font-bold text-gray-800">Power-Graphx BI</h1>
                <div class="flex items-center space-x-2 flex-wrap gap-y-2">
                    <button id="btn-update-source" class="px-4 py-2 text-sm font-medium text-white bg-cyan-600 rounded-md hover:bg-cyan-700">Atualizar Dados</button>
                    <button id="btn-add-csv" class="px-4 py-2 text-sm font-medium text-white bg-orange-600 rounded-md hover:bg-orange-700">Adicionar CSV</button>
                    <button id="btn-new-column" class="px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700">Nova Coluna</button>
                    <button id="btn-edit-data" class="px-4 py-2 text-sm font-medium text-white bg-yellow-600 rounded-md hover:bg-yellow-700">Editar Dados</button>
                    <button id="btn-cond-format" class="px-4 py-2 text-sm font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700">Formatar Tabela</button>
                    <button id="btn-filter" class="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700">Filtrar</button>
                    <button id="btn-toggle-sql" class="px-4 py-2 text-sm font-medium text-white bg-teal-600 rounded-md hover:bg-teal-700">Console SQL</button>
                    <button id="btn-add-chart" class="px-4 py-2 text-sm font-medium text-white bg-purple-600 rounded-md hover:bg-purple-700">Adicionar Gráfico</button>
                    <button id="btn-save-state" class="px-4 py-2 text-sm font-medium text-white bg-pink-600 rounded-md hover:bg-pink-700">Salvar Sessão</button>
                    <button id="btn-download-csv" class="px-4 py-2 text-sm font-medium text-white bg-gray-800 rounded-md hover:bg-gray-900">Baixar CSV</button>
                </div>
            </div>
            <div class="text-xs text-gray-500 mt-1" id="status-label">Carregando...</div>
        </div>
    </header>

    <main class="container mx-auto p-4">
        <section id="sql-section" class="hidden mb-6 bg-white rounded-lg shadow p-6 grid grid-cols-1 md:grid-cols-4 gap-6">
            <div class="md:col-span-3">
                <h2 class="text-2xl font-bold text-gray-800 mb-2">Console SQL (AlaSQL)</h2>
                <textarea id="sql-editor" class="w-full h-32 p-2 font-mono text-sm border border-gray-300 rounded-md" placeholder="SELECT * FROM source_data;">SELECT * FROM source_data;</textarea>
            </div>
            <div class="md:col-span-1">
                <h3 class="text-lg font-bold text-gray-700 mb-2">Tabelas Carregadas</h3>
                <div class="bg-gray-50 p-3 rounded-md h-32 overflow-y-auto">
                    <ul id="table-list" class="list-disc list-inside text-sm font-mono text-gray-800"></ul>
                </div>
            </div>
            <div class="md:col-span-4 mt-2 flex justify-between items-center">
                <div id="sql-status" class="text-sm text-gray-500 italic"></div>
                <div class="flex-shrink-0 flex gap-2">
                    <button id="btn-reset-data" class="px-4 py-2 text-sm font-medium text-gray-700 bg-gray-200 rounded-md hover:bg-gray-300">Resetar Dados</button>
                    <button id="btn-format-sql" class="px-4 py-2 text-sm font-medium text-white bg-indigo-500 rounded-md hover:bg-indigo-600">Formatar SQL</button>
                    <button id="btn-run-sql" class="px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700">Executar</button>
                </div>
            </div>
        </section>
        
        <div id="table-container" class="bg-white rounded-lg shadow overflow-hidden mb-6"></div>
        <div id="charts-container" class="space-y-6"></div>
    </main>
    
    <template id="chart-analysis-template">
        <section id="chart-section-__ID__" class="chart-analysis-section bg-white rounded-lg shadow" data-id="__ID__">
             <div class="p-6">
                  <div class="flex justify-between items-start mb-4">
                      <input type="text" value="Análise Gráfica __ID__" class="chart-title-input-main">
                      <button class="remove-chart-btn text-red-500 hover:text-red-700 font-bold text-2xl leading-none ml-4">&times;</button>
                  </div>
                  <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">
                        <div class="lg:col-span-1 flex flex-col space-y-4">
                            <div>
                                <h3 class="font-bold text-gray-700 mb-2">1. Fonte e Tipo</h3>
                                <select id="chart-data-source-__ID__" class="chart-data-source block w-full rounded-md border-gray-300 text-sm mb-2"></select>
                                <div class="chart-selector grid grid-cols-3 gap-2">
                                     <div><input type="radio" name="chart-type-__ID__" value="bar" id="type-bar-__ID__" checked class="hidden"><label for="type-bar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Barra</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="horizontalBar" id="type-hbar-__ID__" class="hidden"><label for="type-hbar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Barra Horiz.</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="line" id="type-line-__ID__" class="hidden"><label for="type-line-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Linha</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="area" id="type-area-__ID__" class="hidden"><label for="type-area-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Área</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="stackedBar" id="type-stackedBar-__ID__" class="hidden"><label for="type-stackedBar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Empilhada</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="groupedStackedBar" id="type-gstackedBar-__ID__" class="hidden"><label for="type-gstackedBar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Agrup/Empilh.</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="floatingBar" id="type-floatingBar-__ID__" class="hidden"><label for="type-floatingBar-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Flutuante</label></div>
                                     <div><input type="radio" name="chart-type-__ID__" value="combo" id="type-combo-__ID__" class="hidden"><label for="type-combo-__ID__" class="p-2 border rounded-md cursor-pointer flex justify-center items-center text-xs">Combo</label></div>
                                </div>
                            </div>
                            <div>
                                <h3 class="font-bold text-gray-700 mb-2">2. Agrupar Por (Opcional)</h3>
                                <select name="group-by" class="group-by-select mt-1 block w-full rounded-md border-gray-300 text-sm"><option value="">-- Nenhum --</option></select>
                            </div>
                           <div>
                                <div class="flex justify-between items-center mb-2"><h3 class="font-bold text-gray-700">3. Séries de Dados</h3><button class="add-series-btn text-xs bg-blue-500 text-white py-1 px-2 rounded-full hover:bg-blue-600">+ Série</button></div>
                                <div id="series-container-__ID__" class="space-y-3 max-h-60 overflow-y-auto"></div>
                           </div>
                        </div>
                        <div class="lg:col-span-2 bg-gray-50 rounded-lg p-4 flex items-center justify-center min-h-[400px]">
                         <div class="relative w-full h-full"><canvas id="mainChart-__ID__"></canvas></div>
                        </div>
                        <div class="lg:col-span-1 flex flex-col space-y-2 text-sm">
                            <h3 class="font-bold text-gray-700 mb-2">4. Formatar Visual</h3>
                            <div><span class="font-semibold text-gray-700">Títulos</span>
                                <div class="mt-2 space-y-2">
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
                            <div class="point-styling-options" style="display:none;">
                                <div class="divider"></div>
                                <div><span class="font-semibold text-gray-700">Estilo dos Pontos (Linha/Área)</span>
                                    <div class="mt-2"><label class="text-xs text-gray-600">Estilo:</label><select class="point-style mt-1 block w-full rounded-md border-gray-300 text-xs"><option value="circle">Círculo</option><option value="rect">Quadrado</option><option value="rectRot">Diamante</option><option value="star">Estrela</option><option value="triangle">Triângulo</option></select></div>
                                    <div class="mt-2"><label class="text-xs text-gray-600">Tamanho:</label><input type="number" value="3" min="0" class="point-size mt-1 block w-full rounded-md border-gray-300 text-xs"></div>
                                </div>
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
            <div class="flex justify-between items-center"><h3 class="text-lg font-medium">Filtrar Dados</h3><button class="modal-close font-bold text-xl">&times;</button></div>
            <div class="mt-4 space-y-4">
                <div><label class="block text-sm font-medium">Coluna</label><select id="filter-column" class="mt-1 w-full border-gray-300 rounded-md"></select></div>
                <div><label class="block text-sm font-medium">Condição</label><select id="filter-condition" class="mt-1 w-full border-gray-300 rounded-md"><option value="equals">Igual a</option><option value="not_equals">Diferente de</option><option value=">">Maior que</option><option value="<">Menor que</option><option value=">=">Maior ou Igual</option><option value="<=">Menor ou Igual</option><option value="contains">Contém</option><option value="not_contains">Não Contém</option></select></div>
                <div><label class="block text-sm font-medium">Valor</label><input type="text" id="filter-value" class="mt-1 w-full border-gray-300 rounded-md"></div>
            </div>
            <div class="mt-6 flex justify-end"><button id="apply-filter-btn" class="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700">Aplicar Filtro</button></div>
        </div>
    </div>

    <div id="calc-column-modal" class="modal hidden fixed inset-0 bg-gray-600 bg-opacity-50 overflow-y-auto h-full w-full z-30">
        <div class="relative top-20 mx-auto p-5 border w-96 shadow-lg rounded-md bg-white">
            <div class="flex justify-between items-center"><h3 class="text-lg font-medium text-gray-900">Nova Coluna Calculada</h3><button class="modal-close font-bold text-xl">&times;</button></div>
            <div class="mt-4 space-y-4">
                <div><label class="block text-sm font-medium">Tabela de Origem</label><select id="calc-column-table" class="mt-1 block w-full border-gray-300 rounded-md"></select></div>
                <div><label class="block text-sm font-medium">Nome da Nova Coluna</label><input type="text" id="calc-column-name" class="mt-1 block w-full border-gray-300 rounded-md"></div>
                <div><label class="block text-sm font-medium">Fórmula</label><input type="text" id="calc-column-formula" class="mt-1 block w-full border-gray-300 rounded-md" placeholder="Ex: [Receita] - [Custo]"></div>
                <div class="p-3 bg-gray-50 rounded-md text-xs text-gray-600">
                    <h4 class="font-bold mb-1">Como usar:</h4>
                    <p>Use colchetes `[]` para os nomes das colunas. Operadores: +, -, *, /.</p>
                    <ul class="list-disc list-inside mt-2">
                        <li><b>Exemplo 1:</b> `[Receita] - [Custo]`</li>
                        <li><b>Exemplo 2:</b> `([Nota1] + [Nota2]) / 2`</li>
                    </ul>
                </div>
            </div>
            <div class="mt-6 flex justify-end"><button id="apply-calc-column-btn" class="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700">Criar Coluna</button></div>
        </div>
    </div>
    
    <div id="cond-format-modal" class="modal hidden fixed inset-0 bg-gray-600 bg-opacity-50 overflow-y-auto h-full w-full z-30">
        <div class="relative top-20 mx-auto p-5 border w-[500px] shadow-lg rounded-md bg-white">
            <div class="flex justify-between items-center"><h3 class="text-lg font-medium">Formatação Condicional</h3><button class="modal-close font-bold text-xl">&times;</button></div>
            <div class="mt-4">
                <h4 class="font-semibold mb-2">Nova Regra</h4>
                <div class="p-3 border rounded-md">
                    <div class="grid grid-cols-4 gap-2 items-end">
                        <div><label class="block text-xs font-medium">Coluna</label><select id="cond-format-column" class="mt-1 w-full text-sm border-gray-300 rounded-md"></select></div>
                        <div><label class="block text-xs font-medium">Condição</label><select id="cond-format-condition" class="mt-1 w-full text-sm border-gray-300 rounded-md"><option value="greater">Maior que</option><option value="less">Menor que</option><option value="equals">Igual a</option><option value="not_equals">Diferente de</option><option value="contains">Contém</option><option value="not_contains">Não Contém</option></select></div>
                        <div><label class="block text-xs font-medium">Valor</label><input type="text" id="cond-format-value" class="mt-1 w-full text-sm border-gray-300 rounded-md"></div>
                        <div><label class="block text-xs font-medium">Cor</label><input type="color" id="cond-format-color" value="#ef4444" class="mt-1 w-full h-9 p-0 border-0 rounded-md"></div>
                    </div>
                    <div class="mt-3">
                        <label class="block text-xs font-medium mb-1">Aplicar a:</label>
                        <div class="flex items-center space-x-4 text-sm">
                            <label class="flex items-center"><input type="radio" name="cond-format-apply-to" value="cell" checked class="mr-1"> Célula</label>
                            <label class="flex items-center"><input type="radio" name="cond-format-apply-to" value="row" class="mr-1"> Linha Inteira</label>
                        </div>
                    </div>
                </div>
                <button id="add-cond-format-rule-btn" class="mt-3 w-full py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 text-sm">Adicionar Regra</button>
            </div>
            <div class="mt-6"><h4 class="font-semibold mb-2">Regras Ativas</h4><div id="cond-format-rules-list" class="space-y-2 max-h-48 overflow-y-auto"></div></div>
        </div>
    </div>
    
    <script>
        {{JS_CODE}}
    </script>
</body>
</html>
'@

    $template = $template -replace '\{\{CDN_TAGS\}\}', $CdnLibraryTags
    $template = $template -replace '\{\{JSON_DATA\}\}', $JsonData
    $template = $template -replace '\{\{JS_CODE\}\}', $ApplicationJavaScript

    return $template
}

# --- 4. Função Principal de Execução ---
Function Start-WebApp {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV, Salvos HTML (*.csv, *.html)|*.csv;*.html|Todos os arquivos (*.*)|*.*"
    $OpenFileDialog.Title = "Power-Graphx: Selecione o arquivo CSV inicial ou uma sessão salva"

    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FilePath = $OpenFileDialog.FileName
        
        if ($FilePath.EndsWith(".html")) {
            Write-Host "Abrindo sessão salva do Power-Graphx..." -ForegroundColor Cyan
            Start-Process $FilePath
            return
        }
        
        try {
            $firstLine = Get-Content -Path $FilePath -TotalCount 1 -Encoding Default
            
            # CORREÇÃO: Lógica de detecção de delimitador com if/elseif (sugerida por você!)
            $semicolonCount = ([regex]::Matches($firstLine, ';')).Count
            $commaCount = ([regex]::Matches($firstLine, ',')).Count
            $tabCount = ([regex]::Matches($firstLine, "`t")).Count
            
            $bestDelimiter = ','  # Padrão para vírgula
            if ($semicolonCount -gt $commaCount -and $semicolonCount -gt $tabCount) {
                $bestDelimiter = ';'
            }
            elseif ($tabCount -gt $commaCount -and $tabCount -gt $semicolonCount) {
                $bestDelimiter = "`t"
            }
            
            Write-Host "Delimitador detectado: '$bestDelimiter' (';': $semicolonCount | ',': $commaCount | TAB: $tabCount)" -ForegroundColor Yellow
            
            $Data = Import-Csv -Path $FilePath -Delimiter $bestDelimiter -Encoding Default
            if ($null -eq $Data -or $Data.Count -eq 0) { throw "O arquivo CSV está vazio ou em um formato inválido." }
            
            Write-Host "Dados carregados com sucesso! ($($Data.Count) linhas)" -ForegroundColor Cyan
            
            $JsonData = $Data | ConvertTo-Json -Compress -Depth 10
            $cdnTags = Get-CdnLibraryTags
            $HtmlContent = Get-HtmlTemplate -JsonData $JsonData -CdnLibraryTags $cdnTags
            
            $OutputPath = Join-Path $env:TEMP "PowerGraphx_BI_WebApp.html"
            $HtmlContent | Out-File -FilePath $OutputPath -Encoding UTF8
            
            Write-Host "Aplicação gerada: $OutputPath" -ForegroundColor Green
            Start-Process $OutputPath
            
        } catch {
            $errorMsg = "Erro ao processar arquivo:`n`n$($_.Exception.Message)"
            Write-Host $errorMsg -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Erro", "OK", "Error")
        }
    }
}


# --- 5. Iniciar a Aplicação ---
Start-WebApp
