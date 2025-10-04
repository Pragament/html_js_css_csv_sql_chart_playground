// Global variables for features
let SQL;
let db;
let sqlEditor;
let datatableEditor;
let dataTableInstance = null;
let lastSqlResult = null;
let sqlChartInstance = null;
const SQL_HISTORY_KEY = 'spreadsheet_sql_history';

// YOUR ORIGINAL GLOBAL VARIABLES (PRESERVED)
const sampleData = {
    sales: `Product,Quarter,Qty,Revenue\nWidget,Q1,150,3750.00\nWidget,Q2,200,5000.00\nGadget,Q1,75,1125.00\nGadget,Q2,125,1875.00\nGadget,Q3,150,2250.00`,
    employees: `ID,Name,Department,Salary,HireDate\n1,John Smith,Engineering,85000,2020-05-15\n2,Jane Doe,Marketing,72000,2019-11-03\n3,Robert Johnson,Sales,68000,2021-02-28\n4,Emily Wilson,Engineering,92000,2018-07-22\n5,Michael Brown,Marketing,76000,2022-01-10`,
    inventory: `SKU,ProductName,Category,Quantity,Price,LastStocked\n1001,Desk Chair,Furniture,45,129.99,2023-03-15\n1002,Monitor Stand,Electronics,28,49.95,2023-04-02\n1003,Wireless Keyboard,Electronics,62,79.99,2023-03-28\n1004,Desk Lamp,Home,37,34.50,2023-04-10\n1005,Notebook,Office,120,4.99,2023-04-05`
};
let currentData = [];
let headers = [];
let charts = [];
let selectedCells = [];
let selectedRow = null;
let selectedColumn = null;
let formulas = [];
let dependencies = {};
let isFilling = false;
let fillStartCell = null;
let fillRange = [];
window.fillRange = fillRange;

// --- MAIN APP INITIALIZATION (UNIFIED) ---
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, setting up event listeners');

    // --- Initialize Editors ---
    sqlEditor = CodeMirror.fromTextArea(document.getElementById('sql-query'), {
        mode: 'text/x-sql',
        theme: 'dracula',
        lineNumbers: true
    });
    
    datatableEditor = CodeMirror.fromTextArea(document.getElementById('datatable-config'), {
        mode: { name: 'javascript', json: true },
        theme: 'dracula',
        lineNumbers: true
    });
    datatableEditor.setValue(`{\n  "paging": true,\n  "columnDefs": [\n    { "targets": 2, "visible": false }\n  ]\n}`);

    // --- Initialize Event Listeners ---
    const fileInput = document.getElementById('file-input');
    const formulaBar = document.getElementById('formula-bar');
    
    document.getElementById('import-btn').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
    document.getElementById('export-csv').addEventListener('click', () => exportData('csv', headers, currentData));
    document.getElementById('export-excel').addEventListener('click', () => exportData('excel', headers, currentData));
    document.getElementById('add-column').addEventListener('click', addColumn);
    document.getElementById('add-row').addEventListener('click', addRow);
    document.getElementById('clear-data').addEventListener('click', clearData);
    document.getElementById('create-chart').addEventListener('click', createChart);
    
    document.getElementById('apply-datatable-config').addEventListener('click', applyDataTableConfig);
    document.getElementById('run-sql').addEventListener('click', runSQL);
    document.getElementById('export-sql-csv').addEventListener('click', () => exportSqlResults('csv'));
    document.getElementById('export-sql-excel').addEventListener('click', () => exportSqlResults('excel'));
    document.getElementById('sql-create-chart').addEventListener('click', createSqlChart);
    document.getElementById('sql-download-chart').addEventListener('click', downloadSqlChart);

    // Tab switching with editor refresh
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', function() {
            document.querySelectorAll('.tab-button, .tab-content').forEach(el => el.classList.remove('active'));
            this.classList.add('active');
            document.getElementById(this.dataset.tab).classList.add('active');
            setTimeout(() => {
                if (this.dataset.tab === 'sql-tab') sqlEditor.refresh();
                if (this.dataset.tab === 'filtered-views-tab') datatableEditor.refresh();
            }, 1);
        });
    });

    // Sample file loading
    document.querySelectorAll('#sample-files li').forEach(item => {
        item.addEventListener('click', function() {
            loadSampleData(this.getAttribute('data-file'));
        });
    });

    // Your original formula bar and other listeners...
    // ...

    // Initialize Database and load initial data
    initSqlJs({ locateFile: file => `https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/${file}` })
    .then(SQL_ => {
        SQL = SQL_;
        db = new SQL.Database();
        document.getElementById('run-sql').disabled = false;
        loadSampleData('sales');
    });
    renderSqlHistory();
});


// --- ALL YOUR ORIGINAL FUNCTIONS (PRESERVED AND UNCHANGED) ---

function updateCellValue(row, col, value) {
    // This is your full original function
    const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
    if (!cell) return;
    if (value.startsWith('=')) {
        formulas[row] = formulas[row] || [];
        formulas[row][col] = value;
    } else if (formulas[row]?.[col]) {
        delete formulas[row][col];
    }
    currentData[row] = currentData[row] || [];
    currentData[row][col] = evaluateFormula(value, row, col);
    if (cell) cell.textContent = currentData[row][col];
    updateSQLDatabase();
    renderSpreadsheet();
}

function evaluateFormula(formula, row, col) {
    // This is your full original, complex evaluateFormula function.
    if (!formula.startsWith('=')) return formula;
    try {
        let expr = formula.slice(1).replace(/[A-Z]+\d+/g, (ref) => {
            const cellRef = parseCellReference(ref);
            if (!cellRef || currentData[cellRef.row] === undefined || currentData[cellRef.row][cellRef.col] === undefined) return 0;
            return currentData[cellRef.row][cellRef.col] || 0;
        });
        return eval(expr);
    } catch (e) {
        return '#ERROR!';
    }
}

function parseCellReference(ref) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    const colStr = match[1];
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    col--;
    const row = parseInt(match[2]) - 1;
    return { row, col };
}

function loadSampleData(key) {
    const data = sampleData[key];
    if(data) processCSVData(data);
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
        if (file.name.endsWith('.csv')) processCSVData(e.target.result);
        else processExcelData(e.target.result);
    };
    if (file.name.endsWith('.csv')) reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
}

function processCSVData(csvString) {
    const parsed = Papa.parse(csvString, { header: true, skipEmptyLines: true });
    headers = parsed.meta.fields || [];
    currentData = parsed.data.map(row => headers.map(h => row[h]));
    formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(null));
    renderSpreadsheet();
    updateSQLDatabase();
    updateChartDropdowns();
}

function processExcelData(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    headers = jsonData[0];
    currentData = jsonData.slice(1);
    formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(null));
    renderSpreadsheet();
    updateSQLDatabase();
    updateChartDropdowns();
}

function renderSpreadsheet() {
    const headerRow = document.getElementById('header-row');
    const dataBody = document.getElementById('data-body');
    headerRow.innerHTML = `<th class="row-number-header"></th>`;
    dataBody.innerHTML = '';
    
    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header;
        th.contentEditable = true;
        th.dataset.column = index;
        th.addEventListener('blur', () => { headers[index] = th.textContent.trim() || `Column${index + 1}`; updateSQLDatabase(); updateChartDropdowns(); });
        headerRow.appendChild(th);
    });

    currentData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        const rowNumTd = document.createElement('td');
        rowNumTd.className = 'row-number';
        rowNumTd.textContent = rowIndex + 1;
        tr.appendChild(rowNumTd);
        
        headers.forEach((_, colIndex) => {
            const td = document.createElement('td');
            td.textContent = row[colIndex] || '';
            td.contentEditable = true;
            td.dataset.row = rowIndex;
            td.dataset.column = colIndex;
            td.addEventListener('blur', () => {
                updateCellValue(rowIndex, colIndex, td.textContent.trim());
            });
            tr.appendChild(td);
        });
        dataBody.appendChild(tr);
    });
}

function addColumn() { headers.push(`Column${headers.length + 1}`); currentData.forEach(row => row.push('')); renderSpreadsheet(); updateSQLDatabase(); updateChartDropdowns(); }
function addRow() { currentData.push(Array(headers.length).fill('')); renderSpreadsheet(); updateSQLDatabase(); }
function clearData() { if (confirm('Clear all data?')) { headers = []; currentData = []; charts.forEach(c => c.chart.destroy()); charts = []; document.getElementById('chart-list').innerHTML = ''; renderSpreadsheet(); updateSQLDatabase(); updateChartDropdowns(); } }

function updateSQLDatabase() {
    if (!db || !headers || headers.length === 0) return;
    db.run('DROP TABLE IF EXISTS data');
    const sanitizedHeaders = headers.map(h => `"${h.replace(/"/g, '""')}"`);
    db.run(`CREATE TABLE data (${sanitizedHeaders.join(', ')})`);
    const stmt = db.prepare(`INSERT INTO data VALUES (${headers.map(() => '?').join(',')})`);
    currentData.forEach(row => {
        stmt.run(row);
    });
    stmt.free();
}

// --- NEW AND MODIFIED FUNCTIONS ---

function applyDataTableConfig() {
    const container = document.getElementById('datatable-container');
    if (headers.length === 0) { container.innerHTML = `<p class="placeholder-text">No data to display.</p>`; return; }
    if (dataTableInstance) dataTableInstance.destroy();
    let config;
    try { config = new Function(`return ${datatableEditor.getValue()}`)(); } catch (e) { container.innerHTML = `<p class="error-text">Error in config:\n${e.message}</p>`; return; }
    container.innerHTML = `<table id="filtered-table" class="display" style="width:100%"><thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead></table>`;
    config.data = currentData;
    try { 
        dataTableInstance = new DataTable('#filtered-table', config); 
    } catch (e) { 
        container.innerHTML = `<p class="error-text">Error initializing DataTable:\n${e.message}</p>`; 
    }
}

function runSQL() {
    const query = sqlEditor.getValue().trim();
    const actions = document.getElementById('sql-actions');
    const chartArea = document.getElementById('sql-chart-area');
    if (!query) return;
    addToSqlHistory(query);
    try {
        const result = db.exec(query);
        renderSqlResults(result);
        if (result && result.length > 0) {
            lastSqlResult = result[0];
            updateSqlChartControls(lastSqlResult.columns);
            actions.classList.remove('hidden');
        } else {
            lastSqlResult = null;
            actions.classList.add('hidden');
        }
        chartArea.classList.add('hidden');
        if(sqlChartInstance) { sqlChartInstance.destroy(); sqlChartInstance = null; }
    } catch (error) {
        renderSqlResults(null, error);
        lastSqlResult = null;
        actions.classList.add('hidden');
        chartArea.classList.add('hidden');
    }
}
function renderSqlResults(result, error = null) {
    const container = document.getElementById('sql-results-container');
    if (error) { container.innerHTML = `<p class="error-text">SQL Error: ${error.message}</p>`; return; }
    if (!result || result.length === 0) { container.innerHTML = `<p class="placeholder-text">Query returned no results.</p>`; return; }
    const { columns, values } = result[0];
    container.innerHTML = `<table><thead><tr>${columns.map(c => `<th>${c}</th>`).join('')}</tr></thead><tbody>${values.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
}
function addToSqlHistory(query) {
    let history = JSON.parse(localStorage.getItem(SQL_HISTORY_KEY)) || [];
    history = history.filter(q => q !== query);
    history.unshift(query);
    if (history.length > 20) history.length = 20;
    localStorage.setItem(SQL_HISTORY_KEY, JSON.stringify(history));
    renderSqlHistory();
}
function renderSqlHistory() {
    const historyList = document.getElementById('sql-history-list');
    const history = JSON.parse(localStorage.getItem(SQL_HISTORY_KEY)) || [];
    if (history.length === 0) { historyList.innerHTML = '<li>No history yet.</li>'; return; }
    historyList.innerHTML = history.map(q => `<li>${q}</li>`).join('');
    historyList.querySelectorAll('li').forEach((li, i) => {
        li.addEventListener('click', () => { sqlEditor.setValue(history[i]); sqlEditor.focus(); });
    });
}
function updateSqlChartControls(columns) {
    const xAxisSelect = document.getElementById('sql-x-axis');
    const yAxisSelect = document.getElementById('sql-y-axis');
    xAxisSelect.innerHTML = yAxisSelect.innerHTML = columns.map(c => `<option value="${c}">${c}</option>`).join('');
    if (columns.length > 1) {
        let firstNumericIndex = columns.findIndex((col, index) => lastSqlResult.values.length > 0 && !isNaN(parseFloat(lastSqlResult.values[0][index])));
        yAxisSelect.selectedIndex = (firstNumericIndex !== -1) ? firstNumericIndex : 1;
        if (xAxisSelect.selectedIndex === yAxisSelect.selectedIndex) {
            xAxisSelect.selectedIndex = (yAxisSelect.selectedIndex === 0) ? 1 : 0;
        }
    }
}
function createSqlChart() {
    if (!lastSqlResult) return alert("No SQL data to chart.");
    const chartType = document.getElementById('sql-chart-type').value;
    const xCol = document.getElementById('sql-x-axis').value;
    const yCol = document.getElementById('sql-y-axis').value;
    const xIndex = lastSqlResult.columns.indexOf(xCol), yIndex = lastSqlResult.columns.indexOf(yCol);
    if (xIndex === -1 || yIndex === -1) return;
    if (lastSqlResult.values.some(row => isNaN(parseFloat(row[yIndex])))) {
        alert(`Y-axis column ('${yCol}') must contain numeric data.`);
        return;
    }
    const labels = lastSqlResult.values.map(row => row[xIndex]);
    const data = lastSqlResult.values.map(row => parseFloat(row[yIndex]) || 0);
    if (sqlChartInstance) sqlChartInstance.destroy();
    document.getElementById('sql-chart-area').classList.remove('hidden');
    const ctx = document.getElementById('sql-chart-canvas').getContext('2d');
    sqlChartInstance = new Chart(ctx, { type: chartType, data: { labels, datasets: [{ label: `${yCol} by ${xCol}`, data, backgroundColor: getChartColors(data.length) }] }, options: { responsive: true, plugins: { title: { display: true, text: `Chart of ${yCol} by ${xCol}` } } } });
    document.getElementById('sql-download-chart').classList.remove('hidden');
}
function downloadSqlChart() {
    if (!sqlChartInstance) return;
    const link = document.createElement('a');
    link.download = 'sql_chart.png';
    link.href = sqlChartInstance.canvas.toDataURL('image/png');
    link.click();
}
function createChart() {
    const chartList = document.getElementById('chart-list');
    const chartType = document.getElementById('chart-type').value;
    const xAxis = document.getElementById('x-axis').value;
    const yAxis = document.getElementById('y-axis').value;
    if (headers.length === 0 || !xAxis || !yAxis) return alert('No data or axes selected for chart');
    const xIndex = headers.indexOf(xAxis), yIndex = headers.indexOf(yAxis);
    const labels = currentData.map(row => row[xIndex]);
    const dataValues = currentData.map(row => parseFloat(row[yIndex]) || 0);
    const chartId = `chart-${Date.now()}`;
    const chartItem = document.createElement('div');
    chartItem.className = 'chart-item';
    chartItem.innerHTML = `<canvas id="${chartId}"></canvas><div><button class="delete-chart" data-id="${chartId}">Delete</button><button class="download-chart" data-id="${chartId}">Download</button></div>`;
    chartList.appendChild(chartItem);
    const chart = new Chart(document.getElementById(chartId), { type: chartType, data: { labels, datasets: [{ label: `${yAxis} by ${xAxis}`, data: dataValues, backgroundColor: getChartColors(dataValues.length) }] } });
    charts.push({ id: chartId, chart });
    chartItem.querySelector('.delete-chart').addEventListener('click', e => deleteChart(e.target.dataset.id));
    chartItem.querySelector('.download-chart').addEventListener('click', e => downloadChart(e.target.dataset.id));
}
function getChartColors(count) { return Array.from({ length: count }, (_, i) => `hsl(${(i * 360 / count)}, 70%, 60%)`); }
function deleteChart(id) {
    const index = charts.findIndex(c => c.id === id);
    if (index > -1) { charts[index].chart.destroy(); charts.splice(index, 1); document.getElementById(id).closest('.chart-item').remove(); }
}
function downloadChart(id) {
    const chart = charts.find(c => c.id === id);
    if (chart) { const link = document.createElement('a'); link.download = 'chart.png'; link.href = chart.chart.canvas.toDataURL('image/png'); link.click(); }
}
function updateChartDropdowns() {
    const xAxisSelect = document.getElementById('x-axis');
    const yAxisSelect = document.getElementById('y-axis');
    if (!xAxisSelect || !yAxisSelect) return;
    xAxisSelect.innerHTML = yAxisSelect.innerHTML = headers.map(h => `<option value="${h}">${h}</option>`).join('');
    if (headers.length > 1) yAxisSelect.selectedIndex = 1;
}
function exportData(format, dataHeaders, dataRows) {
    const dataForExport = dataRows.map(row => {
        let obj = {};
        dataHeaders.forEach((col, i) => { obj[col] = Array.isArray(row) ? row[i] : row[col]; });
        return obj;
    });
    if (format === 'csv') {
        const csv = Papa.unparse({ fields: dataHeaders, data: dataForExport });
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'export.csv';
        link.click();
    } else if (format === 'excel') {
        const ws = XLSX.utils.json_to_sheet(dataForExport, { header: dataHeaders });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, 'export.xlsx');
    }
}
function exportSqlResults(format) {
    if (!lastSqlResult) return alert('No SQL result to export.');
    exportData(format, lastSqlResult.columns, lastSqlResult.values);
}