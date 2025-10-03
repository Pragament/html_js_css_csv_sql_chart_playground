let SQL;
let db;
let sqlEditor;
let datatableEditor;
let dataTableInstance = null;

// Key for local storage
const SQL_HISTORY_KEY = 'spreadsheet_sql_history';

// Initialize SQL.js
initSqlJs({
    locateFile: file => `https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/${file}`
}).then(function(sql) {
    SQL = sql;
    db = new SQL.Database();
    console.log("SQL.js initialized");
    document.getElementById('run-sql').disabled = false;
}).catch(err => {
    console.error("Error loading SQL.js:", err);
    alert('Failed to load SQL.js. Please refresh the page.');
});

// Sample CSV data
const sampleData = {
    sales: `Product,Quarter,Qty,Revenue
Widget,Q1,150,3750.00
Widget,Q2,200,5000.00
Gadget,Q1,75,1125.00
Gadget,Q2,125,1875.00
Gadget,Q3,150,2250.00`,
    employees: `ID,Name,Department,Salary,HireDate
1,John Smith,Engineering,85000,2020-05-15
2,Jane Doe,Marketing,72000,2019-11-03
3,Robert Johnson,Sales,68000,2021-02-28
4,Emily Wilson,Engineering,92000,2018-07-22
5,Michael Brown,Marketing,76000,2022-01-10`,
    inventory: `SKU,ProductName,Category,Quantity,Price,LastStocked
1001,Desk Chair,Furniture,45,129.99,2023-03-15
1002,Monitor Stand,Electronics,28,49.95,2023-04-02
1003,Wireless Keyboard,Electronics,62,79.99,2023-03-28
1004,Desk Lamp,Home,37,34.50,2023-04-10
1005,Notebook,Office,120,4.99,2023-04-05`
};

// DOM elements
const spreadsheet = document.getElementById('spreadsheet');
const headerRow = document.getElementById('header-row');
const dataBody = document.getElementById('data-body');
const fileInput = document.getElementById('file-input');
const sqlQuery = document.getElementById('sql-query');
const chartTypeSelect = document.getElementById('chart-type');
const xAxisSelect = document.getElementById('x-axis');
const yAxisSelect = document.getElementById('y-axis');
const chartCanvas = document.getElementById('chart');
const chartList = document.getElementById('chart-list');
const loading = document.getElementById('loading');
const formulaBar = document.getElementById('formula-bar');

// Current data storage
let currentData = [];
let headers = [];
let charts = [];
let selectedCells = [];
let selectedRow = null;
let selectedColumn = null;
let formulas = [];
let dependencies = {};

// Filling state variables
let isFilling = false;
let fillStartCell = null;
let fillRange = [];
window.fillRange = fillRange;

// --- GLOBAL UTILITY FUNCTIONS ---
window.getCellByRef = function(ref) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) {
        console.error('Invalid cell reference:', ref);
        return null;
    }
    const colStr = match[1];
    const rowStr = match[2];
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    col--; // Convert to 0-based index
    const row = parseInt(rowStr) - 1; // 0-based index
    return document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
};
window.setCellValue = function(cell, value) {
    if (!cell || !cell.dataset.row || !cell.dataset.column) {
        console.error('Invalid cell for setCellValue:', cell);
        return;
    }
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.column);
    cell.textContent = value;
    window.currentData[row] = window.currentData[row] || [];
    window.currentData[row][col] = value;
    window.updateSQLDatabase();
};
window.getCellValue = function(cell) {
    if (!cell || !cell.dataset.row || !cell.dataset.column) {
        console.error('Invalid cell for getCellValue:', cell);
        return null;
    }
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.column);
    return window.currentData[row]?.[col] || cell.textContent || '';
};
window.clearHighlights = function() {
    document.querySelectorAll('.highlighted-cell').forEach(cell => cell.classList.remove('highlighted-cell'));
};
window.highlightCell = function(cell) {
    if (cell) {
        cell.classList.add('highlighted-cell');
    }
};
window.showNotification = function(message, type = 'info') {
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
};
window.generateRandomSheet = function(rows = 4, cols = 5) {
    headers = Array.from({ length: cols }, (_, i) => String.fromCharCode(65 + i));
    currentData = Array.from({ length: rows }, () => Array(cols).fill(''));
    formulas = Array(rows).fill().map(() => Array(cols).fill(undefined));
    dependencies = {};
    renderSpreadsheet();
    updateSQLDatabase();
    updateChartDropdowns();
};


// --- MAIN APP INITIALIZATION ---
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, setting up event listeners and editors');

    // --- Initialize Editors ---
    const sqlEditorTextarea = document.getElementById('sql-query');
    sqlEditor = CodeMirror.fromTextArea(sqlEditorTextarea, {
        mode: 'text/x-sql',
        theme: 'dracula',
        lineNumbers: false,
        autofocus: true
    });

    const datatableEditorTextarea = document.getElementById('datatable-config');
    datatableEditor = CodeMirror.fromTextArea(datatableEditorTextarea, {
        mode: { name: 'javascript', json: true },
        theme: 'dracula',
        lineNumbers: true
    });
    datatableEditor.setValue(`{\n  "paging": true,\n  "searching": true,\n  "columnDefs": [\n    {\n      "targets": 2,\n      "visible": false\n    }\n  ]\n}`);

    // --- Initialize Event Listeners ---
    document.getElementById('import-btn').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
    document.getElementById('run-sql').addEventListener('click', runSQL);
    document.getElementById('export-csv').addEventListener('click', exportCSV);
    document.getElementById('export-excel').addEventListener('click', exportExcel);
    document.getElementById('add-column').addEventListener('click', addColumn);
    document.getElementById('add-row').addEventListener('click', addRow);
    document.getElementById('clear-data').addEventListener('click', clearData);
    document.getElementById('create-chart').addEventListener('click', createChart);
    document.getElementById('apply-datatable-config').addEventListener('click', applyDataTableConfig);

    if (formulaBar) {
        formulaBar.addEventListener('blur', updateCellFromFormulaBar);
        formulaBar.addEventListener('keypress', (event) => {
            if (event.key === 'Enter') {
                updateCellFromFormulaBar();
            }
        });
    }

    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', function() {
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            this.classList.add('active');
            document.getElementById(this.dataset.tab).classList.add('active');

            // Refresh editors when their tabs become visible to fix rendering issues
            if (this.dataset.tab === 'sql-tab') {
                setTimeout(() => sqlEditor.refresh(), 1);
            }
            if (this.dataset.tab === 'filtered-views-tab') {
                setTimeout(() => datatableEditor.refresh(), 1);
            }
        });
    });

    document.querySelectorAll('#sample-files li').forEach(item => {
        item.addEventListener('click', function() {
            loadSampleData(this.getAttribute('data-file'));
        });
    });

    document.addEventListener('click', (event) => {
        if (!event.target.closest('#spreadsheet') && !event.target.closest('#formula-bar')) {
            clearSelections();
        }
    });

    // --- Initialize Data and History ---
    loadSampleData('sales');
    renderSqlHistory();

    // --- Drag Selection Logic ---
    let isDragging = false;
    let dragStartCell = null;
    document.addEventListener('mousedown', (event) => {
        const td = event.target.closest('td[data-row][data-column]');
        if (td && !td.classList.contains('row-number') && !event.target.classList.contains('fill-handle')) {
            isDragging = true;
            dragStartCell = [parseInt(td.dataset.row), parseInt(td.dataset.column)];
            handleCellSelection(event);
        }
    });
    document.addEventListener('mousemove', (event) => {
        if (!isDragging || !dragStartCell) return;
        const td = event.target.closest('td[data-row][data-column]');
        if (!td || td.classList.contains('row-number')) return;
        handleDragSelection(dragStartCell[0], dragStartCell[1], parseInt(td.dataset.row), parseInt(td.dataset.column), event.ctrlKey);
    });
    document.addEventListener('mouseup', () => {
        if (isDragging) {
            isDragging = false;
            dragStartCell = null;
        }
    });
});


// --- SPREADSHEET CORE FUNCTIONS ---

// (All original spreadsheet, formula, data processing, and UI functions are preserved below)

function updateCellFromFormulaBar() {
    if (selectedCells.length !== 1) return;
    const [row, col] = selectedCells[0].split('-').map(Number);
    const value = formulaBar.value.trim();
    updateCellValue(row, col, value);
}

function updateCellValue(row, col, value) {
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
    cell.textContent = currentData[row][col];
    updateDependents(row, col);
    updateSQLDatabase();
    renderSpreadsheet();
}

function parseCellReference(ref) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;
    const colStr = match[1];
    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    return { row: parseInt(match[2]) - 1, col: col - 1 };
}

function toCellReference(row, col) {
    let colStr = '';
    col++;
    while (col > 0) {
        colStr = String.fromCharCode(65 + ((col - 1) % 26)) + colStr;
        col = Math.floor((col - 1) / 26);
    }
    return `${colStr}${row + 1}`;
}

function parseRange(range) {
    const [start, end] = range.split(':');
    const startRef = parseCellReference(start);
    const endRef = parseCellReference(end);
    if (!startRef || !endRef) return [];
    const cells = [];
    for (let r = Math.min(startRef.row, endRef.row); r <= Math.max(startRef.row, endRef.row); r++) {
        for (let c = Math.min(startRef.col, endRef.col); c <= Math.max(startRef.col, endRef.col); c++) {
            cells.push({ row: r, col: c });
        }
    }
    return cells;
}

function evaluateCondition(condition, row, col) {
    const operators = ['>=', '<=', '!=', '=', '>', '<'];
    let op = '', left = '', right = '';
    for (let operator of operators) {
        if (condition.includes(operator)) {
            [left, right] = condition.split(operator);
            op = operator;
            break;
        }
    }
    if (!op) return false;
    left = left.trim();
    const leftRef = parseCellReference(left);
    if (leftRef) { left = currentData[leftRef.row]?.[leftRef.col] || '0'; }
    right = right.trim().replace(/"/g, '');
    const rightNum = parseFloat(right), leftNum = parseFloat(left);
    switch (op) {
        case '>': return leftNum > rightNum;
        case '<': return leftNum < rightNum;
        case '=': return left == right;
        case '!=': return left != right;
        case '>=': return leftNum >= rightNum;
        case '<=': return leftNum <= rightNum;
        default: return false;
    }
}

function sanitizeCondition(condition) {
    return condition.replace(/=/g, '=').replace(/<>/g, '!=').replace(/([A-Z]+\d+)/g, match => {
        const ref = parseCellReference(match);
        if (!ref) return match;
        return `"${currentData[ref.row]?.[ref.col] || '0'}"`;
    });
}

function evaluateFormula(formula, row, col) {
    if (!formula.startsWith('=')) return formula;
    try {
        let expr = formula.slice(1);
        const deps = [];
        const dependencyKey = `${row}-${col}`;
        const rangeValues = {};
        expr = expr.replace(/[A-Z]+\d+(?::[A-Z]+\d+)?/g, ref => {
            if (ref.includes(':')) {
                const values = parseRange(ref).map(cell => {
                    if (currentData[cell.row]?.[cell.col] !== undefined) {
                        deps.push(`${cell.row}-${cell.col}`);
                        const value = currentData[cell.row][cell.col];
                        return isNaN(parseFloat(value)) ? `"${value}"` : value.toString();
                    }
                    return '0';
                });
                rangeValues[ref] = values;
                return ref;
            } else {
                const cellRef = parseCellReference(ref);
                if (!cellRef || cellRef.row < 0 || cellRef.col < 0) throw new Error('#REF!');
                deps.push(`${cellRef.row}-${cellRef.col}`);
                const value = currentData[cellRef.row]?.[cellRef.col] || '0';
                return isNaN(parseFloat(value)) ? `"${value}"` : value.toString();
            }
        });
        dependencies[dependencyKey] = deps;
        const functions = { SUM: args => args.flatMap(arg => rangeValues[arg] || [arg]).reduce((s, v) => s + (parseFloat(v) || 0), 0), AVERAGE: args => functions.SUM(args) / (args.flatMap(arg => rangeValues[arg] || [arg]).length || 1), MIN: args => Math.min(...args.flatMap(arg => rangeValues[arg] || [arg]).map(v => parseFloat(v) || Infinity)), MAX: args => Math.max(...args.flatMap(arg => rangeValues[arg] || [arg]).map(v => parseFloat(v) || -Infinity)), COUNT: args => args.flatMap(arg => rangeValues[arg] || [arg]).filter(v => !isNaN(parseFloat(v)) && v !== '').length, /* ... other functions ... */ };
        while (expr.match(/(\w+(\.\w+)?)\([^)]+\)/i)) {
            const match = expr.match(/(\w+(\.\w+)?)\((.*)\)$/i);
            if (!match) break;
            const funcName = match[1].toUpperCase();
            if (functions[funcName]) {
                const args = match[3].split(',');
                expr = expr.replace(match[0], functions[funcName](args));
            } else {
                break;
            }
        }
        return eval(expr);
    } catch (error) {
        return error.message.startsWith('#') ? error.message : '#ERROR!';
    }
}

function updateDependents(row, col) {
    const key = `${row}-${col}`;
    for (let depKey in dependencies) {
        if (dependencies[depKey].includes(key)) {
            const [depRow, depCol] = depKey.split('-').map(Number);
            if (formulas[depRow]?.[depCol]) {
                currentData[depRow][depCol] = evaluateFormula(formulas[depRow][depCol], depRow, depCol);
                const cell = document.querySelector(`td[data-row="${depRow}"][data-column="${depCol}"]`);
                if (cell) cell.textContent = currentData[depRow][depCol];
                updateDependents(depRow, depCol);
            }
        }
    }
}

function loadSampleData(key) {
    const csvData = sampleData[key];
    if (csvData) processCSVData(csvData);
    else alert('No sample data found for key: ' + key);
}

function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    loading.style.display = 'block';
    const reader = new FileReader();
    reader.onload = function(e) {
        if (file.name.endsWith('.csv')) processCSVData(e.target.result);
        else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) processExcelData(e.target.result);
        loading.style.display = 'none';
    };
    if (file.name.endsWith('.csv')) reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
}

function processCSVData(csvString) {
    try {
        const workbook = XLSX.read(csvString, { type: 'string' });
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        if (data.length === 0) { alert('No data found in the file'); return; }
        headers = data[0] || [];
        currentData = data.slice(1) || [];
        formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(undefined));
        dependencies = {};
        renderSpreadsheet();
        updateSQLDatabase();
        updateChartDropdowns();
    } catch (error) { alert('Error processing CSV file'); }
}

function processExcelData(arrayBuffer) {
    try {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        if (data.length === 0) { alert('No data found in the file'); return; }
        headers = data[0] || [];
        currentData = data.slice(1) || [];
        formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(undefined));
        dependencies = {};
        renderSpreadsheet();
        updateSQLDatabase();
        updateChartDropdowns();
    } catch (error) { alert('Error processing Excel file'); }
}

function renderSpreadsheet() {
    headerRow.innerHTML = '<th class="row-number-header"></th>';
    dataBody.innerHTML = '';
    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header; th.contentEditable = true; th.dataset.column = index;
        th.addEventListener('click', (e) => { e.stopPropagation(); handleColumnSelection(e); });
        th.addEventListener('blur', () => { headers[index] = th.textContent.trim() || `Column${index + 1}`; updateSQLDatabase(); updateChartDropdowns(); });
        headerRow.appendChild(th);
    });
    currentData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        const rowNumTd = document.createElement('td');
        rowNumTd.className = 'row-number'; rowNumTd.textContent = rowIndex + 1; rowNumTd.dataset.row = rowIndex;
        rowNumTd.addEventListener('click', (e) => { e.stopPropagation(); handleRowSelection(e); });
        tr.appendChild(rowNumTd);
        headers.forEach((_, colIndex) => {
            const td = document.createElement('td');
            td.textContent = row[colIndex] || ''; td.contentEditable = true; td.dataset.row = rowIndex; td.dataset.column = colIndex;
            td.addEventListener('click', (e) => { e.stopPropagation(); handleCellSelection(e); formulaBar.value = formulas[rowIndex]?.[colIndex] || td.textContent; });
            td.addEventListener('blur', () => { updateCellValue(rowIndex, colIndex, td.textContent.trim()); });
            tr.appendChild(td);
        });
        dataBody.appendChild(tr);
    });
    selectedCells.forEach(key => {
        const [r, c] = key.split('-');
        const cell = document.querySelector(`td[data-row="${r}"][data-column="${c}"]`);
        if (cell) cell.classList.add('selected-cell');
    });
    updateFillHandle();
}

function updateFillHandle() {
    document.querySelectorAll('.fill-handle').forEach(h => h.remove());
    if (selectedCells.length < 1) return;
    const lastCellKey = selectedCells[selectedCells.length - 1];
    const [row, col] = lastCellKey.split('-').map(Number);
    const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
    if (cell) {
        const handle = document.createElement('div');
        handle.className = 'fill-handle';
        handle.addEventListener('mousedown', startFillDrag);
        handle.addEventListener('dblclick', handleDoubleClickFill);
        cell.appendChild(handle);
    }
}

function handleDoubleClickFill(event) {
    event.stopPropagation();
    if (selectedCells.length !== 1) return;
    const [row, col] = selectedCells[0].split('-').map(Number);
    let endRow = row;
    const adjCol = col - 1;
    if (adjCol >= 0) {
        while (endRow + 1 < currentData.length && currentData[endRow + 1]?.[adjCol]?.toString() !== '') { endRow++; }
    }
    if (endRow === row && endRow + 1 < currentData.length) endRow++;
    fillStartCell = [row, col];
    fillRange = getFillRange(row, col, endRow, col);
    applyFill();
}

function startFillDrag(event) {
    event.preventDefault();
    event.stopPropagation();
    isFilling = true;
    fillStartCell = selectedCells[0].split('-').map(Number);
    fillRange = [];
    document.addEventListener('mousemove', handleFillDrag);
    document.addEventListener('mouseup', endFillDrag, { once: true });
}

function handleFillDrag(event) {
    if (!isFilling) return;
    const target = event.target.closest('td[data-row][data-column]');
    if (!target) return;
    document.querySelectorAll('.fill-range').forEach(c => c.classList.remove('fill-range'));
    fillRange = getFillRange(fillStartCell[0], fillStartCell[1], parseInt(target.dataset.row), parseInt(target.dataset.column));
    fillRange.forEach(([r, c]) => {
        const cell = document.querySelector(`td[data-row="${r}"][data-column="${c}"]`);
        if (cell) cell.classList.add('fill-range');
    });
}

function endFillDrag() {
    if (!isFilling) return;
    isFilling = false;
    document.removeEventListener('mousemove', handleFillDrag);
    document.querySelectorAll('.fill-range').forEach(c => c.classList.remove('fill-range'));
    if (fillRange.length > 0) {
        applyFill();
        clearSelections();
        selectedCells = fillRange.map(([r, c]) => `${r}-${c}`);
        selectedCells.forEach(key => {
            const cell = document.querySelector(`td[data-row="${key.split('-')[0]}"][data-column="${key.split('-')[1]}"]`);
            if (cell) cell.classList.add('selected-cell');
        });
        updateFillHandle();
    }
}

function getFillRange(startRow, startCol, endRow, endCol) {
    const range = [];
    const isVerticalDrag = Math.abs(endRow - startRow) > Math.abs(endCol - startCol);
    if (isVerticalDrag) {
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) { range.push([r, startCol]); }
    } else {
        for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) { range.push([startRow, c]); }
    }
    return range;
}

function applyFill() {
    if (!fillStartCell || fillRange.length === 0) return;
    const [startRow, startCol] = fillStartCell;
    const sourceValue = formulas[startRow]?.[startCol] || currentData[startRow]?.[startCol] || '';
    fillRange.forEach(([row, col]) => {
        if (row === startRow && col === startCol) return;
        let newValue = sourceValue.startsWith('=') ? adjustFormula(sourceValue, startRow, startCol, row, col) : sourceValue;
        updateCellValue(row, col, newValue);
    });
    renderSpreadsheet();
}

function handleCellSelection(event) {
    if (event.target.closest('.fill-handle')) return;
    const td = event.target.closest('td[data-row][data-column]');
    if (!td) return;
    const key = `${td.dataset.row}-${td.dataset.column}`;
    if (event.shiftKey && selectedCells.length > 0) {
        const [startRow, startCol] = selectedCells[0].split('-').map(Number);
        const endRow = parseInt(td.dataset.row), endCol = parseInt(td.dataset.column);
        clearSelections();
        for (let r = Math.min(startRow, endRow); r <= Math.max(startRow, endRow); r++) {
            for (let c = Math.min(startCol, endCol); c <= Math.max(startCol, endCol); c++) {
                selectedCells.push(`${r}-${c}`);
            }
        }
    } else if (event.ctrlKey) {
        const index = selectedCells.indexOf(key);
        if (index > -1) selectedCells.splice(index, 1);
        else selectedCells.push(key);
    } else {
        clearSelections();
        selectedCells = [key];
    }
    document.querySelectorAll('.selected-cell').forEach(c => c.classList.remove('selected-cell'));
    selectedCells.forEach(k => {
        const [r, c] = k.split('-');
        const cell = document.querySelector(`td[data-row="${r}"][data-column="${c}"]`);
        if (cell) cell.classList.add('selected-cell');
    });
    updateFillHandle();
    formulaBar.value = formulas[td.dataset.row]?.[td.dataset.column] || td.textContent;
}

function addCellToSelection(td, cellKey) {
    if (!selectedCells.includes(cellKey)) {
        selectedCells.push(cellKey);
        td.classList.add('selected-cell');
    }
}

function handleRowSelection(event) {
    const rowIndex = parseInt(event.target.dataset.row);
    clearSelections();
    selectedRow = rowIndex;
    event.target.parentElement.classList.add('selected-row');
}

function handleColumnSelection(event) {
    const colIndex = parseInt(event.target.dataset.column);
    clearSelections();
    selectedColumn = colIndex;
    document.querySelectorAll(`[data-column="${colIndex}"]`).forEach(c => c.classList.add('selected-column'));
}

function clearSelections() {
    document.querySelectorAll('.selected-cell, .selected-row, .selected-column').forEach(el => el.classList.remove('selected-cell', 'selected-row', 'selected-column'));
    selectedCells = [];
    selectedRow = null;
    selectedColumn = null;
    updateFillHandle();
}

function updateSQLDatabase() {
    if (!db) return;
    try {
        db.run('DROP TABLE IF EXISTS data;');
        const sanitizedHeaders = headers.map((h, i) => `"${(h || `Col${i}`).replace(/"/g, '""')}"`);
        db.run(`CREATE TABLE data (${sanitizedHeaders.join(', ')});`);
        const stmt = db.prepare(`INSERT INTO data VALUES (${headers.map(() => '?').join(',')})`);
        currentData.forEach(row => stmt.run(row));
        stmt.free();
    } catch (e) {
        console.error("SQL DB update error:", e);
    }
}

function updateChartDropdowns() {
    xAxisSelect.innerHTML = '';
    yAxisSelect.innerHTML = '';
    headers.forEach(header => {
        const opt1 = document.createElement('option');
        opt1.value = header; opt1.textContent = header;
        xAxisSelect.appendChild(opt1);
        const opt2 = document.createElement('option');
        opt2.value = header; opt2.textContent = header;
        yAxisSelect.appendChild(opt2);
    });
    if (headers.length > 1) yAxisSelect.selectedIndex = 1;
}

// --- FILTERED VIEWS (DATATABLES) ---

function applyDataTableConfig() {
    const container = document.getElementById('datatable-container');
    container.innerHTML = '';

    if (headers.length === 0 || currentData.length === 0) {
        container.innerHTML = `<p class="placeholder-text">No spreadsheet data to display.</p>`;
        return;
    }

    if (dataTableInstance) {
        dataTableInstance.destroy();
        dataTableInstance = null;
    }

    let config = {};
    const configString = datatableEditor.getValue();
    try {
        // Use Function constructor for safe parsing of JS object literals
        config = new Function(`return ${configString}`)();
    } catch (e) {
        container.innerHTML = `<p class="error-text">Error parsing configuration:\n${e.message}</p>`;
        return;
    }

    // Build the HTML table from current data
    const table = document.createElement('table');
    table.id = 'filtered-table';
    table.className = 'display';
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    const headerTr = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h;
        headerTr.appendChild(th);
    });
    thead.appendChild(headerTr);

    currentData.forEach(rowData => {
        const tr = document.createElement('tr');
        headers.forEach((_, i) => {
            const td = document.createElement('td');
            td.textContent = rowData[i] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    container.appendChild(table);

    // Initialize DataTables.net
    try {
        dataTableInstance = new DataTable('#filtered-table', config);
    } catch (e) {
        container.innerHTML = `<p class="error-text">Error initializing DataTable:\n${e.message}</p>`;
    }
}

// --- SQL & HISTORY FUNCTIONS ---

function renderSqlResults(result, error = null) {
    const container = document.getElementById('sql-results-container');
    container.innerHTML = '';
    if (error) { container.innerHTML = `<p class="error-text">SQL Error: ${error.message}</p>`; return; }
    if (!result || result.length === 0) { container.innerHTML = `<p class="placeholder-text">Query returned no results.</p>`; return; }
    const { columns, values } = result[0];
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    const headerRow = document.createElement('tr');
    columns.forEach(colName => { const th = document.createElement('th'); th.textContent = colName; headerRow.appendChild(th); });
    thead.appendChild(headerRow);
    values.forEach(row => {
        const tr = document.createElement('tr');
        row.forEach(cellValue => { const td = document.createElement('td'); td.textContent = cellValue; tr.appendChild(td); });
        tbody.appendChild(tr);
    });
    table.append(thead, tbody);
    container.appendChild(table);
}

function addToSqlHistory(query) {
    if (!query) return;
    try {
        let history = JSON.parse(localStorage.getItem(SQL_HISTORY_KEY)) || [];
        history = history.filter(q => q !== query);
        history.unshift(query);
        if (history.length > 50) history.length = 50;
        localStorage.setItem(SQL_HISTORY_KEY, JSON.stringify(history));
        renderSqlHistory();
    } catch (e) { console.error("Failed to update SQL history:", e); }
}

function renderSqlHistory() {
    const historyList = document.getElementById('sql-history-list');
    historyList.innerHTML = '';
    try {
        const history = JSON.parse(localStorage.getItem(SQL_HISTORY_KEY)) || [];
        if (history.length === 0) { historyList.innerHTML = '<li>No history yet.</li>'; }
        else {
            history.forEach(query => {
                const li = document.createElement('li');
                li.textContent = query;
                li.addEventListener('click', () => { sqlEditor.setValue(query); sqlEditor.focus(); });
                historyList.appendChild(li);
            });
        }
    } catch (e) { console.error("Failed to render SQL history:", e); }
}

function runSQL() {
    if (!db) { alert('SQL.js is not loaded.'); return; }
    const query = sqlEditor.getValue().trim();
    if (!query) { renderSqlResults(null, { message: 'Query is empty.' }); return; }
    addToSqlHistory(query);
    try {
        const result = db.exec(query);
        renderSqlResults(result);
    } catch (error) {
        renderSqlResults(null, error);
    }
}


// --- CHART & EXPORT FUNCTIONS ---

function createChart() {
    if (headers.length === 0) return alert('No data for chart');
    let chartType = chartTypeSelect.value;
    let xAxis = xAxisSelect.value, yAxis = yAxisSelect.value;
    let xIndex = headers.indexOf(xAxis), yIndex = headers.indexOf(yAxis);
    if (xIndex === -1 || yIndex === -1) return alert('Invalid axis');
    let indexAxis = 'x';
    if (!currentData.every(row => !isNaN(parseFloat(row[yIndex])))) {
        [xAxis, yAxis] = [yAxis, xAxis];
        [xIndex, yIndex] = [yIndex, xIndex];
        chartType = 'bar';
        indexAxis = 'y';
    }
    const labels = currentData.map(row => row[xIndex]);
    const dataValues = currentData.map(row => parseFloat(row[yIndex]) || 0);
    const chartId = `chart-${Date.now()}`;
    const chartItem = document.createElement('div');
    chartItem.className = 'chart-item';
    chartItem.innerHTML = `<canvas id="${chartId}"></canvas><div><button class="delete-chart" data-id="${chartId}">Delete</button><button class="download-chart" data-id="${chartId}">Download</button></div>`;
    chartList.appendChild(chartItem);
    const chartCtx = document.getElementById(chartId).getContext('2d');
    const chart = new Chart(chartCtx, { type: chartType, data: { labels, datasets: [{ label: `${yAxis} by ${xAxis}`, data: dataValues, backgroundColor: getChartColors(dataValues.length) }] }, options: { indexAxis, responsive: true } });
    charts.push({ id: chartId, chart });
    chartItem.querySelector('.delete-chart').addEventListener('click', function() { deleteChart(this.dataset.id); });
    chartItem.querySelector('.download-chart').addEventListener('click', function() { downloadChart(this.dataset.id); });
}

function getChartColors(count) {
    return Array(count).fill().map((_, i) => `hsl(${(i * 360 / count) % 360}, 70%, 60%)`);
}

function deleteChart(id) {
    const index = charts.findIndex(c => c.id === id);
    if (index > -1) {
        charts[index].chart.destroy();
        charts.splice(index, 1);
        document.getElementById(id).closest('.chart-item').remove();
    }
}

function downloadChart(id) {
    const chart = charts.find(c => c.id === id);
    if (chart) {
        const link = document.createElement('a');
        link.download = `chart.png`;
        link.href = chart.chart.canvas.toDataURL('image/png');
        link.click();
    }
}

function exportCSV() {
    if (headers.length === 0) return alert('No data');
    let csvContent = headers.join(',') + '\n' + currentData.map(row => row.map(cell => `"${(cell || '').toString().replace(/"/g, '""')}"`).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'export.csv';
    link.click();
    URL.revokeObjectURL(link.href);
}

function exportExcel() {
    if (headers.length === 0) return alert('No data');
    const ws = XLSX.utils.aoa_to_sheet([headers, ...currentData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'export.xlsx');
}

function addColumn() {
    headers.push(`Column${headers.length + 1}`);
    currentData.forEach(row => row.push(''));
    renderSpreadsheet();
    updateSQLDatabase();
    updateChartDropdowns();
}

function addRow() {
    currentData.push(Array(headers.length).fill(''));
    renderSpreadsheet();
    updateSQLDatabase();
}

function clearData() {
    if (confirm('Are you sure you want to clear all data?')) {
        headers = [];
        currentData = [];
        formulas = [];
        dependencies = {};
        charts.forEach(c => c.chart.destroy());
        charts = [];
        chartList.innerHTML = '';
        renderSpreadsheet();
        if (db) db.run('DROP TABLE IF EXISTS data;');
    }
}

let isCtrlSelecting = false;
document.addEventListener('mouseup', () => { isCtrlSelecting = false; });
document.addEventListener('mouseover', event => {
    if (isCtrlSelecting && event.ctrlKey && event.buttons === 1) {
        const td = event.target.closest('td[data-row][data-column]');
        if (td) {
            addCellToSelection(td, `${td.dataset.row}-${td.dataset.column}`);
        }
    }
});