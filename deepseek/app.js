
let SQL;
let db;

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
let formulas = []; // Store raw formulas: formulas[row][col] = "=SUM(A1:A5)"
let dependencies = {}; // Track dependencies: { 'row-col': ['dep-row-dep-col', ...] }

// Filling state variables
let isFilling = false;
let fillStartCell = null;
let fillRange = [];
window.fillRange = fillRange; // Initialize globally
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
    const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
    console.log(`getCellByRef(${ref}): row=${row}, col=${col}, cell=${!!cell}`);
    return cell;
};

window.setCellValue = function(cell, value) {
    if (!cell || !cell.dataset.row || !cell.dataset.column) {
        console.error('Invalid cell for setCellValue:', cell);
        return;
    }
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.column);
    console.log(`setCellValue: row=${row}, col=${col}, value=${value}`);
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
        console.log('Highlighted cell:', cell.dataset.row, cell.dataset.column);
    }
};
window.showNotification = function(message, type = 'info') {
    console.log('Showing notification:', message, type);
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.style.display = 'block';
};

window.generateRandomSheet = function(rows = 4, cols = 5) {
    console.log('Generating sheet:', rows, cols);
    const headers = Array.from({ length: cols }, (_, i) => String.fromCharCode(65 + i));
    const data = Array.from({ length: rows }, () => Array(cols).fill(''));
    window.headers = headers;
    window.currentData = data;
    window.formulas = Array(rows).fill().map(() => Array(cols).fill(undefined));
    window.dependencies = {};
    window.renderSpreadsheet();
    window.updateSQLDatabase();
    window.updateChartDropdowns();
};
// Initialize the app
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, setting up event listeners');
    // Event listeners
    const importBtn = document.getElementById('import-btn');
    if (importBtn) importBtn.addEventListener('click', () => {
        console.log('Import button clicked');
        if (fileInput) fileInput.click();
    });
    if (fileInput) fileInput.addEventListener('change', handleFileSelect);
    const runSqlButton = document.getElementById('run-sql');
    if (runSqlButton) runSqlButton.addEventListener('click', runSQL);
    const exportCsvButton = document.getElementById('export-csv');
    if (exportCsvButton) exportCsvButton.addEventListener('click', exportCSV);
    const exportExcelButton = document.getElementById('export-excel');
    if (exportExcelButton) exportExcelButton.addEventListener('click', exportExcel);
    const addColumnButton = document.getElementById('add-column');
    if (addColumnButton) addColumnButton.addEventListener('click', addColumn);
    const addRowButton = document.getElementById('add-row');
    if (addRowButton) addRowButton.addEventListener('click', addRow);
    const clearDataButton = document.getElementById('clear-data');
    if (clearDataButton) clearDataButton.addEventListener('click', clearData);
    const createChartButton = document.getElementById('create-chart');
    if (createChartButton) createChartButton.addEventListener('click', createChart);

    // Formula bar listeners
    if (formulaBar) {
        formulaBar.addEventListener('blur', updateCellFromFormulaBar);
        formulaBar.addEventListener('keypress', function(event) {
            if (event.key === 'Enter') {
                updateCellFromFormulaBar();
            }
        });
    }

    // Tab switching
    document.querySelectorAll('.tab-button').forEach(button => {
        button.addEventListener('click', function() {
            console.log(`Tab clicked: ${this.dataset.tab}`);
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
            this.classList.add('active');
            document.getElementById(this.dataset.tab).classList.add('active');
        });
    });

    // Sample file click handlers
    document.querySelectorAll('#sample-files li').forEach(item => {
        item.addEventListener('click', function() {
            const fileKey = this.getAttribute('data-file');
            console.log(`Sample file clicked: ${fileKey}`);
            loadSampleData(fileKey);
        });
    });

    // Clear selections when clicking outside
    document.addEventListener('click', function(event) {
        if (!event.target.closest('#spreadsheet') && !event.target.closest('#formula-bar')) {
            console.log('Clicked outside spreadsheet, clearing selections');
            clearSelections();
        }
    });

    // Initialize with sample data
    console.log('Loading initial sales data');
    loadSampleData('sales');

    // Add mousemove and mouseup listeners for drag selection
    let isDragging = false;
    let dragStartCell = null;

    document.addEventListener('mousedown', function(event) {
        const td = event.target.closest('td[data-row][data-column]');
        if (td && !td.classList.contains('row-number') && !event.target.classList.contains('fill-handle')) {
            isDragging = true;
            const rowIndex = parseInt(td.dataset.row);
            const colIndex = parseInt(td.dataset.column);
            dragStartCell = [rowIndex, colIndex];
            handleCellSelection(event);
        }
    });

    document.addEventListener('mousemove', function(event) {
        if (!isDragging || !dragStartCell) return;
        const td = event.target.closest('td[data-row][data-column]');
        if (!td || td.classList.contains('row-number')) return;

        const rowIndex = parseInt(td.dataset.row);
        const colIndex = parseInt(td.dataset.column);
        handleDragSelection(dragStartCell[0], dragStartCell[1], rowIndex, colIndex, event.ctrlKey);
    });

    document.addEventListener('mouseup', function() {
        if (isDragging) {
            isDragging = false;
            dragStartCell = null;
            console.log('Drag selection ended, selected cells:', selectedCells);
        }
    });
});

// Update cell from formula bar
function updateCellFromFormulaBar() {
    if (selectedCells.length !== 1) return;
    const [row, col] = selectedCells[0].split('-').map(Number);
    const value = formulaBar.value.trim();
    console.log(`Updating cell (${row}, ${col}) from formula bar with: ${value}`);
    updateCellValue(row, col, value);
}

// Update cell value and handle formulas
function updateCellValue(row, col, value) {
    const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
    if (!cell) return;

    // Store formula if applicable
    if (value.startsWith('=')) {
        formulas[row] = formulas[row] || [];
        formulas[row][col] = value;
    } else {
        if (formulas[row]?.[col]) {
            delete formulas[row][col];
        }
    }

    // Evaluate and update data
    currentData[row] = currentData[row] || [];
    currentData[row][col] = evaluateFormula(value, row, col);
    cell.textContent = currentData[row][col];

    updateDependents(row, col);
    updateSQLDatabase();
    renderSpreadsheet();
}

// Parse cell reference (e.g., "A1" → { row: 0, col: 0 })
function parseCellReference(ref) {
    const match = ref.match(/^([A-Z]+)(\d+)$/);
    if (!match) return null;

    const colStr = match[1];
    const rowStr = match[2];

    let col = 0;
    for (let i = 0; i < colStr.length; i++) {
        col = col * 26 + (colStr.charCodeAt(i) - 64);
    }
    col--; // 0-based index

    const row = parseInt(rowStr) - 1; // 0-based index
    return { row, col };
}

// Convert row/col to cell reference (e.g., row 0, col 0 → "A1")
function toCellReference(row, col) {
    let colStr = '';
    col++; // 1-based index for display
    while (col > 0) {
        colStr = String.fromCharCode(65 + ((col - 1) % 26)) + colStr;
        col = Math.floor((col - 1) / 26);
    }
    return `${colStr}${row + 1}`;
}

// Parse range (e.g., "A1:A5" → [{ row: 0, col: 0 }, ..., { row: 4, col: 0 }])
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

// Evaluate condition without eval
function evaluateCondition(condition, row, col) {
    console.log(`Evaluating condition: ${condition} at (${row}, ${col})`);
    const operators = ['>=', '<=', '!=', '=', '>', '<'];
    let op = '';
    let left = '';
    let right = '';

    // Find operator
    for (let operator of operators) {
        if (condition.includes(operator)) {
            [left, right] = condition.split(operator);
            op = operator;
            break;
        }
    }

    if (!op) return false;

    // Resolve cell references in left part
    left = left.trim();
    const leftRef = parseCellReference(left);
    if (leftRef) {
        left = currentData[leftRef.row]?.[leftRef.col] || '0';
    }

    // Parse right part (number or string)
    right = right.trim().replace(/"/g, '');
    const rightNum = parseFloat(right);
    const leftNum = parseFloat(left);

    // Compare based on operator
    switch (op) {
        case '>': return leftNum > rightNum;
        case '<': return leftNum < rightNum;
        case '=': return left == right; // Loose equality for text
        case '!=': return left != right;
        case '>=': return leftNum >= rightNum;
        case '<=': return leftNum <= rightNum;
        default: return false;
    }
}

// Sanitize condition for processing
function sanitizeCondition(condition) {
    // Replace operators
    return condition
        .replace(/=/g, '=')
        .replace(/<>/g, '!=')
        .replace(/([A-Z]+\d+)/g, match => {
            const ref = parseCellReference(match);
            if (!ref) return match;
            return `"${currentData[ref.row]?.[ref.col] || '0'}"`;
        });
}

// Evaluate formula
function evaluateFormula(formula, row, col) {
    if (!formula.startsWith('=')) return formula;

    try {
        let expr = formula.slice(1); // Remove '='
        console.log(`Evaluating formula: ${expr} at (${row}, ${col})`);

        // Track dependencies
        const deps = [];
        const dependencyKey = `${row}-${col}`;
        const rangeValues = {}; // Store range values: { 'C2:C3': ['200', '75'] }

        // Replace cell references and ranges
        expr = expr.replace(/[A-Z]+\d+(?::[A-Z]+\d+)?/g, ref => {
            if (ref.includes(':')) {
                // Handle range
                const values = parseRange(ref).map(cell => {
                    if (currentData[cell.row]?.[cell.col] !== undefined) {
                        deps.push(`${cell.row}-${cell.col}`);
                        const value = currentData[cell.row][cell.col];
                        return isNaN(parseFloat(value)) ? `"${value}"` : value.toString();
                    }
                    return '0';
                });
                rangeValues[ref] = values;
                return ref; // Keep reference for function parsing
            } else {
                // Handle single cell
                const cellRef = parseCellReference(ref);
                if (!cellRef || cellRef.row < 0 || cellRef.col < 0) throw new Error('#REF!');
                deps.push(`${cellRef.row}-${cellRef.col}`);
                const value = currentData[cellRef.row]?.[cellRef.col] || '0';
                return isNaN(parseFloat(value)) ? `"${value}"` : value.toString();
            }
        });

        // Update dependencies
        dependencies[dependencyKey] = deps;

        // Define functions
        const functions = {
            SUM: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return flatArgs.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
            },
            AVERAGE: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return functions.SUM(flatArgs) / (flatArgs.length || 1);
            },
            MIN: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return Math.min(...flatArgs.map(val => parseFloat(val) || Infinity));
            },
            MAX: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return Math.max(...flatArgs.map(val => parseFloat(val) || -Infinity));
            },
            COUNT: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return flatArgs.filter(val => !isNaN(parseFloat(val)) && val !== '').length;
            },
            AVERAGEIF: args => {
                const [range, criterion, avgRange] = args;
                const rangeCells = parseRange(range);
                const values = rangeCells.map(cell => currentData[cell.row]?.[cell.col] || '0');
                const avgCells = avgRange ? parseRange(avgRange) : rangeCells;
                const avgValues = avgCells.map(cell => parseFloat(currentData[cell.row]?.[cell.col]) || 0);
                let sum = 0, count = 0;

                for (let i = 0; i < values.length; i++) {
                    if (evaluateCondition(criterion.replace(/"/g, ''), row, col)) {
                        sum += avgValues[i] || 0;
                        count++;
                    }
                }
                return count > 0 ? sum / count : '#DIV/0!';
            },
            AVERAGEIFS: args => {
                const [avgRange, ...criteriaPairs] = args;
                if (criteriaPairs.length % 2 !== 0) throw new Error('#VALUE!');
                const avgCells = parseRange(avgRange);
                const avgValues = avgCells.map(cell => parseFloat(currentData[cell.row]?.[cell.col]) || 0);
                let sum = 0, count = 0;

                for (let i = 0; i < avgCells.length; i++) {
                    let meetsCriteria = true;
                    for (let j = 0; j < criteriaPairs.length; j += 2) {
                        const critRange = parseRange(criteriaPairs[j]);
                        const criterion = criteriaPairs[j + 1];
                        const value = currentData[critRange[i].row]?.[critRange[i].col] || '0';
                        if (!evaluateCondition(criterion.replace(/"/g, ''), row, col)) {
                            meetsCriteria = false;
                            break;
                        }
                    }
                    if (meetsCriteria) {
                        sum += avgValues[i];
                        count++;
                    }
                }
                return count > 0 ? sum / count : '#DIV/0!';
            },
            COUNTA: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return flatArgs.filter(val => val !== '' && val !== '0').length;
            },
            COUNTBLANK: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                return flatArgs.filter(val => val === '' || val === '0').length;
            },
            COUNTIF: args => {
                const [range, criterion] = args;
                const rangeCells = parseRange(range);
                const values = rangeCells.map(cell => currentData[cell.row]?.[cell.col] || '0');
                return values.filter(val => evaluateCondition(criterion.replace(/"/g, ''), row, col)).length;
            },
            COUNTIFS: args => {
                if (args.length % 2 !== 0) throw new Error('#VALUE!');
                const ranges = args.filter((_, i) => i % 2 === 0).map(range => parseRange(range));
                const criteria = args.filter((_, i) => i % 2 === 1);
                if (!ranges.every(r => r.length === ranges[0].length)) throw new Error('#VALUE!');

                let count = 0;
                for (let i = 0; i < ranges[0].length; i++) {
                    let meetsCriteria = true;
                    for (let j = 0; j < ranges.length; j++) {
                        const value = currentData[ranges[j][i].row]?.[ranges[j][i].col] || '0';
                        if (!evaluateCondition(criteria[j].replace(/"/g, ''), row, col)) {
                            meetsCriteria = false;
                            break;
                        }
                    }
                    if (meetsCriteria) count++;
                }
                return count;
            },
            SUMIF: args => {
                const [range, criterion, sumRange] = args;
                const rangeCells = parseRange(range);
                const values = rangeCells.map(cell => currentData[cell.row]?.[cell.col] || '0');
                const sumCells = sumRange ? parseRange(sumRange) : rangeCells;
                const sumValues = sumCells.map(cell => parseFloat(currentData[cell.row]?.[cell.col]) || 0);
                let sum = 0;

                for (let i = 0; i < values.length; i++) {
                    if (evaluateCondition(criterion.replace(/"/g, ''), row, col)) {
                        sum += sumValues[i] || 0;
                    }
                }
                return sum;
            },
            SUMIFS: args => {
                const [sumRange, ...criteriaPairs] = args;
                if (criteriaPairs.length % 2 !== 0) throw new Error('#VALUE!');
                const sumCells = parseRange(sumRange);
                const sumValues = sumCells.map(cell => parseFloat(currentData[cell.row]?.[cell.col]) || 0);
                let sum = 0;

                for (let i = 0; i < sumCells.length; i++) {
                    let meetsCriteria = true;
                    for (let j = 0; j < criteriaPairs.length; j += 2) {
                        const critRange = parseRange(criteriaPairs[j]);
                        const criterion = criteriaPairs[j + 1];
                        const value = currentData[critRange[i].row]?.[critRange[i].col] || '0';
                        if (!evaluateCondition(criterion.replace(/"/g, ''), row, col)) {
                            meetsCriteria = false;
                            break;
                        }
                    }
                    if (meetsCriteria) {
                        sum += sumValues[i];
                    }
                }
                return sum;
            },
            MEDIAN: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                const numbers = flatArgs.map(val => parseFloat(val)).filter(val => !isNaN(val));
                if (numbers.length === 0) return '#NUM!';
                numbers.sort((a, b) => a - b);
                const mid = Math.floor(numbers.length / 2);
                return numbers.length % 2 === 0 ? (numbers[mid - 1] + numbers[mid]) / 2 : numbers[mid];
            },
            MODE: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                const numbers = flatArgs.map(val => parseFloat(val)).filter(val => !isNaN(val));
                if (numbers.length === 0) return '#NUM!';
                const frequency = {};
                let maxFreq = 0;
                let mode = null;
                numbers.forEach(num => {
                    frequency[num] = (frequency[num] || 0) + 1;
                    if (frequency[num] > maxFreq) {
                        maxFreq = frequency[num];
                        mode = num;
                    }
                });
                return mode !== null ? mode : '#N/A';
            },
            STDEV_P: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                const numbers = flatArgs.map(val => parseFloat(val)).filter(val => !isNaN(val));
                if (numbers.length === 0) return '#NUM!';
                const mean = numbers.reduce((sum, val) => sum + val, 0) / numbers.length;
                const variance = numbers.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / numbers.length;
                return Math.sqrt(variance);
            },
            STDEV_S: args => {
                const flatArgs = args.flatMap(arg => rangeValues[arg] || [arg]);
                const numbers = flatArgs.map(val => parseFloat(val)).filter(val => !isNaN(val));
                if (numbers.length <= 1) return '#DIV/0!';
                const mean = numbers.reduce((sum, val) => sum + val, 0) / numbers.length;
                const variance = numbers.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / (numbers.length - 1);
                return Math.sqrt(variance);
            },
            NPV: args => {
                const [rate, ...values] = args;
                const rateVal = parseFloat(rate);
                if (isNaN(rateVal)) throw new Error('#VALUE!');
                const flatValues = values.flatMap(arg => rangeValues[arg] || [arg]).map(val => parseFloat(val)).filter(val => !isNaN(val));
                return flatValues.reduce((npv, val, i) => npv + val / Math.pow(1 + rateVal, i + 1), 0);
            },
            RAND: () => Math.random(),
            AND: args => {
                console.log(`AND args: ${JSON.stringify(args)}`);
                return args.every(arg => {
                    const cond = sanitizeCondition(arg);
                    return evaluateCondition(cond, row, col);
                }) ? 'TRUE' : 'FALSE';
            },
            OR: args => {
                console.log(`OR args: ${JSON.stringify(args)}`);
                return args.some(arg => {
                    const cond = sanitizeCondition(arg);
                    return evaluateCondition(cond, row, col);
                }) ? 'TRUE' : 'FALSE';
            },
            XOR: args => {
                console.log(`XOR args: ${JSON.stringify(args)}`);
                const trueCount = args.filter(arg => {
                    const cond = sanitizeCondition(arg);
                    return evaluateCondition(cond, row, col);
                }).length;
                return (trueCount % 2 === 1) ? 'TRUE' : 'FALSE';
            },
            IF: args => {
                if (args.length !== 3) throw new Error('#VALUE!');
                console.log(`IF args: ${JSON.stringify(args)}`);
                const [condition, trueVal, falseVal] = args;
                const cond = sanitizeCondition(condition);
                return evaluateCondition(cond, row, col) ? trueVal.replace(/"/g, '') : falseVal.replace(/"/g, '');
            },
            IFS: args => {
                if (args.length % 2 !== 0) throw new Error('#VALUE!');
                console.log(`IFS args: ${JSON.stringify(args)}`);
                for (let i = 0; i < args.length; i += 2) {
                    const condition = sanitizeCondition(args[i]);
                    if (evaluateCondition(condition, row, col)) {
                        return args[i + 1].replace(/"/g, '');
                    }
                }
                return '#N/A';
            },
            CONCAT: args => {
                console.log(`CONCAT args: ${JSON.stringify(args)}`);
                return args.map(arg => arg.replace(/"/g, '')).join('');
            },
            LEFT: args => {
                if (args.length < 1 || args.length > 2) throw new Error('#VALUE!');
                console.log(`LEFT args: ${JSON.stringify(args)}`);
                const [text, numChars] = args;
                const n = parseInt(numChars) || 1;
                if (n < 0) return '';
                return text.replace(/"/g, '').slice(0, n);
            },
            RIGHT: args => {
                if (args.length < 1 || args.length > 2) throw new Error('#VALUE!');
                console.log(`RIGHT args: ${JSON.stringify(args)}`);
                const [text, numChars] = args;
                const n = parseInt(numChars) || 1;
                if (n < 0) return '';
                return text.replace(/"/g, '').slice(-n);
            },
            LOWER: args => {
                if (args.length !== 1) throw new Error('#VALUE!');
                console.log(`LOWER args: ${JSON.stringify(args)}`);
                return args[0].replace(/"/g, '').toLowerCase();
            },
            TRIM: args => {
                if (args.length !== 1) throw new Error('#VALUE!');
                console.log(`TRIM raw arg: ${args[0]}`);
                let text = args[0];
                if (text.startsWith('"') && text.endsWith('"')) {
                    text = text.slice(1, -1).replace(/""/g, '"');
                }
                console.log(`TRIM processed text: ${text}`);
                return text.trim().replace(/\s+/g, ' ');
            }
        };

        // Map function names with dots to internal names
        const functionMap = {
            'STDEV.P': 'STDEV_P',
            'STDEV.S': 'STDEV_S'
        };

        // Parse function calls with support for quotes and conditions
        const parseFunctionCalls = (expression) => {
            const match = expression.match(/(\w+(\.\w+)?)\((.*)\)$/i);
            if (!match) throw new Error('#SYNTAX!');
            let funcName = match[1].toUpperCase();
            console.log(`Parsing function: ${funcName}`);
            funcName = functionMap[funcName] || funcName;
            if (!functions[funcName]) throw new Error('#NAME?');
            let argStr = match[3].trim();
            if (!argStr) return { result: functions[funcName]([]), rest: '' };

            const args = [];
            let currentArg = '';
            let inQuotes = false;
            let depth = 0;

            for (let i = 0; i < argStr.length; i++) {
                const char = argStr[i];
                if (char === '"' && argStr[i - 1] !== '\\') {
                    inQuotes = !inQuotes;
                    currentArg += char;
                } else if (char === '(' && !inQuotes) {
                    depth++;
                    currentArg += char;
                } else if (char === ')' && !inQuotes) {
                    depth--;
                    if (depth < 0) {
                        if (currentArg.trim()) args.push(currentArg.trim());
                        console.log(`Parsed ${funcName} args: ${JSON.stringify(args)}`);
                        return { result: functions[funcName](args), rest: argStr.slice(i + 1) };
                    }
                    currentArg += char;
                } else if (char === ',' && !inQuotes && depth === 0) {
                    if (currentArg.trim()) args.push(currentArg.trim());
                    currentArg = '';
                } else {
                    currentArg += char;
                }
            }
            if (currentArg.trim()) args.push(currentArg.trim());

            console.log(`Parsed ${funcName} args: ${JSON.stringify(args)}`);
            return { result: functions[funcName](args), rest: '' };
        };

        // Replace function calls
        while (expr.match(/(\w+(\.\w+)?)\([^)]+\)/i)) {
            const match = expr.match(/(\w+(\.\w+)?)\([^)]+\)/i);
            const funcExpr = match[0];
            const { result } = parseFunctionCalls(funcExpr);
            expr = expr.replace(funcExpr, result);
        }

        return eval(expr);

    } catch (error) {
        console.error(`Formula error at (${row}, ${col}): ${error.message}`);
        return error.message.startsWith('#') ? error.message : '#ERROR!';
    }
}

// Update dependent cells
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

// Adjust formula references for filling
function adjustFormula(formula, srcRow, srcCol, destRow, destCol) {
    console.log(`Adjusting formula "${formula}" from (${srcRow}, ${srcCol}) to (${destRow}, ${destCol})`);
    if (!formula.startsWith('=')) return formula;

    let adjusted = formula;
    const refRegex = /[A-Z]+\d+(?::[A-Z]+\d+)?/g;
    const matches = formula.match(refRegex) || [];

    for (let ref of matches) {
        if (ref.includes(':')) {
            // Handle range (e.g., A1:A5)
            const [start, end] = ref.split(':');
            const startRef = parseCellReference(start);
            const endRef = parseCellReference(end);
            if (!startRef || !endRef) continue;

            const rowDelta = destRow - srcRow;
            const colDelta = destCol - srcCol;

            const newStart = toCellReference(startRef.row + rowDelta, startRef.col + colDelta);
            const newEnd = toCellReference(endRef.row + rowDelta, endRef.col + colDelta);
            const newRange = `${newStart}:${newEnd}`;
            console.log(`Adjusted range ${ref} to ${newRange}`);
            adjusted = adjusted.replace(ref, newRange);
        } else {
            // Handle single cell
            const cellRef = parseCellReference(ref);
            if (!cellRef) continue;

            const rowDelta = destRow - srcRow;
            const colDelta = destCol - srcCol;
            const newRef = toCellReference(cellRef.row + rowDelta, cellRef.col + colDelta);
            console.log(`Adjusted cell ${ref} to ${newRef}`);
            adjusted = adjusted.replace(ref, newRef);
        }
    }

    return adjusted;
}

// Load sample data
function loadSampleData(key) {
    console.log(`Loading sample data: ${key}`);
    const csvData = sampleData[key];
    if (csvData) {
        processCSVData(csvData);
    } else {
        alert('No sample data found for key: ' + key);
    }
}

// Handle file selection
function handleFileSelect(event) {
    console.log('File selected:', event.target.files[0]?.name);
    const file = event.target.files[0];
    if (!file) return;

    loading.style.display = 'block';
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        if (file.name.endsWith('.csv')) {
            processCSVData(data);
        } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            processExcelData(data);
        }
        loading.style.display = 'none';
    };
    
    if (file.name.endsWith('.csv')) {
        reader.readAsText(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

// Process CSV data
function processCSVData(csvString) {
    console.log('Processing CSV data');
    try {
        const workbook = XLSX.read(csvString, { type: 'string' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        if (data.length === 0) {
            console.warn('No data in CSV');
            alert('No data found in the file');
            return;
        }
        
        headers = data[0] || [];
        currentData = data.slice(1) || [];
        formulas = [];
        dependencies = {};
        
        renderSpreadsheet();
        updateSQLDatabase();
        updateChartDropdowns();
    } catch (error) {
        console.error('Error processing CSV:', error);
        alert('Error processing CSV file');
    }
}

// Process Excel data
function processExcelData(arrayBuffer) {
    console.log('Processing Excel data');
    try {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        if (data.length === 0) {
            console.warn('No data in Excel');
            alert('No data found in the file');
            return;
        }
        
        headers = data[0] || [];
        currentData = data.slice(1) || [];
        formulas = [];
        dependencies = {};
        
        renderSpreadsheet();
        updateSQLDatabase();
        updateChartDropdowns();
    } catch (error) {
        console.error('Error processing Excel:', error);
        alert('Error processing Excel file');
    }
}

// Render spreadsheet
function renderSpreadsheet() {
    console.log('Rendering spreadsheet');
    headerRow.innerHTML = '<th class="row-number-header"></th>';
    dataBody.innerHTML = '';

    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header;
        th.contentEditable = true;
        th.dataset.column = index;
        th.addEventListener('click', function(event) {
            console.log(`Column header clicked: ${index}`);
            event.stopPropagation();
            handleColumnSelection(event);
        });
        th.addEventListener('blur', () => {
            headers[index] = th.textContent.trim() || `Column${index + 1}`;
            console.log(`Header updated: ${headers[index]}`);
            updateSQLDatabase();
            updateChartDropdowns();
        });
        headerRow.appendChild(th);
    });

    currentData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');

        const rowNumberTd = document.createElement('td');
        rowNumberTd.className = 'row-number';
        rowNumberTd.textContent = rowIndex + 1;
        rowNumberTd.dataset.row = rowIndex;
        rowNumberTd.addEventListener('click', function(event) {
            console.log(`Row number clicked: ${rowIndex}`);
            event.stopPropagation();
            handleRowSelection(event);
        });
        tr.appendChild(rowNumberTd);

        row.forEach((cell, colIndex) => {
            const td = document.createElement('td');
            td.textContent = cell || '';
            td.contentEditable = true;
            td.dataset.row = rowIndex;
            td.dataset.column = colIndex;
            td.addEventListener('click', function(event) {
                console.log(`Cell clicked: row ${rowIndex}, col ${colIndex}`);
                event.stopPropagation();
                handleCellSelection(event);
                formulaBar.value = formulas[rowIndex]?.[colIndex] || td.textContent;
            });
            td.addEventListener('blur', () => {
                const value = td.textContent.trim();
                console.log(`Cell updated: row ${rowIndex}, col ${colIndex}, value ${value}`);
                updateCellValue(rowIndex, colIndex, value);
            });
            tr.appendChild(td);
        });

        while (tr.children.length - 1 < headers.length) {
            const td = document.createElement('td');
            td.contentEditable = true;
            td.dataset.row = rowIndex;
            td.dataset.column = tr.children.length - 1;
            td.addEventListener('click', function(event) {
                console.log(`Empty cell clicked: row ${rowIndex}, col ${tr.children.length - 1}`);
                event.stopPropagation();
                handleCellSelection(event);
                formulaBar.value = '';
            });
            td.addEventListener('blur', () => {
                const value = td.textContent.trim();
                if (!currentData[rowIndex]) currentData[rowIndex] = [];
                console.log(`Empty cell updated: row ${rowIndex}, col ${tr.children.length - 1}`);
                updateCellValue(rowIndex, tr.children.length - 1, value);
            });
            tr.appendChild(td);
        }

        dataBody.appendChild(tr);
    });

    // Re-apply selected cells highlighting after rendering
    selectedCells.forEach(key => {
        const [row, col] = key.split('-').map(Number);
        const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (cell) cell.classList.add('selected-cell');
    });

    // Add fill handle to selected cell
    updateFillHandle();
    console.log('Spreadsheet rendered, headers:', headers, 'data:', currentData);
}

// Update fill handle visibility
// Update fill handle visibility
function updateFillHandle() {
    document.querySelectorAll('.fill-handle').forEach(handle => handle.remove());
    
    if (selectedCells.length >= 1) {
        let targetCellKey;
        
        // Convert selectedCells to array of { row, col }
        const cells = selectedCells.map(key => {
            const [row, col] = key.split('-').map(Number);
            return { row, col };
        });
        
        // Determine if the selection is vertical (same column) or horizontal (same row)
        const isVertical = cells.every(cell => cell.col === cells[0].col);
        const isHorizontal = cells.every(cell => cell.row === cells[0].row);
        
        if (selectedCells.length === 1) {
            // Single cell case
            targetCellKey = selectedCells[0];
        } else if (selectedCells.length === 2 && isVertical) {
            // Two vertical cells: select the cell with the higher row index (bottom cell)
            const [cell1, cell2] = cells;
            targetCellKey = cell1.row > cell2.row ? `${cell1.row}-${cell1.col}` : `${cell2.row}-${cell2.col}`;
        } else if (isVertical) {
            // Multiple vertical cells: select the cell with the highest row index (bottom cell)
            const maxRowCell = cells.reduce((max, cell) => cell.row > max.row ? cell : max, cells[0]);
            targetCellKey = `${maxRowCell.row}-${maxRowCell.col}`;
        } else if (isHorizontal) {
            // Multiple horizontal cells: select the cell with the highest column index (rightmost cell)
            const maxColCell = cells.reduce((max, cell) => cell.col > max.col ? cell : max, cells[0]);
            targetCellKey = `${maxColCell.row}-${maxColCell.col}`;
        } else {
            // Non-linear selection: fallback to the last cell in selectedCells
            targetCellKey = selectedCells[selectedCells.length - 1];
        }
        
        const [row, col] = targetCellKey.split('-').map(Number);
        const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (cell) {
            const handle = document.createElement('div');
            handle.className = 'fill-handle';
            handle.addEventListener('mousedown', startFillDrag);
            handle.addEventListener('dblclick', handleDoubleClickFill);
            cell.appendChild(handle);
            console.log('Fill handle added to cell:', row, col);
        }
    }
}

// Handle double-click fill
function handleDoubleClickFill(event) {
    event.stopPropagation();
    if (selectedCells.length !== 1) return;
    const [row, col] = selectedCells[0].split('-').map(Number);
    console.log('Double-click fill started from cell:', row, col);
    
    let endRow = row;
    const adjacentCol = col - 1;
    
    // If there's an adjacent column and it has data, extend fill range
    if (adjacentCol >= 0) {
        while (
            endRow + 1 < currentData.length &&
            currentData[endRow + 1] !== undefined &&
            currentData[endRow + 1][adjacentCol] != null &&
            currentData[endRow + 1][adjacentCol].toString() !== ''
        ) {
            endRow++;
        }
    } else {
        // If no adjacent column, fill one row down
        endRow = row + 1 < currentData.length ? row + 1 : row;
    }
    
    // Ensure at least one cell is filled if possible
    if (endRow === row && endRow + 1 < currentData.length) {
        endRow = row + 1;
    }
    
    // Clear previous fill range highlighting
    document.querySelectorAll('.fill-range').forEach(cell => cell.classList.remove('fill-range'));
    
    fillStartCell = [row, col];
    fillRange = getFillRange(row, col, endRow, col);
    
    // Highlight fill range
    fillRange.forEach(([r, c]) => {
        const cell = document.querySelector(`td[data-row="${r}"][data-column="${c}"]`);
        if (cell && !cell.classList.contains('row-number')) {
            cell.classList.add('fill-range');
        }
    });
    
    // Apply fill
    applyFill();
    
    // Remove highlighting after short delay
    setTimeout(() => {
        document.querySelectorAll('.fill-range').forEach(cell => cell.classList.remove('fill-range'));
    }, 500);
    
    console.log('Double-click fill applied to rows:', row, 'to', endRow, 'range:', fillRange);
}

// Start fill drag
function startFillDrag(event) {
    event.preventDefault();
    event.stopPropagation();
    isFilling = true;
    fillStartCell = selectedCells[0].split('-').map(Number);
    fillRange = [];
    console.log('Fill drag started from cell:', fillStartCell);
    
    document.addEventListener('mousemove', handleFillDrag);
    document.addEventListener('mouseup', endFillDrag, { once: true });
}

// Handle fill drag
function handleFillDrag(event) {
    if (!isFilling) return;
    
    const target = event.target.closest('td[data-row][data-column]');
    if (!target || target.classList.contains('row-number')) return;
    
    const endRow = parseInt(target.dataset.row);
    const endCol = parseInt(target.dataset.column);
    
    document.querySelectorAll('.fill-range').forEach(cell => cell.classList.remove('fill-range'));
    
    fillRange = getFillRange(fillStartCell[0], fillStartCell[1], endRow, endCol);
    
    fillRange.forEach(([row, col]) => {
        const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (cell && !cell.classList.contains('row-number')) {
            cell.classList.add('fill-range');
        }
    });
    
    console.log('Fill range updated:', fillRange);
}

// End fill drag
function endFillDrag(event) {
    if (!isFilling) return;
    isFilling = false;
    
    document.removeEventListener('mousemove', handleFillDrag);
    
    document.querySelectorAll('.fill-range').forEach(cell => cell.classList.remove('fill-range'));
    
    if (fillRange.length > 0) {
        applyFill();
        
        // Update selectedCells to include the filled range
        clearSelections();
        selectedCells = fillRange.map(([row, col]) => `${row}-${col}`);
        selectedCells.forEach(key => {
            const [row, col] = key.split('-').map(Number);
            const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
            if (cell) cell.classList.add('selected-cell');
        });
        
        // Update fill handle to appear on the last cell (bottom-right)
        updateFillHandle();
    }
    
    console.log('Fill drag ended, range:', fillRange, 'new selectedCells:', selectedCells);
}

// Get fill range coordinates
function getFillRange(startRow, startCol, endRow, endCol) {
    const range = [];
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);
    
    if (startRow === endRow) {
        for (let col = minCol; col <= maxCol; col++) {
            range.push([startRow, col]);
        }
    } else if (startCol === endCol) {
        for (let row = minRow; row <= maxRow; row++) {
            range.push([row, startCol]);
        }
    }
    
    return range;
}

// Detect pattern for autofill
function detectPattern() {
    console.log('Detecting pattern for fillStartCell:', fillStartCell, 'fillRange:', fillRange);

    if (!fillStartCell) return { type: 'text', base: '', step: 0 };

    const [startRow, startCol] = fillStartCell;
    const startValue = formulas[startRow]?.[startCol] || currentData[startRow]?.[startCol] || '';

    // Date pattern
    const dateMatch = startValue.toString().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (dateMatch) {
        const [_, day, month, year] = dateMatch;
        const paddedDay = day.padStart(2, '0');
        const paddedMonth = month.padStart(2, '0');
        const date = new Date(`${year}-${paddedMonth}-${paddedDay}`);
        if (!isNaN(date.getTime())) {
            return { type: 'date', base: date, inputFormat: { day: day.length === 1, month: month.length === 1 }, step: 1 };
        }
    }

    // Text + number pattern
    const textNumberMatch = startValue.toString().match(/^(.+?)(\d+)$/);
    if (textNumberMatch && isNaN(parseFloat(textNumberMatch[1]))) {
        return { type: 'text-number', prefix: textNumberMatch[1], number: parseInt(textNumberMatch[2]), step: 1 };
    }

    // Numeric pattern
    const num = parseFloat(startValue);
    if (!isNaN(num)) {
        if (selectedCells.length >= 2) {
            // Sort selected cells by row or column to determine direction
            const sortedCells = selectedCells
                .map(key => {
                    const [row, col] = key.split('-').map(Number);
                    return { row, col, value: parseFloat(currentData[row]?.[col]) };
                })
                .filter(cell => !isNaN(cell.value));

            // Determine if cells are in a single row or column
            const isHorizontal = sortedCells.every(cell => cell.row === startRow);
            const isVertical = sortedCells.every(cell => cell.col === startCol);

            if (!isHorizontal && !isVertical) {
                console.log('Selected cells are not in a single row or column: copy mode');
                return { type: 'number', base: num, step: 0 };
            }

            // Sort cells by position (left-to-right or top-to-bottom)
            sortedCells.sort((a, b) => {
                if (isHorizontal) return a.col - b.col;
                return a.row - b.row;
            });

            // Calculate average step (arithmetic sequence)
            let step = 0;
            let validSequence = true;
            const values = sortedCells.map(cell => cell.value);
            if (values.length >= 2) {
                const steps = [];
                for (let i = 1; i < values.length; i++) {
                    steps.push(values[i] - values[i - 1]);
                }
                // Check if steps are consistent (within a small tolerance for floating-point)
                const avgStep = steps.reduce((sum, s) => sum + s, 0) / steps.length;
                validSequence = steps.every(s => Math.abs(s - avgStep) < 0.0001);
                step = validSequence ? avgStep : 0;
            }

            if (validSequence) {
                console.log(`Multi-cell numeric sequence detected: values=${values}, step=${step}, direction=${isHorizontal ? 'horizontal' : 'vertical'}`);
                return {
                    type: 'number',
                    base: sortedCells[0].value, // Use first cell's value as base
                    step: step,
                    direction: isHorizontal ? 'horizontal' : 'vertical',
                    selectedCells: sortedCells.map(cell => [cell.row, cell.col])
                };
            } else {
                console.log('Inconsistent numeric sequence: copy mode');
                return { type: 'number', base: num, step: 0 };
            }
        }
        // Single cell: copy mode
        console.log('Single cell numeric selection: copy mode');
        return { type: 'number', base: num, step: 0 };
    }

    // Text pattern
    return { type: 'text', base: startValue.toString(), step: 0 };
}

// Get next value for autofill
function getNextValue(original, pattern, stepIndex) {
    console.log('Getting next value for pattern:', pattern, 'stepIndex:', stepIndex);
    switch (pattern.type) {
        case 'date':
            const nextDate = new Date(pattern.base);
            nextDate.setDate(nextDate.getDate() + stepIndex * pattern.step);
            const day = pattern.inputFormat.day ? nextDate.getDate() : String(nextDate.getDate()).padStart(2, '0');
            const month = pattern.inputFormat.month ? nextDate.getMonth() + 1 : String(nextDate.getMonth() + 1).padStart(2, '0');
            const year = nextDate.getFullYear();
            return `${day}.${month}.${year}`;
        case 'text-number':
            return `${pattern.prefix}${pattern.number + stepIndex * pattern.step}`;
        case 'number':
            if (pattern.step === 0) {
                return pattern.base.toString(); // Copy mode
            } else {
                // Start series from the base value, applying step for each index
                return (pattern.base + stepIndex * pattern.step).toString();
            }
        case 'text':
            return pattern.base;
        default:
            return original;
    }
}

// Apply fill logic
function applyFill() {
    if (!fillStartCell || fillRange.length === 0) return;

    const [startRow, startCol] = fillStartCell;
    const cell = document.querySelector(`td[data-row="${startRow}"][data-column="${startCol}"]`);
    if (!cell) return;

    const pattern = detectPattern();
    const sourceValue = formulas[startRow]?.[startCol] || cell.textContent.trim();
    console.log('Applying fill from cell:', startRow, startCol, 'value:', sourceValue, 'pattern:', pattern);

    // Determine fill direction
    let isHorizontalFill = true;
    if (pattern.direction) {
        isHorizontalFill = pattern.direction === 'horizontal';
    } else if (fillRange.length > 0) {
        isHorizontalFill = fillRange.every(([r, c]) => r === startRow);
    }

    // Get selected cells to skip
    const selectedCellSet = new Set((pattern.selectedCells || [[startRow, startCol]]).map(([r, c]) => `${r}-${c}`));

    fillRange.forEach(([row, col], index) => {
        // Skip selected cells to preserve their values
        const cellKey = `${row}-${col}`;
        if (selectedCellSet.has(cellKey)) return;

        const targetCell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (targetCell) {
            let newValue;
            if (sourceValue.startsWith('=')) {
                // Adjust formula references
                newValue = adjustFormula(sourceValue, startRow, startCol, row, col);
                formulas[row] = formulas[row] || [];
                formulas[row][col] = newValue;
                newValue = evaluateFormula(newValue, row, col);
                console.log(`Filled formula at (${row}, ${col}) with index ${index}: ${newValue}`);
            } else {
                // Calculate step index based on direction
                let stepIndex;
                if (isHorizontalFill) {
                    stepIndex = col - startCol;
                } else {
                    stepIndex = row - startRow;
                }
                // Adjust step index to start after the last selected cell
                if (pattern.selectedCells) {
                    const lastSelected = pattern.selectedCells[pattern.selectedCells.length - 1];
                    stepIndex = isHorizontalFill
                        ? col - lastSelected[1] + pattern.selectedCells.length - 1
                        : row - lastSelected[0] + pattern.selectedCells.length - 1;
                }
                newValue = getNextValue(sourceValue, pattern, stepIndex);
                if (formulas[row]?.[col]) delete formulas[row][col];
                console.log(`Filled value at (${row}, ${col}) with index ${stepIndex}: ${newValue}`);
            }

            targetCell.textContent = newValue;
            currentData[row] = currentData[row] || [];
            currentData[row][col] = newValue;
            updateDependents(row, col);
        }
    });

    updateSQLDatabase();
    renderSpreadsheet();
}

function handleCellSelection(event) {
    if (window.fillSequenceTestRunning && window.fillSequenceTestSelecting) {
        console.log('handleCellSelection skipped due to fillSequenceTestSelecting');
        return;
    }
    const td = event.target.closest('td[data-row][data-column]');
    if (!td) {
        console.error('No valid cell found for click:', event.target);
        return;
    }
    const rowIndex = parseInt(td.dataset.row);
    const colIndex = parseInt(td.dataset.column);
    if (isNaN(rowIndex) || isNaN(colIndex)) {
        console.error('Invalid cell coordinates:', td.dataset);
        return;
    }
    const cellKey = `${rowIndex}-${colIndex}`;
    event.stopPropagation();

    console.log(`Handling cell selection: row=${rowIndex}, col=${colIndex}, cellKey=${cellKey}`);

    if (event.ctrlKey && event.buttons === 1) {
        isCtrlSelecting = true;
        addCellToSelection(td, cellKey);
        formulaBar.value = formulas[rowIndex]?.[colIndex] || td.textContent || '';
        updateFillHandle();
        return;
    }

    if (event.shiftKey && selectedCells.length > 0) {
        const lastCell = selectedCells[selectedCells.length - 1];
        const [lastRow, lastCol] = lastCell.split('-').map(Number);
        const startRow = Math.min(lastRow, rowIndex);
        const endRow = Math.max(lastRow, rowIndex);
        const startCol = Math.min(lastCol, colIndex);
        const endCol = Math.max(lastCol, colIndex);
        clearSelections();
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const key = `${r}-${c}`;
                const cell = document.querySelector(`td[data-row="${r}"][data-column="${c}"]`);
                if (cell) {
                    selectedCells.push(key);
                    cell.classList.add('selected-cell');
                }
            }
        }
    } else if (event.ctrlKey || event.metaKey) {
        const index = selectedCells.indexOf(cellKey);
        if (index === -1) {
            selectedCells.push(cellKey);
            td.classList.add('selected-cell');
        } else {
            selectedCells.splice(index, 1);
            td.classList.remove('selected-cell');
        }
    } else {
        clearSelections();
        selectedCells = [cellKey];
        td.classList.add('selected-cell');
    }

    window.selectedCells = selectedCells;
    formulaBar.value = formulas[rowIndex]?.[colIndex] || td.textContent || '';
    updateFillHandle();
    console.log('Cell selected:', cellKey, 'selectedCells:', selectedCells);
}

// Handle drag selection
function handleDragSelection(startRow, startCol, endRow, endCol, isCtrl) {
    console.log(`Dragging selection from (${startRow}, ${startCol}) to (${endRow}, ${endCol})`);
    
    if (!isCtrl) {
        clearSelections();
    }

    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);

    const newSelectedCells = [];
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            const key = `${r}-${c}`;
            if (!newSelectedCells.includes(key) && !selectedCells.includes(key)) {
                newSelectedCells.push(key);
                const cell = document.querySelector(`td[data-row="${r}"][data-column="${c}"]`);
                if (cell) cell.classList.add('selected-cell');
            }
        }
    }

    if (isCtrl) {
        selectedCells = [...new Set([...selectedCells, ...newSelectedCells])];
    } else {
        selectedCells = newSelectedCells;
    }

    // Update formula bar for single cell selection
    formulaBar.value = selectedCells.length === 1 ? 
        (formulas[startRow]?.[startCol] || document.querySelector(`td[data-row="${startRow}"][data-column="${startCol}"]`)?.textContent || '') : 
        '';
    
    updateFillHandle();
}

// Add cell to selection
function addCellToSelection(td, cellKey) {
    if (!selectedCells.includes(cellKey)) {
        selectedCells.push(cellKey);
        td.classList.add('selected-cell');
    }
}

function handleRowSelection(event) {
    const rowNumberTd = event.target;
    const rowIndex = parseInt(rowNumberTd.dataset.row);

    console.log('Row clicked:', { rowIndex });

    if (isNaN(rowIndex)) {
        console.error('Invalid row index:', rowNumberTd.dataset);
        return;
    }

    clearSelections();
    selectedRow = rowIndex;
    const row = dataBody.children[rowIndex];
    if (row) {
        row.classList.add('selected-row');
        console.log('Selected row:', rowIndex);
    } else {
        console.error('Row not found:', rowIndex);
    }
}

function handleColumnSelection(event) {
    const th = event.target;
    const colIndex = parseInt(th.dataset.column);

    console.log('Column clicked:', { colIndex });

    if (isNaN(colIndex)) {
        console.error('Invalid column index:', th.dataset);
        return;
    }

    clearSelections();
    selectedColumn = colIndex;
    const cells = document.querySelectorAll(`#spreadsheet td[data-column="${colIndex}"], #spreadsheet th[data-column="${colIndex}"]`);
    cells.forEach(cell => cell.classList.add('selected-column'));
    console.log('Selected column:', colIndex, 'Cells affected:', cells.length);
}
function clearSelections() {
    if (window.fillSequenceTestRunning && window.fillSequenceTestSelecting) {
        console.log('clearSelections skipped due to fillSequenceTestSelecting');
        return;
    }
    console.log('Clearing selections:', { selectedCells, selectedRow, selectedColumn });

    selectedCells.forEach(cellKey => {
        const [row, col] = cellKey.split('-').map(Number);
        const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        if (cell) {
            cell.classList.remove('selected-cell');
        }
    });
    selectedCells = [];
    window.selectedCells = selectedCells;

    if (selectedRow !== null) {
        const row = dataBody.children[selectedRow];
        if (row) {
            row.classList.remove('selected-row');
        }
        selectedRow = null;
    }

    if (selectedColumn !== null) {
        const cells = document.querySelectorAll(`#spreadsheet td[data-column="${selectedColumn}"], #spreadsheet th[data-column="${selectedColumn}"]`);
        cells.forEach(cell => cell.classList.remove('selected-column'));
        selectedColumn = null;
    }

    formulaBar.value = '';
    updateFillHandle();
    console.log('Selections cleared');
}
// Update SQL database
function updateSQLDatabase() {
    if (!SQL || !db) {
        console.warn('SQL.js not ready');
        return;
    }
    
    try {
        db.run('DROP TABLE IF EXISTS data;');
        
        const sanitizedHeaders = headers.map((h, i) => {
            if (!h || h.trim() === '') return `Column${i + 1}`;
            return `"${h.replace(/[^a-zA-Z0-9_]/g, '_')}"`;
        });
        
        const createTableSQL = `CREATE TABLE data (${sanitizedHeaders.join(', ')} TEXT);`;
        console.log('Executing CREATE TABLE:', createTableSQL);
        db.run(createTableSQL);
        
        const placeholders = headers.map(() => '?').join(', ');
        const stmt = db.prepare(`INSERT INTO data VALUES (${placeholders});`);
        
        currentData.forEach((row, rowIndex) => {
            const normalizedRow = headers.map((_, i) => {
                let value = row[i] || '';
                if (typeof value === 'number' || value.startsWith('#')) {
                    return value.toString();
                }
                return value.replace(/"/g, '""');
            });
            console.log(`Inserting row ${rowIndex}:`, normalizedRow);
            stmt.run(normalizedRow);
        });
        
        stmt.free();
        console.log('SQL database updated');
    } catch (error) {
        console.error('SQL Error:', error.message, 'SQL:', createTableSQL || 'N/A');
        alert(`SQL Error: ${error.message}`);
    }
}

// Update chart dropdowns
function updateChartDropdowns() {
    console.log('Updating chart dropdowns');
    xAxisSelect.innerHTML = '';
    yAxisSelect.innerHTML = '';
    
    headers.forEach(header => {
        const option1 = document.createElement('option');
        option1.value = header;
        option1.textContent = header;
        xAxisSelect.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = header;
        option2.textContent = header;
        yAxisSelect.appendChild(option2);
    });
    
    if (headers.length >= 1) xAxisSelect.value = headers[0];
    if (headers.length >= 2) yAxisSelect.value = headers[1];
    console.log('Chart dropdowns updated');
}

// Run SQL query
function runSQL() {
    console.log('Running SQL query');
    if (!SQL || !db) {
        alert('SQL.js is not loaded yet. Please wait a moment and try again.');
        return;
    }
    
    const query = sqlQuery.value.trim();
    if (!query) {
        alert('Please enter a SQL query');
        return;
    }
    
    try {
        const result = db.exec(query);
        
        if (result.length === 0) {
            alert('Query executed successfully but returned no results.');
            return;
        }
        
        const columns = result[0].columns;
        const values = result[0].values;
        
        headers = columns;
        currentData = values.map(row => row.map(cell => cell === null ? '' : cell.toString()));
        
        formulas = [];
        dependencies = {};
        
        renderSpreadsheet();
        updateChartDropdowns();
        console.log('SQL query executed, data updated');
    } catch (error) {
        console.error('SQL Error:', error);
        alert(`SQL Error: ${error.message}`);
    }
}

// Create chart
// ...existing code...
function createChart() {
    console.log('Creating chart');
    if (headers.length === 0 || currentData.length === 0) {
        alert('No data available to create chart');
        return;
    }

    let chartType = chartTypeSelect.value;
    let xAxis = xAxisSelect.value;
    let yAxis = yAxisSelect.value;

    let xIndex = headers.indexOf(xAxis);
    let yIndex = headers.indexOf(yAxis);

    if (xIndex === -1 || yIndex === -1) {
        alert('Invalid axis selection');
        return;
    }

    // Check if Y-axis is numeric
    let yValues = currentData.map(row => row[yIndex]);
    let yIsNumeric = yValues.every(val => !isNaN(parseFloat(val)));

    if (!yIsNumeric) {
        // If not numeric, swap axes and use horizontal bar chart
        [xAxis, yAxis] = [yAxis, xAxis];
        [xIndex, yIndex] = [yIndex, xIndex];
        yValues = currentData.map(row => row[yIndex]);
        chartType = 'bar'; // Chart.js v3+ uses 'bar' with indexAxis option for horizontal
        var indexAxis = 'y';
    } else {
        var indexAxis = 'x';
    }

    // Now, x-axis is always categories, y-axis is always numeric
    const labels = currentData.map(row => row[xIndex]);
    const dataValues = currentData.map(row => {
        const val = row[yIndex];
        return isNaN(parseFloat(val)) ? 0 : parseFloat(val);
    });

    const config = {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: `${yAxis} by ${xAxis}`,
                data: dataValues,
                backgroundColor: getChartColors(chartType, dataValues.length),
                borderColor: '#4CAF50',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            indexAxis: indexAxis,
            plugins: {
                title: {
                    display: true,
                    text: `${yAxis} by ${xAxis}`
                }
            }
        }
    };

    // ...existing code for chart creation...

// ...existing code...
    
    const chartId = `chart-${Date.now()}`;
    const chartItem = document.createElement('div');
    chartItem.className = 'chart-item';
    chartItem.innerHTML = `
        <canvas id="${chartId}"></canvas>
        <div class="chart-item-actions">
            <button class="delete-chart" data-id="${chartId}">Delete</button>
            <button class="download-chart" data-id="${chartId}">Download</button>
        </div>
    `;
    
    chartList.appendChild(chartItem);
    
    const chartCtx = document.getElementById(chartId).getContext('2d');
    const chart = new Chart(chartCtx, config);
    
    charts.push({
        id: chartId,
        chart: chart,
        type: chartType,
        xAxis: xAxis,
        yAxis: yAxis
    });
    
    chartItem.querySelector('.delete-chart').addEventListener('click', function() {
        console.log(`Deleting chart: ${this.dataset.id}`);
        deleteChart(this.dataset.id);
    });
    
    chartItem.querySelector('.download-chart').addEventListener('click', function() {
        console.log(`Downloading chart: ${this.dataset.id}`);
        downloadChart(this.dataset.id);
    });
    console.log('Chart created:', chartId);
}

// Generate colors for charts
function getChartColors(type, count) {
    return Array(count).fill().map((_, i) => {
        const hue = (i * 360 / count) % 360;
        return `hsl(${hue}, 70%, 60%)`;
    });
}

// Delete chart
function deleteChart(id) {
    console.log('Deleting chart:', id);
    const index = charts.findIndex(chart => chart.id === id);
    if (index !== -1) {
        charts[index].chart.destroy();
        charts.splice(index, 1);
        const chartItem = document.getElementById(id)?.parentElement;
        if (chartItem) chartItem.remove();
    }
}

// Download chart
function downloadChart(id) {
    console.log('Downloading chart:', id);
    const chart = charts.find(chart => chart.id === id);
    if (!chart) return;
    
    const link = document.createElement('a');
    link.download = `chart-${chart.type}-${Date.now()}.png`;
    link.href = chart.chart.canvas.toDataURL('image/png');
    link.click();
}

// Export as CSV
function exportCSV() {
    console.log('Exporting as CSV');
    if (headers.length === 0) {
        alert('No data to export');
        return;
    }
    
    let csvContent = headers.join(',') + '\n';
    csvContent += currentData.map(row => 
        row.map(cell => 
            `"${(cell || '').toString().replace(/"/g, '""')}"`
        ).join(',')
    ).join('\n');
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'data_export.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Export as Excel
function exportExcel() {
    console.log('Exporting as Excel');
    if (headers.length === 0) {
        alert('No data to export');
        return;
    }
    
    const exportData = [headers, ...currentData];
    
    const ws = XLSX.utils.aoa_to_sheet(exportData);
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    XLSX.writeFile(wb, 'data_export.xlsx');
}

// Add column
function addColumn() {
    console.log('Adding column');
    const colName = `Column${headers.length + 1}`;
    headers.push(colName);
    
    currentData.forEach(row => row.push(''));
    formulas.forEach(row => row?.push(undefined));
    
    renderSpreadsheet();
    updateSQLDatabase();
    updateChartDropdowns();
}

// Add row
function addRow() {
    console.log('Adding row');
    currentData.push(headers.map(() => ''));
    formulas.push([]);
    
    renderSpreadsheet();
    updateSQLDatabase();
    updateChartDropdowns();
}

// Clear data
function clearData() {
    console.log('Clearing data');
    if (confirm('Are you sure you want to clear all data?')) {
        headers = [];
        currentData = [];
        formulas = [];
        dependencies = {};
        renderSpreadsheet();
        
        charts.forEach(chart => chart.chart.destroy());
        charts = [];
        chartList.innerHTML = '';
        
        if (db) {
            db.run('DROP TABLE IF EXISTS data;');
        }
    }
}

let isCtrlSelecting = false;

document.addEventListener('mouseup', () => {
    isCtrlSelecting = false;
});

document.addEventListener('mouseover', event => {
    if (isCtrlSelecting && event.ctrlKey && event.buttons === 1) {
        const td = event.target.closest('td[data-row][data-column]');
        if (td) {
            const rowIndex = parseInt(td.dataset.row);
            const colIndex = parseInt(td.dataset.column);
            const cellKey = `${rowIndex}-${colIndex}`;
            addCellToSelection(td, cellKey);
        }
    }
});