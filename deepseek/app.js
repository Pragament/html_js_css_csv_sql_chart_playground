// Initialize SQL.js
let SQL;
initSqlJs({
    locateFile: file => `https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/${file}`
}).then(function(sql) {
    SQL = sql;
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

// Initialize Univer
const univer = new Univer();
const universheet = univer.createUniverSheet();

// Add plugins
univer.registerPlugin(Sheets);
univer.registerPlugin(SheetsFormula);
univer.registerPlugin(SheetsNumfmt);
univer.registerPlugin(SheetsUI);
univer.registerPlugin(SheetsDataValidation);
univer.registerPlugin(Drawing);

// Create UI
const container = document.getElementById('univer-container');
const univerAPI = univer.getAPI();
const ui = univerAPI.createUI(container);

// Current data storage
let currentData = [];
let currentSheet = null;

// Event listeners
document.getElementById('import-btn').addEventListener('click', importCSV);
document.getElementById('csv-file').addEventListener('change', handleFileSelect);
document.getElementById('run-sql').addEventListener('click', runSQL);
document.getElementById('export-btn').addEventListener('click', exportCSV);
document.getElementById('create-chart').addEventListener('click', createChart);
document.getElementById('add-sheet').addEventListener('click', addSheet);

// Sample file click handlers
document.querySelectorAll('#sample-files li').forEach(item => {
    item.addEventListener('click', function() {
        const fileKey = this.getAttribute('data-file');
        loadSampleData(fileKey);
    });
});

// Load sample data
function loadSampleData(key) {
    const csvData = sampleData[key];
    if (csvData) {
        processCSVData(csvData);
    }
}

// Handle file selection
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        processCSVData(e.target.result);
    };
    reader.readAsText(file);
}

// Import CSV button handler
function importCSV() {
    document.getElementById('csv-file').click();
}

// Process CSV data
function processCSVData(csvString) {
    const rows = csvString.split('\n').filter(row => row.trim() !== '');
    const data = rows.map(row => row.split(','));
    
    currentData = data;
    
    // Create or clear the sheet
    if (!currentSheet) {
        currentSheet = universheet.addSheet('Sheet1');
    } else {
        universheet.clearSheet(currentSheet.getSheetId());
    }
    
    // Add data to the sheet
    data.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
            currentSheet.setCellValue(rowIndex, colIndex, cell);
        });
    });
    
    // Refresh the UI
    univerAPI.refreshUI();
}

// Run SQL query
function runSQL() {
    if (!SQL) {
        alert('SQL.js is not loaded yet. Please wait a moment and try again.');
        return;
    }
    
    const query = document.getElementById('sql-query').value.trim();
    if (!query) {
        alert('Please enter a SQL query');
        return;
    }
    
    try {
        // Convert current data to SQLite database
        const db = new SQL.Database();
        
        // Create table and insert data (skip header row if it exists)
        const headers = currentData[0];
        const rows = currentData.slice(1);
        
        // Create table
        const createTableSQL = `CREATE TABLE data (${headers.map(h => `${h} TEXT`).join(', ')});`;
        db.run(createTableSQL);
        
        // Insert data
        const insertStmt = db.prepare(`INSERT INTO data VALUES (${headers.map(() => '?').join(', ')});`);
        rows.forEach(row => {
            insertStmt.run(row);
        });
        insertStmt.free();
        
        // Execute query
        const result = db.exec(query);
        
        if (result.length === 0) {
            alert('Query executed successfully but returned no results.');
            return;
        }
        
        // Process results
        const columns = result[0].columns;
        const values = result[0].values;
        
        // Create new sheet for results
        const resultSheet = universheet.addSheet('SQL Results');
        
        // Add column headers
        columns.forEach((col, index) => {
            resultSheet.setCellValue(0, index, col);
        });
        
        // Add data rows
        values.forEach((row, rowIndex) => {
            row.forEach((cell, colIndex) => {
                resultSheet.setCellValue(rowIndex + 1, colIndex, cell);
            });
        });
        
        // Switch to the results sheet
        universheet.setActiveSheet(resultSheet.getSheetId());
        univerAPI.refreshUI();
        
    } catch (error) {
        alert(`SQL Error: ${error.message}`);
    }
}

// Export as CSV
function exportCSV() {
    const activeSheet = universheet.getActiveSheet();
    if (!activeSheet) {
        alert('No sheet to export');
        return;
    }
    
    const data = [];
    const maxRows = 1000; // Limit for performance
    const maxCols = 100; // Limit for performance
    
    // Find data bounds
    let lastRow = 0;
    let lastCol = 0;
    
    for (let r = 0; r < maxRows; r++) {
        let rowHasData = false;
        for (let c = 0; c < maxCols; c++) {
            const value = activeSheet.getCellValue(r, c);
            if (value !== null && value !== undefined && value !== '') {
                rowHasData = true;
                if (c > lastCol) lastCol = c;
            }
        }
        if (rowHasData) lastRow = r;
    }
    
    // Collect data
    for (let r = 0; r <= lastRow; r++) {
        const row = [];
        for (let c = 0; c <= lastCol; c++) {
            const value = activeSheet.getCellValue(r, c) || '';
            row.push(`"${value.toString().replace(/"/g, '""')}"`);
        }
        data.push(row.join(','));
    }
    
    // Create CSV
    const csvContent = data.join('\n');
    
    // Download
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', 'sheet_export.csv');
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Create chart
function createChart() {
    const activeSheet = universheet.getActiveSheet();
    if (!activeSheet) {
        alert('No active sheet to create chart from');
        return;
    }
    
    // For simplicity, we'll just alert about chart creation
    // In a real implementation, you would use Univer's charting API
    alert('Chart creation would be implemented here. In a full implementation, this would open a dialog to select data and chart type.');
}

// Add new sheet
function addSheet() {
    const sheetCount = universheet.getSheets().length;
    const newSheetName = `Sheet${sheetCount + 1}`;
    universheet.addSheet(newSheetName);
    univerAPI.refreshUI();
}

// Initialize with sample data
loadSampleData('sales');
