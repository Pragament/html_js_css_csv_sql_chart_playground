// Initialize SQL.js
let SQL;
let db;

initSqlJs({
    locateFile: file => `https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/${file}`
}).then(function(sql) {
    SQL = sql;
    db = new SQL.Database();
    console.log("SQL.js initialized");
}).catch(err => {
    console.error("Error loading SQL.js:", err);
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

// Current data storage
let currentData = [];
let headers = [];

// Initialize the app
document.addEventListener('DOMContentLoaded', function() {
    // Event listeners
    document.getElementById('import-btn').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
    document.getElementById('run-sql').addEventListener('click', runSQL);
    document.getElementById('export-csv').addEventListener('click', exportCSV);
    document.getElementById('export-excel').addEventListener('click', exportExcel);
    document.getElementById('add-column').addEventListener('click', addColumn);
    document.getElementById('add-row').addEventListener('click', addRow);
    document.getElementById('clear-data').addEventListener('click', clearData);

    // Sample file click handlers
    document.querySelectorAll('#sample-files li').forEach(item => {
        item.addEventListener('click', function() {
            const fileKey = this.getAttribute('data-file');
            loadSampleData(fileKey);
        });
    });

    // Initialize with sample data
    loadSampleData('sales');
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
        const data = e.target.result;
        
        if (file.name.endsWith('.csv')) {
            processCSVData(data);
        } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            processExcelData(data);
        }
    };
    
    if (file.name.endsWith('.csv')) {
        reader.readAsText(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}

// Process CSV data
function processCSVData(csvString) {
    const rows = csvString.split('\n').filter(row => row.trim() !== '');
    const data = rows.map(row => {
        // Simple CSV parsing (for more robust parsing, use a library)
        return row.split(',').map(cell => cell.trim());
    });
    
    if (data.length === 0) return;
    
    headers = data[0];
    currentData = data.slice(1);
    
    renderSpreadsheet();
    updateSQLDatabase();
}

// Process Excel data
function processExcelData(arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    
    if (data.length === 0) return;
    
    headers = data[0];
    currentData = data.slice(1);
    
    renderSpreadsheet();
    updateSQLDatabase();
}

// Render spreadsheet
function renderSpreadsheet() {
    // Clear existing data
    headerRow.innerHTML = '';
    dataBody.innerHTML = '';
    
    // Add headers
    headers.forEach((header, index) => {
        const th = document.createElement('th');
        th.textContent = header;
        th.contentEditable = true;
        th.addEventListener('blur', () => {
            headers[index] = th.textContent;
            updateSQLDatabase();
        });
        headerRow.appendChild(th);
    });
    
    // Add data rows
    currentData.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        row.forEach((cell, colIndex) => {
            const td = document.createElement('td');
            td.textContent = cell;
            td.contentEditable = true;
            td.addEventListener('blur', () => {
                currentData[rowIndex][colIndex] = td.textContent;
                updateSQLDatabase();
            });
            tr.appendChild(td);
        });
        
        // Add empty cells if row is shorter than headers
        while (tr.children.length < headers.length) {
            const td = document.createElement('td');
            td.contentEditable = true;
            td.addEventListener('blur', () => {
                if (!currentData[rowIndex]) currentData[rowIndex] = [];
                currentData[rowIndex][tr.children.length - 1] = td.textContent;
                updateSQLDatabase();
            });
            tr.appendChild(td);
        }
        
        dataBody.appendChild(tr);
    });
}

// Update SQL database
function updateSQLDatabase() {
    if (!SQL || !db) return;
    
    try {
        // Clear existing table
        db.run('DROP TABLE IF EXISTS data;');
        
        // Create table with current headers
        const createTableSQL = `CREATE TABLE data (${headers.map(h => `"${h}" TEXT`).join(', ')});`;
        db.run(createTableSQL);
        
        // Insert data
        const placeholders = headers.map(() => '?').join(', ');
        const stmt = db.prepare(`INSERT INTO data VALUES (${placeholders});`);
        
        currentData.forEach(row => {
            // Ensure row has same number of columns as headers
            const normalizedRow = headers.map((_, i) => row[i] || '');
            stmt.run(normalizedRow);
        });
        
        stmt.free();
    } catch (error) {
        console.error("SQL Error:", error);
    }
}

// Run SQL query
function runSQL() {
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
        
        // Get column names and data
        const columns = result[0].columns;
        const values = result[0].values;
        
        // Update headers and data
        headers = columns;
        currentData = values.map(row => row.map(cell => cell === null ? '' : cell.toString()));
        
        // Re-render spreadsheet
        renderSpreadsheet();
        
    } catch (error) {
        alert(`SQL Error: ${error.message}`);
    }
}

// Export as CSV
function exportCSV() {
    if (headers.length === 0) {
        alert('No data to export');
        return;
    }
    
    // Create CSV content
    let csvContent = headers.join(',') + '\n';
    csvContent += currentData.map(row => 
        row.map(cell => 
            `"${(cell || '').toString().replace(/"/g, '""')}"`
        ).join(',')
    ).join('\n');
    
    // Download
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
    if (headers.length === 0) {
        alert('No data to export');
        return;
    }
    
    // Prepare data
    const exportData = [headers, ...currentData];
    
    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(exportData);
    
    // Create workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    
    // Export
    XLSX.writeFile(wb, 'data_export.xlsx');
}

// Add column
function addColumn() {
    const colName = `Column ${headers.length + 1}`;
    headers.push(colName);
    
    // Add empty values to each row
    currentData.forEach(row => row.push(''));
    
    renderSpreadsheet();
    updateSQLDatabase();
}

// Add row
function addRow() {
    currentData.push(headers.map(() => ''));
    renderSpreadsheet();
    updateSQLDatabase();
}

// Clear data
function clearData() {
    if (confirm('Are you sure you want to clear all data?')) {
        headers = [];
        currentData = [];
        renderSpreadsheet();
        
        if (db) {
            db.run('DROP TABLE IF EXISTS data;');
        }
    }
}
