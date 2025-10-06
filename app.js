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
const loginBtn = document.getElementById('loginBtn');
const logoutBtn = document.getElementById('logoutBtn');
const saveBtn = document.getElementById('saveBtn');
const savedFilesList = document.getElementById('saved-files');
const sheetNameInput = document.getElementById('sheet-name');

// Current data storage
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
    currentData[row] = currentData[row] || [];
    currentData[row][col] = value;
    window.updateSQLDatabase();
};

window.getCellValue = function(cell) {
    if (!cell || !cell.dataset.row || !cell.dataset.column) {
        console.error('Invalid cell for getCellValue:', cell);
        return null;
    }
    const row = parseInt(cell.dataset.row);
    const col = parseInt(cell.dataset.column);
    return currentData[row]?.[col] || cell.textContent || '';
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
    setTimeout(() => {
        notification.style.display = 'none';
    }, 3000);
};

window.generateRandomSheet = function(rows = 4, cols = 5) {
    console.log('Generating sheet:', rows, cols);
    const genHeaders = Array.from({ length: cols }, (_, i) => String.fromCharCode(65 + i));
    const data = Array.from({ length: rows }, () => Array(cols).fill(''));
    headers = genHeaders;
    currentData = data;
    formulas = Array(rows).fill().map(() => Array(cols).fill(undefined));
    dependencies = {};
    window.renderSpreadsheet();
    window.updateSQLDatabase();
    window.updateChartDropdowns();
};

function waitForFirebaseInit(callback) {
    const maxAttempts = 50; // 5 seconds (50 * 100ms)
    let attempts = 0;
    function checkFirebase() {
        if (
            window.firebaseAuth &&
            window.firebaseProvider &&
            window.signInWithPopup &&
            window.signInWithRedirect &&
            window.getRedirectResult &&
            window.signOut &&
            window.doc &&
            window.setDoc &&
            window.getDoc &&
            window.collection &&
            window.query &&
            window.getDocs
        ) {
            console.log('Firebase initialized successfully');
            callback();
        } else {
            attempts++;
            if (attempts >= maxAttempts) {
                console.error('Firebase initialization timed out');
                window.showNotification('Failed to initialize Firebase. Please check your network and refresh.', 'error');
                return;
            }
            console.log(`Waiting for Firebase initialization (attempt ${attempts}/${maxAttempts})`);
            setTimeout(checkFirebase, 100);
        }
    }
    checkFirebase();
}

// Initialize the app
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, setting up event listeners');

    // Verify Firebase script is loaded
    if (!window.firebase) {
        console.error('Firebase script not loaded');
        window.showNotification('Firebase is not available. Please check if firebase.js is loaded.', 'error');
    }

    // Event listeners for other functionalities (unchanged)
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

    // Firebase Authentication Setup
    waitForFirebaseInit(() => {
        console.log('Firebase is ready, setting up auth listeners');

        // Login with popup fallback to redirect
        if (loginBtn) {
            loginBtn.addEventListener('click', async () => {
                if (!window.firebaseAuth || !window.firebaseProvider) {
                    console.error('Firebase auth or provider not initialized');
                    window.showNotification('Authentication service is not available. Please check Firebase setup.', 'error');
                    return;
                }
                try {
                    let result;
                    try {
                        result = await window.signInWithPopup(window.firebaseAuth, window.firebaseProvider);
                        console.log('Popup login successful:', result.user.displayName);
                    } catch (popupErr) {
                        console.warn('Popup login failed:', popupErr.code, popupErr.message);
                        if (popupErr.code === 'auth/popup-blocked' || popupErr.code === 'auth/popup-closed-by-user') {
                            window.showNotification('Popup blocked. Redirecting to login...', 'info');
                            await window.signInWithRedirect(window.firebaseAuth, window.firebaseProvider);
                            return;
                        }
                        throw popupErr;
                    }
                    window.showNotification(`Logged in as ${result.user.displayName || 'User'}`, 'success');
                    updateAuthUI(true);
                    loadSavedList(result.user);
                } catch (err) {
                    console.error('Login error:', err);
                    let errorMessage = 'Login failed. Please try again.';
                    if (err.code === 'auth/network-request-failed') {
                        errorMessage = 'Network error. Please check your internet connection.';
                    } else if (err.code === 'auth/invalid-credential') {
                        errorMessage = 'Invalid credentials. Please check your login details.';
                    } else if (err.code) {
                        errorMessage = `Login failed: ${err.code}`;
                    }
                    window.showNotification(errorMessage, 'error');
                }
            });
        } else {
            console.error('Login button not found');
            window.showNotification('Login button is missing in the UI.', 'error');
        }

        // Handle redirect result and auth state changes
        if (window.firebaseAuth) {
            window.firebaseAuth.onAuthStateChanged(async user => {
                console.log('Auth state changed:', user ? `User logged in: ${user.uid}` : 'No user');
                try {
                    if (user) {
                        const result = await window.getRedirectResult(window.firebaseAuth);
                        if (result) {
                            console.log('Redirect login successful:', result.user.displayName);
                            window.showNotification(`Logged in as ${result.user.displayName || 'User'}`, 'success');
                        }
                        updateAuthUI(true);
                        loadSavedList(user);
                        window.showNotification('Please select a spreadsheet to load.', 'info');
                    } else {
                        updateAuthUI(false);
                        savedFilesList.innerHTML = '';
                        sheetNameInput.value = '';
                        window.showNotification('Logged out', 'info');
                    }
                } catch (err) {
                    console.error('Auth state error:', err);
                    window.showNotification(`Authentication error: ${err.message}`, 'error');
                }
            });
        } else {
            console.error('Firebase auth not initialized');
            window.showNotification('Authentication service is not available. Please check Firebase setup.', 'error');
        }

        // Logout
        if (logoutBtn) {
            logoutBtn.addEventListener('click', async () => {
                if (!window.firebaseAuth) {
                    console.error('Firebase auth not initialized');
                    window.showNotification('Authentication service is not available.', 'error');
                    return;
                }
                try {
                    await window.signOut(window.firebaseAuth);
                    console.log('Logout successful');
                    window.showNotification('Logged out successfully', 'success');
                    updateAuthUI(false);
                    savedFilesList.innerHTML = '';
                    sheetNameInput.value = '';
                } catch (err) {
                    console.error('Logout error:', err);
                    window.showNotification(`Logout failed: ${err.message}`, 'error');
                }
            });
        } else {
            console.error('Logout button not found');
            window.showNotification('Logout button is missing in the UI.', 'error');
        }
    });

    function updateAuthUI(isLoggedIn) {
        loginBtn.style.display = isLoggedIn ? 'none' : 'inline-block';
        logoutBtn.style.display = isLoggedIn ? 'inline-block' : 'none';
        saveBtn.style.display = isLoggedIn ? 'inline-block' : 'none';
        console.log('Auth UI updated:', isLoggedIn);
    }

    // Load list of saved spreadsheets
    async function loadSavedList(user) {
        if (!window.firebaseDb || !window.collection || !window.query || !window.getDocs) {
            console.error('Firestore not initialized');
            window.showNotification('Firestore is not available. Please check Firebase setup.', 'error');
            return;
        }
        try {
            const q = window.query(window.collection(window.firebaseDb, `users/${user.uid}/spreadsheets`));
            const snapshot = await window.getDocs(q);
            savedFilesList.innerHTML = '';
            snapshot.forEach(doc => {
                const li = document.createElement('li');
                li.textContent = doc.data().name || 'Untitled';
                li.dataset.id = doc.id;
                li.addEventListener('click', () => loadSpreadsheet(user, doc.id));
                savedFilesList.appendChild(li);
            });
            console.log('Loaded saved spreadsheets list');
            window.showNotification('Loaded saved spreadsheets list', 'info');
        } catch (err) {
            console.error('Error loading saved list:', err);
            window.showNotification(`Error loading spreadsheets: ${err.message}`, 'error');
        }
    }

    // Load a specific spreadsheet
    async function loadSpreadsheet(user, sheetId) {
        if (!window.firebaseDb || !window.doc || !window.getDoc) {
            console.error('Firestore not initialized');
            window.showNotification('Firestore is not available. Please check Firebase setup.', 'error');
            return;
        }
        try {
            const docRef = window.doc(window.firebaseDb, `users/${user.uid}/spreadsheets/${sheetId}`);
            const docSnap = await window.getDoc(docRef);
            if (docSnap.exists()) {
                const saved = docSnap.data();
                headers = saved.headers || [];
                currentData = saved.data.map((rowObj) => {
                    const row = [];
                    for (let col = 0; col < headers.length; col++) {
                        row.push(rowObj[col] || '');
                    }
                    return row;
                });
                formulas = saved.formulas.map((rowObj) => {
                    const row = [];
                    for (let col = 0; col < headers.length; col++) {
                        row.push(rowObj[col] === null ? undefined : rowObj[col]);
                    }
                    return row;
                });
                dependencies = {};
                for (let r = 0; r < currentData.length; r++) {
                    for (let c = 0; c < headers.length; c++) {
                        if (formulas[r] && formulas[r][c]) {
                            currentData[r][c] = evaluateFormula(formulas[r][c], r, c);
                        }
                    }
                }
                renderSpreadsheet();
                updateSQLDatabase();
                updateChartDropdowns();
                sheetNameInput.value = saved.name || '';
                console.log('Spreadsheet loaded:', saved.name);
                window.showNotification(`Loaded spreadsheet: ${saved.name}`, 'success');
            } else {
                console.error('Spreadsheet not found:', sheetId);
                window.showNotification('Spreadsheet not found', 'error');
            }
        } catch (err) {
            console.error('Error loading spreadsheet:', err);
            window.showNotification(`Error loading spreadsheet: ${err.message}`, 'error');
        }
    }

    // Save with name (new or overwrite)
    if (saveBtn) {
        saveBtn.addEventListener('click', async () => {
            const user = window.firebaseAuth?.currentUser;
            if (!user) {
                console.error('No user logged in for save');
                window.showNotification('Please login first!', 'error');
                return;
            }
            if (!window.firebaseDb || !window.doc || !window.setDoc) {
                console.error('Firestore not initialized');
                window.showNotification('Firestore is not available. Please check Firebase setup.', 'error');
                return;
            }
            const name = sheetNameInput.value.trim();
            if (!name) {
                console.error('No sheet name provided');
                window.showNotification('Enter a sheet name!', 'error');
                return;
            }
            const sheetId = name.replace(/\s/g, '').toLowerCase();
            try {
                if (!formulas || !Array.isArray(formulas)) {
                    formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(undefined));
                }
                const serializedData = currentData.map(row => {
                    const obj = {};
                    row.forEach((val, col) => {
                        obj[col] = val;
                    });
                    return obj;
                });
                const serializedFormulas = formulas.map(row => {
                    const obj = {};
                    if (row) {
                        row.forEach((f, col) => {
                            obj[col] = f || null;
                        });
                    }
                    return obj;
                });
                await window.setDoc(
                    window.doc(window.firebaseDb, 'users', user.uid, 'spreadsheets', sheetId),
                    {
                        name,
                        headers: headers,
                        data: serializedData,
                        formulas: serializedFormulas,
                        savedAt: new Date()
                    }
                );
                console.log(`Saved spreadsheet: ${name}`);
                window.showNotification(`Saved "${name}" to Firestore!`, 'success');
                loadSavedList(user);
            } catch (err) {
                console.error('Save error:', err);
                window.showNotification(`Save failed: ${err.message}`, 'error');
            }
        });
    } else {
        console.error('Save button not found');
        window.showNotification('Save button is missing in the UI.', 'error');
    }
        // Populate test list
    const tests = [
        'Fill Sequence Test',
        'Sum Function Test',
        'Double Click Fill Test'
    ];
    tests.forEach(testName => {
        const li = document.createElement('li');
        li.textContent = testName;
        testList.appendChild(li);
    });
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
        formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(undefined));
        dependencies = {};
        
        renderSpreadsheet();
        updateSQLDatabase();
        updateChartDropdowns();
    } catch (error) {
        console.error('Error processing CSV:', error);
        alert('Error processing CSV file');
    }
}

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
        formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(undefined));
        dependencies = {};
        
        renderSpreadsheet();
        updateSQLDatabase();
        updateChartDropdowns();
    } catch (error) {
        console.error('Error processing Excel:', error);
        alert('Error processing Excel file');
    }
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
        db.run('DROP TABLE IF EXISTS data');
        let columns = headers.map((h, i) => `"${h}" TEXT`).join(', ');
        db.run(`CREATE TABLE data (${columns})`);
        currentData.forEach(row => {
            let values = row.map(val => `'${val || ''}'`).join(', ');
            db.run(`INSERT INTO data VALUES (${values})`);
        });
    } catch (err) {
        console.error('Error updating SQL database:', err);
        window.showNotification('Error updating SQL database', 'error');
    }
}

// Update chart dropdowns
function updateChartDropdowns() {
    xAxisSelect.innerHTML = '<option value="">Select X Axis</option>';
    yAxisSelect.innerHTML = '<option value="">Select Y Axis</option>';
    headers.forEach((header, index) => {
        const optionX = document.createElement('option');
        optionX.value = index;
        optionX.textContent = header;
        xAxisSelect.appendChild(optionX);
        const optionY = document.createElement('option');
        optionY.value = index;
        optionY.textContent = header;
        yAxisSelect.appendChild(optionY);
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
    
    const query = sqlEditor.getValue().trim();
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
        formulas = Array(currentData.length).fill().map(() => Array(headers.length).fill(undefined));
        dependencies = {};
        
        renderSpreadsheet();
        updateChartDropdowns();
        console.log('SQL query executed, data updated');
    } catch (error) {
        console.error('SQL Error:', error);
        alert(`SQL Error: ${error.message}`);
    }
}
function createChart() {
    console.log('Creating chart');
    if (headers.length === 0 || currentData.length === 0) {
        alert('No data available to create chart');
        return;
    }

    let chartType = chartTypeSelect.value;
    let xAxis = xAxisSelect.value;
    let yAxis = yAxisSelect.value;

    let xIndex = parseInt(xAxis);
    let yIndex = parseInt(yAxis);

    if (isNaN(xIndex) || isNaN(yIndex)) {
        alert('Invalid axis selection');
        return;
    }

    // Check if Y-axis is numeric
    let yValues = currentData.map(row => row[yIndex]);
    let yIsNumeric = yValues.every(val => !isNaN(parseFloat(val)) || val === '');

    if (!yIsNumeric) {
        // If not numeric, swap axes and use horizontal bar chart
        [xIndex, yIndex] = [yIndex, xIndex];
        chartType = 'bar'; // Chart.js v3+ uses 'bar' with indexAxis option for horizontal
        var indexAxis = 'y';
    } else {
        var indexAxis = 'x';
    }

    // Now, x-axis is always categories, y-axis is always numeric
    const labels = currentData.map(row => row[xIndex] || '');
    const dataValues = currentData.map(row => {
        const val = row[yIndex];
        return isNaN(parseFloat(val)) ? 0 : parseFloat(val);
    });

    const config = {
        type: chartType,
        data: {
            labels: labels,
            datasets: [{
                label: `${headers[yIndex]} by ${headers[xIndex]}`,
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
                    text: `${headers[yIndex]} by ${headers[xIndex]}`
                }
            },
            scales: {
                x: {
                    display: true,
                    title: {
                        display: true,
                        text: indexAxis === 'y' ? headers[yIndex] : headers[xIndex],
                        font: {
                            size: 14,
                            weight: 'bold'
                        },
                        color: '#333'
                    },
                    ticks: {
                        font: {
                            size: 12
                        }
                    }
                },
                y: {
                    display: true,
                    title: {
                        display: true,
                        text: indexAxis === 'y' ? headers[xIndex] : headers[yIndex],
                        font: {
                            size: 14,
                            weight: 'bold'
                        },
                        color: '#333'
                    },
                    ticks: {
                        font: {
                            size: 12
                        }
                    }
                }
            },
            layout: {
                padding: {
                    top: 30
                }
            }
        }
    };

    const chartId = `chart-${Date.now()}`;
    const chartItem = document.createElement('div');
    chartItem.className = 'chart-item';
    chartItem.innerHTML = `<canvas id="${chartId}"></canvas><div><button class="delete-chart" data-id="${chartId}">Delete</button><button class="download-chart" data-id="${chartId}">Download</button></div>`;
    chartList.appendChild(chartItem);
    
    const chartCtx = document.getElementById(chartId).getContext('2d');
    const chart = new Chart(chartCtx, config);
    
    charts.push({
        id: chartId,
        chart: chart,
        type: chartType,
        xAxis: xIndex,
        yAxis: yIndex
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
function getChartColors(count) { return Array.from({ length: count }, (_, i) => `hsl(${(i * 360 / count)}, 70%, 60%)`); }
function deleteChart(id) {
    const index = charts.findIndex(c => c.id === id);
    if (index > -1) { charts[index].chart.destroy(); charts.splice(index, 1); document.getElementById(id).closest('.chart-item').remove(); }
}
function downloadChart(id) {
    const chart = charts.find(c => c.id === id);
    if (chart) { const link = document.createElement('a'); link.download = 'chart.png'; link.href = chart.chart.canvas.toDataURL('image/png'); link.click(); }
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
    formulas.push(Array(headers.length).fill(undefined));
    
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

window.renderSpreadsheet = renderSpreadsheet;
window.updateSQLDatabase = updateSQLDatabase;
window.updateChartDropdowns = updateChartDropdowns;
window.updateCellValue = updateCellValue;
window.evaluateFormula = evaluateFormula;
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
