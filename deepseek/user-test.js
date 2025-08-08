(function () {
    const { expect } = window.chai || {};

    // --- Test Modules ---
    const testModules = [
        { id: 'fill-sequence-test', name: 'Fill Sequence Test', file: 'tests/fill-sequence-test.js' },
        { id: 'sum-function-test', name: 'Sum Function Test', file: 'tests/sum-function-test.js' },
        { id: 'double-click-fill-test', name: 'Double Click Fill Test', file: 'tests/double-click-fill-test.js' },
    ];

    const loadedTests = [];

    // --- Utility Functions ---
    window.highlightCell = function (cell) {
        if (cell) {
            cell.style.outline = '2px solid orange';
            cell.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
    };

    window.clearHighlights = function () {
        document.querySelectorAll('td, th').forEach(cell => cell.style.outline = '');
    };

    window.getCellByRef = function (ref) {
        const col = ref.charCodeAt(0) - 65; // 'A' => 0
        const row = parseInt(ref.slice(1), 10) - 1; // 1-based to 0-based
        const tbody = document.getElementById('data-body');
        if (!tbody) {
            console.error('tbody not found for ref:', ref);
            return null;
        }
        const tr = tbody.children[row];
        if (!tr) {
            console.error(`Row ${row} not found for ref:`, ref);
            return null;
        }
        const cell = tr.children[col + 1]; // Adjust for row-number cell
        if (!cell) {
            console.error(`Column ${col} not found in row ${row} for ref:`, ref);
            return null;
        }
        return cell;
    };

    window.setCellValue = function (cell, value) {
        if (cell) {
            cell.innerText = value;
            cell.dispatchEvent(new Event('input', { bubbles: true }));
        }
    };

    window.getCellValue = function (cell) {
        if (!cell) return 0;
        return Number(cell.innerText.trim()) || 0; // Fallback to 0 if NaN
    };

    // --- Initialize Mocha ---
    function initMocha() {
        if (!window.mocha || !window.chai) {
            console.error('Mocha or Chai not loaded.');
            window.showNotification && window.showNotification('❌ Error: Mocha or Chai not loaded.', 'error');
            return false;
        }
        window.mocha.setup({
            ui: 'bdd',
            reporter: 'html'
        });
        const mochaDiv = document.getElementById('mocha') || document.createElement('div');
        mochaDiv.id = 'mocha';
        document.body.appendChild(mochaDiv);
        return true;
    }

    // --- Test List UI ---
    async function renderTestList() {
        const testList = document.getElementById('test-list');
        if (!testList) {
            console.error('Test list element not found.');
            return;
        }

        // Wait for spreadsheet initialization
        await new Promise(resolve => {
            const checkSpreadsheet = () => {
                if (document.getElementById('data-body') && window.renderSpreadsheet) {
                    resolve();
                } else {
                    setTimeout(checkSpreadsheet, 100);
                }
            };
            checkSpreadsheet();
        });

        // Initialize Mocha
        if (!initMocha()) {
            console.error('Failed to initialize Mocha.');
            return;
        }

        testList.innerHTML = '';
        testModules.forEach((test, idx) => {
            const li = document.createElement('li');
            li.textContent = test.name;
            li.style.cursor = 'pointer';
            li.onclick = async () => {
                // Clear previous Mocha results
                const mochaDiv = document.getElementById('mocha');
                if (mochaDiv) mochaDiv.innerHTML = '';

                // Reset spreadsheet state - but don't clear selected cells for sum-function-test
                if (test.id !== 'sum-function-test') {
                    window.selectedCells = [];
                }
                window.clearHighlights();
                if (window.renderSpreadsheet) {
                    window.renderSpreadsheet();
                }

                // Reset flags for fill-sequence-test
                if (test.id === 'fill-sequence-test') {
                    window.fillSequenceTestCompleted = false;
                    window.fillSequenceTestRunning = false;
                    const mod = await import(`./${test.file}?cacheBust=${Date.now()}`);
                    await mod.default();
                } else {
                    if (!loadedTests[idx]) {
                        const mod = await import(`./${test.file}`);
                        loadedTests[idx] = mod.default;
                    }
                    await loadedTests[idx]();
                }
            };
            testList.appendChild(li);
        });
    }

    // --- On DOM Ready, render test list ---
    window.addEventListener('DOMContentLoaded', () => {
        console.log('DOM loaded, initializing test list...');
        renderTestList();
    });
})();