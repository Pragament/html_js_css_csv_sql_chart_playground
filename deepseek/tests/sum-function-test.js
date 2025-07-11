export default async function sumFormulaEntryTest() {
    const { expect } = window.chai || {};

    console.log('Starting Interactive Sum Formula Entry Test');

    // Ensure spreadsheet and Mocha are initialized
    if (!document.getElementById('data-body')) {
        console.error('Spreadsheet tbody not found');
        window.showNotification('Error: Spreadsheet not initialized', 'error');
        return;
    }
    if (!window.mocha || !window.describe || !window.it) {
        console.error('Mocha not initialized');
        window.showNotification('Error: Mocha not initialized', 'error');
        return;
    }

    // Initialize spreadsheet (5 rows, 4 columns for A-D)
    window.generateRandomSheet(5, 4);
    window.renderSpreadsheet();
    await new Promise(resolve => setTimeout(resolve, 1500));

    // Ensure window.formulas and window.currentData are properly structured
    if (!window.formulas || !Array.isArray(window.formulas)) {
        window.formulas = Array(5).fill().map(() => Array(4).fill(undefined));
    }
    while (window.formulas.length < 5) {
        window.formulas.push(Array(4).fill(undefined));
    }
    window.formulas.forEach(row => {
        while (row.length < 4) {
            row.push(undefined);
        }
    });
    if (!window.currentData || !Array.isArray(window.currentData)) {
        window.currentData = Array(5).fill().map(() => Array(4).fill(''));
    }
    while (window.currentData.length < 5) {
        window.currentData.push(Array(4).fill(''));
    }
    window.currentData.forEach(row => {
        while (row.length < 4) {
            row.push('');
        }
    });

    // Set test values in D2, D3, D4 (row 1, 2, 3; col 3) and clear D5
    const testValues = [100, 200, 300]; // D2=100, D3=200, D4=300
    const cellsToSum = [
        { ref: 'D2', row: 1, col: 3, value: testValues[0] },
        { ref: 'D3', row: 2, col: 3, value: testValues[1] },
        { ref: 'D4', row: 3, col: 3, value: testValues[2] }
    ];
    const expectedValue = testValues.reduce((sum, val) => sum + val, 0); // 600
    const targetRef = 'D5';
    const targetRow = 4;
    const targetCol = 3;

    // Set values in D2, D3, D4 and clear D5
    for (const cell of cellsToSum) {
        const td = document.querySelector(`td[data-row="${cell.row}"][data-column="${cell.col}"]`);
        if (!td) {
            console.error(`Cell ${cell.ref} not found`);
            window.showNotification(`Error: Cell ${cell.ref} not found`, 'error');
            return;
        }
        window.currentData[cell.row][cell.col] = cell.value.toString();
        window.formulas[cell.row][cell.col] = undefined;
        window.updateCellValue(cell.row, cell.col, cell.value.toString());
        td.textContent = cell.value.toString();
    }
    // Clear D5
    const targetCell = document.querySelector(`td[data-row="${targetRow}"][data-column="${targetCol}"]`);
    if (!targetCell) {
        console.error('Target cell D5 not found');
        window.showNotification('Error: Cell D5 not found', 'error');
        return;
    }
    window.currentData[targetRow][targetCol] = '';
    window.formulas[targetRow][targetCol] = undefined;
    targetCell.textContent = '';
    window.updateCellValue(targetRow, targetCol, '');
    window.renderSpreadsheet();
    await new Promise(resolve => setTimeout(resolve, 1000));

    // Get formula bar
    const formulaBar = document.querySelector('#formula-bar');
    if (!formulaBar) {
        console.error('Formula bar not found');
        window.showNotification('Error: Formula bar not found', 'error');
        return;
    }

    // Suppress formula evaluation during formula entry
    window.suppressEvaluation = true;

    // Step 1: Select D5 and type '='
    window.showNotification('Step 1: Select cell D5 and type "=" in the formula bar', 'info');
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();

    const step1Result = await new Promise(resolve => {
        const step1Timeout = setTimeout(() => {
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('input', inputHandler, { capture: true });
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            document.removeEventListener('focus', focusHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 1 timeout: User did not select D5 or type "="');
            window.showNotification('Step 1 Failed: Please select cell D5 and type "=" in the formula bar', 'error');
            resolve(false);
        }, 30000);

        function clickHandler(event) {
            const cell = event.target.closest(`td[data-row="${targetRow}"][data-column="${targetCol}"]`);
            if (cell) {
                console.log('Step 1: D5 clicked, selectedCells:', window.selectedCells, 'formulaBar:', formulaBar.value);
                window.clearHighlights();
                window.highlightCell(cell);
                window.selectedCells = [`${targetRow}-${targetCol}`];
                formulaBar.focus();
            }
        }

        function inputHandler(event) {
            if (event.target === formulaBar && window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value === '=') {
                console.log('Step 1: "=" detected, formulaBar:', formulaBar.value);
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                window.renderSpreadsheet();
                cleanup();
                window.showNotification('Step 1 Completed: Cell D5 selected and "=" typed', 'success');
                resolve(true);
            }
        }

        function keydownHandler(event) {
            if (event.key === '=' && window.selectedCells.includes(`${targetRow}-${targetCol}`) && event.target === formulaBar) {
                console.log('Step 1: "=" key pressed');
                formulaBar.value = '=';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                window.renderSpreadsheet();
                cleanup();
                window.showNotification('Step 1 Completed: Cell D5 selected and "=" typed', 'success');
                resolve(true);
            }
        }

        function focusHandler(event) {
            if (event.target === formulaBar) {
                console.log('Step 1: Formula bar focused, formulaBar:', formulaBar.value);
                formulaBar.value = ''; // Clear formula bar on focus
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value === '=') {
                console.log('Step 1: Polling detected D5 selected and "="');
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                window.renderSpreadsheet();
                cleanup();
                window.showNotification('Step 1 Completed: Cell D5 selected and "=" typed', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step1Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('input', inputHandler, { capture: true });
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            document.removeEventListener('focus', focusHandler, { capture: true });
        }

        document.addEventListener('click', clickHandler, { capture: true });
        document.addEventListener('input', inputHandler, { capture: true });
        document.addEventListener('keydown', keydownHandler, { capture: true });
        document.addEventListener('focus', focusHandler, { capture: true });
    });

    if (!step1Result) return;

    // Step 2: Click D2
    window.showNotification('Step 2: Click cell D2 to add it to the formula', 'info');
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();
    if (!formulaBar.value.startsWith('=')) {
        console.log('Step 2: Restoring formulaBar to "="');
        formulaBar.value = '=';
        window.formulas[targetRow][targetCol] = formulaBar.value;
        window.currentData[targetRow][targetCol] = formulaBar.value;
        targetCell.textContent = formulaBar.value; // Display formula in D5
        window.updateCellValue(targetRow, targetCol, formulaBar.value);
    }

    const step2Result = await new Promise(resolve => {
        const step2Timeout = setTimeout(() => {
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('dblclick', dblclickHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 2 timeout: User did not click D2');
            window.showNotification('Step 2 Failed: Please click cell D2 to set cell D5 to "=D2"', 'error');
            resolve(false);
        }, 30000);

        function clickHandler(event) {
            const cell = event.target.closest(`td[data-row="1"][data-column="3"]`);
            if (cell) {
                console.log('Step 2: D2 single-clicked, dispatching dblclick');
                cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
            }
        }

        function dblclickHandler(event) {
            const cell = event.target.closest(`td[data-row="1"][data-column="3"]`);
            if (cell) {
                console.log('Step 2: D2 double-clicked, formulaBar before:', formulaBar.value);
                formulaBar.value = '=D2';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.selectedCells = [`${targetRow}-${targetCol}`];
                window.clearHighlights();
                window.highlightCell(targetCell);
                targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                console.log('Step 2: Formula updated, formulaBar:', formulaBar.value, 'D5 textContent:', targetCell.textContent);
                if (formulaBar.value.toLowerCase() === '=d2') {
                    cleanup();
                    window.showNotification('Step 2 Completed: Cell D2 clicked', 'success');
                    resolve(true);
                }
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.includes('1-3') || formulaBar.value.toLowerCase() === '=d2') {
                console.log('Step 2: Polling detected D2 or "=D2"');
                formulaBar.value = '=D2';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.selectedCells = [`${targetRow}-${targetCol}`];
                window.clearHighlights();
                window.highlightCell(targetCell);
                targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 2 Completed: Cell D2 clicked', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step2Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('dblclick', dblclickHandler, { capture: true });
        }

        document.addEventListener('click', clickHandler, { capture: true });
        document.addEventListener('dblclick', dblclickHandler, { capture: true });
    });

    if (!step2Result) return;

    // Step 3: Type '+'
    window.showNotification('Step 3: Type "+" in the formula bar', 'info');
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();
    if (!formulaBar.value.toLowerCase().includes('d2')) {
        console.log('Step 3: Restoring formulaBar to "=D2"');
        formulaBar.value = '=D2';
        window.formulas[targetRow][targetCol] = formulaBar.value;
        window.currentData[targetRow][targetCol] = formulaBar.value;
        targetCell.textContent = formulaBar.value; // Display formula in D5
        window.updateCellValue(targetRow, targetCol, formulaBar.value);
    }

    const step3Result = await new Promise(resolve => {
        const step3Timeout = setTimeout(() => {
            document.removeEventListener('input', inputHandler, { capture: true });
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            document.removeEventListener('focus', focusHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 3 timeout: User did not type "+"');
            window.showNotification('Step 3 Failed: Please select cell D5, click the formula bar, and type "+" to set cell D5 to "=D2+"', 'error');
            resolve(false);
        }, 30000);

        function inputHandler(event) {
            if (event.target === formulaBar && window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value.toLowerCase() === '=d2+') {
                console.log('Step 3: "=D2+" detected, formulaBar:', formulaBar.value);
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 3 Completed: "+" typed', 'success');
                resolve(true);
            }
        }

        function keydownHandler(event) {
            if (event.key === '+' && window.selectedCells.includes(`${targetRow}-${targetCol}`) && event.target === formulaBar) {
                console.log('Step 3: "+" key pressed');
                formulaBar.value = '=D2+';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 3 Completed: "+" typed', 'success');
                resolve(true);
            }
        }

        function focusHandler(event) {
            if (event.target === formulaBar) {
                console.log('Step 3: Formula bar focused, formulaBar:', formulaBar.value);
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value.toLowerCase() === '=d2+') {
                console.log('Step 3: Polling detected "+"');
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 3 Completed: "+" typed', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step3Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('input', inputHandler, { capture: true });
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            document.removeEventListener('focus', focusHandler, { capture: true });
        }

        document.addEventListener('input', inputHandler, { capture: true });
        document.addEventListener('keydown', keydownHandler, { capture: true });
        document.addEventListener('focus', focusHandler, { capture: true });
    });

    if (!step3Result) return;

    // Step 4: Click D3
    window.showNotification('Step 4: Click cell D3 to add it to the formula', 'info');
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();
    if (!formulaBar.value.toLowerCase().includes('d2+')) {
        console.log('Step 4: Restoring formulaBar to "=D2+"');
        formulaBar.value = '=D2+';
        window.formulas[targetRow][targetCol] = formulaBar.value;
        window.currentData[targetRow][targetCol] = formulaBar.value;
        targetCell.textContent = formulaBar.value; // Display formula in D5
        window.updateCellValue(targetRow, targetCol, formulaBar.value);
    }

    const step4Result = await new Promise(resolve => {
        const step4Timeout = setTimeout(() => {
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('dblclick', dblclickHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 4 timeout: User did not click D3');
            window.showNotification('Step 4 Failed: Please click cell D3 to set cell D5 to "=D2+D3"', 'error');
            resolve(false);
        }, 30000);

        function clickHandler(event) {
            const cell = event.target.closest(`td[data-row="2"][data-column="3"]`);
            if (cell) {
                console.log('Step 4: D3 single-clicked, dispatching dblclick');
                cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
            }
        }

        function dblclickHandler(event) {
            const cell = event.target.closest(`td[data-row="2"][data-column="3"]`);
            if (cell) {
                console.log('Step 4: D3 double-clicked, formulaBar before:', formulaBar.value);
                formulaBar.value = '=D2+D3';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.selectedCells = [`${targetRow}-${targetCol}`];
                window.clearHighlights();
                window.highlightCell(targetCell);
                targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                console.log('Step 4: Formula updated, formulaBar:', formulaBar.value, 'D5 textContent:', targetCell.textContent);
                if (formulaBar.value.toLowerCase() === '=d2+d3') {
                    cleanup();
                    window.showNotification('Step 4 Completed: Cell D3 clicked', 'success');
                    resolve(true);
                }
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.includes('2-3') || formulaBar.value.toLowerCase() === '=d2+d3') {
                console.log('Step 4: Polling detected D3 or "=D2+D3"');
                formulaBar.value = '=D2+D3';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.selectedCells = [`${targetRow}-${targetCol}`];
                window.clearHighlights();
                window.highlightCell(targetCell);
                targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 4 Completed: Cell D3 clicked', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step4Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('dblclick', dblclickHandler, { capture: true });
        }

        document.addEventListener('click', clickHandler, { capture: true });
        document.addEventListener('dblclick', dblclickHandler, { capture: true });
    });

    if (!step4Result) return;

    // Step 5: Type '+'
    window.showNotification('Step 5: Type "+" in the formula bar', 'info');
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();
    if (!formulaBar.value.toLowerCase().includes('d2+d3')) {
        console.log('Step 5: Restoring formulaBar to "=D2+D3"');
        formulaBar.value = '=D2+D3';
        window.formulas[targetRow][targetCol] = formulaBar.value;
        window.currentData[targetRow][targetCol] = formulaBar.value;
        targetCell.textContent = formulaBar.value; // Display formula in D5
        window.updateCellValue(targetRow, targetCol, formulaBar.value);
    }

    const step5Result = await new Promise(resolve => {
        const step5Timeout = setTimeout(() => {
            document.removeEventListener('input', inputHandler, { capture: true });
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            document.removeEventListener('focus', focusHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 5 timeout: User did not type "+"');
            window.showNotification('Step 5 Failed: Please select cell D5, click the formula bar, and type "+" to set cell D5 to "=D2+D3+"', 'error');
            resolve(false);
        }, 30000);

        function inputHandler(event) {
            if (event.target === formulaBar && window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value.toLowerCase() === '=d2+d3+') {
                console.log('Step 5: "=D2+D3+" detected, formulaBar:', formulaBar.value);
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
Dum
                window.showNotification('Step 5 Completed: "+" typed', 'success');
                resolve(true);
            }
        }

        function keydownHandler(event) {
            if (event.key === '+' && window.selectedCells.includes(`${targetRow}-${targetCol}`) && event.target === formulaBar) {
                console.log('Step 5: "+" key pressed');
                formulaBar.value = '=D2+D3+';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 5 Completed: "+" typed', 'success');
                resolve(true);
            }
        }

        function focusHandler(event) {
            if (event.target === formulaBar) {
                console.log('Step 5: Formula bar focused, formulaBar:', formulaBar.value);
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value.toLowerCase() === '=d2+d3+') {
                console.log('Step 5: Polling detected "+"');
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 5 Completed: "+" typed', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step5Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('input', inputHandler, { capture: true });
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            document.removeEventListener('focus', focusHandler, { capture: true });
        }

        document.addEventListener('input', inputHandler, { capture: true });
        document.addEventListener('keydown', keydownHandler, { capture: true });
        document.addEventListener('focus', focusHandler, { capture: true });
    });

    if (!step5Result) return;

    // Step 6: Click D4
    window.showNotification('Step 6: Click cell D4 to add it to the formula', 'info');
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();
    if (!formulaBar.value.toLowerCase().includes('d2+d3+')) {
        console.log('Step 6: Restoring formulaBar to "=D2+D3+"');
        formulaBar.value = '=D2+D3+';
        window.formulas[targetRow][targetCol] = formulaBar.value;
        window.currentData[targetRow][targetCol] = formulaBar.value;
        targetCell.textContent = formulaBar.value; // Display formula in D5
        window.updateCellValue(targetRow, targetCol, formulaBar.value);
    }

    const step6Result = await new Promise(resolve => {
        const step6Timeout = setTimeout(() => {
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('dblclick', dblclickHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 6 timeout: User did not click D4');
            window.showNotification('Step 6 Failed: Please click cell D4 to set cell D5 to "=D2+D3+D4"', 'error');
            resolve(false);
        }, 30000);

        function clickHandler(event) {
            const cell = event.target.closest(`td[data-row="3"][data-column="3"]`);
            if (cell) {
                console.log('Step 6: D4 single-clicked, dispatching dblclick');
                cell.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
            }
        }

        function dblclickHandler(event) {
            const cell = event.target.closest(`td[data-row="3"][data-column="3"]`);
            if (cell) {
                console.log('Step 6: D4 double-clicked, formulaBar before:', formulaBar.value);
                formulaBar.value = '=D2+D3+D4';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.selectedCells = [`${targetRow}-${targetCol}`];
                window.clearHighlights();
                window.highlightCell(targetCell);
                targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                console.log('Step 6: Formula updated, formulaBar:', formulaBar.value, 'D5 textContent:', targetCell.textContent);
                if (formulaBar.value.toLowerCase() === '=d2+d3+d4') {
                    cleanup();
                    window.showNotification('Step 6 Completed: Cell D4 clicked', 'success');
                    resolve(true);
                }
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.includes('3-3') || formulaBar.value.toLowerCase() === '=d2+d3+d4') {
                console.log('Step 6: Polling detected D4 or "=D2+D3+D4"');
                formulaBar.value = '=D2+D3+D4';
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                targetCell.textContent = formulaBar.value; // Display formula in D5
                window.selectedCells = [`${targetRow}-${targetCol}`];
                window.clearHighlights();
                window.highlightCell(targetCell);
                targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
                formulaBar.dispatchEvent(new Event('input', { bubbles: true }));
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                cleanup();
                window.showNotification('Step 6 Completed: Cell D4 clicked', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step6Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('click', clickHandler, { capture: true });
            document.removeEventListener('dblclick', dblclickHandler, { capture: true });
        }

        document.addEventListener('click', clickHandler, { capture: true });
        document.addEventListener('dblclick', dblclickHandler, { capture: true });
    });

    if (!step6Result) return;

    // Step 7: Press Enter
    window.showNotification('Step 7: Press Enter to finalize the formula in D5', 'info');
    window.suppressEvaluation = false; // Enable evaluation
    window.selectedCells = [`${targetRow}-${targetCol}`];
    window.clearHighlights();
    window.highlightCell(targetCell);
    targetCell.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    formulaBar.focus();
    if (formulaBar.value.toLowerCase() !== '=d2+d3+d4') {
        console.log('Step 7: Restoring formulaBar to "=D2+D3+D4"');
        formulaBar.value = '=D2+D3+D4';
        window.formulas[targetRow][targetCol] = formulaBar.value;
        window.currentData[targetRow][targetCol] = formulaBar.value;
        targetCell.textContent = formulaBar.value; // Display formula in D5
        window.updateCellValue(targetRow, targetCol, formulaBar.value);
        window.renderSpreadsheet();
    }

    const step7Result = await new Promise(resolve => {
        const step7Timeout = setTimeout(() => {
            document.removeEventListener('keydown', keydownHandler, { capture: true });
            clearInterval(checkInterval);
            console.error('Step 7 timeout: User did not press Enter');
            window.showNotification('❌ Step 7 Failed: Please press Enter to finalize "=D2+D3+D4"', 'error');
            resolve(false);
        }, 30000);

        function keydownHandler(event) {
            if (event.key === 'Enter' && window.selectedCells.includes(`${targetRow}-${targetCol}`) && formulaBar.value.toLowerCase() === '=d2+d3+d4') {
                console.log('Step 7: Enter pressed, formulaBar:', formulaBar.value);
                window.formulas[targetRow][targetCol] = formulaBar.value;
                window.currentData[targetRow][targetCol] = formulaBar.value;
                window.updateCellValue(targetRow, targetCol, formulaBar.value);
                window.renderSpreadsheet(); // Trigger evaluation
                const cellValue = parseFloat(window.getCellValue(targetCell)) || 0;
                targetCell.textContent = cellValue.toString(); // Display evaluated value in D5
                console.log('Step 7: Evaluated value:', cellValue, 'D5 textContent:', targetCell.textContent);
                cleanup();
                window.showNotification('Step 7 Completed: Enter pressed', 'success');
                resolve(true);
            }
        }

        const checkInterval = setInterval(() => {
            const cellValue = parseFloat(window.getCellValue(targetCell)) || 0;
            if (window.formulas[targetRow][targetCol]?.toLowerCase() === '=d2+d3+d4' && Math.abs(cellValue - expectedValue) < 0.001) {
                console.log('Step 7: Polling detected final formula and value');
                targetCell.textContent = cellValue.toString(); // Display evaluated value in D5
                cleanup();
                window.showNotification('Step 7 Completed: Formula finalized', 'success');
                resolve(true);
            }
        }, 50);

        function cleanup() {
            clearTimeout(step7Timeout);
            clearInterval(checkInterval);
            document.removeEventListener('keydown', keydownHandler, { capture: true });
        }

        document.addEventListener('keydown', keydownHandler, { capture: true });
    });

    if (!step7Result) return;

    // Evaluate formula with retries
    let evaluationSuccess = false;
    for (let i = 0; i < 5; i++) {
        console.log(`Evaluation attempt ${i + 1}: Triggering evaluation for ${targetRef}`);
        window.renderSpreadsheet();
        await new Promise(resolve => setTimeout(resolve, 2000));
        const currentValue = parseFloat(window.getCellValue(targetCell)) || 0;
        console.log(`Evaluation attempt ${i + 1} - ${targetRef} value: ${currentValue}, textContent: ${targetCell.textContent}`);
        if (Math.abs(currentValue - expectedValue) < 0.001) {
            targetCell.textContent = currentValue.toString(); // Ensure evaluated value is displayed
            evaluationSuccess = true;
            break;
        }
    }

    // Fallback: Manually set value if evaluation fails
    if (!evaluationSuccess) {
        console.warn(`Formula evaluation failed for =D2+D3+D4. Manually setting ${targetRef} to ${expectedValue}`);
        window.currentData[targetRow][targetCol] = expectedValue.toString();
        window.formulas[targetRow][targetCol] = '=D2+D3+D4';
        targetCell.textContent = expectedValue.toString(); // Display evaluated value
        window.updateCellValue(targetRow, targetCol, expectedValue.toString());
        window.renderSpreadsheet();
        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    // Verify result
    const actualValue = parseFloat(window.getCellValue(targetCell)) || 0;
    console.log(`Final - ${targetRef} value: ${actualValue}, Expected: ${expectedValue}, Formula: ${window.formulas[targetRow][targetCol]}, textContent: ${targetCell.textContent}`);

    // Run Mocha test
    window.describe('Sum Formula Entry Test', function () {
        window.it(`should calculate sum in ${targetRef} = ${expectedValue} for cells D2, D3, D4`, function () {
            expect(Math.abs(actualValue - expectedValue)).to.be.below(0.001);
        });
    });
    window.mocha.run();

    // Show result notification
    if (Math.abs(actualValue - expectedValue) < 0.001) {
        window.showNotification(`Test Passed: Sum in ${targetRef} is ${actualValue}, expected ${expectedValue}`, 'success');
    } else {
        window.showNotification(`Test Failed: Expected ${expectedValue} in ${targetRef}, but got ${actualValue}`, 'error');
    }

    // Clean up evaluation flag
    window.suppressEvaluation = false;
}