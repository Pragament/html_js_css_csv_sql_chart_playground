export default async function fillSequenceTest() {
    const { expect } = window.chai || {};

    if (window.fillSequenceTestRunning) {
        console.log('Fill Sequence Test already running, skipping duplicate load.');
        return;
    }

    if (window.fillSequenceTestCompleted) {
        console.log('Fill Sequence Test already completed, skipping re-run.');
        return;
    }

    window.fillSequenceTestRunning = true;

    // Ensure spreadsheet DOM exists
    if (!document.getElementById('spreadsheet')) {
        console.error('Spreadsheet container not found');
        window.showNotification('Setup Failed: Spreadsheet container not found.', 'error');
        window.fillSequenceTestRunning = false;
        return;
    }

    console.log('Starting Fill Sequence Test setup');
    window.generateRandomSheet(4, 5);
    window.renderSpreadsheet();
    await new Promise(res => setTimeout(res, 100)); // Wait for render

    // Step 1: Select any cell and input a value (numeric or text-number)
    window.clearHighlights();
    window.showNotification("Step 1: Click any cell and enter a value.", 'info');

    const step1Result = await new Promise(resolve => {
        let selectedCell = null;
        let initialValue = null;

        const testValues = ['150', 'A1', 'HELLO1'];
        const suggestedValue = testValues[Math.floor(Math.random() * testValues.length)];
        console.log('Suggested test value:', suggestedValue);

        function clickHandler(event) {
            const cell = event.target.closest('td[data-row][data-column]');
            if (cell && !cell.classList.contains('row-number')) {
                selectedCell = cell;
                window.highlightCell(cell);
                console.log('Step 1: Cell clicked', cell.dataset.row, cell.dataset.column);
            }
        }

        function blurHandler(event) {
            const cell = event.target.closest('td[data-row][data-column]');
            if (cell && cell === selectedCell) {
                const value = String(window.getCellValue(cell) || '');
                console.log('Step 1: Retrieved value:', value, 'type:', typeof value);
                if (value !== '') {
                    initialValue = value;
                    const fillHandle = cell.querySelector('.fill-handle');
                    if (!fillHandle) {
                        console.error('Fill handle not found on selected cell');
                        window.showNotification('Step 1 Failed: Fill handle not found.', 'error');
                        resolve(null);
                        return;
                    }
                    clearTimeout(step1Timeout);
                    document.removeEventListener('click', clickHandler);
                    document.removeEventListener('blur', blurHandler);
                    clearInterval(checkInterval);
                    window.showNotification(`Step 1 Completed: Cell ${toCellRef(cell)} set to "${value}"`, 'success');
                    setTimeout(() => resolve({ cell, value }), 1000);
                }
            }
        }

        const checkInterval = setInterval(() => {
            if (window.selectedCells.length === 1) {
                const cellKey = window.selectedCells[0];
                const [row, col] = cellKey.split('-').map(Number);
                const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
                const value = String(window.getCellValue(cell) || '');
                console.log('Step 1: Check interval value:', value, 'type:', typeof value);
                if (value !== '') {
                    initialValue = value;
                    selectedCell = cell;
                    const fillHandle = cell.querySelector('.fill-handle');
                    if (!fillHandle) {
                        console.error('Fill handle not found on selected cell');
                        window.showNotification('Step 1 Failed: Fill handle not found.', 'error');
                        resolve(null);
                        return;
                    }
                    clearTimeout(step1Timeout);
                    document.removeEventListener('click', clickHandler);
                    document.removeEventListener('blur', blurHandler);
                    clearInterval(checkInterval);
                    window.showNotification(`Step 1 Completed: Cell ${toCellRef(cell)} set to "${value}"`, 'success');
                    setTimeout(() => resolve({ cell, value }), 1000);
                }
            }
        }, 100);

        const step1Timeout = setTimeout(() => {
            clearInterval(checkInterval);
            document.removeEventListener('click', clickHandler);
            document.removeEventListener('blur', blurHandler);
            window.showNotification('Step 1 Failed: Please select a cell and enter a value.', 'error');
            resolve(null);
        }, 20000);

        document.addEventListener('click', clickHandler);
        document.addEventListener('blur', blurHandler);
    });

    if (!step1Result) {
        window.fillSequenceTestRunning = false;
        return;
    }

    const { cell: startCell, value: initialValue } = step1Result;
    const startRow = parseInt(startCell.dataset.row);
    const startCol = parseInt(startCell.dataset.column);
    console.log('Step 1: Initial value confirmed:', initialValue, 'at cell:', toCellRef(startCell));

    await new Promise(res => setTimeout(res, 1000)); // Wait before next step

    // Step 2: Hover over the fill handle
    window.clearHighlights();
    window.highlightCell(startCell);
    window.showNotification(`Step 2: Hover over the fill handle (small square) in the bottom-right corner of ${toCellRef(startCell)}.`, 'info');

    let fillHandle = startCell.querySelector('.fill-handle');
    if (!fillHandle) {
        console.error('Fill handle not found for Step 2');
        window.showNotification('Step 2 Failed: Fill handle not found.', 'error');
        window.fillSequenceTestRunning = false;
        return;
    }

    const step2Result = await new Promise(resolve => {
        function mouseOverHandler(event) {
            const target = event.target.closest('.fill-handle');
            if (target && target === fillHandle) {
                console.log('Step 2: Fill handle hovered');
                clearTimeout(step2Timeout);
                document.removeEventListener('mouseover', mouseOverHandler, { capture: true });
                window.showNotification(`Step 2 Completed: Fill handle hovered.`, 'success');
                setTimeout(() => resolve(true), 1000);
            }
        }

        const step2Timeout = setTimeout(() => {
            document.removeEventListener('mouseover', mouseOverHandler, { capture: true });
            window.showNotification('Step 2 Failed: Please hover over the fill handle.', 'error');
            resolve(false);
        }, 20000);

        document.addEventListener('mouseover', mouseOverHandler, { capture: true });
    });

    if (!step2Result) {
        window.fillSequenceTestRunning = false;
        return;
    }

    await new Promise(res => setTimeout(res, 1000)); // Wait before next step

    // Step 3: Click the fill handle
    window.showNotification(`Step 3: Click the fill handle in ${toCellRef(startCell)}.`, 'info');

    fillHandle = startCell.querySelector('.fill-handle');
    if (!fillHandle) {
        console.error('Fill handle not found for Step 3');
        window.showNotification('Step 3 Failed: Fill handle not found.', 'error');
        window.fillSequenceTestRunning = false;
        return;
    }

    const step3Result = await new Promise(resolve => {
        function clickHandler(event) {
            const target = event.target.closest('.fill-handle') || event.target.closest(`td[data-row="${startRow}"][data-column="${startCol}"]`);
            if (target && (target === fillHandle || target === startCell)) {
                console.log('Step 3: Fill handle or cell clicked', target === fillHandle ? 'fill handle' : 'cell');
                clearTimeout(step3Timeout);
                clearInterval(checkInterval);
                document.removeEventListener('click', clickHandler, { capture: true });
                window.showNotification(`Step 3 Completed: Fill handle clicked.`, 'success');
                setTimeout(() => resolve(true), 1000);
            } else {
                console.log('Step 3: Click ignored, target not fill handle or start cell', event.target);
            }
        }

        const checkInterval = setInterval(() => {
            fillHandle = startCell.querySelector('.fill-handle');
            if (!fillHandle) {
                console.warn('Step 3: Fill handle missing during polling');
            }
        }, 100);

        const step3Timeout = setTimeout(() => {
            clearInterval(checkInterval);
            document.removeEventListener('click', clickHandler, { capture: true });
            window.showNotification('Step 3 Failed: Please click the fill handle.', 'error');
            resolve(false);
        }, 20000);

        document.addEventListener('click', clickHandler, { capture: true });
    });

    if (!step3Result) {
        window.fillSequenceTestRunning = false;
        return;
    }

    await new Promise(res => setTimeout(res, 1000)); // Wait before next step

    // Step 4: Hold the mouse button on the fill handle
    window.showNotification(`Step 4: Hold down the left mouse button on the fill handle in ${toCellRef(startCell)} for at least 0.5 seconds.`, 'info');

    fillHandle = startCell.querySelector('.fill-handle');
    if (!fillHandle) {
        console.error('Fill handle not found for Step 4');
        window.showNotification('Step 4 Failed: Fill handle not found.', 'error');
        window.fillSequenceTestRunning = false;
        return;
    }

    const step4Result = await new Promise(resolve => {
        let isMouseDown = false;
        let mouseDownTime = null;

        function mouseDownHandler(event) {
            const target = event.target.closest('.fill-handle') || event.target.closest(`td[data-row="${startRow}"][data-column="${startCol}"]`);
            if (target && (target === fillHandle || target === startCell)) {
                isMouseDown = true;
                mouseDownTime = Date.now();
                console.log('Step 4: Mouse down detected at', mouseDownTime, 'on', target === fillHandle ? 'fill handle' : 'cell');
            } else {
                console.log('Step 4: Mousedown ignored, target:', event.target);
            }
        }

        function mouseUpHandler(event) {
            if (isMouseDown) {
                const holdDuration = Date.now() - mouseDownTime;
                console.log('Step 4: Mouse up detected, hold duration:', holdDuration, 'ms');
                isMouseDown = false;
                mouseDownTime = null;
                // Do not fail immediately; allow hold to be detected in interval
            }
        }

        const checkInterval = setInterval(() => {
            if (isMouseDown && mouseDownTime && (Date.now() - mouseDownTime >= 300)) {
                console.log('Step 4: Hold detected for at least 300ms');
                clearTimeout(step4Timeout);
                clearInterval(checkInterval);
                document.removeEventListener('mousedown', mouseDownHandler, { capture: true });
                document.removeEventListener('mouseup', mouseUpHandler, { capture: true });
                window.showNotification(`Step 4 Completed: Mouse button held for ${Date.now() - mouseDownTime}ms.`, 'success');
                setTimeout(() => resolve(true), 1000);
            }
        }, 50);

        const step4Timeout = setTimeout(() => {
            clearInterval(checkInterval);
            document.removeEventListener('mousedown', mouseDownHandler, { capture: true });
            document.removeEventListener('mouseup', mouseUpHandler, { capture: true });
            console.log('Step 4: Timeout, isMouseDown:', isMouseDown, 'mouseDownTime:', mouseDownTime);
            window.showNotification('Step 4 Failed: Please hold the mouse button on the fill handle for at least 0.3 seconds.', 'error');
            resolve(false);
        }, 30000);

        document.addEventListener('mousedown', mouseDownHandler, { capture: true });
        document.addEventListener('mouseup', mouseUpHandler, { capture: true });
    });

    if (!step4Result) {
        window.fillSequenceTestRunning = false;
        return;
    }

    await new Promise(res => setTimeout(res, 1000)); // Wait before next step

    // Step 5
    window.clearHighlights();
    window.fillSequenceTestSelecting = true; // Flag to prevent app.js from clearing selections
    window.selectedCells = []; // Initialize empty to allow user selection
    const secondCellRef = { dataset: { row: parseInt(startCell.dataset.row) + 1, column: startCell.dataset.column } };
    window.showNotification(`Step 5: Select two cells in the same column (e.g., ${toCellRef(startCell)} and ${toCellRef(secondCellRef)}), then drag the fill handle downward from the second cell to your desired range.`, 'info');

    const step5Result = await new Promise(resolve => {
        let capturedFillRange = null;
        const handlers = [];

        // Wait for fill handle
        async function waitForFillHandle(cell, timeout = 5000) {
            const start = Date.now();
            while (Date.now() - start < timeout) {
                const fillHandle = cell.querySelector('.fill-handle');
                if (fillHandle) return fillHandle;
                await new Promise(res => setTimeout(res, 100));
            }
            return null;
        }

        // Check for filled cells in the DOM
        function checkFilledCells(secondRow, secondCol) {
            const filledCells = [];
            for (let row = secondRow + 1; row < window.currentData.length; row++) {
                const cell = document.querySelector(`td[data-row="${row}"][data-column="${secondCol}"]`);
                if (cell) {
                    const value = String(window.getCellValue(cell) || '').trim();
                    if (value !== '') {
                        filledCells.push([row, secondCol]);
                        console.log(`Step 5: Found filled cell ${toCellRef({ dataset: { row, column: secondCol } })} with value ${value}`);
                    }
                }
            }
            return filledCells.length > 0 ? filledCells : null;
        }

        // Monitor user selection of two cells
        function checkUserSelection() {
            const selectedCells = window.selectedCells || [];
            console.log('Step 5: Checking user selection, selectedCells:', selectedCells);
            // Fallback: Check DOM for selected cells
            const domSelected = document.querySelectorAll('td[data-row][data-column]:not(.row-number).selected-cell');
            const domKeys = Array.from(domSelected).map(cell => `${cell.dataset.row}-${cell.dataset.column}`);
            console.log('Step 5: DOM selected cells:', domKeys);
            const combinedSelected = [...new Set([...selectedCells, ...domKeys])];
            console.log('Step 5: Combined selected cells:', combinedSelected);

            if (combinedSelected.length >= 2) {
                const selected = combinedSelected.slice(-2).map(key => {
                    const [row, col] = key.split('-').map(Number);
                    return { row, col };
                });
                // Verify cells are in the same column
                if (selected.length === 2 && selected[0].col === selected[1].col && selected[0].col === startCol) {
                    // Sort by row to ensure firstCell is the topmost
                    const [firstCell, secondCell] = selected.sort((a, b) => a.row - b.row);
                    console.log(`Step 5: Valid selection detected: ${toCellRef({ dataset: { row: firstCell.row, column: firstCell.col } })} and ${toCellRef({ dataset: { row: secondCell.row, column: secondCell.col } })}`);
                    window.fillSequenceTestSelecting = false;
                    // Update window.selectedCells to ensure consistency
                    window.selectedCells = [`${firstCell.row}-${firstCell.col}`, `${secondCell.row}-${secondCell.col}`];
                    window.showNotification(`Step 5: Two cells selected (${toCellRef({ dataset: { row: firstCell.row, column: firstCell.col } })} and ${toCellRef({ dataset: { row: secondCell.row, column: secondCell.col } })}), now drag the fill handle downward from the second cell to your desired range.`, 'info');
                    startDragListeners(firstCell.row, firstCell.col, secondCell.row, secondCell.col);
                    return true;
                } else {
                    console.log('Step 5: Invalid selection, cells not in same column or not in startCol:', selected);
                }
            } else {
                console.log('Step 5: Not enough cells selected, combinedSelected length:', combinedSelected.length);
            }
            return false;
        }

        // Click handler for user selection
        function clickHandler(event) {
            const cell = event.target.closest('td[data-row][data-column]:not(.row-number)');
            if (cell) {
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.column);
                const cellKey = `${row}-${col}`;
                console.log(`Step 5: Cell clicked: ${toCellRef(cell)}, current selectedCells:`, window.selectedCells);
                // Manually update selectedCells if not updated by app.js
                if (!window.selectedCells.includes(cellKey)) {
                    window.selectedCells.push(cellKey);
                    cell.classList.add('selected-cell');
                }
                if (checkUserSelection()) {
                    document.removeEventListener('click', clickHandler, { capture: true });
                    handlers.splice(handlers.findIndex(h => h.event === 'click'), 1);
                }
            } else {
                console.log('Step 5: Click ignored, not a valid cell:', event.target);
            }
        }

        // Start drag listeners
        async function startDragListeners(startRow, startCol, secondRow, secondCol) {
            const cell = document.querySelector(`td[data-row="${secondRow}"][data-column="${secondCol}"]`);
            if (!cell) {
                console.error('Step 5: Second cell not found');
                window.showNotification('Step 5 Failed: Second cell not found.', 'error');
                resolve({ success: false, capturedFillRange: null });
                return;
            }
            const fillHandle = await waitForFillHandle(cell);
            if (!fillHandle) {
                console.error('Step 5: Fill handle not found on second cell');
                window.showNotification('Step 5 Failed: Fill handle not found.', 'error');
                resolve({ success: false, capturedFillRange: null });
                return;
            }

            let dragStarted = false;

            function mouseDownHandler(event) {
                const target = event.target.closest('.fill-handle');
                if (target && cell.contains(target)) {
                    dragStarted = true;
                    window.showNotification('Step 5: Dragging started, release to complete.', 'info');
                    console.log('Step 5: Mousedown on fill handle, drag started');
                }
            }

            function mouseMoveHandler(event) {
                if (dragStarted) {
                    console.log('Step 5: Mousemove detected during drag');
                }
            }

            function mouseUpHandler(event) {
                if (dragStarted) {
                    console.log('Step 5: Mouseup detected, fillRange:', window.fillRange, 'selectedCells:', window.selectedCells);
                    capturedFillRange = window.fillRange && window.fillRange.length > 0 ? [...window.fillRange] : checkFilledCells(secondRow, secondCol);
                    if (capturedFillRange && capturedFillRange.length > 0) {
                        console.log('Step 5: Captured fill range:', capturedFillRange, 'source:', window.fillRange && window.fillRange.length > 0 ? 'window.fillRange' : 'DOM check');
                        cleanup();
                        window.showNotification(`Step 5: Dragged successfully.`, 'success');
                        setTimeout(() => resolve({ success: true, capturedFillRange }), 8000);
                    }
                    dragStarted = false;
                }
            }

            function doubleClickHandler(event) {
                const target = event.target.closest('.fill-handle');
                if (target && cell.contains(target) && dragStarted) {
                    console.log('Step 5: Double-click detected, fillRange:', window.fillRange, 'selectedCells:', window.selectedCells);
                    capturedFillRange = window.fillRange && window.fillRange.length > 0 ? [...window.fillRange] : checkFilledCells(secondRow, secondCol);
                    if (capturedFillRange && capturedFillRange.length > 0) {
                        console.log('Step 5: Captured fill range after double-click:', capturedFillRange, 'source:', window.fillRange && window.fillRange.length > 0 ? 'window.fillRange' : 'DOM check');
                        cleanup();
                        window.showNotification(`Step 5: Dragged successfully.`, 'success');
                        setTimeout(() => resolve({ success: true, capturedFillRange }), 8000);
                    }
                    dragStarted = false;
                }
            }

            function cleanup() {
                clearTimeout(step5Timeout);
                clearInterval(checkInterval);
                clearInterval(fillHandleCheck);
                cleanupListeners(handlers);
                if (selectionObserver) selectionObserver.disconnect();
                document.removeEventListener('mousedown', mouseDownHandler, { capture: true });
                document.removeEventListener('mousemove', mouseMoveHandler, { capture: true });
                document.removeEventListener('mouseup', mouseUpHandler, { capture: true });
                document.removeEventListener('dblclick', doubleClickHandler, { capture: true });
                window.fillSequenceTestSelecting = false;
            }

            const checkInterval = setInterval(() => {
                if (dragStarted) {
                    console.log('Step 5: Drag check interval, fillRange:', window.fillRange, 'selectedCells:', window.selectedCells);
                    capturedFillRange = window.fillRange && window.fillRange.length > 0 ? [...window.fillRange] : checkFilledCells(secondRow, secondCol);
                    if (capturedFillRange && capturedFillRange.length > 0) {
                        console.log('Step 5: Captured fill range in interval:', capturedFillRange, 'source:', window.fillRange && window.fillRange.length > 0 ? 'window.fillRange' : 'DOM check');
                        cleanup();
                        window.showNotification(`Step 5: Dragged successfully.`, 'success');
                        setTimeout(() => resolve({ success: true, capturedFillRange }), 8000);
                    }
                }
            }, 50);

            const step5Timeout = setTimeout(() => {
                console.error('Step 5 timeout, fillRange:', window.fillRange, 'selectedCells:', window.selectedCells, 'dragStarted:', dragStarted);
                capturedFillRange = checkFilledCells(secondRow, secondCol);
                if (capturedFillRange && capturedFillRange.length > 0) {
                    console.log('Step 5: Timeout fallback - Captured filled cells from DOM:', capturedFillRange);
                    cleanup();
                    window.showNotification(`Step 5: Dragged successfully.`, 'success');
                    setTimeout(() => resolve({ success: true, capturedFillRange }), 8000);
                } else {
                    cleanup();
                    window.showNotification('Step 5 Failed: Please select two cells in the same column (e.g., C1 and C2) and drag the fill handle downward. Timeout after 30 seconds.', 'error');
                    resolve({ success: false, capturedFillRange: null });
                }
            }, 30000);

            const fillHandleCheck = setInterval(() => {
                const currentFillHandle = cell.querySelector('.fill-handle');
                if (!currentFillHandle) {
                    console.warn('Step 5: Fill handle missing, re-attempting to find it');
                    waitForFillHandle(cell, 2000).then(newFillHandle => {
                        if (newFillHandle) {
                            console.log('Step 5: Re-attached listeners to new fill handle');
                        }
                    });
                }
            }, 500);

            handlers.push(
                { event: 'mouseup', handler: mouseUpHandler, options: { capture: true } },
                { event: 'dblclick', handler: doubleClickHandler, element: document, options: { capture: true } },
                { event: 'interval', id: checkInterval },
                { event: 'timeout', id: step5Timeout },
                { event: 'interval', id: fillHandleCheck }
            );
            document.addEventListener('mousedown', mouseDownHandler, { capture: true });
            document.addEventListener('mousemove', mouseMoveHandler, { capture: true });
            document.addEventListener('mouseup', mouseUpHandler, { capture: true });
            document.addEventListener('dblclick', doubleClickHandler, { capture: true });
        }

        // Start listening for user clicks
        document.addEventListener('click', clickHandler, { capture: true });
        handlers.push({ event: 'click', handler: clickHandler, options: { capture: true } });

        // Periodically check selection
        const selectionCheckInterval = setInterval(() => {
            if (checkUserSelection()) {
                clearInterval(selectionCheckInterval);
                handlers.splice(handlers.findIndex(h => h.event === 'interval' && h.id === selectionCheckInterval), 1);
            }
        }, 50);

        // MutationObserver for selection changes
        let selectionObserver = null;
        if (typeof MutationObserver !== 'undefined') {
            selectionObserver = new MutationObserver(() => {
                console.log('Step 5: DOM mutation detected, re-checking selection');
                if (checkUserSelection()) {
                    selectionObserver.disconnect();
                    handlers.splice(handlers.findIndex(h => h.event === 'observer'), 1);
                }
            });
            selectionObserver.observe(document.body, {
                childList: true,
                subtree: true,
                attributes: true,
                attributeFilter: ['class'],
            });
            handlers.push({ event: 'observer', handler: selectionObserver });
        }

        handlers.push({ event: 'interval', id: selectionCheckInterval });
    });

    if (!step5Result.success) {
        console.error('Step 5 failed, capturedFillRange:', step5Result.capturedFillRange);
        window.fillSequenceTestRunning = false;
        window.fillRange = [];
        window.selectedCells = [];
        return;
    }

    await new Promise(res => setTimeout(res, 3000)).catch(err => {
        console.error('Evaluation wait failed:', err);
    });

    // Evaluate result
    console.log('Starting evaluation, capturedFillRange:', step5Result.capturedFillRange, 'selectedCells:', window.selectedCells);
    window.clearHighlights();

    const notification = document.getElementById('notification');
    if (notification) notification.style.display = 'none'; // Clear previous notifications

    const capturedFillRange = step5Result.capturedFillRange;
    const fillRange = capturedFillRange ? capturedFillRange.filter(([r, c]) => c === startCol && r > startRow) : [];
    console.log('Filtered fillRange:', fillRange, 'startRow:', startRow, 'startCol:', startCol);
    fillRange.sort((a, b) => a[0] - b[0]);

    const filledCellsCount = capturedFillRange ? capturedFillRange.length : 0;

    if (fillRange.length === 0) {
        console.warn('No cells filled below start cell');
        window.showNotification('Step 5 Failed: No cells filled below the start cell.', 'error');
        window.fillSequenceTestRunning = false;
        window.fillRange = [];
        window.selectedCells = [];
        return;
    }

    // Get second cell's value
    let secondCellKey = window.selectedCells.find(key => key !== `${startRow}-${startCol}`);
    if (!secondCellKey) {
        console.warn('Only one cell selected — defaulting secondValue to initialValue');
        secondCellKey = `${startRow + 1}-${startCol}`; // fake second cell just below
    }

    const [secondRow, secondCol] = secondCellKey.split('-').map(Number);

    const secondCell = document.querySelector(`td[data-row="${secondRow}"][data-column="${secondCol}"]`);
    const secondValue = secondCell ? String(window.getCellValue(secondCell) || '') : initialValue;

    const values = [];
    let allCorrect = true;
    let completedCells = 0; // New counter

    fillRange.forEach(([row, col], index) => {
        const cell = document.querySelector(`td[data-row="${row}"][data-column="${col}"]`);
        let actualValue = null;
        let retries = 0;

        const retryInterval = setInterval(() => {
            actualValue = cell ? String(window.getCellValue(cell) || '') : null;
            if (actualValue !== '' || retries >= 10) {
                clearInterval(retryInterval);
                let expectedValue;
                const num1 = parseFloat(initialValue);
                const num2 = parseFloat(secondValue);
                if (!isNaN(num1) && !isNaN(num2)) {
                    const step = num2 - num1;
                    expectedValue = (num1 + step * (row - startRow)).toString();
                } else {
                    expectedValue = secondValue;
                }

                values.push({ cell: toCellRef(cell), actual: actualValue, expected: expectedValue });
                console.log(`Evaluating ${toCellRef(cell)}: actual=${actualValue}, expected=${expectedValue}`);
                if (actualValue !== expectedValue) {
                    allCorrect = false;
                }

                completedCells++;
                if (completedCells === fillRange.length) {
                    finalizeEvaluation();
                }
            }
            retries++;
            console.log(`Step 5: Retrying actualValue for ${toCellRef(cell)}, attempt:`, retries);
        }, 100);
    });

    function finalizeEvaluation() {
        const startCellValue = String(window.getCellValue(startCell) || '');
        const secondCellValue = String(window.getCellValue(secondCell) || '');

        values.unshift(
            { cell: toCellRef(startCell), actual: startCellValue, expected: initialValue },
            { cell: toCellRef(secondCell), actual: secondCellValue, expected: secondValue }
        );

        console.log(`Evaluating ${toCellRef(startCell)}: actual=${startCellValue}, expected=${initialValue}`);
        console.log(`Evaluating ${toCellRef(secondCell)}: actual=${secondCellValue}, expected=${secondValue}`);
        if (startCellValue !== initialValue || secondCellValue !== secondValue) {
            allCorrect = false;
        }

        if (window.describe && window.it && window.mocha) {
            describe('Fill Sequence Test', function () {
                it(`should extend sequence from ${toCellRef(startCell)}=${initialValue} and ${toCellRef(secondCell)}=${secondValue}`, function () {
                    expect(allCorrect).to.be.true;
                });
            });
            mocha.run();
        }

        const sequencePreview = fillRange.length > 0 ? ` to ${values.slice(2).map(v => v.expected).join(', ')}` : '';
        if (allCorrect) {
            window.showNotification(`Test Passed: Sequence Generated.`, 'success');
            window.fillSequenceTestCompleted = true;
        } else {
            const errorDetails = values.map(v => `${v.cell}: expected ${v.expected}, got ${v.actual}`).join('\n');
            window.showNotification(`Test Failed: Sequence incorrect, ${filledCellsCount} cells filled.\n${errorDetails}`, 'error');
        }

        window.fillSequenceTestRunning = false;
        window.fillRange = [];
        window.selectedCells = [];
    }

}

// Helper function to clean up event listeners
function cleanupListeners(handlers) {
    handlers.forEach(({ event, handler, element, options, id }) => {
        if (event === 'interval' || event === 'timeout') {
            if (event === 'interval') clearInterval(id);
            if (event === 'timeout') clearTimeout(id);
        } else {
            (element || document).removeEventListener(event, handler, options);
        }
    });
    console.log('Cleaned up event listeners');
}

// Helper functions
function toCellRef(cell) {
    const row = parseInt(cell.dataset.row) + 1;
    let col = parseInt(cell.dataset.column) + 1;
    let colStr = '';
    while (col > 0) {
        colStr = String.fromCharCode(65 + ((col - 1) % 26)) + colStr;
        col = Math.floor((col - 1) / 26);
    }
    return `${colStr}${row}`;
}

function detectPattern(value) {
    console.log('Detecting pattern for value:', value, 'type:', typeof value);
    if (value === null || value === undefined) {
        console.error('Invalid value in detectPattern:', value);
        return { type: 'text', base: '' };
    }
    value = String(value);

    const textNumberMatch = value.match(/^(.+?)(\d+)$/);
    if (textNumberMatch && isNaN(parseFloat(textNumberMatch[1]))) {
        console.log('Text-number pattern detected:', { prefix: textNumberMatch[1], number: parseInt(textNumberMatch[2]) });
        return { type: 'text-number', prefix: textNumberMatch[1], number: parseInt(textNumberMatch[2]) };
    }

    const dateMatch = value.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (dateMatch) {
        const [_, day, month, year] = dateMatch;
        const paddedDay = day.padStart(2, '0');
        const paddedMonth = month.padStart(2, '0');
        const date = new Date(`${year}-${paddedMonth}-${paddedDay}`);
        if (!isNaN(date.getTime())) {
            console.log('Date pattern detected:', { day, month, year });
            return { type: 'date', base: date, inputFormat: { day: day.length === 1, month: month.length === 1 } };
        }
    }

    const num = parseFloat(value);
    if (!isNaN(num)) {
        if (window.selectedCells.length === 2) {
            const [first, second] = window.selectedCells.map(cellKey => {
                const [row, col] = cellKey.split('-').map(Number);
                return parseFloat(window.currentData[row]?.[col]) || 0;
            });
            const step = second - first;
            console.log('Detected two-cell selection with step:', step);
            return { type: 'number', base: num, step: step };
        }
        console.log('Single cell fill detected, defaulting to step=1');
        return { type: 'number', base: num, step: 1 };
    }

    console.log('Text pattern detected:', value);
    return { type: 'text', base: value };
}

function getNextValue(original, pattern, step) {
    console.log('Getting next value for original:', original, 'pattern:', pattern, 'step:', step);
    switch (pattern.type) {
        case 'date':
            const nextDate = new Date(pattern.base);
            nextDate.setDate(nextDate.getDate() + step);
            const day = pattern.inputFormat.day ? nextDate.getDate() : String(nextDate.getDate()).padStart(2, '0');
            const month = pattern.inputFormat.month ? nextDate.getMonth() + 1 : String(nextDate.getMonth() + 1).padStart(2, '0');
            const year = nextDate.getFullYear();
            return `${day}.${month}.${year}`;
        case 'text-number':
            return `${pattern.prefix}${pattern.number + step}`;
        case 'number':
            return (pattern.base + (pattern.step || 0) * step).toString();
        case 'text':
            return pattern.base;
        default:
            return original;
    }
}