document.addEventListener('DOMContentLoaded', () => {
  const drop = document.getElementById('drop');
  const file = document.getElementById('file');
  const html = document.getElementById('html');
  const autofillOptions = document.getElementById('autofill-options');
  const patternInfo = document.getElementById('pattern-info');
  
  let workbook = null;
  let selectedRange = null;
  let autofillStart = null;
  let autofillDirection = null;
  let autofillHandle = null;
  let isDraggingAutofill = false;
  
  // Initialize event listeners
  drop.addEventListener('dragover', (e) => {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
  });

  drop.addEventListener('drop', (e) => {
    e.stopPropagation();
    e.preventDefault();
    processFile(e.dataTransfer.files[0]);
  });

  file.addEventListener('change', (e) => {
    processFile(e.target.files[0]);
  });

  // Process Excel file
  function processFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array' });
      displaySheet(workbook.Sheets[workbook.SheetNames[0]]);
    };
    reader.readAsArrayBuffer(file);
  }

  // Display sheet data
  function displaySheet(sheet) {
    html.innerHTML = XLSX.utils.sheet_to_html(sheet, { editable: true });
    
    // Add event listeners to cells
    const table = html.querySelector('table');
    if (table) {
      table.addEventListener('mousedown', handleMouseDown);
      table.addEventListener('mousemove', handleMouseMove);
      table.addEventListener('mouseup', handleMouseUp);
      
      // Make sure the table is wide enough
      const cols = table.querySelectorAll('tr')[0].children;
      for (let i = 0; i < cols.length; i++) {
        cols[i].style.minWidth = '100px';
      }
    }
  }

  // Mouse down handler
  function handleMouseDown(e) {
    const cell = e.target.closest('td, th');
    if (!cell) return;
    
    // Check if clicking on autofill handle
    if (e.target.classList.contains('autofill-handle')) {
      isDraggingAutofill = true;
      autofillStart = getCellPosition(cell);
      return;
    }
    
    // Remove any existing autofill handle
    removeAutofillHandle();
    
    // Select cell
    if (e.shiftKey && selectedRange) {
      // Extend selection
      const endCell = getCellPosition(cell);
      selectRange(selectedRange.start, endCell);
    } else {
      // New selection
      const pos = getCellPosition(cell);
      selectedRange = { start: pos, end: pos };
      updateSelection();
    }
    
    // Add autofill handle to the last selected cell
    addAutofillHandle();
  }

  // Mouse move handler
  function handleMouseMove(e) {
    if (isDraggingAutofill) {
      const cell = e.target.closest('td, th');
      if (!cell) return;
      
      const currentPos = getCellPosition(cell);
      autofillDirection = determineAutofillDirection(autofillStart, currentPos);
      previewAutofill(autofillStart, currentPos);
    } else if (e.buttons === 1 && selectedRange) {
      // Regular selection dragging
      const cell = e.target.closest('td, th');
      if (!cell) return;
      
      const pos = getCellPosition(cell);
      selectRange(selectedRange.start, pos);
    }
  }

  // Mouse up handler
  function handleMouseUp(e) {
    if (isDraggingAutofill) {
      const cell = e.target.closest('td, th');
      if (cell) {
        const currentPos = getCellPosition(cell);
        showAutofillOptions(autofillStart, currentPos);
      }
      isDraggingAutofill = false;
    }
    clearAutofillPreview();
  }

  // Get cell position (row and column)
  function getCellPosition(cell) {
    const row = cell.parentNode.rowIndex;
    const col = cell.cellIndex;
    return { row, col };
  }

  // Select range of cells
  function selectRange(start, end) {
    // Normalize range (start should be top-left, end should be bottom-right)
    const minRow = Math.min(start.row, end.row);
    const maxRow = Math.max(start.row, end.row);
    const minCol = Math.min(start.col, end.col);
    const maxCol = Math.max(start.col, end.col);
    
    selectedRange = {
      start: { row: minRow, col: minCol },
      end: { row: maxRow, col: maxCol }
    };
    
    updateSelection();
  }

  // Update cell selection UI
  function updateSelection() {
    // Clear previous selection
    const selectedCells = document.querySelectorAll('.selected');
    selectedCells.forEach(cell => cell.classList.remove('selected'));
    
    if (!selectedRange) return;
    
    // Select new range
    const table = html.querySelector('table');
    if (!table) return;
    
    for (let r = selectedRange.start.row; r <= selectedRange.end.row; r++) {
      const row = table.rows[r];
      if (!row) continue;
      
      for (let c = selectedRange.start.col; c <= selectedRange.end.col; c++) {
        const cell = row.cells[c];
        if (cell) cell.classList.add('selected');
      }
    }
  }

  // Add autofill handle to the last selected cell
  function addAutofillHandle() {
    if (!selectedRange) return;
    
    const table = html.querySelector('table');
    if (!table) return;
    
    const lastRow = table.rows[selectedRange.end.row];
    if (!lastRow) return;
    
    const lastCell = lastRow.cells[selectedRange.end.col];
    if (!lastCell) return;
    
    // Remove any existing handle
    removeAutofillHandle();
    
    // Create new handle
    autofillHandle = document.createElement('div');
    autofillHandle.className = 'autofill-handle';
    lastCell.style.position = 'relative';
    lastCell.appendChild(autofillHandle);
  }

  // Remove autofill handle
  function removeAutofillHandle() {
    if (autofillHandle) {
      autofillHandle.remove();
      autofillHandle = null;
    }
  }

  // Determine autofill direction
  function determineAutofillDirection(start, current) {
    const rowDiff = current.row - start.row;
    const colDiff = current.col - start.col;
    
    if (Math.abs(rowDiff) > Math.abs(colDiff)) {
      return rowDiff > 0 ? 'down' : 'up';
    } else {
      return colDiff > 0 ? 'right' : 'left';
    }
  }

  // Preview autofill
  function previewAutofill(start, current) {
    clearAutofillPreview();
    
    const table = html.querySelector('table');
    if (!table) return;
    
    // Determine range to preview
    let previewCells = [];
    if (autofillDirection === 'right') {
      for (let c = start.col; c <= current.col; c++) {
        for (let r = selectedRange.start.row; r <= selectedRange.end.row; r++) {
          const cell = table.rows[r]?.cells[c];
          if (cell) previewCells.push(cell);
        }
      }
    } else if (autofillDirection === 'down') {
      for (let r = start.row; r <= current.row; r++) {
        for (let c = selectedRange.start.col; c <= selectedRange.end.col; c++) {
          const cell = table.rows[r]?.cells[c];
          if (cell) previewCells.push(cell);
        }
      }
    }
    
    // Highlight preview cells
    previewCells.forEach(cell => {
      if (!cell.classList.contains('selected')) {
        cell.classList.add('autofill-preview');
      }
    });
  }

  // Clear autofill preview
  function clearAutofillPreview() {
    const previewCells = document.querySelectorAll('.autofill-preview');
    previewCells.forEach(cell => cell.classList.remove('autofill-preview'));
  }

  // Show autofill options
  function showAutofillOptions(start, end) {
    if (!selectedRange) return;
    
    // Detect pattern in the selected range
    const pattern = detectPattern();
    patternInfo.textContent = pattern || 'None';
    
    // Position and show options
    autofillOptions.style.display = 'block';
    autofillOptions.style.position = 'absolute';
    autofillOptions.style.left = `${end.col * 100 + 50}px`;
    autofillOptions.style.top = `${end.row * 30 + 100}px`;
    
    // Set up button handlers
    document.getElementById('fill-series').onclick = () => applyAutofill('series');
    document.getElementById('fill-formulas').onclick = () => applyAutofill('formulas');
    document.getElementById('fill-copy').onclick = () => applyAutofill('copy');
    document.getElementById('cancel-autofill').onclick = cancelAutofill;
  }

  // Detect pattern in selected range
  function detectPattern() {
    const table = html.querySelector('table');
    if (!table || !selectedRange) return null;
    
    const values = [];
    const isSingleColumn = selectedRange.start.col === selectedRange.end.col;
    const isSingleRow = selectedRange.start.row === selectedRange.end.row;
    
    // Extract values from selected range
    for (let r = selectedRange.start.row; r <= selectedRange.end.row; r++) {
      const row = [];
      for (let c = selectedRange.start.col; c <= selectedRange.end.col; c++) {
        const cell = table.rows[r]?.cells[c];
        row.push(cell?.textContent.trim());
      }
      values.push(row);
    }
    
    // Check for numeric series
    if (isSingleColumn || isSingleRow) {
      const flatValues = isSingleColumn 
        ? values.map(row => row[0])
        : values[0];
      
      if (flatValues.every(v => !isNaN(parseFloat(v)))) {
        const nums = flatValues.map(Number);
        const diffs = [];
        for (let i = 1; i < nums.length; i++) {
          diffs.push(nums[i] - nums[i - 1]);
        }
        
        if (diffs.every(d => d === diffs[0])) {
          return `Linear series (step ${diffs[0]})`;
        }
      }
    }
    
    // Check for dates
    if (isSingleColumn || isSingleRow) {
      const flatValues = isSingleColumn 
        ? values.map(row => row[0])
        : values[0];
      
      if (flatValues.every(v => isPotentialDate(v))) {
        return 'Date series';
      }
    }
    
    // Check for formulas
    if (values.flat().some(v => v.startsWith('='))) {
      return 'Contains formulas';
    }
    
    return null;
  }

  // Check if a string looks like a date
  function isPotentialDate(str) {
    // Simple checks for common date formats
    if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(str)) return true;
    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) return true;
    if (/^[A-Za-z]{3} \d{1,2},? \d{4}$/.test(str)) return true;
    return false;
  }

  // Apply autofill
  function applyAutofill(mode) {
    if (!selectedRange || !autofillStart || !autofillDirection) return;
    
    const table = html.querySelector('table');
    if (!table) return;
    
    // Determine target range
    const targetRange = calculateTargetRange();
    
    // Get source values
    const sourceValues = getSourceValues();
    
    // Apply autofill based on mode
    if (mode === 'series') {
      fillSeries(sourceValues, targetRange);
    } else if (mode === 'formulas') {
      fillFormulas(sourceValues, targetRange);
    } else if (mode === 'copy') {
      fillCopy(sourceValues, targetRange);
    }
    
    // Clean up
    cancelAutofill();
  }

  // Calculate target range for autofill
  function calculateTargetRange() {
    if (!selectedRange || !autofillStart) return null;
    
    const table = html.querySelector('table');
    if (!table) return null;
    
    // Get the current mouse position (approximated by the options panel position)
    const endCol = parseInt(autofillOptions.style.left) / 100 - 0.5;
    const endRow = (parseInt(autofillOptions.style.top) - 100) / 30;
    
    if (autofillDirection === 'right') {
      return {
        startRow: selectedRange.start.row,
        endRow: selectedRange.end.row,
        startCol: selectedRange.end.col + 1,
        endCol: Math.round(endCol)
      };
    } else if (autofillDirection === 'down') {
      return {
        startRow: selectedRange.end.row + 1,
        endRow: Math.round(endRow),
        startCol: selectedRange.start.col,
        endCol: selectedRange.end.col
      };
    }
    
    return null;
  }

  // Get source values from selected range
  function getSourceValues() {
    const table = html.querySelector('table');
    if (!table || !selectedRange) return [];
    
    const values = [];
    for (let r = selectedRange.start.row; r <= selectedRange.end.row; r++) {
      const row = [];
      for (let c = selectedRange.start.col; c <= selectedRange.end.col; c++) {
        const cell = table.rows[r]?.cells[c];
        row.push(cell?.textContent.trim());
      }
      values.push(row);
    }
    
    return values;
  }

  // Fill series
  function fillSeries(sourceValues, targetRange) {
    const table = html.querySelector('table');
    if (!table || !targetRange) return;
    
    const isVertical = autofillDirection === 'down';
    
    // For simplicity, handle only single row/column cases
    if (isVertical && sourceValues.length === 1) {
      // Copy the single row down
      for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
        for (let c = targetRange.startCol; c <= targetRange.endCol; c++) {
          const cell = ensureCellExists(table, r, c);
          cell.textContent = sourceValues[0][c - selectedRange.start.col];
        }
      }
    } else if (!isVertical && sourceValues[0].length === 1) {
      // Copy the single column right
      for (let c = targetRange.startCol; c <= targetRange.endCol; c++) {
        for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
          const cell = ensureCellExists(table, r, c);
          cell.textContent = sourceValues[r - selectedRange.start.row][0];
        }
      }
    } else {
      // For multi-cell ranges, try to detect and extend patterns
      const pattern = detectPattern();
      
      if (pattern && pattern.includes('Linear series')) {
        const step = parseFloat(pattern.match(/step (-?\d+)/)[1]);
        const nums = sourceValues.flat().map(Number);
        const lastNum = nums[nums.length - 1];
        
        if (isVertical) {
          let current = lastNum;
          for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
            current += step;
            const cell = ensureCellExists(table, r, targetRange.startCol);
            cell.textContent = current;
          }
        } else {
          let current = lastNum;
          for (let c = targetRange.startCol; c <= targetRange.endCol; c++) {
            current += step;
            const cell = ensureCellExists(table, targetRange.startRow, c);
            cell.textContent = current;
          }
        }
      } else {
        // Default to copy if no pattern detected
        fillCopy(sourceValues, targetRange);
      }
    }
  }

  // Fill formulas (with adjustment)
  function fillFormulas(sourceValues, targetRange) {
    const table = html.querySelector('table');
    if (!table || !targetRange) return;
    
    const rowOffset = targetRange.startRow - selectedRange.start.row;
    const colOffset = targetRange.startCol - selectedRange.start.col;
    
    for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
      for (let c = targetRange.startCol; c <= targetRange.endCol; c++) {
        const srcRow = r - rowOffset;
        const srcCol = c - colOffset;
        
        if (srcRow >= selectedRange.start.row && srcRow <= selectedRange.end.row &&
            srcCol >= selectedRange.start.col && srcCol <= selectedRange.end.col) {
          const srcValue = sourceValues[srcRow - selectedRange.start.row][srcCol - selectedRange.start.col];
          
          if (srcValue.startsWith('=')) {
            const cell = ensureCellExists(table, r, c);
            cell.textContent = adjustFormula(srcValue, rowOffset, colOffset);
          }
        }
      }
    }
  }

  // Adjust formula references
  function adjustFormula(formula, rowOffset, colOffset) {
    // Simple formula adjustment - would need more sophisticated parsing for real use
    return formula.replace(/([A-Z]+)(\d+)/g, (match, colLetters, rowNum) => {
      // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
      let colNum = 0;
      for (let i = 0; i < colLetters.length; i++) {
        colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
      }
      
      // Apply offsets
      const newRow = parseInt(rowNum) + rowOffset;
      const newCol = colNum + colOffset;
      
      // Convert back to letters
      let remaining = newCol;
      let newColLetters = '';
      while (remaining > 0) {
        const code = (remaining - 1) % 26;
        newColLetters = String.fromCharCode(65 + code) + newColLetters;
        remaining = Math.floor((remaining - 1) / 26);
      }
      
      return `${newColLetters}${newRow}`;
    });
  }

  // Fill with copied values
  function fillCopy(sourceValues, targetRange) {
    const table = html.querySelector('table');
    if (!table || !targetRange) return;
    
    const rowRepeat = Math.ceil((targetRange.endRow - targetRange.startRow + 1) / sourceValues.length);
    const colRepeat = Math.ceil((targetRange.endCol - targetRange.startCol + 1) / sourceValues[0].length);
    
    for (let r = targetRange.startRow; r <= targetRange.endRow; r++) {
      for (let c = targetRange.startCol; c <= targetRange.endCol; c++) {
        const srcRow = (r - targetRange.startRow) % sourceValues.length;
        const srcCol = (c - targetRange.startCol) % sourceValues[0].length;
        const cell = ensureCellExists(table, r, c);
        cell.textContent = sourceValues[srcRow][srcCol];
      }
    }
  }

  // Ensure cell exists in table (adds rows/columns if needed)
  function ensureCellExists(table, row, col) {
    // Add rows if needed
    while (table.rows.length <= row) {
      const newRow = table.insertRow();
      for (let i = 0; i < table.rows[0].cells.length; i++) {
        newRow.insertCell();
      }
    }
    
    // Add columns if needed
    while (table.rows[row].cells.length <= col) {
      for (let r = 0; r < table.rows.length; r++) {
        table.rows[r].insertCell();
      }
    }
    
    return table.rows[row].cells[col];
  }

  // Cancel autofill
  function cancelAutofill() {
    autofillOptions.style.display = 'none';
    clearAutofillPreview();
    removeAutofillHandle();
  }
});
