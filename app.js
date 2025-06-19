let db;
let csvData = [];
let univer, univerSheet;
let univerResults;

// Initialize SQLite
initSqlJs({ locateFile: file => `https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/${file}` }).then(SQL => {
  db = new SQL.Database();
});

document.getElementById('csvInput').addEventListener('change', handleCSVUpload);

function handleCSVUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: function(results) {
      csvData = results.data;
      createSheet(csvData);
      createSQLiteTable(csvData);
    }
  });
}

function createSheet(data) {
  const headers = Object.keys(data[0]);
  const rows = data.map(row => headers.map(h => row[h]));

  const univerData = {
    name: 'Sheet1',
    sheet: {
      Sheet1: {
        name: 'Data',
        cellData: {
          0: Object.fromEntries(headers.map((h, i) => [i, { v: h }])),
          ...Object.fromEntries(rows.map((r, i) => [i + 1, Object.fromEntries(r.map((val, j) => [j, { v: val }]))]))
        }
      }
    }
  };

  document.getElementById('spreadsheet').innerHTML = '';
  univer = new UniverCore.UniverSheet();
  univer.installPlugin(new UniverUI.UniverUI());
  univer.installPlugin(new UniverSheets.UniverSheets());

  univer.createUniverSheet(document.getElementById('spreadsheet'), univerData);
}

function createSQLiteTable(data) {
  const headers = Object.keys(data[0]);
  db.run('DROP TABLE IF EXISTS data');
  db.run(`CREATE TABLE data (${headers.map(h => `"${h}" TEXT`).join(', ')})`);

  const stmt = db.prepare(`INSERT INTO data VALUES (${headers.map(() => '?').join(',')})`);
  data.forEach(row => stmt.run(headers.map(h => row[h])));
  stmt.free();
}

function runSQL() {
  const query = document.getElementById('sqlQuery').value;
  try {
    const result = db.exec(query);
    if (!result[0]) return alert("No results");

    const columns = result[0].columns;
    const values = result[0].values;

    showSQLResults(columns, values);
  } catch (e) {
    alert("SQL Error: " + e.message);
  }
}

function showSQLResults(columns, values) {
  const data = {
    name: 'SQLSheet',
    sheet: {
      SQLSheet: {
        name: 'SQL Results',
        cellData: {
          0: Object.fromEntries(columns.map((c, i) => [i, { v: c }])),
          ...Object.fromEntries(values.map((row, i) => [i + 1, Object.fromEntries(row.map((val, j) => [j, { v: val }]))]))
        }
      }
    }
  };

  document.getElementById('sql-results').innerHTML = '';
  univerResults = new UniverCore.UniverSheet();
  univerResults.installPlugin(new UniverUI.UniverUI());
  univerResults.installPlugin(new UniverSheets.UniverSheets());

  univerResults.createUniverSheet(document.getElementById('sql-results'), data);
}

function saveAsCSV() {
  if (!univerResults) return alert("No SQL results to save.");

  // Very basic extraction from spreadsheet back to CSV
  const sheet = univerResults.getContext().getUniver().getWorkBook().getSheetByIndex(0);
  const cellMatrix = sheet.getCellMatrix().getMatrix();

  const csvRows = [];
  for (const [rowIndex, rowData] of Object.entries(cellMatrix)) {
    const row = [];
    for (let col = 0; col < Object.keys(rowData).length; col++) {
      const cell = rowData[col];
      row.push(cell?.v ?? "");
    }
    csvRows.push(row.join(","));
  }

  const blob = new Blob([csvRows.join("\n")], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "sql-results.csv";
  a.click();
}
