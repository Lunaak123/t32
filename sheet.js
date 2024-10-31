let data = []; // Stores the initial Excel data

// Load and display the first sheet from the Excel file
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        displaySheet(data);
    } catch (error) {
        console.error("Error loading sheet:", error);
    }
}

// Display Excel sheet data in an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');

    // Create headers
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    // Create data rows
    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Apply selected range and highlight cells
function applyOperation() {
    const rowRangeInput = document.getElementById('row-range').value.trim();
    const colRangeInput = document.getElementById('col-range').value.trim();

    if (!rowRangeInput || !colRangeInput) {
        alert('Please enter both row and column ranges.');
        return;
    }

    // Parse row range
    const rowRange = rowRangeInput.split('-').map(Number);
    const rowFrom = rowRange[0];
    const rowTo = rowRange[1] || rowFrom;

    // Parse column range
    const colRange = colRangeInput.split('-');
    const colFrom = colRange[0];
    const colTo = colRange[1] || colFrom;

    highlightRange(rowFrom, rowTo, colFrom, colTo);
}

// Highlight cells in the specified range
function highlightRange(rowFrom, rowTo, colFrom, colTo) {
    const table = document.querySelector('#sheet-content table');
    const colFromIdx = colFrom.charCodeAt(0) - 65;
    const colToIdx = colTo.charCodeAt(0) - 65;

    Array.from(table.rows).forEach((row, rowIndex) => {
        Array.from(row.cells).forEach((cell, colIndex) => {
            cell.classList.remove('highlight');
            if (
                rowIndex >= rowFrom && rowIndex <= rowTo &&
                colIndex >= colFromIdx && colIndex <= colToIdx
            ) {
                cell.classList.add('highlight');
            }
        });
    });
}

// Event listener to apply operation
document.getElementById('apply-operation').addEventListener('click', applyOperation);

// Load the sheet on page load
window.addEventListener('load', () => {
    const fileUrl = new URLSearchParams(window.location.search).get('fileUrl');
    loadExcelSheet(fileUrl);
});
