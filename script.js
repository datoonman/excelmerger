document.getElementById('mergeButton').addEventListener('click', processFile);

function processFile() {
    const input = document.getElementById('fileInput');
    const statusMessage = document.getElementById('statusMessage');

    if (!input.files[0]) {
        statusMessage.textContent = 'Please upload an Excel file first.';
        return;
    }

    statusMessage.textContent = 'Processing...';

    const file = input.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });
            const headers = json[0];
            const questions = json[1];
            const rows = json.slice(2); // Start from row 3 onwards

            const mergedData = mergeRows(rows, headers.indexOf('Response ID'));

            // Create new workbook
            const newWorkbook = XLSX.utils.book_new();
            const newSheet = XLSX.utils.aoa_to_sheet([headers, questions, ...mergedData]);

            // Apply hyperlinks and fill color to new sheet
            applyHyperlinks(newSheet, sheet);

            // Check for discrepancies and list them in a new sheet
            const discrepanciesSheet = listDiscrepancies(newSheet);
            XLSX.utils.book_append_sheet(newWorkbook, discrepanciesSheet, 'Discrepancies');

            // Add merged data sheet
            XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Merged Data');

            // Save the workbook
            const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'binary' });
            saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), 'merged_output.xlsx');

            statusMessage.textContent = 'File processed and saved successfully!';
        } catch (error) {
            statusMessage.textContent = 'An error occurred during processing. Please try again.';
            console.error('Error:', error.message);
            console.error(error.stack);
        }
    };

    reader.readAsArrayBuffer(file);
}

function mergeRows(rows, idIndex) {
    const mergedData = {};

    rows.forEach((row) => {
        const id = row[idIndex];

        if (!id) {
            return;
        }

        if (!mergedData[id]) {
            mergedData[id] = [...row];
        } else {
            mergedData[id].forEach((cell, i) => {
                if (cell !== row[i] && row[i]) {
                    const formattedCell = formatCell(cell);
                    const formattedRowCell = formatCell(row[i]);

                    if (formattedCell) {
                        mergedData[id][i] = `${formattedCell} | ${formattedRowCell}`;
                    } else {
                        mergedData[id][i] = formattedRowCell;
                    }
                }
            });
        }
    });

    return Object.values(mergedData);
}

function listDiscrepancies(sheet) {
    const discrepancies = [['Cell Reference', 'Discrepancy']];

    Object.keys(sheet).forEach(cellAddress => {
        const cell = sheet[cellAddress];
        if (cell && cell.v && typeof cell.v === 'string' && cell.v.includes('|')) {
            discrepancies.push([cellAddress, cell.v]);
        }
    });

    return XLSX.utils.aoa_to_sheet(discrepancies);
}

function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
}

function applyHyperlinks(newSheet, originalSheet) {
    Object.keys(originalSheet).forEach(cellAddress => {
        const originalCell = originalSheet[cellAddress];
        let newCell = newSheet[cellAddress];

        if (originalCell && originalCell.l) {
            // Initialize newCell if it doesn't exist
            if (!newCell) {
                newCell = { v: '' }; // Initialize with empty value if needed
            }

            newCell.l = originalCell.l; // Copy hyperlink properties
            newSheet[cellAddress] = newCell; // Ensure the cell is updated in newSheet
        }
    });
}

function formatCell(cell) {
    if (typeof cell === 'number') {
        // Assume it's a date if it's a number and format accordingly
        return XLSX.SSF.format('m/d/yyyy h:mm:ss', cell);
    }
    return cell;
}
