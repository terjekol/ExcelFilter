function initData() {
    // const originalRows = model.worksheet.getSheetValues();
    // for (let rowIndex = originalRows.length - 1; rowIndex > 1; rowIndex--) {
    //     const row = originalRows[rowIndex];
    //     if (!isFileNameStartingWith1or4(row)) {
    //         model.worksheet.spliceRows(rowIndex, 1);
    //     }

    // }
    // sortAndSum();
    const rows = model.worksheet.getSheetValues();
    for (let rowIndex = 2; rowIndex < rows.length; rowIndex++) {
        const row = rows[rowIndex];
        if (!row || !row[4] || !row[5]) {
            model.unwantedRows.push(rowIndex);
            continue;
        }
        const infoItem = (row[4] || '').trim().toLowerCase();
        const partNumber = (row[5] || '').toLowerCase();        
        const firstDigit = partNumber.trim()[0];
        const startsWith1or4 = '14'.includes(firstDigit);
        if(rowIndex<10)console.log(partNumber, firstDigit, startsWith1or4);
        if (!startsWith1or4 || infoItem != 'no') {
            model.unwantedRows.push(rowIndex);
        }
    }
    model.skipRows = [...model.unwantedRows];
    model.collapseRows = [];
}

function sortAndSum() {
    let rows = rowsAsArrayOfObjects();
    rows.sort((a, b) => parseInt(a[9]) - parseInt(b[9]));
    let totalQuantity = 0;
    for (let index = rows.length - 1; index > 0; index--) {
        const partNo = parseInt(rows[index][9]);
        const previousPartNo = index == 0 ? 0 : parseInt(rows[index - 1][9])
        const quantity = parseInt(rows[index][3]);
        if (partNo == previousPartNo) {
            totalQuantity += quantity;
            rows.splice(index, 1);
        } else if (totalQuantity > 0) {
            rows[index][3] = '' + (totalQuantity + quantity);
            totalQuantity = 0;
        } else {
            totalQuantity = 0;
        }
    }
    clearWorksheet();
    for (let row of rows) {
        model.worksheet.addRow(row);
    }
}

function clearWorksheet() {
    for (let rowIndex = model.worksheet.actualRowCount; rowIndex > 1; rowIndex--) {
        const worksheetRow = model.worksheet.getRow(rowIndex);
        if (!isFileNameStartingWith1or4(worksheetRow)) {
            model.worksheet.spliceRows(rowIndex, 1);
        }
    }
}

function rowsAsArrayOfObjects() {
    let rows = [];
    for (let i = 2; i <= model.worksheet.actualRowCount; i++) {
        let row = [];
        const worksheetRow = model.worksheet.getRow(i);
        for (let j = 1; j <= model.worksheet.columnCount; j++) {
            row[j] = worksheetRow.getCell(j).value;
        }
        rows.push(row);
    }
    return rows;
}

function toggleCollapse(rowIndex) {
    const rows = model.collapseRows;
    const index = rows.indexOf(rowIndex);
    if (index !== -1) {
        rows.splice(index, 1);
    } else {
        rows.push(rowIndex);
    }
    updateView();
}

function toggleRow(rowIndex) {
    const isSelected = !model.skipRows.includes(rowIndex);
    const level = getLevel(rowIndex);
    setSelectedRow(rowIndex, !isSelected, level, true);
    updateView();
}

function setSelectedRow(rowIndex, isSelected, startLevel, force) {
    const level = getLevel(rowIndex);
    if (!force && level <= startLevel) return;
    const skipRows = model.skipRows;
    if (isSelected && (!model.unwantedRows.includes(rowIndex) || force)) {
        const index = skipRows.indexOf(rowIndex);
        if (index != -1) skipRows.splice(index, 1);
    } else {
        if (!skipRows.includes(rowIndex)) skipRows.push(rowIndex);
    }
    setSelectedRow(rowIndex + 1, isSelected, startLevel);
}

function getLevel(rowIndex) {
    const row = model.worksheet.getRow(rowIndex);
    if (!row) return -1;
    const values = row.values;
    if (!values) return -1;
    const level = values[1];
    if (!level) return -1;
    return parseInt(level);
}