function initData() {
    const rows = model.worksheet.getSheetValues();
    for (let rowIndex = rows.length-1; rowIndex > 1; rowIndex--) {
        if(!isFileNameStartingWith1or4(rows[rowIndex])){
            model.worksheet.spliceRows(rowIndex, 1);
        }
    }
    for (let rowIndex = 2; rowIndex < rows.length; rowIndex++) {
        if (isUnwantedRow(rows[rowIndex])) {
            model.unwantedRows.push(rowIndex);
        }
    }
    model.skipRows = [...model.unwantedRows];
    model.collapseRows = [];
}

function isFileNameStartingWith1or4(row) {
    const number = (row[8] || '').toLowerCase();
    const firstDigit =  number.trim()[0];
    return '14'.includes(firstDigit);
}

function isUnwantedRow(row) {
    const infoItem = (row[3] || '').trim().toLowerCase();
    return infoItem != 'no';
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