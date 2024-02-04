function initData() {
    const rows = model.worksheet.getSheetValues();
    for (let rowIndex = 2; rowIndex < rows.length; rowIndex++) {
        if (isUnwantedRow(rows[rowIndex])) {
            model.unwantedRows.push(rowIndex);
        }
    }
    model.skipRows = [...model.unwantedRows];
    model.collapseRows = [];
}

function isUnwantedRow(row) {
    const infoItem = (row[3]||'').toLowerCase();
    const showRefNo = (row[6] || '').toLowerCase();
    const fileName = (row[7] || '').toLowerCase();
    const name = (row[13] || '').toLowerCase();
    const dependencyType = (row[16] || '').toLowerCase();
    return fileName.includes('_skel.prt')
        || infoItem.includes('yes')
        || fileName.trim()[0] == '1'
        || showRefNo.trim() == 'no'
        || dependencyType.trim() == 'suppressed member'
        || dependencyType.trim() == 'skeleton member'
        || (name.includes('99') && name.includes('part'))
        || (name == 'pipe' || name == 'plate');
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