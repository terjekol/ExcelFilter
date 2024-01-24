const model = {
    colIndexes: [1, 7, 8, 9, 10, 11, 12, 13, 14],
    allColIndexes: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16],
    skipRows: [],
    unwantedRows: [],
    hideCols: true,
};

function handleFile() {
    var fileInput = document.getElementById('fileInput');
    var file = fileInput.files[0];
    if (file) {
        model.fileName = file.name;
        var reader = new FileReader();
        reader.onload = async function (e) {
            var data = new Uint8Array(e.target.result);
            model.workbook = new ExcelJS.Workbook();
            await model.workbook.xlsx.load(data);
            model.workbook.eachSheet(worksheet => model.worksheet = model.worksheet || worksheet);
            initData();
            updateView();
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Velg en Excel-fil først.');
    }
}

function initData() {
    const rows = model.worksheet.getSheetValues();
    for (let rowIndex = 2; rowIndex < rows.length; rowIndex++) {
        if (isUnwantedRow(rows[rowIndex])) {
            model.unwantedRows.push(rowIndex);
        }
    }
    model.skipRows = [...model.unwantedRows];
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

function updateView() {
    var excelDataDiv = document.getElementById('excelData');
    excelDataDiv.innerHTML = '';
    const worksheet = model.worksheet;

    var sheetDiv = document.createElement('div');
    sheetDiv.classList.add('sheet-container');
    sheetDiv.innerHTML = /*HTML*/`
        <input 
        type="checkbox" 
        ${model.hideCols ? 'checked' : ''}
        onchange="model.hideCols=!model.hideCols;updateView()"
        />
        Skjul kolonner
        <h3>${worksheet.name}</h3>    
    `;


    var tableHtml = '<table class="excel-table">' +
        '<thead><tr>' +
        worksheet.getRow(1).values.map((value, index) => (model.colIndexes.includes(index) || !model.hideCols) ? '<th>' + (value || '') + '</th>' : '').join('') +
        '</tr></thead>' +
        '<tbody>' +
        worksheet.getSheetValues().map((row, rowIndex) => {
            if (rowIndex < 2) return '';
            let html = '';
            const isUnwanted = model.unwantedRows.includes(rowIndex);
            const style = isUnwanted ? `style="background-color: #ffeeee; color: darkred"` : '';
            const indexes = model.hideCols ? model.colIndexes : model.allColIndexes;
            for (let colIndex of indexes) {
                html += formatCell(row[colIndex] || '', colIndex, rowIndex);
            }
            return `<tr ${style}>` + html + '</tr>';
        }).join('') +
        '</tbody></table>';

    sheetDiv.innerHTML += tableHtml;
    excelDataDiv.appendChild(sheetDiv);
}

function formatCell(content, colIndex, rowIndex) {
    if (typeof (content) != 'string') content = '';
    const checked = model.skipRows.includes(rowIndex) ? '' : 'checked';
    if (colIndex == 12) content = content.substr(0, 10);
    if (colIndex != 7) return `<td>${content}</td>`;
    const spaceCount = countLeadingSpaces(content);
    const spaces = '&nbsp;'.repeat(spaceCount);
    var checkbox = `<input onclick="toggleRow(${rowIndex})" ${checked} type="checkbox"/>`;
    return `<td>${spaces + checkbox + content}</td>`;
}

function countLeadingSpaces(txt) {
    let index = 0;
    while (txt[index] == ' ') {
        index++;
    }
    return index;
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

async function downloadFile() {
    const skipRows = [...model.skipRows];
    skipRows.sort((a, b) => b - a);
    for (let rowIndex of skipRows) {
        model.worksheet.spliceRows(rowIndex, 1);
    }

    var excelBuffer = await model.workbook.xlsx.writeBuffer();
    var blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'fixed_' + model.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}
