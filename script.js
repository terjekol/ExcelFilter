const model = {
    skipRows: [],
    unwantedRows: [],
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
        if (isSkipRow(rows[rowIndex])) {
            model.unwantedRows.push(rowIndex);
        }
    }
    model.skipRows = [...model.unwantedRows];
}

function isSkipRow(row) {
    const showRefNo = (row[6] || '').toLowerCase();
    const fileName = (row[7] || '').toLowerCase();
    const name = (row[14] || '').toLowerCase();
    const dependencyType = (row[16] || '').toLowerCase();
    return fileName.includes('_skel.prt')
        || fileName.trim()[0] == '1'
        || showRefNo.trim() == 'no'
        || dependencyType.trim() == 'suppressed member'
        || (name.includes('99') && name.includes('part'))
        || (name == 'pipe' || name == 'plate');
}

function updateView() {
    var excelDataDiv = document.getElementById('excelData');
    excelDataDiv.innerHTML = '';
    const worksheet = model.worksheet;

    var sheetDiv = document.createElement('div');
    sheetDiv.classList.add('sheet-container');
    sheetDiv.innerHTML = '<h3>' + worksheet.name + '</h3>';

    var tableHtml = '<table class="excel-table">' +
        '<thead><tr>' +
        worksheet.getRow(1).values.map(value => '<th>' + (value || '') + '</th>').join('') +
        '</tr></thead>' +
        '<tbody>' +
        worksheet.getSheetValues().map((row, rowIndex) => {
            if (rowIndex < 2) return '';
            let html = '';
            const isUnwanted = model.unwantedRows.includes(rowIndex);
            const style = isUnwanted ? `style="color: red"` : '';
            for (let colIndex = 1; colIndex < row.length; colIndex++) {
                html += formatCell(row[colIndex] || '', colIndex, rowIndex);
            }
            return `<tr ${style}>` + html + '</tr>';
        }).join('') +
        '</tbody></table>';

    sheetDiv.innerHTML += tableHtml;
    excelDataDiv.appendChild(sheetDiv);
}

function formatCell(content, colIndex, rowIndex) {
    //if (!content) return '<td>&nbsp;</td>';
    if (typeof (content) != 'string') content = '';
    const checked = model.skipRows.includes(rowIndex) ? '' : 'checked';
    var checkbox = `<input onclick="toggleRow(${rowIndex})" ${checked} type="checkbox"/>`;
    const pre = colIndex == 7 ? checkbox : '';
    return `<td>${pre + content.replaceAll(' ', '&nbsp;')}</td>`;
}

function toggleRow(rowIndex) {
    const isSelected = !model.skipRows.includes(rowIndex);
    const level = getLevel(rowIndex);
    setSelectedRow(rowIndex, !isSelected, level, true);
    updateView();
}

function setSelectedRow(rowIndex, isSelected, startLevel, force) {
    const level = getLevel(rowIndex);
    if (!force && level <= startLevel || model.unwantedRows.includes(rowIndex)) return;
    const skipRows = model.skipRows;
    if (isSelected) {
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
