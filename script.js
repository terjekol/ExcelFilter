const model = {
    skipRows: [],
};

function handleFile() {
    var fileInput = document.getElementById('fileInput');

    var file = fileInput.files[0];
    if (file) {
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = new Uint8Array(e.target.result);
            model.data = data;
            updateView();
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Velg en Excel-fil først.');
    }
}

async function updateView() {
    let workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(model.data);
    var excelDataDiv = document.getElementById('excelData');
    excelDataDiv.innerHTML = '';

    workbook.eachSheet(function (worksheet, sheetId) {
        model.worksheet = worksheet;
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
                var level = row[0];
                return '<tr>' +
                    row.map((content, colIndex) =>
                        formatCell(content, colIndex, level, rowIndex)).join('') +
                    '</tr>';
            }).join('') +
            '</tbody></table>';

        sheetDiv.innerHTML += tableHtml;
        excelDataDiv.appendChild(sheetDiv);
        return;
    });
}

function formatCell(content, colIndex, level, rowIndex) {
    if (!content) return '<td></td>';
    const checked = model.skipRows.includes(rowIndex) ? '' : 'checked';
    var checkbox = `<input onclick="toggleRow(${rowIndex})" ${checked} type="checkbox"/>`;
    const pre = colIndex == 2 ? checkbox : ''
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
    if (!force && level <= startLevel) return;
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
    if(!values) return -1;
    const level = values[1];
    if(!level)return -1;
    return parseInt(level);
}