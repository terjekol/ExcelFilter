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
    var workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(model.data);
    var excelDataDiv = document.getElementById('excelData');
    excelDataDiv.innerHTML = '';

    workbook.eachSheet(function (worksheet, sheetId) {
        var sheetDiv = document.createElement('div');
        sheetDiv.classList.add('sheet-container');
        sheetDiv.innerHTML = '<h3>' + worksheet.name + '</h3>';

        var tableHtml = '<table class="excel-table">' +
            '<thead><tr>' +
            worksheet.getRow(1).values.map(value => '<th>' + (value || '') + '</th>').join('') +
            '</tr></thead>' +
            '<tbody>' +
            worksheet.getSheetValues().slice(1).map((row, rowIndex) => {
                var level = row[0];
                return '<tr>' +
                    row.map((content, colIndex) =>
                        formatCell(content, colIndex, level)).join('') +
                    '</tr>';
            }).join('') +
            '</tbody></table>';

        sheetDiv.innerHTML += tableHtml;
        excelDataDiv.appendChild(sheetDiv);
    });
}

function formatCell(content, colIndex, level) {
    if (!content) return '<td></td>';
    var checkbox = '<input type="checkbox" checked data-level="' + level + '"/>';
    const pre = colIndex == 2 ? checkbox : ''
    return `<td>${pre + content.replaceAll(' ', '&nbsp;')}</td>`;
}

function updateCheckboxes(checkbox, checkboxes) {
    var checked = checkbox.checked;
    var levelColumnIndex = 0; // Anta at level-kolonnen alltid er den første kolonnen
    var selectedRowLevel = parseInt(checkbox.getAttribute('data-level'), 10);

    // Finn indeksen til checkboxen som ble endret
    var checkboxIndex = Array.from(checkboxes).indexOf(checkbox);

    // Gå gjennom de påfølgende radene og oppdater checkboxene
    for (var i = checkboxIndex + 1; i < checkboxes.length; i++) {
        var currentRowLevel = parseInt(checkboxes[i].getAttribute('data-level'), 10);

        if (currentRowLevel > selectedRowLevel) {
            checkboxes[i].checked = checked;
        } else {
            // Hvis raden har samme nivå eller lavere, stopp oppdateringen
            break;
        }
    }
}
