function handleFile() {
    var fileInput = document.getElementById('fileInput');
    var excelDataDiv = document.getElementById('excelData');

    var file = fileInput.files[0];
    if (file) {
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = new Uint8Array(e.target.result);
            window.data = data;            
            processData(data);
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Velg en Excel-fil først.');
    }
}

async function processData(data) {
    var workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);

    var excelDataDiv = document.getElementById('excelData');
    excelDataDiv.innerHTML = '';

    workbook.eachSheet(function (worksheet, sheetId) {
        var sheetDiv = document.createElement('div');
        sheetDiv.classList.add('sheet-container');
        sheetDiv.innerHTML = `<h3>${worksheet.name}</h3>`;

        var tableHtml = /*HTML*/`
            <table class="excel-table">
                <thead>
                    <tr>
                        ${worksheet.getRow(1).values.map(value => `<th>${value || ''}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>
                    ${worksheet.getSheetValues().slice(1).map(row => /*HTML*/`
                        <tr>
                            ${row.map(formatCell).join('')}
                        </tr>`).join('')}
                </tbody>
            </table>`;

        sheetDiv.innerHTML += tableHtml;
        excelDataDiv.appendChild(sheetDiv);
    });
}

function formatCell(content) {
    if (!content) return '<td></td>';
    return '<td>' + content.replaceAll(' ', '&nbsp;') + '</td>';
}
