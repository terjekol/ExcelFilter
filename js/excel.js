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