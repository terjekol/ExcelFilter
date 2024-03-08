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

    let currentCollapseLevel = null;
    var tableHtml = '<table class="excel-table">' +
        '<thead><tr><td></td>' +
        worksheet.getRow(1).values.map((value, index) => (model.colIndexes.includes(index) || !model.hideCols) ? '<th>' + (value || '') + '</th>' : '').join('') +
        '</tr></thead>' +
        '<tbody>' +
        worksheet.getSheetValues().map((row, rowIndex) => {
            if (rowIndex < 2) return '';
            const isCollapsed = model.collapseRows.includes(rowIndex);
            const level = parseInt(row[1]);
            if (isCollapsed) {
                currentCollapseLevel = level;
            } else if (currentCollapseLevel !== null) {
                if (level > currentCollapseLevel) {
                    return '';
                } else {
                    currentCollapseLevel = null;
                }
            }
            let html = '';
            const indexes = model.hideCols ? model.colIndexes : model.allColIndexes;
            for (let colIndex of indexes) {
                html += formatCell(row[colIndex] || '', colIndex, rowIndex);
            }
            const isUnwanted = model.unwantedRows.includes(rowIndex);
            const style =
                isUnwanted ? `style="background-color: #ffeeee; color: darkred"` :
                    isCollapsed ? `style="background-color: lightgray"` :
                        '';
            const collapseChar = isCollapsed ? '+' : '−';
            return `<tr ${style}><td><button onclick="toggleCollapse(${rowIndex})">${collapseChar}</button></td>` + html + '</tr>';
        }).join('') +
        '</tbody></table>';

    sheetDiv.innerHTML += tableHtml;
    excelDataDiv.appendChild(sheetDiv);
}

function formatCell(content, colIndex, rowIndex) {
    if (typeof (content) != 'string') content = '';
    const checked = model.skipRows.includes(rowIndex) ? '' : 'checked';
    if (colIndex == 12) content = content.substr(0, 10);
    if (colIndex != 8) return `<td>${content}</td>`;
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