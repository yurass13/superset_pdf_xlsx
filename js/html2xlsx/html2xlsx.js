// require xlsx-js-style 1.2.0-beta

function isNumeric(str) {
    if (typeof str != "string") return false // we only process strings!
    str = str.replaceAll(' ', '');
    return !isNaN(str) && !isNaN(parseFloat(str))
};

function alphaToNum(alpha) {

    var i = 0,
        num = 0,
        len = alpha.length;

    for (; i < len; i++) {
      num = num * 26 + alpha.charCodeAt(i) - 0x40;
    }

    return num - 1;
};

function numToAlpha(num) {

    var alpha = '';

    for (; num >= 0; num = parseInt(num / 26, 10) - 1) {
        alpha = String.fromCharCode(num % 26 + 0x41) + alpha;
    }

    return alpha;
};

function _buildColumnsArray(range) {

    var i,
        res = [],
        rangeNum = range.split(':').map(function(val) {
            return alphaToNum(val.replace(/[0-9]/g, ''));
        }),
        start = rangeNum[0],
        end = rangeNum[1] + 1;

    for (i = start; i < end ; i++) {
        res.push(numToAlpha(i));
    }

    return res;
};

function replaceAgg(titleRow){
    var startRowID = Array.prototype.indexOf.call(titleRow.parentElement.children, titleRow)
    var endRowID = startRowID + titleRow.firstChild.rowSpan - 1;

    var titleRowClone = titleRow.cloneNode(true);
    var endRowClone = titleRow.parentElement.children[endRowID].cloneNode(true)

    endRowClone.insertBefore(titleRowClone.firstChild, endRowClone.firstChild);
    titleRow.parentElement.replaceChild(endRowClone, titleRow);
    titleRow.removeChild(titleRow.firstChild);
    endRowClone.parentNode.insertBefore(titleRow, endRowClone.nextSibling);
    titleRow.parentElement.removeChild(titleRow.parentElement.children[endRowID + 1])
};

function setColumnWidth(ws) {
    // get maximum character of each column
    for (var l of _buildColumnsArray(ws['!fullref'].replace(/[^A-Z:]/g, ''))){
        var col = []
        let regEx = new RegExp(`^${l}\\d+$`)
        for (key in ws){
            if (regEx.test(key) == true){
                col.push(ws[key].v.toString().length)
            }
        }

        if (col.length > 0){
            ws["!cols"].push(
                // when avg len of col < 2 set fixed width
                {wch: col.reduce((sum,a) => sum + a, 0)/col.length>2?Math.max(...col):10}
            )
        }
        else{
            // hide col if empty
            ws['!cols'].push({hidden: true})
        }

    }
};

function htmlTableToExcel(table, filename='HTMLTableToExcel'){
    let wb = XLSX.utils.book_new();
    // add content into empty cells
    let ws = XLSX.utils.table_to_sheet(table);

    // set cols style
    ws['!cols'] = []
    setColumnWidth(ws);

    // set rows style
    ws['!rows'] = []
    // Table header style
    // IMPORTANT styling thead elements before tbody "!rows" is array!
    Array.prototype.forEach.call(
        table.getElementsByTagName('THEAD')[0].getElementsByTagName('TR'),
        function(tr) {
            ws['!rows'].push({'hpt': 30});
        });

    // Set cells styles
    for (i in ws){
        if (typeof(ws[i])!= 'object') continue;
        let cell = XLSX.utils.decode_cell(i);
        ws[i].s = {
            font: {name: 'arial'},
            alignment: {
                vertical: 'center',
                wrapText: '1',
            },
        }
        // header
        if (cell.r < ws['!rows'].length){
            ws[i].s.fill = {
                patternType: 'solid',
                fgColor: {rgb: 'CCCCCC'},
                bgColor: {rgb: 'CCCCCC'}
            }
            ws[i].s.font.bold = true
            ws[i].s.border = {
                top: {style: 'thin', color: {rgb: 'E0E0E0'}},
                bottom: {style: 'thin', color: {rgb: 'E0E0E0'}},
                left: {style: 'thin', color: {rgb: 'E0E0E0'}},
                right: {style: 'thin', color: {rgb: 'E0E0E0'}},
            }
        }
        // replace spaces in digit
        if (typeof(ws[i].v) == 'string' && isNumeric(ws[i].v)){
            ws[i].v = ws[i].v.replaceAll(' ', '');
            ws[i].t = 'n'
            ws[i].s.alignment = {horizontal: 'right', vertical: 'center'}
        }
    };

    // Set tbody grouping
    let spanCounter = [];
    Array.prototype.forEach.call(
        table.getElementsByTagName('TBODY')[0].getElementsByTagName('TR'),
        function(tr) {
            spanCounter.sort();
            spanCounter.reverse();
            while (spanCounter.includes(0)) {
                spanCounter.pop()
            }
            if (tr.firstChild.rowSpan){
                spanCounter.push(tr.firstChild.rowSpan);
            }
            ws['!rows'].push({'hpt': 20, 'level': spanCounter.length});
            for (i=0; i<spanCounter.length; i++){
                --spanCounter[i];
            };
    });
    // row cells group bgcolor
    for (i in ws){
        if (typeof(ws[i])!= 'object' || String.prototype.startsWith(i.name, '!')) continue;
        let cell = XLSX.utils.decode_cell(i);
        if (cell.r > 0)
            if ('level' in ws['!rows'][cell.r]){
                if (ws['!rows'][cell.r].level == 3){
                    ws[i].s.fill = {
                        patternType: 'solid',
                        fgColor: {rgb: 'c1e1c6'},
                        bgColor: {rgb: 'c1e1c6'}
                    }
                }
                if (ws['!rows'][cell.r].level == 1){
                    ws[i].s.fill = {
                        patternType: 'solid',
                        fgColor: {rgb: '4f8654'},
                        bgColor: {rgb: '4f8654'}
                    }
                    ws[i].s.font = {
                        color: {rgb: 'FFFFFF'}, 
                        italic: true
                    }
                    ws[i].s.border = {
                        bottom: {style: 'thin', color: {rgb: 'E0E0E0'}},
                        top: {style: 'thin', color: {rgb: 'E0E0E0'}},
                    }
                }
                else {ws[i].s.font = {italic: true};}
        }
    }

    // Save file
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    return XLSX.write(wb, {bookType: 'xlsx' , bookSST: false, type: 'base64', cellStyles: true})
};

function saveTable2XLSX(table, headerRow){
    let tableCopy = table.cloneNode(true);
    let head = tableCopy.getElementsByTagName('THEAD')[0];
    let firstRow = tableCopy.getElementsByTagName('TR')[0];
    // Replace total row(s)

    Array.prototype.forEach.call(
        tableCopy.getElementsByClassName('pvtRowTotals'),
        function(row) {
            head.appendChild(row);
        }
    );

    //Add titleRow
    let titleRow = document.createElement('tr');
    let colCount = 0;

    for (elem in firstRow.children){
        if (typeof(firstRow.children[elem]) == 'object')
            colCount += firstRow.children[elem].colSpan;
    }

    titleRow.append(document.createElement('th'));
    titleRow.firstChild.colSpan = colCount/2;
    titleRow.firstChild.append(headerRow);

    titleRow.append(document.createElement('th'));
    titleRow.firstChild.nextSibling.colSpan = colCount - titleRow.firstChild.colSpan;
    titleRow.firstChild.nextSibling.textContent = '\u00A0'
    head.insertBefore(titleRow, firstRow);

    // Regroup table
    Array.prototype.forEach.call(tableCopy.getElementsByTagName('TBODY')[0].getElementsByTagName('TR'),
        function(tr) {
            if (tr.firstChild.rowSpan > 1){
                replaceAgg(tr);
            }
        }
    );
    $('th', tableCopy).each(function() {
        if ($(this, tableCopy).html() === ''){ $(this).append("\u00A0")}
    })

    return htmlTableToExcel(tableCopy, headerRow);
};
