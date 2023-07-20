// require html2pdf

function addSpacers(target){
    // NOTE a4 fromat in portrait(width=height/Math.sqrt(2)) and in landscape(width=height*Math.sqrt(2))
    var pageH = document.body.clientWidth/ Math.sqrt(2);
    var spacer = document.createElement('div');
    for (const row of target.children){
        if ($(row).css('display') !== 'none'){
            var rect = row.getBoundingClientRect();
            // IMPORTANT first page always is 1.
            var startBlockPage = Math.ceil(rect.top/pageH) == 0 ? 1 : Math.ceil(rect.top/pageH);
            var endBlockPage = Math.ceil((rect.bottom)/pageH);

            // block bigger than page need recurse for slove it problem
            // TODO definition for tables
            // INFO on large table with using colSpans n rowSpans sometimes rowSpan'll be ignored.
            if (rect.bottom - rect.top >= pageH){
                addSpacers(row)
            }
            else{
                // all OK
                if(startBlockPage == endBlockPage){}
                // add spacer
                else{
                    spacer.style.width = '100%';
                    spacer.style.backgroundColor = 'transparent';
                    spacer.class = 'html2pdf__spacer';

                    // pageH * (endBlockPage -1)
                    spacer.style.height = (pageH * startBlockPage) - rect.top + 16 + 'px';
                    row.parentNode.insertBefore(spacer.cloneNode(true), row);
                }
            }
        }
    }
};

function stylingPage(){
    // TODO need reset scrollable elements and change overflow and maxHeight for that element and his content.
    // INFO https://github.com/eKoopmans/html2pdf.js/issues/650

    // TODO callback reset changes.

    // set fixed page width
    document.body.style.width = '1833px';

    // set overflow visible
    document.getElementsByClassName('dashboard')[0].style.overflowX = 'visible';

    // styling header with !important static styles
    $('div:has(> .dashboard-chart-id-2)')[0].style.cssText+= 'width: 1801px !important; max-width: 1801px !important; margin-left: 0 !important;'

    // fix bug with the hidding text of styled span objects wich parent has background attribute
    // INFO https://github.com/niklasvh/html2canvas/issues/1497
    $('.css-1wale1b-getColumnConfigs > span').css('position', 'relative').each(
        function() {
            this.removeAttribute('data-column-name');
        }
    );

    // add spacers to ignore broken pagebreak
    // https://github.com/eKoopmans/html2pdf.js/issues/200
    addSpacers($('.grid-content')[0]);
};

function getPdf(){
    // NOTE prepared for calling into selenium.
    stylingPage()
    var opt = {
        margin:       0,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, width: 1833, backgroundColor: '#cccccc'},
        jsPDF:        { unit: 'pt', format: 'a4', orientation: 'l' }
    };

    // NOTE returns binary data!
    return html2pdf().set(opt).from(document.body).outputPdf().then(function(pdf){ return btoa(pdf);})
};
