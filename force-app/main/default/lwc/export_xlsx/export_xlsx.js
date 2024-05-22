export function exportXLSX(config, fileName){
    let uri = 'data:application/vnd.ms-excel;base64;charset=utf-8,';

    let tmplWorkbookXML = `
        <?xml version="1.0" encoding="windows-1252"?>
        <?mso-application progid="Excel.Sheet"?>
        <Workbook 
            xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns:html="http://www.w3.org/TR/REC-html40"
        >
            <Styles>
                {styles}
            </Styles>
            {worksheets}
        </Workbook>
    `;

    let tmplWorksheetXML = `
        <Worksheet ss:Name="{nameWS}">
            <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
                {displayGrid} {zoom}
            </WorksheetOptions>
            <Table>
                {rows}
            </Table>
        </Worksheet>
    `;

    let tmplRowXML = `
        <Row {height}>
    `;

    let tmplCellXML = `
        <Cell {attributeStyleID} {colspan} {rowspan}>
            <Data ss:Type="{nameType}">{data}</Data>
        </Cell>
    `;

    let styles = '';

    config.forEach(table =>{
        if(!table.table){
            table.table = '';
        }

        if(table.zoom === undefined){
            table.zoom = 100;
        }
        if(table.displayGrid === undefined){
            table.displayGrid = true;
        }

        if(table.style){
            styles += table.style;
        }
    });

    if(fileName && !fileName.endsWith('.xls')){
        fileName += '.xls';
    }else{
        fileName = 'File.xls';
    }

    let base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) };
    let format = function(s, c) { return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) };

    let workbookXML = "";
    let worksheetsXML = "";
    let rowsXML = "";

    for(let i = 0; i < config.length; i++){
        for(let j = 0; j < config[i].table.rows.length; j++){
            // row config
            let height = config[i].table.rows[j].getAttribute("data-xls-height");

            let rowCTX = {
                height: height ? `ss:Height="${height}"` : ''
            }

            rowsXML += format(tmplRowXML, rowCTX);

            for(let k = 0; k < config[i].table.rows[j].cells.length; k++){
                // cell config
                let dataType = config[i].table.rows[j].cells[k].getAttribute("data-xls-type");
                let dataStyle = config[i].table.rows[j].cells[k].getAttribute("data-xls-style");
                let dataValue = config[i].table.rows[j].cells[k].getAttribute("data-xls-value");
                let colspan = config[i].table.rows[j].cells[k].getAttribute("colspan");
                let rowspan = config[i].table.rows[j].cells[k].getAttribute("rowspan");

                dataValue = dataValue ? dataValue : config[i].table.rows[j].cells[k].innerHTML;

                if(!dataType){
                    dataType = 'String';
                }else if(!isNaN(dataValue)){
                    dataType = 'Number';
                    dataValue = parseFloat(dataValue);
                }

                let cellCTX = {
                    attributeStyleID: dataStyle ? `ss:StyleID="${dataStyle}"` : '',
                    nameType: dataType,
                    data: dataValue,
                    colspan: colspan ? `ss:MergeAcross="${colspan - 1}"` : '',
                    rowspan: rowspan ? `ss:MergeDown="${rowspan - 1}"` : '',
                };

                rowsXML += format(tmplCellXML, cellCTX);
            }

            rowsXML += '</Row>';
        }

        // workbook config
        let workbookCTX = {
            rows: rowsXML, 
            nameWS: config[i].tabName || 'Sheet ' + i,
            displayGrid: config[i].displayGrid ? '' : '<DoNotDisplayGridlines/>',
            zoom: config[i].zoom ? `<Zoom>${config[i].zoom}</Zoom>` : '',
        };

        worksheetsXML += format(tmplWorksheetXML, workbookCTX);

        rowsXML = "";
    }

    // file config
    let fileCTX = {
        worksheets: worksheetsXML,
        styles: styles,
    };

    workbookXML = format(tmplWorkbookXML, fileCTX);

    // download file
    let link = document.createElement('a');
    link.href = uri + base64('\uFEFF' + workbookXML);
    link.download = fileName;
    link.target = '_blank';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}