function exportDataGrid(options) {
    if(options === undefined) return;

    let { customizeCell, component, worksheet, topLeftCell = { row: 1, column: 1 }, excelFilterEnabled } = options;

    worksheet.properties.outlineProperties = {
        summaryBelow: false,
        summaryRight: false
    };

    let result = {
        from: { row: topLeftCell.row, column: topLeftCell.column },
        to: { row: topLeftCell.row, column: topLeftCell.column }
    };

    let dataProvider = component.getDataProvider();

    return new Promise((resolve) => {
        dataProvider.ready().done(() => {
            let columns = dataProvider.getColumns();
            let headerRowCount = dataProvider.getHeaderRowCount();
            let dataRowsCount = dataProvider.getRowsCount();
            //
            for(let i=0;i<columns.length;i++){
                let coll=worksheet.getColumn(result.from.column+i);
                switch(i){
                  case 0:coll.width=5;
                    break;
                  case 1:coll.width=18;
                    break;
                  case 2:coll.width=15;
                    break;  
                  case 3:coll.width=25;
                    break;
                  case 4:coll.width=15;
                    break;
                  case 5:coll.width=15;
                    break;
                  case 6:coll.width=20;
                    break;
                }
            }
            //
            for(let rowIndex = 0; rowIndex < dataRowsCount; rowIndex++) {
                const row = worksheet.getRow(result.from.row + rowIndex);
                _exportRow(rowIndex, columns.length, row, result.from.column, dataProvider, customizeCell);

                if(rowIndex >= headerRowCount) {
                    row.outlineLevel = dataProvider.getGroupLevel(rowIndex);
                }
                if(rowIndex >= 1) {
                    result.to.row++;
                }
            }

            result.to.column += columns.length > 0 ? columns.length - 1 : 0;

            if(excelFilterEnabled === true) {
                if(dataRowsCount > 0) worksheet.autoFilter = result;
                worksheet.views = [{ state: 'frozen', ySplit: result.from.row + dataProvider.getFrozenArea().y - 1 }];
            }

            resolve(result);
        });
    });
}

function _exportRow(rowIndex, cellCount, row, startColumnIndex, dataProvider, customizeCell) {
    for(let cellIndex = 0; cellIndex < cellCount; cellIndex++) {
        const cellData = dataProvider.getCellData(rowIndex, cellIndex, true);
        const cell = row.getCell(startColumnIndex + cellIndex);
        //
        if(rowIndex<dataProvider.getHeaderRowCount()){
            cell.font=cell.font||{};
            cell.font.bold=true;
          } else{
            switch(cellData.cellSourceData.column.format){
              case 'currency': 
                cell.numFmt = '"$"#,##0';
                break;
              case 'decimal':
                cell.numFmt = '0';
                break;
              case 'billions':
                cell.numFmt = '#,##0,,,"B"';
                break;
              case 'exponential':
                cell.numFmt = '0.E+00';
                break;
              case 'fixedPoint':
                cell.numFmt = '#,##0';
                break;
              case 'millions':
                cell.numFmt = '#,##0,,"M"';
                break;
              case 'percent':
                cell.numFmt = '0%';
                break;
              case 'thousands':
                cell.numFmt = '#,##0,"K"';
                break;
              case 'day':
                cell.numFmt = 'd';
                break;
              case 'longDate':
                cell.numFmt = 'dddd, mmmm d, yyyy';
                break;
              case 'longTime':
                cell.numFmt = 'h:mm:ss AM/PM';
                break;
              case 'millisecond':
                cell.numFmt = 'ms';
                break;
              case 'month':
                cell.numFmt = 'mmmm';
                break;
              case 'monthAndDay':
                cell.numFmt = 'mmmm d';
                break;
              case 'monthAndYear':
                cell.numFmt = 'mmmm yyyy';
                break;
              case 'shortTime':
                cell.numFmt = 'h:mm AM/PM';
                break;
              case 'year':
                cell.numFmt = 'yyyy';
                break;
              case 'dayOfWeek':
                cell.numFmt = 'dddd';
                break;
              case 'hour':
                cell.numFmt = 'hh';
                break;
              case 'longDateLongTime':
                cell.numFmt = 'dddd, mmmm d, yyyy, hh:mm:ss AM/PM';
                break;
              /*case 'minute':
                cell.numFmt = 'm';
                break;*/
              case 'second':
                cell.numFmt = 'ss';
                break;
              case 'shortDateShortTime':
                cell.numFmt = 'd/m/yyyy, h:mm AM/PM';
                break;
            }
          }
          if(cellData.cellSourceData.column.dataType==='date' && rowIndex>=dataProvider.getHeaderRowCount()){
            cellData.value=new Date(Date.UTC(cellData.value.getFullYear(),cellData.value.getMonth(),cellData.value.getDate(),cellData.value.getHours(),cellData.value.getMinutes(),cellData.value.getSeconds()));
          }
        //
        cell.value = cellData.value;

        if(customizeCell !== undefined) {
            customizeCell({
                cell: cell,
                gridCell: {
                    column: cellData.cellSourceData.column,
                    rowType: cellData.cellSourceData.rowType,
                    data: cellData.cellSourceData.data,
                    value: cellData.value,
                    groupIndex: cellData.cellSourceData.groupIndex
                }
            });
        }
    }
}
