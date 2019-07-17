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

            for(let rowIndex = 0; rowIndex < dataRowsCount; rowIndex++) {
                const row = worksheet.getRow(result.from.row + rowIndex);
                _exportRow(rowIndex, columns.length, row, result.from.column, dataProvider, customizeCell);

                //
                if(rowIndex<headerRowCount){
                    for(let cellIndex = 0; cellIndex < columns.length; cellIndex++) {
                      if(dataProvider._options.columns[rowIndex][cellIndex].colspan>1||dataProvider._options.columns[rowIndex][cellIndex].rowspan>1){
                        worksheet.mergeCells(rowIndex+result.from.row,cellIndex+result.from.column,rowIndex+result.from.row+dataProvider._options.columns[rowIndex][cellIndex].rowspan-1,cellIndex+dataProvider._options.columns[rowIndex][cellIndex].colspan+result.from.column-1);
                      }
                    }
                }
                //
                if(rowIndex >= headerRowCount) {
                    row.outlineLevel = dataProvider.getGroupLevel(rowIndex);
                }
                if(rowIndex >= 1) {
                    result.to.row++;
                }
            }//Можно использовать этот фрагмент, но я его закинул в основной цикл
            /*for(let i=0; i<headerRowCount; i++){
              for(let j=0; j<columns.length;j++){
                if(dataProvider._options.columns[i][j].colspan>1||dataProvider._options.columns[i][j].rowspan>1){
                  worksheet.mergeCells(i+result.from.row,j+result.from.column,i+result.from.row+dataProvider._options.columns[i][j].rowspan-1,j+dataProvider._options.columns[i][j].colspan+result.from.column-1);
                }
              }
            }*/

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

        //Если не поставить условие, то некоторые ячейки будут пустыми.
        if(cellData.value!=='') cell.value = cellData.value;//Чтобы работало без условия, нужно использовать закомментированный фрагмент

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
