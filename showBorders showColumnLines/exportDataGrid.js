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
            for(let i=0;i<columns.length;i++){
              let coll=worksheet.getColumn(result.from.column+i);
              if(dataProvider._exportController.component._options.wordWrapEnabled===true) coll.alignment={wrapText:true};
              switch(i){
                case 0:coll.width=10;
                  break;
                case 1:coll.width=30;
                  break;
                case 2:coll.width=30;
                  break;  
                case 3:coll.width=25;
                  break;
              }
            }
            for(let rowIndex = 0; rowIndex < dataRowsCount; rowIndex++) {
                const row = worksheet.getRow(result.from.row + rowIndex);
                _exportRow(rowIndex, columns.length, row, result.from.column, dataProvider, customizeCell);

                //
                let cell;
                for(let cellIndex = 0; cellIndex < columns.length; cellIndex++) {
                    cell = worksheet.getCell(result.from.row + rowIndex,result.from.column + cellIndex);//3
                    if(rowIndex<headerRowCount){
                      if(dataProvider._options.columns[rowIndex][cellIndex].colspan>1||dataProvider._options.columns[rowIndex][cellIndex].rowspan>1){
                        worksheet.mergeCells(rowIndex+result.from.row,cellIndex+result.from.column,rowIndex+result.from.row+dataProvider._options.columns[rowIndex][cellIndex].rowspan-1,cellIndex+dataProvider._options.columns[rowIndex][cellIndex].colspan+result.from.column-1);
                      }//1
                      if(dataProvider._exportController.component._options.showColumnLines===false){
                        cell.border=cell.border={top: {style:'thin', color: {argb:'FF666666'}}, bottom:{style:'thin', color:{argb:'FF666666'}}};
                        if(dataProvider._exportController.component._options.showBorders===true){
                          if(cellIndex===0){
                            cell.border={
                              top: {style:'thin', color: {argb:'FF666666'}},
                              bottom: {style:'thin', color: {argb:'FF666666'}},
                              left: {style:'thin', color: {argb:'FF666666'}}
                            };
                          }
                          if(cellIndex===columns.length-1){
                            cell.border = {
                              top: {style:'thin', color: {argb:'FF666666'}},
                              bottom: {style:'thin', color: {argb:'FF666666'}},
                              right: {style:'thin', color: {argb:'FF666666'}}
                            };
                          }
                        }
                      } else {
                        cell.border = {
                          top: {style:'thin', color: {argb:'FF666666'}},
                          bottom: {style:'thin', color: {argb:'FF666666'}},
                          right: {style:'thin', color: {argb:'FF666666'}},
                          left: {style:'thin', color: {argb:'FF666666'}}
                        }
                        if(dataProvider._exportController.component._options.showBorders===false){
                          if(cellIndex===0){
                            cell.border={
                              top: {style:'thin', color: {argb:'FF666666'}},
                              bottom: {style:'thin', color: {argb:'FF666666'}},
                              right: {style:'thin', color: {argb:'FF666666'}}
                            };
                          }
                          if(cellIndex===columns.length-1){
                            cell.border = {
                              top: {style:'thin', color: {argb:'FF666666'}},
                              bottom: {style:'thin', color: {argb:'FF666666'}},
                              left: {style:'thin', color: {argb:'FF666666'}}
                            };
                          }
                        }
                      }
                      //1
                    } else {
                      if(dataProvider._exportController.component._options.showColumnLines===false && dataProvider._exportController.component._options.showBorders===true){
                        if(cellIndex===0){
                          if(rowIndex===0) cell.border = {top: {style:'thin', color: {argb:'FF666666'}}, left:{style:'thin', color:{argb:'FF666666'}}};
                          else if(rowIndex===dataProvider.getRowsCount()-1) cell.border = {bottom: {style:'thin', color: {argb:'FF666666'}}, left:{style:'thin', color:{argb:'FF666666'}}};
                          else cell.border = {left: {style:'thin', color: {argb:'FF666666'}}};
                        }
                        if(cellIndex===columns.length-1){
                          if(rowIndex===0) cell.border = {top: {style:'thin', color: {argb:'FF666666'}}, right:{style:'thin', color:{argb:'FF666666'}}};
                          else if(rowIndex===dataProvider.getRowsCount()-1) cell.border = {bottom: {style:'thin', color: {argb:'FF666666'}}, right:{style:'thin', color:{argb:'FF666666'}}};
                          else cell.border = {right: {style:'thin', color: {argb:'FF666666'}}};
                        }
                        if(cellIndex>0 && cellIndex<columns.length-1 && rowIndex===headerRowCount) cell.border = {top: {style:'thin', color: {argb:'FF666666'}}};
                        if(cellIndex>0 && cellIndex<columns.length-1 && rowIndex===dataProvider.getRowsCount()-1) cell.border = {bottom: {style:'thin', color: {argb:'FF666666'}}};
                      }
                      if(dataProvider._exportController.component._options.showColumnLines===false && dataProvider._exportController.component._options.showBorders===false && dataProvider._exportController.component._options.showColumnHeaders===false && rowIndex===headerRowCount) cell.border = {top: {style:'thin', color: {argb:'FF666666'}}};
                      if(dataProvider._exportController.component._options.showColumnLines===true){
                        if(dataProvider._exportController.component._options.showBorders===false && cellIndex>0 && cellIndex<columns.length-1) cell.border = {left: {style:'thin', color: {argb:'FF666666'}}, right: {style:'thin', color: {argb:'FF666666'}}};
                        if(dataProvider._exportController.component._options.showBorders===true){
                          if(rowIndex>headerRowCount && rowIndex<dataProvider.getRowsCount()-1) cell.border = {left: {style:'thin', color: {argb:'FF666666'}}, right: {style:'thin', color: {argb:'FF666666'}}};
                          else{
                            if(rowIndex===headerRowCount) cell.border = {left: {style:'thin', color: {argb:'FF666666'}}, right: {style:'thin', color: {argb:'FF666666'}}, top: {style:'thin', color: {argb:'FF666666'}}};
                            if(rowIndex===dataProvider.getRowsCount()-1) cell.border = {left: {style:'thin', color: {argb:'FF666666'}}, right: {style:'thin', color: {argb:'FF666666'}}, bottom: {style:'thin', color: {argb:'FF666666'}}};
                          }
                        }
                      }
                    }
                    //2
                  }
                //
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
        if(rowIndex<dataProvider.getHeaderRowCount()){
            cell.font=cell.font||{};
            cell.font.bold=true;
          }
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
