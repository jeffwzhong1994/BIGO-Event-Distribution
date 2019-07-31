var targetID = '1buA9zuTTPI0iBd13YtY0kRSuSsF5UMKk0_9s_DkBkCs';
var targetSheetName = 'Sheet1';

function onEdit(e) {
  
   // Get the event object properties
  var range = e.range;
  var value = e.value;
  
  //Get the cell position 
  var row = range.getRowIndex();
  var column = range.getColumnIndex();
  
  //Get Date Range and Date values:
  var dateRange = SpreadsheetApp.getActiveSheet().getRange(row, 3)
  var dateColumn = dateRange.getColumnIndex();
  var dateValue = dateRange.getValues();
  
  //Get PST Range and PST values:
  var pstRange = SpreadsheetApp.getActiveSheet().getRange(row, 4)
  var pstColumn = pstRange.getColumnIndex();
  var pstValue = pstRange.getValues();  

  //Main function:
  exportValue(row,dateColumn,dateValue, pstColumn, pstValue, value)
}

function exportValue(row,dateColumn,dateValue, pstColumn, pstValue, value) {
  //Open main sheets:
  var ss = SpreadsheetApp.openById(targetID);
  var s = ss.getSheetByName(targetSheetName); 
  //Start from the second rows:
  var dataRange = s.getRange("A2:N");
  var values = mergeCells(dataRange,s);
  var outRow = findRow(values, dateValue, pstValue);
  Logger.log(outRow);
  var target = s.getRange(outRow, 10);
  target.setValue(value);
}

//Find the row in the main sheet & set values:
function findRow(values, dateValue, pstValue){
  var dateValue = dateValue.toString();
  var pstValue = pstValue.toString();
  var outRow; 
  
  for (var i = 0; i < values.length; i++)
  { 
    if (values[i][2] == dateValue && values[i][3] == pstValue)
    {
      outRow = i + 3;
      break;
    }
  }
  return outRow;
}

//Function dealing with merge cells:
function mergeCells(range, sheet){
  var values = range.getValues();
  var mergeArr = [];
  var splitArr = [];
  var size = 4;
  
  //If cells are part the merge:
  if(range.isPartOfMerge()){
    for(var k = 0; k < range.getMergedRanges().length; k++){
      //Merge Values:
      Col = range.getMergedRanges()[k].getColumn();
      startRow = range.getMergedRanges()[k].getRowIndex();
      lastRow = range.getMergedRanges()[k].getLastRow();
      val = range.getMergedRanges()[k].getCell(1,1).getValues();
      mergeArr.push(Col,startRow,lastRow,val[0][0]);
  }
  }
  
  while(mergeArr.length > 0)
     splitArr.push(mergeArr.splice(0,size));
  
  //Sort 2-D Arrays:
  splitArr.sort(function (element_a, element_b) {
    return element_a[0] - element_b[0];
  });
   splitArr.sort(function (element_a, element_b) {
    return element_a[1] - element_b[1];
  });
  
  //Push Relative Unique start/End row into the list：
  var columnIndex = 1;
  var uniqueArrayStart = function(splitArr, columnIndex){
    var uniqueCol = [];
    splitArr.forEach(function(row){
      if(uniqueCol.indexOf(row[columnIndex]) === -1){
        uniqueCol.push(row[columnIndex]);
      }
    });
    return uniqueCol;
  }
  var start = uniqueArrayStart(splitArr, columnIndex);
  
  var columnIndex = 2;
  var uniqueArrayEnd = function(splitArr, columnIndex){
    var uniqueCol = [];
    splitArr.forEach(function(row){
      if(uniqueCol.indexOf(row[columnIndex]) === -1){
        uniqueCol.push(row[columnIndex]);
      }
    });
    return uniqueCol;
  }
  var end = uniqueArrayEnd(splitArr, columnIndex);
  var startEndArray = [];
  startEndArray.push(start);
  startEndArray.push(end);

    //Column to Groupby:
  var colIndex = 1;
  var colDist = [0,3];
  var arrayGroupBy = function(splitArr,colIndex,colDist){
    var uniqueVal = [];
    splitArr.forEach(function(row){
      if(uniqueVal.indexOf(row[colIndex]) === -1) {
         uniqueVal.push(row[colIndex]);
    }
    });
    var uniqueRow = [];
    uniqueVal.forEach(function(value){     
      //Get Filter Rows:
      var filterRows = splitArr.filter(function(row){
        return row[colIndex] === value;
      });     
      var row = [];
      //Push Row:
      colDist.forEach(function(num){
        row.push(filterRows[0][num]);
        row.push(filterRows[1][num]);
        row.push(filterRows[2][num]);
      });
    uniqueRow.push(row);
    });
    return uniqueRow;
  }
  
  var secondaryArr = arrayGroupBy(splitArr,colIndex,colDist);
  //Third Time Operations:
  var arrary = []
  var arrary1 = [];
  var arrary2 = [];
  var arrary3 = [];
  var terminalArrary = [];
  for(k = 0; k < secondaryArr.length; k++){
    for(j = 0; j < secondaryArr[k].length; j++){
      if(j%3 == 0){
        arrary1.push(secondaryArr[k][j]);
      }
      else{
        if(j%3 == 1){
          arrary2.push(secondaryArr[k][j]);
        }
          else{
            if(j%3 == 2){
              arrary3.push(secondaryArr[k][j]);
            }
          }    
      }
    }
    arrary.push(arrary1, arrary2, arrary3);
    terminalArrary.push(arrary);
    arrary= [];
    arrary2 = [];
    arrary3 = [];
    arrary1 = [];
  }

  //Obtain the values before merging with other cells:
  var setArray = [];
  var setArray2 = [];
  for(l = 0; l < terminalArrary.length; l++){
    for(j = 0; j < terminalArrary[l].length; j++){
      setArray.push(terminalArrary[l][j][1]);
    }
    setArray2.push(setArray);
    setArray= [];
  }
  
  //Loop through Ranges & Merge Cells:
  var arr1 = [];
  var val = [];
  var medium = [];
  var rng = [];
  var middle = 2;
  for(i = 0; i < startEndArray[0].length; i++){
    var start = startEndArray[0][i];
    var end = startEndArray[1][i];
    if(i == 0){
      val[i] = sheet.getRange(start,4,end-middle,sheet.getLastColumn()).getValues();
      for(m = 0; m < val[i].length; m++){
        medium = setArray2[i].slice(0);
        for(n = 0; n < val[i][m].length; n++){
          medium.push(val[i][m][n]);
        }
        arr1.push(medium);
        medium = [];
      }
    }
    else{
      val[i] = sheet.getRange(start,4,end-start+1,sheet.getLastColumn()).getValues();
      for(p = 0; p < val[i].length; p++){
        medium = setArray2[i].slice(0);
        for(q = 0; q < val[i][p].length; q++){
          medium.push(val[i][p][q]);
        }
        arr1.push(medium);
        medium = [];
      }  
    }
  }
  //Return the array back to the main function:
  return arr1;
}
