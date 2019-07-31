var targetID = '1buA9zuTTPI0iBd13YtY0kRSuSsF5UMKk0_9s_DkBkCs';
var targetSheetName = 'Sheet2';

function onEdit(e) {
  
   // Get the event object properties
  var range = e.range;
  var changedValue = e.value;
  Logger.log(changedValue);
  
  var spreadsheet = SpreadsheetApp.openById('1k9hVZU41Wn_7fvjZ-e6CQIC5BTSR7kiDJKPyPfbrMCQ');
  var sheet = spreadsheet.getSheetByName('Sheet2');
  
  //Output values in destinated worksheets into a two dimensional arrays:
  var K = mergeAgencyCells(sheet);
  Logger.log('k %s', K);
  /*
  var T = morph3D(K);
  var array = T
    .map(function (el) {
        return [el];
    });

  var sheet1 = spreadsheet.getSheetByName('Sheet3');
  Logger.log(sheet1.getRange(1,1, array[0].length, array.length).getValues());
  sheet1.getRange(1,1, array[0].length, array.length).setValues(T);
  */
  //Get the cell position:
  var row = range.getRowIndex();
  var column = range.getColumnIndex();
  
  //Make adjustment to exclude the headers:
  var eventType = K[row-4][1];
  var dateVal = K[row-4][2];
  var pstVal = K[row-4][3];
  var agency = K[row-4][9];

  //Set relevant cells:
  exportValue(changedValue, eventType, dateVal, pstVal, agency);

}

//Main functions to collect values from agencies to main sheets:
function exportValue(changedValue, eventType, dateVal, pstVal, agency){
  var ss = SpreadsheetApp.openById(targetID);
  var s = ss.getSheetByName(targetSheetName); 
  var outputArray = mergeCells(s);
  Logger.log(outputArray);
  var outRow = findRow(changedValue, eventType, dateVal, pstVal, agency, outputArray);
  Logger.log(outRow);
  var target = s.getRange(outRow, 11);
  target.setValue(changedValue);
}

//Find Row where two 2-D array's value match & return the destinated rows for the values to be set:
function findRow(changedValue, eventType, dateVal, pstVal, agency, outputArray){
  
  var dateVal = dateVal.toString();
  var pstVal = pstVal.toString();
  var outRow; 
  
  for (var i = 0; i < outputArray.length; i++)
  { 
  
    if (outputArray[i][1] == eventType && outputArray[i][2] == dateVal && outputArray[i][3] == pstVal && outputArray[i][9] == agency)
    {
      Logger.log(outputArray[i][1]);
      Logger.log(outputArray[i][2]);
      Logger.log('idk %s', outputArray[i][9]);
      Logger.log(i);
      outRow = i + 4;
      break;
    }
  }
  return outRow;
}

//Main function to make agencies' worksheets into two dimensional arrays:
function mergeAgencyCells(sheet){  

  var range = sheet.getRange('A4:Q');
  var mergeArr = [];
  var splitArr = [];
  var size = 4;
  
  //If cells are part the merge:
  var A = [];
  var B = [];
  if(range.isPartOfMerge()){
    range.getMergedRanges().forEach(function(range){
      A = [];
      Col = range.getColumn();
      startRow = range.getRowIndex();
      lastRow = range.getLastRow();
      val = range.getCell(1,1).getValues();
      for (i = 0; i < range.getValues().length; i++){
        A.push(val[0][0]);
      }
      mergeArr.push(Col, startRow, lastRow, A);
    });
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
  
    //Logging the range:
  var C = [];
  var D = [];
  var E = [];

  for (l = 0; l < splitArr.length; l++){
    E = [];
    if(splitArr[l][0] == 1){
      C.push(splitArr[l][1]);
      D.push(splitArr[l][2]);
    }
    E.push(C,D);
  }
  
  //Advanced Array Operations:

  var X = [[]];
  Logger.log(E[0][0]);
  X[0] = [];
  X[1] = [];
  X[0].push(E[0][0]);
  var eLen = E[0].length;
  var lastVal = E[1][eLen-1] + 1 ;
  for(d = 1; d < E[0].length; d++){
    if(E[0][d] - E[0][d-1] != 2){
      X[0].push(E[0][d]);
    }
    if(E[1][d] - E[1][d-1] != 2){
      X[1].push(E[1][d-1]);
    }
  }
  X[1].push(lastVal);
  Logger.log('time to test if im right %s', X);
  
  var value = [];
  var F = [];
  var transpose = [];
  var colHeader;
  var J = [];
  
  for (l = 0; l < X[0].length; l++){
    var middle = 3;
    var start = X[0][l];
    var end = X[1][l];
    colHeader  = end + 1;

    //For the first Range:
    if(l == 0){
      value[l] = sheet.getRange(start,10, end-middle, sheet.getLastColumn()).getValues();
      transpose[l] = transposeArray(value[l]);
  
      for(j = 0; j < splitArr.length; j++){
        if(splitArr[j][1] >= start && splitArr[j][2] <= end){
          F.push(splitArr[j]);
        }
      }
      F.sort(function (element_a, element_b) {
        return element_a[0] - element_b[0];
      });
      var H = returnArray(F,transpose[l]);
      var I = sheet.getRange(colHeader,1,1,26).getValues();
      H.forEach(function(Val){
        J.push(Val);
          });
      I.forEach(function(Val){
        J.push(Val);
          });
      // Reset the values:
      transpose = [];
      F = [];
      H = [];
      value = [];
      I = [];
    }
  
    else{
      value[l] = sheet.getRange(start,10, end-start+1, sheet.getLastColumn()).getValues();
      transpose[l] = transposeArray(value[l]);
      Logger.log(transpose[l]);
      for(j = 0; j < splitArr.length; j++){
        if(splitArr[j][1] >= start && splitArr[j][2] <= end){
          F.push(splitArr[j]);
        }
      }

      F.sort(function (element_a, element_b) {
        return element_a[0] - element_b[0];
      });
      var H = returnArray(F,transpose[l]);
      var I = sheet.getRange(colHeader,1,1,26).getValues();

      H.forEach(function(Val){
              J.push(Val);
      });
      I.forEach(function(Val){
        J.push(Val);
      });
      //Reset the values:
      F = [];
      H = [];
      I = [];
    }
  }
  return J;
}


// Turn main sheet's values into two-dimensional arrays:
function mergeCells(sheet){  
  var range = sheet.getRange('A4:Q');
  var mergeArr = [];
  var splitArr = [];
  var size = 4;
  
  //If cells are part the merge:
  var A = [];
  var B = [];
  if(range.isPartOfMerge()){
    range.getMergedRanges().forEach(function(range){
      A = [];
      Col = range.getColumn();
      startRow = range.getRowIndex();
      lastRow = range.getLastRow();
      val = range.getCell(1,1).getValues();
      for (i = 0; i < range.getValues().length; i++){
        A.push(val[0][0]);
      }
      mergeArr.push(Col, startRow, lastRow, A);
    });
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
  
    //Logging the range:
  var C = [];
  var D = [];
  var E = [];

  for (l = 0; l < splitArr.length; l++){
    E = [];
    if(splitArr[l][0] == 1){
      C.push(splitArr[l][1]);
      D.push(splitArr[l][2]);
    }
    E.push(C,D);
  }
  
  var value = [];
  var F = [];
  var transpose = [];
  var colHeader;
  var J = [];
  
  for (l = 0; l < E[0].length; l++){
    var middle = 3;
    var start = E[0][l];
    var end = E[1][l];
    colHeader  = end + 1;

    //For the first Range:
    if(l == 0){
      value[l] = sheet.getRange(start,10, end-middle, sheet.getLastColumn()).getValues();
      Logger.log('vale %s',value[l]);
      transpose[l] = transposeArray(value[l]);
  
      for(j = 0; j < splitArr.length; j++){
        if(splitArr[j][1] >= start && splitArr[j][2] <= end){
          F.push(splitArr[j]);
        }
      }
      F.sort(function (element_a, element_b) {
        return element_a[0] - element_b[0];
      });
      var H = returnArray(F,transpose[l]);
      var I = sheet.getRange(colHeader,1,1,26).getValues();
      H.forEach(function(Val){
        J.push(Val);
          });
      I.forEach(function(Val){
        J.push(Val);
          });
     
      transpose = [];
      F = [];
      H = [];
      value = [];
      I = [];
    }
    //First merged ranges below:
    else{
      value[l] = sheet.getRange(start,10, end-start+1, sheet.getLastColumn()).getValues();
      transpose[l] = transposeArray(value[l]);
      for(j = 0; j < splitArr.length; j++){
        if(splitArr[j][1] >= start && splitArr[j][2] <= end){
          F.push(splitArr[j]);
        }
      }

      F.sort(function (element_a, element_b) {
        return element_a[0] - element_b[0];
      });
      var H = returnArray(F,transpose[l]);
      var I = sheet.getRange(colHeader,1,1,26).getValues();

      H.forEach(function(Val){
              J.push(Val);
      });
      I.forEach(function(Val){
        J.push(Val);
      });
      
      F = [];
      H = [];
      I = [];
    }
  }
  return J;
}

//Return values:
function returnArray(F, transpose){
  var colIndex = 0;
  var uniqueVal = [];
  
  var arrayGroupBy = function(F,colIndex){
    F.forEach(function(row){
      if(uniqueVal.indexOf(row[colIndex]) === -1){
        uniqueVal.push(row[colIndex]);
      }
    });

    var uniqueRow = [];
    uniqueVal.forEach(function(value){
      //Get filter row:
      var filterRows = F.filter(function(row){
        return row[colIndex] === value;
      });
      var row = [];
      for(p = 0; p < filterRows.length; p++){
        for(v = 0 ; v < filterRows[p][3].length; v++){
          row.push(filterRows[p][3][v]);
        }
      }
      uniqueRow.push(row);
    });
    return uniqueRow;
  };
  var G = arrayGroupBy(F,colIndex);
  for (r = 0 ; r < transpose.length; r++){
    G.push(transpose[r]);
  }
  var H = transposeArray(G);
  return H;
}

//Transpose horizontal arrays vertically:
function transposeArray(array){
  var result = [];
  for (var col = 0; col < array[0].length; col++) { // Loop over array cols
    result[col] = [];
    for (var row = 0; row < array.length; row++) { // Loop over array rows
      result[col][row] = array[row][col]; // Rotate
    }
  }
  return result;
}

//Morph 2D arrays into 3D to set multidimensional values:
function morph3D(arr){
  var ar = [];
  var c = [];
  for(var i = 0; i < arr.length; i++){
    c = [];
    for(var j = 0; j < arr[0].length; j++){
      c.push(arr[i][j]);
    }
    ar.push(c);
  }
  return ar;
}