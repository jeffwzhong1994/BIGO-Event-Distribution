var agentID1 = '1k9hVZU41Wn_7fvjZ-e6CQIC5BTSR7kiDJKPyPfbrMCQ';
var agentSheetName1 = 'Sheet1';
var agentID2 = '1lonWd4w4SZ4MD75KEvdtLnydyrtzI9OL5-6L-qAnmpI';
var agentSheetName2 = 'Sheet1';
var agentID3 = '1Vcr3R_iiR60JRrNV1mavNdOf7jFr8AEy9cR5IxZbidM';
var agentSheetName3 = 'Sheet1';

function onEdit(e) {
  
   // Get the event object properties
  var range = e.range;
  var value = e.value;
  //Get the cell position 
  var row = range.getRowIndex();
  var column = range.getColumnIndex();

}

function liveHouse(){
  //Read the First two lines of header:
  copyHeader();

  //Read sheet from the first merge cells:
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(3,1,sheet.getLastRow(),sheet.getLastColumn());
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
  Logger.log(splitArr);
  
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

    //Sort the Array:
  for (k = 0; k < startEndArray.length; k++){
    startEndArray[k].sort(function(a,b){return a-b;});
  }

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
  Logger.log(secondaryArr);
  
  //Intermediate list operations to make it separate by every three cells:
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
  Logger.log(terminalArrary);
  
  //Obtain the values before merging with other cells:
  var setArray = [];
  var setArray2 = [];
  for(l = 0; l < terminalArrary.length; l++){
    for(j = 0; j < terminalArrary[l].length; j++){
      setArray.push([terminalArrary[l][j][1]]);
    }
    setArray2.push(setArray);
    setArray= [];
  }
  Logger.log(setArray2);
  //Range for the final output:
  var arr1 = [];
  var arr2 = [];
  var value = [];
  var medium = [];
  var range = [];

  //Looping through merge cell ranges:
  for(i = 0; i < startEndArray[0].length; i++){
    var start = startEndArray[0][i];
    Logger.log(start);
    var end = startEndArray[1][i];
    var middle = 2;

    //Set Range for colors & stuffs:
    range[i] = sheet.getRange(start,1, 1, 17);
    var fontColors = range[i].getFontColors();
    var backgrounds = range[i].getBackgrounds();
    var fonts = range[i].getFontFamilies();
    var fontWeights = range[i].getFontWeights();
    var fontStyles = range[i].getFontStyles();
    
    //For the first Range:
    if(i == 0){
      value[i] = sheet.getRange(start,4,end-middle,sheet.getLastColumn()).getValues();
      medium = setArray2[i].slice(0);
      for(j = 0; j < value[i].length; j++){
        for(k = 0; k < value[i][j].length; k++){  
          arr1.push([value[i][j][k]]);
        }
        arr2.push(arr1);
        arr1 = [];
      }
      finalOutput(arr2, medium, fontColors, backgrounds, fonts, fontWeights,fontStyles);
    }
    
    //For the first Range Below:
    else {
      Logger.log(arr2);
      arr2 = [];
      medium = [];
 
      value[i] = sheet.getRange(start,4,end-start+1,sheet.getLastColumn()).getValues();
      medium = setArray2[i].slice(0);
      for(j = 0; j < value[i].length; j++){
        for(k = 0; k < value[i][j].length; k++){  
          arr1.push([value[i][j][k]]);
        }
        arr2.push(arr1);
        arr1 = [];
      }
      finalOutput(arr2, medium, fontColors, backgrounds, fonts, fontWeights,fontStyles); 
    } 
  }
}

//Output text, fonts etc. to destinated worksheets:
function finalOutput(arr2, medium, fontColors, backgrounds, fonts, fontWeights,fontStyles){
  
  //Array Operations for Final output:
  var A = [];
  var B = [];
  var C = [];

  for (m = 0; m < arr2.length; m++){
    A = medium.slice(0);
    for(n = 0; n < arr2[i].length; n++){
        C.push(arr2[m][n]);
      }
    for(k = 0; k < C.length; k++){
      A.push(C[k]);
    }
    B.push(A);
    A = [];
    C = [];
  }
  
  for (q = 0; q < B.length; q++){
    //Determine the agencies' names:
    if(B[q][11] == 'Test2'){
      var outputArray1 = morphIntoMatrix(B[q]);
      var len1 = B[q].length;
      copyRow1(outputArray1, len1 ,fontColors, backgrounds, fonts, fontWeights,fontStyles);
    }
    else{
      if(B[q][11] == 'Test3'){
        var outputArray2 = morphIntoMatrix(B[q]);
        var len2 = B[q].length;
        copyRow2(outputArray2, len2, fontColors, backgrounds, fonts, fontWeights,fontStyles);
      }
      else{
        if(B[q][11] == 'Test4'){
            var outputArray3 = morphIntoMatrix(B[q]);
            var len3 = B[q].length;
            copyRow3(outputArray3,len3, fontColors, backgrounds, fonts, fontWeights,fontStyles);
        }
      }
    }
  }
}

//Helper function to decide what's the last Row of the merge Ranges:
function lastRowOfRange(values){
  var row = 0;
  for (var row = 0; row < values.length; row++){
    if(!values[row].join(""))
      break;
  }
  return (row+1);
}

//Copy the Header to every single agencies' sheets:
function copyHeader(){
  //Get Header From the Parent Sheet:
  var sheet = SpreadsheetApp.getActiveSheet();
  var applyRange = sheet.getRange(1,1,2,sheet.getLastColumn());
  var headCol = sheet.getLastColumn();
  var value = applyRange.getValues();
  var bGcolors = applyRange.getBackgrounds();
  var colors = applyRange.getFontColors();
  var fontSizes = applyRange.getFontSizes();
  
  //Set the row values & format to destinated sheets:

  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  var targetRange1 = s1.getRange(1,1,2, headCol);
  s1.clear();
  
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  var targetRange2 = s2.getRange(1,1,2, headCol);
  s2.clear();
  
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  var targetRange3 = s3.getRange(1,1,2, headCol);
  s3.clear();
  
  //Copy the Header:
  var copiedsheet = applyRange.getSheet().copyTo(ss1);
  copiedsheet.getRange(applyRange.getA1Notation()).copyTo(targetRange1);
  copiedsheet.getRange(applyRange.getA1Notation()).copyTo(targetRange1,SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  ss1.deleteSheet(copiedsheet);

  //Copy the Header:
  var copiedsheet = applyRange.getSheet().copyTo(ss2);
  copiedsheet.getRange(applyRange.getA1Notation()).copyTo(targetRange2);
  copiedsheet.getRange(applyRange.getA1Notation()).copyTo(targetRange2,SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  ss2.deleteSheet(copiedsheet);  
  
  //Copy the Header:
  var copiedsheet = applyRange.getSheet().copyTo(ss3);
  copiedsheet.getRange(applyRange.getA1Notation()).copyTo(targetRange3);
  copiedsheet.getRange(applyRange.getA1Notation()).copyTo(targetRange3,SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  ss3.deleteSheet(copiedsheet);  
}

//Morph array into setable formats:
function morphIntoMatrix(array) {

  // Create a new array and set the first row of that array to be the original array
  // This is a sloppy workaround to "morphing" a 1-d array into a 2-d array
  var matrix = new Array();
  matrix[0] = array;

  // "Sanitize" the array by erasing null/"null" values with an empty string ""
  for (var i = 0; i < matrix.length; i ++) {
    for (var j = 0; j < matrix[i].length; j ++) {
      if (matrix[i][j] == null || matrix[i][j] == "null") {
        matrix[i][j] = "";
      }
    }
  }
  return matrix;
}

//Output to first agency:
function copyRow1(array, testLen, fontColors, backgrounds, fonts, fontWeights,fontStyles){
  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  var targetRange1 = s1.getDataRange();
  var val1 = targetRange1.getValues();
  var lastRow1 = lastRowOfRange(val1);
  s1.getRange(lastRow1, 1, 1, testLen).setValues(array);
  s1.getRange(lastRow1, 1, 1, testLen).setFontColors(fontColors);
  s1.getRange(lastRow1, 1, 1, testLen).setBackgrounds(backgrounds);
  s1.getRange(lastRow1, 1, 1, testLen).setFontWeights(fontWeights);
  s1.getRange(lastRow1, 1, 1, testLen).setFontFamilies(fonts);
  s1.getRange(lastRow1, 1, 1, testLen).setFontStyles(fontStyles);
}

//Output to second agency:
function copyRow2(array, testLen, fontColors, backgrounds, fonts, fontWeights,fontStyles){
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  var targetRange2 = s2.getDataRange();
  var val2 = targetRange2.getValues();
  var lastRow2 = lastRowOfRange(val2);
  s2.getRange(lastRow2, 1, 1, testLen).setValues(array);
  s2.getRange(lastRow2, 1, 1, testLen).setFontColors(fontColors);
  s2.getRange(lastRow2, 1, 1, testLen).setBackgrounds(backgrounds);
  s2.getRange(lastRow2, 1, 1, testLen).setFontWeights(fontWeights);
  s2.getRange(lastRow2, 1, 1, testLen).setFontFamilies(fonts);
  s2.getRange(lastRow2, 1, 1, testLen).setFontStyles(fontStyles);
}

//Output to third agency:
function copyRow3(array, testLen, fontColors, backgrounds, fonts, fontWeights,fontStyles){
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  var targetRange3 = s3.getDataRange();
  var val3 = targetRange3.getValues();
  var lastRow3 = lastRowOfRange(val3);
  s3.getRange(lastRow3, 1, 1, testLen).setValues(array);
  s3.getRange(lastRow3, 1, 1, testLen).setFontColors(fontColors);
  s3.getRange(lastRow3, 1, 1, testLen).setBackgrounds(backgrounds);
  s3.getRange(lastRow3, 1, 1, testLen).setFontWeights(fontWeights);
  s3.getRange(lastRow3, 1, 1, testLen).setFontFamilies(fonts);
  s3.getRange(lastRow3, 1, 1, testLen).setFontStyles(fontStyles);
}

// Clear out all the agencies' sheets:
function clear(){
  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  s1.getRange('A:Q').clearContent();
  s1.getRange('A:Q').setBackground(null);
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  s2.getRange('A:Q').clearContent();
  s2.getRange('A:Q').setBackground(null);
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  s3.getRange('A:Q').clearContent();
  s3.getRange('A:Q').setBackground(null);
}

//Set UI:
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🚀Bigo Distribution');
  var item1 = menu.addItem('💫Distribute!', 'liveHouse');
  var item2 = menu.addItem('💫Delete Everything!','clear');
  item1.addToUi();
  item2.addToUi();
}
  
