var agentID1 = '1k9hVZU41Wn_7fvjZ-e6CQIC5BTSR7kiDJKPyPfbrMCQ';
var agentSheetName1 = 'Sheet2';
var agentID2 = '1lonWd4w4SZ4MD75KEvdtLnydyrtzI9OL5-6L-qAnmpI';
var agentSheetName2 = 'Sheet2';
var agentID3 = '1Vcr3R_iiR60JRrNV1mavNdOf7jFr8AEy9cR5IxZbidM';
var agentSheetName3 = 'Sheet2';

//Main Function to distribute:
function PK(){
  //Copy Pk's header (first three rows):
  copyHeaderPK();

  //Read data:
  var spreadsheet = SpreadsheetApp.openById('1buA9zuTTPI0iBd13YtY0kRSuSsF5UMKk0_9s_DkBkCs');
  var sheet = spreadsheet.getSheetByName('Sheet2');
  var range = sheet.getRange(3,1,sheet.getLastRow(),sheet.getLastColumn());
  var values = range.getValues();
  var mergeArr = [];
  var splitArr = [];
  var size = 4;
  var val = [];

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
  
  Logger.log(splitArr);
  //Logging the start/end ranges:
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
  
  for (l = 0; l < E[0].length; l++){
    var middle = 3;
    var start = E[0][l];
    var end = E[1][l];
    var rng;
    var fontColors;
    var backgrounds;
    var fonts;
    var fontWeights;
    var fontStyles;
    colHeader = end + 1;

    //For the first Range:
    if(l == 0){
      value[l] = sheet.getRange(start,10, end-middle, sheet.getLastColumn()).getValues();
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
      //Output the merge cells to relevant work sheets:
      finalOutput(H,start,sheet);
      //Copy Every other rows:
      copyColumnHeader(colHeader);
      transpose = [];
      F = [];
      H = [];
      value = [];
    }
    //For the first range below:
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
      finalOutput(H,start,sheet);
      copyColumnHeader(colHeader);
      F = [];
      H = [];
    }
  }
}

//Copy header function:
function copyColumnHeader(colHeader){
  //Get Header From the Parent Sheet:
  var spreadsheet = SpreadsheetApp.openById('1buA9zuTTPI0iBd13YtY0kRSuSsF5UMKk0_9s_DkBkCs');
  var sheet = spreadsheet.getSheetByName('Sheet2');
  var applyRange = sheet.getRange(colHeader,1,1,sheet.getLastColumn());
  var headCol = sheet.getLastColumn();
  var value = applyRange.getValues();
  var bGcolors = applyRange.getBackgrounds();
  var colors = applyRange.getFontColors();
  var fontSizes = applyRange.getFontSizes();
  var fontWeights = applyRange.getFontWeights();
  var fonts = applyRange.getFontFamilies();

  //Set the row values & format to destinated sheets:
  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  var targetRange1 = s1.getDataRange();
  var val1 = targetRange1.getValues();
  var lastRow1 = lastRowOfRange(val1);
  s1.getRange(lastRow1,1,1,headCol).setValues(value);
  s1.getRange(lastRow1,1,1,headCol).setFontColors(colors);
  s1.getRange(lastRow1,1,1,headCol).setBackgrounds(bGcolors);
  s1.getRange(lastRow1,1,1,headCol).setFontWeights(fontWeights);
  s1.getRange(lastRow1,1,1,headCol).setFontSizes(fontSizes);
  s1.getRange(lastRow1,1,1,headCol).setFontFamilies(fonts);
  
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  var targetRange2 = s2.getDataRange();
  var val2 = targetRange2.getValues();
  var lastRow2 = lastRowOfRange(val2);
  s2.getRange(lastRow2,1,1,headCol).setValues(value);
  s2.getRange(lastRow2,1,1,headCol).setFontColors(colors);
  s2.getRange(lastRow2,1,1,headCol).setBackgrounds(bGcolors);
  s2.getRange(lastRow2,1,1,headCol).setFontWeights(fontWeights);
  s2.getRange(lastRow2,1,1,headCol).setFontSizes(fontSizes);
  s2.getRange(lastRow2,1,1,headCol).setFontFamilies(fonts);
  
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  var targetRange3 = s3.getDataRange();
  var val3 = targetRange3.getValues();
  var lastRow3 = lastRowOfRange(val3);
  s3.getRange(lastRow3,1,1,headCol).setValues(value);
  s3.getRange(lastRow3,1,1,headCol).setFontColors(colors);
  s3.getRange(lastRow3,1,1,headCol).setBackgrounds(bGcolors);
  s3.getRange(lastRow3,1,1,headCol).setFontWeights(fontWeights);
  s3.getRange(lastRow3,1,1,headCol).setFontSizes(fontSizes);
  s3.getRange(lastRow3,1,1,headCol).setFontFamilies(fonts);
}

//Helper function to decide 
function findSameValue(J,q){
  var aa = [];
  for (u = 0; u < J.length; u++){  
    if(J[q][2] == J[u][2] && J[q][3] == J[u][3] && J[q][1] == J[u][1] && J[q][9] != J[u][9]){
      aa.push(J[q]);
      aa.push(J[u]);
    } 
  }
  return aa;
}

//Output functions:
function finalOutput(H,start,sheet){
  var med = [];
  var J = H.slice(0);
  for(q = 0; q < H.length; q++){
    if(H[q][9] == 'test2'){
      Logger.log(H[q]);
      med[q] = findSameValue(J,q);
      Logger.log('med %s', med[q]);
      var outputArray1 = morph3D(med[q]);
      var len1 = H[q].length;
      var rng = sheet.getRange(start,1,2,len1);
      var fontColors = rng.getFontColors();
      var backgrounds = rng.getBackgrounds();
      var fonts = rng.getFontFamilies();
      var fontWeights = rng.getFontWeights();
      var fontStyles = rng.getFontStyles();
      copyRow1(outputArray1, len1,fontColors, backgrounds, fonts, fontWeights,fontStyles);
      med = [];
    }
    else{
      if(H[q][9] == 'test3'){
        med[q] = findSameValue(J,q);
        var outputArray2 = morph3D(med[q]);
        var len2 = H[q].length;
        var rng = sheet.getRange(start,1,2,len2);
        var fontColors = rng.getFontColors();
        var backgrounds = rng.getBackgrounds();
        var fonts = rng.getFontFamilies();
        var fontWeights = rng.getFontWeights();
        var fontStyles = rng.getFontStyles();
        copyRow2(outputArray2, len2,fontColors, backgrounds, fonts, fontWeights,fontStyles);
        med = []
      }
      else{
        if(H[q][9] == 'test4'){
          med[q] = findSameValue(J,q);
          var outputArray3 = morph3D(med[q]);
          var len3 = H[q].length;
          var rng = sheet.getRange(start,1,2,len3);
          var fontColors = rng.getFontColors();
          var backgrounds = rng.getBackgrounds();
          var fonts = rng.getFontFamilies();
          var fontWeights = rng.getFontWeights();
          var fontStyles = rng.getFontStyles();
          copyRow3(outputArray3,len3,fontColors, backgrounds, fonts, fontWeights,fontStyles);
        }
      }
    }
  }
}

//Morph two dimensional array into 3D to set values:
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

//Copy Row to agency 1:
function copyRow1(array, testLen,fontColors, backgrounds, fonts, fontWeights,fontStyles){
  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  var targetRange1 = s1.getDataRange();
  var val1 = targetRange1.getValues();
  var lastRow1 = lastRowOfRange(val1);

  //Hide Agency & UIDs:
  array[1][9] = '';
  array[1][11] = '';
  array[0][11] = '';

  s1.getRange(lastRow1, 1, 2, testLen).setValues(array);
  s1.getRange(lastRow1, 1, 2, testLen).setFontColors(fontColors);
  s1.getRange(lastRow1, 1, 2, testLen).setBackgrounds(backgrounds);
  s1.getRange(lastRow1, 1, 2, testLen).setFontWeights(fontWeights);
  s1.getRange(lastRow1, 1, 2, testLen).setFontFamilies(fonts);
  s1.getRange(lastRow1, 1, 2, testLen).setFontStyles(fontStyles);
  for (i = 1; i < 10; i++){
    s1.getRange(lastRow1, i, 2, 1).merge();
  }
}


//Copy Row to agency2:
function copyRow2(array, testLen,fontColors, backgrounds, fonts, fontWeights,fontStyles){
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  var targetRange2 = s2.getDataRange();
  var val2 = targetRange2.getValues();
  var lastRow2 = lastRowOfRange(val2);
  
  //Hide Agency & UIDs:
  array[1][9] = '';
  array[1][11] = '';
  array[0][11] = '';
  
  s2.getRange(lastRow2, 1, 2, testLen).setValues(array);
  s2.getRange(lastRow2, 1, 2, testLen).setFontColors(fontColors);
  s2.getRange(lastRow2, 1, 2, testLen).setBackgrounds(backgrounds);
  s2.getRange(lastRow2, 1, 2, testLen).setFontWeights(fontWeights);
  s2.getRange(lastRow2, 1, 2, testLen).setFontFamilies(fonts);
  s2.getRange(lastRow2, 1, 2, testLen).setFontStyles(fontStyles);
  for (i = 1; i < 10; i++){
    s2.getRange(lastRow2, i, 2, 1).merge();
  }
}

//Copy Row to agency3:
function copyRow3(array, testLen,fontColors, backgrounds, fonts, fontWeights,fontStyles){
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  var targetRange3 = s3.getDataRange();
  var val3 = targetRange3.getValues();
  var lastRow3 = lastRowOfRange(val3);
  
  //Hide Agency & UIDs:
  array[1][9] = '';
  array[1][11] = '';
  array[0][11] = '';
  
  s3.getRange(lastRow3, 1, 2, testLen).setValues(array);
  s3.getRange(lastRow3, 1, 2, testLen).setFontColors(fontColors);
  s3.getRange(lastRow3, 1, 2, testLen).setBackgrounds(backgrounds);
  s3.getRange(lastRow3, 1, 2, testLen).setFontWeights(fontWeights);
  s3.getRange(lastRow3, 1, 2, testLen).setFontFamilies(fonts);
  s3.getRange(lastRow3, 1, 2, testLen).setFontStyles(fontStyles);
  for (i = 1; i < 10; i++){
    s3.getRange(lastRow3, i, 2, 1).merge();
  }
}

//RETURN the array of unmerge/merge cells:
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
      Logger.log('filter:%s', filterRows);
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
  Logger.log('G: %s', G);
  for (r = 0 ; r < transpose.length; r++){
    G.push(transpose[r]);
  }
  var H = transposeArray(G);
  return H;
}


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

//Transpose horizontal array vertically:
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

//Return the last row of merged ranges:
function lastRowOfRange(values){
  var row = 0;
  for (var row = 0; row < values.length; row++){
    if(!values[row].join(""))
      break;
  }
  return (row+1);
}

//Copy Header:
function copyHeaderPK(){
  //Get Header From the Parent Sheet:
  var spreadsheet = SpreadsheetApp.openById('1buA9zuTTPI0iBd13YtY0kRSuSsF5UMKk0_9s_DkBkCs');
  var sheet = spreadsheet.getSheetByName('Sheet2');
  var applyRange = sheet.getRange(1,1,3,sheet.getLastColumn());
  var headCol = sheet.getLastColumn();
  var value = applyRange.getValues();
  var bGcolors = applyRange.getBackgrounds();
  var colors = applyRange.getFontColors();
  var fontSizes = applyRange.getFontSizes();
  
  //Set the row values & format to destinated sheets:

  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  var targetRange1 = s1.getRange(1,1,3, headCol);
  s1.clear();
  
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  var targetRange2 = s2.getRange(1,1,3, headCol);
  s2.clear();
  
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  var targetRange3 = s3.getRange(1,1,3, headCol);
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

//Clear agencies' worksheets:
function clear(){
  var ss1 = SpreadsheetApp.openById(agentID1);
  var s1 = ss1.getSheetByName(agentSheetName1); 
  s1.getRange('A:Q').clear();
  var ss2 = SpreadsheetApp.openById(agentID2);
  var s2 = ss2.getSheetByName(agentSheetName2); 
  s2.getRange('A:Q').clear();
  var ss3 = SpreadsheetApp.openById(agentID3);
  var s3 = ss3.getSheetByName(agentSheetName3); 
  s3.getRange('A:Q').clear();
}

//Set UI:
function onOpenPk(){
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🚀PK');
  var item1 = menu.addItem('💫PK Distribute!', 'PK');
  var item2 = menu.addItem('💫Delete Everything!','clear');
  item1.addToUi();
  item2.addToUi();
}
