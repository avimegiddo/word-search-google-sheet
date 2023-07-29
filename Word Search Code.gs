// This function generates a word search puzzle based on the words in column A.
// It creates a new sheet with the puzzle and adds the suffix to the sheet name.
// If a sheet with the same name already exists, it increments the suffix.
// The function also creates a duplicate sheet with the answers hidden.
function generateWordSearch() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = ss.getActiveSheet();
 var originalSheetName = sheet.getName();

 // Strip off "ANSWER KEY" and "HIDDEN ANSWERS" from the sheet name
 var strippedSheetName = originalSheetName.replace(/(ANSWER KEY|HIDDEN ANSWERS)/g, '').trim();

 // Set the new sheet name with suffix
 var suffix = 1;
 var match = strippedSheetName.match(/\d+$/);
 var newSheetName;

 if (match) {
   var originalSuffix = parseInt(match[0]);
   if (!isNaN(originalSuffix)) {
     suffix = originalSuffix + 1;
     newSheetName = strippedSheetName.replace(/\d+$/, suffix.toString());
   }
 } else {
   newSheetName = "puzzle " + suffix;
 }

 while (ss.getSheetByName(newSheetName + " - ANSWER KEY") || ss.getSheetByName(newSheetName + " - HIDDEN ANSWERS")) {
   suffix++;
   newSheetName = "puzzle " + suffix;
 }

 // Append the suffix and " - ANSWER KEY" to the new sheet names
 var answerSheetName = newSheetName + " - ANSWER KEY";
 var hideAnswersSheetName = newSheetName + " - HIDDEN ANSWERS";

 var newSheet = sheet.copyTo(ss);

 // Set the name for the answer sheet
 var answerSheet = newSheet.setName(answerSheetName);

 // Move the new sheet to the left-most position
 ss.setActiveSheet(answerSheet);
 ss.moveActiveSheet(1);


 // Make the word search grid more like squares
 resizeColumns();

 // Make all word search words lowercase (to hide them better)
 lowercaseColumnA();

 var wordsRange = SpreadsheetApp.getActiveSpreadsheet().getRange("A2:A");
 var words = wordsRange.getValues().filter(String).map(function (row) { return row[0]; });
 var grid = createEmptyGrid(20, 20); // Change grid size to 20x20
 var directions = [
   { name: "horizontal", rowDelta: 0, colDelta: 1 },
   { name: "reverse horizontal", rowDelta: 0, colDelta: -1 },
   { name: "vertical", rowDelta: 1, colDelta: 0 },
   { name: "reverse vertical", rowDelta: -1, colDelta: 0 },
   { name: "diagonalNE", rowDelta: 1, colDelta: 1 },
   { name: "reverse diagonalSE", rowDelta: 1, colDelta: -1 },
   { name: "diagonalNW", rowDelta: -1, colDelta: 1 },
   { name: "reverse diagonalSW", rowDelta: -1, colDelta: -1 }
 ];

 // Reset all cells to default background color and set font size and alignment
 answerSheet.getRange("B2:U21").setBackground(null);
 answerSheet.getRange("B2:U21").setFontSize(14);
 answerSheet.getRange("B2:U21").setVerticalAlignment("middle");
 answerSheet.getRange("B2:U21").setHorizontalAlignment("center");

 for (var i = 0; i < words.length; i++) {
   var word = words[i];
   var direction = directions[Math.floor(Math.random() * directions.length)];
   var coords = findEmptySpace(grid, word.length, direction.rowDelta, direction.colDelta, word);
   if (coords) {
     placeWord(grid, word, coords.row, coords.col, direction.rowDelta, direction.colDelta);
   } else {
     answerSheet.getRange("A" + (i + 2)).setFontLine("line-through");
     // Couldn't place word, strike through the word in column A
   }
 }

 // Fill the empty cells with random letters
 for (var i = 0; i < grid.length; i++) {
   for (var j = 0; j < grid[0].length; j++) {
     if (grid[i][j] === "") {
       grid[i][j] = String.fromCharCode(97 + Math.floor(Math.random() * 26));
     }
   }
 }
 answerSheet.getRange("B2:U21").setValues(grid); // Change range to B2:U21

 var hideAnswersSheet = newSheet.copyTo(ss);
 hideAnswersSheet.setName(hideAnswersSheetName);

 // Reset all cells to default background color
 hideAnswersSheet.getDataRange().setBackground(null);

 // Move the hide answers sheet to the leftmost position
 ss.setActiveSheet(hideAnswersSheet);
 ss.moveActiveSheet(1);
}


// This function prints the current sheet to PDF format.
// It exports the sheet as a PDF file and saves it to Google Drive.
// The PDF file is then opened in a new Chrome tab.
function printToPDF() {
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getActiveSheet();
 var sheetName = sheet.getName();
 var range = 'A2:U40';
 var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' + '&size=A4&landscape=true&fitw=true' + '&sheetnames=false&printtitle=false&pagenumbers=false' + '&gridlines=false&fzr=false&range=' + range + '&gid=' + sheet.getSheetId();
 var blob = UrlFetchApp.fetch(url, {
   headers: {
     'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
   }
 }).getBlob().setName(sheetName + '.pdf');

 // Create the PDF file in Google Drive
 var folder = DriveApp.getRootFolder();
 var file = folder.createFile(blob);

 // Open the PDF file in a new Chrome tab
 var pdfUrl = file.getUrl();
 var html = '<script>window.open("' + pdfUrl + '");</script>';
 var userInterface = HtmlService.createHtmlOutput(html);
 SpreadsheetApp.getUi().showModalDialog(userInterface, 'Generated PDF');
}

// This function resizes the columns in the active sheet.
// It sets the width of columns 2 to 21 to 36 pixels.
function resizeColumns() {
 var sheet = SpreadsheetApp.getActiveSheet();
 sheet.setColumnWidths(2, 20, 36);
}

// This function converts all values in column A to lowercase.
// It modifies the values directly in the sheet.
function lowercaseColumnA() {
 var sheet = SpreadsheetApp.getActiveSheet();
 var columnA = sheet.getRange("A:A").getValues();
 for (var i = 0; i < columnA.length; i++) {
   columnA[i][0] = columnA[i][0].toString().toLowerCase();
 }
 sheet.getRange("A:A").setValues(columnA);
}

// This function creates an empty grid with the specified number of rows and columns.
// It returns a 2D array representing the grid.
function createEmptyGrid(rows, cols) {
 var grid = [];
 for (var i = 0; i < rows; i++) {
   grid[i] = [];
   for (var j = 0; j < cols; j++) {
     grid[i][j] = "";
   }
 }
 return grid;
}

// This function places a word in the word search grid.
// It fills the cells corresponding to the word with the word's characters.
// It also sets a random pastel color for the word in the answer key.
function placeWord(grid, word, startRow, startCol, rowDelta, colDelta) {
 // Define an array of pastel colors for the answer key
 var pastelColors = ["#b7d8b7", "#b7c9d8", "#d8b7c9", "#d8d0b7", "#c9b7d8", "#d0d8b7", "#b7d8c9", "#d8b7b7", "#d8c9b7", "#c9d8b7", "#d5a6bd", "#fff2cc", "#d9d2e9", "#b6d7a8", "#fce5cd", "#e6b8af", "#d0e0e3", "#f4cccc", "#ead1dc", "#cfe2f3"];

 // Choose a random pastel color from the array for each word
 var color = pastelColors[Math.floor(Math.random() * pastelColors.length)];

 for (var i = 0; i < word.length; i++) {
   var row = startRow + i * rowDelta;
   var col = startCol + i * colDelta;
   grid[row][col] = word.charAt(i);
   var sheet = SpreadsheetApp.getActiveSheet();
   var range = sheet.getRange(row + 2, col + 2);
   range.setBackground(color);
 }
}

// This function finds an empty space in the word search grid for placing a word.
// It searches for a space that can accommodate the word in the specified direction.
// If a space is found, it returns the coordinates of the starting cell.
// If no suitable space is found, it returns null.
function findEmptySpace(grid, length, rowDelta, colDelta, word = "") {
 var rows = grid.length;
 var cols = grid[0].length;
 var maxRow = rows - length * rowDelta;
 var maxCol = cols - length * colDelta;

 // Randomly select a starting row and column
 var startRow = Math.floor(Math.random() * rows);
 var startCol = Math.floor(Math.random() * cols);

 // Move through the grid in a deterministic way
 for (var i = 0; i < rows; i++) {
   for (var j = 0; j < cols; j++) {
     var row = (startRow + i) % rows;
     var col = (startCol + j) % cols;
     var spaceFound = true;
     for (var k = 0; k < length; k++) {
       var checkRow = row + k * rowDelta;
       var checkCol = col + k * colDelta;
       if (checkRow < 0 || checkRow >= rows || checkCol < 0 || checkCol >= cols || (grid[checkRow][checkCol] !== "" && (k >= word.length || word[k] !== grid[checkRow][checkCol]))) {
         spaceFound = false;
         break;
       }
     }
     if (spaceFound) {
       return { row: row, col: col };
     }
   }
 }
 return null;
}
