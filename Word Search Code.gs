 var availableColors = [
    // Original Colors with English Names
    "#b7d8b7", // Soft Mint
    "#b7c9d8", // Light Sky Blue
    "#d8b7c9", // Pale Pink Purple
    "#d8d0b7", // Light Khaki
    "#c9b7d8", // Lavender
    "#d0d8b7", // Pale Lime Green
    "#b7d8c9", // Soft Aqua
    "#d8b7b7", // Soft Pink
    "#d8c9b7", // Pale Sand
    "#c9d8b7", // Light Green
    "#d5a6bd", // Muted Pink
    "#fff2cc", // Pale Yellow
    "#d9d2e9", // Pale Purple
    "#b6d7a8", // Light Green
    "#fce5cd", // Peach Cream
    "#e6b8af", // Soft Coral
    "#d0e0e3", // Pale Turquoise
    "#f4cccc", // Pastel Red
    "#ead1dc", // Pale Lavender Pink
    "#cfe2f3", // Soft Blue
    "#add8e6", // Light Blue
    "#f0e68c", // Khaki
    "#ffb6c1", // Light Pink
    "#d8bfd8", // Thistle
    "#dda0dd", // Plum
    "#ffe4e1", // Misty Rose
    "#ffebcd", // Blanched Almond
    "#fafad2", // Light Goldenrod Yellow
    "#ffe4b5", // Moccasin
    "#ffdead", // Navajo White
    "#f0e5c9", // Cream
    "#faf0e6", // Linen
    "#e6e6fa", // Lavender
    "#fff5ee", // Seashell
    "#f5f5dc", // Beige
    "#fdfd96", // Pastel Yellow
    "#a4c2f4", // Soft Blue
    "#9fc5e8", // Sky Blue
    "#6d9eeb", // Periwinkle Blue
    "#c9daf8", // Lavender Blue
    "#76a5af", // Soft Teal
    "#92cddc", // Light Blue-Green
    "#b4a7d6", // Light Purple
    "#8e7cc3", // Lavender Purple
    "#6fa8dc", // Cornflower Blue
    "#8faabd",  // Slate Gray Blue
    "#e3eaa7", // Pastel Lime
    "#d9e4fc", // Pale Sky Blue
    "#c5e3f6", // Powder Blue
    "#f7fcb9", // Soft Lemon
    "#d8e4bc", // Pale Olive Green
    "#cad2c5", // Sage Green
    "#e4e4c5", // Light Moss
    "#b8e0a2", // Pale Spring Green
    "#a2c4c9", // Soft Cyan
    "#ace1af",  // Celadon Green
    "#FFB3BA", // Light Salmon Pink
    "#FFDFBA", // Light Peach
    "#FFFFBA", // Light Butter Yellow
    "#BAFFC9", // Mint Green
    "#BAE1FF", // Baby Blue
  ];;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Word Search')
    .addItem('Make Puzzle & Answer Key', 'generateWordSearch')
    .addItem('Print PDF', 'printToPDF')
    .addItem('Check Word Selection', 'checkWordSelection')
    .addToUi();
}

var wordLocations = {}; // Global object to store the coordinates of each word


function getWordsFromFirstColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getRange("A2:A"); // Assuming the words start from A2 downwards
  var values = range.getValues();
 
  // Flatten the array and filter out empty cells
  var words = values.map(function(row) {
    return row[0].toLowerCase().trim(); // Convert to lowercase and trim spaces
  }).filter(Boolean); // Remove empty cells

  return words;
}



function createDynamicGrid(words) {
  // Calculate the total characters and average word length
  var totalChars = words.reduce((sum, word) => sum + word.length, 0);
  var longestWordLength = Math.max(...words.map(word => word.length));
  var avgWordLength = totalChars / words.length;

  // Calculate the grid size based on number of words and average word length
  var gridSize = Math.max(Math.ceil(Math.sqrt(words.length * avgWordLength * 2)), longestWordLength, 5);

  // Ensure that grid size is capped at a maximum of 20x20
  gridSize = Math.min(gridSize, 20);

  return createEmptyGrid(gridSize, gridSize); // Create grid with dynamic size
}

function generateWordSearch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var originalSheetName = sheet.getName();

  // Strip off "ANSWER KEY" and "HIDDEN ANSWERS" from the sheet name for clarity
  var strippedSheetName = originalSheetName.replace(/(ANSWER KEY|HIDDEN ANSWERS)/g, '').trim();
  var wordsRange = sheet.getRange("A2:A").getValues()
    .filter(function (row) { return row[0] !== ""; })
    .map(function (row) { return row[0].toLowerCase(); });
  // Set the new sheet name with a numeric suffix to ensure uniqueness
  var suffix = 1;
  var match = strippedSheetName.match(/\d+$/);
  var newSheetName;

  // If the current sheet already has a numeric suffix, increment it
  if (match) {
    var originalSuffix = parseInt(match[0]);
    if (!isNaN(originalSuffix)) {
      suffix = originalSuffix + 1;
      newSheetName = strippedSheetName.replace(/\d+$/, suffix.toString());
    }
  } else {
    newSheetName = "puzzle " + suffix;
  }

  // Check for existing sheets with the same name, increment the suffix if needed
  while (ss.getSheetByName(newSheetName + " - ANSWER KEY") || ss.getSheetByName(newSheetName + " - HIDDEN ANSWERS")) {
    suffix++;
    newSheetName = "puzzle " + suffix;
  }

  // Append " - ANSWER KEY" and " - HIDDEN ANSWERS" to create sheet names for the puzzle and the answer key
  var answerSheetName = newSheetName + " - ANSWER KEY";
  var hideAnswersSheetName = newSheetName + " - HIDDEN ANSWERS";

  // **Reset all formatting in column A** to remove previous strike-throughs or other styles
  sheet.getRange("A:A").clearFormat();

  // **Set font size to 14 in column A**
  sheet.getRange("A:A").setFontSize(14);

  // Call lowercaseColumnA to convert all words in column A to lowercase
  lowercaseColumnA();

  // Get the words from column A starting at A2, and convert all to lowercase
  var wordsRange = sheet.getRange("A2:A");
  var words = wordsRange.getValues()
    .filter(function (row) { return row[0] !== ""; })
    .map(function (row) { return row[0].toLowerCase(); }); // **Apply toLowerCase() here**

  // If no words are found in column A, show an alert and stop execution
  if (words.length === 0) {
    SpreadsheetApp.getUi().alert("No words found! Please input words in column A.");
    return;
  }

  // Duplicate the current sheet as the base for the answer sheet
  var newSheet = sheet.copyTo(ss);

  // Rename the new sheet to be the answer sheet
  var answerSheet = newSheet.setName(answerSheetName);

  // **Reset all formatting in column A in the answerSheet**
  answerSheet.getRange("A:A").clearFormat();

  // **Set font size to 14 in column A in the answerSheet**
  answerSheet.getRange("A:A").setFontSize(14);

  // Move the new answer sheet to the leftmost position
  ss.setActiveSheet(answerSheet);
  ss.moveActiveSheet(1);

  // Clear any content and formatting in the grid range (B2:U21) to reset for a new puzzle
  answerSheet.getRange("B2:U21").clearContent();
  answerSheet.getRange("B2:U21").setBackground(null);

  // **Set font size to 14 in the grid area**
  answerSheet.getRange("B2:U21").setFontSize(14);

  // Resize columns for consistent formatting
  resizeColumns();

  // Now the words are guaranteed to be lowercase, and the rest of the logic can proceed
  // Further logic remains the same (placing words in the grid, calculating difficulty, etc.)

  // Calculate total characters and average word length
  var totalChars = words.reduce((sum, word) => sum + word.length, 0);
  var avgWordLength = totalChars / words.length;

  // Determine difficulty level based on word count and average word length
  var directions;
  if (words.length <= 5 && avgWordLength <= 4) {
    // Easier level: only horizontal and vertical
    directions = [
      { name: "horizontal", rowDelta: 0, colDelta: 1 },
      { name: "vertical", rowDelta: 1, colDelta: 0 }
    ];
  } else if (words.length <= 10 && avgWordLength <= 6) {
    // Middle level: add diagonal but no backward words
    directions = [
      { name: "horizontal", rowDelta: 0, colDelta: 1 },
      { name: "vertical", rowDelta: 1, colDelta: 0 },
      { name: "diagonalNE", rowDelta: 1, colDelta: 1 },
      { name: "diagonalNW", rowDelta: 1, colDelta: -1 }
    ];
  } else {
    // Advanced level: allow backward and diagonal backward words
    directions = [
      { name: "horizontal", rowDelta: 0, colDelta: 1 },
      { name: "reverse horizontal", rowDelta: 0, colDelta: -1 },
      { name: "vertical", rowDelta: 1, colDelta: 0 },
      { name: "reverse vertical", rowDelta: -1, colDelta: 0 },
      { name: "diagonalNE", rowDelta: 1, colDelta: 1 },
      { name: "reverse diagonalSE", rowDelta: 1, colDelta: -1 },
      { name: "diagonalNW", rowDelta: -1, colDelta: 1 },
      { name: "reverse diagonalSW", rowDelta: -1, colDelta: -1 }
    ];
  }

  // Create a dynamic grid based on word list
  var grid = createDynamicGrid(words);

  // Reset all cells in the grid to default settings
  var gridSize = grid.length;
  answerSheet.getRange(2, 2, gridSize, gridSize).setBackground(null);
  answerSheet.getRange(2, 2, gridSize, gridSize).setFontSize(14);
  answerSheet.getRange(2, 2, gridSize, gridSize).setVerticalAlignment("middle");
  answerSheet.getRange(2, 2, gridSize, gridSize).setHorizontalAlignment("center");

  // Place each word on the grid
  for (var i = 0; i < words.length; i++) {
    var word = words[i];
    var direction = directions[Math.floor(Math.random() * directions.length)];
    var coords = findEmptySpace(grid, word.length, direction.rowDelta, direction.colDelta, word);
    if (coords) {
      placeWord(grid, word, coords.row, coords.col, direction.rowDelta, direction.colDelta);
    } else {
      // If the word can't be placed, strike it through in column A
      answerSheet.getRange("A" + (i + 2)).setFontLine("line-through");
    }
  }

  // Fill any remaining empty cells in the grid with random letters
  for (var i = 0; i < grid.length; i++) {
    for (var j = 0; j < grid[0].length; j++) {
      if (grid[i][j] === "") {
        grid[i][j] = String.fromCharCode(97 + Math.floor(Math.random() * 26)); // Random lowercase letters
      }
    }
  }

  // Set the completed grid on the answer sheet
  answerSheet.getRange(2, 2, gridSize, gridSize).setValues(grid);

  // Create a copy of the answer sheet and hide the answers for the puzzle
  var hideAnswersSheet = newSheet.copyTo(ss);
  hideAnswersSheet.setName(hideAnswersSheetName);

  // **Reset all formatting in column A in the hideAnswersSheet**
  hideAnswersSheet.getRange("A:A").clearFormat();

  // **Set font size to 14 in column A in the hideAnswersSheet**
  hideAnswersSheet.getRange("A:A").setFontSize(14);

  // **Set font size to 14 in the grid area of the hideAnswersSheet**
  hideAnswersSheet.getRange("B2:U21").setFontSize(14);

  // Hide the answers by clearing background colors in the grid area
  hideAnswersSheet.getRange("B2:U21").setBackground(null);

  ss.setActiveSheet(hideAnswersSheet);
  ss.moveActiveSheet(1);
}



function printToPDF() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName();
  var range = 'A1:U40';
  var url = 'https://docs.google.com/spreadsheets/d/' + spreadsheet.getId() + '/export?exportFormat=pdf&format=pdf' +
    '&size=A4&landscape=true&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false' +
    '&gridlines=false&fzr=false&range=' + range + '&gid=' + sheet.getSheetId();

  var blob = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  }).getBlob().setName(sheetName + '.pdf');

  return Utilities.base64Encode(blob.getBytes());
}


function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('HandlePDF')
    .setWidth(200)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate PDF');
}




function checkWordSelection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selectedRangeList = sheet.getActiveRangeList(); // Use getActiveRangeList to handle multiple selections

  if (!selectedRangeList) {
    Logger.log("No range selected");
    return false;
  }

  var selectedCoords = [];

  // Collect all selected cells and their coordinates
  selectedRangeList.getRanges().forEach(function(range) {
    var values = range.getValues();
    var startRow = range.getRow();
    var startCol = range.getColumn();

    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        selectedCoords.push({
          row: startRow + i,
          col: startCol + j,
          letter: values[i][j]
        });
      }
    }
  });

  // Sort the selected cells based on their positions to maintain the selection order
  selectedCoords.sort(function(a, b) {
    if (a.row === b.row) {
      return a.col - b.col;
    }
    return a.row - b.row;
  });

  // Check if the selected cells are in a straight line (horizontal, vertical, or diagonal)
  if (!areCellsInLine(selectedCoords)) {
    Logger.log("Selection is not in a straight line");
    return false;
  }

  // Join the selected letters to form the word
  var selectedWord = selectedCoords.map(cell => cell.letter).join('').toLowerCase();
  var reversedWord = selectedCoords.map(cell => cell.letter).reverse().join('').toLowerCase();

  // Retrieve words from column A
  var words = getWordsFromFirstColumn();

  // Check if the selected word or its reverse matches any word in column A
  if (words.includes(selectedWord) || words.includes(reversedWord)) {
    Logger.log("Correct selection: " + selectedWord);

    // Color the selected cells if correct
    colorCells(selectedCoords);

    // Apply strikethrough to the word in column A
    applyStrikethroughToWord(selectedWord, reversedWord, words);

    // **Check if all words have been found**
    if (allWordsFound()) {
      // Show congratulations alert
      SpreadsheetApp.getUi().alert("Congratulations! You have found all of the words. Puzzle complete!");
    }

    return true;
  } else {
    Logger.log("Incorrect selection: " + selectedWord);
    return false;
  }
}



function allWordsFound() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Get the range of words in column A, starting from row 2 (excluding header)
  var wordsRange = sheet.getRange(2, 1, lastRow - 1); // (startRow, column, numRows)
  var fontLines = wordsRange.getFontLines(); // Returns a 2D array

  // Check if all words have 'line-through' applied
  for (var i = 0; i < fontLines.length; i++) {
    if (fontLines[i][0] !== 'line-through') {
      // Found a word that hasn't been found yet
      return false;
    }
  }

  // All words have 'line-through' applied
  return true;
}


// Helper function to apply strikethrough to the found word in column A
function applyStrikethroughToWord(selectedWord, reversedWord, words) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Find the index of the word in the words array
  var wordIndex = words.indexOf(selectedWord);
  if (wordIndex === -1) {
    wordIndex = words.indexOf(reversedWord);
  }

  if (wordIndex !== -1) {
    // Words start from row 2, so row number is wordIndex + 2
    var wordCell = sheet.getRange(wordIndex + 2, 1); // Column A
    wordCell.setFontLine('line-through');
  }
}


function areCellsInLine(coords) {
  if (coords.length < 2) {
    return true; // Single cell is considered a line
  }

  var deltaRow = coords[1].row - coords[0].row;
  var deltaCol = coords[1].col - coords[0].col;

  // Normalize deltas to -1, 0, or 1
  deltaRow = deltaRow === 0 ? 0 : deltaRow / Math.abs(deltaRow);
  deltaCol = deltaCol === 0 ? 0 : deltaCol / Math.abs(deltaCol);

  for (var i = 1; i < coords.length; i++) {
    var expectedRow = coords[0].row + i * deltaRow;
    var expectedCol = coords[0].col + i * deltaCol;

    if (coords[i].row !== expectedRow || coords[i].col !== expectedCol) {
      return false;
    }
  }
  return true;
}




function areCellsAdjacent(coords) {
  for (var i = 1; i < coords.length; i++) {
    var rowDelta = Math.abs(coords[i].row - coords[i - 1].row);
    var colDelta = Math.abs(coords[i].col - coords[i - 1].col);
   
    // Allow horizontal, vertical, or diagonal adjacency
    if (!(rowDelta === 0 && colDelta === 1 ||   // Horizontal
          rowDelta === 1 && colDelta === 0 ||   // Vertical
          rowDelta === 1 && colDelta === 1)) {  // Diagonal
      return false;
    }
  }
  return true;
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
// This function places a word in the word search grid.
// It fills the cells corresponding to the word with the word's characters.
// It also sets a random pastel color for the word in the answer key.



function placeWord(grid, word, startRow, startCol, rowDelta, colDelta) {
  // Define an array of pastel colors for the answer key
 
  var color = availableColors[Math.floor(Math.random() * availableColors.length)];

  var wordCoords = []; // Store the coordinates for the word

  for (var i = 0; i < word.length; i++) {
    var row = startRow + i * rowDelta;
    var col = startCol + i * colDelta;
    grid[row][col] = word.charAt(i);

    // Store the coordinate for the letter
    wordCoords.push({ row: row + 2, col: col + 2 }); // +2 to adjust for grid offset in the sheet

    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getRange(row + 2, col + 2);
    range.setBackground(color);
  }

  // Store the coordinates in the global wordLocations object
  wordLocations[word] = wordCoords;
}

function colorCells(coords) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var color = availableColors[Math.floor(Math.random() * availableColors.length)];

  coords.forEach(function (coord) {
    sheet.getRange(coord.row, coord.col).setBackground(color);
  });
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

/*
// This function prints the current sheet to PDF format.
// It exports the sheet as a PDF file and saves it to Google Drive.
// The PDF file is then opened in a new Chrome tab.
function printToPDF() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var sheetName = sheet.getName();
  var range = 'A1:U40';
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
*/




