//-----------------------------------------------------------------------------------------------------------------
function generateUID() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = spreadsheet.getSheetByName("Data");   
  var columnHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  var UID_COLUMN = columnHeaderValues[0].indexOf("UID") + 1;
  var lastRow     = sheet.getLastRow();
  var uidRange    = sheet.getRange(2, UID_COLUMN, lastRow-1);
  var uidValues   = uidRange.getValues();
  var uidCounterRange = spreadsheet.getRangeByName('UID_Counter');
  var nextCount = uidCounterRange.getValue();

  for (var row in uidValues) {
    if (uidValues[row][0] == 0) {
      uidValues[row][0] = nextCount++;
    }
  }
  
  uidCounterRange.setValue(nextCount);
  uidRange.setValues(uidValues).setHorizontalAlignment("center");
}


//-----------------------------------------------------------------------------------------------------------------
function insertInTodo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues();
  var ADD_TO_TODO2019_COLUMN = columnHeaderValues[0].indexOf("TODO 2019") + 1;
  var currentCell = activeSheet.getCurrentCell();

  if ((firstRow == 1) || (currentCell.getColumn() !== ADD_TO_TODO2019_COLUMN) || (activeSheet.getName() !== "Data") || (currentCell.getValue() == false)) {return};

  var activeRange = currentCell;
  var UID_COLUMN = columnHeaderValues[0].indexOf("UID") + 1;
  var ACKNOWLEDGE_COLUMN = columnHeaderValues[0].indexOf("Vu") + 1;
  var COMMENT_COLUMN = columnHeaderValues[0].indexOf("Commentaire") + 1;
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf("Catégorie") + 1;
  var AUTHOR_COLUMN = columnHeaderValues[0].indexOf("Auteur") + 1;
  var firstRow = currentCell.getRow();
  var todoSpreadsheet = SpreadsheetApp.openById("1J1oFUXOXRSxIpiyLbbA2Nwq31mnSzkFXr6zJ5nOduNE");
  var todoDataSheet = todoSpreadsheet.getSheetByName("Data");
  var todoLastColumn = todoDataSheet.getLastColumn();
  var todoColumnHeaderValues = todoDataSheet.getRange(1, 1, 1, todoLastColumn).getValues();
  var TODO_TASK_COLUMN = todoColumnHeaderValues[0].indexOf("Tâche") + 1;
  var TODO_CATEGORY_COLUMN = todoColumnHeaderValues[0].indexOf("Catégorie") + 1;
  var TODO_AUTHOR_COLUMN = todoColumnHeaderValues[0].indexOf("Auteur") + 1;
  var valuesToAdd = [];
  var uidValue = activeSheet.getRange(currentCell.getRow(), UID_COLUMN).getValue();
  
  activeSheet.getRange(currentCell.getRow(), ACKNOWLEDGE_COLUMN).setValue(true);  
  
  for (i = 1; i <= todoLastColumn; i++) {
    if (i == TODO_TASK_COLUMN) {
      valuesToAdd.push(activeSheet.getRange(firstRow, COMMENT_COLUMN).getValue());
      continue;
    }
    if (i == TODO_CATEGORY_COLUMN) {
      valuesToAdd.push(activeSheet.getRange(firstRow, CATEGORY_COLUMN).getValue());
      continue;
    }
    if (i == TODO_AUTHOR_COLUMN) {
      valuesToAdd.push(activeSheet.getRange(firstRow, AUTHOR_COLUMN).getValue());
      continue;
    }
    valuesToAdd.push("");
  }
  todoDataSheet.appendRow(valuesToAdd);
}

//-----------------------------------------------------------------------------------------------------------------
//function addMenu() {
//  SpreadsheetApp.getUi()
//  .createMenu('S:Triman')
//  .addItem('Insérer NewLine dans Catégorie', 'insertNewLineInCategory')
//  .addToUi();
//}


////-----------------------------------------------------------------------------------------------------------------
//function insertNewLineInCategory() {
//  var spreadsheet        = SpreadsheetApp.getActiveSpreadsheet();
//  var activeSheet        = SpreadsheetApp.getActiveSheet();
//  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues();
//  var CATEGORY_COLUMN    = columnHeaderValues[0].indexOf("Catégorie") + 1;
//  var lastRow            = activeSheet.getLastRow();
//  var categoryRange      = activeSheet.getRange(2, CATEGORY_COLUMN, lastRow-1);
//  var categoryValues     = categoryRange.getValues();
//  var categoryArray      = [];
//
//  for (var row in categoryValues) {
//    categoryValues[row][0] = categoryValues[row][0].toString().replace(/, /g, "\n");
//  }
//  
//  categoryRange.setValues(categoryValues);
//}