/* eslint no-var: 0 */
/* exported onOpen */
/**
 *
 *
 */
function onOpen() {
  addStrimanMenu();
}

/* exported onEdit */
/**
 *
 *
 */
function onEdit() {
  insertInTodo();
}

/**
 * Generate unique identifiers (UID) in the sheet.
 *
 */
function generateUID() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Data');
    var columnHeaderValues = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues();
    var UID_COLUMN = columnHeaderValues[0].indexOf('UID') + 1;
    var lastRow = sheet.getLastRow();
    var uidRange = sheet.getRange(2, UID_COLUMN, lastRow - 1);
    var uidValues = uidRange.getValues();
    var uidCounterRange = spreadsheet.getRangeByName('UID_Counter');
    var nextCount = uidCounterRange.getValue();

    for (var i = 0; i < uidValues; i++) {
      if (uidValues[i][0] == 0) {
        uidValues[i][0] = nextCount++;
      }
    }
    uidCounterRange.setValue(nextCount);
    uidRange.setValues(uidValues).setHorizontalAlignment('center');
  } catch (error) {
    Logger.log('%s : %s', error.name, error.message);
  }
}

/**
 * Inserts comment from this spreadsheet into the
 * corresponding TODO spreadsheet.
 *
 *
 */
function insertInTodo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var ADD_TO_TODO2019_COLUMN = columnHeaderValues[0].indexOf('TODO 2019') + 1;
  var currentCell = activeSheet.getCurrentCell();
  var firstRow = currentCell.getRow();

  if ((firstRow == 1) ||
    (currentCell.getColumn() !== ADD_TO_TODO2019_COLUMN) ||
    (activeSheet.getName() !== 'Data') || (currentCell.getValue() == false)) {
    return;
  }

  var UID_COLUMN = columnHeaderValues[0].indexOf('UID') + 1;
  var ACKNOWLEDGE_COLUMN = columnHeaderValues[0].indexOf('Vu') + 1;
  var COMMENT_COLUMN = columnHeaderValues[0].indexOf('Commentaire') + 1;
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;
  var AUTHOR_COLUMN = columnHeaderValues[0].indexOf('Auteur') + 1;
  var todoSpreadsheet = SpreadsheetApp
    .openById('1J1oFUXOXRSxIpiyLbbA2Nwq31mnSzkFXr6zJ5nOduNE');
  var todoDataSheet = todoSpreadsheet.getSheetByName('Data');
  var todoLastColumn = todoDataSheet.getLastColumn();
  var todoColumnHeaderValues = todoDataSheet
    .getRange(1, 1, 1, todoLastColumn).getValues();
  var TODO_TASK_COLUMN = todoColumnHeaderValues[0].indexOf('Tâche') + 1;
  var TODO_CATEGORY_COLUMN = todoColumnHeaderValues[0]
    .indexOf('Catégorie') + 1;
  var TODO_AUTHOR_COLUMN = todoColumnHeaderValues[0].indexOf('Auteur') + 1;
  var TODO_SOURCE_COLUMN = todoColumnHeaderValues[0].indexOf('Source') + 1;
  var valuesToAdd = [];
  var uidValue = activeSheet.getRange(currentCell.getRow(), UID_COLUMN)
    .getValue();
  var spreadsheetName = spreadsheet.getName();

  activeSheet.getRange(currentCell.getRow(), ACKNOWLEDGE_COLUMN)
    .setValue(true);

  for (i = 1; i <= todoLastColumn; i++) {
    if (i == TODO_TASK_COLUMN) {
      valuesToAdd.push(activeSheet.getRange(firstRow, COMMENT_COLUMN)
        .getValue());
      continue;
    }
    if (i == TODO_CATEGORY_COLUMN) {
      valuesToAdd.push(activeSheet.getRange(firstRow, CATEGORY_COLUMN)
        .getValue());
      continue;
    }
    if (i == TODO_AUTHOR_COLUMN) {
      valuesToAdd.push(activeSheet.getRange(firstRow, AUTHOR_COLUMN)
        .getValue());
      continue;
    }
    if (i == TODO_SOURCE_COLUMN) {
      valuesToAdd.push(spreadsheetName + ' - UID: ' + uidValue);
      continue;
    }
    valuesToAdd.push('');
  }
  todoDataSheet.appendRow(valuesToAdd);
}

/**
 * Create a custom menu for this S:Triman spreadsheet.
 *
 */
function addStrimanMenu() {
  SpreadsheetApp.getUi()
    .createMenu('S:Triman')
    .addItem('Afficher Catégories', 'showDialog')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Options avancées')
      .addItem('Générer UID', 'generateUID')
      .addItem('Insérer NewLine dans Catégorie', 'insertNewLineInCategory')
      .addItem('Supprimer la cache \'Catégories\'', 'removeCachedCategories'))
    .addToUi();

  generateUID();
}

/* exported showDialog */
/**
 * Show a sidebar with a list of check box items.
 * This list comes from a data validation range.
 * This function is called from 'Page.html'.
 *
 */
function showDialog() {
  // var activeSheet = SpreadsheetApp.getActiveSheet();
  // var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
  //   .getLastColumn()).getValues();
  // var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;

  // if (SpreadsheetApp.getCurrentCell().getColumn() != CATEGORY_COLUMN) return;

  var html = HtmlService.createTemplateFromFile('Page').evaluate()
    .setTitle('Liste des catégories');
  SpreadsheetApp.getUi().showSidebar(html);
}


/* exported valid */
/**
 * Returns an array of categories, as defined by a data validation range.
 *
 * @return {array}
 */
var valid = function() {
  try {
    return getCategories();
  } catch (e) {
    return null;
  }
};

/* exported currentSelection */
/**
 * Returns the strings contained in the current cell, in an array and
 * separated by 'new line'.
 *
 * @return {array}
 */
var currentSelection = function() {
  try {
    var arrayOfValues = SpreadsheetApp.getActiveRange().getValue().split('\n');
    return arrayOfValues;
  } catch (e) {
    return null;
  }
};


/* exported fillCell */
/**
 * Sets the value of the current cell with user's sidebar selection.
 *
 * @param {*} e
 */
function fillCell(e) {
  // First, check that the selected cell's column is valid
  var activeSheet = e.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;

  // If not, refresh sidebar. Otherwise fill in selected cell with user choice.
  if (e.getCurrentCell().getColumn() != CATEGORY_COLUMN) {
    showDialog();
  } else {
    var s = [];
    for (var i in e) {
      if (i.substr(0, 2) == 'ch') s.push(e[i]);
    }
    if (s.length) SpreadsheetApp.getActiveRange().setValue(s.join('\n'));
  }
}

/**
 * Returns an array of categories from cached data if available.
 * If not, returns an array of categories from 'Catégorie' named range and
 * put the data in cache for faster future calls.
 *
 * @return {array}
 */
function getCategories() {
  // if (SpreadsheetApp.getActiveRange().getDataValidation() == null)
  //  return null;
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;

  if (SpreadsheetApp.getCurrentCell().getColumn() != CATEGORY_COLUMN) {
    return null;
  }

  var cache = CacheService.getScriptCache();
  var cached = cache.get('categories');
  if (cached != null) {
    var newArray1D = cached.split(',');
    var newArray2D = [];
    while (newArray1D.length) newArray2D.push(newArray1D.splice(0, 1));
    Logger.log('newArray2D = %s', newArray2D);
    return newArray2D;
  }
  // var categoryArray = SpreadsheetApp.getActiveRange().getDataValidation()
  // .getCriteriaValues()[0].getValues();
  var categoryArray = SpreadsheetApp.getActiveSpreadsheet()
    .getRangeByName('Catégories').getValues();
  cache.put('categories', categoryArray, 1500);
  Logger.log('categoryArray = %s', categoryArray);
  return categoryArray;
}

/* exported removeCachedCategories */
/**
 * Remove categories from cache.
 *
 */
function removeCachedCategories() {
  var cache = CacheService.getScriptCache();
  cache.remove('categories');
}

/* exported showCachedCategories */
/**
 *
 *
 */
function showCachedCategories() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('categories');
  if (cached != null) {
    Logger.log('Catégories = %s', cached);
  } else {
    Logger.log('No categories in cache');
  }
}

/* exported insertNewLineInCategory */
/**
 * Insert 'new line' after each string of the cells of a column 'Catégorie'..
 *
 */
function insertNewLineInCategory() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var columnHeaderValues = activeSheet.getRange(1, 1, 1, activeSheet
    .getLastColumn()).getValues();
  var CATEGORY_COLUMN = columnHeaderValues[0].indexOf('Catégorie') + 1;
  var lastRow = activeSheet.getLastRow();
  var categoryRange = activeSheet.getRange(2, CATEGORY_COLUMN, lastRow - 1);
  var categoryValues = categoryRange.getValues();

  for (var i = 0; i < categoryValues.length; i++) {
    categoryValues[i][0] = categoryValues[i][0].toString()
      .replace(/, /g, '\n');
  }
  categoryRange.setValues(categoryValues);
}
