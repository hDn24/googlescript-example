/**
 * A special function that runs when the spreadsheet is opened
 * or reloaded, used to add a custom menu to the spreadsheet.
 */
function onOpen() {
  // Get the spreadsheet's user-interface object.
  var ui = SpreadsheetApp.getUi();

  // Create and add a named menu and its items to the menu bar.
  ui.createMenu('Quick formats')
    .addItem('Format row header', 'formatRowHeader')
    .addItem('Format column header', 'formatColumnHeader')
    .addItem('Format dataset', 'formatDataset')
    .addToUi();
}

/**
 * Formats top row of sheet using our header row style.
 */
function formatRowHeader() {
  // Get the current active sheet and the top row's range.
  var sheet = SpreadsheetApp.getActiveSheet();
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  // Apply each format to the top row: bold white text,
  // blue-green background, and a solid black border
  // around the cells.
  headerRange
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#007272')
    .setBorder(
      true, true, true, true, null, null,
      null,
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

}
