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
