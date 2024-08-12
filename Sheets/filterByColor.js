function filterByColor(sheetName, rangeAddress, color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Hoja "' + sheetName + '" no encontrada.');
  }

  var range = sheet.getRange(rangeAddress);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var result = [];

  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (backgrounds[i][j] == color.toLowerCase()) {
        result.push(values[i][j]);
      }
    }
  }

  return result;
}