/**
 * Google Spreadsheet Utility Functions
 * Collection of utility functions for Google Spreadsheet operations
 */

/**
 * Sets the background color and font color of a cell based on its value
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1", "B2")
 * @param {*} value - The value to check
 * @param {Object} colorRules - Object containing color rules
 * @param {string} colorRules.background - Background color (hex code or color name)
 * @param {string} colorRules.font - Font color (hex code or color name)
 */
function colorCellByValue(sheetName, cellAddress, value, colorRules) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const cell = sheet.getRange(cellAddress);
  cell.setValue(value);
  
  if (colorRules.background) {
    cell.setBackground(colorRules.background);
  }
  
  if (colorRules.font) {
    cell.setFontColor(colorRules.font);
  }
}

/**
 * Applies conditional formatting to a range based on value conditions
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to format (e.g., "A1:A10")
 * @param {Array} conditions - Array of condition objects
 * @param {string} conditions[].condition - The condition type (e.g., "NUMBER_GREATER_THAN", "TEXT_CONTAINS")
 * @param {Array} conditions[].values - Array of values to compare against
 * @param {string} conditions[].background - Background color for this condition
 * @param {string} conditions[].font - Font color for this condition
 */
function applyConditionalFormatting(sheetName, range, conditions) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const rangeObj = sheet.getRange(range);
  
  conditions.forEach(condition => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .setRanges([rangeObj])
      .whenConditionSatisfied(condition.condition, condition.values)
      .setBackground(condition.background)
      .setFontColor(condition.font)
      .build();
    
    sheet.setConditionalFormatRules([rule]);
  });
}

/**
 * Gets the current value of a cell
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1")
 * @returns {*} The value of the cell
 */
function getCellValue(sheetName, cellAddress) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  return sheet.getRange(cellAddress).getValue();
}

/**
 * Sets multiple cell values at once
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to set values for (e.g., "A1:B2")
 * @param {Array} values - 2D array of values to set
 */
function setRangeValues(sheetName, range, values) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  sheet.getRange(range).setValues(values);
}

/**
 * Clears formatting from a range
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to clear formatting from
 */
function clearFormatting(sheetName, range) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  sheet.getRange(range).clearFormat();
}

/**
 * Creates a new sheet if it doesn't exist
 * @param {string} sheetName - The name of the sheet to create
 * @returns {Sheet} The created or existing sheet
 */
function createSheetIfNotExists(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Gets all sheet names in the current spreadsheet
 * @returns {Array<string>} Array of sheet names
 */
function getAllSheetNames() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  return sheets.map(sheet => sheet.getName());
}

/**
 * Deletes a sheet by name
 * @param {string} sheetName - The name of the sheet to delete
 */
function deleteSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  } else {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
}
