/**
 * Google Spreadsheet Utility Functions
 * Collection of utility functions for Google Spreadsheet operations
 */

/**
 * Sets the background colour and font colour of a cell based on its value
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1", "B2")
 * @param {*} value - The value to check
 * @param {Object} colourRules - Object containing colour rules
 * @param {string} colourRules.background - Background colour (hex code or colour name)
 * @param {string} colourRules.font - Font colour (hex code or colour name)
 */
function colourCellByValue(sheetName, cellAddress, value, colourRules) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const cell = sheet.getRange(cellAddress);
  
  // Only set value if it's different from current value
  const currentValue = cell.getValue();
  if (currentValue !== value) {
    cell.setValue(value);
  }
  
  if (colourRules.background) {
    cell.setBackground(colourRules.background);
  }
  
  if (colourRules.font) {
    cell.setFontColor(colourRules.font);
  }
}

/**
 * Applies conditional formatting to a range based on value conditions
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to format (e.g., "A1:A10")
 * @param {Array} conditions - Array of condition objects
 * @param {string} conditions[].condition - The condition type (e.g., "NUMBER_GREATER_THAN", "TEXT_CONTAINS")
 * @param {Array} conditions[].values - Array of values to compare against
 * @param {string} conditions[].background - Background colour for this condition
 * @param {string} conditions[].font - Font colour for this condition
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

/**
 * Custom function that can be called from Google Sheets cells
 * Usage: =colourCell("Sheet1", "A1", "Test", "#ff0000", "#ffffff")
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1")
 * @param {string} value - The value to set in the cell
 * @param {string} backgroundColour - Background colour (hex code or colour name)
 * @param {string} fontColour - Font colour (hex code or colour name)
 * @returns {string} Success message
 */
function colourCell(sheetName, cellAddress, value, backgroundColour, fontColour) {
  try {
    colourCellByValue(sheetName, cellAddress, value, {
      background: backgroundColour,
      font: fontColour
    });
    return "Cell coloured successfully";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom function to only apply formatting (no value setting) - safer for custom functions
 * Usage: =colourCellFormat("Sheet1", "A1", "#ff0000", "#ffffff")
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1")
 * @param {string} backgroundColour - Background colour (hex code or colour name)
 * @param {string} fontColour - Font colour (hex code or colour name)
 * @returns {string} Success message
 */
function colourCellFormat(sheetName, cellAddress, backgroundColour, fontColour) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return "Error: Sheet not found";
    }
    
    const cell = sheet.getRange(cellAddress);
    
    if (backgroundColour) {
      cell.setBackground(backgroundColour);
    }
    
    if (fontColour) {
      cell.setFontColor(fontColour);
    }
    
    return "Formatting applied successfully";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom function to apply conditional formatting from Google Sheets
 * Usage: =applyConditionalFormat("Sheet1", "A1:A10", "NUMBER_GREATER_THAN", 100, "#00ff00", "#000000")
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to format
 * @param {string} condition - The condition type
 * @param {number|string} value - The value to compare against
 * @param {string} backgroundColour - Background colour
 * @param {string} fontColour - Font colour
 * @returns {string} Success message
 */
function applyConditionalFormat(sheetName, range, condition, value, backgroundColour, fontColour) {
  try {
    const conditions = [{
      condition: condition,
      values: [value],
      background: backgroundColour,
      font: fontColour
    }];
    
    applyConditionalFormatting(sheetName, range, conditions);
    return "Conditional formatting applied successfully";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom function to get cell value from Google Sheets
 * Usage: =getValue("Sheet1", "A1")
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address
 * @returns {*} The value of the cell
 */
function getValue(sheetName, cellAddress) {
  try {
    return getCellValue(sheetName, cellAddress);
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom function to clear formatting from Google Sheets
 * Usage: =clearFormat("Sheet1", "A1:A10")
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to clear formatting from
 * @returns {string} Success message
 */
function clearFormat(sheetName, range) {
  try {
    clearFormatting(sheetName, range);
    return "Formatting cleared successfully";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom formula function for conditional formatting
 * Returns TRUE/FALSE based on conditions - can be used in conditional formatting rules
 * Usage in conditional formatting: =colourCondition(A1, "Error", "contains")
 * @param {*} cellValue - The value of the cell to check
 * @param {string} targetValue - The value to compare against
 * @param {string} condition - The condition type ("equals", "contains", "greater", "less", "greaterEqual", "lessEqual")
 * @returns {boolean} TRUE if condition is met, FALSE otherwise
 */
function colourCondition(cellValue, targetValue, condition) {
  try {
    if (cellValue === null || cellValue === undefined) {
      return false;
    }
    
    const cellStr = cellValue.toString().toLowerCase();
    const targetStr = targetValue.toString().toLowerCase();
    
    switch (condition.toLowerCase()) {
      case "equals":
        return cellStr === targetStr;
      case "contains":
        return cellStr.includes(targetStr);
      case "greater":
        return parseFloat(cellValue) > parseFloat(targetValue);
      case "less":
        return parseFloat(cellValue) < parseFloat(targetValue);
      case "greaterequal":
        return parseFloat(cellValue) >= parseFloat(targetValue);
      case "lessequal":
        return parseFloat(cellValue) <= parseFloat(targetValue);
      case "notequals":
        return cellStr !== targetStr;
      case "notcontains":
        return !cellStr.includes(targetStr);
      case "startswith":
        return cellStr.startsWith(targetStr);
      case "endswith":
        return cellStr.endsWith(targetStr);
      default:
        return cellStr === targetStr;
    }
  } catch (error) {
    return false;
  }
}

/**
 * Custom formula for text-based conditional formatting
 * Usage in conditional formatting: =textColourCondition(A1, "Error")
 * @param {*} cellValue - The value of the cell to check
 * @param {string} targetText - The text to look for
 * @returns {boolean} TRUE if cell contains the target text
 */
function textColourCondition(cellValue, targetText) {
  try {
    if (cellValue === null || cellValue === undefined) {
      return false;
    }
    return cellValue.toString().toLowerCase().includes(targetText.toLowerCase());
  } catch (error) {
    return false;
  }
}

/**
 * Custom formula for number-based conditional formatting
 * Usage in conditional formatting: =numberColourCondition(A1, 100, "greater")
 * @param {*} cellValue - The value of the cell to check
 * @param {number} targetNumber - The number to compare against
 * @param {string} operator - The comparison operator ("greater", "less", "equal", "greaterEqual", "lessEqual")
 * @returns {boolean} TRUE if condition is met
 */
function numberColourCondition(cellValue, targetNumber, operator) {
  try {
    const numValue = parseFloat(cellValue);
    if (isNaN(numValue)) {
      return false;
    }
    
    switch (operator.toLowerCase()) {
      case "greater":
        return numValue > targetNumber;
      case "less":
        return numValue < targetNumber;
      case "equal":
        return numValue === targetNumber;
      case "greaterequal":
        return numValue >= targetNumber;
      case "lessequal":
        return numValue <= targetNumber;
      default:
        return numValue > targetNumber;
    }
  } catch (error) {
    return false;
  }
}

/**
 * Custom formula for date-based conditional formatting
 * Usage in conditional formatting: =dateColourCondition(A1, "today", "equals")
 * @param {*} cellValue - The value of the cell to check
 * @param {string} targetDate - The date to compare against ("today", "yesterday", or a specific date)
 * @param {string} operator - The comparison operator ("equals", "greater", "less", "greaterEqual", "lessEqual")
 * @returns {boolean} TRUE if condition is met
 */
function dateColourCondition(cellValue, targetDate, operator) {
  try {
    const cellDate = new Date(cellValue);
    if (isNaN(cellDate.getTime())) {
      return false;
    }
    
    let compareDate;
    if (targetDate.toLowerCase() === "today") {
      compareDate = new Date();
    } else if (targetDate.toLowerCase() === "yesterday") {
      compareDate = new Date();
      compareDate.setDate(compareDate.getDate() - 1);
    } else {
      compareDate = new Date(targetDate);
    }
    
    // Reset time to compare only dates
    cellDate.setHours(0, 0, 0, 0);
    compareDate.setHours(0, 0, 0, 0);
    
    switch (operator.toLowerCase()) {
      case "equals":
        return cellDate.getTime() === compareDate.getTime();
      case "greater":
        return cellDate.getTime() > compareDate.getTime();
      case "less":
        return cellDate.getTime() < compareDate.getTime();
      case "greaterequal":
        return cellDate.getTime() >= compareDate.getTime();
      case "lessequal":
        return cellDate.getTime() <= compareDate.getTime();
      default:
        return cellDate.getTime() === compareDate.getTime();
    }
  } catch (error) {
    return false;
  }
}
