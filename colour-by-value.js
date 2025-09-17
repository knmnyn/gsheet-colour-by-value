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

/**
 * Applies comprehensive formatting style based on cell value
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1")
 * @param {*} value - The value to set in the cell
 * @param {Object} styleRules - Object containing style rules
 * @param {string} styleRules.background - Background colour (hex code or colour name)
 * @param {string} styleRules.font - Font colour (hex code or colour name)
 * @param {boolean} styleRules.bold - Whether text should be bold
 * @param {boolean} styleRules.italic - Whether text should be italic
 * @param {boolean} styleRules.underline - Whether text should be underlined
 * @param {boolean} styleRules.strikethrough - Whether text should be struck through
 * @param {string} styleRules.fontSize - Font size (e.g., "12", "14")
 * @param {string} styleRules.fontFamily - Font family (e.g., "Arial", "Times New Roman")
 * @param {string} styleRules.horizontalAlignment - Horizontal alignment ("left", "center", "right")
 * @param {string} styleRules.verticalAlignment - Vertical alignment ("top", "middle", "bottom")
 */
function formatCellByValue(sheetName, cellAddress, value, styleRules) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const cell = sheet.getRange(cellAddress);
  
  // Set value if provided
  if (value !== undefined && value !== null) {
    const currentValue = cell.getValue();
    if (currentValue !== value) {
      cell.setValue(value);
    }
  }
  
  // Apply background colour
  if (styleRules.background) {
    cell.setBackground(styleRules.background);
  }
  
  // Apply font colour
  if (styleRules.font) {
    cell.setFontColor(styleRules.font);
  }
  
  // Apply font weight (bold)
  if (styleRules.bold !== undefined) {
    cell.setFontWeight(styleRules.bold ? "bold" : "normal");
  }
  
  // Apply font style (italic)
  if (styleRules.italic !== undefined) {
    cell.setFontStyle(styleRules.italic ? "italic" : "normal");
  }
  
  // Apply underline
  if (styleRules.underline !== undefined) {
    cell.setFontLine(styleRules.underline ? "underline" : "none");
  }
  
  // Apply strikethrough
  if (styleRules.strikethrough !== undefined) {
    cell.setFontLine(styleRules.strikethrough ? "line-through" : "none");
  }
  
  // Apply font size
  if (styleRules.fontSize) {
    cell.setFontSize(parseInt(styleRules.fontSize));
  }
  
  // Apply font family
  if (styleRules.fontFamily) {
    cell.setFontFamily(styleRules.fontFamily);
  }
  
  // Apply horizontal alignment
  if (styleRules.horizontalAlignment) {
    cell.setHorizontalAlignment(styleRules.horizontalAlignment);
  }
  
  // Apply vertical alignment
  if (styleRules.verticalAlignment) {
    cell.setVerticalAlignment(styleRules.verticalAlignment);
  }
}

/**
 * Custom function to apply comprehensive formatting from Google Sheets (formatting only)
 * Usage: =formatCell("Sheet1", "A1", "Error", "#ff0000", "#ffffff", true, false, true)
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address
 * @param {string} value - The value to set
 * @param {string} backgroundColour - Background colour
 * @param {string} fontColour - Font colour
 * @param {boolean} bold - Whether to make text bold
 * @param {boolean} italic - Whether to make text italic
 * @param {boolean} underline - Whether to underline text
 * @returns {string} Success message
 */
function formatCell(sheetName, cellAddress, value, backgroundColour, fontColour, bold, italic, underline) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return "Error: Sheet not found";
    }
    
    const cell = sheet.getRange(cellAddress);
    
    // Apply formatting only (no setValue to avoid permission issues)
    if (backgroundColour) {
      cell.setBackground(backgroundColour);
    }
    
    if (fontColour) {
      cell.setFontColor(fontColour);
    }
    
    if (bold !== undefined) {
      cell.setFontWeight(bold ? "bold" : "normal");
    }
    
    if (italic !== undefined) {
      cell.setFontStyle(italic ? "italic" : "normal");
    }
    
    if (underline !== undefined) {
      cell.setFontLine(underline ? "underline" : "none");
    }
    
    return "Cell formatted successfully";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom function for status-based formatting (formatting only)
 * Usage: =formatByStatus("Sheet1", "A1", "Complete")
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address
 * @param {string} status - The status value
 * @returns {string} Success message
 */
function formatByStatus(sheetName, cellAddress, status) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return "Error: Sheet not found";
    }
    
    const cell = sheet.getRange(cellAddress);
    let backgroundColour, fontColour, bold, italic, underline;
    
    switch (status.toLowerCase()) {
      case "complete":
      case "done":
      case "finished":
        backgroundColour = "#d4edda";
        fontColour = "#155724";
        bold = true;
        italic = false;
        break;
      case "in progress":
      case "pending":
      case "working":
        backgroundColour = "#fff3cd";
        fontColour = "#856404";
        bold = false;
        italic = true;
        break;
      case "error":
      case "failed":
      case "issue":
        backgroundColour = "#f8d7da";
        fontColour = "#721c24";
        bold = true;
        underline = true;
        break;
      case "warning":
      case "caution":
        backgroundColour = "#ffeaa7";
        fontColour = "#6c5ce7";
        bold = true;
        italic = true;
        break;
      default:
        backgroundColour = "#f8f9fa";
        fontColour = "#495057";
        bold = false;
        italic = false;
    }
    
    // Apply formatting only (no setValue to avoid permission issues)
    cell.setBackground(backgroundColour);
    cell.setFontColor(fontColour);
    
    if (bold !== undefined) {
      cell.setFontWeight(bold ? "bold" : "normal");
    }
    
    if (italic !== undefined) {
      cell.setFontStyle(italic ? "italic" : "normal");
    }
    
    if (underline !== undefined) {
      cell.setFontLine(underline ? "underline" : "none");
    }
    
    return "Status formatted successfully";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Generates a consistent random colour based on a string value
 * @param {string} value - The value to generate a colour for
 * @param {string} type - "background" or "foreground"
 * @returns {string} Hex colour code
 */
function generateColourFromValue(value, type = "background") {
  // Convert value to string and create a hash
  const str = value.toString();
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const char = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  
  // Use absolute value to ensure positive
  hash = Math.abs(hash);
  
  if (type === "background") {
    // Generate bright, vibrant background colours
    const hue = hash % 360;
    const saturation = 70 + (hash % 30); // 70-100% saturation
    const lightness = 45 + (hash % 20); // 45-65% lightness for good contrast
    return `hsl(${hue}, ${saturation}%, ${lightness}%)`;
  } else {
    // Generate contrasting foreground colours
    const hue = (hash + 180) % 360; // Opposite hue for contrast
    const saturation = 80 + (hash % 20); // 80-100% saturation
    const lightness = hash % 2 === 0 ? 15 : 85; // Very dark or very light
    return `hsl(${hue}, ${saturation}%, ${lightness}%)`;
  }
}

/**
 * Applies random high contrast colours based on cell value
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address (e.g., "A1")
 * @param {*} value - The value to set in the cell
 * @param {Object} options - Optional styling options
 * @param {boolean} options.bold - Whether to make text bold
 * @param {boolean} options.italic - Whether to make text italic
 * @param {string} options.fontSize - Font size
 */
function colourCellByValueRandom(sheetName, cellAddress, value, options = {}) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const cell = sheet.getRange(cellAddress);
  
  // Set value if provided
  if (value !== undefined && value !== null) {
    const currentValue = cell.getValue();
    if (currentValue !== value) {
      cell.setValue(value);
    }
  }
  
  // Generate consistent colours based on value
  const backgroundColour = generateColourFromValue(value, "background");
  const foregroundColour = generateColourFromValue(value, "foreground");
  
  // Apply colours
  cell.setBackground(backgroundColour);
  cell.setFontColor(foregroundColour);
  
  // Apply additional styling options
  if (options.bold !== undefined) {
    cell.setFontWeight(options.bold ? "bold" : "normal");
  }
  
  if (options.italic !== undefined) {
    cell.setFontStyle(options.italic ? "italic" : "normal");
  }
  
  if (options.fontSize) {
    cell.setFontSize(parseInt(options.fontSize));
  }
}

/**
 * Applies random high contrast colours to a range based on cell values
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to format (e.g., "A1:A10")
 * @param {Object} options - Optional styling options
 */
function colourRangeByValueRandom(sheetName, range, options = {}) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
  
  const rangeObj = sheet.getRange(range);
  const values = rangeObj.getValues();
  
  values.forEach((row, rowIndex) => {
    row.forEach((cellValue, colIndex) => {
      if (cellValue !== null && cellValue !== undefined && cellValue !== "") {
        const cellAddress = rangeObj.getCell(rowIndex + 1, colIndex + 1).getA1Notation();
        colourCellByValueRandom(sheetName, cellAddress, cellValue, options);
      }
    });
  });
}

/**
 * Custom function for random high contrast colouring from Google Sheets (formatting only)
 * Usage: =colourCellRandom("Sheet1", "A1", "Error")
 * @param {string} sheetName - The name of the sheet
 * @param {string} cellAddress - The cell address
 * @param {string} value - The value to set
 * @param {boolean} bold - Whether to make text bold
 * @returns {string} Success message
 */
function colourCellRandom(sheetName, cellAddress, value, bold = true) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return "Error: Sheet not found";
    }
    
    const cell = sheet.getRange(cellAddress);
    
    // Generate consistent colours based on value
    const backgroundColour = generateColourFromValue(value, "background");
    const foregroundColour = generateColourFromValue(value, "foreground");
    
    // Apply colours only (no setValue to avoid permission issues)
    cell.setBackground(backgroundColour);
    cell.setFontColor(foregroundColour);
    
    // Apply additional styling options
    if (bold) {
      cell.setFontWeight("bold");
    }
    
    return "Cell coloured with random high contrast colours";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom function for random high contrast colouring of a range
 * Usage: =colourRangeRandom("Sheet1", "A1:A10")
 * @param {string} sheetName - The name of the sheet
 * @param {string} range - The range to colour
 * @param {boolean} bold - Whether to make text bold
 * @returns {string} Success message
 */
function colourRangeRandom(sheetName, range, bold = true) {
  try {
    const options = { bold: bold };
    colourRangeByValueRandom(sheetName, range, options);
    return "Range coloured with random high contrast colours";
  } catch (error) {
    return "Error: " + error.message;
  }
}

/**
 * Custom formula for conditional formatting with random high contrast colours
 * Returns TRUE if the cell has a value (for conditional formatting)
 * Usage in conditional formatting: =hasValueRandom(A1)
 * @param {*} cellValue - The value of the cell to check
 * @returns {boolean} TRUE if cell has a value
 */
function hasValueRandom(cellValue) {
  return cellValue !== null && cellValue !== undefined && cellValue !== "";
}

/**
 * Custom formula for specific value random colouring
 * Usage in conditional formatting: =valueRandomColour(A1, "Error")
 * @param {*} cellValue - The value of the cell to check
 * @param {string} targetValue - The value to match
 * @returns {boolean} TRUE if cell value matches target
 */
function valueRandomColour(cellValue, targetValue) {
  return cellValue !== null && cellValue !== undefined && 
         cellValue.toString().toLowerCase() === targetValue.toLowerCase();
}
