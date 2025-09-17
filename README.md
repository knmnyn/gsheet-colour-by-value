# gsheet-colour-cell

A collection of utility functions for Google Spreadsheet operations, with a focus on cell coloring and formatting based on values.

## Features

- Color cells based on their values
- Apply conditional formatting to ranges
- Manage sheet operations
- Bulk data operations
- Error handling and validation

## Setup

### 1. Google Apps Script Setup

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy the contents of `colour-by-value.js` into your script editor
4. Save the project

### 2. Grant Permissions

When you first run any function, Google Apps Script will prompt you to grant necessary permissions to access your Google Sheets.

## Usage Examples

### Basic Cell Coloring

```javascript
// Color a single cell based on its value
colorCellByValue("Sheet1", "A1", "Error", {
  background: "#ff0000",  // Red background
  font: "#ffffff"         // White text
});
```

### Conditional Formatting

```javascript
// Apply conditional formatting to a range
const conditions = [
  {
    condition: "NUMBER_GREATER_THAN",
    values: [100],
    background: "#00ff00",  // Green for values > 100
    font: "#000000"
  },
  {
    condition: "NUMBER_LESS_THAN",
    values: [50],
    background: "#ff0000",  // Red for values < 50
    font: "#ffffff"
  }
];

applyConditionalFormatting("Sheet1", "A1:A10", conditions);
```

### Data Operations

```javascript
// Get a cell value
const value = getCellValue("Sheet1", "A1");

// Set multiple values at once
const values = [
  ["Name", "Age", "City"],
  ["John", 25, "New York"],
  ["Jane", 30, "London"]
];
setRangeValues("Sheet1", "A1:C3", values);
```

### Custom Functions

```javascript
function highlightErrors() {
  const sheetName = "Data";
  const dataRange = "A1:A100";
  
  // Get all values in the range
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getRange(dataRange).getValues();
  
  // Check each cell and color if it contains "ERROR"
  values.forEach((row, index) => {
    if (row[0] && row[0].toString().toUpperCase().includes("ERROR")) {
      const cellAddress = `A${index + 1}`;
      colorCellByValue(sheetName, cellAddress, row[0], {
        background: "#ffcccc",
        font: "#cc0000"
      });
    }
  });
}
```

### Sheet Management

```javascript
// Create a new sheet if it doesn't exist
const newSheet = createSheetIfNotExists("Analysis");

// Get all sheet names
const allSheets = getAllSheetNames();

// Clear formatting from a range
clearFormatting("Sheet1", "A1:A10");
```

## Available Functions

### Core Functions

- **`colorCellByValue(sheetName, cellAddress, value, colorRules)`** - Sets background and font colors based on cell values
- **`applyConditionalFormatting(sheetName, range, conditions)`** - Applies conditional formatting rules to ranges
- **`getCellValue(sheetName, cellAddress)`** - Retrieves values from specific cells
- **`setRangeValues(sheetName, range, values)`** - Sets multiple cell values at once
- **`clearFormatting(sheetName, range)`** - Removes formatting from ranges

### Sheet Management

- **`createSheetIfNotExists(sheetName)`** - Creates sheets if they don't exist
- **`getAllSheetNames()`** - Lists all sheet names
- **`deleteSheet(sheetName)`** - Removes sheets by name

## Running Functions

### From Google Apps Script Editor
1. Select a function from the dropdown
2. Click the "Run" button
3. Grant necessary permissions when prompted

### From Google Sheets
1. Create custom functions in your script
2. Use them in cells like: `=myCustomFunction()`

### As Triggers
```javascript
// Set up a trigger to run when the sheet is edited
function setupTrigger() {
  ScriptApp.newTrigger('highlightErrors')
    .timeBased()
    .everyMinutes(5)
    .create();
}
```

## Common Use Cases

### Data Validation Highlighting
```javascript
function highlightInvalidData() {
  const sheetName = "Input";
  const range = "B2:B20";
  
  const conditions = [
    {
      condition: "NUMBER_LESS_THAN",
      values: [0],
      background: "#ffcccc",
      font: "#cc0000"
    }
  ];
  
  applyConditionalFormatting(sheetName, range, conditions);
}
```

### Status-Based Coloring
```javascript
function colorByStatus() {
  const statuses = ["Complete", "In Progress", "Pending"];
  const colors = ["#d4edda", "#fff3cd", "#f8d7da"];
  
  // Apply colors based on status values
  // Implementation depends on your specific data structure
}
```

## Best Practices

1. **Always test functions** on a copy of your data first
2. **Use descriptive sheet and range names** for clarity
3. **Handle errors gracefully** - the functions include basic error handling
4. **Batch operations** when possible to improve performance
5. **Use triggers sparingly** to avoid hitting execution limits

## Troubleshooting

- **"Sheet not found" errors**: Check that sheet names match exactly (case-sensitive)
- **Permission errors**: Make sure you've granted necessary permissions
- **Execution timeouts**: Break large operations into smaller chunks
- **Format not applying**: Check that color values are valid (hex codes or standard color names)

## License

This project is licensed under the terms specified in the LICENSE file.
