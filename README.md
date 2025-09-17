# gsheet-colour-cell

A Google Apps Script utility that automatically applies colors and formatting to spreadsheet cells based on their values. The script uses MD5 hashing to ensure consistent coloring - the same value will always get the same color and formatting.

## Features

- **Consistent Color Mapping**: Same cell values always get the same colors using MD5 hashing
- **Random Color Generation**: Uses a palette of 32 web-safe pastel colors
- **Automatic Text Contrast**: Generates contrasting text colors for readability
- **Random Text Formatting**: Applies bold, italic, and underline formatting based on cell values
- **Custom Menu Integration**: Easy-to-use menu in Google Sheets
- **Range Support**: Works on single cells or multiple selected cells

## Setup

### 1. Google Apps Script Setup

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy the contents of `colour-by-value.js` into your script editor
4. Save the project with a descriptive name (e.g., "Cell Color Formatter")

### 2. Link to Google Sheets

1. Open your Google Sheets document
2. Go to **Extensions** → **Apps Script**
3. If you created the script separately, you'll need to copy the code into the Apps Script editor for your specific sheet
4. Save the script

### 3. Grant Permissions

When you first run any function, Google Apps Script will prompt you to grant necessary permissions to access your Google Sheets. Click "Review permissions" and authorize the script.

### 4. Custom Menu

After saving the script, refresh your Google Sheets document. You should see a new **"Cell Formatting"** menu in the menu bar with two options:
- **Random Colour** - Applies colors based on cell values
- **Random Format** - Applies colors and text formatting based on cell values

## Usage

### Using the Custom Menu (Recommended)

1. **Select cells** in your Google Sheets document that you want to format
2. **Click on "Cell Formatting"** in the menu bar
3. **Choose an option**:
   - **Random Colour**: Applies background colors and contrasting text colors based on cell values
   - **Random Format**: Applies colors plus random text formatting (bold, italic, underline) based on cell values

### How It Works

The script uses MD5 hashing to ensure **consistent results**:
- The same cell value will always get the same color and formatting
- Different values get different colors from a palette of 32 pastel colors
- Text colors are automatically chosen for optimal contrast
- Empty cells remain white

### Example Results

| Cell Value | Background Color | Text Color | Formatting |
|------------|------------------|------------|------------|
| "Apple"    | Light Pink       | Dark Gray  | Bold       |
| "Banana"   | Light Green      | Dark Gray  | Normal     |
| "Cherry"   | Light Blue       | Dark Gray  | Italic     |
| "Apple"    | Light Pink       | Dark Gray  | Bold       |

*Note: "Apple" always gets the same formatting because the script uses consistent hashing*

### Manual Function Calls

You can also call the functions directly from the Apps Script editor:

```javascript
// Apply colors to selected cells
randomColour();

// Apply colors and formatting to selected cells  
randomFormat();
```

## Available Functions

### Main Functions

- **`onOpen()`** - Creates the custom "Cell Formatting" menu when the spreadsheet opens
- **`randomColour()`** - Applies background colors and contrasting text colors to selected cells based on their values
- **`randomFormat()`** - Applies colors plus random text formatting (bold, italic, underline) to selected cells based on their values

### Helper Functions

- **`getColorFromValue(value)`** - Generates a consistent color from a cell value using MD5 hashing
- **`getContrastColor(backgroundColor)`** - Generates a contrasting text color for optimal readability
- **`getRandomFormatting(value)`** - Generates random text formatting (bold, italic, underline) based on cell value using MD5 hashing

## How the Color System Works

### Color Palette
The script uses a carefully selected palette of 32 web-safe pastel colors:
- Light pinks, greens, blues, yellows, purples, and cyans
- All colors are pastel tones for easy reading
- Colors are chosen using MD5 hash of the cell value

### Text Formatting Probabilities
When using "Random Format", the script applies formatting with these probabilities:
- **Bold**: 40% chance
- **Italic**: 30% chance  
- **Underline**: 25% chance

### Consistency Guarantee
- Same input value = Same output formatting
- Different values = Different formatting
- Empty cells remain unformatted (white background)

## Common Use Cases

### Data Categorization
Perfect for visually grouping similar data:
- Product categories get consistent colors
- Status values are easily distinguishable
- Duplicate entries are immediately visible

### Data Validation
Use colors to highlight data patterns:
- Duplicate values get the same color
- Unique values get unique colors
- Empty cells stay white for easy identification

### Visual Organization
Great for organizing spreadsheets:
- Color-code by department, priority, or type
- Make data more visually appealing
- Improve readability of large datasets

## Best Practices

1. **Test on a copy first** - Always test on sample data before applying to important sheets
2. **Select appropriate ranges** - Choose the exact cells you want to format
3. **Use consistent data** - Ensure your data is clean and consistent for best results
4. **Refresh after changes** - If you modify the script, refresh your Google Sheets document
5. **Backup your data** - Keep backups of important spreadsheets before applying formatting

## Troubleshooting

- **Menu not appearing**: Refresh your Google Sheets document after saving the script
- **No colors applied**: Make sure you have cells selected before running the function
- **Permission errors**: Grant necessary permissions when prompted by Google Apps Script
- **Script not running**: Check the Apps Script editor for any error messages
- **Inconsistent results**: Ensure your cell values are exactly the same (including spaces and case)

## License

This project is licensed under the terms specified in the LICENSE file.
