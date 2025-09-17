# 🎨 gsheet-colour-by-value

A Google Apps Script utility that automatically applies colours and formatting to spreadsheet cells based on their values. The script uses MD5 hashing to ensure consistent colouring - the same value will always get the same colour and formatting.

## ✨ Features

- 🎯 **Consistent Colour Mapping**: Same cell values always get the same colours using MD5 hashing
- 🌈 **Random Colour Generation**: Uses a palette of 32 web-safe pastel colours
- 🔍 **Automatic Text Contrast**: Generates contrasting text colours for readability
- ✏️ **Random Text Formatting**: Applies bold, italic, and underline formatting based on cell values
- 📋 **Custom Menu Integration**: Easy-to-use menu in Google Sheets
- 📊 **Range Support**: Works on single cells or multiple selected cells

## 🚀 Setup

### 1. 📝 Google Apps Script Setup

1. Go to [script.google.com](https://script.google.com)
2. Create a new project
3. Copy the contents of `colour-by-value.js` into your script editor
4. Save the project with a descriptive name (e.g., "Cell Colour Formatter")

### 2. 🔗 Link to Google Sheets

1. Open your Google Sheets document
2. Go to **Extensions** → **Apps Script**
3. If you created the script separately, you'll need to copy the code into the Apps Script editor for your specific sheet
4. Save the script

### 3. 🔐 Grant Permissions

When you first run any function, Google Apps Script will prompt you to grant necessary permissions to access your Google Sheets. Click "Review permissions" and authorize the script.

### 4. 📋 Custom Menu

After saving the script, refresh your Google Sheets document. You should see a new **"Cell Formatting"** menu in the menu bar with two options:
- **Random Colour** - Applies colors based on cell values
- **Random Format** - Applies colors and text formatting based on cell values

## 📖 Usage

### 🖱️ Using the Custom Menu (Recommended)

1. 📍 **Select cells** in your Google Sheets document that you want to format
2. 🖱️ **Click on "Cell Formatting"** in the menu bar
3. ⚙️ **Choose an option**:
   - 🎨 **Random Colour**: Applies background colours and contrasting text colours based on cell values
   - ✨ **Random Format**: Applies colours plus random text formatting (bold, italic, underline) based on cell values

### ⚙️ How It Works

The script uses MD5 hashing to ensure **consistent results**:
- 🎯 The same cell value will always get the same colour and formatting
- 🌈 Different values get different colours from a palette of 32 pastel colours
- 🔍 Text colours are automatically chosen for optimal contrast
- ⚪ Empty cells remain white

### 📊 Example Results

| Cell Value | Background Colour | Text Colour | Formatting |
|------------|------------------|------------|------------|
| "Apple"    | Light Pink       | Dark Gray  | Bold       |
| "Banana"   | Light Green      | Dark Gray  | Normal     |
| "Cherry"   | Light Blue       | Dark Gray  | Italic     |
| "Apple"    | Light Pink       | Dark Gray  | Bold       |

*Note: "Apple" always gets the same formatting because the script uses consistent hashing*

### 💻 Manual Function Calls

You can also call the functions directly from the Apps Script editor:

```javascript
// Apply colours to selected cells
randomColour();

// Apply colours and formatting to selected cells  
randomFormat();
```

## 🔧 Available Functions

### 🎯 Main Functions

- **`onOpen()`** 📋 - Creates the custom "Cell Formatting" menu when the spreadsheet opens
- **`randomColour()`** 🎨 - Applies background colours and contrasting text colours to selected cells based on their values
- **`randomFormat()`** ✨ - Applies colours plus random text formatting (bold, italic, underline) to selected cells based on their values

### 🛠️ Helper Functions

- **`getColourFromValue(value)`** 🌈 - Generates a consistent colour from a cell value using MD5 hashing
- **`getContrastColour(backgroundColour)`** 🔍 - Generates a contrasting text colour for optimal readability
- **`getRandomFormatting(value)`** ✏️ - Generates random text formatting (bold, italic, underline) based on cell value using MD5 hashing

## 🎨 How the Colour System Works

### 🌈 Colour Palette
The script uses a carefully selected palette of 32 web-safe pastel colours:
- 🎨 Light pinks, greens, blues, yellows, purples, and cyans
- 👁️ All colours are pastel tones for easy reading
- 🔢 Colours are chosen using MD5 hash of the cell value

### 📊 Text Formatting Probabilities
When using "Random Format", the script applies formatting with these probabilities:
- **Bold** 💪: 40% chance
- **Italic** ✨: 30% chance  
- **Underline** 📝: 25% chance

### 🎯 Consistency Guarantee
- 🔄 Same input value = Same output formatting
- 🌈 Different values = Different formatting
- ⚪ Empty cells remain unformatted (white background)

## 💡 Common Use Cases

### 📊 Data Categorization
Perfect for visually grouping similar data:
- 🏷️ Product categories get consistent colours
- 🎯 Status values are easily distinguishable
- 👀 Duplicate entries are immediately visible

### ✅ Data Validation
Use colours to highlight data patterns:
- 🔄 Duplicate values get the same colour
- 🆔 Unique values get unique colours
- ⚪ Empty cells stay white for easy identification

### 🎨 Visual Organization
Great for organizing spreadsheets:
- 🏢 Colour-code by department, priority, or type
- ✨ Make data more visually appealing
- 📖 Improve readability of large datasets

## 💡 Best Practices

1. 🧪 **Test on a copy first** - Always test on sample data before applying to important sheets
2. 📍 **Select appropriate ranges** - Choose the exact cells you want to format
3. 🎯 **Use consistent data** - Ensure your data is clean and consistent for best results
4. 🔄 **Refresh after changes** - If you modify the script, refresh your Google Sheets document
5. 💾 **Backup your data** - Keep backups of important spreadsheets before applying formatting

## 🔧 Troubleshooting

- 📋 **Menu not appearing**: Refresh your Google Sheets document after saving the script
- 🎨 **No colours applied**: Make sure you have cells selected before running the function
- 🔐 **Permission errors**: Grant necessary permissions when prompted by Google Apps Script
- ⚠️ **Script not running**: Check the Apps Script editor for any error messages
- 🔄 **Inconsistent results**: Ensure your cell values are exactly the same (including spaces and case)

## 📄 License

This project is licensed under the terms specified in the LICENSE file.
