/**
 * Creates the custom menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Cell Formatting')
    .addItem('Random Colour', 'randomColour')
    .addItem('Random Format', 'randomFormat')
    .addToUi();
}

/**
 * Applies colors based on the cell value to the selected cell(s)
 */
function randomColour() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  // Get cell values
  const values = range.getValues();
  
  // Apply colors based on each cell's value
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cellValue = values[i][j];
      const cell = range.getCell(i + 1, j + 1);
      
      // Calculate colors based on cell value
      const backgroundColor = getColorFromValue(cellValue);
      const textColor = getContrastColor(backgroundColor);
      
      // Apply colors to the cell
      cell.setBackground(backgroundColor);
      cell.setFontColor(textColor);      
    }
  }
}

/**
 * Applies colors and random text formatting based on the cell value to the selected cell(s)
 */
function randomFormat() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  // First apply colors using the existing randomColour function
  randomColour();
  
  // Get cell values for formatting
  const values = range.getValues();
  
  // Apply random text formatting based on each cell's value
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
      const cellValue = values[i][j];
      const cell = range.getCell(i + 1, j + 1);
      
      // Apply random text formatting based on cell value
      const formatting = getRandomFormatting(cellValue);
      cell.setFontWeight(formatting.bold ? 'bold' : 'normal');
      cell.setFontStyle(formatting.italic ? 'italic' : 'normal');
      cell.setFontLine(formatting.underline ? 'underline' : 'none');
      
      // Store format info for summary
      const formatDesc = [];
      if (formatting.bold) formatDesc.push('bold');
      if (formatting.italic) formatDesc.push('italic');
      if (formatting.underline) formatDesc.push('underline');
      const formatString = formatDesc.length > 0 ? ` (${formatDesc.join(', ')})` : '';
      
      // Get the background color for display
      const backgroundColor = getColorFromValue(cellValue);
    }
  }
}

/**
 * Generates a color based on the cell value
 */
function getColorFromValue(value) {
  if (value === null || value === undefined || value === '') {
    return '#FFFFFF'; // White for empty cells
  }
  
  // Use MD5 hash to pick from websafe pastel colors
  const str = String(value);
  const pastelColors = [
    '#FFB3BA', '#BAFFC9', '#BAE1FF', '#FFFFBA', '#FFB3FF', '#B3FFF9',
    '#FFD1BA', '#E5FFBA', '#BAC3FF', '#FFBAD4', '#BAFFCE', '#BAF0FF',
    '#FFF7BA', '#FFB3F1', '#B3FFE9', '#FFBAC9', '#D4FFBA', '#BAD6FF',
    '#FFE3BA', '#F1B3FF', '#B3FFF1', '#FFBABE', '#C3FFBA', '#BAFFFF',
    '#FFDEBA', '#E1B3FF', '#B3FFD6', '#FFC9BA', '#B3FFBA', '#BAE8FF',
    '#FFCEBA', '#D1B3FF'
  ];
  
  // Get MD5 hash and use first 5 bits (0-31) to select color
  const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  const colorIndex = hashBytes[0] & 0x1F; // Take first byte & mask to get 0-31
  
  // Return the selected pastel color
  return pastelColors[colorIndex];
}

/**
 * Generates a contrasting text color for the given background color
 */
function getContrastColor(backgroundColor) {
  // Array of 16 dark shades for text colors
  const darkShades = [
    '#1A1A1A', '#2B2B2B', '#3C3C3C', '#4D4D4D',
    '#262626', '#333333', '#404040', '#595959',
    '#1F1F1F', '#2E2E2E', '#3D3D3D', '#4C4C4C',
    '#232323', '#303030', '#3E3E3E', '#4B4B4B'
  ];

  // Get MD5 hash of the background color
  const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, backgroundColor);
  
  // Use first 4 bits (0-15) to select dark shade
  const colorIndex = hashBytes[0] & 0x0F;
  
  return darkShades[colorIndex];
}

/**
 * Generates random text formatting based on the cell value
 */
function getRandomFormatting(value) {
  if (value === null || value === undefined || value === '') {
    return { bold: false, italic: false, underline: false };
  }
  
  // Use MD5 hash to determine formatting
  const str = String(value);
  const hashBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  
  // Use different bytes to determine each formatting property
  // This ensures consistent formatting for the same value
  const boldByte = hashBytes[1] & 0xFF;
  const italicByte = hashBytes[2] & 0xFF;
  const underlineByte = hashBytes[3] & 0xFF;
  
  // Apply formatting with different probabilities
  // Bold: 40% chance (0-101 out of 255)
  // Italic: 30% chance (0-76 out of 255)  
  // Underline: 25% chance (0-63 out of 255)
  return {
    bold: boldByte < 102,
    italic: italicByte < 77,
    underline: underlineByte < 64
  };
}
