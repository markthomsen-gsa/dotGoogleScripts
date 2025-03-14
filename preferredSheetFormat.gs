// Configuration object for formatting options.
const FORMAT_CONFIG = {
  deleteEmptyRows: true,
  deleteEmptyColumns: true,
  setBorders: true,
  borderOptions: { top: true, left: true, bottom: true, right: true, vertical: true, horizontal: true },
  headerBold: true,
  customRowHeight: null,
  autoAdjust: false,
  horizontalAlignment: "left",
  verticalAlignment: "middle",
  freezeFirstRow: true,
  freezeFirstColumn: true,
  banding: true,
  bandingTheme: SpreadsheetApp.BandingTheme.BLUE
};

// Main formatting function - acts as a controller
function formatEntireSheet() {
  const ui = getUi();
  
  if (!confirmFormatting(ui)) {
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startTime = new Date();
  const results = {
    errors: [],
    deletedRows: 0,
    deletedCols: 0,
    bandingAppliedMsg: ""
  };
  
  // Check if sheet is too large for single operation
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const isLargeSheet = (lastRow * lastCol) > 10000; // Threshold for large sheets
  
  // Execute each module function with error handling
  if (FORMAT_CONFIG.deleteEmptyColumns) {
    try {
      results.deletedCols = deleteEmptyColumns(sheet);
    } catch (e) {
      results.errors.push("Error deleting blank columns: " + e.message);
    }
  }
  
  if (FORMAT_CONFIG.deleteEmptyRows) {
    try {
      results.deletedRows = deleteEmptyRows(sheet);
    } catch (e) {
      results.errors.push("Error deleting blank rows: " + e.message);
    }
  }
  
  const dataRange = getDataRange(sheet, results.errors);
  if (!dataRange) return;
  
  if (isLargeSheet) {
    processLargeSheet(sheet, results.errors);
  } else {
    // For smaller sheets, apply all formatting at once
    setAlignment(dataRange, results.errors);
    
    if (FORMAT_CONFIG.setBorders) {
      setBorders(dataRange, results.errors);
    }
  }
  
  freezePanes(sheet, results.errors);
  
  if (FORMAT_CONFIG.banding) {
    results.bandingAppliedMsg = applyBanding(sheet, dataRange, results.errors);
  }
  
  if (FORMAT_CONFIG.headerBold) {
    boldHeaderRow(sheet, results.errors);
  }
  
  if (FORMAT_CONFIG.customRowHeight) {
    setCustomRowHeight(sheet, results.errors);
  }
  
  if (FORMAT_CONFIG.autoAdjust) {
    autoResizeDimensions(sheet, results.errors);
  }
  
  const finishTime = new Date();
  const elapsed = ((finishTime - startTime) / 1000).toFixed(2);
  
  displayResults(ui, elapsed, results);
}

// UI Helper Functions
function getUi() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null; // No UI in scheduled runs
  }
}

function confirmFormatting(ui) {
  if (!ui) return true; // Auto-confirm in non-UI mode
  
  const response = ui.alert(
    "Format Entire Sheet",
    "This will delete any rows/columns outside the active data range, set borders and alignment, " +
    "freeze the header, and apply alternating banding (with header/footer) and bold the header row. Proceed?",
    ui.ButtonSet.YES_NO
  );
  
  if (response != ui.Button.YES) {
    ui.alert("Formatting cancelled.");
    return false;
  }
  
  return true;
}

function displayResults(ui, elapsed, results) {
  if (!ui) return;
  
  const errorMsg = results.errors.length > 0 
    ? "\nErrors encountered:\n" + results.errors.join("\n") 
    : "\nNo errors encountered.";
    
  ui.alert("Formatting Complete",
    "Formatting applied in " + elapsed + " seconds.\n" +
    "Deleted " + results.deletedRows + " blank rows and " + 
    results.deletedCols + " blank columns.\n" +
    results.bandingAppliedMsg + errorMsg,
    ui.ButtonSet.OK);
}

// Optimized Formatting Module Functions

// Optimized empty column detection and deletion
function deleteEmptyColumns(sheet) {
  let deletedCols = 0;
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  
  // Get all values at once to minimize API calls
  const allValues = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  const columnsToDelete = [];
  
  // Check each column (process in memory rather than making API calls)
  for (let col = maxCols - 1; col >= 0; col--) {
    let isEmpty = true;
    
    // Check if every cell in this column is empty
    for (let row = 0; row < maxRows; row++) {
      if (allValues[row][col] !== "" && allValues[row][col] !== null) {
        isEmpty = false;
        break;
      }
    }
    
    if (isEmpty) {
      columnsToDelete.push(col + 1); // +1 because sheet columns are 1-indexed
    }
  }
  
  // Delete columns in batches from right to left
  for (let i = 0; i < columnsToDelete.length; i++) {
    sheet.deleteColumn(columnsToDelete[i]);
    deletedCols++;
  }
  
  return deletedCols;
}

// Optimized empty row detection and deletion
function deleteEmptyRows(sheet) {
  let deletedRows = 0;
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  
  // Get all values at once to minimize API calls
  const allValues = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  const rowsToDelete = [];
  
  // Check each row (process in memory rather than making API calls)
  for (let row = maxRows - 1; row >= 0; row--) {
    let isEmpty = true;
    
    // Check if every cell in this row is empty
    for (let col = 0; col < maxCols; col++) {
      if (allValues[row][col] !== "" && allValues[row][col] !== null) {
        isEmpty = false;
        break;
      }
    }
    
    if (isEmpty) {
      rowsToDelete.push(row + 1); // +1 because sheet rows are 1-indexed
    }
  }
  
  // Delete rows in batches from bottom to top
  for (let i = 0; i < rowsToDelete.length; i++) {
    sheet.deleteRow(rowsToDelete[i]);
    deletedRows++;
  }
  
  return deletedRows;
}

function getDataRange(sheet, errors) {
  try {
    return sheet.getDataRange();
  } catch (e) {
    errors.push("Error retrieving data range: " + e.message);
    return null;
  }
}

// Optimized alignment setting using RangeList for batch operations
function setAlignment(dataRange, errors) {
  try {
    // Use RangeList for batch operations
    const sheet = dataRange.getSheet();
    const rangeList = sheet.getRangeList([dataRange.getA1Notation()]);
    
    rangeList
      .setHorizontalAlignment(FORMAT_CONFIG.horizontalAlignment)
      .setVerticalAlignment(FORMAT_CONFIG.verticalAlignment);
  } catch (e) {
    errors.push("Error setting cell alignment: " + e.message);
  }
}

// Optimized border application with batch operations
function setBorders(dataRange, errors) {
  try {
    const options = FORMAT_CONFIG.borderOptions;
    
    // Use setBorder once instead of individual calls
    dataRange.setBorder(
      options.top,
      options.left,
      options.bottom, 
      options.right,
      options.vertical,
      options.horizontal,
      null,
      SpreadsheetApp.BorderStyle.SOLID
    );
  } catch (e) {
    errors.push("Error setting borders: " + e.message);
  }
}

function freezePanes(sheet, errors) {
  try {
    if (FORMAT_CONFIG.freezeFirstRow) sheet.setFrozenRows(1);
    if (FORMAT_CONFIG.freezeFirstColumn) sheet.setFrozenColumns(1);
  } catch (e) {
    errors.push("Error freezing rows/columns: " + e.message);
  }
}

// Optimized banding with caching for performance
function applyBanding(sheet, dataRange, errors) {
  try {
    // Check cached status first
    const scriptProperties = PropertiesService.getScriptProperties();
    const sheetId = sheet.getSheetId();
    const lastBandingKey = `lastBanding_${sheetId}`;
    const lastBandingTime = scriptProperties.getProperty(lastBandingKey);
    
    // Check if banding was recently applied
    if (lastBandingTime) {
      const lastTime = new Date(lastBandingTime).getTime();
      const currentTime = new Date().getTime();
      const timeDiff = (currentTime - lastTime) / (1000 * 60); // Minutes
      
      // If banding was applied in the last hour, skip
      if (timeDiff < 60) {
        return "Skipped banding - recently applied.";
      }
    }
    
    // Apply banding only if not already present
    const existingBandings = sheet.getBandings() || [];
    if (existingBandings.length > 0) {
      return "Alternating banding was already applied; skipping banding.";
    } else {
      dataRange.applyRowBanding(FORMAT_CONFIG.bandingTheme, true, true);
      
      // Cache the current timestamp
      scriptProperties.setProperty(lastBandingKey, new Date().toISOString());
      return "Alternating banding applied with header and footer.";
    }
  } catch (e) {
    errors.push("Error applying row banding: " + e.message);
    return "Error applying banding.";
  }
}

function boldHeaderRow(sheet, errors) {
  try {
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setFontWeight("bold");
  } catch (e) {
    errors.push("Error bolding header row: " + e.message);
  }
}

function setCustomRowHeight(sheet, errors) {
  try {
    const lastRow = sheet.getLastRow();
    sheet.setRowHeights(1, lastRow, FORMAT_CONFIG.customRowHeight);
  } catch (e) {
    errors.push("Error setting row heights: " + e.message);
  }
}

// Optimized auto-resize using batch processing
function autoResizeDimensions(sheet, errors) {
  try {
    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    
    // Batch columns in groups to reduce API calls
    const BATCH_SIZE = 10; // Adjust based on your typical sheet size
    
    // Auto-resize columns in batches
    for (let startCol = 1; startCol <= lastCol; startCol += BATCH_SIZE) {
      const endCol = Math.min(startCol + BATCH_SIZE - 1, lastCol);
      sheet.autoResizeColumns(startCol, endCol - startCol + 1);
    }
    
    // Auto-resize rows in batches
    for (let startRow = 1; startRow <= lastRow; startRow += BATCH_SIZE) {
      const endRow = Math.min(startRow + BATCH_SIZE - 1, lastRow);
      sheet.autoResizeRows(startRow, endRow - startRow + 1);
    }
  } catch (e) {
    errors.push("Error auto-resizing rows/columns: " + e.message);
  }
}

// Process large sheets in chunks to avoid timeout errors
function processLargeSheet(sheet, errors) {
  const MAX_ROWS_PER_CHUNK = 1000;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  try {
    // Process sheet in row chunks to avoid memory issues
    for (let startRow = 1; startRow <= lastRow; startRow += MAX_ROWS_PER_CHUNK) {
      const endRow = Math.min(startRow + MAX_ROWS_PER_CHUNK - 1, lastRow);
      const chunkRange = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol);
      
      // Apply formatting to this chunk
      chunkRange.setHorizontalAlignment(FORMAT_CONFIG.horizontalAlignment);
      chunkRange.setVerticalAlignment(FORMAT_CONFIG.verticalAlignment);
      
      if (FORMAT_CONFIG.setBorders) {
        const options = FORMAT_CONFIG.borderOptions;
        chunkRange.setBorder(
          options.top,
          options.left,
          options.bottom, 
          options.right,
          options.vertical,
          options.horizontal
        );
      }
    }
  } catch (e) {
    errors.push("Error processing large sheet in chunks: " + e.message);
  }
}

// Helper function for creating a menu item to run the script
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Advanced Formatting')
    .addItem('Format Entire Sheet', 'formatEntireSheet')
    .addToUi();
}
