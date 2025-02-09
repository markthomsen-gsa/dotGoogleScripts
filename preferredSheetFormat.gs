// Configuration object for formatting options.
const FORMAT_CONFIG = {
  deleteEmptyRows: true,
  deleteEmptyColumns: true,
  setBorders: true, // Enable borders by default.
  borderOptions: { top: true, left: true, bottom: true, right: true, vertical: true, horizontal: true },
  headerBold: true,
  customRowHeight: null, // Set a fixed row height (in pixels) if desired; leave as null to skip.
  autoAdjust: false,   // Auto-adjust row height and column width (disabled by default).
  horizontalAlignment: "left",
  verticalAlignment: "middle",
  freezeFirstRow: true,
  freezeFirstColumn: true,
  banding: true,
  bandingTheme: SpreadsheetApp.BandingTheme.BLUE
};

function formatEntireSheet() {
  var ui;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {
    ui = null; // No UI in scheduled runs.
  }
  
  // Confirmation dialog if running interactively.
  if (ui) {
    var response = ui.alert(
      "Format Entire Sheet",
      "This will delete any rows/columns outside the active data range, set borders and alignment, freeze the header, and apply alternating banding (with header/footer) and bold the header row. Proceed?",
      ui.ButtonSet.YES_NO
    );
    if (response != ui.Button.YES) {
      ui.alert("Formatting cancelled.");
      return;
    }
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var startTime = new Date();
  var errors = [];
  var deletedRows = 0;
  var deletedCols = 0;
  var bandingAlreadyApplied = false;
  var bandingAppliedMsg = "";
  
  // Get the full grid (using the sheet’s maximum rows and columns).
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  var grid = sheet.getRange(1, 1, maxRows, maxCols).getValues();
  
  // Delete empty columns (from rightmost to leftmost).
  if (FORMAT_CONFIG.deleteEmptyColumns) {
    try {
      for (var col = maxCols; col >= 1; col--) {
        var colRange = sheet.getRange(1, col, maxRows, 1);
        var colValues = colRange.getValues();
        var isEmpty = colValues.every(function(row) {
          return row[0] === "" || row[0] === null;
        });
        if (isEmpty) {
          sheet.deleteColumn(col);
          deletedCols++;
        }
      }
    } catch (e) {
      errors.push("Error deleting blank columns: " + e.message);
    }
  }
  
  // Update maxCols after deleting columns.
  maxCols = sheet.getMaxColumns();
  
  // Delete empty rows (from bottom to top).
  if (FORMAT_CONFIG.deleteEmptyRows) {
    try {
      var currentRows = sheet.getMaxRows();
      for (var row = currentRows; row >= 1; row--) {
        var rowRange = sheet.getRange(row, 1, 1, maxCols);
        var rowValues = rowRange.getValues();
        var isEmpty = rowValues[0].every(function(cell) {
          return cell === "" || cell === null;
        });
        if (isEmpty) {
          sheet.deleteRow(row);
          deletedRows++;
        }
      }
    } catch (e) {
      errors.push("Error deleting blank rows: " + e.message);
    }
  }
  
  // Now assume the active data range is all that matters.
  var dataRange;
  try {
    dataRange = sheet.getDataRange();
  } catch (e) {
    errors.push("Error retrieving data range: " + e.message);
  }
  
  // Set cell alignment.
  try {
    dataRange.setHorizontalAlignment(FORMAT_CONFIG.horizontalAlignment);
    dataRange.setVerticalAlignment(FORMAT_CONFIG.verticalAlignment);
  } catch (e) {
    errors.push("Error setting cell alignment: " + e.message);
  }
  
  // Set borders if enabled.
  if (FORMAT_CONFIG.setBorders) {
    try {
      dataRange.setBorder(
        FORMAT_CONFIG.borderOptions.top,
        FORMAT_CONFIG.borderOptions.left,
        FORMAT_CONFIG.borderOptions.bottom,
        FORMAT_CONFIG.borderOptions.right,
        FORMAT_CONFIG.borderOptions.vertical,
        FORMAT_CONFIG.borderOptions.horizontal
      );
    } catch (e) {
      errors.push("Error setting borders: " + e.message);
    }
  }
  
  // Freeze the first row and first column if enabled.
  try {
    if (FORMAT_CONFIG.freezeFirstRow) sheet.setFrozenRows(1);
    if (FORMAT_CONFIG.freezeFirstColumn) sheet.setFrozenColumns(1);
  } catch (e) {
    errors.push("Error freezing rows/columns: " + e.message);
  }
  
  // Apply alternating banding if enabled—test if already applied.
  if (FORMAT_CONFIG.banding) {
    try {
      var existingBandings = sheet.getBandings() || [];
      if (existingBandings.length > 0) {
        bandingAlreadyApplied = true;
        bandingAppliedMsg = "Alternating banding was already applied; skipping banding.";
      } else {
        // Apply banding with header and footer.
        dataRange.applyRowBanding(FORMAT_CONFIG.bandingTheme, true, true);
        bandingAppliedMsg = "Alternating banding applied with header and footer.";
      }
    } catch (e) {
      errors.push("Error applying row banding: " + e.message);
    }
  }
  
  // Bold the header row.
  if (FORMAT_CONFIG.headerBold) {
    try {
      var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      headerRange.setFontWeight("bold");
    } catch (e) {
      errors.push("Error bolding header row: " + e.message);
    }
  }
  
  // Set custom row height if specified.
  if (FORMAT_CONFIG.customRowHeight) {
    try {
      var lr = sheet.getLastRow();
      sheet.setRowHeights(1, lr, FORMAT_CONFIG.customRowHeight);
    } catch (e) {
      errors.push("Error setting row heights: " + e.message);
    }
  }
  
  // Auto-adjust row heights and column widths if enabled.
  if (FORMAT_CONFIG.autoAdjust) {
    try {
      for (var col = 1; col <= sheet.getLastColumn(); col++) {
        sheet.autoResizeColumn(col);
      }
      for (var row = 1; row <= sheet.getLastRow(); row++) {
        sheet.autoResizeRow(row);
      }
    } catch (e) {
      errors.push("Error auto-resizing rows/columns: " + e.message);
    }
  }
  
  var finishTime = new Date();
  var elapsed = ((finishTime - startTime) / 1000).toFixed(2);
  
  if (ui) {
    var errorMsg = errors.length > 0 ? "\nErrors encountered:\n" + errors.join("\n") : "\nNo errors encountered.";
    ui.alert("Formatting Complete",
      "Formatting applied in " + elapsed + " seconds.\nDeleted " + deletedRows + " blank rows and " + deletedCols + " blank columns.\n" +
      bandingAppliedMsg + errorMsg,
      ui.ButtonSet.OK);
  }
}
