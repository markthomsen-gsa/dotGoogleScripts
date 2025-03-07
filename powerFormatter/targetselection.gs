// Enhanced format target options
const FORMAT_TARGETS = {
  ENTIRE_SHEET: 'entireSheet',
  SELECTED_RANGE: 'selectedRange',
  CUSTOM_RANGE: 'customRange',
  DATA_RANGE: 'dataRange',
  NAMED_RANGE: 'namedRange',
  DETECT_TABLE: 'detectTable',
  FILTERED_ROWS: 'filteredRows',
  CONDITIONAL: 'conditional',
  CURRENT_COLUMN: 'currentColumn',
  CURRENT_ROW: 'currentRow',
  VISIBLE_CELLS: 'visibleCells'
};

// Smart range detection function
function detectTableAroundSelection(sheet) {
  try {
    // Get the current selection
    const activeRange = sheet.getActiveRange();
    if (!activeRange) {
      throw new Error("No cell selected. Please select a cell within your data.");
    }
    
    // Start with active cell position
    const startRow = activeRange.getRow();
    const startCol = activeRange.getColumn();
    
    // Get all values in the sheet
    const allValues = sheet.getDataRange().getValues();
    
    // Find table boundaries
    let tableTop = startRow;
    let tableBottom = startRow;
    let tableLeft = startCol;
    let tableRight = startCol;
    
    // Detect if we're inside a table by checking for content
    const isWithinTable = allValues[startRow-1][startCol-1] !== "";
    
    if (!isWithinTable) {
      return activeRange; // Not in a table, just return original selection
    }
    
    // Find the top boundary (look upward for empty rows)
    for (let row = startRow - 2; row >= 0; row--) {
      let isEmpty = true;
      for (let col = 0; col < allValues[row].length; col++) {
        if (allValues[row][col] !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        break;
      }
      tableTop--;
    }
    
    // Find the bottom boundary (look downward for empty rows)
    for (let row = startRow; row < allValues.length; row++) {
      let isEmpty = true;
      for (let col = 0; col < allValues[row].length; col++) {
        if (allValues[row][col] !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        break;
      }
      tableBottom = row + 1;
    }
    
    // Find the left boundary (look leftward for empty columns)
    for (let col = startCol - 2; col >= 0; col--) {
      let isEmpty = true;
      for (let row = tableTop - 1; row < tableBottom; row++) {
        if (row < allValues.length && col < allValues[row].length && allValues[row][col] !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        break;
      }
      tableLeft--;
    }
    
    // Find the right boundary (look rightward for empty columns)
    const maxCol = allValues[0].length;
    for (let col = startCol; col < maxCol; col++) {
      let isEmpty = true;
      for (let row = tableTop - 1; row < tableBottom; row++) {
        if (row < allValues.length && col < allValues[row].length && allValues[row][col] !== "") {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        break;
      }
      tableRight = col + 1;
    }
    
    // Return the detected table range
    return sheet.getRange(tableTop, tableLeft, tableBottom - tableTop + 1, tableRight - tableLeft + 1);
  } catch (e) {
    Logger.log("Error in table detection: " + e.message);
    return sheet.getActiveRange() || sheet.getDataRange();
  }
}

// Get all named ranges in the spreadsheet
function getNamedRanges() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = spreadsheet.getNamedRanges();
  
  return namedRanges.map(range => ({
    name: range.getName(),
    range: range.getRange().getA1Notation(),
    sheet: range.getRange().getSheet().getName()
  }));
}

// Get filtered rows
function getFilteredRows(sheet) {
  try {
    const filter = sheet.getFilter();
    if (!filter) {
      throw new Error("No filter found on this sheet.");
    }
    
    const range = filter.getRange();
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    
    // We can't directly get filtered rows, so we need to check each row's visibility
    const visibleRows = [];
    
    for (let i = startRow + 1; i < startRow + numRows; i++) {
      // If we can get a value from the row, it's visible
      try {
        // Try to get a value from the first cell in the row
        sheet.getRange(i, range.getColumn()).getValue();
        visibleRows.push(i);
      } catch (e) {
        // If error, row is hidden by filter
        continue;
      }
    }
    
    if (visibleRows.length === 0) {
      throw new Error("No visible filtered rows found.");
    }
    
    // Create a range with just the visible rows
    const ranges = visibleRows.map(row => sheet.getRange(row, range.getColumn(), 1, numCols));
    
    // Return as a RangeList
    return sheet.getRangeList(ranges.map(r => r.getA1Notation()));
  } catch (e) {
    throw new Error("Error getting filtered rows: " + e.message);
  }
}

// Get cells that match condition
function getCellsMatchingCondition(sheet, condition) {
  try {
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const numRows = values.length;
    const numCols = values[0].length;
    const matchingCells = [];
    
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        const cellValue = values[row][col];
        let matches = false;
        
        switch (condition.type) {
          case 'contains':
            matches = String(cellValue).includes(condition.value);
            break;
          case 'equals':
            matches = cellValue == condition.value; // Use loose equality for numeric strings
            break;
          case 'greaterThan':
            matches = cellValue > condition.value;
            break;
          case 'lessThan':
            matches = cellValue < condition.value;
            break;
          case 'blank':
            matches = cellValue === "" || cellValue === null;
            break;
          case 'notBlank':
            matches = cellValue !== "" && cellValue !== null;
            break;
          case 'formula':
            // Need to check directly from the sheet for formulas
            matches = sheet.getRange(row + 1, col + 1).getFormula() !== "";
            break;
          default:
            matches = false;
        }
        
        if (matches) {
          matchingCells.push(sheet.getRange(row + 1, col + 1));
        }
      }
    }
    
    if (matchingCells.length === 0) {
      throw new Error("No cells match the condition.");
    }
    
    // Return as a RangeList
    return sheet.getRangeList(matchingCells.map(r => r.getA1Notation()));
  } catch (e) {
    throw new Error("Error matching condition: " + e.message);
  }
}

// Get all visible cells (not hidden by row/column hiding)
function getVisibleCells(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const visibleRanges = [];
    
    for (let row = 1; row <= maxRows; row++) {
      if (sheet.isRowHiddenByUser(row)) continue;
      
      for (let col = 1; col <= maxCols; col++) {
        if (sheet.isColumnHiddenByUser(col)) continue;
        
        visibleRanges.push(sheet.getRange(row, col));
      }
    }
    
    if (visibleRanges.length === 0) {
      throw new Error("No visible cells found.");
    }
    
    // Return as a RangeList
    return sheet.getRangeList(visibleRanges.map(r => r.getA1Notation()));
  } catch (e) {
    throw new Error("Error getting visible cells: " + e.message);
  }
}

// Main function to determine target range based on selection
function determineTargetRange(config, sheet) {
  try {
    switch(config.formatTarget) {
      case FORMAT_TARGETS.ENTIRE_SHEET:
        // Format entire sheet
        return sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
        
      case FORMAT_TARGETS.SELECTED_RANGE:
        // Format only the currently selected range
        const activeRange = sheet.getActiveRange();
        if (!activeRange) {
          throw new Error("No range selected. Please select cells to format.");
        }
        return activeRange;
        
      case FORMAT_TARGETS.CUSTOM_RANGE:
        // Format a specific range provided by the user
        if (!config.customRange) {
          throw new Error("No custom range specified.");
        }
        try {
          return sheet.getRange(config.customRange);
        } catch (e) {
          throw new Error("Invalid range: " + config.customRange);
        }
        
      case FORMAT_TARGETS.DATA_RANGE:
        // Format only the range containing data
        return sheet.getDataRange();
        
      case FORMAT_TARGETS.NAMED_RANGE:
        // Format a named range
        if (!config.namedRange) {
          throw new Error("No named range specified.");
        }
        
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const namedRanges = spreadsheet.getNamedRanges();
        const targetRange = namedRanges.find(range => range.getName() === config.namedRange);
        
        if (!targetRange) {
          throw new Error("Named range '" + config.namedRange + "' not found.");
        }
        
        return targetRange.getRange();
        
      case FORMAT_TARGETS.DETECT_TABLE:
        // Detect table boundaries around current selection
        return detectTableAroundSelection(sheet);
        
      case FORMAT_TARGETS.FILTERED_ROWS:
        // Get only visible rows in a filtered range
        return getFilteredRows(sheet);
        
      case FORMAT_TARGETS.CONDITIONAL:
        // Get cells matching a condition
        if (!config.condition) {
          throw new Error("No condition specified.");
        }
        return getCellsMatchingCondition(sheet, config.condition);
        
      case FORMAT_TARGETS.CURRENT_COLUMN:
        // Format the current column
        const activeCol = sheet.getActiveRange()?.getColumn();
        if (!activeCol) {
          throw new Error("No column selected.");
        }
        return sheet.getRange(1, activeCol, sheet.getMaxRows(), 1);
        
      case FORMAT_TARGETS.CURRENT_ROW:
        // Format the current row
        const activeRow = sheet.getActiveRange()?.getRow();
        if (!activeRow) {
          throw new Error("No row selected.");
        }
        return sheet.getRange(activeRow, 1, 1, sheet.getMaxColumns());
        
      case FORMAT_TARGETS.VISIBLE_CELLS:
        // Format only visible cells (not hidden by row/column hiding)
        return getVisibleCells(sheet);
        
      default:
        // Default to entire sheet
        return sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    }
  } catch (e) {
    throw new Error("Error determining target range: " + e.message);
  }
}

// Function to preview the target range without applying formatting
function previewTargetRange(config) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const targetRange = determineTargetRange(config, sheet);
    
    if (!targetRange) {
      return {
        success: false,
        message: "Could not determine target range."
      };
    }
    
    // If it's a RangeList (for filtered rows, conditional selection)
    if (typeof targetRange.getRangeList === 'function') {
      const ranges = targetRange.getRangeList().getRanges();
      let totalCells = 0;
      ranges.forEach(range => {
        totalCells += range.getNumRows() * range.getNumColumns();
      });
      
      return {
        success: true,
        info: "Multiple ranges selected",
        rangeCount: ranges.length,
        cellCount: totalCells,
        a1Notation: "Multiple ranges"
      };
    }
    
    // For normal ranges
    const numRows = targetRange.getNumRows();
    const numCols = targetRange.getNumColumns();
    const a1Notation = targetRange.getA1Notation();
    const sheetName = targetRange.getSheet().getName();
    
    // Calculate how many cells have data
    let dataCellCount = 0;
    if (numRows * numCols <= 10000) { // Limit check for performance
      const values = targetRange.getValues();
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (values[i][j] !== "" && values[i][j] !== null) {
            dataCellCount++;
          }
        }
      }
    }
    
    return {
      success: true,
      info: `${sheetName}!${a1Notation}`,
      numRows: numRows,
      numCols: numCols,
      cellCount: numRows * numCols,
      dataCellCount: dataCellCount,
      a1Notation: a1Notation
    };
  } catch (e) {
    return {
      success: false,
      message: e.message
    };
  }
}