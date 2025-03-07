// Functions to handle UI and sidebar interactions

// Show the formatting sidebar
function showFormattingSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('FormattingSidebar')
    .setTitle('Sheet Formatting Options')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Get current configuration for the sidebar
function getCurrentConfig() {
  return FORMAT_CONFIG;
}

// Apply configuration from the sidebar
function applyFormattingWithTarget(config) {
  // Update the global configuration
  Object.assign(FORMAT_CONFIG, config);
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const targetRange = determineTargetRange(config, sheet);
    
    if (config.formatTarget === 'entireSheet') {
      // For entire sheet, just run the existing function
      formatEntireSheet();
      return "Formatting applied to the entire sheet";
    } else {
      // For specific ranges, use targeted formatting
      const results = { errors: [], deletedRows: 0, deletedCols: 0, bandingAppliedMsg: "" };
      formatTargetedRange(sheet, targetRange, results);
      
      if (results.errors.length > 0) {
        return "Formatting applied with errors: " + results.errors.join(", ");
      }
      return "Formatting applied successfully to selected range";
    }
  } catch (e) {
    return "Error: " + e.message;
  }
}

// Configuration Management
function saveConfiguration(configName) {
  if (!configName) return "Please provide a configuration name.";
  
  try {
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('format_' + configName, JSON.stringify(FORMAT_CONFIG));
    return "Configuration '" + configName + "' saved successfully!";
  } catch (e) {
    return "Error saving configuration: " + e.message;
  }
}

function loadConfiguration(configName) {
  try {
    const userProps = PropertiesService.getUserProperties();
    const savedConfig = userProps.getProperty('format_' + configName);
    
    if (!savedConfig) {
      return { success: false, message: "Configuration not found." };
    }
    
    // Update the global configuration
    Object.assign(FORMAT_CONFIG, JSON.parse(savedConfig));
    return { 
      success: true, 
      message: "Configuration '" + configName + "' loaded successfully!",
      config: FORMAT_CONFIG
    };
  } catch (e) {
    return { success: false, message: "Error loading configuration: " + e.message };
  }
}

function getSavedConfigurations() {
  try {
    const userProps = PropertiesService.getUserProperties();
    const props = userProps.getProperties();
    const configs = [];
    
    for (const key in props) {
      if (key.startsWith('format_')) {
        configs.push(key.substring(7)); // Remove 'format_' prefix
      }
    }
    
    return configs;
  } catch (e) {
    return [];
  }
}

// Add the onOpen function to create menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Format Tool')
    .addItem('Show Formatting Options', 'showFormattingSidebar')
    .addSeparator()
    .addItem('Format Entire Sheet (Quick)', 'formatEntireSheet')
    .addToUi();
}