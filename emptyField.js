/**
 * Empty Field Configuration Management - UPDATED WITH ITERATION FIELD
 * Handles the configuration for excluding empty Summary, Fix Version, and Iteration fields
 */

// Configuration storage key
const CONFIG_PROPERTY_KEY = 'EMPTY_FIELD_CONFIG';

// ADD THIS NEW CONSTANT HERE:
const PLACEHOLDER_TEXTS = [
  'title missing',
  'unnamed epic',
  'fix version not available',
  'iteration unknown',
  'n/a',
  'unknown',
  'not available',
  'tbd',
  'none',
  'no title',
  '(no title)',
  '[no title]',
  'missing'
];

/**
 * Add menu item to open configuration dialog
 * Call this from your onOpen() function
 */
function addEmptyFieldConfigMenu() {
  const ui = SpreadsheetApp.getUi();
  // Add to your existing menu or create a new one
  ui.createMenu('Slide Configuration')
      .addItem('Empty Field Display Settings', 'openEmptyFieldConfig')
      .addToUi();
}

/**
 * Open the configuration dialog
 */
function openEmptyFieldConfig() {
  const html = HtmlService.createHtmlOutputFromFile('EmptyFieldConfig')
      .setWidth(600)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Empty Field Display Configuration');
}

/**
 * Get the current configuration
 * @returns {Object} Configuration object with excludeSummary, excludeFixVersion, and excludeIteration properties
 */
function getEmptyFieldConfiguration() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const configJson = scriptProperties.getProperty(CONFIG_PROPERTY_KEY);
  
  if (configJson) {
    try {
      return JSON.parse(configJson);
    } catch (e) {
      Logger.log('Error parsing configuration: ' + e);
      return getDefaultConfiguration();
    }
  }
  
  return getDefaultConfiguration();
}

/**
 * Get default configuration
 * @returns {Object} Default configuration
 */
function getDefaultConfiguration() {
  return {
    excludeSummary: false,
    excludeFixVersion: false,
    excludeIteration: false
  };
}

/**
 * Save the configuration
 * @param {Object} config - Configuration object
 * @returns {boolean} Success status
 */
function saveEmptyFieldConfiguration(config) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(CONFIG_PROPERTY_KEY, JSON.stringify(config));
    Logger.log('Configuration saved: ' + JSON.stringify(config));
    return true;
  } catch (e) {
    Logger.log('Error saving configuration: ' + e);
    throw new Error('Failed to save configuration: ' + e.message);
  }
}

/**
 * Helper function to check if a field should be displayed
 * @param {string} fieldValue - The value of the field
 * @param {string} fieldType - Type of field ('summary', 'fixVersion', or 'iteration')
 * @returns {boolean} True if field should be displayed, false otherwise
 */
function shouldDisplayField(fieldValue, fieldType) {
  const config = getEmptyFieldConfiguration();
  
  // Check if value is empty (null, undefined, empty string, or whitespace only)
  const isEmpty = !fieldValue || fieldValue.toString().trim() === '';
  
  if (!isEmpty) {
    return true; // Always display non-empty values
  }
  
  // If empty, check configuration
  switch (fieldType.toLowerCase()) {
    case 'summary':
      return !config.excludeSummary; // Display if NOT excluded
    case 'fixversion':
      return !config.excludeFixVersion; // Display if NOT excluded
    case 'iteration':
      return !config.excludeIteration; // Display if NOT excluded
    default:
      return true; // Display by default for unknown field types
  }
}

/**
 * Get display value for a field with default text handling
 * @param {string} fieldValue - The actual field value
 * @param {string} fieldType - Type of field ('summary', 'fixVersion', or 'iteration')
 * @param {string} defaultText - Default text to show if field is empty (optional)
 * @returns {string|null} The value to display, or null if field should be hidden
 */
function getFieldDisplayValue(fieldValue, fieldType, defaultText) {
  const config = getEmptyFieldConfiguration();
  
  // Check if value is truly empty (null, undefined, or whitespace)
  const isTrulyEmpty = !fieldValue || fieldValue.toString().trim() === '';
  
  // Check if it's a placeholder text
  let isPlaceholder = false;
  if (!isTrulyEmpty && typeof PLACEHOLDER_TEXTS !== 'undefined') {
    isPlaceholder = PLACEHOLDER_TEXTS.some(placeholder => 
      fieldValue.toString().toLowerCase().trim() === placeholder.toLowerCase()
    );
  }
  
  // Treat both as "empty"
  const isEmpty = isTrulyEmpty || isPlaceholder;
  
  // If has real content, return it
  if (!isEmpty) {
    return fieldValue.toString().trim();
  }
  
  // Field is empty or placeholder - check configuration
  let shouldDisplay = true;
  switch (fieldType.toLowerCase()) {
    case 'summary':
      shouldDisplay = !config.excludeSummary;
      break;
    case 'fixversion':
      shouldDisplay = !config.excludeFixVersion;
      break;
    case 'iteration':
      shouldDisplay = !config.excludeIteration;
      break;
  }
  
  // Return null if should be excluded, otherwise return default
  if (!shouldDisplay) {
    return null;
  }
  
  return defaultText || '';
}

/**
 * Get iteration display text with configuration support
 * @param {string} iterationValue - The iteration value from the data
 * @returns {string|null} The iteration text to display, or null if should be hidden
 */
function getIterationDisplayText(iterationValue) {
  return getFieldDisplayValue(iterationValue, 'iteration', 'Iteration unknown');
}

/**
 * Utility function to log current configuration (for debugging)
 */
function logCurrentConfiguration() {
  const config = getEmptyFieldConfiguration();
  Logger.log('Current Empty Field Configuration:');
  Logger.log('  Exclude Summary: ' + config.excludeSummary);
  Logger.log('  Exclude Fix Version: ' + config.excludeFixVersion);
  Logger.log('  Exclude Iteration: ' + config.excludeIteration);
}

/**
 * Reset configuration to defaults (utility function)
 */
function resetEmptyFieldConfiguration() {
  const defaultConfig = getDefaultConfiguration();
  saveEmptyFieldConfiguration(defaultConfig);
  Logger.log('Configuration reset to defaults');
  SpreadsheetApp.getUi().alert('Configuration has been reset to defaults.');
}

/**
 * Test function to verify configuration is working
 */
function testEmptyFieldConfiguration() {
  Logger.log('=== Testing Empty Field Configuration ===');
  
  // Test 1: Empty string with exclude enabled
  saveEmptyFieldConfiguration({ excludeSummary: true, excludeFixVersion: true, excludeIteration: true });
  var result1 = shouldDisplayField('', 'summary');
  Logger.log('Test 1 - Empty summary with exclude ON: ' + (result1 ? 'FAIL (should be false)' : 'PASS'));
  
  // Test 2: Empty string with exclude disabled
  saveEmptyFieldConfiguration({ excludeSummary: false, excludeFixVersion: false, excludeIteration: false });
  var result2 = shouldDisplayField('', 'summary');
  Logger.log('Test 2 - Empty summary with exclude OFF: ' + (result2 ? 'PASS' : 'FAIL (should be true)'));
  
  // Test 3: Non-empty value
  var result3 = shouldDisplayField('Some value', 'summary');
  Logger.log('Test 3 - Non-empty summary: ' + (result3 ? 'PASS' : 'FAIL (should be true)'));
  
  // Test 4: Getting display value with default
  saveEmptyFieldConfiguration({ excludeSummary: true, excludeFixVersion: true, excludeIteration: true });
  var result4 = getFieldDisplayValue('', 'fixVersion', 'Fix Version not available');
  Logger.log('Test 4 - Display value with exclude ON: ' + (result4 === null ? 'PASS' : 'FAIL (should be null)'));
  
  // Test 5: Getting display value without exclude
  saveEmptyFieldConfiguration({ excludeSummary: false, excludeFixVersion: false, excludeIteration: false });
  var result5 = getFieldDisplayValue('', 'fixVersion', 'Fix Version not available');
  Logger.log('Test 5 - Display value with exclude OFF: ' + (result5 === 'Fix Version not available' ? 'PASS' : 'FAIL'));
  
  // Test 6: Iteration field with exclude enabled
  saveEmptyFieldConfiguration({ excludeSummary: true, excludeFixVersion: true, excludeIteration: true });
  var result6 = getIterationDisplayText('');
  Logger.log('Test 6 - Empty iteration with exclude ON: ' + (result6 === null ? 'PASS' : 'FAIL (should be null)'));
  
  // Test 7: Iteration field with exclude disabled
  saveEmptyFieldConfiguration({ excludeSummary: false, excludeFixVersion: false, excludeIteration: false });
  var result7 = getIterationDisplayText('');
  Logger.log('Test 7 - Empty iteration with exclude OFF: ' + (result7 === 'Iteration unknown' ? 'PASS' : 'FAIL'));
  
  // NEW TESTS for placeholder detection
  saveEmptyFieldConfiguration({ excludeSummary: true, excludeFixVersion: true, excludeIteration: true });
  var result8 = getFieldDisplayValue('title missing', 'summary', 'Unnamed Epic');
  Logger.log('Test 8 - "title missing" with exclude ON: ' + (result8 === null ? 'PASS' : 'FAIL (got: ' + result8 + ')'));
  
  saveEmptyFieldConfiguration({ excludeSummary: false, excludeFixVersion: false, excludeIteration: false });
  var result9 = getFieldDisplayValue('title missing', 'summary', 'Unnamed Epic');
  Logger.log('Test 9 - "title missing" with exclude OFF: ' + (result9 === 'Unnamed Epic' ? 'PASS' : 'FAIL (got: ' + result9 + ')'));
  
  Logger.log('=== Test Complete ===');
}