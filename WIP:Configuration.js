/**
 * JIRA Field Configuration Manager
 * Creates and manages a Configuration tab for controlling slide generation display and ordering
 * 
 * IMPORTANT: When generating slides, always check the Configuration tab first to determine
 * which fields to display and their custom ordering. If the Configuration tab is not found
 * or not populated, use default alphabetical ordering.
 */

// ===== FIELD CONFIGURATION =====
const CONFIG_METADATA_FIELDS = {
  'Value Stream/Org': {
    fieldId: 'customfield_10046',
    hasOrdering: false // <<< No ordering
  },
  'Portfolio Initiative': {
    fieldId: 'customfield_10049',
    hasOrdering: true // <<< Has ordering
  },
  'Program Initiative': {
    fieldId: 'customfield_10050',
    hasOrdering: false // <<< CHANGED: No longer has ordering
  },
};

const CONFIG_SHEET_NAME = 'Configuration';

// ===== MENU ADDITIONS =====
/**
 * Add this to your existing onOpen() function in Menu.gs:
 * .addItem('🔄 Refresh Field Configuration', 'refreshFieldConfiguration')
 */

/**
 * Main function to refresh field configuration in Configuration tab
 */
function refreshFieldConfiguration() {
 return logActivity('JIRA Field Configuration Refresh', () => {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    ss.toast('Fetching field metadata from JIRA...', '⏳ Loading', 10);
    
    // Fetch metadata for all configured fields
    const metadataResults = fetchAllFieldMetadata();
    
    // Write to Configuration sheet with proper structure
    writeConfigurationSheet(metadataResults);
    
    ss.toast('✅ Field configuration refreshed successfully!', '✅ Complete', 5);
    
  } catch (error) {
    console.error('Error refreshing field configuration:', error);
    ui.alert('Error', `Failed to refresh field configuration: ${error.message}`, ui.ButtonSet.OK);
  }
 }, {});
}

function fetchAllFieldMetadata() {
  const jiraConfig = getJiraConfig();
  const results = {};
  
  console.log('Fetching field metadata from JIRA...');
  
  // Fetch all field definitions from JIRA
  const url = `${jiraConfig.baseUrl}/rest/api/3/field`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: getJiraHeaders(),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`JIRA API error: ${response.getResponseCode()} - ${response.getContentText()}`);
    }
    
    const allFields = JSON.parse(response.getContentText());
    
    // For each configured field, find its metadata and get allowed values
    Object.keys(CONFIG_METADATA_FIELDS).forEach(fieldName => {
      const fieldConfig = CONFIG_METADATA_FIELDS[fieldName];
      const fieldId = fieldConfig.fieldId;
      const fieldDef = allFields.find(f => f.id === fieldId);
      
      if (fieldDef) {
        console.log(`Found metadata for ${fieldName} (${fieldId})`);
        
        // Get allowed values
        const allowedValues = fetchFieldAllowedValues(fieldId, fieldDef);
        
        results[fieldName] = {
          fieldId: fieldId,
          fieldType: fieldDef.schema ? fieldDef.schema.type : 'unknown',
          allowedValues: allowedValues,
          hasOrdering: fieldConfig.hasOrdering
        };
      } else {
        console.warn(`Field ${fieldName} (${fieldId}) not found in JIRA`);
        results[fieldName] = {
          fieldId: fieldId,
          fieldType: 'not found',
          allowedValues: [],
          hasOrdering: fieldConfig.hasOrdering
        };
      }
    });
    
  } catch (error) {
    console.error('Error fetching field metadata:', error);
    throw error;
  }
  
  return results;
}

function fetchFieldAllowedValues(fieldId, fieldDef) {
  const jiraConfig = getJiraConfig();
  let allowedValues = [];
  
  // Check if this is a field type that has allowed values
  const fieldType = fieldDef.schema ? fieldDef.schema.type : null;
  
  if (!fieldType || !['option', 'array'].includes(fieldType)) {
    return allowedValues;
  }
  
  console.log(`Fetching allowed values for ${fieldId}...`);
  
  // Strategy 1: Try the custom field context endpoint (more efficient)
  try {
    const contextUrl = `${jiraConfig.baseUrl}/rest/api/3/field/${fieldId}/context`;
    
    const contextResponse = UrlFetchApp.fetch(contextUrl, {
      method: 'GET',
      headers: getJiraHeaders(),
      muteHttpExceptions: true
    });
    
    if (contextResponse.getResponseCode() === 200) {
      const contexts = JSON.parse(contextResponse.getContentText());
      
      if (contexts.values && contexts.values.length > 0) {
        const contextId = contexts.values[0].id;
        
        // Now fetch the options for this context
        const optionsUrl = `${jiraConfig.baseUrl}/rest/api/3/field/${fieldId}/context/${contextId}/option`;
        
        const optionsResponse = UrlFetchApp.fetch(optionsUrl, {
          method: 'GET',
          headers: getJiraHeaders(),
          muteHttpExceptions: true
        });
        
        if (optionsResponse.getResponseCode() === 200) {
          const options = JSON.parse(optionsResponse.getContentText());
          
          if (options.values) {
            options.values.forEach(opt => {
              const value = opt.value || opt.name;
              if (value && !allowedValues.includes(value)) {
                allowedValues.push(value);
              }
            });
            
            console.log(`Found ${allowedValues.length} values using context endpoint`);
            return allowedValues;
          }
        }
      }
    }
  } catch (error) {
    console.warn(`Context endpoint failed for ${fieldId}: ${error.message}`);
  }
  
  // Strategy 2: Query actual issues to find unique values
  try {
    console.log(`Trying issue search for ${fieldId}...`);
    
    // Search for issues that have this field populated
    const searchUrl = `${jiraConfig.baseUrl}/rest/api/3/search/jql`;
    const jql = `${fieldId} is not EMPTY ORDER BY created DESC`;
    
    const searchResponse = UrlFetchApp.fetch(`${searchUrl}?jql=${encodeURIComponent(jql)}&maxResults=100&fields=${fieldId}`, {
      method: 'GET',
      headers: getJiraHeaders(),
      muteHttpExceptions: true
    });
    
    if (searchResponse.getResponseCode() === 200) {
      const searchData = JSON.parse(searchResponse.getContentText());
      
      if (searchData.issues) {
        const uniqueValues = new Set();
        
        searchData.issues.forEach(issue => {
          const fieldValue = issue.fields[fieldId];
          
          if (fieldValue) {
            let value = null;
            
            if (typeof fieldValue === 'string') {
              value = fieldValue;
            } else if (fieldValue.value) {
              value = fieldValue.value;
            } else if (fieldValue.name) {
              value = fieldValue.name;
            } else if (Array.isArray(fieldValue)) {
              fieldValue.forEach(v => {
                const arrValue = v.value || v.name || v.toString();
                if (arrValue) uniqueValues.add(arrValue);
              });
            }
            
            if (value) {
              uniqueValues.add(value);
            }
          }
        });
        
        allowedValues = Array.from(uniqueValues).sort();
        console.log(`Found ${allowedValues.length} unique values from issues`);
      }
    }
    
  } catch (error) {
    console.warn(`Issue search failed for ${fieldId}: ${error.message}`);
  }
  
  return allowedValues;
}

function getDefaultConfiguration() {
  console.log('Using default configuration (alphabetical ordering)');
  return null; // Return null to signal that default behavior should be used
}

function getConfigurationSheetUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    return null;
  }
  
  return `${ss.getUrl()}#gid=${sheet.getSheetId()}`;
}


function writeConfigurationSheet(metadataResults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  // DELETE and recreate sheet to avoid any validation conflicts
  if (sheet) {
    ss.deleteSheet(sheet);
    SpreadsheetApp.flush();
  }
  
  sheet = ss.insertSheet(CONFIG_SHEET_NAME);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontFamily('Comfortaa');
  
  // ===== ROW 1: Title =====
  sheet.getRange(1, 1, 1, 20).merge()
    .setValue('JIRA Field Configuration')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('left');
  
  // ===== ROW 2: Subtitle =====
  sheet.getRange(2, 1, 1, 20).merge()
    .setValue('Field metadata from JIRA')
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('left');
  
  // ===== ROW 3: Timestamp =====
  const timestamp = new Date().toLocaleString();
  sheet.getRange(3, 1, 1, 20).merge()
    .setValue(`Last refreshed: ${timestamp}`)
    .setFontStyle('italic').setFontSize(9).setFontColor('#666666').setHorizontalAlignment('left');
  
  // ===== ROW 4: Default notice =====
  sheet.getRange(4, 1, 1, 20).merge()
    .setValue('Default will be Alphabetical order and all will be displayed')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('left');
  
  // ===== COLUMN LAYOUT =====
  // Column A: Value Stream/Org
  // Column B: Display in Slides
  // Column C: (blank spacing)
  // Column D: Portfolio Initiative  
  // Column E: Display in Slides
  // Column F: Order Display
  // Column G: (blank spacing)
  // Column H: Program Initiative
  // Column I: Display in Slides
  
  const fieldLayout = {
    'Value Stream/Org': { col: 1, hasDisplay: true, hasOrder: false },  // A, B
    'Portfolio Initiative': { col: 4, hasDisplay: true, hasOrder: true }, // D, E, F
    'Program Initiative': { col: 8, hasDisplay: true, hasOrder: false }   // H, I
  };
  
  const headerRow = 5;
  const dataStartRow = 6;
  
  // ===== ROW 5: Headers =====
  Object.keys(fieldLayout).forEach(fieldName => {
    const layout = fieldLayout[fieldName];
    const col = layout.col;
    
    // Field name header
    sheet.getRange(headerRow, col)
      .setValue(fieldName)
      .setFontWeight('bold')
      .setBackground('#9b7bb8')
      .setFontColor('white')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Display in Slides header
    sheet.getRange(headerRow, col + 1)
      .setValue('Display in Slides')
      .setFontWeight('bold')
      .setBackground('#9b7bb8')
      .setFontColor('white')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Order Display header (only for fields with ordering)
    if (layout.hasOrder) {
      sheet.getRange(headerRow, col + 2)
        .setValue('Order Display')
        .setFontWeight('bold')
        .setBackground('#9b7bb8')
        .setFontColor('white')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    }
  });
  
  // ===== DATA ROWS: Write values for each field =====
  Object.keys(fieldLayout).forEach(fieldName => {
    const layout = fieldLayout[fieldName];
    const col = layout.col;
    const metadata = metadataResults[fieldName];
    
    if (!metadata) {
      console.warn(`No metadata found for ${fieldName}`);
      return;
    }
    
    const allowedValues = metadata.allowedValues || [];
    
    // Create validation rules
    const displayRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'], true)
      .setAllowInvalid(false)
      .build();
    
    let orderRule = null;
    if (layout.hasOrder && allowedValues.length > 0) {
      const orderOptions = [];
      for (let i = 1; i <= allowedValues.length; i++) {
        orderOptions.push(i.toString());
      }
      orderRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(orderOptions, true)
        .setAllowInvalid(false)
        .build();
    }
    
    // Write data for each allowed value
    if (allowedValues.length === 0) {
      sheet.getRange(dataStartRow, col).setValue('(no values found)');
      sheet.getRange(dataStartRow, col + 1).setValue('N/A');
      if (layout.hasOrder) {
        sheet.getRange(dataStartRow, col + 2).setValue('N/A');
      }
    } else {
      allowedValues.forEach((value, index) => {
        const row = dataStartRow + index;
        
        // Column: Field value
        sheet.getRange(row, col)
          .setValue(value)
          .setVerticalAlignment('middle')
          .setWrap(false);
        
        // Column + 1: Display in Slides dropdown (default Yes)
        sheet.getRange(row, col + 1)
          .setValue('Yes')
          .setDataValidation(displayRule)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
        
        // Column + 2: Order Display dropdown (if hasOrder) - default to numerical order
        if (layout.hasOrder) {
          sheet.getRange(row, col + 2)
            .setValue((index + 1).toString()) // Default to 1, 2, 3, etc.
            .setDataValidation(orderRule)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        }
      });
    }
    
    // Set column widths
    sheet.setColumnWidth(col, 250);      // Field values column
    sheet.setColumnWidth(col + 1, 120);  // Display in Slides
    if (layout.hasOrder) {
      sheet.setColumnWidth(col + 2, 100); // Order Display
    }
  });
  
  // Set spacing column widths
  sheet.setColumnWidth(3, 20);  // Column C (spacing)
  sheet.setColumnWidth(7, 20);  // Column G (spacing)
  
  // ===== FINAL FORMATTING =====
  sheet.setFrozenRows(5);
  
  console.log(`Configuration sheet created with horizontal layout for ${Object.keys(fieldLayout).length} fields`);
}

function readSlideConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    console.warn('Configuration sheet not found. Using default settings.');
    return getDefaultConfiguration();
  }
  
  const config = {};
  const headerRow = 5;
  const dataStartRow = 6;
  
  // Get all headers to find field locations
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow - 1, 1, 1, lastCol).getValues()[0];
  
  let currentCol = 1;
  
  Object.keys(METADATA_FIELDS).forEach(fieldName => {
    const fieldConfig = METADATA_FIELDS[fieldName];
    const hasOrdering = fieldConfig.hasOrdering;
    const numCols = hasOrdering ? 3 : 2;
    
    // Find the column where this field starts
    let fieldStartCol = -1;
    for (let col = currentCol; col <= lastCol; col++) {
      const cellValue = sheet.getRange(headerRow - 1, col).getValue();
      if (cellValue === fieldName) {
        fieldStartCol = col;
        break;
      }
    }
    
    if (fieldStartCol === -1) {
      console.warn(`Field ${fieldName} not found in Configuration sheet`);
      currentCol += numCols + 1;
      return;
    }
    
    const values = [];
    const displayMap = {};
    const orderMap = {};
    
    // Read data for this field
    const lastRow = sheet.getLastRow();
    let row = dataStartRow;
    
    while (row <= lastRow) {
      const allowedValue = sheet.getRange(row, fieldStartCol).getValue();
      
      if (!allowedValue || allowedValue === '' || allowedValue === '(no values found)') {
        break; // Empty or placeholder means end of this field's data
      }
      
      const displayInSlides = sheet.getRange(row, fieldStartCol + 1).getValue();
      
      if (displayInSlides === 'Yes') {
        values.push(allowedValue);
        displayMap[allowedValue] = true;
        
        if (hasOrdering) {
          const orderValue = sheet.getRange(row, fieldStartCol + 2).getValue();
          if (orderValue && orderValue !== 'n/a') {
            orderMap[allowedValue] = parseInt(orderValue);
          }
        }
      }
      
      row++;
    }
    
    // Determine ordering
    let orderedValues = values;
    const hasCustomOrder = Object.keys(orderMap).length > 0;
    
    if (hasCustomOrder) {
      // Use custom ordering
      orderedValues = values.sort((a, b) => {
        const orderA = orderMap[a] || 999;
        const orderB = orderMap[b] || 999;
        return orderA - orderB;
      });
    } else {
      // Use alphabetical ordering (default)
      orderedValues = values.sort();
    }
    
    config[fieldName] = {
      values: orderedValues,
      displayMap: displayMap,
      orderMap: orderMap,
      hasCustomOrder: hasCustomOrder
    };
    
    currentCol = fieldStartCol + numCols + 1; // Move to next field group
  });
  
  console.log('Configuration loaded:', JSON.stringify(config, null, 2));
  return config;
}

function validateCustomOrdering() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) {
    return { valid: true, errors: [] };
  }
  
  const errors = [];
  let currentRow = 5;
  
  Object.keys(CONFIG_METADATA_FIELDS).forEach(fieldName => {
    const fieldConfig = CONFIG_METADATA_FIELDS[fieldName];
    
    if (!fieldConfig.hasOrdering) {
      return; // Skip fields without ordering capability
    }
    
    currentRow++; // Skip field header
    currentRow++; // Skip column headers
    
    const orderValues = [];
    const rowStart = currentRow;
    
    // Collect all order values for this field
    while (currentRow <= sheet.getLastRow()) {
      const allowedValue = sheet.getRange(currentRow, 1).getValue();
      
      if (!allowedValue || allowedValue === '') {
        break;
      }
      
      const displayInSlides = sheet.getRange(currentRow, 2).getValue();
      const orderValue = sheet.getRange(currentRow, 3).getValue();
      
      if (displayInSlides === 'Yes' && orderValue && orderValue !== 'n/a') {
        orderValues.push({ value: allowedValue, order: orderValue, row: currentRow });
      }
      
      currentRow++;
    }
    
    // Check if partial ordering exists
    const totalDisplayedRows = currentRow - rowStart;
    if (orderValues.length > 0 && orderValues.length < totalDisplayedRows) {
      errors.push(`${fieldName}: Custom ordering is incomplete. Either set order for all displayed values or none.`);
    }
    
    // Check for duplicate order numbers
    const orderNumbers = orderValues.map(ov => ov.order);
    const duplicates = orderNumbers.filter((item, index) => orderNumbers.indexOf(item) !== index);
    if (duplicates.length > 0) {
      errors.push(`${fieldName}: Duplicate order numbers found: ${duplicates.join(', ')}`);
    }
    
    currentRow++; // Skip the spacing row
  });
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}