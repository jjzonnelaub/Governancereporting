/**
 * Predictability Score Report System
 * 
 * This module provides functionality for:
 * 1. Generating Predictability Score metrics from Iteration 6 data
 * 2. Updating the Predictability Score sheet with allocation distributions
 * 3. Calculating ART Velocity metrics
 * 4. (Future) Generating Predictability slide decks
 * 
 * PREREQUISITES:
 * - PI must have an Iteration 6 tab (PI X - Iteration 6)
 * - Predictability Score sheet must exist in the spreadsheet
 */

// ===== CONSTANTS =====

// Predictability Score sheet structure
const PREDICTABILITY_SHEET_NAME = 'Predictability Score';
const PREDICTABILITY_TEMPLATE_ID = '1x08GcFnvUC5g7O-pIP8oyYCj-RztVU4asHBv90VFNh0';
const PREDICTABILITY_DESTINATION_FOLDER_ID = '1PUpzgiqNhc3pbrQqlUY1MtcKbftFLg9M';
const DEFERRAL_ISSUES_PER_SLIDE = 5;

// Table 1: ART Capacity (Rows 1-17) - MANUAL INPUT, WE SKIP THIS
const ART_CAPACITY_START_ROW = 1;
const ART_CAPACITY_END_ROW = 17;

// Table 2: Allocation Distribution (Rows 19-36)
const ALLOCATION_DIST_START_ROW = 20;
const ALLOCATION_DIST_HEADER_ROW = 21;  // Row with PI labels
const ALLOCATION_DIST_SUBHEADER_ROW = 23;  // Row with allocation types
const ALLOCATION_DIST_DATA_START_ROW = 24;  // First data row (AIMM)
const ALLOCATION_DIST_DATA_END_ROW = 37;  // Last data row (Xtract)

// Table 3: ART Velocity (Rows 38-54)
const ART_VELOCITY_START_ROW = 39;
const ART_VELOCITY_HEADER_ROW = 40;  // Row with PI labels
const ART_VELOCITY_DATA_START_ROW = 41;  // First data row (AIMM)
const ART_VELOCITY_DATA_END_ROW = 55;  // Last data row (Xtract)

// Table 4: Epic Status (Rows 57-74)
const EPIC_STATUS_START_ROW = 58;
const EPIC_STATUS_HEADER_ROW = 59;  // Row with PI labels
const EPIC_STATUS_SUBHEADER_ROW = 60;  // Row with status types
const EPIC_STATUS_DATA_START_ROW = 61;  // First data row (AIMM)
const EPIC_STATUS_DATA_END_ROW = 75;  // Last data row (Xtract)

// Each PI spans 3 columns in Epic Status table
const EPIC_STATUS_COLS_PER_PI = 3;
const EPIC_STATUS_TYPES = [
  'Committed',
  'Deferred',
  'Deferred taken into Planning'
];

// PI column mapping for Epic Status table
// PI 11 starts at column B (index 1)
const EPIC_STATUS_PI_COLUMN_START = {
  'PI 11': 1,   // Columns B-D (indices 1-3)
  'PI 12': 4,   // Columns E-G (indices 4-6)
  'PI 13': 7,   // Columns H-J (indices 7-9)
  'PI 14': 10,  // Columns K-M (indices 10-12)
  'PI 15': 13   // Columns N-P (indices 13-15)
};

// Column structure for Allocation Distribution table
// Each PI spans 5 columns (Product-Feature, Product-Compliance, KLO, Quality, Tech / Platform)
const ALLOCATIONS_PER_PI = 5;
const ALLOCATION_TYPES = [
  'Product - Feature',
  'Product - Compliance', 
  'KLO',
  'Quality',
  'Tech / Platform'
];

// PI column mapping for Allocation Distribution
// PI 11 starts at column B (index 1)
const PI_COLUMN_START = {
  'PI 11': 1,   // Columns B-F (indices 1-5)
  'PI 12': 6,   // Columns G-K (indices 6-10)
  'PI 13': 11,  // Columns L-P (indices 11-15)
  'PI 14': 16,  // Columns Q-U (indices 16-20)
  'PI 15': 21   // Columns V-Z (indices 21-25)
};

// Iteration 6 sheet column indices (0-based)
const ITER6_COL_KEY = 0;                    // Column A: Key
const ITER6_COL_ISSUE_TYPE = 2;             // Column C: Issue Type
const ITER6_COL_VALUE_STREAM = 10;          // Column K: Value Stream/Org
const ITER6_COL_PI_COMMITMENT = 12;         // Column M: PI Commitment
const ITER6_COL_ALLOCATION = 15;            // Column P: Allocation
const ITER6_COL_PI_OBJECTIVE_STATUS = 16;   // Column Q: PI Objective Status
const ITER6_COL_STORY_POINT_COMPLETION = 20; // Column U: Story Point Completion
const ITER6_COL_BUSINESS_VALUE = 24;        // Column Y: Business Value
const ITER6_COL_ACTUAL_VALUE = 25;          // Column Z: Actual Value

// Table 5: Program Predictability (Rows 76-93)
const PROGRAM_PREDICTABILITY_START_ROW = 77;
const PROGRAM_PREDICTABILITY_HEADER_ROW = 78;  // Row with PI labels
const PROGRAM_PREDICTABILITY_SUBHEADER_ROW = 79;  // Row with value types
const PROGRAM_PREDICTABILITY_DATA_START_ROW = 80;  // First data row (AIMM)
const PROGRAM_PREDICTABILITY_DATA_END_ROW = 94;  // Last data row (Xtract)

// Each PI spans 3 columns in Program Predictability table
const PROGRAM_PREDICTABILITY_COLS_PER_PI = 3;
const PROGRAM_PREDICTABILITY_TYPES = [
  'Business Value',
  'Actual Value',
  'PI Score'
];

// PI column mapping for Program Predictability table
// PI 11 starts at column B (index 1)
const PROGRAM_PREDICTABILITY_PI_COLUMN_START = {
  'PI 11': 1,   // Columns B-D (indices 1-3)
  'PI 12': 4,   // Columns E-G (indices 4-6)
  'PI 13': 7,   // Columns H-J (indices 7-9)
  'PI 14': 10,  // Columns K-M (indices 10-12)
  'PI 15': 13   // Columns N-P (indices 13-15)
};
const PREDICTABILITY_FONT = 'Comfortaa';
const PREDICTABILITY_FONT_SIZE = 11;
const PREDICTABILITY_FONT_COLOR = '#000000';

const PREDICTABILITY_VALUE_STREAMS = [
  'AIMM',                       // Row 117
  'Cloud Foundation Services',  // Row 118
  'Cloud Operations',           // Row 119
  'Data Platform Engineering',  // Row 120
  'EMA Clinical',               // Row 121
  'EMA RaC',                    // Row 122
  'Fusion and Conversions',     // Row 123
  'MMGI',                       // Row 124
  'MMGI-Cloud',                 // Row 125
  'MMPM',                       // Row 126
  'Patient Collaboration',      // Row 127
  'Platform Engineering',       // Row 128
  'RCM Genie',                  // Row 129
  'Shared Services Platform',   // Row 130
  'Xtract'                      // Row 131
];

// ===== HELPER FUNCTIONS =====

/**
 * Gets list of PIs that have Iteration 6 tabs
 * @returns {Array<Object>} Array of PI objects with number and sheet name
 */
function getAvailablePIsForPredictability() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const availablePIs = [];
  
  // Look for sheets named "PI X - Iteration 6"
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const match = sheetName.match(/^PI (\d+) - Iteration 6$/);
    if (match) {
      const piNumber = parseInt(match[1]);
      availablePIs.push({
        number: piNumber,
        label: `PI ${piNumber}`,
        sheetName: sheetName,
        sheet: sheet
      });
    }
  });
  
  // Sort by PI number
  availablePIs.sort((a, b) => a.number - b.number);
  
  console.log(`Found ${availablePIs.length} PIs with Iteration 6 tabs:`, 
              availablePIs.map(pi => pi.label).join(', '));
  
  return availablePIs;
}

/**
 * Checks if Predictability Score sheet exists
 * @returns {Sheet|null} The Predictability Score sheet or null if not found
 */
function getPredictabilityScoreSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(PREDICTABILITY_SHEET_NAME);
}

/**
 * Gets the column index for a specific PI and allocation type
 * @param {string} piLabel - e.g., "PI 11"
 * @param {string} allocationType - e.g., "Product - Feature"
 * @returns {number} 0-based column index
 */
function getColumnIndexForAllocation(piLabel, allocationType) {
  const piStartCol = PI_COLUMN_START[piLabel];
  if (piStartCol === undefined) {
    throw new Error(`Unknown PI: ${piLabel}`);
  }
  
  const allocationIndex = ALLOCATION_TYPES.indexOf(allocationType);
  if (allocationIndex === -1) {
    throw new Error(`Unknown allocation type: ${allocationType}`);
  }
  
  return piStartCol + allocationIndex;
}

/**
 * Gets the row index for a specific value stream in the Allocation Distribution table
 * @param {Sheet} sheet - The Predictability Score sheet
 * @param {string} valueStream - The value stream name
 * @returns {number|null} 1-based row index or null if not found
 */
function getRowIndexForValueStream(sheet, valueStream) {
  // Read value stream names from column A (rows 22-36)
  const vsRange = sheet.getRange(ALLOCATION_DIST_DATA_START_ROW, 1, 
                                 ALLOCATION_DIST_DATA_END_ROW - ALLOCATION_DIST_DATA_START_ROW + 1, 1);
  const vsValues = vsRange.getValues();
  
  for (let i = 0; i < vsValues.length; i++) {
    if (vsValues[i][0] === valueStream) {
      return ALLOCATION_DIST_DATA_START_ROW + i;  // 1-based row number
    }
  }
  
  console.warn(`Value stream not found in Predictability Score sheet: ${valueStream}`);
  return null;
}

// ===== DIALOG HTML =====

function getPredictabilityDialogHTML() {
  const availablePIs = getAvailablePIsForPredictability();
  
  if (availablePIs.length === 0) {
    return `
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body { 
              font-family: 'Google Sans', 'Roboto', Arial, sans-serif; 
              padding: 20px; 
              text-align: center; 
            }
            .error-icon { font-size: 48px; color: #d93025; margin-bottom: 16px; }
            h2 { color: #d93025; margin-bottom: 16px; }
            p { color: #5f6368; line-height: 1.6; }
          </style>
        </head>
        <body>
          <div class="error-icon">⚠️</div>
          <h2>No Iteration 6 Tabs Found</h2>
          <p>Predictability reports require an Iteration 6 tab for the selected PI.</p>
          <p>Please generate an Iteration 6 report first, then try again.</p>
        </body>
      </html>
    `;
  }
  
  const piOptions = availablePIs.map(pi => 
    `<option value="${pi.number}">${pi.label}</option>`
  ).join('');
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { 
            font-family: 'Google Sans', 'Roboto', Arial, sans-serif; 
            padding: 20px; 
            background: #fff; 
          }
          .container { max-width: 450px; margin: 0 auto; }
          h2 { 
            color: #1a73e8; 
            margin-bottom: 8px; 
            font-size: 20px; 
            font-weight: 500; 
          }
          .subtitle { 
            color: #5f6368; 
            font-size: 14px; 
            margin-bottom: 24px; 
          }
          .section { 
            background: #f8f9fa; 
            border-radius: 8px; 
            padding: 16px; 
            margin-bottom: 20px; 
          }
          .section-title { 
            color: #3c4043; 
            font-weight: 500; 
            margin-bottom: 12px; 
            font-size: 14px; 
            text-transform: uppercase; 
            letter-spacing: 0.5px; 
          }
          .input-group { margin-bottom: 16px; }
          .pi-select { 
            width: 100%; 
            padding: 12px; 
            font-size: 16px; 
            border: 2px solid #dadce0; 
            border-radius: 8px; 
            transition: border-color 0.2s; 
            background: white; 
          }
          .pi-select:focus { 
            outline: none; 
            border-color: #1a73e8; 
          }
          .info-box {
            background: #e3f2fd;
            border-left: 3px solid #1976d2;
            padding: 12px;
            margin: 16px 0;
            border-radius: 4px;
          }
          .info-box p {
            color: #1565c0;
            font-size: 13px;
            margin: 4px 0;
          }
          .info-box ul {
            color: #1565c0;
            font-size: 13px;
            margin: 8px 0 0 20px;
          }
          .button-container { 
            display: flex; 
            gap: 12px; 
            justify-content: flex-end; 
            margin-top: 24px; 
          }
          button { 
            padding: 10px 24px; 
            font-size: 14px; 
            border-radius: 6px; 
            border: none; 
            cursor: pointer; 
            font-weight: 500; 
            transition: all 0.2s; 
            font-family: inherit; 
          }
          .primary-button { 
            background: #1a73e8; 
            color: white; 
          }
          .primary-button:hover:not(:disabled) { 
            background: #1557b0; 
            box-shadow: 0 1px 3px rgba(0,0,0,0.2); 
          }
          .primary-button:disabled { 
            background: #94c1f5; 
            cursor: not-allowed; 
          }
          .secondary-button { 
            background: #ffffff; 
            color: #5f6368; 
            border: 1px solid #dadce0; 
          }
          .secondary-button:hover { background: #f8f9fa; }
          .loading { 
            display: none; 
            text-align: center; 
            padding: 20px; 
          }
          .spinner { 
            border: 3px solid #f3f3f3; 
            border-top: 3px solid #1a73e8; 
            border-radius: 50%; 
            width: 40px; 
            height: 40px; 
            animation: spin 1s linear infinite; 
            margin: 0 auto 12px; 
          }
          @keyframes spin { 
            0% { transform: rotate(0deg); } 
            100% { transform: rotate(360deg); } 
          }
          .error { 
            color: #d93025; 
            font-size: 12px; 
            margin-top: 8px; 
            display: none; 
          }
          .error.show { display: block; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Update Predictability Score</h2>
          <p class="subtitle">Calculate allocation distribution and ART velocity metrics</p>
          
          <div class="section">
            <div class="section-title">Select Program Increment</div>
            <div class="input-group">
              <select id="piSelect" class="pi-select">
                <option value="">-- Select PI --</option>
                ${piOptions}
              </select>
            </div>
          </div>
          
          <div class="info-box">
            <p><strong>📊 What this will update:</strong></p>
            <ul>
              <li>Allocation Distribution table (by Value Stream & Allocation)</li>
              <li>ART Velocity table (total story points per Value Stream)</li>
              <li>Epic Status table (Committed, Deferred, Deferred taken into Planning counts)</li>
              <li>Program Predictability table (Business Value, Actual Value, PI Score)</li>
            </ul>
            <p style="margin-top: 8px;"><strong>📍 Data Source:</strong> Iteration 6 tab for selected PI</p>
          </div>
          
          <div class="button-container">
            <button type="button" class="secondary-button" onclick="google.script.host.close()">
              Cancel
            </button>
            <button type="button" class="primary-button" id="generateBtn" onclick="handleGenerate()">
              Update Predictability Score
            </button>
          </div>
          
          <div class="loading" id="loading">
            <div class="spinner"></div>
            <p>Processing data and updating sheet...</p>
          </div>
          
          <div class="error" id="error"></div>
        </div>
        
        <script>
          function handleGenerate() {
            const piNumber = document.getElementById('piSelect').value;
            
            if (!piNumber) {
              showError('Please select a Program Increment');
              return;
            }
            
            // Show loading state
            document.getElementById('loading').style.display = 'block';
            document.getElementById('generateBtn').disabled = true;
            document.getElementById('error').classList.remove('show');
            
            // Call the backend function
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .updatePredictabilityScore(parseInt(piNumber));
          }
          
          function onSuccess(result) {
            google.script.host.close();
          }
          
          function onFailure(error) {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('generateBtn').disabled = false;
            showError('Error: ' + error.message);
          }
          
          function showError(message) {
            const errorDiv = document.getElementById('error');
            errorDiv.textContent = message;
            errorDiv.classList.add('show');
          }
        </script>
      </body>
    </html>
  `;
}

// ===== MAIN PROCESSING FUNCTION =====

/**
 * Main function to update Predictability Score sheet with data from Iteration 6
 * @param {number} piNumber - The PI number to process
 */
function updatePredictabilityScore(piNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  console.log(`\n========================================`);
  console.log(`UPDATING PREDICTABILITY SCORE FOR PI ${piNumber}`);
  console.log(`========================================`);
  
  // Step 1: Validate inputs
  const predictabilitySheet = getPredictabilityScoreSheet();
  if (!predictabilitySheet) {
    ui.alert('Error', `Sheet "${PREDICTABILITY_SHEET_NAME}" not found. Please create it first.`, ui.ButtonSet.OK);
    return;
  }
  
  const piLabel = `PI ${piNumber}`;
  const iter6SheetName = `${piLabel} - Iteration 6`;
  const iter6Sheet = ss.getSheetByName(iter6SheetName);
  
  if (!iter6Sheet) {
    ui.alert('Error', `Sheet "${iter6SheetName}" not found. Please generate Iteration 6 report first.`, ui.ButtonSet.OK);
    return;
  }
  
  // Check if this PI is supported in the Predictability Score sheet
  if (PI_COLUMN_START[piLabel] === undefined) {
    ui.alert('Error', `${piLabel} is not yet configured in the Predictability Score sheet structure.`, ui.ButtonSet.OK);
    return;
  }
  
  ss.toast(`Processing Iteration 6 data for ${piLabel}...`, '📊 Analyzing', 10);
  
  // Step 1.5: Clear existing data for this PI across all tables
  ss.toast(`Clearing existing PI ${piNumber} data...`, '🧹 Clearing', 5);
  clearPIColumnsInAllTables(predictabilitySheet, piNumber);
  
  // Step 2: Update Allocation Distribution table (Table 2)
  ss.toast('Updating Allocation Distribution table...', '✍️ Writing', 10);
  populateTable2_AllocationDistribution(predictabilitySheet, piNumber);
  
  // Step 3: Update ART Velocity table (Table 3)
  ss.toast('Updating ART Velocity table...', '✍️ Writing', 10);
  populateTable3_ARTVelocity(predictabilitySheet, piNumber);
  
  // Step 4: Update Epic Status table (Table 4)
  ss.toast('Updating Epic Status table...', '✍️ Writing', 10);
  populateTable4_EpicStatus(predictabilitySheet, piNumber);
  
  // Step 5: Update Program Predictability table (Table 5)
  ss.toast('Updating Program Predictability table...', '✍️ Writing', 10);
  populateTable5_ProgramPredictability(predictabilitySheet, piNumber);
  
  // Step 6: Update Objective Status table (Table 6)
  ss.toast('Updating Objective Status table...', '✍️ Writing', 10);
  populateTable6_ObjectiveStatus(predictabilitySheet, piNumber);
  
  // Step 7: Update PI Score Summary table (Table 7)
  ss.toast('Updating PI Score Summary table...', '✍️ Writing', 10);
  populateTable7_PIScoreSummary(predictabilitySheet, piNumber);
  
  // Success!
  console.log('\n✓ Successfully updated Predictability Score for ' + piLabel);
  ss.toast('✅ Predictability Score updated for ' + piLabel, 'Success', 5);
  ui.alert('Success', 'Predictability Score has been updated for ' + piLabel + '!', ui.ButtonSet.OK);
}

/**
 * Clears all data columns for a specific PI across all predictability tables
 * This ensures a clean slate before regenerating data
 * @param {Sheet} sheet - The Predictability Score sheet
 * @param {number} piNumber - The PI number to clear
 */
function clearPIColumnsInAllTables(sheet, piNumber) {
  console.log(`\n🧹 Clearing PI ${piNumber} data from all tables...`);
  
  const BASE_PI = 11;
  const piIndex = piNumber - BASE_PI;
  
  // Get value stream count from Configuration Template
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Configuration Template');
  let numValueStreams = 15; // Default
  
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    let count = 0;
    for (let i = 5; i < configData.length; i++) {
      if (configData[i][0] && configData[i][0].toString().trim() !== '') {
        count++;
      }
    }
    if (count > 0) numValueStreams = count;
  }
  
  console.log(`   Value streams to clear: ${numValueStreams}`);
  
  // Table 2: Allocation Distribution (5 columns per PI, rows 23-37)
  const table2StartCol = 2 + (piIndex * 5); // 5 columns per PI
  const table2StartRow = 23;
  console.log(`   Table 2 (Allocation): Columns ${columnNumberToLetter(table2StartCol)}-${columnNumberToLetter(table2StartCol + 4)}, Rows ${table2StartRow}-${table2StartRow + numValueStreams - 1}`);
  sheet.getRange(table2StartRow, table2StartCol, numValueStreams, 5).clearContent();
  
  // Table 3: ART Velocity (1 column per PI, rows 41-55)
  const table3StartCol = 2 + piIndex; // 1 column per PI
  const table3StartRow = 41;
  console.log(`   Table 3 (Velocity): Column ${columnNumberToLetter(table3StartCol)}, Rows ${table3StartRow}-${table3StartRow + numValueStreams - 1}`);
  sheet.getRange(table3StartRow, table3StartCol, numValueStreams, 1).clearContent();
  
  // Table 4: Epic Status (3 columns per PI, rows 61-75)
  const table4StartCol = 2 + (piIndex * 3); // 3 columns per PI
  const table4StartRow = 61;
  console.log(`   Table 4 (Epic Status): Columns ${columnNumberToLetter(table4StartCol)}-${columnNumberToLetter(table4StartCol + 2)}, Rows ${table4StartRow}-${table4StartRow + numValueStreams - 1}`);
  sheet.getRange(table4StartRow, table4StartCol, numValueStreams, 3).clearContent();
  
  // Table 5: Program Predictability (3 columns per PI, rows 80-94)
  const table5StartCol = 2 + (piIndex * 3); // 3 columns per PI
  const table5StartRow = 80;
  console.log(`   Table 5 (Predictability): Columns ${columnNumberToLetter(table5StartCol)}-${columnNumberToLetter(table5StartCol + 2)}, Rows ${table5StartRow}-${table5StartRow + numValueStreams - 1}`);
  sheet.getRange(table5StartRow, table5StartCol, numValueStreams, 3).clearContent();
  
  // Table 6: Objective Status (5 columns per PI, rows 99-113)
  const table6StartCol = 2 + (piIndex * 5); // 5 columns per PI
  const table6StartRow = 99;
  console.log(`   Table 6 (Objective Status): Columns ${columnNumberToLetter(table6StartCol)}-${columnNumberToLetter(table6StartCol + 4)}, Rows ${table6StartRow}-${table6StartRow + numValueStreams - 1}`);
  sheet.getRange(table6StartRow, table6StartCol, numValueStreams, 5).clearContent();
  
  // Table 7: PI Score Summary (1 column per PI, rows 117-131)
  const table7StartCol = 2 + piIndex; // 1 column per PI
  const table7StartRow = 117;
  console.log(`   Table 7 (PI Score Summary): Column ${columnNumberToLetter(table7StartCol)}, Rows ${table7StartRow}-${table7StartRow + numValueStreams - 1}`);
  sheet.getRange(table7StartRow, table7StartCol, numValueStreams, 1).clearContent();
  
  console.log(`✅ Cleared all PI ${piNumber} data\n`);
}

// ===== DATA READING FUNCTIONS =====

/**
 * Reads Epic data from Iteration 6 sheet
 * @param {Sheet} sheet - The Iteration 6 sheet
 * @returns {Array<Object>} Array of epic records
 */
function readIteration6Data(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 5) {  // Headers are in row 4, data starts at row 5
    console.warn('No data found in Iteration 6 sheet');
    return [];
  }
  
  // Read all data including headers
  const allData = sheet.getRange(4, 1, lastRow - 3, lastCol).getValues();
  const headers = allData[0];
  const dataRows = allData.slice(1);
  
  console.log(`Read ${dataRows.length} data rows from Iteration 6 sheet`);
  console.log(`Headers: ${headers.slice(0, 5).join(', ')}...`);
  
  // Parse into objects - only include Epics with valid data
  const records = [];
  dataRows.forEach(function(row, index) {
    var key = row[ITER6_COL_KEY];
    var issueType = row[ITER6_COL_ISSUE_TYPE];
    var valueStream = row[ITER6_COL_VALUE_STREAM];
    var piCommitment = row[ITER6_COL_PI_COMMITMENT];
    var allocation = row[ITER6_COL_ALLOCATION];
    var piObjectiveStatus = row[ITER6_COL_PI_OBJECTIVE_STATUS];
    var storyPointCompletion = row[ITER6_COL_STORY_POINT_COMPLETION];
    var businessValue = row[ITER6_COL_BUSINESS_VALUE];
    var actualValue = row[ITER6_COL_ACTUAL_VALUE];
    
    // Only process Epics with key and value stream
    if (issueType === 'Epic' && key && valueStream) {
      records.push({
        key: key,
        valueStream: valueStream,
        piCommitment: piCommitment || '',  // Can be blank
        allocation: allocation,
        piObjectiveStatus: piObjectiveStatus || '',  // Can be blank
        storyPointCompletion: parseFloat(storyPointCompletion) || 0,
        businessValue: parseFloat(businessValue) || 0,
        actualValue: parseFloat(actualValue) || 0
      });
    }
  });
  
  console.log('Filtered to ' + records.length + ' Epic records');
  
  return records;
}

// ===== CALCULATION FUNCTIONS =====

/**
 * Calculates allocation metrics grouped by Value Stream and Allocation type
 * @param {Array<Object>} records - Array of epic records from Iteration 6
 * @returns {Object} Nested object: { valueStream: { allocationType: totalPoints } }
 */
function calculateAllocationMetrics(records) {
  const metrics = {};
  
  records.forEach(function(record) {
    const vs = record.valueStream;
    const allocation = record.allocation;
    const points = record.storyPointCompletion;
    
    // Initialize value stream if needed
    if (!metrics[vs]) {
      metrics[vs] = {};
      // Initialize all allocation types to 0
      ALLOCATION_TYPES.forEach(function(type) {
        metrics[vs][type] = 0;
      });
    }
    
    // Add points to the appropriate allocation type
    if (metrics[vs][allocation] !== undefined) {
      metrics[vs][allocation] += points;
    } else {
      console.warn('Unknown allocation type "' + allocation + '" for ' + record.key + ' - skipping');
    }
  });
  
  // Log summary
  console.log('\n=== ALLOCATION METRICS SUMMARY ===');
  
  var sortedValueStreams = Object.keys(metrics).sort();
  for (var i = 0; i < sortedValueStreams.length; i++) {
    var vs = sortedValueStreams[i];
    var vsMetrics = metrics[vs];
    
    // Calculate total points for this value stream
    var total = 0;
    var allocationNames = Object.keys(vsMetrics);
    for (var j = 0; j < allocationNames.length; j++) {
      total += vsMetrics[allocationNames[j]];
    }
    
    var roundedTotal = Math.round(total);
    console.log(vs + ': ' + total.toFixed(2) + ' -> ' + roundedTotal + ' total points (rounded)');
    
    // Log individual allocations
    for (var k = 0; k < allocationNames.length; k++) {
      var allocName = allocationNames[k];
      var points = vsMetrics[allocName];
      if (points > 0) {
        var roundedPoints = Math.round(points);
        console.log('  - ' + allocName + ': ' + points.toFixed(2) + ' -> ' + roundedPoints);
      }
    }
  }
  
  return metrics;
}

/**
 * Calculates epic status counts grouped by Value Stream and PI Commitment status
 * @param {Array<Object>} records - Array of epic records from Iteration 6
 * @returns {Object} Nested object: { valueStream: { statusType: count } }
 */
function calculateEpicStatusCounts(records) {
  const counts = {};
  
  records.forEach(function(record) {
    var vs = record.valueStream;
    var piCommitment = record.piCommitment;
    var allocation = record.allocation;
    
    // EXCLUDE epics with Quality or KLO allocations
    if (allocation === 'Quality' || allocation === 'KLO') {
      return;  // Skip this record
    }
    
    // Initialize value stream if needed
    if (!counts[vs]) {
      counts[vs] = {
        'Committed': 0,
        'Deferred': 0,
        'Deferred taken into Planning': 0
      };
    }
    
    // Categorize based on PI Commitment value
    // Skip if PI Commitment is "Traded" or blank
    if (!piCommitment || piCommitment === 'Traded') {
      return;  // Skip this record
    }
    
    // Committed: "Committed" OR "Committed After Plan"
    if (piCommitment === 'Committed' || piCommitment === 'Committed After Plan') {
      counts[vs]['Committed']++;
    }
    // Deferred: "Deferred" OR "Canceled"
    else if (piCommitment === 'Deferred' || piCommitment === 'Canceled') {
      counts[vs]['Deferred']++;
    }
    // Deferred taken into Planning: TBD - set to 0 for now
    // When rules are defined, add logic here
  });
  
  // Log summary
  console.log('\n=== EPIC STATUS COUNTS SUMMARY ===');
  console.log('(Excludes epics with Quality and KLO allocations)');
  
  var sortedValueStreams = Object.keys(counts).sort();
  for (var i = 0; i < sortedValueStreams.length; i++) {
    var vs = sortedValueStreams[i];
    var vsCounts = counts[vs];
    
    var totalEpics = vsCounts['Committed'] + vsCounts['Deferred'] + vsCounts['Deferred taken into Planning'];
    console.log(vs + ': ' + totalEpics + ' total epics');
    console.log('  - Committed: ' + vsCounts['Committed']);
    console.log('  - Deferred: ' + vsCounts['Deferred']);
    console.log('  - Deferred taken into Planning: ' + vsCounts['Deferred taken into Planning'] + ' (TBD)');
  }
  
  return counts;
}

/**
 * Calculates program predictability metrics (Business Value, Actual Value, PI Score)
 * @param {Array<Object>} records - Array of epic records from Iteration 6
 * @returns {Object} Nested object: { valueStream: { businessValue: sum, actualValue: sum, piScore: percent } }
 */
function calculateProgramPredictability(records) {
  const metrics = {};
  
  // Valid PI Commitment values for inclusion
  const VALID_PI_COMMITMENTS = ['Committed', 'Committed After Plan', 'Deferred'];
  // PI Commitment values to exclude
  const EXCLUDED_PI_COMMITMENTS = ['Canceled', 'Traded'];
  
  records.forEach(function(record) {
    var vs = record.valueStream;
    var allocation = record.allocation;
    var piObjectiveStatus = record.piObjectiveStatus;
    var piCommitment = record.piCommitment;
    var businessValue = record.businessValue;
    var actualValue = record.actualValue;
    
    // FILTER 1: EXCLUDE epics with Quality or KLO allocations
    if (allocation === 'Quality' || allocation === 'KLO') {
      return;  // Skip this record
    }
    
    // FILTER 2: EXCLUDE epics with PI Commitment = "Canceled" or "Traded"
    // Only INCLUDE if PI Commitment = "Committed", "Committed After Plan", or "Deferred"
    if (!piCommitment || !VALID_PI_COMMITMENTS.includes(piCommitment)) {
      return;  // Skip this record (covers Canceled, Traded, blank, and any other values)
    }
    
    // Initialize value stream if needed
    if (!metrics[vs]) {
      metrics[vs] = {
        businessValue: 0,
        actualValue: 0,
        piScore: 0
      };
    }
    
    // ALWAYS sum Business Value (for all non-Quality/KLO epics with valid PI Commitment)
    metrics[vs].businessValue += businessValue;
    
    // ONLY sum Actual Value when PI Objective Status = "Met"
    if (piObjectiveStatus === 'Met') {
      metrics[vs].actualValue += actualValue;
    }
  });
  
  // Calculate PI Score for each value stream
  // PI Score = (Actual Value / Business Value) * 100
  var sortedValueStreams = Object.keys(metrics).sort();
  for (var i = 0; i < sortedValueStreams.length; i++) {
    var vs = sortedValueStreams[i];
    var vsMetrics = metrics[vs];
    
    if (vsMetrics.businessValue > 0) {
      vsMetrics.piScore = (vsMetrics.actualValue / vsMetrics.businessValue) * 100;
    } else {
      vsMetrics.piScore = 0;  // Avoid division by zero
    }
  }
  
  // Log summary
  console.log('\n=== PROGRAM PREDICTABILITY SUMMARY ===');
  console.log('(Excludes epics with Quality/KLO allocations)');
  console.log('(Excludes epics with PI Commitment = Canceled or Traded)');
  console.log('(Actual Value only includes epics with PI Objective Status = "Met")');
  
  for (var i = 0; i < sortedValueStreams.length; i++) {
    var vs = sortedValueStreams[i];
    var vsMetrics = metrics[vs];
    
    console.log(vs + ':');
    console.log('  - Business Value: ' + vsMetrics.businessValue.toFixed(2));
    console.log('  - Actual Value: ' + vsMetrics.actualValue.toFixed(2));
    console.log('  - PI Score: ' + vsMetrics.piScore.toFixed(1) + '%');
  }
  
  return metrics;
}

// ===== SHEET UPDATE FUNCTIONS =====

function populateTable2_AllocationDistribution(sheet, piNumber) {
  console.log(`\n========================================`);
  console.log(`TABLE 2: ALLOCATION DISTRIBUTION - PI ${piNumber}`);
  console.log(`========================================`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = `PI ${piNumber} - Iteration 6`;
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    console.error(`❌ Source sheet not found: ${sourceSheetName}`);
    throw new Error(`Cannot find required sheet: ${sourceSheetName}`);
  }
  
  // Table structure - CORRECTED TO MATCH ACTUAL SHEET
  // Row 20: "Allocation Distribution" header
  // Row 21: "Value Stream/Org" header + PI headers (each PI spans 5 columns)
  // Row 22: Allocation type headers (Quality, Tech/Platform, Product-Feature, Product-Compliance, KLO, Quality, Tech/Platform)
  // Row 23+: Data rows (starting with AIMM at row 23)
  const TABLE_START_ROW = 20;
  const PI_HEADER_ROW = 21;
  const ALLOCATION_HEADER_ROW = 22;
  const FIRST_DATA_ROW = 23; // ✅ FIXED: First data row is 23 (AIMM)
  const BASE_PI = 11;
  const COLUMNS_PER_PI = 5; // Product-Feature, Product-Compliance, KLO, Quality, Tech/Platform
  
  const piIndex = piNumber - BASE_PI;
  const startCol = 2 + (piIndex * COLUMNS_PER_PI); // Column B = index 2 (1-based)
  
  console.log(`📊 PI ${piNumber} columns: ${columnNumberToLetter(startCol)}-${columnNumberToLetter(startCol + COLUMNS_PER_PI - 1)}`);
  console.log(`📊 Column mapping for PI ${piNumber}:`);
  console.log(`   ${columnNumberToLetter(startCol)}: Product-Feature`);
  console.log(`   ${columnNumberToLetter(startCol + 1)}: Product-Compliance`);
  console.log(`   ${columnNumberToLetter(startCol + 2)}: KLO`);
  console.log(`   ${columnNumberToLetter(startCol + 3)}: Quality`);
  console.log(`   ${columnNumberToLetter(startCol + 4)}: Tech/Platform`);
  
  // Read source data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[3]; // Row 4 (0-indexed as 3)
  
  const vsCol = sourceHeaders.indexOf('Value Stream/Org');
  const allocationCol = sourceHeaders.indexOf('Allocation');
  const statusCol = sourceHeaders.indexOf('Status');
  const issueTypeCol = sourceHeaders.indexOf('Issue Type');
  const storyPointsCol = sourceHeaders.indexOf('Story Point Completion');
  
  if (vsCol === -1 || allocationCol === -1 || statusCol === -1 || issueTypeCol === -1 || storyPointsCol === -1) {
    throw new Error('Required columns not found in source sheet');
  }
  
  // Get value streams from Configuration Template
  const configSheet = ss.getSheetByName('Configuration Template');
  if (!configSheet) {
    throw new Error('Configuration Template sheet not found');
  }
  
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0];
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  console.log(`📍 Value Streams (${valueStreams.length}): ${valueStreams.join(', ')}`);
  
  // Process each value stream
  valueStreams.forEach((valueStream, index) => {
    console.log(`\n🔍 Processing: ${valueStream}`);
    
    const metrics = {
      'Product - Feature': 0,
      'Product - Compliance': 0,
      'KLO': 0,
      'Quality': 0,
      'Tech / Platform': 0
    };
    
    let totalEpicsChecked = 0;
    let closedEpicsProcessed = 0;
    
    // Scan source data (start from row 5, index 4)
    for (let i = 4; i < sourceData.length; i++) {
      const row = sourceData[i];
      const rowVS = row[vsCol];
      const rowAllocation = row[allocationCol];
      const rowStatus = row[statusCol];
      const rowIssueType = row[issueTypeCol];
      const rowStoryPoints = row[storyPointsCol];
      
      // Only process CLOSED EPICS for the current value stream
      if (rowIssueType !== 'Epic') continue;
      if (rowVS !== valueStream) continue;
      
      totalEpicsChecked++;
      
      const normalizedStatus = (rowStatus || '').toString().trim().toLowerCase();
      
      // ⚠️ CRITICAL FILTER: Only include if Status = "Closed"
      if (normalizedStatus !== 'closed' && normalizedStatus !== 'pending acceptance') continue;
      
      closedEpicsProcessed++;
      
      const storyPoints = parseFloat(rowStoryPoints) || 0;
      const allocation = (rowAllocation || '').toString().trim();
      
      // Map allocation to metrics (handle variations)
      if (allocation === 'Product - Feature') {
        metrics['Product - Feature'] += storyPoints;
      } else if (allocation === 'Product - Compliance' || allocation === 'Compliance') {
        metrics['Product - Compliance'] += storyPoints;
      } else if (allocation === 'KLO' || allocation === 'KLO / Quality') {
        metrics['KLO'] += storyPoints;
      } else if (allocation === 'Quality') {
        metrics['Quality'] += storyPoints;
      } else if (allocation === 'Tech / Platform' || allocation === 'Tech/Platform') {
        metrics['Tech / Platform'] += storyPoints;
      }
    }
    
    console.log(`   Total epics checked: ${totalEpicsChecked}`);
    console.log(`   Closed epics processed: ${closedEpicsProcessed}`);
    console.log(`   Product-Feature: ${metrics['Product - Feature'].toFixed(2)} pts`);
    console.log(`   Product-Compliance: ${metrics['Product - Compliance'].toFixed(2)} pts`);
    console.log(`   KLO: ${metrics['KLO'].toFixed(2)} pts`);
    console.log(`   Quality: ${metrics['Quality'].toFixed(2)} pts`);
    console.log(`   Tech/Platform: ${metrics['Tech / Platform'].toFixed(2)} pts`);
    
    // Write to sheet
    const rowNum = FIRST_DATA_ROW + index;
    
    // ✅ FIXED: Column offsets to match actual sheet structure (Row 22)
    // Column order: Product-Feature, Product-Compliance, KLO, Quality, Tech/Platform
    const colOffsets = {
      'Product - Feature': 0,
      'Product - Compliance': 1,
      'KLO': 2,
      'Quality': 3,
      'Tech / Platform': 4
    };
    
    // Write ALL 5 allocations to their correct columns
    ['Product - Feature', 'Product - Compliance', 'KLO', 'Quality', 'Tech / Platform'].forEach(allocation => {
      const cellCol = startCol + colOffsets[allocation];
      const cell = sheet.getRange(rowNum, cellCol);
      const roundedValue = Math.round(metrics[allocation]); // Round to whole number
      
      console.log(`   Writing ${allocation}: ${roundedValue} to row ${rowNum}, col ${cellCol} (${columnNumberToLetter(cellCol)})`);
      
      cell.setValue(roundedValue)
        .setFontFamily('Comfortaa')
        .setFontSize(11)
        .setFontColor('#000000')
        .setHorizontalAlignment('center')
        .setNumberFormat('0');
    });
  });
  
  console.log(`\n✅ Table 2 complete!`);
  console.log(`========================================\n`);
}

function populateTable3_ARTVelocity(sheet, piNumber) {
  console.log(`\n========================================`);
  console.log(`📊 TABLE 3: ART VELOCITY CHART - PI ${piNumber}`);
  console.log(`========================================\n`);
  
  // ====================================================================
  // STEP 1: SETUP
  // ====================================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(`PI ${piNumber} - Iteration 6`);
  
  if (!sheet) {
    throw new Error('Predictability Score sheet not found');
  }
  
  if (!sourceSheet) {
    throw new Error(`Source sheet not found: PI ${piNumber} - Iteration 6`);
  }
  
  // Table 3 configuration
  const ART_VELOCITY_START_ROW = 39;
  const ART_VELOCITY_HEADER_ROW = 40;
  const ART_VELOCITY_DATA_START_ROW = 41;
  
  // ART VELOCITY: ONE COLUMN PER PI
  // PI 11 starts at column L (index 12)
  const ART_VELOCITY_PI_COLUMN_MAP = {
    11: 12,  // Column L
    12: 13,  // Column M
    13: 14,  // Column N
    14: 15,  // Column O
    15: 16   // Column P
  };
  
  const targetCol = ART_VELOCITY_PI_COLUMN_MAP[piNumber];
  
  if (!targetCol) {
    throw new Error(`No column mapping defined for PI ${piNumber} in ART Velocity table`);
  }
  
  console.log(`📍 Table Configuration:`);
  console.log(`   PI ${piNumber} → Column ${columnNumberToLetter(targetCol)} (index ${targetCol})`);
  console.log(`   Data rows start at: Row ${ART_VELOCITY_DATA_START_ROW}`);
  
  // ====================================================================
  // STEP 2: READ SOURCE DATA
  // ====================================================================
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[3]; // Row 4 (0-indexed as 3) has headers
  
  // Find required columns
  const vsCol = sourceHeaders.indexOf('Value Stream/Org');
  const statusCol = sourceHeaders.indexOf('Status');
  const issueTypeCol = sourceHeaders.indexOf('Issue Type');
  const storyPointsCol = sourceHeaders.indexOf('Story Point Completion');
  
  console.log(`\n📋 Source Sheet Columns:`);
  console.log(`   Value Stream/Org: Column ${vsCol + 1}`);
  console.log(`   Status: Column ${statusCol + 1}`);
  console.log(`   Issue Type: Column ${issueTypeCol + 1}`);
  console.log(`   Story Point Completion: Column ${storyPointsCol + 1}`);
  
  if (vsCol === -1 || statusCol === -1 || issueTypeCol === -1 || storyPointsCol === -1) {
    throw new Error('Required columns not found in source sheet');
  }
  
  // ====================================================================
  // STEP 3: GET VALUE STREAM LIST FROM CONFIGURATION
  // ====================================================================
  const configSheet = ss.getSheetByName('Configuration Template');
  if (!configSheet) {
    throw new Error('Configuration Template sheet not found');
  }
  
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  
  // Value streams start at row 6 in Configuration Template
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0]; // Column A
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  console.log(`\n📍 Processing ${valueStreams.length} Value Streams`);
  
  // ====================================================================
  // STEP 4: CALCULATE TOTAL STORY POINTS FOR EACH VALUE STREAM
  // ====================================================================
  valueStreams.forEach((valueStream, index) => {
    console.log(`\n📊 ${valueStream}:`);
    
    let totalStoryPoints = 0;
    let totalEpicsChecked = 0;
    let closedEpicsProcessed = 0;
    
    // Scan all rows in source sheet (start from row 5, which is index 4)
    for (let i = 4; i < sourceData.length; i++) {
      const row = sourceData[i];
      const rowVS = row[vsCol];
      const rowStatus = row[statusCol];
      const rowIssueType = row[issueTypeCol];
      const rowStoryPoints = row[storyPointsCol];
      
      // Only process epics for this value stream
      if (rowIssueType !== 'Epic') continue;
      if (rowVS !== valueStream) continue;
      
      totalEpicsChecked++;
      
      // ⚠️ CRITICAL FILTER: Only include if Status = "Closed"
      const normalizedStatus = (rowStatus || '').toString().trim().toLowerCase();
      if (normalizedStatus !== 'closed' && normalizedStatus !== 'pending acceptance') continue;

      
      closedEpicsProcessed++;
      
      // Sum story points (ALL allocations, no filtering by allocation type)
      const storyPoints = parseFloat(rowStoryPoints) || 0;
      totalStoryPoints += storyPoints;
    }
    
    // Calculate target row
    const rowNum = ART_VELOCITY_DATA_START_ROW + index;
    
    console.log(`   Closed epics: ${closedEpicsProcessed}/${totalEpicsChecked}`);
    console.log(`   Total story points: ${totalStoryPoints.toFixed(2)}`);
    console.log(`   Writing to: ${columnNumberToLetter(targetCol)}${rowNum}`);
    
    // ====================================================================
    // STEP 5: WRITE TO SHEET
    // ====================================================================
    const cell = sheet.getRange(rowNum, targetCol);
    const roundedValue = Math.round(totalStoryPoints);
    
    cell.setValue(roundedValue)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    console.log(`   ✓ Wrote ${roundedValue}`);
  });
  
  console.log(`\n✅ Table 3 (ART Velocity) complete!`);
  console.log(`========================================\n`);
}

function populateTable4_EpicStatus(sheet, piNumber) {
  console.log(`\n========================================`);
  console.log(`📊 TABLE 4: EPIC STATUS - PI ${piNumber}`);
  console.log(`========================================\n`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const iter6SheetName = `PI ${piNumber} - Iteration 6`;
  const iter6Sheet = ss.getSheetByName(iter6SheetName);
  
  if (!iter6Sheet) {
    throw new Error(`Iteration 6 sheet not found: ${iter6SheetName}`);
  }
  
  if (!sheet) {
    throw new Error('Predictability Score sheet not found');
  }
  
  console.log(`✅ Found sheet: ${iter6SheetName}`);
  
  const configSheet = ss.getSheetByName('Configuration Template');
  if (!configSheet) {
    throw new Error('Configuration Template sheet not found');
  }
  
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0];
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  console.log(`\n📍 Processing ${valueStreams.length} Value Streams`);
  
  const iter6Data = iter6Sheet.getDataRange().getValues();
  console.log(`📊 Read ${iter6Data.length} rows from ${iter6SheetName}`);
  
  const dataRows = iter6Data.slice(1);
  
  const EPIC_STATUS_PI_COLUMN_START = {
    11: 2,
    12: 5,
    13: 8,
    14: 11,
    15: 14
  };
  
  const startCol = EPIC_STATUS_PI_COLUMN_START[piNumber];
  
  if (!startCol) {
    throw new Error(`No column mapping defined for PI ${piNumber} in Epic Status table`);
  }
  
  console.log(`📍 PI ${piNumber} Epic Status columns start at: Column ${columnNumberToLetter(startCol)}`);
  console.log(`   Committed: Column ${columnNumberToLetter(startCol)}`);
  console.log(`   Deferred: Column ${columnNumberToLetter(startCol + 1)}`);
  console.log(`   Deferred taken into Planning: Column ${columnNumberToLetter(startCol + 2)} (will be left empty)`);
  
  valueStreams.forEach((valueStream, index) => {
    console.log(`\n${'─'.repeat(60)}`);
    console.log(`📊 ${valueStream}`);
    console.log(`${'─'.repeat(60)}`);
    
    const vsRows = dataRows.filter(row => {
      const rowVS = row[ITER6_COL_VALUE_STREAM];
      const issueType = row[ITER6_COL_ISSUE_TYPE];
      return rowVS === valueStream && issueType === 'Epic';
    });
    
    console.log(`   Found ${vsRows.length} Epics for ${valueStream}`);
    
    let committedCount = 0;
    let deferredCount = 0;
    let skippedQualityKLO = 0;
    let skippedTradedBlank = 0;
    let skippedOther = 0;
    
    // Track unique commitment values we see
    const uniqueCommitments = new Set();
    
    vsRows.forEach(row => {
      const commitment = row[ITER6_COL_PI_COMMITMENT];
      const allocation = row[ITER6_COL_ALLOCATION];
      const key = row[ITER6_COL_KEY];
      
      // Track all commitment values
      if (commitment) {
        uniqueCommitments.add(commitment);
      }
      
      // EXCLUDE epics with Quality or KLO allocations
      if (allocation === 'Quality' || allocation === 'KLO') {
        skippedQualityKLO++;
        return;  // Skip this epic
      }
      
      // Skip if PI Commitment is "Traded" or blank
      if (!commitment || commitment === 'Traded') {
        skippedTradedBlank++;
        return;  // Skip this epic
      }
      
      if (commitment === 'Committed' || commitment === 'Committed After Plan') {
        committedCount++;
        console.log(`   ✓ COMMITTED: ${key} - "${commitment}"`);
      } else if (commitment === 'Deferred' || commitment === 'Canceled') {
        deferredCount++;
        console.log(`   ✓ DEFERRED: ${key} - "${commitment}"`);
      } else {
        skippedOther++;
        console.log(`   ⚠ SKIPPED (other): ${key} - Commitment: "${commitment}", Allocation: "${allocation}"`);
      }
    });
    
    console.log(`\n   📋 SUMMARY FOR ${valueStream}:`);
    console.log(`   - Committed: ${committedCount}`);
    console.log(`   - Deferred: ${deferredCount}`);
    console.log(`   - Skipped (Quality/KLO): ${skippedQualityKLO}`);
    console.log(`   - Skipped (Traded/Blank): ${skippedTradedBlank}`);
    console.log(`   - Skipped (Other): ${skippedOther}`);
    console.log(`   - Unique PI Commitment values seen: ${Array.from(uniqueCommitments).join(', ')}`);
    
    const EPIC_STATUS_DATA_START_ROW = 61;
    const rowNum = EPIC_STATUS_DATA_START_ROW + index;
    
    console.log(`   Writing to row ${rowNum}...`);
    
    const committedCell = sheet.getRange(rowNum, startCol);
    committedCell.setValue(committedCount)
      .setFontFamily(PREDICTABILITY_FONT)
      .setFontSize(PREDICTABILITY_FONT_SIZE)
      .setFontColor(PREDICTABILITY_FONT_COLOR)
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    const deferredCell = sheet.getRange(rowNum, startCol + 1);
    deferredCell.setValue(deferredCount)
      .setFontFamily(PREDICTABILITY_FONT)
      .setFontSize(PREDICTABILITY_FONT_SIZE)
      .setFontColor(PREDICTABILITY_FONT_COLOR)
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    const deferredTakenCell = sheet.getRange(rowNum, startCol + 2);
    deferredTakenCell.clearContent()
      .setFontFamily(PREDICTABILITY_FONT)
      .setFontSize(PREDICTABILITY_FONT_SIZE)
      .setFontColor(PREDICTABILITY_FONT_COLOR)
      .setHorizontalAlignment('center');
    
    console.log(`   ✓ Wrote: Committed=${committedCount}, Deferred=${deferredCount}, Deferred taken into Planning=(empty)`);
  });
  
  console.log(`\n✅ Table 4 complete!`);
  console.log(`========================================\n`);
}

function populateTable5_ProgramPredictability(sheet, piNumber) {
  console.log(`\n========================================`);
  console.log(`TABLE 5: PROGRAM PREDICTABILITY - PI ${piNumber}`);
  console.log(`========================================`);
  console.log(`Business Rules:`);
  console.log(`  • EXCLUDE: PI Commitment = "Canceled" or "Traded" (from both BV and AV)`);
  console.log(`  • INCLUDE: PI Commitment = "Committed", "Committed After Plan", or "Deferred"`);
  console.log(`  • EXCLUDE: Allocation = Quality or KLO`);
  console.log(`  • BV: ALL epics (any status) that pass above filters`);
  console.log(`  • AV: Only Closed/Pending Acceptance epics with PI Objective Status = "Met"`);
  console.log(`  • MMPM: Exclude Scrum Team = "Appeal Engine" from both BV and AV`);
  console.log(`========================================`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = `PI ${piNumber} - Iteration 6`;
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    console.error(`❌ Source sheet not found: ${sourceSheetName}`);
    throw new Error(`Cannot find required sheet: ${sourceSheetName}`);
  }
  
  // Valid PI Commitment values for inclusion
  const VALID_PI_COMMITMENTS = ['Committed', 'Committed After Plan', 'Deferred'];
  
  // Table structure
  const TABLE_START_ROW = 76;
  const PI_HEADER_ROW = 77;
  const CRITERIA_ROW = 78;
  const FIRST_DATA_ROW = 80;
  const BASE_PI = 11;
  const COLUMNS_PER_PI = 3; // Business Value, Actual Value, PI Score
  
  const piIndex = piNumber - BASE_PI;
  const startCol = 2 + (piIndex * COLUMNS_PER_PI);
  
  console.log(`📊 PI ${piNumber} columns: ${columnNumberToLetter(startCol)}-${columnNumberToLetter(startCol + COLUMNS_PER_PI - 1)}`);
  
  // Read source data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[3];
  
  const vsCol = sourceHeaders.indexOf('Value Stream/Org');
  const allocationCol = sourceHeaders.indexOf('Allocation');
  const statusCol = sourceHeaders.indexOf('Status');
  const issueTypeCol = sourceHeaders.indexOf('Issue Type');
  const piObjectiveStatusCol = sourceHeaders.indexOf('PI Objective Status');
  const businessValueCol = sourceHeaders.indexOf('Business Value');
  const actualValueCol = sourceHeaders.indexOf('Actual Value');
  const scrumTeamCol = sourceHeaders.indexOf('Scrum Team');
  const keyCol = sourceHeaders.indexOf('Key');
  const piCommitmentCol = sourceHeaders.indexOf('PI Commitment');
  
  if (vsCol === -1 || allocationCol === -1 || statusCol === -1 || issueTypeCol === -1 || 
      piObjectiveStatusCol === -1 || businessValueCol === -1 || actualValueCol === -1) {
    throw new Error('Required columns not found in source sheet');
  }
  
  if (scrumTeamCol === -1) {
    console.warn(`⚠️ Scrum Team column not found - MMPM Appeal Engine filter will not be applied`);
  }
  
  if (piCommitmentCol === -1) {
    console.warn(`⚠️ PI Commitment column not found - PI Commitment filter will not be applied`);
  }
  
  // Get value streams
  const configSheet = ss.getSheetByName('Configuration Template');
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0];
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  // Process each value stream
  valueStreams.forEach((valueStream, index) => {
    console.log(`\n🔍 Processing: ${valueStream}`);
    
    let totalBusinessValue = 0;
    let totalActualValue = 0;
    let totalEpicsChecked = 0;
    let bvEpicsIncluded = 0;
    let avEpicsIncluded = 0;
    let excludedQualityKLO = 0;
    let excludedAppealEngine = 0;
    let excludedPICommitment = 0;
    
    // Check if this is MMPM (case-insensitive)
    const isMMPM = valueStream.toUpperCase() === 'MMPM';
    if (isMMPM) {
      console.log(`   ⚠️ MMPM detected - will exclude Scrum Team = "Appeal Engine"`);
    }
    
    for (let i = 4; i < sourceData.length; i++) {
      const row = sourceData[i];
      const rowVS = row[vsCol];
      const rowAllocation = row[allocationCol];
      const rowStatus = row[statusCol];
      const rowIssueType = row[issueTypeCol];
      const rowPIObjectiveStatus = row[piObjectiveStatusCol];
      const rowBusinessValue = row[businessValueCol];
      const rowActualValue = row[actualValueCol];
      const rowScrumTeam = scrumTeamCol !== -1 ? row[scrumTeamCol] : '';
      const rowKey = keyCol !== -1 ? row[keyCol] : '';
      const rowPICommitment = piCommitmentCol !== -1 ? row[piCommitmentCol] : '';
      
      // Only process Epics for this value stream
      if (rowIssueType !== 'Epic') continue;
      if (rowVS !== valueStream) continue;
      
      totalEpicsChecked++;
      
      const normalizedStatus = (rowStatus || '').toString().trim().toLowerCase();
      const normalizedPIObjectiveStatus = (rowPIObjectiveStatus || '').toString().trim().toLowerCase();
      const allocation = (rowAllocation || '').toString().trim();
      const scrumTeam = (rowScrumTeam || '').toString().trim();
      const piCommitment = (rowPICommitment || '').toString().trim();
      
      // FILTER 1: Exclude if PI Commitment is NOT in valid list (Committed, Committed After Plan, Deferred)
      // This excludes: Canceled, Traded, blank, and any other values
      if (piCommitmentCol !== -1 && !VALID_PI_COMMITMENTS.includes(piCommitment)) {
        excludedPICommitment++;
        console.log(`   Excluding epic (PI Commitment = "${piCommitment}"): ${rowKey}`);
        continue;
      }
      
      // FILTER 2: Exclude Quality and KLO allocations (applies to both BV and AV)
      if (allocation === 'Quality' || allocation === 'KLO / Quality' || allocation === 'KLO') {
        excludedQualityKLO++;
        continue;
      }
      
      // FILTER 3: MMPM-specific - Exclude Appeal Engine scrum team (applies to both BV and AV)
      if (isMMPM && scrumTeam.toLowerCase() === 'appeal engine') {
        excludedAppealEngine++;
        console.log(`   Excluding Appeal Engine epic: ${rowKey}`);
        continue;
      }
      
      // Parse values
      const businessValue = parseFloat(rowBusinessValue) || 0;
      const actualValue = parseFloat(rowActualValue) || 0;
      
      // ✅ BV: Include ALL epics (any status) that passed all filters above
      totalBusinessValue += businessValue;
      bvEpicsIncluded++;
      
      // ✅ AV: Only include if Status = Closed/Pending Acceptance AND PI Objective Status = "Met"
      const isClosedOrPending = normalizedStatus === 'closed' || normalizedStatus === 'pending acceptance';
      if (isClosedOrPending && normalizedPIObjectiveStatus === 'met') {
        totalActualValue += actualValue;
        avEpicsIncluded++;
      }
    }
    
    // Calculate PI Score: (Actual Value / Business Value) * 100%
    const piScore = totalBusinessValue > 0 
      ? (totalActualValue / totalBusinessValue) * 100 
      : 0;
    
    console.log(`   Total epics checked: ${totalEpicsChecked}`);
    console.log(`   Excluded (PI Commitment = Canceled/Traded/blank): ${excludedPICommitment}`);
    console.log(`   Excluded (Quality/KLO): ${excludedQualityKLO}`);
    if (isMMPM) {
      console.log(`   Excluded (Appeal Engine): ${excludedAppealEngine}`);
    }
    console.log(`   BV epics included (all statuses): ${bvEpicsIncluded}`);
    console.log(`   AV epics included (Closed/Pending + Met): ${avEpicsIncluded}`);
    console.log(`   Business Value: ${totalBusinessValue}`);
    console.log(`   Actual Value: ${totalActualValue}`);
    console.log(`   PI Score: ${piScore.toFixed(1)}%`);
    
    // Write to sheet
    const rowNum = FIRST_DATA_ROW + index;
    
    // Business Value (Column B/E/H/etc)
    const bvCell = sheet.getRange(rowNum, startCol);
    bvCell.setValue(totalBusinessValue)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    // Actual Value (Column C/F/I/etc)
    const avCell = sheet.getRange(rowNum, startCol + 1);
    avCell.setValue(totalActualValue)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    // PI Score (Column D/G/J/etc)
    const scoreCell = sheet.getRange(rowNum, startCol + 2);
    scoreCell.setValue(piScore)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0.0"%"');
  });
  
  console.log(`\n✅ Table 5 complete!`);
  console.log(`========================================\n`);
}

function populateTable6_ObjectiveStatus(sheet, piNumber) {
  console.log(`\n========================================`);
  console.log(`TABLE 6: OBJECTIVE STATUS - PI ${piNumber}`);
  console.log(`========================================`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = `PI ${piNumber} - Iteration 6`;
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    console.error(`❌ Source sheet not found: ${sourceSheetName}`);
    throw new Error(`Cannot find required sheet: ${sourceSheetName}`);
  }
  
  // Table structure
  const TABLE_START_ROW = 96;
  const PI_HEADER_ROW = 97;
  const CRITERIA_ROW = 98;
  const FIRST_DATA_ROW = 99;
  const BASE_PI = 11;
  const COLUMNS_PER_PI = 5; // Met, Partial, Not Met, ALL Epics, Objective Score
  
  const piIndex = piNumber - BASE_PI;
  const startCol = 2 + (piIndex * COLUMNS_PER_PI);
  
  console.log(`📊 PI ${piNumber} columns: ${columnNumberToLetter(startCol)}-${columnNumberToLetter(startCol + COLUMNS_PER_PI - 1)}`);
  
  // Read source data
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[3];
  
  const vsCol = sourceHeaders.indexOf('Value Stream/Org');
  const statusCol = sourceHeaders.indexOf('Status');
  const issueTypeCol = sourceHeaders.indexOf('Issue Type');
  const piObjectiveStatusCol = sourceHeaders.indexOf('PI Objective Status');
  
  if (vsCol === -1 || statusCol === -1 || issueTypeCol === -1 || piObjectiveStatusCol === -1) {
    throw new Error('Required columns not found in source sheet');
  }
  
  // Get value streams
  const configSheet = ss.getSheetByName('Configuration Template');
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0];
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  // Process each value stream
  valueStreams.forEach((valueStream, index) => {
    console.log(`\n🔍 Processing: ${valueStream}`);
    
    let metCount = 0;
    let partialCount = 0;
    let notMetCount = 0;
    
    for (let i = 4; i < sourceData.length; i++) {
      const row = sourceData[i];
      const rowVS = row[vsCol];
      const rowStatus = row[statusCol];
      const rowIssueType = row[issueTypeCol];
      const rowPIObjectiveStatus = row[piObjectiveStatusCol];
      
      if (rowIssueType !== 'Epic') continue;
      if (rowVS !== valueStream) continue;
      
      const normalizedStatus = (rowStatus || '').toString().trim().toLowerCase();
      const normalizedPIObjectiveStatus = (rowPIObjectiveStatus || '').toString().trim().toLowerCase();
      
      // ⚠️ CRITICAL: Only count epics with Status = "Closed" or "Pending Acceptance"
      if (normalizedStatus !== 'closed' && normalizedStatus !== 'pending acceptance') continue;
      
      // Count by PI Objective Status (only count if they have a valid status)
      if (normalizedPIObjectiveStatus === 'met') {
        metCount++;
      } else if (normalizedPIObjectiveStatus === 'partial') {
        partialCount++;
      } else if (normalizedPIObjectiveStatus === 'not met') {
        notMetCount++;
      }
      // ✅ If PI Objective Status is blank or other value, don't count it anywhere
    }
    
    // ✅ FIXED: ALL = Sum of Met + Partial + Not Met (only epics with valid PI Objective Status)
    const allClosedEpicsCount = metCount + partialCount + notMetCount;
    
    // Calculate Objective Score: (Met / ALL) * 100%
    const objectiveScore = allClosedEpicsCount > 0 
      ? (metCount / allClosedEpicsCount) * 100 
      : 0;
    
    console.log(`   Met (Closed + Met): ${metCount}`);
    console.log(`   Partial (Closed + Partial): ${partialCount}`);
    console.log(`   Not Met (Closed + Not Met): ${notMetCount}`);
    console.log(`   ALL (Met + Partial + Not Met): ${allClosedEpicsCount}`);
    console.log(`   Objective Score: ${objectiveScore.toFixed(1)}%`);
    
    // Write to sheet
    const rowNum = FIRST_DATA_ROW + index;
    
    // Met count (Column B/G/L/etc)
    const metCell = sheet.getRange(rowNum, startCol);
    metCell.setValue(metCount)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    // Partial count (Column C/H/M/etc)
    const partialCell = sheet.getRange(rowNum, startCol + 1);
    partialCell.setValue(partialCount)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    // Not Met count (Column D/I/N/etc)
    const notMetCell = sheet.getRange(rowNum, startCol + 2);
    notMetCell.setValue(notMetCount)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    // ALL Closed Epics count (Column E/J/O/etc) - ✅ FIXED: Now sum of Met + Partial + Not Met
    const allEpicsCell = sheet.getRange(rowNum, startCol + 3);
    allEpicsCell.setValue(allClosedEpicsCount)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0');
    
    // Objective Score (Column F/K/P/etc)
    const scoreCell = sheet.getRange(rowNum, startCol + 4);
    scoreCell.setValue(objectiveScore)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0.0"%"');
  });
  
  console.log(`\n✅ Table 6 complete!`);
  console.log(`========================================\n`);
}
function populateTable7_PIScoreSummary(sheet, piNumber) {
  console.log(`\n========================================`);
  console.log(`📊 TABLE 7: PI SCORE SUMMARY - PI ${piNumber}`);
  console.log(`========================================\n`);
  
  // ====================================================================
  // STEP 1: SETUP
  // ====================================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!sheet) {
    throw new Error('Predictability Score sheet not found');
  }
  
  // Table 7 configuration
  const TABLE7_START_ROW = 115;
  const TABLE7_HEADER_ROW = 116;
  const TABLE7_DATA_START_ROW = 117;
  
  // Table 5 configuration (source of PI Score data)
  const TABLE5_DATA_START_ROW = 80;  // First data row in Program Predictability
  const TABLE5_BASE_PI = 11;
  const TABLE5_COLUMNS_PER_PI = 3;   // Business Value, Actual Value, PI Score
  const TABLE5_PI_SCORE_OFFSET = 2;  // PI Score is the 3rd column (offset 2)
  
  // PI Score Summary: ONE COLUMN PER PI (same as ART Velocity)
  const PI_SCORE_SUMMARY_COLUMN_MAP = {
    11: 12,  // Column L
    12: 13,  // Column M
    13: 14,  // Column N
    14: 15,  // Column O
    15: 16   // Column P
  };
  
  const table7TargetCol = PI_SCORE_SUMMARY_COLUMN_MAP[piNumber];
  
  if (!table7TargetCol) {
    throw new Error(`No column mapping defined for PI ${piNumber} in Table 7`);
  }
  
  // Calculate source column in Table 5 (Program Predictability)
  const piIndex = piNumber - TABLE5_BASE_PI;
  const table5StartCol = 2 + (piIndex * TABLE5_COLUMNS_PER_PI); // Column B = index 2
  const table5PIScoreCol = table5StartCol + TABLE5_PI_SCORE_OFFSET; // PI Score is 3rd column
  
  console.log(`📍 Table Configuration:`);
  console.log(`   Table 7 starts at: Row ${TABLE7_START_ROW}`);
  console.log(`   PI headers in: Row ${TABLE7_HEADER_ROW}`);
  console.log(`   Data starts at: Row ${TABLE7_DATA_START_ROW}`);
  console.log(`   PI ${piNumber} → Column ${columnNumberToLetter(table7TargetCol)} (index ${table7TargetCol})`);
  console.log(`\n📍 Source: Table 5 (Program Predictability)`);
  console.log(`   Reading from: Column ${columnNumberToLetter(table5PIScoreCol)} (PI Score column)`);
  console.log(`   Starting at: Row ${TABLE5_DATA_START_ROW}`);
  
  // ====================================================================
  // STEP 2: GET VALUE STREAM LIST FROM CONFIGURATION
  // ====================================================================
  const configSheet = ss.getSheetByName('Configuration Template');
  if (!configSheet) {
    throw new Error('Configuration Template sheet not found');
  }
  
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  
  // Value streams start at row 6 in Configuration Template
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0]; // Column A
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  console.log(`\n📍 Processing ${valueStreams.length} Value Streams`);
  
  // ====================================================================
  // STEP 3: COPY PI SCORE VALUES FROM TABLE 5 TO TABLE 7
  // ====================================================================
  console.log(`\n🔄 Copying PI Score data from Table 5 to Table 7...`);
  
  valueStreams.forEach((valueStream, index) => {
    console.log(`\n📊 ${valueStream}:`);
    
    // Calculate source row in Table 5
    const table5RowNum = TABLE5_DATA_START_ROW + index;
    
    // Read PI Score value from Table 5
    const sourceCell = sheet.getRange(table5RowNum, table5PIScoreCol);
    const piScoreValue = sourceCell.getValue();
    
    // Calculate target row in Table 7
    const table7RowNum = TABLE7_DATA_START_ROW + index;
    
    console.log(`   Source: ${columnNumberToLetter(table5PIScoreCol)}${table5RowNum} (Table 5)`);
    console.log(`   PI Score: ${piScoreValue}`);
    console.log(`   Target: ${columnNumberToLetter(table7TargetCol)}${table7RowNum} (Table 7)`);
    
    // ====================================================================
    // STEP 4: WRITE TO TABLE 7
    // ====================================================================
    const targetCell = sheet.getRange(table7RowNum, table7TargetCol);
    
    // If the source value is already a number (not a formula), write it directly
    // If it's a string like "75.0%", parse it
    let valueToWrite = piScoreValue;
    
    if (typeof piScoreValue === 'string' && piScoreValue.includes('%')) {
      // Parse percentage string like "75.0%" to number 75.0
      valueToWrite = parseFloat(piScoreValue.replace('%', ''));
    } else if (typeof piScoreValue === 'number') {
      // Value is already a number, use as-is
      valueToWrite = piScoreValue;
    }
    
    targetCell.setValue(valueToWrite)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontColor('#000000')
      .setHorizontalAlignment('center')
      .setNumberFormat('0.0"%"');
    
    console.log(`   ✓ Wrote ${valueToWrite}% to ${columnNumberToLetter(table7TargetCol)}${table7RowNum}`);
  });
  
  console.log(`\n✅ Table 7 (PI Score Summary) complete!`);
  console.log(`========================================\n`);
}


function columnNumberToLetter(column) {
  let letter = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

function addDeferralsToPredictabilityReport(presentationId, piNumber, sortBy) {
  return logActivity('Add Deferrals to Predictability Report', () => {
    const startTime = Date.now();
    const ISSUES_PER_SLIDE = 5;
    const DEFERRAL_TEMPLATE_SLIDE_INDEX = 2; // Page 3 (0-indexed)
    const JIRA_BASE_URL = 'https://modmedrnd.atlassian.net/browse/';
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    sortBy = sortBy || 'program_initiative';
    
    console.log(`\n${'='.repeat(80)}`);
    console.log(`ADD DEFERRALS TO PREDICTABILITY REPORT`);
    console.log(`Presentation ID: ${presentationId}`);
    console.log(`PI: ${piNumber}`);
    console.log(`Sort By: ${sortBy === 'program_initiative' ? 'Program Initiative' : 'Value Stream/Org'}`);
    console.log(`${'='.repeat(80)}\n`);
    
    try {
      // ═══════════════════════════════════════════════════════════════════
      // STEP 1: OPEN EXISTING PRESENTATION
      // ═══════════════════════════════════════════════════════════════════
      const presentation = SlidesApp.openById(presentationId);
      const slides = presentation.getSlides();
      
      if (slides.length < 3) {
        throw new Error('Predictability presentation must have at least 3 slides. Page 3 should be the deferral template.');
      }
      
      const templateSlide = slides[DEFERRAL_TEMPLATE_SLIDE_INDEX];
      console.log(`✓ Using slide ${DEFERRAL_TEMPLATE_SLIDE_INDEX + 1} (page 3) as deferral template`);
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 2: GET DEFERRAL DATA
      // ═══════════════════════════════════════════════════════════════════
      const sourceSheetName = `PI ${piNumber} - Iteration 6`;
      const sourceSheet = ss.getSheetByName(sourceSheetName);
      
      if (!sourceSheet) {
        console.log(`⚠️ No deferrals: ${sourceSheetName} sheet not found`);
        return { slideCount: 0, deferralCount: 0, skipped: true };
      }
      
      // Read and filter deferral data (same logic as standalone)
      const dataRange = sourceSheet.getRange(4, 1, sourceSheet.getLastRow() - 3, sourceSheet.getLastColumn());
      const values = dataRange.getValues();
      const headers = values.shift();

      // Find columns
      let piCommitmentCol = headers.indexOf('PI Commitment');
      if (piCommitmentCol === -1) {
        piCommitmentCol = headers.findIndex(h => 
          h && h.toString().toLowerCase().includes('pi commitment')
        );
      }
      
      const ragNoteCol = headers.indexOf('RAG Note');
      const initiativeCol = headers.indexOf('Program Initiative');
      const keyCol = headers.indexOf('Issue key');
      const summaryCol = headers.indexOf('Summary');
      const valueStreamCol = headers.indexOf('Value Stream/Org');
      const aiSummaryCol = headers.indexOf('RAG Note - AI Summary');

      if (piCommitmentCol === -1 || initiativeCol === -1) {
        throw new Error("Required columns not found: 'PI Commitment' or 'Program Initiative'");
      }
      
      console.log(`✓ Found PI Commitment column: "${headers[piCommitmentCol]}"`);

      // Filter deferrals
      const deferralStatuses = ['deferred', 'canceled', 'traded', 'cancelled'];
      const issues = values.map(row => ({
        key: row[keyCol],
        programInitiative: row[initiativeCol],
        summary: row[summaryCol],
        piCommitment: row[piCommitmentCol],
        ragNote: row[ragNoteCol] || '',
        valueStreamOrg: row[valueStreamCol],
        aiRagNote: row[aiSummaryCol] || ''
      })).filter(issue => {
        const commitment = (issue.piCommitment || '').toString().trim().toLowerCase();
        return deferralStatuses.includes(commitment);
      });

      if (issues.length === 0) {
        console.log('⚠️ No deferrals found - skipping deferral section');
        return { slideCount: 0, deferralCount: 0, skipped: true };
      }

      // Sort
      if (sortBy === 'value_stream') {
        issues.sort((a, b) => {
          const vsCompare = (a.valueStreamOrg || '').localeCompare(b.valueStreamOrg || '');
          if (vsCompare !== 0) return vsCompare;
          return (a.programInitiative || '').localeCompare(b.programInitiative || '');
        });
      } else {
        issues.sort((a, b) => {
          const piCompare = (a.programInitiative || '').localeCompare(b.programInitiative || '');
          if (piCompare !== 0) return piCompare;
          return (a.valueStreamOrg || '').localeCompare(b.valueStreamOrg || '');
        });
      }
      
      console.log(`✓ Found ${issues.length} deferrals`);

      // ═══════════════════════════════════════════════════════════════════
      // STEP 3: INSERT DEFERRAL SLIDES AFTER TEMPLATE
      // ═══════════════════════════════════════════════════════════════════
      let slidesCreated = 0;
      const lastRunDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');
      
      // Calculate where to insert (after the template slide)
      const insertPosition = DEFERRAL_TEMPLATE_SLIDE_INDEX + 1;
      
      for (let i = 0; i < issues.length; i += ISSUES_PER_SLIDE) {
        const issuesBatch = issues.slice(i, i + ISSUES_PER_SLIDE);
        
        // Insert at the correct position (pushes everything else back)
        const newSlide = presentation.insertSlide(insertPosition + slidesCreated, templateSlide);
        slidesCreated++;

        // Populate placeholders
        issuesBatch.forEach((issue, index) => {
          const rowNum = index + 1;

          newSlide.replaceAllText(`{{issue_key_${rowNum}}}`, issue.key || 'N/A');
          newSlide.replaceAllText(`{{program_initiative_${rowNum}}}`, issue.programInitiative || 'N/A');
          newSlide.replaceAllText(`{{epic_summary_${rowNum}}}`, issue.summary || 'N/A');
          newSlide.replaceAllText(`{{rag_note_${rowNum}}}`, issue.ragNote || '');
          newSlide.replaceAllText(`{{value_stream_org_${rowNum}}}`, issue.valueStreamOrg || 'N/A');
          newSlide.replaceAllText(`{{ai_rag_note_${rowNum}}}`, issue.aiRagNote || '');
        });
        
        // Add hyperlinks AFTER all text replacements are complete
        issuesBatch.forEach((issue, index) => {
          if (!issue.key || !issue.summary || issue.summary === 'N/A') {
            return; // Skip if no valid data
          }
          
          try {
            const shapes = newSlide.getShapes();
            let hyperlinkAdded = false;
            
            for (let i = 0; i < shapes.length && !hyperlinkAdded; i++) {
              const shape = shapes[i];
              
              try {
                const textRange = shape.getText();
                const fullText = textRange.asString();
                
                // Look for the exact summary text
                const summaryIndex = fullText.indexOf(issue.summary);
                
                if (summaryIndex !== -1) {
                  // Found the summary text - add hyperlink
                  const summaryEndIndex = summaryIndex + issue.summary.length;
                  const summaryRange = textRange.getRange(summaryIndex, summaryEndIndex);
                  
                  // Set hyperlink
                  summaryRange.getTextStyle().setLinkUrl(JIRA_BASE_URL + issue.key);
                  
                  // Style it (black, no underline, Lato 8pt)
                  summaryRange.getTextStyle()
                    .setForegroundColor('#000000')
                    .setUnderline(false)
                    .setBold(false)
                    .setFontFamily('Lato')
                    .setFontSize(8);
                  
                  hyperlinkAdded = true;
                  console.log(`  ✓ Added hyperlink to Epic Summary: "${issue.summary.substring(0, 50)}..." → ${issue.key}`);
                }
              } catch (shapeError) {
                // This shape doesn't have text or can't be processed - skip it
                continue;
              }
            }
            
            if (!hyperlinkAdded) {
              console.warn(`  ⚠️ Could not find Epic Summary text to hyperlink for ${issue.key}`);
            }
          } catch (error) {
            console.error(`  ❌ Error adding hyperlink for ${issue.key}: ${error.message}`);
          }
        });

        newSlide.replaceAllText(`{{last_run_date}}`, lastRunDate);
        console.log(`  ✓ Created deferral slide ${slidesCreated} with ${issuesBatch.length} issue(s)`);
      }
      
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      console.log(`\n✅ DEFERRALS ADDED TO PREDICTABILITY REPORT`);
      console.log(`  Slides Created: ${slidesCreated}`);
      console.log(`  Deferrals: ${issues.length}`);
      console.log(`  Duration: ${duration}s\n`);

      return {
        slideCount: slidesCreated,
        deferralCount: issues.length,
        skipped: false
      };

    } catch (error) {
      console.error(`\n❌ ERROR adding deferrals: ${error.message}`);
      console.error(error.stack);
      throw error;
    }
  }, { 
    piNumber: piNumber,
    presentationId: presentationId
  });
}