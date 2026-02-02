/**
 * Predictability Diff Validation (Enhanced v2)
 * 
 * This module validates predictability by checking:
 * 1. Epics where Program Increment was changed/removed during PI execution window
 * 2. Epics with PI assigned but missing PI Commitment field
 * 3. Whether proper cloning workflow was followed when PI was changed
 * 
 * CLONE WORKFLOW VALIDATION:
 * Correct process: When deferring Epic A from PI 12 to PI 13:
 *   - Clone Epic A → Epic B (clone stays in PI 12, marked Deferred/Canceled)
 *   - Epic A (original) moves to PI 13
 * 
 * This report flags:
 *   - PI changed with NO clone left behind (audit trail broken)
 *   - Clone was moved instead of original (wrong epic moved)
 * 
 * Output: New standalone spreadsheet "PI {#} - Predictability Diff"
 * 
 * DEPENDENCIES (from Core.gs):
 * - FIELD_MAPPINGS (constant)
 * - getJiraConfig() 
 * - getJiraHeaders()
 * - searchJiraIssuesLimited()
 * 
 * REQUIRED SHEETS:
 * - "Iteration Parameters" with columns: PI (A), Iteration (B), Start Date (C), End Date (D)
 */

// ===== CONFIGURATION =====

// Destination folder for generated reports
const PREDICTABILITY_DIFF_FOLDER_ID = '1PUpzgiqNhc3pbrQqlUY1MtcKbftFLg9M';

// ===== DIALOG FUNCTIONS =====

/**
 * Shows dialog to select PI and Value Stream for predictability diff analysis
 */
function showPredictabilityDiffDialog() {
  const html = HtmlService.createHtmlOutput(getPredictabilityDiffDialogHtml())
    .setWidth(400)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(html, 'Predictability Diff Validation');
}

/**
 * Returns the HTML for the Predictability Diff dialog
 */
function getPredictabilityDiffDialogHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      padding: 20px;
      background-color: #f8f9fa;
    }
    .form-group {
      margin-bottom: 15px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      font-weight: 600;
      color: #333;
    }
    select, input {
      width: 100%;
      padding: 10px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 14px;
      box-sizing: border-box;
    }
    select:focus, input:focus {
      outline: none;
      border-color: #9b7bb8;
      box-shadow: 0 0 0 2px rgba(155, 123, 184, 0.2);
    }
    .btn-container {
      display: flex;
      gap: 10px;
      margin-top: 20px;
    }
    button {
      flex: 1;
      padding: 12px 20px;
      border: none;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.2s;
    }
    .btn-primary {
      background-color: #9b7bb8;
      color: white;
    }
    .btn-primary:hover {
      background-color: #8a6aa7;
    }
    .btn-secondary {
      background-color: #e9ecef;
      color: #495057;
    }
    .btn-secondary:hover {
      background-color: #dee2e6;
    }
    .description {
      font-size: 11px;
      color: #666;
      margin-top: 15px;
      padding: 10px;
      background-color: #fff;
      border-radius: 6px;
      border-left: 3px solid #9b7bb8;
    }
    #status {
      margin-top: 10px;
      padding: 10px;
      border-radius: 6px;
      display: none;
    }
    .status-processing {
      background-color: #fff3cd;
      color: #856404;
    }
    .status-error {
      background-color: #f8d7da;
      color: #721c24;
    }
  </style>
</head>
<body>
  <div class="form-group">
    <label for="piNumber">Program Increment (PI)</label>
    <input type="number" id="piNumber" min="1" max="99" value="12" placeholder="Enter PI number">
  </div>
  
  <div class="form-group">
    <label for="valueStream">Value Stream</label>
    <select id="valueStream">
      <option value="ALL">-- All Value Streams --</option>
      <option value="AIMM">AIMM</option>
      <option value="Cloud Foundation Services">Cloud Foundation Services</option>
      <option value="Cloud Operations">Cloud Operations</option>
      <option value="Data Platform Engineering">Data Platform Engineering</option>
      <option value="EMA Clinical">EMA Clinical</option>
      <option value="EMA RaC">EMA RaC</option>
      <option value="Fusion and Conversions">Fusion and Conversions</option>
      <option value="MMGI">MMGI</option>
      <option value="MMGI-Cloud">MMGI-Cloud</option>
      <option value="MMPM">MMPM</option>
      <option value="Patient Collaboration">Patient Collaboration</option>
      <option value="Platform Engineering">Platform Engineering</option>
      <option value="RCM Genie">RCM Genie</option>
      <option value="Security Engineering">Security Engineering</option>
      <option value="Shared Services Platform">Shared Services Platform</option>
      <option value="Xtract">Xtract</option>
    </select>
  </div>
  
  <div class="description">
    <strong>This validation checks:</strong><br>
    1. Epics where PI was changed/removed during execution<br>
    2. Epics with PI assigned but missing PI Commitment<br>
    3. Clone workflow compliance (was a clone left in original PI?)
  </div>
  
  <div id="status"></div>
  
  <div class="btn-container">
    <button class="btn-secondary" onclick="google.script.host.close()">Cancel</button>
    <button class="btn-primary" onclick="runValidation()">Run Validation</button>
  </div>
  
  <script>
    function runValidation() {
      const piNumber = document.getElementById('piNumber').value;
      const valueStream = document.getElementById('valueStream').value;
      
      if (!piNumber || piNumber < 1) {
        showStatus('Please enter a valid PI number', 'error');
        return;
      }
      
      showStatus('Running validation... This may take several minutes.', 'processing');
      
      google.script.run
        .withSuccessHandler(onSuccess)
        .withFailureHandler(onFailure)
        .generatePredictabilityDiff(parseInt(piNumber), valueStream);
    }
    
    function showStatus(message, type) {
      const status = document.getElementById('status');
      status.textContent = message;
      status.className = 'status-' + type;
      status.style.display = 'block';
    }
    
    function onSuccess() {
      google.script.host.close();
    }
    
    function onFailure(error) {
      showStatus('Error: ' + error.message, 'error');
    }
  </script>
</body>
</html>
  `;
}

// ===== DATE RANGE FUNCTIONS =====

/**
 * Gets the full PI execution window from Iteration Parameters
 */
function getPIDateRange(piNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Iteration Parameters');
  
  if (!sheet) {
    throw new Error('Sheet "Iteration Parameters" not found.');
  }
  
  const data = sheet.getDataRange().getValues();
  
  let firstStartDate = null;
  let lastEndDate = null;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowPI = row[0];
    
    if (rowPI == piNumber) {
      let startDate = null;
      if (row[2] instanceof Date) {
        startDate = new Date(row[2]);
      } else if (typeof row[2] === 'string' && row[2]) {
        startDate = new Date(row[2]);
      }
      
      let endDate = null;
      if (row[3] instanceof Date) {
        endDate = new Date(row[3]);
      } else if (typeof row[3] === 'string' && row[3]) {
        endDate = new Date(row[3]);
      }
      
      if (startDate && !isNaN(startDate.getTime())) {
        if (!firstStartDate || startDate < firstStartDate) {
          firstStartDate = startDate;
        }
      }
      
      if (endDate && !isNaN(endDate.getTime())) {
        if (!lastEndDate || endDate > lastEndDate) {
          lastEndDate = endDate;
        }
      }
    }
  }
  
  if (!firstStartDate || !lastEndDate) {
    throw new Error(`Could not find date range for PI ${piNumber} in Iteration Parameters sheet.`);
  }
  
  firstStartDate.setHours(0, 0, 0, 0);
  lastEndDate.setHours(23, 59, 59, 999);
  
  console.log(`PI ${piNumber} Date Range: ${firstStartDate.toISOString()} to ${lastEndDate.toISOString()}`);
  
  return {
    startDate: firstStartDate,
    endDate: lastEndDate
  };
}

// ===== PAGINATION HELPER =====

/**
 * Search JIRA with ASC/DESC pagination to handle >100 results
 * This works around the 100-row limit by querying from both ends
 * 
 * @param {string} jql - Base JQL query (without ORDER BY)
 * @returns {Array} All unique issues found
 */
function searchJiraWithPagination(jql) {
  const allIssues = [];
  const seenKeys = new Set();
  
  // First call: ORDER BY key ASC (gets first 100 alphabetically)
  console.log(`  Pagination 1/2: ORDER BY key ASC`);
  const ascJql = `${jql} ORDER BY key ASC`;
  const ascIssues = searchJiraIssuesLimited(ascJql, 100);
  
  ascIssues.forEach(issue => {
    if (!seenKeys.has(issue.key)) {
      seenKeys.add(issue.key);
      allIssues.push(issue);
    }
  });
  
  console.log(`  ASC query returned: ${ascIssues.length} issues`);
  
  // If we hit the limit, fetch from the other end
  if (ascIssues.length >= 100) {
    console.log(`  ⚠️ HIT 100-ROW LIMIT - Fetching from other end...`);
    
    Utilities.sleep(200);
    
    // Second call: ORDER BY key DESC (gets last 100 alphabetically)
    console.log(`  Pagination 2/2: ORDER BY key DESC`);
    const descJql = `${jql} ORDER BY key DESC`;
    const descIssues = searchJiraIssuesLimited(descJql, 100);
    
    let newCount = 0;
    descIssues.forEach(issue => {
      if (!seenKeys.has(issue.key)) {
        seenKeys.add(issue.key);
        allIssues.push(issue);
        newCount++;
      }
    });
    
    console.log(`  DESC query returned: ${descIssues.length} issues (${newCount} new)`);
    
    // Warn if we might still be missing data
    if (descIssues.length >= 100 && newCount > 50) {
      console.warn(`  ⚠️ WARNING: Both queries hit 100 limit. May have >200 results - some data could be missing.`);
    }
  }
  
  console.log(`  ✅ Total unique issues: ${allIssues.length}`);
  return allIssues;
}

// ===== MAIN VALIDATION FUNCTION =====

/**
 * Main function to generate predictability diff validation
 */
function generatePredictabilityDiff(piNumber, valueStream) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  
  // Handle "ALL" value streams
  if (valueStream === 'ALL') {
    generatePredictabilityDiffAllValueStreams(piNumber);
    return;
  }
  
  console.log(`\n${'='.repeat(80)}`);
  console.log(`PREDICTABILITY DIFF VALIDATION`);
  console.log(`PI: ${piNumber} | Value Stream: ${valueStream}`);
  console.log(`Started: ${new Date().toISOString()}`);
  console.log(`${'='.repeat(80)}\n`);
  
  try {
    // Step 1: Get PI date range
    ss.toast('Step 1/6: Getting PI date range...', '⏳ Processing', 30);
    const dateRange = getPIDateRange(piNumber);
    console.log(`PI ${piNumber} execution window: ${dateRange.startDate.toDateString()} to ${dateRange.endDate.toDateString()}`);
    
    // Step 2: Find epics with missing PI Commitment (with pagination)
    ss.toast('Step 2/6: Finding epics with missing PI Commitment...', '⏳ Processing', 30);
    const missingCommitmentEpics = findEpicsWithMissingCommitment(piNumber, valueStream);
    console.log(`Found ${missingCommitmentEpics.length} epics missing PI Commitment`);
    
    // Step 3: Find epics where PI was changed during execution (with pagination)
    ss.toast('Step 3/6: Finding epics with PI changes...', '⏳ Processing', 60);
    const changedPIEpics = findEpicsWithPIChanges(piNumber, valueStream, dateRange);
    console.log(`Found ${changedPIEpics.length} epics with PI changes`);
    
    // Step 4: Check clone workflow compliance
    ss.toast('Step 4/6: Checking clone workflow compliance...', '⏳ Processing', 60);
    const epicsWithCloneCheck = checkCloneWorkflowCompliance(changedPIEpics, piNumber);
    
    // Step 5: Combine and deduplicate results
    ss.toast('Step 5/6: Processing results...', '⏳ Processing', 10);
    const combinedResults = combineValidationResults(missingCommitmentEpics, epicsWithCloneCheck);
    console.log(`Total unique issues: ${combinedResults.length}`);
    
    // Step 6: Write to new spreadsheet
    ss.toast('Step 6/6: Creating report spreadsheet...', '⏳ Processing', 10);
    const sheetName = `PI ${piNumber} - Predictability Diff - ${valueStream}`;
    const reportUrl = writeValidationReport(combinedResults, sheetName, piNumber, valueStream, dateRange);
    
    const duration = (Date.now() - startTime) / 1000;
    console.log(`\n${'='.repeat(80)}`);
    console.log(`VALIDATION COMPLETE in ${duration.toFixed(1)} seconds`);
    console.log(`${'='.repeat(80)}\n`);
    
    // Count stats for summary
    const cloneProperCount = combinedResults.filter(r => r.cloneWorkflowStatus === 'PROPER').length;
    const cloneMissingCount = combinedResults.filter(r => r.cloneWorkflowStatus === 'NO_CLONE').length;
    const cloneWrongCount = combinedResults.filter(r => r.cloneWorkflowStatus === 'WRONG_EPIC_MOVED').length;
    
    ss.toast(`✅ Found ${combinedResults.length} issues. New spreadsheet created!`, '✅ Complete', 10);
    
    // Show completion dialog with clickable link
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 10px;">
        <p style="font-size: 14px; margin-bottom: 15px;">
          Found <strong>${combinedResults.length}</strong> issues requiring attention.
        </p>
        <ul style="font-size: 13px; margin-bottom: 15px;">
          <li>Missing PI Commitment: <strong>${missingCommitmentEpics.length}</strong></li>
          <li>PI Changed During Execution: <strong>${changedPIEpics.length}</strong></li>
        </ul>
        <p style="font-size: 13px; margin-bottom: 15px;"><strong>Clone Workflow:</strong></p>
        <ul style="font-size: 12px; margin-bottom: 15px;">
          <li>✅ Proper (clone left behind): <strong>${cloneProperCount}</strong></li>
          <li>⚠️ No clone found: <strong>${cloneMissingCount}</strong></li>
          <li>❌ Wrong epic moved: <strong>${cloneWrongCount}</strong></li>
        </ul>
        <p style="font-size: 14px; margin-bottom: 10px;">
          <a href="${reportUrl}" target="_blank" style="color: #1a73e8; font-weight: bold;">
            📊 Open Report Spreadsheet
          </a>
        </p>
        <p style="font-size: 11px; color: #666;">
          Completed in ${duration.toFixed(1)} seconds
        </p>
      </div>`
    )
    .setWidth(380)
    .setHeight(280);
    
    ui.showModalDialog(htmlOutput, '✅ Validation Complete');
      
  } catch (error) {
    console.error('Validation error:', error);
    console.error('Stack:', error.stack);
    ui.alert('Error', `Validation failed: ${error.message}`, ui.ButtonSet.OK);
    throw error;
  }
}

/**
 * Generate predictability diff for ALL value streams
 */
function generatePredictabilityDiffAllValueStreams(piNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  
  const VALUE_STREAMS_LIST = [
    'AIMM', 'Cloud Foundation Services', 'Cloud Operations', 'Data Platform Engineering',
    'EMA Clinical', 'EMA RaC', 'Fusion and Conversions', 'MMGI', 'MMGI-Cloud', 'MMPM',
    'Patient Collaboration', 'Platform Engineering', 'RCM Genie', 'Security Engineering',
    'Shared Services Platform', 'Xtract'
  ];
  
  console.log(`\n${'='.repeat(80)}`);
  console.log(`PREDICTABILITY DIFF VALIDATION - ALL VALUE STREAMS`);
  console.log(`PI: ${piNumber}`);
  console.log(`Started: ${new Date().toISOString()}`);
  console.log(`${'='.repeat(80)}\n`);
  
  try {
    // Step 1: Get PI date range
    ss.toast('Step 1: Getting PI date range...', '⏳ Processing', 30);
    const dateRange = getPIDateRange(piNumber);
    console.log(`PI ${piNumber} execution window: ${dateRange.startDate.toDateString()} to ${dateRange.endDate.toDateString()}`);
    
    const allMissingCommitment = [];
    const allChangedPI = [];
    
    // Step 2: Process each value stream INDIVIDUALLY (to avoid hitting limits)
    for (let i = 0; i < VALUE_STREAMS_LIST.length; i++) {
      const vs = VALUE_STREAMS_LIST[i];
      ss.toast(`Step 2: Processing ${vs} (${i + 1}/${VALUE_STREAMS_LIST.length})...`, '⏳ Processing', 30);
      
      console.log(`\n${'─'.repeat(60)}`);
      console.log(`Processing Value Stream: ${vs}`);
      console.log(`${'─'.repeat(60)}`);
      
      // Find missing commitment epics for this VS (with pagination)
      const missingCommitment = findEpicsWithMissingCommitment(piNumber, vs);
      console.log(`  Missing commitment: ${missingCommitment.length}`);
      allMissingCommitment.push(...missingCommitment);
      
      // Find PI changed epics for this VS (with pagination)
      const changedPI = findEpicsWithPIChanges(piNumber, vs, dateRange);
      console.log(`  PI changed: ${changedPI.length}`);
      allChangedPI.push(...changedPI);
      
      // Small delay between value streams
      Utilities.sleep(300);
    }
    
    console.log(`\n${'─'.repeat(60)}`);
    console.log(`TOTALS ACROSS ALL VALUE STREAMS:`);
    console.log(`  Missing commitment: ${allMissingCommitment.length}`);
    console.log(`  PI changed: ${allChangedPI.length}`);
    console.log(`${'─'.repeat(60)}\n`);
    
    // Step 3: Check clone workflow compliance for all changed PI epics
    ss.toast('Step 3: Checking clone workflow compliance...', '⏳ Processing', 60);
    const epicsWithCloneCheck = checkCloneWorkflowCompliance(allChangedPI, piNumber);
    
    // Step 4: Combine and deduplicate
    ss.toast('Step 4: Combining results...', '⏳ Processing', 10);
    const combinedResults = combineValidationResults(allMissingCommitment, epicsWithCloneCheck);
    console.log(`Total unique issues: ${combinedResults.length}`);
    
    // Step 5: Write report
    ss.toast('Step 5: Creating report spreadsheet...', '⏳ Processing', 10);
    const sheetName = `PI ${piNumber} - Predictability Diff - All Value Streams`;
    const reportUrl = writeValidationReport(combinedResults, sheetName, piNumber, 'All Value Streams', dateRange);
    
    const duration = (Date.now() - startTime) / 1000;
    console.log(`\n${'='.repeat(80)}`);
    console.log(`VALIDATION COMPLETE in ${duration.toFixed(1)} seconds`);
    console.log(`${'='.repeat(80)}\n`);
    
    // Count stats
    const cloneProperCount = combinedResults.filter(r => r.cloneWorkflowStatus === 'PROPER').length;
    const cloneMissingCount = combinedResults.filter(r => r.cloneWorkflowStatus === 'NO_CLONE').length;
    const cloneWrongCount = combinedResults.filter(r => r.cloneWorkflowStatus === 'WRONG_EPIC_MOVED').length;
    
    ss.toast(`✅ Found ${combinedResults.length} issues. New spreadsheet created!`, '✅ Complete', 10);
    
    // Show completion dialog
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family: 'Segoe UI', Arial, sans-serif; padding: 10px;">
        <p style="font-size: 14px; margin-bottom: 15px;">
          Found <strong>${combinedResults.length}</strong> issues across all value streams.
        </p>
        <ul style="font-size: 13px; margin-bottom: 15px;">
          <li>Missing PI Commitment: <strong>${allMissingCommitment.length}</strong></li>
          <li>PI Changed During Execution: <strong>${allChangedPI.length}</strong></li>
        </ul>
        <p style="font-size: 13px; margin-bottom: 15px;"><strong>Clone Workflow:</strong></p>
        <ul style="font-size: 12px; margin-bottom: 15px;">
          <li>✅ Proper (clone left behind): <strong>${cloneProperCount}</strong></li>
          <li>⚠️ No clone found: <strong>${cloneMissingCount}</strong></li>
          <li>❌ Wrong epic moved: <strong>${cloneWrongCount}</strong></li>
        </ul>
        <p style="font-size: 14px; margin-bottom: 10px;">
          <a href="${reportUrl}" target="_blank" style="color: #1a73e8; font-weight: bold;">
            📊 Open Report Spreadsheet
          </a>
        </p>
        <p style="font-size: 11px; color: #666;">
          Completed in ${duration.toFixed(1)} seconds
        </p>
      </div>`
    )
    .setWidth(380)
    .setHeight(280);
    
    ui.showModalDialog(htmlOutput, '✅ Validation Complete (All Value Streams)');
      
  } catch (error) {
    console.error('Validation error:', error);
    console.error('Stack:', error.stack);
    ui.alert('Error', `Validation failed: ${error.message}`, ui.ButtonSet.OK);
    throw error;
  }
}

// ===== QUERY FUNCTIONS =====

/**
 * Find epics that currently have the PI assigned but are missing PI Commitment
 * Uses pagination to handle >100 results
 */
function findEpicsWithMissingCommitment(piNumber, valueStream) {
  const programIncrement = `PI ${piNumber}`;
  const piFieldId = FIELD_MAPPINGS.programIncrement.replace('customfield_', '');
  const vsFieldId = FIELD_MAPPINGS.valueStream.replace('customfield_', '');
  const commitFieldId = FIELD_MAPPINGS.piCommitment.replace('customfield_', '');
  
  // JQL: Epics with this PI and Value Stream where PI Commitment is empty
  const jql = `issuetype = Epic AND cf[${piFieldId}] = "${programIncrement}" AND cf[${vsFieldId}] = "${valueStream}" AND cf[${commitFieldId}] IS EMPTY`;
  
  console.log(`\n--- Finding epics missing PI Commitment ---`);
  console.log(`JQL: ${jql}`);
  
  // Use pagination to get all results
  const epics = searchJiraWithPagination(jql);
  
  console.log(`Found ${epics.length} epics missing PI Commitment`);
  
  // Add reason to each epic
  return epics.map(epic => ({
    ...epic,
    reason: 'PI Commitment field is empty',
    reasonCategory: 'Missing PI Commitment',
    cloneWorkflowStatus: 'N/A',
    cloneDetails: ''
  }));
}

/**
 * Find epics where PI was changed/removed during the PI execution window
 * Uses changelog analysis with pagination
 */
function findEpicsWithPIChanges(piNumber, valueStream, dateRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const programIncrement = `PI ${piNumber}`;
  const piFieldId = FIELD_MAPPINGS.programIncrement.replace('customfield_', '');
  const vsFieldId = FIELD_MAPPINGS.valueStream.replace('customfield_', '');
  
  // Format dates for JQL
  const startDateStr = Utilities.formatDate(dateRange.startDate, 'GMT', 'yyyy-MM-dd');
  const endDateStr = Utilities.formatDate(dateRange.endDate, 'GMT', 'yyyy-MM-dd');
  
  console.log(`\n--- Finding epics with PI changes ---`);
  console.log(`Looking for PI: "${programIncrement}"`);
  console.log(`Date range: ${startDateStr} to ${endDateStr}`);
  
  // Strategy: Query epics updated during the PI window that had this PI at some point
  // We need to get epics from two queries and then analyze changelogs
  
  // Query 1: Epics that CURRENTLY have this PI (to check if PI commitment changed)
  const currentJql = `issuetype = Epic AND cf[${piFieldId}] = "${programIncrement}" AND cf[${vsFieldId}] = "${valueStream}"`;
  
  // Query 2: Epics updated during PI window in this VS that DON'T currently have this PI
  // (these are potential de-commits)
  const updatedJql = `issuetype = Epic AND cf[${vsFieldId}] = "${valueStream}" AND updated >= "${startDateStr}" AND updated <= "${endDateStr}" AND cf[${piFieldId}] != "${programIncrement}"`;
  
  console.log(`Query 1 (current PI): ${currentJql}`);
  const currentPIEpics = searchJiraWithPagination(currentJql);
  console.log(`  Found ${currentPIEpics.length} epics currently with PI ${piNumber}`);
  
  console.log(`Query 2 (updated, no current PI): ${updatedJql}`);
  const updatedEpics = searchJiraWithPagination(updatedJql);
  console.log(`  Found ${updatedEpics.length} epics updated during PI without current PI`);
  
  // Combine unique keys
  const allKeys = new Set();
  currentPIEpics.forEach(e => allKeys.add(e.key));
  updatedEpics.forEach(e => allKeys.add(e.key));
  
  console.log(`Total unique epics to check changelogs: ${allKeys.size}`);
  
  if (allKeys.size === 0) {
    return [];
  }
  
  // Fetch changelogs and analyze
  ss.toast(`Fetching changelogs for ${allKeys.size} epics...`, '⏳ Processing', 60);
  const epicsWithChangelogs = fetchEpicChangelogsWithLinks(Array.from(allKeys));
  
  // Analyze changelogs to find PI changes during the window
  return analyzeChangelogsForPIChanges(epicsWithChangelogs, programIncrement, dateRange);
}

/**
 * Fetch changelogs AND issuelinks for epic keys (for clone detection)
 */
function fetchEpicChangelogsWithLinks(epicKeys) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jiraConfig = getJiraConfig();
  const allIssues = [];
  const BATCH_SIZE = 20;
  
  console.log(`\n--- Fetching changelogs for ${epicKeys.length} epics ---`);
  
  for (let i = 0; i < epicKeys.length; i += BATCH_SIZE) {
    const batch = epicKeys.slice(i, i + BATCH_SIZE);
    const progress = Math.min(i + BATCH_SIZE, epicKeys.length);
    
    if (i > 0) {
      ss.toast(`Fetching changelogs (${progress}/${epicKeys.length})...`, '⏳ Processing', 30);
    }
    
    const jql = `key IN (${batch.map(k => `"${k}"`).join(',')})`;
    
    // Build URL with changelog expansion and issuelinks for clone detection
    const fields = [
      'key', 'summary', 'issuetype', 'status', 'created', 'issuelinks',
      FIELD_MAPPINGS.programIncrement,
      FIELD_MAPPINGS.valueStream,
      FIELD_MAPPINGS.piCommitment,
      FIELD_MAPPINGS.scrumTeam,
      FIELD_MAPPINGS.allocation,
      FIELD_MAPPINGS.portfolioInitiative,
      FIELD_MAPPINGS.programInitiative,
      FIELD_MAPPINGS.businessValue,
      FIELD_MAPPINGS.actualValue
    ].join(',');
    
    const url = `${jiraConfig.baseUrl}/rest/api/3/search/jql?` +
                `jql=${encodeURIComponent(jql)}` +
                `&expand=changelog` +
                `&maxResults=${BATCH_SIZE}` +
                `&fields=${encodeURIComponent(fields)}`;
    
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: getJiraHeaders(),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        if (data.issues) {
          allIssues.push(...data.issues);
        }
      } else {
        console.error(`JIRA API error: ${response.getResponseCode()}`);
      }
    } catch (e) {
      console.error(`Failed to fetch changelog batch: ${e.message}`);
    }
    
    Utilities.sleep(300);
  }
  
  console.log(`Fetched ${allIssues.length} issues with changelogs`);
  return allIssues;
}

/**
 * Analyze changelogs to find PI changes during the PI window
 */
function analyzeChangelogsForPIChanges(issues, programIncrement, dateRange) {
  const results = [];
  const piFieldNames = ['Program Increment', 'Program Increment (PI)', 'customfield_10113'];
  const commitFieldNames = ['PI Commitment', 'customfield_10063'];
  const statusFieldNames = ['status', 'Status'];
  
  console.log(`\n--- Analyzing changelogs for PI changes ---`);
  console.log(`Looking for changes to: "${programIncrement}"`);
  console.log(`Window: ${dateRange.startDate.toISOString()} to ${dateRange.endDate.toISOString()}`);
  
  issues.forEach(issue => {
    if (!issue.changelog || !issue.changelog.histories) {
      return;
    }
    
    const fields = issue.fields || {};
    const currentPI = extractFieldValue(fields[FIELD_MAPPINGS.programIncrement]);
    const createdDate = fields.created ? new Date(fields.created) : null;
    
    // Extract clone info from issuelinks
    const cloneInfo = extractCloneInfo(issue);
    
    // Look through changelog for PI field changes during the window
    issue.changelog.histories.forEach(history => {
      const changeDate = new Date(history.created);
      
      // Check if change is within the PI window
      if (changeDate >= dateRange.startDate && changeDate <= dateRange.endDate) {
        history.items.forEach(item => {
          const isPI = piFieldNames.some(name => 
            item.field === name || 
            item.fieldId === 'customfield_10113' ||
            (item.field && item.field.toLowerCase().includes('program increment'))
          );
          
          if (isPI) {
            const fromValue = item.fromString || '';
            const toValue = item.toString || '';
            
            const hadThisPI = fromValue === programIncrement;
            const nowHasThisPI = currentPI === programIncrement;
            
            // We want epics where the PI was REMOVED (had it, but no longer has it)
            if (hadThisPI && !nowHasThisPI) {
              console.log(`  FOUND PI CHANGE: ${issue.key}`);
              console.log(`    Changed on: ${changeDate.toISOString()}`);
              console.log(`    From: "${fromValue}" → To: "${toValue || '(empty)'}"`);
              console.log(`    Is clone: ${cloneInfo.isClone}`);
              console.log(`    Created: ${createdDate}`);
              
              // Reconstruct historical values
              const historicalValues = reconstructHistoricalValues(
                issue, changeDate, commitFieldNames, statusFieldNames
              );
              
              // Parse the issue into our standard format
              const parsedIssue = parseJiraIssueForDiff(issue);
              parsedIssue.reason = `PI changed from "${fromValue}" to "${toValue || '(empty)'}"`;
              parsedIssue.reasonCategory = 'PI Changed During Execution';
              parsedIssue.changeDate = changeDate;
              parsedIssue.changeDateFormatted = formatDateTime(changeDate);
              parsedIssue.previousPI = fromValue;
              parsedIssue.newPI = toValue || '(empty)';
              parsedIssue.piCommitmentAtChange = historicalValues.piCommitmentAtChange;
              parsedIssue.statusAtChange = historicalValues.statusAtChange;
              parsedIssue.changedBy = history.author ? (history.author.displayName || history.author.name || '') : '';
              
              // Clone info
              parsedIssue.isClone = cloneInfo.isClone;
              parsedIssue.clonedFromKey = cloneInfo.clonedFromKey;
              parsedIssue.clonedByKeys = cloneInfo.clonedByKeys; // Issues that cloned this one
              parsedIssue.createdDate = createdDate;
              
              results.push(parsedIssue);
            }
          }
        });
      }
    });
  });
  
  console.log(`Found ${results.length} epics with PI removed during execution window`);
  return results;
}

/**
 * Extract clone information from issuelinks
 * Returns info about both directions:
 * - isClone: true if this issue was cloned FROM another
 * - clonedFromKey: the issue this was cloned from
 * - clonedByKeys: array of issues that were cloned FROM this one
 */
function extractCloneInfo(issue) {
  const fields = issue.fields || {};
  const issueLinks = fields.issuelinks || [];
  
  let isClone = false;
  let clonedFromKey = null;
  const clonedByKeys = [];
  
  for (const link of issueLinks) {
    const linkType = link.type || {};
    const linkTypeName = (linkType.name || '').toLowerCase();
    
    if (linkTypeName.includes('clone') || linkTypeName.includes('cloner')) {
      // inwardIssue = this issue "clones" another (this is a clone)
      if (link.inwardIssue) {
        isClone = true;
        clonedFromKey = link.inwardIssue.key;
      }
      
      // outwardIssue = another issue "clones" this one (this was cloned)
      if (link.outwardIssue) {
        clonedByKeys.push(link.outwardIssue.key);
      }
    }
  }
  
  return { isClone, clonedFromKey, clonedByKeys };
}

/**
 * Check clone workflow compliance for epics with PI changes
 * 
 * Correct workflow: When moving Epic A from PI 12 to PI 13:
 *   - Clone Epic A → Epic B
 *   - Epic B (clone) stays in PI 12, marked Deferred
 *   - Epic A (original) moves to PI 13
 * 
 * What we check:
 *   1. If this epic IS a clone and it was moved → WRONG (should have moved original)
 *   2. If this epic was cloned (has clonedByKeys) → Check if clone stayed in original PI
 *   3. If no clone relationship exists → NO_CLONE (missing audit trail)
 */
function checkCloneWorkflowCompliance(changedEpics, piNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jiraConfig = getJiraConfig();
  const programIncrement = `PI ${piNumber}`;
  
  console.log(`\n--- Checking clone workflow compliance ---`);
  console.log(`Checking ${changedEpics.length} epics with PI changes`);
  
  // Collect all clone keys we need to verify
  const cloneKeysToCheck = new Set();
  changedEpics.forEach(epic => {
    if (epic.clonedByKeys && epic.clonedByKeys.length > 0) {
      epic.clonedByKeys.forEach(k => cloneKeysToCheck.add(k));
    }
  });
  
  // Fetch current PI for all clones in one batch
  let clonePIMap = {};
  if (cloneKeysToCheck.size > 0) {
    console.log(`Fetching PI info for ${cloneKeysToCheck.size} clone(s)...`);
    clonePIMap = fetchClonePIInfo(Array.from(cloneKeysToCheck));
  }
  
  // Now evaluate each epic
  const results = changedEpics.map(epic => {
    const result = { ...epic };
    
    // Case 1: This epic IS a clone and it was moved
    if (epic.isClone) {
      result.cloneWorkflowStatus = 'WRONG_EPIC_MOVED';
      result.cloneDetails = `This epic is a CLONE of ${epic.clonedFromKey}. The clone should stay in original PI; the original should move.`;
      console.log(`  ${epic.key}: WRONG_EPIC_MOVED (is clone of ${epic.clonedFromKey})`);
      return result;
    }
    
    // Case 2: This epic was cloned - check if any clone stayed in original PI
    if (epic.clonedByKeys && epic.clonedByKeys.length > 0) {
      let cloneInOriginalPI = null;
      let cloneWithDeferral = null;
      
      for (const cloneKey of epic.clonedByKeys) {
        const clonePI = clonePIMap[cloneKey]?.pi || '';
        const cloneCommitment = clonePIMap[cloneKey]?.commitment || '';
        
        if (clonePI === programIncrement) {
          cloneInOriginalPI = cloneKey;
          
          // Check if commitment indicates deferral
          if (cloneCommitment.toLowerCase().includes('defer') || 
              cloneCommitment.toLowerCase().includes('cancel')) {
            cloneWithDeferral = cloneKey;
          }
        }
      }
      
      if (cloneWithDeferral) {
        result.cloneWorkflowStatus = 'PROPER';
        result.cloneDetails = `✅ Clone ${cloneWithDeferral} was left in ${programIncrement} with commitment "${clonePIMap[cloneWithDeferral]?.commitment}"`;
        console.log(`  ${epic.key}: PROPER (clone ${cloneWithDeferral} in original PI)`);
      } else if (cloneInOriginalPI) {
        result.cloneWorkflowStatus = 'PROPER';
        result.cloneDetails = `Clone ${cloneInOriginalPI} exists in ${programIncrement} (commitment: ${clonePIMap[cloneInOriginalPI]?.commitment || 'not set'})`;
        console.log(`  ${epic.key}: PROPER (clone ${cloneInOriginalPI} in original PI)`);
      } else {
        result.cloneWorkflowStatus = 'NO_CLONE';
        result.cloneDetails = `Clone(s) exist (${epic.clonedByKeys.join(', ')}) but none remain in ${programIncrement}`;
        console.log(`  ${epic.key}: NO_CLONE in original PI (clones: ${epic.clonedByKeys.join(', ')})`);
      }
      return result;
    }
    
    // Case 3: No clone relationship at all
    result.cloneWorkflowStatus = 'NO_CLONE';
    result.cloneDetails = 'No clone was created when PI was changed';
    console.log(`  ${epic.key}: NO_CLONE (no clone relationships found)`);
    return result;
  });
  
  // Summary
  const proper = results.filter(r => r.cloneWorkflowStatus === 'PROPER').length;
  const noClone = results.filter(r => r.cloneWorkflowStatus === 'NO_CLONE').length;
  const wrong = results.filter(r => r.cloneWorkflowStatus === 'WRONG_EPIC_MOVED').length;
  
  console.log(`\nClone workflow summary:`);
  console.log(`  PROPER: ${proper}`);
  console.log(`  NO_CLONE: ${noClone}`);
  console.log(`  WRONG_EPIC_MOVED: ${wrong}`);
  
  return results;
}

/**
 * Fetch current PI and PI Commitment for a list of clone keys
 */
function fetchClonePIInfo(cloneKeys) {
  const jiraConfig = getJiraConfig();
  const result = {};
  
  if (cloneKeys.length === 0) return result;
  
  const jql = `key IN (${cloneKeys.map(k => `"${k}"`).join(',')})`;
  const fields = `key,${FIELD_MAPPINGS.programIncrement},${FIELD_MAPPINGS.piCommitment}`;
  
  const url = `${jiraConfig.baseUrl}/rest/api/3/search/jql?` +
              `jql=${encodeURIComponent(jql)}` +
              `&maxResults=100` +
              `&fields=${encodeURIComponent(fields)}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: getJiraHeaders(),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.issues) {
        data.issues.forEach(issue => {
          result[issue.key] = {
            pi: extractFieldValue(issue.fields[FIELD_MAPPINGS.programIncrement]),
            commitment: extractFieldValue(issue.fields[FIELD_MAPPINGS.piCommitment])
          };
        });
      }
    }
  } catch (e) {
    console.error(`Failed to fetch clone PI info: ${e.message}`);
  }
  
  return result;
}

/**
 * Reconstruct historical field values at a specific point in time
 */
function reconstructHistoricalValues(issue, targetDate, commitFieldNames, statusFieldNames) {
  const fields = issue.fields || {};
  const changelog = issue.changelog;
  
  let piCommitmentValue = extractFieldValue(fields[FIELD_MAPPINGS.piCommitment]) || '';
  let statusValue = fields.status ? fields.status.name : '';
  
  if (!changelog || !changelog.histories) {
    return { piCommitmentAtChange: piCommitmentValue, statusAtChange: statusValue };
  }
  
  const sortedHistories = [...changelog.histories].sort((a, b) => 
    new Date(b.created) - new Date(a.created)
  );
  
  sortedHistories.forEach(history => {
    const historyDate = new Date(history.created);
    
    if (historyDate > targetDate) {
      history.items.forEach(item => {
        const isCommitField = commitFieldNames.some(name => 
          item.field === name || 
          item.fieldId === 'customfield_10063' ||
          (item.field && item.field.toLowerCase().includes('pi commitment'))
        );
        
        if (isCommitField) {
          piCommitmentValue = item.fromString || '';
        }
        
        const isStatusField = statusFieldNames.some(name => 
          item.field === name || item.fieldId === 'status'
        );
        
        if (isStatusField) {
          statusValue = item.fromString || '';
        }
      });
    }
  });
  
  return { piCommitmentAtChange: piCommitmentValue, statusAtChange: statusValue };
}

/**
 * Format date/time for display
 */
function formatDateTime(date) {
  if (!date) return '';
  
  const options = {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    hour12: true
  };
  
  return date.toLocaleString('en-US', options);
}

/**
 * Parse a JIRA issue for the diff report
 */
function parseJiraIssueForDiff(rawIssue) {
  const fields = rawIssue.fields || {};
  
  return {
    key: rawIssue.key,
    summary: fields.summary || '',
    status: fields.status ? fields.status.name : '',
    valueStream: extractFieldValue(fields[FIELD_MAPPINGS.valueStream]),
    scrumTeam: extractFieldValue(fields[FIELD_MAPPINGS.scrumTeam]),
    allocation: extractFieldValue(fields[FIELD_MAPPINGS.allocation]),
    portfolioInitiative: extractFieldValue(fields[FIELD_MAPPINGS.portfolioInitiative]),
    programInitiative: extractFieldValue(fields[FIELD_MAPPINGS.programInitiative]),
    programIncrement: extractFieldValue(fields[FIELD_MAPPINGS.programIncrement]),
    piCommitment: extractFieldValue(fields[FIELD_MAPPINGS.piCommitment]),
    businessValue: parseFloat(fields[FIELD_MAPPINGS.businessValue]) || 0,
    actualValue: parseFloat(fields[FIELD_MAPPINGS.actualValue]) || 0
  };
}

/**
 * Extract value from JIRA field
 */
function extractFieldValue(fieldValue) {
  if (!fieldValue) return '';
  if (typeof fieldValue === 'string') return fieldValue.trim();
  if (typeof fieldValue === 'number') return fieldValue.toString();
  if (Array.isArray(fieldValue)) {
    if (fieldValue.length === 0) return '';
    return extractFieldValue(fieldValue[0]);
  }
  if (typeof fieldValue === 'object') {
    if (fieldValue.value) return extractFieldValue(fieldValue.value);
    if (fieldValue.name) return String(fieldValue.name).trim();
    if (fieldValue.displayName) return String(fieldValue.displayName).trim();
    if (fieldValue.key) return String(fieldValue.key).trim();
  }
  return '';
}

// ===== RESULT PROCESSING =====

/**
 * Combine and deduplicate validation results
 */
function combineValidationResults(missingCommitment, changedPI) {
  const seenKeys = new Set();
  const combined = [];
  
  missingCommitment.forEach(epic => {
    if (!seenKeys.has(epic.key)) {
      seenKeys.add(epic.key);
      combined.push(epic);
    }
  });
  
  changedPI.forEach(epic => {
    if (!seenKeys.has(epic.key)) {
      seenKeys.add(epic.key);
      combined.push(epic);
    } else {
      const existing = combined.find(e => e.key === epic.key);
      if (existing && !existing.reason.includes('PI changed')) {
        existing.reason += ` | ${epic.reason}`;
        existing.reasonCategory = 'Both Issues';
        // Copy clone info
        existing.cloneWorkflowStatus = epic.cloneWorkflowStatus;
        existing.cloneDetails = epic.cloneDetails;
      }
    }
  });
  
  return combined;
}

// ===== REPORT WRITING =====

/**
 * Write validation results to a new standalone spreadsheet
 */
function writeValidationReport(results, sheetName, piNumber, valueStream, dateRange) {
  const jiraConfig = getJiraConfig();
  
  console.log(`Creating new spreadsheet: ${sheetName}`);
  const newSpreadsheet = SpreadsheetApp.create(sheetName);
  const newSpreadsheetUrl = newSpreadsheet.getUrl();
  const newSpreadsheetId = newSpreadsheet.getId();
  
  // Move to destination folder
  try {
    const file = DriveApp.getFileById(newSpreadsheetId);
    const destinationFolder = DriveApp.getFolderById(PREDICTABILITY_DIFF_FOLDER_ID);
    file.moveTo(destinationFolder);
    console.log(`✓ Moved spreadsheet to destination folder`);
  } catch (e) {
    console.warn(`Could not move to destination folder: ${e.message}`);
  }
  
  const sheet = newSpreadsheet.getSheets()[0];
  sheet.setName('Validation Results');
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily('Comfortaa');
  
  // Headers - includes clone workflow columns
  const headers = [
    'Epic Key',
    'Value Stream',
    'Scrum Team',
    'Allocation',
    'Portfolio Initiative',
    'Program Initiative',
    '(Current) Program Increment',
    '(Current) PI Commitment',
    '(Current) Status',
    '(Current) Business Value',
    '(Current) Actual Value',
    'Issue Category',
    'Change Date/Time',
    'Changed By',
    'Status (at change)',
    'PI Commitment (at change)',
    'Clone Workflow',
    'Clone Details',
    'Reason'
  ];
  
  // Title row
  sheet.getRange(1, 1).setValue(`Predictability Diff Validation - PI ${piNumber}`)
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1a73e8');
  
  sheet.getRange(2, 1).setValue(`Value Stream: ${valueStream} | PI Window: ${dateRange.startDate.toLocaleDateString()} to ${dateRange.endDate.toLocaleDateString()}`)
    .setFontSize(11)
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  sheet.getRange(3, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontSize(10)
    .setFontColor('#999999');
  
  // Summary counts
  const missingCommitCount = results.filter(r => r.reasonCategory === 'Missing PI Commitment').length;
  const changedPICount = results.filter(r => r.reasonCategory === 'PI Changed During Execution').length;
  const bothCount = results.filter(r => r.reasonCategory === 'Both Issues').length;
  const properCount = results.filter(r => r.cloneWorkflowStatus === 'PROPER').length;
  const noCloneCount = results.filter(r => r.cloneWorkflowStatus === 'NO_CLONE').length;
  const wrongCount = results.filter(r => r.cloneWorkflowStatus === 'WRONG_EPIC_MOVED').length;
  
  sheet.getRange(4, 1).setValue(
    `Summary: ${results.length} total | Missing Commitment: ${missingCommitCount} | PI Changed: ${changedPICount} | Both: ${bothCount} | ` +
    `Clone Workflow: ✅${properCount} ⚠️${noCloneCount} ❌${wrongCount}`
  )
    .setFontSize(10)
    .setFontWeight('bold')
    .setFontColor('#333333');
  
  // Header row
  sheet.getRange(6, 1, 1, headers.length).setValues([headers])
    .setBackground('#9b7bb8')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(10);
  
  // Data rows
  if (results.length > 0) {
    const dataRows = results.map(epic => {
      const epicUrl = `${jiraConfig.baseUrl}/browse/${epic.key}`;
      const hyperlink = `=HYPERLINK("${epicUrl}", "${epic.key}")`;
      
      const isPIChanged = epic.reasonCategory === 'PI Changed During Execution' || epic.reasonCategory === 'Both Issues';
      
      // Clone workflow display
      let cloneWorkflowDisplay = epic.cloneWorkflowStatus || 'N/A';
      if (cloneWorkflowDisplay === 'PROPER') cloneWorkflowDisplay = '✅ PROPER';
      else if (cloneWorkflowDisplay === 'NO_CLONE') cloneWorkflowDisplay = '⚠️ NO CLONE';
      else if (cloneWorkflowDisplay === 'WRONG_EPIC_MOVED') cloneWorkflowDisplay = '❌ WRONG EPIC MOVED';
      
      return [
        hyperlink,
        epic.valueStream || '',
        epic.scrumTeam || '',
        epic.allocation || '',
        epic.portfolioInitiative || '',
        epic.programInitiative || '',
        epic.programIncrement || '',
        epic.piCommitment || '',
        epic.status || '',
        epic.businessValue || 0,
        epic.actualValue || 0,
        epic.reasonCategory || '',
        isPIChanged ? (epic.changeDateFormatted || '') : '',
        isPIChanged ? (epic.changedBy || '') : '',
        isPIChanged ? (epic.statusAtChange || '') : '',
        isPIChanged ? (epic.piCommitmentAtChange || '') : '',
        isPIChanged ? cloneWorkflowDisplay : 'N/A',
        epic.cloneDetails || '',
        epic.reason || ''
      ];
    });
    
    sheet.getRange(7, 1, dataRows.length, headers.length).setValues(dataRows)
      .setFontSize(10);
    
    // Color-coding
    const categoryCol = headers.indexOf('Issue Category') + 1;
    const cloneWorkflowCol = headers.indexOf('Clone Workflow') + 1;
    
    for (let i = 0; i < results.length; i++) {
      const rowNum = 7 + i;
      const result = results[i];
      
      // Category coloring
      const categoryCell = sheet.getRange(rowNum, categoryCol);
      if (result.reasonCategory === 'Missing PI Commitment') {
        categoryCell.setBackground('#fff3cd').setFontColor('#856404');
      } else if (result.reasonCategory === 'PI Changed During Execution') {
        categoryCell.setBackground('#f8d7da').setFontColor('#721c24');
      } else if (result.reasonCategory === 'Both Issues') {
        categoryCell.setBackground('#f5c6cb').setFontColor('#721c24');
      }
      
      // Clone workflow coloring
      const cloneCell = sheet.getRange(rowNum, cloneWorkflowCol);
      if (result.cloneWorkflowStatus === 'PROPER') {
        cloneCell.setBackground('#d4edda').setFontColor('#155724');
      } else if (result.cloneWorkflowStatus === 'NO_CLONE') {
        cloneCell.setBackground('#fff3cd').setFontColor('#856404');
      } else if (result.cloneWorkflowStatus === 'WRONG_EPIC_MOVED') {
        cloneCell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
      }
    }
    
    // Format numeric columns
    const bvCol = headers.indexOf('(Current) Business Value') + 1;
    const avCol = headers.indexOf('(Current) Actual Value') + 1;
    sheet.getRange(7, bvCol, dataRows.length, 1).setNumberFormat('#,##0');
    sheet.getRange(7, avCol, dataRows.length, 1).setNumberFormat('#,##0');
    
    // Darker purple for historical columns header
    const changeDateCol = headers.indexOf('Change Date/Time') + 1;
    sheet.getRange(6, changeDateCol, 1, 4).setBackground('#6c5b8a');
    
  } else {
    sheet.getRange(7, 1).setValue('✅ No predictability issues found!')
      .setFontSize(12)
      .setFontColor('#28a745')
      .setFontWeight('bold');
  }
  
  // Freeze header rows
  sheet.setFrozenRows(6);
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Set specific column widths
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(headers.indexOf('Clone Details') + 1, 300);
  sheet.setColumnWidth(headers.indexOf('Reason') + 1, 280);
  
  console.log(`✅ Report written: ${newSpreadsheetUrl}`);
  
  return newSpreadsheetUrl;
}

// ===== MENU REGISTRATION =====

function addPredictabilityDiffMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Predictability Diff')
    .addItem('Run Validation...', 'showPredictabilityDiffDialog')
    .addToUi();
}