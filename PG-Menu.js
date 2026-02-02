// ═══════════════════════════════════════════════════════════════════════════════
// MENU.GS - Menu System, Dialogs, and Pipeline Functions
// ═══════════════════════════════════════════════════════════════════════════════
//
// This file contains:
// 1. Menu registration (onOpen)
// 2. Dialog display functions
// 3. Report generation functions
// 4. Pipeline functions (sequential and parallel)
// 5. Scheduler functions
// 6. Utility/helper functions
//
// ═══════════════════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 1: MENU REGISTRATION
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Program Governance')
    // Individual Steps (for manual/debugging)
    //.addItem('📊 STEP 1: Generate Iteration Report', 'showIterationReportDialog')
    .addItem('⚡ STEP 1: Run Full Pipeline (Parallel)', 'showParallelPipelineDialog')
    //.addItem('🔗 STEP 2: Enrich with Initiative Data', 'enrichWithInitiativeData')
    .addItem('🔍 STEP 2: Analyze Changes for Iteration', 'showComparisonDialog')
    .addItem('🎯 STEP 3: Generate Governance Slides', 'showSlideGeneratorDialog')
    //.addItem('🔍 STEP 3: Generate Epic Perspectives', 'generateEpicPerspectives')
    //.addItem('🔗 STEP 4: Generate Initiative Perspectives', 'generateInitiativePerspectives')
    //.addSeparator()
    // Full Pipeline Options
    //.addItem('🚀 Run Full Pipeline (Sequential)', 'showFullPipelineDialog')
    .addSeparator()
    // Scheduler
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⏰ Scheduler')
        .addItem('🔧 Setup Pipeline Scheduler', 'setupPipelineScheduler')
        .addItem('⚙️ Update Scheduler Config', 'updateSchedulerConfig')
        .addItem('🧪 Test Scheduled Run', 'testScheduledPipeline')
        .addItem('📋 View Scheduler Log', 'viewSchedulerLog')
        .addItem('🗑️ Remove Scheduler', 'removeScheduler')
    )
    //.addSeparator()
    //.addItem('🔄 Add Child Ticket Metrics', 'showRefreshChildMetricsDialog')
    //.addItem('🔍 Analyze Changes for Iteration', 'showComparisonDialog')
    //.addSeparator()
    //.addItem('🎯 Generate Governance Slides', 'showSlideGeneratorDialog')
    .addItem('✅ Generate Closed Items Report', 'showClosedItemsReportDialog')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('PI Deferral Metrics')
        .addItem('∑ Generate PI Deferrals - AI Summary', 'summarizeDeferrals')
        .addItem('🔴 Generate PI Deferrals Slides', 'generateDeferralsSlide')
    )              
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📈 Planning Reports')
        .addItem('📋 Generate Initial Plan Review', 'showInitialPlanDialog')
        .addItem('✅ Generate Final Plan Review - VS ', 'showFinalPlanDialog')
        .addItem('🎯 Generate Final Plan - non-VS', 'showInitiativeDialog')
    )
    .addSeparator()  
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📈 Predictability Reports')
        .addItem('📊 Update Predictability Score', 'showPredictabilityDialog')
        .addItem('🎯 Generate Predictability Presentation', 'showPredictabilityPresentationDialog')
        .addItem('🔍 Validate Predictability', 'showPredictabilityValidationDialog')
        .addItem('Predictability Diff Validation...', 'showPredictabilityDiffDialog')
        .addSeparator()
        .addItem('📈 Generate VS Charts Sheet', 'generateAllValueStreamChartsSheet')
        .addItem('🔄 Refresh Chart Boxes (1,3,6)', 'refreshFormattedBoxes')
        .addItem('✅ Check VS Alignment', 'checkPredictabilityAlignment')
    )
    .addSeparator()  
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📝 Activity Log')
        .addItem('📊 View Activity Summary', 'showActivitySummary')
        .addItem('🗑️ Clear Old Logs', 'clearOldActivityLogs')
        .addItem('📈 View Full Log', 'openActivityLogSheet')
    )
    .addSeparator()
    .addItem('⚙️ Configure JIRA', 'showJiraConfigurationDialog')
    .addItem('⚙️ Configure Gemini AI', 'setupGeminiConfiguration')
    .addItem('📊 Configure Portfolio Slides', 'showPortfolioConfigDialog')
    .addItem('⚙️ Empty Field Display Settings', 'openEmptyFieldConfig')
    .addItem('🔄 Refresh Config Metadata', 'refreshFieldMetadata') 
    .addItem('🔄 Re-apply Governance Filter', 'showReapplyGovernanceFilterDialog')
    .addToUi();
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 2: DIALOG DISPLAY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Shows the main dialog for generating a governance report for a specific iteration.
 */
function showIterationReportDialog() {
  const html = HtmlService.createHtmlOutput(getIterationReportDialogHTML())
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Governance Report');
}

/**
 * Shows the sequential full pipeline dialog
 */
function showFullPipelineDialog() {
  const html = HtmlService.createHtmlOutput(getFullPipelineDialogHTML())
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 Run Full Pipeline');
}

/**
 * Shows the parallel pipeline dialog
 */
function showParallelPipelineDialog() {
  const html = HtmlService.createHtmlOutput(getParallelPipelineDialogHTML())
    .setWidth(500)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ Run Parallel Pipeline');
}

/**
 * Shows the slide generator dialog
 */
function showSlideGeneratorDialog() {
  const html = HtmlService.createHtmlOutput(getSlideGeneratorHTML())
    .setWidth(500)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Governance Slides');
}

/**
 * Shows the refresh value stream dialog
 */
function showRefreshValueStreamDialog() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet.getName().includes('Governance')) {
    ui.alert('Error', 'Please select a Governance sheet first', ui.ButtonSet.OK);
    return;
  }

  const piMatch = sheet.getName().match(/PI (\d+)/);
  if (!piMatch) {
    ui.alert('Error', 'Could not determine PI number from sheet name.', ui.ButtonSet.OK);
    return;
  }
  const piNumber = piMatch[1];
  
  const html = HtmlService.createHtmlOutputFromFile('RefreshDialog.html')
      .setWidth(400)
      .setHeight(450);
  ui.showModalDialog(html, `Refresh Value Stream for PI ${piNumber}`);
}

/**
 * Shows the predictability dialog
 */
function showPredictabilityDialog() {
  const html = HtmlService.createHtmlOutput(getPredictabilityDialogHTML())
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Update Predictability Score');
}

/**
 * Shows the refresh child metrics dialog
 */
function showRefreshChildMetricsDialog() {
  const html = HtmlService.createHtmlOutput(getRefreshChildMetricsHTML())
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '🔄 Refresh Child Ticket Metrics');
}

/**
 * Shows the predictability validation dialog
 */
function showPredictabilityValidationDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get available PIs
  const availablePIs = [];
  for (let pi = 11; pi <= 15; pi++) {
    const sheetName = `PI ${pi} - Iteration 6`;
    if (ss.getSheetByName(sheetName)) {
      availablePIs.push(pi);
    }
  }
  
  if (availablePIs.length === 0) {
    ui.alert('No Data', 'No Iteration 6 sheets found for PI 11-15.', ui.ButtonSet.OK);
    return;
  }
  
  // Get value streams
  const configSheet = ss.getSheetByName('Configuration Template');
  if (!configSheet) {
    ui.alert('Error', 'Configuration Template sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  const configData = configSheet.getDataRange().getValues();
  const valueStreams = [];
  for (let i = 5; i < configData.length; i++) {
    const vs = configData[i][0];
    if (vs && vs.toString().trim() !== '') {
      valueStreams.push(vs.toString().trim());
    }
  }
  
  if (valueStreams.length === 0) {
    ui.alert('Error', 'No Value Streams found in Configuration Template.', ui.ButtonSet.OK);
    return;
  }
  
  // Show HTML dialog
  const html = getValidationDialogHTML(availablePIs, valueStreams);
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600);
  ui.showModalDialog(htmlOutput, '🔍 Predictability Validation Report');
}

/**
 * Shows allocation diagnostic dialog
 */
function showAllocationDiagnosticDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get available PIs (11-15)
  const availablePIs = [];
  for (let pi = 11; pi <= 15; pi++) {
    const sheetName = `PI ${pi} - Iteration 6`;
    if (ss.getSheetByName(sheetName)) {
      availablePIs.push(pi);
    }
  }
  
  if (availablePIs.length === 0) {
    ui.alert('No Data', 'No Iteration 6 sheets found for PI 11-15.\n\nPlease generate Iteration 6 data first using "Generate Iteration Report".', ui.ButtonSet.OK);
    return;
  }
  
  // Prompt for PI selection
  const piPrompt = ui.prompt(
    'Allocation Distribution Diagnostic',
    `This diagnostic identifies child tickets (Stories/Bugs) that are counted in ART Velocity\nbut excluded from Allocation Distribution due to missing or invalid Epic allocation types.\n\nAvailable PIs: ${availablePIs.join(', ')}\n\nEnter PI number to analyze:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (piPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const piNumber = parseInt(piPrompt.getResponseText().trim());
  
  if (!availablePIs.includes(piNumber)) {
    ui.alert('Invalid PI', `PI ${piNumber} does not have an Iteration 6 sheet.\n\nAvailable PIs: ${availablePIs.join(', ')}`, ui.ButtonSet.OK);
    return;
  }
  
  // Generate diagnostic report
  ss.toast(`Analyzing allocation distribution for PI ${piNumber}...`, '🩺 Diagnostic', -1);
  
  try {
    generateAllocationDiagnosticReport(piNumber);
  } catch (error) {
    console.error('Error generating allocation diagnostic:', error);
    ui.alert('Error', `Failed to generate diagnostic report:\n\n${error.message}`, ui.ButtonSet.OK);
    ss.toast('', '', 1);
  }
}

/**
 * Shows reapply governance filter dialog
 */
function showReapplyGovernanceFilterDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Re-apply Governance Filter',
    'Enter the PI number to re-apply governance filtering:',
    ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const piNumber = parseInt(response.getResponseText().trim());
    if (isNaN(piNumber)) {
      ui.alert('Error', 'Please enter a valid PI number.', ui.ButtonSet.OK);
      return;
    }
    reapplyGovernanceFilter(piNumber);
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 3: DIALOG HTML GENERATORS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generates HTML for the Parallel Pipeline dialog
 */
function getParallelPipelineDialogHTML() {
  const valueStreamCheckboxes = VALUE_STREAMS.map(vs => {
    const vsId = vs.replace(/[^a-zA-Z0-9]/g, '');
    return `
      <div class="checkbox-item" onclick="toggleCheckbox('${vsId}')">
        <input type="checkbox" id="vs_${vsId}" value="${vs}" checked onchange="updateSelectAll()">
        <label for="vs_${vsId}">${vs}</label>
      </div>
    `;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Google Sans', 'Roboto', Arial, sans-serif; padding: 20px; background: #fff; }
          .container { max-width: 480px; margin: 0 auto; }
          h2 { color: #7b1fa2; margin-bottom: 8px; font-size: 20px; font-weight: 500; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 16px; }
          
          .info-box { 
            background: linear-gradient(135deg, #f3e5f5 0%, #e8eaf6 100%); 
            border-radius: 8px; 
            padding: 12px 16px; 
            margin-bottom: 16px; 
            border-left: 4px solid #7b1fa2;
          }
          .info-box strong { color: #7b1fa2; }
          .info-box ul { margin: 8px 0 0 20px; font-size: 13px; color: #5f6368; }
          .info-box li { margin: 4px 0; }
          
          .section { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
          .section-title { color: #3c4043; font-weight: 500; margin-bottom: 12px; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }
          
          .section-purple {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 16px;
            margin-bottom: 16px;
            border: 2px solid #7b1fa2;
          }
          .section-title-purple {
            background: #7b1fa2;
            color: white;
            padding: 8px 12px;
            border-radius: 4px;
            margin: -16px -16px 12px -16px;
            font-weight: 500;
            font-size: 14px;
          }
          
          .input-group { display: flex; gap: 12px; }
          .input-wrapper { flex: 1; }
          .input-label { display: block; font-size: 12px; color: #5f6368; margin-bottom: 4px; }
          input[type="number"] { 
            width: 100%; 
            padding: 12px; 
            font-size: 16px; 
            border: 2px solid #dadce0; 
            border-radius: 8px; 
            transition: border-color 0.2s;
          }
          input[type="number"]:focus { outline: none; border-color: #7b1fa2; }
          
          .checkbox-group { 
            background: white; 
            border-radius: 8px; 
            padding: 8px; 
            max-height: 140px; 
            overflow-y: auto; 
            border: 1px solid #e0e0e0;
          }
          .checkbox-item { 
            display: flex; 
            align-items: center; 
            padding: 6px 8px; 
            margin: 2px 0; 
            border-radius: 6px; 
            cursor: pointer; 
            transition: background 0.15s;
          }
          .checkbox-item:hover { background: #f1f3f4; }
          .checkbox-item.select-all { 
            background: #f3e5f5; 
            margin-bottom: 8px; 
            border-bottom: 1px solid #e1bee7; 
            font-weight: 500; 
          }
          .checkbox-item input { 
            width: 18px; 
            height: 18px; 
            margin-right: 12px; 
            accent-color: #7b1fa2; 
            cursor: pointer;
          }
          .checkbox-item label { cursor: pointer; font-size: 14px; }
          
          .radio-group { display: flex; flex-direction: column; gap: 8px; }
          .radio-option { 
            display: flex; 
            align-items: center; 
            padding: 10px 12px; 
            background: white;
            border: 2px solid #dadce0; 
            border-radius: 8px; 
            cursor: pointer; 
            transition: all 0.2s; 
          }
          .radio-option:hover { border-color: #9c27b0; background: #faf5fc; }
          .radio-option.selected { border-color: #7b1fa2; background: #f3e5f5; }
          .radio-option input[type="radio"] { 
            width: 18px; 
            height: 18px; 
            margin-right: 12px; 
            accent-color: #7b1fa2; 
            cursor: pointer;
          }
          .radio-option label { 
            cursor: pointer; 
            flex: 1; 
            font-size: 14px; 
            color: #3c4043; 
          }
          
          .existing-sheet-container {
            margin-top: 12px;
            display: none;
          }
          .existing-sheet-container.visible { display: block; }
          .existing-sheet-label {
            display: block;
            font-size: 12px;
            color: #5f6368;
            margin-bottom: 4px;
          }
          .existing-sheet-select {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #dadce0;
            border-radius: 8px;
            font-size: 14px;
            background: white;
            cursor: pointer;
            transition: border-color 0.2s;
          }
          .existing-sheet-select:focus { outline: none; border-color: #7b1fa2; }
          
          .btn-group { display: flex; gap: 12px; margin-top: 20px; }
          .btn { 
            flex: 1; 
            padding: 14px 24px; 
            border: none; 
            border-radius: 8px; 
            font-size: 14px; 
            cursor: pointer; 
            font-weight: 500; 
            transition: all 0.2s; 
          }
          .btn-primary { 
            background: linear-gradient(135deg, #9c27b0 0%, #7b1fa2 100%); 
            color: white; 
            box-shadow: 0 2px 8px rgba(123, 31, 162, 0.3);
          }
          .btn-primary:hover { 
            background: linear-gradient(135deg, #8e24aa 0%, #6a1b9a 100%); 
            transform: translateY(-1px); 
            box-shadow: 0 4px 12px rgba(123, 31, 162, 0.4);
          }
          .btn-primary:disabled { 
            background: #bdbdbd; 
            cursor: not-allowed; 
            transform: none; 
            box-shadow: none;
          }
          .btn-secondary { background: #f1f3f4; color: #5f6368; }
          .btn-secondary:hover { background: #e8eaed; }
          
          .note { 
            font-size: 12px; 
            color: #9e9e9e; 
            text-align: center; 
            margin-top: 16px; 
            padding: 8px;
            border-radius: 4px;
            transition: all 0.3s;
          }
          .note.running { background: #fff3e0; color: #e65100; }
          .note.success { background: #e8f5e9; color: #2e7d32; }
          .note.error { background: #ffebee; color: #c62828; }
          .note a { color: #7b1fa2; font-weight: 500; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>⚡ Parallel Pipeline</h2>
          <p class="subtitle">Run the full governance pipeline with parallel JIRA fetching</p>
          
          <div class="info-box">
            <strong>Pipeline Steps:</strong>
            <ul>
              <li>Step 1: Fetch epics, dependencies & initiatives (parallel)</li>
              <li>Step 2: Generate epic perspectives (AI)</li>
              <li>Step 3: Generate initiative perspectives (AI)</li>
            </ul>
          </div>
          
          <div class="section">
            <div class="section-title">Configuration</div>
            <div class="input-group">
              <div class="input-wrapper">
                <span class="input-label">PI Number</span>
                <input type="number" id="piNumber" value="13" min="1" max="99">
              </div>
              <div class="input-wrapper">
                <span class="input-label">Iteration</span>
                <input type="number" id="iterationNumber" value="2" min="0" max="6">
              </div>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Value Streams</div>
            <div class="checkbox-group">
              <div class="checkbox-item select-all" onclick="toggleSelectAll()">
                <input type="checkbox" id="selectAll" checked>
                <label for="selectAll">Select All</label>
              </div>
              ${valueStreamCheckboxes}
            </div>
          </div>
          
          <div class="section-purple">
            <div class="section-title-purple">📁 Report Destination</div>
            <div class="radio-group">
              <div class="radio-option selected" onclick="selectDestination('current')">
                <input type="radio" id="dest_current" name="destination" value="current" checked>
                <label for="dest_current">Add to Current Spreadsheet</label>
              </div>
              <div class="radio-option" onclick="selectDestination('new')">
                <input type="radio" id="dest_new" name="destination" value="new">
                <label for="dest_new">Create New Report Spreadsheet</label>
              </div>
              <div class="radio-option" onclick="selectDestination('existing')">
                <input type="radio" id="dest_existing" name="destination" value="existing">
                <label for="dest_existing">Update Existing Report</label>
              </div>
            </div>
            <div id="existingSheetContainer" class="existing-sheet-container">
              <span class="existing-sheet-label">Select sheet to update:</span>
              <select id="existingSheetSelect" class="existing-sheet-select">
                <option value="">Loading sheets...</option>
              </select>
            </div>
          </div>
          
          <div class="btn-group">
            <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
            <button class="btn btn-primary" id="runBtn" onclick="runPipeline()">⚡ Run Pipeline</button>
          </div>
          
          <p class="note" id="statusNote">Progress will be shown in toast notifications</p>
        </div>
        
        <script>
          let selectedDestination = 'current';
          let existingSheetsLoaded = false;
          
          function toggleCheckbox(id) {
            const cb = document.getElementById('vs_' + id);
            if (cb) {
              cb.checked = !cb.checked;
              updateSelectAll();
            }
          }
          
          function toggleSelectAll() {
            const selectAll = document.getElementById('selectAll');
            const checkboxes = document.querySelectorAll('input[id^="vs_"]');
            checkboxes.forEach(cb => cb.checked = selectAll.checked);
          }
          
          function updateSelectAll() {
            const checkboxes = document.querySelectorAll('input[id^="vs_"]');
            const allChecked = Array.from(checkboxes).every(cb => cb.checked);
            document.getElementById('selectAll').checked = allChecked;
          }
          
          function selectDestination(dest) {
            selectedDestination = dest;
            
            // Update radio buttons
            document.getElementById('dest_current').checked = (dest === 'current');
            document.getElementById('dest_new').checked = (dest === 'new');
            document.getElementById('dest_existing').checked = (dest === 'existing');
            
            // Update visual selection
            document.querySelectorAll('.radio-option').forEach(opt => opt.classList.remove('selected'));
            event.currentTarget.classList.add('selected');
            
            // Show/hide existing sheet dropdown
            const container = document.getElementById('existingSheetContainer');
            if (dest === 'existing') {
              container.classList.add('visible');
              if (!existingSheetsLoaded) {
                loadExistingSheets();
              }
            } else {
              container.classList.remove('visible');
            }
          }
          
          function loadExistingSheets() {
            const select = document.getElementById('existingSheetSelect');
            select.innerHTML = '<option value="">Loading sheets...</option>';
            select.disabled = true;
            
            google.script.run
              .withSuccessHandler(function(sheets) {
                select.innerHTML = '<option value="">-- Select a sheet --</option>';
                if (sheets && sheets.length > 0) {
                  sheets.forEach(function(sheetName) {
                    const opt = document.createElement('option');
                    opt.value = sheetName;
                    opt.textContent = sheetName;
                    select.appendChild(opt);
                  });
                } else {
                  select.innerHTML = '<option value="">No iteration sheets found</option>';
                }
                select.disabled = false;
                existingSheetsLoaded = true;
              })
              .withFailureHandler(function(err) {
                select.innerHTML = '<option value="">Error loading sheets</option>';
                select.disabled = false;
                console.error('Failed to load sheets:', err);
              })
              .getExistingIterationSheets();
          }
          
          function validateInputs() {
            const pi = parseInt(document.getElementById('piNumber').value);
            const iter = parseInt(document.getElementById('iterationNumber').value);
            
            if (isNaN(pi) || pi < 1 || pi > 99) {
              alert('Please enter a valid PI number (1-99)');
              return null;
            }
            
            if (isNaN(iter) || iter < 0 || iter > 6) {
              alert('Please enter a valid iteration number (0-6)');
              return null;
            }
            
            const checkboxes = document.querySelectorAll('input[id^="vs_"]:checked');
            const selectedStreams = Array.from(checkboxes).map(cb => cb.value);
            
            if (selectedStreams.length === 0) {
              alert('Please select at least one value stream');
              return null;
            }
            
            let existingSheetName = null;
            if (selectedDestination === 'existing') {
              existingSheetName = document.getElementById('existingSheetSelect').value;
              if (!existingSheetName) {
                alert('Please select an existing sheet to update');
                return null;
              }
            }
            
            return {
              pi: pi,
              iteration: iter,
              valueStreams: selectedStreams,
              destinationConfig: {
                type: selectedDestination,
                existingSheetName: existingSheetName
              }
            };
          }
          
          function setRunningState(running) {
            const btn = document.getElementById('runBtn');
            const note = document.getElementById('statusNote');
            
            if (running) {
              btn.disabled = true;
              btn.textContent = '⏳ Running...';
              note.className = 'note running';
              note.innerHTML = '⏳ <strong>Pipeline running...</strong> Watch for toast notifications in the spreadsheet.';
            } else {
              btn.disabled = false;
              btn.textContent = '⚡ Run Pipeline';
              note.className = 'note';
            }
          }
          
          function runPipeline() {
            const inputs = validateInputs();
            if (!inputs) return;
            
            setRunningState(true);
            
            google.script.run
              .withSuccessHandler(function(result) {
                const btn = document.getElementById('runBtn');
                const note = document.getElementById('statusNote');
                
                btn.textContent = '✅ Complete!';
                btn.style.background = 'linear-gradient(135deg, #4caf50 0%, #388e3c 100%)';
                note.className = 'note success';
                
                if (result && result.newSpreadsheetUrl) {
                  note.innerHTML = '✅ Pipeline complete! <a href="' + result.newSpreadsheetUrl + '" target="_blank">Open new spreadsheet →</a>';
                } else {
                  const rows = result ? result.totalRows : 'unknown';
                  const time = result ? result.timeSeconds : 'unknown';
                  note.innerHTML = '✅ Pipeline complete! Processed ' + rows + ' rows in ' + time + 's';
                }
                
                // Auto-close after 4 seconds
                setTimeout(function() { 
                  google.script.host.close(); 
                }, 4000);
              })
              .withFailureHandler(function(err) {
                setRunningState(false);
                
                const note = document.getElementById('statusNote');
                note.className = 'note error';
                note.innerHTML = '❌ Error: ' + (err.message || err.toString());
              })
              .runFullPipelineParallel(
                inputs.pi, 
                inputs.iteration, 
                inputs.valueStreams, 
                inputs.destinationConfig
              );
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Generates HTML for the Iteration Change Analysis dialog
 */
function getIterationDialogHTML() {
  const valueStreamCheckboxes = VALUE_STREAMS.map(vs => {
    const vsId = vs.replace(/[^a-zA-Z0-9]/g, '');
    return `
      <div class="checkbox-item" onclick="toggleCheckbox('${vsId}')">
        <input type="checkbox" id="vs_${vsId}" value="${vs}" onchange="updateSelectAll()">
        <label for="vs_${vsId}">${vs}</label>
      </div>
    `;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Google Sans', 'Roboto', Arial, sans-serif; padding: 20px; background: #fff; }
          .container { max-width: 450px; margin: 0 auto; }
          h2 { color: #1a73e8; margin-bottom: 8px; font-size: 20px; font-weight: 500; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 24px; }
          .section { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 20px; }
          .section-title { color: #3c4043; font-weight: 500; margin-bottom: 12px; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }
          .input-group { margin-bottom: 16px; }
          .pi-input, .iteration-select { width: 100%; padding: 12px; font-size: 16px; border: 2px solid #dadce0; border-radius: 8px; transition: border-color 0.2s; background: white; }
          .pi-input:focus, .iteration-select:focus { outline: none; border-color: #1a73e8; }
          .checkbox-group { background: white; border-radius: 8px; padding: 12px; max-height: 200px; overflow-y: auto; }
          .checkbox-item { display: flex; align-items: center; padding: 8px; margin: 2px 0; border-radius: 6px; transition: background 0.2s; cursor: pointer; user-select: none; }
          .checkbox-item:hover { background: #f1f3f4; }
          .checkbox-item.select-all { background: #e8f0fe; margin-bottom: 8px; border-bottom: 1px solid #dadce0; font-weight: 500; }
          .checkbox-item input[type="checkbox"] { width: 18px; height: 18px; margin-right: 12px; cursor: pointer; accent-color: #1a73e8; }
          .checkbox-item label { cursor: pointer; flex: 1; color: #3c4043; font-size: 14px; }
          .info-box { background: #e3f2fd; border-left: 3px solid #1976d2; padding: 12px; margin: 16px 0; border-radius: 4px; }
          .info-box p { color: #1565c0; font-size: 13px; margin: 0; }
          .button-container { display: flex; gap: 12px; justify-content: flex-end; margin-top: 24px; }
          button { padding: 10px 24px; font-size: 14px; border-radius: 6px; border: none; cursor: pointer; font-weight: 500; transition: all 0.2s; font-family: inherit; }
          .primary-button { background: #1a73e8; color: white; }
          .primary-button:hover:not(:disabled) { background: #1557b0; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
          .primary-button:disabled { background: #94c1f5; cursor: not-allowed; }
          .secondary-button { background: #ffffff; color: #5f6368; border: 1px solid #dadce0; }
          .secondary-button:hover { background: #f8f9fa; }
          .error { color: #d93025; font-size: 12px; margin-top: 8px; display: none; }
          .error.show { display: block; }
          .loading { display: none; text-align: center; padding: 20px; }
          .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1a73e8; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Track Iteration Changes</h2>
          <div class="subtitle">Analyze what changed during a specific iteration</div>
          <div class="section">
            <div class="section-title">Program Increment</div>
            <div class="input-group">
              <input type="text" id="piNumber" class="pi-input" placeholder="Enter PI number (e.g., 11, 12, 13)" autocomplete="off" onkeypress="return event.charCode >= 48 && event.charCode <= 57" />
              <div id="piError" class="error">Please enter a valid PI number</div>
            </div>
          </div>
          <div class="section">
            <div class="section-title">Iteration</div>
            <div class="input-group">
              <select id="iterationNumber" class="iteration-select">
                <option value="">Select iteration...</option>
                <option value="0">Iteration 0 (Pre-Planning) 📋</option>
                <option value="1">Iteration 1 (2 weeks)</option>
                <option value="2">Iteration 2 (2 weeks)</option>
                <option value="3">Iteration 3 (2 weeks)</option>
                <option value="4">Iteration 4 (2 weeks)</option>
                <option value="5">Iteration 5 (2 weeks)</option>
                <option value="6">Iteration 6 (3 weeks - IP)</option>
              </select>
              <div id="iterationError" class="error">Please select an iteration</div>
            </div>
          </div>
          <div class="section">
            <div class="section-title">Value Streams</div>
            <div class="checkbox-group">
              <div class="checkbox-item select-all" onclick="toggleAll(this)">
                <input type="checkbox" id="selectAll" onchange="toggleAll(this)">
                <label for="selectAll">Select All Value Streams</label>
              </div>
              ${valueStreamCheckboxes}
            </div>
          </div>
          <div class="info-box">
            <p>📊 This will analyze all changes made during the selected iteration period plus 3 days after the iteration ends.</p>
          </div>
          <div id="loading" class="loading">
            <div class="spinner"></div>
            <p style="margin-top: 16px; color: #5f6368;">Analyzing changes...<br>This may take a few moments.</p>
          </div>
          <div class="button-container">
            <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
            <button id="analyzeBtn" class="primary-button" onclick="analyzeChanges()">Analyze Changes</button>
          </div>
        </div>
        <script>
          document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('piNumber').focus();
          });
          function toggleCheckbox(vsId) {
            const checkbox = document.getElementById('vs_' + vsId);
            if (event.target.type !== 'checkbox') {
              checkbox.checked = !checkbox.checked;
            }
            updateSelectAll();
          }
          function toggleAll(element) {
            const selectAll = document.getElementById('selectAll');
            if (element.tagName === 'DIV') {
              selectAll.checked = !selectAll.checked;
            }
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll)');
            checkboxes.forEach(cb => cb.checked = selectAll.checked);
          }
          function updateSelectAll() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll)');
            const selectAll = document.getElementById('selectAll');
            const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
            selectAll.checked = checkedCount === checkboxes.length;
            selectAll.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
          }
          function analyzeChanges() {
            document.querySelectorAll('.error').forEach(el => el.classList.remove('show'));
            const piNumber = document.getElementById('piNumber').value.trim();
            if (!piNumber || isNaN(piNumber)) {
              document.getElementById('piError').classList.add('show');
              return;
            }
            const iterationNumber = document.getElementById('iterationNumber').value;
            if (!iterationNumber) {
              document.getElementById('iterationError').classList.add('show');
              return;
            }
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll):checked');
            if (checkboxes.length === 0) {
              alert('Please select at least one value stream');
              return;
            }
            const selectedValueStreams = Array.from(checkboxes).map(cb => cb.value);
            const allCheckboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll)');
            const isAllSelected = selectedValueStreams.length === allCheckboxes.length;
            document.getElementById('loading').style.display = 'block';
            document.getElementById('analyzeBtn').disabled = true;
            google.script.run
              .withSuccessHandler(() => {
                document.getElementById('loading').style.display = 'none';
                google.script.host.close();
              })
              .withFailureHandler(err => {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('analyzeBtn').disabled = false;
                alert('Error: ' + err);
              })
              .analyzeIterationChanges(
                parseInt(piNumber), 
                parseInt(iterationNumber), 
                isAllSelected ? null : selectedValueStreams
              );
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Generates HTML for the Iteration Report dialog
 */
function getIterationReportDialogHTML() {
  const valueStreamCheckboxes = VALUE_STREAMS.map(vs => {
    const vsId = vs.replace(/[^a-zA-Z0-9]/g, '');
    return `
      <div class="checkbox-item" onclick="toggleCheckbox('${vsId}')">
        <input type="checkbox" id="vs_${vsId}" value="${vs}" onchange="updateSelectAll()">
        <label for="vs_${vsId}">${vs}</label>
      </div>
    `;
  }).join('');

  const portfolios = getAvailablePortfolios();
  const portfolioOptions = portfolios.map(p => 
    `<option value="${p}">${p}</option>`
  ).join('');

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Google Sans', 'Roboto', Arial, sans-serif; padding: 20px; background: #fff; }
          .container { max-width: 480px; margin: 0 auto; }
          h2 { color: #1a73e8; margin-bottom: 8px; font-size: 20px; font-weight: 500; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 20px; }
          .section { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
          .section-title { color: #3c4043; font-weight: 500; margin-bottom: 12px; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }
          .input-group { margin-bottom: 12px; }
          .input-group:last-child { margin-bottom: 0; }
          .pi-input, .iteration-select, .portfolio-select { 
            width: 100%; padding: 12px; font-size: 16px; 
            border: 2px solid #dadce0; border-radius: 8px; 
            transition: border-color 0.2s; background: white; 
          }
          .pi-input:focus, .iteration-select:focus, .portfolio-select:focus { outline: none; border-color: #1a73e8; }
          .checkbox-group { background: white; border-radius: 8px; padding: 12px; max-height: 180px; overflow-y: auto; }
          .checkbox-item { display: flex; align-items: center; padding: 8px; margin: 2px 0; border-radius: 6px; transition: background 0.2s; cursor: pointer; user-select: none; }
          .checkbox-item:hover { background: #f1f3f4; }
          .checkbox-item.select-all { background: #e8f0fe; margin-bottom: 8px; border-bottom: 1px solid #dadce0; font-weight: 500; }
          .checkbox-item input[type="checkbox"] { width: 18px; height: 18px; margin-right: 12px; cursor: pointer; accent-color: #1a73e8; }
          .checkbox-item label { cursor: pointer; flex: 1; color: #3c4043; font-size: 14px; }
          .deferrals-section { background: #fce4ec; border: 2px solid #e91e63; border-radius: 8px; padding: 16px; margin-bottom: 16px; display: none; }
          .deferrals-section .section-title { color: #c2185b; }
          .deferrals-section .checkbox-item { background: white; }
          .help-text { font-size: 12px; color: #5f6368; margin-top: 6px; line-height: 1.4; }
          .button-container { display: flex; gap: 12px; justify-content: flex-end; margin-top: 20px; }
          button { padding: 10px 24px; font-size: 14px; border-radius: 6px; border: none; cursor: pointer; font-weight: 500; transition: all 0.2s; font-family: inherit; }
          .primary-button { background: #1a73e8; color: white; }
          .primary-button:hover:not(:disabled) { background: #1557b0; }
          .primary-button:disabled { background: #94c1f5; cursor: not-allowed; }
          .secondary-button { background: #ffffff; color: #5f6368; border: 1px solid #dadce0; }
          .secondary-button:hover { background: #f8f9fa; }
          .error { color: #d93025; font-size: 12px; margin-top: 8px; display: none; }
          .error.show { display: block; }
          .loading { display: none; text-align: center; padding: 30px; }
          .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1a73e8; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
          #status-text { margin-top: 16px; color: #5f6368; font-size: 14px; }
          .progress-bar-container { width: 100%; height: 8px; background: #e0e0e0; border-radius: 4px; margin-top: 16px; overflow: hidden; }
          .progress-bar { height: 100%; background: #1a73e8; border-radius: 4px; width: 0%; transition: width 0.3s ease; }
          #progress-detail { font-size: 12px; color: #757575; margin-top: 8px; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Generate Iteration Report</h2>
          <div class="subtitle">Select the PI, Iteration, and Value Streams to report on.</div>
          
          <div id="form-container">
            <div class="section">
              <div class="section-title">Program Increment</div>
              <div class="input-group">
                <input type="text" id="piNumber" class="pi-input" placeholder="Enter PI number (e.g., 13)" autocomplete="off" onkeypress="return event.charCode >= 48 && event.charCode <= 57" oninput="updateDeferralsVisibility()" />
                <div id="piError" class="error">Please enter a valid PI number</div>
              </div>
            </div>
            
            <div class="section">
              <div class="section-title">Iteration</div>
              <div class="input-group">
                <select id="iterationNumber" class="iteration-select" onchange="updateDeferralsVisibility()">
                  <option value="">Select iteration...</option>
                  <option value="0">Iteration 0 (Pre-Planning) 📋</option>
                  <option value="1">Iteration 1</option>
                  <option value="2">Iteration 2</option>
                  <option value="3">Iteration 3</option>
                  <option value="4">Iteration 4</option>
                  <option value="5">Iteration 5</option>
                  <option value="6">Iteration 6 (IP) 🔴</option>
                </select>
                <div id="iterationError" class="error">Please select an iteration</div>
              </div>
            </div>
            
            <div class="section">
              <div class="section-title">Portfolio Filter (Optional)</div>
              <div class="input-group">
                <select id="portfolioFilter" class="portfolio-select">
                  <option value="">All Portfolios</option>
                  ${portfolioOptions}
                </select>
                <p class="help-text">Filter report to a specific portfolio. Leave as "All Portfolios" for a complete report.</p>
              </div>
            </div>
            
            <div class="section">
              <div class="section-title">Value Streams</div>
              <div class="checkbox-group">
                <div class="checkbox-item select-all" onclick="toggleAll(this)">
                  <input type="checkbox" id="selectAll" onchange="toggleAll(this)">
                  <label for="selectAll">Select All Value Streams</label>
                </div>
                ${valueStreamCheckboxes}
              </div>
            </div>
            
            <div class="deferrals-section" id="deferralsSection">
              <div class="section-title">🔴 Deferrals Report</div>
              <div class="checkbox-item" onclick="toggleDeferralsCheckbox()">
                <input type="checkbox" id="generateDeferrals" checked onclick="event.stopPropagation();">
                <label for="generateDeferrals">Auto-generate PI Deferrals presentation</label>
              </div>
              <p class="help-text">When checked, a Deferrals presentation will be created automatically after the Iteration 6 report completes.</p>
            </div>
            
            <div class="button-container">
              <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
              <button id="generateBtn" class="primary-button" onclick="runReport()">Generate Report</button>
            </div>
          </div>
          
          <div id="loading" class="loading">
            <div class="spinner"></div>
            <p id="status-text">Initializing...</p>
            <div class="progress-bar-container">
              <div class="progress-bar" id="progressBar"></div>
            </div>
            <p id="progress-detail"></p>
          </div>
        </div>
        
        <script>
          let pollingInterval;
          let totalStreams = 0;

          function toggleCheckbox(vsId) {
            const checkbox = document.getElementById('vs_' + vsId);
            if (event.target.type !== 'checkbox') {
              checkbox.checked = !checkbox.checked;
            }
            updateSelectAll();
          }

          function toggleAll(element) {
            const selectAllCheckbox = document.getElementById('selectAll');
            if (element.tagName === 'DIV') {
              selectAllCheckbox.checked = !selectAllCheckbox.checked;
            }
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll):not(#generateDeferrals)');
            checkboxes.forEach(cb => cb.checked = selectAllCheckbox.checked);
          }
          
          function updateSelectAll() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll):not(#generateDeferrals)');
            const selectAllCheckbox = document.getElementById('selectAll');
            const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
            selectAllCheckbox.checked = checkedCount === checkboxes.length;
            selectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
          }

          function toggleDeferralsCheckbox() {
            const checkbox = document.getElementById('generateDeferrals');
            if (event.target.type !== 'checkbox') {
              checkbox.checked = !checkbox.checked;
            }
          }

          function updateDeferralsVisibility() {
            const iterSelect = document.getElementById('iterationNumber');
            const deferralsSection = document.getElementById('deferralsSection');
            if (iterSelect && deferralsSection) {
              deferralsSection.style.display = (iterSelect.value === '6') ? 'block' : 'none';
            }
          }

          function runReport() {
            document.querySelectorAll('.error').forEach(el => el.classList.remove('show'));
            
            const pi = document.getElementById('piNumber').value.trim();
            const iter = document.getElementById('iterationNumber').value;
            const portfolio = document.getElementById('portfolioFilter').value;
            
            if (!pi || isNaN(pi)) { 
              document.getElementById('piError').classList.add('show'); 
              return; 
            }
            if (!iter) { 
              document.getElementById('iterationError').classList.add('show'); 
              return; 
            }
            
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll):not(#generateDeferrals):checked');
            const allStreamsCheckbox = document.getElementById('selectAll');
            
            let selectedStreams = null;
            if (!allStreamsCheckbox.checked) {
              if (checkboxes.length === 0) { 
                alert('Please select at least one value stream, or check "Select All".'); 
                return; 
              }
              selectedStreams = Array.from(checkboxes).map(cb => cb.value);
            }
            
            const generateDeferrals = iter === '6' && document.getElementById('generateDeferrals').checked;
            totalStreams = selectedStreams ? selectedStreams.length : ${VALUE_STREAMS.length};
            
            document.getElementById('form-container').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('progressBar').style.width = '0%';
            document.getElementById('progress-detail').textContent = 'Preparing to process ' + totalStreams + ' value streams...';
            
            pollingInterval = setInterval(checkStatus, 1500);

            google.script.run
              .withSuccessHandler(onReportComplete)
              .withFailureHandler(onReportError)
              .generateIterationReport(parseInt(pi), parseInt(iter), selectedStreams, true, generateDeferrals, portfolio || null);
          }

          function checkStatus() {
            google.script.run.withSuccessHandler(updateStatusText).getReportGenerationStatus();
          }

          function updateStatusText(status) {
            if (status) {
              const statusEl = document.getElementById('status-text');
              const progressBar = document.getElementById('progressBar');
              const progressDetail = document.getElementById('progress-detail');
              
              if (statusEl.textContent !== status) {
                statusEl.textContent = status;
                const match = status.match(/\\((\\d+)\\/(\\d+)\\)/);
                if (match) {
                  const current = parseInt(match[1]);
                  const total = parseInt(match[2]);
                  const percent = Math.round((current / total) * 100);
                  progressBar.style.width = percent + '%';
                  progressDetail.textContent = current + ' of ' + total + ' value streams processed';
                }
              }
              
              if (status.includes('Complete') || status.includes('Error')) {
                onReportComplete();
              }
            }
          }

          function onReportComplete() {
            clearInterval(pollingInterval);
            const progressBar = document.getElementById('progressBar');
            progressBar.style.width = '100%';
            document.getElementById('status-text').textContent = 'Complete!';
            document.getElementById('progress-detail').textContent = 'Report generated successfully';
            setTimeout(() => google.script.host.close(), 1500);
          }

          function onReportError(err) {
            clearInterval(pollingInterval);
            document.getElementById('status-text').textContent = 'An error occurred.';
            document.getElementById('progress-detail').textContent = '';
            alert('An error occurred: ' + err);
            google.script.host.close();
          }

          document.addEventListener('DOMContentLoaded', function() {
            updateDeferralsVisibility();
          });
        </script>
      </body>
    </html>
  `;
}

/**
 * Generates HTML for the Full Pipeline dialog (sequential)
 */
function getFullPipelineDialogHTML() {
  const valueStreamCheckboxes = VALUE_STREAMS.map(vs => {
    const vsId = vs.replace(/[^a-zA-Z0-9]/g, '');
    return `
      <div class="checkbox-item" onclick="toggleCheckbox('${vsId}')">
        <input type="checkbox" id="vs_${vsId}" value="${vs}" onchange="updateSelectAll()">
        <label for="vs_${vsId}">${vs}</label>
      </div>
    `;
  }).join('');

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Google Sans', 'Roboto', Arial, sans-serif; padding: 20px; background: #fff; }
          .container { max-width: 480px; margin: 0 auto; }
          h2 { color: #1a73e8; margin-bottom: 8px; font-size: 20px; font-weight: 500; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 20px; }
          .section { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
          .section-title { color: #3c4043; font-weight: 500; margin-bottom: 12px; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }
          .input-group { margin-bottom: 12px; }
          .input-group:last-child { margin-bottom: 0; }
          .pi-input, .iteration-select { width: 100%; padding: 12px; font-size: 16px; border: 2px solid #dadce0; border-radius: 8px; transition: border-color 0.2s; background: white; }
          .pi-input:focus, .iteration-select:focus { outline: none; border-color: #1a73e8; }
          .checkbox-group { background: white; border-radius: 8px; padding: 12px; max-height: 160px; overflow-y: auto; }
          .checkbox-item { display: flex; align-items: center; padding: 8px; margin: 2px 0; border-radius: 6px; transition: background 0.2s; cursor: pointer; user-select: none; }
          .checkbox-item:hover { background: #f1f3f4; }
          .checkbox-item.select-all { background: #e8f0fe; margin-bottom: 8px; border-bottom: 1px solid #dadce0; font-weight: 500; }
          .checkbox-item input[type="checkbox"] { width: 18px; height: 18px; margin-right: 12px; cursor: pointer; accent-color: #1a73e8; }
          .checkbox-item label { cursor: pointer; flex: 1; color: #3c4043; font-size: 14px; }
          .pipeline-info { background: #e8f5e9; border: 2px solid #4caf50; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
          .pipeline-info .section-title { color: #2e7d32; margin-bottom: 8px; }
          .pipeline-steps { font-size: 13px; color: #1b5e20; }
          .pipeline-steps div { padding: 4px 0; display: flex; align-items: center; }
          .pipeline-steps .step-num { background: #4caf50; color: white; border-radius: 50%; width: 20px; height: 20px; display: inline-flex; align-items: center; justify-content: center; font-size: 11px; font-weight: bold; margin-right: 8px; }
          .button-container { display: flex; gap: 12px; justify-content: flex-end; margin-top: 20px; }
          button { padding: 10px 24px; font-size: 14px; border-radius: 6px; border: none; cursor: pointer; font-weight: 500; transition: all 0.2s; font-family: inherit; }
          .primary-button { background: #4caf50; color: white; }
          .primary-button:hover:not(:disabled) { background: #388e3c; }
          .primary-button:disabled { background: #a5d6a7; cursor: not-allowed; }
          .secondary-button { background: #ffffff; color: #5f6368; border: 1px solid #dadce0; }
          .secondary-button:hover { background: #f8f9fa; }
          .error { color: #d93025; font-size: 12px; margin-top: 8px; display: none; }
          .error.show { display: block; }
          .loading { display: none; text-align: center; padding: 30px; }
          .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #4caf50; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
          #status-text { margin-top: 16px; color: #5f6368; font-size: 14px; font-weight: 500; }
          #step-indicator { margin-top: 12px; padding: 12px; background: #f5f5f5; border-radius: 8px; }
          .step-status { display: flex; align-items: center; padding: 6px 0; font-size: 13px; }
          .step-status .icon { width: 20px; margin-right: 8px; text-align: center; }
          .step-status.pending { color: #9e9e9e; }
          .step-status.active { color: #1a73e8; font-weight: 500; }
          .step-status.complete { color: #4caf50; }
          .step-status.error { color: #d32f2f; }
          #progress-detail { font-size: 12px; color: #757575; margin-top: 8px; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>🚀 Run Full Pipeline</h2>
          <div class="subtitle">Generate complete iteration report with AI-powered perspectives</div>
          
          <div id="form-container">
            <div class="pipeline-info">
              <div class="section-title">This will run all 4 steps automatically:</div>
              <div class="pipeline-steps">
                <div><span class="step-num">1</span> Generate Iteration Report (JIRA data)</div>
                <div><span class="step-num">2</span> Enrich with Initiative Data</div>
                <div><span class="step-num">3</span> Generate Epic Perspectives (AI)</div>
                <div><span class="step-num">4</span> Generate Initiative Perspectives (AI)</div>
              </div>
            </div>
            
            <div class="section">
              <div class="section-title">Program Increment</div>
              <div class="input-group">
                <input type="text" id="piNumber" class="pi-input" placeholder="Enter PI number (e.g., 13)" autocomplete="off" onkeypress="return event.charCode >= 48 && event.charCode <= 57" />
                <div id="piError" class="error">Please enter a valid PI number</div>
              </div>
            </div>
            
            <div class="section">
              <div class="section-title">Iteration</div>
              <div class="input-group">
                <select id="iterationNumber" class="iteration-select">
                  <option value="">Select iteration...</option>
                  <option value="0">Iteration 0 (Pre-Planning) 📋</option>
                  <option value="1">Iteration 1</option>
                  <option value="2">Iteration 2</option>
                  <option value="3">Iteration 3</option>
                  <option value="4">Iteration 4</option>
                  <option value="5">Iteration 5</option>
                  <option value="6">Iteration 6 (IP)</option>
                </select>
                <div id="iterationError" class="error">Please select an iteration</div>
              </div>
            </div>
            
            <div class="section">
              <div class="section-title">Value Streams</div>
              <div class="checkbox-group">
                <div class="checkbox-item select-all" onclick="toggleAll(this)">
                  <input type="checkbox" id="selectAll" onchange="toggleAll(this)" checked>
                  <label for="selectAll">Select All Value Streams</label>
                </div>
                ${valueStreamCheckboxes}
              </div>
            </div>
            
            <div class="button-container">
              <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
              <button id="runBtn" class="primary-button" onclick="runPipeline()">🚀 Run Pipeline</button>
            </div>
          </div>
          
          <div id="loading" class="loading">
            <div class="spinner"></div>
            <p id="status-text">Initializing pipeline...</p>
            <div id="step-indicator">
              <div class="step-status pending" id="step1-status"><span class="icon">○</span> Step 1: Generate Iteration Report</div>
              <div class="step-status pending" id="step2-status"><span class="icon">○</span> Step 2: Enrich with Initiative Data</div>
              <div class="step-status pending" id="step3-status"><span class="icon">○</span> Step 3: Generate Epic Perspectives</div>
              <div class="step-status pending" id="step4-status"><span class="icon">○</span> Step 4: Generate Initiative Perspectives</div>
            </div>
            <p id="progress-detail"></p>
          </div>
        </div>
        
        <script>
          let pollingInterval;

          document.addEventListener('DOMContentLoaded', function() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll)');
            checkboxes.forEach(cb => cb.checked = true);
          });

          function toggleCheckbox(vsId) {
            const checkbox = document.getElementById('vs_' + vsId);
            if (event.target.type !== 'checkbox') {
              checkbox.checked = !checkbox.checked;
            }
            updateSelectAll();
          }

          function toggleAll(element) {
            const selectAllCheckbox = document.getElementById('selectAll');
            if (element.tagName === 'DIV') {
              selectAllCheckbox.checked = !selectAllCheckbox.checked;
            }
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll)');
            checkboxes.forEach(cb => cb.checked = selectAllCheckbox.checked);
          }
          
          function updateSelectAll() {
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll)');
            const selectAllCheckbox = document.getElementById('selectAll');
            const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
            selectAllCheckbox.checked = checkedCount === checkboxes.length;
            selectAllCheckbox.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
          }

          function runPipeline() {
            document.querySelectorAll('.error').forEach(el => el.classList.remove('show'));
            
            const pi = document.getElementById('piNumber').value.trim();
            const iter = document.getElementById('iterationNumber').value;
            
            if (!pi || isNaN(pi)) { 
              document.getElementById('piError').classList.add('show'); 
              return; 
            }
            if (!iter) { 
              document.getElementById('iterationError').classList.add('show'); 
              return; 
            }
            
            const checkboxes = document.querySelectorAll('input[type="checkbox"]:not(#selectAll):checked');
            const allStreamsCheckbox = document.getElementById('selectAll');
            
            let selectedStreams = null;
            if (!allStreamsCheckbox.checked) {
              if (checkboxes.length === 0) { 
                alert('Please select at least one value stream, or check "Select All".'); 
                return; 
              }
              selectedStreams = Array.from(checkboxes).map(cb => cb.value);
            }
            
            document.getElementById('form-container').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            
            pollingInterval = setInterval(checkPipelineStatus, 2000);

            google.script.run
              .withSuccessHandler(onPipelineComplete)
              .withFailureHandler(onPipelineError)
              .runFullPipeline(parseInt(pi), parseInt(iter), selectedStreams);
          }

          function checkPipelineStatus() {
            google.script.run
              .withSuccessHandler(updateStatusDisplay)
              .withFailureHandler(function() {})
              .getFullPipelineStatus();
          }

          function updateStatusDisplay(status) {
            if (!status) return;
            
            document.getElementById('status-text').textContent = status.message || 'Processing...';
            document.getElementById('progress-detail').textContent = status.detail || '';
            
            for (let i = 1; i <= 4; i++) {
              const stepEl = document.getElementById('step' + i + '-status');
              const stepStatus = status['step' + i] || 'pending';
              
              stepEl.className = 'step-status ' + stepStatus;
              
              if (stepStatus === 'complete') {
                stepEl.querySelector('.icon').textContent = '✓';
              } else if (stepStatus === 'active') {
                stepEl.querySelector('.icon').textContent = '●';
              } else if (stepStatus === 'error') {
                stepEl.querySelector('.icon').textContent = '✗';
              } else {
                stepEl.querySelector('.icon').textContent = '○';
              }
            }
          }

          function onPipelineComplete(result) {
            clearInterval(pollingInterval);
            
            for (let i = 1; i <= 4; i++) {
              const stepEl = document.getElementById('step' + i + '-status');
              stepEl.className = 'step-status complete';
              stepEl.querySelector('.icon').textContent = '✓';
            }
            
            document.getElementById('status-text').textContent = '✅ Pipeline Complete!';
            document.getElementById('progress-detail').textContent = 'Total time: ' + (result.duration || 'N/A') + ' seconds';
            
            setTimeout(function() { google.script.host.close(); }, 3000);
          }

          function onPipelineError(error) {
            clearInterval(pollingInterval);
            document.getElementById('status-text').textContent = '❌ Pipeline Error';
            document.getElementById('progress-detail').textContent = error.toString();
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Generates HTML for refresh child metrics dialog
 */
function getRefreshChildMetricsHTML() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Google Sans', 'Roboto', Arial, sans-serif; padding: 20px; background: #fff; }
          .container { max-width: 450px; margin: 0 auto; }
          h2 { color: #1a73e8; margin-bottom: 8px; font-size: 20px; font-weight: 500; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 24px; }
          .section { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 20px; }
          .section-title { color: #3c4043; font-weight: 500; margin-bottom: 12px; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; }
          .input-group { margin-bottom: 16px; }
          .pi-input, .sheet-select { width: 100%; padding: 12px; font-size: 16px; border: 2px solid #dadce0; border-radius: 8px; transition: border-color 0.2s; background: white; }
          .pi-input:focus, .sheet-select:focus { outline: none; border-color: #1a73e8; }
          .info-box { background: #e3f2fd; border-left: 3px solid #1976d2; padding: 12px; margin: 16px 0; border-radius: 4px; }
          .info-box p { color: #1565c0; font-size: 13px; margin: 4px 0; }
          .info-box strong { color: #0d47a1; }
          .button-container { display: flex; gap: 12px; justify-content: flex-end; margin-top: 24px; }
          button { padding: 10px 24px; font-size: 14px; border-radius: 6px; border: none; cursor: pointer; font-weight: 500; transition: all 0.2s; font-family: inherit; }
          .primary-button { background: #1a73e8; color: white; }
          .primary-button:hover:not(:disabled) { background: #1557b0; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
          .primary-button:disabled { background: #94c1f5; cursor: not-allowed; }
          .secondary-button { background: #ffffff; color: #5f6368; border: 1px solid #dadce0; }
          .secondary-button:hover { background: #f8f9fa; }
          .diagnostic-button { background: #ea4335; color: white; margin-right: auto; }
          .diagnostic-button:hover { background: #d33426; }
          .error { color: #d93025; font-size: 12px; margin-top: 8px; display: none; }
          .error.show { display: block; }
          .loading { display: none; text-align: center; padding: 20px; }
          .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1a73e8; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        </style>
      </head>
      <body>
        <div class="container" id="mainContent">
          <h2>Refresh Child Ticket Metrics</h2>
          <div class="subtitle">Update story point and completion data from JIRA</div>
          
          <div class="section">
            <div class="section-title">Program Increment</div>
            <div class="input-group">
              <input type="text" id="piNumber" class="pi-input" placeholder="Enter PI number (e.g., 11, 12, 13)" autocomplete="off" onkeypress="return event.charCode >= 48 && event.charCode <= 57" />
              <div id="piError" class="error">Please enter a valid PI number</div>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Target Sheet</div>
            <div class="input-group">
              <select id="sheetName" class="sheet-select">
                <option value="">Select iteration sheet...</option>
                <option value="Iteration 0 (Pre-Planning)">Iteration 0 (Pre-Planning)</option>
                <option value="Iteration 1">Iteration 1</option>
                <option value="Iteration 2">Iteration 2</option>
                <option value="Iteration 3">Iteration 3</option>
                <option value="Iteration 4">Iteration 4</option>
                <option value="Iteration 5">Iteration 5</option>
                <option value="Iteration 6">Iteration 6 (IP)</option>
              </select>
              <div id="sheetError" class="error">Please select a sheet</div>
            </div>
          </div>
          
          <div class="info-box">
            <p><strong>This will update columns T-W:</strong></p>
            <p>• T: Total Story Points (all Stories/Bugs)</p>
            <p>• U: Closed Story Points (Closed Stories/Bugs)</p>
            <p>• V: Total Child Count</p>
            <p>• W: Closed Child Count</p>
            <p>• X: % Complete (auto-calculated)</p>
          </div>
          
          <div class="button-container">
            <button class="diagnostic-button" onclick="runDiagnostic()">🔍 Diagnose Fields</button>
            <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
            <button id="refreshBtn" class="primary-button" onclick="refreshMetrics()">Refresh Metrics</button>
          </div>
        </div>
        
        <div id="loading" class="loading">
          <div class="spinner"></div>
          <p style="margin-top: 16px; color: #5f6368;">Fetching child tickets from JIRA...<br>This may take a few minutes.</p>
        </div>
        
        <script>
          document.addEventListener('DOMContentLoaded', function() {
            document.getElementById('piNumber').focus();
          });
          
          function refreshMetrics() {
            const piNumber = document.getElementById('piNumber').value.trim();
            const sheetName = document.getElementById('sheetName').value;
            
            let isValid = true;
            
            if (!piNumber || parseInt(piNumber) < 1) {
              document.getElementById('piError').classList.add('show');
              isValid = false;
            } else {
              document.getElementById('piError').classList.remove('show');
            }
            
            if (!sheetName) {
              document.getElementById('sheetError').classList.add('show');
              isValid = false;
            } else {
              document.getElementById('sheetError').classList.remove('show');
            }
            
            if (!isValid) return;
            
            document.getElementById('mainContent').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .refreshChildTicketMetrics(piNumber, sheetName);
          }
          
          function runDiagnostic() {
            document.getElementById('mainContent').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            document.getElementById('loading').querySelector('p').innerHTML = 'Running field diagnostics...<br>Checking your JIRA configuration.';
            
            google.script.run
              .withSuccessHandler(onDiagnosticSuccess)
              .withFailureHandler(onFailure)
              .diagnoseStoryPointFields();
          }
          
          function onSuccess(result) {
            alert('✅ Success!\\n\\n' + result);
            google.script.host.close();
          }
          
          function onDiagnosticSuccess(result) {
            alert('🔍 Field Diagnostic Results\\n\\n' + result);
            document.getElementById('loading').style.display = 'none';
            document.getElementById('mainContent').style.display = 'block';
          }
          
          function onFailure(error) {
            alert('❌ Error:\\n\\n' + error.message);
            document.getElementById('loading').style.display = 'none';
            document.getElementById('mainContent').style.display = 'block';
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Generates HTML for predictability validation dialog
 */
function getValidationDialogHTML(availablePIs, valueStreams) {
  const piOptions = availablePIs.map(pi => 
    `<label class="radio-option">
      <input type="radio" name="pi" value="${pi}" ${pi === availablePIs[0] ? 'checked' : ''}>
      <span>PI ${pi}</span>
    </label>`
  ).join('');
  
  const vsOptions = valueStreams.map((vs, index) => 
    `<label class="radio-option">
      <input type="radio" name="valueStream" value="${vs}" ${index === 0 ? 'checked' : ''}>
      <span>${vs}</span>
    </label>`
  ).join('');
  
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          * { box-sizing: border-box; margin: 0; padding: 0; }
          body { font-family: 'Google Sans', 'Roboto', Arial, sans-serif; padding: 20px; background: #fff; }
          .container { max-width: 460px; margin: 0 auto; }
          h2 { color: #1a73e8; margin-bottom: 8px; font-size: 20px; font-weight: 500; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 24px; }
          .section { background: #f8f9fa; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
          .section-title { font-weight: 500; color: #202124; margin-bottom: 12px; font-size: 14px; }
          .radio-option { display: block; padding: 10px 12px; margin-bottom: 8px; background: white; border: 1px solid #dadce0; border-radius: 6px; cursor: pointer; transition: all 0.2s; }
          .radio-option:hover { border-color: #1a73e8; background: #f8f9fa; }
          .radio-option input[type="radio"] { margin-right: 10px; cursor: pointer; }
          .radio-option input[type="radio"]:checked + span { color: #1a73e8; font-weight: 500; }
          .radio-option span { color: #202124; font-size: 14px; }
          .button-container { display: flex; justify-content: flex-end; gap: 12px; margin-top: 24px; }
          button { padding: 10px 24px; border: none; border-radius: 6px; font-size: 14px; font-weight: 500; cursor: pointer; transition: all 0.2s; }
          .btn-cancel { background: white; color: #5f6368; border: 1px solid #dadce0; }
          .btn-cancel:hover { background: #f8f9fa; }
          .btn-generate { background: #1a73e8; color: white; }
          .btn-generate:hover { background: #1557b0; }
          .btn-generate:disabled { background: #dadce0; cursor: not-allowed; }
          #loading { display: none; text-align: center; padding: 20px; }
          .spinner { border: 3px solid #f3f3f3; border-top: 3px solid #1a73e8; border-radius: 50%; width: 40px; height: 40px; animation: spin 1s linear infinite; margin: 0 auto 16px; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
          .loading-text { color: #5f6368; font-size: 14px; }
          .scroll-container { max-height: 300px; overflow-y: auto; padding-right: 8px; }
          .scroll-container::-webkit-scrollbar { width: 8px; }
          .scroll-container::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 4px; }
          .scroll-container::-webkit-scrollbar-thumb { background: #dadce0; border-radius: 4px; }
          .scroll-container::-webkit-scrollbar-thumb:hover { background: #bdc1c6; }
        </style>
      </head>
      <body>
        <div class="container">
          <div id="mainContent">
            <h2>Predictability Validation Report</h2>
            <p class="subtitle">Generate a detailed validation report comparing data sources</p>
            
            <div class="section">
              <div class="section-title">Select Program Increment</div>
              ${piOptions}
            </div>
            
            <div class="section">
              <div class="section-title">Select Value Stream</div>
              <div class="scroll-container">
                ${vsOptions}
              </div>
            </div>
            
            <div class="button-container">
              <button type="button" class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
              <button type="button" class="btn-generate" onclick="generateReport()">Generate Report</button>
            </div>
          </div>
          
          <div id="loading">
            <div class="spinner"></div>
            <div class="loading-text">Generating validation report...</div>
          </div>
        </div>
        
        <script>
          function generateReport() {
            const pi = document.querySelector('input[name="pi"]:checked').value;
            const valueStream = document.querySelector('input[name="valueStream"]:checked').value;
            
            document.getElementById('mainContent').style.display = 'none';
            document.getElementById('loading').style.display = 'block';
            
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .generatePredictabilityValidationFromDialog(parseInt(pi), valueStream);
          }
          
          function onSuccess() {
            google.script.host.close();
          }
          
          function onFailure(error) {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('mainContent').style.display = 'block';
            alert('Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `;
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 4: REPORT GENERATION FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generates an iteration report for the specified PI and iteration
 */
function generateIterationReport(piNumber, iterationNumber, selectedValueStreams, includeStoryPoints = true, generateDeferrals = false, portfolioFilter = null) {
  return logActivity('PI Iteration Report Generation', () => {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const startTime = Date.now();
    const userCache = CacheService.getUserCache();

    const streamsToProcess = (selectedValueStreams && selectedValueStreams.length > 0)
      ? selectedValueStreams
      : VALUE_STREAMS;
    
    const totalStreams = streamsToProcess.length;
    
    console.log('\n╔══════════════════════════════════════════════════════════════╗');
    console.log('║           ITERATION REPORT GENERATION                        ║');
    console.log('╠══════════════════════════════════════════════════════════════╣');
    console.log(`║  PI: ${piNumber}  |  Iteration: ${iterationNumber}  |  ${new Date().toLocaleString()}`);
    console.log(`║  Value Streams: ${totalStreams}  |  Story Points: ${includeStoryPoints}`);
    console.log(`║  Portfolio Filter: ${portfolioFilter || 'ALL'}`);
    if (iterationNumber == 6 && generateDeferrals) {
      console.log(`║  Auto-Generate Deferrals: Yes`);
    }
    console.log('╚══════════════════════════════════════════════════════════════╝\n');
    
    try {
      userCache.put('report_status', 'Initializing...', 120);

      // Record refresh timestamp BEFORE data fetching
      const refreshTimestamp = new Date();
      const reportKey = `PI${piNumber}_Iteration${iterationNumber}`;
      recordRefreshTimestamp(reportKey, refreshTimestamp);
      console.log(`✓ Recorded refresh timestamp: ${refreshTimestamp.toLocaleString()}`);

      for (let i = 0; i < totalStreams; i++) {
        const valueStream = streamsToProcess[i];
        const progressMsg = `(${i + 1}/${totalStreams}) Processing: ${valueStream}`;
        
        userCache.put('report_status', progressMsg, 120);
        ss.toast(progressMsg, '⏳ Generating Report', 30);
        
        const clearExisting = (i === 0);
        const isFirstValueStream = (i === 0);
        processAndWriteDataForIteration(piNumber, iterationNumber, valueStream, includeStoryPoints, clearExisting, isFirstValueStream);
      }
      
      userCache.put('report_status', 'Complete!', 120);
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      ss.toast('', '', 1);
      Utilities.sleep(300);
      ss.toast(`Report complete in ${duration}s`, '✅ Success', 4);
      
      // Auto-generate deferrals for Iteration 6 if requested
      if (iterationNumber == 6 && generateDeferrals) {
        console.log(`\n${'─'.repeat(60)}`);
        console.log(`AUTO-GENERATING DEFERRALS FOR PI ${piNumber}`);
        console.log(`${'─'.repeat(60)}\n`);
        
        ss.toast('', '', 1);
        Utilities.sleep(500);
        
        generateDeferralsSlideFromDialog(piNumber, null, 'program_initiative');
      } else {
        ui.alert('✅ Success', `Report generation complete in ${duration} seconds.`, ui.ButtonSet.OK);
      }

    } catch (error) {
      console.error('Report Generation Error:', error);
      userCache.put('report_status', 'Error!', 120);
      ss.toast('', '', 1);
      ui.alert('Error', `An error occurred: ${error.toString()}`, ui.ButtonSet.OK);
    }
  }, { piNumber, iterationNumber, streamCount: selectedValueStreams?.length, includeStoryPoints, generateDeferrals, portfolioFilter });
}

/**
 * Processes and writes data for a single value stream iteration
 */
function processAndWriteDataForIteration(piNumber, iterationNumber, valueStream, includeStoryPoints = true, clearExisting = true, isFirstValueStream = false) {
  const programIncrement = `PI ${piNumber}`;
  const sheetName = `${programIncrement} - Iteration ${iterationNumber}`;
  
  console.log(`Processing ${valueStream} for Iteration ${iterationNumber}`);
  
  const includePrePlanning = (iterationNumber == 0);
  
  const epics = fetchEpicsForPIOptimized(programIncrement, [valueStream], includePrePlanning);
  if (epics.length === 0) {
    console.log(`No epics found for ${valueStream} in ${programIncrement}. Clearing its data.`);
    writeGovernanceDataToSheet([], programIncrement, valueStream, sheetName, clearExisting, isFirstValueStream);
    return;
  }
  
  console.log(`Found ${epics.length} epics`);
  
  const epicDataMap = fetchChildDataForEpicsOptimized(epics, includeStoryPoints);
  const { relevantEpics, filteredEpicDataMap } = filterDataForIteration(epics, epicDataMap, iterationNumber);
  const processedData = processGovernanceData(relevantEpics, filteredEpicDataMap);
  writeGovernanceDataToSheet(processedData, programIncrement, valueStream, sheetName, clearExisting, isFirstValueStream);

  console.log(`Finished processing ${valueStream} for Iteration ${iterationNumber}. Wrote ${processedData.length} rows.`);
}

/**
 * Filters data for a specific iteration (currently passes through all data)
 */
function filterDataForIteration(epics, epicDataMap, iterationNumber) {
  console.log(`Including ALL ${epics.length} epics with ALL dependencies`);
  
  const relevantEpics = [];
  const filteredEpicDataMap = {};
  
  epics.forEach(epic => {
    const originalData = epicDataMap[epic.key];
    if (!originalData) return;
    
    relevantEpics.push(epic);
    
    filteredEpicDataMap[epic.key] = {
      totalStoryPoints: originalData.totalStoryPoints,
      closedStoryPoints: originalData.closedStoryPoints,
      totalChildren: originalData.totalChildren,
      closedChildren: originalData.closedChildren,
      risks: originalData.risks || [],
      dependencies: originalData.dependencies || [],
      allChildren: originalData.allChildren || []
    };
  });
  
  return { relevantEpics, filteredEpicDataMap };
}

/**
 * Generates governance slides
 */
function generateGovernanceSlides(piNumber, iterationNumber, presentationName, showDependencies = true, noInitiativeMode = 'show', hideSameTeamDeps = true, singlePortfolio = null, showAllEpics = false, includeAtRisk = true, highlightSchedulingRisk = true, showPreviouslyClosed = true, valueStreamFilter = null, groupByFixVersion = true) {

  return logActivity('Governance Presentation Generation', () => {
    try {
      console.log('========== CHECKING CONFIGURATION ==========');
      console.log(`Show Dependencies: ${showDependencies}`);
      console.log(`Hide Same-Team Dependencies: ${hideSameTeamDeps}`);
      console.log(`Single Portfolio: ${singlePortfolio || 'ALL'}`);
      console.log(`Show All Epics: ${showAllEpics}`);
      console.log(`Include At-Risk: ${includeAtRisk}`);
      console.log(`Show Previously Closed: ${showPreviouslyClosed}`);
      console.log(`Value Stream Filter: ${valueStreamFilter ? valueStreamFilter.join(', ') : 'ALL'}`);
      console.log(`Group by Fix Version: ${groupByFixVersion}`);

      
      const slideConfig = readSlideConfiguration();
      
      if (slideConfig) {
        console.log('✓ Configuration tab found - using custom field display and ordering');
      } else {
        console.log('⚠ Configuration tab not found or empty - using default alphabetical ordering');
      }
      
      const data = processSlideData(piNumber, iterationNumber, slideConfig, showAllEpics, includeAtRisk);
      
      // Apply "Exclude Already Closed" filter if enabled
      if (!showPreviouslyClosed && data.epics) {
        const beforeCount = data.epics.length;
        data.epics = data.epics.filter(epic => {
          // Keep if NOT (closed but not closed this iteration)
          const isClosed = epic.isDone || epic.isClosed || 
            (epic['Status'] && epic['Status'].toLowerCase() === 'closed');
          const closedThisIteration = epic.closedThisIteration === true;
          const alreadyClosed = epic.alreadyClosed === true;
          
          // Exclude items that were already closed (closed but NOT closedThisIteration)
          if (isClosed && !closedThisIteration && alreadyClosed) {
            console.log(`  🚫 Excluding already-closed: ${epic['Key']} (Status: ${epic['Status']})`);
            return false;
          }
          return true;
        });
        console.log(`Exclude Already Closed: ${beforeCount} -> ${data.epics.length} epics`);
      }
      
      if (iterationNumber <= 1) {
        console.log(`Running BASELINE report for Iteration ${iterationNumber}...`);
        return createFormattedPresentation(presentationName, data, showDependencies, noInitiativeMode, hideSameTeamDeps, singlePortfolio, valueStreamFilter, groupByFixVersion);
      } else {
        console.log(`Running CHANGES-ONLY report for Iteration ${iterationNumber}...`);
        const changesOnlyName = presentationName.replace('Governance Slides', 'Governance Changes');
        return createChangesOnlyPresentation(changesOnlyName, data, piNumber, iterationNumber, showDependencies, hideSameTeamDeps, highlightSchedulingRisk, singlePortfolio, valueStreamFilter, groupByFixVersion);
      }

    } catch (error) {
      console.error('Slide generation error:', error);
      SpreadsheetApp.getUi().alert('Error', `Failed to generate slides: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
      return { success: false, error: error.message };
    }
  }, { piNumber, iterationNumber, presentationName, showDependencies, hideSameTeamDeps, singlePortfolio, valueStreamFilter, showPreviouslyClosed });
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 5: SEQUENTIAL PIPELINE FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Runs the full pipeline: STEP 1 → STEP 2 → STEP 3 → STEP 4 (Sequential)
 */
function runFullPipeline(piNumber, iterationNumber, selectedValueStreams) {
  return logActivity('Full Pipeline (Steps 1-4)', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userCache = CacheService.getUserCache();
    const startTime = Date.now();
    
    const streamsToProcess = selectedValueStreams || VALUE_STREAMS;
    const sheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
    
    console.log('\n' + '═'.repeat(70));
    console.log('  🚀 FULL PIPELINE EXECUTION');
    console.log('═'.repeat(70));
    console.log(`  PI: ${piNumber}  |  Iteration: ${iterationNumber}`);
    console.log(`  Value Streams: ${streamsToProcess.length}`);
    console.log(`  Started: ${new Date().toLocaleString()}`);
    console.log('═'.repeat(70) + '\n');
    
    try {
      updatePipelineStatus(userCache, {
        message: 'Starting pipeline...',
        step1: 'active', step2: 'pending', step3: 'pending', step4: 'pending'
      });
      
      // STEP 1
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 1: Generate Iteration Report');
      console.log('─'.repeat(60));
      
      updatePipelineStatus(userCache, {
        message: 'Step 1: Generating iteration report...',
        detail: `Processing ${streamsToProcess.length} value streams`,
        step1: 'active', step2: 'pending', step3: 'pending', step4: 'pending'
      });
      
      runStep1_GenerateIterationReport(piNumber, iterationNumber, streamsToProcess, ss, userCache);
      
      updatePipelineStatus(userCache, {
        message: 'Step 1 complete',
        step1: 'complete', step2: 'active', step3: 'pending', step4: 'pending'
      });
      console.log('✓ Step 1 complete\n');
      
      // STEP 2
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 2: Enrich with Initiative Data');
      console.log('─'.repeat(60));
      
      updatePipelineStatus(userCache, {
        message: 'Step 2: Enriching with initiative data...',
        detail: 'Fetching initiative details from JIRA',
        step1: 'complete', step2: 'active', step3: 'pending', step4: 'pending'
      });
      
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) {
        throw new Error(`Sheet "${sheetName}" not found after Step 1`);
      }
      ss.setActiveSheet(targetSheet);
      
      enrichWithInitiativeData();
      
      updatePipelineStatus(userCache, {
        message: 'Step 2 complete',
        step1: 'complete', step2: 'complete', step3: 'active', step4: 'pending'
      });
      console.log('✓ Step 2 complete\n');
      
      // STEP 3
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 3: Generate Epic Perspectives');
      console.log('─'.repeat(60));
      
      updatePipelineStatus(userCache, {
        message: 'Step 3: Generating epic perspectives...',
        detail: 'AI processing business & technical summaries',
        step1: 'complete', step2: 'complete', step3: 'active', step4: 'pending'
      });
      
      ss.setActiveSheet(targetSheet);
      generateEpicPerspectives();
      
      updatePipelineStatus(userCache, {
        message: 'Step 3 complete',
        step1: 'complete', step2: 'complete', step3: 'complete', step4: 'active'
      });
      console.log('✓ Step 3 complete\n');
      
      // STEP 4
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 4: Generate Initiative Perspectives');
      console.log('─'.repeat(60));
      
      updatePipelineStatus(userCache, {
        message: 'Step 4: Generating initiative perspectives...',
        detail: 'AI merging epic perspectives into initiatives',
        step1: 'complete', step2: 'complete', step3: 'complete', step4: 'active'
      });
      
      ss.setActiveSheet(targetSheet);
      generateInitiativePerspectives();
      
      updatePipelineStatus(userCache, {
        message: 'Pipeline complete!',
        step1: 'complete', step2: 'complete', step3: 'complete', step4: 'complete'
      });
      console.log('✓ Step 4 complete\n');
      
      // COMPLETE
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      console.log('\n' + '═'.repeat(70));
      console.log('  ✅ PIPELINE COMPLETE');
      console.log('═'.repeat(70));
      console.log(`  Total Duration: ${duration} seconds`);
      console.log(`  Sheet: ${sheetName}`);
      console.log('═'.repeat(70) + '\n');
      
      ss.toast(`Pipeline complete in ${duration}s`, '✅ Success', 5);
      
      return {
        success: true,
        duration: duration,
        sheetName: sheetName,
        valueStreams: streamsToProcess.length
      };
      
    } catch (error) {
      console.error('❌ Pipeline Error:', error);
      
      updatePipelineStatus(userCache, {
        message: 'Pipeline error: ' + error.message,
        detail: 'Check logs for details'
      });
      
      throw error;
    }
  }, { piNumber, iterationNumber, valueStreamCount: (selectedValueStreams || VALUE_STREAMS).length });
}

/**
 * Helper: Run Step 1 without UI alerts (for pipeline use)
 */
function runStep1_GenerateIterationReport(piNumber, iterationNumber, streamsToProcess, ss, userCache) {
  const programIncrement = `PI ${piNumber}`;
  const sheetName = `${programIncrement} - Iteration ${iterationNumber}`;
  
  const refreshTimestamp = new Date();
  const reportKey = `PI${piNumber}_Iteration${iterationNumber}`;
  recordRefreshTimestamp(reportKey, refreshTimestamp);
  
  for (let i = 0; i < streamsToProcess.length; i++) {
    const valueStream = streamsToProcess[i];
    const progressMsg = `(${i + 1}/${streamsToProcess.length}) ${valueStream}`;
    
    updatePipelineStatus(userCache, {
      message: 'Step 1: Generating iteration report...',
      detail: progressMsg,
      step1: 'active', step2: 'pending', step3: 'pending', step4: 'pending'
    });
    
    ss.toast(progressMsg, '⏳ Step 1', 30);
    
    const clearExisting = (i === 0);
    const isFirstValueStream = (i === 0);
    processAndWriteDataForIteration(piNumber, iterationNumber, valueStream, true, clearExisting, isFirstValueStream);
  }
  
  ss.toast('', '', 1);
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 6: PARALLEL PIPELINE FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════
//
// TODO: Future optimization opportunity
// ────────────────────────────────────────────────────────────────────────────────
// Currently Step 2 (Initiative Enrichment) must run AFTER Step 1 because it reads
// Parent Key from the sheet. Steps 3+4 require Initiative Title/Description from
// Step 2 for AI context.
//
// Potential future refactor:
// - Modify processGovernanceData() to include Initiative Title/Description columns
// - Fetch initiatives in parallel with epics in Step 1
// - Write everything in one pass
// - This would allow Steps 1+2 to merge, potentially saving 10-15 seconds
// ────────────────────────────────────────────────────────────────────────────────

/**
 * Runs the full pipeline (Steps 1-4) with parallel JIRA fetching
 * Uses toast notifications for progress (no polling)
 */
function runFullPipelineParallel(piNumber, iterationNumber, selectedValueStreams, destinationConfig) {
  return logActivity('Full Pipeline PARALLEL (Steps 1-3)', () => {
    const startTime = Date.now();
    const sourceSS = SpreadsheetApp.getActiveSpreadsheet();
    
    // Default destination config for backward compatibility
    destinationConfig = destinationConfig || { type: 'current' };
    
    const streamsToProcess = selectedValueStreams || VALUE_STREAMS;
    const programIncrement = `PI ${piNumber}`;
    const includePrePlanning = (iterationNumber == 0);
    
    // Determine target spreadsheet and sheet name based on destination
    let targetSS;
    let sheetName;
    let newSpreadsheetUrl = null;
    
    if (destinationConfig.type === 'new') {
      const newName = `PI ${piNumber} - Iteration ${iterationNumber} Report - ${new Date().toLocaleDateString()}`;
      const newSpreadsheet = SpreadsheetApp.create(newName);
      targetSS = newSpreadsheet;
      sheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
      newSpreadsheetUrl = newSpreadsheet.getUrl();
      
      try {
        const sourceFile = DriveApp.getFileById(sourceSS.getId());
        const parentFolders = sourceFile.getParents();
        if (parentFolders.hasNext()) {
          const folder = parentFolders.next();
          DriveApp.getFileById(newSpreadsheet.getId()).moveTo(folder);
        }
      } catch (e) {
        console.warn('Could not move new spreadsheet to source folder:', e.message);
      }
      
      console.log(`✓ Created new spreadsheet: ${newName}`);
      console.log(`  URL: ${newSpreadsheetUrl}`);
      
    } else if (destinationConfig.type === 'existing' && destinationConfig.existingSheetName) {
      targetSS = sourceSS;
      sheetName = destinationConfig.existingSheetName;
      console.log(`✓ Will update existing sheet: ${sheetName}`);
      
    } else {
      targetSS = sourceSS;
      sheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
    }
    
    console.log('\n' + '═'.repeat(70));
    console.log('  ⚡ PARALLEL PIPELINE EXECUTION (OPTIMIZED)');
    console.log('═'.repeat(70));
    console.log(`  PI: ${piNumber}  |  Iteration: ${iterationNumber}`);
    console.log(`  Value Streams: ${streamsToProcess.length}`);
    console.log(`  Pre-Planning: ${includePrePlanning}`);
    console.log(`  Destination: ${destinationConfig.type}`);
    console.log(`  Target Sheet: ${sheetName}`);
    console.log(`  Started: ${new Date().toLocaleString()}`);
    console.log('═'.repeat(70) + '\n');
    
    try {
      // ═══════════════════════════════════════════════════════════════════════
      // STEP 1: PARALLEL FETCH → PROCESS ALL → WRITE ONCE
      // ═══════════════════════════════════════════════════════════════════════
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 1: Parallel Fetch, Process & Batched Write');
      console.log('─'.repeat(60));
      
      // 1a. Fetch all epics in parallel
      sourceSS.toast(`Step 1/3: Fetching epics for ${streamsToProcess.length} value streams...`, '⚡ Parallel Pipeline', -1);
      
      const epicsByValueStream = fetchEpicsForAllValueStreamsParallel(
        programIncrement, 
        streamsToProcess, 
        includePrePlanning
      );
      
      // 1b. Fetch dependencies in parallel
      sourceSS.toast('Step 1/3: Fetching dependencies...', '⚡ Parallel Pipeline', -1);
      const dependenciesByEpic = fetchDependenciesParallel(epicsByValueStream);
      
      // 1c. Collect all unique parent keys for initiative fetch
      sourceSS.toast('Step 1/3: Fetching initiative data...', '⚡ Parallel Pipeline', -1);
      const allParentKeys = new Set();
      
      Object.values(epicsByValueStream).forEach(epics => {
        epics.forEach(epic => {
          if (epic.parentKey) {
            allParentKeys.add(epic.parentKey);
          }
        });
      });
      
      console.log(`  Found ${allParentKeys.size} unique parent keys across all value streams`);
      
      // 1d. Fetch initiative data in parallel
      const initiativeMap = fetchInitiativesParallel(allParentKeys);
      
      // Record refresh timestamp
      const refreshTimestamp = new Date();
      const reportKey = `PI${piNumber}_Iteration${iterationNumber}`;
      recordRefreshTimestamp(reportKey, refreshTimestamp);
      
      // ─────────────────────────────────────────────────────────────────────
      // 1e. OPTIMIZED: Process ALL value streams, collect data in memory
      // ─────────────────────────────────────────────────────────────────────
      sourceSS.toast('Step 1/3: Processing all value streams...', '⚡ Parallel Pipeline', -1);
      
      const allProcessedData = [];
      const rowCountByVS = {};
      
      for (const valueStream of streamsToProcess) {
        const epics = epicsByValueStream[valueStream] || [];
        
        if (epics.length === 0) {
          console.log(`  ${valueStream}: No epics, skipping`);
          rowCountByVS[valueStream] = 0;
          continue;
        }
        
        const epicDataMap = {};
        epics.forEach(epic => {
          epicDataMap[epic.key] = {
            totalStoryPoints: 0,
            closedStoryPoints: 0,
            totalChildren: 0,
            closedChildren: 0,
            risks: [],
            dependencies: dependenciesByEpic[epic.key] || [],
            allChildren: dependenciesByEpic[epic.key] || []
          };
        });
        
        // Process data (in memory only, no write yet)
        const processedData = processGovernanceData(epics, epicDataMap, initiativeMap);
        allProcessedData.push(...processedData);
        rowCountByVS[valueStream] = processedData.length;
        
        console.log(`  ${valueStream}: ${processedData.length} rows processed`);
      }
      
      // ─────────────────────────────────────────────────────────────────────
      // 1f. OPTIMIZED: Write ALL data in ONE operation
      // ─────────────────────────────────────────────────────────────────────
      sourceSS.toast(`Step 1/3: Writing ${allProcessedData.length} rows to sheet...`, '⚡ Parallel Pipeline', -1);
      
      // Single write call: headers + all data + enrichment cache/restore ONCE
      writeGovernanceDataToSheet(
        allProcessedData, 
        programIncrement, 
        'All Value Streams',  // Combined label for logging
        sheetName, 
        true,   // clearExisting - clear sheet and write headers
        true,   // isFirstValueStream - ensure headers are written
        targetSS
      );
      
      const totalRows = allProcessedData.length;
      
      // Log per-VS breakdown
      console.log('\n📊 Rows by Value Stream:');
      Object.entries(rowCountByVS)
        .filter(([vs, count]) => count > 0)
        .sort((a, b) => b[1] - a[1])  // Sort by count descending
        .forEach(([vs, count]) => console.log(`  ${vs}: ${count} rows`));
      
      const step1Time = ((Date.now() - startTime) / 1000).toFixed(1);
      console.log(`\n✓ Step 1 complete in ${step1Time}s - ${totalRows} rows written (single batch)`);
      
      // Get target sheet for subsequent steps
      const targetSheet = targetSS.getSheetByName(sheetName);
      if (!targetSheet) {
        throw new Error(`Sheet "${sheetName}" not found after Step 1`);
      }
      
      // ═══════════════════════════════════════════════════════════════════════
      // STEP 2: EPIC PERSPECTIVES (AI)
      // ═══════════════════════════════════════════════════════════════════════
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 2: Generate Epic Perspectives');
      console.log('─'.repeat(60));
      
      sourceSS.toast('Step 2/3: Generating epic perspectives (AI)...', '⚡ Parallel Pipeline', -1);
      
      generateEpicPerspectivesOnSheet(targetSheet);
      
      const step2Time = ((Date.now() - startTime) / 1000).toFixed(1);
      console.log(`\n✓ Step 2 complete in ${step2Time}s (cumulative)`);
      
      // ═══════════════════════════════════════════════════════════════════════
      // STEP 3: INITIATIVE PERSPECTIVES (AI)
      // ═══════════════════════════════════════════════════════════════════════
      console.log('\n' + '─'.repeat(60));
      console.log('  STEP 3: Generate Initiative Perspectives');
      console.log('─'.repeat(60));
      
      sourceSS.toast('Step 3/3: Generating initiative perspectives (AI)...', '⚡ Parallel Pipeline', -1);
      
      generateInitiativePerspectivesOnSheet(targetSheet);
      
      // COMPLETE
      const totalTime = ((Date.now() - startTime) / 1000).toFixed(1);
      
      console.log('\n' + '═'.repeat(70));
      console.log('  ⚡ PARALLEL PIPELINE COMPLETE');
      console.log('═'.repeat(70));
      console.log(`  Total Time: ${totalTime}s`);
      console.log(`  Rows Processed: ${totalRows}`);
      console.log(`  Initiatives Fetched: ${initiativeMap.size}`);
      console.log(`  Destination: ${destinationConfig.type}`);
      if (newSpreadsheetUrl) {
        console.log(`  New Spreadsheet: ${newSpreadsheetUrl}`);
      }
      console.log(`  Completed: ${new Date().toLocaleString()}`);
      console.log('═'.repeat(70) + '\n');
      
      sourceSS.toast(`✅ Complete! ${totalRows} rows in ${totalTime}s`, '⚡ Pipeline Done', 10);
      
      return { 
        success: true, 
        totalRows, 
        timeSeconds: parseFloat(totalTime),
        sheetName,
        newSpreadsheetUrl,
        initiativeCount: initiativeMap.size
      };
      
    } catch (error) {
      console.error('❌ Pipeline failed:', error);
      sourceSS.toast(`❌ Error: ${error.message}`, 'Pipeline Failed', 10);
      throw error;
    }
  }, { piNumber, iterationNumber, valueStreams: selectedValueStreams?.length || VALUE_STREAMS.length, destination: destinationConfig?.type });
}

/**
 * Parallel version of enrichWithInitiativeData
 * Uses fetchAll() for faster initiative fetching
 */
function enrichWithInitiativeDataParallel() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('  Starting parallel initiative enrichment...');
  
  const sheetName = sheet.getName();
  
  if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
    console.warn('Not a valid report sheet for enrichment');
    return;
  }
  
  try {
    const allHeaders = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const parentKeyCol = allHeaders.indexOf('Parent Key') + 1;
    const issueTypeCol = allHeaders.indexOf('Issue Type') + 1;
    
    if (parentKeyCol === 0 || issueTypeCol === 0) {
      console.error('Could not find required columns (Parent Key, Issue Type)');
      return;
    }
    
    let titleCol = allHeaders.indexOf('Initiative Title') + 1;
    let descCol = allHeaders.indexOf('Initiative Description') + 1;
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 4) {
      console.log('No data rows found');
      return;
    }
    
    const dataRange = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
    
    const initiativeKeys = new Set();
    
    dataRange.forEach((row) => {
      const issueType = row[issueTypeCol - 1];
      let parentKey = row[parentKeyCol - 1];
      
      if (typeof parentKey === 'string' && parentKey.startsWith('=HYPERLINK')) {
        const keyMatch = parentKey.match(/"([A-Z]+-\d+)"/);
        if (keyMatch) {
          parentKey = keyMatch[1];
        }
      }
      
      if (issueType === 'Epic' && parentKey) {
        initiativeKeys.add(parentKey);
      }
    });
    
    console.log(`  Found ${initiativeKeys.size} unique initiatives to fetch`);
    
    if (initiativeKeys.size === 0) {
      console.log('  No parent keys found for Epics');
      return;
    }
    
    const initiativeMap = fetchInitiativesParallel(initiativeKeys);
    
    console.log(`  Fetched ${initiativeMap.size} initiatives`);
    
    if (titleCol === 0) {
      titleCol = sheet.getLastColumn() + 1;
      sheet.getRange(4, titleCol).setValue('Initiative Title')
        .setFontWeight('bold')
        .setBackground('#7e57c2')
        .setFontColor('white');
      sheet.setColumnWidth(titleCol, 350);
    }
    if (descCol === 0) {
      descCol = sheet.getLastColumn() + 1;
      sheet.getRange(4, descCol).setValue('Initiative Description')
        .setFontWeight('bold')
        .setBackground('#7e57c2')
        .setFontColor('white');
      sheet.setColumnWidth(descCol, 350);
    }
    
    const titleData = [];
    const descData = [];
    
    for (let i = 0; i < dataRange.length; i++) {
      let parentKey = dataRange[i][parentKeyCol - 1];
      
      if (typeof parentKey === 'string' && parentKey.startsWith('=HYPERLINK')) {
        const keyMatch = parentKey.match(/"([A-Z]+-\d+)"/);
        if (keyMatch) {
          parentKey = keyMatch[1];
        }
      }
      
      if (parentKey && initiativeMap.has(parentKey)) {
        const initiative = initiativeMap.get(parentKey);
        titleData.push([initiative.title]);
        descData.push([initiative.description]);
      } else {
        titleData.push(['']);
        descData.push(['']);
      }
    }
    
    if (titleData.length > 0) {
      sheet.getRange(5, titleCol, titleData.length, 1).setValues(titleData);
      sheet.getRange(5, descCol, descData.length, 1).setValues(descData);
      sheet.getRange(5, titleCol, titleData.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.getRange(5, descCol, descData.length, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    }
    
    const enrichedCount = titleData.filter(row => row[0] !== '').length;
    console.log(`  ✓ Initiative enrichment complete: ${enrichedCount} rows enriched`);
    
  } catch (error) {
    console.error('Initiative Enrichment Error:', error);
    throw error;
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 7: SCHEDULER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Sets up the bi-weekly pipeline scheduler
 */
function setupPipelineScheduler() {
  const ui = SpreadsheetApp.getUi();
  
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runScheduledPipelineCheck') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });
  
  if (removedCount > 0) {
    console.log(`Removed ${removedCount} existing scheduler triggers`);
  }
  
  ScriptApp.newTrigger('runScheduledPipelineCheck')
    .timeBased()
    .everyMinutes(30)
    .create();
  
  const referenceDate = new Date('2025-01-08T08:00:00');
  PropertiesService.getScriptProperties().setProperty('SCHEDULER_REFERENCE_DATE', referenceDate.toISOString());
  
  const config = {
    piNumber: 13,
    iterationNumber: 2,
    valueStreams: null,
    enabled: true
  };
  PropertiesService.getScriptProperties().setProperty('SCHEDULER_CONFIG', JSON.stringify(config));
  
  ui.alert('Scheduler Setup', 
    'Pipeline scheduler has been set up!\n\n' +
    '• Runs every 30 minutes\n' +
    '• Active: Every other Wednesday 8am ET - Thursday 3pm ET\n' +
    '• Current config: PI 13, Iteration 2, All value streams\n\n' +
    'Use "Update Scheduler Config" to change settings.',
    ui.ButtonSet.OK);
  
  console.log('✅ Pipeline scheduler configured');
}

/**
 * Removes all scheduler triggers
 */
function removeScheduler() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  let removedCount = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runScheduledPipelineCheck') {
      ScriptApp.deleteTrigger(trigger);
      removedCount++;
    }
  });
  
  PropertiesService.getScriptProperties().deleteProperty('SCHEDULER_CONFIG');
  
  ui.alert('Scheduler Removed', 
    `Removed ${removedCount} scheduler trigger(s).\nThe pipeline will no longer run automatically.`,
    ui.ButtonSet.OK);
}

/**
 * Updates the scheduler configuration
 */
function updateSchedulerConfig() {
  const ui = SpreadsheetApp.getUi();
  
  const configStr = PropertiesService.getScriptProperties().getProperty('SCHEDULER_CONFIG');
  const currentConfig = configStr ? JSON.parse(configStr) : { piNumber: 13, iterationNumber: 2, valueStreams: null, enabled: true };
  
  const piResponse = ui.prompt('Update Scheduler', 
    `Enter PI number (current: ${currentConfig.piNumber}):`, 
    ui.ButtonSet.OK_CANCEL);
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const newPi = parseInt(piResponse.getResponseText().trim()) || currentConfig.piNumber;
  
  const iterResponse = ui.prompt('Update Scheduler', 
    `Enter Iteration number (current: ${currentConfig.iterationNumber}):`, 
    ui.ButtonSet.OK_CANCEL);
  
  if (iterResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const newIter = parseInt(iterResponse.getResponseText().trim()) || currentConfig.iterationNumber;
  
  const newConfig = {
    piNumber: newPi,
    iterationNumber: newIter,
    valueStreams: null,
    enabled: true
  };
  
  PropertiesService.getScriptProperties().setProperty('SCHEDULER_CONFIG', JSON.stringify(newConfig));
  
  ui.alert('Config Updated', 
    `Scheduler will now run:\n• PI: ${newPi}\n• Iteration: ${newIter}\n• All value streams`,
    ui.ButtonSet.OK);
}

/**
 * Triggered function - checks if in scheduled window and runs pipeline
 */
function runScheduledPipelineCheck() {
  const props = PropertiesService.getScriptProperties();
  
  const configStr = props.getProperty('SCHEDULER_CONFIG');
  if (!configStr) {
    console.log('No scheduler config found, skipping');
    return;
  }
  
  const config = JSON.parse(configStr);
  if (!config.enabled) {
    console.log('Scheduler is disabled, skipping');
    return;
  }
  
  const now = new Date();
  const etOffset = isDSTActive(now) ? -4 : -5;
  const etNow = new Date(now.getTime() + (now.getTimezoneOffset() + etOffset * 60) * 60000);
  
  const dayOfWeek = etNow.getDay();
  const hour = etNow.getHours();
  const minute = etNow.getMinutes();
  
  console.log(`Scheduler check: ${etNow.toLocaleString()} ET (Day: ${dayOfWeek}, Hour: ${hour})`);
  
  let inWindow = false;
  
  if (dayOfWeek === 3) {
    inWindow = (hour >= 8);
  } else if (dayOfWeek === 4) {
    inWindow = (hour < 15 || (hour === 15 && minute === 0));
  }
  
  if (!inWindow) {
    console.log('Outside scheduled window, skipping');
    return;
  }
  
  const referenceStr = props.getProperty('SCHEDULER_REFERENCE_DATE');
  if (referenceStr) {
    const referenceDate = new Date(referenceStr);
    const daysSinceReference = Math.floor((etNow - referenceDate) / (1000 * 60 * 60 * 24));
    const weeksSinceReference = Math.floor(daysSinceReference / 7);
    
    if (weeksSinceReference % 2 !== 0) {
      console.log(`Off-week (week ${weeksSinceReference}), skipping`);
      return;
    }
  }
  
  const lastRunStr = props.getProperty('SCHEDULER_LAST_RUN');
  if (lastRunStr) {
    const lastRun = new Date(lastRunStr);
    const minutesSinceLastRun = (now - lastRun) / (1000 * 60);
    
    if (minutesSinceLastRun < 25) {
      console.log(`Already ran ${minutesSinceLastRun.toFixed(0)} minutes ago, skipping`);
      return;
    }
  }
  
  props.setProperty('SCHEDULER_LAST_RUN', now.toISOString());
  
  console.log('═'.repeat(60));
  console.log('SCHEDULED PIPELINE RUN STARTING');
  console.log(`PI: ${config.piNumber}, Iteration: ${config.iterationNumber}`);
  console.log('═'.repeat(60));
  
  try {
    const result = runFullPipelineParallel(
      config.piNumber,
      config.iterationNumber,
      config.valueStreams
    );
    
    appendToSchedulerLog({
      timestamp: now.toISOString(),
      status: 'success',
      pi: config.piNumber,
      iteration: config.iterationNumber,
      rows: result.totalRows,
      timeSeconds: result.timeSeconds
    });
    
    console.log('✅ Scheduled run complete');
    
  } catch (error) {
    console.error('❌ Scheduled run failed:', error);
    
    appendToSchedulerLog({
      timestamp: now.toISOString(),
      status: 'error',
      pi: config.piNumber,
      iteration: config.iterationNumber,
      error: error.message
    });
  }
}

/**
 * Checks if Daylight Saving Time is active
 */
function isDSTActive(date) {
  const jan = new Date(date.getFullYear(), 0, 1);
  const jul = new Date(date.getFullYear(), 6, 1);
  const stdOffset = Math.max(jan.getTimezoneOffset(), jul.getTimezoneOffset());
  return date.getTimezoneOffset() < stdOffset;
}

/**
 * Appends entry to scheduler log
 */
function appendToSchedulerLog(entry) {
  const props = PropertiesService.getScriptProperties();
  const logStr = props.getProperty('SCHEDULER_LOG') || '[]';
  const log = JSON.parse(logStr);
  
  log.push(entry);
  
  while (log.length > 100) {
    log.shift();
  }
  
  props.setProperty('SCHEDULER_LOG', JSON.stringify(log));
}

/**
 * Shows the scheduler log
 */
function viewSchedulerLog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const logStr = props.getProperty('SCHEDULER_LOG') || '[]';
  const log = JSON.parse(logStr);
  
  if (log.length === 0) {
    ui.alert('Scheduler Log', 'No scheduled runs recorded yet.', ui.ButtonSet.OK);
    return;
  }
  
  const recent = log.slice(-10).reverse();
  let message = 'Last 10 scheduled runs:\n\n';
  
  recent.forEach((entry, idx) => {
    const date = new Date(entry.timestamp);
    const dateStr = date.toLocaleString();
    
    if (entry.status === 'success') {
      message += `${idx + 1}. ✅ ${dateStr}\n   PI ${entry.pi}, Iter ${entry.iteration}: ${entry.rows} rows in ${entry.timeSeconds}s\n\n`;
    } else {
      message += `${idx + 1}. ❌ ${dateStr}\n   PI ${entry.pi}, Iter ${entry.iteration}: ${entry.error}\n\n`;
    }
  });
  
  ui.alert('Scheduler Log', message, ui.ButtonSet.OK);
}

/**
 * Tests the scheduled pipeline manually
 */
function testScheduledPipeline() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const configStr = props.getProperty('SCHEDULER_CONFIG');
  if (!configStr) {
    ui.alert('Error', 'No scheduler config found. Run "Setup Pipeline Scheduler" first.', ui.ButtonSet.OK);
    return;
  }
  
  const config = JSON.parse(configStr);
  
  const confirm = ui.alert('Test Scheduled Pipeline',
    `This will run the parallel pipeline with:\n• PI: ${config.piNumber}\n• Iteration: ${config.iterationNumber}\n• All value streams\n\nContinue?`,
    ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  runFullPipelineParallel(config.piNumber, config.iterationNumber, config.valueStreams);
}


// ═══════════════════════════════════════════════════════════════════════════════
// SECTION 8: UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Updates pipeline status in cache for UI polling
 */
function updatePipelineStatus(cache, status) {
  cache.put('pipeline_status', JSON.stringify(status), 300);
}

/**
 * Gets current pipeline status for UI polling
 */
function getFullPipelineStatus() {
  const cache = CacheService.getUserCache();
  const status = cache.get('pipeline_status');
  return status ? JSON.parse(status) : null;
}

/**
 * Gets report generation status
 */
function getReportGenerationStatus() {
  return CacheService.getUserCache().get('report_status') || 'Initializing...';
}

/**
 * Gets changelog analysis status
 */
function getChangelogAnalysisStatus() {
  return CacheService.getUserCache().get('changelog_status') || 'Initializing...';
}

/**
 * Gets data for refresh dialog
 */
function getRefreshDialogData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const piMatch = sheet.getName().match(/PI (\d+)/);
  const piNumber = piMatch ? piMatch[1] : null;
  return {
    piNumber: piNumber,
    valueStreams: VALUE_STREAMS
  };
}

/**
 * Runs value stream refresh
 */
function runValueStreamRefresh(piNumber, valueStream) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Refreshing data for ${valueStream}...`, '🔄 Refreshing', 10
  );
  refreshValueStreamData(piNumber, valueStream);
}

/**
 * Menu entry point for initiative enrichment
 */
function menuEnrichWithInitiativeData() {
  enrichWithInitiativeData();
}

/**
 * Sets up Gemini AI configuration
 */
function setupGeminiConfiguration() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Gemini API Configuration', 
    'Enter your Gemini API key:', 
    ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
      ui.alert('Success', 'Gemini API key saved successfully!', ui.ButtonSet.OK);
    }
  }
}

/**
 * Get list of available Portfolio Initiatives
 * Reads from Configuration sheet with fallback to iteration sheets
 */
function getAvailablePortfolios() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const possibleNames = ['Configuration', 'Configuration Template', 'Config', 'Slide Configuration'];
    let configSheet = null;
    let foundSheetName = '';
    
    for (const name of possibleNames) {
      configSheet = ss.getSheetByName(name);
      if (configSheet) {
        foundSheetName = name;
        console.log(`Found configuration sheet: "${name}"`);
        break;
      }
    }
    
    if (!configSheet) {
      console.warn('No configuration sheet found. Tried:', possibleNames.join(', '));
      return getPortfoliosFromIterationSheets(ss);
    }
    
    const lastRow = configSheet.getLastRow();
    const lastCol = configSheet.getLastColumn();
    
    if (lastRow < 2) {
      console.warn('Configuration sheet appears empty');
      return getPortfoliosFromIterationSheets(ss);
    }
    
    let headerRow = null;
    let headers = [];
    
    for (let row = 1; row <= Math.min(10, lastRow); row++) {
      const rowData = configSheet.getRange(row, 1, 1, lastCol).getValues()[0];
      const nonEmptyCount = rowData.filter(cell => cell && cell.toString().trim() !== '').length;
      
      if (nonEmptyCount >= 3) {
        headerRow = row;
        headers = rowData;
        console.log(`Found header row at row ${row} with ${nonEmptyCount} headers`);
        break;
      }
    }
    
    if (!headerRow) {
      console.warn('Could not find header row in configuration sheet');
      return getPortfoliosFromIterationSheets(ss);
    }
    
    const portfolioHeaders = ['Portfolio Initiative', 'Portfolio', 'Program', 'Value Stream'];
    let portfolioColIndex = -1;
    
    for (const searchHeader of portfolioHeaders) {
      portfolioColIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().includes(searchHeader.toLowerCase())
      );
      if (portfolioColIndex >= 0) {
        console.log(`Found portfolio column "${headers[portfolioColIndex]}" at column ${portfolioColIndex + 1}`);
        break;
      }
    }
    
    if (portfolioColIndex < 0) {
      console.warn('Could not find portfolio column. Headers:', headers.filter(h => h).join(', '));
      return getPortfoliosFromIterationSheets(ss);
    }
    
    const dataStartRow = headerRow + 1;
    const dataRows = lastRow - headerRow;
    
    if (dataRows <= 0) {
      console.warn('No data rows after header');
      return getPortfoliosFromIterationSheets(ss);
    }
    
    const portfolioCol = portfolioColIndex + 1;
    const portfolioValues = configSheet.getRange(dataStartRow, portfolioCol, dataRows, 1).getValues();
    
    const portfolios = [];
    portfolioValues.forEach(row => {
      const value = row[0];
      if (value && value.toString().trim() !== '') {
        portfolios.push(value.toString().trim());
      }
    });
    
    const uniquePortfolios = [...new Set(portfolios)];
    
    uniquePortfolios.sort((a, b) => {
      if (a === 'INFOSEC') return -1;
      if (b === 'INFOSEC') return 1;
      return a.localeCompare(b);
    });
    
    console.log(`Found ${uniquePortfolios.length} portfolios from "${foundSheetName}"`);
    return uniquePortfolios;
    
  } catch (error) {
    console.error('Error getting available portfolios:', error);
    return [];
  }
}

/**
 * Fallback: get portfolios from iteration sheets
 */
function getPortfoliosFromIterationSheets(ss) {
  console.log('Falling back to iteration sheets for portfolio list...');
  
  const sheets = ss.getSheets();
  const portfolios = new Set();
  
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name.match(/^PI \d+ - Iteration \d+$/)) {
      try {
        const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
        const portfolioCol = headers.indexOf('Portfolio Initiative');
        
        if (portfolioCol >= 0) {
          const lastRow = sheet.getLastRow();
          if (lastRow > 4) {
            const values = sheet.getRange(5, portfolioCol + 1, lastRow - 4, 1).getValues();
            values.forEach(row => {
              if (row[0] && row[0].toString().trim()) {
                portfolios.add(row[0].toString().trim());
              }
            });
          }
        }
      } catch (e) {
        // Skip sheets that cause errors
      }
      
      if (portfolios.size > 0) break;
    }
  }
  
  const result = [...portfolios].sort((a, b) => {
    if (a === 'INFOSEC') return -1;
    if (b === 'INFOSEC') return 1;
    return a.localeCompare(b);
  });
  
  console.log(`Found ${result.length} portfolios from iteration sheets: ${result.join(', ')}`);
  return result;
}

/**
 * Gets portfolios for dialogs (alias for backward compatibility)
 */
function getAvailablePortfoliosForDialog() {
  return getAvailablePortfolios();
}

/**
 * Gets slide dialog data
 */
function getSlideDialogData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const piNumbers = new Set();
  const iterationsByPI = {};
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    const match = name.match(/^PI (\d+) - Iteration (\d+)$/);
    if (match) {
      const pi = parseInt(match[1]);
      const iter = parseInt(match[2]);
      
      piNumbers.add(pi);
      
      if (!iterationsByPI[pi]) {
        iterationsByPI[pi] = new Set();
      }
      iterationsByPI[pi].add(iter);
    }
  });
  
  const pis = Array.from(piNumbers).sort((a, b) => b - a);
  
  Object.keys(iterationsByPI).forEach(pi => {
    iterationsByPI[pi] = Array.from(iterationsByPI[pi]).sort((a, b) => a - b);
  });
  
  const portfolios = getAvailablePortfolios();
  
  // Get value streams from the constant defined in CORE.gs
  const valueStreams = typeof VALUE_STREAMS !== 'undefined' ? VALUE_STREAMS : [];
  
  console.log(`getSlideDialogData: ${pis.length} PIs, ${portfolios.length} portfolios, ${valueStreams.length} value streams`);
  
  return {
    pis: pis,
    iterationsByPI: iterationsByPI,
    portfolios: portfolios,
    valueStreams: valueStreams
  };
}

/**
 * Gets available PIs for slides
 */
function getAvailablePIsForSlides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const piNumbers = new Set();
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    const match = name.match(/^PI (\d+) - /);
    if (match) {
      piNumbers.add(parseInt(match[1]));
    }
  });
  
  return Array.from(piNumbers).sort((a, b) => b - a);
}

/**
 * Gets available iterations for a PI
 */
function getAvailableIterationsForPI(piNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const iterations = new Set();
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    const match = name.match(new RegExp(`^PI ${piNumber} - Iteration (\\d+)$`));
    if (match) {
      iterations.add(parseInt(match[1]));
    }
  });
  
  return Array.from(iterations).sort((a, b) => a - b);
}

/**
 * Generates validation report from dialog
 */
function generatePredictabilityValidationFromDialog(piNumber, valueStream) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`VALIDATION GENERATION FROM DIALOG`);
  console.log(`PI: ${piNumber}, Value Stream: ${valueStream}`);
  console.log(`${'='.repeat(60)}\n`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast(`Generating validation for PI ${piNumber} - ${valueStream}...`, 'Processing', -1);
  
  try {
    generatePredictabilityValidation(piNumber, valueStream);
    ss.toast(`✅ Validation sheet created successfully!`, 'Complete', 5);
  } catch (error) {
    console.error('Error generating validation:', error);
    ss.toast('', '', 1);
    throw error;
  }
}

/**
 * Opens the Activity Log sheet
 */
function openActivityLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(ACTIVITY_LOG_CONFIG.SHEET_NAME);
  
  if (logSheet) {
    ss.setActiveSheet(logSheet);
    SpreadsheetApp.getUi().alert('Activity Log', 
      'Viewing Activity Log sheet.\n\nTip: You can filter and sort this data like any other sheet!', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No Activity Log', 
      'No activity has been logged yet. The log sheet will be created automatically when functions are executed.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Adds badge columns to changelog
 */
function addBadgeColumnsToChangelog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.prompt(
    'Add Badge Columns',
    'Enter the PI number:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const piNumber = parseInt(response.getResponseText().trim());
  if (isNaN(piNumber)) {
    ui.alert('Invalid PI number');
    return;
  }
  
  const changelogSheetName = `PI ${piNumber} Changelog`;
  const sheet = ss.getSheetByName(changelogSheetName);
  
  if (!sheet) {
    ui.alert(`Changelog sheet "${changelogSheetName}" not found.`);
    return;
  }
  
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(5, 1, 1, lastCol).getValues()[0];
  
  const badgeColumnSuffixes = [
    'Badge',
    'Status Badge', 
    'Status Note',
    'At Risk',
    'Iteration Risk',
    'Closed This Iter',
    'Deferred This Iter',
    'Is New',
    'Qualifying Reasons'
  ];
  
  let columnsAdded = 0;
  
  for (let iter = 2; iter <= 6; iter++) {
    const noChangesHeader = `Iteration ${iter} - NO CHANGES`;
    const noChangesIndex = headers.indexOf(noChangesHeader);
    
    if (noChangesIndex === -1) {
      console.log(`NO CHANGES column not found for Iteration ${iter} - skipping`);
      continue;
    }
    
    badgeColumnSuffixes.forEach((suffix, idx) => {
      const fullHeader = `Iteration ${iter} - ${suffix}`;
      if (!headers.includes(fullHeader)) {
        const insertPosition = noChangesIndex + 2 + columnsAdded;
        sheet.insertColumnAfter(insertPosition);
        sheet.getRange(5, insertPosition + 1).setValue(fullHeader)
          .setFontWeight('bold')
          .setBackground('#e8d5f0')
          .setFontColor('#333333');
        
        headers.splice(insertPosition, 0, fullHeader);
        columnsAdded++;
        console.log(`Added column: ${fullHeader}`);
      }
    });
  }
  
  if (columnsAdded > 0) {
    ui.alert(`Added ${columnsAdded} badge columns to the changelog.`);
  } else {
    ui.alert('All badge columns already exist.');
  }
}
function showClosedItemsReportDialog() {
  const html = HtmlService.createHtmlOutput(getClosedItemsReportDialogHTML())
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Closed Items Report');
}

/**
 * Returns HTML for the Closed Items Report dialog
 */
function getClosedItemsReportDialogHTML() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Google Sans', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            color: #202124;
          }
          h2 {
            color: #1a73e8;
            margin: 0 0 8px 0;
            font-size: 20px;
            font-weight: 400;
          }
          .subtitle {
            color: #5f6368;
            font-size: 13px;
            margin-bottom: 20px;
          }
          .section {
            margin-bottom: 16px;
          }
          .section-title {
            font-size: 13px;
            font-weight: 500;
            color: #202124;
            margin-bottom: 6px;
          }
          select, input[type="text"] {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            color: #202124;
            background-color: white;
            box-sizing: border-box;
          }
          select:focus, input:focus {
            outline: none;
            border-color: #1a73e8;
          }
          .info-box {
            background: #e8f5e9;
            border-left: 4px solid #4caf50;
            padding: 12px 16px;
            margin: 16px 0;
            border-radius: 4px;
            font-size: 12px;
          }
          .button-container {
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            margin-top: 20px;
          }
          .primary-button {
            background: #1a73e8;
            color: white;
            border: none;
            padding: 10px 24px;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
          }
          .primary-button:hover { background: #1557b0; }
          .primary-button:disabled { background: #dadce0; cursor: not-allowed; }
          .secondary-button {
            background: white;
            color: #1a73e8;
            border: 1px solid #dadce0;
            padding: 10px 24px;
            border-radius: 4px;
            font-size: 14px;
            cursor: pointer;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 30px;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #1a73e8;
            border-radius: 50%;
            width: 36px;
            height: 36px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .error { color: #d93025; font-size: 12px; margin-top: 4px; display: none; }
        </style>
      </head>
      <body>
        <h2>✅ Closed Items Report</h2>
        <p class="subtitle">Generate a presentation of all items closed during this PI</p>
        
        <div class="info-box">
          <strong>📋 What's included:</strong><br>
          All epics and features closed across ALL iterations of the selected PI, 
          organized by portfolio with closure details.
        </div>
        
        <div class="section">
          <div class="section-title">Program Increment (PI)</div>
          <select id="piSelect" onchange="updatePresentationName()">
            <option value="">Select PI...</option>
          </select>
          <div id="piError" class="error">Please select a PI</div>
        </div>
        
        <div class="section">
          <div class="section-title">Presentation Name</div>
          <input type="text" id="presentationName" placeholder="Auto-generated..." />
        </div>
        
        <div id="loading" class="loading">
          <div class="spinner"></div>
          <p style="margin-top: 12px; color: #5f6368;">Generating report...</p>
        </div>
        
        <div class="button-container">
          <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
          <button id="generateBtn" class="primary-button" onclick="generateReport()">Generate Report</button>
        </div>
        
        <script>
          window.onload = function() {
            google.script.run
              .withSuccessHandler(function(data) {
                const piSelect = document.getElementById('piSelect');
                data.pis.forEach(function(pi) {
                  const opt = document.createElement('option');
                  opt.value = pi;
                  opt.textContent = 'PI ' + pi;
                  piSelect.appendChild(opt);
                });
              })
              .getSlideDialogData();
          };
          
          function updatePresentationName() {
            const pi = document.getElementById('piSelect').value;
            if (pi) {
              const today = new Date();
              const dateStr = (today.getMonth()+1) + '.' + today.getDate() + '.' + today.getFullYear();
              document.getElementById('presentationName').value = 
                'PI ' + pi + ' - Closed Items Report - ' + dateStr;
            }
          }
          
          function generateReport() {
            const pi = document.getElementById('piSelect').value;
            const name = document.getElementById('presentationName').value.trim();
            
            if (!pi) {
              document.getElementById('piError').style.display = 'block';
              return;
            }
            document.getElementById('piError').style.display = 'none';
            
            if (!name) {
              alert('Please enter a presentation name');
              return;
            }
            
            document.getElementById('loading').style.display = 'block';
            document.getElementById('generateBtn').disabled = true;
            
            google.script.run
              .withSuccessHandler(function(result) {
                document.getElementById('loading').style.display = 'none';
                if (result && result.success && result.url) {
                  window.open(result.url, '_blank');
                  alert('✅ Closed Items Report generated!\\n\\n' + result.closedCount + ' closed items found.');
                  google.script.host.close();
                } else {
                  alert('Error: ' + (result.error || 'Unknown error'));
                  document.getElementById('generateBtn').disabled = false;
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('generateBtn').disabled = false;
                alert('Error: ' + error.message);
              })
              .generateClosedItemsReport(parseInt(pi), name);
          }
        </script>
      </body>
    </html>
  `;
}

/**
 * Generates a report of all closed items across all iterations of a PI
 */
function generateClosedItemsReport(piNumber, presentationName) {
  return logActivity('Closed Items Report Generation', () => {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      console.log('\n╔══════════════════════════════════════════════════════════════╗');
      console.log('║           CLOSED ITEMS REPORT                                ║');
      console.log('╠══════════════════════════════════════════════════════════════╣');
      console.log(`║  PI: ${piNumber}  |  ${new Date().toLocaleString()}`);
      console.log('╚══════════════════════════════════════════════════════════════╝\n');
      
      // Find all iteration sheets for this PI
      const sheets = ss.getSheets();
      const iterationSheets = sheets.filter(s => {
        const match = s.getName().match(/^PI (\d+) - Iteration (\d+)$/);
        return match && parseInt(match[1]) === piNumber;
      }).sort((a, b) => {
        const iterA = parseInt(a.getName().match(/Iteration (\d+)/)[1]);
        const iterB = parseInt(b.getName().match(/Iteration (\d+)/)[1]);
        return iterA - iterB;
      });
      
      if (iterationSheets.length === 0) {
        throw new Error(`No iteration sheets found for PI ${piNumber}`);
      }
      
      console.log(`Found ${iterationSheets.length} iteration sheets for PI ${piNumber}`);
      
      // Collect all closed items from all iterations
      const closedItems = [];
      const seenKeys = new Set();
      
      iterationSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        const iterMatch = sheetName.match(/Iteration (\d+)/);
        const iterNum = iterMatch ? parseInt(iterMatch[1]) : 0;
        
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;
        
        const headers = data[0];
        const keyIdx = headers.indexOf('Key');
        const statusIdx = headers.indexOf('Status');
        const summaryIdx = headers.indexOf('Summary');
        const vsIdx = headers.indexOf('Value Stream/Org');
        const portfolioIdx = headers.indexOf('Portfolio Initiative');
        const programIdx = headers.indexOf('Program Initiative');
        const typeIdx = headers.indexOf('Issue Type');
        const resolutionIdx = headers.indexOf('Resolution');
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const key = row[keyIdx];
          const status = (row[statusIdx] || '').toString().toLowerCase();
          const resolution = (row[resolutionIdx] || '').toString().toLowerCase();
          
          // Check if closed (and not duplicate resolution)
          if (status === 'closed' && resolution !== 'duplicate' && !seenKeys.has(key)) {
            seenKeys.add(key);
            closedItems.push({
              key: key,
              summary: row[summaryIdx] || '',
              valueStream: row[vsIdx] || 'Unknown',
              portfolio: row[portfolioIdx] || 'No Portfolio',
              program: row[programIdx] || 'No Program',
              issueType: row[typeIdx] || 'Epic',
              closedInIteration: iterNum,
              resolution: row[resolutionIdx] || 'Done'
            });
          }
        }
      });
      
      console.log(`Found ${closedItems.length} unique closed items across all iterations`);
      
      if (closedItems.length === 0) {
        SpreadsheetApp.getUi().alert('No Closed Items', 
          `No closed items found for PI ${piNumber}.`, 
          SpreadsheetApp.getUi().ButtonSet.OK);
        return { success: true, url: null, closedCount: 0 };
      }
      
      // Group by portfolio
      const byPortfolio = {};
      closedItems.forEach(item => {
        const portfolio = item.portfolio || 'No Portfolio';
        if (!byPortfolio[portfolio]) {
          byPortfolio[portfolio] = [];
        }
        byPortfolio[portfolio].push(item);
      });
      
      // Create presentation
      const templateFile = DriveApp.getFileById(TEMPLATE_CONFIG.TEMPLATE_ID);
      const copiedFile = templateFile.makeCopy(presentationName);
      const presentation = SlidesApp.openById(copiedFile.getId());
      
      // Update title slide
      const titleSlide = presentation.getSlides()[0];
      const shapes = titleSlide.getShapes();
      shapes.forEach(shape => {
        const text = shape.getText().asString();
        if (text.includes('{{TITLE}}') || text.includes('Governance')) {
          shape.getText().setText(`PI ${piNumber} - Closed Items Report`);
        }
        if (text.includes('{{SUBTITLE}}') || text.includes('Iteration')) {
          shape.getText().setText(`All closed items across all iterations\nGenerated: ${new Date().toLocaleDateString()}`);
        }
      });
      
      // Create slides for each portfolio
      const sortedPortfolios = Object.keys(byPortfolio).sort();
      let slideIndex = 1;
      
      sortedPortfolios.forEach(portfolioName => {
        const items = byPortfolio[portfolioName];
        
        // Group items by iteration for display
        const byIteration = {};
        items.forEach(item => {
          const iter = item.closedInIteration || 'Unknown';
          if (!byIteration[iter]) byIteration[iter] = [];
          byIteration[iter].push(item);
        });
        
        // Build slide content
        let content = `${portfolioName}\n${items.length} items closed\n\n`;
        
        Object.keys(byIteration).sort((a, b) => parseInt(a) - parseInt(b)).forEach(iter => {
          content += `── Iteration ${iter} ──\n`;
          byIteration[iter].forEach(item => {
            content += `• ${item.key}: ${item.summary.substring(0, 60)}${item.summary.length > 60 ? '...' : ''}\n`;
            content += `  [${item.valueStream}] ${item.resolution}\n`;
          });
          content += '\n';
        });
        
        // Create slide
        const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        
        // Add title
        const titleBox = slide.insertTextBox(portfolioName, 20, 20, 680, 40);
        titleBox.getText().getTextStyle()
          .setFontFamily('Montserrat')
          .setFontSize(18)
          .setBold(true)
          .setForegroundColor('#663399');
        
        // Add count
        const countBox = slide.insertTextBox(`${items.length} items closed in PI ${piNumber}`, 20, 55, 680, 25);
        countBox.getText().getTextStyle()
          .setFontFamily('Lato')
          .setFontSize(11)
          .setForegroundColor('#666666');
        
        // Add items list
        let yPos = 90;
        Object.keys(byIteration).sort((a, b) => parseInt(a) - parseInt(b)).forEach(iter => {
          // Iteration header
          const iterBox = slide.insertTextBox(`Iteration ${iter}`, 20, yPos, 680, 20);
          iterBox.getText().getTextStyle()
            .setFontFamily('Lato')
            .setFontSize(10)
            .setBold(true)
            .setForegroundColor('#1a73e8');
          yPos += 22;
          
          // Items
          byIteration[iter].slice(0, 8).forEach(item => { // Limit to 8 per iteration to fit
            const itemText = `• ${item.key}: ${item.summary.substring(0, 55)}${item.summary.length > 55 ? '...' : ''}`;
            const itemBox = slide.insertTextBox(itemText, 30, yPos, 660, 16);
            itemBox.getText().getTextStyle()
              .setFontFamily('Lato')
              .setFontSize(9)
              .setForegroundColor('#333333');
            yPos += 18;
            
            // Value stream line
            const vsText = `  [${item.valueStream}]`;
            const vsBox = slide.insertTextBox(vsText, 40, yPos, 640, 14);
            vsBox.getText().getTextStyle()
              .setFontFamily('Lato')
              .setFontSize(8)
              .setForegroundColor('#666666');
            yPos += 16;
            
            if (yPos > 360) return; // Stop if we're running out of space
          });
          
          if (byIteration[iter].length > 8) {
            const moreBox = slide.insertTextBox(`  ... and ${byIteration[iter].length - 8} more`, 40, yPos, 640, 14);
            moreBox.getText().getTextStyle()
              .setFontFamily('Lato')
              .setFontSize(8)
              .setItalic(true)
              .setForegroundColor('#999999');
            yPos += 18;
          }
          
          yPos += 8; // Space between iterations
          
          if (yPos > 360) return; // Stop if running out of space
        });
        
        slideIndex++;
      });
      
      // Clean up template slides (keep only title and our new slides)
      const allSlides = presentation.getSlides();
      // Remove template slides (indexes 1-4 typically) but keep title (0) and our new slides
      for (let i = Math.min(4, allSlides.length - sortedPortfolios.length - 1); i >= 1; i--) {
        try {
          allSlides[i].remove();
        } catch (e) {
          console.log(`Could not remove template slide ${i}: ${e.message}`);
        }
      }
      
      const url = presentation.getUrl();
      
      console.log(`\n✅ Closed Items Report complete: ${url}`);
      console.log(`   Total slides: ${presentation.getSlides().length}`);
      console.log(`   Portfolios: ${sortedPortfolios.length}`);
      console.log(`   Closed items: ${closedItems.length}`);
      
      return {
        success: true,
        url: url,
        closedCount: closedItems.length,
        portfolioCount: sortedPortfolios.length
      };
      
    } catch (error) {
      console.error('Closed Items Report error:', error);
      return { success: false, error: error.message };
    }
  }, { piNumber, presentationName });
}
/**
 * Creates presentation with slides (router function)
 */
function createPresentationWithSlides(presentationName, data) {
  if (data.metadata && data.metadata.iterationNumber > 1) {
    return createChangesOnlyPresentation(
      presentationName, 
      data, 
      data.metadata.piNumber, 
      data.metadata.iterationNumber
    );
  } else {
    return createFormattedPresentation(presentationName, data);
  }
}
function getExistingIterationSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  const iterationSheets = sheets
    .map(sheet => sheet.getName())
    .filter(name => name.match(/^PI \d+ - Iteration \d+$/))
    .sort((a, b) => {
      const matchA = a.match(/PI (\d+) - Iteration (\d+)/);
      const matchB = b.match(/PI (\d+) - Iteration (\d+)/);
      if (matchA && matchB) {
        const piDiff = parseInt(matchB[1]) - parseInt(matchA[1]);
        if (piDiff !== 0) return piDiff;
        return parseInt(matchA[2]) - parseInt(matchB[2]);
      }
      return a.localeCompare(b);
    });
  
  return iterationSheets;
}
