/**
 * ════════════════════════════════════════════════════════════════════════════
 * PI DEFERRALS SYSTEM - WITH PI SELECTION DIALOG
 * ════════════════════════════════════════════════════════════════════════════
 * 
 * Generates deferral presentations from Iteration 6 data
 * Can be run from ANY sheet - presents dialog to select PI
 * Matches predictability report workflow
 */

function generateDeferralsSlide()
{

  showDeferralsDialog()

}

// ═══════════════════════════════════════════════════════════════════════════
// DIALOG FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Shows dialog for generating PI Deferrals presentation
 * Can be called from any sheet
 */
function showDeferralsDialog() {
  const html = HtmlService.createHtmlOutput(getDeferralsDialogHTML())
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '🔴 Generate PI Deferrals');
}

/**
 * Gets available PIs that have Iteration 6 tabs
 * Returns array of {number, label} objects
 */
function getAvailablePIsForDeferrals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const availablePIs = [];
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const match = sheetName.match(/^PI (\d+) - Iteration 6$/i);
    if (match) {
      const piNumber = match[1];
      availablePIs.push({
        number: piNumber,
        label: `PI ${piNumber}`
      });
    }
  });
  
  // Sort by PI number (descending - most recent first)
  availablePIs.sort((a, b) => parseInt(b.number) - parseInt(a.number));
  
  return availablePIs;
}

/**
 * Generates HTML for the deferrals dialog
 */
function getDeferralsDialogHTML() {
  const availablePIs = getAvailablePIsForDeferrals();
  
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
          <p>Deferral presentations require an Iteration 6 tab for the selected PI.</p>
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
          .input-group:last-child { margin-bottom: 0; }
          label {
            display: block;
            color: #5f6368;
            font-size: 13px;
            margin-bottom: 6px;
            font-weight: 500;
          }
          .pi-select, .folder-input { 
            width: 100%; 
            padding: 12px; 
            font-size: 16px; 
            border: 2px solid #dadce0; 
            border-radius: 8px; 
            transition: border-color 0.2s; 
            background: white; 
          }
          .pi-select:focus, .folder-input:focus { 
            outline: none; 
            border-color: #1a73e8; 
          }
          .info-box {
            background: #fef7e0;
            border-left: 3px solid #f9ab00;
            border-radius: 4px;
            padding: 12px;
            margin-bottom: 20px;
          }
          .info-box p {
            color: #5f6368;
            font-size: 13px;
            line-height: 1.5;
            margin-bottom: 8px;
          }
          .info-box p:last-child { margin-bottom: 0; }
          .info-box strong { color: #3c4043; }
          .info-box ul {
            margin: 8px 0 0 20px;
            color: #5f6368;
            font-size: 13px;
          }
          .info-box ul li { margin-bottom: 4px; }
          .help-text {
            color: #5f6368;
            font-size: 12px;
            margin-top: 6px;
            line-height: 1.4;
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
            font-weight: 500; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            transition: all 0.2s; 
          }
          .primary-button { 
            background: #1a73e8; 
            color: white; 
          }
          .primary-button:hover { 
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
          .radio-group {
            display: flex;
            flex-direction: column;
            gap: 12px;
          }
          .radio-option {
            display: flex;
            align-items: center;
            padding: 12px;
            border: 2px solid #dadce0;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.2s;
          }
          .radio-option:hover {
            border-color: #1a73e8;
            background: #f8f9fa;
          }
          .radio-option input[type="radio"] {
            margin-right: 12px;
            width: 18px;
            height: 18px;
            cursor: pointer;
            accent-color: #1a73e8;
          }
          .radio-option label {
            cursor: pointer;
            flex: 1;
            font-size: 14px;
            color: #3c4043;
          }
          .radio-option.selected {
            border-color: #1a73e8;
            background: #e8f0fe;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Generate PI Deferrals</h2>
          <p class="subtitle">Create deferral slides from Iteration 6 data</p>
          
          <div class="section">
            <div class="section-title">Select Program Increment</div>
            <div class="input-group">
              <select id="piSelect" class="pi-select">
                <option value="">-- Select PI --</option>
                ${piOptions}
              </select>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Sort Deferrals By</div>
            <div class="radio-group">
              <div class="radio-option selected" onclick="selectSortOption(this, 'program_initiative')">
                <input type="radio" id="sort_pi" name="sortBy" value="program_initiative" checked>
                <label for="sort_pi">Program Initiative</label>
              </div>
              <div class="radio-option" onclick="selectSortOption(this, 'value_stream')">
                <input type="radio" id="sort_vs" name="sortBy" value="value_stream">
                <label for="sort_vs">Value Stream/Org</label>
              </div>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Destination Folder (Optional)</div>
            <div class="input-group">
              <label for="folderInput">Google Drive Folder ID</label>
              <input 
                type="text" 
                id="folderInput" 
                class="folder-input" 
                placeholder="Leave blank to use default APMO folder"
              >
              <p class="help-text">
                To use a specific folder, paste its ID from the Drive URL. If left blank, 
                the presentation will be saved to the default APMO folder.
              </p>
            </div>
          </div>
          
          <div class="info-box">
            <p><strong>🔴 Report Contents:</strong></p>
            <ul>
              <li>5 deferrals per slide</li>
              <li>Sorted by your selection (Program Initiative or Value Stream)</li>
              <li>Columns: Issue Key, Program Initiative, <strong>Epic Summary (clickable link to JIRA)</strong>, RAG Note, Value Stream/Org, AI Summary</li>
              <li>Only includes issues with PI Commitment = "Deferred", "Canceled", or "Traded"</li>
              <li>RAG Note may be empty for some issues</li>
            </ul>
            <p style="margin-top: 8px;"><strong>📍 Data Source:</strong> PI X - Iteration 6 tab</p>
          </div>
          
          <div class="button-container">
            <button type="button" class="secondary-button" onclick="google.script.host.close()">
              Cancel
            </button>
            <button type="button" class="primary-button" id="generateBtn" onclick="handleGenerate()">
              Generate Deferrals
            </button>
          </div>
          
          <div class="loading" id="loading">
            <div class="spinner"></div>
            <p>Creating deferral slides...</p>
          </div>
          
          <div class="error" id="error"></div>
        </div>
        
        <script>
          function selectSortOption(element, value) {
            // Remove selected class from all options
            document.querySelectorAll('.radio-option').forEach(opt => {
              opt.classList.remove('selected');
            });
            
            // Add selected class to clicked option
            element.classList.add('selected');
            
            // Check the radio button
            const radioInput = element.querySelector('input[type="radio"]');
            if (radioInput) {
              radioInput.checked = true;
            }
          }
          
          // Alternative: handle clicks on the input directly
          document.addEventListener('DOMContentLoaded', function() {
            document.querySelectorAll('.radio-option input[type="radio"]').forEach(radio => {
              radio.addEventListener('change', function() {
                document.querySelectorAll('.radio-option').forEach(opt => {
                  opt.classList.remove('selected');
                });
                this.closest('.radio-option').classList.add('selected');
              });
            });
          });
          
          function handleGenerate() {
            const piNumber = document.getElementById('piSelect').value;
            const folderId = document.getElementById('folderInput').value.trim();
            const sortBy = document.querySelector('input[name="sortBy"]:checked').value;
            
            if (!piNumber) {
              showError('Please select a Program Increment');
              return;
            }
            
            // Show loading state
            document.getElementById('loading').style.display = 'block';
            document.getElementById('generateBtn').disabled = true;
            document.getElementById('error').classList.remove('show');
            
            // Call the backend function with sort parameter
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .generateDeferralsSlideFromDialog(
                parseInt(piNumber), 
                folderId || null,
                sortBy
              );
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

// ═══════════════════════════════════════════════════════════════════════════
// GENERATION FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Generates deferral slides from dialog selection
 * Called by the dialog UI
 * 
 * @param {number} piNumber - The PI number selected in dialog
 * @param {string} destinationFolderId - Optional folder ID for saving presentation
 * @param {string} sortBy - Sort option: 'program_initiative' or 'value_stream'
 */
function generateDeferralsSlideFromDialog(piNumber, destinationFolderId, sortBy) {
  return logActivity('PI Deferrals Slide Generation', () => {
    const startTime = Date.now();
    const ISSUES_PER_SLIDE = 5;
    const TEMPLATE_SLIDE_ID = '1iT7sFMa3W4-Gbf7MxC69V-lWYrGIBCZNuY-FvTye_Vw'; // Correct APMO template
    const DEFAULT_FOLDER_ID = '1PUpzgiqNhc3pbrQqlUY1MtcKbftFLg9M'; // APMO Google Drive
    const JIRA_BASE_URL = 'https://modmedrnd.atlassian.net/browse/';
    
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    sortBy = sortBy || 'program_initiative'; // Default to program initiative if not specified
    
    console.log(`\n${'='.repeat(80)}`);
    console.log(`PI DEFERRALS GENERATION - PI ${piNumber}`);
    console.log(`Sort By: ${sortBy === 'program_initiative' ? 'Program Initiative' : 'Value Stream/Org'}`);
    console.log(`${'='.repeat(80)}\n`);
    
    ss.toast('Starting deferrals generation...', '🔴 Initializing', 5);
    
    try {
      // ═══════════════════════════════════════════════════════════════════
      // STEP 1: VALIDATE AND GET SOURCE SHEET
      // ═══════════════════════════════════════════════════════════════════
      const sourceSheetName = `PI ${piNumber} - Iteration 6`;
      const sourceSheet = ss.getSheetByName(sourceSheetName);
      
      if (!sourceSheet) {
        throw new Error(
          `Source sheet not found: ${sourceSheetName}\n\n` +
          `Deferrals require an Iteration 6 tab.\n` +
          `Please generate the Iteration 6 report first.`
        );
      }
      
      console.log(`✓ Found source sheet: ${sourceSheetName}`);
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 2: READ AND FILTER DEFERRAL DATA
      // ═══════════════════════════════════════════════════════════════════
      ss.toast('Reading deferral data from Iteration 6...', '📊 Step 1/3', 5);
      
      const dataRange = sourceSheet.getRange(4, 1, sourceSheet.getLastRow() - 3, sourceSheet.getLastColumn());
      const values = dataRange.getValues();
      const headers = values.shift();

      // Find column indexes dynamically
      // Look for PI Commitment column (flexible matching)
      let piCommitmentCol = headers.indexOf('PI Commitment');
      if (piCommitmentCol === -1) {
        // Try flexible matching for variations like "PI Commitment field"
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

      if (piCommitmentCol === -1) {
        throw new Error(
          "Required column not found: 'PI Commitment'\n\n" +
          "Looking for a column containing 'PI Commitment' in the header.\n" +
          "Available columns:\n" + headers.filter(h => h).join(', ')
        );
      }
      
      if (initiativeCol === -1) {
        throw new Error("Required column not found: 'Program Initiative'");
      }
      
      console.log(`✓ Found PI Commitment column: "${headers[piCommitmentCol]}" at index ${piCommitmentCol}`);

      // Filter rows where PI Commitment is "Deferred", "Canceled", or "Traded"
      // Using case-insensitive and trimmed comparison for robustness
      const deferralStatuses = ['deferred', 'canceled', 'traded', 'cancelled']; // Include UK spelling
      const issues = values.map(row => ({
        key: row[keyCol],
        programInitiative: row[initiativeCol],
        summary: row[summaryCol],
        piCommitment: row[piCommitmentCol],
        ragNote: row[ragNoteCol] || '', // Empty string if no RAG Note
        valueStreamOrg: row[valueStreamCol],
        aiRagNote: row[aiSummaryCol] || '' // Empty string if no AI summary
      })).filter(issue => {
        const commitment = (issue.piCommitment || '').toString().trim().toLowerCase();
        return deferralStatuses.includes(commitment);
      });
      
      console.log(`  Filtering on values: "Deferred", "Canceled", "Cancelled", "Traded" (case-insensitive)`);
      
      // Debug: show what values exist in the PI Commitment column
      const uniqueCommitments = [...new Set(values.map(row => 
        (row[piCommitmentCol] || '').toString().trim()
      ).filter(v => v))];
      console.log(`  Unique PI Commitment values found: ${uniqueCommitments.join(', ')}`);
      console.log(`  Total rows before filtering: ${values.length}`);

      if (issues.length === 0) {
        console.log('No deferrals found with PI Commitment = Deferred/Canceled/Traded');
        
        // Show what values were actually found
        const uniqueCommitments = [...new Set(values.map(row => 
          (row[piCommitmentCol] || '').toString().trim()
        ).filter(v => v))];
        
        console.log(`Values found in PI Commitment column: ${uniqueCommitments.join(', ')}`);
        
        ui.alert(
          'No Deferrals Found', 
          `No issues with PI Commitment = "Deferred", "Canceled", or "Traded" found in ${sourceSheetName}.\n\n` +
          `PI Commitment values found:\n${uniqueCommitments.map(v => `• ${v}`).join('\n')}\n\n` +
          `Nothing to generate.`,
          ui.ButtonSet.OK
        );
        return;
      }

      // Sort by selected option
      if (sortBy === 'value_stream') {
        issues.sort((a, b) => {
          const vsCompare = (a.valueStreamOrg || '').localeCompare(b.valueStreamOrg || '');
          if (vsCompare !== 0) return vsCompare;
          // Secondary sort by Program Initiative
          return (a.programInitiative || '').localeCompare(b.programInitiative || '');
        });
        console.log(`✓ Sorted ${issues.length} deferrals by Value Stream/Org (then Program Initiative)`);
      } else {
        // Default: Sort by Program Initiative
        issues.sort((a, b) => {
          const piCompare = (a.programInitiative || '').localeCompare(b.programInitiative || '');
          if (piCompare !== 0) return piCompare;
          // Secondary sort by Value Stream
          return (a.valueStreamOrg || '').localeCompare(b.valueStreamOrg || '');
        });
        console.log(`✓ Sorted ${issues.length} deferrals by Program Initiative (then Value Stream/Org)`);
      }

      // ═══════════════════════════════════════════════════════════════════
      // STEP 3: CREATE PRESENTATION
      // ═══════════════════════════════════════════════════════════════════
      ss.toast(`Generating presentation with ${issues.length} deferrals...`, '📊 Step 2/3', 10);
      
      const folderId = destinationFolderId || DEFAULT_FOLDER_ID;
      const presentationName = `PI ${piNumber} Deferrals - Iteration 6 - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')}`;
      
      const templateFile = DriveApp.getFileById(TEMPLATE_SLIDE_ID);
      const destinationFolder = DriveApp.getFolderById(folderId);
      const newFile = templateFile.makeCopy(presentationName, destinationFolder);
      const newPresentation = SlidesApp.openById(newFile.getId());
      const templateSlide = newPresentation.getSlides()[0];
      const lastRunDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');

      // Clean up template copy
      newPresentation.getSlides().slice(1).forEach(slide => slide.remove());

      // ═══════════════════════════════════════════════════════════════════
      // STEP 4: POPULATE SLIDES WITH DEFERRALS
      // ═══════════════════════════════════════════════════════════════════
      let slidesCreated = 0;
      
      for (let i = 0; i < issues.length; i += ISSUES_PER_SLIDE) {
        const issuesBatch = issues.slice(i, i + ISSUES_PER_SLIDE);
        const newSlide = newPresentation.insertSlide(newPresentation.getSlides().length, templateSlide);
        slidesCreated++;
        addDisclosureToSlide(newSlide);

        // Populate placeholders for each issue in the batch
        issuesBatch.forEach((issue, index) => {
          const rowNum = index + 1;

          // Populate the 6 data columns
          newSlide.replaceAllText(`{{issue_key_${rowNum}}}`, issue.key || 'N/A');
          newSlide.replaceAllText(`{{program_initiative_${rowNum}}}`, issue.programInitiative || 'N/A');
          newSlide.replaceAllText(`{{epic_summary_${rowNum}}}`, issue.summary || 'N/A');
          newSlide.replaceAllText(`{{rag_note_${rowNum}}}`, issue.ragNote || ''); // Empty if no RAG note
          newSlide.replaceAllText(`{{value_stream_org_${rowNum}}}`, issue.valueStreamOrg || 'N/A');
          newSlide.replaceAllText(`{{ai_rag_note_${rowNum}}}`, issue.aiRagNote || ''); // Empty if no AI summary
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
              console.warn(`  ⚠️ Could not find Epic Summary text to hyperlink for ${issue.key}: "${issue.summary.substring(0, 50)}..."`);
            }
          } catch (error) {
            console.error(`  ❌ Error adding hyperlink for ${issue.key}: ${error.message}`);
          }
        });

        newSlide.replaceAllText(`{{last_run_date}}`, lastRunDate);
        console.log(`  ✓ Created slide ${slidesCreated} with ${issuesBatch.length} issue(s)`);
      }
      
      templateSlide.remove();

      // ═══════════════════════════════════════════════════════════════════
      // STEP 5: FINALIZE AND REPORT SUCCESS
      // ═══════════════════════════════════════════════════════════════════
      const presentationUrl = newFile.getUrl();
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      console.log(`\n${'='.repeat(80)}`);
      console.log(`✅ DEFERRALS PRESENTATION COMPLETE`);
      console.log(`${'='.repeat(80)}`);
      console.log(`PI: ${piNumber}`);
      console.log(`Total Slides: ${slidesCreated}`);
      console.log(`  - ${issues.length} deferral(s) across ${slidesCreated} slide(s)`);
      console.log(`Duration: ${duration}s`);
      console.log(`URL: ${presentationUrl}`);
      console.log(`${'='.repeat(80)}\n`);

      ss.toast('✅ PI Deferrals created!', 'Success', 5);
      
      // Show clickable button popup (matching predictability report style)
      const htmlOutput = HtmlService.createHtmlOutput(
        `<div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2 style="color: #502D7F; margin-top: 0;">✅ Success!</h2>
          <p style="font-size: 16px;"><strong>PI ${piNumber} Deferrals has been created!</strong></p>
          <p style="margin: 25px 0;">
            <a href="${presentationUrl}" target="_blank" style="
              display: inline-block;
              background: #502D7F;
              color: white;
              padding: 14px 28px;
              text-decoration: none;
              border-radius: 6px;
              font-weight: bold;
              font-size: 16px;
            ">📊 Open Presentation</a>
          </p>
          <hr style="margin: 25px 0; border: none; border-top: 1px solid #ddd;">
          <div style="background: #f5f5f5; padding: 15px; border-radius: 6px;">
            <p style="font-size: 14px; margin: 5px 0;"><strong>Total Slides:</strong> ${slidesCreated}</p>
            <p style="font-size: 14px; margin: 5px 0;"><strong>Deferrals:</strong> ${issues.length}</p>
            <p style="font-size: 14px; margin: 5px 0;"><strong>Duration:</strong> ${duration} seconds</p>
          </div>
          <p style="font-size: 12px; color: #999; margin-top: 20px; text-align: center;">
            Click the button above to open your presentation in a new tab
          </p>
        </div>`
      )
        .setWidth(500)
        .setHeight(380);
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'PI Deferrals Created');
      
      return {
        url: presentationUrl,
        id: newFile.getId(),
        slideCount: slidesCreated,
        deferralCount: issues.length
      };

    } catch (error) {
      console.error(`\n❌ ERROR: ${error.message}`);
      console.error(error.stack);
      ss.toast(`❌ Error: ${error.message}`, 'Error', 10);
      ui.alert('Error', `Failed to generate deferrals:\n\n${error.message}`, ui.ButtonSet.OK);
      throw error;
    }
  }, { 
    piNumber: piNumber,
    sourceSheet: `PI ${piNumber} - Iteration 6` 
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// AI SUMMARY FUNCTION (unchanged)
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Reads all 'RAG Note' entries, summarizes them using a batch AI call,
 * and populates the 'RAG Note - AI Summary' column.
 */
function summarizeDeferrals() {
 return logActivity('PI Deferrals AI Summary Generation', () => {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetName = sheet.getName();
  if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
    ui.alert('Error', 'Please run this function on a valid report sheet.', ui.ButtonSet.OK);
    return;
  }
  
  ss.toast('Starting Deferral Note summarization...', '🔍 Analyzing Sheet', 10);

  try {
    const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
    const ragNoteCol = headers.indexOf('RAG Note') + 1;
    let aiRagSummaryCol = headers.indexOf('RAG Note - AI Summary') + 1;
    
    if (ragNoteCol === 0) {
      ui.alert('Column Not Found', "Could not find the 'RAG Note' column.", ui.ButtonSet.OK);
      return;
    }
    
    if (aiRagSummaryCol === 0) {
      aiRagSummaryCol = sheet.getLastColumn() + 1;
      sheet.getRange(4, aiRagSummaryCol).setValue('RAG Note - AI Summary');
      sheet.setColumnWidth(aiRagSummaryCol, 350);
    }

    const dataRange = sheet.getRange(5, 1, sheet.getLastRow() - 4, sheet.getLastColumn());
    const dataValues = dataRange.getValues();
    
    const ragNoteTasks = [];
    dataValues.forEach((row, index) => {
      const ragNoteText = row[ragNoteCol - 1];
      if (ragNoteText && ragNoteText.trim() !== '' && ragNoteText.trim().toLowerCase() !== 'n/a') {
        ragNoteTasks.push({ text: ragNoteText, index: index });
      }
    });

    if (ragNoteTasks.length === 0) {
      ui.alert('No Notes Found', 'There are no new RAG notes to summarize.', ui.ButtonSet.OK);
      return;
    }

    const outputRange = sheet.getRange(5, aiRagSummaryCol, dataValues.length, 1);
    const existingSummaries = outputRange.getValues();

    ss.toast(`Found ${ragNoteTasks.length} notes. Processing in batches...`, '🧠 Thinking...', 30);
    
    const ragPrompt = "Summarize the following project status note for an executive. In one sentence, identify the primary reason for deferral and its impact. If the note is positive or has no clear risk, state 'On track with no significant risks noted.'";
    const BATCH_SIZE = 10;

    for (let i = 0; i < ragNoteTasks.length; i += BATCH_SIZE) {
      const batchTasks = ragNoteTasks.slice(i, i + BATCH_SIZE);
      const batchTexts = batchTasks.map(task => task.text);
      
      const currentBatchNum = (i / BATCH_SIZE) + 1;
      const totalBatches = Math.ceil(ragNoteTasks.length / BATCH_SIZE);
      ss.toast(`Processing batch ${currentBatchNum} of ${totalBatches}...`, `Batch ${currentBatchNum}/${totalBatches}`, 10);
      
      const batchSummaries = batchSummarize(batchTexts, ragPrompt);

      if (batchSummaries.length !== batchTasks.length) {
        console.error(`Mismatched count in batch starting at index ${i}.`);
        batchTasks.forEach(task => {
          existingSummaries[task.index][0] = "Error: AI response dropped this item.";
        });
        continue;
      }
      
      batchSummaries.forEach((summary, j) => {
        const originalIndex = batchTasks[j].index;
        existingSummaries[originalIndex][0] = summary;
      });
    }

    outputRange.setValues(existingSummaries);
    outputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    ss.toast(`✅ Summarization complete! Updated ${ragNoteTasks.length} notes.`, '✅ Success', 8);

  } catch (error) {
    console.error('Deferral Summarization Error:', error);
    ui.alert('An Error Occurred', error.toString(), ui.ButtonSet.OK);
  }
 }, { sheetName: SpreadsheetApp.getActiveSheet().getName() });
}

function queryDeferralsFromJira()
{
  const JIRA_DOMAIN = 'modmedrnd.atlassian.net';
  const JIRA_EMAIL = PropertiesService.getScriptProperties().getProperty('JIRA_EMAIL');
  const JIRA_API_TOKEN = PropertiesService.getScriptProperties().getProperty('JIRA_API_TOKEN');
 
  // JQL query to find all relevant Epics.
  const JQL_QUERY = '"Program Increment[Dropdown]" = "PI 12" and "PI Commitment[Dropdown]" in (Deferred, Traded, Canceled) and issuetype = Epic and "Allocation[Dropdown]" not in (KLO, Quality) ORDER BY "cf[10050]" ASC, "cf[10049]" ASC';

  // The custom field IDs for the data you want to capture.
  const fields = [
    'customfield_10050', // Program Initiative
    'customfield_10067', // Rag Note
    'customfield_10018', // Epic Summary
    'summary',           // To get the Epic Summary
    'customfield_10046'  // The new field for Value Stream/Org
  ];

  // --- API CALL & DATA PARSING ---
  const jiraApiUrl = `https://${JIRA_DOMAIN}/rest/api/3/search/jql?jql=${encodeURIComponent(JQL_QUERY)}&fields=${fields.join(',')}&maxResults=100`;

  const options = {
    'method': 'GET',
    'headers': {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${JIRA_EMAIL}:${JIRA_API_TOKEN}`),
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };

  let response;
  try {
    response = UrlFetchApp.fetch(jiraApiUrl, options);
  } catch (e) {
    Logger.log('Jira API call failed: ' + e.message);
    return;
  }

  if (response.getResponseCode() !== 200) {
    Logger.log('Jira API returned a non-200 status code: ' + response.getResponseCode());
    Logger.log('Jira API Response: ' + response.getContentText());
    return;
  }

  const data = JSON.parse(response.getContentText());
  const issues = data.issues;

  if (!issues || issues.length === 0) {
    Logger.log('No issues found with the specified JQL.');
    return;
  }
  for(let i =0; i < issues.length; i++)
  {
      Logger.log(i);
      Logger.log(issues[i]);
      Logger.log("\n");
  }
  return issues;
}

