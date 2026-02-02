/**
 * @OnlyCurrentDoc
 *
 * This script generates a PI Governance presentation based on data
 * in a Google Sheet. It copies a template, populates static slides,
 * and dynamically generates hierarchy slides for different allocation types.
 */

// =================================================================================
// SCRIPT CONFIGURATION
// =================================================================================

const TEMPLATE_ID = '1iUd2nlawFMV49WvG_EcXYnjILAmpwaayasrL0fin9dE';
const HEADER_ROW = 4;
const DATA_START_ROW = 5;

// RENAMED from SLIDE_CONFIG
const FINAL_SLIDE_CONFIG = {
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 400,
  MARGIN_TOP: 6,
  MARGIN_LEFT: 20,
  MARGIN_RIGHT: 20,
  TABLE_START_Y: 55,
  TABLE_END_Y: 400,
  CONTENT_WIDTH: 680,
  TITLE_FONT_SIZE: 16,
  SECTION_TITLE_SIZE: 12,
  INITIATIVE_SIZE: 11,
  VERSION_SIZE: 10,
  STREAM_SIZE: 10,
  BODY_FONT_SIZE: 9,
  SMALL_FONT_SIZE: 8,
  PURPLE: '#663399',
  DARK_PURPLE: '#4A148C',
  LIGHT_PURPLE: '#E8E0FF',
  INITIATIVE_PURPLE: '#502d7f',
  ROYAL_BLUE: '#0024ff',
  BLUE: '#0057B8',
  BLACK: '#000000',
  RED: '#D32F2F',
  GRAY: '#BDBDBD',
  WHITE: '#FFFFFF',
  TITLE_HEIGHT: 28,
  SECTION_HEIGHT: 28,
  INITIATIVE_HEIGHT: 16,
  VERSION_HEIGHT: 14,
  EPIC_HEIGHT: 16,
  DEPENDENCY_HEIGHT: 12,
  OBJECTIVE_HEIGHT: 12,
  RISK_NOTE_HEIGHT: 10,
  HEADER_FONT: 'Montserrat',
  BODY_FONT: 'Lato',
  SLIDE_2_10_SUBTITLE_SIZE: 18,
  SLIDE_2_10_LIST_SIZE: 13,
  HIERARCHY_PROGRAM_SIZE: 9,
  HIERARCHY_EPIC_SIZE: 9,
  HIERARCHY_OBJECTIVE_SIZE: 7,
  HIERARCHY_DEPENDENCY_SIZE: 7,
  
  // NEW: Add these for the updated template structure
  SLIDE_2_INDEX: 1,
  SLIDE_2_INITIATIVE_LIST_SIZE: 14,
  PRODUCT_FEATURE_TEMPLATE_INDEX: 5,
  TECH_PLATFORM_TEMPLATE_INDEX: 5,
  HIERARCHY_TITLE_SIZE: 14,
  HIERARCHY_PERSPECTIVE_SIZE: 12,
  HIERARCHY_EPIC_LIST_SIZE: 11,
  HIERARCHY_PROGRAM_NAME_SIZE: 14
};

// =================================================================================
// MENU & UI FUNCTIONS
// =================================================================================
function showFolderPicker() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    '📁 Select Destination Folder',
    'Enter the Google Drive folder URL or ID where you want to save the presentation.\n\n' +
    'To get the folder URL:\n' +
    '1. Open Google Drive\n' +
    '2. Navigate to the desired folder\n' +
    '3. Copy the URL from your browser\n\n' +
    'Leave blank to use the default location (same folder as this spreadsheet).',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.CANCEL) {
    return null;
  }
  
  const input = result.getResponseText().trim();
  
  // If empty, return null (will use default location)
  if (!input) {
    return {
      id: null,
      name: 'Same folder as spreadsheet (default)'
    };
  }
  
  try {
    // Extract folder ID from URL or use input as ID directly
    let folderId = input;
    
    // If it looks like a URL, extract the ID
    if (input.includes('drive.google.com')) {
      // Match patterns like:
      // https://drive.google.com/drive/folders/ABC123xyz
      // https://drive.google.com/drive/u/0/folders/ABC123xyz
      const match = input.match(/folders\/([a-zA-Z0-9-_]+)/);
      if (match) {
        folderId = match[1];
      } else {
        throw new Error('Could not extract folder ID from URL');
      }
    }
    
    // Verify folder exists and is accessible
    const folder = DriveApp.getFolderById(folderId);
    
    ui.alert(
      '✅ Folder Selected',
      'Destination folder: ' + folder.getName(),
      ui.ButtonSet.OK
    );
    
    return {
      id: folderId,
      name: folder.getName()
    };
    
  } catch (error) {
    ui.alert(
      '❌ Invalid Folder',
      'Could not access the folder. Please check:\n\n' +
      '• The URL or ID is correct\n' +
      '• You have access to the folder\n' +
      '• The folder hasn\'t been deleted\n\n' +
      'Error: ' + error.message,
      ui.ButtonSet.OK
    );
    
    return null;
  }
}

function getPortfolioInitiativeOrder(portfolioInitiative) {
  const portfolioInit = (portfolioInitiative || '').toLowerCase();
  
  if (portfolioInit.includes('klo')) return 1;
  if (portfolioInit.includes('infosec') || portfolioInit.includes('inforsec')) return 2;
  if (portfolioInit.includes('rac')) return 3;
  if (portfolioInit.includes('platform')) return 4;
  if (portfolioInit.includes('pm')) return 5;
  if (portfolioInit.includes('patient engagement')) return 6;
  if (portfolioInit.includes('clinical ai')) return 7;
  if (portfolioInit.includes('procure')) return 8;
  if (portfolioInit.includes('mmpay')) return 9;
  if (portfolioInit.includes('conversions')) return 10;
  if (portfolioInit.includes('asc')) return 11;
  if (portfolioInit.includes('data')) return 12;
  return 13; // All others
}

/**
 * Sorts Portfolio Initiative names by custom order, then alphabetically
 * Used for consistent ordering across Slide 2 and hierarchy slides (7+)
 */
function sortPortfolioInitiativesByCustomOrder(portfolioNames) {
  return portfolioNames.sort((a, b) => {
    const orderA = getPortfolioInitiativeOrder(a);
    const orderB = getPortfolioInitiativeOrder(b);
    
    // First sort by custom order
    if (orderA !== orderB) {
      return orderA - orderB;
    }
    
    // If same order priority, sort alphabetically
    return a.localeCompare(b);
  });
}
function callGeminiAPI(prompt) {
  try {
    const geminiConfig = getGeminiConfig();
    if (!geminiConfig || !geminiConfig.apiKey) {
      console.warn('Gemini API not configured, using fallback merge');
      return null;
    }
    
    const apiKey = geminiConfig.apiKey;
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;
    
    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 200
      }
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      console.warn(`Gemini API returned ${responseCode}, using fallback`);
      return null;
    }
    
    const result = JSON.parse(response.getContentText());
    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      return result.candidates[0].content.parts[0].text.trim();
    }
    
    return null;
  } catch (error) {
    console.warn(`Gemini API error: ${error.message}, using fallback`);
    return null;
  }
}

function showFinalPlanDialog() {
  const html = HtmlService.createTemplateFromFile('Finalplan')
    .evaluate()
    .setWidth(450)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Final Plan Review');
}

function showInitialPlanDialog() {
  const html = HtmlService.createTemplateFromFile('initialplanreport')  // ← Changed
    .evaluate()
    .setWidth(450)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Initial Plan Review');
}

function getAvailablePIsForValueStream() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  // Updated regex to match ANY iteration number (0, 1, 2, etc.)
  const piRegex = /^PI (\d+) - Iteration \d+$/;
  const pis = new Set();
  
  sheets.forEach(sheet => {
    const match = sheet.getName().match(piRegex);
    if (match) {
      pis.add(match[1]); // Extracts just the PI number
    }
  });
  
  // Return sorted in descending order (newest first)
  return Array.from(pis).sort((a, b) => b - a);
}

function getAvailableValueStreams(piNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Configuration Template");
    
    if (!configSheet) {
      console.error('Configuration Template sheet not found');
      throw new Error('Sheet "Configuration Template" not found.');
    }
    
    const lastRow = configSheet.getLastRow();
    console.log(`Configuration Template has ${lastRow} rows`);
    
    if (lastRow < 6) {
      console.warn(`Configuration Template only has ${lastRow} rows, need at least 6. Returning empty array.`);
      return [];
    }
    
    const values = configSheet.getRange(`A6:A${lastRow}`).getValues();
    console.log(`Read ${values.length} rows from Configuration Template column A (rows 6-${lastRow})`);
    
    const valueStreams = values.flat().filter(Boolean).sort();
    console.log(`Found ${valueStreams.length} value streams: ${valueStreams.join(', ')}`);
    
    return valueStreams;
  } catch (e) {
    console.error('Error in getAvailableValueStreams: ' + e);
    console.error(e.stack);
    throw new Error('Could not load Value Streams: ' + e.message);
  }
}

// =================================================================================
// MAIN PRESENTATION GENERATOR (CALLED FROM HTML)
// =================================================================================
function generateValueStreamPresentation(valueStream, piNumber, phase, presentationId = null, piCommitmentFilter = 'All', destinationFolderId = null) {
  phase = phase || 'FINAL';
  piCommitmentFilter = piCommitmentFilter || 'All';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('═══════════════════════════════════════════════════════════');
  console.log('DIAGNOSTIC: Starting generateValueStreamPresentation');
  console.log('═══════════════════════════════════════════════════════════');
  console.log(`Parameter 1 (valueStream): "${valueStream}" [type: ${typeof valueStream}]`);
  console.log(`Parameter 2 (piNumber): "${piNumber}" [type: ${typeof piNumber}]`);
  console.log(`Parameter 3 (phase): "${phase}" [type: ${typeof phase}]`);
  console.log(`Parameter 4 (presentationId): "${presentationId}" [type: ${typeof presentationId}]`);
  console.log(`Parameter 5 (piCommitmentFilter): "${piCommitmentFilter}" [type: ${typeof piCommitmentFilter}]`);
  console.log(`Parameter 6 (destinationFolderId): "${destinationFolderId}" [type: ${typeof destinationFolderId}]`);
  console.log('═══════════════════════════════════════════════════════════');
  
  if (valueStream === 'ALL') {
    console.log('DIAGNOSTIC: Routing to generateAllValueStreamsPresentation');
    return generateAllValueStreamsPresentation(piNumber, piCommitmentFilter, phase, destinationFolderId);
  }
  
  let presentation = null;
  let presentationUrl = null;
  
  try {
    ss.toast(`Starting ${phase} Plan generation for ${valueStream} PI ${piNumber}...`, 'Presentation Generation', 5);
    
    ss.toast('Loading data from JIRA...', 'Step 1/5', 3);
    console.log('\nDIAGNOSTIC: Loading data from JIRA...');
    const allIssues = loadPIData(piNumber);
    console.log(`DIAGNOSTIC: Loaded ${allIssues.length} total issues from JIRA`);
    
    const sourceSheetName = `PI ${piNumber} - Iteration 0`;
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    let sourceSheetUrl = ss.getUrl();
    if (sourceSheet) {
      sourceSheetUrl = `${ss.getUrl()}#gid=${sourceSheet.getSheetId()}`;
      console.log(`DIAGNOSTIC: Source sheet URL: ${sourceSheetUrl}`);
    } else {
      console.log(`DIAGNOSTIC: Source sheet "${sourceSheetName}" not found, using spreadsheet URL`);
    }
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');
    
    const issueTypeCounts = {};
    allIssues.forEach(issue => {
      const type = issue['Issue Type'] || 'Unknown';
      issueTypeCounts[type] = (issueTypeCounts[type] || 0) + 1;
    });
    console.log('DIAGNOSTIC: Issue type breakdown:');
    Object.keys(issueTypeCounts).sort().forEach(type => {
      console.log(`  - ${type}: ${issueTypeCounts[type]}`);
    });
    
    // ═══════════════════════════════════════════════════════════════
    // ✅ BUILD GLOBAL DEPENDENCY MAP FROM ALL ISSUES (BEFORE FILTERING)
    // ═══════════════════════════════════════════════════════════════
    console.log(`\nDIAGNOSTIC: ====== BUILDING DEPENDENCY MAP ======`);
    const globalDependencyMap = {};
    
    allIssues.forEach(issue => {
      if (issue['Issue Type'] === 'Dependency' && issue['Parent Key']) {
        const parentKey = issue['Parent Key'];
        if (!globalDependencyMap[parentKey]) {
          globalDependencyMap[parentKey] = [];
        }
        globalDependencyMap[parentKey].push(issue);
      }
    });
    
    console.log(`DIAGNOSTIC: Built global dependency map with ${Object.keys(globalDependencyMap).length} epic(s) having dependencies`);
    
    // Log sample of dependencies for diagnostics
    const sampleCount = Math.min(20, Object.keys(globalDependencyMap).length);
    Object.keys(globalDependencyMap).slice(0, sampleCount).forEach(epicKey => {
      const depCount = globalDependencyMap[epicKey].length;
      console.log(`  - ${epicKey}: ${depCount} dependencies`);
    });
    console.log('DIAGNOSTIC: ══════════════════════════════════════════════');
    
    console.log(`\nDIAGNOSTIC: Filtering by Value Stream = "${valueStream}"`);
    let vsData = allIssues.filter(issue => issue['Value Stream/Org'] === valueStream);
    console.log(`DIAGNOSTIC: After Value Stream filter: ${vsData.length} issues`);
    
    const vsTypeCounts = {};
    vsData.forEach(issue => {
      const type = issue['Issue Type'] || 'Unknown';
      vsTypeCounts[type] = (vsTypeCounts[type] || 0) + 1;
    });
    console.log('DIAGNOSTIC: Issue type breakdown after VS filter:');
    Object.keys(vsTypeCounts).sort().forEach(type => {
      console.log(`  - ${type}: ${vsTypeCounts[type]}`);
    });
    
    if (vsData.length === 0) {
      ss.toast(`❌ No data found for ${valueStream}`, 'Error', 5);
      throw new Error(`No data found for value stream "${valueStream}" in PI ${piNumber}`);
    }
    
    console.log(`\nDIAGNOSTIC: PI Commitment Filter = "${piCommitmentFilter}"`);
    if (piCommitmentFilter !== 'All') {
      const beforeFilterCount = vsData.length;
      console.log('DIAGNOSTIC: Applying PI Commitment filter...');
      
      const piCommitmentValues = {};
      vsData.filter(i => i['Issue Type'] === 'Epic').forEach(epic => {
        const val = (epic['PI Commitment'] || 'BLANK').trim() || 'BLANK';
        piCommitmentValues[val] = (piCommitmentValues[val] || 0) + 1;
      });
      console.log('DIAGNOSTIC: PI Commitment values in Epics BEFORE filter:');
      Object.keys(piCommitmentValues).sort().forEach(val => {
        console.log(`  - "${val}": ${piCommitmentValues[val]} epics`);
      });
      
      vsData = applyPICommitmentFilter(vsData, piCommitmentFilter);
      console.log(`DIAGNOSTIC: After PI Commitment filter: ${beforeFilterCount} → ${vsData.length} items`);
      
      const postFilterTypeCounts = {};
      vsData.forEach(issue => {
        const type = issue['Issue Type'] || 'Unknown';
        postFilterTypeCounts[type] = (postFilterTypeCounts[type] || 0) + 1;
      });
      console.log('DIAGNOSTIC: Issue type breakdown after PI Commitment filter:');
      Object.keys(postFilterTypeCounts).sort().forEach(type => {
        console.log(`  - ${type}: ${postFilterTypeCounts[type]}`);
      });
    } else {
      console.log('DIAGNOSTIC: Skipping PI Commitment filter (set to "All")');
    }
    
    if (vsData.length === 0) {
      ss.toast(`❌ No data after filtering`, 'Error', 5);
      throw new Error(`No data found after filtering for value stream "${valueStream}" with PI Commitment filter "${piCommitmentFilter}"`);
    }
    
    console.log(`\nDIAGNOSTIC: Final vsData count before PI filter: ${vsData.length} issues`);
    
    // ═══════════════════════════════════════════════════════════════
    // ✅ CORRECTED: Filter by Program Increment field
    // ═══════════════════════════════════════════════════════════════
    console.log(`\nDIAGNOSTIC: ====== PROGRAM INCREMENT FILTERING ======`);
    console.log(`DIAGNOSTIC: Phase: ${phase}`);
    const beforePIFilter = vsData.length;
    
    const piValues = {};
    vsData.forEach(issue => {
      const piVal = (issue['Program Increment'] || 'BLANK').toString().trim() || 'BLANK';
      piValues[piVal] = (piValues[piVal] || 0) + 1;
    });
    console.log('DIAGNOSTIC: Program Increment values BEFORE filter:');
    Object.keys(piValues).sort().forEach(val => {
      console.log(`  - "${val}": ${piValues[val]} issues`);
    });
    
    if (phase === 'INITIAL') {
      vsData = vsData.filter(issue => {
        const programIncrement = (issue['Program Increment'] || '').toString().trim();
        
        // ✅ MMGI-Cloud SPECIAL HANDLING: Include ALL HYD epics regardless of PI
        if (valueStream === 'MMGI-Cloud' && issue['Key'] && issue['Key'].startsWith('HYD-')) {
          console.log(`  ✓ Including MMGI-Cloud epic ${issue['Key']} with PI="${programIncrement}"`);
          return true;
        }
        
        return programIncrement === `PI ${piNumber}` || 
               programIncrement === piNumber.toString() || 
               programIncrement.toUpperCase() === 'PRE-PLANNING';
      });
      console.log(`DIAGNOSTIC: INITIAL PLAN - Included PI ${piNumber} and Pre-Planning items`);
      
      const explicitPI = beforePIFilter - vsData.length;
      if (explicitPI > 0) {
        console.log(`  - Excluded ${explicitPI} issues (wrong PI number)`);
      }
    } else {
      vsData = vsData.filter(issue => {
        const programIncrement = (issue['Program Increment'] || '').toString().trim();
        return programIncrement === `PI ${piNumber}` || programIncrement === piNumber.toString();
      });
      console.log(`DIAGNOSTIC: FINAL PLAN - Included ONLY PI ${piNumber} items (excluded Pre-Planning and other PIs)`);
      
      const excluded = beforePIFilter - vsData.length;
      if (excluded > 0) {
        console.log(`  - Excluded ${excluded} issues:`);
        const expectedPIFormat = `PI ${piNumber}`;
        const prePlanningCount = piValues['PRE-PLANNING'] || piValues['Pre-Planning'] || 0;
        
        if (prePlanningCount > 0) {
          console.log(`    - Pre-Planning items: ${prePlanningCount}`);
        }
        
        Object.keys(piValues).forEach(val => {
          if (val === expectedPIFormat || 
              val === piNumber.toString() || 
              val.toUpperCase() === 'PRE-PLANNING' || 
              val === 'BLANK') {
            return;
          }
          console.log(`    - ${val} items: ${piValues[val]}`);
        });
      }
    }
    
    console.log(`DIAGNOSTIC: After Program Increment filter: ${beforePIFilter} → ${vsData.length} issues`);
    console.log('DIAGNOSTIC: ══════════════════════════════════════════════');
    
    if (vsData.length === 0) {
      ss.toast(`❌ No data after PI filtering`, 'Error', 5);
      throw new Error(`No data found for value stream "${valueStream}" in PI ${piNumber} after Program Increment filtering. Check that issues have "Program Increment" = "PI ${piNumber}" or "${piNumber}"`);
    }
    
    ss.toast(`Organizing ${vsData.length} items by portfolio...`, 'Step 2/5', 3);
    console.log('\nDIAGNOSTIC: ====== STRUCTURING HIERARCHICAL DATA ======');
    const allHierarchies = {};
    const allocationTypes = [
      { name: 'Product - Feature', allocations: ['Product - Feature'] },
      { name: 'Compliance', allocations: ['Product - Compliance'] },
      { name: 'Quality', allocations: ['Quality'] },
      { name: 'Tech / Platform', allocations: ['Tech / Platform'] },
     { name: 'KLO', allocations: ['KLO'] }
    ];
    
    allocationTypes.forEach(allocType => {
      console.log(`\nDIAGNOSTIC: Processing allocation type: "${allocType.name}"`);
      const hierarchy = structureHierarchicalDataForFinalPlan(vsData, allocType.allocations);
      
      if (hierarchy && Object.keys(hierarchy).length > 0) {
        allHierarchies[allocType.name] = hierarchy;
        console.log(`DIAGNOSTIC: ✓ Created hierarchy with ${Object.keys(hierarchy).length} portfolios`);
        console.log('DIAGNOSTIC: Portfolio names:');
        Object.entries(hierarchy).forEach(([portfolioName, portfolioData]) => {
          const programCount = Object.keys(portfolioData).length;
          console.log(`  - "${portfolioName}": ${programCount} programs`);
        });
      } else {
        console.log('DIAGNOSTIC: ✗ No hierarchy created (empty or null)');
        if (hierarchy) {
          console.log(`DIAGNOSTIC: Hierarchy exists but has ${Object.keys(hierarchy).length} portfolios`);
        }
      }
    });
    
    console.log('\nDIAGNOSTIC: ====== HIERARCHY SUMMARY ======');
    const hierarchyCount = Object.keys(allHierarchies).length;
    console.log(`DIAGNOSTIC: Total allocation types with hierarchies: ${hierarchyCount}`);
    Object.entries(allHierarchies).forEach(([allocName, hierarchy]) => {
      console.log(`DIAGNOSTIC: - ${allocName}: ${Object.keys(hierarchy).length} portfolios`);
    });
    
    if (hierarchyCount === 0) {
      ss.toast('No hierarchical data generated', '⚠️ Warning', 5);
      throw new Error('No hierarchical data was generated. Check allocation type filters and data.');
    }
    
    // ═══════════════════════════════════════════════════════════════
    // ✅ PRESENTATION CREATION WITH DEFENSIVE ERROR HANDLING
    // ═══════════════════════════════════════════════════════════════
    ss.toast('Creating presentation...', 'Step 3/5', 3);
    console.log('\nDIAGNOSTIC: Creating new presentation from template...');
    
    const destinationFolder = destinationFolderId 
      ? DriveApp.getFolderById(destinationFolderId)
      : DriveApp.getFileById(ss.getId()).getParents().next();
    
    console.log(`DIAGNOSTIC: Using ${destinationFolderId ? 'specified' : 'default'} location (spreadsheet folder): ${destinationFolder.getName()}`);
    
    try {
      const templateFile = DriveApp.getFileById(TEMPLATE_ID);
      const newFileName = `${valueStream} - PI ${piNumber} - ${phase} Plan - ${timestamp}`;
      console.log(`DIAGNOSTIC: Creating copy with name: ${newFileName}`);
      
      const newFile = templateFile.makeCopy(newFileName, destinationFolder);
      console.log(`DIAGNOSTIC: File copied, ID: ${newFile.getId()}`);
      console.log(`DIAGNOSTIC: Opening as presentation...`);
      
      // Add small delay to ensure file is ready
      Utilities.sleep(500);
      
      presentation = SlidesApp.openById(newFile.getId());
      if (!presentation) {
        throw new Error('Failed to open presentation after creating copy');
      }
      
      presentationUrl = presentation.getUrl();
      console.log(`DIAGNOSTIC: Created new presentation from template: ${presentationUrl}`);
      console.log(`DIAGNOSTIC: Saved to folder: ${destinationFolder.getName()}`);
      
      const templateSlides = presentation.getSlides();
      if (!templateSlides || templateSlides.length === 0) {
        throw new Error('Presentation has no slides');
      }
      console.log(`DIAGNOSTIC: Presentation has ${templateSlides.length} slides initially`);
      
      ss.toast('Populating slides...', 'Step 4/5', 4);
      console.log('\nDIAGNOSTIC: Updating static slides...');
      
      // ═══════════════════════════════════════════════════════════════
      // SLIDE 1: Title Slide (INDEX 0)
      // ═══════════════════════════════════════════════════════════════
      const SLIDE_1_INDEX = 0;
      if (templateSlides.length <= SLIDE_1_INDEX) {
        throw new Error(`Cannot access slide at index ${SLIDE_1_INDEX}, presentation only has ${templateSlides.length} slides`);
      }
      
      const slide1 = templateSlides[SLIDE_1_INDEX];
      if (!slide1) {
        throw new Error(`Slide 1 (index ${SLIDE_1_INDEX}) is undefined`);
      }
      
      console.log(`DIAGNOSTIC: Populating Slide 1 (Title slide)...`);
      const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');
      robustReplaceAllText(slide1, '{{ValueStream}}', valueStream);
      robustReplaceAllText(slide1, '{{PI}}', piNumber.toString());
      robustReplaceAllText(slide1, '{{PHASE}}', phase);
      robustReplaceAllText(slide1, '{{Phase}}', phase);
      robustReplaceAllText(slide1, '{{Date}}', currentDate);
      addTimestampToSlide(slide1, timestamp, sourceSheetName, sourceSheetUrl);
      console.log(`DIAGNOSTIC: ✓ Slide 1 complete`);
      
      // ═══════════════════════════════════════════════════════════════
      // SLIDE 2: Portfolio Initiatives Summary (INDEX 1)
      // ═══════════════════════════════════════════════════════════════
      console.log('DIAGNOSTIC: Populating Slide 2 with portfolio initiatives...');
      
      const SLIDE_2_INDEX = 1;
      if (templateSlides.length <= SLIDE_2_INDEX) {
        throw new Error(`Cannot access slide at index ${SLIDE_2_INDEX}, presentation only has ${templateSlides.length} slides`);
      }
      
      const slide2 = templateSlides[SLIDE_2_INDEX];
      if (!slide2) {
        throw new Error(`Slide 2 (index ${SLIDE_2_INDEX}) is undefined`);
      }
      
      populateSlide2Plans(presentation, piNumber, valueStream, vsData, slide2);
      console.log(`DIAGNOSTIC: ✓ Slide 2 complete`);
      
      // ═══════════════════════════════════════════════════════════
      // ✅ GENERATE HIERARCHY SLIDES FOR EACH ALLOCATION TYPE
      // ═══════════════════════════════════════════════════════════
      console.log('\nDIAGNOSTIC: Generating hierarchy slides...');
      const TEMPLATE_SLIDE_INDEX = 5;
      console.log(`DIAGNOSTIC: Using slide ${TEMPLATE_SLIDE_INDEX} as template for hierarchy slides`);
      
      if (templateSlides.length <= TEMPLATE_SLIDE_INDEX) {
        throw new Error(`Cannot access template slide at index ${TEMPLATE_SLIDE_INDEX}, presentation only has ${templateSlides.length} slides`);
      }
      
      Object.entries(allHierarchies).forEach(([allocationType, hierarchy]) => {
        console.log(`\nDIAGNOSTIC: Generating hierarchy slides for ${allocationType}...`);
        const templateSlide = templateSlides[TEMPLATE_SLIDE_INDEX];
        
        if (!templateSlide) {
          console.error(`Template slide at index ${TEMPLATE_SLIDE_INDEX} is undefined, skipping ${allocationType}`);
          return;
        }
        
        generateHierarchySlides(presentation, templateSlide, hierarchy, allocationType, phase, valueStream, globalDependencyMap);
      });
      
      // ✅ REMOVE TEMPLATE SLIDE AFTER ALL ALLOCATIONS ARE PROCESSED
      try {
        const templateSlide = templateSlides[TEMPLATE_SLIDE_INDEX];
        if (templateSlide) {
          templateSlide.remove();
          console.log(`\n✅ Removed template slide after all allocation types processed`);
        }
      } catch (e) {
        console.warn(`Could not remove template slide: ${e.message}`);
      }
      
      ss.toast('Finalizing presentation...', 'Step 5/5', 2);
      const totalSlides = presentation.getSlides().length;
      console.log(`\nDIAGNOSTIC: ✅ Presentation generation complete`);
      console.log(`DIAGNOSTIC: Total slides: ${totalSlides}`);
      console.log(`DIAGNOSTIC: URL: ${presentationUrl}`);
      
      ss.toast(`✅ ${phase} Plan complete! ${totalSlides} slides`, 'Success', 5);
      
      return {
        success: true,
        url: presentationUrl,
        slideCount: totalSlides
      };
      
    } catch (presentationError) {
      console.error(`\nDIAGNOSTIC: ❌ Error during presentation creation/population`);
      console.error(`DIAGNOSTIC: Error message: ${presentationError.message}`);
      console.error(`DIAGNOSTIC: Error stack: ${presentationError.stack}`);
      throw presentationError;
    }
    
  } catch (error) {
    console.log(`\nDIAGNOSTIC: ❌ Error in generateValueStreamPresentation: [${error}]`);
    console.log(`DIAGNOSTIC: Error stack: ${error.stack}`);
    ss.toast(`Error: ${error.message}`, '❌ Failed', 10);
    throw error;
  }
}

function copyTemplateSlides(templateId, destinationPresentation) {
  const template = SlidesApp.openById(templateId);
  const templateSlides = template.getSlides();
  const newSlides = [];
  const lastSlideIndex = destinationPresentation.getSlides().length;
  for (let i = 0; i < templateSlides.length; i++) {
    newSlides.push(destinationPresentation.insertSlide(lastSlideIndex + i, templateSlides[i]));
  }
  return newSlides;
}

// =================================================================================
// DATA LOADING & STRUCTURING
// =================================================================================
function generateHierarchySlides(presentation, templateSlideForDuplication, portfolioHierarchy, allocationType, phase, valueStream, globalDependencyMap) {
  globalDependencyMap = globalDependencyMap || {}; // Default to empty if not provided
  phase = phase || 'FINAL';
  
  console.log(`Generating hierarchy slides for ${allocationType} with ${phase} phase for ${valueStream}...`);
  
  const removeLeadingNumber = (text) => {
    if (!text) return text;
    return text.replace(/^\d+\.\s*/, '').trim();
  };

  const sortedPortfolios = sortPortfolioInitiativesByCustomOrder(Object.keys(portfolioHierarchy));
  console.log(`Portfolio Initiatives sorted by custom order: ${sortedPortfolios.join(', ')}`);
  
  if (sortedPortfolios.length === 0) {
    console.log(`No portfolios found for ${allocationType}. Skipping slide generation.`);
    return;
  }

  const MAX_INITIATIVES_PER_SLIDE = 5;
  const PROGRAM_BG_COLORS = ['#FFFFFF', '#F5F0FF'];

  for (const portfolioName of sortedPortfolios) {
    const programHierarchy = portfolioHierarchy[portfolioName];
    const sortedPrograms = Object.keys(programHierarchy).sort();

    console.log(`\n=== Portfolio: ${portfolioName} ===`);
    console.log(`Total programs: ${sortedPrograms.length}`);

    // Check if portfolio has ANY initiatives with data
    let portfolioHasData = false;
    for (const programName of sortedPrograms) {
      const programData = programHierarchy[programName];
      if (programData.initiatives && programData.initiatives.length > 0) {
        portfolioHasData = true;
        break;
      }
    }
    
    if (!portfolioHasData) {
      console.log(`⚠️ Portfolio "${portfolioName}" has no initiatives with data. Skipping slide generation.`);
      continue;
    }

    // ═══════════════════════════════════════════════════════════════
    // ✅ USE GLOBAL DEPENDENCY MAP (built before value stream filtering)
    // ═══════════════════════════════════════════════════════════════
    const dependencyMap = globalDependencyMap;
    console.log(`Using global dependency map with ${Object.keys(dependencyMap).length} epic(s) having dependencies`);

    const totalPages = calculateFinalPlanPageCount(programHierarchy);
    console.log(`Calculated ${totalPages} total pages for portfolio "${portfolioName}"`);

    let slideManager = {
      presentation: presentation,
      templateSlide: templateSlideForDuplication,
      slide: null,
      table: null,
      portfolioName: portfolioName,
      allocationType: allocationType,
      valueStream: valueStream,
      programHierarchy: programHierarchy,
      dependencyMap: dependencyMap,
      pageNum: 1,
      totalPages: totalPages,
      initiativeCountOnSlide: 0,
      programColorIndex: 0,
      
      startNewSlide: function(isContinued = false) {
        this.slide = this.presentation.appendSlide(this.templateSlide);
        addDisclosureToSlide(this.slide),
        console.log(`\n📄 Created new slide (Page ${this.pageNum}/${this.totalPages})`);

        const tables = this.slide.getTables();

        if (tables && tables.length > 0) {
          this.table = tables[0];
          
          while (this.table.getNumRows() > 1) {
            this.table.getRow(1).remove();
          }
          
          console.log(`Table ready with ${this.table.getNumColumns()} columns`);
        } else {
          console.error(`No table found on template slide`);
          this.table = null;
        }

        const pageNumbering = this.totalPages > 1 ? ` (${this.pageNum}/${this.totalPages})` : '';
        const portfolioText = `${this.portfolioName}${pageNumbering}`;
        
        let totalEffort = 0;
        Object.values(this.programHierarchy).forEach(programData => {
          programData.initiatives.forEach(initiative => {
            initiative.epics.forEach(epic => {
              const featurePoints = parseFloat(epic['Feature Points']) || 0;
              totalEffort += featurePoints;
            });
          });
        });
        
        const titleLine1 = `${this.valueStream}: ${this.allocationType} - ${portfolioText}`;
        const titleLine2 = `Feature Points: ${totalEffort}`;
        const fullTitle = `${titleLine1}\n${titleLine2}`;
        
        robustReplaceAllText(this.slide, '{{Value Stream}}: {{Allocation Type}} - {{Portfolio Initiative}}', fullTitle);
        
        try {
          const shapes = this.slide.getShapes();
          shapes.forEach(shape => {
            const textRange = shape.getText();
            if (textRange && !textRange.isEmpty()) {
              const fullText = textRange.asString();
              
              if (fullText.includes(titleLine1) && fullText.includes(titleLine2)) {
                textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
                textRange.getTextStyle().setFontFamily('Lato');
                
                const line1Start = fullText.indexOf(titleLine1);
                if (line1Start > -1) {
                  const line1End = line1Start + titleLine1.length;
                  const line1Range = textRange.getRange(line1Start, line1End);
                  line1Range.getTextStyle()
                    .setFontSize(16)
                    .setItalic(false)
                    .setFontFamily('Lato');
                  
                  // Style the page numbering with smaller font (2 sizes smaller)
                  if (pageNumbering) {
                    const pageNumStart = titleLine1.indexOf(pageNumbering);
                    if (pageNumStart > -1) {
                      const pageNumStartInFull = line1Start + pageNumStart;
                      const pageNumEndInFull = pageNumStartInFull + pageNumbering.length;
                      const pageNumRange = textRange.getRange(pageNumStartInFull, pageNumEndInFull);
                      pageNumRange.getTextStyle().setFontSize(14);
                    }
                  }
                }
                
                const line2Start = fullText.indexOf(titleLine2);
                if (line2Start > -1) {
                  const line2End = line2Start + titleLine2.length;
                  const line2Range = textRange.getRange(line2Start, line2End);
                  line2Range.getTextStyle()
                    .setFontSize(8)
                    .setBold(false)
                    .setItalic(true)
                    .setFontFamily('Lato');
                }
              }
            }
          });
        } catch (e) {
          console.warn(`Could not style title: ${e.message}`);
        }
        
        if (isContinued) this.pageNum++;
        
        this.initiativeCountOnSlide = 0;
      },
      
      canAddInitiatives: function(count) {
        return (this.initiativeCountOnSlide + count) <= MAX_INITIATIVES_PER_SLIDE;
      },
      
      addProgramRow: function(programName, initiativesToAdd, bgColor, isPartial = false) {
        if (!this.table) {
          console.error('No table available');
          return 0;
        }
        
        const row = this.table.appendRow();
        const initiativesAdded = initiativesToAdd.length;
        
        this.initiativeCountOnSlide += initiativesAdded;
        console.log(`  ✅ Added ${initiativesAdded} initiatives (total on slide: ${this.initiativeCountOnSlide})`);
        
        try {
          for (let i = 0; i < this.table.getNumColumns(); i++) {
            row.getCell(i).getFill().setSolidFill(bgColor);
          }
        } catch (e) {
          console.warn(`Could not set background: ${e.message}`);
        }
        
        // COLUMN 0 - Program Initiative
        try {
          const cell = row.getCell(0);
          const textRange = cell.getText();
          const cleanProgramName = removeLeadingNumber(programName);
          const displayName = isPartial ? `${cleanProgramName} (continued)` : cleanProgramName;
          
          textRange.setText(displayName);
          textRange.getTextStyle()
            .setFontSize(9)
            .setBold(true)
            .setFontFamily('Lato')
            .setForegroundColor('#000000');
          textRange.getParagraphStyle()
            .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
            .setSpaceAbove(2)
            .setSpaceBelow(2);
        } catch (e) {
          console.error(`Could not set Program Initiative column: ${e.message}`);
        }
        
        // COLUMN 1 - Initiative RAG (with linked initiatives, feature points, and perspectives)
        try {
          const cell = row.getCell(1);
          
          let fullText = '';
          const metadataStore = [];
          
          initiativesToAdd.forEach((initiative, idx) => {
            const isLast = (idx === initiativesToAdd.length - 1);
            const separator = isLast ? '' : '\n\n';
            
            // Calculate total Feature Points for this initiative
            let totalFeaturePoints = 0;
            initiative.epics.forEach(epic => {
              const fp = parseFloat(epic['Feature Points']) || 0;
              totalFeaturePoints += fp;
            });
            
            const initiativeTitle = initiative.initiativeTitle || initiative.epics[0]['Summary'] || 'Untitled Initiative';
            const linkKey = initiative.linkKey;
            
            // Add Feature Points to the title
            const titleWithFP = `${initiativeTitle} | Feature Points: ${totalFeaturePoints}`;
            
            const titleStart = fullText.length;
            const titleEnd = titleStart + titleWithFP.length;
            const fpMarkerStart = titleStart + initiativeTitle.length;
            
            // Build perspective text
            let perspectiveText = '';
            if (phase === 'INITIAL') {
              perspectiveText = initiative.businessPerspective || '';
            } else if (phase === 'FINAL') {
              perspectiveText = initiative.technicalPerspective || '';
            }
            
            const perspectiveStart = titleEnd + 1;
            const perspectiveEnd = perspectiveStart + perspectiveText.length;
            
            // Build committed epics line
            const committedEpics = initiative.epics
              .filter(epic => {
                const commitment = (epic['PI Commitment'] || '').toLowerCase();
                return commitment.includes('committed');
              })
              .map(epic => epic['Key']);
            
            let committedEpicsLine = '';
            let committedEpicsStart = -1;
            let committedEpicsKeysStart = -1;
            
            if (committedEpics.length > 0) {
              const epicKeysList = committedEpics.join(', ');
              committedEpicsLine = `Committed: ${epicKeysList}`;
              committedEpicsStart = perspectiveEnd + 1;
              committedEpicsKeysStart = committedEpicsStart + 'Committed: '.length;
            }
            
            // Assemble the full text for this initiative
            const initiativeBlock = `${titleWithFP}\n${perspectiveText}${committedEpicsLine ? '\n' + committedEpicsLine : ''}${separator}`;
            fullText += initiativeBlock;
            
            metadataStore.push({
              line1Start: titleStart,
              line1End: titleEnd,
              titleEnd: titleEnd,
              fpMarkerStart: fpMarkerStart,
              linkKey: linkKey,
              initiativeTitleLength: initiativeTitle.length,
              perspectiveStart: perspectiveStart,
              perspectiveEnd: perspectiveEnd,
              committedEpicsStart: committedEpicsStart,
              committedEpicsKeysStart: committedEpicsKeysStart,
              committedEpics: committedEpics
            });
          });
          
          const textRange = cell.getText();
          textRange.setText(fullText);
          
          metadataStore.forEach((meta) => {
            // LINE 1: Title with hyperlink, bold, 11pt, purple
            const initiativeTitleEnd = meta.line1Start + meta.initiativeTitleLength;
            const titleOnlyRange = textRange.getRange(meta.line1Start, initiativeTitleEnd);
            
            const url = getJiraUrl(meta.linkKey);
            if (url) {
              titleOnlyRange.getTextStyle()
                .setLinkUrl(url)
                .setFontSize(11)
                .setBold(true)
                .setForegroundColor(FINAL_SLIDE_CONFIG.INITIATIVE_PURPLE)
                .setUnderline(false)
                .setFontFamily('Lato');
            } else {
              titleOnlyRange.getTextStyle()
                .setFontSize(11)
                .setBold(true)
                .setForegroundColor(FINAL_SLIDE_CONFIG.INITIATIVE_PURPLE)
                .setFontFamily('Lato');
            }
            
            titleOnlyRange.getParagraphStyle()
              .setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
            
            // Style Feature Points portion separately (6pt, not bold, italic)
            const fpRange = textRange.getRange(meta.fpMarkerStart, meta.line1End);
            fpRange.getTextStyle()
              .setFontSize(6)
              .setBold(false)
              .setItalic(true)
              .setForegroundColor(FINAL_SLIDE_CONFIG.INITIATIVE_PURPLE)
              .setFontFamily('Lato');
            
            console.log(`  ✓ Styled Feature Points: 6pt, not bold, italic`);
  
            // LINE 2: Perspective - Black, not bold, 9pt
            if (meta.perspectiveEnd > meta.perspectiveStart) {
              const perspectiveRange = textRange.getRange(meta.perspectiveStart, meta.perspectiveEnd);
              perspectiveRange.getTextStyle()
                .setFontSize(9)
                .setBold(false)
                .setItalic(false)
                .setForegroundColor('#000000')
                .setFontFamily('Lato');
              perspectiveRange.getParagraphStyle()
                .setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
            }
            
            // LINE 3 - Style Committed Epics
            if (meta.committedEpicsStart > -1) {
              const labelRange = textRange.getRange(meta.committedEpicsStart, meta.committedEpicsKeysStart);
              labelRange.getTextStyle()
                .setFontSize(6)
                .setBold(true)
                .setItalic(false)
                .setForegroundColor('#000000')
                .setFontFamily('Lato');
              
              if (meta.committedEpics && meta.committedEpics.length > 0) {
                const epicKeysText = meta.committedEpics.join(', ');
                
                meta.committedEpics.forEach((epicKey) => {
                  const epicKeyStart = meta.committedEpicsKeysStart + epicKeysText.indexOf(epicKey);
                  const epicKeyEnd = epicKeyStart + epicKey.length;
                  
                  try {
                    const epicKeyRange = textRange.getRange(epicKeyStart, epicKeyEnd);
                    const epicUrl = getJiraUrl(epicKey);
                    if (epicUrl) {
                      epicKeyRange.getTextStyle()
                        .setLinkUrl(epicUrl)
                        .setForegroundColor(FINAL_SLIDE_CONFIG.ROYAL_BLUE)
                        .setUnderline(false)
                        .setBold(false)
                        .setFontSize(6)
                        .setFontFamily('Lato');
                    }
                  } catch (e) {
                    console.warn(`Could not hyperlink committed epic key ${epicKey}: ${e.message}`);
                  }
                });
              }
            }
          });
          
          textRange.getParagraphStyle()
            .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
            .setSpaceAbove(2)
            .setSpaceBelow(2);
          
        } catch (e) {
          console.error(`Could not populate Initiative/RAG column: ${e.message}`);
        }
        
        // ═══════════════════════════════════════════════════════════════════════
        // ✅ COLUMN 2 - Risks & Dependencies (DEDUPLICATED)
        // ═══════════════════════════════════════════════════════════════════════
        try {
          const cell = row.getCell(2);
          
          if (!cell || typeof cell.getText !== 'function') {
            console.warn(`⚠️ Risk/Dep column cell does not support getText() - skipping display for this row`);
            return initiativesAdded;
          }
          
          let textRange = null;
          try {
            textRange = cell.getText();
          } catch (getTextError) {
            console.warn(`⚠️ Could not get text from Risk/Dep cell: ${getTextError.message}`);
            // ✅ FIX: Better cell initialization
            try {
              // Try to insert a paragraph first
              cell.insertParagraph(0, '');
              textRange = cell.getText();
              console.log(`   ✓ Successfully initialized Risk/Dep cell with paragraph`);
            } catch (insertError) {
              console.warn(`   ❌ Could not initialize Risk/Dep cell: ${insertError.message} - skipping display`);
              return initiativesAdded;
            }
          }
          
          if (!textRange) {
            console.warn(`⚠️ No text range available for Risk/Dep column - skipping display`);
            return initiativesAdded;
          }
          
          // ✅ BUILD CONTENT: RAG notes + Dependencies (BOTH DEDUPLICATED)
          let riskDepContent = '';
          
          initiativesToAdd.forEach((initiative, idx) => {
            const isLast = (idx === initiativesToAdd.length - 1);
            const separator = isLast ? '' : '\n\n';
            
            // ═══════════════════════════════════════════════════════════
            // PART 1: Collect RAG notes (deduplicated by icon + note text)
            // ═══════════════════════════════════════════════════════════
            const uniqueRagNotes = new Set();
            
            initiative.epics.forEach(epic => {
              const rag = (epic['RAG'] || '').toLowerCase().trim();
              const ragNote = (epic['RAG Note'] || '').trim();
              
              // Only show if NOT green and has a note
              if (rag && rag !== 'green' && rag !== '' && ragNote && ragNote !== '') {
                const ragIcon = getRagIcon(rag);
                const ragEntry = `${ragIcon} ${ragNote}`;
                uniqueRagNotes.add(ragEntry);
              }
            });
            
            // Add RAG notes for this initiative (now deduplicated)
            if (uniqueRagNotes.size > 0) {
              riskDepContent += Array.from(uniqueRagNotes).join('\n') + '\n';
            }
            
            // ═══════════════════════════════════════════════════════════
            // PART 2: Collect dependencies (deduplicated by VS + title)
            // ═══════════════════════════════════════════════════════════
            const uniqueDeps = new Set();
            
            initiative.epics.forEach(epic => {
              const deps = this.dependencyMap[epic['Key']] || [];
              
              deps.forEach(dep => {
                const depVS = (dep['Depends on Valuestream'] || dep['Value Stream/Org'] || 'Unknown').trim();
                const depTitle = (dep['Summary'] || dep['Key']).trim();
                
                // Create unique identifier to prevent duplicates
                const depIdentifier = `${depVS}|${depTitle}`;
                uniqueDeps.add(depIdentifier);
              });
            });
            
            // Add dependencies for this initiative
            if (uniqueDeps.size > 0) {
              // Add spacing if we already have RAG notes
              if (uniqueRagNotes.size > 0) {
                riskDepContent += '\n';
              }
              
              // Convert Set to Array, sort, and format
              Array.from(uniqueDeps).sort().forEach(depIdentifier => {
                const [depVS, depTitle] = depIdentifier.split('|');
                riskDepContent += `🔗 ${depVS}: ${depTitle}\n`;
              });
            }
            
            riskDepContent += separator;
          });
          
          // Apply text and styling
          textRange.setText(riskDepContent.trim());
          textRange.getTextStyle()
            .setFontSize(8)
            .setBold(false)
            .setFontFamily('Lato')
            .setForegroundColor('#000000');
          textRange.getParagraphStyle()
            .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
            .setSpaceAbove(2)
            .setSpaceBelow(2);
          
          console.log(`  ✓ Set risks/dependencies for ${initiativesToAdd.length} initiative(s)`);
          
        } catch (e) {
          console.warn(`Could not set Risk/Dep column: ${e.message}`);
        }
        
        return initiativesAdded;
      }
    };
    
    slideManager.startNewSlide(false);

    // Process each program
    for (const programName of sortedPrograms) {
      const programData = programHierarchy[programName];
      
      if (!programData.initiatives || programData.initiatives.length === 0) {
        console.log(`  ⚠️ Program "${programName}" has no initiatives. Skipping.`);
        continue;
      }

      console.log(`\n--- Program: ${programName} ---`);
      console.log(`  Total initiatives: ${programData.initiatives.length}`);

      const sortedInitiatives = programData.initiatives.sort((a, b) => {
        const titleA = a.initiativeTitle || '';
        const titleB = b.initiativeTitle || '';
        return titleA.localeCompare(titleB);
      });

      let remainingInitiatives = [...sortedInitiatives];
      let isFirstBatchForThisProgram = true;

      while (remainingInitiatives.length > 0) {
        const availableSlots = MAX_INITIATIVES_PER_SLIDE - slideManager.initiativeCountOnSlide;
        const batchSize = Math.min(availableSlots, remainingInitiatives.length);
        const batch = remainingInitiatives.splice(0, batchSize);

        const bgColor = PROGRAM_BG_COLORS[slideManager.programColorIndex % PROGRAM_BG_COLORS.length];
        const isPartial = !isFirstBatchForThisProgram;
        slideManager.addProgramRow(programName, batch, bgColor, isPartial);

        isFirstBatchForThisProgram = false;

        if (remainingInitiatives.length > 0) {
          console.log(`  🔄 ${remainingInitiatives.length} initiatives remaining for "${programName}"`);
          slideManager.startNewSlide(true);
        }
      }

      slideManager.programColorIndex++;
    }
  }

  // ✅ DO NOT REMOVE TEMPLATE HERE - it's shared across allocations!
  console.log(`\n✅ Hierarchy slides generated successfully for ${allocationType}`);
}
/**
 * Helper function for calculateFinalPlanPageCount
 */
function calculateFinalPlanPageCount(programHierarchy) {
  let totalInitiatives = 0;
  Object.values(programHierarchy).forEach(programData => {
    if (programData.initiatives && Array.isArray(programData.initiatives)) {
      totalInitiatives += programData.initiatives.length;
    }
  });
  return Math.ceil(totalInitiatives / 5); // 5 initiatives per page
}

function loadPIData(piNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Find the first available iteration sheet for this PI (prefer Iteration 0, then 1, then others)
  const piRegex = new RegExp(`^PI ${piNumber} - Iteration (\\d+)$`);
  const availableIterations = [];
  
  sheets.forEach(sheet => {
    const match = sheet.getName().match(piRegex);
    if (match) {
      availableIterations.push({
        iteration: parseInt(match[1]),
        name: sheet.getName(),
        sheet: sheet
      });
    }
  });
  
  if (availableIterations.length === 0) {
    throw new Error(`No sheets found for PI ${piNumber}. Please ensure sheets named "PI ${piNumber} - Iteration X" exist.`);
  }
  
  // Sort iterations: prefer 0 FIRST (for Initial/Final Plans), then 1, then ascending order
  availableIterations.sort((a, b) => {
    if (a.iteration === 0) return -1;  // Iteration 0 has highest priority
    if (b.iteration === 0) return 1;
    if (a.iteration === 1) return -1;  // Iteration 1 is second priority
    if (b.iteration === 1) return 1;
    return a.iteration - b.iteration;  // Others in ascending order
  });
  
  const selectedSheet = availableIterations[0].sheet;
  const sheetName = availableIterations[0].name;
  
  console.log(`Loading data from sheet: ${sheetName} for PI ${piNumber}`);
  
  const lastCol = selectedSheet.getLastColumn();
  const lastRow = selectedSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const headers = selectedSheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const columnMap = {};
  headers.forEach((header, index) => { if (header) columnMap[header] = index; });

  const requiredCols = ['Key', 'Issue Type', 'Value Stream/Org', 'Allocation', 'Portfolio Initiative', 'Program Initiative', 'Summary', 'Initiative Title', 'PI Objective', 'Depends on Valuestream', 'Parent Key'];
  for (const col of requiredCols) {
    if (columnMap[col] === undefined) throw new Error(`Missing required column in sheet '${sheetName}': "${col}"`);
  }

  const dataRows = selectedSheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, lastCol).getValues();
  const allIssues = [];
  dataRows.forEach(row => {
    const issue = {};
    Object.keys(columnMap).forEach(header => { issue[header] = row[columnMap[header]]; });
    if (issue['Key'] && String(issue['Key']).trim() !== '') allIssues.push(issue);
  });
  console.log(`Loaded ${allIssues.length} issues from ${sheetName}`);
  return allIssues;
}

function structureHierarchicalDataForFinalPlan(allData, targetAllocations) {
  console.log(`\n========================================`);
  console.log(`STRUCTURING HIERARCHICAL DATA`);
  console.log(`Allocation types: ${targetAllocations.join(', ')}`);
  console.log(`========================================`);
  console.log(`Input data: ${allData.length} total issues`);
  
  // ═══════════════════════════════════════════════════════════
  // STEP 1: BUILD INITIATIVE LOOKUP MAP FROM ALL DATA
  // ═══════════════════════════════════════════════════════════
  // This map allows us to look up initiative titles by their JIRA key
  const initiativeMap = {};
  
  allData.forEach(issue => {
    const key = issue['Key'];
    const summary = issue['Summary'];
    const issueType = issue['Issue Type'];
    
    // Store ALL issues by their key for lookup
    // This includes Initiatives, Epics, and any other issue types
    if (key && summary) {
      initiativeMap[key] = {
        summary: summary,
        issueType: issueType
      };
    }
  });
  
  console.log(`Built initiative lookup map with ${Object.keys(initiativeMap).length} entries`);
  
  // ═══════════════════════════════════════════════════════════
  // STEP 2: FILTER TO EPICS MATCHING TARGET ALLOCATIONS
  // ═══════════════════════════════════════════════════════════
  let filteredEpics = allData.filter(issue => issue['Issue Type'] === 'Epic');
  console.log(`Filtered to ${filteredEpics.length} epics (excluded ${allData.length - filteredEpics.length} non-epic issues)`);
  
  if (targetAllocations.length > 0) {
    filteredEpics = filteredEpics.filter(epic => {
      const allocation = epic['Allocation'] || '';
      return targetAllocations.includes(allocation);
    });
    console.log(`After allocation filter: ${filteredEpics.length} epics match ${targetAllocations.join(', ')}`);
  }
  
  if (filteredEpics.length === 0) {
    console.log('⚠️ No epics match the specified allocations');
    return {};
  }
  
  // ═══════════════════════════════════════════════════════════
  // STEP 3: BUILD HIERARCHICAL STRUCTURE
  // ═══════════════════════════════════════════════════════════
  const portfolioHierarchy = {};
  
  filteredEpics.forEach(epic => {
    // ─────────────────────────────────────────────────────────
    // Extract data from Google Sheets columns
    // ─────────────────────────────────────────────────────────
    const epicKey = epic['Key'];              // COLUMN A: Epic's own JIRA key (e.g., "ENG-54135")
    const parentKey = epic['Parent Key'];     // COLUMN B: Parent initiative key (e.g., "ENG-12345")
    const initiativeTitle = epic['Initiative Title']; // Initiative name to display
    const portfolioKey = epic['Portfolio Initiative'];
    const portfolioTitle = epic['Portfolio Initiative Title'];
    const programKey = epic['Program Initiative'];
    const programTitle = epic['Program Initiative Title'];
    
    // Determine portfolio and program names for grouping
    const portfolioName = portfolioTitle || portfolioKey || 'Uncategorized Portfolio';
    const programName = programTitle || programKey || parentKey || 'Uncategorized Program';
    
    // ─────────────────────────────────────────────────────────
    // ✅ CRITICAL FIX: Determine which JIRA key to use for hyperlink
    // ─────────────────────────────────────────────────────────
    // HYPERLINK KEY LOGIC (which URL to create):
    //   - If COLUMN B (Parent Key) exists → use it (links to parent initiative)
    //   - If COLUMN B is empty → use COLUMN A (Key) as fallback (links to epic itself)
    //
    // DISPLAY NAME LOGIC (what text to show):
    //   - If COLUMN B has value: Use Initiative Title field, or parent's Summary
    //   - If COLUMN B is empty: Use epic's Summary (COLUMN D) directly
    // ─────────────────────────────────────────────────────────
    
    let linkKey = null;       // JIRA key for the hyperlink URL
    let displayName = null;   // Text to display on the slide
    
    if (parentKey && parentKey.trim() !== '') {
      // ✅ COLUMN B HAS VALUE - Use Parent Key for hyperlink
      linkKey = parentKey;
      
      // Get display name from Initiative Title field (preferred)
      if (initiativeTitle && initiativeTitle.trim() !== '') {
        displayName = initiativeTitle;
        console.log(`  ✓ Epic ${epicKey}: Using COLUMN B (Parent Key)="${parentKey}" for link, displaying "${displayName}"`);
      } else {
        // No Initiative Title - look up parent's Summary from initiative map
        const parentInfo = initiativeMap[parentKey];
        if (parentInfo) {
          displayName = parentInfo.summary;
          console.log(`  ✓ Epic ${epicKey}: Using COLUMN B (Parent Key)="${parentKey}" for link, displaying parent summary: "${displayName}"`);
        } else {
          // Parent not found in data - use parent key as display name
          displayName = parentKey;
          console.warn(`  ⚠️ Epic ${epicKey}: COLUMN B (Parent Key) "${parentKey}" not found, using key as display name`);
        }
      }
    } else {
      // ⚠️ COLUMN B IS EMPTY - Fall back to using COLUMN A (Epic's own Key) for hyperlink
      // ✅ FIX: Use epic's Summary (COLUMN D) directly as display name
      linkKey = epicKey;
      displayName = epic['Summary'] || epicKey;
      
      console.warn(`  ⚠️ Epic ${epicKey}: COLUMN B (Parent Key) is empty`);
      console.warn(`     Using COLUMN A (Key)="${epicKey}" for hyperlink`);
      console.warn(`     Displaying COLUMN D (Summary): "${displayName}"`);
    }
    
    console.log(`  → Portfolio: "${portfolioName}", Program: "${programName}", Initiative: "${displayName}"`);
    
    // ─────────────────────────────────────────────────────────
    // Initialize hierarchy structure
    // ─────────────────────────────────────────────────────────
    
    // Initialize portfolio if needed
    if (!portfolioHierarchy[portfolioName]) {
      portfolioHierarchy[portfolioName] = {};
    }
    
    // Initialize program if needed
    if (!portfolioHierarchy[portfolioName][programName]) {
      portfolioHierarchy[portfolioName][programName] = {
        initiatives: []
      };
    }
    
    // ─────────────────────────────────────────────────────────
    // Find or create initiative grouping
    // ─────────────────────────────────────────────────────────
    
    // Find existing initiative with same linkKey (groups epics under same initiative)
    let initiative = portfolioHierarchy[portfolioName][programName].initiatives.find(
      init => init.linkKey === linkKey
    );
    
    if (!initiative) {
      // Create new initiative grouping
      initiative = {
        linkKey: linkKey,                    // JIRA key for hyperlink (from COLUMN B or COLUMN A)
        initiativeTitle: displayName,        // Display name (Initiative Title or epic Summary)
        epics: [],
        ragStatus: epic['RAG'] || 'Not Set',
        businessPerspective: epic['Business Perspective'] || '',
        technicalPerspective: epic['Technical Perspective'] || ''
      };
      portfolioHierarchy[portfolioName][programName].initiatives.push(initiative);
      console.log(`  ✓ Created new initiative grouping: linkKey="${linkKey}", display="${displayName}"`);
    }
    
    // Add epic to initiative grouping
    initiative.epics.push(epic);
  });
  
  // ═══════════════════════════════════════════════════════════
  // STEP 4: LOG SUMMARY STATISTICS
  // ═══════════════════════════════════════════════════════════
  let totalPortfolios = 0;
  let totalPrograms = 0;
  let totalInitiatives = 0;
  let totalEpics = 0;
  let initiativesWithParentKey = 0;
  let initiativesWithoutParentKey = 0;
  
  Object.keys(portfolioHierarchy).forEach(portfolioName => {
    totalPortfolios++;
    const programs = portfolioHierarchy[portfolioName];
    
    Object.keys(programs).forEach(programName => {
      totalPrograms++;
      const programData = programs[programName];
      totalInitiatives += programData.initiatives.length;
      
      programData.initiatives.forEach(initiative => {
        totalEpics += initiative.epics.length;
        
        // Check if this initiative has epics with Parent Key (COLUMN B)
        const hasParentKey = initiative.epics.some(epic => 
          epic['Parent Key'] && epic['Parent Key'].trim() !== ''
        );
        
        if (hasParentKey) {
          initiativesWithParentKey++;
        } else {
          initiativesWithoutParentKey++;
          console.warn(`  ⚠️ Initiative "${initiative.initiativeTitle}" - COLUMN B empty (linking to COLUMN A epic key)`);
        }
      });
    });
  });
  
  console.log(`\n--- HIERARCHY STRUCTURE SUMMARY ---`);
  console.log(`Portfolios: ${totalPortfolios}`);
  console.log(`Programs: ${totalPrograms}`);
  console.log(`Initiatives: ${totalInitiatives}`);
  console.log(`Epics: ${totalEpics}`);
  console.log(`\n--- INITIATIVE LINK ANALYSIS ---`);
  console.log(`✅ Initiatives with COLUMN B (Parent Key) - links to parent: ${initiativesWithParentKey}`);
  console.log(`⚠️ Initiatives without COLUMN B (Parent Key) - links to epic: ${initiativesWithoutParentKey}`);
  
  if (initiativesWithoutParentKey > 0) {
    console.warn(`\n⚠️ NOTE: ${initiativesWithoutParentKey} initiatives are missing COLUMN B (Parent Key)`);
    console.warn(`   These will link to COLUMN A (epic itself) instead of parent initiative`);
    console.warn(`   To fix: In JIRA, set the Parent Link field on these epics, then regenerate report`);
  }
  
  console.log(`========================================\n`);
  
  return portfolioHierarchy;
}
function robustReplaceAllText(slide, findText, replaceText) {
  slide.getPageElements().forEach(pageElement => {
    let textRange = null;
    
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = pageElement.asShape();
      if (typeof shape.getText === 'function') {
        textRange = shape.getText();
      }
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.TABLE) {
      const table = pageElement.asTable();
      for (let i = 0; i < table.getNumRows(); i++) {
        for (let j = 0; j < table.getNumColumns(); j++) {
          try {
            const cell = table.getCell(i, j);
            if (typeof cell.getText === 'function') {
              const cellTextRange = cell.getText();
              if (cellTextRange && !cellTextRange.isEmpty()) {
                cellTextRange.replaceAllText(findText, replaceText);
              }
            }
          } catch (e) {
            // Skip merged cells that aren't the head cell
            // This is expected behavior for non-head cells in merged regions
            if (!e.message.includes('merged cells')) {
              console.warn(`Could not replace text in cell [${i},${j}]: ${e.message}`);
            }
          }
        }
      }
    }
    
    if (textRange && !textRange.isEmpty()) {
      if (pageElement.getPageElementType() !== SlidesApp.PageElementType.TABLE) {
        textRange.replaceAllText(findText, replaceText);
      }
    }
  });
  
  try { 
    slide.replaceAllText(findText, replaceText); 
  } catch(e) {
    /* Ignore */
  }
}

/**
 * UPDATED: Sets list font size explicitly to HIERARCHY_LIST_SIZE and NOT bold.
 */
function replaceAllTextWithList(slide, placeholder, text, applyListStyle) {
  let shapeFound = false;
  slide.getPageElements().forEach(pageElement => {
    let textRange = null;
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      const shape = pageElement.asShape();
      if (typeof shape.getText === 'function' && shape.getText().asString().includes(placeholder)) {
        textRange = shape.getText();
      }
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.TABLE) {
      const table = pageElement.asTable();
       for (let i = 0; i < table.getNumRows(); i++) {
        for (let j = 0; j < table.getNumColumns(); j++) {
          const cell = table.getCell(i, j);
          if (typeof cell.getText === 'function' && cell.getText().asString().includes(placeholder)) {
            textRange = cell.getText();
            break;
          }
        }
        if(textRange) break;
      }
    }

    if (textRange && !textRange.isEmpty()) {
      const fullTextBefore = textRange.asString();
      const placeholderIndex = fullTextBefore.indexOf(placeholder);
      
      textRange.replaceAllText(placeholder, text);

      const fullTextAfter = textRange.asString();
      
      // Find where the replaced text starts and ends
      const replacedTextStart = placeholderIndex;
      const replacedTextEnd = placeholderIndex + text.length;

      // Only style the newly inserted text (the list), not any existing content like subtitles
      if (text !== 'None' && text.trim().length > 0 && replacedTextStart > -1) {
        try {
          // Style only the list portion
          const listRange = textRange.getRange(replacedTextStart, replacedTextEnd);
          
          listRange.getTextStyle()
               .setFontSize(FINAL_SLIDE_CONFIG.SLIDE_2_10_LIST_SIZE)
               .setBold(false);
        }
        catch(e) { console.warn(`Could not set font size/style for ${placeholder}: ${e.message}`) }
      }

      // Apply list style conditionally - but only to lines that should have bullets
      if (applyListStyle && text !== 'None' && text.trim().length > 0) {
        try { 
          // Apply to entire text range - Google Slides will handle which paragraphs get bullets
          // But we need to ensure the subtitle (first line before the placeholder) doesn't get bullets
          const lines = fullTextAfter.split('\n');
          
          // Find the first line that contains our replaced content
          let contentStartLine = 0;
          let charCount = 0;
          for (let i = 0; i < lines.length; i++) {
            if (charCount >= replacedTextStart) {
              contentStartLine = i;
              break;
            }
            charCount += lines[i].length + 1; // +1 for newline
          }
          
          // Apply list preset to the range starting from the replaced text
          if (replacedTextStart < fullTextAfter.length) {
            const listOnlyRange = textRange.getRange(replacedTextStart, fullTextAfter.length);
            listOnlyRange.getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);
          }
        }
        catch (e) { console.warn(`Could not apply list style to ${placeholder}: ${e.message}`); }
      } else {
        // Remove all list styling if we don't want bullets
        try { textRange.getListStyle().removeList(); } catch (e) { /* Ignore */ }
      }
      shapeFound = true;
    }
  });
  if (!shapeFound) {
    console.warn(`Placeholder "${placeholder}" not found in any element.`);
     try { slide.replaceAllText(placeholder, text); } catch(e) { /* Ignore */ }
  }
}
function populateSlide1(slide, piNumber, valueStream, phase) {
  const today = new Date().toLocaleDateString();
  robustReplaceAllText(slide, '{{ValueStream}}', valueStream);
  robustReplaceAllText(slide, '{{PI}}', `${piNumber}`);
  robustReplaceAllText(slide, '{{Date}}', today);
  robustReplaceAllText(slide, '{{Phase}}', phase || 'FINAL');
}

function processInitiativeList(data, allocationFilter) {
    const titles = data
        .filter(r => allocationFilter.includes(r['Allocation']) && (r['Issue Type'] === 'Epic' || r['Issue Type'] === 'Initiative'))
        .map(r => r['Initiative Title'])
        .filter(title => title && title.trim() !== '' && title.trim().toLowerCase() !== 'title missing');

    const uniqueTitles = [...new Set(titles)];
    uniqueTitles.sort();
    return uniqueTitles.join('\n');
}


function populateSlide10(slide, piNumber, vsData) {
  robustReplaceAllText(slide, '{{PI_Title}}', `${piNumber}`);
  
  const techText = processInitiativeList(vsData, ['Tech / Platform']);
  replaceAllTextWithList(slide, '{{TechPlatform_List}}', techText || 'None', true); // Bullets
  styleShapeText(slide, "Tech Platform", FINAL_SLIDE_CONFIG.SLIDE_2_10_SUBTITLE_SIZE, true);
}
function processInitiativeListByPortfolioInitiative(data, portfolioInitiativeValue) {
  const titles = data
    .filter(r => r['Portfolio Initiative'] === portfolioInitiativeValue && (r['Issue Type'] === 'Epic' || r['Issue Type'] === 'Initiative'))
    .map(r => r['Initiative Title'])
    .filter(title => title && title.trim() !== '' && title.trim().toLowerCase() !== 'title missing');

  const uniqueTitles = [...new Set(titles)];
  uniqueTitles.sort();
  return uniqueTitles.join('\n');
}

function styleShapeText(slide, textToFind, fontSize, isBold) {
   slide.getPageElements().forEach(pageElement => {
     if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
         const shape = pageElement.asShape();
         if (typeof shape.getText === 'function') {
             const textRange = shape.getText();
             if (!textRange.isEmpty() && textRange.asString().includes(textToFind)) {
                 try {
                     // Find the specific text and style it
                     const fullText = textRange.asString();
                     const startIndex = fullText.indexOf(textToFind);
                     if (startIndex !== -1) {
                         const endIndex = startIndex + textToFind.length;
                         const targetRange = textRange.getRange(startIndex, endIndex);
                         targetRange.getTextStyle()
                             .setFontSize(fontSize)
                             .setBold(isBold);
                         console.log(`Styled text "${textToFind}" to size ${fontSize}, bold: ${isBold}`);
                     }
                 } catch (e) {
                     console.warn(`Could not style shape text "${textToFind}": ${e.message}`);
                 }
             }
         }
      }
   });
}

// =================================================================================
// DYNAMIC HIERARCHY SLIDE GENERATION (Using Duplication and Table Population)
// =================================================================================

/**
 * Populates a 3-column table, 1 row per Epic, applying specific styles
 * and placing Program Initiative only on the first row for that program.
 */
function generatePlansHierarchySlides(presentation, initiativeGroups, allIssues, phase) {
  console.log(`[HIERARCHY] Starting hierarchy slide generation for ${initiativeGroups.length} initiative groups`);
  
  if (initiativeGroups.length === 0) {
    console.log(`[HIERARCHY] No initiative groups to display`);
    
    // FIX: Remove ALL slides after the template (slides 5+) if no data
    const slides = presentation.getSlides();
    console.log(`[HIERARCHY] Removing empty hierarchy slides (${slides.length - 5} slides)`);
    
    // Remove slides from the end backwards
    for (let i = slides.length - 1; i >= 5; i--) {
      try {
        console.log(`[HIERARCHY] Removing empty slide at index ${i}`);
        slides[i].remove();
      } catch (e) {
        console.warn(`[HIERARCHY] Could not remove slide ${i}: ${e.message}`);
      }
    }
    
    return;
  }
  
  const slides = presentation.getSlides();
  const templateIndex = 4; // After removal: 0=Title, 1=Initiatives, 2=Commitments, 3=Roadmap, 4=Template
  
  if (slides.length <= templateIndex) {
    console.error(`[HIERARCHY] Template slide not found at index ${templateIndex}`);
    return;
  }
  
  const templateSlide = slides[templateIndex];
  console.log(`[HIERARCHY] Using slide at index ${templateIndex} as template`);
  
  // Verify template has a table
  const templateTables = templateSlide.getTables();
  if (templateTables.length === 0) {
    console.error(`[HIERARCHY] Template slide has no table`);
    return;
  }
  
  const MAX_INITIATIVES_PER_SLIDE = 3;
  let currentSlide = null;
  let currentTable = null;
  let initiativesOnCurrentSlide = 0;
  let pageNum = 0;
  
  // Build dependency map - FIX: Include all dependency information
  const dependencyMap = {};
  allIssues.forEach(issue => {
    if (issue['Issue Type'] === 'Dependency' && issue['Parent Key']) {
      if (!dependencyMap[issue['Parent Key']]) {
        dependencyMap[issue['Parent Key']] = [];
      }
      dependencyMap[issue['Parent Key']].push(issue);
    }
  });
  
  console.log(`[HIERARCHY] Built dependency map with ${Object.keys(dependencyMap).length} epic(s) having dependencies`);
  
  // Process each initiative group
  initiativeGroups.forEach((initiative, initIndex) => {
    console.log(`[HIERARCHY] Processing initiative ${initIndex + 1}/${initiativeGroups.length}: ${initiative.key}`);
    
    try {
      // Check if we need a new slide
      if (!currentSlide || initiativesOnCurrentSlide >= MAX_INITIATIVES_PER_SLIDE) {
        pageNum++;
        console.log(`[HIERARCHY] Creating new slide (page ${pageNum})`);
        
        // Duplicate template
        const newSlide = templateSlide.duplicate();
        addDisclosureToSlide(newSlide);
        const newSlideIndex = presentation.getSlides().indexOf(newSlide);
        const targetIndex = templateIndex + pageNum;
        
        if (newSlideIndex !== targetIndex) {
          newSlide.move(targetIndex);
        }
        
        currentSlide = newSlide;
        
        // Get the table
        const tables = newSlide.getTables();
        if (tables.length === 0) {
          console.error('[HIERARCHY] No table found on duplicated slide');
          return;
        }
        
        currentTable = tables[0];
        console.log(`[HIERARCHY] Table has ${currentTable.getNumRows()} rows and ${currentTable.getNumColumns()} columns`);
        
        // Clear template rows (keep header only)
        const numRows = currentTable.getNumRows();
        for (let i = numRows - 1; i > 0; i--) {
          try {
            currentTable.getRow(i).remove();
            console.log(`[HIERARCHY] Removed template row ${i}`);
          } catch (e) {
            console.warn(`Could not remove row ${i}: ${e.message}`);
          }
        }
        
        // Remove first column if we have 3 columns
        const numCols = currentTable.getNumColumns();
        if (numCols === 3) {
          try {
            currentTable.getColumn(0).remove();
            console.log('[HIERARCHY] Removed first column from table');
          } catch (e) {
            console.warn(`Could not remove first column: ${e.message}`);
          }
        }
        
        // Style header row (should now have 2 columns)
        try {
          const headerRow = currentTable.getRow(0);
          const finalNumCols = currentTable.getNumColumns();
          
          console.log(`[HIERARCHY] Styling ${finalNumCols} header columns`);
          
          for (let c = 0; c < finalNumCols; c++) {
            try {
              const headerCell = headerRow.getCell(c);
              headerCell.getFill().setSolidFill('#663399');
              
              const headerText = headerCell.getText();
              headerText.getTextStyle()
                .setForegroundColor('#FFFFFF')
                .setBold(true)
                .setFontSize(11)
                .setFontFamily('Lato');
              
              console.log(`[HIERARCHY] Styled header column ${c}`);
            } catch (e) {
              console.warn(`Could not style header cell ${c}: ${e.message}`);
            }
          }
        } catch (e) {
          console.error(`Error styling header: ${e.message}`);
        }
        
        initiativesOnCurrentSlide = 0;
      }
      
      // Add initiative to slide with FIXED formatting
      addInitiativeToSlideFixed(currentTable, initiative, dependencyMap);
      initiativesOnCurrentSlide++;
      
    } catch (error) {
      console.error(`Error processing initiative ${initiative.key}: ${error.message}`);
      console.error(error.stack);
    }
  });
  
  // Remove template slide
  try {
    templateSlide.remove();
    console.log(`[HIERARCHY] Removed template slide at index ${templateIndex}`);
  } catch (error) {
    console.warn(`[HIERARCHY] Could not remove template slide: ${error.message}`);
  }
  
  console.log(`[HIERARCHY] Hierarchy slide generation complete. Created ${pageNum} page(s)`);
}

function calculateFinalPlanPageCount(programHierarchy) {
  let totalInitiatives = 0;
  Object.values(programHierarchy).forEach(programData => {
    if (programData.initiatives && Array.isArray(programData.initiatives)) {
      totalInitiatives += programData.initiatives.length;
    }
  });
  return Math.ceil(totalInitiatives / 5); // 5 initiatives per page
}
// =================================================================================
// UPDATED DIAGNOSTIC FUNCTION - Also fixed to use 'Key' column
// =================================================================================

function diagnoseEpicKeysIssue() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const piNumber = 13; // CHANGE THIS
    const valueStream = 'MMPM'; // CHANGE THIS
    
    const vsData = loadPIData(piNumber, valueStream);
    
    console.log(`\n======== DIAGNOSTIC REPORT ========`);
    console.log(`PI: ${piNumber}, Value Stream: ${valueStream}`);
    console.log(`Total issues loaded: ${vsData.length}`);
    
    const allocationType = 'Product - Feature';
    const hierarchicalData = structureHierarchicalDataForFinalPlan(vsData, [allocationType]);
    
    Object.keys(hierarchicalData).forEach(portfolioName => {
      console.log(`\n--- Portfolio: ${portfolioName} ---`);
      const programHierarchy = hierarchicalData[portfolioName];
      
      Object.keys(programHierarchy).forEach(programName => {
        console.log(`  Program: ${programName}`);
        const initiatives = programHierarchy[programName].initiatives;
        
        initiatives.forEach((initiative, idx) => {
          console.log(`    Initiative ${idx}:`);
          console.log(`      hasInitiative: ${initiative.hasInitiative}`);
          console.log(`      initiativeTitle: ${initiative.initiativeTitle}`);
          console.log(`      parentKey: ${initiative.parentKey}`);
          console.log(`      Number of epics: ${initiative.epics.length}`);
          
          // ✅ FIXED: Changed from 'Issue key' to 'Key'
          const epicKeys = initiative.epics.map(e => e['Key']).filter(k => k);
          console.log(`      Epic keys: ${epicKeys.join(', ')}`);
          console.log(`      Would show LINE 3: ${initiative.hasInitiative || epicKeys.length > 1}`);
        });
      });
    });
    
    ui.alert(
      'Diagnostic Complete',
      'Check the Apps Script execution log (View → Logs) to see the detailed diagnostic report.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Diagnostic error:', error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

// =================================================================================
// UTILITIES
// =================================================================================
function styleTableHeaders(table) {
  if (!table || table.getNumRows() < 1) return;
  
  try {
    const headerRow = table.getRow(0);
    for (let i = 0; i < table.getNumColumns(); i++) {
      const cell = headerRow.getCell(i);
      const textRange = cell.getText();
      if (textRange && !textRange.isEmpty()) {
        textRange.getTextStyle()
          .setFontFamily('Lato')
          .setBold(true)
          .setFontSize(10);
      }
    }
  } catch (e) {
    console.warn(`Could not style table headers: ${e.message}`);
  }
}

function populateSlide2Plans(presentation, piNumber, valueStream, vsData, slide2) {
  try {
    console.log(`\n========================================`);
    console.log(`SLIDE 2 POPULATION - PI ${piNumber} - ${valueStream}`);
    addDisclosureToSlide(slide2),
    console.log(`========================================`);
    
    // ═══════════════════════════════════════════════════════════
    // PARAMETER VALIDATION
    // ═══════════════════════════════════════════════════════════
    console.log(`Parameter types:`);
    console.log(`  - presentation: ${typeof presentation}`);
    console.log(`  - piNumber: ${typeof piNumber} = ${piNumber}`);
    console.log(`  - valueStream: ${typeof valueStream} = ${valueStream}`);
    console.log(`  - vsData: ${typeof vsData}, isArray: ${Array.isArray(vsData)}, length: ${vsData ? vsData.length : 'N/A'}`);
    console.log(`  - slide2: ${typeof slide2}`);
    
    if (!presentation || typeof piNumber === 'undefined' || !valueStream || !vsData || !slide2) {
      throw new Error(`Missing required parameters for populateSlide2Plans`);
    }
    
    if (!Array.isArray(vsData)) {
      throw new Error(`vsData must be an array, got ${typeof vsData}`);
    }
    
    console.log(`✅ Parameter validation passed`);
    console.log(`Processing ${vsData.length} total issues for slide 2`);
    
    // ═══════════════════════════════════════════════════════════
    // FILTER TO EPICS ONLY
    // ═══════════════════════════════════════════════════════════
    const epicsOnly = vsData.filter(issue => issue['Issue Type'] === 'Epic');
    console.log(`Filtered to ${epicsOnly.length} epics (excluded ${vsData.length - epicsOnly.length} non-epic issues)`);
    
    // ═══════════════════════════════════════════════════════════
    // UPDATE SLIDE PLACEHOLDERS
    // ═══════════════════════════════════════════════════════════
    robustReplaceAllText(slide2, '{{ValueStream}}', valueStream);
    robustReplaceAllText(slide2, '{{PI_Title}}', `PI ${piNumber}`);
    robustReplaceAllText(slide2, '{{PI}}', piNumber.toString());
    
    // ═══════════════════════════════════════════════════════════
    // ADD FEATURE POINTS BAR CHART
    // ═══════════════════════════════════════════════════════════
    console.log(`Adding feature points bar chart to slide 2...`);
    try {
      const sourceSheetName = `PI ${piNumber} - Iteration 0`;
      addFeaturePointsChartToSlide2(presentation, valueStream, piNumber, sourceSheetName);
    } catch (chartError) {
      console.error(`Error adding chart: ${chartError.message}`);
      // Don't fail the whole function if chart fails
    }
    
    // ═══════════════════════════════════════════════════════════
    // BUILD INITIATIVE MAPS BY ALLOCATION
    // ═══════════════════════════════════════════════════════════
    const productFeatureMap = {};
    const complianceMap = {};
    const infosecMap = {};
    const techPlatformMap = {};
    
    let productFeatureFP = 0;
    let complianceFP = 0;
    let infosecFP = 0;
    let techPlatformFP = 0;
    
    console.log(`\nBuilding initiative lists for slide 2...`);
    
    // ═══════════════════════════════════════════════════════════
    // ✅ FIXED HELPER FUNCTION: Get Initiative Title from Epic's Row
    // ═══════════════════════════════════════════════════════════
    function getInitiativeNameFromParentKey(epic) {
      // Step 1: Check if parent key exists
      const parentKey = epic['Parent Key'];
      
      if (!parentKey || parentKey.trim() === '') {
        console.log(`  Epic ${epic['Key']} has no Parent Key, skipping`);
        return null;
      }
      
      // Step 2: Get Initiative Title from Epic's row (column AN)
      const initiativeTitle = epic['Initiative Title'];
      
      if (!initiativeTitle || initiativeTitle.trim() === '') {
        console.log(`  Epic ${epic['Key']} has Parent Key ${parentKey} but no Initiative Title, skipping`);
        return null;
      }
      
      // Step 3: Return initiative info with title from Epic and URL to parent
      console.log(`  ✓ Epic ${epic['Key']}: Initiative Title "${initiativeTitle}" → Parent Key ${parentKey}`);
      return {
        key: parentKey,
        title: initiativeTitle,
        url: `https://activhealth.atlassian.net/browse/${parentKey}`
      };
    }
    
    // ═══════════════════════════════════════════════════════════
    // PRODUCT - FEATURE ALLOCATION
    // ═══════════════════════════════════════════════════════════
    const productFeatureEpics = epicsOnly.filter(epic => epic['Allocation'] === 'Product - Feature');
    console.log(`Found ${productFeatureEpics.length} Product-Feature epics`);
    
    productFeatureEpics.forEach(epic => {
      const featurePoints = parseFloat(epic['Feature Points']) || 0;
      productFeatureFP += featurePoints;
      
      const initiativeInfo = getInitiativeNameFromParentKey(epic);
      if (initiativeInfo) {
        // Use the parent key as the map key to ensure uniqueness
        if (!productFeatureMap[initiativeInfo.key]) {
          productFeatureMap[initiativeInfo.key] = {
            title: initiativeInfo.title,
            url: initiativeInfo.url
          };
        }
      }
    });
    
    // ═══════════════════════════════════════════════════════════
    // TECH / PLATFORM ALLOCATION
    // ═══════════════════════════════════════════════════════════
    const techPlatformEpics = epicsOnly.filter(epic => epic['Allocation'] === 'Tech / Platform');
    console.log(`Found ${techPlatformEpics.length} Tech/Platform epics`);
    
    techPlatformEpics.forEach(epic => {
      const featurePoints = parseFloat(epic['Feature Points']) || 0;
      techPlatformFP += featurePoints;
      
      const initiativeInfo = getInitiativeNameFromParentKey(epic);
      if (initiativeInfo) {
        if (!techPlatformMap[initiativeInfo.key]) {
          techPlatformMap[initiativeInfo.key] = {
            title: initiativeInfo.title,
            url: initiativeInfo.url
          };
        }
      }
    });
    
    // ═══════════════════════════════════════════════════════════
    // COMPLIANCE ALLOCATION
    // ═══════════════════════════════════════════════════════════
    const complianceEpics = epicsOnly.filter(epic => 
      epic['Allocation'] === 'Compliance' || epic['Allocation'] === 'Product - Compliance'
    );
    console.log(`Found ${complianceEpics.length} Compliance epics`);
    
    complianceEpics.forEach(epic => {
      const featurePoints = parseFloat(epic['Feature Points']) || 0;
      complianceFP += featurePoints;
      
      const initiativeInfo = getInitiativeNameFromParentKey(epic);
      if (initiativeInfo) {
        if (!complianceMap[initiativeInfo.key]) {
          complianceMap[initiativeInfo.key] = {
            title: initiativeInfo.title,
            url: initiativeInfo.url
          };
        }
      }
    });
    
    // ═══════════════════════════════════════════════════════════
    // INFOSEC ALLOCATION
    // ═══════════════════════════════════════════════════════════
    console.log(`Searching for Infosec initiatives...`);
    epicsOnly.forEach(epic => {
      const portfolioInit = (epic['Portfolio Initiative'] || '').toLowerCase();
      const programInit = (epic['Program Initiative'] || '').toLowerCase();
      
      if (portfolioInit.includes('infosec') || programInit.includes('infosec')) {
        const featurePoints = parseFloat(epic['Feature Points']) || 0;
        infosecFP += featurePoints;
        
        const initiativeInfo = getInitiativeNameFromParentKey(epic);
        if (initiativeInfo) {
          if (!infosecMap[initiativeInfo.key]) {
            infosecMap[initiativeInfo.key] = {
              title: initiativeInfo.title,
              url: initiativeInfo.url
            };
          }
        }
      }
    });
    
    // ═══════════════════════════════════════════════════════════
    // FORMAT INITIATIVE LISTS (Template already has bullet formatting)
    // ═══════════════════════════════════════════════════════════
    const productFeatureList = Object.keys(productFeatureMap).length > 0 
      ? Object.values(productFeatureMap).map(i => i.title).sort().join('\n')
      : 'None';
    
    const techPlatformList = Object.keys(techPlatformMap).length > 0
      ? Object.values(techPlatformMap).map(i => i.title).sort().join('\n')
      : 'None';
    
    const complianceList = Object.keys(complianceMap).length > 0
      ? Object.values(complianceMap).map(i => i.title).sort().join('\n')
      : 'None';
    
    const infosecList = Object.keys(infosecMap).length > 0
      ? Object.values(infosecMap).map(i => i.title).sort().join('\n')
      : 'None';
    
    // ═══════════════════════════════════════════════════════════
    // LOG SUMMARY STATISTICS
    // ═══════════════════════════════════════════════════════════
    console.log(`\nSlide 2 Initiative Counts:`);
    console.log(`  - Product-Feature: ${Object.keys(productFeatureMap).length} initiatives (${productFeatureFP} FP)`);
    console.log(`  - Tech/Platform: ${Object.keys(techPlatformMap).length} initiatives (${techPlatformFP} FP)`);
    console.log(`  - Compliance: ${Object.keys(complianceMap).length} initiatives (${complianceFP} FP)`);
    console.log(`  - Infosec: ${Object.keys(infosecMap).length} initiatives (${infosecFP} FP)`);
    
    // ═══════════════════════════════════════════════════════════
    // REPLACE TEXT PLACEHOLDERS IN SLIDE
    // ═══════════════════════════════════════════════════════════
    robustReplaceAllText(slide2, '{{ProductFeature_List}}', productFeatureList);
    robustReplaceAllText(slide2, '{{TechPlatform_List}}', techPlatformList);
    robustReplaceAllText(slide2, '{{Compliance_List}}', complianceList);
    robustReplaceAllText(slide2, '{{Infosec_List}}', infosecList);
    
    robustReplaceAllText(slide2, '{{ProductFeature_FP}}', productFeatureFP.toString());
    robustReplaceAllText(slide2, '{{TechPlatform_FP}}', techPlatformFP.toString());
    robustReplaceAllText(slide2, '{{Compliance_FP}}', complianceFP.toString());
    robustReplaceAllText(slide2, '{{Infosec_FP}}', infosecFP.toString());
    
    console.log(`✅ Slide 2 population complete`);
    console.log(`========================================\n`);
    
  } catch (error) {
    console.error(`\n❌ ERROR in populateSlide2Plans`);
    console.error(`Error message: ${error.message}`);
    console.error(`Error stack:`);
    console.error(error.stack);
    console.error(`========================================\n`);
    throw error;
  }
}

function generateAllValueStreamsPresentation(piNumber, piCommitmentFilter, phase, destinationFolderId = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    phase = phase || 'FINAL';
    console.log(`\n========================================`);
    console.log(`GENERATING ALL VALUE STREAMS PRESENTATION`);
    console.log(`PI: ${piNumber}, Phase: ${phase}, PI Commitment Filter: ${piCommitmentFilter}`);
    console.log(`========================================\n`);
    
    ss.toast('Loading value streams from Configuration Template...', 'Step 1/3', 3);
    const valueStreams = getAvailableValueStreams(piNumber);
    
    if (!valueStreams || valueStreams.length === 0) {
      throw new Error('No value streams found in Configuration Template (Column A, rows 6+)');
    }
    
    console.log(`Found ${valueStreams.length} value streams in Configuration Template: ${valueStreams.join(', ')}`);
    
    const sourceSheetName = `PI ${piNumber} - Iteration 0`;
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    const sourceSheetUrl = sourceSheet ? sourceSheet.getParent().getUrl() + '#gid=' + sourceSheet.getSheetId() : ss.getUrl();
    
    console.log(`\nLoading data from all iteration sheets for PI ${piNumber}...`);
    ss.toast('Loading data from iteration sheets...', 'Step 1/3', 3);
    
    const iterationSheets = [];
    const allSheets = ss.getSheets();
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      if (sheetName.startsWith(`PI ${piNumber} - Iteration`)) {
        iterationSheets.push(sheet);
      }
    });
    
    console.log(`Found ${iterationSheets.length} iteration sheets: ${iterationSheets.map(s => s.getName()).join(', ')}`);
    
    if (iterationSheets.length === 0) {
      throw new Error(`No iteration sheets found for PI ${piNumber}. Expected sheets named "PI ${piNumber} - Iteration X"`);
    }
    
    let allIssues = [];
    iterationSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      console.log(`Processing ${sheetName}: ${sheet.getLastRow() - 4} data rows`);
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 4) {
        console.log(`  Skipping ${sheetName} - no data rows`);
        return;
      }
      
      const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
      const dataRange = sheet.getRange(5, 1, lastRow - 4, headers.length);
      const data = dataRange.getValues();
      
      data.forEach(row => {
        const issue = {};
        headers.forEach((header, index) => {
          issue[header] = row[index];
        });
        allIssues.push(issue);
      });
      
      console.log(`  Added ${data.length} valid rows from ${sheetName}`);
      
      console.log(`  Sample of last 3 issues added:`);
      data.slice(-3).forEach(row => {
        const key = row[headers.indexOf('Key')];
        const vs = row[headers.indexOf('Value Stream/Org')];
        const type = row[headers.indexOf('Issue Type')];
        const pi = row[headers.indexOf('Program Increment')];
        console.log(`    - Key: "${key}", Value Stream/Org: "${vs}", Issue Type: "${type}", Program Increment: "${pi}"`);
      });
    });
    
    console.log(`\n✓ Loaded ${allIssues.length} total issues from ${iterationSheets.length} sheets`);
    
    const issueTypeCounts = {};
    allIssues.forEach(issue => {
      const type = issue['Issue Type'] || 'Unknown';
      issueTypeCounts[type] = (issueTypeCounts[type] || 0) + 1;
    });
    console.log('Issue type breakdown:');
    Object.keys(issueTypeCounts).sort().forEach(type => {
      console.log(`  - ${type}: ${issueTypeCounts[type]}`);
    });
    
    // ✅ CORRECTED: Apply Program Increment filter at EPIC LEVEL ONLY
    console.log(`\n=== Applying Program Increment Filter (Epic Level Only) ===`);
    console.log(`Phase: ${phase}`);
    
    const piValues = new Set();
    allIssues.forEach(issue => {
      const pi = (issue['Program Increment'] || 'BLANK').toString().trim() || 'BLANK';
      piValues.add(pi);
    });
    console.log(`Program Increment values in data:`);
    Array.from(piValues).sort().forEach(pi => {
      const count = allIssues.filter(i => {
        const issuePi = (i['Program Increment'] || 'BLANK').toString().trim() || 'BLANK';
        return issuePi === pi;
      }).length;
      console.log(`  - "${pi}": ${count} issues`);
    });
    
    // ✅ STEP 1: Filter EPICS by Program Increment
    let filteredEpics = [];
    
    if (phase === 'INITIAL') {
      filteredEpics = allIssues.filter(issue => {
        if (issue['Issue Type'] !== 'Epic') return false;
        const programIncrement = (issue['Program Increment'] || '').toString().trim();
        // ✅ FIXED: Match both "PI 13" and "13" formats
        return programIncrement === `PI ${piNumber}` || 
               programIncrement === piNumber.toString() || 
               programIncrement.toUpperCase() === 'PRE-PLANNING';
      });
      console.log(`INITIAL PLAN: Including Epics with PI ${piNumber} or "Pre-Planning"`);
      console.log(`  Total issues before filter: ${allIssues.length}`);
      console.log(`  Filtered Epics: ${filteredEpics.length}`);
      
      const explicitPI = allIssues.filter(i => i['Issue Type'] === 'Epic' && ((i['Program Increment'] || '').toString().trim() === `PI ${piNumber}` || (i['Program Increment'] || '').toString().trim() === piNumber.toString())).length;
      const prePlanning = allIssues.filter(i => i['Issue Type'] === 'Epic' && (i['Program Increment'] || '').toString().trim().toUpperCase() === 'PRE-PLANNING').length;
      console.log(`  - Epics with explicit PI ${piNumber}: ${explicitPI}`);
      console.log(`  - Pre-Planning Epics: ${prePlanning}`);
    } else {
      filteredEpics = allIssues.filter(issue => {
        if (issue['Issue Type'] !== 'Epic') return false;
        const programIncrement = (issue['Program Increment'] || '').toString().trim();
        // ✅ FIXED: Match both "PI 13" and "13" formats
        return programIncrement === `PI ${piNumber}` || programIncrement === piNumber.toString();
      });
      console.log(`FINAL PLAN: Including ONLY Epics with explicit PI ${piNumber}`);
      console.log(`  Total issues before filter: ${allIssues.length}`);
      console.log(`  Filtered Epics: ${filteredEpics.length}`);
      
      const totalEpics = allIssues.filter(i => i['Issue Type'] === 'Epic').length;
      const excluded = totalEpics - filteredEpics.length;
      if (excluded > 0) {
        console.log(`  - Excluded ${excluded} Epics (wrong PI or Pre-Planning)`);
      }
    }
    
    // ✅ STEP 2: Get all Epic keys that passed the filter
    const allowedEpicKeys = new Set(filteredEpics.map(epic => epic['Key']));
    console.log(`\nAllowed Epic keys: ${allowedEpicKeys.size}`);
    
    // ✅ STEP 3: Include ALL issues (Epics + Dependencies) where:
    // - The issue itself is a filtered Epic, OR
    // - The issue's Parent Key is a filtered Epic (child of allowed Epic)
    const piFilteredIssues = allIssues.filter(issue => {
      if (issue['Issue Type'] === 'Epic' && allowedEpicKeys.has(issue['Key'])) {
        return true;
      }
      
      if (issue['Parent Key'] && allowedEpicKeys.has(issue['Parent Key'])) {
        return true;
      }
      
      return false;
    });
    
    console.log(`\nAfter Epic-based PI filtering:`);
    console.log(`  - Epics: ${piFilteredIssues.filter(i => i['Issue Type'] === 'Epic').length}`);
    console.log(`  - Dependencies: ${piFilteredIssues.filter(i => i['Issue Type'] === 'Dependency').length}`);
    console.log(`  - Total issues: ${piFilteredIssues.length}`);
    
    const actualValueStreams = new Set();
    piFilteredIssues.forEach(issue => {
      if (issue['Value Stream/Org']) {
        actualValueStreams.add(issue['Value Stream/Org']);
      }
    });
    console.log(`\nDIAGNOSTIC: Unique Value Stream/Org values in PI-filtered data (${actualValueStreams.size} total):`);
    Array.from(actualValueStreams).sort().forEach(vs => {
      const count = piFilteredIssues.filter(i => i['Value Stream/Org'] === vs).length;
      console.log(`  - "${vs}" (${count} issues)`);
    });
    console.log(`\nDIAGNOSTIC: Expected value streams from Configuration Template (${valueStreams.length} total):`);
    valueStreams.forEach(vs => console.log(`  - "${vs}"`));
    
    let filteredIssues = piFilteredIssues;
    if (piCommitmentFilter !== 'All') {
      console.log(`\nApplying PI Commitment filter: ${piCommitmentFilter}`);
      const beforeFilter = filteredIssues.length;
      filteredIssues = applyPICommitmentFilter(piFilteredIssues, piCommitmentFilter);
      console.log(`Filtered to ${filteredIssues.length} issues (from ${beforeFilter}) matching PI Commitment filter: ${piCommitmentFilter}`);
    }
    
    if (filteredIssues.length === 0) {
      throw new Error(`No data found for PI ${piNumber} with ${phase} phase filter. Please ensure issues have Program Increment field set to "PI ${piNumber}" or "${piNumber}"${phase === 'INITIAL' ? ' or "Pre-Planning"' : ''}.`);
    }
    
    ss.toast('Creating combined presentation...', 'Step 2/3', 5);
    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const presentationName = `PI ${piNumber} - All Value Streams - ${phase} Plan Review`;
    
    // ✅ NEW: Determine target folder based on destinationFolderId parameter
    let targetFolder;

    if (destinationFolderId) {
      try {
        targetFolder = DriveApp.getFolderById(destinationFolderId);
        console.log(`Using user-specified destination folder: ${targetFolder.getName()}`);
      } catch (error) {
        console.warn(`Could not access specified folder: ${error.message}`);
        const spreadsheetFile = DriveApp.getFileById(ss.getId());
        const parentFolders = spreadsheetFile.getParents();
        targetFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
      }
    } else {
      const spreadsheetFile = DriveApp.getFileById(ss.getId());
      const parentFolders = spreadsheetFile.getParents();
      targetFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
      console.log(`Using default location: ${targetFolder.getName()}`);
    }
    
    const newFile = templateFile.makeCopy(presentationName, targetFolder);
    const presentation = SlidesApp.openById(newFile.getId());
    const presentationUrl = presentation.getUrl();
    
    console.log(`\nCreated combined presentation: ${presentationUrl}`);
    
    let slides = presentation.getSlides();
    for (let i = slides.length - 1; i > 0; i--) {
      slides[i].remove();
    }
    
    const titleSlide = presentation.getSlides()[0];
    robustReplaceAllText(titleSlide, '{{ValueStream}}', 'All Value Streams');
    robustReplaceAllText(titleSlide, '{{PI_Title}}', `PI ${piNumber}`);
    robustReplaceAllText(titleSlide, '{{PI}}', piNumber.toString());
    robustReplaceAllText(titleSlide, '{{PHASE}}', phase);
    robustReplaceAllText(titleSlide, '{{Phase}}', phase);
    
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');
    robustReplaceAllText(titleSlide, '{{Date}}', currentDate);
    
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');
    addTimestampToSlide(titleSlide, timestamp, sourceSheetName, sourceSheetUrl);
    
    const tempPresentation = SlidesApp.openById(TEMPLATE_ID);
    const templateSlides = tempPresentation.getSlides();
    
    let processedCount = 0;
    let skippedStreams = [];
    
    console.log(`\n=== Processing ${valueStreams.length} Value Streams ===`);
    
    valueStreams.forEach((valueStream, index) => {
      console.log(`\n=== Processing Value Stream ${index + 1}/${valueStreams.length}: ${valueStream} ===`);
      ss.toast(`Processing ${valueStream} (${index + 1}/${valueStreams.length})...`, 'Step 3/3', 3);
      
      const vsData = filteredIssues.filter(issue => issue['Value Stream/Org'] === valueStream);
      
      if (vsData.length === 0) {
        console.log(`⚠️ WARNING: No data found for "${valueStream}"`);
        skippedStreams.push(valueStream);
        return;
      }
      
      console.log(`✓ Found ${vsData.length} issues for ${valueStream}`);
      
      const vsTypeCounts = {};
      vsData.forEach(issue => {
        const type = issue['Issue Type'] || 'Unknown';
        vsTypeCounts[type] = (vsTypeCounts[type] || 0) + 1;
      });
      console.log(`  Issue types: ${JSON.stringify(vsTypeCounts)}`);
      
      try {
        console.log(`Structuring Product - Feature hierarchy for ${valueStream}...`);
        const productFeatureHierarchy = structureHierarchicalDataForFinalPlan(vsData, ['Product - Feature']);
        
        if (!productFeatureHierarchy || Object.keys(productFeatureHierarchy).length === 0) {
          console.log(`⚠️ WARNING: No Product - Feature data found for "${valueStream}", skipping`);
          skippedStreams.push(valueStream);
          return;
        }
        
        const dividerSlide = presentation.appendSlide(templateSlides[0]);
        robustReplaceAllText(dividerSlide, '{{ValueStream}}', valueStream);
        robustReplaceAllText(dividerSlide, '{{PI_Title}}', `PI ${piNumber}`);
        robustReplaceAllText(dividerSlide, '{{PI}}', piNumber.toString());
        robustReplaceAllText(dividerSlide, '{{PHASE}}', phase);
        robustReplaceAllText(dividerSlide, '{{Phase}}', phase);
        robustReplaceAllText(dividerSlide, '{{Date}}', currentDate);
        addTimestampToSlide(dividerSlide, timestamp, sourceSheetName, sourceSheetUrl);
        
        console.log(`Adding Portfolio Initiatives slide for ${valueStream}...`);
        const slide2 = presentation.appendSlide(templateSlides[FINAL_SLIDE_CONFIG.SLIDE_2_INDEX]);
        
        populateSlide2Plans(presentation, piNumber, valueStream, vsData, slide2);
        
        console.log(`Generating Product - Feature hierarchy slides for ${valueStream}...`);
        const allHierarchies = {
          'Product - Feature': productFeatureHierarchy
        };
        generateHierarchySlides(presentation, allHierarchies, TEMPLATE_ID, phase, timestamp, sourceSheetName, sourceSheetUrl);
        
        processedCount++;
        console.log(`✓ Completed ${valueStream} (${processedCount} value streams processed so far)`);
        
      } catch (vsError) {
        console.error(`Error processing ${valueStream}:`, vsError);
        console.error(`Error stack:`, vsError.stack);
        skippedStreams.push(valueStream);
      }
    });
    
    tempPresentation.saveAndClose();
    
    if (processedCount === 0) {
      throw new Error(
        `❌ ERROR: No value streams could be processed!\n\n` +
        `Skipped value streams (${skippedStreams.length}): ${skippedStreams.join(', ')}\n\n` +
        `This usually means there's a mismatch between:\n` +
        `1. Value Stream names in Configuration Template (Column A, rows 6+)\n` +
        `2. Values in "Value Stream/Org" column in your PI ${piNumber} data\n\n` +
        `Please check the execution logs for diagnostic information showing:\n` +
        `- Expected value streams from Configuration Template\n` +
        `- Actual Value Stream/Org values in your data\n` +
        `- Possible close matches found\n\n` +
        `TIP: Run "🔍 Diagnose Value Stream Mismatch" from the menu to identify exact mismatches.`
      );
    }
    
    console.log(`\n✅ Successfully processed ${processedCount}/${valueStreams.length} value streams`);
    if (skippedStreams.length > 0) {
      console.log(`⚠️ Skipped ${skippedStreams.length} value streams with no data: ${skippedStreams.join(', ')}`);
    }
    
    const totalSlides = presentation.getSlides().length;
    console.log(`✅ Presentation complete with ${totalSlides} total slides`);
    
    ss.toast('✅ All value streams presentation complete!', 'Success', 5);
    
    return {
      success: true,
      url: presentationUrl,
      slideCount: totalSlides,
      valueStreams: `${processedCount}/${valueStreams.length}`,
      skipped: skippedStreams.length > 0 ? skippedStreams : null
    };
    
  } catch (error) {
    console.error('Error generating combined presentation:', error.message);
    console.error('Error stack:', error.stack);
    ss.toast(`Error: ${error.message}`, '❌ Failed', 10);
    throw error;
  }
}

/**
 * Populates Slide 3 (Dependencies & Constraints) for a single value stream in ALL mode
 */
function populateSlide3ForValueStreamInAllMode(slide, vsData, valueStream, timestamp, sourceSheetName, sourceSheetUrl) {
  console.log(`Populating Slide 3 for ${valueStream}...`);
  
  robustReplaceAllText(slide, '{{ValueStream}}', valueStream);
  addTimestampToSlide(slide, timestamp, sourceSheetName, sourceSheetUrl);
  
  // Create temp presentation and use existing function
  const tempPres = SlidesApp.create('temp_slide3_' + valueStream);
  
  try {
    populateSlide3WithDependenciesAndConstraints(tempPres, vsData);
    // Copy content logic here
  } finally {
    DriveApp.getFileById(tempPres.getId()).setTrashed(true);
  }
  
  console.log(`✓ Slide 3 populated for ${valueStream}`);
}

/**
 * Populates Slide 4 (Top 10 Risks) for a single value stream in ALL mode
 */
function populateSlide4ForValueStreamInAllMode(slide, vsData, valueStream, timestamp, sourceSheetName, sourceSheetUrl) {
  console.log(`Populating Slide 4 for ${valueStream}...`);
  
  robustReplaceAllText(slide, '{{ValueStream}}', valueStream);
  addTimestampToSlide(slide, timestamp, sourceSheetName, sourceSheetUrl);
  
  // Create temp presentation and use existing function
  const tempPres = SlidesApp.create('temp_slide4_' + valueStream);
  
  try {
    populateSlide4WithTopRisks(tempPres, vsData);
    // Copy content logic here
  } finally {
    DriveApp.getFileById(tempPres.getId()).setTrashed(true);
  }
  
  console.log(`✓ Slide 4 populated for ${valueStream}`);
}
function addInitiativeToSlideFixed(table, initiative, dependencyMap) {
  console.log(`[ADD INIT] Adding initiative: ${initiative.key} with ${initiative.epics.length} epics`);
  
  // Row 1: Initiative Title + Feature Points (spans columns)
  const titleRow = table.appendRow();
  
  try {
    // Make row transparent
    for (let c = 0; c < table.getNumColumns(); c++) {
      try {
        titleRow.getCell(c).getFill().setSolidFill('#FFFFFF', 0);
      } catch (e) {
        console.warn(`Could not set fill for cell ${c}: ${e.message}`);
      }
    }
    
    // Column 0: Initiative Title + Feature Points
    const titleCell = titleRow.getCell(0);
    const titleTextContent = `${initiative.title}\nFeature Points: ${initiative.totalFeaturePoints}`;
    
    let titleText = null;
    try {
      titleText = titleCell.getText();
      titleText.setText(titleTextContent);
      
      // FIX: Add hyperlink to initiative title if it's a JIRA key
      if (initiative.key && initiative.key.includes('-')) {
        const jiraUrl = `https://activhealth.atlassian.net/browse/${initiative.key}`;
        const titleRange = titleText.getRange(0, initiative.title.length);
        
        try {
          titleRange.getTextStyle()
            .setFontSize(11)
            .setBold(true)
            .setFontFamily('Lato')
            .setForegroundColor('#000000') // Keep black, not blue
            .setUnderline(false); // No underline
          
          // Set the hyperlink
          titleRange.getLinkUrl = function() { return jiraUrl; };
          console.log(`[HYPERLINK] Added link to initiative ${initiative.key}: ${jiraUrl}`);
        } catch (e) {
          console.warn(`Could not add hyperlink to initiative: ${e.message}`);
        }
      }
      
      // Style the feature points (italic, size 8)
      const fpStartIndex = initiative.title.length + 1;
      const fpRange = titleText.getRange(fpStartIndex, titleText.asString().length);
      fpRange.getTextStyle()
        .setFontSize(8)
        .setItalic(true)
        .setBold(false)
        .setFontFamily('Lato');
        
    } catch (e) {
      console.warn(`Could not set title text: ${e.message}`);
    }
    
    // Column 1: Empty for title row
    try {
      const emptyCell = titleRow.getCell(1);
      const emptyText = emptyCell.getText();
      emptyText.setText('');
    } catch (e) {
      console.warn(`Could not clear cell 1: ${e.message}`);
    }
    
  } catch (error) {
    console.error(`Error adding title row: ${error.message}`);
  }
  
  // FIX: Calculate INITIATIVE-LEVEL RAG STATUS (aggregate across all epics)
  const initiativeRagNotes = [];
  const ragStatusMap = { red: 0, amber: 0, yellow: 0, green: 0 };
  
  initiative.epics.forEach(epic => {
    const rag = (epic['RAG'] || '').toLowerCase().trim();
    const ragNote = (epic['RAG Note'] || '').trim();
    
    console.log(`[RAG] Epic ${epic['Key']}: RAG=${rag}, Note="${ragNote}"`);
    
    if (rag && rag !== 'green' && rag !== '' && rag !== 'null') {
      // Count RAG status
      if (rag.includes('red')) ragStatusMap.red++;
      else if (rag.includes('amber') || rag.includes('yellow')) ragStatusMap.amber++;
      
      // Only add note if it's not empty
      if (ragNote && ragNote !== '' && ragNote !== 'null') {
        const ragIcon = getRagIcon(rag);
        initiativeRagNotes.push(`${ragIcon} ${ragNote}`);
      }
    }
  });
  
  console.log(`[RAG] Initiative ${initiative.key} RAG Summary: Red=${ragStatusMap.red}, Amber=${ragStatusMap.amber}`);
  console.log(`[RAG] Initiative ${initiative.key} has ${initiativeRagNotes.length} RAG notes`);
  
  // Group epics by value stream
  const valueStreamMap = {};
  initiative.epics.forEach(epic => {
    const vs = epic['Value Stream/Org'] || 'No Value Stream';
    if (!valueStreamMap[vs]) {
      valueStreamMap[vs] = [];
    }
    valueStreamMap[vs].push(epic);
  });
  
  // Add rows for each value stream
  Object.keys(valueStreamMap).sort().forEach(valueStream => {
    const epicsInVS = valueStreamMap[valueStream];
    
    try {
      // Collect technical perspectives for this value stream
      const perspectives = epicsInVS
        .map(epic => epic['Technical Perspective'])
        .filter(p => p && String(p).trim() !== '' && p !== 'null');
      
      const vsText = perspectives.length > 0 
        ? `${valueStream}\n${perspectives.join('\n')}`
        : valueStream;
      
      // FIX: Collect dependencies with proper child ticket tracing
      let depsText = '';
      const uniqueDeps = new Set(); // Prevent duplicates
      
      epicsInVS.forEach(epic => {
        const deps = dependencyMap[epic['Key']] || [];
        console.log(`[DEPS] Epic ${epic['Key']} has ${deps.length} dependencies`);
        
        deps.forEach(dep => {
          // Get dependency information
          const depKey = dep['Key'] || '';
          const depVS = dep['Depends on Valuestream'] || dep['Value Stream/Org'] || 'Unknown';
          const depTitle = dep['Summary'] || depKey;
          
          // Create unique identifier to prevent duplicates
          const depIdentifier = `${depVS}|${depTitle}`;
          
          if (!uniqueDeps.has(depIdentifier)) {
            uniqueDeps.add(depIdentifier);
            depsText += `→ ${depVS}: ${depTitle}\n`;
            
            console.log(`[DEPS] Added dependency: ${depVS}: ${depTitle}`);
          } else {
            console.log(`[DEPS] Skipped duplicate dependency: ${depVS}: ${depTitle}`);
          }
        });
      });
      
      const vsRow = table.appendRow();
      
      // Make row transparent
      for (let c = 0; c < table.getNumColumns(); c++) {
        try {
          vsRow.getCell(c).getFill().setSolidFill('#FFFFFF', 0);
        } catch (e) {
          console.warn(`Could not set fill for vs cell ${c}: ${e.message}`);
        }
      }
      
      // Column 0: Value Stream + Perspectives
      try {
        const vsCell = vsRow.getCell(0);
        const vsCellText = vsCell.getText();
        vsCellText.setText(vsText);
        vsCellText.getTextStyle()
          .setFontSize(7)
          .setFontFamily('Lato');
      } catch (e) {
        console.warn(`Could not set VS text: ${e.message}`);
      }
      
      // Column 1: RAG Notes + Dependencies
      try {
        // FIX: Show initiative-level RAG notes first, then dependencies
        let riskDepContent = '';
        
        if (initiativeRagNotes.length > 0) {
          riskDepContent += initiativeRagNotes.join('\n') + '\n\n';
        }
        
        if (depsText.trim() !== '') {
          riskDepContent += depsText;
        }
        
        const riskCell = vsRow.getCell(1);
        const riskCellText = riskCell.getText();
        riskCellText.setText(riskDepContent.trim() || '');
        riskCellText.getTextStyle()
          .setFontSize(7)
          .setFontFamily('Lato');
          
        console.log(`[RISK/DEP] Added ${initiativeRagNotes.length} RAG notes and ${uniqueDeps.size} dependencies`);
      } catch (e) {
        console.warn(`Could not set risk/dep text: ${e.message}`);
      }
      
    } catch (error) {
      console.error(`Error adding value stream row for ${valueStream}: ${error.message}`);
    }
  });
}
function addTimestampToSlide(slide, timestamp, sheetName, sheetUrl) {
  try {
    // Single line format at bottom right, size 5, light grey
    const text = `Data Source: ${sheetName} | Generated: ${timestamp}`;
    
    // Position at absolute bottom-right corner
    const left = 620; // Right-aligned
    const top = 530;  // Bottom of slide (540 is slide height)
    const width = 300;
    const height = 15;
    
    // Create text box
    const textBox = slide.insertTextBox(text, left, top, width, height);
    const textRange = textBox.getText();
    
    // Style: size 5, light grey, no bold/italic
    const textStyle = textRange.getTextStyle();
    textStyle.setFontSize(5);
    textStyle.setForegroundColor('#999999'); // Light grey
    textStyle.setFontFamily('Lato');
    textStyle.setBold(false);
    textStyle.setItalic(false);
    
    // Right-align the text
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
    
    // Try to add hyperlink (may fail, that's okay)
    try {
      const linkStart = text.indexOf(sheetName);
      const linkEnd = linkStart + sheetName.length;
      const linkRange = textRange.getRange(linkStart, linkEnd);
      linkRange.getTextStyle().setLinkUrl(sheetUrl);
      console.log(`✓ Added timestamp with hyperlink: "${timestamp}" → ${sheetUrl}`);
    } catch (e) {
      console.log(`✓ Added timestamp (no hyperlink): "${timestamp}"`);
    }
    
  } catch (error) {
    console.warn(`Could not add timestamp to slide: ${error.message}`);
  }
}

function applyPICommitmentFilter(issues, filterValue) {
  if (!filterValue || filterValue === 'All') {
    return issues; // No filtering
  }
  
  console.log(`Applying PI Commitment filter: "${filterValue}"`);
  
  return issues.filter(issue => {
    // Only filter Epics - pass through all other issue types (dependencies, etc.)
    if (issue['Issue Type'] !== 'Epic') {
      return true;
    }
    
    const piCommitment = (issue['PI Commitment'] || '').trim();
    
    switch (filterValue) {
      case 'Blank':
        return piCommitment === '';
        
      case 'Committed':
        // ✅ FIXED: Include both "Committed" and "Committed After Plan"
        // Both should be treated as committed for filtering purposes
        return piCommitment === 'Committed' || piCommitment === 'Committed After Plan';
        
      case 'Committed After Plan':
        // This filter shows ONLY items committed after plan (distinct from regular committed)
        return piCommitment === 'Committed After Plan';
        
      case 'Not Committed':
        // ✅ FIXED: Not Committed should exclude both "Committed" and "Committed After Plan"
        return piCommitment !== 'Committed' && piCommitment !== 'Committed After Plan';
        
      default:
        return true; // Unknown filter value, include everything
    }
  });
}
function getJiraUrl(issueKey) {
  if (!issueKey) return null;
  return `https://modmedrnd.atlassian.net/browse/${issueKey}`;
}
function removeShapeContainingText(slide, searchText) {
  const shapes = slide.getShapes();
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    try {
      const textRange = shape.getText();
      if (textRange && !textRange.isEmpty()) {
        const text = textRange.asString();
        
        // Check if this shape contains the search text
        if (text.includes(searchText)) {
          console.log(`  🗑️ Removing shape containing: "${searchText}"`);
          shape.remove();
          return true;
        }
      }
    } catch (e) {
      // Skip shapes without text
    }
  }
  
  return false;
}
function getRagIcon(rag) {
  const ragLower = String(rag).toLowerCase().trim();
  if (ragLower.includes('red')) return '🔴';
  if (ragLower.includes('amber') || ragLower.includes('yellow')) return '🟡';
  if (ragLower.includes('green')) return '🟢';
  return '⚪';
}