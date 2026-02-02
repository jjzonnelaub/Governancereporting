/**
 * ════════════════════════════════════════════════════════════════════════════
 * PREDICTABILITY PRESENTATION - UPDATED FOR CHART GENERATION
 * ════════════════════════════════════════════════════════════════════════════
 * 
 * Based on template: POC - Program Predictability Template
 * 
 * Slide Structure:
 * - Slide 1: Title ({{PI #}}, {{Date Generated}})
 * - Slide 2: SKIP (was "PI Commitment Changelog: Deferred" header)
 * - Slide 3: Deferral table data
 * - Slide 4: "Program Predictability by Value Stream" section header
 * - Slide 5+: Value stream slides with 6 charts each:
 *   1. Capacity & Velocity Chart
 *   2. Allocation Chart
 *   3. Epic Status Chart
 *   4. Velocity Trending Chart
 *   5. Program Predictability Chart
 *   6. BV/AV Summary Chart
 */

// ════════════════════════════════════════════════════════════════════════════════════════════════════════
// CONSTANTS
// ════════════════════════════════════════════════════════════════════════════════════════════════════════
const DEFERRAL_CATEGORIES = ['Infosec', 'Product - Compliance', 'Product - Feature', 'Platform', 'Unknown'];
const ISSUES_PER_SLIDE = 7;
const JIRA_BASE_URL_PRED = 'https://modmedrnd.atlassian.net/browse/';
/**
 * MAIN GENERATION FUNCTION - REPLACE YOUR CURRENT generatePredictabilityPresentation
 */
function generatePredictabilityPresentation(piNumber, destinationFolderId, deferralSortBy, regenerateCharts, includeObjectiveNotMet) {
  return logActivity('Predictability Report Presentation Generation', () => {
    const startTime = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Default parameter values
    deferralSortBy = deferralSortBy || 'category';
    includeObjectiveNotMet = includeObjectiveNotMet !== false; // Default to true
    
    console.log(`\n${'='.repeat(80)}`);
    console.log(`PREDICTABILITY REPORT PRESENTATION - PI ${piNumber}`);
    console.log(`Include Objective Not Met: ${includeObjectiveNotMet}`);
    console.log(`Regenerate Charts: ${regenerateCharts}`);
    console.log(`${'='.repeat(80)}\n`);
    
    ss.toast('Starting Predictability Report generation...', '📊 Initializing', 10);
    
    try {
      // ═══════════════════════════════════════════════════════════════════
      // STEP 1: VALIDATE SOURCE DATA
      // ═══════════════════════════════════════════════════════════════════
      const sourceSheetName = `PI ${piNumber} - Iteration 6`;
      const sourceSheet = ss.getSheetByName(sourceSheetName);
      
      if (!sourceSheet) {
        throw new Error(`Source sheet not found: ${sourceSheetName}`);
      }
      
      const predScoreSheet = ss.getSheetByName('Predictability Score');
      if (!predScoreSheet) {
        throw new Error('Predictability Score sheet not found');
      }
      
      console.log(`✓ Source sheet found: ${sourceSheetName}`);
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 2: CREATE PRESENTATION FROM TEMPLATE
      // ═══════════════════════════════════════════════════════════════════
      ss.toast('Creating presentation from template...', '📑 Creating', 10);
      
      const templateFile = DriveApp.getFileById(PREDICTABILITY_TEMPLATE_ID);
      
      let destinationFolder;
      if (destinationFolderId) {
        destinationFolder = DriveApp.getFolderById(destinationFolderId);
      } else {
        const spreadsheetFile = DriveApp.getFileById(ss.getId());
        const parentFolders = spreadsheetFile.getParents();
        destinationFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
      }
      
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');
      const presentationName = `Predictability Report - PI ${piNumber} - ${timestamp}`;
      
      const newFile = templateFile.makeCopy(presentationName, destinationFolder);
      const presentation = SlidesApp.openById(newFile.getId());
      const presentationUrl = presentation.getUrl();
      
      console.log(`✓ Created presentation: ${presentationUrl}`);
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 3: UPDATE TITLE SLIDE
      // ═══════════════════════════════════════════════════════════════════
      const slides = presentation.getSlides();
      const titleSlide = slides[0];
      
      robustReplaceAllText(titleSlide, '{{PI #}}', `PI ${piNumber}`);
      const dateGenerated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy hh:mm a');
      robustReplaceAllText(titleSlide, '{{Date Generated}}', dateGenerated);
      
      console.log(`✓ Title slide updated`);
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 4: GENERATE DEFERRAL SLIDES
      // ═══════════════════════════════════════════════════════════════════
      ss.toast('Generating deferral slides...', '✏️ Step 2/3', 10);
      console.log(`\n📊 Generating Deferral Slides...`);

      let deferralSlides = 0;

      try {
        // ✅ Pass includeObjectiveNotMet as 3rd parameter
        const deferredIssues = extractDeferredIssuesFromSheet(
          sourceSheet, 
          deferralSortBy,
          includeObjectiveNotMet
        );
        
        console.log(`Found ${deferredIssues ? deferredIssues.length : 0} issues`);
        
        if (deferredIssues && deferredIssues.length > 0) {
          deferralSlides = generateDeferralSlidesForPredictability(
            presentation,
            deferredIssues,
            undefined
          );
          console.log(`✓ Generated ${deferralSlides} deferral slide(s)`);
        } else {
          console.log('⚠️  No deferred issues found');
        }
      } catch (error) {
        console.error(`❌ Deferral error: ${error.message}`);
        deferralSlides = 0;
      }
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 5: GENERATE CHARTS
      // ═══════════════════════════════════════════════════════════════════
      const chartSheetName = `PI ${piNumber} - VS Charts`;
      const existingChartSheet = ss.getSheetByName(chartSheetName);
      
      if (!existingChartSheet || regenerateCharts) {
        ss.toast('Generating charts...', '📈 Charts', 10);
        generateValueStreamChartsSheetProgrammatically(piNumber, ss, predScoreSheet, sourceSheet);
      }
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 6: GENERATE VALUE STREAM SLIDES
      // ═══════════════════════════════════════════════════════════════════
      ss.toast('Generating value stream slides...', '✏️ Step 3/3', 15);
      
      const valueStreamSlides = generateValueStreamSlides(
        presentation,
        piNumber,
        predScoreSheet
      );
      
      // ═══════════════════════════════════════════════════════════════════
      // STEP 7: FINALIZE
      // ═══════════════════════════════════════════════════════════════════
      const totalSlides = presentation.getSlides().length;
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      console.log(`\n${'='.repeat(80)}`);
      console.log(`✅ COMPLETE - ${totalSlides} slides in ${duration}s`);
      console.log(`${'='.repeat(80)}\n`);
      
      ss.toast('✅ Predictability Report created!', 'Success', 5);
      
      // Show success popup
      const htmlOutput = HtmlService.createHtmlOutput(
        `<div style="font-family: Arial, sans-serif; padding: 20px;">
          <h2 style="color: #502D7F; margin-top: 0;">✅ Success!</h2>
          <p style="margin: 25px 0;">
            <a href="${presentationUrl}" target="_blank" style="display: inline-block; background: #502D7F; color: white; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold;">📊 Open Presentation</a>
          </p>
          <div style="background: #f5f5f5; padding: 15px; border-radius: 6px;">
            <p style="margin: 5px 0;"><strong>Total Slides:</strong> ${totalSlides}</p>
            <p style="margin: 5px 0;"><strong>Deferral Slides:</strong> ${deferralSlides}</p>
            <p style="margin: 5px 0;"><strong>Value Stream Slides:</strong> ${valueStreamSlides}</p>
            <p style="margin: 5px 0;"><strong>Duration:</strong> ${duration}s</p>
          </div>
        </div>`
      ).setWidth(500).setHeight(350);
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Predictability Report Created');
      
      return { url: presentationUrl, slideCount: totalSlides };
      
    } catch (error) {
      console.error(`❌ ERROR: ${error.message}`);
      ui.alert('Error', `Failed to generate report:\n\n${error.message}`, ui.ButtonSet.OK);
      throw error;
    }
  }, { piNumber: piNumber });
}

// ════════════════════════════════════════════════════════════════════════════════════════════════════════
// extractDeferredIssuesFromSheet - UPDATED with Allocation and Objective Not Met
// ════════════════════════════════════════════════════════════════════════════════════════════════════════

function extractDeferredIssuesFromSheet(sourceSheet, sortBy, includeObjectiveNotMet) {
  console.log('\n📋 Extracting deferred issues from sheet...');
  console.log(`   Sort by: ${sortBy}`);
  console.log(`   Include Objective Not Met: ${includeObjectiveNotMet !== false}`);
  
  // Default includeObjectiveNotMet to true if not specified
  if (includeObjectiveNotMet === undefined) {
    includeObjectiveNotMet = true;
  }
  
  // ════════════════════════════════════════════════════════════════════════
  // CONFIGURATION
  // ════════════════════════════════════════════════════════════════════════
  const HEADER_ROW = 4;
  const DATA_START_ROW = 5;
  
  const lastCol = sourceSheet.getLastColumn();
  const lastRow = sourceSheet.getLastRow();
  
  if (lastRow < DATA_START_ROW) {
    console.log('   No data rows found');
    return [];
  }
  
  // Get headers from row 4
  const headers = sourceSheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(h => h.toString().toLowerCase().trim());
  
  // Get data starting from row 5
  const dataRows = sourceSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol).getValues();
  
  console.log(`   Total data rows: ${dataRows.length}`);
  console.log(`   Headers: ${headers.filter(h => h).join(', ')}`);
  
  // ════════════════════════════════════════════════════════════════════════
  // COLUMN DETECTION - Helper function
  // ════════════════════════════════════════════════════════════════════════
  function findColumn(possibleNames) {
    for (const name of possibleNames) {
      const idx = headers.indexOf(name);
      if (idx >= 0) return idx;
    }
    return -1;
  }
  
  const colIndices = {
    // ✅ CRITICAL: Issue Type column for filtering Epics only
    issueType: findColumn(['issue type', 'issuetype', 'type']),
    
    piCommitment: findColumn(['pi commitment']),
    programInitiative: findColumn(['program initiative']),
    portfolioInitiative: findColumn(['portfolio initiative']),
    summary: findColumn(['summary']),
    valuestream: findColumn(['valuestream', 'value stream', 'value stream/org']),
    aiRagNote: findColumn(['rag note - ai summary', 'ai rag note', 'rca deferral ai summary']),
    issueKey: findColumn(['issue key', 'key']),
    status: findColumn(['status']),
    piObjectiveStatus: findColumn(['pi objective status', 'objective status']),
    allocation: findColumn(['allocation'])
  };
  
  // ════════════════════════════════════════════════════════════════════════
  // LOGGING COLUMN INDICES
  // ════════════════════════════════════════════════════════════════════════
  console.log('\n   Column indices found:');
  Object.keys(colIndices).forEach(key => {
    const idx = colIndices[key];
    const status = idx >= 0 ? `✓ Column ${idx} ("${headers[idx]}")` : '✗ NOT FOUND';
    console.log(`      ${key}: ${status}`);
  });
  
  // ════════════════════════════════════════════════════════════════════════
  // VALIDATION - Critical columns
  // ════════════════════════════════════════════════════════════════════════
  if (colIndices.issueType === -1) {
    console.error('\n   ❌ CRITICAL: "Issue Type" column not found!');
    console.error('   Cannot filter for Epics only. Check sheet headers.');
    console.error('   Expected column names: "Issue Type", "IssueType", or "Type"');
    throw new Error('Required column "Issue Type" not found in sheet. Cannot filter for Epics.');
  }
  
  if (colIndices.piCommitment === -1) {
    console.error('   ❌ "PI Commitment" column not found!');
    throw new Error('Required column "PI Commitment" not found in sheet');
  }
  
  console.log('\n   ✅ Issue Type filter ENABLED - Only Epics will be included');
  
  // ════════════════════════════════════════════════════════════════════════
  // DATA EXTRACTION
  // ════════════════════════════════════════════════════════════════════════
  const issues = [];
  const DEFERRAL_STATUSES = ['deferred', 'canceled', 'cancelled', 'traded'];
  const CLOSED_STATUSES = ['closed', 'pending acceptance'];
  
  // Track keys to prevent duplicates
  const processedKeys = new Set();
  
  // Counters for logging
  let totalRows = 0;
  let skippedNonEpics = 0;
  let skippedNoMatch = 0;
  let skippedDuplicates = 0;
  let deferredCount = 0;
  let objectiveNotMetCount = 0;
  
  // Helper to get cell value safely
  function getVal(row, colIndex) {
    if (colIndex < 0 || colIndex >= row.length) return '';
    return (row[colIndex] || '').toString().trim();
  }
  
  console.log(`\n   ═══════════════════════════════════════════════════════════`);
  console.log(`   FILTERING CRITERIA:`);
  console.log(`   ═══════════════════════════════════════════════════════════`);
  console.log(`   Step 1: Issue Type = "Epic" (REQUIRED)`);
  console.log(`   Step 2: PI Commitment IN ["${DEFERRAL_STATUSES.join('", "')}"]`);
  if (includeObjectiveNotMet) {
    console.log(`   Step 3: OR (Status IN ["${CLOSED_STATUSES.join('", "')}"]`);
    console.log(`               AND PI Objective Status ≠ "Met")`);
  }
  console.log(`   ═══════════════════════════════════════════════════════════\n`);
  
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    totalRows++;
    
    // ════════════════════════════════════════════════════════════════════
    // ✅ FILTER 1: Issue Type MUST be "Epic"
    // ════════════════════════════════════════════════════════════════════
    const issueType = getVal(row, colIndices.issueType).toLowerCase();
    
    if (issueType !== 'epic') {
      skippedNonEpics++;
      continue;  // Skip Dependencies, Stories, Tasks, etc.
    }
    
    // Get key values
    const issueKey = getVal(row, colIndices.issueKey);
    const piCommitment = getVal(row, colIndices.piCommitment);
    const piCommitmentLower = piCommitment.toLowerCase();
    const status = getVal(row, colIndices.status);
    const statusLower = status.toLowerCase();
    const piObjectiveStatus = getVal(row, colIndices.piObjectiveStatus);
    const piObjectiveStatusLower = piObjectiveStatus.toLowerCase();
    
    // ════════════════════════════════════════════════════════════════════
    // FILTER 2: Check deferral/objective conditions
    // ════════════════════════════════════════════════════════════════════
    const isDeferred = DEFERRAL_STATUSES.some(s => piCommitmentLower.includes(s));
    const isClosedOrPending = CLOSED_STATUSES.some(s => statusLower === s);
    const isObjectiveNotMet = piObjectiveStatusLower !== 'met';  // ≠ "Met"
    
    let reason = null;
    
    // Condition 1: Deferred commitment
    if (isDeferred) {
      reason = 'Deferred';
      deferredCount++;
    }
    // Condition 2: Closed/Pending AND PI Objective Status ≠ "Met"
    else if (includeObjectiveNotMet && isClosedOrPending && isObjectiveNotMet) {
      reason = 'Objective Not Met';
      objectiveNotMetCount++;
    }
    
    // Skip if doesn't match any condition
    if (!reason) {
      skippedNoMatch++;
      continue;
    }
    
    // ════════════════════════════════════════════════════════════════════
    // FILTER 3: De-duplicate by issue key
    // ════════════════════════════════════════════════════════════════════
    if (issueKey && processedKeys.has(issueKey)) {
      skippedDuplicates++;
      continue;
    }
    if (issueKey) {
      processedKeys.add(issueKey);
    }
    
    // Build issue object
    const issue = {
      programInitiative: getVal(row, colIndices.programInitiative),
      portfolioInitiative: getVal(row, colIndices.portfolioInitiative),
      epicSummary: getVal(row, colIndices.summary),
      valuestream: getVal(row, colIndices.valuestream),
      aiRagNote: getVal(row, colIndices.aiRagNote),
      issueKey: issueKey,
      piCommitment: piCommitment,
      status: status,
      piObjectiveStatus: piObjectiveStatus,
      allocation: getVal(row, colIndices.allocation),
      reason: reason  // "Deferred" or "Objective Not Met"
    };
    
    // Only add if has meaningful data
    if (issue.programInitiative || issue.epicSummary) {
      issues.push(issue);
    }
  }
  
  // ════════════════════════════════════════════════════════════════════════
  // LOGGING RESULTS
  // ════════════════════════════════════════════════════════════════════════
  console.log(`\n   ═══════════════════════════════════════════════════════════`);
  console.log(`   EXTRACTION RESULTS:`);
  console.log(`   ═══════════════════════════════════════════════════════════`);
  console.log(`   Total rows scanned:     ${totalRows}`);
  console.log(`   Skipped (not Epic):     ${skippedNonEpics}`);
  console.log(`   Skipped (no match):     ${skippedNoMatch}`);
  console.log(`   Skipped (duplicates):   ${skippedDuplicates}`);
  console.log(`   ───────────────────────────────────────────────────────────`);
  console.log(`   Deferred Epics:         ${deferredCount}`);
  console.log(`   Objective Not Met:      ${objectiveNotMetCount}`);
  console.log(`   ───────────────────────────────────────────────────────────`);
  console.log(`   TOTAL EPICS INCLUDED:   ${issues.length}`);
  console.log(`   ═══════════════════════════════════════════════════════════\n`);
  
  if (issues.length === 0) {
    console.log('   No matching Epics found.');
    return [];
  }
  
  // ════════════════════════════════════════════════════════════════════════
  // SORTING - Custom Group Ordering for Deferrals
  // ════════════════════════════════════════════════════════════════════════
  console.log(`   Applying category-based sorting...`);
  
  /**
   * Determines the group priority for a deferral issue
   */
  function getDeferralGroupPriority(issue) {
    const portfolioInit = (issue.portfolioInitiative || '').toLowerCase();
    const programInit = (issue.programInitiative || '').toLowerCase();
    const allocation = (issue.allocation || '').toLowerCase();
    
    if (portfolioInit.includes('infosec') || programInit.includes('infosec')) return 1;
    if (allocation === 'product - compliance') return 2;
    if (allocation === 'product - feature') return 3;
    if (allocation.includes('platform')) return 4;
    return 5;  // Unknown
  }
  
  const groupNames = {
    1: 'Infosec',
    2: 'Product - Compliance',
    3: 'Product - Feature',
    4: 'Platform',
    5: 'Unknown'
  };
  
  // Get portfolio sort order
  let portfolioOrder = {};
  try {
    portfolioOrder = getPortfolioSortOrder();
  } catch (e) {
    console.log('   Using default portfolio order');
    const defaultOrder = [
      'Information Security', 'EMR Platform Integrity', 'Insights', 'Compliance',
      'Provider Experience', 'Patient Engagement', 'Cloud Migration', 'PM Rewrite', 'Practice Efficiency'
    ];
    defaultOrder.forEach((name, i) => portfolioOrder[name] = i);
  }
  
  // Sort: Category → Portfolio Initiative → Program Initiative → Value Stream
  issues.sort((a, b) => {
    // Level 1: Category/Group Priority
    const groupA = getDeferralGroupPriority(a);
    const groupB = getDeferralGroupPriority(b);
    if (groupA !== groupB) return groupA - groupB;
    
    // Level 2: Portfolio Initiative (governance order)
    const portOrderA = portfolioOrder[a.portfolioInitiative] !== undefined ? portfolioOrder[a.portfolioInitiative] : 999;
    const portOrderB = portfolioOrder[b.portfolioInitiative] !== undefined ? portfolioOrder[b.portfolioInitiative] : 999;
    if (portOrderA !== portOrderB) return portOrderA - portOrderB;
    
    // Level 3: Program Initiative (alpha)
    const progA = (a.programInitiative || '').toLowerCase();
    const progB = (b.programInitiative || '').toLowerCase();
    if (progA !== progB) return progA.localeCompare(progB);
    
    // Level 4: Value Stream (alpha)
    const vsA = (a.valuestream || '').toLowerCase();
    const vsB = (b.valuestream || '').toLowerCase();
    return vsA.localeCompare(vsB);
  });
  
  // Log group distribution
  const groupCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };
  issues.forEach(issue => {
    const group = getDeferralGroupPriority(issue);
    groupCounts[group]++;
  });
  
  console.log(`\n   ═══════════════════════════════════════════════════════════`);
  console.log(`   CATEGORY DISTRIBUTION:`);
  console.log(`   ═══════════════════════════════════════════════════════════`);
  Object.keys(groupCounts).forEach(key => {
    if (groupCounts[key] > 0) {
      console.log(`   ${groupNames[key]}: ${groupCounts[key]} epics`);
    }
  });
  console.log(`   ═══════════════════════════════════════════════════════════`);
  
  // Log sample issues
  if (issues.length > 0) {
    console.log(`\n   Sample extracted Epics:`);
    const samplesToShow = Math.min(5, issues.length);
    for (let i = 0; i < samplesToShow; i++) {
      const issue = issues[i];
      const group = getDeferralGroupPriority(issue);
      console.log(`      ${i + 1}. [${groupNames[group]}] [${issue.reason}] ${issue.issueKey}: ${(issue.epicSummary || '').substring(0, 40)}...`);
    }
    if (issues.length > samplesToShow) {
      console.log(`      ... and ${issues.length - samplesToShow} more`);
    }
  }
  
  return issues;
}

// ════════════════════════════════════════════════════════════════════════════════════════════════════════
// generateDeferralSlidesForPredictability - COMPLETE REWRITE
// ════════════════════════════════════════════════════════════════════════════════════════════════════════

function generateDeferralSlidesForPredictability(presentation, issues, deferralTemplateIndex) {
  console.log(`\n${'═'.repeat(80)}`);
  console.log(`GENERATING DEFERRAL SLIDES`);
  console.log(`${'═'.repeat(80)}`);
  console.log(`Total issues: ${issues ? issues.length : 0}`);
  
  if (!issues || issues.length === 0) return 0;
  
  const slides = presentation.getSlides();
  
  // Find template slide
  let deferralTemplate;
  let templateIndex;
  
  if (deferralTemplateIndex !== undefined && deferralTemplateIndex < slides.length) {
    deferralTemplate = slides[deferralTemplateIndex];
    templateIndex = deferralTemplateIndex;
  } else {
    const result = findDeferralTemplateSlide(presentation);
    deferralTemplate = result.slide;
    templateIndex = result.index;
  }
  
  if (!deferralTemplate) throw new Error('Deferral template not found');
  
  console.log(`Template index: ${templateIndex}`);
  
  // Remove header slide if present
  const allSlides = presentation.getSlides();
  if (allSlides.length > 1) {
    const headerSlide = allSlides[1];
    const shapes = headerSlide.getShapes();
    for (const shape of shapes) {
      try {
        const text = shape.getText().asString();
        if (text.includes('PI Commitment Changelog') && text.includes('Deferred')) {
          console.log(`🗑️  Removing header slide`);
          headerSlide.remove();
          if (templateIndex > 1) {
            templateIndex--;
            deferralTemplate = presentation.getSlides()[templateIndex];
          }
          break;
        }
      } catch (e) {}
    }
  }
  
  // Group issues by category
  const categorizedIssues = {};
  DEFERRAL_CATEGORIES.forEach(cat => categorizedIssues[cat] = []);
  
  issues.forEach(issue => {
    const category = getDeferralCategory(issue);
    categorizedIssues[category].push(issue);
  });
  
  // Log distribution
  console.log(`\n📊 Category Distribution:`);
  DEFERRAL_CATEGORIES.forEach(cat => {
    if (categorizedIssues[cat].length > 0) {
      console.log(`   ${cat}: ${categorizedIssues[cat].length} issues`);
    }
  });
  
  // Generate slides for each category
  let slidesCreated = 0;
  let insertPosition = templateIndex;
  
  for (const category of DEFERRAL_CATEGORIES) {
    const categoryIssues = categorizedIssues[category];
    
    if (categoryIssues.length === 0) {
      console.log(`\n⏭️  Skipping ${category} (no issues)`);
      continue;
    }
    
    console.log(`\n📂 Processing: ${category} (${categoryIssues.length} issues)`);
    
    // Sort within category
    sortIssuesWithinCategory(categoryIssues);
    
    const slidesForCategory = Math.ceil(categoryIssues.length / ISSUES_PER_SLIDE);
    
    for (let slideNum = 0; slideNum < slidesForCategory; slideNum++) {
      const startIdx = slideNum * ISSUES_PER_SLIDE;
      const endIdx = Math.min(startIdx + ISSUES_PER_SLIDE, categoryIssues.length);
      const issuesBatch = categoryIssues.slice(startIdx, endIdx);
      
      console.log(`   📄 Slide ${slideNum + 1}: Issues ${startIdx + 1}-${endIdx}`);
      
      // Duplicate template
      const newSlide = deferralTemplate.duplicate();
      
      // Update title
      updateDeferralSlideTitle(newSlide, category);
      
      // Populate table
      try {
        populateDeferralTableForSlide(newSlide, issuesBatch);
      } catch (error) {
        console.error(`   ❌ Failed: ${error.message}`);
        throw error;
      }
      
      // Move slide
      insertPosition++;
      try {
        newSlide.move(insertPosition);
      } catch (error) {
        console.warn(`   ⚠️ Move failed: ${error.message}`);
      }
      
      slidesCreated++;
    }
  }
  
  // Remove template
  console.log(`\n🗑️  Removing template slide`);
  deferralTemplate.remove();
  
  console.log(`\n✅ Created ${slidesCreated} deferral slides`);
  return slidesCreated;
}

function generateValueStreamSlides(presentation, piNumber, predScoreSheet) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`GENERATING VALUE STREAM SLIDES`);
  console.log(`${'='.repeat(80)}\n`);
  
  const slides = presentation.getSlides();
  
  let templateSlide = null;
  let templateSlideIndex = -1;
  
  console.log(`📋 Searching for value stream template in ${slides.length} slides...`);
  
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    const shapes = slide.getShapes();
    
    for (let j = 0; j < shapes.length; j++) {
      try {
        const text = shapes[j].getText().asString();
        
        if (text.includes('{{Value Stream}}') || text.includes('{{PI #}}')) {
          templateSlide = slide;
          templateSlideIndex = i;
          console.log(`✓ Found value stream template at slide ${i + 1} (0-indexed: ${i})`);
          break;
        }
      } catch (e) {}
    }
    
    if (templateSlide) break;
  }
  
  if (!templateSlide) {
    console.error('❌ ERROR: Could not find value stream template slide!');
    return 0;
  }
  
  console.log(`✓ Using slide ${templateSlideIndex + 1} as template\n`);
  
  let slidesCreated = 0;
  
  for (let i = 0; i < PREDICTABILITY_VALUE_STREAMS.length; i++) {
    const valueStream = PREDICTABILITY_VALUE_STREAMS[i];
    console.log(`\n📊 Processing: ${valueStream} (Index: ${i})`);
    
    try {
      // Duplicate template
      const newSlide = templateSlide.duplicate();
      
      // ✅ FIX: Move THEN get fresh reference by ID, not by position
      const targetPosition = templateSlideIndex + 1 + i;
      newSlide.move(targetPosition);
      
      // ✅ CRITICAL: Get the slide object ID before it moves
      const slideId = newSlide.getObjectId();
      
      // ✅ Get fresh reference using the ID
      const allSlides = presentation.getSlides();
      let workingSlide = null;
      for (let j = 0; j < allSlides.length; j++) {
        if (allSlides[j].getObjectId() === slideId) {
          workingSlide = allSlides[j];
          break;
        }
      }
      
      if (!workingSlide) {
        console.error(`  ❌ Could not find working slide after move`);
        continue;
      }
      
      console.log(`  📍 Created slide at position ${targetPosition}`);
      
      addDisclosureToSlide(workingSlide);
      robustReplaceAllText(workingSlide, '{{Value Stream}}', valueStream);
      robustReplaceAllText(workingSlide, '{{PI #}}', `PI ${piNumber}`);
      setSlideTitle(workingSlide, `${valueStream} - PI ${piNumber} Program Predictability`);
      
      populateValueStreamChartsWithData(workingSlide, valueStream, piNumber, predScoreSheet, i);
      slidesCreated++;
      console.log(`  ✓ Slide completed: "${valueStream} - PI ${piNumber} Program Predictability"`);
      
    } catch (error) {
      console.error(`  ❌ Error creating slide for ${valueStream}: ${error.message}`);
      console.error(`  Stack trace: ${error.stack}`);
    }
  }
  
  templateSlide.remove();
  console.log(`\n✓ Removed template slide`);
  console.log(`✓ Created ${slidesCreated} value stream slides`);
  
  return slidesCreated;
}

function setSlideTitle(slide, titleText) {
  const shapes = slide.getShapes();
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    
    try {
      const text = shape.getText().asString();
      
      if (text.includes('{{Value Stream}}') || 
          text.includes('{{PI #}}') || 
          text.includes('Program Predictability')) {
        
        shape.getText().setText(titleText);
        
        shape.getText().getTextStyle()
          .setFontFamily('Arial')
          .setFontSize(24)
          .setBold(true)
          .setForegroundColor('#502D7F');
        
        console.log(`    ✓ Title set: "${titleText}"`);
        return true;
      }
    } catch (e) {
      // Skip shapes without text
    }
  }
  
  console.warn(`    ⚠️ Could not find title shape to update`);
  return false;
}
function populateValueStreamChartsWithData(slide, valueStream, piNumber, predScoreSheet, vsIndex) {
  console.log(`    📈 Generating charts for ${valueStream} (vsIndex: ${vsIndex})...`);
  
  try {
    drawCapacityVelocityChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex);
    drawAllocationChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex);
    drawEpicStatusChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex);
    drawVelocityTrendChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex);
    drawProgramPredictabilityChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex);
    drawBVAVSummaryChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex, piNumber - 11);
    
    console.log(`    ✅ All charts generated for ${valueStream}`);
  } catch (error) {
    console.error(`    ❌ Error generating charts for ${valueStream}: ${error.message}`);
    console.error(`    Stack trace: ${error.stack}`);
    throw error;
  }
}

/**
 * CHART 1: Capacity & Velocity Boxes
 */
function drawCapacityVelocityChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex) {
  console.log(`      📊 Chart 1: Capacity & Velocity`);
  
  try {
    // ✅ FIXED: Capacity column varies by PI (2 columns per PI)
    const capacityRow = 4 + vsIndex;
    const capacityCol = 2 + ((piNumber - 11) * 2);  // PI 11→2, PI 12→4, PI 13→6
    const capacity = parseFloat(predScoreSheet.getRange(capacityRow, capacityCol).getValue()) || 0;
    
    // ✅ FIXED: Velocity - 1 column per PI starting at B for PI 1
    const velocityRow = 41 + vsIndex;
    const velocityCol = 1 + piNumber;  // PI 1→2(B), PI 13→14(N)
    const velocity = parseFloat(predScoreSheet.getRange(velocityRow, velocityCol).getValue()) || 0;
    
    console.log(`        Capacity: ${capacity} (col ${capacityCol}), Velocity: ${velocity} (col ${velocityCol})`);
    
    const placeholder = findPlaceholder(slide, '{{Capacity & Velocity Chart}}');
    if (!placeholder) {
      console.warn(`        ⚠️  Placeholder not found`);
      return;
    }
    
    const x = placeholder.left;
    const y = placeholder.top;
    const width = placeholder.width;
    const height = placeholder.height;
    
    placeholder.shape.remove();
    
    const boxHeight = height / 2 - 5;
    const boxWidth = Math.min(width, 300);
    
    // CAPACITY BOX (Top)
    const capacityBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, boxWidth, boxHeight);
    capacityBox.getFill().setSolidFill('#9575cd');
    capacityBox.getBorder().setWeight(3);
    capacityBox.getBorder().getLineFill().setSolidFill('#FFC82E');
    
    const capacityText = capacityBox.getText();
    capacityText.setText(`${Math.round(capacity)}\nCapacity`);
    capacityText.getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(14)
      .setBold(true)
      .setForegroundColor('#ffffff');
    capacityText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // VELOCITY BOX (Bottom)
    const velocityBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y + boxHeight + 10, boxWidth, boxHeight);
    velocityBox.getFill().setSolidFill('#b39ddb');
    velocityBox.getBorder().setWeight(3);
    velocityBox.getBorder().getLineFill().setSolidFill('#FFC82E');
    
    const velocityText = velocityBox.getText();
    velocityText.setText(`${Math.round(velocity)}\nVelocity`);
    velocityText.getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(14)
      .setBold(true)
      .setForegroundColor('#ffffff');
    velocityText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    console.log(`        ✅ Created capacity/velocity boxes (double width, font 14)`);
  } catch (error) {
    console.error(`        ❌ Error: ${error.message}`);
    throw error;
  }
}

/**
 * CHART 2: Allocation Pie Chart
 */
function drawAllocationChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex) {
  try {
    console.log(`      📊 Chart 2: Allocation Pie Chart`);
    console.log(`        🔍 Looking for PIE chart for ${valueStream} (vsIndex=${vsIndex})`);
    
    // ✅ FIX: Use SECTION_HEIGHT = 26, not 25!
    const SECTION_HEIGHT = 26;
    const sectionStartRow = 2 + (vsIndex * SECTION_HEIGHT);
    console.log(`        📍 Expected position: row ${sectionStartRow}, cols 4-5`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const chartSheetName = `PI ${piNumber} - VS Charts`;
    const chartSheet = ss.getSheetByName(chartSheetName);
    
    if (!chartSheet) {
      console.log(`        ❌ Chart sheet not found: ${chartSheetName}`);
      return;
    }
    
    const charts = chartSheet.getCharts();
    console.log(`        📊 Total charts on sheet: ${charts.length}`);
    
    const chart = findChartByPosition(chartSheet, sectionStartRow, [4, 5]);
    
    if (!chart) {
      console.log(`        ❌ PIE chart not found at expected position`);
      return;
    }
    
    console.log(`        ✅ Found correct pie chart at row ${sectionStartRow}`);
    
    const shapes = slide.getShapes();
    let placeholder = null;
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      const text = shape.getText().asString();
      if (text.includes('{{Allocation Chart}}')) {
        placeholder = shape;
        break;
      }
    }
    
    if (!placeholder) {
      console.log(`        ⚠️  Allocation chart placeholder not found`);
      return;
    }
    
    const left = placeholder.getLeft();
    const top = placeholder.getTop();
    const width = placeholder.getWidth();
    const height = placeholder.getHeight();
    
    const chartBlob = chart.getAs('image/png');
    const image = slide.insertImage(chartBlob, left, top, width, height);
    
    placeholder.remove();
    
    console.log(`        ✅ Embedded pie chart for ${valueStream}`);
    
  } catch (error) {
    console.log(`        ❌ Error: ${error.message}`);
    throw error;
  }
}


/**
 * CHART 3: Epic Status Boxes
 */
function drawEpicStatusChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex) {
  console.log(`      📊 Chart 3: Epic Status`);
  
  try {
    const piOffset = piNumber - 11;
    const epicRow = 61 + vsIndex;
    const epicStartCol = 2 + (piOffset * 3);
    
    const committed = predScoreSheet.getRange(epicRow, epicStartCol).getValue() || 0;
    const deferred = predScoreSheet.getRange(epicRow, epicStartCol + 1).getValue() || 0;
    
    console.log(`        Committed: ${committed}, Deferred: ${deferred}`);
    
    const placeholder = findPlaceholder(slide, '{{Epic Status Chart}}');
    if (!placeholder) {
      console.warn(`        ⚠️  Placeholder not found`);
      return;
    }
    
    const x = placeholder.left;
    const y = placeholder.top;
    const width = placeholder.width;
    const height = placeholder.height;
    
    placeholder.shape.remove();
    
    const boxHeight = height / 2 - 5;
    // ✅ CHANGED: Double the width (removed the /2 and increased max)
    const boxWidth = Math.min(width, 300);
    
    // COMMITTED BOX (Top)
    const committedBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, boxWidth, boxHeight);
    committedBox.getFill().setSolidFill('#7e57c2');
    committedBox.getBorder().setWeight(3);
    committedBox.getBorder().getLineFill().setSolidFill('#FFC82E');
    
    const committedText = committedBox.getText();
    committedText.setText(`${committed}\nFeatures Committed`);
    committedText.getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)  // ✅ CHANGED: 15 → 11 (reduced by 4)
      .setBold(true)
      .setForegroundColor('#ffffff');
    committedText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // DEFERRED BOX (Bottom)
    const deferredBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y + boxHeight + 10, boxWidth, boxHeight);
    deferredBox.getFill().setSolidFill('#9575cd');
    deferredBox.getBorder().setWeight(3);
    deferredBox.getBorder().getLineFill().setSolidFill('#FFC82E');
    
    const deferredText = deferredBox.getText();
    deferredText.setText(`${deferred}\nFeatures Deferred`);
    deferredText.getTextStyle()
      .setFontFamily('Arial')
      .setFontSize(11)  // ✅ CHANGED: 15 → 11 (reduced by 4)
      .setBold(true)
      .setForegroundColor('#ffffff');
    deferredText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    console.log(`        ✅ Created epic status boxes (double width, font 11)`);
  } catch (error) {
    console.error(`        ❌ Error: ${error.message}`);
    throw error;
  }
}

/**
 * CHART 4: Velocity Trending Line Chart
 */
function drawVelocityTrendChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex) {
  try {
    console.log(`      📊 Chart 4: Velocity Trending Chart`);
    console.log(`        🔍 Looking for VELOCITY chart for ${valueStream} (vsIndex=${vsIndex})`);
    
    // ✅ FIX: Use SECTION_HEIGHT = 26, not 25!
    const SECTION_HEIGHT = 26;
    const sectionStartRow = 2 + (vsIndex * SECTION_HEIGHT);
    console.log(`        📍 Expected position: row ${sectionStartRow}, cols 15-16`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const chartSheetName = `PI ${piNumber} - VS Charts`;
    const chartSheet = ss.getSheetByName(chartSheetName);
    
    if (!chartSheet) {
      console.log(`        ❌ Chart sheet not found: ${chartSheetName}`);
      return;
    }
    
    const chart = findChartByPosition(chartSheet, sectionStartRow, [15, 16]);
    
    if (!chart) {
      console.log(`        ❌ VELOCITY chart not found at expected position`);
      return;
    }
    
    console.log(`        ✅ Found correct velocity chart at row ${sectionStartRow}`);
    
    const shapes = slide.getShapes();
    let placeholder = null;
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      const text = shape.getText().asString();
      if (text.includes('{{Velocity Trending Chart}}')) {
        placeholder = shape;
        break;
      }
    }
    
    if (!placeholder) {
      console.log(`        ⚠️  Velocity chart placeholder not found`);
      return;
    }
    
    const left = placeholder.getLeft();
    const top = placeholder.getTop();
    const width = placeholder.getWidth();
    const height = placeholder.getHeight();
    
    const chartBlob = chart.getAs('image/png');
    const image = slide.insertImage(chartBlob, left, top, width, height);
    
    placeholder.remove();
    
    console.log(`        ✅ Embedded velocity chart for ${valueStream}`);
    
  } catch (error) {
    console.log(`        ❌ Error: ${error.message}`);
    throw error;
  }
}

/**
 * CHART 5: Program Predictability Line Chart
 */
function drawProgramPredictabilityChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex) {
  try {
    console.log(`      📊 Chart 5: Program Predictability Chart`);
    console.log(`        🔍 Looking for PREDICTABILITY chart for ${valueStream} (vsIndex=${vsIndex})`);
    
    // ✅ FIX: Use SECTION_HEIGHT = 26, not 25!
    const SECTION_HEIGHT = 26;
    const sectionStartRow = 2 + (vsIndex * SECTION_HEIGHT);
    console.log(`        📍 Expected position: row ${sectionStartRow}, cols 21-22`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const chartSheetName = `PI ${piNumber} - VS Charts`;
    const chartSheet = ss.getSheetByName(chartSheetName);
    
    if (!chartSheet) {
      console.log(`        ❌ Chart sheet not found: ${chartSheetName}`);
      return;
    }
    
    const chart = findChartByPosition(chartSheet, sectionStartRow, [21, 22]);
    
    if (!chart) {
      console.log(`        ❌ PREDICTABILITY chart not found at expected position`);
      return;
    }
    
    console.log(`        ✅ Found correct predictability chart at row ${sectionStartRow}`);
    
    const shapes = slide.getShapes();
    let placeholder = null;
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      const text = shape.getText().asString();
      if (text.includes('{{Program Predictability Chart}}')) {
        placeholder = shape;
        break;
      }
    }
    
    if (!placeholder) {
      console.log(`        ⚠️  Predictability chart placeholder not found`);
      return;
    }
    
    const left = placeholder.getLeft();
    const top = placeholder.getTop();
    const width = placeholder.getWidth();
    const height = placeholder.getHeight();
    
    const chartBlob = chart.getAs('image/png');
    const image = slide.insertImage(chartBlob, left, top, width, height);
    
    placeholder.remove();
    
    console.log(`        ✅ Embedded predictability chart for ${valueStream}`);
    
  } catch (error) {
    console.log(`        ❌ Error: ${error.message}`);
    throw error;
  }
}


/**
 * Helper: Finds a placeholder text in a slide
 */
function findPlaceholder(slide, placeholderText) {
  const shapes = slide.getShapes();
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    try {
      const text = shape.getText().asString();
      if (text.includes(placeholderText)) {
        return {
          shape: shape,
          left: shape.getLeft(),
          top: shape.getTop(),
          width: shape.getWidth(),
          height: shape.getHeight()
        };
      }
    } catch (e) {}
  }
  
  return null;
}
/**
 * ════════════════════════════════════════════════════════════════════════════
 * VALUE STREAM CHART GENERATION
 * ════════════════════════════════════════════════════════════════════════════
/**
 * Populates charts on a value stream slide
 */

/**
 * CHART 6: BV/AV Summary Table
 * Structure: Header → Allocation Groups (yellow) → Initiatives (white) → Total → PI Score
 */
function drawBVAVSummaryChartEmbedded(slide, valueStream, piNumber, predScoreSheet, vsIndex, piIndex) {
  const TABLE_LEFT = 480;
  const TABLE_TOP = 50;
  const TABLE_WIDTH = 260;
  const TABLE_HEIGHT = 280;

  try {
    console.log(`    Reading BV/AV data from Iteration 6 sheet...`);
    console.log(`    Looking for Value Stream: "${valueStream}" (vsIndex: ${vsIndex})`);
    
    const bvavData = getBVAVSummaryData(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`PI ${piNumber} - Iteration 6`),
      piNumber,
      valueStream,
      predScoreSheet,
      vsIndex,
      piIndex
    );
    
    if (!bvavData || !bvavData.allocations || bvavData.allocations.length === 0) {
      console.log(`        ⚠️  No BV/AV data for ${valueStream}`);
      return;
    }

    // Calculate total rows needed:
    // 1 header + allocation rows (with subtotals) + initiative rows + 1 total + 1 PI score
    let totalInitiatives = 0;
    bvavData.allocations.forEach(allocGroup => {
      totalInitiatives += allocGroup.initiatives.length;
    });
    const numRows = 1 + bvavData.allocations.length + totalInitiatives + 2;
    
    console.log(`        Creating table with ${numRows} rows`);
    
    const table = slide.insertTable(numRows, 3, TABLE_LEFT, TABLE_TOP, TABLE_WIDTH, TABLE_HEIGHT);
    
    // Set column widths
    table.getColumn(0).getTableColumnProperties().setColumnWidth(200);
    table.getColumn(1).getTableColumnProperties().setColumnWidth(30);
    table.getColumn(2).getTableColumnProperties().setColumnWidth(30);
    
    let rowIndex = 0;
    
    // Header row
    const headerRow = table.getRow(rowIndex++);
    headerRow.getCell(0).getText().setText('Program Initiative');
    headerRow.getCell(1).getText().setText('BV');
    headerRow.getCell(2).getText().setText('AV');
    
    for (let col = 0; col < 3; col++) {
      const cell = headerRow.getCell(col);
      cell.getFill().setSolidFill('#7030A0');
      const textStyle = cell.getText().getTextStyle();
      textStyle.setForegroundColor('#FFFFFF');
      textStyle.setBold(true);
      textStyle.setFontSize(11);
    }
    
    // Allocation groups with subtotals
    bvavData.allocations.forEach(allocGroup => {
      // Allocation row with SUBTOTALS
      const allocRow = table.getRow(rowIndex++);
      allocRow.getCell(0).getText().setText(allocGroup.allocation);
      allocRow.getCell(1).getText().setText(String(allocGroup.subtotalBV));
      allocRow.getCell(2).getText().setText(String(allocGroup.subtotalAV));
      
      for (let col = 0; col < 3; col++) {
        const cell = allocRow.getCell(col);
        cell.getFill().setSolidFill('#FFF2CC');
        cell.getText().getTextStyle().setBold(true);
        cell.getText().getTextStyle().setFontSize(10);
      }
      
      // Initiative rows
      allocGroup.initiatives.forEach(init => {
        const initRow = table.getRow(rowIndex++);
        initRow.getCell(0).getText().setText('   ' + init.name);
        initRow.getCell(1).getText().setText(String(init.bv));
        initRow.getCell(2).getText().setText(String(init.av));
        
        for (let col = 0; col < 3; col++) {
          const cell = initRow.getCell(col);
          cell.getFill().setSolidFill('#FFFFFF');
          cell.getText().getTextStyle().setFontSize(9);
        }
        
        // Highlight red if AV < BV
        if (init.av < init.bv) {
          for (let col = 0; col < 3; col++) {
            initRow.getCell(col).getFill().setSolidFill('#FFCCCC');
          }
        }
      });
    });
    
    // Total row (sum of allocation subtotals)
    const totalRow = table.getRow(rowIndex++);
    totalRow.getCell(0).getText().setText('Total');
    totalRow.getCell(1).getText().setText(String(bvavData.totalBV));
    totalRow.getCell(2).getText().setText(String(bvavData.totalAV));
    
    for (let col = 0; col < 3; col++) {
      const cell = totalRow.getCell(col);
      cell.getFill().setSolidFill('#B4C7E7');
      cell.getText().getTextStyle().setBold(true);
      cell.getText().getTextStyle().setFontSize(11);
    }
    
    // PI Score row
    const scoreRow = table.getRow(rowIndex);
    
    try {
      scoreRow.getCell(0).merge(scoreRow.getCell(1));
    } catch (e) {
      console.log(`        ⚠️  Cell merge failed (non-critical): ${e.message}`);
    }
    
    scoreRow.getCell(0).getText().setText('Program Predictability Score');
    scoreRow.getCell(2).getText().setText(`${bvavData.piScore.toFixed(1)}%`);
    
    scoreRow.getCell(2).getFill().setSolidFill('#C6E0B4');
    scoreRow.getCell(2).getText().getTextStyle().setBold(true);
    scoreRow.getCell(2).getText().getTextStyle().setFontSize(12);
    
    console.log(`        ✅ Created BV/AV summary table for ${valueStream}`);
    
  } catch (error) {
    console.log(`        ❌ Error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
  }
}


/**
 * Helper to replace chart placeholder with text (for now)
 */
function replaceChartPlaceholder(slide, placeholder, description) {
  const shapes = slide.getShapes();
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    try {
      const text = shape.getText().asString();
      if (text.includes(placeholder)) {
        // For now, replace with description text
        // In full implementation, this would create actual charts
        shape.getText().setText(description);
        console.log(`    ✓ Replaced: ${placeholder}`);
        return;
      }
    } catch (e) {
      // Skip shapes without text
    }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTION: ROBUST TEXT REPLACEMENT
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Robustly replaces text in all shapes and tables on a slide
 * (Include this if not already present in your codebase)
 */
function robustReplaceAllText(slide, placeholder, replacement) {
  let replacementCount = 0;
  
  // Replace in all shapes
  const shapes = slide.getShapes();
  for (let i = 0; i < shapes.length; i++) {
    try {
      const shape = shapes[i];
      const textRange = shape.getText();
      
      if (textRange && textRange.asString().includes(placeholder)) {
        textRange.replaceAllText(placeholder, replacement);
        replacementCount++;
      }
    } catch (e) {
      // Skip shapes without text
    }
  }
  
  // Replace in all tables
  const tables = slide.getTables();
  for (let i = 0; i < tables.length; i++) {
    try {
      const table = tables[i];
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();
      
      for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
          try {
            const cell = table.getCell(row, col);
            const textRange = cell.getText();
            
            if (textRange && textRange.asString().includes(placeholder)) {
              textRange.replaceAllText(placeholder, replacement);
              replacementCount++;
            }
          } catch (e) {
            // Skip cells that can't be accessed
          }
        }
      }
    } catch (e) {
      // Skip tables that can't be accessed
    }
  }
  
  return replacementCount;
}


/**
 * Shows dialog for generating Predictability Report presentation
 */
function showPredictabilityPresentationDialog() {
  const html = getPredictabilityPresentationDialogHTML();
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(520)
    .setHeight(620);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Generate Predictability Presentation');
}

function getPredictabilityPresentationDialogHTML() {
  const availablePIs = getAvailablePIsForPredictability();
  
  if (availablePIs.length === 0) {
    return `
      <!DOCTYPE html>
      <html>
        <head>
          <style>
            body { 
              font-family: 'Google Sans', Arial, sans-serif; 
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
          <h2>No Iteration 6 Data Found</h2>
          <p>Predictability presentations require Iteration 6 data.</p>
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
            font-family: 'Google Sans', Arial, sans-serif; 
            padding: 20px; 
            background: #fff; 
          }
          .container { max-width: 480px; margin: 0 auto; }
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
            margin-bottom: 16px; 
          }
          .section-title { 
            color: #3c4043; 
            font-weight: 500; 
            margin-bottom: 12px; 
            font-size: 13px; 
            text-transform: uppercase; 
            letter-spacing: 0.5px; 
          }
          .pi-select { 
            width: 100%; 
            padding: 12px; 
            font-size: 16px; 
            border: 2px solid #dadce0; 
            border-radius: 8px; 
            background: white; 
          }
          .pi-select:focus { 
            outline: none; 
            border-color: #1a73e8; 
          }
          .checkbox-option {
            display: flex;
            align-items: flex-start;
            padding: 12px;
            background: white;
            border: 2px solid #dadce0;
            border-radius: 6px;
            cursor: pointer;
            margin-bottom: 10px;
            transition: all 0.2s;
          }
          .checkbox-option:hover {
            border-color: #1a73e8;
            background: #f8f9fa;
          }
          .checkbox-option:last-child {
            margin-bottom: 0;
          }
          .checkbox-option input[type="checkbox"] {
            margin-right: 12px;
            margin-top: 2px;
            width: 18px;
            height: 18px;
            cursor: pointer;
          }
          .checkbox-label {
            flex: 1;
          }
          .checkbox-label strong {
            display: block;
            font-size: 14px;
            color: #3c4043;
            margin-bottom: 2px;
          }
          .checkbox-label span {
            font-size: 12px;
            color: #5f6368;
            line-height: 1.4;
          }
          .info-box {
            background: #e8f0fe;
            border-left: 4px solid #1a73e8;
            padding: 12px 14px;
            margin: 16px 0;
            border-radius: 0 6px 6px 0;
          }
          .info-box p {
            color: #1967d2;
            font-size: 13px;
            margin: 4px 0;
          }
          .info-box ul {
            color: #1967d2;
            font-size: 12px;
            margin: 8px 0 0 18px;
            line-height: 1.6;
          }
          .button-container { 
            display: flex; 
            gap: 12px; 
            justify-content: flex-end; 
            margin-top: 20px; 
          }
          button { 
            padding: 10px 24px; 
            font-size: 14px; 
            border-radius: 6px; 
            border: none; 
            cursor: pointer; 
            font-weight: 500; 
            transition: all 0.2s; 
          }
          .primary-button { 
            background: #1a73e8; 
            color: white; 
          }
          .primary-button:hover:not(:disabled) { 
            background: #1557b0; 
          }
          .primary-button:disabled { 
            background: #94c1f5; 
            cursor: not-allowed; 
          }
          .secondary-button { 
            background: #fff; 
            color: #5f6368; 
            border: 1px solid #dadce0; 
          }
          .secondary-button:hover { 
            background: #f8f9fa; 
          }
          .loading { 
            display: none; 
            text-align: center; 
            padding: 30px 20px; 
          }
          .spinner { 
            border: 3px solid #f3f3f3; 
            border-top: 3px solid #1a73e8; 
            border-radius: 50%; 
            width: 40px; 
            height: 40px; 
            animation: spin 1s linear infinite; 
            margin: 0 auto 16px; 
          }
          .loading p {
            color: #5f6368;
            font-size: 14px;
          }
          @keyframes spin { 
            0% { transform: rotate(0deg); } 
            100% { transform: rotate(360deg); } 
          }
          .error { 
            background: #fce8e6;
            color: #c5221f; 
            font-size: 13px; 
            padding: 12px;
            border-radius: 6px;
            margin-top: 12px; 
            display: none; 
          }
          .error.show { 
            display: block; 
          }
          .form-content.hidden {
            display: none;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Generate Predictability Presentation</h2>
          <p class="subtitle">Create PI predictability report with deferral analysis</p>
          
          <div id="formContent" class="form-content">
            <!-- PI Selection -->
            <div class="section">
              <div class="section-title">Program Increment</div>
              <select id="piSelect" class="pi-select">
                <option value="">-- Select PI --</option>
                ${piOptions}
              </select>
            </div>
            
            <!-- Deferral Options -->
            <div class="section">
              <div class="section-title">Deferral Report Options</div>
              
              <div class="checkbox-option" onclick="toggleCheckbox('includeObjectiveNotMet')">
                <input type="checkbox" id="includeObjectiveNotMet" checked>
                <div class="checkbox-label">
                  <strong>Include "Objective Not Met" Epics</strong>
                  <span>Include closed epics where PI Objective Status ≠ "Met" (in addition to Deferred/Canceled/Traded)</span>
                </div>
              </div>
            </div>
            
            <!-- Chart Options -->
            <div class="section">
              <div class="section-title">Chart Generation</div>
              
              <div class="checkbox-option" onclick="toggleCheckbox('regenerateCharts')">
                <input type="checkbox" id="regenerateCharts">
                <div class="checkbox-label">
                  <strong>Regenerate Charts</strong>
                  <span>Recreate the VS Charts sheet from scratch. Leave unchecked to use existing charts (faster).</span>
                </div>
              </div>
            </div>
            
            <!-- Info Box -->
            <div class="info-box">
              <p><strong>📊 Deferral Categories (in order):</strong></p>
              <ul>
                <li><strong>Infosec</strong> - Initiative contains "Infosec"</li>
                <li><strong>Product - Compliance</strong> - Allocation match</li>
                <li><strong>Product - Feature</strong> - Allocation match</li>
                <li><strong>Platform</strong> - Allocation contains "Platform"</li>
                <li><strong>Unknown</strong> - Everything else</li>
              </ul>
            </div>
            
            <!-- Buttons -->
            <div class="button-container">
              <button type="button" class="secondary-button" onclick="google.script.host.close()">
                Cancel
              </button>
              <button type="button" class="primary-button" id="generateBtn" onclick="handleGenerate()">
                Generate Report
              </button>
            </div>
            
            <div class="error" id="error"></div>
          </div>
          
          <!-- Loading State -->
          <div class="loading" id="loading">
            <div class="spinner"></div>
            <p>Creating presentation...</p>
            <p style="font-size: 12px; color: #999; margin-top: 8px;">This may take 1-2 minutes</p>
          </div>
        </div>
        
        <script>
          // Toggle checkbox when clicking the row
          function toggleCheckbox(id) {
            const checkbox = document.getElementById(id);
            // Only toggle if the click wasn't directly on the checkbox
            if (event.target.type !== 'checkbox') {
              checkbox.checked = !checkbox.checked;
            }
          }
          
          function handleGenerate() {
            // Get values
            const piNumber = document.getElementById('piSelect').value;
            const includeObjectiveNotMet = document.getElementById('includeObjectiveNotMet').checked;
            const regenerateCharts = document.getElementById('regenerateCharts').checked;
            
            // Validate
            if (!piNumber) {
              showError('Please select a Program Increment');
              return;
            }
            
            // Show loading state
            document.getElementById('formContent').classList.add('hidden');
            document.getElementById('loading').style.display = 'block';
            document.getElementById('error').classList.remove('show');
            
            console.log('Calling generatePredictabilityPresentation with:');
            console.log('  piNumber:', parseInt(piNumber));
            console.log('  destinationFolderId:', null);
            console.log('  deferralSortBy:', 'category');
            console.log('  regenerateCharts:', regenerateCharts);
            console.log('  includeObjectiveNotMet:', includeObjectiveNotMet);
            
            // Call backend - 5 parameters
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .generatePredictabilityPresentation(
                parseInt(piNumber),      // 1. PI number
                null,                    // 2. Destination folder (null = same as spreadsheet)
                'category',              // 3. Sort by category grouping
                regenerateCharts,        // 4. Regenerate charts flag
                includeObjectiveNotMet   // 5. Include Objective Not Met flag
              );
          }
          
          function onSuccess(result) {
            console.log('Success:', result);
            google.script.host.close();
          }
          
          function onFailure(error) {
            console.error('Error:', error);
            document.getElementById('formContent').classList.remove('hidden');
            document.getElementById('loading').style.display = 'none';
            showError('Error: ' + (error.message || error));
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

/**
 * ════════════════════════════════════════════════════════════════════════════
 * PREDICTABILITY PRESENTATION DATA POPULATION
 * ════════════════════════════════════════════════════════════════════════════
 * 
 * Populates presentation slides with data from the Predictability Score sheet.
 * 
 * RULES FOR ART VELOCITY CHART (Table 3):
 * - PI 11 data is in Column L (index 12)
 * - PI numbers are in Row 40
 * - Data rows start at Row 41
 * - ONE column per PI (not 5 like Allocation Distribution)
 * 
 * Data Sources:
 * - Table 1: ART Capacity (Rows 1-17) - MANUAL, SKIP
 * - Table 2: Allocation Distribution (Rows 20-37) - 5 columns per PI
 * - Table 3: ART Velocity (Rows 39-55) - 1 column per PI ← KEY TABLE
 * - Table 4: Epic Status (Rows 58-75) - 3 columns per PI
 * - Table 5: Program Predictability (Rows 77-94) - 3 columns per PI
 * - Table 6: Objective Status (Rows 96-113) - 5 columns per PI
 * - Table 7: PI Score Summary (Rows 115+) - 1 column per PI
 * 
 * NOTE: Uses existing VALUE_STREAMS constant from core.gs
 */

// ═══════════════════════════════════════════════════════════════════════════
// CONSTANTS - PREDICTABILITY SCORE SHEET STRUCTURE
// ═══════════════════════════════════════════════════════════════════════════

const PRED_SHEET_CONFIG = {
  // Table 2: Allocation Distribution
  TABLE2: {
    HEADER_ROW: 21,
    SUBHEADER_ROW: 23,
    DATA_START_ROW: 24,
    DATA_END_ROW: 37,
    COLS_PER_PI: 5, // Product-Feature, Tech/Platform, Compliance, Infosec, Quality
    PI_11_START_COL: 2 // Column B
  },
  
  // Table 3: ART Velocity
  TABLE3: {
    HEADER_ROW: 40,
    DATA_START_ROW: 41,
    DATA_END_ROW: 55,
    COLS_PER_PI: 1, // ONE column per PI
    PI_11_COL: 12   // Column L
  },
  
  // Table 4: Epic Status
  TABLE4: {
    HEADER_ROW: 59,
    SUBHEADER_ROW: 60,
    DATA_START_ROW: 61,
    DATA_END_ROW: 75,
    COLS_PER_PI: 3, // Committed, Uncommitted, Total
    PI_11_START_COL: 2 // Column B
  },
  
  // Table 5: Program Predictability
  TABLE5: {
    HEADER_ROW: 78,
    SUBHEADER_ROW: 79,
    DATA_START_ROW: 80,
    DATA_END_ROW: 94,
    COLS_PER_PI: 3, // Business Value, Actual Value, PI Score
    PI_11_START_COL: 2 // Column B
  },
  
  // Table 6: Objective Status
  TABLE6: {
    HEADER_ROW: 97,
    SUBHEADER_ROW: 98,
    DATA_START_ROW: 99,
    DATA_END_ROW: 113,
    COLS_PER_PI: 5, // Met, Partial, Not Met, All Epics, Score
    PI_11_START_COL: 2 // Column B
  },
  
  // Table 7: PI Score Summary
  TABLE7: {
    HEADER_ROW: 116,
    DATA_START_ROW: 117,
    DATA_END_ROW: 130,
    COLS_PER_PI: 1, // ONE column per PI
    PI_11_COL: 12   // Column L
  }
};

// NOTE: VALUE_STREAMS constant is already defined in core.gs
// We use that existing constant instead of declaring it again

// ═══════════════════════════════════════════════════════════════════════════
// MAIN POPULATION FUNCTION - INSERT THIS INTO generatePredictabilityPresentation
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Populates data slides in the Predictability Report presentation
 * Call this AFTER slide 2 and BEFORE generating deferral slides
 * 
 * @param {Presentation} presentation - The presentation to populate
 * @param {number} piNumber - The PI number
 * @param {Sheet} predScoreSheet - The Predictability Score sheet
 */
function populatePredictabilityDataSlides(presentation, piNumber, predScoreSheet) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`POPULATING PREDICTABILITY DATA SLIDES - PI ${piNumber}`);
  console.log(`${'='.repeat(80)}\n`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get or create Predictability Score sheet
    let scoreSheet = predScoreSheet;
    if (!scoreSheet) {
      scoreSheet = ss.getSheetByName('Predictability Score');
      if (!scoreSheet) {
        throw new Error('Predictability Score sheet not found. Please run "Update Predictability Score" first.');
      }
    }
    
    console.log(`✓ Using Predictability Score sheet`);
    console.log(`✓ Using VALUE_STREAMS from core.gs (${VALUE_STREAMS.length} value streams)`);
    
    // ═══════════════════════════════════════════════════════════════════
    // STEP 1: POPULATE SLIDE 3 - ART VELOCITY CHART
    // ═══════════════════════════════════════════════════════════════════
    console.log(`\n📊 Populating Slide 3: ART Velocity Chart...`);
    populateARTVelocitySlide(presentation, piNumber, scoreSheet);
    
    // ═══════════════════════════════════════════════════════════════════
    // STEP 2: POPULATE SLIDE 4 - ALLOCATION DISTRIBUTION
    // ═══════════════════════════════════════════════════════════════════
    console.log(`\n📊 Populating Slide 4: Allocation Distribution...`);
    populateAllocationDistributionSlide(presentation, piNumber, scoreSheet);
    
    // ═══════════════════════════════════════════════════════════════════
    // STEP 3: POPULATE SLIDE 5 - EPIC STATUS
    // ═══════════════════════════════════════════════════════════════════
    console.log(`\n📊 Populating Slide 5: Epic Status...`);
    populateEpicStatusSlide(presentation, piNumber, scoreSheet);
    
    // ═══════════════════════════════════════════════════════════════════
    // STEP 4: POPULATE SLIDE 6 - PROGRAM PREDICTABILITY
    // ═══════════════════════════════════════════════════════════════════
    console.log(`\n📊 Populating Slide 6: Program Predictability...`);
    populateProgramPredictabilitySlide(presentation, piNumber, scoreSheet);
    
    // ═══════════════════════════════════════════════════════════════════
    // STEP 5: POPULATE SLIDE 7 - OBJECTIVE STATUS
    // ═══════════════════════════════════════════════════════════════════
    console.log(`\n📊 Populating Slide 7: Objective Status...`);
    populateObjectiveStatusSlide(presentation, piNumber, scoreSheet);
    
    // ═══════════════════════════════════════════════════════════════════
    // STEP 6: POPULATE SLIDE 8 - PI SCORE SUMMARY
    // ═══════════════════════════════════════════════════════════════════
    console.log(`\n📊 Populating Slide 8: PI Score Summary...`);
    populatePIScoreSummarySlide(presentation, piNumber, scoreSheet);
    
    console.log(`\n${'='.repeat(80)}`);
    console.log(`✅ ALL DATA SLIDES POPULATED`);
    console.log(`${'='.repeat(80)}\n`);
    
  } catch (error) {
    console.error(`\n❌ ERROR populating data slides: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 3: ART VELOCITY CHART
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Populates Slide 3 with ART Velocity data from Table 3
 * Table 3 Rules:
 * - PI 11 is in Column L (index 12)
 * - PI numbers in Row 40
 * - Data starts Row 41
 * - ONE column per PI
 */
function populateARTVelocitySlide(presentation, piNumber, scoreSheet) {
  console.log(`  📌 Fetching ART Velocity data for PI ${piNumber}...`);
  
  const slides = presentation.getSlides();
  if (slides.length < 3) {
    console.warn(`  ⚠️ Slide 3 not found, skipping ART Velocity population`);
    return;
  }
  
  const slide = slides[2]; // Slide 3 (0-indexed)
  
  // Calculate column for this PI
  const piOffset = piNumber - 11; // PI 11 = 0, PI 12 = 1, etc.
  const dataColumn = PRED_SHEET_CONFIG.TABLE3.PI_11_COL + piOffset;
  
  console.log(`  📍 Reading from Column ${columnNumberToLetter(dataColumn)} (PI ${piNumber})`);
  
  // Read data from Table 3
  const dataStartRow = PRED_SHEET_CONFIG.TABLE3.DATA_START_ROW;
  const dataEndRow = PRED_SHEET_CONFIG.TABLE3.DATA_END_ROW;
  const numRows = dataEndRow - dataStartRow + 1;
  
  const velocityData = [];
  
  for (let i = 0; i < VALUE_STREAMS.length && i < numRows; i++) {
    const row = dataStartRow + i;
    const valueStream = VALUE_STREAMS[i];
    const velocity = scoreSheet.getRange(row, dataColumn).getValue();
    
    velocityData.push({
      valueStream: valueStream,
      velocity: velocity || 0
    });
    
    console.log(`  ${valueStream}: ${velocity}`);
  }
  
  // Replace placeholder table {{ART_VELOCITY_TABLE}}
  replaceTablePlaceholderWithData(
    slide,
    '{{ART_VELOCITY_TABLE}}',
    velocityData,
    ['Value Stream', 'Velocity (Story Points)'],
    (row) => [row.valueStream, Math.round(row.velocity)]
  );
  
  console.log(`  ✓ ART Velocity slide populated`);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 4: ALLOCATION DISTRIBUTION
// ═══════════════════════════════════════════════════════════════════════════

function populateAllocationDistributionSlide(presentation, piNumber, scoreSheet) {
  console.log(`  📌 Fetching Allocation Distribution data for PI ${piNumber}...`);
  
  const slides = presentation.getSlides();
  if (slides.length < 4) {
    console.warn(`  ⚠️ Slide 4 not found, skipping Allocation Distribution`);
    return;
  }
  
  const slide = slides[3]; // Slide 4
  
  // Calculate starting column for this PI
  const piOffset = piNumber - 11;
  const startCol = PRED_SHEET_CONFIG.TABLE2.PI_11_START_COL + (piOffset * PRED_SHEET_CONFIG.TABLE2.COLS_PER_PI);
  
  console.log(`  📍 Reading from Columns ${columnNumberToLetter(startCol)}-${columnNumberToLetter(startCol + 4)} (PI ${piNumber})`);
  
  // Read data
  const dataStartRow = PRED_SHEET_CONFIG.TABLE2.DATA_START_ROW;
  const dataEndRow = PRED_SHEET_CONFIG.TABLE2.DATA_END_ROW;
  const numRows = dataEndRow - dataStartRow + 1;
  
  const allocationData = [];
  
  for (let i = 0; i < VALUE_STREAMS.length && i < numRows; i++) {
    const row = dataStartRow + i;
    const valueStream = VALUE_STREAMS[i];
    
    const productFeature = scoreSheet.getRange(row, startCol).getValue() || 0;
    const techPlatform = scoreSheet.getRange(row, startCol + 1).getValue() || 0;
    const compliance = scoreSheet.getRange(row, startCol + 2).getValue() || 0;
    const KLO = scoreSheet.getRange(row, startCol + 3).getValue() || 0;
    const quality = scoreSheet.getRange(row, startCol + 4).getValue() || 0;
    
    allocationData.push({
      valueStream: valueStream,
      productFeature: Math.round(productFeature),
      techPlatform: Math.round(techPlatform),
      compliance: Math.round(compliance),
      KLO: Math.round(KLO),
      quality: Math.round(quality)
    });
  }
  
  // Replace placeholder
  replaceTablePlaceholderWithData(
    slide,
    '{{ALLOCATION_TABLE}}',
    allocationData,
    ['Value Stream', 'Product-Feature', 'Tech/Platform', 'Compliance', 'KLO', 'Quality'],
    (row) => [
      row.valueStream,
      row.productFeature,
      row.techPlatform,
      row.compliance,
      row.KLO,
      row.quality
    ]
  );
  
  console.log(`  ✓ Allocation Distribution slide populated`);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 5: EPIC STATUS
// ═══════════════════════════════════════════════════════════════════════════

function populateEpicStatusSlide(presentation, piNumber, scoreSheet) {
  console.log(`  📌 Fetching Epic Status data for PI ${piNumber}...`);
  
  const slides = presentation.getSlides();
  if (slides.length < 5) {
    console.warn(`  ⚠️ Slide 5 not found, skipping Epic Status`);
    return;
  }
  
  const slide = slides[4]; // Slide 5
  
  // Calculate starting column
  const piOffset = piNumber - 11;
  const startCol = PRED_SHEET_CONFIG.TABLE4.PI_11_START_COL + (piOffset * PRED_SHEET_CONFIG.TABLE4.COLS_PER_PI);
  
  // Read data
  const dataStartRow = PRED_SHEET_CONFIG.TABLE4.DATA_START_ROW;
  const dataEndRow = PRED_SHEET_CONFIG.TABLE4.DATA_END_ROW;
  const numRows = dataEndRow - dataStartRow + 1;
  
  const epicStatusData = [];
  
  for (let i = 0; i < VALUE_STREAMS.length && i < numRows; i++) {
    const row = dataStartRow + i;
    const valueStream = VALUE_STREAMS[i];
    
    const committed = scoreSheet.getRange(row, startCol).getValue() || 0;
    const uncommitted = scoreSheet.getRange(row, startCol + 1).getValue() || 0;
    const total = scoreSheet.getRange(row, startCol + 2).getValue() || 0;
    
    epicStatusData.push({
      valueStream: valueStream,
      committed: Math.round(committed),
      uncommitted: Math.round(uncommitted),
      total: Math.round(total)
    });
  }
  
  replaceTablePlaceholderWithData(
    slide,
    '{{EPIC_STATUS_TABLE}}',
    epicStatusData,
    ['Value Stream', 'Committed', 'Uncommitted', 'Total'],
    (row) => [row.valueStream, row.committed, row.uncommitted, row.total]
  );
  
  console.log(`  ✓ Epic Status slide populated`);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 6: PROGRAM PREDICTABILITY
// ═══════════════════════════════════════════════════════════════════════════

function populateProgramPredictabilitySlide(presentation, piNumber, scoreSheet) {
  console.log(`  📌 Fetching Program Predictability data for PI ${piNumber}...`);
  
  const slides = presentation.getSlides();
  if (slides.length < 6) {
    console.warn(`  ⚠️ Slide 6 not found, skipping Program Predictability`);
    return;
  }
  
  const slide = slides[5]; // Slide 6
  
  // Calculate starting column
  const piOffset = piNumber - 11;
  const startCol = PRED_SHEET_CONFIG.TABLE5.PI_11_START_COL + (piOffset * PRED_SHEET_CONFIG.TABLE5.COLS_PER_PI);
  
  // Read data
  const dataStartRow = PRED_SHEET_CONFIG.TABLE5.DATA_START_ROW;
  const dataEndRow = PRED_SHEET_CONFIG.TABLE5.DATA_END_ROW;
  const numRows = dataEndRow - dataStartRow + 1;
  
  const predData = [];
  
  for (let i = 0; i < VALUE_STREAMS.length && i < numRows; i++) {
    const row = dataStartRow + i;
    const valueStream = VALUE_STREAMS[i];
    
    const businessValue = scoreSheet.getRange(row, startCol).getValue() || 0;
    const actualValue = scoreSheet.getRange(row, startCol + 1).getValue() || 0;
    const piScore = scoreSheet.getRange(row, startCol + 2).getValue() || 0;
    
    predData.push({
      valueStream: valueStream,
      businessValue: Math.round(businessValue),
      actualValue: Math.round(actualValue),
      piScore: typeof piScore === 'number' ? piScore.toFixed(1) + '%' : piScore
    });
  }
  
  replaceTablePlaceholderWithData(
    slide,
    '{{PROGRAM_PREDICTABILITY_TABLE}}',
    predData,
    ['Value Stream', 'Business Value', 'Actual Value', 'PI Score'],
    (row) => [row.valueStream, row.businessValue, row.actualValue, row.piScore]
  );
  
  console.log(`  ✓ Program Predictability slide populated`);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 7: OBJECTIVE STATUS
// ═══════════════════════════════════════════════════════════════════════════

function populateObjectiveStatusSlide(presentation, piNumber, scoreSheet) {
  console.log(`  📌 Fetching Objective Status data for PI ${piNumber}...`);
  
  const slides = presentation.getSlides();
  if (slides.length < 7) {
    console.warn(`  ⚠️ Slide 7 not found, skipping Objective Status`);
    return;
  }
  
  const slide = slides[6]; // Slide 7
  
  // Calculate starting column
  const piOffset = piNumber - 11;
  const startCol = PRED_SHEET_CONFIG.TABLE6.PI_11_START_COL + (piOffset * PRED_SHEET_CONFIG.TABLE6.COLS_PER_PI);
  
  // Read data
  const dataStartRow = PRED_SHEET_CONFIG.TABLE6.DATA_START_ROW;
  const dataEndRow = PRED_SHEET_CONFIG.TABLE6.DATA_END_ROW;
  const numRows = dataEndRow - dataStartRow + 1;
  
  const objectiveData = [];
  
  for (let i = 0; i < VALUE_STREAMS.length && i < numRows; i++) {
    const row = dataStartRow + i;
    const valueStream = VALUE_STREAMS[i];
    
    const met = scoreSheet.getRange(row, startCol).getValue() || 0;
    const partial = scoreSheet.getRange(row, startCol + 1).getValue() || 0;
    const notMet = scoreSheet.getRange(row, startCol + 2).getValue() || 0;
    const allEpics = scoreSheet.getRange(row, startCol + 3).getValue() || 0;
    const score = scoreSheet.getRange(row, startCol + 4).getValue() || 0;
    
    objectiveData.push({
      valueStream: valueStream,
      met: Math.round(met),
      partial: Math.round(partial),
      notMet: Math.round(notMet),
      allEpics: Math.round(allEpics),
      score: typeof score === 'number' ? score.toFixed(1) + '%' : score
    });
  }
  
  replaceTablePlaceholderWithData(
    slide,
    '{{OBJECTIVE_STATUS_TABLE}}',
    objectiveData,
    ['Value Stream', 'Met', 'Partial', 'Not Met', 'All Epics', 'Score'],
    (row) => [row.valueStream, row.met, row.partial, row.notMet, row.allEpics, row.score]
  );
  
  console.log(`  ✓ Objective Status slide populated`);
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE 8: PI SCORE SUMMARY
// ═══════════════════════════════════════════════════════════════════════════

function populatePIScoreSummarySlide(presentation, piNumber, scoreSheet) {
  console.log(`  📌 Fetching PI Score Summary data for PI ${piNumber}...`);
  
  const slides = presentation.getSlides();
  if (slides.length < 8) {
    console.warn(`  ⚠️ Slide 8 not found, skipping PI Score Summary`);
    return;
  }
  
  const slide = slides[7]; // Slide 8
  
  // Calculate column (same as ART Velocity - ONE column per PI)
  const piOffset = piNumber - 11;
  const dataColumn = PRED_SHEET_CONFIG.TABLE7.PI_11_COL + piOffset;
  
  // Read data
  const dataStartRow = PRED_SHEET_CONFIG.TABLE7.DATA_START_ROW;
  const dataEndRow = PRED_SHEET_CONFIG.TABLE7.DATA_END_ROW;
  const numRows = dataEndRow - dataStartRow + 1;
  
  const scoreData = [];
  
  for (let i = 0; i < VALUE_STREAMS.length && i < numRows; i++) {
    const row = dataStartRow + i;
    const valueStream = VALUE_STREAMS[i];
    const score = scoreSheet.getRange(row, dataColumn).getValue() || 0;
    
    scoreData.push({
      valueStream: valueStream,
      score: typeof score === 'number' ? score.toFixed(1) + '%' : score
    });
  }
  
  replaceTablePlaceholderWithData(
    slide,
    '{{PI_SCORE_SUMMARY_TABLE}}',
    scoreData,
    ['Value Stream', 'PI Score'],
    (row) => [row.valueStream, row.score]
  );
  
  console.log(`  ✓ PI Score Summary slide populated`);
}

// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Replaces a table placeholder with actual data
 */
function replaceTablePlaceholderWithData(slide, placeholder, data, headers, rowMapper) {
  // Find and remove placeholder
  const shapes = slide.getShapes();
  let tablePosition = null;
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    try {
      const text = shape.getText().asString();
      if (text.includes(placeholder)) {
        tablePosition = {
          left: shape.getLeft(),
          top: shape.getTop(),
          width: shape.getWidth(),
          height: shape.getHeight()
        };
        shape.remove();
        break;
      }
    } catch (e) {
      // Skip shapes without text
    }
  }
  
  if (!tablePosition) {
    console.warn(`  ⚠️ Placeholder ${placeholder} not found`);
    return;
  }
  
  // Create table
  const numRows = data.length + 1; // +1 for header
  const numCols = headers.length;
  
  const table = slide.insertTable(
    numRows,
    numCols,
    tablePosition.left,
    tablePosition.top,
    tablePosition.width,
    tablePosition.height
  );
  
  // Set headers
  for (let col = 0; col < headers.length; col++) {
    const cell = table.getCell(0, col);
    cell.getText().setText(headers[col]);
    cell.getText().getTextStyle()
      .setBold(true)
      .setFontSize(10)
      .setForegroundColor('#FFFFFF');
    cell.getFill().setSolidFill('#4A86E8');
  }
  
  // Set data rows
  for (let row = 0; row < data.length; row++) {
    const rowData = rowMapper(data[row]);
    for (let col = 0; col < rowData.length; col++) {
      const cell = table.getCell(row + 1, col);
      cell.getText().setText(String(rowData[col]));
      cell.getText().getTextStyle().setFontSize(9);
      
      // Alternate row colors
      if ((row + 1) % 2 === 0) {
        cell.getFill().setSolidFill('#F3F3F3');
      }
    }
  }
  
  console.log(`  ✓ Table created: ${numRows} rows × ${numCols} cols`);
}

/**
 * Converts column number to letter (1 = A, 2 = B, etc.)
 * Note: There's likely already a columnNumberToLetter function in your project,
 * but including this for safety. If you get a duplicate declaration error,
 * just remove this function.
 */
function columnNumberToLetter(column) {
  let letter = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

function replaceTextWithFormatting(slide, placeholder, replacement, fontSize) {
  fontSize = fontSize || 9; // Default to 9 if not specified
  
  try {
    // Get all shapes on the slide
    const shapes = slide.getShapes();
    
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      
      try {
        const textRange = shape.getText();
        const currentText = textRange.asString();
        
        // Check if this shape contains the placeholder
        if (currentText.includes(placeholder)) {
          // Find the position of the placeholder in the text
          const startIndex = currentText.indexOf(placeholder);
          const endIndex = startIndex + placeholder.length;
          
          // Replace the placeholder with the new text
          const newText = currentText.split(placeholder).join(replacement);
          textRange.setText(newText);
          
          // Set font size for the replaced text
          // Calculate new end index based on replacement length
          const newEndIndex = startIndex + replacement.length;
          
          if (replacement.length > 0) {
            // Get the range of just the replaced text
            const replacedRange = textRange.getRange(startIndex, newEndIndex);
            
            // Apply font size
            replacedRange.getTextStyle().setFontSize(fontSize);
          }
          
          return true;
        }
      } catch (e) {
        // Shape doesn't have text, skip it
        continue;
      }
    }
    
    // Also check tables
    const tables = slide.getTables();
    for (let i = 0; i < tables.length; i++) {
      const table = tables[i];
      const numRows = table.getNumRows();
      const numCols = table.getNumColumns();
      
      for (let row = 0; row < numRows; row++) {
        for (let col = 0; col < numCols; col++) {
          try {
            const cell = table.getCell(row, col);
            const textRange = cell.getText();
            const currentText = textRange.asString();
            
            if (currentText.includes(placeholder)) {
              // Replace text
              const newText = currentText.split(placeholder).join(replacement);
              textRange.setText(newText);
              
              // Set font size for entire cell
              textRange.getTextStyle().setFontSize(fontSize);
              
              return true;
            }
          } catch (e) {
            continue;
          }
        }
      }
    }
    
  } catch (error) {
    console.error(`  ❌ Error replacing "${placeholder}" with formatting: ${error.message}`);
  }
  
  return false;
}
function populateDeferralTableForSlide(slide, issuesBatch) {
  console.log(`      Populating ${issuesBatch.length} issues (dynamic rows)`);
  
  const tables = slide.getTables();
  if (!tables || tables.length === 0) {
    throw new Error('No table found');
  }
  
  const table = tables[0];
  const numCols = table.getNumColumns();
  const initialRows = table.getNumRows();
  
  console.log(`      Table: ${initialRows} rows × ${numCols} cols`);
  
  const HEADER_ROW = 0;
  const TEMPLATE_DATA_ROW = 1;
  const FONT_SIZE = 9;
  
  // Add rows if needed (template has 1 header + 1 data row)
  const rowsToAdd = issuesBatch.length - 1;
  if (rowsToAdd > 0) {
    console.log(`      Adding ${rowsToAdd} rows`);
    for (let i = 0; i < rowsToAdd; i++) {
      table.appendRow();
    }
    console.log(`      Table now: ${table.getNumRows()} rows`);
  }
  
  // Populate each row
  for (let i = 0; i < issuesBatch.length; i++) {
    const issue = issuesBatch[i];
    const dataRow = TEMPLATE_DATA_ROW + i;
    
    try {
      // Col 0: Program Initiative
      setCellTextHelper(table, dataRow, 0, issue.programInitiative || '', FONT_SIZE);
      
      // Col 1: Epic Summary with JIRA hyperlink
      const epicSummary = issue.epicSummary || '';
      const cell1 = table.getCell(dataRow, 1);
      const text1 = cell1.getText();
      text1.setText(epicSummary);
      
      if (issue.issueKey && epicSummary) {
        const jiraUrl = JIRA_BASE_URL_PRED + issue.issueKey;
        text1.getTextStyle()
          .setLinkUrl(jiraUrl)
          .setForegroundColor('#000000')
          .setBold(true)
          .setUnderline(false)
          .setFontSize(FONT_SIZE);
      } else {
        text1.getTextStyle().setFontSize(FONT_SIZE);
      }
      
      // Col 2: Value Stream
      setCellTextHelper(table, dataRow, 2, issue.valuestream || '', FONT_SIZE);
      
      // Col 3: RCA AI Summary
      setCellTextHelper(table, dataRow, 3, issue.aiRagNote || '', FONT_SIZE);
      
      // Col 4: Reason (color-coded)
      const reason = issue.reason || 'Deferred';
      const cell4 = table.getCell(dataRow, 4);
      const text4 = cell4.getText();
      text4.setText(reason);
      
      const reasonColor = reason === 'Objective Not Met' ? '#FF6600' : '#CC0000';
      text4.getTextStyle()
        .setFontSize(FONT_SIZE)
        .setForegroundColor(reasonColor)
        .setBold(true);
      
      // Col 5: Impact
      setCellTextHelper(table, dataRow, 5, 'RTE TO COMPLETE', FONT_SIZE, '#DA1230', true);
      
    } catch (error) {
      console.error(`      ❌ Row ${dataRow}: ${error.message}`);
    }
  }
  
  console.log(`      ✅ Populated ${issuesBatch.length} rows`);
}

function setCellTextHelper(table, row, col, value, fontSize, color, bold) {
  try {
    const cell = table.getCell(row, col);
    const textRange = cell.getText();
    textRange.setText(value || '');
    
    const style = textRange.getTextStyle();
    style.setFontSize(fontSize || 9);
    
    if (color) style.setForegroundColor(color);
    if (bold) style.setBold(true);
  } catch (e) {
    console.warn(`      ⚠️ Cell [${row}, ${col}]: ${e.message}`);
  }
}

function getPortfolioSortOrder() {
  const order = [
    'Information Security', 'EMR Platform Integrity', 'Insights', 'Compliance',
    'Provider Experience', 'Patient Engagement', 'Cloud Migration', 'PM Rewrite', 'Practice Efficiency'
  ];
  const map = {};
  order.forEach((name, i) => map[name] = i);
  return map;
}

function findDeferralTemplateSlide(presentation) {
  console.log('\n🔍 Finding deferral template...');
  
  const slides = presentation.getSlides();
  const DEFERRAL_SLIDE_INDEX = 2;
  
  if (slides.length <= DEFERRAL_SLIDE_INDEX) {
    throw new Error(`Template needs at least ${DEFERRAL_SLIDE_INDEX + 1} slides`);
  }
  
  const deferralSlide = slides[DEFERRAL_SLIDE_INDEX];
  
  const tables = deferralSlide.getTables();
  if (tables && tables.length > 0) {
    const table = tables[0];
    console.log(`   ✓ Found at slide ${DEFERRAL_SLIDE_INDEX + 1}: ${table.getNumRows()} rows × ${table.getNumColumns()} cols`);
  } else {
    console.warn(`   ⚠️ Slide ${DEFERRAL_SLIDE_INDEX + 1} has no table!`);
  }
  
  return { slide: deferralSlide, index: DEFERRAL_SLIDE_INDEX };
}

function generateValueStreamChartsSheetProgrammatically(piNumber, ss, predScoreSheet, iter6Sheet) {
  console.log(`  Generating chart sheet for PI ${piNumber}...`);
  
  const chartSheetName = `PI ${piNumber} - VS Charts`;
  let chartSheet = ss.getSheetByName(chartSheetName);
  
  // Clear existing or create new
  if (chartSheet) {
    chartSheet.getCharts().forEach(chart => chartSheet.removeChart(chart));
    chartSheet.clear();
    console.log(`  ✓ Cleared existing chart sheet`);
  } else {
    chartSheet = ss.insertSheet(chartSheetName);
    console.log(`  ✓ Created new chart sheet`);
  }
  
  const piOffset = piNumber - 11;
  let currentRow = 1;
  const SECTION_HEIGHT = 25;
  
  // Generate charts for each value stream
  PREDICTABILITY_VALUE_STREAMS.forEach((valueStream, vsIndex) => {
    console.log(`    Processing charts for: ${valueStream} (${vsIndex + 1}/${PREDICTABILITY_VALUE_STREAMS.length})`);
    
    // Section Header
    const headerCell = chartSheet.getRange(currentRow, 1, 1, 28);
    headerCell.merge();
    headerCell.setValue(valueStream);
    headerCell.setFontSize(14);
    headerCell.setFontWeight('bold');
    headerCell.setBackground('#674ea7');
    headerCell.setFontColor('#ffffff');
    headerCell.setHorizontalAlignment('center');
    headerCell.setVerticalAlignment('middle');
    chartSheet.setRowHeight(currentRow, 35);
    currentRow++;
    
    const sectionStartRow = currentRow;
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 1: Capacity & Velocity (formatted boxes with RichTextValue)
    // ═══════════════════════════════════════════════════════════════════════════
    const capacityRow = 4 + vsIndex;
    const capacityCol = 2 + ((piNumber - 11) * 2);  // PI 11→2, PI 12→4, PI 13→6
    const capacity = parseFloat(predScoreSheet.getRange(capacityRow, capacityCol).getValue()) || 0;
    const velocityRow = 41 + vsIndex;
    const velocityCol = 1 + piNumber;  // PI 1→2(B), PI 13→14(N)
    const velocity = parseFloat(predScoreSheet.getRange(velocityRow, velocityCol).getValue()) || 0;
    
    console.log(`      Chart 1: Capacity=${capacity}, Velocity=${velocity}`);
    
    // Capacity Box (Column A, rows 2-4 of section)
    const capacityBoxRange = chartSheet.getRange(sectionStartRow, 1, 3, 1);
    capacityBoxRange.merge();
    capacityBoxRange.setBackground('#674ea7');
    capacityBoxRange.setFontColor('#ffffff');
    capacityBoxRange.setFontSize(14);
    capacityBoxRange.setFontWeight('bold');
    capacityBoxRange.setHorizontalAlignment('center');
    capacityBoxRange.setVerticalAlignment('middle');
    
    const capacityRichText = SpreadsheetApp.newRichTextValue()
      .setText(`${Math.round(capacity)}\nART Capacity`)
      .setTextStyle(0, String(Math.round(capacity)).length, SpreadsheetApp.newTextStyle().setFontSize(18).setBold(true).build())
      .setTextStyle(String(Math.round(capacity)).length + 1, `${Math.round(capacity)}\nART Capacity`.length, 
        SpreadsheetApp.newTextStyle().setFontSize(10).setBold(false).build())
      .build();
    capacityBoxRange.setRichTextValue(capacityRichText);
    
    // Velocity Box (Column B, rows 2-4 of section)
    const velocityBoxRange = chartSheet.getRange(sectionStartRow, 2, 3, 1);
    velocityBoxRange.merge();
    velocityBoxRange.setBackground('#9b7bb8');
    velocityBoxRange.setFontColor('#ffffff');
    velocityBoxRange.setFontSize(14);
    velocityBoxRange.setFontWeight('bold');
    velocityBoxRange.setHorizontalAlignment('center');
    velocityBoxRange.setVerticalAlignment('middle');
    
    const velocityRichText = SpreadsheetApp.newRichTextValue()
      .setText(`${Math.round(velocity)}\nART Velocity`)
      .setTextStyle(0, String(Math.round(velocity)).length, SpreadsheetApp.newTextStyle().setFontSize(18).setBold(true).build())
      .setTextStyle(String(Math.round(velocity)).length + 1, `${Math.round(velocity)}\nART Velocity`.length, 
        SpreadsheetApp.newTextStyle().setFontSize(10).setBold(false).build())
      .build();
    velocityBoxRange.setRichTextValue(velocityRichText);
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 2: Allocation Pie Chart
    // ═══════════════════════════════════════════════════════════════════════════
    const allocRow = 23 + vsIndex;
    const allocStartCol = 2 + (piOffset * 5);
    const productFeature = predScoreSheet.getRange(allocRow, allocStartCol + 0).getValue() || 0;
    const compliance = predScoreSheet.getRange(allocRow, allocStartCol + 1).getValue() || 0;
    const klo = predScoreSheet.getRange(allocRow, allocStartCol + 2).getValue() || 0;
    const quality = predScoreSheet.getRange(allocRow, allocStartCol + 3).getValue() || 0;
    const techPlatform = predScoreSheet.getRange(allocRow, allocStartCol + 4).getValue() || 0;
    
    const pieDataRow = sectionStartRow;
    chartSheet.getRange(pieDataRow, 4).setValue('Category');
    chartSheet.getRange(pieDataRow, 5).setValue('Points');
    chartSheet.getRange(pieDataRow + 1, 4).setValue('Product-Feature');
    chartSheet.getRange(pieDataRow + 1, 5).setValue(productFeature);
    chartSheet.getRange(pieDataRow + 2, 4).setValue('Compliance');
    chartSheet.getRange(pieDataRow + 2, 5).setValue(compliance);
    chartSheet.getRange(pieDataRow + 3, 4).setValue('Tech/Platform');
    chartSheet.getRange(pieDataRow + 3, 5).setValue(techPlatform);
    chartSheet.getRange(pieDataRow + 4, 4).setValue('Quality');
    chartSheet.getRange(pieDataRow + 4, 5).setValue(quality);
    chartSheet.getRange(pieDataRow + 5, 4).setValue('KLO');
    chartSheet.getRange(pieDataRow + 5, 5).setValue(klo);
    
    const pieChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(chartSheet.getRange(pieDataRow + 1, 4, 5, 2))
      .setPosition(pieDataRow, 4, 0, 0)
      .setOption('title', 'Capacity Allocation Delivered')
      .setOption('titleTextStyle', { fontSize: 11, bold: true, alignment: 'center' })
      .setOption('width', 320)
      .setOption('height', 240)
      .setOption('is3D', true)
      .setOption('pieSliceText', 'percentage')
      .setOption('pieSliceTextStyle', { color: '#ffffff' })
      .setOption('legend', { position: 'bottom', alignment: 'center', textStyle: { fontSize: 9 } })
      .setOption('fontName', 'Tahoma')
      .setOption('colors', ['#674ea7', '#9900ff', '#f1c232', '#46bdc6', '#cccccc'])
      .build();
    chartSheet.insertChart(pieChart);
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 3: Completed & Deferred Boxes (from Iteration 6 sheet)
    // ═══════════════════════════════════════════════════════════════════════════
    const epicCounts = getEpicCountsForValueStream(iter6Sheet, valueStream);
    const completed = epicCounts.completed || 0;
    const deferred = epicCounts.deferred || 0;
    
    console.log(`      Chart 3: Completed=${completed}, Deferred=${deferred}`);
    
    // Completed Box (Column A, rows 8-10 of section)
    const completedBoxRange = chartSheet.getRange(sectionStartRow + 6, 1, 3, 1);
    completedBoxRange.merge();
    completedBoxRange.setBackground('#00AA00');
    completedBoxRange.setFontColor('#ffffff');
    completedBoxRange.setFontSize(14);
    completedBoxRange.setFontWeight('bold');
    completedBoxRange.setHorizontalAlignment('center');
    completedBoxRange.setVerticalAlignment('middle');
    
    const completedRichText = SpreadsheetApp.newRichTextValue()
      .setText(`${completed}\nFeatures Completed`)
      .setTextStyle(0, String(completed).length, SpreadsheetApp.newTextStyle().setFontSize(18).setBold(true).build())
      .setTextStyle(String(completed).length + 1, `${completed}\nFeatures Completed`.length, 
        SpreadsheetApp.newTextStyle().setFontSize(10).setBold(false).build())
      .build();
    completedBoxRange.setRichTextValue(completedRichText);
    
    // Deferred Box (Column B, rows 8-10 of section)
    const deferredBoxRange = chartSheet.getRange(sectionStartRow + 6, 2, 3, 1);
    deferredBoxRange.merge();
    deferredBoxRange.setBackground('#FFC82E');
    deferredBoxRange.setFontColor('#ffffff');
    deferredBoxRange.setFontSize(14);
    deferredBoxRange.setFontWeight('bold');
    deferredBoxRange.setHorizontalAlignment('center');
    deferredBoxRange.setVerticalAlignment('middle');
    
    const deferredRichText = SpreadsheetApp.newRichTextValue()
      .setText(`${deferred}\nFeatures Deferred`)
      .setTextStyle(0, String(deferred).length, SpreadsheetApp.newTextStyle().setFontSize(18).setBold(true).build())
      .setTextStyle(String(deferred).length + 1, `${deferred}\nFeatures Deferred`.length, 
        SpreadsheetApp.newTextStyle().setFontSize(10).setBold(false).build())
      .build();
    deferredBoxRange.setRichTextValue(deferredRichText);
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 4: Velocity Trending
    // ═══════════════════════════════════════════════════════════════════════════
    const velocityTrendRow = 41 + vsIndex;
    const velocityHistData = [];
    const startCol = 2;
    const currentPICol = 12 + (piNumber - 11);
    
    for (let col = startCol; col <= currentPICol; col++) {
      const value = predScoreSheet.getRange(velocityTrendRow, col).getValue() || 0;
      const pi = 11 + (col - 12);
      if (value > 0) {
        velocityHistData.push([`PI ${pi}`, value]);
      }
    }
    
    if (velocityHistData.length > 0) {
      const velChartDataRow = sectionStartRow;
      velocityHistData.forEach(([pi, value], idx) => {
        chartSheet.getRange(velChartDataRow + idx, 15).setValue(pi);
        chartSheet.getRange(velChartDataRow + idx, 16).setValue(value);
      });
      
      const velocityValues = velocityHistData.map(([pi, value]) => value);
      const minVelocity = Math.min(...velocityValues);
      const maxVelocity = Math.max(...velocityValues);
      const yAxisMin = Math.max(0, minVelocity - 200);
      const yAxisMax = maxVelocity + 200;
      
      const velChart = chartSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartSheet.getRange(velChartDataRow, 15, velocityHistData.length, 2))
        .setPosition(velChartDataRow, 15, 0, 0)
        .setOption('title', `${valueStream} ART Velocity`)
        .setOption('titleTextStyle', { fontSize: 10, bold: true })
        .setOption('width', 350)
        .setOption('height', 240)
        .setOption('curveType', 'function')
        .setOption('colors', ['#674ea7'])
        .setOption('pointSize', 5)
        .setOption('lineWidth', 3)
        .setOption('vAxis', { viewWindow: { min: yAxisMin, max: yAxisMax } })
        .build();
      chartSheet.insertChart(velChart);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 5: Program Predictability
    // ═══════════════════════════════════════════════════════════════════════════
    const predTrendRow = 117 + vsIndex;
    const predHistData = [];
    const predStartCol = 2;
    const predCurrentPICol = 12 + (piNumber - 11);
    
    for (let col = predStartCol; col <= predCurrentPICol; col++) {
      const rawValue = predScoreSheet.getRange(predTrendRow, col).getValue() || 0;
      const pi = 11 + (col - 12);
      if (rawValue > 0) {
        let value = (pi < 11) ? rawValue * 100 : rawValue;
        predHistData.push([`PI ${pi}`, value]);
      }
    }
    
    if (predHistData.length > 0) {
      const predChartDataRow = sectionStartRow;
      predHistData.forEach(([pi, value], idx) => {
        chartSheet.getRange(predChartDataRow + idx, 21).setValue(pi);
        chartSheet.getRange(predChartDataRow + idx, 22).setValue(value);
      });
      
      const predChart = chartSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartSheet.getRange(predChartDataRow, 21, predHistData.length, 2))
        .setPosition(predChartDataRow, 21, 0, 0)
        .setOption('title', `${valueStream} Program Predictability`)
        .setOption('titleTextStyle', { fontSize: 10, bold: true })
        .setOption('width', 350)
        .setOption('height', 240)
        .setOption('curveType', 'function')
        .setOption('colors', ['#674ea7'])
        .setOption('pointSize', 5)
        .setOption('lineWidth', 3)
        .setOption('vAxis', { viewWindow: { min: 65, max: 115 } })
        .build();
      chartSheet.insertChart(predChart);
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 6: BV/AV Summary Table
    // ═══════════════════════════════════════════════════════════════════════════
    const bvAvData = getBVAVSummaryData(iter6Sheet, piNumber, valueStream, predScoreSheet, vsIndex, piOffset);
    
    if (bvAvData && (bvAvData.totalBV > 0 || bvAvData.totalAV > 0)) {
      console.log(`      Chart 6: BV/AV - ${bvAvData.allocations.length} allocations, Total BV=${bvAvData.totalBV}, AV=${bvAvData.totalAV}`);
      
      const bvAvStartRow = sectionStartRow;
      const bvAvStartCol = 24; // Column X
      let tableRow = bvAvStartRow;
      
      // Header row
      chartSheet.getRange(tableRow, bvAvStartCol).setValue('Program Initiative');
      chartSheet.getRange(tableRow, bvAvStartCol + 1).setValue('BV');
      chartSheet.getRange(tableRow, bvAvStartCol + 2).setValue('AV');
      chartSheet.getRange(tableRow, bvAvStartCol, 1, 3)
        .setBackground('#674ea7')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setFontSize(9);
      tableRow++;
      
      // Data rows by allocation
      bvAvData.allocations.forEach(allocGroup => {
        // Allocation header row (merged)
        const allocHeaderRange = chartSheet.getRange(tableRow, bvAvStartCol, 1, 3);
        allocHeaderRange.merge();
        allocHeaderRange.setValue(allocGroup.allocation);
        allocHeaderRange.setBackground('#fff2cc');
        allocHeaderRange.setFontWeight('bold');
        allocHeaderRange.setFontSize(9);
        tableRow++;
        
        // Initiative rows
        allocGroup.initiatives.forEach(init => {
          chartSheet.getRange(tableRow, bvAvStartCol).setValue(init.name);
          chartSheet.getRange(tableRow, bvAvStartCol + 1).setValue(init.bv);
          chartSheet.getRange(tableRow, bvAvStartCol + 2).setValue(init.av);
          
          if (init.av < init.bv) {
            chartSheet.getRange(tableRow, bvAvStartCol + 2).setBackground('#ffcccc');
          }
          
          chartSheet.getRange(tableRow, bvAvStartCol, 1, 3).setFontSize(8);
          tableRow++;
        });
      });
      
      // Total row
      chartSheet.getRange(tableRow, bvAvStartCol).setValue('TOTAL');
      chartSheet.getRange(tableRow, bvAvStartCol + 1).setValue(bvAvData.totalBV);
      chartSheet.getRange(tableRow, bvAvStartCol + 2).setValue(bvAvData.totalAV);
      chartSheet.getRange(tableRow, bvAvStartCol, 1, 3)
        .setBackground('#add8e6')
        .setFontWeight('bold')
        .setFontSize(9);
      tableRow++;
      
      // PI Score row - merge first two columns, score in third
      chartSheet.getRange(tableRow, bvAvStartCol).setValue('Program Predictability Score');
      chartSheet.getRange(tableRow, bvAvStartCol, 1, 2).merge();
      chartSheet.getRange(tableRow, bvAvStartCol).setBackground('#fff2cc').setFontWeight('bold').setFontSize(9);
      
      chartSheet.getRange(tableRow, bvAvStartCol + 2).setValue(`${bvAvData.piScore.toFixed(1)}%`);
      const scoreColor = bvAvData.piScore >= 90 ? '#d4edda' : '#ffcccc';
      chartSheet.getRange(tableRow, bvAvStartCol + 2).setBackground(scoreColor).setFontWeight('bold').setFontSize(9);
      
    } else {
      console.log(`      Chart 6: No BV/AV data for ${valueStream}`);
    }
    
    currentRow += SECTION_HEIGHT;
  });
  
  // Set column widths
  chartSheet.setColumnWidths(1, 2, 100);   // Columns A-B
  chartSheet.setColumnWidth(24, 200);       // Column X - Program Initiative (wider)
  chartSheet.setColumnWidth(25, 40);        // Column Y - BV (narrower)
  chartSheet.setColumnWidth(26, 40);        // Column Z - AV (narrower)
  
  console.log(`  ✓ Generated charts for ${PREDICTABILITY_VALUE_STREAMS.length} value streams`);
  console.log(`  ✓ Charts 1, 3, 6 now populated on sheet`);
}

function findChartByPosition(chartSheet, targetRow, columnRange) {
  const charts = chartSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    const chart = charts[i];
    const containerInfo = chart.getContainerInfo();
    const chartRow = containerInfo.getAnchorRow();
    const chartCol = containerInfo.getAnchorColumn();
    if (chartRow === targetRow && chartCol >= columnRange[0] && chartCol <= columnRange[1]) {
      return chart;
    }
  }
  return null;
}
function updateDeferralSlideTitle(slide, category) {
  const shapes = slide.getShapes();
  for (const shape of shapes) {
    try {
      const textRange = shape.getText();
      const text = textRange.asString();
      if (text.includes('{{Category}}')) {
        textRange.setText(text.replace('{{Category}}', category));
        console.log(`      ✓ Title: "${category}"`);
        return true;
      }
    } catch (e) {}
  }
  console.warn(`      ⚠️ No {{Category}} placeholder`);
  return false;
}
function getDeferralCategory(issue) {
  const portfolioInit = (issue.portfolioInitiative || '').toLowerCase();
  const programInit = (issue.programInitiative || '').toLowerCase();
  const allocation = (issue.allocation || '').toLowerCase().trim();
  
  if (portfolioInit.includes('infosec') || programInit.includes('infosec')) return 'Infosec';
  if (allocation === 'product - compliance') return 'Product - Compliance';
  if (allocation === 'product - feature') return 'Product - Feature';
  if (allocation.includes('platform')) return 'Platform';
  return 'Unknown';
}

function sortIssuesWithinCategory(issues) {
  let portfolioOrder = {};
  try {
    portfolioOrder = getPortfolioSortOrder();
  } catch (e) {}
  
  issues.sort((a, b) => {
    const portA = a.portfolioInitiative || '';
    const portB = b.portfolioInitiative || '';
    const orderA = portfolioOrder[portA] !== undefined ? portfolioOrder[portA] : 999;
    const orderB = portfolioOrder[portB] !== undefined ? portfolioOrder[portB] : 999;
    
    if (orderA !== orderB) return orderA - orderB;
    
    const progA = (a.programInitiative || '').toLowerCase();
    const progB = (b.programInitiative || '').toLowerCase();
    if (progA < progB) return -1;
    if (progA > progB) return 1;
    
    const vsA = (a.valuestream || '').toLowerCase();
    const vsB = (b.valuestream || '').toLowerCase();
    if (vsA < vsB) return -1;
    if (vsA > vsB) return 1;
    
    return 0;
  });
}