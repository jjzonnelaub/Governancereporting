function loadAllIssuesFromSheets(sheetNames, ss) {
  const allIssues = [];
  
  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.warn(`Sheet not found: ${sheetName}`);
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < DATA_START_ROW) {
      console.log(`Sheet ${sheetName} has no data rows`);
      return;
    }
    
    // Get headers from row 4
    const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    
    // Build column map
    const columnMap = {};
    headers.forEach((header, index) => {
      if (header) columnMap[header] = index;
    });
    
    // Get all data rows (starting from row 5)
    const dataRows = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, lastCol).getValues();
    
    // Convert each row to an issue object
    dataRows.forEach(row => {
      const issue = {};
      headers.forEach((header, colIndex) => {
        issue[header] = row[colIndex];
      });
      
      // Only include rows with a valid key
      if (issue['Key'] && String(issue['Key']).trim() !== '') {
        allIssues.push(issue);
      }
    });
  });
  
  console.log(`Loaded ${allIssues.length} total issues from ${sheetNames.length} sheets`);
  return allIssues;
}

function showInitiativeDialog() {
  const html = HtmlService.createHtmlOutputFromFile('initiativedialog')
    .setWidth(500)
    .setHeight(680)
    .setTitle('Generate Initiative Report');
  SpreadsheetApp.getUi().showModalDialog(html, 'Generate Initiative Report');
}

function addInitiativeSlide(presentation, insertIndex, initiatives, piNumber, sourceSheetUrl, timestamp) {
  console.log(`[ADD INIT] Starting addInitiativeSlide at index ${insertIndex} with ${initiatives.length} initiative(s)`);
  
  const WIDTH = 720;
  const HEIGHT = 540;
  const MARGIN = 20;
  const HEADER_HEIGHT = 60;
  const FOOTER_HEIGHT = 30;
  const CONTENT_HEIGHT = HEIGHT - HEADER_HEIGHT - FOOTER_HEIGHT;
  
  const INITIATIVES_PER_PAGE = 3;
  const INITIATIVE_HEIGHT = CONTENT_HEIGHT / INITIATIVES_PER_PAGE;
  
  // Insert blank slide at the specified position
  const slide = presentation.insertSlide(insertIndex);
  console.log(`[ADD INIT] Inserted blank slide at index ${insertIndex}`);
  
  // Add header with PI number
  const headerText = slide.insertTextBox(`PI ${piNumber} - Initiative Details`);
  headerText.setLeft(MARGIN);
  headerText.setTop(10);
  headerText.setWidth(WIDTH - 2 * MARGIN);
  headerText.setHeight(40);
  
  const headerTextRange = headerText.getText();
  headerTextRange.getTextStyle()
    .setFontSize(24)
    .setBold(true)
    .setForegroundColor('#1a73e8');
  headerTextRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Add each initiative to the slide
  initiatives.forEach((initiative, idx) => {
    const yPos = HEADER_HEIGHT + (idx * INITIATIVE_HEIGHT);
    console.log(`[ADD INIT]   Initiative ${idx + 1}/${initiatives.length}: ${initiative.key} at Y=${yPos}`);
    
    // Initiative title with hyperlink
    const titleBox = slide.insertTextBox(initiative.title);
    titleBox.setLeft(MARGIN);
    titleBox.setTop(yPos + 5);
    titleBox.setWidth(WIDTH - 2 * MARGIN);
    titleBox.setHeight(30);
    
    const titleText = titleBox.getText();
    titleText.getTextStyle()
      .setFontSize(16)
      .setBold(true)
      .setForegroundColor('#333333');
    
    // Add hyperlink to source sheet
    if (sourceSheetUrl) {
      titleText.getRange(0, initiative.title.length).getTextStyle().setLinkUrl(sourceSheetUrl);
      console.log(`[ADD INIT]     Added hyperlink to title`);
    }
    
    // Epic summary section
    const epicSummaryY = yPos + 40;
    let epicText = `Epics (${initiative.epics.length}): `;
    
    initiative.epics.forEach((epic, epicIdx) => {
      const epicKey = epic['Key'];
      const epicSummary = (epic['Summary'] || '').trim();
      const displayText = epicSummary ? `${epicKey} - ${epicSummary}` : epicKey;
      
      if (epicIdx > 0) epicText += ', ';
      epicText += displayText;
      
      // Truncate if too long (approximate)
      if (epicText.length > 200) {
        epicText = epicText.substring(0, 197) + '...';
        return;
      }
    });
    
    const epicBox = slide.insertTextBox(epicText);
    epicBox.setLeft(MARGIN + 10);
    epicBox.setTop(epicSummaryY);
    epicBox.setWidth(WIDTH - 2 * MARGIN - 20);
    epicBox.setHeight(60);
    
    const epicTextRange = epicBox.getText();
    epicTextRange.getTextStyle()
      .setFontSize(11)
      .setForegroundColor('#666666');
    
    // Feature Points
    const fpY = epicSummaryY + 65;
    const fpText = `Total Feature Points: ${initiative.totalFeaturePoints.toFixed(1)}`;
    const fpBox = slide.insertTextBox(fpText);
    fpBox.setLeft(MARGIN + 10);
    fpBox.setTop(fpY);
    fpBox.setWidth(200);
    fpBox.setHeight(20);
    
    const fpTextRange = fpBox.getText();
    fpTextRange.getTextStyle()
      .setFontSize(10)
      .setBold(true)
      .setForegroundColor('#1a73e8');
    
    // Add separator line if not the last initiative on the page
    if (idx < initiatives.length - 1) {
      const lineY = yPos + INITIATIVE_HEIGHT - 5;
      const line = slide.insertLine(
        SlidesApp.LineCategory.STRAIGHT,
        MARGIN,
        lineY,
        WIDTH - MARGIN,
        lineY
      );
      line.getLineFill().setSolidFill('#e0e0e0');
      line.setWeight(1);
    }
  });
  
  // Add footer with timestamp
  const footerText = slide.insertTextBox(`Generated: ${timestamp}`);
  footerText.setLeft(MARGIN);
  footerText.setTop(HEIGHT - FOOTER_HEIGHT + 5);
  footerText.setWidth(WIDTH - 2 * MARGIN);
  footerText.setHeight(20);
  
  const footerTextRange = footerText.getText();
  footerTextRange.getTextStyle()
    .setFontSize(8)
    .setForegroundColor('#999999');
  footerTextRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  console.log(`[ADD INIT] ✓ Completed\n`);
}

function structureInitiativeReportData(issues, targetAllocations, initiativeField, initiativeValue) {
  console.log(`\n========================================`);
  console.log(`[STRUCTURE] Starting initiative report structuring`);
  console.log(`[STRUCTURE] Initiative Field: ${initiativeField}`);
  console.log(`[STRUCTURE] Initiative Value: ${initiativeValue}`);
  console.log(`[STRUCTURE] Processing ${issues.length} total issues`);
  console.log(`========================================`);
  
  // Filter to only epics that match our criteria
  const epics = issues.filter(issue => {
    if (issue['Issue Type'] !== 'Epic') return false;
    if (issue[initiativeField] !== initiativeValue) return false;
    const allocation = issue['Allocation'] || '';
    return targetAllocations.length === 0 || targetAllocations.includes(allocation);
  });
  
  console.log(`[STRUCTURE] Found ${epics.length} epics matching criteria`);
  
  // Build issue map for parent lookup
  const issueMap = {};
  issues.forEach(issue => {
    issueMap[issue['Key']] = issue;
  });
  console.log(`[STRUCTURE] Built issue map with ${Object.keys(issueMap).length} issues`);
  
  // Group epics by parent key (or by epic key if no parent)
  const initiativeMap = {};
  
  epics.forEach((epic, idx) => {
    const parentKey = epic['Parent Key'];
    const epicKey = epic['Key'];
    
    console.log(`\n[STRUCTURE] Processing epic ${idx + 1}/${epics.length}: ${epicKey}`);
    console.log(`[STRUCTURE]   Parent Key: ${parentKey || 'NONE'}`);
    
    // Determine grouping key:
    // - If epic HAS parent key -> use parent key (group with other epics under same parent)
    // - If epic has NO parent key -> use epic key itself (standalone)
    const groupKey = (parentKey && parentKey.trim() !== '') ? parentKey : epicKey;
    
    console.log(`[STRUCTURE]   Group Key: ${groupKey}`);
    
    // Create initiative group if it doesn't exist
    if (!initiativeMap[groupKey]) {
      let initiativeTitle;
      let useParentLink = false;
      
      if (parentKey && parentKey.trim() !== '') {
        // Epic HAS parent - use Initiative Title from the EPIC'S row (same logic as slide 2)
        const epicInitTitle = (epic['Initiative Title'] || '').trim();
        
        console.log(`[STRUCTURE]   Epic has parent: ${parentKey}`);
        console.log(`[STRUCTURE]     Epic's Initiative Title: "${epicInitTitle}"`);
        
        // Use Initiative Title from epic's row, fallback to parent key if empty
        if (epicInitTitle !== '') {
          initiativeTitle = epicInitTitle;
          console.log(`[STRUCTURE]   Using epic's Initiative Title: "${initiativeTitle}"`);
        } else {
          initiativeTitle = parentKey;
          console.log(`[STRUCTURE]   Epic's Initiative Title empty, using parent key: "${initiativeTitle}"`);
        }
        
        useParentLink = true;
      } else {
        // Epic has NO parent - use epic's own information
        const epicInitTitle = (epic['Initiative Title'] || '').trim();
        const epicSummary = (epic['Summary'] || '').trim();
        
        console.log(`[STRUCTURE]   No parent - using epic's own data`);
        console.log(`[STRUCTURE]     Initiative Title: "${epicInitTitle}"`);
        console.log(`[STRUCTURE]     Summary: "${epicSummary}"`);
        
        // Use Initiative Title if available, otherwise Summary, otherwise key
        if (epicInitTitle !== '') {
          initiativeTitle = epicInitTitle;
          console.log(`[STRUCTURE]   Using epic Initiative Title: "${initiativeTitle}"`);
        } else if (epicSummary !== '') {
          initiativeTitle = epicSummary;
          console.log(`[STRUCTURE]   Using epic Summary: "${initiativeTitle}"`);
        } else {
          initiativeTitle = epicKey;
          console.log(`[STRUCTURE]   Epic fields empty, using epic key: "${initiativeTitle}"`);
        }
        
        useParentLink = false;
      }
      
      initiativeMap[groupKey] = {
        key: groupKey,
        title: initiativeTitle,
        hasParent: useParentLink,
        epics: [],
        totalFeaturePoints: 0
      };
      
      console.log(`[STRUCTURE]   ✓ Created new group: ${groupKey} with title: "${initiativeTitle}"`);
    }
    
    // Add epic to the group
    initiativeMap[groupKey].epics.push(epic);
    initiativeMap[groupKey].totalFeaturePoints += parseFloat(epic['Feature Points']) || 0;
    
    console.log(`[STRUCTURE]   ✓ Added epic to group ${groupKey}`);
    console.log(`[STRUCTURE]   Group now has ${initiativeMap[groupKey].epics.length} epic(s), ${initiativeMap[groupKey].totalFeaturePoints} FP`);
  });
  
  // Check if this is an INFOSEC initiative report
  const isInfosec = initiativeValue && initiativeValue.toUpperCase().includes('INFOSEC');
  console.log(`[STRUCTURE] Is INFOSEC initiative: ${isInfosec}`);
  
  let initiativeGroups;
  
  if (isInfosec) {
    // INFOSEC-specific grouping by value stream groups
    console.log(`[STRUCTURE] Applying INFOSEC-specific value stream grouping`);
    
    // Define value stream groups
    const vsGroups = {
      'EMA': [
        'EMA Clinical', 'EMA RaC', 'Xtract', 'Platform Engineering', 
        'Shared Services Platform', 'Fusion and Conversions', 'MMPM', 
        'RCM Genie', 'Cloud Foundation Services', 'Cloud Operations'
      ],
      'MMGI': ['MMGI-Cloud', 'MMGI'],
      'Other': ['Patient Collaboration', 'AIMM Data Platform Engineering']
    };
    
    // Create a reverse lookup map: value stream -> group name
    const vsToGroup = {};
    Object.keys(vsGroups).forEach(groupName => {
      vsGroups[groupName].forEach(vs => {
        vsToGroup[vs] = groupName;
      });
    });
    
    // Analyze each initiative to determine which VS groups it spans
    const initiativesByVSGroup = {
      'EMA': [],
      'MMGI': [],
      'Other': []
    };
    
    Object.values(initiativeMap).forEach(initiative => {
      // Get all unique value streams for this initiative
      const valueStreams = [...new Set(initiative.epics.map(e => (e['Value Stream/Org'] || '').trim()))];
      
      // Determine which VS groups this initiative belongs to
      const groupsForInitiative = new Set();
      valueStreams.forEach(vs => {
        const group = vsToGroup[vs];
        if (group) {
          groupsForInitiative.add(group);
        } else {
          // If value stream not in any defined group, put in 'Other'
          console.warn(`[STRUCTURE] Value Stream "${vs}" not in predefined groups, adding to 'Other'`);
          groupsForInitiative.add('Other');
        }
      });
      
      console.log(`[STRUCTURE] Initiative ${initiative.key} spans groups: ${Array.from(groupsForInitiative).join(', ')}`);
      
      // Add initiative to each group it belongs to
      groupsForInitiative.forEach(groupName => {
        // Filter epics to only those in this group's value streams
        const epicsInGroup = initiative.epics.filter(epic => {
          const epicVS = (epic['Value Stream/Org'] || '').trim();
          return vsToGroup[epicVS] === groupName || (!vsToGroup[epicVS] && groupName === 'Other');
        });
        
        if (epicsInGroup.length > 0) {
          // Create a copy of the initiative with only the epics for this group
          const groupedInitiative = {
            ...initiative,
            epics: epicsInGroup,
            totalFeaturePoints: epicsInGroup.reduce((sum, e) => sum + (parseFloat(e['Feature Points']) || 0), 0),
            vsGroup: groupName // Add VS group metadata
          };
          
          initiativesByVSGroup[groupName].push(groupedInitiative);
        }
      });
    });
    
    // Sort within each group by value stream order, then by initiative title
    ['EMA', 'MMGI', 'Other'].forEach(groupName => {
      const vsOrder = vsGroups[groupName];
      
      initiativesByVSGroup[groupName].sort((a, b) => {
        // Get primary value stream (first VS in the VS order list that appears in epics)
        const getFirstVS = (initiative) => {
          for (const vs of vsOrder) {
            if (initiative.epics.some(e => (e['Value Stream/Org'] || '').trim() === vs)) {
              return vs;
            }
          }
          // If no match found, use first epic's VS
          return (initiative.epics[0]?.['Value Stream/Org'] || '').trim();
        };
        
        const vsA = getFirstVS(a);
        const vsB = getFirstVS(b);
        
        const indexA = vsOrder.indexOf(vsA);
        const indexB = vsOrder.indexOf(vsB);
        
        // Sort by VS order
        if (indexA !== indexB) {
          // Handle case where VS not in list (put at end)
          if (indexA === -1) return 1;
          if (indexB === -1) return -1;
          return indexA - indexB;
        }
        
        // If same VS, sort by title
        return a.title.localeCompare(b.title);
      });
    });
    
    // Flatten into single array, maintaining group order: EMA, MMGI, Other
    initiativeGroups = [
      ...initiativesByVSGroup['EMA'],
      ...initiativesByVSGroup['MMGI'],
      ...initiativesByVSGroup['Other']
    ];
    
    console.log(`[STRUCTURE] INFOSEC Grouping Results:`);
    console.log(`[STRUCTURE]   EMA: ${initiativesByVSGroup['EMA'].length} initiatives`);
    console.log(`[STRUCTURE]   MMGI: ${initiativesByVSGroup['MMGI'].length} initiatives`);
    console.log(`[STRUCTURE]   Other: ${initiativesByVSGroup['Other'].length} initiatives`);
    
  } else {
    // Standard sorting by Value Stream alphabetically
    initiativeGroups = Object.values(initiativeMap).sort((a, b) => {
      // Get Value Stream from first epic in each group
      const vsA = (a.epics[0]?.['Value Stream/Org'] || '').trim();
      const vsB = (b.epics[0]?.['Value Stream/Org'] || '').trim();
      
      // First sort by Value Stream
      if (vsA !== vsB) {
        return vsA.localeCompare(vsB);
      }
      
      // If same Value Stream, sort by Initiative Title
      return a.title.localeCompare(b.title);
    });
    console.log(`[STRUCTURE] Sorted ${initiativeGroups.length} groups by Value Stream, then Title`);
  }
  
  console.log(`\n========================================`);
  console.log(`[STRUCTURE] Final Results:`);
  console.log(`[STRUCTURE] Created ${initiativeGroups.length} initiative groups`);
  initiativeGroups.forEach((group, idx) => {
    console.log(`[STRUCTURE]   ${idx + 1}. ${group.title} (${group.key})`);
    console.log(`[STRUCTURE]      - Has Parent: ${group.hasParent}`);
    console.log(`[STRUCTURE]      - Epic Count: ${group.epics.length}`);
    console.log(`[STRUCTURE]      - Feature Points: ${group.totalFeaturePoints}`);
    if (group.vsGroup) {
      console.log(`[STRUCTURE]      - VS Group: ${group.vsGroup}`);
    }
    group.epics.forEach((epic, epicIdx) => {
      console.log(`[STRUCTURE]        Epic ${epicIdx + 1}: ${epic['Key']} - ${epic['Summary']} (${epic['Value Stream/Org']})`);
    });
  });
  console.log(`========================================\n`);
  
  return initiativeGroups;
}

function generateInitiativePresentation(piNumber, initiativeField, initiativeValue, phase, presentationId, piCommitmentFilter) {
  phase = phase || 'FINAL';
  piCommitmentFilter = piCommitmentFilter || 'All';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log(`Generating presentation for ${initiativeField}: ${initiativeValue}, PI ${piNumber}, Phase: ${phase}`);
  
  try {
    // STEP 1: Load all issues for this PI
    ss.toast('Loading PI data...', 'Step 1/5', 5);
    const sheetNames = [];
    const allSheets = ss.getSheets();
    for (let i = 0; i < allSheets.length; i++) {
      const sheetName = allSheets[i].getName();
      if (sheetName.startsWith(`PI ${piNumber} - Iteration`)) {
        sheetNames.push(sheetName);
      }
    }
    
    if (sheetNames.length === 0) {
      throw new Error(`No sheets found for PI ${piNumber}`);
    }
    
    // Get source sheet information for hyperlink
    const sourceSheetName = sheetNames[0];
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    let sourceSheetUrl = ss.getUrl();
    if (sourceSheet) {
      sourceSheetUrl = `${ss.getUrl()}#gid=${sourceSheet.getSheetId()}`;
      console.log(`Source sheet URL: ${sourceSheetUrl}`);
    } else {
      console.log(`Source sheet "${sourceSheetName}" not found, using spreadsheet URL`);
    }
    
    // Generate timestamp once for all slides
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy hh:mm a');
    
    const allIssues = loadAllIssuesFromSheets(sheetNames, ss);
    console.log(`Loaded ${allIssues.length} total issues from ${sheetNames.length} sheets`);
    
    // ═══════════════════════════════════════════════════════════════
    // ✅ NEW: STEP 1.5 - Apply Program Increment Filter (Epic Level)
    // ═══════════════════════════════════════════════════════════════
    console.log(`\n=== Applying Program Increment Filter ===`);
    console.log(`Phase: ${phase}`);
    
    // Log PI values before filtering
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
    
    // Filter EPICS by Program Increment
    let filteredEpics = [];
    
    if (phase === 'INITIAL') {
      filteredEpics = allIssues.filter(issue => {
        if (issue['Issue Type'] !== 'Epic') return false;
        const programIncrement = (issue['Program Increment'] || '').toString().trim();
        return programIncrement === `PI ${piNumber}` || 
               programIncrement === piNumber.toString() || 
               programIncrement.toUpperCase() === 'PRE-PLANNING';
      });
      console.log(`INITIAL PLAN: Including Epics with PI ${piNumber} or "Pre-Planning"`);
      console.log(`  Filtered Epics: ${filteredEpics.length}`);
      
    } else {
      // ✅ FINAL PHASE: Exclude pre-planning items
      filteredEpics = allIssues.filter(issue => {
        if (issue['Issue Type'] !== 'Epic') return false;
        const programIncrement = (issue['Program Increment'] || '').toString().trim();
        return programIncrement === `PI ${piNumber}` || programIncrement === piNumber.toString();
      });
      console.log(`FINAL PLAN: Including ONLY Epics with explicit PI ${piNumber}`);
      console.log(`  Filtered Epics: ${filteredEpics.length}`);
      
      const totalEpics = allIssues.filter(i => i['Issue Type'] === 'Epic').length;
      const excluded = totalEpics - filteredEpics.length;
      if (excluded > 0) {
        console.log(`  - Excluded ${excluded} Epics (wrong PI or Pre-Planning)`);
      }
    }
    
    // Get all Epic keys that passed the filter
    const allowedEpicKeys = new Set(filteredEpics.map(epic => epic['Key']));
    console.log(`Allowed Epic keys: ${allowedEpicKeys.size}`);
    
    // Include ALL issues (Epics + Dependencies) where:
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
    
    // ═══════════════════════════════════════════════════════════════
    // STEP 2: Filter by initiative (INCLUDE BOTH EPICS AND THEIR DEPENDENCIES)
    // ═══════════════════════════════════════════════════════════════
    ss.toast('Filtering by initiative...', 'Step 2/5', 5);
    
    // ✅ NOW USING piFilteredIssues instead of allIssues
    // Get all epics that match the initiative
    const initiativeEpics = piFilteredIssues.filter(issue => 
      issue['Issue Type'] === 'Epic' && issue[initiativeField] === initiativeValue
    );
    
    // Get all epic keys
    const epicKeys = new Set(initiativeEpics.map(e => e['Key']));
    
    // Get dependencies for these epics (where Parent Key matches an epic)
    const dependencies = piFilteredIssues.filter(issue =>
      issue['Issue Type'] === 'Dependency' && epicKeys.has(issue['Parent Key'])
    );
    
    // Combine epics and dependencies
    const filteredIssues = [...initiativeEpics, ...dependencies];
    
    console.log(`Found ${initiativeEpics.length} epics and ${dependencies.length} dependencies for ${initiativeValue}`);
    console.log(`Total issues: ${filteredIssues.length}`);
    
    if (initiativeEpics.length === 0) {
      throw new Error(`No epics found for ${initiativeField}: ${initiativeValue} in PI ${piNumber}${phase === 'FINAL' ? ' (excluding Pre-Planning items)' : ''}`);
    }
    
    // Apply PI Commitment filter if specified
    const finalIssues = piCommitmentFilter !== 'All' 
      ? filteredIssues.filter(issue => 
          issue['Issue Type'] === 'Dependency' || issue['PI Commitment'] === piCommitmentFilter
        )
      : filteredIssues;
    
    console.log(`After PI Commitment filter (${piCommitmentFilter}): ${finalIssues.length} issues`);
    
    // STEP 3: Create presentation from template
    ss.toast('Creating presentation...', 'Step 3/5', 5);
    const presentationName = presentationId || 
      `PI ${piNumber} ${phase} Plan - ${initiativeValue} - ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`;
    
    const templateFile = DriveApp.getFileById('1KzmwWW_KvzBkfyPgEAkgOGiPI9mJCn6K4e5DH4hDjhs');
    const parentFolders = templateFile.getParents();
    const targetFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
    
    const newFile = templateFile.makeCopy(presentationName, targetFolder);
    const presentation = SlidesApp.openById(newFile.getId());
    
    console.log(`Created presentation: ${presentation.getUrl()}`);
    
    // STEP 4: Update title slide
    const titleSlide = presentation.getSlides()[0];
    robustReplaceAllText(titleSlide, '{{ValueStream}}', initiativeValue);
    robustReplaceAllText(titleSlide, '{{PI_Title}}', `PI ${piNumber}`);
    robustReplaceAllText(titleSlide, '{{PI}}', piNumber.toString());
    robustReplaceAllText(titleSlide, '{{Phase}}', phase);
    robustReplaceAllText(titleSlide, '{{PHASE}}', phase);
    
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');
    robustReplaceAllText(titleSlide, '{{Date}}', currentDate);
    
    // Add timestamp with hyperlink to ALL initial slides
    const slides = presentation.getSlides();
    console.log(`Adding timestamp to all initial slides (${slides.length} slides)`);
    slides.forEach((slide, index) => {
      addTimestampToSlide(slide, timestamp, sourceSheetName, sourceSheetUrl);
      console.log(`Added timestamp to slide ${index + 1}`);
    });
    
    // STEP 5: Populate initiative list slide (Slide 2)
    ss.toast('Populating initiative list...', 'Step 4/5', 3);
    if (slides[1]) {
      populateSlide2Initiatives(presentation, piNumber, initiativeField, initiativeValue, finalIssues, slides[1]);
    }
    
    // STEP 6: Remove unnecessary template slides for initiative reports    
    console.log('Removing extra template slides (keeping Slide 6 at index 5 as template)...');
    let currentSlides = presentation.getSlides();
    console.log(`Current slide count: ${currentSlides.length}`);
    
    // Remove slides at index 6 and beyond (keep indices 0-5)
    for (let i = currentSlides.length - 1; i >= 6; i--) {
      try {
        const slideToRemove = currentSlides[i];
        console.log(`Removing slide at index ${i}`);
        slideToRemove.remove();
      } catch (e) {
        console.warn(`Could not remove slide at index ${i}: ${e.message}`);
      }
    }
    
    // Refresh slides array after removal
    currentSlides = presentation.getSlides();
    console.log(`After cleanup: ${currentSlides.length} slides remaining (should be 6: indices 0-5)`);
    console.log(`Slide 6 (index 5) will be used as template for hierarchy generation`);
        
    // STEP 7: Generate hierarchy slides using initiative report structure
    ss.toast('Generating hierarchy slides...', 'Step 5/5', 10);
    console.log('Building initiative report data structures...');
    
    // Structure data by parent/epic grouping (NOT by allocation types)
    const initiativeGroups = structureInitiativeReportData(
      finalIssues,
      [], // allAllocations - empty array means include all
      initiativeField,
      initiativeValue
    );
    
    console.log(`Structured ${initiativeGroups.length} initiative groups for hierarchy slides`);
    
    let totalHierarchySlides = 0;
    
    if (initiativeGroups.length > 0) {
      const slidesBeforeHierarchy = presentation.getSlides().length;
      
      // Generate hierarchy slides with proper formatting
      generateInitiativeHierarchySlides(
        presentation,
        initiativeGroups,
        finalIssues,
        phase,
        initiativeValue 
      );
      
      const slidesAfterHierarchy = presentation.getSlides().length;
      const hierarchySlidesAdded = slidesAfterHierarchy - slidesBeforeHierarchy;
      totalHierarchySlides += hierarchySlidesAdded;
      
      // Add timestamp to all newly created hierarchy slides
      console.log(`Adding timestamp to ${hierarchySlidesAdded} new hierarchy slides`);
      for (let slideIdx = slidesBeforeHierarchy; slideIdx < slidesAfterHierarchy; slideIdx++) {
        addTimestampToSlide(presentation.getSlides()[slideIdx], timestamp, sourceSheetName, sourceSheetUrl);
      }
      
      console.log(`✓ Added ${hierarchySlidesAdded} hierarchy slide(s)`);
    } else {
      console.warn('No initiatives found to populate hierarchy slides');
    }
    
    console.log(`\n✅ Total hierarchy slides added: ${totalHierarchySlides}`);
    console.log(`\n✅ Presentation generation complete!`);
    console.log(`URL: ${presentation.getUrl()}`);
    console.log(`Total slides: ${presentation.getSlides().length}`);
    console.log(`Hierarchy slides generated: ${totalHierarchySlides}`);
    
    ss.toast('✅ Presentation created!', 'Success', 5);
    
    return { 
      url: presentation.getUrl(), 
      id: newFile.getId(),
      slideCount: presentation.getSlides().length
    };
    
  } catch (error) {
    console.error(`Error generating initiative presentation: ${error.message}`);
    console.error(error.stack);
    ss.toast(`❌ Error: ${error.message}`, 'Error', 10);
    throw error;
  }
}

function populateSlide2Initiatives(presentation, piNumber, initiativeField, initiativeValue, allIssues, slide2) {
  try {
    console.log(`\n========== POPULATING SLIDE 2 ==========`);
    console.log(`Initiative Field: ${initiativeField}`);
    console.log(`Initiative Value: ${initiativeValue}`);
    console.log(`Total issues: ${allIssues.length}`);
    
    // Get only epics for this initiative
    const epics = allIssues.filter(issue => 
      issue['Issue Type'] === 'Epic' && issue[initiativeField] === initiativeValue
    );
    console.log(`Found ${epics.length} epics to display on slide 2`);
    
    // Update title placeholders
    robustReplaceAllText(slide2, '{{PI_Title}}', `PI ${piNumber}`);
    robustReplaceAllText(slide2, '{{PI}}', piNumber);
    robustReplaceAllText(slide2, '{{ValueStream}}', initiativeValue);
    
    // Organize INITIATIVES (not epics) by allocation using Maps to store {key, title}
    const allocationGroups = {
      'Product - Feature': new Map(),
      'Product - Compliance': new Map(),
      'Compliance': new Map(),
      'Infosec': new Map(),
      'Tech / Platform': new Map()
    };
    
    epics.forEach(epic => {
      const allocation = epic['Allocation'] || 'Unknown';
      const parentKey = epic['Parent Key'];
      const initiativeTitle = (epic['Initiative Title'] || '').trim();
      
      // CRITICAL: Only include if BOTH Parent Key AND Initiative Title exist
      if (!parentKey || parentKey === '' || !initiativeTitle || initiativeTitle === '') {
        console.log(`Epic ${epic['Key']} skipped - Missing Parent Key or Initiative Title`);
        return;
      }
      
      // Normalize Product - Compliance to Compliance
      const normalizedAllocation = allocation === 'Product - Compliance' ? 'Compliance' : allocation;
      
      if (allocationGroups[normalizedAllocation] !== undefined) {
        // Store with parent key as map key to avoid duplicates
        if (!allocationGroups[normalizedAllocation].has(parentKey)) {
          allocationGroups[normalizedAllocation].set(parentKey, {
            key: parentKey,
            title: initiativeTitle
          });
          console.log(`Added initiative "${initiativeTitle}" (${parentKey}) to ${normalizedAllocation}`);
        }
      } else {
        console.warn(`Unknown allocation: ${allocation} for epic ${epic['Key']}`);
      }
    });
    
    // Special handling for Infosec - also check Portfolio Initiative field
    console.log(`\n----- INFOSEC SPECIAL HANDLING -----`);
    epics.forEach(epic => {
      const portfolioInit = (epic['Portfolio Initiative'] || '').toLowerCase();
      const parentKey = epic['Parent Key'];
      const initiativeTitle = (epic['Initiative Title'] || '').trim();
      
      // Check if Portfolio Initiative contains 'infosec' (case insensitive)
      if (portfolioInit.includes('infosec')) {
        if (parentKey && parentKey !== '' && initiativeTitle && initiativeTitle !== '') {
          if (!allocationGroups['Infosec'].has(parentKey)) {
            allocationGroups['Infosec'].set(parentKey, {
              key: parentKey,
              title: initiativeTitle
            });
            console.log(`Added Infosec initiative via Portfolio field: "${initiativeTitle}" (${parentKey})`);
          }
        }
      }
    });
    
    // Secondary query: Get all epics with Program Initiative starting with "Infosec"
    console.log(`\n----- SECONDARY INFOSEC QUERY (Program Initiative) -----`);
    allIssues.forEach(issue => {
      if (issue['Issue Type'] === 'Epic') {
        const programInit = (issue['Program Initiative'] || '').trim();
        const parentKey = issue['Parent Key'];
        const initiativeTitle = (issue['Initiative Title'] || '').trim();
        
        // Check if Program Initiative starts with "Infosec" (case insensitive)
        if (programInit.toLowerCase().startsWith('infosec')) {
          if (parentKey && parentKey !== '' && initiativeTitle && initiativeTitle !== '') {
            if (!allocationGroups['Infosec'].has(parentKey)) {
              allocationGroups['Infosec'].set(parentKey, {
                key: parentKey,
                title: initiativeTitle
              });
              console.log(`Added Infosec initiative via Program Initiative: "${initiativeTitle}" (${parentKey})`);
            }
          }
        }
      }
    });
    
    // Log final counts
    console.log(`\n----- FINAL ALLOCATION COUNTS -----`);
    Object.keys(allocationGroups).forEach(allocation => {
      console.log(`${allocation}: ${allocationGroups[allocation].size} unique initiatives`);
    });
    
    // Build bullet lists with newlines (one initiative per line)
    const productFeatureList = allocationGroups['Product - Feature'].size > 0 
      ? Array.from(allocationGroups['Product - Feature'].values())
          .sort((a, b) => a.title.localeCompare(b.title))
          .map(init => init.title)
          .join('\n')
      : 'None';
    
    const complianceList = allocationGroups['Compliance'].size > 0
      ? Array.from(allocationGroups['Compliance'].values())
          .sort((a, b) => a.title.localeCompare(b.title))
          .map(init => init.title)
          .join('\n')
      : 'None';
    
    const infosecList = allocationGroups['Infosec'].size > 0
      ? Array.from(allocationGroups['Infosec'].values())
          .sort((a, b) => a.title.localeCompare(b.title))
          .map(init => init.title)
          .join('\n')
      : 'None';
    
    const techPlatformList = allocationGroups['Tech / Platform'].size > 0
      ? Array.from(allocationGroups['Tech / Platform'].values())
          .sort((a, b) => a.title.localeCompare(b.title))
          .map(init => init.title)
          .join('\n')
      : 'None';
    
    console.log(`Product - Feature list:\n${productFeatureList}`);
    console.log(`Compliance list:\n${complianceList}`);
    console.log(`Infosec list:\n${infosecList}`);
    console.log(`Tech / Platform list:\n${techPlatformList}`);
    
    // Replace the placeholders with text (preserves section titles)
    robustReplaceAllText(slide2, '{{ProductFeature_List}}', productFeatureList);
    robustReplaceAllText(slide2, '{{Compliance_List}}', complianceList);
    robustReplaceAllText(slide2, '{{Infosec_List}}', infosecList);
    robustReplaceAllText(slide2, '{{TechPlatform_List}}', techPlatformList);
    
    console.log(`Replaced text placeholders with initiative lists`);
    
    // Now add hyperlinks and styling to each initiative name
    const shapes = slide2.getShapes();
    
    shapes.forEach(shape => {
      try {
        const textRange = shape.getText();
        if (textRange && !textRange.isEmpty()) {
          const fullText = textRange.asString();
          
          // Check each allocation group
          const allocationsToProcess = [
            { name: 'Product - Feature', map: allocationGroups['Product - Feature'], list: productFeatureList },
            { name: 'Compliance', map: allocationGroups['Compliance'], list: complianceList },
            { name: 'Infosec', map: allocationGroups['Infosec'], list: infosecList },
            { name: 'Tech / Platform', map: allocationGroups['Tech / Platform'], list: techPlatformList }
          ];
          
          allocationsToProcess.forEach(allocation => {
            if (allocation.list !== 'None' && fullText.includes(allocation.list)) {
              console.log(`\n----- Processing ${allocation.name} in shape -----`);
              
              const initiatives = Array.from(allocation.map.values()).sort((a, b) => a.title.localeCompare(b.title));
              
              initiatives.forEach(init => {
                // Find each occurrence of this initiative title in the text
                let searchPos = 0;
                while (true) {
                  const titleStart = fullText.indexOf(init.title, searchPos);
                  if (titleStart === -1) break;
                  
                  const titleEnd = titleStart + init.title.length;
                  
                  try {
                    const linkRange = textRange.getRange(titleStart, titleEnd);
                    // CORRECTED: Use modmedrnd.atlassian.net
                    linkRange.getTextStyle()
                      .setLinkUrl(`https://modmedrnd.atlassian.net/browse/${init.key}`)
                      .setForegroundColor('#663399')  // Purple
                      .setBold(false)                 // Not bold
                      .setUnderline(false);           // Not underlined
                    
                    console.log(`✓ Styled & hyperlinked "${init.title}" → https://modmedrnd.atlassian.net/browse/${init.key}`);
                  } catch (styleError) {
                    console.warn(`Could not style "${init.title}" at position ${titleStart}: ${styleError.message}`);
                  }
                  
                  searchPos = titleEnd;
                }
              });
            }
          });
        }
      } catch (e) {
        // Skip shapes that don't have text
      }
    });
    
    // Add feature points chart by Value Stream
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    let sheetName = null;
    
    for (let i = 0; i < allSheets.length; i++) {
      const name = allSheets[i].getName();
      if (name.startsWith(`PI ${piNumber} - Iteration`)) {
        sheetName = name;
        break;
      }
    }
    
    // Remove any remaining placeholder shapes
    const shapesAfter = slide2.getShapes();
    for (let i = shapesAfter.length - 1; i >= 0; i--) {
      const shape = shapesAfter[i];
      try {
        const textRange = shape.getText();
        if (textRange && !textRange.isEmpty()) {
          const text = textRange.asString().trim();
          if (text.startsWith('{{') && text.endsWith('}}') && !text.includes('{{BAR CHART}}')) {
            console.log(`Removing shape with placeholder: ${text.substring(0, 50)}`);
            shape.remove();
          }
        }
      } catch (e) {
        // Ignore
      }
    }
    
    console.log(`========== SLIDE 2 COMPLETE ==========\n`);
    
  } catch (error) {
    console.error(`Error populating Slide 2: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}

function generateInitiativeHierarchySlides(presentation, initiativeGroups, allIssues, phase, initiativeValue) {
  console.log(`[HIERARCHY] Starting hierarchy slide generation for ${initiativeGroups.length} initiative groups`);
  
  if (initiativeGroups.length === 0) {
    console.log(`[HIERARCHY] No initiative groups to display`);
    return;
  }
  
  const slides = presentation.getSlides();
  const templateIndex = 5;
  
  if (slides.length <= templateIndex) {
    console.error(`[HIERARCHY] Template slide not found at index ${templateIndex}`);
    return;
  }
  
  const templateSlide = slides[templateIndex];
  const templateTables = templateSlide.getTables();
  
  if (templateTables.length === 0) {
    console.error(`[HIERARCHY] Template slide has no table`);
    return;
  }
  
  const MAX_INITIATIVES_PER_SLIDE = 3;
  let currentSlide = null;
  let currentTable = null;
  let itemsOnCurrentSlide = 0;  // Count items, not rows
  let pageNum = 0;
  
  // Build dependency map
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
      if (!currentSlide || itemsOnCurrentSlide >= MAX_INITIATIVES_PER_SLIDE) {
        pageNum++;
        console.log(`[HIERARCHY] Creating new slide (page ${pageNum})`);
        
        // Duplicate template
        const newSlide = templateSlide.duplicate();
        const newSlideIndex = presentation.getSlides().indexOf(newSlide);
        addDisclosureToSlide(newSlide);
        const targetIndex = templateIndex + pageNum;
        
        if (newSlideIndex !== targetIndex) {
          newSlide.move(targetIndex);
        }
        
        currentSlide = newSlide;
        
        // Update slide title
        const shapes = currentSlide.getShapes();
        shapes.forEach(shape => {
          try {
            const text = shape.getText();
            if (text && !text.isEmpty()) {
              const textContent = text.asString();
              if (textContent.includes('{{Value Stream}}') || textContent.includes('{{Allocation Type}}') || textContent.includes('{{Portfolio Initiative}}')) {
                text.setText(initiativeValue);
                text.getTextStyle()
                  .setFontSize(24)
                  .setBold(true)
                  .setFontFamily('Lato')
                  .setForegroundColor('#663399');
                console.log(`[HIERARCHY] Updated slide title to: ${initiativeValue}`);
              }
            }
          } catch (e) {
            // Skip non-text shapes
          }
        });
        
        // Get table
        const tables = newSlide.getTables();
        if (tables.length === 0) {
          console.error('[HIERARCHY] No table found on duplicated slide');
          return;
        }
        
        currentTable = tables[0];
        console.log(`[HIERARCHY] Table has ${currentTable.getNumRows()} rows and ${currentTable.getNumColumns()} columns`);
        
        // Clear template data rows (keep header)
        const numRows = currentTable.getNumRows();
        for (let i = numRows - 1; i > 0; i--) {
          currentTable.getRow(i).remove();
        }
        
        // Update header text
        const headerRow = currentTable.getRow(0);
        headerRow.getCell(0).getText().setText('Initiative Details');
        headerRow.getCell(1).getText().setText('Team Information');
        headerRow.getCell(2).getText().setText('Risks / Dependencies');
        
        // Style headers
        for (let c = 0; c < 3; c++) {
          const headerCell = headerRow.getCell(c);
          headerCell.getFill().setSolidFill('#663399');
          
          const headerText = headerCell.getText();
          headerText.getTextStyle()
            .setForegroundColor('#FFFFFF')
            .setBold(true)
            .setFontSize(11)
            .setFontFamily('Lato');
          
          if (c === 0) {
            headerText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
          } else {
            headerText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
          }
        }
        
        // Set column widths
        try {
          currentTable.getColumn(0).setWidth(342);
          currentTable.getColumn(1).setWidth(274);
          currentTable.getColumn(2).setWidth(144);
          console.log('[HIERARCHY] ✓ Set column widths: 342pts, 274pts, 144pts');
        } catch (e) {
          console.error(`[HIERARCHY] Error setting column widths: ${e.message}`);
        }
        
        itemsOnCurrentSlide = 0;
      }
      
      // Add initiative/epic to slide
      addInitiativeToTable(currentTable, initiative, dependencyMap);
      itemsOnCurrentSlide++;
      
      console.log(`[HIERARCHY] Items on current slide: ${itemsOnCurrentSlide}/${MAX_INITIATIVES_PER_SLIDE}`);
      
    } catch (error) {
      console.error(`Error processing initiative ${initiative.key}: ${error.message}`);
      console.error(error.stack);
    }
  });
  
  // Remove template slide
  try {
    templateSlide.remove();
    console.log(`[HIERARCHY] ✓ Removed template slide`);
  } catch (error) {
    console.warn(`[HIERARCHY] Could not remove template slide: ${error.message}`);
  }
  
  console.log(`[HIERARCHY] ✅ Created ${pageNum} page(s)`);
}

function addInitiativeToTable(table, initiative, dependencyMap) {
  console.log(`\n[ADD INIT] Adding: ${initiative.title}`);
  console.log(`[ADD INIT] Key: ${initiative.key}`);
  console.log(`[ADD INIT] Has Parent: ${initiative.hasParent}`);
  console.log(`[ADD INIT] Epics: ${initiative.epics.length}, FP: ${initiative.totalFeaturePoints}`);
  
  const row = table.appendRow();
  
  // Make transparent
  for (let c = 0; c < 3; c++) {
    row.getCell(c).getFill().setSolidFill('#FFFFFF', 0);
  }
  
  if (initiative.hasParent) {
    // ============================================
    // INITIATIVE (HAS PARENT): Normal Layout
    // ============================================
    
    // COLUMN 0: Initiative Details
    const col0Cell = row.getCell(0);
    const col0Text = col0Cell.getText();
    
    let col0Content = '';
    col0Content += initiative.title + '\n\n';
    const titleEnd = initiative.title.length;
    
    console.log(`[COL 0] Displaying initiative NAME: "${initiative.title}" (NOT key: ${initiative.key})`);
    
    // Add merged technical perspective (prose, not bullets)
    const allPerspectives = initiative.epics
      .map(e => e['Technical Perspective'])
      .filter(p => p && p.trim() !== '');
    const mergedPerspective = allPerspectives.join(' ');
    
    const perspStart = col0Content.length;
    if (mergedPerspective) {
      col0Content += mergedPerspective;
    }
    
    col0Text.setText(col0Content.trim());
    
    // Style title: BOLD, PURPLE, NOT underlined
    // IMPORTANT: Set link URL FIRST, then override the default blue/underlined link style
    const titleRange = col0Text.getRange(0, titleEnd);
    titleRange.getTextStyle()
      .setLinkUrl(`https://modmedrnd.atlassian.net/browse/${initiative.key}`);
    
    // Now override the default link styling
    titleRange.getTextStyle()
      .setFontSize(9)
      .setBold(true)              // BOLD
      .setItalic(false)
      .setFontFamily('Lato')
      .setForegroundColor('#663399')  // Purple (overrides blue)
      .setUnderline(false);            // No underline (overrides underline)
    
    console.log(`[COL 0] ✓ Title: "${initiative.title}" → Hyperlinked to: ${initiative.key}`);
    
    // Style perspective
    if (mergedPerspective) {
      const perspEnd = col0Text.asString().length;
      const perspRange = col0Text.getRange(perspStart, perspEnd);
      perspRange.getTextStyle()
        .setFontSize(8)
        .setBold(false)
        .setItalic(false)
        .setFontFamily('Lato')
        .setForegroundColor('#000000');
    }
    
    col0Text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    
    // COLUMN 1: Team Information (grouped by value stream)
    const col1Cell = row.getCell(1);
    const col1Text = col1Cell.getText();
    
    const vsMap = {};
    initiative.epics.forEach(epic => {
      const vs = epic['Value Stream/Org'] || 'No Value Stream';
      if (!vsMap[vs]) {
        vsMap[vs] = {
          committed: [],
          nonCommitted: []
        };
      }
      
      // FIXED: Commitment logic
      // Committed = "Committed" OR "Committed After Plan"
      // Everything else = Non-Committed
      const commitment = (epic['PI Commitment'] || '').trim();
      if (commitment === 'Committed' || commitment === 'Committed After Plan') {
        vsMap[vs].committed.push(epic);
        console.log(`[COMMITMENT] Epic ${epic['Key']}: "${commitment}" → Committed`);
      } else {
        vsMap[vs].nonCommitted.push(epic);
        console.log(`[COMMITMENT] Epic ${epic['Key']}: "${commitment}" → Non-Committed`);
      }
    });
    
    const valueStreams = Object.keys(vsMap).sort();
    let col1Content = '';
    
    valueStreams.forEach((vs, idx) => {
      if (idx > 0) col1Content += '\n\n';
      
      col1Content += vs + '\n';
      
      const allEpics = [...vsMap[vs].committed, ...vsMap[vs].nonCommitted];
      const descriptions = allEpics
        .map(e => e['Technical Perspective'])
        .filter(p => p && p.trim() !== '');
      
      if (descriptions.length > 0) {
        col1Content += descriptions.join(' ') + '\n';
      }
      
      // Show epic KEYS (not names/summary) - comma-separated on same line
      if (vsMap[vs].committed.length > 0) {
        const committedKeys = vsMap[vs].committed.map(epic => {
          if (piCommitmentFilter === 'All') {
            const status = (epic['PI Commitment'] || '').trim();
            return `${epic['Key']} - ${status || 'Blank'}`;
          }
          return epic['Key'];
        }).join(', ');
        col1Content += '\nCommitted: ' + committedKeys + '\n';
      }

      if (vsMap[vs].nonCommitted.length > 0) {
        const nonCommittedKeys = vsMap[vs].nonCommitted.map(epic => {
          if (piCommitmentFilter === 'All') {
            const status = (epic['PI Commitment'] || '').trim();
            return `${epic['Key']} - ${status || 'Blank'}`;
          }
          return epic['Key'];
        }).join(', ');
        col1Content += '\nNon-Committed: ' + nonCommittedKeys + '\n';
      }
    });
    
    col1Text.setText(col1Content.trim());
    col1Text.getTextStyle()
      .setFontSize(8)
      .setFontFamily('Lato')
      .setForegroundColor('#000000')
      .setBold(false);
    
    // Apply styling to middle column
    const finalText = col1Content.trim();
    
    // Bold value stream names
    valueStreams.forEach(vs => {
      let searchPos = 0;
      while (true) {
        const vsStart = finalText.indexOf(vs, searchPos);
        if (vsStart === -1) break;
        
        const vsEnd = vsStart + vs.length;
        try {
          col1Text.getRange(vsStart, vsEnd).getTextStyle().setBold(true);
        } catch (e) {
          console.warn(`Could not bold VS at position ${vsStart}`);
        }
        searchPos = vsEnd;
      }
    });
    
    // Make "Non-Committed:" RED
    let searchPos = 0;
    while (true) {
      const ncStart = finalText.indexOf('Non-Committed:', searchPos);
      if (ncStart === -1) break;
      
      const ncEnd = ncStart + 'Non-Committed:'.length;
      try {
        col1Text.getRange(ncStart, ncEnd).getTextStyle().setForegroundColor('#FF0000');
      } catch (e) {
        console.warn(`Could not color Non-Committed at position ${ncStart}`);
      }
      searchPos = ncEnd;
    }
    
    // Hyperlink epic KEYS to their JIRA URLs: BOLD, PURPLE, NOT underlined
    // IMPORTANT: Set link URL FIRST, then override the default blue/underlined link style
    initiative.epics.forEach(epic => {
      const epicKey = epic['Key'];
      
      let searchPos = 0;
      while (true) {
        const keyPos = finalText.indexOf(epicKey, searchPos);
        if (keyPos === -1) break;
        
        const keyEnd = keyPos + epicKey.length;
        try {
          const keyRange = col1Text.getRange(keyPos, keyEnd);
          
          // Set link first
          keyRange.getTextStyle()
            .setLinkUrl(`https://modmedrnd.atlassian.net/browse/${epicKey}`);
          
          // Then override the default link styling
          keyRange.getTextStyle()
            .setForegroundColor('#663399')  // Purple (overrides blue)
            .setBold(true)                  // BOLD
            .setUnderline(false);           // Not underlined (overrides underline)
          
          console.log(`[COL 1] ✓ Hyperlinked epic key: ${epicKey}`);
        } catch (e) {
          console.warn(`Could not hyperlink "${epicKey}" at position ${keyPos}`);
        }
        searchPos = keyEnd;
      }
    });
    
    col1Text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    
  } else {
    // ============================================
    // STANDALONE EPIC (NO PARENT): Special Layout
    // Column 0: BLANK
    // Column 1: Epic info goes here
    // ============================================
    
    console.log(`[SPECIAL] Standalone epic - Column 0 blank, data in Column 1`);
    
    // COLUMN 0: Leave BLANK
    const col0Cell = row.getCell(0);
    col0Cell.getText().setText('');
    console.log(`[COL 0] ✓ Left blank for standalone epic`);
    
    // COLUMN 1: Epic information goes here instead
    const col1Cell = row.getCell(1);
    const col1Text = col1Cell.getText();
    
    const epic = initiative.epics[0];
    const vs = epic ? (epic['Value Stream/Org'] || 'No Value Stream') : 'No Value Stream';
    
    // FIXED: Epic info ABOVE value stream
    // Order: Epic name -> Technical perspective -> Value stream
    let col1Content = '';
    
    // Add epic SUMMARY/NAME (NOT key)
    const epicName = epic ? (epic['Summary'] || epic['Key']) : initiative.title;
    console.log(`[COL 1] Displaying epic NAME: "${epicName}" (NOT key: ${initiative.key})`);
    
    col1Content += epicName + '\n\n';
    const epicNameStart = 0;
    const epicNameEnd = epicName.length;
    
    // Add technical perspective
    const techPerspective = epic ? (epic['Technical Perspective'] || '') : '';
    const perspStart = col1Content.length;
    if (techPerspective && techPerspective.trim() !== '') {
      col1Content += techPerspective + '\n\n';
    }
    
    // Add value stream at the END (below epic info)
    const vsLineStart = col1Content.length;
    col1Content += 'Value stream: ' + vs;
    const vsLabelEnd = vsLineStart + 'Value stream:'.length;
    
    col1Text.setText(col1Content.trim());
    col1Text.getTextStyle()
      .setFontSize(8)
      .setFontFamily('Lato')
      .setForegroundColor('#000000')
      .setBold(false);
    
    // Hyperlink epic name: BOLD, PURPLE, NOT underlined
    // IMPORTANT: Set link URL FIRST, then override the default blue/underlined link style
    try {
      const epicNameRange = col1Text.getRange(epicNameStart, epicNameEnd);
      epicNameRange.getTextStyle()
        .setLinkUrl(`https://modmedrnd.atlassian.net/browse/${initiative.key}`);
      
      // Now override the default link styling
      epicNameRange.getTextStyle()
        .setForegroundColor('#663399')  // Purple (overrides blue)
        .setBold(true)                  // BOLD
        .setUnderline(false);           // Not underlined (overrides underline)
      
      console.log(`[COL 1] ✓ Epic name: "${epicName}" → Hyperlinked to: ${initiative.key}`);
    } catch (e) {
      console.warn(`Could not hyperlink epic name: ${e.message}`);
    }
    
    // Bold "Value stream:" label
    try {
      const vsLabelRange = col1Text.getRange(vsLabelStart, vsLabelEnd);
      vsLabelRange.getTextStyle().setBold(true);
      console.log(`[COL 1] ✓ Bolded "Value stream:" label`);
    } catch (e) {
      console.warn(`Could not bold "Value stream:" label: ${e.message}`);
    }
    
    col1Text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  }
  
  // ============================================
  // COLUMN 2: Risks / Dependencies (SAME FOR BOTH)
  // ============================================
  const col2Cell = row.getCell(2);
  const col2Text = col2Cell.getText();
  
  let col2Content = '';
  
  initiative.epics.forEach(epic => {
    const rag = (epic['RAG'] || '').toLowerCase().trim();
    const ragNote = (epic['RAG Note'] || '').trim();
    
    if (rag && rag !== 'green' && rag !== '' && ragNote && ragNote !== '') {
      col2Content += `${getRagIcon(rag)} ${ragNote}\n`;
    }
    
    const deps = dependencyMap[epic['Key']] || [];
    deps.forEach(dep => {
      const depVS = (dep['Depends on Valuestream'] || dep['Value Stream/Org'] || 'Unknown').trim();
      const depTitle = (dep['Summary'] || dep['Key']).trim();
      col2Content += `🔗 ${depVS}: ${depTitle}\n`;
    });
  });
  
  col2Text.setText(col2Content.trim());
  col2Text.getTextStyle()
    .setFontSize(8)
    .setFontFamily('Lato')
    .setForegroundColor('#000000')
    .setBold(false);
  
  col2Text.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
  
  console.log(`[ADD INIT] ✓ Completed\n`);
}

/**
 * Helper function to get RAG icon
 */
function getRagIcon(rag) {
  const ragLower = rag.toLowerCase();
  if (ragLower.includes('red')) return '🔴';
  if (ragLower.includes('amber') || ragLower.includes('yellow')) return '🟡';
  if (ragLower.includes('green')) return '🟢';
  return '⚪';
}

// Helper to get all Program Initiatives for a PI
function getAvailableInitiatives(piNumber, initiativeField) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const initiatives = new Set();
  
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const sheetName = allSheets[i].getName();
    if (sheetName.startsWith(`PI ${piNumber} - Iteration`)) {
      const sheet = allSheets[i];
      const lastRow = sheet.getLastRow();
      if (lastRow < DATA_START_ROW) continue;
      
      const headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
      const fieldIndex = headers.indexOf(initiativeField);
      
      if (fieldIndex !== -1) {
        const data = sheet.getRange(DATA_START_ROW, fieldIndex + 1, lastRow - DATA_START_ROW + 1, 1).getValues();
        data.forEach(row => {
          const value = String(row[0]).trim();
          if (value && value !== '' && value !== 'null') {
            initiatives.add(value);
          }
        });
      }
    }
  }
  
  return Array.from(initiatives).sort();
}

function structureHierarchicalData(issues, targetAllocations) {
  console.log(`[STRUCTURE] Starting hierarchical data structuring with ${issues.length} issues`);
  console.log(`[STRUCTURE] Target allocations: ${targetAllocations.join(', ')}`);
  
  const hierarchy = {};
  
  // Filter to only epics that match target allocations
  const epics = issues.filter(issue => {
    if (issue['Issue Type'] !== 'Epic') return false;
    const allocation = issue['Allocation'];
    const matches = targetAllocations.includes(allocation);
    if (!matches) {
      console.log(`[STRUCTURE] Skipping epic ${issue['Key']} - allocation '${allocation}' not in target list`);
    }
    return matches;
  });
  
  console.log(`[STRUCTURE] Found ${epics.length} epics matching target allocations`);
  
  // Build issue map for parent lookup
  const issueMap = {};
  issues.forEach(issue => {
    issueMap[issue['Key']] = issue;
  });
  console.log(`[STRUCTURE] Built issue map with ${Object.keys(issueMap).length} issues`);
  
  // Group epics by parent initiative
  const initiativeMap = {}; // Key = parent key, Value = initiative data + child epics
  
  epics.forEach(epic => {
    const parentKey = epic['Parent Key'];
    
    if (!parentKey || parentKey === '') {
      console.warn(`[STRUCTURE] Epic ${epic['Key']} has no parent key, skipping`);
      return;
    }
    
    if (!initiativeMap[parentKey]) {
      // Create initiative entry from parent issue
      const parentIssue = issueMap[parentKey];
      
      if (!parentIssue) {
        console.warn(`[STRUCTURE] Parent issue ${parentKey} not found in issue map, using key only`);
      }
      
      initiativeMap[parentKey] = {
        key: parentKey,
        summary: parentIssue ? (parentIssue['Summary'] || parentIssue['Initiative Title'] || parentKey) : parentKey,
        businessValue: parentIssue ? (parentIssue['Business Value'] || '') : '',
        successCriteria: parentIssue ? (parentIssue['Benefit Hypothesis'] || parentIssue['Acceptance Criteria'] || '') : '',
        priority: parentIssue ? (parentIssue['PI Commitment'] || '') : '',
        status: parentIssue ? (parentIssue['Status'] || '') : '',
        piObjective: parentIssue ? (parentIssue['PI Objective'] || '') : '',
        allocation: parentIssue ? (parentIssue['Allocation'] || '') : '',
        portfolioInitiative: epic['Portfolio Initiative'] || 'No Portfolio Initiative',
        programInitiative: epic['Program Initiative'] || 'No Program Initiative',
        epics: [],
        totalFeaturePoints: 0,
        epicCount: 0
      };
      
      console.log(`[STRUCTURE] Created initiative entry for ${parentKey}: "${initiativeMap[parentKey].summary}"`);
    }
    
    // Add epic to initiative and aggregate
    initiativeMap[parentKey].epics.push(epic);
    initiativeMap[parentKey].epicCount++;
    
    const featurePoints = parseFloat(epic['Feature Points']) || 0;
    initiativeMap[parentKey].totalFeaturePoints += featurePoints;
    
    console.log(`[STRUCTURE] Added epic ${epic['Key']} to initiative ${parentKey} (FP: ${featurePoints}, Total: ${initiativeMap[parentKey].totalFeaturePoints})`);
  });
  
  console.log(`[STRUCTURE] Created ${Object.keys(initiativeMap).length} initiative groups`);
  
  // Organize initiatives into Portfolio > Program hierarchy
  Object.values(initiativeMap).forEach(initiative => {
    const portfolioInitiative = initiative.portfolioInitiative;
    const programInitiative = initiative.programInitiative;
    
    // Initialize portfolio level
    if (!hierarchy[portfolioInitiative]) {
      hierarchy[portfolioInitiative] = {
        name: portfolioInitiative,
        programs: {}
      };
      console.log(`[STRUCTURE] Created portfolio: ${portfolioInitiative}`);
    }
    
    // Initialize program level
    if (!hierarchy[portfolioInitiative].programs[programInitiative]) {
      hierarchy[portfolioInitiative].programs[programInitiative] = {
        name: programInitiative,
        initiatives: []
      };
      console.log(`[STRUCTURE] Created program: ${programInitiative} under ${portfolioInitiative}`);
    }
    
    // Add initiative to program
    hierarchy[portfolioInitiative].programs[programInitiative].initiatives.push(initiative);
    
    console.log(`[STRUCTURE] Added initiative ${initiative.key} to ${portfolioInitiative} > ${programInitiative} (${initiative.epicCount} epics, ${initiative.totalFeaturePoints} FP)`);
  });
  
  const portfolioCount = Object.keys(hierarchy).length;
  console.log(`[STRUCTURE] Created hierarchy with ${portfolioCount} portfolio initiative(s)`);
  
  // Log summary statistics
  Object.keys(hierarchy).forEach(portfolioName => {
    const portfolio = hierarchy[portfolioName];
    const programCount = Object.keys(portfolio.programs).length;
    let totalInitiatives = 0;
    
    Object.values(portfolio.programs).forEach(program => {
      totalInitiatives += program.initiatives.length;
    });
    
    console.log(`[STRUCTURE] ${portfolioName}: ${programCount} program(s), ${totalInitiatives} initiative(s)`);
  });
  
  return hierarchy;
}
function calculateInitiativeReportPageCount(programHierarchy) {
  let totalInitiatives = 0;
  Object.values(programHierarchy).forEach(programData => {
    if (programData.initiatives && Array.isArray(programData.initiatives)) {
      totalInitiatives += programData.initiatives.length;
    }
  });
  return Math.ceil(totalInitiatives / 3); // 3 initiatives per page for initiative reports
}

function getRagIcon(rag) {
  const ragLower = rag.toLowerCase();
  if (ragLower.includes('red')) return '🔴';
  if (ragLower.includes('amber') || ragLower.includes('yellow')) return '🟡';
  if (ragLower.includes('green')) return '🟢';
  return '⚪';
}

function generateInitiativeReport(piNumber, initiativeField, initiativeValue, piCommitmentFilter, includePrePlanning) {
  console.log(`[INITIATIVE REPORT] Starting report generation`);
  console.log(`[INITIATIVE REPORT] PI: ${piNumber}, Field: ${initiativeField}, Value: ${initiativeValue}`);
  console.log(`[INITIATIVE REPORT] PI Commitment Filter: ${piCommitmentFilter}`);
  console.log(`[INITIATIVE REPORT] Include Pre-Planning: ${includePrePlanning}`);
  
  // Call the main function with phase='FINAL' and no custom presentation ID
  return generateInitiativePresentation(
    piNumber,           // PI number
    initiativeField,    // 'Program Initiative' or 'Portfolio Initiative'
    initiativeValue,    // The specific initiative name
    'FINAL',           // Always FINAL for initiative reports
    null,              // Let function auto-generate presentation name
    piCommitmentFilter, // Commitment filter
    includePrePlanning  // Include Pre-Planning toggle
  );
}
