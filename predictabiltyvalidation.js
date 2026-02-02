function generatePredictabilityValidation(piNumber, valueStream) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log(`\n${'='.repeat(80)}`);
  console.log(`PREDICTABILITY VALIDATION - PI ${piNumber} - ${valueStream}`);
  console.log(`${'='.repeat(80)}\n`);
  
  // Create or clear validation sheet
  const sheetName = `Validation - PI ${piNumber} - ${valueStream}`;
  let validationSheet = ss.getSheetByName(sheetName);
  
  if (validationSheet) {
    validationSheet.clear();
  } else {
    validationSheet = ss.insertSheet(sheetName);
  }
  
  // Get source data
  const sourceSheetName = `PI ${piNumber} - Iteration 6`;
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error(`Source sheet not found: ${sourceSheetName}`);
  }
  
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData[3]; // Row 4 (0-indexed as 3)
  const dataRows = sourceData.slice(4); // Start from row 5
  
  console.log(`Source sheet: ${sourceSheetName}`);
  console.log(`Total rows in source: ${dataRows.length}`);
  
  // Get column indices
  const colIndices = {
    key: headers.indexOf('Key'),
    summary: headers.indexOf('Summary'),
    issueType: headers.indexOf('Issue Type'),
    valueStream: headers.indexOf('Value Stream/Org'),
    status: headers.indexOf('Status'),
    piCommitment: headers.indexOf('PI Commitment'),
    allocation: headers.indexOf('Allocation'),
    piObjectiveStatus: headers.indexOf('PI Objective Status'),
    storyPoints: headers.indexOf('Story Point Completion'),
    businessValue: headers.indexOf('Business Value'),
    actualValue: headers.indexOf('Actual Value'),
    scrumTeam: headers.indexOf('Scrum Team'),
    programInitiative: headers.indexOf('Program Initiative')
  };
  
  // Track exclusions with reasons
  const exclusions = {
    notEpic: [],
    wrongValueStream: [],
    emptyKey: [],
    total: 0
  };
  
  // Filter to epics for this value stream with exclusion tracking
  const allEpics = [];
  
  dataRows.forEach(row => {
    // Check if empty key
    if (!row[colIndices.key] || row[colIndices.key].toString().trim() === '') {
      exclusions.emptyKey.push({
        summary: row[colIndices.summary] || '(No summary)',
        reason: 'Missing JIRA key - cannot identify issue'
      });
      exclusions.total++;
      return;
    }
    
    // Check if not an epic
    if (row[colIndices.issueType] !== 'Epic') {
      exclusions.notEpic.push({
        key: row[colIndices.key],
        issueType: row[colIndices.issueType] || '(Unknown)',
        summary: row[colIndices.summary],
        reason: `Issue Type = "${row[colIndices.issueType]}" (only Epics included in predictability metrics)`
      });
      exclusions.total++;
      return;
    }
    
    // Check if wrong value stream
    if (row[colIndices.valueStream] !== valueStream) {
      exclusions.wrongValueStream.push({
        key: row[colIndices.key],
        valueStream: row[colIndices.valueStream] || '(None)',
        summary: row[colIndices.summary],
        reason: `Value Stream = "${row[colIndices.valueStream]}" (validating only "${valueStream}")`
      });
      exclusions.total++;
      return;
    }
    
    // Include this epic
    allEpics.push({
      key: row[colIndices.key],
      summary: row[colIndices.summary],
      status: (row[colIndices.status] || '').toString().trim(),
      piCommitment: (row[colIndices.piCommitment] || '').toString().trim(),
      allocation: (row[colIndices.allocation] || '').toString().trim(),
      piObjectiveStatus: (row[colIndices.piObjectiveStatus] || '').toString().trim(),
      storyPoints: parseFloat(row[colIndices.storyPoints]) || 0,
      businessValue: parseFloat(row[colIndices.businessValue]) || 0,
      actualValue: parseFloat(row[colIndices.actualValue]) || 0,
      scrumTeam: colIndices.scrumTeam !== -1 ? (row[colIndices.scrumTeam] || '').toString().trim() : '',
           programInitiative: (row[colIndices.programInitiative] || '').toString().trim() 
    });
  });
  
  console.log(`✓ Found ${allEpics.length} epics for ${valueStream}`);
  console.log(`✓ Excluded ${exclusions.total} items:`);
  console.log(`  - Not Epic: ${exclusions.notEpic.length}`);
  console.log(`  - Wrong Value Stream: ${exclusions.wrongValueStream.length}`);
  console.log(`  - Empty Key: ${exclusions.emptyKey.length}`);
  
  // Set up sheet header
  let currentRow = 1;
  
  validationSheet.getRange(currentRow, 1).setValue(`PREDICTABILITY VALIDATION REPORT`)
    .setFontFamily('Comfortaa')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#1a73e8');
  currentRow++;
  
  validationSheet.getRange(currentRow, 1).setValue(`PI ${piNumber} - ${valueStream}`)
    .setFontFamily('Comfortaa')
    .setFontSize(14)
    .setFontWeight('bold');
  currentRow++;
  
  validationSheet.getRange(currentRow, 1).setValue(`Generated: ${new Date().toLocaleString()}`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontColor('#666666');
  currentRow += 2;
  
  // Summary statistics
  validationSheet.getRange(currentRow, 1).setValue(`📊 DATA SUMMARY`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#1a73e8');
  currentRow++;
  
  validationSheet.getRange(currentRow, 1).setValue(`Total Epics Included: ${allEpics.length}`)
    .setFontFamily('Comfortaa')
    .setFontSize(11);
  currentRow++;
  
  validationSheet.getRange(currentRow, 1).setValue(`Total Items Excluded: ${exclusions.total}`)
    .setFontFamily('Comfortaa')
    .setFontSize(11)
    .setFontColor('#d93025');
  currentRow += 2;
  
  // Generate validation for each table
  currentRow = writeTable2Validation(validationSheet, currentRow, allEpics, piNumber, valueStream);
  currentRow = writeTable3Validation(validationSheet, currentRow, allEpics, piNumber, valueStream);
  currentRow = writeTable4Validation(validationSheet, currentRow, allEpics, piNumber, valueStream);
  currentRow = writeTable5Validation(validationSheet, currentRow, allEpics, piNumber, valueStream);
  currentRow = writeTable6Validation(validationSheet, currentRow, allEpics, piNumber, valueStream);
  
  // Add exclusion section at the end
  currentRow = writeExclusionSection(validationSheet, currentRow, exclusions);
  
  // Format the sheet
  validationSheet.setColumnWidth(1, 150);
  validationSheet.setColumnWidth(2, 400);
  validationSheet.setColumnWidth(3, 150);
  validationSheet.setColumnWidth(4, 150);
  
  console.log(`\n✅ Validation sheet created: ${sheetName}`);
}

/**
 * NEW FUNCTION - Writes exclusion details to validation sheet
 * ADD this function to predictatibilityvalidatio.gs
 */
function writeExclusionSection(sheet, startRow, exclusions) {
  let currentRow = startRow + 2;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`🚫 EXCLUDED ITEMS`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setFontColor('#d93025');
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`These items were excluded from the validation report:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontColor('#666666');
  currentRow += 2;
  
  // Not Epic section
  if (exclusions.notEpic.length > 0) {
    sheet.getRange(currentRow, 1).setValue(`❌ Not Epic Type (${exclusions.notEpic.length} items)`)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontWeight('bold');
    currentRow++;
    
    // Headers
    sheet.getRange(currentRow, 1, 1, 4).setValues([['Key', 'Issue Type', 'Summary', 'Reason']])
      .setFontFamily('Comfortaa')
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground('#f8f9fa');
    currentRow++;
    
    // Data
    exclusions.notEpic.forEach(item => {
      sheet.getRange(currentRow, 1, 1, 4).setValues([[item.key, item.issueType, item.summary, item.reason]])
        .setFontFamily('Comfortaa')
        .setFontSize(9);
      currentRow++;
    });
    
    currentRow += 1;
  }
  
  // Wrong Value Stream section
  if (exclusions.wrongValueStream.length > 0) {
    sheet.getRange(currentRow, 1).setValue(`❌ Wrong Value Stream (${exclusions.wrongValueStream.length} items)`)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontWeight('bold');
    currentRow++;
    
    // Headers
    sheet.getRange(currentRow, 1, 1, 4).setValues([['Key', 'Value Stream', 'Summary', 'Reason']])
      .setFontFamily('Comfortaa')
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground('#f8f9fa');
    currentRow++;
    
    // Data
    exclusions.wrongValueStream.forEach(item => {
      sheet.getRange(currentRow, 1, 1, 4).setValues([[item.key, item.valueStream, item.summary, item.reason]])
        .setFontFamily('Comfortaa')
        .setFontSize(9);
      currentRow++;
    });
    
    currentRow += 1;
  }
  
  // Empty Key section
  if (exclusions.emptyKey.length > 0) {
    sheet.getRange(currentRow, 1).setValue(`❌ Missing JIRA Key (${exclusions.emptyKey.length} items)`)
      .setFontFamily('Comfortaa')
      .setFontSize(11)
      .setFontWeight('bold');
    currentRow++;
    
    // Headers
    sheet.getRange(currentRow, 1, 1, 2).setValues([['Summary', 'Reason']])
      .setFontFamily('Comfortaa')
      .setFontWeight('bold')
      .setFontSize(9)
      .setBackground('#f8f9fa');
    currentRow++;
    
    // Data
    exclusions.emptyKey.forEach(item => {
      sheet.getRange(currentRow, 1, 1, 2).setValues([[item.summary, item.reason]])
        .setFontFamily('Comfortaa')
        .setFontSize(9);
      currentRow++;
    });
  }
  
  return currentRow;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLE 2: ALLOCATION DISTRIBUTION VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

function writeTable2Validation(sheet, startRow, allEpics, piNumber, valueStream) {
  let currentRow = startRow;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`TABLE 2: ALLOCATION DISTRIBUTION`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  currentRow++;
  
  // Business rules
  sheet.getRange(currentRow, 1).setValue(`Business Rules:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold');
  currentRow++;
  
  const rules2 = [
    '• Include: Status = "Closed" OR "Pending Acceptance"',
    '• Group by: Allocation Type',
    '• Sum: Story Point Completion by allocation',
    '• Allocation Types: Product-Feature, Tech/Platform, Compliance, Infosec, Quality, KLO'
  ];
  
  rules2.forEach(rule => {
    sheet.getRange(currentRow, 1).setValue(rule)
      .setFontFamily('Comfortaa')
      .setFontSize(9)
      .setFontColor('#666666');
    currentRow++;
  });
  currentRow++;
  
  // Filter epics - include both Closed AND Pending Acceptance
  const closedEpics = allEpics.filter(e => {
    const normalizedStatus = e.status.toLowerCase();
    return normalizedStatus === 'closed' || normalizedStatus === 'pending acceptance';
  });
  
  // Write column headers
  const headerRow = currentRow;
  sheet.getRange(headerRow, 1, 1, 7).setValues([[
    'Key', 'Summary', 'Allocation', 'Story Points', 'Status', 'Included?', 'Reason'
  ]]).setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  currentRow++;
  
  // Write epic details
  allEpics.forEach(epic => {
    const normalizedStatus = epic.status.toLowerCase();
    const included = normalizedStatus === 'closed' || normalizedStatus === 'pending acceptance';
    let reason = '';
    
    if (included) {
      reason = `Status = "${epic.status}"`;
    } else {
      reason = `Status = "${epic.status}" (excluded - not Closed or Pending Acceptance)`;
    }
    
    sheet.getRange(currentRow, 1, 1, 7).setValues([[
      epic.key,
      epic.summary,
      epic.allocation,
      epic.storyPoints,
      epic.status,
      included ? 'YES' : 'NO',
      reason
    ]]).setFontFamily('Comfortaa').setFontSize(9);
    
    if (included) {
      sheet.getRange(currentRow, 6).setBackground('#d4edda').setFontWeight('bold');
    } else {
      sheet.getRange(currentRow, 6).setBackground('#f8d7da');
    }
    
    currentRow++;
  });
  
  // Summary totals
  currentRow++;
  const allocationTotals = {};
  const allocationTypes = ['Product - Feature', 'Tech / Platform', 'Compliance', 'Infosec', 'Quality'];
  
  allocationTypes.forEach(type => allocationTotals[type] = 0);
  
  closedEpics.forEach(epic => {
    if (allocationTotals.hasOwnProperty(epic.allocation)) {
      allocationTotals[epic.allocation] += epic.storyPoints;
    }
  });
  
  sheet.getRange(currentRow, 1).setValue('TOTALS BY ALLOCATION:')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  Object.entries(allocationTotals).forEach(([allocation, total]) => {
    sheet.getRange(currentRow, 1, 1, 2).setValues([[allocation, total]])
      .setFontFamily('Comfortaa')
      .setFontSize(9);
    currentRow++;
  });
  
  sheet.getRange(currentRow, 1).setValue(`Total Epics Included: ${closedEpics.length} of ${allEpics.length}`)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow += 3;
  
  return currentRow;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLE 3: ART VELOCITY VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

function writeTable3Validation(sheet, startRow, allEpics, piNumber, valueStream) {
  let currentRow = startRow;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`TABLE 3: ART VELOCITY`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  currentRow++;
  
  // Business rules
  sheet.getRange(currentRow, 1).setValue(`Business Rules:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold');
  currentRow++;
  
  const rules3 = [
    '• Include: Status = "Closed" AND Issue Type = "Epic"',
    '• Sum: Story Point Completion for all closed epics',
    '• No allocation filtering - includes all allocation types'
  ];
  
  rules3.forEach(rule => {
    sheet.getRange(currentRow, 1).setValue(rule)
      .setFontFamily('Comfortaa')
      .setFontSize(9)
      .setFontColor('#666666');
    currentRow++;
  });
  currentRow++;
  
  // Filter epics
  const closedEpics = allEpics.filter(e => e.status.toLowerCase() === 'closed');
  
  // Write column headers
  const headerRow = currentRow;
  sheet.getRange(headerRow, 1, 1, 6).setValues([[
    'Key', 'Summary', 'Allocation', 'Story Points', 'Status', 'Included?'
  ]]).setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  currentRow++;
  
  // Write epic details
  allEpics.forEach(epic => {
    const included = epic.status.toLowerCase() === 'closed';
    
    sheet.getRange(currentRow, 1, 1, 6).setValues([[
      epic.key,
      epic.summary,
      epic.allocation,
      epic.storyPoints,
      epic.status,
      included ? 'YES' : 'NO'
    ]]).setFontFamily('Comfortaa').setFontSize(9);
    
    if (included) {
      sheet.getRange(currentRow, 6).setBackground('#d4edda').setFontWeight('bold');
    } else {
      sheet.getRange(currentRow, 6).setBackground('#f8d7da');
    }
    
    currentRow++;
  });
  
  // Summary totals
  currentRow++;
  const totalVelocity = closedEpics.reduce((sum, epic) => sum + epic.storyPoints, 0);
  
  sheet.getRange(currentRow, 1).setValue('TOTAL ART VELOCITY:')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.getRange(currentRow, 2).setValue(totalVelocity)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`Total Epics Included: ${closedEpics.length} of ${allEpics.length}`)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow += 3;
  
  return currentRow;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLE 4: EPIC STATUS VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

function writeTable4Validation(sheet, startRow, allEpics, piNumber, valueStream) {
  let currentRow = startRow;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`TABLE 4: EPIC STATUS`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  currentRow++;
  
  // Business rules
  sheet.getRange(currentRow, 1).setValue(`Business Rules:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold');
  currentRow++;
  
  const rules4 = [
    '• EXCLUDE: Epics with Allocation = "Quality" or "KLO"',
    '• Committed: PI Commitment = "Committed" OR "Committed After Plan"',
    '• Deferred: PI Commitment = "Deferred" OR "Canceled"',
    '• Deferred taken into Planning: TBD (currently 0)'
  ];
  
  rules4.forEach(rule => {
    sheet.getRange(currentRow, 1).setValue(rule)
      .setFontFamily('Comfortaa')
      .setFontSize(9)
      .setFontColor('#666666');
    currentRow++;
  });
  currentRow++;
  
  // Filter epics
  const nonQualityKLOEpics = allEpics.filter(e => 
    e.allocation !== 'Quality' && e.allocation !== 'KLO'
  );
  
  const committedEpics = nonQualityKLOEpics.filter(e => 
    e.piCommitment === 'Committed' || e.piCommitment === 'Committed After Plan'
  );
  
  const deferredEpics = nonQualityKLOEpics.filter(e => 
    e.piCommitment === 'Deferred' || e.piCommitment === 'Canceled'
  );
  
  // Write column headers
  const headerRow = currentRow;
  sheet.getRange(headerRow, 1, 1, 7).setValues([[
    'Key', 'Summary', 'PI Commitment', 'Allocation', 'Status', 'Category', 'Reason'
  ]]).setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  currentRow++;
  
  // Write epic details
  allEpics.forEach(epic => {
    let category = 'EXCLUDED (Quality/KLO)';
    let bgColor = '#f8d7da';
    let reason = '';
    
    if (epic.allocation === 'Quality' || epic.allocation === 'KLO') {
      category = 'EXCLUDED (Quality/KLO)';
      bgColor = '#f8d7da';
      reason = `Allocation = "${epic.allocation}" (excluded from Epic Status count per business rules)`;
    } else if (epic.piCommitment === 'Committed' || epic.piCommitment === 'Committed After Plan') {
      category = 'COMMITTED';
      bgColor = '#d4edda';
      reason = `PI Commitment = "${epic.piCommitment}"`;
    } else if (epic.piCommitment === 'Deferred' || epic.piCommitment === 'Canceled') {
      category = 'DEFERRED';
      bgColor = '#fff3cd';
      reason = `PI Commitment = "${epic.piCommitment}"`;
    } else {
      category = 'Other';
      bgColor = '#e2e3e5';
      reason = `PI Commitment = "${epic.piCommitment}" (not Committed or Deferred)`;
    }
    
    sheet.getRange(currentRow, 1, 1, 7).setValues([[
      epic.key,
      epic.summary,
      epic.piCommitment,
      epic.allocation,
      epic.status,
      category,
      reason
    ]]).setFontFamily('Comfortaa').setFontSize(9);
    
    sheet.getRange(currentRow, 6).setBackground(bgColor).setFontWeight('bold');
    
    currentRow++;
  });
  
  // Summary totals
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('EPIC STATUS COUNTS:')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Committed:', committedEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Deferred:', deferredEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Deferred taken into Planning:', 0]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`Total Non-Quality/KLO Epics: ${nonQualityKLOEpics.length}`)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow += 3;
  
  return currentRow;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLE 5: PROGRAM PREDICTABILITY VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

function writeTable5Validation(sheet, startRow, allEpics, piNumber, valueStream) {
  let currentRow = startRow;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`TABLE 5: PROGRAM PREDICTABILITY`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  currentRow++;
  
  // Business rules
  sheet.getRange(currentRow, 1).setValue(`Business Rules:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold');
  currentRow++;
  
  const rules5 = [
    '• FILTER: Status = "Closed" OR "Pending Acceptance"',
    '• EXCLUDE: Epics with Allocation = "Quality" or "KLO"',
    '• INCLUDE: Only epics where PI Objective Status = "Met"',
    '• Sum: Business Value (all non-Quality/KLO epics) and Actual Value (only Met epics)',
    '• Calculate: PI Score = (Actual Value / Business Value) × 100%'
  ];
  
  rules5.forEach(rule => {
    sheet.getRange(currentRow, 1).setValue(rule)
      .setFontFamily('Comfortaa')
      .setFontSize(9)
      .setFontColor('#666666');
    currentRow++;
  });
  currentRow++;
  
  // Filter epics - must match predictability.gs logic exactly
  const includedEpics = allEpics.filter(e => {
    const normalizedStatus = e.status.toLowerCase();
    const normalizedPIObjectiveStatus = e.piObjectiveStatus.toLowerCase();
    
    // Must be Closed or Pending Acceptance
    if (normalizedStatus !== 'closed' && normalizedStatus !== 'pending acceptance') return false;
    
    // Must not be Quality or KLO
    if (e.allocation === 'Quality' || e.allocation === 'KLO') return false;
    
    // Must have PI Objective Status = Met
    if (normalizedPIObjectiveStatus !== 'met') return false;
    
    return true;
  });
  
  // Write column headers - NOW INCLUDES PROGRAM INITIATIVE
  const headerRow = currentRow;
  sheet.getRange(headerRow, 1, 1, 10).setValues([[
    'Key', 'Program Initiative', 'Summary', 'Status', 'Allocation', 'PI Objective Status', 'Business Value', 'Actual Value', 'Included?', 'Reason'
  ]]).setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  currentRow++;
  
  // Write epic details
  allEpics.forEach(epic => {
    const normalizedStatus = epic.status.toLowerCase();
    const normalizedPIObjectiveStatus = epic.piObjectiveStatus.toLowerCase();
    const isStatusValid = normalizedStatus === 'closed' || normalizedStatus === 'pending acceptance';
    const isAllocationValid = epic.allocation !== 'Quality' && epic.allocation !== 'KLO';
    const isPIObjectiveStatusMet = normalizedPIObjectiveStatus === 'met';
    
    const included = isStatusValid && isAllocationValid && isPIObjectiveStatusMet;
    
    let reason = '';
    if (!isStatusValid) {
      reason = `Status = "${epic.status}" (must be Closed or Pending Acceptance)`;
    } else if (!isAllocationValid) {
      reason = `Allocation = "${epic.allocation}" (excluded per business rules)`;
    } else if (!isPIObjectiveStatusMet) {
      reason = `PI Objective Status = "${epic.piObjectiveStatus}" (only Met epics included for Actual Value)`;
    } else {
      reason = 'All criteria met';
    }
    
    sheet.getRange(currentRow, 1, 1, 10).setValues([[
      epic.key,
      epic.programInitiative || '(No Initiative)',
      epic.summary,
      epic.status,
      epic.allocation,
      epic.piObjectiveStatus,
      epic.businessValue,
      epic.actualValue,
      included ? 'YES' : 'NO',
      reason
    ]]).setFontFamily('Comfortaa').setFontSize(9);
    
    if (included) {
      sheet.getRange(currentRow, 9).setBackground('#d4edda').setFontWeight('bold');
    } else {
      sheet.getRange(currentRow, 9).setBackground('#f8d7da');
    }
    
    currentRow++;
  });
  
  // Summary totals
  currentRow++;
  const totalBV = includedEpics.reduce((sum, epic) => sum + epic.businessValue, 0);
  const totalAV = includedEpics.reduce((sum, epic) => sum + epic.actualValue, 0);
  const piScore = totalBV > 0 ? (totalAV / totalBV) * 100 : 0;
  
  sheet.getRange(currentRow, 1).setValue('PROGRAM PREDICTABILITY TOTALS:')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Business Value:', totalBV]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Actual Value:', totalAV]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['PI Score:', `${piScore.toFixed(1)}%`]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`Total Epics Included: ${includedEpics.length} of ${allEpics.length}`)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow += 3;
  
  return currentRow;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLE 6: OBJECTIVE STATUS VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

function writeTable6Validation(sheet, startRow, allEpics, piNumber, valueStream) {
  let currentRow = startRow;
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`TABLE 6: OBJECTIVE STATUS`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  currentRow++;
  
  // Business rules
  sheet.getRange(currentRow, 1).setValue(`Business Rules:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold');
  currentRow++;
  
  const rules6 = [
    '• Include: Status = "Closed" OR "Pending Acceptance"',
    '• Count ONLY epics with valid PI Objective Status (Met, Partial, Not Met)',
    '• ALL = Met + Partial + Not Met (epics with blank status excluded from count)',
    '• Calculate: Objective Score = (Met / ALL) × 100%'
  ];
  
  rules6.forEach(rule => {
    sheet.getRange(currentRow, 1).setValue(rule)
      .setFontFamily('Comfortaa')
      .setFontSize(9)
      .setFontColor('#666666');
    currentRow++;
  });
  currentRow++;
  
  // Filter epics to match predictability.gs logic exactly
  const closedEpics = allEpics.filter(e => {
    const normalizedStatus = e.status.toLowerCase();
    return normalizedStatus === 'closed' || normalizedStatus === 'pending acceptance';
  });
  
  const metEpics = closedEpics.filter(e => e.piObjectiveStatus.toLowerCase() === 'met');
  const partialEpics = closedEpics.filter(e => e.piObjectiveStatus.toLowerCase() === 'partial');
  const notMetEpics = closedEpics.filter(e => e.piObjectiveStatus.toLowerCase() === 'not met');
  
  // ALL = Sum of Met + Partial + Not Met (only epics with valid PI Objective Status)
  const allClosedEpicsCount = metEpics.length + partialEpics.length + notMetEpics.length;
  
  // Write column headers
  const headerRow = currentRow;
  sheet.getRange(headerRow, 1, 1, 7).setValues([[
    'Key', 'Summary', 'Status', 'PI Objective Status', 'Allocation', 'Category', 'Reason'
  ]]).setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold')
    .setBackground('#f0f0f0');
  currentRow++;
  
  // Write epic details
  allEpics.forEach(epic => {
    const normalizedStatus = epic.status.toLowerCase();
    const normalizedPIObjectiveStatus = epic.piObjectiveStatus.toLowerCase();
    const isStatusValid = normalizedStatus === 'closed' || normalizedStatus === 'pending acceptance';
    
    let category = 'Not Closed/Pending Acceptance';
    let bgColor = '#f8d7da';
    let reason = '';
    
    if (!isStatusValid) {
      category = 'Not Closed/Pending Acceptance';
      bgColor = '#f8d7da';
      reason = `Status = "${epic.status}" (must be Closed or Pending Acceptance)`;
    } else if (normalizedPIObjectiveStatus === 'met') {
      category = 'MET';
      bgColor = '#d4edda';
      reason = 'Closed/Pending + Met';
    } else if (normalizedPIObjectiveStatus === 'partial') {
      category = 'PARTIAL';
      bgColor = '#fff3cd';
      reason = 'Closed/Pending + Partial';
    } else if (normalizedPIObjectiveStatus === 'not met') {
      category = 'NOT MET';
      bgColor = '#f8d7da';
      reason = 'Closed/Pending + Not Met';
    } else {
      category = 'Closed (No Status)';
      bgColor = '#e2e3e5';
      reason = `PI Objective Status blank or invalid (excluded from ALL count)`;
    }
    
    sheet.getRange(currentRow, 1, 1, 7).setValues([[
      epic.key,
      epic.summary,
      epic.status,
      epic.piObjectiveStatus,
      epic.allocation,
      category,
      reason
    ]]).setFontFamily('Comfortaa').setFontSize(9);
    
    sheet.getRange(currentRow, 6).setBackground(bgColor).setFontWeight('bold');
    
    currentRow++;
  });
  
  // Summary totals
  currentRow++;
  const objectiveScore = allClosedEpicsCount > 0 ? (metEpics.length / allClosedEpicsCount) * 100 : 0;
  
  sheet.getRange(currentRow, 1).setValue('OBJECTIVE STATUS COUNTS:')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Met:', metEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Partial:', partialEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Not Met:', notMetEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['ALL (Met + Partial + Not Met):', allClosedEpicsCount]])
    .setFontFamily('Comfortaa').setFontSize(9).setFontWeight('bold');
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Objective Score:', `${objectiveScore.toFixed(1)}%`]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`Total Closed/Pending Epics: ${closedEpics.length} of ${allEpics.length}`)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  sheet.getRange(currentRow, 1).setValue(`(Epics with valid PI Objective Status: ${allClosedEpicsCount})`)
    .setFontFamily('Comfortaa')
    .setFontSize(9)
    .setFontColor('#666666')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow += 3;
  
  return currentRow;
}

// ═══════════════════════════════════════════════════════════════════════════
// TABLE 7: PI SCORE SUMMARY VALIDATION
// ═══════════════════════════════════════════════════════════════════════════

function writeTable7Validation(sheet, startRow, allEpics, piNumber, valueStream) {
  let currentRow = startRow;
  
  // Check if this is MMPM (case-insensitive)
  const isMMPM = valueStream.toUpperCase() === 'MMPM';
  
  // Section header
  sheet.getRange(currentRow, 1).setValue(`TABLE 7: PI SCORE SUMMARY`)
    .setFontFamily('Comfortaa')
    .setFontSize(12)
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  currentRow++;
  
  // Business rules
  sheet.getRange(currentRow, 1).setValue(`Business Rules:`)
    .setFontFamily('Comfortaa')
    .setFontSize(10)
    .setFontWeight('bold');
  currentRow++;
  
  const rules7 = [
    '• This is a SUMMARY row pulling data from other tables',
    '• Reads the PI Score value from Table 5 (Program Predictability)',
    '• PI Score = (Actual Value / Business Value) × 100%',
    '• BV: ALL epics (any status), excluding Quality/KLO',
    '• AV: Only Closed/Pending Acceptance epics with PI Objective Status = "Met"',
    '• See Table 5 validation above for calculation details'
  ];
  
  if (isMMPM) {
    rules7.splice(4, 0, '• MMPM: Excludes Scrum Team = "Appeal Engine" from both BV and AV');
  }
  
  rules7.forEach(rule => {
    sheet.getRange(currentRow, 1).setValue(rule)
      .setFontFamily('Comfortaa')
      .setFontSize(9)
      .setFontColor('#666666');
    currentRow++;
  });
  currentRow++;
  
  // Calculate for reference (same as Table 5 - corrected logic)
  // BV: ALL epics (any status), non-Quality/KLO
  const bvEpics = allEpics.filter(e => {
    if (e.allocation === 'Quality' || e.allocation === 'KLO' || e.allocation === 'KLO / Quality') return false;
    if (isMMPM && e.scrumTeam && e.scrumTeam.toLowerCase() === 'appeal engine') return false;
    return true;
  });
  
  // AV: Closed/Pending + Met
  const avEpics = bvEpics.filter(e => {
    const normalizedStatus = e.status.toLowerCase();
    if (normalizedStatus !== 'closed' && normalizedStatus !== 'pending acceptance') return false;
    return e.piObjectiveStatus.toLowerCase() === 'met';
  });
  
  const totalBV = bvEpics.reduce((sum, epic) => sum + epic.businessValue, 0);
  const totalAV = avEpics.reduce((sum, epic) => sum + epic.actualValue, 0);
  const piScore = totalBV > 0 ? (totalAV / totalBV) * 100 : 0;
  
  sheet.getRange(currentRow, 1).setValue('PI SCORE (from Table 5):')
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.getRange(currentRow, 2).setValue(`${piScore.toFixed(1)}%`)
    .setFontFamily('Comfortaa')
    .setFontWeight('bold')
    .setFontSize(10);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['BV Epics:', bvEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['AV Epics (Met):', avEpics.length]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Total BV:', totalBV]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 2).setValues([['Total AV:', totalAV]])
    .setFontFamily('Comfortaa').setFontSize(9);
  currentRow += 3;
  
  return currentRow;
}