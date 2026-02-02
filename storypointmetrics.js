function diagnoseStoryPointFields() {
  console.log('\n' + '='.repeat(70));
  console.log('STORY POINT FIELD DIAGNOSTIC');
  console.log('='.repeat(70) + '\n');
  
  // Get a sample of Story and Bug tickets
  const jql = 'type in (Story, Bug) ORDER BY updated DESC';
  console.log(`Fetching sample tickets with JQL: ${jql}`);
  
  const sampleTickets = searchJiraIssuesLimited(jql, 10);
  
  if (sampleTickets.length === 0) {
    return '❌ No Story or Bug tickets found in JIRA.\n\nPlease check your JIRA connection.';
  }
  
  console.log(`Found ${sampleTickets.length} sample tickets\n`);
  
  let report = `📊 STORY POINT FIELD DIAGNOSTIC\n`;
  report += `${'='.repeat(50)}\n\n`;
  report += `Analyzed ${sampleTickets.length} recent Story/Bug tickets:\n\n`;
  
  let field1Count = 0;
  let field2Count = 0;
  let bothFieldsCount = 0;
  let noFieldsCount = 0;
  
  const field1Name = 'storyPoints (customfield_10037)';
  const field2Name = 'storyPointEstimate (customfield_10016)';
  
  sampleTickets.forEach(ticket => {
    const hasField1 = ticket.storyPoints && ticket.storyPoints > 0;
    const hasField2 = ticket.storyPointEstimate && ticket.storyPointEstimate > 0;
    
    if (hasField1 && hasField2) {
      bothFieldsCount++;
      console.log(`${ticket.key}: Both fields (${ticket.storyPoints} / ${ticket.storyPointEstimate})`);
    } else if (hasField1) {
      field1Count++;
      console.log(`${ticket.key}: Field 1 only (${ticket.storyPoints})`);
    } else if (hasField2) {
      field2Count++;
      console.log(`${ticket.key}: Field 2 only (${ticket.storyPointEstimate})`);
    } else {
      noFieldsCount++;
      console.log(`${ticket.key}: No story points`);
    }
  });
  
  report += `Results:\n`;
  report += `• ${field1Name}: ${field1Count} tickets\n`;
  report += `• ${field2Name}: ${field2Count} tickets\n`;
  report += `• Both fields populated: ${bothFieldsCount} tickets\n`;
  report += `• No story points: ${noFieldsCount} tickets\n\n`;
  
  // Recommendation
  if (field1Count > field2Count) {
    report += `✅ RECOMMENDATION:\n`;
    report += `Your primary field is: ${field1Name}\n`;
    report += `The refresh function will check this field first.`;
  } else if (field2Count > field1Count) {
    report += `✅ RECOMMENDATION:\n`;
    report += `Your primary field is: ${field2Name}\n`;
    report += `The refresh function will check this field second (fallback).`;
  } else {
    report += `⚠️ RECOMMENDATION:\n`;
    report += `Both fields are equally used.\n`;
    report += `The refresh function will check both fields (storyPoints first, then storyPointEstimate).`;
  }
  
  console.log('\n' + report);
  return report;
}

// ═══════════════════════════════════════════════════════════════════════════
// MAIN REFRESH FUNCTION
// ═══════════════════════════════════════════════════════════════════════════
function refreshChildTicketMetrics(piNumber, sheetBaseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ===== BUILD CORRECT SHEET NAME =====
  let fullSheetName;
  if (sheetBaseName === 'Iteration 0 (Pre-Planning)') {
    fullSheetName = `PI ${piNumber} - Iteration 0 - Pre-Planning`;
  } else if (sheetBaseName === 'Iteration 6 (IP)') {
    fullSheetName = `PI ${piNumber} - Iteration 6`;
  } else {
    // Standard format: "PI 11 - Iteration 1", "PI 11 - Iteration 2", etc.
    const iterNum = sheetBaseName.replace('Iteration ', '');
    fullSheetName = `PI ${piNumber} - Iteration ${iterNum}`;
  }
  
  const sheet = ss.getSheetByName(fullSheetName);
  
  if (!sheet) {
    throw new Error(`Sheet "${fullSheetName}" not found. Please check PI number and sheet selection.\n\nLooking for sheet name: "${fullSheetName}"`);
  }
  
  console.log('\n' + '='.repeat(70));
  console.log('REFRESHING CHILD TICKET METRICS');
  console.log(`Sheet: ${fullSheetName}`);
  console.log('='.repeat(70) + '\n');
  
  // Get all data from the sheet
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) {
    throw new Error('No data found in sheet. Please run the main governance report first.');
  }
  
  const dataRange = sheet.getRange(5, 1, lastRow - 4, 40);
  const allData = dataRange.getValues();
  
  // Column indices (0-based)
  const COL_KEY = 0;           // A: Key
  const COL_ISSUE_TYPE = 2;    // C: Issue Type
  const COL_STORY_POINTS = 19; // T: Investment (Story Points)
  const COL_CLOSED_POINTS = 20; // U: Story Point Completion
  const COL_TOTAL_CHILDREN = 21; // V: Total Child Issues
  const COL_CLOSED_CHILDREN = 22; // W: Closed Child Issues
  const COL_PERCENT_COMPLETE = 23; // X: % Complete
  
  // Extract Epic keys
  const epicKeys = [];
  const epicRowMap = {}; // Map epic key to row number
  
  allData.forEach((row, index) => {
    const key = row[COL_KEY];
    const issueType = row[COL_ISSUE_TYPE];
    
    if (issueType === 'Epic' && key) {
      epicKeys.push(key);
      epicRowMap[key] = index + 5; // +5 to account for header rows
    }
  });
  
  if (epicKeys.length === 0) {
    throw new Error('No Epic tickets found in the sheet.');
  }
  
  console.log(`Found ${epicKeys.length} Epic tickets to process`);
  ss.toast(`Processing ${epicKeys.length} epics...`, '🔍 Step 1/2', 30);
  
  // Fetch child tickets using the JQL query
  const childMetricsMap = fetchChildMetricsForEpics(epicKeys);
  
  console.log('\n' + '='.repeat(70));
  console.log('UPDATING SHEET WITH METRICS');
  console.log('='.repeat(70) + '\n');
  
  ss.toast('Writing results to sheet...', '✍️ Step 2/2', 30);
  
  // Update the sheet
  let updatedCount = 0;
  
  epicKeys.forEach(epicKey => {
    const metrics = childMetricsMap[epicKey];
    const rowNumber = epicRowMap[epicKey];
    
    if (metrics) {
      // Update columns T, U, V, W
      sheet.getRange(rowNumber, COL_STORY_POINTS + 1).setValue(metrics.totalStoryPoints);
      sheet.getRange(rowNumber, COL_CLOSED_POINTS + 1).setValue(metrics.closedStoryPoints);
      sheet.getRange(rowNumber, COL_TOTAL_CHILDREN + 1).setValue(metrics.totalChildren);
      sheet.getRange(rowNumber, COL_CLOSED_CHILDREN + 1).setValue(metrics.closedChildren);
      
      // Update % Complete formula in column X
      if (metrics.totalStoryPoints > 0) {
        const percentFormula = `=ROUND((U${rowNumber}/T${rowNumber})*100, 0)`;
        sheet.getRange(rowNumber, COL_PERCENT_COMPLETE + 1).setFormula(percentFormula);
      } else {
        sheet.getRange(rowNumber, COL_PERCENT_COMPLETE + 1).setValue(0);
      }
      
      updatedCount++;
      
      console.log(`${epicKey}: ${metrics.totalStoryPoints} pts (${metrics.closedStoryPoints} closed), ` +
                  `${metrics.totalChildren} items (${metrics.closedChildren} closed)`);
    }
  });
  
  console.log('\n' + '='.repeat(70));
  console.log(`✅ COMPLETED: Updated ${updatedCount} epics`);
  console.log('='.repeat(70) + '\n');
  
  ss.toast('✅ Metrics refreshed successfully!', 'Complete', 5);
  
  return `Successfully updated ${updatedCount} epics:\n\n` +
         `• Story Points (Column T)\n` +
         `• Closed Story Points (Column U)\n` +
         `• Total Children (Column V)\n` +
         `• Closed Children (Column W)\n` +
         `• % Complete (Column X)\n\n` +
         `Sheet: ${fullSheetName}`;
}

// ═══════════════════════════════════════════════════════════════════════════
// CHILD TICKET FETCHING - Uses ascending/descending strategy to overcome 100 limit
// ═══════════════════════════════════════════════════════════════════════════
function fetchChildMetricsForEpics(epicKeys) {
  const BATCH_SIZE = 4;
  const metricsMap = {};
  
  // Initialize all epics with zero values
  epicKeys.forEach(key => {
    metricsMap[key] = {
      totalStoryPoints: 0,
      closedStoryPoints: 0,
      totalChildren: 0,
      closedChildren: 0
    };
  });
  
  console.log('\n' + '='.repeat(70));
  console.log('FETCHING CHILD TICKETS (Stories & Bugs)');
  console.log('='.repeat(70));
  console.log('JQL Pattern: parent in (Epic Keys) AND type in (Story, Bug)');
  console.log('Strategy: Ascending + Descending when hitting 100 limit');
  console.log('');
  console.log('Story Point Field Configuration:');
  console.log(`  Primary: storyPoints = ${FIELD_MAPPINGS.storyPoints}`);
  console.log(`  Fallback: storyPointEstimate = ${FIELD_MAPPINGS.storyPointEstimate}`);
  console.log('');
  
  const totalBatches = Math.ceil(epicKeys.length / BATCH_SIZE);
  let totalChildrenFetched = 0;
  let field1Count = 0;
  let field2Count = 0;
  let noFieldCount = 0;
  
  for (let i = 0; i < epicKeys.length; i += BATCH_SIZE) {
    const batch = epicKeys.slice(i, Math.min(i + BATCH_SIZE, epicKeys.length));
    const batchNum = Math.floor(i / BATCH_SIZE) + 1;
    
    console.log(`${'─'.repeat(70)}`);
    console.log(`BATCH ${batchNum}/${totalBatches} - Processing ${batch.length} epics`);
    console.log(`${'─'.repeat(70)}`);
    
    // Build JQL with parent keys
    const epicKeysList = batch.map(key => `"${key}"`).join(',');
    
    try {
      // Fetch with ascending/descending strategy
      const children = fetchChildrenWithAscDesc(epicKeysList);
      
      console.log(`✓ Found ${children.length} child tickets`);
      totalChildrenFetched += children.length;
      console.log('');
      
      if (children.length > 0) {
        console.log('Child Ticket Details:');
        console.log('');
      }
      
      // Process each child ticket
      children.forEach(child => {
        const parentKey = child.parentKey;
        
        if (parentKey && metricsMap[parentKey]) {
          let storyPoints = 0;
          let fieldSource = 'none';
          
          // Try primary field first (customfield_10037)
          if (child.storyPoints && child.storyPoints > 0) {
            storyPoints = child.storyPoints;
            fieldSource = FIELD_MAPPINGS.storyPoints;
            field1Count++;
          }
          // Try fallback field (customfield_10016)
          else if (child.storyPointEstimate && child.storyPointEstimate > 0) {
            storyPoints = child.storyPointEstimate;
            fieldSource = FIELD_MAPPINGS.storyPointEstimate;
            field2Count++;
          }
          else {
            noFieldCount++;
          }
          
          // Add to totals
          metricsMap[parentKey].totalStoryPoints += storyPoints;
          metricsMap[parentKey].totalChildren++;
          
          // ✅ FIXED: Check if closed - Pending Acceptance is NOT considered closed
          const status = (child.status || '').toLowerCase();
          const isClosed = status === 'closed' || 
                          status === 'done' || 
                          status === 'resolved' || 
                          status === 'completed';
          
          if (isClosed) {
            metricsMap[parentKey].closedStoryPoints += storyPoints;
            metricsMap[parentKey].closedChildren++;
          }
          
          // Detailed logging
          const closedMarker = isClosed ? '✓' : '○';
          const pointsDisplay = storyPoints > 0 ? `${storyPoints} pts` : 'NO PTS';
          console.log(`  ${closedMarker} ${child.key} (${child.issueType}): ${pointsDisplay} from [${fieldSource}], Status=${child.status}, Parent=${parentKey}`);
        }
      });
      
      console.log('');
      
    } catch (error) {
      console.error(`❌ Error in batch ${batchNum}: ${error.message}`);
    }
    
    // Rate limiting
    if (i + BATCH_SIZE < epicKeys.length) {
      Utilities.sleep(300);
    }
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SUMMARY REPORT
  // ═══════════════════════════════════════════════════════════════════════════
  console.log('='.repeat(70));
  console.log('FIELD USAGE SUMMARY');
  console.log('='.repeat(70));
  console.log(`Total child tickets fetched: ${totalChildrenFetched}`);
  console.log('');
  console.log('Story Point Field Distribution:');
  console.log(`  • ${FIELD_MAPPINGS.storyPoints}: ${field1Count} tickets (${((field1Count/totalChildrenFetched)*100).toFixed(1)}%)`);
  console.log(`  • ${FIELD_MAPPINGS.storyPointEstimate}: ${field2Count} tickets (${((field2Count/totalChildrenFetched)*100).toFixed(1)}%)`);
  console.log(`  • No story points: ${noFieldCount} tickets (${((noFieldCount/totalChildrenFetched)*100).toFixed(1)}%)`);
  console.log('');
  
  // Calculate totals
  let totalPoints = 0;
  let totalClosedPoints = 0;
  let totalItems = 0;
  let totalClosedItems = 0;
  
  console.log('='.repeat(70));
  console.log('EPIC METRICS SUMMARY');
  console.log('='.repeat(70));
  
  Object.entries(metricsMap).forEach(([key, metrics]) => {
    totalPoints += metrics.totalStoryPoints;
    totalClosedPoints += metrics.closedStoryPoints;
    totalItems += metrics.totalChildren;
    totalClosedItems += metrics.closedChildren;
    
    if (metrics.totalChildren > 0) {
      const completion = metrics.totalStoryPoints > 0 
        ? Math.round((metrics.closedStoryPoints / metrics.totalStoryPoints) * 100)
        : 0;
      console.log(`${key}: ${metrics.totalStoryPoints} pts (${completion}% complete), ${metrics.totalChildren} items`);
    }
  });
  
  console.log('');
  console.log('Overall Totals:');
  console.log(`  Total Story Points: ${totalPoints}`);
  console.log(`  Closed Story Points: ${totalClosedPoints}`);
  console.log(`  Total Children: ${totalItems}`);
  console.log(`  Closed Children: ${totalClosedItems}`);
  console.log(`  Completion Rate: ${totalPoints > 0 ? Math.round((totalClosedPoints/totalPoints)*100) : 0}%`);
  console.log('='.repeat(70) + '\n');
  
  return metricsMap;
}
function refreshChildMetrics(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  
  console.log('');
  console.log('╔══════════════════════════════════════════════════════════════╗');
  console.log('║           CHILD METRICS REFRESH                              ║');
  console.log('╠══════════════════════════════════════════════════════════════╣');
  console.log(`║  Sheet: ${sheetName}  |  ${new Date().toLocaleString()}`);
  console.log('╚══════════════════════════════════════════════════════════════╝');
  console.log('');
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
    
    // Get epic keys from the sheet
    const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
    const keyCol = headers.indexOf('Key') + 1;
    const issueTypeCol = headers.indexOf('Issue Type') + 1;
    
    if (keyCol === 0) {
      throw new Error('Key column not found');
    }
    
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(5, 1, lastRow - 4, sheet.getLastColumn()).getValues();
    
    // Get all epic keys
    const epicKeys = [];
    data.forEach((row, idx) => {
      if (row[issueTypeCol - 1] === 'Epic' && row[keyCol - 1]) {
        epicKeys.push(row[keyCol - 1]);
      }
    });
    
    console.log(`Found ${epicKeys.length} epics to process`);
    
    ss.toast(`Processing ${epicKeys.length} epics...`, '🔍 Step 1/2', 60);
    
    // Fetch child metrics for all epics
    const metricsMap = fetchChildMetricsForEpics(epicKeys);
    
    ss.toast('Writing results to sheet...', '✍️ Step 2/2', 30);
    
    // Find or create metric columns
    let totalChildCol = headers.indexOf('Total Child Issues') + 1;
    let closedChildCol = headers.indexOf('Closed Child Issues') + 1;
    let percentCompleteCol = headers.indexOf('% Complete') + 1;
    
    // Create columns if they don't exist
    const lastCol = sheet.getLastColumn();
    if (totalChildCol === 0) {
      totalChildCol = lastCol + 1;
      sheet.getRange(4, totalChildCol).setValue('Total Child Issues').setFontWeight('bold');
    }
    if (closedChildCol === 0) {
      closedChildCol = lastCol + 2;
      sheet.getRange(4, closedChildCol).setValue('Closed Child Issues').setFontWeight('bold');
    }
    if (percentCompleteCol === 0) {
      percentCompleteCol = lastCol + 3;
      sheet.getRange(4, percentCompleteCol).setValue('% Complete').setFontWeight('bold');
    }
    
    // Write metrics to sheet
    const outputData = [];
    data.forEach((row) => {
      const key = row[keyCol - 1];
      const metrics = metricsMap[key] || { total: 0, closed: 0 };
      const percent = metrics.total > 0 ? Math.round((metrics.closed / metrics.total) * 100) + '%' : '0%';
      outputData.push([metrics.total, metrics.closed, percent]);
    });
    
    sheet.getRange(5, totalChildCol, outputData.length, 3).setValues(outputData);
    
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    
    console.log('');
    console.log('╔══════════════════════════════════════════════════════════════╗');
    console.log('║                    ✅ METRICS REFRESHED                      ║');
    console.log('╠══════════════════════════════════════════════════════════════╣');
    console.log(`║  Epics: ${epicKeys.length}  |  Duration: ${duration}s`);
    console.log('╚══════════════════════════════════════════════════════════════╝');
    
    // Dismiss working toast
    ss.toast(`✅ Metrics refreshed in ${duration}s`, '✅ Complete', 5);
    
  } catch (error) {
    console.error('Error refreshing metrics:', error);
    
    // Dismiss working toast
    ss.toast('An error occurred', '❌ Error', 3);
    
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}
// ═══════════════════════════════════════════════════════════════════════════
// HELPER FUNCTION: Fetch children with ascending/descending strategy
// ═══════════════════════════════════════════════════════════════════════════
function fetchChildrenWithAscDesc(epicKeysList) {
  const jqlBase = `parent in (${epicKeysList}) AND type in (Story, Bug)`;
  
  // Try ascending first (ORDER BY key ASC)
  const jqlAsc = `${jqlBase} ORDER BY key ASC`;
  console.log(`JQL: ${jqlAsc}`);
  console.log('');
  
  const childrenAsc = searchJiraIssuesLimited(jqlAsc, 500);
  
  // If we got less than 100, we have everything
  if (childrenAsc.length < 100) {
    console.log(`✓ Got ${childrenAsc.length} results (< 100, no pagination needed)`);
    return childrenAsc;
  }
  
  // We hit the 100 limit - try descending
  console.warn(`⚠️  HIT 100 LIMIT on ascending (${childrenAsc.length} results) - trying descending...`);
  
  const jqlDesc = `${jqlBase} ORDER BY key DESC`;
  console.log(`JQL (DESC): ${jqlDesc}`);
  console.log('');
  
  const childrenDesc = searchJiraIssuesLimited(jqlDesc, 500);
  
  console.log(`Ascending results: ${childrenAsc.length}`);
  console.log(`Descending results: ${childrenDesc.length}`);
  
  // Check if descending also hit the limit
  if (childrenDesc.length >= 100) {
    console.warn(`⚠️  WARNING: Both ascending and descending hit 100 limit!`);
    console.warn(`⚠️  This batch may have more than 200 children - some data will be missing.`);
    console.warn(`⚠️  Consider reducing BATCH_SIZE from 6 to 4 or 3 epics.`);
  }
  
  // Merge and deduplicate
  const seenKeys = new Set();
  const merged = [];
  
  childrenAsc.forEach(child => {
    if (!seenKeys.has(child.key)) {
      seenKeys.add(child.key);
      merged.push(child);
    }
  });
  
  childrenDesc.forEach(child => {
    if (!seenKeys.has(child.key)) {
      seenKeys.add(child.key);
      merged.push(child);
    }
  });
  
  const duplicateCount = (childrenAsc.length + childrenDesc.length) - merged.length;
  console.log(`✓ Merged and deduplicated: ${merged.length} unique children`);
  console.log(`  (Removed ${duplicateCount} duplicates from ${childrenAsc.length + childrenDesc.length} total results)`);
  console.log('');
  
  return merged;
}