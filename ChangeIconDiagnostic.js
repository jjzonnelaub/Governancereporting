/**
 * DIAGNOSTIC FUNCTION: Traces the change icon data flow
 * Run this to identify why change icons aren't displaying
 * 
 * Add this function to your project and run it from Script Editor
 */
function diagnoseChangeIconIssue() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Prompt for PI and Iteration
  const piResponse = ui.prompt('Enter PI Number (e.g., 15):');
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  const piNumber = parseInt(piResponse.getResponseText().trim());
  
  const iterResponse = ui.prompt('Enter Iteration Number (e.g., 2):');
  if (iterResponse.getSelectedButton() !== ui.Button.OK) return;
  const iterationNumber = parseInt(iterResponse.getResponseText().trim());
  
  console.log('\n' + '='.repeat(80));
  console.log('CHANGE ICON DIAGNOSTIC');
  console.log('='.repeat(80));
  console.log(`PI: ${piNumber}, Iteration: ${iterationNumber}`);
  console.log('='.repeat(80) + '\n');
  
  const diagnosticResults = [];
  
  // CHECK 1: Iteration sheet exists
  console.log('CHECK 1: Iteration Sheet');
  const iterSheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
  const iterSheet = ss.getSheetByName(iterSheetName);
  if (!iterSheet) {
    diagnosticResults.push(`❌ CHECK 1 FAILED: Sheet "${iterSheetName}" not found`);
    console.log(`  ❌ FAILED: Sheet "${iterSheetName}" not found`);
  } else {
    const iterRowCount = iterSheet.getLastRow() - 4;
    diagnosticResults.push(`✅ CHECK 1 PASSED: Found "${iterSheetName}" with ${iterRowCount} data rows`);
    console.log(`  ✅ PASSED: Found "${iterSheetName}" with ${iterRowCount} data rows`);
  }
  
  // CHECK 2: Previous iteration sheet exists (for iteration > 1)
  console.log('\nCHECK 2: Previous Iteration Sheet');
  if (iterationNumber > 1) {
    const prevSheetName = `PI ${piNumber} - Iteration ${iterationNumber - 1}`;
    const prevSheet = ss.getSheetByName(prevSheetName);
    if (!prevSheet) {
      diagnosticResults.push(`❌ CHECK 2 FAILED: Previous iteration sheet "${prevSheetName}" not found`);
      console.log(`  ❌ FAILED: Previous iteration sheet "${prevSheetName}" not found`);
    } else {
      const prevRowCount = prevSheet.getLastRow() - 4;
      diagnosticResults.push(`✅ CHECK 2 PASSED: Found "${prevSheetName}" with ${prevRowCount} data rows`);
      console.log(`  ✅ PASSED: Found "${prevSheetName}" with ${prevRowCount} data rows`);
    }
  } else {
    diagnosticResults.push(`⏭️ CHECK 2 SKIPPED: Iteration 1 is baseline (no previous iteration needed)`);
    console.log(`  ⏭️ SKIPPED: Iteration 1 is baseline`);
  }
  
  // CHECK 3: Changelog sheet exists
  console.log('\nCHECK 3: Changelog Sheet');
  if (iterationNumber > 1) {
    const changelogName = `PI ${piNumber} Changelog`;
    const changelogSheet = ss.getSheetByName(changelogName);
    if (!changelogSheet) {
      diagnosticResults.push(`❌ CHECK 3 FAILED: Changelog sheet "${changelogName}" not found. Run "Analyze Changes for Iteration" first!`);
      console.log(`  ❌ FAILED: Changelog sheet "${changelogName}" not found`);
    } else {
      const headers = changelogSheet.getRange(5, 1, 1, changelogSheet.getLastColumn()).getValues()[0];
      const iterColumn = `Iteration ${iterationNumber} - NO CHANGES`;
      
      if (!headers.includes(iterColumn)) {
        diagnosticResults.push(`❌ CHECK 3 FAILED: Changelog hasn't been analyzed for Iteration ${iterationNumber}. Column "${iterColumn}" missing.`);
        console.log(`  ❌ FAILED: Changelog column "${iterColumn}" not found`);
        console.log(`  Available columns: ${headers.filter(h => h.includes('Iteration')).join(', ')}`);
      } else {
        const changelogRows = changelogSheet.getLastRow() - 5;
        diagnosticResults.push(`✅ CHECK 3 PASSED: Changelog analyzed for Iteration ${iterationNumber} (${changelogRows} tracked issues)`);
        console.log(`  ✅ PASSED: Changelog has ${changelogRows} tracked issues with Iteration ${iterationNumber} analysis`);
      }
    }
  } else {
    diagnosticResults.push(`⏭️ CHECK 3 SKIPPED: Iteration 1 is baseline (changelog not required)`);
    console.log(`  ⏭️ SKIPPED: Iteration 1 is baseline`);
  }
  
  // CHECK 4: Sample epic change detection (only if iteration > 1)
  console.log('\nCHECK 4: Sample Change Detection');
  if (iterationNumber > 1 && iterSheet) {
    const prevSheetName = `PI ${piNumber} - Iteration ${iterationNumber - 1}`;
    const prevSheet = ss.getSheetByName(prevSheetName);
    
    if (prevSheet) {
      // Load headers
      const currHeaders = iterSheet.getRange(4, 1, 1, iterSheet.getLastColumn()).getValues()[0];
      const prevHeaders = prevSheet.getRange(4, 1, 1, prevSheet.getLastColumn()).getValues()[0];
      
      // Find key columns
      const currKeyIdx = currHeaders.indexOf('Key');
      const prevKeyIdx = prevHeaders.indexOf('Key');
      const currStatusIdx = currHeaders.indexOf('Status');
      const prevStatusIdx = prevHeaders.indexOf('Status');
      const currIterIdx = currHeaders.indexOf('End Iteration Name') >= 0 ? currHeaders.indexOf('End Iteration Name') : currHeaders.indexOf('PI Target Iteration');
      const prevIterIdx = prevHeaders.indexOf('End Iteration Name') >= 0 ? prevHeaders.indexOf('End Iteration Name') : prevHeaders.indexOf('PI Target Iteration');
      const currRagIdx = currHeaders.indexOf('RAG');
      const prevRagIdx = prevHeaders.indexOf('RAG');
      
      // Load first 10 epics from current iteration
      const currLastRow = iterSheet.getLastRow();
      const currData = iterSheet.getRange(5, 1, Math.min(50, currLastRow - 4), currHeaders.length).getValues();
      
      // Build previous iteration lookup
      const prevLastRow = prevSheet.getLastRow();
      const prevData = prevSheet.getRange(5, 1, prevLastRow - 4, prevHeaders.length).getValues();
      const prevLookup = {};
      prevData.forEach(row => {
        const key = row[prevKeyIdx];
        if (key) prevLookup[key] = row;
      });
      
      // Analyze changes
      let newCount = 0;
      let modifiedCount = 0;
      let closedCount = 0;
      let unchangedCount = 0;
      const sampleChanges = [];
      
      currData.forEach(row => {
        const key = row[currKeyIdx];
        const issueType = row[currHeaders.indexOf('Issue Type')];
        if (!key || issueType !== 'Epic') return;
        
        const currStatus = (row[currStatusIdx] || '').toString().toUpperCase();
        const currIter = (row[currIterIdx] || '').toString();
        const currRag = (row[currRagIdx] || '').toString();
        
        const prev = prevLookup[key];
        
        if (!prev) {
          newCount++;
          if (sampleChanges.length < 5) {
            sampleChanges.push(`  📗 NEW: ${key} - not in previous iteration`);
          }
        } else {
          const prevStatus = (prev[prevStatusIdx] || '').toString().toUpperCase();
          const prevIter = (prev[prevIterIdx] || '').toString();
          const prevRag = (prev[prevRagIdx] || '').toString();
          
          const isNowClosed = currStatus === 'DONE' || currStatus === 'CLOSED';
          const wasNotClosed = prevStatus !== 'DONE' && prevStatus !== 'CLOSED';
          
          if (isNowClosed && wasNotClosed) {
            closedCount++;
            if (sampleChanges.length < 5) {
              sampleChanges.push(`  ✅ CLOSED: ${key} - Status: ${prevStatus} → ${currStatus}`);
            }
          } else if (currIter !== prevIter || currRag !== prevRag) {
            modifiedCount++;
            if (sampleChanges.length < 5) {
              const changes = [];
              if (currIter !== prevIter) changes.push(`Iter: ${prevIter || '(blank)'} → ${currIter || '(blank)'}`);
              if (currRag !== prevRag) changes.push(`RAG: ${prevRag || '(blank)'} → ${currRag || '(blank)'}`);
              sampleChanges.push(`  📙 MODIFIED: ${key} - ${changes.join(', ')}`);
            }
          } else {
            unchangedCount++;
          }
        }
      });
      
      console.log(`  Change Detection Summary:`);
      console.log(`    - NEW (added this iteration): ${newCount}`);
      console.log(`    - MODIFIED (changed fields): ${modifiedCount}`);
      console.log(`    - CLOSED (status → Done/Closed): ${closedCount}`);
      console.log(`    - UNCHANGED: ${unchangedCount}`);
      
      if (sampleChanges.length > 0) {
        console.log(`\n  Sample Changes Detected:`);
        sampleChanges.forEach(c => console.log(c));
      }
      
      const totalWithChanges = newCount + modifiedCount + closedCount;
      if (totalWithChanges > 0) {
        diagnosticResults.push(`✅ CHECK 4 PASSED: Detected ${totalWithChanges} epics with changes (${newCount} NEW, ${modifiedCount} MODIFIED, ${closedCount} CLOSED)`);
      } else {
        diagnosticResults.push(`⚠️ CHECK 4 WARNING: No changes detected between iterations. Verify data is different.`);
      }
    } else {
      diagnosticResults.push(`⏭️ CHECK 4 SKIPPED: Previous iteration sheet not found`);
    }
  } else if (iterationNumber === 1) {
    diagnosticResults.push(`⏭️ CHECK 4 SKIPPED: Iteration 1 is baseline - all items will show without badges`);
    console.log(`  ⏭️ SKIPPED: Iteration 1 is baseline`);
  }
  
  // CHECK 5: Configuration sheet
  console.log('\nCHECK 5: Configuration Sheet');
  const configSheet = ss.getSheetByName('Configuration');
  if (!configSheet) {
    diagnosticResults.push(`⚠️ CHECK 5 WARNING: Configuration sheet not found. Using default settings.`);
    console.log(`  ⚠️ WARNING: Configuration sheet not found`);
  } else {
    diagnosticResults.push(`✅ CHECK 5 PASSED: Configuration sheet found`);
    console.log(`  ✅ PASSED: Configuration sheet found`);
  }
  
  // SUMMARY
  console.log('\n' + '='.repeat(80));
  console.log('DIAGNOSTIC SUMMARY');
  console.log('='.repeat(80));
  diagnosticResults.forEach(r => console.log(r));
  
  // Check for critical failures
  const failures = diagnosticResults.filter(r => r.includes('❌'));
  
  let alertMessage;
  if (failures.length > 0) {
    alertMessage = 'ISSUES FOUND:\n\n' + failures.map(f => f.replace('❌ ', '• ')).join('\n') + 
      '\n\nRecommendation: Fix the above issues before generating slides.';
  } else {
    alertMessage = 'All checks passed!\n\n' + diagnosticResults.join('\n') +
      '\n\nIf badges still aren\'t showing, check the Execution Log for detailed output.';
  }
  
  ui.alert('Change Icon Diagnostic Results', alertMessage, ui.ButtonSet.OK);
  
  return diagnosticResults;
}

/**
 * Additional diagnostic: Dump epic change flags for a specific epic
 */
function diagnoseSpecificEpic() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const response = ui.prompt('Enter Epic Key (e.g., PROJ-123):');
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const epicKey = response.getResponseText().trim().toUpperCase();
  
  // Find the epic in all PI sheets
  const sheets = ss.getSheets();
  const results = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!name.match(/^PI \d+ - Iteration \d+$/)) return;
    
    const headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0];
    const keyIdx = headers.indexOf('Key');
    if (keyIdx < 0) return;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 5) return;
    
    const keys = sheet.getRange(5, keyIdx + 1, lastRow - 4, 1).getValues();
    const rowIdx = keys.findIndex(r => r[0] === epicKey);
    
    if (rowIdx >= 0) {
      const rowData = sheet.getRange(5 + rowIdx, 1, 1, headers.length).getValues()[0];
      const epicData = {};
      headers.forEach((h, i) => { epicData[h] = rowData[i]; });
      results.push({ sheet: name, data: epicData });
    }
  });
  
  if (results.length === 0) {
    ui.alert('Epic Not Found', `Could not find epic "${epicKey}" in any iteration sheet.`, ui.ButtonSet.OK);
    return;
  }
  
  console.log(`\n${'='.repeat(80)}`);
  console.log(`EPIC HISTORY: ${epicKey}`);
  console.log('='.repeat(80));
  
  results.forEach(r => {
    console.log(`\n📋 ${r.sheet}:`);
    console.log(`  Status: ${r.data['Status']}`);
    console.log(`  PI Commitment: ${r.data['PI Commitment']}`);
    console.log(`  RAG: ${r.data['RAG']}`);
    console.log(`  RAG Note: ${(r.data['RAG Note'] || '').substring(0, 50)}...`);
    console.log(`  End Iteration: ${r.data['End Iteration Name'] || r.data['PI Target Iteration']}`);
    console.log(`  Fix Versions: ${r.data['Fix Versions']}`);
  });
  
  ui.alert('Epic History', 
    `Found "${epicKey}" in ${results.length} iteration sheets.\n\nCheck Execution Log for detailed history.`, 
    ui.ButtonSet.OK);
}