//PG-Slide Processing 
// ═══════════════════════════════════════════════════════════════════════════════
// REFACTORED: Now reads pre-computed badge flags from changelog instead of
// recomputing them. Badge computation happens in diffunction.gs during
// "Analyze Changes for Iteration" step.
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Main entry point for slide data processing
 * Uses pre-computed badge flags from changelog
 * 
 * @param {number} piNumber - PI number
 * @param {number} iterationNumber - Current iteration
 * @param {Object} slideConfig - Configuration for slide display
 * @param {boolean} showAllEpics - If true, show all epics (not just those with changes)
 * @param {boolean} includeAtRisk - If true, include at-risk epics with no changes
 * @param {boolean} includePreviousClosed - If true, include previously closed epics
 * @returns {Object} Processed data structured for slide generation
 */
function processSlideData(piNumber, iterationNumber, slideConfig = null, showAllEpics = false, includeAtRisk = true) {
  console.log(`\n========== PROCESSING SLIDE DATA ==========`);
  console.log(`PI: ${piNumber}, Iteration: ${iterationNumber}`);
  console.log(`Show All: ${showAllEpics}, Include At-Risk: ${includeAtRisk}`);
  
  if (slideConfig) {
    console.log('Using Configuration tab settings for filtering and ordering');
  } else {
    console.log('Using default settings (alphabetical ordering)');
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 1. Verify preconditions
    verifyPreconditions(ss, piNumber, iterationNumber);
    
    // 2. Load iteration sheet data
    const iterationData = loadIterationSheetData(ss, piNumber, iterationNumber);
    console.log(`  Loaded ${iterationData.length} items from iteration sheet`);
    
    // 3. Load pre-computed badge flags from changelog (for iterations 2+)
    let badgeFlags = new Map();
    if (iterationNumber > 1) {
      badgeFlags = loadChangelogBadgeFlags(ss, piNumber, iterationNumber);
      
      if (badgeFlags.size === 0) {
        throw new Error(`Badge flags not found for Iteration ${iterationNumber}. Run "Analyze Changes for Iteration ${iterationNumber}" first.`);
      }
    }
    
    // 4. Load previous iteration data (needed for dependency processing and change details)
    const previousIterData = iterationNumber > 1 ? loadPreviousIterationData(ss, piNumber, iterationNumber) : {};
    
    // 5. Merge badge flags into iteration data
    iterationData.forEach(item => {
      const flags = badgeFlags.get(item['Key']);
      if (flags) {
        // Merge all pre-computed flags into the item
        Object.assign(item, flags);
      }
    });
    
    // 6. Apply governance filter (exclude KLO/Quality unless INFOSEC)
    let filteredData = applyGovernanceFilter(iterationData, badgeFlags);
    console.log(`  After governance filter: ${filteredData.length} items`);
    
    // 7. Handle INFOSEC bypass and Duplicate exclusion
    const { nonInfosecData, infosecEpics, infosecDependencies, duplicateKeys } = 
      separateSpecialCases(filteredData, iterationNumber);
    
    // 8. Apply changes-only filter based on iteration and options
    if (iterationNumber <= 1) {
      // Iteration 0/1: Baseline - include all (already filtered by governance)
      console.log('Iteration 0/1: Including all data (baseline)');
      filteredData = nonInfosecData;
    } else if (showAllEpics) {
      // Show All Epics: Include all but badges are still applied for display
      console.log('Show All Epics: Including all data with badge indicators');
      filteredData = nonInfosecData;
    } else {
      // Iteration 2+: Filter to only items with changes/badges
      console.log(`Iteration ${iterationNumber}: Filtering by changes`);
      filteredData = applyChangesOnlyFilter(nonInfosecData, includeAtRisk, duplicateKeys);
      console.log(`  After changes-only filter: ${filteredData.length} items`);
    }
    
    // 9. Process dependencies (apply visibility rules based on parent status)
    filteredData = processDependencyVisibility(filteredData, previousIterData, iterationNumber);
    
    // 10. Remove orphaned at-risk-only epics (SKIP when showAllEpics is true)
    if (!showAllEpics) {
      filteredData = removeOrphanedAtRiskEpics(filteredData);
    } else {
      console.log('  Skipping orphan removal (showAllEpics=true)');
    }
    
    // 11. Build change details for each epic (for slide display)
    filteredData.forEach(item => {
      if (item['Issue Type'] === 'Epic') {
        const prev = previousIterData[item['Key']] || {};
        item.changes = buildChangeDetails(item, prev, iterationNumber);
      }
    });
    
    // 12. Add back INFOSEC items (bypass change filtering)
    filteredData = addInfosecItems(filteredData, infosecEpics, infosecDependencies, iterationNumber, previousIterData);
    
    // 13. Process and structure for slides WITH CONFIGURATION
    const processedData = structureForSlides(filteredData, piNumber, iterationNumber, slideConfig);
    
    console.log(`========== PROCESSING COMPLETE ==========`);
    console.log(`Total epics for slides: ${processedData.epics.length}`);
    console.log(`Portfolio initiatives: ${processedData.portfolioInitiatives.length}`);
    
    return processedData;
    
  } catch (error) {
    console.error('Error processing slide data:', error);
    throw error;
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// PRECONDITION VERIFICATION
// ═══════════════════════════════════════════════════════════════════════════════

function verifyPreconditions(ss, piNumber, iterationNumber) {
  console.log('Verifying preconditions...');
  
  // Check iteration sheet exists
  const iterationSheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
  const iterationSheet = ss.getSheetByName(iterationSheetName);
  
  if (!iterationSheet) {
    throw new Error(`Required sheet "${iterationSheetName}" not found. Please generate the iteration report first.`);
  }
  console.log(`✓ Iteration sheet found: ${iterationSheetName}`);
  
  // For iterations > 1, check changelog exists and has badge columns
  if (iterationNumber > 1) {
    const changelogSheetName = `PI ${piNumber} Changelog`;
    const changelogSheet = ss.getSheetByName(changelogSheetName);
    
    if (!changelogSheet) {
      throw new Error(`Required sheet "${changelogSheetName}" not found. Please run changelog analysis first.`);
    }
    console.log(`✓ Changelog sheet found: ${changelogSheetName}`);
    
    // Verify changelog has been run for this iteration (check for Badge column)
    const headers = changelogSheet.getRange(5, 1, 1, changelogSheet.getLastColumn()).getValues()[0];
    const badgeColumn = `Iteration ${iterationNumber} - Badge`;
    const noChangesColumn = `Iteration ${iterationNumber} - NO CHANGES`;
    
    if (!headers.includes(noChangesColumn)) {
      throw new Error(`Changelog has not been analyzed for Iteration ${iterationNumber}. Please run "Analyze Changes for Iteration" first.`);
    }
    
    if (!headers.includes(badgeColumn)) {
      throw new Error(`Badge columns not found for Iteration ${iterationNumber}. Please run "Analyze Changes for Iteration" again to compute badges, or run "Add Badge Columns to Changelog" first.`);
    }
    
    console.log(`✓ Changelog verified for Iteration ${iterationNumber} (with badge columns)`);
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// DATA LOADING
// ═══════════════════════════════════════════════════════════════════════════════

function loadIterationSheetData(ss, piNumber, iterationNumber) {
  const sheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
  const sheet = ss.getSheetByName(sheetName);
  
  console.log(`Loading data from ${sheetName}...`);
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow < 5) {
    throw new Error(`No data found in ${sheetName}`);
  }
  
  const headers = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(5, 1, lastRow - 4, lastCol);
  const values = dataRange.getValues();
  
  const epics = [];
  values.forEach(row => {
    const epic = {};
    headers.forEach((header, index) => {
      epic[header] = row[index];
    });
    
    // Only include Epics and Dependencies
    if (epic['Issue Type'] === 'Epic' || epic['Issue Type'] === 'Dependency') {
      epics.push(epic);
    }
  });
  
  console.log(`Loaded ${epics.length} epics/dependencies from iteration sheet`);
  return epics;
}

function loadPreviousIterationData(ss, piNumber, iterationNumber) {
  if (iterationNumber <= 1) {
    return {};
  }
  
  const prevSheetName = `PI ${piNumber} - Iteration ${iterationNumber - 1}`;
  const prevSheet = ss.getSheetByName(prevSheetName);
  
  if (!prevSheet) {
    console.log(`  Previous iteration sheet "${prevSheetName}" not found`);
    return {};
  }
  
  const headers = prevSheet.getRange(4, 1, 1, prevSheet.getLastColumn()).getValues()[0];
  const lastRow = prevSheet.getLastRow();
  
  if (lastRow < 5) {
    return {};
  }
  
  const data = prevSheet.getRange(5, 1, lastRow - 4, prevSheet.getLastColumn()).getValues();
  const result = {};
  
  data.forEach(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    
    if (item['Key']) {
      result[item['Key']] = item;
    }
  });
  
  console.log(`  Loaded ${Object.keys(result).length} items from previous iteration`);
  return result;
}


// ═══════════════════════════════════════════════════════════════════════════════
// BADGE FLAG LOADING (FROM CHANGELOG)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Loads pre-computed badge flags from the changelog sheet
 * This replaces the complex filterByChanges() badge computation
 * 
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {number} piNumber - PI number
 * @param {number} iterationNumber - Current iteration
 * @returns {Map<string, Object>} Map of issue key to badge flags
 */
function loadChangelogBadgeFlags(ss, piNumber, iterationNumber) {
  const changelogSheetName = `PI ${piNumber} Changelog`;
  const changelogSheet = ss.getSheetByName(changelogSheetName);
  
  if (!changelogSheet) {
    throw new Error(`Changelog sheet "${changelogSheetName}" not found. Run "Analyze Changes for Iteration" first.`);
  }
  
  console.log(`  Loading badge flags from ${changelogSheetName}...`);
  
  const lastCol = changelogSheet.getLastColumn();
  const headers = changelogSheet.getRange(5, 1, 1, lastCol).getValues()[0];
  
  // Build header index map
  const headerIndexMap = {};
  headers.forEach((header, index) => {
    headerIndexMap[header] = index;
  });
  
  // Find required columns
  const keyCol = headerIndexMap['Key'];
  const issueTypeCol = headerIndexMap['Issue Type'];
  const governanceCol = headerIndexMap['Include in Governance'];
  const addedInIterCol = headerIndexMap['Added in Iteration'];
  
  if (keyCol === undefined) {
    throw new Error('Key column not found in changelog');
  }
  
  // Find badge columns for this iteration
  const iterPrefix = `Iteration ${iterationNumber}`;
  const cols = {
    badge: headerIndexMap[`${iterPrefix} - Badge`],
    statusBadge: headerIndexMap[`${iterPrefix} - Status Badge`],
    statusNote: headerIndexMap[`${iterPrefix} - Status Note`],
    atRisk: headerIndexMap[`${iterPrefix} - At Risk`],
    iterRisk: headerIndexMap[`${iterPrefix} - Iteration Risk`],
    closedThisIter: headerIndexMap[`${iterPrefix} - Closed This Iter`],
    deferredThisIter: headerIndexMap[`${iterPrefix} - Deferred This Iter`],
    canceledThisIter: headerIndexMap[`${iterPrefix} - Canceled This Iter`],
    isNew: headerIndexMap[`${iterPrefix} - Is New`],
    reasons: headerIndexMap[`${iterPrefix} - Qualifying Reasons`],
    noChanges: headerIndexMap[`${iterPrefix} - NO CHANGES`]
  };
  
  // Verify badge columns exist
  if (cols.badge === undefined) {
    console.warn(`  ⚠️ Badge columns not found for Iteration ${iterationNumber}`);
    console.warn(`  Run "Analyze Changes for Iteration ${iterationNumber}" to compute badges`);
    return new Map();
  }
  
  // Read all data
  const lastRow = changelogSheet.getLastRow();
  if (lastRow < 6) {
    console.log(`  No data rows in changelog`);
    return new Map();
  }
  
  const data = changelogSheet.getRange(6, 1, lastRow - 5, lastCol).getValues();
  const flagsMap = new Map();
  
  // Process each row
  data.forEach(row => {
    const key = row[keyCol];
    if (!key) return;
    
    const badge = row[cols.badge] || '';
    const statusNote = row[cols.statusNote] || '';
    const isAtRisk = row[cols.atRisk] === 'Yes';
    const isIterationRisk = row[cols.iterRisk] === 'Yes';
    const closedThisIteration = row[cols.closedThisIter] === 'Yes';
    const deferredThisIteration = row[cols.deferredThisIter] === 'Yes';
    const canceledThisIteration = cols.canceledThisIter !== undefined ? row[cols.canceledThisIter] === 'Yes' : false;
    const isNew = row[cols.isNew] === 'Yes';
    const qualifyingReasons = (row[cols.reasons] || '').split('; ').filter(r => r);
    const noChanges = row[cols.noChanges];
    const includeInGovernance = governanceCol !== undefined ? row[governanceCol] : 'Yes';
    const addedInIter = addedInIterCol !== undefined ? row[addedInIterCol] : null;
    
    // Derive convenience flags from badge
    const isDone = badge === 'DONE';
    const isPendingClosure = badge === 'PENDING';
    const isDeferred = badge === 'DEF';
    const isCanceled = badge === 'CANCELED';
    const isModified = badge === 'CHG';
    const isOverdue = badge === 'OVERDUE';
    
    // Determine if this was already done/deferred/canceled/pending in previous iterations
    const alreadyClosed = isDone && !closedThisIteration;
    const alreadyPendingClosure = isPendingClosure && !closedThisIteration; // Uses closedThisIter as proxy
    const alreadyDeferred = isDeferred && !deferredThisIteration;
    const alreadyCanceled = isCanceled && !canceledThisIteration;
    
    // Derive pendingClosureThisIteration from badge + status note
    const pendingClosureThisIteration = isPendingClosure && 
      (statusNote.includes('this iteration') || closedThisIteration);
    
    flagsMap.set(key, {
      // Primary badge
      badge: badge,
      statusNote: statusNote,
      
      // Boolean flags
      isNew: isNew,
      isModified: isModified,
      isDone: isDone,
      isPendingClosure: isPendingClosure,
      isDeferred: isDeferred,
      isCanceled: isCanceled,
      isOverdue: isOverdue,
      isAtRisk: isAtRisk,
      isIterationRisk: isIterationRisk,
      
      // Timing flags
      closedThisIteration: closedThisIteration,
      pendingClosureThisIteration: pendingClosureThisIteration,
      deferredThisIteration: deferredThisIteration,
      canceledThisIteration: canceledThisIteration,
      alreadyClosed: alreadyClosed,
      alreadyPendingClosure: alreadyPendingClosure,
      alreadyDeferred: alreadyDeferred,
      alreadyCanceled: alreadyCanceled,
      
      // Audit info
      qualifyingReasons: qualifyingReasons,
      noChanges: noChanges === 'NO CHANGES' || noChanges === 'Baseline',
      addedInIteration: addedInIter,
      
      // Governance inclusion
      includeInGovernance: includeInGovernance === 'Yes',
      
      // Issue type
      issueType: issueTypeCol !== undefined ? row[issueTypeCol] : ''
    });
  });
  
  console.log(`  Loaded ${flagsMap.size} badge flags from changelog`);
  
  // Log badge distribution
  let stats = { NEW: 0, CHG: 0, DONE: 0, PENDING: 0, DEF: 0, CANCELED: 0, OVERDUE: 0, ATRISK: 0 };
  flagsMap.forEach(flags => {
    if (flags.badge) stats[flags.badge] = (stats[flags.badge] || 0) + 1;
    if (flags.isAtRisk && !flags.badge) stats.ATRISK++;
  });
  console.log(`  Badge distribution: NEW=${stats.NEW}, CHG=${stats.CHG}, DONE=${stats.DONE}, PENDING=${stats.PENDING}, DEF=${stats.DEF}, CANCELED=${stats.CANCELED}, OVERDUE=${stats.OVERDUE}, ATRISK=${stats.ATRISK}`);
  
  return flagsMap;
}


// ═══════════════════════════════════════════════════════════════════════════════
// FILTERING FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Applies governance filter using pre-computed includeInGovernance flag
 * Also applies global KLO/Quality exclusion unless INFOSEC
 */
function applyGovernanceFilter(iterationData, badgeFlags) {
  const beforeCount = iterationData.length;
  
  const filtered = iterationData.filter(item => {
    const key = item['Key'];
    const flags = badgeFlags.get(key);
    
    // If we have flags and governance is explicitly 'No', exclude
    if (flags && flags.includeInGovernance === false) {
      return false;
    }
    
    // For dependencies, check parent's governance status
    if (item['Issue Type'] === 'Dependency') {
      const parentKey = item['Parent Key'];
      const parentFlags = badgeFlags.get(parentKey);
      if (parentFlags && parentFlags.includeInGovernance === false) {
        return false;
      }
    }
    
    // Global filter: Exclude KLO/Quality allocation UNLESS Portfolio = INFOSEC
    const allocation = (item['Allocation'] || '').toLowerCase().trim();
    const portfolio = (item['Portfolio Initiative'] || '').toUpperCase().trim();
    const isKloOrQuality = allocation === 'klo' || allocation === 'quality' || allocation === 'klo/quality';
    
    if (isKloOrQuality && !portfolio.startsWith('INFOSEC')) {
      return false;
    }
    
    return true;
  });
  
  const excludedCount = beforeCount - filtered.length;
  if (excludedCount > 0) {
    console.log(`  Governance filter: excluded ${excludedCount} items`);
  }
  
  return filtered;
}

/**
 * Separates INFOSEC items and identifies duplicates for special handling
 */
function separateSpecialCases(iterationData, iterationNumber) {
  const infosecEpics = [];  // Keep name for compatibility
  const infosecDependencies = [];  // Keep name for compatibility
  const nonInfosecData = [];
  const duplicateKeys = new Set();
  
  // First pass: identify duplicates
  iterationData.forEach(item => {
    if (item['Issue Type'] === 'Epic') {
      const resolution = (item['Resolution'] || '').toString().trim().toLowerCase();
      if (resolution === 'duplicate') {
        duplicateKeys.add(item['Key']);
        console.log(`  🚫 Excluding ${item['Key']} - Resolution = "Duplicate"`);
      }
    }
  });
  
  // Second pass: separate show-all portfolios vs normal
  iterationData.forEach(item => {
    const key = item['Key'];
    
    // Skip duplicates
    if (item['Issue Type'] === 'Epic' && duplicateKeys.has(key)) {
      return;
    }
    if (item['Issue Type'] === 'Dependency' && duplicateKeys.has(item['Parent Key'])) {
      return;
    }
    
    const portfolioInit = (item['Portfolio Initiative'] || '').toString().trim();
    
    // Check if this portfolio should show all epics (configurable)
    if (shouldShowAllEpics(portfolioInit)) {
      item.isInfosecBypass = true;  // Keep flag name for compatibility
      item.isShowAllPortfolio = true;
      item.showAllReason = portfolioInit;
      
      if (item['Issue Type'] === 'Epic') {
        infosecEpics.push(item);
      } else if (item['Issue Type'] === 'Dependency') {
        infosecDependencies.push(item);
      }
    } else {
      nonInfosecData.push(item);
    }
  });
  
  if (infosecEpics.length > 0) {
    const showAllConfig = getPortfolioShowAllConfig();
    console.log(`  📋 SHOW-ALL PORTFOLIOS (${showAllConfig.join(', ')}): ${infosecEpics.length} epics + ${infosecDependencies.length} dependencies`);
  }
  if (duplicateKeys.size > 0) {
    console.log(`  🚫 Excluded ${duplicateKeys.size} duplicate epics`);
  }
  
  return { nonInfosecData, infosecEpics, infosecDependencies, duplicateKeys };
}

/**
 * Applies the changes-only filter for governance reports
 * Only includes items with badges or that meet at-risk/previously-closed criteria
 */
function applyChangesOnlyFilter(iterationData, includeAtRisk, duplicateKeys) {
  // First, determine which epics should be included
  const includedEpicKeys = new Set();
  
  iterationData.forEach(item => {
    if (item['Issue Type'] !== 'Epic') return;
    
    const key = item['Key'];
    
    // Skip duplicates
    if (duplicateKeys.has(key)) return;
    
    // Has any badge = potentially include (check options)
    if (item.badge) {
      // ATRISK badge - check if includeAtRisk is true
      if (item.badge === 'ATRISK') {
        if (includeAtRisk) {
          includedEpicKeys.add(key);
        }
        return;
      }
      
      if (item.badge === 'DONE') {
        item.includeInGovernance = true;
        return true;
      }
      
      // DEF badge with alreadyDeferred - always include (important to show)
      if (item.badge === 'DEF' && item.alreadyDeferred) {
        includedEpicKeys.add(key);
        return;
      }
      
      // CANCELED badge - always include (important to show status)
      if (item.badge === 'CANCELED') {
        includedEpicKeys.add(key);
        return;
      }
      
      // All other badges (NEW, CHG, DONE this iter, DEF this iter) = include
      includedEpicKeys.add(key);
      return;
    }

    // No badge but has iteration risk - include
    if (item.isIterationRisk) {
      includedEpicKeys.add(key);
      return;
    }
    
    // Check for at-risk RAG (fallback if badge wasn't set)
    if (includeAtRisk) {
      const rag = (item['RAG'] || '').toString().trim().toUpperCase();
      if (rag === 'AMBER' || rag === 'RED') {
        includedEpicKeys.add(key);
      }
    }
  });
  
  // Now filter the data - include matching epics and their dependencies
  return iterationData.filter(item => {
    const issueType = item['Issue Type'];
    const key = item['Key'];
    
    if (issueType === 'Epic') {
      return includedEpicKeys.has(key);
    } else if (issueType === 'Dependency') {
      // Dependencies: include if parent is included
      const parentKey = item['Parent Key'];
      return includedEpicKeys.has(parentKey);
    }
    
    return false;
  });
}


// ═══════════════════════════════════════════════════════════════════════════════
// DEPENDENCY VISIBILITY PROCESSING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Processes dependencies to determine visibility based on business rules
 * 
 * Business Rules:
 * - Canceled previously: HIDE
 * - Canceled this iteration: SHOW with CANCELED badge
 * - Deferred previously: HIDE (or show with DEF based on settings)
 * - Deferred this iteration: SHOW with DEF badge
 * - Done this iteration + parent has changes: SHOW with CHG + DONE
 * - Done this iteration + parent at-risk only: HIDE
 * - Done previously + parent has changes: SHOW with DONE
 * - Done previously + parent at-risk only: HIDE
 */
function processDependencyVisibility(filteredData, previousIterData, iterationNumber) {
  console.log('\n  Processing dependency visibility...');
  
  // Build a lookup of epic status for parent checks
  const epicStatusLookup = {};
  
  filteredData.forEach(item => {
    if (item['Issue Type'] === 'Epic') {
      epicStatusLookup[item['Key']] = {
        hasChanges: item.badge === 'NEW' || item.badge === 'CHG' || item.closedThisIteration || item.deferredThisIteration,
        isAtRiskOnly: item.badge === 'ATRISK' || (item.isAtRisk && !item.badge),
        badge: item.badge
      };
    }
  });
  
  // Track stats
  let depHiddenCount = 0;
  
  // Process each dependency to set visibility flags
  filteredData.forEach(item => {
    if (item['Issue Type'] !== 'Dependency') return;
    
    const key = item['Key'];
    const parentKey = item['Parent Key'];
    const parentStatus = epicStatusLookup[parentKey] || { hasChanges: false, isAtRiskOnly: false };
    
    // Get current and previous values
    const status = (item['Status'] || '').toString().trim().toUpperCase();
    const piCommitment = (item['PI Commitment'] || '').toString().trim().toUpperCase();
    const prev = previousIterData[key] || {};
    const prevStatus = (prev['Status'] || '').toString().trim().toUpperCase();
    const prevCommitment = (prev['PI Commitment'] || '').toString().trim().toUpperCase();
    
    // Check DONE status
    const isNowDone = status === 'DONE' || status === 'CLOSED';
    const wasNotDone = prevStatus !== 'DONE' && prevStatus !== 'CLOSED';
    
    item.depIsDone = isNowDone;
    item.depDoneThisIteration = isNowDone && wasNotDone;
    item.depDonePreviously = isNowDone && !wasNotDone;
    
    // Check CANCELED status
    const isNowCanceled = piCommitment === 'CANCELED' || piCommitment === 'CANCELLED';
    const wasNotCanceled = prevCommitment !== 'CANCELED' && prevCommitment !== 'CANCELLED';
    
    item.depIsCanceled = isNowCanceled;
    item.depCanceledThisIteration = isNowCanceled && wasNotCanceled;
    item.depCanceledPreviously = isNowCanceled && !wasNotCanceled;
    
    // Check DEFERRED status (Not Committed, Deferred)
    const deferredValues = ['NOT COMMITTED', 'DEFERRED'];
    const isNowDeferred = deferredValues.includes(piCommitment);
    const wasNotDeferred = !deferredValues.includes(prevCommitment);
    
    item.depIsDeferred = isNowDeferred;
    item.depDeferredThisIteration = isNowDeferred && wasNotDeferred;
    item.depDeferredPreviously = isNowDeferred && !wasNotDeferred;
    
    // Check ITERATION RISK - Dependency due this iteration but not complete
    const depEndIterName = (item['End Iteration Name'] || '').toString().trim();
    const depIterMatch = depEndIterName.match(/iteration\s*(\d+)/i);
    const depIterNum = depIterMatch ? parseInt(depIterMatch[1], 10) : null;
    
    item.depIsIterationRisk = false;
    if (depIterNum !== null && depIterNum === iterationNumber && !isNowDone && !isNowCanceled && !isNowDeferred) {
      item.depIsIterationRisk = true;
    }
    
    // Initialize visibility
    item.depShouldShow = true;
    item.depBadges = [];
    item.depHideReason = '';
    
    // Apply visibility rules in priority order
    
    // CANCELED rules
    if (item.depCanceledPreviously) {
      item.depShouldShow = false;
      item.depHideReason = 'Canceled in previous iteration';
      depHiddenCount++;
    } else if (item.depCanceledThisIteration) {
      item.depShouldShow = true;
      item.depBadges.push('CANCELED');
    }
    // DEFERRED rules
    else if (item.depDeferredThisIteration) {
      item.depShouldShow = true;
      item.depBadges.push('DEF');
    } else if (item.depDeferredPreviously) {
      item.depShouldShow = true;
      item.depBadges.push('DEF');
    }
    // DONE rules
    else if (item.depDoneThisIteration) {
      if (parentStatus.hasChanges) {
        item.depShouldShow = true;
        item.depBadges.push('CHG');
        item.depBadges.push('DONE');
      } else {
        item.depShouldShow = false;
        item.depHideReason = 'Done this iteration, parent at-risk only';
        depHiddenCount++;
      }
    } else if (item.depDonePreviously) {
      if (parentStatus.hasChanges) {
        item.depShouldShow = true;
        item.depBadges.push('DONE');
      } else {
        item.depShouldShow = false;
        item.depHideReason = 'Done previously, parent at-risk only';
        depHiddenCount++;
      }
    }
    
    // Add RISK badge if due this iteration but not complete
    if (item.depIsIterationRisk && item.depShouldShow) {
      item.depBadges.unshift('RISK');
    }
  });
  
  // Filter out hidden dependencies
  const visibleData = filteredData.filter(item => {
    if (item['Issue Type'] === 'Dependency') {
      if (!item.depShouldShow) {
        console.log(`    🚫 ${item['Key']}: HIDE - ${item.depHideReason}`);
        return false;
      }
    }
    return true;
  });
  
  if (depHiddenCount > 0) {
    console.log(`  Hidden ${depHiddenCount} dependencies based on visibility rules`);
  }
  
  return visibleData;
}


// ═══════════════════════════════════════════════════════════════════════════════
// ORPHANED EPIC REMOVAL
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Removes at-risk-only epics that have no visible at-risk dependencies
 */
function removeOrphanedAtRiskEpics(filteredData) {
  console.log('\n  Checking for orphaned at-risk-only epics...');
  
  // Build a map of visible at-risk dependencies per parent epic
  const visibleAtRiskDepsByParent = {};
  filteredData.forEach(item => {
    if (item['Issue Type'] !== 'Dependency') return;
    if (item.depShouldShow === false) return;
    
    const rag = (item['RAG'] || '').toString().trim().toUpperCase();
    if (rag !== 'AMBER' && rag !== 'RED') return;
    
    const parentKey = item['Parent Key'];
    if (!visibleAtRiskDepsByParent[parentKey]) {
      visibleAtRiskDepsByParent[parentKey] = [];
    }
    visibleAtRiskDepsByParent[parentKey].push(item['Key']);
  });
  
  // Identify epics to remove
  const epicsToRemove = new Set();
  
  filteredData.forEach(item => {
    if (item['Issue Type'] !== 'Epic') return;
    
    const key = item['Key'];
    const epicRag = (item['RAG'] || '').toString().trim().toUpperCase();
    const epicIsAtRisk = epicRag === 'AMBER' || epicRag === 'RED';
    
    // Check if this epic has any actual changes
    const hasRealChanges = item.isNew || item.isModified || item.closedThisIteration || 
                           item.deferredThisIteration || item.isIterationRisk ||
                           item.badge === 'NEW' || item.badge === 'CHG' || 
                           item.badge === 'DONE' || item.badge === 'DEF';
    const hasAlreadyClosedOrDeferred = item.alreadyClosed || item.alreadyDeferred;
    
    // If epic has real changes or is already closed/deferred, keep it
    if (hasRealChanges || hasAlreadyClosedOrDeferred) return;
    
    // If the EPIC ITSELF is at-risk (amber/red), keep it
    if (epicIsAtRisk) return;
    
    // At this point, the epic has no changes and no at-risk RAG itself
    // It was ONLY included because of at-risk dependencies
    
    // Check if it has any VISIBLE at-risk dependencies
    const visibleAtRiskDeps = visibleAtRiskDepsByParent[key] || [];
    
    if (visibleAtRiskDeps.length === 0) {
      epicsToRemove.add(key);
      console.log(`    🗑️ Removing ${key}: No changes, no at-risk RAG, no visible at-risk dependencies`);
    }
  });
  
  // Remove the orphaned epics and their dependencies
  if (epicsToRemove.size > 0) {
    const beforeCount = filteredData.length;
    
    const result = filteredData.filter(item => {
      const key = item['Key'];
      const parentKey = item['Parent Key'];
      
      if (item['Issue Type'] === 'Epic' && epicsToRemove.has(key)) {
        return false;
      }
      if (item['Issue Type'] === 'Dependency' && epicsToRemove.has(parentKey)) {
        return false;
      }
      return true;
    });
    
    console.log(`    Removed ${epicsToRemove.size} orphaned epics (${beforeCount - result.length} total items)`);
    return result;
  }
  
  console.log(`    No orphaned epics found`);
  return filteredData;
}


// ═══════════════════════════════════════════════════════════════════════════════
// INFOSEC BYPASS - ADD BACK INFOSEC ITEMS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Adds INFOSEC items back (they bypass change filtering)
 */
function addInfosecItems(filteredData, infosecEpics, infosecDependencies, iterationNumber, previousIterData = {}) {
  if (infosecEpics.length === 0) {
    return filteredData;
  }
  
  console.log('\n  Adding show-all portfolio items (bypass change filtering)...');  
  // Add INFOSEC epics
  infosecEpics.forEach(epic => {
    const key = epic['Key'];
    const alreadyIncluded = filteredData.some(item => item['Key'] === key);
    
    if (!alreadyIncluded) {
      // Ensure fields are strings
      if (epic['RAG'] !== undefined && epic['RAG'] !== null) {
        epic['RAG'] = String(epic['RAG']);
      } else {
        epic['RAG'] = '';
      }
      
      // Check for iteration risk
      const endIterName = (epic['End Iteration Name'] || '').toString().trim();
      const iterMatch = endIterName.match(/iteration\s*(\d+)/i);
      const epicIterNum = iterMatch ? parseInt(iterMatch[1], 10) : null;
      const currentStatus = (epic['Status'] || '').toString().trim().toUpperCase();
      const isNowClosed = currentStatus === 'DONE' || currentStatus === 'CLOSED';
      
      epic.isIterationRisk = (epicIterNum !== null && epicIterNum === iterationNumber && !isNowClosed);
      
      // Set minimal display flags (no badges for bypassed items)
      epic.isNew = false;
      epic.isModified = false;
      epic.isDone = false;
      epic.isDeferred = false;
      epic.isAtRisk = false;
      epic.closedThisIteration = false;
      epic.deferredThisIteration = false;
      epic.alreadyClosed = false;
      epic.alreadyDeferred = false;
      epic.qualifyingChanges = [];
      const prev = previousIterData[key] || {};
      epic.changes = buildChangeDetails(epic, prev, iterationNumber);
      filteredData.push(epic);
    }
  });
  
  // Add INFOSEC dependencies
  infosecDependencies.forEach(dep => {
    const key = dep['Key'];
    const parentKey = dep['Parent Key'];
    
    const parentIncluded = filteredData.some(item => item['Key'] === parentKey);
    const alreadyIncluded = filteredData.some(item => item['Key'] === key);
    
    if (parentIncluded && !alreadyIncluded) {
      if (dep['RAG'] !== undefined && dep['RAG'] !== null) {
        dep['RAG'] = String(dep['RAG']);
      } else {
        dep['RAG'] = '';
      }
      
      // Check for iteration risk
      const depEndIterName = (dep['End Iteration Name'] || '').toString().trim();
      const depIterMatch = depEndIterName.match(/iteration\s*(\d+)/i);
      const depIterNum = depIterMatch ? parseInt(depIterMatch[1], 10) : null;
      const depStatus = (dep['Status'] || '').toString().trim().toUpperCase();
      const depIsNowDone = depStatus === 'DONE' || depStatus === 'CLOSED';
      
      dep.depIsIterationRisk = (depIterNum !== null && depIterNum === iterationNumber && !depIsNowDone);
      dep.depShouldShow = true;
      dep.depBadges = [];
      
      if (dep.depIsIterationRisk) {
        dep.depBadges.push('RISK');
      }
      
      filteredData.push(dep);
    }
  });
  
  const finalInfosecEpics = filteredData.filter(i => i.isInfosecBypass && i['Issue Type'] === 'Epic').length;
  const finalInfosecDeps = filteredData.filter(i => i.isInfosecBypass && i['Issue Type'] === 'Dependency').length;
  const finalShowAllEpics = filteredData.filter(i => i.isInfosecBypass && i['Issue Type'] === 'Epic').length;
  const finalShowAllDeps = filteredData.filter(i => i.isInfosecBypass && i['Issue Type'] === 'Dependency').length;
  console.log(`  Added ${finalShowAllEpics} show-all portfolio epics + ${finalShowAllDeps} dependencies`);
  
  return filteredData;
}


// ═══════════════════════════════════════════════════════════════════════════════
// CHANGE DETAILS BUILDER (FOR SLIDE DISPLAY)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Builds change details object for slide display
 * This tracks field changes between iterations for strikethrough/highlight display
 */
function buildChangeDetails(current, previous, currentIterationNumber) {
  const changes = {
    hasChanges: false,
    
    // Iteration changes - for strikethrough display
    iterationChanged: false,
    iterationPulledIn: false,    // Positive: moved to earlier iteration
    iterationPushedOut: false,   // Negative: moved to later iteration
    previousIteration: null,
    currentIteration: null,
    
    // RAG changes
    ragChanged: false,
    ragMitigated: false,         // Positive: Yellow/Red → Green
    ragNoteChanged: false,
    previousRag: null,
    currentRag: null,
    previousRagNote: null,
    currentRagNote: null,
    ragNotes: [],
    
    // Track if RAG is unchanged from previous iteration
    ragUnchangedFromPrevious: false,
    ragUnchangedIterations: 0,
    
    // Depends on Valuestream changes
    dependsOnChanged: false,
    previousDependsOn: null,
    currentDependsOn: null,
    
    // Fix Versions
    fixVersions: null,
    fixVersionsChanged: false,
    previousFixVersions: null,
    currentFixVersions: null,
    
    // PI Commitment changes
    piCommitmentChanged: false,
    piCommitmentToCommitted: false,
    piCommitmentFromCommitted: false,
    previousPiCommitment: null,
    currentPiCommitment: null,
    
    // Program Increment
    programIncrementChanged: false,
    programIncrementFromBlank: false,
    previousProgramIncrement: null,
    currentProgramIncrement: null
  };
  
  // Get current values
  const currIterEnd = (current['End Iteration Name'] || current['PI Target Iteration'] || '').toString().trim();
  const currRag = (current['RAG'] || '').toString().trim();
  const currRagNote = (current['RAG Note'] || '').toString().trim();
  const currDependsOn = (current['Depends on Valuestream'] || '').toString().trim();
  const currFixVersions = (current['Fix Versions'] || '').toString().trim();
  const currPiCommitment = (current['PI Commitment'] || '').toString().trim();
  const currProgramIncrement = (current['Program Increment'] || '').toString().trim();
  
  // Get previous values
  const prevIterEnd = (previous['End Iteration Name'] || previous['PI Target Iteration'] || '').toString().trim();
  const prevRag = (previous['RAG'] || '').toString().trim();
  const prevRagNote = (previous['RAG Note'] || '').toString().trim();
  const prevDependsOn = (previous['Depends on Valuestream'] || '').toString().trim();
  const prevFixVersions = (previous['Fix Versions'] || '').toString().trim();
  const prevPiCommitment = (previous['PI Commitment'] || '').toString().trim();
  const prevProgramIncrement = (previous['Program Increment'] || '').toString().trim();
  
  // Check iteration change and direction
  if (currIterEnd !== prevIterEnd && prevIterEnd !== '') {
    changes.hasChanges = true;
    changes.iterationChanged = true;
    changes.previousIteration = prevIterEnd;
    changes.currentIteration = currIterEnd;
    
    // Determine direction: extract iteration numbers and compare
    const prevIterNum = extractIterationNumber(prevIterEnd);
    const currIterNum = extractIterationNumber(currIterEnd);
    
    if (prevIterNum !== null && currIterNum !== null) {
      if (currIterNum < prevIterNum) {
        changes.iterationPulledIn = true;
        console.log(`✅ ITERATION PULLED IN: ${current['Key']}: "${prevIterEnd}" → "${currIterEnd}" (Iter ${prevIterNum} → ${currIterNum})`);
      } else if (currIterNum > prevIterNum) {
        changes.iterationPushedOut = true;
        console.log(`⚠️ ITERATION PUSHED OUT: ${current['Key']}: "${prevIterEnd}" → "${currIterEnd}" (Iter ${prevIterNum} → ${currIterNum})`);
      }
    } else {
      console.log(`✅ ITERATION CHANGED: ${current['Key']}: "${prevIterEnd}" → "${currIterEnd}"`);
    }
  }
  
  // Check RAG change and detect mitigation (Yellow/Red → Green)
  const currRagUpper = currRag.toUpperCase();
  const prevRagUpper = prevRag.toUpperCase();
  const isBlankToGreen = prevRag === '' && currRagUpper.includes('GREEN');
  
  if (currRag !== prevRag && !isBlankToGreen) {
    changes.hasChanges = true;
    changes.ragChanged = true;
    changes.previousRag = prevRag;
    changes.currentRag = currRag;
    
    // Detect RAG mitigation: Yellow/Red → Green
    const prevWasAtRisk = (prevRagUpper === 'AMBER' || prevRagUpper === 'RED' || 
                          prevRagUpper === 'YELLOW');
    const currIsGreen = currRagUpper.includes('GREEN');
    
    if (prevWasAtRisk && currIsGreen) {
      changes.ragMitigated = true;
      console.log(`✅ RAG MITIGATED: ${current['Key']}: "${prevRag}" → "${currRag}"`);
    }
  } else {
    if ((currRagUpper === 'AMBER' || currRagUpper === 'RED') && currRagUpper === prevRagUpper) {
      changes.ragUnchangedFromPrevious = true;
    }
  }
  
  // Check RAG Note change
  const isStatusOnlyNote = currRagNote.toLowerCase() === 'green' || 
                           currRagNote.toLowerCase().startsWith('green -') ||
                           currRagNote.toLowerCase() === 'on track';
  const isBlankToStatusOnly = prevRagNote === '' && isStatusOnlyNote;
  
  if (currRagNote !== prevRagNote && !isBlankToStatusOnly) {
    changes.hasChanges = true;
    changes.ragNoteChanged = true;
    changes.previousRagNote = prevRagNote;
    changes.currentRagNote = currRagNote;
  }
  
  // Build RAG notes array for at-risk epics
  if (currRagUpper === 'AMBER' || currRagUpper === 'RED') {
    if (currRagNote) {
      changes.ragNotes.push({
        iteration: currentIterationNumber,
        note: currRagNote,
        isCurrent: true
      });
    }
    
    const prevWasAtRisk = (prevRagUpper === 'AMBER' || prevRagUpper === 'RED');
    
    if (prevRagNote && prevRagNote !== currRagNote) {
      if (changes.ragNoteChanged || prevWasAtRisk) {
        changes.ragNotes.push({
          iteration: currentIterationNumber - 1,
          note: prevRagNote,
          isCurrent: false
        });
      }
    }
  }
  
  // Check Depends on Valuestream change
  if (currDependsOn !== prevDependsOn && prevDependsOn !== '') {
    changes.dependsOnChanged = true;
    changes.previousDependsOn = prevDependsOn;
    changes.currentDependsOn = currDependsOn;
  }
  
  // Check Fix Versions change
  if (currFixVersions) {
    changes.fixVersions = currFixVersions;
    changes.currentFixVersions = currFixVersions;
  }
  
  if (currFixVersions !== prevFixVersions && prevFixVersions !== '') {
    changes.hasChanges = true;
    changes.fixVersionsChanged = true;
    changes.previousFixVersions = prevFixVersions;
    changes.currentFixVersions = currFixVersions || '';
  }
  
  // Check PI Commitment change
  if (currPiCommitment !== prevPiCommitment) {
    changes.hasChanges = true;
    changes.piCommitmentChanged = true;
    changes.previousPiCommitment = prevPiCommitment;
    changes.currentPiCommitment = currPiCommitment;
    
    const prevUpper = prevPiCommitment.toUpperCase();
    const currUpper = currPiCommitment.toUpperCase();
    
    const committedValues = ['COMMITTED', 'COMMITTED AFTER PLAN', 'COMMITTED AFTER PLANNING'];
    const wasCommitted = committedValues.includes(prevUpper);
    const isNowCommitted = committedValues.includes(currUpper);
    
    if (!wasCommitted && isNowCommitted) {
      changes.piCommitmentToCommitted = true;
    } else if (wasCommitted && !isNowCommitted) {
      changes.piCommitmentFromCommitted = true;
    }
  }
  
  // Check Program Increment change
  if (currProgramIncrement !== prevProgramIncrement) {
    changes.hasChanges = true;
    changes.programIncrementChanged = true;
    changes.previousProgramIncrement = prevProgramIncrement;
    changes.currentProgramIncrement = currProgramIncrement;
    
    if (prevProgramIncrement === '' && currProgramIncrement !== '') {
      changes.programIncrementFromBlank = true;
    }
  }
  
  return changes;
}


// ═══════════════════════════════════════════════════════════════════════════════
// STRUCTURE FOR SLIDES
// ═══════════════════════════════════════════════════════════════════════════════

function structureForSlides(filteredData, piNumber, iterationNumber, slideConfig = null) {
  console.log('Structuring data for slides with configuration support...');
  
  const epics = filteredData.filter(item => item['Issue Type'] === 'Epic');
  const allDependencies = filteredData.filter(item => item['Issue Type'] === 'Dependency');
  
  // Apply configuration filtering and ordering to epics
  let processedEpics = applyConfigurationToEpics(epics, slideConfig);
  
  // Get unique portfolio initiatives WITH CONFIGURATION
  const portfolioInitiatives = getConfiguredPortfolioInitiatives(processedEpics, slideConfig);
  
  console.log(`Processed ${processedEpics.length} epics across ${portfolioInitiatives.length} portfolio initiatives`);
  
  // Build the final data structure
  const slideMetadata = {
    piNumber: piNumber,
    iterationNumber: iterationNumber,
    isBaseline: iterationNumber === 1,
    title: iterationNumber === 1 
      ? `PI ${piNumber} - Full Governance Report`
      : `PI ${piNumber} - Iteration ${iterationNumber} Changes`,
    generatedAt: new Date().toISOString()
  };
  
  return {
    epics: processedEpics,
    dependencies: allDependencies,
    portfolioInitiatives: portfolioInitiatives,
    metadata: slideMetadata,
    stats: {
      totalEpics: processedEpics.length,
      totalDependencies: allDependencies.length,
      portfolioCount: portfolioInitiatives.length,
      epicsWithDependencies: processedEpics.filter(e => {
        const epicDeps = allDependencies.filter(dep => dep['Parent Key'] === e['Key']);
        return epicDeps.length > 0;
      }).length
    }
  };
}


// ═══════════════════════════════════════════════════════════════════════════════
// CONFIGURATION FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Applies configuration to epics for filtering and ordering
 */
function applyConfigurationToEpics(epics, slideConfig) {
  // Global filter: Exclude KLO/Quality allocation UNLESS Portfolio = INFOSEC
  const beforeGlobalFilter = epics.length;
  let filtered = epics.filter(epic => {
    const allocation = (epic['Allocation'] || '').toLowerCase().trim();
    const portfolio = (epic['Portfolio Initiative'] || '').toUpperCase().trim();
    const isKloOrQuality = allocation === 'klo' || allocation === 'quality' || allocation === 'klo/quality';
    
    if (isKloOrQuality && !portfolio.startsWith('INFOSEC')) {
      return false;
    }
    return true;
  });
  
  const excludedByGlobal = beforeGlobalFilter - filtered.length;
  if (excludedByGlobal > 0) {
    console.log(`  GLOBAL FILTER: Excluded ${excludedByGlobal} KLO/Quality epics (non-INFOSEC)`);
  }
  
  if (!slideConfig) {
    console.log('No configuration - using all epics with default ordering');
    return filtered;
  }
  
  console.log('Applying configuration filters to epics...');
  
  // Filter by PI Commitment if configured
  if (slideConfig['PI Commitment']) {
    filtered = filterByField(filtered, 'PI Commitment', slideConfig['PI Commitment']);
  }
  
  // Filter by Value Stream/Org if configured
  if (slideConfig['Value Stream/Org']) {
    filtered = filterByField(filtered, 'Value Stream/Org', slideConfig['Value Stream/Org']);
  }
  
  // Filter by RAG if configured
  if (slideConfig['RAG']) {
    filtered = filterByField(filtered, 'RAG', slideConfig['RAG']);
  }
  
  // Filter by Allocation if configured
  if (slideConfig['Allocation']) {
    filtered = filterByField(filtered, 'Allocation', slideConfig['Allocation']);
  }
  
  console.log(`After configuration filtering: ${filtered.length} epics`);
  return filtered;
}

function filterByField(epics, fieldName, fieldConfig) {
  const originalCount = epics.length;
  
  const filtered = epics.filter(epic => {
    const fieldValue = epic[fieldName] || '';
    return fieldConfig.displayMap[fieldValue] === true;
  });
  
  console.log(`  ${fieldName}: ${originalCount} → ${filtered.length} epics`);
  return filtered;
}

function getConfiguredPortfolioInitiatives(epics, slideConfig) {
  const initiatives = new Set();
  epics.forEach(epic => {
    const initiative = epic['Portfolio Initiative'] || 'Unassigned';
    initiatives.add(initiative);
  });
  
  let initiativeList = Array.from(initiatives);
  
  if (slideConfig && slideConfig['Portfolio Initiative']) {
    const config = slideConfig['Portfolio Initiative'];
    
    initiativeList = initiativeList.filter(init => config.displayMap[init] === true);
    
    if (config.hasCustomOrder) {
      initiativeList.sort((a, b) => {
        const orderA = config.orderMap[a] || 999;
        const orderB = config.orderMap[b] || 999;
        return orderA - orderB;
      });
      console.log('Portfolio Initiatives ordered by custom configuration');
    } else {
      initiativeList.sort();
      console.log('Portfolio Initiatives ordered alphabetically');
    }
  } else {
    initiativeList.sort();
    console.log('Portfolio Initiatives ordered alphabetically (no configuration)');
  }
  
  return initiativeList;
}

function groupEpicsByField(epics, fieldName, slideConfig) {
  const groups = {};
  const fieldConfig = slideConfig ? slideConfig[fieldName] : null;
  
  epics.forEach(epic => {
    const fieldValue = epic[fieldName] || 'Unassigned';
    
    if (!fieldConfig || fieldConfig.displayMap[fieldValue] === true) {
      if (!groups[fieldValue]) {
        groups[fieldValue] = [];
      }
      groups[fieldValue].push(epic);
    }
  });
  
  let sortedKeys;
  if (fieldConfig && fieldConfig.hasCustomOrder) {
    sortedKeys = Object.keys(groups).sort((a, b) => {
      const orderA = fieldConfig.orderMap[a] || 999;
      const orderB = fieldConfig.orderMap[b] || 999;
      return orderA - orderB;
    });
    console.log(`${fieldName}: Custom ordering applied`);
  } else {
    sortedKeys = Object.keys(groups).sort();
    console.log(`${fieldName}: Alphabetical ordering applied`);
  }
  
  return { groups, sortedKeys };
}

function sortEpicsByField(epics, fieldName, slideConfig) {
  const fieldConfig = slideConfig ? slideConfig[fieldName] : null;
  
  if (fieldConfig && fieldConfig.hasCustomOrder) {
    return epics.sort((a, b) => {
      const valueA = a[fieldName] || '';
      const valueB = b[fieldName] || '';
      const orderA = fieldConfig.orderMap[valueA] || 999;
      const orderB = fieldConfig.orderMap[valueB] || 999;
      return orderA - orderB;
    });
  } else {
    return epics.sort((a, b) => {
      const valueA = a[fieldName] || '';
      const valueB = b[fieldName] || '';
      return valueA.localeCompare(valueB);
    });
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function getEpicChangeSummary(ss, epicKey, piNumber, iterationNumber) {
  const changelogSheetName = `PI ${piNumber} Changelog`;
  const changelogSheet = ss.getSheetByName(changelogSheetName);
  
  if (!changelogSheet) return null;
  
  const headers = changelogSheet.getRange(5, 1, 1, changelogSheet.getLastColumn()).getValues()[0];
  
  const lastRow = changelogSheet.getLastRow();
  if (lastRow < 6) return null;
  
  const keys = changelogSheet.getRange(6, 1, lastRow - 5, 1).getValues();
  let epicRow = -1;
  
  for (let i = 0; i < keys.length; i++) {
    if (keys[i][0] === epicKey) {
      epicRow = i + 6;
      break;
    }
  }
  
  if (epicRow === -1) return null;
  
  const iterationPrefix = `Iteration ${iterationNumber} -`;
  const changes = [];
  
  headers.forEach((header, index) => {
    if (header.startsWith(iterationPrefix) && !header.includes('NO CHANGES')) {
      const value = changelogSheet.getRange(epicRow, index + 1).getValue();
      if (value && value !== '') {
        changes.push({
          field: header.replace(iterationPrefix, '').trim(),
          value: value
        });
      }
    }
  });
  
  return changes;
}

function getFieldKey(fieldName) {
  const fieldMap = {
    'Portfolio Initiative': 'portfolioInitiative',
    'Program Initiative': 'programInitiative',
    'Allocation': 'allocation',
    'PI Commitment': 'piCommitment',
    'Value Stream/Org': 'valueStream',
    'RAG': 'rag'
  };
  return fieldMap[fieldName];
}


// ═══════════════════════════════════════════════════════════════════════════════
// HTML DIALOG FOR SLIDE GENERATION
// ═══════════════════════════════════════════════════════════════════════════════

function hasConfigurationTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Configuration');
  return sheet !== null;
}

function getSlideGeneratorHTML() {
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
          .container {
            max-width: 460px;
          }
          h2 {
            color: #1a73e8;
            margin: 0 0 8px 0;
            font-size: 22px;
            font-weight: 400;
          }
          .subtitle {
            color: #5f6368;
            font-size: 13px;
            margin-bottom: 24px;
          }
          .config-notice {
            background: #e8f0fe;
            border-left: 4px solid #1a73e8;
            padding: 12px 16px;
            margin-bottom: 20px;
            border-radius: 4px;
            font-size: 12px;
            line-height: 1.6;
          }
          .config-notice strong {
            color: #1a73e8;
            display: block;
            margin-bottom: 4px;
          }
          .config-link {
            color: #1a73e8;
            text-decoration: none;
            font-weight: 500;
          }
          .config-link:hover {
            text-decoration: underline;
          }
          .section {
            margin-bottom: 20px;
          }
          .section-title {
            font-size: 13px;
            font-weight: 500;
            color: #202124;
            margin-bottom: 8px;
          }
          .input-group {
            position: relative;
          }
          select, .text-input {
            width: 100%;
            padding: 10px 12px;
            border: 1px solid #dadce0;
            border-radius: 4px;
            font-size: 14px;
            color: #202124;
            background-color: white;
            box-sizing: border-box;
          }
          select:focus, .text-input:focus {
            outline: none;
            border-color: #1a73e8;
          }
          select:disabled {
            background-color: #f1f3f4;
            color: #80868b;
          }
          .error {
            color: #d93025;
            font-size: 12px;
            margin-top: 4px;
            display: none;
          }
          .checkbox-group {
            display: flex;
            align-items: center;
            padding: 8px 0;
            cursor: pointer;
          }
          .checkbox-group input {
            width: 18px;
            height: 18px;
            margin-right: 10px;
            cursor: pointer;
          }
          .checkbox-group label {
            font-size: 14px;
            cursor: pointer;
          }
          .checkbox-hint {
            font-size: 11px;
            color: #757575;
            margin: 4px 0 0 28px;
          }
          .multi-select-container {
            border: 1px solid #dadce0;
            border-radius: 4px;
            max-height: 120px;
            overflow-y: auto;
            padding: 8px;
            background: white;
          }
          .multi-select-item {
            display: flex;
            align-items: center;
            padding: 4px 8px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 13px;
          }
          .multi-select-item:hover {
            background: #f1f3f4;
          }
          .multi-select-item input {
            margin-right: 8px;
            width: 16px;
            height: 16px;
          }
          .multi-select-item.select-all {
            border-bottom: 1px solid #dadce0;
            margin-bottom: 4px;
            padding-bottom: 8px;
            font-weight: 500;
          }
          .info-box {
            background: #f8f9fa;
            border: 1px solid #dadce0;
            border-radius: 8px;
            padding: 16px;
            margin: 20px 0;
            font-size: 13px;
          }
          .info-box p {
            margin: 0 0 8px 0;
          }
          .info-box p:last-child {
            margin: 0;
          }
          .button-container {
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            margin-top: 24px;
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
          .primary-button:hover {
            background: #1557b0;
          }
          .primary-button:disabled {
            background: #dadce0;
            cursor: not-allowed;
          }
          .secondary-button {
            background: white;
            color: #1a73e8;
            border: 1px solid #dadce0;
            padding: 10px 24px;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
          }
          .secondary-button:hover {
            background: #f8f9fa;
          }
          .loading {
            display: none;
            text-align: center;
            padding: 40px;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #1a73e8;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Generate Governance Slides</h2>
          <p class="subtitle">Create a presentation from governance data</p>
          
          <div class="config-notice">
            <strong>📋 Field Configuration</strong>
            Fields will be displayed using the order and visibility settings from the 
            <a href="#" class="config-link" onclick="openConfigSheet(); return false;">Configuration</a> tab.
          </div>
          
          <div class="section">
            <div class="section-title">Program Increment (PI)</div>
            <div class="input-group">
              <select id="piSelect" onchange="loadIterations()">
                <option value="">Select PI...</option>
              </select>
              <div id="piError" class="error">Please select a PI</div>
            </div>
          </div>
          
          <!-- Display Options Section -->
          <div class="section">
            <div class="section-title">Display Options</div>
            <div class="input-group">
              <div class="checkbox-group" onclick="document.getElementById('showDependencies').click();">
                <input type="checkbox" id="showDependencies" checked onclick="event.stopPropagation();">
                <label for="showDependencies">Show Dependencies</label>
              </div>
              <p class="checkbox-hint">🔗 Include dependency information under each epic</p>
            </div>
            <div class="input-group" style="margin-top: 12px;">
              <label style="font-size: 13px; color: #5f6368; margin-bottom: 6px; display: block;">No Initiative Handling</label>
              <select id="noInitiativeMode">
                <option value="skip" selected>Skip header, show Value Streams only</option>
                <option value="show">Show "No Initiative" (highlighted)</option>
                <option value="hide">Hide epics with no initiative</option>
              </select>
            </div>
            <div class="input-group" style="margin-top: 8px;">
              <div class="checkbox-group" onclick="document.getElementById('hideSameTeamDeps').click();">
                <input type="checkbox" id="hideSameTeamDeps" checked onclick="event.stopPropagation();">
                <label for="hideSameTeamDeps">Hide Same-Team Dependencies</label>
              </div>
              <p class="checkbox-hint">📇 Hide dependencies where dependent team = epic's Value Stream</p>
            </div>
            <div class="input-group" style="margin-top: 8px;">
              <div class="checkbox-group" onclick="document.getElementById('highlightSchedulingRisk').click();">
                <input type="checkbox" id="highlightSchedulingRisk" checked onclick="event.stopPropagation();">
                <label for="highlightSchedulingRisk">Highlight Scheduling Risks</label>
              </div>
              <p class="checkbox-hint">⚠️ Orange highlight when dependency iteration > parent epic iteration</p>
            </div>
            <div class="input-group" style="margin-top: 8px;">
              <div class="checkbox-group" onclick="document.getElementById('showAllEpics').click();">
                <input type="checkbox" id="showAllEpics" onclick="event.stopPropagation(); toggleContinuityOptions();">
                <label for="showAllEpics">Show All Epics</label>
              </div>
              <p class="checkbox-hint">📋 Include all epics regardless of changes (Iteration 2+ only)</p>
            </div>
            <div class="input-group" style="margin-top: 8px;" id="atRiskContainer">
              <div class="checkbox-group" onclick="document.getElementById('includeAtRisk').click();">
                <input type="checkbox" id="includeAtRisk" checked onclick="event.stopPropagation();">
                <label for="includeAtRisk">Include At-Risk (no changes)</label>
              </div>
              <p class="checkbox-hint">🔴🟡 Include Amber/Red RAG epics even if they haven't changed</p>
            </div>
            <div class="input-group" style="margin-top: 8px;" id="showPreviouslyClosedContainer">
              <div class="checkbox-group" onclick="document.getElementById('showPreviouslyClosed').click();">
                <input type="checkbox" id="showPreviouslyClosed" checked onclick="event.stopPropagation();">
                <label for="showPreviouslyClosed">Show Previously Closed</label>
              </div>
              <p class="checkbox-hint">✅ Include epics closed in previous iterations for continuity</p>
            </div>
            <div class="input-group" style="margin-top: 8px;">
              <div class="checkbox-group" onclick="document.getElementById('groupByFixVersion').click();">
                <input type="checkbox" id="groupByFixVersion" checked onclick="event.stopPropagation();">
                <label for="groupByFixVersion">Group by Fix Version</label>
              </div>
              <p class="checkbox-hint">📦 Group epics by Fix Version and Iteration (cleaner display)</p>
            </div>
          </div>
          
          <!-- Iteration Section -->
          <div class="section">
            <div class="section-title">Iteration</div>
            <div class="input-group">
              <select id="iterSelect" onchange="updateInfo()" disabled>
                <option value="">Select iteration...</option>
              </select>
              <div id="iterationError" class="error">Please select an iteration</div>
            </div>
          </div>
          
          <!-- Filtering Section -->
          <div class="section">
            <div class="section-title">Filtering (Optional)</div>
            
            <!-- Portfolio Selection -->
            <div class="input-group">
              <label style="font-size: 12px; color: #5f6368; margin-bottom: 4px; display: block;">Portfolio</label>
              <select id="portfolioSelect" onchange="updateInfo()">
                <option value="">All Portfolios</option>
              </select>
            </div>
            
            <!-- Value Stream Multi-Select -->
            <div class="input-group" style="margin-top: 12px;">
              <label style="font-size: 12px; color: #5f6368; margin-bottom: 4px; display: block;">Value Streams</label>
              <div class="multi-select-container" id="valueStreamContainer">
                <div class="multi-select-item select-all" onclick="toggleAllValueStreams(event)">
                  <input type="checkbox" id="vsSelectAll" checked onclick="event.stopPropagation(); toggleAllValueStreams(event);">
                  <label for="vsSelectAll">All Value Streams</label>
                </div>
              </div>
              <p style="font-size: 11px; color: #757575; margin-top: 4px;">🏢 Filter by specific value streams</p>
            </div>
          </div>
          
          <div class="section">
            <div class="section-title">Presentation Name</div>
            <div class="input-group">
              <input type="text" id="presentationName" class="text-input" 
                     placeholder="Presentation name will be auto-generated" />
              <div id="nameError" class="error">Please enter a presentation name</div>
            </div>
          </div>
          
          <div id="infoBox" class="info-box">
            <p>Select a PI and iteration to see generation details.</p>
          </div>
          
          <div id="loading" class="loading">
            <div class="spinner"></div>
            <p style="margin-top: 16px; color: #5f6368;">Generating slides...</p>
          </div>
          
          <div class="button-container">
            <button class="secondary-button" onclick="google.script.host.close()">Cancel</button>
            <button id="generateBtn" class="primary-button" onclick="generateSlides()">Generate Slides</button>
          </div>
        </div>
        
        <script>
          let dialogData = null;
          
          window.onload = function() {
            document.getElementById('piSelect').innerHTML = '<option value="">Loading...</option>';
            document.getElementById('piSelect').disabled = true;
            
            google.script.run
              .withSuccessHandler(function(data) {
                dialogData = data;
                populatePIs(data.pis);
                populatePortfolios(data.portfolios);
                populateValueStreams(data.valueStreams || []);
                document.getElementById('piSelect').disabled = false;
                document.getElementById('portfolioSelect').disabled = false;
              })
              .withFailureHandler(function(error) {
                console.error('Error loading:', error);
                document.getElementById('piSelect').innerHTML = '<option value="">Error - retry</option>';
              })
              .getSlideDialogData();
          };
          
          function openConfigSheet() {
            google.script.run.withSuccessHandler(function(url) {
              if (url) window.open(url, '_blank');
            }).getConfigurationSheetUrl();
          }
          
          function populatePIs(pis) {
            const sel = document.getElementById('piSelect');
            sel.innerHTML = '<option value="">Select PI...</option>';
            pis.forEach(function(pi) {
              const opt = document.createElement('option');
              opt.value = pi;
              opt.textContent = 'PI ' + pi;
              sel.appendChild(opt);
            });
          }
          
          function populatePortfolios(portfolios) {
            const sel = document.getElementById('portfolioSelect');
            portfolios.forEach(function(p) {
              const opt = document.createElement('option');
              opt.value = p;
              opt.textContent = p;
              sel.appendChild(opt);
            });
          }
          
          function populateValueStreams(valueStreams) {
            const container = document.getElementById('valueStreamContainer');
            valueStreams.forEach(function(vs) {
              const div = document.createElement('div');
              div.className = 'multi-select-item';
              div.onclick = function(e) { toggleVS(e, vs); };
              
              const cb = document.createElement('input');
              cb.type = 'checkbox';
              cb.id = 'vs_' + vs.replace(/[^a-zA-Z0-9]/g, '_');
              cb.value = vs;
              cb.checked = true;
              cb.onclick = function(e) { e.stopPropagation(); updateSelectAll(); };
              
              const lbl = document.createElement('label');
              lbl.htmlFor = cb.id;
              lbl.textContent = vs;
              
              div.appendChild(cb);
              div.appendChild(lbl);
              container.appendChild(div);
            });
          }
          
          function toggleVS(event, vs) {
            const cb = document.getElementById('vs_' + vs.replace(/[^a-zA-Z0-9]/g, '_'));
            if (event.target.type !== 'checkbox') cb.checked = !cb.checked;
            updateSelectAll();
          }
          
          function toggleAllValueStreams(event) {
            const selectAll = document.getElementById('vsSelectAll');
            if (event.target.type !== 'checkbox') selectAll.checked = !selectAll.checked;
            
            const cbs = document.querySelectorAll('#valueStreamContainer input[type="checkbox"]:not(#vsSelectAll)');
            cbs.forEach(function(cb) { cb.checked = selectAll.checked; });
          }
          
          function updateSelectAll() {
            const cbs = document.querySelectorAll('#valueStreamContainer input[type="checkbox"]:not(#vsSelectAll)');
            const selectAll = document.getElementById('vsSelectAll');
            const checked = Array.from(cbs).filter(cb => cb.checked).length;
            selectAll.checked = checked === cbs.length;
            selectAll.indeterminate = checked > 0 && checked < cbs.length;
          }
          
          function getSelectedValueStreams() {
            const selectAll = document.getElementById('vsSelectAll');
            if (selectAll.checked && !selectAll.indeterminate) return null;
            const cbs = document.querySelectorAll('#valueStreamContainer input[type="checkbox"]:not(#vsSelectAll):checked');
            return Array.from(cbs).map(cb => cb.value);
          }
          
          function loadIterations() {
            const pi = parseInt(document.getElementById('piSelect').value);
            const iterSel = document.getElementById('iterSelect');
            
            if (!pi) {
              iterSel.disabled = true;
              iterSel.innerHTML = '<option value="">Select iteration...</option>';
              return;
            }
            
            const iters = dialogData && dialogData.iterationsByPI[pi] ? dialogData.iterationsByPI[pi] : [];
            iterSel.innerHTML = '<option value="">Select iteration...</option>';
            iters.forEach(function(iter) {
              const opt = document.createElement('option');
              opt.value = iter;
              opt.textContent = 'Iteration ' + iter;
              iterSel.appendChild(opt);
            });
            iterSel.disabled = false;
          }
          
          function updateInfo() {
            const pi = parseInt(document.getElementById('piSelect').value);
            const iter = parseInt(document.getElementById('iterSelect').value);
            if (!iter) return;
            
            const portfolio = document.getElementById('portfolioSelect').value;
            const today = new Date();
            const dateStr = (today.getMonth()+1).toString().padStart(2,'0') + '.' + 
                           today.getDate().toString().padStart(2,'0') + '.' + today.getFullYear();
            
            let name = 'PI ' + pi + ' - Iteration ' + iter;
            if (portfolio) name = portfolio + ' - ' + name;
            
            const infoBox = document.getElementById('infoBox');
            const nameField = document.getElementById('presentationName');
            
            if (iter <= 1) {
              nameField.value = name + ' Governance Slides - ' + dateStr;
              infoBox.innerHTML = '<p><strong>📊 Baseline Report</strong></p><p>Full governance presentation with all epics.</p>';
            } else {
              nameField.value = name + ' Governance Changes - ' + dateStr;
              infoBox.innerHTML = '<p><strong>🔄 Changes-Only Report</strong></p><p>New, modified, completed, and deferred epics.</p>';
            }
          }
          
          function toggleContinuityOptions() {
            const showAll = document.getElementById('showAllEpics').checked;
            
            const atRiskContainer = document.getElementById('atRiskContainer');
            const includeAtRisk = document.getElementById('includeAtRisk');
            const showPreviouslyClosedContainer = document.getElementById('showPreviouslyClosedContainer');
            const showPreviouslyClosed = document.getElementById('showPreviouslyClosed');
            
            if (showAll) {
              atRiskContainer.style.opacity = '0.5';
              includeAtRisk.disabled = true;
              showPreviouslyClosedContainer.style.opacity = '0.5';
              showPreviouslyClosed.disabled = true;
            } else {
              atRiskContainer.style.opacity = '1';
              includeAtRisk.disabled = false;
              showPreviouslyClosedContainer.style.opacity = '1';
              showPreviouslyClosed.disabled = false;
            }
          }

          function generateSlides() {
            const pi = document.getElementById('piSelect').value;
            const iter = document.getElementById('iterSelect').value;
            const name = document.getElementById('presentationName').value.trim();

            
            
            let valid = true;
            if (!pi) { document.getElementById('piError').style.display = 'block'; valid = false; }
            else { document.getElementById('piError').style.display = 'none'; }
            
            if (!iter) { document.getElementById('iterationError').style.display = 'block'; valid = false; }
            else { document.getElementById('iterationError').style.display = 'none'; }
            
            if (!name) { document.getElementById('nameError').style.display = 'block'; valid = false; }
            else { document.getElementById('nameError').style.display = 'none'; }
            
            const selectedVS = getSelectedValueStreams();
            if (selectedVS && selectedVS.length === 0) {
              alert('Please select at least one value stream.');
              return;
            }
            
            if (!valid) return;
            
            document.getElementById('loading').style.display = 'block';
            document.getElementById('generateBtn').disabled = true;
            
            google.script.run
              .withSuccessHandler(function(result) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('generateBtn').disabled = false;
                
                if (result && result.success && result.url) {
                  window.open(result.url, '_blank');
                  alert('✅ Slides generated successfully!');
                  google.script.host.close();
                } else {
                  alert('Error: ' + (result.error || 'Unknown error'));
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('loading').style.display = 'none';
                document.getElementById('generateBtn').disabled = false;
                alert('Error: ' + error.message);
              })
              .generateGovernanceSlides(
                parseInt(pi),
                parseInt(iter),
                name,
                document.getElementById('showDependencies').checked,
                document.getElementById('noInitiativeMode').value,
                document.getElementById('hideSameTeamDeps').checked,
                document.getElementById('portfolioSelect').value || null,
                document.getElementById('showAllEpics').checked,
                document.getElementById('includeAtRisk').checked,
                document.getElementById('highlightSchedulingRisk').checked,
                document.getElementById('showPreviouslyClosed').checked,
                selectedVS,
                document.getElementById('groupByFixVersion').checked
              );
          }
        </script>
      </body>
    </html>
  `;
}

function shouldShowAllEpics(portfolioName) {
  if (!portfolioName) return false;
  
  const showAllList = getPortfolioShowAllConfig();
  const portfolioUpper = portfolioName.toUpperCase();
  
  return showAllList.some(keyword => {
    const keywordUpper = keyword.toUpperCase();
    return portfolioUpper.startsWith(keywordUpper) || portfolioUpper.includes(keywordUpper);
  });
}

/**
 * Get the list of portfolios that should show all epics
 * Uses User Properties if saved, otherwise returns defaults
 */
function getPortfolioShowAllConfig() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const savedConfig = userProperties.getProperty('portfolioShowAllConfig');
    
    if (savedConfig) {
      const parsed = JSON.parse(savedConfig);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    }
  } catch (e) {
    console.warn('Error reading show-all config, using defaults');
  }
  
  // Default list - configurable via savePortfolioShowAllConfig()
  return [
    'RAC: REGULATORY & COMPLIANCE',
    'RCM AUTOMATION',
    'INFOSEC'
  ];
}

/**
 * Save the list of portfolios that should show all epics
 * @param {string[]} portfolioList - Array of portfolio keywords
 */
function savePortfolioShowAllConfig(portfolioList) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const listToSave = Array.isArray(portfolioList) ? portfolioList : getPortfolioShowAllConfig();
    userProperties.setProperty('portfolioShowAllConfig', JSON.stringify(listToSave));
    console.log(`Saved portfolio show-all config: ${listToSave.join(', ')}`);
    return { success: true, portfolios: listToSave };
  } catch (e) {
    console.error('Error saving portfolio show-all config:', e);
    return { success: false, error: e.toString() };
  }
}