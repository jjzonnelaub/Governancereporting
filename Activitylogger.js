/**
 * ============================================================================
 * ACTIVITY LOGGER
 * ============================================================================
 * Automatically logs function execution to an "Activity Log" sheet tab
 * 
 * Usage:
 * 1. Wrap any function you want to track with logActivity()
 * 2. All executions are automatically logged with metadata
 * 
 * Example:
 *   function myFunction(pi, iteration) {
 *     return logActivity('myFunction', () => {
 *       // your existing code here
 *       return result;
 *     }, { pi, iteration });
 *   }
 */

// ===== CONFIGURATION =====
const ACTIVITY_LOG_CONFIG = {
  SHEET_NAME: 'Activity Log',
  MAX_ROWS: 5000,  // Keep last 5000 entries
  RETENTION_DAYS: 90,  // Keep logs for 90 days
  HEADERS: [
    'Timestamp',
    'Function Name',
    'Status',
    'Duration (seconds)',
    'User Email',
    'Parameters',
    'Result Summary',
    'Error Message',
    'Sheet Context',
    'Memory Used (MB)',
    'Trigger Type',
    'Session ID'
  ]
};

// ===== MAIN LOGGING FUNCTION =====

/**
 * Wraps a function execution with automatic activity logging
 * 
 * @param {string} functionName - Name of the function being executed
 * @param {Function} fn - The function to execute
 * @param {Object} params - Parameters passed to the function (optional)
 * @param {Object} options - Additional logging options (optional)
 * @returns {*} - Returns the result of the wrapped function
 */
function logActivity(functionName, fn, params = {}, options = {}) {
  const startTime = new Date();
  const sessionId = Utilities.getUuid().substring(0, 8);
  const userEmail = Session.getActiveUser().getEmail() || 'Unknown';
  const triggerType = detectTriggerType();
  
  let status = 'SUCCESS';
  let errorMessage = '';
  let result = null;
  let resultSummary = '';
  
  try {
    // Execute the wrapped function
    result = fn();
    
    // Generate result summary
    if (result !== null && result !== undefined) {
      if (typeof result === 'object') {
        if (Array.isArray(result)) {
          resultSummary = `Array with ${result.length} items`;
        } else {
          resultSummary = `Object with keys: ${Object.keys(result).slice(0, 3).join(', ')}`;
        }
      } else {
        resultSummary = String(result).substring(0, 100);
      }
    }
    
  } catch (error) {
    status = 'ERROR';
    errorMessage = error.toString().substring(0, 500);
    throw error; // Re-throw to maintain original error handling
    
  } finally {
    // Calculate duration
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    // Log the activity
    try {
      writeActivityLog({
        timestamp: startTime,
        functionName: functionName,
        status: status,
        duration: duration,
        userEmail: userEmail,
        parameters: JSON.stringify(params).substring(0, 500),
        resultSummary: resultSummary,
        errorMessage: errorMessage,
        sheetContext: getSheetContext(),
        memoryUsed: getMemoryUsage(),
        triggerType: triggerType,
        sessionId: sessionId,
        ...options
      });
    } catch (logError) {
      // Don't let logging errors break the main function
      console.error('Failed to write activity log:', logError);
    }
  }
  
  return result;
}

// ===== HELPER FUNCTIONS =====

/**
 * Writes a log entry to the Activity Log sheet
 */
function writeActivityLog(logEntry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(ACTIVITY_LOG_CONFIG.SHEET_NAME);
  
  // Create sheet if it doesn't exist
  if (!logSheet) {
    logSheet = createActivityLogSheet(ss);
  }
  
  // Prepare row data
  const rowData = [
    logEntry.timestamp,
    logEntry.functionName,
    logEntry.status,
    logEntry.duration.toFixed(3),
    logEntry.userEmail,
    logEntry.parameters,
    logEntry.resultSummary,
    logEntry.errorMessage,
    logEntry.sheetContext,
    logEntry.memoryUsed,
    logEntry.triggerType,
    logEntry.sessionId
  ];
  
  // Insert at row 2 (after header)
  logSheet.insertRowAfter(1);
  logSheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
  
  // Apply formatting to the new row
  formatLogRow(logSheet, 2, logEntry.status);
  
  // Cleanup old entries if needed
  maintainLogSize(logSheet);
}

/**
 * Creates the Activity Log sheet with headers and formatting
 */
function createActivityLogSheet(ss) {
  const logSheet = ss.insertSheet(ACTIVITY_LOG_CONFIG.SHEET_NAME);
  
  // Set headers
  const headerRange = logSheet.getRange(1, 1, 1, ACTIVITY_LOG_CONFIG.HEADERS.length);
  headerRange.setValues([ACTIVITY_LOG_CONFIG.HEADERS]);
  
  // Format headers
  headerRange
    .setBackground('#ffffff')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  
  // Set column widths
  const widths = [150, 200, 80, 100, 180, 250, 200, 300, 150, 100, 100, 100];
  widths.forEach((width, index) => {
    logSheet.setColumnWidth(index + 1, width);
  });
  
  // Freeze header row
  logSheet.setFrozenRows(1);
  
  return logSheet;
}

/**
 * Formats a log row based on status
 */
function formatLogRow(sheet, row, status) {
  const range = sheet.getRange(row, 1, 1, ACTIVITY_LOG_CONFIG.HEADERS.length);
  
  // Apply color based on status
  if (status === 'ERROR') {
    range.setBackground('#FFEBEE'); // Light red
  } else if (status === 'SUCCESS') {
    range.setBackground('#E8F5E9'); // Light green
  } else {
    range.setBackground('#FFF9C4'); // Light yellow (warning)
  }
  
  // Format duration column to highlight long operations
  const durationCell = sheet.getRange(row, 4);
  const duration = parseFloat(durationCell.getValue());
  if (duration > 30) {
    durationCell.setFontColor('#D32F2F'); // Red for operations > 30s
    durationCell.setFontWeight('bold');
  }
}

/**
 * Maintains log size by removing old entries
 */
function maintainLogSize(logSheet) {
  const lastRow = logSheet.getLastRow();
  
  // Remove rows exceeding MAX_ROWS
  if (lastRow > ACTIVITY_LOG_CONFIG.MAX_ROWS + 1) {
    const rowsToDelete = lastRow - ACTIVITY_LOG_CONFIG.MAX_ROWS - 1;
    logSheet.deleteRows(ACTIVITY_LOG_CONFIG.MAX_ROWS + 2, rowsToDelete);
  }
  
  // Remove entries older than RETENTION_DAYS
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - ACTIVITY_LOG_CONFIG.RETENTION_DAYS);
  
  const timestamps = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let deleteFromRow = null;
  
  for (let i = timestamps.length - 1; i >= 0; i--) {
    if (timestamps[i][0] < cutoffDate) {
      deleteFromRow = i + 2; // +2 because array is 0-indexed and row 1 is header
      break;
    }
  }
  
  if (deleteFromRow && deleteFromRow <= lastRow) {
    logSheet.deleteRows(deleteFromRow, lastRow - deleteFromRow + 1);
  }
}

/**
 * Gets the current sheet context
 */
function getSheetContext() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    return sheet ? sheet.getName() : 'No active sheet';
  } catch (e) {
    return 'Unknown';
  }
}

/**
 * Estimates memory usage
 */
function getMemoryUsage() {
  try {
    const cache = CacheService.getScriptCache();
    // This is an approximation - Google Apps Script doesn't provide direct memory usage
    return 'N/A';
  } catch (e) {
    return 'Unknown';
  }
}

/**
 * Detects how the function was triggered
 */
function detectTriggerType() {
  try {
    const trigger = ScriptApp.getScriptTriggers()[0];
    if (!trigger) {
      return 'Manual';
    }
    
    const eventType = trigger.getEventType();
    switch (eventType) {
      case ScriptApp.EventType.CLOCK:
        return 'Time-driven';
      case ScriptApp.EventType.ON_OPEN:
        return 'On Open';
      case ScriptApp.EventType.ON_EDIT:
        return 'On Edit';
      case ScriptApp.EventType.ON_FORM_SUBMIT:
        return 'Form Submit';
      default:
        return 'Manual';
    }
  } catch (e) {
    return 'Manual';
  }
}

// ===== ANALYSIS & REPORTING FUNCTIONS =====

/**
 * Gets activity statistics for a date range
 */
function getActivityStats(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(ACTIVITY_LOG_CONFIG.SHEET_NAME);
  
  if (!logSheet) {
    return { error: 'Activity Log sheet not found' };
  }
  
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) {
    return { error: 'No activity data available' };
  }
  
  const data = logSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  const stats = {
    totalExecutions: 0,
    successCount: 0,
    errorCount: 0,
    totalDuration: 0,
    averageDuration: 0,
    functionCounts: {},
    slowestFunctions: []
  };
  
  data.forEach(row => {
    const timestamp = row[0];
    const functionName = row[1];
    const status = row[2];
    const duration = parseFloat(row[3]);
    
    if (timestamp >= startDate && timestamp <= endDate) {
      stats.totalExecutions++;
      stats.totalDuration += duration;
      
      if (status === 'SUCCESS') stats.successCount++;
      if (status === 'ERROR') stats.errorCount++;
      
      stats.functionCounts[functionName] = (stats.functionCounts[functionName] || 0) + 1;
    }
  });
  
  stats.averageDuration = stats.totalExecutions > 0 
    ? stats.totalDuration / stats.totalExecutions 
    : 0;
  
  return stats;
}

/**
 * Creates a summary report of recent activity
 */
function createActivitySummaryReport() {
  const today = new Date();
  const sevenDaysAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  const stats = getActivityStats(sevenDaysAgo, today);
  
  if (stats.error) {
    return stats.error;
  }
  
  const report = `
Activity Summary (Last 7 Days)
================================
Total Executions: ${stats.totalExecutions}
Successful: ${stats.successCount}
Errors: ${stats.errorCount}
Total Duration: ${stats.totalDuration.toFixed(2)}s
Average Duration: ${stats.averageDuration.toFixed(2)}s

Most Used Functions:
${Object.entries(stats.functionCounts)
  .sort((a, b) => b[1] - a[1])
  .slice(0, 5)
  .map(([name, count]) => `  ${name}: ${count} times`)
  .join('\n')}
  `;
  
  return report;
}

/**
 * Menu function to view activity summary
 */
function showActivitySummary() {
  const report = createActivitySummaryReport();
  const ui = SpreadsheetApp.getUi();
  ui.alert('Activity Summary', report, ui.ButtonSet.OK);
}

/**
 * Menu function to clear old logs
 */
function clearOldActivityLogs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Old Logs',
    `This will remove all activity logs older than ${ACTIVITY_LOG_CONFIG.RETENTION_DAYS} days. Continue?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(ACTIVITY_LOG_CONFIG.SHEET_NAME);
    
    if (logSheet) {
      maintainLogSize(logSheet);
      ui.alert('Success', 'Old logs have been cleared.', ui.ButtonSet.OK);
    }
  }
}