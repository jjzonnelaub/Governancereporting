/**
 * RefreshTracker.gs
 * 
 * NEW FILE - Add this as a new .gs file in your Apps Script project
 * 
 * This file manages refresh timestamps for reports to show when data was last
 * fetched from JIRA. It tracks both cached and non-cached data fetches.
 * 
 * Usage:
 * - Call recordRefreshTimestamp() when starting a report generation
 * - Call getRefreshTimestamp() to retrieve the timestamp for display
 */

// ===== CONFIGURATION =====
const REFRESH_TRACKER_CONFIG = {
  PROPERTY_PREFIX: 'REFRESH_TIMESTAMP_',
  CACHE_EXPIRATION: 86400 // 24 hours in seconds
};

// ===== CORE TRACKING FUNCTIONS =====

/**
 * Records a refresh timestamp for a specific report
 * 
 * @param {string} reportKey - Unique identifier for the report (e.g., "PI13_Iteration2")
 * @param {Date} timestamp - The timestamp to record (defaults to current time)
 * @returns {boolean} Success status
 */
function recordRefreshTimestamp(reportKey, timestamp) {
  try {
    if (!reportKey) {
      console.error('RefreshTracker: reportKey is required');
      return false;
    }
    
    const ts = timestamp || new Date();
    const propertyKey = REFRESH_TRACKER_CONFIG.PROPERTY_PREFIX + reportKey;
    
    // Store in Script Properties for persistence across sessions
    PropertiesService.getScriptProperties().setProperty(
      propertyKey, 
      ts.toISOString()
    );
    
    console.log(`✓ Refresh timestamp recorded for ${reportKey}: ${ts.toLocaleString()}`);
    return true;
    
  } catch (error) {
    console.error(`RefreshTracker: Error recording timestamp for ${reportKey}:`, error);
    return false;
  }
}

/**
 * Retrieves the last refresh timestamp for a specific report
 * 
 * @param {string} reportKey - Unique identifier for the report (e.g., "PI13_Iteration2")
 * @returns {Date|null} The timestamp, or null if not found
 */
function getRefreshTimestamp(reportKey) {
  try {
    if (!reportKey) {
      console.error('RefreshTracker: reportKey is required');
      return null;
    }
    
    const propertyKey = REFRESH_TRACKER_CONFIG.PROPERTY_PREFIX + reportKey;
    const storedValue = PropertiesService.getScriptProperties().getProperty(propertyKey);
    
    if (!storedValue) {
      console.log(`RefreshTracker: No timestamp found for ${reportKey}`);
      return null;
    }
    
    const timestamp = new Date(storedValue);
    
    if (isNaN(timestamp.getTime())) {
      console.error(`RefreshTracker: Invalid timestamp for ${reportKey}: ${storedValue}`);
      return null;
    }
    
    return timestamp;
    
  } catch (error) {
    console.error(`RefreshTracker: Error retrieving timestamp for ${reportKey}:`, error);
    return null;
  }
}

/**
 * Gets a formatted refresh timestamp string for display
 * 
 * @param {string} reportKey - Unique identifier for the report
 * @param {string} format - Format style ('long' or 'short', defaults to 'long')
 * @returns {string} Formatted timestamp string or empty string if not found
 */
function getFormattedRefreshTimestamp(reportKey, format = 'long') {
  const timestamp = getRefreshTimestamp(reportKey);
  
  if (!timestamp) {
    return '';
  }
  
  if (format === 'short') {
    return timestamp.toLocaleDateString() + ' ' + timestamp.toLocaleTimeString();
  }
  
  // Long format (default)
  return timestamp.toLocaleString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  });
}

/**
 * Clears refresh timestamp for a specific report
 * 
 * @param {string} reportKey - Unique identifier for the report
 * @returns {boolean} Success status
 */
function clearRefreshTimestamp(reportKey) {
  try {
    if (!reportKey) {
      console.error('RefreshTracker: reportKey is required');
      return false;
    }
    
    const propertyKey = REFRESH_TRACKER_CONFIG.PROPERTY_PREFIX + reportKey;
    PropertiesService.getScriptProperties().deleteProperty(propertyKey);
    
    console.log(`✓ Refresh timestamp cleared for ${reportKey}`);
    return true;
    
  } catch (error) {
    console.error(`RefreshTracker: Error clearing timestamp for ${reportKey}:`, error);
    return false;
  }
}

/**
 * Clears all refresh timestamps (useful for cleanup)
 * 
 * @returns {number} Number of timestamps cleared
 */
function clearAllRefreshTimestamps() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProps = properties.getProperties();
    
    let clearedCount = 0;
    
    Object.keys(allProps).forEach(key => {
      if (key.startsWith(REFRESH_TRACKER_CONFIG.PROPERTY_PREFIX)) {
        properties.deleteProperty(key);
        clearedCount++;
      }
    });
    
    console.log(`✓ Cleared ${clearedCount} refresh timestamps`);
    return clearedCount;
    
  } catch (error) {
    console.error('RefreshTracker: Error clearing all timestamps:', error);
    return 0;
  }
}

/**
 * Lists all stored refresh timestamps (useful for debugging)
 * 
 * @returns {Object} Object with reportKey: timestamp pairs
 */
function listAllRefreshTimestamps() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const allProps = properties.getProperties();
    const timestamps = {};
    
    Object.keys(allProps).forEach(key => {
      if (key.startsWith(REFRESH_TRACKER_CONFIG.PROPERTY_PREFIX)) {
        const reportKey = key.replace(REFRESH_TRACKER_CONFIG.PROPERTY_PREFIX, '');
        const timestamp = new Date(allProps[key]);
        timestamps[reportKey] = timestamp.toLocaleString();
      }
    });
    
    console.log('Stored refresh timestamps:', timestamps);
    return timestamps;
    
  } catch (error) {
    console.error('RefreshTracker: Error listing timestamps:', error);
    return {};
  }
}

// ===== HELPER FUNCTIONS =====

/**
 * Checks if a refresh timestamp is stale (older than threshold)
 * 
 * @param {string} reportKey - Unique identifier for the report
 * @param {number} hoursThreshold - Number of hours to consider stale (default: 24)
 * @returns {boolean} True if timestamp is stale or missing
 */
function isRefreshTimestampStale(reportKey, hoursThreshold = 24) {
  const timestamp = getRefreshTimestamp(reportKey);
  
  if (!timestamp) {
    return true; // No timestamp means it's stale
  }
  
  const now = new Date();
  const ageInHours = (now - timestamp) / (1000 * 60 * 60);
  
  return ageInHours > hoursThreshold;
}

/**
 * Gets time since last refresh in human-readable format
 * 
 * @param {string} reportKey - Unique identifier for the report
 * @returns {string} Human-readable time since refresh (e.g., "2 hours ago")
 */
function getTimeSinceRefresh(reportKey) {
  const timestamp = getRefreshTimestamp(reportKey);
  
  if (!timestamp) {
    return 'Never refreshed';
  }
  
  const now = new Date();
  const diffMs = now - timestamp;
  const diffMins = Math.floor(diffMs / (1000 * 60));
  const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  
  if (diffMins < 1) return 'Just now';
  if (diffMins < 60) return `${diffMins} minute${diffMins > 1 ? 's' : ''} ago`;
  if (diffHours < 24) return `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
  return `${diffDays} day${diffDays > 1 ? 's' : ''} ago`;
}