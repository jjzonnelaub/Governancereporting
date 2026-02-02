/**
 * Program Governance Report - Portfolio Configuration Management
 * Handles saving and loading the custom sort order and include/exclude settings for Portfolio Initiatives.
 */

// Keys used to store configuration in User Properties
const PORTFOLIO_ORDER_KEY = 'portfolioCustomOrder';
const PORTFOLIO_INCLUDE_KEY = 'portfolioIncludeConfig';
const PORTFOLIO_SHOW_ALL_KEY = 'portfolioShowAllConfig';  // NEW: For "show all epics" setting

// ===== KEYWORD-BASED PORTFOLIO CONFIGURATION =====
// These keywords are used for "contains" matching (case-sensitive)

/**
 * Portfolio sort order based on keyword matching.
 * Portfolios containing these keywords will be sorted in this order.
 * Portfolios not matching any keyword go at the end, sorted alphabetically.
 */
const PORTFOLIO_SORT_KEYWORDS = [
  'INFOSEC',
  'RAC: REGULATORY & COMPLIANCE',
  'RCM AUTOMATION',
  'KLARA / PT COLLAB & ENGAGEMENT',
  'AI INITIATIVES',
  'DATA & ANALYTICS',
  'MMPay',
  'CLINICAL ENHANCEMENTS',
  'Enterprise PM',
  'Integration & Interop',
  'Platform Scale'
];

/**
 * Portfolio exclusion keywords.
 * Portfolios containing ANY of these keywords will be excluded from slides.
 */
const PORTFOLIO_EXCLUDE_KEYWORDS = [
  'KLO',
  'QUALITY',
  'OTHER'
];

/**
 * Value Streams to exclude (exact match).
 * Epics with these exact Value Stream values will be excluded.
 */
const VALUE_STREAM_EXCLUDE_EXACT = [
  'Xtract'
];

// ═══════════════════════════════════════════════════════════════════════════════
// NEW: PORTFOLIOS THAT SHOW ALL EPICS (bypass iteration/commitment filtering)
// ═══════════════════════════════════════════════════════════════════════════════
/**
 * Portfolios that show ALL epics for the PI regardless of:
 * - PI Commitment status
 * - Target iteration
 * - Other exclusion rules
 * 
 * Badges and highlighting are still applied based on current iteration rules.
 * This is the DEFAULT list - can be overridden via User Properties.
 */
const PORTFOLIO_SHOW_ALL_EPICS_DEFAULT = [
  'RAC: REGULATORY & COMPLIANCE',
  'RCM AUTOMATION',
  'INFOSEC'
];

/**
 * Check if a portfolio should show all epics (bypass filtering)
 * @param {string} portfolioName - The portfolio initiative name
 * @returns {boolean} - True if this portfolio should show all epics
 */
function shouldShowAllEpics(portfolioName) {
  if (!portfolioName) return false;
  
  const showAllList = getPortfolioShowAllConfig();
  const portfolioUpper = portfolioName.toUpperCase();
  
  // Check if portfolio matches any in the show-all list (case-insensitive, starts-with match)
  return showAllList.some(keyword => {
    const keywordUpper = keyword.toUpperCase();
    return portfolioUpper.startsWith(keywordUpper) || portfolioUpper.includes(keywordUpper);
  });
}

/**
 * Get the list of portfolios that should show all epics
 * @returns {string[]} Array of portfolio keywords
 */
function getPortfolioShowAllConfig() {
  const userProperties = PropertiesService.getUserProperties();
  const savedConfig = userProperties.getProperty(PORTFOLIO_SHOW_ALL_KEY);
  
  if (savedConfig) {
    try {
      const parsed = JSON.parse(savedConfig);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed;
      }
    } catch (e) {
      console.warn('Error parsing saved show-all config, using defaults');
    }
  }
  
  return PORTFOLIO_SHOW_ALL_EPICS_DEFAULT;
}

/**
 * Save the list of portfolios that should show all epics
 * @param {string[]} portfolioList - Array of portfolio keywords
 */
function savePortfolioShowAllConfig(portfolioList) {
  const userProperties = PropertiesService.getUserProperties();
  
  try {
    const listToSave = Array.isArray(portfolioList) ? portfolioList : PORTFOLIO_SHOW_ALL_EPICS_DEFAULT;
    userProperties.setProperty(PORTFOLIO_SHOW_ALL_KEY, JSON.stringify(listToSave));
    console.log(`Saved portfolio show-all config: ${listToSave.length} portfolios`);
    return { success: true, portfolios: listToSave };
  } catch (e) {
    console.error('Error saving portfolio show-all config:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * Reset show-all config to defaults
 */
function resetPortfolioShowAllConfig() {
  return savePortfolioShowAllConfig(PORTFOLIO_SHOW_ALL_EPICS_DEFAULT);
}

// ===== ORDER MANAGEMENT =====

/**
 * Retrieves the custom portfolio order from user properties.
 * @returns {string[]} The list of Portfolio Initiatives in the user's preferred order, or empty array if none saved.
 */
function getCustomPortfolioOrder() {
  const userProperties = PropertiesService.getUserProperties();
  const savedOrderJson = userProperties.getProperty(PORTFOLIO_ORDER_KEY);
  
  if (savedOrderJson) {
    try {
      return JSON.parse(savedOrderJson);
    } catch (e) {
      console.error("Error parsing saved custom portfolio order:", e);
      return [];
    }
  }
  
  return [];
}

/**
 * Saves the user's custom portfolio order.
 * @param {string[]} orderedList - The new array of portfolio names in preferred order.
 */
function saveCustomPortfolioOrder(orderedList) {
  const userProperties = PropertiesService.getUserProperties();
  const listToSave = Array.isArray(orderedList) ? orderedList : [];
  
  try {
    userProperties.setProperty(PORTFOLIO_ORDER_KEY, JSON.stringify(listToSave));
    console.log(`Saved custom portfolio order: ${listToSave.length} portfolios`);
    return { success: true };
  } catch (e) {
    console.error("Error saving custom portfolio order:", e);
    return { success: false, error: e.toString() };
  }
}

// ===== INCLUDE/EXCLUDE MANAGEMENT =====

/**
 * Retrieves the portfolio include/exclude configuration.
 * @returns {Object} Map of portfolio names to boolean (true = included, false = excluded)
 */
function getPortfolioIncludeConfig() {
  const userProperties = PropertiesService.getUserProperties();
  const savedConfigJson = userProperties.getProperty(PORTFOLIO_INCLUDE_KEY);
  
  if (savedConfigJson) {
    try {
      return JSON.parse(savedConfigJson);
    } catch (e) {
      console.error("Error parsing saved include config:", e);
      return {};
    }
  }
  
  return {};
}

/**
 * Saves the portfolio include/exclude configuration.
 * @param {Object} includeConfig - Map of portfolio names to boolean
 */
function savePortfolioIncludeConfig(includeConfig) {
  const userProperties = PropertiesService.getUserProperties();
  
  try {
    userProperties.setProperty(PORTFOLIO_INCLUDE_KEY, JSON.stringify(includeConfig));
    console.log(`Saved portfolio include config`);
    return { success: true };
  } catch (e) {
    console.error("Error saving portfolio include config:", e);
    return { success: false, error: e.toString() };
  }
}

// ===== DIALOG FUNCTIONS =====

/**
 * Shows the HTML dialog for configuring portfolio order and include/exclude.
 * This is the function called by the menu item.
 */
function showPortfolioConfigDialog() {
  const html = HtmlService.createHtmlOutputFromFile('PortfolioConfig')
      .setWidth(550)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Portfolio Slide Configuration');
}

/**
 * Fetches data required to populate the configuration dialog.
 * Scans all iteration sheets to find unique Portfolio Initiatives.
 * @returns {Object} Contains portfolios array with {name, included, showAll} objects and customOrder array
 */
function getPortfolioConfigData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Find all unique Portfolio Initiatives from iteration sheets
  const portfolioSet = new Set();
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    // Match iteration sheets: "PI X - Iteration Y"
    if (sheetName.match(/^PI \d+ - Iteration \d+$/)) {
      try {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        
        if (lastRow >= 5 && lastCol > 0) {
          // Headers are in row 4
          const headers = sheet.getRange(4, 1, 1, lastCol).getValues()[0];
          const portfolioColIndex = headers.indexOf('Portfolio Initiative');
          
          if (portfolioColIndex !== -1) {
            // Data starts in row 5
            const dataRows = lastRow - 4;
            if (dataRows > 0) {
              const portfolioValues = sheet.getRange(5, portfolioColIndex + 1, dataRows, 1).getValues();
              portfolioValues.forEach(row => {
                const value = row[0];
                if (value && value.toString().trim() !== '') {
                  portfolioSet.add(value.toString().trim());
                }
              });
            }
          }
        }
      } catch (e) {
        console.log(`Error reading sheet ${sheetName}: ${e.message}`);
      }
    }
  });
  
  // Get saved configurations
  const savedOrder = getCustomPortfolioOrder();
  const savedIncludeConfig = getPortfolioIncludeConfig();
  const showAllList = getPortfolioShowAllConfig();
  
  // Convert to array and build portfolio objects
  const allPortfolios = Array.from(portfolioSet);
  
  // Sort portfolios: use saved order first, then alphabetical for any new ones
  const orderedPortfolios = [];
  const unorderedPortfolios = [];
  
  allPortfolios.forEach(name => {
    const orderIndex = savedOrder.indexOf(name);
    if (orderIndex !== -1) {
      orderedPortfolios.push({ name, orderIndex });
    } else {
      unorderedPortfolios.push(name);
    }
  });
  
  // Sort ordered portfolios by their saved index
  orderedPortfolios.sort((a, b) => a.orderIndex - b.orderIndex);
  
  // Sort unordered portfolios alphabetically
  unorderedPortfolios.sort();
  
  // Combine: ordered first, then unordered
  const finalOrder = [
    ...orderedPortfolios.map(p => p.name),
    ...unorderedPortfolios
  ];
  
  // Build result with included status and showAll status
  const portfolios = finalOrder.map(name => {
    // Check if explicitly set in saved config, otherwise check keyword exclusions
    let included = true;
    if (savedIncludeConfig.hasOwnProperty(name)) {
      included = savedIncludeConfig[name];
    } else {
      // Check if excluded by keyword
      included = !PORTFOLIO_EXCLUDE_KEYWORDS.some(keyword => 
        name.toUpperCase().includes(keyword.toUpperCase())
      );
    }
    
    // Check if this portfolio shows all epics
    const showAll = shouldShowAllEpics(name);
    
    return { name, included, showAll };
  });
  
  return {
    portfolios: portfolios,
    customOrder: finalOrder,
    showAllList: showAllList
  };
}

// ===== SORTING AND FILTERING =====

/**
 * Gets the filtered and sorted list of portfolios for slide generation.
 * Applies keyword-based sorting and exclusion, respecting any user customizations.
 * @param {string[]} portfolioNames - Array of portfolio names to sort/filter
 * @returns {string[]} Sorted and filtered portfolio names
 */
function getFilteredAndSortedPortfolios(portfolioNames) {
  if (!portfolioNames || portfolioNames.length === 0) {
    return [];
  }
  
  // Get saved configurations
  const savedOrder = getCustomPortfolioOrder();
  const savedIncludeConfig = getPortfolioIncludeConfig();
  
  // Filter out excluded portfolios
  const includedPortfolios = portfolioNames.filter(name => {
    // Check if explicitly set in saved config
    if (savedIncludeConfig.hasOwnProperty(name)) {
      return savedIncludeConfig[name];
    }
    
    // Check if excluded by keyword
    const isExcludedByKeyword = PORTFOLIO_EXCLUDE_KEYWORDS.some(keyword =>
      name.toUpperCase().includes(keyword.toUpperCase())
    );
    
    return !isExcludedByKeyword;
  });
  
  // Sort portfolios
  // Priority 1: Saved custom order
  // Priority 2: Keyword sort order
  // Priority 3: Alphabetical
  
  const sortedPortfolios = [...includedPortfolios].sort((a, b) => {
    // Check saved custom order first
    const aCustomIndex = savedOrder.indexOf(a);
    const bCustomIndex = savedOrder.indexOf(b);
    
    if (aCustomIndex !== -1 && bCustomIndex !== -1) {
      return aCustomIndex - bCustomIndex;
    }
    if (aCustomIndex !== -1) return -1;
    if (bCustomIndex !== -1) return 1;
    
    // Then check keyword sort order
    const aKeywordIndex = PORTFOLIO_SORT_KEYWORDS.findIndex(keyword =>
      a.toUpperCase().includes(keyword.toUpperCase())
    );
    const bKeywordIndex = PORTFOLIO_SORT_KEYWORDS.findIndex(keyword =>
      b.toUpperCase().includes(keyword.toUpperCase())
    );
    
    if (aKeywordIndex !== -1 && bKeywordIndex !== -1) {
      return aKeywordIndex - bKeywordIndex;
    }
    if (aKeywordIndex !== -1) return -1;
    if (bKeywordIndex !== -1) return 1;
    
    // Finally, alphabetical
    return a.localeCompare(b);
  });
  
  return sortedPortfolios;
}

/**
 * Resets all configurations to defaults.
 */
function resetAllPortfolioConfig() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(PORTFOLIO_ORDER_KEY);
  userProperties.deleteProperty(PORTFOLIO_INCLUDE_KEY);
  userProperties.deleteProperty(PORTFOLIO_SHOW_ALL_KEY);
  console.log('Reset all portfolio configurations to defaults');
  return { success: true };
}