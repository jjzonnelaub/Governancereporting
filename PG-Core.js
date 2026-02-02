function reauthorize() {
  SpreadsheetApp.getUi().alert('Authorization successful!');
}
// ===== CONFIGURATION & SETUP =====

// Projects to exclude from governance reporting
const EXCLUDED_PROJECTS = ['UX', 'AUTOMATION', 'OVERHEAD'];

// Available value streams
const VALUE_STREAMS = [
  'AIMM',
  'Cloud Foundation Services',
  'Cloud Operations',
  'Data Platform Engineering',
  'EMA Clinical',
  'EMA RaC',
  'Fusion and Conversions',
  'MMGI',
  'MMGI-Cloud',
  'MMPM',
  'Other',
  'Patient Collaboration',
  'Platform Engineering',
  'RCM Genie',
  'Security Engineering',
  'Shared Services Platform', 
  'Xtract'

];

// Field mappings for JIRA custom fields
const FIELD_MAPPINGS = {
  summary: 'summary',
  status: 'status',
  storyPoints: 'customfield_10037',
  storyPointEstimate: 'customfield_10016',
  epicLink: 'customfield_10014',
  epicName: 'customfield_10011',
  piObjective: 'customfield_10047',
  programIncrement: 'customfield_10113',
  valueStream: 'customfield_10046',
  piCommitment: 'customfield_10063',
  portfolioInitiative: 'customfield_10049',
  programInitiative: 'customfield_10050',
  scrumTeam: 'customfield_10040',
  piTargetIteration: 'customfield_10061',
  allocation: 'customfield_10043',
  piObjectiveStatus: 'customfield_10224', 
  businessValue: 'customfield_10071',
  actualValue: 'customfield_10060',
  rag: 'customfield_10068',
  ragNote: 'customfield_10067',
  resolution: 'resolution',
  fixVersions: 'fixVersions',
  benefitHypothesis: 'customfield_10044',
  acceptanceCriteria: 'customfield_10045',
  iterationStart: 'customfield_10069',
  iterationEnd: 'customfield_10070',
  dependsOnValuestream: 'customfield_10114',
  dependsOnTeam: 'customfield_10120',
  featurePoints: 'customfield_10066'
};

// ===== AI LABEL DETECTION CONFIGURATION =====
// Define which labels indicate Governance items
const GOVERNANCE_LABEL_PATTERNS = [
  'tech_initative',      
  'tech_initiative',     
  'governance',
  'strategic',
  'board_item',
  'executive_priority'
];

// Define which labels indicate Momentum items  
const MOMENTUM_LABEL_PATTERNS = [
  'momentum',
  'momentum_2025',
  'momentum_25'
];
const INFOSEC_SUBSECTION_LABELS = {
  'infra-vulns': 'Infrastructure Vulnerabilities',
  'appsec-vulns': 'AppSec Vulnerabilities', 
  'product-infosec': 'Product InfoSec'
};
function getInfosecSubsection(labels) {
  if (!Array.isArray(labels) || labels.length === 0) {
    return '';
  }
  
  const matchedLabels = [];
  let primarySubsection = '';
  
  // Check labels in priority order (as defined in INFOSEC_SUBSECTION_LABELS)
  for (const [labelPattern, subsectionName] of Object.entries(INFOSEC_SUBSECTION_LABELS)) {
    const hasLabel = labels.some(label => 
      label.toLowerCase() === labelPattern.toLowerCase()
    );
    if (hasLabel) {
      matchedLabels.push(labelPattern);
      if (!primarySubsection) {
        primarySubsection = subsectionName;
      }
    }
  }
  
  // Log warning if multiple InfoSec labels found (data quality issue)
  if (matchedLabels.length > 1) {
    console.warn(`⚠️ DATA QUALITY: Multiple InfoSec subsection labels found: [${matchedLabels.join(', ')}]. Using "${primarySubsection}" based on priority order.`);
  }
  
  return primarySubsection;
}

// ===== JIRA CONFIGURATION =====


function getJiraConfig() {
  const properties = PropertiesService.getScriptProperties();
  
  const config = {
    baseUrl: properties.getProperty('JIRA_BASE_URL'),
    email: properties.getProperty('JIRA_EMAIL'),
    apiToken: properties.getProperty('JIRA_API_TOKEN'),
    apiVersion: properties.getProperty('JIRA_API_VERSION') || '3',
    maxResults: parseInt(properties.getProperty('JIRA_MAX_RESULTS') || '100'),
    rateLimitDelay: parseInt(properties.getProperty('JIRA_RATE_LIMIT_DELAY') || '1000'),
    retryAttempts: parseInt(properties.getProperty('JIRA_RETRY_ATTEMPTS') || '3'),
    timeout: parseInt(properties.getProperty('JIRA_TIMEOUT') || '30000')
  };
  
  if (!config.baseUrl || !config.email || !config.apiToken) {
    throw new Error('JIRA configuration not found. Please run setupJiraConfiguration() first.');
  }
  
  config.baseUrl = config.baseUrl.replace(/\/$/, '');
  
  return config;
}

function getJiraHeaders() {
  const config = getJiraConfig();
  return {
    'Authorization': 'Basic ' + Utilities.base64Encode(config.email + ':' + config.apiToken),
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  };
}

function getJiraHeaders() {
  const config = getJiraConfig();
  return {
    'Authorization': 'Basic ' + Utilities.base64Encode(config.email + ':' + config.apiToken),
    'Content-Type': 'application/json',
    'Accept': 'application/json'
  };
}

// ===== MAIN ANALYSIS FUNCTION =====

function refreshValueStreamData(piNumber, valueStream) {
  const ui = SpreadsheetApp.getUi();
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    console.log(`Starting refresh for PI ${piNumber} - ${valueStream}`);
    const programIncrement = `PI ${piNumber}`;
    
    ss.toast(`Step 1/3: Fetching epics for ${valueStream}...`, '⏳ Refreshing', 20);
    const epics = fetchEpicsForPIOptimized(programIncrement, [valueStream]);
    
    if (epics.length === 0) {
        ss.toast(`No epics found for ${valueStream}.`, 'ℹ️ Notice', 5);
        const sheetName = `${programIncrement} - Governance`;
        writeGovernanceDataToSheet([], programIncrement, valueStream, sheetName, true);
        return;
    }
    
    console.log(`Found ${epics.length} epics`);
    
    ss.toast(`Step 2/3: Analyzing ${epics.length} epics...`, '⏳ Refreshing', 20);
    const epicDataMap = fetchChildDataForEpicsOptimized(epics);
    
    ss.toast('Step 3/3: Writing data to sheet...', '⏳ Refreshing', 10);
    const processedData = processGovernanceData(epics, epicDataMap);
    
    const sheetName = `${programIncrement} - Governance`;
    // Pass true for clearExisting to replace all data when refreshing
    writeGovernanceDataToSheet(processedData, programIncrement, valueStream, sheetName, true);
    
    const duration = (Date.now() - startTime) / 1000;
    ss.toast(
      `✅ Successfully refreshed ${valueStream} in ${duration.toFixed(1)} seconds!`,
      '✅ Refresh Complete',
      10
    );
    
  } catch (error) {
    console.error('Refresh error:', error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}

// ===== JIRA API FUNCTIONS =====

function fetchEpicsForPIOptimized(programIncrement, valueStream = null, includePrePlanning = false) {
  const excludedProjectsList = EXCLUDED_PROJECTS.map(p => `"${p}"`).join(', ');
  let allEpics = [];
  const seenKeys = new Set();
  
  let singleValueStream = null;
  if (Array.isArray(valueStream) && valueStream.length === 1) {
    singleValueStream = valueStream[0];
  } else if (typeof valueStream === 'string') {
    singleValueStream = valueStream;
  }
  
  console.log(`\n========================================`);
  console.log(`Fetching epics for PI: ${programIncrement}`);
  console.log(`Value Stream: ${singleValueStream || 'All'}`);
  console.log(`Include Pre-Planning: ${includePrePlanning}`);
  console.log(`========================================\n`);
  
  const piFieldId = FIELD_MAPPINGS.programIncrement.replace('customfield_', '');
  const vsFieldId = FIELD_MAPPINGS.valueStream.replace('customfield_', '');
  
  // Helper function to fetch and deduplicate epics
  const fetchAndAddEpics = (label, epics) => {
    let newCount = 0;
    epics.forEach(epic => {
      if (!seenKeys.has(epic.key)) {
        seenKeys.add(epic.key);
        allEpics.push(epic);
        newCount++;
      }
    });
    console.log(`${label}: Added ${newCount} unique epics (${epics.length - newCount} duplicates skipped)`);
  };
  
  // Build JQL for each query type
  const buildJQL = (vsName, piValue) => {
    const vsFilter = `cf[${vsFieldId}] = "${vsName}"`;
    const piFilterStr = `cf[${piFieldId}] = "${piValue}"`;
    
    // MMPM and EMA Clinical don't have project exclusions
    if (vsName === 'MMPM' || vsName === 'EMA Clinical') {
      return `issuetype = Epic AND ${piFilterStr} AND ${vsFilter}`;
    } else {
      return `issuetype = Epic AND ${piFilterStr} AND ${vsFilter} AND project NOT IN (${excludedProjectsList})`;
    }
  };
  
  // Handle MMPM value stream
  if (singleValueStream === 'MMPM') {
    console.log(`\n--- Fetching MMPM Value Stream ---`);
    
    // Query 1: PI X epics
    console.log(`\n[1/2] Fetching "${programIncrement}" epics...`);
    const piJql = buildJQL('MMPM', programIncrement);
    console.log(`JQL: ${piJql}`);
    try {
      const piEpics = fetchEpicsWithPagination(piJql, 100);
      fetchAndAddEpics('PI epics', piEpics);
    } catch (e) {
      console.error(`Error fetching MMPM PI epics: ${e.message}`);
    }
    
    // Query 2: Pre-Planning epics (if requested)
    if (includePrePlanning) {
      console.log(`\n[2/2] Fetching "Pre-Planning" epics...`);
      const prePlanJql = buildJQL('MMPM', 'Pre-Planning');
      console.log(`JQL: ${prePlanJql}`);
      try {
        const prePlanEpics = fetchEpicsWithPagination(prePlanJql, 100);
        fetchAndAddEpics('Pre-Planning epics', prePlanEpics);
      } catch (e) {
        console.error(`Error fetching MMPM Pre-Planning epics: ${e.message}`);
      }
    } else {
      console.log(`\n[2/2] Skipping Pre-Planning (not requested)`);
    }
  } 
  // Handle EMA Clinical value stream
  else if (singleValueStream === 'EMA Clinical') {
    console.log(`\n--- Fetching EMA Clinical Value Stream ---`);
    
    // Query 1: PI X epics
    console.log(`\n[1/2] Fetching "${programIncrement}" epics...`);
    const piJql = buildJQL('EMA Clinical', programIncrement);
    console.log(`JQL: ${piJql}`);
    try {
      const piEpics = fetchEpicsWithPagination(piJql, 100);
      fetchAndAddEpics('PI epics', piEpics);
    } catch (e) {
      console.error(`Error fetching EMA Clinical PI epics: ${e.message}`);
    }
    
    // Query 2: Pre-Planning epics (if requested)
    if (includePrePlanning) {
      console.log(`\n[2/2] Fetching "Pre-Planning" epics...`);
      const prePlanJql = buildJQL('EMA Clinical', 'Pre-Planning');
      console.log(`JQL: ${prePlanJql}`);
      try {
        const prePlanEpics = fetchEpicsWithPagination(prePlanJql, 100);
        fetchAndAddEpics('Pre-Planning epics', prePlanEpics);
      } catch (e) {
        console.error(`Error fetching EMA Clinical Pre-Planning epics: ${e.message}`);
      }
    } else {
      console.log(`\n[2/2] Skipping Pre-Planning (not requested)`);
    }
  } 
  // ═══════════════════════════════════════════════════════════
  // ✅ HANDLE MMGI-CLOUD VALUE STREAM (Project HYD)
  // ═══════════════════════════════════════════════════════════
  else if (singleValueStream === 'MMGI-Cloud') {
    console.log(`\n--- Fetching MMGI-Cloud Value Stream (Project HYD) ---`);
    
    // Query 1: PI X epics using project filter instead of Value Stream field
    console.log(`\n[1/2] Fetching "${programIncrement}" epics from project HYD...`);
    const piJql = `issuetype = Epic AND cf[${piFieldId}] = "${programIncrement}" AND project = HYD`;
    console.log(`JQL: ${piJql}`);
    try {
      const piEpics = fetchEpicsWithPagination(piJql, 100);
      fetchAndAddEpics('PI epics (HYD project)', piEpics);
    } catch (e) {
      console.error(`Error fetching MMGI-Cloud PI epics: ${e.message}`);
    }
    
    // Query 2: Pre-Planning epics (if requested)
    if (includePrePlanning) {
      console.log(`\n[2/2] Fetching "Pre-Planning" epics from project HYD...`);
      const prePlanJql = `issuetype = Epic AND cf[${piFieldId}] = "Pre-Planning" AND project = HYD`;
      console.log(`JQL: ${prePlanJql}`);
      try {
        const prePlanEpics = fetchEpicsWithPagination(prePlanJql, 100);
        fetchAndAddEpics('Pre-Planning epics (HYD project)', prePlanEpics);
      } catch (e) {
        console.error(`Error fetching MMGI-Cloud Pre-Planning epics: ${e.message}`);
      }
    } else {
      console.log(`\n[2/2] Skipping Pre-Planning (not requested)`);
    }
  }
  // ═══════════════════════════════════════════════════════════
  // END MMGI-CLOUD HANDLING
  // ═══════════════════════════════════════════════════════════
  // Handle single specific value stream
  else if (singleValueStream) {
    console.log(`\n--- Fetching ${singleValueStream} Value Stream ---`);
    
    // Query 1: PI X epics
    console.log(`\n[1/2] Fetching "${programIncrement}" epics...`);
    const piJql = buildJQL(singleValueStream, programIncrement);
    console.log(`JQL: ${piJql}`);
    try {
      const piEpics = fetchEpicsWithPagination(piJql, 100);
      fetchAndAddEpics('PI epics', piEpics);
    } catch (e) {
      console.error(`Error fetching ${singleValueStream} PI epics: ${e.message}`);
    }
    
    // Query 2: Pre-Planning epics (if requested)
    if (includePrePlanning) {
      console.log(`\n[2/2] Fetching "Pre-Planning" epics...`);
      const prePlanJql = buildJQL(singleValueStream, 'Pre-Planning');
      console.log(`JQL: ${prePlanJql}`);
      try {
        const prePlanEpics = fetchEpicsWithPagination(prePlanJql, 100);
        fetchAndAddEpics('Pre-Planning epics', prePlanEpics);
      } catch (e) {
        console.error(`Error fetching ${singleValueStream} Pre-Planning epics: ${e.message}`);
      }
    } else {
      console.log(`\n[2/2] Skipping Pre-Planning (not requested)`);
    }
  } 
  // Handle all value streams or multiple value streams
  else {
    console.log(`\n--- Fetching All/Multiple Value Streams ---`);
    
    const valueStreams = Array.isArray(valueStream) && valueStream.length > 0 
      ? valueStream 
      : ['MMPM', 'EMA Clinical', 'XTRAC', 'RCM Genie'];
    
    valueStreams.forEach((vs, index) => {
      console.log(`\n--- Value Stream ${index + 1}/${valueStreams.length}: ${vs} ---`);
      
      // Query 1: PI X epics
      console.log(`\n[1/2] Fetching "${programIncrement}" epics...`);
      const piJql = buildJQL(vs, programIncrement);
      console.log(`JQL: ${piJql}`);
      try {
        const piEpics = fetchEpicsWithPagination(piJql, 100);
        fetchAndAddEpics(`${vs} PI epics`, piEpics);
      } catch (e) {
        console.error(`Error fetching ${vs} PI epics: ${e.message}`);
      }
      
      // Query 2: Pre-Planning epics (if requested)
      if (includePrePlanning) {
        console.log(`\n[2/2] Fetching "Pre-Planning" epics...`);
        const prePlanJql = buildJQL(vs, 'Pre-Planning');
        console.log(`JQL: ${prePlanJql}`);
        try {
          const prePlanEpics = fetchEpicsWithPagination(prePlanJql, 100);
          fetchAndAddEpics(`${vs} Pre-Planning epics`, prePlanEpics);
        } catch (e) {
          console.error(`Error fetching ${vs} Pre-Planning epics: ${e.message}`);
        }
      } else {
        console.log(`\n[2/2] Skipping Pre-Planning (not requested)`);
      }
    });
  }
  
  console.log(`\n========================================`);
  console.log(`✅ FETCH COMPLETE`);
  console.log(`Total unique epics: ${allEpics.length}`);
  console.log(`========================================\n`);
  
  // ═══════════════════════════════════════════════════════════
  // ✅ POST-PROCESSING: Override Value Stream for MMGI-Cloud
  // ALL epics from HYD project should be tagged as MMGI-Cloud
  // regardless of their Value Stream field in JIRA
  // ═══════════════════════════════════════════════════════════
  if (singleValueStream === 'MMGI-Cloud') {
    console.log(`\n--- Post-Processing: Overriding Value Stream for ALL HYD epics ---`);
    let overrideCount = 0;
    
    allEpics.forEach(epic => {
      // Check if this is a HYD project epic (key starts with "HYD-")
      if (epic.key && epic.key.startsWith('HYD-')) {
        const originalVS = epic.valueStream;
        epic.valueStream = 'MMGI-Cloud';
        console.log(`  ${epic.key}: "${originalVS}" → "MMGI-Cloud"`);
        overrideCount++;
      }
    });
    
    console.log(`✓ Overridden ${overrideCount} HYD epic(s) to Value Stream = "MMGI-Cloud"`);
    
    // Also log any non-HYD epics that were fetched (shouldn't happen, but good to know)
    const nonHydEpics = allEpics.filter(e => !e.key.startsWith('HYD-'));
    if (nonHydEpics.length > 0) {
      console.warn(`⚠️ Warning: Found ${nonHydEpics.length} non-HYD epics in MMGI-Cloud query:`);
      nonHydEpics.forEach(e => console.warn(`  - ${e.key}: Value Stream = "${e.valueStream}"`));
    }
    console.log(`========================================\n`);
  }
  // ═══════════════════════════════════════════════════════════
  
  return allEpics;
}

function searchJiraIssues(jql, expand = null) {
  const initialAttempt = searchJiraIssuesLimited(jql, 100);
  
  if (initialAttempt.length < 100) {
    console.log(`Found ${initialAttempt.length} issues, which is under the 100 limit. Returning results directly.`);
    return initialAttempt;
  }

  console.warn(`Query returned 100 results. This may indicate a broken pagination limit. Starting to slice the query by month.`);
  const allSlicedIssues = [];
  const issueKeys = new Set();

  const today = new Date();
  for (let i = 0; i < 24; i++) {
    const startDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const endDate = new Date(today.getFullYear(), today.getMonth() - i + 1, 1);
    
    const formattedStartDate = Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd");
    const formattedEndDate = Utilities.formatDate(endDate, "GMT", "yyyy-MM-dd");

    const slicedJql = `${jql} AND created >= "${formattedStartDate}" AND created < "${formattedEndDate}"`;
    
    console.log(`Fetching slice: ${slicedJql}`);
    const monthlyIssues = searchJiraIssuesLimited(slicedJql, expand);
    
    monthlyIssues.forEach(issue => {
      const issueKey = expand ? issue.key : issue.key; 
      if (!issueKeys.has(issueKey)) {
        issueKeys.add(issueKey);
        allSlicedIssues.push(issue);
      }
    });
  }

  console.log(`Total unique issues found after slicing: ${allSlicedIssues.length}`);
  return allSlicedIssues;
}

function searchJiraIssuesLimited(jql, maxResults = 100) {
  const jiraConfig = getJiraConfig();
  
  // ⚠️ ENFORCE 100 LIMIT - JIRA pagination doesn't work
  if (maxResults > 100) {
    console.warn(`⚠️  searchJiraIssuesLimited called with maxResults=${maxResults}, forcing to 100 (pagination broken)`);
    maxResults = 100;
  }
  
  console.log(`Executing JQL (limit 100): ${jql.substring(0, 200)}...`);
  
  const baseFields = ['key', 'issuetype', 'parent', 'created', 'updated', 'labels', 'summary', 'status', 'resolution', 'description'];
  const customFields = Object.values(FIELD_MAPPINGS);
  const fields = [...new Set([...baseFields, ...customFields])].join(',');

  const baseUrl = `${jiraConfig.baseUrl}/rest/api/3/search/jql`;
  
  let url = `${baseUrl}?`;
  url += `jql=${encodeURIComponent(jql)}`;
  url += `&startAt=0`;
  url += `&maxResults=100`;
  url += `&fields=${encodeURIComponent(fields)}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: getJiraHeaders(),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`JIRA API error: ${response.getResponseCode()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.issues || data.issues.length === 0) {
      console.log(`No issues found`);
      return [];
    }
    
    const parsed = data.issues.map(issue => parseJiraIssue(issue)).filter(Boolean);
    
    // Warn if we hit exactly 100
    if (parsed.length === 100) {
      console.warn(`⚠️  HIT 100-ROW LIMIT - Query may have more results: ${jql.substring(0, 100)}...`);
    }
    
    console.log(`Found ${parsed.length} issues`);
    return parsed;
    
  } catch (error) {
    console.error(`Error executing JQL: ${error.message}`);
    throw error;
  }
}

function fetchAllJiraIssues(jql, expand = null) {
  console.log(`Initial JQL: ${jql}`);
  
  // First, try the query as-is.
  const initialAttempt = searchJiraIssuesLimited(jql, expand);

  // If we get less than 100 results, pagination is not an issue. We're done.
  if (initialAttempt.length < 100) {
    console.log(`Found ${initialAttempt.length} issues, which is under the 100 limit. Returning results directly.`);
    return initialAttempt;
  }

  // If we hit the 100 limit, we must assume there are more results and start slicing.
  console.warn(`Query returned 100 results. This may indicate a broken pagination limit. Starting to slice the query by month.`);
  const allSlicedIssues = [];
  const issueKeys = new Set(); // Use a Set to prevent duplicates

  // We'll slice a 2-year period month by month. Adjust if your PIs are longer.
  const today = new Date();
  for (let i = 0; i < 24; i++) {
    const startDate = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const endDate = new Date(today.getFullYear(), today.getMonth() - i + 1, 1);
    
    const formattedStartDate = Utilities.formatDate(startDate, "GMT", "yyyy-MM-dd");
    const formattedEndDate = Utilities.formatDate(endDate, "GMT", "yyyy-MM-dd");

    // Add the date slice to the original JQL
    const slicedJql = `${jql} AND created >= "${formattedStartDate}" AND created < "${formattedEndDate}"`;
    
    console.log(`Fetching slice: ${slicedJql}`);
    const monthlyIssues = searchJiraIssuesLimited(slicedJql, expand);
    
    monthlyIssues.forEach(issue => {
      // The issue object from the API has a 'key' property
      const issueKey = expand ? issue.key : issue.key; 
      if (!issueKeys.has(issueKey)) {
        issueKeys.add(issueKey);
        allSlicedIssues.push(issue);
      }
    });
  }

  console.log(`Total unique issues found after slicing: ${allSlicedIssues.length}`);
  return allSlicedIssues;
}

function fetchChildDataForEpicsOptimized(epics, includeStoryPoints = true) {
  const epicDataMap = {};
  
  // Initialize map
  epics.forEach(epic => {
    epicDataMap[epic.key] = {
      totalStoryPoints: 0,
      closedStoryPoints: 0,
      totalChildren: 0,
      closedChildren: 0,
      risks: [],
      dependencies: [],
      allChildren: []
    };
  });

  if (epics.length === 0) return epicDataMap;
  
  const epicKeys = epics.map(epic => epic.key);
  console.log(`Processing ${epicKeys.length} epics`);
  
  // ═══════════════════════════════════════════════════════════
  // FETCH ALL DEPENDENCIES (BATCHED TO AVOID URL LENGTH LIMITS)
  // ═══════════════════════════════════════════════════════════
  console.log('\nFetching ALL dependencies for this value stream...');
  
  // Batch size of 40 epics keeps URL well under the ~2KB limit
  // Each epic key is ~10-15 chars + quotes + comma = ~20 chars
  // 40 epics * 20 chars = ~800 chars for the IN clause, leaving room for rest of URL
  const DEP_BATCH_SIZE = 40;
  
  try {
    let allDependencies = [];
    const totalBatches = Math.ceil(epicKeys.length / DEP_BATCH_SIZE);
    
    if (epicKeys.length > DEP_BATCH_SIZE) {
      console.log(`Batching dependency fetch: ${epicKeys.length} epics in ${totalBatches} batches of ${DEP_BATCH_SIZE}`);
    }
    
    for (let i = 0; i < epicKeys.length; i += DEP_BATCH_SIZE) {
      const batch = epicKeys.slice(i, i + DEP_BATCH_SIZE);
      const batchNum = Math.floor(i / DEP_BATCH_SIZE) + 1;
      
      if (totalBatches > 1) {
        console.log(`  Dependency batch ${batchNum}/${totalBatches}: ${batch.length} epics`);
      }
      
      const epicKeysList = batch.map(key => `"${key}"`).join(',');
      const depJql = `parent IN (${epicKeysList}) AND issuetype = Dependency`;
      
      const batchDependencies = searchJiraIssuesLimited(depJql, 100);
      allDependencies = allDependencies.concat(batchDependencies);
      
      // Small delay between batches to be respectful to the API
      if (i + DEP_BATCH_SIZE < epicKeys.length) {
        Utilities.sleep(150);
      }
    }
    
    console.log(`✓ Found ${allDependencies.length} total dependencies`);
    
    // Deduplicate in case a dependency somehow appeared in multiple batches
    const seenDepKeys = new Set();
    const uniqueDependencies = [];
    allDependencies.forEach(dep => {
      if (!seenDepKeys.has(dep.key)) {
        seenDepKeys.add(dep.key);
        uniqueDependencies.push(dep);
      }
    });
    
    if (uniqueDependencies.length < allDependencies.length) {
      console.log(`  (Deduplicated: ${allDependencies.length} → ${uniqueDependencies.length})`);
    }
    
    // Group dependencies by parent epic key
    uniqueDependencies.forEach(dep => {
      const parentKey = dep.parentKey;
      if (parentKey && epicDataMap[parentKey]) {
        epicDataMap[parentKey].dependencies.push(dep);
        epicDataMap[parentKey].allChildren.push(dep);
      } else if (parentKey) {
        console.warn(`  Warning: Dependency ${dep.key} has parent ${parentKey} which is not in our epic list`);
      }
    });
    
    // Log summary
    let epicsWithDeps = 0;
    Object.entries(epicDataMap).forEach(([key, data]) => {
      if (data.dependencies.length > 0) {
        epicsWithDeps++;
        console.log(`  ${key}: ${data.dependencies.length} dependencies`);
      }
    });
    console.log(`${epicsWithDeps} epics have dependencies`);
    
  } catch (error) {
    console.error(`Error fetching dependencies: ${error.message}`);
  }
  
  // ═══════════════════════════════════════════════════════════
  // FETCH ALL STORIES/BUGS (BATCHED - ONE QUERY PER BATCH)
  // ═══════════════════════════════════════════════════════════
  if (false && includeStoryPoints) { 
    console.log('\nFetching ALL stories/bugs for this value stream...');
    
    const epicLinkField = FIELD_MAPPINGS.epicLink.replace('customfield_', '');
    const STORY_BATCH_SIZE = 40;
    
    try {
      let allStories = [];
      const totalBatches = Math.ceil(epicKeys.length / STORY_BATCH_SIZE);
      
      if (epicKeys.length > STORY_BATCH_SIZE) {
        console.log(`Batching story fetch: ${epicKeys.length} epics in ${totalBatches} batches of ${STORY_BATCH_SIZE}`);
      }
      
      for (let i = 0; i < epicKeys.length; i += STORY_BATCH_SIZE) {
        const batch = epicKeys.slice(i, i + STORY_BATCH_SIZE);
        const batchNum = Math.floor(i / STORY_BATCH_SIZE) + 1;
        
        if (totalBatches > 1) {
          console.log(`  Story batch ${batchNum}/${totalBatches}: ${batch.length} epics`);
        }
        
        const epicKeysList = batch.map(key => `"${key}"`).join(',');
        const storyJql = `(cf[${epicLinkField}] IN (${epicKeysList}) OR parent IN (${epicKeysList})) AND issuetype IN (Story, Bug)`;
        
        const batchStories = searchJiraIssuesLimited(storyJql, 100);
        allStories = allStories.concat(batchStories);
        
        if (batchStories.length === 100) {
          console.warn(`  ⚠️ Batch ${batchNum} hit 100-row limit - some stories may be missed`);
        }
        
        // Small delay between batches
        if (i + STORY_BATCH_SIZE < epicKeys.length) {
          Utilities.sleep(150);
        }
      }
      
      console.log(`✓ Found ${allStories.length} total stories/bugs`);
      
      // Deduplicate stories
      const seenStoryKeys = new Set();
      const uniqueStories = [];
      allStories.forEach(story => {
        if (!seenStoryKeys.has(story.key)) {
          seenStoryKeys.add(story.key);
          uniqueStories.push(story);
        }
      });
      
      if (uniqueStories.length < allStories.length) {
        console.log(`  (Deduplicated: ${allStories.length} → ${uniqueStories.length})`);
      }
      
      // Group stories by parent epic key
      uniqueStories.forEach(story => {
        const parentKey = story.epicLink || story.parentKey;
        
        if (parentKey && epicDataMap[parentKey]) {
          const points = story.storyPoints || story.storyPointEstimate || 0;
          epicDataMap[parentKey].totalStoryPoints += points;
          epicDataMap[parentKey].totalChildren++;
          
          if (story.status === 'Closed' || story.status === 'Done' || story.resolution) {
            epicDataMap[parentKey].closedStoryPoints += points;
            epicDataMap[parentKey].closedChildren++;
          }
        } else if (parentKey) {
          console.warn(`  Warning: Story ${story.key} has parent ${parentKey} which is not in our epic list`);
        }
      });
      
      // Log summary
      let epicsWithPoints = 0;
      let totalPoints = 0;
      Object.entries(epicDataMap).forEach(([key, data]) => {
        if (data.totalStoryPoints > 0) {
          epicsWithPoints++;
          totalPoints += data.totalStoryPoints;
          console.log(`  ${key}: ${data.totalChildren} stories/bugs, ${data.totalStoryPoints} points`);
        }
      });
      console.log(`${epicsWithPoints} epics have story points (${totalPoints} total)`);
      
    } catch (error) {
      console.error(`Error fetching stories: ${error.message}`);
    }
  }
  
  // ═══════════════════════════════════════════════════════════
  // FINAL SUMMARY
  // ═══════════════════════════════════════════════════════════
  let epicsWithDeps = 0;
  let epicsWithPoints = 0;
  let totalDeps = 0;
  let totalPoints = 0;
  
  Object.entries(epicDataMap).forEach(([key, data]) => {
    if (data.dependencies.length > 0) {
      epicsWithDeps++;
      totalDeps += data.dependencies.length;
    }
    if (data.totalStoryPoints > 0) {
      epicsWithPoints++;
      totalPoints += data.totalStoryPoints;
    }
  });
  
  console.log(`\n========================================`);
  console.log(`SUMMARY:`);
  console.log(`  Total epics: ${epicKeys.length}`);
  console.log(`  Epics with dependencies: ${epicsWithDeps} (${totalDeps} total)`);
  console.log(`  Epics with story points: ${epicsWithPoints} (${totalPoints} total)`);
  console.log(`========================================\n`);
  
  return epicDataMap;
}







function fetchChildIssues(epicKeys) {
  const jql = `"Epic Link" IN (${epicKeys.join(',')}) OR parent IN (${epicKeys.join(',')})`;
  return searchJiraIssuesLimited(jql);
}

function parseJiraIssue(rawIssue) {
  if (!rawIssue || !rawIssue.key) {
    console.warn('parseJiraIssue: Received invalid issue object (no key)');
    return null;
  }

  try {
    const fields = rawIssue.fields || {};
    const issueType = fields.issuetype ? fields.issuetype.name : '';
    
    console.log(`Parsing ${rawIssue.key} (${issueType})`);

    // ========== FIXED EXTRACT FIELD HELPER ==========
    // This is the critical fix for the [object Object] issue
    const extractField = (fieldValue) => {
      // Handle null/undefined
      if (!fieldValue) return '';
      
      // Handle simple strings - trim whitespace
      if (typeof fieldValue === 'string') {
        return fieldValue.trim();
      }
      
      // Handle numbers
      if (typeof fieldValue === 'number') {
        return fieldValue.toString();
      }
      
      // Handle booleans
      if (typeof fieldValue === 'boolean') {
        return fieldValue.toString();
      }
      
      // Handle arrays - recursively extract from first element
      if (Array.isArray(fieldValue)) {
        if (fieldValue.length === 0) return '';
        return extractField(fieldValue[0]);
      }
      
      // Handle objects - this is where the fix is most important
      if (typeof fieldValue === 'object') {
        // Handle Atlassian Document Format (ADF) - used for rich text fields
        if (fieldValue.type === 'doc' && Array.isArray(fieldValue.content)) {
          let extractedText = '';
          fieldValue.content.forEach(block => {
            if (block.content && Array.isArray(block.content)) {
              block.content.forEach(element => {
                if (element.type === 'text' && element.text) {
                  extractedText += element.text;
                }
              });
              extractedText += '\n';
            }
          });
          return extractedText.trim();
        }
        
        // CRITICAL FIX: Try extracting from common JIRA object properties
        // Most select/dropdown fields use .value
        if (fieldValue.value !== undefined && fieldValue.value !== null) {
          // If value is itself an object, recurse
          if (typeof fieldValue.value === 'object') {
            return extractField(fieldValue.value);
          }
          return String(fieldValue.value).trim();
        }
        
        // Some fields use .name (like user fields, some select fields)
        if (fieldValue.name !== undefined && fieldValue.name !== null) {
          return String(fieldValue.name).trim();
        }
        
        // User fields sometimes use .displayName
        if (fieldValue.displayName !== undefined && fieldValue.displayName !== null) {
          return String(fieldValue.displayName).trim();
        }
        
        // Epic Link and issue references use .key
        if (fieldValue.key !== undefined && fieldValue.key !== null) {
          return String(fieldValue.key).trim();
        }
        
        // If we can't extract a value, log for debugging
        console.warn(`extractField: Could not extract value from object with keys: ${Object.keys(fieldValue).join(', ')}`);
      }
      
      // Last resort: return empty string
      return '';
    };
    // ========== END OF EXTRACT FIELD HELPER ==========

    // Safely extract labels
    const labels = Array.isArray(fields.labels) ? fields.labels : [];
    
    // Check for governance/momentum labels
    const hasGovernance = labels.some(label => 
      GOVERNANCE_LABEL_PATTERNS.some(pattern => 
        label.toLowerCase().includes(pattern)
      )
    );
    
    const hasMomentum = labels.some(label => 
      MOMENTUM_LABEL_PATTERNS.some(pattern => 
        label.toLowerCase().includes(pattern)
      )
    );

    // Check for InfoSec subsection labels (only relevant for INFOSEC portfolio items)
    const infosecSubsection = getInfosecSubsection(labels);

    // Extract parent key
    const parentKeyText = (fields.parent && fields.parent.key) ? fields.parent.key : '';

    // Helper to safely parse numeric fields
    const safeParseFloat = (value) => {
      if (value === null || value === undefined || value === '') return 0;
      const parsed = parseFloat(value);
      return isNaN(parsed) ? 0 : parsed;
    };

    // Build the complete parsed issue object
    const parsedIssue = {
      key: rawIssue.key,
      summary: fields.summary || '',
      description: extractField(fields.description),
      status: fields.status ? fields.status.name : '',
      issueType: issueType,
      resolution: fields.resolution ? fields.resolution.name : '',
      
      // Story Points - handle safely
      storyPoints: safeParseFloat(fields[FIELD_MAPPINGS.storyPoints]),
      storyPointEstimate: safeParseFloat(fields[FIELD_MAPPINGS.storyPointEstimate]),
      
      // Epic fields
      epicLink: extractField(fields[FIELD_MAPPINGS.epicLink]),
      epicName: extractField(fields[FIELD_MAPPINGS.epicName]),
      
      // PI fields - these use the FIXED extractField function
      piObjective: extractField(fields[FIELD_MAPPINGS.piObjective]),
      portfolioInitiative: extractField(fields[FIELD_MAPPINGS.portfolioInitiative]),
      programInitiative: extractField(fields[FIELD_MAPPINGS.programInitiative]),
      programIncrement: extractField(fields[FIELD_MAPPINGS.programIncrement]),
      valueStream: extractField(fields[FIELD_MAPPINGS.valueStream]),
      piCommitment: extractField(fields[FIELD_MAPPINGS.piCommitment]),
      piTargetIteration: extractField(fields[FIELD_MAPPINGS.piTargetIteration]),
      allocation: extractField(fields[FIELD_MAPPINGS.allocation]),
      piObjectiveStatus: extractField(fields[FIELD_MAPPINGS.piObjectiveStatus]),
      scrumTeam: extractField(fields[FIELD_MAPPINGS.scrumTeam]),
      
      // Business value
      businessValue: safeParseFloat(fields[FIELD_MAPPINGS.businessValue]),
      actualValue: safeParseFloat(fields[FIELD_MAPPINGS.actualValue]),
      
      // RAG status
      rag: extractField(fields[FIELD_MAPPINGS.rag]),
      ragNote: extractField(fields[FIELD_MAPPINGS.ragNote]),
      
      // Feature Points
      featurePoints: safeParseFloat(fields[FIELD_MAPPINGS.featurePoints]),
      
      // Other fields
      fixVersions: Array.isArray(fields.fixVersions) ? 
        fields.fixVersions.map(v => v.name).join(', ') : '',
      labels: labels,
      hasGovernanceLabel: hasGovernance,
      hasMomentumLabel: hasMomentum,
      governance: hasGovernance ? 'Yes' : 'No',
      momentum: hasMomentum ? 'Yes' : 'No',
      infosecSubsection: infosecSubsection,
      created: fields.created ? new Date(fields.created).toLocaleDateString() : '',
      updated: fields.updated ? new Date(fields.updated).toLocaleDateString() : '',
      iterationStart: extractField(fields[FIELD_MAPPINGS.iterationStart]),
      iterationEnd: extractField(fields[FIELD_MAPPINGS.iterationEnd]),
      parentKey: parentKeyText,
      benefitHypothesis: extractField(fields[FIELD_MAPPINGS.benefitHypothesis]),
      acceptanceCriteria: extractField(fields[FIELD_MAPPINGS.acceptanceCriteria]),
      dependsOnValuestream: extractField(fields[FIELD_MAPPINGS.dependsOnValuestream]),
      dependsOnTeam: extractField(fields[FIELD_MAPPINGS.dependsOnTeam])
    };
    
    console.log(`✓ Successfully parsed ${rawIssue.key}`);
    return parsedIssue;
    
  } catch (error) {
    console.error(`Error parsing issue ${rawIssue.key}:`, error.toString());
    console.error('Error stack:', error.stack);
    
    // Try to log the problematic raw data for debugging
    try {
      console.error('Raw issue data (first 1000 chars):', JSON.stringify(rawIssue, null, 2).substring(0, 1000));
    } catch (e) {
      console.error('Could not stringify raw issue data');
    }
    
    return null;
  }
}
// ===== DATA PROCESSING =====

function processGovernanceData(epics, epicDataMap, initiativeMap = null) {
  const processedData = [];
  
  epics.forEach(epic => {
    const epicData = epicDataMap[epic.key] || {
      totalStoryPoints: 0,
      closedStoryPoints: 0,
      totalChildren: 0,
      closedChildren: 0,
      risks: [],
      dependencies: [],
      allChildren: []
    };
    
    const totalStoryPoints = epicData.totalStoryPoints || 0;
    const closedStoryPoints = epicData.closedStoryPoints || 0;
    const percentComplete = totalStoryPoints > 0 
      ? Math.round((closedStoryPoints / totalStoryPoints) * 100) 
      : 0;
    
    const totalChildren = epicData.totalChildren || 0;
    const closedChildren = epicData.closedChildren || 0;
    const itemsPercentComplete = totalChildren > 0
      ? Math.round((closedChildren / totalChildren) * 100)
      : 0;
    
    const dependencyCount = epicData.dependencies ? epicData.dependencies.length : 0;
    const riskCount = epicData.risks ? epicData.risks.length : 0;
    
    // OVERRIDE: Force Value Stream to "MMGI-Cloud" for HYD epics
    const finalValueStream = (epic.key && epic.key.startsWith('HYD-')) 
      ? 'MMGI-Cloud' 
      : epic.valueStream;
    
    // Get initiative data from map if available
    let initiativeTitle = '';
    let initiativeDescription = '';
    if (initiativeMap && epic.parentKey) {
      const initiative = initiativeMap.get(epic.parentKey);
      if (initiative) {
        initiativeTitle = initiative.title || '';
        initiativeDescription = initiative.description || '';
      }
    }
    
    // Add the Epic row
    processedData.push({
      'Key': epic.key,
      'Summary': epic.summary,
      'Issue Type': epic.issueType,
      'Status': epic.status,
      'Value Stream/Org': finalValueStream,
      'Program Increment': epic.programIncrement, 
      'PI Objective': epic.piObjective,
      'PI Objective Status': epic.piObjectiveStatus,
      'PI Commitment': epic.piCommitment,
      'Portfolio Initiative': epic.portfolioInitiative,
      'Program Initiative': epic.programInitiative,
      'Infosec Subsection': epic.infosecSubsection || '',
      'Initiative Title': initiativeTitle,
      'Initiative Description': initiativeDescription,
      'Scrum Team': epic.scrumTeam,
      'PI Target Iteration': epic.piTargetIteration,
      'Allocation': epic.allocation,
      'Business Value': epic.businessValue,
      'Actual Value': epic.actualValue,
      'RAG': epic.rag,
      'RAG Note': epic.ragNote,
      'Resolution': epic.resolution,
      'Fix Versions': epic.fixVersions,
      'Benefit Hypothesis': epic.benefitHypothesis,
      'Acceptance Criteria': epic.acceptanceCriteria,
      'Parent Key': epic.parentKey,
      'Total Story Points': totalStoryPoints,
      'Closed Story Points': closedStoryPoints,
      '% Complete (Points)': percentComplete,
      'Total Items': totalChildren,
      'Closed Items': closedChildren,
      '% Complete (Items)': itemsPercentComplete,
      'Dependencies': dependencyCount,
      'Risks': riskCount,
      'Feature Points': epic.featurePoints || '',
      'Has Governance Label': epic.hasGovernanceLabel ? 'Yes' : 'No',
      'Has Momentum Label': epic.hasMomentumLabel ? 'Yes' : 'No',
      'Start Iteration Name': epic.iterationStart || '',
      'End Iteration Name': epic.iterationEnd || '',
      'Depends on Valuestream': '',
      'Depends on Team': '',
      'Created Date': epic.created || '',
      'Updated Date': epic.updated || ''
    });
    
    // Add child dependency rows under this Epic
    if (epicData.dependencies && epicData.dependencies.length > 0) {
      epicData.dependencies.forEach(dep => {
        processedData.push({
          'Key': dep.key,
          'Summary': dep.summary,
          'Issue Type': dep.issueType,
          'Status': dep.status,
          'Value Stream/Org': dep.valueStream || '',
          'Program Increment': epic.programIncrement || '',
          'PI Objective': '',
          'PI Objective Status': '',
          'PI Commitment': '',
          'Portfolio Initiative': epic.portfolioInitiative || '',
          'Program Initiative': epic.programInitiative || '',
          'Infosec Subsection': '',
          'Initiative Title': initiativeTitle,
          'Initiative Description': initiativeDescription,
          'Scrum Team': dep.scrumTeam || '',
          'PI Target Iteration': dep.piTargetIteration || '',
          'Allocation': '',
          'Business Value': '',
          'Actual Value': '',
          'RAG': dep.rag || '',
          'RAG Note': dep.ragNote || '',
          'Resolution': dep.resolution || '',
          'Fix Versions': dep.fixVersions || '',
          'Benefit Hypothesis': '',
          'Acceptance Criteria': '',
          'Parent Key': epic.key,
          'Total Story Points': 0,
          'Closed Story Points': 0,
          '% Complete (Points)': 0,
          'Total Items': 0,
          'Closed Items': 0,
          '% Complete (Items)': 0,
          'Dependencies': 0,
          'Risks': 0,
          'Feature Points': '',
          'Has Governance Label': 'No',
          'Has Momentum Label': 'No',
          'Depends on Valuestream': dep.dependsOnValuestream || '',
          'Depends on Team': dep.dependsOnTeam || '',
          'Created Date': dep.created || '',
          'Updated Date': dep.updated || '',
          'End Iteration Name': dep.iterationEnd || '',
          'Start Iteration Name': dep.iterationStart || ''
        });
      });
    }
  });
  
  return processedData;
}

function getIterationDates(piNumber, iterationNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Iteration Parameters');
  
  if (!sheet) {
    const message = 'Error: The required sheet "Iteration Parameters" was not found.';
    SpreadsheetApp.getUi().alert(message);
    throw new Error(message);
  }
  
  if (iterationNumber == 0) {
    console.log(`Iteration 0 selected - No date filtering will be applied`);
    return { start: null, end: null, isPrePlanning: true };
  }
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] == piNumber && row[1] == iterationNumber) {
      const startDate = row[2];
      const endDate = row[3];
      
      if (!startDate || !endDate) {
        throw new Error(`Missing date data for PI ${piNumber}, Iteration ${iterationNumber} in Iteration Parameters sheet.`);
      }
      
      console.log(`Found dates for PI ${piNumber}, Iteration ${iterationNumber}: Start=${startDate}, End=${endDate}`);
      return { start: startDate, end: endDate, isPrePlanning: false };
    }
  }
  
  throw new Error(`No matching iteration found for PI ${piNumber}, Iteration ${iterationNumber} in Iteration Parameters sheet.`);
}

function analyzeIterationChanges(piNumber, iterationNumber, valueStreams) {
 return logActivity('PI Iteration Change Analysis', () => {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startTime = Date.now();

  try {
    const programIncrement = `PI ${piNumber}`;
    
    // *** NEW: Check if this is Iteration 0 (Pre-Planning) ***
    const isPrePlanning = (iterationNumber == 0);
    
    if (isPrePlanning) {
      // For Iteration 0, fetch all data for the PI without date filtering
      ss.toast("Fetching all data for PI (Pre-Planning mode)...", "⏳ Processing...", 20);
      console.log("Step 1: Pre-Planning mode - skipping date lookup");
      
      ss.toast("Step 2/4: Fetching epics from Jira...", "⏳ Processing...", 20);
      const baseEpics = fetchEpicsForPIOptimized(programIncrement, valueStreams);
      
      if (baseEpics.length === 0) {
        ui.alert('No Data', `No epics found for the selected criteria in ${programIncrement}.`, ui.ButtonSet.OK);
        return;
      }
      
      console.log(`Step 2 Complete: Found ${baseEpics.length} epics.`);
      
      ss.toast(`Step 3/4: Fetching child issues...`, "⏳ Processing...", 20);
      const epicKeys = baseEpics.map(e => e.key);
      const childIssues = fetchChildIssuesForEpics(epicKeys);
      console.log(`Step 3 Complete: Found ${childIssues.length} child issues.`);
      
      ss.toast(`Step 4/4: Processing data...`, "⏳ Processing...", 20);
      const epicDataMap = calculateMetricsFromChildren(baseEpics, childIssues);
      const processedData = processGovernanceData(baseEpics, epicDataMap);
      
      const valueStreamForWrite = (Array.isArray(valueStreams) && valueStreams.length === 1) ? valueStreams[0] : null;
      const sheetName = `PI ${piNumber} - Iteration 0 - Pre-Planning`;
      writeGovernanceDataToSheet(processedData, programIncrement, valueStreamForWrite, sheetName, true);
      
      const duration = (Date.now() - startTime) / 1000;
      ss.toast(`✅ Pre-Planning report complete in ${duration.toFixed(1)} seconds!`, "✅ Success", 10);
      return;
    }
    
    // *** EXISTING CODE: Normal iteration processing with date filtering ***
    ss.toast("Step 1/6: Getting iteration dates...", "⏳ Processing...", 20);
    const dates = getIterationDates(piNumber, iterationNumber);
    const analysisDate = new Date(dates.end.getTime() + (3 * 24 * 60 * 60 * 1000));
    console.log("Step 1 Complete: Dates retrieved.");

    ss.toast("Step 2/6: Fetching epics from Jira...", "⏳ Processing...", 20);
    const baseEpics = fetchEpicsForPIOptimized(programIncrement, valueStreams);
    if (baseEpics.length === 0) {
      ui.alert('No Data', `No epics found for the selected criteria in ${programIncrement}.`, ui.ButtonSet.OK);
      return;
    }
    const epicKeys = baseEpics.map(e => e.key);
    console.log(`Step 2 Complete: Found ${epicKeys.length} base epics.`);
    
    ss.toast(`Step 3/6: Fetching child issues...`, "⏳ Processing...", 20);
    const childIssues = fetchChildIssuesForEpics(epicKeys);
    console.log(`Step 3 Complete: Found ${childIssues.length} child issues.`);
    
    const allIssueKeys = [...new Set(epicKeys.concat(childIssues.map(c => c.key)))];
    console.log(`Total of ${allIssueKeys.length} unique issues to analyze.`);
    
    ss.toast(`Step 4/6: Fetching historical data...`, "⏳ This is the longest step...", 30);
    const allIssuesWithChangelog = fetchIssuesInChunks(allIssueKeys, 25, 'changelog');
    console.log(`Step 4 Complete: Fetched ${allIssuesWithChangelog.length} issues with historical data.`);

    ss.toast(`Step 5/6: Reconstructing data...`, "⏳ Processing...", 20);
    const reconstructedIssues = allIssuesWithChangelog
      .map(issue => {
        if (new Date(issue.fields.created) > analysisDate) return null; 
        const reconstructedFields = reconstructIssueFields(issue.fields, issue.changelog, analysisDate);
        return parseJiraIssue({ key: issue.key, fields: reconstructedFields });
      })
      .filter(Boolean);
    console.log(`Step 5 Complete: Reconstructed ${reconstructedIssues.length} issues to the state on ${analysisDate.toLocaleDateString()}.`);

    ss.toast("Step 6/6: Calculating metrics and writing report...", "⏳ Processing...", 20);
    const reconstructedEpics = reconstructedIssues.filter(i => i.issueType === 'Epic');
    const epicDataMap = calculateMetricsFromChildren(reconstructedEpics, reconstructedIssues);
    const processedData = processGovernanceData(reconstructedEpics, epicDataMap);
    const valueStreamForWrite = (Array.isArray(valueStreams) && valueStreams.length === 1) ? valueStreams[0] : null;
    const sheetName = `PI ${piNumber} - Iteration ${iterationNumber} - Changes`;
    writeGovernanceDataToSheet(processedData, programIncrement, valueStreamForWrite, sheetName, true);
    console.log("Step 6 Complete: Report written to sheet.");

    const duration = (Date.now() - startTime) / 1000;
    ss.toast(`✅ Analysis complete in ${duration.toFixed(1)} seconds!`, "✅ Success", 10);

  } catch (error) {
    console.error("Iteration Analysis Error:", error);
    ui.alert("Error", error.toString(), ui.ButtonSet.OK);
  }
 }, { piNumber, iterationNumber, streamCount: selectedValueStreams?.length });
}

function fetchIssuesInChunks(issueKeys, chunkSize = 25, expand = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allFetchedIssues = [];
  const totalChunks = Math.ceil(issueKeys.length / chunkSize);
  console.log(`Fetching ${issueKeys.length} issues in chunks of ${chunkSize}.`);

  for (let i = 0; i < issueKeys.length; i += chunkSize) {
    const currentChunk = Math.floor(i / chunkSize) + 1;
    // Update the UI toast for each new chunk.
    ss.toast(`Step 4/6: Fetching historical data (batch ${currentChunk} of ${totalChunks})...`, "⏳ Processing...", 30);
    
    const chunk = issueKeys.slice(i, i + chunkSize);
    const jql = `key IN (${chunk.join(',')})`;
    
    console.log(`--- Fetching chunk ${currentChunk} of ${totalChunks} ---`);
    const chunkResults = searchJiraIssuesLimited(jql, expand);
    allFetchedIssues = allFetchedIssues.concat(chunkResults);
    Utilities.sleep(250);
  }

  console.log(`Finished fetching in chunks. Total issues retrieved: ${allFetchedIssues.length}`);
  return allFetchedIssues;
}
function fetchChildIssuesInBatches(epicKeys, batchSize = 15) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allChildIssues = [];
  const uniqueChildKeys = new Set(); // Use a Set to prevent adding duplicate issues
  const totalBatches = Math.ceil(epicKeys.length / batchSize);

  console.log(`Fetching child issues for ${epicKeys.length} epics in ${totalBatches} batches of up to ${batchSize} epics each.`);

  for (let i = 0; i < epicKeys.length; i += batchSize) {
    const batch = epicKeys.slice(i, i + batchSize);
    const currentBatchNum = Math.floor(i / batchSize) + 1;
    
    // Provide real-time feedback to the user in the spreadsheet UI
    ss.toast(`Step 2/3: Fetching children (batch ${currentBatchNum} of ${totalBatches})...`, "⏳ Refreshing", 20);
    
    // Construct a JQL query for the current small batch of epics
    const jql = `"Epic Link" IN (${batch.join(',')}) OR parent IN (${batch.join(',')})`;
    const batchResults = searchJiraIssuesLimited(jql); // This function already handles the basic API call

    // Process the results from the batch, ensuring no duplicates are added
    batchResults.forEach(issue => {
      if (!uniqueChildKeys.has(issue.key)) {
        uniqueChildKeys.add(issue.key);
        allChildIssues.push(issue);
      }
    });
    
    // A small delay to be respectful to the JIRA API and avoid rate limiting
    Utilities.sleep(250); 
  }

  console.log(`Finished fetching children in batches. Total unique child issues found: ${allChildIssues.length}`);
  return allChildIssues;
}
function fetchChildIssuesForEpics(epicKeys) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let allChildIssues = [];
  const uniqueChildKeys = new Set();
  console.log(`Fetching child issues for ${epicKeys.length} epics one-by-one.`);

  epicKeys.forEach((epicKey, index) => {
    if (index > 0 && index % 10 === 0) {
      ss.toast(`Fetching children for epic ${index + 1} of ${epicKeys.length}...`, "🔍 Analyzing", 20);
    }
    
    // FIX: Added quotation marks around the ${epicKey} to create a valid JQL query.
    const jql = `"Epic Link" = "${epicKey}" OR parent = "${epicKey}"`;
    
    // --- DIAGNOSTIC LOG --- (You can remove this later)
    console.log(`Executing JQL to find children: ${jql}`); 
    
    const childResults = searchJiraIssuesLimited(jql);

    // --- DIAGNOSTIC LOG --- (You can remove this later)
    console.log(`Found ${childResults.length} children for ${epicKey}.`);

    childResults.forEach(issue => {
      if (issue && issue.key && !uniqueChildKeys.has(issue.key)) {
        uniqueChildKeys.add(issue.key);
        allChildIssues.push(issue);
      }
    });
    Utilities.sleep(100); // Small delay to be respectful to the API
  });

  console.log(`Finished fetching children. Total unique child issues found: ${allChildIssues.length}`);
  return allChildIssues;
}
function reconstructIssueFields(currentFields, changelog, analysisDate) {
  // Start with a clone of the current fields
  const reconstructed = JSON.parse(JSON.stringify(currentFields));

  if (!changelog || !changelog.histories || !analysisDate) {
    return reconstructed; // Return current state if no history or date
  }

  // Iterate through the changelog histories backwards (from newest to oldest)
  for (let i = changelog.histories.length - 1; i >= 0; i--) {
    const history = changelog.histories[i];
    const changeDate = new Date(history.created);

    // If the change happened AFTER our target analysis date, revert it
    if (changeDate > analysisDate) {
      history.items.forEach(item => {
        // The 'fromString' value is the state of the field BEFORE the change
        const previousValue = item.fromString; 
        const fieldId = item.fieldId;

        // Simple logic to revert the field.
        // NOTE: This needs to be expanded for complex fields (like status, resolution, etc.)
        // which are objects, not simple strings.
        if (fieldId in reconstructed) {
           // This is a simplified example. For fields like 'status' or 'resolution',
           // you'd need to reconstruct the object: e.g., { name: previousValue }
           reconstructed[fieldId] = previousValue;
        }
      });
    }
  }
  
  return reconstructed;
}
function calculateMetricsFromChildren(epics, allChildIssues) {
  const epicDataMap = {};
  
  // Initialize map
  epics.forEach(epic => {
    epicDataMap[epic.key] = {
      totalStoryPoints: 0,
      closedStoryPoints: 0,
      totalChildren: 0,
      closedChildren: 0,
      risks: [],
      dependencies: [],
      allChildren: []
    };
  });
  
  // Process child issues
  allChildIssues.forEach(child => {
    const parentEpicKey = child.epicLink || child.parentKey;
    
    if (parentEpicKey && epicDataMap[parentEpicKey]) {
      const points = child.storyPoints || child.storyPointEstimate || 0;
      epicDataMap[parentEpicKey].totalStoryPoints += points;
      epicDataMap[parentEpicKey].totalChildren++;
      epicDataMap[parentEpicKey].allChildren.push(child);
      
      if (child.status === 'Closed' || child.status === 'Done' || child.resolution) {
        epicDataMap[parentEpicKey].closedStoryPoints += points;
        epicDataMap[parentEpicKey].closedChildren++;
      }
      
      // Track dependencies
      if (child.issueType === 'Dependency') {
        epicDataMap[parentEpicKey].dependencies.push(child);
      }
    }
  });
  
  return epicDataMap;
}
function fetchEpicsWithPagination(jql, maxResults = 100) {
  const allEpics = [];
  const seenKeys = new Set();
  
  // First call: ORDER BY key ASC (gets first 100 alphabetically)
  console.log(`Fetch 1/2: ORDER BY key ASC`);
  const ascJql = `${jql} ORDER BY key ASC`;
  const ascEpics = searchJiraIssuesLimited(ascJql, maxResults);
  
  ascEpics.forEach(epic => {
    if (!seenKeys.has(epic.key)) {
      seenKeys.add(epic.key);
      allEpics.push(epic);
    }
  });
  
  console.log(`Fetch 1/2: Got ${ascEpics.length} epics`);
  
  // If we hit the limit, fetch from the other end
  if (ascEpics.length === maxResults) {
    console.log(`⚠️ HIT ${maxResults} ROW LIMIT - Fetching from other end...`);
    
    Utilities.sleep(200);
    
    // Second call: ORDER BY key DESC (gets last 100 alphabetically)
    console.log(`Fetch 2/2: ORDER BY key DESC`);
    const descJql = `${jql} ORDER BY key DESC`;
    const descEpics = searchJiraIssuesLimited(descJql, maxResults);
    
    let newCount = 0;
    descEpics.forEach(epic => {
      if (!seenKeys.has(epic.key)) {
        seenKeys.add(epic.key);
        allEpics.push(epic);
        newCount++;
      }
    });
    
    console.log(`Fetch 2/2: Got ${descEpics.length} epics (${newCount} new, ${descEpics.length - newCount} duplicates)`);
  }
  
  console.log(`✅ Pagination complete: ${allEpics.length} total unique epics`);
  return allEpics;
}
function getIterationDisplayText(iterationValue) {
  // If no value, empty, or whitespace only
  if (!iterationValue || iterationValue.toString().trim() === '') {
    // Return null to indicate "hide this" or return placeholder
    // Based on your current code pattern, return null and let callers use || 'Iteration unknown'
    return null;
  }
  
  const trimmed = iterationValue.toString().trim();
  
  // Check for placeholder values that should be treated as missing
  if (trimmed.toLowerCase() === 'unknown' || 
      trimmed.toLowerCase() === 'tbd' ||
      trimmed === '-') {
    return null;
  }
  
  // Return the cleaned iteration value
  return trimmed;
}

function dismissToast(message = '', title = '', duration = 3) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (message) {
    ss.toast(message, title, duration);
  } else {
    ss.toast('', '', 1); // Dismiss immediately
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PARALLEL JIRA FETCHING - Add to end of Core.gs
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Executes multiple JQL queries in parallel using UrlFetchApp.fetchAll()
 * Returns a map of queryId -> parsed issues
 * 
 * @param {Array<{id: string, jql: string}>} queries - Array of query objects
 * @returns {Object} Map of id -> array of parsed issues
 */
function searchJiraIssuesParallel(queries) {
  if (!queries || queries.length === 0) return {};
  
  const jiraConfig = getJiraConfig();
  const headers = getJiraHeaders();
  const baseUrl = `${jiraConfig.baseUrl}/rest/api/3/search/jql`;
  
  const baseFields = ['key', 'issuetype', 'parent', 'created', 'updated', 'labels', 'summary', 'status', 'resolution', 'description'];
  const customFields = Object.values(FIELD_MAPPINGS);
  const fields = [...new Set([...baseFields, ...customFields])].join(',');
  
  console.log(`⚡ Parallel fetch: ${queries.length} queries`);
  
  // Build all requests
  const requests = queries.map(q => ({
    url: `${baseUrl}?jql=${encodeURIComponent(q.jql)}&startAt=0&maxResults=100&fields=${encodeURIComponent(fields)}`,
    method: 'GET',
    headers: headers,
    muteHttpExceptions: true,
    _queryId: q.id  // Store for mapping results back
  }));
  
  // Execute in batches of 15 to avoid overwhelming the API
  const BATCH_SIZE = 15;
  const results = {};
  
  for (let i = 0; i < requests.length; i += BATCH_SIZE) {
    const batch = requests.slice(i, i + BATCH_SIZE);
    const batchNum = Math.floor(i / BATCH_SIZE) + 1;
    const totalBatches = Math.ceil(requests.length / BATCH_SIZE);
    
    if (totalBatches > 1) {
      console.log(`  Batch ${batchNum}/${totalBatches}: ${batch.length} requests`);
    }
    
    try {
      const responses = UrlFetchApp.fetchAll(batch);
      
      responses.forEach((response, idx) => {
        const queryId = batch[idx]._queryId;
        
        if (response.getResponseCode() === 200) {
          const data = JSON.parse(response.getContentText());
          const parsed = (data.issues || []).map(issue => parseJiraIssue(issue)).filter(Boolean);
          
          if (parsed.length === 100) {
            console.warn(`  ⚠️ Query "${queryId}" hit 100-row limit`);
          }
          
          results[queryId] = parsed;
        } else {
          console.error(`  ✗ Query "${queryId}" failed: ${response.getResponseCode()}`);
          results[queryId] = [];
        }
      });
      
    } catch (e) {
      console.error(`  Parallel batch ${batchNum} failed: ${e.message}`);
      batch.forEach(req => {
        results[req._queryId] = [];
      });
    }
    
    // Small delay between batches
    if (i + BATCH_SIZE < requests.length) {
      Utilities.sleep(100);
    }
  }
  
  return results;
}

/**
 * Fetches epics for multiple value streams in parallel
 * 
 * @param {string} programIncrement - e.g., "PI 13"
 * @param {Array<string>} valueStreams - Array of value stream names
 * @param {boolean} includePrePlanning - Whether to include Pre-Planning epics
 * @returns {Object} Map of valueStream -> array of epics
 */
function fetchEpicsForAllValueStreamsParallel(programIncrement, valueStreams, includePrePlanning = false) {
  const excludedProjectsList = EXCLUDED_PROJECTS.map(p => `"${p}"`).join(', ');
  const piFieldId = FIELD_MAPPINGS.programIncrement.replace('customfield_', '');
  const vsFieldId = FIELD_MAPPINGS.valueStream.replace('customfield_', '');
  
  console.log(`\n${'═'.repeat(60)}`);
  console.log(`⚡ PARALLEL EPIC FETCH: ${valueStreams.length} value streams`);
  console.log(`PI: ${programIncrement} | Pre-Planning: ${includePrePlanning}`);
  console.log(`${'═'.repeat(60)}\n`);
  
  // Build queries for all value streams
  const queries = [];
  
  valueStreams.forEach(vs => {
    // Standard PI query
    let jql;
    if (vs === 'MMPM' || vs === 'EMA Clinical') {
      jql = `issuetype = Epic AND cf[${piFieldId}] = "${programIncrement}" AND cf[${vsFieldId}] = "${vs}"`;
    } else if (vs === 'MMGI-Cloud') {
      jql = `issuetype = Epic AND cf[${piFieldId}] = "${programIncrement}" AND project = HYD`;
    } else {
      jql = `issuetype = Epic AND cf[${piFieldId}] = "${programIncrement}" AND cf[${vsFieldId}] = "${vs}" AND project NOT IN (${excludedProjectsList})`;
    }
    
    queries.push({ id: `${vs}_ASC`, jql: `${jql} ORDER BY created ASC` });
    queries.push({ id: `${vs}_DESC`, jql: `${jql} ORDER BY created DESC` });
    
    // Pre-Planning queries if needed
    if (includePrePlanning) {
      let prePlanJql;
      if (vs === 'MMPM' || vs === 'EMA Clinical') {
        prePlanJql = `issuetype = Epic AND cf[${piFieldId}] = "Pre-Planning" AND cf[${vsFieldId}] = "${vs}"`;
      } else if (vs === 'MMGI-Cloud') {
        prePlanJql = `issuetype = Epic AND cf[${piFieldId}] = "Pre-Planning" AND project = HYD`;
      } else {
        prePlanJql = `issuetype = Epic AND cf[${piFieldId}] = "Pre-Planning" AND cf[${vsFieldId}] = "${vs}" AND project NOT IN (${excludedProjectsList})`;
      }
      
      queries.push({ id: `${vs}_PREPLAN`, jql: prePlanJql });
    }
  });
  
  console.log(`Built ${queries.length} queries for parallel execution`);
  
  // Execute all queries in parallel
  const results = searchJiraIssuesParallel(queries);
  
  // Combine results per value stream (deduplicate ASC/DESC)
  const epicsByValueStream = {};
  
  valueStreams.forEach(vs => {
    const seenKeys = new Set();
    const allEpics = [];
    
    // Combine ASC and DESC results
    const ascResults = results[`${vs}_ASC`] || [];
    const descResults = results[`${vs}_DESC`] || [];
    const prePlanResults = results[`${vs}_PREPLAN`] || [];
    
    [...ascResults, ...descResults, ...prePlanResults].forEach(epic => {
      if (!seenKeys.has(epic.key)) {
        seenKeys.add(epic.key);
        allEpics.push(epic);
      }
    });
    
    epicsByValueStream[vs] = allEpics;
    console.log(`  ${vs}: ${allEpics.length} unique epics (ASC: ${ascResults.length}, DESC: ${descResults.length}${includePrePlanning ? `, PrePlan: ${prePlanResults.length}` : ''})`);
  });
  
  return epicsByValueStream;
}

/**
 * Fetches dependencies for multiple epic sets in parallel
 * 
 * @param {Object} epicsByValueStream - Map of valueStream -> array of epics
 * @returns {Object} Map of epicKey -> array of dependencies
 */
function fetchDependenciesParallel(epicsByValueStream) {
  // Collect all epic keys
  const allEpicKeys = [];
  Object.values(epicsByValueStream).forEach(epics => {
    epics.forEach(epic => allEpicKeys.push(epic.key));
  });
  
  if (allEpicKeys.length === 0) return {};
  
  console.log(`\n⚡ Parallel dependency fetch: ${allEpicKeys.length} epics`);
  
  // Build batched queries (40 epics per query for URL length limits)
  const BATCH_SIZE = 40;
  const queries = [];
  
  for (let i = 0; i < allEpicKeys.length; i += BATCH_SIZE) {
    const batch = allEpicKeys.slice(i, i + BATCH_SIZE);
    const epicKeysList = batch.map(key => `"${key}"`).join(',');
    const jql = `parent IN (${epicKeysList}) AND issuetype = Dependency`;
    
    queries.push({ id: `dep_batch_${i}`, jql });
  }
  
  // Execute in parallel
  const results = searchJiraIssuesParallel(queries);
  
  // Collect and deduplicate all dependencies
  const seenDepKeys = new Set();
  const allDependencies = [];
  
  Object.values(results).forEach(deps => {
    deps.forEach(dep => {
      if (!seenDepKeys.has(dep.key)) {
        seenDepKeys.add(dep.key);
        allDependencies.push(dep);
      }
    });
  });
  
  console.log(`  Found ${allDependencies.length} total dependencies`);
  
  // Group by parent epic key
  const dependenciesByEpic = {};
  allDependencies.forEach(dep => {
    const parentKey = dep.parentKey;
    if (parentKey) {
      if (!dependenciesByEpic[parentKey]) {
        dependenciesByEpic[parentKey] = [];
      }
      dependenciesByEpic[parentKey].push(dep);
    }
  });
  
  return dependenciesByEpic;
}

/**
 * Fetches initiative details in parallel
 * 
 * @param {Set<string>} initiativeKeys - Set of initiative keys to fetch
 * @returns {Map<string, Object>} Map of key -> {title, description}
 */
function fetchInitiativesParallel(parentKeys) {
  const keysArray = Array.from(parentKeys);
  
  if (keysArray.length === 0) {
    return new Map();
  }
  
  console.log(`⚡ Fetching ${keysArray.length} initiatives in parallel...`);
  
  const jiraConfig = getJiraConfig();
  const headers = getJiraHeaders();
  const initiativeMap = new Map();
  
  // Helper to extract text from JIRA fields (including ADF format)
  const extractText = (fieldValue) => {
    if (!fieldValue) return '';
    if (typeof fieldValue === 'string') return fieldValue.trim();
    if (typeof fieldValue === 'number') return fieldValue.toString();
    
    if (Array.isArray(fieldValue)) {
      return fieldValue.length > 0 ? extractText(fieldValue[0]) : '';
    }
    
    if (typeof fieldValue === 'object') {
      // Handle Atlassian Document Format (ADF) - used for description fields
      if (fieldValue.type === 'doc' && Array.isArray(fieldValue.content)) {
        let text = '';
        fieldValue.content.forEach(block => {
          if (block.content && Array.isArray(block.content)) {
            block.content.forEach(element => {
              if (element.type === 'text' && element.text) {
                text += element.text;
              }
            });
            text += '\n';
          }
        });
        return text.trim();
      }
      
      // Standard JIRA field object properties
      if (fieldValue.value !== undefined) return extractText(fieldValue.value);
      if (fieldValue.name) return String(fieldValue.name).trim();
      if (fieldValue.displayName) return String(fieldValue.displayName).trim();
      if (fieldValue.key) return String(fieldValue.key).trim();
    }
    
    return '';
  };
  
  // Batch keys into groups of 25 for JQL IN clause
  const BATCH_SIZE = 25;
  const batches = [];
  
  for (let i = 0; i < keysArray.length; i += BATCH_SIZE) {
    batches.push(keysArray.slice(i, i + BATCH_SIZE));
  }
  
  // Build requests for all batches
  const requests = batches.map((batch, idx) => {
    const keyList = batch.map(k => `"${k}"`).join(',');
    const jql = `key IN (${keyList})`;
    const fields = 'summary,description';
    
    return {
      url: `${jiraConfig.baseUrl}/rest/api/3/search/jql?jql=${encodeURIComponent(jql)}&maxResults=100&fields=${encodeURIComponent(fields)}`,
      method: 'GET',
      headers: headers,
      muteHttpExceptions: true,
      _batchIndex: idx
    };
  });
  
  // Execute all requests in parallel
  try {
    const responses = UrlFetchApp.fetchAll(requests);
    
    responses.forEach((response, idx) => {
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        
        if (data.issues) {
          data.issues.forEach(issue => {
            initiativeMap.set(issue.key, {
              title: issue.fields.summary || '',
              description: extractText(issue.fields.description)
            });
          });
        }
      } else {
        console.warn(`  Batch ${idx} failed: ${response.getResponseCode()}`);
      }
    });
    
  } catch (e) {
    console.error(`Initiative fetch failed: ${e.message}`);
  }
  
  console.log(`  ✓ Fetched ${initiativeMap.size} initiatives`);
  return initiativeMap;
}