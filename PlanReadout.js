/**
 * =================================================================
 * PI Readout Summarizer (Controller) - ENHANCED WITH INITIATIVE CONTEXT
 * =================================================================
 * Orchestrates the generation of both Business and Technical summaries
 * by analyzing initiative hierarchy and PI-specific epic goals.
 * 
 * Key enhancements:
 * - Understands initiative → epic parent-child relationships
 * - Incorporates initiative-level context (Title, Description)
 * - Summarizes what will be accomplished in THIS PI
 * - Handles epics without initiatives gracefully
 */
function generateEpicPerspectives() {
  return logActivity('Epic-Level Perspectives', () => {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    const sheetName = sheet.getName();
    if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
      ui.alert('Error', 'Please run this function on a valid report sheet.', ui.ButtonSet.OK);
      return;
    }
    
    ss.toast('Starting epic-level perspective generation...', '🔍 Analyzing Sheet', 10);
    const headerRow = 4;

    try {
      const COLUMN_MAPPINGS = {
        key: ['Key'],
        parentKey: ['Parent Key'],
        issueType: ['Issue Type'],
        summary: ['Summary'],
        piObjective: ['PI Objective'],
        benefitHypothesis: ['Benefit Hypothesis'],
        acceptanceCriteria: ['Acceptance Criteria'],
        initiativeTitle: ['Initiative Title'],
        initiativeDescription: ['Initiative Description']
      };
      
      let headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      const dataValues = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn()).getValues();

      // Business Perspective Job
      const businessJob = {
        newColumnHeader: 'Business Perspective',
        summaryPrompt: `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

TASK:
Write an executive micro-summary of the epic's BUSINESS VALUE.

STYLE RULES:
- Maximum 25 words. If your response exceeds 25 words, shorten it.
- First word must be a present-participle verb (e.g., Delivering, Enabling, Modernizing, Establishing, Streamlining, Automating)
- BANNED first words: "Optimizing", "Enhancing", "Improving", "Updating" (too vague)
- Format: [Verb] + [capability/outcome] + [who/why it matters]
- Focus on business outcome: time saved, risk reduced, revenue protected, adoption improved
- Avoid technical implementation details
- BANNED phrases: "This PI will", "In this PI", "This work will focus on", "The team will"
- Ignore any blank fields and work with what is available

EXAMPLE (22 words):
"Delivering automated payment reconciliation for the billing team to reduce manual processing errors and accelerate month-end close by 3 days."`,
        sourceColumns: ['key', 'parentKey', 'summary', 'piObjective', 'benefitHypothesis', 'acceptanceCriteria', 'initiativeTitle', 'initiativeDescription'],
        filterColumn: 'issueType',
        filterValue: 'Epic'
      };
      
      // Technical Perspective Job
      const technicalJob = {
        newColumnHeader: 'Technical Perspective',
        summaryPrompt: `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

TASK:
Write an executive micro-summary of WHAT IS BEING BUILT (technical deliverable).

STYLE RULES:
- Maximum 20 words. If your response exceeds 20 words, shorten it.
- Start with: "The feature...", "The system...", or "The platform..."
- Use exactly one main verb from: builds, integrates, migrates, implements, establishes, automates. Avoid chaining with "and."
- Optional: one short "to..." clause at the end (maximum 6 words)
- Focus on systems, capabilities, integrations, infrastructure — not process or staffing
- BANNED phrases: "Engineers will", "We will", "The team will", "This PI will"
- BANNED words: Epic, Initiative, Story — use "feature", "system", "platform" instead
- Ignore any blank fields and work with what is available

EXAMPLE (16 words):
"The platform integrates OAuth 2.0 authentication with legacy billing systems to enable single sign-on access."`,
        sourceColumns: ['key', 'parentKey', 'summary', 'piObjective', 'benefitHypothesis', 'acceptanceCriteria', 'initiativeTitle', 'initiativeDescription'],
        filterColumn: 'issueType',
        filterValue: 'Epic'
      };
      
      // Execute
      ss.toast('Step 1/2: Generating Business Perspective...', '🤖 AI Processing', 5);
      generateEnhancedPIReadoutHelper(ss, ui, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, businessJob);
      
      ss.toast('Step 2/2: Generating Technical Perspective...', '🤖 AI Processing', 5);
      headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      generateEnhancedPIReadoutHelper(ss, ui, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, technicalJob);
      
      ss.toast('✅ Epic-level perspectives complete! Now run "Generate Initiative Perspectives"', '✅ Success', 10);

    } catch (error) {
      console.error('Epic Perspective Error:', error);
      ui.alert('An Error Occurred', error.toString(), ui.ButtonSet.OK);
    }
  }, { sheetName: SpreadsheetApp.getActiveSheet().getName() });
}

function generateInitiativePerspectives() {
  return logActivity('Initiative-Level Perspectives', () => {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    const sheetName = sheet.getName();
    if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
      ui.alert('Error', 'Please run this function on a valid report sheet.', ui.ButtonSet.OK);
      return;
    }
    
    // Verify prerequisites
    let headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    if (headers.indexOf('Business Perspective') === -1 || headers.indexOf('Technical Perspective') === -1) {
      ui.alert('Missing Prerequisites', 
        'Please run "Generate Epic Perspectives" first!\n\nThis function requires the Business Perspective and Technical Perspective columns to already exist.', 
        ui.ButtonSet.OK);
      return;
    }
    
    ss.toast('Starting initiative-level perspective generation...', '🔍 Analyzing Sheet', 10);
    const headerRow = 4;

    try {
      const COLUMN_MAPPINGS = {
        key: ['Key'],
        parentKey: ['Parent Key'],
        issueType: ['Issue Type'],
        summary: ['Summary'],
        piObjective: ['PI Objective'],
        benefitHypothesis: ['Benefit Hypothesis'],
        acceptanceCriteria: ['Acceptance Criteria'],
        initiativeTitle: ['Initiative Title'],
        initiativeDescription: ['Initiative Description']
      };
      
      const dataValues = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn()).getValues();

      ss.toast('Generating merged perspectives...', '🤖 AI Processing', 5);
      generateMergedInitiativePerspectives(ss, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS);
      
      ss.toast('✅ Initiative-level perspectives complete!', '✅ Success', 10);

    } catch (error) {
      console.error('Initiative Perspective Error:', error);
      ui.alert('An Error Occurred', error.toString(), ui.ButtonSet.OK);
    }
  }, { sheetName: SpreadsheetApp.getActiveSheet().getName() });
}

function summarizePlanReadout() {
  return logActivity('Business & Technical Perspective Summary', () => {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    const sheetName = sheet.getName();
    if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
      ui.alert('Error', 'Please run this function on a valid report sheet.', ui.ButtonSet.OK);
      return;
    }
    
    ss.toast('Starting Plan Readout summary generation with initiative context...', '🔍 Analyzing Sheet', 10);
    const headerRow = 4;

    try {
      const COLUMN_MAPPINGS = {
        key: ['Key'],
        parentKey: ['Parent Key'],
        issueType: ['Issue Type'],
        summary: ['Summary'],
        piObjective: ['PI Objective'],
        benefitHypothesis: ['Benefit Hypothesis'],
        acceptanceCriteria: ['Acceptance Criteria'],
        initiativeTitle: ['Initiative Title'],
        initiativeDescription: ['Initiative Description']
      };
      
      let headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      const dataValues = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn()).getValues();

      // Business Perspective Job
      const businessJob = {
        newColumnHeader: 'Business Perspective',
        summaryPrompt: `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

TASK:
Write an executive micro-summary of the epic's BUSINESS VALUE.

STYLE RULES:
- Maximum 25 words. If your response exceeds 25 words, shorten it.
- First word must be a present-participle verb (e.g., Delivering, Enabling, Modernizing, Establishing, Streamlining, Automating)
- BANNED first words: "Optimizing", "Enhancing", "Improving", "Updating" (too vague)
- Format: [Verb] + [capability/outcome] + [who/why it matters]
- Focus on business outcome: time saved, risk reduced, revenue protected, adoption improved
- Avoid technical implementation details
- BANNED phrases: "This PI will", "In this PI", "This work will focus on", "The team will"
- Ignore any blank fields and work with what is available

EXAMPLE (22 words):
"Delivering automated payment reconciliation for the billing team to reduce manual processing errors and accelerate month-end close by 3 days."`,
        sourceColumns: ['key', 'parentKey', 'summary', 'piObjective', 'benefitHypothesis', 'acceptanceCriteria', 'initiativeTitle', 'initiativeDescription'],
        filterColumn: 'issueType',
        filterValue: 'Epic'
      };
      
      // Technical Perspective Job
      const technicalJob = {
        newColumnHeader: 'Technical Perspective',
        summaryPrompt: `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

TASK:
Write an executive micro-summary of WHAT IS BEING BUILT (technical deliverable).

STYLE RULES:
- Maximum 20 words. If your response exceeds 20 words, shorten it.
- Start with: "The feature...", "The system...", or "The platform..."
- Use exactly one main verb from: builds, integrates, migrates, implements, establishes, automates. Avoid chaining with "and."
- Optional: one short "to..." clause at the end (maximum 6 words)
- Focus on systems, capabilities, integrations, infrastructure — not process or staffing
- BANNED phrases: "Engineers will", "We will", "The team will", "This PI will"
- BANNED words: Epic, Initiative, Story — use "feature", "system", "platform" instead
- Ignore any blank fields and work with what is available

EXAMPLE (16 words):
"The platform integrates OAuth 2.0 authentication with legacy billing systems to enable single sign-on access."`,
        sourceColumns: ['key', 'parentKey', 'summary', 'piObjective', 'benefitHypothesis', 'acceptanceCriteria', 'initiativeTitle', 'initiativeDescription'],
        filterColumn: 'issueType',
        filterValue: 'Epic'
      };
      
      // Execute epic-level jobs
      ss.toast('Step 1/4: Generating Business Perspective (epic-level)...', '🤖 AI Processing', 5);
      generateEnhancedPIReadoutHelper(ss, ui, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, businessJob);
      
      headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      console.log(`Headers after Business Perspective: ${headers.length} columns`);
      
      ss.toast('Step 2/4: Generating Technical Perspective (epic-level)...', '🤖 AI Processing', 5);
      generateEnhancedPIReadoutHelper(ss, ui, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, technicalJob);
      
      headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      console.log(`Headers after Technical Perspective: ${headers.length} columns`);
      
      // Generate merged initiative-level perspectives
      ss.toast('Step 3/4: Generating Merged Business Perspective (initiative-level)...', '🤖 AI Processing', 5);
      generateMergedInitiativePerspectives(ss, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, 'Business');
      
      headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      console.log(`Headers after Merged Business Perspective: ${headers.length} columns`);
      
      ss.toast('Step 4/4: Generating Merged Technical Perspective (initiative-level)...', '🤖 AI Processing', 5);
      generateMergedInitiativePerspectives(ss, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, 'Technical');
      
      ss.toast('✅ Plan Readout generation complete!', '✅ Success', 8);

    } catch (error) {
      console.error('Plan Readout Error:', error);
      ui.alert('An Error Occurred', error.toString(), ui.ButtonSet.OK);
    }
  }, { sheetName: SpreadsheetApp.getActiveSheet().getName() });
}

function generateSingleEpicPrompt(epicKey, contextText, perspectiveType) {
  return buildSingleEpicPrompt(epicKey, contextText, perspectiveType);
}

function generateSingleEpicMergedPerspective(epicKey, contextText, perspectiveType) {
  return callGeminiAPIWithFallback(
    buildSingleEpicPrompt(epicKey, contextText, perspectiveType),
    'N/A*'
  );
}

function generateMergedInitiativePerspectives(ss, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS) {
  
  console.log('\n=== GENERATING MERGED PERSPECTIVES (BATCHED) ===');
  
  const getColIndex = (mappingKey) => {
    const possibleNames = COLUMN_MAPPINGS[mappingKey];
    for (const name of possibleNames) {
      const idx = headers.indexOf(name);
      if (idx !== -1) return idx;
    }
    return -1;
  };
  
  const keyCol = getColIndex('key');
  const parentKeyCol = getColIndex('parentKey');
  const issueTypeCol = getColIndex('issueType');
  
  const summaryCol = getColIndex('summary');
  const epicNameCol = headers.indexOf('Epic Name');
  const piObjectiveCol = getColIndex('piObjective');
  const benefitHypothesisCol = getColIndex('benefitHypothesis');
  const acceptanceCriteriaCol = getColIndex('acceptanceCriteria');
  const initiativeTitleCol = getColIndex('initiativeTitle');
  const initiativeDescCol = getColIndex('initiativeDescription');
  
  const businessPerspectiveCol = headers.indexOf('Business Perspective');
  const technicalPerspectiveCol = headers.indexOf('Technical Perspective');
  
  if (businessPerspectiveCol === -1) {
    throw new Error('Business Perspective column not found! Make sure epic-level perspectives are generated first.');
  }
  if (technicalPerspectiveCol === -1) {
    throw new Error('Technical Perspective column not found! Make sure epic-level perspectives are generated first.');
  }
  
  const mergedBusinessColumnName = 'Merged Business Perspective';
  const mergedTechnicalColumnName = 'Merged Technical Perspective';
  
  let mergedBusinessColNumber;
  let mergedTechnicalColNumber;
  
  const mergedBusinessIdx = headers.indexOf(mergedBusinessColumnName);
  if (mergedBusinessIdx === -1) {
    const currentLastCol = sheet.getLastColumn();
    sheet.insertColumnAfter(currentLastCol);
    mergedBusinessColNumber = currentLastCol + 1;
    sheet.getRange(headerRow, mergedBusinessColNumber)
      .setValue(mergedBusinessColumnName)
      .setFontWeight('bold')
      .setBackground('#9b7bb8');
    headers.push(mergedBusinessColumnName);
    console.log(`Created "${mergedBusinessColumnName}" at column ${mergedBusinessColNumber} (${columnNumberToLetter(mergedBusinessColNumber)})`);
  } else {
    mergedBusinessColNumber = mergedBusinessIdx + 1;
    console.log(`Found "${mergedBusinessColumnName}" at column ${mergedBusinessColNumber} (${columnNumberToLetter(mergedBusinessColNumber)})`);
  }
  
  const mergedTechnicalIdx = headers.indexOf(mergedTechnicalColumnName);
  if (mergedTechnicalIdx === -1) {
    const currentLastCol = sheet.getLastColumn();
    sheet.insertColumnAfter(currentLastCol);
    mergedTechnicalColNumber = currentLastCol + 1;
    sheet.getRange(headerRow, mergedTechnicalColNumber)
      .setValue(mergedTechnicalColumnName)
      .setFontWeight('bold')
      .setBackground('#9b7bb8');
    headers.push(mergedTechnicalColumnName);
    console.log(`Created "${mergedTechnicalColumnName}" at column ${mergedTechnicalColNumber} (${columnNumberToLetter(mergedTechnicalColNumber)})`);
  } else {
    mergedTechnicalColNumber = mergedTechnicalIdx + 1;
    console.log(`Found "${mergedTechnicalColumnName}" at column ${mergedTechnicalColNumber} (${columnNumberToLetter(mergedTechnicalColNumber)})`);
  }
  
  const initiativeGroups = {};
  
  dataValues.forEach((row, rowIndex) => {
    const issueType = row[issueTypeCol];
    const parentKey = row[parentKeyCol];
    
    if (issueType === 'Epic' && parentKey) {
      if (!initiativeGroups[parentKey]) {
        initiativeGroups[parentKey] = [];
      }
      
      const getVal = (colIdx) => {
        return (colIdx >= 0 && row[colIdx]) ? String(row[colIdx]).trim() : '';
      };
      
      initiativeGroups[parentKey].push({
        rowIndex: rowIndex,
        key: getVal(keyCol),
        summary: getVal(summaryCol),
        epicName: getVal(epicNameCol),
        piObjective: getVal(piObjectiveCol),
        benefitHypothesis: getVal(benefitHypothesisCol),
        acceptanceCriteria: getVal(acceptanceCriteriaCol),
        initiativeTitle: getVal(initiativeTitleCol),
        initiativeDescription: getVal(initiativeDescCol)
      });
    }
  });
  
  const initiativeKeys = Object.keys(initiativeGroups);
  console.log(`Processing ${initiativeKeys.length} initiatives with BATCHED AI calls`);
  
  const businessTasks = [];
  const technicalTasks = [];
  
  for (const [parentKey, epics] of Object.entries(initiativeGroups)) {
    const epicRowIndices = epics.map(e => e.rowIndex);
    
    if (epics.length === 1) {
      const epic = epics[0];
      
      let contextText = '';
      if (epic.initiativeTitle) contextText += `Initiative: ${epic.initiativeTitle}\n`;
      if (epic.initiativeDescription) contextText += `Description: ${epic.initiativeDescription}\n`;
      if (epic.epicName) contextText += `Epic: ${epic.epicName}\n`;
      if (epic.summary) contextText += `Summary: ${epic.summary}\n`;
      if (epic.piObjective) contextText += `PI Objective: ${epic.piObjective}\n`;
      if (epic.benefitHypothesis) contextText += `Benefit: ${epic.benefitHypothesis}\n`;
      if (epic.acceptanceCriteria) contextText += `Acceptance: ${epic.acceptanceCriteria}\n`;
      
      if (contextText.length > 50) {
        const businessPrompt = buildSingleEpicPrompt(epic.key, contextText, 'Business');
        const technicalPrompt = buildSingleEpicPrompt(epic.key, contextText, 'Technical');
        
        businessTasks.push({ parentKey, prompt: businessPrompt, epicRowIndices });
        technicalTasks.push({ parentKey, prompt: technicalPrompt, epicRowIndices });
      } else {
        businessTasks.push({ parentKey, prompt: null, directValue: 'N/A*', epicRowIndices });
        technicalTasks.push({ parentKey, prompt: null, directValue: 'N/A*', epicRowIndices });
      }
      
    } else {
      const initiativeTitle = epics[0].initiativeTitle || 'Unknown Initiative';
      const initiativeDescription = epics[0].initiativeDescription || '';
      
      let epicContextDetails = '';
      epics.forEach((epic, idx) => {
        epicContextDetails += `\n=== EPIC ${idx + 1} (${epic.key}) ===\n`;
        if (epic.epicName) epicContextDetails += `Epic Name: ${epic.epicName}\n`;
        if (epic.summary) epicContextDetails += `Summary: ${epic.summary}\n`;
        if (epic.piObjective) epicContextDetails += `PI Objective: ${epic.piObjective}\n`;
        if (epic.benefitHypothesis) epicContextDetails += `Benefit: ${epic.benefitHypothesis}\n`;
        if (epic.acceptanceCriteria) epicContextDetails += `Acceptance: ${epic.acceptanceCriteria}\n`;
      });
      
      const businessPrompt = buildMergedPrompt(initiativeTitle, parentKey, initiativeDescription, epicContextDetails, 'Business');
      const technicalPrompt = buildMergedPrompt(initiativeTitle, parentKey, initiativeDescription, epicContextDetails, 'Technical');
      
      businessTasks.push({ parentKey, prompt: businessPrompt, epicRowIndices });
      technicalTasks.push({ parentKey, prompt: technicalPrompt, epicRowIndices });
    }
  }
  
  console.log(`Built ${businessTasks.length} business tasks and ${technicalTasks.length} technical tasks`);
  
  const businessResults = processBatchedTasks(businessTasks, 'Business', ss);
  const technicalResults = processBatchedTasks(technicalTasks, 'Technical', ss);
  
  console.log('Writing all results to sheet...');
  let writtenCount = 0;
  
  businessTasks.forEach((task, idx) => {
    const businessValue = task.directValue || businessResults.get(task.parentKey) || 'N/A*';
    const technicalValue = technicalTasks[idx].directValue || technicalResults.get(task.parentKey) || 'N/A*';
    
    task.epicRowIndices.forEach(rowIndex => {
      const actualRow = headerRow + 1 + rowIndex;
      sheet.getRange(actualRow, mergedBusinessColNumber).setValue(businessValue);
      sheet.getRange(actualRow, mergedTechnicalColNumber).setValue(technicalValue);
    });
    
    writtenCount++;
    if (writtenCount % 20 === 0) {
      ss.toast(`Written ${writtenCount}/${businessTasks.length} merged perspectives...`, 'Progress', 1);
    }
  });
  
  console.log(`✓ Generated and wrote ${businessTasks.length} merged perspectives (batched processing)`);
}

function buildSingleEpicPrompt(epicKey, contextText, perspectiveType) {
  const header = `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

AVAILABLE DATA:
${contextText}
`;

  if (perspectiveType === 'Business') {
    return `${header}
TASK:
Write an executive micro-summary of the epic's BUSINESS VALUE.

STYLE RULES:
- Maximum 25 words. If your response exceeds 25 words, shorten it.
- First word must be a present-participle verb (e.g., Delivering, Enabling, Modernizing, Establishing, Streamlining, Automating)
- BANNED first words: "Optimizing", "Enhancing", "Improving", "Updating" (too vague)
- Format: [Verb] + [capability/outcome] + [who/why it matters]
- Focus on business outcome: time saved, risk reduced, revenue protected, adoption improved
- Avoid technical implementation details
- BANNED phrases: "This PI will", "In this PI", "This work will focus on", "The team will"
- Ignore any blank fields and work with what is available

EXAMPLE (22 words):
"Delivering automated payment reconciliation for the billing team to reduce manual processing errors and accelerate month-end close by 3 days."`;
  }

  return `${header}
TASK:
Write an executive micro-summary of WHAT IS BEING BUILT (technical deliverable).

STYLE RULES:
- Maximum 20 words. If your response exceeds 20 words, shorten it.
- Start with: "The feature...", "The system...", or "The platform..."
- Use exactly one main verb from: builds, integrates, migrates, implements, establishes, automates. Avoid chaining with "and."
- Optional: one short "to..." clause at the end (maximum 6 words)
- Focus on systems, capabilities, integrations, infrastructure — not process or staffing
- BANNED phrases: "Engineers will", "We will", "The team will", "This PI will"
- BANNED words: Epic, Initiative, Story — use "feature", "system", "platform" instead
- Ignore any blank fields and work with what is available

EXAMPLE (16 words):
"The platform integrates OAuth 2.0 authentication with legacy billing systems to enable single sign-on access."`;
}

function buildMergedPrompt(initiativeTitle, parentKey, initiativeDescription, epicContextDetails, perspectiveType) {
  const header = `RETURN RULES (HARD):
- Return ONLY the summary sentences. No labels, no quotes, no markdown, no extra text.
- Each sentence ends with exactly one period. No other punctuation used as sentence separators.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A`;

  if (perspectiveType === 'Business') {
    return `${header}

INITIATIVE:
Title: ${initiativeTitle}
Key: ${parentKey}
${initiativeDescription ? `Description: ${initiativeDescription}` : ''}

CHILD EPICS:
${epicContextDetails}

TASK:
Synthesize the epics into ONE cohesive initiative-level value proposition.

STYLE RULES:
- 2-3 sentences, maximum 50 words total. If your response exceeds 50 words, shorten it.
- Sentence 1: The business problem or opportunity this initiative addresses
- Sentence 2: What capability is being delivered across the epics
- Sentence 3 (optional): Expected business outcome or impact
- SYNTHESIZE across all epics — do not list them individually
- Start with: "The program..." or "This initiative..."
- BANNED: "Epic 1", "Epic 2", "This PI will"
- Use "program" or "feature" instead of "Epic" or "Initiative"

GOOD: "The authentication program addresses growing security risks across customer-facing platforms. It delivers unified identity management with multi-factor authentication, reducing unauthorized access incidents while improving login experience."

BAD: "Epic 1 implements OAuth. Epic 2 adds MFA. Epic 3 handles permissions."`;
  } else {
    return `${header}

INITIATIVE:
Title: ${initiativeTitle}
Key: ${parentKey}
${initiativeDescription ? `Description: ${initiativeDescription}` : ''}

CHILD EPICS:
${epicContextDetails}

TASK:
Synthesize the epics into ONE cohesive initiative-level technical overview.

STYLE RULES:
- 2-3 sentences, maximum 50 words total. If your response exceeds 50 words, shorten it.
- Describe the OVERALL technical approach and key systems involved
- SYNTHESIZE across all epics — do not list them individually
- Start with: "The program builds..." or "The system delivers..."
- BANNED: "Epic 1", "Epic 2", "This PI will"
- Use "program" or "system" instead of "Epic" or "Initiative"
- Technical language appropriate for executive audience

GOOD: "The program builds a cloud-native authentication framework using OAuth 2.0 and SAML, integrating with existing identity providers to enable unified access management across web and mobile platforms."

BAD: "Epic 1 does OAuth API. Epic 2 builds frontend. Epic 3 adds database."`;
  }
}

function processBatchedTasks(tasks, perspectiveType, ss) {
  const tasksNeedingAI = tasks.filter(t => t.prompt !== null);
  const results = new Map();
  
  if (tasksNeedingAI.length === 0) {
    console.log(`No ${perspectiveType} tasks need AI generation`);
    return results;
  }
  
  const BATCH_SIZE = 25;
  const totalBatches = Math.ceil(tasksNeedingAI.length / BATCH_SIZE);
  
  console.log(`Processing ${tasksNeedingAI.length} ${perspectiveType} tasks in ${totalBatches} batches`);
  
  for (let i = 0; i < tasksNeedingAI.length; i += BATCH_SIZE) {
    const batch = tasksNeedingAI.slice(i, i + BATCH_SIZE);
    const batchPrompts = batch.map(task => task.prompt);
    
    const currentBatchNum = Math.floor(i / BATCH_SIZE) + 1;
    
    ss.toast(
      `Generating ${perspectiveType} perspectives: batch ${currentBatchNum}/${totalBatches} (${batch.length} initiatives)...`,
      '🧠 AI Processing',
      10
    );
    
    console.log(`Processing ${perspectiveType} batch ${currentBatchNum}/${totalBatches}`);
    
    const batchInstruction = getMergedPerspectiveInstructions(perspectiveType);
    const batchResults = batchSummarize(batchPrompts, batchInstruction);
    
    batch.forEach((task, idx) => {
      results.set(task.parentKey, batchResults[idx]);
    });
    
    Utilities.sleep(500);
  }
  
  console.log(`✓ Completed ${perspectiveType} perspective generation`);
  return results;
}

function columnNumberToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function generateMergedPrompt(initiativeTitle, initiativeDescription, epicPerspectives, perspectiveType) {
  const header = `RETURN RULES (HARD):
- Return ONLY the summary sentences. No labels, no quotes, no markdown, no extra text.
- Each sentence ends with exactly one period. No other punctuation used as sentence separators.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A`;

  if (perspectiveType === 'Business') {
    return `${header}

INITIATIVE: ${initiativeTitle}
${initiativeDescription ? `Description: ${initiativeDescription}` : ''}

CHILD EPIC SUMMARIES:
${epicPerspectives}

TASK:
Synthesize these epic-level business summaries into ONE cohesive initiative-level value proposition.

STYLE RULES:
- 2-3 sentences, maximum 50 words total. If your response exceeds 50 words, shorten it.
- Sentence 1: The business problem or opportunity this initiative addresses
- Sentence 2: What capability is being delivered across the epics
- Sentence 3 (optional): Expected business outcome or impact
- SYNTHESIZE across all epics — do not list them individually
- Start with: "The program..." or "This initiative..."
- BANNED: "Epic 1", "Epic 2", "This PI will"
- Use "program" or "feature" instead of "Epic" or "Initiative"

GOOD: "The authentication program addresses growing security risks across customer-facing platforms. It delivers unified identity management with multi-factor authentication, reducing unauthorized access incidents while improving login experience."

BAD: "Epic 1 implements OAuth. Epic 2 adds MFA. Epic 3 handles permissions."`;
  } else {
    return `${header}

INITIATIVE: ${initiativeTitle}
${initiativeDescription ? `Description: ${initiativeDescription}` : ''}

CHILD EPIC SUMMARIES:
${epicPerspectives}

TASK:
Synthesize these epic-level technical summaries into ONE cohesive initiative-level technical overview.

STYLE RULES:
- 2-3 sentences, maximum 50 words total. If your response exceeds 50 words, shorten it.
- Describe the OVERALL technical approach and key systems involved
- SYNTHESIZE across all epics — do not list them individually
- Start with: "The program builds..." or "The system delivers..."
- BANNED: "Epic 1", "Epic 2", "This PI will"
- Use "program" or "system" instead of "Epic" or "Initiative"
- Technical language appropriate for executive audience

GOOD: "The program builds a cloud-native authentication framework using OAuth 2.0 and SAML, integrating with existing identity providers to enable unified access management across web and mobile platforms."

BAD: "Epic 1 does OAuth API. Epic 2 builds frontend. Epic 3 adds database."`;
  }
}

function getMergedPerspectiveInstructions(perspectiveType) {
  if (perspectiveType === 'Business') {
    return `You are creating executive business value propositions. Each input describes an initiative with child epics. Synthesize into ONE cohesive statement: what business problem is addressed, what capability is delivered, and expected impact. Maximum 50 words per response. No bullet points. No "Epic 1/2/3" references. No colons or em dashes to chain ideas. Each sentence ends with exactly one period.`;
  } else {
    return `You are creating executive technical overviews. Each input describes an initiative with child epics. Synthesize into ONE cohesive statement: what systems are being built, key technologies involved, and technical approach. Maximum 50 words per response. No bullet points. No "Epic 1/2/3" references. No colons or em dashes to chain ideas. Each sentence ends with exactly one period.`;
  }
}

function generateEnhancedPIReadoutHelper(ss, ui, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, jobConfig) {
  
  const { newColumnHeader, summaryPrompt, sourceColumns, filterColumn, filterValue } = jobConfig;
  ss.toast(`Processing "${newColumnHeader}" with initiative context...`, '🧠 Thinking...', 10);

  const colMap = {};
  for (const key of sourceColumns) {
    const possibleNames = COLUMN_MAPPINGS[key];
    const index = headers.findIndex(header => possibleNames.includes(header));
    if (index === -1) {
      console.warn(`Column not found: ${possibleNames.join(', ')} - will use empty values`);
      colMap[key] = -1;
    } else {
      colMap[key] = index;
    }
  }
  
  const filterCol = headers.findIndex(header => COLUMN_MAPPINGS[filterColumn].includes(header));
  if (filterCol === -1) {
    throw new Error(`Could not find filter column. Missing: [${COLUMN_MAPPINGS[filterColumn].join(', ')}]`);
  }

  let destCol = headers.indexOf(newColumnHeader);
  if (destCol === -1) {
    destCol = sheet.getLastColumn() + 1;
    sheet.getRange(headerRow, destCol).setValue(newColumnHeader)
      .setFontWeight('bold').setBackground('#9b7bb8');
    sheet.setColumnWidth(destCol, 400);
  } else {
    destCol += 1;
  }

  const summaryTasks = [];
  
  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    
    if (row[filterCol] !== filterValue) continue;
    
    const getVal = (colKey) => {
      const idx = colMap[colKey];
      return (idx >= 0 && row[idx]) ? String(row[idx]).trim() : '';
    };
    
    const epicKey = getVal('key');
    const parentKey = getVal('parentKey');
    const epicSummary = getVal('summary');
    const piObjective = getVal('piObjective');
    const benefitHypothesis = getVal('benefitHypothesis');
    const acceptanceCriteria = getVal('acceptanceCriteria');
    const initiativeTitle = getVal('initiativeTitle');
    const initiativeDescription = getVal('initiativeDescription');
    
    let combinedText = '';
    
    if (parentKey && (initiativeTitle || initiativeDescription)) {
      combinedText += `=== OVERALL INITIATIVE (Parent: ${parentKey}) ===\n`;
      if (initiativeTitle) {
        combinedText += `Initiative: ${initiativeTitle}\n`;
      }
      if (initiativeDescription) {
        combinedText += `Description: ${initiativeDescription}\n`;
      }
      combinedText += `\n=== THIS PI'S EPIC (${epicKey}) ===\n`;
    } else {
      combinedText += `=== EPIC ${epicKey} (No Parent Initiative) ===\n`;
    }
    
    if (epicSummary) {
      combinedText += `Epic Summary: ${epicSummary}\n`;
    }
    if (piObjective) {
      combinedText += `PI Objective: ${piObjective}\n`;
    }
    if (benefitHypothesis) {
      combinedText += `Benefit Hypothesis: ${benefitHypothesis}\n`;
    }
    if (acceptanceCriteria) {
      combinedText += `Acceptance Criteria: ${acceptanceCriteria}\n`;
    }
    
    if (combinedText.length > 100) {
      summaryTasks.push({ 
        text: combinedText, 
        index: i,
        epicKey: epicKey,
        hasInitiative: !!(parentKey && (initiativeTitle || initiativeDescription))
      });
    }
  }

  if (summaryTasks.length === 0) {
    if (ui) {
      ui.alert(`"${newColumnHeader}" Job: No epics with sufficient data were found on this sheet.`);
    }
    return;
  }
  
  console.log(`Found ${summaryTasks.length} epics to summarize (${summaryTasks.filter(t => t.hasInitiative).length} with initiative context)`);

  const BATCH_SIZE = 25;
  const outputRange = sheet.getRange(headerRow + 1, destCol, dataValues.length, 1);
  const outputValues = outputRange.getValues();

  for (let i = 0; i < summaryTasks.length; i += BATCH_SIZE) {
    const batchTasks = summaryTasks.slice(i, i + BATCH_SIZE);
    const batchTexts = batchTasks.map(task => task.text);
    
    const currentBatchNum = Math.floor(i / BATCH_SIZE) + 1;
    const totalBatches = Math.ceil(summaryTasks.length / BATCH_SIZE);
    ss.toast(
      `"${newColumnHeader}": Processing batch ${currentBatchNum}/${totalBatches} (${batchTasks.length} epics)...`, 
      '🧠 Thinking...', 
      15
    );

    const batchSummaries = batchSummarize(batchTexts, summaryPrompt);

    if (batchSummaries.length !== batchTasks.length) {
      console.error(`Batch ${currentBatchNum} mismatch: Expected ${batchTasks.length}, got ${batchSummaries.length}`);
      batchTasks.forEach(task => {
        outputValues[task.index][0] = "Error: AI response mismatch.";
      });
      continue;
    }

    batchSummaries.forEach((summary, j) => {
      const originalIndex = batchTasks[j].index;
      outputValues[originalIndex][0] = summary;
    });
  }

  outputRange.setValues(outputValues);
  outputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  ss.toast(`✅ "${newColumnHeader}" complete!`, '✅ Success', 5);
}

function generateEpicPerspectivesOnSheet(targetSheet) {
  const ss = targetSheet.getParent();
  const sheet = targetSheet;
  
  const sheetName = sheet.getName();
  if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
    console.warn('Not a valid report sheet for epic perspectives');
    return;
  }
  
  console.log(`Starting epic-level perspective generation on sheet: ${sheetName}`);
  const headerRow = 4;

  try {
    const COLUMN_MAPPINGS = {
      key: ['Key'],
      parentKey: ['Parent Key'],
      issueType: ['Issue Type'],
      summary: ['Summary'],
      piObjective: ['PI Objective'],
      benefitHypothesis: ['Benefit Hypothesis'],
      acceptanceCriteria: ['Acceptance Criteria'],
      initiativeTitle: ['Initiative Title'],
      initiativeDescription: ['Initiative Description']
    };
    
    let headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const dataValues = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn()).getValues();

    const businessJob = {
      newColumnHeader: 'Business Perspective',
      summaryPrompt: `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

TASK:
Write an executive micro-summary of the epic's BUSINESS VALUE.

STYLE RULES:
- Maximum 25 words. If your response exceeds 25 words, shorten it.
- First word must be a present-participle verb (e.g., Delivering, Enabling, Modernizing, Establishing, Streamlining, Automating)
- BANNED first words: "Optimizing", "Enhancing", "Improving", "Updating" (too vague)
- Format: [Verb] + [capability/outcome] + [who/why it matters]
- Focus on business outcome: time saved, risk reduced, revenue protected, adoption improved
- Avoid technical implementation details
- BANNED phrases: "This PI will", "In this PI", "This work will focus on", "The team will"
- Ignore any blank fields and work with what is available

EXAMPLE (22 words):
"Delivering automated payment reconciliation for the billing team to reduce manual processing errors and accelerate month-end close by 3 days."`,
      sourceColumns: ['key', 'summary', 'piObjective', 'benefitHypothesis', 'acceptanceCriteria', 'initiativeTitle', 'initiativeDescription'],
      filterColumn: 'issueType',
      filterValue: 'Epic'
    };
    
    generateEnhancedPIReadoutHelper(ss, null, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, businessJob);
    
    headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
    
    const technicalJob = {
      newColumnHeader: 'Technical Perspective',
      summaryPrompt: `RETURN RULES (HARD):
- Return ONLY the final sentence. No labels, no quotes, no markdown, no extra text.
- Exactly ONE sentence. No semicolons. No line breaks.
- Exactly one period at the very end. No other periods in the response.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A

TASK:
Write an executive micro-summary of WHAT IS BEING BUILT (technical deliverable).

STYLE RULES:
- Maximum 20 words. If your response exceeds 20 words, shorten it.
- Start with: "The feature...", "The system...", or "The platform..."
- Use exactly one main verb from: builds, integrates, migrates, implements, establishes, automates. Avoid chaining with "and."
- Optional: one short "to..." clause at the end (maximum 6 words)
- Focus on systems, capabilities, integrations, infrastructure — not process or staffing
- BANNED phrases: "Engineers will", "We will", "The team will", "This PI will"
- BANNED words: Epic, Initiative, Story — use "feature", "system", "platform" instead
- Ignore any blank fields and work with what is available

EXAMPLE (16 words):
"The platform integrates OAuth 2.0 authentication with legacy billing systems to enable single sign-on access."`,
      sourceColumns: ['key', 'summary', 'piObjective', 'benefitHypothesis', 'acceptanceCriteria', 'initiativeTitle', 'initiativeDescription'],
      filterColumn: 'issueType',
      filterValue: 'Epic'
    };
    
    generateEnhancedPIReadoutHelper(ss, null, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS, technicalJob);
    
    console.log('✓ Epic perspectives complete on sheet: ' + sheetName);

  } catch (error) {
    console.error('Epic Perspective Error:', error);
    throw error;
  }
}

function generateInitiativePerspectivesOnSheet(targetSheet) {
  const ss = targetSheet.getParent();
  const sheet = targetSheet;
  
  const sheetName = sheet.getName();
  if (!sheetName.includes('Governance') && !sheetName.match(/PI \d+ - Iteration \d+/)) {
    console.warn('Not a valid report sheet for initiative perspectives');
    return;
  }
  
  let headers = sheet.getRange(4, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
  
  if (headers.indexOf('Business Perspective') === -1 || headers.indexOf('Technical Perspective') === -1) {
    console.error('Missing prerequisites: Business Perspective and Technical Perspective columns required');
    return;
  }
  
  console.log(`Starting initiative-level perspective generation on sheet: ${sheetName}`);
  const headerRow = 4;

  try {
    const COLUMN_MAPPINGS = {
      key: ['Key'],
      parentKey: ['Parent Key'],
      issueType: ['Issue Type'],
      summary: ['Summary'],
      piObjective: ['PI Objective'],
      benefitHypothesis: ['Benefit Hypothesis'],
      acceptanceCriteria: ['Acceptance Criteria'],
      initiativeTitle: ['Initiative Title'],
      initiativeDescription: ['Initiative Description']
    };
    
    const dataValues = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn()).getValues();

    generateMergedInitiativePerspectives(ss, sheet, headers, dataValues, headerRow, COLUMN_MAPPINGS);
    
    console.log('✓ Initiative perspectives complete on sheet: ' + sheetName);

  } catch (error) {
    console.error('Initiative Perspective Error:', error);
    throw error;
  }
}

function generateInitiativePerspective(initiativeKey, initiativeTitle, initiativeDescription, childEpics, perspectiveType) {
  
  const epicSummaries = childEpics.map((e, idx) => 
    `Epic ${idx + 1} (${e.issueKey}): ${e.summary}`
  ).join('\n');
  
  const header = `RETURN RULES (HARD):
- Return ONLY the summary sentences. No labels, no quotes, no markdown, no extra text.
- Each sentence ends with exactly one period. No other punctuation used as sentence separators.
- No colons (:), semicolons (;), or em dashes (—) to chain multiple ideas.
- No bullet characters (-, •, *). No numbered lists.
- If insufficient data, return exactly: N/A`;

  let prompt = '';
  
  if (perspectiveType === 'Business') {
    prompt = `${header}

INITIATIVE: ${initiativeTitle}
Key: ${initiativeKey}
Description: ${initiativeDescription || 'No description provided'}

CHILD EPICS:
${epicSummaries}

TASK:
Synthesize these epic-level details into ONE cohesive initiative-level value proposition.

STYLE RULES:
- 2-3 sentences, maximum 50 words total. If your response exceeds 50 words, shorten it.
- Sentence 1: The business problem or opportunity this initiative addresses
- Sentence 2: What capability is being delivered across the epics
- Sentence 3 (optional): Expected business outcome or impact
- SYNTHESIZE across all epics — do not list them individually
- Start with: "The program..." or "This initiative..."
- BANNED: "Epic 1", "Epic 2", "This PI will"
- Use "program" or "feature" instead of "Epic" or "Initiative"`;
    
  } else {
    prompt = `${header}

INITIATIVE: ${initiativeTitle}
Key: ${initiativeKey}
Description: ${initiativeDescription || 'No description provided'}

CHILD EPICS:
${epicSummaries}

TASK:
Synthesize these epic-level details into ONE cohesive initiative-level technical overview.

STYLE RULES:
- 2-3 sentences, maximum 50 words total. If your response exceeds 50 words, shorten it.
- Describe the OVERALL technical approach and key systems involved
- SYNTHESIZE across all epics — do not list them individually
- Start with: "The program builds..." or "The system delivers..."
- BANNED: "Epic 1", "Epic 2", "This PI will"
- Use "program" or "system" instead of "Epic" or "Initiative"
- Technical language appropriate for executive audience`;
  }
  
  return callGeminiAPIWithFallback(prompt, 'N/A*');
}
