// =================================================================================
// SCRIPT CONFIGURATION - HYBRID VERSION
// Combines FinalPlanReport.gs table layout with PG-Slideformatter.gs dependency handling
// =================================================================================

const SLIDE_CONFIG = {
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 540,
  MARGIN_TOP: 6,
  TITLE_TOP: 55,
  MARGIN_LEFT: 20,
  MARGIN_RIGHT: 20,
  TABLE_START_Y: 108,
  TABLE_END_Y: 396,
  CONTENT_WIDTH: 680,
  
  // Font sizes
  TITLE_FONT_SIZE: 16,
  TITLE_PAGE_NUM_SIZE: 14,
  SECTION_TITLE_SIZE: 11,
  INITIATIVE_SIZE: 11,
  PERSPECTIVE_SIZE: 8,
  VERSION_SIZE: 10,
  STREAM_SIZE: 9,
  VALUE_STREAM_SIZE: 9,
  BODY_FONT_SIZE: 9,
  SMALL_FONT_SIZE: 8,
  EPIC_KEY_SIZE: 6,
  FEATURE_POINTS_SIZE: 6,
  DISCLOSURE_FONT_SIZE: 6,
  PAGE_NUMBER_SIZE: 8,
  RAG_INDICATOR_SIZE: 6,
  EPIC_ITERATION_SIZE: 8,
  DEP_ITERATION_SIZE: 8,
  DEP_ICON_SIZE: 5,
  TIMESTAMP_SIZE: 6,
  INITIATIVE_SEPARATOR_SIZE: 4,
  RISK_LABEL_SIZE: 8,
  
  // Table settings
  MAX_INITIATIVES_PER_SLIDE: 5,
  TABLE_HEADER_HEIGHT: 24,
  TABLE_ROW_MIN_HEIGHT: 60,
  COL_PROGRAM_WIDTH: 50,
  COL_INITIATIVE_WIDTH: 630,
  
  // Indentation
  EPIC_INDENT: '  ',
  DEPENDENCY_INDENT: '      ',
  
  // Colors
  PURPLE: '#663399',
  DARK_PURPLE: '#4A148C',
  LIGHT_PURPLE: '#E8E0FF',
  VALUE_STREAM_PURPLE: '#7a68b3',
  PIPE_PURPLE: '#502d7f',
  TABLE_HEADER_BG: '#c8bdd7',
  TABLE_TITLE_BG: '#f0ebfa',
  ALT_ROW_COLOR: '#f5f0ff',
  ROW_COLOR_WHITE: '#FFFFFF',
  INITIATIVE_PURPLE: '#502d7f',
  ROYAL_BLUE: '#0024ff',
  BLUE: '#0057B8',
  BLACK: '#000000',
  RED: '#D32F2F',
  GRAY: '#BDBDBD',
  LIGHT_GRAY: '#999999',
  DARK_GRAY: '#555555',
  DEP_GRAY: '#777777',
  TIMESTAMP_GRAY: '#666666',
  WHITE: '#FFFFFF',
  MISSING_DATA_HIGHLIGHT: '#F9CB9C',
  MISSING_DATA_TEXT: '#000000',
  
  // Element heights
  EPIC_HEIGHT: 16,
  DEPENDENCY_HEIGHT: 12,
  RISK_NOTE_HEIGHT: 10,
  
  // Fonts
  HEADER_FONT: 'Lato',
  BODY_FONT: 'Lato'
};

const PROGRAM_BG_COLORS = ['#FFFFFF', '#f5f0ff'];

const MISSING_DATA_PLACEHOLDERS = [
  'No Program Initiative',
  'No Initiative',
  'No Initiative Associated',
  'NO RAG NOTE',
  'Iteration unknown',
  'Unnamed Epic',
  'VS Unknown'
];
const INFOSEC_SUBSECTION_ORDER = [
  'Infrastructure Vulnerabilities',
  'AppSec Vulnerabilities',
  'Product InfoSec',
  'Other' 
];
// =================================================================================
// TEMPLATE CONFIGURATION
// =================================================================================

const TEMPLATE_CONFIG = {
  TEMPLATE_ID: '1Pmd4tKSMs_j5kyQHVTheSmsRJE1I9jEhbV6Ha0g8WgI',
  TITLE_SLIDE_INDEX: 0,
  PORTFOLIO_DIVIDER_INDEX: 1,
  TABLE_SLIDE_INDEX: 2,
  END_SLIDE_INDEX: 3,
};

// =================================================================================
// HELPER FUNCTIONS
// =================================================================================
function getEpicInfosecSubsection(epic) {
  if (!epic) return 'Other';
  
  const field = epic['Infosec Subsection'];
  
  // No field value
  if (!field) return 'Other';
  
  // Already a string
  if (typeof field === 'string') {
    return field.trim() || 'Other';
  }
  
  // JIRA returns objects like { value: "Infrastructure Vulnerabilities" }
  if (typeof field === 'object') {
    const extracted = field.value || field.name || '';
    return (typeof extracted === 'string' ? extracted.trim() : String(extracted).trim()) || 'Other';
  }
  
  // Fallback for other types
  return String(field).trim() || 'Other';
}
function isInfosecPortfolio(portfolioName) {
  return portfolioName && portfolioName.toUpperCase().startsWith('INFOSEC');
}

function groupInfosecEpicsBySubsection(portfolioData) {
  const SUBSECTION_ORDER = [
    'Infrastructure Vulnerabilities',
    'AppSec Vulnerabilities', 
    'Product InfoSec'
  ];
  
  // Collect all epics with their subsection
  const epicsBySubsection = {};
  
  Object.entries(portfolioData.programs || {}).forEach(([programName, programData]) => {
    Object.entries(programData.initiatives || {}).forEach(([initTitle, initiative]) => {
      (initiative.epics || []).forEach(epic => {
        const subsection = (epic['Infosec Subsection'] || '').toString().trim() || 'Other';
        
        if (!epicsBySubsection[subsection]) {
          epicsBySubsection[subsection] = {
            programs: {},
            totalFeaturePoints: 0
          };
        }
        
        // Preserve original program/initiative structure within subsection
        if (!epicsBySubsection[subsection].programs[programName]) {
          epicsBySubsection[subsection].programs[programName] = { initiatives: {} };
        }
        
        if (!epicsBySubsection[subsection].programs[programName].initiatives[initTitle]) {
          epicsBySubsection[subsection].programs[programName].initiatives[initTitle] = {
            name: initiative.name || initTitle,
            linkKey: initiative.linkKey,
            fixVersion: initiative.fixVersion || '',
            businessPerspective: initiative.businessPerspective || '',
            technicalPerspective: initiative.technicalPerspective || '',
            epics: [],
            totalFeaturePoints: 0
          };
        }
        
        epicsBySubsection[subsection].programs[programName].initiatives[initTitle].epics.push(epic);
        const fp = parseFloat(epic['Feature Points']) || 0;
        epicsBySubsection[subsection].programs[programName].initiatives[initTitle].totalFeaturePoints += fp;
        epicsBySubsection[subsection].totalFeaturePoints += fp;
      });
    });
  });
  
  // Sort subsections in priority order
  const sortedSubsections = Object.keys(epicsBySubsection).sort((a, b) => {
    const indexA = SUBSECTION_ORDER.indexOf(a);
    const indexB = SUBSECTION_ORDER.indexOf(b);
    if (indexA !== -1 && indexB !== -1) return indexA - indexB;
    if (indexA !== -1) return -1;
    if (indexB !== -1) return 1;
    return a.localeCompare(b);
  });
  
  // Build ordered result
  const result = {};
  sortedSubsections.forEach(subsec => {
    result[subsec] = epicsBySubsection[subsec];
  });
  
  console.log(`🔐 INFOSEC grouped into ${Object.keys(result).length} subsections:`);
  Object.entries(result).forEach(([subsec, data]) => {
    console.log(`   📋 ${subsec}: ${data.totalFeaturePoints} FP`);
  });
  
  return result;
}
function createInfosecSubsectionTableSlides(presentation, tableTemplate, basePortfolioName, subsectionName, subsectionData, showDependencies, insertPosition, timestampText, noInitiativeMode, hideSameTeamDeps) {
  const slideTitle = `${basePortfolioName} - ${subsectionName}`;
  
  // Build program rows from subsection data
  const programRows = [];
  const sortedPrograms = Object.keys(subsectionData.programs || {}).sort();
  
  sortedPrograms.forEach(programName => {
    const programData = subsectionData.programs[programName];
    const initiatives = [];
    
    Object.keys(programData.initiatives || {}).sort().forEach(initTitle => {
      initiatives.push({
        title: initTitle,
        ...programData.initiatives[initTitle]
      });
    });
    
    if (initiatives.length > 0) {
      programRows.push({
        programName: programName,
        initiatives: initiatives
      });
    }
  });
  
  if (programRows.length === 0) {
    console.log(`  ⚠️ Subsection "${subsectionName}" has no content, skipping`);
    return 0;
  }
  
  // Build content lines
  const contentLines = buildContentLinesFromPrograms(programRows, showDependencies, noInitiativeMode);
  
  if (contentLines.length === 0) {
    console.log(`  ⚠️ Subsection "${subsectionName}" has no content lines, skipping`);
    return 0;
  }
  
  console.log(`  🔐 Creating slides for "${slideTitle}" - ${contentLines.length} content lines`);
  
  // Create table slides for this subsection
  const MAX_LINES_PER_SLIDE = 24;
  const ROLL_THRESHOLD = 22;
  
  const tableSlides = [];
  let slidesCreated = 0;
  let currentSlide = null;
  let currentTable = null;
  let currentLineCount = 0;
  let programColorIndex = 0;
  let lastProgramName = null;
  
  let currentContext = {
    program: null,
    programBgColor: null,
    initiative: null,
    valueStream: null
  };
  
  function createNewTableSlide() {
    const newSlide = presentation.insertSlide(insertPosition + slidesCreated, tableTemplate);
    updateTableSlideHeader(newSlide, slideTitle);  // Use subsection-specific title
    addTimestampToSlide(newSlide, timestampText);
    const newTable = getOrCreateTable(newSlide);
    tableSlides.push(newSlide);
    slidesCreated++;
    currentLineCount = 0;
    console.log(`    -> Created table slide #${tableSlides.length} for "${slideTitle}"`);
    return { slide: newSlide, table: newTable };
  }
  
  let i = 0;
  let needsContinuation = false;
  
  while (i < contentLines.length) {
    const item = contentLines[i];
    
    if (item.program !== lastProgramName) {
      if (lastProgramName !== null) {
        programColorIndex++;
      }
      lastProgramName = item.program;
    }
    const bgColor = PROGRAM_BG_COLORS[programColorIndex % 2];
    
    if (!currentSlide || !currentTable) {
      const result = createNewTableSlide();
      currentSlide = result.slide;
      currentTable = result.table;
    }
    
    const isNewMajorElement = item.type === 'initiative' || item.type === 'valueStream' || item.type === 'epic';
    const isNearEndOfSlide = currentLineCount >= ROLL_THRESHOLD;
    const shouldRollForNewContent = isNearEndOfSlide && isNewMajorElement && 
      (item.program !== currentContext.program || item.type === 'initiative' || item.type === 'valueStream');
    
    if (currentLineCount + item.lines > MAX_LINES_PER_SLIDE || shouldRollForNewContent) {
      const result = createNewTableSlide();
      currentSlide = result.slide;
      currentTable = result.table;
      needsContinuation = true;
    }
    
    const programBatch = [];
    const batchStartProgram = item.program;
    let batchLineCount = 0;
    
    let continuationLineCount = 0;
    if (needsContinuation && currentContext.program === batchStartProgram) {
      continuationLineCount = 1;
      if (currentContext.initiative) {
        continuationLineCount += 1;
        if (currentContext.valueStream) {
          continuationLineCount += 1;
        }
      }
    }
    
    while (i < contentLines.length && 
           contentLines[i].program === batchStartProgram &&
           currentLineCount + continuationLineCount + batchLineCount + contentLines[i].lines <= MAX_LINES_PER_SLIDE) {
      
      const nextItem = contentLines[i];
      const nextIsNewMajor = nextItem.type === 'initiative' || nextItem.type === 'valueStream';
      const wouldBeNearEnd = currentLineCount + continuationLineCount + batchLineCount >= ROLL_THRESHOLD;
      
      if (wouldBeNearEnd && nextIsNewMajor && batchLineCount > 0) {
        break;
      }
      
      programBatch.push(contentLines[i]);
      batchLineCount += contentLines[i].lines;
      i++;
    }
    
    if (programBatch.length > 0) {
      const isSameProgram = currentContext.program === batchStartProgram;
      const displayName = isSameProgram ? `${batchStartProgram} (continued)` : batchStartProgram;
      
      const continuationContext = (needsContinuation && isSameProgram) ? {
        initiative: currentContext.initiative,
        valueStream: currentContext.valueStream
      } : null;
      
      const batchInitiatives = buildInitiativesFromContentBatch(programBatch, showDependencies, continuationContext);
      
      if (batchInitiatives.length > 0) {
        const totalLines = batchLineCount + (continuationContext ? continuationLineCount : 0);
        addProgramRow(currentTable, displayName, batchInitiatives, bgColor, showDependencies, noInitiativeMode, hideSameTeamDeps, true);
        currentLineCount += totalLines;
      }
      
      const lastItem = programBatch[programBatch.length - 1];
      currentContext.program = lastItem.program;
      currentContext.programBgColor = bgColor;
      currentContext.initiative = lastItem.initiativeObj || null;
      currentContext.valueStream = lastItem.valueStream || null;
      
      needsContinuation = false;
    }
  }
  
  // Add page numbers for this subsection's slides
  const totalTableSlides = tableSlides.length;
  tableSlides.forEach((slide, index) => {
    addPortfolioPageNumber(slide, index + 1, totalTableSlides);
  });
  
  console.log(`    ✓ Created ${slidesCreated} table slides for "${slideTitle}"`);
  return slidesCreated;
}
function isMissingDataPlaceholder(value) {
  if (!value) return false;
  const trimmed = String(value).trim();
  return MISSING_DATA_PLACEHOLDERS.includes(trimmed);
}

function showProgress(message, title = '📊 Generating Slides', duration = 30) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title, duration);
  } catch (e) {
    console.log(`Progress: ${title} - ${message}`);
  }
}

function formatDateLong(date) {
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 
                  'July', 'August', 'September', 'October', 'November', 'December'];
  return `${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

function formatTimestamp(date) {
  if (!date) date = new Date();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `Generated: ${month}/${day}/${year} ${hours}:${minutes}`;
}

function removeLeadingNumber(programName) {
  return programName.replace(/^\d+[\s\-\.]+/, '').trim();
}

function getJiraUrl(issueKey) {
  if (!issueKey || issueKey.toString().trim() === '') {
    return null;
  }
  return `https://modmedrnd.atlassian.net/browse/${issueKey.toString().trim()}`;
}

function getRagIndicator(ragValue) {
  if (!ragValue) return '';
  const rag = String(ragValue).toUpperCase();
  if (rag.includes('GREEN')) return '';
  if (rag.includes('AMBER') || rag.includes('YELLOW')) return '🟡';
  if (rag.includes('RED')) return '🔴';
  return '';
}

/**
 * Get display text for iteration field
 * Extracts "Iteration X" from full iteration name
 */
function getIterationDisplayText(iterationValue) {
  if (!iterationValue) return null;
  
  const iterStr = String(iterationValue).trim();
  if (!iterStr) return null;
  
  // If it already looks like "Iteration X", return as-is
  if (iterStr.match(/^Iteration\s+\d+$/i)) {
    return iterStr;
  }
  
  // Try to extract iteration number from various formats
  // Format: "PI 13 - Iteration 1" or "PI13-Iteration1" or just "1"
  const iterMatch = iterStr.match(/Iteration\s*(\d+)/i);
  if (iterMatch) {
    return `Iteration ${iterMatch[1]}`;
  }
  
  // If just a number
  if (iterStr.match(/^\d+$/)) {
    return `Iteration ${iterStr}`;
  }
  
  // Return the original value if we can't parse it
  return iterStr;
}

/**
 * Get display value for a field, handling missing/empty values
 * @param {*} value - The field value
 * @param {string} fieldName - Name of the field (for logging)
 * @param {string} defaultValue - Default value if missing
 * @returns {string|null} - The display value or null if should be excluded
 */
function getFieldDisplayValue(value, fieldName, defaultValue) {
  if (value === null || value === undefined) {
    return defaultValue || null;
  }
  
  const strValue = String(value).trim();
  
  // Empty string
  if (!strValue) {
    return defaultValue || null;
  }
  
  // Check for placeholder/exclusion values
  const excludeValues = ['n/a', 'na', 'none', 'null', 'undefined', '-', '--'];
  if (excludeValues.includes(strValue.toLowerCase())) {
    return defaultValue || null;
  }
  
  return strValue;
}

function getJiraRefreshTimestamp(piNumber, iterationNumber) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return new Date().toLocaleString();
    }
    
    const row3Value = sheet.getRange(3, 1).getValue();
    if (!row3Value) {
      return new Date().toLocaleString();
    }
    
    const timestampStr = String(row3Value);
    const colonIndex = timestampStr.indexOf(':');
    
    if (colonIndex !== -1 && colonIndex < timestampStr.length - 1) {
      const dateTimePart = timestampStr.substring(colonIndex + 1).trim();
      console.log(`✓ Extracted JIRA refresh timestamp: ${dateTimePart}`);
      return dateTimePart;
    }
    
    return timestampStr;
  } catch (error) {
    console.error(`❌ Error getting JIRA refresh timestamp: ${error.message}`);
    return new Date().toLocaleString();
  }
}

// =================================================================================
// MAIN PRESENTATION FUNCTION
// =================================================================================

/**
 * Main function to create governance slides
 * @param {string} presentationName - Name for the presentation
 * @param {Object} data - Epic and dependency data
 * @param {boolean} showDependencies - Whether to show dependencies
 * @param {string} noInitiativeMode - 'show', 'hide', or 'skip'
 * @param {boolean} hideSameTeamDeps - Whether to hide same-team dependencies
 * @param {string|null} singlePortfolio - If provided, only generate slides for this portfolio
 */
function createFormattedPresentation(presentationName, data, showDependencies = true, noInitiativeMode = 'show', hideSameTeamDeps = false, singlePortfolio = null, valueStreamFilter = null, groupByFixVersion = true) {
  // noInitiativeMode options:
  // 'show' (default) - Show "No Initiative" highlighted
  // 'hide' - Hide epics with no initiative entirely
  // 'skip' - Skip initiative header but show value streams/epics directly
  try {
    // Store generation timestamp for slides
    const generatedTimestamp = new Date();
    
    // Step 1: Copy the template
    showProgress('Step 1 of 5: Copying template...', '📊 Starting');
    const templateFile = DriveApp.getFileById(TEMPLATE_CONFIG.TEMPLATE_ID);
    const copiedFile = templateFile.makeCopy(presentationName);
    const presentation = SlidesApp.openById(copiedFile.getId());
    console.log(`Created presentation from template: ${presentationName} (ID: ${presentation.getId()})`);
    
    // Step 2: Update title slide
    showProgress('Step 2 of 5: Updating title slide...', '📊 Preparing');
    updateTitleSlide(presentation, data.metadata);
    
    // Step 3: Process data
    showProgress('Step 3 of 5: Processing epic and dependency data...', '📊 Processing Data');
    const hideNoInitiative = noInitiativeMode === 'hide';
    const portfolioData = processDataForTableSlides(data, showDependencies, hideNoInitiative);
    
    // Step 4: Build slides for each portfolio using template
    let sortedPortfolios = getFilteredAndSortedPortfolios(Object.keys(portfolioData));
    
    // NEW: Filter to single portfolio if specified
    if (singlePortfolio) {
      sortedPortfolios = sortedPortfolios.filter(p => p === singlePortfolio);
      if (sortedPortfolios.length === 0) {
        throw new Error(`Portfolio "${singlePortfolio}" not found. Available portfolios: ${Object.keys(portfolioData).join(', ')}`);
      }
      console.log(`Single portfolio mode: generating slides only for "${singlePortfolio}"`);
    }
    
    const totalPortfolios = sortedPortfolios.length;
    console.log(`Starting to build slides for ${totalPortfolios} portfolios...`);
        // Filter by Value Stream if specified
    if (valueStreamFilter && valueStreamFilter.length > 0) {
      console.log(`Value Stream filter: ${valueStreamFilter.join(', ')}`);
      // Filter epics within each portfolio to only include matching value streams
      Object.keys(portfolioData).forEach(portfolioName => {
        const portfolio = portfolioData[portfolioName];
        Object.keys(portfolio).forEach(programKey => {
          const program = portfolio[programKey];
          if (program.initiatives) {
            program.initiatives.forEach(initiative => {
              if (initiative.epics) {
                initiative.epics = initiative.epics.filter(epic => {
                  const epicVS = epic['Value Stream/Org'] || '';
                  return valueStreamFilter.some(vs => epicVS.toLowerCase().includes(vs.toLowerCase()));
                });
              }
            });
          }
        });
      });
    }
    
    // Get template slides (before we start duplicating)
    const dividerTemplate = presentation.getSlides()[TEMPLATE_CONFIG.PORTFOLIO_DIVIDER_INDEX];
    const tableTemplate = presentation.getSlides()[TEMPLATE_CONFIG.TABLE_SLIDE_INDEX];
    
    // Track where to insert new slides (after End of Governance slide position)
    // We'll insert before the End slide, then move End slide to the end
    let insertPosition = TEMPLATE_CONFIG.END_SLIDE_INDEX; // Start inserting at position 3 (0-indexed)
    
    sortedPortfolios.forEach((portfolioName, index) => {
      const progress = Math.round(((index + 1) / totalPortfolios) * 100);
      showProgress(
        `Step 4 of 5: Building portfolio ${index + 1} of ${totalPortfolios} (${progress}%)\n${portfolioName}`,
        '📊 Building Slides'
      );
      console.log(`[${index + 1}/${totalPortfolios}] Processing Portfolio: ${portfolioName}`);
      
      let slidesCreated = 0;
      
      if (isInfosecPortfolio(portfolioName)) {
        console.log(`  🔐 INFOSEC portfolio detected - using subsection grouping`);
        
        const subsectionGroups = groupInfosecEpicsBySubsection(portfolioData[portfolioName]);
        const subsectionNames = Object.keys(subsectionGroups);
        
        if (subsectionNames.length === 0) {
          console.log(`  ⚠️ No INFOSEC subsections with data - skipping portfolio`);
        } else {
          console.log(`  🔐 Found ${subsectionNames.length} subsections with data`);
          
          // Step 1: Create ONE divider slide for "INFOSEC"
          const dividerSlide = presentation.insertSlide(insertPosition + slidesCreated, dividerTemplate);
          updatePortfolioDividerSlide(dividerSlide, portfolioName, portfolioData[portfolioName].totalFeaturePoints);
          slidesCreated++;
          console.log(`  🔐 Created single INFOSEC divider slide`);
          
          // Step 2: Create table slides for each subsection (no dividers)
          subsectionNames.forEach((subsectionName, subIndex) => {
            const subsectionData = subsectionGroups[subsectionName];
            const subsectionTitle = `INFOSEC - ${subsectionName}`;
            
            console.log(`  🔐 [${subIndex + 1}/${subsectionNames.length}] Creating table slides for: ${subsectionTitle}`);
            
            const tableSlides = createPortfolioTableSlides(
              presentation, 
              tableTemplate, 
              subsectionTitle,
              subsectionData,
              data.dependencies, 
              showDependencies,
              insertPosition + slidesCreated,
              generatedTimestamp,
              'skip',
              hideSameTeamDeps,
              groupByFixVersion
            );
            
            slidesCreated += tableSlides;
            console.log(`    📄 Created ${tableSlides} table slides for ${subsectionTitle}`);
          });
        }
      } else {
        // ═══════════════════════════════════════════════════════════════════════
        // NON-INFOSEC: Standard processing
        // ═══════════════════════════════════════════════════════════════════════
        slidesCreated = createPortfolioSlidesFromTemplate(
          presentation, 
          dividerTemplate, 
          tableTemplate, 
          portfolioName, 
          portfolioData[portfolioName], 
          data.dependencies, 
          showDependencies,
          insertPosition,
          generatedTimestamp,
          noInitiativeMode,
          hideSameTeamDeps
        );
      }
      
      insertPosition += slidesCreated;
    });
    
    // Step 5: Clean up - delete template slides and format guidelines
    showProgress('Step 5 of 5: Finalizing presentation...', '📊 Finishing Up');
    cleanupTemplateSlides(presentation, totalPortfolios);
    
    const url = presentation.getUrl();
    const slideCount = presentation.getSlides().length;
    console.log(`Presentation complete: ${url}`);
    
    // Show completion
    showProgress(
      `✅ Complete! Created ${slideCount} slides across ${totalPortfolios} portfolios.`,
      '🎉 Success',
      10
    );
    
    return {
      success: true,
      url: url,
      presentationId: presentation.getId(),
      slideCount: slideCount
    };
  } catch (error) {
    showProgress(`❌ Error: ${error.message}`, '⚠️ Failed', 10);
    console.error('Error creating presentation:', error);
    throw error;
  }
}

// =================================================================================
// SLIDE CREATION FUNCTIONS
// =================================================================================

function addTimestampToSlide(slide, timestampText) {
  if (!timestampText) return false;
  
  try {
    const xPos = 520;
    const yPos = 12;
    const width = 180;
    const height = 16;
    
    const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, xPos, yPos, width, height);
    shape.getFill().setTransparent();
    shape.getBorder().setTransparent();
    
    const textRange = shape.getText();
    textRange.setText(timestampText);
    textRange.getTextStyle()
      .setFontFamily('Lato')
      .setFontSize(7)
      .setForegroundColor('#666666')
      .setBold(false)
      .setItalic(false);
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
    
    return true;
  } catch (e) {
    console.error(`❌ Failed to add timestamp: ${e.message}`);
    return false;
  }
}

function updateTitleSlide(presentation, metadata) {
  const titleSlide = presentation.getSlides()[TEMPLATE_CONFIG.TITLE_SLIDE_INDEX];
  const shapes = titleSlide.getShapes();
  
  const piNumber = metadata.piNumber;
  const iterationNumber = metadata.iterationNumber;
  const jiraTimestamp = getJiraRefreshTimestamp(piNumber, iterationNumber);
  
  shapes.forEach(shape => {
    if (shape.getText) {
      const text = shape.getText().asString();
      
      if (text.includes('{{PI#}}')) {
        shape.getText().replaceAllText('{{PI#}}', piNumber.toString());
      }
      
      if (text.includes('{{Iter#}}')) {
        shape.getText().replaceAllText('{{Iter#}}', iterationNumber.toString());
      }
      
      if (text.includes('{{Date Ran}}')) {
        shape.getText().replaceAllText('{{Date Ran}}', jiraTimestamp);
        try {
          shape.getText().getTextStyle()
            .setFontSize(4)
            .setForegroundColor('#999999')
            .setFontFamily('Lato')
            .setBold(false)
            .setItalic(false);
        } catch (e) {
          console.warn(`Could not style Date Ran on title: ${e.message}`);
        }
      }
    }
  });
  
  console.log(`✓ Updated title slide: PI ${piNumber}, Iteration ${iterationNumber}, JIRA Data: ${jiraTimestamp}`);
}

function updateTableSlideHeader(slide, portfolioName, jiraTimestamp) {
  const shapes = slide.getShapes();
  
  shapes.forEach(shape => {
    if (shape.getText) {
      const text = shape.getText().asString();
      
      if (text.includes('{{Portfolio Initiative}}')) {
        shape.getText().replaceAllText('{{Portfolio Initiative}}', portfolioName);
      }
      
      if (text.includes('{{ # of #}}')) {
        shape.getText().replaceAllText('{{ # of #}}', '');
      }
      
      if (text.includes('{{page # of #}}')) {
        shape.getText().replaceAllText('{{page # of #}}', '');
      }
      
      if (text.includes('{{Date Ran}}')) {
        const dateText = jiraTimestamp || new Date().toLocaleString();
        shape.getText().replaceAllText('{{Date Ran}}', dateText);
        try {
          shape.getText().getTextStyle()
            .setFontSize(4)
            .setForegroundColor('#999999')
            .setFontFamily('Lato')
            .setBold(false)
            .setItalic(false);
        } catch (e) {
          console.warn(`Could not style Date Ran: ${e.message}`);
        }
      }
    }
  });
}

function updatePortfolioDividerSlide(slide, portfolioName, featurePoints) {
  const shapes = slide.getShapes();
  
  shapes.forEach(shape => {
    if (shape.getText) {
      const text = shape.getText().asString();
      
      if (text.includes('{{Portfolio Initiative}}')) {
        const newText = `${portfolioName}\nFeature Points: ${featurePoints}`;
        shape.getText().setText(newText);
        
        const textRange = shape.getText();
        const fullText = textRange.asString();
        
        const line1End = portfolioName.length;
        const line1Range = textRange.getRange(0, line1End);
        line1Range.getTextStyle()
          .setBold(true)
          .setFontSize(44)
          .setForegroundColor('#FFFFFF');
        
        if (fullText.length > line1End + 1) {
          const line2Start = line1End + 1;
          const line2Range = textRange.getRange(line2Start, fullText.length);
          line2Range.getTextStyle()
            .setBold(false)
            .setFontSize(18)
            .setItalic(true)
            .setForegroundColor('#FFFFFF');
        }
      }
    }
  });
}

function addPortfolioPageNumber(slide, pageNum, totalPages) {
  const shapes = slide.getShapes();
  let foundPlaceholder = false;
  
  shapes.forEach(shape => {
    if (shape.getText) {
      const text = shape.getText().asString();
      const left = shape.getLeft();
      const top = shape.getTop();
      
      if (left > 600 && top < 80 && text.trim() === '') {
        shape.getText().setText(`${pageNum} of ${totalPages}`);
        shape.getText().getTextStyle()
          .setFontSize(6)
          .setItalic(true)
          .setForegroundColor(SLIDE_CONFIG.PURPLE);
        shape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
        foundPlaceholder = true;
      }
    }
  });
  
  if (!foundPlaceholder) {
    const pageNumBox = slide.insertTextBox(
      `${pageNum} of ${totalPages}`,
      SLIDE_CONFIG.SLIDE_WIDTH - 70,
      38,
      50,
      20
    );
    pageNumBox.getText().getTextStyle()
      .setFontSize(6)
      .setItalic(true)
      .setForegroundColor(SLIDE_CONFIG.PURPLE);
    pageNumBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);
  }
}

function getOrCreateTable(slide) {
  const existingTables = slide.getTables();
  
  if (existingTables.length > 0) {
    const table = existingTables[0];
    
    while (table.getNumRows() > 1) {
      table.getRow(table.getNumRows() - 1).remove();
    }
    
    const tableTop = table.getTop();
    console.log(`Using template table at top=${tableTop}, height=${table.getHeight()}, bottom=${tableTop}`);
    
    return table;
  }
  
  console.log(`No template table found, creating new one`);
  return createInitiativeTable(slide);
}

function createInitiativeTable(slide) {
  const tableTop = 108;
  const tableWidth = SLIDE_CONFIG.CONTENT_WIDTH;
  
  const table = slide.insertTable(1, 2, SLIDE_CONFIG.MARGIN_LEFT, tableTop, tableWidth, SLIDE_CONFIG.TABLE_HEADER_HEIGHT);
  
  try {
    const col0 = table.getColumn(0);
    const col1 = table.getColumn(1);
    if (col0 && col0.setWidth) col0.setWidth(SLIDE_CONFIG.COL_PROGRAM_WIDTH);
    if (col1 && col1.setWidth) col1.setWidth(SLIDE_CONFIG.COL_INITIATIVE_WIDTH);
  } catch (e) {
    console.log(`Note: Could not set column widths (${e.message})`);
  }
  
  const headerRow = table.getRow(0);
  
  const cell0 = headerRow.getCell(0);
  cell0.getFill().setSolidFill(SLIDE_CONFIG.TABLE_HEADER_BG);
  cell0.getText().setText('Program Initiative');
  cell0.getText().getTextStyle()
    .setForegroundColor(SLIDE_CONFIG.BLACK)
    .setBold(true)
    .setFontSize(SLIDE_CONFIG.SECTION_TITLE_SIZE)
    .setFontFamily(SLIDE_CONFIG.HEADER_FONT);
  
  const cell1 = headerRow.getCell(1);
  cell1.getFill().setSolidFill(SLIDE_CONFIG.TABLE_HEADER_BG);
  cell1.getText().setText('PI Objective');
  cell1.getText().getTextStyle()
    .setForegroundColor(SLIDE_CONFIG.BLACK)
    .setBold(true)
    .setFontSize(SLIDE_CONFIG.SECTION_TITLE_SIZE)
    .setFontFamily(SLIDE_CONFIG.HEADER_FONT);
  
  return table;
}

// =================================================================================
// PORTFOLIO SLIDE CREATION
// =================================================================================

function createPortfolioSlidesFromTemplate(presentation, dividerTemplate, tableTemplate, portfolioName, portfolioData, allDependencies, showDependencies, insertPosition, generatedTimestamp, noInitiativeMode = 'show', hideSameTeamDeps = false) {
  if (!portfolioData || !portfolioData.programs || Object.keys(portfolioData.programs).length === 0) {
    console.log(`Portfolio "${portfolioName}" has no data, skipping`);
    return 0;
  }
  
  let slidesCreated = 0;
  
  // Format timestamp for display
  const timestampText = formatTimestamp(generatedTimestamp);
  console.log(`📅 Timestamp text formatted: "${timestampText}"`);
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // INFOSEC SPECIAL HANDLING: Group by subsection with separate slide titles
  // ═══════════════════════════════════════════════════════════════════════════════
  if (isInfosecPortfolio(portfolioName)) {
    console.log(`\n🔐 INFOSEC Portfolio detected: "${portfolioName}" - using subsection grouping`);
    
    // Group epics by subsection
    const subsectionGroups = groupInfosecEpicsBySubsection(portfolioData);
    
    if (Object.keys(subsectionGroups).length === 0) {
      console.log(`⚠️ INFOSEC portfolio "${portfolioName}" has no subsections, skipping`);
      return 0;
    }
    
    // Create ONE divider slide for the whole INFOSEC portfolio
    const dividerSlide = presentation.insertSlide(insertPosition, dividerTemplate);
    updatePortfolioDividerSlide(dividerSlide, portfolioName, portfolioData.totalFeaturePoints);
    slidesCreated++;
    
    // Create table slides for each subsection with its own title and page numbers
    Object.entries(subsectionGroups).forEach(([subsectionName, subsectionData]) => {
      const subsectionSlides = createInfosecSubsectionTableSlides(
        presentation,
        tableTemplate,
        portfolioName,
        subsectionName,
        subsectionData,
        showDependencies,
        insertPosition + slidesCreated,
        timestampText,
        noInitiativeMode,
        hideSameTeamDeps
      );
      slidesCreated += subsectionSlides;
    });
    
    console.log(`✓ INFOSEC complete: ${slidesCreated} total slides`);
    return slidesCreated;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // STANDARD PORTFOLIO HANDLING (non-INFOSEC)
  // ═══════════════════════════════════════════════════════════════════════════════
  
  // 2. Build program data for table slides FIRST (before creating divider)
  const programRows = [];
  const sortedPrograms = Object.keys(portfolioData.programs).sort();
  
  sortedPrograms.forEach(programName => {
    const programData = portfolioData.programs[programName];
    const initiatives = [];
    
    Object.keys(programData.initiatives).sort().forEach(initTitle => {
      initiatives.push({
        title: initTitle,
        ...programData.initiatives[initTitle]
      });
    });
    
    if (initiatives.length > 0) {
      programRows.push({
        programName: programName,
        initiatives: initiatives
      });
    }
  });
  
  const totalInitiatives = programRows.reduce((sum, row) => sum + row.initiatives.length, 0);
  console.log(`Portfolio "${portfolioName}": ${programRows.length} programs, ${totalInitiatives} initiatives`);
  
  // Build content lines and check if there's actual content
  const contentLines = buildContentLinesFromPrograms(programRows, showDependencies, noInitiativeMode);
  console.log(`Total content lines for "${portfolioName}": ${contentLines.length}`);
  
  // Skip portfolio entirely if no content to display
  if (programRows.length === 0 || contentLines.length === 0) {
    console.log(`⚠️ Portfolio "${portfolioName}" has no content to display, skipping entirely`);
    return 0;
  }
  
  // 1. NOW create portfolio divider slide (only if we have content)
  const dividerSlide = presentation.insertSlide(insertPosition, dividerTemplate);
  updatePortfolioDividerSlide(dividerSlide, portfolioName, portfolioData.totalFeaturePoints);
  slidesCreated++;
  
  // =============================================================================
  // LINE-BASED PAGINATION: Max 24 lines per slide
  // =============================================================================
  const MAX_LINES_PER_SLIDE = 24;
  
  const tableSlides = [];
  let currentSlide = null;
  let currentTable = null;
  let currentLineCount = 0;
  let programColorIndex = 0;
  let lastProgramName = null;
  
  // Track context for continuation headers
  let currentContext = {
    program: null,
    programBgColor: null,
    initiative: null,
    valueStream: null
  };
  
  // Helper to create a new table slide
  function createNewTableSlide() {
    const newSlide = presentation.insertSlide(insertPosition + slidesCreated, tableTemplate);
    updateTableSlideHeader(newSlide, portfolioName);
    
    // Add timestamp to slide
    addTimestampToSlide(newSlide, timestampText);
    
    const newTable = getOrCreateTable(newSlide);
    tableSlides.push(newSlide);
    slidesCreated++;
    currentLineCount = 0;
    console.log(`  -> Created new table slide #${tableSlides.length} for "${portfolioName}"`);
    return { slide: newSlide, table: newTable };
  }
  
  // Process content lines
  let i = 0;
  let needsContinuation = false;
  
  // Define threshold for "near end of slide" - if starting new major element on line 23+, roll to next
  const ROLL_THRESHOLD = 22;
  
  while (i < contentLines.length) {
    const item = contentLines[i];
    
    // Track program changes for alternating colors
    if (item.program !== lastProgramName) {
      if (lastProgramName !== null) {
        programColorIndex++;
      }
      lastProgramName = item.program;
    }
    const bgColor = PROGRAM_BG_COLORS[programColorIndex % 2];
    
    // Ensure we have a slide
    if (!currentSlide || !currentTable) {
      const result = createNewTableSlide();
      currentSlide = result.slide;
      currentTable = result.table;
    }
    
    // Check if we need a new slide
    const isNewMajorElement = item.type === 'initiative' || item.type === 'valueStream' || item.type === 'epic';
    const isNearEndOfSlide = currentLineCount >= ROLL_THRESHOLD;
    const shouldRollForNewContent = isNearEndOfSlide && isNewMajorElement && 
      (item.program !== currentContext.program || item.type === 'initiative' || item.type === 'valueStream');
    
    if (currentLineCount + item.lines > MAX_LINES_PER_SLIDE || shouldRollForNewContent) {
      if (shouldRollForNewContent) {
        console.log(`  Line ${currentLineCount + 1} is near end and starts new ${item.type}, rolling to next slide`);
      } else {
        console.log(`  Line ${currentLineCount} + ${item.lines} would exceed ${MAX_LINES_PER_SLIDE}, starting new slide`);
      }
      
      const result = createNewTableSlide();
      currentSlide = result.slide;
      currentTable = result.table;
      needsContinuation = true;
    }
    
    // Collect consecutive items from same program that fit on this slide
    const programBatch = [];
    const batchStartProgram = item.program;
    let batchLineCount = 0;
    
    // Calculate extra lines needed for continuation headers
    let continuationLineCount = 0;
    if (needsContinuation && currentContext.program === batchStartProgram) {
      continuationLineCount = 1;  // Program (continued)
      if (currentContext.initiative) {
        continuationLineCount += 1;  // Initiative (continued)
        if (currentContext.valueStream) {
          continuationLineCount += 1;  // Value Stream (continued)
        }
      }
    }
    
    while (i < contentLines.length && 
           contentLines[i].program === batchStartProgram &&
           currentLineCount + continuationLineCount + batchLineCount + contentLines[i].lines <= MAX_LINES_PER_SLIDE) {
      
      const nextItem = contentLines[i];
      const nextIsNewMajor = nextItem.type === 'initiative' || nextItem.type === 'valueStream';
      const wouldBeNearEnd = currentLineCount + continuationLineCount + batchLineCount >= ROLL_THRESHOLD;
      
      if (wouldBeNearEnd && nextIsNewMajor && batchLineCount > 0) {
        break;
      }
      
      programBatch.push(contentLines[i]);
      batchLineCount += contentLines[i].lines;
      i++;
    }
    
    // Build and add this batch to the table
    if (programBatch.length > 0) {
      const isSameProgram = currentContext.program === batchStartProgram;
      const displayName = isSameProgram ? `${batchStartProgram} (continued)` : batchStartProgram;
      
      const continuationContext = (needsContinuation && isSameProgram) ? {
        initiative: currentContext.initiative,
        valueStream: currentContext.valueStream
      } : null;
      
      const batchInitiatives = buildInitiativesFromContentBatch(programBatch, showDependencies, continuationContext);
      
      if (batchInitiatives.length > 0) {
        const totalLines = batchLineCount + (continuationContext ? continuationLineCount : 0);
        console.log(`  Adding "${displayName}" with ${batchInitiatives.length} initiative(s), ${totalLines} lines`);
        addProgramRow(currentTable, displayName, batchInitiatives, bgColor, showDependencies, noInitiativeMode, hideSameTeamDeps, true);
        
        currentLineCount += totalLines;
      }

      // Update context for continuation tracking
      const lastItem = programBatch[programBatch.length - 1];
      currentContext.program = lastItem.program;
      currentContext.programBgColor = bgColor;
      currentContext.initiative = lastItem.initiativeObj || null;
      currentContext.valueStream = lastItem.valueStream || null;
      
      needsContinuation = false;
    }
  }
  
  // Add page numbers to table slides
  const totalTableSlides = tableSlides.length;
  tableSlides.forEach((slide, index) => {
    addPortfolioPageNumber(slide, index + 1, totalTableSlides);
  });
  
  console.log(`Created ${slidesCreated} slides for portfolio "${portfolioName}" (1 divider + ${totalTableSlides} table slides)`);
  return slidesCreated;
}
// =================================================================================
// TABLE ROW CREATION
// =================================================================================

function addProgramRow(table, programName, initiatives, bgColor, showDependencies = true, noInitiativeMode = 'show', hideSameTeamDeps = false, groupByFixVersion = true) {
  let cellContent = buildProgramCellContent(initiatives, showDependencies, noInitiativeMode, hideSameTeamDeps, groupByFixVersion);
  if (!cellContent.text || cellContent.text.trim() === '') {
    console.log(`⏭️ Skipping empty row for "${programName}" - no displayable content`);
    return false;
  }
  
  const row = table.appendRow();
  const numCols = table.getNumColumns();
  
  // Apply background color to all cells
  for (let i = 0; i < numCols; i++) {
    row.getCell(i).getFill().setSolidFill(bgColor);
  }
  
  // Column 0: Program Initiative (single name for all initiatives)
  const programCell = row.getCell(0);
  const cleanProgramName = removeLeadingNumber(programName);
  programCell.getText().setText(cleanProgramName);
  
  // Check if program name is a missing data placeholder
  const isProgramMissing = MISSING_DATA_PLACEHOLDERS.includes(cleanProgramName);
  
  programCell.getText().getTextStyle()
    .setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE)
    .setBold(true)
    .setFontFamily(SLIDE_CONFIG.BODY_FONT)
    .setForegroundColor(isProgramMissing ? SLIDE_CONFIG.MISSING_DATA_TEXT : SLIDE_CONFIG.BLACK);
  
  // Apply orange highlight if missing data
  if (isProgramMissing) {
    programCell.getFill().setSolidFill(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
  }
  
  programCell.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
    .setSpaceAbove(2)
    .setSpaceBelow(2);
  
  // Column 1: PI Objective (all initiatives with Value Stream grouping)
  const objectiveCell = row.getCell(1);
  const textRange = objectiveCell.getText();
  
  // Set content and apply styling
  textRange.setText(cellContent.text);
  applyProgramCellStyling(textRange, cellContent);
  
  textRange.getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.START)
    .setSpaceAbove(2)
    .setSpaceBelow(2);
  
  return true;  // Row was successfully added
}
// =================================================================================
// CELL CONTENT BUILDING - SINGLE CONSOLIDATED FUNCTION
// =================================================================================

/**
 * Groups epics by Fix Version and Iteration for cleaner display
 * Returns categorized epics: changed (at top), grouped, and no-fix-version
 * 
 * @param {Array} epics - Array of epic objects
 * @returns {Object} { changedEpics, groups, noFixVersionEpics }
 */
function groupEpicsForFixVersionDisplay(epics) {
  const changedEpics = [];
  const groupableEpics = [];
  const noFixVersionEpics = [];
  
  // Badge priority for sorting: NEW=1, CHG=2, DONE=3, none=4
  function getBadgePriority(epic) {
    const epicData = processEpicData(epic);
    const changes = epicData.changes || {};
    const isNewlyCommitted = changes.piCommitmentToCommitted === true;
    
    if (epicData.isNew || isNewlyCommitted) return 1;  // NEW
    if (epicData.isModified || epicData.closedThisIteration || epicData.deferredThisIteration) return 2;  // CHG
    if (epicData.isClosed || epicData.isDone) return 3;  // DONE
    return 4;  // No badge
  }
  
  // Step 1: Categorize epics
  epics.forEach(epic => {
    const epicData = processEpicData(epic);
    const changes = epicData.changes || {};
    const fixVersions = (epic['Fix Versions'] || '').toString().trim();
    const endIteration = (epic['End Iteration Name'] || epic['PI Target Iteration'] || '').toString().trim();
    
    // Epics with iteration changes go to changed section (shown first with strikethrough)
    if (changes.iterationChanged) {
      epic._displayMode = 'full';
      epic._badgePriority = getBadgePriority(epic);
      changedEpics.push(epic);
    }
    // Epics with fix version AND iteration can be grouped
    else if (fixVersions && endIteration) {
      epic._groupKey = `${fixVersions}|||${endIteration}`;
      epic._fixVersion = fixVersions;
      epic._iteration = endIteration;
      epic._badgePriority = getBadgePriority(epic);
      groupableEpics.push(epic);
    }
    // Epics without fix version go to separate section
    else if (!fixVersions) {
      epic._displayMode = 'iterationOnly';
      epic._badgePriority = getBadgePriority(epic);
      noFixVersionEpics.push(epic);
    }
    // Epics with fix version but no iteration - show full
    else {
      epic._displayMode = 'full';
      epic._badgePriority = getBadgePriority(epic);
      changedEpics.push(epic);
    }
  });
  
  // Step 2: Group groupable epics by (fixVersion, iteration)
  const groupMap = {};
  groupableEpics.forEach(epic => {
    const key = epic._groupKey;
    if (!groupMap[key]) {
      groupMap[key] = {
        fixVersion: epic._fixVersion,
        iteration: epic._iteration,
        iterationNum: extractIterationNumber(epic._iteration) || 99,
        epics: []
      };
    }
    groupMap[key].epics.push(epic);
  });
  
  // Step 3: Separate singleton groups (1 epic) vs real groups (2+ epics)
  const realGroups = [];
  Object.values(groupMap).forEach(group => {
    if (group.epics.length === 1) {
      // Singleton - treat as changed epic (full display)
      group.epics[0]._displayMode = 'full';
      changedEpics.push(group.epics[0]);
    } else {
      // Real group - mark epics for summary-only display
      group.epics.forEach(epic => {
        epic._displayMode = 'summaryOnly';
      });
      
      // Sort epics within group by badge priority: NEW > CHG > DONE > none
      group.epics.sort((a, b) => {
        const priorityDiff = a._badgePriority - b._badgePriority;
        if (priorityDiff !== 0) return priorityDiff;
        // Secondary sort by summary alphabetically
        return (a['Summary'] || '').localeCompare(b['Summary'] || '');
      });
      
      realGroups.push(group);
    }
  });
  
  // Step 4: Sort groups by fix version (alpha), then iteration (numeric ascending)
  realGroups.sort((a, b) => {
    const fvCompare = a.fixVersion.localeCompare(b.fixVersion);
    if (fvCompare !== 0) return fvCompare;
    return a.iterationNum - b.iterationNum;
  });
  
  // Step 5: Sort no-fix-version epics by badge priority, then iteration
  noFixVersionEpics.sort((a, b) => {
    const priorityDiff = a._badgePriority - b._badgePriority;
    if (priorityDiff !== 0) return priorityDiff;
    const iterA = extractIterationNumber(a['End Iteration Name'] || a['PI Target Iteration']) || 99;
    const iterB = extractIterationNumber(b['End Iteration Name'] || b['PI Target Iteration']) || 99;
    return iterA - iterB;
  });
  
  // Step 6: Sort changed epics by badge priority, then iteration
  changedEpics.sort((a, b) => {
    const priorityDiff = a._badgePriority - b._badgePriority;
    if (priorityDiff !== 0) return priorityDiff;
    const iterA = extractIterationNumber(a['End Iteration Name'] || a['PI Target Iteration']) || 99;
    const iterB = extractIterationNumber(b['End Iteration Name'] || b['PI Target Iteration']) || 99;
    return iterA - iterB;
  });
  
  return {
    changedEpics,      // Show first, full line format with strikethrough
    groups: realGroups, // Show grouped with headers
    noFixVersionEpics  // Show last in "No Fix Version Assigned" section
  };
}

function buildProgramCellContent(initiatives, showDependencies = true, noInitiativeMode = 'show', hideSameTeamDeps = false, groupByFixVersion = true) {
  let text = '';
  const metadata = {
    initiativeSections: [],
    valueStreamSections: [],
    epicSections: [],
    dependencySections: [],
    riskSections: [],
    fixVersionHeaders: [],    // Fix version group headers
    iterationHeaders: [],     // Iteration sub-headers within fix version groups
    noFixVersionHeader: null  // "No Fix Version Assigned" header
  };
  
  let previousInitiativeWasSkipped = false;
  let isFirstContent = true;
  
  initiatives.forEach((initiative, initIndex) => {
    const titleText = initiative.name;
    const isNoInitiative = titleText && titleText.startsWith('No Initiative');
    const skipInitiativeHeader = isNoInitiative && noInitiativeMode === 'skip';
    
    // Add separator between initiatives (but not if previous was skipped or this is skipped)
    if (initIndex > 0 && !previousInitiativeWasSkipped && !skipInitiativeHeader) {
      const separatorStart = text.length;
      text += '\n';  // Single newline instead of double
      const separatorEnd = text.length;
      
      metadata.separatorSections = metadata.separatorSections || [];
      metadata.separatorSections.push({
        start: separatorStart,
        end: separatorEnd
      });
    }
    
    // Count epics for this initiative
    const epicCount = (initiative.epics || []).filter(e => {
      const data = processEpicData(e);
      return !data.summaryIsEmpty;
    }).length;
    
    if (!skipInitiativeHeader) {
      const titleStart = text.length;
      const fpText = ` | Feature Points: ${initiative.totalFeaturePoints}`;
      
      text += titleText + fpText;
      
      const isInitiativeMissing = isMissingDataPlaceholder(titleText);
      
      let effectiveLinkKey = initiative.linkKey;
      if (!effectiveLinkKey || effectiveLinkKey.trim() === '') {
        if (initiative.epics && initiative.epics.length > 0) {
          const firstEpic = initiative.epics[0];
          effectiveLinkKey = firstEpic['Parent Key'] || firstEpic['Key'] || null;
          if (effectiveLinkKey) {
            console.log(`🔗 Initiative "${titleText.substring(0, 40)}..." → ${effectiveLinkKey}`);
          }
        }
      }
      
      metadata.initiativeSections.push({
        titleStart: titleStart,
        titleEnd: titleStart + titleText.length,
        fpStart: titleStart + titleText.length,
        fpEnd: text.length,
        linkKey: effectiveLinkKey,
        isMissing: isInitiativeMissing
      });
      
      // Technical Perspective only if 2+ epics
      const perspective = initiative.technicalPerspective || '';
      if (perspective && epicCount >= 2) {
        const perspectiveStart = text.length + 1;
        text += '\n' + perspective;
        metadata.initiativeSections[metadata.initiativeSections.length - 1].perspectiveStart = perspectiveStart;
        metadata.initiativeSections[metadata.initiativeSections.length - 1].perspectiveEnd = text.length;
      }
    }
    
    // Group epics by Value Stream
    const valueStreamGroups = {};
    initiative.epics.forEach(epic => {
      const vs = epic['Value Stream/Org'] || 'Other';
      if (!valueStreamGroups[vs]) {
        valueStreamGroups[vs] = [];
      }
      valueStreamGroups[vs].push(epic);
    });
    
    const sortedValueStreams = Object.keys(valueStreamGroups).sort();
    let isFirstValueStream = true;
    
    sortedValueStreams.forEach(vsName => {
      let vsEpics = valueStreamGroups[vsName];
      
      const hasVsContinuation = isFirstValueStream && 
                                initiative.isContinuation && 
                                vsEpics.length > 0 && 
                                vsEpics[0]._valueStreamContinued;
      
      const vsDisplayName = hasVsContinuation ? `${vsName} (continued)` : vsName;
      
      // Only add newline before Value Stream if there's content before it
      // (avoids blank line when initiative header was skipped)
      const needsNewline = text.length > 0;
      const vsStart = needsNewline ? text.length + 1 : text.length;
      
      if (needsNewline) {
        text += '\n' + vsDisplayName;
      } else {
        text += vsDisplayName;
      }
      isFirstContent = false;  
      
      metadata.valueStreamSections.push({
        start: vsStart,
        end: text.length,
        name: vsName
      });
      
      // ═══════════════════════════════════════════════════════════════════
      // PROCESS EPICS - Either grouped or line-by-line
      // ═══════════════════════════════════════════════════════════════════
      let orderedEpics = [];
      
      if (groupByFixVersion) {
        // GROUP EPICS BY FIX VERSION AND ITERATION
        const grouped = groupEpicsForFixVersionDisplay(vsEpics);
        
        // 1. Changed epics first (full format with strikethrough)
        grouped.changedEpics.forEach(epic => {
          orderedEpics.push({ epic, displayMode: 'full', groupHeader: null });
        });
        
        // 2. Grouped epics with headers
        grouped.groups.forEach(group => {
          // Add fix version header before first epic of group
          orderedEpics.push({ 
            epic: null, 
            displayMode: null, 
            groupHeader: { type: 'fixVersion', text: `Fix: ${group.fixVersion}` }
          });
          // Add iteration sub-header
          orderedEpics.push({ 
            epic: null, 
            displayMode: null, 
            groupHeader: { type: 'iteration', text: getIterationDisplayText(group.iteration) || group.iteration }
          });
          // Add epics in the group (summary only)
          group.epics.forEach(epic => {
            orderedEpics.push({ epic, displayMode: 'summaryOnly', groupHeader: null });
          });
        });
        
        // 3. No Fix Version section
        if (grouped.noFixVersionEpics.length > 0) {
          // Add "No Fix Version Assigned" header
          orderedEpics.push({ 
            epic: null, 
            displayMode: null, 
            groupHeader: { type: 'noFixVersion', text: 'No Fix Version Assigned' }
          });
          // Add epics (iteration only)
          grouped.noFixVersionEpics.forEach(epic => {
            orderedEpics.push({ epic, displayMode: 'iterationOnly', groupHeader: null });
          });
        }
      } else {
        // ORIGINAL LINE-BY-LINE FORMAT - Sort by iteration and show full info on each line
        vsEpics.sort((a, b) => {
          const iterA = (a['End Iteration Name'] || a['PI Target Iteration'] || '').toString();
          const iterB = (b['End Iteration Name'] || b['PI Target Iteration'] || '').toString();
          return iterA.localeCompare(iterB);
        });
        
        vsEpics.forEach(epic => {
          orderedEpics.push({ epic, displayMode: 'full', groupHeader: null });
        });
      }
      
      // Process ordered epics
      orderedEpics.forEach(item => {
        // Handle group headers
        if (item.groupHeader) {
          const headerIndent = SLIDE_CONFIG.EPIC_INDENT || '  ';
          const headerLine = `\n${item.groupHeader.text}`;
          const headerStart = text.length + 1;
          text += headerLine;
          
          if (item.groupHeader.type === 'fixVersion') {
            metadata.fixVersionHeaders.push({
              start: headerStart,
              end: text.length,
              text: item.groupHeader.text
            });
          } else if (item.groupHeader.type === 'iteration') {
            metadata.iterationHeaders.push({
              start: headerStart,
              end: text.length,
              text: item.groupHeader.text
            });
          } else if (item.groupHeader.type === 'noFixVersion') {
            metadata.noFixVersionHeader = {
              start: headerStart,
              end: text.length,
              text: item.groupHeader.text
            };
          }
          return;
        }
        
        // Handle epics
        const epic = item.epic;
        const displayMode = item.displayMode;
        const epicData = processEpicData(epic);
        
        if (epicData.summaryIsEmpty) return;
        
        // ═══════════════════════════════════════════════════════════════════
        // BADGE LOGIC - Determines which badge(s) to show
        // Priority: CANCELED > DEF > DONE > NEW > OVERDUE > CHG > ATRISK
        // ═══════════════════════════════════════════════════════════════════
        const changes = epicData.changes || {};
        const badges = [];
        
        const isNewlyCommitted = changes.piCommitmentToCommitted === true;
        
        // CHG badge: Has changes but not transitioning to NEW or other status badges
        const showChgBadge = (epicData.closedThisIteration || 
                            epicData.deferredThisIteration || 
                            (epicData.isModified && !epicData.isNew)) &&
                            !isNewlyCommitted &&
                            !epicData.isCanceled &&
                            !epicData.isDeferred &&
                            !epicData.isOverdue;
        
        if (showChgBadge) {
          badges.push('CHG');
        }
        
        // Status badges - CANCELED takes highest priority
        if (epicData.isCanceled) {
          badges.push('CANCELED');
        } else if (epicData.isClosed) {
          badges.push('DONE');
        } else if (epicData.isDeferred) {
          badges.push('DEF');
        } else if (epicData.isNew || isNewlyCommitted) {
          if (badges.length > 0 && badges[0] === 'CHG') {
            badges[0] = 'NEW';  // Replace CHG with NEW
          } else if (badges.length === 0) {
            badges.push('NEW');
          }
        } else if (epicData.isOverdue) {
          // OVERDUE: Due this iteration but not complete
          if (badges.length > 0 && badges[0] === 'CHG') {
            badges[0] = 'OVERDUE';  // Replace CHG with OVERDUE
          } else if (badges.length === 0) {
            badges.push('OVERDUE');
          }
        }
        
        let bulletChar = '•';
        if (badges.length > 0) {
          bulletChar = badges.join('');
        } else if (epicData.showRag) {
          bulletChar = '• ' + epicData.ragIndicator;
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // ITERATION TEXT - With strikethrough for changes or *new* indicator
        // Format: "Iter 3 Iter 5" where "Iter 3" gets strikethrough
        // Format: "Iter 5 *new*" when iteration was blank before
        // SKIP if displayMode is 'summaryOnly' (grouped epics)
        // ═══════════════════════════════════════════════════════════════════
        let iterationText = '';
        let isIterationMissing = false;
        let hasStrikethroughIter = false;
        let oldIterText = '';
        let hasNewIterIndicator = false;
        let shouldHighlightMissingIter = false;
        
        // Only include iteration text for 'full' and 'iterationOnly' display modes
        const showIterationInLine = (displayMode !== 'summaryOnly');
        
        if (showIterationInLine) {
          iterationText = getIterationDisplayText(epicData.endIteration) || 'Iteration unknown';
          isIterationMissing = (iterationText === 'Iteration unknown');
          
          if (changes.iterationChanged && changes.previousIteration) {
            // Iteration CHANGED from one value to another - show strikethrough
            oldIterText = getIterationDisplayText(changes.previousIteration) || changes.previousIteration;
            iterationText = oldIterText + ' ' + iterationText;  // "Iter 3 Iter 5"
            hasStrikethroughIter = true;
          } else if (changes.iterationChanged && !changes.previousIteration && !isIterationMissing) {
            // Iteration was BLANK before, now has value - show *new*
            iterationText = iterationText + ' *new*';
            hasNewIterIndicator = true;
          }
          
          // Determine if iteration should be highlighted as missing
          // DON'T highlight if item is CLOSED (they don't need an iteration anymore)
          const isClosed = epicData.isClosed || epicData.isDone || 
            (epicData.status && epicData.status.toLowerCase() === 'closed');
          shouldHighlightMissingIter = isIterationMissing && !isClosed;
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // FIX VERSIONS - With strikethrough for changes or *new* indicator
        // Format: "Fix: Release 7.13 Release 7.14" where old gets strikethrough
        // Format: "Fix: Release 7.14 *new*" when fix version was blank before
        // SKIP if displayMode is 'summaryOnly' or 'iterationOnly'
        // ═══════════════════════════════════════════════════════════════════
        let fixVersionsText = '';
        let hasStrikethroughFV = false;
        let oldFVText = '';
        let newFVText = '';
        let hasNewFVIndicator = false;
        
        // Only include fix version text for 'full' display mode
        const showFixVersionInLine = (displayMode === 'full');
        
        if (showFixVersionInLine) {
          if (changes.fixVersionsChanged && changes.previousFixVersions) {
            // Fix version CHANGED from one value to another - show strikethrough
            oldFVText = changes.previousFixVersions;
            newFVText = changes.currentFixVersions || '(removed)';
            fixVersionsText = ` | Fix: ${oldFVText} ${newFVText}`;
            hasStrikethroughFV = true;
          } else if (changes.fixVersionsChanged && !changes.previousFixVersions && changes.currentFixVersions) {
            // Fix version was BLANK before, now has value - show *new*
            fixVersionsText = ' | Fix: ' + changes.currentFixVersions + ' *new*';
            hasNewFVIndicator = true;
          } else if (epicData.fixVersions) {
            // No change, just show current
            fixVersionsText = ' | Fix: ' + epicData.fixVersions;
          }
        }
        
        const epicIterNum = extractIterationNumber(epicData.endIteration);
        
        // ═══════════════════════════════════════════════════════════════════
        // BUILD EPIC LINE - Format varies by displayMode
        // 'full': summary | iteration | fix version
        // 'iterationOnly': summary | iteration  
        // 'summaryOnly': summary only
        // ═══════════════════════════════════════════════════════════════════
        const epicIndent = SLIDE_CONFIG.EPIC_INDENT || '  ';
        let epicLine;
        
        if (displayMode === 'summaryOnly') {
          // Grouped epics - just show summary
          epicLine = `\n${epicIndent}${bulletChar} ${epicData.summary}`;
        } else if (displayMode === 'iterationOnly') {
          // No fix version section - show summary + iteration
          epicLine = `\n${epicIndent}${bulletChar} ${epicData.summary} | ${iterationText}`;
        } else {
          // Full format - summary + iteration + fix version
          epicLine = `\n${epicIndent}${bulletChar} ${epicData.summary} | ${iterationText}${fixVersionsText}`;
        }
        
        const epicStart = text.length + 1;
        text += epicLine;
        
        // ═══════════════════════════════════════════════════════════════════
        // CALCULATE POSITIONS - Vary by displayMode
        // ═══════════════════════════════════════════════════════════════════
        const bulletStart = epicStart + epicIndent.length;
        const bulletEnd = bulletStart + bulletChar.length;
        const summaryStart = bulletEnd + 1;
        const summaryEnd = summaryStart + epicData.summary.length;
        
        // Pipe and iteration positions - only for full and iterationOnly modes
        let pipeStart = null;
        let pipeEnd = null;
        let iterationStart = null;
        let strikethroughStart = null;
        let strikethroughEnd = null;
        let newIterStart = null;
        let baseIterEnd = summaryEnd;
        
        if (displayMode !== 'summaryOnly' && iterationText) {
          pipeStart = summaryEnd + 1;
          pipeEnd = summaryEnd + 2;
          iterationStart = summaryEnd + 3;
          
          // Iteration strikethrough positions
          if (hasStrikethroughIter && oldIterText) {
            strikethroughStart = iterationStart;
            strikethroughEnd = iterationStart + oldIterText.length;
            newIterStart = strikethroughEnd + 1;
          }
          
          baseIterEnd = iterationStart + iterationText.length;
        }
        
        // Fix versions positions - only for full mode
        let fixVersionsStart = null;
        let fixVersionsEnd = null;
        let fvStrikethroughStart = null;
        let fvStrikethroughEnd = null;
        let fvNewStart = null;
        let fvNewEnd = null;
        let fvLabelEnd = null;
        
        if (displayMode === 'full' && fixVersionsText) {
          fixVersionsStart = baseIterEnd;  // Start of " | Fix: ..."
          fixVersionsEnd = baseIterEnd + fixVersionsText.length;
          fvLabelEnd = baseIterEnd + 8;  // " | Fix: " is 8 chars
          
          if (hasStrikethroughFV && oldFVText) {
            // Calculate positions within fix versions text
            // Format: " | Fix: oldValue newValue"
            fvStrikethroughStart = fvLabelEnd;
            fvStrikethroughEnd = fvLabelEnd + oldFVText.length;
            fvNewStart = fvStrikethroughEnd + 1;  // After space
            fvNewEnd = fixVersionsEnd;
          }
        }
        
        const iterationEnd = (displayMode === 'full' && fixVersionsText) ? (baseIterEnd + fixVersionsText.length) : baseIterEnd;
        
        console.log(`Epic: "${epicData.summary.substring(0, 30)}..." badges=[${badges.join(',')}], iter=${iterationText}${hasStrikethroughIter ? ' (strikethrough)' : ''}${hasNewIterIndicator ? ' (*new*)' : ''}${hasStrikethroughFV ? ', FV (strikethrough)' : ''}${hasNewFVIndicator ? ', FV (*new*)' : ''}`);
        
        const isSummaryMissing = isMissingDataPlaceholder(epicData.summary);
        const piCommitment = (epic['PI Commitment'] || '').toString().trim();
        const isNotCommitted = piCommitment.toUpperCase() === 'NOT COMMITTED';
        
        const epicSectionIndex = metadata.epicSections.length;
        
        metadata.epicSections.push({
          type: 'epic',
          start: epicStart,
          end: text.length,
          bulletStart: bulletStart,
          bulletEnd: bulletEnd,
          summaryStart: summaryStart,
          summaryEnd: summaryEnd,
          pipeStart: pipeStart,
          pipeEnd: pipeEnd,
          url: epicData.url,
          hasIteration: epicData.endIteration && epicData.endIteration.trim() !== '',
          hasRagBullet: epicData.showRag || badges.length > 0,
          ragIndicator: epicData.ragIndicator,
          isClosed: epicData.isClosed,
          badges: badges,
          badgeType: badges.length > 0 ? badges[badges.length - 1] : null,
          isSummaryMissing: isSummaryMissing,
          iterationMissing: isIterationMissing,
          shouldHighlightMissingIter: shouldHighlightMissingIter,  // Don't highlight if closed
          iterationStart: iterationStart,
          iterationEnd: iterationEnd,
          isNotCommitted: isNotCommitted,
          hasLaterDependency: false,
          epicIterNum: epicIterNum,
          // Iteration strikethrough
          hasStrikethroughIter: hasStrikethroughIter,
          strikethroughStart: strikethroughStart,
          strikethroughEnd: strikethroughEnd,
          newIterStart: newIterStart,
          // *new* indicators
          hasNewIterIndicator: hasNewIterIndicator,
          hasNewFVIndicator: hasNewFVIndicator,
          // Fix Versions with strikethrough
          fixVersionsStart: fixVersionsStart,
          fixVersionsEnd: fixVersionsEnd,
          hasStrikethroughFV: hasStrikethroughFV,
          fvStrikethroughStart: fvStrikethroughStart,
          fvStrikethroughEnd: fvStrikethroughEnd,
          fvNewStart: fvNewStart,
          fvNewEnd: fvNewEnd,
          fvLabelEnd: fvLabelEnd,
          changes: changes
        });
        
        // ═══════════════════════════════════════════════════════════════════
        // RAG NOTES - ENHANCED: Show BOTH old and new when ragNoteChanged
        // Latest iteration on top
        // ═══════════════════════════════════════════════════════════════════
        if (epicData.showRag) {
          const ragNotes = changes.ragNotes || [];
          
          if (ragNotes.length > 1) {
            // Multiple RAG notes - show each with iteration label
            const ragColor = epicData.rag || 'Unknown';
            const riskLabelLine = `\n${SLIDE_CONFIG.EPIC_INDENT}Risk - ${ragColor}`;
            const riskLabelStart = text.length + 1;
            text += riskLabelLine;
            
            const labelContentStart = riskLabelStart + (SLIDE_CONFIG.EPIC_INDENT || '  ').length;
            
            metadata.riskSections.push({
              type: 'ragLabel',
              start: riskLabelStart,
              end: text.length,
              labelStart: labelContentStart,
              labelEnd: text.length,
              noteStart: null,
              noteEnd: null,
              hasNote: true,
              isMissing: false,
              isDependency: false,
              isLabelOnly: true
            });
            
            // Note lines with iteration labels (newest first - already sorted)
            ragNotes.forEach((noteInfo, idx) => {
              const iterLabel = `Iter ${noteInfo.iteration}`;
              const noteLine = `\n${SLIDE_CONFIG.EPIC_INDENT}  - ${iterLabel}: ${noteInfo.note}`;
              const noteStart = text.length + 1;
              text += noteLine;
              
              const dashStart = noteStart + (SLIDE_CONFIG.EPIC_INDENT || '  ').length + 2;
              const iterLabelEnd = dashStart + 2 + iterLabel.length + 1;
              const noteTextStart = iterLabelEnd + 1;
              
              metadata.riskSections.push({
                type: 'ragNoteWithIter',
                start: noteStart,
                end: text.length,
                labelStart: dashStart,
                labelEnd: iterLabelEnd,
                noteStart: noteTextStart,
                noteEnd: text.length,
                hasNote: true,
                isMissing: false,
                isDependency: false,
                isCurrent: noteInfo.isCurrent,
                iteration: noteInfo.iteration,
                isPreviousNote: !noteInfo.isCurrent
              });
            });
            
          } else {
            // Single RAG note
            const ragNoteText = epicData.ragNote || 'NO RAG NOTE';
            const isRagNoteMissing = (ragNoteText === 'NO RAG NOTE');
            
            let riskLabel = 'Risk - ';
            if (ragNotes.length === 1) {
              riskLabel = `Risk - Iter ${ragNotes[0].iteration}: `;
            }
            
            const riskLine = `\n${SLIDE_CONFIG.EPIC_INDENT}${riskLabel}${ragNoteText}`;
            const riskStart = text.length + 1;
            text += riskLine;
            
            const labelStart = riskStart + (SLIDE_CONFIG.EPIC_INDENT || '  ').length;
            const labelEnd = labelStart + riskLabel.length - 1;
            const noteStart = labelEnd + 1;
            const noteEnd = text.length;
            
            metadata.riskSections.push({
              start: riskStart,
              end: text.length,
              labelStart: labelStart,
              labelEnd: labelEnd,
              noteStart: noteStart,
              noteEnd: noteEnd,
              hasNote: !!epicData.ragNote,
              isMissing: isRagNoteMissing,
              isDependency: false
            });
          }
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // RAG MITIGATION NOTE - When risk status moves to Green
        // ═══════════════════════════════════════════════════════════════════
        if (changes.ragMitigated) {
          const mitigationLabel = 'Note - ';
          const mitigationText = 'Risk has been mitigated';
          const mitigationIndent = SLIDE_CONFIG.EPIC_INDENT || '  ';
          const mitigationLine = `\n${mitigationIndent}${mitigationLabel}${mitigationText}`;
          const mitigationStart = text.length + 1;
          text += mitigationLine;
          
          const labelStart = mitigationStart + mitigationIndent.length;
          const labelEnd = labelStart + mitigationLabel.length - 1;
          const noteTextStart = labelEnd + 1;
          const noteTextEnd = text.length;
          
          console.log(`  ✅ Adding RAG mitigation note for ${epicData.key}`);
          
          metadata.noteSections = metadata.noteSections || [];
          metadata.noteSections.push({
            start: mitigationStart,
            end: text.length,
            labelStart: labelStart,
            labelEnd: labelEnd,
            noteStart: noteTextStart,
            noteEnd: noteTextEnd,
            isDependency: false,
            isPositiveNote: true
          });
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // NOTE SECTION - RAG Note exists but no RAG status
        // ═══════════════════════════════════════════════════════════════════
        else if (epicData.showNote) {
          const noteLabel = 'Note - ';
          const noteText = epicData.noteText || '';
          
          const noteLine = `\n${SLIDE_CONFIG.EPIC_INDENT}${noteLabel}${noteText}`;
          const noteStart = text.length + 1;
          text += noteLine;
          
          const labelStart = noteStart + (SLIDE_CONFIG.EPIC_INDENT || '  ').length;
          const labelEnd = labelStart + noteLabel.length - 1;
          const noteTextStart = labelEnd + 1;
          const noteTextEnd = text.length;
          
          metadata.noteSections = metadata.noteSections || [];
          metadata.noteSections.push({
            start: noteStart,
            end: text.length,
            labelStart: labelStart,
            labelEnd: labelEnd,
            noteStart: noteTextStart,
            noteEnd: noteTextEnd,
            isDependency: false
          });
          
          console.log(`  -> Adding epic note line: "${noteLabel}${noteText.substring(0, 50)}..."`);
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // DEPENDENCIES
        // ═══════════════════════════════════════════════════════════════════
        if (showDependencies && epicData.dependencies && epicData.dependencies.length > 0) {
          let sortedDeps = [...epicData.dependencies].sort((a, b) => {
            const iterA = (a['End Iteration Name'] || '').toString();
            const iterB = (b['End Iteration Name'] || '').toString();
            return iterA.localeCompare(iterB);
          });
          
          if (hideSameTeamDeps) {
            const epicValueStream = (epic['Value Stream/Org'] || '').toLowerCase().trim();
            sortedDeps = sortedDeps.filter(dep => {
              const depTeam = (dep['Depends on Valuestream'] || '').toLowerCase().trim();
              const isSameTeam = epicValueStream && depTeam && epicValueStream === depTeam;
              if (isSameTeam) {
                console.log(`🔇 Hiding same-team dep: "${(dep['Summary'] || '').substring(0, 40)}..." (${depTeam} = ${epicValueStream})`);
              }
              return !isSameTeam;
            });
          }
          
          const beforeStatusFilter = sortedDeps.length;
          sortedDeps = sortedDeps.filter(dep => {
            if (dep.depShouldShow === false) {
              console.log(`🚫 Filtering out dep ${dep['Key']}: ${dep.depHideReason || 'Hidden by business rule'}`);
              return false;
            }
            return true;
          });
          
          if (beforeStatusFilter !== sortedDeps.length) {
            console.log(`  Status filter: ${beforeStatusFilter} -> ${sortedDeps.length} dependencies`);
          }
          
          sortedDeps.forEach(dep => {
            const depRagValue = dep['RAG'] || '';
            const depRagNote = dep['RAG Note'] || '';
            const showRag = !!getRagIndicator(depRagValue);
            
            console.log(`Dependency ${dep['Key']}: RAG="${depRagValue}", RAG Note="${depRagNote}", showRag=${showRag}`);
            
            const rawDepIter = dep['End Iteration Name'];
            const depIterText = getIterationDisplayText(rawDepIter) || 'Iteration unknown';
            const isDepIterMissing = !rawDepIter || 
                                     rawDepIter.toString().trim() === '' ||
                                     depIterText === 'Iteration unknown' ||
                                     depIterText.toLowerCase().includes('unknown');
            
            const depIterNum = extractIterationNumber(rawDepIter);
            const iterationAfterEpic = (epicIterNum !== null && depIterNum !== null && depIterNum > epicIterNum);
            
            if (iterationAfterEpic) {
              console.log(`⚠️ SCHEDULING RISK: Dependency "${dep['Key']}" iteration ${depIterNum} > Epic iteration ${epicIterNum}`);
              metadata.epicSections[epicSectionIndex].hasLaterDependency = true;
            }
            
            console.log(`Dependency iteration: raw="${rawDepIter}", display="${depIterText}", depIterNum=${depIterNum}, missing=${isDepIterMissing}, afterEpic=${iterationAfterEpic}`);
            
            let depBullet = '• 🔗';
            if (showRag) {
              depBullet = '• ' + getRagIndicator(depRagValue) + ' 🔗';
            }
            
            let depBadgeText = '';
            if (dep.depBadges && dep.depBadges.length > 0) {
              depBadgeText = ' ' + dep.depBadges.map(b => `[${b}]`).join(' ');
              console.log(`  Adding badges to ${dep['Key']}: ${depBadgeText}`);
            }
            
            const depVS = dep['Depends on Valuestream'] || 'VS Unknown';
            const isDepVSMissing = (depVS === 'VS Unknown' || !dep['Depends on Valuestream'] || dep['Depends on Valuestream'].trim() === '');
            
            const depIndent = SLIDE_CONFIG.DEPENDENCY_INDENT || '      ';
            const depSummary = dep['Summary'] || '';
            
            const depLine = `\n${depIndent}${depBullet} ${depSummary}${depBadgeText} | ${depVS} | ${depIterText}`;
            
            const depStart = text.length + 1;
            text += depLine;
            
            const bulletStart = depStart + depIndent.length;
            const bulletEnd = bulletStart + depBullet.length;
            const depSummaryStart = bulletEnd + 1;
            const depSummaryEnd = depSummaryStart + depSummary.length;
            
            const badgeStart = depSummaryEnd;
            const badgeEnd = depSummaryEnd + depBadgeText.length;
            
            const pipe1Start = badgeEnd + 2;
            const pipe1End = badgeEnd + 3;
            
            const vsStart = pipe1End + 1;
            const vsEnd = vsStart + depVS.length;
            
            const pipe2Start = vsEnd + 2;
            const pipe2End = vsEnd + 3;
            
            const depIterStart = text.length - depIterText.length;
            const depIterEnd = text.length;
            
            metadata.dependencySections.push({
              start: depStart,
              end: text.length,
              bulletStart: bulletStart,
              bulletEnd: bulletEnd,
              summaryStart: depSummaryStart,
              summaryEnd: depSummaryEnd,
              hasBadges: dep.depBadges && dep.depBadges.length > 0,
              badges: dep.depBadges || [],
              badgeStart: badgeStart,
              badgeEnd: badgeEnd,
              pipe1Start: pipe1Start,
              pipe1End: pipe1End,
              pipe2Start: pipe2Start,
              pipe2End: pipe2End,
              key: dep['Key'],
              hasIteration: dep['End Iteration Name'] && dep['End Iteration Name'].trim() !== '',
              hasRagBullet: showRag,
              vsMissing: isDepVSMissing,
              vsStart: vsStart,
              vsEnd: vsEnd,
              iterationMissing: isDepIterMissing,
              iterationStart: depIterStart,
              iterationEnd: depIterEnd,
              iterationAfterEpic: iterationAfterEpic
            });
            
            // Dependency RAG note
            if (showRag) {
              const depRagNoteText = dep['RAG Note'] || 'NO RAG NOTE';
              const isDepRagNoteMissing = (depRagNoteText === 'NO RAG NOTE' || depRagNoteText.trim() === '');
              const riskIndent = SLIDE_CONFIG.DEPENDENCY_INDENT || '      ';
              const riskLabel = 'Risk - ';
              const depRiskLine = `\n${riskIndent}  ${riskLabel}${depRagNoteText}`;
              const depRiskStart = text.length + 1;
              text += depRiskLine;
              
              console.log(`  -> Adding dependency risk line: "${riskLabel}${depRagNoteText}" (missing=${isDepRagNoteMissing})`);
              
              const labelStart = depRiskStart + riskIndent.length + 2;
              const labelEnd = labelStart + riskLabel.length - 1;
              const depNoteStart = labelEnd + 1;
              const depNoteEnd = text.length;
              
              metadata.riskSections.push({
                start: depRiskStart,
                end: text.length,
                labelStart: labelStart,
                labelEnd: labelEnd,
                noteStart: depNoteStart,
                noteEnd: depNoteEnd,
                hasNote: !!dep['RAG Note'] && dep['RAG Note'].trim() !== '',
                isDependency: true,
                isMissing: isDepRagNoteMissing
              });
            }
            
            // ═══════════════════════════════════════════════════════════════════
            // DEPENDENCY COMPLETION NOTE - When dependency is completed this iteration
            // ═══════════════════════════════════════════════════════════════════
            if (dep.depDoneThisIteration) {
              const completionLabel = 'Note - ';
              const completionText = 'Dependency has been completed';
              const completionIndent = SLIDE_CONFIG.DEPENDENCY_INDENT || '      ';
              const completionLine = `\n${completionIndent}  ${completionLabel}${completionText}`;
              const completionStart = text.length + 1;
              text += completionLine;
              
              const labelStart = completionStart + completionIndent.length + 2;
              const labelEnd = labelStart + completionLabel.length - 1;
              const noteTextStart = labelEnd + 1;
              const noteTextEnd = text.length;
              
              console.log(`  ✅ Adding dependency completion note for ${dep['Key']}`);
              
              metadata.noteSections = metadata.noteSections || [];
              metadata.noteSections.push({
                start: completionStart,
                end: text.length,
                labelStart: labelStart,
                labelEnd: labelEnd,
                noteStart: noteTextStart,
                noteEnd: noteTextEnd,
                isDependency: true,
                isPositiveNote: true
              });
            }
            
            // Dependency NOTE (not RAG risk, just informational)
            else if (depRagNote && depRagNote.trim() !== '') {
              const depNoteText = depRagNote.trim();
              const isJustStatusIndicator = depNoteText.toLowerCase() === 'green' || 
                                            depNoteText.toLowerCase().startsWith('green -') ||
                                            depNoteText.toLowerCase().startsWith('green:') ||
                                            depNoteText.toLowerCase() === 'on track' ||
                                            depNoteText.toLowerCase() === 'no issues';
              
              if (!isJustStatusIndicator) {
                const noteLabel = 'Note - ';
                const noteIndent = SLIDE_CONFIG.DEPENDENCY_INDENT || '      ';
                const depNoteLine = `\n${noteIndent}  ${noteLabel}${depRagNote}`;
                const depNoteStart = text.length + 1;
                text += depNoteLine;
                
                console.log(`  -> Adding dependency note line: "${noteLabel}${depRagNote.substring(0, 50)}..."`);
                
                const labelStart = depNoteStart + noteIndent.length + 2;
                const labelEnd = labelStart + noteLabel.length - 1;
                const noteTextStart = labelEnd + 1;
                const noteTextEnd = text.length;
                
                metadata.noteSections = metadata.noteSections || [];
                metadata.noteSections.push({
                  start: depNoteStart,
                  end: text.length,
                  labelStart: labelStart,
                  labelEnd: labelEnd,
                  noteStart: noteTextStart,
                  noteEnd: noteTextEnd,
                  isDependency: true
                });
              }
            }
          });
        }
      });
      
      isFirstValueStream = false;
    });
    
    // Track whether this initiative was skipped for separator logic
    previousInitiativeWasSkipped = skipInitiativeHeader;
  });
  
  return { text, metadata };
}
// =================================================================================
// CELL STYLING - SINGLE CONSOLIDATED FUNCTION
// =================================================================================

// ═══════════════════════════════════════════════════════════════════════════════
// FILE: slideformatter.gs
// FUNCTION: applyProgramCellStyling - COMPLETE FIXED VERSION
// REPLACE THE ENTIRE FUNCTION
// ═══════════════════════════════════════════════════════════════════════════════

function applyProgramCellStyling(textRange, content) {
  const meta = content.metadata;
  const fullText = content.text;
  
  console.log(`\n=== STYLING CELL CONTENT (${fullText.length} chars) ===`);
  
  // Default styling for entire cell - explicitly set bold to false
  textRange.getTextStyle()
    .setFontFamily(SLIDE_CONFIG.BODY_FONT)
    .setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE)
    .setForegroundColor(SLIDE_CONFIG.BLACK)
    .setBold(false)
    .setItalic(false);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // INITIATIVE SECTIONS
  // ═══════════════════════════════════════════════════════════════════════════
  meta.initiativeSections.forEach((section, idx) => {
    if (section.titleEnd > section.titleStart) {
      try {
        const titleRange = textRange.getRange(section.titleStart, section.titleEnd);
        const titleText = fullText.substring(section.titleStart, section.titleEnd);
        const isNoInitiative = titleText === 'No Initiative' || titleText === 'No Initiative Associated';
        
        if (isNoInitiative) {
          console.log(`🟠 MISSING: No Initiative "${titleText}"`);
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // INITIATIVE HYPERLINK - Apply FIRST (before styling)
        // ═══════════════════════════════════════════════════════════════════
        if (!isNoInitiative && section.linkKey && section.linkKey.trim() !== '') {
          try {
            const linkUrl = getJiraUrl(section.linkKey);
            if (linkUrl) {
              console.log(`🔗 INITIATIVE: "${titleText.substring(0, 40)}..." → ${section.linkKey}`);
              titleRange.getTextStyle().setLinkUrl(linkUrl);
            } else {
              console.warn(`⚠️ INITIATIVE: "${titleText.substring(0, 40)}..." - getJiraUrl returned null for "${section.linkKey}"`);
            }
          } catch (linkError) {
            console.error(`❌ INITIATIVE link error for "${titleText.substring(0, 40)}...": ${linkError.message}`);
          }
        } else if (!isNoInitiative) {
          console.warn(`⚠️ INITIATIVE: "${titleText.substring(0, 40)}..." - No linkKey available`);
        }
        
        // Apply styling AFTER hyperlink (to override blue link color)
        if (isNoInitiative) {
          titleRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.INITIATIVE_SIZE)
            .setBold(true)
            .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
            .setFontFamily(SLIDE_CONFIG.BODY_FONT)
            .setUnderline(false)
            .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
        } else {
          titleRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.INITIATIVE_SIZE)
            .setBold(true)
            .setForegroundColor(SLIDE_CONFIG.INITIATIVE_PURPLE)
            .setFontFamily(SLIDE_CONFIG.BODY_FONT)
            .setUnderline(false);
        }
        
      } catch (e) {
        console.warn(`Could not style initiative title: ${e.message}`);
      }
    }
    
    // Feature points - highlight if 0
    if (section.fpEnd > section.fpStart) {
      try {
        const fpRange = textRange.getRange(section.fpStart, section.fpEnd);
        const fpText = fullText.substring(section.fpStart, section.fpEnd);
        // Check if Feature Points: 0 (exactly zero)
        const isMissingFP = fpText.includes(': 0') && !fpText.match(/: 0\d/);  // 0 but not 01, 02, etc.
        
        if (isMissingFP) {
          console.log(`🟠 MISSING: Feature Points = 0`);
          fpRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.FEATURE_POINTS_SIZE)
            .setItalic(true)
            .setBold(true)
            .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
            .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
        } else {
          fpRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.FEATURE_POINTS_SIZE)
            .setItalic(true)
            .setBold(false)
            .setForegroundColor(SLIDE_CONFIG.INITIATIVE_PURPLE);
        }
      } catch (e) {
        console.warn(`Could not style feature points: ${e.message}`);
      }
    }
    
    // Perspective (italics, bold, size 8)
    if (section.perspectiveEnd && section.perspectiveEnd > section.perspectiveStart) {
      try {
        const perspectiveRange = textRange.getRange(section.perspectiveStart, section.perspectiveEnd);
        perspectiveRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.PERSPECTIVE_SIZE)
          .setBold(true)
          .setItalic(true)
          .setForegroundColor(SLIDE_CONFIG.BLACK);
      } catch (e) {
        console.warn(`Could not style perspective: ${e.message}`);
      }
    }
  });
  
  // ═══════════════════════════════════════════════════════════════════════════
  // VALUE STREAM SECTIONS
  // ═══════════════════════════════════════════════════════════════════════════
  meta.valueStreamSections.forEach(section => {
    try {
      const vsRange = textRange.getRange(section.start, section.end);
      const vsText = fullText.substring(section.start, section.end);
      const isMissingVS = vsText === 'Other' || vsText === 'VS Unknown' || !vsText.trim();
      
      if (isMissingVS) {
        console.log(`🟠 MISSING: Value Stream "${vsText}"`);
        vsRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.STREAM_SIZE)
          .setBold(true)
          .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
          .setFontFamily(SLIDE_CONFIG.BODY_FONT)
          .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
      } else {
        vsRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.STREAM_SIZE)
          .setBold(true)
          .setForegroundColor(SLIDE_CONFIG.VALUE_STREAM_PURPLE)
          .setFontFamily(SLIDE_CONFIG.BODY_FONT);
      }
    } catch (e) {
      console.warn(`Could not style value stream: ${e.message}`);
    }
  });
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FIX VERSION HEADERS - Bold, black text
  // ═══════════════════════════════════════════════════════════════════════════
  if (meta.fixVersionHeaders) {
    meta.fixVersionHeaders.forEach(section => {
      try {
        const fvRange = textRange.getRange(section.start, section.end);
        fvRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE)
          .setBold(true)
          .setItalic(false)
          .setForegroundColor(SLIDE_CONFIG.BLACK)
          .setFontFamily(SLIDE_CONFIG.BODY_FONT);
      } catch (e) {
        console.warn(`Could not style fix version header: ${e.message}`);
      }
    });
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // ITERATION HEADERS (within fix version groups) - Italic, gray text
  // ═══════════════════════════════════════════════════════════════════════════
  if (meta.iterationHeaders) {
    meta.iterationHeaders.forEach(section => {
      try {
        const iterRange = textRange.getRange(section.start, section.end);
        iterRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE)
          .setBold(false)
          .setItalic(true)
          .setForegroundColor('#666666')
          .setFontFamily(SLIDE_CONFIG.BODY_FONT);
      } catch (e) {
        console.warn(`Could not style iteration header: ${e.message}`);
      }
    });
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // NO FIX VERSION HEADER - Bold, gray text
  // ═══════════════════════════════════════════════════════════════════════════
  if (meta.noFixVersionHeader) {
    try {
      const nfvRange = textRange.getRange(meta.noFixVersionHeader.start, meta.noFixVersionHeader.end);
      nfvRange.getTextStyle()
        .setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE)
        .setBold(true)
        .setItalic(false)
        .setForegroundColor('#666666')
        .setFontFamily(SLIDE_CONFIG.BODY_FONT);
    } catch (e) {
      console.warn(`Could not style no fix version header: ${e.message}`);
    }
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // EPIC SECTIONS
  // ═══════════════════════════════════════════════════════════════════════════
  meta.epicSections.forEach(section => {
    if (section.type === 'epic' && section.summaryEnd > section.summaryStart) {
      try {
        // ═══════════════════════════════════════════════════════════════════
        // BADGE STYLING - Supports multiple badges with CORRECT COLORS
        // ═══════════════════════════════════════════════════════════════════
        if (section.bulletStart !== undefined && section.bulletEnd !== undefined) {
          try {
            const badgeColors = {
              NEW: { bg: '#00BCD4', text: '#FFFFFF' },      // Teal/Cyan - Added this iteration
              CHG: { bg: '#FF9800', text: '#FFFFFF' },      // Orange - Modified since last
              DONE: { bg: '#4CAF50', text: '#FFFFFF' },     // Green - Completed/Closed
              PENDING: { bg: '#81C784', text: '#FFFFFF' },  // Light Green - Pending Closure
              DEF: { bg: '#F44336', text: '#FFFFFF' },      // Red - Deferred
              RISK: { bg: '#FFEB3B', text: '#000000' },     // Yellow - Iteration risk
              CANCELED: { bg: '#9E9E9E', text: '#FFFFFF' }, // Gray - Canceled/Won't Do
              OVERDUE: { bg: '#E65100', text: '#FFFFFF' }   // Burnt Orange - Due this iteration but not complete
            };
            const badges = section.badges || [];
            
            if (badges.length === 0) {
              // No badges - style as normal bullet/RAG indicator
              const bulletRange = textRange.getRange(section.bulletStart, section.bulletEnd);
              bulletRange.getTextStyle()
                .setFontSize(SLIDE_CONFIG.RAG_INDICATOR_SIZE);
            } else if (badges.length === 1) {
              // Single badge
              const bulletRange = textRange.getRange(section.bulletStart, section.bulletEnd);
              const colors = badgeColors[badges[0]];
              if (colors) {
                bulletRange.getTextStyle()
                  .setFontSize(SLIDE_CONFIG.RAG_INDICATOR_SIZE)
                  .setBold(true)
                  .setForegroundColor(colors.text)
                  .setBackgroundColor(colors.bg);
              }
            } else if (badges.length === 2) {
              // Two badges - style each separately (e.g., "CHGDONE")
              const badge1 = badges[0];
              const badge2 = badges[1];
              const badge1End = section.bulletStart + badge1.length;
              
              // First badge (CHG)
              const colors1 = badgeColors[badge1];
              if (colors1) {
                const range1 = textRange.getRange(section.bulletStart, badge1End);
                range1.getTextStyle()
                  .setFontSize(SLIDE_CONFIG.RAG_INDICATOR_SIZE)
                  .setBold(true)
                  .setForegroundColor(colors1.text)
                  .setBackgroundColor(colors1.bg);
              }
              
              // Second badge (DONE or DEF)
              const colors2 = badgeColors[badge2];
              if (colors2) {
                const range2 = textRange.getRange(badge1End, section.bulletEnd);
                range2.getTextStyle()
                  .setFontSize(SLIDE_CONFIG.RAG_INDICATOR_SIZE)
                  .setBold(true)
                  .setForegroundColor(colors2.text)
                  .setBackgroundColor(colors2.bg);
              }
            }
          } catch (e) {
            console.warn(`Badge styling error: ${e.message}`);
          }
        }
        
        const summaryRange = textRange.getRange(section.summaryStart, section.summaryEnd);
        const summaryText = fullText.substring(section.summaryStart, section.summaryEnd);
        
        // ═══════════════════════════════════════════════════════════════════
        // EPIC HYPERLINK - Apply FIRST (before styling)
        // ═══════════════════════════════════════════════════════════════════
        if (section.url && section.url.trim() !== '') {
          try {
            console.log(`🔗 EPIC: "${summaryText.substring(0, 40)}..." → ${section.url}`);
            summaryRange.getTextStyle().setLinkUrl(section.url);
          } catch (linkError) {
            console.error(`❌ EPIC link error: ${linkError.message}`);
          }
        } else {
          console.warn(`⚠️ EPIC: "${summaryText.substring(0, 40)}..." - No URL available`);
        }
        
        // Apply styling AFTER hyperlink (to override blue link color)
        if (section.isSummaryMissing || section.isNotCommitted) {
          if (section.isNotCommitted) console.log(`🟠 NOT COMMITTED epic`);
          summaryRange.getTextStyle()
            .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
            .setBold(true)
            .setUnderline(false)
            .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
        } else {
          summaryRange.getTextStyle()
            .setForegroundColor(SLIDE_CONFIG.BLACK)
            .setBold(false)
            .setUnderline(false);
        }
        
        // Style pipe | (purple, bold, not italic)
        if (section.pipeStart && section.pipeEnd) {
          try {
            const pipeRange = textRange.getRange(section.pipeStart, section.pipeEnd);
            pipeRange.getTextStyle()
              .setForegroundColor(SLIDE_CONFIG.PIPE_PURPLE)
              .setBold(true)
              .setItalic(false);
          } catch (e) {
            console.warn(`Could not style epic pipe: ${e.message}`);
          }
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // STRIKETHROUGH FOR CHANGED ITERATION
        // ═══════════════════════════════════════════════════════════════════
        if (section.hasStrikethroughIter && section.strikethroughStart && section.strikethroughEnd) {
          try {
            // Style the old iteration with strikethrough (gray)
            const oldIterRange = textRange.getRange(section.strikethroughStart, section.strikethroughEnd);
            oldIterRange.getTextStyle()
              .setStrikethrough(true)
              .setForegroundColor('#9E9E9E')  // Gray
              .setItalic(true);
            
            // Style the new iteration with normal formatting
            if (section.newIterStart && section.newIterStart < section.iterationEnd) {
              const newIterEnd = section.fixVersionsStart ? 
                (section.fixVersionsStart - 3) : section.iterationEnd;
              
              if (section.newIterStart < newIterEnd) {
                const newIterRange = textRange.getRange(section.newIterStart, newIterEnd);
                newIterRange.getTextStyle()
                  .setStrikethrough(false)
                  .setForegroundColor(SLIDE_CONFIG.BLACK)
                  .setBold(false)
                  .setItalic(true)
                  .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE);
              }
            }
            
            console.log(`⚡ Applied strikethrough to changed iteration`);
          } catch (e) {
            console.warn(`Could not apply iteration strikethrough: ${e.message}`);
          }
        } else {
          // Style iteration normally - orange if missing OR if dependency has later iteration
          if (section.iterationStart && section.iterationEnd && section.iterationStart < section.iterationEnd) {
            try {
              const iterRange = textRange.getRange(section.iterationStart, section.iterationEnd);
              
              // Use shouldHighlightMissingIter (false for closed items)
              if (section.shouldHighlightMissingIter || section.hasLaterDependency) {
                if (section.hasLaterDependency && !section.iterationMissing) {
                  console.log(`⚠️ SCHEDULING RISK: Epic iteration highlighted (dependency has later iteration)`);
                } else {
                  console.log(`🟠 MISSING: Epic iteration`);
                }
                iterRange.getTextStyle()
                  .setFontFamily(SLIDE_CONFIG.BODY_FONT)
                  .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE)
                  .setBold(false)
                  .setItalic(true)
                  .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
                  .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
              } else {
                iterRange.getTextStyle()
                  .setFontFamily(SLIDE_CONFIG.BODY_FONT)
                  .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE)
                  .setBold(false)
                  .setItalic(true)
                  .setForegroundColor(SLIDE_CONFIG.BLACK);
              }
            } catch (e) {
              console.warn(`Could not style epic iteration: ${e.message}`);
            }
          }
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // FIX VERSIONS STYLING
        // ═══════════════════════════════════════════════════════════════════
        if (section.fixVersionsStart && section.fixVersionsEnd && 
            section.fixVersionsStart < section.fixVersionsEnd) {
          try {
            if (section.hasStrikethroughFV && section.fvStrikethroughStart && section.fvStrikethroughEnd) {
              // Style the entire fix versions section first
              const fixRange = textRange.getRange(section.fixVersionsStart, section.fixVersionsEnd);
              fixRange.getTextStyle()
                .setForegroundColor('#616161')
                .setItalic(true)
                .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE);
              
              // Style the OLD Fix Version with strikethrough
              const oldFVRange = textRange.getRange(section.fvStrikethroughStart, section.fvStrikethroughEnd);
              oldFVRange.getTextStyle()
                .setStrikethrough(true)
                .setForegroundColor('#9E9E9E')
                .setItalic(true);
              
              // Style the NEW Fix Version normally
              if (section.fvNewStart && section.fvNewEnd && section.fvNewStart < section.fvNewEnd) {
                const newFVRange = textRange.getRange(section.fvNewStart, section.fvNewEnd);
                newFVRange.getTextStyle()
                  .setStrikethrough(false)
                  .setForegroundColor('#616161')
                  .setBold(false)
                  .setItalic(true);
              }
              
              console.log(`📦 Applied strikethrough to changed Fix Versions`);
            } else {
              // No change - style normally
              const fixRange = textRange.getRange(section.fixVersionsStart, section.fixVersionsEnd);
              fixRange.getTextStyle()
                .setForegroundColor('#616161')
                .setItalic(true)
                .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE);
              
              console.log(`📦 Styled Fix Versions`);
            }
          } catch (e) {
            console.warn(`Could not style fix versions: ${e.message}`);
          }
        }
        
        // ═══════════════════════════════════════════════════════════════════
        // *NEW* INDICATOR STYLING - Style " *new*" in italic green
        // ═══════════════════════════════════════════════════════════════════
        if (section.hasNewIterIndicator) {
          try {
            const newIndicatorLength = 6;  // " *new*" is 6 characters
            const iterTextEnd = section.fixVersionsStart ? section.fixVersionsStart : section.iterationEnd;
            const newIndicatorStart = iterTextEnd - newIndicatorLength;
            
            if (newIndicatorStart > section.iterationStart) {
              const newRange = textRange.getRange(newIndicatorStart, iterTextEnd);
              newRange.getTextStyle()
                .setItalic(true)
                .setBold(false)
                .setForegroundColor('#2E7D32')  // Dark green
                .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE);
              console.log(`✨ Styled iteration *new* indicator`);
            }
          } catch (e) {
            console.warn(`Could not style iteration *new* indicator: ${e.message}`);
          }
        }
        
        if (section.hasNewFVIndicator && section.fixVersionsStart && section.fixVersionsEnd) {
          try {
            const newIndicatorLength = 6;
            const newIndicatorStart = section.fixVersionsEnd - newIndicatorLength;
            
            if (newIndicatorStart > section.fixVersionsStart) {
              const newRange = textRange.getRange(newIndicatorStart, section.fixVersionsEnd);
              newRange.getTextStyle()
                .setItalic(true)
                .setBold(false)
                .setForegroundColor('#2E7D32')  // Dark green
                .setFontSize(SLIDE_CONFIG.EPIC_ITERATION_SIZE);
              console.log(`✨ Styled fix version *new* indicator`);
            }
          } catch (e) {
            console.warn(`Could not style fix version *new* indicator: ${e.message}`);
          }
        }
        
      } catch (e) {
        console.warn(`Could not style epic section: ${e.message}`);
      }
    }
  });

  // ═══════════════════════════════════════════════════════════════════════════
  // DEPENDENCY SECTIONS
  // ═══════════════════════════════════════════════════════════════════════════
  meta.dependencySections.forEach(section => {
    try {
      // STEP 1: Style entire dependency line first (sets base styling)
      const depLineRange = textRange.getRange(section.start, section.end);
      depLineRange.getTextStyle()
        .setFontFamily(SLIDE_CONFIG.BODY_FONT)
        .setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE)
        .setBold(false)
        .setItalic(true)
        .setForegroundColor(SLIDE_CONFIG.DEP_GRAY);
      
      // STEP 2: Style bullet/icon (smaller size, not italic)
      if (section.bulletStart !== undefined && section.bulletEnd !== undefined) {
        try {
          const bulletRange = textRange.getRange(section.bulletStart, section.bulletEnd);
          bulletRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.DEP_ICON_SIZE)
            .setItalic(false)
            .setForegroundColor(SLIDE_CONFIG.DEP_GRAY);
        } catch (e) {
          console.warn(`Could not style dependency bullet: ${e.message}`);
        }
      }
      
      // STEP 3: Style summary text
      if (section.summaryEnd > section.summaryStart) {
        try {
          const summaryRange = textRange.getRange(section.summaryStart, section.summaryEnd);
          summaryRange.getTextStyle()
            .setUnderline(false)
            .setForegroundColor(SLIDE_CONFIG.DEP_GRAY);
        } catch (e) {
          console.warn(`Could not style dependency summary: ${e.message}`);
        }
      }
      
      // STEP 4: Style first pipe |
      if (section.pipe1Start !== undefined && section.pipe1End !== undefined) {
        try {
          const pipe1Range = textRange.getRange(section.pipe1Start, section.pipe1End);
          pipe1Range.getTextStyle()
            .setForegroundColor(SLIDE_CONFIG.PIPE_PURPLE)
            .setBold(true)
            .setItalic(false);
        } catch (e) {
          console.warn(`Could not style dependency pipe1: ${e.message}`);
        }
      }
      
      // STEP 5: Style VS
      if (section.vsStart !== undefined && section.vsEnd !== undefined) {
        try {
          const vsRange = textRange.getRange(section.vsStart, section.vsEnd);
          
          if (section.vsMissing) {
            console.log(`🟠 MISSING: Dependency VS`);
            vsRange.getTextStyle()
              .setBold(true)
              .setItalic(false)
              .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
              .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
          } else {
            vsRange.getTextStyle()
              .setBold(false)
              .setItalic(false)
              .setForegroundColor(SLIDE_CONFIG.DEP_GRAY);
          }
        } catch (e) {
          console.warn(`Could not style dependency VS: ${e.message}`);
        }
      }
      
      // STEP 6: Style second pipe |
      if (section.pipe2Start !== undefined && section.pipe2End !== undefined) {
        try {
          const pipe2Range = textRange.getRange(section.pipe2Start, section.pipe2End);
          pipe2Range.getTextStyle()
            .setForegroundColor(SLIDE_CONFIG.PIPE_PURPLE)
            .setBold(true)
            .setItalic(false);
        } catch (e) {
          console.warn(`Could not style dependency pipe2: ${e.message}`);
        }
      }
      
      // STEP 7: Style iteration
      if (section.iterationStart !== undefined && section.iterationEnd !== undefined) {
        try {
          const iterRange = textRange.getRange(section.iterationStart, section.iterationEnd);
          
          if (section.iterationMissing || section.iterationAfterEpic) {
            if (section.iterationAfterEpic && !section.iterationMissing) {
              console.log(`⚠️ SCHEDULING RISK: Dependency iteration highlighted (after parent epic)`);
            } else {
              console.log(`🟠 MISSING: Dependency iteration`);
            }
            iterRange.getTextStyle()
              .setFontSize(SLIDE_CONFIG.DEP_ITERATION_SIZE)
              .setBold(false)
              .setItalic(true)
              .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
              .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
          } else {
            iterRange.getTextStyle()
              .setFontSize(SLIDE_CONFIG.DEP_ITERATION_SIZE)
              .setBold(false)
              .setItalic(true)
              .setForegroundColor(SLIDE_CONFIG.DEP_GRAY);
          }
        } catch (e) {
          console.warn(`Could not style dependency iteration: ${e.message}`);
        }
      }
      
      // DEPENDENCY HYPERLINK
      const depSummaryText = fullText.substring(section.summaryStart, section.summaryEnd);
      if (section.key && section.key.trim() !== '' && section.summaryEnd > section.summaryStart) {
        try {
          const linkUrl = getJiraUrl(section.key);
          if (linkUrl) {
            console.log(`🔗 DEPENDENCY: "${depSummaryText.substring(0, 40)}..." → ${section.key}`);
            const summaryRange = textRange.getRange(section.summaryStart, section.summaryEnd);
            summaryRange.getTextStyle().setLinkUrl(linkUrl);
            summaryRange.getTextStyle()
              .setForegroundColor(SLIDE_CONFIG.DEP_GRAY)
              .setUnderline(false);
          } else {
            console.warn(`⚠️ DEPENDENCY: "${depSummaryText.substring(0, 40)}..." - getJiraUrl returned null for "${section.key}"`);
          }
        } catch (linkError) {
          console.error(`❌ DEPENDENCY link error for ${section.key}: ${linkError.message}`);
        }
      } else {
        console.warn(`⚠️ DEPENDENCY: "${depSummaryText.substring(0, 40)}..." - No key available`);
      }
      
      // DEPENDENCY STRIKETHROUGH FOR CHANGED ITERATION
      if (section.hasStrikethroughIter && section.strikethroughIterStart && section.strikethroughIterEnd) {
        try {
          const oldIterRange = textRange.getRange(section.strikethroughIterStart, section.strikethroughIterEnd);
          oldIterRange.getTextStyle()
            .setStrikethrough(true)
            .setForegroundColor('#9E9E9E');
          
          if (section.newIterStart && section.iterationEnd && section.newIterStart < section.iterationEnd) {
            const newIterRange = textRange.getRange(section.newIterStart, section.iterationEnd);
            newIterRange.getTextStyle()
              .setStrikethrough(false)
              .setForegroundColor(SLIDE_CONFIG.DEP_GRAY)
              .setBold(false)
              .setItalic(true)
              .setFontSize(SLIDE_CONFIG.DEP_ITERATION_SIZE);
          }
          
          console.log(`⚡ Dep ${section.key}: iteration strikethrough applied`);
        } catch (e) {
          console.warn(`Could not apply dep iteration strikethrough: ${e.message}`);
        }
      }
      
      // DEPENDENCY STRIKETHROUGH FOR CHANGED FIX VERSIONS
      if (section.hasStrikethroughFV && section.fvStrikethroughStart && section.fvStrikethroughEnd) {
        try {
          const oldFVRange = textRange.getRange(section.fvStrikethroughStart, section.fvStrikethroughEnd);
          oldFVRange.getTextStyle()
            .setStrikethrough(true)
            .setForegroundColor('#9E9E9E');
          
          if (section.fvNewStart && section.fvNewEnd && section.fvNewStart < section.fvNewEnd) {
            const newFVRange = textRange.getRange(section.fvNewStart, section.fvNewEnd);
            newFVRange.getTextStyle()
              .setStrikethrough(false)
              .setForegroundColor(SLIDE_CONFIG.DEP_GRAY)
              .setBold(false)
              .setItalic(true);
          }
          
          console.log(`📦 Dep ${section.key}: fix versions strikethrough applied`);
        } catch (e) {
          console.warn(`Could not apply dep FV strikethrough: ${e.message}`);
        }
      }
      
    } catch (e) {
      console.warn(`Could not style dependency section: ${e.message}`);
    }
  });
  
  // ═══════════════════════════════════════════════════════════════════════════
  // RISK SECTIONS
  // ═══════════════════════════════════════════════════════════════════════════
  meta.riskSections.forEach(section => {
    try {
      // Handle label-only lines (for multi-line RAG notes)
      if (section.isLabelOnly) {
        const labelRange = textRange.getRange(section.labelStart, section.labelEnd);
        labelRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
          .setBold(true)
          .setForegroundColor(SLIDE_CONFIG.BLUE);
        return;
      }
      
      // Handle notes with iteration labels
      if (section.type === 'ragNoteWithIter') {
        if (section.labelStart && section.labelEnd) {
          const iterLabelRange = textRange.getRange(section.labelStart, section.labelEnd);
          iterLabelRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
            .setBold(true)
            .setForegroundColor(section.isCurrent ? SLIDE_CONFIG.BLUE : '#757575');
        }
        
        if (section.noteStart && section.noteEnd) {
          const noteRange = textRange.getRange(section.noteStart, section.noteEnd);
          if (section.isPreviousNote || !section.isCurrent) {
            noteRange.getTextStyle()
              .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
              .setItalic(true)
              .setStrikethrough(true)
              .setForegroundColor('#9E9E9E');
            console.log(`  📝 Previous RAG note (Iter ${section.iteration}): strikethrough applied`);
          } else {
            noteRange.getTextStyle()
              .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
              .setItalic(true)
              .setStrikethrough(false)
              .setForegroundColor(SLIDE_CONFIG.BLACK);
          }
        }
        return;
      }
      
      // Style "Risk - " label (bold, blue)
      if (section.labelStart && section.labelEnd) {
        try {
          const labelRange = textRange.getRange(section.labelStart, section.labelEnd);
          labelRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
            .setBold(true)
            .setForegroundColor(SLIDE_CONFIG.BLUE);
        } catch (e) {
          console.warn(`Could not style risk label: ${e.message}`);
        }
      }
      
      // Style note (italics, orange if missing)
      if (section.noteStart && section.noteEnd) {
        try {
          const noteRange = textRange.getRange(section.noteStart, section.noteEnd);
          
          if (section.isMissing) {
            console.log(`🟠 MISSING: RAG note`);
            noteRange.getTextStyle()
              .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
              .setBold(true)
              .setItalic(false)
              .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
              .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT);
          } else {
            noteRange.getTextStyle()
              .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
              .setBold(false)
              .setItalic(true)
              .setForegroundColor(SLIDE_CONFIG.BLUE);
          }
        } catch (e) {
          console.warn(`Could not style risk note: ${e.message}`);
        }
      }
    } catch (e) {
      console.warn(`Could not style risk section: ${e.message}`);
    }
  });
  
  // ═══════════════════════════════════════════════════════════════════════════
  // NOTE SECTIONS - Informational notes
  // Uses blue styling for positive notes (mitigation, completion)
  // Uses gray styling for regular informational notes
  // ═══════════════════════════════════════════════════════════════════════════
  if (meta.noteSections) {
    meta.noteSections.forEach(section => {
      try {
        // Determine color based on note type
        const noteColor = section.isPositiveNote ? SLIDE_CONFIG.BLUE : '#666666';
        
        if (section.labelStart && section.labelEnd) {
          const labelRange = textRange.getRange(section.labelStart, section.labelEnd);
          labelRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
            .setBold(true)
            .setForegroundColor(noteColor);
        }
        
        if (section.noteStart && section.noteEnd) {
          const noteRange = textRange.getRange(section.noteStart, section.noteEnd);
          noteRange.getTextStyle()
            .setFontSize(SLIDE_CONFIG.RISK_LABEL_SIZE)
            .setBold(false)
            .setItalic(true)
            .setForegroundColor(noteColor);
        }
      } catch (e) {
        console.warn(`Could not style note section: ${e.message}`);
      }
    });
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SEPARATOR SECTIONS
  // ═══════════════════════════════════════════════════════════════════════════
  if (meta.separatorSections) {
    meta.separatorSections.forEach(section => {
      try {
        const sepRange = textRange.getRange(section.start, section.end);
        sepRange.getTextStyle()
          .setFontSize(SLIDE_CONFIG.INITIATIVE_SEPARATOR_SIZE);
      } catch (e) {
        console.warn(`Could not style separator: ${e.message}`);
      }
    });
  }
  
  console.log(`=== STYLING COMPLETE ===\n`);
}
function processDataForTableSlides(data, showDependencies = true, hideNoInitiative = false) {
  const portfolioData = {};
  const piNumber = data.metadata?.piNumber;
  
  console.log(`Processing data for PI ${piNumber}, filtering out Pre-Planning and other PIs...`);
  console.log(`Show Dependencies: ${showDependencies}`);
  console.log(`Hide No Initiative: ${hideNoInitiative}`);
  console.log(`Total dependencies in data: ${data.dependencies ? data.dependencies.length : 0}`);
  
  if (data.dependencies && data.dependencies.length > 0) {
    const sampleDep = data.dependencies[0];
    console.log(`Sample dependency keys: ${Object.keys(sampleDep).join(', ')}`);
    console.log(`Sample dependency Parent Key: "${sampleDep['Parent Key']}"`);
  }
  
  showProgress(`Filtering ${data.epics.length} epics for PI ${piNumber}...`, '📊 Processing Data');
  
  let filteredEpics = data.epics;
  
  // Filter by Program Increment
  if (piNumber) {
    const expectedPI = `PI ${piNumber}`;
    const beforeCount = filteredEpics.length;
    
    filteredEpics = data.epics.filter(epic => {
      const programIncrement = (epic['Program Increment'] || '').toString().trim();
      const initiativeTitle = (epic['Initiative Title'] || '').toString().trim();
      
      const isMatchingPI = programIncrement === expectedPI || 
                          programIncrement === piNumber.toString() ||
                          programIncrement === '';
      
      const isPIPrePlanning = programIncrement.toLowerCase().includes('pre-planning') ||
                              programIncrement.toLowerCase().includes('preplanning');
      
      const isInitiativePrePlanning = initiativeTitle.toLowerCase().includes('pre-planning') ||
                                      initiativeTitle.toLowerCase().includes('preplanning');
      
      const differentPIPattern = new RegExp(`PI\\s*(\\d+)`, 'gi');
      let isDifferentPI = false;
      let match;
      while ((match = differentPIPattern.exec(initiativeTitle)) !== null) {
        const foundPI = parseInt(match[1]);
        if (foundPI !== piNumber) {
          isDifferentPI = true;
          break;
        }
      }
      
      return isMatchingPI && !isPIPrePlanning && !isInitiativePrePlanning && !isDifferentPI;
    });
    
    console.log(`Program Increment filter: ${beforeCount} -> ${filteredEpics.length} epics (filtered out ${beforeCount - filteredEpics.length} non-PI ${piNumber} items)`);
  }
  
  // Filter by PI Commitment
  // IMPORTANT: Do NOT filter out items with blank PI Commitment if they are:
  // - Already closed (continuity) - epic.alreadyClosed = true
  // - Already deferred (continuity) - epic.alreadyDeferred = true
  // - Closed/Done status (may have cleared commitment)
  // - Deferred this iteration (transitioning out)
  // These items were explicitly included by filterByChanges and should appear on slides
  const beforeCommitmentFilter = filteredEpics.length;
  filteredEpics = filteredEpics.filter(epic => {
    const piCommitment = (epic['PI Commitment'] || '').toString().trim();
    const status = (epic['Status'] || '').toString().toUpperCase();
    const isClosed = status === 'CLOSED' || status === 'DONE';
    
    // Keep the epic if:
    // 1. Has a PI Commitment value, OR
    // 2. Is already closed (continuity), OR
    // 3. Is already deferred (continuity), OR
    // 4. Status is Closed/Done (may have cleared commitment), OR
    // 5. Was closed this iteration, OR
    // 6. Was deferred this iteration
    const keepForContinuity = epic.alreadyClosed || 
                               epic.alreadyDeferred || 
                               epic.closedThisIteration || 
                               epic.deferredThisIteration ||
                               isClosed;
    
    if (piCommitment !== '') {
      return true;  // Has commitment - keep it
    } else if (keepForContinuity) {
      console.log(`  📌 Keeping ${epic['Key']} despite blank PI Commitment (continuity: alreadyClosed=${epic.alreadyClosed}, alreadyDeferred=${epic.alreadyDeferred}, isClosed=${isClosed})`);
      return true;  // Continuity item - keep it
    } else {
      return false;  // No commitment and not a continuity item - filter out
    }
  });
  console.log(`PI Commitment filter: ${beforeCommitmentFilter} -> ${filteredEpics.length} epics (filtered out ${beforeCommitmentFilter - filteredEpics.length} with blank PI Commitment, kept continuity items)`);
  
  // Filter Value Streams (except INFOSEC)
  const valueStreamExclusions = typeof VALUE_STREAM_EXCLUDE_EXACT !== 'undefined' 
    ? VALUE_STREAM_EXCLUDE_EXACT 
    : ['Xtract'];
  
  if (valueStreamExclusions.length > 0) {
    const beforeVSFilter = filteredEpics.length;
    filteredEpics = filteredEpics.filter(epic => {
      const valueStream = epic['Value Stream/Org'] || '';
      const portfolioInitiative = epic['Portfolio Initiative'] || '';
      
      // INFOSEC bypasses ALL allocation/value stream filtering
      if (portfolioInitiative.startsWith('INFOSEC')) {

        return true;
      }
      
      return !valueStreamExclusions.includes(valueStream);
    });
    
    const vsExcluded = beforeVSFilter - filteredEpics.length;
    if (vsExcluded > 0) {
      console.log(`Value Stream exclusions: Removed ${vsExcluded} epics from excluded Value Streams (${valueStreamExclusions.join(', ')})`);
    }
  }
  
  showProgress(`Grouping ${filteredEpics.length} epics by portfolio and matching dependencies...`, '📊 Processing Data');
  
  // Group epics by Portfolio > Program > Initiative
  filteredEpics.forEach(epic => {
    const portfolioName = epic['Portfolio Initiative'];
    if (!portfolioName) return;
    
    const programName = epic['Program Initiative'] || 'No Program Initiative';
    const initiativeTitle = (epic['Initiative Title'] || '').trim() || 'No Initiative';
    const parentKey = epic['Parent Key'] || epic['Key'];
    
    if (!portfolioData[portfolioName]) {
      portfolioData[portfolioName] = {
        programs: {},
        totalFeaturePoints: 0
      };
    }
    
    if (!portfolioData[portfolioName].programs[programName]) {
      portfolioData[portfolioName].programs[programName] = {
        initiatives: {}
      };
    }
    
    if (!portfolioData[portfolioName].programs[programName].initiatives[initiativeTitle]) {
      portfolioData[portfolioName].programs[programName].initiatives[initiativeTitle] = {
        name: initiativeTitle,
        linkKey: parentKey,
        fixVersion: epic['Fix Versions'] || '',
        businessPerspective: epic['Business Perspective'] || '',
        technicalPerspective: epic['Technical Perspective'] || '',
        epics: [],
        totalFeaturePoints: 0
      };
    } else {
      const existingInit = portfolioData[portfolioName].programs[programName].initiatives[initiativeTitle];
      const epicParentKey = epic['Parent Key'];
      const epicKey = epic['Key'];
      
      if (epicParentKey && epicParentKey.trim() !== '') {
        if (!existingInit.linkKey || existingInit.linkKey.trim() === '' || existingInit.linkKey === epicKey) {
          existingInit.linkKey = epicParentKey;
        }
      }
    }
    
    const initiative = portfolioData[portfolioName].programs[programName].initiatives[initiativeTitle];
    
    let epicDeps = [];
    if (showDependencies && data.dependencies) {
      epicDeps = data.dependencies.filter(dep => dep['Parent Key'] === epic['Key']);
      
      if (epicDeps.length > 0) {
        console.log(`Epic ${epic['Key']}: Found ${epicDeps.length} dependencies`);
      }
      
      if (piNumber && epicDeps.length > 0) {
        const beforePIFilter = epicDeps.length;
        const expectedPI = `PI ${piNumber}`;
        epicDeps = epicDeps.filter(dep => {
          const depPI = (dep['Program Increment'] || '').toString().trim();
          const isMatchingPI = depPI === expectedPI || 
                              depPI === piNumber.toString() ||
                              depPI === '';
          const isPrePlanning = depPI.toLowerCase().includes('pre-planning');
          return isMatchingPI && !isPrePlanning;
        });
        
        if (beforePIFilter !== epicDeps.length) {
          console.log(`  -> PI filter: ${beforePIFilter} -> ${epicDeps.length} dependencies`);
        }
      }
    }
    
    initiative.epics.push({
      ...epic,
      dependencies: epicDeps
    });
    
    const fp = parseFloat(epic['Feature Points']) || 0;
    initiative.totalFeaturePoints += fp;
    portfolioData[portfolioName].totalFeaturePoints += fp;
  });
  
  // Log summary
  let totalDepsAttached = 0;
  let epicsWithDeps = 0;
  Object.values(portfolioData).forEach(portfolio => {
    Object.values(portfolio.programs).forEach(program => {
      Object.values(program.initiatives).forEach(initiative => {
        initiative.epics.forEach(epic => {
          if (epic.dependencies && epic.dependencies.length > 0) {
            totalDepsAttached += epic.dependencies.length;
            epicsWithDeps++;
          }
        });
      });
    });
  });
  console.log(`Dependencies summary: ${totalDepsAttached} dependencies attached to ${epicsWithDeps} epics`);
  
  return portfolioData;
}
// =================================================================================
// CONTENT LINE BUILDING
// =================================================================================

function buildContentLinesFromPrograms(programRows, showDependencies, noInitiativeMode = 'show') {
  const contentLines = [];
  
  programRows.forEach((programRow) => {
    const programName = programRow.programName;
    
    programRow.initiatives.forEach(initiative => {
      const initName = initiative.name || initiative.title || 'Unnamed';
      const isNoInitiative = initName === 'No Initiative' || initName === 'No Initiative Associated';
      const allEpics = initiative.epics || [];
      const displayableEpics = allEpics.filter(epic => {
        const data = processEpicData(epic);
        return !data.summaryIsEmpty;
      });
      
      // Skip entire initiative if no displayable epics
      if (displayableEpics.length === 0) {
        console.log(`  ⏭️ Skipping empty initiative: "${initName}" (no displayable epics)`);
        return;  // Skip this initiative entirely
      }
      
      const epicCount = displayableEpics.length;
      const skipInitiativeHeader = isNoInitiative && noInitiativeMode === 'skip';
      
      if (!skipInitiativeHeader) {
        contentLines.push({
          type: 'initiative',
          lines: 1,
          program: programName,
          initiativeObj: initiative,
          initiativeName: initName,
          featurePoints: initiative.totalFeaturePoints || 0,
          isNoInitiative: isNoInitiative,
          epicCount: epicCount
        });
        
        const perspective = initiative.technicalPerspective || '';
        if (perspective && epicCount >= 2) {
          contentLines.push({
            type: 'perspective',
            lines: 2,
            program: programName,
            initiativeObj: initiative,
            text: perspective
          });
        }
      }
      
      // Group epics by Value Stream
      const valueStreamGroups = {};
      displayableEpics.forEach(epic => {
        const vs = epic['Value Stream/Org'] || 'Other';
        if (!valueStreamGroups[vs]) valueStreamGroups[vs] = [];
        valueStreamGroups[vs].push(epic);
      });
      
      const sortedVS = Object.keys(valueStreamGroups).sort();
      
      sortedVS.forEach(vsName => {
        const vsEpics = valueStreamGroups[vsName];
        
        // ═══════════════════════════════════════════════════════════════════════
        // FIX: Double-check value stream has displayable epics
        // (Should always be true now, but defensive check)
        // ═══════════════════════════════════════════════════════════════════════
        if (vsEpics.length === 0) {
          console.log(`  ⏭️ Skipping empty value stream: ${vsName}`);
          return;
        }
        
        contentLines.push({
          type: 'valueStream',
          lines: 1,
          program: programName,
          initiativeObj: initiative,
          valueStream: vsName
        });
        
        // Sort epics by iteration
        vsEpics.sort((a, b) => {
          const iterA = (a['End Iteration Name'] || a['PI Target Iteration'] || '').toString();
          const iterB = (b['End Iteration Name'] || b['PI Target Iteration'] || '').toString();
          return iterA.localeCompare(iterB);
        });
        
        vsEpics.forEach(epic => {
          const epicData = processEpicData(epic);
          
          // This check is now redundant but kept for safety
          if (epicData.summaryIsEmpty) return;
          
          contentLines.push({
            type: 'epic',
            lines: 1,
            program: programName,
            initiativeObj: initiative,
            valueStream: vsName,
            epic: epic,
            epicData: epicData
          });
          
          if (epicData.showRag) {
            contentLines.push({
              type: 'epicRag',
              lines: 1,
              program: programName,
              initiativeObj: initiative,
              valueStream: vsName,
              epic: epic,
              epicData: epicData
            });
          }
          
          if (showDependencies && epicData.dependencies && epicData.dependencies.length > 0) {
            const sortedDeps = [...epicData.dependencies].sort((a, b) => {
              const iterA = (a['End Iteration Name'] || '').toString();
              const iterB = (b['End Iteration Name'] || '').toString();
              return iterA.localeCompare(iterB);
            });
            
            sortedDeps.forEach(dep => {
              contentLines.push({
                type: 'dependency',
                lines: 1,
                program: programName,
                initiativeObj: initiative,
                valueStream: vsName,
                epic: epic,
                dependency: dep
              });
              
              const depRag = getRagIndicator(dep['RAG']);
              if (depRag) {
                contentLines.push({
                  type: 'depRag',
                  lines: 1,
                  program: programName,
                  initiativeObj: initiative,
                  valueStream: vsName,
                  epic: epic,
                  dependency: dep
                });
              }
            });
          }
        });
      });
    });
  });
  
  return contentLines;
}

function buildInitiativesFromContentBatch(batch, showDependencies, continuationContext) {
  const initiatives = [];
  let currentInitiative = null;
  let isFirstItem = true;
  let continuedValueStream = null;
  
  // Handle continuation from previous slide
  if (continuationContext && continuationContext.initiative && batch.length > 0 && batch[0].type !== 'initiative') {
    const contInit = continuationContext.initiative;
    currentInitiative = {
      name: `${contInit.name || contInit.title || 'Unknown'} (continued)`,
      totalFeaturePoints: contInit.totalFeaturePoints || 0,
      technicalPerspective: '',
      businessPerspective: '',
      epics: [],
      linkKey: contInit.linkKey || null,
      isContinuation: true
    };
    initiatives.push(currentInitiative);
    
    if (continuationContext.valueStream) {
      continuedValueStream = continuationContext.valueStream;
    }
  }
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // FIX: Create default initiative if batch has items but no 'initiative' type
  // This handles the case when noInitiativeMode = 'skip' causes all 'initiative' 
  // items to be skipped (because all initiatives are named "No Initiative")
  // ═══════════════════════════════════════════════════════════════════════════════
  if (!currentInitiative && batch.length > 0 && batch[0].type !== 'initiative') {
    // Get initiative info from the first item if available
    const firstItem = batch[0];
    const initiativeObj = firstItem.initiativeObj;
    
    currentInitiative = {
      name: initiativeObj ? (initiativeObj.name || initiativeObj.title || 'No Initiative') : 'No Initiative',
      totalFeaturePoints: initiativeObj ? (initiativeObj.totalFeaturePoints || 0) : 0,
      technicalPerspective: initiativeObj ? (initiativeObj.technicalPerspective || '') : '',
      businessPerspective: initiativeObj ? (initiativeObj.businessPerspective || '') : '',
      epics: [],
      linkKey: initiativeObj ? initiativeObj.linkKey : null,
      isDefaultInitiative: true  // Flag to indicate this was auto-created
    };
    initiatives.push(currentInitiative);
    
    console.log(`  ⚡ Created default initiative for batch starting with '${firstItem.type}' (noInitiativeMode may have skipped header)`);
  }
  // ═══════════════════════════════════════════════════════════════════════════════
  
  batch.forEach(item => {
    if (item.type === 'initiative') {
      currentInitiative = {
        name: item.initiativeName,
        totalFeaturePoints: item.featurePoints,
        technicalPerspective: '',
        businessPerspective: '',
        epics: [],
        linkKey: item.initiativeObj ? item.initiativeObj.linkKey : null
      };
      initiatives.push(currentInitiative);
      continuedValueStream = null;
    } else if (item.type === 'perspective' && currentInitiative) {
      currentInitiative.technicalPerspective = item.text;
    } else if (item.type === 'valueStream' && currentInitiative) {
      if (isFirstItem && continuedValueStream === item.valueStream) {
        item.isContinuedValueStream = true;
      }
      continuedValueStream = null;
    } else if (item.type === 'epic' && currentInitiative) {
      const epicWithMeta = {
        ...item.epic,
        dependencies: [],
        _valueStreamContinued: isFirstItem && continuationContext && continuationContext.valueStream === item.valueStream
      };
      currentInitiative.epics.push(epicWithMeta);
    } else if (item.type === 'dependency' && currentInitiative && currentInitiative.epics.length > 0) {
      const lastEpic = currentInitiative.epics[currentInitiative.epics.length - 1];
      if (!lastEpic.dependencies) lastEpic.dependencies = [];
      lastEpic.dependencies.push(item.dependency);
    }
    
    isFirstItem = false;
  });
  
  return initiatives;
}

// =================================================================================
// EPIC DATA PROCESSING
// =================================================================================

function processEpicData(epic) {
  const epicData = {
    key: epic['Key'],
    summary: getFieldDisplayValue(epic['Summary'], 'summary', 'Unnamed Epic') || 'Unnamed Epic',
    summaryIsEmpty: getFieldDisplayValue(epic['Summary'], 'summary', 'Unnamed Epic') === null,
    rag: epic['RAG'] || '',
    ragNote: epic['RAG Note'] || '',
    endIteration: epic['End Iteration Name'] || epic['PI Target Iteration'] || '',
    dependencies: epic.dependencies || [],
    url: getJiraUrl(epic['Key']),
    status: epic['Status'] || '',
    fixVersions: epic['Fix Versions'] || ''
  };
  
const status = (epic['Status'] || '').toString().toUpperCase();
  epicData.isClosed = status === 'CLOSED' || status === 'DONE';
  
  // Check for Pending Acceptance status
  const isPendingAcceptanceStatus = status === 'PENDING ACCEPTANCE';
  
  const piCommitment = (epic['PI Commitment'] || '').toString().trim();
  
  // Valid commitment values - epic is NOT deferred if it has one of these
  const validCommitments = ['Committed', 'Committed After Plan', 'Committed After Planning'];
  // Treat "Not Committed", "Deferred" explicitly as deferred status
  const deferredValues = ['Not Committed', 'Deferred'];
  // Treat "Canceled" as canceled status (separate from deferred)
  const canceledValues = ['Canceled'];
  
  const isCanceledByCommitment = canceledValues.some(v => v.toLowerCase() === piCommitment.toLowerCase());
  const isDeferredByCommitment = piCommitment !== '' && 
                     (!validCommitments.includes(piCommitment) && 
                      !isCanceledByCommitment &&
                      deferredValues.some(v => v.toLowerCase() === piCommitment.toLowerCase()));
  
  // Set flags based on PI Commitment
  if (isCanceledByCommitment) {
    epicData.ragIndicator = '🔴';
    epicData.showRag = true;
    epicData.isCanceled = true;
    epicData.ragNote = 'canceled for this PI';
  } else if (isDeferredByCommitment) {
    epicData.ragIndicator = '🔴';
    epicData.showRag = true;
    epicData.isDeferred = true;
    epicData.ragNote = 'deferred for this PI';
  } else if (isPendingAcceptanceStatus) {
    // Pending Acceptance gets its own visual treatment
    epicData.ragIndicator = getRagIndicator(epicData.rag);
    epicData.showRag = epicData.ragIndicator !== '';
    epicData.isPendingClosure = true;
  } else {
    epicData.ragIndicator = getRagIndicator(epicData.rag);
    epicData.showRag = epicData.ragIndicator !== '';
    epicData.isDeferred = false;
  }
  
  // Handle RAG Note without RAG status - show as "Note" instead of "Risk"
  const ragNoteText = (epicData.ragNote || '').trim();
  const isJustStatusIndicator = ragNoteText.toLowerCase() === 'green' || 
                                 ragNoteText.toLowerCase().startsWith('green -') ||
                                 ragNoteText.toLowerCase().startsWith('green:') ||
                                 ragNoteText.toLowerCase() === 'on track' ||
                                 ragNoteText.toLowerCase() === 'no issues';
  
  if (!epicData.showRag && ragNoteText !== '' && !isJustStatusIndicator) {
    epicData.showNote = true;
    epicData.noteText = ragNoteText;
  } else {
    epicData.showNote = false;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // Use pre-computed badge from changelog as the source of truth
  // ═══════════════════════════════════════════════════════════════════════════════
  const badge = (epic.badge || '').toUpperCase();
  
  // Check if epic is NEW
  epicData.isNew = epic.isNew || badge === 'NEW' || false;
  
  // Check if epic is MODIFIED
  epicData.isModified = epic.isModified || badge === 'CHG' || false;
  
  // Check if epic is DONE (closed)
  if (!epicData.isClosed && badge === 'DONE') {
    epicData.isClosed = true;
  }
  
  // Check if epic is PENDING CLOSURE
  if (!epicData.isPendingClosure && badge === 'PENDING') {
    epicData.isPendingClosure = true;
  }
  epicData.pendingClosureThisIteration = epic.pendingClosureThisIteration || false;
  epicData.alreadyPendingClosure = epic.alreadyPendingClosure || false;
  
  // Check if epic is DEFERRED
  if (!epicData.isDeferred && badge === 'DEF') {
    epicData.isDeferred = true;
    epicData.ragIndicator = '🔴';
    epicData.showRag = true;
    epicData.ragNote = 'deferred for this PI';
  }
  
  // Pass through transition flags and change details
  epicData.closedThisIteration = epic.closedThisIteration || false;
  epicData.deferredThisIteration = epic.deferredThisIteration || false;
  epicData.changes = epic.changes || {};
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // CANCELED handling - check both flag AND badge
  // ═══════════════════════════════════════════════════════════════════════════════
  if (!epicData.isCanceled && badge === 'CANCELED') {
    epicData.isCanceled = true;
    epicData.ragIndicator = '🔴';
    epicData.showRag = true;
    epicData.ragNote = 'canceled for this PI';
  }
  epicData.canceledReason = epic.canceledReason || '';
  epicData.canceledThisIteration = epic.canceledThisIteration || false;
  epicData.alreadyCanceled = epic.alreadyCanceled || false;
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // OVERDUE handling - check both flag AND badge
  // ═══════════════════════════════════════════════════════════════════════════════
  epicData.isOverdue = epic.isOverdue || epic.isIterationRisk || badge === 'OVERDUE' || false;
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // AT-RISK handling - check both flag AND badge
  // ═══════════════════════════════════════════════════════════════════════════════
  epicData.isAtRisk = epic.isAtRisk || badge === 'ATRISK' || badge === 'AT-RISK' || false;
  
  return epicData;
}
// =================================================================================
// CLEANUP
// =================================================================================

// =================================================================================
// NOTE: getFilteredAndSortedPortfolios is defined in MENUCONFIG.gs
// It handles portfolio ordering and filtering based on user configuration
// =================================================================================

function cleanupTemplateSlides(presentation, totalPortfolios) {
  console.log('\n=== CLEANUP: Starting template slide cleanup ===');
  
  const slides = presentation.getSlides();
  const slidesToDelete = [];
  const referenceSlideIndices = []; // Slides to keep and move to end
  let endSlideIndex = -1;
  
  // First pass: identify slides
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    const shapes = slide.getShapes();
    let hasTemplatePlaceholder = false;
    let isReferenceSlide = false;
    let isEndSlide = false;
    let slideIdentifier = '';
    
    shapes.forEach(shape => {
      if (shape.getText) {
        const text = shape.getText().asString();
        
        // Check for unfilled template placeholders that should be deleted
        if (text.includes('{{Portfolio Initiative}}') || 
            text.includes('{{page # of #}}') ||
            text.includes('{{ # of #}}') ||
            text.includes('Portfolio Distribution')) {
          hasTemplatePlaceholder = true;
        }
        
        // Check for reference slides to KEEP and move to end
        // Format Guidelines divider
        if (text === 'Format Guidelines' || 
            (text.includes('Format Guidelines') && !text.includes('📖') && !text.includes('How to'))) {
          isReferenceSlide = true;
          slideIdentifier = 'Format Guidelines Divider';
        }
        // Format Guidelines content (How to Read)
        if (text.includes('📖 Format Guidelines') || 
            text.includes('How to Read this Deck') ||
            text.includes('Slide Anatomy')) {
          isReferenceSlide = true;
          slideIdentifier = 'Format Guidelines Content';
        }
        // RTE Guidelines (How to USE)
        if (text.includes('📖 RTE Guidelines') || 
            text.includes('RTE Guidelines') ||
            text.includes('How to USE this Deck') ||
            text.includes('COLOR REFERENCE') ||
            text.includes('FONT SPECIFICATION')) {
          isReferenceSlide = true;
          slideIdentifier = 'RTE Guidelines';
        }
        // Changes Report Legend
        if (text.includes('Changes Report Legend') || 
            text.includes('AUDIT REFERENCE - REMOVE BEFORE DISTRIBUTION') ||
            text.includes('WHAT APPEARS ON CHANGES REPORT')) {
          isReferenceSlide = true;
          slideIdentifier = 'Changes Report Legend';
        }
        
        // Check for End of Governance slide
        if (text.includes('End of Governance') || 
            text.includes('End of Report')) {
          isEndSlide = true;
        }
      }
    });
    
    // Mark for deletion if it's a template placeholder but NOT a reference slide
    if (hasTemplatePlaceholder && !isReferenceSlide) {
      slidesToDelete.push(i);
      console.log(`  📌 Marking slide ${i} for deletion (unfilled template)`);
    }
    
    // Track reference slides
    if (isReferenceSlide) {
      referenceSlideIndices.push({ index: i, name: slideIdentifier });
      console.log(`  📚 Found reference slide at index ${i}: ${slideIdentifier}`);
    }
    
    // Track End slide
    if (isEndSlide) {
      endSlideIndex = i;
      console.log(`  🏁 Found End of Governance slide at index ${i}`);
    }
  }
  
  // Delete unfilled template slides in reverse order
  for (let j = slidesToDelete.length - 1; j >= 0; j--) {
    const index = slidesToDelete[j];
    try {
      presentation.getSlides()[index].remove();
      console.log(`  ✓ Deleted slide at index ${index}`);
      
      // Adjust tracked indices
      referenceSlideIndices.forEach(ref => {
        if (ref.index > index) ref.index--;
      });
      if (endSlideIndex > index) endSlideIndex--;
    } catch (e) {
      console.warn(`  ⚠️ Could not delete slide at index ${index}: ${e.message}`);
    }
  }
  
  // Now rearrange slides to final order:
  // [Content...] → [End of Governance] → [Reference Slides in order]
  
  const currentSlides = presentation.getSlides();
  const totalCount = currentSlides.length;
  
  console.log(`\n  📋 Rearranging ${referenceSlideIndices.length} reference slides...`);
  
  // Sort reference slides by their intended order
  const slideOrder = ['Format Guidelines Divider', 'Format Guidelines Content', 'RTE Guidelines', 'Changes Report Legend'];
  referenceSlideIndices.sort((a, b) => {
    return slideOrder.indexOf(a.name) - slideOrder.indexOf(b.name);
  });
  
  // Move End of Governance to just before the reference slides position
  // First, find where it currently is
  let currentEndIndex = -1;
  for (let i = 0; i < presentation.getSlides().length; i++) {
    const shapes = presentation.getSlides()[i].getShapes();
    for (let s = 0; s < shapes.length; s++) {
      if (shapes[s].getText && shapes[s].getText().asString().includes('End of Governance')) {
        currentEndIndex = i;
        break;
      }
    }
    if (currentEndIndex >= 0) break;
  }
  
  // Calculate target position for End slide (after all content, before reference slides)
  // Reference slides should be at the very end
  const numRefSlides = referenceSlideIndices.length;
  const endSlideTargetPos = presentation.getSlides().length - numRefSlides;
  
  if (currentEndIndex >= 0 && currentEndIndex < endSlideTargetPos - 1) {
    try {
      presentation.getSlides()[currentEndIndex].move(endSlideTargetPos);
      console.log(`  ✓ Moved End of Governance to position ${endSlideTargetPos}`);
    } catch (e) {
      console.warn(`  ⚠️ Could not move End slide: ${e.message}`);
    }
  }
  
  // Move each reference slide to the end in order
  // We need to re-find them since indices may have shifted
  const finalSlideCount = presentation.getSlides().length;
  
  slideOrder.forEach((targetName, orderIndex) => {
    // Find this slide
    for (let i = 0; i < presentation.getSlides().length - 1; i++) {
      const shapes = presentation.getSlides()[i].getShapes();
      let isMatch = false;
      
      shapes.forEach(shape => {
        if (!shape.getText) return;
        const text = shape.getText().asString();
        
        if (targetName === 'Format Guidelines Divider' && 
            text === 'Format Guidelines') {
          isMatch = true;
        }
        if (targetName === 'Format Guidelines Content' && 
            (text.includes('📖 Format Guidelines') || text.includes('Slide Anatomy'))) {
          isMatch = true;
        }
        if (targetName === 'RTE Guidelines' && 
            (text.includes('RTE Guidelines') || text.includes('COLOR REFERENCE'))) {
          isMatch = true;
        }
        if (targetName === 'Changes Report Legend' && 
            (text.includes('Changes Report Legend') || text.includes('WHAT APPEARS ON CHANGES REPORT'))) {
          isMatch = true;
        }
      });
      
      if (isMatch) {
        try {
          presentation.getSlides()[i].move(finalSlideCount);
          console.log(`  ✓ Moved "${targetName}" to end`);
        } catch (e) {
          console.warn(`  ⚠️ Could not move "${targetName}": ${e.message}`);
        }
        break;
      }
    }
  });
  
  console.log(`=== CLEANUP: Complete. Final slide count: ${presentation.getSlides().length} ===\n`);
}
function extractIterationNumber(iterString) {
  if (!iterString || typeof iterString !== 'string') return null;
  
  const str = iterString.trim();
  if (!str || str.toLowerCase().includes('unknown')) return null;
  
  // Try to match "Iteration X" pattern (case insensitive)
  const iterMatch = str.match(/iteration\s*(\d+)/i);
  if (iterMatch) {
    return parseInt(iterMatch[1], 10);
  }
  
  // Try to match just a number
  const numMatch = str.match(/(\d+)/);
  if (numMatch) {
    return parseInt(numMatch[1], 10);
  }
  
  return null;
}
function createChangesOnlyPresentation(presentationName, data, piNumber, iterationNumber, showDependencies = true, hideSameTeamDeps = true, highlightSchedulingRisk = true, singlePortfolio = null, valueStreamFilter = null, groupByFixVersion = true) {
  try {
    console.log('\n╔══════════════════════════════════════════════════════════════╗');
    console.log('║     CHANGES-ONLY PRESENTATION - CONSOLIDATED AUDIT SECTION   ║');
    console.log('╠══════════════════════════════════════════════════════════════╣');
    console.log(`║  PI: ${piNumber}  |  Iteration: ${iterationNumber}  |  ${new Date().toLocaleString()}`);
    console.log(`║  Epics/Dependencies: ${data.epics ? data.epics.length : 0}  |  Show Deps: ${showDependencies}`);
    console.log(`║  Portfolio Filter: ${singlePortfolio || 'ALL'}`);
    console.log(`║  Value Stream Filter: ${valueStreamFilter ? valueStreamFilter.join(', ') : 'ALL'}`);
    console.log('╚══════════════════════════════════════════════════════════════╝\n');
    
    const generatedTimestamp = new Date();
    
    // Step 1: Copy the template
    showProgress('Step 1 of 7: Copying template...', '📊 Starting');
    const templateFile = DriveApp.getFileById(TEMPLATE_CONFIG.TEMPLATE_ID);
    const copiedFile = templateFile.makeCopy(presentationName);
    const presentation = SlidesApp.openById(copiedFile.getId());
    console.log(`Created presentation from template: ${presentationName} (ID: ${presentation.getId()})`);
    
    // Step 2: Update title slide with changes-specific title
    showProgress('Step 2 of 7: Updating title slide...', '📊 Preparing');
    updateChangesOnlyTitleSlide(presentation, data.metadata, piNumber, iterationNumber);
    
    // Step 3: Process data (already filtered to changes only)
    showProgress('Step 3 of 7: Processing changed epics...', '📊 Processing Data');
    const portfolioData = processDataForTableSlides(data, showDependencies, false);
    
    // Step 4: Build slides for each portfolio (collect audit data but don't create audit slides yet)
    let sortedPortfolios = getFilteredAndSortedPortfolios(Object.keys(portfolioData));
    
    // ═══════════════════════════════════════════════════════════════════════════
    // APPLY PORTFOLIO FILTER
    // ═══════════════════════════════════════════════════════════════════════════
    if (singlePortfolio) {
      sortedPortfolios = sortedPortfolios.filter(p => p === singlePortfolio);
      if (sortedPortfolios.length === 0) {
        throw new Error(`Portfolio "${singlePortfolio}" not found. Available: ${Object.keys(portfolioData).join(', ')}`);
      }
      console.log(`Single portfolio mode: generating slides only for "${singlePortfolio}"`);
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // APPLY VALUE STREAM FILTER
    // ═══════════════════════════════════════════════════════════════════════════
    if (valueStreamFilter && valueStreamFilter.length > 0) {
      console.log(`Value Stream filter active: ${valueStreamFilter.join(', ')}`);
      
      // Filter epics within each portfolio to only include matching value streams
      sortedPortfolios.forEach(portfolioName => {
        const portfolio = portfolioData[portfolioName];
        if (portfolio && portfolio.programs) {
          Object.keys(portfolio.programs).forEach(programKey => {
            const program = portfolio.programs[programKey];
            if (program && program.initiatives) {
              Object.keys(program.initiatives).forEach(initKey => {
                const initiative = program.initiatives[initKey];
                if (initiative && initiative.epics) {
                  const beforeCount = initiative.epics.length;
                  initiative.epics = initiative.epics.filter(epic => {
                    const epicVS = epic['Value Stream/Org'] || '';
                    return valueStreamFilter.some(vs => 
                      epicVS.toLowerCase().includes(vs.toLowerCase()) ||
                      vs.toLowerCase().includes(epicVS.toLowerCase())
                    );
                  });
                  if (initiative.epics.length < beforeCount) {
                    console.log(`  VS Filter: ${portfolioName}/${programKey}/${initKey}: ${beforeCount} -> ${initiative.epics.length} epics`);
                  }
                }
              });
            }
          });
        }
      });
      
      // Remove portfolios that now have no epics after filtering
      sortedPortfolios = sortedPortfolios.filter(portfolioName => {
        const portfolio = portfolioData[portfolioName];
        if (!portfolio || !portfolio.programs) return false;
        
        const hasEpics = Object.values(portfolio.programs).some(program => 
          program && program.initiatives && Object.values(program.initiatives).some(init => 
            init && init.epics && init.epics.length > 0
          )
        );
        if (!hasEpics) {
          console.log(`  Removing empty portfolio after VS filter: ${portfolioName}`);
        }
        return hasEpics;
      });
    }
    
    const totalPortfolios = sortedPortfolios.length;
    console.log(`Building slides for ${totalPortfolios} portfolios with changes...`);
    
    // Handle empty results case
    if (totalPortfolios === 0) {
      console.log('No portfolios with changes found');
      showProgress('No changes found for this iteration', '⚠️ Warning');
      
      const slides = presentation.getSlides();
      if (slides.length > 1) {
        const noChangesSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        const textBox = noChangesSlide.insertTextBox(
          'No governance changes detected for Iteration ' + iterationNumber + '\n\n' +
          'All epics are unchanged with Green or no RAG status.',
          50, 200, 620, 100
        );
        textBox.getText().getTextStyle()
          .setFontFamily('Lato')
          .setFontSize(18)
          .setForegroundColor('#5f6368');
        textBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      }
      
      cleanupTemplateSlides(presentation, 0);
      
      const url = presentation.getUrl();
      SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
      Utilities.sleep(300);
      SpreadsheetApp.getUi().alert('Changes Report', 
        `No governance changes detected for PI ${piNumber} Iteration ${iterationNumber}.\n\nPresentation created: ${url}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
      
      return { success: true, url: url, slideCount: presentation.getSlides().length, changesCount: 0 };
    }
    
    // Get template slides (before we start duplicating)
    const dividerTemplate = presentation.getSlides()[TEMPLATE_CONFIG.PORTFOLIO_DIVIDER_INDEX];
    const tableTemplate = presentation.getSlides()[TEMPLATE_CONFIG.TABLE_SLIDE_INDEX];
    
    // Track insertPosition for slide placement
    let insertPosition = TEMPLATE_CONFIG.END_SLIDE_INDEX;
    
    // Collect audit data for all portfolios (will create consolidated audit section at the very end)
    const allAuditData = [];
    
    sortedPortfolios.forEach((portfolioName, index) => {
      const progress = Math.round(((index + 1) / totalPortfolios) * 100);
      showProgress(
        `Step 4 of 7: Building portfolio ${index + 1} of ${totalPortfolios} (${progress}%)\n${portfolioName}`,
        '📊 Building Slides'
      );
      console.log(`\n[${index + 1}/${totalPortfolios}] Processing Portfolio: ${portfolioName}`);
      console.log(`  📍 Current insertPosition: ${insertPosition}`);
      
      // ═══════════════════════════════════════════════════════════════════════
      // STEP A: Collect and analyze epics for audit data
      // ═══════════════════════════════════════════════════════════════════════
      const portfolioEpics = collectPortfolioEpics(portfolioData[portfolioName]);
      const auditData = buildPortfolioAuditDataEnhanced(portfolioEpics, portfolioName);
      
      console.log(`  📋 Audit: NEW=${auditData.new.length}, CHG=${auditData.chg.length}, DONE=${auditData.done.length}, DEF=${auditData.def.length}, AT-RISK=${auditData.atRisk.length}`);
      
      // Store for consolidated audit section at end
      allAuditData.push(auditData);
      
      // ═══════════════════════════════════════════════════════════════════════
      // STEP B: Create DIVIDER + TABLE slides
      // INFOSEC SPECIAL HANDLING: Group by subsection with separate slide sections
      // ═══════════════════════════════════════════════════════════════════════
      let slidesCreatedByPortfolio = 0;
      
      if (isInfosecPortfolio(portfolioName)) {
        console.log(`  🔐 INFOSEC portfolio detected - using subsection grouping`);
        
        const subsectionGroups = groupInfosecEpicsBySubsection(portfolioData[portfolioName]);
        const subsectionNames = Object.keys(subsectionGroups);
        
        if (subsectionNames.length === 0) {
          console.log(`  ⚠️ No INFOSEC subsections with data - skipping portfolio`);
        } else {
          console.log(`  🔐 Found ${subsectionNames.length} subsections with data`);
          
          // Step 1: Create ONE divider slide for "INFOSEC" (not per-subsection)
          const dividerSlide = presentation.insertSlide(insertPosition + slidesCreatedByPortfolio, dividerTemplate);
          updatePortfolioDividerSlide(dividerSlide, portfolioName, portfolioData[portfolioName].totalFeaturePoints);
          slidesCreatedByPortfolio++;
          console.log(`  🔐 Created single INFOSEC divider slide`);
          
          // Step 2: Create table slides for each subsection (no dividers)
          subsectionNames.forEach((subsectionName, subIndex) => {
            const subsectionData = subsectionGroups[subsectionName];
            const subsectionTitle = `INFOSEC - ${subsectionName}`;
            
            console.log(`  🔐 [${subIndex + 1}/${subsectionNames.length}] Creating table slides for: ${subsectionTitle}`);
            
            // Create ONLY table slides (no divider) using the subsection title in header
            const tableSlides = createPortfolioTableSlides(
              presentation, 
              tableTemplate, 
              subsectionTitle,  // Shows "INFOSEC - Subsection" in table header
              subsectionData,   // Use subsection-filtered data
              data.dependencies, 
              showDependencies,
              insertPosition + slidesCreatedByPortfolio,
              generatedTimestamp,
              'skip',  // noInitiativeMode
              hideSameTeamDeps,
              groupByFixVersion
            );
            
            slidesCreatedByPortfolio += tableSlides;
            console.log(`    📄 Created ${tableSlides} table slides for ${subsectionTitle}`);
          });
        }
      } else {
        // ─────────────────────────────────────────────────────────────────────
        // NON-INFOSEC: Standard processing (unchanged from original)
        // ─────────────────────────────────────────────────────────────────────
        slidesCreatedByPortfolio = createPortfolioSlidesFromTemplate(
          presentation, 
          dividerTemplate, 
          tableTemplate, 
          portfolioName, 
          portfolioData[portfolioName], 
          data.dependencies, 
          showDependencies,
          insertPosition,
          generatedTimestamp,
          'skip',  // noInitiativeMode - skip for changes report (no "No Initiative" sections)
          hideSameTeamDeps
        );
      }
      
      console.log(`  📄 Total slides created for portfolio: ${slidesCreatedByPortfolio}`);
      
      // ═══════════════════════════════════════════════════════════════════════
      // STEP C: UPDATE insertPosition for next portfolio
      // ═══════════════════════════════════════════════════════════════════════
      insertPosition += slidesCreatedByPortfolio;
      console.log(`  📍 Updated insertPosition: ${insertPosition}`);
      
      // NOTE: Audit slides are NOT created here - they will be in consolidated section at the end
      console.log(`  ⏭️ Audit data collected (will be in consolidated section at end)`);
    });
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Step 5: Clean up template slides and arrange End/Format slides
    // This MUST run BEFORE audit section so Format Guidelines are in place
    // ═══════════════════════════════════════════════════════════════════════════
    showProgress('Step 5 of 7: Organizing slide order...', '📊 Arranging');
    cleanupTemplateSlides(presentation, totalPortfolios);
    console.log('✓ Cleanup complete - End of Governance and Format Guidelines positioned');
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Step 6: CREATE CONSOLIDATED AUDIT SECTION at the ABSOLUTE END
    // This runs AFTER cleanup so audit slides are truly at the very end
    // ═══════════════════════════════════════════════════════════════════════════
    showProgress('Step 6 of 7: Creating audit section...', '📊 Building Audit Trail');
    
    let auditSlidesCreated = 0;
    if (allAuditData.length > 0 && allAuditData.some(a => a.hasChanges)) {
      auditSlidesCreated = createConsolidatedAuditSection(presentation, allAuditData, {
        piNumber: piNumber,
        iterationNumber: iterationNumber
      });
      console.log(`✓ Created ${auditSlidesCreated} audit slides at absolute end of presentation`);
    } else {
      console.log('⏭️ No audit section needed - no portfolios with changes');
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Step 7: Final verification and completion
    // ═══════════════════════════════════════════════════════════════════════════
    showProgress('Step 7 of 7: Completing presentation...', '📊 Finishing Up');
    
    const url = presentation.getUrl();
    const slideCount = presentation.getSlides().length;
    const epicCount = data.epics ? data.epics.filter(e => e['Issue Type'] === 'Epic').length : 0;
    
    console.log(`\n${'═'.repeat(60)}`);
    console.log(`Changes presentation complete: ${url}`);
    console.log(`Total slides: ${slideCount} (including ${auditSlidesCreated} audit slides at end)`);
    console.log(`Portfolios with changes: ${allAuditData.filter(a => a.hasChanges).length}`);
    console.log(`${'═'.repeat(60)}\n`);
    
    // Show completion
    showProgress(
      `✅ Complete! Created ${slideCount} slides showing ${epicCount} changed epics.`,
      '✅ Done'
    );
    
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    Utilities.sleep(300);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Created ${slideCount} slides with ${epicCount} changed/at-risk epics`, 
      '✅ Changes Report Complete', 
      5
    );
    
    return { 
      success: true, 
      url: url, 
      presentationId: presentation.getId(),
      slideCount: slideCount, 
      changesCount: epicCount 
    };
    
  } catch (error) {
    console.error('Changes-only presentation error:', error);
    showProgress(`❌ Error: ${error.message}`, '⚠️ Failed', 10);
    SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
    throw error;
  }
}

/**
 * Create table slides for a portfolio (extracted from createPortfolioSlidesFromTemplate)
 * This is called AFTER the audit slide is created
 */
function createPortfolioTableSlides(presentation, tableTemplate, portfolioName, portfolioData, allDependencies, showDependencies, insertPosition, generatedTimestamp, noInitiativeMode, hideSameTeamDeps, groupByFixVersion = true) {
  if (!portfolioData || !portfolioData.programs || Object.keys(portfolioData.programs).length === 0) {
    console.log(`Portfolio "${portfolioName}" has no data, skipping tables`);
    return 0;
  }
  
  let slidesCreated = 0;
  
  // Format timestamp for display
  const timestampText = formatTimestamp(generatedTimestamp);
  
  // Build program data for table slides
  const programRows = [];
  const sortedPrograms = Object.keys(portfolioData.programs).sort();
  
  sortedPrograms.forEach(programName => {
    const programData = portfolioData.programs[programName];
    const initiatives = [];
    
    Object.keys(programData.initiatives).sort().forEach(initTitle => {
      initiatives.push({
        title: initTitle,
        ...programData.initiatives[initTitle]
      });
    });
    
    if (initiatives.length > 0) {
      programRows.push({
        programName: programName,
        initiatives: initiatives
      });
    }
  });
  
  if (programRows.length === 0) {
    return 0;
  }
  
  // =============================================================================
  // LINE-BASED PAGINATION: Max 20 lines per slide
  // =============================================================================
  const MAX_LINES_PER_SLIDE = 20;
  
  // Build flat list of content lines with context
  const contentLines = buildContentLinesFromPrograms(programRows, showDependencies, noInitiativeMode);
  console.log(`  Total content lines for "${portfolioName}": ${contentLines.length}`);
  
  const tableSlides = [];
  let currentSlide = null;
  let currentTable = null;
  let currentLineCount = 0;
  let programColorIndex = 0;
  let lastProgramName = null;
  
  // Track context for continuation headers
  let currentContext = {
    program: null,
    programBgColor: null,
    initiative: null,
    valueStream: null
  };
  
  // Helper to create a new table slide
  function createNewTableSlide() {
    const newSlide = presentation.insertSlide(insertPosition + slidesCreated, tableTemplate);
    updateTableSlideHeader(newSlide, portfolioName);
    
    // Add timestamp to slide
    addTimestampToSlide(newSlide, timestampText);
    
    const newTable = getOrCreateTable(newSlide);
    tableSlides.push(newSlide);
    slidesCreated++;
    currentLineCount = 0;
    console.log(`    -> Created table slide #${tableSlides.length} for "${portfolioName}"`);
    return { slide: newSlide, table: newTable };
  }
  
  // Process content lines
  let i = 0;
  let needsContinuation = false;
  const ROLL_THRESHOLD = 22;
  
  while (i < contentLines.length) {
    const item = contentLines[i];
    
    // Track program changes for alternating colors
    if (item.program !== lastProgramName) {
      if (lastProgramName !== null) {
        programColorIndex++;
      }
      lastProgramName = item.program;
    }
    const bgColor = PROGRAM_BG_COLORS[programColorIndex % 2];
    
    // Ensure we have a slide
    if (!currentSlide || !currentTable) {
      const result = createNewTableSlide();
      currentSlide = result.slide;
      currentTable = result.table;
    }
    
    // Check if we need a new slide
    const isNewMajorElement = item.type === 'initiative' || item.type === 'valueStream' || item.type === 'epic';
    const isNearEndOfSlide = currentLineCount >= ROLL_THRESHOLD;
    const shouldRollForNewContent = isNearEndOfSlide && isNewMajorElement && 
      (item.program !== currentContext.program || item.type === 'initiative' || item.type === 'valueStream');
    
    if (currentLineCount + item.lines > MAX_LINES_PER_SLIDE || shouldRollForNewContent) {
      const result = createNewTableSlide();
      currentSlide = result.slide;
      currentTable = result.table;
      needsContinuation = true;
    }
    
    // Collect consecutive items from same program that fit on this slide
    const programBatch = [];
    const batchStartProgram = item.program;
    let batchLineCount = 0;
    
    // Calculate extra lines needed for continuation headers
    let continuationLineCount = 0;
    if (needsContinuation && currentContext.program === batchStartProgram) {
      continuationLineCount = 1;
      if (currentContext.initiative) {
        continuationLineCount += 1;
        if (currentContext.valueStream) {
          continuationLineCount += 1;
        }
      }
    }
    
    while (i < contentLines.length && 
           contentLines[i].program === batchStartProgram &&
           currentLineCount + continuationLineCount + batchLineCount + contentLines[i].lines <= MAX_LINES_PER_SLIDE) {
      
      const nextItem = contentLines[i];
      const nextIsNewMajor = nextItem.type === 'initiative' || nextItem.type === 'valueStream';
      const wouldBeNearEnd = currentLineCount + continuationLineCount + batchLineCount >= ROLL_THRESHOLD;
      
      if (wouldBeNearEnd && nextIsNewMajor && batchLineCount > 0) {
        break;
      }
      
      programBatch.push(contentLines[i]);
      batchLineCount += contentLines[i].lines;
      i++;
    }
    
    // Build and add this batch to the table
    if (programBatch.length > 0) {
      const isSameProgram = currentContext.program === batchStartProgram;
      const displayName = isSameProgram ? `${batchStartProgram} (continued)` : batchStartProgram;
      
      const continuationContext = (needsContinuation && isSameProgram) ? {
        initiative: currentContext.initiative,
        valueStream: currentContext.valueStream
      } : null;
      
      const batchInitiatives = buildInitiativesFromContentBatch(programBatch, showDependencies, continuationContext);
      
      if (batchInitiatives.length > 0) {
        const totalLines = batchLineCount + (continuationContext ? continuationLineCount : 0);
        addProgramRow(currentTable, displayName, batchInitiatives, bgColor, showDependencies, noInitiativeMode, hideSameTeamDeps, groupByFixVersion);
        currentLineCount += totalLines;
      }

      // Update context for continuation tracking
      const lastItem = programBatch[programBatch.length - 1];
      currentContext.program = lastItem.program;
      currentContext.programBgColor = bgColor;
      currentContext.initiative = lastItem.initiativeObj || null;
      currentContext.valueStream = lastItem.valueStream || null;
      
      needsContinuation = false;
    }
  }
  
  // Add page numbers to table slides
  const totalTableSlides = tableSlides.length;
  tableSlides.forEach((slide, index) => {
    addPortfolioPageNumber(slide, index + 1, totalTableSlides);
  });
  
  return slidesCreated;
}
function updateChangesOnlyTitleSlide(presentation, metadata, piNumber, iterationNumber) {
  try {
    const titleSlide = presentation.getSlides()[0];
    const shapes = titleSlide.getShapes();
    
    // Find and update text boxes on the title slide
    shapes.forEach(shape => {
      if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
        const text = shape.getText().asString().toLowerCase();
        
        // Update main title placeholder
        if (text.includes('{{title}}') || text.includes('governance') || text.includes('program increment')) {
          shape.getText().setText(`PI ${piNumber} - Iteration ${iterationNumber}\nGovernance Changes`);
          const textStyle = shape.getText().getTextStyle();
          textStyle.setFontFamily('Lato');
          textStyle.setBold(true);
          textStyle.setFontSize(28);
        }
        
        // Update subtitle/date placeholder  
        if (text.includes('{{date}}') || text.includes('{{subtitle}}') || text.includes('date')) {
          const dateStr = new Date().toLocaleDateString('en-US', { 
            weekday: 'long', 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric' 
          });
          shape.getText().setText(`Changes since Iteration ${iterationNumber - 1}\n${dateStr}`);
          const textStyle = shape.getText().getTextStyle();
          textStyle.setFontFamily('Lato');
          textStyle.setFontSize(14);
        }
      }
    });
    
    console.log('✓ Title slide updated for changes-only report');
  } catch (error) {
    console.warn(`Could not update title slide: ${error.message}`);
    // Non-fatal - continue with presentation generation
  }
}
// =================================================================================
// AUDIT SLIDE FUNCTIONS - FIXED VERSION (no BLANK layout required)
// =================================================================================
const AUDIT_SLIDE_CONFIG = {
  // Layout
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 540,
  HEADER_HEIGHT: 28,
  TITLE_HEIGHT: 40,
  TITLE_TOP: 32,
  CONTENT_TOP: 75,
  CONTENT_BOTTOM: 520,
  MARGIN: 20,
  COLUMN_GAP: 16,
  
  // Two columns
  get COLUMN_WIDTH() { return (this.SLIDE_WIDTH - (2 * this.MARGIN) - this.COLUMN_GAP) / 2; },
  get LEFT_COLUMN_X() { return this.MARGIN; },
  get RIGHT_COLUMN_X() { return this.MARGIN + this.COLUMN_WIDTH + this.COLUMN_GAP; },
  get CONTENT_HEIGHT() { return this.CONTENT_BOTTOM - this.CONTENT_TOP; },
  
  // Rows per column (each item = ~4 lines at font size 8)
  LINES_PER_ITEM: 4,
  FONT_SIZE: 8,
  LINE_HEIGHT: 10,  // pixels per line
  get MAX_LINES_PER_COLUMN() { return Math.floor(this.CONTENT_HEIGHT / this.LINE_HEIGHT); },
  get MAX_ITEMS_PER_COLUMN() { return Math.floor(this.MAX_LINES_PER_COLUMN / this.LINES_PER_ITEM); },
  get MAX_ITEMS_PER_SLIDE() { return this.MAX_ITEMS_PER_COLUMN * 2; }  // Two columns
};

function collectPortfolioEpics(portfolioData) {
  const epics = [];
  
  if (!portfolioData || !portfolioData.programs) {
    return epics;
  }
  
  Object.keys(portfolioData.programs).forEach(programName => {
    const program = portfolioData.programs[programName];
    
    if (program.initiatives) {
      Object.keys(program.initiatives).forEach(initName => {
        const initiative = program.initiatives[initName];
        
        // Use initiative.epics (the correct property name)
        if (initiative.epics && Array.isArray(initiative.epics)) {
          initiative.epics.forEach(epic => {
            if (epic['Issue Type'] === 'Epic') {
              epics.push(epic);
            }
          });
        }
      });
    }
  });
  
  return epics;
}

function buildPortfolioAuditDataEnhanced(epics, portfolioName) {
  const auditData = {
    portfolioName: portfolioName,
    new: [],
    chg: [],
    done: [],
    pending: [],
    def: [],
    canceled: [],
    overdue: [],
    atRisk: [],
    info: [],
    hasChanges: false
  };
  
  epics.forEach(epic => {
    // ═══════════════════════════════════════════════════════════════════════════
    // EXCLUSION 1: Skip epics with Resolution = "Duplicate"
    // ═══════════════════════════════════════════════════════════════════════════
    const resolution = (epic['Resolution'] || '').toString().trim().toLowerCase();
    if (resolution === 'duplicate') {
      console.log(`  🚫 Audit skip: ${epic['Key']} - Resolution = "Duplicate"`);
      return;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // EXCLUSION 2: Skip epics with empty summaries
    // ═══════════════════════════════════════════════════════════════════════════
    const rawSummary = epic['Summary'];
    const summaryIsEmpty = !rawSummary || 
                           rawSummary.toString().trim() === '' ||
                           ['n/a', 'na', 'none', 'null', 'undefined', '-', '--'].includes(
                             rawSummary.toString().trim().toLowerCase()
                           );
    if (summaryIsEmpty) {
      console.log(`  ⏭️ Audit skip: ${epic['Key']} - empty/excluded summary`);
      return;
    }
    
    const key = epic['Key'] || 'Unknown';
    const summary = (epic['Summary'] || 'No summary').substring(0, 60);
    const valueStream = epic['Value Stream/Org'] || 'Unknown';
    const rag = epic['RAG'] || '';
    const ragNote = epic['RAG Note'] || '';
    const status = epic['Status'] || '';
    const piCommitment = epic['PI Commitment'] || '';
    const changes = epic.changes || {};
    
    // Get badge from changelog (source of truth)
    const badge = (epic.badge || '').toUpperCase();
    
    const item = {
      key: key,
      summary: summary + (summary.length >= 60 ? '...' : ''),
      valueStream: valueStream,
      ragLevel: rag.toUpperCase(),
      ragNote: ragNote
    };
    
    // Track which category this epic goes into (for deduplication)
    let categorized = false;
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CANCELED: Epic canceled (check first - takes precedence)
    // ═══════════════════════════════════════════════════════════════════════════
    if (epic.isCanceled || badge === 'CANCELED') {
      if (epic.canceledThisIteration) {
        auditData.canceled.push({
          ...item,
          details: ['Canceled this iteration'],
          businessRule: `PI Commitment changed to Canceled during current iteration`
        });
      } else {
        auditData.canceled.push({
          ...item,
          details: ['Already canceled (continuity)'],
          businessRule: `PI Commitment = Canceled (canceled in a previous iteration, included for continuity)`
        });
      }
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // DEF: Epic deferred (check BEFORE DONE - higher priority)
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isDeferred || badge === 'DEF')) {
      if (epic.deferredThisIteration) {
        const prevCommit = changes.previousPiCommitment || '(blank)';
        auditData.def.push({
          ...item, 
          details: [`PI Commitment: ${prevCommit} → ${piCommitment}`],
          businessRule: `PI Commitment changed to "${piCommitment}" (previously "${prevCommit}")`
        });
      } else if (changes.piCommitmentFromCommitted) {
        const prevCommit = changes.previousPiCommitment || 'Committed';
        const currCommit = changes.currentPiCommitment || '(blank)';
        auditData.def.push({
          ...item,
          details: [`PI Commitment: ${prevCommit} → ${currCommit}`],
          businessRule: `PI Commitment changed from "${prevCommit}" to "${currCommit}" - no longer committed`
        });
      } else {
        auditData.def.push({
          ...item, 
          details: ['Already deferred (continuity)'],
          businessRule: `PI Commitment = "${piCommitment}" (deferred in a previous iteration, included for continuity)`
        });
      }
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // NEW: Added this iteration OR PI Commitment changed TO Committed
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isNew || badge === 'NEW')) {
      auditData.new.push({
        ...item, 
        details: ['Added this iteration'],
        businessRule: 'Epic created/added during current iteration (Added in Iteration = current)'
      });
      categorized = true;
    } else if (!categorized && changes.piCommitmentToCommitted) {
      const prevCommit = changes.previousPiCommitment || '(blank)';
      const currCommit = changes.currentPiCommitment || 'Committed';
      auditData.new.push({
        ...item,
        details: [`PI Commitment: ${prevCommit} → ${currCommit}`],
        businessRule: `PI Commitment changed to "${currCommit}" (previously "${prevCommit}") - newly committed this iteration`
      });
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // DONE: Epic closed/completed
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isDone || badge === 'DONE')) {
      if (epic.closedThisIteration) {
        auditData.done.push({
          ...item, 
          details: ['Closed this iteration'],
          businessRule: `Status transitioned to "${status}" during current iteration`
        });
      } else {
        auditData.done.push({
          ...item, 
          details: ['Already closed (continuity)'],
          businessRule: `Status = "${status}" (closed in a previous iteration, included for continuity)`
        });
      }
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // PENDING: Epic in Pending Acceptance (Pending Closure)
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isPendingClosure || badge === 'PENDING')) {
      if (epic.pendingClosureThisIteration) {
        auditData.pending.push({
          ...item, 
          details: ['Moved to Pending Acceptance this iteration'],
          businessRule: `Status transitioned to "Pending Acceptance" during current iteration`
        });
      } else {
        auditData.pending.push({
          ...item, 
          details: ['Already in Pending Acceptance (continuity)'],
          businessRule: `Status = "Pending Acceptance" (entered in a previous iteration, included for continuity)`
        });
      }
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // OVERDUE: Due this iteration but not complete
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isOverdue || badge === 'OVERDUE')) {
      const endIteration = epic['End Iteration Name'] || epic['PI Target Iteration'] || '';
      auditData.overdue.push({
        ...item, 
        details: [`Due: ${endIteration}`],
        businessRule: `End Iteration = current iteration but Status ≠ Closed/Done`
      });
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // CHG: Has qualifying changes
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isModified || badge === 'CHG')) {
      const reasons = epic.qualifyingReasons || [];
      auditData.chg.push({
        ...item,
        details: reasons.length > 0 ? reasons : ['Has qualifying changes'],
        businessRule: reasons.length > 0 ? reasons.join('; ') : 'Field changes detected during iteration'
      });
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // AT-RISK: Has Amber/Red RAG with no other changes
    // ═══════════════════════════════════════════════════════════════════════════
    if (!categorized && (epic.isAtRisk || badge === 'ATRISK' || badge === 'AT-RISK')) {
      auditData.atRisk.push({
        ...item,
        details: [`RAG: ${rag}`],
        businessRule: `RAG status is "${rag}" - included for visibility`,
        ragUnchanged: true
      });
      categorized = true;
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // INFO: Has RAG Note without other categorization
    // ═══════════════════════════════════════════════════════════════════════════
    const noteText = (ragNote || '').toLowerCase();
    const hasNote = noteText && noteText.trim() !== '';
    const isJustStatusIndicator = noteText === 'green' || 
                                   noteText.startsWith('green -') ||
                                   noteText.startsWith('green:') ||
                                   noteText === 'on track' ||
                                   noteText === 'no issues';
    
    if (hasNote && !isJustStatusIndicator && !categorized) {
      let businessRule = '';
      
      if (piCommitment === 'Deferred' || piCommitment === 'Canceled') {
        businessRule = `Deferred/Canceled item with Note - displayed with "Note - " prefix for context`;
      } else {
        businessRule = `RAG Note exists ("${ragNote.substring(0, 40)}...") - displayed on slide for visibility`;
      }
      
      auditData.info.push({
        ...item,
        details: ['Has informational note displayed on slide'],
        businessRule: businessRule
      });
    }
  });
  
  // Check if there are any changes
  auditData.hasChanges = (
    auditData.new.length > 0 ||
    auditData.chg.length > 0 ||
    auditData.done.length > 0 ||
    auditData.pending.length > 0 ||
    auditData.def.length > 0 ||
    auditData.canceled.length > 0 ||
    auditData.overdue.length > 0 ||
    auditData.atRisk.length > 0 ||
    auditData.info.length > 0
  );
  
  // Log summary
  console.log(`  📋 Audit data for ${portfolioName}:`);
  console.log(`     NEW: ${auditData.new.length}, OVERDUE: ${auditData.overdue.length}, CHG: ${auditData.chg.length}, DONE: ${auditData.done.length}`);
  console.log(`     PENDING: ${auditData.pending.length}, DEF: ${auditData.def.length}, CANCELED: ${auditData.canceled.length}, AT-RISK: ${auditData.atRisk.length}, INFO: ${auditData.info.length}`);
  
  return auditData;
}
function addAuditSlidesToPresentationWithPagination(presentation, portfolioData, metadata) {
  console.log('\n' + '═'.repeat(60));
  console.log('ADDING AUDIT SLIDES (Two-Column Layout with Pagination)');
  console.log('═'.repeat(60));
  
  // Collect all portfolio audit data
  const allAuditData = [];
  
  Object.keys(portfolioData).forEach(portfolioName => {
    const epics = collectPortfolioEpics(portfolioData[portfolioName]);
    
    if (epics.length > 0) {
      // Use enhanced audit data builder
      const auditData = buildPortfolioAuditDataEnhanced(epics, portfolioName);
      
      if (auditData.hasChanges) {
        allAuditData.push(auditData);
      }
    }
  });
  
  if (allAuditData.length === 0) {
    console.log('No portfolios have changes - skipping audit section');
    return 0;
  }
  
  const portfoliosWithChanges = allAuditData.filter(d => d.hasChanges);
  console.log(`Found ${portfoliosWithChanges.length} portfolios with changes`);
  
  let slidesCreated = 0;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 1: Audit Section Divider
  // ═══════════════════════════════════════════════════════════════════════════
  const dividerSlide = createBlankSlideAppend(presentation);
  dividerSlide.getBackground().setSolidFill(AUDIT_COLORS.BACKGROUND);
  
  // Warning banner
  const dividerBanner = dividerSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 40);
  dividerBanner.getFill().setSolidFill(AUDIT_COLORS.HEADER_BG);
  dividerBanner.getBorder().setTransparent();
  
  const dividerBannerText = dividerBanner.getText();
  dividerBannerText.setText('⚠️  AUDIT SECTION - DELETE BEFORE DISTRIBUTION  ⚠️');
  dividerBannerText.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(16)
    .setBold(true)
    .setForegroundColor('#FFFFFF');
  dividerBannerText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Title
  const dividerTitle = dividerSlide.insertTextBox(
    `PI ${metadata.piNumber} - Iteration ${metadata.iterationNumber}\nBusiness Rules Audit`,
    60, 100, 600, 80
  );
  dividerTitle.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(32)
    .setBold(true)
    .setForegroundColor('#502d7f');
  dividerTitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Summary counts including OVERDUE
  const totals = { new: 0, chg: 0, done: 0, def: 0, canceled: 0, overdue: 0, atRisk: 0, info: 0 };
  allAuditData.forEach(pd => {
    totals.new += pd.new.length;
    totals.chg += pd.chg.length;
    totals.done += pd.done.length;
    totals.def += pd.def.length;
    totals.canceled += (pd.canceled || []).length;
    totals.overdue += (pd.overdue || []).length;
    totals.atRisk += pd.atRisk.length;
    totals.info += (pd.info || []).length;
  });
  
  const subtitleText = `${portfoliosWithChanges.length} Portfolio${portfoliosWithChanges.length !== 1 ? 's' : ''} with Changes\n` +
    `NEW: ${totals.new}  |  OVERDUE: ${totals.overdue}  |  CHG: ${totals.chg}  |  DONE: ${totals.done}  |  DEF: ${totals.def}  |  CANCELED: ${totals.canceled}  |  AT-RISK: ${totals.atRisk}`;
  
  const subtitle = dividerSlide.insertTextBox(subtitleText, 60, 260, 600, 120);
  subtitle.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(16)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT || '#4A4A4A');
  subtitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Instructions
  const instructions = dividerSlide.insertTextBox(
    'These slides show business rules that triggered each item to appear in the report.\n' +
    'Two-column layout • Overflow items continue to additional slides.\n' +
    'Delete this entire section before sharing with stakeholders.',
    60, 400, 600, 80
  );
  instructions.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setItalic(true)
    .setForegroundColor(AUDIT_COLORS.WARNING_TEXT || '#8B5A2B');
  instructions.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  slidesCreated++;
  console.log('    ✅ Created audit section divider slide');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 2: Master Summary (all portfolios combined)
  // ═══════════════════════════════════════════════════════════════════════════
  createMasterAuditSlideWithColumns(presentation, allAuditData, metadata);
  slidesCreated++;
  console.log('    ✅ Created master audit summary slide');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDES 3+: Individual Portfolio Audit Slides (with pagination)
  // ═══════════════════════════════════════════════════════════════════════════
  portfoliosWithChanges.forEach((auditData, index) => {
    const portfolioSlides = createPortfolioAuditSlidesWithPagination(presentation, auditData);
    slidesCreated += portfolioSlides;
    console.log(`    ✅ Created ${portfolioSlides} audit slide(s) for ${auditData.portfolioName}`);
  });
  
  console.log(`\n✅ Total audit slides created: ${slidesCreated}`);
  return slidesCreated;
}

// =================================================================================
// HELPER: Create Master Audit Slide with Two Columns
// =================================================================================

function createMasterAuditSlide(presentation, allAuditData, metadata) {
  // Append to end of presentation
  const slide = createBlankSlideAppend(presentation);
  
  // Set soft background
  slide.getBackground().setSolidFill(AUDIT_COLORS.BACKGROUND);
  
  // Warning banner
  const warningBanner = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 28);
  warningBanner.getFill().setSolidFill(AUDIT_COLORS.HEADER_BG);
  warningBanner.getBorder().setTransparent();
  
  const warningText = warningBanner.getText();
  warningText.setText('⚠   AUDIT SLIDE - REMOVE BEFORE DISTRIBUTION   ⚠');
  warningText.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.TITLE_TEXT);
  warningText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Title
  const titleBox = slide.insertTextBox(
    `PI ${metadata.piNumber} - Iteration ${metadata.iterationNumber} | ALL CHANGES SUMMARY`,
    20, 40, 680, 35
  );
  titleBox.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor('#502d7f');
  
  // Aggregate all changes
  const totals = { new: 0, chg: 0, done: 0, def: 0, canceled: 0, overdue: 0, atRisk: 0 };
  allAuditData.forEach(pd => {
    totals.new += pd.new.length;
    totals.chg += pd.chg.length;
    totals.done += pd.done.length;
    totals.def += pd.def.length;
    totals.canceled += (pd.canceled || []).length;
    totals.overdue += (pd.overdue || []).length;
    totals.atRisk += pd.atRisk.length;
  });
  
  // Summary line
  const summaryBox = slide.insertTextBox(
    `NEW: ${totals.new}  |  OVERDUE: ${totals.overdue}  |  CHG: ${totals.chg}  |  DONE: ${totals.done}  |  DEF: ${totals.def}  |  CANCELED: ${totals.canceled}  |  AT-RISK: ${totals.atRisk}`,
    20, 80, 680, 25
  );
  summaryBox.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT);
  
  // Build combined content by category
  let content = '';
  
  ['done', 'overdue', 'chg', 'def', 'canceled', 'atRisk', 'new'].forEach(category => {
    const categoryTitle = category.toUpperCase().replace('ATRISK', 'AT-RISK');
    const items = [];
    
    allAuditData.forEach(pd => {
      const catItems = pd[category] || [];
      catItems.forEach(item => {
        items.push({
          ...item,
          portfolio: pd.portfolioName.substring(0, 12)
        });
      });
    });
    
    if (items.length > 0) {
      content += `${categoryTitle} (${items.length})\n`;
      items.forEach(item => {
        const vsDisplay = item.valueStream ? item.valueStream.substring(0, 25) : item.portfolio;
        content += `  • ${item.key} [${vsDisplay}]\n`;
        content += `    ${item.summary.substring(0, 45)}...\n`;
        // Add business rule on master slide with prominent WHY format
        if (item.businessRule) {
          // Truncate long rules for master slide
          const ruleText = item.businessRule.length > 70 
            ? item.businessRule.substring(0, 67) + '...' 
            : item.businessRule;
          content += `    ➤ WHY: ${ruleText}\n`;
        }
      });
      content += '\n';
    }
  });
  
  if (content.trim() === '') {
    content = 'No changes detected.';
  }
  
  // Add content
  const contentBox = slide.insertTextBox(content.trim(), 20, 110, 680, 400);
  const textRange = contentBox.getText();
  
  textRange.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(8)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT);
  
  // Style section headers with softer colors
  const fullText = textRange.asString();
  const sections = [
    { pattern: /^NEW \(\d+\)/m, color: '#00ACC1' },
    { pattern: /^OVERDUE \(\d+\)/m, color: '#E65100' },
    { pattern: /^CHG \(\d+\)/m, color: '#FB8C00' },
    { pattern: /^DONE \(\d+\)/m, color: '#43A047' },
    { pattern: /^DEF \(\d+\)/m, color: '#E53935' },
    { pattern: /^CANCELED \(\d+\)/m, color: '#757575' },
    { pattern: /^AT-RISK \(\d+\)/m, color: '#FB8C00' }
  ];
  
  sections.forEach(section => {
    const match = fullText.match(section.pattern);
    if (match) {
      const startIdx = fullText.indexOf(match[0]);
      const endIdx = startIdx + match[0].length;
      try {
        textRange.getRange(startIdx, endIdx).getTextStyle()
          .setBold(true)
          .setFontSize(10)
          .setForegroundColor(section.color);
      } catch (e) {}
    }
  });
  
  console.log('  Created master audit slide');
}

function buildMasterColumnContent(items) {
  let text = '';
  const headerPositions = [];
  
  items.forEach(item => {
    if (item.type === 'header') {
      const startPos = text.length;
      text += `${item.text}\n`;
      headerPositions.push({
        start: startPos,
        end: text.length - 1,
        color: item.color
      });
    } else {
      // Compact format for master slide - show valueStream for context
      const vsDisplay = item.valueStream ? item.valueStream.substring(0, 20) : item.portfolio;
      text += `• ${item.key} [${vsDisplay}]\n`;
      text += `  ${item.summary.substring(0, 40)}...\n`;
      if (item.businessRule) {
        const shortRule = item.businessRule.length > 55 
          ? item.businessRule.substring(0, 52) + '...'
          : item.businessRule;
        text += `  ➤ ${shortRule}\n`;
      }
    }
  });
  
  return { text: text.trim(), headerPositions };
}

function styleMasterColumnContent(textBox, headerPositions) {
  const textRange = textBox.getText();
  
  textRange.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(7)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT || '#4A4A4A');
  
  headerPositions.forEach(header => {
    try {
      const headerRange = textRange.getRange(header.start, header.end);
      headerRange.getTextStyle()
        .setBold(true)
        .setFontSize(9)
        .setForegroundColor(header.color);
    } catch (e) {
      // Ignore styling errors
    }
  });
}
// =================================================================================
// AUDIT SLIDE COLORS - Bright warning colors for easy identification
// =================================================================================
const AUDIT_COLORS = {
  BACKGROUND: '#FFF8E7',      // Soft cream/ivory - much easier on eyes
  HEADER_BG: '#E8A838',       // Muted golden/amber header bar
  WARNING_TEXT: '#8B5A2B',    // Muted brown for warning text
  TITLE_TEXT: '#FFFFFF',      // White text on header
  BODY_TEXT: '#4A4A4A',       // Softer dark gray for content
  SECTION_DIVIDER_BG: '#F5EED6'  // Slightly darker cream for section divider
};

function createPortfolioAuditSlide(presentation, auditData, insertPosition) {
  if (!auditData.hasChanges) {
    console.log(`  No changes for ${auditData.portfolioName}, skipping audit slide`);
    return 0;
  }
  
  // Create a blank slide using helper function (works with any template)
  const slide = createBlankSlide(presentation, insertPosition);
  
  // Set WARNING background color (light orange)
  slide.getBackground().setSolidFill(AUDIT_COLORS.BACKGROUND);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // WARNING BANNER at top
  // ═══════════════════════════════════════════════════════════════════════════
  const warningBanner = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 28);
  warningBanner.getFill().setSolidFill(AUDIT_COLORS.HEADER_BG);
  warningBanner.getBorder().setTransparent();
  
  const warningText = warningBanner.getText();
  warningText.setText('⚠️  AUDIT SLIDE - REMOVE BEFORE DISTRIBUTION  ⚠️');
  warningText.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.TITLE_TEXT);
  warningText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // TITLE: Portfolio Name - Changes Summary
  // ═══════════════════════════════════════════════════════════════════════════
  const titleBox = slide.insertTextBox(
    `${auditData.portfolioName} - Changes Summary`,
    20, 40, 680, 35
  );
  titleBox.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor('#502d7f');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CONTENT: Build the changes list with business rules
  // ═══════════════════════════════════════════════════════════════════════════
  let content = '';
  
  function addSection(title, items) {
    if (items.length === 0) return '';
    
    let section = `${title} (${items.length})\n`;
    items.forEach(item => {
      // Main item line with key and value stream
      section += `  • ${item.key} [${item.valueStream}]\n`;
      
      // Summary on its own line for readability
      section += `      ${item.summary}\n`;
      
      // PROMINENT: WHY this is highlighted (moved to immediately after item)
      if (item.businessRule) {
        section += `      ➤ WHY: ${item.businessRule}\n`;
      }
      
      // Details (additional context)
      if (item.details && item.details.length > 0) {
        section += `      Details: ${item.details.join(', ')}\n`;
      }
      
      // RAG info if present
      if (item.ragLevel && item.ragLevel !== '') {
        section += `      ${item.ragLevel}: ${item.ragNote || 'No note'}\n`;
      }
      
      section += '\n';
    });
    return section;
  }
  
  content += addSection('DONE', auditData.done);
  content += addSection('OVERDUE', auditData.overdue || []);
  content += addSection('AT-RISK', auditData.atRisk);
  content += addSection('CHG', auditData.chg);
  content += addSection('NEW', auditData.new);
  content += addSection('DEF', auditData.def);
  content += addSection('CANCELED', auditData.canceled || []);
  content += addSection('INFO', auditData.info || []);
  
  if (content.trim() === '') {
    content = 'No significant changes detected.';
  }
  
  // Add content text box
  const contentBox = slide.insertTextBox(content.trim(), 20, 80, 680, 440);
  const textRange = contentBox.getText();
  
  // Default styling
  textRange.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(9)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT);
  
  // Style section headers with colors
  const fullText = textRange.asString();
  
  const sections = [
    { pattern: /^NEW \(\d+\)/m, color: '#00ACC1' },
    { pattern: /^CHG \(\d+\)/m, color: '#FB8C00' },
    { pattern: /^DONE \(\d+\)/m, color: '#43A047' },
    { pattern: /^DEF \(\d+\)/m, color: '#E53935' },
    { pattern: /^CANCELED \(\d+\)/m, color: '#757575' }, 
    { pattern: /^AT-RISK \(\d+\)/m, color: '#FB8C00' },
    { pattern: /^INFO \(\d+\)/m, color: '#9C27B0' }
  ];
  
  sections.forEach(section => {
    const match = fullText.match(section.pattern);
    if (match) {
      const startIdx = fullText.indexOf(match[0]);
      const endIdx = startIdx + match[0].length;
      try {
        const headerRange = textRange.getRange(startIdx, endIdx);
        headerRange.getTextStyle()
          .setBold(true)
          .setFontSize(11)
          .setForegroundColor(section.color);
      } catch (e) {
        // Ignore styling errors
      }
    }
  });
  
  console.log(`  ✅ Created audit slide for ${auditData.portfolioName}`);
  return 1;
}

function createMasterAuditSlide(presentation, allAuditData, metadata) {
  // Append to end of presentation
  const slide = createBlankSlideAppend(presentation);
  
  // Set soft background
  slide.getBackground().setSolidFill(AUDIT_COLORS.BACKGROUND);
  
  // Warning banner
  const warningBanner = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 28);
  warningBanner.getFill().setSolidFill(AUDIT_COLORS.HEADER_BG);
  warningBanner.getBorder().setTransparent();
  
  const warningText = warningBanner.getText();
  warningText.setText('⚠   AUDIT SLIDE - REMOVE BEFORE DISTRIBUTION   ⚠');
  warningText.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.TITLE_TEXT);
  warningText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Title
  const titleBox = slide.insertTextBox(
    `PI ${metadata.piNumber} - Iteration ${metadata.iterationNumber} | ALL CHANGES SUMMARY`,
    20, 40, 680, 35
  );
  titleBox.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(20)
    .setBold(true)
    .setForegroundColor('#502d7f');
  
  // Aggregate all changes
  const totals = { new: 0, chg: 0, done: 0, def: 0, canceled: 0, overdue: 0, atRisk: 0 };
  allAuditData.forEach(pd => {
    totals.new += pd.new.length;
    totals.chg += pd.chg.length;
    totals.done += pd.done.length;
    totals.def += pd.def.length;
    totals.canceled += (pd.canceled || []).length;
    totals.overdue += (pd.overdue || []).length;
    totals.atRisk += pd.atRisk.length;
  });
  
  // Summary line
  const summaryBox = slide.insertTextBox(
    `NEW: ${totals.new}  |  OVERDUE: ${totals.overdue}  |  CHG: ${totals.chg}  |  DONE: ${totals.done}  |  DEF: ${totals.def}  |  CANCELED: ${totals.canceled}  |  AT-RISK: ${totals.atRisk}`,
    20, 80, 680, 25
  );
  summaryBox.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT);
  
  // Build combined content by category
  let content = '';
  
  ['done', 'overdue', 'chg', 'def', 'canceled', 'atRisk', 'new'].forEach(category => {
    const categoryTitle = category.toUpperCase().replace('ATRISK', 'AT-RISK');
    const items = [];
    
    allAuditData.forEach(pd => {
      const catItems = pd[category] || [];
      catItems.forEach(item => {
        items.push({
          ...item,
          portfolio: pd.portfolioName.substring(0, 12)
        });
      });
    });
    
    if (items.length > 0) {
      content += `${categoryTitle} (${items.length})\n`;
      items.forEach(item => {
        content += `  • ${item.key} [${item.portfolio}] ${item.summary.substring(0, 50)}...\n`;
        // Add business rule on master slide with prominent WHY format
        if (item.businessRule) {
          // Truncate long rules for master slide
          const ruleText = item.businessRule.length > 80 
            ? item.businessRule.substring(0, 77) + '...' 
            : item.businessRule;
          content += `      ➤ WHY: ${ruleText}\n`;
        }
      });
      content += '\n';
    }
  });
  
  if (content.trim() === '') {
    content = 'No changes detected.';
  }
  
  // Add content
  const contentBox = slide.insertTextBox(content.trim(), 20, 110, 680, 400);
  const textRange = contentBox.getText();
  
  textRange.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(8)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT);
  
  // Style section headers with softer colors
  const fullText = textRange.asString();
  const sections = [
    { pattern: /^NEW \(\d+\)/m, color: '#00ACC1' },
    { pattern: /^OVERDUE \(\d+\)/m, color: '#E65100' },
    { pattern: /^CHG \(\d+\)/m, color: '#FB8C00' },
    { pattern: /^DONE \(\d+\)/m, color: '#43A047' },
    { pattern: /^DEF \(\d+\)/m, color: '#E53935' },
    { pattern: /^CANCELED \(\d+\)/m, color: '#757575' }, 
    { pattern: /^AT-RISK \(\d+\)/m, color: '#FB8C00' },
    { pattern: /^INFO \(\d+\)/m, color: '#9C27B0' }
  ];
  sections.forEach(section => {
    const match = fullText.match(section.pattern);
    if (match) {
      const startIdx = fullText.indexOf(match[0]);
      const endIdx = startIdx + match[0].length;
      try {
        textRange.getRange(startIdx, endIdx).getTextStyle()
          .setBold(true)
          .setFontSize(10)
          .setForegroundColor(section.color);
      } catch (e) {}
    }
  });
  
  console.log('  Created master audit slide');
}

function createBlankSlide(presentation, insertPosition) {
  // Duplicate an existing slide (use index 2 - the table template)
  const slides = presentation.getSlides();
  const templateSlide = slides[Math.min(2, slides.length - 1)];
  
  // Duplicate and move to position
  const newSlide = templateSlide.duplicate();
  newSlide.move(insertPosition + 1);  // move() is 1-indexed
  
  // Clear all content from the slide
  const shapes = newSlide.getShapes();
  shapes.forEach(shape => {
    try {
      shape.remove();
    } catch (e) {
      // Some shapes can't be removed
    }
  });
  
  const tables = newSlide.getTables();
  tables.forEach(table => {
    try {
      table.remove();
    } catch (e) {}
  });
  
  const images = newSlide.getImages();
  images.forEach(image => {
    try {
      image.remove();
    } catch (e) {}
  });
  
  return newSlide;
}
function createConsolidatedAuditSection(presentation, allAuditData, metadata) {
  let slidesCreated = 0;
  
  // Filter to only portfolios with changes
  const portfoliosWithChanges = allAuditData.filter(a => a.hasChanges);
  
  if (portfoliosWithChanges.length === 0) {
    console.log('  No audit slides needed - no portfolios with changes');
    return 0;
  }
  
  console.log(`  Creating consolidated audit section with ${portfoliosWithChanges.length} portfolio slides...`);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 1: Audit Section Divider
  // ═══════════════════════════════════════════════════════════════════════════
  const dividerSlide = createBlankSlideAppend(presentation);
  dividerSlide.getBackground().setSolidFill(AUDIT_COLORS.SECTION_DIVIDER_BG || '#F5EED6');
  
  // Warning banner at top
  const dividerBanner = dividerSlide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, 40);
  dividerBanner.getFill().setSolidFill(AUDIT_COLORS.HEADER_BG || '#E8A838');
  dividerBanner.getBorder().setTransparent();
  
  const dividerBannerText = dividerBanner.getText();
  dividerBannerText.setText('⚠️  AUDIT SECTION - REMOVE ALL SLIDES BEFORE DISTRIBUTION  ⚠️');
  dividerBannerText.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(14)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.TITLE_TEXT || '#FFFFFF');
  dividerBannerText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Section title
  const sectionTitle = dividerSlide.insertTextBox('Audit Trail', 60, 180, 600, 60);
  sectionTitle.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(36)
    .setBold(true)
    .setForegroundColor('#502d7f');
  sectionTitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Subtitle with counts - INCLUDING CANCELED
  const totals = { new: 0, chg: 0, done: 0, def: 0, canceled: 0, atRisk: 0, info: 0 };
  portfoliosWithChanges.forEach(pd => {
    totals.new += pd.new.length;
    totals.chg += pd.chg.length;
    totals.done += pd.done.length;
    totals.def += pd.def.length;
    totals.canceled += (pd.canceled || []).length;
    totals.atRisk += pd.atRisk.length;
    totals.info += (pd.info || []).length;
  });
  
  const subtitleText = `PI ${metadata.piNumber} - Iteration ${metadata.iterationNumber}\n\n` +
    `${portfoliosWithChanges.length} Portfolio${portfoliosWithChanges.length > 1 ? 's' : ''} with Changes\n` +
    `NEW: ${totals.new}  |  CHG: ${totals.chg}  |  DONE: ${totals.done}  |  DEF: ${totals.def}  |  CANCELED: ${totals.canceled}  |  AT-RISK: ${totals.atRisk}` +
    (totals.info > 0 ? `  |  INFO: ${totals.info}` : '');
  
  const subtitle = dividerSlide.insertTextBox(subtitleText, 60, 260, 600, 120);
  subtitle.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(16)
    .setForegroundColor(AUDIT_COLORS.BODY_TEXT || '#4A4A4A');
  subtitle.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Instructions
  const instructions = dividerSlide.insertTextBox(
    'These slides show business rules that triggered each item to appear in the report.\n' +
    'Delete this entire section before sharing with stakeholders.',
    60, 420, 600, 60
  );
  instructions.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setItalic(true)
    .setForegroundColor(AUDIT_COLORS.WARNING_TEXT || '#8B5A2B');
  instructions.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  slidesCreated++;
  console.log('    ✅ Created audit section divider slide');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDE 2: Master Summary (all portfolios combined)
  // ═══════════════════════════════════════════════════════════════════════════
  createMasterAuditSlide(presentation, allAuditData, metadata);
  slidesCreated++;
  console.log('    ✅ Created master audit summary slide');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SLIDES 3+: Individual Portfolio Audit Slides (with pagination)
  // ═══════════════════════════════════════════════════════════════════════════
  portfoliosWithChanges.forEach((auditData, index) => {
    const portfolioSlides = createPortfolioAuditSlidesWithPagination(presentation, auditData);
    slidesCreated += portfolioSlides;
    console.log(`    ✅ Created ${portfolioSlides} audit slide(s) for ${auditData.portfolioName}`);
  });
  
  console.log(`\n✅ Total audit slides created: ${slidesCreated}`);
  return slidesCreated;
}

function createPortfolioAuditSlidesWithPagination(presentation, auditData) {
  if (!auditData.hasChanges) {
    console.log(`  No changes for ${auditData.portfolioName}, skipping audit slide`);
    return 0;
  }
  
  // Flatten all items into a single array with category labels
  const allItems = [];
  
  // Order: DONE → OVERDUE → AT-RISK → CHG → NEW → DEF → CANCELED → INFO
  const categories = [
    { key: 'done', title: 'DONE', color: '#43A047' },
    { key: 'pending', title: 'PENDING CLOSURE', color: '#81C784' },
    { key: 'overdue', title: 'OVERDUE', color: '#E65100' },
    { key: 'atRisk', title: 'AT-RISK', color: '#FB8C00' },
    { key: 'chg', title: 'CHG', color: '#FB8C00' },
    { key: 'new', title: 'NEW', color: '#00ACC1' },
    { key: 'def', title: 'DEF', color: '#E53935' },
    { key: 'canceled', title: 'CANCELED', color: '#757575' }, 
    { key: 'info', title: 'INFO', color: '#9C27B0' }
  ];

  
  categories.forEach(cat => {
    const items = auditData[cat.key] || [];
    if (items.length > 0) {
      // Add category header
      allItems.push({
        type: 'header',
        text: `${cat.title} (${items.length})`,
        color: cat.color
      });
      
      // Add items
      items.forEach(item => {
        allItems.push({
          type: 'item',
          key: item.key,
          valueStream: item.valueStream,
          summary: item.summary,
          businessRule: item.businessRule,
          details: item.details,
          ragLevel: item.ragLevel,
          ragNote: item.ragNote,
          ragUnchanged: item.ragUnchanged
        });
      });
    }
  });
  
  if (allItems.length === 0) {
    console.log(`  No items to display for ${auditData.portfolioName}`);
    return 0;
  }
  
  // Calculate how many slides we need
  // Each item takes ~4 lines, headers take ~2 lines
  let totalLines = 0;
  allItems.forEach(item => {
    totalLines += item.type === 'header' ? 2 : AUDIT_SLIDE_CONFIG.LINES_PER_ITEM;
  });
  
  const linesPerSlide = AUDIT_SLIDE_CONFIG.MAX_LINES_PER_COLUMN * 2;  // Two columns
  const numSlides = Math.ceil(totalLines / linesPerSlide);
  
  console.log(`  📊 Audit: ${auditData.portfolioName} has ${allItems.length} items, ${totalLines} lines, needs ${numSlides} slide(s)`);
  
  // Split items across slides
  let currentSlideItems = [];
  let currentSlideLines = 0;
  const slideItemGroups = [];
  
  allItems.forEach(item => {
    const itemLines = item.type === 'header' ? 2 : AUDIT_SLIDE_CONFIG.LINES_PER_ITEM;
    
    if (currentSlideLines + itemLines > linesPerSlide && currentSlideItems.length > 0) {
      // Current slide is full, start a new one
      slideItemGroups.push(currentSlideItems);
      currentSlideItems = [];
      currentSlideLines = 0;
    }
    
    currentSlideItems.push(item);
    currentSlideLines += itemLines;
  });
  
  // Don't forget the last group
  if (currentSlideItems.length > 0) {
    slideItemGroups.push(currentSlideItems);
  }
  
  // Create slides
  let slidesCreated = 0;
  slideItemGroups.forEach((items, pageIndex) => {
    createSingleAuditSlide(
      presentation, 
      auditData.portfolioName, 
      items, 
      pageIndex + 1, 
      slideItemGroups.length
    );
    slidesCreated++;
  });
  
  console.log(`  ✅ Created ${slidesCreated} audit slide(s) for ${auditData.portfolioName}`);
  return slidesCreated;
}
function createSingleAuditSlide(presentation, portfolioName, items, pageNum, totalPages) {
  const slide = createBlankSlideAppend(presentation);
  
  // Set background
  slide.getBackground().setSolidFill(AUDIT_COLORS.BACKGROUND || '#FFF8E7');
  
  // Warning banner at top
  const warningBanner = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, 720, AUDIT_SLIDE_CONFIG.HEADER_HEIGHT);
  warningBanner.getFill().setSolidFill(AUDIT_COLORS.HEADER_BG || '#E8A838');
  warningBanner.getBorder().setTransparent();
  
  const warningText = warningBanner.getText();
  warningText.setText('⚠️  AUDIT SLIDE - REMOVE BEFORE DISTRIBUTION  ⚠️');
  warningText.getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor(AUDIT_COLORS.TITLE_TEXT || '#FFFFFF');
  warningText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  
  // Title with page number
  const pageIndicator = totalPages > 1 ? ` (${pageNum}/${totalPages})` : '';
  const titleBox = slide.insertTextBox(
    `${portfolioName} - Changes Summary${pageIndicator}`,
    AUDIT_SLIDE_CONFIG.MARGIN, 
    AUDIT_SLIDE_CONFIG.TITLE_TOP, 
    AUDIT_SLIDE_CONFIG.SLIDE_WIDTH - (2 * AUDIT_SLIDE_CONFIG.MARGIN), 
    AUDIT_SLIDE_CONFIG.TITLE_HEIGHT
  );
  titleBox.getText().getTextStyle()
    .setFontFamily('Lato')
    .setFontSize(18)
    .setBold(true)
    .setForegroundColor('#502d7f');
  
  // Split items into two columns
  const midPoint = Math.ceil(items.length / 2);
  const leftItems = items.slice(0, midPoint);
  const rightItems = items.slice(midPoint);
  
  // Create LEFT column
  const leftContent = buildColumnContent(leftItems);
  const leftBox = slide.insertTextBox(
    leftContent.text,
    AUDIT_SLIDE_CONFIG.LEFT_COLUMN_X,
    AUDIT_SLIDE_CONFIG.CONTENT_TOP,
    AUDIT_SLIDE_CONFIG.COLUMN_WIDTH,
    AUDIT_SLIDE_CONFIG.CONTENT_HEIGHT
  );
  styleColumnContent(leftBox, leftContent.headerPositions);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // DEFENSIVE: Only create RIGHT column if there are items for it
  // ═══════════════════════════════════════════════════════════════════════════
  if (rightItems.length > 0) {
    const rightContent = buildColumnContent(rightItems);
    const rightBox = slide.insertTextBox(
      rightContent.text,
      AUDIT_SLIDE_CONFIG.RIGHT_COLUMN_X,
      AUDIT_SLIDE_CONFIG.CONTENT_TOP,
      AUDIT_SLIDE_CONFIG.COLUMN_WIDTH,
      AUDIT_SLIDE_CONFIG.CONTENT_HEIGHT
    );
    styleColumnContent(rightBox, rightContent.headerPositions);
  }
  
  // Add column divider line
  const dividerX = AUDIT_SLIDE_CONFIG.MARGIN + AUDIT_SLIDE_CONFIG.COLUMN_WIDTH + (AUDIT_SLIDE_CONFIG.COLUMN_GAP / 2);
  const divider = slide.insertLine(
    SlidesApp.LineCategory.STRAIGHT,
    dividerX, AUDIT_SLIDE_CONFIG.CONTENT_TOP,
    dividerX, AUDIT_SLIDE_CONFIG.CONTENT_BOTTOM
  );
  divider.getLineFill().setSolidFill('#CCCCCC');
  divider.setWeight(1);
  
  return slide;
}

function buildColumnContent(items) {
  let text = '';
  const headerPositions = [];  // Track header positions for styling
  
  items.forEach(item => {
    if (item.type === 'header') {
      const startPos = text.length;
      text += `${item.text}\n`;
      headerPositions.push({
        start: startPos,
        end: text.length - 1,
        color: item.color
      });
    } else {
      // Item format (compact for two columns)
      text += `• ${item.key} [${item.valueStream}]\n`;
      
      // Summary (truncated for space)
      const shortSummary = item.summary.length > 40 
        ? item.summary.substring(0, 37) + '...' 
        : item.summary;
      text += `  ${shortSummary}\n`;
      
      // Business Rule (the WHY)
      if (item.businessRule) {
        const shortRule = item.businessRule.length > 60
          ? item.businessRule.substring(0, 57) + '...'
          : item.businessRule;
        text += `  ➤ WHY: ${shortRule}\n`;
      }
      
      text += '\n';  // Blank line between items
    }
  });
  
  return { text: text.trim(), headerPositions };
}
function styleColumnContent(textBox, headerPositions) {
  // ═══════════════════════════════════════════════════════════════════════════
  // DEFENSIVE: Wrap entire function in try-catch to prevent slide generation failure
  // ═══════════════════════════════════════════════════════════════════════════
  try {
    const textRange = textBox.getText();
    
    // DEFENSIVE: Check if text box has any content before styling
    const textContent = textRange.asString();
    if (!textContent || textContent.trim() === '') {
      console.log('  ⚠️ styleColumnContent: Empty text box, skipping styling');
      return;
    }
    
    // Default styling
    textRange.getTextStyle()
      .setFontFamily('Lato')
      .setFontSize(AUDIT_SLIDE_CONFIG.FONT_SIZE)
      .setForegroundColor(AUDIT_COLORS.BODY_TEXT || '#4A4A4A');
    
    // Style headers with colors
    headerPositions.forEach(header => {
      try {
        const headerRange = textRange.getRange(header.start, header.end);
        headerRange.getTextStyle()
          .setBold(true)
          .setFontSize(10)
          .setForegroundColor(header.color);
      } catch (e) {
        console.warn(`Could not style header: ${e.message}`);
      }
    });
  } catch (e) {
    console.error(`styleColumnContent error: ${e.message}`);
    // Non-fatal - continue with slide generation
  }
}

function addAuditSection(title, items) {
  if (items.length === 0) return '';
  
  let section = `${title} (${items.length})\n`;
  items.forEach(item => {
    // Add indicator for unchanged RAG in AT-RISK section
    let prefix = '•';
    if (item.ragUnchanged) {
      prefix = '⚡';  // Lightning bolt to indicate persisting risk
    }
    
    // Main item line with key and value stream
    section += `  ${prefix} ${item.key} [${item.valueStream}]\n`;
    
    // Summary on its own line
    section += `      ${item.summary}\n`;
    
    // PROMINENT: WHY this is highlighted
    if (item.businessRule) {
      section += `      ➤ WHY: ${item.businessRule}\n`;
    }
    
    // Details (additional context)
    if (item.details && item.details.length > 0) {
      section += `      Details: ${item.details.join(', ')}\n`;
    }
    
    // RAG info if present
    if (item.ragLevel && item.ragLevel !== '') {
      section += `      ${item.ragLevel}: ${item.ragNote || 'No note'}\n`;
    }
    
    section += '\n';
  });
  return section;
}
function createBlankSlideAppend(presentation) {
  try {
    // ═══════════════════════════════════════════════════════════════════════════
    // Method 1: Try to find a blank layout in the presentation's masters
    // ═══════════════════════════════════════════════════════════════════════════
    const masters = presentation.getMasters();
    if (masters && masters.length > 0) {
      const layouts = masters[0].getLayouts();
      
      // Look for a blank or minimal layout
      for (const layout of layouts) {
        try {
          const layoutName = layout.getLayoutName().toLowerCase();
          if (layoutName.includes('blank') || layoutName.includes('empty')) {
            return presentation.appendSlide(layout);
          }
        } catch (layoutError) {
          // Continue to next layout
        }
      }
      
      // If no blank layout found, use the first layout and clear it
      if (layouts && layouts.length > 0) {
        try {
          const newSlide = presentation.appendSlide(layouts[0]);
          
          // Remove all placeholder elements
          newSlide.getShapes().forEach(shape => {
            try { shape.remove(); } catch (e) {}
          });
          newSlide.getTables().forEach(table => {
            try { table.remove(); } catch (e) {}
          });
          newSlide.getImages().forEach(image => {
            try { image.remove(); } catch (e) {}
          });
          
          return newSlide;
        } catch (layoutSlideError) {
          console.warn(`Could not use layout: ${layoutSlideError.message}`);
        }
      }
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Method 2: Duplicate last slide and clear it
    // ═══════════════════════════════════════════════════════════════════════════
    const slides = presentation.getSlides();
    if (slides && slides.length > 0) {
      try {
        const lastSlide = slides[slides.length - 1];
        const newSlide = lastSlide.duplicate();
        
        // Clear all content
        newSlide.getShapes().forEach(shape => {
          try { shape.remove(); } catch (e) {}
        });
        newSlide.getTables().forEach(table => {
          try { table.remove(); } catch (e) {}
        });
        newSlide.getImages().forEach(image => {
          try { image.remove(); } catch (e) {}
        });
        newSlide.getLines().forEach(line => {
          try { line.remove(); } catch (e) {}
        });
        
        // Reset background
        newSlide.getBackground().setSolidFill('#FFFFFF');
        
        return newSlide;
      } catch (dupError) {
        console.warn(`Could not duplicate slide: ${dupError.message}`);
      }
    }
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Method 3: Last resort - appendSlide with no arguments
    // ═══════════════════════════════════════════════════════════════════════════
    console.log('Using fallback appendSlide()');
    return presentation.appendSlide();
    
  } catch (e) {
    console.error(`Error creating blank slide: ${e.message}`);
    // Ultimate fallback
    return presentation.appendSlide();
  }
}
function extractStringValue(fieldValue, defaultValue = '') {
  if (!fieldValue) return defaultValue;
  if (typeof fieldValue === 'string') return fieldValue.trim() || defaultValue;
  if (typeof fieldValue === 'object') {
    return (fieldValue.value || fieldValue.name || String(fieldValue)).trim() || defaultValue;
  }
  return String(fieldValue).trim() || defaultValue;
}