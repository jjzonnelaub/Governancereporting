// ═══════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════

const PORTFOLIO_DIST_CONFIG = {
  // Value Stream Colors (matches existing featurepoint.gs)
  VS_COLORS: {
    'MMPM': '#6C207F',                    // Purple
    'EMA Clinical': '#1769FC',            // Blue
    'RCM Genie': '#FFC42E',              // Yellow/Gold
    'Clinical Quality': '#3F68DD',        // Light Blue
    'Data Governance': '#A8DADC',         // Teal
    'Enterprise': '#FFE2C8',              // Peach
    'Patient Collaboration': '#FF6B9D',   // Pink
    'MMGI': '#2ECC71',                    // Green
    'MMGI-Cloud': '#27AE60',              // Darker Green
    'Cloud Operations': '#9B59B6',        // Violet
    'Platform Engineering': '#E74C3C',    // Red
    'Revenue Cycle': '#F39C12',           // Orange
    'Infosec': '#1ABC9C',                 // Turquoise
    'Other': '#95A5A6'                    // Gray
  },
  
  // Program Initiative Colors (distinct palette)
  PROG_INIT_COLORS: [
    '#2E86AB',  // Steel Blue
    '#A23B72',  // Mulberry
    '#F18F01',  // Tangerine
    '#C73E1D',  // Vermillion
    '#3B1F2B',  // Dark Purple
    '#6B4226',  // Brown
    '#44AF69',  // Emerald
    '#FCAB10',  // Amber
    '#2C514C',  // Dark Teal
    '#E84855',  // Coral Red
    '#7209B7',  // Purple
    '#3A0CA3',  // Indigo
    '#4361EE',  // Royal Blue
    '#4CC9F0',  // Sky Blue
    '#F72585'   // Hot Pink
  ],
  
  // Sheet configuration
  HEADER_ROW: 4,
  DATA_START_ROW: 5,
  
  // Chart dimensions
  CHART_WIDTH: 600,
  CHART_HEIGHT: 300,
  
  // Slide layout
  SLIDE_WIDTH: 720,
  SLIDE_HEIGHT: 540,
  MARGIN: 20
};

// ═══════════════════════════════════════════════════════════════════════════
// MAIN ENTRY POINT - CREATE SUMMARY SHEET WITH FORMULAS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Creates or updates the Feature Point Summary sheet with SUMIFS formulas
 * These formulas automatically update when PI report data changes
 * 
 * @param {number} piNumber - The PI number (e.g., 13)
 * @param {string} iterationNumber - The iteration to pull data from (default: 0)
 * @returns {Object} - Sheet reference and chart positions
 */
function createFeaturePointSummarySheet(piNumber, iterationNumber) {
  iterationNumber = iterationNumber || 0;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheetName = `PI ${piNumber} - Iteration ${iterationNumber}`;
  const summarySheetName = `PI ${piNumber} - Feature Point Summary`;
  
  console.log(`\n${'='.repeat(80)}`);
  console.log(`CREATING FEATURE POINT SUMMARY SHEET - PI ${piNumber}`);
  console.log(`Source: ${sourceSheetName}`);
  console.log(`${'='.repeat(80)}\n`);
  
  // Validate source sheet exists
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) {
    throw new Error(`Source sheet not found: ${sourceSheetName}`);
  }
  
  // Get or create summary sheet
  let summarySheet = ss.getSheetByName(summarySheetName);
  if (summarySheet) {
    // Clear existing charts and data
    summarySheet.getCharts().forEach(chart => summarySheet.removeChart(chart));
    summarySheet.clear();
    console.log(`✓ Cleared existing summary sheet: ${summarySheetName}`);
  } else {
    summarySheet = ss.insertSheet(summarySheetName);
    console.log(`✓ Created new summary sheet: ${summarySheetName}`);
  }
  
  // Set font for entire sheet
  summarySheet.getRange(1, 1, summarySheet.getMaxRows(), summarySheet.getMaxColumns())
    .setFontFamily('Lato');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 1: Get unique Portfolio Initiatives, Value Streams, and Program Initiatives
  // ═══════════════════════════════════════════════════════════════════════════
  
  const sourceData = sourceSheet.getDataRange().getValues();
  const headers = sourceData[PORTFOLIO_DIST_CONFIG.HEADER_ROW - 1];
  
  const colIndices = {
    portfolioInit: headers.indexOf('Portfolio Initiative'),
    programInit: headers.indexOf('Program Initiative'),
    valueStream: headers.indexOf('Value Stream/Org'),
    featurePoints: headers.indexOf('Feature Points'),
    piCommitment: headers.indexOf('PI Commitment'),
    issueType: headers.indexOf('Issue Type')
  };
  
  console.log('Column indices:', JSON.stringify(colIndices));
  
  // Validate columns exist
  if (colIndices.portfolioInit === -1 || colIndices.valueStream === -1 || 
      colIndices.featurePoints === -1 || colIndices.piCommitment === -1) {
    throw new Error('Required columns not found in source sheet');
  }
  
  // Extract unique values (only from committed epics)
  const portfolioInits = new Set();
  const valueStreams = new Set();
  const programInits = new Set();
  
  for (let i = PORTFOLIO_DIST_CONFIG.DATA_START_ROW - 1; i < sourceData.length; i++) {
    const row = sourceData[i];
    const piCommitment = (row[colIndices.piCommitment] || '').toString().trim();
    const issueType = (row[colIndices.issueType] || '').toString().trim();
    
    // Only include Committed or Committed After Plan epics
    if (issueType === 'Epic' && 
        (piCommitment === 'Committed' || piCommitment === 'Committed After Plan')) {
      
      const portfolioInit = (row[colIndices.portfolioInit] || '').toString().trim();
      const valueStream = (row[colIndices.valueStream] || '').toString().trim();
      const programInit = (row[colIndices.programInit] || '').toString().trim();
      
      if (portfolioInit) portfolioInits.add(portfolioInit);
      if (valueStream) valueStreams.add(valueStream);
      if (programInit) programInits.add(programInit);
    }
  }
  
  // Sort alphabetically
  const sortedPortfolioInits = Array.from(portfolioInits).sort();
  const sortedValueStreams = Array.from(valueStreams).sort();
  const sortedProgramInits = Array.from(programInits).sort();
  
  console.log(`Found ${sortedPortfolioInits.length} Portfolio Initiatives`);
  console.log(`Found ${sortedValueStreams.length} Value Streams`);
  console.log(`Found ${sortedProgramInits.length} Program Initiatives`);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 2: Create Table 1 - Portfolio Initiative × Value Stream
  // ═══════════════════════════════════════════════════════════════════════════
  
  console.log('\n📊 Creating Table 1: Portfolio Initiative × Value Stream...');
  
  const table1StartRow = 1;
  const table1StartCol = 1;
  
  // Title
  summarySheet.getRange(table1StartRow, table1StartCol)
    .setValue(`PI ${piNumber} - Feature Points by Portfolio Initiative & Value Stream`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#663399');
  
  // Subtitle with filter info
  summarySheet.getRange(table1StartRow + 1, table1StartCol)
    .setValue('Filter: PI Commitment = Committed OR Committed After Plan')
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  // Header row
  const table1HeaderRow = table1StartRow + 3;
  const table1Headers = ['Portfolio Initiative', ...sortedValueStreams, 'Total'];
  summarySheet.getRange(table1HeaderRow, table1StartCol, 1, table1Headers.length)
    .setValues([table1Headers])
    .setBackground('#663399')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Data rows with SUMIFS formulas
  const table1DataStartRow = table1HeaderRow + 1;
  
  sortedPortfolioInits.forEach((portfolioInit, rowIndex) => {
    const currentRow = table1DataStartRow + rowIndex;
    
    // Portfolio Initiative name
    summarySheet.getRange(currentRow, table1StartCol).setValue(portfolioInit);
    
    // SUMIFS formulas for each Value Stream
    sortedValueStreams.forEach((vs, colIndex) => {
      const formula = createSUMIFSFormula(
        sourceSheetName,
        colIndices,
        portfolioInit,
        vs,
        null  // No Program Initiative filter
      );
      summarySheet.getRange(currentRow, table1StartCol + 1 + colIndex).setFormula(formula);
    });
    
    // Total formula (SUM of row)
    const startCol = columnToLetter(table1StartCol + 1);
    const endCol = columnToLetter(table1StartCol + sortedValueStreams.length);
    summarySheet.getRange(currentRow, table1StartCol + sortedValueStreams.length + 1)
      .setFormula(`=SUM(${startCol}${currentRow}:${endCol}${currentRow})`)
      .setFontWeight('bold');
  });
  
  // Total row
  const table1TotalRow = table1DataStartRow + sortedPortfolioInits.length;
  summarySheet.getRange(table1TotalRow, table1StartCol)
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setBackground('#E8E0FF');
  
  for (let colIndex = 0; colIndex <= sortedValueStreams.length; colIndex++) {
    const col = table1StartCol + 1 + colIndex;
    const colLetter = columnToLetter(col);
    summarySheet.getRange(table1TotalRow, col)
      .setFormula(`=SUM(${colLetter}${table1DataStartRow}:${colLetter}${table1TotalRow - 1})`)
      .setFontWeight('bold')
      .setBackground('#E8E0FF');
  }
  
  // Format numbers
  summarySheet.getRange(table1DataStartRow, table1StartCol + 1, 
    sortedPortfolioInits.length + 1, sortedValueStreams.length + 1)
    .setNumberFormat('#,##0');
  
  // Add borders
  summarySheet.getRange(table1HeaderRow, table1StartCol, 
    sortedPortfolioInits.length + 2, table1Headers.length)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 3: Create Table 2 - Portfolio Initiative × Program Initiative
  // ═══════════════════════════════════════════════════════════════════════════
  
  console.log('\n📊 Creating Table 2: Portfolio Initiative × Program Initiative...');
  
  const table2StartRow = table1TotalRow + 4;
  const table2StartCol = 1;
  
  // Title
  summarySheet.getRange(table2StartRow, table2StartCol)
    .setValue(`PI ${piNumber} - Feature Points by Portfolio Initiative & Program Initiative`)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor('#663399');
  
  // Subtitle
  summarySheet.getRange(table2StartRow + 1, table2StartCol)
    .setValue('Filter: PI Commitment = Committed OR Committed After Plan')
    .setFontSize(10)
    .setFontStyle('italic')
    .setFontColor('#666666');
  
  // Header row
  const table2HeaderRow = table2StartRow + 3;
  const table2Headers = ['Portfolio Initiative', ...sortedProgramInits, 'Total'];
  summarySheet.getRange(table2HeaderRow, table2StartCol, 1, table2Headers.length)
    .setValues([table2Headers])
    .setBackground('#663399')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Data rows with SUMIFS formulas
  const table2DataStartRow = table2HeaderRow + 1;
  
  sortedPortfolioInits.forEach((portfolioInit, rowIndex) => {
    const currentRow = table2DataStartRow + rowIndex;
    
    // Portfolio Initiative name
    summarySheet.getRange(currentRow, table2StartCol).setValue(portfolioInit);
    
    // SUMIFS formulas for each Program Initiative
    sortedProgramInits.forEach((progInit, colIndex) => {
      const formula = createSUMIFSFormula(
        sourceSheetName,
        colIndices,
        portfolioInit,
        null,  // No Value Stream filter
        progInit
      );
      summarySheet.getRange(currentRow, table2StartCol + 1 + colIndex).setFormula(formula);
    });
    
    // Total formula (SUM of row)
    const startCol = columnToLetter(table2StartCol + 1);
    const endCol = columnToLetter(table2StartCol + sortedProgramInits.length);
    summarySheet.getRange(currentRow, table2StartCol + sortedProgramInits.length + 1)
      .setFormula(`=SUM(${startCol}${currentRow}:${endCol}${currentRow})`)
      .setFontWeight('bold');
  });
  
  // Total row
  const table2TotalRow = table2DataStartRow + sortedPortfolioInits.length;
  summarySheet.getRange(table2TotalRow, table2StartCol)
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setBackground('#E8E0FF');
  
  for (let colIndex = 0; colIndex <= sortedProgramInits.length; colIndex++) {
    const col = table2StartCol + 1 + colIndex;
    const colLetter = columnToLetter(col);
    summarySheet.getRange(table2TotalRow, col)
      .setFormula(`=SUM(${colLetter}${table2DataStartRow}:${colLetter}${table2TotalRow - 1})`)
      .setFontWeight('bold')
      .setBackground('#E8E0FF');
  }
  
  // Format numbers
  summarySheet.getRange(table2DataStartRow, table2StartCol + 1, 
    sortedPortfolioInits.length + 1, sortedProgramInits.length + 1)
    .setNumberFormat('#,##0');
  
  // Add borders
  summarySheet.getRange(table2HeaderRow, table2StartCol, 
    sortedPortfolioInits.length + 2, table2Headers.length)
    .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 4: Create Stacked Bar Charts
  // ═══════════════════════════════════════════════════════════════════════════
  
  console.log('\n📊 Creating stacked bar charts...');
  
  // Auto-fit columns
  summarySheet.autoResizeColumns(1, Math.max(table1Headers.length, table2Headers.length));
  
  // Chart 1: Portfolio Initiative × Value Stream
  const chart1 = createStackedBarChart(
    summarySheet,
    table1HeaderRow,
    table1StartCol,
    sortedPortfolioInits.length + 1,  // +1 for header
    sortedValueStreams.length + 1,     // +1 for Portfolio Init column
    `PI ${piNumber}: Feature Points by Portfolio Initiative & Value Stream`,
    table1TotalRow + 2,  // Position below table
    table1StartCol
  );
  
  // Chart 2: Portfolio Initiative × Program Initiative
  const chart2 = createStackedBarChart(
    summarySheet,
    table2HeaderRow,
    table2StartCol,
    sortedPortfolioInits.length + 1,
    sortedProgramInits.length + 1,
    `PI ${piNumber}: Feature Points by Portfolio Initiative & Program Initiative`,
    table2TotalRow + 2,
    table2StartCol
  );
  
  console.log(`\n✅ Feature Point Summary sheet created: ${summarySheetName}`);
  
  return {
    sheet: summarySheet,
    sheetName: summarySheetName,
    chart1Row: table1TotalRow + 2,
    chart2Row: table2TotalRow + 2,
    portfolioInits: sortedPortfolioInits,
    valueStreams: sortedValueStreams,
    programInits: sortedProgramInits
  };
}

/**
 * Creates a SUMIFS formula for the Feature Point Summary
 */
function createSUMIFSFormula(sourceSheetName, colIndices, portfolioInit, valueStream, programInit) {
  const fpCol = columnToLetter(colIndices.featurePoints + 1);
  const portfolioCol = columnToLetter(colIndices.portfolioInit + 1);
  const vsCol = columnToLetter(colIndices.valueStream + 1);
  const progInitCol = columnToLetter(colIndices.programInit + 1);
  const piCommitCol = columnToLetter(colIndices.piCommitment + 1);
  const issueTypeCol = columnToLetter(colIndices.issueType + 1);
  
  // Escape sheet name for formula
  const sheetRef = `'${sourceSheetName}'`;
  
  // Build criteria array
  let formula = `=SUMIFS(${sheetRef}!${fpCol}:${fpCol}`;
  
  // Always filter by Portfolio Initiative
  formula += `,${sheetRef}!${portfolioCol}:${portfolioCol},"${portfolioInit}"`;
  
  // Filter by Value Stream if provided
  if (valueStream) {
    formula += `,${sheetRef}!${vsCol}:${vsCol},"${valueStream}"`;
  }
  
  // Filter by Program Initiative if provided
  if (programInit) {
    formula += `,${sheetRef}!${progInitCol}:${progInitCol},"${programInit}"`;
  }
  
  // Always filter by Issue Type = Epic
  formula += `,${sheetRef}!${issueTypeCol}:${issueTypeCol},"Epic"`;
  
  // Filter by PI Commitment (Committed OR Committed After Plan)
  // Use SUMIFS twice and add them together for OR logic
  const formula2 = formula.replace('=SUMIFS', 'SUMIFS');
  
  formula += `,${sheetRef}!${piCommitCol}:${piCommitCol},"Committed")`;
  const formulaAfterPlan = formula2 + `,${sheetRef}!${piCommitCol}:${piCommitCol},"Committed After Plan")`;
  
  return `=${formula.substring(1)}+${formulaAfterPlan}`;
}

/**
 * Converts column number to letter (1=A, 2=B, etc.)
 */
function columnToLetter(columnNum) {
  let letter = '';
  while (columnNum > 0) {
    const mod = (columnNum - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    columnNum = Math.floor((columnNum - 1) / 26);
  }
  return letter;
}

/**
 * Creates a stacked bar chart from table data
 */
function createStackedBarChart(sheet, headerRow, startCol, numRows, numCols, title, chartRow, chartCol) {
  try {
    // Get data range (excluding Total column)
    const dataRange = sheet.getRange(headerRow, startCol, numRows, numCols);
    
    // Build the chart
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dataRange)
      .setPosition(chartRow, chartCol, 0, 0)
      .setOption('title', title)
      .setOption('titleTextStyle', { fontSize: 14, bold: true, color: '#333333' })
      .setOption('isStacked', true)
      .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
      .setOption('hAxis', { 
        title: 'Feature Points',
        titleTextStyle: { fontSize: 11, italic: true },
        textStyle: { fontSize: 10 }
      })
      .setOption('vAxis', { 
        title: 'Portfolio Initiative',
        titleTextStyle: { fontSize: 11, italic: true },
        textStyle: { fontSize: 10 }
      })
      .setOption('chartArea', { left: '25%', top: '10%', width: '50%', height: '75%' })
      .setOption('bar', { groupWidth: '80%' })
      .setOption('width', 700)
      .setOption('height', 400);
    
    const chart = chartBuilder.build();
    sheet.insertChart(chart);
    
    console.log(`  ✅ Created chart: ${title}`);
    return chart;
    
  } catch (error) {
    console.error(`  ❌ Error creating chart: ${error.message}`);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SLIDE GENERATION - ADD TO GOVERNANCE PRESENTATION
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Adds Portfolio Distribution slides to the end of a governance presentation
 * 
 * @param {Presentation} presentation - The Google Slides presentation
 * @param {number} piNumber - The PI number
 * @param {string} valueStream - The value stream (or 'ALL' for all value streams)
 */
function addPortfolioDistributionSlides(presentation, piNumber, valueStream) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`ADDING PORTFOLIO DISTRIBUTION SLIDES - PI ${piNumber}`);
  console.log(`Value Stream: ${valueStream}`);
  console.log(`${'='.repeat(80)}\n`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Create/update the summary sheet first
    const summaryInfo = createFeaturePointSummarySheet(piNumber, 0);
    
    if (!summaryInfo || !summaryInfo.sheet) {
      console.error('❌ Failed to create summary sheet');
      return;
    }
    
    // Give Sheets a moment to render charts
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    
    // Get the charts from the summary sheet
    const charts = summaryInfo.sheet.getCharts();
    console.log(`Found ${charts.length} charts on summary sheet`);
    
    if (charts.length < 2) {
      console.warn('⚠️ Expected 2 charts, found ' + charts.length);
    }
    
    // ═══════════════════════════════════════════════════════════════════════
    // SLIDE 1: Section Header - "Portfolio Distribution"
    // ═══════════════════════════════════════════════════════════════════════
    
    const headerSlide = createBlankSlide(presentation);
    addDisclosureToSlide(headerSlide);
    
    // Purple background for section header
    headerSlide.getBackground().setSolidFill('#663399');
    
    // Title
    const titleBox = headerSlide.insertTextBox('Portfolio Distribution', 60, 200, 600, 80);
    titleBox.getText().getTextStyle()
      .setFontSize(36)
      .setForegroundColor('#FFFFFF')
      .setBold(true)
      .setFontFamily('Montserrat');
    titleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Subtitle
    const subtitleBox = headerSlide.insertTextBox(
      `Feature Point Distribution by Portfolio Initiative\nPI ${piNumber}`, 
      60, 290, 600, 60
    );
    subtitleBox.getText().getTextStyle()
      .setFontSize(18)
      .setForegroundColor('#FFFFFF')
      .setFontFamily('Lato');
    subtitleBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    console.log('✅ Added Portfolio Distribution section header slide');
    
    // ═══════════════════════════════════════════════════════════════════════
    // SLIDE 2: Charts (try to fit both on one slide)
    // ═══════════════════════════════════════════════════════════════════════
    
    const chartSlide = createBlankSlide(presentation);
    addDisclosureToSlide(chartSlide);
    
    // Slide title
    const slideTitleBox = chartSlide.insertTextBox(
      `PI ${piNumber}: Feature Point Distribution`, 
      20, 10, 680, 30
    );
    slideTitleBox.getText().getTextStyle()
      .setFontSize(16)
      .setForegroundColor('#663399')
      .setBold(true)
      .setFontFamily('Montserrat');
    
    // Embed charts as images
    if (charts.length >= 1) {
      try {
        const chart1Blob = charts[0].getAs('image/png');
        const chart1Image = chartSlide.insertImage(chart1Blob, 20, 50, 340, 220);
        console.log('✅ Embedded Chart 1 (Portfolio × Value Stream)');
      } catch (e) {
        console.error(`❌ Failed to embed Chart 1: ${e.message}`);
      }
    }
    
    if (charts.length >= 2) {
      try {
        const chart2Blob = charts[1].getAs('image/png');
        const chart2Image = chartSlide.insertImage(chart2Blob, 370, 50, 340, 220);
        console.log('✅ Embedded Chart 2 (Portfolio × Program Initiative)');
      } catch (e) {
        console.error(`❌ Failed to embed Chart 2: ${e.message}`);
      }
    }
    
    // Add data source note
    const dataSourceBox = chartSlide.insertTextBox(
      `Data Source: ${summaryInfo.sheetName} | Filter: Committed + Committed After Plan`,
      20, 280, 680, 20
    );
    dataSourceBox.getText().getTextStyle()
      .setFontSize(8)
      .setForegroundColor('#999999')
      .setFontFamily('Lato')
      .setItalic(true);
    
    // Add summary statistics
    const statsText = `Portfolio Initiatives: ${summaryInfo.portfolioInits.length} | ` +
                     `Value Streams: ${summaryInfo.valueStreams.length} | ` +
                     `Program Initiatives: ${summaryInfo.programInits.length}`;
    const statsBox = chartSlide.insertTextBox(statsText, 20, 300, 680, 20);
    statsBox.getText().getTextStyle()
      .setFontSize(9)
      .setForegroundColor('#666666')
      .setFontFamily('Lato');
    
    console.log('✅ Added Portfolio Distribution chart slide');
    
    return {
      success: true,
      slidesAdded: 2,
      summarySheetName: summaryInfo.sheetName
    };
    
  } catch (error) {
    console.error(`❌ Error adding Portfolio Distribution slides: ${error.message}`);
    console.error(error.stack);
    return {
      success: false,
      error: error.message
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// STANDALONE TEST FUNCTION
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Test function to create summary sheet without generating slides
 * Run this from the Apps Script editor to test the sheet creation
 */
function testCreateFeaturePointSummarySheet() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Test Feature Point Summary',
    'Enter PI number:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const piNumber = parseInt(response.getResponseText().trim());
  if (isNaN(piNumber)) {
    ui.alert('Invalid PI number');
    return;
  }
  
  try {
    const result = createFeaturePointSummarySheet(piNumber, 0);
    ui.alert('Success', 
      `Created sheet: ${result.sheetName}\n` +
      `Portfolio Initiatives: ${result.portfolioInits.length}\n` +
      `Value Streams: ${result.valueStreams.length}\n` +
      `Program Initiatives: ${result.programInits.length}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Error', error.message, ui.ButtonSet.OK);
  }
}

/**
 * Test function to add slides to an existing presentation
 */
function testAddPortfolioDistributionSlides() {
  const ui = SpreadsheetApp.getUi();
  
  const piResponse = ui.prompt('Enter PI number:', ui.ButtonSet.OK_CANCEL);
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  const piNumber = parseInt(piResponse.getResponseText().trim());
  
  const presResponse = ui.prompt('Enter Presentation ID:', ui.ButtonSet.OK_CANCEL);
  if (presResponse.getSelectedButton() !== ui.Button.OK) return;
  const presentationId = presResponse.getResponseText().trim();
  
  try {
    const presentation = SlidesApp.openById(presentationId);
    const result = addPortfolioDistributionSlides(presentation, piNumber, 'ALL');
    
    if (result.success) {
      ui.alert('Success', 
        `Added ${result.slidesAdded} slides\n` +
        `Summary sheet: ${result.summarySheetName}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Error', result.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', error.message, ui.ButtonSet.OK);
  }
}
function createBlankSlide(presentation) {
  try {
    // Method 1: Try to find a blank or simple layout in the presentation's layouts
    const masters = presentation.getMasters();
    if (masters.length > 0) {
      const layouts = masters[0].getLayouts();
      
      // Look for a blank or minimal layout
      for (const layout of layouts) {
        const layoutName = layout.getLayoutName().toLowerCase();
        if (layoutName.includes('blank') || layoutName.includes('empty')) {
          return presentation.appendSlide(layout);
        }
      }
      
      // If no blank layout found, use the first layout and clear it
      if (layouts.length > 0) {
        const newSlide = presentation.appendSlide(layouts[0]);
        
        // Remove all placeholder elements
        const shapes = newSlide.getShapes();
        shapes.forEach(shape => {
          try {
            shape.remove();
          } catch (e) {
            // Some shapes may not be removable
          }
        });
        
        return newSlide;
      }
    }
    
    // Method 2: Duplicate an existing slide and clear it
    const slides = presentation.getSlides();
    if (slides.length > 0) {
      const newSlide = slides[0].duplicate();
      
      // Clear all content from the duplicated slide
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
      
      // Reset background to white
      newSlide.getBackground().setSolidFill('#FFFFFF');
      
      return newSlide;
    }
    
    // Method 3: Last resort - use any predefined layout
    return presentation.appendSlide();
    
  } catch (e) {
    console.error(`Error creating blank slide: ${e.message}`);
    // Final fallback - just append a slide with no arguments
    return presentation.appendSlide();
  }
}