function calculateBVAVTableHeight(bvAvData) {
  if (!bvAvData || !bvAvData.allocations || bvAvData.allocations.length === 0) {
    return 0;
  }
  
  let rowCount = 1;  // Header row
  
  bvAvData.allocations.forEach(allocGroup => {
    rowCount += 1;  // Allocation header row
    rowCount += allocGroup.initiatives.length;  // Initiative rows
  });
  
  rowCount += 1;  // TOTAL row
  rowCount += 1;  // PI Score row
  
  return rowCount;
}

function generateAllValueStreamChartsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const piResponse = ui.prompt('Generate Value Stream Charts', 
    'Enter PI number (e.g., 12):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  const piNumber = parseInt(piResponse.getResponseText().trim());
  
  if (isNaN(piNumber)) {
    ui.alert('Invalid input', 'Please provide valid PI number.', ui.ButtonSet.OK);
    return;
  }
  
  Logger.log(`\n${'='.repeat(80)}`);
  Logger.log(`GENERATING ALL VALUE STREAM CHARTS - PI ${piNumber} (DYNAMIC HEIGHTS)`);
  Logger.log(`${'='.repeat(80)}\n`);
  
  const predScoreSheet = ss.getSheetByName('Predictability Score');
  if (!predScoreSheet) {
    ui.alert('Error', 'Predictability Score sheet not found!', ui.ButtonSet.OK);
    return;
  }
  
  const iter6Sheet = ss.getSheetByName(`PI ${piNumber} - Iteration 6`);
  if (!iter6Sheet) {
    ui.alert('Error', `PI ${piNumber} - Iteration 6 sheet not found!`, ui.ButtonSet.OK);
    return;
  }
  
  const chartSheetName = `PI ${piNumber} - VS Charts`;
  let chartSheet = ss.getSheetByName(chartSheetName);
  if (chartSheet) {
    chartSheet.getCharts().forEach(chart => chartSheet.removeChart(chart));
    chartSheet.clear();
  } else {
    chartSheet = ss.insertSheet(chartSheetName);
  }
  
  const piOffset = piNumber - 11;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // HELPER FUNCTION - Define before use
  // ═══════════════════════════════════════════════════════════════════════════
  function colLetter(colNum) {
    let letter = '';
    while (colNum > 0) {
      const mod = (colNum - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      colNum = Math.floor((colNum - 1) / 26);
    }
    return letter;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 1: PRE-CALCULATE ALL BV/AV DATA AND SECTION HEIGHTS
  // ═══════════════════════════════════════════════════════════════════════════
  Logger.log(`\n📊 PHASE 1: Pre-calculating section heights...`);
  
  const MIN_SECTION_HEIGHT = 15;  // Minimum rows needed for charts to render properly
  const sectionData = [];  // Array of { valueStream, bvAvData, sectionHeight, startRow }
  
  let currentRow = 1;
  
  PREDICTABILITY_VALUE_STREAMS.forEach((valueStream, vsIndex) => {
    const bvAvData = getBVAVSummaryData(iter6Sheet, piNumber, valueStream, predScoreSheet, vsIndex, piOffset);
    const bvAvHeight = calculateBVAVTableHeight(bvAvData);
    
    // Section height = max of (BV/AV table + 2 for header/padding) OR minimum chart height
    // Add 2 extra rows: 1 for section header, 1 for bottom padding
    const neededHeight = bvAvHeight > 0 ? bvAvHeight + 2 : 0;
    const sectionHeight = Math.max(MIN_SECTION_HEIGHT, neededHeight);
    
    sectionData.push({
      valueStream: valueStream,
      vsIndex: vsIndex,
      bvAvData: bvAvData,
      bvAvHeight: bvAvHeight,
      sectionHeight: sectionHeight,
      headerRow: currentRow,
      dataStartRow: currentRow + 1
    });
    
    Logger.log(`  ${valueStream}: BV/AV rows=${bvAvHeight}, Section height=${sectionHeight}, Start row=${currentRow}`);
    
    currentRow += sectionHeight;
  });
  
  Logger.log(`\n📊 Total rows needed: ${currentRow - 1}`);
  
  // Calculate column letters for formulas
  const allocStartColNum = 2 + (piOffset * 5);
  const velocityColNum = 12 + piOffset;
  const predScoreColNum = 12 + piOffset;
  
  const allocCols = [
    colLetter(allocStartColNum),
    colLetter(allocStartColNum + 1),
    colLetter(allocStartColNum + 2),
    colLetter(allocStartColNum + 3),
    colLetter(allocStartColNum + 4)
  ];
  const velocityCol = colLetter(velocityColNum);
  const predScoreCol = colLetter(predScoreColNum);
  
  Logger.log(`Formula columns for PI ${piNumber}:`);
  Logger.log(`  Allocation: ${allocCols.join(', ')}`);
  Logger.log(`  Velocity/PredScore: ${velocityCol}`);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PHASE 2: GENERATE CHARTS FOR EACH VALUE STREAM
  // ═══════════════════════════════════════════════════════════════════════════
  Logger.log(`\n📊 PHASE 2: Generating charts...`);
  
  // Store section boundaries in Column AA (metadata for refresh function)
  chartSheet.getRange(1, 27).setValue('SECTION_METADATA');
  chartSheet.getRange(1, 27).setFontWeight('bold').setBackground('#cccccc');
  
  sectionData.forEach((section, idx) => {
    const { valueStream, vsIndex, bvAvData, sectionHeight, headerRow, dataStartRow } = section;
    const sectionEndRow = headerRow + sectionHeight - 1;
    
    Logger.log(`\n${'─'.repeat(60)}`);
    Logger.log(`${valueStream} (Rows ${headerRow}-${sectionEndRow}, Height: ${sectionHeight})`);
    Logger.log(`${'─'.repeat(60)}`);
    
    // Store section boundary metadata in Column AA
    // Format: "VS_INDEX|HEADER_ROW|DATA_START|SECTION_END"
    chartSheet.getRange(headerRow, 27).setValue(`${vsIndex}|${headerRow}|${dataStartRow}|${sectionEndRow}`);
    chartSheet.getRange(headerRow, 27).setFontColor('#999999').setFontSize(8);
    
    // ═══════════════════════════════════════════════════════════════════════════
    // SECTION HEADER
    // ═══════════════════════════════════════════════════════════════════════════
    const headerCell = chartSheet.getRange(headerRow, 1, 1, 26);
    headerCell.merge();
    headerCell.setValue(valueStream);
    headerCell.setFontSize(14);
    headerCell.setFontWeight('bold');
    headerCell.setBackground('#674ea7');
    headerCell.setFontColor('#ffffff');
    headerCell.setHorizontalAlignment('center');
    headerCell.setVerticalAlignment('middle');
    chartSheet.setRowHeight(headerRow, 35);
    
    const sectionStartRow = dataStartRow;
    
    // Row references in Predictability Score sheet
    const allocRow = 23 + vsIndex;
    const predScoreRow = 117 + vsIndex;
    
    // ═══════════════════════════════════════════════════════════════════
    // CHART 1: CAPACITY & VELOCITY BOXES (Cols 1-3) - FORMULA BASED
    // ═══════════════════════════════════════════════════════════════════
    Logger.log(`  Chart 1: Capacity & Velocity...`);
    
    const capacityRow = 4 + vsIndex;
    const capacityColLetter = colLetter(2 + ((piNumber - 11) * 2));  // PI 11→B, PI 12→D, PI 13→F
    
    const velocityRow = 41 + vsIndex;
    const velocityColLetter = colLetter(1 + piNumber);  // PI 11→L, PI 12→M, PI 13→N
    
    const BOX_HEIGHT = 2;
    const BOX_WIDTH = 2;
    const BOX_GAP = 1;
    
    // Write formulas to data cells (Column 3)
    chartSheet.getRange(sectionStartRow, 3).setFormula(`='Predictability Score'!${capacityColLetter}${capacityRow}`);
    chartSheet.getRange(sectionStartRow + BOX_HEIGHT + BOX_GAP, 3).setFormula(`='Predictability Score'!${velocityColLetter}${velocityRow}`);
    chartSheet.setColumnWidth(3, 50);
    
    Logger.log(`    Capacity formula: ='Predictability Score'!${capacityColLetter}${capacityRow}`);
    Logger.log(`    Velocity formula: ='Predictability Score'!${velocityColLetter}${velocityRow}`);

    // CAPACITY BOX - references formula cell
    const capacityBox = chartSheet.getRange(sectionStartRow, 1, BOX_HEIGHT, BOX_WIDTH);
    capacityBox.merge();
    capacityBox.setFormula(`=ROUND(C${sectionStartRow})&CHAR(10)&"Capacity"`);
    capacityBox.setBackground('#9575cd');
    capacityBox.setFontColor('#ffffff');
    capacityBox.setFontWeight('bold');
    capacityBox.setFontSize(14);
    capacityBox.setHorizontalAlignment('center');
    capacityBox.setVerticalAlignment('middle');
    capacityBox.setWrap(true);
    capacityBox.setBorder(true, true, true, true, false, false, '#FFC82E', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // VELOCITY BOX - references formula cell
    const velocityBox = chartSheet.getRange(sectionStartRow + BOX_HEIGHT + BOX_GAP, 1, BOX_HEIGHT, BOX_WIDTH);
    velocityBox.merge();
    velocityBox.setFormula(`=ROUND(C${sectionStartRow + BOX_HEIGHT + BOX_GAP})&CHAR(10)&"Velocity"`);
    velocityBox.setBackground('#b39ddb');
    velocityBox.setFontColor('#ffffff');
    velocityBox.setFontWeight('bold');
    velocityBox.setFontSize(14);
    velocityBox.setHorizontalAlignment('center');
    velocityBox.setVerticalAlignment('middle');
    velocityBox.setWrap(true);
    velocityBox.setBorder(true, true, true, true, false, false, '#FFC82E', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    Logger.log(`  ✅ Chart 1 created (FORMULA-BASED - auto-updates)`);
    
    // ═══════════════════════════════════════════════════════════════════
    // CHART 2: ALLOCATION PIE CHART (Cols 4-9) - FORMULA BASED
    // ═══════════════════════════════════════════════════════════════════
    const pieDataRow = sectionStartRow;
    
    // Header row (static)
    chartSheet.getRange(pieDataRow, 4).setValue('Category');
    chartSheet.getRange(pieDataRow, 5).setValue('Points');
    
    // Data rows with FORMULAS
    chartSheet.getRange(pieDataRow + 1, 4).setValue('Product-Feature');
    chartSheet.getRange(pieDataRow + 1, 5).setFormula(`='Predictability Score'!${allocCols[0]}${allocRow}`);
    
    chartSheet.getRange(pieDataRow + 2, 4).setValue('Compliance');
    chartSheet.getRange(pieDataRow + 2, 5).setFormula(`='Predictability Score'!${allocCols[1]}${allocRow}`);
    
    chartSheet.getRange(pieDataRow + 3, 4).setValue('Tech/Platform');
    chartSheet.getRange(pieDataRow + 3, 5).setFormula(`='Predictability Score'!${allocCols[4]}${allocRow}`);
    
    chartSheet.getRange(pieDataRow + 4, 4).setValue('Quality');
    chartSheet.getRange(pieDataRow + 4, 5).setFormula(`='Predictability Score'!${allocCols[3]}${allocRow}`);
    
    chartSheet.getRange(pieDataRow + 5, 4).setValue('KLO');
    chartSheet.getRange(pieDataRow + 5, 5).setFormula(`='Predictability Score'!${allocCols[2]}${allocRow}`);

    const pieChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(chartSheet.getRange(pieDataRow + 1, 4, 5, 2))
      .setPosition(pieDataRow, 4, 0, 0)
      .setOption('title', 'Capacity Allocation Delivered')
      .setOption('titleTextStyle', { fontSize: 11, bold: true, alignment: 'center' })
      .setOption('width', 320)
      .setOption('height', 240)
      .setOption('backgroundColor', 'transparent')
      .setOption('is3D', true)
      .setOption('pieSliceText', 'percentage')
      .setOption('pieSliceTextStyle', { color: '#ffffff' })
      .setOption('legend', { position: 'bottom', alignment: 'center', textStyle: { fontSize: 9 } })
      .setOption('fontName', 'Tahoma')
      .setOption('colors', ['#674ea7', '#9900ff', '#f1c232', '#46bdc6', '#cccccc'])
      .build();

    chartSheet.insertChart(pieChart);
    Logger.log(`  ✅ Chart 2 created (FORMULA-BASED - auto-updates)`);
    
    // ═══════════════════════════════════════════════════════════════════
    // CHART 3: EPIC STATUS BOXES (Cols 12-14) - FORMULA BASED
    // ═══════════════════════════════════════════════════════════════════
    Logger.log(`  Chart 3: Epic Status...`);
    
    // Epic Status is in Table 4, row 61 + vsIndex
    // PI 11: Columns B-D (indices 2-4), Committed at B, Deferred at C
    // PI 12: Columns E-G (indices 5-7), Committed at E, Deferred at F
    const epicStatusRow = 61 + vsIndex;
    const epicStatusStartCol = 2 + ((piNumber - 11) * 3);  // PI 11→2, PI 12→5, PI 13→8
    const committedColLetter = colLetter(epicStatusStartCol);      // Committed
    const deferredColLetter = colLetter(epicStatusStartCol + 1);   // Deferred
    
    // Write formulas to data cells (Column 14)
    chartSheet.getRange(sectionStartRow, 14).setFormula(`='Predictability Score'!${committedColLetter}${epicStatusRow}`);
    chartSheet.getRange(sectionStartRow + BOX_HEIGHT + BOX_GAP, 14).setFormula(`='Predictability Score'!${deferredColLetter}${epicStatusRow}`);
    chartSheet.setColumnWidth(14, 50);
    
    Logger.log(`    Completed formula: ='Predictability Score'!${committedColLetter}${epicStatusRow}`);
    Logger.log(`    Deferred formula: ='Predictability Score'!${deferredColLetter}${epicStatusRow}`);
    
    // COMPLETED BOX - references formula cell
    const completedBox = chartSheet.getRange(sectionStartRow, 12, BOX_HEIGHT, BOX_WIDTH);
    completedBox.merge();
    completedBox.setFormula(`=N${sectionStartRow}&CHAR(10)&"Features Completed"`);
    completedBox.setBackground('#7e57c2');
    completedBox.setFontColor('#ffffff');
    completedBox.setFontWeight('bold');
    completedBox.setFontSize(12);
    completedBox.setHorizontalAlignment('center');
    completedBox.setVerticalAlignment('middle');
    completedBox.setWrap(true);
    completedBox.setBorder(true, true, true, true, false, false, '#FFC82E', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // DEFERRED BOX - references formula cell
    const deferredBox = chartSheet.getRange(sectionStartRow + BOX_HEIGHT + BOX_GAP, 12, BOX_HEIGHT, BOX_WIDTH);
    deferredBox.merge();
    deferredBox.setFormula(`=N${sectionStartRow + BOX_HEIGHT + BOX_GAP}&CHAR(10)&"Features Deferred"`);
    deferredBox.setBackground('#9575cd');
    deferredBox.setFontColor('#ffffff');
    deferredBox.setFontWeight('bold');
    deferredBox.setFontSize(12);
    deferredBox.setHorizontalAlignment('center');
    deferredBox.setVerticalAlignment('middle');
    deferredBox.setWrap(true);
    deferredBox.setBorder(true, true, true, true, false, false, '#FFC82E', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    Logger.log(`  ✅ Chart 3 created (FORMULA-BASED - auto-updates)`);
    
    // ═══════════════════════════════════════════════════════════════════
    // CHART 4: VELOCITY TRENDING (Cols 15-16) - FORMULA BASED
    // ═══════════════════════════════════════════════════════════════════
    Logger.log(`  Chart 4: Velocity Trending...`);
    
    const velocityTrendRow = 41 + vsIndex;
    let velocityDataCount = 0;
    
    // Scan from column B (index 2) through column P (index 16) to find all data
    const VELOCITY_START_COL = 2;   // Column B
    const VELOCITY_END_COL = 16;    // Column P (PI 15)
    
    Logger.log(`    Scanning row ${velocityTrendRow}, columns B-P for velocity data...`);
    
    // First pass: find which columns have data
    const velocityDataCols = [];
    for (let col = VELOCITY_START_COL; col <= VELOCITY_END_COL; col++) {
      const value = predScoreSheet.getRange(velocityTrendRow, col).getValue();
      if (value && value > 0) {
        velocityDataCols.push({ col: col, pi: col - 1 });
      }
    }
    
    Logger.log(`    Found ${velocityDataCols.length} velocity data points`);
    
    if (velocityDataCols.length > 0) {
      const velChartDataRow = sectionStartRow;
      
      // Write formulas for each data point
      velocityDataCols.forEach((data, idx) => {
        const colLetterForFormula = colLetter(data.col);
        chartSheet.getRange(velChartDataRow + idx, 15).setValue(`PI ${data.pi}`);
        chartSheet.getRange(velChartDataRow + idx, 16).setFormula(`='Predictability Score'!${colLetterForFormula}${velocityTrendRow}`);
        Logger.log(`    PI ${data.pi} formula: ='Predictability Score'!${colLetterForFormula}${velocityTrendRow}`);
      });
      
      velocityDataCount = velocityDataCols.length;
      
      // Get current values for Y-axis calculation
      const velocityValues = velocityDataCols.map(data => 
        predScoreSheet.getRange(velocityTrendRow, data.col).getValue()
      );
      const minVelocity = Math.min(...velocityValues);
      const maxVelocity = Math.max(...velocityValues);
      const yAxisMin = Math.max(0, minVelocity - 200);
      const yAxisMax = maxVelocity + 200;
      
      Logger.log(`    Velocity range: ${minVelocity}-${maxVelocity}, Y-axis: ${yAxisMin}-${yAxisMax}`);
      
      const velChart = chartSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartSheet.getRange(velChartDataRow, 15, velocityDataCount, 2))
        .setPosition(velChartDataRow, 15, 0, 0)
        .setOption('title', `${valueStream} ART Velocity (Story Points)`)
        .setOption('titleTextStyle', { fontSize: 10, bold: true, alignment: 'center' })
        .setOption('width', 350)
        .setOption('height', 240)
        .setOption('backgroundColor', 'transparent')
        .setOption('curveType', 'function')
        .setOption('colors', ['#674ea7'])
        .setOption('pointSize', 5)
        .setOption('lineWidth', 3)
        .setOption('hAxis', { 
          title: 'Program Increment',
          textStyle: { fontSize: 8 }
        })
        .setOption('vAxis', { 
          title: 'Story Points',
          viewWindow: { min: yAxisMin, max: yAxisMax },
          gridlines: { count: -1 },
          textStyle: { fontSize: 8 }
        })
        .setOption('legend', { position: 'none' })
        .setOption('trendlines', { 
          0: { 
            color: '#fff2cc',
            lineWidth: 2,
            opacity: 0.5,
            type: 'linear',
            visibleInLegend: false
          }
        })
        .build();
      
      chartSheet.insertChart(velChart);
      Logger.log(`  ✅ Chart 4 created (FORMULA-BASED - auto-updates) with ${velocityDataCount} PIs`);
    }
    
    // ═══════════════════════════════════════════════════════════════════
    // CHART 5: PROGRAM PREDICTABILITY (Cols 21-22) - FORMULA BASED
    // ═══════════════════════════════════════════════════════════════════
    Logger.log(`  Chart 5: Program Predictability...`);

    // Read from Table 7 (PI Score Summary) - same structure as velocity (ONE column per PI)
    const TABLE7_DATA_START_ROW = 117;
    const predTrendRow = TABLE7_DATA_START_ROW + vsIndex;
    let predDataCount = 0;

    // Scan from column B (index 2) through column P (index 16) to find all data
    const PRED_START_COL = 2;   // Column B
    const PRED_END_COL = 16;    // Column P (PI 15)

    Logger.log(`    Scanning Table 7 row ${predTrendRow}, columns B-P for predictability data...`);

    // First pass: find which columns have data
    const predDataCols = [];
    for (let col = PRED_START_COL; col <= PRED_END_COL; col++) {
      const value = predScoreSheet.getRange(predTrendRow, col).getValue();
      if (value && value > 0) {
        predDataCols.push({ col: col, pi: col - 1 });
      }
    }

    Logger.log(`    Found ${predDataCols.length} predictability data points`);

    if (predDataCols.length > 0) {
      const predChartDataRow = sectionStartRow;
      
      // Write formulas for each data point
      predDataCols.forEach((data, idx) => {
        const colLetterForFormula = colLetter(data.col);
        chartSheet.getRange(predChartDataRow + idx, 21).setValue(`PI ${data.pi}`);
        // Multiply by 100 if value is decimal (< 2), otherwise use as-is
        chartSheet.getRange(predChartDataRow + idx, 22).setFormula(
          `=IF('Predictability Score'!${colLetterForFormula}${predTrendRow}<2,'Predictability Score'!${colLetterForFormula}${predTrendRow}*100,'Predictability Score'!${colLetterForFormula}${predTrendRow})`
        );
        Logger.log(`    PI ${data.pi} formula: ='Predictability Score'!${colLetterForFormula}${predTrendRow}`);
      });
      
      predDataCount = predDataCols.length;
      
      // Get current values for Y-axis calculation
      const predValues = predDataCols.map(data => {
        const val = predScoreSheet.getRange(predTrendRow, data.col).getValue();
        return val < 2 ? val * 100 : val;
      });
      const minPred = Math.min(...predValues);
      const maxPred = Math.max(...predValues);
      const yAxisMin = Math.max(0, minPred - 10);
      const yAxisMax = Math.min(120, maxPred + 10);
      
      Logger.log(`    Predictability range: ${minPred.toFixed(1)}-${maxPred.toFixed(1)}%, Y-axis: ${yAxisMin}-${yAxisMax}`);
      
      const predChart = chartSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartSheet.getRange(predChartDataRow, 21, predDataCount, 2))
        .setPosition(predChartDataRow, 21, 0, 0)
        .setOption('title', `${valueStream} Program Predictability Score`)
        .setOption('titleTextStyle', { fontSize: 10, bold: true, alignment: 'center' })
        .setOption('width', 350)
        .setOption('height', 240)
        .setOption('backgroundColor', 'transparent')
        .setOption('curveType', 'function')
        .setOption('colors', ['#674ea7'])
        .setOption('pointSize', 5)
        .setOption('lineWidth', 3)
        .setOption('hAxis', { 
          title: 'Program Increment',
          textStyle: { fontSize: 8 }
        })
        .setOption('vAxis', { 
          title: 'Score (%)',
          viewWindow: { min: yAxisMin, max: yAxisMax },
          gridlines: { count: -1 },
          textStyle: { fontSize: 8 }
        })
        .setOption('legend', { position: 'none' })
        .setOption('trendlines', { 
          0: { 
            color: '#fff2cc',
            lineWidth: 2,
            opacity: 0.5,
            type: 'linear',
            visibleInLegend: false
          }
        })
        .build();
      
      chartSheet.insertChart(predChart);
      Logger.log(`  ✅ Chart 5 created (FORMULA-BASED - auto-updates) with ${predDataCount} PIs`);
    }

    // ═══════════════════════════════════════════════════════════════════════════
    // CHART 6: BV/AV SUMMARY TABLE (Cols 24-26) - STATIC DATA, FORMULA PI SCORE
    // ═══════════════════════════════════════════════════════════════════════════
    Logger.log(`  Chart 6: BV/AV Summary...`);
    
    if (bvAvData) {
      let tableRow = sectionStartRow;
      
      chartSheet.setColumnWidth(24, 200);
      chartSheet.setColumnWidth(25, 48);
      chartSheet.setColumnWidth(26, 48);
      
      // Header row
      chartSheet.getRange(tableRow, 24).setValue('Program Initiative');
      chartSheet.getRange(tableRow, 25).setValue('BV');
      chartSheet.getRange(tableRow, 26).setValue('AV');
      chartSheet.getRange(tableRow, 24, 1, 3)
        .setFontWeight('bold')
        .setBackground('#674ea7')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      tableRow++;
      
      // Allocation groups with subtotals
      bvAvData.allocations.forEach(allocGroup => {
        // Allocation header row with SUBTOTALS
        chartSheet.getRange(tableRow, 24).setValue(allocGroup.allocation);
        chartSheet.getRange(tableRow, 25).setValue(allocGroup.subtotalBV);
        chartSheet.getRange(tableRow, 26).setValue(allocGroup.subtotalAV);
        chartSheet.getRange(tableRow, 24, 1, 3)
          .setBackground('#fff2cc')
          .setFontWeight('bold')
          .setHorizontalAlignment('center')
          .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        tableRow++;
        
        // Initiative rows (indented)
        allocGroup.initiatives.forEach(init => {
          chartSheet.getRange(tableRow, 24).setValue('   ' + init.name);
          chartSheet.getRange(tableRow, 25).setValue(init.bv);
          chartSheet.getRange(tableRow, 26).setValue(init.av);
          
          if (init.av < init.bv) {
            chartSheet.getRange(tableRow, 26).setBackground('#ffcccc');
          }
          
          chartSheet.getRange(tableRow, 24, 1, 3)
            .setHorizontalAlignment('center')
            .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
          tableRow++;
        });
      });
      
      // TOTAL row (sum of allocation subtotals)
      chartSheet.getRange(tableRow, 24).setValue('TOTAL');
      chartSheet.getRange(tableRow, 25).setValue(bvAvData.totalBV);
      chartSheet.getRange(tableRow, 26).setValue(bvAvData.totalAV);
      chartSheet.getRange(tableRow, 24, 1, 3)
        .setBackground('#add8e6')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      tableRow++;
      
      // PI Score row - FORMULA BASED from Table 5
      const TABLE5_DATA_START_ROW = 80;
      const TABLE5_PI_11_START_COL = 2;
      const TABLE5_COLS_PER_PI = 3;
      const TABLE5_PI_SCORE_OFFSET = 2;
      const piIndex = piNumber - 11;
      const piScoreCol = TABLE5_PI_11_START_COL + (piIndex * TABLE5_COLS_PER_PI) + TABLE5_PI_SCORE_OFFSET;
      const piScoreRowNum = TABLE5_DATA_START_ROW + vsIndex;
      const piScoreColLetter = colLetter(piScoreCol);
      
      chartSheet.getRange(tableRow, 24).setValue('Program Predictability Score');
      chartSheet.getRange(tableRow, 24).setFontWeight('bold').setBackground('#fff2cc');
      chartSheet.getRange(tableRow, 25, 1, 2).merge();
      
      // Formula that handles decimal vs percentage and formats with %
      chartSheet.getRange(tableRow, 25).setFormula(
        `=IF('Predictability Score'!${piScoreColLetter}${piScoreRowNum}<2,TEXT('Predictability Score'!${piScoreColLetter}${piScoreRowNum}*100,"0.0")&"%",TEXT('Predictability Score'!${piScoreColLetter}${piScoreRowNum},"0.0")&"%")`
      );
      
      // Conditional formatting for score color (green >= 90, red < 90)
      const currentScore = bvAvData.piScore;
      const scoreColor = currentScore >= 90 ? '#d4edda' : '#ffcccc';
      chartSheet.getRange(tableRow, 25).setBackground(scoreColor).setFontWeight('bold');
      chartSheet.getRange(tableRow, 24, 1, 3)
        .setHorizontalAlignment('center')
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      
      Logger.log(`    PI Score formula: ='Predictability Score'!${piScoreColLetter}${piScoreRowNum}`);
      Logger.log(`  ✅ Chart 6 created (PI Score FORMULA-BASED - auto-updates)`);
    }
  });  // End of sectionData.forEach
  
  chartSheet.setColumnWidths(1, 3, 70);
  
  // Hide metadata column
  chartSheet.hideColumns(27);
  
  Logger.log(`\n✅ COMPLETE - Dynamic section heights used`);
  
  ui.alert('Charts Generated!', 
    `Created all 6 charts for ${PREDICTABILITY_VALUE_STREAMS.length} value streams.\n\n` +
    `📊 DYNAMIC HEIGHTS: Each section sized to fit its BV/AV data.\n\n` +
    `📊 AUTO-REFRESH (formula-based):\n` +
    `• Chart 1: Capacity & Velocity boxes\n` +
    `• Chart 2: Allocation Pie\n` +
    `• Chart 3: Epic Status boxes\n` +
    `• Chart 4: Velocity Trending\n` +
    `• Chart 5: Predictability Trending\n` +
    `• Chart 6: PI Score (in BV/AV table)\n\n` +
    `Note: BV/AV initiative data is static and requires regeneration to update.`,
    ui.ButtonSet.OK);
}


function checkPredictabilityAlignment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const predSheet = ss.getSheetByName('Predictability Score');
  
  if (!predSheet) {
    SpreadsheetApp.getUi().alert('Error', 'Predictability Score sheet not found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  console.log('═'.repeat(80));
  console.log('PREDICTABILITY VALUE STREAMS ALIGNMENT CHECK');
  console.log('═'.repeat(80));
  console.log('');
  
  const START_ROW = 117;
  const END_ROW = 131;
  const VALUE_STREAM_COL = 1;
  
  console.log('Reading from Predictability Score sheet:');
  console.log('-'.repeat(80));
  
  const csvValueStreams = [];
  for (let row = START_ROW; row <= END_ROW; row++) {
    const vsName = predSheet.getRange(row, VALUE_STREAM_COL).getValue();
    if (vsName && vsName.toString().trim()) {
      csvValueStreams.push(vsName.toString().trim());
      console.log(`Row ${row}: ${vsName}`);
    }
  }
  
  console.log('');
  console.log('═'.repeat(80));
  console.log('COMPARISON: PREDICTABILITY_VALUE_STREAMS vs ACTUAL CSV');
  console.log('═'.repeat(80));
  console.log('');
  
  let allMatch = true;
  const maxLen = Math.max(PREDICTABILITY_VALUE_STREAMS.length, csvValueStreams.length);
  
  console.log('Index | PREDICTABILITY_VALUE_STREAMS | CSV Row | CSV Name                     | Match');
  console.log('-'.repeat(100));
  
  for (let i = 0; i < maxLen; i++) {
    const codeVS = i < PREDICTABILITY_VALUE_STREAMS.length ? PREDICTABILITY_VALUE_STREAMS[i] : '---';
    const csvVS = i < csvValueStreams.length ? csvValueStreams[i] : '---';
    const csvRow = START_ROW + i;
    const match = codeVS === csvVS;
    const matchIcon = match ? '✅' : '❌';
    
    if (!match) allMatch = false;
    
    const codeVSPadded = (codeVS + ' '.repeat(30)).substring(0, 28);
    const csvVSPadded = (csvVS + ' '.repeat(30)).substring(0, 28);
    
    console.log(`${i.toString().padStart(5)} | ${codeVSPadded} | ${csvRow.toString().padStart(7)} | ${csvVSPadded} | ${matchIcon}`);
  }
  
  console.log('');
  console.log('═'.repeat(80));
  console.log('RESULT');
  console.log('═'.repeat(80));
  
  if (allMatch) {
    console.log('✅ SUCCESS: All value streams are correctly aligned!');
    SpreadsheetApp.getUi().alert('✅ Success', 
      'All value streams are correctly aligned with the Predictability Score sheet!', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    console.log('❌ ERROR: Misalignment detected!');
    console.log('');
    console.log('Please ensure PREDICTABILITY_VALUE_STREAMS constant matches the exact');
    console.log('order of value streams in the Predictability Score sheet (rows 117-131).');
    
    SpreadsheetApp.getUi().alert('❌ Error', 
      'Misalignment detected! Check console logs for details.\n\n' +
      'PREDICTABILITY_VALUE_STREAMS constant does not match the Predictability Score sheet.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  console.log('');
}
function refreshFormattedBoxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const piResponse = ui.prompt('Refresh Formatted Boxes', 
    'Enter PI number (e.g., 12):', 
    ui.ButtonSet.OK_CANCEL);
  
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  const piNumber = parseInt(piResponse.getResponseText().trim());
  
  if (isNaN(piNumber)) {
    ui.alert('Invalid input', 'Please provide valid PI number.', ui.ButtonSet.OK);
    return;
  }
  
  const chartSheetName = `PI ${piNumber} - VS Charts`;
  const chartSheet = ss.getSheetByName(chartSheetName);
  
  if (!chartSheet) {
    ui.alert('Error', `Sheet "${chartSheetName}" not found!`, ui.ButtonSet.OK);
    return;
  }
  
  const predScoreSheet = ss.getSheetByName('Predictability Score');
  if (!predScoreSheet) {
    ui.alert('Error', 'Predictability Score sheet not found!', ui.ButtonSet.OK);
    return;
  }
  
  const iter6Sheet = ss.getSheetByName(`PI ${piNumber} - Iteration 6`);
  if (!iter6Sheet) {
    ui.alert('Error', `PI ${piNumber} - Iteration 6 sheet not found!`, ui.ButtonSet.OK);
    return;
  }
  
  Logger.log(`\n${'='.repeat(80)}`);
  Logger.log(`REFRESHING FORMATTED BOXES - PI ${piNumber} (DYNAMIC BOUNDARIES)`);
  Logger.log(`${'='.repeat(80)}\n`);
  
  const piOffset = piNumber - 11;
  const BOX_HEIGHT = 2;
  const BOX_WIDTH = 2;
  const BOX_GAP = 1;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Read section boundaries from Column AA metadata
  // ═══════════════════════════════════════════════════════════════════════════
  const metadataCol = 27;  // Column AA
  const sectionBoundaries = [];
  
  // Scan column AA for metadata entries
  const lastRow = chartSheet.getLastRow();
  const metadataRange = chartSheet.getRange(1, metadataCol, lastRow, 1).getValues();
  
  for (let i = 0; i < metadataRange.length; i++) {
    const cellValue = metadataRange[i][0];
    if (cellValue && cellValue.toString().includes('|')) {
      const parts = cellValue.toString().split('|');
      if (parts.length === 4) {
        sectionBoundaries.push({
          vsIndex: parseInt(parts[0]),
          headerRow: parseInt(parts[1]),
          dataStartRow: parseInt(parts[2]),
          sectionEndRow: parseInt(parts[3])
        });
      }
    }
  }
  
  // Fallback to fixed calculation if no metadata found (backward compatibility)
  if (sectionBoundaries.length === 0) {
    Logger.log(`⚠️ No section metadata found - using fixed SECTION_HEIGHT=25 (legacy mode)`);
    const SECTION_HEIGHT = 25;
    PREDICTABILITY_VALUE_STREAMS.forEach((vs, idx) => {
      const headerRow = 1 + (idx * SECTION_HEIGHT);
      sectionBoundaries.push({
        vsIndex: idx,
        headerRow: headerRow,
        dataStartRow: headerRow + 1,
        sectionEndRow: headerRow + SECTION_HEIGHT - 1
      });
    });
  } else {
    Logger.log(`✅ Found ${sectionBoundaries.length} section boundaries from metadata`);
  }
  
  // Sort by vsIndex to ensure correct order
  sectionBoundaries.sort((a, b) => a.vsIndex - b.vsIndex);
  
  sectionBoundaries.forEach((section, idx) => {
    const { vsIndex, headerRow, dataStartRow, sectionEndRow } = section;
    const valueStream = PREDICTABILITY_VALUE_STREAMS[vsIndex];
    
    if (!valueStream) {
      Logger.log(`⚠️ Skipping invalid vsIndex: ${vsIndex}`);
      return;
    }
    
    const sectionStartRow = dataStartRow;
    
    Logger.log(`\n${valueStream} (Rows ${headerRow}-${sectionEndRow})`);
    
    // ─────────────────────────────────────────────────────────────────
    // CHART 1: CAPACITY & VELOCITY BOXES
    // ─────────────────────────────────────────────────────────────────
    const capacityRowIdx = 4 + vsIndex;
    const capacityColIdx = 2 + ((piNumber - 11) * 2);
    const velocityRowIdx = 41 + vsIndex;
    const velocityColIdx = 1 + piNumber;
    
    const capacity = parseFloat(predScoreSheet.getRange(capacityRowIdx, capacityColIdx).getValue()) || 0;
    const velocity = parseFloat(predScoreSheet.getRange(velocityRowIdx, velocityColIdx).getValue()) || 0;
    
    // Update Capacity Box
    const capacityBox = chartSheet.getRange(sectionStartRow, 1, BOX_HEIGHT, BOX_WIDTH);
    const capacityText = SpreadsheetApp.newRichTextValue()
      .setText(`${Math.round(capacity)}\nCapacity`)
      .setTextStyle(0, Math.round(capacity).toString().length, 
        SpreadsheetApp.newTextStyle().setFontSize(18).setForegroundColor('#ffffff').setBold(true).build())
      .setTextStyle(Math.round(capacity).toString().length, `${Math.round(capacity)}\nCapacity`.length,
        SpreadsheetApp.newTextStyle().setFontSize(11).setForegroundColor('#ffffff').setBold(true).build())
      .build();
    capacityBox.setRichTextValue(capacityText);
    
    // Update Velocity Box
    const velocityBox = chartSheet.getRange(sectionStartRow + BOX_HEIGHT + BOX_GAP, 1, BOX_HEIGHT, BOX_WIDTH);
    const velocityText = SpreadsheetApp.newRichTextValue()
      .setText(`${Math.round(velocity)}\nVelocity`)
      .setTextStyle(0, Math.round(velocity).toString().length, 
        SpreadsheetApp.newTextStyle().setFontSize(18).setForegroundColor('#ffffff').setBold(true).build())
      .setTextStyle(Math.round(velocity).toString().length, `${Math.round(velocity)}\nVelocity`.length,
        SpreadsheetApp.newTextStyle().setFontSize(11).setForegroundColor('#ffffff').setBold(true).build())
      .build();
    velocityBox.setRichTextValue(velocityText);
    
    Logger.log(`  ✓ Chart 1: Capacity=${capacity}, Velocity=${velocity}`);
    
    // ─────────────────────────────────────────────────────────────────
    // CHART 3: EPIC STATUS BOXES
    // ─────────────────────────────────────────────────────────────────
    const iter6Data = iter6Sheet.getDataRange().getValues();
    const iter6Headers = iter6Data[3];
    
    const iter6VsCol = iter6Headers.indexOf('Value Stream/Org');
    const iter6StatusCol = iter6Headers.indexOf('Status');
    const iter6TypeCol = iter6Headers.indexOf('Issue Type');
    
    let completedCount = 0;
    let deferredCount = 0;
    
    for (let i = 4; i < iter6Data.length; i++) {
      const row = iter6Data[i];
      const rowVS = (row[iter6VsCol] || '').toString().trim().toLowerCase();
      const targetVS = valueStream.toLowerCase();
      const status = (row[iter6StatusCol] || '').toString();
      const type = (row[iter6TypeCol] || '').toString();
      
      if (rowVS === targetVS && type === 'Epic') {
        if (status === 'Closed') {
          completedCount++;
        } else if (status === 'Deferred') {
          deferredCount++;
        }
      }
    }
    
    // Update Committed Box
    const committedBox = chartSheet.getRange(sectionStartRow, 12, BOX_HEIGHT, BOX_WIDTH);
    const committedText = SpreadsheetApp.newRichTextValue()
      .setText(`${completedCount}\nFeatures Completed`)
      .setTextStyle(0, completedCount.toString().length, 
        SpreadsheetApp.newTextStyle().setFontSize(15).setForegroundColor('#ffffff').setBold(true).build())
      .setTextStyle(completedCount.toString().length, `${completedCount}\nFeatures Completed`.length,
        SpreadsheetApp.newTextStyle().setFontSize(11).setForegroundColor('#ffffff').setBold(true).build())
      .build();
    committedBox.setRichTextValue(committedText);
    
    // Update Deferred Box
    const deferredBox = chartSheet.getRange(sectionStartRow + BOX_HEIGHT + BOX_GAP, 12, BOX_HEIGHT, BOX_WIDTH);
    const deferredText = SpreadsheetApp.newRichTextValue()
      .setText(`${deferredCount}\nFeatures Deferred`)
      .setTextStyle(0, deferredCount.toString().length, 
        SpreadsheetApp.newTextStyle().setFontSize(15).setForegroundColor('#ffffff').setBold(true).build())
      .setTextStyle(deferredCount.toString().length, `${deferredCount}\nFeatures Deferred`.length,
        SpreadsheetApp.newTextStyle().setFontSize(11).setForegroundColor('#ffffff').setBold(true).build())
      .build();
    deferredBox.setRichTextValue(deferredText);
    
    Logger.log(`  ✓ Chart 3: Completed=${completedCount}, Deferred=${deferredCount}`);

    // ─────────────────────────────────────────────────────────────────
    // CHART 6: BV/AV SUMMARY TABLE - DYNAMIC HEIGHT
    // ─────────────────────────────────────────────────────────────────
    const bvAvData = getBVAVSummaryData(iter6Sheet, piNumber, valueStream, predScoreSheet, vsIndex, piOffset);
    
    if (bvAvData && bvAvData.allocations && bvAvData.allocations.length > 0) {
      // Calculate available rows for BV/AV table
      const maxBvAvRows = sectionEndRow - sectionStartRow;
      
      Logger.log(`  Clearing BV/AV area: rows ${sectionStartRow}-${sectionEndRow}, cols 24-26 (${maxBvAvRows} rows available)`);
      
      // Break apart merged ranges first
      const bvAvArea = chartSheet.getRange(sectionStartRow, 24, maxBvAvRows, 3);
      const mergedRanges = bvAvArea.getMergedRanges();
      Logger.log(`  Found ${mergedRanges.length} merged ranges to break apart`);
      mergedRanges.forEach(mergedRange => {
        mergedRange.breakApart();
      });
      
      // Clear the range
      bvAvArea.clearContent();
      bvAvArea.clearFormat();
      
      let tableRow = sectionStartRow;
      
      // Set column widths
      chartSheet.setColumnWidth(24, 200);
      chartSheet.setColumnWidth(25, 100);
      chartSheet.setColumnWidth(26, 100);
      
      // Rebuild table header
      chartSheet.getRange(tableRow, 24).setValue('Program Initiative');
      chartSheet.getRange(tableRow, 25).setValue('BV');
      chartSheet.getRange(tableRow, 26).setValue('AV');
      chartSheet.getRange(tableRow, 24, 1, 3)
        .setFontWeight('bold')
        .setBackground('#674ea7')
        .setFontColor('#ffffff')
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      tableRow++;
      
      // Rebuild allocations and initiatives - NO TRUNCATION with dynamic heights
      bvAvData.allocations.forEach(allocGroup => {
        // Allocation row with SUBTOTALS
        chartSheet.getRange(tableRow, 24).setValue(allocGroup.allocation);
        chartSheet.getRange(tableRow, 25).setValue(allocGroup.subtotalBV);
        chartSheet.getRange(tableRow, 26).setValue(allocGroup.subtotalAV);
        chartSheet.getRange(tableRow, 24, 1, 3)
          .setBackground('#fff2cc')
          .setFontWeight('bold')
          .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        tableRow++;
        
        // Initiative rows (indented)
        allocGroup.initiatives.forEach(init => {
          chartSheet.getRange(tableRow, 24).setValue('   ' + init.name);
          chartSheet.getRange(tableRow, 25).setValue(init.bv);
          chartSheet.getRange(tableRow, 26).setValue(init.av);
          
          if (init.av < init.bv) {
            chartSheet.getRange(tableRow, 26).setBackground('#ffcccc');
          }
          
          chartSheet.getRange(tableRow, 24, 1, 3)
            .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
          tableRow++;
        });
      });
      
      // Total row
      chartSheet.getRange(tableRow, 24).setValue('TOTAL');
      chartSheet.getRange(tableRow, 25).setValue(bvAvData.totalBV);
      chartSheet.getRange(tableRow, 26).setValue(bvAvData.totalAV);
      chartSheet.getRange(tableRow, 24, 1, 3)
        .setBackground('#add8e6')
        .setFontWeight('bold')
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      tableRow++;
      
      // PI Score row
      chartSheet.getRange(tableRow, 24).setValue('Program Predictability Score');
      chartSheet.getRange(tableRow, 24).setFontWeight('bold').setBackground('#fff2cc');
      chartSheet.getRange(tableRow, 25, 1, 2).merge();
      chartSheet.getRange(tableRow, 25).setValue(`${bvAvData.piScore.toFixed(1)}%`);
      const scoreColor = bvAvData.piScore >= 90 ? '#d4edda' : '#ffcccc';
      chartSheet.getRange(tableRow, 25).setBackground(scoreColor).setFontWeight('bold');
      chartSheet.getRange(tableRow, 24, 1, 3)
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      
      const rowsUsed = tableRow - sectionStartRow + 1;
      Logger.log(`  ✓ Chart 6: ${rowsUsed} rows used, BV=${bvAvData.totalBV}, AV=${bvAvData.totalAV}, Score=${bvAvData.piScore.toFixed(1)}%`);
    } else {
      Logger.log(`  ⚠️ Chart 6: No BV/AV data found`);
    }
  });
  
  Logger.log(`\n✅ FORMATTED BOXES REFRESH COMPLETE`);
  
  ui.alert('Success!', 
    `Refreshed Charts 1, 3, 6 for all ${PREDICTABILITY_VALUE_STREAMS.length} value streams.\n\n` +
    `• Chart 1: Capacity & Velocity boxes\n` +
    `• Chart 3: Completed & Deferred boxes\n` +
    `• Chart 6: BV/AV Summary tables (dynamic heights)`,
    ui.ButtonSet.OK);
}
function getBVAVSummaryData(iter6Sheet, piNumber, valueStream, predScoreSheet, vsIndex, piOffset) {
  Logger.log(`    Reading BV/AV data from Iteration 6 sheet...`);
  Logger.log(`    Looking for Value Stream: "${valueStream}" (vsIndex: ${vsIndex})`);
  
  // Ensure valueStream is a string
  const targetVS = (valueStream || '').toString().toLowerCase().trim();
  
  if (!targetVS) {
    Logger.log(`    ❌ Invalid value stream: "${valueStream}"`);
    return null;
  }
  
  const data = iter6Sheet.getDataRange().getValues();
  const headers = data[3];
  
  const vsCol = headers.indexOf('Value Stream/Org');
  const allocCol = headers.indexOf('Allocation');
  const statusCol = headers.indexOf('Status');
  const typeCol = headers.indexOf('Issue Type');
  const piObjStatusCol = headers.indexOf('PI Objective Status');
  const bvCol = headers.indexOf('Business Value');
  const avCol = headers.indexOf('Actual Value');
  const programInitCol = headers.indexOf('Program Initiative');
  const scrumTeamCol = headers.indexOf('Scrum Team');  // ADDED: For MMPM filtering
  
  if (vsCol === -1 || allocCol === -1 || statusCol === -1 || typeCol === -1 || 
      piObjStatusCol === -1 || bvCol === -1 || avCol === -1 || programInitCol === -1) {
    Logger.log(`    ❌ Required columns not found`);
    Logger.log(`    vsCol=${vsCol}, allocCol=${allocCol}, statusCol=${statusCol}, typeCol=${typeCol}`);
    Logger.log(`    piObjStatusCol=${piObjStatusCol}, bvCol=${bvCol}, avCol=${avCol}, programInitCol=${programInitCol}`);
    return null;
  }
  
  // Check if MMPM - special filtering applies
  const isMMPM = targetVS === 'mmpm';
  if (isMMPM) {
    Logger.log(`    ⚠️ MMPM detected - will exclude Scrum Team = "Appeal Engine"`);
    if (scrumTeamCol === -1) {
      Logger.log(`    ⚠️ Warning: Scrum Team column not found, cannot apply MMPM filter`);
    }
  }
  
  const allocationsToInclude = ['Product - Feature', 'Product - Compliance', 'Tech / Platform'];
  const grouped = {};
  let bvRowsFound = 0;
  let avRowsFound = 0;
  let vsMatchCount = 0;
  let appealEngineSkipped = 0;
  
  for (let i = 4; i < data.length; i++) {
    const row = data[i];
    const rowVS = (row[vsCol] || '').toString().trim();
    
    // CASE-INSENSITIVE MATCHING
    const vsMatch = rowVS.toLowerCase() === targetVS;
    
    if (vsMatch) vsMatchCount++;
    
    if (!vsMatch) continue;
    if (row[typeCol] !== 'Epic') continue;
    
    // ═══════════════════════════════════════════════════════════════════
    // MMPM SPECIAL FILTER: Exclude Scrum Team = "Appeal Engine"
    // ═══════════════════════════════════════════════════════════════════
    if (isMMPM && scrumTeamCol !== -1) {
      const scrumTeam = (row[scrumTeamCol] || '').toString().trim();
      if (scrumTeam === 'Appeal Engine') {
        appealEngineSkipped++;
        Logger.log(`    ✖ Skipping Appeal Engine epic: ${row[headers.indexOf('Key')] || 'unknown'}`);
        continue;
      }
    }
    
    const allocation = row[allocCol];
    
    // Skip Quality and KLO allocations (excluded from both BV and AV)
    if (allocation === 'Quality' || allocation === 'KLO' || allocation === 'KLO / Quality') {
      continue;
    }
    
    // Only include specific allocations
    if (!allocationsToInclude.includes(allocation)) {
      Logger.log(`    Skipping allocation: ${allocation}`);
      continue;
    }
    
    const programInit = row[programInitCol] || 'Unassigned';
    const bv = parseFloat(row[bvCol]) || 0;
    const av = parseFloat(row[avCol]) || 0;
    const status = (row[statusCol] || '').toString().toLowerCase();
    const piObjStatus = (row[piObjStatusCol] || '').toString().toLowerCase();
    
    // Initialize grouped structure
    if (!grouped[allocation]) {
      grouped[allocation] = {};
    }
    if (!grouped[allocation][programInit]) {
      grouped[allocation][programInit] = { bv: 0, av: 0 };
    }
    
    // ═══════════════════════════════════════════════════════════════════
    // BV: ALL epics (any status), excluding Quality/KLO
    // ═══════════════════════════════════════════════════════════════════
    grouped[allocation][programInit].bv += bv;
    bvRowsFound++;
    
    // ═══════════════════════════════════════════════════════════════════
    // AV: Only Closed/Pending Acceptance epics with PI Objective Status = "Met"
    // ═══════════════════════════════════════════════════════════════════
    const isClosedOrPending = (status === 'closed' || status === 'pending acceptance');
    const isMet = (piObjStatus === 'met');
    
    if (isClosedOrPending && isMet) {
      grouped[allocation][programInit].av += av;
      avRowsFound++;
      Logger.log(`    ✓ AV counted: ${programInit} - BV=${bv}, AV=${av}, Status=${status}, PIObjStatus=${piObjStatus}`);
    } else {
      Logger.log(`    ○ BV only: ${programInit} - BV=${bv}, Status=${status}, PIObjStatus=${piObjStatus}`);
    }
  }
  
  Logger.log(`    VS matches: ${vsMatchCount}, BV rows: ${bvRowsFound}, AV rows: ${avRowsFound} for "${valueStream}"`);
  if (isMMPM) {
    Logger.log(`    Appeal Engine epics skipped: ${appealEngineSkipped}`);
  }
  
  // Build allocations array with subtotals
  const allocations = [];
  let totalBV = 0;
  let totalAV = 0;
  
  allocationsToInclude.forEach(alloc => {
    if (grouped[alloc]) {
      const initiatives = [];
      let allocBV = 0;
      let allocAV = 0;
      
      Object.keys(grouped[alloc]).forEach(init => {
        const initBV = Math.round(grouped[alloc][init].bv);
        const initAV = Math.round(grouped[alloc][init].av);
        
        initiatives.push({
          name: init,
          bv: initBV,
          av: initAV
        });
        
        // Sum for allocation subtotal
        allocBV += initBV;
        allocAV += initAV;
      });
      
      allocations.push({
        allocation: alloc,
        initiatives: initiatives,
        subtotalBV: allocBV,
        subtotalAV: allocAV
      });
      
      // Sum allocation subtotals for grand total
      totalBV += allocBV;
      totalAV += allocAV;
    }
  });
  
  if (totalBV === 0 && totalAV === 0) {
    Logger.log(`    ⚠️ No BV/AV data found for ${valueStream} - returning null`);
    return null;
  }
  
  // ═══════════════════════════════════════════════════════════════════
  // READ PI SCORE FROM TABLE 5 (PROGRAM PREDICTABILITY) - NOT CALCULATED
  // ═══════════════════════════════════════════════════════════════════
  let piScore = 0;
  
  if (predScoreSheet) {
    const TABLE5_DATA_START_ROW = 80;
    const TABLE5_PI_11_START_COL = 2;  // Column B
    const TABLE5_COLS_PER_PI = 3;      // BV, AV, PI Score
    const TABLE5_PI_SCORE_OFFSET = 2;  // PI Score is 3rd column (offset 2)
    
    const piIndex = piNumber - 11;
    const piScoreCol = TABLE5_PI_11_START_COL + (piIndex * TABLE5_COLS_PER_PI) + TABLE5_PI_SCORE_OFFSET;
    const piScoreRow = TABLE5_DATA_START_ROW + vsIndex;
    
    const rawPiScore = predScoreSheet.getRange(piScoreRow, piScoreCol).getValue();
    
    // Convert decimal to percentage if needed (values < 2 are likely decimals)
    if (rawPiScore && rawPiScore > 0) {
      piScore = rawPiScore < 2 ? rawPiScore * 100 : rawPiScore;
    }
    
    Logger.log(`    Read PI Score from Table 5: Row ${piScoreRow}, Col ${piScoreCol} = ${rawPiScore} → ${piScore.toFixed(1)}%`);
  } else {
    // Fallback: calculate if predScoreSheet not available
    piScore = totalBV > 0 ? (totalAV / totalBV) * 100 : 0;
    Logger.log(`    ⚠️ predScoreSheet not available, calculated PI Score: ${piScore.toFixed(1)}%`);
  }
  
  Logger.log(`    Total BV=${totalBV}, Total AV=${totalAV}`);
  Logger.log(`    PI Score (from Table 5): ${piScore.toFixed(1)}%`);
  Logger.log(`    ✓ Returning ${allocations.length} allocations`);
  
  return {
    allocations: allocations,
    totalBV: totalBV,
    totalAV: totalAV,
    piScore: piScore
  };
}


/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * DIAGNOSTIC: Check Value Stream names in Iteration 6 sheet
 * ═══════════════════════════════════════════════════════════════════════════════
 * Run this to see what value stream names actually exist in your data
 */
function diagnoseBVAVValueStreams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Get PI number
  const piResponse = ui.prompt('Diagnose BV/AV Value Streams', 'Enter PI number:', ui.ButtonSet.OK_CANCEL);
  if (piResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const piNumber = parseInt(piResponse.getResponseText());
  const iter6Sheet = ss.getSheetByName(`PI ${piNumber} - Iteration 6`);
  
  if (!iter6Sheet) {
    ui.alert('Error', `Sheet "PI ${piNumber} - Iteration 6" not found`, ui.ButtonSet.OK);
    return;
  }
  
  const data = iter6Sheet.getDataRange().getValues();
  const headers = data[3];
  const vsCol = headers.indexOf('Value Stream/Org');
  
  if (vsCol === -1) {
    ui.alert('Error', 'Column "Value Stream/Org" not found in headers', ui.ButtonSet.OK);
    return;
  }
  
  // Collect unique value streams
  const valueStreamCounts = {};
  for (let i = 4; i < data.length; i++) {
    const vs = (data[i][vsCol] || '').toString().trim();
    if (vs) {
      valueStreamCounts[vs] = (valueStreamCounts[vs] || 0) + 1;
    }
  }
  
  // Build report
  let report = 'VALUE STREAMS FOUND IN ITERATION 6 DATA:\n';
  report += '=' .repeat(50) + '\n\n';
  
  const sortedVS = Object.keys(valueStreamCounts).sort();
  sortedVS.forEach(vs => {
    report += `"${vs}" - ${valueStreamCounts[vs]} rows\n`;
  });
  
  report += '\n' + '=' .repeat(50) + '\n';
  report += 'EXPECTED VALUE STREAMS:\n\n';
  
  PREDICTABILITY_VALUE_STREAMS.forEach(vs => {
    const found = sortedVS.find(s => s.toLowerCase() === vs.toLowerCase());
    const match = found ? `✓ Found as "${found}"` : '❌ NOT FOUND';
    report += `"${vs}" - ${match}\n`;
  });
  
  console.log(report);
  
  // Show in dialog
  const htmlOutput = HtmlService.createHtmlOutput(
    `<pre style="font-family: monospace; font-size: 12px; white-space: pre-wrap;">${report}</pre>`
  ).setWidth(600).setHeight(500);
  
  ui.showModalDialog(htmlOutput, 'Value Stream Diagnosis');
}
