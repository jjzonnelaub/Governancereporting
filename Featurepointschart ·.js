/**
 * FINAL FIXED VERSION
 * - No labels for allocations with 0% FP
 * - No labels for segments that are too narrow (< 30px)
 * - Labels positioned higher above bar
 */

function addFeaturePointsChartToSlide2(presentation, valueStream, programIncrement, sheetName) {
  console.log('\n🔍 Starting chart generation');
  
  try {
    const slides = presentation.getSlides();
    if (slides.length < 2) {
      console.error('❌ ERROR: Presentation does not have slide 2');
      return;
    }
    
    const slide2 = slides[1];
    
    // Find {{BAR CHART}} placeholder
    const shapes = slide2.getShapes();
    let chartX = 50;
    let chartY = 100;
    let chartWidth = 880;
    let chartHeight = 50;
    
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      try {
        const textRange = shape.getText();
        if (textRange && !textRange.isEmpty()) {
          const text = textRange.asString();
          
          if (text.includes('{{BAR CHART}}')) {
            chartX = shape.getLeft();
            chartY = shape.getTop();
            chartWidth = shape.getWidth();
            chartHeight = shape.getHeight();
            console.log(`✅ Found placeholder at (${chartX}, ${chartY})`);
            shape.remove();
            break;
          }
        }
      } catch (e) {
        // Skip shapes without text
      }
    }
    
    // Get data - ONLY returns allocations with FP > 0
    const aggregatedData = aggregateFeaturePointsByAllocation(sheetName, valueStream);
    
    if (!aggregatedData || aggregatedData.length === 0) {
      console.error('❌ No data returned!');
      return;
    }
    
    console.log(`✅ Got ${aggregatedData.length} allocations (0% excluded)`);
    
    // Draw chart
    drawHybridChart(slide2, aggregatedData, chartX, chartY, chartWidth, chartHeight);
    
    console.log('✅ Chart complete!\n');
    
  } catch (error) {
    console.error(`❌ ERROR: ${error.message}`);
    throw error;
  }
}

function aggregateFeaturePointsByAllocation(sheetName, valueStream) {
  try {
    console.log(`\n🔍 CHART AGGREGATION DEBUG`);
    console.log(`Sheet name: "${sheetName}"`);
    console.log(`Value Stream: "${valueStream}"`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error(`❌ Sheet not found: ${sheetName}`);
      console.log(`Available sheets: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    console.log(`Total rows in sheet: ${data.length}`);
    
    // Find header row
    let headerRowIndex = -1;
    let headers = [];
    
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const row = data[i];
      const rowText = row.join('|').toLowerCase();
      
      if (rowText.includes('value stream') && rowText.includes('allocation')) {
        headerRowIndex = i;
        headers = data[i];
        console.log(`✅ Found header row at index ${i}`);
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      console.error('❌ Header row not found!');
      console.log('First 3 rows:');
      for (let i = 0; i < Math.min(3, data.length); i++) {
        console.log(`Row ${i}: ${data[i].slice(0, 5).join(' | ')}`);
      }
      return [];
    }
    
    const vsCol = headers.indexOf('Value Stream/Org');
    const allocationCol = headers.indexOf('Allocation');
    const fpCol = headers.indexOf('Feature Points');
    
    console.log(`Column indices - VS: ${vsCol}, Allocation: ${allocationCol}, FP: ${fpCol}`);
    
    if (vsCol === -1 || allocationCol === -1 || fpCol === -1) {
      console.error('❌ Required columns not found!');
      console.log(`Headers: ${headers.slice(0, 10).join(' | ')}`);
      return [];
    }
    
    // Aggregate - only if FP > 0
    const aggregation = {};
    let totalRowsChecked = 0;
    let rowsMatchingVS = 0;
    let rowsWithFP = 0;
    
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      const rowVS = row[vsCol];
      const allocation = row[allocationCol];
      const fpValue = row[fpCol];
      
      totalRowsChecked++;
      
      if (rowVS === valueStream) {
        rowsMatchingVS++;
        
        const fp = (fpValue !== null && fpValue !== undefined && fpValue !== '') 
          ? parseFloat(fpValue) 
          : 0;
        
        // ONLY aggregate if FP > 0
        if (allocation && allocation !== '' && !isNaN(fp) && fp > 0) {
          rowsWithFP++;
          if (!aggregation[allocation]) {
            aggregation[allocation] = 0;
          }
          aggregation[allocation] += fp;
          
          if (rowsWithFP <= 5) {
            console.log(`  Row ${i}: ${allocation} = ${fp} FP`);
          }
        }
      }
    }
    
    console.log(`\n📊 AGGREGATION RESULTS:`);
    console.log(`  Total rows checked: ${totalRowsChecked}`);
    console.log(`  Rows matching VS "${valueStream}": ${rowsMatchingVS}`);
    console.log(`  Rows with FP > 0: ${rowsWithFP}`);
    
    const total = Object.values(aggregation).reduce((sum, val) => sum + val, 0);
    
    if (total === 0) {
      console.error('❌ Total FP = 0!');
      console.log('Aggregation object:', JSON.stringify(aggregation));
      return [];
    }
    
    console.log(`  Total FP: ${total}`);
    
    // Colors - Purple/Gold/Grey palette for professional appearance
    const colorMap = {
      'Product - Feature': '#6C207F',      // Purple (dark)
      'Product - Compliance': '#8B5CA8',   // Purple (medium)
      'Compliance': '#A77BC2',             // Purple (light)
      'Infosec': '#C4A8D8',                // Purple (lighter)
      'Tech / Platform': '#FFC42E',        // Gold (primary)
      'KLO': '#757575',                    // Grey (medium)
      'KLO / Quality': '#757575',          // Grey (medium)
      'Quality': '#9E9E9E'                 // Grey (light)
    };
    
    // ONLY return allocations with FP > 0
    const result = Object.entries(aggregation)
      .filter(([allocation, featurePoints]) => featurePoints > 0)
      .map(([allocation, featurePoints]) => ({
        allocation: allocation,
        featurePoints: featurePoints,
        percentage: (featurePoints / total) * 100,
        color: colorMap[allocation] || '#CCCCCC'
      }));
    
    result.sort((a, b) => b.percentage - a.percentage);
    
    console.log('Aggregated data (only FP > 0):');
    result.forEach(item => {
      console.log(`  ${item.allocation}: ${item.featurePoints} FP (${item.percentage.toFixed(1)}%)`);
    });
    
    return result;
    
  } catch (error) {
    console.error(`Aggregation error: ${error.message}`);
    return [];
  }
}

function drawHybridChart(slide, data, x, y, width, height) {
  console.log('\n🎨 Drawing chart...');
  console.log(`Drawing ${data.length} allocations (0% excluded)`);
  
  try {
    // Reduce bar width for labels
    const SPACE_RIGHT = 60;
    const SPACE_LEFT = 40;
    const barWidth = width - SPACE_RIGHT - SPACE_LEFT;
    const barX = x + SPACE_LEFT;
    const barY = y;
    
    // Labels above bar
    const LABEL_HEIGHT = 12;
    const LABEL_GAP = 6;
    const LABEL_Y = y - LABEL_HEIGHT - LABEL_GAP;
    
    // FIX: Minimum width to show label above
    const MIN_WIDTH_FOR_LABEL = 30;  // Don't show label if segment < 30px
    
    console.log(`Bar: ${barWidth.toFixed(1)}x${height.toFixed(1)} at (${barX.toFixed(1)}, ${barY.toFixed(1)})`);
    console.log(`Labels at Y=${LABEL_Y.toFixed(1)} (${LABEL_GAP}px gap)`);
    console.log(`Min width for labels: ${MIN_WIDTH_FOR_LABEL}px`);
    
    // REMOVED: 0% and 100% labels completely removed per user request
    
    // Draw segments
    let currentX = barX;
    
    data.forEach((item, idx) => {
      const segmentWidth = (item.percentage / 100) * barWidth;
      
      // Skip if too small to render
      if (segmentWidth < 0.5) {
        console.log(`[${idx}] SKIP "${item.allocation}" - too small (${segmentWidth.toFixed(2)}px)`);
        return;
      }
      
      console.log(`\n[${idx}] Drawing "${item.allocation}": ${item.percentage.toFixed(1)}% (${segmentWidth.toFixed(1)}px)`);
      
      // Draw colored segment
      const segment = slide.insertShape(
        SlidesApp.ShapeType.RECTANGLE,
        currentX,
        barY,
        segmentWidth,
        height
      );
      segment.getFill().setSolidFill(item.color);
      segment.getBorder().setTransparent();
      console.log(`  ✓ Segment drawn (color: ${item.color})`);
      
      // Only show label above if segment is wide enough
      if (segmentWidth >= MIN_WIDTH_FOR_LABEL) {
        const labelAbove = slide.insertTextBox(
          item.allocation,
          currentX,
          LABEL_Y,
          segmentWidth,
          LABEL_HEIGHT
        );
        labelAbove.getText().getTextStyle()
          .setFontFamily('Lato')
          .setFontSize(7)
          .setBold(true)
          .setForegroundColor('#000000');
        labelAbove.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        console.log(`  ✓ Label above: "${item.allocation}"`);
      } else {
        console.log(`  ⚠️ Segment too narrow for label (${segmentWidth.toFixed(1)}px < ${MIN_WIDTH_FOR_LABEL}px) - SKIPPED`);
      }
      
      // Percentage inside bar (if wide enough)
      if (segmentWidth >= 40) {
        const percentInside = slide.insertTextBox(
          `${item.percentage.toFixed(1)}%`,
          currentX,
          barY + (height / 2) - 7,
          segmentWidth,
          14
        );
        
        const textStyle = percentInside.getText().getTextStyle();
        textStyle
          .setFontFamily('Lato')
          .setFontSize(10)
          .setBold(true);
        
        // White text on dark backgrounds
        const isDark = (item.color === '#6C207F' || item.color === '#1769FC' || item.color === '#3F68DD');
        textStyle.setForegroundColor(isDark ? '#FFFFFF' : '#000000');
        
        percentInside.getText().getParagraphStyle()
          .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        
        console.log(`  ✓ Percentage inside: ${item.percentage.toFixed(1)}%`);
      } else {
        console.log(`  ⚠️ Too narrow for inside % (${segmentWidth.toFixed(1)}px < 40px)`);
      }
      
      currentX += segmentWidth;
    });
    
    console.log(`\n✅ Chart complete! Drew ${data.length} segments`);
    
  } catch (error) {
    console.error(`Drawing error: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}
function testFeaturePointsChart() {
  try {
    const presentationId = 'YOUR_PRESENTATION_ID';
    const presentation = SlidesApp.openById(presentationId);
    const valueStream = 'MMPM';
    const sheetName = 'PI 13 - Iteration 0';
    
    addFeaturePointsChartToSlide2(presentation, valueStream, 'PI 13', sheetName);
    
    console.log('✅ Test complete!');
  } catch (error) {
    console.error(`❌ Test failed: ${error.message}`);
  }
}
function addFeaturePointsChartToSlide2ForInitiative(presentation, initiativeField, initiativeName, programIncrement, sheetName) {
  console.log('\n🔍 Starting Initiative chart generation');
  console.log(`Initiative Field: ${initiativeField}`);
  console.log(`Initiative Name: ${initiativeName}`);
  console.log(`PI: ${programIncrement}`);
  console.log(`Sheet: ${sheetName}`);
  
  try {
    const slides = presentation.getSlides();
    if (slides.length < 2) {
      console.error('❌ ERROR: Presentation does not have slide 2');
      return;
    }
    
    const slide2 = slides[1];
    
    // Find {{BAR CHART}} placeholder
    const shapes = slide2.getShapes();
    let chartX = 50;
    let chartY = 100;
    let chartWidth = 880;
    let chartHeight = 50;
    let foundPlaceholder = false;
    
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      try {
        const textRange = shape.getText();
        if (textRange && !textRange.isEmpty()) {
          const text = textRange.asString();
          
          if (text.includes('{{BAR CHART}}')) {
            chartX = shape.getLeft();
            chartY = shape.getTop();
            chartWidth = shape.getWidth();
            chartHeight = shape.getHeight();
            console.log(`✅ Found {{BAR CHART}} placeholder at (${chartX}, ${chartY}) size ${chartWidth}x${chartHeight}`);
            shape.remove();
            foundPlaceholder = true;
            break;
          }
        }
      } catch (e) {
        // Skip shapes without text
      }
    }
    
    if (!foundPlaceholder) {
      console.warn('⚠️ {{BAR CHART}} placeholder not found on slide 2');
    }
    
    // Get data - aggregated by Value Stream for this initiative
    const aggregatedData = aggregateFeaturePointsByValueStreamForInitiative(
      sheetName, 
      initiativeField, 
      initiativeName, 
      programIncrement
    );
    
    if (!aggregatedData || aggregatedData.length === 0) {
      console.error('❌ No data returned for initiative!');
      console.error(`   Initiative Field: ${initiativeField}`);
      console.error(`   Initiative Name: ${initiativeName}`);
      console.error(`   PI: ${programIncrement}`);
      console.error(`   Sheet: ${sheetName}`);
      return;
    }
    
    console.log(`✅ Got ${aggregatedData.length} value streams (0% excluded)`);
    
    // Draw chart using same function as value stream reports
    drawHybridChart(slide2, aggregatedData, chartX, chartY, chartWidth, chartHeight);
    
    console.log('✅ Initiative chart complete!\n');
    
  } catch (error) {
    console.error(`❌ ERROR in addFeaturePointsChartToSlide2ForInitiative: ${error.message}`);
    console.error(error.stack);
    throw error;
  }
}

/**
 * Aggregate Feature Points by Value Stream for a specific Initiative
 * Only includes Product-Feature allocation type
 */
function aggregateFeaturePointsByValueStreamForInitiative(sheetName, initiativeField, initiativeName, programIncrement) {
  try {
    console.log(`\n📊 Aggregating data for initiative chart:`);
    console.log(`   Sheet: ${sheetName}`);
    console.log(`   Initiative Field: ${initiativeField}`);
    console.log(`   Initiative Name: ${initiativeName}`);
    console.log(`   PI: ${programIncrement}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error(`❌ Sheet not found: ${sheetName}`);
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    console.log(`   Total rows in sheet: ${data.length}`);
    
    // Find header row
    let headerRowIndex = -1;
    let headers = [];
    
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const row = data[i];
      const rowText = row.join('|').toLowerCase();
      
      if (rowText.includes('value stream') && rowText.includes('allocation')) {
        headerRowIndex = i;
        headers = data[i];
        console.log(`✓ Found header row at index ${i}`);
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      console.error('❌ Header row not found!');
      return [];
    }
    
    // Find required columns - USE initiativeField parameter
    const vsCol = headers.indexOf('Value Stream/Org');
    const initiativeCol = headers.indexOf(initiativeField);  // <-- KEY FIX: Use parameter instead of hardcoded 'Initiative'
    const allocationCol = headers.indexOf('Allocation');
    const piCol = headers.indexOf('Program Increment');
    const fpCol = headers.indexOf('Feature Points');
    
    console.log(`\n📍 Column indices:`);
    console.log(`   Value Stream/Org: ${vsCol}`);
    console.log(`   ${initiativeField}: ${initiativeCol}`);
    console.log(`   Allocation: ${allocationCol}`);
    console.log(`   Program Increment: ${piCol}`);
    console.log(`   Feature Points: ${fpCol}`);
    
    if (vsCol === -1 || initiativeCol === -1 || allocationCol === -1 || piCol === -1 || fpCol === -1) {
      console.error('❌ Required columns not found!');
      if (vsCol === -1) console.error('   Missing: Value Stream/Org');
      if (initiativeCol === -1) console.error(`   Missing: ${initiativeField}`);
      if (allocationCol === -1) console.error('   Missing: Allocation');
      if (piCol === -1) console.error('   Missing: Program Increment');
      if (fpCol === -1) console.error('   Missing: Feature Points');
      return [];
    }
    
    // Aggregate by Value Stream (only Product-Feature allocation)
    const aggregation = {};
    let rowsChecked = 0;
    let rowsMatched = 0;
    
    console.log(`\n🔍 Scanning rows for matches...`);
    console.log(`   Looking for: ${initiativeField}="${initiativeName}", Allocation="Product - Feature", PI="${programIncrement}"`);
    
    for (let i = headerRowIndex + 1; i < data.length; i++) {
      const row = data[i];
      const rowVS = row[vsCol];
      const rowInitiative = row[initiativeCol];
      const rowAllocation = row[allocationCol];
      const rowPI = row[piCol];
      const fpValue = row[fpCol];
      
      rowsChecked++;
      
      const fp = (fpValue !== null && fpValue !== undefined && fpValue !== '') 
        ? parseFloat(fpValue) 
        : 0;
      
      // Log first few rows for debugging
      if (rowsChecked <= 3) {
        console.log(`\n   Row ${i}:`);
        console.log(`      ${initiativeField}: "${rowInitiative}"`);
        console.log(`      Allocation: "${rowAllocation}"`);
        console.log(`      PI: "${rowPI}"`);
        console.log(`      Value Stream: "${rowVS}"`);
        console.log(`      Feature Points: ${fp}`);
      }
      
      // Filter: same initiative, Product-Feature allocation only, current PI, FP > 0
      if (rowInitiative === initiativeName && 
          rowAllocation === 'Product - Feature' && 
          rowPI === programIncrement &&
          rowVS && rowVS !== '' && 
          !isNaN(fp) && fp > 0) {
        
        if (!aggregation[rowVS]) {
          aggregation[rowVS] = 0;
        }
        aggregation[rowVS] += fp;
        rowsMatched++;
        
        console.log(`   ✓ Match #${rowsMatched}: ${rowVS} - ${fp} FP (Row ${i})`);
      }
    }
    
    console.log(`\n📈 Scan complete:`);
    console.log(`   Rows checked: ${rowsChecked}`);
    console.log(`   Rows matched: ${rowsMatched}`);
    
    const total = Object.values(aggregation).reduce((sum, val) => sum + val, 0);
    
    if (total === 0) {
      console.error('❌ Total = 0! No Product-Feature epics found for this initiative.');
      console.error('   Possible causes:');
      console.error(`   1. No epics with ${initiativeField} = "${initiativeName}"`);
      console.error(`   2. No epics with Allocation = "Product - Feature"`);
      console.error(`   3. No epics with Program Increment = "${programIncrement}"`);
      console.error('   4. All matching epics have 0 Feature Points');
      return [];
    }
    
    console.log(`✓ Total FP for initiative "${initiativeName}": ${total}`);
    
    // Value Stream colors - UPDATE THESE TO MATCH YOUR VALUE STREAMS
    const colorMap = {
      'MMPM': '#6C207F',                    // Purple
      'EMA Clinical': '#1769FC',            // Blue
      'RCM Genie': '#FFC42E',              // Yellow/Gold
      'Clinical Quality': '#3F68DD',        // Light Blue
      'Data Governance': '#A8DADC',         // Teal
      'Enterprise': '#FFE2C8',             // Peach
      'Patient Collaboration': '#FF6B9D',   // Pink (NEW)
      'MMGI': '#2ECC71'                    // Green (NEW)
    };
    
    // Return data in same format as existing function
    const result = Object.entries(aggregation)
      .filter(([valueStream, featurePoints]) => featurePoints > 0)
      .map(([valueStream, featurePoints]) => ({
        allocation: valueStream,  // Using 'allocation' field name for compatibility with drawHybridChart
        featurePoints: featurePoints,
        percentage: (featurePoints / total) * 100,
        color: colorMap[valueStream] || '#CCCCCC'
      }));
    
    result.sort((a, b) => b.percentage - a.percentage);
    
    console.log('\n✅ Aggregated data by Value Stream (only FP > 0):');
    result.forEach(item => {
      console.log(`   ${item.allocation}: ${item.featurePoints} FP (${item.percentage.toFixed(1)}%)`);
    });
    
    return result;
    
  } catch (error) {
    console.error(`❌ Aggregation error: ${error.message}`);
    console.error(error.stack);
    return [];
  }
}
