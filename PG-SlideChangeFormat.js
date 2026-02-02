// =================================================================================
// NEW SCRIPT FOR "CHANGES-ONLY" PRESENTATIONS (ITERATION 2+)
// =================================================================================

const ICON_CONFIG = {
  NEW: '1HS_X47dlqhmYBhLt168oNmvFKwuWjebl',
  CHANGED: '1lcwYNb2yfafsCBGJzuPBNJHLa3Xd74Au',
  COMPLETED: '1ll-aS1qEpaAHtETNtkTyEbCqAgiBADLf',
  DEFERRED: '1_93UrZLtoIAyUVohMHqfSdHCuDXSDNW7',
  
  SIZE: 20, // The display size of the icons in points
  SPACING: 4  // Space between icons if multiple are shown
};

function createFormattedPresentation(presentationName, data, showDependencies = true, noInitiativeMode = 'show') {
  // noInitiativeMode options:
  // 'show' (default) - Show "No Initiative" highlighted
  // 'hide' - Hide epics with no initiative entirely
  // 'skip' - Skip initiative header but show value streams/epics directly
  try {
    // Store generation timestamp for slides
    const generatedTimestamp = new Date();
    
    // Step 1: Copy the template
    showProgress('Step 1 of 6: Copying template...', '📊 Starting');
    const templateFile = DriveApp.getFileById(TEMPLATE_CONFIG.TEMPLATE_ID);
    const copiedFile = templateFile.makeCopy(presentationName);
    const presentation = SlidesApp.openById(copiedFile.getId());
    console.log(`Created presentation from template: ${presentationName} (ID: ${presentation.getId()})`);
    
    // Step 2: Update title slide
    showProgress('Step 2 of 6: Updating title slide...', '📊 Preparing');
    updateTitleSlide(presentation, data.metadata);
    
    // Step 3: Process data
    showProgress('Step 3 of 6: Processing epic and dependency data...', '📊 Processing Data');
    const hideNoInitiative = noInitiativeMode === 'hide';
    const portfolioData = processDataForTableSlides(data, showDependencies, hideNoInitiative);
    
    // Step 4: Build slides for each portfolio using template
    const sortedPortfolios = getFilteredAndSortedPortfolios(Object.keys(portfolioData));
    const totalPortfolios = sortedPortfolios.length;
    console.log(`Starting to build slides for ${totalPortfolios} portfolios...`);
    
    // Get template slides (before we start duplicating)
    const dividerTemplate = presentation.getSlides()[TEMPLATE_CONFIG.PORTFOLIO_DIVIDER_INDEX];
    const tableTemplate = presentation.getSlides()[TEMPLATE_CONFIG.TABLE_SLIDE_INDEX];
    
    // Track where to insert new slides (after End of Governance slide position)
    // We'll insert before the End slide, then move End slide to the end
    let insertPosition = TEMPLATE_CONFIG.END_SLIDE_INDEX; // Start inserting at position 3 (0-indexed)
    
    sortedPortfolios.forEach((portfolioName, index) => {
      const progress = Math.round(((index + 1) / totalPortfolios) * 100);
      showProgress(
        `Step 4 of 6: Building portfolio ${index + 1} of ${totalPortfolios} (${progress}%)\n${portfolioName}`,
        '📊 Building Slides'
      );
      console.log(`[${index + 1}/${totalPortfolios}] Processing Portfolio: ${portfolioName}`);
      
      const slidesCreated = createPortfolioSlidesFromTemplate(
        presentation, 
        dividerTemplate, 
        tableTemplate, 
        portfolioName, 
        portfolioData[portfolioName], 
        data.dependencies, 
        showDependencies,
        insertPosition,
        generatedTimestamp,
        noInitiativeMode
      );
      
      insertPosition += slidesCreated;
    });
    
    // ═══════════════════════════════════════════════════════════════════════
    // Step 5: ADD PORTFOLIO DISTRIBUTION SLIDES (Feature Point Summary)
    // ═══════════════════════════════════════════════════════════════════════
    try {
      showProgress('Step 5 of 6: Adding Portfolio Distribution slides...', '📊 Building Charts');
      console.log('\n' + '═'.repeat(60));
      console.log('ADDING PORTFOLIO DISTRIBUTION SLIDES');
      console.log('═'.repeat(60));
      
      const piNumber = data.metadata?.piNumber;
      if (piNumber) {
        const portfolioResult = addPortfolioDistributionSlides(presentation, piNumber, 'ALL');
        
        if (portfolioResult && portfolioResult.success) {
          console.log(`✅ Portfolio Distribution: Added ${portfolioResult.slidesAdded} slides`);
          console.log(`   Summary sheet: ${portfolioResult.summarySheetName}`);
        } else {
          console.warn(`⚠️ Portfolio Distribution: ${portfolioResult ? portfolioResult.error : 'Unknown error'}`);
        }
      } else {
        console.warn('⚠️ Could not determine PI number for Portfolio Distribution slides');
      }
    } catch (portfolioError) {
      console.warn(`⚠️ Could not add Portfolio Distribution slides: ${portfolioError.message}`);
      // Non-fatal error - continue with presentation generation
    }
    
    // Step 6: Clean up - delete template slides and format guidelines
    showProgress('Step 6 of 6: Finalizing presentation...', '📊 Finishing Up');
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

/**
 * Filters the main data object to only include epics and dependencies that have changes.
 */
function filterDataByChanges(data, changelogData, showDependencies = true) {
  const changedKeys = new Set(Object.keys(changelogData));
  const epicsWithChanges = data.epics.filter(epic => changedKeys.has(epic['Key']));
  const includedEpicKeys = new Set(epicsWithChanges.map(e => e['Key']));
  
  // Only include dependencies if showDependencies is true
  let dependenciesWithChanges = [];
  if (showDependencies) {
    dependenciesWithChanges = data.dependencies.filter(dep => 
      changedKeys.has(dep['Key']) || includedEpicKeys.has(dep.parentEpicKey)
    );
  }
  
  return { ...data, epics: epicsWithChanges, dependencies: dependenciesWithChanges };
}

/**
 * A new version of createPortfolioSlides that passes changelog data down.
 */
function createPortfolioSlidesWithIcons(presentation, portfolioName, portfolioData, logoBlob, changelogData, iterationNumber, showDependencies = true) {
  if (!portfolioData.hasProgramData) return;

  // This reuses the same slideManager and pagination logic from your baseline report
  const totalPages = calculatePortfolioPageCount(portfolioData); // Reuse existing calculator
  let pageNum = 1;

  let slideManager = {
    presentation, portfolioName, slide: null, y: SLIDE_CONFIG.TABLE_START_Y, rows: 0,
    startNewSlide: function(isContinued = false) {
      if (isContinued) pageNum++;
      this.slide = this.presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      applyTemplateLayout(this.slide, false);
      addLogoToSlide(this.slide, logoBlob);
      addDisclosureToSlide(this.slide);
      
      // Add page number to bottom right if multiple pages
      if (totalPages > 1) {
        addPageNumberToSlide(this.slide, pageNum, totalPages);
      }
      
      // Title without page numbers (they're now in bottom right)
      addSlideHeader(this.slide, this.portfolioName, '');
      const bottomMargin = 70;
      const contentBoxHeight = SLIDE_CONFIG.SLIDE_HEIGHT - SLIDE_CONFIG.TABLE_START_Y - bottomMargin;
      this.slide.insertShape(SlidesApp.ShapeType.RECTANGLE, SLIDE_CONFIG.MARGIN_LEFT, SLIDE_CONFIG.TABLE_START_Y, SLIDE_CONFIG.CONTENT_WIDTH, contentBoxHeight)
          .getFill().setTransparent();
      this.y = SLIDE_CONFIG.TABLE_START_Y;
      this.rows = 0;
    }
  };

  slideManager.startNewSlide(false);
  const programs = Object.keys(portfolioData.programs).sort();
  for (let i = 0; i < programs.length; i++) {
    const programName = programs[i];
    const programData = portfolioData.programs[programName];
    if (i > 0) {
        if (slideManager.rows + 1 > SLIDE_CONFIG.MAX_ROWS_PER_SLIDE) { slideManager.startNewSlide(true); } 
        else { slideManager.y += 15; slideManager.rows++; }
    }
    addProgramHeaderToSlide(slideManager.slide, programName, slideManager.y);
    slideManager.y += SLIDE_CONFIG.SECTION_HEIGHT;
    slideManager.rows++;
    const sortedInitiatives = Object.keys(programData.initiatives).sort();
    for (const initiativeName of sortedInitiatives) {
      const initiativeData = programData.initiatives[initiativeName];
      let headerBlockDrawnForInitiative = false;
      const sortedStreams = Object.keys(initiativeData.valueStreams).sort();
      for (const streamName of sortedStreams) {
        const epics = initiativeData.valueStreams[streamName];
        for (const epic of epics) {
          let requiredRows = getEpicRowCount(epic); // Reuse existing row counter
          if (!headerBlockDrawnForInitiative) { requiredRows += (initiativeData.name !== 'No Initiative') ? 3 : 2; }
          if (slideManager.rows > 0 && (slideManager.rows + requiredRows) > SLIDE_CONFIG.MAX_ROWS_PER_SLIDE) {
            slideManager.startNewSlide(true);
            addProgramHeaderToSlide(slideManager.slide, `${programName} (Continued)`, slideManager.y);
            slideManager.y += SLIDE_CONFIG.SECTION_HEIGHT;
            slideManager.rows++;
            headerBlockDrawnForInitiative = false;
          }
          if (!headerBlockDrawnForInitiative) {
            // Pass changelogData and iterationNumber to addInitiativeToSlide
            slideManager.y = addInitiativeToSlide(slideManager.slide, initiativeData, slideManager.y, epic, changelogData, iterationNumber);
            slideManager.y = addValueStreamToSlide(slideManager.slide, { name: streamName }, slideManager.y);
            slideManager.rows += requiredRows - getEpicRowCount(epic);
            headerBlockDrawnForInitiative = true;
          }
          // **Call the new function to draw epics WITH icons**
          // Pass showDependencies to control whether dependencies are rendered
          slideManager.y = addEpicToSlideWithIcons(slideManager.slide, epic, slideManager.y, changelogData, iterationNumber, showDependencies);
          slideManager.rows += getEpicRowCount(epic);
        }
      }
    }
  }
}


function addEpicToSlideWithIcons(slide, epic, startY, changelogData, iterationNumber) {
  let currentY = startY;
  // Fixed indent for all epics (accounts for max 2 icons width)
  const BASE_INDENT = SLIDE_CONFIG.MARGIN_LEFT + 30;
  const MAX_ICON_WIDTH = (ICON_CONFIG.SIZE * 2) + (ICON_CONFIG.SPACING * 2) + 10;
  const TEXT_INDENT = BASE_INDENT + MAX_ICON_WIDTH; // All text starts here regardless of icon count
  
  const epicChanges = changelogData[epic['Key']] || { details: {} };
  
  // --- ICON LOGIC ---
  const iconsToDraw = [];
  
  // NEW: Item was added in this iteration
  if (epicChanges.isNew) {
    iconsToDraw.push(ICON_CONFIG.NEW);
  }
  
  // CHANGED: Has any field changes (but not new)
  if (!epicChanges.isNew && Object.keys(epicChanges.details).length > 0) {
    iconsToDraw.push(ICON_CONFIG.CHANGED);
  }
  
  // COMPLETED: Status changed TO 'Closed' or 'Done'
  if (epicChanges.details['Status'] && epicChanges.details['Status'].new) {
    const newStatus = epicChanges.details['Status'].new.toLowerCase();
    if (newStatus.includes('closed') || newStatus.includes('done')) {
      iconsToDraw.push(ICON_CONFIG.COMPLETED);
    }
  }
  
  // DEFERRED: PI Commitment changed TO 'Deferred'
  if (epicChanges.details['PI Commitment'] && epicChanges.details['PI Commitment'].new) {
    const newCommitment = epicChanges.details['PI Commitment'].new.toLowerCase();
    if (newCommitment.includes('deferred')) {
      iconsToDraw.push(ICON_CONFIG.DEFERRED);
    }
  }
  
  // Store icon images to send to back
  const iconImages = [];
  
  // Draw icons with horizontal staggering pattern
  const STAGGER_OFFSET = ICON_CONFIG.SIZE + ICON_CONFIG.SPACING;
  iconsToDraw.forEach((iconId, index) => {
    try {
      const iconBlob = DriveApp.getFileById(iconId).getBlob();
      
      // Horizontal staggering: alternate between left position and indented position
      let iconX;
      if (index % 2 === 0) {
        // Even index (0, 2, 4...): normal position
        iconX = BASE_INDENT;
      } else {
        // Odd index (1, 3, 5...): indented to the right
        iconX = BASE_INDENT + STAGGER_OFFSET;
      }
      
      // Vertical positioning: stack them with slight overlap
      const iconY = currentY + (SLIDE_CONFIG.EPIC_HEIGHT - ICON_CONFIG.SIZE) / 2 + (Math.floor(index / 2) * (ICON_CONFIG.SIZE * 0.7));
      
      const iconImage = slide.insertImage(iconBlob, iconX, iconY, ICON_CONFIG.SIZE, ICON_CONFIG.SIZE);
      iconImages.push(iconImage);
    } catch (e) { console.error(`Could not add icon with ID: ${iconId}`); }
  });
  
  // Send all icons to back
  iconImages.forEach(img => img.sendToBack());
  
  // --- TEXT LOGIC ---
  const epicData = processEpicData(epic);
  // ADD THESE LINES:
  if (epicData.summaryIsEmpty) {
  return currentY; // Skip epics with no title when user excludes them
  }
  const iterChanges = epicChanges.details['Iteration End'] || epicChanges.details['End Iteration Name'] || epicChanges.details['PI Target Iteration'];
  const hasIteration = (iterChanges ? iterChanges.new : epicData.endIteration)?.trim();
  
  // Use configuration-aware function for iteration text
  let iterationText = getIterationDisplayText(hasIteration);
  
  // If null is returned, user wants to hide empty iterations
  // In this case, we might still want to show it if there's a CHANGE
  // But if both old and new are empty, hide it
  if (iterationText === null && !iterChanges) {
    // No iteration and user wants to hide empties - we'll handle this in line building
    iterationText = null;
  } else if (iterationText === null && iterChanges) {
    // There's a change, so we should show something
    iterationText = getIterationDisplayText(hasIteration) || '';
  }
  
  // Build the complete line (no extra spacing between summary and iteration)
  // Using • (bullet character U+2022)
  let fullLine;
  if (iterationText !== null) {
    fullLine = `• ${epicData.showRag ? epicData.ragIndicator + ' ' : ''}${epicData.summary} | `;
  } else {
    // User configured to hide empty iterations and there's no value
    fullLine = `• ${epicData.showRag ? epicData.ragIndicator + ' ' : ''}${epicData.summary}`;
  }
  
  // Calculate text width for main part
  const mainTextLength = fullLine.length;
  
  // Add old iteration if changed (only if we're showing iterations)
  // Using → (rightwards arrow U+2192)
  if (iterationText !== null) {
    if (iterChanges && iterChanges.old) {
      fullLine += `${iterChanges.old} → `;
    }
    fullLine += iterationText;
  }
  
  // Create single text box
  const epicBox = slide.insertTextBox(fullLine, TEXT_INDENT, currentY, 
    SLIDE_CONFIG.CONTENT_WIDTH - TEXT_INDENT + SLIDE_CONFIG.MARGIN_LEFT, SLIDE_CONFIG.EPIC_HEIGHT);
  const epicText = epicBox.getText();
  epicText.getTextStyle().setFontFamily(SLIDE_CONFIG.BODY_FONT).setFontSize(SLIDE_CONFIG.BODY_FONT_SIZE).setForegroundColor(SLIDE_CONFIG.BLACK);
  addDisclosureToSlide(slide);
  
  // Style the summary with link
  const summaryStart = fullLine.indexOf(epicData.summary);
  if (summaryStart > -1) {
    const summaryRange = epicText.getRange(summaryStart, summaryStart + epicData.summary.length);
    summaryRange.getTextStyle().setLinkUrl(epicData.url).setForegroundColor(SLIDE_CONFIG.BLACK).setBold(true).setUnderline(false);
  }
  
  // Style old iteration with strikethrough if exists
  if (iterChanges && iterChanges.old) {
    const oldIterStart = fullLine.indexOf(iterChanges.old, mainTextLength);
    if (oldIterStart > -1) {
      epicText.getRange(oldIterStart, oldIterStart + iterChanges.old.length).getTextStyle()
        .setStrikethrough(true).setForegroundColor(SLIDE_CONFIG.GRAY);
    }
    // Highlight ONLY the new iteration value
    const newIterStart = fullLine.lastIndexOf(iterationText);
    if (newIterStart > -1) {
      // Create highlight shape behind text for just the iteration
      const highlightBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 
        TEXT_INDENT + (newIterStart * 4.5), currentY, 
        iterationText.length * 4.5, SLIDE_CONFIG.EPIC_HEIGHT);
      highlightBox.getFill().setSolidFill('#e8e0ff');
      highlightBox.getBorder().setTransparent();
      highlightBox.sendToBack();
      
      epicText.getRange(newIterStart, fullLine.length).getTextStyle()
            .setForegroundColor(hasIteration ? SLIDE_CONFIG.PURPLE : SLIDE_CONFIG.RED)
            .setFontSize(8)
            .setBold(false)
            .setItalic(true);
    }
  } else {
    // Style the iteration text normally
    const iterStart = fullLine.lastIndexOf(iterationText);
    if (iterStart > -1) {
      epicText.getRange(iterStart, fullLine.length).getTextStyle()
        .setForegroundColor(hasIteration ? SLIDE_CONFIG.PURPLE : SLIDE_CONFIG.RED)
        .setFontSize(8)
        .setBold(false)
        .setItalic(true);
    }
  }
  
  currentY += SLIDE_CONFIG.EPIC_HEIGHT;

  // Dependencies
  if (epicData.dependencies && epicData.dependencies.length > 0) {
    epicData.dependencies.forEach(dep => { 
        currentY = addDependencyToSlideWithIcons(slide, dep, currentY, changelogData, iterationNumber);
    });
  }
  
  // RAG Risk Note
  if (epicData.showRag) {
    let riskText = 'Epic Risk - ';
    
    // Check if RAG Note was updated
    if (epicChanges.details['RAG Note']) {
      // Add update indicator in BLUE and BOLD
      const updateText = `(Updated in Iteration ${iterationNumber}) `;
      riskText += updateText + (epicData.ragNote || 'NO RAG NOTE');
      
      const riskBox = slide.insertTextBox(riskText, TEXT_INDENT + 10, currentY, 
        SLIDE_CONFIG.CONTENT_WIDTH - TEXT_INDENT - 20, SLIDE_CONFIG.RISK_NOTE_HEIGHT);
      const riskTextObj = riskBox.getText();
      riskTextObj.getTextStyle().setFontFamily(SLIDE_CONFIG.BODY_FONT).setFontSize(SLIDE_CONFIG.SMALL_FONT_SIZE);
      
      // Style "Epic Risk -" in blue
      riskTextObj.getRange(0, 12).getTextStyle().setForegroundColor(SLIDE_CONFIG.BLUE).setBold(true);
      
      // Style the update indicator in BLUE and BOLD
      const updateStart = 12;
      const updateEnd = updateStart + updateText.length;
      riskTextObj.getRange(updateStart, updateEnd).getTextStyle()
        .setForegroundColor(SLIDE_CONFIG.BLUE).setBold(true);
      
      // Style NO RAG NOTE in red if missing
      if (!epicData.ragNote) {
        riskTextObj.getRange(updateEnd, riskText.length).getTextStyle()
          .setForegroundColor(SLIDE_CONFIG.RED).setBold(true);
      }
    } else {
      // No update, standard formatting
      riskText += epicData.ragNote || 'NO RAG NOTE';
      const riskBox = slide.insertTextBox(riskText, TEXT_INDENT + 10, currentY, 
        SLIDE_CONFIG.CONTENT_WIDTH - TEXT_INDENT - 20, SLIDE_CONFIG.RISK_NOTE_HEIGHT);
      const riskTextObj = riskBox.getText();
      riskTextObj.getTextStyle().setFontFamily(SLIDE_CONFIG.BODY_FONT).setFontSize(SLIDE_CONFIG.SMALL_FONT_SIZE);
      riskTextObj.getRange(0, 12).getTextStyle().setForegroundColor(SLIDE_CONFIG.BLUE).setBold(true);
      
      if (!epicData.ragNote) {
        riskTextObj.getRange(12, riskText.length).getTextStyle()
          .setForegroundColor(SLIDE_CONFIG.RED).setBold(true);
      }
    }
    currentY += SLIDE_CONFIG.RISK_NOTE_HEIGHT;
  }
  return currentY;
}

function addDependencyToSlideWithIcons(slide, dep, startY, changelogData, iterationNumber) {
    let currentY = startY;
    const indentX = SLIDE_CONFIG.MARGIN_LEFT + 40;
    const depChanges = changelogData[dep['Key']] || { details: {} };

    // --- ICONS LOGIC (Allow multiple icons) ---
    const iconsToDraw = [];
    
    // Check if it's new
    if (depChanges.isNew) {
        iconsToDraw.push(ICON_CONFIG.NEW);
    }
    
    // Check if it has changes (not new)
    if (!depChanges.isNew && Object.keys(depChanges.details).length > 0) {
        iconsToDraw.push(ICON_CONFIG.CHANGED);
    }
    
    // Check if Status changed TO 'Closed' or 'Done'
    if (depChanges.details['Status'] && depChanges.details['Status'].new) {
        const newStatus = depChanges.details['Status'].new.toLowerCase();
        if (newStatus.includes('closed') || newStatus.includes('done')) {
            iconsToDraw.push(ICON_CONFIG.COMPLETED);
        }
    }
    
    // Check if PI Commitment changed TO 'Deferred'
    if (depChanges.details['PI Commitment'] && depChanges.details['PI Commitment'].new) {
        const newCommitment = depChanges.details['PI Commitment'].new.toLowerCase();
        if (newCommitment.includes('deferred')) {
            iconsToDraw.push(ICON_CONFIG.DEFERRED);
        }
    }

    // Store icon images to send to back later
    const iconImages = [];
    
    // Draw icons on the LEFT side with zigzag pattern
    let iconX = SLIDE_CONFIG.MARGIN_LEFT + 10;
    const zigzagOffset = 2; // Smaller offset for dependencies
    iconsToDraw.forEach((iconId, index) => {
        try {
            const iconBlob = DriveApp.getFileById(iconId).getBlob();
            // Apply zigzag pattern
            const iconY = currentY + (SLIDE_CONFIG.DEPENDENCY_HEIGHT - ICON_CONFIG.SIZE) / 2 + (index % 2 * zigzagOffset);
            const iconImage = slide.insertImage(iconBlob, iconX, iconY, ICON_CONFIG.SIZE, ICON_CONFIG.SIZE);
            iconImages.push(iconImage);
            iconX += (ICON_CONFIG.SIZE + ICON_CONFIG.SPACING);
        } catch (e) { console.error(`Could not add icon with ID: ${iconId}`); }
    });
    
    // Send all icons to back
    iconImages.forEach(img => img.sendToBack());
    
    // Adjust text indent if icons were drawn
    const textIndent = iconsToDraw.length > 0 ? iconX + 8 : indentX;

    // --- Text Logic with highlighting ---
    // --- Text Logic with highlighting ---
    const iterChanges = depChanges.details['Iteration End'];
    const hasIter = (iterChanges ? iterChanges.new : dep['End Iteration Name'])?.trim();

    // Use configuration-aware function for iteration
    const iterText = getIterationDisplayText(hasIter);
    
    // Check if there are changes
    const hasChanges = Object.keys(depChanges.details).length > 0 || depChanges.isNew;
    
    // Build complete line
    let fullDepLine;
    if (iterText === null) {
    // User configured to hide empty iterations
    fullDepLine = `  ◦ 🔗 ${dep['Summary']} | ${dep['Depends on Valuestream'] || 'Unknown'}`;
    } else {
    // Show iteration
    fullDepLine = `  ◦ 🔗 ${dep['Summary']} | ${dep['Depends on Valuestream'] || 'Unknown'} | `;
    if (iterChanges && iterChanges.old) {
        fullDepLine += iterChanges.old + ' → ';
    }
    fullDepLine += iterText;
    }

    const depBox = slide.insertTextBox(fullDepLine, textIndent, currentY, SLIDE_CONFIG.CONTENT_WIDTH - (textIndent - SLIDE_CONFIG.MARGIN_LEFT) - 10, SLIDE_CONFIG.DEPENDENCY_HEIGHT);
    
    // Add highlight background if there are changes
    if (hasChanges) {
        depBox.getFill().setSolidFill('#e8e0ff');
    }
    
    const depText = depBox.getText();
    depText.getTextStyle().setFontFamily(SLIDE_CONFIG.BODY_FONT).setFontSize(SLIDE_CONFIG.SMALL_FONT_SIZE).setForegroundColor(SLIDE_CONFIG.BLACK);

    // Style the summary with link
    const summaryStart = fullDepLine.indexOf(dep['Summary']);
    if (summaryStart > -1) {
        const summaryRange = depText.getRange(summaryStart, summaryStart + dep['Summary'].length);
        summaryRange.getTextStyle().setLinkUrl(getJiraUrl(dep['Key'])).setForegroundColor(SLIDE_CONFIG.BLACK).setUnderline(false);
    }
    
    // Style old iteration with strikethrough if exists
    if (iterChanges && iterChanges.old) {
        const oldIterStart = fullDepLine.lastIndexOf(iterChanges.old);
        if (oldIterStart > -1) {
            depText.getRange(oldIterStart, oldIterStart + iterChanges.old.length).getTextStyle()
                .setStrikethrough(true).setForegroundColor(SLIDE_CONFIG.GRAY);
        }
    }
    
    // Style the iteration text (only if it exists)
    if (iterText !== null) {
      const iterStart = fullDepLine.lastIndexOf(iterText);
      if (iterStart > -1) {
        const iterRange = depText.getRange(iterStart, fullDepLine.length);
        
        if (hasIter) {
          // Has valid iteration - purple text
          iterRange.getTextStyle()
            .setForegroundColor(SLIDE_CONFIG.PURPLE)
            .setBold(true);
        } else {
          // Missing iteration - orange highlight background
          iterRange.getTextStyle()
            .setForegroundColor(SLIDE_CONFIG.MISSING_DATA_TEXT)
            .setBackgroundColor(SLIDE_CONFIG.MISSING_DATA_HIGHLIGHT)
            .setBold(true);
        }
      }
    }
    
    currentY += SLIDE_CONFIG.DEPENDENCY_HEIGHT;

    // Risk Note Logic remains the same...
    
    return currentY;
}
function applyGovernanceFilterToIterationData(ss, iterationData, piNumber) {
  const changelogSheetName = `PI ${piNumber} Changelog`;
  const changelogSheet = ss.getSheetByName(changelogSheetName);
  
  // If changelog exists, use its governance column
  if (changelogSheet) {
    console.log('Applying governance filter from changelog sheet...');
    
    const headers = changelogSheet.getRange(5, 1, 1, changelogSheet.getLastColumn()).getValues()[0];
    const governanceIndex = headers.indexOf('Include in Governance');
    
    if (governanceIndex === -1) {
      console.warn('Include in Governance column not found in changelog - no filtering applied');
      return iterationData;
    }
    
    const lastRow = changelogSheet.getLastRow();
    if (lastRow < 6) {
      return iterationData;
    }
    
    // Read key and governance columns
    const changelogData = changelogSheet.getRange(6, 1, lastRow - 5, governanceIndex + 1).getValues();
    
    // Build set of excluded keys
    const excludedKeys = new Set();
    changelogData.forEach(row => {
      const key = row[0];
      const governance = row[governanceIndex];
      if (key && governance === 'No') {
        excludedKeys.add(key);
      }
    });
    
    const filtered = iterationData.filter(epic => !excludedKeys.has(epic['Key']));
    console.log(`  Governance filter: ${iterationData.length} → ${filtered.length} (${excludedKeys.size} excluded)`);
    return filtered;
  }
  
  // No changelog - apply inline governance logic
  console.log('No changelog found - applying inline governance filter...');
  
  // Build epic governance map first (for dependencies to inherit)
  const epicGovernanceMap = {};
  iterationData.forEach(item => {
    if (item['Issue Type'] === 'Epic') {
      const allocation = item['Allocation'] || '';
      const portfolioInitiative = item['Portfolio Initiative'] || '';
      epicGovernanceMap[item['Key']] = shouldIncludeInGovernance(allocation, portfolioInitiative);
    }
  });
  
  // Filter all items
  const filtered = iterationData.filter(item => {
    if (item['Issue Type'] === 'Epic') {
      return epicGovernanceMap[item['Key']] === 'Yes';
    } else if (item['Issue Type'] === 'Dependency') {
      const parentKey = item['Parent Key'];
      return epicGovernanceMap[parentKey] === 'Yes';
    }
    return true; // Include other types by default
  });
  
  console.log(`  Inline governance filter: ${iterationData.length} → ${filtered.length}`);
  return filtered;
}