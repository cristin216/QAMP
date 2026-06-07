/*
================================================================================
TEACHER INVOICE CODE
================================================================================
Version: 32
Total Functions: 59
Documentation: See Teacher-Invoice-Functions.md
================================================================================
*/

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Teacher Invoice Tools')
      .addItem('Collect Monthly Invoice Data', 'showInvoiceGenerationUI')
      .addItem('Add Late Teacher', 'addLateTeacherToInvoice')
      .addItem('Generate Invoice Documents', 'generateTeacherInvoiceDocuments')
      .addItem('Print Documents', 'convertFolderDocsToPdfUI')
      .addItem('Verify Logs vs Invoices', 'verifyAttendanceVsInvoices')
      .addToUi();

    UtilityScriptLibrary.debugLog('onOpen', 'INFO', 'Teacher Invoice Tools menu created', '', '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('onOpen', 'ERROR', 'Error creating menu', '', error.message);
  }
}

// === MAIN UI FUNCTION ===


function showInvoiceGenerationUI() {
  try {
    UtilityScriptLibrary.debugLog('showInvoiceGenerationUI', 'INFO', 'Starting invoice generation UI', '', '');
    
    var ui = SpreadsheetApp.getUi();
    
    // Step 1: Get cutoff date first
    var cutoffDate = promptForCutoffDate();
    if (!cutoffDate) {
      ui.alert('❌ Cutoff date is required for invoice generation.');
      return;
    }
    
    // Step 2: Get invoice date
    var invoiceDate = promptForInvoiceDate();
    if (!invoiceDate) {
      ui.alert('❌ Invoice date is required for invoice generation.');
      return;
    }
    
    // Step 3: Get month name for sheet (with default based on cutoff date)
    var month = promptForMonthName(cutoffDate);
    if (!month) {
      ui.alert('❌ Month name is required for invoice generation.');
      return;
    }
    
    // Step 4: Get invoice period
    var invoicePeriod = promptForInvoicePeriod(cutoffDate);
    if (!invoicePeriod) {
      ui.alert('❌ Invoice period is required for invoice generation.');
      return;
    }
    
    // Step 5: Collect lesson data using parameters (no additional prompts)
    var lessonResults = collectUninvoicedLessonsUpToDate(cutoffDate, invoiceDate, invoicePeriod);
    
    if (!lessonResults) {
      ui.alert('❌ Lesson data collection failed.');
      return;
    }
    
    // Step 6: Generate invoice sheet
    var invoiceResults = generateMonthlyTeacherInvoices(month, cutoffDate, invoiceDate, lessonResults);
    
    if (!invoiceResults.success) {
      ui.alert('❌ Invoice generation failed: ' + invoiceResults.message);
      return;
    }
    
    // Step 7: Write metadata
    try {
      writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod);
      UtilityScriptLibrary.debugLog('showInvoiceGenerationUI', 'SUCCESS', 'Metadata written successfully', '', '');
    } catch (metadataError) {
      UtilityScriptLibrary.debugLog('showInvoiceGenerationUI', 'ERROR', 'Metadata write failed (non-critical)', '', metadataError.message);
      // Continue - this is not critical to invoice generation
    }
    
    // Step 8: Show comprehensive results summary
    showInvoiceGenerationResults(lessonResults, invoiceResults);
    
    UtilityScriptLibrary.debugLog('showInvoiceGenerationUI', 'SUCCESS', 'Invoice generation workflow completed', '', '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('showInvoiceGenerationUI', 'ERROR', 'Invoice generation workflow failed', '', error.message);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('❌ Error: ' + error.message + '\n\nCheck the Teacher_Invoice_Debug sheet for details.');
  }
}

// === GENERAL FUNCTIONS ===

function addLateTeacherToInvoice() {
  try {
    var ui = SpreadsheetApp.getUi();
    var activeSheet = SpreadsheetApp.getActiveSheet();
    
    // Validate that this is a monthly invoice sheet
    if (!isMonthlyInvoiceSheet(activeSheet)) {
      ui.alert('This does not appear to be a monthly invoice sheet. Please select the correct sheet.');
      return;
    }
    
    var monthName = activeSheet.getName();
    
    // Get metadata for this invoice month
    var metadata;
    try {
      metadata = getTeacherInvoicingMetadata(monthName);
      UtilityScriptLibrary.debugLog("addLateTeacherToInvoice", "INFO", 
                                    "Loaded metadata", 
                                    "Month: " + monthName, "");
    } catch (metadataError) {
      ui.alert('Could not load metadata for this month. Please ensure invoice data was collected first.');
      return;
    }
    
    // Get active teacher list
    var allTeachers = getActiveTeacherList();
    if (allTeachers.length === 0) {
      ui.alert('No active teachers found.');
      return;
    }
    
    // Build teacher selection prompt
    var teacherNames = [];
    for (var i = 0; i < allTeachers.length; i++) {
      teacherNames.push((i + 1) + '. ' + allTeachers[i].teacherName);
    }
    
    var response = ui.prompt(
      'Select Late Teacher',
      'Enter the number of the teacher to add:\n\n' + teacherNames.join('\n'),
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    var selectedIndex = parseInt(response.getResponseText()) - 1;
    if (isNaN(selectedIndex) || selectedIndex < 0 || selectedIndex >= allTeachers.length) {
      ui.alert('Invalid selection.');
      return;
    }
    
    var selectedTeacher = allTeachers[selectedIndex];
    
    // Generate invoice number for this teacher
    var invoiceNumber = generateInvoiceNumber(
      selectedTeacher.teacherId, 
      new Date(metadata.invoiceDate)
    );
    
    // Add invoice number to teacher data for collection
    selectedTeacher.invoiceNumber = invoiceNumber;
    
    // Collect lessons for just this teacher
    UtilityScriptLibrary.debugLog("addLateTeacherToInvoice", "INFO", 
                                  "Collecting lessons for late teacher", 
                                  "Teacher: " + selectedTeacher.teacherName, "");
    
    var teacherLessons = getUninvoicedLessonsForTeacher(
      selectedTeacher,
      new Date(metadata.lessonsEndingDate),
      new Date(metadata.invoiceDate),
      invoiceNumber
    );
    
    if (teacherLessons.length === 0) {
      ui.alert('No uninvoiced lessons found for ' + selectedTeacher.teacherName);
      return;
    }
    
    // Group the lessons
    var groupedLessons = groupLessonsByTeacherAndType(teacherLessons);
    
    // Validate
    var validation = validateLessonData(groupedLessons);
    if (!validation.isValid) {
      ui.alert('Validation issues found:\n\n' + validation.issues.join('\n'));
      return;
    }
    
    // Get invoice period from metadata
    var invoicePeriod = '';
    if (metadata.lessonsStartingDate && metadata.lessonsEndingDate) {
      var startDate = new Date(metadata.lessonsStartingDate);
      var endDate = new Date(metadata.lessonsEndingDate);
      var formattedStart = UtilityScriptLibrary.formatDateFlexible(startDate, "M/d/yyyy");
      var formattedEnd = UtilityScriptLibrary.formatDateFlexible(endDate, "M/d/yyyy");
      invoicePeriod = formattedStart + " - " + formattedEnd;
    }
    
    // Append to existing sheet
    var result = appendTeacherToInvoiceSheet(
      activeSheet, 
      groupedLessons, 
      new Date(metadata.invoiceDate),
      invoicePeriod,
      metadata.ratesLookup
    );
    
    ui.alert(
      'Successfully added late teacher:\n\n' +
      'Teacher: ' + selectedTeacher.teacherName + '\n' +
      'Line items: ' + result.lineItemCount + '\n' +
      'Lessons marked: ' + teacherLessons.length + '\n\n' +
      'You can now run "Generate Invoice Documents" to create the document for this teacher.'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addLateTeacherToInvoice", "ERROR", 
                                  "Failed to add late teacher", "", error.message);
    SpreadsheetApp.getUi().alert('Error adding late teacher: ' + error.message);
  }
}

function addStudentLineItem(sheet, row, columnMap, lessonGroup, teacherName, ratesCache, programRateKeysCache) {
  try {
    // Calculate rate and cost for this lesson type using cached rates
    var rateCalculation = calculateLessonRateAndCost(lessonGroup, ratesCache, programRateKeysCache);
    
    // Parse student name into first and last names
    var studentNames = parseStudentName(lessonGroup.studentName);
    
    var rowData = new Array(sheet.getLastColumn()).fill('');
    
    // Populate student line item row using normalized column names (convert to 0-based)
    if (columnMap['lastname'] !== undefined) rowData[columnMap['lastname'] - 1] = studentNames.lastName;
    if (columnMap['firstname'] !== undefined) rowData[columnMap['firstname'] - 1] = studentNames.firstName;
    if (columnMap['id'] !== undefined) rowData[columnMap['id'] - 1] = lessonGroup.studentId;
    if (columnMap['teacher'] !== undefined) rowData[columnMap['teacher'] - 1] = teacherName;
    if (columnMap['duration'] !== undefined) rowData[columnMap['duration'] - 1] = lessonGroup.lessonLength;
    if (columnMap['quantity'] !== undefined) rowData[columnMap['quantity'] - 1] = lessonGroup.quantity;
    if (columnMap['rate'] !== undefined) rowData[columnMap['rate'] - 1] = rateCalculation.lessonRate;
    if (columnMap['cost'] !== undefined) rowData[columnMap['cost'] - 1] = rateCalculation.totalCost;
    // Leave other columns blank for student rows
    
    // Add the row data
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    
    // Apply attendance-style formatting
    var dataRange = sheet.getRange(row, 1, 1, rowData.length);
    dataRange.setFontWeight('normal')
             .setFontColor('black')
             .setWrap(true);
    
    // Apply alternating background
    var isEvenRow = row % 2 === 0;
    if (isEvenRow) {
      dataRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      dataRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }
    
    UtilityScriptLibrary.debugLog("addStudentLineItem", "DEBUG", 
                                  "Added student line item with 0-based indices", 
                                  "Student: " + lessonGroup.studentName + ", Row: " + row + ", Rate: " + rateCalculation.lessonRate + ", Cost: " + rateCalculation.totalCost, "");
    
    return row + 1;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentLineItem", "ERROR", 
                                  "Failed to add student line item", 
                                  "Student: " + lessonGroup.studentName, error.message);
    throw error;
  }
}

function addTeacherHeaderRow(sheet, row, columnMap, teacherInfo, invoiceDate, invoiceNumber, invoicePeriod) {
  try {
    var rowData = new Array(sheet.getLastColumn()).fill('');
    
    // Populate teacher header row using normalized column names (convert to 0-based)
    if (columnMap['lastname'] !== undefined) rowData[columnMap['lastname'] - 1] = teacherInfo.lastName || '';
    if (columnMap['firstname'] !== undefined) rowData[columnMap['firstname'] - 1] = teacherInfo.firstName || '';
    if (columnMap['id'] !== undefined) rowData[columnMap['id'] - 1] = teacherInfo.teacherId || '';
    if (columnMap['teacher'] !== undefined) rowData[columnMap['teacher'] - 1] = teacherInfo.teacherName || '';
    if (columnMap['invoicedate'] !== undefined) rowData[columnMap['invoicedate'] - 1] = UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MM/dd/yyyy');
    if (columnMap['invoicenumber'] !== undefined) rowData[columnMap['invoicenumber'] - 1] = invoiceNumber || '';
    if (columnMap['invoiceperiod'] !== undefined) rowData[columnMap['invoiceperiod'] - 1] = invoicePeriod || '';
    if (columnMap['address'] !== undefined) rowData[columnMap['address'] - 1] = teacherInfo.address || '';
    
   
    
    // Add the row data
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    
    // Apply teacher header formatting
    var dataRange = sheet.getRange(row, 1, 1, rowData.length);
    dataRange.setFontWeight('bold')
             .setFontColor(UtilityScriptLibrary.STYLES.SUBHEADER.text)
             .setWrap(true)
             .setBackground(UtilityScriptLibrary.STYLES.SUBHEADER.background);
    
    UtilityScriptLibrary.debugLog("addTeacherHeaderRow", "DEBUG", 
                                  "Added teacher header row with normalized columns", 
                                  "Teacher: " + teacherInfo.teacherName + ", Row: " + row, "");
    
    return row + 1;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addTeacherHeaderRow", "ERROR", "Failed to add teacher header row", 
                                  "Teacher: " + (teacherInfo.teacherName || 'unknown'), error.message);
    throw error;
  }
}

function appendTeacherToInvoiceSheet(sheet, groupedLessons, invoiceDate, invoicePeriod, ratesColumnHeader) {
  try {
    var columnMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Load caches once for all rate lookups
    var ratesCache = loadRatesCache(ratesColumnHeader);
    var programRateKeysCache = loadProgramRateKeysCache();
    
    // Start at the next available row
    var currentRow = sheet.getLastRow() + 1;
    
    // Add blank separator row
    currentRow++;
    
    var lineItemCount = 0;
    var currentTeacher = null;
    
    // Process each lesson group
    for (var key in groupedLessons) {
      if (groupedLessons.hasOwnProperty(key)) {
        var lessonGroup = groupedLessons[key];
        
        if (lessonGroup.teacherName !== currentTeacher) {
          currentTeacher = lessonGroup.teacherName;
          
          var teacherInfo = getTeacherInfoFromLessonGroup(lessonGroup);
          
          currentRow = addTeacherHeaderRow(
            sheet, currentRow, columnMap, teacherInfo, 
            invoiceDate, lessonGroup.invoiceNumber, invoicePeriod
          );
        }
        
        currentRow = addStudentLineItem(
          sheet, currentRow, columnMap, lessonGroup, currentTeacher, 
          ratesCache, programRateKeysCache
        );
        lineItemCount++;
      }
    }
    
    UtilityScriptLibrary.debugLog("appendTeacherToInvoiceSheet", "SUCCESS", 
                                  "Appended teacher data", 
                                  "Line items: " + lineItemCount, "");
    
    return {
      success: true,
      lineItemCount: lineItemCount
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("appendTeacherToInvoiceSheet", "ERROR", 
                                  "Failed to append teacher", "", error.message);
    throw error;
  }
}

function buildTeacherInvoiceFileName(teacherData, variables) {
  try {
    var teacherName = variables.TeacherFirstName && variables.TeacherLastName ? 
                      variables.TeacherFirstName + ' ' + variables.TeacherLastName : 
                      teacherData.teacherName;
    
    var invoiceNumber = variables.InvoiceNumber || 'NOINVOICE';
    
    return teacherName + ' - ' + invoiceNumber;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("buildTeacherInvoiceFileName", "ERROR", "Failed to build filename", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    return 'TeacherInvoice - ' + teacherData.teacherName;
  }
}

function buildTeacherInvoiceVariableMap(teacherData, teacherContactInfo, metadata) {
  try {
    // Build dynamic content directly from the sheet's student rows
    var dynamicDescription = [];
    var dynQty = [];
    var dynRate = [];
    var dynCost = [];
    var totalCost = 0;
    
    for (var i = 0; i < teacherData.studentRows.length; i++) {
      var row = teacherData.studentRows[i];
      
      // Format: "Last, First [Duration] minute lessons" (matches existing sheet format)
      var description = row.lastName + ', ' + row.firstName + ' ' + row.duration + ' minute lessons';
      dynamicDescription.push(description);
      
      dynQty.push(row.quantity.toString());
      dynRate.push(UtilityScriptLibrary.formatCurrency(row.rate));
      dynCost.push(UtilityScriptLibrary.formatCurrency(row.cost));
      
      totalCost += row.cost;
    }
    
    // Use teacher's first and last name directly from the sheet
    var teacherFirstName = teacherData.teacherFirstName || '';
    var teacherLastName = teacherData.teacherLastName || '';
    var teacherAddress = teacherData.teacherAddress || '';
    
    // Only fall back to parsing teacherName if we don't have the names from sheet
    if (!teacherFirstName && !teacherLastName && teacherData.teacherName) {
      var nameParts = teacherData.teacherName.split(' ');
      if (nameParts.length >= 2) {
        teacherFirstName = nameParts[0];
        teacherLastName = nameParts.slice(1).join(' ');
      } else {
        teacherLastName = teacherData.teacherName; // Use full name as last name
      }
    }
    
    // Format invoice date
    var formattedInvoiceDate = '';
    if (teacherData.invoiceDate) {
      var invoiceDate = typeof teacherData.invoiceDate === 'string' ?
        new Date(teacherData.invoiceDate) : teacherData.invoiceDate;
      formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MMMM d, yyyy');
    }
    
    // Format invoice period - USE METADATA FOR DATE RANGE
    var formattedInvoicePeriod = '';
    if (metadata && metadata.lessonsStartingDate && metadata.lessonsEndingDate) {
      // Use metadata date range
      var startDate = new Date(metadata.lessonsStartingDate);
      var endDate = new Date(metadata.lessonsEndingDate);
      
      var formattedStart = UtilityScriptLibrary.formatDateFlexible(startDate, "MMMM d, yyyy");
      var formattedEnd = UtilityScriptLibrary.formatDateFlexible(endDate, "MMMM d, yyyy");
      formattedInvoicePeriod = formattedStart + " - " + formattedEnd;
      
      UtilityScriptLibrary.debugLog("buildTeacherInvoiceVariableMap", "DEBUG", 
                                    "Using metadata for invoice period", 
                                    "Period: " + formattedInvoicePeriod, "");
    } else {
      // Fallback to teacherData (old behavior)
      if (teacherData.invoicePeriod) {
        if (typeof teacherData.invoicePeriod === 'string') {
          formattedInvoicePeriod = teacherData.invoicePeriod; // Already a string
        } else {
          // It's a Date object, format it as "MMMM yyyy"
          formattedInvoicePeriod = UtilityScriptLibrary.formatDateFlexible(teacherData.invoicePeriod, 'MMMM yyyy');
        }
      }
      
      UtilityScriptLibrary.debugLog("buildTeacherInvoiceVariableMap", "WARNING", 
                                    "Using fallback for invoice period (no metadata)", 
                                    "Period: " + formattedInvoicePeriod, "");
    }
    
    return {
      'InvoiceNumber': teacherData.invoiceNumber || '',
      'InvoiceDate': formattedInvoiceDate,
      'InvoicePeriod': formattedInvoicePeriod,
      'TeacherFirstName': teacherFirstName,
      'TeacherLastName': teacherLastName,
      'TeacherAddress': teacherAddress,
      'DYNAMIC_DESCRIPTION': dynamicDescription.join('\n'),
      'DYN_QTY': dynQty.join('\n'),
      'DYN_RATE': dynRate.join('\n'),
      'DYN_COST': dynCost.join('\n'),
      'Total': UtilityScriptLibrary.formatCurrency(totalCost),
      'Comment': teacherData.comment || ''
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("buildTeacherInvoiceVariableMap", "ERROR", "Failed to build variable map", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    throw error;
  }
}

function calculateLessonRateAndCost(lessonGroup, ratesCache, programRateKeysCache) {
  try {
    // Debug logging to track the issue
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                  "Starting rate calculation", 
                                  "Student/Group: " + lessonGroup.studentName + ", ID: " + lessonGroup.studentId + ", Length: " + lessonGroup.lessonLength + ", Quantity: " + lessonGroup.quantity, "");
    
    // Get the semester for rate lookup using the first lesson date
    if (!lessonGroup.lessons || lessonGroup.lessons.length === 0) {
      throw new Error("No lessons found in lesson group");
    }
    
    var firstLessonDate = lessonGroup.lessons[0].lessonDate;
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                  "First lesson date", 
                                  "Date: " + UtilityScriptLibrary.formatDateFlexible(firstLessonDate, 'MMM d, yyyy'), "");
    
    var semesterName = getSemesterForDate(firstLessonDate);
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                  "Semester determined", 
                                  "Semester: " + semesterName, "");
    
    // Determine if this is a group session (ID starts with "G")
    var isGroupSession = lessonGroup.studentId && lessonGroup.studentId.toString().charAt(0) === 'G';
    
    var hourlyRate;
    var lessonRate;
    var totalCost;
    
    if (isGroupSession) {
      // GROUP SESSION RATE CALCULATION
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Detected group session", 
                                    "Group ID: " + lessonGroup.studentId, "");
      
      // Get the Rate Key from cache
      var rateKey = getRateKeyForProgram(lessonGroup.studentName, programRateKeysCache);
      
      if (!rateKey) {
        throw new Error("Could not find Rate Key for group: " + lessonGroup.studentName);
      }
      
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Rate Key found", 
                                    "Rate Key: " + rateKey, "");
      
      // Look up group pay rate using Rate Key + " Pay" from cache
      var rateType = rateKey + ' Pay';
      hourlyRate = getRateForSemester(semesterName, rateType, ratesCache);
      
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Group hourly rate retrieved", 
                                    "Rate Type: " + rateType + ", Rate: " + hourlyRate, "");
      
      // Calculate lesson rate based on duration (if provided)
      if (lessonGroup.lessonLength && lessonGroup.lessonLength > 0) {
        var lessonRateMultiplier = lessonGroup.lessonLength / 60; // Convert minutes to hours
        lessonRate = hourlyRate * lessonRateMultiplier;
      } else {
        // If no duration specified, use hourly rate as-is
        lessonRate = hourlyRate;
      }
      
      // Calculate total cost
      totalCost = lessonRate * lessonGroup.quantity;
      
    } else {
      // STUDENT LESSON RATE CALCULATION
      hourlyRate = getRateForSemester(semesterName, 'Lessons Pay', ratesCache);
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Student hourly rate retrieved", 
                                    "Rate: " + hourlyRate, "");
      
      // Calculate lesson rate based on duration
      var lessonRateMultiplier = lessonGroup.lessonLength / 60; // Convert minutes to hours
      lessonRate = hourlyRate * lessonRateMultiplier;
      
      // Calculate total cost
      totalCost = lessonRate * lessonGroup.quantity;
    }
    
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "INFO", 
                                  "Rate calculation completed", 
                                  "Lesson rate: " + lessonRate + ", Total cost: " + totalCost, "");
    
    return {
      hourlyRate: hourlyRate,
      lessonRate: lessonRate,
      totalCost: totalCost
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "ERROR", 
                                  "Rate calculation failed", 
                                  "Student/Group: " + lessonGroup.studentName, error.message);
    throw error;
  }
}

function calculateLessonsStartingDate(metadataSheet) {
  try {
    var lastRow = metadataSheet.getLastRow();
    
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "INFO", 
                                    "First invoice ever - no previous metadata", 
                                    "Will use first day of invoice month", "");
      return null;
    }
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    var endingDateCol = headerMap[UtilityScriptLibrary.normalizeHeader("Lessons Ending Date")];
    
    if (!endingDateCol) {
      throw new Error("Lessons Ending Date column not found in metadata");
    }
    
    var previousEndingDate = metadataSheet.getRange(lastRow, endingDateCol).getValue();
    
    if (!previousEndingDate || !(previousEndingDate instanceof Date)) {
      UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "WARNING", 
                                    "Invalid previous ending date", 
                                    "Value: " + previousEndingDate, "");
      return null;
    }
    
    var startingDate = new Date(previousEndingDate);
    startingDate.setDate(startingDate.getDate() + 1);
    
    UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "INFO", 
                                  "Calculated starting date", 
                                  "Previous ending: " + UtilityScriptLibrary.formatDateFlexible(previousEndingDate, 'M/d/yy') + 
                                  ", New starting: " + UtilityScriptLibrary.formatDateFlexible(startingDate, 'M/d/yy'), "");
    
    return startingDate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "ERROR", 
                                  "Failed to calculate starting date", 
                                  "", error.message);
    return null;
  }
}

function collectUninvoicedLessonsUpToDate(cutoffDate, invoiceDate, invoicePeriod) {
  try {
    UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'INFO', 
                                  'Starting lesson collection', 
                                  'Cutoff: ' + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy') + ', Invoice Date: ' + UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'M/d/yy'), '');
    
    var formResponsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
    var lookupSheet = formResponsesSS.getSheetByName('Teacher Roster Lookup');
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      throw new Error('Teacher Roster Lookup sheet not found or empty');
    }
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(lookupSheet);
    var firstNameCol   = headerMap[UtilityScriptLibrary.normalizeHeader('First Name')];
    var lastNameCol    = headerMap[UtilityScriptLibrary.normalizeHeader('Last Name')];
    var rosterUrlCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Roster URL')];
    var teacherIdCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    var displayNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Display Name')];
    var statusCol      = headerMap[UtilityScriptLibrary.normalizeHeader('Status')];

    var data = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, lookupSheet.getLastColumn()).getValues();
    var allTeachers = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var firstName   = firstNameCol   ? row[firstNameCol - 1].toString().trim()   : '';
      var lastName    = lastNameCol    ? row[lastNameCol - 1].toString().trim()    : '';
      var rosterUrl   = rosterUrlCol   ? row[rosterUrlCol - 1].toString().trim()   : '';
      var teacherId   = teacherIdCol   ? row[teacherIdCol - 1].toString().trim()   : '';
      var displayName = displayNameCol ? row[displayNameCol - 1].toString().trim() : '';
      var status      = statusCol      ? row[statusCol - 1].toString().trim()      : '';
      var teacherName = (firstName + ' ' + lastName).trim();
      
      if (teacherName && rosterUrl !== '') {
        var invoiceNumber = generateInvoiceNumber(teacherId || teacherName, invoiceDate);
        
        allTeachers.push({
          teacherName: teacherName,
          firstName: firstName,
          lastName: lastName,
          rosterUrl: rosterUrl,
          teacherId: teacherId || '',
          displayName: displayName || '',
          status: status || 'Unknown',
          invoiceNumber: invoiceNumber
        });
      }
    }
    
    UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'INFO', 
                                  'Found teachers', 
                                  'Count: ' + allTeachers.length, '');
    
    var allUninvoicedLessons = [];
    var errors = [];
    
    for (var i = 0; i < allTeachers.length; i++) {
      var teacher = allTeachers[i];
      try {
        var teacherLessons = getUninvoicedLessonsForTeacher(
          teacher, 
          cutoffDate, 
          invoiceDate, 
          teacher.invoiceNumber
        );
        
        if (teacherLessons.length > 0) {
          allUninvoicedLessons = allUninvoicedLessons.concat(teacherLessons);
          UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'SUCCESS', 
                                        'Teacher processed', 
                                        teacher.teacherName + ': ' + teacherLessons.length + ' lessons', '');
        }
      } catch (error) {
        errors.push({
          teacher: teacher.teacherName,
          error: error.message,
          timestamp: new Date()
        });
        UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'ERROR', 
                                      'Teacher failed', 
                                      teacher.teacherName, error.message);
      }
    }
    
    var groupedLessons = groupLessonsByTeacherAndType(allUninvoicedLessons);
    var validationResults = validateLessonData(groupedLessons);
    
    return {
      lessons: groupedLessons,
      allLessons: allUninvoicedLessons,
      errors: errors,
      validation: validationResults,
      cutoffDate: cutoffDate,
      invoiceDate: invoiceDate,
      invoicePeriod: invoicePeriod,
      summary: {
        totalTeachers: allTeachers.length,
        successfulTeachers: allTeachers.length - errors.length,
        totalLessons: allUninvoicedLessons.length,
        totalLineItems: Object.keys(groupedLessons).length
      }
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'ERROR', 
                                  'Lesson collection failed', '', error.message);
    throw error;
  }
}

function convertFolderDocsToPdfUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeTabName = activeSheet.getName();
  
  // Validate that this is a monthly invoice sheet
  if (!isMonthlyInvoiceSheet(activeSheet)) {
    SpreadsheetApp.getUi().alert('This does not appear to be a monthly invoice sheet. Please select the correct sheet.');
    return;
  }
  
  // Get metadata for this invoice month
  var metadata;
  try {
    metadata = getTeacherInvoicingMetadata(activeTabName);
    UtilityScriptLibrary.debugLog("convertFolderDocsToPdfUI", "INFO", 
                                  "Loaded metadata", 
                                  "Month: " + activeTabName, "");
  } catch (metadataError) {
    UtilityScriptLibrary.debugLog("convertFolderDocsToPdfUI", "WARNING", 
                                  "Could not load metadata (will proceed without)", 
                                  "Month: " + activeTabName, metadataError.message);
    metadata = null;
  }
  
  // Extract teacher data from sheet to match PDFs back to teachers
  var teacherData = extractTeachersFromFormattedSheet(activeSheet);
  var teacherMap = {};
  var invoiceHeaderMap = UtilityScriptLibrary.getHeaderMap(activeSheet);
  var pdfCol = invoiceHeaderMap['pdf'];
  
  // Build a map of "FirstName LastName - InvoiceNumber" to teacher data
  for (var i = 0; i < teacherData.length; i++) {
    var teacher = teacherData[i];
    var teacherName = teacher.teacherFirstName && teacher.teacherLastName ? 
                      teacher.teacherFirstName + ' ' + teacher.teacherLastName : 
                      teacher.teacherName;
    var key = teacherName + ' - ' + teacher.invoiceNumber;
    teacherMap[key] = teacher;
  }
  
  // Get environment automatically from EnvironmentManager
  var env = UtilityScriptLibrary.EnvironmentManager.get();
  var config = UtilityScriptLibrary.getConfig();
  var parentFolder = DriveApp.getFolderById(config[env].generatedDocumentsFolderId);
  
  // Navigate to Teacher Invoices folder
  var teacherInvoicesIterator = parentFolder.getFoldersByName('Teacher Invoices');
  if (!teacherInvoicesIterator.hasNext()) {
    SpreadsheetApp.getUi().alert('Teacher Invoices folder not found');
    return;
  }
  var teacherInvoicesFolder = teacherInvoicesIterator.next();
  
  // Find the folder matching the active tab name
  var folderIterator = teacherInvoicesFolder.getFoldersByName(activeTabName);
  if (!folderIterator.hasNext()) {
    SpreadsheetApp.getUi().alert('Folder "' + activeTabName + '" not found in Teacher Invoices');
    return;
  }
  var monthFolder = folderIterator.next();
  
  // Create a "[folder name] PDFs" subfolder if it doesn't exist
  var pdfFolderName = activeTabName + ' PDFs';
  var pdfFolder;
  var pdfFolderIterator = monthFolder.getFoldersByName(pdfFolderName);
  if (pdfFolderIterator.hasNext()) {
    pdfFolder = pdfFolderIterator.next();
  } else {
    pdfFolder = monthFolder.createFolder(pdfFolderName);
  }
  
  // Get existing PDFs to check for duplicates
  var existingPdfs = {};
  var pdfFiles = pdfFolder.getFiles();
  while (pdfFiles.hasNext()) {
    existingPdfs[pdfFiles.next().getName()] = true;
  }
  
  // Get all Google Docs files in the month folder
  var files = monthFolder.getFilesByType(MimeType.GOOGLE_DOCS);
  var convertedCount = 0;
  var skippedCount = 0;
  var rosterUpdateCount = 0;
  
  while (files.hasNext()) {
    var file = files.next();
    var docId = file.getId();
    var docName = file.getName();
    var pdfName = docName + '.pdf';
    
    // Check if PDF already exists
    if (existingPdfs[pdfName]) {
      skippedCount++;
      continue;
    }
    
    // Get the document and convert to PDF
    var doc = DocumentApp.openById(docId);
    var pdfBlob = doc.getAs('application/pdf');
    pdfBlob.setName(pdfName);
    
    // Save PDF to the PDFs folder and set view-only sharing for anyone with the link
    var pdfFile = pdfFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    convertedCount++;
    
    // Match PDF to teacher data and update records
    var teacher = teacherMap[docName];
    if (teacher) {
      // Write PDF URL to the PDF column on the invoice sheet
      if (pdfCol && teacher.headerRowIndex) {
        activeSheet.getRange(teacher.headerRowIndex, pdfCol)
                   .setFormula('=HYPERLINK("' + pdfFile.getUrl() + '", "' + pdfFile.getId() + '")');
      }
      
      try {
        var pdfResult = {
          success: true,
          fileId: pdfFile.getId(),
          url: pdfFile.getUrl()
        };
        
        updateTeacherInvoiceHistory(teacher, pdfResult, metadata);
        rosterUpdateCount++;
        
        UtilityScriptLibrary.debugLog("convertFolderDocsToPdfUI", "SUCCESS", 
                                      "Updated roster with PDF URL", 
                                      "Teacher: " + teacher.teacherName, "");
      } catch (error) {
        UtilityScriptLibrary.debugLog("convertFolderDocsToPdfUI", "ERROR", 
                                      "Failed to update roster", 
                                      "Teacher: " + teacher.teacherName, error.message);
      }
    } else {
      UtilityScriptLibrary.debugLog("convertFolderDocsToPdfUI", "WARNING", 
                                    "Could not match PDF to teacher data", 
                                    "Document name: " + docName, "");
    }
  }
  
  var envLabel = env.toUpperCase();
  SpreadsheetApp.getUi().alert(
    '[' + envLabel + '] Converted ' + convertedCount + ' new document(s) to PDF.\n' +
    'Skipped ' + skippedCount + ' already converted.\n' +
    'Updated ' + rosterUpdateCount + ' teacher roster(s) with PDF URLs.\n' +
    'PDF folder: ' + pdfFolder.getUrl()
  );
}

function createMonthlyInvoiceSheet(month) {
  try {
    // Access active Teacher Invoices workbook
    var teacherInvoicesSS = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if month sheet already exists
    var monthSheet = teacherInvoicesSS.getSheetByName(month);
    if (monthSheet) {
      UtilityScriptLibrary.debugLog("createMonthlyInvoiceSheet", "WARNING", 
                                    "Sheet already exists", "Month: " + month, "");
      
      // Clear existing data but keep headers
      if (monthSheet.getLastRow() > 1) {
        monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, monthSheet.getLastColumn()).clearContent();
      }
      return monthSheet;
    }
    
    // Get the Monthly Template sheet
    var templateSheet = teacherInvoicesSS.getSheetByName("Monthly Template");
    if (!templateSheet) {
      throw new Error("Monthly Template sheet not found in Teacher Invoices workbook");
    }
    
    // Copy template to create new month sheet
    monthSheet = templateSheet.copyTo(teacherInvoicesSS);
    monthSheet.setName(month);
    
    UtilityScriptLibrary.debugLog("createMonthlyInvoiceSheet", "INFO", 
                                  "Created month sheet from template", "Month: " + month, "");
    
    return monthSheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("createMonthlyInvoiceSheet", "ERROR", 
                                  "Failed to create month sheet", "Month: " + month, error.message);
    throw error;
  }
}

function createStudentReconciliationReport(workbook, attendanceData, invoiceData) {
  var sheetName = 'Student Reconciliation Report';
  var sheet = workbook.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = workbook.insertSheet(sheetName);
  }
  
  // Headers
  var headers = ['Teacher', 'Invoice Date', 'Student ID', 'Attendance Hours', 'Invoice Hours', 'Match', 'Issue'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff');
  
  var outputData = [];
  
  // Get all teachers
  var allTeachers = {};
  for (var teacher in attendanceData) {
    allTeachers[teacher] = true;
  }
  for (var teacher in invoiceData) {
    allTeachers[teacher] = true;
  }
  
  // Process each teacher × invoice date combination
  for (var teacher in allTeachers) {
    var teacherAttendance = attendanceData[teacher] || {};
    var teacherInvoice = invoiceData[teacher] || {};
    
    // Get all invoice dates for this teacher
    var allInvoiceDates = {};
    for (var invDate in teacherAttendance) {
      allInvoiceDates[invDate] = true;
    }
    for (var invDate in teacherInvoice) {
      allInvoiceDates[invDate] = true;
    }
    
    for (var invoiceDate in allInvoiceDates) {
      var attendanceStudents = teacherAttendance[invoiceDate] || {};
      var invoiceStudents = teacherInvoice[invoiceDate] || {};
      
      // Get all students for this combination
      var allStudents = {};
      for (var studentId in attendanceStudents) {
        allStudents[studentId] = true;
      }
      for (var studentId in invoiceStudents) {
        allStudents[studentId] = true;
      }
      
      // Compare each student
      for (var studentId in allStudents) {
        var attendanceHours = attendanceStudents[studentId] || 0;
        var invoiceHours = invoiceStudents[studentId] || 0;
        
        var match = '';
        var issue = '';
        
        if (attendanceHours > 0 && invoiceHours === 0) {
          match = 'X';
          issue = 'Missing from invoice';
        } else if (attendanceHours === 0 && invoiceHours > 0) {
          match = 'X';
          issue = 'Not in attendance';
        } else if (Math.abs(attendanceHours - invoiceHours) > 0.01) {
          match = 'X';
          issue = 'Hour mismatch';
        } else {
          match = '✓';
          issue = '';
        }
        
        outputData.push([
          teacher,
          invoiceDate,
          studentId,
          attendanceHours > 0 ? attendanceHours : '',
          invoiceHours > 0 ? invoiceHours : '',
          match,
          issue
        ]);
      }
    }
  }
  
  // Write data
  if (outputData.length > 0) {
    sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
    
    // Apply conditional formatting
    for (var row = 0; row < outputData.length; row++) {
      if (outputData[row][5] === 'X') {
        sheet.getRange(row + 2, 6).setBackground('#ffcccc');
      } else if (outputData[row][5] === '✓') {
        sheet.getRange(row + 2, 6).setBackground('#ccffcc');
      }
    }
  }
  
  UtilityScriptLibrary.debugLog('createStudentReconciliationReport', 'SUCCESS', 
                                'Report created', outputData.length + ' student comparisons', '');
}

function createUnpaidLessonsReport(workbook, unpaidLessons) {
  var sheetName = 'Unpaid Lessons Report';
  var sheet = workbook.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = workbook.insertSheet(sheetName);
  }
  
  // Headers
  var headers = ['Teacher', 'Attendance Sheet', 'Student ID', 'Date', 'Hours'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('#ffffff');
  
  // Data
  if (unpaidLessons.length > 0) {
    var outputData = [];
    for (var i = 0; i < unpaidLessons.length; i++) {
      var lesson = unpaidLessons[i];
      outputData.push([
        lesson.teacher,
        lesson.sheet,
        lesson.studentId,
        lesson.date,
        lesson.hours
      ]);
    }
    
    sheet.getRange(2, 1, outputData.length, headers.length).setValues(outputData);
  }
  
  UtilityScriptLibrary.debugLog('createUnpaidLessonsReport', 'SUCCESS', 
                                'Report created', unpaidLessons.length + ' unpaid lessons', '');
}

function extractAndMarkLessonsFromSheet(sheet, teacherData, cutoffDate, formattedInvoiceDate, invoiceNumber) {
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // Empty sheet or only headers
    }
    
    // Use UtilityScriptLibrary to get header map (returns 1-based column numbers)
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Validate required columns exist
    var requiredColumns = ['studentid', 'studentname', 'date', 'length', 'status'];
    for (var j = 0; j < requiredColumns.length; j++) {
      var col = requiredColumns[j];
      if (!headerMap[col]) {
        throw new Error('Required column "' + col + '" not found in ' + sheet.getName());
      }
    }
    
    // Get column indices (convert to 0-based for array access, but keep 1-based for setRange)
    var studentIdCol = headerMap['studentid'];
    var studentNameCol = headerMap['studentname'];
    var dateCol = headerMap['date'];
    var lengthCol = headerMap['length'];
    var statusCol = headerMap['status'];
    var invoiceDateCol = headerMap['invoicedate'];
    var invoiceNumberCol = headerMap['invoicenumber'];
    
    var lessons = [];
    var rowsToUpdate = []; // Track which rows need invoice data written
    
    // Process data rows (skip header row 0)
    for (var k = 1; k < data.length; k++) {
      var row = data[k];
      var sheetRowIndex = k + 1; // Convert 0-based to 1-based sheet row
      
      // Extract lesson data using 0-based indices
      var studentId = row[studentIdCol - 1];
      var studentName = row[studentNameCol - 1];
      var dateValue = row[dateCol - 1];
      var lengthValue = row[lengthCol - 1];
      var statusValue = row[statusCol - 1];
      var existingInvoiceDate = invoiceDateCol ? row[invoiceDateCol - 1] : null;
      
      // Skip if already invoiced
      if (existingInvoiceDate && existingInvoiceDate.toString().trim() !== '') {
        continue;
      }
      
      // Skip if no student ID or student name
      if (!studentId || !studentName || studentId.toString().trim() === '' || studentName.toString().trim() === '') {
        continue;
      }
      
      // Parse lesson date
      var lessonDate = parseLessonDate(dateValue);
      if (!lessonDate || isNaN(lessonDate.getTime())) {
        continue;
      }
      
      // Skip if after cutoff date
      if (lessonDate > cutoffDate) {
        continue;
      }
      
      // Parse status
      var status = statusValue ? statusValue.toString().toLowerCase().trim() : '';
      
      // Skip invalid statuses
      if (status !== 'lesson' && status !== 'no show') {
        continue;
      }
      
      // Parse lesson length
      var lessonLength = parseInt(lengthValue) || 30;
      
      // This is a valid uninvoiced lesson - collect it
      var lesson = {
        teacherName: teacherData.teacherName,
        teacherId: teacherData.teacherId,
        studentId: studentId.toString().trim(),
        studentName: studentName.toString().trim(),
        lessonDate: lessonDate,
        status: status === 'no show' ? 'No Show' : 'Lesson',
        lessonLength: lessonLength,
        sheetName: sheet.getName(),
        sheetRow: sheetRowIndex
      };
      
      lessons.push(lesson);
      
      // Track that this row needs invoice data written
      rowsToUpdate.push({
        row: sheetRowIndex,
        studentId: lesson.studentId,
        lessonDate: lesson.lessonDate
      });
    }
    
    // WRITE INVOICE DATA IMMEDIATELY (while workbook is still open)
    if (rowsToUpdate.length > 0 && invoiceDateCol && invoiceNumberCol) {
      UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "INFO", 
                                    "Writing invoice data to sheet", 
                                    "Sheet: " + sheet.getName() + ", Rows: " + rowsToUpdate.length, "");
      
      for (var m = 0; m < rowsToUpdate.length; m++) {
        var rowInfo = rowsToUpdate[m];
        
        // Write Invoice Date
        sheet.getRange(rowInfo.row, invoiceDateCol).setValue(formattedInvoiceDate);
        
        // Write Invoice Number
        sheet.getRange(rowInfo.row, invoiceNumberCol).setValue(invoiceNumber);
      }
      
      UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "SUCCESS", 
                                    "Invoice data written", 
                                    "Sheet: " + sheet.getName() + ", Marked: " + rowsToUpdate.length, "");
    } else if (rowsToUpdate.length > 0 && (!invoiceDateCol || !invoiceNumberCol)) {
      UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "WARNING", 
                                    "Cannot write invoice data - columns missing", 
                                    "Sheet: " + sheet.getName() + ", InvoiceDateCol: " + invoiceDateCol + ", InvoiceNumberCol: " + invoiceNumberCol, "");
    }
    
    return lessons;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "ERROR", 
                                  "Failed to extract and mark lessons", 
                                  "Sheet: " + sheet.getName(), error.message);
    throw error;
  }
}

function extractLessonFromRow(row, headerMap, teacherData, cutoffDate, sheet, sheetRowIndex) {
  try {
    // Extract values using headerMap (convert 1-based to 0-based for array access)
    var studentId = row[headerMap['studentid'] - 1];
    var studentName = row[headerMap['studentname'] - 1];
    var dateValue = row[headerMap['date'] - 1];
    var lengthValue = row[headerMap['length'] - 1];
    var statusValue = row[headerMap['status'] - 1];
    var invoiceDateValue = headerMap['invoicedate'] ? row[headerMap['invoicedate'] - 1] : null;
    
    // Skip if essential data is missing
    if (!studentId || !studentName || !dateValue || !statusValue) {
      return null;
    }
    
    // Skip if already invoiced (Invoice Date column has value)
    if (invoiceDateValue && invoiceDateValue.toString().trim() !== '') {
      return null;
    }
    
    // Parse the lesson date
    var lessonDate = parseLessonDate(dateValue);
    if (!lessonDate || isNaN(lessonDate.getTime())) {
      return null; // Invalid date
    }
    
    // Only include lessons on or before cutoff date
    if (lessonDate > cutoffDate) {
      return null;
    }
    
    // Only include payable attendance ("lesson" or "no show")
    var status = statusValue.toString().toLowerCase().trim();
    if (['lesson', 'no show'].indexOf(status) === -1) {
      return null;
    }
    
    // Parse lesson length
    var lessonLength = parseInt(lengthValue) || 0;
    
    // Determine if this is a group session (ID starts with "G")
    var isGroupSession = studentId && studentId.toString().charAt(0) === 'G';
    
    // Validate lesson length based on type
    var isValidLength = false;
    
    if (isGroupSession) {
      // GROUP SESSION: 15-240 minutes, must end in 0 or 5
      if (lessonLength >= 15 && lessonLength <= 240) {
        var lastDigit = lessonLength % 10;
        if (lastDigit === 0 || lastDigit === 5) {
          isValidLength = true;
        }
      }
    } else {
      // STUDENT LESSON: Only 30, 45, or 60
      if (lessonLength === 30 || lessonLength === 45 || lessonLength === 60) {
        isValidLength = true;
      }
    }
    
    // If invalid length, write to Admin Comments and skip row
    if (!isValidLength) {
      if (headerMap['admincomments'] && sheet && sheetRowIndex) {
        var adminCommentsCol = headerMap['admincomments'];
        var currentComments = row[adminCommentsCol - 1];
        var errorMessage = 'Fix lesson length';
        
        // Append error message if comments already exist
        var newComments = currentComments && currentComments.toString().trim() !== '' 
          ? currentComments + ' | ' + errorMessage 
          : errorMessage;
        
        sheet.getRange(sheetRowIndex, adminCommentsCol).setValue(newComments);
        
        UtilityScriptLibrary.debugLog('extractLessonFromRow', 'WARNING', 'Invalid lesson length',
          (isGroupSession ? 'Group' : 'Student') + ' length: ' + lengthValue + ', Row: ' + sheetRowIndex, '');
      }
      return null; // Skip this row
    }
    
    return {
      teacherName: teacherData.teacherName,
      teacherId: teacherData.teacherId,
      studentId: studentId.toString().trim(),
      studentName: studentName.toString().trim(),
      lessonDate: lessonDate,
      status: status === 'no show' ? 'No Show' : 'Lesson',
      lessonLength: lessonLength
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('extractLessonFromRow', 'ERROR', 'Row extraction failed', '', error.message);
    return null;
  }
}

function extractLessonsFromAttendanceSheet(sheet, teacherData, cutoffDate) {
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // Empty sheet or only headers
    }
    
    // Use UtilityScriptLibrary to get header map (returns 1-based column numbers)
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Validate required columns exist
    var requiredColumns = ['studentid', 'studentname', 'date', 'length', 'status'];
    for (var j = 0; j < requiredColumns.length; j++) {
      var col = requiredColumns[j];
      if (!headerMap[col]) {
        throw new Error('Required column "' + col + '" not found in ' + sheet.getName());
      }
    }
    
    var lessons = [];
    
    // Process data rows (skip header row 0)
    for (var k = 1; k < data.length; k++) {
      var sheetRowIndex = k + 1; // Convert 0-based data index to 1-based sheet row
      var lesson = extractLessonFromRow(data[k], headerMap, teacherData, cutoffDate, sheet, sheetRowIndex);
      if (lesson) {
        lessons.push(lesson);
      }
    }
    
    UtilityScriptLibrary.debugLog('extractLessonsFromAttendanceSheet', 'INFO', 'Lessons extracted',
      lessons.length + ' lessons from ' + sheet.getName(), '');
    return lessons;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('extractLessonsFromAttendanceSheet', 'ERROR', 'Extraction failed',
      sheet.getName(), error.message);
    throw error;
  }
}

function extractTeachersFromFormattedSheet(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Get column indices
    var teacherCol = headerMap['teacher'];
    var urlCol = headerMap['url'];
    var pdfCol = headerMap['pdf'];
    var idCol = headerMap['id'];
    var invoiceDateCol = headerMap['invoicedate'];
    var invoiceNumberCol = headerMap['invoicenumber'];
    var invoicePeriodCol = headerMap['invoiceperiod'];
    var addressCol = headerMap['address'];
    var durationCol = headerMap['duration'];
    var lastNameCol = headerMap['lastname'];
    var firstNameCol = headerMap['firstname'];
    var quantityCol = headerMap['quantity'];
    var rateCol = headerMap['rate'];
    var costCol = headerMap['cost'];
    var commentsCol = headerMap['comments'];
    
    var teachers = [];
    var currentTeacher = null;
    
    for (var row = 1; row < data.length; row++) {
      var rowData = data[row];
      var teacherName = rowData[teacherCol - 1];
      var studentId = idCol ? rowData[idCol - 1] : '';
      var invoiceDate = invoiceDateCol ? rowData[invoiceDateCol - 1] : '';
      var invoiceNumber = invoiceNumberCol ? rowData[invoiceNumberCol - 1] : '';
      var duration = durationCol ? rowData[durationCol - 1] : '';
      
      if (!teacherName) continue;
      
      var isTeacherHeader = (studentId && studentId.toString().startsWith('T')) ||
                           (invoiceDate && invoiceDate.toString().trim() !== '') ||
                           (invoiceNumber && invoiceNumber.toString().trim() !== '') ||
                           (!duration || duration.toString().trim() === '');
      
      if (isTeacherHeader) {
        var hasPdfUrl = pdfCol && rowData[pdfCol - 1] && rowData[pdfCol - 1].toString().trim() !== '';
        
        // Only process teachers without a PDF already generated
        if (!hasPdfUrl) {
          var teacherFirstName = firstNameCol ? rowData[firstNameCol - 1] : '';
          var teacherLastName = lastNameCol ? rowData[lastNameCol - 1] : '';
          var teacherAddress = addressCol ? rowData[addressCol - 1] : '';
          var teacherComment = commentsCol ? rowData[commentsCol - 1] : '';
          
          currentTeacher = {
            teacherName: teacherName,
            teacherFirstName: teacherFirstName,
            teacherLastName: teacherLastName,
            teacherAddress: teacherAddress,
            comment: teacherComment,
            headerRowIndex: row + 1,
            invoiceNumber: invoiceNumber,
            invoiceDate: invoiceDate,
            invoicePeriod: invoicePeriodCol ? rowData[invoicePeriodCol - 1] : '',
            studentRows: []
          };
          teachers.push(currentTeacher);
        } else {
          currentTeacher = null; // Skip - PDF already generated
        }
      } else if (currentTeacher) {
        var lastName = lastNameCol ? rowData[lastNameCol - 1] : '';
        var firstName = firstNameCol ? rowData[firstNameCol - 1] : '';
        var quantity = quantityCol ? parseInt(rowData[quantityCol - 1]) || 0 : 0;
        var rate = rateCol ? parseFloat(rowData[rateCol - 1]) || 0 : 0;
        var cost = costCol ? parseFloat(rowData[costCol - 1]) || 0 : 0;
        
        currentTeacher.studentRows.push({
          lastName: lastName || '',
          firstName: firstName || '',
          duration: duration,
          quantity: quantity,
          rate: rate,
          cost: cost
        });
      }
    }
    
    return teachers;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractTeachersFromFormattedSheet", "ERROR", "Failed to extract teacher data", "", error.message);
    throw error;
  }
}

function formatInvoiceSheet(sheet) {
  try {
    UtilityScriptLibrary.debugLog("formatInvoiceSheet", "INFO", "Starting invoice sheet formatting", 
                                  "Sheet: " + sheet.getName(), "");
    
    // Get column mappings directly from utility library
    var columnMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol <= 0) {
      UtilityScriptLibrary.debugLog("formatInvoiceSheet", "WARNING", "No data to format", 
                                    "LastRow: " + lastRow + ", LastCol: " + lastCol, "");
      return;
    }
    
    // Apply basic formatting to data range
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange.setFontFamily("Arial")
             .setFontSize(10)
             .setWrap(true);
    
    // Format currency columns if they exist
    var currencyColumns = ['rate', 'cost'];
    for (var i = 0; i < currencyColumns.length; i++) {
      var colName = currencyColumns[i];
      if (columnMap[colName] !== undefined) {
        var colIndex = columnMap[colName] + 1; // Convert to 1-based
        var currencyRange = sheet.getRange(2, colIndex, lastRow - 1, 1);
        currencyRange.setNumberFormat('$#,##0.00');
      }
    }
    
    // REMOVED: Auto-resize columns - this was overriding template column widths
    // sheet.autoResizeColumns(1, lastCol);
    
    UtilityScriptLibrary.debugLog("formatInvoiceSheet", "INFO", "Invoice sheet formatting completed", "", "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("formatInvoiceSheet", "ERROR", "Failed to format invoice sheet", "", error.message);
    // Don't throw error - formatting failure shouldn't stop invoice generation
  }
}

function generateDefaultInvoicePeriod(cutoffDate) {
  try {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    var month = monthNames[cutoffDate.getMonth()];
    var year = cutoffDate.getFullYear();
    
    return month + ' ' + year;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('generateDefaultInvoicePeriod', 'ERROR', 'Error generating default invoice period', '', error.message);
    return "Unknown Period";
  }
}

function generateDefaultMonthName(cutoffDate) {
  try {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    return monthNames[cutoffDate.getMonth()];
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('generateDefaultMonthName', 'ERROR', 'Error generating default month name', '', error.message);
    return "January"; // Fallback
  }
}

function generateInvoiceNumber(teacherId, invoiceDate) {
  try {
    var year = invoiceDate.getFullYear();
    var month = String(invoiceDate.getMonth() + 1);
    var day = String(invoiceDate.getDate());
    
    // Pad month and day with leading zeros (ES5 compatible)
    if (month.length === 1) month = '0' + month;
    if (day.length === 1) day = '0' + day;
    
    var invoiceNumber = teacherId + '-' + year + month + day;
    
    UtilityScriptLibrary.debugLog("generateInvoiceNumber", "DEBUG", 
                                  "Generated invoice number", 
                                  "Teacher ID: " + teacherId + ", Date: " + UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MM/dd/yyyy') + ", Number: " + invoiceNumber, "");
    
    return invoiceNumber;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateInvoiceNumber", "ERROR", 
                                  "Failed to generate invoice number", 
                                  "Teacher ID: " + teacherId, error.message);
    return teacherId + '-ERROR';
  }
}

function generateMonthlyTeacherInvoices(month, cutoffDate, invoiceDate, lessonResults) {
  try {
    UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "INFO", 
                                  "Starting monthly invoice generation", 
                                  "Month: " + month + ", Cutoff: " + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy'), "");
    
    if (!lessonResults || Object.keys(lessonResults.lessons).length === 0) {
      UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "WARNING", 
                                    "No uninvoiced lessons found", "", "");
      return { success: false, message: "No uninvoiced lessons found" };
    }
    
    var semesterName = getSemesterForDate(cutoffDate);
    var ratesLookup = getRatesLookupForSemester(semesterName);
    
    var invoiceSheet = createMonthlyInvoiceSheet(month);
    
    var populationResult = populateInvoiceSheetFromLessons(
      invoiceSheet, 
      lessonResults.lessons, 
      invoiceDate,
      lessonResults.invoicePeriod,
      ratesLookup
    );
    
    formatInvoiceSheet(invoiceSheet);
    
    UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "SUCCESS", 
                                  "Invoice generation completed", 
                                  "Teachers: " + populationResult.teacherCount + ", Line items: " + populationResult.lineItemCount, "");
    
    return {
      success: true,
      teacherCount: populationResult.teacherCount,
      lineItemCount: populationResult.lineItemCount,
      markedCount: lessonResults.summary.totalLessons
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "ERROR", 
                                  "Invoice generation failed", "", error.message);
    throw error;
  }
}

function generateSingleTeacherInvoice(teacherData, sheet, metadata) {
  try {
    // No need for contact lookup anymore - we have everything from the sheet
    var teacherContactInfo = {}; // Empty object since we're not using it
    
    // Build template variables (pass metadata for date range)
    var variables = buildTeacherInvoiceVariableMap(teacherData, teacherContactInfo, metadata);
    
    // Build filename
    var fileName = buildTeacherInvoiceFileName(teacherData, variables);
    
    // Get destination folder - pass sheet name for monthly organization
    var monthName = sheet.getName();
    var destinationFolder = getTeacherInvoicesFolder(monthName);
    
    // Check if document already exists
    if (UtilityScriptLibrary.documentAlreadyExists(fileName, destinationFolder)) {
      return {
        success: false,
        message: "Document already exists",
        fileName: fileName
      };
    }
    
    // Generate document
    var result = UtilityScriptLibrary.generateDocumentFromTemplate(
      'teacherInvoice',
      variables,
      fileName,
      destinationFolder
    );
    
    if (result.success) {
      // Set the document to view-only (readers can view but not edit)
      try {
        var file = DriveApp.getFileById(result.fileId);
        
        // Remove editor access and set to viewer access for anyone with the link
        // This makes it so teachers can view but not edit
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        UtilityScriptLibrary.debugLog("generateSingleTeacherInvoice", "INFO", 
                                      "Document set to view-only", 
                                      "FileId: " + result.fileId, "");
      } catch (sharingError) {
        // Log but don't fail - this shouldn't break invoice generation
        UtilityScriptLibrary.debugLog("generateSingleTeacherInvoice", "WARNING", 
                                      "Could not set document to view-only", 
                                      "FileId: " + result.fileId, 
                                      sharingError.message);
      }
      
      // Write URL back to teacher header row as hyperlink (text = fileId, link = URL)
      if (teacherData.headerRowIndex) {
        var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
        var urlCol = headerMap['url'];
        if (urlCol) {
          sheet.getRange(teacherData.headerRowIndex, urlCol).setFormula('=HYPERLINK("' + result.url + '", "' + result.fileId + '")');
        }
      }
      
      return {
        success: true,
        fileId: result.fileId,
        url: result.url,
        fileName: fileName
      };
    } else {
      return {
        success: false,
        error: result.error,
        fileName: fileName
      };
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateSingleTeacherInvoice", "ERROR", "Failed to generate single teacher invoice", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    throw error;
  }
}

function generateTeacherInvoiceDocuments() {
  try {
    var ui = SpreadsheetApp.getUi();
    var activeSheet = SpreadsheetApp.getActiveSheet();
    
    // Validate that this is a monthly invoice sheet
    if (!isMonthlyInvoiceSheet(activeSheet)) {
      ui.alert('This does not appear to be a monthly invoice sheet. Please select the correct sheet.');
      return;
    }
    
    var monthName = activeSheet.getName();
    
    // Get metadata for this invoice month
    var metadata;
    try {
      metadata = getTeacherInvoicingMetadata(monthName);
      UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "INFO", 
                                    "Loaded metadata", 
                                    "Month: " + monthName, "");
    } catch (metadataError) {
      UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "WARNING", 
                                    "Could not load metadata (will proceed without)", 
                                    "Month: " + monthName, metadataError.message);
      metadata = null;
    }
    
    // Extract teacher data from the already-formatted sheet
    var teacherData = extractTeachersFromFormattedSheet(activeSheet);
    
    if (teacherData.length === 0) {
      ui.alert('No teachers found that need invoice documents generated.');
      return;
    }
    
    // Generate documents for each teacher
    var results = {
      successful: 0,
      skipped: 0,
      errors: []
    };
    
    for (var i = 0; i < teacherData.length; i++) {
      var teacher = teacherData[i];
      
      try {
        var result = generateSingleTeacherInvoice(teacher, activeSheet, metadata);
        
        if (result.success) {
          results.successful++;
          UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "SUCCESS", 
                                        "Generated invoice for teacher", 
                                        "Teacher: " + teacher.teacherName, "");
          
          // NOTE: Roster update removed - will be done during PDF conversion
          
        } else {
          results.skipped++;
          UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "WARNING", 
                                        "Skipped teacher invoice", 
                                        "Teacher: " + teacher.teacherName + ", Reason: " + result.message, "");
        }
        
      } catch (error) {
        results.errors.push({
          teacher: teacher.teacherName,
          error: error.message
        });
        UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "ERROR", 
                                      "Failed to generate teacher invoice", 
                                      "Teacher: " + teacher.teacherName, error.message);
      }
    }
    
    // Update metadata status to "Generated" if any documents were created
    if (results.successful > 0 && metadata) {
      try {
        updateMetadataStatus(monthName, 'Generated');
      } catch (statusError) {
        UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "WARNING", 
                                      "Could not update metadata status", 
                                      "", statusError.message);
      }
    }
    
    // Show results summary
    showTeacherInvoiceResults(results);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "ERROR", "Teacher invoice generation failed", "", error.message);
    SpreadsheetApp.getUi().alert('Error generating teacher invoices: ' + error.message);
  }
}

function getActiveTeacherList() {
  try {
    UtilityScriptLibrary.debugLog("getActiveTeacherList", "INFO", "Getting active teacher list", "", "");
    
    var formResponsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
    var teacherLookupSheet = formResponsesSS.getSheetByName('Teacher Roster Lookup');
    
    if (!teacherLookupSheet) {
      throw new Error('Teacher Roster Lookup sheet not found in form responses workbook');
    }
    
    var data = teacherLookupSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(teacherLookupSheet);
    
    var firstNameCol   = headerMap[UtilityScriptLibrary.normalizeHeader('First Name')];
    var lastNameCol    = headerMap[UtilityScriptLibrary.normalizeHeader('Last Name')];
    var rosterUrlCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Roster URL')];
    var teacherIdCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    var statusCol      = headerMap[UtilityScriptLibrary.normalizeHeader('Status')];
    
    var activeTeachers = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var firstName = firstNameCol ? row[firstNameCol - 1].toString().trim() : '';
      var lastName  = lastNameCol  ? row[lastNameCol - 1].toString().trim()  : '';
      var rosterUrl = rosterUrlCol ? row[rosterUrlCol - 1].toString().trim() : '';
      var teacherId = teacherIdCol ? row[teacherIdCol - 1].toString().trim() : '';
      var status    = statusCol    ? row[statusCol - 1].toString().trim()    : '';
      var teacherName = (firstName + ' ' + lastName).trim();
      
      if (status === 'Active' && rosterUrl && teacherName) {
        activeTeachers.push({
          teacherName: teacherName,
          firstName: firstName,
          lastName: lastName,
          teacherId: teacherId || teacherName,
          rosterUrl: rosterUrl
        });
      }
    }
    
    UtilityScriptLibrary.debugLog("getActiveTeacherList", "INFO", 
                          "Active teachers found", 
                          "Count: " + activeTeachers.length, "");
    
    return activeTeachers;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getActiveTeacherList", "ERROR", 
                          "Failed to get active teacher list", "", error.message);
    throw error;
  }
}

function getAttendanceSheetsFromWorkbook(spreadsheet) {
  try {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    var allSheets = spreadsheet.getSheets();
    var attendanceSheets = [];
    
    for (var i = 0; i < allSheets.length; i++) {
      var sheet = allSheets[i];
      var sheetName = sheet.getName();
      
      // Check if sheet name matches any month name (case-insensitive)
      for (var j = 0; j < monthNames.length; j++) {
        if (sheetName.toLowerCase() === monthNames[j].toLowerCase()) {
          attendanceSheets.push(sheet);
          UtilityScriptLibrary.debugLog('getAttendanceSheetsFromWorkbook', 'INFO', 'Found attendance sheet', sheet.getName(), '');
          break; // No need to check other months once we find a match
        }
      }
    }
    
    UtilityScriptLibrary.debugLog('getAttendanceSheetsFromWorkbook', 'INFO', 'Attendance sheets found', 
      'Count: ' + attendanceSheets.length, '');
    
    return attendanceSheets;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('getAttendanceSheetsFromWorkbook', 'ERROR', 'Failed to get attendance sheets', '', error.message);
    return [];
  }
}

function getRateForSemester(semesterName, rateType, ratesCache) {
  try {
    UtilityScriptLibrary.debugLog("getRateForSemester", "DEBUG", 
                                  "Looking up rate from cache", 
                                  "Semester: " + semesterName + ", Type: " + rateType, "");
    
    if (!ratesCache[rateType]) {
      throw new Error('Rate type "' + rateType + '" not found in rates cache');
    }
    
    var rate = ratesCache[rateType];
    
    UtilityScriptLibrary.debugLog("getRateForSemester", "INFO", 
                                  "Rate found", 
                                  "Type: " + rateType + ", Rate: " + rate, "");
    
    return rate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getRateForSemester", "ERROR", 
                                  "Rate lookup failed", 
                                  "Semester: " + semesterName + ", Type: " + rateType, error.message);
    throw error;
  }
}

function getRateKeyForProgram(programName, programRateKeysCache) {
  try {
    UtilityScriptLibrary.debugLog("getRateKeyForProgram", "DEBUG", 
                                  "Looking up Rate Key from cache", 
                                  "Program: " + programName, "");
    
    var normalizedProgramName = programName.toLowerCase().trim();
    
    if (!programRateKeysCache[normalizedProgramName]) {
      UtilityScriptLibrary.debugLog("getRateKeyForProgram", "WARNING", 
                                    "Rate Key not found for program", 
                                    "Program: " + programName, "");
      return null;
    }
    
    var rateKey = programRateKeysCache[normalizedProgramName];
    
    UtilityScriptLibrary.debugLog("getRateKeyForProgram", "INFO", 
                                  "Rate Key found", 
                                  "Program: " + programName + ", Rate Key: " + rateKey, "");
    
    return rateKey;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getRateKeyForProgram", "ERROR", 
                                  "Failed to get Rate Key", 
                                  "Program: " + programName, error.message);
    throw error;
  }
}

function getRatesLookupForSemester(semesterName) {
  try {
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var semesterSheet = billingWorkbook.getSheetByName('Semester Metadata');

    if (!semesterSheet) {
      throw new Error('Semester Metadata sheet not found');
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(semesterSheet);
    var semesterNameCol      = headerMap[UtilityScriptLibrary.normalizeHeader('Semester Name')];
    var ratesVerificationCol = headerMap[UtilityScriptLibrary.normalizeHeader('Rates Verification')];

    if (!semesterNameCol || !ratesVerificationCol) {
      throw new Error('Required columns not found in Semester Metadata');
    }

    var data = semesterSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][semesterNameCol - 1] === semesterName) {
        var ratesLookup = data[i][ratesVerificationCol - 1];
        UtilityScriptLibrary.debugLog("getRatesLookupForSemester", "INFO",
                                      "Found rates lookup",
                                      "Semester: " + semesterName + ", Rates: " + ratesLookup, "");
        return ratesLookup || "";
      }
    }

    UtilityScriptLibrary.debugLog("getRatesLookupForSemester", "WARNING",
                                  "Semester not found in metadata",
                                  "Semester: " + semesterName, "");
    return "";

  } catch (error) {
    UtilityScriptLibrary.debugLog("getRatesLookupForSemester", "ERROR",
                                  "Failed to get rates lookup",
                                  "Semester: " + semesterName, error.message);
    return "";
  }
}

function getSemesterForDate(date) {
  try {
    UtilityScriptLibrary.debugLog("getSemesterForDate", "DEBUG",
                                  "Looking up semester for date",
                                  "Date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy'), "");

    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var semesterSheet = billingWorkbook.getSheetByName('Semester Metadata');

    if (!semesterSheet) {
      throw new Error('Semester Metadata sheet not found in billing workbook');
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(semesterSheet);
    var semesterNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Semester Name')];
    var startDateCol    = headerMap[UtilityScriptLibrary.normalizeHeader('Start Date')];
    var endDateCol      = headerMap[UtilityScriptLibrary.normalizeHeader('End Date')];

    if (!semesterNameCol || !startDateCol || !endDateCol) {
      throw new Error('Required columns not found in Semester Metadata');
    }

    var data = semesterSheet.getDataRange().getValues();
    UtilityScriptLibrary.debugLog("getSemesterForDate", "DEBUG",
                                  "Retrieved semester metadata",
                                  "Rows: " + data.length, "");

    for (var i = 1; i < data.length; i++) {
      var semesterName = data[i][semesterNameCol - 1];
      var startDate    = new Date(data[i][startDateCol - 1]);
      var endDate      = new Date(data[i][endDateCol - 1]);

      UtilityScriptLibrary.debugLog("getSemesterForDate", "DEBUG",
                                    "Checking semester",
                                    "Semester: " + semesterName + ", Start: " + UtilityScriptLibrary.formatDateFlexible(startDate, 'MMM d, yyyy') + ", End: " + UtilityScriptLibrary.formatDateFlexible(endDate, 'MMM d, yyyy'), "");

      if (date >= startDate && date <= endDate) {
        UtilityScriptLibrary.debugLog("getSemesterForDate", "INFO",
                                      "Found matching semester",
                                      "Date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy') + ", Semester: " + semesterName, "");
        return semesterName;
      }
    }

    throw new Error("No semester found for date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy'));

  } catch (error) {
    UtilityScriptLibrary.debugLog("getSemesterForDate", "ERROR",
                                  "Failed to get semester for date",
                                  "Date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy'), error.message);
    throw error;
  }
}

function getTeacherContactInfo(teacherId) {
  try {
    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "DEBUG", "Looking up contact info", "Teacher ID: " + teacherId, "");

    if (!teacherId || teacherId.toString().trim() === '') {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "WARNING", "No teacher ID provided", "", "");
      return { firstName: '', lastName: '', address: '' };
    }

    var sheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
    if (!sheet) {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", "Teachers and Admin sheet not found", "", "");
      return { firstName: '', lastName: '', address: '' };
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var idCol        = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    var firstNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('First Name')];
    var lastNameCol  = headerMap[UtilityScriptLibrary.normalizeHeader('Last Name')];
    var addressCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Address')];

    if (!idCol) {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", "Teacher ID column not found", "", "");
      return { firstName: '', lastName: '', address: '' };
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol - 1]).trim() === String(teacherId).trim()) {
        var firstName = firstNameCol ? String(data[i][firstNameCol - 1] || '').trim() : '';
        var lastName  = lastNameCol  ? String(data[i][lastNameCol - 1]  || '').trim() : '';
        var address   = addressCol   ? String(data[i][addressCol - 1]   || '').trim() : '';

        UtilityScriptLibrary.debugLog("getTeacherContactInfo", "DEBUG", "Found teacher contact info",
                                      "ID: " + teacherId + ", First: " + firstName + ", Last: " + lastName, "");
        return { firstName: firstName, lastName: lastName, address: address };
      }
    }

    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "WARNING", "Teacher not found in contacts", "ID: " + teacherId, "");
    return { firstName: '', lastName: '', address: '' };

  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", "Failed to get teacher contact info", "ID: " + teacherId, error.message);
    return { firstName: '', lastName: '', address: '' };
  }
}
 
function getTeacherInfoByName(teacherName) {
  try {
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "INFO", "Looking up teacher info", 
                 "Teacher: " + teacherName, "");
    
    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog("getTeacherInfoByName", "WARNING", "Lookup sheet not found or empty", "", "");
      return null;
    }
    
    var data = lookupSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(lookupSheet);

    var firstNameCol   = headerMap[UtilityScriptLibrary.normalizeHeader('First Name')];
    var lastNameCol    = headerMap[UtilityScriptLibrary.normalizeHeader('Last Name')];
    var rosterUrlCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Roster URL')];
    var teacherIdCol   = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    var displayNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Display Name')];
    var statusCol      = headerMap[UtilityScriptLibrary.normalizeHeader('Status')];
    var lastUpdatedCol = headerMap[UtilityScriptLibrary.normalizeHeader('Last Updated')];

    var searchName = String(teacherName).trim();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var firstName    = firstNameCol   ? row[firstNameCol - 1].toString().trim()   : '';
      var lastName     = lastNameCol    ? row[lastNameCol - 1].toString().trim()    : '';
      var rowFullName  = (firstName + ' ' + lastName).trim();

      // Try full name exact match, then last name only
      if (rowFullName === searchName || lastName === searchName) {
        var teacherInfo = {
          teacherName:  rowFullName,
          firstName:    firstName,
          lastName:     lastName,
          rosterUrl:    rosterUrlCol   ? row[rosterUrlCol - 1].toString().trim()   : '',
          teacherId:    teacherIdCol   ? row[teacherIdCol - 1].toString().trim()   : '',
          displayName:  displayNameCol ? row[displayNameCol - 1].toString().trim() : '',
          status:       statusCol      ? row[statusCol - 1].toString().trim()      : '',
          lastUpdated:  lastUpdatedCol ? row[lastUpdatedCol - 1]                   : ''
        };

        UtilityScriptLibrary.debugLog("getTeacherInfoByName", "SUCCESS", 
                     rowFullName === searchName ? "Found exact match" : "Found last name match", 
                     "Searched: " + searchName + ", Found: " + rowFullName, "");
        return teacherInfo;
      }
    }
    
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "WARNING", "Teacher not found", 
                 "Teacher: " + searchName, "");
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "ERROR", "Lookup failed", 
                 "Teacher: " + teacherName, error.message);
    return null;
  }
}

function getTeacherInfoFromLessonGroup(lessonGroup) {
  try {
    var teacherName = lessonGroup.teacherName;
    var teacherId = lessonGroup.teacherId;
    
    // Get teacher ID from Teacher Roster Lookup if not available
    var finalTeacherId = teacherId;
    if (!finalTeacherId) {
      // For now, just use empty string - we can improve this later
      finalTeacherId = '';
    }
    
    // Get comprehensive teacher contact info (firstName, lastName, address)
    var teacherContactInfo = getTeacherContactInfo(finalTeacherId);
    
    // Parse last name from teacher name if contact lookup didn't provide it
    var lastName = teacherContactInfo.lastName || teacherName;
    if (!teacherContactInfo.lastName && teacherName.indexOf(',') !== -1) {
      var nameParts = teacherName.split(',');
      lastName = nameParts[0] ? nameParts[0].trim() : teacherName;
    }
    
    UtilityScriptLibrary.debugLog("getTeacherInfoFromLessonGroup", "DEBUG", 
                                  "Teacher info compiled", 
                                  "Name: " + teacherName + ", ID: " + finalTeacherId, "");
    
    return {
      teacherName: teacherName,
      teacherId: finalTeacherId,
      lastName: lastName,
      firstName: teacherContactInfo.firstName,
      address: teacherContactInfo.address
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInfoFromLessonGroup", "ERROR", 
                                  "Failed to get teacher info", 
                                  "Teacher: " + lessonGroup.teacherName, error.message);
    return {
      teacherName: lessonGroup.teacherName,
      teacherId: '',
      lastName: lessonGroup.teacherName,
      firstName: '',
      address: ''
    };
  }
}

function getTeacherInvoicesFolder(monthName) {
  try {
    var mainFolder = UtilityScriptLibrary.getGeneratedDocumentsFolder();
    var teacherFolders = mainFolder.getFoldersByName("Teacher Invoices");
    
    if (!teacherFolders.hasNext()) {
      throw new Error("Teacher Invoices folder not found. Please create it manually in the generated documents folder.");
    }
    
    var teacherInvoicesFolder = teacherFolders.next();
    
    // Create or get the monthly subfolder
    var monthlyFolders = teacherInvoicesFolder.getFoldersByName(monthName);
    
    if (monthlyFolders.hasNext()) {
      UtilityScriptLibrary.debugLog("getTeacherInvoicesFolder", "INFO", "Using existing monthly folder", 
                    "Month: " + monthName, "");
      return monthlyFolders.next();
    } else {
      var newMonthFolder = teacherInvoicesFolder.createFolder(monthName);
      UtilityScriptLibrary.debugLog("getTeacherInvoicesFolder", "INFO", "Created new monthly folder", 
                    "Month: " + monthName, "");
      return newMonthFolder;
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInvoicesFolder", "ERROR", "Failed to get teacher invoices folder", 
                  "Month: " + monthName, error.message);
    throw error;
  }
}

function getTeacherInvoicingMetadata(invoiceMonth) {
  try {
    UtilityScriptLibrary.debugLog("getTeacherInvoicingMetadata", "INFO", 
                                  "Reading metadata", 
                                  "Month: " + invoiceMonth, "");
    
    var teacherInvoicesSS = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = teacherInvoicesSS.getSheetByName('Teacher Invoicing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Teacher Invoicing Metadata sheet not found');
    }
    
    var data = metadataSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    
    // Find row matching the invoice month
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var monthCol = headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Month')];
      
      if (row[monthCol - 1] === invoiceMonth) {
        var metadata = {
          invoiceMonth: row[headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Month')] - 1],
          lessonsStartingDate: row[headerMap[UtilityScriptLibrary.normalizeHeader('Lessons Starting Date')] - 1],
          lessonsEndingDate: row[headerMap[UtilityScriptLibrary.normalizeHeader('Lessons Ending Date')] - 1],
          invoiceDate: row[headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Date')] - 1],
          ratesLookup: row[headerMap[UtilityScriptLibrary.normalizeHeader('Rates Lookup')] - 1] || '',
          semesterName: row[headerMap[UtilityScriptLibrary.normalizeHeader('Semester Name')] - 1] || '',
          status: row[headerMap[UtilityScriptLibrary.normalizeHeader('Status')] - 1] || ''
        };
        
        UtilityScriptLibrary.debugLog("getTeacherInvoicingMetadata", "SUCCESS", 
                                      "Found metadata", 
                                      "Month: " + invoiceMonth + ", Semester: " + metadata.semesterName, "");
        
        return metadata;
      }
    }
    
    throw new Error("Metadata not found for month: " + invoiceMonth);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInvoicingMetadata", "ERROR", 
                                  "Failed to read metadata", 
                                  "Month: " + invoiceMonth, error.message);
    throw error;
  }
}

function getUninvoicedLessonsForTeacher(teacherData, cutoffDate, invoiceDate, invoiceNumber) {
  try {
    UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "INFO", 
                                  "Starting single-pass lesson collection and marking", 
                                  "Teacher: " + teacherData.teacherName + ", Cutoff: " + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy'), "");
    
    var rosterUrl = teacherData.rosterUrl;
    if (!rosterUrl) {
      throw new Error('No roster URL found for ' + teacherData.teacherName);
    }
    
    var fileIdMatch = rosterUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!fileIdMatch) {
      throw new Error('Invalid roster URL format for ' + teacherData.teacherName);
    }
    
    var rosterSS = SpreadsheetApp.openById(fileIdMatch[1]);
    var attendanceSheets = getAttendanceSheetsFromWorkbook(rosterSS);
    
    if (attendanceSheets.length === 0) {
      UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "WARNING", 
                                    "No attendance sheets found", 
                                    "Teacher: " + teacherData.teacherName, "");
      return [];
    }
    
    var uninvoicedLessons = [];
    var formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(invoiceDate, "MM/dd/yyyy");
    
    for (var i = 0; i < attendanceSheets.length; i++) {
      var sheet = attendanceSheets[i];
      try {
        var sheetLessons = extractAndMarkLessonsFromSheet(
          sheet, 
          teacherData, 
          cutoffDate, 
          formattedInvoiceDate, 
          invoiceNumber
        );
        uninvoicedLessons = uninvoicedLessons.concat(sheetLessons);
        
        UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "INFO", 
                                      "Processed sheet", 
                                      "Sheet: " + sheet.getName() + ", Lessons: " + sheetLessons.length, "");
      } catch (sheetError) {
        UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "ERROR", 
                                      "Error processing sheet", 
                                      "Teacher: " + teacherData.teacherName + ", Sheet: " + sheet.getName(), 
                                      sheetError.message);
      }
    }
    
    UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "SUCCESS", 
                                  "Completed single-pass processing", 
                                  "Teacher: " + teacherData.teacherName + ", Total lessons: " + uninvoicedLessons.length, "");
    
    return uninvoicedLessons;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "ERROR", 
                                  "Failed to process teacher", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    throw error;
  }
}

function groupLessonsByTeacherAndType(allLessons) {
  try {
    var grouped = {};
    
    for (var i = 0; i < allLessons.length; i++) {
      var lesson = allLessons[i];
      // Create grouping key: teacher + student + lesson length (one line per student+length)
      var groupKey = lesson.teacherName + '|' + lesson.studentId + '|' + lesson.lessonLength;
      
      if (!grouped[groupKey]) {
        grouped[groupKey] = {
          teacherName: lesson.teacherName,
          teacherId: lesson.teacherId,
          studentId: lesson.studentId,
          studentName: lesson.studentName,
          lessonLength: lesson.lessonLength,
          quantity: 0,
          lessons: []
        };
      }
      
      // Add lesson to group and increment quantity
      grouped[groupKey].quantity++;
      grouped[groupKey].lessons.push(lesson);
    }
    
    UtilityScriptLibrary.debugLog('groupLessonsByTeacherAndType', 'INFO', 'Lessons grouped',
      allLessons.length + ' lessons into ' + Object.keys(grouped).length + ' line items', '');
    return grouped;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('groupLessonsByTeacherAndType', 'ERROR', 'Error grouping lessons', '', error.message);
    throw error;
  }
}

function isMonthlyInvoiceSheet(sheet) {
  try {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var expectedHeaders = ['Last Name', 'First Name', 'Teacher', 'Rate', 'Cost', 'URL'];
    
    var foundCount = 0;
    for (var i = 0; i < expectedHeaders.length; i++) {
      for (var j = 0; j < headers.length; j++) {
        if (UtilityScriptLibrary.normalizeHeader(headers[j]) === UtilityScriptLibrary.normalizeHeader(expectedHeaders[i])) {
          foundCount++;
          break;
        }
      }
    }
    
    return foundCount >= 4; // Must have at least 4 key headers
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("isMonthlyInvoiceSheet", "ERROR", "Error validating sheet", "", error.message);
    return false;
  }
}

function loadProgramRateKeysCache() {
  try {
    UtilityScriptLibrary.debugLog("loadProgramRateKeysCache", "INFO", "Loading program rate keys cache", "", "");

    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var programsSheet = billingWorkbook.getSheetByName('Programs List');

    if (!programsSheet) {
      throw new Error('Programs List sheet not found in billing workbook');
    }

    var data = programsSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('No data found in Programs List sheet');
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(programsSheet);
    var programNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Program Name')];
    var rateKeyCol     = headerMap[UtilityScriptLibrary.normalizeHeader('Rate Key')];

    if (!programNameCol || !rateKeyCol) {
      throw new Error('Required columns not found in Programs List (Program Name, Rate Key)');
    }

    var programRateKeysCache = {};

    for (var i = 1; i < data.length; i++) {
      var programName = data[i][programNameCol - 1] ? data[i][programNameCol - 1].toString().toLowerCase().trim() : '';
      var rateKey     = data[i][rateKeyCol - 1]     ? data[i][rateKeyCol - 1].toString().trim()                   : '';

      if (programName && rateKey) {
        programRateKeysCache[programName] = rateKey;
      }
    }

    UtilityScriptLibrary.debugLog("loadProgramRateKeysCache", "SUCCESS", "Program rate keys cache loaded",
                                  "Total programs: " + Object.keys(programRateKeysCache).length, "");
    return programRateKeysCache;

  } catch (error) {
    UtilityScriptLibrary.debugLog("loadProgramRateKeysCache", "ERROR", "Failed to load program rate keys cache", "", error.message);
    throw error;
  }
}

function loadRatesCache(ratesColumnHeader) {
  try {
    UtilityScriptLibrary.debugLog("loadRatesCache", "INFO", "Loading rates cache", 
                                  "Column: " + ratesColumnHeader, "");
    
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var ratesSheet = billingWorkbook.getSheetByName('Rates');
    
    if (!ratesSheet) {
      throw new Error('Rates sheet not found in billing workbook');
    }
    
    var headers = ratesSheet.getRange(1, 1, 1, ratesSheet.getLastColumn()).getValues()[0];
    var colIndex = headers.indexOf(ratesColumnHeader);
    
    if (colIndex === -1) {
      throw new Error('Rates column "' + ratesColumnHeader + '" not found in Rates sheet');
    }
    
    var ratesCache = UtilityScriptLibrary.buildRateMapFromSheet(ratesSheet, headers, colIndex);
    
    UtilityScriptLibrary.debugLog("loadRatesCache", "SUCCESS", "Rates cache loaded", 
                                  "Column: " + ratesColumnHeader + ", Total rates: " + Object.keys(ratesCache).length, "");
    
    return ratesCache;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("loadRatesCache", "ERROR", "Failed to load rates cache", "", error.message);
    throw error;
  }
}

function logSheetHeaders() {
  UtilityScriptLibrary.logAllSheetHeaders();
}

function normalizeNameForMatching(name) {
  if (!name) return '';
  
  // Remove accents/diacritics
  var normalized = name.toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
  
  return normalized;
}

function parseLessonDate(dateValue) {
  try {
    // Handle different date formats
    if (dateValue instanceof Date) {
      return dateValue;
    }
        
    if (!dateValue) {
      return null;
    }
        
    var dateStr = dateValue.toString().trim();
        
    // Handle "Fri, 1/5" format - extract just the "1/5" part
    var dayDateMatch = dateStr.match(/\w+,\s*(\d{1,2}\/\d{1,2})/);
    if (dayDateMatch) {
      var dateOnly = dayDateMatch[1]; // e.g., "1/5"
      var currentYear = new Date().getFullYear();
      var fullDateStr = dateOnly + "/" + currentYear; // e.g., "1/5/2024"
            
      // Use UtilityScriptLibrary to parse
      var parsedDate = UtilityScriptLibrary.parseDateFromString(fullDateStr);
      if (parsedDate) {
        return parsedDate;
      }
    }
        
    // Try direct parsing with UtilityScriptLibrary
    var directParse = UtilityScriptLibrary.parseDateFromString(dateStr);
    if (directParse) {
      return directParse;
    }
        
    // Fallback to native Date constructor
    var nativeDate = new Date(dateStr);
    if (!isNaN(nativeDate.getTime())) {
      return nativeDate;
    }
        
    return null;
      
  } catch (error) {
    UtilityScriptLibrary.debugLog('parseLessonDate', 'ERROR', 'Date parsing failed', 'Input: ' + dateValue, error.message);
    return null;
  }
}

function parseStudentName(studentName) {
  try {
    if (!studentName) {
      return { firstName: '', lastName: '' };
    }
    
    var cleanName = studentName;
    
    // Remove instrument suffix if present (e.g., "Cacacho, Isaac - Cello" -> "Cacacho, Isaac")
    var dashIndex = cleanName.indexOf(' - ');
    if (dashIndex !== -1) {
      cleanName = cleanName.substring(0, dashIndex);
    }
    
    // Check if this looks like a group name (no comma means it's not "Last, First" format)
    if (cleanName.indexOf(',') === -1) {
      // This is a group name - return it as lastName only
      UtilityScriptLibrary.debugLog("parseStudentName", "DEBUG", "Detected group name", 
                                    "Input: '" + studentName + "' -> Group: '" + cleanName + "'", "");
      return { firstName: '', lastName: cleanName };
    }
    
    // Parse "Last, First" format for student names
    var nameParts = cleanName.split(',');
    var lastName = nameParts[0] ? nameParts[0].trim() : cleanName;
    var firstName = nameParts[1] ? nameParts[1].trim() : '';
    
    UtilityScriptLibrary.debugLog("parseStudentName", "DEBUG", "Parsed student name", 
                                  "Input: '" + studentName + "' -> Last: '" + lastName + "', First: '" + firstName + "'", "");
    
    return {
      firstName: firstName,
      lastName: lastName
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("parseStudentName", "ERROR", "Failed to parse student name", 
                                  "Input: " + studentName, error.message);
    return { firstName: '', lastName: studentName || '' };
  }
}

function populateInvoiceSheetFromLessons(sheet, groupedLessons, invoiceDate, invoicePeriod, ratesColumnHeader) {
  try {
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "INFO", 
                                  "Starting invoice sheet population", 
                                  "Line items: " + Object.keys(groupedLessons).length, "");
    
    // Load rates and program rate keys once
    var ratesCache = loadRatesCache(ratesColumnHeader);
    var programRateKeysCache = loadProgramRateKeysCache();
    
    // Get column mappings directly from utility library
    var columnMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Log available columns for debugging
    var availableColumns = Object.keys(columnMap);
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "DEBUG", "Available columns", 
                                  "Columns: " + availableColumns.join(', '), "");
    
    var currentRow = 2; // Start after header
    var teacherCount = 0;
    var lineItemCount = 0;
    var currentTeacher = null;
    
    // Process each lesson group (student+duration combination)
    for (var key in groupedLessons) {
      if (groupedLessons.hasOwnProperty(key)) {
        var lessonGroup = groupedLessons[key];
        
        // If this is a new teacher, add teacher header row
        if (lessonGroup.teacherName !== currentTeacher) {
          // Add blank separator row (except for first teacher)
          if (currentTeacher !== null) {
            currentRow++;
          }
          
          teacherCount++;
          currentTeacher = lessonGroup.teacherName;
          
          // Get teacher information
          var teacherInfo = getTeacherInfoFromLessonGroup(lessonGroup);
          
          // Generate invoice number
          var invoiceNumber = generateInvoiceNumber(teacherInfo.teacherId, invoiceDate);
          
          // Add teacher header row
          currentRow = addTeacherHeaderRow(
            sheet, currentRow, columnMap, teacherInfo, 
            invoiceDate, invoiceNumber, invoicePeriod
          );
        }
        
        // Add student line item row with caches
        currentRow = addStudentLineItem(
          sheet, currentRow, columnMap, lessonGroup, currentTeacher, 
          ratesCache, programRateKeysCache
        );
        lineItemCount++;
      }
    }
    
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "INFO", 
                                  "Invoice sheet population completed", 
                                  "Teachers: " + teacherCount + ", Line items: " + lineItemCount, "");
    
    return {
      teacherCount: teacherCount,
      lineItemCount: lineItemCount
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "ERROR", 
                                  "Invoice sheet population failed", "", error.message);
    throw error;
  }
}

function processAttendanceSheet(sheet, teacherLastName, sheetName, unpaidLessons, attendanceData) {
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Check for required columns
    if (!headerMap['studentid'] || !headerMap['date'] || !headerMap['length'] || !headerMap['status']) {
      return;
    }
    
    var studentIdCol = headerMap['studentid'] - 1;
    var dateCol = headerMap['date'] - 1;
    var lengthCol = headerMap['length'] - 1;
    var statusCol = headerMap['status'] - 1;
    var invoiceDateCol = headerMap['invoicedate'] ? headerMap['invoicedate'] - 1 : -1;
    
    // Normalize teacher name for matching
    var teacherKey = normalizeNameForMatching(teacherLastName);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var studentId = row[studentIdCol];
      var dateValue = row[dateCol];
      var lengthValue = row[lengthCol];
      var statusValue = row[statusCol];
      var invoiceDateValue = invoiceDateCol !== -1 ? row[invoiceDateCol] : null;
      
      // Skip if missing essential data
      if (!studentId || !dateValue || !statusValue) continue;
      
      // Parse status
      var status = statusValue.toString().toLowerCase().trim();
      if (status !== 'lesson' && status !== 'no show') continue;
      
      // Parse length
      var length = parseFloat(lengthValue) || 0;
      if (length <= 0) continue;
      
      var hours = length / 60;
      var studentIdStr = studentId.toString().trim();
      
      // Check if Invoice Date is filled
      if (!invoiceDateValue || invoiceDateValue.toString().trim() === '') {
        // UNPAID LESSON
        unpaidLessons.push({
          teacher: teacherLastName, // Keep original name for display
          sheet: sheetName,
          studentId: studentIdStr,
          date: dateValue,
          hours: hours
        });
      } else {
        // Has Invoice Date - add to attendance data
        var invoiceDate = UtilityScriptLibrary.formatDateFlexible(
          invoiceDateValue instanceof Date ? invoiceDateValue : new Date(invoiceDateValue), 
          'MM/dd/yyyy'
        );
        
        if (!attendanceData[teacherKey]) {
          attendanceData[teacherKey] = {};
        }
        if (!attendanceData[teacherKey][invoiceDate]) {
          attendanceData[teacherKey][invoiceDate] = {};
        }
        if (!attendanceData[teacherKey][invoiceDate][studentIdStr]) {
          attendanceData[teacherKey][invoiceDate][studentIdStr] = 0;
        }
        
        attendanceData[teacherKey][invoiceDate][studentIdStr] += hours;
      }
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('processAttendanceSheet', 'ERROR', 
                                  'Failed', 
                                  teacherLastName + ' - ' + sheetName, error.message);
  }
}

function processInvoiceSheetForVerification(sheet, invoiceData) {
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Required columns
    if (!headerMap['teacher'] || !headerMap['id'] || !headerMap['duration'] || !headerMap['quantity']) {
      return;
    }
    
    var teacherCol = headerMap['teacher'] - 1;
    var idCol = headerMap['id'] - 1;
    var durationCol = headerMap['duration'] - 1;
    var quantityCol = headerMap['quantity'] - 1;
    var invoiceDateCol = headerMap['invoicedate'] ? headerMap['invoicedate'] - 1 : -1;
    
    var currentTeacher = null;
    var currentInvoiceDate = null;
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      var teacherName = row[teacherCol];
      var studentId = row[idCol];
      var duration = row[durationCol];
      var quantity = row[quantityCol];
      var invoiceDateValue = invoiceDateCol !== -1 ? row[invoiceDateCol] : null;
      
      if (!teacherName) continue;
      
      var teacherStr = teacherName.toString().trim();
      
      // Extract last name from teacher name
      var nameParts = teacherStr.split(' ');
      var teacherLastName = nameParts[nameParts.length - 1];
      
      // Check if this is a teacher header row
      var isTeacherHeader = (studentId && studentId.toString().startsWith('T')) ||
                           (invoiceDateValue && invoiceDateValue.toString().trim() !== '');
      
      if (isTeacherHeader) {
        // Normalize teacher name for matching
        currentTeacher = normalizeNameForMatching(teacherLastName);
        if (invoiceDateValue) {
          currentInvoiceDate = UtilityScriptLibrary.formatDateFlexible(
            invoiceDateValue instanceof Date ? invoiceDateValue : new Date(invoiceDateValue),
            'MM/dd/yyyy'
          );
        }
      } else if (currentTeacher && currentInvoiceDate && studentId) {
        // Student row
        var studentIdStr = studentId.toString().trim();
        
        // Skip if ID starts with G (group) or T (teacher)
        if (studentIdStr.charAt(0) === 'G' || studentIdStr.charAt(0) === 'T') {
          continue;
        }
        
        var dur = parseFloat(duration) || 0;
        var qty = parseFloat(quantity) || 0;
        
        // Duration is in minutes, convert to hours
        var hours = (dur / 60) * qty;
        
        if (hours > 0) {
          if (!invoiceData[currentTeacher]) {
            invoiceData[currentTeacher] = {};
          }
          if (!invoiceData[currentTeacher][currentInvoiceDate]) {
            invoiceData[currentTeacher][currentInvoiceDate] = {};
          }
          if (!invoiceData[currentTeacher][currentInvoiceDate][studentIdStr]) {
            invoiceData[currentTeacher][currentInvoiceDate][studentIdStr] = 0;
          }
          
          invoiceData[currentTeacher][currentInvoiceDate][studentIdStr] += hours;
        }
      }
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('processInvoiceSheetForVerification', 'ERROR', 
                                  'Failed', sheet.getName(), error.message);
  }
}

function promptForCutoffDate() {
  try {
    UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'Starting cutoff date prompt', '', '');
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Lesson Cutoff Date',
      'Enter the cutoff date for lessons to include (MM/DD/YYYY).\n\n' +
      'All lessons on or before this date will be included in the invoice.\n\n' +
      'Example: 01/31/2024',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'User response received',
      'Button: ' + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"', '');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'User cancelled cutoff date prompt', '', '');
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'Processing user input', userInput, '');
    
    var cutoffDate;
    try {
      cutoffDate = UtilityScriptLibrary.parseDateFromString(userInput);
      UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'parseDateFromString result',
        cutoffDate ? UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy') : 'null', '');
    } catch (utilityError) {
      UtilityScriptLibrary.debugLog('promptForCutoffDate', 'WARNING', 'parseDateFromString failed',
        userInput, utilityError.message);
      cutoffDate = null;
    }
    
    if (!cutoffDate) {
      UtilityScriptLibrary.debugLog('promptForCutoffDate', 'WARNING', 'Falling back to native Date constructor', '', '');
      
      try {
        cutoffDate = new Date(userInput);
        if (isNaN(cutoffDate.getTime())) {
          cutoffDate = null;
        }
        UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'Native Date constructor result',
          cutoffDate ? UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy') : 'null', '');
      } catch (nativeError) {
        UtilityScriptLibrary.debugLog('promptForCutoffDate', 'ERROR', 'Native Date constructor failed',
          userInput, nativeError.message);
        cutoffDate = null;
      }
    }
    
    if (!cutoffDate) {
      var errorMsg = '❌ Invalid date format. Please use MM/DD/YYYY format (example: 01/31/2024).';
      ui.alert(errorMsg);
      UtilityScriptLibrary.debugLog('promptForCutoffDate', 'ERROR', 'Date parsing completely failed', userInput, '');
      return null;
    }
    
    UtilityScriptLibrary.debugLog('promptForCutoffDate', 'INFO', 'Successfully parsed cutoff date',
      'Input: ' + userInput + ', Result: ' + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy'), '');
    
    return cutoffDate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('promptForCutoffDate', 'ERROR', 'Unexpected error', '', error.message);
    return null;
  }
}

function promptForInvoiceDate() {
  try {
    UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'Starting invoice date prompt', '', '');
    
    var ui = SpreadsheetApp.getUi();
    var today = new Date();
    var defaultDateStr = UtilityScriptLibrary.formatDateFlexible(today, 'MM/dd/yyyy');
    
    UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'Default date calculated',
      'Today: ' + UtilityScriptLibrary.formatDateFlexible(today, 'M/d/yy') + ', Formatted: ' + defaultDateStr, '');
    
    var response = ui.prompt(
      'Invoice Date',
      'Enter the invoice date to mark these lessons (MM/DD/YYYY).\n\n' +
      'This date will be stamped on all processed lessons.\n\n' +
      'Default (today): ' + defaultDateStr + '\n' +
      'Press OK to use default, or enter a different date:',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'User response received',
      'Button: ' + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"', '');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'User cancelled invoice date prompt', '', '');
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    
    var invoiceDate;
    if (!userInput) {
      invoiceDate = today;
      UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'Using default date (today)', UtilityScriptLibrary.formatDateFlexible(today, 'M/d/yy'), '');
    } else {
      UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'Parsing user input', userInput, '');
      
      try {
        invoiceDate = UtilityScriptLibrary.parseDateFromString(userInput);
        UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'parseDateFromString result',
          invoiceDate ? UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'M/d/yy') : 'null', '');
      } catch (utilityError) {
        UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'WARNING', 'parseDateFromString failed',
          userInput, utilityError.message);
        invoiceDate = null;
      }
      
      if (!invoiceDate) {
        try {
          invoiceDate = new Date(userInput);
          if (isNaN(invoiceDate.getTime())) {
            invoiceDate = null;
          }
          UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'Native Date constructor result',
            invoiceDate ? UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'M/d/yy') : 'null', '');
        } catch (nativeError) {
          UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'ERROR', 'Native Date constructor failed',
            userInput, nativeError.message);
          invoiceDate = null;
        }
      }
      
      if (!invoiceDate) {
        var errorMsg = '❌ Invalid date format. Please use MM/DD/YYYY format (example: 01/31/2024).';
        ui.alert(errorMsg);
        UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'ERROR', 'Date parsing failed', userInput, '');
        return null;
      }
    }
    
    UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'INFO', 'Successfully determined invoice date',
      UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'M/d/yy'), '');
    
    return invoiceDate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('promptForInvoiceDate', 'ERROR', 'Unexpected error', '', error.message);
    return null;
  }
}

function promptForInvoicePeriod(cutoffDate) {
  try {
    UtilityScriptLibrary.debugLog('promptForInvoicePeriod', 'INFO', 'Starting invoice period prompt', '', '');
    
    var ui = SpreadsheetApp.getUi();
    var defaultPeriod = generateDefaultInvoicePeriod(cutoffDate);
    
    UtilityScriptLibrary.debugLog('promptForInvoicePeriod', 'INFO', 'Default period calculated',
      'Cutoff: ' + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy') + ', Default period: ' + defaultPeriod, '');
    
    var response = ui.prompt(
      'Invoice Period',
      'Invoice period description.\n\n' +
      'Default (based on cutoff date): ' + defaultPeriod + '\n\n' +
      'You can override with custom text like:\n' +
      '• "Summer 2025"\n' +
      '• "October-November 2024"\n' +
      '• "Q1 2024"\n\n' +
      'Press OK to use default, or enter custom period:',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog('promptForInvoicePeriod', 'INFO', 'User response received',
      'Button: ' + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"', '');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('promptForInvoicePeriod', 'INFO', 'User cancelled invoice period prompt', '', '');
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    var finalPeriod = userInput || defaultPeriod;
    
    UtilityScriptLibrary.debugLog('promptForInvoicePeriod', 'INFO', 'Invoice period determined',
      'User input: ' + userInput + ', Final period: ' + finalPeriod, '');
    
    return finalPeriod;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('promptForInvoicePeriod', 'ERROR', 'Unexpected error', '', error.message);
    return generateDefaultInvoicePeriod(cutoffDate);
  }
}

function promptForMonthName(cutoffDate) {
  try {
    UtilityScriptLibrary.debugLog('promptForMonthName', 'INFO', 'Starting month name prompt', '', '');
    
    var defaultMonth = generateDefaultMonthName(cutoffDate);
    
    UtilityScriptLibrary.debugLog('promptForMonthName', 'INFO', 'Default month calculated',
      'Cutoff: ' + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy') + ', Default month: ' + defaultMonth, '');
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Invoice Sheet Month',
      'Enter the month name for the invoice sheet.\n\n' +
      'Default (based on cutoff date): ' + defaultMonth + '\n\n' +
      'Examples:\n' +
      '• January\n' +
      '• February\n' +
      '• March\n\n' +
      'Press OK to use default, or enter a different month:',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog('promptForMonthName', 'INFO', 'User response received',
      'Button: ' + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"', '');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('promptForMonthName', 'INFO', 'User cancelled month name prompt', '', '');
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    
    var finalMonth;
    if (!userInput) {
      finalMonth = defaultMonth;
      UtilityScriptLibrary.debugLog('promptForMonthName', 'INFO', 'Using default month', finalMonth, '');
    } else {
      finalMonth = userInput.charAt(0).toUpperCase() + userInput.slice(1).toLowerCase();
      UtilityScriptLibrary.debugLog('promptForMonthName', 'INFO', 'Using custom month',
        'Input: ' + userInput + ', Final: ' + finalMonth, '');
    }
    
    return finalMonth;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('promptForMonthName', 'ERROR', 'Unexpected error', '', error.message);
    return null;
  }
}

function showCombinedErrorDetails(lessonResults, invoiceResults) {
  try {
    var ui = SpreadsheetApp.getUi();
    var detailsMessage = '❌ ERRORS AND ISSUES:\n\n';
    
    // Show lesson collection errors
    if (lessonResults.errors.length > 0) {
      detailsMessage += '🚫 LESSON COLLECTION ERRORS (' + lessonResults.errors.length + '):\n';
      for (var i = 0; i < lessonResults.errors.length; i++) {
        var error = lessonResults.errors[i];
        if (i < 5) { // Limit to first 5
          detailsMessage += '• ' + error.teacher + ': ' + error.error + '\n';
        }
      }
      if (lessonResults.errors.length > 5) {
        detailsMessage += '• ... and ' + (lessonResults.errors.length - 5) + ' more errors\n';
      }
      detailsMessage += '\n';
    }
    
    // Show invoice generation errors
    if (invoiceResults.errors && invoiceResults.errors.length > 0) {
      detailsMessage += '🚫 INVOICE GENERATION ERRORS (' + invoiceResults.errors.length + '):\n';
      for (var j = 0; j < invoiceResults.errors.length; j++) {
        var invoiceError = invoiceResults.errors[j];
        if (j < 5) { // Limit to first 5
          detailsMessage += '• ' + invoiceError.teacher + ': ' + invoiceError.error + '\n';
        }
      }
      if (invoiceResults.errors.length > 5) {
        detailsMessage += '• ... and ' + (invoiceResults.errors.length - 5) + ' more errors\n';
      }
      detailsMessage += '\n';
    }
    
    // Show validation issues
    if (lessonResults.validation.issues.length > 0) {
      detailsMessage += '⚠️ VALIDATION ISSUES (' + lessonResults.validation.issues.length + '):\n';
      for (var k = 0; k < lessonResults.validation.issues.length; k++) {
        var issue = lessonResults.validation.issues[k];
        if (k < 5) { // Limit to first 5
          detailsMessage += '• ' + issue + '\n';
        }
      }
      if (lessonResults.validation.issues.length > 5) {
        detailsMessage += '• ... and ' + (lessonResults.validation.issues.length - 5) + ' more issues\n';
      }
    }
    
    detailsMessage += '\nCheck Teacher_Invoice_Debug sheet for complete details.';
    
    ui.alert('Error Details', detailsMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('showCombinedErrorDetails', 'ERROR', 'Error in showCombinedErrorDetails', '', error.message);
  }
}

function showInvoiceGenerationResults(lessonResults, invoiceResults) {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var summaryMessage = '🎉 INVOICE GENERATION COMPLETE!\n\n';
    
    summaryMessage += '📋 INVOICE SHEET:\n';
    summaryMessage += '• Sheet name: ' + invoiceResults.sheetName + '\n';
    summaryMessage += '• Teachers processed: ' + invoiceResults.teacherCount + '\n';
    summaryMessage += '• Line items created: ' + invoiceResults.lineItemCount + '\n\n';
    
    summaryMessage += '📅 LESSON DATA:\n';
    summaryMessage += '• Cutoff date: ' + UtilityScriptLibrary.formatDateFlexible(lessonResults.cutoffDate, 'M/d/yy') + '\n';
    summaryMessage += '• Invoice date: ' + UtilityScriptLibrary.formatDateFlexible(lessonResults.invoiceDate, 'M/d/yy') + '\n';
    summaryMessage += '• Invoice period: ' + lessonResults.invoicePeriod + '\n';
    summaryMessage += '• Total lessons found: ' + lessonResults.summary.totalLessons + '\n\n';
    
    var totalErrors = lessonResults.errors.length + (invoiceResults.errors ? invoiceResults.errors.length : 0);
    var totalIssues = lessonResults.validation.issues.length;
    
    if (totalErrors === 0 && totalIssues === 0) {
      summaryMessage += '✅ SUCCESS: Invoice sheet generated successfully!\n';
    } else {
      summaryMessage += '⚠️ COMPLETED WITH ISSUES:\n';
      summaryMessage += '• Errors: ' + totalErrors + '\n';
      summaryMessage += '• Validation issues: ' + totalIssues + '\n';
    }
    
    summaryMessage += '\nCheck Teacher_Invoice_Debug sheet for detailed logs.';
    
    ui.alert('Invoice Generation Complete', summaryMessage, ui.ButtonSet.OK);
    
    if (totalErrors > 0 || totalIssues > 0) {
      var showDetailsResponse = ui.alert(
        'Show Details?',
        'Would you like to see details about the ' + totalErrors + ' errors and ' + totalIssues + ' validation issues?',
        ui.ButtonSet.YES_NO
      );
      
      if (showDetailsResponse === ui.Button.YES) {
        showCombinedErrorDetails(lessonResults, invoiceResults);
      }
    }
    
    UtilityScriptLibrary.debugLog('showInvoiceGenerationResults', 'INFO', 'Results summary displayed to user', '', '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('showInvoiceGenerationResults', 'ERROR', 'Error showing results summary', '', error.message);
  }
}

function showTeacherInvoiceResults(results) {
  try {
    var ui = SpreadsheetApp.getUi();
    var message = 'Teacher Invoice Generation Complete!\n\n';
    
    message += 'Successfully generated: ' + results.successful + '\n';
    message += 'Skipped (already exist): ' + results.skipped + '\n';
    message += 'Errors: ' + results.errors.length + '\n';
    
    if (results.errors.length > 0) {
      message += '\nErrors:\n';
      for (var i = 0; i < Math.min(results.errors.length, 3); i++) {
        message += '• ' + results.errors[i].teacher + ': ' + results.errors[i].error + '\n';
      }
      if (results.errors.length > 3) {
        message += '• ... and ' + (results.errors.length - 3) + ' more (check debug log)\n';
      }
    }
    
    ui.alert('Invoice Generation Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showTeacherInvoiceResults", "ERROR", "Failed to show results", "", error.message);
  }
}

function updateMetadataStatus(invoiceMonth, newStatus) {
  try {
    UtilityScriptLibrary.debugLog("updateMetadataStatus", "INFO", 
                                  "Updating status", 
                                  "Month: " + invoiceMonth + ", Status: " + newStatus, "");
    
    var teacherInvoicesSS = UtilityScriptLibrary.getWorkbook('teacherInvoices');
    var metadataSheet = teacherInvoicesSS.getSheetByName('Teacher Invoicing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Teacher Invoicing Metadata sheet not found');
    }
    
    var data = metadataSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    
    var monthCol = headerMap[UtilityScriptLibrary.normalizeHeader("Invoice Month")];
    var statusCol = headerMap[UtilityScriptLibrary.normalizeHeader("Status")];
    
    // Find the row matching the invoice month
    for (var i = 1; i < data.length; i++) {
      if (data[i][monthCol - 1] === invoiceMonth) {
        metadataSheet.getRange(i + 1, statusCol).setValue(newStatus);
        
        UtilityScriptLibrary.debugLog("updateMetadataStatus", "SUCCESS", 
                                      "Status updated", 
                                      "Month: " + invoiceMonth + ", New status: " + newStatus, "");
        return;
      }
    }
    
    throw new Error("Metadata not found for month: " + invoiceMonth);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("updateMetadataStatus", "ERROR", 
                                  "Failed to update status", 
                                  "Month: " + invoiceMonth, error.message);
    throw error;
  }
}

function updateTeacherInvoiceHistory(teacherData, invoiceResult, metadata) {
  try {
    // Only log if we have a successful invoice
    if (!invoiceResult || !invoiceResult.success) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "No successful invoice to log", 
                   "Teacher: " + teacherData.teacherName, "");
      return;
    }
    
    // Get teacher info using existing utility
    var teacherInfo = getTeacherInfoByName(teacherData.teacherName);
    if (!teacherInfo || !teacherInfo.rosterUrl) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Teacher info not found", 
                   "Teacher: " + teacherData.teacherName, "");
      return;
    }
    
    // Open teacher workbook using existing URL parsing pattern
    var fileIdMatch = teacherInfo.rosterUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!fileIdMatch) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Invalid roster URL", 
                   "Teacher: " + teacherData.teacherName + ", URL: " + teacherInfo.rosterUrl, "");
      return;
    }
    
    var teacherWorkbook = SpreadsheetApp.openById(fileIdMatch[1]);
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Opened teacher workbook", 
                 "Teacher: " + teacherData.teacherName + ", Workbook name: " + teacherWorkbook.getName(), "");
    
    // Get or create Invoice Log sheet
    var invoiceLogSheet = teacherWorkbook.getSheetByName('Invoice Log');
    if (!invoiceLogSheet) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Invoice Log sheet not found", 
                   "Teacher: " + teacherData.teacherName + 
                   ", Available sheets: " + teacherWorkbook.getSheets().map(function(s) { return s.getName(); }).join(", "), "");
      return; // Sheet should exist if created in Responses, but handle gracefully
    }
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Found Invoice Log sheet", 
                 "Teacher: " + teacherData.teacherName + ", Last row: " + invoiceLogSheet.getLastRow(), "");
    
    // Calculate total from teacher's student rows
    var totalAmount = 0;
    if (!teacherData.studentRows || teacherData.studentRows.length === 0) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "No student rows found", 
                   "Teacher: " + teacherData.teacherName, "");
    } else {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Calculating total amount", 
                   "Teacher: " + teacherData.teacherName + ", Student rows: " + teacherData.studentRows.length, "");
      
      for (var i = 0; i < teacherData.studentRows.length; i++) {
        var rowCost = teacherData.studentRows[i].cost || 0;
        totalAmount += rowCost;
        UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Adding student cost", 
                     "Row " + i + ": " + rowCost + ", Running total: " + totalAmount, "");
      }
    }
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Total calculated", 
                 "Teacher: " + teacherData.teacherName + ", Total: " + totalAmount, "");
    
    // Format dates and create invoice period string
    var formattedInvoiceDate = '';
    var invoicePeriodString = '';
    
    if (metadata && metadata.lessonsStartingDate && metadata.lessonsEndingDate) {
      // Use metadata for date range
      var startDate = new Date(metadata.lessonsStartingDate);
      var endDate = new Date(metadata.lessonsEndingDate);
      
      var formattedStart = UtilityScriptLibrary.formatDateFlexible(startDate, "MM/dd/yyyy");
      var formattedEnd = UtilityScriptLibrary.formatDateFlexible(endDate, "MM/dd/yyyy");
      invoicePeriodString = formattedStart + " - " + formattedEnd;
      
      if (metadata.invoiceDate) {
        formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(new Date(metadata.invoiceDate), "MM/dd/yyyy");
      }
      
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Using metadata for dates", 
                   "Period: " + invoicePeriodString + ", Invoice Date: " + formattedInvoiceDate, "");
    } else {
      // Fallback: use teacher data (old behavior)
      if (teacherData.invoiceDate) {
        formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(new Date(teacherData.invoiceDate), "MM/dd/yyyy");
      }
      invoicePeriodString = teacherData.invoicePeriod || '';
      
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Using fallback for dates (no metadata)", 
                   "Period: " + invoicePeriodString, "");
    }
    
    // Add new invoice record - append first to get the row number
    var newRow = [
      teacherData.invoiceNumber || '',           // Invoice Number
      formattedInvoiceDate || '',                // Invoice Date
      invoicePeriodString || '',                 // Invoice Period (now a date range)
      '',                                        // Invoice URL (will be set as formula)
      UtilityScriptLibrary.formatCurrency(totalAmount) // Total Amount
    ];
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Prepared row data", 
                 "Teacher: " + teacherData.teacherName + 
                 ", Invoice#: " + newRow[0] + 
                 ", Date: " + newRow[1] + 
                 ", Period: " + newRow[2] + 
                 ", Has URL: " + (invoiceResult.url ? "YES" : "NO") + 
                 ", Total: " + newRow[4], "");
    
    invoiceLogSheet.appendRow(newRow);
    var newRowNumber = invoiceLogSheet.getLastRow();
    
    // Now set the URL column (column 4) as a HYPERLINK formula (text = fileId, link = URL)
    if (invoiceResult.url && invoiceResult.fileId) {
      var urlFormula = '=HYPERLINK("' + invoiceResult.url + '", "' + (teacherData.invoiceNumber || invoiceResult.fileId) + '")';
      invoiceLogSheet.getRange(newRowNumber, 4).setFormula(urlFormula);
      
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Set URL as hyperlink formula", 
                   "Row: " + newRowNumber + ", FileId: " + invoiceResult.fileId, "");
    }
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "SUCCESS", "Added invoice to log", 
                 "Teacher: " + teacherData.teacherName + ", Invoice: " + teacherData.invoiceNumber + 
                 ", New last row: " + newRowNumber, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "ERROR", "Failed to update invoice history", 
                 "Teacher: " + teacherData.teacherName + ", Error: " + error.message + 
                 ", Stack: " + error.stack, "");
    // Don't throw - this is not critical to the invoice generation process
  }
}

function validateLessonData(groupedLessons) {
  var issues = [];
  
  var groupKeys = Object.keys(groupedLessons);
  for (var i = 0; i < groupKeys.length; i++) {
    var groupKey = groupKeys[i];
    var group = groupedLessons[groupKey];
    
    // Check for missing required data
    if (!group.teacherName) {
      issues.push(groupKey + ': Missing teacher name');
    }
    if (!group.studentId) {
      issues.push(groupKey + ': Missing student ID');
    }
    if (!group.studentName) {
      issues.push(groupKey + ': Missing student name');
    }
    if (!group.lessonLength || ![30, 45, 60].includes(group.lessonLength)) {
      issues.push(groupKey + ': Invalid lesson length (' + group.lessonLength + ')');
    }
    if (group.quantity <= 0) {
      issues.push(groupKey + ': Invalid quantity (' + group.quantity + ')');
    }
    
    // Validate individual lessons in group
    for (var j = 0; j < group.lessons.length; j++) {
      var lesson = group.lessons[j];
      if (!lesson.lessonDate || isNaN(lesson.lessonDate.getTime())) {
        issues.push(groupKey + ': Invalid lesson date in group');
      }
    }
  }
  
  UtilityScriptLibrary.debugLog('validateLessonData', 'INFO', 'Validation complete',
    issues.length + ' issues found', '');
  
  return {
    issues: issues,
    isValid: issues.length === 0
  };
}

function verifyAttendanceVsInvoices() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Prompt for year
    var yearResponse = ui.prompt(
      'Verify Attendance vs Invoices',
      'Enter year (e.g., 2025):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (yearResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    var year = yearResponse.getResponseText().trim();
    if (!year || isNaN(year)) {
      ui.alert('Invalid year');
      return;
    }
    
    // Get rosters folder from config
    var env = UtilityScriptLibrary.EnvironmentManager.get();
    var config = UtilityScriptLibrary.getConfig();
    var rosterFolderId = config[env].rosterFolderId;
    
    if (!rosterFolderId) {
      throw new Error('Roster folder ID not found in config');
    }
    
    var rostersFolder = DriveApp.getFolderById(rosterFolderId);
    
    // Find year subfolder
    var yearFolderName = year + ' Rosters';
    var yearFolder = null;
    var subfolders = rostersFolder.getFolders();
    
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      if (folder.getName() === yearFolderName) {
        yearFolder = folder;
        break;
      }
    }
    
    if (!yearFolder) {
      throw new Error('Year folder not found: ' + yearFolderName);
    }
    
    UtilityScriptLibrary.debugLog('verifyAttendanceVsInvoices', 'INFO', 
                                  'Found year folder', yearFolderName, '');
    
    // Get Teacher Invoices workbook (current)
    var teacherInvoicesWB = SpreadsheetApp.getActiveSpreadsheet();
    
    // Data structures
    var unpaidLessons = [];
    var attendanceData = {};
    var invoiceData = {};
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    // Process each teacher roster file in the year folder
    var files = yearFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var fileCount = 0;
    var processedCount = 0;
    
    while (files.hasNext()) {
      var file = files.next();
      var fileName = file.getName();
      fileCount++;
      
      // Check if filename matches pattern: "QAMP yyyy LastName"
      var pattern = new RegExp('^QAMP ' + year + ' (.+)$');
      var match = fileName.match(pattern);
      
      if (!match) {
        UtilityScriptLibrary.debugLog('verifyAttendanceVsInvoices', 'INFO', 
                                      'Skipping file — does not match QAMP pattern', fileName, '');
        continue;
      }
      
      var lastName = match[1].trim();
      processedCount++;
      
      if (processedCount % 5 === 0) {
        UtilityScriptLibrary.debugLog('verifyAttendanceVsInvoices', 'INFO', 
                                      'Processing files', 
                                      processedCount + ' of ' + fileCount, '');
      }
      
      try {
        var teacherWB = SpreadsheetApp.openById(file.getId());
        var sheets = teacherWB.getSheets();
        
        for (var i = 0; i < sheets.length; i++) {
          var sheet = sheets[i];
          var sheetName = sheet.getName();
          
          if (monthNames.indexOf(sheetName) === -1) {
            continue;
          }
          
          processAttendanceSheet(sheet, lastName, sheetName, unpaidLessons, attendanceData);
        }
        
      } catch (error) {
        UtilityScriptLibrary.debugLog('verifyAttendanceVsInvoices', 'ERROR', 
                                      'Failed to process file', 
                                      fileName, error.message);
      }
    }
    
    UtilityScriptLibrary.debugLog('verifyAttendanceVsInvoices', 'INFO', 
                                  'Processed roster files', 
                                  processedCount + ' files processed', '');
    
    // Process invoice sheets
    var allSheets = teacherInvoicesWB.getSheets();
    
    for (var i = 0; i < allSheets.length; i++) {
      var sheet = allSheets[i];
      var name = sheet.getName();
      
      if (name === 'Monthly Template' || 
          name === 'Teacher Invoicing Metadata' ||
          name === 'Instructions' ||
          name.indexOf('Verification') !== -1 ||
          name.indexOf('Totals') !== -1 ||
          name.indexOf('Report') !== -1) {
        continue;
      }
      
      processInvoiceSheetForVerification(sheet, invoiceData);
    }
    
    // Create verification reports
    createUnpaidLessonsReport(teacherInvoicesWB, unpaidLessons);
    createStudentReconciliationReport(teacherInvoicesWB, attendanceData, invoiceData);
    
    ui.alert('✅ Verification complete!\n\n' +
             'Files processed: ' + processedCount + '\n' +
             'Unpaid lessons: ' + unpaidLessons.length + '\n\n' +
             'Check:\n- Unpaid Lessons Report\n- Student Reconciliation Report');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('verifyAttendanceVsInvoices', 'ERROR', 
                                  'Verification failed', '', error.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod) {
  try {
    UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "INFO", 
                                  "Writing metadata", 
                                  "Month: " + month, "");
    
    var teacherInvoicesSS = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = teacherInvoicesSS.getSheetByName('Teacher Invoicing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Teacher Invoicing Metadata sheet not found');
    }
    
    var lessonsStartingDate = calculateLessonsStartingDate(metadataSheet);
    
    if (!lessonsStartingDate) {
      lessonsStartingDate = new Date(cutoffDate.getFullYear(), cutoffDate.getMonth(), 1);
      UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "INFO", 
                                    "Using first day of month for starting date", 
                                    "Date: " + UtilityScriptLibrary.formatDateFlexible(lessonsStartingDate, 'M/d/yy'), "");
    }
    
    var semesterName = UtilityScriptLibrary.getCurrentSemesterName();
    var ratesLookup = getRatesLookupForSemester(semesterName);
    
    var newRow = [
      month,
      lessonsStartingDate,
      cutoffDate,
      invoiceDate,
      ratesLookup,
      semesterName,
      'Collected'
    ];
    
    metadataSheet.appendRow(newRow);
    
    UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "SUCCESS", 
                                  "Metadata written", 
                                  "Month: " + month + ", Period: " + UtilityScriptLibrary.formatDateFlexible(lessonsStartingDate, 'M/d/yy') + " to " + UtilityScriptLibrary.formatDateFlexible(cutoffDate, 'M/d/yy'), "");
    
    return {
      success: true,
      lessonsStartingDate: lessonsStartingDate,
      lessonsEndingDate: cutoffDate
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "ERROR", 
                                  "Failed to write metadata", 
                                  "Month: " + month, error.message);
    throw error;
  }
}
