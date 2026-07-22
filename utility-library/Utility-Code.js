/*
================================================================================
UTILITY LIBRARY CODE
================================================================================
Version: 91
Total Functions: 137 (135 standard + 2 EnvironmentManager methods)
Documentation: See Utility-Functions.md
================================================================================
*/
// #region
var _executionCache = {};

var EnvironmentManager = (function () {
  var currentEnv = 'test'; // set to 'test' as default

  return {
    set: function (env) {
      if (env !== 'test' && env !== 'prod') {
        throw new Error('Invalid environment: "' + env + '"');
      }
      currentEnv = env;
    },
    get: function () {
      return currentEnv;
    }
  };
})();

// Set to true for historical data import
// Set to false for normal system generation
var HISTORICAL_DATA_MODE = false;  // 👈 Just change this one line!

var SHEET_MAP = {
  // Teacher Interest Workbook
  teacherResponses: {
    name: 'New Teacher Responses'
  },
  teacherFieldMap: {
    name: 'TeacherFieldMap'
  },
  teacherReturningResponses: {
    name: 'Returning Teacher Responses'
  },
  teacherReturningFieldMap: {
    name: 'Returning Teacher FieldMap'
  },
  // Contacts Workbook
  teachersAndAdmin: {
    name: 'Teachers and Admin'
  },
  students: {
    name: 'Students'
  },
  parents: {
    name: 'Parents'
  },
  instrumentList: {
    name: 'Instrument List'
  },
  futureTeachers: {
    name: 'Future Teacher Contacts'
  },
  // Billing Workbook
  billingMetadata: {
    name: 'Billing Metadata'
  },
  semesterMetadata: {
    name: 'Semester Metadata'
  },
  yearMetadata: {
    name: 'Year Metadata'
  },
  rates: {
    name: 'Rates'
  },
  programList: {
    name: 'Programs List'
  },
  packages: {
    name: 'Packages'
  },
  billingTemplate: {
    name: 'Billing Template'
  },
  reregistration: {
    name: 'Reregistration'
  },
  reregistrationQueue: {
    name: 'Reregistration Queue'
  },
  // Registration Workbook
  calendar: {
    name: 'Calendar'
  },
  fieldMap: {
    name: 'FieldMap'
  },
  teacherRosterLookup: {
    name: 'Teacher Roster Lookup'
  },
  // Payments Workbook
  ledgerTemplate: {
    name: 'Ledger Template'
  },
  // Scheduling Workbook
  teacherPreferences: {
    name: 'Teacher Preferences'
  }
};

var TEMPLATE_MAP = {
  // Welcome Letters for Print
  newFamilyPrint: 'NewFamilyLetter-print',
  newAdultPrint: 'NewAdultLetter-print', 
  returningFamilyPrint: 'ReturningFamilyLetter-print',
  returningAdultPrint: 'ReturningAdultLetter-print',
  secondInvoicePrint: 'SecondInvoiceLetter-print',
  revisedInvoicePrint: 'RevisedInvoiceLetter-print',
  missingDocumentPrint: 'MissingDocumentLetter-print',

  // Welcome Letters for Email
  newFamilyEmail: 'NewFamilyLetter-email',
  newAdultEmail: 'NewAdultLetter-email', 
  returningFamilyEmail: 'ReturningFamilyLetter-email',
  returningAdultEmail: 'ReturningAdultLetter-email',
  secondInvoiceEmail: 'SecondInvoiceLetter-email',
  revisedInvoiceEmail: 'RevisedInvoiceLetter-email',
  missingDocumentEmail: 'MissingDocumentLetter-email',
  
  // Legal & Business
  invoice: 'Invoice',
  agreement: 'Agreement',
  mediaReleaseChild: 'MediaConsentChild',
  mediaReleaseAdult: 'MediaConsentAdult',
  teacherInvoice: 'TeacherInvoiceTemplate', 
  
  // Administrative
  paymentReminder: 'Payment Reminder Letter',
  revisedInvoice: 'Revised Invoice Letter',
  paymentReceipt: 'Receipt',
};

/**
 * Standard styles for consistent formatting across all sheets
 * Access via UtilityScriptLibrary.STYLES.HEADER.background, etc.
 * For borders, use STYLES.HEADER.background (same green color)
 */
var STYLES = {
  HEADER: {
    background: '#37a247',
    text: '#ffffff',
    bold: true
  },
  SUBHEADER: {
    background: '#e0f0fa',
    text: '#2a4d7c',
    bold: true
  },
  WARNING: {
    background: '#f4cccc',
    text: '#7a1f1f'
  },
  ALTERNATING_DARK: {
    background: '#f3f3f3'
  },
  ALTERNATING_LIGHT: {
    background: '#ffffff'
  }
};

var INSTRUMENT_FAMILIES = {
  Strings: ['Violin', 'Viola', 'Cello', 'Bass'],
  Winds: ['Flute', 'Piccolo', 'Oboe', 'Bassoon', 'Clarinet', 'Saxophone'],
  Brass: ['French Horn', 'Trumpet', 'Trombone', 'Baritone', 'Euphonium', 'Tuba'],
  Auxiliary: ['Harp', 'Piano', 'Percussion', 'Voice', 'Guitar']
};

var MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];
// #endregion
function addToCurrencyCols(currencyCols, columnNumber, headerName) {
  // CURRENCY VALIDATION: Never add hours columns to currencyCols
  if (typeof columnNumber !== 'number' || columnNumber <= 0) {
    return; // Invalid column number
  }
  
  // Only add if it should be currency formatted
  if (shouldBeCurrency(headerName)) {
    currencyCols.push(columnNumber);
  } else {
    debugLog('addToCurrencyCols', 'INFO', 'Skipped non-currency column', 
             'Column: ' + columnNumber + ', Header: ' + headerName, '');
  }
}

function appendToMetadataWithVerification(metadataSheet, rowData, verificationSteps, options) {
  options = options || {};
  var requireAllVerifications = options.requireAllVerifications !== false; // Default true
  var showFinalConfirmation = options.showFinalConfirmation !== false; // Default true
  
  try {
    var ui = SpreadsheetApp.getUi();
    var verificationResults = [];
    
    // Execute each verification step
    for (var i = 0; i < verificationSteps.length; i++) {
      var step = verificationSteps[i];
      var stepResult = {
        name: step.name || 'Step ' + (i + 1),
        success: false,
        data: null,
        userConfirmed: false
      };
      
      try {
        // Execute verification function if provided
        if (step.verifyFunction && typeof step.verifyFunction === 'function') {
          stepResult.data = step.verifyFunction();
          stepResult.success = true;
        }
        
        // Show user confirmation if required
        if (step.confirmationRequired !== false) { // Default true
          var confirmationMessage = step.message || 'Please confirm this step';
          var details = step.formatData && stepResult.data ? 
            step.formatData(stepResult.data) : 
            (stepResult.data || '');
          
          var confirmed = showConfirmationDialog(
            step.title || stepResult.name,
            confirmationMessage,
            details,
            { buttonSet: ui.ButtonSet.YES_NO, confirmButton: ui.Button.YES }
          );
          
          stepResult.userConfirmed = confirmed;
          
          if (!confirmed && requireAllVerifications) {
            throw new Error('User cancelled verification at step: ' + stepResult.name);
          }
        } else {
          stepResult.userConfirmed = true;
        }
        
      } catch (stepError) {
        stepResult.error = stepError.message;
        if (requireAllVerifications) {
          throw new Error('Verification failed at step "' + stepResult.name + '": ' + stepError.message);
        }
      }
      
      verificationResults.push(stepResult);
    }
    
    // Final confirmation if requested
    if (showFinalConfirmation) {
      var summaryLines = ['Verification Summary:'];
      for (var j = 0; j < verificationResults.length; j++) {
        var result = verificationResults[j];
        var status = result.userConfirmed ? 'Confirmed' : 'Skipped';
        if (result.error) status = 'Failed';
        summaryLines.push('• ' + result.name + ': ' + status);
      }
      
      var finalConfirmed = showConfirmationDialog(
        'Final Confirmation',
        'Ready to add metadata entry?',
        summaryLines,
        { buttonSet: ui.ButtonSet.YES_NO, confirmButton: ui.Button.YES }
      );
      
      if (!finalConfirmed) {
        throw new Error('User cancelled final confirmation');
      }
    }
    
    // Append the data
    metadataSheet.appendRow(rowData);
    
    debugLog('appendToMetadataWithVerification', 'SUCCESS', 'Metadata appended', 'Steps: ' + verificationSteps.length, '');
    
    return {
      success: true,
      message: 'Metadata appended successfully',
      verificationResults: verificationResults,
      rowData: rowData
    };
    
  } catch (error) {
    debugLog('appendToMetadataWithVerification', 'ERROR', 'Failed to append metadata', '', error.message);
    return {
      success: false,
      error: error.message,
      verificationResults: verificationResults || []
    };
  }
}

function buildAcademicYearVariables() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var semesterSheet = ss.getSheetByName('Semester Metadata');
    
    if (!semesterSheet) {
      throw new Error('Semester Metadata sheet not found');
    }
    
    var data = semesterSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('No semester data found');
    }
    
    // Get the most recent semester (last row)
    var latestSemester = data[data.length - 1];
    var semesterName = latestSemester[0];
    
    // Extract year from semester name (e.g., "Fall 2024" -> "2024")
    var yearMatch = semesterName.match(/\d{4}/);
    if (!yearMatch) {
      throw new Error('Could not extract year from semester: ' + semesterName);
    }
    
    var currentYear = parseInt(yearMatch[0], 10);
    var nextYear = currentYear + 1;
    var academicYear = currentYear + '-' + nextYear;
    
    return {
      'CurrentAcademicYear': academicYear,
      'AcademicYearStart': currentYear.toString(),
      'AcademicYearEnd': nextYear.toString()
    };
    
  } catch (error) {
    return {
      'CurrentAcademicYear': '2024-2025',
      'AcademicYearStart': '2024',
      'AcademicYearEnd': '2025'
    };
  }
}

function buildRateMapFromSheet(sheet, headers, yearColIndex) {
  var data = sheet.getDataRange().getValues();
  var rateMap = {};
  for (var i = 1; i < data.length; i++) {
    var title = data[i][0];
    var value = data[i][yearColIndex];
    if (title && typeof value !== 'undefined') {
      rateMap[title] = value;
    }
  }
  return rateMap;
}

function buildRateVariables() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rateSheet = ss.getSheetByName('Rates');
    
    if (!rateSheet) {
      throw new Error('Rates sheet not found');
    }
    
    var rateData = rateSheet.getDataRange().getValues();
    var rateHeaders = rateData[0];
    
    // FIXED: Always use column B (index 1) for rates
    var bestColIndex = 1;
    
    var rateMap = buildRateMapFromSheet(rateSheet, rateHeaders, bestColIndex);
    
    var hourlyRate = parseFloat(rateMap['Lesson']);
    var lateFee = parseFloat(rateMap['Late Fee']);
    
    if (isNaN(hourlyRate)) {
      throw new Error('Invalid Lesson rate in rates sheet');
    }
    
    if (isNaN(lateFee)) {
      throw new Error('Invalid Late Fee rate in rates sheet');
    }
    
    var halfHourRate = hourlyRate / 2;
    
    return {
      'HourlyRate': formatCurrency(hourlyRate),
      'HalfHourRate': formatCurrency(halfHourRate),
      'LateFee': formatCurrency(lateFee),
      'LateFeeGracePeriod': '10'
    };
    
  } catch (error) {
    return {
      'HourlyRate': '$0.00',
      'HalfHourRate': '$0.00',
      'LateFee': '$0.00',
      'LateFeeGracePeriod': '10'
    };
  }
}

function bulkUpdateStudentStatus(studentsSheet, statusColumn, newValue, options) {
  options = options || {};
  var condition = options.condition; // Function to test each row
  var whereColumn = options.whereColumn; // Column to check
  var whereValue = options.whereValue; // Value to match
  var skipEmpty = options.skipEmpty !== false; // Default true
  
  try {
    var data = studentsSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find target column
    var targetCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (normalizeHeader(headers[i]) === normalizeHeader(statusColumn)) {
        targetCol = i;
        break;
      }
    }
    
    if (targetCol === -1) {
      throw new Error('Column "' + statusColumn + '" not found in students sheet');
    }
    
    // Find condition column if specified
    var conditionCol = -1;
    if (whereColumn) {
      for (var j = 0; j < headers.length; j++) {
        if (normalizeHeader(headers[j]) === normalizeHeader(whereColumn)) {
          conditionCol = j;
          break;
        }
      }
      if (conditionCol === -1) {
        throw new Error('Condition column "' + whereColumn + '" not found');
      }
    }
    
    var updatedCount = 0;
    var changedRows = [];
    
    // Process each row
    for (var k = 1; k < data.length; k++) {
      var row = data[k];
      var shouldUpdate = false;
      
      // Skip empty rows if requested
      if (skipEmpty) {
        var hasData = false;
        for (var col = 0; col < row.length; col++) {
          if (row[col] && row[col].toString().trim() !== '') {
            hasData = true;
            break;
          }
        }
        if (!hasData) continue;
      }
      
      // Apply condition logic
      if (condition && typeof condition === 'function') {
        try {
          shouldUpdate = condition(row, k + 1);
        } catch (conditionError) {
          debugLog('bulkUpdateStudentStatus', 'WARNING', 'Condition function error', 'Row: ' + (k + 1), conditionError.message);
          continue;
        }
      } else if (whereColumn && whereValue !== undefined) {
        shouldUpdate = row[conditionCol] === whereValue;
      } else {
        // Update all non-empty rows if no condition specified
        shouldUpdate = true;
      }
      
      // Update the row if condition met
      if (shouldUpdate) {
        var oldValue = row[targetCol];
        data[k][targetCol] = newValue;
        updatedCount++;
        
        changedRows.push({
          rowNumber: k + 1,
          oldValue: oldValue,
          newValue: newValue
        });
      }
    }
    
    // Write back the updated data
    if (updatedCount > 0) {
      studentsSheet.getDataRange().setValues(data);
    }
    
    debugLog('bulkUpdateStudentStatus', 'SUCCESS', 'Bulk update completed', 'Rows updated: ' + updatedCount, '');
    
    return {
      success: true,
      updatedCount: updatedCount,
      changedRows: changedRows,
      message: 'Updated ' + updatedCount + ' student records'
    };
    
  } catch (error) {
    debugLog('bulkUpdateStudentStatus', 'ERROR', 'Bulk update failed', '', error.message);
    return {
      success: false,
      error: error.message,
      updatedCount: 0,
      changedRows: []
    };
  }
}

function calculateGraduationYear(grade) {
  try {
    debugLog('calculateGraduationYear', 'INFO', 'Starting', 'Grade: ' + JSON.stringify(grade), '');

    if (!grade) {
      return '';
    }

    var currentYear = new Date().getFullYear();
    var gradeStr = String(grade).toLowerCase().trim();

    if (gradeStr.indexOf('adult') !== -1 || gradeStr.indexOf('college') !== -1) {
      return 'Adult';
    }

    if (gradeStr.indexOf('k') !== -1 || gradeStr === 'kindergarten') {
      return currentYear + 13;
    }

    var gradeNum = parseInt(gradeStr.replace(/[^\d]/g, ''));

    if (isNaN(gradeNum) || gradeNum < 1 || gradeNum > 12) {
      debugLog('calculateGraduationYear', 'WARNING', 'Invalid grade', 'Grade: ' + grade, '');
      return '';
    }

    var result = currentYear + (13 - gradeNum);
    debugLog('calculateGraduationYear', 'SUCCESS', 'Calculated', 'Year: ' + result, '');
    return result;

  } catch (error) {
    debugLog('calculateGraduationYear', 'ERROR', 'Failed', '', error.message);
    return '';
  }
}

function cascadeFormerStatus(teacherId) {
  try {
    debugLog('cascadeFormerStatus', 'INFO', 'Starting former status cascade',
             'Teacher ID: ' + teacherId, '');

    if (!teacherId || String(teacherId).trim() === '') {
      throw new Error('Teacher ID is required');
    }

    var instrumentSheet = getSheet('instrumentList');
    var headerMap = getHeaderMap(instrumentSheet);

    var teacherIdCol = headerMap[normalizeHeader('Teacher ID')];
    var statusCol = headerMap[normalizeHeader('Status')];

    if (!teacherIdCol) {
      throw new Error('Teacher ID column not found in Instrument List');
    }
    if (!statusCol) {
      throw new Error('Status column not found in Instrument List');
    }

    var data = instrumentSheet.getDataRange().getValues();
    var updatedCount = 0;

    for (var i = 1; i < data.length; i++) {
      var rowTeacherId = String(data[i][teacherIdCol - 1] || '').trim();
      if (rowTeacherId === String(teacherId).trim()) {
        instrumentSheet.getRange(i + 1, statusCol).setValue('former');
        updatedCount++;
      }
    }

    debugLog('cascadeFormerStatus', 'SUCCESS', 'Former cascade complete',
             'Teacher ID: ' + teacherId + ', Instrument rows updated: ' + updatedCount, '');

    return updatedCount;

  } catch (error) {
    debugLog('cascadeFormerStatus', 'ERROR', 'Former cascade failed',
             'Teacher ID: ' + teacherId, error.message);
    throw error;
  }
}

function cleanName(name) {
  return name ? name.toString().trim() : '';
}

function clearCache() {
  _executionCache = {};
}

function clearEmptyRows(sheet) {
  var lastRow = sheet.getLastRow();
  var maxRows = sheet.getMaxRows();
  if (maxRows > lastRow) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

function clearOldDebugEntries(debugSheet) {
  try {
    var lastRow = debugSheet.getLastRow();
    if (lastRow > 500) {
      // Keep header + last 100 entries, delete the rest
      var rowsToDelete = lastRow - 101;
      if (rowsToDelete > 0) {
        debugSheet.deleteRows(2, rowsToDelete);
        Logger.log("🧹 Cleared " + rowsToDelete + " old debug entries");
      }
    }
  } catch (error) {
    Logger.log("❌ Error clearing old debug entries: " + error.message);
  }
}

function clearUtilityDebugLog() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName("Debug");
    
    if (debugSheet) {
      var lastRow = debugSheet.getLastRow();
      if (lastRow > 1) {
        debugSheet.getRange(2, 1, lastRow - 1, 6).clear();
        debugLog("clearUtilityDebugLog", "INFO", "Utility debug log cleared", "", "");
      }
    }
  } catch (error) {
    Logger.log("❌ Error clearing utility debug log: " + error.message);
  }
}

function columnToLetter(column) {
  var temp = "", letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function combineDocumentsIntoPDF(documents, fileName, destinationFolder) {
  try {
    if (!documents || documents.length === 0) {
      throw new Error('No documents provided to combine');
    }
    
    // Create individual PDFs only
    var individualPdfFiles = [];
    for (var i = 0; i < documents.length; i++) {
      var doc = documents[i];
      
      try {
        var docFile = DriveApp.getFileById(doc.fileId);
        var pdfBlob = docFile.getAs('application/pdf');
        var individualPdfName = doc.name + ' - ' + fileName.replace('Registration Packet - ', '');
        var individualPdf = destinationFolder.createFile(pdfBlob.setName(individualPdfName + '.pdf'));
        individualPdfFiles.push(individualPdf);
        
      } catch (docError) {
        debugLog('combineDocumentsIntoPDF', 'ERROR', 'Failed to convert document to PDF', doc.name, docError.message);
      }
    }
    
    if (individualPdfFiles.length === 0) {
      throw new Error('No documents could be converted to PDF');
    }
    
    // Return success with individual PDFs only (no combined PDF)
    return {
      success: true,
      fileId: null, // No combined PDF
      url: null,    // No combined PDF
      fileName: fileName + ' (Individual Documents Only)',
      documentsIncluded: individualPdfFiles.length,
      individualPdfs: individualPdfFiles.map(function(file) {
        return {
          fileId: file.getId(),
          url: file.getUrl(),
          name: file.getName()
        };
      }),
      message: 'Individual PDFs created successfully. No combined packet generated.'
    };
    
  } catch (error) {
    debugLog('combineDocumentsIntoPDF', 'ERROR', 'Failed to create individual PDFs', '', error.message);
    return {
      success: false,
      error: error.message,
      fileId: null,
      url: null
    };
  }
}

function copyPreviousColumnToNew(newRow, prevRow, currMap, prevMap, mapping) {
  var newCol  = currMap[normalizeHeader(mapping.newCol)];
  var prevCol = prevMap[normalizeHeader(mapping.prevCol)];

  if (newCol && prevCol && prevRow.length >= prevCol) {
    newRow[newCol - 1] = prevRow[prevCol - 1];
    return true;
  }
  return false;
}

function copyTextAttributes(sourcePara, targetPara) {
  try {
    var text = sourcePara.getText();
    for (var i = 0; i < text.length; i++) {
      var charAttrs = sourcePara.getAttributes(i);
      if (charAttrs && Object.keys(charAttrs).length > 0) {
        targetPara.setAttributes(i, i + 1, charAttrs);
      }
    }
  } catch (e) {
    // Skip if attribute copying fails
  }
}

function convertYesNoToBoolean(value) {
  if (typeof value === 'boolean') {
    return value;
  }
  if (typeof value === 'number') {
    return value === 1 || value > 0;
  }
  if (typeof value === 'string') {
    var str = value.toLowerCase().trim();
    return str === 'yes' || str === 'true' || str === 'y';
  }
  return false;
}

function copySheetWithProtections(sourceSheet, targetWorkbook, newName, options) {
  options = options || {};
  var preserveProtections = options.preserveProtections !== false; // Default true
  var preserveFormatting = options.preserveFormatting !== false; // Default true
  var clearContent = options.clearContent || false;
  
  try {
    // Check if sheet with same name already exists
    var existingSheet = targetWorkbook.getSheetByName(newName);
    if (existingSheet) {
      throw new Error('A sheet named "' + newName + '" already exists in the target workbook');
    }
    
    // Copy the sheet
    var newSheet = sourceSheet.copyTo(targetWorkbook);
    newSheet.setName(newName);
    
    // Clear content if requested but preserve structure
    if (clearContent) {
      var lastRow = newSheet.getLastRow();
      var lastCol = newSheet.getLastColumn();
      
      if (lastRow > 1 && lastCol > 0) {
        // Clear content starting from row 2 (preserve headers)
        newSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
    
    // Copy protections if requested
    if (preserveProtections) {
      var sourceProtections = sourceSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      
      for (var i = 0; i < sourceProtections.length; i++) {
        var sourceProtection = sourceProtections[i];
        var range = sourceProtection.getRange();
        
        // Create corresponding range in new sheet
        var newRange = newSheet.getRange(
          range.getRow(),
          range.getColumn(),
          range.getNumRows(),
          range.getNumColumns()
        );
        
        // Apply protection
        var newProtection = newRange.protect();
        newProtection.setDescription(sourceProtection.getDescription() || 'Copied Protection');
        
        // Copy warning-only setting if applicable
        if (sourceProtection.isWarningOnly()) {
          newProtection.setWarningOnly(true);
        }
      }
    }
    
    debugLog('copySheetWithProtections', 'SUCCESS', 'Sheet copied', sourceSheet.getName() + ' → ' + newName, '');
    
    return {
      success: true,
      sheet: newSheet,
      message: 'Sheet copied successfully'
    };
    
  } catch (error) {
    debugLog('copySheetWithProtections', 'ERROR', 'Failed to copy sheet', '', error.message);
    return {
      success: false,
      sheet: null,
      error: error.message
    };
  }
}

function createColumnFinder(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return function(name) {
    for (var i = 0; i < headers.length; i++) {
      if (normalizeHeader(String(headers[i])) === normalizeHeader(name)) {
        return i + 1;  // Return 1-based column index
      }
    }
    return 0;  // Return 0 if not found
  };
}

function createGroupSections(sheet, groupEntries) {
  try {
    if (!groupEntries || groupEntries.length === 0) {
      debugLog('createGroupSections', 'DEBUG', 'No group entries to create', '', '');
      return;
    }
    
    var lastRow = sheet.getLastRow();
    var monthName = sheet.getName(); // Get month name from sheet name
    
    for (var i = 0; i < groupEntries.length; i++) {
      var groupEntry = groupEntries[i];
      
      // Add spacing before group (like between students)
      if (lastRow > 1) {
        lastRow += 1; // One empty row before this group
      }
      
      var startRow = lastRow + 1;
      
      // Create group header row (same format as student header)
      var headerData = [
        groupEntry.groupId,                           // A - Group ID
        groupEntry.groupName,                         // B - Group Name
        monthName,                                    // C - Month name (like students)
        '',                                           // D - Length (leave empty for groups)
        '',                                           // E - Status (empty)
        '',                                           // F - Comments (empty)
        '',                                           // G - Admin Review Date (empty)
        '',                                           // H - Invoice Date (empty)
        '',                                           // I - Payment Date (empty)
        '',                                           // J - Invoice Number (empty)
        '',                                           // K - Admin Comments (empty)
        ''                                            // L - Calendar Event ID (empty)
      ];
      
      sheet.getRange(startRow, 1, 1, headerData.length).setValues([headerData]);
      
      // Style group header (same as student header - light blue)
      var headerRange = sheet.getRange(startRow, 1, 1, headerData.length);
      headerRange.setBackground(STYLES.SUBHEADER.background)
                 .setFontColor(STYLES.SUBHEADER.text)
                 .setFontWeight('bold');
      
      // Add green border between Comments (F) and Admin Review Date (G)
      headerRange.getCell(1, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
      
      protectStudentHeaderRow(sheet, startRow);
      
      // Create 5 lesson rows for this group (same as students)
      for (var j = 1; j <= 5; j++) {
        var rowNum = startRow + j;
        
        var lessonData = [
          groupEntry.groupId,                         // A - Group ID
          groupEntry.groupName,                       // B - Group Name
          '',                                         // C - Date (teacher fills)
          '',                                         // D - Length (teacher fills)
          '',                                         // E - Status (teacher fills)
          '',                                         // F - Comments (teacher fills)
          '',                                         // G - Admin Review Date
          '',                                         // H - Invoice Date
          '',                                         // I - Payment Date
          '',                                         // J - Invoice Number
          '',                                         // K - Admin Comments
          ''                                          // L - Calendar Event ID
        ];
        
        sheet.getRange(rowNum, 1, 1, lessonData.length).setValues([lessonData]);
        
        // Apply alternating row colors (same as students - light gray for even rows)
        if (j % 2 === 0) {
          sheet.getRange(rowNum, 1, 1, lessonData.length).setBackground(STYLES.ALTERNATING_DARK.background);
        }
        
        // Add green border between Comments (F) and Admin Review Date (G)
        sheet.getRange(rowNum, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
      }
      
      lastRow = startRow + 5; // Header + 5 lesson rows
      
      debugLog('createGroupSections', 'DEBUG', 'Created group section', 
                                   'Group: ' + groupEntry.groupName + ', Rows: ' + startRow + '-' + lastRow, '');
    }
    
    debugLog('createGroupSections', 'INFO', 'Created all group sections', 
                                 'Count: ' + groupEntries.length, '');
    
  } catch (error) {
    debugLog('createGroupSections', 'ERROR', 'Error creating group sections', '', error.message);
  }
}

function createLessonRows(sheet, student, startRow, numRows) {
  for (var i = 0; i < numRows; i++) {
    var rowNum = startRow + i;
    
    // Extract numeric lesson length
    var numericLessonLength = extractNumericLessonLength(student.lessonLength);
    
    // Pre-fill lesson row data (12 columns now - includes Calendar Event ID)
    var lessonData = [
      student.id || '',                    // A - Student ID
      (student.lastName || '') + ', ' + (student.firstName || ''), // B - Student Name
      '',                                  // C - Date (teacher fills)
      numericLessonLength,                 // D - Length (NUMERIC VALUE ONLY - no " minutes")
      '',                                  // E - Status (teacher fills via dropdown)
      '',                                  // F - Comments (teacher fills)
      '',                                  // G - Admin Review Date (admin fills)
      '',                                  // H - Invoice Date (admin fills)
      '',                                  // I - Payment Date (admin fills)
      '',                                  // J - Invoice Number (admin fills)
      '',                                  // K - Admin Comments (admin fills)
      ''                                   // L - Calendar Event ID (scheduling script fills)
    ];
    
    sheet.getRange(rowNum, 1, 1, lessonData.length).setValues([lessonData]);
    
    // Apply alternating row colors (light gray for even rows within this student's section)
    if (i % 2 === 1) {
      sheet.getRange(rowNum, 1, 1, lessonData.length).setBackground(STYLES.ALTERNATING_DARK.background);
    }
    
    // Add green border between Comments (F) and Admin Review Date (G)
    sheet.getRange(rowNum, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
  
  return startRow + numRows;
}

function createMonthlyAttendanceSheet(workbook, monthName, rosterData) {
  try {
    debugLog('createMonthlyAttendanceSheet', 'INFO', 'Creating attendance sheet', monthName, '');
    
    var sheet = workbook.insertSheet(monthName);
    
    setupAttendanceHeaders(sheet);
    
    var teacherId = extractTeacherIdFromWorkbook(workbook);
    if (teacherId) {
      var groupEntries = getTeacherGroupAssignments(teacherId);
      if (groupEntries.length > 0) {
        createGroupSections(sheet, groupEntries);
        debugLog('createMonthlyAttendanceSheet', 'INFO', 'Created group sections', groupEntries.length + ' groups', '');
      }
    }
    
    if (rosterData && rosterData.length > 0) {
      createStudentSections(sheet, rosterData);
      debugLog('createMonthlyAttendanceSheet', 'INFO', 'Created student sections', rosterData.length + ' students', '');
    } else {
      debugLog('createMonthlyAttendanceSheet', 'INFO', 'Created empty attendance sheet', monthName, '');
    }
    
    formatAttendanceSheet(sheet);
    protectAttendanceSheet(sheet);
    
    debugLog('createMonthlyAttendanceSheet', 'SUCCESS', 'Attendance sheet created', monthName, '');
    return sheet;
    
  } catch (error) {
    debugLog('createMonthlyAttendanceSheet', 'ERROR', 'Failed to create attendance sheet', monthName, error.message);
    throw error;
  }
}

function createSignOffRow(sheet) {
  try {
    var numCols = 12;

    // Set entire row to white base
    sheet.getRange(2, 1, 1, numCols).setBackground('#ffffff')
                                     .setFontColor('#333333')
                                     .setFontWeight('normal');

    // B2: label
    sheet.getRange(2, 2)
         .setValue('Month Complete:')
         .setFontWeight('bold')
         .setHorizontalAlignment('right');

    // C2: initials entry cell — plain text, centered, white
    sheet.getRange(2, 3)
         .setNumberFormat('@')
         .setHorizontalAlignment('center')
         .clearDataValidations();

    // D2: hint text
    sheet.getRange(2, 4)
         .setValue('Initial here when done')
         .setFontColor('#999999')
         .setFontStyle('italic')
         .setFontWeight('normal');

    // Conditional formatting: yellow when past day 21 and C2 is blank
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(DAY(TODAY())>21,ISBLANK(C2))')
      .setBackground('#FFD966')
      .setFontColor('#333333')
      .setRanges([sheet.getRange(2, 1, 1, numCols)])
      .build();

    var rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);

    debugLog('createSignOffRow', 'INFO', 'Sign-off row created at row 2', '', '');

  } catch (error) {
    debugLog('createSignOffRow', 'ERROR', 'Failed to create sign-off row', '', error.message);
    throw error;
  }
}

function createStudentHeader(sheet, student, row) {
  var monthName = sheet.getName();
  
  var lessonLengthForHeader = formatLessonLengthWithMinutes(student.lessonLength);
  
  var headerData = [
    student.id || '',
    (student.lastName || '') + ', ' + (student.firstName || '') + ' - ' + (student.instrument || ''),
    monthName,
    lessonLengthForHeader,
    '', '', '', '', '', '', '', ''  // E-L
  ];
  
  sheet.getRange(row, 1, 1, headerData.length).setValues([headerData]);
  
  var lessonsRemaining = student.lessonsRemaining || 0;
  var isWarning = lessonsRemaining < 2 && lessonsRemaining > 0;
  
  var headerRange = sheet.getRange(row, 1, 1, headerData.length);
  if (isWarning) {
    headerRange.setBackground(STYLES.WARNING.background)
               .setFontColor(STYLES.WARNING.text)
               .setFontWeight('bold');
    
    var numericLength = extractNumericLessonLength(student.lessonLength);
    sheet.getRange(row, 4).setValue(numericLength + ' minutes - WARNING: Only ' + lessonsRemaining + ' lessons left!')
                          .setWrap(false);
  } else {
    headerRange.setBackground(STYLES.SUBHEADER.background)
               .setFontColor(STYLES.SUBHEADER.text)
               .setFontWeight('bold');
  }
  
  headerRange.getCell(1, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  protectStudentHeaderRow(sheet, row);
  
  return row + 1;
}

function createStudentSections(sheet, rosterData) {
  var lastRow = sheet.getLastRow();
  var currentRow;

  // Rows 1 (header) and 2 (sign-off) are reserved — start at row 3
  if (lastRow <= 2) {
    currentRow = 3;
  } else {
    currentRow = lastRow + 2;
  }

  for (var i = 0; i < rosterData.length; i++) {
    var student = rosterData[i];

    currentRow = createStudentHeader(sheet, student, currentRow);
    currentRow = createLessonRows(sheet, student, currentRow, 5);

    if (i < rosterData.length - 1) {
      currentRow += 1;
    }
  }

  debugLog('createStudentSections', 'INFO', 'Created sections for ' + rosterData.length + ' students', '', '');
}

function createUtilityDebugSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheetName = "Debug";
    var debugSheet = ss.insertSheet(debugSheetName);
    
    // Set up headers matching your billing script format
    var headers = [
      "Timestamp",
      "Function",
      "Event Type", 
      "Message",
      "Data",
      "Error Details"
    ];
    
    debugSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format headers
    var headerRange = debugSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground(STYLES.HEADER.background);
    headerRange.setFontColor(STYLES.HEADER.text);
    
    // Set column widths for better readability
    debugSheet.setColumnWidth(1, 150); // Timestamp
    debugSheet.setColumnWidth(2, 200); // Function
    debugSheet.setColumnWidth(3, 120); // Event Type
    debugSheet.setColumnWidth(4, 300); // Message
    debugSheet.setColumnWidth(5, 200); // Data
    debugSheet.setColumnWidth(6, 250); // Error Details
    
    // Freeze header row
    debugSheet.setFrozenRows(1);
    
    Logger.log("✅ Utility debug sheet created successfully");
    return debugSheet;
    
  } catch (error) {
    Logger.log("❌ Error creating utility debug sheet: " + error.message);
    return null;
  }
}

function debugLog(functionName, eventType, message, data, errorDetails) {
  try {
    // Handle backward compatibility - if called with single parameter (old way)
    if (arguments.length === 1 && typeof functionName === 'string') {
      // Legacy call - treat as message only
      var legacyMessage = functionName;
      functionName = "LEGACY_CALL";
      eventType = "DEBUG";
      message = legacyMessage;
      data = "";
      errorDetails = "";
    }
    
    // Set default values if not provided
    functionName = functionName || "UNKNOWN_FUNCTION";
    eventType = eventType || "DEBUG";
    message = message || "";
    data = data || "";
    errorDetails = errorDetails || "";
    
    // Convert objects and handle special values for message
    var messageStr = formatLogValue(message);
    var dataStr = formatLogValue(data);
    var errorStr = formatLogValue(errorDetails);
    
    // Truncate very long values to prevent sheet errors
    messageStr = truncateString(messageStr, 1000);
    dataStr = truncateString(dataStr, 800);
    errorStr = truncateString(errorStr, 800);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName("Debug");
    
    if (!debugSheet) {
      debugSheet = createUtilityDebugSheet();
    }
    
    var timestamp = new Date();
    var logData = [timestamp, functionName, eventType, messageStr, dataStr, errorStr];
    
    try {
      debugSheet.appendRow(logData);
    } catch (appendError) {
      // Fallback: try to clear some space and retry once
      Logger.log("⚠️ Debug sheet append failed, attempting to clear old entries: " + appendError.message);
      try {
        clearOldDebugEntries(debugSheet);
        debugSheet.appendRow(logData);
      } catch (retryError) {
        Logger.log("❌ Debug sheet retry failed: " + retryError.message);
      }
    }
    
    // Also log to Google Apps Script logger with emoji prefixes
    var logMessage = "[" + eventType + "] " + functionName + ": " + messageStr;
    if (eventType === "ERROR") {
      Logger.log("❌ " + logMessage + (errorStr ? " | Error: " + errorStr : ""));
    } else if (eventType === "WARNING") {
      Logger.log("⚠️ " + logMessage);
    } else if (eventType === "INFO") {
      Logger.log("ℹ️ " + logMessage);
    } else {
      Logger.log("🔍 " + logMessage);
    }
    
  } catch (error) {
    Logger.log("❌ Error in debugLog function: " + error.message);
  }
}

function deleteExtraColumns(sheet, keepColumns) {
  var totalColumns = sheet.getMaxColumns();
  if (totalColumns > keepColumns) {
    sheet.deleteColumns(keepColumns + 1, totalColumns - keepColumns);
  }
}

function determineIfStudentIsAdult(studentData) {
  // Check the Age field (should be "Adult" or "Student")
  if (studentData.age) {
    var age = studentData.age.toString().toLowerCase();
    if (age === 'adult') {
      return true;
    }
    if (age === 'student' || age === 'child') {
      return false;
    }
  }
  
  // Default to child if Age field is missing or unclear
  return false;
}

function documentAlreadyExists(fileName, folder) {
  try {
    var files = folder.getFilesByName(fileName);
    return files.hasNext();
  } catch (error) {
    debugLog('documentAlreadyExists', 'ERROR', 'Failed to check document existence', fileName, error.message);
    return false; // If we can't check, assume it doesn't exist and proceed
  }
}

function enableDatePickerForColumn(sheet, columnIndex, startRow) {
  startRow = startRow || 2;

  var lastRow = sheet.getMaxRows();
  var range = sheet.getRange(startRow, columnIndex, lastRow - startRow + 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  range.setDataValidation(rule);
}

function executeWithErrorHandling(operation, successMessage, context, options) {
  options = options || {};
  var showUI = options.showUI !== undefined ? options.showUI : true;
  var logLevel = options.logLevel || 'INFO';
  
  try {
    var result = operation();
    
    // Log success
    if (logLevel !== 'NONE') {
      debugLog('executeWithErrorHandling', logLevel, successMessage, context, '');
    }
    
    // Show UI feedback if requested
    if (showUI && successMessage) {
      SpreadsheetApp.getUi().alert('Success', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return {
      success: true,
      data: result,
      message: successMessage
    };
    
  } catch (error) {
    var errorMessage = error.message || 'Unknown error occurred';
    
    // Log error with stack trace
    debugLog('executeWithErrorHandling', 'ERROR', errorMessage, context, error.stack || '');
    
    // Show UI error if requested
    if (showUI) {
      SpreadsheetApp.getUi().alert('Error', errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    return {
      success: false,
      error: errorMessage,
      data: null
    };
  }
}

function extractLessonQuantityFromPackage(packageText) {
  try {
    if (!packageText || typeof packageText !== 'string' || packageText.trim() === '') {
      return 0;
    }
    
    // Try format: (5 lessons $260)
    var match = packageText.match(/\((\d+)\s*lessons/i);
    
    if (!match) {
      // Try alternate format: : 15 lessons ($487.50)
      match = packageText.match(/:\s*(\d+)\s*lessons/i);
    }
    
    if (match && match[1]) {
      return parseInt(match[1], 10);
    }
    
    return 0;
    
  } catch (error) {
    return 0;
  }
}

function extractNumericLessonLength(lessonLengthValue) {
  var value = lessonLengthValue || 30;
  
  if (typeof value === 'string' && value.indexOf(' minutes') !== -1) {
    return parseInt(value.replace(' minutes', '')) || 30;
  } else if (typeof value === 'string') {
    return parseInt(value) || 30;
  } else {
    return value;
  }
}

function extractRosterData(rosterSheet) {
  var data = rosterSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Find column indices with fallback handling
  var getCol = function(name) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toString().toLowerCase().indexOf(name.toLowerCase()) !== -1) {
        return i;
      }
    }
    return -1;
  };
  
  var students = [];
  
  // Process each row (skip header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Skip empty rows - check both last name and first name
    var lastNameCol = getCol('Last Name');
    var firstNameCol = getCol('First Name');
    
    if (lastNameCol === -1 || firstNameCol === -1 || 
        !row[lastNameCol] || !row[firstNameCol]) {
      continue;
    }
    
    var student = {
      id: row[getCol('Student ID')] || '',
      lastName: row[lastNameCol] || '',
      firstName: row[firstNameCol] || '',
      instrument: row[getCol('Instrument')] || '',
      lessonLength: row[getCol('Length')] || 30,
      lessonsRegistered: row[getCol('Quantity')] || 0,
      lessonsCompleted: 0, // Will be calculated from lesson rows
      lessonsRemaining: row[getCol('Quantity')] || 0,
      status: row[getCol('Status')] || 'active'
    };
    
    // Only include active students (not dropped)
    if (student.status.toString().toLowerCase() !== 'dropped') {
      students.push(student);
    }
  }
  
  debugLog('extractRosterData', 'INFO', 'Extracted data for active students', students.length + ' students', '');
  return students;
}

function extractSeasonFromSemester(semesterName) {
  try {
    if (!semesterName || typeof semesterName !== 'string') {
      debugLog("extractSeasonFromSemester", "ERROR", "Invalid semester name", 
                            "Input: " + semesterName, "");
      return null;
    }
    
    var trimmed = String(semesterName).replace(/^\s+|\s+$/g, ''); // ES5 trim equivalent
    if (trimmed === '') {
      debugLog("extractSeasonFromSemester", "ERROR", "Empty semester name", "", "");
      return null;
    }
    
    // Split by space and take first part
    var parts = trimmed.split(' ');
    if (parts.length < 2) {
      debugLog("extractSeasonFromSemester", "ERROR", "Semester name format invalid", 
                            "Expected format: 'Season YYYY', got: " + semesterName, "");
      return null;
    }
    
    var season = parts[0];
    
    // Validate that it looks like a season (ES5 compatible check)
    var validSeasons = ['Spring', 'Summer', 'Fall', 'Winter'];
    var isValidSeason = false;
    for (var i = 0; i < validSeasons.length; i++) {
      if (validSeasons[i] === season) {
        isValidSeason = true;
        break;
      }
    }
    
    if (!isValidSeason) {
      debugLog("extractSeasonFromSemester", "WARNING", "Unrecognized season name", 
                            "Season: " + season + ", Expected one of: " + validSeasons.join(', '), "");
      // Return it anyway in case there are custom season names
    }
    
    debugLog("extractSeasonFromSemester", "INFO", "Extracted season", 
                          "Input: " + semesterName + ", Season: " + season, "");
    return season;
    
  } catch (error) {
    debugLog("extractSeasonFromSemester", "ERROR", "Error extracting season", 
                          "Input: " + semesterName, error.message);
    return null;
  }
}

function extractTeacherIdFromWorkbook(workbook) {
  try {
    var title = workbook.getName();
    var match = title.match(/QAMP\s+\d{4}\s+\S+\s+(T\d+)/);
    if (match && match[1]) {
      debugLog('extractTeacherIdFromWorkbook', 'DEBUG', 'Extracted teacher ID', 'ID: ' + match[1], '');
      return match[1];
    }
    debugLog('extractTeacherIdFromWorkbook', 'WARNING', 'Could not extract teacher ID from workbook title', 'Title: ' + title, '');
    return null;
  } catch (error) {
    debugLog('extractTeacherIdFromWorkbook', 'ERROR', 'Error extracting teacher ID', '', error.message);
    return null;
  }
}

function extractTeacherNameFromWorkbook(workbook) {
  try {
    // Extract teacher name from workbook title
    // Format is typically "QAMP [year] [teachername]"
    var title = workbook.getName();
    var match = title.match(/QAMP\s+\d{4}\s+(.+)/);
    
    if (match && match[1]) {
      var teacherName = match[1].trim();
      debugLog('extractTeacherNameFromWorkbook', 'DEBUG', 'Extracted teacher name', 'Name: ' + teacherName, '');
      return teacherName;
    }
    
    debugLog('extractTeacherNameFromWorkbook', 'WARN', 'Could not extract teacher name from workbook title', 'Title: ' + title, '');
    return null;
    
  } catch (error) {
    debugLog('extractTeacherNameFromWorkbook', 'ERROR', 'Error extracting teacher name', '', error.message);
    return null;
  }
}

function extractTotalLessonsFromPackages(qty30, qty45, qty60) {
  var total30 = extractLessonQuantityFromPackage(qty30) || 0;
  var total45 = extractLessonQuantityFromPackage(qty45) || 0;
  var total60 = extractLessonQuantityFromPackage(qty60) || 0;
  
  return total30 + total45 + total60;
}

function findColumnByPartialName(headers, searchTerm) {
  if (!headers || !searchTerm) return -1;

  var search = searchTerm.toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().toLowerCase().indexOf(search) !== -1) {
      return i + 1; // 1-based index
    }
  }
  return -1;
}

function findMostRecentRosterSheet(workbook) {
  try {
    var allSheets = workbook.getSheets();
    var rosterSheets = [];

    for (var i = 0; i < allSheets.length; i++) {
      var sheetName = allSheets[i].getName();
      if (sheetName.indexOf(' Roster') !== -1) {
        var season = sheetName.replace(' Roster', '').trim();
        rosterSheets.push({ sheet: allSheets[i], season: season });
      }
    }

    if (rosterSheets.length === 0) {
      debugLog('findMostRecentRosterSheet', 'INFO', 'No roster sheets found in workbook', '', '');
      return null;
    }

    var semesterMetadataSheet = getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      debugLog('findMostRecentRosterSheet', 'WARNING', 'Semester Metadata sheet not found - using first roster', '', '');
      return rosterSheets[0].sheet;
    }

    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol = -1;
    var startCol = -1;

    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      }
    }

    if (nameCol === -1 || startCol === -1) {
      debugLog('findMostRecentRosterSheet', 'WARNING', 'Required columns not found in Semester Metadata - using first roster', '', '');
      return rosterSheets[0].sheet;
    }

    var semesters = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[startCol]) {
        var semesterName = row[nameCol].toString().trim();
        var season = extractSeasonFromSemester(semesterName);
        if (season) {
          semesters.push({ season: season, semesterName: semesterName, startDate: new Date(row[startCol]) });
        }
      }
    }

    semesters.sort(function(a, b) { return b.startDate - a.startDate; });

    for (var i = 0; i < semesters.length; i++) {
      var semester = semesters[i];
      for (var j = 0; j < rosterSheets.length; j++) {
        if (rosterSheets[j].season === semester.season) {
          debugLog('findMostRecentRosterSheet', 'SUCCESS', 'Found most recent roster',
            semester.season + ' Roster (' + semester.semesterName + ')', '');
          return rosterSheets[j].sheet;
        }
      } // closes inner for
    } // closes outer for

    debugLog('findMostRecentRosterSheet', 'WARNING', 'No semester match found - using first available roster',
      rosterSheets[0].season + ' Roster', '');
    return rosterSheets[0].sheet;

  } catch (error) {
    debugLog('findMostRecentRosterSheet', 'ERROR', 'Error finding roster sheet', '', error.message);
    return null;
  }
}

function findParentRow(parentsSheet, parentId, fallbackKey) {
  try {
    var data = parentsSheet.getDataRange().getValues();
    var headers = data[0];
    var idCol = -1;
    var lookupCol = -1;

    for (var i = 0; i < headers.length; i++) {
      var normalizedHeader = normalizeHeader(headers[i]);
      if (normalizedHeader === 'parentid') idCol = i;
      if (normalizedHeader === 'parentlookup') lookupCol = i;
    }

    if (idCol === -1 || lookupCol === -1) {
      debugLog('findParentRow', 'ERROR', 'Missing required columns',
        'Parent ID col: ' + idCol + ', Parent Lookup col: ' + lookupCol, '');
      throw new Error("Missing 'Parent ID' or 'Parent Lookup' column in Parents sheet.");
    }

    debugLog('findParentRow', 'DEBUG', 'Searching for parent',
      'Parent ID: "' + parentId + '", Key: "' + fallbackKey + '", Rows to check: ' + (data.length - 1), '');

    for (var j = 1; j < data.length; j++) {
      var rowId = (data[j][idCol] || '').toString().trim();
      var rowKey = (data[j][lookupCol] || '').toString().toLowerCase().trim();

      if (rowId === parentId || rowKey === fallbackKey) {
        debugLog('findParentRow', 'DEBUG', 'Found parent match', 'Row: ' + (j + 1) + ', ID: ' + rowId, '');
        return j + 1;
      }
    }

    debugLog('findParentRow', 'DEBUG', 'No parent match found',
      'Parent ID: "' + parentId + '", Key: "' + fallbackKey + '"', '');
    return -1;

  } catch (error) {
    debugLog('findParentRow', 'ERROR', 'Failed', '', error.message);
    throw error;
  }
}

function findRosterSheetForMonth(workbook, targetMonthName) {
  try {
    var parts = targetMonthName.split(' ');
    if (parts.length !== 2) {
      debugLog('findRosterSheetForMonth', 'ERROR', 'Invalid targetMonthName format', targetMonthName, '');
      return null;
    }

    var monthIndex = MONTH_NAMES.indexOf(parts[0]);
    var year = parseInt(parts[1]);
    if (monthIndex === -1 || isNaN(year)) {
      debugLog('findRosterSheetForMonth', 'ERROR', 'Could not parse month/year', targetMonthName, '');
      return null;
    }

    var targetDate = new Date(year, monthIndex, 1);

    var semesterMetadataSheet = getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      debugLog('findRosterSheetForMonth', 'ERROR', 'Semester Metadata sheet not found', '', '');
      return null;
    }

    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];
    var headerMap = getHeaderMap(semesterMetadataSheet);
    var norm = normalizeHeader;

    var nameCol = headerMap[norm('Semester Name')];
    var startCol = headerMap[norm('Start Date')];
    var endCol = headerMap[norm('End Date')];

    if (!nameCol || !startCol || !endCol) {
      debugLog('findRosterSheetForMonth', 'ERROR', 'Required columns not found in Semester Metadata', '', '');
      return null;
    }

    var matchedSemester = null;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var semesterName = (row[nameCol - 1] || '').toString().trim();
      var startDate = new Date(row[startCol - 1]);
      var endDate = new Date(row[endCol - 1]);

      if (!semesterName || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) continue;

      if (targetDate >= startDate && targetDate <= endDate) {
        matchedSemester = semesterName;
        break;
      }
    }

    if (!matchedSemester) {
      debugLog('findRosterSheetForMonth', 'WARNING', 'No semester found for month', targetMonthName, '');
      return null;
    }

    var season = extractSeasonFromSemester(matchedSemester);
    if (!season) {
      debugLog('findRosterSheetForMonth', 'WARNING', 'Could not extract season from semester', matchedSemester, '');
      return null;
    }

    var rosterSheetName = season + ' Roster';
    var rosterSheet = workbook.getSheetByName(rosterSheetName);

    if (!rosterSheet) {
      debugLog('findRosterSheetForMonth', 'INFO', 'No matching roster sheet found', rosterSheetName, '');
      return null;
    }

    debugLog('findRosterSheetForMonth', 'SUCCESS', 'Found roster sheet', rosterSheetName + ' for ' + targetMonthName, '');
    return rosterSheet;

  } catch (error) {
    debugLog('findRosterSheetForMonth', 'ERROR', 'Error finding roster sheet', targetMonthName, error.message);
    return null;
  }
}

function findStudentInContacts(contactsData, studentIdCol, targetStudentId) {
  for (var i = 1; i < contactsData.length; i++) {
    if (contactsData[i][studentIdCol - 1] === targetStudentId) {
      return i;
    }
  }
  return -1;
}

function findStudentRow(studentSheet, studentKey) {
  var data = studentSheet.getDataRange().getValues();
  var headerRow = data[0];
  var lookupColIndex = -1;
  
  for (var i = 0; i < headerRow.length; i++) {
    if (normalizeHeader(headerRow[i]) === 'studentlookup') {
      lookupColIndex = i;
      break;
    }
  }
  
  if (lookupColIndex === -1) {
    throw new Error("Missing 'Student Lookup' column in Students sheet.");
  }
  
  for (var j = 1; j < data.length; j++) {
    var rowKey = data[j][lookupColIndex].toString().trim().toLowerCase();
    if (rowKey === studentKey) {
      return j + 1;
    }
  }
  return -1;
}

function findTeacherInRosterLookup(lookupSheet, teacherId) {
  try {
    var getCol = createColumnFinder(lookupSheet);
    var teacherIdCol = getCol('Teacher ID');

    if (!teacherIdCol) {
      debugLog('findTeacherInRosterLookup', 'ERROR', 'Teacher ID column not found', '', '');
      return -1;
    }

    var data = lookupSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][teacherIdCol - 1]).trim() === String(teacherId).trim()) {
        return i + 1;
      }
    }

    return -1;

  } catch (error) {
    debugLog('findTeacherInRosterLookup', 'ERROR', 'Failed', teacherId, error.message);
    return -1;
  }
}

function formatAddress(street, city, zip) {
  var parts = [];
  if (street) parts.push(street);
  
  var cityStateZip = [];
  if (city) cityStateZip.push(city);
  cityStateZip.push("NY");
  if (zip) cityStateZip.push(zip);
  
  if (cityStateZip.length > 1) {
    parts.push(cityStateZip.join(', '));
  }
  
  return parts.join('\n');
}

function formatAttendanceColumns(sheet, studentCount) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    for (var i = 0; i < headers.length; i++) {
      var key = headers[i].toLowerCase();
      var col = i + 1;
      if (key.includes('add row')) sheet.setColumnWidth(col, 60);
      else if (key.includes('last name') || key.includes('first name')) sheet.setColumnWidth(col, 120);
      else if (key.includes('instrument')) sheet.setColumnWidth(col, 90);
      else if (key.includes('student id')) sheet.setColumnWidth(col, 100);
      else if (key.includes('lessons completed') || key.includes('lessons registered')) sheet.setColumnWidth(col, 80);
      else if (key.includes('status')) sheet.setColumnWidth(col, 90);
      else if (key.includes('drop date')) sheet.setColumnWidth(col, 90);
      else if (key.includes('comments')) sheet.setColumnWidth(col, 250);
      else if (key.match(/\d{1,2} \- \d{1,2}/)) sheet.setColumnWidth(col, 100); // Week range
    }

    if (lastRow > 1 && lastCol > 0) {
      var bandings = sheet.getBandings();
      for (var j = 0; j < bandings.length; j++) {
        bandings[j].remove();
      }
      var bandingRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      var banding = bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding.setHeaderRowColor(null)
             .setFirstRowColor(STYLES.ALTERNATING_LIGHT.background)
             .setSecondRowColor(STYLES.ALTERNATING_DARK.background);
    }
  } catch (e) {
    debugLog('formatAttendanceColumns', 'ERROR', 'Formatting failed', '', e.message);
  }
}

function formatAttendanceSheet(sheet) {
  try {
    sheet.setFrozenRows(1);

    var maxRows = sheet.getMaxRows();

    // Date format and validation for column C — start at row 3 to preserve sign-off row
    var dateRange = sheet.getRange(3, 3, maxRows - 2, 1);
    dateRange.setNumberFormat('ddd, M/d');

    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .build();
    dateRange.setDataValidation(dateRule);

    // Status dropdown for column E — start at row 3 to preserve sign-off row
    var statusOptions = ['Lesson', 'No Show', 'No Lesson'];
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(statusOptions)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(3, 5, maxRows - 2, 1).setDataValidation(statusRule);

    // Protect admin columns G-L with warning
    var adminRange = sheet.getRange(1, 7, maxRows, 6);
    var protection = adminRange.protect();
    protection.setDescription('Admin columns - automated data');
    protection.setWarningOnly(true);

    // Text wrapping for all columns
    sheet.getRange(1, 1, maxRows, 12).setWrap(true);

    debugLog('formatAttendanceSheet', 'INFO', 'Attendance sheet formatted', '', '');

  } catch (error) {
    debugLog('formatAttendanceSheet', 'ERROR', 'Failed to format attendance sheet', '', error.message);
    throw error;
  }
}

function formatCurrency(amount) {
  if (isNaN(amount)) return '$0.00';
  return '$' + amount.toFixed(2);
}

function formatDateFlexible(date, format) {
  format = format || 'MMM d';
  // Fix the same instanceof Date issue
  if (!date || Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date.getTime())) {
    return '';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}

function formatLessonLengthWithMinutes(lessonLengthValue) {
  var value = lessonLengthValue || 30;
  
  if (typeof value === 'string' && value.indexOf(' minutes') !== -1) {
    return value;
  } else {
    return value + ' minutes';
  }
}

function formatLogValue(value) {
  if (typeof value === 'object' && value !== null) {
    try {
      return JSON.stringify(value);
    } catch (jsonError) {
      return '[Object - JSON Error: ' + jsonError.message + ']';
    }
  } else if (typeof value === 'undefined') {
    return 'undefined';
  } else if (value === null) {
    return 'null';
  } else {
    var strValue = String(value);
    // CRITICAL FIX: Escape strings that start with = to prevent formula interpretation
    // When a string starting with = is written to a sheet cell, Google Sheets
    // interprets it as a formula. Prepending a single quote forces text interpretation.
    if (strValue.length > 0 && strValue.charAt(0) === '=') {
      return "'" + strValue;  // Prepend single quote to escape the formula
    }
    return strValue;
  }
}

function formatPhoneNumber(phoneRaw) {
  var digits = phoneRaw.toString().replace(/\D/g, '');
  if (digits.length === 10) {
    return '(' + digits.slice(0, 3) + ') ' + digits.slice(3, 6) + '-' + digits.slice(6);
  }
  return phoneRaw;
}

function formatRosterColumns(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    if (lastRow > 1 && lastCol > 0) {
      var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
      dataRange.setVerticalAlignment("top").setWrap(true).setFontWeight("normal");

      sheet.getRange(2, 1, lastRow - 1, 1).setFontColor("#5f6368").setHorizontalAlignment("center"); // Checkbox
      if (lastCol > 1) {
        sheet.getRange(2, 2, lastRow - 1, lastCol - 1).setFontColor("#000000");
      }

      sheet.getRange(2, 8, lastRow - 1, 1).setHorizontalAlignment("center"); // Parent Last Name
      sheet.getRange(2, 9, lastRow - 1, 1).setHorizontalAlignment("center"); // Parent First Name
    }

    var columnWidths = [100, 120, 100, 100, 200, 120, 120, 80, 80, 120, 120, 120, 80, 120, 120, 150, 200];
    for (var i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }
    sheet.setColumnWidth(16, 220); // Emphasize column 16

    var bandings = sheet.getBandings();
    for (var j = 0; j < bandings.length; j++) {
      bandings[j].remove();
    }

    if (lastRow >= 2 && lastCol > 0) {
      var banding = sheet.getRange(2, 1, lastRow - 1, lastCol)
        .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding.setHeaderRowColor(null)
             .setFirstRowColor(STYLES.ALTERNATING_LIGHT.background)
             .setSecondRowColor(STYLES.ALTERNATING_DARK.background);
    }
  } catch (e) {
    debugLog('formatRosterColumns', 'ERROR', 'Formatting failed', '', e.message);
  }
}

function freezeSheetFormulas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange();
  var formulas = range.getFormulas();
  var values = range.getValues();

  for (var row = 0; row < formulas.length; row++) {
    for (var col = 0; col < formulas[0].length; col++) {
      if (formulas[row][col]) {
        sheet.getRange(row + 1, col + 1).setValue(values[row][col]);
      }
    }
  }
}

function generateDocumentFromTemplate(templateKey, variableData, fileName, destinationFolder) {
  try {
    // Get the template file
    var templateFile = getTemplate(templateKey);
    
    // Make copy of template
    var docCopy = templateFile.makeCopy(fileName, destinationFolder);
    var newDoc = DocumentApp.openById(docCopy.getId());
    var body = newDoc.getBody();
    
    // Replace all template variables
    for (var variable in variableData) {
      var placeholder = '{{' + variable + '}}';
      var value = variableData[variable] || '';
      
      // Convert value to string if it's not already
      if (typeof value !== 'string') {
        value = value.toString();
      }
      
      body.replaceText(placeholder, value);
    }
    
    // Save and close
    newDoc.saveAndClose();
    
    return {
      success: true,
      fileId: docCopy.getId(),
      url: docCopy.getUrl(),
      fileName: fileName
    };
    
  } catch (error) {
    debugLog('generateDocumentFromTemplate', 'ERROR', 'Failed to generate document', '', error.message);
    return {
      success: false,
      error: error.message,
      fileId: null,
      url: null
    };
  }
}

function generateKey() {
  var fields = Array.prototype.slice.call(arguments);
  var result = [];
  for (var i = 0; i < fields.length; i++) {
    var field = fields[i] || '';
    result.push(field.toString().trim().toLowerCase());
  }
  return result.join('_');
}

function generateNextId(sheet, columnName, prefix, recordName) {
  // Check if historical data input is enabled
  if (isHistoricalDataInputEnabled()) {
    return promptForHistoricalId(sheet, columnName, prefix, recordName);
  }
  
  // Original auto-generation logic
  return generateNextIdDirect(sheet, columnName, prefix);
}

function generateNextIdDirect(sheet, columnName, prefix) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = -1;
  
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === normalizeHeader(columnName)) {
      idCol = i;
      break;
    }
  }
  
  if (idCol === -1) throw new Error("Missing '" + columnName + "' column.");
  
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var val = data[i][idCol];
    if (val && typeof val === 'string') {
      var match = val.match(new RegExp("^" + prefix + "(\\d+)$"));
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
  }
  
  var nextNum = maxNum + 1;
  var paddedNum = ('0000' + nextNum).slice(-4);
  return prefix + paddedNum;
}

function getAttendanceSheetForDate(ss, targetDate) {
  if (!targetDate || !(targetDate instanceof Date)) {
    return null;
  }
  
  try {
    // First try the target month
    var targetMonthName = getMonthNameFromDate(targetDate);
    if (targetMonthName) {
      var targetSheet = ss.getSheetByName(targetMonthName);
      if (targetSheet) {
        return targetSheet;
      }
    }
    
    // Fallback to previous month if target month not found
    var previousMonth = new Date(targetDate);
    previousMonth.setMonth(targetDate.getMonth() - 1);
    
    var previousMonthName = getMonthNameFromDate(previousMonth);
    if (previousMonthName) {
      var previousSheet = ss.getSheetByName(previousMonthName);
      if (previousSheet) {
        return previousSheet;
      }
    }
    
    // If both fail, try to find any attendance sheet
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (isMonthSheet(sheets[i].getName())) {
        return sheets[i];
      }
    }
    
    return null;
    
  } catch (error) {
    debugLog('getAttendanceSheetForDate', 'ERROR', 'Error getting attendance sheet', 'Target date: ' + targetDate, error.message);
    return null;
  }
}

function getCached(key) {
  return _executionCache[key];
}

function getColumnHeaders(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var result = [];
  for (var i = 0; i < headers.length; i++) {
    result.push(headers[i] ? headers[i].toString().trim() : '');
  }
  return result;
}

function getColumnIndices(sheet, columnNames) {
  if (!sheet || !columnNames) return {};
  
  // Create cache key from sheet ID and column names
  var sheetId = sheet.getParent().getId() + '_' + sheet.getName();
  var cacheKey = 'colIdx_' + sheetId + '_' + columnNames.join('_');
  var cached = getCached(cacheKey);
  if (cached !== undefined) return cached;
  
  // Get headers and find indices
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var indices = {};
  
  for (var i = 0; i < headers.length; i++) {
    var normalizedHeader = normalizeHeader(headers[i]);
    if (columnNames.indexOf(normalizedHeader) !== -1) {
      indices[normalizedHeader] = i;
    }
  }
  
  // Cache and return
  return setCached(cacheKey, indices);
}

function getCurrentAcademicYearInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var semesterSheet = ss.getSheetByName('Semester Metadata');
    
    if (semesterSheet) {
      var data = semesterSheet.getDataRange().getValues();
      if (data.length > 1) {
        var latestSemester = data[data.length - 1];
        var semesterName = latestSemester[0];
        var startDate = latestSemester[1];
        
        // Extract academic year from semester name (e.g., "Fall 2024" -> "2024-2025")
        var year = semesterName.match(/\d{4}/);
        if (year) {
          var currentYear = parseInt(year[0], 10);
          var academicYear = currentYear + '-' + (currentYear + 1);
          var formattedStartDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), 'MMMM d, yyyy');
          
          return {
            academicYear: academicYear,
            startDate: formattedStartDate,
            semesterName: semesterName
          };
        }
      }
    }
    
    // Fallback
    var currentYear = new Date().getFullYear();
    return {
      academicYear: currentYear + '-' + (currentYear + 1),
      startDate: 'September 1, ' + currentYear,
      semesterName: 'Current Semester'
    };
    
  } catch (error) {
    var currentYear = new Date().getFullYear();
    return {
      academicYear: currentYear + '-' + (currentYear + 1),
      startDate: 'September 1, ' + currentYear,
      semesterName: 'Current Semester'
    };
  }
}

function getCurrentSemesterMonth(semesterName) {
  try {
    var semesterMetadataSheet = getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      return 'January';
    }

    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];

    var nameCol = -1, startCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      }
    }

    if (nameCol === -1 || startCol === -1) {
      return 'January';
    }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[nameCol].toString().trim() === semesterName) {
        var startDate = new Date(row[startCol]);
        return MONTH_NAMES[startDate.getMonth()];
      }
    }

    return 'January';

  } catch (error) {
    return 'January';
  }
}

function getCurrentSemesterName(asOfDate) {
  try {
    var semesterSheet = getSheet('semesterMetadata');
    if (!semesterSheet) {
      debugLog('getCurrentSemesterName', 'WARNING', 'Semester Metadata sheet not found', '', '');
      return null;
    }

    var data = semesterSheet.getDataRange().getValues();
    if (data.length < 2) {
      debugLog('getCurrentSemesterName', 'WARNING', 'No semester data found in metadata', '', '');
      return null;
    }

    var headers = data[0];
    var nameCol = -1, startCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      if (header.indexOf('start') !== -1 || header.indexOf('begin') !== -1) {
        startCol = i;
      } else if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      }
    }

    var resolvedNameCol = nameCol !== -1 ? nameCol : 0;

    // No date provided or no startCol — fall back to last row
    if (!asOfDate || startCol === -1) {
      var currentSemester = data[data.length - 1][resolvedNameCol];
      if (!currentSemester || String(currentSemester).trim() === '') {
        debugLog('getCurrentSemesterName', 'WARNING', 'Empty semester name in metadata', '', '');
        return null;
      }
      debugLog('getCurrentSemesterName', 'INFO', 'Found semester (last row fallback)', String(currentSemester).trim(), '');
      return String(currentSemester).trim();
    }

    var compareDate = new Date(asOfDate);
    compareDate.setHours(0, 0, 0, 0);
    var result = null;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var semesterName = row[resolvedNameCol];
      if (!semesterName || String(semesterName).trim() === '') continue;

      var startDate = new Date(row[startCol]);
      startDate.setHours(0, 0, 0, 0);

      if (startDate <= compareDate) {
        result = String(semesterName).trim();
      }
    }

    if (!result) {
      result = String(data[1][resolvedNameCol]).trim();
      debugLog('getCurrentSemesterName', 'WARNING', 'No started semester found, using first', result, '');
    } else {
      debugLog('getCurrentSemesterName', 'INFO', 'Found active semester by date', result, '');
    }

    return result;

  } catch (error) {
    debugLog('getCurrentSemesterName', 'ERROR', 'Failed to get current semester', '', error.message);
    return null;
  }
}

function getDateForWeekday(weekStartDate, weekdayName) {
  if (!weekStartDate || !weekdayName) return null;

  var dayMap = {
    sunday: 0,
    monday: 1,
    tuesday: 2,
    wednesday: 3,
    thursday: 4,
    friday: 5,
    saturday: 6
  };

  var targetDay = dayMap[weekdayName.toLowerCase()];
  if (targetDay === undefined) return null;

  var startDay = weekStartDate.getDay(); // 0 (Sun) to 6 (Sat)
  var diff = targetDay - startDay;

  var resultDate = new Date(weekStartDate);
  resultDate.setDate(weekStartDate.getDate() + diff);
  return resultDate;
}

function getFieldMappingFromSheet(fieldMapSheet) {
  var data = fieldMapSheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var formHeader = normalizeHeader(data[i][0]);
    var internalName = data[i][1].toString().trim();
    if (formHeader && internalName) {
      map[formHeader] = internalName;
    }
  }
  return map;
}

function getGeneratedDocumentsFolder() {
  var env = EnvironmentManager.get();
  var folderId = CONFIG && CONFIG[env] && CONFIG[env].generatedDocumentsFolderId;
  if (!folderId) {
    throw new Error('❌ generatedDocumentsFolderId not found in config for environment "' + env + '"');
  }
  return DriveApp.getFolderById(folderId);
}

function getHeaderMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[normalizeHeader(headers[i])] = i + 1;
  }
  return map;
}

function getInstrumentFamily(instrument) {
  if (!instrument) return '';
  var normalized = instrument.toString().trim().toLowerCase();
  for (var family in INSTRUMENT_FAMILIES) {
    var members = INSTRUMENT_FAMILIES[family];
    for (var i = 0; i < members.length; i++) {
      if (members[i].toLowerCase() === normalized) {
        return family;
      }
    }
  }
  return '';
}

function getLessonLengthFromPackages(qty30Package, qty45Package, qty60Package) {
  // This function determines which lesson length was selected and returns that length
  try {
    var qty30 = extractLessonQuantityFromPackage(qty30Package);
    var qty45 = extractLessonQuantityFromPackage(qty45Package);
    var qty60 = extractLessonQuantityFromPackage(qty60Package);
    
    if (qty60 > 0) return 60;
    if (qty45 > 0) return 45;
    if (qty30 > 0) return 30;
    
    return 30; // Default
    
  } catch (error) {
    return 30;
  }
}

function getMonthNameFromDate(date, capitalize) {
  if (!date || typeof date.getMonth !== 'function') return null;
  var monthName = MONTH_NAMES[date.getMonth()];
  if (capitalize) return monthName;
  return monthName.toLowerCase();
}

function getMonthNames() {
  return MONTH_NAMES;
}

function getMonthSheets(ss) {
  if (!ss) return [];
  var ssId = ss.getId();
  var cacheKey = 'monthSheets_' + ssId;
  var cached = getCached(cacheKey);
  if (cached !== undefined) return cached;
  var sheets = ss.getSheets();
  var monthSheets = [];
  for (var i = 0; i < sheets.length; i++) {
    if (isMonthSheet(sheets[i].getName())) {
      monthSheets.push(sheets[i]);
    }
  }
  return setCached(cacheKey, monthSheets);
}

function getMostRecentRateColumn(headers) {
  var bestColIndex = -1;
  var bestStartYear = 0;
  var regex = /^(\d{4})-(\d{4})$/;

  for (var i = 0; i < headers.length; i++) {
    var match = regex.exec(headers[i]);
    if (match) {
      var startYear = parseInt(match[1], 10);
      if (startYear > bestStartYear) {
        bestStartYear = startYear;
        bestColIndex = i;
      }
    }
  }

  return bestColIndex;
}

function getNextSemester(currentSemesterName) {
  try {
    var semesterMetadataSheet = getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      return null;
    }
    
    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find columns
    var nameCol = -1, startCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      }
    }
    
    if (nameCol === -1 || startCol === -1) {
      return null;
    }
    
    // Build array of semesters with start dates
    var semesters = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[startCol]) {
        semesters.push({
          name: row[nameCol],
          startDate: new Date(row[startCol])
        });
      }
    }
    
    // Sort by start date
    semesters.sort(function(a, b) {
      return a.startDate - b.startDate;
    });
    
    // Find current semester and return next one
    for (var i = 0; i < semesters.length; i++) {
      if (semesters[i].name === currentSemesterName) {
        if (i < semesters.length - 1) {
          debugLog("getNextSemester", "INFO", 
                   "Found next semester", 
                   "Current: " + currentSemesterName + ", Next: " + semesters[i+1].name, "");
          return semesters[i+1].name;
        } else {
          debugLog("getNextSemester", "INFO", 
                   "No next semester (last semester)", 
                   currentSemesterName, "");
          return null;
        }
      }
    }
    
    return null;
    
  } catch (error) {
    debugLog("getNextSemester", "ERROR", "Error finding next semester", "", error.message);
    return null;
  }
}

function getOrCreateSubfolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function getRateSummary() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Rates');
  if (!sheet) return '⚠️ Rates sheet not found.';

  var data = sheet.getDataRange().getValues();
  if (data.length < 2 || data[0].length < 3) return '⚠️ Rates sheet is missing expected structure.';

  var headers = data[0];
  var bestColIndex = getMostRecentRateColumn(headers);
  if (bestColIndex === -1) return '⚠️ No valid rate year columns found.';

  var summary = ['📅 Current rates for ' + headers[bestColIndex] + ':'];
  for (var i = 1; i < data.length; i++) {
    var title = data[i][0];
    var value = data[i][bestColIndex];
    var rateType = data[i][headers.length - 1];

    if (value === '' || value === null || typeof value === 'undefined') continue;

    if (typeof value === 'number' && rateType === 'Currency') {
      value = '$' + value.toFixed(2);
    }

    summary.push(title + ': ' + value);
  }

  return summary.join('\n');
}

function getRosterFolder() {
  var env = EnvironmentManager.get();
  var folderId = CONFIG && CONFIG[env] && CONFIG[env].rosterFolderId;
  if (!folderId) {
    throw new Error('❌ rosterFolderId not found in config for environment "' + env + '"');
  }
  return DriveApp.getFolderById(folderId);
}

function getRosterFolderUrlForYear(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var yearSheet = ss.getSheetByName('Year Metadata');
  if (!yearSheet) return '';
  var data = yearSheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === year) {
      return data[i][1];
    }
  }
  return '';
}

function getSemesterDates(semesterName) {
  try {
    var semesterSheet = getSheet('semesterMetadata');
    if (!semesterSheet) {
      return null;
    }
    
    var data = semesterSheet.getDataRange().getValues();
    var headers = data[0];
    
    var nameCol = -1, startCol = -1, endCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      } else if (header.indexOf('end') !== -1) {
        endCol = i;
      }
    }
    
    if (nameCol === -1 || startCol === -1 || endCol === -1) {
      return null;
    }
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[nameCol].toString().trim() === semesterName) {
        var startDate = new Date(row[startCol]);
        var endDate = new Date(row[endCol]);
        
        return {
          start: startDate,
          end: endDate,
          name: semesterName
        };
      }
    }
    
    return null;
    
  } catch (error) {
    return null;
  }
}

function getSemesterForDate(targetDate) {
  try {
    var semesterMetadataSheet = getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      debugLog("getSemesterForDate", "ERROR", "Semester Metadata sheet not found", "", "");
      return null;
    }

    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];

    var nameCol = -1, startCol = -1, endCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1 || header.indexOf('begin') !== -1) {
        startCol = i;
      } else if (header.indexOf('end') !== -1) {
        endCol = i;
      }
    }

    if (nameCol === -1 || startCol === -1 || endCol === -1) {
      debugLog("getSemesterForDate", "ERROR", "Required columns not found", "", "");
      return null;
    }

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var semesterName = row[nameCol];
      var startDate = new Date(row[startCol]);
      var endDate = new Date(row[endCol]);

      if (targetDate >= startDate && targetDate <= endDate) {
        debugLog("getSemesterForDate", "INFO", "Found semester for date",
                 "Date: " + formatDateFlexible(targetDate, 'M/d/yy') + ", Semester: " + semesterName, "");
        return semesterName;
      }
    }

    debugLog("getSemesterForDate", "WARNING", "No semester found for date",
             formatDateFlexible(targetDate, 'M/d/yy'), "");
    return null;

  } catch (error) {
    debugLog("getSemesterForDate", "ERROR", "Error finding semester", "", error.message);
    return null;
  }
}

function getSheet(sheetKey) {
  var env = EnvironmentManager.get();
  var entry = SHEET_MAP && SHEET_MAP[sheetKey];
  if (!entry) {
    throw new Error('❌ Sheet key "' + sheetKey + '" not found in SHEET_MAP');
  }

  var workbookKey = inferWorkbookKey(sheetKey);
  var workbookId = CONFIG && CONFIG[env] && CONFIG[env][workbookKey + 'Id'];
  if (!workbookId) {
    throw new Error('❌ Workbook ID for sheet "' + sheetKey + '" not found in environment "' + env + '"');
  }

  var workbook = SpreadsheetApp.openById(workbookId);
  var sheet = workbook.getSheetByName(entry.name);
  if (!sheet) {
    throw new Error('❌ Sheet "' + entry.name + '" not found in workbook "' + workbookKey + '"');
  }

  return sheet;
}

function getStudentIdFromRow(row, headerMap) {
  var norm = normalizeHeader;
  var sidCol = headerMap[norm("Student ID")];
  var idCol = headerMap[norm("ID")];
  return sidCol ? row[sidCol - 1] : (idCol ? row[idCol - 1] : null);
}

function getTeacherGroupAssignments(teacherId) {
  try {
    var teacherLookupSheet = getSheet('teacherRosterLookup');
    if (!teacherLookupSheet) {
      debugLog('getTeacherGroupAssignments', 'ERROR', 'Teacher Roster Lookup sheet not found', '', '');
      return [];
    }

    var getCol = createColumnFinder(teacherLookupSheet);
    var teacherIdCol = getCol('Teacher ID');
    var groupAssignmentCol = getCol('Group Assignment');

    if (teacherIdCol === 0 || groupAssignmentCol === 0) {
      debugLog('getTeacherGroupAssignments', 'ERROR', 'Required columns not found',
               'Teacher ID col: ' + teacherIdCol + ', Group col: ' + groupAssignmentCol, '');
      return [];
    }

    var teacherData = teacherLookupSheet.getDataRange().getValues();
    var groupAssignmentRaw = '';

    for (var i = 1; i < teacherData.length; i++) {
      if (String(teacherData[i][teacherIdCol - 1]).trim() === String(teacherId).trim()) {
        groupAssignmentRaw = teacherData[i][groupAssignmentCol - 1] || '';
        break;
      }
    }

    if (!groupAssignmentRaw) {
      debugLog('getTeacherGroupAssignments', 'DEBUG', 'No group assignment found', 'Teacher ID: ' + teacherId, '');
      return [];
    }

    var assignedPrograms = groupAssignmentRaw.toString().split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p; });

    var billingWorkbook = getWorkbook('billing');
    var programsSheet = billingWorkbook.getSheetByName('Programs List');
    if (!programsSheet) {
      debugLog('getTeacherGroupAssignments', 'ERROR', 'Programs List sheet not found', '', '');
      return [];
    }

    var programsData = programsSheet.getDataRange().getValues();
    var getProgCol = createColumnFinder(programsSheet);
    var programNameCol = getProgCol('Program Name');
    var groupIdCol = getProgCol('Group ID');
    var typeCol = getProgCol('Type');
    var aliasCol = getProgCol('Alias For');
    var activeCol = getProgCol('Active');

    if (programNameCol === 0 || groupIdCol === 0) {
      debugLog('getTeacherGroupAssignments', 'ERROR', 'Required program columns not found', '', '');
      return [];
    }

    var groupEntries = [];

    for (var j = 0; j < assignedPrograms.length; j++) {
      var programName = assignedPrograms[j];

      for (var k = 1; k < programsData.length; k++) {
        var row = programsData[k];
        var rowProgramName = row[programNameCol - 1];
        var isActive = row[activeCol - 1] === true;
        var type = row[typeCol - 1];
        var groupId = row[groupIdCol - 1];
        var aliasFor = row[aliasCol - 1];

        if (!isActive) continue;

        if (rowProgramName === programName) {
          if (type === 'Package' && aliasFor) {
            var components = aliasFor.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c; });
            for (var c = 0; c < components.length; c++) {
              var componentName = components[c];
              for (var m = 1; m < programsData.length; m++) {
                var componentRow = programsData[m];
                if (componentRow[programNameCol - 1] === componentName &&
                    componentRow[activeCol - 1] === true &&
                    componentRow[groupIdCol - 1]) {
                  groupEntries.push({
                    groupId: componentRow[groupIdCol - 1],
                    groupName: componentName
                  });
                }
              }
            }
          } else if (groupId) {
            groupEntries.push({
              groupId: groupId,
              groupName: programName
            });
          }
          break;
        }
      }
    }

    debugLog('getTeacherGroupAssignments', 'INFO', 'Found group assignments',
             'Teacher ID: ' + teacherId + ', Groups: ' + groupEntries.length, '');
    return groupEntries;

  } catch (error) {
    debugLog('getTeacherGroupAssignments', 'ERROR', 'Error getting teacher group assignments',
             'Teacher ID: ' + teacherId, error.message);
    return [];
  }
}

function getTeacherIdByDisplayName(displayName) {
  try {
    if (!displayName || String(displayName).trim() === '') {
      debugLog('getTeacherIdByDisplayName', 'WARNING', 'No display name provided', '', '');
      return '';
    }

    var name = String(displayName).trim();

    // If it already looks like a Teacher ID, return as-is
    if (/^T\d+$/.test(name)) return name;

    var lookupSheet = getSheet('teacherRosterLookup');
    var getCol = createColumnFinder(lookupSheet);
    var displayNameCol = getCol('Display Name');
    var teacherIdCol   = getCol('Teacher ID');

    if (!displayNameCol || !teacherIdCol) {
      debugLog('getTeacherIdByDisplayName', 'ERROR', 'Required columns not found in Teacher Roster Lookup', '', '');
      return '';
    }

    var data = lookupSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][displayNameCol - 1]).trim() === name) {
        var teacherId = String(data[i][teacherIdCol - 1]).trim();
        debugLog('getTeacherIdByDisplayName', 'SUCCESS', 'Teacher ID resolved', name + ' -> ' + teacherId, '');
        return teacherId;
      }
    }

    debugLog('getTeacherIdByDisplayName', 'WARNING', 'No match found for display name', name, '');
    return '';

  } catch (error) {
    debugLog('getTeacherIdByDisplayName', 'ERROR', 'Failed', displayName, error.message);
    return '';
  }
}

function getTeacherInfoByDisplayName(displayName) {
  try {
    var lookupSheet = getSheet('teacherRosterLookup');

    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      debugLog('getTeacherInfoByDisplayName', 'WARNING', 'Teacher Roster Lookup sheet not found or empty', '', '');
      return null;
    }

    var getCol = createColumnFinder(lookupSheet);
    var firstNameCol   = getCol('First Name');
    var lastNameCol    = getCol('Last Name');
    var rosterUrlCol   = getCol('Roster URL');
    var teacherIdCol   = getCol('Teacher ID');
    var displayNameCol = getCol('Display Name');
    var statusCol      = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');

    if (!teacherIdCol || !displayNameCol) {
      debugLog('getTeacherInfoByDisplayName', 'ERROR', 'Required columns not found', '', '');
      return null;
    }

    var data = lookupSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][displayNameCol - 1]).trim() !== displayName) continue;

      return {
        firstName:   firstNameCol   ? String(data[i][firstNameCol - 1]).trim()  : '',
        lastName:    lastNameCol     ? String(data[i][lastNameCol - 1]).trim()   : '',
        rosterUrl:   rosterUrlCol    ? String(data[i][rosterUrlCol - 1]).trim()  : '',
        teacherId:   String(data[i][teacherIdCol - 1]).trim(),
        status:      statusCol       ? String(data[i][statusCol - 1]).trim()     : '',
        lastUpdated: lastUpdatedCol  ? data[i][lastUpdatedCol - 1]               : ''
      };
    }

    debugLog('getTeacherInfoByDisplayName', 'WARNING', 'Display name not found', displayName, '');
    return null;

  } catch (error) {
    debugLog('getTeacherInfoByDisplayName', 'ERROR', 'Failed', displayName, error.message);
    return null;
  }
}

function getTeacherInfoByFullName(firstName, lastName) {
  try {
    var lookupSheet = getSheet('teacherRosterLookup');

    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      debugLog('getTeacherInfoByFullName', 'WARNING', 'Teacher Roster Lookup sheet not found or empty', '', '');
      return null;
    }

    var getCol = createColumnFinder(lookupSheet);
    var firstNameCol   = getCol('First Name');
    var lastNameCol    = getCol('Last Name');
    var rosterUrlCol   = getCol('Roster URL');
    var teacherIdCol   = getCol('Teacher ID');
    var statusCol      = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');

    if (!firstNameCol || !lastNameCol) {
      debugLog('getTeacherInfoByFullName', 'ERROR', 'Required columns not found', '', '');
      return null;
    }

    var data = lookupSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowFirstName = String(data[i][firstNameCol - 1]).trim();
      var rowLastName  = String(data[i][lastNameCol - 1]).trim();
      if (rowFirstName.toLowerCase() !== firstName.trim().toLowerCase()) continue;
      if (rowLastName.toLowerCase()  !== lastName.trim().toLowerCase())  continue;

      return {
        firstName:   rowFirstName,
        lastName:    rowLastName,
        rosterUrl:   rosterUrlCol   ? String(data[i][rosterUrlCol - 1]).trim() : '',
        teacherId:   teacherIdCol   ? String(data[i][teacherIdCol - 1]).trim() : '',
        status:      statusCol      ? String(data[i][statusCol - 1]).trim()    : '',
        lastUpdated: lastUpdatedCol ? data[i][lastUpdatedCol - 1]              : ''
      };
    }

    debugLog('getTeacherInfoByFullName', 'WARNING', 'Teacher not found', firstName + ' ' + lastName, '');
    return null;

  } catch (error) {
    debugLog('getTeacherInfoByFullName', 'ERROR', 'Failed', firstName + ' ' + lastName, error.message);
    return null;
  }
}

function getTeacherNameById(teacherId) {
  try {
    if (!teacherId || String(teacherId).trim() === '') {
      debugLog('getTeacherNameById', 'WARNING', 'No Teacher ID provided', '', '');
      return '';
    }

    var sheet = getSheet('teachersAndAdmin');
    var headerMap = getHeaderMap(sheet);
    var norm = normalizeHeader;

    var teacherIdCol = headerMap[norm('Teacher ID')];
    var firstNameCol = headerMap[norm('First Name')];
    var lastNameCol = headerMap[norm('Last Name')];

    if (!teacherIdCol || !firstNameCol || !lastNameCol) {
      debugLog('getTeacherNameById', 'ERROR', 'Required columns not found in Teachers and Admin', '', '');
      return '';
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][teacherIdCol - 1]).trim() === String(teacherId).trim()) {
        var firstName = String(data[i][firstNameCol - 1] || '').trim();
        var lastName = String(data[i][lastNameCol - 1] || '').trim();
        var fullName = (firstName + ' ' + lastName).trim();
        debugLog('getTeacherNameById', 'SUCCESS', 'Teacher name resolved', teacherId + ' -> ' + fullName, '');
        return fullName;
      }
    }

    debugLog('getTeacherNameById', 'WARNING', 'Teacher ID not found in Teachers and Admin', teacherId, '');
    return '';

  } catch (error) {
    debugLog('getTeacherNameById', 'ERROR', 'Failed', teacherId, error.message);
    return '';
  }
}

function getTemplate(templateKey) {
  var templateName = TEMPLATE_MAP[templateKey];
  if (!templateName) {
    throw new Error('❌ Template key "' + templateKey + '" not found in TEMPLATE_MAP');
  }
  
  var templateFolder = getTemplateFolder();
  var files = templateFolder.getFilesByName(templateName);
  
  if (!files.hasNext()) {
    throw new Error('❌ Template "' + templateName + '" not found in template folder');
  }
  
  var file = files.next();
  
  // Verify it's a Google Doc
  if (file.getMimeType() !== 'application/vnd.google-apps.document') {
    throw new Error('❌ Template "' + templateName + '" is not a Google Doc (found: ' + file.getMimeType() + ')');
  }
  
  return file;
}

function getTemplateFolder() {
  var env = EnvironmentManager.get();
  var folderId = CONFIG && CONFIG[env] && CONFIG[env].templateFolderId;
  if (!folderId) {
    throw new Error('❌ templateFolderId not found in config for environment "' + env + '"');
  }
  return DriveApp.getFolderById(folderId);
}

function getWeekdayName(startDate, weekdayName) {
  var targetDay = getWeekdayNumber(weekdayName);
  var date = new Date(startDate);
  var offset = (targetDay - date.getDay() + 7) % 7;
  date.setDate(date.getDate() + offset);
  return date;
}

function getWeekdayNumber(weekdayName) {
  var days = {
    Sunday: 0, Monday: 1, Tuesday: 2, Wednesday: 3,
    Thursday: 4, Friday: 5, Saturday: 6
  };
  return days[weekdayName.trim()];
}

function getWorkbook(workbookKey) {
  var env = EnvironmentManager.get();
  var workbookId = CONFIG && CONFIG[env] && CONFIG[env][workbookKey + 'Id'];
  if (!workbookId) {
    throw new Error('❌ Workbook ID not found for key "' + workbookKey + '" in environment "' + env + '"');
  }
  return SpreadsheetApp.openById(workbookId);
}

function getYearFromSemesterName(semesterName) {
  var match = semesterName.match(/\b(20\d{2})\b/);
  return match ? match[1] : '';
}

function inferWorkbookKey(sheetKey) {
  var keyToWorkbook = {
    teacherResponses: 'teacherInterest',
    teacherFieldMap: 'teacherInterest',
    teacherReturningResponses: 'teacherInterest',
    teacherReturningFieldMap: 'teacherInterest',
    teachersAndAdmin: 'contacts',
    students: 'contacts',
    parents: 'contacts',
    instrumentList: 'contacts',
    futureTeachers: 'contacts',
    billingMetadata: 'billing',
    semesterMetadata: 'billing',
    yearMetadata: 'billing',
    rates: 'billing',
    programList: 'billing',
    packages: 'billing',
    billingTemplate: 'billing',
    reregistration: 'billing',
    reregistrationQueue: 'billing',
    calendar: 'formResponses',
    fieldMap: 'formResponses',
    teacherRosterLookup: 'formResponses',
    ledgerTemplate: 'payments',
    teacherPreferences: 'scheduling'
  };
  return keyToWorkbook[sheetKey];
}

function insertCountFormula(sheet, row, startCol, endCol, targetCol) {
  startCol = startCol || 6;
  if (!endCol) {
    endCol = sheet.getLastColumn();
  }

  var startLetter = columnToLetter(startCol);
  var endLetter = columnToLetter(endCol);
  var formula = '=COUNTIF(' + startLetter + row + ':' + endLetter + row + ', "lesson") + COUNTIF(' + startLetter + row + ':' + endLetter + row + ', "no show")';

  // If targetCol is specified, use it
  if (targetCol) {
    sheet.getRange(row, targetCol).setFormula(formula);
    return;
  }
}

function interpretAgeField(ageResponse) {
  if (!ageResponse) return '';
  var firstChar = ageResponse.toString().trim().toUpperCase().charAt(0);
  if (firstChar === 'Y') return 'Adult';
  if (firstChar === 'N') return 'Child';
  return ageResponse.toString().trim();
}

function isCurrentOrFutureMonth(sheetName, targetDate) {
  var monthMap = {};
  for (var i = 0; i < MONTH_NAMES.length; i++) {
    monthMap[MONTH_NAMES[i].toLowerCase()] = i;
  }

  var sheetNameLower = sheetName.toLowerCase().trim();
  var sheetMonth = monthMap[sheetNameLower];

  if (sheetMonth === undefined) {
    return false;
  }

  var targetMonth = targetDate.getMonth();
  return sheetMonth >= targetMonth;
}

function isHistoricalDataInputEnabled() {
  return HISTORICAL_DATA_MODE;
}

function isIdAlreadyUsed(sheet, columnName, idToCheck) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = -1;
  
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === normalizeHeader(columnName)) {
      idCol = i;
      break;
    }
  }
  
  if (idCol === -1) return false;
  
  for (var j = 1; j < data.length; j++) {
    if (data[j][idCol] === idToCheck) {
      return true;
    }
  }
  
  return false;
}

function isMonthSheet(sheetName) {
  if (!sheetName) return false;
  var cacheKey = 'isMonth_' + sheetName;
  var cached = getCached(cacheKey);
  if (cached !== undefined) return cached;
  var result = MONTH_NAMES.indexOf(sheetName) !== -1 ||
               MONTH_NAMES.map(function(m) { return m.toLowerCase(); }).indexOf(sheetName.toLowerCase().trim()) !== -1;
  return setCached(cacheKey, result);
}

function logAllSheetHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var lastCol = sheet.getLastColumn();
    
    if (lastCol === 0) {
      debugLog('logAllSheetHeaders', 'INFO', 'Sheet is empty', sheet.getName(), '');
      continue;
    }
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    debugLog('logAllSheetHeaders', 'INFO', sheet.getName(), headers.join(' | '), '');
  }
}

function normalizeHeader(header) {
  if (!header) return '';
  var str = header.toString();
  return str.replace(/["\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function ocrPdfPage(pageBlob) {
  var resource = {
    name: pageBlob.getName(),
    mimeType: MimeType.GOOGLE_DOCS
  };
  var ocrOptions = {
    ocr: true,
    ocrLanguage: 'en'
  };

  var docFile = Drive.Files.create(resource, pageBlob, ocrOptions);
  var doc = DocumentApp.openById(docFile.id);
  var text = doc.getBody().getText();

  DriveApp.getFileById(docFile.id).setTrashed(true);

  return text;
}

function parseAllPackageQuantities(qty30Package, qty45Package, qty60Package) {
  try {
    return {
      qty30: extractLessonQuantityFromPackage(qty30Package),
      qty45: extractLessonQuantityFromPackage(qty45Package),
      qty60: extractLessonQuantityFromPackage(qty60Package),
      totalQuantity: extractLessonQuantityFromPackage(qty30Package) + 
                    extractLessonQuantityFromPackage(qty45Package) + 
                    extractLessonQuantityFromPackage(qty60Package)
    };
  } catch (error) {
    return {
      qty30: 0,
      qty45: 0,
      qty60: 0,
      totalQuantity: 0
    };
  }
}

function parseAndFormatAddress(rawAddress) {
  if (!rawAddress || typeof rawAddress !== 'string') return '';

  var lines = rawAddress.split(/\r?\n/);
  var filteredLines = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (line) {
      filteredLines.push(line);
    }
  }
  
  var street = '', city = '', zip = '';

  if (filteredLines.length === 0) return '';

  if (filteredLines.length === 1) {
    var match = filteredLines[0].match(/^(.+?)\s*,?\s*([A-Za-z\s]+)?\s*NY\s+(\d{5})$/);
    if (match) {
      street = match[1].trim();
      city = match[2] ? match[2].trim() : '';
      zip = match[3];
    } else {
      street = filteredLines[0]; // fallback
    }
  } else {
    street = filteredLines.slice(0, -1).join(', ');
    var lastLineMatch = filteredLines[filteredLines.length - 1].match(/^([A-Za-z\s]+),?\s*NY\s+(\d{5})$/);
    if (lastLineMatch) {
      city = lastLineMatch[1].trim();
      zip = lastLineMatch[2];
    }
  }

  return street + '\n' + city + ', NY ' + zip;
}

function parseCityZipMessy(input) {
  if (!input) return { city: '', zip: '' };

  // Convert to string first to handle numbers or other types
  var cleaned = String(input).trim().replace(/\s+/g, ' ').replace(/[.]/g, ',');
  
  // Extract zip (last 5-digit number)
  var zipMatch = cleaned.match(/(\d{5})(?!.*\d)/);
  var zip = zipMatch ? zipMatch[1] : '';

  // Remove zip and everything after it
  var cityPart = zip ? cleaned.substring(0, cleaned.lastIndexOf(zip)).trim() : cleaned;

  // Remove trailing commas or "NY"/"New York"
  cityPart = cityPart.replace(/,\s*(NY|New York)?$/i, '').trim();
  cityPart = cityPart.replace(/(NY|New York)$/i, '').trim();
  cityPart = cityPart.replace(/,+$/, '').trim();

  // Capitalize city nicely
  var words = cityPart.toLowerCase().split(' ');
  var capitalizedWords = [];
  for (var i = 0; i < words.length; i++) {
    if (words[i]) {
      capitalizedWords.push(words[i].charAt(0).toUpperCase() + words[i].slice(1));
    }
  }
  var city = capitalizedWords.join(' ');

  return { city: city, zip: zip };
}

function parseDateFromString(str) {
  if (!str || typeof str !== 'string') return null;
  var parts = str.trim().split('/');
  if (parts.length !== 3) return null;
  var month = parseInt(parts[0], 10);
  var day = parseInt(parts[1], 10);
  var year = parseInt(parts[2], 10);
  var date = new Date(year, month - 1, day);
  return isNaN(date.getTime()) ? null : date;
}

function parseGridInstruments(headers, rowValues) {
  var GRID_PREFIX = 'What instrument(s) are you interested in teaching? Please check the level(s) you are comfortable teaching for those instruments you wish to teach.';
  var BRACKET_REGEX = /\[(.+?)\]\s*$/;
  var LEVEL_MAP = { 'beginning': 'beg', 'intermediate': 'int', 'advanced': 'adv' };
  var results = [];

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i].toString();
    if (header.toLowerCase().indexOf(GRID_PREFIX.toLowerCase()) !== 0) continue;

    var bracketMatch = header.match(BRACKET_REGEX);
    if (!bracketMatch) continue;

    var cellValue = (rowValues[i] || '').toString().trim();
    if (!cellValue) continue;

    var levels = cellValue.split(',').map(function(l) {
      var t = l.trim().toLowerCase();
      return LEVEL_MAP[t] || t;
    }).filter(Boolean).join('/');

    results.push({ instrument: bracketMatch[1].trim(), levels: levels });
  }

  return results;
}

function parseRosterData(row, headerMap, fieldMap, studentIdOverride) {
  studentIdOverride = studentIdOverride || '';
  try {
    var get = function(internalField) {
      var formHeader = null;
      for (var key in fieldMap) {
        if (fieldMap[key] === internalField) {
          formHeader = key;
          break;
        }
      }
      if (!formHeader) return '';
      var normHeader = normalizeHeader(formHeader);
      var colIndex = headerMap[normHeader] - 1;
      return row[colIndex] !== undefined ? row[colIndex] : '';
    };

    var email = get("Email");
    var firstName = get("Student First Name");
    var lastName = get("Student Last Name");
    var instrument = get("Instrument");
    var experience = get("Experience");
    var lengthRaw = get("Length");
    
    // Use utility function to extract quantities from package text
    var qty30Package = get("Qty30");
    var qty45Package = get("Qty45");
    var qty60Package = get("Qty60");
    var qty30 = extractLessonQuantityFromPackage(qty30Package);
    var qty45 = extractLessonQuantityFromPackage(qty45Package);
    var qty60 = extractLessonQuantityFromPackage(qty60Package);
    
    var grade = get("Grade");
    var parentFirstName = get("Parent First Name");
    var parentLastName = get("Parent Last Name");
    var additionalInfo = get("Additional contacts");
    var phoneRaw = get("Phone");
    var studentId = studentIdOverride || get("Student ID");

    // FIXED: Complete the length determination logic
    var quantity = qty60 || qty45 || qty30 || 1;
    var length = lengthRaw ? lengthRaw.toString() : (qty60 > 0 ? '60' : qty45 > 0 ? '45' : qty30 > 0 ? '30' : '30');
    
    var phone = formatPhoneNumber(phoneRaw);
    var address = formatAddress(get("Address"), get("City"), get("Zip"));
    
    // Return parsed data array where index [3] is the length
    return [
      quantity,           // [0] - Total quantity of lessons
      address,            // [1] - Formatted address
      phone,              // [2] - Formatted phone number
      length,             // [3] - Lesson length (30, 45, or 60)
      email,              // [4] - Email
      firstName,          // [5] - Student first name
      lastName,           // [6] - Student last name
      instrument,         // [7] - Instrument
      experience,         // [8] - Experience level
      grade,              // [9] - Grade
      parentFirstName,    // [10] - Parent first name
      parentLastName,     // [11] - Parent last name
      additionalInfo,     // [12] - Additional contacts
      studentId          // [13] - Student ID
    ];
    
  } catch (error) {
    debugLog('parseRosterData', 'ERROR', 'Error parsing roster data', '', error.message);
    return ['', '', '', '30', '', '', '', '', '', '', '', '', '', '']; // Default with '30' as length
  }
}

function prefillAttendanceDatesForStudent(rowValues, headers) {
  var get = function(label) {
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].toLowerCase() === label.toLowerCase()) {
        return rowValues[i];
      }
    }
    return '';
  };

  var studentId = get("Student ID");
  var firstLessonDate = new Date(get("First Lesson Date"));
  var lessonDay = get("Lesson Day");

  if (!studentId || isNaN(firstLessonDate) || !lessonDay) return;

  var env = EnvironmentManager.get();
  var config = CONFIG[env];
  var folder = DriveApp.getFolderById(config.rosterFolderId);
  var fileName = get("Teacher") + " Roster";
  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return;

  var rosterFile = SpreadsheetApp.openById(files.next().getId());
  var attendanceSheet = rosterFile.getSheetByName("Attendance");
  if (!attendanceSheet) return;

  var calendarSheet = SpreadsheetApp.openById(config.formResponsesId).getSheetByName("Calendar");
  var calendarData = calendarSheet.getDataRange().getValues();
  var rosterSheetName = rosterFile.getSheetByName("Roster") ? rosterFile.getSheetByName("Roster").getName() : '';
  var filteredCalendarData = [];
  for (var i = 0; i < calendarData.length; i++) {
    if (calendarData[i][3] === rosterSheetName) {
      filteredCalendarData.push(calendarData[i]);
    }
  }

  var headerMap = getHeaderMap(attendanceSheet);
  var data = attendanceSheet.getDataRange().getValues();
  var studentRowIndex = -1;
  
  for (var j = 0; j < data.length; j++) {
    if (j === 0) continue;
    var cellValue = data[j][headerMap["student id"] - 1];
    if (cellValue && cellValue.toString().trim() === studentId) {
      studentRowIndex = j;
      break;
    }
  }

  if (studentRowIndex === -1) return;

  for (var k = 0; k < filteredCalendarData.length; k++) {
    var row = filteredCalendarData[k];
    var weekStart = new Date(row[1]);
    var lessonDate = getDateForWeekday(weekStart, lessonDay);
    if (lessonDate >= firstLessonDate) {
      var col = 9 + k * 2; // baseHeaders = 9
      attendanceSheet.getRange(studentRowIndex + 1, col).setValue(lessonDate);
    }
  }
}

function promptForCustomToday() {
  if (!isHistoricalDataInputEnabled()) {
    return new Date();
  }

  return promptForDate({
    title: 'Historical Data Entry',
    message: 'Enter the date to use as "today" (MM/DD/YYYY).',
    defaultDate: new Date(),
    cancelMessage: 'Billing cycle setup cancelled.'
  });
}

function promptForDate(config) {
  // config: {
  //   title: string,
  //   message: string,
  //   defaultDate: Date|null,  // if provided, blank input returns this date
  //   cancelMessage: string    // optional
  // }
  // Returns: Date or null (cancelled)

  var ui = SpreadsheetApp.getUi();
  var cancelMessage = config.cancelMessage || 'Cancelled.';

  var message = config.message;
  if (config.defaultDate) {
    var defaultStr = formatDateFlexible(config.defaultDate, "MM/dd/yyyy");
    message += '\nLeave blank for ' + defaultStr + '.';
  }

  var response = ui.prompt(config.title, message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('❌ ' + cancelMessage);
    return null;
  }

  var dateString = response.getResponseText().trim();

  if (dateString === '') {
    if (config.defaultDate) return config.defaultDate;
    ui.alert('❌ A date is required.');
    return null;
  }

  var parsed = parseDateFromString(dateString);
  if (!parsed) {
    ui.alert('❌ Invalid date format. Please use MM/DD/YYYY.');
    return null;
  }

  return parsed;
}

function promptForHistoricalId(sheet, columnName, prefix, recordName) {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Get the auto-generated ID as a suggestion - call direct generation to avoid recursion
    var suggestedId = generateNextIdDirect(sheet, columnName, prefix);
    
    // Build prompt message with record name if provided
    var recordInfo = recordName && String(recordName).trim() !== '' ? 
      ' for ' + recordName : 
      ' for this record';
    
    var prompt = ui.prompt(
      'Historical Data Entry', 
      'Enter the ' + columnName + recordInfo + ':\n\n' +
      '(Suggested next auto-ID: ' + suggestedId + ')\n\n' +
      'Enter ' + prefix + 'XXXX format (e.g., ' + prefix + '0001):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (prompt.getSelectedButton() === ui.Button.OK) {
      var enteredId = prompt.getResponseText().trim().toUpperCase();
      
      // Validate format
      var regex = new RegExp("^" + prefix + "\\d{4}$");
      if (!regex.test(enteredId)) {
        ui.alert('Invalid Format', 'ID must be in format ' + prefix + 'XXXX (e.g., ' + prefix + '0001)', ui.ButtonSet.OK);
        return promptForHistoricalId(sheet, columnName, prefix, recordName); // Try again
      }
      
      // Check for duplicates
      if (isIdAlreadyUsed(sheet, columnName, enteredId)) {
        var overwrite = ui.alert('Duplicate ID', 'ID "' + enteredId + '" already exists. Use it anyway?', ui.ButtonSet.YES_NO);
        if (overwrite === ui.Button.NO) {
          return promptForHistoricalId(sheet, columnName, prefix, recordName); // Try again
        }
      }
      
      return enteredId;
    } else {
      // User cancelled - use auto-generated ID
      return suggestedId;
    }
    
  } catch (error) {
    debugLog('promptForHistoricalId', 'ERROR', 'Failed, falling back to auto-generate', '', error.message);
    // Fallback to auto-generation
    return generateNextIdDirect(sheet, columnName, prefix);
  }
}

function promptForNameWithDefault(config) {
  var ui = SpreadsheetApp.getUi();
  
  // Show default confirmation if provided
  if (config.defaultValue) {
    var confirm = ui.alert(
      'Is "' + config.defaultValue + '" the ' + config.entityType + '?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirm === ui.Button.YES) {
      return config.defaultValue;
    }
  }
  
  // Prompt for custom value
  var response = ui.prompt(
    config.promptTitle,
    config.promptMessage,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert(config.entityType + ' setup cancelled.');
    return null;
  }
  
  var customName = response.getResponseText().trim();
  if (!customName) {
    ui.alert(config.entityType + ' name cannot be empty.');
    return null;
  }
  
  // Optional validation
  if (config.minLength && customName.length < config.minLength) {
    ui.alert(config.entityType + ' name is too short. Please provide a meaningful name.');
    return null;
  }
  
  return customName;
}

function protectAttendanceSheet(sheet) {
  try {
    var maxRows = sheet.getMaxRows();

    protectSheetRanges(sheet, {
      warningOnly: false,
      ranges: [
        { range: sheet.getRange(1, 1, 1, 12),      description: 'Header row - do not edit' },
        { range: sheet.getRange(2, 1),              description: 'Sign-off label - do not edit' },
        { range: sheet.getRange(2, 4, 1, 8),        description: 'Sign-off row - do not edit' },
        { range: sheet.getRange(1, 1, maxRows, 1),  description: 'Student ID column - do not edit' },
        { range: sheet.getRange(1, 2, maxRows, 1),  description: 'Student Name column - do not edit' }
      ]
    });

    debugLog('protectAttendanceSheet', 'INFO', 'Sheet protected', sheet.getName(), '');

  } catch (error) {
    debugLog('protectAttendanceSheet', 'ERROR', 'Failed to protect sheet', sheet.getName(), error.message);
    throw error;
  }
}

function protectSheetRanges(sheet, options) {
  options = options || {};
  var columns = options.columns || [];
  var ranges = options.ranges || [];
  var warningOnly = options.warningOnly !== false; // Default true
  var clearExisting = options.clearExisting || false;

  try {
    if (clearExisting) {
      var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < existingProtections.length; i++) {
        existingProtections[i].remove();
      }
      debugLog('protectSheetRanges', 'INFO', 'Cleared existing protections', 'Count: ' + existingProtections.length, '');
    }

    // Column string ranges (e.g. "A:A", "G:K")
    for (var j = 0; j < columns.length; j++) {
      var protection = sheet.getRange(columns[j]).protect();
      protection.setDescription('Protected columns: ' + columns[j]);
      protection.setWarningOnly(warningOnly);
      if (!warningOnly) {
        var editors = protection.getEditors();
        if (editors.length > 0) protection.removeEditors(editors);
        if (protection.canDomainEdit()) protection.setDomainEdit(false);
      }
    }

    // Explicit Range objects (e.g. sheet.getRange(1,1,1,11))
    for (var k = 0; k < ranges.length; k++) {
      var rProtection = ranges[k].range.protect();
      rProtection.setDescription(ranges[k].description || 'Protected range');
      rProtection.setWarningOnly(warningOnly);
      if (!warningOnly) {
        var rEditors = rProtection.getEditors();
        if (rEditors.length > 0) rProtection.removeEditors(rEditors);
        if (rProtection.canDomainEdit()) rProtection.setDomainEdit(false);
      }
    }

    debugLog('protectSheetRanges', 'INFO', 'Protections applied',
             'Columns: ' + columns.length + ', Ranges: ' + ranges.length +
             (warningOnly ? ' (warning only)' : ' (hard)'), '');

    return { success: true, protectedRanges: columns.length + ranges.length };

  } catch (error) {
    debugLog('protectSheetRanges', 'ERROR', 'Error protecting sheet ranges', '', error.message);
    return { success: false, error: error.message };
  }
}

function protectStudentHeaderRow(sheet, row) {
  try {
    protectSheetRanges(sheet, {
      warningOnly: false,
      ranges: [
        { range: sheet.getRange(row, 1, 1, 12), description: 'Student header row - do not edit' }
      ]
    });
    debugLog('protectStudentHeaderRow', 'INFO', 'Protected student header row', 'Row: ' + row, '');
  } catch (error) {
    debugLog('protectStudentHeaderRow', 'ERROR', 'Failed to protect student header row', 'Row: ' + row, error.message);
    throw error;
  }
}

function safeGet(row, index) {
  return Array.isArray(row) && index >= 0 && index < row.length ? row[index] : '';
}

function safeParseFloat(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    return parseFloat(value.replace(/[^0-9.\-]/g, "")) || 0;
  }
  return 0;
}

function sendEmail(to, subject, body, options) {
  try {
    options = options || {};
    var params = {
      name: 'Quaker Arts Music Program'
    };
    if (options.cc)       params.cc       = options.cc;
    if (options.bcc)      params.bcc      = options.bcc;
    if (options.htmlBody) params.htmlBody = options.htmlBody;

    GmailApp.sendEmail(to, subject, body, params);
    debugLog('sendEmail', 'SUCCESS', 'Email sent', to, '');
  } catch (error) {
    debugLog('sendEmail', 'ERROR', 'Failed to send email', to, error.message);
    throw error;
  }
}

function setCached(key, value) {
  _executionCache[key] = value;
  return value;
}

function setupAttendanceHeaders(sheet) {
  var headers = [
    'Student ID',           // A
    'Student Name',         // B
    'Date',                // C
    'Length',              // D
    'Status',              // E
    'Comments',            // F
    'Admin Review Date',   // G
    'Invoice Date',        // H
    'Payment Date',        // I
    'Invoice Number',      // J
    'Admin Comments',      // K
    'Calendar Event ID'    // L
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(STYLES.HEADER.background)
             .setFontColor(STYLES.HEADER.text)
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setWrap(true);
  
  var widths = [60, 220, 80, 80, 95, 220, 75, 110, 75, 100, 220];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  headerRange.getCell(1, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  createSignOffRow(sheet);
  
  debugLog('setupAttendanceHeaders', 'INFO', 'Headers and sign-off row created', '', '');
}

function setupRosterTemplateProtection(sheet) {
  try {
    protectSheetRanges(sheet, {
      columns: ['E:U'],
      warningOnly: true,
      clearExisting: true
    });

    var headerMap = getHeaderMap(sheet);

    var firstLessonDateCol = headerMap[normalizeHeader('First Lesson Date')];
    if (!firstLessonDateCol) {
      debugLog('setupRosterTemplateProtection', 'WARNING', 'First Lesson Date column not found, skipping date validation', '', '');
    } else {
      var dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, firstLessonDateCol, sheet.getMaxRows() - 1, 1).setDataValidation(dateRule);
    }

    var statusCol = headerMap[normalizeHeader('Status')];
    if (!statusCol) {
      debugLog('setupRosterTemplateProtection', 'WARNING', 'Status column not found, skipping status validation', '', '');
    } else {
      var statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['active', 'dropped', 'carryover', 'transferred'], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(statusRule);
    }

    debugLog('setupRosterTemplateProtection', 'SUCCESS', 'Roster protection, date validation, and status dropdown applied', '', '');

  } catch (error) {
    debugLog('setupRosterTemplateProtection', 'ERROR', 'Error in roster protection', '', error.message);
  }
}

function setupStatusValidation(sheet, lastRow) {
  try {
    // Get all Student ID values to determine which are groups vs students
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var studentIdIdx = -1;
    var dateIdx = -1;
    var lengthIdx = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i]).toLowerCase().trim();
      if (header === 'student id' || header === 'id') studentIdIdx = i;
      if (header === 'date') dateIdx = i;
      if (header === 'length') lengthIdx = i;
    }
    
    if (studentIdIdx === -1) {
      debugLog('setupStatusValidation', 'ERROR', 'Student ID column not found', '', '');
      return;
    }
    
    // Student status options: Lesson, No Show, No Lesson
    var studentStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Lesson', 'No Show', 'No Lesson'])
      .setAllowInvalid(false)
      .setHelpText('Select lesson status')
      .build();
    
    // Group status options: Lesson, No Lesson (no "No Show" for groups)
    var groupStatusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Lesson', 'No Lesson'])
      .setAllowInvalid(false)
      .setHelpText('Select group session status')
      .build();
    
    var processedRows = 0;
    
    // Apply appropriate validation to each row (skip header row)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentId = row[studentIdIdx];
      var rowNum = i + 1; // Convert to 1-based row number
      var statusCell = sheet.getRange(rowNum, 5); // Status column is E (5)
      
      // Skip if no student ID
      if (!studentId || studentId.toString().trim() === '') {
        continue;
      }
      
      // Check if this is a header row
      var isHeaderRow = false;
      
      if (dateIdx !== -1 && lengthIdx !== -1) {
        var dateValue = row[dateIdx];
        var lengthValue = row[lengthIdx];
        
        // Check if date contains a month name
        // Check if date contains a month name
        if (dateValue && typeof dateValue === 'string') {
          var allMonthNames = MONTH_NAMES;
          var dateLower = dateValue.toLowerCase();
          for (var m = 0; m < allMonthNames.length; m++) {
            if (dateLower.indexOf(allMonthNames[m].toLowerCase()) !== -1) {
              isHeaderRow = true;
              break;
            }
          }
        }
                
        // Check if length contains " minutes" suffix
        if (!isHeaderRow && lengthValue && typeof lengthValue === 'string' && lengthValue.indexOf(' minutes') !== -1) {
          isHeaderRow = true;
        }
      }
      
      // Skip header rows - they should not have dropdowns
      if (isHeaderRow) {
        continue;
      }
      
      // Apply validation based on ID prefix
      var studentIdStr = studentId.toString();
      if (studentIdStr.match(/^G\d{4}$/)) {
        // This is a group entry - use group status validation
        statusCell.setDataValidation(groupStatusRule);
        processedRows++;
      } else if (studentIdStr.match(/^Q\d{4}$/)) {
        // This is a student entry - use student status validation
        statusCell.setDataValidation(studentStatusRule);
        processedRows++;
      }
    }
    
    debugLog('setupStatusValidation', 'INFO', 'Applied status validation', 
                                 'Rows processed: ' + processedRows, '');
    
  } catch (error) {
    debugLog('setupStatusValidation', 'ERROR', 'Error setting up status validation', '', error.message);
  }
}

function shouldBeCurrency(columnName) {
  if (!columnName) return false;
  var normalized = normalizeHeader(columnName);
  
  // NEVER format hours/time columns as currency
  if (normalized.indexOf("hours") !== -1 || 
      normalized.indexOf("remaining") !== -1 ||
      normalized.indexOf("taught") !== -1) {
    return false;
  }
  
  return normalized.indexOf("price") !== -1 ||
         normalized.indexOf("total") !== -1 ||
         normalized.indexOf("balance") !== -1 ||
         normalized.indexOf("fee") !== -1;
}

function showConfirmationDialog(title, message, details, options) {
  options = options || {};
  var buttonSet = options.buttonSet || SpreadsheetApp.getUi().ButtonSet.YES_NO;
  var confirmButton = options.confirmButton || SpreadsheetApp.getUi().Button.YES;
  
  try {
    var formattedDetails = '';
    
    // Format details
    if (details) {
      if (typeof details === 'string') {
        formattedDetails = details;
      } else if (Array.isArray(details)) {
        formattedDetails = details.join('\n');
      } else {
        formattedDetails = details.toString();
      }
    }
    
    // Build full message
    var fullMessage = message;
    if (formattedDetails) {
      fullMessage += '\n\n' + formattedDetails;
    }
    
    // Show dialog
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(title, fullMessage, buttonSet);
    
    return response === confirmButton;
    
  } catch (error) {
    debugLog('showConfirmationDialog', 'ERROR', 'Dialog failed', '', error.message);
    return false;
  }
}

async function splitPdfIntoSinglePages(pdfBlob) {
  var sourceBytes = new Uint8Array(pdfBlob.getBytes());
  var srcDoc = await PDFLib.PDFDocument.load(sourceBytes);
  var pageCount = srcDoc.getPageCount();
  var pageBlobs = [];

  for (var i = 0; i < pageCount; i++) {
    var newDoc = await PDFLib.PDFDocument.create();
    var copiedPages = await newDoc.copyPages(srcDoc, [i]);
    newDoc.addPage(copiedPages[0]);

    var pdfBytes = await newDoc.save();
    var pageBlob = Utilities.newBlob(pdfBytes, 'application/pdf', 'page' + (i + 1) + '.pdf');
    pageBlobs.push(pageBlob);
  }

  return pageBlobs;
}

function styleHeaderRow(sheet, headers) {
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setFontWeight("bold")
    .setFontColor(STYLES.HEADER.text)
    .setBackground(STYLES.HEADER.background)
    .setHorizontalAlignment("center")
    .setWrap(true);

  var dataRange = sheet.getRange(2, 1, sheet.getMaxRows() - 1, headers.length);
  dataRange.setFontWeight("normal").setBackground(null);
}

function truncateString(str, maxLength) {
  if (str && str.length > maxLength) {
    return str.substring(0, maxLength - 3) + '...';
  }
  return str;
}

function updateFieldMappings(fieldMapSheet, newHeaders, sourceSheetName, options) {
  options = options || {};
  var highlightDuplicates = options.highlightDuplicates !== false; // Default true
  var addMissingHeaders = options.addMissingHeaders !== false; // Default true
  
  try {
    // Get existing mappings
    var existingData = fieldMapSheet.getDataRange().getValues();
    var existingHeaders = existingData[0];
    
    // Find required columns
    var formHeaderCol = -1;
    var internalFieldCol = -1;
    var activeCol = -1;
    var notesCol = -1;
    
    for (var i = 0; i < existingHeaders.length; i++) {
      var header = normalizeHeader(existingHeaders[i]);
      if (header.indexOf('form') !== -1 && header.indexOf('header') !== -1) {
        formHeaderCol = i;
      } else if (header.indexOf('internal') !== -1 || header.indexOf('field') !== -1) {
        internalFieldCol = i;
      } else if (header.indexOf('active') !== -1 || header.indexOf('updated') !== -1) {
        activeCol = i;
      } else if (header.indexOf('notes') !== -1) {
        notesCol = i;
      }
    }
    
    if (formHeaderCol === -1) {
      throw new Error('Could not find form header column in field map sheet');
    }
    
    // Build normalized existing headers map
    var existingNormalized = {};
    var duplicates = [];
    
    for (var j = 1; j < existingData.length; j++) {
      var formHeader = existingData[j][formHeaderCol];
      if (!formHeader) continue;
      
      var normalized = normalizeHeader(formHeader);
      
      if (existingNormalized[normalized]) {
        duplicates.push({
          row: j + 1,
          header: formHeader,
          normalized: normalized
        });
      } else {
        existingNormalized[normalized] = j + 1;
      }
    }
    
    // Find new headers to add
    var newRows = [];
    var newHeaderCount = 0;
    
    if (addMissingHeaders) {
      for (var k = 0; k < newHeaders.length; k++) {
        var header = newHeaders[k];
        if (!header || typeof header !== 'string') continue;
        
        var normalized = normalizeHeader(header);
        
        if (!existingNormalized[normalized]) {
          var newRow = [];
          
          // Build new row based on available columns
          for (var col = 0; col < existingHeaders.length; col++) {
            if (col === formHeaderCol) {
              newRow.push(header);
            } else if (col === internalFieldCol) {
              newRow.push(''); // Leave blank for manual mapping
            } else if (col === activeCol) {
              newRow.push(true); // Mark as active by default
            } else if (col === notesCol) {
              newRow.push('Added from ' + sourceSheetName);
            } else {
              newRow.push(''); // Default empty value
            }
          }
          
          newRows.push(newRow);
          existingNormalized[normalized] = true; // Prevent duplicates in this batch
          newHeaderCount++;
        }
      }
    }
    
    // Add new rows if any
    if (newRows.length > 0) {
      var startRow = fieldMapSheet.getLastRow() + 1;
      var range = fieldMapSheet.getRange(startRow, 1, newRows.length, existingHeaders.length);
      range.setValues(newRows);
      
      // Apply default formatting to new rows
      range.setFontWeight('normal');
      range.setBackground(null);
      range.setFontColor('black');
    }
    
    // Highlight duplicates if requested
    if (highlightDuplicates && duplicates.length > 0) {
      for (var d = 0; d < duplicates.length; d++) {
        var duplicate = duplicates[d];
        var duplicateRange = fieldMapSheet.getRange(duplicate.row, 1, 1, existingHeaders.length);
        duplicateRange.setBackground(STYLES.WARNING.background);
        duplicateRange.setFontColor(STYLES.WARNING.text);
        
        if (notesCol !== -1) {
          fieldMapSheet.getRange(duplicate.row, notesCol + 1).setValue('Duplicate header');
        }
      }
    }
    
    return {
      success: true,
      newHeaders: newHeaderCount,
      duplicates: duplicates.length,
      details: {
        added: newRows.length,
        duplicatesFound: duplicates
      }
    };
    
  } catch (error) {
    debugLog('updateFieldMappings', 'ERROR', 'Failed to update field mappings', '', error.message);
    return {
      success: false,
      error: error.message,
      newHeaders: 0,
      duplicates: 0
    };
  }
}

function updateParentContactFields(parentsSheet, parentRow, fieldsToUpdate, options) {
  options = options || {};
  var updateCheckbox = options.updateUpdatedCheckbox !== false; // default true

  try {
    var norm = normalizeHeader;
    var headers = parentsSheet.getRange(1, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
    var rowValues = parentsSheet.getRange(parentRow, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
    var changesMade = false;

    // Find a column index by field name
    var findCol = function(fieldName) {
      for (var i = 0; i < headers.length; i++) {
        if (norm(String(headers[i])) === norm(fieldName)) return i + 1;
      }
      return 0;
    };

    // Update text fields — only write if value differs
    for (var field in fieldsToUpdate) {
      if (!fieldsToUpdate.hasOwnProperty(field)) continue;
      var col = findCol(field);
      if (!col) continue;
      var newValue = fieldsToUpdate[field] !== null && fieldsToUpdate[field] !== undefined
        ? String(fieldsToUpdate[field]) : '';
      var currentValue = String(rowValues[col - 1] || '');
      if (newValue !== currentValue) {
        parentsSheet.getRange(parentRow, col).setValue(newValue);
        changesMade = true;
        debugLog('updateParentContactFields', 'INFO', 'Updated ' + field,
          "'" + currentValue + "' → '" + newValue + "'", '');
      }
    }

    // Update Parent Lookup key if provided
    if (options.newLookupKey) {
      var lookupCol = findCol('Parent Lookup');
      if (lookupCol) {
        var currentLookup = String(rowValues[lookupCol - 1] || '').toLowerCase().trim();
        var newLookup = options.newLookupKey.toLowerCase().trim();
        if (currentLookup !== newLookup) {
          parentsSheet.getRange(parentRow, lookupCol).setValue(options.newLookupKey);
          changesMade = true;
          debugLog('updateParentContactFields', 'INFO', 'Updated Parent Lookup',
            "'" + currentLookup + "' → '" + newLookup + "'", '');
        }
      }
    }

    // Update Updated checkbox
    if (updateCheckbox) {
      var updatedCol = findCol('Updated');
      if (updatedCol) {
        var updatedCell = parentsSheet.getRange(parentRow, updatedCol);
        updatedCell.insertCheckboxes();
        updatedCell.setValue(changesMade);
      }
    }

    return { changesMade: changesMade };

  } catch (error) {
    debugLog('updateParentContactFields', 'ERROR', 'Failed', '', error.message);
    throw error;
  }
}

function validateProgramConfiguration(programSheet, options) {
  options = options || {};
  var checkPackages = options.checkPackages !== false; // Default true
  var checkRates = options.checkRates !== false; // Default true
  
  try {
    var data = programSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find required columns
    var nameCol = headers.indexOf('Program Name');
    var activeCol = headers.indexOf('Active');
    var typeCol = headers.indexOf('Type');
    var aliasCol = headers.indexOf('Alias For');
    var rateKeyCol = headers.indexOf('Rate Key');
    
    if (nameCol === -1 || activeCol === -1 || typeCol === -1) {
      throw new Error('Required columns not found in Programs sheet');
    }
    
    var issues = [];
    var activePrograms = [];
    
    // Collect active programs
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[activeCol] === true) {
        activePrograms.push(row[nameCol]);
      }
    }
    
    // Check package configurations
    if (checkPackages) {
      for (var j = 1; j < data.length; j++) {
        var row = data[j];
        var isActive = row[activeCol] === true;
        var type = row[typeCol];
        var programName = row[nameCol];
        var aliasFor = aliasCol !== -1 ? row[aliasCol] : '';
        
        if (isActive && type === 'Package') {
          if (!aliasFor || typeof aliasFor !== 'string') {
            issues.push({
              type: 'MISSING_ALIAS',
              program: programName,
              message: 'Package "' + programName + '" has no "Alias For" value'
            });
            continue;
          }
          
          var aliasArray = aliasFor.split(',');
          var cleanAliases = [];
          for (var k = 0; k < aliasArray.length; k++) {
            var alias = aliasArray[k].trim();
            if (alias) {
              cleanAliases.push(alias);
            }
          }
          
          var missingPrograms = [];
          for (var l = 0; l < cleanAliases.length; l++) {
            var alias = cleanAliases[l];
            if (activePrograms.indexOf(alias) === -1) {
              missingPrograms.push(alias);
            }
          }
          
          if (missingPrograms.length > 0) {
            issues.push({
              type: 'MISSING_COMPONENTS',
              program: programName,
              missing: missingPrograms,
              message: 'Package "' + programName + '" references inactive programs: ' + missingPrograms.join(', ')
            });
          }
        }
      }
    }
    
    // Check rate configurations
    if (checkRates && rateKeyCol !== -1) {
      for (var m = 1; m < data.length; m++) {
        var row = data[m];
        var isActive = row[activeCol] === true;
        var type = row[typeCol];
        var programName = row[nameCol];
        var rateKey = row[rateKeyCol];
        
        if (isActive && type !== 'Package') {
          if (!rateKey || typeof rateKey !== 'string' || rateKey.trim() === '') {
            issues.push({
              type: 'MISSING_RATE_KEY',
              program: programName,
              message: 'Active program "' + programName + '" has no rate key'
            });
          }
        }
      }
    }
    
    return {
      success: true,
      isValid: issues.length === 0,
      issues: issues,
      activePrograms: activePrograms,
      summary: issues.length + ' issues found'
    };
    
  } catch (error) {
    debugLog('validateProgramConfiguration', 'ERROR', 'Validation failed', '', error.message);
    return {
      success: false,
      error: error.message,
      issues: [],
      activePrograms: []
    };
  }
}

function validateTemplateVariables(variableMap) {
  try {
    var validatedMap = {};
    
    for (var variable in variableMap) {
      var value = variableMap[variable];
      
      // Convert to string if not already
      if (typeof value !== 'string') {
        if (value === null || value === undefined) {
          value = '';
        } else {
          value = value.toString();
        }
      }
      
      // Clean up the value
      value = value.trim();
      
      validatedMap[variable] = value;
    }
    
    return {
      success: true,
      variables: validatedMap,
      errors: []
    };
    
  } catch (error) {
    return {
      success: false,
      variables: {},
      errors: ['Error validating template variables: ' + error.message]
    };
  }
}

function verifyConfigurationWithUser(title, itemName, data, formatFunction, options) {
  options = options || {};
  var allowSkip = options.allowSkip || false;
  var skipButtonText = options.skipButtonText || 'Skip';
  var confirmButtonText = options.confirmButtonText || 'Confirm';
  var customMessage = options.message;
  
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Format the data for display
    var formattedData = '';
    if (formatFunction && typeof formatFunction === 'function') {
      try {
        formattedData = formatFunction(data);
      } catch (formatError) {
        formattedData = 'Error formatting data: ' + formatError.message;
        debugLog('verifyConfigurationWithUser', 'WARNING', 'Format function error', itemName, formatError.message);
      }
    } else if (data) {
      // Default formatting
      if (typeof data === 'string') {
        formattedData = data;
      } else if (Array.isArray(data)) {
        formattedData = data.join('\n');
      } else if (typeof data === 'object') {
        var lines = [];
        for (var key in data) {
          if (data.hasOwnProperty(key)) {
            lines.push(key + ': ' + data[key]);
          }
        }
        formattedData = lines.join('\n');
      } else {
        formattedData = data.toString();
      }
    }
    
    // Build message
    var message = customMessage || ('Please review the ' + itemName + ' configuration:');
    
    // Determine button set
    var buttonSet;
    var confirmButton;
    
    if (allowSkip) {
      // Create three-button dialog using YES_NO_CANCEL
      buttonSet = ui.ButtonSet.YES_NO_CANCEL;
      confirmButton = ui.Button.YES;
      
      // Modify message to explain buttons
      message += '\n\nClick "' + confirmButton + '" to confirm, "No" to skip, or "Cancel" to abort.';
    } else {
      buttonSet = ui.ButtonSet.YES_NO;
      confirmButton = ui.Button.YES;
    }
    
    // Show dialog
    var response = ui.alert(title, message + '\n\n' + formattedData, buttonSet);
    
    // Handle response
    if (response === confirmButton) {
      debugLog('verifyConfigurationWithUser', 'INFO', 'User confirmed', itemName, '');
      return true;
    } else if (allowSkip && response === ui.Button.NO) {
      debugLog('verifyConfigurationWithUser', 'INFO', 'User skipped', itemName, '');
      return false; // Skipped, but continue
    } else {
      debugLog('verifyConfigurationWithUser', 'INFO', 'User cancelled', itemName, '');
      throw new Error('User cancelled ' + itemName + ' verification');
    }
    
  } catch (error) {
    debugLog('verifyConfigurationWithUser', 'ERROR', 'Verification dialog error', itemName, error.message);
    
    // Re-throw user cancellation errors
    if (error.message.indexOf('cancelled') !== -1) {
      throw error;
    }
    
    // For other errors, ask user what to do
    try {
      var fallbackResponse = SpreadsheetApp.getUi().alert(
        'Verification Error',
        'Error showing ' + itemName + ' verification: ' + error.message + '\n\nContinue anyway?',
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      return fallbackResponse === SpreadsheetApp.getUi().Button.YES;
    } catch (fallbackError) {
      return false; // Ultimate fallback
    }
  }
}