/*
================================================================================
RESPONSES CODE
================================================================================
Version: 113
Total Functions: 76
Documentation: See Responses-Functions.md
================================================================================
*/

function authorizeScript() {
  // This function just needs to be run once to authorize the script
  // It accesses the UI which triggers the authorization prompt
  try {
    SpreadsheetApp.getUi();
    SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Script authorized successfully!');
  } catch (e) {
    Logger.log('Authorization error: ' + e.message);
  }
}

function handleFormEdit(e) {
  if (!e) {
    UtilityScriptLibrary.debugLog("⚠️ handleFormEdit called without event object");
    return;
  }
  
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();
  
  // Get current semester from Calendar D2
  var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');
  if (!calendarSheet) {
    return;
  }
  
  var currentSemester = calendarSheet.getRange(2, 4).getValue();
  if (!currentSemester || String(currentSemester).trim() === '') {
    return;
  }
  
  // Only process if sheet matches current semester
  if (sheetName !== String(currentSemester).trim()) {
    return;
  }
  
  // Get Teacher column from headerMap
  var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
  var teacherCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher')];
  
  if (!teacherCol) {
    return;
  }
  
  // Only process if edited column is Teacher column
  if (editedCol !== teacherCol) {
    return;
  }
  
  // Only process if editing row 2 or higher (skip header)
  if (editedRow < 2) {
    return;
  }
  
  UtilityScriptLibrary.debugLog("🔍 Form edit detected on sheet: " + sheetName);
  
  if (!shouldProcessEdit(e, sheet)) {
    UtilityScriptLibrary.debugLog("⏭️ Skipping form edit processing");
    return;
  }

  // LOCK MECHANISM - Prevent concurrent executions
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Could not obtain lock - another execution in progress. EXITING.");
    return;
  }

  try {
    UtilityScriptLibrary.debugLog("=== STARTING handleFormEdit ===");
    UtilityScriptLibrary.debugLog("Triggered handleFormEdit on sheet: " + sheetName + ", row: " + editedRow);
    
    processSingleRow(sheet, editedRow, headerMap);
    
    Browser.msgBox("🎉 Student successfully added! You can add the next student now.");
    
    UtilityScriptLibrary.debugLog("=== COMPLETED handleFormEdit ===");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in handleFormEdit: " + error.message);
    Browser.msgBox("Error processing student: " + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  var currentSemester = UtilityScriptLibrary.getCurrentSemesterName();
  if (!currentSemester) return;

  if (sheetName !== String(currentSemester).trim()) return;

  var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
  var teacherCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher')];

  if (!teacherCol) return;
  if (editedCol !== teacherCol) return;
  if (editedRow < 2) return;

  UtilityScriptLibrary.debugLog("🔍 Form edit detected on sheet: " + sheetName);

  // Swap display name -> Teacher ID before any further processing
  var selectedValue = String(e.range.getValue()).trim();
  if (selectedValue) {
    var lookupSheet = getTeacherRosterLookupSheet();
    if (lookupSheet) {
      var displayNameMap = generateTeacherDisplayNames(lookupSheet);
      // Reverse lookup: displayName -> teacherId
      for (var teacherId in displayNameMap) {
        if (displayNameMap[teacherId] === selectedValue) {
          e.range.setValue(teacherId);
          UtilityScriptLibrary.debugLog('onEdit', 'INFO', 'Swapped display name to Teacher ID', selectedValue + ' -> ' + teacherId, '');
          break;
        }
      }
    }
  }

  if (!shouldProcessEdit(e, sheet)) {
    UtilityScriptLibrary.debugLog("⏭️ Skipping form edit processing");
    return;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Could not obtain lock - another execution in progress. EXITING.");
    return;
  }

  try {
    UtilityScriptLibrary.debugLog("=== STARTING handleFormEdit ===");
    UtilityScriptLibrary.debugLog("Triggered handleFormEdit on sheet: " + sheetName + ", row: " + editedRow);

    processSingleRow(sheet, editedRow, headerMap);

    Browser.msgBox("🎉 Student successfully added! You can add the next student now.");

    UtilityScriptLibrary.debugLog("=== COMPLETED handleFormEdit ===");

  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in handleFormEdit: " + error.message);
    Browser.msgBox("Error processing student: " + error.message);
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QAMP Tools')
    .addItem('Refresh Teacher Dropdown', 'refreshCurrentSemesterTeacherDropdown')
    .addItem('Update Roster Groups', 'updateAllTeacherGroupAssignments')
    .addItem('Process Teacher Assignments', 'processPendingAssignments')
    .addItem('Reassign Student to Different Teacher', 'reassignStudentToNewTeacher')
    .addItem('Clear Reports', 'clearReports')
    .addItem('Verify by Drive ID', 'verifyByDriveIdWithPrompt')
    .addItem('Create New Year Workbooks with Continuing Students', 'createNewYearWorkbooksWithContinuingStudents')
    .addToUi();
  
  applyTeacherDropdownToCurrentSemester();
}


function addCarryoverStudentsToNewRoster(spreadsheet, newRosterSheet, currentSemesterName) {
  try {
    UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Starting carryover process for semester: " + currentSemesterName);
    
    var previousRosterSheetName = findPreviousSemesterRoster(spreadsheet, currentSemesterName);
    if (!previousRosterSheetName) {
      UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - No previous semester roster found");
      return 0;
    }
    
    var previousRosterSheet = spreadsheet.getSheetByName(previousRosterSheetName);
    if (!previousRosterSheet) {
      UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Previous roster sheet not found: " + previousRosterSheetName);
      return 0;
    }
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(previousRosterSheet);
    
    UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Header map keys: " + Object.keys(headerMap).join(", "));
    
    var data = previousRosterSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Previous roster has no student data (only headers)");
      return 0;
    }
    
    // FIXED: Use normalized keys (no spaces)
    var statusCol = headerMap['status'];
    var lessonsRemainingCol = headerMap['lessonsremaining'];
    var studentIdCol = headerMap['studentid'];
    
    UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Found columns - Status: " + statusCol + ", Lessons Remaining: " + lessonsRemainingCol + ", Student ID: " + studentIdCol);
    
    if (!statusCol || !lessonsRemainingCol || !studentIdCol) {
      UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Required columns not found in previous roster");
      return 0;
    }
    
    var studentsToCarryOver = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = row[statusCol - 1];
      var lessonsRemaining = parseFloat(row[lessonsRemainingCol - 1]) || 0;
      var studentId = row[studentIdCol - 1];
      
      if (status && status.toString().toLowerCase() === 'active' && lessonsRemaining > 0 && studentId) {
        studentsToCarryOver.push({
          rowIndex: i + 1,
          rowData: row,
          studentId: studentId,
          lessonsRemaining: lessonsRemaining
        });
      }
    }
    
    if (studentsToCarryOver.length === 0) {
      UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - No students to carry over (no Active students with lessons remaining)");
      return 0;
    }
    
    UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Found " + studentsToCarryOver.length + " students to carry over");
    
    var newHeaderMap = UtilityScriptLibrary.getHeaderMap(newRosterSheet);
    var newStatusCol = newHeaderMap['status'];
    
    var addedCount = 0;
    for (var i = 0; i < studentsToCarryOver.length; i++) {
      var student = studentsToCarryOver[i];
      
      try {
        var alreadyExists = checkIfStudentExists(newRosterSheet, student.studentId, newHeaderMap);
        if (alreadyExists) {
          UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Student already exists, skipping: " + student.studentId);
          continue;
        }
        
        // Map data from old roster columns to new roster columns by name
        // FIXED: Initialize array without .fill() - not supported in Google Apps Script
        var newRowData = [];
        for (var j = 0; j < 23; j++) {
          newRowData[j] = '';
        }
        
        // Map each column by name from old to new
        var fieldsToMap = [
          'contacted', 'firstlessondate', 'firstlessontime', 'comments',
          'lastname', 'firstname', 'instrument', 'length', 'experience',
          'grade', 'school', 'schoolteacher',
          'parentlastname', 'parentfirstname', 'phone', 'email', 'additionalcontacts',
          'hoursremaining', 'lessonsremaining', 'status', 'studentid', 'admincomments', 'systemcomments'
        ];
        
        for (var k = 0; k < fieldsToMap.length; k++) {
          var fieldName = fieldsToMap[k];
          var oldColIndex = headerMap[fieldName];
          var newColIndex = newHeaderMap[fieldName];
          
          if (oldColIndex && newColIndex) {
            newRowData[newColIndex - 1] = student.rowData[oldColIndex - 1];
          }
        }
        
        // Override status to Carryover
        newRowData[newStatusCol - 1] = 'Carryover';
        
        // Add system comment about carryover
        var systemCommentsCol = newHeaderMap['systemcomments'];
        if (systemCommentsCol) {
          var oldComments = newRowData[systemCommentsCol - 1] || '';
          newRowData[systemCommentsCol - 1] = 'Carried over from ' + previousRosterSheetName + ' on ' + new Date().toDateString() + '. ' + oldComments;
        }
        
        var targetRow = newRosterSheet.getLastRow() + 1;
        newRosterSheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
        
        newRosterSheet.getRange(targetRow, 1).insertCheckboxes().setValue(false);
        
        var rowRange = newRosterSheet.getRange(targetRow, 1, 1, 23);
        rowRange.setBackground(UtilityScriptLibrary.STYLES.WARNING.background)
                .setFontColor(UtilityScriptLibrary.STYLES.WARNING.text)
                .setFontWeight('bold');
        
        addedCount++;
        UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Added carryover student: " + student.studentId + " to row " + targetRow);
        
      } catch (studentError) {
        UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - ERROR adding student " + student.studentId + ": " + studentError.message);
      }
    }
    
    UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - Successfully added " + addedCount + " carryover students");
    return addedCount;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addCarryoverStudentsToNewRoster - ERROR: " + error.message);
    return 0;
  }
}

function addRosterTemplateBorders(sheet) {
  try {
    var maxRows = sheet.getMaxRows();
    
    // Thick green border between D and E (between editable and admin)
    var borderRange1 = sheet.getRange(1, 4, maxRows, 1);
    borderRange1.setBorder(null, null, null, true, null, null, UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
    
    // Dotted border between H and I (between Length and Experience)
    var borderRange2 = sheet.getRange(1, 8, maxRows, 1);
    borderRange2.setBorder(null, null, null, true, null, null, UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.DOTTED);
    
    // Thick green border between Q and R (before Hours/Lessons Remaining) - UPDATED position
    var borderRange3 = sheet.getRange(1, 17, maxRows, 1);
    borderRange3.setBorder(null, null, null, true, null, null, UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
    
    UtilityScriptLibrary.debugLog("✅ Green borders added to roster template");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Error adding roster borders: " + error.message);
  }
}

function addStudentToAttendanceSheet(attendanceSheet, studentData) {
  //only one student
  try {
    UtilityScriptLibrary.debugLog("📅 Adding single student to attendance sheet: " + attendanceSheet.getName());
    
    // Log current sheet state BEFORE adding
    var lastRowBefore = attendanceSheet.getLastRow();
    UtilityScriptLibrary.debugLog("📊 Sheet state BEFORE: lastRow=" + lastRowBefore);
    
    // CRITICAL FIX: Update date validation to allow empty cells
    var maxRows = attendanceSheet.getMaxRows();
    var dateColumn = attendanceSheet.getRange(1, 3, maxRows, 1); // Column C (Date column)
    
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .build();
    dateColumn.setDataValidation(dateRule);
    
    UtilityScriptLibrary.debugLog("✅ Updated date validation to allow empty cells in column C");
    
    // Convert studentData array to student object
    var student = createStudentObjectForAttendance(studentData);
    UtilityScriptLibrary.debugLog("👤 Student object created: " + JSON.stringify(student));
    
    // Use Utility's createStudentSections with single-student array
    UtilityScriptLibrary.createStudentSections(attendanceSheet, [student]);
    
    // Log current sheet state AFTER adding
    var lastRowAfter = attendanceSheet.getLastRow();
    UtilityScriptLibrary.debugLog("📊 Sheet state AFTER: lastRow=" + lastRowAfter);
    UtilityScriptLibrary.debugLog("📍 Expected to write starting at row: " + (lastRowBefore <= 1 ? 2 : lastRowBefore + 2));
    
    // Apply status dropdown validation to all rows in the sheet
    UtilityScriptLibrary.setupStatusValidation(attendanceSheet, lastRowAfter);
    UtilityScriptLibrary.debugLog("✅ Applied status dropdown validation");
    
    // Verify data was actually written
    if (lastRowAfter > lastRowBefore) {
      UtilityScriptLibrary.debugLog("✅ Successfully added student " + student.firstName + " " + student.lastName + " to " + attendanceSheet.getName());
    } else {
      UtilityScriptLibrary.debugLog("⚠️ WARNING: lastRow did not increase - data may not have been written!");
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ ERROR in addStudentToAttendanceSheet: " + error.message);
    throw error;
  }
}

function addStudentToAttendanceSheetsFromDate(workbook, studentInfo, effectiveDate) {
  try {
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "INFO",
                                  "Starting attendance addition from date",
                                  "Student: " + studentInfo.studentId + ", Date: " + effectiveDate.toDateString(), "");
    
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var effectiveMonthIndex = effectiveDate.getMonth(); // 0-11
    
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "INFO",
                                  "Effective month calculated",
                                  "Index: " + effectiveMonthIndex + " (" + monthNames[effectiveMonthIndex] + ")", "");
    
    // Get all sheets in the workbook
    var allSheets = workbook.getSheets();
    var attendanceSheetsToUpdate = [];
    
    // Find all month sheets from effective date forward
    for (var i = 0; i < allSheets.length; i++) {
      var sheetName = allSheets[i].getName();
      var monthIndex = monthNames.indexOf(sheetName);
      
      if (monthIndex !== -1 && monthIndex >= effectiveMonthIndex) {
        attendanceSheetsToUpdate.push({
          sheet: allSheets[i],
          name: sheetName,
          monthIndex: monthIndex
        });
      }
    }
    
    // Sort by month order
    attendanceSheetsToUpdate.sort(function(a, b) {
      return a.monthIndex - b.monthIndex;
    });
    
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "INFO",
                                  "Found attendance sheets to update",
                                  attendanceSheetsToUpdate.length + " sheets: " + 
                                  attendanceSheetsToUpdate.map(function(s) { return s.name; }).join(', '), "");
    
    if (attendanceSheetsToUpdate.length === 0) {
      UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "INFO",
                                    "No attendance sheets found from " + monthNames[effectiveMonthIndex] + " forward",
                                    "Student will be added when monthly attendance sheets are created via bulk process", "");
      
    } else {
      // Add student to existing attendance sheets
      for (var j = 0; j < attendanceSheetsToUpdate.length; j++) {
        var attendanceSheet = attendanceSheetsToUpdate[j].sheet;
        
        // Check if student already exists in this sheet
        if (studentExistsInAttendanceSheet(attendanceSheet, studentInfo.studentId)) {
          UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "INFO",
                                        "Student already exists - skipping",
                                        "Sheet: " + attendanceSheet.getName(), "");
          continue;
        }
        
        // Convert student info to format expected by attendance sheet
        var studentObj = convertStudentInfoToAttendanceObject(studentInfo);
        
        // Use existing function to add student section
        UtilityScriptLibrary.createStudentSections(attendanceSheet, [studentObj]);
        
        // Apply status dropdown validation to all rows in the sheet
        var lastRow = attendanceSheet.getLastRow();
        UtilityScriptLibrary.setupStatusValidation(attendanceSheet, lastRow);
        
        UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "SUCCESS",
                                      "Added student to attendance sheet",
                                      "Sheet: " + attendanceSheet.getName(), "");
      }
    }
    
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "SUCCESS",
                                  "Completed attendance sheet updates", "", "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "ERROR",
                                  "Failed attendance addition",
                                  "", error.message);
    throw error;
  }
}

function addStudentToNewRosterTemplate(sheet, formData, studentId) {
  try {
    UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - Adding student: " + studentId + " to roster template");
    
    // Determine lesson length from form data
    var lessonLength = '';
    var lengthFromField = formData["Length"];
    
    if (lengthFromField && lengthFromField.toString().trim() !== '') {
      // Extract numeric value from "30 minutes" format or plain "30"
      lessonLength = UtilityScriptLibrary.extractNumericLessonLength(lengthFromField);
    } else {
      // Fallback: determine from package quantities
      var qty60 = UtilityScriptLibrary.extractLessonQuantityFromPackage(formData["Qty60"]);
      var qty45 = UtilityScriptLibrary.extractLessonQuantityFromPackage(formData["Qty45"]);
      var qty30 = UtilityScriptLibrary.extractLessonQuantityFromPackage(formData["Qty30"]);
      
      lessonLength = (qty60 > 0) ? 60 : (qty45 > 0) ? 45 : (qty30 > 0) ? 30 : 30;
    }
    
    UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - Determined lesson length: " + lessonLength);
    
    // Create new row data array - 23 columns
    var newRowData = new Array(23);
    for (var i = 0; i < newRowData.length; i++) {
      newRowData[i] = '';
    }
    
    // Grade logic: Check if person IS an adult (Yes response)
    var gradeValue = '';
    var ageValue = formData["Age"] || '';
    
    if (ageValue.toString().toLowerCase().indexOf('yes') === 0) {
      gradeValue = 'Adult';
    } else {
      gradeValue = formData["Grade"] || '';
    }
    
    UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - Grade logic: Age='" + ageValue + "', Final='" + gradeValue + "'");
    
    // Map form data to roster column structure
    newRowData[0] = false;                               // A: Contacted (checkbox)
    newRowData[1] = '';                                  // B: First Lesson Date (teacher fills)
    newRowData[2] = '';                                  // C: First Lesson Time (teacher fills)  
    newRowData[3] = '';                                  // D: Comments (teacher fills)
    newRowData[4] = formData["Student Last Name"] || ''; // E: Last Name
    newRowData[5] = formData["Student First Name"] || '';// F: First Name
    newRowData[6] = formData["Instrument"] || '';        // G: Instrument
    newRowData[7] = lessonLength;                        // H: Length (numeric: 30, 45, 60)
    newRowData[8] = formData["Experience"] || '';        // I: Experience
    newRowData[9] = gradeValue;                          // J: Grade (Adult or grade)
    newRowData[10] = formData["School"] || '';           // K: School
    newRowData[11] = formData["SchoolTeacher"] || '';    // L: School Teacher
    newRowData[12] = formData["Parent Last Name"] || ''; // M: Parent Last Name
    newRowData[13] = formData["Parent First Name"] || '';// N: Parent First Name
    newRowData[14] = formData["Phone"] || '';            // O: Phone
    newRowData[15] = formData["Email"] || '';            // P: Email
    newRowData[16] = formData["Additional contacts"] || '';// Q: Additional contacts
    newRowData[17] = 0;                                  // R: Hours Remaining (starts at 0, updated by sync)
    newRowData[18] = 0;                                  // S: Lessons Remaining (starts at 0, updated by sync)
    newRowData[19] = 'Active';                           // T: Status
    newRowData[20] = studentId;                          // U: Student ID
    newRowData[21] = '';                                 // V: Admin Comments
    newRowData[22] = 'Added: ' + new Date().toDateString();  // W: System Comments
    
    // Find an empty row or append
    var lastRow = sheet.getLastRow();
    var targetRow = lastRow + 1;
    
    // Look for empty rows (starting from row 2, after header)
    for (var i = 2; i <= lastRow; i++) {
      var existingData = sheet.getRange(i, 1, 1, 23).getValues()[0];
      var isEmpty = true;
      for (var j = 0; j < existingData.length; j++) {
        if (existingData[j] !== '' && existingData[j] !== null && existingData[j] !== undefined) {
          isEmpty = false;
          break;
        }
      }
      
      if (isEmpty) {
        targetRow = i;
        break;
      }
    }
    
    // Insert the student data
    if (targetRow <= lastRow) {
      sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
      UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - Inserted student into empty row: " + targetRow);
    } else {
      sheet.appendRow(newRowData);
      targetRow = sheet.getLastRow();
      UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - Appended student to new row: " + targetRow);
    }
    
    // Set checkbox for Contacted column (A)
    sheet.getRange(targetRow, 1).insertCheckboxes().setValue(false);
    
    // Apply alternating row color
    if (targetRow % 2 === 0) {
      sheet.getRange(targetRow, 1, 1, newRowData.length).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      sheet.getRange(targetRow, 1, 1, newRowData.length).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }
    
    UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - Successfully added student to roster template. Student: " + studentId + ", Row: " + targetRow);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentToNewRosterTemplate - ERROR: Error adding student to roster template. Student: " + studentId + ". Error: " + error.message);
    throw error;
  }
}

function addStudentToRosterFromData(rosterSheet, studentInfo, headerMap) {
  try {
    UtilityScriptLibrary.debugLog("addStudentToRosterFromData - Adding reassigned student: " + studentInfo.studentId);

    var newRowData = new Array(23);
    for (var i = 0; i < newRowData.length; i++) {
      newRowData[i] = '';
    }

    newRowData[0]  = false;                                          // A: Contacted (checkbox)
    newRowData[1]  = studentInfo.firstLessonDate || '';             // B: First Lesson Date
    newRowData[2]  = studentInfo.firstLessonTime || '';             // C: First Lesson Time
    newRowData[3]  = studentInfo.comments || '';                    // D: Comments
    newRowData[4]  = studentInfo.lastName || '';                    // E: Last Name
    newRowData[5]  = studentInfo.firstName || '';                   // F: First Name
    newRowData[6]  = studentInfo.instrument || '';                  // G: Instrument
    newRowData[7]  = studentInfo.length || 30;                      // H: Length
    newRowData[8]  = studentInfo.experience || '';                  // I: Experience
    newRowData[9]  = studentInfo.grade || '';                       // J: Grade
    newRowData[10] = studentInfo.school || '';                      // K: School
    newRowData[11] = studentInfo.schoolTeacher || '';               // L: School Teacher
    newRowData[12] = studentInfo.parentLastName || '';              // M: Parent Last Name
    newRowData[13] = studentInfo.parentFirstName || '';             // N: Parent First Name
    newRowData[14] = studentInfo.phone || '';                       // O: Phone
    newRowData[15] = studentInfo.email || '';                       // P: Email
    newRowData[16] = studentInfo.additionalContacts || '';          // Q: Additional contacts
    newRowData[17] = studentInfo.hoursRemaining || 0;               // R: Hours Remaining
    newRowData[18] = studentInfo.lessonsRemaining || 0;             // S: Lessons Remaining
    newRowData[19] = 'Active';                                      // T: Status
    newRowData[20] = studentInfo.studentId || '';                   // U: Student ID
    newRowData[21] = '';                                            // V: Admin Comments
    newRowData[22] = 'Reassigned: ' + new Date().toDateString();   // W: System Comments

    var lastRow = rosterSheet.getLastRow();
    var targetRow = lastRow + 1;

    // Look for an empty row before appending
    for (var i = 2; i <= lastRow; i++) {
      var existingData = rosterSheet.getRange(i, 1, 1, 23).getValues()[0];
      var isEmpty = true;
      for (var j = 0; j < existingData.length; j++) {
        if (existingData[j] !== '' && existingData[j] !== null && existingData[j] !== undefined) {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        targetRow = i;
        break;
      }
    }

    if (targetRow <= lastRow) {
      rosterSheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
    } else {
      rosterSheet.appendRow(newRowData);
      targetRow = rosterSheet.getLastRow();
    }

    rosterSheet.getRange(targetRow, 1).insertCheckboxes().setValue(false);

    if (targetRow % 2 === 0) {
      rosterSheet.getRange(targetRow, 1, 1, 23).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      rosterSheet.getRange(targetRow, 1, 1, 23).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }

    UtilityScriptLibrary.debugLog("addStudentToRosterFromData - Successfully added student at row: " + targetRow);

  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentToRosterFromData - ERROR: " + error.message);
    throw error;
  }
}

function addStudentToSemesterRoster(workbook, formData, studentId, semesterName) {
  try {
    UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Adding student to semester roster. Student: " + studentId + ", Semester: " + semesterName);
    
    // Extract season from semesterName
    var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
    if (!season) {
      throw new Error("Could not extract season from semester name: " + semesterName);
    }
    
    // Find or create the season roster sheet
    var rosterSheetName = season + " Roster";
    var rosterSheet = workbook.getSheetByName(rosterSheetName);
    
    var isNewSheet = false;
    if (!rosterSheet) {
      // Create the season roster sheet if it doesn't exist
      UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Creating missing season roster sheet: " + rosterSheetName);
      try {
        rosterSheet = workbook.insertSheet(rosterSheetName);
        setupNewRosterTemplate(rosterSheet);
        isNewSheet = true;
        UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Successfully created season roster sheet: " + rosterSheetName);
      } catch (createError) {
        throw new Error("Failed to create season roster sheet '" + rosterSheetName + "': " + createError.message);
      }
    } else {
      UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Using existing season roster sheet: " + rosterSheetName);
    }
    
    // If we just created a new sheet, add carryover students from previous semester
    if (isNewSheet) {
      try {
        var carryoverCount = addCarryoverStudentsToNewRoster(workbook, rosterSheet, semesterName);
        if (carryoverCount > 0) {
          UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Added " + carryoverCount + " carryover students from previous semester");
        } else {
          UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - No carryover students found or no previous semester roster");
        }
      } catch (carryoverError) {
        UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - WARNING: Error adding carryover students: " + carryoverError.message);
        // Non-critical - continue with student registration
      }
    }
    
    // Check if student already exists
    try {
      var headerMap = UtilityScriptLibrary.getHeaderMap(rosterSheet);
      var existsResult = checkIfStudentExists(rosterSheet, studentId, headerMap);
      
      if (existsResult === 'CARRYOVER') {
        // Convert Carryover to Active
        UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Converting Carryover student to Active: " + studentId);
        convertCarryoverToActive(rosterSheet, studentId, formData, headerMap);
        UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Successfully converted Carryover to Active");
        return;
      } else if (existsResult === true) {
        UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Student already exists in roster: " + studentId);
        return;
      }
    } catch (checkError) {
      UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - WARNING: Could not check for existing student: " + checkError.message);
    }
    
    // Add student to the roster
    try {
      addStudentToNewRosterTemplate(rosterSheet, formData, studentId);
      UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - Successfully added student to season roster.");
    } catch (addError) {
      throw new Error("Failed to add student to roster: " + addError.message);
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentToSemesterRoster - ERROR: " + error.message);
    throw error;
  }
}

function appendToReports(detailIssues, summaryData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Detail Report
    var detailSheet = ss.getSheetByName('Student ID Detail Report');
    if (!detailSheet) {
      detailSheet = ss.insertSheet('Student ID Detail Report');
      var headers = ['Workbook Name', 'Sheet Name', 'Row Number', 'First Name', 'Last Name', 'Found Student ID', 'Expected Student ID', 'Issue Type'];
      detailSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      detailSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#37a247')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      detailSheet.setColumnWidth(1, 200);
      detailSheet.setColumnWidth(2, 180);
      detailSheet.setColumnWidth(3, 80);
      detailSheet.setColumnWidth(4, 120);
      detailSheet.setColumnWidth(5, 120);
      detailSheet.setColumnWidth(6, 120);
      detailSheet.setColumnWidth(7, 120);
      detailSheet.setColumnWidth(8, 150);
      detailSheet.setFrozenRows(1);
    }
    
    if (detailIssues.length > 0) {
      var detailData = detailIssues.map(function(issue) {
        return [issue.workbookName, issue.sheetName, issue.rowNumber, issue.firstName, issue.lastName, issue.foundId, issue.expectedId, issue.issueType];
      });
      var lastRow = detailSheet.getLastRow();
      detailSheet.getRange(lastRow + 1, 1, detailData.length, 8).setValues(detailData);
    }
    
    // Summary Report
    var summarySheet = ss.getSheetByName('Student ID Summary Report');
    if (!summarySheet) {
      summarySheet = ss.insertSheet('Student ID Summary Report');
      var headers = ['Workbook Name', 'Sheet Name', 'Status', 'Issue Count'];
      summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      summarySheet.getRange(1, 1, 1, headers.length)
        .setBackground('#37a247')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      summarySheet.setColumnWidth(1, 200);
      summarySheet.setColumnWidth(2, 180);
      summarySheet.setColumnWidth(3, 80);
      summarySheet.setColumnWidth(4, 100);
      summarySheet.setFrozenRows(1);
    }
    
    if (summaryData.length > 0) {
      var summaryRows = summaryData.map(function(item) {
        return [item.workbookName, item.sheetName, item.status, item.issueCount];
      });
      var lastRow = summarySheet.getLastRow();
      summarySheet.getRange(lastRow + 1, 1, summaryRows.length, 4).setValues(summaryRows);
    }
    
    Logger.log('✅ Appended to reports');
    
  } catch (error) {
    Logger.log('❌ Error appending to reports: ' + error.message);
    throw error;
  }
}

function applyTeacherDropdownToCurrentSemester() {
  try {
    UtilityScriptLibrary.debugLog('applyTeacherDropdownToCurrentSemester', 'INFO', 'Starting auto-apply teacher dropdown', '', '');
    
    // Get current semester name from calendar
    var currentSemester = UtilityScriptLibrary.getCurrentSemesterName();
    if (!currentSemester) {
      UtilityScriptLibrary.debugLog('applyTeacherDropdownToCurrentSemester', 'WARNING', 'No current semester found', '', '');
      return;
    }
    
    // Find the current semester sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var semesterSheet = ss.getSheetByName(currentSemester);
    
    if (!semesterSheet) {
      UtilityScriptLibrary.debugLog('applyTeacherDropdownToCurrentSemester', 'WARNING', 'Current semester sheet not found', currentSemester, '');
      return;
    }
    
    // Apply teacher dropdown
    applyTeacherDropdownToSheet(semesterSheet);
    
    UtilityScriptLibrary.debugLog('applyTeacherDropdownToCurrentSemester', 'INFO', 'Successfully applied teacher dropdown', currentSemester, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('applyTeacherDropdownToCurrentSemester', 'ERROR', 'Failed to apply teacher dropdown', '', error.message);
    // Don't show errors to user on onOpen - just log them
  }
}

function applyTeacherDropdownToSheet(sheet) {
  try {
    // Get active teachers for dropdown (using existing function)
    var activeTeachers = getActiveTeachersForDropdown();
    
    if (activeTeachers.length === 0) {
      UtilityScriptLibrary.debugLog('applyTeacherDropdownToSheet', 'WARNING', 'No active teachers found for dropdown', sheet.getName(), '');
      return;
    }
    
    // Find Teacher column using headerMap
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var teacherCol = headerMap["teacher"] || headerMap["instructor"];
    
    if (!teacherCol) {
      UtilityScriptLibrary.debugLog('applyTeacherDropdownToSheet', 'WARNING', 'Teacher column not found in sheet', sheet.getName(), '');
      return;
    }
    
    // Create validation rule
    var teacherRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(activeTeachers)
      .setAllowInvalid(true)  // Allow existing values that might not be in current list
      .setHelpText('Select a teacher from the dropdown. List shows currently active teachers.')
      .build();
    
    // Apply to entire Teacher column (skip header row)
    var lastRow = Math.max(sheet.getLastRow(), 1000); // Ensure we cover future rows
    var teacherRange = sheet.getRange(2, teacherCol, lastRow - 1, 1);
    teacherRange.setDataValidation(teacherRule);
    
    UtilityScriptLibrary.debugLog('applyTeacherDropdownToSheet', 'INFO', 'Applied teacher dropdown validation', 
                                  'Sheet: ' + sheet.getName() + ', Column: ' + teacherCol + ', Teachers: ' + activeTeachers.length, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('applyTeacherDropdownToSheet', 'ERROR', 'Failed to apply teacher dropdown', sheet.getName(), error.message);
    throw error;
  }
}

function calculateExperienceStartRange(experience) {
  try {
    UtilityScriptLibrary.debugLog("calculateExperienceStartRange - START - Input: " + JSON.stringify(experience));
    
    if (!experience) return '';
    
    var currentYear = new Date().getFullYear();
    var exp = String(experience).toLowerCase().trim();
    
    UtilityScriptLibrary.debugLog("calculateExperienceStartRange - Processing exp: '" + exp + "'");
    
    if (exp === 'none') {
      return String(currentYear);
    } else if (exp.indexOf('less than 2') !== -1) {
      return (currentYear - 1) + '-' + currentYear;
    } else if (exp.indexOf('2-4') !== -1) {
      return (currentYear - 4) + '-' + (currentYear - 2);
    } else if (exp.indexOf('4-6') !== -1) {
      return (currentYear - 6) + '-' + (currentYear - 4);
    } else if (exp.indexOf('more than 6') !== -1) {
      return 'Before ' + (currentYear - 6);
    }
    
    return '';
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("calculateExperienceStartRange - ERROR: " + error.message);
    return '';
  }
}

function calculateGraduationYear(grade) {
  try {
    UtilityScriptLibrary.debugLog("calculateGraduationYear - START - Input grade: " + JSON.stringify(grade) + " (type: " + typeof grade + ")");
    
    if (!grade) {
      UtilityScriptLibrary.debugLog("calculateGraduationYear - EARLY EXIT - Grade is falsy");
      return '';
    }
    
    var currentYear = new Date().getFullYear();
    var gradeStr = String(grade).toLowerCase().trim();
    
    UtilityScriptLibrary.debugLog("calculateGraduationYear - Processing gradeStr: '" + gradeStr + "'");
    
    // Replace includes() with indexOf() for compatibility
    if (gradeStr.indexOf('adult') !== -1 || gradeStr.indexOf('college') !== -1) {
      UtilityScriptLibrary.debugLog("calculateGraduationYear - Found adult/college keyword");
      return 'Adult';
    }
    
    // Handle various grade formats
    if (gradeStr.indexOf('k') !== -1 || gradeStr === 'kindergarten') {
      UtilityScriptLibrary.debugLog("calculateGraduationYear - Found kindergarten");
      return currentYear + 13;
    }
    
    // Extract number from grade (handles "1st", "2nd", "1", "2", etc.)
    var gradeNum = parseInt(gradeStr.replace(/[^\d]/g, ''));
    UtilityScriptLibrary.debugLog("calculateGraduationYear - Extracted grade number: " + gradeNum);
    
    if (isNaN(gradeNum) || gradeNum < 1 || gradeNum > 12) {
      UtilityScriptLibrary.debugLog("calculateGraduationYear - Invalid grade number, returning empty string");
      return '';
    }
    
    var result = currentYear + (13 - gradeNum);
    UtilityScriptLibrary.debugLog("calculateGraduationYear - SUCCESS - Calculated graduation year: " + result);
    return result;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("calculateGraduationYear - ERROR: " + error.message);
    return '';
  }
}

function checkIfStudentExists(rosterSheet, studentId, headerMap) {
  try {
    var rosterData = rosterSheet.getDataRange().getValues();
    var studentIdCol = headerMap["studentid"]; // FIXED: removed space (normalized key)
    
    if (!studentIdCol) {
      UtilityScriptLibrary.debugLog("checkIfStudentExists - Student ID column not found in roster");
      return false;
    }
    
    for (var i = 1; i < rosterData.length; i++) {
      var existingId = rosterData[i][studentIdCol - 1];
      if (existingId && existingId.toString().trim() === studentId.toString().trim()) {
        // Student exists - check if they're a Carryover student
        var statusCol = headerMap["status"];
        if (statusCol) {
          var currentStatus = rosterData[i][statusCol - 1];
          if (currentStatus && currentStatus.toString() === 'Carryover') {
            UtilityScriptLibrary.debugLog("checkIfStudentExists - Found Carryover student, will convert to Active: " + studentId);
            return 'CARRYOVER';
          }
        }
        
        UtilityScriptLibrary.debugLog("checkIfStudentExists - Student already exists in roster: " + studentId);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("checkIfStudentExists - ERROR: " + error.message);
    return false;
  }
}

function checkSheet(workbook, workbookName, sheet, studentMap, detailIssues, summaryData) {
  try {
    var sheetName = sheet.getName();
    var data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      return;
    }
    
    var headers = data[0];
    var norm = UtilityScriptLibrary.normalizeHeader;
    
    var idCol = -1;
    var firstNameCol = -1;
    var lastNameCol = -1;
    
    for (var h = 0; h < headers.length; h++) {
      var normalizedHeader = norm(headers[h]);
      
      // ID column - try "Student ID" first, then "ID"
      if (normalizedHeader === norm('Student ID')) {
        idCol = h;
      } else if (idCol === -1 && normalizedHeader === norm('ID')) {
        idCol = h;
      }
      
      // First Name - prioritize "Student First Name" over "First Name"
      if (normalizedHeader === norm('Student First Name')) {
        firstNameCol = h;
      } else if (firstNameCol === -1 && normalizedHeader === norm('First Name')) {
        firstNameCol = h;
      }
      
      // Last Name - prioritize "Student Last Name" over "Last Name"
      if (normalizedHeader === norm('Student Last Name')) {
        lastNameCol = h;
      } else if (lastNameCol === -1 && normalizedHeader === norm('Last Name')) {
        lastNameCol = h;
      }
    }
    
    if (idCol === -1 || firstNameCol === -1 || lastNameCol === -1) {
      return;
    }
    
    var sheetIssues = [];
    
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var foundId = (row[idCol] || '').toString().trim();
      var firstName = (row[firstNameCol] || '').toString().trim();
      var lastName = (row[lastNameCol] || '').toString().trim();
      
      // Skip completely empty rows
      if (!foundId && !firstName && !lastName) {
        continue;
      }
      
      // Skip rows with non-Q IDs (P, T, G, etc.)
      if (foundId && foundId.charAt(0) !== 'Q') {
        continue;
      }
      
      var issue = null;
      
      if (foundId.charAt(0) === 'Q' && (!firstName || !lastName)) {
        // ID without name
        issue = {
          workbookName: workbookName,
          sheetName: sheetName,
          rowNumber: r + 1,
          firstName: firstName,
          lastName: lastName,
          foundId: foundId,
          expectedId: '',
          issueType: 'ID without name'
        };
        
      } else if (firstName && lastName && !foundId) {
        // Name without ID
        var key = firstName.toLowerCase() + '|' + lastName.toLowerCase();
        var expectedId = studentMap[key] || 'NOT FOUND';
        
        issue = {
          workbookName: workbookName,
          sheetName: sheetName,
          rowNumber: r + 1,
          firstName: firstName,
          lastName: lastName,
          foundId: '',
          expectedId: expectedId,
          issueType: 'Name without ID'
        };
        
      } else if (foundId.charAt(0) === 'Q' && firstName && lastName) {
        // Both present - check for mismatch
        var key = firstName.toLowerCase() + '|' + lastName.toLowerCase();
        var expectedId = studentMap[key];
        
        if (!expectedId) {
          issue = {
            workbookName: workbookName,
            sheetName: sheetName,
            rowNumber: r + 1,
            firstName: firstName,
            lastName: lastName,
            foundId: foundId,
            expectedId: 'NOT FOUND IN CONTACTS',
            issueType: 'Student not in Contacts'
          };
        } else if (foundId !== expectedId) {
          issue = {
            workbookName: workbookName,
            sheetName: sheetName,
            rowNumber: r + 1,
            firstName: firstName,
            lastName: lastName,
            foundId: foundId,
            expectedId: expectedId,
            issueType: 'ID Mismatch'
          };
        }
      }
      
      if (issue) {
        sheetIssues.push(issue);
        detailIssues.push(issue);
      }
    }
    
    summaryData.push({
      workbookName: workbookName,
      sheetName: sheetName,
      status: sheetIssues.length === 0 ? '✓' : '✗',
      issueCount: sheetIssues.length
    });
    
    if (sheetIssues.length > 0) {
      Logger.log('    ⚠️ Found ' + sheetIssues.length + ' issues in sheet: ' + sheetName);
    }
    
  } catch (error) {
    Logger.log('❌ Error checking sheet ' + workbookName + ' - ' + sheetName + ': ' + error.message);
  }
}

function checkWorkbook(workbook, workbookName, studentMap, detailIssues, summaryData) {
  try {
    var sheets = workbook.getSheets();
    Logger.log('  📄 Checking ' + sheets.length + ' sheets in: ' + workbookName);
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      checkSheet(workbook, workbookName, sheet, studentMap, detailIssues, summaryData);
    }
    
  } catch (error) {
    Logger.log('❌ Error checking workbook ' + workbookName + ': ' + error.message);
  }
}

function checkWorkbooksInFolder(folder, studentMap, detailIssues, summaryData, isHomeFolder) {
  try {
    var files = folder.getFiles();
    var excludedNames = ['Contacts', 'Teacher Interest Survey Responses'];
    var processedCount = 0;
    
    while (files.hasNext()) {
      var file = files.next();
      
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
        continue;
      }
      
      var workbookName = file.getName();
      
      if (isHomeFolder && excludedNames.indexOf(workbookName) !== -1) {
        Logger.log('⏭️ Skipping excluded workbook: ' + workbookName);
        continue;
      }
      
      try {
        Logger.log('📖 Opening workbook: ' + workbookName);
        var workbook = SpreadsheetApp.openById(file.getId());
        checkWorkbook(workbook, workbookName, studentMap, detailIssues, summaryData);
        processedCount++;
        
        if (processedCount % 5 === 0) {
          Utilities.sleep(1000);
          Logger.log('⏸️ Processed ' + processedCount + ' workbooks, pausing briefly...');
        }
        
      } catch (error) {
        Logger.log('⚠️ Could not access workbook ' + workbookName + ': ' + error.message);
        continue;
      }
    }
    
    Logger.log('✅ Completed folder check. Processed ' + processedCount + ' workbooks.');
    
  } catch (error) {
    Logger.log('❌ Error checking workbooks in folder: ' + error.message);
  }
}

function clearReports() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var detailSheet = ss.getSheetByName('Student ID Detail Report');
    if (detailSheet) {
      ss.deleteSheet(detailSheet);
    }
    
    var summarySheet = ss.getSheetByName('Student ID Summary Report');
    if (summarySheet) {
      ss.deleteSheet(summarySheet);
    }
    
    Logger.log('✅ Reports cleared');
    Browser.msgBox('Reports cleared. Ready for new verification run.');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    Browser.msgBox('Error: ' + error.message);
  }
}

function convertCarryoverToActive(rosterSheet, studentId, formData, headerMap) {
  try {
    UtilityScriptLibrary.debugLog("convertCarryoverToActive - Converting student: " + studentId);
    
    var data = rosterSheet.getDataRange().getValues();
    var studentIdCol = headerMap["studentid"];
    
    // Find the student row
    var studentRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol - 1] && data[i][studentIdCol - 1].toString().trim() === studentId.toString().trim()) {
        studentRow = i + 1; // Convert to 1-based
        break;
      }
    }
    
    if (studentRow === -1) {
      throw new Error("Could not find student row for ID: " + studentId);
    }
    
    // Get new lesson package from form data
    var newLessonsRegistered = parseInt(formData["Lesson Quantity"]) || 0;
    
    // Determine lesson length from Qty fields
    var lessonLength = 30; // default
    if (formData["Qty60"]) {
      lessonLength = 60;
    } else if (formData["Qty45"]) {
      lessonLength = 45;
    } else if (formData["Qty30"]) {
      lessonLength = 30;
    }
    
    // Determine grade value
    var ageValue = formData["Age"] || '';
    var gradeValue = '';
    if (ageValue.toString().toLowerCase().charAt(0) === 'y') {
      gradeValue = 'Adult';
    } else {
      gradeValue = formData["Grade"] || '';
    }
    
    // Update all relevant fields from new registration
    var updates = {
      'length': lessonLength,
      'experience': formData["Experience"] || '',
      'grade': gradeValue,
      'school': formData["School"] || '',
      'schoolteacher': formData["SchoolTeacher"] || '',
      'parentlastname': formData["Parent Last Name"] || '',
      'parentfirstname': formData["Parent First Name"] || '',
      'phone': formData["Phone"] || '',
      'email': formData["Email"] || '',
      'additionalcontacts': formData["Additional contacts"] || '',
      'hoursremaining': 0, // Reset, will be updated by sync
      'lessonsremaining': 0, // Reset, will be updated by sync
      'status': 'Active'
    };
    
    // Apply all updates
    for (var fieldName in updates) {
      var colIndex = headerMap[fieldName];
      if (colIndex) {
        rosterSheet.getRange(studentRow, colIndex).setValue(updates[fieldName]);
      }
    }
    
    // Update System Comments
    var systemCommentsCol = headerMap['systemcomments'];
    if (systemCommentsCol) {
      var oldComments = data[studentRow - 1][systemCommentsCol - 1] || '';
      var newComment = 'Re-registered on ' + new Date().toDateString() + ' with ' + newLessonsRegistered + ' lessons. ';
      rosterSheet.getRange(studentRow, systemCommentsCol).setValue(newComment + oldComments);
    }
    
    // Remove WARNING formatting from entire row (A-W = 23 columns)
    var rowRange = rosterSheet.getRange(studentRow, 1, 1, 23);
    rowRange.setBackground('#ffffff')
            .setFontColor('#000000')
            .setFontWeight('normal');
    
    // Reapply alternating row color
    if (studentRow % 2 === 0) {
      rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }
    
    UtilityScriptLibrary.debugLog("convertCarryoverToActive - Successfully converted student. New lessons: " + newLessonsRegistered);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("convertCarryoverToActive - ERROR: " + error.message);
    throw error;
  }
}

function convertStudentInfoToAttendanceObject(studentInfo) {
  return {
    id: studentInfo.studentId || '',
    lastName: studentInfo.lastName || '',
    firstName: studentInfo.firstName || '',
    instrument: studentInfo.instrument || '',
    lessonLength: studentInfo.length || 30,
    lessonsRegistered: studentInfo.lessonsRemaining || 0,
    lessonsCompleted: 0,
    lessonsRemaining: studentInfo.lessonsRemaining || 0,
    status: 'Active'
  };
}

function createInvoiceLogSheet(spreadsheet) {
  try {
    UtilityScriptLibrary.debugLog('Creating Invoice Log sheet');
    
    // Create the sheet
    var sheet = spreadsheet.insertSheet('Invoice Log');
    
    // Set up column headers
    setupInvoiceLogHeaders(sheet);
    
    // Apply formatting
    formatInvoiceLogSheet(sheet);
    
    UtilityScriptLibrary.debugLog('✅ Successfully created Invoice Log sheet');
    return sheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error creating Invoice Log sheet: ' + error.message);
    throw error;
  }
}

function createNewYearWorkbooksWithContinuingStudents() {
  var ui = SpreadsheetApp.getUi();

  try {
    var verifyResponse = ui.alert(
      'Teacher Status Verification',
      'Have you verified teacher status in Teacher Roster Lookup?\n\n' +
      '(Only Active teachers will get new workbooks)',
      ui.ButtonSet.YES_NO
    );

    if (verifyResponse !== ui.Button.YES) {
      ui.alert('Cancelled', 'Please verify teacher status first.', ui.ButtonSet.OK);
      return;
    }

    var semesterName = UtilityScriptLibrary.getCurrentSemesterName();
    if (!semesterName) throw new Error('No current semester found in Semester Metadata');

    var newYear = UtilityScriptLibrary.getYearFromSemesterName(semesterName);
    if (!newYear) throw new Error('Could not extract year from semester name: ' + semesterName);

    var previousYear = String(parseInt(newYear) - 1);

    var confirm = ui.alert(
      'Create ' + newYear + ' Teacher Workbooks',
      'This will:\n' +
      '• Create new QAMP ' + newYear + ' workbooks\n' +
      '• Copy continuing students from ' + previousYear + ' rosters\n' +
      '• Create "' + semesterName + '" roster sheets\n' +
      '• Update Teacher Roster Lookup URLs\n\n' +
      'Continuing students: Lessons Remaining > 0 AND Status = Active/Carryover\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (confirm !== ui.Button.YES) return;

    var folders = getYearRosterFolders(previousYear, newYear);

    var teacherLookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    if (!teacherLookupSheet) throw new Error('Teacher Roster Lookup sheet not found');

    var getCol = UtilityScriptLibrary.createColumnFinder(teacherLookupSheet);
    var firstNameCol = getCol('First Name');
    var lastNameCol = getCol('Last Name');
    var teacherIdCol = getCol('Teacher ID');
    var displayNameCol = getCol('Display Name');
    var statusCol = getCol('Status');
    var urlCol = getCol('Roster URL');

    if (!firstNameCol || !lastNameCol || !teacherIdCol || !statusCol || !urlCol) {
      throw new Error('Required columns not found in Teacher Roster Lookup');
    }

    var teacherData = teacherLookupSheet.getDataRange().getValues();
    var stats = { processed: 0, skipped: 0, errors: [] };

    for (var i = 1; i < teacherData.length; i++) {
      var row = teacherData[i];
      var firstName = String(row[firstNameCol - 1]).trim();
      var lastName = String(row[lastNameCol - 1]).trim();
      var teacherId = String(row[teacherIdCol - 1]).trim();
      var displayName = displayNameCol ? String(row[displayNameCol - 1]).trim() : '';
      var status = String(row[statusCol - 1]).trim().toLowerCase();

      if (!firstName || !lastName || status !== 'active') {
        if (firstName || lastName) {
          UtilityScriptLibrary.debugLog('createNewYearWorkbooksWithContinuingStudents', 'INFO',
            'Skipping', firstName + ' ' + lastName + ' (' + status + ')', '');
          stats.skipped++;
        }
        continue;
      }

      var teacherInfo = {
        firstName: firstName,
        lastName: lastName,
        teacherId: teacherId,
        displayName: displayName
      };

      try {
        var result = createNewYearWorkbookForTeacher(teacherInfo, previousYear, newYear, folders.previous, folders.next, semesterName);

        if (!result) {
          stats.skipped++;
          continue;
        }

        teacherLookupSheet.getRange(i + 1, urlCol).setValue(result.url);
        stats.processed++;
        UtilityScriptLibrary.debugLog('createNewYearWorkbooksWithContinuingStudents', 'INFO',
          'Created workbook', firstName + ' ' + lastName + ' — ' + result.studentCount + ' students', '');

      } catch (teacherError) {
        stats.errors.push({ teacher: firstName + ' ' + lastName, error: teacherError.message });
        UtilityScriptLibrary.debugLog('createNewYearWorkbooksWithContinuingStudents', 'ERROR',
          'Failed', firstName + ' ' + lastName, teacherError.message);
      }
    }

    var summary = 'Workbook Creation Complete\n\n' +
                  'Semester: ' + semesterName + '\n' +
                  'Workbooks Created: ' + stats.processed + '\n' +
                  'Skipped: ' + stats.skipped;

    if (stats.errors.length > 0) {
      summary += '\n\nErrors (' + stats.errors.length + '):';
      for (var j = 0; j < Math.min(stats.errors.length, 5); j++) {
        summary += '\n• ' + stats.errors[j].teacher + ': ' + stats.errors[j].error;
      }
      if (stats.errors.length > 5) {
        summary += '\n• ... and ' + (stats.errors.length - 5) + ' more';
      }
    }

    ui.alert('✅ Complete', summary, ui.ButtonSet.OK);

  } catch (error) {
    UtilityScriptLibrary.debugLog('createNewYearWorkbooksWithContinuingStudents', 'ERROR', 'Fatal error', '', error.message);
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}

function createNewYearWorkbookForTeacher(teacherInfo, previousYear, newYear, previousFolder, newFolder, semesterName) {
  var fullName = teacherInfo.firstName + ' ' + teacherInfo.lastName;
  var previousFileName = 'QAMP ' + previousYear + ' ' + teacherInfo.lastName + ' ' + teacherInfo.teacherId;
  var previousFiles = previousFolder.getFilesByName(previousFileName);

  if (!previousFiles.hasNext()) {
    UtilityScriptLibrary.debugLog('createNewYearWorkbookForTeacher', 'INFO', 'No previous workbook found', previousFileName, '');
    return null;
  }

  var previousSS = SpreadsheetApp.openById(previousFiles.next().getId());
  var continuingStudents = getContinuingStudentsFromWorkbook(previousSS);

  UtilityScriptLibrary.debugLog('createNewYearWorkbookForTeacher', 'INFO',
    'Found continuing students', fullName + ' — ' + continuingStudents.length, '');

  var newSS = getOrCreateRosterFromTemplate(teacherInfo, newFolder, newYear, semesterName);

  if (continuingStudents.length > 0) {
    populateRosterWithContinuingStudents(newSS, semesterName, continuingStudents);
  }

  return {
    url: newSS.getUrl(),
    studentCount: continuingStudents.length
  };
}

function createStudentObjectForAttendance(studentData) {
  return {
    id: studentData[studentData.length - 1] || '',  // Last element is Student ID
    lastName: studentData[0] || '',
    firstName: studentData[1] || '',
    instrument: studentData[4] || '',
    lessonLength: studentData[3] || 30,
    lessonsRegistered: studentData[2] || 0,
    lessonsCompleted: 0,
    lessonsRemaining: studentData[2] || 0,
    status: 'Active'
  };
}

function enterEffectiveDate(newTeacherDisplay) {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();
    
    if (!newTeacherDisplay) {
      UtilityScriptLibrary.debugLog("User cancelled at New Teacher prompt");
      return;
    }
    
    var oldTeacherDisplay = scriptProps.getProperty('reassign_oldTeacherDisplay');
    
    if (newTeacherDisplay === oldTeacherDisplay) {
      ui.alert('Error', 'New teacher cannot be the same as old teacher.', ui.ButtonSet.OK);
      return;
    }
    
    UtilityScriptLibrary.debugLog("New teacher (display name): " + newTeacherDisplay);
    
    // Store new teacher
    scriptProps.setProperty('reassign_newTeacherDisplay', newTeacherDisplay);
    
    // Prompt for Effective Date
    var effectiveDateResponse = ui.prompt(
      'Reassign Students - Step 4 of 4',
      'Enter the effective date (MM/DD/YYYY):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (effectiveDateResponse.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog("User cancelled at Effective Date prompt");
      return;
    }
    
    var effectiveDateStr = effectiveDateResponse.getResponseText().trim();
    var effectiveDate;
    try {
      effectiveDate = new Date(effectiveDateStr);
      if (isNaN(effectiveDate.getTime())) {
        throw new Error("Invalid date");
      }
    } catch (error) {
      ui.alert('Error', 'Invalid date format. Please use MM/DD/YYYY.', ui.ButtonSet.OK);
      return;
    }
    
    UtilityScriptLibrary.debugLog("Effective date: " + effectiveDate);
    
    // Store effective date
    scriptProps.setProperty('reassign_effectiveDate', effectiveDate.toISOString());
    
    // Move to final step: process reassignment
    processReassignment();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in enterEffectiveDate: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 4 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function extractFormData(sheet, row, headerMap, fieldMap) {
  var rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var formData = {};
  for (var normHeader in headerMap) {
    var colIndex = headerMap[normHeader] - 1;
    var internalName = fieldMap[normHeader];
    if (internalName) {
      formData[internalName] = rowValues[colIndex];
    }
  }
  return formData;
}

function extractStudentDataFromRoster(studentRow, headerMap) {
  // Extract all student data from roster row using header map
  var getVal = function(fieldName) {
    var col = headerMap[fieldName.toLowerCase().replace(/ /g, '')];
    return col ? studentRow[col - 1] : '';
  };
  
  return {
    studentId: getVal('Student ID'),
    lastName: getVal('Last Name'),
    firstName: getVal('First Name'),
    instrument: getVal('Instrument'),
    length: getVal('Length'),
    experience: getVal('Experience'),
    grade: getVal('Grade'),
    school: getVal('School'),
    schoolTeacher: getVal('School Teacher'),
    parentLastName: getVal('Parent Last Name'),
    parentFirstName: getVal('Parent First Name'),
    phone: getVal('Phone'),
    email: getVal('Email'),
    additionalContacts: getVal('Additional contacts'),
    hoursRemaining: getVal('Hours Remaining'),
    lessonsRemaining: getVal('Lessons Remaining'),
    status: 'Active',
    firstLessonDate: getVal('First Lesson Date'),
    firstLessonTime: getVal('First Lesson Time'),
    comments: getVal('Comments')
  };
}

function findMostRecentRosterSheet(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var rosterSheets = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    // Look for sheets with "Roster" in the name
    if (sheetName.toLowerCase().indexOf('roster') !== -1) {
      rosterSheets.push(sheets[i]);
    }
  }
  
  // Return the first roster sheet found (usually there's only one named "[Season] Roster")
  // If multiple, they're typically in chronological order, so last one is most recent
  return rosterSheets.length > 0 ? rosterSheets[rosterSheets.length - 1] : null;
}

function findPreviousSemesterRoster(spreadsheet, currentSemesterName) {
  try {
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Looking for previous semester roster. Current: " + currentSemesterName);

    var semesterSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterSheet) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Semester Metadata sheet not found");
      return null;
    }

    var data = semesterSheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol = -1, startDateCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var norm = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (norm === 'semestername') nameCol = i;
      if (norm === 'startdate') startDateCol = i;
    }

    if (nameCol === -1 || startDateCol === -1) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Required columns not found in Semester Metadata");
      return null;
    }

    var currentStartDate = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][nameCol] === currentSemesterName) {
        currentStartDate = new Date(data[i][startDateCol]);
        break;
      }
    }

    if (!currentStartDate) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Current semester not found in metadata: " + currentSemesterName);
      return null;
    }

    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Current semester start date: " + currentStartDate);

    var sheets = spreadsheet.getSheets();
    var existingRosterSheets = [];
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName.toLowerCase().indexOf('roster') !== -1) {
        existingRosterSheets.push(sheetName);
      }
    }

    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Found " + existingRosterSheets.length + " roster sheets in workbook");

    var candidates = [];
    for (var i = 1; i < data.length; i++) {
      var semesterName = data[i][nameCol];
      var startDate = new Date(data[i][startDateCol]);
      if (startDate >= currentStartDate) continue;
      var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
      if (!season) continue;
      var expectedSheetName = season + " Roster";
      var hasSheet = false;
      for (var j = 0; j < existingRosterSheets.length; j++) {
        if (existingRosterSheets[j] === expectedSheetName) {
          hasSheet = true;
          break;
        }
      }
      if (hasSheet) {
        candidates.push({ semesterName: semesterName, startDate: startDate, sheetName: expectedSheetName });
      }
    }

    if (candidates.length === 0) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - No previous semester rosters found in this workbook");
      return null;
    }

    candidates.sort(function(a, b) { return b.startDate - a.startDate; });
    var previousSemester = candidates[0];
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Found previous semester: " + previousSemester.semesterName + " (Sheet: " + previousSemester.sheetName + ")");
    return previousSemester.sheetName;

  } catch (error) {
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - ERROR: " + error.message);
    return null;
  }
}

function findSemesterRoster(workbook, semesterName) {
  try {
    // Extract season from semester name (e.g., "Spring 2024" -> "Spring")
    var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
    
    if (!season) {
      UtilityScriptLibrary.debugLog("findSemesterRoster", "ERROR", "Could not extract season from semester name", 
                            "Semester: " + semesterName, "");
      return null;
    }
    
    var rosterSheetName = season + " Roster";
    var rosterSheet = workbook.getSheetByName(rosterSheetName);
    
    if (!rosterSheet) {
      UtilityScriptLibrary.debugLog("findSemesterRoster", "WARNING", "Semester roster sheet not found", 
                            "Sheet: " + rosterSheetName, "");
      return null;
    }
    
    UtilityScriptLibrary.debugLog("findSemesterRoster", "DEBUG", "Found semester roster sheet", 
                          "Sheet: " + rosterSheetName, "");
    return rosterSheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("findSemesterRoster", "ERROR", "Error finding semester roster", 
                          "Semester: " + semesterName, error.message);
    return null;
  }
}

function formatInvoiceLogSheet(sheet) {
  try {
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set up date format for Invoice Date column (column B)
    var maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 2, maxRows - 1, 1).setNumberFormat('MM/dd/yyyy');
    }
    
    // Set up currency format for Total Amount column (column E)  
    if (maxRows > 1) {
      sheet.getRange(2, 5, maxRows - 1, 1).setNumberFormat('$#,##0.00');
    }
    
    // Set text wrapping for all columns
    sheet.getRange(1, 1, maxRows, sheet.getLastColumn()).setWrap(true);
    
    // PROTECT THE ENTIRE SHEET - view-only for teachers
    var protection = sheet.protect();
    protection.setDescription('Invoice Log - View Only (automated data)');
    protection.setWarningOnly(false); // Hard protection, not just warning
    
    UtilityScriptLibrary.debugLog('✅ Applied Invoice Log formatting and protection');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('⚠️ Error in Invoice Log formatting: ' + error.message);
  }
}

function generateAttendanceSheetFromRoster(teacherWorkbook, monthName) {
  try {
    UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - Generating ' + monthName + ' attendance sheet');

    var rosterSheet = findMostRecentRosterSheet(teacherWorkbook);
    if (!rosterSheet) {
      throw new Error('No roster sheet found in workbook');
    }

    var rosterData = UtilityScriptLibrary.extractRosterData(rosterSheet);

    // Check if month sheet already exists
    var existingSheet = teacherWorkbook.getSheetByName(monthName);
    if (existingSheet) {
      UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - Sheet ' + monthName + ' already exists, skipping');
      return existingSheet;
    }

    var attendanceSheet = UtilityScriptLibrary.createMonthlyAttendanceSheet(teacherWorkbook, monthName, rosterData);

    UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - ✅ Generated ' + monthName + ' attendance sheet with ' + rosterData.length + ' students');
    return attendanceSheet;

  } catch (error) {
    UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - ❌ Error generating attendance sheet: ' + error.message);
    throw error;
  }
}

function generateTeacherDisplayNames(lookupSheet) {
  try {
    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var firstNameCol = getCol('First Name');
    var lastNameCol = getCol('Last Name');
    var teacherIdCol = getCol('Teacher ID');
    var statusCol = getCol('Status');

    if (!firstNameCol || !lastNameCol || !teacherIdCol || !statusCol) {
      UtilityScriptLibrary.debugLog('generateTeacherDisplayNames', 'ERROR', 'Required columns not found', '', '');
      return {};
    }

    var validStatuses = ['potential', 'active', 'returning'];
    var data = lookupSheet.getDataRange().getValues();
    var teachers = [];

    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][statusCol - 1]).trim().toLowerCase();
      if (validStatuses.indexOf(status) === -1) continue;
      var firstName = String(data[i][firstNameCol - 1]).trim();
      var lastName = String(data[i][lastNameCol - 1]).trim();
      var teacherId = String(data[i][teacherIdCol - 1]).trim();
      if (!firstName || !lastName || !teacherId) continue;
      teachers.push({ firstName: firstName, lastName: lastName, teacherId: teacherId, row: i + 1 });
    }

    // Count last name occurrences
    var lastNameCounts = {};
    for (var i = 0; i < teachers.length; i++) {
      var ln = teachers[i].lastName;
      lastNameCounts[ln] = (lastNameCounts[ln] || 0) + 1;
    }

    // Generate display names with collision detection
    var displayNameMap = {}; // teacherId -> displayName

    // First pass: assign last-name-only where no collision
    var collisionGroups = {}; // lastName -> [teacher, ...]
    for (var i = 0; i < teachers.length; i++) {
      var t = teachers[i];
      if (lastNameCounts[t.lastName] === 1) {
        displayNameMap[t.teacherId] = t.lastName;
      } else {
        if (!collisionGroups[t.lastName]) collisionGroups[t.lastName] = [];
        collisionGroups[t.lastName].push(t);
      }
    }

    // Second pass: resolve collisions by appending minimum first name chars
    for (var lastName in collisionGroups) {
      var group = collisionGroups[lastName];
      var charsNeeded = 1;
      var resolved = false;

      while (!resolved && charsNeeded <= group[0].firstName.length) {
        var candidates = {};
        var collision = false;

        for (var i = 0; i < group.length; i++) {
          var suffix = group[i].firstName.substring(0, charsNeeded);
          if (candidates[suffix]) {
            collision = true;
            break;
          }
          candidates[suffix] = true;
        }

        if (!collision) {
          resolved = true;
          for (var i = 0; i < group.length; i++) {
            displayNameMap[group[i].teacherId] = group[i].lastName + group[i].firstName.substring(0, charsNeeded);
          }
        } else {
          charsNeeded++;
        }
      }

      // Fallback: full first name still collides (identical names), append teacher ID
      if (!resolved) {
        for (var i = 0; i < group.length; i++) {
          displayNameMap[group[i].teacherId] = group[i].lastName + group[i].firstName + group[i].teacherId;
        }
      }
    }

    return displayNameMap; // { teacherId: displayName, ... }

  } catch (error) {
    UtilityScriptLibrary.debugLog('generateTeacherDisplayNames', 'ERROR', 'Failed', '', error.message);
    return {};
  }
}

function getActiveStudentsFromRoster(rosterSheet) {
  try {
    UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "DEBUG", "Starting to extract students from roster", "Sheet: " + rosterSheet.getName(), "");
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(rosterSheet);
    
    // DEBUG: Log all available headers
    var allHeaders = [];
    for (var key in headerMap) {
      allHeaders.push(key + "=" + headerMap[key]);
    }
    UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "DEBUG", "Available headers in roster", allHeaders.join(", "), "");
    
    var studentIdCol = headerMap['student id'] || headerMap['studentid'];
    var statusCol = headerMap['status'];
    var firstNameCol = headerMap['first name'] || headerMap['firstname'];
    var lastNameCol = headerMap['last name'] || headerMap['lastname'];
    var instrumentCol = headerMap['instrument'];
    
    if (!studentIdCol) {
      UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "ERROR", "Student ID column not found", "Available headers: " + allHeaders.join(", "), "");
      throw new Error('Student ID column not found in roster');
    }
    
    if (!firstNameCol || !lastNameCol || !instrumentCol) {
      UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "ERROR", "Required columns missing", 
        "FirstName: " + firstNameCol + ", LastName: " + lastNameCol + ", Instrument: " + instrumentCol, "");
      throw new Error('Required columns (First Name, Last Name, or Instrument) not found in roster');
    }
    
    UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "DEBUG", "Column positions", 
      "StudentID: " + studentIdCol + ", Status: " + statusCol + ", FirstName: " + firstNameCol + 
      ", LastName: " + lastNameCol + ", Instrument: " + instrumentCol, "");
    
    var data = rosterSheet.getDataRange().getValues();
    var students = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentId = row[studentIdCol - 1];
      var status = statusCol ? row[statusCol - 1] : '';
      
      // Include students with IDs and Active or Carryover status (or empty status)
      if (studentId && String(studentId).trim() !== '' && 
          (!status || String(status).trim() === '' || 
           String(status).trim().toLowerCase() === 'active' ||
           String(status).trim().toLowerCase() === 'carryover')) {
        
        var studentInfo = extractStudentDataFromRoster(row, headerMap);
        studentInfo.rowNumber = i + 1; // Store row number for later updates
        students.push(studentInfo);
      }
    }
    
    UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "INFO", "Extracted students from roster", 
      "Found " + students.length + " active students", "");
    return students;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getActiveStudentsFromRoster", "ERROR", "Failed to extract students", "", error.message);
    throw error;
  }
}

function getActiveTeachersForDropdown() {
  try {
    var lookupSheet = getTeacherRosterLookupSheet();

    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'WARNING', 'Teacher Roster Lookup sheet not found or empty', '', '');
      return [];
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var teacherIdCol = getCol('Teacher ID');
    var statusCol = getCol('Status');

    if (!teacherIdCol || !statusCol) {
      UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'ERROR', 'Required columns not found', '', '');
      return [];
    }

    var displayNameMap = generateTeacherDisplayNames(lookupSheet);

    if (Object.keys(displayNameMap).length === 0) {
      UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'WARNING', 'No display names generated', '', '');
      return [];
    }

    var validStatuses = ['potential', 'active', 'returning'];
    var displayNames = [];
    var data = lookupSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var teacherId = String(data[i][teacherIdCol - 1]).trim();
      var status = String(data[i][statusCol - 1]).trim().toLowerCase();
      if (validStatuses.indexOf(status) === -1) continue;
      if (!displayNameMap[teacherId]) continue;
      displayNames.push(displayNameMap[teacherId]);
    }

    displayNames.sort();

    UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'INFO', 'Found teachers for dropdown', displayNames.length + ' teachers', '');
    return displayNames;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'ERROR', 'Failed', '', error.message);
    return [];
  }
}

function getAllTeachersWithGroupAssignments() {
  try {
    var lookupSheet = getTeacherRosterLookupSheet();

    if (!lookupSheet) {
      UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'WARNING', 'Teacher Roster Lookup sheet not found', '', '');
      return [];
    }

    var data = lookupSheet.getDataRange().getValues();
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'INFO', 'No teachers found in lookup', '', '');
      return [];
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var firstNameCol = getCol('First Name');
    var lastNameCol = getCol('Last Name');
    var groupAssignmentCol = getCol('Group Assignment');

    if (!firstNameCol || !lastNameCol || !groupAssignmentCol) {
      UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'ERROR', 'Required columns not found', '', '');
      return [];
    }

    var teachers = [];
    for (var i = 1; i < data.length; i++) {
      var firstName = String(data[i][firstNameCol - 1]).trim();
      var lastName = String(data[i][lastNameCol - 1]).trim();
      var groupAssignment = String(data[i][groupAssignmentCol - 1]).trim();

      if (firstName && lastName && groupAssignment) {
        teachers.push({ firstName: firstName, lastName: lastName });
      }
    }

    UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'INFO', 'Found teachers with groups', 'Count: ' + teachers.length, '');
    return teachers;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'ERROR', 'Failed', '', error.message);
    return [];
  }
}

function getContinuingStudentsFromWorkbook(workbook) {
  try {
    var rosterSheet = findMostRecentRosterSheet(workbook);
    if (!rosterSheet) {
      UtilityScriptLibrary.debugLog('getContinuingStudentsFromWorkbook - No roster sheet found');
      return [];
    }

    UtilityScriptLibrary.debugLog('getContinuingStudentsFromWorkbook - Using roster sheet: ' + rosterSheet.getName());

    var data = rosterSheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var headerMap = UtilityScriptLibrary.getHeaderMap(rosterSheet);

    var lastNameCol      = headerMap['lastname'];
    var firstNameCol     = headerMap['firstname'];
    var statusCol        = headerMap['status'];
    var lessonsCol       = headerMap['lessonsremaining'];
    var instrumentCol    = headerMap['instrument'];
    var lengthCol        = headerMap['length'];
    var experienceCol    = headerMap['experience'];
    var gradeCol         = headerMap['grade'];
    var schoolCol        = headerMap['school'];
    var schoolTeacherCol = headerMap['schoolteacher'];
    var parentLastCol    = headerMap['parentlastname'];
    var parentFirstCol   = headerMap['parentfirstname'];
    var phoneCol         = headerMap['phone'];
    var emailCol         = headerMap['email'];
    var addlContactsCol  = headerMap['additionalcontacts'];
    var studentIdCol     = headerMap['studentid'];

    if (!lastNameCol || !firstNameCol || !statusCol || !lessonsCol || !studentIdCol) {
      throw new Error('Required columns not found in roster: ' + rosterSheet.getName());
    }

    var continuingStudents = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[lastNameCol - 1] || !row[firstNameCol - 1]) continue;

      var status = (row[statusCol - 1] || '').toString().trim().toLowerCase();
      var lessonsRemaining = parseFloat(row[lessonsCol - 1]) || 0;

      if (lessonsRemaining > 0 && (status === 'active' || status === 'carryover')) {
        continuingStudents.push({
          lastName:         row[lastNameCol - 1]      || '',
          firstName:        row[firstNameCol - 1]     || '',
          instrument:       row[instrumentCol - 1]    || '',
          length:           row[lengthCol - 1]        || 30,
          experience:       row[experienceCol - 1]    || '',
          grade:            row[gradeCol - 1]         || '',
          school:           row[schoolCol - 1]        || '',
          schoolTeacher:    row[schoolTeacherCol - 1] || '',
          parentLastName:   row[parentLastCol - 1]    || '',
          parentFirstName:  row[parentFirstCol - 1]   || '',
          phone:            row[phoneCol - 1]         || '',
          email:            row[emailCol - 1]         || '',
          additionalContacts: row[addlContactsCol - 1] || '',
          studentId:        row[studentIdCol - 1]     || '',
          lessonsRemaining: lessonsRemaining
        });
      }
    }

    return continuingStudents;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getContinuingStudentsFromWorkbook - Error: ' + error.message);
    return [];
  }
}

function getExistingGroupIds(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var studentIdCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (UtilityScriptLibrary.normalizeHeader(headers[i]) === 'studentid') {
        studentIdCol = i;
        break;
      }
    }
    if (studentIdCol === -1) {
      UtilityScriptLibrary.debugLog('getExistingGroupIds', 'ERROR', 'Student ID column not found', '', '');
      return [];
    }
    var existingGroupIds = [];
    for (var i = 1; i < data.length; i++) {
      var studentId = data[i][studentIdCol];
      if (studentId && String(studentId).match(/^G\d{4}$/)) {
        var groupIdStr = String(studentId).trim();
        var alreadyInList = false;
        for (var j = 0; j < existingGroupIds.length; j++) {
          if (existingGroupIds[j] === groupIdStr) {
            alreadyInList = true;
            break;
          }
        }
        if (!alreadyInList) {
          existingGroupIds.push(groupIdStr);
        }
      }
    }
    UtilityScriptLibrary.debugLog('getExistingGroupIds', 'DEBUG', 'Found existing group IDs', existingGroupIds.join(', '), '');
    return existingGroupIds;
  } catch (error) {
    UtilityScriptLibrary.debugLog('getExistingGroupIds', 'ERROR', 'Error getting existing group IDs', '', error.message);
    return [];
  }
}

function getMostRecentMonthSheet(workbook) {
  try {
    var sheets = workbook.getSheets();
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    var foundMonthSheets = [];
    
    // Find all sheets that match month names
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      for (var j = 0; j < monthNames.length; j++) {
        if (sheetName.toLowerCase() === monthNames[j].toLowerCase()) {
          foundMonthSheets.push({
            sheet: sheets[i],
            monthIndex: j,
            sheetName: sheetName
          });
          break;
        }
      }
    }
    
    if (foundMonthSheets.length === 0) {
      UtilityScriptLibrary.debugLog('getMostRecentMonthSheet', 'WARN', 'No month sheets found', '', '');
      return null;
    }
    
    // Sort by month order and return the last one (most recent chronologically)
    foundMonthSheets.sort(function(a, b) {
      return a.monthIndex - b.monthIndex;
    });
    
    var mostRecentMonthSheet = foundMonthSheets[foundMonthSheets.length - 1];
    
    UtilityScriptLibrary.debugLog('getMostRecentMonthSheet', 'INFO', 'Found most recent month sheet', 
                                 'Sheet: ' + mostRecentMonthSheet.sheetName, '');
    
    return mostRecentMonthSheet.sheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('getMostRecentMonthSheet', 'ERROR', 'Error finding most recent month sheet', '', error.message);
    return null;
  }
}

function getMostRecentMonthSheets(workbook) {
  try {
    var allSheets = workbook.getSheets();
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var monthSheets = [];
    
    for (var i = 0; i < allSheets.length; i++) {
      var sheetName = allSheets[i].getName();
      var monthIndex = monthNames.indexOf(sheetName);
      if (monthIndex !== -1) {
        monthSheets.push({
          sheet: allSheets[i],
          name: sheetName,
          monthIndex: monthIndex
        });
      }
    }
    
    if (monthSheets.length === 0) {
      UtilityScriptLibrary.debugLog("getMostRecentMonthSheets", "INFO", "No month sheets found", "", "");
      return { mostRecent: null, secondMostRecent: null };
    }
    
    monthSheets.sort(function(a, b) {
      return a.monthIndex - b.monthIndex;
    });
    
    var mostRecent = monthSheets[monthSheets.length - 1];
    var secondMostRecent = monthSheets.length >= 2 ? monthSheets[monthSheets.length - 2] : null;
    
    UtilityScriptLibrary.debugLog("getMostRecentMonthSheets", "INFO", "Found month sheets", 
                          "Most recent: " + mostRecent.name + 
                          (secondMostRecent ? ", Second: " + secondMostRecent.name : ""), 
                          "Total: " + monthSheets.length);
    
    return {
      mostRecent: mostRecent,
      secondMostRecent: secondMostRecent
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getMostRecentMonthSheets", "ERROR", "Error finding month sheets", "", error.message);
    return { mostRecent: null, secondMostRecent: null };
  }
}

function getOrCreateRosterFromTemplate(teacherInfo, rosterFolder, year, semesterName, registrationTimestamp) {
  try {
    if (!teacherInfo || !teacherInfo.teacherId) {
      throw new Error('Invalid teacherInfo object — missing Teacher ID');
    }

    var teacherFullName = teacherInfo.firstName + ' ' + teacherInfo.lastName;
    var fileName = 'QAMP ' + year + ' ' + teacherInfo.lastName + ' ' + teacherInfo.teacherId;

    UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'INFO', 'Processing roster for teacher', teacherFullName, '');

    var files = rosterFolder.getFilesByName(fileName);

    if (files.hasNext()) {
      var existingFile = files.next();
      UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'INFO', 'Roster file already exists', fileName, '');
      return SpreadsheetApp.openById(existingFile.getId());
    }

    var newSpreadsheet = SpreadsheetApp.create(fileName);
    var newFile = DriveApp.getFileById(newSpreadsheet.getId());
    rosterFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);

    UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'INFO', 'Created new spreadsheet', fileName, '');

    try {
      setupCompleteRosterWorkbook(newSpreadsheet, teacherFullName, year, semesterName, registrationTimestamp);
      UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'INFO', 'Roster workbook setup complete', fileName, '');
    } catch (setupError) {
      UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'WARNING', 'Roster workbook setup failed — file still created', fileName, setupError.message);
    }

    try {
      updateTeacherRosterLookup(teacherInfo.teacherId, newFile.getUrl());
      UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'INFO', 'Updated teacher roster lookup', teacherInfo.teacherId, '');
    } catch (lookupError) {
      UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'WARNING', 'Could not update teacher roster lookup', teacherInfo.teacherId, lookupError.message);
    }

    return newSpreadsheet;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getOrCreateRosterFromTemplate', 'ERROR', 'Failed', '', error.message);
    throw error;
  }
}

function getTeacherInfoByDisplayName(displayName) {
  try {
    var lookupSheet = getTeacherRosterLookupSheet();

    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog('getTeacherInfoByDisplayName', 'WARNING', 'Teacher Roster Lookup sheet not found or empty', '', '');
      return null;
    }

    var displayNameMap = generateTeacherDisplayNames(lookupSheet);

    // Reverse lookup: displayName -> teacherId
    var resolvedTeacherId = null;
    for (var teacherId in displayNameMap) {
      if (displayNameMap[teacherId] === displayName) {
        resolvedTeacherId = teacherId;
        break;
      }
    }

    if (!resolvedTeacherId) {
      UtilityScriptLibrary.debugLog('getTeacherInfoByDisplayName', 'WARNING', 'Teacher not found', displayName, '');
      return null;
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var firstNameCol = getCol('First Name');
    var lastNameCol = getCol('Last Name');
    var rosterUrlCol = getCol('Roster URL');
    var teacherIdCol = getCol('Teacher ID');
    var statusCol = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');

    if (!teacherIdCol) {
      UtilityScriptLibrary.debugLog('getTeacherInfoByDisplayName', 'ERROR', 'Required columns not found', '', '');
      return null;
    }

    var data = lookupSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][teacherIdCol - 1]).trim() !== resolvedTeacherId) continue;

      return {
        firstName:   firstNameCol   ? String(data[i][firstNameCol - 1]).trim()  : '',
        lastName:    lastNameCol     ? String(data[i][lastNameCol - 1]).trim()   : '',
        rosterUrl:   rosterUrlCol    ? String(data[i][rosterUrlCol - 1]).trim()  : '',
        teacherId:   resolvedTeacherId,
        status:      statusCol       ? String(data[i][statusCol - 1]).trim()     : '',
        lastUpdated: lastUpdatedCol  ? data[i][lastUpdatedCol - 1]               : ''
      };
    }

    UtilityScriptLibrary.debugLog('getTeacherInfoByDisplayName', 'WARNING', 'Teacher ID not found in data rows', resolvedTeacherId, '');
    return null;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getTeacherInfoByDisplayName', 'ERROR', 'Failed', displayName, error.message);
    return null;
  }
}

function getTeacherInfoByFullName(firstName, lastName) {
  try {
    var lookupSheet = getTeacherRosterLookupSheet();

    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog('getTeacherInfoByFullName', 'WARNING', 'Teacher Roster Lookup sheet not found or empty', '', '');
      return null;
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var firstNameCol = getCol('First Name');
    var lastNameCol = getCol('Last Name');
    var rosterUrlCol = getCol('Roster URL');
    var teacherIdCol = getCol('Teacher ID');
    var statusCol = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');

    if (!firstNameCol || !lastNameCol) {
      UtilityScriptLibrary.debugLog('getTeacherInfoByFullName', 'ERROR', 'Required columns not found', '', '');
      return null;
    }

    var data = lookupSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowFirstName = String(data[i][firstNameCol - 1]).trim();
      var rowLastName = String(data[i][lastNameCol - 1]).trim();
      if (rowFirstName.toLowerCase() !== firstName.trim().toLowerCase()) continue;
      if (rowLastName.toLowerCase() !== lastName.trim().toLowerCase()) continue;

      return {
        firstName:   rowFirstName,
        lastName:    rowLastName,
        rosterUrl:   rosterUrlCol   ? String(data[i][rosterUrlCol - 1]).trim() : '',
        teacherId:   teacherIdCol   ? String(data[i][teacherIdCol - 1]).trim() : '',
        status:      statusCol      ? String(data[i][statusCol - 1]).trim()    : '',
        lastUpdated: lastUpdatedCol ? data[i][lastUpdatedCol - 1]              : ''
      };
    }

    UtilityScriptLibrary.debugLog('getTeacherInfoByFullName', 'WARNING', 'Teacher not found', firstName + ' ' + lastName, '');
    return null;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getTeacherInfoByFullName', 'ERROR', 'Failed', firstName + ' ' + lastName, error.message);
    return null;
  }
}

function getTeacherRosterLookupSheet() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return activeSpreadsheet.getSheetByName("Teacher Roster Lookup");
}

function getYearRosterFolders(previousYear, newYear) {
  var rostersFolder = UtilityScriptLibrary.getRosterFolder();
  var previousFolder = null;
  var nextFolder = null;

  var subfolders = rostersFolder.getFolders();
  while (subfolders.hasNext()) {
    var folder = subfolders.next();
    var name = folder.getName();
    if (name === previousYear + ' Rosters') previousFolder = folder;
    if (name === newYear + ' Rosters') nextFolder = folder;
  }

  if (!previousFolder) {
    throw new Error('Previous year folder not found: ' + previousYear + ' Rosters');
  }
  if (!nextFolder) {
    throw new Error('New year folder not found: ' + newYear + ' Rosters\n\nPlease run semester setup in Billing first.');
  }

  UtilityScriptLibrary.debugLog('getYearRosterFolders - Found folders for ' + previousYear + ' and ' + newYear);
  return { previous: previousFolder, next: nextFolder };
}

function hasMonthBeenInvoiced(sheet) {
  try {
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var invoiceDateCol = headerMap['invoicedate'];
    var invoiceNumberCol = headerMap['invoicenumber'];
    
    if (!invoiceDateCol && !invoiceNumberCol) {
      return false;
    }
    
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      var invoiceDate = invoiceDateCol ? data[i][invoiceDateCol - 1] : null;
      var invoiceNumber = invoiceNumberCol ? data[i][invoiceNumberCol - 1] : null;
      
      if ((invoiceDate && invoiceDate.toString().trim() !== '') ||
          (invoiceNumber && invoiceNumber.toString().trim() !== '')) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("hasMonthBeenInvoiced", "ERROR", "Error checking invoice status", 
                          "Sheet: " + sheet.getName(), error.message);
    return false;
  }
}

function loadStudentMapFromContacts() {
  try {
    var contactsSheet = UtilityScriptLibrary.getSheet('students');
    var data = contactsSheet.getDataRange().getValues();
    var headers = data[0];
    
    var getCol = UtilityScriptLibrary.createColumnFinder(contactsSheet);
    var idCol = getCol('Student ID');
    var firstNameCol = getCol('Student First Name');
    var lastNameCol = getCol('Student Last Name');
    
    if (!idCol || !firstNameCol || !lastNameCol) {
      throw new Error('Required columns not found in Contacts students sheet');
    }
    
    var studentMap = {};
    for (var i = 1; i < data.length; i++) {
      var id = (data[i][idCol - 1] || '').toString().trim();
      var firstName = (data[i][firstNameCol - 1] || '').toString().trim().toLowerCase();
      var lastName = (data[i][lastNameCol - 1] || '').toString().trim().toLowerCase();
      
      if (id && id.charAt(0) === 'Q' && firstName && lastName) {
        var key = firstName + '|' + lastName;
        studentMap[key] = id;
      }
    }
    
    Logger.log('📚 Loaded ' + Object.keys(studentMap).length + ' students from Contacts');
    return studentMap;
    
  } catch (error) {
    Logger.log('❌ Error loading student map: ' + error.message);
    throw error;
  }
}

function populateRosterWithContinuingStudents(workbook, semesterName, students) {
  try {
    var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
    var rosterSheetName = season + ' Roster';
    var rosterSheet = workbook.getSheetByName(rosterSheetName);

    if (!rosterSheet) {
      throw new Error('Roster sheet not found: ' + rosterSheetName);
    }

    // Sort students alphabetically by last name then first name
    students.sort(function(a, b) {
      var lastNameCompare = (a.lastName || '').localeCompare(b.lastName || '');
      if (lastNameCompare !== 0) return lastNameCompare;
      return (a.firstName || '').localeCompare(b.firstName || '');
    });

    var dataRows = [];
    for (var i = 0; i < students.length; i++) {
      var s = students[i];
      dataRows.push([
        false,                               // A - Contacted checkbox
        '',                                  // B - First Lesson Date
        '',                                  // C - First Lesson Time
        '',                                  // D - Comments
        s.lastName,                          // E - Last Name
        s.firstName,                         // F - First Name
        s.instrument,                        // G - Instrument
        s.length,                            // H - Length
        s.experience,                        // I - Experience
        s.grade,                             // J - Grade
        s.school,                            // K - School
        s.schoolTeacher,                     // L - School Teacher
        s.parentLastName,                    // M - Parent Last Name
        s.parentFirstName,                   // N - Parent First Name
        s.phone,                             // O - Phone
        s.email,                             // P - Email
        s.additionalContacts,                // Q - Additional contacts
        0,                                   // R - Hours Remaining (reset; updated by sync)
        s.lessonsRemaining,                  // S - Lessons Remaining (carry forward)
        'Carryover',                         // T - Status
        s.studentId,                         // U - Student ID
        '',                                  // V - Admin Comments
        'Carried over from previous year'    // W - System Comments
      ]);
    }

    if (dataRows.length > 0) {
      rosterSheet.getRange(2, 1, dataRows.length, 23).setValues(dataRows);
      // Insert checkboxes for Contacted column (A)
      rosterSheet.getRange(2, 1, dataRows.length, 1).insertCheckboxes();
      UtilityScriptLibrary.debugLog('✅ Populated roster with ' + students.length + ' continuing students');
    }

  } catch (error) {
    UtilityScriptLibrary.debugLog('Error populating roster: ' + error.message);
    throw error;
  }
}

function processParent(formData, parentsSheet, studentId, existingParentId) {
  try {
    UtilityScriptLibrary.debugLog("👨‍👩‍👧‍👦 Starting processParent");

    existingParentId = existingParentId || '';

    var headers = parentsSheet.getRange(1, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
    var getCol = function(name) {
      for (var i = 0; i < headers.length; i++) {
        if (UtilityScriptLibrary.normalizeHeader(String(headers[i])) === UtilityScriptLibrary.normalizeHeader(name)) {
          return i + 1;
        }
      }
      return 0;
    };

    if (formData["Phone"]) {
      formData["Phone"] = UtilityScriptLibrary.formatPhoneNumber(String(formData["Phone"]));
    }

    var cityZipRaw = formData["CityZip"];
    if (cityZipRaw) {
      var parsed = UtilityScriptLibrary.parseCityZipMessy(cityZipRaw);
      formData["Address City"] = parsed.city;
      formData["Address Zip Code"] = parsed.zip;
      formData["Address Formatted"] = UtilityScriptLibrary.formatAddress(formData["Address Street"] || '', parsed.city, parsed.zip);
    }

    var parentKey = UtilityScriptLibrary.generateKey(
      formData["Parent Last Name"] || '',
      formData["Parent First Name"] || '',
      formData["Email"] || ''
    );

    var parentIdCol = getCol("Parent ID");
    var lookupCol = getCol("Parent Lookup");
    var studentIdsCol = getCol("Student IDs");
    var updatedCol = getCol("Updated");
    var parentGroupCol = getCol("Parent Group Interest");

    var textFields = [
      "Parent Last Name", "Parent First Name", "Salutation", "Email", "Email 2",
      "Phone", "Address Formatted", "Billing Preference", "Additional Contacts", "Referral"
    ];

    var parentId = existingParentId;
    var parentRow = UtilityScriptLibrary.findParentRow(parentsSheet, parentId, parentKey);

    UtilityScriptLibrary.debugLog("=== PARENT DUPLICATE CHECK ===");
    UtilityScriptLibrary.debugLog("Looking for parentId: '" + parentId + "' and parentKey: '" + parentKey + "'");
    UtilityScriptLibrary.debugLog("findParentRow result: " + parentRow);
    UtilityScriptLibrary.debugLog("=== END PARENT DEBUG ===");

    if (parentRow !== -1) {
      UtilityScriptLibrary.debugLog("📄 Updating existing parent");

      // Build fields object from formData
      var fieldsObj = {};
      for (var j = 0; j < textFields.length; j++) {
        fieldsObj[textFields[j]] = formData[textFields[j]] || '';
      }

      var updateResult = UtilityScriptLibrary.updateParentContactFields(
        parentsSheet, parentRow, fieldsObj, { newLookupKey: parentKey }
      );
      var changesMade = updateResult.changesMade;

      // Update Student IDs list (add if not already present)
      if (studentIdsCol) {
        var rowValues = parentsSheet.getRange(parentRow, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
        var currentStudentIds = String(rowValues[studentIdsCol - 1] || '');
        var studentIdArray = currentStudentIds ? currentStudentIds.split(',').map(function(id) { return id.trim(); }) : [];
        if (studentIdArray.indexOf(studentId) === -1) {
          studentIdArray.push(studentId);
          parentsSheet.getRange(parentRow, studentIdsCol).setValue(studentIdArray.join(', '));
          UtilityScriptLibrary.debugLog("Added student ID " + studentId + " to parent's student list");
        }
      }

      // Update Parent Group Interest checkbox
      if (parentGroupCol) {
        var parentGroupCell = parentsSheet.getRange(parentRow, parentGroupCol);
        parentGroupCell.insertCheckboxes();
        parentGroupCell.setValue(formData["Parent Group Interest"] === "Yes");
      }

    } else {
      UtilityScriptLibrary.debugLog("➕ Creating new parent");
      parentId = UtilityScriptLibrary.generateNextId(parentsSheet, "Parent ID", "P");

      var newRow = new Array(headers.length);
      for (var m = 0; m < headers.length; m++) {
        newRow[m] = '';
      }

      newRow[parentIdCol - 1] = parentId;
      newRow[lookupCol - 1] = parentKey;
      newRow[studentIdsCol - 1] = studentId;

      for (var n = 0; n < textFields.length; n++) {
        var field = textFields[n];
        var col = getCol(field);
        if (col) {
          newRow[col - 1] = formData[field] || '';
        }
      }

      parentsSheet.appendRow(newRow);
      var newRowIndex = parentsSheet.getLastRow();

      if (parentGroupCol) {
        var parentGroupCell = parentsSheet.getRange(newRowIndex, parentGroupCol);
        parentGroupCell.insertCheckboxes();
        parentGroupCell.setValue(formData["Parent Group Interest"] === "Yes");
      }

      if (updatedCol) {
        var updatedCell = parentsSheet.getRange(newRowIndex, updatedCol);
        updatedCell.insertCheckboxes();
        updatedCell.setValue(true);
      }
    }

    UtilityScriptLibrary.debugLog("✅ Completed processParent - ID: " + parentId);
    return parentId;

  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in processParent: " + error.message);
    throw error;
  }
}

function processPendingAssignments() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var calendarSheet = ss.getSheetByName('Calendar');
    
    if (!calendarSheet) {
      SpreadsheetApp.getUi().alert('Calendar sheet not found');
      return;
    }
    
    var currentSemester = calendarSheet.getRange(2, 4).getValue();
    if (!currentSemester || String(currentSemester).trim() === '') {
      SpreadsheetApp.getUi().alert('No current semester defined in Calendar');
      return;
    }
    
    var sheet = ss.getSheetByName(String(currentSemester).trim());
    if (!sheet) {
      SpreadsheetApp.getUi().alert('Current semester sheet not found: ' + currentSemester);
      return;
    }
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var teacherCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher')];
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    
    if (!teacherCol || !studentIdCol) {
      SpreadsheetApp.getUi().alert('Required columns not found');
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    var pendingRows = [];
    
    // Find all rows that need processing
    for (var i = 1; i < data.length; i++) {
      var teacher = data[i][teacherCol - 1];
      var studentId = data[i][studentIdCol - 1];
      
      if (teacher && String(teacher).trim() !== '' && 
          (!studentId || String(studentId).trim() === '')) {
        pendingRows.push(i + 1);
      }
    }
    
    if (pendingRows.length === 0) {
      SpreadsheetApp.getUi().alert('No pending assignments found');
      return;
    }
    
    // Confirm before processing
    var response = SpreadsheetApp.getUi().alert(
      'Found ' + pendingRows.length + ' pending assignment(s). Process them now?\n\n' +
      'This may take several minutes.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
    
    // Process each row
    var successCount = 0;
    var errorCount = 0;
    var errors = [];
    
    for (var i = 0; i < pendingRows.length; i++) {
      try {
        processSingleRow(sheet, pendingRows[i], headerMap);
        successCount++;
        UtilityScriptLibrary.debugLog('✅ Processed row ' + pendingRows[i]);
      } catch (error) {
        errorCount++;
        errors.push('Row ' + pendingRows[i] + ': ' + error.message);
        UtilityScriptLibrary.debugLog('❌ Error on row ' + pendingRows[i] + ': ' + error.message);
      }
    }
    
    // Show results
    var message = 'Batch Processing Complete\n\n';
    message += 'Successfully processed: ' + successCount + '\n';
    if (errorCount > 0) {
      message += 'Errors: ' + errorCount + '\n\n';
      message += 'Failed rows:\n' + errors.join('\n');
    }
    
    SpreadsheetApp.getUi().alert(message);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    UtilityScriptLibrary.debugLog('❌ Batch processing error: ' + error.message);
  }
}

function processReassignment() {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();
    
    // Retrieve all stored data
    var currentSemester = scriptProps.getProperty('reassign_currentSemester');
    var year = scriptProps.getProperty('reassign_year');
    var season = scriptProps.getProperty('reassign_season');
    var oldTeacherDisplay = scriptProps.getProperty('reassign_oldTeacherDisplay');
    var newTeacherDisplay = scriptProps.getProperty('reassign_newTeacherDisplay');
    var oldTeacherInfo = JSON.parse(scriptProps.getProperty('reassign_oldTeacherInfo'));
    var selectedStudents = JSON.parse(scriptProps.getProperty('reassign_selectedStudents'));
    var effectiveDate = new Date(scriptProps.getProperty('reassign_effectiveDate'));
    var oldRosterWorkbookId = scriptProps.getProperty('reassign_oldRosterWorkbookId');
    var oldRosterSheetName = scriptProps.getProperty('reassign_oldRosterSheetName');
    
    UtilityScriptLibrary.debugLog("Processing reassignment: " + selectedStudents.length + " students from " + oldTeacherDisplay + " to " + newTeacherDisplay);
    
    // Get roster folder from Year Metadata
    var yearMetadataSheet = UtilityScriptLibrary.getSheet('yearMetadata');
    if (!yearMetadataSheet) {
      throw new Error("Year Metadata sheet not found");
    }
    
    var metadataRows = yearMetadataSheet.getDataRange().getValues();
    var headerRow = metadataRows[0];
    var yearColIndex = -1;
    var folderIdColIndex = -1;
    
    for (var i = 0; i < headerRow.length; i++) {
      if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader("Year")) {
        yearColIndex = i;
      }
      if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader("Roster Folder ID")) {
        folderIdColIndex = i;
      }
    }
    
    if (yearColIndex === -1 || folderIdColIndex === -1) {
      throw new Error("Required columns not found in Year Metadata sheet");
    }
    
    var yearRow = null;
    for (var i = 0; i < metadataRows.length; i++) {
      if (metadataRows[i][yearColIndex] && metadataRows[i][yearColIndex].toString() === year) {
        yearRow = metadataRows[i];
        break;
      }
    }
    
    if (!yearRow) {
      throw new Error("No roster folder found for year: " + year);
    }
    
    var rosterFolderId = yearRow[folderIdColIndex];
    var rosterFolder = DriveApp.getFolderById(rosterFolderId);
    UtilityScriptLibrary.debugLog("✅ Found roster folder for year " + year);
    
    // Get new teacher info
    var newTeacherInfo = getTeacherInfoByDisplayName(newTeacherDisplay);
    if (!newTeacherInfo) {
      throw new Error("Could not find new teacher info in Teacher Roster Lookup");
    }
    
    var newRosterWorkbook = getOrCreateRosterFromTemplate(newTeacherInfo, rosterFolder, year, currentSemester);
    var newRosterSheet = findSemesterRoster(newRosterWorkbook, currentSemester);
    
    if (!newRosterSheet) {
      throw new Error("Could not find or create roster sheet for " + currentSemester + " in new teacher's workbook");
    }
    
    UtilityScriptLibrary.debugLog("✅ Got/created new teacher roster: " + newRosterSheet.getName());
    
    // Update Teacher Roster Lookup with URL
    var newRosterFile = DriveApp.getFileById(newRosterWorkbook.getId());
    var newRosterUrl = newRosterFile.getUrl();
    
    UtilityScriptLibrary.debugLog("📝 Updating Teacher Roster Lookup for display name: " + newTeacherInfo.displayName);
    updateTeacherRosterLookup(newTeacherInfo.teacherId, newRosterUrl);
    UtilityScriptLibrary.debugLog("✅ Teacher Roster Lookup update completed");
    
    // Open old teacher's roster
    var oldRosterWorkbook = SpreadsheetApp.openById(oldRosterWorkbookId);
    var oldRosterSheet = oldRosterWorkbook.getSheetByName(oldRosterSheetName);
    
    if (!oldRosterSheet) {
      throw new Error("Could not find old roster sheet: " + oldRosterSheetName);
    }

    // Pre-fetch headers and data needed across the student loop
    var oldHeaderMap = UtilityScriptLibrary.getHeaderMap(oldRosterSheet);
    var newHeaderMap = UtilityScriptLibrary.getHeaderMap(newRosterSheet);

    var contactsSheet = UtilityScriptLibrary.getSheet('students');
    var contactsHeaderMap = contactsSheet ? UtilityScriptLibrary.getHeaderMap(contactsSheet) : null;
    var teacherCol = contactsHeaderMap ? contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Teacher')] : null;
    var studentIdCol = contactsHeaderMap ? contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Student ID')] : null;
    var contactsData = contactsSheet ? contactsSheet.getDataRange().getValues() : null;
    
    // Process each selected student
    var successCount = 0;
    var skipCount = 0;
    var errorCount = 0;
    var errors = [];
    
    for (var i = 0; i < selectedStudents.length; i++) {
      var student = selectedStudents[i];
      
      try {
        UtilityScriptLibrary.debugLog("Processing student: " + student.firstName + " " + student.lastName + " (" + student.studentId + ")");
        
        // Mark student as "Transferred" in old roster and clear warning color
        var statusCol = oldHeaderMap['status'];
        oldRosterSheet.getRange(student.rowNumber, statusCol).setValue('Transferred');
        
        // Apply normal alternating row style (background + text color) for columns A-W (23 columns)
        var rowRange = oldRosterSheet.getRange(student.rowNumber, 1, 1, 23);
        if (student.rowNumber % 2 === 0) {
          rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
          rowRange.setFontColor(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.text);
        } else {
          rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
          rowRange.setFontColor(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.text);
        }
        
        UtilityScriptLibrary.debugLog("Marked as 'Transferred' in old teacher's roster and cleared warning color");
        
        // Check if student already exists in new roster
        var studentExists = checkIfStudentExists(newRosterSheet, student.studentId, newHeaderMap);
        
        if (studentExists) {
          UtilityScriptLibrary.debugLog("Student already exists in new roster - skipping roster addition");
          skipCount++;
        } else {
          addStudentToRosterFromData(newRosterSheet, student, newHeaderMap);
          UtilityScriptLibrary.debugLog("Added to new teacher's roster");
        }
        
        // Add to attendance sheets
        addStudentToAttendanceSheetsFromDate(newRosterWorkbook, student, effectiveDate);

        // Update Teacher ID in Students (Contacts) sheet
        if (contactsSheet && teacherCol && studentIdCol && contactsData) {
          for (var r = 1; r < contactsData.length; r++) {
            if (String(contactsData[r][studentIdCol - 1]).trim() === String(student.studentId).trim()) {
              contactsSheet.getRange(r + 1, teacherCol).setValue(newTeacherInfo.teacherId);
              UtilityScriptLibrary.debugLog('processReassignment', 'INFO', 'Updated Teacher ID in Contacts',
                'Student: ' + student.studentId + ', Teacher ID: ' + newTeacherInfo.teacherId, '');
              break;
            }
          }
        }
        
        successCount++;
        UtilityScriptLibrary.debugLog("Successfully processed student: " + student.studentId);
        
      } catch (error) {
        errorCount++;
        var errorMsg = student.firstName + ' ' + student.lastName + ' (' + student.studentId + '): ' + error.message;
        errors.push(errorMsg);
        UtilityScriptLibrary.debugLog("ERROR processing student: " + errorMsg);
      }
    }
    
    // Clean up script properties
    var keysToDelete = [
      'reassign_currentSemester', 'reassign_year', 'reassign_season',
      'reassign_activeTeachers', 'reassign_oldTeacherDisplay', 'reassign_oldTeacherInfo',
      'reassign_studentList', 'reassign_oldRosterSheetName', 'reassign_oldRosterWorkbookId',
      'reassign_selectedStudents', 'reassign_newTeacherDisplay', 'reassign_effectiveDate'
    ];
    for (var i = 0; i < keysToDelete.length; i++) {
      scriptProps.deleteProperty(keysToDelete[i]);
    }
    
    UtilityScriptLibrary.debugLog("=== STUDENT REASSIGNMENT COMPLETE ===");
    
    // Show results
    var message = 'Student Reassignment Complete!\n\n';
    message += 'Successfully processed: ' + successCount + ' student(s)\n';
    if (skipCount > 0) {
      message += 'Already in new roster (skipped): ' + skipCount + '\n';
    }
    if (errorCount > 0) {
      message += 'Errors: ' + errorCount + '\n\n';
      message += 'Failed students:\n' + errors.join('\n');
    }
    
    ui.alert('Success', message, ui.ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in processReassignment: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 5 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function processRoster(formData, sheet, editedRow, headerMap, fieldMap, studentId, teacherId, rosterFolder, year, semesterName) {
  try {
    UtilityScriptLibrary.debugLog('processRoster', 'INFO', 'Starting', 'Teacher ID: ' + teacherId + ', Semester: ' + semesterName, '');

    if (!teacherId || teacherId.trim() === '') {
      UtilityScriptLibrary.debugLog('processRoster', 'WARNING', 'Skipping — missing Teacher ID', '', '');
      return;
    }

    if (!studentId || studentId.toString().trim() === '') {
      UtilityScriptLibrary.debugLog('processRoster', 'WARNING', 'Skipping — missing student ID', '', '');
      return;
    }

    var lookupSheet = getTeacherRosterLookupSheet();
    if (!lookupSheet) {
      throw new Error('Teacher Roster Lookup sheet not found');
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var teacherIdCol = getCol('Teacher ID');
    var firstNameCol = getCol('First Name');
    var lastNameCol = getCol('Last Name');

    if (!teacherIdCol || !firstNameCol || !lastNameCol) {
      throw new Error('Required columns not found in Teacher Roster Lookup');
    }

    var data = lookupSheet.getDataRange().getValues();
    var teacherInfo = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][teacherIdCol - 1]).trim() === teacherId) {
        teacherInfo = {
          teacherId: teacherId,
          firstName: String(data[i][firstNameCol - 1]).trim(),
          lastName:  String(data[i][lastNameCol - 1]).trim()
        };
        break;
      }
    }

    if (!teacherInfo) {
      throw new Error('Teacher ID not found in lookup sheet: ' + teacherId);
    }

    var registrationTimestamp = formData['Timestamp'] || formData['timestamp'] || formData['Registration Date'] || null;

    var rosterSS = getOrCreateRosterFromTemplate(teacherInfo, rosterFolder, year, semesterName, registrationTimestamp);
    if (!rosterSS) {
      throw new Error('Could not create or access roster workbook for Teacher ID: ' + teacherId);
    }

    addStudentToSemesterRoster(rosterSS, formData, studentId, semesterName, registrationTimestamp);

    UtilityScriptLibrary.debugLog('processRoster', 'INFO', 'Completed', 'Teacher ID: ' + teacherId + ', Student: ' + studentId, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('processRoster', 'ERROR', 'Failed', 'Teacher ID: ' + teacherId + ', Semester: ' + semesterName, error.message);
    throw error;
  }
}

function processSingleRow(sheet, row, headerMap) {
  UtilityScriptLibrary.debugLog("Edit meets criteria - PROCEEDING with processing");

  var fieldMapSheet = UtilityScriptLibrary.getSheet('fieldMap');
  if (!fieldMapSheet) throw new Error("FieldMap sheet not found.");

  var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
  var formData = extractFormData(sheet, row, headerMap, fieldMap);

  UtilityScriptLibrary.debugLog("Extracted formData: " + JSON.stringify(formData));

  var contactsSheet = UtilityScriptLibrary.getSheet('students');
  var studentResult = processStudent(formData, contactsSheet, sheet.getName());
  var studentId = studentResult.studentId;
  var parentId = studentResult.parentId;
  var studentRow = studentResult.studentRow;

  var idCol = headerMap[UtilityScriptLibrary.normalizeHeader("Student ID")];
  if (!idCol) {
    idCol = headerMap[UtilityScriptLibrary.normalizeHeader("ID")];
    UtilityScriptLibrary.debugLog("Fallback to 'ID' column. Column index: " + idCol);
  }
  if (idCol) {
    UtilityScriptLibrary.debugLog("Writing studentId: " + studentId + " to row: " + row + ", col: " + idCol);
    sheet.getRange(row, idCol).setValue(studentId);
  }

  var parentsSheet = UtilityScriptLibrary.getSheet('parents');
  if (!parentsSheet) throw new Error("Parents sheet not found.");

  var existingParentId = studentResult.parentId || '';
  UtilityScriptLibrary.debugLog("Using Parent ID from student record: '" + existingParentId + "'");

  parentId = processParent(formData, parentsSheet, studentId, existingParentId);
  updateStudentWithParentId(contactsSheet, studentRow, parentId);

  // === ROSTER PROCESSING ===
  try {
    var sheetName = sheet.getName();
    var yearMatch = sheetName.match(/\d{4}/);
    if (!yearMatch) {
      UtilityScriptLibrary.debugLog("⚠️ Could not extract year from sheet name: " + sheetName);
      return;
    }
    var year = yearMatch[0];
    var semesterName = sheetName;
    UtilityScriptLibrary.debugLog("Extracted year: " + year + ", semester: " + semesterName);

    var yearMetadataSheet = UtilityScriptLibrary.getSheet('yearMetadata');
    if (!yearMetadataSheet) {
      UtilityScriptLibrary.debugLog("⚠️ Year Metadata sheet not found.");
      return;
    }

    var metadataRows = yearMetadataSheet.getDataRange().getValues();
    var headerRow = metadataRows[0];
    var yearColIndex = -1;
    var folderIdColIndex = -1;
    for (var i = 0; i < headerRow.length; i++) {
      if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader("Year")) {
        yearColIndex = i;
      }
      if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader("Roster Folder ID")) {
        folderIdColIndex = i;
      }
    }
    
    if (yearColIndex === -1 || folderIdColIndex === -1) {
      UtilityScriptLibrary.debugLog("⚠️ Required columns not found in Year Metadata sheet.");
      return;
    }

    var yearRow = null;
    for (var i = 0; i < metadataRows.length; i++) {
      if (metadataRows[i][yearColIndex] && metadataRows[i][yearColIndex].toString() === year) {
        yearRow = metadataRows[i];
        break;
      }
    }
    if (!yearRow) {
      UtilityScriptLibrary.debugLog("⚠️ No roster folder found for year: " + year);
      return;
    }
  
    var rosterFolderId = yearRow[folderIdColIndex];
    var rosterFolder = DriveApp.getFolderById(rosterFolderId);
    UtilityScriptLibrary.debugLog("✅ Found roster folder for year " + year);

    processRoster(formData, sheet, row, headerMap, fieldMap, studentId, rosterFolder, year, semesterName);

  } catch (rosterError) {
    UtilityScriptLibrary.debugLog("⚠️ Error in roster processing: " + rosterError.message);
  }
}

function processStudent(formData, contactsSheet, enrollmentTerm) {
  try {
    UtilityScriptLibrary.debugLog("👤 Starting processStudent");
    
    var studentKey = UtilityScriptLibrary.generateKey(
      formData["Student Last Name"] || '',
      formData["Student First Name"] || '',
      formData["Instrument"] || ''
    );

    var studentRow = UtilityScriptLibrary.findStudentRow(contactsSheet, studentKey);
    
    UtilityScriptLibrary.debugLog("=== STUDENT DUPLICATE CHECK ===");
    UtilityScriptLibrary.debugLog("Looking for key: '" + studentKey + "'");
    UtilityScriptLibrary.debugLog("findStudentRow result: " + studentRow);
    UtilityScriptLibrary.debugLog("=== END STUDENT DEBUG ===");

    var headers = contactsSheet.getRange(1, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
    var getCol = function(name) {
      for (var i = 0; i < headers.length; i++) {
        if (UtilityScriptLibrary.normalizeHeader(String(headers[i])) === UtilityScriptLibrary.normalizeHeader(name)) {
          return i + 1;
        }
      }
      return 0;
    };

    var requiredFields = [
      "Student Last Name", "Student First Name", "Instrument", "Teacher",
      "Age", "Currently Registered", "Student ID", "Parent ID", "Student Lookup", "First Enrollment Term"
    ];
    var missingFields = [];
    for (var i = 0; i < requiredFields.length; i++) {
      if (getCol(requiredFields[i]) === 0) {
        missingFields.push(requiredFields[i]);
      }
    }

    if (missingFields.length > 0) {
      throw new Error("Missing required columns in Students sheet: " + missingFields.join(", "));
    }

    // Teacher ID is stored directly — onEdit already swapped display name to ID
    var teacherId = String(formData["Teacher"] || '').trim();
    if (!teacherId) {
      UtilityScriptLibrary.debugLog("processStudent", "WARNING", "No Teacher ID found in form data", "", "");
    }

    var studentId = formData["Student ID"] || '';
    var parentId = '';

    // Convert age response to standardized Adult/Child values
    var ageResponse = formData["Age"] || '';
    var standardizedAge = '';
    if (ageResponse.toString().toLowerCase().indexOf('yes') === 0) {
      standardizedAge = 'Adult';
    } else if (ageResponse.toString().toLowerCase().indexOf('no') === 0) {
      standardizedAge = 'Child';
    } else {
      standardizedAge = 'Child';
    }
    
    UtilityScriptLibrary.debugLog("Age conversion: Original='" + ageResponse + "', Standardized='" + standardizedAge + "'");

    if (studentRow !== -1) {
      UtilityScriptLibrary.debugLog("📄 Updating existing student");
      var rowData = contactsSheet.getRange(studentRow, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
      studentId = String(rowData[getCol("Student ID") - 1] || '');
      parentId = String(rowData[getCol("Parent ID") - 1] || '');

      var registeredCol = getCol("Currently Registered");
      if (registeredCol) {
        var checkboxCell = contactsSheet.getRange(studentRow, registeredCol);
        checkboxCell.insertCheckboxes();
        checkboxCell.setValue(true);
      }

      if (getCol("Teacher")) {
        contactsSheet.getRange(studentRow, getCol("Teacher")).setValue(teacherId);
      }
      
      if (getCol("Age")) {
        contactsSheet.getRange(studentRow, getCol("Age")).setValue(standardizedAge);
      }
      
      if (getCol("Graduation Year")) {
        var graduationYear = calculateGraduationYear(formData["Grade"]);
        contactsSheet.getRange(studentRow, getCol("Graduation Year")).setValue(graduationYear);
      }

      if (getCol("School District")) {
        contactsSheet.getRange(studentRow, getCol("School District")).setValue(formData["School"] || '');
      }

      if (getCol("School Teacher")) {
        contactsSheet.getRange(studentRow, getCol("School Teacher")).setValue(formData["SchoolTeacher"] || '');
      }

      if (getCol("Experience")) {
        contactsSheet.getRange(studentRow, getCol("Experience")).setValue(formData["Experience"] || '');
      }

      if (getCol("Experience Start Range")) {
        var experienceStartRange = calculateExperienceStartRange(formData["Experience"]);
        contactsSheet.getRange(studentRow, getCol("Experience Start Range")).setValue(experienceStartRange);
      }

    } else {
      UtilityScriptLibrary.debugLog("➕ Creating new student");
      studentId = UtilityScriptLibrary.generateNextId(contactsSheet, 'Student ID', 'Q', (formData["Student First Name"] || '') + ' ' + (formData["Student Last Name"] || ''));

      var newRow = new Array(headers.length);
      for (var i = 0; i < headers.length; i++) {
        newRow[i] = '';
      }

      newRow[getCol("Student ID") - 1] = studentId;
      newRow[getCol("Student Last Name") - 1] = formData["Student Last Name"] || '';
      newRow[getCol("Student First Name") - 1] = formData["Student First Name"] || '';
      newRow[getCol("Instrument") - 1] = formData["Instrument"] || '';
      newRow[getCol("Teacher") - 1] = teacherId;
      newRow[getCol("Age") - 1] = standardizedAge;
      newRow[getCol("First Enrollment Term") - 1] = enrollmentTerm || '';
      newRow[getCol("Student Lookup") - 1] = studentKey;

      if (getCol("Graduation Year")) {
        var graduationYear = calculateGraduationYear(formData["Grade"]);
        newRow[getCol("Graduation Year") - 1] = graduationYear;
      }

      if (getCol("School District")) {
        newRow[getCol("School District") - 1] = formData["School"] || '';
      }

      if (getCol("School Teacher")) {
        newRow[getCol("School Teacher") - 1] = formData["SchoolTeacher"] || '';
      }

      if (getCol("Experience")) {
        newRow[getCol("Experience") - 1] = formData["Experience"] || '';
      }

      if (getCol("Experience Start Range")) {
        var experienceStartRange = calculateExperienceStartRange(formData["Experience"]);
        newRow[getCol("Experience Start Range") - 1] = experienceStartRange;
      }

      contactsSheet.appendRow(newRow);

      var registeredCol = getCol("Currently Registered");
      if (registeredCol) {
        var newRowIndex = contactsSheet.getLastRow();
        var cell = contactsSheet.getRange(newRowIndex, registeredCol);
        cell.insertCheckboxes();
        cell.setValue(true);
      }
    }

    UtilityScriptLibrary.debugLog("✅ Completed processStudent - ID: " + studentId);
    return { studentId: studentId, parentId: parentId, studentRow: studentRow, teacherId: teacherId };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in processStudent: " + error.message);
    throw error;
  }
}

function processStudentSelection(selectedIndices) {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();
    
    if (!selectedIndices || selectedIndices.length === 0) {
      UtilityScriptLibrary.debugLog("User cancelled at Student Selection prompt");
      return;
    }
    
    var studentList = JSON.parse(scriptProps.getProperty('reassign_studentList'));
    
    // Get selected students by indices
    var selectedStudents = [];
    for (var i = 0; i < selectedIndices.length; i++) {
      var index = selectedIndices[i];
      if (index >= 0 && index < studentList.length) {
        selectedStudents.push(studentList[index]);
      }
    }
    
    if (selectedStudents.length === 0) {
      ui.alert('Error', 'No students were selected.', ui.ButtonSet.OK);
      return;
    }
    
    UtilityScriptLibrary.debugLog("User selected " + selectedStudents.length + " students");
    
    // Store selected students
    scriptProps.setProperty('reassign_selectedStudents', JSON.stringify(selectedStudents));
    
    // Move to step 3: select new teacher
    selectNewTeacher();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in processStudentSelection: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Student selection failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function reassignStudentToNewTeacher() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    UtilityScriptLibrary.debugLog("=== STARTING STUDENT REASSIGNMENT ===");
    
    // Get current semester
    var currentSemester = UtilityScriptLibrary.getCurrentSemesterName();
    if (!currentSemester) {
      ui.alert('Error', 'Could not determine current semester from Calendar sheet.', ui.ButtonSet.OK);
      return;
    }
    
    var year = UtilityScriptLibrary.getYearFromSemesterName(currentSemester);
    var season = UtilityScriptLibrary.extractSeasonFromSemester(currentSemester);
    
    UtilityScriptLibrary.debugLog("Current semester: " + currentSemester + " (Year: " + year + ", Season: " + season + ")");
    
    // Store semester info for later steps
    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('reassign_currentSemester', currentSemester);
    scriptProps.setProperty('reassign_year', year);
    scriptProps.setProperty('reassign_season', season);
    
    // Step 1: Get active teachers and show dropdown for Old Teacher
    var activeTeachers = getActiveTeachersForDropdown();
    if (activeTeachers.length === 0) {
      ui.alert('Error', 'No active teachers found. Please check Teacher Roster Lookup.', ui.ButtonSet.OK);
      return;
    }
    
    // Store active teachers list for step 3
    scriptProps.setProperty('reassign_activeTeachers', JSON.stringify(activeTeachers));
    
    // Show HTML dropdown for old teacher - with callback to step 2
    showTeacherDropdownDialog(
      'Reassign Students - Step 1 of 4',
      'Select the OLD teacher (current teacher):',
      activeTeachers,
      'selectStudents'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in reassignStudentToNewTeacher: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Reassignment failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function refreshCurrentSemesterTeacherDropdown() {
  try {
    UtilityScriptLibrary.debugLog('refreshCurrentSemesterTeacherDropdown', 'INFO', 'Manual refresh triggered', '', '');
    
    var currentSemester = UtilityScriptLibrary.getCurrentSemesterName();
    if (!currentSemester) {
      SpreadsheetApp.getUi().alert('❌ No current semester found. Please ensure calendar is set up correctly.');
      return;
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var semesterSheet = ss.getSheetByName(currentSemester);
    
    if (!semesterSheet) {
      SpreadsheetApp.getUi().alert('❌ Current semester sheet "' + currentSemester + '" not found.');
      return;
    }
    
    // Apply teacher dropdown
    applyTeacherDropdownToSheet(semesterSheet);
    
    SpreadsheetApp.getUi().alert('✅ Teacher dropdown refreshed for "' + currentSemester + '"');
    
    UtilityScriptLibrary.debugLog('refreshCurrentSemesterTeacherDropdown', 'INFO', 'Manual refresh completed', currentSemester, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('refreshCurrentSemesterTeacherDropdown', 'ERROR', 'Manual refresh failed', '', error.message);
    SpreadsheetApp.getUi().alert('❌ Error refreshing teacher dropdown: ' + error.message);
  }
}

function runLogHeaders() {
  UtilityScriptLibrary.logAllSheetHeaders();
}

function selectNewTeacher() {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();
    
    var oldTeacherDisplay = scriptProps.getProperty('reassign_oldTeacherDisplay');
    var activeTeachers = JSON.parse(scriptProps.getProperty('reassign_activeTeachers'));
    
    // Filter out old teacher
    var availableTeachers = activeTeachers.filter(function(t) { return t !== oldTeacherDisplay; });
    
    if (availableTeachers.length === 0) {
      ui.alert('Error', 'No other active teachers available for reassignment.', ui.ButtonSet.OK);
      return;
    }
    
    // Show HTML dropdown for new teacher - with callback to step 4
    showTeacherDropdownDialog(
      'Reassign Students - Step 3 of 4',
      'Select the NEW teacher:',
      availableTeachers,
      'enterEffectiveDate'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in selectNewTeacher: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 3 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function selectStudents(oldTeacherDisplay) {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();
    
    if (!oldTeacherDisplay) {
      UtilityScriptLibrary.debugLog("User cancelled at Old Teacher prompt");
      return;
    }
    
    UtilityScriptLibrary.debugLog("Old teacher (display name): " + oldTeacherDisplay);
    
    // Store old teacher for later steps
    scriptProps.setProperty('reassign_oldTeacherDisplay', oldTeacherDisplay);
    
    // Get old teacher's info and roster
    var oldTeacherInfo = getTeacherInfoByDisplayName(oldTeacherDisplay);
    if (!oldTeacherInfo || !oldTeacherInfo.rosterUrl) {
      ui.alert('Error', 'Could not find roster URL for old teacher: ' + oldTeacherDisplay, ui.ButtonSet.OK);
      return;
    }
    
    UtilityScriptLibrary.debugLog("Old teacher full name: " + oldTeacherInfo.firstName + ' ' + oldTeacherInfo.lastName + ", URL: " + oldTeacherInfo.rosterUrl);
    
    // Store for later
    scriptProps.setProperty('reassign_oldTeacherInfo', JSON.stringify(oldTeacherInfo));
    
    // Open old teacher's roster and get students
    var currentSemester = scriptProps.getProperty('reassign_currentSemester');
    var oldRosterWorkbook = SpreadsheetApp.openByUrl(oldTeacherInfo.rosterUrl);
    var oldRosterSheet = findSemesterRoster(oldRosterWorkbook, currentSemester);
    
    if (!oldRosterSheet) {
      ui.alert('Error', 'Could not find roster sheet for ' + currentSemester + ' in old teacher\'s workbook.', ui.ButtonSet.OK);
      return;
    }
    
    // Get list of students from old teacher's roster
    var studentList = getActiveStudentsFromRoster(oldRosterSheet);
    
    if (studentList.length === 0) {
      ui.alert('Error', 'No active students found in old teacher\'s roster.', ui.ButtonSet.OK);
      return;
    }
    
    UtilityScriptLibrary.debugLog("Found " + studentList.length + " active students in old teacher's roster");
    
    // Store student list and old roster sheet NAME (not ID) for later
    scriptProps.setProperty('reassign_studentList', JSON.stringify(studentList));
    scriptProps.setProperty('reassign_oldRosterSheetName', oldRosterSheet.getName()); // CHANGED
    scriptProps.setProperty('reassign_oldRosterWorkbookId', oldRosterWorkbook.getId());
    
    // Show checkbox dialog for student selection
    showStudentCheckboxDialog(
      'Reassign Students - Step 2 of 4',
      'Select students to transfer:',
      studentList,
      'processStudentSelection'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in selectStudents: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 2 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function setupCompleteRosterWorkbook(spreadsheet, teacher, year, semesterName, registrationTimestamp) {
  try {
    UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Setting up complete roster workbook structure for Teacher: " + teacher + ", Year: " + year + ", Semester: " + semesterName);
    
    // Extract season from semesterName (e.g., "Spring 2024" -> "Spring")
    var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
    if (!season) {
      throw new Error("Could not extract season from semester name: " + semesterName);
    }
    
    // 1. Create and set up season-named roster sheet (empty)
    var rosterSheetName = season + " Roster";
    var rosterSheet = spreadsheet.insertSheet(rosterSheetName);
    setupNewRosterTemplate(rosterSheet);
    UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Created season roster sheet: " + rosterSheetName);
    
    // 2. Create attendance sheet for current semester month (empty)
    var currentSemesterMonth = UtilityScriptLibrary.getCurrentSemesterMonth(semesterName);
    if (!currentSemesterMonth) {
      throw new Error("Could not determine current semester month for: " + semesterName);
    }
    
    var attendanceSheet = UtilityScriptLibrary.createMonthlyAttendanceSheet(spreadsheet, currentSemesterMonth, []); // EMPTY
    UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Created empty attendance sheet for Month: " + currentSemesterMonth + ", Semester: " + semesterName);
    
    // 3. Create Invoice Log sheet
    createInvoiceLogSheet(spreadsheet);
    UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Created Invoice Log sheet");
    
    // 4. Remove default "Sheet1" (after we have other sheets)
    try {
      var sheets = spreadsheet.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() === "Sheet1") {
          spreadsheet.deleteSheet(sheets[i]);
          UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Removed default Sheet1");
          break;
        }
      }
    } catch (deleteError) {
      UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - WARNING: Could not delete Sheet1: " + deleteError.message);
      // Non-critical error, continue
    }
    
    // 5. Set sheet order (Roster first, then current month)
    try {
      spreadsheet.setActiveSheet(rosterSheet);
      spreadsheet.moveActiveSheet(1);
    } catch (orderError) {
      UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - WARNING: Could not set sheet order: " + orderError.message);
      // Non-critical error, continue
    }
    
    UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Complete workbook structure created. Roster: " + rosterSheetName + ", Attendance: " + currentSemesterMonth + ", Invoice Log: created");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - ERROR: Error setting up complete workbook for Teacher: " + teacher + ", Semester: " + semesterName + ". Error: " + error.message);
    throw error;
  }
}

function setupInvoiceLogHeaders(sheet) {
  var headers = [
    'Invoice Number',       // A
    'Invoice Date',         // B
    'Invoice Period',       // C
    'Invoice URL',          // D
    'Total Amount'          // E
  ];
  
  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Style header row using utility function
  UtilityScriptLibrary.styleHeaderRow(sheet, headers);
  
  // Set column widths
  var widths = [120, 100, 155, 400, 100];  // Invoice Number, Date, Period, URL, Amount
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  UtilityScriptLibrary.debugLog('✅ Invoice Log headers set up with formatting');
}

function setupNewRosterTemplate(sheet) {
  try {
    UtilityScriptLibrary.debugLog("🎨 Setting up new roster template structure");
    
    // Clear existing content but keep the sheet
    sheet.clear();
    
    // Set up headers - UPDATED: Changed columns R and S
    var headers = [
      'Contacted',              // A - Checkbox (editable)
      'First Lesson Date',      // B - Date (editable) 
      'First Lesson Time',      // C - Time (editable)
      'Comments',               // D - Text (editable)
      'Last Name',              // E - Admin only
      'First Name',             // F - Admin only
      'Instrument',             // G - Admin only
      'Length',                 // H - Admin only
      'Experience',             // I - Admin only
      'Grade',                  // J - Admin only
      'School',                 // K - Admin only
      'School Teacher',         // L - Admin only
      'Parent Last Name',       // M - Admin only
      'Parent First Name',      // N - Admin only
      'Phone',                  // O - Admin only
      'Email',                  // P - Admin only
      'Additional contacts',    // Q - Admin only
      'Hours Remaining',        // R - Admin only (CHANGED from Lessons Registered)
      'Lessons Remaining',      // S - Admin only (CHANGED from Lessons Completed)
      'Status',                 // T - Admin only (MOVED from U)
      'Student ID',             // U - Admin only (MOVED from V)
      'Admin Comments',         // V - Admin only (MOVED from W)
      'System Comments'         // W - Admin only (MOVED from X)
    ];
    
    // Set headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply formatting
    setupRosterTemplateFormatting(sheet);
    
    UtilityScriptLibrary.debugLog("✅ New roster template structure applied");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error setting up new roster template: " + error.message);
  }
}

function setupRosterTemplateFormatting(sheet) {
  try {
    // Style header row - WITH text wrapping
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground(UtilityScriptLibrary.STYLES.HEADER.background)
               .setFontColor(UtilityScriptLibrary.STYLES.HEADER.text)
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle')
               .setWrap(true);
    
    // Set specific column widths - UPDATED for new structure
    var columnWidths = [
      75,   // A - Contacted (checkbox)
      95,   // B - First Lesson Date
      95,   // C - First Lesson Time  
      220,  // D - Comments
      120,  // E - Last Name
      120,  // F - First Name
      80,   // G - Instrument
      55,   // H - Length
      120,  // I - Experience
      55,   // J - Grade
      110,  // K - School
      110,  // L - School Teacher
      120,  // M - Parent Last Name
      120,  // N - Parent First Name
      100,  // O - Phone
      220,  // P - Email
      200,  // Q - Additional contacts
      80,   // R - Hours Remaining (CHANGED)
      80,   // S - Lessons Remaining (CHANGED)
      80,   // T - Status (MOVED)
      60,   // U - Student ID (MOVED)
      220,  // V - Admin Comments (MOVED)
      220   // W - System Comments (MOVED)
    ];
    
    // Apply column widths
    for (var i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }
    
    // Set text wrapping for specific columns
    var maxRows = sheet.getMaxRows();
    
    // No wrap for phone number (column O)
    sheet.getRange(1, 15, maxRows, 1).setWrap(false);
    
    // Wrap for lesson tracking columns (R, S) 
    sheet.getRange(1, 18, maxRows, 2).setWrap(true);
    
    // Number format for Hours Remaining (column R) - 2 decimal places
    sheet.getRange(2, 18, maxRows - 1, 1).setNumberFormat('0.00');
    
    // Number format for Lessons Remaining (column S) - whole numbers
    sheet.getRange(2, 19, maxRows - 1, 1).setNumberFormat('0');
    
    // Add thick green borders
    addRosterTemplateBorders(sheet);
    
    // Set up protection and validation
    UtilityScriptLibrary.setupRosterTemplateProtection(sheet);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    UtilityScriptLibrary.debugLog("✅ Roster template formatting applied");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Error in formatting: " + error.message);
  }
}

function shouldProcessEdit(e, sheet) {
  //UPDATED 11/6/25
  try {
    var getCol = UtilityScriptLibrary.createColumnFinder(sheet);

    var teacherCol = getCol("Teacher");
    var idCol = getCol("Student ID");

    UtilityScriptLibrary.debugLog("Teacher column index: " + teacherCol);
    UtilityScriptLibrary.debugLog("Student ID column index: " + idCol);
    UtilityScriptLibrary.debugLog("Edited column: " + e.range.getColumn());

    // Skip if editing the Student ID column itself
    if (e.range.getColumn() === idCol) {
      UtilityScriptLibrary.debugLog("🛑 Edit was to Student ID column. Ignoring.");
      return false;
    }

    // Skip if not editing Teacher column
    if (e.range.getColumn() !== teacherCol) {
      UtilityScriptLibrary.debugLog("🛑 Edit was not to Teacher column. Ignoring.");
      return false;
    }

    var editedRow = e.range.getRow();
    var studentIdValue = sheet.getRange(editedRow, idCol).getValue();
    var teacherValue = sheet.getRange(editedRow, teacherCol).getValue();

    UtilityScriptLibrary.debugLog("Student ID cell value (live): '" + String(studentIdValue) + "'");
    UtilityScriptLibrary.debugLog("Teacher cell value (live): '" + String(teacherValue) + "'");

    // Skip if teacher field is empty
    if (!teacherValue || String(teacherValue).trim() === "") {
      UtilityScriptLibrary.debugLog("❌ Teacher field is empty. Skipping.");
      return false;
    }

    // Skip if Student ID ALREADY EXISTS (duplicate/reprocessing prevention)
    if (studentIdValue && String(studentIdValue).trim() !== '') {
      UtilityScriptLibrary.debugLog("❌ Student ID already exists. Skipping to prevent duplicate processing.");
      return false;
    }

    UtilityScriptLibrary.debugLog("✅ Edit validation passed - processing row " + editedRow);
    return true;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in shouldProcessEdit: " + error.message);
    return false;
  }
}

function showStudentCheckboxDialog(title, message, studentList, callbackFunctionName) {
  var html = HtmlService.createHtmlOutput()
    .setWidth(500)
    .setHeight(400);
  
  var htmlContent = '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }' +
    'h3 { margin-top: 0; color: #333; }' +
    'p { color: #666; margin-bottom: 15px; }' +
    '.student-list { max-height: 250px; overflow-y: auto; border: 1px solid #ddd; border-radius: 4px; padding: 10px; margin: 10px 0; }' +
    '.student-item { padding: 8px; margin: 4px 0; }' +
    '.student-item:hover { background-color: #f5f5f5; }' +
    '.student-item label { cursor: pointer; display: block; }' +
    '.student-item input[type="checkbox"] { margin-right: 10px; cursor: pointer; }' +
    'button { padding: 10px 20px; margin: 5px; font-size: 14px; cursor: pointer; border-radius: 4px; border: none; }' +
    '.ok-btn { background-color: #4CAF50; color: white; }' +
    '.ok-btn:hover { background-color: #45a049; }' +
    '.cancel-btn { background-color: #f44336; color: white; }' +
    '.cancel-btn:hover { background-color: #da190b; }' +
    '.select-btn { background-color: #2196F3; color: white; }' +
    '.select-btn:hover { background-color: #0b7dda; }' +
    '.button-container { margin-top: 15px; text-align: right; }' +
    '.utility-buttons { margin-bottom: 10px; }' +
    '</style>' +
    '<div>' +
    '<h3>' + title + '</h3>' +
    '<p>' + message + '</p>' +
    '<div class="utility-buttons">' +
    '<button class="select-btn" onclick="selectAll()">Select All</button>' +
    '<button class="select-btn" onclick="clearAll()">Clear All</button>' +
    '</div>' +
    '<div class="student-list" id="studentList">';
  
  // Add checkboxes for each student
  for (var i = 0; i < studentList.length; i++) {
    var student = studentList[i];
    var displayName = student.firstName + ' ' + student.lastName + ' (' + student.studentId + ') - ' + student.instrument;
    htmlContent += '<div class="student-item">' +
      '<label>' +
      '<input type="checkbox" value="' + i + '" class="student-checkbox"> ' +
      displayName +
      '</label>' +
      '</div>';
  }
  
  htmlContent += '</div>' +
    '<div class="button-container">' +
    '<button class="cancel-btn" onclick="handleCancel()">Cancel</button>' +
    '<button class="ok-btn" onclick="handleOk()">OK</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'function selectAll() {' +
    '  var checkboxes = document.querySelectorAll(".student-checkbox");' +
    '  for (var i = 0; i < checkboxes.length; i++) {' +
    '    checkboxes[i].checked = true;' +
    '  }' +
    '}' +
    'function clearAll() {' +
    '  var checkboxes = document.querySelectorAll(".student-checkbox");' +
    '  for (var i = 0; i < checkboxes.length; i++) {' +
    '    checkboxes[i].checked = false;' +
    '  }' +
    '}' +
    'function handleOk() {' +
    '  var checkboxes = document.querySelectorAll(".student-checkbox");' +
    '  var selected = [];' +
    '  for (var i = 0; i < checkboxes.length; i++) {' +
    '    if (checkboxes[i].checked) {' +
    '      selected.push(parseInt(checkboxes[i].value));' +
    '    }' +
    '  }' +
    '  if (selected.length === 0) {' +
    '    alert("Please select at least one student.");' +
    '    return;' +
    '  }' +
    '  google.script.run' +
    '    .withSuccessHandler(function() {' +
    '      google.script.host.close();' +
    '    })' +
    '    .' + callbackFunctionName + '(selected);' +
    '}' +
    'function handleCancel() {' +
    '  google.script.run' +
    '    .withSuccessHandler(function() {' +
    '      google.script.host.close();' +
    '    })' +
    '    .' + callbackFunctionName + '([]);' +
    '}' +
    '</script>';
  
  html.setContent(htmlContent);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function showTeacherDropdownDialog(title, message, teacherList, callbackFunctionName) {
  var html = HtmlService.createHtmlOutput()
    .setWidth(400)
    .setHeight(300);
  
  var htmlContent = '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }' +
    'h3 { margin-top: 0; color: #333; }' +
    'p { color: #666; margin-bottom: 15px; }' +
    'select { width: 100%; padding: 10px; font-size: 14px; margin: 10px 0; border: 1px solid #ddd; border-radius: 4px; }' +
    'select option { padding: 8px; }' +
    'button { padding: 10px 20px; margin: 5px; font-size: 14px; cursor: pointer; border-radius: 4px; border: none; }' +
    '.ok-btn { background-color: #4CAF50; color: white; }' +
    '.ok-btn:hover { background-color: #45a049; }' +
    '.cancel-btn { background-color: #f44336; color: white; }' +
    '.cancel-btn:hover { background-color: #da190b; }' +
    '.button-container { margin-top: 20px; text-align: right; }' +
    '</style>' +
    '<div>' +
    '<h3>' + title + '</h3>' +
    '<p>' + message + '</p>' +
    '<select id="teacherSelect" size="10">';
  
  for (var i = 0; i < teacherList.length; i++) {
    htmlContent += '<option value="' + teacherList[i] + '">' + teacherList[i] + '</option>';
  }
  
  htmlContent += '</select>' +
    '<div class="button-container">' +
    '<button class="cancel-btn" onclick="handleCancel()">Cancel</button>' +
    '<button class="ok-btn" onclick="handleOk()">OK</button>' +
    '</div>' +
    '</div>' +
    '<script>' +
    'document.getElementById("teacherSelect").focus();' +
    'document.getElementById("teacherSelect").selectedIndex = 0;' +
    'document.getElementById("teacherSelect").addEventListener("dblclick", function() { handleOk(); });' +
    'document.getElementById("teacherSelect").addEventListener("keydown", function(e) {' +
    '  if (e.key === "Enter") handleOk();' +
    '  if (e.key === "Escape") handleCancel();' +
    '});' +
    'function handleOk() {' +
    '  var select = document.getElementById("teacherSelect");' +
    '  if (select.selectedIndex >= 0) {' +
    '    google.script.run' +
    '      .withSuccessHandler(function() {' +
    '        google.script.host.close();' +
    '      })' +
    '      .' + callbackFunctionName + '(select.value);' +
    '  }' +
    '}' +
    'function handleCancel() {' +
    '  google.script.run' +
    '    .withSuccessHandler(function() {' +
    '      google.script.host.close();' +
    '    })' +
    '    .' + callbackFunctionName + '(null);' +
    '}' +
    '</script>';
  
  html.setContent(htmlContent);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function studentExistsInAttendanceSheet(attendanceSheet, studentId) {
  try {
    var data = attendanceSheet.getDataRange().getValues();
    var headers = data[0];
    var studentIdCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (UtilityScriptLibrary.normalizeHeader(headers[i]) === 'studentid') {
        studentIdCol = i;
        break;
      }
    }
    if (studentIdCol === -1) {
      UtilityScriptLibrary.debugLog('studentExistsInAttendanceSheet', 'ERROR', 'Student ID column not found', '', '');
      return false;
    }
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol] && data[i][studentIdCol].toString().trim() === studentId.toString().trim()) {
        return true;
      }
    }
    return false;
  } catch (error) {
    UtilityScriptLibrary.debugLog('studentExistsInAttendanceSheet', 'ERROR', 'Error checking if student exists', '', error.message);
    return false;
  }
}

function updateAllTeacherGroupAssignments() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'INFO', 'Starting batch group assignment update', '', '');
    
    var semesterName = UtilityScriptLibrary.getCurrentSemesterName();
    if (!semesterName) {
      ui.alert('Error: Could not determine current semester from Calendar sheet.');
      return;
    }
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'INFO', 'Found current semester', semesterName, '');
    
    var allTeachers = getAllTeachersWithGroupAssignments();
    
    if (!allTeachers || allTeachers.length === 0) {
      ui.alert('No teachers with group assignments found.');
      return;
    }
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'INFO', 'Found teachers with groups', 
                                 'Count: ' + allTeachers.length, '');
    
    var successCount = 0;
    var errorCount = 0;
    var errors = [];
    
    for (var i = 0; i < allTeachers.length; i++) {
      var teacherName = allTeachers[i];
      try {
        updateGroupAssignmentsForCurrentMonth(teacherName.firstName, teacherName.lastName, semesterName);
        successCount++;
      } catch (error) {
        errorCount++;
        errors.push(teacherName.firstName + ' ' + teacherName.lastName + ': ' + error.message);
        UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'ERROR', 'Failed for teacher',
                                     teacherName.firstName + ' ' + teacherName.lastName, error.message);
      }
    }
    
    var message = 'Group Assignment Update Complete\n\n';
    message += 'Successfully updated: ' + successCount + ' teacher(s)\n';
    if (errorCount > 0) {
      message += 'Errors: ' + errorCount + '\n\n';
      message += 'Failed teachers:\n' + errors.join('\n');
    }
    
    ui.alert(message);
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'SUCCESS', 'Batch update complete', 
                                 'Success: ' + successCount + ', Errors: ' + errorCount, '');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'ERROR', 'Batch update failed', '', error.message);
  }
}

function updateGroupAssignmentsForCurrentMonth(firstName, lastName, semesterName) {
  try {
    UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'Starting',
      'Teacher: ' + firstName + ' ' + lastName + ', Semester: ' + semesterName, '');

    var groupAssignments = UtilityScriptLibrary.getTeacherGroupAssignments(firstName + ' ' + lastName);
    if (!groupAssignments || groupAssignments.length === 0) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'No group assignments found',
        'Teacher: ' + firstName + ' ' + lastName, '');
      return;
    }

    var teacherInfo = getTeacherInfoByFullName(firstName, lastName);
    if (!teacherInfo || !teacherInfo.rosterUrl) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'Teacher roster URL not found',
        'Teacher: ' + firstName + ' ' + lastName, '');
      return;
    }

    var rosterSS = SpreadsheetApp.openByUrl(teacherInfo.rosterUrl);
    if (!rosterSS) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'Could not open teacher workbook',
        'Teacher: ' + firstName + ' ' + lastName, '');
      return;
    }

    var attendanceSheet = getMostRecentMonthSheet(rosterSS);
    if (!attendanceSheet) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'No month attendance sheet found',
        'Teacher: ' + firstName + ' ' + lastName, '');
      return;
    }

    var existingGroupIds = getExistingGroupIds(attendanceSheet);

    var newGroupAssignments = [];
    for (var i = 0; i < groupAssignments.length; i++) {
      var alreadyExists = false;
      for (var j = 0; j < existingGroupIds.length; j++) {
        if (existingGroupIds[j] === groupAssignments[i].groupId) {
          alreadyExists = true;
          break;
        }
      }
      if (!alreadyExists) newGroupAssignments.push(groupAssignments[i]);
    }

    if (newGroupAssignments.length === 0) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'All group assignments already exist',
        'Teacher: ' + firstName + ' ' + lastName, '');
      return;
    }

    UtilityScriptLibrary.createGroupSections(attendanceSheet, newGroupAssignments);

    var totalRows = attendanceSheet.getLastRow();
    if (totalRows > 1) {
      UtilityScriptLibrary.setupStatusValidation(attendanceSheet, totalRows);
    }

    UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'Successfully added group assignments',
      'Teacher: ' + firstName + ' ' + lastName + ', Added: ' + newGroupAssignments.length, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'Failed',
      'Teacher: ' + firstName + ' ' + lastName, error.message);
    throw error;
  }
}

function updateStudentWithParentId(contactsSheet, studentRow, parentId) {
  //UPDATED 11-5-25
  try {
    UtilityScriptLibrary.debugLog("🔗 Linking student with parent ID: " + parentId);
    
    var getCol = UtilityScriptLibrary.createColumnFinder(contactsSheet);
    
    var parentIdCol = getCol("Parent ID");
    if (parentIdCol === 0) {
      throw new Error("Parent ID column not found in students sheet");
    }

    if (studentRow !== -1) {
      var currentParentId = contactsSheet.getRange(studentRow, parentIdCol).getValue();
      if (!currentParentId || String(currentParentId).trim() === '') {
        contactsSheet.getRange(studentRow, parentIdCol).setValue(parentId);
        UtilityScriptLibrary.debugLog("✅ Updated existing student with parent ID");
      }
    } else {
      var lastRow = contactsSheet.getLastRow();
      contactsSheet.getRange(lastRow, parentIdCol).setValue(parentId);
      UtilityScriptLibrary.debugLog("✅ Updated new student with parent ID");
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in updateStudentWithParentId: " + error.message);
    throw error;
  }
}

function updateTeacherRosterLookup(teacherId, fileUrl) {
  try {
    var lookupSheet = getTeacherRosterLookupSheet();

    if (!lookupSheet) {
      UtilityScriptLibrary.debugLog('updateTeacherRosterLookup', 'ERROR', 'Teacher Roster Lookup sheet not found', '', '');
      return;
    }

    var existingRow = UtilityScriptLibrary.findTeacherInRosterLookup(lookupSheet, teacherId);

    if (existingRow === -1) {
      UtilityScriptLibrary.debugLog('updateTeacherRosterLookup', 'ERROR', 'Teacher not found in lookup', teacherId, '');
      return;
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var rosterUrlCol = getCol('Roster URL');
    var statusCol = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');

    if (rosterUrlCol) lookupSheet.getRange(existingRow, rosterUrlCol).setValue(fileUrl);
    if (statusCol) lookupSheet.getRange(existingRow, statusCol).setValue('Active');
    if (lastUpdatedCol) lookupSheet.getRange(existingRow, lastUpdatedCol).setValue(new Date());

    UtilityScriptLibrary.debugLog('updateTeacherRosterLookup', 'INFO', 'Updated teacher roster lookup', teacherId, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('updateTeacherRosterLookup', 'ERROR', 'Failed', teacherId, error.message);
  }
}

function verifyByDriveId(driveId) {
  try {
    if (!driveId) {
      Logger.log('❌ No Drive ID provided');
      return;
    }
    
    driveId = driveId.trim();
    Logger.log('🔍 Processing Drive ID: ' + driveId);
    
    var studentMap = loadStudentMapFromContacts();
    var detailIssues = [];
    var summaryData = [];
    
    // Try as folder first
    try {
      var folder = DriveApp.getFolderById(driveId);
      Logger.log('📁 Processing folder: ' + folder.getName());
      checkWorkbooksInFolder(folder, studentMap, detailIssues, summaryData, false);
    } catch (e) {
      // Not a folder, try as file
      try {
        var file = DriveApp.getFileById(driveId);
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
          throw new Error('File is not a Google Spreadsheet');
        }
        Logger.log('📄 Processing file: ' + file.getName());
        var workbook = SpreadsheetApp.openById(driveId);
        checkWorkbook(workbook, file.getName(), studentMap, detailIssues, summaryData);
      } catch (e2) {
        throw new Error('Invalid ID or no access: ' + e2.message);
      }
    }
    
    appendToReports(detailIssues, summaryData);
    Logger.log('✅ Complete. Found ' + detailIssues.length + ' issues.');
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
  }
}

function verifyByDriveIdWithPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Verify Student IDs',
    'Enter Google Drive ID (folder or spreadsheet):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var driveId = response.getResponseText().trim();
  if (!driveId) {
    ui.alert('No ID provided');
    return;
  }
  
  verifyByDriveId(driveId);
  ui.alert('Complete! Check the report sheets.');
}
