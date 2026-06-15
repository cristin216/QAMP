/*
================================================================================
REGISTRATION CODE
================================================================================
Version: 119
Total Functions: 75
Documentation: See Registration-Functions.md
================================================================================
*/

function backfillParentIds() {
  var norm = UtilityScriptLibrary.normalizeHeader;

  // === PART 1: Form responses sheet ===
  var formSS = UtilityScriptLibrary.getWorkbook('formResponses');
  var formSheet = formSS.getSheetByName('Summer 2026');
  if (!formSheet) throw new Error('Summer 2026 form sheet not found.');

  var formHeaders = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
  var formGetCol = function(name) {
    for (var i = 0; i < formHeaders.length; i++) {
      if (norm(String(formHeaders[i])) === norm(name)) return i + 1;
    }
    return 0;
  };

  var fStudentIdCol = formGetCol('Student ID');
  var fParentIdCol  = formGetCol('Parent ID');
  if (!fStudentIdCol || !fParentIdCol) throw new Error('Student ID or Parent ID column not found in form sheet.');

  // === PART 2: Billing sheet ===
  var billingSS = UtilityScriptLibrary.getWorkbook('billing');
  var billingSheet = billingSS.getSheetByName('May 2026');
  if (!billingSheet) throw new Error('May 2026 billing sheet not found.');

  var billingHeaders = billingSheet.getRange(1, 1, 1, billingSheet.getLastColumn()).getValues()[0];
  var billGetCol = function(name) {
    for (var i = 0; i < billingHeaders.length; i++) {
      if (norm(String(billingHeaders[i])) === norm(name)) return i + 1;
    }
    return 0;
  };

  var bStudentIdCol = billGetCol('Student ID');
  var bParentIdCol  = billGetCol('Parent ID');
  if (!bStudentIdCol || !bParentIdCol) throw new Error('Student ID or Parent ID column not found in billing sheet.');

  // === PART 3: Build Student ID -> Parent ID map from contacts ===
  var contactsSheet = UtilityScriptLibrary.getSheet('students');
  var contactsData  = contactsSheet.getDataRange().getValues();
  var contactHeaders = contactsData[0];
  var cGetCol = function(name) {
    for (var i = 0; i < contactHeaders.length; i++) {
      if (norm(String(contactHeaders[i])) === norm(name)) return i + 1;
    }
    return 0;
  };

  var cStudentIdCol = cGetCol('Student ID');
  var cParentIdCol  = cGetCol('Parent ID');
  if (!cStudentIdCol || !cParentIdCol) throw new Error('Student ID or Parent ID column not found in contacts sheet.');

  var studentToParent = {};
  for (var i = 1; i < contactsData.length; i++) {
    var sid = String(contactsData[i][cStudentIdCol - 1] || '').trim();
    var pid = String(contactsData[i][cParentIdCol - 1] || '').trim();
    if (sid && pid) studentToParent[sid] = pid;
  }

  // === PART 4: Backfill form responses sheet ===
  var formData = formSheet.getDataRange().getValues();
  var formUpdated = 0;
  var formMissed = [];
  for (var r = 1; r < formData.length; r++) {
    var sid = String(formData[r][fStudentIdCol - 1] || '').trim();
    var pid = String(formData[r][fParentIdCol - 1] || '').trim();
    if (sid && !pid) {
      var foundPid = studentToParent[sid];
      if (foundPid) {
        formSheet.getRange(r + 1, fParentIdCol).setValue(foundPid);
        formUpdated++;
      } else {
        formMissed.push('Row ' + (r + 1) + ': Student ID ' + sid);
      }
    }
  }

  // === PART 5: Backfill billing sheet ===
  var billingData = billingSheet.getDataRange().getValues();
  var billUpdated = 0;
  var billMissed = [];
  for (var r = 1; r < billingData.length; r++) {
    var sid = String(billingData[r][bStudentIdCol - 1] || '').trim();
    var pid = String(billingData[r][bParentIdCol - 1] || '').trim();
    if (sid && !pid) {
      var foundPid = studentToParent[sid];
      if (foundPid) {
        billingSheet.getRange(r + 1, bParentIdCol).setValue(foundPid);
        billUpdated++;
      } else {
        billMissed.push('Row ' + (r + 1) + ': Student ID ' + sid);
      }
    }
  }

  // === REPORT ===
  var msg = 'Form sheet: ' + formUpdated + ' updated.\n' +
            'Billing sheet: ' + billUpdated + ' updated.\n';
  if (formMissed.length > 0) msg += '\nForm rows with no match:\n' + formMissed.join('\n');
  if (billMissed.length > 0) msg += '\nBilling rows with no match:\n' + billMissed.join('\n');

  Logger.log(msg);
  console.log(msg);
}

function authorizeScript() {
  // This function just needs to be run once to authorize the script
  // It accesses the UI which triggers the authorization prompt
  try {
    SpreadsheetApp.getUi();
    SpreadsheetApp.getActiveSpreadsheet();
    UtilityScriptLibrary.debugLog('authorizeScript', 'SUCCESS', 'Script authorized successfully', '', '');
  } catch (e) {
    UtilityScriptLibrary.debugLog('authorizeScript', 'ERROR', 'Authorization error', '', e.message);
  }
}

function handleFormEdit(e) {
  if (!e) {
    UtilityScriptLibrary.debugLog('handleFormEdit', 'WARNING', 'Called without event object', '', '');
    return;
  }

  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  // Get current semester from Calendar D2
  var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');
  if (!calendarSheet) {
    UtilityScriptLibrary.debugLog('handleFormEdit', 'ERROR', 'Calendar sheet not found', '', '');
    return;
  }

  var currentSemester = calendarSheet.getRange(2, 4).getValue();
  if (!currentSemester || String(currentSemester).trim() === '') {
    UtilityScriptLibrary.debugLog('handleFormEdit', 'WARNING', 'No current semester in Calendar D2', '', '');
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
    UtilityScriptLibrary.debugLog('handleFormEdit', 'ERROR', 'Teacher column not found in sheet', sheetName, '');
    return;
  }

  if (editedCol !== teacherCol) return;
  if (editedRow < 2) return;

  if (!shouldProcessEdit(e, headerMap)) {
    return;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    UtilityScriptLibrary.debugLog('handleFormEdit', 'WARNING', 'Could not obtain lock - another execution in progress', 'Sheet: ' + sheetName + ', Row: ' + editedRow, '');
    return;
  }

  try {
    UtilityScriptLibrary.debugLog('handleFormEdit', 'INFO', 'Processing row', 'Sheet: ' + sheetName + ', Row: ' + editedRow, '');

    processSingleRow(sheet, editedRow, headerMap);

    Browser.msgBox("🎉 Student successfully added! You can add the next student now.");

    UtilityScriptLibrary.debugLog('handleFormEdit', 'SUCCESS', 'Completed', 'Sheet: ' + sheetName + ', Row: ' + editedRow, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('handleFormEdit', 'ERROR', 'Failed', 'Sheet: ' + sheetName + ', Row: ' + editedRow, error.message);
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
    .addItem('Send Teacher Assignment Emails', 'sendTeacherAssignmentEmailsUI')
    .addItem('Reassign Student to Different Teacher', 'reassignStudentToNewTeacher')
    .addItem('Clear Reports', 'clearReports')
    .addItem('Verify by Drive ID', 'verifyByDriveIdWithPrompt')
    .addItem('Create New Year Workbooks with Continuing Students', 'createNewYearWorkbooksWithContinuingStudents')
    .addItem('Generate Enrollment Comparison Graph', 'generateEnrollmentComparisonGraph')
    .addToUi();
}

function addCarryoverStudentsToNewRoster(spreadsheet, newRosterSheet, currentSemesterName) {
  try {
    UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'INFO', 'Starting', 'Semester: ' + currentSemesterName, '');

    var previousRosterSheetName = findPreviousSemesterRoster(spreadsheet, currentSemesterName);
    if (!previousRosterSheetName) {
      UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'INFO', 'No previous semester roster found', currentSemesterName, '');
      return 0;
    }

    var previousRosterSheet = spreadsheet.getSheetByName(previousRosterSheetName);
    if (!previousRosterSheet) {
      UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'ERROR', 'Previous roster sheet not found', previousRosterSheetName, '');
      return 0;
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(previousRosterSheet);
    var data = previousRosterSheet.getDataRange().getValues();

    if (data.length <= 1) {
      UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'INFO', 'Previous roster has no student data', previousRosterSheetName, '');
      return 0;
    }

    var statusCol = headerMap['status'];
    var lessonsRemainingCol = headerMap['lessonsremaining'];
    var studentIdCol = headerMap['studentid'];

    if (!statusCol || !lessonsRemainingCol || !studentIdCol) {
      UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'ERROR', 'Required columns not found in previous roster',
        'Status: ' + statusCol + ', LessonsRemaining: ' + lessonsRemainingCol + ', StudentID: ' + studentIdCol, '');
      return 0;
    }

    var studentsToCarryOver = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = row[statusCol - 1];
      var lessonsRemaining = parseFloat(row[lessonsRemainingCol - 1]) || 0;
      var studentId = row[studentIdCol - 1];

      if (status && ['active', 'carryover', 'transferred'].indexOf(status.toString().toLowerCase()) !== -1 && lessonsRemaining > 0 && studentId) {
        studentsToCarryOver.push({
          rowIndex: i + 1,
          rowData: row,
          studentId: studentId,
          lessonsRemaining: lessonsRemaining
        });
      }
    }

    if (studentsToCarryOver.length === 0) {
      UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'INFO', 'No active students with lessons remaining to carry over', previousRosterSheetName, '');
      return 0;
    }

    UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'INFO', 'Students to carry over',
      'Count: ' + studentsToCarryOver.length + ', From: ' + previousRosterSheetName, '');

    var newHeaderMap = UtilityScriptLibrary.getHeaderMap(newRosterSheet);
    var newStatusCol = newHeaderMap['status'];

    var addedCount = 0;
    var fieldsToMap = [
      'contacted', 'firstlessondate', 'firstlessontime', 'comments',
      'lastname', 'firstname', 'instrument', 'length', 'experience',
      'grade', 'school', 'schoolteacher',
      'parentlastname', 'parentfirstname', 'phone', 'email', 'additionalcontacts',
      'hoursremaining', 'lessonsremaining', 'status', 'studentid', 'admincomments', 'systemcomments'
    ];

    for (var i = 0; i < studentsToCarryOver.length; i++) {
      var student = studentsToCarryOver[i];
      try {
        var alreadyExists = checkIfStudentExists(newRosterSheet, student.studentId, newHeaderMap);
        if (alreadyExists) {
          UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'DEBUG', 'Student already exists, skipping', student.studentId, '');
          continue;
        }

        var newRowData = [];
        for (var j = 0; j < 23; j++) newRowData[j] = '';

        for (var k = 0; k < fieldsToMap.length; k++) {
          var fieldName = fieldsToMap[k];
          var oldColIndex = headerMap[fieldName];
          var newColIndex = newHeaderMap[fieldName];
          if (oldColIndex && newColIndex) newRowData[newColIndex - 1] = student.rowData[oldColIndex - 1];
        }

        newRowData[newStatusCol - 1] = 'carryover';

        var systemCommentsCol = newHeaderMap['systemcomments'];
        if (systemCommentsCol) {
          var oldComments = newRowData[systemCommentsCol - 1] || '';
          newRowData[systemCommentsCol - 1] = 'Carried over from ' + previousRosterSheetName + ' on ' + UtilityScriptLibrary.formatDateFlexible(new Date(), 'M/d/yy') + '. ' + oldComments;
        }

        var targetRow = newRosterSheet.getLastRow() + 1;
        newRosterSheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
        newRosterSheet.getRange(targetRow, 1).insertCheckboxes().setValue(false);
        newRosterSheet.getRange(targetRow, 1, 1, 23)
          .setBackground(UtilityScriptLibrary.STYLES.WARNING.background)
          .setFontColor(UtilityScriptLibrary.STYLES.WARNING.text)
          .setFontWeight('bold');

        addedCount++;
        UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'DEBUG', 'Added carryover student',
          'Student: ' + student.studentId + ', Row: ' + targetRow, '');

      } catch (studentError) {
        UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'ERROR', 'Failed to add student',
          student.studentId, studentError.message);
      }
    }

    UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'SUCCESS', 'Completed',
      'Added: ' + addedCount + ' of ' + studentsToCarryOver.length, '');
    return addedCount;

  } catch (error) {
    UtilityScriptLibrary.debugLog('addCarryoverStudentsToNewRoster', 'ERROR', 'Failed', currentSemesterName, error.message);
    return 0;
  }
}

function addRosterTemplateBorders(sheet) {
  try {
    var maxRows = sheet.getMaxRows();
    sheet.getRange(1, 4, maxRows, 1).setBorder(null, null, null, true, null, null,
      UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(1, 8, maxRows, 1).setBorder(null, null, null, true, null, null,
      UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.DOTTED);
    sheet.getRange(1, 17, maxRows, 1).setBorder(null, null, null, true, null, null,
      UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
  } catch (error) {
    UtilityScriptLibrary.debugLog('addRosterTemplateBorders', 'ERROR', 'Failed', sheet.getName(), error.message);
  }
}

function addStudentToAttendanceSheet(attendanceSheet, studentData) {
  try {
    UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'INFO', 'Adding single student', attendanceSheet.getName(), '');

    var lastRowBefore = attendanceSheet.getLastRow();
    UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'INFO', 'Sheet state before', 'lastRow=' + lastRowBefore, '');

    // Update date validation starting at row 3 to preserve sign-off row
    var maxRows = attendanceSheet.getMaxRows();
    var dateColumn = attendanceSheet.getRange(3, 3, maxRows - 2, 1);
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .build();
    dateColumn.setDataValidation(dateRule);

    var student = createStudentObjectForAttendance(studentData);
    UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'INFO', 'Student object created', JSON.stringify(student), '');

    UtilityScriptLibrary.createStudentSections(attendanceSheet, [student]);

    var lastRowAfter = attendanceSheet.getLastRow();
    UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'INFO', 'Sheet state after', 'lastRow=' + lastRowAfter, '');

    UtilityScriptLibrary.setupStatusValidation(attendanceSheet, lastRowAfter);

    if (lastRowAfter > lastRowBefore) {
      UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'SUCCESS', 'Student added', student.firstName + ' ' + student.lastName, '');
    } else {
      UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'WARNING', 'lastRow did not increase - data may not have been written', '', '');
    }

  } catch (error) {
    UtilityScriptLibrary.debugLog('addStudentToAttendanceSheet', 'ERROR', 'Failed to add student', attendanceSheet.getName(), error.message);
    throw error;
  }
}

function addStudentToAttendanceSheetsFromDate(workbook, studentInfo, effectiveDate) {
  try {
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheetsFromDate", "INFO",
                                  "Starting attendance addition from date",
                                  "Student: " + studentInfo.studentId + ", Date: " + UtilityScriptLibrary.formatDateFlexible(effectiveDate, 'M/d/yy'), "");
    
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

function addStudentToRosterFromData(rosterSheet, studentInfo, headerMap) {
  try {
    var norm = UtilityScriptLibrary.normalizeHeader;
    var numCols = rosterSheet.getLastColumn();
    var newRowData = [];
    for (var i = 0; i < numCols; i++) newRowData[i] = '';

    var contacted        = headerMap[norm('Contacted')];
    var firstLessonDate  = headerMap[norm('First Lesson Date')];
    var firstLessonTime  = headerMap[norm('First Lesson Time')];
    var comments         = headerMap[norm('Comments')];
    var lastName         = headerMap[norm('Last Name')];
    var firstName        = headerMap[norm('First Name')];
    var instrument       = headerMap[norm('Instrument')];
    var length           = headerMap[norm('Length')];
    var experience       = headerMap[norm('Experience')];
    var grade            = headerMap[norm('Grade')];
    var school           = headerMap[norm('School')];
    var schoolTeacher    = headerMap[norm('School Teacher')];
    var parentLastName   = headerMap[norm('Parent Last Name')];
    var parentFirstName  = headerMap[norm('Parent First Name')];
    var phone            = headerMap[norm('Phone')];
    var email            = headerMap[norm('Email')];
    var additionalContacts = headerMap[norm('Additional contacts')];
    var hoursRemaining   = headerMap[norm('Hours Remaining')];
    var lessonsRemaining = headerMap[norm('Lessons Remaining')];
    var status           = headerMap[norm('Status')];
    var studentId        = headerMap[norm('Student ID')];
    var systemComments   = headerMap[norm('System Comments')];

    if (contacted)        newRowData[contacted - 1]        = false;
    if (firstLessonDate)  newRowData[firstLessonDate - 1]  = studentInfo.firstLessonDate || '';
    if (firstLessonTime)  newRowData[firstLessonTime - 1]  = studentInfo.firstLessonTime || '';
    if (comments)         newRowData[comments - 1]         = studentInfo.comments || '';
    if (lastName)         newRowData[lastName - 1]         = studentInfo.lastName || '';
    if (firstName)        newRowData[firstName - 1]        = studentInfo.firstName || '';
    if (instrument)       newRowData[instrument - 1]       = studentInfo.instrument || '';
    if (length)           newRowData[length - 1]           = studentInfo.length || 30;
    if (experience)       newRowData[experience - 1]       = studentInfo.experience || '';
    if (grade)            newRowData[grade - 1]            = studentInfo.grade || '';
    if (school)           newRowData[school - 1]           = studentInfo.school || '';
    if (schoolTeacher)    newRowData[schoolTeacher - 1]    = studentInfo.schoolTeacher || '';
    if (parentLastName)   newRowData[parentLastName - 1]   = studentInfo.parentLastName || '';
    if (parentFirstName)  newRowData[parentFirstName - 1]  = studentInfo.parentFirstName || '';
    if (phone)            newRowData[phone - 1]            = studentInfo.phone || '';
    if (email)            newRowData[email - 1]            = studentInfo.email || '';
    if (additionalContacts) newRowData[additionalContacts - 1] = studentInfo.additionalContacts || '';
    if (hoursRemaining)   newRowData[hoursRemaining - 1]   = studentInfo.hoursRemaining || 0;
    if (lessonsRemaining) newRowData[lessonsRemaining - 1] = studentInfo.lessonsRemaining || 0;
    if (status)           newRowData[status - 1]           = 'active';
    if (studentId)        newRowData[studentId - 1]        = studentInfo.studentId || '';
    if (systemComments)   newRowData[systemComments - 1]   = studentInfo.systemComment || '';

    // Find empty row or append
    var lastRow = rosterSheet.getLastRow();
    var targetRow = lastRow + 1;

    if (lastRow >= 2) {
      var existingData = rosterSheet.getRange(2, 1, lastRow - 1, numCols).getValues();
      for (var i = 0; i < existingData.length; i++) {
        var row = existingData[i];
        var isEmpty = true;
        for (var j = 0; j < row.length; j++) {
          if (row[j] !== '' && row[j] !== null && row[j] !== undefined) {
            isEmpty = false;
            break;
          }
        }
        if (isEmpty) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow <= lastRow) {
      rosterSheet.getRange(targetRow, 1, 1, numCols).setValues([newRowData]);
    } else {
      rosterSheet.appendRow(newRowData);
      targetRow = rosterSheet.getLastRow();
    }

    rosterSheet.getRange(targetRow, 1).insertCheckboxes().setValue(false);

    if (targetRow % 2 === 0) {
      rosterSheet.getRange(targetRow, 1, 1, numCols).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      rosterSheet.getRange(targetRow, 1, 1, numCols).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }

    UtilityScriptLibrary.debugLog('addStudentToRosterFromData', 'SUCCESS', 'Added student to roster',
      'Student: ' + studentInfo.studentId + ', Row: ' + targetRow, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('addStudentToRosterFromData', 'ERROR', 'Failed',
      'Student: ' + (studentInfo.studentId || 'unknown'), error.message);
    throw error;
  }
}

function addStudentToSemesterRoster(workbook, formData, studentId, semesterName) {
  try {
    UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'INFO', 'Starting',
      'Student: ' + studentId + ', Semester: ' + semesterName, '');

    var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
    if (!season) {
      throw new Error('Could not extract season from semester name: ' + semesterName);
    }

    var rosterSheetName = season + ' Roster';
    var rosterSheet = workbook.getSheetByName(rosterSheetName);
    var isNewSheet = false;

    if (!rosterSheet) {
      UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'INFO', 'Creating missing season roster sheet', rosterSheetName, '');
      try {
        rosterSheet = workbook.insertSheet(rosterSheetName);
        setupNewRosterTemplate(rosterSheet);
        isNewSheet = true;
      } catch (createError) {
        throw new Error("Failed to create season roster sheet '" + rosterSheetName + "': " + createError.message);
      }
    } else {
      UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'DEBUG', 'Using existing season roster sheet', rosterSheetName, '');
    }

    if (isNewSheet) {
      try {
        var carryoverCount = addCarryoverStudentsToNewRoster(workbook, rosterSheet, semesterName);
        UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'INFO', 'Carryover complete',
          'Count: ' + carryoverCount, '');
      } catch (carryoverError) {
        UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'WARNING', 'Error adding carryover students',
          rosterSheetName, carryoverError.message);
      }
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(rosterSheet);

    try {
      var existsResult = checkIfStudentExists(rosterSheet, studentId, headerMap);

      if (existsResult === 'CARRYOVER') {
        UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'INFO', 'Converting carryover to active', studentId, '');
        convertCarryoverToActive(rosterSheet, studentId, formData, headerMap);
        return;
      } else if (existsResult === true) {
        UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'INFO', 'Student already exists in roster', studentId, '');
        return;
      }
    } catch (checkError) {
      UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'WARNING', 'Could not check for existing student',
        studentId, checkError.message);
    }

    try {
      var lessonLength = '';
      var lengthFromField = formData['Length'];
      if (lengthFromField && lengthFromField.toString().trim() !== '') {
        lessonLength = UtilityScriptLibrary.extractNumericLessonLength(lengthFromField);
      } else {
        var qty60 = UtilityScriptLibrary.extractLessonQuantityFromPackage(formData['Qty60']);
        var qty45 = UtilityScriptLibrary.extractLessonQuantityFromPackage(formData['Qty45']);
        var qty30 = UtilityScriptLibrary.extractLessonQuantityFromPackage(formData['Qty30']);
        lessonLength = (qty60 > 0) ? 60 : (qty45 > 0) ? 45 : (qty30 > 0) ? 30 : 30;
      }

      var ageValue = formData['Age'] || '';
      var gradeValue = (ageValue.toString().toLowerCase().indexOf('yes') === 0) ? 'Adult' : (formData['Grade'] || '');

      var studentInfo = {
        studentId:         studentId,
        lastName:          formData['Student Last Name'] || '',
        firstName:         formData['Student First Name'] || '',
        instrument:        formData['Instrument'] || '',
        length:            lessonLength,
        experience:        formData['Experience'] || '',
        grade:             gradeValue,
        school:            formData['School'] || '',
        schoolTeacher:     formData['SchoolTeacher'] || '',
        parentLastName:    formData['Parent Last Name'] || '',
        parentFirstName:   formData['Parent First Name'] || '',
        phone:             formData['Phone'] || '',
        email:             formData['Email'] || '',
        additionalContacts: formData['Additional contacts'] || '',
        hoursRemaining:    0,
        lessonsRemaining:  0,
        systemComment:     'Added: ' + UtilityScriptLibrary.formatDateFlexible(new Date(), 'M/d')
      };

      addStudentToRosterFromData(rosterSheet, studentInfo, headerMap);
    } catch (addError) {
      throw new Error('Failed to add student to roster: ' + addError.message);
    }

    UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'SUCCESS', 'Completed',
      'Student: ' + studentId + ', Roster: ' + rosterSheetName, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('addStudentToSemesterRoster', 'ERROR', 'Failed',
      'Student: ' + studentId + ', Semester: ' + semesterName, error.message);
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
    
    UtilityScriptLibrary.debugLog('appendToReports', 'SUCCESS', 'Appended to reports', '', '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('appendToReports', 'ERROR', 'Failed to append to reports', '', error.message);
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
    if (!experience) return '';

    var currentYear = new Date().getFullYear();
    var exp = String(experience).toLowerCase().trim();

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
    UtilityScriptLibrary.debugLog('calculateExperienceStartRange', 'ERROR', 'Failed', String(experience), error.message);
    return '';
  }
}

function checkIfStudentExists(rosterSheet, studentId, headerMap) {
  try {
    var studentIdCol = headerMap['studentid'];
    if (!studentIdCol) {
      UtilityScriptLibrary.debugLog('checkIfStudentExists', 'ERROR', 'Student ID column not found in roster',
        rosterSheet.getName(), '');
      return false;
    }

    var rosterData = rosterSheet.getDataRange().getValues();
    var statusCol = headerMap['status'];

    for (var i = 1; i < rosterData.length; i++) {
      var existingId = rosterData[i][studentIdCol - 1];
      if (existingId && existingId.toString().trim() === studentId.toString().trim()) {
        if (statusCol && rosterData[i][statusCol - 1] && rosterData[i][statusCol - 1].toString() === 'carryover') {
          UtilityScriptLibrary.debugLog('checkIfStudentExists', 'DEBUG', 'Found carryover student', studentId, '');
          return 'CARRYOVER';
        }
        UtilityScriptLibrary.debugLog('checkIfStudentExists', 'DEBUG', 'Student already exists', studentId, '');
        return true;
      }
    }

    return false;

  } catch (error) {
    UtilityScriptLibrary.debugLog('checkIfStudentExists', 'ERROR', 'Failed', studentId, error.message);
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
      
      if (normalizedHeader === norm('Student ID')) {
        idCol = h;
      } else if (idCol === -1 && normalizedHeader === norm('ID')) {
        idCol = h;
      }
      
      if (normalizedHeader === norm('Student First Name')) {
        firstNameCol = h;
      } else if (firstNameCol === -1 && normalizedHeader === norm('First Name')) {
        firstNameCol = h;
      }
      
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
      
      if (!foundId && !firstName && !lastName) {
        continue;
      }
      
      if (foundId && foundId.charAt(0) !== 'Q') {
        continue;
      }
      
      var issue = null;
      
      if (foundId.charAt(0) === 'Q' && (!firstName || !lastName)) {
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
      UtilityScriptLibrary.debugLog('checkSheet', 'WARNING', 'Issues found in sheet', sheetName + ': ' + sheetIssues.length, '');
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('checkSheet', 'ERROR', 'Failed to check sheet', workbookName + ' - ' + sheetName, error.message);
  }
}

function checkWorkbook(workbook, workbookName, studentMap, detailIssues, summaryData) {
  try {
    var sheets = workbook.getSheets();
    UtilityScriptLibrary.debugLog('checkWorkbook', 'INFO', 'Checking sheets', workbookName + ': ' + sheets.length + ' sheets', '');
    
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      checkSheet(workbook, workbookName, sheet, studentMap, detailIssues, summaryData);
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('checkWorkbook', 'ERROR', 'Failed to check workbook', workbookName, error.message);
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
        UtilityScriptLibrary.debugLog('checkWorkbooksInFolder', 'INFO', 'Skipping excluded workbook', workbookName, '');
        continue;
      }
      
      try {
        UtilityScriptLibrary.debugLog('checkWorkbooksInFolder', 'INFO', 'Opening workbook', workbookName, '');
        var workbook = SpreadsheetApp.openById(file.getId());
        checkWorkbook(workbook, workbookName, studentMap, detailIssues, summaryData);
        processedCount++;
        
        if (processedCount % 5 === 0) {
          Utilities.sleep(1000);
          UtilityScriptLibrary.debugLog('checkWorkbooksInFolder', 'INFO', 'Pausing briefly', 'Processed: ' + processedCount + ' workbooks', '');
        }
        
      } catch (error) {
        UtilityScriptLibrary.debugLog('checkWorkbooksInFolder', 'WARNING', 'Could not access workbook', workbookName, error.message);
        continue;
      }
    }
    
    UtilityScriptLibrary.debugLog('checkWorkbooksInFolder', 'SUCCESS', 'Folder check complete', 'Processed: ' + processedCount + ' workbooks', '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('checkWorkbooksInFolder', 'ERROR', 'Failed', '', error.message);
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
    
    UtilityScriptLibrary.debugLog('clearReports', 'SUCCESS', 'Reports cleared', '', '');
    Browser.msgBox('Reports cleared. Ready for new verification run.');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('clearReports', 'ERROR', 'Failed to clear reports', '', error.message);
    Browser.msgBox('Error: ' + error.message);
  }
}

function convertCarryoverToActive(rosterSheet, studentId, formData, headerMap) {
  try {
    UtilityScriptLibrary.debugLog('convertCarryoverToActive', 'INFO', 'Starting', studentId, '');

    var data = rosterSheet.getDataRange().getValues();
    var studentIdCol = headerMap['studentid'];

    var studentRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol - 1] && data[i][studentIdCol - 1].toString().trim() === studentId.toString().trim()) {
        studentRow = i + 1;
        break;
      }
    }

    if (studentRow === -1) {
      throw new Error('Could not find student row for ID: ' + studentId);
    }

    var newLessonsRegistered = parseInt(formData['Lesson Quantity']) || 0;

    var lessonLength = 30;
    if (formData['Qty60']) lessonLength = 60;
    else if (formData['Qty45']) lessonLength = 45;

    var ageValue = formData['Age'] || '';
    var gradeValue = (ageValue.toString().toLowerCase().charAt(0) === 'y') ? 'Adult' : (formData['Grade'] || '');

    var updates = {
      'length':           lessonLength,
      'experience':       formData['Experience'] || '',
      'grade':            gradeValue,
      'school':           formData['School'] || '',
      'schoolteacher':    formData['SchoolTeacher'] || '',
      'parentlastname':   formData['Parent Last Name'] || '',
      'parentfirstname':  formData['Parent First Name'] || '',
      'phone':            formData['Phone'] || '',
      'email':            formData['Email'] || '',
      'additionalcontacts': formData['Additional contacts'] || '',
      'hoursremaining':   0,
      'lessonsremaining': 0,
      'status':           'active'
    };

    for (var fieldName in updates) {
      var colIndex = headerMap[fieldName];
      if (colIndex) rosterSheet.getRange(studentRow, colIndex).setValue(updates[fieldName]);
    }

    var systemCommentsCol = headerMap['systemcomments'];
    if (systemCommentsCol) {
      var oldComments = data[studentRow - 1][systemCommentsCol - 1] || '';
      rosterSheet.getRange(studentRow, systemCommentsCol).setValue(
        'Re-registered on ' + UtilityScriptLibrary.formatDateFlexible(new Date(), 'M/d/yy') + ' with ' + newLessonsRegistered + ' lessons. ' + oldComments
      );
    }

    var rowRange = rosterSheet.getRange(studentRow, 1, 1, 23);
    if (studentRow % 2 === 0) {
      rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }
    rowRange.setFontColor('#000000').setFontWeight('normal');

    UtilityScriptLibrary.debugLog('convertCarryoverToActive', 'SUCCESS', 'Converted',
      'Student: ' + studentId + ', Lessons: ' + newLessonsRegistered + ', Length: ' + lessonLength, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('convertCarryoverToActive', 'ERROR', 'Failed', studentId, error.message);
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
    status: 'active'
  };
}

function createInvoiceLogSheet(spreadsheet) {
  try {
    UtilityScriptLibrary.debugLog('createInvoiceLogSheet', 'INFO', 'Starting', spreadsheet.getName(), '');
    
    // Create the sheet
    var sheet = spreadsheet.insertSheet('Invoice Log');
    
    // Set up column headers
    setupInvoiceLogHeaders(sheet);
    
    // Apply formatting
    formatInvoiceLogSheet(sheet);
    
    UtilityScriptLibrary.debugLog('createInvoiceLogSheet', 'SUCCESS', 'Created', spreadsheet.getName(), '');
    return sheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('createInvoiceLogSheet', 'ERROR', 'Failed', spreadsheet.getName(), error.message);
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
    status: 'active'
  };
}

function enterEffectiveDate(newTeacherDisplay) {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();

    if (!newTeacherDisplay) {
      UtilityScriptLibrary.debugLog('enterEffectiveDate', 'INFO', 'User cancelled at new teacher prompt', '', '');
      return;
    }

    var oldTeacherDisplay = scriptProps.getProperty('reassign_oldTeacherDisplay');
    if (newTeacherDisplay === oldTeacherDisplay) {
      ui.alert('Error', 'New teacher cannot be the same as old teacher.', ui.ButtonSet.OK);
      return;
    }

    UtilityScriptLibrary.debugLog('enterEffectiveDate', 'INFO', 'New teacher selected', newTeacherDisplay, '');
    scriptProps.setProperty('reassign_newTeacherDisplay', newTeacherDisplay);

    var effectiveDateResponse = ui.prompt(
      'Reassign Students - Step 4 of 4',
      'Enter the effective date (MM/DD/YYYY):',
      ui.ButtonSet.OK_CANCEL
    );

    if (effectiveDateResponse.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('enterEffectiveDate', 'INFO', 'User cancelled at effective date prompt', '', '');
      return;
    }

    var effectiveDateStr = effectiveDateResponse.getResponseText().trim();
    var effectiveDate;
    try {
      effectiveDate = new Date(effectiveDateStr);
      if (isNaN(effectiveDate.getTime())) throw new Error('Invalid date');
    } catch (error) {
      ui.alert('Error', 'Invalid date format. Please use MM/DD/YYYY.', ui.ButtonSet.OK);
      return;
    }

    UtilityScriptLibrary.debugLog('enterEffectiveDate', 'INFO', 'Effective date set', UtilityScriptLibrary.formatDateFlexible(effectiveDate, 'M/d/yy'), '');
    scriptProps.setProperty('reassign_effectiveDate', effectiveDate.toISOString());

    processReassignment();

  } catch (error) {
    UtilityScriptLibrary.debugLog('enterEffectiveDate', 'ERROR', 'Failed', '', error.message);
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
    status: 'active',
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
    UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'INFO', 'Starting', currentSemesterName, '');

    var semesterSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterSheet) {
      UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'ERROR', 'Semester Metadata sheet not found', '', '');
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
      UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'ERROR', 'Required columns not found in Semester Metadata',
        'Name col: ' + nameCol + ', Start date col: ' + startDateCol, '');
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
      UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'WARNING', 'Current semester not found in metadata',
        currentSemesterName, '');
      return null;
    }

    var sheets = spreadsheet.getSheets();
    var existingRosterSheets = [];
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName.toLowerCase().indexOf('roster') !== -1) {
        existingRosterSheets.push(sheetName);
      }
    }

    var candidates = [];
    for (var i = 1; i < data.length; i++) {
      var semesterName = data[i][nameCol];
      var startDate = new Date(data[i][startDateCol]);
      if (startDate >= currentStartDate) continue;
      var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
      if (!season) continue;
      var expectedSheetName = season + ' Roster';
      for (var j = 0; j < existingRosterSheets.length; j++) {
        if (existingRosterSheets[j] === expectedSheetName) {
          candidates.push({ semesterName: semesterName, startDate: startDate, sheetName: expectedSheetName });
          break;
        }
      }
    }

    if (candidates.length === 0) {
      UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'INFO', 'No previous semester rosters found', spreadsheet.getName(), '');
      return null;
    }

    candidates.sort(function(a, b) { return b.startDate - a.startDate; });
    var previousSemester = candidates[0];
    UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'INFO', 'Found previous semester',
      previousSemester.semesterName + ' → ' + previousSemester.sheetName, '');
    return previousSemester.sheetName;

  } catch (error) {
    UtilityScriptLibrary.debugLog('findPreviousSemesterRoster', 'ERROR', 'Failed', currentSemesterName, error.message);
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

function generateEnrollmentComparisonGraph() {
  try {
    UtilityScriptLibrary.debugLog('generateEnrollmentComparisonGraph', 'INFO', 'Starting', '', '');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var MAY_15 = { month: 4, day: 15 };

    var sheet2025 = ss.getSheetByName('Summer 2025');
    var sheet2026 = ss.getSheetByName('Summer 2026');
    if (!sheet2025) throw new Error('Sheet "Summer 2025" not found.');
    if (!sheet2026) throw new Error('Sheet "Summer 2026" not found.');

    function getTimestamps(sheet) {
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      var floored = new Date(2000, MAY_15.month, MAY_15.day);
      return data
        .map(function(row) { return row[0]; })
        .filter(function(val) { return val instanceof Date && !isNaN(val); })
        .map(function(d) {
          var asNeutral = new Date(2000, d.getMonth(), d.getDate());
          return asNeutral < floored ? floored : asNeutral;
        })
        .sort(function(a, b) { return a - b; });
    }

    var timestamps2025 = getTimestamps(sheet2025);
    var timestamps2026 = getTimestamps(sheet2026);

    function buildCumulative(timestamps) {
      return timestamps.map(function(d, i) {
        return {
          label: UtilityScriptLibrary.formatDateFlexible(d),
          count: i + 1
        };
      });
    }

    var cum2025 = buildCumulative(timestamps2025);
    var cum2026 = buildCumulative(timestamps2026);

    var allLabels = [];
    var labelSet = {};
    cum2025.concat(cum2026).forEach(function(pt) {
      if (!labelSet[pt.label]) {
        labelSet[pt.label] = true;
        allLabels.push(pt.label);
      }
    });
    allLabels.sort(function(a, b) {
      return new Date('2000 ' + a) - new Date('2000 ' + b);
    });

    function mapToLabels(cum, labels) {
      var lookup = {};
      var lastLabel = null;
      cum.forEach(function(pt) {
        lookup[pt.label] = pt.count;
        lastLabel = pt.label;
      });
      var last = 0;
      var done = false;
      return labels.map(function(lbl) {
        if (done) return null;
        if (lookup[lbl] !== undefined) last = lookup[lbl];
        if (lbl === lastLabel) done = true;
        return last || null;
      });
    }

    var values2025 = mapToLabels(cum2025, allLabels);
    var values2026 = mapToLabels(cum2026, allLabels);

    var graphSheet = ss.getSheetByName('Graph');
    if (graphSheet) {
      graphSheet.clearContents();
      graphSheet.getCharts().forEach(function(c) { graphSheet.removeChart(c); });
    } else {
      graphSheet = ss.insertSheet('Graph');
    }

    graphSheet.getRange(1, 1, 1, 4).setValues([['Date', 'Summer 2025', 'Summer 2026', '% Change']]);

    var rows = allLabels.map(function(lbl, i) {
      var v25 = values2025[i];
      var v26 = values2026[i];
      var pct = (v25 && v26) ? (v26 / v25 - 1) : null;
      return [lbl, v25, v26, pct];
    });
    graphSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    graphSheet.getRange(2, 4, rows.length, 1).setNumberFormat('0.0%');

    var chart = graphSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(graphSheet.getRange(1, 1, rows.length + 1, 3))
      .setPosition(2, 5, 0, 0)
      .setOption('title', 'Summer Enrollment: 2025 vs 2026')
      .setOption('hAxis.title', '')
      .setOption('vAxis.title', 'Cumulative Registrations')
      .setOption('legend.position', 'top')
      .setOption('interpolateNulls', true)
      .setOption('width', 700)
      .setOption('height', 400)
      .build();

    graphSheet.insertChart(chart);

    UtilityScriptLibrary.debugLog('generateEnrollmentComparisonGraph', 'INFO', 'Complete',
      '2025: ' + timestamps2025.length + ' registrations, 2026: ' + timestamps2026.length + ' registrations', '');

    SpreadsheetApp.getUi().alert('Enrollment comparison graph generated successfully.');

  } catch (error) {
    UtilityScriptLibrary.debugLog('generateEnrollmentComparisonGraph', 'ERROR', 'Failed', error.message, error.stack);
    SpreadsheetApp.getUi().alert('Error generating graph: ' + error.message);
  }
}

function getActiveStudentsFromRoster(rosterSheet) {
  try {
    var headerMap = UtilityScriptLibrary.getHeaderMap(rosterSheet);

    var studentIdCol = headerMap['studentid'] || headerMap['student id'];
    var statusCol    = headerMap['status'];
    var firstNameCol = headerMap['firstname'] || headerMap['first name'];
    var lastNameCol  = headerMap['lastname'] || headerMap['last name'];
    var instrumentCol = headerMap['instrument'];

    if (!studentIdCol) {
      throw new Error('Student ID column not found in roster');
    }
    if (!firstNameCol || !lastNameCol || !instrumentCol) {
      UtilityScriptLibrary.debugLog('getActiveStudentsFromRoster', 'ERROR', 'Required columns missing',
        'FirstName: ' + firstNameCol + ', LastName: ' + lastNameCol + ', Instrument: ' + instrumentCol, '');
      throw new Error('Required columns (First Name, Last Name, or Instrument) not found in roster');
    }

    var data = rosterSheet.getDataRange().getValues();
    var students = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentId = row[studentIdCol - 1];
      var status = statusCol ? String(row[statusCol - 1] || '').trim().toLowerCase() : '';

      if (studentId && String(studentId).trim() !== '' &&
          (!status || status === '' || status === 'active' || status === 'carryover')) {
        var studentInfo = extractStudentDataFromRoster(row, headerMap);
        studentInfo.rowNumber = i + 1;
        students.push(studentInfo);
      }
    }

    UtilityScriptLibrary.debugLog('getActiveStudentsFromRoster', 'INFO', 'Extracted active students',
      'Sheet: ' + rosterSheet.getName() + ', Count: ' + students.length, '');
    return students;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getActiveStudentsFromRoster', 'ERROR', 'Failed',
      rosterSheet.getName(), error.message);
    throw error;
  }
}

function getActiveTeachersForDropdown() {
  try {
    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');

    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'WARNING', 'Teacher Roster Lookup sheet not found or empty', '', '');
      return [];
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var displayNameCol = getCol('Display Name');
    var statusCol      = getCol('Status');

    if (!displayNameCol || !statusCol) {
      UtilityScriptLibrary.debugLog('getActiveTeachersForDropdown', 'ERROR', 'Required columns not found', '', '');
      return [];
    }

    var validStatuses = ['potential', 'active', 'returning'];
    var displayNames = [];
    var data = lookupSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var status      = String(data[i][statusCol - 1]).trim().toLowerCase();
      var displayName = String(data[i][displayNameCol - 1]).trim();
      if (validStatuses.indexOf(status) === -1) continue;
      if (!displayName) continue;
      displayNames.push(displayName);
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
    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');

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

function getTeacherContactForAssignmentEmail(teacherId) {
  try {
    var norm = UtilityScriptLibrary.normalizeHeader;
    var sheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);

    var idCol        = headerMap[norm('Teacher ID')];
    var firstCol     = headerMap[norm('First Name')];
    var lastCol      = headerMap[norm('Last Name')];
    var salutCol     = headerMap[norm('Salutation')];
    var emailCol     = headerMap[norm('Email')];

    if (!idCol || !firstCol || !lastCol || !emailCol) {
      UtilityScriptLibrary.debugLog('getTeacherContactForAssignmentEmail', 'ERROR',
        'Required columns not found in Teachers and Admin', '', '');
      return null;
    }

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol - 1]).trim() === String(teacherId).trim()) {
        return {
          firstName:  String(data[i][firstCol - 1]  || '').trim(),
          lastName:   String(data[i][lastCol - 1]   || '').trim(),
          salutation: salutCol ? String(data[i][salutCol - 1] || '').trim() : '',
          email:      String(data[i][emailCol - 1]  || '').trim()
        };
      }
    }

    UtilityScriptLibrary.debugLog('getTeacherContactForAssignmentEmail', 'WARNING',
      'Teacher ID not found', teacherId, '');
    return null;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getTeacherContactForAssignmentEmail', 'ERROR',
      'Failed', teacherId, error.message);
    return null;
  }
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

function handleNewStudentFormSubmit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var currentSemester = UtilityScriptLibrary.getCurrentSemesterName();
  if (!currentSemester || sheet.getName() !== String(currentSemester).trim()) return;

  try {
    var row = e.range.getRow();
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var fieldMapSheet = UtilityScriptLibrary.getSheet('fieldMap');
    if (!fieldMapSheet) throw new Error('FieldMap sheet not found');
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
    var formData = extractFormData(sheet, row, headerMap, fieldMap);

    var email = String(formData['Email'] || '').trim();
    if (!email) {
      UtilityScriptLibrary.debugLog('handleNewStudentFormSubmit', 'WARNING', 'No email address found — skipping confirmation', '', '');
      return;
    }

    var studentFirst = String(formData['Student First Name'] || '').trim();
    var instrument   = String(formData['Instrument'] || '').trim();
    var isAdult      = String(formData['Age'] || '').toLowerCase().indexOf('yes') === 0;

    var title      = String(formData['Salutation'] || '').trim();
    var firstName  = String(formData['Parent First Name'] || '').trim();
    var lastName   = String(formData['Parent Last Name'] || '').trim();
    var greeting   = (title && lastName) ? 'Dear ' + title + ' ' + lastName : 'Dear ' + (firstName || 'Music Family');

    var registrationRef = isAdult
      ? 'your registration for ' + instrument + ' lessons'
      : 'We\'ve received ' + studentFirst + '\'s registration for ' + instrument + ' lessons';

    var subject = 'Quaker Arts Summer Music Lessons \u2014 Registration Received';
    var body = greeting + ',\n\n'
      + (isAdult ? 'Thank you for registering! We\'ve received ' + registrationRef + '.'
                 : 'Thank you for registering! ' + registrationRef + '.') + '\n\n'
      + 'Here\'s what happens next:\n'
      + '- You\'ll hear from your teacher by mid-June to arrange a lesson time and location\n'
      + '- Once scheduling is confirmed, you\'ll receive a billing statement\n\n'
      + 'If you have any questions in the meantime, don\'t hesitate to reach out.\n\n'
      + 'We look forward to a great summer!\n\n'
      + 'Warm regards,\n'
      + 'Cristin';

    UtilityScriptLibrary.sendEmail(email, subject, body);
    UtilityScriptLibrary.debugLog('handleNewStudentFormSubmit', 'SUCCESS', 'Confirmation email sent', email, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('handleNewStudentFormSubmit', 'ERROR', 'Failed', '', error.message);
  }
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
    
    UtilityScriptLibrary.debugLog('loadStudentMapFromContacts', 'SUCCESS', 'Loaded students from Contacts', Object.keys(studentMap).length + ' students', '');
    return studentMap;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('loadStudentMapFromContacts', 'ERROR', 'Failed to load student map', '', error.message);
    throw error;
  }
}

function logSheetHeaders() {
  UtilityScriptLibrary.logAllSheetHeaders();
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
        'carryover',                         // T - Status
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

    if (formData['Phone']) {
      formData['Phone'] = UtilityScriptLibrary.formatPhoneNumber(String(formData['Phone']));
    }

    var city = String(formData['City'] || '').trim().toLowerCase().split(' ').map(function(w) {
      return w.charAt(0).toUpperCase() + w.slice(1);
    }).join(' ');
    var zip  = String(formData['Zip'] || '').trim();
    formData['Address City']      = city;
    formData['Address Zip Code']  = zip;
    formData['Address Formatted'] = UtilityScriptLibrary.formatAddress(formData['Address Street'] || '', city, zip);

    var parentKey = UtilityScriptLibrary.generateKey(
      formData['Parent Last Name'] || '',
      formData['Parent First Name'] || '',
      formData['Email'] || ''
    );

    var parentIdCol    = getCol('Parent ID');
    var lookupCol      = getCol('Parent Lookup');
    var studentIdsCol  = getCol('Student IDs');
    var updatedCol     = getCol('Updated');
    var parentGroupCol = getCol('Parent Group Interest');

    var textFields = [
      'Parent Last Name', 'Parent First Name', 'Salutation', 'Email',
      'Phone', 'Address Formatted', 'Billing Preference', 'Additional contacts', 'Referral'
    ];

    var parentId  = existingParentId;
    var parentRow = UtilityScriptLibrary.findParentRow(parentsSheet, parentId, parentKey);

    UtilityScriptLibrary.debugLog('processParent', 'DEBUG', 'Duplicate check',
      'Parent ID: "' + parentId + '", Key: "' + parentKey + '", Found row: ' + parentRow, '');

    if (parentRow !== -1) {
      if (!parentId) {
        var foundRowValues = parentsSheet.getRange(parentRow, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
        parentId = String(foundRowValues[parentIdCol - 1] || '').trim();
        UtilityScriptLibrary.debugLog('processParent', 'DEBUG', 'Resolved parent ID from sheet', 'ID: ' + parentId, '');
      }

      UtilityScriptLibrary.debugLog('processParent', 'INFO', 'Updating existing parent', 'Parent ID: ' + parentId, '');

      var fieldsObj = {};
      for (var j = 0; j < textFields.length; j++) {
        fieldsObj[textFields[j]] = formData[textFields[j]] || '';
      }

      UtilityScriptLibrary.updateParentContactFields(parentsSheet, parentRow, fieldsObj, { newLookupKey: parentKey });

      if (studentIdsCol) {
        var rowValues = parentsSheet.getRange(parentRow, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
        var currentStudentIds = String(rowValues[studentIdsCol - 1] || '');
        var studentIdArray = currentStudentIds ? currentStudentIds.split(',').map(function(id) { return id.trim(); }) : [];
        if (studentIdArray.indexOf(studentId) === -1) {
          studentIdArray.push(studentId);
          parentsSheet.getRange(parentRow, studentIdsCol).setValue(studentIdArray.join(', '));
          UtilityScriptLibrary.debugLog('processParent', 'DEBUG', 'Added student to parent record', 'Student: ' + studentId + ', Parent: ' + parentId, '');
        }
      }

      if (parentGroupCol) {
        var parentGroupCell = parentsSheet.getRange(parentRow, parentGroupCol);
        parentGroupCell.insertCheckboxes();
        parentGroupCell.setValue(formData['Parent Group Interest'] === 'Yes');
      }

    } else {
      UtilityScriptLibrary.debugLog('processParent', 'INFO', 'Creating new parent',
        (formData['Parent First Name'] || '') + ' ' + (formData['Parent Last Name'] || ''), '');
      parentId = UtilityScriptLibrary.generateNextId(parentsSheet, 'Parent ID', 'P');

      var newRow = new Array(headers.length);
      for (var m = 0; m < headers.length; m++) newRow[m] = '';

      newRow[parentIdCol - 1]   = parentId;
      newRow[lookupCol - 1]     = parentKey;
      newRow[studentIdsCol - 1] = studentId;

      for (var n = 0; n < textFields.length; n++) {
        var field = textFields[n];
        var col = getCol(field);
        if (col) newRow[col - 1] = formData[field] || '';
      }

      parentsSheet.appendRow(newRow);
      var newRowIndex = parentsSheet.getLastRow();

      if (parentGroupCol) {
        var parentGroupCell = parentsSheet.getRange(newRowIndex, parentGroupCol);
        parentGroupCell.insertCheckboxes();
        parentGroupCell.setValue(formData['Parent Group Interest'] === 'Yes');
      }

      if (updatedCol) {
        var updatedCell = parentsSheet.getRange(newRowIndex, updatedCol);
        updatedCell.insertCheckboxes();
        updatedCell.setValue(true);
      }
    }

    UtilityScriptLibrary.debugLog('processParent', 'SUCCESS', 'Completed', 'ID: ' + parentId, '');
    return parentId;

  } catch (error) {
    UtilityScriptLibrary.debugLog('processParent', 'ERROR', 'Failed', 'Student: ' + studentId, error.message);
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
        UtilityScriptLibrary.debugLog('processPendingAssignments', 'INFO', 'Processed row', 'Row: ' + pendingRows[i], '');
      } catch (error) {
        errorCount++;
        errors.push('Row ' + pendingRows[i] + ': ' + error.message);
        UtilityScriptLibrary.debugLog('processPendingAssignments', 'ERROR', 'Failed to process row', 'Row: ' + pendingRows[i], error.message);
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
    UtilityScriptLibrary.debugLog('processPendingAssignments', 'ERROR', 'Batch processing failed', '', error.message);
  }
}

function processReassignment() {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();

    var currentSemester    = scriptProps.getProperty('reassign_currentSemester');
    var year               = scriptProps.getProperty('reassign_year');
    var oldTeacherDisplay  = scriptProps.getProperty('reassign_oldTeacherDisplay');
    var newTeacherDisplay  = scriptProps.getProperty('reassign_newTeacherDisplay');
    var oldTeacherInfo     = JSON.parse(scriptProps.getProperty('reassign_oldTeacherInfo'));
    var selectedStudents   = JSON.parse(scriptProps.getProperty('reassign_selectedStudents'));
    var effectiveDate      = new Date(scriptProps.getProperty('reassign_effectiveDate'));
    var oldRosterWorkbookId = scriptProps.getProperty('reassign_oldRosterWorkbookId');
    var oldRosterSheetName  = scriptProps.getProperty('reassign_oldRosterSheetName');

    UtilityScriptLibrary.debugLog('processReassignment', 'INFO', 'Starting',
      selectedStudents.length + ' students from ' + oldTeacherDisplay + ' to ' + newTeacherDisplay, '');

    var yearMetadataSheet = UtilityScriptLibrary.getSheet('yearMetadata');
    if (!yearMetadataSheet) throw new Error('Year Metadata sheet not found');

    var metadataRows = yearMetadataSheet.getDataRange().getValues();
    var headerRow = metadataRows[0];
    var yearColIndex = -1, folderIdColIndex = -1;

    for (var i = 0; i < headerRow.length; i++) {
      if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader('Year')) yearColIndex = i;
      if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader('Roster Folder ID')) folderIdColIndex = i;
    }

    if (yearColIndex === -1 || folderIdColIndex === -1) {
      throw new Error('Required columns not found in Year Metadata sheet');
    }

    var yearRow = null;
    for (var i = 0; i < metadataRows.length; i++) {
      if (metadataRows[i][yearColIndex] && metadataRows[i][yearColIndex].toString() === year) {
        yearRow = metadataRows[i];
        break;
      }
    }
    if (!yearRow) throw new Error('No roster folder found for year: ' + year);

    var rosterFolder = DriveApp.getFolderById(yearRow[folderIdColIndex]);
    UtilityScriptLibrary.debugLog('processReassignment', 'DEBUG', 'Found roster folder', 'Year: ' + year, '');

    var newTeacherInfo = UtilityScriptLibrary.getTeacherInfoByDisplayName(newTeacherDisplay);
    if (!newTeacherInfo) throw new Error('Could not find new teacher info in Teacher Roster Lookup');

    var newRosterWorkbook = getOrCreateRosterFromTemplate(newTeacherInfo, rosterFolder, year, currentSemester);
    var newRosterSheet = findSemesterRoster(newRosterWorkbook, currentSemester);
    if (!newRosterSheet) throw new Error('Could not find or create roster sheet for ' + currentSemester + ' in new teacher workbook');

    var newRosterFile = DriveApp.getFileById(newRosterWorkbook.getId());
    updateTeacherRosterLookup(newTeacherInfo.teacherId, newRosterFile.getUrl());

    var oldRosterWorkbook = SpreadsheetApp.openById(oldRosterWorkbookId);
    var oldRosterSheet = oldRosterWorkbook.getSheetByName(oldRosterSheetName);
    if (!oldRosterSheet) throw new Error('Could not find old roster sheet: ' + oldRosterSheetName);

    var oldHeaderMap = UtilityScriptLibrary.getHeaderMap(oldRosterSheet);
    var newHeaderMap = UtilityScriptLibrary.getHeaderMap(newRosterSheet);

    var contactsSheet = UtilityScriptLibrary.getSheet('students');
    var contactsHeaderMap = contactsSheet ? UtilityScriptLibrary.getHeaderMap(contactsSheet) : null;
    var teacherCol    = contactsHeaderMap ? contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')] : null;
    var studentIdCol  = contactsHeaderMap ? contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Student ID')] : null;
    var contactsData  = contactsSheet ? contactsSheet.getDataRange().getValues() : null;

    var successCount = 0, skipCount = 0, errorCount = 0, errors = [];

    for (var i = 0; i < selectedStudents.length; i++) {
      var student = selectedStudents[i];
      try {
        var statusCol = oldHeaderMap['status'];
        oldRosterSheet.getRange(student.rowNumber, statusCol).setValue('transferred');

        var rowRange = oldRosterSheet.getRange(student.rowNumber, 1, 1, 23);
        if (student.rowNumber % 2 === 0) {
          rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background)
                  .setFontColor(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.text);
        } else {
          rowRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background)
                  .setFontColor(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.text);
        }

        var studentExists = checkIfStudentExists(newRosterSheet, student.studentId, newHeaderMap);
        if (studentExists) {
          skipCount++;
        } else {
          student.systemComment = 'Reassigned: ' + UtilityScriptLibrary.formatDateFlexible(new Date(), 'M/d');
          addStudentToRosterFromData(newRosterSheet, student, newHeaderMap);
        }

        addStudentToAttendanceSheetsFromDate(newRosterWorkbook, student, effectiveDate);

        if (contactsSheet && teacherCol && studentIdCol && contactsData) {
          for (var r = 1; r < contactsData.length; r++) {
            if (String(contactsData[r][studentIdCol - 1]).trim() === String(student.studentId).trim()) {
              contactsSheet.getRange(r + 1, teacherCol).setValue(newTeacherInfo.teacherId);
              break;
            }
          }
        }

        successCount++;
        UtilityScriptLibrary.debugLog('processReassignment', 'DEBUG', 'Processed student',
          student.firstName + ' ' + student.lastName + ' (' + student.studentId + ')', '');

      } catch (studentError) {
        errorCount++;
        errors.push(student.firstName + ' ' + student.lastName + ' (' + student.studentId + '): ' + studentError.message);
        UtilityScriptLibrary.debugLog('processReassignment', 'ERROR', 'Failed to process student',
          student.studentId, studentError.message);
      }
    }

    var keysToDelete = [
      'reassign_currentSemester', 'reassign_year', 'reassign_season',
      'reassign_activeTeachers', 'reassign_oldTeacherDisplay', 'reassign_oldTeacherInfo',
      'reassign_studentList', 'reassign_oldRosterSheetName', 'reassign_oldRosterWorkbookId',
      'reassign_selectedStudents', 'reassign_newTeacherDisplay', 'reassign_effectiveDate'
    ];
    for (var i = 0; i < keysToDelete.length; i++) scriptProps.deleteProperty(keysToDelete[i]);

    UtilityScriptLibrary.debugLog('processReassignment', 'SUCCESS', 'Completed',
      'Success: ' + successCount + ', Skipped: ' + skipCount + ', Errors: ' + errorCount, '');

    var message = 'Student Reassignment Complete!\n\nSuccessfully processed: ' + successCount + ' student(s)\n';
    if (skipCount > 0) message += 'Already in new roster (skipped): ' + skipCount + '\n';
    if (errorCount > 0) message += 'Errors: ' + errorCount + '\n\nFailed students:\n' + errors.join('\n');

    ui.alert('Success', message, ui.ButtonSet.OK);

  } catch (error) {
    UtilityScriptLibrary.debugLog('processReassignment', 'ERROR', 'Failed', '', error.message);
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

    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
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
  try {
    UtilityScriptLibrary.debugLog('processSingleRow', 'INFO', 'Starting', 'Sheet: ' + sheet.getName() + ', Row: ' + row, '');

    var fieldMapSheet = UtilityScriptLibrary.getSheet('fieldMap');
    if (!fieldMapSheet) throw new Error('FieldMap sheet not found.');

    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
    var formData = extractFormData(sheet, row, headerMap, fieldMap);

    var contactsSheet = UtilityScriptLibrary.getSheet('students');
    var studentResult = processStudent(formData, contactsSheet, sheet.getName());
    var studentId = studentResult.studentId;
    var studentRow = studentResult.studentRow;

    var idCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    if (!idCol) {
      idCol = headerMap[UtilityScriptLibrary.normalizeHeader('ID')];
      UtilityScriptLibrary.debugLog('processSingleRow', 'WARNING', 'Student ID column not found, falling back to ID column', 'Col: ' + idCol, '');
    }

    var parentsSheet = UtilityScriptLibrary.getSheet('parents');
    if (!parentsSheet) throw new Error('Parents sheet not found.');

    var existingParentId = studentResult.parentId || '';
    var parentId = processParent(formData, parentsSheet, studentId, existingParentId);
    updateStudentWithParentId(contactsSheet, studentRow, parentId);

    if (idCol) {
      sheet.getRange(row, idCol).setValue(studentId);
      UtilityScriptLibrary.debugLog('processSingleRow', 'DEBUG', 'Wrote Student ID to sheet', 'ID: ' + studentId + ', Row: ' + row + ', Col: ' + idCol, '');
    }

    var parentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Parent ID')];
    if (parentIdCol) {
      sheet.getRange(row, parentIdCol).setValue(parentId);
      UtilityScriptLibrary.debugLog('processSingleRow', 'DEBUG', 'Wrote Parent ID to sheet', 'ID: ' + parentId + ', Row: ' + row + ', Col: ' + parentIdCol, '');
    }

    // === ROSTER PROCESSING ===
    try {
      var sheetName = sheet.getName();
      var yearMatch = sheetName.match(/\d{4}/);
      if (!yearMatch) {
        UtilityScriptLibrary.debugLog('processSingleRow', 'WARNING', 'Could not extract year from sheet name', sheetName, '');
        return;
      }
      var year = yearMatch[0];
      var semesterName = sheetName;

      var yearMetadataSheet = UtilityScriptLibrary.getSheet('yearMetadata');
      if (!yearMetadataSheet) {
        UtilityScriptLibrary.debugLog('processSingleRow', 'ERROR', 'Year Metadata sheet not found', '', '');
        return;
      }

      var metadataRows = yearMetadataSheet.getDataRange().getValues();
      var headerRow = metadataRows[0];
      var yearColIndex = -1;
      var folderIdColIndex = -1;
      for (var i = 0; i < headerRow.length; i++) {
        if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader('Year')) yearColIndex = i;
        if (UtilityScriptLibrary.normalizeHeader(headerRow[i]) === UtilityScriptLibrary.normalizeHeader('Roster Folder ID')) folderIdColIndex = i;
      }

      if (yearColIndex === -1 || folderIdColIndex === -1) {
        UtilityScriptLibrary.debugLog('processSingleRow', 'ERROR', 'Required columns not found in Year Metadata', 'Year col: ' + yearColIndex + ', Folder ID col: ' + folderIdColIndex, '');
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
        UtilityScriptLibrary.debugLog('processSingleRow', 'ERROR', 'No roster folder found for year', year, '');
        return;
      }

      var rosterFolderId = yearRow[folderIdColIndex];
      var rosterFolder = DriveApp.getFolderById(rosterFolderId);
      UtilityScriptLibrary.debugLog('processSingleRow', 'DEBUG', 'Found roster folder', 'Year: ' + year, '');

      processRoster(formData, sheet, row, headerMap, fieldMap, studentId, studentResult.teacherId, rosterFolder, year, semesterName);

    } catch (rosterError) {
      UtilityScriptLibrary.debugLog('processSingleRow', 'ERROR', 'Roster processing failed', 'Row: ' + row, rosterError.message);
    }

    UtilityScriptLibrary.debugLog('processSingleRow', 'SUCCESS', 'Completed', 'Student: ' + studentId + ', Row: ' + row, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('processSingleRow', 'ERROR', 'Failed', 'Row: ' + row, error.message);
    throw error;
  }
}

function processStudent(formData, contactsSheet, enrollmentTerm) {
  try {
    var studentKey = UtilityScriptLibrary.generateKey(
      formData['Student Last Name'] || '',
      formData['Student First Name'] || '',
      formData['Instrument'] || ''
    );

    var studentRow = UtilityScriptLibrary.findStudentRow(contactsSheet, studentKey);
    UtilityScriptLibrary.debugLog('processStudent', 'DEBUG', 'Duplicate check',
      'Key: ' + studentKey + ', Found row: ' + studentRow, '');

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
      'Student Last Name', 'Student First Name', 'Instrument', 'Teacher ID',
      'Age', 'Currently Registered', 'Student ID', 'Parent ID', 'Student Lookup', 'First Enrollment Term'
    ];
    var missingFields = [];
    for (var i = 0; i < requiredFields.length; i++) {
      if (getCol(requiredFields[i]) === 0) missingFields.push(requiredFields[i]);
    }
    if (missingFields.length > 0) {
      throw new Error('Missing required columns in Students sheet: ' + missingFields.join(', '));
    }

    var teacherId = UtilityScriptLibrary.getTeacherIdByDisplayName(String(formData['Teacher'] || '').trim());
    if (!teacherId) {
      UtilityScriptLibrary.debugLog('processStudent', 'WARNING', 'Could not resolve Teacher ID from display name', formData['Teacher'] || '', '');
    }

    var studentId = formData['Student ID'] || '';
    var parentId = '';

    var ageResponse = formData['Age'] || '';
    var standardizedAge = (ageResponse.toString().toLowerCase().indexOf('yes') === 0) ? 'Adult' : 'Child';
    UtilityScriptLibrary.debugLog('processStudent', 'DEBUG', 'Age standardized',
      'Original: "' + ageResponse + '", Result: ' + standardizedAge, '');

    if (studentRow !== -1) {
      UtilityScriptLibrary.debugLog('processStudent', 'INFO', 'Updating existing student', studentKey, '');
      var rowData = contactsSheet.getRange(studentRow, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
      studentId = String(rowData[getCol('Student ID') - 1] || '');
      parentId = String(rowData[getCol('Parent ID') - 1] || '');

      var registeredCol = getCol('Currently Registered');
      if (registeredCol) {
        var checkboxCell = contactsSheet.getRange(studentRow, registeredCol);
        checkboxCell.insertCheckboxes();
        checkboxCell.setValue(true);
      }

      if (getCol('Teacher ID')) contactsSheet.getRange(studentRow, getCol('Teacher ID')).setValue(teacherId);
      if (getCol('Age')) contactsSheet.getRange(studentRow, getCol('Age')).setValue(standardizedAge);
      if (getCol('Graduation Year')) {
        contactsSheet.getRange(studentRow, getCol('Graduation Year')).setValue(UtilityScriptLibrary.calculateGraduationYear(formData['Grade']));
      }
      if (getCol('School District')) contactsSheet.getRange(studentRow, getCol('School District')).setValue(formData['School'] || '');
      if (getCol('School Teacher')) contactsSheet.getRange(studentRow, getCol('School Teacher')).setValue(formData['SchoolTeacher'] || '');
      if (getCol('Experience')) contactsSheet.getRange(studentRow, getCol('Experience')).setValue(formData['Experience'] || '');
      if (getCol('Experience Start Range')) contactsSheet.getRange(studentRow, getCol('Experience Start Range')).setValue(calculateExperienceStartRange(formData['Experience']));

    } else {
      UtilityScriptLibrary.debugLog('processStudent', 'INFO', 'Creating new student',
        (formData['Student First Name'] || '') + ' ' + (formData['Student Last Name'] || '') + ' / ' + (formData['Instrument'] || ''), '');
      studentId = UtilityScriptLibrary.generateNextId(contactsSheet, 'Student ID', 'Q',
        (formData['Student First Name'] || '') + ' ' + (formData['Student Last Name'] || ''));

      var newRow = new Array(headers.length);
      for (var i = 0; i < headers.length; i++) newRow[i] = '';

      newRow[getCol('Student ID') - 1]           = studentId;
      newRow[getCol('Student Last Name') - 1]     = formData['Student Last Name'] || '';
      newRow[getCol('Student First Name') - 1]    = formData['Student First Name'] || '';
      newRow[getCol('Instrument') - 1]            = formData['Instrument'] || '';
      newRow[getCol('Teacher ID') - 1]            = teacherId;
      newRow[getCol('Age') - 1]                   = standardizedAge;
      newRow[getCol('First Enrollment Term') - 1] = enrollmentTerm || '';
      newRow[getCol('Student Lookup') - 1]        = studentKey;

      if (getCol('Graduation Year')) newRow[getCol('Graduation Year') - 1] = UtilityScriptLibrary.calculateGraduationYear(formData['Grade']);
      if (getCol('School District')) newRow[getCol('School District') - 1] = formData['School'] || '';
      if (getCol('School Teacher')) newRow[getCol('School Teacher') - 1] = formData['SchoolTeacher'] || '';
      if (getCol('Experience')) newRow[getCol('Experience') - 1] = formData['Experience'] || '';
      if (getCol('Experience Start Range')) newRow[getCol('Experience Start Range') - 1] = calculateExperienceStartRange(formData['Experience']);

      contactsSheet.appendRow(newRow);

      var registeredCol = getCol('Currently Registered');
      if (registeredCol) {
        var newRowIndex = contactsSheet.getLastRow();
        var cell = contactsSheet.getRange(newRowIndex, registeredCol);
        cell.insertCheckboxes();
        cell.setValue(true);
      }
    }

    UtilityScriptLibrary.debugLog('processStudent', 'SUCCESS', 'Completed', 'ID: ' + studentId, '');
    return { studentId: studentId, parentId: parentId, studentRow: studentRow, teacherId: teacherId };

  } catch (error) {
    UtilityScriptLibrary.debugLog('processStudent', 'ERROR', 'Failed', '', error.message);
    throw error;
  }
}

function processStudentSelection(selectedIndices) {
  try {
    if (!selectedIndices || selectedIndices.length === 0) {
      UtilityScriptLibrary.debugLog('processStudentSelection', 'INFO', 'User cancelled — no students selected', '', '');
      return;
    }

    var scriptProps = PropertiesService.getScriptProperties();
    var studentList = JSON.parse(scriptProps.getProperty('reassign_studentList'));

    var selectedStudents = [];
    for (var i = 0; i < selectedIndices.length; i++) {
      var index = selectedIndices[i];
      if (index >= 0 && index < studentList.length) selectedStudents.push(studentList[index]);
    }

    if (selectedStudents.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No students were selected.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    UtilityScriptLibrary.debugLog('processStudentSelection', 'INFO', 'Students selected', 'Count: ' + selectedStudents.length, '');
    scriptProps.setProperty('reassign_selectedStudents', JSON.stringify(selectedStudents));
    selectNewTeacher();

  } catch (error) {
    UtilityScriptLibrary.debugLog('processStudentSelection', 'ERROR', 'Failed', '', error.message);
    SpreadsheetApp.getUi().alert('Error', 'Student selection failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function reassignStudentToNewTeacher() {
  try {
    var ui = SpreadsheetApp.getUi();

    var currentSemester = UtilityScriptLibrary.getCurrentSemesterName();
    if (!currentSemester) {
      ui.alert('Error', 'Could not determine current semester from Calendar sheet.', ui.ButtonSet.OK);
      return;
    }

    var year = UtilityScriptLibrary.getYearFromSemesterName(currentSemester);
    var season = UtilityScriptLibrary.extractSeasonFromSemester(currentSemester);

    UtilityScriptLibrary.debugLog('reassignStudentToNewTeacher', 'INFO', 'Starting',
      'Semester: ' + currentSemester + ', Year: ' + year + ', Season: ' + season, '');

    var scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('reassign_currentSemester', currentSemester);
    scriptProps.setProperty('reassign_year', year);
    scriptProps.setProperty('reassign_season', season);

    var activeTeachers = getActiveTeachersForDropdown();
    if (activeTeachers.length === 0) {
      ui.alert('Error', 'No active teachers found. Please check Teacher Roster Lookup.', ui.ButtonSet.OK);
      return;
    }

    scriptProps.setProperty('reassign_activeTeachers', JSON.stringify(activeTeachers));

    showTeacherDropdownDialog(
      'Reassign Students - Step 1 of 4',
      'Select the OLD teacher (current teacher):',
      activeTeachers,
      'selectStudents'
    );

  } catch (error) {
    UtilityScriptLibrary.debugLog('reassignStudentToNewTeacher', 'ERROR', 'Failed', '', error.message);
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

function selectNewTeacher() {
  try {
    var scriptProps = PropertiesService.getScriptProperties();
    var oldTeacherDisplay = scriptProps.getProperty('reassign_oldTeacherDisplay');
    var activeTeachers = JSON.parse(scriptProps.getProperty('reassign_activeTeachers'));

    var availableTeachers = activeTeachers.filter(function(t) { return t !== oldTeacherDisplay; });

    if (availableTeachers.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No other active teachers available for reassignment.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    showTeacherDropdownDialog(
      'Reassign Students - Step 3 of 4',
      'Select the NEW teacher:',
      availableTeachers,
      'enterEffectiveDate'
    );

  } catch (error) {
    UtilityScriptLibrary.debugLog('selectNewTeacher', 'ERROR', 'Failed', '', error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 3 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function selectStudents(oldTeacherDisplay) {
  try {
    var ui = SpreadsheetApp.getUi();
    var scriptProps = PropertiesService.getScriptProperties();

    if (!oldTeacherDisplay) {
      UtilityScriptLibrary.debugLog('selectStudents', 'INFO', 'User cancelled at old teacher prompt', '', '');
      return;
    }

    UtilityScriptLibrary.debugLog('selectStudents', 'INFO', 'Old teacher selected', oldTeacherDisplay, '');
    scriptProps.setProperty('reassign_oldTeacherDisplay', oldTeacherDisplay);

    var oldTeacherInfo = UtilityScriptLibrary.getTeacherInfoByDisplayName(oldTeacherDisplay);
    if (!oldTeacherInfo || !oldTeacherInfo.rosterUrl) {
      ui.alert('Error', 'Could not find roster URL for old teacher: ' + oldTeacherDisplay, ui.ButtonSet.OK);
      return;
    }

    scriptProps.setProperty('reassign_oldTeacherInfo', JSON.stringify(oldTeacherInfo));

    var currentSemester = scriptProps.getProperty('reassign_currentSemester');
    var oldRosterWorkbook = SpreadsheetApp.openByUrl(oldTeacherInfo.rosterUrl);
    var oldRosterSheet = findSemesterRoster(oldRosterWorkbook, currentSemester);

    if (!oldRosterSheet) {
      ui.alert('Error', 'Could not find roster sheet for ' + currentSemester + ' in old teacher\'s workbook.', ui.ButtonSet.OK);
      return;
    }

    var studentList = getActiveStudentsFromRoster(oldRosterSheet);

    if (studentList.length === 0) {
      ui.alert('Error', 'No active students found in old teacher\'s roster.', ui.ButtonSet.OK);
      return;
    }

    UtilityScriptLibrary.debugLog('selectStudents', 'INFO', 'Found active students',
      'Teacher: ' + oldTeacherDisplay + ', Count: ' + studentList.length, '');

    scriptProps.setProperty('reassign_studentList', JSON.stringify(studentList));
    scriptProps.setProperty('reassign_oldRosterSheetName', oldRosterSheet.getName());
    scriptProps.setProperty('reassign_oldRosterWorkbookId', oldRosterWorkbook.getId());

    showStudentCheckboxDialog(
      'Reassign Students - Step 2 of 4',
      'Select students to transfer:',
      studentList,
      'processStudentSelection'
    );

  } catch (error) {
    UtilityScriptLibrary.debugLog('selectStudents', 'ERROR', 'Failed', oldTeacherDisplay, error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 2 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sendTeacherAssignmentEmailsUI() {
  var ui = SpreadsheetApp.getUi();

  // Get current semester from Calendar D2
  var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');
  if (!calendarSheet) {
    ui.alert('Calendar sheet not found.');
    return;
  }
  var currentSemester = String(calendarSheet.getRange(2, 4).getValue()).trim();
  if (!currentSemester) {
    ui.alert('No current semester defined in Calendar D2.');
    return;
  }

  // Get active teachers for dropdown
  var teachers = getActiveTeachersForDropdown();
  if (!teachers || teachers.length === 0) {
    ui.alert('No active teachers found.');
    return;
  }

  // Prompt for teacher selection
  var teacherList = teachers.map(function(name, i) { return (i + 1) + '. ' + name; }).join('\n');
  var response = ui.prompt(
    'Send Teacher Assignment Emails',
    'Enter the number of the teacher to send for:\n\n' + teacherList,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var selection = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(selection) || selection < 1 || selection > teachers.length) {
    ui.alert('Invalid selection.');
    return;
  }

  var selectedDisplayName = teachers[selection - 1];
  var teacherInfo = UtilityScriptLibrary.getTeacherInfoByDisplayName(selectedDisplayName);
  if (!teacherInfo || !teacherInfo.teacherId) {
    ui.alert('Could not resolve teacher ID for: ' + selectedDisplayName);
    return;
  }

  // Look up teacher contact info (name + email) from Teachers and Admin
  var teacherContact = getTeacherContactForAssignmentEmail(teacherInfo.teacherId);
  if (!teacherContact) {
    ui.alert('Could not find contact info for teacher: ' + selectedDisplayName);
    return;
  }
  if (!teacherContact.email) {
    ui.alert('No email address on file for ' + selectedDisplayName + '. Please add one before sending.');
    return;
  }

  // Get form sheet for current semester
  var formSS = UtilityScriptLibrary.getWorkbook('formResponses');
  var formSheet = formSS.getSheetByName(currentSemester);
  UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'DEBUG',
    'Pre-fieldmap checkpoint', 
    'formSheet: ' + (formSheet ? formSheet.getName() : 'NOT FOUND') + ', currentSemester: ' + currentSemester, '');
  if (!formSheet) {
    ui.alert('Form sheet not found for semester: ' + currentSemester);
    return;
  }

  var norm = UtilityScriptLibrary.normalizeHeader;
  var formHeaderMap = UtilityScriptLibrary.getHeaderMap(formSheet);

  var teacherCol       = formHeaderMap[norm('Teacher')];
  var studentIdCol     = formHeaderMap[norm('Student ID')];
  var parentIdCol      = formHeaderMap[norm('Parent ID')];
  var sentCol          = formHeaderMap[norm('Teacher Assignment Sent')];
  var fieldMapSheet = UtilityScriptLibrary.getSheet('fieldMap');
  if (!fieldMapSheet) {
    ui.alert('FieldMap sheet not found.');
    return;
  }
  var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);

  // Resolve instrument and student first name columns via internal field names
  var instrumentCol   = null;
  var studentFirstCol = null;
  for (var normHeader in formHeaderMap) {
    var internalName = fieldMap[normHeader];
    if (internalName === 'Instrument')        instrumentCol   = formHeaderMap[normHeader];
    if (internalName === 'Student First Name') studentFirstCol = formHeaderMap[normHeader];
  }
  UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'DEBUG',
    'Column resolution', 
    'instrumentCol: ' + instrumentCol + ', studentFirstCol: ' + studentFirstCol, '');
    UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'DEBUG',
    'Column check', 
    'teacherCol: ' + teacherCol + ', studentIdCol: ' + studentIdCol + ', parentIdCol: ' + parentIdCol + ', sentCol: ' + sentCol, '');
    if (!teacherCol || !studentIdCol || !parentIdCol) {
    ui.alert('Required columns missing from form sheet. Please check Teacher, Student ID, and Parent ID columns.');
    return;
  }
  if (!sentCol) {
    ui.alert('Teacher Assignment Sent column not found. Please add it to the form sheet before running.');
    return;
  }

  // Build parent lookup
  var parentsSheet = UtilityScriptLibrary.getSheet('parents');
  var parentHeaderMap = UtilityScriptLibrary.getHeaderMap(parentsSheet);
  var parentsData = parentsSheet.getDataRange().getValues();
  var pIdCol        = parentHeaderMap[norm('Parent ID')];
  var pFirstCol     = parentHeaderMap[norm('Parent First Name')];
  var pLastCol      = parentHeaderMap[norm('Parent Last Name')];
  var pSalutCol     = parentHeaderMap[norm('Salutation')];
  var pEmailCol     = parentHeaderMap[norm('Email')];

  var parentMap = {};
  for (var p = 1; p < parentsData.length; p++) {
    var pid = String(parentsData[p][pIdCol - 1] || '').trim();
    if (pid) {
      parentMap[pid] = {
        firstName:  String(parentsData[p][pFirstCol - 1]  || '').trim(),
        lastName:   String(parentsData[p][pLastCol - 1]   || '').trim(),
        salutation: String(parentsData[p][pSalutCol - 1]  || '').trim(),
        email:      String(parentsData[p][pEmailCol - 1]  || '').trim()
      };
    }
  }

  // Resolve teacher ID from display name on the form
  // Form stores display name in Teacher column, not ID
  var formData = formSheet.getDataRange().getValues();
  var pendingByParent = {}; // key: parentId, value: { parentData, students: [{firstName, instrument, rowIndex}] }

  for (var i = 1; i < formData.length; i++) {
    var row = formData[i];
    var rowTeacherRaw  = String(row[teacherCol - 1] || '').trim();
    var rowStudentId   = String(row[studentIdCol - 1] || '').trim();
    var rowParentId    = String(row[parentIdCol - 1] || '').trim();
    var rowSent        = row[sentCol - 1];

    if (!rowTeacherRaw || !rowStudentId || !rowParentId) continue;
    if (rowSent) continue; // already sent

    // Resolve teacher ID from display name
    var rowTeacherInfo = UtilityScriptLibrary.getTeacherInfoByDisplayName(rowTeacherRaw);
    if (!rowTeacherInfo || rowTeacherInfo.teacherId !== teacherInfo.teacherId) continue;

    var parentData = parentMap[rowParentId];
    if (!parentData) {
      UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'WARNING',
        'Parent not found — skipping row', 'Row: ' + (i + 1) + ', Parent ID: ' + rowParentId, '');
      continue;
    }
    if (!parentData.email) {
      UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'WARNING',
        'Parent has no email — skipping row', 'Row: ' + (i + 1) + ', Parent ID: ' + rowParentId, '');
      continue;
    }

    var studentFirst = studentFirstCol ? String(row[studentFirstCol - 1] || '').trim() : rowStudentId;
    var instrument   = instrumentCol   ? String(row[instrumentCol - 1]   || '').trim() : '';

    if (!pendingByParent[rowParentId]) {
      pendingByParent[rowParentId] = {
        parentData: parentData,
        students: [],
        rowIndices: []
      };
    }
    pendingByParent[rowParentId].students.push({ firstName: studentFirst, instrument: instrument });
    pendingByParent[rowParentId].rowIndices.push(i + 1); // 1-based sheet row
  }

  var parentIds = Object.keys(pendingByParent);
  if (parentIds.length === 0) {
    ui.alert('No unsent assignments found for ' + selectedDisplayName + '.');
    return;
  }

  var confirm = ui.alert(
    'Send Teacher Assignment Emails',
    'Ready to send to ' + parentIds.length + ' parent(s) for teacher ' + selectedDisplayName + '. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var sent = 0;
  var failed = 0;
  var timestamp = UtilityScriptLibrary.formatDateFlexible(new Date(), 'MM/dd/yy');

  for (var k = 0; k < parentIds.length; k++) {
    var parentId = parentIds[k];
    var entry    = pendingByParent[parentId];
    var parent   = entry.parentData;

    try {
      var greeting = (parent.salutation && parent.lastName)
        ? 'Dear ' + parent.salutation + ' ' + parent.lastName
        : 'Dear ' + parent.firstName;

      // Build student name + instrument string
      var allSameInstrument = true;
      var firstInstrument = entry.students[0].instrument;
      for (var s = 1; s < entry.students.length; s++) {
        if (entry.students[s].instrument !== firstInstrument) {
          allSameInstrument = false;
          break;
        }
      }

      var studentString;
      if (allSameInstrument) {
        var names = entry.students.map(function(stu) { return stu.firstName; });
        var nameList = names.length === 1
          ? names[0]
          : names.slice(0, -1).join(', ') + ' and ' + names[names.length - 1];
        studentString = nameList + ' will be taking ' + firstInstrument + ' lessons';
      } else {
        var parts = entry.students.map(function(stu) {
          return stu.firstName + (stu.instrument ? ' (' + stu.instrument + ')' : '');
        });
        var partList = parts.length === 1
          ? parts[0]
          : parts.slice(0, -1).join(', ') + ' and ' + parts[parts.length - 1];
        studentString = partList + ' will be taking lessons';
      }

      var teacherAddressName = (teacherContact.salutation && teacherContact.lastName)
        ? teacherContact.salutation + ' ' + teacherContact.lastName
        : teacherContact.firstName;

      var subject = 'Your QAMP Teacher Assignment \u2013 ' + currentSemester;
      var body =
        greeting + ',\n\n' +
        'We\'re excited to let you know that ' + studentString + ' with ' +
        teacherContact.firstName + ' ' + teacherContact.lastName + ' this ' + currentSemester + '.\n\n' +
        'You can reach ' + teacherAddressName + ' at ' + teacherContact.email + '. ' +
        'They will be in touch with you soon to arrange lesson times.\n\n' +
        'If you have any questions in the meantime, please don\'t hesitate to reach out at admin@quakermusic.org.\n\n' +
        'We look forward to a wonderful semester!\n\n' +
        'Warm regards,\n' +
        'Cristin Kalinowski\n' +
        'Quaker Arts Music Program';

      UtilityScriptLibrary.sendEmail(parent.email, subject, body);

      // Write timestamp to Teacher Assignment Sent column for each row
      for (var r = 0; r < entry.rowIndices.length; r++) {
        formSheet.getRange(entry.rowIndices[r], sentCol).setValue(timestamp);
      }

      sent++;
      UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'INFO',
        'Email sent', 'Parent ID: ' + parentId + ', Email: ' + parent.email, '');

    } catch (emailError) {
      failed++;
      UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'ERROR',
        'Email failed', 'Parent ID: ' + parentId, emailError.message);
    }
  }

  var resultMsg = 'Done. Sent: ' + sent;
  if (failed > 0) resultMsg += ', Failed: ' + failed + '. Check debug log for details.';
  ui.alert(resultMsg);

  UtilityScriptLibrary.debugLog('sendTeacherAssignmentEmailsUI', 'SUCCESS',
    'Batch complete', 'Sent: ' + sent + ', Failed: ' + failed, '');
}

function setupCompleteRosterWorkbook(spreadsheet, teacher, year, semesterName, registrationTimestamp) {
  try {
    UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'INFO', 'Starting',
      'Teacher: ' + teacher + ', Year: ' + year + ', Semester: ' + semesterName, '');

    var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
    if (!season) {
      throw new Error('Could not extract season from semester name: ' + semesterName);
    }

    var rosterSheetName = season + ' Roster';
    var rosterSheet = spreadsheet.insertSheet(rosterSheetName);
    setupNewRosterTemplate(rosterSheet);
    UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'INFO', 'Created roster sheet', rosterSheetName, '');

    var currentSemesterMonth = UtilityScriptLibrary.getCurrentSemesterMonth(semesterName);
    if (!currentSemesterMonth) {
      throw new Error('Could not determine current semester month for: ' + semesterName);
    }

    UtilityScriptLibrary.createMonthlyAttendanceSheet(spreadsheet, currentSemesterMonth, []);
    UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'INFO', 'Created attendance sheet',
      'Month: ' + currentSemesterMonth, '');

    createInvoiceLogSheet(spreadsheet);
    UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'INFO', 'Created Invoice Log sheet', '', '');

    try {
      var sheets = spreadsheet.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() === 'Sheet1') {
          spreadsheet.deleteSheet(sheets[i]);
          break;
        }
      }
    } catch (deleteError) {
      UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'WARNING', 'Could not delete Sheet1', '', deleteError.message);
    }

    try {
      spreadsheet.setActiveSheet(rosterSheet);
      spreadsheet.moveActiveSheet(1);
    } catch (orderError) {
      UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'WARNING', 'Could not set sheet order', '', orderError.message);
    }

    UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'SUCCESS', 'Completed',
      'Roster: ' + rosterSheetName + ', Attendance: ' + currentSemesterMonth, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('setupCompleteRosterWorkbook', 'ERROR', 'Failed',
      'Teacher: ' + teacher + ', Semester: ' + semesterName, error.message);
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
  
  UtilityScriptLibrary.debugLog('setupInvoiceLogHeaders', 'SUCCESS', 'Headers set up', sheet.getName(), '');
}

function setupNewRosterTemplate(sheet) {
  try {
    sheet.clear();

    var headers = [
      'Contacted',
      'First Lesson Date',
      'First Lesson Time',
      'Comments',
      'Last Name',
      'First Name',
      'Instrument',
      'Length',
      'Experience',
      'Grade',
      'School',
      'School Teacher',
      'Parent Last Name',
      'Parent First Name',
      'Phone',
      'Email',
      'Additional contacts',
      'Hours Remaining',
      'Lessons Remaining',
      'Status',
      'Student ID',
      'Admin Comments',
      'System Comments'
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    setupRosterTemplateFormatting(sheet);

    UtilityScriptLibrary.debugLog('setupNewRosterTemplate', 'SUCCESS', 'Roster template created', sheet.getName(), '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('setupNewRosterTemplate', 'ERROR', 'Failed', sheet.getName(), error.message);
  }
}

function setupRosterTemplateFormatting(sheet) {
  try {
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground(UtilityScriptLibrary.STYLES.HEADER.background)
               .setFontColor(UtilityScriptLibrary.STYLES.HEADER.text)
               .setFontWeight('bold')
               .setHorizontalAlignment('center')
               .setVerticalAlignment('middle')
               .setWrap(true);

    var columnWidths = [75, 95, 95, 220, 120, 120, 80, 55, 120, 55, 110, 110, 120, 120, 100, 220, 200, 80, 80, 80, 60, 220, 220];
    for (var i = 0; i < columnWidths.length; i++) sheet.setColumnWidth(i + 1, columnWidths[i]);

    var maxRows = sheet.getMaxRows();
    sheet.getRange(1, 15, maxRows, 1).setWrap(false);
    sheet.getRange(1, 18, maxRows, 2).setWrap(true);
    sheet.getRange(2, 18, maxRows - 1, 1).setNumberFormat('0.00');
    sheet.getRange(2, 19, maxRows - 1, 1).setNumberFormat('0');

    addRosterTemplateBorders(sheet);
    UtilityScriptLibrary.setupRosterTemplateProtection(sheet);
    sheet.setFrozenRows(1);

    UtilityScriptLibrary.debugLog('setupRosterTemplateFormatting', 'SUCCESS', 'Formatting applied', sheet.getName(), '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('setupRosterTemplateFormatting', 'ERROR', 'Failed', sheet.getName(), error.message);
  }
}

function shouldProcessEdit(e, headerMap) {
  try {
    var teacherCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher')];
    var idCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var parentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Parent ID')];

    UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Checking edit',
      'Teacher col: ' + teacherCol + ', Student ID col: ' + idCol + ', Edited col: ' + e.range.getColumn(), '');

    if (e.range.getColumn() === idCol) {
      UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Edit was to Student ID column - ignoring', '', '');
      return false;
    }

    if (e.range.getColumn() === parentIdCol) {
      UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Edit was to Parent ID column - ignoring', '', '');
      return false;
    }

    if (e.range.getColumn() !== teacherCol) {
      UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Edit was not to Teacher column - ignoring', '', '');
      return false;
    }

    var editedRow = e.range.getRow();
    var sheet = e.range.getSheet();
    var studentIdValue = sheet.getRange(editedRow, idCol).getValue();
    var teacherValue = sheet.getRange(editedRow, teacherCol).getValue();

    UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Cell values',
      'Student ID: "' + String(studentIdValue) + '", Teacher: "' + String(teacherValue) + '"', '');

    if (!teacherValue || String(teacherValue).trim() === '') {
      UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Teacher field is empty - skipping', '', '');
      return false;
    }

    if (studentIdValue && String(studentIdValue).trim() !== '') {
      UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Student ID already exists - skipping', '', '');
      return false;
    }

    UtilityScriptLibrary.debugLog('shouldProcessEdit', 'DEBUG', 'Edit validation passed - processing row ' + editedRow, '', '');
    return true;

  } catch (error) {
    UtilityScriptLibrary.debugLog('shouldProcessEdit', 'ERROR', 'Failed', '', error.message);
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

    var teacherInfo = UtilityScriptLibrary.getTeacherInfoByFullName(firstName, lastName);
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
  try {
    var getCol = UtilityScriptLibrary.createColumnFinder(contactsSheet);
    var parentIdCol = getCol('Parent ID');

    if (parentIdCol === 0) {
      throw new Error('Parent ID column not found in students sheet');
    }

    if (studentRow !== -1) {
      var currentParentId = contactsSheet.getRange(studentRow, parentIdCol).getValue();
      if (!currentParentId || String(currentParentId).trim() === '') {
        contactsSheet.getRange(studentRow, parentIdCol).setValue(parentId);
        UtilityScriptLibrary.debugLog('updateStudentWithParentId', 'DEBUG', 'Linked parent to existing student',
          'Student row: ' + studentRow + ', Parent: ' + parentId, '');
      }
    } else {
      var lastRow = contactsSheet.getLastRow();
      contactsSheet.getRange(lastRow, parentIdCol).setValue(parentId);
      UtilityScriptLibrary.debugLog('updateStudentWithParentId', 'DEBUG', 'Linked parent to new student',
        'Row: ' + lastRow + ', Parent: ' + parentId, '');
    }

  } catch (error) {
    UtilityScriptLibrary.debugLog('updateStudentWithParentId', 'ERROR', 'Failed',
      'Student row: ' + studentRow + ', Parent: ' + parentId, error.message);
    throw error;
  }
}

function updateTeacherRosterLookup(teacherId, fileUrl) {
  try {
    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');

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
    if (statusCol) lookupSheet.getRange(existingRow, statusCol).setValue('active');
    if (lastUpdatedCol) lookupSheet.getRange(existingRow, lastUpdatedCol).setValue(new Date());

    UtilityScriptLibrary.debugLog('updateTeacherRosterLookup', 'INFO', 'Updated teacher roster lookup', teacherId, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('updateTeacherRosterLookup', 'ERROR', 'Failed', teacherId, error.message);
  }
}

function verifyByDriveId(driveId) {
  try {
    if (!driveId) {
      UtilityScriptLibrary.debugLog('verifyByDriveId', 'WARNING', 'No Drive ID provided', '', '');
      return;
    }
    
    driveId = driveId.trim();
    UtilityScriptLibrary.debugLog('verifyByDriveId', 'INFO', 'Processing Drive ID', driveId, '');
    
    var studentMap = loadStudentMapFromContacts();
    var detailIssues = [];
    var summaryData = [];
    
    // Try as folder first
    try {
      var folder = DriveApp.getFolderById(driveId);
      UtilityScriptLibrary.debugLog('verifyByDriveId', 'INFO', 'Processing folder', folder.getName(), '');
      checkWorkbooksInFolder(folder, studentMap, detailIssues, summaryData, false);
    } catch (e) {
      // Not a folder, try as file
      try {
        var file = DriveApp.getFileById(driveId);
        if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) {
          throw new Error('File is not a Google Spreadsheet');
        }
        UtilityScriptLibrary.debugLog('verifyByDriveId', 'INFO', 'Processing file', file.getName(), '');
        var workbook = SpreadsheetApp.openById(driveId);
        checkWorkbook(workbook, file.getName(), studentMap, detailIssues, summaryData);
      } catch (e2) {
        throw new Error('Invalid ID or no access: ' + e2.message);
      }
    }
    
    appendToReports(detailIssues, summaryData);
    UtilityScriptLibrary.debugLog('verifyByDriveId', 'SUCCESS', 'Complete', 'Issues found: ' + detailIssues.length, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('verifyByDriveId', 'ERROR', 'Failed', '', error.message);
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
