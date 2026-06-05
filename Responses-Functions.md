================================================================================
RESPONSES FUNCTION DIRECTORY
================================================================================
    Total Functions: 75
    Most Recent version: 118

    This directory provides a quick reference for all functions in Responses script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), UtilityScriptLibrary.function2()

  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    addCarryoverStudentsToNewRoster
    addRosterTemplateBorders
    addStudentToAttendanceSheet
    addStudentToAttendanceSheetsFromDate
    addStudentToRosterFromData
    addStudentToSemesterRoster
    appendToReports
    applyTeacherDropdownToCurrentSemester
    applyTeacherDropdownToSheet
    authorizeScript
    backfillParentIds
    calculateExperienceStartRange
    checkIfStudentExists
    checkSheet
    checkWorkbook
    checkWorkbooksInFolder
    clearReports
    convertCarryoverToActive
    convertStudentInfoToAttendanceObject
    createInvoiceLogSheet
    createNewYearWorkbookForTeacher
    createNewYearWorkbooksWithContinuingStudents
    createStudentObjectForAttendance
    enterEffectiveDate
    extractFormData
    extractStudentDataFromRoster
    findMostRecentRosterSheet
    findPreviousSemesterRoster
    findSemesterRoster
    formatInvoiceLogSheet
    generateAttendanceSheetFromRoster
    generateEnrollmentComparisonGraph
    getActiveStudentsFromRoster
    getActiveTeachersForDropdown
    getAllTeachersWithGroupAssignments
    getContinuingStudentsFromWorkbook
    getExistingGroupIds
    getMostRecentMonthSheet
    getMostRecentMonthSheets
    getOrCreateRosterFromTemplate
    getTeacherInfoByDisplayName
    getTeacherInfoByFullName
    getYearRosterFolders
    handleFormEdit
    handleNewStudentFormSubmit
    hasMonthBeenInvoiced
    loadStudentMapFromContacts
    onOpen
    populateRosterWithContinuingStudents
    processParent
    processPendingAssignments
    processReassignment
    processRoster
    processSingleRow
    processStudent
    processStudentSelection
    reassignStudentToNewTeacher
    refreshCurrentSemesterTeacherDropdown
    runLogHeaders
    selectNewTeacher
    selectStudents
    setupCompleteRosterWorkbook
    setupInvoiceLogHeaders
    setupNewRosterTemplate
    setupRosterTemplateFormatting
    shouldProcessEdit
    showStudentCheckboxDialog
    showTeacherDropdownDialog
    studentExistsInAttendanceSheet
    updateAllTeacherGroupAssignments
    updateGroupAssignmentsForCurrentMonth
    updateStudentWithParentId
    updateTeacherRosterLookup
    verifyByDriveId
    verifyByDriveIdWithPrompt

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addCarryoverStudentsToNewRoster(spreadsheet, newRosterSheet, currentSemesterName) -> void
        Adds carryover students from the previous semester roster to a new roster sheet.
        Reads the previous semester roster, filters for active/carryover students with
        remaining lessons, and appends them with status set to 'carryover'.
        Category: ROSTER_MANAGEMENT
        Local functions used: findPreviousSemesterRoster(), addStudentToRosterFromData()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.extractSeasonFromSemester(),
                                UtilityScriptLibrary.debugLog()

    addRosterTemplateBorders(sheet) -> void
        Applies a right border to column D of the roster template sheet for visual
        separation.
        Category: FORMATTING
        Local functions used: None
        Utility functions used: None

    addStudentToAttendanceSheet(attendanceSheet, studentData) -> void
        Adds a single student to an existing attendance sheet. Checks for duplicates
        before appending.
        Category: ATTENDANCE
        Local functions used: studentExistsInAttendanceSheet(), createStudentObjectForAttendance()
        Utility functions used: UtilityScriptLibrary.createMonthlyAttendanceSheet(),
                                UtilityScriptLibrary.debugLog()

    addStudentToAttendanceSheetsFromDate(workbook, studentInfo, effectiveDate) -> void
        Adds a student to all attendance sheets in a workbook starting from the given
        effective date month forward.
        Category: ATTENDANCE
        Local functions used: addStudentToAttendanceSheet(), convertStudentInfoToAttendanceObject()
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.debugLog()

    addStudentToRosterFromData(rosterSheet, studentInfo, headerMap) -> void
        Appends a new student row to a roster sheet using the provided studentInfo
        object and headerMap for column mapping.
        Category: ROSTER_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    addStudentToSemesterRoster(workbook, formData, studentId, semesterName) -> void
        Adds or updates a student on the semester roster for a teacher workbook.
        Handles new students, carryover conversion, and duplicate detection.
        Category: ROSTER_MANAGEMENT
        Local functions used: findSemesterRoster(), checkIfStudentExists(),
                              convertCarryoverToActive(), addStudentToRosterFromData()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    appendToReports(detailIssues, summaryData) -> void
        Appends verification detail issues and summary data to the Detail Issues and
        Summary report sheets in the active spreadsheet.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    applyTeacherDropdownToCurrentSemester() -> void
        UI-triggered function that applies the teacher dropdown to the current semester
        form responses sheet.
        Category: UI
        Local functions used: applyTeacherDropdownToSheet(), getActiveTeachersForDropdown()
        Utility functions used: UtilityScriptLibrary.getCurrentSemesterName(), UtilityScriptLibrary.debugLog()

    applyTeacherDropdownToSheet(sheet) -> void
        Applies a data validation dropdown of active teachers to the Teacher column
        of a given sheet.
        Category: ROSTER_MANAGEMENT
        Local functions used: getActiveTeachersForDropdown()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    authorizeScript() -> void
        One-time authorization helper. Run manually to trigger the Apps Script
        authorization prompt.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    backfillParentIds() -> void
        One-time utility to backfill Parent ID values into the form responses sheet
        and the billing sheet for existing records. Reads Parent ID from the Contacts
        sheet and writes it back to matching rows.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getWorkbook(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    calculateExperienceStartRange(experience) -> String
        Converts an experience level string to a lesson start range string used in
        roster and contact records.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    checkIfStudentExists(rosterSheet, studentId, headerMap) -> Boolean
        Returns true if a student with the given ID already exists in the roster sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    checkSheet(workbook, workbookName, sheet, studentMap, detailIssues, summaryData) -> void
        Verifies a single attendance or roster sheet against the student map, collecting
        detail issues and summary data for the verification report.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    checkWorkbook(workbook, workbookName, studentMap, detailIssues, summaryData) -> void
        Iterates all sheets in a workbook and calls checkSheet() on each.
        Category: VERIFICATION
        Local functions used: checkSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    checkWorkbooksInFolder(folder, studentMap, detailIssues, summaryData, isHomeFolder) -> void
        Recursively checks all teacher workbooks in a Drive folder against the student
        map for verification.
        Category: VERIFICATION
        Local functions used: checkWorkbook()
        Utility functions used: UtilityScriptLibrary.debugLog()

    clearReports() -> void
        Clears the Detail Issues and Summary report sheets in the active spreadsheet.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    convertCarryoverToActive(rosterSheet, studentId, formData, headerMap) -> void
        Updates a carryover student's roster row to active status and updates their
        lesson quantity and other fields from the new form submission.
        Category: ROSTER_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    convertStudentInfoToAttendanceObject(studentInfo) -> Object
        Converts a studentInfo object to the format expected by attendance sheet
        functions.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    createInvoiceLogSheet(spreadsheet) -> Sheet
        Creates or clears an Invoice Log sheet in the given spreadsheet with standard
        headers.
        Category: SHEET_OPERATIONS
        Local functions used: setupInvoiceLogHeaders(), formatInvoiceLogSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createNewYearWorkbookForTeacher(teacherInfo, previousYear, newYear, previousFolder, newFolder, semesterName) -> Object
        Creates a new year roster workbook for a single teacher by copying the previous
        year's workbook and populating it with continuing students.
        Category: ROSTER_MANAGEMENT
        Local functions used: getContinuingStudentsFromWorkbook(), populateRosterWithContinuingStudents(),
                              setupCompleteRosterWorkbook()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createNewYearWorkbooksWithContinuingStudents() -> void
        UI-triggered function to create new academic year workbooks for all active
        teachers, copying continuing students from the previous year.
        Category: UI
        Local functions used: getYearRosterFolders(), getActiveTeachersForDropdown(),
                              createNewYearWorkbookForTeacher()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createStudentObjectForAttendance(studentData) -> Object
        Converts a student data array (from a roster row) to the attendance object
        format.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    enterEffectiveDate(newTeacherDisplay) -> void
        Prompts the user to enter an effective date for a teacher reassignment and
        stores it in script properties for later use by processReassignment().
        Category: REASSIGNMENT
        Local functions used: selectNewTeacher()
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractFormData(sheet, row, headerMap, fieldMap) -> Object
        Extracts form data from a specific row of a sheet using the headerMap and
        fieldMap. Returns an object with internal field names as keys.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: None

    extractStudentDataFromRoster(studentRow, headerMap) -> Object
        Extracts student fields from a roster row array using the headerMap. Returns
        an object with student field names as keys.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: None

    findMostRecentRosterSheet(spreadsheet) -> Sheet | null
        Returns the most recent semester roster sheet from a teacher workbook. Local
        version scoped to Responses.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    findPreviousSemesterRoster(spreadsheet, currentSemesterName) -> Sheet | null
        Finds the roster sheet for the semester immediately preceding the given semester
        name in a teacher workbook.
        Category: DATA_LOOKUP
        Local functions used: findSemesterRoster()
        Utility functions used: UtilityScriptLibrary.extractSeasonFromSemester(),
                                UtilityScriptLibrary.debugLog()

    findSemesterRoster(workbook, semesterName) -> Sheet | null
        Finds the roster sheet matching a given semester name (by season) in a teacher
        workbook.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.extractSeasonFromSemester()

    formatInvoiceLogSheet(sheet) -> void
        Applies formatting to the Invoice Log sheet: frozen header row, column widths,
        header background and font styles.
        Category: FORMATTING
        Local functions used: None
        Utility functions used: None

    generateAttendanceSheetFromRoster(teacherWorkbook, monthName) -> void
        Generates a monthly attendance sheet for a teacher workbook from the current
        roster data.
        Category: ATTENDANCE
        Local functions used: getActiveStudentsFromRoster(), findMostRecentRosterSheet()
        Utility functions used: UtilityScriptLibrary.createMonthlyAttendanceSheet(),
                                UtilityScriptLibrary.debugLog()

    generateEnrollmentComparisonGraph() -> void
        Generates a cumulative enrollment comparison line graph between Summer 2025 and
        Summer 2026, written to a 'Graph' sheet. Column D shows % change over prior year.
        Category: REPORTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    getActiveStudentsFromRoster(rosterSheet) -> Array
        Returns an array of active and carryover student objects from a roster sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    getActiveTeachersForDropdown() -> Array
        Returns an array of display name strings for active teachers from the Teacher
        Roster Lookup sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    getAllTeachersWithGroupAssignments() -> Array
        Returns an array of teacher objects including group assignment data from the
        Teacher Roster Lookup sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    getContinuingStudentsFromWorkbook(workbook) -> Array
        Returns an array of continuing student objects from the most recent roster
        sheet in a teacher workbook.
        Category: DATA_LOOKUP
        Local functions used: findMostRecentRosterSheet(), extractStudentDataFromRoster()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    getExistingGroupIds(sheet) -> Array
        Returns an array of existing group ID values from a sheet's Group ID column.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    getMostRecentMonthSheet(workbook) -> Sheet | null
        Returns the most recent month-named sheet from a teacher workbook.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames()

    getMostRecentMonthSheets(workbook) -> Array
        Returns all month-named sheets from the most recent month found in a teacher
        workbook.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames()

    getOrCreateRosterFromTemplate(teacherInfo, rosterFolder, year, semesterName, registrationTimestamp) -> Spreadsheet
        Gets or creates a teacher roster workbook for the given semester. If no workbook
        exists, creates one from the template and sets it up.
        Category: ROSTER_MANAGEMENT
        Local functions used: setupCompleteRosterWorkbook()
        Utility functions used: UtilityScriptLibrary.getTemplate(), UtilityScriptLibrary.debugLog()

    getTeacherInfoByDisplayName(displayName) -> Object | null
        Looks up a teacher in the Teacher Roster Lookup sheet by display name. Returns
        an object with teacherId, rosterUrl, firstName, lastName, status. Local version
        scoped to Responses.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    getTeacherInfoByFullName(firstName, lastName) -> Object | null
        Looks up a teacher in the Teacher Roster Lookup sheet by first and last name
        (case-insensitive). Returns an object with teacher fields. Local version scoped
        to Responses.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    getYearRosterFolders(previousYear, newYear) -> Object
        Returns or creates Drive folder objects for the previous and new academic year
        roster folders.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getRosterFolder()

    handleFormEdit(e) -> void
        Main event handler for form response sheet edits. Acquires a script lock,
        determines the edited row, and routes to processSingleRow() if the edit is
        relevant.
        Category: FORM_PROCESSING
        Local functions used: shouldProcessEdit(), processSingleRow()
        Utility functions used: UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    handleNewStudentFormSubmit(e) -> void
        Event handler for new student form submissions from the re-registration HTML
        form. Processes the submitted data and routes to processSingleRow().
        Category: FORM_PROCESSING
        Local functions used: processSingleRow()
        Utility functions used: UtilityScriptLibrary.debugLog()

    hasMonthBeenInvoiced(sheet) -> Boolean
        Returns true if any row in a monthly attendance sheet has an Invoice Date value,
        indicating the month has already been invoiced.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    loadStudentMapFromContacts() -> Object
        Loads all students from the Contacts sheet into a map keyed by Student ID.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap()

    onOpen() -> void
        Installs the QAMP Tools menu in the spreadsheet UI on open.
        Category: UI
        Local functions used: None
        Utility functions used: None

    populateRosterWithContinuingStudents(workbook, semesterName, students) -> void
        Adds continuing students to the semester roster sheet in a teacher workbook.
        Category: ROSTER_MANAGEMENT
        Local functions used: findSemesterRoster(), addStudentToRosterFromData()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    processParent(formData, parentsSheet, studentId, existingParentId) -> String
        Creates or updates a parent record in the Parents sheet. Returns the Parent ID.
        Category: FORM_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.generateNextId(), UtilityScriptLibrary.generateKey(),
                                UtilityScriptLibrary.findParentRow(), UtilityScriptLibrary.updateParentContactFields(),
                                UtilityScriptLibrary.debugLog()

    processPendingAssignments() -> void
        UI-triggered function to process pending teacher assignments from the Calendar
        sheet. Reassigns students whose effective date has been reached.
        Category: REASSIGNMENT
        Local functions used: processReassignment()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processReassignment() -> void
        Executes a teacher reassignment using data stored in script properties. Moves
        the student to a new teacher's roster and attendance sheets.
        Category: REASSIGNMENT
        Local functions used: getTeacherInfoByDisplayName(), addStudentToSemesterRoster(),
                              addStudentToAttendanceSheetsFromDate()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    processRoster(formData, sheet, editedRow, headerMap, fieldMap, studentId, teacherId, rosterFolder, year, semesterName) -> void
        Handles all roster operations for a processed form row: gets or creates the
        teacher workbook, adds/updates the student on the roster, and adds to attendance
        sheets.
        Category: FORM_PROCESSING
        Local functions used: getOrCreateRosterFromTemplate(), addStudentToSemesterRoster(),
                              addStudentToAttendanceSheetsFromDate()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processSingleRow(sheet, row, headerMap) -> void
        Core processing function for a single form response row. Orchestrates contact
        creation/update, parent processing, roster update, and ID writebacks.
        Category: FORM_PROCESSING
        Local functions used: extractFormData(), processStudent(), processParent(),
                              processRoster()
        Utility functions used: UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getCurrentSemesterName(),
                                UtilityScriptLibrary.getRosterFolder(), UtilityScriptLibrary.debugLog()

    processStudent(formData, contactsSheet, enrollmentTerm) -> Object
        Creates or updates a student contact record in the Contacts sheet. Returns an
        object with studentId, parentId, and isNew.
        Category: FORM_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.generateKey(), UtilityScriptLibrary.generateNextId(),
                                UtilityScriptLibrary.findStudentInContacts(), UtilityScriptLibrary.interpretAgeField(),
                                UtilityScriptLibrary.calculateGraduationYear(), UtilityScriptLibrary.formatPhoneNumber(),
                                UtilityScriptLibrary.debugLog()

    processStudentSelection(selectedIndices) -> void
        Callback from the student checkbox dialog. Stores selected students in script
        properties and proceeds to enterEffectiveDate().
        Category: REASSIGNMENT
        Local functions used: enterEffectiveDate()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reassignStudentToNewTeacher() -> void
        UI entry point for the teacher reassignment workflow. Prompts the user to
        select the old teacher, then routes to selectStudents().
        Category: REASSIGNMENT
        Local functions used: selectStudents()
        Utility functions used: UtilityScriptLibrary.debugLog()

    refreshCurrentSemesterTeacherDropdown() -> void
        Refreshes the teacher dropdown on the current semester form responses sheet.
        Category: UI
        Local functions used: applyTeacherDropdownToCurrentSemester()
        Utility functions used: UtilityScriptLibrary.debugLog()

    runLogHeaders() -> void
        One-liner wrapper that calls UtilityScriptLibrary.logAllSheetHeaders() for
        debugging.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.logAllSheetHeaders()

    selectNewTeacher() -> void
        Second step of the reassignment workflow. Presents a dropdown to select the
        new teacher and stores the selection in script properties.
        Category: REASSIGNMENT
        Local functions used: showTeacherDropdownDialog(), getActiveTeachersForDropdown()
        Utility functions used: UtilityScriptLibrary.debugLog()

    selectStudents(oldTeacherDisplay) -> void
        First step of the reassignment workflow. Presents a checkbox dialog to select
        which students to reassign from the old teacher.
        Category: REASSIGNMENT
        Local functions used: showStudentCheckboxDialog(), getActiveStudentsFromRoster(),
                              getTeacherInfoByDisplayName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    setupCompleteRosterWorkbook(spreadsheet, teacher, year, semesterName, registrationTimestamp) -> void
        Sets up a complete new roster workbook: creates the semester roster sheet,
        invoice log sheet, and applies all formatting and protections.
        Category: ROSTER_MANAGEMENT
        Local functions used: setupNewRosterTemplate(), createInvoiceLogSheet(),
                              setupRosterTemplateFormatting(), addRosterTemplateBorders()
        Utility functions used: UtilityScriptLibrary.setupRosterTemplateProtection(),
                                UtilityScriptLibrary.debugLog()

    setupInvoiceLogHeaders(sheet) -> void
        Writes the standard column headers to an Invoice Log sheet.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: None

    setupNewRosterTemplate(sheet) -> void
        Clears a sheet and writes the standard roster template headers and structure.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.setupStatusValidation(),
                                UtilityScriptLibrary.enableDatePickerForColumn()

    setupRosterTemplateFormatting(sheet) -> void
        Applies standard formatting to a roster template sheet: header styles, column
        widths, frozen rows.
        Category: FORMATTING
        Local functions used: None
        Utility functions used: None

    shouldProcessEdit(e, headerMap) -> Boolean
        Returns true if a form edit event should trigger processing. Checks that the
        edited column is Teacher or Student ID and that the row has a timestamp.
        Category: FORM_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    showStudentCheckboxDialog(title, message, studentList, callbackFunctionName) -> void
        Displays a modal HTML dialog with a checkbox list of students. On submission,
        calls the specified callback function with selected indices.
        Category: UI
        Local functions used: None
        Utility functions used: None

    showTeacherDropdownDialog(title, message, teacherList, callbackFunctionName) -> void
        Displays a modal HTML dialog with a teacher dropdown. On submission, calls the
        specified callback function with the selected teacher.
        Category: UI
        Local functions used: None
        Utility functions used: None

    studentExistsInAttendanceSheet(attendanceSheet, studentId) -> Boolean
        Returns true if a student with the given ID already exists in an attendance
        sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    updateAllTeacherGroupAssignments() -> void
        UI-triggered function to update group assignments for all active teachers for
        the current semester.
        Category: UI
        Local functions used: getAllTeachersWithGroupAssignments(), updateGroupAssignmentsForCurrentMonth()
        Utility functions used: UtilityScriptLibrary.getCurrentSemesterName(), UtilityScriptLibrary.debugLog()

    updateGroupAssignmentsForCurrentMonth(firstName, lastName, semesterName) -> void
        Updates group lesson assignments for a single teacher for the current month's
        attendance sheet.
        Category: ROSTER_MANAGEMENT
        Local functions used: getTeacherInfoByDisplayName(), getMostRecentMonthSheets(),
                              getExistingGroupIds()
        Utility functions used: UtilityScriptLibrary.getTeacherGroupAssignments(),
                                UtilityScriptLibrary.debugLog()

    updateStudentWithParentId(contactsSheet, studentRow, parentId) -> void
        Writes a Parent ID value to a student's row in the Contacts sheet.
        Category: DATA_UPDATE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder(), UtilityScriptLibrary.debugLog()

    updateTeacherRosterLookup(teacherId, fileUrl) -> void
        Updates the Roster URL for a teacher in the Teacher Roster Lookup sheet.
        Category: DATA_UPDATE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.findTeacherInRosterLookup(),
                                UtilityScriptLibrary.debugLog()

    verifyByDriveId(driveId) -> Object
        Verifies student IDs in a teacher workbook identified by Drive ID. Returns a
        result object with counts and issues found.
        Category: VERIFICATION
        Local functions used: checkWorkbook(), loadStudentMapFromContacts()
        Utility functions used: UtilityScriptLibrary.debugLog()

    verifyByDriveIdWithPrompt() -> void
        UI-triggered wrapper for verifyByDriveId(). Prompts for a Drive ID and
        displays verification results.
        Category: VERIFICATION
        Local functions used: verifyByDriveId(), appendToReports()
        Utility functions used: UtilityScriptLibrary.debugLog()

  ================================================================================
  FUNCTIONS BY CATEGORY
  ================================================================================

  ATTENDANCE (3 functions):
    addStudentToAttendanceSheet
    addStudentToAttendanceSheetsFromDate
    generateAttendanceSheetFromRoster

  DATA_EXTRACTION (2 functions):
    extractFormData
    extractStudentDataFromRoster

  DATA_LOOKUP (16 functions):
    checkIfStudentExists
    findMostRecentRosterSheet
    findPreviousSemesterRoster
    findSemesterRoster
    getActiveStudentsFromRoster
    getActiveTeachersForDropdown
    getAllTeachersWithGroupAssignments
    getContinuingStudentsFromWorkbook
    getExistingGroupIds
    getMostRecentMonthSheet
    getMostRecentMonthSheets
    getTeacherInfoByDisplayName
    getTeacherInfoByFullName
    hasMonthBeenInvoiced
    loadStudentMapFromContacts
    studentExistsInAttendanceSheet

  DATA_UPDATE (2 functions):
    updateStudentWithParentId
    updateTeacherRosterLookup

  FORM_PROCESSING (7 functions):
    handleFormEdit
    handleNewStudentFormSubmit
    processRoster
    processSingleRow
    processStudent
    processParent
    shouldProcessEdit

  FORMATTING (3 functions):
    addRosterTemplateBorders
    formatInvoiceLogSheet
    setupRosterTemplateFormatting

  HELPERS (7 functions):
    authorizeScript
    backfillParentIds
    calculateExperienceStartRange
    convertStudentInfoToAttendanceObject
    createStudentObjectForAttendance
    getYearRosterFolders
    runLogHeaders

  REASSIGNMENT (7 functions):
    enterEffectiveDate
    processPendingAssignments
    processReassignment
    processStudentSelection
    reassignStudentToNewTeacher
    selectNewTeacher
    selectStudents

  REPORTING (1 function):
    generateEnrollmentComparisonGraph

  ROSTER_MANAGEMENT (9 functions):
    addCarryoverStudentsToNewRoster
    addStudentToRosterFromData
    addStudentToSemesterRoster
    applyTeacherDropdownToSheet
    convertCarryoverToActive
    createNewYearWorkbookForTeacher
    getOrCreateRosterFromTemplate
    populateRosterWithContinuingStudents
    setupCompleteRosterWorkbook

  SHEET_OPERATIONS (3 functions):
    createInvoiceLogSheet
    setupInvoiceLogHeaders
    setupNewRosterTemplate

  UI (7 functions):
    applyTeacherDropdownToCurrentSemester
    createNewYearWorkbooksWithContinuingStudents
    onOpen
    refreshCurrentSemesterTeacherDropdown
    showStudentCheckboxDialog
    showTeacherDropdownDialog
    updateAllTeacherGroupAssignments

  VERIFICATION (7 functions):
    appendToReports
    checkSheet
    checkWorkbook
    checkWorkbooksInFolder
    clearReports
    verifyByDriveId
    verifyByDriveIdWithPrompt

================================================================================
END OF FUNCTION DIRECTORY
================================================================================