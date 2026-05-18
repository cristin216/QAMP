/* 
================================================================================
WORKBOOK DOCUMENTATION
================================================================================
  Workbook Name: Responses
  Most Recent version: 117

  Primary Purpose:
      Stores all form submissions, teacher and student metadata, semester registration
      data, calendar information, and debugging logs. Used by scripts for enrollment,
      roster management, lesson scheduling, billing, reporting, and debugging.

  Access Pattern:
      Scripts use UtilityScriptLibrary.getSheet() to access sheets dynamically.
      Column headers are referenced dynamically; avoid hardcoding indexes.

  --------------------------------------------------------------------------------
  SHEET: Semester Metadata
  --------------------------------------------------------------------------------
    Purpose:
        Stores semester-level metadata such as names, start/end dates, and verification status.
    Columns:
    | Semester Name       | Date   | Example: "Spring 2024" | Name of the semester. |
    | Start Date          | Date   | 1/1/2024               | Semester start date. |
    | End Date            | Date   | 5/31/2024              | Semester end date. |
    | Rates Verification  | String | "2023-2024"            | Verification period for rates. |
    | Program Verification| String | "Yes"                  | Verification for program availability. |
    Example Row:
    | Spring 2024 | 1/1/2024 | 5/31/2024 | 2023-2024 | Yes |

  --------------------------------------------------------------------------------
  SHEET: Calendar
  --------------------------------------------------------------------------------
    Purpose:
        Tracks weekly calendar information for multiple concurrent semesters.
        Each row represents one week; columns repeat in groups of three for each semester.
    Column Structure (repeating groups):
    | Week       | Number | 1           | Week number. |
    | Week Start | Date   | 6/1/2024    | Week start date for this semester group. |
    | Week End   | Date   | 6/7/2024    | Week end date for this semester group. |
    | Semester   | String | Summer 2024 | Semester name for this column group. |
    Notes:
        Column D (row 2) is read directly by handleFormEdit and processPendingAssignments
        to get the current semester name without requiring openById permissions.

  --------------------------------------------------------------------------------
  SHEET: FieldMap
  --------------------------------------------------------------------------------
    Purpose:
        Maps form headers from Form Responses to internal field names used in scripts.
    Columns:
    | Form Header (from Form Responses)    | String | "City, State, Zip" | Original form field name. |
    | Internal Field Name (used in script) | String | "CityZip"          | Field name scripts reference. |
    Notes:
        Required because Google Form question text changes between semesters.
        Scripts always use internal field names; never reference raw form question text directly.

  --------------------------------------------------------------------------------
  SHEET: Teacher Roster Lookup
  --------------------------------------------------------------------------------
    Purpose:
        Stores metadata for all teachers for roster management and display name assignments.
        Teacher ID is the unique backend key; Display Name is used only for UX dropdowns.
    Columns:
    | First Name       | String | "Angel"      | Teacher first name. |
    | Last Name        | String | "Rhodes"      | Teacher last name. |
    | Roster URL       | String | anonymized    | Link to individual teacher roster workbook. |
    | Teacher ID       | String | "T0002"       | Unique backend ID; format: T + 4 digits. |
    | Display Name     | String | "Rhodes A-2"  | Used in dropdown only; not for lookups. |
    | Group Assignment | String | "Suzuki"      | Optional group assignment. |
    | Status           | String | "active"      | "active", "potential", "returning", or "former". |
    | Last Updated     | Date   | 10/14/2025    | Row last updated. |
    Notes:
        Scripts use Teacher ID for all backend processing.
        Display Name is resolved to Teacher ID via getTeacherIdByDisplayName() in Utility.

  --------------------------------------------------------------------------------
  SHEET: Debug
  --------------------------------------------------------------------------------
    Purpose:
        Centralized log for script activity, function calls, events, messages, and errors.
    Columns:
    | Timestamp    | DateTime    | 10/14/2025 14:32       | Log entry timestamp. |
    | Function     | String      | "processStudent"       | Function generating the log. |
    | Event Type   | String      | "INFO"                 | Severity: INFO, WARNING, ERROR, SUCCESS, DEBUG. |
    | Message      | String      | "Created new student"  | Human-readable description. |
    | Data         | String/JSON | {studentId:"Q0006"}    | Optional structured data. |
    | Error Details| String      | "TypeError: undefined" | Stack trace or error message. |
    Notes:
        Append-only. Cleared automatically when entry count exceeds 500 (keeps last 100).

  --------------------------------------------------------------------------------
  SHEET: Semester Registration Sheets (e.g., Spring 2026, Summer 2026)
  --------------------------------------------------------------------------------
    Purpose:
        Stores student registration information for a given semester.
        Each semester has its own sheet named "Season YYYY".
    Notes:
        Column headers are raw Google Form question text, which varies between semesters
        as forms are updated. The FieldMap sheet maps these question strings to internal
        field names used by scripts. Do not rely on column position or exact header text.
        Key logical fields (regardless of exact header wording):
    | Timestamp            | DateTime | 12/31/2023 10:00        | Form submission timestamp. |
    | Email Address        | String   | anonymized@example.com  | Student/guardian contact email. |
    | Student First Name   | String   | Mary                    | Student first name. |
    | Student Last Name    | String   | Cool                    | Student last name. |
    | Instrument           | String   | Violin                  | Instrument student is registering for. |
    | Teacher              | String   | Rhodes A-2              | Assigned teacher display name. |
    | Experience Level     | String   | 2-4 years               | Experience level. |
    | Lesson Length        | String   | 30 min                  | Lesson duration chosen. |
    | Qty30                | Number   | 15                      | Number of 30-min lessons. |
    | Qty45                | Number   | 0                       | Number of 45-min lessons. |
    | Qty60                | Number   | 0                       | Number of 60-min lessons. |
    | Age (Is Adult)       | Boolean  | No                      | Whether student is adult. |
    | Grade                | String   | 4                       | Grade or upcoming grade. |
    | School               | String   | Orchard Park            | School district. |
    | SchoolTeacher        | String   | Mrs. Awesome            | School music teacher. |
    | Salutation           | String   | Mrs.                    | Mr./Mrs./Ms./Dr., etc. |
    | Parent First Name    | String   | Jane                    | Billing/guardian first name. |
    | Parent Last Name     | String   | Cool                    | Billing/guardian last name. |
    | Address Street       | String   | 36 Main St.             | Mailing address. |
    | CityZip              | String   | Orchard Park, NY 14127  | Full address or city/zip (varies by semester). |
    | Phone                | String   | (716) 123-4567          | Contact phone. |
    | Billing Preference   | String   | Mail                    | Mail or Email. |
    | Additional Contacts  | String   | N/A                     | Optional additional contact info. |
    | Parent Group Interest| Boolean  | No                      | Guardian interest in volunteering. |
    | Referral             | String   | Online Search           | How student heard about program. |
    | Student ID           | String   | Q0006                   | Unique student ID: Q + 4 digits. Written by script. |
    Notes:
        - City and Zip were combined ("City, State, Zip") in older semesters;
          split into separate City and Zip columns in Summer 2026+.
        - Some semesters include Enrollment or Suzuki interest columns; others do not.
        - Scripts read each semester sheet dynamically via FieldMap.

  --------------------------------------------------------------------------------
  SHEET: Random Crap
  --------------------------------------------------------------------------------
    Purpose:
        Legacy test sheet retained from early development. Not used by any script.
        Can be safely ignored.

================================================================================
RESPONSES FUNCTION DIRECTORY
================================================================================
    Total Functions: 73

    This directory provides a quick reference for all functions in Responses script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: function1(), function2()

  --------------------------------------------------------------------------------
  ALPHABETICAL INDEX:
  --------------------------------------------------------------------------------
    addCarryoverStudentsToNewRoster
    addRosterTemplateBorders
    addStudentToAttendanceSheet
    addStudentToAttendanceSheetsFromDate
    addStudentToNewRosterTemplate
    addStudentToRosterFromData
    addStudentToSemesterRoster
    appendToReports
    applyTeacherDropdownToCurrentSemester
    applyTeacherDropdownToSheet
    authorizeScript
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

  --------------------------------------------------------------------------------
  FUNCTION REFERENCE (Alphabetical)
  --------------------------------------------------------------------------------
    addCarryoverStudentsToNewRoster(spreadsheet, newRosterSheet, currentSemesterName) -> Number
        Adds students from previous semester roster who have Status="active" AND Lessons Remaining > 0 to new roster.
        Sets Status to "Carryover" and applies WARNING formatting to entire row (A-X).
        Category: ROSTER_OPERATIONS
        Local functions used: findPreviousSemesterRoster(), checkIfStudentExists()
        Utility functions used: debugLog(), getHeaderMap(), STYLES.WARNING.background, STYLES.WARNING.text

    addRosterTemplateBorders(sheet) -> void
        Adds thick green borders and dotted borders to roster template sheet to separate sections.
        Borders separate editable area, admin area, and hours/lessons remaining columns.
        Category: SHEET_FORMATTING
        Local functions used: None
        Utility functions used: STYLES

    addStudentToAttendanceSheet(attendanceSheet, studentData) -> void
        Adds a single student to an attendance sheet with proper formatting and validation.
        Includes status dropdown, formatted lesson length, and proper column values.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: createStudentObjectForAttendance()
        Utility functions used: debugLog(), createStudentSections(), setupStatusValidation()

    addStudentToAttendanceSheetsFromDate(workbook, studentInfo, effectiveDate) -> void
        Adds student to attendance sheets starting from a specific effective date.
        Used for mid-semester enrollments or teacher reassignments.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getMostRecentMonthSheets(), convertStudentInfoToAttendanceObject(),
                             studentExistsInAttendanceSheet()
        Utility functions used: debugLog(), createStudentSections(), setupStatusValidation()

    addStudentToNewRosterTemplate(sheet, formData, studentId) -> void
        Adds a new student to the roster template sheet with all form data.
        Batch-reads existing rows to find empty slot before appending.
        Handles grade, experience, instrument, lesson length, and parent information.
        Category: ROSTER_OPERATIONS
        Local functions used: calculateExperienceStartRange()
        Utility functions used: debugLog(), STYLES

    addStudentToRosterFromData(rosterSheet, studentInfo, headerMap) -> void
        Adds student to roster using extracted student data with proper formatting.
        Batch-reads existing rows to find empty slot before appending.
        Used when processing pending assignments or reassignments.
        Category: ROSTER_OPERATIONS
        Local functions used: None
        Utility functions used: debugLog(), STYLES

    addStudentToSemesterRoster(workbook, formData, studentId, semesterName) -> void
        Adds student to the semester-specific roster sheet within a teacher workbook.
        Finds or creates season roster; handles carryover conversion and duplicate checks.
        Category: ROSTER_OPERATIONS
        Local functions used: addCarryoverStudentsToNewRoster(), addStudentToNewRosterTemplate(),
                             checkIfStudentExists(), convertCarryoverToActive(), setupNewRosterTemplate()
        Utility functions used: debugLog(), extractSeasonFromSemester()

    appendToReports(detailIssues, summaryData) -> void
        Appends verification results to Student ID Detail Report and Student ID Summary Report sheets.
        Creates sheets with headers if they don't exist.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    applyTeacherDropdownToCurrentSemester() -> void
        Applies teacher dropdown validation to the current semester sheet.
        Called manually via Refresh Teacher Dropdown menu item.
        Category: UI_OPERATIONS
        Local functions used: applyTeacherDropdownToSheet()
        Utility functions used: debugLog(), getCurrentSemesterName()

    applyTeacherDropdownToSheet(sheet) -> void
        Applies teacher dropdown data validation to the Teacher column of a sheet.
        Uses active teacher list from Teacher Roster Lookup.
        Category: UI_OPERATIONS
        Local functions used: getActiveTeachersForDropdown()
        Utility functions used: debugLog(), getHeaderMap(), normalizeHeader()

    authorizeScript() -> void
        One-time authorization function to grant script necessary permissions.
        Accesses UI which triggers the authorization prompt.
        Category: SETUP
        Local functions used: None
        Utility functions used: None

    calculateExperienceStartRange(experience) -> String
        Converts experience level string (e.g., "2-4 years") to an estimated start year range.
        Returns empty string for unrecognized input; no logging on clean execution.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: debugLog()

    checkIfStudentExists(rosterSheet, studentId, headerMap) -> Boolean|'CARRYOVER'
        Checks if a student already exists in a roster sheet by Student ID.
        Returns 'CARRYOVER' if found with Carryover status, true if found active, false if not found.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: debugLog()

    checkSheet(workbook, workbookName, sheet, studentMap, detailIssues, summaryData) -> void
        Checks a single sheet for student ID mismatches against the Contacts student map.
        Looks for Student ID and name columns; records issues to detailIssues/summaryData arrays.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: normalizeHeader()

    checkWorkbook(workbook, workbookName, studentMap, detailIssues, summaryData) -> void
        Iterates all sheets in a workbook and delegates to checkSheet() for each.
        Category: VERIFICATION
        Local functions used: checkSheet()
        Utility functions used: None

    checkWorkbooksInFolder(folder, studentMap, detailIssues, summaryData, isHomeFolder) -> void
        Iterates all Google Sheets files in a Drive folder and verifies each via checkWorkbook().
        Skips non-spreadsheet files and excluded workbook names (Contacts, Teacher Interest Survey).
        Category: VERIFICATION
        Local functions used: checkWorkbook()
        Utility functions used: None

    clearReports() -> void
        Deletes Student ID Detail Report and Student ID Summary Report sheets if they exist.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    convertCarryoverToActive(rosterSheet, studentId, formData, headerMap) -> void
        Converts a Carryover student to Active status when they re-register.
        Updates lesson/grade/contact fields, resets formatting to normal alternating rows.
        Category: ROSTER_OPERATIONS
        Local functions used: None
        Utility functions used: debugLog(), STYLES

    convertStudentInfoToAttendanceObject(studentInfo) -> Object
        Converts named-object roster student data to attendance sheet format.
        Used by addStudentToAttendanceSheetsFromDate (named object input path).
        See also: createStudentObjectForAttendance (positional array input path).
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    createInvoiceLogSheet(spreadsheet) -> Sheet
        Creates the Invoice Log sheet in a teacher roster workbook.
        Sets up headers and applies formatting and protection.
        Category: SHEET_CREATION
        Local functions used: setupInvoiceLogHeaders(), formatInvoiceLogSheet()
        Utility functions used: debugLog()

    createNewYearWorkbookForTeacher(teacherInfo, previousYear, newYear, previousFolder, newFolder, semesterName) -> Object|null
        Creates a new-year roster workbook for one teacher by copying continuing students
        from the previous year's workbook. Returns {url, studentCount} or null if no prior workbook found.
        Category: YEAR_ROLLOVER
        Local functions used: getContinuingStudentsFromWorkbook(), getOrCreateRosterFromTemplate(),
                             populateRosterWithContinuingStudents()
        Utility functions used: debugLog()

    createNewYearWorkbooksWithContinuingStudents() -> void
        UI entry point for new-year workbook creation. Confirms teacher status verification,
        then iterates all active teachers and calls createNewYearWorkbookForTeacher() for each.
        Updates Teacher Roster Lookup URLs on success.
        Category: YEAR_ROLLOVER
        Local functions used: createNewYearWorkbookForTeacher(), getYearRosterFolders(),
                             updateTeacherRosterLookup()
        Utility functions used: debugLog(), getCurrentSemesterName(), getYearFromSemesterName()

    createStudentObjectForAttendance(studentData) -> Object
        Creates a student data object formatted for attendance sheet insertion.
        Accepts a positional array as input.
        Used by addStudentToAttendanceSheet (positional array input path).
        See also: convertStudentInfoToAttendanceObject (named object input path).
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    enterEffectiveDate(newTeacherDisplay) -> void
        Step 4 of 4 in reassignment flow (called as HTML dialog callback). Stores new teacher,
        prompts for effective date, validates format, then calls processReassignment().
        Category: UI_OPERATIONS
        Local functions used: processReassignment()
        Utility functions used: debugLog()

    extractFormData(sheet, row, headerMap, fieldMap) -> Object
        Extracts form data from a sheet row using header and field mappings.
        Normalizes headers and handles data type conversions.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: normalizeHeader()

    extractStudentDataFromRoster(studentRow, headerMap) -> Object
        Extracts complete student data from a roster row.
        Returns object with all student fields.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: debugLog()

    findMostRecentRosterSheet(spreadsheet) -> Sheet|null
        Finds the most recent roster sheet (name contains "Roster") in a workbook.
        Returns last matching sheet or null if none found.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: None

    findPreviousSemesterRoster(spreadsheet, currentSemesterName) -> String|null
        Searches Semester Metadata and workbook sheets to find the most recent previous semester roster sheet.
        Returns sheet name if found, null otherwise.
        Category: ROSTER_OPERATIONS
        Local functions used: None
        Utility functions used: debugLog(), getSheet(), extractSeasonFromSemester(), normalizeHeader()

    findSemesterRoster(workbook, semesterName) -> Sheet|null
        Finds a semester-specific roster sheet in a workbook by matching season name.
        Returns the sheet if found, null otherwise.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: debugLog()

    formatInvoiceLogSheet(sheet) -> void
        Applies formatting to Invoice Log sheet including date/currency formats.
        Protects sheet from editing (view-only for teachers).
        Category: SHEET_FORMATTING
        Local functions used: None
        Utility functions used: debugLog()

    generateAttendanceSheetFromRoster(teacherWorkbook, monthName) -> Sheet
        Generates a new attendance sheet for a given month from roster data.
        Populates with all active students from roster.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getActiveStudentsFromRoster(), addStudentToAttendanceSheet()
        Utility functions used: debugLog()

    getActiveStudentsFromRoster(rosterSheet) -> Array
        Retrieves all active and carryover students from a roster sheet.
        Returns array of student data objects with rowNumber attached.
        Category: DATA_RETRIEVAL
        Local functions used: extractStudentDataFromRoster()
        Utility functions used: debugLog()

    getActiveTeachersForDropdown() -> Array
        Gets sorted list of display names for teacher dropdown validation.
        Regenerates teacher display names, writes changes back to lookup sheet,
        and syncs renamed display names into the current semester Teacher column.
        Returns display names for teachers with status: potential, active, or returning.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: debugLog(), getSheet(), getCurrentSemesterName(), getHeaderMap(), normalizeHeader()

    getAllTeachersWithGroupAssignments() -> Array
        Retrieves all teachers with their group assignments from lookup sheet.
        Returns array of teacher objects with name and group assignment.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: debugLog(), getSheet(), createColumnFinder()

    getContinuingStudentsFromWorkbook(workbook) -> Array
        Extracts all continuing students (Status=Active or Carryover, Lessons Remaining > 0)
        from the most recent roster sheet in a teacher workbook.
        Returns array of student data objects.
        Category: DATA_RETRIEVAL
        Local functions used: findMostRecentRosterSheet()
        Utility functions used: debugLog(), getHeaderMap()

    getExistingGroupIds(sheet) -> Array
        Extracts all existing group IDs (G####) from an attendance sheet.
        Returns array of group ID strings.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: None

    getMostRecentMonthSheet(workbook) -> Sheet|null
        Finds and returns the most recent month attendance sheet in a workbook.
        Compares sheet names against month name list to find latest.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: debugLog()

    getMostRecentMonthSheets(workbook) -> Array
        Gets all month sheets from the most recent month onwards.
        Returns array of sheet objects sorted chronologically.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: debugLog()

    getOrCreateRosterFromTemplate(teacherInfo, rosterFolder, year, semesterName, registrationTimestamp?) -> Spreadsheet
        Gets existing teacher roster or creates new one from template.
        Creates complete workbook structure including Invoice Log.
        Updates Teacher Roster Lookup URL on creation.
        Category: ROSTER_OPERATIONS
        Local functions used: setupCompleteRosterWorkbook(), updateTeacherRosterLookup()
        Utility functions used: debugLog()

    getTeacherInfoByDisplayName(displayName) -> Object|null
        Gets complete teacher information from Teacher Roster Lookup using display name.
        Returns full info object {firstName, lastName, teacherId, rosterUrl, status, lastUpdated} or null.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: debugLog(), getSheet(), createColumnFinder()

    getTeacherInfoByFullName(firstName, lastName) -> Object|null
        Gets complete teacher information using split first and last name.
        Returns full info object {firstName, lastName, teacherId, rosterUrl, status, lastUpdated} or null.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: debugLog(), getSheet(), createColumnFinder()

    getYearRosterFolders(previousYear, newYear) -> Object
        Locates and returns {previous, next} Drive folder references for the given years.
        Throws if either folder is not found under the main Rosters folder.
        Category: YEAR_ROLLOVER
        Local functions used: None
        Utility functions used: debugLog(), getRosterFolder()

    handleFormEdit(e) -> void
        Installable trigger handler for Teacher column edits in the current semester sheet.
        Reads current semester from Calendar D2 directly (no openById needed).
        Validates edit via shouldProcessEdit(), acquires script lock, calls processSingleRow().
        Displays success or error message via Browser.msgBox().
        Category: EVENT_HANDLER
        Local functions used: shouldProcessEdit(), processSingleRow()
        Utility functions used: debugLog(), getHeaderMap(), normalizeHeader()

    hasMonthBeenInvoiced(sheet) -> Boolean
        Checks if a month attendance sheet has already been invoiced.
        Looks for "Invoiced" status in metadata or specific cells.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: debugLog()

    loadStudentMapFromContacts() -> Object
        Loads all students from Contacts into a firstName|lastName -> studentId lookup map.
        Only includes records with Q-prefixed IDs.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: getSheet(), createColumnFinder()

    onOpen() -> void
        Simple trigger. Runs when spreadsheet opens. Creates custom QAMP Tools menu.
        Menu items: Refresh Teacher Dropdown, Update Roster Groups, Process Teacher Assignments,
        Reassign Student to Different Teacher, Clear Reports, Verify by Drive ID,
        Create New Year Workbooks with Continuing Students.
        Note: Does not apply teacher dropdown on open; use Refresh Teacher Dropdown menu item instead.
        Category: EVENT_HANDLER
        Local functions used: None
        Utility functions used: None

    populateRosterWithContinuingStudents(workbook, semesterName, students) -> void
        Writes continuing student data into the season roster sheet of a workbook.
        Sorts alphabetically by last/first name; sets Status to "Carryover".
        Category: YEAR_ROLLOVER
        Local functions used: None
        Utility functions used: debugLog(), extractSeasonFromSemester()

    processParent(formData, parentsSheet, studentId, existingParentId?) -> String
        Processes parent information from form submission.
        Creates new parent record or updates existing parent contact fields. Returns Parent ID.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: debugLog(), findParentRow(), formatAddress(), formatPhoneNumber(),
                                generateKey(), generateNextId(), normalizeHeader(),
                                parseCityZipMessy(), updateParentContactFields()

    processPendingAssignments() -> void
        Menu entry point. Processes all pending teacher assignments from current semester sheet.
        Reads current semester from Calendar D2. Batch processes rows where Teacher is
        assigned but Student ID is missing. Confirms with user before processing.
        Category: BATCH_OPERATIONS
        Local functions used: processSingleRow()
        Utility functions used: debugLog(), getHeaderMap(), normalizeHeader()

    processReassignment() -> void
        Final step of reassignment flow. Reads all stored ScriptProperties, retrieves
        roster folder from Year Metadata, moves students from old to new teacher's roster
        and attendance sheets. Updates Contacts with new Teacher ID.
        Category: ROSTER_OPERATIONS
        Local functions used: getTeacherInfoByDisplayName(), getOrCreateRosterFromTemplate(),
                             findSemesterRoster(), checkIfStudentExists(), addStudentToRosterFromData(),
                             addStudentToAttendanceSheetsFromDate(), updateTeacherRosterLookup()
        Utility functions used: debugLog(), getSheet(), normalizeHeader()

    processRoster(formData, sheet, editedRow, headerMap, fieldMap, studentId, teacherId, rosterFolder, year, semesterName) -> void
        Processes roster update for assigned teacher.
        Coordinates student data preparation and roster/attendance updates.
        Category: ROSTER_OPERATIONS
        Local functions used: getOrCreateRosterFromTemplate(), addStudentToSemesterRoster(),
                             addStudentToAttendanceSheetsFromDate()
        Utility functions used: debugLog()

    processSingleRow(sheet, row, headerMap) -> void
        Processes a single form submission row end-to-end.
        Orchestrates student creation/update, parent linking, and roster assignment.
        Retrieves fieldMap internally; roster folder resolved from Year Metadata sheet.
        Category: DATA_PROCESSING
        Local functions used: extractFormData(), processStudent(), processParent(),
                             updateStudentWithParentId(), processRoster()
        Utility functions used: debugLog(), getSheet(), normalizeHeader()

    processStudent(formData, contactsSheet, enrollmentTerm) -> Object
        Processes student information from form submission.
        Creates new student record or updates existing. Returns {studentId, parentId, studentRow, teacherId}.
        Category: DATA_PROCESSING
        Local functions used: calculateExperienceStartRange()
        Utility functions used: debugLog(), findStudentRow(), generateNextId(),
                                generateKey(), getTeacherIdByDisplayName(), calculateGraduationYear(),
                                normalizeHeader()

    processStudentSelection(selectedIndices) -> void
        Step 2 of 4 in reassignment flow (called as HTML dialog callback). Resolves selected
        student indices against stored student list and chains to selectNewTeacher().
        Category: UI_OPERATIONS
        Local functions used: selectNewTeacher()
        Utility functions used: debugLog()

    reassignStudentToNewTeacher() -> void
        Entry point (Step 1 of 4) for reassignment flow. Reads current semester, stores
        semester/year/season in ScriptProperties, shows teacher dropdown dialog with
        callback to selectStudents().
        Category: UI_OPERATIONS
        Local functions used: getActiveTeachersForDropdown(), showTeacherDropdownDialog()
        Utility functions used: debugLog(), getCurrentSemesterName(), getYearFromSemesterName(),
                                extractSeasonFromSemester()

    refreshCurrentSemesterTeacherDropdown() -> void
        Menu entry point. Refreshes teacher dropdown for the current semester sheet.
        Category: UI_OPERATIONS
        Local functions used: applyTeacherDropdownToSheet()
        Utility functions used: debugLog(), getCurrentSemesterName()

    runLogHeaders() -> void
        Diagnostic utility. Calls UtilityScriptLibrary.logAllSheetHeaders() to log all
        sheet headers to the execution log.
        Category: SETUP
        Local functions used: None
        Utility functions used: logAllSheetHeaders()

    selectNewTeacher() -> void
        Step 3 of 4 in reassignment flow. Retrieves active teacher list from ScriptProperties,
        filters out old teacher, and shows teacher dropdown dialog with callback to enterEffectiveDate().
        Category: UI_OPERATIONS
        Local functions used: showTeacherDropdownDialog()
        Utility functions used: debugLog()

    selectStudents(oldTeacherDisplay) -> void
        Step 2 of 4 in reassignment flow (called as HTML dialog callback). Stores old teacher,
        opens their roster, retrieves active students, and shows checkbox dialog
        with callback to processStudentSelection().
        Category: UI_OPERATIONS
        Local functions used: getTeacherInfoByDisplayName(), findSemesterRoster(),
                             getActiveStudentsFromRoster(), showStudentCheckboxDialog()
        Utility functions used: debugLog()

    setupCompleteRosterWorkbook(spreadsheet, teacher, year, semesterName, registrationTimestamp?) -> void
        Sets up complete roster workbook structure: creates season roster sheet, current month
        attendance sheet, Invoice Log sheet, and removes default Sheet1.
        Category: SHEET_CREATION
        Local functions used: setupNewRosterTemplate(), createInvoiceLogSheet()
        Utility functions used: debugLog(), extractSeasonFromSemester(), getCurrentSemesterMonth(),
                                createMonthlyAttendanceSheet()

    setupInvoiceLogHeaders(sheet) -> void
        Sets up column headers and formatting for the Invoice Log sheet.
        Columns: Invoice Number, Invoice Date, Invoice Period, Invoice URL, Total Amount.
        Category: SHEET_SETUP
        Local functions used: None
        Utility functions used: debugLog(), styleHeaderRow()

    setupNewRosterTemplate(sheet) -> void
        Sets up a new roster template sheet by clearing, writing headers, and applying formatting.
        Category: SHEET_SETUP
        Local functions used: setupRosterTemplateFormatting()
        Utility functions used: debugLog()

    setupRosterTemplateFormatting(sheet) -> void
        Applies column widths, number formats, wrap settings, borders, protection, and frozen rows
        to a roster template sheet.
        Category: SHEET_FORMATTING
        Local functions used: addRosterTemplateBorders()
        Utility functions used: debugLog(), setupRosterTemplateProtection()

    shouldProcessEdit(e, headerMap) -> Boolean
        Determines if a Teacher column edit should trigger student processing.
        Returns false if: edit is not on Teacher column, Teacher field is empty,
        or Student ID already exists on the row (already processed).
        Category: VALIDATION
        Local functions used: None
        Utility functions used: debugLog(), normalizeHeader()

    showStudentCheckboxDialog(title, message, studentList, callbackFunctionName) -> void
        Displays custom HTML dialog with student checkboxes for multi-student selection.
        Used in reassignment workflow. Callback receives array of selected indices.
        Category: UI_OPERATIONS
        Local functions used: None
        Utility functions used: None

    showTeacherDropdownDialog(title, message, teacherList, callbackFunctionName) -> void
        Displays custom HTML dialog with teacher dropdown for selection.
        Used in reassignment and other workflows requiring teacher selection.
        Category: UI_OPERATIONS
        Local functions used: None
        Utility functions used: None

    studentExistsInAttendanceSheet(attendanceSheet, studentId) -> Boolean
        Checks if a student already exists in an attendance sheet by Student ID.
        Returns true if found, false otherwise.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: debugLog(), normalizeHeader()

    updateAllTeacherGroupAssignments() -> void
        Menu entry point. Updates group assignments for all teachers with group assignments.
        Batch processes group assignments to current month attendance sheets.
        Category: BATCH_OPERATIONS
        Local functions used: getAllTeachersWithGroupAssignments(), updateGroupAssignmentsForCurrentMonth()
        Utility functions used: debugLog()

    updateGroupAssignmentsForCurrentMonth(firstName, lastName, semesterName) -> void
        Updates group assignments in teacher's current month attendance sheet.
        Adds new group sections that don't already exist.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getTeacherInfoByFullName(), getMostRecentMonthSheet(),
                             getExistingGroupIds()
        Utility functions used: debugLog(), createGroupSections(), setupStatusValidation(),
                                getTeacherGroupAssignments()

    updateStudentWithParentId(contactsSheet, studentRow, parentId) -> void
        Updates a student record with their parent's ID.
        Links student to parent in Contacts sheet. Handles both new (last row) and
        existing (known row) student cases.
        Category: DATA_UPDATE
        Local functions used: None
        Utility functions used: debugLog(), createColumnFinder()

    updateTeacherRosterLookup(teacherId, fileUrl) -> void
        Updates Teacher Roster Lookup with roster file URL and status.
        Marks teacher as active and updates last modified timestamp.
        Category: DATA_UPDATE
        Local functions used: None
        Utility functions used: debugLog(), getSheet(), createColumnFinder(), findTeacherInRosterLookup()

    verifyByDriveId(driveId) -> void
        Accepts a Drive ID (folder or spreadsheet), loads student map from Contacts,
        and runs student ID verification across all workbooks found. Appends results to report sheets.
        Category: VERIFICATION
        Local functions used: loadStudentMapFromContacts(), checkWorkbooksInFolder(),
                             checkWorkbook(), appendToReports()
        Utility functions used: None

    verifyByDriveIdWithPrompt() -> void
        UI wrapper for verifyByDriveId(). Prompts user for Drive ID and calls verifyByDriveId().
        Category: VERIFICATION
        Local functions used: verifyByDriveId()
        Utility functions used: None

  --------------------------------------------------------------------------------
  CATEGORIES:
  --------------------------------------------------------------------------------

    ATTENDANCE_OPERATIONS (4 functions):
      addStudentToAttendanceSheet
      addStudentToAttendanceSheetsFromDate
      generateAttendanceSheetFromRoster
      updateGroupAssignmentsForCurrentMonth

    BATCH_OPERATIONS (2 functions):
      processPendingAssignments
      updateAllTeacherGroupAssignments

    DATA_PROCESSING (9 functions):
      calculateExperienceStartRange
      convertStudentInfoToAttendanceObject
      createStudentObjectForAttendance
      extractFormData
      extractStudentDataFromRoster
      processParent
      processSingleRow
      processStudent

    DATA_RETRIEVAL (10 functions):
      getActiveStudentsFromRoster
      getActiveTeachersForDropdown
      getAllTeachersWithGroupAssignments
      getContinuingStudentsFromWorkbook
      getExistingGroupIds
      getMostRecentMonthSheet
      getMostRecentMonthSheets
      getTeacherInfoByDisplayName
      getTeacherInfoByFullName
      loadStudentMapFromContacts

    DATA_UPDATE (2 functions):
      updateStudentWithParentId
      updateTeacherRosterLookup

    EVENT_HANDLER (2 functions):
      handleFormEdit
      onOpen

    ROSTER_OPERATIONS (9 functions):
      addCarryoverStudentsToNewRoster
      addStudentToNewRosterTemplate
      addStudentToRosterFromData
      addStudentToSemesterRoster
      convertCarryoverToActive
      findPreviousSemesterRoster
      getOrCreateRosterFromTemplate
      processReassignment
      processRoster

    SETUP (2 functions):
      authorizeScript
      runLogHeaders

    SHEET_CREATION (2 functions):
      createInvoiceLogSheet
      setupCompleteRosterWorkbook

    SHEET_FORMATTING (3 functions):
      addRosterTemplateBorders
      formatInvoiceLogSheet
      setupRosterTemplateFormatting

    SHEET_OPERATIONS (5 functions):
      findMostRecentRosterSheet
      findSemesterRoster
      getMostRecentMonthSheet
      getMostRecentMonthSheets
      populateRosterWithContinuingStudents

    SHEET_SETUP (2 functions):
      setupInvoiceLogHeaders
      setupNewRosterTemplate

    UI_OPERATIONS (10 functions):
      applyTeacherDropdownToCurrentSemester
      applyTeacherDropdownToSheet
      enterEffectiveDate
      processStudentSelection
      reassignStudentToNewTeacher
      refreshCurrentSemesterTeacherDropdown
      selectNewTeacher
      selectStudents
      showStudentCheckboxDialog
      showTeacherDropdownDialog

    VALIDATION (4 functions):
      checkIfStudentExists
      hasMonthBeenInvoiced
      shouldProcessEdit
      studentExistsInAttendanceSheet

    VERIFICATION (7 functions):
      appendToReports
      checkSheet
      checkWorkbook
      checkWorkbooksInFolder
      clearReports
      verifyByDriveId
      verifyByDriveIdWithPrompt

    YEAR_ROLLOVER (4 functions):
      createNewYearWorkbookForTeacher
      createNewYearWorkbooksWithContinuingStudents
      getYearRosterFolders
      populateRosterWithContinuingStudents

================================================================================
END OF FUNCTION DIRECTORY
================================================================================
*/
