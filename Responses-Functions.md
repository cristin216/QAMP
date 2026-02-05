/* 
================================================================================
WORKBOOK DOCUMENTATION
================================================================================
  Workbook Name: Responses
  Most Recent version: 110
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
    | Semester Name | Date       | Example: "Spring 2024"      | Name of the semester. |
    | Start Date    | Date       | 1/1/2024                    | Semester start date. |
    | End Date      | Date       | 5/31/2024                   | Semester end date. |
    | Rates Verification | String | "2023-2024"                | Verification period for rates. |
    | Program Verification | String | "Yes"                     | Verification for program availability. |
    Example Row:
    | Spring 2024 | 1/1/2024 | 5/31/2024 | 2023-2024 | Yes |

  --------------------------------------------------------------------------------
  SHEET: Calendar
  --------------------------------------------------------------------------------
    Purpose:
        Tracks weekly calendar information for multiple semesters.
    Columns:
    | Week | Number | 1                           | Week number. |
    | Week Start | Date | 6/1/2024                    | Week start date. |
    | Week End   | Date | 6/7/2024                    | Week end date. |
    | Semester   | String | Summer 2024                | Semester name. |
    Notes:
        Each week may have overlapping entries for multiple semesters.

  --------------------------------------------------------------------------------
  SHEET: FieldMap
  --------------------------------------------------------------------------------
    Purpose:
        Maps form headers from Form Responses to internal field names used in scripts.
    Columns:
    | Form Header (from Form Responses) | String | "City, State, Zip" | Original form field name. |
    | Internal Field Name (used in script)| String | "CityZip" | Field name scripts reference. |
    Notes:
        Helps scripts dynamically map form submissions to internal processing.

  --------------------------------------------------------------------------------
  SHEET: Teacher Roster Lookup
  --------------------------------------------------------------------------------
    Purpose:
        Stores metadata for all teachers for roster management, instrument assignments, and display names.
    Columns:
    | Teacher Name | String | "John Rhodes" | Full name. |
    | Roster URL   | String | anonymized | Link to individual roster. |
    | Teacher ID   | String | "T0002" | Unique ID, format: 1 capital letter + 4 digits. Letter may denote teacher type. |
    | Display Name | String | "Rhodes" | Typically last name; used as primary identifier in scripts. |
    | Group Assignment | String | "Suzuki" | Optional group assignment. |
    | Status       | String | "active" | "active" or "former". |
    | Last Updated | Date   | 10/14/2025 | Row last updated. |
    Notes:
        Scripts rely on IDs and Display Names for processing.

  --------------------------------------------------------------------------------
  SHEET: Debug
  --------------------------------------------------------------------------------
    Purpose:
        Centralized log for script activity, function calls, events, messages, and errors.
    Columns:
    | Timestamp | DateTime | 10/14/2025 14:32 | Log entry timestamp. |
    | Function  | String   | "processTeacher" | Function generating the log. |
    | Event Type | String  | "INFO" | Severity: INFO, WARNING, ERROR. |
    | Message   | String   | "Processed teacher successfully" | Human-readable description. |
    | Data      | String/JSON | {teacherId:"T0002"} | Optional structured data. |
    | Error Details | String | "TypeError: undefined" | Stack trace or error message. |
    Notes:
        Append-only

  --------------------------------------------------------------------------------
  SHEET: Semester Registration Sheets (e.g., Spring 2024, Summer 2024)
  --------------------------------------------------------------------------------
    Purpose:
        Stores student registration information for a given semester.
        Each semester has its own sheet named "Season YYYY".
    Columns:
    | Timestamp | DateTime | 12/31/2023 10:00 | Form submission timestamp. |
    | Email Address | String | anonymized@example.com | Student/guardian contact email. |
    | Student First Name | String | Mary | Student first name. |
    | Student Last Name | String | Cool | Student last name. |
    | Instrument | String | Violin | Instrument student is registering for. |
    | Teacher | String | Johnston | Assigned teacher. |
    | Student Experience Level | String | 2-4 years | Experience level. |
    | Lesson Length | String | 30 min | Lesson duration chosen. |
    | Lesson Quantity (30-min) | Number | 15 | Number of 30-min lessons. |
    | Lesson Quantity (45-min) | Number | 0 | Number of 45-min lessons. |
    | Lesson Quantity (60-min) | Number | 0 | Number of 60-min lessons. |
    | Is Adult | Boolean | No | Whether student is adult. |
    | Student Grade | String | 4 | Grade or upcoming grade. |
    | School District | String | Orchard Park | School district. |
    | In-School Teacher | String | Mrs. Awesome Ms. | School music teacher. |
    | Term of Address | String | Mrs. | Mr./Mrs./Ms./Dr., etc. |
    | Guardian First Name | String | Jane | Billing/guardian first name. |
    | Guardian Last Name | String | Cool | Billing/guardian last name. |
    | Street Address | String | 36 Main St. | Mailing address. |
    | City, State, Zip | String | Orchard Park, NY 14127 | Full address. |
    | Phone Number | String | (716) 123-4567 | Contact phone. |
    | Billing Preference | String | Mail | Mail or Email. |
    | Additional Contact | String | N/A | Optional additional contact info. |
    | Fundraising/Volunteer Interest | Boolean | No | Guardian interest in volunteering. |
    | Referral Source | String | Online Search | How student heard about program. |
    | Student ID | String | Q0006 | Unique ID: 1 capital letter + 4 digits. |
    | Enrollment Type | String | Suzuki | Program or lesson type. |
    | Comments | String | N/A | Optional free-text comments. |
    Notes:
        - Scripts read each semester sheet dynamically.
        - Multi-line form prompts are simplified into field names.

================================================================================
RESPONSES FUNCTION DIRECTORY
================================================================================
    Total Functions: 75

    This directory provides a quick reference for all functions in Responses script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), UtilityScriptLibrary.function2()

  --------------------------------------------------------------------------------
  ALPHABETICAL INDEX:
  --------------------------------------------------------------------------------
    addCarryoverStudentsToNewRoster, addRosterTemplateBorders, addStudentToAttendanceSheet, addStudentToAttendanceSheets,
    addStudentToAttendanceSheetsFromDate, addStudentToNewRosterTemplate, addStudentToRosterFromData,
    addStudentToSemesterRoster, applyTeacherDropdownToCurrentSemester, applyTeacherDropdownToSheet,
    authorizeScript, calculateExperienceStartRange, calculateGraduationYear, checkIfStudentExists,
    convertCarryoverToActive, convertStudentInfoToAttendanceObject, createCurrentMonthAttendance, createInvoiceLogSheet,
    createStudentObjectForAttendance, createTeacherInfoObject, extractDisplayNameFromFullName,
    extractFormData, extractNumericLessonLength, extractStudentDataFromRoster,
    findPreviousSemesterRoster, findSemesterRoster, findTeacherInEnhancedRosterLookup, formatInvoiceLogSheet,
    formatLessonLengthWithMinutes, generateAttendanceSheetFromRoster, getActiveStudentsFromRoster,
    getActiveTeachersForDropdown, getAllTeachersWithGroupAssignments, getCurrentSemesterMonth,
    getCurrentSemesterName, getExistingGroupIds, getMonthName, getMostRecentMonthSheet,
    getMostRecentMonthSheets, getOrCreateRosterFromTemplate, getSemesterDates,
    getTeacherIdFromContacts, getTeacherInfoByDisplayName, getTeacherInfoByFullName,
    getTeacherRosterLookupSheet, handleFormEdit, handleRosterUpdate, hasMonthBeenInvoiced,
    markTeacherAsActiveInLookup, onOpen, processParent, processPendingAssignments,
    processRoster, processSingleRow, processStudent, reassignStep2_selectStudents,
    reassignStep2b_processStudentSelection, reassignStep3_selectNewTeacher,
    reassignStep4_enterEffectiveDate, reassignStep5_processReassignment, reassignStudentToNewTeacher,
    refreshCurrentSemesterTeacherDropdown, setupCompleteRosterWorkbook, setupInvoiceLogHeaders,
    setupNewRosterTemplate, setupRosterTemplateFormatting, setupRosterTemplateProtection,
    setupStatusValidation, shouldProcessEdit, showStudentCheckboxDialog, showTeacherDropdownDialog,
    studentExistsInAttendanceSheet, updateAllTeacherGroupAssignments, updateGroupAssignmentsForCurrentMonth,
    updateStudentWithParentId, updateTeacherRosterLookup, validateEnhancedTeacherRosterLookup

  --------------------------------------------------------------------------------
  FUNCTION REFERENCE (Alphabetical)
  --------------------------------------------------------------------------------
    addCarryoverStudentsToNewRoster(spreadsheet, newRosterSheet, currentSemesterName) -> Number
        Adds students from previous semester roster who have Status="active" AND Lessons Remaining > 0 to new roster.
        Sets Status to "Carryover" and applies WARNING formatting to entire row (A-X).
        Category: ROSTER_OPERATIONS
        Local functions used: findPreviousSemesterRoster(), checkIfStudentExists()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.getHeaderMap(), 
                                UtilityScriptLibrary.STYLES.WARNING.background, UtilityScriptLibrary.STYLES.WARNING.text

    addRosterTemplateBorders(sheet) -> void
        Adds thick green borders and dotted borders to roster template sheet to separate sections.
        Borders separate editable area, admin area, and hours/lessons remaining columns.
        Category: SHEET_FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.STYLES

    addStudentToAttendanceSheet(attendanceSheet, studentData) -> void
        Adds a single student to an attendance sheet with proper formatting and validation.
        Includes status dropdown, formatted lesson length, and proper column values.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: setupStatusValidation()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    addStudentToAttendanceSheets(rosterWorkbook, studentData, semesterName, customToday?) -> void
        Adds student to all relevant attendance sheets in the roster workbook.
        Handles both current and future month sheets based on effective date.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getMostRecentMonthSheets(), addStudentToAttendanceSheet(), getMonthName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addStudentToAttendanceSheetsFromDate(workbook, studentInfo, effectiveDate) -> void
        Adds student to attendance sheets starting from a specific effective date.
        Used for mid-semester enrollments or teacher reassignments.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getMostRecentMonthSheets(), convertStudentInfoToAttendanceObject(), 
                             studentExistsInAttendanceSheet(), addStudentToAttendanceSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addStudentToNewRosterTemplate(sheet, formData, studentId) -> void
        Adds a new student to the roster template sheet with all form data.
        Handles grade, experience, instrument, lesson length, and parent information.
        Category: ROSTER_OPERATIONS
        Local functions used: calculateGraduationYear(), calculateExperienceStartRange()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    addStudentToRosterFromData(rosterSheet, studentInfo, headerMap) -> void
        Adds student to roster using extracted student data with proper formatting.
        Used when processing pending assignments or reassignments.
        Category: ROSTER_OPERATIONS
        Local functions used: calculateGraduationYear(), calculateExperienceStartRange()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addStudentToSemesterRoster(workbook, formData, studentId, semesterName) -> void
        Adds student to the semester-specific roster sheet within a teacher workbook.
        Finds or creates semester roster and populates with student data.
        Category: ROSTER_OPERATIONS
        Local functions used: findSemesterRoster(), addStudentToNewRosterTemplate(), checkIfStudentExists()
        Utility functions used: UtilityScriptLibrary.debugLog()

    applyTeacherDropdownToCurrentSemester() -> void
        Applies teacher dropdown validation to the current semester sheet.
        Called on spreadsheet open to ensure dropdown is available.
        Category: UI_OPERATIONS
        Local functions used: getCurrentSemesterName(), applyTeacherDropdownToSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    applyTeacherDropdownToSheet(sheet) -> void
        Applies teacher dropdown data validation to the Teacher column of a sheet.
        Uses active teacher list from current semester.
        Category: UI_OPERATIONS
        Local functions used: getActiveTeachersForDropdown()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.getHeaderMap(), 
                                 UtilityScriptLibrary.normalizeHeader()

    authorizeScript() -> void
        One-time authorization function to grant script necessary permissions.
        Accesses UI which triggers the authorization prompt.
        Category: SETUP
        Local functions used: None
        Utility functions used: None

    calculateExperienceStartRange(experience) -> String
        Converts experience level to a range string (e.g., "3-5 years" from "4").
        Handles various input formats and edge cases.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    calculateGraduationYear(grade) -> Number|String
        Calculates graduation year from current grade level.
        Returns "N/A" for adult students.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    checkIfStudentExists(rosterSheet, studentId, headerMap) -> Boolean
        Checks if a student already exists in a roster sheet by Student ID.
        Returns true if found, false otherwise.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    convertCarryoverToActive(rosterSheet, studentId, formData, headerMap) -> void
        Converts a Carryover student to Active status when they re-register.
        Replaces Lessons Registered with new value, resets Lessons Completed to 0, and removes WARNING formatting.
        Category: ROSTER_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    convertStudentInfoToAttendanceObject(studentInfo) -> Object
        Converts roster student data to attendance sheet format.
        Maps fields appropriately for attendance tracking.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    createCurrentMonthAttendance() -> void
        Creates attendance sheet for the current month.
        Used for manual month sheet creation.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getCurrentSemesterName()
        Utility functions used: None

    createInvoiceLogSheet(spreadsheet) -> Sheet
        Creates the Invoice Log sheet in a teacher roster workbook.
        Sets up headers and applies formatting and protection.
        Category: SHEET_CREATION
        Local functions used: setupInvoiceLogHeaders(), formatInvoiceLogSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createStudentObjectForAttendance(studentData) -> Object
        Creates a student data object formatted for attendance sheet insertion.
        Includes all necessary fields for attendance tracking.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    createTeacherInfoObject(dataRow) -> Object
        Creates a teacher info object from a data row.
        Standardizes teacher data structure for processing.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    extractDisplayNameFromFullName(fullName) -> String
        Extracts display name (first name + last initial) from full name.
        Handles various name formats.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    extractFormData(sheet, row, headerMap, fieldMap) -> Object
        Extracts form data from a sheet row using header and field mappings.
        Normalizes headers and handles data type conversions.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    extractNumericLessonLength(lessonLengthValue) -> Number
        Extracts numeric minutes from lesson length value (e.g., "60 min" → 60).
        Handles various input formats.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    extractStudentDataFromRoster(studentRow, headerMap) -> Object
        Extracts complete student data from a roster row.
        Returns object with all student fields.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    findPreviousSemesterRoster(spreadsheet, currentSemesterName) -> String|null
        Searches Semester Metadata and workbook sheets to find the most recent previous semester roster sheet.
        Returns sheet name if found, null otherwise.
        Category: ROSTER_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.extractSeasonFromSemester(), 
                                UtilityScriptLibrary.debugLog()

    findSemesterRoster(workbook, semesterName) -> Sheet|null
        Finds a semester-specific roster sheet in a workbook.
        Returns the sheet if found, null otherwise.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    findTeacherInEnhancedRosterLookup(lookupSheet, teacherName) -> Number
        Finds a teacher's row in the Enhanced Teacher Roster Lookup sheet.
        Returns row number or -1 if not found.
        Category: LOOKUP_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    formatInvoiceLogSheet(sheet) -> void
        Applies formatting to Invoice Log sheet including date/currency formats.
        Protects sheet from editing (view-only for teachers).
        Category: SHEET_FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    formatLessonLengthWithMinutes(lessonLengthValue) -> String
        Formats lesson length value to include "min" suffix if not present.
        Standardizes lesson length display.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None

    generateAttendanceSheetFromRoster(teacherWorkbook, monthName) -> Sheet
        Generates a new attendance sheet for a given month from roster data.
        Populates with all active students from roster.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getActiveStudentsFromRoster(), addStudentToAttendanceSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getActiveStudentsFromRoster(rosterSheet) -> Array
        Retrieves all active students from a roster sheet.
        Excludes withdrawn students and returns array of student data objects.
        Category: DATA_RETRIEVAL
        Local functions used: extractStudentDataFromRoster()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    getActiveTeachersForDropdown() -> Array
        Gets list of active teachers for dropdown validation.
        Retrieves from Teacher Roster Lookup sheet.
        Category: DATA_RETRIEVAL
        Local functions used: getCurrentSemesterName(), getTeacherRosterLookupSheet()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.normalizeHeader()

    getAllTeachersWithGroupAssignments() -> Array
        Retrieves all teachers with their group assignments from lookup sheet.
        Returns array of teacher objects with name and group assignment.
        Category: DATA_RETRIEVAL
        Local functions used: getTeacherRosterLookupSheet()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    getCurrentSemesterMonth(semesterName) -> String
        Gets the current month name based on semester and current date.
        Handles semester date ranges.
        Category: DATE_OPERATIONS
        Local functions used: getSemesterDates(), getMonthName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getCurrentSemesterName() -> String
        Gets the current semester name from Calendar sheet D2.
        Returns semester identifier string.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: None

    getExistingGroupIds(sheet) -> Array
        Extracts all existing group IDs (G####) from an attendance sheet.
        Returns array of group ID strings.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getMonthName(date) -> String
        Returns month name from a date object (e.g., "January").
        Category: DATE_OPERATIONS
        Local functions used: None
        Utility functions used: None

    getMostRecentMonthSheet(workbook) -> Sheet|null
        Finds and returns the most recent month attendance sheet in a workbook.
        Compares sheet names to find latest month.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getMostRecentMonthSheets(workbook) -> Array
        Gets all month sheets from the most recent month onwards.
        Returns array of sheet objects sorted chronologically.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getOrCreateRosterFromTemplate(teacher, rosterFolder, year, semesterName) -> Spreadsheet
        Gets existing teacher roster or creates new one from template.
        Creates complete workbook structure including Invoice Log.
        Category: ROSTER_OPERATIONS
        Local functions used: setupCompleteRosterWorkbook()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getSemesterDates(semesterName) -> Object
        Returns start and end dates for a given semester.
        Retrieves from Calendar sheet configuration.
        Category: DATE_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getTeacherIdFromContacts(teacherName) -> String
        Gets teacher ID from Contacts sheet by teacher name.
        Returns empty string if not found.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.normalizeHeader()

    getTeacherInfoByDisplayName(displayName) -> Object|null
        Gets complete teacher information using display name (First Last-Initial).
        Returns teacher info object or null if not found.
        Category: DATA_RETRIEVAL
        Local functions used: getTeacherRosterLookupSheet()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    getTeacherInfoByFullName(teacherName) -> Object|null
        Gets complete teacher information using full name.
        Returns teacher info object or null if not found.
        Category: DATA_RETRIEVAL
        Local functions used: findTeacherInEnhancedRosterLookup(), getTeacherRosterLookupSheet()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    getTeacherRosterLookupSheet() -> Sheet|null
        Gets the Teacher Roster Lookup sheet from active spreadsheet.
        Returns sheet or null if not found.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: None

    handleFormEdit(e) -> void
        Main event handler for form submissions/edits in current semester sheet.
        Triggers student processing when Teacher column is edited.
        Category: EVENT_HANDLER
        Local functions used: shouldProcessEdit(), processSingleRow()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.getHeaderMap(), 
                                 UtilityScriptLibrary.normalizeHeader()

    handleRosterUpdate(teacher, rosterFolder, studentData, year, semesterName) -> void
        Handles adding/updating student in teacher roster and attendance sheets.
        Orchestrates roster creation, student addition, and attendance updates.
        Category: ROSTER_OPERATIONS
        Local functions used: getOrCreateRosterFromTemplate(), addStudentToSemesterRoster(), 
                             addStudentToAttendanceSheets()
        Utility functions used: UtilityScriptLibrary.debugLog()

    hasMonthBeenInvoiced(sheet) -> Boolean
        Checks if a month attendance sheet has already been invoiced.
        Looks for "Invoiced" status in metadata or specific cells.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    markTeacherAsActiveInLookup(teacherName) -> void
        Marks a teacher as active in the Teacher Roster Lookup sheet.
        Updates status and last updated timestamp.
        Category: DATA_UPDATE
        Local functions used: getTeacherRosterLookupSheet(), findTeacherInEnhancedRosterLookup()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    onOpen() -> void
        Runs when spreadsheet opens. Creates custom menu and applies teacher dropdown.
        Provides access to refresh, roster, and reassignment functions.
        Category: EVENT_HANDLER
        Local functions used: applyTeacherDropdownToCurrentSemester()
        Utility functions used: None

    processParent(formData, parentsSheet, studentId, existingParentId?) -> String
        Processes parent information from form submission.
        Creates new parent record or links to existing parent. Returns Parent ID.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.generateUniqueId(), 
                                 UtilityScriptLibrary.normalizeHeader()

    processPendingAssignments() -> void
        Processes all pending teacher assignments from current semester sheet.
        Batch processes multiple students that need teacher roster updates.
        Category: BATCH_OPERATIONS
        Local functions used: getCurrentSemesterName(), processSingleRow()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.getHeaderMap(), 
                                 UtilityScriptLibrary.normalizeHeader()

    processRoster(formData, sheet, editedRow, headerMap, fieldMap, studentId, rosterFolder, year, semesterName) -> void
        Processes roster update for assigned teacher.
        Coordinates student data preparation and roster/attendance updates.
        Category: ROSTER_OPERATIONS
        Local functions used: handleRosterUpdate(), extractStudentDataFromRoster()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processSingleRow(sheet, row, headerMap) -> void
        Processes a single row from form responses sheet.
        Handles student, parent, and roster processing for one student.
        Category: DATA_PROCESSING
        Local functions used: extractFormData(), processStudent(), processParent(), updateStudentWithParentId(), 
                             processRoster()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.getFieldMap()

    processStudent(formData, contactsSheet, enrollmentTerm) -> Object
        Processes student information from form data.
        Creates new student record or updates existing. Returns {studentId, isNew, studentRow}.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.generateUniqueId(), 
                                 UtilityScriptLibrary.normalizeHeader()

    reassignStep2_selectStudents(oldTeacherDisplay) -> void
        Step 2 of reassignment: Shows checkbox dialog to select students to reassign.
        Displays list of active students from old teacher's roster.
        Category: UI_OPERATIONS
        Local functions used: showStudentCheckboxDialog(), getTeacherInfoByDisplayName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reassignStep2b_processStudentSelection(selectedIndices) -> void
        Step 2b of reassignment: Processes student selection and moves to teacher selection.
        Stores selected students and calls next step.
        Category: UI_OPERATIONS
        Local functions used: reassignStep3_selectNewTeacher()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reassignStep3_selectNewTeacher() -> void
        Step 3 of reassignment: Shows dropdown to select new teacher.
        Displays active teachers for selection.
        Category: UI_OPERATIONS
        Local functions used: getActiveTeachersForDropdown(), showTeacherDropdownDialog()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reassignStep4_enterEffectiveDate(newTeacherDisplay) -> void
        Step 4 of reassignment: Prompts for effective date of reassignment.
        Validates date and proceeds to processing.
        Category: UI_OPERATIONS
        Local functions used: reassignStep5_processReassignment()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reassignStep5_processReassignment() -> void
        Step 5 of reassignment: Executes the actual student reassignment.
        Moves students from old to new teacher roster and updates attendance sheets.
        Category: ROSTER_OPERATIONS
        Local functions used: getTeacherInfoByDisplayName(), addStudentToRosterFromData(), 
                             addStudentToAttendanceSheetsFromDate()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reassignStudentToNewTeacher() -> void
        Main entry point for student reassignment workflow.
        Initiates multi-step reassignment process with teacher selection.
        Category: UI_OPERATIONS
        Local functions used: getActiveTeachersForDropdown(), showTeacherDropdownDialog()
        Utility functions used: UtilityScriptLibrary.debugLog()

    refreshCurrentSemesterTeacherDropdown() -> void
        Refreshes teacher dropdown in current semester sheet.
        Updates validation list with current active teachers.
        Category: UI_OPERATIONS
        Local functions used: getCurrentSemesterName(), applyTeacherDropdownToSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    setupCompleteRosterWorkbook(spreadsheet, teacher, year, semesterName) -> void
        Sets up complete roster workbook structure with all sheets.
        Creates roster template, Invoice Log, semester rosters, and applies formatting.
        Category: SHEET_CREATION
        Local functions used: setupNewRosterTemplate(), createInvoiceLogSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    setupInvoiceLogHeaders(sheet) -> void
        Sets up column headers for Invoice Log sheet.
        Defines Invoice Number, Invoice Date, Student Name, Program Type, and Total Amount columns.
        Category: SHEET_SETUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.styleHeaderRow()

    setupNewRosterTemplate(sheet) -> void
        Creates the roster template structure with headers and initial setup.
        Sets up all student tracking columns and applies formatting/protection.
        Category: SHEET_SETUP
        Local functions used: setupRosterTemplateFormatting(), addRosterTemplateBorders(), 
                             setupRosterTemplateProtection()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.styleHeaderRow()

    setupRosterTemplateFormatting(sheet) -> void
        Applies formatting to roster template including column widths and colors.
        Sets frozen rows/columns and conditional formatting.
        Category: SHEET_FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    setupRosterTemplateProtection(sheet) -> void
        Applies protection to non-editable columns in roster template.
        Allows editing only in designated columns.
        Category: SHEET_PROTECTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    setupStatusValidation(sheet, lastRow) -> void
        Sets up status dropdown validation for attendance sheets.
        Applies data validation with status options (Present, Absent, Excused, etc.).
        Category: SHEET_SETUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    shouldProcessEdit(e, sheet) -> Boolean
        Determines if a form edit should be processed based on various conditions.
        Checks for duplicate submissions and validates edit context.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    showStudentCheckboxDialog(title, message, studentList, callbackFunctionName) -> void
        Displays custom HTML dialog with student checkboxes for selection.
        Used in reassignment workflow for multi-student selection.
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
        Checks if a student already exists in an attendance sheet.
        Returns true if found, false otherwise.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder()

    updateAllTeacherGroupAssignments() -> void
        Updates group assignments for all teachers with roster files.
        Batch processes group assignments to current month attendance sheets.
        Category: BATCH_OPERATIONS
        Local functions used: getAllTeachersWithGroupAssignments(), updateGroupAssignmentsForCurrentMonth()
        Utility functions used: UtilityScriptLibrary.debugLog()

    updateGroupAssignmentsForCurrentMonth(teacherName, semesterName) -> void
        Updates group assignments in teacher's current month attendance sheet.
        Adds new group sections that don't already exist.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: getTeacherInfoByFullName(), getMostRecentMonthSheet(), getExistingGroupIds(), 
                             setupStatusValidation()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createGroupSections()

    updateStudentWithParentId(contactsSheet, studentRow, parentId, getCol?) -> void
        Updates a student record with their parent's ID.
        Links student to parent in Contacts sheet.
        Category: DATA_UPDATE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.normalizeHeader()

    updateTeacherRosterLookup(teacherName, fileUrl) -> void
        Updates Teacher Roster Lookup with roster file URL and status.
        Marks teacher as active and updates last modified timestamp.
        Category: DATA_UPDATE
        Local functions used: findTeacherInEnhancedRosterLookup()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.createColumnFinder()

    validateEnhancedTeacherRosterLookup() -> Boolean
        Validates and repairs Enhanced Teacher Roster Lookup sheet structure.
        Ensures all required columns exist with correct headers.
        Category: VALIDATION
        Local functions used: getTeacherRosterLookupSheet(), extractDisplayNameFromFullName()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.styleHeaderRow()

  --------------------------------------------------------------------------------
  CATEGORIES:
  --------------------------------------------------------------------------------
  
    ATTENDANCE_OPERATIONS (7 functions):
      addStudentToAttendanceSheet, addStudentToAttendanceSheets, addStudentToAttendanceSheetsFromDate,
      createCurrentMonthAttendance, generateAttendanceSheetFromRoster, updateGroupAssignmentsForCurrentMonth

    BATCH_OPERATIONS (2 functions):
      processPendingAssignments, updateAllTeacherGroupAssignments

    DATA_PROCESSING (13 functions):
      calculateExperienceStartRange, calculateGraduationYear, convertStudentInfoToAttendanceObject,
      createStudentObjectForAttendance, createTeacherInfoObject, extractDisplayNameFromFullName,
      extractFormData, extractNumericLessonLength, extractStudentDataFromRoster,
      formatLessonLengthWithMinutes, processParent, processSingleRow, processStudent

    DATA_RETRIEVAL (10 functions):
      getActiveStudentsFromRoster, getActiveTeachersForDropdown, getAllTeachersWithGroupAssignments,
      getCurrentSemesterName, getExistingGroupIds, getTeacherIdFromContacts, getTeacherInfoByDisplayName,
      getTeacherInfoByFullName

    DATA_UPDATE (3 functions):
      markTeacherAsActiveInLookup, updateStudentWithParentId, updateTeacherRosterLookup

    DATE_OPERATIONS (3 functions):
      getCurrentSemesterMonth, getMonthName, getSemesterDates

    EVENT_HANDLER (2 functions):
      handleFormEdit, onOpen

    LOOKUP_OPERATIONS (1 function):
      findTeacherInEnhancedRosterLookup

    ROSTER_OPERATIONS (8 functions):
      addStudentToNewRosterTemplate, addStudentToRosterFromData, addStudentToSemesterRoster,
      getOrCreateRosterFromTemplate, handleRosterUpdate, processRoster, reassignStep5_processReassignment

    SETUP (1 function):
      authorizeScript

    SHEET_CREATION (2 functions):
      createInvoiceLogSheet, setupCompleteRosterWorkbook

    SHEET_FORMATTING (3 functions):
      addRosterTemplateBorders, formatInvoiceLogSheet, setupRosterTemplateFormatting

    SHEET_OPERATIONS (5 functions):
      findSemesterRoster, getMostRecentMonthSheet, getMostRecentMonthSheets, getTeacherRosterLookupSheet

    SHEET_PROTECTION (1 function):
      setupRosterTemplateProtection

    SHEET_SETUP (3 functions):
      setupInvoiceLogHeaders, setupNewRosterTemplate, setupStatusValidation

    UI_OPERATIONS (9 functions):
      applyTeacherDropdownToCurrentSemester, applyTeacherDropdownToSheet, reassignStep2_selectStudents,
      reassignStep2b_processStudentSelection, reassignStep3_selectNewTeacher, reassignStep4_enterEffectiveDate,
      reassignStudentToNewTeacher, refreshCurrentSemesterTeacherDropdown, showStudentCheckboxDialog,
      showTeacherDropdownDialog

    VALIDATION (5 functions):
      checkIfStudentExists, hasMonthBeenInvoiced, shouldProcessEdit, studentExistsInAttendanceSheet,
      validateEnhancedTeacherRosterLookup
================================================================================
END OF FUNCTION DIRECTORY
================================================================================    
*/