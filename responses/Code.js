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

// === MAIN FORM EDIT HANDLER ===

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

// === TEMPLATE-BASED ROSTER CREATION ===

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

function getOrCreateRosterFromTemplate(teacher, rosterFolder, year, semesterName, registrationTimestamp) {
  try {
    UtilityScriptLibrary.debugLog("📄 Creating roster from script for teacher: " + teacher);
    
    var fileName = "QAMP " + year + " " + teacher;
    var files = rosterFolder.getFilesByName(fileName);
    
    if (files.hasNext()) {
      var existingFile = files.next();
      UtilityScriptLibrary.debugLog("📄 Roster file already exists: " + fileName);
      return SpreadsheetApp.openById(existingFile.getId());
    }

    // Create new spreadsheet entirely from script
    UtilityScriptLibrary.debugLog("🔧 Creating new spreadsheet from scratch");
    var newSpreadsheet = SpreadsheetApp.create(fileName);
    
    // Move the new file to the correct folder
    var newFile = DriveApp.getFileById(newSpreadsheet.getId());
    rosterFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
    
    UtilityScriptLibrary.debugLog("📄 Created new spreadsheet: " + fileName);
    
    // CHANGED: Pass timestamp to setup function
    try {
      setupCompleteRosterWorkbook(newSpreadsheet, teacher, year, semesterName, registrationTimestamp);
      UtilityScriptLibrary.debugLog("✅ Successfully set up complete roster workbook");
    } catch (setupError) {
      UtilityScriptLibrary.debugLog("⚠️ Error setting up roster workbook (will continue): " + setupError.message);
      // Continue anyway - at least we have the file created
    }
    
    // CRITICAL: Always update Teacher Roster Lookup, even if setup failed
    try {
      updateTeacherRosterLookup(teacher, newFile.getUrl());
      UtilityScriptLibrary.debugLog("✅ Updated teacher roster lookup");
    } catch (lookupError) {
      UtilityScriptLibrary.debugLog("⚠️ Could not update teacher roster lookup: " + lookupError.message);
    }
    
    UtilityScriptLibrary.debugLog("✅ Roster creation process completed for: " + teacher);
    return newSpreadsheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error creating roster from script: " + error.message);
    throw error;
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
    
    // 1.5 ADD CARRYOVER STUDENTS FROM PREVIOUS SEMESTER
    try {
      var carryoverCount = addCarryoverStudentsToNewRoster(spreadsheet, rosterSheet, semesterName);
      if (carryoverCount > 0) {
        UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - Added " + carryoverCount + " carryover students from previous semester");
      } else {
        UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - No carryover students found or no previous semester roster");
      }
    } catch (carryoverError) {
      UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - WARNING: Error adding carryover students: " + carryoverError.message);
      // Non-critical - continue with roster creation
    }
    
    // 2. Create attendance sheet for CURRENT SEMESTER MONTH (empty)
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
    } catch (activeError) {
      UtilityScriptLibrary.debugLog("setupCompleteRosterWorkbook - WARNING: Could not set active sheet: " + activeError.message);
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
    setupRosterTemplateProtection(sheet);
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    UtilityScriptLibrary.debugLog("✅ Roster template formatting applied");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Error in formatting: " + error.message);
  }
}

function setupRosterTemplateProtection(sheet) {
  try {
    // Protect admin columns (E through U) with warning - UPDATED range
    var adminRange = sheet.getRange(1, 5, sheet.getMaxRows(), 17); // Columns E-U (17 columns)
    var protection = adminRange.protect();
    protection.setDescription('Admin columns - automated data only');
    protection.setWarningOnly(true);
    
    // Set up date validation for First Lesson Date column (B)
    var dateRange = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1);
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();
    dateRange.setDataValidation(dateRule);
    
    UtilityScriptLibrary.debugLog("✅ Roster protection and validation applied");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Error in roster protection: " + error.message);
  }
}

// === GROUP ASSIGNMENT INTEGRATION ===

function getAllTeachersWithGroupAssignments() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var teacherLookupSheet = ss.getSheetByName('Teacher Roster Lookup');
    
    if (!teacherLookupSheet) {
      UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'WARNING', 'Teacher Roster Lookup sheet not found', '', '');
      return [];
    }
    
    var data = teacherLookupSheet.getDataRange().getValues();
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'INFO', 'No teachers found in lookup', '', '');
      return [];
    }
    
    // Find columns using createColumnFinder
    var getCol = UtilityScriptLibrary.createColumnFinder(teacherLookupSheet);
    var nameCol = getCol('Teacher Name');
    var groupAssignmentCol = getCol('Group Assignment');
    
    if (nameCol === 0 || groupAssignmentCol === 0) {
      UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'ERROR', 'Required columns not found', 
                                   'Name col: ' + nameCol + ', Group col: ' + groupAssignmentCol, '');
      return [];
    }
    
    // Get teachers who have group assignments (non-empty Group Assignment column)
    var teachersWithGroups = [];
    for (var i = 1; i < data.length; i++) {
      var teacherName = data[i][nameCol - 1];
      var groupAssignment = data[i][groupAssignmentCol - 1];
      
      if (teacherName && String(teacherName).trim() !== '' && 
          groupAssignment && String(groupAssignment).trim() !== '') {
        teachersWithGroups.push(String(teacherName).trim());
      }
    }
    
    UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'INFO', 'Found teachers with groups', 
                                 'Count: ' + teachersWithGroups.length, '');
    
    return teachersWithGroups;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('getAllTeachersWithGroupAssignments', 'ERROR', 'Failed to get teachers', '', error.message);
    return [];
  }
}

function updateAllTeacherGroupAssignments() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'INFO', 'Starting batch group assignment update', '', '');
    
    // Get current semester from Calendar sheet (D2)
    var semesterName = getCurrentSemesterName();
    if (!semesterName) {
      ui.alert('Error: Could not determine current semester from Calendar sheet.');
      return;
    }
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'INFO', 'Found current semester', semesterName, '');
    
    // Get all teachers who have group assignments
    var allTeachers = getAllTeachersWithGroupAssignments();
    
    if (!allTeachers || allTeachers.length === 0) {
      ui.alert('No teachers with group assignments found.');
      return;
    }
    
    UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'INFO', 'Found teachers with groups', 
                                 'Count: ' + allTeachers.length + ', Teachers: ' + allTeachers.join(', '), '');
    
    // Update each teacher
    var successCount = 0;
    var errorCount = 0;
    var errors = [];
    
    for (var i = 0; i < allTeachers.length; i++) {
      var teacherName = allTeachers[i];
      try {
        updateGroupAssignmentsForCurrentMonth(teacherName, semesterName);
        successCount++;
      } catch (error) {
        errorCount++;
        errors.push(teacherName + ': ' + error.message);
        UtilityScriptLibrary.debugLog('updateAllTeacherGroupAssignments', 'ERROR', 'Failed for teacher', 
                                     teacherName, error.message);
      }
    }
    
    // Show results
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

// === ROSTER AND ATTENDANCE UPDATE HANDLERS ===
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

function addStudentToAttendanceSheets(rosterWorkbook, studentData, semesterName, referenceDate) {
  try {
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", "Starting attendance addition", 
                          "Semester: " + semesterName, "");
    
    // Use referenceDate if provided, otherwise use current date
    var today = referenceDate || new Date();
    var currentMonthIndex = today.getMonth();  // 0-11
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", "Using date for attendance", 
                          "Date: " + today.toDateString() + ", Month: " + monthNames[currentMonthIndex], "");
    
    // Get semester date range
    var semesterDates = UtilityScriptLibrary.getSemesterDates(semesterName);
    if (!semesterDates) {
      UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "WARNING", 
                            "Could not find semester dates - skipping attendance addition", 
                            "Semester: " + semesterName, "");
      return;
    }
    
    var sheets = getMostRecentMonthSheets(rosterWorkbook);
    
    // Check if attendance sheets exist
    if (sheets.mostRecent && sheets.mostRecent.monthIndex >= currentMonthIndex) {
      // Teacher HAS current/future sheets - check if they belong to the same semester
      
      // Helper function to check if a month sheet belongs to this semester
      var isSheetInSemester = function(sheetMonthIndex) {
        // Create a date for the 1st of the month to compare
        var sheetDate = new Date(today.getFullYear(), sheetMonthIndex, 1);
        
        // Check if this date falls within semester range
        return sheetDate >= semesterDates.start && sheetDate <= semesterDates.end;
      };
      
      // Check most recent sheet
      if (isSheetInSemester(sheets.mostRecent.monthIndex)) {
        UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", 
                              "Most recent sheet belongs to semester - adding student", 
                              "Sheet: " + sheets.mostRecent.name + ", Semester: " + semesterName, "");
        addStudentToAttendanceSheet(sheets.mostRecent.sheet, studentData);
      } else {
        UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", 
                              "Most recent sheet belongs to different semester - skipping", 
                              "Sheet: " + sheets.mostRecent.name + ", Student semester: " + semesterName, "");
      }
      
      // Check second-most-recent if not invoiced AND in same semester
      if (sheets.secondMostRecent) {
        if (isSheetInSemester(sheets.secondMostRecent.monthIndex)) {
          var isInvoiced = hasMonthBeenInvoiced(sheets.secondMostRecent.sheet);
          if (!isInvoiced) {
            UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", 
                                  "Adding to second sheet (not invoiced, same semester)", 
                                  "Sheet: " + sheets.secondMostRecent.name, "");
            addStudentToAttendanceSheet(sheets.secondMostRecent.sheet, studentData);
          } else {
            UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", 
                                  "Second sheet already invoiced - skipping", 
                                  "Sheet: " + sheets.secondMostRecent.name, "");
          }
        } else {
          UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", 
                                "Second sheet belongs to different semester - skipping", 
                                "Sheet: " + sheets.secondMostRecent.name + ", Student semester: " + semesterName, "");
        }
      }
      
    } else {
      // No current/future sheets exist - SKIP attendance creation
      // Student will be picked up when bulk attendance creation is run
      UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "INFO", 
                            "No current attendance sheets found - skipping attendance creation", 
                            "Student will be added when monthly attendance sheets are created via bulk process", "");
    }
    
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "SUCCESS", "Completed attendance addition", "", "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentToAttendanceSheets", "ERROR", "Failed attendance addition", "", error.message);
    throw error;
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
    // The existing validation with setAllowInvalid(false) rejects empty cells
    // We need to update it to setAllowInvalid(true) for the entire column
    var maxRows = attendanceSheet.getMaxRows();
    var dateColumn = attendanceSheet.getRange(1, 3, maxRows, 1); // Column C (Date column)
    
    // Apply permissive date validation that allows empty cells
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)  // Allow empty cells - they can be filled in later
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
    setupStatusValidation(attendanceSheet, lastRowAfter);
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

function handleRosterUpdate(teacher, rosterFolder, studentData, year, semesterName) {
  try {
    UtilityScriptLibrary.debugLog("🔄 Starting handleRosterUpdate for teacher: " + teacher);
    
    if (!teacher || typeof teacher !== 'string' || teacher.trim() === '') {
      UtilityScriptLibrary.debugLog("❌ Skipping roster update: Missing teacher name.");
      return;
    }

    var studentId = studentData[studentData.length - 1];
    if (!studentId || typeof studentId !== 'string' || studentId.trim() === '') {
      UtilityScriptLibrary.debugLog("❌ Skipping roster update: Missing student ID.");
      return;
    }

    var rosterSS = getOrCreateRosterFromTemplate(teacher, rosterFolder, year, semesterName);
    var rosterSheet = rosterSS.getSheetByName("Roster");
    
    if (!rosterSheet) {
      throw new Error("Roster sheet not found in template-based spreadsheet");
    }

    UtilityScriptLibrary.debugLog("📊 Roster sheet retrieved: " + rosterSheet.getName());

    // Check if student already exists in roster
    var map = UtilityScriptLibrary.getHeaderMap(rosterSheet);
    var alreadyExists = checkIfStudentExists(rosterSheet, studentId, map);

    if (alreadyExists) {
      UtilityScriptLibrary.debugLog("⏩ Student with ID " + studentId + " already exists in " + teacher + "'s roster.");
      return;
    }

    UtilityScriptLibrary.debugLog("➕ Adding student to roster...");
    
    // Add student to roster sheet
    addStudentToNewRosterTemplate(rosterSheet, studentData, map);
    UtilityScriptLibrary.debugLog("✅ Added student " + studentData[1] + " " + studentData[0] + " to roster sheet.");
    
    // Add student to attendance sheet
    var currentSemesterMonth = UtilityScriptLibrary.getCurrentSemesterMonth(semesterName);
    var attendanceSheet = rosterSS.getSheetByName(currentSemesterMonth);
    
    if (attendanceSheet) {
      addStudentToAttendanceSheet(attendanceSheet, studentData);
      UtilityScriptLibrary.debugLog("✅ Added student " + studentData[1] + " " + studentData[0] + " to " + currentSemesterMonth + " attendance sheet.");
    } else {
      UtilityScriptLibrary.debugLog("ℹ️ Attendance sheet " + currentSemesterMonth + " not found - student will be added when sheets are created via bulk process");
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ ERROR in handleRosterUpdate: " + error.message);
    UtilityScriptLibrary.debugLog("❌ Stack trace: " + error.stack);
    throw error;
  }
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

function addStudentToRosterFromData(rosterSheet, studentInfo, headerMap) {
  try {
    UtilityScriptLibrary.debugLog("Adding student to roster from extracted data");
    
    // Create new row data array - 23 columns to match roster template
    var newRowData = new Array(23);
    for (var i = 0; i < newRowData.length; i++) {
      newRowData[i] = '';
    }
    
    // Map student data to roster columns
    newRowData[0] = false;                              // A: Contacted (checkbox)
    newRowData[1] = studentInfo.firstLessonDate || '';  // B: First Lesson Date
    newRowData[2] = studentInfo.firstLessonTime || '';  // C: First Lesson Time  
    newRowData[3] = studentInfo.comments || '';         // D: Comments
    newRowData[4] = studentInfo.lastName || '';         // E: Last Name
    newRowData[5] = studentInfo.firstName || '';        // F: First Name
    newRowData[6] = studentInfo.instrument || '';       // G: Instrument
    newRowData[7] = studentInfo.length || 30;           // H: Length (numeric)
    newRowData[8] = studentInfo.experience || '';       // I: Experience
    newRowData[9] = studentInfo.grade || '';            // J: Grade
    newRowData[10] = studentInfo.school || '';          // K: School
    newRowData[11] = studentInfo.schoolTeacher || '';   // L: School Teacher
    newRowData[12] = studentInfo.parentLastName || '';  // M: Parent Last Name
    newRowData[13] = studentInfo.parentFirstName || ''; // N: Parent First Name
    newRowData[14] = studentInfo.phone || '';           // O: Phone
    newRowData[15] = studentInfo.email || '';           // P: Email
    newRowData[16] = studentInfo.additionalContacts || ''; // Q: Additional contacts
    newRowData[17] = studentInfo.hoursRemaining || 0;   // R: Hours Remaining
    newRowData[18] = studentInfo.lessonsRemaining || 0; // S: Lessons Remaining
    newRowData[19] = 'Active';                          // T: Status
    newRowData[20] = studentInfo.studentId || '';       // U: Student ID
    newRowData[21] = '';                                // V: Admin Comments
    newRowData[22] = 'Reassigned: ' + new Date().toDateString();  // W: System Comments
    
    // Find target row (look for empty row or append)
    var lastRow = rosterSheet.getLastRow();
    var targetRow = lastRow + 1;
    
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
    
    // Insert the student data
    if (targetRow <= lastRow) {
      rosterSheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
      UtilityScriptLibrary.debugLog("Inserted student into empty row: " + targetRow);
    } else {
      rosterSheet.appendRow(newRowData);
      targetRow = rosterSheet.getLastRow();
      UtilityScriptLibrary.debugLog("Appended student to new row: " + targetRow);
    }
    
    // Set checkbox for Contacted column (A)
    rosterSheet.getRange(targetRow, 1).insertCheckboxes().setValue(false);
    
    // Apply alternating row color
    if (targetRow % 2 === 0) {
      rosterSheet.getRange(targetRow, 1, 1, newRowData.length).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      rosterSheet.getRange(targetRow, 1, 1, newRowData.length).setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }
    
    UtilityScriptLibrary.debugLog("Successfully added student to roster at row " + targetRow);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in addStudentToRosterFromData: " + error.message);
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
      // No attendance sheets exist - skip creation
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
        setupStatusValidation(attendanceSheet, lastRow);
        
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

function studentExistsInAttendanceSheet(attendanceSheet, studentId) {
  try {
    var data = attendanceSheet.getDataRange().getValues();
    
    // Check column A (Student ID column) for matching ID
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === studentId.toString().trim()) {
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("Error checking if student exists in attendance sheet: " + error.message);
    return false;
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
// === STUDENT AND PARENT PROCESSING ===

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
      var cleaned = String(cityZipRaw).trim().replace(/\s+/g, ' ').replace(/[.]/g, ',');
      var zipMatch = cleaned.match(/(\d{5})(?!.*\d)/);
      var zip = zipMatch ? zipMatch[1] : '';
      var cityPart = zip ? cleaned.substring(0, cleaned.lastIndexOf(zip)).trim() : cleaned;
      cityPart = cityPart.replace(/,\s*(NY|New York)?$/i, '').trim().replace(/,+$/, '');
      var city = cityPart.toLowerCase().split(' ').map(function(w) {
        return w.charAt(0).toUpperCase() + w.slice(1);
      }).join(' ');
      formData["Address City"] = city;
      formData["Address Zip Code"] = zip;
      formData["Address Formatted"] = (formData["Address Street"] || '') + '\n' + city + ', NY ' + zip;
    }

    // Generate parent key based on current form data
    var parentKey = UtilityScriptLibrary.generateKey(
      formData["Parent Last Name"] || '',
      formData["Parent First Name"] || '', 
      formData["Email"] || ''
    );

    // Get column positions dynamically
    var parentIdCol = getCol("Parent ID");
    var lookupCol = getCol("Parent Lookup");
    var studentIdsCol = getCol("Student IDs");
    var updatedCol = getCol("Updated");
    var parentGroupCol = getCol("Parent Group Interest");

    var parentId = existingParentId;
    var parentRow = UtilityScriptLibrary.findParentRow(parentsSheet, parentId, parentKey);

    UtilityScriptLibrary.debugLog("=== PARENT DUPLICATE CHECK ===");
    UtilityScriptLibrary.debugLog("Looking for parentId: '" + parentId + "' and parentKey: '" + parentKey + "'");
    UtilityScriptLibrary.debugLog("findParentRow result: " + parentRow);
    UtilityScriptLibrary.debugLog("=== END PARENT DEBUG ===");

    // Text fields to update
    var textFields = [
      "Parent Last Name", "Parent First Name", "Salutation", "Email", "Email 2", 
      "Phone", "Address Formatted", "Billing Preference", "Additional Contacts", "Referral"
    ];

    if (parentRow !== -1) {
      // UPDATE EXISTING PARENT
      UtilityScriptLibrary.debugLog("📄 Updating existing parent");
      var rowValues = parentsSheet.getRange(parentRow, 1, 1, parentsSheet.getLastColumn()).getValues()[0];
      var changesMade = false;
      
      // Update text fields
      for (var j = 0; j < textFields.length; j++) {
        var field = textFields[j];
        var col = getCol(field);
        if (col) {
          var newValue = formData[field] || '';
          var currentValue = rowValues[col - 1] || '';
          if (String(newValue) !== String(currentValue)) {
            parentsSheet.getRange(parentRow, col).setValue(newValue);
            changesMade = true;
            UtilityScriptLibrary.debugLog("Updated " + field + ": '" + currentValue + "' → '" + newValue + "'");
          }
        }
      }
      
      // Update Parent Lookup key to reflect current info (fixes email/name changes)
      if (lookupCol) {
        var currentLookup = String(rowValues[lookupCol - 1] || '').toLowerCase().trim();
        if (currentLookup !== parentKey) {
          parentsSheet.getRange(parentRow, lookupCol).setValue(parentKey);
          changesMade = true;
          UtilityScriptLibrary.debugLog("Updated Parent Lookup: '" + currentLookup + "' → '" + parentKey + "'");
        }
      }
      
      // Update Student IDs list (add if not already present)
      if (studentIdsCol) {
        var currentStudentIds = String(rowValues[studentIdsCol - 1] || '');
        var studentIdArray = currentStudentIds ? currentStudentIds.split(',').map(function(id) { return id.trim(); }) : [];
        
        // ES5 compatible check - use indexOf instead of includes
        if (studentIdArray.indexOf(studentId) === -1) {
          studentIdArray.push(studentId);
          var updatedStudentIds = studentIdArray.join(', ');
          parentsSheet.getRange(parentRow, studentIdsCol).setValue(updatedStudentIds);
          changesMade = true;
          UtilityScriptLibrary.debugLog("Added student ID " + studentId + " to parent's student list");
        }
      }
      
      // Update Parent Group Interest checkbox
      if (parentGroupCol) {
        var parentGroupCell = parentsSheet.getRange(parentRow, parentGroupCol);
        parentGroupCell.insertCheckboxes();
        var parentGroupValue = (formData["Parent Group Interest"] === "Yes");
        parentGroupCell.setValue(parentGroupValue);
      }

      // Update Updated checkbox
      if (updatedCol) {
        var updatedCell = parentsSheet.getRange(parentRow, updatedCol);
        updatedCell.insertCheckboxes();
        updatedCell.setValue(changesMade);
      }

    } else {
      // CREATE NEW PARENT
      UtilityScriptLibrary.debugLog("➕ Creating new parent");
      parentId = UtilityScriptLibrary.generateNextId(parentsSheet, "Parent ID", "P");
      
      var newRow = new Array(headers.length);
      for (var m = 0; m < headers.length; m++) {
        newRow[m] = '';
      }
      
      // Set basic parent information
      newRow[parentIdCol - 1] = parentId;
      newRow[lookupCol - 1] = parentKey;
      newRow[studentIdsCol - 1] = studentId;
      
      // Set text fields
      for (var n = 0; n < textFields.length; n++) {
        var field = textFields[n];
        var col = getCol(field);
        if (col) {
          newRow[col - 1] = formData[field] || '';
        }
      }
      
      parentsSheet.appendRow(newRow);
      var newRowIndex = parentsSheet.getLastRow();
      
      // Set Parent Group Interest checkbox
      if (parentGroupCol) {
        var parentGroupCell = parentsSheet.getRange(newRowIndex, parentGroupCol);
        parentGroupCell.insertCheckboxes();
        var parentGroupValue = (formData["Parent Group Interest"] === "Yes");
        parentGroupCell.setValue(parentGroupValue);
      }
      
      // Set Updated checkbox  
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
      standardizedAge = 'Child'; // Default to Child for safety
    }
    
    UtilityScriptLibrary.debugLog("Age conversion: Original='" + ageResponse + "', Standardized='" + standardizedAge + "'");

    if (studentRow !== -1) {
      UtilityScriptLibrary.debugLog("📄 Updating existing student");
      var rowData = contactsSheet.getRange(studentRow, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
      studentId = String(rowData[getCol("Student ID") - 1] || '');
      parentId = String(rowData[getCol("Parent ID") - 1] || '');

      // Update currently registered status
      var registeredCol = getCol("Currently Registered");
      if (registeredCol) {
        var checkboxCell = contactsSheet.getRange(studentRow, registeredCol);
        checkboxCell.insertCheckboxes();
        checkboxCell.setValue(true);
      }

      // Update teacher assignment
      if (getCol("Teacher")) {
        contactsSheet.getRange(studentRow, getCol("Teacher")).setValue(formData["Teacher"] || '');
      }
      
      // Update age with standardized value
      if (getCol("Age")) {
        contactsSheet.getRange(studentRow, getCol("Age")).setValue(standardizedAge);
      }
      
      // Update the 5 new fields for existing students
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

      // Set basic student information
      newRow[getCol("Student ID") - 1] = studentId;
      newRow[getCol("Student Last Name") - 1] = formData["Student Last Name"] || '';
      newRow[getCol("Student First Name") - 1] = formData["Student First Name"] || '';
      newRow[getCol("Instrument") - 1] = formData["Instrument"] || '';
      newRow[getCol("Teacher") - 1] = formData["Teacher"] || '';
      newRow[getCol("Age") - 1] = standardizedAge; // Use standardized age value
      newRow[getCol("First Enrollment Term") - 1] = enrollmentTerm || '';
      newRow[getCol("Student Lookup") - 1] = studentKey;

      // Add the 5 new fields for new students
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

      // Set up Currently Registered checkbox for new students
      var registeredCol = getCol("Currently Registered");
      if (registeredCol) {
        var newRowIndex = contactsSheet.getLastRow();
        var cell = contactsSheet.getRange(newRowIndex, registeredCol);
        cell.insertCheckboxes();
        cell.setValue(true);
      }
    }

    UtilityScriptLibrary.debugLog("✅ Completed processStudent - ID: " + studentId);
    return { studentId: studentId, parentId: parentId, studentRow: studentRow };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in processStudent: " + error.message);
    throw error;
  }
}

// === UTILITY FUNCTIONS ===

function applyTeacherDropdownToCurrentSemester() {
  try {
    UtilityScriptLibrary.debugLog('applyTeacherDropdownToCurrentSemester', 'INFO', 'Starting auto-apply teacher dropdown', '', '');
    
    // Get current semester name from calendar
    var currentSemester = getCurrentSemesterName();
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

function createCurrentMonthAttendance() {
  //UPDATED 11-5-25
  try {
    var workbook = SpreadsheetApp.getActiveSpreadsheet();
    var currentMonth = UtilityScriptLibrary.getMonthNameFromDate(new Date());
    
    generateAttendanceSheetFromRoster(workbook, currentMonth);
    SpreadsheetApp.getUi().alert('✅ ' + currentMonth + ' attendance sheet created successfully!');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function createTeacherInfoObject(dataRow) {
  return {
    teacherName: dataRow[0],
    rosterUrl: dataRow[1],
    teacherId: dataRow[2],
    displayName: dataRow[3],
    status: dataRow[4],
    lastUpdated: dataRow[5]
  };
}

function extractDisplayNameFromFullName(fullName) {
  try {
    if (!fullName || String(fullName).trim() === '') {
      return '';
    }
    
    var nameParts = String(fullName).trim().split(' ');
    if (nameParts.length === 0) {
      return '';
    }
    
    // Get the last part of the name and clean it (remove non-alphanumeric characters)
    var lastName = nameParts[nameParts.length - 1];
    return lastName.replace(/[^a-zA-Z0-9]/g, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error extracting display name: " + error.message);
    return '';
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

function findPreviousSemesterRoster(spreadsheet, currentSemesterName) {
  try {
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Looking for previous semester roster. Current: " + currentSemesterName);
    
    // Get Semester Metadata
    var semesterSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterSheet) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Semester Metadata sheet not found");
      return null;
    }
    
    var data = semesterSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var nameCol = -1, startDateCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().toLowerCase().trim();
      if (header === 'semester name') nameCol = i;
      if (header === 'start date') startDateCol = i;
    }
    
    if (nameCol === -1 || startDateCol === -1) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Required columns not found in Semester Metadata");
      return null;
    }
    
    // Find current semester's start date
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
    
    // Get all sheets in workbook
    var sheets = spreadsheet.getSheets();
    var existingRosterSheets = [];
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName.toLowerCase().indexOf('roster') !== -1) {
        existingRosterSheets.push(sheetName);
      }
    }
    
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Found " + existingRosterSheets.length + " roster sheets in workbook");
    
    // Find all previous semesters that have roster sheets in this workbook
    var candidates = [];
    for (var i = 1; i < data.length; i++) {
      var semesterName = data[i][nameCol];
      var startDate = new Date(data[i][startDateCol]);
      
      // Must be before current semester
      if (startDate >= currentStartDate) continue;
      
      // Must have a roster sheet in this workbook
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
        candidates.push({
          semesterName: semesterName,
          startDate: startDate,
          sheetName: expectedSheetName
        });
      }
    }
    
    if (candidates.length === 0) {
      UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - No previous semester rosters found in this workbook");
      return null;
    }
    
    // Sort by start date descending and pick most recent
    candidates.sort(function(a, b) {
      return b.startDate - a.startDate;
    });
    
    var previousSemester = candidates[0];
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - Found previous semester: " + previousSemester.semesterName + " (Sheet: " + previousSemester.sheetName + ")");
    
    return previousSemester.sheetName;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("findPreviousSemesterRoster - ERROR: " + error.message);
    return null;
  }
}



function findTeacherInEnhancedRosterLookup(lookupSheet, teacherName) {
  try {
    if (lookupSheet.getLastRow() <= 1) {
      return -1;
    }
    
    var data = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 4).getValues();
    var searchName = String(teacherName).trim();
    
    for (var i = 0; i < data.length; i++) {
      var displayName = String(data[i][3]).trim();
      
      if (displayName === searchName) {
        UtilityScriptLibrary.debugLog("✅ Found teacher by display name: " + displayName + " (row " + (i + 2) + ")");
        return i + 2;
      }
    }
    
    UtilityScriptLibrary.debugLog("⚠️ Teacher not found in lookup: " + searchName);
    return -1;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error finding teacher in enhanced roster lookup: " + error.message);
    return -1;
  }
}

function generateAttendanceSheetFromRoster(teacherWorkbook, monthName) {
  try {
    UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - Generating ' + monthName + ' attendance sheet');
    
    // Get roster data
    var rosterSheet = teacherWorkbook.getSheetByName('Roster');
    if (!rosterSheet) {
      throw new Error('Roster sheet not found');
    }
    
    var rosterData = UtilityScriptLibrary.extractRosterData(rosterSheet);
    
    // Check if month sheet already exists
    var existingSheet = teacherWorkbook.getSheetByName(monthName);
    if (existingSheet) {
      UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - Sheet ' + monthName + ' already exists, skipping');
      return existingSheet;
    }
    
    // Create the attendance sheet (works even with empty roster)
    var attendanceSheet = UtilityScriptLibrary.createMonthlyAttendanceSheet(teacherWorkbook, monthName, rosterData);
    
    UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - ✅ Generated ' + monthName + ' attendance sheet with ' + rosterData.length + ' students');
    return attendanceSheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('generateAttendanceSheetFromRoster - ❌ Error generating attendance sheet: ' + error.message);
    throw error;
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

// === MAIN ENTRY POINT ===
function reassignStudentToNewTeacher() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    UtilityScriptLibrary.debugLog("=== STARTING STUDENT REASSIGNMENT ===");
    
    // Get current semester
    var currentSemester = getCurrentSemesterName();
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
      'reassignStep2_selectStudents'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in reassignStudentToNewTeacher: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Reassignment failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === STEP 2: SELECT STUDENTS (UPDATED WITH CHECKBOX DIALOG) ===
function reassignStep2_selectStudents(oldTeacherDisplay) {
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
    
    UtilityScriptLibrary.debugLog("Old teacher full name: " + oldTeacherInfo.teacherName + ", URL: " + oldTeacherInfo.rosterUrl);
    
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
      'reassignStep2b_processStudentSelection'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in reassignStep2_selectStudents: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 2 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === STEP 2B: PROCESS STUDENT SELECTION ===
function reassignStep2b_processStudentSelection(selectedIndices) {
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
    reassignStep3_selectNewTeacher();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in reassignStep2b_processStudentSelection: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Student selection failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === STUDENT CHECKBOX DIALOG ===
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

// === STEP 3: SELECT NEW TEACHER ===
function reassignStep3_selectNewTeacher() {
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
      'reassignStep4_enterEffectiveDate'
    );
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in reassignStep3_selectNewTeacher: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 3 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === STEP 4: ENTER EFFECTIVE DATE ===
function reassignStep4_enterEffectiveDate(newTeacherDisplay) {
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
    reassignStep5_processReassignment();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("ERROR in reassignStep4_enterEffectiveDate: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 4 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === STEP 5: PROCESS REASSIGNMENT ===
function reassignStep5_processReassignment() {
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
    
    // FIXED: Use displayName instead of teacherName for roster file lookup
    var newRosterWorkbook = getOrCreateRosterFromTemplate(newTeacherInfo.displayName, rosterFolder, year, currentSemester);
    var newRosterSheet = findSemesterRoster(newRosterWorkbook, currentSemester);
    
    if (!newRosterSheet) {
      throw new Error("Could not find or create roster sheet for " + currentSemester + " in new teacher's workbook");
    }
    
    UtilityScriptLibrary.debugLog("✅ Got/created new teacher roster: " + newRosterSheet.getName());
    
    // Update Teacher Roster Lookup with URL - FIXED: use displayName instead of teacherName
    var newRosterFile = DriveApp.getFileById(newRosterWorkbook.getId());
    var newRosterUrl = newRosterFile.getUrl();
    
    UtilityScriptLibrary.debugLog("📝 Updating Teacher Roster Lookup for display name: " + newTeacherInfo.displayName);
    updateTeacherRosterLookup(newTeacherInfo.displayName, newRosterUrl);
    UtilityScriptLibrary.debugLog("✅ Teacher Roster Lookup update completed");
    
    // Open old teacher's roster
    var oldRosterWorkbook = SpreadsheetApp.openById(oldRosterWorkbookId);
    var oldRosterSheet = oldRosterWorkbook.getSheetByName(oldRosterSheetName);
    
    if (!oldRosterSheet) {
      throw new Error("Could not find old roster sheet: " + oldRosterSheetName);
    }
    
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
        var oldHeaderMap = UtilityScriptLibrary.getHeaderMap(oldRosterSheet);
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
        var newHeaderMap = UtilityScriptLibrary.getHeaderMap(newRosterSheet);
        var studentExists = checkIfStudentExists(newRosterSheet, student.studentId, newHeaderMap);
        
        if (studentExists) {
          UtilityScriptLibrary.debugLog("Student already exists in new roster - skipping roster addition");
          skipCount++;
        } else {
          // Add to new roster
          addStudentToRosterFromData(newRosterSheet, student, newHeaderMap);
          UtilityScriptLibrary.debugLog("Added to new teacher's roster");
        }
        
        // Add to attendance sheets
        addStudentToAttendanceSheetsFromDate(newRosterWorkbook, student, effectiveDate);
        
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
    UtilityScriptLibrary.debugLog("ERROR in reassignStep5_processReassignment: " + error.message);
    SpreadsheetApp.getUi().alert('Error', 'Step 5 failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// === DIALOG FUNCTION ===
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

function getActiveTeachersForDropdown() {
  try {
    UtilityScriptLibrary.debugLog("🎯 Getting active and potential teachers for dropdown");
    
    var lookupSheet = getTeacherRosterLookupSheet();
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog("⚠️ Teacher Roster Lookup sheet not found or empty");
      return [];
    }
    
    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var displayNameCol = getCol('Display Name');
    var statusCol = getCol('Status');
    
    if (!displayNameCol || !statusCol) {
      UtilityScriptLibrary.debugLog("⚠️ Required columns (Display Name or Status) not found");
      return [];
    }
    
    var lastCol = Math.max(displayNameCol, statusCol);
    var data = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, lastCol).getValues();
    var displayNames = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var status = String(row[statusCol - 1]).trim().toLowerCase();
      var displayName = row[displayNameCol - 1];
      
      if ((status === 'active' || status === 'potential') && displayName && String(displayName).trim() !== '') {
        displayNames.push(String(displayName).trim());
      }
    }
    
    displayNames = displayNames.sort().filter(function(value, index, self) {
      return self.indexOf(value) === index;
    });
    
    UtilityScriptLibrary.debugLog("✅ Found " + displayNames.length + " teachers for dropdown: " + displayNames.join(', '));
    return displayNames;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error getting teachers for dropdown: " + error.message);
    return [];
  }
}

function getAppropriateAttendanceMonth(semesterName, registrationTimestamp) {
  try {
    UtilityScriptLibrary.debugLog("getAppropriateAttendanceMonth", "INFO", "Determining appropriate attendance month", "Semester: " + semesterName, "");
    
    // Get semester start date
    var semesterMetadataSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      UtilityScriptLibrary.debugLog("getAppropriateAttendanceMonth", "WARNING", "Semester Metadata sheet not found", "", "");
      return "January"; // Fallback
    }
    
    var getCol = UtilityScriptLibrary.createColumnFinder(semesterMetadataSheet);
    var nameCol = getCol('Semester Name');
    var startCol = getCol('Start Date');
    
    if (nameCol === 0 || startCol === 0) {
      UtilityScriptLibrary.debugLog("getAppropriateAttendanceMonth", "WARNING", "Required columns not found", "", "");
      return "January"; // Fallback
    }
    
    var data = semesterMetadataSheet.getDataRange().getValues();
    var semesterStartDate = null;
    
    // Find semester start date
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol - 1] && row[nameCol - 1].toString().trim() === semesterName) {
        semesterStartDate = new Date(row[startCol - 1]);
        break;
      }
    }
    
    if (!semesterStartDate) {
      UtilityScriptLibrary.debugLog("getAppropriateAttendanceMonth", "WARNING", "Semester not found", "Semester: " + semesterName, "");
      return "January"; // Fallback
    }
    
    // Get registration date (use timestamp if provided, otherwise use current date)
    var registrationDate = registrationTimestamp ? new Date(registrationTimestamp) : new Date();
    
    // Use whichever date is LATER (registration or semester start)
    var appropriateDate = registrationDate > semesterStartDate ? registrationDate : semesterStartDate;
    
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var monthName = monthNames[appropriateDate.getMonth()];
    
    UtilityScriptLibrary.debugLog("getAppropriateAttendanceMonth", "INFO", "Determined month", 
                          "Semester: " + semesterName + ", Start: " + semesterStartDate.toDateString() + 
                          ", Registration: " + registrationDate.toDateString() + ", Using: " + monthName, "");
    
    return monthName;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getAppropriateAttendanceMonth", "ERROR", "Error determining month", 
                          "Semester: " + semesterName, error.message);
    return "January"; // Fallback
  }
}

function getCurrentSemesterName() {
  try {
    var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Calendar');
    if (!calendarSheet || calendarSheet.getLastRow() < 2) {
      UtilityScriptLibrary.debugLog('getCurrentSemesterName', 'WARNING', 'Calendar sheet not found or empty', '', '');
      return null;
    }
    
    // Current semester is always in D2 (column 4, row 2)
    var currentSemester = calendarSheet.getRange(2, 4).getValue();
    
    if (!currentSemester || String(currentSemester).trim() === '') {
      UtilityScriptLibrary.debugLog('getCurrentSemesterName', 'WARNING', 'No current semester in calendar', '', '');
      return null;
    }
    
    var semesterName = String(currentSemester).trim();
    UtilityScriptLibrary.debugLog('getCurrentSemesterName', 'INFO', 'Found current semester', semesterName, '');
    return semesterName;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('getCurrentSemesterName', 'ERROR', 'Failed to get current semester', '', error.message);
    return null;
  }
}

function getExistingGroupIds(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    var existingGroupIds = [];
    
    // Check Student ID column (column A, index 0)
    for (var i = 1; i < data.length; i++) {
      var studentId = data[i][0];
      if (studentId && String(studentId).match(/^G\d{4}$/)) {
        // This is a group ID (G followed by exactly 4 digits)
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
    
    UtilityScriptLibrary.debugLog('getExistingGroupIds', 'DEBUG', 'Found existing group IDs', 
                                 'Count: ' + existingGroupIds.length + ', IDs: ' + existingGroupIds.join(', '), '');
    
    return existingGroupIds;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('getExistingGroupIds', 'ERROR', 'Error getting existing group IDs', '', error.message);
    return [];
  }
}

function getMostRecentMonthSheet(workbook) {
  try {
    var sheets = workbook.getSheet();
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

function getTeacherIdFromContacts(teacherName) {
  try {
    UtilityScriptLibrary.debugLog("🔍 Looking up Teacher ID for: '" + teacherName + "'");
    
    var contactsWorkbook = UtilityScriptLibrary.getWorkbook('contacts');
    var teachersSheet = contactsWorkbook.getSheetByName('Teachers and Admin');
    
    if (!teachersSheet) {
      UtilityScriptLibrary.debugLog("❌ Teachers and Admin sheet not found");
      return 'not found';
    }
    
    UtilityScriptLibrary.debugLog("🔍 Searching by last name only: '" + teacherName + "'");
    
    // Get header map
    var headerMap = UtilityScriptLibrary.getHeaderMap(teachersSheet);
    var data = teachersSheet.getDataRange().getValues();
    
    var lastNameCol = headerMap["last name"];
    var teacherIdCol = headerMap["teacher id"];
    
    UtilityScriptLibrary.debugLog("📍 Column indices - Last Name: " + lastNameCol + ", Teacher ID: " + teacherIdCol);
    
    if (!lastNameCol || !teacherIdCol) {
      UtilityScriptLibrary.debugLog("❌ Required columns not found in Teachers and Admin sheet");
      return 'not found';
    }
    
    // Search for matching teacher by last name only
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowLastName = (row[lastNameCol - 1] || '').toString().trim();
      var rowTeacherId = (row[teacherIdCol - 1] || '').toString().trim();
      
      UtilityScriptLibrary.debugLog("Checking row " + i + " - Last Name: '" + rowLastName + "', ID: '" + rowTeacherId + "'");
      
      if (rowLastName === teacherName) {
        UtilityScriptLibrary.debugLog("✅ Found matching teacher with ID: " + rowTeacherId);
        return rowTeacherId || 'no ID';
      }
    }
    
    UtilityScriptLibrary.debugLog("❌ No matching teacher found with last name: '" + teacherName + "'");
    return 'not found';
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error getting Teacher ID from contacts: " + error.message);
    return 'error: ' + error.message;
  }
}

function getTeacherInfoByDisplayName(displayName) {
  try {
    UtilityScriptLibrary.debugLog("🔍 Looking up teacher info for display name: " + displayName);
    
    var lookupSheet = getTeacherRosterLookupSheet();
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog("⚠️ Teacher Roster Lookup sheet not found or empty");
      return null;
    }
    
    // Use dynamic column finder
    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var teacherNameCol = getCol('Teacher Name');
    var rosterUrlCol = getCol('Roster URL');
    var teacherIdCol = getCol('Teacher ID');
    var displayNameCol = getCol('Display Name');
    var statusCol = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');
    
    if (!displayNameCol) {
      UtilityScriptLibrary.debugLog("⚠️ Display Name column not found");
      return null;
    }
    
    // Get all data
    var lastCol = Math.max(teacherNameCol || 1, rosterUrlCol || 1, teacherIdCol || 1, displayNameCol, statusCol || 1, lastUpdatedCol || 1);
    var data = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, lastCol).getValues();
    
    // Search for teacher by display name
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowDisplayName = String(row[displayNameCol - 1]).trim();
      
      if (rowDisplayName === displayName) {
        var teacherInfo = {
          teacherName: teacherNameCol ? row[teacherNameCol - 1] : '',
          rosterUrl: rosterUrlCol ? row[rosterUrlCol - 1] : '',
          teacherId: teacherIdCol ? row[teacherIdCol - 1] : '',
          displayName: row[displayNameCol - 1],
          status: statusCol ? row[statusCol - 1] : '',
          lastUpdated: lastUpdatedCol ? row[lastUpdatedCol - 1] : ''
        };
        
        UtilityScriptLibrary.debugLog("✅ Found teacher: " + teacherInfo.teacherName);
        return teacherInfo;
      }
    }
    
    UtilityScriptLibrary.debugLog("⚠️ Teacher not found with display name: " + displayName);
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error getting teacher info by display name: " + error.message);
    return null;
  }
}

function getTeacherInfoByFullName(teacherName) {
  try {
    UtilityScriptLibrary.debugLog("🔍 Looking up teacher info for full name: " + teacherName);
    
    var lookupSheet = getTeacherRosterLookupSheet();
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog("⚠️ Teacher Roster Lookup sheet not found or empty");
      return null;
    }
    
    var teacherRow = findTeacherInEnhancedRosterLookup(lookupSheet, teacherName);
    
    if (teacherRow !== -1) {
      var data = lookupSheet.getRange(teacherRow, 1, 1, 6).getValues()[0];
      
      var teacherInfo = createTeacherInfoObject(data);
      
      UtilityScriptLibrary.debugLog("✅ Found teacher info: " + JSON.stringify(teacherInfo));
      return teacherInfo;
    }
    
    UtilityScriptLibrary.debugLog("❌ Teacher not found with full name: " + teacherName);
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error getting teacher info by full name: " + error.message);
    return null;
  }
}

function getTeacherRosterLookupSheet() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return activeSpreadsheet.getSheetByName("Teacher Roster Lookup");
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

function markTeacherAsActiveInLookup(teacherName) {
  try {
    UtilityScriptLibrary.debugLog("✅ Marking teacher as active: " + teacherName);
    
    var lookupSheet = getTeacherRosterLookupSheet();
    
    if (!lookupSheet) {
      UtilityScriptLibrary.debugLog("⚠️ Teacher Roster Lookup sheet not found");
      return;
    }
    
    var teacherRow = findTeacherInEnhancedRosterLookup(lookupSheet, teacherName);
    
    if (teacherRow !== -1) {
      // Update status to Active and timestamp
      lookupSheet.getRange(teacherRow, 5).setValue('Active');     // Column E: Status
      lookupSheet.getRange(teacherRow, 6).setValue(new Date());   // Column F: Last Updated
      
      UtilityScriptLibrary.debugLog("✅ Successfully marked teacher as active");
    } else {
      UtilityScriptLibrary.debugLog("⚠️ Teacher not found in lookup - cannot mark as active");
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error marking teacher as active: " + error.message);
  }
}

function processRoster(formData, sheet, editedRow, headerMap, fieldMap, studentId, rosterFolder, year, semesterName) {
  try {
    var teacher = formData["Teacher"];
    UtilityScriptLibrary.debugLog("processRoster - Starting processRoster for teacher: " + teacher + ", Semester: " + semesterName);
    
    if (!teacher || teacher.trim() === "") {
      UtilityScriptLibrary.debugLog("processRoster - WARNING: Skipping roster update: Missing teacher name");
      return;
    }

    if (!studentId || studentId.toString().trim() === "") {
      UtilityScriptLibrary.debugLog("processRoster - WARNING: Skipping roster update: Missing student ID");
      return;
    }

    // CRITICAL: Extract timestamp from formData for historical data support
    // Try multiple possible field names for timestamp
    var registrationTimestamp = formData["Timestamp"] || formData["timestamp"] || formData["Registration Date"] || null;
    
    if (registrationTimestamp) {
      UtilityScriptLibrary.debugLog("processRoster - Using registration timestamp: " + registrationTimestamp);
    } else {
      UtilityScriptLibrary.debugLog("processRoster - No timestamp found, will use current date");
    }

    // Get or create teacher roster workbook - PASS THE TIMESTAMP
    var rosterSS = getOrCreateRosterFromTemplate(teacher, rosterFolder, year, semesterName, registrationTimestamp);
    if (!rosterSS) {
      throw new Error("Could not create or access roster workbook for " + teacher);
    }

    // Add student to semester-specific roster - PASS THE TIMESTAMP
    addStudentToSemesterRoster(rosterSS, formData, studentId, semesterName, registrationTimestamp);
    
    UtilityScriptLibrary.debugLog("processRoster - Successfully processed roster for Teacher: " + teacher + ", Student: " + studentId + ", Semester: " + semesterName);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("processRoster - ERROR: Error in processRoster for Teacher: " + teacher + ", Semester: " + semesterName + ". Error: " + error.message);
    throw error;
  }
}

function refreshCurrentSemesterTeacherDropdown() {
  try {
    UtilityScriptLibrary.debugLog('refreshCurrentSemesterTeacherDropdown', 'INFO', 'Manual refresh triggered', '', '');
    
    var currentSemester = getCurrentSemesterName();
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
      UtilityScriptLibrary.debugLog('setupStatusValidation', 'ERROR', 'Student ID column not found', '', '');
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
        if (dateValue && typeof dateValue === 'string') {
          var monthNames = ['january', 'february', 'march', 'april', 'may', 'june',
                           'july', 'august', 'september', 'october', 'november', 'december'];
          var dateLower = dateValue.toLowerCase();
          for (var m = 0; m < monthNames.length; m++) {
            if (dateLower.indexOf(monthNames[m]) !== -1) {
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
    
    UtilityScriptLibrary.debugLog('setupStatusValidation', 'INFO', 'Applied status validation', 
                                 'Rows processed: ' + processedRows, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('setupStatusValidation', 'ERROR', 'Error setting up status validation', '', error.message);
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

function updateGroupAssignmentsForCurrentMonth(teacherName, semesterName) {
  try {
    UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'Starting group assignment update', 
                                 'Teacher: ' + teacherName + ', Semester: ' + semesterName, '');
    
    // Get teacher's current group assignments
    var groupAssignments = UtilityScriptLibrary.getTeacherGroupAssignments(teacherName);
    if (!groupAssignments || groupAssignments.length === 0) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'No group assignments found for teacher', 
                                   'Teacher: ' + teacherName, '');
      return;
    }
    
    // Get teacher info from Teacher Roster Lookup (includes roster URL)
    var teacherInfo = getTeacherInfoByFullName(teacherName);
    if (!teacherInfo || !teacherInfo.rosterUrl) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'Teacher roster URL not found', 
                                   'Teacher: ' + teacherName, '');
      return;
    }
    
    // Open teacher's roster workbook directly by URL
    var rosterSS = SpreadsheetApp.openByUrl(teacherInfo.rosterUrl);
    
    if (!rosterSS) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'Could not open teacher workbook', 
                                   'Teacher: ' + teacherName, '');
      return;
    }
    
    // Get the most recent month attendance sheet
    var attendanceSheet = getMostRecentMonthSheet(rosterSS);
    if (!attendanceSheet) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'No month attendance sheet found', 
                                   'Teacher: ' + teacherName, '');
      return;
    }
    
    // Check which G#### sections already exist in the sheet
    var existingGroupIds = getExistingGroupIds(attendanceSheet);
    
    // Filter out group assignments that already exist
    var newGroupAssignments = [];
    for (var i = 0; i < groupAssignments.length; i++) {
      var groupId = groupAssignments[i].groupId;
      var alreadyExists = false;
      for (var j = 0; j < existingGroupIds.length; j++) {
        if (existingGroupIds[j] === groupId) {
          alreadyExists = true;
          break;
        }
      }
      if (!alreadyExists) {
        newGroupAssignments.push(groupAssignments[i]);
      }
    }
    
    if (newGroupAssignments.length === 0) {
      UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'All group assignments already exist', 
                                   'Teacher: ' + teacherName, '');
      return;
    }
    
    // Add new group sections
    UtilityScriptLibrary.createGroupSections(attendanceSheet, newGroupAssignments);
    
    // Apply status validation to all rows (students and groups)
    var totalRows = attendanceSheet.getLastRow();
    if (totalRows > 1) {
      setupStatusValidation(attendanceSheet, totalRows);
    }
    
    UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'INFO', 'Successfully added group assignments', 
                                 'Teacher: ' + teacherName + ', Added: ' + newGroupAssignments.length, '');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('updateGroupAssignmentsForCurrentMonth', 'ERROR', 'Failed to update group assignments', 
                                 'Teacher: ' + teacherName, error.message);
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

function updateTeacherRosterLookup(teacherName, fileUrl) {
  try {
    UtilityScriptLibrary.debugLog("🔍 Updating teacher roster lookup for: " + teacherName);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lookupSheet = ss.getSheetByName("Teacher Roster Lookup");
    
    if (!lookupSheet) {
      UtilityScriptLibrary.debugLog("❌ Teacher Roster Lookup sheet not found");
      return;
    }
    
    var teacherRow = findTeacherInEnhancedRosterLookup(lookupSheet, teacherName);
    
    if (teacherRow === -1) {
      UtilityScriptLibrary.debugLog("❌ Teacher not found in lookup: " + teacherName);
      return;
    }
    
    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var rosterUrlCol = getCol('Roster URL');
    var statusCol = getCol('Status');
    var lastUpdatedCol = getCol('Last Updated');
    
    if (rosterUrlCol) lookupSheet.getRange(teacherRow, rosterUrlCol).setValue(fileUrl);
    if (statusCol) lookupSheet.getRange(teacherRow, statusCol).setValue('active');
    if (lastUpdatedCol) lookupSheet.getRange(teacherRow, lastUpdatedCol).setValue(new Date());
    
    UtilityScriptLibrary.debugLog("✅ Updated existing teacher roster lookup entry");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in updateTeacherRosterLookup: " + error.message);
  }
}

function validateEnhancedTeacherRosterLookup() {
  try {
    UtilityScriptLibrary.debugLog("🔍 Validating Enhanced Teacher Roster Lookup structure");
    
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var lookupSheet = getTeacherRosterLookupSheet();
    
    if (!lookupSheet) {
      UtilityScriptLibrary.debugLog("⚠️ Teacher Roster Lookup sheet not found - creating new one");
      lookupSheet = createEnhancedTeacherRosterLookupSheet(activeSpreadsheet);
      return true;
    }
    
    var headers = lookupSheet.getRange(1, 1, 1, lookupSheet.getLastColumn()).getValues()[0];
    var expectedHeaders = ["Teacher Name", "Roster URL", "Teacher ID", "Display Name", "Group Assignment", "Status", "Last Updated"];
    
    var needsRepair = false;
    
    if (headers.length < 7) {
      UtilityScriptLibrary.debugLog("⚠️ Header structure incomplete - has " + headers.length + " columns, needs 7");
      needsRepair = true;
    }
    
    for (var i = 0; i < expectedHeaders.length && i < headers.length; i++) {
      if (String(headers[i]).trim() !== expectedHeaders[i]) {
        UtilityScriptLibrary.debugLog("⚠️ Header mismatch at column " + (i + 1) + ": expected '" + expectedHeaders[i] + "', found '" + headers[i] + "'");
        needsRepair = true;
      }
    }
    
    if (needsRepair) {
      UtilityScriptLibrary.debugLog("🔧 Repairing Enhanced Teacher Roster Lookup structure");
      
      var existingData = [];
      if (lookupSheet.getLastRow() > 1) {
        existingData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, Math.min(lookupSheet.getLastColumn(), 7)).getValues();
      }
      
      lookupSheet.clear();
      lookupSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      
      try {
        UtilityScriptLibrary.styleHeaderRow(lookupSheet, expectedHeaders);
      } catch (styleError) {
        var headerRange = lookupSheet.getRange(1, 1, 1, expectedHeaders.length);
        headerRange.setFontWeight("bold")
                   .setBackground(UtilityScriptLibrary.STYLES.HEADER.background)
                   .setFontColor(UtilityScriptLibrary.STYLES.HEADER.text);
      }
      
      if (existingData.length > 0) {
        for (var i = 0; i < existingData.length; i++) {
          var row = existingData[i];
          var teacherName = row[0] || '';
          var rosterUrl = row[1] || '';
          var teacherId = row[2] || '';
          var displayName = row[3] || extractDisplayNameFromFullName(teacherName);
          var groupAssignment = row[4] || '';
          var status = row[5] || (rosterUrl ? 'active' : 'potential');
          var lastUpdated = row[6] || new Date();
          
          if (teacherName && String(teacherName).trim() !== '') {
            lookupSheet.appendRow([teacherName, rosterUrl, teacherId, displayName, groupAssignment, status, lastUpdated]);
          }
        }
      }
      
      UtilityScriptLibrary.debugLog("✅ Enhanced Teacher Roster Lookup structure repaired");
    } else {
      UtilityScriptLibrary.debugLog("✅ Enhanced Teacher Roster Lookup structure is valid");
    }
    
    return true;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error validating Enhanced Teacher Roster Lookup: " + error.message);
    return false;
  }
}

// ============================================
// VERIFY: By Google Drive ID (folder or file)
// ============================================
// ============================================
// VERIFY: By Google Drive ID (folder or file)
// ============================================
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

// ============================================
// HELPER: Clear Reports (run this first)
// ============================================
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

// ============================================
// SHARED FUNCTIONS
// ============================================

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

function createNewYearWorkbooksWithContinuingStudents() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Step 1: Verify teacher status
    var verifyResponse = ui.alert(
      'Teacher Status Verification',
      'Have you verified teacher status in Teacher Roster Lookup?\n\n' +
      '(Only active teachers will get new workbooks)',
      ui.ButtonSet.YES_NO
    );
    
    if (verifyResponse !== ui.Button.YES) {
      ui.alert('Cancelled', 'Please verify teacher status first.', ui.ButtonSet.OK);
      return;
    }
    
    // Step 2: Get most recent semester from Billing
    var semesterMetadata = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterMetadata) {
      throw new Error('Semester Metadata sheet not found in Billing');
    }
    
    var semesterData = semesterMetadata.getDataRange().getValues();
    var headers = semesterData[0];
    
    var nameCol = -1, startCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (header === 'semester name' || header === 'semestername') nameCol = i;
      if (header === 'start date' || header === 'startdate') startCol = i;
    }
    
    if (nameCol === -1 || startCol === -1) {
      throw new Error('Required columns not found in Semester Metadata');
    }
    
    // Find most recent semester (latest start date)
    var mostRecentSemester = null;
    var mostRecentDate = null;
    
    for (var i = 1; i < semesterData.length; i++) {
      var semesterName = semesterData[i][nameCol];
      var startDate = new Date(semesterData[i][startCol]);
      
      if (!semesterName || !startDate) continue;
      
      if (!mostRecentDate || startDate > mostRecentDate) {
        mostRecentDate = startDate;
        mostRecentSemester = semesterName;
      }
    }
    
    if (!mostRecentSemester) {
      throw new Error('No valid semester found in Semester Metadata');
    }
    
    UtilityScriptLibrary.debugLog('Most recent semester: ' + mostRecentSemester);
    
    // Step 3: Extract year from semester name
    var yearMatch = mostRecentSemester.match(/\d{4}/);
    if (!yearMatch) {
      throw new Error('Could not extract year from semester name: ' + mostRecentSemester);
    }
    
    var newYear = yearMatch[0];
    var previousYear = String(parseInt(newYear) - 1);
    
    UtilityScriptLibrary.debugLog('New year: ' + newYear + ', Previous year: ' + previousYear);
    
    // Step 4: Confirm with user
    var confirm = ui.alert(
      'Create ' + newYear + ' Teacher Workbooks',
      'This will:\n' +
      '• Create new QAMP ' + newYear + ' workbooks\n' +
      '• Copy continuing students from ' + previousYear + ' rosters\n' +
      '• Create "' + mostRecentSemester + '" roster sheets\n' +
      '• Update Teacher Roster Lookup URLs\n\n' +
      'Continuing students: Lessons Remaining > 0 AND Status = active/carryover\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirm !== ui.Button.YES) {
      return;
    }
    
    // Step 5: Get roster folders
    var rostersFolder = UtilityScriptLibrary.getRosterFolder();
    
    var previousYearFolderName = previousYear + ' Rosters';
    var newYearFolderName = newYear + ' Rosters';
    
    var previousYearFolder = null;
    var newYearFolder = null;
    
    var subfolders = rostersFolder.getFolders();
    while (subfolders.hasNext()) {
      var folder = subfolders.next();
      var folderName = folder.getName();
      
      if (folderName === previousYearFolderName) {
        previousYearFolder = folder;
      }
      if (folderName === newYearFolderName) {
        newYearFolder = folder;
      }
    }
    
    if (!previousYearFolder) {
      throw new Error('Previous year folder not found: ' + previousYearFolderName);
    }
    
    if (!newYearFolder) {
      throw new Error('New year folder not found: ' + newYearFolderName + '\n\nPlease run semester setup in Billing first.');
    }
    
    UtilityScriptLibrary.debugLog('Found folders - Previous: ' + previousYearFolderName + ', New: ' + newYearFolderName);
    
    // Step 6: Get active teachers
    var teacherLookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    if (!teacherLookupSheet) {
      throw new Error('Teacher Roster Lookup sheet not found');
    }
    
    var teacherData = teacherLookupSheet.getDataRange().getValues();
    var getCol = UtilityScriptLibrary.createColumnFinder(teacherLookupSheet);
    
    var teacherNameCol = getCol('Teacher Name') - 1;
    var displayNameCol = getCol('Display Name') - 1;
    var statusCol = getCol('Status') - 1;
    var urlCol = getCol('Roster URL') - 1;
    
    if (teacherNameCol === -1 || displayNameCol === -1 || statusCol === -1 || urlCol === -1) {
      throw new Error('Required columns not found in Teacher Roster Lookup');
    }
    
    // Step 7: Process each active teacher
    var stats = {
      processed: 0,
      created: 0,
      skipped: 0,
      errors: []
    };
    
    for (var i = 1; i < teacherData.length; i++) {
      var row = teacherData[i];
      var teacherName = row[teacherNameCol];
      var displayName = row[displayNameCol];
      var status = row[statusCol];
      
      if (!teacherName || status !== 'active') {
        if (teacherName) {
          UtilityScriptLibrary.debugLog('Skipping ' + teacherName + ' (not active)');
          stats.skipped++;
        }
        continue;
      }
      
      try {
        UtilityScriptLibrary.debugLog('Processing teacher: ' + teacherName);
        
        // Find previous year workbook
        var previousYearFileName = 'QAMP ' + previousYear + ' ' + (displayName || teacherName);
        var previousYearFiles = previousYearFolder.getFilesByName(previousYearFileName);
        
        if (!previousYearFiles.hasNext()) {
          UtilityScriptLibrary.debugLog('No ' + previousYear + ' workbook found for ' + teacherName + ' - skipping');
          stats.skipped++;
          continue;
        }
        
        var previousYearFile = previousYearFiles.next();
        var previousYearSS = SpreadsheetApp.openById(previousYearFile.getId());
        
        // Get continuing students
        var continuingStudents = getContinuingStudentsFromWorkbook(previousYearSS);
        
        UtilityScriptLibrary.debugLog('Found ' + continuingStudents.length + ' continuing students for ' + teacherName);
        
        // Create new year workbook
        var newYearSS = getOrCreateRosterFromTemplate(
          displayName || teacherName,
          newYearFolder,
          newYear,
          mostRecentSemester
        );
        
        // Populate with continuing students if any
        if (continuingStudents.length > 0) {
          populateRosterWithContinuingStudents(newYearSS, mostRecentSemester, continuingStudents);
        }
        
        // Update Teacher Roster Lookup URL
        var newUrl = newYearSS.getUrl();
        teacherLookupSheet.getRange(i + 1, urlCol + 1).setValue(newUrl);
        
        stats.processed++;
        stats.created++;
        
        UtilityScriptLibrary.debugLog('✅ Created ' + newYear + ' workbook for ' + teacherName + ' with ' + continuingStudents.length + ' students');
        
      } catch (error) {
        stats.errors.push({
          teacher: teacherName,
          error: error.message
        });
        UtilityScriptLibrary.debugLog('❌ Error processing ' + teacherName + ': ' + error.message);
      }
    }
    
    // Step 8: Show summary
    var summary = 'Workbook Creation Complete\n\n' +
                  'Semester: ' + mostRecentSemester + '\n' +
                  'Teachers Processed: ' + stats.processed + '\n' +
                  'Workbooks Created: ' + stats.created + '\n' +
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
    UtilityScriptLibrary.debugLog('❌ Fatal error: ' + error.message);
    ui.alert('❌ Error', error.message, ui.ButtonSet.OK);
  }
}

function getContinuingStudentsFromWorkbook(workbook) {
  try {
    // Find the most recent roster sheet
    var rosterSheet = findMostRecentRosterSheet(workbook);
    
    if (!rosterSheet) {
      UtilityScriptLibrary.debugLog('No roster sheet found in workbook');
      return [];
    }
    
    UtilityScriptLibrary.debugLog('Using roster sheet: ' + rosterSheet.getName());
    
    var data = rosterSheet.getDataRange().getValues();
    var headers = data[0];
    
    var getCol = function(name) {
      for (var i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().toLowerCase().indexOf(name.toLowerCase()) !== -1) {
          return i;
        }
      }
      return -1;
    };
    
    var continuingStudents = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Skip empty rows
      if (!row[getCol('Last Name')] || !row[getCol('First Name')]) {
        continue;
      }
      
      var status = (row[getCol('Status')] || '').toString().trim().toLowerCase();
      var lessonsRemaining = parseFloat(row[getCol('Lessons Remaining')]) || 0;
      
      // Include students with: Lessons Remaining > 0 AND (Status = 'active' OR 'carryover')
      if (lessonsRemaining > 0 && (status === 'active' || status === 'carryover')) {
        continuingStudents.push({
          lastName: row[getCol('Last Name')] || '',
          firstName: row[getCol('First Name')] || '',
          instrument: row[getCol('Instrument')] || '',
          length: row[getCol('Length')] || 30,
          experience: row[getCol('Experience')] || '',
          grade: row[getCol('Grade')] || '',
          school: row[getCol('School')] || '',
          schoolTeacher: row[getCol('School Teacher')] || '',
          parentLastName: row[getCol('Parent Last Name')] || '',
          parentFirstName: row[getCol('Parent First Name')] || '',
          phone: row[getCol('Phone')] || '',
          email: row[getCol('Email')] || '',
          additionalContacts: row[getCol('Additional contacts')] || '',
          studentId: row[getCol('Student ID')] || '',
          lessonsRemaining: lessonsRemaining
        });
      }
    }
    
    return continuingStudents;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('Error getting continuing students: ' + error.message);
    return [];
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
    
    // Sort students alphabetically
    students.sort(function(a, b) {
      var lastNameCompare = (a.lastName || '').localeCompare(b.lastName || '');
      if (lastNameCompare !== 0) {
        return lastNameCompare;
      }
      return (a.firstName || '').localeCompare(b.firstName || '');
    });
    
    // Prepare data rows
    var dataRows = [];
    for (var i = 0; i < students.length; i++) {
      var s = students[i];
      dataRows.push([
        false,                      // A - Contacted checkbox
        '',                         // B - First Lesson Date
        '',                         // C - First Lesson Time
        '',                         // D - Comments
        s.lastName,                 // E - Last Name
        s.firstName,                // F - First Name
        s.instrument,               // G - Instrument
        s.length,                   // H - Length
        s.experience,               // I - Experience
        s.grade,                    // J - Grade
        s.school,                   // K - School
        s.schoolTeacher,            // L - School Teacher
        s.parentLastName,           // M - Parent Last Name
        s.parentFirstName,          // N - Parent First Name
        s.phone,                    // O - Phone
        s.email,                    // P - Email
        s.additionalContacts,       // Q - Additional contacts
        s.lessonsRemaining,         // R - Lessons Registered (carry forward remaining)
        0,                          // S - Lessons Completed
        s.lessonsRemaining,         // T - Lessons Remaining
        'Carryover',                // U - Status
        s.studentId,                // V - Student ID
        'Carried over from previous year', // W - Admin Comments
        ''                          // X - System Comments
      ]);
    }
    
    // Write all data at once
    if (dataRows.length > 0) {
      rosterSheet.getRange(2, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
      UtilityScriptLibrary.debugLog('✅ Populated roster with ' + students.length + ' students');
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('Error populating roster: ' + error.message);
    throw error;
  }
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
