================================================================================
UTILITY LIBRARY FUNCTION DIRECTORY
================================================================================
    Total Functions: 128 (126 standard functions + 2 EnvironmentManager methods)
     Most Recent version: 88

    This directory provides a quick reference for all functions in Utility script.
    Parameters marked with ? are optional. 'options' parameters are optional 
    configuration objects with properties documented inline in the function code.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Dependencies: function1(), function2()
            Called by: 
  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    addToCurrencyCols
    appendToMetadataWithVerification
    buildAcademicYearVariables
    buildRateMapFromSheet
    buildRateVariables
    calculateGraduationYear
    bulkUpdateStudentStatus
    cascadeFormerStatus
    cleanName
    clearCache
    clearEmptyRows
    clearOldDebugEntries
    clearUtilityDebugLog
    columnToLetter
    combineDocumentsIntoPDF
    convertYesNoToBoolean
    copyPreviousColumnToNew
    copySheetWithProtections
    copyTextAttributes
    createColumnFinder
    createGroupSections
    createLessonRows
    createMonthlyAttendanceSheet
    createStudentHeader
    createStudentSections
    createUtilityDebugSheet
    debugLog
    deleteExtraColumns
    determineIfStudentIsAdult
    documentAlreadyExists
    enableDatePickerForColumn
    executeWithErrorHandling
    extractLessonQuantityFromPackage
    extractNumericLessonLength
    extractRosterData
    extractSeasonFromSemester
    extractTeacherNameFromWorkbook
    extractTotalLessonsFromPackages
    findColumnByPartialName
    findTeacherInRosterLookup
    findParentRow
    findStudentInContacts
    findStudentRow
    formatAddress
    formatAttendanceColumns
    formatAttendanceSheet
    formatCurrency
    formatDateFlexible
    formatLessonLengthWithMinutes
    formatLogValue
    formatPhoneNumber
    formatRosterColumns
    freezeSheetFormulas
    generateDocumentFromTemplate
    generateKey
    generateNextId
    generateNextIdDirect
    getAttendanceSheetForDate
    getCached
    getColumnHeaders
    getColumnIndices
    getCurrentAcademicYearInfo
    getCurrentSemesterMonth
    getCurrentSemesterName
    getDateForWeekday
    getFieldMappingFromSheet
    getGeneratedDocumentsFolder
    getHeaderMap
    getInstrumentFamily
    getLessonLengthFromPackages
    getMonthNameFromDate
    getMonthNames
    getMonthSheets
    getMostRecentRateColumn
    getRateSummary
    getRosterFolder
    getRosterFolderUrlForYear
    getSemesterDates
    getSemesterForDate
    getSheet
    getStudentIdFromRow
    getTeacherGroupAssignments
    getTeacherIdByDisplayName
    getTeacherNameById
    getTemplate
    getTemplateFolder
    getWeekdayName
    getWeekdayNumber
    getWorkbook
    getYearFromSemesterName
    inferWorkbookKey
    insertCountFormula
    interpretAgeField
    isCurrentOrFutureMonth
    isHistoricalDataInputEnabled
    isIdAlreadyUsed
    isMonthSheet
    logAllSheetHeaders
    normalizeHeader
    parseAllPackageQuantities
    parseGridInstruments
    parseAndFormatAddress
    parseCityZipMessy
    parseDateFromString
    parseRosterData
    prefillAttendanceDatesForStudent
    promptForCustomToday
    promptForDate
    promptForHistoricalId
    promptForNameWithDefault
    protectSheetRanges
    safeGet
    safeParseFloat
    setCached
    setupAttendanceHeaders
    setupRosterTemplateProtection
    setupStatusValidation
    shouldBeCurrency
    showConfirmationDialog
    styleHeaderRow
    truncateString
    updateFieldMappings
    updateParentContactFields
    validateProgramConfiguration
    validateTemplateVariables
    verifyConfigurationWithUser
  ================================================================================
  ================================================================================
  ENVIRONMENT MANAGER (Special Module)
  ================================================================================
    EnvironmentManager.set(env) -> void
        Sets the current environment to 'test' or 'prod'. Throws error if invalid.
        Use: EnvironmentManager.set('prod')
        Category: CONFIGURATION
        Dependencies: None
        Called by: 

    EnvironmentManager.get() -> String
        Returns the current environment setting ('test' or 'prod')
        Use: var env = EnvironmentManager.get()
        Category: CONFIGURATION
        Dependencies: None
        Called by: (internal: getGeneratedDocumentsFolder, getRosterFolder, getSheet, getTemplateFolder, getWorkbook, prefillAttendanceDatesForStudent)

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addToCurrencyCols(currencyCols, columnNumber, headerName) -> void
        Adds a column number to the currencyCols array if the column should be 
        currency formatted. Uses shouldBeCurrency() to validate. Critical for 
        preventing hours columns from being formatted as currency.
        Category: UTILITIES
        Dependencies: debugLog(), shouldBeCurrency()
        Called by: Billing

    appendToMetadataWithVerification(metadataSheet, rowData, verificationSteps, options?) -> Object
        Appends a row to metadata sheet with multi-step user verification process.
        Executes verification functions and shows confirmation dialogs before appending.
        Returns: {success, message, verificationResults, rowData, error?}
        Category: METADATA
        Dependencies: showConfirmationDialog()
        Called by: 

    buildAcademicYearVariables() -> Object
        Builds template variables for the current academic year from Semester Metadata.
        Extracts year from latest semester (e.g., "Fall 2024" -> "2024-2025").
        Returns: {CurrentAcademicYear, AcademicYearStart, AcademicYearEnd}
        Category: VARIABLE_BUILDING
        Dependencies: None
        Called by: (internal: getRosterFolder)

    buildRateMapFromSheet(sheet, headers, yearColIndex) -> Object
        Internal helper that builds a key-value map of rates from a sheet column.
        Maps first column (titles) to specified year column values.
        Category: VARIABLE_BUILDING
        Dependencies: None
        Called by: Billing, TeacherInvoice (internal: buildRateVariables)

    buildRateVariables() -> Object
        Builds template variables for lesson rates from the Rates sheet.
        Always uses column B (index 1) for rate values.
        Returns: {HourlyRate, HalfHourRate, LateFee, LateFeeGracePeriod}
        Category: VARIABLE_BUILDING
        Dependencies: buildRateMapFromSheet(), formatCurrency()
        Called by: Billing (internal: getRateSummary)


    calculateGraduationYear(grade) -> String|Number
        Calculates a student's expected graduation year from their current grade.
        Returns "Adult" for adult/college students, adds 13 years for Kindergarten,
        or computes (currentYear + 12 - gradeNum). Returns empty string if grade is falsy.
        Category: DATA_MANIPULATION
        Dependencies: debugLog()
        Called by: Billing, Responses

    bulkUpdateStudentStatus(studentsSheet, statusColumn, newValue, options?) -> Object
        Updates a status column for multiple students based on conditions.
        Options: {condition, whereColumn, whereValue, skipEmpty}
        Returns: {success, updatedCount, changedRows, error?}
        Category: BULK_OPERATIONS
        Dependencies: normalizeHeader()
        Called by: Billing

    cascadeFormerStatus(teacherId) -> void
        Marks all instrument rows for the given teacher ID as "Former" in the Instrument
        List sheet, and updates the teacher's status in the Teacher Roster Lookup sheet.
        Category: SHEET_OPERATIONS
        Dependencies: getSheet(), getHeaderMap(), normalizeHeader(), debugLog()
        Called by: Contacts, TeacherResponses

    cleanName(name) -> String
        Cleans and standardizes a name by removing extra whitespace and 
        standardizing capitalization.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Billing

    clearCache() -> void
        Clears the execution cache (_executionCache object).
        Category: CACHE
        Dependencies: None
        Called by: 

    clearEmptyRows(sheet) -> void
        Deletes empty rows from the bottom of a sheet to clean up formatting.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

    clearOldDebugEntries(debugSheet) -> void
        Removes old entries from debug sheet to prevent performance issues.
        Called automatically when debug sheet append fails.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: debugLog)

    clearUtilityDebugLog() -> void
        Completely clears the Utility Debug sheet by deleting all rows except header.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

    columnToLetter(column) -> String
        Converts a column number to its letter representation (e.g., 1 -> 'A', 27 -> 'AA').
        Category: UTILITIES
        Dependencies: None
        Called by: Billing, TeacherInvoice

    combineDocumentsIntoPDF(documents, fileName, destinationFolder) -> Object
        Combines multiple Google Docs into a single PDF file in specified folder.
        Returns: {success, pdfFile, pdfUrl, error?}
        Category: DOCUMENT_GENERATION
        Dependencies: None
        Called by: Billing

    convertYesNoToBoolean(value) -> Boolean
        Converts "Yes"/"No" strings to boolean values. Returns false for anything else.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Billing

    copyPreviousColumnToNew(newRow, prevRow, currMap, prevMap, mapping) -> Boolean
        Copies a value from a previous row into a new row using header map lookups.
        Returns true if copy succeeded, false if columns not found.
        Category: DATA_MANIPULATION
        Dependencies: normalizeHeader()
        Called by: Billing

    copySheetWithProtections(sourceSheet, targetWorkbook, newName, options?) -> Object
        Copies a sheet to another workbook while preserving protections and formatting.
        Options: {clearContents, preserveValidation, copyProtections}
        Returns: {success, sheet, sheetName, error?}
        Category: SHEET_OPERATIONS
        Dependencies: debugLog()
        Called by: Billing

    copyTextAttributes(sourcePara, targetPara) -> void
        Copies text formatting attributes from source paragraph to target paragraph
        in Google Docs.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: generateDocumentFromTemplate)

    createColumnFinder(sheet) -> Function
        Returns a closure function that finds column indices by header name.
        Usage: var findCol = createColumnFinder(sheet); var nameCol = findCol('Name');
        Category: SHEET_OPERATIONS
        Dependencies: normalizeHeader()
        Called by: Billing, Responses, TeacherResponses (internal: findTeacherInRosterLookup, getTeacherGroupAssignments, getTeacherIdByDisplayName)

    createGroupSections(sheet, groupEntries) -> void
        Creates formatted group sections in a roster sheet with headers and student rows.
        Used for group lesson rosters.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: Responses (internal: createMonthlyAttendanceSheet)

    createLessonRows(sheet, student, startRow, numRows) -> Number
        Creates formatted lesson tracking rows for a student in a roster sheet.
        Returns the next available row number.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: createStudentSections)

    createMonthlyAttendanceSheet(workbook, monthName, rosterData) -> Sheet
        Creates a new monthly attendance sheet with student sections and formatting.
        Returns the newly created sheet.
        Category: SHEET_OPERATIONS
        Dependencies: createGroupSections(), createStudentSections(), debugLog(), extractTeacherNameFromWorkbook(), formatAttendanceSheet(), setupAttendanceHeaders(), styleHeaderRow()
        Called by: Billing, Responses

    createStudentHeader(sheet, student, row) -> Number
        Creates a formatted header row for a student in a roster sheet.
        Returns the next available row number.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: createStudentSections)

    createStudentSections(sheet, rosterData) -> void
        Creates all student sections in a roster sheet based on roster data array.
        Category: SHEET_OPERATIONS
        Dependencies: createLessonRows(), createStudentHeader(), debugLog()
        Called by: Billing, Responses (internal: createMonthlyAttendanceSheet)

    createUtilityDebugSheet() -> Sheet
        Creates and formats the "Debug" sheet for logging with standard headers.
        Returns the debug sheet or null if creation fails.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: debugLog)

    debugLog(functionName, eventType, message, data?, errorDetails?) -> void
        Logs debug information to the Debug sheet with timestamp and formatting.
        EventType: "ERROR", "INFO", "DEBUG", "WARNING", etc.
        Handles backward compatibility for single-parameter legacy calls.
        Category: DEBUG
        Dependencies: clearOldDebugEntries(), createUtilityDebugSheet(), formatLogValue(), truncateString()
        Called by: Billing, Responses, Contacts, TeacherResponses, TeacherInvoice, Payments (internal: addToCurrencyCols, calculateGraduationYear, cascadeFormerStatus, copySheetWithProtections, createMonthlyAttendanceSheet, createStudentSections, executeWithErrorHandling, findTeacherInRosterLookup, generateDocumentFromTemplate, getAttendanceSheetForDate, getCurrentSemesterName, getSemesterForDate, getTeacherGroupAssignments, getTeacherIdByDisplayName, getTeacherNameById, parseRosterData, promptForDate, updateParentContactFields)

    deleteExtraColumns(sheet, keepColumns) -> void
        Deletes all columns after the specified keepColumns count to clean up sheets.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

    determineIfStudentIsAdult(studentData) -> Boolean
        Determines if a student should be treated as an adult based on age field.
        Returns true if age is 18+, "Adult", or "Parent".
        Category: VALIDATION
        Dependencies: None
        Called by: Billing

    documentAlreadyExists(fileName, folder) -> File | Boolean
        Checks if a document with the given name exists in folder.
        Returns the File object if found, false otherwise.
        Category: UTILITIES
        Dependencies: None
        Called by: Billing, TeacherInvoice

    enableDatePickerForColumn(sheet, columnIndex, startRow) -> void
        Enables date picker data validation for a column starting at specified row.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

    executeWithErrorHandling(operation, successMessage, context?, options?) -> Object
        Wraps an operation with standardized error handling and debug logging.
        Options: {showSuccessAlert, returnData, suppressErrors}
        Returns: {success, data?, error?}
        Category: ERROR_HANDLING
        Dependencies: debugLog()
        Called by: Billing

    extractLessonQuantityFromPackage(packageText) -> Number
        Extracts the numeric quantity from package text (e.g., "5x 30min" -> 5).
        Returns 0 if no quantity found.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: Billing, Responses (internal: extractTotalLessonsFromPackages, parseAllPackageQuantities, parseRosterData)

    extractNumericLessonLength(lessonLengthValue) -> Number
        Extracts numeric lesson length in minutes from various formats.
        Handles "30", "30 min", "30min", "0.5 hour" formats.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: Billing, Responses

    extractRosterData(rosterSheet) -> Array
        Extracts all roster data from a teacher's roster sheet.
        Returns array of student objects with lesson details.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: Responses

    extractSeasonFromSemester(semesterName) -> String
        Extracts season from semester name (e.g., "Fall 2024" -> "Fall").
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: Billing, Responses (internal: getSemesterDates)

    extractTeacherNameFromWorkbook(workbook) -> String
        Extracts teacher name from workbook title by removing " Roster" suffix.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: (internal: createMonthlyAttendanceSheet)

    extractTotalLessonsFromPackages(qty30, qty45, qty60) -> Number
        Sums lesson quantities from all three package fields (30min, 45min, 60min).
        Handles empty strings and null values gracefully.
        Returns total number of lessons across all packages.
        Category: DATA_EXTRACTION
        Dependencies: extractLessonQuantityFromPackage()
        Called by: 
        
    findColumnByPartialName(headers, searchTerm) -> Number
        Finds column index by partial header name match (case-insensitive).
        Returns -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: None
        Called by: (internal: updateFieldMappings)

    findParentRow(parentsSheet, parentId, fallbackKey?) -> Number
        Finds the row number for a parent by Parent ID or fallback lookup key.
        Returns -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: normalizeHeader()
        Called by: Billing, Responses

    findStudentInContacts(contactsData, studentIdCol, targetStudentId) -> Number
        Searches a contacts data array for a student by ID.
        Returns 0-based row index, or -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: None
        Called by: Billing

    findStudentRow(studentSheet, studentKey) -> Number
        Finds the row number for a student using various lookup strategies
        (Student ID, first/last name combinations). Returns -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: normalizeHeader()
        Called by: Responses

    findTeacherInRosterLookup(lookupSheet, teacherId) -> Number
        Searches the Teacher Roster Lookup sheet for a teacher by Teacher ID.
        Returns the 1-based row number, or -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: createColumnFinder(), debugLog()
        Called by: Responses, TeacherResponses

    formatAddress(street, city, zip) -> String
        Formats address components into a single line address string.
        Handles missing components gracefully.
        Category: FORMATTING
        Dependencies: None
        Called by: Billing, Responses, TeacherResponses (internal: parseAndFormatAddress, parseRosterData)

    formatAttendanceColumns(sheet, studentCount) -> void
        Formats date columns in attendance sheet with proper widths and styling.
        Category: FORMATTING
        Dependencies: None
        Called by: 

    formatAttendanceSheet(sheet) -> void
        Applies comprehensive formatting to an attendance sheet including headers,
        protections, and column widths.
        Category: FORMATTING
        Dependencies: None
        Called by: (internal: createMonthlyAttendanceSheet)

    formatCurrency(amount) -> String
        Formats a number as USD currency string (e.g., 45.5 -> "$45.50").
        Category: FORMATTING
        Dependencies: None
        Called by: Billing, TeacherInvoice (internal: buildRateVariables)

    formatDateFlexible(date, format?) -> String
        Formats a date object with flexible format string support.
        Default format: "MM/dd/yyyy"
        Category: FORMATTING
        Dependencies: None
        Called by: Billing, TeacherInvoice (internal: promptForDate)

    formatLessonLengthWithMinutes(lessonLengthValue) -> String
        Formats lesson length to include "min" suffix (e.g., "30" -> "30 min").
        Category: FORMATTING
        Dependencies: None
        Called by: 

    formatLogValue(value) -> String
        Formats various data types for logging (handles objects, arrays, nulls).
        Used internally by debugLog().
        Category: FORMATTING
        Dependencies: truncateString()
        Called by: (internal: debugLog)

    formatPhoneNumber(phoneRaw) -> String
        Formats phone number to (XXX) XXX-XXXX format. Returns original if invalid.
        Category: FORMATTING
        Dependencies: None
        Called by: Responses, TeacherResponses (internal: parseRosterData)

    formatRosterColumns(sheet) -> void
        Applies standard column width formatting to roster sheets.
        Category: FORMATTING
        Dependencies: None
        Called by: 

    freezeSheetFormulas() -> void
        Converts all formulas in active spreadsheet to static values.
        Used for archiving or finalizing data.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

    generateDocumentFromTemplate(templateKey, variableData, fileName, destinationFolder) -> Object
        Generates a Google Doc from template with variable substitution.
        Returns: {success, document, documentId, documentUrl, error?}
        Category: DOCUMENT_GENERATION
        Dependencies: copyTextAttributes(), debugLog(), getTemplate()
        Called by: Billing, TeacherInvoice

    generateKey() -> String
        Generates a random alphanumeric key (8 characters) for lookups.
        Category: ID_GENERATION
        Dependencies: None
        Called by: Billing, Responses, TeacherResponses

    generateNextId(sheet, columnName, prefix, recordName?) -> String
        Generates the next sequential ID in format PREFIX#### (e.g., "S0042").
        Scans column for highest ID and increments.
        Category: ID_GENERATION
        Dependencies: isHistoricalDataInputEnabled(), normalizeHeader(), promptForHistoricalId()
        Called by: Responses, TeacherResponses

    generateNextIdDirect(sheet, columnName, prefix) -> String
        Generates next sequential ID without checking historical mode.
        Internal helper function that performs direct auto-generation.
        Category: ID_GENERATION
        Local functions used: normalizeHeader()
        Utility functions used: None

    getAttendanceSheetForDate(ss, targetDate) -> Sheet | null
        Returns the attendance sheet for a given date's month or null if not found.
        Category: ATTENDANCE
        Dependencies: debugLog(), getMonthNameFromDate(), isMonthSheet()
        Called by: 

    getCached(key) -> Any
        Retrieves value from execution cache. Returns undefined if not found.
        Category: CACHE
        Dependencies: None
        Called by: (internal: getColumnIndices, getMonthSheets, isMonthSheet)

    getColumnHeaders(sheet) -> Array
        Returns array of header values from first row of sheet.
        Category: DATA_RETRIEVAL
        Dependencies: None
        Called by: TeacherResponses

    getColumnIndices(sheet, columnNames) -> Object
        Returns object mapping column names to their indices.
        Example: {Name: 0, Email: 1, Phone: 2}
        Category: DATA_RETRIEVAL
        Dependencies: getCached(), normalizeHeader(), setCached()
        Called by: 

    getCurrentAcademicYearInfo() -> Object
        Gets current academic year info from most recent semester metadata.
        Returns: {year, startYear, endYear, semester, dates: {start, end}}
        Category: DATE_TIME
        Dependencies: getSemesterDates(), getYearFromSemesterName()
        Called by: 

    getCurrentSemesterMonth(semesterName) -> String
        Returns the starting month name for a semester (e.g., "Fall 2024" -> "September").
        Category: DATE_TIME
        Dependencies: None
        Called by: Responses

    getCurrentSemesterName() -> String|null
        Returns the current semester name from the last row of Semester Metadata.
        Returns null if sheet not found or empty.
        Category: CONFIGURATION
        Dependencies: debugLog(), getSheet()
        Called by: Billing, Responses, TeacherInvoice

    getDateForWeekday(weekStartDate, weekdayName) -> Date
        Calculates date for a specific weekday within a week starting at weekStartDate.
        Category: DATE_TIME
        Dependencies: getWeekdayNumber()
        Called by: (internal: getWeekdayName, prefillAttendanceDatesForStudent)

    getFieldMappingFromSheet(fieldMapSheet) -> Object
        Reads field mapping configuration from FieldMap sheet.
        Returns object mapping form fields to internal field names.
        Category: FIELD_MAPPING
        Dependencies: None
        Called by: Billing, Responses, TeacherResponses

    getGeneratedDocumentsFolder() -> Folder
        Returns the Google Drive folder for generated documents based on environment.
        Category: CONFIGURATION
        Dependencies: EnvironmentManager.get()
        Called by: Billing, TeacherInvoice

    getHeaderMap(sheet) -> Object
        Creates object mapping normalized header names to column indices.
        Used for flexible column lookups.
        Category: DATA_RETRIEVAL
        Dependencies: normalizeHeader()
        Called by: Billing, Responses, Contacts, TeacherResponses, TeacherInvoice, Payments (internal: cascadeFormerStatus, getTeacherNameById)

    getInstrumentFamily(instrument) -> String
        Looks up an instrument in the INSTRUMENT_FAMILIES constant map and returns its
        family name (e.g., "violin" -> "Strings"). Returns empty string if not found.
        Category: DATA_RETRIEVAL
        Dependencies: None
        Called by: TeacherResponses

    getLessonLengthFromPackages(qty30Package, qty45Package, qty60Package) -> Number
        Determines lesson length (30, 45, or 60) from package quantities.
        Returns 30 if no packages have quantity.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Billing

    getMonthNameFromDate(date, capitalize?) -> String
        Returns month name from date object. Optionally capitalizes first letter.
        Category: DATE_TIME
        Dependencies: None
        Called by: (internal: getAttendanceSheetForDate)

    getMonthNames() -> Array
        Returns array of all month names: ["January", "February", ...]
        Category: DATE_TIME
        Dependencies: None
        Called by: Billing, Responses, TeacherInvoice (internal: isCurrentOrFutureMonth, isMonthSheet)

    getMonthSheets(ss) -> Array
        Returns array of all month-based attendance sheets from spreadsheet.
        Category: DATA_RETRIEVAL
        Dependencies: getCached(), getSheet(), isMonthSheet(), setCached()
        Called by: Billing

    getMostRecentRateColumn(headers) -> Number
        Finds the most recent year column in Rates sheet headers.
        Currently returns 1 (column B) as fixed implementation.
        Category: RATES
        Dependencies: None
        Called by: Billing

    getRateSummary() -> String
        Returns formatted summary of current rates for user display/verification.
        Category: CONFIGURATION
        Dependencies: buildRateVariables()
        Called by: Billing

    getRosterFolder() -> Folder
        Returns the Google Drive folder containing roster files for current academic year.
        Category: CONFIGURATION
        Dependencies: buildAcademicYearVariables(), getConfig(), getRosterFolderUrlForYear()
        Called by: Billing, Responses

    getRosterFolderUrlForYear(year) -> String
        Returns the Drive URL for the roster folder of a specific year.
        Category: CONFIGURATION
        Dependencies: EnvironmentManager.get()
        Called by: (internal: getRosterFolder)

    getSemesterDates(semesterName) -> Object
        Parses semester name and returns start/end dates based on semester type.
        Returns: {start: Date, end: Date}
        Category: DATE_TIME
        Dependencies: extractSeasonFromSemester(), getYearFromSemesterName()
        Called by: (internal: getCurrentAcademicYearInfo)

    getSemesterForDate(targetDate) -> String|null
        Finds and returns the semester name whose date range contains the given date.
        Returns null if no matching semester found.
        Category: DATE_TIME
        Dependencies: debugLog(), getSheet(), normalizeHeader()
        Called by: Billing

    getSheet(sheetKey) -> Sheet
        Returns a sheet object by its SHEET_MAP key (e.g., 'students', 'parents').
        Automatically infers which workbook to use.
        Category: CONFIGURATION
        Dependencies: getWorkbook(), inferWorkbookKey()
        Called by: Billing, Responses, Contacts, TeacherResponses, TeacherInvoice (internal: cascadeFormerStatus, getCurrentSemesterName, getMonthSheets, getSemesterForDate, getTeacherGroupAssignments, getTeacherIdByDisplayName, getTeacherNameById)

    getStudentIdFromRow(row, headerMap) -> String|null
        Extracts student ID from a data row using flexible column name matching.
        Checks for both "Student ID" and "ID" column headers.
        Returns null if student ID not found.
        Category: DATA_EXTRACTION
        Dependencies: normalizeHeader()
        Called by: Billing

    getTeacherGroupAssignments(teacherName) -> Array
        Returns array of group lesson assignments for a teacher from lookup sheet.
        Category: ROSTER
        Dependencies: createColumnFinder(), debugLog(), getSheet(), getWorkbook()
        Called by: Responses

    getTeacherIdByDisplayName(displayName) -> String
        Looks up a Teacher ID from the Teacher Roster Lookup sheet by display name.
        If the input already matches /^T\d+$/, returns it as-is.
        Returns empty string if not found or on error.
        Category: DATA_RETRIEVAL
        Dependencies: getSheet(), createColumnFinder(), debugLog()
        Called by: Billing, Responses, Contacts

    getTeacherNameById(teacherId) -> String
        Looks up a teacher's full name ("First Last") from the Teachers and Admin sheet
        by Teacher ID. Returns empty string if not found or on error.
        Category: DATA_RETRIEVAL
        Dependencies: getSheet(), getHeaderMap(), normalizeHeader(), debugLog()
        Called by: Billing

    getTemplate(templateKey) -> File
        Returns Google Doc template file by TEMPLATE_MAP key.
        Category: CONFIGURATION
        Dependencies: getTemplateFolder()
        Called by: (internal: generateDocumentFromTemplate)

    getTemplateFolder() -> Folder
        Returns the Google Drive folder containing document templates.
        Category: CONFIGURATION
        Dependencies: EnvironmentManager.get()
        Called by: (internal: getTemplate)

    getWeekdayName(startDate, weekdayName) -> String
        Returns formatted date string for a weekday (e.g., "Monday 9/15").
        Category: DATE_TIME
        Dependencies: getDateForWeekday()
        Called by: 

    getWeekdayNumber(weekdayName) -> Number
        Converts weekday name to number (0=Sunday, 6=Saturday).
        Category: DATE_TIME
        Dependencies: None
        Called by: (internal: getDateForWeekday)

    getWorkbook(workbookKey) -> Spreadsheet
        Returns a spreadsheet by CONFIG key (e.g., 'paymentsId', 'contactsId').
        Category: CONFIGURATION
        Dependencies: EnvironmentManager.get()
        Called by: Billing, TeacherInvoice (internal: getSheet, getTeacherGroupAssignments)

    getYearFromSemesterName(semesterName) -> String
        Extracts 4-digit year from semester name (e.g., "Fall 2024" -> "2024").
        Category: DATE_TIME
        Dependencies: None
        Called by: Billing, Responses (internal: getCurrentAcademicYearInfo, getSemesterDates)

    inferWorkbookKey(sheetKey) -> String
        Determines which CONFIG workbook key to use based on SHEET_MAP sheet key.
        Example: 'students' -> 'contactsId'
        Category: UTILITIES
        Dependencies: None
        Called by: (internal: getSheet)

    insertCountFormula(sheet, row, startCol, endCol, targetCol) -> void
        Inserts a COUNTA formula in target column to count filled cells in a range.
        Category: UTILITIES
        Dependencies: None
        Called by: 

    interpretAgeField(ageResponse) -> String
        Interprets age field from forms into standardized format ("Minor", "Adult", or age).
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: 

    isCurrentOrFutureMonth(sheetName, targetDate) -> Boolean
        Returns true if sheet represents current or future month relative to targetDate.
        Category: VALIDATION
        Dependencies: getMonthNames(), parseDateFromString()
        Called by: Billing

    isHistoricalDataInputEnabled() -> Boolean
        Returns value of HISTORICAL_DATA_MODE flag.
        Category: VALIDATION
        Dependencies: None
        Called by: (internal: generateNextId, promptForCustomToday)

    isIdAlreadyUsed(sheet, columnName, idToCheck) -> Boolean
        Checks if an ID already exists in a column. Returns true if found.
        Category: VALIDATION
        Dependencies: normalizeHeader()
        Called by: (internal: promptForHistoricalId)

    isMonthSheet(sheetName) -> Boolean
        Returns true if sheet name exactly matches a month name.
        Category: VALIDATION
        Dependencies: getCached(), getMonthNames(), setCached()
        Called by: Billing (internal: getAttendanceSheetForDate, getMonthSheets)

    logAllSheetHeaders() -> void
        Logs all sheet names and their header rows to the Apps Script logger.
        Used for diagnostics.
        Category: DEBUG
        Dependencies: None
        Called by: Responses

    normalizeHeader(header) -> String
        Normalizes header string for comparison (lowercase, no spaces/punctuation).
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Billing, Responses, Contacts, TeacherResponses, TeacherInvoice (internal: bulkUpdateStudentStatus, cascadeFormerStatus, copyPreviousColumnToNew, createColumnFinder, findParentRow, findStudentRow, generateNextId, getColumnIndices, getHeaderMap, getSemesterForDate, getStudentIdFromRow, getTeacherNameById, isIdAlreadyUsed, parseRosterData, updateFieldMappings, updateParentContactFields)

    parseAllPackageQuantities(qty30Package, qty45Package, qty60Package) -> Object
        Parses all package quantity strings and returns structured object.
        Returns: {qty30, qty45, qty60, totalLessons}
        Category: DATA_MANIPULATION
        Dependencies: extractLessonQuantityFromPackage()
        Called by: Billing

    parseGridInstruments(headers, rowValues) -> Array
        Parses grid-style instrument checkboxes from a teacher interest form submission.
        Identifies columns matching the grid question prefix, extracts instrument name
        from bracket notation, and maps level responses to beg/int/adv abbreviations.
        Returns array of {instrument, levels} objects.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: TeacherResponses

    parseAndFormatAddress(rawAddress) -> String
        Parses messy address input and returns formatted single-line address.
        Category: DATA_MANIPULATION
        Dependencies: formatAddress(), parseCityZipMessy()
        Called by: 

    parseCityZipMessy(input) -> Object
        Parses various formats of city/zip input into structured data.
        Returns: {city, state, zip}
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Billing, Responses (internal: parseAndFormatAddress)

    parseDateFromString(str) -> Date | null
        Attempts to parse a date from string in various formats. Returns null if fails.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: TeacherInvoice (internal: isCurrentOrFutureMonth)

    parseRosterData(row, headerMap, fieldMap, studentIdOverride?) -> Object
        Parses a roster data row into structured student object using header and field maps.
        Category: DATA_MANIPULATION
        Dependencies: debugLog(), extractLessonQuantityFromPackage(), formatAddress(), formatPhoneNumber(), normalizeHeader()
        Called by: 

    prefillAttendanceDatesForStudent(rowValues, headers) -> Array
        Prefills attendance dates in row values based on form responses for weekday lessons.
        Category: DATA_MANIPULATION
        Dependencies: getDateForWeekday()
        Called by: 

    promptForCustomToday() -> Date
        Returns the current date if historical data mode is disabled. Otherwise calls
        promptForDate() with a "Historical Data Entry" title and cancel message.
        Category: UI_INTERACTION
        Dependencies: isHistoricalDataInputEnabled(), promptForDate()
        Called by: Billing
        Note: Internal caller of promptForDate().

    promptForDate(config) -> Date|null
        Generic date prompt UI. Accepts config object with title, message, optional
        defaultDate, and optional cancelMessage. Returns parsed Date or null if cancelled.
        Appends default date hint to message when defaultDate is provided.
        Category: UI_INTERACTION
        Dependencies: formatDateFlexible(), debugLog()
        Called by: Billing (internal: promptForCustomToday)

    promptForHistoricalId(sheet, columnName, prefix, recordName?) -> String
        Shows UI prompt for user to manually enter a historical ID.
        Used when HISTORICAL_DATA_MODE is enabled.
        Category: UI_INTERACTION
        Dependencies: isIdAlreadyUsed()
        Called by: (internal: generateNextId)

    promptForNameWithDefault(config) -> String
        Shows UI prompt with default value for user to enter/confirm a name.
        Config: {title, message, defaultValue}
        Category: UI_INTERACTION
        Dependencies: None
        Called by: Billing

    protectSheetRanges(sheet, options?) -> Object
        Protects specified column ranges in a sheet with optional settings.
        Options: {columns, warningOnly, clearExisting}
        Returns: {success, protectedRanges?, error?}
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: Billing

    safeGet(row, index) -> Any
        Safely retrieves value from array at index. Returns empty string if out of bounds.
        Category: UTILITIES
        Dependencies: None
        Called by: 

    safeParseFloat(value) -> Number
        Safely parses any value to a float, handling null, undefined, empty strings,
        and formatted strings (e.g., "$1,234.56" or "-$50.00"). Returns 0 for invalid
        or missing values. Removes all non-numeric characters except decimal point
        and minus sign before parsing.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Billing

    setCached(key, value) -> void
        Stores value in execution cache for duration of script execution.
        Category: CACHE
        Dependencies: None
        Called by: (internal: getColumnIndices, getMonthSheets, isMonthSheet)

    setupAttendanceHeaders(sheet) -> void
        Creates and formats header row for attendance sheet with standard columns.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: createMonthlyAttendanceSheet)

    setupRosterTemplateProtection(sheet) -> void
        Sets up roster sheet protection (admin columns E:U) and data validation 
        (date picker for First Lesson Date, dropdown for Status).
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: Responses

    setupStatusValidation(sheet, lastRow) -> void
        Adds data validation dropdown for student status column.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: Responses

    shouldBeCurrency(columnName) -> Boolean
        Returns true if column should be formatted as currency.
        CRITICAL: Prevents hours columns from being formatted as currency.
        Category: VALIDATION
        Dependencies: None
        Called by: (internal: addToCurrencyCols)

    showConfirmationDialog(title, message, details, options?) -> Boolean
        Shows UI confirmation dialog with optional formatted details.
        Options: {buttonSet, confirmButton}
        Returns true if confirmed, false/throws error otherwise.
        Category: UI_INTERACTION
        Dependencies: None
        Called by: (internal: appendToMetadataWithVerification, verifyConfigurationWithUser)

    styleHeaderRow(sheet, headers) -> void
        Applies standard STYLES.HEADER formatting to first row of sheet.
        Category: UTILITIES
        Dependencies: None
        Called by: Responses (internal: createMonthlyAttendanceSheet)

    truncateString(str, maxLength) -> String
        Truncates string to maxLength, adding "..." if truncated.
        Category: UTILITIES
        Dependencies: None
        Called by: (internal: debugLog, formatLogValue)

    updateFieldMappings(fieldMapSheet, newHeaders, sourceSheetName, options?) -> Object
        Updates field mapping sheet with new form headers, checking for duplicates.
        Options: {autoAddNew, confirmBeforeAdd}
        Returns: {success, newHeaders, duplicates, error?}
        Category: FIELD_MAPPING
        Dependencies: findColumnByPartialName(), normalizeHeader()
        Called by: 


    updateParentContactFields(parentsSheet, parentRow, fieldsToUpdate, options?) -> Object
        Updates contact fields for an existing parent row. Only writes cells where
        the new value differs from the current value. Optionally updates Parent
        Lookup key and Updated checkbox.
        Options: {newLookupKey, updateUpdatedCheckbox}
        Returns: {changesMade}
        Category: SHEET_OPERATIONS
        Dependencies: debugLog(), normalizeHeader()
        Called by: Billing, Responses

    validateProgramConfiguration(programSheet, options?) -> Object
        Validates program configuration in Program List sheet.
        Options: {checkPackages, checkRates}
        Returns: {success, isValid, issues[], activePrograms[], summary, error?}
        Category: VALIDATION
        Dependencies: None
        Called by: 

    validateTemplateVariables(variableMap) -> Object
        Validates and normalizes template variables for document generation.
        Returns: {success, variables, errors[]}
        Category: VALIDATION
        Dependencies: None
        Called by: 

    verifyConfigurationWithUser(title, itemName, data, formatFunction?, options?) -> Boolean
        Shows user a formatted confirmation dialog for configuration data.
        Options: {allowSkip, skipButtonText, confirmButtonText, message}
        Returns true if confirmed, false if skipped, throws error if cancelled.
        Category: VALIDATION
        Dependencies: showConfirmationDialog()
        Called by: 


  --------------------------------------------------------------------------------
ATTENDANCE (1 function):
    getAttendanceSheetForDate

  BULK_OPERATIONS (1 function):
    bulkUpdateStudentStatus

  CACHE (3 functions):
    clearCache
    getCached
    setCached

  CONFIGURATION (12 functions):
    EnvironmentManager.get
    EnvironmentManager.set
    getCurrentSemesterName
    getGeneratedDocumentsFolder
    getRateSummary
    getRosterFolder
    getRosterFolderUrlForYear
    getSheet
    getTeacherIdByDisplayName
    getTeacherNameById
    getTemplate
    getTemplateFolder
    getWorkbook

  DATA_EXTRACTION (8 functions):
    extractLessonQuantityFromPackage
    extractNumericLessonLength
    extractRosterData
    extractSeasonFromSemester
    extractTeacherNameFromWorkbook
    extractTotalLessonsFromPackages
    getStudentIdFromRow
    parseGridInstruments

  DATA_MANIPULATION (14 functions):
    cascadeFormerStatus
    cleanName
    convertYesNoToBoolean
    copyPreviousColumnToNew
    getLessonLengthFromPackages
    interpretAgeField
    normalizeHeader
    parseAllPackageQuantities
    parseGridInstruments
    parseAndFormatAddress
    parseCityZipMessy
    parseDateFromString
    parseRosterData
    prefillAttendanceDatesForStudent
    safeParseFloat

  DATA_RETRIEVAL (8 functions):
    findColumnByPartialName
    findTeacherInRosterLookup
    findParentRow
    findStudentInContacts
    findStudentRow
    getColumnHeaders
    getColumnIndices
    getHeaderMap
    getInstrumentFamily
    getMonthSheets

  DATE_TIME (10 functions):
    getCurrentAcademicYearInfo
    getCurrentSemesterMonth
    getDateForWeekday
    getMonthNameFromDate
    getMonthNames
    getSemesterDates
    getSemesterForDate
    getWeekdayName
    getWeekdayNumber
    getYearFromSemesterName

  DEBUG (2 functions):
    debugLog
    logAllSheetHeaders

  DOCUMENT_GENERATION (2 functions):
    combineDocumentsIntoPDF
    generateDocumentFromTemplate

  ERROR_HANDLING (1 function):
    executeWithErrorHandling

  FIELD_MAPPING (2 functions):
    getFieldMappingFromSheet
    updateFieldMappings

  FORMATTING (9 functions):
    formatAddress
    formatAttendanceColumns
    formatAttendanceSheet
    formatCurrency
    formatDateFlexible
    formatLessonLengthWithMinutes
    formatLogValue
    formatPhoneNumber
    formatRosterColumns

  ID_GENERATION (3 functions):
    generateKey
    generateNextId
    generateNextIdDirect

  METADATA (1 function):
    appendToMetadataWithVerification

  RATES (1 function):
    getMostRecentRateColumn

  ROSTER (1 function):
    getTeacherGroupAssignments

  SHEET_OPERATIONS (21 functions):
    cascadeFormerStatus
    clearEmptyRows
    clearOldDebugEntries
    clearUtilityDebugLog
    copySheetWithProtections
    copyTextAttributes
    createColumnFinder
    createGroupSections
    createLessonRows
    createMonthlyAttendanceSheet
    createStudentHeader
    createStudentSections
    createUtilityDebugSheet
    deleteExtraColumns
    enableDatePickerForColumn
    freezeSheetFormulas
    protectSheetRanges
    setupAttendanceHeaders
    setupRosterTemplateProtection
    setupStatusValidation
    updateParentContactFields

  UI_INTERACTION (3 functions):
    promptForCustomToday
    promptForDate
    promptForHistoricalId
    promptForNameWithDefault
    showConfirmationDialog

  UTILITIES (8 functions):
    addToCurrencyCols
    columnToLetter
    documentAlreadyExists
    inferWorkbookKey
    insertCountFormula
    safeGet
    styleHeaderRow
    truncateString

  VALIDATION (9 functions):
    determineIfStudentIsAdult
    isCurrentOrFutureMonth
    isHistoricalDataInputEnabled
    isIdAlreadyUsed
    isMonthSheet
    shouldBeCurrency
    validateProgramConfiguration
    validateTemplateVariables
    verifyConfigurationWithUser

  VARIABLE_BUILDING (3 functions):
    buildAcademicYearVariables
    buildRateMapFromSheet
    buildRateVariables
    calculateGraduationYear

================================================================================
END OF FUNCTION DIRECTORY
================================================================================
