/*
================================================================================
UTILITY LIBRARY FUNCTION DIRECTORY
================================================================================
    Total Functions: 112 (110 standard functions + 2 EnvironmentManager methods)
     Most Recent version: 86

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
    addToCurrencyCols, appendToMetadataWithVerification, buildAcademicYearVariables,
    buildRateMapFromSheet, buildRateVariables, bulkUpdateStudentStatus, cleanName,
    clearCache, clearEmptyRows, clearOldDebugEntries, clearUtilityDebugLog,
    columnToLetter, combineDocumentsIntoPDF, convertYesNoToBoolean,
    copySheetWithProtections, copyTextAttributes, createColumnFinder, createDisplayName, 
    createGroupSections, createLessonRows, createMonthlyAttendanceSheet,
    createStudentHeader, createStudentSections, createUtilityDebugSheet, debugLog,
    deleteExtraColumns, determineIfStudentIsAdult, determineLessonLengthFromPackages,
    documentAlreadyExists, enableDatePickerForColumn, executeWithErrorHandling,
    extractLessonQuantityFromPackage, extractNumericLessonLength, extractRosterData,
    extractSeasonFromSemester, extractTeacherNameFromWorkbook, extractTotalLessonsFromPackages, findColumnByPartialName,
    findParentRow, findStudentRow, formatAddress, formatAttendanceColumns,
    formatAttendanceSheet, formatCurrency, formatDateFlexible,
    formatLessonLengthWithMinutes, formatLogValue, formatPhoneNumber,
    formatRosterColumns, freezeSheetFormulas, generateAutoId,
    generateDocumentFromTemplate, generateKey, generateNextId,generateNextIdDirect, 
    getAttendanceSheetForDate, getCached, getColumnHeaders, getColumnIndices,
    getConfig, getCurrentAcademicYearInfo, getCurrentMonthName, getCurrentSemesterMonth,
    getDateForWeekday, getFieldMappingFromSheet, getGeneratedDocumentsFolder,
    getHeaderMap, getLessonLengthFromPackages, getMonthNameFromDate, getMonthNames,
    getMonthSheets, getMostRecentRateColumn, getRateSummary, getRosterFolder,
    getRosterFolderUrlForYear, getSemesterDates, getSheet, getStudentIdFromRow, getTeacherGroupAssignments,
    getTemplate, getTemplateFolder, getWeekdayName, getWeekdayNumber, getWorkbook,
    getYearFromSemesterName, inferWorkbookKey, insertCountFormula, interpretAgeField,
    isAttendanceSheet, isCurrentOrFutureMonth, isHistoricalDataInputEnabled,
    isIdAlreadyUsed, isMonthSheet, normalizeHeader, parseAllPackageQuantities,
    parseAndFormatAddress, parseCityZipMessy, parseDateFromString, parseRosterData,
    prefillAttendanceDatesForStudent, promptForHistoricalId, promptForNameWithDefault,
    protectSheetRanges, safeGet, safeParseFloat, setCached, setupAttendanceHeaders, setupRosterTemplateProtection,
    setupStatusValidation, shouldBeCurrency, showConfirmationDialog, styleHeaderRow,
    truncateString, updateFieldMappings, validateProgramConfiguration,
    validateTemplateVariables, verifyConfigurationWithUser

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
        Called by: (internal: getConfig, getGeneratedDocumentsFolder, getRosterFolder, getSheet, getTemplateFolder, getWorkbook, prefillAttendanceDatesForStudent)

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addToCurrencyCols(currencyCols, columnNumber, headerName) -> void
        Adds a column number to the currencyCols array if the column should be 
        currency formatted. Uses shouldBeCurrency() to validate. Critical for 
        preventing hours columns from being formatted as currency.
        Category: UTILITIES
        Dependencies: debugLog(), shouldBeCurrency()
        Called by: 

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
        Called by: 

    buildRateMapFromSheet(sheet, headers, yearColIndex) -> Object
        Internal helper that builds a key-value map of rates from a sheet column.
        Maps first column (titles) to specified year column values.
        Category: VARIABLE_BUILDING
        Dependencies: None
        Called by: (internal: buildRateVariables)

    buildRateVariables() -> Object
        Builds template variables for lesson rates from the Rates sheet.
        Always uses column B (index 1) for rate values.
        Returns: {HourlyRate, HalfHourRate, LateFee, LateFeeGracePeriod}
        Category: VARIABLE_BUILDING
        Dependencies: buildRateMapFromSheet(), formatCurrency()
        Called by: 

    bulkUpdateStudentStatus(studentsSheet, statusColumn, newValue, options?) -> Object
        Updates a status column for multiple students based on conditions.
        Options: {condition, whereColumn, whereValue, skipEmpty}
        Returns: {success, updatedCount, changedRows, error?}
        Category: BULK_OPERATIONS
        Dependencies: normalizeHeader()
        Called by: 

    cleanName(name) -> String
        Cleans and standardizes a name by removing extra whitespace and 
        standardizing capitalization.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: 

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
        Called by: (internal: insertCountFormula)

    combineDocumentsIntoPDF(documents, fileName, destinationFolder) -> Object
        Combines multiple Google Docs into a single PDF file in specified folder.
        Returns: {success, pdfFile, pdfUrl, error?}
        Category: DOCUMENT_GENERATION
        Dependencies: None
        Called by: 

    convertYesNoToBoolean(value) -> Boolean
        Converts "Yes"/"No" strings to boolean values. Returns false for anything else.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: 

    copySheetWithProtections(sourceSheet, targetWorkbook, newName, options?) -> Object
        Copies a sheet to another workbook while preserving protections and formatting.
        Options: {clearContents, preserveValidation, copyProtections}
        Returns: {success, sheet, sheetName, error?}
        Category: SHEET_OPERATIONS
        Dependencies: debugLog()
        Called by: 

    copyTextAttributes(sourcePara, targetPara) -> void
        Copies text formatting attributes from source paragraph to target paragraph
        in Google Docs.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

    createColumnFinder(sheet) -> Function
        Returns a closure function that finds column indices by header name.
        Usage: var findCol = createColumnFinder(sheet); var nameCol = findCol('Name');
        Category: SHEET_OPERATIONS
        Dependencies: normalizeHeader()
        Called by: Responses, TeacherResponses, (internal: getTeacherGroupAssignments)

    createDisplayName(lastName) -> String
        Creates a display name from a last name by removing all non-alphanumeric 
        characters and trimming whitespace. Returns empty string for invalid input.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: 

    createGroupSections(sheet, groupEntries) -> void
        Creates formatted group sections in a roster sheet with headers and student rows.
        Used for group lesson rosters.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: Responses, (internal: createMonthlyAttendanceSheet)

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
        Called by: 

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
        Called by: (internal: createMonthlyAttendanceSheet)

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
        Called by: Billing, Responses, TeacherInvoice, TeacherResponses, (internal: addToCurrencyCols, clearUtilityDebugLog, createGroupSections, createMonthlyAttendanceSheet, createStudentSections, extractRosterData, extractSeasonFromSemester, extractTeacherNameFromWorkbook, formatAttendanceSheet, getAttendanceSheetForDate, getTeacherGroupAssignments, parseRosterData, setupAttendanceHeaders, setupRosterTemplateProtection, setupStatusValidation)

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
        Called by: 

    determineLessonLengthFromPackages() -> DELETED
        DEPRECATED: Use getLessonLengthFromPackages() instead.

    documentAlreadyExists(fileName, folder) -> File | Boolean
        Checks if a document with the given name exists in folder.
        Returns the File object if found, false otherwise.
        Category: UTILITIES
        Dependencies: None
        Called by: TeacherInvoice

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
        Called by: 

    extractLessonQuantityFromPackage(packageText) -> Number
        Extracts the numeric quantity from package text (e.g., "5x 30min" -> 5).
        Returns 0 if no quantity found.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: (internal: getLessonLengthFromPackages, parseAllPackageQuantities, parseRosterData)

    extractNumericLessonLength(lessonLengthValue) -> Number
        Extracts numeric lesson length in minutes from various formats.
        Handles "30", "30 min", "30min", "0.5 hour" formats.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: (internal: createLessonRows, createStudentHeader)

    extractRosterData(rosterSheet) -> Array
        Extracts all roster data from a teacher's roster sheet.
        Returns array of student objects with lesson details.
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: 

    extractSeasonFromSemester(semesterName) -> String
        Extracts season from semester name (e.g., "Fall 2024" -> "Fall").
        Category: DATA_EXTRACTION
        Dependencies: None
        Called by: 

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
        Called by: Responses
        
    findColumnByPartialName(headers, searchTerm) -> Number
        Finds column index by partial header name match (case-insensitive).
        Returns -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: None
        Called by: 

    findParentRow(parentsSheet, parentId, fallbackKey?) -> Number
        Finds the row number for a parent by Parent ID or fallback lookup key.
        Returns -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: normalizeHeader()
        Called by: 

    findStudentRow(studentSheet, studentKey) -> Number
        Finds the row number for a student using various lookup strategies
        (Student ID, first/last name combinations). Returns -1 if not found.
        Category: DATA_RETRIEVAL
        Dependencies: normalizeHeader()
        Called by: 

    formatAddress(street, city, zip) -> String
        Formats address components into a single line address string.
        Handles missing components gracefully.
        Category: FORMATTING
        Dependencies: None
        Called by: TeacherResponses, (internal: parseRosterData)

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
        Called by: TeacherInvoice, (internal: buildRateVariables)

    formatDateFlexible(date, format?) -> String
        Formats a date object with flexible format string support.
        Default format: "MM/dd/yyyy"
        Category: FORMATTING
        Dependencies: None
        Called by: TeacherInvoice

    formatLessonLengthWithMinutes(lessonLengthValue) -> String
        Formats lesson length to include "min" suffix (e.g., "30" -> "30 min").
        Category: FORMATTING
        Dependencies: None
        Called by: (internal: createStudentHeader)

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
        Called by: TeacherResponses, (internal: parseRosterData)

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

    generateAutoId() -> DELETED
        DEPRECATED: Use generateNextId() instead.

    generateDocumentFromTemplate(templateKey, variableData, fileName, destinationFolder) -> Object
        Generates a Google Doc from template with variable substitution.
        Returns: {success, document, documentId, documentUrl, error?}
        Category: DOCUMENT_GENERATION
        Dependencies: copyTextAttributes(), debugLog(), getTemplate()
        Called by: TeacherInvoice

    generateKey() -> String
        Generates a random alphanumeric key (8 characters) for lookups.
        Category: ID_GENERATION
        Dependencies: None
        Called by: TeacherResponses

    generateNextId(sheet, columnName, prefix, recordName?) -> String
        Generates the next sequential ID in format PREFIX#### (e.g., "S0042").
        Scans column for highest ID and increments.
        Category: ID_GENERATION
        Dependencies: isHistoricalDataInputEnabled(), normalizeHeader(), promptForHistoricalId()
        Called by: TeacherResponses, (internal: promptForHistoricalId)

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

    getConfig() -> Object
        Returns configuration object for current environment (test/prod).
        Access via: getConfig().paymentsId, etc.
        Category: CONFIGURATION
        Dependencies: None
        Called by: 

    getCurrentAcademicYearInfo() -> Object
        Gets current academic year info from most recent semester metadata.
        Returns: {year, startYear, endYear, semester, dates: {start, end}}
        Category: DATE_TIME
        Dependencies: getSemesterDates(), getYearFromSemesterName()
        Called by: 

    getCurrentMonthName() -> DELETED
        DEPRECATED: Use getMonthNameFromDate(date, capitalize?) instead.

    getCurrentSemesterMonth(semesterName) -> String
        Returns the starting month name for a semester (e.g., "Fall 2024" -> "September").
        Category: DATE_TIME
        Dependencies: None
        Called by: 

    getDateForWeekday(weekStartDate, weekdayName) -> Date
        Calculates date for a specific weekday within a week starting at weekStartDate.
        Category: DATE_TIME
        Dependencies: getWeekdayNumber()
        Called by: (internal: prefillAttendanceDatesForStudent)

    getFieldMappingFromSheet(fieldMapSheet) -> Object
        Reads field mapping configuration from FieldMap sheet.
        Returns object mapping form fields to internal field names.
        Category: FIELD_MAPPING
        Dependencies: None
        Called by: TeacherResponses

    getGeneratedDocumentsFolder() -> Folder
        Returns the Google Drive folder for generated documents based on environment.
        Category: CONFIGURATION
        Dependencies: getConfig()
        Called by: TeacherInvoice

    getHeaderMap(sheet) -> Object
        Creates object mapping normalized header names to column indices.
        Used for flexible column lookups.
        Category: DATA_RETRIEVAL
        Dependencies: normalizeHeader()
        Called by: Billing, Responses, TeacherInvoice, (internal: prefillAttendanceDatesForStudent)

    getLessonLengthFromPackages(qty30Package, qty45Package, qty60Package) -> Number
        Determines lesson length (30, 45, or 60) from package quantities.
        Returns 30 if no packages have quantity.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: 

    getMonthNameFromDate(date, capitalize?) -> String
        Returns month name from date object. Optionally capitalizes first letter.
        Category: DATE_TIME
        Dependencies: None
        Called by: (internal: getAttendanceSheetForDate)

    getMonthNames() -> Array
        Returns array of all month names: ["January", "February", ...]
        Category: DATE_TIME
        Dependencies: None
        Called by: TeacherInvoice, (internal: getMonthNameFromDate, isMonthSheet, setupStatusValidation)

    getMonthSheets(ss) -> Array
        Returns array of all month-based attendance sheets from spreadsheet.
        Category: DATA_RETRIEVAL
        Dependencies: getCached(), getSheet(), isMonthSheet(), setCached()
        Called by: 

    getMostRecentRateColumn(headers) -> Number
        Finds the most recent year column in Rates sheet headers.
        Currently returns 1 (column B) as fixed implementation.
        Category: RATES
        Dependencies: None
        Called by: (internal: getRateSummary)

    getRateSummary() -> String
        Returns formatted summary of current rates for user display/verification.
        Category: CONFIGURATION
        Dependencies: buildRateVariables()
        Called by: 

    getRosterFolder() -> Folder
        Returns the Google Drive folder containing roster files for current academic year.
        Category: CONFIGURATION
        Dependencies: buildAcademicYearVariables(), getConfig(), getRosterFolderUrlForYear()
        Called by: 

    getRosterFolderUrlForYear(year) -> String
        Returns the Drive URL for the roster folder of a specific year.
        Category: CONFIGURATION
        Dependencies: getConfig()
        Called by: 

    getSemesterDates(semesterName) -> Object
        Parses semester name and returns start/end dates based on semester type.
        Returns: {start: Date, end: Date}
        Category: DATE_TIME
        Dependencies: extractSeasonFromSemester(), getYearFromSemesterName()
        Called by: 

    getSheet(sheetKey) -> Sheet
        Returns a sheet object by its SHEET_MAP key (e.g., 'students', 'parents').
        Automatically infers which workbook to use.
        Category: CONFIGURATION
        Dependencies: getWorkbook(), inferWorkbookKey()
        Called by: TeacherInvoice, TeacherResponses, (internal: getCurrentSemesterMonth, getMonthSheets, getSemesterDates, getTeacherGroupAssignments)

    getStudentIdFromRow(row, headerMap) -> String|null
        Extracts student ID from a data row using flexible column name matching.
        Checks for both "Student ID" and "ID" column headers.
        Returns null if student ID not found.
        Category: DATA_EXTRACTION
        Dependencies: normalizeHeader()
        Called by: Billing (buildBillingRowFromPrevious)

    getTeacherGroupAssignments(teacherName) -> Array
        Returns array of group lesson assignments for a teacher from lookup sheet.
        Category: ROSTER
        Dependencies: createColumnFinder(), debugLog(), getSheet(), getWorkbook()
        Called by: (internal: createMonthlyAttendanceSheet)

    getTemplate(templateKey) -> File
        Returns Google Doc template file by TEMPLATE_MAP key.
        Category: CONFIGURATION
        Dependencies: getTemplateFolder()
        Called by: (internal: generateDocumentFromTemplate)

    getTemplateFolder() -> Folder
        Returns the Google Drive folder containing document templates.
        Category: CONFIGURATION
        Dependencies: getConfig()
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
        Called by: (internal: getWeekdayName)

    getWorkbook(workbookKey) -> Spreadsheet
        Returns a spreadsheet by CONFIG key (e.g., 'paymentsId', 'contactsId').
        Category: CONFIGURATION
        Dependencies: getConfig()
        Called by: TeacherInvoice, TeacherResponses, (internal: getTeacherGroupAssignments)

    getYearFromSemesterName(semesterName) -> String
        Extracts 4-digit year from semester name (e.g., "Fall 2024" -> "2024").
        Category: DATE_TIME
        Dependencies: None
        Called by: 

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

    isAttendanceSheet() -> DELETED
        DEPRECATED: Use isMonthSheet(sheetName) instead.

    isCurrentOrFutureMonth(sheetName, targetDate) -> Boolean
        Returns true if sheet represents current or future month relative to targetDate.
        Category: VALIDATION
        Dependencies: getMonthNames(), parseDateFromString()
        Called by: 

    isHistoricalDataInputEnabled() -> Boolean
        Returns value of HISTORICAL_DATA_MODE flag.
        Category: VALIDATION
        Dependencies: None
        Called by: (internal: generateNextId)

    isIdAlreadyUsed(sheet, columnName, idToCheck) -> Boolean
        Checks if an ID already exists in a column. Returns true if found.
        Category: VALIDATION
        Dependencies: normalizeHeader()
        Called by: (internal: promptForHistoricalId)

    isMonthSheet(sheetName) -> Boolean
        Returns true if sheet name exactly matches a month name.
        Category: VALIDATION
        Dependencies: getCached(), getMonthNames(), setCached()
        Called by: (internal: getAttendanceSheetForDate, getMonthSheets)

    normalizeHeader(header) -> String
        Normalizes header string for comparison (lowercase, no spaces/punctuation).
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: Responses, TeacherInvoice, TeacherResponses, (internal: bulkUpdateStudentStatus, createColumnFinder, findParentRow, findStudentRow, generateNextId, getColumnIndices, getCurrentSemesterMonth, getFieldMappingFromSheet, getHeaderMap, getSemesterDates, isIdAlreadyUsed, parseRosterData, shouldBeCurrency, updateFieldMappings)

    parseAllPackageQuantities(qty30Package, qty45Package, qty60Package) -> Object
        Parses all package quantity strings and returns structured object.
        Returns: {qty30, qty45, qty60, totalLessons}
        Category: DATA_MANIPULATION
        Dependencies: extractLessonQuantityFromPackage()
        Called by: 

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
        Called by: 

    parseDateFromString(str) -> Date | null
        Attempts to parse a date from string in various formats. Returns null if fails.
        Category: DATA_MANIPULATION
        Dependencies: None
        Called by: TeacherInvoice

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
        Called by: 

    protectSheetRanges(sheet, options?) -> Object
        Protects specified column ranges in a sheet with optional settings.
        Options: {columns, warningOnly, clearExisting}
        Returns: {success, protectedRanges?, error?}
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: (internal: setupRosterTemplateProtection)

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
        Called by: Billing (applyPastDataToRow, multiple locations)

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
        Called by: 

    setupStatusValidation(sheet, lastRow) -> void
        Adds data validation dropdown for student status column.
        Category: SHEET_OPERATIONS
        Dependencies: None
        Called by: 

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
        Called by: (internal: appendToMetadataWithVerification)

    styleHeaderRow(sheet, headers) -> void
        Applies standard STYLES.HEADER formatting to first row of sheet.
        Category: UTILITIES
        Dependencies: None
        Called by: Responses

    truncateString(str, maxLength) -> String
        Truncates string to maxLength, adding "..." if truncated.
        Category: UTILITIES
        Dependencies: None
        Called by: (internal: debugLog)

    updateFieldMappings(fieldMapSheet, newHeaders, sourceSheetName, options?) -> Object
        Updates field mapping sheet with new form headers, checking for duplicates.
        Options: {autoAddNew, confirmBeforeAdd}
        Returns: {success, newHeaders, duplicates, error?}
        Category: FIELD_MAPPING
        Dependencies: findColumnByPartialName(), normalizeHeader()
        Called by: 

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
  CATEGORIES:
  --------------------------------------------------------------------------------
  
    ATTENDANCE (1 function):
      getAttendanceSheetForDate

    BULK_OPERATIONS (1 function):
      bulkUpdateStudentStatus

    CACHE (3 functions):
      clearCache, getCached, setCached

    CONFIGURATION (11 functions):
      EnvironmentManager.get, EnvironmentManager.set, getConfig, getGeneratedDocumentsFolder,
      getRateSummary, getRosterFolder, getRosterFolderUrlForYear, getSheet, getTemplate,
      getTemplateFolder, getWorkbook

    DATA_EXTRACTION (7 functions):
      extractLessonQuantityFromPackage, extractNumericLessonLength, extractRosterData,
      extractSeasonFromSemester, extractTeacherNameFromWorkbook, extractTotalLessonsFromPackages,
      getStudentIdFromRow

    DATA_MANIPULATION (13 functions):
      cleanName, convertYesNoToBoolean, createDisplayName, getLessonLengthFromPackages,
      interpretAgeField, normalizeHeader, parseAllPackageQuantities, parseAndFormatAddress,
      parseCityZipMessy, parseDateFromString, parseRosterData, prefillAttendanceDatesForStudent,
      safeParseFloat

    DATA_RETRIEVAL (7 functions):
      findColumnByPartialName, findParentRow, findStudentRow, getColumnHeaders, getColumnIndices,
      getHeaderMap, getMonthSheets

    DATE_TIME (9 functions):
      getCurrentAcademicYearInfo, getCurrentSemesterMonth, getDateForWeekday, getMonthNameFromDate,
      getMonthNames, getSemesterDates, getWeekdayName, getWeekdayNumber, getYearFromSemesterName

    DEBUG (1 function):
      debugLog

    DOCUMENT_GENERATION (2 functions):
      combineDocumentsIntoPDF, generateDocumentFromTemplate

    ERROR_HANDLING (1 function):
      executeWithErrorHandling

    FIELD_MAPPING (2 functions):
      getFieldMappingFromSheet, updateFieldMappings

    FORMATTING (9 functions):
      formatAddress, formatAttendanceColumns, formatAttendanceSheet, formatCurrency,
      formatDateFlexible, formatLessonLengthWithMinutes, formatLogValue, formatPhoneNumber,
      formatRosterColumns

    ID_GENERATION (3 functions):
      generateKey, generateNextId, generateNextIdDirect

    METADATA (1 function):
      appendToMetadataWithVerification

    RATES (1 function):
      getMostRecentRateColumn

    ROSTER (1 function):
      getTeacherGroupAssignments

    SHEET_OPERATIONS (19 functions):
      clearEmptyRows, clearOldDebugEntries, clearUtilityDebugLog, copySheetWithProtections,
      copyTextAttributes, createColumnFinder, createGroupSections, createLessonRows,
      createMonthlyAttendanceSheet, createStudentHeader, createStudentSections,
      createUtilityDebugSheet, deleteExtraColumns, enableDatePickerForColumn,
      freezeSheetFormulas, protectSheetRanges, setupAttendanceHeaders,
      setupRosterTemplateProtection, setupStatusValidation

    UI_INTERACTION (3 functions):
      promptForHistoricalId, promptForNameWithDefault, showConfirmationDialog

    UTILITIES (8 functions):
      addToCurrencyCols, columnToLetter, documentAlreadyExists, inferWorkbookKey,
      insertCountFormula, safeGet, styleHeaderRow, truncateString

    VALIDATION (9 functions):
      determineIfStudentIsAdult, isCurrentOrFutureMonth, isHistoricalDataInputEnabled,
      isIdAlreadyUsed, isMonthSheet, shouldBeCurrency, validateProgramConfiguration,
      validateTemplateVariables, verifyConfigurationWithUser

    VARIABLE_BUILDING (3 functions):
      buildAcademicYearVariables, buildRateMapFromSheet, buildRateVariables
================================================================================
END OF FUNCTION DIRECTORY
================================================================================
*/

var _executionCache = {};

var CONFIG = {
  test: {
    paymentsId: '1cFj26gstCX6nGwqqC0dMOE8cVVY-UVE3mq5oV_zb_gE',
    contactsId: '1MjDBD9NVX9TuZ0rMYFOll5NKEq2qtwuR4Bn9keUMam8',
    billingId: '1IIs5fSsyWxa-BML9_VUpABkehltHxB8deZsLJzONhDE',
    formResponsesId: '1qbOz_jpmyhVXYDOabPV3bK1Cj0gHvmbV7wmSl-NPyqI',
    rosterFolderId: '1-1srnA3vcVDdJCkNU9Q--DY3Iem054sd',
    templateFolderId: '1zeh7Kj8ky5cdcfU5wHdu6fDpUxaDd1WF', 
    metadataSheetName: 'Semester Metadata',
    teacherInvoicesId: '1MU0uTlzUneWA3aXAQp5GLW5ZdVmMFCyAiquZvKHW6bc',
    teacherInterestId: '1rmWN4GlqTJHQSjN9ATxUwBs-cC4SUhZeDHxHHTwkPOs',
    generatedDocumentsFolderId: '1njXC_lsloSrofCCu4DQOt4hD6ncDnej0'
  },
  prod: {
    paymentsId: '1xF-HzGDIJQEwW1bxdY2RWXiI3GYA_VpTULnl00iRI9w',
    contactsId: '1fGS7CY6a4IASllyJ9vOblWDB9l4D-GbuHr4SMEyuIzo',
    billingId: '1kBZ1aKNzze-uyTtvjV-18gIaosDdDB11qih3Q3Toh1w',
    formResponsesId: '13G5U7OCLDrdZEytGIUTks8LLRXtLEd30EdEsxnN_pes',
    rosterFolderId: '1vVK02xiLYmiyRZa0zQ6eqbw81MToxuNR',
    templateFolderId: '1zeh7Kj8ky5cdcfU5wHdu6fDpUxaDd1WF', 
    metadataSheetName: 'Semester Metadata',
    teacherInvoicesId: '1ThWe85gZXnbHG5bjMXM98h0_8rvcdp-vEOohaGeSnIU',
    teacherInterestId: '1QOW0enzP1oqZIXeXx8oXKuKsF-0Eitcy3cW7cBTsCt0',
    generatedDocumentsFolderId: '1njXC_lsloSrofCCu4DQOt4hD6ncDnej0'
  }
};

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
var HISTORICAL_DATA_MODE = true;  // 👈 Just change this one line!

var SHEET_MAP = {
  // Teacher Interest Workbook
  teacherResponses: {
    name: 'Teacher Responses'
  },
  teacherFieldMap: {
    name: 'TeacherFieldMap'
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
    name: 'Program List'
  },
  billingTemplate: {
    name: 'Billing Template'
  },
  // Form Responses Workbook
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
  revisedInvoice: 'Revised Invoice Letter'
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
    
    Logger.log('Successfully appended metadata with ' + verificationSteps.length + ' verification steps');
    
    return {
      success: true,
      message: 'Metadata appended successfully',
      verificationResults: verificationResults,
      rowData: rowData
    };
    
  } catch (error) {
    Logger.log('Error appending metadata with verification: ' + error.message);
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
          Logger.log('Condition function error for row ' + (k + 1) + ': ' + conditionError.message);
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
    
    Logger.log('Bulk update completed: ' + updatedCount + ' rows updated');
    
    return {
      success: true,
      updatedCount: updatedCount,
      changedRows: changedRows,
      message: 'Updated ' + updatedCount + ' student records'
    };
    
  } catch (error) {
    Logger.log('Error in bulk student status update: ' + error.message);
    return {
      success: false,
      error: error.message,
      updatedCount: 0,
      changedRows: []
    };
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
        Logger.log('Error converting document ' + doc.name + ' to PDF: ' + docError.message);
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
    Logger.log('Error creating individual PDFs: ' + error.message);
    return {
      success: false,
      error: error.message,
      fileId: null,
      url: null
    };
  }
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
    
    Logger.log('Successfully copied sheet "' + sourceSheet.getName() + '" to "' + newName + '"');
    
    return {
      success: true,
      sheet: newSheet,
      message: 'Sheet copied successfully'
    };
    
  } catch (error) {
    Logger.log('Error copying sheet: ' + error.message);
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

function createDisplayName(lastName) {
    if (!lastName || String(lastName).trim() === '') {
    return '';
  }
  return String(lastName).trim().replace(/[^a-zA-Z0-9]/g, '');
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
        ''                                            // K - Admin Comments (empty)
      ];
      
      sheet.getRange(startRow, 1, 1, headerData.length).setValues([headerData]);
      
      // Style group header (same as student header - light blue)
      var headerRange = sheet.getRange(startRow, 1, 1, headerData.length);
      headerRange.setBackground(STYLES.SUBHEADER.background)
                 .setFontColor(STYLES.SUBHEADER.text)
                 .setFontWeight('bold');
      
      // Add green border between Comments (F) and Admin Review Date (G)
      headerRange.getCell(1, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
      
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
          ''                                          // K - Admin Comments
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
    
    // Pre-fill lesson row data (11 columns now - includes Admin Review Date)
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
      ''                                   // K - Admin Comments (admin fills)
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
    debugLog('Creating attendance sheet for ' + monthName);
    
    // Create the sheet
    var sheet = workbook.insertSheet(monthName);
    
    // Set up column headers
    setupAttendanceHeaders(sheet);
    
    // NEW ORDER: Add group sections FIRST based on teacher's group assignments
    var teacherName = extractTeacherNameFromWorkbook(workbook);
    if (teacherName) {
      var groupEntries = getTeacherGroupAssignments(teacherName);
      if (groupEntries.length > 0) {
        createGroupSections(sheet, groupEntries);
        debugLog('✅ Created group sections for ' + groupEntries.length + ' groups');
      }
    }
    
    // THEN create student sections from roster data (if any)
    if (rosterData && rosterData.length > 0) {
      createStudentSections(sheet, rosterData);
      debugLog('✅ Created sections for ' + rosterData.length + ' students');
    } else {
      debugLog('📝 Created empty attendance sheet (no students yet)');
    }
    
    // Apply formatting and protection
    formatAttendanceSheet(sheet);
    
    debugLog('✅ Successfully created ' + monthName + ' attendance sheet');
    return sheet;
    
  } catch (error) {
    debugLog('❌ Error creating attendance sheet: ' + error.message);
    throw error;
  }
}

function createStudentHeader(sheet, student, row) {
  // Get the current month name from the sheet name
  var monthName = sheet.getName(); // Should be "January", "February", etc.
  
  // Format lesson length with " minutes" suffix for header
  var lessonLengthForHeader = formatLessonLengthWithMinutes(student.lessonLength);
  
  var headerData = [
    student.id || '',                                                                    // A - Student ID
    (student.lastName || '') + ', ' + (student.firstName || '') + ' - ' + (student.instrument || ''), // B - Student Name + Instrument
    monthName,                                                                           // C - Month name
    lessonLengthForHeader,                                                              // D - Length WITH " minutes" for header
    '', '', '', '', '', '', ''                                                           // E-K - Empty columns (11 total now)
  ];
  
  // Set header data
  sheet.getRange(row, 1, 1, headerData.length).setValues([headerData]);
  
  // Check if student needs warning (less than 2 lessons remaining)
  var lessonsRemaining = student.lessonsRemaining || 0;
  var isWarning = lessonsRemaining < 2 && lessonsRemaining > 0;
  
  // Style student header
  var headerRange = sheet.getRange(row, 1, 1, headerData.length);
  if (isWarning) {
    // Warning style (pink)
    headerRange.setBackground(STYLES.WARNING.background)
               .setFontColor(STYLES.WARNING.text)
               .setFontWeight('bold');
    
    // Update Length column to include warning with " minutes" suffix
    var numericLength = extractNumericLessonLength(student.lessonLength);
    sheet.getRange(row, 4).setValue(numericLength + ' minutes - WARNING: Only ' + lessonsRemaining + ' lessons left!')
                          .setWrap(false);  // Use overflow instead of text wrap
  } else {
    // Normal style (light blue)
    headerRange.setBackground(STYLES.SUBHEADER.background)
               .setFontColor(STYLES.SUBHEADER.text)
               .setFontWeight('bold');
  }
  
  // Add green border between Comments (F) and Admin Review Date (G)
  headerRange.getCell(1, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  return row + 1;
}

function createStudentSections(sheet, rosterData) {
  var lastRow = sheet.getLastRow();
  var currentRow;
  
  // If sheet only has headers (row 1), start at row 2
  // Otherwise, start after the last row with data
  if (lastRow <= 1) {
    currentRow = 2;
  } else {
    currentRow = lastRow + 2; // +1 for spacing, +1 for next available row
  }
  
  for (var i = 0; i < rosterData.length; i++) {
    var student = rosterData[i];
    
    // Create student header row
    currentRow = createStudentHeader(sheet, student, currentRow);
    
    // Create 5 lesson rows for this student
    currentRow = createLessonRows(sheet, student, currentRow, 5);
    
    // Add spacing between students (except last one)
    if (i < rosterData.length - 1) {
      currentRow += 1; // One empty row between students
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
    Logger.log('Error checking if document exists: ' + error.message);
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
      Logger.log('[' + logLevel + '] ' + context + ': ' + successMessage);
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
    Logger.log('[ERROR] ' + context + ': ' + errorMessage);
    if (error.stack) {
      Logger.log('Stack trace: ' + error.stack);
    }
    
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
      status: row[getCol('Status')] || 'Active'
    };
    
    // Only include active students (not dropped)
    if (student.status.toString().toLowerCase() !== 'dropped') {
      students.push(student);
    }
  }
  
  debugLog('extractRosterData - Extracted data for ' + students.length + ' active students');
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
      return i + 1; // Return 1-based index
    }
  }
  return -1;
}

function findParentRow(parentsSheet, parentId, fallbackKey) {
  var data = parentsSheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = -1;
  var lookupCol = -1;
  
  // Debug the headers first
  console.log("findParentRow - Available headers: " + JSON.stringify(headers));
  
  for (var i = 0; i < headers.length; i++) {
    var normalizedHeader = normalizeHeader(headers[i]);
    if (normalizedHeader === 'parentid') {
      idCol = i;
      console.log("Found Parent ID column at index: " + i);
    }
    if (normalizedHeader === 'parentlookup') {
      lookupCol = i;
      console.log("Found Parent Lookup column at index: " + i);
    }
  }
  
  if (idCol === -1 || lookupCol === -1) {
    console.log("❌ Missing columns - idCol: " + idCol + ", lookupCol: " + lookupCol);
    throw new Error("Missing 'Parent ID' or 'Parent Lookup' column in Parents sheet.");
  }
  
  // Debug logging
  console.log("findParentRow - Searching for parentId: '" + parentId + "', fallbackKey: '" + fallbackKey + "'");
  console.log("Found " + (data.length - 1) + " existing parent rows to check");
  
  for (var j = 1; j < data.length; j++) {
    var row = data[j];
    var rowId = (row[idCol] || '').toString().trim();
    var rowKey = (row[lookupCol] || '').toString().toLowerCase().trim();
    
    console.log("Row " + j + " - ID: '" + rowId + "', Key: '" + rowKey + "', Comparing to: '" + fallbackKey + "'");
    console.log("Match check - ID match: " + (rowId === parentId) + ", Key match: " + (rowKey === fallbackKey));
    
    if (rowId === parentId || rowKey === fallbackKey) {
      console.log("✅ Found match at row " + (j + 1));
      return j + 1;
    }
  }
  
  console.log("❌ No parent match found");
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
    Logger.log("Error in formatAttendanceColumns: " + e.message);
  }
}

function formatAttendanceSheet(sheet) {
  try {
    // Freeze header row
    sheet.setFrozenRows(1);
    
    var lastRow = sheet.getLastRow();
    var maxRows = sheet.getMaxRows();
    
    // Set up ROSTER-STYLE date format and validation for Date column (column C)
    // Apply to entire column so it works even on empty sheets
    var dateRange = sheet.getRange(2, 3, maxRows - 1, 1);
    dateRange.setNumberFormat('ddd, M/d');  // Format: "Tue, 1/15"
    
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .build();
    dateRange.setDataValidation(dateRule);
    
    // Set up data validation for Status column (column E) - skip student header rows
    if (lastRow > 1) {
      var data = sheet.getDataRange().getValues();
      var monthNames = getMonthNames().map(function(name) {
        return name.toLowerCase();
      });
      
      var statusOptions = ['Lesson', 'No Show', 'No Lesson'];
      var statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statusOptions)
        .setAllowInvalid(false)
        .build();
      
      // Apply validation to each row individually, skipping student header rows
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowNum = i + 1;
        var dateValue = row[2];  // Column C (Date)
        var lengthValue = row[3]; // Column D (Length)
        
        var isHeaderRow = false;
        
        // Check if date contains a month name
        if (dateValue && typeof dateValue === 'string') {
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
        
        // Only apply status dropdown to non-header rows
        if (!isHeaderRow) {
          sheet.getRange(rowNum, 5).setDataValidation(statusRule);
        }
      }
    }
    
    // Protect admin columns (G through K) with warning (updated range for new column count)
    var adminRange = sheet.getRange(1, 7, maxRows, 5);  // 5 columns now (G through K)
    var protection = adminRange.protect();
    protection.setDescription('Admin columns - automated data');
    protection.setWarningOnly(true);
    
    // Set text wrapping for ALL columns (as requested)
    var maxCols = sheet.getLastColumn();
    sheet.getRange(1, 1, maxRows, maxCols).setWrap(true);
    
    // Column D (Length) should overflow instead of wrap
    sheet.getRange(1, 4, maxRows, 1).setWrap(false);
    
    debugLog('✅ Applied formatting, protection, dropdown validation (skipping header rows), roster-style dates, and text wrapping');
    
  } catch (error) {
    debugLog('⚠️ Error in formatting: ' + error.message);
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
    Logger.log("❌ Error in formatRosterColumns: " + e.message);
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
    Logger.log('Error generating document from template: ' + error.message);
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
    debugLog('getAttendanceSheetForDate - Error: ' + error.message);
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

function getConfig() {
  var env = EnvironmentManager.get();
  return CONFIG[env];
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
      return "January";
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
      return "January";
    }
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[nameCol].toString().trim() === semesterName) {
        var startDate = new Date(row[startCol]);
        var monthNames = [
          "January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"
        ];
        var monthName = monthNames[startDate.getMonth()];
        return monthName;
      }
    }
    
    return "January";
    
  } catch (error) {
    return "January";
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
  
  var monthNames = getMonthNames();
  var monthName = monthNames[date.getMonth()];
  
  // If capitalize parameter is true, return capitalized
  if (capitalize) {
    return monthName;
  }
  
  return monthName.toLowerCase(); // Default lowercase for backend
}

function getMonthNames() {
  return ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December'];
}

function getMonthSheets(ss) {
  if (!ss) return [];
  
  // Check cache first
  var ssId = ss.getId();
  var cacheKey = 'monthSheets_' + ssId;
  var cached = getCached(cacheKey);
  if (cached !== undefined) return cached;
  
  // Perform search
  var sheets = ss.getSheet();
  var monthSheets = [];
  
  for (var i = 0; i < sheets.length; i++) {
    if (isMonthSheet(sheets[i].getName())) {
      monthSheets.push(sheets[i]);
    }
  }
  
  // Cache and return
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

function getTeacherGroupAssignments(teacherName) {
  try {
    // Get Teacher Roster Lookup sheet from formResponses workbook
    var teacherLookupSheet = getSheet('teacherRosterLookup');
    
    if (!teacherLookupSheet) {
      debugLog('getTeacherGroupAssignments', 'ERROR', 'Teacher Roster Lookup sheet not found', '', '');
      return [];
    }
    
    // Find teacher's group assignment using createColumnFinder
    var getCol = createColumnFinder(teacherLookupSheet);
    var nameCol = getCol('Teacher Name');
    var groupAssignmentCol = getCol('Group Assignment');
    
    if (nameCol === 0 || groupAssignmentCol === 0) {
      debugLog('getTeacherGroupAssignments', 'ERROR', 'Required columns not found', 
               'Name col: ' + nameCol + ', Group col: ' + groupAssignmentCol, '');
      return [];
    }
    
    var teacherData = teacherLookupSheet.getDataRange().getValues();
    var groupAssignmentRaw = '';
    
    for (var i = 1; i < teacherData.length; i++) {
      if (teacherData[i][nameCol - 1] === teacherName) {
        groupAssignmentRaw = teacherData[i][groupAssignmentCol - 1] || '';
        break;
      }
    }
    
    if (!groupAssignmentRaw) {
      debugLog('getTeacherGroupAssignments', 'DEBUG', 'No group assignment found', 'Teacher: ' + teacherName, '');
      return [];
    }
    
    // Parse group assignments (comma-separated)
    var assignedPrograms = groupAssignmentRaw.toString().split(',').map(function(p) { return p.trim(); }).filter(function(p) { return p; });
    
    // Get Programs List to map program names to G-IDs
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
    
    // Process each assigned program
    for (var j = 0; j < assignedPrograms.length; j++) {
      var programName = assignedPrograms[j];
      
      // Find program in Programs List
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
            // Handle packages - expand to component programs
            var components = aliasFor.split(',').map(function(c) { return c.trim(); }).filter(function(c) { return c; });
            
            // Find G-IDs for each component
            for (var c = 0; c < components.length; c++) {
              var componentName = components[c];
              for (var m = 1; m < programsData.length; m++) {
                var componentRow = programsData[m];
                if (componentRow[programNameCol - 1] === componentName && componentRow[activeCol - 1] === true && componentRow[groupIdCol - 1]) {
                  groupEntries.push({
                    groupId: componentRow[groupIdCol - 1],
                    groupName: componentName
                  });
                }
              }
            }
          } else if (groupId) {
            // Individual program with G-ID
            groupEntries.push({
              groupId: groupId,
              groupName: programName
            });
          }
          break;
        }
      }
    }
    
    debugLog('getTeacherGroupAssignments', 'INFO', 'Found group assignments', 'Teacher: ' + teacherName + ', Groups: ' + groupEntries.length, '');
    return groupEntries;
    
  } catch (error) {
    debugLog('getTeacherGroupAssignments', 'ERROR', 'Error getting teacher group assignments', 'Teacher: ' + teacherName, error.message);
    return [];
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
    teachersAndAdmin: 'contacts',
    students: 'contacts',
    parents: 'contacts',
    instrumentList: 'contacts',
    billingMetadata: 'billing',
    semesterMetadata: 'billing',
    yearMetadata: 'billing',
    rates: 'billing',
    programList: 'billing',
    billingTemplate: 'billing',
    calendar: 'formResponses',
    fieldMap: 'formResponses',
    teacherRosterLookup: 'formResponses', 
    ledgerTemplate: 'payments'
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
  // Map month names to numbers (0-11)
  var monthMap = {
    'january': 0, 'february': 1, 'march': 2, 'april': 3,
    'may': 4, 'june': 5, 'july': 6, 'august': 7,
    'september': 8, 'october': 9, 'november': 10, 'december': 11
  };
  
  var sheetNameLower = sheetName.toLowerCase().trim();
  var sheetMonth = monthMap[sheetNameLower];
  
  // If sheet name isn't a recognized month, return false
  if (sheetMonth === undefined) {
    return false;
  }
  
  var targetMonth = targetDate.getMonth(); // 0-11
  
  // Simple comparison: sheet month >= target month means current or future
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
  
  // Check cache first
  var cacheKey = 'isMonth_' + sheetName;
  var cached = getCached(cacheKey);
  if (cached !== undefined) return cached;
  
  // Perform check
  var lowerSheetName = sheetName.toLowerCase();
  var monthNames = getMonthNames().map(function(name) {
    return name.toLowerCase();
  });
  
  var result = monthNames.indexOf(lowerSheetName) !== -1;
  
  // Cache and return
  return setCached(cacheKey, result);
}

function normalizeHeader(header) {
  if (!header) return '';
  var str = header.toString();
  return str.replace(/["\n\r]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
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
    debugLog('ERROR in parseRosterData: ' + error.message);
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
    Logger.log('Error in promptForHistoricalId: ' + error.message);
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

function protectSheetRanges(sheet, options) {
  options = options || {};
  var columns = options.columns || [];
  var warningOnly = options.warningOnly !== false; // Default true
  var clearExisting = options.clearExisting || false;
  
  try {
    // Clear existing protections if requested
    if (clearExisting) {
      var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < existingProtections.length; i++) {
        existingProtections[i].remove();
      }
      debugLog('protectSheetRanges', 'INFO', 'Cleared existing protections', 'Count: ' + existingProtections.length, '');
    }
    
    // Apply protection to each specified column range
    for (var j = 0; j < columns.length; j++) {
      var columnRange = columns[j];
      
      // Parse column range (e.g., "E:U" or "A:A")
      var range = sheet.getRange(columnRange);
      
      // Protect the range
      var protection = range.protect();
      protection.setDescription('Protected columns: ' + columnRange);
      
      // Set warning-only mode if requested
      if (warningOnly) {
        protection.setWarningOnly(true);
      }
    }
    
    debugLog('protectSheetRanges', 'INFO', 'Protected columns successfully', 
             'Columns: ' + columns.join(', ') + (warningOnly ? ' (warning only)' : ''), '');
    
    return {
      success: true,
      protectedRanges: columns.length
    };
    
  } catch (error) {
    debugLog('protectSheetRanges', 'ERROR', 'Error protecting sheet ranges', '', error.message);
    return {
      success: false,
      error: error.message
    };
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
    'Admin Review Date',   // G - NEW COLUMN
    'Invoice Date',        // H
    'Payment Date',        // I
    'Invoice Number',      // J
    'Admin Comments'       // K
  ];
  
  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Style header row - WITH text wrapping
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground(STYLES.HEADER.background)
             .setFontColor(STYLES.HEADER.text)
             .setFontWeight('bold')
             .setHorizontalAlignment('center')
             .setWrap(true);
  
  // Updated column widths (added Admin Review Date column - same width as other date columns)
  var widths = [60, 220, 80, 80, 95, 220, 75, 110, 75, 100, 220];  // 11 columns now
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
  
  // Add thick green border between Comments (F) and Admin Review Date (G)
  headerRange.getCell(1, 6).setBorder(null, null, null, true, null, null, STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
  
  debugLog('✅ Headers set up with formatting and green border between F|G');
}

function setupRosterTemplateProtection(sheet) {
  try {
    // Protect admin columns (E through U) with warning
    protectSheetRanges(sheet, {
      columns: ['E:U'],
      warningOnly: true,
      clearExisting: true
    });
    
    // Set up date validation for First Lesson Date column (B)
    var dateRange = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1);
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .build();
    dateRange.setDataValidation(dateRule);
    
    // Set up dropdown validation for Status column (T)
    var statusRange = sheet.getRange(2, 20, sheet.getMaxRows() - 1, 1);
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['active', 'dropped'], true)
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    debugLog("✅ Roster protection, date validation, and status dropdown applied");
    
  } catch (error) {
    debugLog("⚠️ Error in roster protection: " + error.message);
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
          var allMonthNames = getMonthNames();
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
    Logger.log('Error in showConfirmationDialog: ' + error.message);
    return false;
  }
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
    Logger.log('Error updating field mappings: ' + error.message);
    return {
      success: false,
      error: error.message,
      newHeaders: 0,
      duplicates: 0
    };
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
    Logger.log('Error validating program configuration: ' + error.message);
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
        Logger.log('Format function error: ' + formatError.message);
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
      Logger.log('User confirmed ' + itemName + ' configuration');
      return true;
    } else if (allowSkip && response === ui.Button.NO) {
      Logger.log('User skipped ' + itemName + ' verification');
      return false; // Skipped, but continue
    } else {
      Logger.log('User cancelled ' + itemName + ' verification');
      throw new Error('User cancelled ' + itemName + ' verification');
    }
    
  } catch (error) {
    Logger.log('Error in verification dialog: ' + error.message);
    
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