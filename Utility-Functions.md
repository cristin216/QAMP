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
