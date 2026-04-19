================================================================================
TEACHERINVOICE FUNCTION DIRECTORY
================================================================================
    Total Functions: 57
    Most Recent version: 29

    This directory provides a quick reference for all functions in TeacherInvoice script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), UtilityScriptLibrary.function2()

  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    addStudentLineItem, addTeacherHeaderRow, buildTeacherInvoiceFileName, 
    buildTeacherInvoiceVariableMap, calculateLessonRateAndCost, calculateLessonsStartingDate,
    collectUninvoicedLessonsUpToDate, createMonthlyInvoiceSheet, extractAndMarkLessonsFromSheet,
    extractLessonFromRow, extractLessonsFromAttendanceSheet, extractTeacherInvoiceNumbers,
    extractTeachersFromFormattedSheet, formatDateForInput, formatInvoiceSheet,
    generateDefaultInvoicePeriod, generateDefaultMonthName, generateInvoiceNumber,
    generateMonthlyTeacherInvoices, generateSingleTeacherInvoice, generateTeacherInvoiceDocuments,
    getActiveTeacherList, getAttendanceSheetsFromWorkbook, getCurrentSemesterName,
    getRateForSemester, getRateKeyForProgram, getRatesLookupForSemester, getSemesterForDate,
    getTeacherContactInfo, getTeacherInfoByName, getTeacherInfoFromLessonGroup,
    getTeacherInvoicesFolder, getTeacherInvoicingMetadata, getUninvoicedLessonsForTeacher,
    groupLessonsByTeacherAndType, isMonthlyInvoiceSheet, loadProgramRateKeysCache,
    loadRatesCache, onOpen, parseLessonDate, parseStudentName, populateInvoiceSheetFromLessons,
    promptForCutoffDate, promptForInvoiceDate, promptForInvoicePeriod, promptForMonthName,
    setupMetadataStatusDropdown, showCombinedErrorDetails, showInvoiceGenerationResults,
    showInvoiceGenerationUI, showLessonCollectionUI, showResultsSummaryUI, showTeacherInvoiceResults,
    updateMetadataStatus, updateTeacherInvoiceHistory, validateLessonData, writeTeacherInvoicingMetadata

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addStudentLineItem(sheet, row, columnMap, lessonGroup, teacherName, ratesCache, programRateKeysCache) -> Number
        Adds a student line item row to the monthly invoice sheet with calculated rate and cost.
        Uses cached rates to avoid repeated database lookups. Returns the next available row number.
        Category: SHEET_OPERATIONS
        Local functions used: calculateLessonRateAndCost(), parseStudentName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addTeacherHeaderRow(sheet, row, columnMap, teacherInfo, invoiceDate, invoiceNumber, invoicePeriod) -> Number
        Adds a teacher header row to the monthly invoice sheet with invoice metadata.
        Applies bold formatting with light green background. Returns the next available row number.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    buildTeacherInvoiceFileName(teacherData, variables) -> String
        Constructs the filename for a teacher's invoice document using teacher name and invoice number.
        Format: "Teacher Name - InvoiceNumber". Returns filename string.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildTeacherInvoiceVariableMap(teacherData, teacherContactInfo, metadata) -> Object
        Builds a variable mapping object for populating the teacher invoice template.
        Creates dynamic content for student rows and formats dates using metadata when available.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.formatCurrency(), 
                                UtilityScriptLibrary.debugLog()

    calculateLessonRateAndCost(lessonGroup, ratesCache, programRateKeysCache) -> Object
        Calculates the hourly rate, lesson rate, and total cost for a lesson group.
        Handles both student lessons and group sessions with different rate structures.
        Returns object with hourlyRate, lessonRate, and totalCost.
        Category: RATE_CALCULATION
        Local functions used: getSemesterForDate(), getRateKeyForProgram(), getRateForSemester()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.formatDateFlexible()

    calculateLessonsStartingDate(metadataSheet) -> Date|null
        Calculates the starting date for lessons by adding 1 day to the previous invoice's ending date.
        Returns null if no previous metadata exists (signals to use first day of cutoff month).
        Category: DATE_CALCULATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    collectUninvoicedLessonsUpToDate(cutoffDate?, invoiceDate?, invoicePeriod?) -> Object
        Main data collection function. Gathers all uninvoiced lessons from active teachers up to cutoff date.
        If no parameters provided, prompts user for input. Returns comprehensive results object with
        lessons, errors, validation, summary statistics, and parameters used.
        Category: DATA_COLLECTION
        Local functions used: promptForCutoffDate(), promptForInvoiceDate(), promptForInvoicePeriod(),
                              getActiveTeacherList(), generateInvoiceNumber(), getUninvoicedLessonsForTeacher(),
                              groupLessonsByTeacherAndType(), validateLessonData()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.formatDateFlexible()

    createMonthlyInvoiceSheet(month) -> Sheet
        Creates a new monthly invoice sheet by copying the Monthly Template.
        If sheet already exists, clears data while preserving headers. Returns the sheet object.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractAndMarkLessonsFromSheet(sheet, teacherData, cutoffDate, formattedInvoiceDate, invoiceNumber) -> Array
        Single-pass function that extracts uninvoiced lessons from an attendance sheet AND
        immediately marks them with invoice data. Returns array of lesson objects.
        Category: DATA_EXTRACTION
        Local functions used: parseLessonDate()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    extractLessonFromRow(row, headerMap, teacherData, cutoffDate, sheet, sheetRowIndex) -> Object|null
        Extracts and validates a single lesson from a data row. Performs validation on lesson length
        based on type (student vs group). Writes validation errors to Admin Comments column.
        Returns lesson object or null if invalid.
        Category: DATA_EXTRACTION
        Local functions used: parseLessonDate()
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractLessonsFromAttendanceSheet(sheet, teacherData, cutoffDate) -> Array
        Extracts uninvoiced lessons from a single attendance sheet without marking them.
        Used for read-only lesson collection. Returns array of lesson objects.
        Category: DATA_EXTRACTION
        Local functions used: extractLessonFromRow()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    extractTeacherInvoiceNumbers(invoiceSheet) -> Object
        Extracts existing invoice numbers for each teacher from a monthly invoice sheet.
        Returns object mapping teacher names to invoice numbers.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    extractTeachersFromFormattedSheet(sheet) -> Array
        Parses a formatted monthly invoice sheet to extract teacher data and their student rows.
        Only includes teachers without existing URLs (not yet generated). Returns array of teacher objects
        with studentRows nested arrays.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    formatDateForInput(date) -> String
        Formats a date object for UI input prompts. Returns string in MM/DD/YYYY format.
        Category: DATE_FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    formatInvoiceSheet(sheet) -> void
        Applies formatting to the monthly invoice sheet including font, wrapping, and currency formats.
        Does not auto-resize columns to preserve template column widths.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    generateDefaultInvoicePeriod(cutoffDate) -> String
        Generates default invoice period text from cutoff date. Returns string in "Month Year" format.
        Category: DATE_FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.debugLog()

    generateDefaultMonthName(cutoffDate) -> String
        Extracts month name from cutoff date for sheet naming. Returns month name string.
        Category: DATE_FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.debugLog()

    generateInvoiceNumber(teacherId, invoiceDate) -> String
        Generates a unique invoice number combining teacher ID and date.
        Format: "TeacherID-YYYYMMDD". Returns invoice number string.
        Category: INVOICE_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    generateMonthlyTeacherInvoices(month, cutoffDate, invoiceDate, lessonResults) -> Object
        Main invoice sheet generation function. Creates monthly sheet, populates with lesson data,
        and formats it. Returns success status with teacher count, line item count, and marked count.
        Category: INVOICE_GENERATION
        Local functions used: createMonthlyInvoiceSheet(), populateInvoiceSheetFromLessons(),
                              formatInvoiceSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateSingleTeacherInvoice(teacherData, sheet, metadata) -> Object
        Generates a single teacher's invoice PDF document from template and uploads to Drive.
        Sets document to view-only and writes URL back to invoice sheet. Returns result object
        with success status, fileId, and url.
        Category: DOCUMENT_GENERATION
        Local functions used: buildTeacherInvoiceVariableMap(), buildTeacherInvoiceFileName(),
                              getTeacherInvoicesFolder()
        Utility functions used: UtilityScriptLibrary.documentAlreadyExists(),
                                UtilityScriptLibrary.generateDocumentFromTemplate(),
                                UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    generateTeacherInvoiceDocuments() -> void
        Main UI function for bulk PDF generation. Validates sheet, extracts teachers,
        generates documents for each, updates metadata status, and shows results summary.
        Category: DOCUMENT_GENERATION
        Local functions used: isMonthlyInvoiceSheet(), getTeacherInvoicingMetadata(),
                              extractTeachersFromFormattedSheet(), generateSingleTeacherInvoice(),
                              updateTeacherInvoiceHistory(), updateMetadataStatus(),
                              showTeacherInvoiceResults()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getActiveTeacherList() -> Array
        Retrieves list of active teachers from Teacher Roster Lookup sheet.
        Only includes teachers with "Active" status and roster URLs. Returns array of teacher objects.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    getAttendanceSheetsFromWorkbook(spreadsheet) -> Array
        Identifies and returns all attendance sheets (monthly sheets) from a teacher's roster workbook.
        Matches sheet names against standard month names. Returns array of Sheet objects.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.debugLog()

    getCurrentSemesterName() -> String
        Retrieves the most recent semester name from Semester Metadata sheet in billing workbook.
        Returns semester name string or "Unknown Semester" if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog()

    getRateForSemester(semesterName, rateType, ratesCache) -> Number
        Looks up a specific rate from the pre-loaded rates cache for a given semester.
        Returns the rate value.
        Category: RATE_CALCULATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRateKeyForProgram(programName, programRateKeysCache) -> String|null
        Retrieves the rate key for a given program name from the pre-loaded program rate keys cache.
        Returns rate key string or null if not found.
        Category: RATE_CALCULATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRatesLookupForSemester(semesterName) -> String
        Retrieves the Rates Verification value for a specific semester from Semester Metadata sheet.
        Returns rates lookup string or empty string if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    getSemesterForDate(date) -> String
        Determines which semester a given date falls within by querying Semester Metadata sheet.
        Returns semester name string.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.formatDateFlexible(),
                                UtilityScriptLibrary.debugLog()

    getTeacherContactInfo(teacherId) -> Object
        Looks up teacher's first name, last name, and address from Teachers and Admin sheet
        in contacts workbook. Returns object with firstName, lastName, and address properties.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    getTeacherInfoByName(teacherName) -> Object|null
        Looks up comprehensive teacher info from Teacher Roster Lookup sheet by name.
        Tries exact match first, then falls back to last name matching. Returns teacher info object
        or null if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    getTeacherInfoFromLessonGroup(lessonGroup) -> Object
        Extracts teacher information from a lesson group object and enriches it with contact info.
        Returns object with teacherName, teacherId, lastName, firstName, and address.
        Category: DATA_EXTRACTION
        Local functions used: getTeacherContactInfo()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getTeacherInvoicesFolder(monthName) -> Folder
        Gets or creates the monthly subfolder within Teacher Invoices folder for storing
        generated invoice documents. Returns Drive Folder object.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getGeneratedDocumentsFolder(),
                                UtilityScriptLibrary.debugLog()

    getTeacherInvoicingMetadata(invoiceMonth) -> Object
        Retrieves metadata for a specific invoice month from Teacher Invoicing Metadata sheet.
        Returns object with invoiceMonth, lessonsStartingDate, lessonsEndingDate, invoiceDate,
        ratesLookup, semesterName, and status.
        Category: METADATA_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    getUninvoicedLessonsForTeacher(teacherData, cutoffDate, invoiceDate, invoiceNumber) -> Array
        Single-pass function that opens a teacher's roster, processes all attendance sheets,
        extracts lessons, and marks them as invoiced. Returns array of uninvoiced lesson objects.
        Category: DATA_COLLECTION
        Local functions used: getAttendanceSheetsFromWorkbook(), extractAndMarkLessonsFromSheet()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    groupLessonsByTeacherAndType(allLessons) -> Object
        Groups flat array of lessons by teacher, student, and lesson length for invoicing.
        Creates aggregated line items with quantity counts. Returns object keyed by
        "teacherName|studentId|lessonLength".
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    isMonthlyInvoiceSheet(sheet) -> Boolean
        Validates whether a sheet is a monthly invoice sheet by checking for expected headers.
        Returns true if at least 4 key headers are found.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    loadProgramRateKeysCache() -> Object
        Loads all program rate keys from Programs List sheet into a cache object for fast lookups.
        Returns object mapping normalized program names to rate keys.
        Category: RATE_CALCULATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog()

    loadRatesCache() -> Object
        Loads all current rates from Rates sheet into a cache object for fast lookups.
        Returns object mapping rate types to rate values.
        Category: RATE_CALCULATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog()

    onOpen() -> void
        Creates custom menu in the spreadsheet UI when the Teacher Invoices workbook opens.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    parseLessonDate(dateValue) -> Date|null
        Parses various date formats from attendance sheets into Date objects.
        Handles "Day, M/D" format, standard date strings, and Date objects. Returns Date or null.
        Category: DATE_PARSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.parseDateFromString(), UtilityScriptLibrary.debugLog()

    parseStudentName(studentName) -> Object
        Parses student/group name into firstName and lastName components.
        Handles "Last, First" format for students and plain text for group names.
        Returns object with firstName and lastName properties.
        Category: DATA_PARSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateInvoiceSheetFromLessons(sheet, groupedLessons, invoiceDate, invoicePeriod) -> Object
        Main population function that writes all teacher headers and student line items to
        the monthly invoice sheet. Returns object with teacherCount and lineItemCount.
        Category: SHEET_OPERATIONS
        Local functions used: loadRatesCache(), loadProgramRateKeysCache(), getTeacherInfoFromLessonGroup(),
                              generateInvoiceNumber(), addTeacherHeaderRow(), addStudentLineItem()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    promptForCutoffDate() -> Date|null
        Prompts user for lesson cutoff date with validation. Tries UtilityScriptLibrary parser
        first, then falls back to native Date constructor. Returns Date object or null if cancelled.
        Category: UI_PROMPT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.parseDateFromString(), UtilityScriptLibrary.debugLog()

    promptForInvoiceDate() -> Date|null
        Prompts user for invoice date with today's date as default. Returns Date object or null
        if cancelled.
        Category: UI_PROMPT
        Local functions used: formatDateForInput()
        Utility functions used: UtilityScriptLibrary.parseDateFromString(), UtilityScriptLibrary.debugLog()

    promptForInvoicePeriod(cutoffDate) -> String|null
        Prompts user for invoice period description with default generated from cutoff date.
        Allows custom text input. Returns period string or null if cancelled.
        Category: UI_PROMPT
        Local functions used: generateDefaultInvoicePeriod()
        Utility functions used: UtilityScriptLibrary.debugLog()

    promptForMonthName(cutoffDate) -> String|null
        Prompts user for invoice sheet month name with default generated from cutoff date.
        Capitalizes first letter of input. Returns month name or null if cancelled.
        Category: UI_PROMPT
        Local functions used: generateDefaultMonthName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    setupMetadataStatusDropdown(metadataSheet) -> void
        Sets up data validation dropdown for Status column in metadata sheet with values:
        'Collected', 'Generated', 'Sent', 'Paid'.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    showCombinedErrorDetails(lessonResults, invoiceResults) -> void
        Displays detailed error and validation issue alert to user. Shows up to 5 errors
        from each category with count of additional errors.
        Category: UI_DISPLAY
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    showInvoiceGenerationResults(lessonResults, invoiceResults) -> void
        Displays comprehensive invoice generation results to user including sheet info,
        lesson data, and status. Offers to show error details if issues exist.
        Category: UI_DISPLAY
        Local functions used: showCombinedErrorDetails()
        Utility functions used: UtilityScriptLibrary.debugLog()

    showInvoiceGenerationUI() -> void
        Main UI workflow function that orchestrates the entire invoice generation process.
        Prompts for all parameters, collects lessons, generates sheet, writes metadata,
        and shows results.
        Category: UI_WORKFLOW
        Local functions used: promptForCutoffDate(), promptForInvoiceDate(), promptForMonthName(),
                              promptForInvoicePeriod(), collectUninvoicedLessonsUpToDate(),
                              generateMonthlyTeacherInvoices(), writeTeacherInvoicingMetadata(),
                              showInvoiceGenerationResults()
        Utility functions used: UtilityScriptLibrary.debugLog()

    showLessonCollectionUI() -> void
        Alternative UI workflow for lesson collection only (without invoice sheet generation).
        Collects data and shows summary results.
        Category: UI_WORKFLOW
        Local functions used: collectUninvoicedLessonsUpToDate(), showResultsSummaryUI()
        Utility functions used: UtilityScriptLibrary.debugLog()

    showResultsSummaryUI(results) -> void
        Displays comprehensive lesson collection results to user including cutoff date,
        invoice date, period, summary statistics, and status.
        Category: UI_DISPLAY
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    showTeacherInvoiceResults(results) -> void
        Displays teacher invoice PDF generation results including counts of successful,
        skipped, and error documents. Shows first 3 errors inline.
        Category: UI_DISPLAY
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    updateMetadataStatus(invoiceMonth, newStatus) -> void
        Updates the Status field for a specific invoice month in Teacher Invoicing Metadata sheet.
        Category: METADATA_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    updateTeacherInvoiceHistory(teacherData, invoiceResult, metadata) -> void
        Adds a new invoice record to the teacher's Invoice Log sheet in their roster workbook.
        Includes invoice number, date, period (date range from metadata), URL, and total amount.
        Category: METADATA_MANAGEMENT
        Local functions used: getTeacherInfoByName()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.formatCurrency(),
                                UtilityScriptLibrary.debugLog()

    validateLessonData(groupedLessons) -> Object
        Validates grouped lesson data for completeness and correctness. Checks for required fields,
        valid lesson lengths, positive quantities, and valid lesson dates. Returns validation object
        with issues array and isValid boolean.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod) -> Object
        Writes new invoice metadata record to Teacher Invoicing Metadata sheet including date ranges,
        semester info, and rates lookup. Returns success object with lessonsStartingDate and
        lessonsEndingDate.
        Category: METADATA_MANAGEMENT
        Local functions used: calculateLessonsStartingDate(), getCurrentSemesterName(),
                              getRatesLookupForSemester()
        Utility functions used: UtilityScriptLibrary.debugLog()
================================================================================
END OF FUNCTION DIRECTORY
================================================================================