================================================================================
TEACHERINVOICE FUNCTION DIRECTORY
================================================================================
    Total Functions: 85
    Most Recent version: 31

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
addLateTeacherToInvoice,
    addStudentLineItem,
    addTeacherHeaderRow,
    appendTeacherToInvoiceSheet,
    buildTeacherInvoiceFileName,
    buildTeacherInvoiceVariableMap,
    calculateLessonRateAndCost,
    calculateLessonsStartingDate,
    checkTemplateFormatting,
    collectUninvoicedLessonsUpToDate,
    convertFolderDocsToPdfUI,
    createAdminDetailReport,
    createAdminSummaryReport,
    createAugDecVerificationSheet,
    createInvoiceVsBillingSheet,
    createMonthlyInvoiceSheet,
    createStudentHoursByAdminMonthReport,
    createStudentReconciliationReport,
    createUnpaidLessonsReport,
    createVerificationSheet,
    debugMonthlySheetStructure,
    extractAndMarkLessonsFromSheet,
    extractLessonFromRow,
    extractLessonsFromAttendanceSheet,
    extractTeacherInvoiceNumbers,
    extractTeachersFromFormattedSheet,
    formatDateForComparison,
    formatDateForInput,
    formatInvoiceSheet,
    generate2025AdminReports,
    generateAndVerifyHours,
    generateDefaultInvoicePeriod,
    generateDefaultMonthName,
    generateInvoiceNumber,
    generateMonthlyTeacherInvoices,
    generateSingleTeacherInvoice,
    generateTeacherInvoiceDocuments,
    generateYearlyStudentTotals,
    getActiveTeacherList,
    getAttendanceSheetsFromWorkbook,
    getDecemberCumulativeData,
    getRateForSemester,
    getRateKeyForProgram,
    getRatesLookupForSemester,
    getSemesterForDate,
    getTeacherContactInfo,
    getTeacherInfoByName,
    getTeacherInfoFromLessonGroup,
    getTeacherInvoicesFolder,
    getTeacherInvoicingMetadata,
    getUninvoicedLessonsForTeacher,
    groupLessonsByTeacherAndType,
    isMonthlyInvoiceSheet,
    loadProgramRateKeysCache,
    loadRatesCache,
    normalizeNameForMatching,
    onOpen,
    parseLessonDate,
    parseStudentName,
    populateInvoiceSheetFromLessons,
    processBillingSheetData,
    processBillingSheetFor2025,
    processBillingSheetForAugDec,
    processAttendanceSheet,
    processAttendanceSheetForAdminReports,
    processInvoiceSheetForVerification,
    processSheetForMonthlyTotals,
    promptForCutoffDate,
    promptForInvoiceDate,
    promptForInvoicePeriod,
    promptForMonthName,
    setupMetadataStatusDropdown,
    showCombinedErrorDetails,
    showInvoiceGenerationResults,
    showInvoiceGenerationUI,
    showResultsSummaryUI,
    showTeacherInvoiceResults,
    updateMetadataStatus,
    updateTeacherInvoiceHistory,
    validateLessonData,
    verifyAttendanceVsInvoices,
    verifyHoursAugustDecember2025,
    verifyHoursTaught,
    verifyInvoiceVsBilling2025,
    writeTeacherInvoicingMetadata

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addLateTeacherToInvoice() -> void
        UI workflow for adding a late teacher to an already-generated monthly invoice sheet.
        Validates active sheet, loads metadata, prompts user to select a teacher, collects and
        marks their uninvoiced lessons, and appends their rows to the existing sheet.
        Category: UI_WORKFLOW
        Local functions used: isMonthlyInvoiceSheet(), getTeacherInvoicingMetadata(),
                              getActiveTeacherList(), generateInvoiceNumber(),
                              getUninvoicedLessonsForTeacher(), groupLessonsByTeacherAndType(),
                              validateLessonData(), appendTeacherToInvoiceSheet()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

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

    appendTeacherToInvoiceSheet(sheet, groupedLessons, invoiceDate, invoicePeriod) -> Object
        Appends a single teacher's header row and student line items to an existing invoice sheet
        starting after the last used row. Returns object with success status and lineItemCount.
        Category: SHEET_OPERATIONS
        Local functions used: loadRatesCache(), loadProgramRateKeysCache(),
                              getTeacherInfoFromLessonGroup(), addTeacherHeaderRow(), addStudentLineItem()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

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

    checkTemplateFormatting() -> Array
        Debug utility that reads the Monthly Template sheet and analyzes number formats for each column.
        Categorizes formats as CURRENCY, DATE, DECIMAL, INTEGER, GENERAL, or OTHER and displays summary.
        Returns array of column analysis objects.
        Category: DEBUG
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.columnToLetter(), UtilityScriptLibrary.debugLog()

    collectUninvoicedLessonsUpToDate(cutoffDate, invoiceDate, invoicePeriod) -> Object
        Main data collection function. Gathers all uninvoiced lessons from active teachers up to cutoff date.
        Returns comprehensive results object with lessons, errors, validation, summary statistics, and parameters used.
        Category: DATA_COLLECTION
        Local functions used: getActiveTeacherList(), generateInvoiceNumber(), getUninvoicedLessonsForTeacher(),
                              groupLessonsByTeacherAndType(), validateLessonData()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.formatDateFlexible()

    convertFolderDocsToPdfUI() -> void
        Converts all Google Docs in the active month's Teacher Invoices Drive folder to PDFs,
        saves them in a '[Month] PDFs' subfolder, and updates each teacher's roster Invoice Log
        with the PDF URL. Skips already-converted documents.
        Category: UI_WORKFLOW
        Local functions used: isMonthlyInvoiceSheet(), getTeacherInvoicingMetadata(),
                              extractTeachersFromFormattedSheet(), updateTeacherInvoiceHistory()
        Utility functions used: UtilityScriptLibrary.EnvironmentManager.get(), UtilityScriptLibrary.getConfig(),
                                UtilityScriptLibrary.debugLog()

    createAdminDetailReport(workbook, detailErrors) -> void
        Creates or clears 'Admin Detail' sheet and writes row-level validation errors from
        processAttendanceSheetForAdminReports(). Highlights error cells in red.
        Category: REPORTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createAdminSummaryReport(workbook, summaryData) -> void
        Creates or clears 'Admin Summary' sheet and writes per-workbook error counts
        from generate2025AdminReports(). Sorted by workbook name.
        Category: REPORTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createAugDecVerificationSheet(currentWorkbook, ourData, billingData, monthsToVerify) -> void
        Creates or clears 'Aug-Dec 2025 Verification' sheet. Compares calculated invoice hours
        against billing sheet hours for August–December for each student with match indicators.
        Category: REPORTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createInvoiceVsBillingSheet(currentWorkbook, ourData, billingData, monthNames) -> void
        Creates or clears 'Invoice vs Billing 2025' sheet. Compares admin-month invoice hours
        against billing cycle hours for all 12 months with per-student match indicators and totals row.
        Category: REPORTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()
                
    createMonthlyInvoiceSheet(month) -> Sheet
        Creates a new monthly invoice sheet by copying the Monthly Template.
        If sheet already exists, clears data while preserving headers. Returns the sheet object.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createStudentHoursByAdminMonthReport(workbook, studentHours, monthNames) -> void
        Creates or clears the '2025 Hours by Admin Month' sheet in the given workbook.
        Writes a header row with Student ID plus each month name, then populates one row
        per student with their hours per month from the studentHours object.
        Category: REPORTING
        Local functions used: None
        Utility functions used: None
        Called by: generate2025AdminReports()

    createStudentReconciliationReport(workbook, attendanceData, invoiceData) -> void
        Creates or clears the 'Student Reconciliation Report' sheet. Compares attendance
        hours against invoice hours per teacher/invoice date/student, flagging mismatches.
        Category: REPORTING
        Local functions used: None
        Utility functions used: None
        Called by: verifyAttendanceVsInvoices()

    createUnpaidLessonsReport(workbook, unpaidLessons) -> void
        Creates or clears the 'Unpaid Lessons Report' sheet. Writes all lessons found in
        attendance sheets that have no matching invoice entry.
        Category: REPORTING
        Local functions used: None
        Utility functions used: None
        Called by: verifyAttendanceVsInvoices()

    createVerificationSheet(currentWorkbook, ourData, billingData, cleanMonthColumns, decemberCumulativeData) -> void
        Creates or clears the 'Hours Taught Verification' sheet. Compares invoice-derived
        hours against billing sheet hours for each student across months, with Expected/Actual/Match
        column groups per month and December cumulative totals column.
        Category: REPORTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()
        Called by: verifyHoursTaught()

    debugMonthlySheetStructure() -> void
        Debug utility. Reads the active sheet and logs column mapping (teacher, url, lastName,
        firstName) and first few data rows to the debug log for structural analysis.
        Category: DEBUG
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

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

    formatDateForComparison(dateValue) -> String
        Formats a Date object or date-parseable value to MM/DD/YYYY string for string comparison.
        Returns the original toString() value if the date is invalid.
        Category: DATE_FORMATTING
        Local functions used: None
        Utility functions used: None
        Called by: processAttendanceSheet(), processInvoiceSheetForVerification()

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

    generate2025AdminReports() -> void
        UI entry point. Iterates all teacher roster workbooks in the rosters folder,
        processes each attendance sheet via processAttendanceSheetForAdminReports(), and
        produces an Admin Summary and Admin Detail report sheet in the active workbook.
        Category: UI_WORKFLOW
        Local functions used: processAttendanceSheetForAdminReports(),
                              createStudentHoursByAdminMonthReport(), createAdminSummaryReport(),
                              createAdminDetailReport()
        Utility functions used: UtilityScriptLibrary.CONFIG, UtilityScriptLibrary.debugLog()

    generateAndVerifyHours() -> void
        UI entry point. Runs generateYearlyStudentTotals() then verifyHoursTaught() in sequence.
        Displays an error alert if either step fails.
        Category: UI_WORKFLOW
        Local functions used: generateYearlyStudentTotals(), verifyHoursTaught()
        Utility functions used: UtilityScriptLibrary.debugLog()

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

    generateYearlyStudentTotals() -> void
        Reads Teacher Invoicing Metadata to identify all invoice months, then processes
        each monthly invoice sheet via processSheetForMonthlyTotals() to aggregate student
        lesson totals, writing results to the 'Yearly Student Totals' sheet.
        Category: REPORTING
        Local functions used: processSheetForMonthlyTotals()
        Utility functions used: UtilityScriptLibrary.debugLog()
        Called by: generateAndVerifyHours(), verifyHoursTaught()

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

    getDecemberCumulativeData(sheet, decemberCumulativeData) -> void
        Reads a billing sheet and populates the decemberCumulativeData object with each
        student's Current Cumulative Hours Taught value, keyed by Student ID.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: None
        Called by: verifyHoursTaught()

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

    normalizeNameForMatching(name) -> String
        Strips diacritics/accents via NFD normalization, trims whitespace, and lowercases
        a name string for fuzzy comparison.
        Returns normalized string, or empty string if input is falsy.
        Category: DATA_PROCESSING
        Local functions used: None
        Utility functions used: None
        Called by: processAttendanceSheet(), processInvoiceSheetForVerification()

    onOpen() -> void
        Creates custom menu 'Teacher Invoice Tools' in the spreadsheet UI when the Teacher Invoices workbook opens.
        Menu items: Collect Monthly Invoice Data, Add Late Teacher, Generate Invoice Documents,
        Print Documents, Generate & Verify Hours, Verify Logs vs Invoices,
        Generate 2025 Admin Reports, Verify Invoice vs Billing 2025.
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

    processAttendanceSheet(sheet, teacherLastName, sheetName, unpaidLessons, attendanceData) -> void
        Reads an attendance sheet and identifies completed lessons not found in invoice data.
        Appends unpaid lessons to the unpaidLessons array and builds attendanceData map
        keyed by normalized teacher name and invoice date.
        Category: DATA_COLLECTION
        Local functions used: formatDateForComparison(), normalizeNameForMatching()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()
        Called by: verifyAttendanceVsInvoices(), generate2025AdminReports()

    processAttendanceSheetForAdminReports(sheet, workbookName, sheetName, studentHours, monthNames) -> Array
        Processes a single attendance sheet for admin report generation. Aggregates completed
        lesson hours per student per admin month into the studentHours object.
        Returns array of error objects encountered during processing.
        Category: DATA_COLLECTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()
        Called by: generate2025AdminReports()

    processBillingSheetData(sheet, cleanMonth, billingData) -> void
        Reads a billing sheet and populates the billingData object with each student's
        Current Hours Taught This Billing Cycle value for the given cleanMonth key.
        Category: DATA_COLLECTION
        Local functions used: None
        Utility functions used: None
        Called by: verifyHoursTaught()

    processBillingSheetFor2025(sheet, month, billingData) -> void
        Reads a 2025 billing cycle sheet and populates billingData with student hours
        for the given month. Variant of processBillingSheetData scoped to 2025 data.
        Category: DATA_COLLECTION
        Local functions used: None
        Utility functions used: None
        Called by: verifyInvoiceVsBilling2025()

    processBillingSheetForAugDec(sheet, month, billingData) -> void
        Reads an August–December billing sheet and populates billingData with student hours
        for the given month. Variant scoped to the Aug–Dec date range.
        Category: DATA_COLLECTION
        Local functions used: None
        Utility functions used: None
        Called by: verifyHoursAugustDecember2025()

    processInvoiceSheetForVerification(sheet, invoiceData) -> void
        Reads a monthly invoice sheet and builds the invoiceData map keyed by normalized
        teacher name and invoice date, storing student lesson durations and quantities
        for attendance vs invoice comparison.
        Category: DATA_EXTRACTION
        Local functions used: formatDateForComparison(), normalizeNameForMatching()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()
        Called by: verifyAttendanceVsInvoices()

    processSheetForMonthlyTotals(sheet, cleanMonth, studentData) -> void
        Reads a monthly invoice sheet and aggregates each student's lesson hours for
        the given cleanMonth into the studentData object, accumulating totals across calls.
        Category: DATA_COLLECTION
        Local functions used: None
        Utility functions used: None
        Called by: generateYearlyStudentTotals()

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

    verifyAttendanceVsInvoices() -> void
        UI entry point. Prompts for a year, scans all teacher roster workbooks for that year,
        compares attendance sheet data against invoice sheet data via processAttendanceSheet()
        and processInvoiceSheetForVerification(), then generates a Student Reconciliation Report
        and an Unpaid Lessons Report.
        Category: UI_WORKFLOW
        Local functions used: processAttendanceSheet(), processInvoiceSheetForVerification(),
                              createStudentReconciliationReport(), createUnpaidLessonsReport()
        Utility functions used: UtilityScriptLibrary.CONFIG, UtilityScriptLibrary.debugLog()

    verifyHoursAugustDecember2025() -> void
        Reads the 'Yearly Student Totals' sheet, processes August–December billing sheets via
        processBillingSheetForAugDec(), and produces the 'Aug-Dec 2025 Verification' sheet
        via createAugDecVerificationSheet() comparing invoice hours against billing hours.
        Category: UI_WORKFLOW
        Local functions used: processBillingSheetForAugDec(), createAugDecVerificationSheet()
        Utility functions used: UtilityScriptLibrary.CONFIG, UtilityScriptLibrary.debugLog()
        Note: Defined in JS without a leading newline separator (}function) — formatting issue.

    verifyHoursTaught() -> void
        Reads Teacher Invoicing Metadata to enumerate invoice months, processes corresponding
        billing sheets via processBillingSheetData() and getDecemberCumulativeData(), then
        calls generateYearlyStudentTotals() and createVerificationSheet() to produce the
        'Hours Taught Verification' comparison sheet.
        Category: UI_WORKFLOW
        Local functions used: generateYearlyStudentTotals(), processBillingSheetData(),
                              getDecemberCumulativeData(), createVerificationSheet()
        Utility functions used: UtilityScriptLibrary.CONFIG, UtilityScriptLibrary.debugLog()
        Called by: generateAndVerifyHours()

    verifyInvoiceVsBilling2025() -> void
        Reads the '2025 Hours by Admin Month' sheet and iterates billing sheets to compare
        invoice-derived hours against billing cycle hours for each student, producing the
        'Invoice vs Billing 2025' sheet via createInvoiceVsBillingSheet().
        Category: UI_WORKFLOW
        Local functions used: processBillingSheetFor2025(), createInvoiceVsBillingSheet()
        Utility functions used: UtilityScriptLibrary.CONFIG, UtilityScriptLibrary.debugLog()

    writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod) -> Object
        Writes new invoice metadata record to Teacher Invoicing Metadata sheet including date ranges,
        semester info, and rates lookup. Returns success object with lessonsStartingDate and
        lessonsEndingDate.
        Category: METADATA_MANAGEMENT
        Local functions used: calculateLessonsStartingDate(), getRatesLookupForSemester()
        Utility functions used: UtilityScriptLibrary.getCurrentSemesterName(), UtilityScriptLibrary.debugLog()
================================================================================
END OF FUNCTION DIRECTORY
================================================================================