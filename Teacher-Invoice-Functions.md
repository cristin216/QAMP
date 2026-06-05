================================================================================
TEACHER INVOICE FUNCTION DIRECTORY
================================================================================
    Total Functions: 59
    Most Recent version: 32

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
    addLateTeacherToInvoice
    addStudentLineItem
    addTeacherHeaderRow
    appendTeacherToInvoiceSheet
    buildTeacherInvoiceFileName
    buildTeacherInvoiceVariableMap
    calculateLessonRateAndCost
    calculateLessonsStartingDate
    collectUninvoicedLessonsUpToDate
    convertFolderDocsToPdfUI
    createMonthlyInvoiceSheet
    createStudentReconciliationReport
    createUnpaidLessonsReport
    extractAndMarkLessonsFromSheet
    extractLessonFromRow
    extractLessonsFromAttendanceSheet
    extractTeachersFromFormattedSheet
    formatInvoiceSheet
    generateDefaultInvoicePeriod
    generateDefaultMonthName
    generateInvoiceNumber
    generateMonthlyTeacherInvoices
    generateSingleTeacherInvoice
    generateTeacherInvoiceDocuments
    getActiveTeacherList
    getAttendanceSheetsFromWorkbook
    getRateForSemester
    getRateKeyForProgram
    getRatesLookupForSemester
    getSemesterForDate
    getTeacherContactInfo
    getTeacherInfoByName
    getTeacherInfoFromLessonGroup
    getTeacherInvoicesFolder
    getTeacherInvoicingMetadata
    getUninvoicedLessonsForTeacher
    groupLessonsByTeacherAndType
    isMonthlyInvoiceSheet
    loadProgramRateKeysCache
    loadRatesCache
    normalizeNameForMatching
    onOpen
    parseLessonDate
    parseStudentName
    populateInvoiceSheetFromLessons
    processAttendanceSheet
    processInvoiceSheetForVerification
    promptForCutoffDate
    promptForInvoiceDate
    promptForInvoicePeriod
    promptForMonthName
    showCombinedErrorDetails
    showInvoiceGenerationResults
    showInvoiceGenerationUI
    showTeacherInvoiceResults
    updateMetadataStatus
    updateTeacherInvoiceHistory
    validateLessonData
    verifyAttendanceVsInvoices
    writeTeacherInvoicingMetadata

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addLateTeacherToInvoice() -> void
        UI-triggered function to add a teacher who was missing from the original invoice
        run to the active invoice sheet. Prompts for invoice date, period, and rates
        column, then appends the teacher's lessons.
        Category: INVOICE_GENERATION
        Local functions used: appendTeacherToInvoiceSheet(), getActiveTeacherList(),
                              collectUninvoicedLessonsUpToDate(), groupLessonsByTeacherAndType(),
                              promptForInvoiceDate(), promptForInvoicePeriod(), promptForCutoffDate()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addStudentLineItem(sheet, row, columnMap, lessonGroup, teacherName, ratesCache, programRateKeysCache) -> void
        Writes a single student lesson line item row to the invoice sheet using cached
        rate data. Calculates rate and cost via calculateLessonRateAndCost().
        Category: INVOICE_GENERATION
        Local functions used: calculateLessonRateAndCost()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addTeacherHeaderRow(sheet, row, columnMap, teacherInfo, invoiceDate, invoiceNumber, invoicePeriod) -> void
        Writes a teacher header row to the invoice sheet including teacher name,
        invoice date, invoice number, and invoice period.
        Category: INVOICE_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    appendTeacherToInvoiceSheet(sheet, groupedLessons, invoiceDate, invoicePeriod, ratesColumnHeader) -> void
        Appends all lesson line items for a teacher to the invoice sheet, loading rate
        caches once and writing header and student rows.
        Category: INVOICE_GENERATION
        Local functions used: addTeacherHeaderRow(), addStudentLineItem(),
                              loadRatesCache(), loadProgramRateKeysCache()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    buildTeacherInvoiceFileName(teacherData, variables) -> String
        Builds the file name for a teacher invoice document using teacher name and
        invoice metadata from the variable map.
        Category: INVOICE_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildTeacherInvoiceVariableMap(teacherData, teacherContactInfo, metadata) -> Object
        Builds the template variable map for a teacher invoice document, including
        dynamic description rows, quantities, rates, costs, and total.
        Category: INVOICE_GENERATION
        Local functions used: generateInvoiceNumber()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    calculateLessonRateAndCost(lessonGroup, ratesCache, programRateKeysCache) -> Object
        Calculates the hourly rate and total cost for a lesson group using the rates
        cache and program rate keys cache. Returns an object with rate, cost, and
        rateKey.
        Category: RATES
        Local functions used: getRateKeyForProgram()
        Utility functions used: UtilityScriptLibrary.debugLog()

    calculateLessonsStartingDate(metadataSheet) -> Date | null
        Reads the metadata sheet to determine the lessons starting date for the current
        invoice period based on the previous entry's cutoff date.
        Category: METADATA
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    collectUninvoicedLessonsUpToDate(cutoffDate, invoiceDate, invoicePeriod) -> Object
        Iterates all active teachers and collects uninvoiced lessons up to the cutoff
        date. Returns an object with lessonResults and allGroupedLessons.
        Category: LESSON_COLLECTION
        Local functions used: getActiveTeacherList(), getUninvoicedLessonsForTeacher(),
                              groupLessonsByTeacherAndType()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.formatDateFlexible()

    convertFolderDocsToPdfUI() -> void
        UI-triggered function to convert all Google Docs in the active invoice sheet's
        teacher invoices folder to PDF format.
        Category: DOCUMENT_GENERATION
        Local functions used: getTeacherInvoicesFolder()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createMonthlyInvoiceSheet(month) -> Sheet
        Creates or clears a monthly invoice sheet in the active Teacher Invoices
        workbook with standard column headers.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createStudentReconciliationReport(workbook, attendanceData, invoiceData) -> void
        Creates or overwrites a Student Reconciliation Report sheet in the given
        workbook, comparing attendance records against invoice records by student.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    createUnpaidLessonsReport(workbook, unpaidLessons) -> void
        Creates or overwrites an Unpaid Lessons Report sheet listing all lessons
        that have not yet been invoiced.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    extractAndMarkLessonsFromSheet(sheet, teacherData, cutoffDate, formattedInvoiceDate, invoiceNumber) -> Array
        Extracts uninvoiced lessons from an attendance sheet up to the cutoff date and
        marks them with the invoice date and invoice number in the sheet.
        Category: LESSON_COLLECTION
        Local functions used: extractLessonFromRow()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog(),
                                UtilityScriptLibrary.formatDateFlexible()

    extractLessonFromRow(row, headerMap, teacherData, cutoffDate, sheet, sheetRowIndex) -> Object | null
        Extracts and validates a single lesson row from an attendance sheet. Returns a
        lesson object or null if the row should be skipped.
        Category: LESSON_COLLECTION
        Local functions used: parseLessonDate(), parseStudentName(), normalizeNameForMatching()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.formatDateFlexible()

    extractLessonsFromAttendanceSheet(sheet, teacherData, cutoffDate) -> Array
        Extracts all uninvoiced lessons from an attendance sheet without marking them.
        Used for verification workflows.
        Category: LESSON_COLLECTION
        Local functions used: extractLessonFromRow()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    extractTeachersFromFormattedSheet(sheet) -> Array
        Parses a formatted invoice sheet to extract teacher and student lesson data
        grouped by teacher. Returns an array of teacher data objects.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    formatInvoiceSheet(sheet) -> void
        Applies standard formatting to an invoice sheet: bolds headers, sets column
        widths, applies alternating row colors for teacher sections.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateDefaultInvoicePeriod(cutoffDate) -> String
        Generates a default invoice period string (e.g. "May 2026") based on the
        cutoff date.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames()

    generateDefaultMonthName(cutoffDate) -> String
        Returns the month name string for the given cutoff date.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames()

    generateInvoiceNumber(teacherId, invoiceDate) -> String
        Generates a unique invoice number in the format TXXXX-YYYYMMDD.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateMonthlyTeacherInvoices(month, cutoffDate, invoiceDate, lessonResults) -> Object
        Generates invoice documents for all teachers with lessons in the current
        month. Returns an object with invoiceResults.
        Category: INVOICE_GENERATION
        Local functions used: generateSingleTeacherInvoice(), createMonthlyInvoiceSheet(),
                              populateInvoiceSheetFromLessons(), formatInvoiceSheet()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.formatDateFlexible()

    generateSingleTeacherInvoice(teacherData, sheet, metadata) -> Object
        Generates a single teacher invoice document from the invoice sheet using the
        template variable map. Returns a result object with success status and file info.
        Category: INVOICE_GENERATION
        Local functions used: buildTeacherInvoiceVariableMap(), buildTeacherInvoiceFileName(),
                              getTeacherContactInfo(), updateTeacherInvoiceHistory()
        Utility functions used: UtilityScriptLibrary.generateDocumentFromTemplate(),
                                UtilityScriptLibrary.debugLog()

    generateTeacherInvoiceDocuments() -> void
        UI-triggered function to generate PDF invoice documents from the active invoice
        sheet. Prompts for confirmation and processes each teacher section.
        Category: INVOICE_GENERATION
        Local functions used: extractTeachersFromFormattedSheet(), generateSingleTeacherInvoice(),
                              showTeacherInvoiceResults()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getActiveTeacherList() -> Array
        Returns an array of active teacher objects from the Teacher Roster Lookup sheet,
        each with teacherId, teacherName, and rosterUrl.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getAttendanceSheetsFromWorkbook(spreadsheet) -> Array
        Returns all month-named attendance sheets from a teacher's roster workbook.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.debugLog()

    getRateForSemester(semesterName, rateType, ratesCache) -> Number | null
        Looks up a rate from the rates cache for a given semester and rate type.
        Category: RATES
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRateKeyForProgram(programName, programRateKeysCache) -> String | null
        Looks up the rate key for a program name from the program rate keys cache.
        Category: RATES
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRatesLookupForSemester(semesterName) -> String
        Returns the rates column header string for the given semester by reading the
        billing workbook's Semester Metadata sheet.
        Category: RATES
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getSemesterForDate(date) -> String | null
        Returns the semester name for a given date by reading the billing workbook's
        Semester Metadata sheet. Local version scoped to Teacher-Invoice.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    getTeacherContactInfo(teacherId) -> Object
        Looks up a teacher's contact info (name, email, address) from the Teachers and
        Admin contacts sheet by teacher ID. Returns an object with contact fields.
        Category: DATA_RETRIEVAL
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getTeacherInfoByName(teacherName) -> Object | null
        Looks up a teacher in the Teacher Roster Lookup sheet by display name or full
        name. Returns an object with teacherId, rosterUrl, and name fields, or null.
        Category: DATA_RETRIEVAL
        Local functions used: normalizeNameForMatching()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getTeacherInfoFromLessonGroup(lessonGroup) -> Object | null
        Extracts teacher info from a lesson group object, falling back to a name lookup
        if the teacher ID is not present.
        Category: DATA_RETRIEVAL
        Local functions used: getTeacherInfoByName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getTeacherInvoicesFolder(monthName) -> Folder
        Returns or creates the Teacher Invoices subfolder for the given month under
        the main generated documents folder.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getGeneratedDocumentsFolder(),
                                UtilityScriptLibrary.debugLog()

    getTeacherInvoicingMetadata(invoiceMonth) -> Object | null
        Reads the Teacher Invoicing Metadata sheet and returns the metadata row for
        the given month as an object with date and status fields.
        Category: METADATA
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getUninvoicedLessonsForTeacher(teacherData, cutoffDate, invoiceDate, invoiceNumber) -> Object
        Single-pass function that reads all attendance sheets for a teacher, extracts
        uninvoiced lessons up to the cutoff date, and marks them with the invoice info.
        Returns an object with lessons array and counts.
        Category: LESSON_COLLECTION
        Local functions used: getAttendanceSheetsFromWorkbook(), extractAndMarkLessonsFromSheet()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog(),
                                UtilityScriptLibrary.formatDateFlexible()

    groupLessonsByTeacherAndType(allLessons) -> Object
        Groups an array of lesson objects by teacher and lesson type (program/length).
        Returns a nested object keyed by teacherId and lessonType.
        Category: DATA_MANIPULATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    isMonthlyInvoiceSheet(sheet) -> Boolean
        Returns true if a sheet has the expected monthly invoice sheet column headers.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    loadProgramRateKeysCache() -> Object
        Loads the program-to-rate-key mapping from the billing workbook's Program Rates
        sheet into a cache object keyed by program name.
        Category: RATES
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    loadRatesCache(ratesColumnHeader) -> Object
        Loads lesson rates from the billing workbook's Rates sheet into a cache object
        keyed by semester and rate type.
        Category: RATES
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    normalizeNameForMatching(name) -> String
        Normalizes a name string for fuzzy matching by removing diacritics, lowercasing,
        and stripping extra whitespace.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    onOpen() -> void
        Installs the Teacher Invoice Tools menu in the spreadsheet UI on open.
        Category: UI
        Local functions used: None
        Utility functions used: None

    parseLessonDate(dateValue) -> Date | null
        Parses a lesson date value that may be a Date object or a string in various
        formats. Returns a Date or null if unparseable.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    parseStudentName(studentName) -> Object
        Parses a student name string into firstName and lastName components.
        Returns an object with firstName and lastName.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateInvoiceSheetFromLessons(sheet, groupedLessons, invoiceDate, invoicePeriod, ratesColumnHeader) -> void
        Populates a monthly invoice sheet with all grouped lessons, writing teacher
        header rows and student line items.
        Category: INVOICE_GENERATION
        Local functions used: appendTeacherToInvoiceSheet(), getTeacherInfoFromLessonGroup()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processAttendanceSheet(sheet, teacherLastName, sheetName, unpaidLessons, attendanceData) -> void
        Processes a single attendance sheet for the verification workflow, collecting
        unpaid lesson and attendance data into the provided accumulator arrays.
        Category: VERIFICATION
        Local functions used: parseLessonDate()
        Utility functions used: None

    processInvoiceSheetForVerification(sheet, invoiceData) -> void
        Processes a single invoice sheet for the verification workflow, extracting
        invoice records into the provided accumulator object.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    promptForCutoffDate() -> Date | null
        Prompts the user to enter a cutoff date, defaulting to today. Returns a Date
        or null if cancelled.
        Category: UI
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.promptForCustomToday(),
                                UtilityScriptLibrary.debugLog()

    promptForInvoiceDate() -> Date | null
        Prompts the user to enter an invoice date. Returns a Date or null if cancelled.
        Category: UI
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.promptForDate(), UtilityScriptLibrary.debugLog()

    promptForInvoicePeriod(cutoffDate) -> String | null
        Prompts the user to confirm or enter an invoice period string, defaulting to
        the month of the cutoff date. Returns a string or null if cancelled.
        Category: UI
        Local functions used: generateDefaultInvoicePeriod()
        Utility functions used: UtilityScriptLibrary.debugLog()

    promptForMonthName(cutoffDate) -> String | null
        Prompts the user to confirm or enter a month name, defaulting to the month of
        the cutoff date. Returns a string or null if cancelled.
        Category: UI
        Local functions used: generateDefaultMonthName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    showCombinedErrorDetails(lessonResults, invoiceResults) -> void
        Displays a combined error details dialog from both lesson collection and invoice
        generation result objects.
        Category: UI
        Local functions used: None
        Utility functions used: None

    showInvoiceGenerationResults(lessonResults, invoiceResults) -> void
        Displays a summary dialog of invoice generation results including counts of
        successes, skips, and errors.
        Category: UI
        Local functions used: showCombinedErrorDetails()
        Utility functions used: None

    showInvoiceGenerationUI() -> void
        Main UI entry point for the invoice generation workflow. Prompts for cutoff
        date, invoice date, invoice period, and month name, then runs lesson collection
        and invoice generation.
        Category: UI
        Local functions used: promptForCutoffDate(), promptForInvoiceDate(),
                              promptForInvoicePeriod(), promptForMonthName(),
                              collectUninvoicedLessonsUpToDate(), generateMonthlyTeacherInvoices(),
                              showInvoiceGenerationResults()
        Utility functions used: UtilityScriptLibrary.debugLog()

    showTeacherInvoiceResults(results) -> void
        Displays a results dialog after generating teacher invoice documents, showing
        counts of successes and failures.
        Category: UI
        Local functions used: None
        Utility functions used: None

    updateMetadataStatus(invoiceMonth, newStatus) -> void
        Updates the status field for a given month in the Teacher Invoicing Metadata
        sheet.
        Category: METADATA
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    updateTeacherInvoiceHistory(teacherData, invoiceResult, metadata) -> void
        Logs a successful invoice to the Teacher Invoice History sheet with teacher,
        date, amount, and document link.
        Category: METADATA
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    validateLessonData(groupedLessons) -> Array
        Validates grouped lesson data and returns an array of issue strings for any
        lessons with missing or invalid fields.
        Category: VALIDATION
        Local functions used: None
        Utility functions used: None

    verifyAttendanceVsInvoices() -> void
        UI-triggered verification tool that compares attendance records against invoice
        records across all teachers, generating reconciliation and unpaid lessons reports.
        Category: VERIFICATION
        Local functions used: getActiveTeacherList(), getAttendanceSheetsFromWorkbook(),
                              processAttendanceSheet(), processInvoiceSheetForVerification(),
                              createStudentReconciliationReport(), createUnpaidLessonsReport()
        Utility functions used: UtilityScriptLibrary.getConfig(), UtilityScriptLibrary.debugLog()

    writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod) -> void
        Writes a new row to the Teacher Invoicing Metadata sheet for the given month
        with cutoff date, invoice date, period, rates lookup, semester name, and status.
        Category: METADATA
        Local functions used: calculateLessonsStartingDate(), getRatesLookupForSemester()
        Utility functions used: UtilityScriptLibrary.getCurrentSemesterName(),
                                UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

  ================================================================================
  FUNCTIONS BY CATEGORY
  ================================================================================

  DATA_EXTRACTION (1 function):
    extractTeachersFromFormattedSheet

  DATA_MANIPULATION (1 function):
    groupLessonsByTeacherAndType

  DATA_RETRIEVAL (6 functions):
    getActiveTeacherList
    getAttendanceSheetsFromWorkbook
    getSemesterForDate
    getTeacherContactInfo
    getTeacherInfoByName
    getTeacherInfoFromLessonGroup

  DOCUMENT_GENERATION (2 functions):
    convertFolderDocsToPdfUI
    getTeacherInvoicesFolder

  HELPERS (7 functions):
    generateDefaultInvoicePeriod
    generateDefaultMonthName
    generateInvoiceNumber
    isMonthlyInvoiceSheet
    normalizeNameForMatching
    parseLessonDate
    parseStudentName

  INVOICE_GENERATION (10 functions):
    addLateTeacherToInvoice
    addStudentLineItem
    addTeacherHeaderRow
    appendTeacherToInvoiceSheet
    buildTeacherInvoiceFileName
    buildTeacherInvoiceVariableMap
    generateMonthlyTeacherInvoices
    generateSingleTeacherInvoice
    generateTeacherInvoiceDocuments
    populateInvoiceSheetFromLessons

  LESSON_COLLECTION (5 functions):
    collectUninvoicedLessonsUpToDate
    extractAndMarkLessonsFromSheet
    extractLessonFromRow
    extractLessonsFromAttendanceSheet
    getUninvoicedLessonsForTeacher

  METADATA (5 functions):
    calculateLessonsStartingDate
    getTeacherInvoicingMetadata
    updateMetadataStatus
    updateTeacherInvoiceHistory
    writeTeacherInvoicingMetadata

  RATES (6 functions):
    calculateLessonRateAndCost
    getRateForSemester
    getRateKeyForProgram
    getRatesLookupForSemester
    loadProgramRateKeysCache
    loadRatesCache

  SHEET_OPERATIONS (2 functions):
    createMonthlyInvoiceSheet
    formatInvoiceSheet

  UI (9 functions):
    onOpen
    promptForCutoffDate
    promptForInvoiceDate
    promptForInvoicePeriod
    promptForMonthName
    showCombinedErrorDetails
    showInvoiceGenerationResults
    showInvoiceGenerationUI
    showTeacherInvoiceResults

  VALIDATION (1 function):
    validateLessonData

  VERIFICATION (5 functions):
    createStudentReconciliationReport
    createUnpaidLessonsReport
    processAttendanceSheet
    processInvoiceSheetForVerification
    verifyAttendanceVsInvoices

================================================================================
END OF FUNCTION DIRECTORY
================================================================================