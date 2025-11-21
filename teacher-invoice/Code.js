/* 
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
*/
       
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Teacher Invoice Tools')
      .addItem('Collect Monthly Invoice Data', 'showInvoiceGenerationUI')
      .addItem('Generate Invoice Documents', 'generateTeacherInvoiceDocuments')
      .addToUi();
    
    UtilityScriptLibrary.debugLog("✅ Teacher Invoice Tools menu created");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error creating menu: " + error.message);
  }
}

// === MAIN UI FUNCTION ===


function showInvoiceGenerationUI() {
  try {
    UtilityScriptLibrary.debugLog("showInvoiceGenerationUI - Starting invoice generation UI");
    
    var ui = SpreadsheetApp.getUi();
    
    // Step 1: Get cutoff date first
    var cutoffDate = promptForCutoffDate();
    if (!cutoffDate) {
      ui.alert('❌ Cutoff date is required for invoice generation.');
      return;
    }
    
    // Step 2: Get invoice date
    var invoiceDate = promptForInvoiceDate();
    if (!invoiceDate) {
      ui.alert('❌ Invoice date is required for invoice generation.');
      return;
    }
    
    // Step 3: Get month name for sheet (with default based on cutoff date)
    var month = promptForMonthName(cutoffDate);
    if (!month) {
      ui.alert('❌ Month name is required for invoice generation.');
      return;
    }
    
    // Step 4: Get invoice period
    var invoicePeriod = promptForInvoicePeriod(cutoffDate);
    if (!invoicePeriod) {
      ui.alert('❌ Invoice period is required for invoice generation.');
      return;
    }
    
    // Step 5: Collect lesson data using parameters (no additional prompts)
    var lessonResults = collectUninvoicedLessonsUpToDate(cutoffDate, invoiceDate, invoicePeriod);
    
    if (!lessonResults) {
      ui.alert('❌ Lesson data collection failed.');
      return;
    }
    
    // Step 6: Generate invoice sheet
    var invoiceResults = generateMonthlyTeacherInvoices(month, cutoffDate, invoiceDate, lessonResults);
    
    if (!invoiceResults.success) {
      ui.alert('❌ Invoice generation failed: ' + invoiceResults.message);
      return;
    }
    
    // Step 7: Write metadata
    try {
      writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod);
      UtilityScriptLibrary.debugLog("showInvoiceGenerationUI", "SUCCESS", 
                                    "Metadata written successfully", "", "");
    } catch (metadataError) {
      UtilityScriptLibrary.debugLog("showInvoiceGenerationUI", "ERROR", 
                                    "Metadata write failed (non-critical)", 
                                    "", metadataError.message);
      // Continue - this is not critical to invoice generation
    }
    
    // Step 8: Show comprehensive results summary
    showInvoiceGenerationResults(lessonResults, invoiceResults);
    
    UtilityScriptLibrary.debugLog("showInvoiceGenerationUI - Invoice generation workflow completed successfully");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showInvoiceGenerationUI - ERROR: Invoice generation workflow failed - " + error.message);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('❌ Error: ' + error.message + '\n\nCheck the Teacher_Invoice_Debug sheet for details.');
  }
}

function showLessonCollectionUI() {
  try {
    UtilityScriptLibrary.debugLog("showLessonCollectionUI - Starting lesson collection UI");
    
    var ui = SpreadsheetApp.getUi();
    
    // Collect the data using the improved prompts
    var results = collectUninvoicedLessonsUpToDate();
    
    if (!results) {
      ui.alert('❌ Data collection was cancelled or failed.');
      return;
    }
    
    // Show comprehensive results summary
    showResultsSummaryUI(results);
    
    UtilityScriptLibrary.debugLog("showLessonCollectionUI - UI workflow completed successfully");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showLessonCollectionUI - ERROR: UI workflow failed - " + error.message);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('❌ Error: ' + error.message + '\n\nCheck the Teacher_Invoice_Debug sheet for details.');
  }
}

function showResultsSummaryUI(results) {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Build summary message
    var summaryMessage = '📊 LESSON COLLECTION RESULTS\n\n';
    
    // Basic info
    summaryMessage += '📅 Cutoff Date: ' + results.cutoffDate.toDateString() + '\n';
    summaryMessage += '📝 Invoice Date: ' + results.invoiceDate.toDateString() + '\n';
    summaryMessage += '📋 Invoice Period: ' + results.invoicePeriod + '\n\n';
    
    // Summary stats
    summaryMessage += '📈 SUMMARY:\n';
    summaryMessage += '• Teachers checked: ' + results.summary.totalTeachers + '\n';
    summaryMessage += '• Successful: ' + results.summary.successfulTeachers + '\n';
    summaryMessage += '• Total lessons found: ' + results.summary.totalLessons + '\n';
    summaryMessage += '• Invoice line items: ' + results.summary.totalLineItems + '\n';
    summaryMessage += '• Errors: ' + results.errors.length + '\n';
    summaryMessage += '• Validation issues: ' + results.validation.issues.length + '\n\n';
    
    // Status
    if (results.errors.length === 0 && results.validation.issues.length === 0) {
      summaryMessage += '✅ SUCCESS: All data collected successfully!\n';
    } else if (results.errors.length > 0 || results.validation.issues.length > 0) {
      summaryMessage += '⚠️ PARTIAL SUCCESS: Some issues found.\n';
    }
    
    summaryMessage += '\nCheck Teacher_Invoice_Debug sheet for detailed logs.';
    
    // Show the summary
    ui.alert('Data Collection Complete', summaryMessage, ui.ButtonSet.OK);
    
    UtilityScriptLibrary.debugLog("showResultsSummaryUI - Results summary displayed to user");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showResultsSummaryUI - ERROR: Error showing results summary - " + error.message);
  }
}

// === GENERAL FUNCTIONS ===


function addStudentLineItem(sheet, row, columnMap, lessonGroup, teacherName, ratesCache, programRateKeysCache) {
  try {
    // Calculate rate and cost for this lesson type using cached rates
    var rateCalculation = calculateLessonRateAndCost(lessonGroup, ratesCache, programRateKeysCache);
    
    // Parse student name into first and last names
    var studentNames = parseStudentName(lessonGroup.studentName);
    
    var rowData = new Array(sheet.getLastColumn()).fill('');
    
    // Populate student line item row using normalized column names (convert to 0-based)
    if (columnMap['lastname'] !== undefined) rowData[columnMap['lastname'] - 1] = studentNames.lastName;
    if (columnMap['firstname'] !== undefined) rowData[columnMap['firstname'] - 1] = studentNames.firstName;
    if (columnMap['id'] !== undefined) rowData[columnMap['id'] - 1] = lessonGroup.studentId;
    if (columnMap['teacher'] !== undefined) rowData[columnMap['teacher'] - 1] = teacherName;
    if (columnMap['duration'] !== undefined) rowData[columnMap['duration'] - 1] = lessonGroup.lessonLength;
    if (columnMap['quantity'] !== undefined) rowData[columnMap['quantity'] - 1] = lessonGroup.quantity;
    if (columnMap['rate'] !== undefined) rowData[columnMap['rate'] - 1] = rateCalculation.lessonRate;
    if (columnMap['cost'] !== undefined) rowData[columnMap['cost'] - 1] = rateCalculation.totalCost;
    // Leave other columns blank for student rows
    
    // Add the row data
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    
    // Apply attendance-style formatting
    var dataRange = sheet.getRange(row, 1, 1, rowData.length);
    dataRange.setFontWeight('normal')
             .setFontColor('black')
             .setWrap(true);
    
    // Apply alternating background
    var isEvenRow = row % 2 === 0;
    if (isEvenRow) {
      dataRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background);
    } else {
      dataRange.setBackground(UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background);
    }
    
    UtilityScriptLibrary.debugLog("addStudentLineItem", "DEBUG", 
                                  "Added student line item with 0-based indices", 
                                  "Student: " + lessonGroup.studentName + ", Row: " + row + ", Rate: " + rateCalculation.lessonRate + ", Cost: " + rateCalculation.totalCost, "");
    
    return row + 1;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addStudentLineItem", "ERROR", 
                                  "Failed to add student line item", 
                                  "Student: " + lessonGroup.studentName, error.message);
    throw error;
  }
}

function addTeacherHeaderRow(sheet, row, columnMap, teacherInfo, invoiceDate, invoiceNumber, invoicePeriod) {
  try {
    var rowData = new Array(sheet.getLastColumn()).fill('');
    
    // Populate teacher header row using normalized column names (convert to 0-based)
    if (columnMap['lastname'] !== undefined) rowData[columnMap['lastname'] - 1] = teacherInfo.lastName || '';
    if (columnMap['firstname'] !== undefined) rowData[columnMap['firstname'] - 1] = teacherInfo.firstName || '';
    if (columnMap['id'] !== undefined) rowData[columnMap['id'] - 1] = teacherInfo.teacherId || '';
    if (columnMap['teacher'] !== undefined) rowData[columnMap['teacher'] - 1] = teacherInfo.teacherName || '';
    if (columnMap['invoicedate'] !== undefined) rowData[columnMap['invoicedate'] - 1] = UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MM/dd/yyyy');
    if (columnMap['invoicenumber'] !== undefined) rowData[columnMap['invoicenumber'] - 1] = invoiceNumber || '';
    if (columnMap['invoiceperiod'] !== undefined) rowData[columnMap['invoiceperiod'] - 1] = invoicePeriod || '';
    if (columnMap['address'] !== undefined) rowData[columnMap['address'] - 1] = teacherInfo.address || '';
    
   
    
    // Add the row data
    sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
    
    // Apply teacher header formatting
    var dataRange = sheet.getRange(row, 1, 1, rowData.length);
    dataRange.setFontWeight('bold')
             .setFontColor(UtilityScriptLibrary.STYLES.SUBHEADER.text)
             .setWrap(true)
             .setBackground(UtilityScriptLibrary.STYLES.SUBHEADER.background);
    
    UtilityScriptLibrary.debugLog("addTeacherHeaderRow", "DEBUG", 
                                  "Added teacher header row with normalized columns", 
                                  "Teacher: " + teacherInfo.teacherName + ", Row: " + row, "");
    
    return row + 1;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addTeacherHeaderRow", "ERROR", "Failed to add teacher header row", 
                                  "Teacher: " + (teacherInfo.teacherName || 'unknown'), error.message);
    throw error;
  }
}

function buildTeacherInvoiceFileName(teacherData, variables) {
  try {
    var teacherName = variables.TeacherFirstName && variables.TeacherLastName ? 
                      variables.TeacherFirstName + ' ' + variables.TeacherLastName : 
                      teacherData.teacherName;
    
    var invoiceNumber = variables.InvoiceNumber || 'NOINVOICE';
    
    return teacherName + ' - ' + invoiceNumber;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("buildTeacherInvoiceFileName", "ERROR", "Failed to build filename", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    return 'TeacherInvoice - ' + teacherData.teacherName;
  }
}

function buildTeacherInvoiceVariableMap(teacherData, teacherContactInfo, metadata) {
  try {
    // Build dynamic content directly from the sheet's student rows
    var dynamicDescription = [];
    var dynQty = [];
    var dynRate = [];
    var dynCost = [];
    var totalCost = 0;
    
    for (var i = 0; i < teacherData.studentRows.length; i++) {
      var row = teacherData.studentRows[i];
      
      // Format: "Last, First [Duration] minute lessons" (matches existing sheet format)
      var description = row.lastName + ', ' + row.firstName + ' ' + row.duration + ' minute lessons';
      dynamicDescription.push(description);
      
      dynQty.push(row.quantity.toString());
      dynRate.push(UtilityScriptLibrary.formatCurrency(row.rate));
      dynCost.push(UtilityScriptLibrary.formatCurrency(row.cost));
      
      totalCost += row.cost;
    }
    
    // Use teacher's first and last name directly from the sheet
    var teacherFirstName = teacherData.teacherFirstName || '';
    var teacherLastName = teacherData.teacherLastName || '';
    var teacherAddress = teacherData.teacherAddress || '';
    
    // Only fall back to parsing teacherName if we don't have the names from sheet
    if (!teacherFirstName && !teacherLastName && teacherData.teacherName) {
      var nameParts = teacherData.teacherName.split(' ');
      if (nameParts.length >= 2) {
        teacherFirstName = nameParts[0];
        teacherLastName = nameParts.slice(1).join(' ');
      } else {
        teacherLastName = teacherData.teacherName; // Use full name as last name
      }
    }
    
    // Format invoice date
    var formattedInvoiceDate = '';
    if (teacherData.invoiceDate) {
      var invoiceDate = typeof teacherData.invoiceDate === 'string' ?
        new Date(teacherData.invoiceDate) : teacherData.invoiceDate;
      formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MMMM d, yyyy');
    }
    
    // Format invoice period - USE METADATA FOR DATE RANGE
    var formattedInvoicePeriod = '';
    if (metadata && metadata.lessonsStartingDate && metadata.lessonsEndingDate) {
      // Use metadata date range
      var startDate = new Date(metadata.lessonsStartingDate);
      var endDate = new Date(metadata.lessonsEndingDate);
      
      var formattedStart = UtilityScriptLibrary.formatDateFlexible(startDate, "MMMM d, yyyy");
      var formattedEnd = UtilityScriptLibrary.formatDateFlexible(endDate, "MMMM d, yyyy");
      formattedInvoicePeriod = formattedStart + " - " + formattedEnd;
      
      UtilityScriptLibrary.debugLog("buildTeacherInvoiceVariableMap", "DEBUG", 
                                    "Using metadata for invoice period", 
                                    "Period: " + formattedInvoicePeriod, "");
    } else {
      // Fallback to teacherData (old behavior)
      if (teacherData.invoicePeriod) {
        if (typeof teacherData.invoicePeriod === 'string') {
          formattedInvoicePeriod = teacherData.invoicePeriod; // Already a string
        } else {
          // It's a Date object, format it as "MMMM yyyy"
          formattedInvoicePeriod = UtilityScriptLibrary.formatDateFlexible(teacherData.invoicePeriod, 'MMMM yyyy');
        }
      }
      
      UtilityScriptLibrary.debugLog("buildTeacherInvoiceVariableMap", "WARNING", 
                                    "Using fallback for invoice period (no metadata)", 
                                    "Period: " + formattedInvoicePeriod, "");
    }
    
    return {
      'InvoiceNumber': teacherData.invoiceNumber || '',
      'InvoiceDate': formattedInvoiceDate,
      'InvoicePeriod': formattedInvoicePeriod,
      'TeacherFirstName': teacherFirstName,
      'TeacherLastName': teacherLastName,
      'TeacherAddress': teacherAddress,
      'DYNAMIC_DESCRIPTION': dynamicDescription.join('\n'),
      'DYN_QTY': dynQty.join('\n'),
      'DYN_RATE': dynRate.join('\n'),
      'DYN_COST': dynCost.join('\n'),
      'Total': UtilityScriptLibrary.formatCurrency(totalCost),
      'Comment': teacherData.comment || ''
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("buildTeacherInvoiceVariableMap", "ERROR", "Failed to build variable map", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    throw error;
  }
}

function calculateLessonRateAndCost(lessonGroup, ratesCache, programRateKeysCache) {
  try {
    // Debug logging to track the issue
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                  "Starting rate calculation", 
                                  "Student/Group: " + lessonGroup.studentName + ", ID: " + lessonGroup.studentId + ", Length: " + lessonGroup.lessonLength + ", Quantity: " + lessonGroup.quantity, "");
    
    // Get the semester for rate lookup using the first lesson date
    if (!lessonGroup.lessons || lessonGroup.lessons.length === 0) {
      throw new Error("No lessons found in lesson group");
    }
    
    var firstLessonDate = lessonGroup.lessons[0].lessonDate;
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                  "First lesson date", 
                                  "Date: " + UtilityScriptLibrary.formatDateFlexible(firstLessonDate, 'MMM d, yyyy'), "");
    
    var semesterName = getSemesterForDate(firstLessonDate);
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                  "Semester determined", 
                                  "Semester: " + semesterName, "");
    
    // Determine if this is a group session (ID starts with "G")
    var isGroupSession = lessonGroup.studentId && lessonGroup.studentId.toString().charAt(0) === 'G';
    
    var hourlyRate;
    var lessonRate;
    var totalCost;
    
    if (isGroupSession) {
      // GROUP SESSION RATE CALCULATION
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Detected group session", 
                                    "Group ID: " + lessonGroup.studentId, "");
      
      // Get the Rate Key from cache
      var rateKey = getRateKeyForProgram(lessonGroup.studentName, programRateKeysCache);
      
      if (!rateKey) {
        throw new Error("Could not find Rate Key for group: " + lessonGroup.studentName);
      }
      
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Rate Key found", 
                                    "Rate Key: " + rateKey, "");
      
      // Look up group pay rate using Rate Key + " Pay" from cache
      var rateType = rateKey + ' Pay';
      hourlyRate = getRateForSemester(semesterName, rateType, ratesCache);
      
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Group hourly rate retrieved", 
                                    "Rate Type: " + rateType + ", Rate: " + hourlyRate, "");
      
      // Calculate lesson rate based on duration (if provided)
      if (lessonGroup.lessonLength && lessonGroup.lessonLength > 0) {
        var lessonRateMultiplier = lessonGroup.lessonLength / 60; // Convert minutes to hours
        lessonRate = hourlyRate * lessonRateMultiplier;
      } else {
        // If no duration specified, use hourly rate as-is
        lessonRate = hourlyRate;
      }
      
      // Calculate total cost
      totalCost = lessonRate * lessonGroup.quantity;
      
    } else {
      // STUDENT LESSON RATE CALCULATION
      hourlyRate = getRateForSemester(semesterName, 'Lessons Pay', ratesCache);
      UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "DEBUG", 
                                    "Student hourly rate retrieved", 
                                    "Rate: " + hourlyRate, "");
      
      // Calculate lesson rate based on duration
      var lessonRateMultiplier = lessonGroup.lessonLength / 60; // Convert minutes to hours
      lessonRate = hourlyRate * lessonRateMultiplier;
      
      // Calculate total cost
      totalCost = lessonRate * lessonGroup.quantity;
    }
    
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "INFO", 
                                  "Rate calculation completed", 
                                  "Lesson rate: " + lessonRate + ", Total cost: " + totalCost, "");
    
    return {
      hourlyRate: hourlyRate,
      lessonRate: lessonRate,
      totalCost: totalCost
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("calculateLessonRateAndCost", "ERROR", 
                                  "Rate calculation failed", 
                                  "Student/Group: " + lessonGroup.studentName, error.message);
    throw error;
  }
}

function calculateLessonsStartingDate(metadataSheet) {
  try {
    var lastRow = metadataSheet.getLastRow();
    
    // If only headers exist (row 1), use first day of cutoff date's month
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "INFO", 
                                    "First invoice ever - no previous metadata", 
                                    "Will use first day of invoice month", "");
      return null; // Return null to signal that we should calculate from cutoff date
    }
    
    // Get the previous row's "Lessons Ending Date" 
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    var endingDateCol = headerMap[UtilityScriptLibrary.normalizeHeader("Lessons Ending Date")];
    
    if (!endingDateCol) {
      throw new Error("Lessons Ending Date column not found in metadata");
    }
    
    var previousEndingDate = metadataSheet.getRange(lastRow, endingDateCol).getValue();
    
    if (!previousEndingDate || !(previousEndingDate instanceof Date)) {
      UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "WARNING", 
                                    "Invalid previous ending date", 
                                    "Value: " + previousEndingDate, "");
      return null; // Return null to signal that we should calculate from cutoff date
    }
    
    // Add 1 day to previous ending date
    var startingDate = new Date(previousEndingDate);
    startingDate.setDate(startingDate.getDate() + 1);
    
    UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "INFO", 
                                  "Calculated starting date", 
                                  "Previous ending: " + previousEndingDate.toDateString() + 
                                  ", New starting: " + startingDate.toDateString(), "");
    
    return startingDate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("calculateLessonsStartingDate", "ERROR", 
                                  "Failed to calculate starting date", 
                                  "", error.message);
    return null; // Return null to fallback to calculation from cutoff date
  }
}

function collectUninvoicedLessonsUpToDate(cutoffDate, invoiceDate, invoicePeriod) {
  try {
    // If no parameters provided, prompt user (backward compatibility)
    if (!cutoffDate) {
      cutoffDate = promptForCutoffDate();
      if (!cutoffDate) {
        throw new Error('No cutoff date provided');
      }
    }
    
    if (!invoiceDate) {
      invoiceDate = promptForInvoiceDate();
      if (!invoiceDate) {
        throw new Error('No invoice date provided');
      }
    }
    
    if (!invoicePeriod) {
      invoicePeriod = promptForInvoicePeriod(cutoffDate);
      if (!invoicePeriod) {
        throw new Error('No invoice period provided');
      }
    }
    
    UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'INFO', 
                                  'Starting lesson collection', 
                                  'Cutoff: ' + cutoffDate.toDateString() + ', Invoice Date: ' + invoiceDate.toDateString(), '');
    
    // Get Form Responses workbook ONCE to get teacher list
    var formResponsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
    var lookupSheet = formResponsesSS.getSheetByName('Teacher Roster Lookup');
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      throw new Error('Teacher Roster Lookup sheet not found or empty');
    }
    
    // Read all teachers into memory
    var data = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 6).getValues();
    var allTeachers = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var teacherName = row[0];      // Column A: Teacher Name
      var rosterUrl = row[1];        // Column B: Roster URL
      var teacherId = row[2];        // Column C: Teacher ID
      var displayName = row[3];      // Column D: Display Name
      var status = row[4];           // Column E: Status
      
      // Only include teachers that have roster URLs
      if (teacherName && rosterUrl && rosterUrl.toString().trim() !== '') {
        // Generate invoice number NOW (before processing teacher)
        var invoiceNumber = generateInvoiceNumber(teacherId || teacherName, invoiceDate);
        
        allTeachers.push({
          teacherName: teacherName,
          rosterUrl: rosterUrl,
          teacherId: teacherId || '',
          displayName: displayName || '',
          status: status || 'Unknown',
          invoiceNumber: invoiceNumber  // Store invoice number with teacher data
        });
      }
    }
    
    UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'INFO', 
                                  'Found teachers', 
                                  'Count: ' + allTeachers.length, '');
    
    // Collect lessons from each teacher (AND mark them as invoiced in same pass)
    var allUninvoicedLessons = [];
    var errors = [];
    
    for (var i = 0; i < allTeachers.length; i++) {
      var teacher = allTeachers[i];
      try {
        // SINGLE PASS: Collect AND mark lessons
        var teacherLessons = getUninvoicedLessonsForTeacher(
          teacher, 
          cutoffDate, 
          invoiceDate, 
          teacher.invoiceNumber  // Pass invoice number
        );
        
        if (teacherLessons.length > 0) {
          allUninvoicedLessons = allUninvoicedLessons.concat(teacherLessons);
          UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'SUCCESS', 
                                        'Teacher processed', 
                                        teacher.teacherName + ': ' + teacherLessons.length + ' lessons', '');
        }
      } catch (error) {
        errors.push({
          teacher: teacher.teacherName,
          error: error.message,
          timestamp: new Date()
        });
        UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'ERROR', 
                                      'Teacher failed', 
                                      teacher.teacherName, error.message);
      }
    }
    
    // Group lessons by teacher and student+length
    var groupedLessons = groupLessonsByTeacherAndType(allUninvoicedLessons);
    
    // Basic validation
    var validationResults = validateLessonData(groupedLessons);
    
    return {
      lessons: groupedLessons,
      allLessons: allUninvoicedLessons,  // Keep flat list for reference
      errors: errors,
      validation: validationResults,
      cutoffDate: cutoffDate,
      invoiceDate: invoiceDate,
      invoicePeriod: invoicePeriod,
      summary: {
        totalTeachers: allTeachers.length,
        successfulTeachers: allTeachers.length - errors.length,
        totalLessons: allUninvoicedLessons.length,
        totalLineItems: Object.keys(groupedLessons).length
      }
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('collectUninvoicedLessonsUpToDate', 'ERROR', 
                                  'Lesson collection failed', '', error.message);
    throw error;
  }
}

function createMonthlyInvoiceSheet(month) {
  try {
    // Access active Teacher Invoices workbook
    var teacherInvoicesSS = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if month sheet already exists
    var monthSheet = teacherInvoicesSS.getSheetByName(month);
    if (monthSheet) {
      UtilityScriptLibrary.debugLog("createMonthlyInvoiceSheet", "WARNING", 
                                    "Sheet already exists", "Month: " + month, "");
      
      // Clear existing data but keep headers
      if (monthSheet.getLastRow() > 1) {
        monthSheet.getRange(2, 1, monthSheet.getLastRow() - 1, monthSheet.getLastColumn()).clearContent();
      }
      return monthSheet;
    }
    
    // Get the Monthly Template sheet
    var templateSheet = teacherInvoicesSS.getSheetByName("Monthly Template");
    if (!templateSheet) {
      throw new Error("Monthly Template sheet not found in Teacher Invoices workbook");
    }
    
    // Copy template to create new month sheet
    monthSheet = templateSheet.copyTo(teacherInvoicesSS);
    monthSheet.setName(month);
    
    UtilityScriptLibrary.debugLog("createMonthlyInvoiceSheet", "INFO", 
                                  "Created month sheet from template", "Month: " + month, "");
    
    return monthSheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("createMonthlyInvoiceSheet", "ERROR", 
                                  "Failed to create month sheet", "Month: " + month, error.message);
    throw error;
  }
}

function extractAndMarkLessonsFromSheet(sheet, teacherData, cutoffDate, formattedInvoiceDate, invoiceNumber) {
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // Empty sheet or only headers
    }
    
    // Use UtilityScriptLibrary to get header map (returns 1-based column numbers)
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Validate required columns exist
    var requiredColumns = ['studentid', 'studentname', 'date', 'length', 'status'];
    for (var j = 0; j < requiredColumns.length; j++) {
      var col = requiredColumns[j];
      if (!headerMap[col]) {
        throw new Error('Required column "' + col + '" not found in ' + sheet.getName());
      }
    }
    
    // Get column indices (convert to 0-based for array access, but keep 1-based for setRange)
    var studentIdCol = headerMap['studentid'];
    var studentNameCol = headerMap['studentname'];
    var dateCol = headerMap['date'];
    var lengthCol = headerMap['length'];
    var statusCol = headerMap['status'];
    var invoiceDateCol = headerMap['invoicedate'];
    var invoiceNumberCol = headerMap['invoicenumber'];
    
    var lessons = [];
    var rowsToUpdate = []; // Track which rows need invoice data written
    
    // Process data rows (skip header row 0)
    for (var k = 1; k < data.length; k++) {
      var row = data[k];
      var sheetRowIndex = k + 1; // Convert 0-based to 1-based sheet row
      
      // Extract lesson data using 0-based indices
      var studentId = row[studentIdCol - 1];
      var studentName = row[studentNameCol - 1];
      var dateValue = row[dateCol - 1];
      var lengthValue = row[lengthCol - 1];
      var statusValue = row[statusCol - 1];
      var existingInvoiceDate = invoiceDateCol ? row[invoiceDateCol - 1] : null;
      
      // Skip if already invoiced
      if (existingInvoiceDate && existingInvoiceDate.toString().trim() !== '') {
        continue;
      }
      
      // Skip if no student ID or student name
      if (!studentId || !studentName || studentId.toString().trim() === '' || studentName.toString().trim() === '') {
        continue;
      }
      
      // Parse lesson date
      var lessonDate = parseLessonDate(dateValue);
      if (!lessonDate || isNaN(lessonDate.getTime())) {
        continue;
      }
      
      // Skip if after cutoff date
      if (lessonDate > cutoffDate) {
        continue;
      }
      
      // Parse status
      var status = statusValue ? statusValue.toString().toLowerCase().trim() : '';
      
      // Skip invalid statuses
      if (status !== 'lesson' && status !== 'no show') {
        continue;
      }
      
      // Parse lesson length
      var lessonLength = parseInt(lengthValue) || 30;
      
      // This is a valid uninvoiced lesson - collect it
      var lesson = {
        teacherName: teacherData.teacherName,
        teacherId: teacherData.teacherId,
        studentId: studentId.toString().trim(),
        studentName: studentName.toString().trim(),
        lessonDate: lessonDate,
        status: status === 'no show' ? 'No Show' : 'Lesson',
        lessonLength: lessonLength,
        sheetName: sheet.getName(),
        sheetRow: sheetRowIndex
      };
      
      lessons.push(lesson);
      
      // Track that this row needs invoice data written
      rowsToUpdate.push({
        row: sheetRowIndex,
        studentId: lesson.studentId,
        lessonDate: lesson.lessonDate
      });
    }
    
    // WRITE INVOICE DATA IMMEDIATELY (while workbook is still open)
    if (rowsToUpdate.length > 0 && invoiceDateCol && invoiceNumberCol) {
      UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "INFO", 
                                    "Writing invoice data to sheet", 
                                    "Sheet: " + sheet.getName() + ", Rows: " + rowsToUpdate.length, "");
      
      for (var m = 0; m < rowsToUpdate.length; m++) {
        var rowInfo = rowsToUpdate[m];
        
        // Write Invoice Date
        sheet.getRange(rowInfo.row, invoiceDateCol).setValue(formattedInvoiceDate);
        
        // Write Invoice Number
        sheet.getRange(rowInfo.row, invoiceNumberCol).setValue(invoiceNumber);
      }
      
      UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "SUCCESS", 
                                    "Invoice data written", 
                                    "Sheet: " + sheet.getName() + ", Marked: " + rowsToUpdate.length, "");
    } else if (rowsToUpdate.length > 0 && (!invoiceDateCol || !invoiceNumberCol)) {
      UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "WARNING", 
                                    "Cannot write invoice data - columns missing", 
                                    "Sheet: " + sheet.getName() + ", InvoiceDateCol: " + invoiceDateCol + ", InvoiceNumberCol: " + invoiceNumberCol, "");
    }
    
    return lessons;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractAndMarkLessonsFromSheet", "ERROR", 
                                  "Failed to extract and mark lessons", 
                                  "Sheet: " + sheet.getName(), error.message);
    throw error;
  }
}

function extractLessonFromRow(row, headerMap, teacherData, cutoffDate, sheet, sheetRowIndex) {
  try {
    // Extract values using headerMap (convert 1-based to 0-based for array access)
    var studentId = row[headerMap['studentid'] - 1];
    var studentName = row[headerMap['studentname'] - 1];
    var dateValue = row[headerMap['date'] - 1];
    var lengthValue = row[headerMap['length'] - 1];
    var statusValue = row[headerMap['status'] - 1];
    var invoiceDateValue = headerMap['invoicedate'] ? row[headerMap['invoicedate'] - 1] : null;
    
    // Skip if essential data is missing
    if (!studentId || !studentName || !dateValue || !statusValue) {
      return null;
    }
    
    // Skip if already invoiced (Invoice Date column has value)
    if (invoiceDateValue && invoiceDateValue.toString().trim() !== '') {
      return null;
    }
    
    // Parse the lesson date
    var lessonDate = parseLessonDate(dateValue);
    if (!lessonDate || isNaN(lessonDate.getTime())) {
      return null; // Invalid date
    }
    
    // Only include lessons on or before cutoff date
    if (lessonDate > cutoffDate) {
      return null;
    }
    
    // Only include payable attendance ("lesson" or "no show")
    var status = statusValue.toString().toLowerCase().trim();
    if (['lesson', 'no show'].indexOf(status) === -1) {
      return null;
    }
    
    // Parse lesson length
    var lessonLength = parseInt(lengthValue) || 0;
    
    // Determine if this is a group session (ID starts with "G")
    var isGroupSession = studentId && studentId.toString().charAt(0) === 'G';
    
    // Validate lesson length based on type
    var isValidLength = false;
    
    if (isGroupSession) {
      // GROUP SESSION: 15-240 minutes, must end in 0 or 5
      if (lessonLength >= 15 && lessonLength <= 240) {
        var lastDigit = lessonLength % 10;
        if (lastDigit === 0 || lastDigit === 5) {
          isValidLength = true;
        }
      }
    } else {
      // STUDENT LESSON: Only 30, 45, or 60
      if (lessonLength === 30 || lessonLength === 45 || lessonLength === 60) {
        isValidLength = true;
      }
    }
    
    // If invalid length, write to Admin Comments and skip row
    if (!isValidLength) {
      if (headerMap['admincomments'] && sheet && sheetRowIndex) {
        var adminCommentsCol = headerMap['admincomments'];
        var currentComments = row[adminCommentsCol - 1];
        var errorMessage = 'Fix lesson length';
        
        // Append error message if comments already exist
        var newComments = currentComments && currentComments.toString().trim() !== '' 
          ? currentComments + ' | ' + errorMessage 
          : errorMessage;
        
        sheet.getRange(sheetRowIndex, adminCommentsCol).setValue(newComments);
        
        UtilityScriptLibrary.debugLog("extractLessonFromRow - VALIDATION ERROR: Invalid length " + lengthValue + 
                                      " for " + (isGroupSession ? "group" : "student") + 
                                      " - Row: " + sheetRowIndex + " - Wrote to Admin Comments");
      }
      return null; // Skip this row
    }
    
    return {
      teacherName: teacherData.teacherName,
      teacherId: teacherData.teacherId,
      studentId: studentId.toString().trim(),
      studentName: studentName.toString().trim(),
      lessonDate: lessonDate,
      status: status === 'no show' ? 'No Show' : 'Lesson',
      lessonLength: lessonLength
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractLessonFromRow - ERROR: Row extraction failed - " + error.message);
    return null;
  }
}

function extractLessonsFromAttendanceSheet(sheet, teacherData, cutoffDate) {
  try {
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // Empty sheet or only headers
    }
    
    // Use UtilityScriptLibrary to get header map (returns 1-based column numbers)
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Validate required columns exist
    var requiredColumns = ['studentid', 'studentname', 'date', 'length', 'status'];
    for (var j = 0; j < requiredColumns.length; j++) {
      var col = requiredColumns[j];
      if (!headerMap[col]) {
        throw new Error('Required column "' + col + '" not found in ' + sheet.getName());
      }
    }
    
    var lessons = [];
    
    // Process data rows (skip header row 0)
    for (var k = 1; k < data.length; k++) {
      var sheetRowIndex = k + 1; // Convert 0-based data index to 1-based sheet row
      var lesson = extractLessonFromRow(data[k], headerMap, teacherData, cutoffDate, sheet, sheetRowIndex);
      if (lesson) {
        lessons.push(lesson);
      }
    }
    
    UtilityScriptLibrary.debugLog('extractLessonsFromAttendanceSheet - Extracted ' + lessons.length + ' lessons from ' + sheet.getName());
    return lessons;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('extractLessonsFromAttendanceSheet - ERROR: ' + error.message);
    throw error;
  }
}

function extractTeacherInvoiceNumbers(invoiceSheet) {
  try {
    var data = invoiceSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(invoiceSheet);
    
    if (!headerMap['teacher'] || !headerMap['invoicenumber']) {
      UtilityScriptLibrary.debugLog("extractTeacherInvoiceNumbers", "WARNING", 
                                   "Required columns not found", "", "");
      return {};
    }
    
    var teacherCol = headerMap['teacher'] - 1;
    var invoiceNumberCol = headerMap['invoicenumber'] - 1;
    
    var teacherInvoiceNumbers = {};
    
    // Scan through rows looking for teacher header rows (have invoice number but no student name)
    for (var i = 1; i < data.length; i++) {
      var teacher = data[i][teacherCol];
      var invoiceNumber = data[i][invoiceNumberCol];
      
      if (teacher && invoiceNumber) {
        var teacherStr = teacher.toString().trim();
        var invoiceNumberStr = invoiceNumber.toString().trim();
        
        if (teacherStr && invoiceNumberStr && !teacherInvoiceNumbers[teacherStr]) {
          teacherInvoiceNumbers[teacherStr] = invoiceNumberStr;
        }
      }
    }
    
    UtilityScriptLibrary.debugLog("extractTeacherInvoiceNumbers", "INFO", 
                                 "Extracted invoice numbers", 
                                 "Count: " + Object.keys(teacherInvoiceNumbers).length, "");
    
    return teacherInvoiceNumbers;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractTeacherInvoiceNumbers", "ERROR", 
                                 "Failed to extract invoice numbers", "", error.message);
    return {};
  }
}

function extractTeachersFromFormattedSheet(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Get column indices
    var teacherCol = headerMap['teacher'];
    var urlCol = headerMap['url'];
    var idCol = headerMap['id'];
    var invoiceDateCol = headerMap['invoicedate'];
    var invoiceNumberCol = headerMap['invoicenumber'];
    var invoicePeriodCol = headerMap['invoiceperiod'];
    var addressCol = headerMap['address'];
    var durationCol = headerMap['duration'];
    var lastNameCol = headerMap['lastname'];
    var firstNameCol = headerMap['firstname'];
    var quantityCol = headerMap['quantity'];
    var rateCol = headerMap['rate'];
    var costCol = headerMap['cost'];
    var commentsCol = headerMap['comments'];  // ADD THIS LINE
    
    var teachers = [];
    var currentTeacher = null;
    
    for (var row = 1; row < data.length; row++) {
      var rowData = data[row];
      var teacherName = rowData[teacherCol - 1];
      var studentId = idCol ? rowData[idCol - 1] : '';
      var invoiceDate = invoiceDateCol ? rowData[invoiceDateCol - 1] : '';
      var invoiceNumber = invoiceNumberCol ? rowData[invoiceNumberCol - 1] : '';
      var duration = durationCol ? rowData[durationCol - 1] : '';
      
      if (!teacherName) continue;
      
      // Check if this is a teacher header row using multiple criteria:
      // 1. ID starts with 'T' (teacher pattern)
      // 2. Has invoice date/number (only teacher headers have these)
      // 3. No duration (student rows have duration)
      var isTeacherHeader = (studentId && studentId.toString().startsWith('T')) ||
                           (invoiceDate && invoiceDate.toString().trim() !== '') ||
                           (invoiceNumber && invoiceNumber.toString().trim() !== '') ||
                           (!duration || duration.toString().trim() === '');
      
      if (isTeacherHeader) {
        // This is a teacher header row
        var hasUrl = urlCol && rowData[urlCol - 1] && rowData[urlCol - 1].toString().trim() !== '';
        
        // Only process teachers without URLs
        if (!hasUrl) {
          // Get teacher's first and last name directly from the sheet
          var teacherFirstName = firstNameCol ? rowData[firstNameCol - 1] : '';
          var teacherLastName = lastNameCol ? rowData[lastNameCol - 1] : '';
          var teacherAddress = addressCol ? rowData[addressCol - 1] : '';
          var teacherComment = commentsCol ? rowData[commentsCol - 1] : '';  // ADD THIS LINE
          
          currentTeacher = {
            teacherName: teacherName,
            teacherFirstName: teacherFirstName,
            teacherLastName: teacherLastName,
            teacherAddress: teacherAddress,
            comment: teacherComment,  // ADD THIS LINE
            headerRowIndex: row + 1,
            invoiceNumber: invoiceNumber,
            invoiceDate: invoiceDate,
            invoicePeriod: invoicePeriodCol ? rowData[invoicePeriodCol - 1] : '',
            studentRows: []
          };
          teachers.push(currentTeacher);
        } else {
          currentTeacher = null; // Skip this teacher
        }
      } else if (currentTeacher) {
        // This is a student row for the current teacher
        var lastName = lastNameCol ? rowData[lastNameCol - 1] : '';
        var firstName = firstNameCol ? rowData[firstNameCol - 1] : '';
        var quantity = quantityCol ? parseInt(rowData[quantityCol - 1]) || 0 : 0;
        var rate = rateCol ? parseFloat(rowData[rateCol - 1]) || 0 : 0;
        var cost = costCol ? parseFloat(rowData[costCol - 1]) || 0 : 0;
        
        currentTeacher.studentRows.push({
          lastName: lastName || '',
          firstName: firstName || '',
          duration: duration,
          quantity: quantity,
          rate: rate,
          cost: cost
        });
      }
    }
    
    return teachers;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractTeachersFromFormattedSheet", "ERROR", "Failed to extract teacher data", "", error.message);
    throw error;
  }
}

function formatDateForInput(date) {
  try {
    var month = String(date.getMonth() + 1);
    var day = String(date.getDate());
    var year = date.getFullYear();
    
    // ES5-compatible zero padding
    if (month.length === 1) month = '0' + month;
    if (day.length === 1) day = '0' + day;
    
    return month + '/' + day + '/' + year;
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error formatting date: ' + error.message);
    return '01/01/2024'; // Fallback
  }
}

function formatInvoiceSheet(sheet) {
  try {
    UtilityScriptLibrary.debugLog("formatInvoiceSheet", "INFO", "Starting invoice sheet formatting", 
                                  "Sheet: " + sheet.getName(), "");
    
    // Get column mappings directly from utility library
    var columnMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol <= 0) {
      UtilityScriptLibrary.debugLog("formatInvoiceSheet", "WARNING", "No data to format", 
                                    "LastRow: " + lastRow + ", LastCol: " + lastCol, "");
      return;
    }
    
    // Apply basic formatting to data range
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange.setFontFamily("Arial")
             .setFontSize(10)
             .setWrap(true);
    
    // Format currency columns if they exist
    var currencyColumns = ['rate', 'cost'];
    for (var i = 0; i < currencyColumns.length; i++) {
      var colName = currencyColumns[i];
      if (columnMap[colName] !== undefined) {
        var colIndex = columnMap[colName] + 1; // Convert to 1-based
        var currencyRange = sheet.getRange(2, colIndex, lastRow - 1, 1);
        currencyRange.setNumberFormat('$#,##0.00');
      }
    }
    
    // REMOVED: Auto-resize columns - this was overriding template column widths
    // sheet.autoResizeColumns(1, lastCol);
    
    UtilityScriptLibrary.debugLog("formatInvoiceSheet", "INFO", "Invoice sheet formatting completed", "", "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("formatInvoiceSheet", "ERROR", "Failed to format invoice sheet", "", error.message);
    // Don't throw error - formatting failure shouldn't stop invoice generation
  }
}

function generateDefaultInvoicePeriod(cutoffDate) {
  try {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    var month = monthNames[cutoffDate.getMonth()];
    var year = cutoffDate.getFullYear();
    
    return month + ' ' + year;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error generating default invoice period: ' + error.message);
    return "Unknown Period";
  }
}

function generateDefaultMonthName(cutoffDate) {
  try {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    return monthNames[cutoffDate.getMonth()];
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error generating default month name: ' + error.message);
    return "January"; // Fallback
  }
}

function generateInvoiceNumber(teacherId, invoiceDate) {
  try {
    var year = invoiceDate.getFullYear();
    var month = String(invoiceDate.getMonth() + 1);
    var day = String(invoiceDate.getDate());
    
    // Pad month and day with leading zeros (ES5 compatible)
    if (month.length === 1) month = '0' + month;
    if (day.length === 1) day = '0' + day;
    
    var invoiceNumber = teacherId + '-' + year + month + day;
    
    UtilityScriptLibrary.debugLog("generateInvoiceNumber", "DEBUG", 
                                  "Generated invoice number", 
                                  "Teacher ID: " + teacherId + ", Date: " + UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MM/dd/yyyy') + ", Number: " + invoiceNumber, "");
    
    return invoiceNumber;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateInvoiceNumber", "ERROR", 
                                  "Failed to generate invoice number", 
                                  "Teacher ID: " + teacherId, error.message);
    return teacherId + '-ERROR';
  }
}

function generateMonthlyTeacherInvoices(month, cutoffDate, invoiceDate, lessonResults) {
  try {
    UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "INFO", 
                                  "Starting monthly invoice generation", 
                                  "Month: " + month + ", Cutoff: " + cutoffDate.toDateString(), "");
    
    // Use provided lesson results (already marked as invoiced during collection)
    if (!lessonResults || Object.keys(lessonResults.lessons).length === 0) {
      UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "WARNING", 
                                    "No uninvoiced lessons found", "", "");
      return { success: false, message: "No uninvoiced lessons found" };
    }
    
    // Create monthly invoice sheet
    var invoiceSheet = createMonthlyInvoiceSheet(month);
    
    // Populate invoice sheet with lesson data
    var populationResult = populateInvoiceSheetFromLessons(
      invoiceSheet, 
      lessonResults.lessons, 
      invoiceDate,
      lessonResults.invoicePeriod
    );
    
    // Format the sheet
    formatInvoiceSheet(invoiceSheet);
    
    UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "SUCCESS", 
                                  "Invoice generation completed", 
                                  "Teachers: " + populationResult.teacherCount + ", Line items: " + populationResult.lineItemCount, "");
    
    return {
      success: true,
      teacherCount: populationResult.teacherCount,
      lineItemCount: populationResult.lineItemCount,
      markedCount: lessonResults.summary.totalLessons  // Already marked during collection
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateMonthlyTeacherInvoices", "ERROR", 
                                  "Invoice generation failed", "", error.message);
    throw error;
  }
}

function generateSingleTeacherInvoice(teacherData, sheet, metadata) {
  try {
    // No need for contact lookup anymore - we have everything from the sheet
    var teacherContactInfo = {}; // Empty object since we're not using it
    
    // Build template variables (pass metadata for date range)
    var variables = buildTeacherInvoiceVariableMap(teacherData, teacherContactInfo, metadata);
    
    // Build filename
    var fileName = buildTeacherInvoiceFileName(teacherData, variables);
    
    // Get destination folder - pass sheet name for monthly organization
    var monthName = sheet.getName();
    var destinationFolder = getTeacherInvoicesFolder(monthName);
    
    // Check if document already exists
    if (UtilityScriptLibrary.documentAlreadyExists(fileName, destinationFolder)) {
      return {
        success: false,
        message: "Document already exists",
        fileName: fileName
      };
    }
    
    // Generate document
    var result = UtilityScriptLibrary.generateDocumentFromTemplate(
      'teacherInvoice',
      variables,
      fileName,
      destinationFolder
    );
    
    if (result.success) {
      // Set the document to view-only (readers can view but not edit)
      try {
        var file = DriveApp.getFileById(result.fileId);
        
        // Remove editor access and set to viewer access for anyone with the link
        // This makes it so teachers can view but not edit
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        UtilityScriptLibrary.debugLog("generateSingleTeacherInvoice", "INFO", 
                                      "Document set to view-only", 
                                      "FileId: " + result.fileId, "");
      } catch (sharingError) {
        // Log but don't fail - this shouldn't break invoice generation
        UtilityScriptLibrary.debugLog("generateSingleTeacherInvoice", "WARNING", 
                                      "Could not set document to view-only", 
                                      "FileId: " + result.fileId, 
                                      sharingError.message);
      }
      
      // Write URL back to teacher header row as hyperlink (text = fileId, link = URL)
      if (teacherData.headerRowIndex) {
        var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
        var urlCol = headerMap['url'];
        if (urlCol) {
          sheet.getRange(teacherData.headerRowIndex, urlCol).setFormula('=HYPERLINK("' + result.url + '", "' + result.fileId + '")');
        }
      }
      
      return {
        success: true,
        fileId: result.fileId,
        url: result.url,
        fileName: fileName
      };
    } else {
      return {
        success: false,
        error: result.error,
        fileName: fileName
      };
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateSingleTeacherInvoice", "ERROR", "Failed to generate single teacher invoice", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    throw error;
  }
}

function generateTeacherInvoiceDocuments() {
  try {
    var ui = SpreadsheetApp.getUi();
    var activeSheet = SpreadsheetApp.getActiveSheet();
    
    // Validate that this is a monthly invoice sheet
    if (!isMonthlyInvoiceSheet(activeSheet)) {
      ui.alert('This does not appear to be a monthly invoice sheet. Please select the correct sheet.');
      return;
    }
    
    var monthName = activeSheet.getName();
    
    // Get metadata for this invoice month
    var metadata;
    try {
      metadata = getTeacherInvoicingMetadata(monthName);
      UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "INFO", 
                                    "Loaded metadata", 
                                    "Month: " + monthName, "");
    } catch (metadataError) {
      UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "WARNING", 
                                    "Could not load metadata (will proceed without)", 
                                    "Month: " + monthName, metadataError.message);
      metadata = null;
    }
    
    // Extract teacher data from the already-formatted sheet
    var teacherData = extractTeachersFromFormattedSheet(activeSheet);
    
    if (teacherData.length === 0) {
      ui.alert('No teachers found that need invoice documents generated.');
      return;
    }
    
    // Generate documents for each teacher
    var results = {
      successful: 0,
      skipped: 0,
      errors: []
    };
    
    for (var i = 0; i < teacherData.length; i++) {
      var teacher = teacherData[i];
      
      try {
        var result = generateSingleTeacherInvoice(teacher, activeSheet, metadata);
        
        if (result.success) {
          results.successful++;
          UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "SUCCESS", 
                                        "Generated invoice for teacher", 
                                        "Teacher: " + teacher.teacherName, "");
          
          // Update teacher's invoice history (pass metadata for date range)
          updateTeacherInvoiceHistory(teacher, result, metadata);
          
        } else {
          results.skipped++;
          UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "WARNING", 
                                        "Skipped teacher invoice", 
                                        "Teacher: " + teacher.teacherName + ", Reason: " + result.message, "");
        }
        
      } catch (error) {
        results.errors.push({
          teacher: teacher.teacherName,
          error: error.message
        });
        UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "ERROR", 
                                      "Failed to generate teacher invoice", 
                                      "Teacher: " + teacher.teacherName, error.message);
      }
    }
    
    // Update metadata status to "Generated" if any documents were created
    if (results.successful > 0 && metadata) {
      try {
        updateMetadataStatus(monthName, 'Generated');
      } catch (statusError) {
        UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "WARNING", 
                                      "Could not update metadata status", 
                                      "", statusError.message);
      }
    }
    
    // Show results summary
    showTeacherInvoiceResults(results);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateTeacherInvoiceDocuments", "ERROR", "Teacher invoice generation failed", "", error.message);
    SpreadsheetApp.getUi().alert('Error generating teacher invoices: ' + error.message);
  }
}

function getActiveTeacherList() {
  try {
    UtilityScriptLibrary.debugLog("getActiveTeacherList", "INFO", "Getting active teacher list", "", "");
    
    // Use utility library to get form responses workbook
    var formResponsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
    var teacherLookupSheet = formResponsesSS.getSheetByName('Teacher Roster Lookup');
    
    if (!teacherLookupSheet) {
      throw new Error('Teacher Roster Lookup sheet not found in form responses workbook');
    }
    
    var data = teacherLookupSheet.getDataRange().getValues();
    
    // Use utility function for header mapping
    var headerMap = UtilityScriptLibrary.getHeaderMap(teacherLookupSheet);
    
    var teacherNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher Name')];
    var rosterUrlCol = headerMap[UtilityScriptLibrary.normalizeHeader('Roster URL')];
    var teacherIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    var statusCol = headerMap[UtilityScriptLibrary.normalizeHeader('Status')];
    
    var activeTeachers = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var teacherName = teacherNameCol ? row[teacherNameCol - 1] : '';
      var rosterUrl = rosterUrlCol ? row[rosterUrlCol - 1] : '';
      var teacherId = teacherIdCol ? row[teacherIdCol - 1] : '';
      var status = statusCol ? row[statusCol - 1] : '';
      
      // Only include active teachers with roster URLs
      if (status === 'Active' && rosterUrl && teacherName) {
        activeTeachers.push({
          teacherName: teacherName,
          teacherId: teacherId || teacherName, // fallback to name if no ID
          rosterUrl: rosterUrl
        });
      }
    }
    
    UtilityScriptLibrary.debugLog("getActiveTeacherList", "INFO", 
                          "Active teachers found", 
                          "Count: " + activeTeachers.length, "");
    
    return activeTeachers;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getActiveTeacherList", "ERROR", 
                          "Failed to get active teacher list", "", error.message);
    throw error;
  }
}

function getAttendanceSheetsFromWorkbook(spreadsheet) {
  try {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    var allSheets = spreadsheet.getSheets();
    var attendanceSheets = [];
    
    for (var i = 0; i < allSheets.length; i++) {
      var sheet = allSheets[i];
      var sheetName = sheet.getName();
      
      // Check if sheet name matches any month name (case-insensitive)
      for (var j = 0; j < monthNames.length; j++) {
        if (sheetName.toLowerCase() === monthNames[j].toLowerCase()) {
          attendanceSheets.push(sheet);
          UtilityScriptLibrary.debugLog("getAttendanceSheetsFromWorkbook - Found attendance sheet: " + sheet.getName());
          break; // No need to check other months once we find a match
        }
      }
    }
    
    UtilityScriptLibrary.debugLog("getAttendanceSheetsFromWorkbook - Attendance sheets found: " + attendanceSheets.length);
    
    return attendanceSheets;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getAttendanceSheetsFromWorkbook - ERROR: Failed to get attendance sheets - " + error.message);
    return [];
  }
}

function getCurrentSemesterName() {
  try {
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var semesterSheet = billingWorkbook.getSheetByName('Semester Metadata');
    
    if (!semesterSheet) {
      throw new Error('Semester Metadata sheet not found');
    }
    
    var data = semesterSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('No semester data found');
    }
    
    // Get the most recent semester (last row)
    var currentSemester = data[data.length - 1][0];
    
    UtilityScriptLibrary.debugLog("getCurrentSemesterName", "INFO", 
                                  "Found current semester", 
                                  "Semester: " + currentSemester, "");
    
    return currentSemester;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getCurrentSemesterName", "ERROR", 
                                  "Failed to get current semester", 
                                  "", error.message);
    return "Unknown Semester";
  }
}

function getRateForSemester(semesterName, rateType, ratesCache) {
  try {
    UtilityScriptLibrary.debugLog("getRateForSemester", "DEBUG", 
                                  "Looking up rate from cache", 
                                  "Semester: " + semesterName + ", Type: " + rateType, "");
    
    if (!ratesCache[rateType]) {
      throw new Error('Rate type "' + rateType + '" not found in rates cache');
    }
    
    var rate = ratesCache[rateType];
    
    UtilityScriptLibrary.debugLog("getRateForSemester", "INFO", 
                                  "Rate found", 
                                  "Type: " + rateType + ", Rate: " + rate, "");
    
    return rate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getRateForSemester", "ERROR", 
                                  "Rate lookup failed", 
                                  "Semester: " + semesterName + ", Type: " + rateType, error.message);
    throw error;
  }
}

function getRateKeyForProgram(programName, programRateKeysCache) {
  try {
    UtilityScriptLibrary.debugLog("getRateKeyForProgram", "DEBUG", 
                                  "Looking up Rate Key from cache", 
                                  "Program: " + programName, "");
    
    var normalizedProgramName = programName.toLowerCase().trim();
    
    if (!programRateKeysCache[normalizedProgramName]) {
      UtilityScriptLibrary.debugLog("getRateKeyForProgram", "WARNING", 
                                    "Rate Key not found for program", 
                                    "Program: " + programName, "");
      return null;
    }
    
    var rateKey = programRateKeysCache[normalizedProgramName];
    
    UtilityScriptLibrary.debugLog("getRateKeyForProgram", "INFO", 
                                  "Rate Key found", 
                                  "Program: " + programName + ", Rate Key: " + rateKey, "");
    
    return rateKey;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getRateKeyForProgram", "ERROR", 
                                  "Failed to get Rate Key", 
                                  "Program: " + programName, error.message);
    throw error;
  }
}

function getRatesLookupForSemester(semesterName) {
  try {
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var semesterSheet = billingWorkbook.getSheetByName('Semester Metadata');
    
    if (!semesterSheet) {
      throw new Error('Semester Metadata sheet not found');
    }
    
    var data = semesterSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var semesterNameCol = -1;
    var ratesVerificationCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var normalizedHeader = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (normalizedHeader === UtilityScriptLibrary.normalizeHeader("Semester Name")) {
        semesterNameCol = i;
      }
      if (normalizedHeader === UtilityScriptLibrary.normalizeHeader("Rates Verification")) {
        ratesVerificationCol = i;
      }
    }
    
    if (semesterNameCol === -1 || ratesVerificationCol === -1) {
      throw new Error("Required columns not found in Semester Metadata");
    }
    
    // Find the row matching the semester name
    for (var i = 1; i < data.length; i++) {
      if (data[i][semesterNameCol] === semesterName) {
        var ratesLookup = data[i][ratesVerificationCol];
        
        UtilityScriptLibrary.debugLog("getRatesLookupForSemester", "INFO", 
                                      "Found rates lookup", 
                                      "Semester: " + semesterName + ", Rates: " + ratesLookup, "");
        
        return ratesLookup || "";
      }
    }
    
    UtilityScriptLibrary.debugLog("getRatesLookupForSemester", "WARNING", 
                                  "Semester not found in metadata", 
                                  "Semester: " + semesterName, "");
    return "";
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getRatesLookupForSemester", "ERROR", 
                                  "Failed to get rates lookup", 
                                  "Semester: " + semesterName, error.message);
    return "";
  }
}

function getSemesterForDate(date) {
  try {
    UtilityScriptLibrary.debugLog("getSemesterForDate", "DEBUG", 
                                  "Looking up semester for date", 
                                  "Date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy'), "");
    
    // Get semester metadata from billing workbook
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var semesterSheet = billingWorkbook.getSheetByName('Semester Metadata');
    
    if (!semesterSheet) {
      throw new Error('Semester Metadata sheet not found in billing workbook');
    }
    
    var data = semesterSheet.getDataRange().getValues();
    UtilityScriptLibrary.debugLog("getSemesterForDate", "DEBUG", 
                                  "Retrieved semester metadata", 
                                  "Rows: " + data.length, "");
    
    // Find semester that contains this date
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var semesterName = row[0];
      var startDate = new Date(row[1]);
      var endDate = new Date(row[2]);
      
      UtilityScriptLibrary.debugLog("getSemesterForDate", "DEBUG", 
                                    "Checking semester", 
                                    "Semester: " + semesterName + ", Start: " + UtilityScriptLibrary.formatDateFlexible(startDate, 'MMM d, yyyy') + ", End: " + UtilityScriptLibrary.formatDateFlexible(endDate, 'MMM d, yyyy'), "");
      
      if (date >= startDate && date <= endDate) {
        UtilityScriptLibrary.debugLog("getSemesterForDate", "INFO", 
                                      "Found matching semester", 
                                      "Date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy') + ", Semester: " + semesterName, "");
        return semesterName;
      }
    }
    
    throw new Error("No semester found for date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy'));
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getSemesterForDate", "ERROR", 
                                  "Failed to get semester for date", 
                                  "Date: " + UtilityScriptLibrary.formatDateFlexible(date, 'MMM d, yyyy'), error.message);
    throw error;
  }
}

function getTeacherContactInfo(teacherId) {
  try {
    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "DEBUG", "Looking up contact info", "Teacher ID: " + teacherId, "");
    
    if (!teacherId || teacherId.toString().trim() === '') {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "WARNING", "No teacher ID provided", "", "");
      return { firstName: '', lastName: '', address: '' };
    }
    
    // Get the contacts workbook first
    var contactsWorkbook = UtilityScriptLibrary.getWorkbook('contacts');
    var teachersSheet = contactsWorkbook.getSheetByName('Teachers and Admin');
    
    if (!teachersSheet) {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", "Teachers and Admin sheet not found in contacts workbook", "", "");
      return { firstName: '', lastName: '', address: '' };
    }
    
    var data = teachersSheet.getDataRange().getValues();
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", "No data in Teachers and Admin sheet", "", "");
      return { firstName: '', lastName: '', address: '' };
    }
    
    var headers = data[0];
    
    // Find column indices using utility normalization - FIXED: Use correct normalized strings
    var idCol = -1;
    var firstNameCol = -1;
    var lastNameCol = -1;
    var addressCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var normalizedHeader = UtilityScriptLibrary.normalizeHeader(headers[i]);
      
      // FIXED: Match the actual normalized output (no spaces)
      if (normalizedHeader === 'teacherid') idCol = i;
      if (normalizedHeader === 'firstname') firstNameCol = i;
      if (normalizedHeader === 'lastname') lastNameCol = i;
      if (normalizedHeader === 'address') addressCol = i;
    }
    
    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "DEBUG", "Column indices found", 
                          "ID: " + idCol + ", First: " + firstNameCol + ", Last: " + lastNameCol + ", Address: " + addressCol, "");
    
    if (idCol === -1) {
      UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", 
                            "Teacher ID column not found in contacts", 
                            "Available headers: " + headers.join(", "), "");
      return { firstName: '', lastName: '', address: '' };
    }
    
    // Find teacher row by matching Teacher ID exactly
    for (var j = 1; j < data.length; j++) {
      var rowTeacherId = String(data[j][idCol] || '').trim();
      var searchTeacherId = String(teacherId).trim();
      
      if (rowTeacherId === searchTeacherId) {
        var firstName = firstNameCol !== -1 ? String(data[j][firstNameCol] || '').trim() : '';
        var lastName = lastNameCol !== -1 ? String(data[j][lastNameCol] || '').trim() : '';
        var address = addressCol !== -1 ? String(data[j][addressCol] || '').trim() : '';
        
        UtilityScriptLibrary.debugLog("getTeacherContactInfo", "DEBUG", 
                              "Found teacher contact info", 
                              "ID: " + teacherId + ", First: " + firstName + ", Last: " + lastName, "");
        
        return { firstName: firstName, lastName: lastName, address: address };
      }
    }
    
    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "WARNING", 
                          "Teacher not found in contacts", 
                          "ID: " + teacherId, "");
    return { firstName: '', lastName: '', address: '' };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherContactInfo", "ERROR", 
                          "Failed to get teacher contact info", 
                          "ID: " + teacherId, error.message);
    return { firstName: '', lastName: '', address: '' };
  }
}
 
function getTeacherInfoByName(teacherName) {
  try {
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "INFO", "Looking up teacher info", 
                 "Teacher: " + teacherName, "");
    
    // Use getSheet instead of getWorkbook + getSheetByName
    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    
    if (!lookupSheet || lookupSheet.getLastRow() <= 1) {
      UtilityScriptLibrary.debugLog("getTeacherInfoByName", "WARNING", "Lookup sheet not found or empty", "", "");
      return null;
    }
    
    // Get all data from lookup sheet
    var data = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 6).getValues();
    var searchName = String(teacherName).trim();
    
    // Try exact match first
    for (var i = 0; i < data.length; i++) {
      var rowTeacherName = String(data[i][0]).trim();
      
      if (rowTeacherName === searchName) {
        var teacherInfo = {
          teacherName: data[i][0],
          rosterUrl: data[i][1],
          teacherId: data[i][2],
          displayName: data[i][3],
          status: data[i][4],
          lastUpdated: data[i][5]
        };
        
        UtilityScriptLibrary.debugLog("getTeacherInfoByName", "SUCCESS", "Found exact match", 
                     "Teacher: " + searchName, "");
        return teacherInfo;
      }
    }
    
    // If no exact match, try matching last name
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "DEBUG", "No exact match, trying last name match", 
                 "Teacher: " + searchName, "");
    
    for (var i = 0; i < data.length; i++) {
      var rowTeacherName = String(data[i][0]).trim();
      var rowLastName = rowTeacherName.split(' ').pop();
      
      if (rowLastName === searchName || rowTeacherName.endsWith(' ' + searchName)) {
        var teacherInfo = {
          teacherName: data[i][0],
          rosterUrl: data[i][1],
          teacherId: data[i][2],
          displayName: data[i][3],
          status: data[i][4],
          lastUpdated: data[i][5]
        };
        
        UtilityScriptLibrary.debugLog("getTeacherInfoByName", "SUCCESS", "Found last name match", 
                     "Searched: " + searchName + ", Found: " + rowTeacherName, "");
        return teacherInfo;
      }
    }
    
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "WARNING", "Teacher not found", 
                 "Teacher: " + searchName, "");
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInfoByName", "ERROR", "Lookup failed", 
                 "Teacher: " + teacherName, error.message);
    return null;
  }
}

function getTeacherInfoFromLessonGroup(lessonGroup) {
  try {
    var teacherName = lessonGroup.teacherName;
    var teacherId = lessonGroup.teacherId;
    
    // Get teacher ID from Teacher Roster Lookup if not available
    var finalTeacherId = teacherId;
    if (!finalTeacherId) {
      // For now, just use empty string - we can improve this later
      finalTeacherId = '';
    }
    
    // Get comprehensive teacher contact info (firstName, lastName, address)
    var teacherContactInfo = getTeacherContactInfo(finalTeacherId);
    
    // Parse last name from teacher name if contact lookup didn't provide it
    var lastName = teacherContactInfo.lastName || teacherName;
    if (!teacherContactInfo.lastName && teacherName.indexOf(',') !== -1) {
      var nameParts = teacherName.split(',');
      lastName = nameParts[0] ? nameParts[0].trim() : teacherName;
    }
    
    UtilityScriptLibrary.debugLog("getTeacherInfoFromLessonGroup", "DEBUG", 
                                  "Teacher info compiled", 
                                  "Name: " + teacherName + ", ID: " + finalTeacherId, "");
    
    return {
      teacherName: teacherName,
      teacherId: finalTeacherId,
      lastName: lastName,
      firstName: teacherContactInfo.firstName,
      address: teacherContactInfo.address
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInfoFromLessonGroup", "ERROR", 
                                  "Failed to get teacher info", 
                                  "Teacher: " + lessonGroup.teacherName, error.message);
    return {
      teacherName: lessonGroup.teacherName,
      teacherId: '',
      lastName: lessonGroup.teacherName,
      firstName: '',
      address: ''
    };
  }
}

function getTeacherInvoicesFolder(monthName) {
  try {
    var mainFolder = UtilityScriptLibrary.getGeneratedDocumentsFolder();
    var teacherFolders = mainFolder.getFoldersByName("Teacher Invoices");
    
    if (!teacherFolders.hasNext()) {
      throw new Error("Teacher Invoices folder not found. Please create it manually in the generated documents folder.");
    }
    
    var teacherInvoicesFolder = teacherFolders.next();
    
    // Create or get the monthly subfolder
    var monthlyFolders = teacherInvoicesFolder.getFoldersByName(monthName);
    
    if (monthlyFolders.hasNext()) {
      UtilityScriptLibrary.debugLog("getTeacherInvoicesFolder", "INFO", "Using existing monthly folder", 
                    "Month: " + monthName, "");
      return monthlyFolders.next();
    } else {
      var newMonthFolder = teacherInvoicesFolder.createFolder(monthName);
      UtilityScriptLibrary.debugLog("getTeacherInvoicesFolder", "INFO", "Created new monthly folder", 
                    "Month: " + monthName, "");
      return newMonthFolder;
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInvoicesFolder", "ERROR", "Failed to get teacher invoices folder", 
                  "Month: " + monthName, error.message);
    throw error;
  }
}

function getTeacherInvoicingMetadata(invoiceMonth) {
  try {
    UtilityScriptLibrary.debugLog("getTeacherInvoicingMetadata", "INFO", 
                                  "Reading metadata", 
                                  "Month: " + invoiceMonth, "");
    
    var teacherInvoicesSS = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = teacherInvoicesSS.getSheetByName('Teacher Invoicing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Teacher Invoicing Metadata sheet not found');
    }
    
    var data = metadataSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    
    // Find row matching the invoice month
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var monthCol = headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Month')];
      
      if (row[monthCol - 1] === invoiceMonth) {
        var metadata = {
          invoiceMonth: row[headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Month')] - 1],
          lessonsStartingDate: row[headerMap[UtilityScriptLibrary.normalizeHeader('Lessons Starting Date')] - 1],
          lessonsEndingDate: row[headerMap[UtilityScriptLibrary.normalizeHeader('Lessons Ending Date')] - 1],
          invoiceDate: row[headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Date')] - 1],
          ratesLookup: row[headerMap[UtilityScriptLibrary.normalizeHeader('Rates Lookup')] - 1] || '',
          semesterName: row[headerMap[UtilityScriptLibrary.normalizeHeader('Semester Name')] - 1] || '',
          status: row[headerMap[UtilityScriptLibrary.normalizeHeader('Status')] - 1] || ''
        };
        
        UtilityScriptLibrary.debugLog("getTeacherInvoicingMetadata", "SUCCESS", 
                                      "Found metadata", 
                                      "Month: " + invoiceMonth + ", Semester: " + metadata.semesterName, "");
        
        return metadata;
      }
    }
    
    throw new Error("Metadata not found for month: " + invoiceMonth);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getTeacherInvoicingMetadata", "ERROR", 
                                  "Failed to read metadata", 
                                  "Month: " + invoiceMonth, error.message);
    throw error;
  }
}

function getUninvoicedLessonsForTeacher(teacherData, cutoffDate, invoiceDate, invoiceNumber) {
  try {
    UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "INFO", 
                                  "Starting single-pass lesson collection and marking", 
                                  "Teacher: " + teacherData.teacherName + ", Cutoff: " + cutoffDate.toDateString(), "");
    
    // Open teacher's roster workbook ONCE
    var rosterUrl = teacherData.rosterUrl;
    if (!rosterUrl) {
      throw new Error('No roster URL found for ' + teacherData.teacherName);
    }
    
    // Extract file ID from URL and open spreadsheet
    var fileIdMatch = rosterUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!fileIdMatch) {
      throw new Error('Invalid roster URL format for ' + teacherData.teacherName);
    }
    
    var rosterSS = SpreadsheetApp.openById(fileIdMatch[1]);
    
    // Get all attendance sheets (monthly sheets like "January", "February", etc.)
    var attendanceSheets = getAttendanceSheetsFromWorkbook(rosterSS);
    
    if (attendanceSheets.length === 0) {
      UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "WARNING", 
                                    "No attendance sheets found", 
                                    "Teacher: " + teacherData.teacherName, "");
      return [];
    }
    
    var uninvoicedLessons = [];
    var formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(invoiceDate, "MM/dd/yyyy");
    
    // Process each attendance sheet
    for (var i = 0; i < attendanceSheets.length; i++) {
      var sheet = attendanceSheets[i];
      try {
        // SINGLE PASS: Read lessons AND mark them as invoiced
        var sheetLessons = extractAndMarkLessonsFromSheet(
          sheet, 
          teacherData, 
          cutoffDate, 
          formattedInvoiceDate, 
          invoiceNumber
        );
        uninvoicedLessons = uninvoicedLessons.concat(sheetLessons);
        
        UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "INFO", 
                                      "Processed sheet", 
                                      "Sheet: " + sheet.getName() + ", Lessons: " + sheetLessons.length, "");
      } catch (sheetError) {
        UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "ERROR", 
                                      "Error processing sheet", 
                                      "Teacher: " + teacherData.teacherName + ", Sheet: " + sheet.getName(), 
                                      sheetError.message);
        // Continue with other sheets
      }
    }
    
    UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "SUCCESS", 
                                  "Completed single-pass processing", 
                                  "Teacher: " + teacherData.teacherName + ", Total lessons: " + uninvoicedLessons.length, "");
    
    return uninvoicedLessons;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getUninvoicedLessonsForTeacher", "ERROR", 
                                  "Failed to process teacher", 
                                  "Teacher: " + teacherData.teacherName, error.message);
    throw error;
  }
}

function groupLessonsByTeacherAndType(allLessons) {
  try {
    var grouped = {};
    
    for (var i = 0; i < allLessons.length; i++) {
      var lesson = allLessons[i];
      // Create grouping key: teacher + student + lesson length (one line per student+length)
      var groupKey = lesson.teacherName + '|' + lesson.studentId + '|' + lesson.lessonLength;
      
      if (!grouped[groupKey]) {
        grouped[groupKey] = {
          teacherName: lesson.teacherName,
          teacherId: lesson.teacherId,
          studentId: lesson.studentId,
          studentName: lesson.studentName,
          lessonLength: lesson.lessonLength,
          quantity: 0,
          lessons: []
        };
      }
      
      // Add lesson to group and increment quantity
      grouped[groupKey].quantity++;
      grouped[groupKey].lessons.push(lesson);
    }
    
    UtilityScriptLibrary.debugLog('📊 Grouped ' + allLessons.length + ' lessons into ' + Object.keys(grouped).length + ' line items');
    return grouped;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error in groupLessonsByTeacherAndType: ' + error.message);
    throw error;
  }
}

function isMonthlyInvoiceSheet(sheet) {
  try {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var expectedHeaders = ['Last Name', 'First Name', 'Teacher', 'Rate', 'Cost', 'URL'];
    
    var foundCount = 0;
    for (var i = 0; i < expectedHeaders.length; i++) {
      for (var j = 0; j < headers.length; j++) {
        if (UtilityScriptLibrary.normalizeHeader(headers[j]) === UtilityScriptLibrary.normalizeHeader(expectedHeaders[i])) {
          foundCount++;
          break;
        }
      }
    }
    
    return foundCount >= 4; // Must have at least 4 key headers
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("isMonthlyInvoiceSheet", "ERROR", "Error validating sheet", "", error.message);
    return false;
  }
}

function loadProgramRateKeysCache() {
  try {
    UtilityScriptLibrary.debugLog("loadProgramRateKeysCache", "INFO", "Loading program rate keys cache", "", "");
    
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var programsSheet = billingWorkbook.getSheetByName('Programs List');
    
    if (!programsSheet) {
      throw new Error('Programs List sheet not found in billing workbook');
    }
    
    var data = programsSheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error('No data found in Programs List sheet');
    }
    
    var headers = data[0];
    var programNameCol = -1;
    var rateKeyCol = -1;
    
    // Find column indices
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().toLowerCase().trim();
      if (header === 'program name') {
        programNameCol = i;
      } else if (header === 'rate key') {
        rateKeyCol = i;
      }
    }
    
    if (programNameCol === -1 || rateKeyCol === -1) {
      throw new Error('Required columns not found in Programs List (Program Name, Rate Key)');
    }
    
    var programRateKeysCache = {};
    
    // Load all program rate keys into cache (skip header row)
    for (var j = 1; j < data.length; j++) {
      var row = data[j];
      var programName = row[programNameCol] ? row[programNameCol].toString().toLowerCase().trim() : '';
      var rateKey = row[rateKeyCol] ? row[rateKeyCol].toString().trim() : '';
      
      if (programName && rateKey) {
        programRateKeysCache[programName] = rateKey;
      }
    }
    
    UtilityScriptLibrary.debugLog("loadProgramRateKeysCache", "SUCCESS", "Program rate keys cache loaded", 
                                  "Total programs: " + Object.keys(programRateKeysCache).length, "");
    
    return programRateKeysCache;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("loadProgramRateKeysCache", "ERROR", "Failed to load program rate keys cache", "", error.message);
    throw error;
  }
}

function loadRatesCache() {
  try {
    UtilityScriptLibrary.debugLog("loadRatesCache", "INFO", "Loading rates cache", "", "");
    
    var billingWorkbook = UtilityScriptLibrary.getWorkbook('billing');
    var ratesSheet = billingWorkbook.getSheetByName('Rates');
    
    if (!ratesSheet) {
      throw new Error('Rates sheet not found in billing workbook');
    }
    
    var data = ratesSheet.getDataRange().getValues();
    var currentRatesCol = 1; // Column B (0-indexed)
    
    var ratesCache = {};
    
    // Load all rates into cache (skip header row)
    for (var i = 1; i < data.length; i++) {
      var rateType = data[i][0]; // Column A contains rate titles
      var rate = data[i][currentRatesCol];
      
      if (rateType && rate !== undefined && rate !== null && rate !== '') {
        ratesCache[rateType] = rate;
      }
    }
    
    UtilityScriptLibrary.debugLog("loadRatesCache", "SUCCESS", "Rates cache loaded", 
                                  "Total rates: " + Object.keys(ratesCache).length, "");
    
    return ratesCache;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("loadRatesCache", "ERROR", "Failed to load rates cache", "", error.message);
    throw error;
  }
}

function parseLessonDate(dateValue) {
  try {
    // Handle different date formats
    if (dateValue instanceof Date) {
      return dateValue;
    }
        
    if (!dateValue) {
      return null;
    }
        
    var dateStr = dateValue.toString().trim();
        
    // Handle "Fri, 1/5" format - extract just the "1/5" part
    var dayDateMatch = dateStr.match(/\w+,\s*(\d{1,2}\/\d{1,2})/);
    if (dayDateMatch) {
      var dateOnly = dayDateMatch[1]; // e.g., "1/5"
      var currentYear = new Date().getFullYear();
      var fullDateStr = dateOnly + "/" + currentYear; // e.g., "1/5/2024"
            
      // Use UtilityScriptLibrary to parse
      var parsedDate = UtilityScriptLibrary.parseDateFromString(fullDateStr);
      if (parsedDate) {
        return parsedDate;
      }
    }
        
    // Try direct parsing with UtilityScriptLibrary
    var directParse = UtilityScriptLibrary.parseDateFromString(dateStr);
    if (directParse) {
      return directParse;
    }
        
    // Fallback to native Date constructor
    var nativeDate = new Date(dateStr);
    if (!isNaN(nativeDate.getTime())) {
      return nativeDate;
    }
        
    return null;
      
  } catch (error) {
    UtilityScriptLibrary.debugLog("parseLessonDate - ERROR: Date parsing failed - Input: " + dateValue + " - " + error.message);
    return null;
  }
}

function parseStudentName(studentName) {
  try {
    if (!studentName) {
      return { firstName: '', lastName: '' };
    }
    
    var cleanName = studentName;
    
    // Remove instrument suffix if present (e.g., "Cacacho, Isaac - Cello" -> "Cacacho, Isaac")
    var dashIndex = cleanName.indexOf(' - ');
    if (dashIndex !== -1) {
      cleanName = cleanName.substring(0, dashIndex);
    }
    
    // Check if this looks like a group name (no comma means it's not "Last, First" format)
    if (cleanName.indexOf(',') === -1) {
      // This is a group name - return it as lastName only
      UtilityScriptLibrary.debugLog("parseStudentName", "DEBUG", "Detected group name", 
                                    "Input: '" + studentName + "' -> Group: '" + cleanName + "'", "");
      return { firstName: '', lastName: cleanName };
    }
    
    // Parse "Last, First" format for student names
    var nameParts = cleanName.split(',');
    var lastName = nameParts[0] ? nameParts[0].trim() : cleanName;
    var firstName = nameParts[1] ? nameParts[1].trim() : '';
    
    UtilityScriptLibrary.debugLog("parseStudentName", "DEBUG", "Parsed student name", 
                                  "Input: '" + studentName + "' -> Last: '" + lastName + "', First: '" + firstName + "'", "");
    
    return {
      firstName: firstName,
      lastName: lastName
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("parseStudentName", "ERROR", "Failed to parse student name", 
                                  "Input: " + studentName, error.message);
    return { firstName: '', lastName: studentName || '' };
  }
}

function populateInvoiceSheetFromLessons(sheet, groupedLessons, invoiceDate, invoicePeriod) {
  try {
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "INFO", 
                                  "Starting invoice sheet population", 
                                  "Line items: " + Object.keys(groupedLessons).length, "");
    
    // Load rates and program rate keys once
    var ratesCache = loadRatesCache();
    var programRateKeysCache = loadProgramRateKeysCache();
    
    // Get column mappings directly from utility library
    var columnMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Log available columns for debugging
    var availableColumns = Object.keys(columnMap);
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "DEBUG", "Available columns", 
                                  "Columns: " + availableColumns.join(', '), "");
    
    var currentRow = 2; // Start after header
    var teacherCount = 0;
    var lineItemCount = 0;
    var currentTeacher = null;
    
    // Process each lesson group (student+duration combination)
    for (var key in groupedLessons) {
      if (groupedLessons.hasOwnProperty(key)) {
        var lessonGroup = groupedLessons[key];
        
        // If this is a new teacher, add teacher header row
        if (lessonGroup.teacherName !== currentTeacher) {
          // Add blank separator row (except for first teacher)
          if (currentTeacher !== null) {
            currentRow++;
          }
          
          teacherCount++;
          currentTeacher = lessonGroup.teacherName;
          
          // Get teacher information
          var teacherInfo = getTeacherInfoFromLessonGroup(lessonGroup);
          
          // Generate invoice number
          var invoiceNumber = generateInvoiceNumber(teacherInfo.teacherId, invoiceDate);
          
          // Add teacher header row
          currentRow = addTeacherHeaderRow(
            sheet, currentRow, columnMap, teacherInfo, 
            invoiceDate, invoiceNumber, invoicePeriod
          );
        }
        
        // Add student line item row with caches
        currentRow = addStudentLineItem(
          sheet, currentRow, columnMap, lessonGroup, currentTeacher, 
          ratesCache, programRateKeysCache
        );
        lineItemCount++;
      }
    }
    
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "INFO", 
                                  "Invoice sheet population completed", 
                                  "Teachers: " + teacherCount + ", Line items: " + lineItemCount, "");
    
    return {
      teacherCount: teacherCount,
      lineItemCount: lineItemCount
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateInvoiceSheetFromLessons", "ERROR", 
                                  "Invoice sheet population failed", "", error.message);
    throw error;
  }
}

function promptForCutoffDate() {
  try {
    UtilityScriptLibrary.debugLog("promptForCutoffDate - Starting cutoff date prompt");
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Lesson Cutoff Date',
      'Enter the cutoff date for lessons to include (MM/DD/YYYY).\n\n' +
      'All lessons on or before this date will be included in the invoice.\n\n' +
      'Example: 01/31/2024',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog("promptForCutoffDate - User response: Button: " + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog("promptForCutoffDate - User cancelled cutoff date prompt");
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    UtilityScriptLibrary.debugLog("promptForCutoffDate - Processing user input: " + userInput);
    
    // Try UtilityScriptLibrary first
    var cutoffDate;
    try {
      cutoffDate = UtilityScriptLibrary.parseDateFromString(userInput);
      UtilityScriptLibrary.debugLog("promptForCutoffDate - UtilityScriptLibrary.parseDateFromString result: " + (cutoffDate ? cutoffDate.toDateString() : 'null'));
    } catch (utilityError) {
      UtilityScriptLibrary.debugLog("promptForCutoffDate - WARNING: UtilityScriptLibrary.parseDateFromString failed for input: " + userInput + " - " + utilityError.message);
      cutoffDate = null;
    }
    
    // Fallback to native Date constructor
    if (!cutoffDate) {
      UtilityScriptLibrary.debugLog("promptForCutoffDate - WARNING: Falling back to native Date constructor");
      
      try {
        cutoffDate = new Date(userInput);
        if (isNaN(cutoffDate.getTime())) {
          cutoffDate = null;
        }
        UtilityScriptLibrary.debugLog("promptForCutoffDate - Native Date constructor result: " + (cutoffDate ? cutoffDate.toDateString() : 'null'));
      } catch (nativeError) {
        UtilityScriptLibrary.debugLog("promptForCutoffDate - ERROR: Native Date constructor failed for input: " + userInput + " - " + nativeError.message);
        cutoffDate = null;
      }
    }
    
    if (!cutoffDate) {
      var errorMsg = '❌ Invalid date format. Please use MM/DD/YYYY format (example: 01/31/2024).';
      ui.alert(errorMsg);
      UtilityScriptLibrary.debugLog("promptForCutoffDate - ERROR: Date parsing completely failed for input: " + userInput);
      return null;
    }
    
    UtilityScriptLibrary.debugLog("promptForCutoffDate - Successfully parsed cutoff date - Input: " + userInput + ", Result: " + cutoffDate.toDateString());
    
    return cutoffDate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("promptForCutoffDate - ERROR: Unexpected error - " + error.message);
    return null;
  }
}

function promptForInvoiceDate() {
  try {
    UtilityScriptLibrary.debugLog("promptForInvoiceDate - Starting invoice date prompt");
    
    var ui = SpreadsheetApp.getUi();
    var today = new Date();
    var defaultDateStr = formatDateForInput(today);
    
    UtilityScriptLibrary.debugLog("promptForInvoiceDate - Default date calculated - Today: " + today.toDateString() + ", Formatted: " + defaultDateStr);
    
    var response = ui.prompt(
      'Invoice Date',
      'Enter the invoice date to mark these lessons (MM/DD/YYYY).\n\n' +
      'This date will be stamped on all processed lessons.\n\n' +
      'Default (today): ' + defaultDateStr + '\n' +
      'Press OK to use default, or enter a different date:',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog("promptForInvoiceDate - User response: Button: " + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog("promptForInvoiceDate - User cancelled invoice date prompt");
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    
    // Use default (today) if user left it blank
    var invoiceDate;
    if (!userInput) {
      invoiceDate = today;
      UtilityScriptLibrary.debugLog("promptForInvoiceDate - Using default date (today): " + invoiceDate.toDateString());
    } else {
      // Parse user input
      UtilityScriptLibrary.debugLog("promptForInvoiceDate - Parsing user input: " + userInput);
      
      // Try UtilityScriptLibrary first
      try {
        invoiceDate = UtilityScriptLibrary.parseDateFromString(userInput);
        UtilityScriptLibrary.debugLog("promptForInvoiceDate - UtilityScriptLibrary.parseDateFromString result: " + (invoiceDate ? invoiceDate.toDateString() : 'null'));
      } catch (utilityError) {
        UtilityScriptLibrary.debugLog("promptForInvoiceDate - WARNING: UtilityScriptLibrary.parseDateFromString failed for input: " + userInput + " - " + utilityError.message);
        invoiceDate = null;
      }
      
      // Fallback to native Date constructor
      if (!invoiceDate) {
        try {
          invoiceDate = new Date(userInput);
          if (isNaN(invoiceDate.getTime())) {
            invoiceDate = null;
          }
          UtilityScriptLibrary.debugLog("promptForInvoiceDate - Native Date constructor result: " + (invoiceDate ? invoiceDate.toDateString() : 'null'));
        } catch (nativeError) {
          UtilityScriptLibrary.debugLog("promptForInvoiceDate - ERROR: Native Date constructor failed for input: " + userInput + " - " + nativeError.message);
          invoiceDate = null;
        }
      }
      
      if (!invoiceDate) {
        var errorMsg = '❌ Invalid date format. Please use MM/DD/YYYY format (example: 01/31/2024).';
        ui.alert(errorMsg);
        UtilityScriptLibrary.debugLog("promptForInvoiceDate - ERROR: Date parsing failed for input: " + userInput);
        return null;
      }
    }
    
    UtilityScriptLibrary.debugLog("promptForInvoiceDate - Successfully determined invoice date: " + invoiceDate.toDateString());
    
    return invoiceDate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("promptForInvoiceDate - ERROR: Unexpected error - " + error.message);
    return null;
  }
}

function promptForInvoicePeriod(cutoffDate) {
  try {
    UtilityScriptLibrary.debugLog("promptForInvoicePeriod - Starting invoice period prompt");
    
    var ui = SpreadsheetApp.getUi();
    
    // Generate default period from cutoff date (month/year)
    var defaultPeriod = generateDefaultInvoicePeriod(cutoffDate);
    
    UtilityScriptLibrary.debugLog("promptForInvoicePeriod - Default period calculated - Cutoff date: " + cutoffDate.toDateString() + ", Default period: " + defaultPeriod);
    
    var response = ui.prompt(
      'Invoice Period',
      'Invoice period description.\n\n' +
      'Default (based on cutoff date): ' + defaultPeriod + '\n\n' +
      'You can override with custom text like:\n' +
      '• "Summer 2025"\n' +
      '• "October-November 2024"\n' +
      '• "Q1 2024"\n\n' +
      'Press OK to use default, or enter custom period:',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog("promptForInvoicePeriod - User response: Button: " + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog("promptForInvoicePeriod - User cancelled invoice period prompt");
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    
    // Use default if user left it blank, otherwise use their input
    var finalPeriod = userInput || defaultPeriod;
    
    UtilityScriptLibrary.debugLog("promptForInvoicePeriod - Invoice period determined - User input: " + userInput + ", Final period: " + finalPeriod);
    
    return finalPeriod;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("promptForInvoicePeriod - ERROR: Unexpected error - " + error.message);
    // Return default as fallback
    return generateDefaultInvoicePeriod(cutoffDate);
  }
}

function promptForMonthName(cutoffDate) {
  try {
    UtilityScriptLibrary.debugLog("promptForMonthName - Starting month name prompt");
    
    // Generate default month name from cutoff date
    var defaultMonth = generateDefaultMonthName(cutoffDate);
    
    UtilityScriptLibrary.debugLog("promptForMonthName - Default month calculated - Cutoff date: " + cutoffDate.toDateString() + ", Default month: " + defaultMonth);
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Invoice Sheet Month',
      'Enter the month name for the invoice sheet.\n\n' +
      'Default (based on cutoff date): ' + defaultMonth + '\n\n' +
      'Examples:\n' +
      '• January\n' +
      '• February\n' +
      '• March\n\n' +
      'Press OK to use default, or enter a different month:',
      ui.ButtonSet.OK_CANCEL
    );
    
    UtilityScriptLibrary.debugLog("promptForMonthName - User response: Button: " + response.getSelectedButton() + ', Text: "' + response.getResponseText() + '"');
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog("promptForMonthName - User cancelled month name prompt");
      return null;
    }
    
    var userInput = response.getResponseText().trim();
    
    // Use default if user left it blank, otherwise use their input
    var finalMonth;
    if (!userInput) {
      finalMonth = defaultMonth;
      UtilityScriptLibrary.debugLog("promptForMonthName - Using default month: " + finalMonth);
    } else {
      // Capitalize first letter
      finalMonth = userInput.charAt(0).toUpperCase() + userInput.slice(1).toLowerCase();
      UtilityScriptLibrary.debugLog("promptForMonthName - Using custom month - Input: " + userInput + ", Final: " + finalMonth);
    }
    
    return finalMonth;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("promptForMonthName - ERROR: Unexpected error - " + error.message);
    return null;
  }
}

function setupMetadataStatusDropdown(metadataSheet) {
  try {
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    var statusCol = headerMap[UtilityScriptLibrary.normalizeHeader("Status")];
    
    if (!statusCol) {
      UtilityScriptLibrary.debugLog("setupMetadataStatusDropdown", "WARNING", 
                                    "Status column not found", "", "");
      return;
    }
    
    // Set up data validation for Status column (starting from row 2)
    var lastRow = Math.max(metadataSheet.getLastRow(), 100); // At least 100 rows for future use
    var statusRange = metadataSheet.getRange(2, statusCol, lastRow - 1, 1);
    
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Collected', 'Generated', 'Sent', 'Paid'], true)
      .setAllowInvalid(false)
      .build();
    
    statusRange.setDataValidation(rule);
    
    UtilityScriptLibrary.debugLog("setupMetadataStatusDropdown", "SUCCESS", 
                                  "Status dropdown configured", 
                                  "Column: " + statusCol + ", Rows: 2-" + lastRow, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("setupMetadataStatusDropdown", "ERROR", 
                                  "Failed to setup status dropdown", 
                                  "", error.message);
  }
}

function showCombinedErrorDetails(lessonResults, invoiceResults) {
  try {
    var ui = SpreadsheetApp.getUi();
    var detailsMessage = '❌ ERRORS AND ISSUES:\n\n';
    
    // Show lesson collection errors
    if (lessonResults.errors.length > 0) {
      detailsMessage += '🚫 LESSON COLLECTION ERRORS (' + lessonResults.errors.length + '):\n';
      for (var i = 0; i < lessonResults.errors.length; i++) {
        var error = lessonResults.errors[i];
        if (i < 5) { // Limit to first 5
          detailsMessage += '• ' + error.teacher + ': ' + error.error + '\n';
        }
      }
      if (lessonResults.errors.length > 5) {
        detailsMessage += '• ... and ' + (lessonResults.errors.length - 5) + ' more errors\n';
      }
      detailsMessage += '\n';
    }
    
    // Show invoice generation errors
    if (invoiceResults.errors && invoiceResults.errors.length > 0) {
      detailsMessage += '🚫 INVOICE GENERATION ERRORS (' + invoiceResults.errors.length + '):\n';
      for (var j = 0; j < invoiceResults.errors.length; j++) {
        var invoiceError = invoiceResults.errors[j];
        if (j < 5) { // Limit to first 5
          detailsMessage += '• ' + invoiceError.teacher + ': ' + invoiceError.error + '\n';
        }
      }
      if (invoiceResults.errors.length > 5) {
        detailsMessage += '• ... and ' + (invoiceResults.errors.length - 5) + ' more errors\n';
      }
      detailsMessage += '\n';
    }
    
    // Show validation issues
    if (lessonResults.validation.issues.length > 0) {
      detailsMessage += '⚠️ VALIDATION ISSUES (' + lessonResults.validation.issues.length + '):\n';
      for (var k = 0; k < lessonResults.validation.issues.length; k++) {
        var issue = lessonResults.validation.issues[k];
        if (k < 5) { // Limit to first 5
          detailsMessage += '• ' + issue + '\n';
        }
      }
      if (lessonResults.validation.issues.length > 5) {
        detailsMessage += '• ... and ' + (lessonResults.validation.issues.length - 5) + ' more issues\n';
      }
    }
    
    detailsMessage += '\nCheck Teacher_Invoice_Debug sheet for complete details.';
    
    ui.alert('Error Details', detailsMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error in showCombinedErrorDetails: ' + error.message);
  }
}

function showInvoiceGenerationResults(lessonResults, invoiceResults) {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Build comprehensive summary message
    var summaryMessage = '🎉 INVOICE GENERATION COMPLETE!\n\n';
    
    // Invoice info
    summaryMessage += '📋 INVOICE SHEET:\n';
    summaryMessage += '• Sheet name: ' + invoiceResults.sheetName + '\n';
    summaryMessage += '• Teachers processed: ' + invoiceResults.teacherCount + '\n';
    summaryMessage += '• Line items created: ' + invoiceResults.lineItemCount + '\n\n';
    
    // Lesson data info
    summaryMessage += '📅 LESSON DATA:\n';
    summaryMessage += '• Cutoff date: ' + lessonResults.cutoffDate.toDateString() + '\n';
    summaryMessage += '• Invoice date: ' + lessonResults.invoiceDate.toDateString() + '\n';
    summaryMessage += '• Invoice period: ' + lessonResults.invoicePeriod + '\n';
    summaryMessage += '• Total lessons found: ' + lessonResults.summary.totalLessons + '\n\n';
    
    // Status
    var totalErrors = lessonResults.errors.length + (invoiceResults.errors ? invoiceResults.errors.length : 0);
    var totalIssues = lessonResults.validation.issues.length;
    
    if (totalErrors === 0 && totalIssues === 0) {
      summaryMessage += '✅ SUCCESS: Invoice sheet generated successfully!\n';
    } else {
      summaryMessage += '⚠️ COMPLETED WITH ISSUES:\n';
      summaryMessage += '• Errors: ' + totalErrors + '\n';
      summaryMessage += '• Validation issues: ' + totalIssues + '\n';
    }
    
    summaryMessage += '\nCheck Teacher_Invoice_Debug sheet for detailed logs.';
    
    // Show the summary
    ui.alert('Invoice Generation Complete', summaryMessage, ui.ButtonSet.OK);
    
    // If there are errors or issues, offer to show details
    if (totalErrors > 0 || totalIssues > 0) {
      var showDetailsResponse = ui.alert(
        'Show Details?',
        'Would you like to see details about the ' + totalErrors + ' errors and ' + totalIssues + ' validation issues?',
        ui.ButtonSet.YES_NO
      );
      
      if (showDetailsResponse === ui.Button.YES) {
        showCombinedErrorDetails(lessonResults, invoiceResults);
      }
    }
    
    UtilityScriptLibrary.debugLog("showInvoiceGenerationResults - Results summary displayed to user");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showInvoiceGenerationResults - ERROR: Error showing results summary - " + error.message);
  }
}

function showTeacherInvoiceResults(results) {
  try {
    var ui = SpreadsheetApp.getUi();
    var message = 'Teacher Invoice Generation Complete!\n\n';
    
    message += 'Successfully generated: ' + results.successful + '\n';
    message += 'Skipped (already exist): ' + results.skipped + '\n';
    message += 'Errors: ' + results.errors.length + '\n';
    
    if (results.errors.length > 0) {
      message += '\nErrors:\n';
      for (var i = 0; i < Math.min(results.errors.length, 3); i++) {
        message += '• ' + results.errors[i].teacher + ': ' + results.errors[i].error + '\n';
      }
      if (results.errors.length > 3) {
        message += '• ... and ' + (results.errors.length - 3) + ' more (check debug log)\n';
      }
    }
    
    ui.alert('Invoice Generation Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showTeacherInvoiceResults", "ERROR", "Failed to show results", "", error.message);
  }
}

function updateMetadataStatus(invoiceMonth, newStatus) {
  try {
    UtilityScriptLibrary.debugLog("updateMetadataStatus", "INFO", 
                                  "Updating status", 
                                  "Month: " + invoiceMonth + ", Status: " + newStatus, "");
    
    var teacherInvoicesSS = UtilityScriptLibrary.getWorkbook('teacherInvoices');
    var metadataSheet = teacherInvoicesSS.getSheetByName('Teacher Invoicing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Teacher Invoicing Metadata sheet not found');
    }
    
    var data = metadataSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    
    var monthCol = headerMap[UtilityScriptLibrary.normalizeHeader("Invoice Month")];
    var statusCol = headerMap[UtilityScriptLibrary.normalizeHeader("Status")];
    
    // Find the row matching the invoice month
    for (var i = 1; i < data.length; i++) {
      if (data[i][monthCol - 1] === invoiceMonth) {
        metadataSheet.getRange(i + 1, statusCol).setValue(newStatus);
        
        UtilityScriptLibrary.debugLog("updateMetadataStatus", "SUCCESS", 
                                      "Status updated", 
                                      "Month: " + invoiceMonth + ", New status: " + newStatus, "");
        return;
      }
    }
    
    throw new Error("Metadata not found for month: " + invoiceMonth);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("updateMetadataStatus", "ERROR", 
                                  "Failed to update status", 
                                  "Month: " + invoiceMonth, error.message);
    throw error;
  }
}

function updateTeacherInvoiceHistory(teacherData, invoiceResult, metadata) {
  try {
    // Only log if we have a successful invoice
    if (!invoiceResult || !invoiceResult.success) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "No successful invoice to log", 
                   "Teacher: " + teacherData.teacherName, "");
      return;
    }
    
    // Get teacher info using existing utility
    var teacherInfo = getTeacherInfoByName(teacherData.teacherName);
    if (!teacherInfo || !teacherInfo.rosterUrl) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Teacher info not found", 
                   "Teacher: " + teacherData.teacherName, "");
      return;
    }
    
    // Open teacher workbook using existing URL parsing pattern
    var fileIdMatch = teacherInfo.rosterUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!fileIdMatch) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Invalid roster URL", 
                   "Teacher: " + teacherData.teacherName + ", URL: " + teacherInfo.rosterUrl, "");
      return;
    }
    
    var teacherWorkbook = SpreadsheetApp.openById(fileIdMatch[1]);
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Opened teacher workbook", 
                 "Teacher: " + teacherData.teacherName + ", Workbook name: " + teacherWorkbook.getName(), "");
    
    // Get or create Invoice Log sheet
    var invoiceLogSheet = teacherWorkbook.getSheetByName('Invoice Log');
    if (!invoiceLogSheet) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Invoice Log sheet not found", 
                   "Teacher: " + teacherData.teacherName + 
                   ", Available sheets: " + teacherWorkbook.getSheets().map(function(s) { return s.getName(); }).join(", "), "");
      return; // Sheet should exist if created in Responses, but handle gracefully
    }
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Found Invoice Log sheet", 
                 "Teacher: " + teacherData.teacherName + ", Last row: " + invoiceLogSheet.getLastRow(), "");
    
    // Calculate total from teacher's student rows
    var totalAmount = 0;
    if (!teacherData.studentRows || teacherData.studentRows.length === 0) {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "No student rows found", 
                   "Teacher: " + teacherData.teacherName, "");
    } else {
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Calculating total amount", 
                   "Teacher: " + teacherData.teacherName + ", Student rows: " + teacherData.studentRows.length, "");
      
      for (var i = 0; i < teacherData.studentRows.length; i++) {
        var rowCost = teacherData.studentRows[i].cost || 0;
        totalAmount += rowCost;
        UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Adding student cost", 
                     "Row " + i + ": " + rowCost + ", Running total: " + totalAmount, "");
      }
    }
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Total calculated", 
                 "Teacher: " + teacherData.teacherName + ", Total: " + totalAmount, "");
    
    // Format dates and create invoice period string
    var formattedInvoiceDate = '';
    var invoicePeriodString = '';
    
    if (metadata && metadata.lessonsStartingDate && metadata.lessonsEndingDate) {
      // Use metadata for date range
      var startDate = new Date(metadata.lessonsStartingDate);
      var endDate = new Date(metadata.lessonsEndingDate);
      
      var formattedStart = UtilityScriptLibrary.formatDateFlexible(startDate, "MM/dd/yyyy");
      var formattedEnd = UtilityScriptLibrary.formatDateFlexible(endDate, "MM/dd/yyyy");
      invoicePeriodString = formattedStart + " - " + formattedEnd;
      
      if (metadata.invoiceDate) {
        formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(new Date(metadata.invoiceDate), "MM/dd/yyyy");
      }
      
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Using metadata for dates", 
                   "Period: " + invoicePeriodString + ", Invoice Date: " + formattedInvoiceDate, "");
    } else {
      // Fallback: use teacher data (old behavior)
      if (teacherData.invoiceDate) {
        formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(new Date(teacherData.invoiceDate), "MM/dd/yyyy");
      }
      invoicePeriodString = teacherData.invoicePeriod || '';
      
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "WARNING", "Using fallback for dates (no metadata)", 
                   "Period: " + invoicePeriodString, "");
    }
    
    // Add new invoice record - append first to get the row number
    var newRow = [
      teacherData.invoiceNumber || '',           // Invoice Number
      formattedInvoiceDate || '',                // Invoice Date
      invoicePeriodString || '',                 // Invoice Period (now a date range)
      '',                                        // Invoice URL (will be set as formula)
      UtilityScriptLibrary.formatCurrency(totalAmount) // Total Amount
    ];
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Prepared row data", 
                 "Teacher: " + teacherData.teacherName + 
                 ", Invoice#: " + newRow[0] + 
                 ", Date: " + newRow[1] + 
                 ", Period: " + newRow[2] + 
                 ", Has URL: " + (invoiceResult.url ? "YES" : "NO") + 
                 ", Total: " + newRow[4], "");
    
    invoiceLogSheet.appendRow(newRow);
    var newRowNumber = invoiceLogSheet.getLastRow();
    
    // Now set the URL column (column 4) as a HYPERLINK formula (text = fileId, link = URL)
    if (invoiceResult.url && invoiceResult.fileId) {
      var urlFormula = '=HYPERLINK("' + invoiceResult.url + '", "' + invoiceResult.fileId + '")';
      invoiceLogSheet.getRange(newRowNumber, 4).setFormula(urlFormula);
      
      UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "DEBUG", "Set URL as hyperlink formula", 
                   "Row: " + newRowNumber + ", FileId: " + invoiceResult.fileId, "");
    }
    
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "SUCCESS", "Added invoice to log", 
                 "Teacher: " + teacherData.teacherName + ", Invoice: " + teacherData.invoiceNumber + 
                 ", New last row: " + newRowNumber, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("updateTeacherInvoiceHistory", "ERROR", "Failed to update invoice history", 
                 "Teacher: " + teacherData.teacherName + ", Error: " + error.message + 
                 ", Stack: " + error.stack, "");
    // Don't throw - this is not critical to the invoice generation process
  }
}

function validateLessonData(groupedLessons) {
  var issues = [];
  
  var groupKeys = Object.keys(groupedLessons);
  for (var i = 0; i < groupKeys.length; i++) {
    var groupKey = groupKeys[i];
    var group = groupedLessons[groupKey];
    
    // Check for missing required data
    if (!group.teacherName) {
      issues.push(groupKey + ': Missing teacher name');
    }
    if (!group.studentId) {
      issues.push(groupKey + ': Missing student ID');
    }
    if (!group.studentName) {
      issues.push(groupKey + ': Missing student name');
    }
    if (!group.lessonLength || ![30, 45, 60].includes(group.lessonLength)) {
      issues.push(groupKey + ': Invalid lesson length (' + group.lessonLength + ')');
    }
    if (group.quantity <= 0) {
      issues.push(groupKey + ': Invalid quantity (' + group.quantity + ')');
    }
    
    // Validate individual lessons in group
    for (var j = 0; j < group.lessons.length; j++) {
      var lesson = group.lessons[j];
      if (!lesson.lessonDate || isNaN(lesson.lessonDate.getTime())) {
        issues.push(groupKey + ': Invalid lesson date in group');
      }
    }
  }
  
  UtilityScriptLibrary.debugLog('🔍 Validation: ' + issues.length + ' issues found');
  
  return {
    issues: issues,
    isValid: issues.length === 0
  };
}

function writeTeacherInvoicingMetadata(month, cutoffDate, invoiceDate, invoicePeriod) {
  try {
    UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "INFO", 
                                  "Writing metadata", 
                                  "Month: " + month, "");
    
    var teacherInvoicesSS = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = teacherInvoicesSS.getSheetByName('Teacher Invoicing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Teacher Invoicing Metadata sheet not found');
    }
    
    // Calculate lessons starting date
    var lessonsStartingDate = calculateLessonsStartingDate(metadataSheet);
    
    // If no previous metadata, calculate from cutoff date (first day of month)
    if (!lessonsStartingDate) {
      lessonsStartingDate = new Date(cutoffDate.getFullYear(), cutoffDate.getMonth(), 1);
      UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "INFO", 
                                    "Using first day of month for starting date", 
                                    "Date: " + lessonsStartingDate.toDateString(), "");
    }
    
    // Get current semester
    var semesterName = getCurrentSemesterName();
    
    // Get rates lookup for semester
    var ratesLookup = getRatesLookupForSemester(semesterName);
    
    // Prepare row data
    var newRow = [
      month,                          // Invoice Month
      lessonsStartingDate,            // Lessons Starting Date
      cutoffDate,                     // Lessons Ending Date (cutoff date)
      invoiceDate,                    // Invoice Date
      ratesLookup,                    // Rates Lookup
      semesterName,                   // Semester Name
      'Collected'                     // Status (initially "Collected", then becomes "Generated")
    ];
    
    metadataSheet.appendRow(newRow);
    
    UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "SUCCESS", 
                                  "Metadata written", 
                                  "Month: " + month + ", Period: " + lessonsStartingDate.toDateString() + " to " + cutoffDate.toDateString(), "");
    
    return {
      success: true,
      lessonsStartingDate: lessonsStartingDate,
      lessonsEndingDate: cutoffDate
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("writeTeacherInvoicingMetadata", "ERROR", 
                                  "Failed to write metadata", 
                                  "Month: " + month, error.message);
    throw error;
  }
}

// TEST FUNCTIONS //

function checkTemplateFormatting() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Monthly Template");
    
    if (!templateSheet) {
      UtilityScriptLibrary.debugLog("checkTemplateFormatting", "ERROR", "Monthly Template sheet not found", "", "");
      return;
    }
    
    var lastCol = templateSheet.getLastColumn();
    
    // Get headers from row 1
    var headers = templateSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Get formats from row 2 (first data row)
    var formats = templateSheet.getRange(2, 1, 1, lastCol).getNumberFormats()[0];
    
    UtilityScriptLibrary.debugLog("checkTemplateFormatting", "INFO", "Starting template format analysis", 
                                  "Sheet: Monthly Template, Columns: " + lastCol, "");
    
    var results = [];
    
    for (var col = 0; col < lastCol; col++) {
      var header = headers[col] || "No Header";
      var format = formats[col] || "General";
      var colLetter = UtilityScriptLibrary.columnToLetter(col + 1);
      
      // Categorize the format
      var category = "";
      if (format.indexOf("$") !== -1) {
        category = "CURRENCY";
      } else if (format === "0.00") {
        category = "DECIMAL";
      } else if (format === "0") {
        category = "INTEGER";
      } else if (format.indexOf("M/d") !== -1 || format.indexOf("MM/dd") !== -1) {
        category = "DATE";
      } else if (format === "General") {
        category = "GENERAL";
      } else {
        category = "OTHER";
      }
      
      var result = {
        column: col + 1,
        letter: colLetter,
        header: header,
        format: format,
        category: category
      };
      
      results.push(result);
      
      UtilityScriptLibrary.debugLog("checkTemplateFormatting", "DEBUG", "Column analysis", 
                                    "Col " + colLetter + ": " + header, 
                                    "Format: " + format + " | Category: " + category);
    }
    
    // Summary counts using ES5-compatible loops
    var currencyCount = 0;
    var dateCount = 0;
    var generalCount = 0;
    var otherCount = 0;
    
    for (var i = 0; i < results.length; i++) {
      if (results[i].category === "CURRENCY") currencyCount++;
      else if (results[i].category === "DATE") dateCount++;
      else if (results[i].category === "GENERAL") generalCount++;
      else if (results[i].category === "OTHER") otherCount++;
    }
    
    UtilityScriptLibrary.debugLog("checkTemplateFormatting", "INFO", "Template formatting summary", 
                                  "Currency: " + currencyCount + ", Date: " + dateCount + ", General: " + generalCount + ", Other: " + otherCount, "");
    
    // Find specific columns of interest using ES5-compatible loops
    var rateCol = null;
    var costCol = null;
    
    for (var j = 0; j < results.length; j++) {
      if (results[j].header === "Rate") {
        rateCol = results[j];
      }
      if (results[j].header === "Cost") {
        costCol = results[j];
      }
    }
    
    if (rateCol) {
      UtilityScriptLibrary.debugLog("checkTemplateFormatting", "INFO", "Rate column analysis", 
                                    "Column " + rateCol.letter + ": " + rateCol.format, 
                                    "Category: " + rateCol.category);
    }
    
    if (costCol) {
      UtilityScriptLibrary.debugLog("checkTemplateFormatting", "INFO", "Cost column analysis", 
                                    "Column " + costCol.letter + ": " + costCol.format, 
                                    "Category: " + costCol.category);
    }
    
    // Show user-friendly alert
    var ui = SpreadsheetApp.getUi();
    var alertMessage = "Template Format Analysis Complete!\n\n" +
                      "Summary:\n" +
                      "• Currency columns: " + currencyCount + "\n" +
                      "• Date columns: " + dateCount + "\n" +
                      "• General format: " + generalCount + "\n" +
                      "• Other formats: " + otherCount + "\n\n";
    
    if (rateCol) {
      alertMessage += "Rate column (" + rateCol.letter + "): " + rateCol.category + "\n";
    }
    
    if (costCol) {
      alertMessage += "Cost column (" + costCol.letter + "): " + costCol.category + "\n";
    }
    
    alertMessage += "\nCheck Debug sheet for detailed results.";
    
    ui.alert("Template Format Check", alertMessage, ui.ButtonSet.OK);
    
    return results;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("checkTemplateFormatting", "ERROR", "Failed to check template formatting", "", error.message);
    throw error;
  }
}

function debugMonthlySheetStructure() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    UtilityScriptLibrary.debugLog("DEBUG", "INFO", "Sheet analysis starting", "Sheet: " + sheet.getName(), "");
    
    var teacherCol = headerMap['teacher'];
    var urlCol = headerMap['url'];
    var lastNameCol = headerMap['lastname'];
    var firstNameCol = headerMap['firstname'];
    
    UtilityScriptLibrary.debugLog("DEBUG", "INFO", "Column mapping", 
                                  "Teacher col: " + teacherCol + ", URL col: " + urlCol + ", LastName col: " + lastNameCol + ", FirstName col: " + firstNameCol, "");
    
    for (var row = 1; row < Math.min(data.length, 10); row++) {
      var rowData = data[row];
      var teacherName = teacherCol ? rowData[teacherCol - 1] : '';
      var url = urlCol ? rowData[urlCol - 1] : '';
      var lastName = lastNameCol ? rowData[lastNameCol - 1] : '';
      var firstName = firstNameCol ? rowData[firstNameCol - 1] : '';
      
      var rowType = (!lastName && !firstName) ? "TEACHER HEADER" : "STUDENT ROW";
      
      UtilityScriptLibrary.debugLog("DEBUG", "INFO", "Row " + (row + 1) + " analysis", 
                                    "Type: " + rowType + ", Teacher: '" + teacherName + "', URL: '" + url + "', Student: '" + lastName + ", " + firstName + "'", "");
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("DEBUG", "ERROR", "Debug function failed", "", error.message);
  }
}

function testHeaderReading() {
  var contactsWorkbook = UtilityScriptLibrary.getWorkbook('contacts');
  var teachersSheet = contactsWorkbook.getSheetByName('Teachers and Admin');
  var headers = teachersSheet.getRange(1, 1, 1, teachersSheet.getLastColumn()).getValues()[0];
  
  UtilityScriptLibrary.debugLog("=== HEADER ANALYSIS ===");
  for (var i = 0; i < headers.length; i++) {
    var original = String(headers[i]);
    var normalized = UtilityScriptLibrary.normalizeHeader(original);
    UtilityScriptLibrary.debugLog("Column " + i + ": Original='" + original + "' | Normalized='" + normalized + "'");
  }
  
  // Test the specific function with T0003
  UtilityScriptLibrary.debugLog("=== TESTING getTeacherContactInfo ===");
  var result = getTeacherContactInfo("T0003");
  UtilityScriptLibrary.debugLog("Result: " + JSON.stringify(result));
}
