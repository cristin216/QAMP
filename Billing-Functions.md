================================================================================
BILLING FUNCTION DIRECTORY
================================================================================
    Total Functions: 180
    Most Recent version: 168

    This directory provides a quick reference for all functions in Billing script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), UtilityScriptLibrary.function2()

  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    addInvoiceTotalFormula
    addLateRegistrationsToBillingCycle
    addLateRegistrationsUI
    addMissingStudentsToAttendanceSheet
    appendReregistrationNewStudents
    appendStudentsFromContext
    appendToBillingMetadata
    appendToSemesterMetadata
    applyBillingConditionalFormatting
    applyCumulativeHistory
    applyLateFeeToRow
    applyLetterTypeValidation
    applyMultiStudentDiscount
    applyReregistrationOverwrites
    applyWarningsToTeacherWorkbook
    backfillCumulativeTracking
    buildBillingContext
    buildBillingRowFromForm
    buildBillingRowFromPrevious
    buildCarryoverStudents
    buildDocumentFileName
    buildDocumentSentence
    buildDynamicAmounts
    buildDynamicLineItems
    buildDynamicProgramColumns
    buildFormStudents
    buildFutureSemesterContext
    buildInvoiceTotalFormula
    buildInvoiceVariableMap
    buildMissingDocumentSentence
    buildPastVlookupFormula
    buildProgramDescription
    buildReregistrationQueue
    buildTemplateVariables
    calculateTotalCreditsApplied
    cancelDocumentGeneration
    checkIfDocumentStillNeeded
    checkIfMediaReleaseNeeded
    clearDocIdFromBillingSheet
    collectBillingDataDetailed
    collectPaymentsDataDetailed
    continueAttendanceSheetCreation
    continuePacketGeneration
    convertFolderDocsToPdfUI
    copyStaticFieldsToBillingRow
    createBillingSheet
    createDetailedPaymentReport
    createNewAttendanceSheets
    createPaymentsTab
    createPaymentVerificationReport
    createRosterFolder
    createSingleRegistrationPacketWithSelection
    determineIfNewStudent
    determinePacketVersions
    doGet
    executeDocumentGeneration
    expandSheetAttendanceRows
    expandTeacherAttendanceRows
    expandTeacherAttendanceSheets
    extractBillingDataFromRow
    extractDeliveryPreference
    extractDocumentNames
    extractPreviousBillingData
    extractProgramTotals
    extractRosterDataForAttendance
    extractStudentDataFromBillingRow
    finalizeBillingSheet
    findBillingRowByStudentId
    formatRow
    generateCalendarForSemester
    generateDocumentForStudent
    generateInvoiceForStudent
    generateInvoicesForBillingCycle
    generateProgramFormulas
    generateReconciliationSummary
    generateRefundInvoicesForBillingCycle
    generateRegistrationPacketForStudentWithSelection
    generateRegistrationPacketsForBillingCycle
    getActivePrograms
    getBillingSheet
    getCumulativeHistory
    getCurrentBillingCycleDates
    getCurrentBillingSheet
    getCurrentRateChartName
    getCurrentSemesterFromBillingMetadata
    getCurrentSemesterInfo
    getCurrentSemesterRateForLength
    getDocIdColumnName
    getDocIdFromBillingSheet
    getDocumentSelectionHtml
    getExpandedPrograms
    getFilterDate
    getFormsDataFromContacts
    getFormSheet
    getFutureSemesterName
    getInvoiceNumber
    getLessonLengthFromRow
    getNextMonthName
    getPreviousSemester
    getRateColumnFromMetadata
    getRateMap
    getStudentBalancesFromBilling
    getStudentDocumentsFolder
    getStudentRegisteredLessonLength
    getStudentsNeedingPackets
    handleMonthYearCancel
    handleMonthYearSelection
    identifyWarningStudents
    isHeaderRow
    loadProgramConfig
    loadReregistrationData
    locateStudentRecord
    locateStudentRecordEnhanced
    logMysteryStudents
    markReregistrationProcessed
    markStudentsInactive
    migrateTeacherDisplayNamesToIds
    onOpen
    parseMonthYear
    populateAllCumulativeColumns
    populateBillingSheet
    populateBillingSheetContinuingSemester
    populateCurrentBalanceFormula
    populateDeliveryPreference
    populateDeliveryPreferenceFromPrevious
    populateInvoiceMetadata
    populateLateFee
    populateLetterType
    processDocumentSelection
    processFieldMapForSemester
    processFormsData
    processFormsReconciliationForRow
    processPaymentReconciliationForRow
    processTeacherAttendanceForBilling
    processTeacherForNewAttendance
    processTeacherReconciliation
    promptForBillingCycleName
    promptForMonthAndYear
    promptForSemesterDates
    promptForSemesterName
    protectBillingSheet
    protectPreviousBillingCycle
    reconcilePayment
    renameLatestFormSheet
    runBillingCycleAutomation
    runFullReconciliation
    runFullReconciliationUI
    runLogHeaders
    runPaymentsReconciliation
    runPaymentsReconciliationUI
    runRegistrationPacketGenerationUI
    runWeeklyLessonReconciliation
    runWeeklyLessonReconciliationUI
    selectDocumentTemplate
    sendReregistrationLinks
    setDeliveryPreference
    setPastColumnFormulas
    setProgramQuantitiesForCarryover
    setupNewSemester
    shouldGenerateInvoice
    shouldIncludeAgreement
    shouldIncludeDocument
    shouldIncludeMediaRelease
    shouldUseMissingDocumentLetter
    showSimpleDocumentSelectionDialog
    submitReregistration
    sumPayments
    updateBillingForStudents
    updateCumulativeTracking
    updateDocIdInBillingSheet
    updateInvoiceUrlInBillingSheet
    updateLastReconciliationDate
    updateSheetStudentWarnings
    updateTeacherRosterBalances
    verifyAndGetParentData
    verifyPaymentsDetailed
    verifyPaymentsDetailedForStudents
    verifyProgramsForSemester
    verifyRatesEnhanced
    writeAndFormatRows

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addInvoiceTotalFormula(newRow, context, rowIndex, currencyCols) -> void
        Adds a formula to the Current Invoice Total column of a billing row array.
        Category: BILLING_ROW_BUILDERS
        Local functions used: buildInvoiceTotalFormula()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    addLateRegistrationsToBillingCycle() -> Object
        Adds late-registering students to the current billing cycle sheet by processing
        new form submissions that arrived after the initial billing run.
        Category: BILLING_CYCLE
        Local functions used: buildBillingContext(), loadProgramConfig(), getFormSheet(),
                              buildFormStudents(), writeAndFormatRows(), finalizeBillingSheet()
        Utility functions used: UtilityScriptLibrary.promptForCustomToday(), UtilityScriptLibrary.debugLog()

    addLateRegistrationsUI() -> void
        UI wrapper for addLateRegistrationsToBillingCycle(). Displays result alert.
        Category: UI
        Local functions used: addLateRegistrationsToBillingCycle()
        Utility functions used: UtilityScriptLibrary.debugLog()

    addMissingStudentsToAttendanceSheet(sheet, activeStudents) -> Number
        Adds any active roster students not already present to an existing attendance
        sheet. Returns the count of students added.
        Category: ATTENDANCE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    appendReregistrationNewStudents(billingSheet, reregMap, overwroteIds, context) -> void
        Appends new re-registration students (those not already in the billing sheet)
        from the re-registration queue to the billing sheet.
        Category: REREGISTRATION
        Local functions used: buildBillingRowFromForm(), copyStaticFieldsToBillingRow(),
                              applyCumulativeHistory(), populateInvoiceMetadata()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    appendStudentsFromContext(billingSheet, context, existingIds) -> void
        Appends new students from the form responses sheet to the billing sheet,
        skipping students already present.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm(), writeAndFormatRows()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.getHeaderMap()

    appendToBillingMetadata(billingCycleName, customToday, semesterName) -> void
        Appends a new row to the Billing Metadata sheet for a completed billing cycle.
        Category: METADATA
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.formatDateFlexible(),
                                UtilityScriptLibrary.debugLog()

    appendToSemesterMetadata(semesterName, startDate, endDate) -> Object
        Appends semester start and end dates to the Semester Metadata sheet.
        Category: METADATA
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.debugLog()

    applyBillingConditionalFormatting(billingSheet) -> void
        Applies conditional formatting rules to a billing sheet for visual indicators
        on balance, status, and warning columns.
        Category: FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.getHeaderMap()

    applyCumulativeHistory(newRow, studentId, context, currencyCols) -> void
        Populates cumulative taught/billed/balance columns in a billing row from the
        Cumulative Tracking sheet.
        Category: BILLING_ROW_BUILDERS
        Local functions used: getCumulativeHistory()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    applyLateFeeToRow(newRow, context, currencyCols, pastBalanceValue) -> void
        Applies a late fee to a billing row if the student's past balance qualifies.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    applyLetterTypeValidation(billingSheet) -> void
        Applies a data validation dropdown to the Letter Type column of a billing sheet.
        Category: FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    applyMultiStudentDiscount(billingSheet) -> void
        Identifies families with multiple students and applies a multi-student discount
        to qualifying rows in the billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    applyReregistrationOverwrites(billingSheet, reregMap, formStudentIds) -> void
        Overwrites billing row fields for re-registering students with updated data
        from the re-registration queue.
        Category: REREGISTRATION
        Local functions used: copyStaticFieldsToBillingRow()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    applyWarningsToTeacherWorkbook(teacherSS, warningStudents, targetDate) -> void
        Writes warning indicators to the appropriate attendance sheets in a teacher
        workbook for students with low balances.
        Category: RECONCILIATION
        Local functions used: updateSheetStudentWarnings()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.getMonthNames()

    backfillCumulativeTracking() -> void
        One-time utility to populate the Cumulative Tracking sheet from historical
        billing cycle data. Iterates all billing sheets in reverse chronological order
        and writes cumTaught, cumBilled, and balance per student.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    buildBillingContext(customToday, semesterName, billingCycleName, programConfig) -> Object
        Builds the billing context object used throughout a billing cycle run. Contains
        header maps, sheet references, semester info, rate maps, and helper functions.
        Category: BILLING_CYCLE
        Local functions used: loadProgramConfig(), getRateMap(), getFormSheet(),
                              getRateColumnFromMetadata(), getCurrentSemesterFromBillingMetadata()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getSheet(),
                                UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    buildBillingRowFromForm(formRow, prevRow, context, rowIndex) -> Array
        Builds a complete new billing row array from a form response row, merging with
        previous billing data where applicable.
        Category: BILLING_ROW_BUILDERS
        Local functions used: copyStaticFieldsToBillingRow(), buildDynamicProgramColumns(),
                              populateDeliveryPreference(), populateInvoiceMetadata(),
                              populateLateFee(), populateLetterType(), populateCurrentBalanceFormula(),
                              applyCumulativeHistory(), setPastColumnFormulas(), generateProgramFormulas(),
                              addInvoiceTotalFormula()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    buildBillingRowFromPrevious(prevRow, context, rowIndex) -> Array
        Builds a complete new billing row array from a previous billing cycle row for
        carryover students.
        Category: BILLING_ROW_BUILDERS
        Local functions used: copyStaticFieldsToBillingRow(), buildDynamicProgramColumns(),
                              populateDeliveryPreferenceFromPrevious(), populateInvoiceMetadata(),
                              populateLateFee(), populateLetterType(), populateCurrentBalanceFormula(),
                              applyCumulativeHistory(), setPastColumnFormulas(), generateProgramFormulas(),
                              addInvoiceTotalFormula()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    buildCarryoverStudents(previousData, existingStudentIds, context, allStudents, startingRowIndex) -> Number
        Iterates previous billing data and builds carryover rows for students not
        already in the new billing sheet. Returns the next available row index.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromPrevious()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    buildDocumentFileName(studentData, billingData, documentType, deliveryMethod) -> String
        Builds a document file name string for a registration packet document.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: None

    buildDocumentSentence(billingData, deliveryMethod) -> String
        Builds the document list sentence for a registration packet letter, describing
        which documents are included.
        Category: DOCUMENT_GENERATION
        Local functions used: extractDocumentNames()
        Utility functions used: None

    buildDynamicAmounts(billingData) -> Object
        Builds dynamic amount variables (quantities, rates, costs) for invoice template
        population from a billingData object.
        Category: DOCUMENT_GENERATION
        Local functions used: extractProgramTotals()
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildDynamicLineItems(billingData) -> Object
        Builds dynamic line item strings for invoice PDF template population.
        Category: DOCUMENT_GENERATION
        Local functions used: buildProgramDescription(), extractProgramTotals()
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildDynamicProgramColumns(newRow, row, context, quantityCols, currencyCols, rowNum, lessonQuantity) -> void
        Populates program quantity and cost columns in a billing row array based on
        the student's lesson package and rates.
        Category: BILLING_ROW_BUILDERS
        Local functions used: getLessonLengthFromRow()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    buildFormStudents(formSheet, filterDate, previousData, context, allStudents, startingRowIndex) -> Number
        Iterates form response rows after the filter date and builds billing rows for
        new registrants. Returns the next available row index.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    buildFutureSemesterContext(futureSemesterName, customToday, billingCycleName, programConfig, prevHeaderMap) -> Object
        Builds a billing context for a future semester, used when processing
        pre-registrations for the next term.
        Category: BILLING_CYCLE
        Local functions used: buildBillingContext()
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildInvoiceTotalFormula(headerMap, rowNum) -> String
        Builds a spreadsheet SUM formula string for the Current Invoice Total column
        spanning all program cost columns.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.columnToLetter(), UtilityScriptLibrary.normalizeHeader()

    buildInvoiceVariableMap(studentData, billingData, isRefund) -> Object
        Builds the template variable map for an invoice document.
        Category: DOCUMENT_GENERATION
        Local functions used: buildDynamicLineItems(), buildDynamicAmounts()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible()

    buildMissingDocumentSentence(billingData) -> String
        Builds the missing document list sentence for a missing document letter.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: None

    buildPastVlookupFormula(prevSheetName, prevColIndex, rowNum) -> String
        Builds a VLOOKUP formula string to pull a value from the previous billing
        cycle sheet.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: None

    buildProgramDescription(programTotals, lessonLength) -> String
        Builds a human-readable program description string from program totals and
        lesson length.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: None

    buildReregistrationQueue() -> Object
        Reads the re-registration sheet and builds a queue map of students to be
        processed in the current billing cycle.
        Category: REREGISTRATION
        Local functions used: loadReregistrationData()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    buildTemplateVariables(studentData, billingData, templateType) -> Object
        Builds the full template variable map for any registration packet document
        type.
        Category: DOCUMENT_GENERATION
        Local functions used: buildInvoiceVariableMap(), buildDocumentSentence(),
                              buildMissingDocumentSentence()
        Utility functions used: UtilityScriptLibrary.formatDateFlexible()

    calculateTotalCreditsApplied(billingData) -> Number
        Calculates the sum of all credit amounts in a billingData object.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    cancelDocumentGeneration() -> void
        Clears stored document selection from script properties, cancelling a pending
        packet generation.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: None

    checkIfDocumentStillNeeded(studentId, docType, billingSheet, headerMap) -> Boolean
        Returns true if a student still needs a given document type based on the
        billing sheet's document status columns.
        Category: DOCUMENT_GENERATION
        Local functions used: shouldIncludeAgreement(), shouldIncludeMediaRelease()
        Utility functions used: None

    checkIfMediaReleaseNeeded(studentId) -> Boolean
        Returns true if a student does not yet have a media release on file in the
        Contacts workbook.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    clearDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap) -> void
        Clears the document ID and URL from the billing sheet for a given student and
        document type.
        Category: DOCUMENT_GENERATION
        Local functions used: getDocIdColumnName()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    collectBillingDataDetailed(billingWB, studentIds) -> Object
        Collects billing row data for a list of student IDs across all billing sheets
        in a workbook. Returns a map keyed by student ID.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    collectPaymentsDataDetailed(paymentsWB, studentIds) -> Object
        Collects payment row data for a list of student IDs across all payment sheets.
        Returns a map keyed by student ID.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    continueAttendanceSheetCreation(targetMonthName, targetYear) -> void
        Processes all active teachers and creates or updates attendance sheets for the
        given month/year. Called after the month/year UI prompt completes.
        Category: ATTENDANCE
        Local functions used: getActivePrograms(), processTeacherForNewAttendance()
        Utility functions used: UtilityScriptLibrary.debugLog()

    continuePacketGeneration() -> void
        Continues packet generation after document type selection. Reads stored
        selection from script properties and calls executeDocumentGeneration().
        Category: DOCUMENT_GENERATION
        Local functions used: executeDocumentGeneration()
        Utility functions used: UtilityScriptLibrary.debugLog()

    convertFolderDocsToPdfUI() -> void
        UI-triggered function to convert all Google Docs in a student documents folder
        to PDF format.
        Category: DOCUMENT_GENERATION
        Local functions used: getStudentDocumentsFolder()
        Utility functions used: UtilityScriptLibrary.debugLog()

    copyStaticFieldsToBillingRow(newRow, sourceRow, context, getFn) -> void
        Copies static student/parent contact fields into a new billing row array from
        a source row using the field map.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    createBillingSheet(billingMonth, programConfig) -> Sheet
        Creates a new billing sheet for the given month with all required column
        headers based on the program configuration.
        Category: SHEET_OPERATIONS
        Local functions used: loadProgramConfig(), applyLetterTypeValidation(),
                              applyBillingConditionalFormatting(), protectBillingSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createDetailedPaymentReport(billingWB, paymentsData, billingData, studentIds) -> void
        Creates a detailed payment report sheet in the billing workbook comparing
        payment and billing data for specific students.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    createNewAttendanceSheets() -> void
        UI entry point for attendance sheet creation. Prompts for month/year and
        calls continueAttendanceSheetCreation().
        Category: UI
        Local functions used: promptForMonthAndYear()
        Utility functions used: None

    createPaymentsTab(semesterName) -> Sheet
        Creates a new payments tab in the billing workbook for the given semester with
        standard column headers.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createPaymentVerificationReport(billingWB, paymentsData, billingData) -> void
        Creates a payment verification report sheet comparing all payment and billing
        data across the workbook.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: None

    createRosterFolder(semesterName) -> Folder
        Creates a new Drive roster folder for the given semester under the main roster
        folder.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getRosterFolder(), UtilityScriptLibrary.debugLog()

    createSingleRegistrationPacketWithSelection(studentData, billingData, packetType, currentSemester, destinationFolder, forceRegenerate, selectedDocTypes) -> Object
        Generates a registration packet for a single student using only the selected
        document types. Returns a result object.
        Category: DOCUMENT_GENERATION
        Local functions used: generateDocumentForStudent(), shouldIncludeDocument(),
                              determinePacketVersions()
        Utility functions used: UtilityScriptLibrary.debugLog()

    determineIfNewStudent(studentId, currentSemester) -> Boolean
        Returns true if the student's first enrollment term matches the current
        semester, indicating they are new this term.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    determinePacketVersions(deliveryPreference) -> Object
        Returns an object with email and print boolean flags based on the delivery
        preference string.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    doGet(e) -> HtmlOutput
        Web app GET handler. Serves the ReRegistration HTML page with parent ID
        pre-populated from URL parameter.
        Category: UI
        Local functions used: None
        Utility functions used: None

    executeDocumentGeneration() -> void
        Executes packet generation for all selected students using stored script
        properties for document type selection.
        Category: DOCUMENT_GENERATION
        Local functions used: getStudentsNeedingPackets(), generateRegistrationPacketForStudentWithSelection(),
                              continuePacketGeneration()
        Utility functions used: UtilityScriptLibrary.debugLog()

    expandSheetAttendanceRows(sheet) -> Object
        Expands the lesson rows section of a single attendance sheet to accommodate
        more entries. Returns a result object.
        Category: ATTENDANCE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    expandTeacherAttendanceRows(teacherSS) -> void
        Expands attendance rows in all month sheets of a teacher workbook.
        Category: ATTENDANCE
        Local functions used: expandSheetAttendanceRows()
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.debugLog()

    expandTeacherAttendanceSheets() -> void
        UI entry point. Prompts for a teacher workbook Drive ID and calls
        expandTeacherAttendanceRows().
        Category: UI
        Local functions used: expandTeacherAttendanceRows()
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractBillingDataFromRow(billingRowData, headerMap) -> Object
        Extracts all billing fields from a billing row array using the header map.
        Returns a billingData object.
        Category: DATA_EXTRACTION
        Local functions used: extractDeliveryPreference(), extractProgramTotals()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    extractDeliveryPreference(billingRowData, headerMap) -> String
        Extracts and normalizes the delivery preference value from a billing row.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    extractDocumentNames(documents) -> Array
        Extracts display name strings from an array of document objects.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: None

    extractPreviousBillingData(options?) -> Object
        Reads the previous billing cycle sheet and returns its data as an array.
        options.includeAll controls whether inactive students are included.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    extractProgramTotals(row, headerMap) -> Array
        Extracts program quantity, rate, and cost columns from a billing row using
        the header map. Returns an array of program total objects.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    extractRosterDataForAttendance(rosterSheet) -> Array
        Reads a roster sheet and returns an array of active/carryover student objects
        for use in attendance sheet creation.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: None

    extractStudentDataFromBillingRow(row, headerMap) -> Object
        Extracts student contact fields from a billing row array using the header map.
        Returns a studentData object.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    finalizeBillingSheet(billingSheet, context, formStudentIdsMap) -> void
        Runs final steps after populating a billing sheet: applies multi-student
        discounts, reregistration overwrites, conditional formatting, and protections.
        Category: BILLING_CYCLE
        Local functions used: applyMultiStudentDiscount(), applyReregistrationOverwrites(),
                              applyBillingConditionalFormatting(), protectBillingSheet(),
                              markReregistrationProcessed()
        Utility functions used: UtilityScriptLibrary.debugLog()

    findBillingRowByStudentId(billingData, studentId, studentIdColIndex) -> Number
        Finds the row index of a student in a billing data array by student ID.
        Returns -1 if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    formatRow(sheet, rowIndex, quantityCols, currencyCols) -> void
        Applies number formatting to quantity and currency columns for a single
        billing row.
        Category: FORMATTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.addToCurrencyCols(), UtilityScriptLibrary.debugLog()

    generateCalendarForSemester(semesterName, startDate, endDate) -> Object
        Generates a calendar sheet for a semester with week-by-week date columns.
        Category: SEMESTER_SETUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.formatDateFlexible(), UtilityScriptLibrary.debugLog()

    generateDocumentForStudent(studentData, billingData, templateType, deliveryMethod, currentSemester, billingSheetName, billingSheet, headerMap) -> Object
        Generates a single document (invoice, letter, agreement, or media release) for
        a student from the appropriate template. Returns a result object.
        Category: DOCUMENT_GENERATION
        Local functions used: selectDocumentTemplate(), buildTemplateVariables(),
                              buildDocumentFileName(), updateDocIdInBillingSheet()
        Utility functions used: UtilityScriptLibrary.generateDocumentFromTemplate(),
                                UtilityScriptLibrary.debugLog()

    generateInvoiceForStudent(studentData, row, headerMap, invoiceOptions, billingSheet, billingSheetName) -> Object
        Generates an invoice document for a single student. Returns a result object
        with success status and file info.
        Category: DOCUMENT_GENERATION
        Local functions used: extractBillingDataFromRow(), extractStudentDataFromBillingRow(),
                              generateDocumentForStudent(), updateInvoiceUrlInBillingSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateInvoicesForBillingCycle(billingSheetName, options?) -> Object
        Generates invoice documents for all eligible students in a billing cycle sheet.
        options.includeNegativeBalances controls refund invoice behavior.
        Category: DOCUMENT_GENERATION
        Local functions used: generateInvoiceForStudent(), shouldGenerateInvoice()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    generateProgramFormulas(newRow, context, rowNum, quantityCols, currencyCols) -> void
        Writes cost calculation formulas into program cost columns of a billing row.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.columnToLetter()

    generateReconciliationSummary(results) -> String
        Builds a summary string from a reconciliation results object for display.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    generateRefundInvoicesForBillingCycle(billingSheetName) -> Object
        Generates refund invoices for students with negative balances in a billing
        cycle sheet.
        Category: DOCUMENT_GENERATION
        Local functions used: generateInvoicesForBillingCycle()
        Utility functions used: None

    generateRegistrationPacketForStudentWithSelection(studentData, billingData, selectedDocumentTypes, deliveryMethods, currentSemester) -> Object
        Generates a full registration packet for a student using selected document
        types and delivery methods. Returns a result object.
        Category: DOCUMENT_GENERATION
        Local functions used: createSingleRegistrationPacketWithSelection()
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateRegistrationPacketsForBillingCycle() -> void
        UI entry point to generate registration packets for all students in the active
        billing sheet who need them.
        Category: UI
        Local functions used: getStudentsNeedingPackets(), showSimpleDocumentSelectionDialog()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getActivePrograms() -> Array
        Returns an array of active program names from the Programs List sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    getBillingSheet(paymentDate, activeSheetName, shouldLog?) -> Object
        Finds the billing sheet that covers a given payment date. Returns an object
        with the sheet and its header map.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getCumulativeHistory(studentId) -> Object | null
        Reads the Cumulative Tracking sheet and returns the most recent cumulative
        history object for a student, skipping rows with 0/0 hours.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    getCurrentBillingCycleDates() -> Object
        Returns the start and end dates for the current billing cycle from billing
        metadata.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getCurrentBillingSheet() -> Sheet | null
        Returns the billing sheet for the current semester.
        Category: DATA_LOOKUP
        Local functions used: getCurrentSemesterInfo()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getCurrentRateChartName() -> String | null
        Returns the name of the current rate chart column from the Rates sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getMostRecentRateColumn(),
                                UtilityScriptLibrary.getColumnHeaders()

    getCurrentSemesterFromBillingMetadata(billingSheetName) -> String | null
        Returns the semester name associated with a given billing sheet name from
        the Billing Metadata sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    getCurrentSemesterInfo() -> Object | null
        Returns the current semester name and billing sheet name from Billing Metadata.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getCurrentSemesterRateForLength(lessonLength) -> Number | null
        Returns the current semester's lesson rate for a given lesson length.
        Category: DATA_LOOKUP
        Local functions used: getCurrentRateChartName()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap()

    getDocIdColumnName(docType) -> String
        Returns the billing sheet column name for the document ID of a given document
        type.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    getDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap) -> String
        Returns the stored document ID for a given student and document type from the
        billing sheet.
        Category: DATA_LOOKUP
        Local functions used: getDocIdColumnName()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    getDocumentSelectionHtml() -> String
        Returns the HTML string for the document selection dialog.
        Category: UI
        Local functions used: None
        Utility functions used: None

    getExpandedPrograms(row, context) -> Array
        Reserved for future multi-program support. Currently returns the lesson
        program from the row.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    getFilterDate(context, previousStartDate?) -> Date
        Returns the filter date for form response processing — either the previous
        start date or the context's today date.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    getFormsDataFromContacts() -> Object
        Reads the Contacts workbook's form responses sheet and returns a map of
        student data keyed by student ID.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    getFormSheet(context) -> Sheet
        Returns the form responses sheet for the current semester from the form
        responses workbook.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook()

    getFutureSemesterName(currentSemesterName) -> String | null
        Returns the next semester name from Semester Metadata, if one exists.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    getInvoiceNumber(billingSheet, billingRowIndex) -> String
        Returns the invoice number for a student row in a billing sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    getLessonLengthFromRow(row, get) -> Number
        Determines lesson length (30, 45, or 60) from a billing row's quantity columns.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    getNextMonthName(targetDate) -> String
        Returns the month name for the month following the given date.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames()

    getPreviousSemester(currentSemesterName) -> Object | null
        Returns the previous semester's name and billing sheet name from Semester
        Metadata.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    getRateColumnFromMetadata(semesterName) -> String | null
        Returns the rates column header name for a given semester from the Semester
        Metadata sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: None

    getRateMap(context) -> Object
        Returns a rate map object for the billing context, caching on the context
        object after first load.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.buildRateMapFromSheet()

    getStudentBalancesFromBilling(billingSheet, teacherId) -> Array
        Returns an array of student balance objects for a given teacher from a billing
        sheet.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    getStudentDocumentsFolder(monthName) -> Folder
        Returns or creates the student documents subfolder for the given month.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getGeneratedDocumentsFolder(),
                                UtilityScriptLibrary.debugLog()

    getStudentRegisteredLessonLength(studentId) -> Number | null
        Looks up the lesson length a student is registered for from the current
        billing sheet.
        Category: DATA_LOOKUP
        Local functions used: getCurrentBillingSheet()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    getStudentsNeedingPackets() -> Array
        Returns an array of student objects from the active billing sheet who still
        need registration packets generated.
        Category: DATA_LOOKUP
        Local functions used: extractStudentDataFromBillingRow(), extractBillingDataFromRow()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    handleMonthYearCancel() -> void
        Logs cancellation when the user dismisses the month/year selection dialog.
        Category: UI
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    handleMonthYearSelection(month, year) -> void
        Callback from the month/year dialog. Stores the selection in script properties
        and calls continueAttendanceSheetCreation().
        Category: UI
        Local functions used: continueAttendanceSheetCreation()
        Utility functions used: UtilityScriptLibrary.debugLog()

    identifyWarningStudents(studentBalances) -> Array
        Filters a student balances array to return only students whose balance is at
        or below the warning threshold.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    isHeaderRow(row, dateIdx, lengthIdx) -> Boolean
        Returns true if a row appears to be a header row based on date and length
        column values.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    loadProgramConfig(ss) -> Object
        Reads the Programs List sheet and returns a program configuration object.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    loadReregistrationData() -> Array
        Reads the re-registration sheet and returns all pending re-registration rows
        as an array of objects.
        Category: REREGISTRATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    locateStudentRecord(rowData, paymentSheet, billingInfo) -> Object | null
        Locates a student's billing row from a payment sheet row using student ID
        or name matching.
        Category: RECONCILIATION
        Local functions used: locateStudentRecordEnhanced()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.getHeaderMap()

    locateStudentRecordEnhanced(rowData, billingInfo, studentIdInput, invoiceNumberInput, lastNameCol, firstNameCol) -> Object | null
        Enhanced student record lookup that tries student ID, then invoice number,
        then name matching to find a billing row.
        Category: RECONCILIATION
        Local functions used: findBillingRowByStudentId()
        Utility functions used: UtilityScriptLibrary.debugLog()

    logMysteryStudents(mysteryStudents) -> void
        Logs an array of unmatched student records to the debug sheet.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    markReregistrationProcessed(reregSheet, processedIds) -> void
        Marks re-registration rows as processed in the re-registration sheet.
        Category: REREGISTRATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    markStudentsInactive() -> Object
        Marks students who did not re-register as inactive in the Contacts workbook.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    migrateTeacherDisplayNamesToIds() -> void
        One-time migration utility to replace teacher display names with teacher IDs
        in billing sheets.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    onOpen() -> void
        Installs the QAMP Tools menu in the spreadsheet UI on open.
        Category: UI
        Local functions used: None
        Utility functions used: None

    parseMonthYear(monthStr) -> Object
        Parses a "Month YYYY" string into an object with month name and year number.
        Category: HELPERS
        Local functions used: None
        Utility functions used: None

    populateAllCumulativeColumns() -> void
        UI-triggered utility to populate cumulative tracking columns across all
        billing sheets for all students.
        Category: HELPERS
        Local functions used: updateCumulativeTracking()
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateBillingSheet(context, carryOverData) -> void
        Main function to populate a new billing sheet. Handles both new and carryover
        students from the previous billing cycle and form responses.
        Category: BILLING_CYCLE
        Local functions used: populateBillingSheetContinuingSemester(), appendStudentsFromContext()
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateBillingSheetContinuingSemester(context, billingSheet, existingStudentIds, previousData, previousStartDate) -> void
        Populates the billing sheet with carryover students from the previous semester.
        Category: BILLING_CYCLE
        Local functions used: buildCarryoverStudents(), writeAndFormatRows()
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateCurrentBalanceFormula(newRow, context, rowIndex) -> void
        Adds a formula to the Current Balance column of a billing row that subtracts
        payments from the invoice total.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.columnToLetter()

    populateDeliveryPreference(newRow, formRow, context) -> void
        Sets the delivery preference in a new billing row from a form response row.
        Category: BILLING_ROW_BUILDERS
        Local functions used: setDeliveryPreference()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    populateDeliveryPreferenceFromPrevious(newRow, prevRow, context) -> void
        Sets the delivery preference in a new billing row from a previous billing row.
        Category: BILLING_ROW_BUILDERS
        Local functions used: setDeliveryPreference()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    populateInvoiceMetadata(newRow, studentId, context, rowIndex) -> void
        Populates invoice number, invoice date, and billing cycle name columns in a
        billing row.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.formatDateFlexible()

    populateLateFee(newRow, context, currencyCols) -> void
        Populates the late fee column in a billing row if the student qualifies.
        Category: BILLING_ROW_BUILDERS
        Local functions used: applyLateFeeToRow()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    populateLetterType(newRow, context, sourceType, prevRow?) -> void
        Determines and sets the letter type for a billing row based on student
        history and registration source.
        Category: BILLING_ROW_BUILDERS
        Local functions used: determineIfNewStudent()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    processDocumentSelection(selectedTypes) -> void
        Callback from the document selection dialog. Stores the selection and calls
        continuePacketGeneration().
        Category: DOCUMENT_GENERATION
        Local functions used: continuePacketGeneration()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processFieldMapForSemester(semesterName) -> Object
        Updates the field map sheet for the given semester's form responses sheet.
        Category: SEMESTER_SETUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.updateFieldMappings(),
                                UtilityScriptLibrary.debugLog()

    processFormsData(paymentSheet, studentsSheet) -> Object
        Processes form response data for reconciliation, comparing against the
        students sheet for agreement and media release status.
        Category: RECONCILIATION
        Local functions used: processFormsReconciliationForRow()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader()

    processFormsReconciliationForRow(rowData, studentsSheet, studentIdInput, hasAgreement, hasMediaResponse, studentIdCol, agreementCol, mediaCol) -> Object
        Processes a single form response row for reconciliation, checking agreement
        and media release status.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processPaymentReconciliationForRow(rowData, paymentSheet, rowNumber, studentIdInput, invoiceNumberInput, amountCol, dateCol, lastNameCol, firstNameCol) -> Object
        Processes a single payment row for reconciliation against billing data.
        Category: RECONCILIATION
        Local functions used: locateStudentRecord(), reconcilePayment()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processTeacherAttendanceForBilling(teacherSS, targetDate) -> Object
        Reads all attendance sheets in a teacher workbook and collects lesson hours
        per student per month up to the target date. Returns studentHoursByMonth
        and studentDataCurrentThrough.
        Category: RECONCILIATION
        Local functions used: isHeaderRow()
        Utility functions used: UtilityScriptLibrary.getMonthNames(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    processTeacherForNewAttendance(teacherId, rosterUrl, targetMonthName, targetYear) -> Object
        Creates or updates attendance sheets for a single teacher for the target
        month. Returns a result object indicating created, updated, or skipped.
        Category: ATTENDANCE
        Local functions used: extractRosterDataForAttendance(), addMissingStudentsToAttendanceSheet()
        Utility functions used: UtilityScriptLibrary.findRosterSheetForMonth(),
                                UtilityScriptLibrary.createMonthlyAttendanceSheet(),
                                UtilityScriptLibrary.debugLog()

    processTeacherReconciliation(teacherId, rosterUrl, currentBillingSheet, targetDate, currentSemester) -> Object
        Runs a full reconciliation for a single teacher: collects attendance, updates
        billing hours, updates roster balances, and applies warnings.
        Category: RECONCILIATION
        Local functions used: processTeacherAttendanceForBilling(), updateBillingForStudents(),
                              updateTeacherRosterBalances(), getStudentBalancesFromBilling(),
                              identifyWarningStudents(), applyWarningsToTeacherWorkbook()
        Utility functions used: UtilityScriptLibrary.debugLog()

    promptForBillingCycleName(customToday) -> String | null
        Prompts the user to confirm or enter a billing cycle name based on the current
        date. Returns the name or null if cancelled.
        Category: UI
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getMonthNames()

    promptForMonthAndYear() -> void
        Prompts the user to select a month and year via an HTML dialog for attendance
        sheet creation.
        Category: UI
        Local functions used: None
        Utility functions used: None

    promptForSemesterDates() -> Object
        Prompts the user to enter semester start and end dates. Returns an object
        with startDate and endDate or null if cancelled.
        Category: UI
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.promptForDate(), UtilityScriptLibrary.debugLog()

    promptForSemesterName() -> String | null
        Prompts the user to enter a semester name. Returns the name or null if
        cancelled.
        Category: UI
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.promptForNameWithDefault(), UtilityScriptLibrary.debugLog()

    protectBillingSheet(sheet) -> void
        Applies warning-only protection to a billing sheet.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.protectSheetRanges()

    protectPreviousBillingCycle() -> void
        Applies protection to the billing sheet from the previous cycle.
        Category: SHEET_OPERATIONS
        Local functions used: getPreviousSemester(), protectBillingSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    reconcilePayment(paymentSheet, billingSheet, billingRowIndex, paymentRowNumber, invoiceValue, totalPayments) -> Object
        Writes a reconciled payment amount to the billing sheet and marks the payment
        row as reconciled. Returns a result object.
        Category: RECONCILIATION
        Local functions used: sumPayments()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    renameLatestFormSheet(semesterName) -> Object
        Renames the most recent 'Form Responses' sheet to the semester name. Adds
        Teacher, Student ID, and Parent ID columns if missing. Returns a result object.
        Category: SEMESTER_SETUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.executeWithErrorHandling(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    runBillingCycleAutomation() -> void
        Main entry point for a full billing cycle run. Prompts for date, semester,
        and cycle name, then creates the billing sheet and populates it.
        Category: BILLING_CYCLE
        Local functions used: buildBillingContext(), loadProgramConfig(), createBillingSheet(),
                              populateBillingSheet(), finalizeBillingSheet(),
                              appendToBillingMetadata(), protectPreviousBillingCycle()
        Utility functions used: UtilityScriptLibrary.promptForCustomToday(),
                                UtilityScriptLibrary.getCurrentSemesterName(),
                                UtilityScriptLibrary.debugLog()

    runFullReconciliation(reconciliationDate?) -> Object
        Runs a full lesson and payment reconciliation across all active teachers.
        Returns a results object with per-teacher stats.
        Category: RECONCILIATION
        Local functions used: processTeacherReconciliation(), updateLastReconciliationDate(),
                              generateReconciliationSummary()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    runFullReconciliationUI() -> void
        UI wrapper for runFullReconciliation(). Prompts for date and displays results.
        Category: UI
        Local functions used: runFullReconciliation()
        Utility functions used: UtilityScriptLibrary.promptForCustomToday(), UtilityScriptLibrary.debugLog()

    runLogHeaders() -> void
        One-liner wrapper calling UtilityScriptLibrary.logAllSheetHeaders() for
        debugging.
        Category: HELPERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.logAllSheetHeaders()

    runPaymentsReconciliation() -> Object
        Runs payment reconciliation by comparing the Payments sheet against billing
        data. Returns a results object.
        Category: RECONCILIATION
        Local functions used: processPaymentReconciliationForRow(), processFormsData(),
                              verifyAndGetParentData()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    runPaymentsReconciliationUI() -> void
        UI wrapper for runPaymentsReconciliation(). Displays a results alert.
        Category: UI
        Local functions used: runPaymentsReconciliation()
        Utility functions used: UtilityScriptLibrary.debugLog()

    runRegistrationPacketGenerationUI() -> void
        UI entry point that calls generateRegistrationPacketsForBillingCycle().
        Category: UI
        Local functions used: generateRegistrationPacketsForBillingCycle()
        Utility functions used: None

    runWeeklyLessonReconciliation(customDate?) -> Object
        Runs weekly lesson reconciliation for all active teachers up to the given
        date. Returns a results object.
        Category: RECONCILIATION
        Local functions used: processTeacherReconciliation(), updateLastReconciliationDate(),
                              generateReconciliationSummary()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    runWeeklyLessonReconciliationUI() -> void
        UI wrapper for runWeeklyLessonReconciliation(). Prompts for date and displays
        results.
        Category: UI
        Local functions used: runWeeklyLessonReconciliation()
        Utility functions used: UtilityScriptLibrary.promptForCustomToday(), UtilityScriptLibrary.debugLog()

    selectDocumentTemplate(templateType, studentData, deliveryMethod, currentSemester, billingData) -> String
        Returns the template key for a given document type, delivery method, and
        student context.
        Category: DOCUMENT_GENERATION
        Local functions used: determineIfNewStudent()
        Utility functions used: None

    sendReregistrationLinks() -> void
        Sends re-registration email links to parents of active students via the
        re-registration queue.
        Category: REREGISTRATION
        Local functions used: buildReregistrationQueue(), verifyAndGetParentData()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.sendEmail(), UtilityScriptLibrary.debugLog()

    setDeliveryPreference(newRow, sourceRow, context, sourceIsForm) -> void
        Sets the delivery preference column in a new billing row from either a form
        row or a previous billing row.
        Category: BILLING_ROW_BUILDERS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    setPastColumnFormulas(newRow, context, rowNum, currencyCols) -> void
        Adds VLOOKUP formulas to past balance and payment columns in a new billing
        row, pulling from the previous cycle sheet.
        Category: BILLING_ROW_BUILDERS
        Local functions used: buildPastVlookupFormula()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    setProgramQuantitiesForCarryover(newRow, context, quantityCols, currencyCols) -> void
        Sets lesson quantity columns in a carryover billing row based on the student's
        registered lesson length and rate.
        Category: BILLING_ROW_BUILDERS
        Local functions used: getRateMap()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    setupNewSemester() -> void
        UI entry point to set up a new semester: prompts for name and dates, creates
        roster folder, billing sheet, payments tab, calendar, and field map.
        Category: SEMESTER_SETUP
        Local functions used: promptForSemesterName(), promptForSemesterDates(),
                              createRosterFolder(), createBillingSheet(), createPaymentsTab(),
                              generateCalendarForSemester(), processFieldMapForSemester(),
                              appendToSemesterMetadata(), renameLatestFormSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    shouldGenerateInvoice(billingData) -> Boolean
        Returns true if a student should receive an invoice based on their current
        balance and program enrollment.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    shouldIncludeAgreement(studentId, billingSheet, headerMap) -> Boolean
        Returns true if a student still needs to return their parent agreement form.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    shouldIncludeDocument(studentId, docType, billingSheet, headerMap) -> Boolean
        Returns true if a student still needs a given document type.
        Category: DOCUMENT_GENERATION
        Local functions used: checkIfDocumentStillNeeded()
        Utility functions used: None

    shouldIncludeMediaRelease(studentId, billingSheet, headerMap) -> Boolean
        Returns true if a student still needs to return their media release form.
        Category: DOCUMENT_GENERATION
        Local functions used: checkIfMediaReleaseNeeded()
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    shouldUseMissingDocumentLetter(billingData) -> Boolean
        Returns true if a missing document letter should be used instead of a standard
        welcome letter based on billing data.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    showSimpleDocumentSelectionDialog() -> void
        Displays an HTML dialog for selecting which document types to include in a
        registration packet run.
        Category: UI
        Local functions used: getDocumentSelectionHtml()
        Utility functions used: None

    submitReregistration(data) -> Object
        Web app POST handler for re-registration form submissions. Validates parent
        ID, finds the student's billing row, and updates re-registration queue.
        Category: UI
        Local functions used: verifyAndGetParentData(), loadReregistrationData()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    sumPayments(sheet, studentId, startDate, endDate) -> Number
        Sums all payments for a student in a payments sheet between the start and end
        dates.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    updateBillingForStudents(billingSheet, studentHours) -> void
        Updates lesson hours and invoice date columns in a billing sheet from a
        studentHours map collected during reconciliation.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    updateCumulativeTracking(billingSheetName) -> void
        Updates the Cumulative Tracking sheet with current cumTaught, cumBilled, and
        balance values from a billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    updateDocIdInBillingSheet(studentId, docType, fileId, docUrl, billingSheet, headerMap) -> void
        Writes a document ID and URL to the appropriate columns in a billing sheet row.
        Category: DOCUMENT_GENERATION
        Local functions used: getDocIdColumnName()
        Utility functions used: UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    updateInvoiceUrlInBillingSheet(billingSheet, rowNumber, headerMap, url) -> void
        Writes an invoice URL to the Invoice URL column of a billing row.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    updateLastReconciliationDate(billingSheet, studentDataCurrentThrough) -> void
        Updates the Last Reconciliation Date column for each student in a billing
        sheet from the studentDataCurrentThrough map.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.formatDateFlexible()

    updateSheetStudentWarnings(sheet, warningStudentMap, formattedDate) -> void
        Writes low-balance warning indicators to a single attendance sheet for
        students in the warning map.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    updateTeacherRosterBalances(teacherSS, studentBalances, currentSemester) -> void
        Updates the lesson balance column on the current semester roster sheet in a
        teacher workbook.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.extractSeasonFromSemester(),
                                UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    verifyAndGetParentData(parentId, lastFourPhone) -> Object
        Verifies a parent's identity by matching parent ID and last four digits of
        phone against the Contacts workbook. Returns parent data if verified.
        Category: REREGISTRATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    verifyPaymentsDetailed() -> void
        UI entry point that prompts for student IDs and calls
        verifyPaymentsDetailedForStudents().
        Category: UI
        Local functions used: verifyPaymentsDetailedForStudents()
        Utility functions used: None

    verifyPaymentsDetailedForStudents(studentIds) -> void
        Generates detailed payment verification reports for a list of student IDs
        by comparing billing and payment data.
        Category: VERIFICATION
        Local functions used: collectBillingDataDetailed(), collectPaymentsDataDetailed(),
                              createDetailedPaymentReport(), createPaymentVerificationReport()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog()

    verifyProgramsForSemester(semesterName, billingSS) -> Object
        Verifies that all programs in the billing sheet are configured in the Programs
        List for the given semester. Returns a results object.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    verifyRatesEnhanced() -> void
        UI-triggered function to verify that all rates are properly configured for
        the current semester. Logs detailed results.
        Category: VERIFICATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.debugLog()

    writeAndFormatRows(billingSheet, allStudents) -> void
        Batch-writes all student rows to the billing sheet and applies formatting to
        each row.
        Category: BILLING_CYCLE
        Local functions used: formatRow()
        Utility functions used: UtilityScriptLibrary.debugLog()

  ================================================================================
  FUNCTIONS BY CATEGORY
  ================================================================================

  ATTENDANCE (5 functions):
    addMissingStudentsToAttendanceSheet
    continueAttendanceSheetCreation
    expandSheetAttendanceRows
    expandTeacherAttendanceRows
    processTeacherForNewAttendance

  BILLING_CYCLE (13 functions):
    addLateRegistrationsToBillingCycle
    appendStudentsFromContext
    applyMultiStudentDiscount
    buildBillingContext
    buildCarryoverStudents
    buildFormStudents
    buildFutureSemesterContext
    finalizeBillingSheet
    populateBillingSheet
    populateBillingSheetContinuingSemester
    runBillingCycleAutomation
    updateCumulativeTracking
    writeAndFormatRows

  BILLING_ROW_BUILDERS (19 functions):
    addInvoiceTotalFormula
    applyCumulativeHistory
    applyLateFeeToRow
    buildBillingRowFromForm
    buildBillingRowFromPrevious
    buildDynamicProgramColumns
    buildInvoiceTotalFormula
    buildPastVlookupFormula
    copyStaticFieldsToBillingRow
    generateProgramFormulas
    populateCurrentBalanceFormula
    populateDeliveryPreference
    populateDeliveryPreferenceFromPrevious
    populateInvoiceMetadata
    populateLateFee
    populateLetterType
    setDeliveryPreference
    setPastColumnFormulas
    setProgramQuantitiesForCarryover

  DATA_EXTRACTION (7 functions):
    extractBillingDataFromRow
    extractDeliveryPreference
    extractDocumentNames
    extractPreviousBillingData
    extractProgramTotals
    extractRosterDataForAttendance
    extractStudentDataFromBillingRow

  DATA_LOOKUP (22 functions):
    findBillingRowByStudentId
    getActivePrograms
    getBillingSheet
    getCumulativeHistory
    getCurrentBillingCycleDates
    getCurrentBillingSheet
    getCurrentRateChartName
    getCurrentSemesterFromBillingMetadata
    getCurrentSemesterInfo
    getCurrentSemesterRateForLength
    getDocIdFromBillingSheet
    getFormsDataFromContacts
    getFormSheet
    getFutureSemesterName
    getInvoiceNumber
    getPreviousSemester
    getRateColumnFromMetadata
    getRateMap
    getStudentBalancesFromBilling
    getStudentRegisteredLessonLength
    getStudentsNeedingPackets
    loadProgramConfig

  DOCUMENT_GENERATION (30 functions):
    buildDocumentFileName
    buildDocumentSentence
    buildDynamicAmounts
    buildDynamicLineItems
    buildInvoiceVariableMap
    buildMissingDocumentSentence
    buildProgramDescription
    buildTemplateVariables
    cancelDocumentGeneration
    checkIfDocumentStillNeeded
    checkIfMediaReleaseNeeded
    clearDocIdFromBillingSheet
    continuePacketGeneration
    convertFolderDocsToPdfUI
    createSingleRegistrationPacketWithSelection
    executeDocumentGeneration
    generateDocumentForStudent
    generateInvoiceForStudent
    generateInvoicesForBillingCycle
    generateRefundInvoicesForBillingCycle
    generateRegistrationPacketForStudentWithSelection
    processDocumentSelection
    selectDocumentTemplate
    shouldGenerateInvoice
    shouldIncludeAgreement
    shouldIncludeDocument
    shouldIncludeMediaRelease
    shouldUseMissingDocumentLetter
    updateDocIdInBillingSheet
    updateInvoiceUrlInBillingSheet

  FORMATTING (3 functions):
    applyBillingConditionalFormatting
    applyLetterTypeValidation
    formatRow

  HELPERS (18 functions):
    backfillCumulativeTracking
    calculateTotalCreditsApplied
    createRosterFolder
    determineIfNewStudent
    determinePacketVersions
    getDocIdColumnName
    getExpandedPrograms
    getFilterDate
    getNextMonthName
    getLessonLengthFromRow
    getStudentDocumentsFolder
    isHeaderRow
    logMysteryStudents
    markStudentsInactive
    migrateTeacherDisplayNamesToIds
    parseMonthYear
    populateAllCumulativeColumns
    runLogHeaders

  METADATA (2 functions):
    appendToBillingMetadata
    appendToSemesterMetadata

  RECONCILIATION (19 functions):
    applyWarningsToTeacherWorkbook
    generateReconciliationSummary
    identifyWarningStudents
    locateStudentRecord
    locateStudentRecordEnhanced
    processFormsData
    processFormsReconciliationForRow
    processPaymentReconciliationForRow
    processTeacherAttendanceForBilling
    processTeacherReconciliation
    reconcilePayment
    runFullReconciliation
    runPaymentsReconciliation
    runWeeklyLessonReconciliation
    sumPayments
    updateBillingForStudents
    updateLastReconciliationDate
    updateSheetStudentWarnings
    updateTeacherRosterBalances

  REREGISTRATION (7 functions):
    appendReregistrationNewStudents
    applyReregistrationOverwrites
    buildReregistrationQueue
    loadReregistrationData
    markReregistrationProcessed
    sendReregistrationLinks
    verifyAndGetParentData

  SEMESTER_SETUP (4 functions):
    generateCalendarForSemester
    processFieldMapForSemester
    renameLatestFormSheet
    setupNewSemester

  SHEET_OPERATIONS (4 functions):
    createBillingSheet
    createPaymentsTab
    protectBillingSheet
    protectPreviousBillingCycle

  UI (20 functions):
    addLateRegistrationsUI
    createNewAttendanceSheets
    doGet
    expandTeacherAttendanceSheets
    generateRegistrationPacketsForBillingCycle
    getDocumentSelectionHtml
    handleMonthYearCancel
    handleMonthYearSelection
    onOpen
    promptForBillingCycleName
    promptForMonthAndYear
    promptForSemesterDates
    promptForSemesterName
    runFullReconciliationUI
    runPaymentsReconciliationUI
    runRegistrationPacketGenerationUI
    runWeeklyLessonReconciliationUI
    showSimpleDocumentSelectionDialog
    submitReregistration
    verifyPaymentsDetailed

  VERIFICATION (7 functions):
    collectBillingDataDetailed
    collectPaymentsDataDetailed
    createDetailedPaymentReport
    createPaymentVerificationReport
    verifyPaymentsDetailedForStudents
    verifyProgramsForSemester
    verifyRatesEnhanced

================================================================================
END OF FUNCTION DIRECTORY
================================================================================