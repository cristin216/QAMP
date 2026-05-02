================================================================================
BILLING FUNCTION DIRECTORY
================================================================================
    Total Functions: 196
    Most Recent version: 164

    This directory provides a quick reference for all functions in Billing script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: function1(), function2()

  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    addCumulativeFormulas
    addInvoiceTotalFormula
    addMissingStudentsToAttendanceSheet
    appendReregistrationNewStudents
    appendToBillingMetadata
    appendToSemesterMetadata
    applyAdminVisualFormatting
    applyCumulativeHistory
    applyLateFeeToRow
    applyLessonEquivalentCredits
    applyLetterTypeValidation
    applyMultiStudentDiscount
    applyPastDataToRow
    applyReregistrationOverwrites
    applyWarningsToTeacherWorkbook
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
    buildInvoiceTotalFormula
    buildInvoiceVariableMap
    buildMissingDocumentSentence
    buildProgramDescription
    buildTemplateVariables
    calculateLateFee
    calculateLessonEquivalents
    calculateTotalCreditsApplied
    cancelDocumentGeneration
    checkIfDocumentStillNeeded
    checkIfMediaReleaseNeeded
    checkRowFormatting
    clearDocIdFromBillingSheet
    collectBillingData
    collectBillingDataDetailed
    collectPaymentsData
    collectPaymentsDataDetailed
    continueAttendanceSheetCreation
    continuePacketGeneration
    convertFolderDocsToPdfUI
    copyStaticFieldsToBillingRow
    createBillingSheet
    createDetailedPaymentReport
    createNewAttendanceSheets
    createPaymentVerificationReport
    createPaymentsTab
    createRosterFolder
    createSingleRegistrationPacketWithSelection
    debugBillingSheetColumns
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
    extractNumericLessonLength
    extractPreviousBillingData
    extractProgramTotals
    extractRosterDataForAttendance
    extractStudentDataFromBillingRow
    findBillingRowByStudentId
    findMostRecentRosterSheet
    formatRow
    generateCalendarForSemester
    generateDocumentForStudent
    generateInvoiceForStudent
    generateInvoicesForBillingCycle
    generateProgramFormulas
    generateReconciliationSummary
    generateReconciliationSummaryUpdated
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
    getFormSheet
    getFormsDataFromContacts
    getInvoiceNumber
    getLessonLengthFromRow
    getNextMonthName
    getPreviousSemester
    getPreviousSemesterBalance
    getRateColumnFromMetadata
    getRateMap
    getStudentBalancesFromBilling
    getStudentDocumentsFolder
    getStudentRegisteredLessonLength
    getStudentsNeedingPackets
    grantDocumentPermissions
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
    populatePastBalanceAndCredit
    processDocumentSelection
    processFieldMapForSemester
    processFormsData
    processFormsReconciliationForRow
    processPaymentReconciliationForRow
    processPaymentRecord
    processSemesterEndCredits
    processTeacherAttendanceForBilling
    processTeacherForNewAttendance
    processTeacherReconciliation
    promptForBillingCycleName
    promptForCustomToday
    promptForMonthAndYear
    promptForSemesterDates
    promptForSemesterName
    protectBillingSheet
    protectPreviousBillingCycle
    reconcilePayment
    renameLatestFormSheet
    runBillingCycleAutomation
    runCombinedReconciliation
    runFullReconciliation
    runFullReconciliationUI
    runPaymentReconciliation
    runPaymentReconciliationUI
    runRegistrationPacketGenerationUI
    runWeeklyLessonReconciliation
    runWeeklyLessonReconciliationUI
    selectDocumentTemplate
    sendReregistrationLinks
    setDeliveryPreference
    setProgramQuantitiesForCarryover
    setupNewSemester
    setupRosterTemplateProtection
    shouldGenerateInvoice
    shouldIncludeAgreement
    shouldIncludeDocument
    shouldIncludeMediaRelease
    shouldUseMissingDocumentLetter
    showSimpleDocumentSelectionDialog
    storeSemesterEndBalances
    submitReregistration
    sumPayments
    testDynamicInvoiceFunctions
    testDynamicInvoiceFunctionsOnCorrectSheet
    testExtractStudentDataFromBillingRow
    testFormatDateFunction
    testFullAgreementGeneration
    testPrompt
    testTemplateLiteral
    updateBillingForTeacherStudents
    updateCumulativeTracking
    updateDocIdInBillingSheet
    updateInvoiceUrlInBillingSheet
    updateSheetStudentWarnings
    updateTeacherRosterBalances
    verifyAndGetParentData
    verifyCumulativeFormulas
    verifyPaymentsDetailed
    verifyPaymentsDetailedForStudents
    verifyProgramsForSemester
    verifyRatesEnhanced
    writeAndFormatRows  ================================================================================
  FUNCTION CATEGORIES:
  ================================================================================

    UI_MENU (14 functions):
      getNextMonthName
      handleMonthYearCancel
      handleMonthYearSelection
      onOpen
      promptForMonthAndYear
      runBillingCycleAutomation
      runFullReconciliation
      runFullReconciliationUI
      runPaymentReconciliation
      runPaymentReconciliationUI
      runRegistrationPacketGenerationUI
      runWeeklyLessonReconciliationUI
      setupNewSemester
      setupRosterTemplateProtection
      verifyPaymentsDetailed

    SETUP_SEMESTER (8 functions):
      appendToSemesterMetadata
      createPaymentsTab
      createRosterFolder
      generateCalendarForSemester
      markStudentsInactive
      processFieldMapForSemester
      renameLatestFormSheet
      verifyProgramsForSemester

    BILLING_CYCLE (41 functions):
      addCumulativeFormulas
      addInvoiceTotalFormula
      appendToBillingMetadata
      applyCumulativeHistory
      applyLateFeeToRow
      applyLessonEquivalentCredits
      applyLetterTypeValidation
      applyMultiStudentDiscount
      applyPastDataToRow
      buildBillingContext
      buildBillingRowFromForm
      buildBillingRowFromPrevious
      buildCarryoverStudents
      buildDynamicAmounts
      buildDynamicLineItems
      buildDynamicProgramColumns
      buildFormStudents
      buildInvoiceTotalFormula
      calculateLateFee
      copyStaticFieldsToBillingRow
      createBillingSheet
      formatRow
      generateProgramFormulas
      populateAllCumulativeColumns
      populateBillingSheet
      populateBillingSheetContinuingSemester
      populateCurrentBalanceFormula
      populateDeliveryPreference
      populateDeliveryPreferenceFromPrevious
      populateInvoiceMetadata
      populateLateFee
      populateLetterType
      populatePastBalanceAndCredit
      processSemesterEndCredits
      protectPreviousBillingCycle
      setDeliveryPreference
      setProgramQuantitiesForCarryover
      updateCumulativeTracking
      writeAndFormatRows

    GENERATE_DOCUMENTS (38 functions):
      addMissingStudentsToAttendanceSheet
      buildDocumentFileName
      buildDocumentSentence
      buildInvoiceVariableMap
      buildMissingDocumentSentence
      buildProgramDescription
      buildTemplateVariables
      cancelDocumentGeneration
      checkIfDocumentStillNeeded
      checkIfMediaReleaseNeeded
      clearDocIdFromBillingSheet
      continueAttendanceSheetCreation
      continuePacketGeneration
      convertFolderDocsToPdfUI
      createSingleRegistrationPacketWithSelection
      determineIfNewStudent
      determinePacketVersions
      executeDocumentGeneration
      extractDocumentNames
      extractRosterDataForAttendance
      generateDocumentForStudent
      generateInvoiceForStudent
      generateInvoicesForBillingCycle
      generateRefundInvoicesForBillingCycle
      generateRegistrationPacketForStudentWithSelection
      generateRegistrationPacketsForBillingCycle
      getDocIdColumnName
      getDocIdFromBillingSheet
      getDocumentSelectionHtml
      processDocumentSelection
      processTeacherForNewAttendance
      selectDocumentTemplate
      shouldGenerateInvoice
      shouldIncludeAgreement
      shouldIncludeDocument
      shouldIncludeMediaRelease
      shouldUseMissingDocumentLetter
      showSimpleDocumentSelectionDialog
      updateDocIdInBillingSheet
      updateInvoiceUrlInBillingSheet

    RECONCILIATION (27 functions):
      applyAdminVisualFormatting
      applyWarningsToTeacherWorkbook
      expandSheetAttendanceRows
      findBillingRowByStudentId
      generateReconciliationSummary
      generateReconciliationSummaryUpdated
      getBillingSheet
      getInvoiceNumber
      getStudentBalancesFromBilling
      identifyWarningStudents
      locateStudentRecord
      locateStudentRecordEnhanced
      logMysteryStudents
      processFormsData
      processFormsReconciliationForRow
      processPaymentReconciliationForRow
      processPaymentRecord
      processTeacherAttendanceForBilling
      processTeacherReconciliation
      reconcilePayment
      runCombinedReconciliation
      runWeeklyLessonReconciliation
      sumPayments
      updateBillingForTeacherStudents
      updateSheetStudentWarnings
      updateTeacherRosterBalances

    RE_REGISTRATION (8 functions):
      appendReregistrationNewStudents
      applyReregistrationOverwrites
      doGet
      loadReregistrationData
      markReregistrationProcessed
      sendReregistrationLinks
      submitReregistration
      verifyAndGetParentData

    HELPER_FUNCTIONS (47 functions):
      applyLessonEquivalentCredits (note: also referenced from BILLING_CYCLE)
      calculateLessonEquivalents
      calculateTotalCreditsApplied
      collectBillingData
      collectBillingDataDetailed
      collectPaymentsData
      collectPaymentsDataDetailed
      createDetailedPaymentReport
      createNewAttendanceSheets
      createPaymentVerificationReport
      expandSheetAttendanceRows
      expandTeacherAttendanceRows
      expandTeacherAttendanceSheets
      extractBillingDataFromRow
      extractDeliveryPreference
      extractNumericLessonLength
      extractPreviousBillingData
      extractProgramTotals
      extractStudentDataFromBillingRow
      findMostRecentRosterSheet
      getActivePrograms
      getCurrentBillingCycleDates
      getCurrentBillingSheet
      getCurrentRateChartName
      getCurrentSemesterFromBillingMetadata
      getCurrentSemesterInfo
      getCurrentSemesterRateForLength
      getCumulativeHistory
      getExpandedPrograms
      getFilterDate
      getFormsDataFromContacts
      getFormSheet
      getLessonLengthFromRow
      getPreviousSemester
      getPreviousSemesterBalance
      getRateColumnFromMetadata
      getRateMap
      getStudentDocumentsFolder
      getStudentRegisteredLessonLength
      getStudentsNeedingPackets
      isHeaderRow
      loadProgramConfig
      parseMonthYear
      promptForBillingCycleName
      promptForCustomToday
      promptForSemesterDates
      promptForSemesterName
      protectBillingSheet
      storeSemesterEndBalances
      verifyPaymentsDetailedForStudents
      verifyRatesEnhanced

    TESTING (11 functions):
      checkRowFormatting
      debugBillingSheetColumns
      grantDocumentPermissions
      testDynamicInvoiceFunctions
      testDynamicInvoiceFunctionsOnCorrectSheet
      testExtractStudentDataFromBillingRow
      testFormatDateFunction
      testFullAgreementGeneration
      testPrompt
      testTemplateLiteral
      verifyCumulativeFormulas
      
  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addCumulativeFormulas(newRow, context, rowIndex, quantityCols) -> void
        Writes cumulative tracking formulas into the appropriate columns of a new billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: normalizeHeader(), columnToLetter()

    addInvoiceTotalFormula(newRow, context, rowIndex, currencyCols) -> void
        Writes the Current Invoice Total formula into the new billing row.
        Category: BILLING_CYCLE
        Local functions used: buildInvoiceTotalFormula()
        Utility functions used: normalizeHeader(), addToCurrencyCols()

    addMissingStudentsToAttendanceSheet(sheet, activeStudents) -> void
        Adds any active students missing from an attendance sheet.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    appendReregistrationNewStudents(billingSheet, reregMap, overwroteIds, context) -> void
        Appends to the billing sheet any re-registered students not already present.
        Category: RE_REGISTRATION
        Local functions used: buildBillingRowFromForm()
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    appendToBillingMetadata(billingCycleName, billingDate, semesterName) -> void
        Appends new billing cycle information to Billing Metadata sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    appendToSemesterMetadata(semesterName, startDate, endDate, fieldMap) -> void
        Appends new semester information to Semester Metadata sheet.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    applyAdminVisualFormatting(sheet, dataRowCount) -> void
        Applies visual formatting to reconciliation sheets for admin viewing.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    applyCumulativeHistory(newRow, studentId, context, currencyCols) -> void
        Applies historical cumulative hours and billed data from Cumulative Tracking sheet to billing row.
        Category: BILLING_CYCLE
        Local functions used: getCumulativeHistory()
        Utility functions used: normalizeHeader(), debugLog()

    applyLateFeeToRow(newRow, context, currencyCols) -> void
        Calculates and writes late fee into the appropriate column of a new billing row.
        Category: BILLING_CYCLE
        Local functions used: calculateLateFee()
        Utility functions used: normalizeHeader()

    applyLessonEquivalentCredits(billingContext, studentData, row, billingSheet) -> void
        Applies lesson equivalent credits for students switching from packages to private lessons.
        Category: BILLING_CYCLE
        Local functions used: calculateLessonEquivalents()
        Utility functions used: getHeaderMap(), debugLog()

    applyLetterTypeValidation(sheet) -> void
        Applies data validation dropdown to Letter Type column in billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    applyMultiStudentDiscount(billingSheet) -> void
        Applies multi-student family discount to all eligible rows in the billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: normalizeHeader(), getHeaderMap(), debugLog()

    applyPastDataToRow(newRow, prevRow, context, currencyCols) -> void
        Copies past-cycle balance and invoice data from the previous billing row into the new row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: normalizeHeader()

    applyReregistrationOverwrites(billingSheet, reregMap, formStudentIds) -> void
        Overwrites lesson quantity, hours, and length for existing students who re-registered.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    applyWarningsToTeacherWorkbook(teacherWorkbook, warningStudents) -> void
        Highlights warning students in teacher's attendance sheets with red rows.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    buildBillingContext(customToday, semesterName, billingCycleName) -> Object
        Builds comprehensive context object containing all billing cycle information.
        Returns object with dates, rates, semester info, and configuration.
        Category: BILLING_CYCLE
        Local functions used: getCurrentBillingCycleDates(), getCurrentRateChartName(), 
                              getRateColumnFromMetadata()
        Utility functions used: debugLog()

    buildBillingRowFromForm(context, formData, studentId) -> Array
        Constructs billing row data from student form submission for new students.
        Returns array representing complete billing row.
        Category: BILLING_CYCLE
        Local functions used: buildDynamicProgramColumns(), populateDeliveryPreference()
        Utility functions used: getHeaderMap(), debugLog()

    buildBillingRowFromPrevious(context, prevRow, prevHeaderMap, studentId) -> Array
        Constructs billing row data from previous billing cycle for continuing students.
        Returns array representing complete billing row.
        Category: BILLING_CYCLE
        Local functions used: buildDynamicProgramColumns(), populateDeliveryPreferenceFromPrevious()
        Utility functions used: getHeaderMap(), debugLog()

    buildCarryoverStudents(previousData, existingStudentIds, context, allStudents, startingRowIndex) -> Number
        Builds billing rows for all carryover students from the previous billing cycle.
        Returns the next available row index after all carryover rows are added.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromPrevious()
        Utility functions used: debugLog()

    buildDocumentFileName(studentData, docType, newOrReturning, deliveryMode) -> String
        Generates standardized document filename for generated PDFs.
        Returns formatted filename string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildDocumentSentence(documentStatus) -> String
        Builds human-readable sentence describing document generation status.
        Returns formatted status string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildDynamicAmounts(billingData) -> String
        Builds formatted string of dollar amounts for invoice display.
        Returns newline-separated string of amounts.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    buildDynamicLineItems(billingData) -> String
        Builds formatted string of line item descriptions for invoice display.
        Returns newline-separated string of items.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    buildDynamicProgramColumns(context, studentData, programMap?) -> Array
        Builds array of program registration data for billing row.
        Returns array of program values for dynamic columns.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    buildFormStudents(formSheet, filterDate, previousData, context, allStudents, startingRowIndex) -> Number
        Builds billing rows from new form submissions not present in carryover data.
        Returns the next available row index after all form student rows are added.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm()
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    buildInvoiceTotalFormula(row, headerMap) -> String
        Generates Excel formula to calculate invoice total from program columns.
        Returns formula string.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    buildInvoiceVariableMap(studentData, billingData) -> Object
        Creates variable map for invoice template merging.
        Returns object with all invoice template variables.
        Category: GENERATE_DOCUMENTS
        Local functions used: buildDynamicLineItems(), buildDynamicAmounts(), buildProgramDescription()
        Utility functions used: None

    buildMissingDocumentSentence(missingDocuments) -> String
        Builds formatted sentence listing missing required documents.
        Returns formatted document list string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildProgramDescription(programTotals) -> String
        Builds formatted description of student's program registrations.
        Returns formatted program description string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildTemplateVariables(studentData, billingData, docType) -> Object
        Builds complete variable map for all document template types.
        Returns object with all template merge variables.
        Category: GENERATE_DOCUMENTS
        Local functions used: buildInvoiceVariableMap(), buildDocumentSentence(), 
                              buildProgramDescription()
        Utility functions used: debugLog()

    calculateLateFee(pastBalance, paymentReceived, pastInvoiceNumber, rateMap) -> Number
        Calculates late fee amount based on past balance, payments received, and grace period from rateMap.
        Returns late fee amount (0 if not applicable).
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    calculateLessonEquivalents(previousLessons, currentLessonLength) -> Number
        Calculates lesson equivalents when converting between lesson lengths.
        Returns number of equivalent lessons.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    calculateTotalCreditsApplied(billingRow, headerMap) -> Number
        Calculates total credits applied across all programs for a student.
        Returns total credit amount.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    cancelDocumentGeneration() -> void
        Cancels document generation process in progress.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    checkIfDocumentStillNeeded(studentId, docType, billingSheet, headerMap) -> Boolean
        Checks whether a given document type still needs to be generated for a student.
        Returns true if document is still needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: shouldIncludeAgreement(), shouldIncludeMediaRelease()
        Utility functions used: None

    checkIfMediaReleaseNeeded(studentData) -> Boolean
        Determines if media release document is needed for student.
        Returns true if media release needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    checkRowFormatting(sheet, rowNum) -> void
        Debug function to check cell formatting in a specific row.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

    clearDocIdFromBillingSheet(studentId, docType) -> void
        Clears document ID from billing sheet for regeneration.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName(), getCurrentBillingSheet()
        Utility functions used: getHeaderMap(), debugLog()

    collectBillingData(billingWB) -> Object
        Collects billed amounts per student across all billing cycle sheets.
        Returns object mapping studentId to total billed and monthly breakdown.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    collectBillingDataDetailed(billingWB, studentIds) -> Object
        Collects detailed per-month billing data for a specific list of student IDs.
        Returns object mapping studentId → month → amount.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader()

    collectPaymentsData(paymentsWB) -> Object
        Collects payment amounts per student across all payment sheets.
        Returns object mapping studentId to total paid and monthly breakdown.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    collectPaymentsDataDetailed(paymentsWB, studentIds) -> Object
        Collects detailed per-month payment data for a specific list of student IDs.
        Returns object mapping studentId → month → amount.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader()

    continueAttendanceSheetCreation(targetMonthName, targetYear) -> void
        Continues attendance sheet creation after month/year selection from dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: extractRosterDataForAttendance(), processTeacherForNewAttendance()
        Utility functions used: getMonthNames(), debugLog()

    continuePacketGeneration() -> void
        Continues registration packet generation after document selection.
        Category: GENERATE_DOCUMENTS
        Local functions used: processDocumentSelection()
        Utility functions used: None

    convertFolderDocsToPdfUI() -> void
        Converts all Google Docs in the active billing cycle's folder to PDFs, skipping duplicates.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: EnvironmentManager.get(), getConfig()

    copyStaticFieldsToBillingRow(row, headerMap, prevRow, prevHeaderMap) -> void
        Copies unchanging fields from previous billing row to new row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    createBillingSheet(billingCycleName) -> Sheet
        Creates new billing cycle sheet with headers and formatting.
        Returns newly created billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    createDetailedPaymentReport(billingWB, paymentsData, billingData, studentIds) -> void
        Creates a detail sheet in the billing workbook showing month-by-month billed vs. paid for specific students.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    createNewAttendanceSheets() -> void
        Creates new month attendance sheets for all active teachers.
        Category: HELPER_FUNCTIONS
        Local functions used: extractRosterDataForAttendance(), processTeacherForNewAttendance()
        Utility functions used: debugLog()

    createPaymentVerificationReport(billingWB, paymentsData, billingData) -> void
        Creates a Payment Verification summary sheet comparing all billed vs. paid amounts for 2024-2025.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    createPaymentsTab() -> Sheet
        Creates Payments tracking tab in billing workbook.
        Returns newly created payments sheet.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    createRosterFolder(teacherName, semesterName) -> Folder
        Creates Google Drive folder for teacher roster workbook.
        Returns newly created folder.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    createSingleRegistrationPacketWithSelection(studentId, billingCycleName) -> void
        Generates registration packet for single student with document selection dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: showSimpleDocumentSelectionDialog()
        Utility functions used: debugLog()

    debugBillingSheetColumns(sheetName?) -> void
        Logs billing sheet column information for debugging.
        Category: TESTING
        Local functions used: getCurrentBillingSheet()
        Utility functions used: getHeaderMap()

    determineIfNewStudent(studentId, semesterName) -> Boolean
        Determines if student is new to the current semester.
        Returns true if new student.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: debugLog()

    determinePacketVersions(newOrReturning, deliveryMode) -> Object
        Determines which document versions to use based on student status and delivery.
        Returns object with document version selections.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    doGet(e) -> HtmlOutput
        Web app entry point. Serves the ReRegistration HTML page with parent ID from URL parameter.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: None

    executeDocumentGeneration(studentId, billingCycleName, selectedDocs) -> void
        Executes document generation with user-selected documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: generateRegistrationPacketForStudentWithSelection()
        Utility functions used: debugLog()

    expandSheetAttendanceRows(sheet, requiredRows) -> void
        Expands attendance sheet to have minimum required rows for all students.
        Category: RECONCILIATION or HELPER_FUNCTIONS (duplicate in both sections)
        Local functions used: None
        Utility functions used: debugLog()

    expandTeacherAttendanceRows(teacherWorkbook, monthName, studentCount) -> void
        Expands specific month attendance sheet in teacher workbook.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    expandTeacherAttendanceSheets(teacherWorkbook, studentCount) -> void
        Expands all attendance sheets in teacher workbook to accommodate students.
        Category: HELPER_FUNCTIONS
        Local functions used: expandTeacherAttendanceRows()
        Utility functions used: debugLog()

    extractBillingDataFromRow(billingRow, headerMap) -> Object
        Extracts billing-specific data from billing sheet row.
        Returns object with billing amounts and settings.
        Category: HELPER_FUNCTIONS
        Local functions used: extractProgramTotals()
        Utility functions used: debugLog()

    extractDeliveryPreference(contactsSheet, studentId) -> String
        Extracts delivery preference (email/print/both) from Contacts sheet.
        Returns delivery preference string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap()

    extractDocumentNames(selectedCheckboxes) -> Array
        Extracts document names from checkbox selection array.
        Returns array of document name strings.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    extractNumericLessonLength(lengthValue) -> Number
        Extracts a numeric lesson length from a string such as "30 minutes" or "30".
        Returns numeric lesson length, defaulting to 30 if unparseable.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    extractPreviousBillingData(options?) -> Object
        Extracts student billing data from previous billing cycle sheet.
        Returns object with previous billing data and mappings.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    extractProgramTotals(billingRow, headerMap) -> Object
        Extracts program totals from dynamic program columns in billing row.
        Returns object mapping program names to amounts.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    extractRosterDataForAttendance() -> Array
        Extracts teacher and roster data needed for attendance sheet generation.
        Returns array of teacher/roster data objects.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    extractStudentDataFromBillingRow(billingRow, headerMap) -> Object
        Extracts student-specific data from billing sheet row.
        Returns object with student demographic and program data.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    findBillingRowByStudentId(billingData, studentId, studentIdColIndex) -> Array or null
        Finds a billing data row by student ID given raw sheet data and a column index.
        Returns the matching row array, or null if not found.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    findMostRecentRosterSheet(workbook) -> Sheet or null
        Finds the most recent roster sheet in a teacher workbook based on Semester Metadata ordering.
        Searches for sheets ending in " Roster", extracts their season names, matches them against
        Semester Metadata, and returns the roster sheet corresponding to the most recent semester.
        Returns null if no roster sheets exist.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: None
        Utility functions used: getSheet(), normalizeHeader(),
                              extractSeasonFromSemester(), debugLog()

    formatRow(sheet, rowIndex) -> void
        Applies formatting to a specific row in billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    generateCalendarForSemester(semesterName, startDate, endDate) -> void
        Generates semester calendar with holidays and important dates.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    generateDocumentForStudent(studentId, billingCycleName, docType, newOrReturning, deliveryMode) -> String
        Generates single document for student and returns document ID.
        Returns Google Docs document ID.
        Category: GENERATE_DOCUMENTS
        Local functions used: buildTemplateVariables(), selectDocumentTemplate(), 
                              buildDocumentFileName()
        Utility functions used: debugLog()

    generateInvoiceForStudent(studentData, row, headerMap, invoiceOptions, billingSheet, billingSheetName) -> String
        Generates a single invoice document for a student and returns the document ID.
        Returns Google Docs document ID.
        Category: GENERATE_DOCUMENTS
        Local functions used: extractBillingDataFromRow(), extractDeliveryPreference(),
                              buildTemplateVariables(), generateDocumentForStudent()
        Utility functions used: debugLog()

    generateInvoicesForBillingCycle(billingCycleName?) -> void
        Generates invoices for all students in billing cycle.
        Category: GENERATE_DOCUMENTS
        Local functions used: shouldGenerateInvoice(), buildInvoiceVariableMap()
        Utility functions used: debugLog()

    generateProgramFormulas(row, context, billingSheet) -> void
        Generates formulas in program columns for automatic calculation.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    generateReconciliationSummary(reconcileData) -> void
        Generates and displays reconciliation summary in UI.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    generateReconciliationSummaryUpdated(totalStudents, updatedStudents, unchangedStudents) -> void
        Generates updated reconciliation summary with detailed counts.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    generateRefundInvoicesForBillingCycle(billingCycleName?) -> void
        Generates refund invoices for students with negative balances.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: debugLog()

    generateRegistrationPacketForStudentWithSelection(studentId, billingCycleName, selectedDocs) -> void
        Generates registration packet with user-selected documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: determineIfNewStudent(), shouldIncludeDocument(), 
                              generateDocumentForStudent()
        Utility functions used: debugLog()

    generateRegistrationPacketsForBillingCycle(billingCycleName?) -> void
        Generates registration packets for all students needing documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: getStudentsNeedingPackets(), determineIfNewStudent()
        Utility functions used: debugLog()

    getActivePrograms() -> Array
        Gets list of currently active programs from Programs List sheet.
        Returns array of active program names.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getBillingSheet(billingCycleName?) -> Sheet
        Gets billing sheet by name or most recent if no name provided.
        Returns billing sheet object.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    getCumulativeHistory(studentId) -> Object or null
        Retrieves cumulative lesson history for a student from the Cumulative Tracking sheet.
        Returns object with cumulative hours taught and billed, or null if not found.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getCurrentBillingCycleDates(customToday, semesterName) -> Object
        Calculates billing cycle start and end dates based on current date.
        Returns object with startDate and endDate.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getCurrentBillingSheet() -> Sheet
        Gets the most recent billing cycle sheet.
        Returns current billing sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getCurrentRateChartName(semesterName) -> String
        Gets the rate chart name for current semester.
        Returns rate chart name string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getCurrentSemesterFromBillingMetadata() -> String
        Gets current semester name from Billing Metadata sheet.
        Returns semester name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getCurrentSemesterInfo() -> Object
        Gets comprehensive current semester information.
        Returns object with semester name, dates, and configuration.
        Category: HELPER_FUNCTIONS
        Local functions used: getCurrentSemesterName()
        Utility functions used: debugLog()

    getCurrentSemesterRateForLength(lessonLength) -> Number
        Gets current semester rate for specified lesson length.
        Returns rate amount.
        Category: HELPER_FUNCTIONS
        Local functions used: getCurrentSemesterInfo(), getRateForSemester()
        Utility functions used: None

    getDocIdColumnName(docType) -> String
        Gets column name for storing document ID based on document type.
        Returns column name string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    getDocIdFromBillingSheet(studentId, docType) -> String
        Retrieves document ID from billing sheet for specific student and document.
        Returns document ID or null.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName(), getCurrentBillingSheet()
        Utility functions used: getHeaderMap()

    getDocumentSelectionHtml(studentData, newOrReturning, deliveryMode) -> String
        Generates HTML for document selection checkbox dialog.
        Returns HTML string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    getExpandedPrograms() -> Array
        Gets expanded list of programs including group/private variations.
        Returns array of expanded program objects.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getFilterDate(context, previousStartDate?) -> Date or null
        Returns the date to use as a filter cutoff for form submissions, from previousStartDate or Billing Metadata.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getFormSheet(context) -> Sheet
        Gets the semester's form responses sheet from the Form Responses workbook.
        Returns the sheet, or throws if not found.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getWorkbook()

    getFormsDataFromContacts() -> Object
        Retrieves form response data from Contacts sheet.
        Returns object with form data indexed by student ID.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    getInvoiceNumber(billingCycleName) -> String
        Generates sequential invoice number for billing cycle.
        Returns invoice number string.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    getLessonLengthFromRow(row, headerMap) -> String
        Extracts lesson length value from billing row.
        Returns lesson length string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getNextMonthName(currentMonth) -> String
        Gets next month name for semester progression.
        Returns month name string.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    getPreviousSemester(currentSemester) -> String
        Gets previous semester name from current semester.
        Returns previous semester name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getPreviousSemesterBalance(studentId, programName) -> Number
        Gets student's balance for specific program from previous semester.
        Returns balance amount.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getRateColumnFromMetadata(rateChartName) -> Number
        Gets column index for rate chart in Rates sheet.
        Returns column index.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getRateMap(context) -> Object
        Returns a cached rate map from context, or builds and caches it from the Rates sheet.
        Returns object mapping rate keys to values.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getMostRecentRateColumn(), buildRateMapFromSheet()

    getStudentBalancesFromBilling(billingSheet?) -> Object
        Gets all student balances from billing sheet.
        Returns object mapping student IDs to balance data.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getHeaderMap()

    getStudentDocumentsFolder(studentId, semesterName) -> Folder
        Gets or creates student's documents folder in Google Drive.
        Returns student documents folder.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getStudentRegisteredLessonLength(studentId) -> String
        Gets student's registered lesson length from Contacts sheet.
        Returns lesson length string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap()

    getStudentsNeedingPackets(billingCycleName) -> Array
        Gets list of students needing registration packets generated.
        Returns array of student IDs.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    grantDocumentPermissions(docId, email) -> void
        Grants view permissions on generated document to student/parent email.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

    handleMonthYearCancel() -> void
        Logs that the user cancelled the month/year selection dialog.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: debugLog()

    handleMonthYearSelection(month, year) -> void
        Handles the month/year selection from the attendance sheet creation dialog and continues the process.
        Category: UI_MENU
        Local functions used: continueAttendanceSheetCreation()
        Utility functions used: debugLog()

    identifyWarningStudents(teacherBalances) -> Array
        Identifies students with attendance/billing discrepancies needing warnings.
        Returns array of warning student data.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    isHeaderRow(row) -> Boolean
        Checks if row is a header row (contains "Student Name" or similar).
        Returns true if header row.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    loadProgramConfig(ss) -> Object
        Loads and returns program configuration from the Programs List sheet.
        Returns object with programMap and related column indices.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    loadReregistrationData() -> Object or null
        Loads pending re-registration submissions from the Reregistration sheet.
        Returns object mapping studentId to re-registration data, or null if sheet not found.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: getSheet(), getHeaderMap(), normalizeHeader(), debugLog()

    locateStudentRecord(teacherId, studentId) -> Object
        Locates student record in teacher roster or attendance sheets.
        Returns object with sheet and row information.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    locateStudentRecordEnhanced(teacherId, studentId, monthName?) -> Object
        Enhanced version of locateStudentRecord with month-specific search.
        Returns object with detailed location information.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    logMysteryStudents(teacherData) -> void
        Logs students found in teacher attendance but not in billing sheet.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    markReregistrationProcessed(reregSheet, processedIds) -> void
        Marks re-registration sheet rows as processed for the given student IDs.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: normalizeHeader(), debugLog()

    markStudentsInactive(semesterName) -> void
        Marks students as inactive when semester ends.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    onOpen() -> void
        Creates custom menu in spreadsheet UI when workbook opens.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    parseMonthYear(monthStr) -> Object
        Parses a "Month YYYY" string into separate month and year values.
        Returns object with month (String) and year (Number).
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    populateAllCumulativeColumns() -> void
        Populates all cumulative total columns in current billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    populateBillingSheet(context, carryOverData) -> void
        Populates billing sheet for first cycle of new semester.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm(), getFormsDataFromContacts()
        Utility functions used: debugLog()

    populateBillingSheetContinuingSemester(context, billingSheet, existingIds, carryOverData, previousDate?) -> void
        Populates billing sheet for continuing semester billing cycle.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromPrevious(), processSemesterEndCredits()
        Utility functions used: debugLog()

    populateCurrentBalanceFormula(row, headerMap) -> void
        Populates formula calculating current balance in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    populateDeliveryPreference(rowIndex, billingSheet, studentId) -> void
        Populates delivery preference in billing row from Contacts sheet.
        Category: BILLING_CYCLE
        Local functions used: extractDeliveryPreference()
        Utility functions used: getHeaderMap()

    populateDeliveryPreferenceFromPrevious(rowIndex, billingSheet, prevRow, prevHeaderMap) -> void
        Copies delivery preference from previous billing cycle.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateInvoiceMetadata(rowIndex, billingSheet, billingCycleName) -> void
        Populates invoice metadata fields (number, URL, etc.) in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateLateFee(rowIndex, billingSheet) -> void
        Calculates and populates late fee based on past balance.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateLetterType(rowIndex, billingSheet, newOrReturning) -> void
        Populates letter type (New/Returning) in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populatePastBalanceAndCredit(row, headerMap, prevRow, prevHeaderMap) -> void
        Populates past balance and credit from previous billing cycle.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    processDocumentSelection(studentId, billingCycleName) -> void
        Processes document selection from checkbox dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: executeDocumentGeneration()
        Utility functions used: None

    processFieldMapForSemester(semesterName) -> void
        Processes and stores field map for semester form responses.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    processFormsData(formsData, billingData) -> Array
        Processes form submissions and compares against billing data.
        Returns array of form reconciliation results.
        Category: RECONCILIATION
        Local functions used: processFormsReconciliationForRow()
        Utility functions used: debugLog()

    processFormsReconciliationForRow(studentId, formData, billingRow, headerMap) -> Object
        Processes single student's form reconciliation.
        Returns reconciliation result object.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    processPaymentReconciliationForRow(payment, billingSheet, billingData) -> Object
        Processes single payment record against billing data.
        Returns payment reconciliation result.
        Category: RECONCILIATION
        Local functions used: processPaymentRecord()
        Utility functions used: debugLog()

    processPaymentRecord(payment, billingSheet, studentRow, headerMap) -> void
        Processes payment and updates billing sheet.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    processSemesterEndCredits(billingRow, headerMap, previousBalance) -> void
        Processes credits from previous semester end balances.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    processTeacherAttendanceForBilling(teacherId, billingSheet, billingData) -> Object
        Processes teacher's attendance data and reconciles with billing.
        Returns reconciliation results for teacher.
        Category: RECONCILIATION
        Local functions used: locateStudentRecordEnhanced()
        Utility functions used: debugLog()

    processTeacherForNewAttendance(teacherData) -> void
        Processes teacher for new attendance sheet generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: debugLog()

    processTeacherReconciliation(teacherWorkbook, teacherId, billingCycleName, billingSheet) -> Object
        Reconciles teacher's roster and attendance with billing data.
        Returns reconciliation summary for teacher.
        Category: RECONCILIATION
        Local functions used: processTeacherAttendanceForBilling()
        Utility functions used: debugLog()

    promptForBillingCycleName(customToday) -> String
        Prompts user to enter billing cycle name.
        Returns billing cycle name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    promptForCustomToday() -> Date
        Prompts user to enter custom date for billing operations.
        Returns Date object.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    promptForMonthAndYear() -> void
        Shows an HTML dialog for the user to select a month and year for attendance sheet creation.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: getMonthNames()

    promptForSemesterDates() -> Object
        Prompts user to enter semester start and end dates.
        Returns object with dates.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    promptForSemesterName() -> String
        Prompts user to enter semester name.
        Returns semester name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    protectBillingSheet(sheet) -> void
        Protects billing sheet from accidental edits.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    protectPreviousBillingCycle() -> void
        Protects previous billing cycle sheet after new cycle created.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    reconcilePayment(studentRow, headerMap, amount, paymentDate) -> void
        Reconciles single payment against student's billing row.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    renameLatestFormSheet(semesterName) -> void
        Renames most recent form responses sheet to include semester name.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    runBillingCycleAutomation() -> void
        Main function to run complete billing cycle automation process.
        Category: UI_MENU
        Local functions used: promptForCustomToday(), promptForBillingCycleName(), 
                              createBillingSheet(), populateBillingSheet()
        Utility functions used: debugLog()

    runCombinedReconciliation(billingCycleName?) -> void
        Runs combined reconciliation of payments, lessons, and forms.
        Category: RECONCILIATION
        Local functions used: runPaymentReconciliation(), runWeeklyLessonReconciliation(), 
                              runFormsReconciliation()
        Utility functions used: debugLog()

    runFullReconciliation(billingCycleName?) -> void
        Runs complete reconciliation process (payments, lessons, forms).
        Category: UI_MENU
        Local functions used: runPaymentReconciliation(), runWeeklyLessonReconciliation(), 
                              runFormsReconciliation()
        Utility functions used: debugLog()

    runFullReconciliationUI() -> void
        UI wrapper for full reconciliation with user prompts.
        Category: UI_MENU
        Local functions used: runFullReconciliation()
        Utility functions used: None

    runPaymentReconciliation(billingCycleName?) -> Object
        Reconciles payments from Payments sheet with billing data.
        Returns reconciliation results object.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: debugLog()

    runPaymentReconciliationUI() -> void
        UI wrapper for payment reconciliation with user prompts.
        Category: UI_MENU
        Local functions used: runPaymentReconciliation()
        Utility functions used: None

    runRegistrationPacketGenerationUI() -> void
        UI wrapper for registration packet generation with user prompts.
        Category: UI_MENU
        Local functions used: generateRegistrationPacketsForBillingCycle()
        Utility functions used: None

    runWeeklyLessonReconciliation(billingCycleName?) -> Object
        Reconciles teacher attendance records with billing data.
        Returns reconciliation results object.
        Category: RECONCILIATION
        Local functions used: processTeacherReconciliation()
        Utility functions used: debugLog()

    runWeeklyLessonReconciliationUI() -> void
        UI wrapper for lesson reconciliation with user prompts.
        Category: UI_MENU
        Local functions used: runWeeklyLessonReconciliation()
        Utility functions used: None

    selectDocumentTemplate(docType, newOrReturning, deliveryMode) -> String
        Selects appropriate document template ID based on parameters.
        Returns template document ID.
        Category: GENERATE_DOCUMENTS
        Local functions used: determinePacketVersions()
        Utility functions used: None

    sendReregistrationLinks() -> void
        Emails a unique re-registration link to all parents with active students in the
        current billing cycle. Menu-triggered. Confirms count before sending.
        Category: RE_REGISTRATION
        Local functions used: getCurrentBillingSheet()
        Utility functions used: debugLog(), getHeaderMap(), getSheet(), normalizeHeader()

    setDeliveryPreference(newRow, sourceRow, context, sourceIsForm) -> void
        Sets the Delivery Preference field in a new billing row from either a form row or previous billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: normalizeHeader(), getHeaderMap()

    setProgramQuantitiesForCarryover(newRow, context, quantityCols, currencyCols) -> void
        Sets program quantities and prices for carryover students based on the current rate map.
        Category: BILLING_CYCLE
        Local functions used: getRateMap()
        Utility functions used: normalizeHeader()

    setupNewSemester() -> void
        Main function to set up new semester with all required configuration.
        Category: UI_MENU
        Local functions used: promptForSemesterName(), promptForSemesterDates(), 
                              appendToSemesterMetadata()
        Utility functions used: debugLog()

    setupRosterTemplateProtection() -> void
        Sets up protection on roster template sheets.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    shouldGenerateInvoice(billingRow, headerMap) -> Boolean
        Determines if invoice should be generated for student.
        Returns true if invoice needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    shouldIncludeAgreement(studentData) -> Boolean
        Determines if agreement document should be included in packet.
        Returns true if agreement needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    shouldIncludeDocument(docType, newOrReturning, studentData) -> Boolean
        Determines if specific document type should be included in packet.
        Returns true if document needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: shouldIncludeAgreement(), checkIfMediaReleaseNeeded()
        Utility functions used: None

    shouldIncludeMediaRelease(studentData) -> Boolean
        Determines if media release should be included in packet.
        Returns true if media release needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: checkIfMediaReleaseNeeded()
        Utility functions used: None

    shouldUseMissingDocumentLetter(documentStatus) -> Boolean
        Determines if missing document letter should be used instead of welcome letter.
        Returns true if missing documents exist.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    showSimpleDocumentSelectionDialog(studentId, billingCycleName) -> void
        Displays HTML dialog for manual document selection.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocumentSelectionHtml()
        Utility functions used: None

    storeSemesterEndBalances() -> void
        Stores end-of-semester balances for all students.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    submitReregistration(data) -> Object
        Processes re-registration form submission. Updates lesson quantity in current billing
        sheet and updates parent contact info in Contacts if changes were submitted.
        Returns {success, updatedStudents[], errors[], message}
        Category: RE_REGISTRATION
        Local functions used: getCurrentBillingSheet()
        Utility functions used: debugLog(), findParentRow(), generateKey(), getHeaderMap(),
                                getSheet(), normalizeHeader(), updateParentContactFields()

    sumPayments(studentId, throughDate?) -> Number
        Sums all payments for student up to specified date.
        Returns total payment amount.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    testDynamicInvoiceFunctions() -> void
        Tests dynamic invoice line items and amounts functions.
        Category: TESTING
        Local functions used: extractBillingDataFromRow(), buildDynamicLineItems(), 
                              buildDynamicAmounts()
        Utility functions used: None

    testDynamicInvoiceFunctionsOnCorrectSheet() -> void
        Tests dynamic invoice functions on correct billing sheet.
        Category: TESTING
        Local functions used: extractBillingDataFromRow(), buildDynamicLineItems(), 
                              buildDynamicAmounts()
        Utility functions used: None

    testExtractStudentDataFromBillingRow() -> void
        Tests student data extraction from billing row.
        Category: TESTING
        Local functions used: extractStudentDataFromBillingRow()
        Utility functions used: debugLog()

    testFormatDateFunction() -> void
        Tests date formatting functionality.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

    testFullAgreementGeneration() -> void
        Tests complete agreement generation process.
        Category: TESTING
        Local functions used: extractStudentDataFromBillingRow(), extractBillingDataFromRow(), 
                              buildTemplateVariables()
        Utility functions used: debugLog()

    testPrompt() -> void
        Tests user prompt dialogs.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

    testTemplateLiteral() -> void
        Tests template literal string functionality.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

    updateBillingForTeacherStudents(billingSheet, teacherData, billingData) -> Number
        Updates billing sheet with reconciled attendance data for teacher's students.
        Returns count of updated students.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    updateCumulativeTracking(billingSheetName) -> void
        Updates the Cumulative Tracking sheet with totals from the specified billing cycle sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    updateDocIdInBillingSheet(studentId, docType, docId) -> void
        Updates document ID in billing sheet after document generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName(), getCurrentBillingSheet()
        Utility functions used: getHeaderMap(), debugLog()

    updateInvoiceUrlInBillingSheet(studentId, invoiceUrl) -> void
        Updates invoice URL in billing sheet after invoice generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: getCurrentBillingSheet()
        Utility functions used: getHeaderMap()

    updateSheetStudentWarnings(sheet, warningStudents) -> void
        Updates attendance sheet with warning highlights for students.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    updateTeacherRosterBalances(teacherWorkbook, teacherData) -> void
        Updates teacher roster with current balance information.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    verifyAndGetParentData(parentId, lastFourPhone) -> Object
        Verifies parent identity by matching last 4 digits of phone on file.
        Returns parent contact info and associated student records on success.
        Returns {success, parent?, students?, message?}
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: debugLog(), findParentRow(), getHeaderMap(), getSheet(),
                                normalizeHeader()

    verifyCumulativeFormulas() -> Object
        Checks all billing cycle sheets for missing or incorrect cumulative formulas and returns a summary.
        Returns result object from executeWithErrorHandling.
        Category: TESTING
        Local functions used: None
        Utility functions used: executeWithErrorHandling(), debugLog()

    verifyPaymentsDetailed() -> void
        UI entry point: prompts for student IDs and runs detailed payment verification for those students.
        Category: UI_MENU
        Local functions used: verifyPaymentsDetailedForStudents()
        Utility functions used: None

    verifyPaymentsDetailedForStudents(studentIds) -> void
        Runs detailed payment vs. billed comparison for specific students and generates a report sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: collectBillingDataDetailed(), collectPaymentsDataDetailed(),
                              createDetailedPaymentReport()
        Utility functions used: EnvironmentManager.get(), getConfig(), debugLog()

    verifyProgramsForSemester(semesterName) -> Boolean
        Verifies all required programs are set up for semester.
        Returns true if verification passed.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()
        
    verifyRatesEnhanced() -> Boolean
        Verifies all rates are properly configured for billing.
        Returns true if verification passed.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    writeAndFormatRows(billingSheet, allStudents) -> void
        Writes all student rows to the billing sheet in a single batch operation and applies formatting.
        Category: BILLING_CYCLE
        Local functions used: formatRow()
        Utility functions used: None
================================================================================
END OF FUNCTION DIRECTORY
================================================================================    
