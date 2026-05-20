================================================================================
BILLING FUNCTION DIRECTORY
================================================================================
    Total Functions: 182
    Most Recent version: 167

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
    createPaymentVerificationReport
    createPaymentsTab
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
    extractNumericLessonLength
    extractPreviousBillingData
    extractProgramTotals
    extractRosterDataForAttendance
    extractStudentDataFromBillingRow
    finalizeBillingSheet
    findBillingRowByStudentId
    findMostRecentRosterSheet
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
    getFormSheet
    getFormsDataFromContacts
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
    setupRosterTemplateProtection
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
  FUNCTION CATEGORIES:
  ================================================================================

    UI_MENU (16 functions):
      addLateRegistrationsUI
      getNextMonthName
      handleMonthYearCancel
      handleMonthYearSelection
      onOpen
      promptForMonthAndYear
      runBillingCycleAutomation
      runFullReconciliation
      runFullReconciliationUI
      runLogHeaders
      runPaymentsReconciliationUI
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

    BILLING_CYCLE (40 functions):
      addInvoiceTotalFormula
      addLateRegistrationsToBillingCycle
      appendStudentsFromContext
      appendToBillingMetadata
      applyBillingConditionalFormatting
      applyCumulativeHistory
      applyLateFeeToRow
      applyLetterTypeValidation
      applyMultiStudentDiscount
      buildBillingContext
      buildBillingRowFromForm
      buildBillingRowFromPrevious
      buildCarryoverStudents
      buildDynamicAmounts
      buildDynamicLineItems
      buildDynamicProgramColumns
      buildFormStudents
      buildFutureSemesterContext
      buildInvoiceTotalFormula
      buildPastVlookupFormula
      copyStaticFieldsToBillingRow
      createBillingSheet
      finalizeBillingSheet
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
      protectPreviousBillingCycle
      setDeliveryPreference
      setPastColumnFormulas
      setProgramQuantitiesForCarryover
      updateCumulativeTracking
      writeAndFormatRows

    GENERATE_DOCUMENTS (40 functions):
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

    RECONCILIATION (24 functions):
      applyWarningsToTeacherWorkbook
      expandSheetAttendanceRows
      findBillingRowByStudentId
      generateReconciliationSummary
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
      processTeacherAttendanceForBilling
      processTeacherReconciliation
      reconcilePayment
      runPaymentsReconciliation
      runWeeklyLessonReconciliation
      sumPayments
      updateBillingForStudents
      updateLastReconciliationDate
      updateSheetStudentWarnings
      updateTeacherRosterBalances

    RE_REGISTRATION (9 functions):
      appendReregistrationNewStudents
      applyReregistrationOverwrites
      buildReregistrationQueue
      doGet
      loadReregistrationData
      markReregistrationProcessed
      sendReregistrationLinks
      submitReregistration
      verifyAndGetParentData

    HELPER_FUNCTIONS (45 functions):
      calculateTotalCreditsApplied
      collectBillingDataDetailed
      collectPaymentsDataDetailed
      createDetailedPaymentReport
      createNewAttendanceSheets
      createPaymentVerificationReport
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
      getFutureSemesterName
      getLessonLengthFromRow
      getPreviousSemester
      getRateColumnFromMetadata
      getRateMap
      getStudentDocumentsFolder
      getStudentRegisteredLessonLength
      getStudentsNeedingPackets
      isHeaderRow
      loadProgramConfig
      migrateTeacherDisplayNamesToIds
      parseMonthYear
      promptForBillingCycleName
      promptForSemesterDates
      promptForSemesterName
      protectBillingSheet
      verifyPaymentsDetailedForStudents
      verifyRatesEnhanced

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addInvoiceTotalFormula(newRow, context, rowIndex, currencyCols) -> void
        Writes the Current Invoice Total formula into the new billing row.
        Category: BILLING_CYCLE
        Local functions used: buildInvoiceTotalFormula()
        Utility functions used: normalizeHeader(), addToCurrencyCols()

    addLateRegistrationsToBillingCycle() -> Object
        Checks for new form submissions (current and future semester) not yet in the current
        billing cycle and appends them. Returns {added: N} or null if cancelled.
        Handles both the active semester's form sheet and the next semester's form sheet
        if one exists (via getFutureSemesterName / buildFutureSemesterContext).
        Category: BILLING_CYCLE
        Local functions used: getCurrentBillingSheet(), getCurrentSemesterFromBillingMetadata(),
                              loadProgramConfig(), buildBillingContext(), appendStudentsFromContext(),
                              getFutureSemesterName(), buildFutureSemesterContext(),
                              applyMultiStudentDiscount(), populateAllCumulativeColumns(),
                              applyBillingConditionalFormatting()
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    addLateRegistrationsUI() -> void
        UI wrapper for addLateRegistrationsToBillingCycle(). Displays result count or
        "no new registrations" alert, with error handling.
        Category: UI_MENU
        Local functions used: addLateRegistrationsToBillingCycle()
        Utility functions used: debugLog()

    addMissingStudentsToAttendanceSheet(sheet, activeStudents) -> void
        Adds any active students missing from an attendance sheet.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    appendReregistrationNewStudents(billingSheet, reregMap, overwroteIds, context) -> void
        Appends to the billing sheet any re-registered students not already present.
        Writes Package Quantity, Lesson Length, Lesson Quantity, Lesson Price, and Lesson Hours
        directly from reregistration entry data.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    appendStudentsFromContext(billingSheet, context, existingIds) -> Number
        Scans the context's form sheet for student IDs not in existingIds, builds a billing row
        for each via buildBillingRowFromForm(), appends it to billingSheet, and marks the ID
        as seen. Returns count of students appended.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm(), formatRow()
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    appendToBillingMetadata(billingCycleName, customToday, semesterName) -> void
        Appends new billing cycle information to Billing Metadata sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    appendToSemesterMetadata(semesterName, startDate, endDate) -> void
        Appends new semester information to Semester Metadata sheet.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    applyBillingConditionalFormatting(billingSheet) -> void
        Applies conditional formatting rules to Hours Remaining, Current Balance, and Student Last Name
        columns, highlighting rows in warning style when hours are low or balance is negative.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: normalizeHeader(), getHeaderMap(), columnToLetter(), debugLog(), STYLES

    applyCumulativeHistory(newRow, studentId, context, currencyCols) -> void
        Applies historical cumulative hours and billed data from Cumulative Tracking sheet to billing row.
        Category: BILLING_CYCLE
        Local functions used: getCumulativeHistory()
        Utility functions used: normalizeHeader(), debugLog()

    applyLateFeeToRow(newRow, context, currencyCols, pastBalanceValue) -> void
        Calculates and writes late fee into the appropriate column of a new billing row.
        Uses pastBalanceValue (passed in rather than re-read) to determine if fee applies.
        Category: BILLING_CYCLE
        Local functions used: getRateMap()
        Utility functions used: normalizeHeader(), addToCurrencyCols()

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

    applyReregistrationOverwrites(billingSheet, reregMap, formStudentIds) -> void
        Overwrites Package Quantity, Lesson Quantity, Lesson Hours, Lesson Length, Lesson Price,
        and Enrollment for existing students who re-registered.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    applyWarningsToTeacherWorkbook(teacherSS, warningStudents, targetDate) -> void
        Highlights warning students in teacher's attendance sheets with red rows.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    buildBillingContext(customToday, semesterName, billingCycleName, programConfig) -> Object
        Builds comprehensive context object containing all billing cycle information.
        Returns object with dates, rates, semester info, header maps, and program configuration.
        Also includes carryOverForms (Boolean), prevHeaderMap, previousSheetName, reregMap,
        and reregSheet when set by callers.
        Category: BILLING_CYCLE
        Local functions used: getCurrentBillingCycleDates(), getCurrentRateChartName(),
                              getRateColumnFromMetadata()
        Utility functions used: debugLog()

    buildBillingRowFromForm(formRow, prevRow, context, rowIndex) -> Array
        Constructs billing row array from a new student's form submission row.
        Sets Package Quantity from parsed lesson count (qty30/45/60 package text).
        Respects context.carryOverForms to determine whether form data should be carried
        across a school year boundary.
        Returns array representing complete billing row.
        Category: BILLING_CYCLE
        Local functions used: copyStaticFieldsToBillingRow(), setDeliveryPreference(),
                              populateInvoiceMetadata(), populateLetterType(), getExpandedPrograms(),
                              buildDynamicProgramColumns(), setPastColumnFormulas(),
                              addInvoiceTotalFormula(), applyCumulativeHistory(), applyLateFeeToRow()
        Utility functions used: parseAllPackageQuantities(), getLessonLengthFromPackages(),
                                normalizeHeader(), debugLog()

    buildBillingRowFromPrevious(prevRow, context, rowIndex) -> Array
        Constructs billing row array from a continuing student's previous billing cycle row.
        Carries over both Lesson Length and Package Quantity from the previous row.
        Respects context.carryOverForms for school year boundary handling.
        Returns array representing complete billing row.
        Category: BILLING_CYCLE
        Local functions used: copyStaticFieldsToBillingRow(), setDeliveryPreference(),
                              populateInvoiceMetadata(), populateLetterType(), buildDynamicProgramColumns(),
                              setPastColumnFormulas(), addInvoiceTotalFormula(), applyCumulativeHistory(),
                              applyLateFeeToRow(), setProgramQuantitiesForCarryover()
        Utility functions used: getStudentIdFromRow(), copyPreviousColumnToNew(), normalizeHeader(), debugLog()

    buildCarryoverStudents(previousData, existingStudentIds, context, allStudents, startingRowIndex) -> Number
        Builds billing rows for all carryover students from the previous billing cycle.
        Returns the next available row index after all carryover rows are added.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromPrevious()
        Utility functions used: debugLog()

    buildDocumentFileName(studentData, billingData, documentType, deliveryMethod) -> String
        Generates standardized document filename for generated PDFs.
        Returns formatted filename string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildDocumentSentence(billingData, deliveryMethod) -> String
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

    buildDynamicProgramColumns(newRow, row, enrolledPrograms, context, quantityCols, currencyCols, rowNum, lessonQuantity) -> void
        Populates dynamic program quantity and rate columns in the new billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    buildFormStudents(formSheet, filterDate, previousData, context, allStudents, startingRowIndex) -> Number
        Builds billing rows from new form submissions not present in carryover data.
        Returns the next available row index after all form student rows are added.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm()
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    buildFutureSemesterContext(futureSemesterName, customToday, billingCycleName, programConfig, prevHeaderMap) -> Object or null
        Builds a billing context for the next semester's form sheet so that early registrants
        for the upcoming semester can be appended to the current billing cycle.
        Returns null (silently) if the future semester's form sheet does not yet exist.
        Category: BILLING_CYCLE
        Local functions used: buildBillingContext()
        Utility functions used: debugLog()

    buildInvoiceTotalFormula(headerMap, rowNum) -> String
        Generates spreadsheet formula to calculate invoice total from program columns.
        Returns formula string.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    buildInvoiceVariableMap(studentData, billingData, isRefund) -> Object
        Creates variable map for invoice template merging.
        Returns object with all invoice template variables.
        Category: GENERATE_DOCUMENTS
        Local functions used: buildDynamicLineItems(), buildDynamicAmounts(), buildProgramDescription()
        Utility functions used: None

    buildMissingDocumentSentence(billingData) -> String
        Builds formatted sentence listing missing required documents.
        Returns formatted document list string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildPastVlookupFormula(prevSheetName, prevColIndex, rowNum) -> String
        Generates a VLOOKUP formula string that pulls a value from the previous billing cycle sheet
        by matching on Student ID in column C.
        Returns formula string.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None
        Called by: setPastColumnFormulas()

    buildProgramDescription(programTotals, lessonLength) -> String
        Builds formatted description of student's program registrations.
        Returns formatted program description string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    buildReregistrationQueue() -> void
        Refreshes the Reregistration Queue as a rolling work list. Adds newly qualifying students,
        updates lesson counts on existing rows, and removes students who have submitted an unprocessed
        reregistration response. Preserves Send?/Sent state for existing rows. Does not prompt to
        clear existing data. Threshold is 3 lessons for all packages; 1 lesson for packages with
        Package Quantity <= 4.
        Category: RE_REGISTRATION
        Local functions used: getCurrentBillingSheet()
        Utility functions used: getSheet(), getHeaderMap(), normalizeHeader(), debugLog()

    buildTemplateVariables(studentData, billingData, templateType) -> Object
        Builds complete variable map for all document template types.
        Returns object with all template merge variables.
        Category: GENERATE_DOCUMENTS
        Local functions used: buildInvoiceVariableMap(), buildDocumentSentence(),
                              buildProgramDescription()
        Utility functions used: debugLog()

    calculateTotalCreditsApplied(billingData) -> Number
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

    checkIfMediaReleaseNeeded(studentId) -> Boolean
        Determines if media release document is needed for student.
        Returns true if media release needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    clearDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap) -> void
        Clears document ID from billing sheet for regeneration.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName()
        Utility functions used: getHeaderMap(), debugLog()

    collectBillingDataDetailed(billingWB, studentIds) -> Object
        Collects detailed per-month billing data for a specific list of student IDs.
        Returns object mapping studentId -> month -> amount.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader()

    collectPaymentsDataDetailed(paymentsWB, studentIds) -> Object
        Collects detailed per-month payment data for a specific list of student IDs.
        Returns object mapping studentId -> month -> amount.
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

    copyStaticFieldsToBillingRow(newRow, sourceRow, context, getFn) -> void
        Copies unchanging fields from previous or form billing row to new row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    createBillingSheet(billingMonth, programConfig) -> Sheet
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
        Creates a Payment Verification summary sheet comparing all billed vs. paid amounts.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    createPaymentsTab(semesterName) -> Sheet
        Creates Payments tracking tab in billing workbook for a given semester.
        Returns newly created payments sheet.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    createRosterFolder(semesterName) -> Folder
        Creates Google Drive folder for teacher rosters for the given semester.
        Returns newly created folder.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    createSingleRegistrationPacketWithSelection(studentData, billingData, packetType, currentSemester, destinationFolder, forceRegenerate, selectedDocTypes) -> void
        Generates registration packet for single student with document selection dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: showSimpleDocumentSelectionDialog()
        Utility functions used: debugLog()

    determineIfNewStudent(studentId, currentSemester) -> Boolean
        Determines if student is new to the current semester.
        Returns true if new student.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: debugLog()

    determinePacketVersions(deliveryPreference) -> Object
        Determines which document versions to use based on delivery preference.
        Returns object with document version selections.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    doGet(e) -> HtmlOutput
        Web app entry point. Serves the ReRegistration HTML page with parent ID from URL parameter.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: None

    executeDocumentGeneration() -> void
        Executes document generation with user-selected documents from the dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: generateRegistrationPacketForStudentWithSelection()
        Utility functions used: debugLog()

    expandSheetAttendanceRows(sheet) -> void
        Expands attendance sheet rows to accommodate all student lesson slots.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    expandTeacherAttendanceRows(teacherSS) -> void
        Expands attendance rows in all monthly sheets of a teacher workbook.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    expandTeacherAttendanceSheets() -> void
        Expands attendance sheets for all teachers in the roster folder.
        Category: HELPER_FUNCTIONS
        Local functions used: expandTeacherAttendanceRows()
        Utility functions used: debugLog()

    extractBillingDataFromRow(billingRowData, headerMap) -> Object
        Extracts billing-specific data from billing sheet row.
        Returns object with billing amounts and settings.
        Category: HELPER_FUNCTIONS
        Local functions used: extractProgramTotals()
        Utility functions used: debugLog()

    extractDeliveryPreference(billingRowData, headerMap) -> String
        Extracts delivery preference (email/print/both) from a billing row.
        Returns delivery preference string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap()

    extractDocumentNames(documents) -> Array
        Extracts document names from document array.
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

    extractProgramTotals(row, headerMap) -> Object
        Extracts program totals from dynamic program columns in billing row.
        Returns object mapping program names to amounts.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    extractRosterDataForAttendance(rosterSheet) -> Array
        Extracts active student data from a teacher roster sheet for attendance generation.
        Returns array of student data objects.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: getHeaderMap(), debugLog()

    extractStudentDataFromBillingRow(row, headerMap) -> Object
        Extracts student-specific data from billing sheet row.
        Returns object with student demographic and program data.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    finalizeBillingSheet(billingSheet, context, formStudentIdsMap) -> void
        Shared post-population step for both populateBillingSheet and
        populateBillingSheetContinuingSemester. Applies re-registration overwrites,
        appends re-registration new students, sets context.processedReregIds,
        calls populateAllCumulativeColumns(), and applies conditional formatting.
        Category: BILLING_CYCLE
        Local functions used: applyReregistrationOverwrites(), appendReregistrationNewStudents(),
                              populateAllCumulativeColumns(), applyBillingConditionalFormatting()
        Utility functions used: None

    findBillingRowByStudentId(billingData, studentId, studentIdColIndex) -> Number or null
        Finds the row index of a student in billing sheet data.
        Returns row index (1-based), or null if not found.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    findMostRecentRosterSheet(workbook) -> Sheet or null
        Finds the most recent roster sheet in a teacher workbook based on Semester Metadata ordering.
        Returns null if no roster sheets exist.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getSheet(), normalizeHeader(), extractSeasonFromSemester(), debugLog()

    formatRow(sheet, rowIndex, quantityCols, currencyCols) -> void
        Applies number formatting to a specific billing row, including currency and quantity columns.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    generateCalendarForSemester(semesterName, startDate, endDate) -> void
        Generates semester calendar with holidays and important dates.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    generateDocumentForStudent(studentData, billingData, templateType, deliveryMethod, currentSemester, billingSheetName, billingSheet, headerMap) -> String
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

    generateInvoicesForBillingCycle(billingSheetName, options?) -> void
        Generates invoices for all students in billing cycle.
        Category: GENERATE_DOCUMENTS
        Local functions used: shouldGenerateInvoice(), buildInvoiceVariableMap()
        Utility functions used: debugLog()

    generateProgramFormulas(newRow, context, rowNum, quantityCols, currencyCols) -> void
        Generates formulas in program columns for automatic calculation.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    generateReconciliationSummary(results) -> String
        Builds and returns a formatted multi-line summary string from a reconciliation results
        object. Used by runPaymentsReconciliationUI() and runFullReconciliationUI() to display
        the alert. results object: {attendance, combined: {paymentsProcessed, paymentsErrors,
        agreementUpdates, mediaUpdates, formsErrors}, errors[]}.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    generateRefundInvoicesForBillingCycle(billingSheetName?) -> void
        Generates refund invoices for students with negative balances.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: debugLog()

    generateRegistrationPacketForStudentWithSelection(studentData, billingData, selectedDocumentTypes, deliveryMethods, currentSemester) -> void
        Generates registration packet with user-selected documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: determineIfNewStudent(), shouldIncludeDocument(),
                              generateDocumentForStudent()
        Utility functions used: debugLog()

    generateRegistrationPacketsForBillingCycle() -> void
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

    getBillingSheet(paymentDate, activeSheetName?, shouldLog?) -> Sheet
        Gets billing sheet by date match or active sheet name, with optional logging.
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

    getCurrentBillingCycleDates() -> Object
        Returns billing cycle start and end dates based on the current billing metadata.
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

    getCurrentRateChartName() -> String
        Gets the rate chart name for current semester from Billing Metadata.
        Returns rate chart name string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getCurrentSemesterFromBillingMetadata(billingSheetName) -> String
        Gets current semester name from Billing Metadata sheet by matching to billing sheet name.
        Returns semester name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getCurrentSemesterInfo() -> Object
        Gets comprehensive current semester information.
        Returns object with semester name, dates, and configuration.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

    getCurrentSemesterRateForLength(lessonLength) -> Number
        Gets current semester rate for specified lesson length.
        Returns rate amount.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getDocIdColumnName(docType) -> String
        Gets column name for storing document ID based on document type.
        Returns column name string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    getDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap) -> String
        Retrieves document ID from billing sheet for specific student and document type.
        Returns document ID or null.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName()
        Utility functions used: getHeaderMap()

    getDocumentSelectionHtml() -> String
        Generates HTML for document selection checkbox dialog.
        Returns HTML string.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    getExpandedPrograms(row, context) -> Array
        Gets expanded list of enrolled programs for a student based on form row and program config.
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

    getFutureSemesterName(currentSemesterName) -> String or null
        Looks up the next semester in Semester Metadata and returns its name.
        Returns null if currentSemesterName is the last row or sheet is not found.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getInvoiceNumber(billingSheet, billingRowIndex) -> String
        Gets the invoice number for a billing row, checking Invoice Number first then Past Invoice Number.
        Returns invoice value or null.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    getLessonLengthFromRow(row, get) -> String
        Extracts lesson length value from billing row using a column-getter function.
        Returns lesson length string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getNextMonthName(targetDate) -> String
        Gets the name of the month following the given date.
        Returns month name string.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    getPreviousSemester(currentSemesterName) -> String
        Gets previous semester name from current semester by looking up Semester Metadata.
        Returns previous semester name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getRateColumnFromMetadata(semesterName) -> Number
        Gets column index for the rate chart matching the given semester name in the Rates sheet.
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

    getStudentBalancesFromBilling(billingSheet, teacherId) -> Object
        Gets all student balances from billing sheet, optionally filtered by teacher ID.
        Returns object mapping student IDs to balance data.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getHeaderMap()

    getStudentDocumentsFolder(monthName) -> Folder
        Gets or creates the monthly subfolder under Student Documents in the generated documents folder.
        Returns the monthly folder.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getGeneratedDocumentsFolder(), debugLog()

    getStudentRegisteredLessonLength(studentId) -> String
        Gets student's registered lesson length from current billing sheet.
        Returns lesson length string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getHeaderMap()

    getStudentsNeedingPackets() -> Array
        Gets list of students in the current billing cycle who still need registration packets.
        Returns array of student data objects.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: debugLog()

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

    identifyWarningStudents(studentBalances) -> Array
        Identifies students with low hours or negative balances needing warnings.
        Returns array of warning student data.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: None

    isHeaderRow(row, dateIdx, lengthIdx) -> Boolean
        Checks if a row is a student header row by inspecting date and length column values.
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

    locateStudentRecord(rowData, paymentSheet, billingInfo) -> Object or null
        Locates a student's row in the billing sheet by matching invoice number or student ID
        from the given payment row data.
        Returns object with billingRowIndex and studentId, or null if not found.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    locateStudentRecordEnhanced(rowData, billingInfo, studentIdInput, invoiceNumberInput, lastNameCol, firstNameCol) -> Object or null
        Enhanced student record locator using multiple lookup strategies: invoice number,
        past invoice number, student ID, and last/first name fallback.
        Returns object with billingRowIndex, studentId, and match strategy used, or null if not found.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    logMysteryStudents(mysteryStudents) -> void
        Logs students found in teacher attendance but not in billing sheet.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    markReregistrationProcessed(reregSheet, processedIds) -> void
        Marks re-registration sheet rows as processed for the given student IDs.
        Category: RE_REGISTRATION
        Local functions used: None
        Utility functions used: normalizeHeader(), debugLog()

    markStudentsInactive() -> void
        Marks students as inactive in the current billing sheet at semester end.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    migrateTeacherDisplayNamesToIds() -> void
        One-time migration utility. Scans all billing cycle sheets and replaces teacher display name
        values with Teacher IDs from the Teachers and Admin sheet. Also renames the "Teacher" column
        header to "Teacher ID" if needed.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getSheet(), getHeaderMap(), normalizeHeader(), getMonthNames(), debugLog()

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
        Local functions used: buildBillingRowFromForm(), getFormsDataFromContacts(),
                              finalizeBillingSheet()
        Utility functions used: debugLog()

    populateBillingSheetContinuingSemester(context, billingSheet, existingStudentIds, previousData, previousStartDate?) -> void
        Populates billing sheet for continuing semester billing cycle.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromPrevious(), finalizeBillingSheet()
        Utility functions used: debugLog()

    populateCurrentBalanceFormula(newRow, context, rowIndex) -> void
        Populates formula calculating current balance in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    populateDeliveryPreference(newRow, formRow, context) -> void
        Populates delivery preference in billing row from form row data.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateDeliveryPreferenceFromPrevious(newRow, prevRow, context) -> void
        Copies delivery preference from previous billing cycle row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateInvoiceMetadata(newRow, studentId, context, rowIndex) -> void
        Populates invoice metadata fields (number, URL, etc.) in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateLateFee(newRow, context, currencyCols) -> void
        Calculates and populates late fee in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    populateLetterType(newRow, context, sourceType, prevRow) -> void
        Populates letter type (New/Returning) in billing row based on source type and previous data.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap()

    processDocumentSelection(selectedTypes) -> void
        Processes document type selection from the checkbox dialog and triggers generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: executeDocumentGeneration()
        Utility functions used: None

    processFieldMapForSemester(semesterName) -> void
        Processes and stores field map for semester form responses.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    processFormsData(paymentSheet, studentsSheet) -> Array
        Processes form submissions from the payment sheet and reconciles against the students sheet.
        Returns array of form reconciliation results.
        Category: RECONCILIATION
        Local functions used: processFormsReconciliationForRow()
        Utility functions used: debugLog()

    processFormsReconciliationForRow(rowData, studentsSheet, studentIdInput, hasAgreement, hasMediaResponse, studentIdCol, agreementCol, mediaCol) -> Object
        Processes a single form submission row against the students sheet, updating agreement
        and media consent status in both the students sheet and current billing sheet.
        Returns reconciliation result object.
        Category: RECONCILIATION
        Local functions used: getCurrentBillingSheet()
        Utility functions used: debugLog()

    processPaymentReconciliationForRow(rowData, paymentSheet, rowNumber, studentIdInput, invoiceNumberInput, amountCol, dateCol, lastNameCol, firstNameCol) -> Object
        Processes a single payment row against billing data using multiple lookup strategies.
        Writes total payments to billing sheet and returns result with studentId and invoiceNumber.
        Returns payment reconciliation result object.
        Category: RECONCILIATION
        Local functions used: locateStudentRecordEnhanced(), sumPayments(), reconcilePayment()
        Utility functions used: getHeaderMap(), debugLog()

    processTeacherAttendanceForBilling(teacherSS, targetDate) -> Object
        Processes teacher's attendance data from monthly sheets and aggregates hours by student and month.
        Returns object with studentHoursByMonth and aggregated data.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getMonthNames(), debugLog()

    processTeacherForNewAttendance(teacherId, rosterUrl, targetMonthName) -> void
        Processes a single teacher's roster to create a new monthly attendance sheet.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: debugLog()

    processTeacherReconciliation(teacherId, rosterUrl, currentBillingSheet, targetDate, currentSemester) -> Object
        Reconciles teacher's roster and attendance with billing data for a given period.
        Returns reconciliation summary for teacher.
        Category: RECONCILIATION
        Local functions used: processTeacherAttendanceForBilling(), updateBillingForStudents()
        Utility functions used: debugLog()

    promptForBillingCycleName(customToday) -> String
        Prompts user to enter or confirm billing cycle name based on the custom date.
        Returns billing cycle name.
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
        Protects billing sheet columns from accidental edits using warning-only protection.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    protectPreviousBillingCycle() -> void
        Protects previous billing cycle sheet after new cycle is created.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: debugLog()

    reconcilePayment(paymentSheet, billingSheet, billingRowIndex, paymentRowNumber, invoiceValue, totalPayments) -> void
        Writes the total payment amount to the Payment Received column of the matching billing row
        and marks the payment row as reconciled.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    renameLatestFormSheet(semesterName) -> void
        Renames most recent form responses sheet to include semester name.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: debugLog()

    runBillingCycleAutomation() -> void
        Main entry point for the full billing cycle automation process. Prompts for cycle name,
        verifies programs and rates, creates the billing sheet, detects school year boundaries
        via Rates Verification in Semester Metadata (carryOverForms), populates from previous
        or form data, appends future-semester students if a next-semester form sheet exists,
        applies discounts, marks re-registrations, and appends billing metadata.
        Category: UI_MENU
        Local functions used: promptForBillingCycleName(), verifyProgramsForSemester(),
                              verifyRatesEnhanced(), createBillingSheet(), extractPreviousBillingData(),
                              updateCumulativeTracking(), loadReregistrationData(), buildBillingContext(),
                              populateBillingSheet(), populateBillingSheetContinuingSemester(),
                              getFutureSemesterName(), buildFutureSemesterContext(),
                              appendStudentsFromContext(), applyMultiStudentDiscount(),
                              populateAllCumulativeColumns(), applyBillingConditionalFormatting(),
                              markReregistrationProcessed(), appendToBillingMetadata()
        Utility functions used: getCurrentSemesterName(), getHeaderMap(), normalizeHeader(), debugLog()

    runFullReconciliation(reconciliationDate) -> void
        Runs complete reconciliation process for a given date. Calls lesson reconciliation first,
        then payment/forms reconciliation. Returns results object.
        Category: UI_MENU
        Local functions used: runWeeklyLessonReconciliation(), runPaymentsReconciliation(),
                              generateReconciliationSummary()
        Utility functions used: debugLog()

    runFullReconciliationUI() -> void
        UI wrapper for full reconciliation. Prompts for date, confirms, runs reconciliation,
        and shows summary alert.
        Category: UI_MENU
        Local functions used: runFullReconciliation()
        Utility functions used: promptForDate(), debugLog()

    runLogHeaders() -> void
        Debug utility. Calls UtilityScriptLibrary.logAllSheetHeaders() to dump all sheet
        headers to the debug log.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: logAllSheetHeaders()

    runPaymentsReconciliation() -> Object
        Consolidated payments reconciliation. For each unreconciled row in the current semester's
        Payments sheet, processes payment data via processPaymentReconciliationForRow() and form data
        (agreement / media release) via processFormsReconciliationForRow(). Writes Reconciled
        checkbox, Student ID, Invoice Number, and System Comments back to the Payments sheet.
        Returns results object: {paymentsProcessed, paymentsErrors, agreementUpdates, mediaUpdates,
        formsErrors, details[]}.
        Category: RECONCILIATION
        Local functions used: processPaymentReconciliationForRow(), processFormsReconciliationForRow()
        Utility functions used: getWorkbook(), getCurrentSemesterName(), getHeaderMap(),
                                normalizeHeader(), debugLog()

    runPaymentsReconciliationUI() -> void
        UI wrapper for runPaymentsReconciliation(). Displays the formatted summary alert on success.
        Category: UI_MENU
        Local functions used: runPaymentsReconciliation(), generateReconciliationSummary()
        Utility functions used: debugLog()

    runRegistrationPacketGenerationUI() -> void
        UI wrapper for registration packet generation with error handling.
        Category: UI_MENU
        Local functions used: generateRegistrationPacketsForBillingCycle()
        Utility functions used: debugLog()

    runWeeklyLessonReconciliation(customDate?) -> void
        Reconciles teacher attendance records with billing data for all active teachers.
        Category: RECONCILIATION
        Local functions used: processTeacherReconciliation(), applyBillingConditionalFormatting(),
                              updateLastReconciliationDate()
        Utility functions used: getCurrentSemesterName(), getSheet(), debugLog()

    runWeeklyLessonReconciliationUI() -> void
        UI wrapper for lesson reconciliation. Prompts for date, confirms, runs reconciliation,
        and displays success alert.
        Category: UI_MENU
        Local functions used: runWeeklyLessonReconciliation()
        Utility functions used: promptForDate(), formatDateFlexible(), debugLog()

    selectDocumentTemplate(templateType, studentData, deliveryMethod, currentSemester, billingData) -> Object
        Selects appropriate document template based on parameters and student data.
        Returns template file object.
        Category: GENERATE_DOCUMENTS
        Local functions used: determinePacketVersions()
        Utility functions used: None

    sendReregistrationLinks() -> void
        Sends re-registration emails to parents whose rows are checked (Send? = true) and not yet
        sent in the Reregistration Queue. Confirms count before sending. Writes sent date as
        MM/dd/yy to the Sent column on success.
        Category: RE_REGISTRATION
        Local functions used: getCurrentBillingSheet()
        Utility functions used: debugLog(), getHeaderMap(), getSheet(), normalizeHeader()

    setDeliveryPreference(newRow, sourceRow, context, sourceIsForm) -> void
        Sets the Delivery Preference field in a new billing row from either a form row or previous billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: normalizeHeader(), getHeaderMap()

    setPastColumnFormulas(newRow, context, rowNum, currencyCols) -> void
        Writes VLOOKUP formulas into all "Past" columns (Past Balance, Past Invoice Number,
        Past Hours Remaining, Past Cumulative Hours) by referencing the previous billing cycle sheet.
        Marks currency and hours columns appropriately.
        Category: BILLING_CYCLE
        Local functions used: buildPastVlookupFormula()
        Utility functions used: normalizeHeader(), addToCurrencyCols()
        Called by: buildBillingRowFromForm(), buildBillingRowFromPrevious()

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

    setupRosterTemplateProtection(sheet) -> void
        Sets up protection, date validation, and status dropdown on a roster template sheet.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: protectSheetRanges()

    shouldGenerateInvoice(billingData) -> Boolean
        Determines if invoice should be generated for student based on billing data.
        Returns true if invoice needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    shouldIncludeAgreement(studentId, billingSheet, headerMap) -> Boolean
        Determines if agreement document should be included in packet.
        Returns true if agreement needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    shouldIncludeDocument(studentId, docType, billingSheet, headerMap) -> Boolean
        Determines if specific document type should be included in packet.
        Returns true if document needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: shouldIncludeAgreement(), shouldIncludeMediaRelease()
        Utility functions used: None

    shouldIncludeMediaRelease(studentId, billingSheet, headerMap) -> Boolean
        Determines if media release should be included in packet.
        Returns true if media release needed.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    shouldUseMissingDocumentLetter(billingData) -> Boolean
        Determines if missing document letter should be used instead of welcome letter.
        Returns true if missing documents exist.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    showSimpleDocumentSelectionDialog() -> void
        Displays HTML dialog for manual document type selection.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocumentSelectionHtml()
        Utility functions used: None

    submitReregistration(data) -> Object
        Processes re-registration form submission. Writes student package selection to the
        Reregistration sheet and updates parent contact info in Contacts if changes were submitted.
        Returns {success, updatedStudents[], errors[], message}
        Category: RE_REGISTRATION
        Local functions used: getCurrentBillingSheet()
        Utility functions used: debugLog(), findParentRow(), generateKey(), getHeaderMap(),
                                getSheet(), normalizeHeader(), updateParentContactFields()

    sumPayments(sheet, studentId, startDate, endDate) -> Number
        Sums all payments for a student within the given date range from the provided sheet.
        Returns total payment amount.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    updateBillingForStudents(billingSheet, studentHours) -> Number
        Updates the Current Hours Taught This Billing Cycle column in the billing sheet
        for each student ID found in the studentHours map.
        Returns count of updated students.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()
        Called by: processTeacherReconciliation(), runWeeklyLessonReconciliation()

    updateCumulativeTracking(billingSheetName) -> void
        Updates the Cumulative Tracking sheet with totals from the specified billing cycle sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    updateDocIdInBillingSheet(studentId, docType, fileId, docUrl, billingSheet, headerMap) -> void
        Updates document ID and URL in billing sheet as a HYPERLINK formula after document generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName()
        Utility functions used: getHeaderMap(), debugLog()

    updateInvoiceUrlInBillingSheet(billingSheet, rowNumber, headerMap, url) -> void
        Updates invoice URL in the Invoice URL column of the billing sheet for the given row.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: normalizeHeader()

    updateLastReconciliationDate(billingSheet, studentDataCurrentThrough) -> void
        Updates the Last Reconciliation Date column for each student in studentDataCurrentThrough
        map. Only overwrites if the new date is more recent than the existing value.
        studentDataCurrentThrough maps studentId -> Date.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    updateSheetStudentWarnings(sheet, warningStudentMap, formattedDate) -> void
        Updates attendance sheet with warning highlights for students with low hours or balances.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: debugLog()

    updateTeacherRosterBalances(teacherSS, studentBalances, currentSemester) -> void
        Updates teacher roster with current balance information from billing sheet.
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

    verifyProgramsForSemester(semesterName, billingSS) -> Object
        Verifies all required programs are set up for semester. Checks that all Package programs
        have valid Alias For references. Prompts user to confirm active programs list.
        Returns {status, activePrograms[], issues[]}.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: getHeaderMap(), normalizeHeader(), debugLog()

    verifyRatesEnhanced() -> Boolean
        Verifies all rates are properly configured for billing. Displays current rate chart
        values and prompts user to confirm.
        Returns true if confirmed.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: getMostRecentRateColumn(), getHeaderMap(), normalizeHeader(),
                                formatCurrency(), debugLog()

    writeAndFormatRows(billingSheet, allStudents) -> void
        Writes all student rows to the billing sheet in a single batch operation and applies formatting.
        Category: BILLING_CYCLE
        Local functions used: formatRow()
        Utility functions used: None
================================================================================
END OF FUNCTION DIRECTORY
================================================================================