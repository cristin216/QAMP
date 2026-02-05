================================================================================
BILLING FUNCTION DIRECTORY
================================================================================
    Total Functions: 161
    Most Recent version: 163

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
    addOveragesToBillingSheet, appendToBillingMetadata, appendToSemesterMetadata, 
    applyAdminVisualFormatting, applyLessonEquivalentCredits, applyLetterTypeValidation, 
    applyWarningsToTeacherWorkbook, buildBillingContext, buildBillingRowFromForm, 
    buildBillingRowFromPrevious, buildDocumentFileName, buildDocumentSentence, 
    buildDynamicAmounts, buildDynamicLineItems, buildDynamicProgramColumns, 
    buildDynamicProgramColumnsWithCredits, buildInvoiceTotalFormula, buildInvoiceVariableMap, 
    buildMissingDocumentSentence, buildProgramDescription, buildTemplateVariables, 
    calculateLessonEquivalents, calculateTotalCreditsApplied, cancelDocumentGeneration, 
    checkIfMediaReleaseNeeded, checkRowFormatting, clearDocIdFromBillingSheet, 
    continuePacketGeneration, copyStaticFieldsToBillingRow, createBillingSheet, 
    createNewAttendanceSheets, createPaymentsTab, createRosterFolder, 
    createSingleRegistrationPacketWithSelection, debugBillingSheetColumns, 
    detectAndBillOverages, determineIfNewStudent, determinePacketVersions, 
    executeDocumentGeneration, expandSheetAttendanceRows, expandTeacherAttendanceRows, 
    expandTeacherAttendanceSheets, extractBillingDataFromRow, extractDeliveryPreference, 
    extractDocumentNames, extractPreviousBillingData, extractProgramTotals, 
    extractRosterDataForAttendance, extractStudentDataFromBillingRow, formatRow, 
    generateCalendarForSemester, generateDocumentForStudent, generateInvoicesForBillingCycle, 
    generateProgramFormulas, generateReconciliationSummary, generateReconciliationSummaryUpdated, 
    generateRefundInvoicesForBillingCycle, generateRegistrationPacketForStudentWithSelection, 
    generateRegistrationPacketsForBillingCycle, getActivePrograms, getBillingSheet, 
    getCurrentBillingCycleDates, getCurrentBillingSheet, getCurrentRateChartName, 
    getCurrentSemesterFromBillingMetadata, getCurrentSemesterInfo, getCurrentSemesterName, 
    getCurrentSemesterRateForLength, getDocIdColumnName, getDocIdFromBillingSheet, 
    getDocumentSelectionHtml, getExpandedPrograms, getFormsDataFromContacts, 
    getInvoiceNumber, getLessonLengthFromRow, getNextMonthName, getPreviousSemester, 
    getPreviousSemesterBalance, getRateChartForSemester, getRateColumnFromMetadata, 
    getRateForSemester, getSemesterForDate, getStudentBalancesFromBilling, 
    getStudentDocumentsFolder, getStudentRegisteredLessonLength, getStudentsNeedingPackets, 
    grantDocumentPermissions, identifyWarningStudents, isHeaderRow, locateStudentRecord, 
    locateStudentRecordEnhanced, logMysteryStudents, markStudentsInactive, onOpen, 
    populateAllCumulativeColumns, populateBillingSheet, populateBillingSheetContinuingSemester, 
    populateCurrentBalanceFormula, populateDeliveryPreference, populateDeliveryPreferenceFromPrevious, 
    populateInvoiceMetadata, populateLateFee, populateLetterType, populatePastBalanceAndCredit, 
    processDocumentSelection, processFieldMapForSemester, processFormsData, 
    processFormsReconciliationForRow, processPaymentReconciliationForRow, processPaymentRecord, 
    processSemesterEndCredits, processTeacherAttendanceForBilling, processTeacherForNewAttendance, 
    processTeacherReconciliation, promptForBillingCycleName, promptForCustomToday, 
    promptForSemesterDates, promptForSemesterName, protectBillingSheet, 
    protectPreviousBillingCycle, reconcilePayment, renameLatestFormSheet, 
    runBillingCycleAutomation, runCombinedReconciliation, runFormsReconciliation, 
    runFormsReconciliationUI, runFullReconciliation, runFullReconciliationUI, 
    runPaymentReconciliation, runPaymentReconciliationUI, runRegistrationPacketGenerationUI, 
    runWeeklyLessonReconciliation, runWeeklyLessonReconciliationUI, selectDocumentTemplate, 
    setupNewSemester, setupRosterTemplateProtection, shouldGenerateInvoice, 
    shouldIncludeAgreement, shouldIncludeDocument, shouldIncludeMediaRelease, 
    shouldUseMissingDocumentLetter, showSimpleDocumentSelectionDialog, storeSemesterEndBalances, 
    sumPayments, testDynamicInvoiceFunctions, testDynamicInvoiceFunctionsOnCorrectSheet, 
    testExtractStudentDataFromBillingRow, testFormatDateFunction, testFullAgreementGeneration, 
    testPrompt, testTemplateLiteral, updateBillingForTeacherStudents, 
    updateDocIdInBillingSheet, updateInvoiceUrlInBillingSheet, updateSheetStudentWarnings, 
    updateTeacherRosterBalances, verifyProgramsForSemester, verifyRatesEnhanced

  ================================================================================
  FUNCTION CATEGORIES:
  ================================================================================

    UI_MENU (13 functions):
      getNextMonthName, onOpen, runBillingCycleAutomation, runFormsReconciliation, 
      runFormsReconciliationUI, runFullReconciliation, runFullReconciliationUI, 
      runPaymentReconciliation, runPaymentReconciliationUI, runRegistrationPacketGenerationUI, 
      runWeeklyLessonReconciliationUI, setupNewSemester, setupRosterTemplateProtection

    SETUP_SEMESTER (8 functions):
      appendToSemesterMetadata, createPaymentsTab, createRosterFolder, generateCalendarForSemester, 
      markStudentsInactive, processFieldMapForSemester, renameLatestFormSheet, 
      verifyProgramsForSemester

    BILLING_CYCLE (29 functions):
      addOveragesToBillingSheet, appendToBillingMetadata, applyLessonEquivalentCredits, 
      applyLetterTypeValidation, buildBillingContext, buildBillingRowFromForm, 
      buildBillingRowFromPrevious, buildDynamicAmounts, buildDynamicLineItems, 
      buildDynamicProgramColumns, buildDynamicProgramColumnsWithCredits, buildInvoiceTotalFormula, 
      copyStaticFieldsToBillingRow, createBillingSheet, detectAndBillOverages, formatRow, 
      generateProgramFormulas, populateAllCumulativeColumns, populateBillingSheet, 
      populateBillingSheetContinuingSemester, populateCurrentBalanceFormula, 
      populateDeliveryPreference, populateDeliveryPreferenceFromPrevious, populateInvoiceMetadata, 
      populateLateFee, populateLetterType, populatePastBalanceAndCredit, processSemesterEndCredits, 
      protectPreviousBillingCycle

    GENERATE_DOCUMENTS (35 functions):
      buildDocumentFileName, buildDocumentSentence, buildInvoiceVariableMap, 
      buildMissingDocumentSentence, buildProgramDescription, buildTemplateVariables, 
      cancelDocumentGeneration, checkIfMediaReleaseNeeded, clearDocIdFromBillingSheet, 
      continuePacketGeneration, createSingleRegistrationPacketWithSelection, 
      determineIfNewStudent, determinePacketVersions, executeDocumentGeneration, 
      extractDocumentNames, extractRosterDataForAttendance, generateDocumentForStudent, 
      generateInvoicesForBillingCycle, generateRefundInvoicesForBillingCycle, 
      generateRegistrationPacketForStudentWithSelection, generateRegistrationPacketsForBillingCycle, 
      getDocIdColumnName, getDocIdFromBillingSheet, getDocumentSelectionHtml, 
      processDocumentSelection, processTeacherForNewAttendance, selectDocumentTemplate, 
      shouldGenerateInvoice, shouldIncludeAgreement, shouldIncludeDocument, 
      shouldIncludeMediaRelease, shouldUseMissingDocumentLetter, showSimpleDocumentSelectionDialog, 
      updateDocIdInBillingSheet, updateInvoiceUrlInBillingSheet

    RECONCILIATION (27 functions):
      applyAdminVisualFormatting, applyWarningsToTeacherWorkbook, expandSheetAttendanceRows, 
      findBillingRowByStudentId, findStudentInContacts, generateReconciliationSummary, 
      generateReconciliationSummaryUpdated, getBillingSheet, getInvoiceNumber, 
      getStudentBalancesFromBilling, identifyWarningStudents, locateStudentRecord, 
      locateStudentRecordEnhanced, logMysteryStudents, processFormsData, 
      processFormsReconciliationForRow, processPaymentReconciliationForRow, processPaymentRecord, 
      processTeacherAttendanceForBilling, processTeacherReconciliation, reconcilePayment, 
      runCombinedReconciliation, runWeeklyLessonReconciliation, sumPayments, 
      updateBillingForTeacherStudents, updateSheetStudentWarnings, updateTeacherRosterBalances

    HELPER_FUNCTIONS (39 functions):
      calculateLessonEquivalents, calculateTotalCreditsApplied, createNewAttendanceSheets, 
      expandSheetAttendanceRows, expandTeacherAttendanceRows, expandTeacherAttendanceSheets, 
      extractBillingDataFromRow, extractDeliveryPreference, extractPreviousBillingData, 
      extractProgramTotals, extractStudentDataFromBillingRow, getActivePrograms, 
      getCurrentBillingCycleDates, getCurrentBillingSheet, getCurrentRateChartName, 
      getCurrentSemesterFromBillingMetadata, getCurrentSemesterInfo, getCurrentSemesterName, 
      getCurrentSemesterRateForLength, getExpandedPrograms, getFormsDataFromContacts, 
      getLessonLengthFromRow, getPreviousSemester, getPreviousSemesterBalance, 
      getRateChartForSemester, getRateColumnFromMetadata, getRateForSemester, getSemesterForDate, 
      getStudentDocumentsFolder, getStudentRegisteredLessonLength, getStudentsNeedingPackets, 
      isHeaderRow, promptForBillingCycleName, promptForCustomToday, promptForSemesterDates, 
      promptForSemesterName, protectBillingSheet, storeSemesterEndBalances, verifyRatesEnhanced

    TESTING (10 functions):
      checkRowFormatting, debugBillingSheetColumns, grantDocumentPermissions, 
      testDynamicInvoiceFunctions, testDynamicInvoiceFunctionsOnCorrectSheet, 
      testExtractStudentDataFromBillingRow, testFormatDateFunction, testFullAgreementGeneration, 
      testPrompt, testTemplateLiteral

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addOveragesToBillingSheet(billingSheet, overageData) -> void
        Adds overage charges to billing sheet for students who exceeded lesson registrations.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    appendToBillingMetadata(billingCycleName, billingDate, semesterName) -> void
        Appends new billing cycle information to Billing Metadata sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    appendToSemesterMetadata(semesterName, startDate, endDate, fieldMap) -> void
        Appends new semester information to Semester Metadata sheet.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    applyAdminVisualFormatting(sheet, dataRowCount) -> void
        Applies visual formatting to reconciliation sheets for admin viewing.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    applyLessonEquivalentCredits(billingContext, studentData, row, billingSheet) -> void
        Applies lesson equivalent credits for students switching from packages to private lessons.
        Category: BILLING_CYCLE
        Local functions used: calculateLessonEquivalents()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    applyLetterTypeValidation(sheet) -> void
        Applies data validation dropdown to Letter Type column in billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    applyWarningsToTeacherWorkbook(teacherWorkbook, warningStudents) -> void
        Highlights warning students in teacher's attendance sheets with red rows.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildBillingContext(customToday, semesterName, billingCycleName) -> Object
        Builds comprehensive context object containing all billing cycle information.
        Returns object with dates, rates, semester info, and configuration.
        Category: BILLING_CYCLE
        Local functions used: getCurrentBillingCycleDates(), getCurrentRateChartName(), 
                              getRateColumnFromMetadata()
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildBillingRowFromForm(context, formData, studentId) -> Array
        Constructs billing row data from student form submission for new students.
        Returns array representing complete billing row.
        Category: BILLING_CYCLE
        Local functions used: buildDynamicProgramColumns(), populateDeliveryPreference()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    buildBillingRowFromPrevious(context, prevRow, prevHeaderMap, studentId) -> Array
        Constructs billing row data from previous billing cycle for continuing students.
        Returns array representing complete billing row.
        Category: BILLING_CYCLE
        Local functions used: buildDynamicProgramColumns(), populateDeliveryPreferenceFromPrevious()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    buildDynamicProgramColumnsWithCredits(context, studentData, previousPrograms, existingCredits?) -> Array
        Builds program columns with credit carryover for continuing semester students.
        Returns array of program values including credits.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    calculateLessonEquivalents(previousLessons, currentLessonLength) -> Number
        Calculates lesson equivalents when converting between lesson lengths.
        Returns number of equivalent lessons.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    continuePacketGeneration() -> void
        Continues registration packet generation after document selection.
        Category: GENERATE_DOCUMENTS
        Local functions used: processDocumentSelection()
        Utility functions used: None

    copyStaticFieldsToBillingRow(row, headerMap, prevRow, prevHeaderMap) -> void
        Copies unchanging fields from previous billing row to new row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createBillingSheet(billingCycleName) -> Sheet
        Creates new billing cycle sheet with headers and formatting.
        Returns newly created billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createNewAttendanceSheets() -> void
        Creates new month attendance sheets for all active teachers.
        Category: HELPER_FUNCTIONS
        Local functions used: extractRosterDataForAttendance(), processTeacherForNewAttendance()
        Utility functions used: UtilityScriptLibrary.debugLog()

    createPaymentsTab() -> Sheet
        Creates Payments tracking tab in billing workbook.
        Returns newly created payments sheet.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createRosterFolder(teacherName, semesterName) -> Folder
        Creates Google Drive folder for teacher roster workbook.
        Returns newly created folder.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    createSingleRegistrationPacketWithSelection(studentId, billingCycleName) -> void
        Generates registration packet for single student with document selection dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: showSimpleDocumentSelectionDialog()
        Utility functions used: UtilityScriptLibrary.debugLog()

    debugBillingSheetColumns(sheetName?) -> void
        Logs billing sheet column information for debugging.
        Category: TESTING
        Local functions used: getCurrentBillingSheet()
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    detectAndBillOverages(billingSheet, billingCycleName) -> Array
        Detects and bills students with lesson overages from previous cycle.
        Returns array of overage records.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    determineIfNewStudent(studentId, semesterName) -> Boolean
        Determines if student is new to the current semester.
        Returns true if new student.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    determinePacketVersions(newOrReturning, deliveryMode) -> Object
        Determines which document versions to use based on student status and delivery.
        Returns object with document version selections.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    executeDocumentGeneration(studentId, billingCycleName, selectedDocs) -> void
        Executes document generation with user-selected documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: generateRegistrationPacketForStudentWithSelection()
        Utility functions used: UtilityScriptLibrary.debugLog()

    expandSheetAttendanceRows(sheet, requiredRows) -> void
        Expands attendance sheet to have minimum required rows for all students.
        Category: RECONCILIATION or HELPER_FUNCTIONS (duplicate in both sections)
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    expandTeacherAttendanceRows(teacherWorkbook, monthName, studentCount) -> void
        Expands specific month attendance sheet in teacher workbook.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    expandTeacherAttendanceSheets(teacherWorkbook, studentCount) -> void
        Expands all attendance sheets in teacher workbook to accommodate students.
        Category: HELPER_FUNCTIONS
        Local functions used: expandTeacherAttendanceRows()
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractBillingDataFromRow(billingRow, headerMap) -> Object
        Extracts billing-specific data from billing sheet row.
        Returns object with billing amounts and settings.
        Category: HELPER_FUNCTIONS
        Local functions used: extractProgramTotals()
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractDeliveryPreference(contactsSheet, studentId) -> String
        Extracts delivery preference (email/print/both) from Contacts sheet.
        Returns delivery preference string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    extractDocumentNames(selectedCheckboxes) -> Array
        Extracts document names from checkbox selection array.
        Returns array of document name strings.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: None

    extractPreviousBillingData(options?) -> Object
        Extracts student billing data from previous billing cycle sheet.
        Returns object with previous billing data and mappings.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    extractStudentDataFromBillingRow(billingRow, headerMap) -> Object
        Extracts student-specific data from billing sheet row.
        Returns object with student demographic and program data.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    findMostRecentRosterSheet(workbook) -> Sheet or null
        Finds the most recent roster sheet in a teacher workbook based on Semester Metadata ordering.
        Searches for sheets ending in " Roster", extracts their season names, matches them against
        Semester Metadata, and returns the roster sheet corresponding to the most recent semester.
        Returns null if no roster sheets exist.
        Category: ATTENDANCE_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.normalizeHeader(),
                              UtilityScriptLibrary.extractSeasonFromSemester(), UtilityScriptLibrary.debugLog()

    formatRow(sheet, rowIndex) -> void
        Applies formatting to a specific row in billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    generateCalendarForSemester(semesterName, startDate, endDate) -> void
        Generates semester calendar with holidays and important dates.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateDocumentForStudent(studentId, billingCycleName, docType, newOrReturning, deliveryMode) -> String
        Generates single document for student and returns document ID.
        Returns Google Docs document ID.
        Category: GENERATE_DOCUMENTS
        Local functions used: buildTemplateVariables(), selectDocumentTemplate(), 
                              buildDocumentFileName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateInvoicesForBillingCycle(billingCycleName?) -> void
        Generates invoices for all students in billing cycle.
        Category: GENERATE_DOCUMENTS
        Local functions used: shouldGenerateInvoice(), buildInvoiceVariableMap()
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateProgramFormulas(row, context, billingSheet) -> void
        Generates formulas in program columns for automatic calculation.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateRegistrationPacketForStudentWithSelection(studentId, billingCycleName, selectedDocs) -> void
        Generates registration packet with user-selected documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: determineIfNewStudent(), shouldIncludeDocument(), 
                              generateDocumentForStudent()
        Utility functions used: UtilityScriptLibrary.debugLog()

    generateRegistrationPacketsForBillingCycle(billingCycleName?) -> void
        Generates registration packets for all students needing documents.
        Category: GENERATE_DOCUMENTS
        Local functions used: getStudentsNeedingPackets(), determineIfNewStudent()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getActivePrograms() -> Array
        Gets list of currently active programs from Programs List sheet.
        Returns array of active program names.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getBillingSheet(billingCycleName?) -> Sheet
        Gets billing sheet by name or most recent if no name provided.
        Returns billing sheet object.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getCurrentBillingCycleDates(customToday, semesterName) -> Object
        Calculates billing cycle start and end dates based on current date.
        Returns object with startDate and endDate.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getCurrentBillingSheet() -> Sheet
        Gets the most recent billing cycle sheet.
        Returns current billing sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    getCurrentSemesterInfo() -> Object
        Gets comprehensive current semester information.
        Returns object with semester name, dates, and configuration.
        Category: HELPER_FUNCTIONS
        Local functions used: getCurrentSemesterName()
        Utility functions used: UtilityScriptLibrary.debugLog()

    getCurrentSemesterName() -> String
        Gets current semester name from Semester Metadata sheet.
        Returns semester name string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    getFormsDataFromContacts() -> Object
        Retrieves form response data from Contacts sheet.
        Returns object with form data indexed by student ID.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRateChartForSemester(semesterName) -> String
        Gets rate chart name for specified semester.
        Returns rate chart name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRateColumnFromMetadata(rateChartName) -> Number
        Gets column index for rate chart in Rates sheet.
        Returns column index.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getRateForSemester(lessonLength, semesterInfo) -> Number
        Gets rate for lesson length in specified semester.
        Returns rate amount.
        Category: HELPER_FUNCTIONS
        Local functions used: getRateColumnFromMetadata()
        Utility functions used: None

    getSemesterForDate(date) -> String
        Determines semester name for given date.
        Returns semester name.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getStudentBalancesFromBilling(billingSheet?) -> Object
        Gets all student balances from billing sheet.
        Returns object mapping student IDs to balance data.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    getStudentDocumentsFolder(studentId, semesterName) -> Folder
        Gets or creates student's documents folder in Google Drive.
        Returns student documents folder.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    getStudentRegisteredLessonLength(studentId) -> String
        Gets student's registered lesson length from Contacts sheet.
        Returns lesson length string.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    getStudentsNeedingPackets(billingCycleName) -> Array
        Gets list of students needing registration packets generated.
        Returns array of student IDs.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    grantDocumentPermissions(docId, email) -> void
        Grants view permissions on generated document to student/parent email.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

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

    locateStudentRecord(teacherId, studentId) -> Object
        Locates student record in teacher roster or attendance sheets.
        Returns object with sheet and row information.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    locateStudentRecordEnhanced(teacherId, studentId, monthName?) -> Object
        Enhanced version of locateStudentRecord with month-specific search.
        Returns object with detailed location information.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    logMysteryStudents(teacherData) -> void
        Logs students found in teacher attendance but not in billing sheet.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    markStudentsInactive(semesterName) -> void
        Marks students as inactive when semester ends.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    onOpen() -> void
        Creates custom menu in spreadsheet UI when workbook opens.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    populateAllCumulativeColumns() -> void
        Populates all cumulative total columns in current billing sheet.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    populateBillingSheet(context, carryOverData) -> void
        Populates billing sheet for first cycle of new semester.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromForm(), getFormsDataFromContacts()
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateBillingSheetContinuingSemester(context, billingSheet, existingIds, carryOverData, previousDate?) -> void
        Populates billing sheet for continuing semester billing cycle.
        Category: BILLING_CYCLE
        Local functions used: buildBillingRowFromPrevious(), processSemesterEndCredits()
        Utility functions used: UtilityScriptLibrary.debugLog()

    populateCurrentBalanceFormula(row, headerMap) -> void
        Populates formula calculating current balance in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: None

    populateDeliveryPreference(rowIndex, billingSheet, studentId) -> void
        Populates delivery preference in billing row from Contacts sheet.
        Category: BILLING_CYCLE
        Local functions used: extractDeliveryPreference()
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    populateDeliveryPreferenceFromPrevious(rowIndex, billingSheet, prevRow, prevHeaderMap) -> void
        Copies delivery preference from previous billing cycle.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    populateInvoiceMetadata(rowIndex, billingSheet, billingCycleName) -> void
        Populates invoice metadata fields (number, URL, etc.) in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    populateLateFee(rowIndex, billingSheet) -> void
        Calculates and populates late fee based on past balance.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    populateLetterType(rowIndex, billingSheet, newOrReturning) -> void
        Populates letter type (New/Returning) in billing row.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    populatePastBalanceAndCredit(row, headerMap, prevRow, prevHeaderMap) -> void
        Populates past balance and credit from previous billing cycle.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processDocumentSelection(studentId, billingCycleName) -> void
        Processes document selection from checkbox dialog.
        Category: GENERATE_DOCUMENTS
        Local functions used: executeDocumentGeneration()
        Utility functions used: None

    processFieldMapForSemester(semesterName) -> void
        Processes and stores field map for semester form responses.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processFormsData(formsData, billingData) -> Array
        Processes form submissions and compares against billing data.
        Returns array of form reconciliation results.
        Category: RECONCILIATION
        Local functions used: processFormsReconciliationForRow()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processFormsReconciliationForRow(studentId, formData, billingRow, headerMap) -> Object
        Processes single student's form reconciliation.
        Returns reconciliation result object.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processPaymentReconciliationForRow(payment, billingSheet, billingData) -> Object
        Processes single payment record against billing data.
        Returns payment reconciliation result.
        Category: RECONCILIATION
        Local functions used: processPaymentRecord()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processPaymentRecord(payment, billingSheet, studentRow, headerMap) -> void
        Processes payment and updates billing sheet.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processSemesterEndCredits(billingRow, headerMap, previousBalance) -> void
        Processes credits from previous semester end balances.
        Category: BILLING_CYCLE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processTeacherAttendanceForBilling(teacherId, billingSheet, billingData) -> Object
        Processes teacher's attendance data and reconciles with billing.
        Returns reconciliation results for teacher.
        Category: RECONCILIATION
        Local functions used: locateStudentRecordEnhanced()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processTeacherForNewAttendance(teacherData) -> void
        Processes teacher for new attendance sheet generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    processTeacherReconciliation(teacherWorkbook, teacherId, billingCycleName, billingSheet) -> Object
        Reconciles teacher's roster and attendance with billing data.
        Returns reconciliation summary for teacher.
        Category: RECONCILIATION
        Local functions used: processTeacherAttendanceForBilling()
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    reconcilePayment(studentRow, headerMap, amount, paymentDate) -> void
        Reconciles single payment against student's billing row.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    renameLatestFormSheet(semesterName) -> void
        Renames most recent form responses sheet to include semester name.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    runBillingCycleAutomation() -> void
        Main function to run complete billing cycle automation process.
        Category: UI_MENU
        Local functions used: promptForCustomToday(), promptForBillingCycleName(), 
                              createBillingSheet(), populateBillingSheet()
        Utility functions used: UtilityScriptLibrary.debugLog()

    runCombinedReconciliation(billingCycleName?) -> void
        Runs combined reconciliation of payments, lessons, and forms.
        Category: RECONCILIATION
        Local functions used: runPaymentReconciliation(), runWeeklyLessonReconciliation(), 
                              runFormsReconciliation()
        Utility functions used: UtilityScriptLibrary.debugLog()

    runFormsReconciliation(billingCycleName?) -> Object
        Reconciles form submissions with billing data.
        Returns reconciliation results object.
        Category: UI_MENU
        Local functions used: getFormsDataFromContacts(), processFormsData()
        Utility functions used: UtilityScriptLibrary.debugLog()

    runFormsReconciliationUI() -> void
        UI wrapper for forms reconciliation with user prompts.
        Category: UI_MENU
        Local functions used: runFormsReconciliation()
        Utility functions used: None

    runFullReconciliation(billingCycleName?) -> void
        Runs complete reconciliation process (payments, lessons, forms).
        Category: UI_MENU
        Local functions used: runPaymentReconciliation(), runWeeklyLessonReconciliation(), 
                              runFormsReconciliation()
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

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

    setupNewSemester() -> void
        Main function to set up new semester with all required configuration.
        Category: UI_MENU
        Local functions used: promptForSemesterName(), promptForSemesterDates(), 
                              appendToSemesterMetadata()
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    sumPayments(studentId, throughDate?) -> Number
        Sums all payments for student up to specified date.
        Returns total payment amount.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

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
        Utility functions used: UtilityScriptLibrary.debugLog()

    updateDocIdInBillingSheet(studentId, docType, docId) -> void
        Updates document ID in billing sheet after document generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: getDocIdColumnName(), getCurrentBillingSheet()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.debugLog()

    updateInvoiceUrlInBillingSheet(studentId, invoiceUrl) -> void
        Updates invoice URL in billing sheet after invoice generation.
        Category: GENERATE_DOCUMENTS
        Local functions used: getCurrentBillingSheet()
        Utility functions used: UtilityScriptLibrary.getHeaderMap()

    updateSheetStudentWarnings(sheet, warningStudents) -> void
        Updates attendance sheet with warning highlights for students.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    updateTeacherRosterBalances(teacherWorkbook, teacherData) -> void
        Updates teacher roster with current balance information.
        Category: RECONCILIATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    verifyProgramsForSemester(semesterName) -> Boolean
        Verifies all required programs are set up for semester.
        Returns true if verification passed.
        Category: SETUP_SEMESTER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    verifyRatesEnhanced() -> Boolean
        Verifies all rates are properly configured for billing.
        Returns true if verification passed.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()
================================================================================
END OF FUNCTION DIRECTORY
================================================================================    
