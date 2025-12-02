function diagnosticCheckQtyColumns() {
  // Access the Responses spreadsheet
  var responsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
  var formSheet = responsesSS.getSheetByName('Summer 2025');
  
  if (!formSheet) {
    UtilityScriptLibrary.debugLog('diagnosticCheckQtyColumns', 'ERROR', 
      'Sheet "Summer 2025" not found. Available: ' + responsesSS.getSheets().map(function(s) { return s.getName(); }).join(', '), 
      '', '');
    return;
  }
  
  // Get the fieldMap properly
  var fieldMapSheet = responsesSS.getSheetByName('FieldMap');
  if (!fieldMapSheet) {
    UtilityScriptLibrary.debugLog('diagnosticCheckQtyColumns', 'ERROR', 
      'FieldMap sheet not found in Form Responses workbook', '', '');
    return;
  }
  var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
  
  var headers = formSheet.getDataRange().getValues()[0];
  
  // Find columns that contain "30-minute" or "45-minute" or "60-minute"
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var headerStr = header.toString();
    
    if (headerStr.indexOf('30-minute') > -1 || 
        headerStr.indexOf('45-minute') > -1 || 
        headerStr.indexOf('60-minute') > -1) {
      
      // FIXED: Normalize the header before looking it up
      var normalizedHeader = UtilityScriptLibrary.normalizeHeader(header);
      var mappedName = fieldMap[normalizedHeader] || 'NOT FOUND';
      var sampleValue = formSheet.getLastRow() > 1 ? formSheet.getRange(2, i + 1).getValue() : 'N/A';
      
      Logger.log('===== COLUMN ' + i + ' =====');
      Logger.log('Header: ' + JSON.stringify(header));
      Logger.log('Normalized: ' + normalizedHeader);
      Logger.log('Mapped to: ' + mappedName);
      Logger.log('Sample value: ' + JSON.stringify(sampleValue));
      Logger.log('');
    }
  }
  
  Logger.log('Diagnostic complete');
}

/* 
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
*/

// ============================================================================
// SECTION 1: UI MENU
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('QAMP Tools')
    .addItem('Set Up New Semester', 'setupNewSemester')
    .addItem('Begin New Billing Cycle', 'runBillingCycleAutomation')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Reconciliation')
        .addItem('Full Reconciliation', 'runFullReconciliationUI')
        .addSeparator()
        .addItem('Lessons Only', 'runWeeklyLessonReconciliationUI')
        .addItem('Payments Only', 'runPaymentReconciliationUI')
        .addItem('Forms Only', 'runFormsReconciliationUI')
    )
    .addSeparator()
    .addItem('Generate Documents', 'runRegistrationPacketGenerationUI')
    .addItem('Print Documents', 'convertFolderDocsToPdfUI')
    .addItem('Create New Attendance Sheets', 'createNewAttendanceSheets')
    .addToUi();
}

function runBillingCycleAutomation() {
  try {
    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "INFO", "Starting billing cycle automation", "", "");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    var customToday = promptForCustomToday();
    var billingCycleName = promptForBillingCycleName(customToday);
    if (!billingCycleName) return;

    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "DEBUG", "User inputs confirmed", 
                  "Date: " + customToday.toLocaleDateString() + ", Cycle: " + billingCycleName, "");

    var semesterSheet = ss.getSheetByName('Semester Metadata');
    if (!semesterSheet) throw new Error('Semester Metadata sheet not found.');
    var semesterData = semesterSheet.getDataRange().getValues();
    if (semesterData.length < 2) throw new Error('No semesters found in Semester Metadata.');
    
    var semesterName = semesterData[semesterData.length - 1][0];

    verifyProgramsForSemester(semesterName, ss);

    if (!verifyRatesEnhanced()) {
      ui.alert('Please verify the Rates sheet before proceeding.');
      return;
    }

    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "DEBUG", "All validations passed", "", "");

    createBillingSheet(billingCycleName);
    var newBillingSheet = ss.getSheetByName(billingCycleName);
    ss.setActiveSheet(newBillingSheet);

    protectPreviousBillingCycle();

    var billingMetaSheet = ss.getSheetByName('Billing Metadata');
    var billingMetaData = billingMetaSheet.getDataRange().getValues();
    
    var metaHeaderMap = UtilityScriptLibrary.getHeaderMap(billingMetaSheet);
    var semesterCol = metaHeaderMap[UtilityScriptLibrary.normalizeHeader("Semester Name")];

    if (!semesterCol) {
      UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "ERROR", "Semester Name column not found in Billing Metadata", "", "");
      throw new Error("Semester Name column not found in Billing Metadata sheet");
    }

    var isFirstCycle = !billingMetaData.some(function(row) { 
      return row[semesterCol - 1] === semesterName; 
    });

    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "DEBUG", "Billing cycle type determined", 
                  "First cycle: " + isFirstCycle, "");

    var carryOverData = extractPreviousBillingData({ includeAll: !isFirstCycle });
    if (!carryOverData || !carryOverData.previousSheetName) {
      if (isFirstCycle) {
        UtilityScriptLibrary.debugLog('No previous billing data found. Initializing empty carryOverData for first cycle.');
        carryOverData = { rowsToCarry: {}, prevHeaderMap: {}, previousSheetName: null };
      } else {
        throw new Error('Could not retrieve previous billing data.');
      }
    }

    // NEW: Prompt for forms carry-over if this is the first cycle of a new semester
    var carryOverForms = true; // Default to true for within-semester cycles
    if (isFirstCycle && carryOverData && carryOverData.previousSheetName) {
      var response = ui.alert(
        'Forms Data Carry-Over',
        'This is the first billing cycle of a new semester.\n\n' +
        'Do you want to carry over Agreement Form and Media Release status from the previous semester?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        carryOverForms = true;
        UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "INFO", "User selected forms carry-over", 
                      "Carry over: true", "");
      } else if (response === ui.Button.NO) {
        carryOverForms = false;
        UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "INFO", "User selected forms carry-over", 
                      "Carry over: false", "");
      } else {
        ui.alert('❌ Billing cycle cancelled.');
        return;
      }
    }

    // NEW: Update cumulative tracking from previous cycle before building new rows
    if (carryOverData.previousSheetName) {
      updateCumulativeTracking(carryOverData.previousSheetName);
      UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "INFO", "Cumulative tracking updated", 
                    "From sheet: " + carryOverData.previousSheetName, "");
    }

    ss.setActiveSheet(newBillingSheet);
    var context = buildBillingContext(customToday, semesterName, billingCycleName);
    context.prevHeaderMap = carryOverData.previousSheetName
      ? UtilityScriptLibrary.getHeaderMap(ss.getSheetByName(carryOverData.previousSheetName))
      : {};
    context.carryOverForms = carryOverForms; // NEW: Add forms carry-over flag to context

    ss.setActiveSheet(newBillingSheet);
    var context = buildBillingContext(customToday, semesterName, billingCycleName);
    context.prevHeaderMap = carryOverData.previousSheetName
      ? UtilityScriptLibrary.getHeaderMap(ss.getSheetByName(carryOverData.previousSheetName))
      : {};

    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "DEBUG", "Context built", 
                  "Previous sheet: " + (carryOverData.previousSheetName || "None"), "");

    if (isFirstCycle) {
      populateBillingSheet(context, carryOverData);
    } else {
      var previousStartDate = new Date(context.billingStartDate);
      var existingStudentIds = [];
      populateBillingSheetContinuingSemester(context, newBillingSheet, existingStudentIds, carryOverData.rowsToCarry, null);
    }

    appendToBillingMetadata(billingCycleName, customToday, semesterName);
    
    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "INFO", "Billing cycle completed successfully", 
                  "Cycle: " + billingCycleName, "");
    
    ui.alert("New billing cycle \"" + billingCycleName + "\" successfully created.");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("runBillingCycleAutomation", "ERROR", "Billing cycle failed", 
                  "", error.message + " | " + error.stack);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
    UtilityScriptLibrary.debugLog("Error in runBillingCycleAutomation: " + error.stack);
  }
}

function runFullReconciliation(reconciliationDate) {
  try {
    UtilityScriptLibrary.debugLog("runFullReconciliation", "INFO", "Starting full reconciliation", 
                 "Date: " + reconciliationDate, "");
    
    var ui = SpreadsheetApp.getUi();
    var results = {
      attendance: null,
      combined: null,
      success: false,
      errors: []
    };
    
    // 1. Run attendance reconciliation
    try {
      UtilityScriptLibrary.debugLog("runFullReconciliation", "INFO", "Running attendance reconciliation", "", "");
      results.attendance = runWeeklyLessonReconciliation(reconciliationDate);
      UtilityScriptLibrary.debugLog("runFullReconciliation", "SUCCESS", "Attendance reconciliation completed", "", "");
    } catch (error) {
      results.errors.push("Attendance reconciliation failed: " + error.message);
      UtilityScriptLibrary.debugLog("runFullReconciliation", "ERROR", "Attendance reconciliation failed", "", error.message);
    }
    
    // 2. Run combined payment and forms reconciliation
    try {
      UtilityScriptLibrary.debugLog("runFullReconciliation", "INFO", "Running combined payment and forms reconciliation", "", "");
      results.combined = runCombinedReconciliation();
      UtilityScriptLibrary.debugLog("runFullReconciliation", "SUCCESS", "Combined reconciliation completed", "", "");
    } catch (error) {
      results.errors.push("Payment/Forms reconciliation failed: " + error.message);
      UtilityScriptLibrary.debugLog("runFullReconciliation", "ERROR", "Combined reconciliation failed", "", error.message);
    }
    
    // Generate summary
    var summary = generateReconciliationSummaryUpdated(results);
    results.success = (results.errors.length === 0);
    
    // Show results to user
    ui.alert(
      results.success ? '✅ Reconciliation Complete' : '⚠️ Reconciliation Completed with Errors',
      summary,
      ui.ButtonSet.OK
    );
    
    UtilityScriptLibrary.debugLog("runFullReconciliation", results.success ? "SUCCESS" : "WARNING", 
                 "Full reconciliation completed", "Errors: " + results.errors.length, "");
    
    return results;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("runFullReconciliation", "ERROR", "Full reconciliation failed", "", error.message);
    SpreadsheetApp.getUi().alert('❌ Reconciliation Error', 'Failed to run reconciliation: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
    throw error;
  }
}

function runFullReconciliationUI() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Prompt for reconciliation date
    var response = ui.prompt(
      'Full Reconciliation',
      'Enter the date to reconcile up to (MM/DD/YYYY):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('❌ Reconciliation cancelled.');
      return;
    }
    
    // Parse the date
    var dateString = response.getResponseText().trim();
    var reconciliationDate = UtilityScriptLibrary.parseDateFromString(dateString);
    
    if (!reconciliationDate) {
      ui.alert('❌ Invalid date format. Please use MM/DD/YYYY.');
      return;
    }
    
    // Run full reconciliation
    runFullReconciliation(reconciliationDate);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("runFullReconciliationUI", "ERROR", "UI reconciliation failed", "", error.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function runRegistrationPacketGenerationUI() {
  try {
    generateRegistrationPacketsForBillingCycle();
  } catch (error) {
    UtilityScriptLibrary.debugLog("runRegistrationPacketGenerationUI", "ERROR", "Registration packet UI failed", "", error.message);
    SpreadsheetApp.getUi().alert('Error', 'Registration packet generation failed: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function runPaymentReconciliation() {
  try {
    UtilityScriptLibrary.debugLog('💰 Starting payment reconciliation...');
    
    // Get payment workbook and current semester sheet
    var paymentsSS = UtilityScriptLibrary.getWorkbook('payments');
    UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'Got payments workbook', '', '');
    
    var currentSemester = getCurrentSemesterName();
    UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'Current semester', currentSemester, '');
    
    var paymentSheet = paymentsSS.getSheetByName(currentSemester);
    
    if (!paymentSheet) {
      throw new Error('Payment sheet not found for semester: ' + currentSemester);
    }
    UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'Found payment sheet', currentSemester, '');
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(paymentSheet);
    UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'Got header map', 'Keys: ' + Object.keys(headerMap).length, '');
    
    var data = paymentSheet.getDataRange().getValues();
    UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'Got payment data', 'Rows: ' + data.length, '');
    
    var processed = 0;
    var errors = [];
    var results = {
      processed: 0,
      errors: 0,
      details: []
    };
    
    // Process each payment record (skip header row)
    for (var i = 1; i < data.length; i++) {
      var rowData = data[i];
      var rowNumber = i + 1;
      
      UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'Processing row', 'Row: ' + rowNumber, '');
      
      try {
        var result = processPaymentRecord(rowData, paymentSheet, headerMap, rowNumber);
        UtilityScriptLibrary.debugLog('runPaymentReconciliation', 'DEBUG', 'processPaymentRecord result', 
                     'Row: ' + rowNumber + ', Processed: ' + result.processed + ', Message: ' + result.message, '');
        
        if (result.processed) {
          results.processed++;
          results.details.push(result.message);
        }
      } catch (error) {
        results.errors++;
        results.details.push('Row ' + rowNumber + ': ' + error.message);
        UtilityScriptLibrary.debugLog('❌ Error processing row ' + rowNumber + ': ' + error.message);
      }
    }
    
    UtilityScriptLibrary.debugLog('✅ Payment reconciliation completed. Processed: ' + results.processed + ', Errors: ' + results.errors);
    return results;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Payment reconciliation failed: ' + error.message);
    throw error;
  }
}

function runPaymentReconciliationUI() {
  try {
    var results = runPaymentReconciliation();
    var ui = SpreadsheetApp.getUi();
    ui.alert('💰 Payment Reconciliation Complete', 
             'Processed: ' + results.processed + ' payments\nErrors: ' + results.errors, 
             ui.ButtonSet.OK);
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function runWeeklyLessonReconciliationUI() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Prompt for admin review date
    var response = ui.prompt(
      'Admin Review Date',
      'Enter the date to pull lessons up to (MM/DD/YYYY):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('❌ Lesson reconciliation cancelled.');
      return;
    }
    
    // Parse the date
    var dateString = response.getResponseText().trim();
    var reconciliationDate = UtilityScriptLibrary.parseDateFromString(dateString);
    
    if (!reconciliationDate) {
      ui.alert('❌ Invalid date format. Please use MM/DD/YYYY.');
      return;
    }
    
    // Confirm the date
    var formattedDate = UtilityScriptLibrary.formatDateFlexible(reconciliationDate, "MM/dd/yyyy");
    var confirm = ui.alert(
      `🔄 Reconcile lessons up to ${formattedDate}?`,
      ui.ButtonSet.YES_NO
    );
    
    if (confirm !== ui.Button.YES) {
      ui.alert('❌ Lesson reconciliation cancelled.');
      return;
    }
    
    // Run the reconciliation
    runWeeklyLessonReconciliation(reconciliationDate);
    
    // Success message
    ui.alert(`✅ Lesson reconciliation completed for ${formattedDate}!`);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog(`❌ Error in lesson reconciliation UI: ${error.message}`);
    SpreadsheetApp.getUi().alert(`❌ Error: ${error.message}`);
  }
}

function getNextMonthName(targetDate) {
  var monthNames = UtilityScriptLibrary.getMonthNames();
  
  var nextMonth = targetDate.getMonth() + 1;
  if (nextMonth > 11) {
    nextMonth = 0; // Wrap to January
  }
  
  return monthNames[nextMonth];
}

function setupNewSemester() {
  var semesterName = promptForSemesterName();
  if (!semesterName) return;

  try {
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Starting setup for semester', semesterName, '');

    // Process previous semester credits before creating new semester
    var creditBalances = processSemesterEndCredits(getCurrentSemesterName());
    
    // FIXED: Handle null return value safely
    if (creditBalances && creditBalances.length > 0) {
      SpreadsheetApp.getUi().alert('📊 SEMESTER CREDITS PROCESSED\n\nProcessed lesson balances for ' + creditBalances.length + ' students.\nCredits will be applied to new semester registrations.');
    }

    // Step 1: Prompt for dates
    var semesterDates = promptForSemesterDates();
    
    if (!semesterDates) {
      throw new Error('❌ No semester dates returned from prompt.');
    }
    
    var startDate = semesterDates.startDate;
    var endDate = semesterDates.endDate;
    
    // Force conversion to Date objects (handles serialization issues)
    startDate = new Date(startDate);
    endDate = new Date(endDate);
    
    UtilityScriptLibrary.debugLog('🔧 Converted dates - startDate: ' + startDate + ' (instanceof Date: ' + (startDate instanceof Date) + ')');
    UtilityScriptLibrary.debugLog('🔧 Converted dates - endDate: ' + endDate + ' (instanceof Date: ' + (endDate instanceof Date) + ')');
    
    // Validate dates
    if (!startDate || !(startDate instanceof Date) || isNaN(startDate.getTime())) {
      throw new Error('❌ Invalid startDate after conversion. Value: ' + startDate);
    }
    if (!endDate || !(endDate instanceof Date) || isNaN(endDate.getTime())) {
      throw new Error('❌ Invalid endDate after conversion. Value: ' + endDate);
    }

    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Semester dates confirmed', 
                                  'Start: ' + startDate + ', End: ' + endDate, '');

    // Step 2: Generate Calendar
    generateCalendarForSemester(semesterName, startDate, endDate);
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Calendar generated', '', '');

    // Step 3: Rename Form Sheet
    renameLatestFormSheet(semesterName);
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Form response sheet renamed', '', '');

    // Step 4: Process Field Map
    processFieldMapForSemester(semesterName);
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Field map processed', '', '');

    // Step 5: Create Payments Tab
    createPaymentsTab(semesterName);
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Payments tab created', '', '');

    // Step 6: Metadata Append
    var folderId = appendToSemesterMetadata(semesterName, startDate, endDate);
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Semester metadata appended', 
                                  'Folder ID: ' + folderId, '');

    // Step 7: Mark Students Inactive
    markStudentsInactive();
    UtilityScriptLibrary.debugLog('setupNewSemester', 'INFO', 'Students marked inactive', '', '');

    // Step 8: Final UI Confirmation
    SpreadsheetApp.getUi().alert('✅ New semester "' + semesterName + '" setup completed.');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('setupNewSemester', 'ERROR', 'Setup failed', semesterName, error.message);
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

function setupRosterTemplateProtection(sheet) {
  try {
    // Protect admin columns (E through U) with warning - UPDATED range
    var adminRange = sheet.getRange(1, 5, sheet.getMaxRows(), 17); // Columns E-U (17 columns)
    var protection = adminRange.protect();
    protection.setDescription('Admin columns - automated data only');
    protection.setWarningOnly(true);
    
    // Set up date validation for First Lesson Date column (B)
    var dateRange = sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1);
    var dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();
    dateRange.setDataValidation(dateRule);
    
    // Set up dropdown validation for Status column (T)
    var statusRange = sheet.getRange(2, 20, sheet.getMaxRows() - 1, 1);
    var statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['active', 'dropped'], true)
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    UtilityScriptLibrary.debugLog("✅ Roster protection, date validation, and status dropdown applied");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Error in roster protection: " + error.message);
  }
}

// ============================================================================
// SECTION 2: SET UP NEW SEMESTER MID-LEVEL AND HELPER FUNCTIONS
// ============================================================================

function appendToSemesterMetadata(semesterName, startDate, endDate) {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    
    var billingSS = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = UtilityScriptLibrary.getSheet('semesterMetadata', billingSS);
    if (!metadataSheet) {
      throw new Error('Semester Metadata sheet not found.');
    }

    // Generate folder using existing utility function
    var folderResult = createRosterFolder(semesterName);
    UtilityScriptLibrary.debugLog('appendToSemesterMetadata', 'INFO', 'Roster folder created', 
                                  'Folder ID: ' + folderResult.folderId, '');

    // Verify rates using utility functions
    var rateChartName = getCurrentRateChartName();
    var rateSummary = UtilityScriptLibrary.getRateSummary();
    var ui = SpreadsheetApp.getUi();
    
    var rateConfirm = ui.alert(
      'Current Rates (' + rateChartName + '):\n\n' + rateSummary + '\n\nDo you confirm these are correct for "' + semesterName + '"?',
      ui.ButtonSet.YES_NO
    );

    if (rateConfirm !== ui.Button.YES) {
      throw new Error('Semester setup cancelled - rates not confirmed.');
    }

    // Verify programs using extracted helper function
    var programVerificationResult = verifyProgramsForSemester(semesterName, billingSS);
    
    // Build and append metadata row
    var rowData = [
      semesterName,
      startDate,
      endDate,
      rateChartName,
      programVerificationResult.status
    ];

    metadataSheet.appendRow(rowData);

    UtilityScriptLibrary.debugLog('appendToSemesterMetadata', 'INFO', 'Semester metadata appended', 
                                  'Data: ' + JSON.stringify(rowData), '');
    
    return {
      folderId: folderResult.folderId,
      rateChart: rateChartName,
      programStatus: programVerificationResult.status,
      activePrograms: programVerificationResult.activePrograms
    };
    
  }, 'Semester metadata appended successfully', 'appendToSemesterMetadata', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

function convertFolderDocsToPdfUI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeTabName = ss.getActiveSheet().getName();
  
  // Get environment automatically from EnvironmentManager
  var env = UtilityScriptLibrary.EnvironmentManager.get();
  var config = UtilityScriptLibrary.getConfig();
  var parentFolder = DriveApp.getFolderById(config[env].generatedDocumentsFolderId);
  
  // Navigate to Student Documents folder
  var studentDocsIterator = parentFolder.getFoldersByName('Student Documents');
  if (!studentDocsIterator.hasNext()) {
    SpreadsheetApp.getUi().alert('Student Documents folder not found');
    return;
  }
  var studentDocsFolder = studentDocsIterator.next();
  
  // Find the folder matching the active tab name
  var folderIterator = studentDocsFolder.getFoldersByName(activeTabName);
  
  if (!folderIterator.hasNext()) {
    SpreadsheetApp.getUi().alert('Folder "' + activeTabName + '" not found in Student Documents');
    return;
  }
  
  var billingFolder = folderIterator.next();
  
  // Create a "[folder name] PDFs" subfolder if it doesn't exist
  var pdfFolderName = activeTabName + ' PDFs';
  var pdfFolder;
  var pdfFolderIterator = billingFolder.getFoldersByName(pdfFolderName);
  if (pdfFolderIterator.hasNext()) {
    pdfFolder = pdfFolderIterator.next();
  } else {
    pdfFolder = billingFolder.createFolder(pdfFolderName);
  }
  
  // Get existing PDFs to check for duplicates
  var existingPdfs = {};
  var pdfFiles = pdfFolder.getFiles();
  while (pdfFiles.hasNext()) {
    existingPdfs[pdfFiles.next().getName()] = true;
  }
  
  // Get all Google Docs files in the billing folder
  var files = billingFolder.getFilesByType(MimeType.GOOGLE_DOCS);
  var convertedCount = 0;
  var skippedCount = 0;
  
  while (files.hasNext()) {
    var file = files.next();
    var docId = file.getId();
    var docName = file.getName();
    var pdfName = docName + '.pdf';
    
    // Check if PDF already exists
    if (existingPdfs[pdfName]) {
      skippedCount++;
      continue;
    }
    
    // Get the document and convert to PDF
    var doc = DocumentApp.openById(docId);
    var pdfBlob = doc.getAs('application/pdf');
    pdfBlob.setName(pdfName);
    
    // Save PDF to the PDFs folder
    pdfFolder.createFile(pdfBlob);
    convertedCount++;
  }
  
  var envLabel = env.toUpperCase();
  SpreadsheetApp.getUi().alert(
    '[' + envLabel + '] Converted ' + convertedCount + ' new document(s) to PDF.\n' +
    'Skipped ' + skippedCount + ' already converted.\n' +
    'PDF folder: ' + pdfFolder.getUrl()
  );
}

function createPaymentsTab(semesterName) {
  try {
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'INFO', 'Starting payments tab creation', 
                                  'Semester: ' + semesterName, '');

    var paymentsSS = UtilityScriptLibrary.getWorkbook('payments');
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Retrieved payments workbook', 
                                  'ID: ' + paymentsSS.getId(), '');
    
    var template = paymentsSS.getSheetByName('Ledger Template');
    
    if (!template) {
      UtilityScriptLibrary.debugLog('createPaymentsTab', 'ERROR', 'Ledger Template sheet not found', '', '');
      throw new Error('Ledger Template sheet not found in Payments spreadsheet.');
    }
    
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Found template sheet', 
                                  'Template exists: true', '');

    // Check if sheet already exists
    var existingSheet = paymentsSS.getSheetByName(semesterName);
    if (existingSheet) {
      UtilityScriptLibrary.debugLog('createPaymentsTab', 'ERROR', 'Sheet already exists', 
                                    'Sheet: ' + semesterName, '');
      throw new Error('A sheet named "' + semesterName + '" already exists.');
    }
    
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Sheet name is available', 
                                  'Sheet: ' + semesterName, '');

    // Set template as active and duplicate it
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Setting template as active sheet', '', '');
    paymentsSS.setActiveSheet(template);
    
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Duplicating template sheet', '', '');
    var newSheet = paymentsSS.duplicateActiveSheet();
    
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Renaming duplicated sheet', 
                                  'New name: ' + semesterName, '');
    newSheet.setName(semesterName);
    
    // Move new sheet to position after template (position 2)
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Moving sheet to position 2', '', '');
    paymentsSS.setActiveSheet(newSheet);
    paymentsSS.moveActiveSheet(2);
    
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Flushing changes', '', '');
    SpreadsheetApp.flush();

    // Verify sheet was created
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Verifying sheet creation', '', '');
    var verifySheet = paymentsSS.getSheetByName(semesterName);
    if (!verifySheet) {
      UtilityScriptLibrary.debugLog('createPaymentsTab', 'ERROR', 'Sheet verification failed', 
                                    'Sheet not found after creation', '');
      throw new Error('Sheet creation verification failed - sheet "' + semesterName + '" not found after creation.');
    }

    UtilityScriptLibrary.debugLog('createPaymentsTab', 'INFO', 'Sheet created and verified', 
                                  'Sheet: ' + semesterName, '');

    // Apply basic sheet setup
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Freezing header row', '', '');
    newSheet.setFrozenRows(1);

    // Set up protections
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Applying sheet protections', '', '');
    var protectionConfig = [
      { range: '1:1', description: 'Protect Header Row' },
      { range: 'H:K', description: 'Protect Admin Columns (H through K)' }
    ];

    for (var i = 0; i < protectionConfig.length; i++) {
      var config = protectionConfig[i];
      var range = newSheet.getRange(config.range);
      var protection = range.protect().setDescription(config.description);
      protection.setWarningOnly(true);
      UtilityScriptLibrary.debugLog('createPaymentsTab', 'DEBUG', 'Applied protection', 
                                    'Range: ' + config.range, '');
    }

    UtilityScriptLibrary.debugLog('createPaymentsTab', 'SUCCESS', 'Payments tab created successfully', 
                                  'Sheet: ' + semesterName + ', Protections: ' + protectionConfig.length, '');

    return {
      sheetName: semesterName,
      protectedRanges: protectionConfig.length,
      success: true
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('createPaymentsTab', 'ERROR', 'Failed to create payments tab', 
                                  'Semester: ' + semesterName, error.message + ' | ' + error.stack);
    throw error;
  }
}

function createRosterFolder(semesterName) {
   logPrefix = 'createRosterFolder("' + semesterName + '")';

  try {
    UtilityScriptLibrary.debugLog(logPrefix + ' - Accessing parent folder via UtilityScriptLibrary...');
    var parentFolder = UtilityScriptLibrary.getRosterFolder();

    UtilityScriptLibrary.debugLog(logPrefix + ' - Extracting year from semester name...');
    var year = UtilityScriptLibrary.getYearFromSemesterName(semesterName);
    if (!year) throw new Error('Year could not be parsed from "' + semesterName + '"');

    var folderName = year + ' Rosters';
    UtilityScriptLibrary.debugLog(logPrefix + ' - Searching for existing folder: "' + folderName + '"');
    var folders = parentFolder.getFoldersByName(folderName);
    var yearFolder = folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);

    var folderId = yearFolder.getId();
    var folderUrl = 'https://drive.google.com/drive/folders/' + folderId;
    UtilityScriptLibrary.debugLog(logPrefix + ' - Folder resolved: ' + folderUrl);

    UtilityScriptLibrary.debugLog(logPrefix + ' - Logging folder info to Year Metadata...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearSheet = ss.getSheetByName('Year Metadata');
    if (!yearSheet) {
      yearSheet = ss.insertSheet('Year Metadata');
      // Add headers if new sheet
      yearSheet.getRange(1, 1, 1, 3).setValues([['Year', 'Roster Folder ID', 'Roster Folder URL']]);
      yearSheet.setFrozenRows(1);
    }

    var data = yearSheet.getDataRange().getValues();
    var yearString = String(year);
    
    // Find existing row (skip header row at index 0)
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === yearString) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex !== -1) {
      // Update existing row (rowIndex is 0-based, sheet rows are 1-based)
      yearSheet.getRange(rowIndex + 1, 2, 1, 2).setValues([[folderId, folderUrl]]);
      UtilityScriptLibrary.debugLog(logPrefix + ' - Updated existing entry for "' + year + '" at row ' + (rowIndex + 1));
    } else {
      // Append new row
      yearSheet.appendRow([year, folderId, folderUrl]);
      UtilityScriptLibrary.debugLog(logPrefix + ' - Appended new entry for "' + year + '"');
    }

    return { folderId: folderId, folderUrl: folderUrl };

  } catch (e) {
    UtilityScriptLibrary.debugLog(logPrefix + ' âŒ Error: ' + e.message);
    throw e;
  }
}

function generateCalendarForSemester(semesterName, startDate, endDate) {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    UtilityScriptLibrary.debugLog('📅 generateCalendarForSemester - Starting calendar generation for: ' + semesterName);
    
    var calendarSheet = UtilityScriptLibrary.getSheet('calendar');
    if (!calendarSheet) {
      UtilityScriptLibrary.debugLog('📅 generateCalendarForSemester - Calendar sheet not found');
      throw new Error('Calendar sheet not found in Responses workbook.');
    }

    var sheet = calendarSheet;
    var lastRow = sheet.getLastRow();
    
    UtilityScriptLibrary.debugLog('📅 generateCalendarForSemester - Processing existing calendar data, lastRow: ' + lastRow);
    
    if (lastRow >= 2) {
      // FIXED: Shift existing calendar INCLUDING HEADERS (B1:D) → E1:G
      var existingRange = sheet.getRange(1, 2, lastRow, 3); // Start from row 1 to include headers
      sheet.insertColumnsAfter(4, 3); // Insert E:G
      existingRange.copyTo(sheet.getRange(1, 5)); // Copy to E1, including headers and formatting
      sheet.getRange(1, 2, lastRow, 3).clearContent(); // Clear B1:D including headers
      UtilityScriptLibrary.debugLog('📅 generateCalendarForSemester - Shifted existing calendar data with headers');
    }

    // Write headers in B1:D1
    sheet.getRange(1, 2, 1, 3).setValues([['Week Start', 'Week End', 'Semester']]);
    UtilityScriptLibrary.debugLog('📅 generateCalendarForSemester - Headers written');

    // Generate week data
    var rows = [];
    var currentDate = new Date(startDate);

    while (currentDate <= endDate) {
      var weekStart = new Date(currentDate);
      var weekEnd = new Date(weekStart);
      weekEnd.setDate(weekEnd.getDate() + 6);
      if (weekEnd > endDate) {
        weekEnd.setTime(endDate.getTime());
      }

      rows.push([
        Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
        Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
        semesterName
      ]);

      currentDate.setDate(currentDate.getDate() + 7);
    }

    // Write new semester rows starting from row 2, columns B–D
    sheet.getRange(2, 2, rows.length, 3).setValues(rows);
    
    UtilityScriptLibrary.debugLog('📅 generateCalendarForSemester - Calendar populated with ' + rows.length + ' weeks for ' + semesterName);
    
    return {
      weeks: rows.length,
      semesterName: semesterName
    };
    
  }, 'Calendar generated successfully for ' + semesterName, 'generateCalendarForSemester', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

function markStudentsInactive() {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    
    // Get the students sheet
    var studentsSheet = UtilityScriptLibrary.getSheet('students');
    
    // FIXED: Use the correct function name from Utility-0043.txt
    var result = UtilityScriptLibrary.bulkUpdateStudentStatus(
      studentsSheet,               // sheet
      'Currently Registered',      // columnName
      false,                       // newValue (set to false for inactive)
      {
        skipEmpty: true            // Only update rows that have data
      }
    );

    if (!result.success) {
      throw new Error('Failed to mark students inactive: ' + result.error);
    }

    UtilityScriptLibrary.debugLog('markStudentsInactive', 'INFO', 'Students marked inactive', 
                                  'Students processed: ' + result.updatedCount, '');

    return {
      studentsProcessed: result.updatedCount,
      changedRows: result.changedRows,
      message: result.message
    };
    
  }, 'Students marked inactive successfully', 'markStudentsInactive', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

function processFieldMapForSemester(semesterName) {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    
    var formResponsesWorkbook = UtilityScriptLibrary.getWorkbook('formResponses');
    var formSheet = formResponsesWorkbook.getSheetByName(semesterName);
    if (!formSheet) {
      throw new Error('Form sheet "' + semesterName + '" not found in Form Responses workbook.');
    }

    // FIXED: Use correct SHEET_MAP key
    var fieldMapSheet = UtilityScriptLibrary.getSheet('fieldMap');
    if (!fieldMapSheet) {
      throw new Error('Field Map sheet not found.');
    }

    var headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    var existingFieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
    var existingHeaders = Object.keys(existingFieldMap);
    
    var newRows = [];
    var duplicates = [];
    
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i];
      var normalizedHeader = UtilityScriptLibrary.normalizeHeader(header);
      
      var isDuplicate = false;
      for (var j = 0; j < existingHeaders.length; j++) {
        if (UtilityScriptLibrary.normalizeHeader(existingHeaders[j]) === normalizedHeader) {
          isDuplicate = true;
          duplicates.push(header);
          break;
        }
      }
      
      if (!isDuplicate) {
        newRows.push([header, '', false]);
      }
    }
    
    if (newRows.length > 0) {
      var lastRow = fieldMapSheet.getLastRow();
      var newRange = fieldMapSheet.getRange(lastRow + 1, 1, newRows.length, 3);
      newRange.setValues(newRows);
    }

    // Reset Updated? column if it exists - reset flags to false
    var headers = fieldMapSheet.getRange(1, 1, 1, fieldMapSheet.getLastColumn()).getValues()[0];
    var updatedCol = -1;
    var updatedColumnProcessed = false;
    
    for (var k = 0; k < headers.length; k++) {
      if (UtilityScriptLibrary.normalizeHeader(headers[k]) === "updated?") {
        updatedCol = k + 1;
        break;
      }
    }
    
    if (updatedCol > 0) {
      var allData = fieldMapSheet.getDataRange().getValues();
      for (var m = 1; m < allData.length; m++) {
        if (allData[m][updatedCol - 1] === true) {
          fieldMapSheet.getRange(m + 1, updatedCol).setValue(false);
          updatedColumnProcessed = true;
        }
      }
    }

    UtilityScriptLibrary.debugLog('processFieldMapForSemester', 'INFO', 'Field map processing completed', 
                                  'New fields: ' + newRows.length + ', Duplicates: ' + duplicates.length, '');

    return {
      newFieldsAdded: newRows.length,
      duplicatesFound: duplicates.length,
      updatedColumnProcessed: updatedColumnProcessed
    };
    
  }, 'Field map processed successfully', 'processFieldMapForSemester', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

function renameLatestFormSheet(semesterName) {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    UtilityScriptLibrary.debugLog('📝 renameLatestFormSheet - Starting form sheet rename for: ' + semesterName);
    
    var ss = UtilityScriptLibrary.getWorkbook('formResponses');
    var sheets = ss.getSheets();
    var formSheets = sheets.filter(function(s) { 
      return s.getName().toLowerCase().startsWith("form responses"); 
    });

    if (formSheets.length === 0) {
      UtilityScriptLibrary.debugLog('📝 renameLatestFormSheet - No form response sheets found');
      throw new Error("Registration Form not yet linked. Please link to Responses workbook and try again.");
    }

    var latestSheet = formSheets[formSheets.length - 1];
    var currentName = latestSheet.getName();
    
    UtilityScriptLibrary.debugLog('📝 renameLatestFormSheet - Found latest form sheet: ' + currentName);
    
    // Check if Teacher and Student ID columns already exist
    var currentHeaders = latestSheet.getRange(1, 1, 1, latestSheet.getLastColumn()).getValues()[0];
    var hasTeacher = false;
    var hasStudentId = false;
    
    for (var i = 0; i < currentHeaders.length; i++) {
      var normalizedHeader = UtilityScriptLibrary.normalizeHeader(currentHeaders[i]);
      if (normalizedHeader === 'teacher') {
        hasTeacher = true;
      }
      if (normalizedHeader === 'student id' || normalizedHeader === 'studentid') {
        hasStudentId = true;
      }
    }
    
    // Add missing columns
    if (!hasTeacher) {
      // Insert Teacher column at position F (column 6)
      latestSheet.insertColumnAfter(5); // Insert after column E
      latestSheet.getRange(1, 6).setValue('Teacher');
      
      UtilityScriptLibrary.debugLog('📝 renameLatestFormSheet - Added Teacher column at position F');
    }
    
    if (!hasStudentId) {
      // Add Student ID column at the end
      var lastCol = latestSheet.getLastColumn();
      latestSheet.getRange(1, lastCol + 1).setValue('Student ID');
      
      UtilityScriptLibrary.debugLog('📝 renameLatestFormSheet - Added Student ID column at end');
    }
    
    // Rename the sheet
    latestSheet.setName(semesterName);
    
    UtilityScriptLibrary.debugLog('📝 renameLatestFormSheet - Sheet renamed from "' + currentName + '" to "' + semesterName + '"');
    
    return {
      oldName: currentName,
      newName: semesterName,
      teacherAdded: !hasTeacher,
      studentIdAdded: !hasStudentId
    };
    
  }, 'Form response sheet renamed to ' + semesterName, 'renameLatestFormSheet', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

function verifyProgramsForSemester(semesterName, billingSS) {
  // FIX: Use direct sheet access since we're already in the billing spreadsheet
  var programSheet = billingSS.getSheetByName('Programs List');
  if (!programSheet) {
    throw new Error('Programs List sheet not found in billing spreadsheet.');
  }
  
  var programData = programSheet.getDataRange().getValues();
  var headers = programData[0];
  
  var nameCol = headers.indexOf('Program Name');
  var activeCol = headers.indexOf('Active');
  var typeCol = headers.indexOf('Type');
  var aliasCol = headers.indexOf('Alias For');

  // Get active programs
  var activePrograms = [];
  for (var i = 1; i < programData.length; i++) {
    if (programData[i][activeCol] === true) {
      activePrograms.push(programData[i][nameCol]);
    }
  }

  // Check for package issues
  var packageIssues = [];
  for (var j = 1; j < programData.length; j++) {
    var isActive = programData[j][activeCol] === true;
    var type = programData[j][typeCol];
    if (!isActive || type !== 'Package') continue;

    var programName = programData[j][nameCol];
    var aliasFor = programData[j][aliasCol] || '';
    
    if (!aliasFor) {
      packageIssues.push('Package "' + programName + '" has no "Alias For" value.');
      continue;
    }

    var aliasArray = aliasFor.split(',').map(function(alias) { return alias.trim(); }).filter(function(alias) { return alias; });
    var missingPrograms = aliasArray.filter(function(alias) { return activePrograms.indexOf(alias) === -1; });
    
    if (missingPrograms.length > 0) {
      packageIssues.push('Package "' + programName + '" references inactive programs: ' + missingPrograms.join(', '));
    }
  }

  var ui = SpreadsheetApp.getUi();
  var programStatus = 'Yes';
  
  // Handle package issues if any exist
  if (packageIssues.length > 0) {
    var packageAlert = ui.alert(
      'Program Issues Found:\n\n' + packageIssues.join('\n') + '\n\nContinue anyway?',
      ui.ButtonSet.YES_NO
    );
    if (packageAlert !== ui.Button.YES) {
      throw new Error('Semester setup cancelled - program issues not resolved.');
    }
    programStatus = 'Yes (with issues)';
  }
  
  // Always show active programs for confirmation
  var programList = activePrograms.join('\n');
  var programConfirm = ui.alert(
    'Active Programs for ' + semesterName + ':\n\n' + programList + '\n\nDo you confirm these are correct for this semester?',
    ui.ButtonSet.YES_NO
  );
  
  if (programConfirm !== ui.Button.YES) {
    throw new Error('Semester setup cancelled - programs not confirmed.');
  }

  return {
    status: programStatus,
    activePrograms: activePrograms,
    issues: packageIssues
  };
}

// ============================================================================
// SECTION 3: BEGIN NEW BILLING CYCLE
// ============================================================================

function addOveragesToBillingSheet(billingSheet, overageStudents, headerMap) {
  try {
    // Find or create "Lesson Overages" column
    var overageAmountCol = headerMap[UtilityScriptLibrary.normalizeHeader('Lesson Overages')];
    
    if (!overageAmountCol) {
      // Add column if it doesn't exist
      var lastCol = billingSheet.getLastColumn();
      billingSheet.getRange(1, lastCol + 1).setValue('Lesson Overages');
      overageAmountCol = lastCol + 1;
    }
    
    // FIXED: Use correct column name "Current Cumulative Hours Billed"
    var hoursBilledCol = headerMap[UtilityScriptLibrary.normalizeHeader('Current Cumulative Hours Billed')];
    
    if (!hoursBilledCol) {
      UtilityScriptLibrary.debugLog('addOveragesToBillingSheet', 'ERROR', 
                    'Current Cumulative Hours Billed column not found', '', '');
      throw new Error('Current Cumulative Hours Billed column not found in billing sheet');
    }
    
    for (var i = 0; i < overageStudents.length; i++) {
      var student = overageStudents[i];
      var rate = getCurrentSemesterRateForLength(student.lessonLength);
      var overageAmount = student.overageHours * rate;
      
      // Add overage amount to billing sheet
      billingSheet.getRange(student.rowIndex, overageAmountCol).setValue(overageAmount);
      
      // Update cumulative hours billed
      var newBilled = student.billedHours + student.overageHours;
      billingSheet.getRange(student.rowIndex, hoursBilledCol).setValue(newBilled);
      
      UtilityScriptLibrary.debugLog('💰 Added overage to billing sheet - Student ' + student.studentId + ': $' + overageAmount);
    }
    
    UtilityScriptLibrary.debugLog('✅ Added overages to billing sheet for ' + overageStudents.length + ' students');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error adding overages to billing sheet: ' + error.message);
    throw error;
  }
}

function appendToBillingMetadata(billingCycleName, customToday, semesterName) {
  UtilityScriptLibrary.debugLog("appendToBillingMetadata", "INFO", "Adding billing metadata", 
                "Cycle: " + billingCycleName + ", Semester: " + semesterName, "");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = ss.getSheetByName('Billing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Billing Metadata sheet not found.');
    }
    
    var data = metadataSheet.getDataRange().getValues();
    var lastRow = data.length;
    
    var prevMonth = '';
    if (lastRow > 1) {
      prevMonth = data[lastRow - 1][0];
      
      // Update previous row's payment ending date to customToday - 1
      var previousEndDate = new Date(customToday);
      previousEndDate.setDate(previousEndDate.getDate() - 1);
      previousEndDate.setHours(0, 0, 0, 0);  // Strip time
      metadataSheet.getRange(lastRow, 4).setValue(previousEndDate);
    }
    
    var invoicingDate = new Date(customToday);
    invoicingDate.setHours(0, 0, 0, 0);  // Strip time
    
    var paymentStartingDate = new Date(customToday);
    paymentStartingDate.setHours(0, 0, 0, 0);  // Strip time
    
    var rateChart = getRateColumnFromMetadata(semesterName);
    
    var rowData = [
      billingCycleName,           // Billing Month
      invoicingDate,              // Invoicing Date
      paymentStartingDate,        // Payment Starting Date
      '=TODAY()',                 // FIXED: Payment Ending Date as formula
      prevMonth,                  // Previous Month
      rateChart,                  // Rates
      semesterName                // Semester Name
    ];
    
    metadataSheet.appendRow(rowData);
    
    // IMPORTANT: After appending, we need to set the formula in the Payment Ending Date column
    // because appendRow() treats '=TODAY()' as a string, not a formula
    var newRow = metadataSheet.getLastRow();
    metadataSheet.getRange(newRow, 4).setFormula('=TODAY()');
    
    UtilityScriptLibrary.debugLog("appendToBillingMetadata", "SUCCESS", "Billing metadata appended", 
                  "Cycle: " + billingCycleName + ", Row: " + newRow, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("appendToBillingMetadata", "ERROR", "Failed to append billing metadata", 
                  "", error.message);
    throw error;
  }
}

function applyLessonEquivalentCredits(studentId, newRegistrationData) {
  try {
    var carryoverBalance = getPreviousSemesterBalance(studentId);
    
    if (!carryoverBalance || carryoverBalance.lessonBalance <= 0) {
      // No credits to apply
      return {
        adjustment: null,
        explanation: null
      };
    }
    
    var newLessonLength = parseInt(newRegistrationData.lessonLength);
    var carryoverLessonLength = carryoverBalance.lessonLength;
    var newRegistrationQuantity = newRegistrationData.quantity;
    
    // Convert carryover lessons to new lesson length equivalents
    var convertedCarryover = (carryoverBalance.lessonBalance * carryoverLessonLength) / newLessonLength;
    
    // Calculate billing adjustment
    var totalLessons = newRegistrationQuantity + convertedCarryover;
    var lessonsToAdd = Math.max(0, newRegistrationQuantity - convertedCarryover);
    var billingReduction = newRegistrationQuantity - lessonsToAdd;
    
    var adjustment = {
      totalLessons: totalLessons,
      lessonsToAdd: lessonsToAdd,
      carryoverApplied: convertedCarryover,
      billingReduction: billingReduction,
      carryoverSemester: carryoverBalance.semesterName,
      originalCarryover: carryoverBalance.lessonBalance,
      originalLessonLength: carryoverLessonLength
    };
    
    var explanation = newRegistrationQuantity + ' new lessons + ' + convertedCarryover + ' carryover = ' + totalLessons + ' total lessons (billing for ' + lessonsToAdd + ' lessons)';
    
    UtilityScriptLibrary.debugLog('💰 Credit applied - Student ' + studentId + ': ' + explanation);
    
    return {
      adjustment: adjustment,
      explanation: explanation
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error applying lesson credits for student ' + studentId + ': ' + error.message);
    return {
      adjustment: null,
      explanation: null
    };
  }
}

function buildBillingContext(customToday, semesterName, billingCycleName) {
  UtilityScriptLibrary.debugLog("buildBillingContext", "INFO", "Building billing context for new cycle", 
          "Date: " + customToday.toLocaleDateString() + ", Semester: " + semesterName, "");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formResponsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
  var formSheet;
  
  // STANDARDIZED: Get form sheet by exact semester name, fail if not found
  try {
    formSheet = formResponsesSS.getSheetByName(semesterName);
    if (!formSheet) {
      throw new Error("Semester '" + semesterName + "' not found in Responses workbook. Please set up the semester first.");
    }
  } catch (e) {
    UtilityScriptLibrary.debugLog("buildBillingContext", "ERROR", "Could not access form sheet", "", e.message);
    throw new Error("Could not access form responses sheet: " + e.message);
  }
  
  var billingSheet = ss.getSheetByName(billingCycleName);
  if (!billingSheet) {
    throw new Error("Billing sheet not found: " + billingCycleName);
  }
  
  var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
  
  // Build program map
  var programSheet = ss.getSheetByName("Programs List");
  var programData = programSheet.getDataRange().getValues();
  var headers = programData[0];
  var nameCol = headers.indexOf("Program Name");
  var typeCol = headers.indexOf("Type");
  var prefixCol = headers.indexOf("Billing Sheet Prefix");
  var rateKeyCol = headers.indexOf("Rate Key");
  var quantityTypeCol = headers.indexOf("Quantity Type");
  
  var programMap = {};
  for (var i = 1; i < programData.length; i++) {
    var row = programData[i];
    var active = row[headers.indexOf("Active")];
    if (active === true && row[typeCol] !== "Package") {
      var programName = row[nameCol];
      programMap[programName] = {
        prefix: row[prefixCol],
        rateKey: row[rateKeyCol],
        quantityType: row[quantityTypeCol]
      };
    }
  }
  
  // Load field mapping
  var fieldMapSheet = formResponsesSS.getSheetByName("FieldMap");
  if (!fieldMapSheet) {
    throw new Error("FieldMap sheet not found in Form Responses workbook");
  }
  var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
  
  // Create column index helper with field mapping support
  var formHeaderMap = UtilityScriptLibrary.getHeaderMap(formSheet);
  
  var getColIndex = function(internalFieldName) {
  var formHeader = null;
    for (var key in fieldMap) {
      if (fieldMap[key] === internalFieldName) {
        formHeader = key;
        break;
      }
    }
    
    if (!formHeader) {
      return -1;
    }
    
    // formHeader is already normalized (it's a key from fieldMap)
    var colIndex = formHeaderMap[formHeader];
    return colIndex ? colIndex - 1 : -1;
  };
  
  var context = {
    ss: ss,
    semesterName: semesterName,
    billingCycleName: billingCycleName,
    today: UtilityScriptLibrary.formatDateFlexible(customToday, "yyyyMMdd"),
    billingSheet: billingSheet,
    formSheet: formSheet,
    headerMap: headerMap,
    programMap: programMap,
    programData: programData,
    nameCol: nameCol,
    typeCol: typeCol,
    aliasCol: headers.indexOf("Alias For"),
    getColIndex: getColIndex,
    fieldMap: fieldMap,
    prevHeaderMap: {} // Will be set later if there's previous data
  };
  
  UtilityScriptLibrary.debugLog("buildBillingContext", "INFO", "Billing context built successfully", 
                "Programs: " + Object.keys(programMap).length, "");
  
  return context;
}

function buildDynamicAmounts(billingData) {
  try {
    UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'DEBUG', 'Starting function', 
                  'billingData keys: ' + Object.keys(billingData).join(', '), '');
    
    var amounts = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get Programs List to understand packages
    var programSheet = ss.getSheetByName('Programs List');
    if (!programSheet) {
      UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'ERROR', 'Programs List sheet not found', '', '');
      throw new Error('Programs List sheet not found');
    }
    
    var programData = programSheet.getDataRange().getValues();
    var headers = programData[0];
    var nameCol = headers.indexOf('Program Name');
    var typeCol = headers.indexOf('Type');
    var aliasCol = headers.indexOf('Alias For');
    var activeCol = headers.indexOf('Active');
    
    UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'DEBUG', 'Column indices', 
                  'Name: ' + nameCol + ', Type: ' + typeCol + ', Alias: ' + aliasCol + ', Active: ' + activeCol, '');
    
    if (nameCol === -1 || typeCol === -1 || activeCol === -1) {
      throw new Error('Required columns not found in Programs List');
    }
    
    // Build package lookup
    var packageAliases = {};
    
    for (var i = 1; i < programData.length; i++) {
      var row = programData[i];
      var programName = row[nameCol];
      var isActive = row[activeCol] === true;
      var type = row[typeCol];
      var aliasFor = row[aliasCol] || '';
      
      if (isActive && type === 'Package' && aliasFor) {
        var components = aliasFor.split(',')
          .map(function(comp) { return comp.trim().toLowerCase(); })
          .filter(function(comp) { return comp; });
        packageAliases[programName.toLowerCase()] = components;
      }
    }
    
    // Track covered programs
    var coveredPrograms = [];
    
    // 1. Previous balance amount
    if (billingData.pastBalance && billingData.pastBalance > 0) {
      amounts.push(UtilityScriptLibrary.formatCurrency(billingData.pastBalance));
    }
    
    // 1b. Credit from previous billing (negative past balance)
    if (billingData.pastBalance && billingData.pastBalance < 0) {
      amounts.push(UtilityScriptLibrary.formatCurrency(billingData.pastBalance));
    }
    
    // 2. PASS 1: Add charge amounts (quantity > 0)
    if (billingData.programTotals && billingData.programTotals.programs) {
      UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'DEBUG', 'Processing programs for charge amounts', 
                    'Program count: ' + billingData.programTotals.programs.length, '');
      
      for (var i = 0; i < billingData.programTotals.programs.length; i++) {
        var program = billingData.programTotals.programs[i];
        var programNameLower = program.name.toLowerCase();
        
        // Skip if covered by a package
        if (coveredPrograms.indexOf(programNameLower) !== -1) {
          continue;
        }
        
        // Check if this is a package that covers other programs
        if (packageAliases[programNameLower]) {
          coveredPrograms = coveredPrograms.concat(packageAliases[programNameLower]);
        }
        
        // Add charge amount if quantity > 0
        var quantity = parseFloat(program.quantity) || 0;
        if (quantity > 0) {
          var price = parseFloat(program.price) || 0;
          var chargeAmount = 0;
          
          if (programNameLower === 'lesson') {
            // For lessons, calculate based on hours
            var lessonLength = parseInt(billingData.lessonLength) || 30;
            var hours = quantity * lessonLength / 60;
            chargeAmount = hours * price;
          } else {
            // For other programs
            chargeAmount = quantity * price;
          }
          
          amounts.push(UtilityScriptLibrary.formatCurrency(chargeAmount));
        }
      }
    }
    
    // 3. PASS 2: Add credit amounts (credit > 0)
    if (billingData.programTotals && billingData.programTotals.programs) {
      UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'DEBUG', 'Processing programs for credit amounts', 
                    'Program count: ' + billingData.programTotals.programs.length, '');
      
      for (var i = 0; i < billingData.programTotals.programs.length; i++) {
        var program = billingData.programTotals.programs[i];
        var programNameLower = program.name.toLowerCase();
        
        // Skip if covered by a package
        if (coveredPrograms.indexOf(programNameLower) !== -1) {
          continue;
        }
        
        // Add credit amount if credit > 0
        var credit = parseFloat(program.credit) || 0;
        if (credit > 0) {
          var price = parseFloat(program.price) || 0;
          var creditAmount = 0;
          
          if (programNameLower === 'lesson') {
            // For lessons, calculate based on hours
            var lessonLength = parseInt(billingData.lessonLength) || 30;
            var hours = credit * lessonLength / 60;
            creditAmount = -(hours * price); // Make negative
          } else {
            // For other programs
            creditAmount = -(credit * price); // Make negative
          }
          
          amounts.push(UtilityScriptLibrary.formatCurrency(creditAmount));
        }
      }
    }
    
    // 4. Late fee amount
    if (billingData.lateFee && billingData.lateFee > 0) {
      amounts.push(UtilityScriptLibrary.formatCurrency(billingData.lateFee));
    }
    
    UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'INFO', 'Amounts built successfully', 
                  'Amount count: ' + amounts.length, '');
    
    return amounts.join('\n');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('buildDynamicAmounts', 'ERROR', 'Failed to build amounts', 
                  '', error.message + ' | ' + error.stack);
    return 'Error: ' + error.message;
  }
}

function buildDynamicProgramColumnsWithCredits(newRow, row, enrolledPrograms, context, quantityCols, currencyCols, rowNum, creditAdjustment) {
  var h = context.headerMap;
  var get = context.getColIndex;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ratesSheet = ss.getSheetByName("Rates");
  var rateHeaders = ratesSheet.getRange(1, 1, 1, ratesSheet.getLastColumn()).getValues()[0];
  var rateYearCol = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
  var rateMap = UtilityScriptLibrary.buildRateMapFromSheet(ratesSheet, rateHeaders, rateYearCol);
  
  UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Starting program processing", 
                "Row: " + rowNum + ", Enrolled programs count: " + enrolledPrograms.length, "");
  
  for (var i = 0; i < enrolledPrograms.length; i++) {
    var program = enrolledPrograms[i];
    var programName = program.name;
    var prefix = program.prefix;
    var isEnrolled = program.enrolled;
    var qtyType = program.quantityType;
    var price = parseFloat(rateMap[programName + " Price"]) || 0;
    
    UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Processing program", 
                  "Program: " + programName + ", Prefix: " + prefix + ", Enrolled: " + isEnrolled + ", QtyType: " + qtyType, "");
    
    var qtyCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Quantity")];
    var creditCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Credit")];
    var priceCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Price")];
    var hoursCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Hours")];

    UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Column mapping", 
                  "Program: " + programName + ", QtyCol: " + qtyCol + ", CreditCol: " + creditCol, "");

    UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'Prefix check',
                              'Program: ' + programName + ', Prefix: "' + prefix + '", prefix.toLowerCase(): "' + prefix.toLowerCase() + '"', '');

    // Extract actual quantities from form data
    var quantity = 0;
    
    if (isEnrolled && programName.toLowerCase() === "lessons") {
      UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'ENTERED LESSON BLOCK',
                                    'Program: ' + programName, '');
      // Extract lesson quantities from package selections
      var qty60Col = get("Qty60");
      var qty45Col = get("Qty45");
      var qty30Col = get("Qty30");
      
      var qty60 = 0;
      var qty45 = 0;
      var qty30 = 0;
      
      if (qty60Col !== -1 && row[qty60Col]) {
        qty60 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty60Col]) || 0;
      }
      if (qty45Col !== -1 && row[qty45Col]) {
        qty45 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty45Col]) || 0;
      }
      if (qty30Col !== -1 && row[qty30Col]) {
        qty30 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty30Col]) || 0;
      }
      
      quantity = qty60 + qty45 + qty30;
    } else if (isEnrolled && qtyType === "Multiple") {
      quantity = row[get(qtyType)] || 1;
      UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Multiple quantity type", 
                    "Program: " + programName + ", Quantity: " + quantity, "");
      
    } else if (isEnrolled) {
      UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Enrolled non-lesson program", 
                    "Program: " + programName + ", Prefix: " + prefix, "");
      
      // FIXED: For non-lesson programs, look up quantity from rates
      if (prefix.toLowerCase() !== "lessons") {
        UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Looking up rates for non-lesson program", 
                      "Program: " + programName + ", Prefix: " + prefix, "");
        
        var rateSheet = ss.getSheetByName('Rates');
        if (rateSheet) {
          var rateData = rateSheet.getDataRange().getValues();
          var rateHeaders = rateData[0];
          var bestColIndex = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
          
          UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Rate sheet access", 
                        "Headers: " + rateHeaders.join(', ') + ", BestColIndex: " + bestColIndex, "");
          
          if (bestColIndex !== -1) {
            // Look for "[Prefix] Number" in rates - e.g. "Group Number", "Concerts Number", etc.
            var rateKey = prefix + " Number";
            UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Looking for rate key", 
                          "RateKey: '" + rateKey + "'", "");
            
            for (var j = 1; j < rateData.length; j++) {
              var rateTitle = rateData[j][0];
              UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Checking rate row", 
                            "Row " + j + ": Title='" + rateTitle + "', Value=" + rateData[j][bestColIndex], "");
              
              if (rateTitle === rateKey) {
                quantity = rateData[j][bestColIndex] || 0; // Use 0 as fallback to indicate problem
                UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "FOUND MATCHING RATE", 
                              "RateKey: '" + rateKey + "', Quantity: " + quantity, "");
                break;
              }
            }
            
            if (quantity === 0) {
              UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "WARNING", "Rate not found", 
                            "RateKey: '" + rateKey + "' not found in rates", "");
            }
          } else {
            UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "WARNING", "No valid rate column found", "", "");
          }
        } else {
          UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "ERROR", "Rates sheet not found", "", "");
        }
      } else {
        UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Lesson program in else clause - should not happen", 
                      "Program: " + programName, "");
      }
      
      // Fallback for any programs that don't have rates entries: keep at 0
      // This way you'll know immediately if a rate is missing
    } else {
      UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Program not enrolled or other condition", 
                    "Program: " + programName + ", Enrolled: " + isEnrolled, "");
    }

    UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Final quantity setting", 
                  "Program: " + programName + ", Final Quantity: " + quantity + ", Column: " + qtyCol, "");

    if (qtyCol) {
      newRow[qtyCol - 1] = quantity;
      quantityCols.push(qtyCol);
    }
    if (creditCol) {
      newRow[creditCol - 1] = 0;
      quantityCols.push(creditCol);
    }
    if (priceCol) {
      newRow[priceCol - 1] = price;
      currencyCols.push(priceCol);
    }
    
    // Hours column for lesson programs
    if (hoursCol && prefix.toLowerCase() === "lesson") {
      var lessonLength = getLessonLengthFromRow(row, get);
      var lengthMinutes = parseInt(lessonLength) || 30;
      var hours = quantity * (lengthMinutes / 60);
      newRow[hoursCol - 1] = hours;
      quantityCols.push(hoursCol);
    }
  }
  
  // Generate formulas
  generateProgramFormulas(newRow, context, rowNum, quantityCols, currencyCols);
  
  return { quantityCols: quantityCols, currencyCols: currencyCols };
}

function buildDynamicProgramColumns(newRow, row, enrolledPrograms, context, quantityCols, currencyCols, rowNum) {
    UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'FUNCTION ENTERED',
                                'enrolledPrograms: ' + JSON.stringify(enrolledPrograms) + ', type: ' + (Array.isArray(enrolledPrograms) ? 'array' : typeof enrolledPrograms), '');
  var get = context.getColIndex;
  var h = context.headerMap;
  var ratesSheet = context.ss.getSheetByName("Rates");
  var rateHeaders = ratesSheet.getRange(1, 1, 1, ratesSheet.getLastColumn()).getValues()[0];
  var rateYearCol = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
  var rateMap = UtilityScriptLibrary.buildRateMapFromSheet(ratesSheet, rateHeaders, rateYearCol);

  // FIXED: Convert Set to Array for ES5 compatibility
  var enrolledArray = [];
  if (enrolledPrograms && typeof enrolledPrograms === 'object') {
    if (enrolledPrograms.forEach) {
      // It's a Set - convert to array
      enrolledPrograms.forEach(function(program) {
        enrolledArray.push(program);
      });
    } else if (enrolledPrograms.length !== undefined) {
      // It's already an array
      enrolledArray = enrolledPrograms;
    }
  }

  for (var programName in context.programMap) {
    var program = context.programMap[programName];
    var prefix = program.prefix;
    var rateKey = program.rateKey;
    var qtyType = program.quantityType;
    
    // FIXED: ES5-compatible enrollment check
    var isEnrolled = false;
    for (var i = 0; i < enrolledArray.length; i++) {
      if (enrolledArray[i] === programName.toLowerCase()) {
        isEnrolled = true;
        break;
      }
    }
    
    UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'Checking program enrollment',
                                  'Program: ' + programName + ', isEnrolled: ' + isEnrolled + ', enrolledArray: ' + JSON.stringify(enrolledArray), '');
    
    var price = rateMap[rateKey] || 0;
    var qtyCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Quantity")];
    var creditCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Credit")];
    var priceCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Price")];
    var hoursCol = h[UtilityScriptLibrary.normalizeHeader(prefix + " Hours")];

    UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'Prefix check',
                              'Program: ' + programName + ', Prefix: "' + prefix + '", prefix.toLowerCase(): "' + prefix.toLowerCase() + '"', '');

    // Extract actual quantities from form data
    var quantity = 0;
    if (isEnrolled && programName.toLowerCase() === "lessons") {
      UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'ENTERED LESSON BLOCK',
                                    'Program: ' + programName, '');
      
      // Add this diagnostic block
      var qty60Col = get("Qty60");
      var qty45Col = get("Qty45");
      var qty30Col = get("Qty30");
      
      UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'Column indices retrieved',
                                    'qty60Col: ' + qty60Col + ', qty45Col: ' + qty45Col + ', qty30Col: ' + qty30Col, '');
      
      var qty60 = 0;
      var qty45 = 0;
      var qty30 = 0;
      
      if (qty60Col !== -1 && row[qty60Col]) {
        qty60 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty60Col]) || 0;
      }
      if (qty45Col !== -1 && row[qty45Col]) {
        qty45 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty45Col]) || 0;
      }
      if (qty30Col !== -1 && row[qty30Col]) {
        qty30 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty30Col]) || 0;
      }
      
      UtilityScriptLibrary.debugLog('buildDynamicProgramColumns', 'DEBUG', 'Quantities extracted',
                                    'qty60: ' + qty60 + ', qty45: ' + qty45 + ', qty30: ' + qty30, '');
      
      quantity = qty60 + qty45 + qty30;
      UtilityScriptLibrary.debugLog("buildDynamicProgramColumnsWithCredits", "DEBUG", "Lesson quantity calculated", 
                    "Qty60: " + qty60 + ", Qty45: " + qty45 + ", Qty30: " + qty30 + ", Total: " + quantity, "");
      
    } else if (isEnrolled && qtyType === "Multiple") {
      quantity = row[get(qtyType)] || 1;
    } else if (isEnrolled) {
      // FIXED: For non-lesson programs, look up quantity from rates
      if (prefix.toLowerCase() !== "lessons") {
        var rateSheet = context.ss.getSheetByName('Rates');
        if (rateSheet) {
          var rateData = rateSheet.getDataRange().getValues();
          var rateHeaders = rateData[0];
          var bestColIndex = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
          
          if (bestColIndex !== -1) {
            // Look for "[Prefix] Number" in rates - e.g. "Group Number", "Concerts Number", etc.
            var rateKey = prefix + " Number";
            for (var j = 1; j < rateData.length; j++) {
              if (rateData[j][0] === rateKey) {
                quantity = rateData[j][bestColIndex] || 0; // Use 0 as fallback to indicate problem
                break;
              }
            }
          }
        }
      }
      
      // Fallback for any programs that don't have rates entries: keep at 0
      // This way you'll know immediately if a rate is missing
    }

    if (qtyCol) {
      newRow[qtyCol - 1] = quantity;
      quantityCols.push(qtyCol);
    }
    if (creditCol) {
      newRow[creditCol - 1] = 0;
      quantityCols.push(creditCol);
    }
    if (priceCol) {
      newRow[priceCol - 1] = price;
      currencyCols.push(priceCol);
    }
    
    // Hours column for lesson programs
    if (hoursCol && prefix.toLowerCase() === "lesson") {
      var lessonLength = getLessonLengthFromRow(row, get);
      var lengthMinutes = parseInt(lessonLength) || 30;
      var hours = quantity * (lengthMinutes / 60);
      newRow[hoursCol - 1] = hours;
      quantityCols.push(hoursCol);
    }
  }
  
  // Generate formulas
  generateProgramFormulas(newRow, context, rowNum, quantityCols, currencyCols);
  
  return { quantityCols: quantityCols, currencyCols: currencyCols };
}

function buildDynamicLineItems(billingData) {
  try {
    UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'DEBUG', 'Starting function', 
                  'billingData keys: ' + Object.keys(billingData).join(', '), '');
    
    var lineItems = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get Programs List to understand packages
    var programSheet = ss.getSheetByName('Programs List');
    if (!programSheet) {
      UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'ERROR', 'Programs List sheet not found', '', '');
      throw new Error('Programs List sheet not found');
    }
    
    var programData = programSheet.getDataRange().getValues();
    var headers = programData[0];
    var nameCol = headers.indexOf('Program Name');
    var typeCol = headers.indexOf('Type');
    var aliasCol = headers.indexOf('Alias For');
    var activeCol = headers.indexOf('Active');
    
    UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'DEBUG', 'Column indices', 
                  'Name: ' + nameCol + ', Type: ' + typeCol + ', Alias: ' + aliasCol + ', Active: ' + activeCol, '');
    
    if (nameCol === -1 || typeCol === -1 || activeCol === -1) {
      throw new Error('Required columns not found in Programs List');
    }
    
    // Build lookup maps
    var packageAliases = {}; // package name -> array of component programs
    var activePrograms = [];
    
    for (var i = 1; i < programData.length; i++) {
      var row = programData[i];
      var programName = row[nameCol];
      var isActive = row[activeCol] === true;
      var type = row[typeCol];
      var aliasFor = row[aliasCol] || '';
      
      if (isActive) {
        activePrograms.push(programName.toLowerCase());
        
        if (type === 'Package' && aliasFor) {
          var components = aliasFor.split(',')
            .map(function(comp) { return comp.trim().toLowerCase(); })
            .filter(function(comp) { return comp; });
          packageAliases[programName.toLowerCase()] = components;
        }
      }
    }
    
    // Track which individual programs are covered by packages
    var coveredPrograms = [];
    
    // 1. Previous balance (if >0)
    if (billingData.pastBalance && billingData.pastBalance > 0) {
      lineItems.push('Previous Balance');
    }
    
    // 1b. Credit from previous billing (if <0)
    if (billingData.pastBalance && billingData.pastBalance < 0) {
      lineItems.push('Credit from Previous Billing');
    }
    
    // 2. PASS 1: Add charge line items (quantity > 0)
    if (billingData.programTotals && billingData.programTotals.programs) {
      UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'DEBUG', 'Processing programs for charges', 
                    'Program count: ' + billingData.programTotals.programs.length, '');
      
      for (var i = 0; i < billingData.programTotals.programs.length; i++) {
        var program = billingData.programTotals.programs[i];
        var programNameLower = program.name.toLowerCase();
        
        var quantity = parseFloat(program.quantity) || 0;
        
        if (quantity > 0) {
          // Check if this is a package
          if (packageAliases[programNameLower]) {
            // This is a package - build description with components
            var packageComponents = [];
            var components = packageAliases[programNameLower];
            
            for (var k = 0; k < components.length; k++) {
              var componentName = components[k];
              
              if (componentName === 'lesson') {
                // Special handling for lessons - include quantity and length
                packageComponents.push(quantity + ' ' + billingData.lessonLength + '-minute lessons');
              } else {
                // For other components, find their quantity if available
                var componentQuantity = 1; // default
                for (var l = 0; l < billingData.programTotals.programs.length; l++) {
                  var otherProgram = billingData.programTotals.programs[l];
                  if (otherProgram.name.toLowerCase() === componentName) {
                    componentQuantity = otherProgram.quantity;
                    break;
                  }
                }
                packageComponents.push(componentQuantity + ' ' + componentName + ' session' + (componentQuantity === 1 ? '' : 's'));
              }
            }
            
            var packageDescription = program.name + ' (' + packageComponents.join(', ') + ')';
            lineItems.push(packageDescription);
            
            // Mark components as covered
            for (var m = 0; m < components.length; m++) {
              coveredPrograms.push(components[m]);
            }
          } else {
            // Check if this individual program is not covered by a package
            if (coveredPrograms.indexOf(programNameLower) === -1) {
              if (programNameLower === 'lesson') {
                // Special formatting for individual lessons
                lineItems.push(quantity + ' ' + billingData.lessonLength + '-minute lessons');
              } else {
                // Other individual programs
                lineItems.push(quantity + ' ' + program.name.toLowerCase() + ' session' + (quantity === 1 ? '' : 's'));
              }
            }
          }
        }
      }
    }
    
    // 3. PASS 2: Add credit line items (credit > 0)
    if (billingData.programTotals && billingData.programTotals.programs) {
      UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'DEBUG', 'Processing programs for credits', 
                    'Program count: ' + billingData.programTotals.programs.length, '');
      
      for (var i = 0; i < billingData.programTotals.programs.length; i++) {
        var program = billingData.programTotals.programs[i];
        var programNameLower = program.name.toLowerCase();
        
        // Skip if covered by a package
        if (coveredPrograms.indexOf(programNameLower) !== -1) {
          continue;
        }
        
        var credit = parseFloat(program.credit) || 0;
        
        if (credit > 0) {
          if (programNameLower === 'lesson') {
            // Special formatting for lesson credits
            lineItems.push(credit + ' ' + billingData.lessonLength + '-minute lessons Credit');
          } else {
            // Other program credits
            lineItems.push(credit + ' ' + program.name.toLowerCase() + ' session' + (credit === 1 ? '' : 's') + ' Credit');
          }
        }
      }
    }
    
    // 4. Late fee (if >0)
    if (billingData.lateFee && billingData.lateFee > 0) {
      lineItems.push('Late Fee');
    }
    
    UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'INFO', 'Line items built successfully', 
                  'Item count: ' + lineItems.length, '');
    
    return lineItems.join('\n');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('buildDynamicLineItems', 'ERROR', 'Failed to build line items', 
                  '', error.message + ' | ' + error.stack);
    return 'Error: ' + error.message;
  }
}

function buildInvoiceTotalFormula(headerMap, rowNum) {
  var colLetter = UtilityScriptLibrary.columnToLetter;
  var norm = UtilityScriptLibrary.normalizeHeader;
  var parts = [];

  for (var header in headerMap) {
    var normalized = norm(header);
    var colIndex = headerMap[header];
    var cellRef = colLetter(colIndex) + rowNum;
    
    // FIXED: Use normalized comparison (no spaces)
    if (normalized === "currentinvoicetotal") {
      continue;
    }
    
    // Add program total columns and static components
    if (normalized.endsWith("total")) {
      parts.push(cellRef);
    } else if (normalized === "pastbalance" || normalized === "latefee") {
      parts.push(cellRef);
    }
  }

  if (parts.length === 0) {
    return "=0";
  }

  return "=" + parts.join(" + ");
}

function copyStaticFieldsToBillingRow(newRow, sourceRow, context, getFn) {
  var fields = [
    "Student Last Name", "Student First Name", "Student ID", "Instrument",
    "Teacher", "Enrollment", "Salutation", "Parent First Name",
    "Parent Last Name"
  ];

  for (var i = 0; i < fields.length; i++) {
    var field = fields[i];
    var normField = UtilityScriptLibrary.normalizeHeader(field);
    var colIndex = context.headerMap[normField];
    if (!colIndex) continue;

    var value;
    if (getFn) {
      var formColIndex = getFn(field);
      value = formColIndex === -1 ? '' : sourceRow[formColIndex];
    } else {
      var prevColIndex = context.prevHeaderMap[normField];
      value = prevColIndex ? sourceRow[prevColIndex - 1] : '';
    }
    newRow[colIndex - 1] = value || '';
  }

  // Handle Parent Address
  var addressCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Parent Address")];
  if (addressCol) {
    if (getFn) {
      var street = sourceRow[getFn("Address Street")];
      var cityZip = sourceRow[getFn("CityZip")];
      var cityZipParsed = UtilityScriptLibrary.parseCityZipMessy(cityZip);
      var city = cityZipParsed.city;
      var zip = cityZipParsed.zip;
      newRow[addressCol - 1] = street + '\n' + city + ', NY ' + zip;
    } else {
      var prevAddressCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Parent Address")];
      newRow[addressCol - 1] = prevAddressCol ? sourceRow[prevAddressCol - 1] : '';
    }
  }

  // Handle Lesson Length from previous billing cycle
  var lessonLengthCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Lesson Length")];
  if (lessonLengthCol && !getFn) {
    var prevLessonLengthCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Lesson Length")];
    newRow[lessonLengthCol - 1] = prevLessonLengthCol ? sourceRow[prevLessonLengthCol - 1] : 0;
  }

  // Handle Lesson Price from previous billing cycle
  var lessonPriceCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Lesson Price")];
  if (lessonPriceCol && !getFn) {
    var prevLessonPriceCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Lesson Price")];
    newRow[lessonPriceCol - 1] = prevLessonPriceCol ? sourceRow[prevLessonPriceCol - 1] : 0;
  }

  // NEW: Handle Agreement Form and Media Release from previous billing cycle
  // These carry over within semester automatically, and cross-semester if user confirms
  if (!getFn) {
    var agreementFormCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Agreement Form")];
    var mediaReleaseCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Media Release")];
    
    if (agreementFormCol) {
      var prevAgreementCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Agreement Form")];
      // Check if we should carry over forms (context.carryOverForms is set during billing cycle creation)
      if (prevAgreementCol && context.carryOverForms !== false) {
        newRow[agreementFormCol - 1] = sourceRow[prevAgreementCol - 1] === true;
      } else {
        newRow[agreementFormCol - 1] = false; // Default to false if not carrying over
      }
    }
    
    if (mediaReleaseCol) {
      var prevMediaCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Media Release")];
      if (prevMediaCol && context.carryOverForms !== false) {
        newRow[mediaReleaseCol - 1] = sourceRow[prevMediaCol - 1] === true;
      } else {
        newRow[mediaReleaseCol - 1] = false; // Default to false if not carrying over
      }
    }
  } else {
    // When building from form (new student), default to false
    var agreementFormCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Agreement Form")];
    var mediaReleaseCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Media Release")];
    
    if (agreementFormCol) {
      newRow[agreementFormCol - 1] = false;
    }
    if (mediaReleaseCol) {
      newRow[mediaReleaseCol - 1] = false;
    }
  }

  // Handle Past data fields (for continuing semester billing only)
  if (!getFn) {
    var pastFieldMappings = [
      { pastField: "Past Balance", currentField: "Current Balance" },
      { pastField: "Past Invoice Number", currentField: "Invoice Number" },
      { pastField: "Past Cumulative Hours Taught", currentField: "Current Cumulative Hours Taught" },
      { pastField: "Past Cumulative Hours Billed", currentField: "Current Cumulative Hours Billed" }
    ];
    
    for (var j = 0; j < pastFieldMappings.length; j++) {
      var mapping = pastFieldMappings[j];
      var pastCol = context.headerMap[UtilityScriptLibrary.normalizeHeader(mapping.pastField)];
      if (pastCol) {
        var prevCurrentCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader(mapping.currentField)];
        newRow[pastCol - 1] = prevCurrentCol ? sourceRow[prevCurrentCol - 1] : '';
        
        // Track hours columns for special formatting (CURRENCY FIX)
        if (mapping.pastField.indexOf("Hours") !== -1) {
          if (!context.hoursColumnsToFormat) {
            context.hoursColumnsToFormat = [];
          }
          context.hoursColumnsToFormat.push(pastCol);
        }
      }
    }
  }
}

function populateLetterType(newRow, context, sourceType, prevRow) {
  try {
    var letterTypeCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Letter Type")];
    if (!letterTypeCol) {
      UtilityScriptLibrary.debugLog("populateLetterType", "WARNING", "Letter Type column not found", "", "");
      return;
    }
    
    // Check if Letter Type is already set
    if (newRow[letterTypeCol - 1] && newRow[letterTypeCol - 1] !== "") {
      UtilityScriptLibrary.debugLog("populateLetterType", "DEBUG", "Letter Type already set, not overwriting", 
                    "Value: " + newRow[letterTypeCol - 1], "");
      return;
    }
    
    var letterType = ""; // Default to blank
    
    if (sourceType === "form") {
      // NEW LOGIC: Check if student is new or returning
      var studentIdCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Student ID")];
      var studentId = newRow[studentIdCol - 1];
      
      if (studentId) {
        var isNewStudent = determineIfNewStudent(studentId, context.semesterName);
        letterType = isNewStudent ? "welcome" : "returning";
        UtilityScriptLibrary.debugLog("populateLetterType", "DEBUG", "Set letter type from form", 
                      "Type: " + letterType + ", Student ID: " + studentId, "");
      } else {
        // Fallback if no student ID
        letterType = "welcome";
        UtilityScriptLibrary.debugLog("populateLetterType", "WARNING", "No student ID found, defaulting to welcome", "", "");
      }
    } else if (sourceType === "previous" && prevRow) {
      // Students from previous billing cycle - check their current balance
      var prevCurrentBalanceCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Current Balance")];
      
      if (prevCurrentBalanceCol) {
        var previousBalance = parseFloat(prevRow[prevCurrentBalanceCol - 1]) || 0;
        
        if (previousBalance > 0) {
          letterType = "second";
          UtilityScriptLibrary.debugLog("populateLetterType", "DEBUG", "Set letter type from previous with balance", 
                        "Type: second, Balance: " + previousBalance, "");
        } else {
          UtilityScriptLibrary.debugLog("populateLetterType", "DEBUG", "Previous student with no balance, leaving blank", 
                        "Balance: " + previousBalance, "");
        }
      } else {
        UtilityScriptLibrary.debugLog("populateLetterType", "WARNING", "Current Balance column not found in previous data", "", "");
      }
    }
    
    newRow[letterTypeCol - 1] = letterType;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateLetterType", "ERROR", "Failed to populate letter type", 
                  "", error.message);
    // Don't throw - just log the error and continue
  }
}

function applyLetterTypeValidation(billingSheet) {
  try {
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    var letterTypeCol = headerMap[UtilityScriptLibrary.normalizeHeader("Letter Type")];
    
    if (!letterTypeCol) {
      UtilityScriptLibrary.debugLog('applyLetterTypeValidation', 'WARNING', 'Letter Type column not found', '', '');
      return;
    }
    
    var lastRow = billingSheet.getLastRow();
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog('applyLetterTypeValidation', 'INFO', 'No data rows to format', '', '');
      return;
    }
    
    // Apply Letter Type dropdown validation to all data rows at once
    var letterTypeOptions = ['welcome', 'returning', 'second', 'revised', 'missing', ''];
    var letterTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(letterTypeOptions, true)
      .setAllowInvalid(false)
      .build();
    
    billingSheet.getRange(2, letterTypeCol, lastRow - 1, 1).setDataValidation(letterTypeRule);
    
    UtilityScriptLibrary.debugLog('applyLetterTypeValidation', 'SUCCESS', 'Letter Type dropdown applied to all rows', 
                                  'Column: ' + letterTypeCol + ', Rows: ' + (lastRow - 1), '');
  } catch (error) {
    UtilityScriptLibrary.debugLog('applyLetterTypeValidation', 'ERROR', 'Failed to apply Letter Type validation', 
                                  '', error.message);
  }
}

function createBillingSheet(billingMonth) {
  UtilityScriptLibrary.debugLog("createBillingSheet", "INFO", "Creating billing sheet", 
                "Billing month: " + billingMonth, "");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var templateSheet = ss.getSheetByName("Billing Template");
    
    if (!templateSheet) {
      throw new Error("Billing Template sheet not found.");
    }

    // Sanitize name (ES5 compatible)
    var safeName = billingMonth.replace(/[\\\/:*?\[\]\n\r]/g, '').substring(0, 100);
    if (!safeName) {
      throw new Error("Billing month name is invalid or empty.");
    }
    
    if (ss.getSheetByName(safeName)) {
      throw new Error('A sheet named "' + safeName + '" already exists.');
    }

    // Use utility function for sheet copying with protections
    var copyResult = UtilityScriptLibrary.copySheetWithProtections(templateSheet, ss, safeName);
    if (!copyResult.success) {
      throw new Error("Failed to copy template sheet: " + copyResult.error);
    }
    
    var newSheet = copyResult.sheet;
    ss.setActiveSheet(newSheet);
    
    // Move new sheet to first position (leftmost)
    ss.moveActiveSheet(1);

    // Find dynamic columns marker
    var headerRow = newSheet.getRange(1, 1, 1, newSheet.getLastColumn()).getValues()[0];
    var markerIndex = -1;
    for (var i = 0; i < headerRow.length; i++) {
      if (headerRow[i] === '<<DYNAMIC_COLUMNS_START>>') {
        markerIndex = i;
        break;
      }
    }
    
    if (markerIndex === -1) {
      throw new Error("Marker <<DYNAMIC_COLUMNS_START>> not found.");
    }

    // Get active programs for dynamic headers
    var programSheet = ss.getSheetByName("Programs List");
    var data = programSheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol = headers.indexOf("Program Name");
    var activeCol = headers.indexOf("Active");
    var prefixCol = headers.indexOf("Billing Sheet Prefix");
    var typeCol = headers.indexOf("Type");

    var dynamicHeaders = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[activeCol] === true && row[typeCol] !== "Package") {
        var prefix = row[prefixCol] ? row[prefixCol].toString().trim() : '';
        if (prefix) {
          // Updated column order for Lesson programs
          if (prefix.toLowerCase() === 'lesson') {
            dynamicHeaders.push(
              prefix + ' Quantity', 
              prefix + ' Credit',
              prefix + ' Hours',
              prefix + ' Price',
              prefix + ' Total'
            );
          } else {
            dynamicHeaders.push(
              prefix + ' Quantity',
              prefix + ' Credit', 
              prefix + ' Price',
              prefix + ' Total'
            );
          }
        }
      }
    }

    // Insert dynamic columns
    var insertAt = markerIndex + 2;
    newSheet.insertColumnsBefore(insertAt, dynamicHeaders.length);
    newSheet.getRange(1, insertAt, 1, dynamicHeaders.length).setValues([dynamicHeaders]);
    newSheet.deleteColumn(markerIndex + 1);

    // REMOVED: newSheet.getDataRange().clearDataValidations();
    // This line was clearing ALL data validations including the Letter Type dropdown
    // Data validations don't interfere with number formatting, so this line serves no purpose
    
    UtilityScriptLibrary.debugLog("createBillingSheet", "INFO", "Billing sheet created successfully", 
                  "Sheet: " + safeName + ", Dynamic columns: " + dynamicHeaders.length, "");
    
    return newSheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("createBillingSheet", "ERROR", "Failed to create billing sheet", 
                  "Billing month: " + billingMonth, error.message);
    throw error;
  }
}

function detectAndBillOverages(billingSheet, billingCycleName) {
  try {
    UtilityScriptLibrary.debugLog('Detecting lesson overages for billing cycle: ' + billingCycleName);
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    var lastRow = billingSheet.getLastRow();
    
    // Check if there are any data rows (must have at least row 2)
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog('No data rows found in billing sheet - skipping overage detection');
      return [];
    }
    
    var numDataRows = lastRow - 1; // Subtract header row
    if (numDataRows <= 0) {
      UtilityScriptLibrary.debugLog('No data rows to process - skipping overage detection');
      return [];
    }
    
    var dataRange = billingSheet.getRange(2, 1, numDataRows, billingSheet.getLastColumn());
    var data = dataRange.getValues();
    
    // Use the actual column names from your billing template
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var currentCumTaughtCol = headerMap[UtilityScriptLibrary.normalizeHeader('Current Cumulative Hours Taught')];
    var currentCumBilledCol = headerMap[UtilityScriptLibrary.normalizeHeader('Current Cumulative Hours Billed')];
    var lessonLengthCol = headerMap[UtilityScriptLibrary.normalizeHeader('Lesson Length')];
    var lessonHoursCol = headerMap[UtilityScriptLibrary.normalizeHeader('Lesson Hours')];
    
    // Check for required columns
    if (!studentIdCol || !lessonLengthCol) {
      UtilityScriptLibrary.debugLog('Required overage detection columns not found - Student ID: ' + studentIdCol + ', Lesson Length: ' + lessonLengthCol);
      return [];
    }
    
    // If cumulative columns don't exist, skip overage detection
    if (!currentCumTaughtCol || !currentCumBilledCol || !lessonHoursCol) {
      UtilityScriptLibrary.debugLog('Cumulative hours columns not found - skipping overage detection for this cycle');
      return [];
    }
    
    var overageStudents = [];
    
    // Check each student for overages
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var studentId = row[studentIdCol - 1];
      
      if (!studentId) continue;
      
      var registeredHours = parseFloat(row[lessonHoursCol - 1]) || 0;
      var taughtHours = parseFloat(row[currentCumTaughtCol - 1]) || 0;
      var billedHours = parseFloat(row[currentCumBilledCol - 1]) || 0;
      var lessonLength = row[lessonLengthCol - 1] || '30';
      
      // Calculate unbilled overages: hours taught beyond registered amount that haven't been billed yet
      var unbilledOverage = Math.max(0, taughtHours - registeredHours - (billedHours - registeredHours));
      
      if (unbilledOverage > 0) {
        overageStudents.push({
          studentId: studentId,
          overageHours: unbilledOverage,
          lessonLength: lessonLength,
          rowIndex: i + 2, // +2 for header row and 0-based index
          registeredHours: registeredHours,
          taughtHours: taughtHours,
          billedHours: billedHours
        });
        
        UtilityScriptLibrary.debugLog('Overage detected - Student ' + studentId + ': ' + unbilledOverage + ' hour equivalents');
      }
    }
    
    // Add overages to billing sheet if any found
    if (overageStudents.length > 0) {
      addOveragesToBillingSheet(billingSheet, overageStudents, headerMap);
    }
    
    UtilityScriptLibrary.debugLog('Overage detection complete: ' + overageStudents.length + ' students with overages');
    return overageStudents;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('Error detecting overages: ' + error.message);
    throw error;
  }
}

function formatRow(sheet, rowIndex, quantityCols, currencyCols) {
  UtilityScriptLibrary.debugLog('formatRow', 'DEBUG', 'Formatting row ' + rowIndex, 
                                'Quantity cols: ' + quantityCols.length + ', Currency cols: ' + currencyCols.length, '');
  
  // Get header map to identify lesson length column and letter type column
  var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
  var lessonLengthCol = headerMap[UtilityScriptLibrary.normalizeHeader("Lesson Length")];
  var letterTypeCol = headerMap[UtilityScriptLibrary.normalizeHeader("Letter Type")];

  // Format quantity columns
  quantityCols.forEach(function(col) {
    if (typeof col === 'number' && col > 0) {
      if (col === lessonLengthCol) {
        // Lesson Length stays as integer
        sheet.getRange(rowIndex, col).setNumberFormat("0");
      } else {
        // All other quantity columns get 2 decimal places
        sheet.getRange(rowIndex, col).setNumberFormat("0.00");
      }
    }
  });

  // Format currency columns
  currencyCols.forEach(function(col) {
    if (typeof col === 'number' && col > 0) {
      sheet.getRange(rowIndex, col).setNumberFormat("$#,##0.00");
    }
  });
  
  // Add dropdown validation for Letter Type column
  if (letterTypeCol) {
    var letterTypeOptions = ['welcome', 'returning', 'revised', 'second', 'missing', ''];
    var letterTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(letterTypeOptions, true)  // true = show dropdown
      .setAllowInvalid(false)
      .build();
    
    sheet.getRange(rowIndex, letterTypeCol).setDataValidation(letterTypeRule);
    
    UtilityScriptLibrary.debugLog('formatRow', 'DEBUG', 'Added Letter Type dropdown validation', 
                                  'Row: ' + rowIndex + ', Col: ' + letterTypeCol, '');
  }
  
  // NEW: Format Agreement Form and Media Release as checkboxes
  var agreementCol = headerMap[UtilityScriptLibrary.normalizeHeader('Agreement Form')];
  var mediaCol = headerMap[UtilityScriptLibrary.normalizeHeader('Media Release')];
  
  if (agreementCol) {
    var agreementCell = sheet.getRange(rowIndex, agreementCol);
    agreementCell.insertCheckboxes();
    UtilityScriptLibrary.debugLog('formatRow', 'DEBUG', 'Added Agreement Form checkbox', 
                                  'Row: ' + rowIndex + ', Col: ' + agreementCol, '');
  }
  
  if (mediaCol) {
    var mediaCell = sheet.getRange(rowIndex, mediaCol);
    mediaCell.insertCheckboxes();
    UtilityScriptLibrary.debugLog('formatRow', 'DEBUG', 'Added Media Release checkbox', 
                                  'Row: ' + rowIndex + ', Col: ' + mediaCol, '');
  }
}

function protectPreviousBillingCycle() {
  UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "INFO", "Starting previous billing cycle protection", "", "");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = ss.getSheetByName("Billing Metadata");
    
    if (!metadataSheet) {
      UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "WARNING", "Billing Metadata sheet not found", "", "");
      return;
    }
    
    var lastRow = metadataSheet.getLastRow();
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "WARNING", "No billing cycles found to protect", 
                    "Last row: " + lastRow, "");
      return;
    }
    
    // Get the previous billing cycle name (second to last row)
    var billingMonth;
    if (lastRow === 2) {
      // Only one billing cycle exists, protect it
      billingMonth = metadataSheet.getRange(lastRow, 1).getDisplayValue();
      UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "INFO", "Only one billing cycle exists, protecting current", 
                    "Billing month: " + billingMonth, "");
    } else {
      // Protect the previous cycle (second to last row)
      billingMonth = metadataSheet.getRange(lastRow, 1).getDisplayValue();
      UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "INFO", "Protecting previous billing cycle", 
                    "Billing month: " + billingMonth, "");
    }
    
    if (!billingMonth) {
      UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "WARNING", "No billing month name found", "", "");
      return;
    }
    
    var billingSheet = ss.getSheetByName(billingMonth);
    if (!billingSheet) {
      UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "WARNING", "Billing sheet not found", 
                    "Sheet name: " + billingMonth, "");
      return;
    }
    
    // Protect the entire sheet with warning-only mode
    protectBillingSheet(billingSheet);
    
    UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "INFO", "Successfully protected billing cycle", 
                  "Sheet: " + billingMonth, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("protectPreviousBillingCycle", "ERROR", "Failed to protect previous billing cycle", 
                  "", error.message);
    throw error;
  }
}

function generateProgramFormulas(newRow, context, rowNum, quantityCols, currencyCols) {
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  for (var programName in context.programMap) {
    var program = context.programMap[programName];
    var prefix = program.prefix;
    
    var qtyCol = context.headerMap[norm(prefix + " Quantity")];
    var creditCol = context.headerMap[norm(prefix + " Credit")];
    var priceCol = context.headerMap[norm(prefix + " Price")];
    var totalCol = context.headerMap[norm(prefix + " Total")];
    var hoursCol = context.headerMap[norm(prefix + " Hours")];
    
    // Generate Hours formula (lesson programs only)
    if (hoursCol && prefix.toLowerCase() === "lesson") {
      var lengthCol = context.headerMap[norm("Lesson Length")];
      if (lengthCol && qtyCol && creditCol) {
        var qtyCell = UtilityScriptLibrary.columnToLetter(qtyCol) + rowNum;
        var creditCell = UtilityScriptLibrary.columnToLetter(creditCol) + rowNum;
        var lengthCell = UtilityScriptLibrary.columnToLetter(lengthCol) + rowNum;
        var hoursFormula = "=(" + qtyCell + " - " + creditCell + ") * " + lengthCell + " / 60";
        newRow[hoursCol - 1] = hoursFormula;
        
        // Track hours columns for special formatting
        if (!context.hoursColumnsToFormat) {
          context.hoursColumnsToFormat = [];
        }
        context.hoursColumnsToFormat.push(hoursCol);
      }
    }
    
    // Generate Total formula
    if (totalCol && priceCol) {
      if (prefix.toLowerCase() === "lesson" && hoursCol) {
        // For lessons: Total = Hours * HourlyRate
        var hoursCell = UtilityScriptLibrary.columnToLetter(hoursCol) + rowNum;
        var priceCell = UtilityScriptLibrary.columnToLetter(priceCol) + rowNum;
        var totalFormula = "=" + hoursCell + " * " + priceCell;
        newRow[totalCol - 1] = totalFormula;
      } else if (qtyCol && creditCol) {
        // For other programs: Total = (Quantity - Credit) * Price
        var qtyCell = UtilityScriptLibrary.columnToLetter(qtyCol) + rowNum;
        var creditCell = UtilityScriptLibrary.columnToLetter(creditCol) + rowNum;
        var priceCell = UtilityScriptLibrary.columnToLetter(priceCol) + rowNum;
        var totalFormula = "=(" + qtyCell + " - " + creditCell + ") * " + priceCell;
        newRow[totalCol - 1] = totalFormula;
      }
      currencyCols.push(totalCol);
    }
  }
}

function populateAllCumulativeColumns() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheet = ss.getActiveSheet();
    var headerMap = UtilityScriptLibrary.getHeaderMap(currentSheet);
    
    // RANGE VALIDATION: Check if sheet has data
    var lastRow = currentSheet.getLastRow();
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog('No data rows found in sheet - skipping cumulative columns population');
      return;
    }
    
    // Find ALL required columns using new names
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader("Student ID")];
    var currentHoursTaughtThisCycleCol = headerMap[UtilityScriptLibrary.normalizeHeader("Current Hours Taught This Billing Cycle")];
    var lessonHoursCol = headerMap[UtilityScriptLibrary.normalizeHeader("Lesson Hours")];
    
    // Past cumulative columns
    var pastCumTaughtCol = headerMap[UtilityScriptLibrary.normalizeHeader("Past Cumulative Hours Taught")];
    var pastCumBilledCol = headerMap[UtilityScriptLibrary.normalizeHeader("Past Cumulative Hours Billed")];
    
    // Current cumulative columns (these will get formulas)
    var currentCumTaughtCol = headerMap[UtilityScriptLibrary.normalizeHeader("Current Cumulative Hours Taught")];
    var currentCumBilledCol = headerMap[UtilityScriptLibrary.normalizeHeader("Current Cumulative Hours Billed")];
    
    // Hours Remaining column
    var hoursRemainingCol = headerMap[UtilityScriptLibrary.normalizeHeader("Hours Remaining")];
    
    if (!studentIdCol || !currentHoursTaughtThisCycleCol || !lessonHoursCol ||
        !pastCumTaughtCol || !pastCumBilledCol || !currentCumTaughtCol || 
        !currentCumBilledCol || !hoursRemainingCol) {
      UtilityScriptLibrary.debugLog('Required columns not found for cumulative hours calculations');
      return;
    }
    
    // RANGE VALIDATION: Ensure we have data rows to process
    var numDataRows = lastRow - 1;
    if (numDataRows <= 0) {
      UtilityScriptLibrary.debugLog('No data rows to process in cumulative columns population');
      return;
    }
    
    var dataRange = currentSheet.getRange(2, 1, numDataRows, currentSheet.getLastColumn());
    var data = dataRange.getValues();
    
    // Process each student row
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rowIndex = i + 2; // Row numbers are 1-indexed, data starts at row 2
      var studentId = row[studentIdCol - 1];
      
      if (!studentId) continue;
      
      // Generate formulas for cumulative columns
      var pastTaughtCell = UtilityScriptLibrary.columnToLetter(pastCumTaughtCol) + rowIndex;
      var currentTaughtCell = UtilityScriptLibrary.columnToLetter(currentHoursTaughtThisCycleCol) + rowIndex;
      var pastBilledCell = UtilityScriptLibrary.columnToLetter(pastCumBilledCol) + rowIndex;
      var lessonHoursCell = UtilityScriptLibrary.columnToLetter(lessonHoursCol) + rowIndex;
      
      // Current Cumulative Hours Taught = Past + Current This Cycle
      var currentCumTaughtFormula = "=" + pastTaughtCell + " + " + currentTaughtCell;
      currentSheet.getRange(rowIndex, currentCumTaughtCol).setFormula(currentCumTaughtFormula);
      currentSheet.getRange(rowIndex, currentCumTaughtCol).setNumberFormat("0.00");
      
      // Current Cumulative Hours Billed = Past + Lesson Hours
      var currentCumBilledFormula = "=" + pastBilledCell + " + " + lessonHoursCell;
      currentSheet.getRange(rowIndex, currentCumBilledCol).setFormula(currentCumBilledFormula);
      currentSheet.getRange(rowIndex, currentCumBilledCol).setNumberFormat("0.00");
      
      // Hours Remaining = Current Cumulative Taught - Current Cumulative Billed
      var currentCumTaughtCell = UtilityScriptLibrary.columnToLetter(currentCumTaughtCol) + rowIndex;
      var currentCumBilledCell = UtilityScriptLibrary.columnToLetter(currentCumBilledCol) + rowIndex;
      var hoursRemainingFormula = "=" + currentCumBilledCell + " - " + currentCumTaughtCell;
      currentSheet.getRange(rowIndex, hoursRemainingCol).setFormula(hoursRemainingFormula);
      currentSheet.getRange(rowIndex, hoursRemainingCol).setNumberFormat("0.00");
    }
    
    UtilityScriptLibrary.debugLog('Successfully populated cumulative columns for ' + data.length + ' students');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('Error populating cumulative columns: ' + error.message);
    throw error;
  }
}



function populateCurrentBalanceFormula(newRow, context, rowIndex) {
  var h = context.headerMap;
  var invoiceTotalCol = h[UtilityScriptLibrary.normalizeHeader("Current Invoice Total")];
  var paymentReceivedCol = h[UtilityScriptLibrary.normalizeHeader("Payment Received")];
  var currentBalanceCol = h[UtilityScriptLibrary.normalizeHeader("Current Balance")];

  if (invoiceTotalCol && paymentReceivedCol && currentBalanceCol && rowIndex) {
    var formula = `=${UtilityScriptLibrary.columnToLetter(invoiceTotalCol)}${rowIndex} - ${UtilityScriptLibrary.columnToLetter(paymentReceivedCol)}${rowIndex}`;
    newRow[currentBalanceCol - 1] = formula;
  }
}

function populateDeliveryPreference(newRow, formRow, context) {
  try {
    var deliveryPrefCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Delivery Preference")];
    if (!deliveryPrefCol) {
      UtilityScriptLibrary.debugLog("populateDeliveryPreference", "WARNING", "Delivery Preference column not found", "", "");
      return;
    }
    
    // Get delivery preference from form data
    var deliveryPrefFromForm = context.getColIndex("Delivery Preference");
    var deliveryPreference = "Email"; // Default value
    
    if (deliveryPrefFromForm !== -1 && formRow[deliveryPrefFromForm]) {
      deliveryPreference = formRow[deliveryPrefFromForm];
    }
    
    newRow[deliveryPrefCol - 1] = deliveryPreference;
    
    UtilityScriptLibrary.debugLog("populateDeliveryPreference", "DEBUG", "Set delivery preference", 
                  "Preference: " + deliveryPreference, "");
                  
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateDeliveryPreference", "ERROR", "Failed to populate delivery preference", 
                  "", error.message);
    // Don't throw - just log the error and continue
  }
}

function populateDeliveryPreferenceFromPrevious(newRow, prevRow, context) {
  try {
    var deliveryPrefCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Delivery Preference")];
    var prevDeliveryPrefCol = context.prevHeaderMap[UtilityScriptLibrary.normalizeHeader("Delivery Preference")];
    
    if (!deliveryPrefCol) {
      UtilityScriptLibrary.debugLog("populateDeliveryPreferenceFromPrevious", "WARNING", "Delivery Preference column not found in current sheet", "", "");
      return;
    }
    
    var deliveryPreference = "Email"; // Default value
    
    // Try to get from previous billing cycle
    if (prevDeliveryPrefCol && prevRow[prevDeliveryPrefCol - 1]) {
      deliveryPreference = prevRow[prevDeliveryPrefCol - 1];
    }
    
    newRow[deliveryPrefCol - 1] = deliveryPreference;
    
    UtilityScriptLibrary.debugLog("populateDeliveryPreferenceFromPrevious", "DEBUG", "Carried over delivery preference", 
                  "Preference: " + deliveryPreference, "");
                  
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateDeliveryPreferenceFromPrevious", "ERROR", "Failed to populate delivery preference from previous", 
                  "", error.message);
    // Don't throw - just log the error and continue
  }
}

function populateInvoiceMetadata(newRow, studentId, context, rowIndex) {
  var h = context.headerMap;
  var today = context.today;
  var year = parseInt(today.slice(0, 4), 10);
  var month = parseInt(today.slice(4, 6), 10);
  var day = parseInt(today.slice(6), 10);

  var invoiceNumberCol = h[UtilityScriptLibrary.normalizeHeader("Invoice Number")];
  var invoiceDateCol = h[UtilityScriptLibrary.normalizeHeader("Invoice Date")];
  var dueDateCol = h[UtilityScriptLibrary.normalizeHeader("Due Date")];
  var invoiceUrlCol = h[UtilityScriptLibrary.normalizeHeader("Invoice URL")];

  if (invoiceNumberCol) {
    newRow[invoiceNumberCol - 1] = `${studentId}-${today}`;
  }
  if (invoiceDateCol) {
    newRow[invoiceDateCol - 1] = `=DATE(${year}, ${month}, ${day})`;
  }
  if (dueDateCol) {
    newRow[dueDateCol - 1] = `=DATE(${year}, ${month}, ${day})+14`;
  }
  if (invoiceUrlCol) {
    newRow[invoiceUrlCol - 1] = "";
  }
}

function populateLateFee(newRow, context, currencyCols) {
  var h = context.headerMap;
  var lateFeeCol = h[UtilityScriptLibrary.normalizeHeader("Late Fee")];
  var pastBalanceCol = h[UtilityScriptLibrary.normalizeHeader("Past Balance")];
  var paymentReceivedCol = h[UtilityScriptLibrary.normalizeHeader("Payment Received")];
  var pastInvoiceNumberCol = h[UtilityScriptLibrary.normalizeHeader("Past Invoice Number")];
  
  if (lateFeeCol && pastBalanceCol && paymentReceivedCol && pastInvoiceNumberCol) {
    var pastBalance = parseFloat(newRow[pastBalanceCol - 1]) || 0;
    var paymentReceived = parseFloat(newRow[paymentReceivedCol - 1]) || 0;
    var pastInvoiceNumber = newRow[pastInvoiceNumberCol - 1] || '';
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ratesSheet = ss.getSheetByName("Rates");
    var rateHeaders = ratesSheet.getRange(1, 1, 1, ratesSheet.getLastColumn()).getValues()[0];
    var rateYearCol = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
    var rateMap = UtilityScriptLibrary.buildRateMapFromSheet(ratesSheet, rateHeaders, rateYearCol);
    
    var lessonRate = parseFloat(rateMap["Lessons"]) || 0;
    var gracePeriod = parseInt(rateMap["Grace Period"]) || 10;
    
    if (pastBalance > lessonRate && paymentReceived === 0 && pastInvoiceNumber) {
      var dateMatch = pastInvoiceNumber.match(/-(\d{8})$/);
      if (dateMatch) {
        var invoiceDateStr = dateMatch[1];
        var invoiceDate = new Date(
          parseInt(invoiceDateStr.substr(0,4)),
          parseInt(invoiceDateStr.substr(4,2)) - 1,
          parseInt(invoiceDateStr.substr(6,2))
        );
        var dueDate = new Date(invoiceDate);
        dueDate.setDate(dueDate.getDate() + 14);
        
        var lateFeeDate = new Date(dueDate);
        lateFeeDate.setDate(lateFeeDate.getDate() + gracePeriod);
        
        var today = new Date();
        if (today > lateFeeDate) {
          var lateFeeAmount = parseFloat(rateMap["Late Fee"]) || 0;
          if (lateFeeAmount > 0) {
            newRow[lateFeeCol - 1] = lateFeeAmount;
            UtilityScriptLibrary.addToCurrencyCols(currencyCols, lateFeeCol, "Late Fee");
          }
        }
      }
    }
  }
}

function populatePastBalanceAndCredit(newRow, pastRow, context, currencyCols) {
  var h = context.headerMap;
  var prevH = context.prevHeaderMap;
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  // === MONETARY BALANCE HANDLING ===
  var pastBalanceCol = h[norm("Past Balance")];
  var creditCol = h[norm("Credit")];
  var currentBalanceCol = prevH[norm("Current Balance")];

  if (currentBalanceCol && Array.isArray(pastRow) && pastRow.length >= currentBalanceCol) {
    var rawValue = pastRow[currentBalanceCol - 1];
    var previousBalance = typeof rawValue === "string"
      ? parseFloat(rawValue.replace(/[^0-9.\-]/g, ""))
      : parseFloat(rawValue);

    if (previousBalance !== 0 && pastBalanceCol) {
      newRow[pastBalanceCol - 1] = previousBalance; // Keep the sign (+/-)
      UtilityScriptLibrary.addToCurrencyCols(currencyCols, pastBalanceCol, "Past Balance");
    }
  }
  
  // === PAST INVOICE NUMBER ===
  var pastInvoiceNumberCol = h[norm("Past Invoice Number")];
  var prevInvoiceNumberCol = prevH[norm("Invoice Number")];
  if (pastInvoiceNumberCol && prevInvoiceNumberCol) {
    newRow[pastInvoiceNumberCol - 1] = pastRow[prevInvoiceNumberCol - 1] || '';
  }
  
  // === PAST CUMULATIVE HOURS (CRITICAL FOR CONTINUING CYCLES) ===
  var pastCumTaughtCol = h[norm("Past Cumulative Hours Taught")];
  var pastCumBilledCol = h[norm("Past Cumulative Hours Billed")];
  var prevCumTaughtCol = prevH[norm("Current Cumulative Hours Taught")];
  var prevCumBilledCol = prevH[norm("Current Cumulative Hours Billed")];
  
  if (pastCumTaughtCol && prevCumTaughtCol) {
    var prevCumTaught = pastRow[prevCumTaughtCol - 1];
    newRow[pastCumTaughtCol - 1] = (typeof prevCumTaught === 'number' || typeof prevCumTaught === 'string') 
      ? parseFloat(prevCumTaught) || 0 
      : 0;
    
    UtilityScriptLibrary.debugLog('populatePastBalanceAndCredit', 'DEBUG', 
                                  'Set Past Cumulative Hours Taught',
                                  'Value: ' + newRow[pastCumTaughtCol - 1], '');
  }
  
  if (pastCumBilledCol && prevCumBilledCol) {
    var prevCumBilled = pastRow[prevCumBilledCol - 1];
    newRow[pastCumBilledCol - 1] = (typeof prevCumBilled === 'number' || typeof prevCumBilled === 'string') 
      ? parseFloat(prevCumBilled) || 0 
      : 0;
    
    UtilityScriptLibrary.debugLog('populatePastBalanceAndCredit', 'DEBUG', 
                                  'Set Past Cumulative Hours Billed',
                                  'Value: ' + newRow[pastCumBilledCol - 1], '');
  }
  
  // === HOURS REMAINING HANDLING ===
  var pastHoursRemainingCol = h[norm("Past Hours Remaining")];
  var lessonCreditCol = h[norm("Lesson Credit")];
  var lessonQuantityCol = h[norm("Lesson Quantity")];
  var lessonLengthCol = h[norm("Lesson Length")];
  var prevHoursRemainingCol = prevH[norm("Hours Remaining")];
  
  if (!pastHoursRemainingCol) {
    UtilityScriptLibrary.debugLog('populatePastBalanceAndCredit', 'WARNING', 
                                  'Past Hours Remaining column not found', '', '');
    return;
  }
  
  // Get previous Hours Remaining
  var pastHoursRemaining = 0;
  if (prevHoursRemainingCol && Array.isArray(pastRow) && pastRow.length >= prevHoursRemainingCol) {
    var rawHours = pastRow[prevHoursRemainingCol - 1];
    pastHoursRemaining = typeof rawHours === "string"
      ? parseFloat(rawHours.replace(/[^0-9.\-]/g, ""))
      : parseFloat(rawHours);
    
    if (isNaN(pastHoursRemaining)) {
      pastHoursRemaining = 0;
    }
  }
  
  // Set Past Hours Remaining
  newRow[pastHoursRemainingCol - 1] = pastHoursRemaining;
  
  UtilityScriptLibrary.debugLog('populatePastBalanceAndCredit', 'DEBUG', 
                                'Set Past Hours Remaining',
                                'Value: ' + pastHoursRemaining, '');
  
  // Calculate and set Lesson Credit ONLY if there's a quantity to bill
  if (lessonCreditCol && lessonQuantityCol && lessonLengthCol && pastHoursRemaining > 0) {
    var lessonQuantity = parseFloat(newRow[lessonQuantityCol - 1]) || 0;
    var lessonLength = parseFloat(newRow[lessonLengthCol - 1]) || 0;
    
    if (lessonQuantity > 0 && lessonLength > 0) {
      var lessonCredit = pastHoursRemaining / (lessonLength / 60);
      newRow[lessonCreditCol - 1] = lessonCredit;
      
      UtilityScriptLibrary.debugLog('populatePastBalanceAndCredit', 'INFO', 
                                    'Calculated Lesson Credit',
                                    'Past Hours: ' + pastHoursRemaining + 
                                    ', Length: ' + lessonLength + 
                                    ', Quantity: ' + lessonQuantity +
                                    ', Credit: ' + lessonCredit, '');
    } else {
      // No quantity to bill - don't apply credit
      newRow[lessonCreditCol - 1] = 0;
      
      UtilityScriptLibrary.debugLog('populatePastBalanceAndCredit', 'DEBUG', 
                                    'No credit applied - no quantity',
                                    'Quantity: ' + lessonQuantity, '');
    }
  } else if (lessonCreditCol) {
    // No credit to apply
    newRow[lessonCreditCol - 1] = 0;
  }
}

function processSemesterEndCredits(currentSemesterName) {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    UtilityScriptLibrary.debugLog('💳 processSemesterEndCredits - Starting credit processing for: ' + currentSemesterName);
    
    var currentBillingSheet = getCurrentBillingSheet();
    if (!currentBillingSheet) {
      UtilityScriptLibrary.debugLog('💳 processSemesterEndCredits - No current billing sheet found - skipping credit processing');
      return [];
    }
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(currentBillingSheet);
    var dataRange = currentBillingSheet.getRange(2, 1, currentBillingSheet.getLastRow() - 1, currentBillingSheet.getLastColumn());
    var data = dataRange.getValues();
    
    // Required column indices using utility normalization
    var norm = UtilityScriptLibrary.normalizeHeader;
    var studentIdCol = headerMap[norm('Student ID')];
    var hoursTaughtCol = headerMap[norm('Current Cumulative Hours Taught')];
    var lessonLengthCol = headerMap[norm('Lesson Length')];
    var lessonHoursCol = headerMap[norm('Lesson Hours')];
    
    if (!studentIdCol || !hoursTaughtCol || !lessonLengthCol || !lessonHoursCol) {
      UtilityScriptLibrary.debugLog('💳 processSemesterEndCredits - Required columns not found - skipping credit processing');
      return [];
    }
    
    var creditBalances = [];
    
    // Process each student
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var studentId = row[studentIdCol - 1];
      
      if (!studentId) continue;
      
      var registeredHours = parseFloat(row[lessonHoursCol - 1]) || 0;
      var taughtHours = parseFloat(row[hoursTaughtCol - 1]) || 0;
      var lessonLength = parseFloat(row[lessonLengthCol - 1]) || 30;
      
      // Calculate credit balance (hours paid for but not taken)
      var creditHours = registeredHours - taughtHours;
      
      if (creditHours > 0) {
        creditBalances.push({
          studentId: studentId,
          creditLessons: creditHours,
          lessonLength: lessonLength,
          semesterName: currentSemesterName,
          registeredLessons: registeredHours,
          taughtLessons: taughtHours,
          lessonBalance: creditHours,
          status: 'pending'
        });
        
        UtilityScriptLibrary.debugLog('💰 Credit balance for student ' + studentId + ': ' + 
                                    creditHours + ' hours at ' + lessonLength + 'min');
      }
    }
    
    // Store credit balances for next semester
    if (creditBalances.length > 0) {
      storeSemesterEndBalances(creditBalances);
      UtilityScriptLibrary.debugLog('✅ Stored ' + creditBalances.length + ' student credit balances');
    } else {
      UtilityScriptLibrary.debugLog('ℹ️ No credit balances to process');
    }
    
    return creditBalances;
    
  }, 'Semester end credits processed successfully', 'processSemesterEndCredits', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

// ============================================================================
// SECTION 4: GENERATE DOCUMENTS
// ============================================================================

function buildDocumentFileName(studentData, billingData, documentType, deliveryMethod) {
  var documentTypeFormatted = documentType.charAt(0).toUpperCase() + documentType.slice(1);
  
  var invoiceNumber = '';
  if (billingData && billingData.invoiceNumber) {
    invoiceNumber = billingData.invoiceNumber;
  } else {
    // Fallback: create invoice number format using student ID and current date
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    invoiceNumber = studentData.studentId + '-' + today;
  }
  
  return invoiceNumber + ' - ' + studentData.firstName + ' ' + studentData.lastName + ' - ' + documentTypeFormatted;
}

function buildDocumentSentence(billingData, deliveryMethod) {
  try {
    var documents = [];
    
    // Check Agreement Form - if FALSE or missing, we need it
    if (!billingData.agreementForm) {
      documents.push('Agreement Form');
    }
    
    // Check Media Release - if FALSE or missing, we need it
    if (!billingData.mediaRelease) {
      documents.push('Media Consent Form');
    }
    
    // Determine prefix based on delivery method
    var prefix = (deliveryMethod && deliveryMethod.toLowerCase() === 'email') ? 'Attached' : 'Enclosed';
    
    // Build the sentence
    var sentence = prefix + ' please find your invoice';
    
    if (documents.length === 0) {
      // Just invoice, no additional forms needed
      sentence += '.';
    } else if (documents.length === 1) {
      // Invoice + one form
      sentence += ', as well as your ' + documents[0] + '. Please complete and return this form at your earliest convenience.';
    } else {
      // Invoice + two forms
      sentence += ', as well as your ' + documents[0] + ' and ' + documents[1] + '. Please complete and return these forms at your earliest convenience.';
    }
    
    return sentence;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('buildDocumentSentence', 'ERROR', 'Failed to build document sentence', 
                                  '', error.message);
    return 'Please find your invoice and forms.';
  }
}

function buildMissingDocumentSentence(billingData) {
  try {
    var documents = [];
    
    // Check Agreement Form - if FALSE or missing, we need it
    if (!billingData.agreementForm) {
      documents.push('your Agreement Form');
    }
    
    // Check Media Release - if FALSE or missing, we need it
    if (!billingData.mediaRelease) {
      documents.push('your Media Consent Form');
    }
    
    // Build the sentence
    var sentence = '';
    
    if (documents.length === 0) {
      // No documents missing - shouldn't normally happen for a missing letter
      sentence = 'the required documentation';
    } else if (documents.length === 1) {
      // One document missing
      sentence = documents[0];
    } else {
      // Both documents missing
      sentence = documents[0] + ' and ' + documents[1];
    }
    
    return sentence;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('buildMissingDocumentSentence', 'ERROR', 'Failed to build missing document sentence', 
                                  '', error.message);
    return 'the required documentation';
  }
}


function buildInvoiceVariableMap(studentData, billingData, isRefund) {
  var studentFullName = studentData.firstName + ' ' + studentData.lastName;
  var billingFullName = billingData.parentFirstName + ' ' + billingData.parentLastName;
  
  var billingHonorific = '';
  if (billingData.termOfAddress) {
    billingHonorific = billingData.termOfAddress;
  } else {
    billingHonorific = 'Dear ' + billingData.parentFirstName;
  }
  
  var invoiceDate = billingData.invoiceDate;
  if (typeof invoiceDate === 'string') {
    invoiceDate = new Date(invoiceDate);
  }
  var formattedInvoiceDate = UtilityScriptLibrary.formatDateFlexible(invoiceDate, 'MMMM d, yyyy');
  
  var dueDate = billingData.dueDate;
  if (typeof dueDate === 'string') {
    dueDate = new Date(dueDate);
  }
  var formattedDueDate = UtilityScriptLibrary.formatDateFlexible(dueDate, 'MMMM d, yyyy');
  
  var balance = billingData.currentBalance;
  var refundAmount = isRefund ? Math.abs(balance) : 0;
  var amountDue = isRefund ? 0 : Math.max(0, balance);
  
  var dynamicLineItems = buildDynamicLineItems(billingData);
  var dynamicAmounts = buildDynamicAmounts(billingData);
  
  return {
    'InvoiceNumber': billingData.invoiceNumber || '',
    'InvoiceDate': formattedInvoiceDate || '',
    'DueDate': formattedDueDate || '',
    'BillingFullName': billingFullName,
    'BillingAddress': billingData.parentAddress || '',
    'StudentFullName': studentFullName,
    'Instrument': studentData.instrument || '',
    'Teacher': studentData.teacher || '',
    'DYNAMIC_LINE_ITEMS': dynamicLineItems,
    'DYNAMIC_AMOUNTS': dynamicAmounts,
    'BillTotal': UtilityScriptLibrary.formatCurrency(billingData.currentInvoiceTotal || 0),
    'Received': UtilityScriptLibrary.formatCurrency(billingData.paymentReceived || 0),
    'Credit': UtilityScriptLibrary.formatCurrency(calculateTotalCreditsApplied(billingData)),
    'Balance': isRefund ? UtilityScriptLibrary.formatCurrency(refundAmount) : UtilityScriptLibrary.formatCurrency(amountDue),
    'Comment': billingData.comments || ''
  };
}

function buildProgramDescription(programTotals, lessonLength) {
  var descriptions = [];
  
  for (var i = 0; i < programTotals.programs.length; i++) {
    var program = programTotals.programs[i];
    if (program.name.toLowerCase() === 'lesson' && program.quantity > 0) {
      descriptions.push(program.quantity + ' ' + lessonLength + '-minute lessons');
    } else if (program.quantity > 0) {
      descriptions.push(program.quantity + ' ' + program.name.toLowerCase() + ' sessions');
    }
  }
  
  return descriptions.join(', ');
}

function buildTemplateVariables(studentData, billingData, templateType) {
  var variables = {};
  
  variables.StudentFirstName = studentData.firstName || '';
  variables.FirstName = studentData.firstName || '';
  variables.LastName = studentData.lastName || '';
  variables.Instrument = studentData.instrument || '';
  variables.Teacher = studentData.teacher || '';
  variables.LsnLength = studentData.lessonLength || '';
  variables.LsnQuantity = studentData.lessonQuantity || '';
  
  if (billingData) {
    UtilityScriptLibrary.debugLog("buildTemplateVariables", "DEBUG", "Billing data found", 
                  "dueDate raw: " + billingData.dueDate, "");
    
    var billingHonorific = '';
    if (billingData.termOfAddress) {
      billingHonorific = billingData.termOfAddress;
    } else {
      billingHonorific = billingData.parentFirstName || '';
    }
    
    variables.BillingHonorific = billingHonorific;
    variables.BillingLastName = billingData.parentLastName || '';
    
    var dueDate = billingData.dueDate;
    if (typeof dueDate === 'string') {
      dueDate = new Date(dueDate);
    }
    
    var formattedDueDate = UtilityScriptLibrary.formatDateFlexible(dueDate, 'MMMM d, yyyy') || 'TBD';
    
    UtilityScriptLibrary.debugLog("buildTemplateVariables", "DEBUG", "Due date formatting", 
                  "Input: " + billingData.dueDate + ", Output: '" + formattedDueDate + "'", "");
    
    variables.DueDate = formattedDueDate;
    
    // Add DocumentSentence for revised, second, and missing letter types
    if (billingData.letterType) {
      var letterType = billingData.letterType.toLowerCase().trim();
      if (letterType === 'revised' || letterType === 'second') {
        var deliveryMethod = billingData.deliveryPreference || 'print';
        variables.DocumentSentence = buildDocumentSentence(billingData, deliveryMethod);
      } else if (letterType === 'missing') {
        variables.DocumentSentence = buildMissingDocumentSentence(billingData);
      }
    }
  } else {
    UtilityScriptLibrary.debugLog("buildTemplateVariables", "DEBUG", "No billing data provided", "", "");
  }
  
  if (templateType === 'agreement' || templateType === 'invoice') {
    var rateVariables = UtilityScriptLibrary.buildRateVariables();
    for (var rateVar in rateVariables) {
      variables[rateVar] = rateVariables[rateVar];
    }
    
    if (templateType === 'invoice') {
      var invoiceVariables = buildInvoiceVariableMap(studentData, billingData, false);
      for (var invoiceVar in invoiceVariables) {
        variables[invoiceVar] = invoiceVariables[invoiceVar];
      }
    }
    
    if (templateType === 'agreement') {
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var semesterSheet = ss.getSheetByName('Semester Metadata');
        var data = semesterSheet.getDataRange().getValues();
        var latestSemester = data[data.length - 1];
        variables.CurrentAcademicYear = latestSemester[3] || '';
        variables.AcademicYear = latestSemester[3] || '';
        
        var startDate = latestSemester[1];
        if (startDate) {
          variables.AcademicYearStartDate = UtilityScriptLibrary.formatDateFlexible(startDate, 'MMMM d, yyyy') || '';
        }
      } catch (error) {
        UtilityScriptLibrary.debugLog("buildTemplateVariables", "ERROR", "Error getting academic year", "", error.message);
        variables.CurrentAcademicYear = '';
        variables.AcademicYear = '';
        variables.AcademicYearStartDate = '';
      }
    }
  }
  
  return variables;
}

function cancelDocumentGeneration() {
  try {
    // Clear any stored selection
    PropertiesService.getScriptProperties().deleteProperty('documentSelection');
    
    UtilityScriptLibrary.debugLog("cancelDocumentGeneration", "INFO", "Document generation cancelled by user", "", "");
    
    return { cancelled: true };
  } catch (error) {
    UtilityScriptLibrary.debugLog("cancelDocumentGeneration", "ERROR", "Error cancelling generation", "", error.message);
    throw error;
  }
}

function checkIfMediaReleaseNeeded(studentId) {
  try {
    var contactsSS = UtilityScriptLibrary.getWorkbook('contacts');
    var studentsSheet = contactsSS.getSheetByName('students');
    
    if (!studentsSheet) {
      return true; // Default to needing media release if we can't check
    }
    
    var data = studentsSheet.getDataRange().getValues();
    var headers = data[0];
    var norm = UtilityScriptLibrary.normalizeHeader;
    
    var studentIdCol = -1;
    var currentMediaCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var normalizedHeader = norm(headers[i]);
      if (normalizedHeader === norm('Student ID')) {
        studentIdCol = i;
      } else if (normalizedHeader === norm('Current Media')) {
        currentMediaCol = i;
      }
    }
    
    if (studentIdCol === -1 || currentMediaCol === -1) {
      return true; // Default to needing media release
    }
    
    // Find the student and check media status
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol] === studentId) {
        var hasCurrentMedia = data[i][currentMediaCol];
        return !UtilityScriptLibrary.convertYesNoToBoolean(hasCurrentMedia);
      }
    }
    
    return true; // Student not found, assume needs media release
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("checkIfMediaReleaseNeeded", "ERROR", "Error checking media release status", "", error.message);
    return true; // Default to needing media release on error
  }
}

function clearDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap) {
  var columnName = getDocIdColumnName(docType);
  var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
  var docIdCol = headerMap[UtilityScriptLibrary.normalizeHeader(columnName)];
  
  if (!studentIdCol || !docIdCol) return;
  
  var data = billingSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][studentIdCol - 1] === studentId) {
      billingSheet.getRange(i + 1, docIdCol).setValue('');
      return;
    }
  }
}

function continuePacketGeneration() {
  try {
    var selectedDocTypesJson = PropertiesService.getScriptProperties().getProperty('documentSelection');
    var studentsToProcessJson = PropertiesService.getScriptProperties().getProperty('studentsToProcess');
    
    if (!selectedDocTypesJson || !studentsToProcessJson) {
      throw new Error('Missing generation data in properties');
    }
    
    var selectedDocTypes = JSON.parse(selectedDocTypesJson);
    var studentsToProcess = JSON.parse(studentsToProcessJson);
    
    var results = {
      generated: 0,
      skipped: 0,
      errors: 0
    };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getActiveSheet();
    var currentSemester = getCurrentSemesterFromBillingMetadata(billingSheet);
    
    for (var i = 0; i < studentsToProcess.length; i++) {
      var studentInfo = studentsToProcess[i];
      
      try {
        var generationResult = createSingleRegistrationPacketWithSelection(
          studentInfo.studentData,
          studentInfo.billingData,
          studentInfo.deliveryMethod,
          currentSemester,
          null,
          false,
          selectedDocTypes
        );
        
        if (generationResult.documents && generationResult.documents.length > 0) {
          results.generated++;
          UtilityScriptLibrary.debugLog("continuePacketGeneration", "SUCCESS", "Generated packet", 
                        "Student: " + studentInfo.studentData.firstName + " " + studentInfo.studentData.lastName + 
                        ", Documents: " + generationResult.documents.length, "");
        } else if (generationResult.allExisted) {
          results.skipped++;
          UtilityScriptLibrary.debugLog("continuePacketGeneration", "INFO", "All documents already existed", 
                        "Student: " + studentInfo.studentData.firstName + " " + studentInfo.studentData.lastName, 
                        "Skipped");
        }
        
      } catch (studentError) {
        results.errors++;
        UtilityScriptLibrary.debugLog("continuePacketGeneration", "ERROR", "Error processing student", 
                      "Student: " + studentInfo.studentData.firstName + " " + studentInfo.studentData.lastName, 
                      studentError.message);
      }
    }
    
    // Clean up any one-time triggers we created
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'continuePacketGeneration') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    PropertiesService.getScriptProperties().deleteProperty('studentsToProcess');
    PropertiesService.getScriptProperties().deleteProperty('documentSelection');
    
    var message = "Packet generation completed!\n\n";
    message += "Generated: " + results.generated + "\n";
    message += "Skipped: " + results.skipped + "\n";
    message += "Errors: " + results.errors + "\n";
    message += "Document types included: " + selectedDocTypes.join(', ');
    
    SpreadsheetApp.getUi().alert('Generation Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
    UtilityScriptLibrary.debugLog("continuePacketGeneration", "INFO", "Batch packet generation completed", 
                  "Generated: " + results.generated + ", Skipped: " + results.skipped + ", Errors: " + results.errors, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("continuePacketGeneration", "ERROR", "Failed to continue generation", "", error.message);
    SpreadsheetApp.getUi().alert('Error continuing generation: ' + error.message);
  }
}

function createSingleRegistrationPacketWithSelection(studentData, billingData, packetType, currentSemester, destinationFolder, forceRegenerate, selectedDocTypes) {
  // This is similar to createSingleRegistrationPacket but only generates selected document types
  
  var documents = [];
  var errors = [];
  
  // Only generate documents that are in the selection
  if (selectedDocTypes.indexOf('welcome letter') !== -1) {
    var welcomeResult = generateDocumentForStudent(studentData, billingData, 'welcome letter', packetType, currentSemester);
    if (welcomeResult.success && !welcomeResult.alreadyExists) {
      documents.push({
        name: 'Welcome Letter',
        fileId: welcomeResult.fileId,
        url: welcomeResult.url
      });
    } else if (!welcomeResult.success) {
      errors.push('Welcome Letter: ' + welcomeResult.error);
    }
  }
  
  if (selectedDocTypes.indexOf('invoice') !== -1) {
    var invoiceResult = generateDocumentForStudent(studentData, billingData, 'invoice', null, currentSemester);
    if (invoiceResult.success && !invoiceResult.alreadyExists) {
      documents.push({
        name: 'Invoice',
        fileId: invoiceResult.fileId,
        url: invoiceResult.url
      });
    } else if (!invoiceResult.success) {
      errors.push('Invoice: ' + invoiceResult.error);
    }
  }
  
  if (selectedDocTypes.indexOf('agreement') !== -1 && shouldIncludeAgreement(studentData.studentId)) {
    var agreementResult = generateDocumentForStudent(studentData, billingData, 'agreement', null, currentSemester);
    if (agreementResult.success && !agreementResult.alreadyExists) {
      documents.push({
        name: 'Agreement',
        fileId: agreementResult.fileId,
        url: agreementResult.url
      });
    } else if (!agreementResult.success) {
      errors.push('Agreement: ' + agreementResult.error);
    }
  }
  
  if (selectedDocTypes.indexOf('media release') !== -1 && shouldIncludeMediaRelease(studentData.studentId)) {
    var mediaResult = generateDocumentForStudent(studentData, billingData, 'media release', null, currentSemester);
    if (mediaResult.success && !mediaResult.alreadyExists) {
      documents.push({
        name: 'Media Release',
        fileId: mediaResult.fileId,
        url: mediaResult.url
      });
    } else if (!mediaResult.success) {
      errors.push('Media Release: ' + mediaResult.error);
    }
  }
  
  // Create individual PDFs (no combined packet)
  if (documents.length > 0) {
    var packetFileName = "Registration Packet - " + studentData.lastName + ", " + studentData.firstName + 
                        (packetType === 'print' ? ' (Print Version)' : ' (Email Version)');
    
    var combinedResult = UtilityScriptLibrary.combineDocumentsIntoPDF(documents, packetFileName, destinationFolder);
    
    if (combinedResult.success) {
      return {
        success: true,
        message: "Selected documents generated successfully",
        documentsIncluded: extractDocumentNames(documents),
        packetType: packetType
      };
    } else {
      return {
        success: false,
        error: "Failed to create documents: " + combinedResult.error,
        packetType: packetType
      };
    }
  } else {
    return {
      success: true,
      message: "No documents selected or all documents already exist",
      documentsIncluded: [],
      packetType: packetType
    };
  }
}

function determineIfNewStudent(studentId, currentSemester) {
  try {
    // Get the contacts workbook to check First Enrollment Term
    var contactsSS = UtilityScriptLibrary.getWorkbook('contacts');
    var studentsSheet = contactsSS.getSheetByName('students');
    
    if (!studentsSheet) {
      UtilityScriptLibrary.debugLog("determineIfNewStudent", "ERROR", "Students sheet not found", "", "");
      return true; // Default to new student if we can't determine
    }
    
    var data = studentsSheet.getDataRange().getValues();
    var headers = data[0];
    var norm = UtilityScriptLibrary.normalizeHeader;
    
    // Find column indices
    var studentIdCol = -1;
    var enrollmentTermCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var normalizedHeader = norm(headers[i]);
      if (normalizedHeader === norm('Student ID')) {
        studentIdCol = i;
      } else if (normalizedHeader === norm('First Enrollment Term')) {
        enrollmentTermCol = i;
      }
    }
    
    if (studentIdCol === -1 || enrollmentTermCol === -1) {
      UtilityScriptLibrary.debugLog("determineIfNewStudent", "ERROR", "Required columns not found", 
                    "StudentID col: " + studentIdCol + ", Enrollment col: " + enrollmentTermCol, "");
      return true; // Default to new student
    }
    
    // Find the student's first enrollment term
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol] === studentId) {
        var firstEnrollmentTerm = data[i][enrollmentTermCol];
        var isNew = (firstEnrollmentTerm === currentSemester);
        
        UtilityScriptLibrary.debugLog("determineIfNewStudent", "DEBUG", "Student enrollment check", 
                      "First term: " + firstEnrollmentTerm + ", Current: " + currentSemester + ", Is new: " + isNew, "");
        
        return isNew;
      }
    }
    
    // Student not found, assume new
    UtilityScriptLibrary.debugLog("determineIfNewStudent", "WARN", "Student not found in contacts", "ID: " + studentId, "");
    return true;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("determineIfNewStudent", "ERROR", "Error determining student status", 
                  "", error.message);
    return true; // Default to new student on error
  }
}

function showSimpleDocumentSelectionDialog() {
  try {
    // Define the standard document types that are always available
    var availableDocuments = [
      {
        name: 'Welcome Letter',
        type: 'welcome letter',
        description: 'Introductory letter with program information'
      },
      {
        name: 'Invoice',
        type: 'invoice', 
        description: 'Billing statement with payment details'
      },
      {
        name: 'Agreement',
        type: 'agreement',
        description: 'Terms and conditions that need to be signed'
      },
      {
        name: 'Media Release',
        type: 'media release',
        description: 'Permission form for using photos/videos'
      }
    ];
    
    var html = HtmlService.createTemplate(getDocumentSelectionHtml());
    html.documentsNeededJson = Utilities.jsonStringify(availableDocuments);
    
    var htmlOutput = html.evaluate()
      .setWidth(500)
      .setHeight(400)
      .setTitle('Select Documents to Generate');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Document Selection');
    
    return { showingDialog: true };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("showSimpleDocumentSelectionDialog", "ERROR", "Error showing dialog", "", error.message);
    
    // Fallback: generate all document types
    var allDocTypes = ['welcome letter', 'invoice', 'agreement', 'media release'];
    
    PropertiesService.getScriptProperties().setProperties({
      'documentSelection': JSON.stringify(allDocTypes),
      'selectionTimestamp': new Date().getTime().toString()
    });
    
    continuePacketGeneration();
    
    return { 
      generateAll: true, 
      selectedTypes: allDocTypes
    };
  }
}

function determinePacketVersions(deliveryPreference) {
  UtilityScriptLibrary.debugLog("determinePacketVersions", "DEBUG", "Input delivery preference", 
                "Original: '" + deliveryPreference + "'", "");
  
  var preference = deliveryPreference ? deliveryPreference.toLowerCase() : 'email';
  
  UtilityScriptLibrary.debugLog("determinePacketVersions", "DEBUG", "Processed preference", 
                "Lowercase: '" + preference + "'", "");
  
  var result;
  if (preference.indexOf('both') !== -1) {
    result = ['print', 'email'];
    UtilityScriptLibrary.debugLog("determinePacketVersions", "DEBUG", "Selected both path", 
                  "Result: " + JSON.stringify(result), "");
  } else if (preference === 'mail' || preference === 'print') {
    result = ['print'];
    UtilityScriptLibrary.debugLog("determinePacketVersions", "DEBUG", "Selected print path", 
                  "mail check: " + (preference === 'mail') + ", print check: " + (preference === 'print'), "");
  } else {
    result = ['email'];
    UtilityScriptLibrary.debugLog("determinePacketVersions", "DEBUG", "Selected email path (default)", 
                  "Result: " + JSON.stringify(result), "");
  }
  
  return result;
}

function extractDocumentNames(documents) {
  var names = [];
  for (var i = 0; i < documents.length; i++) {
    names.push(documents[i].name);
  }
  return names;
}

function extractRosterDataForAttendance(rosterSheet) {
  var data = rosterSheet.getDataRange().getValues();
  var headers = data[0];
  
  // Dynamic column finder
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
    
    // Find lessons remaining column
    var lessonsRemainingCol = getCol('Lessons Remaining');
    var lessonsRemaining = 0;
    
    if (lessonsRemainingCol !== -1 && row[lessonsRemainingCol]) {
      lessonsRemaining = parseFloat(row[lessonsRemainingCol]) || 0;
    }
    
    var student = {
      id: row[getCol('Student ID')] || '',
      lastName: row[lastNameCol] || '',
      firstName: row[firstNameCol] || '',
      instrument: row[getCol('Instrument')] || '',
      lessonLength: row[getCol('Length')] || 30,
      lessonsRegistered: 0,
      lessonsCompleted: 0,
      lessonsRemaining: lessonsRemaining,
      status: row[getCol('Status')] || 'active'
    };
    
    // Only include active students (not dropped)
    if (student.status.toString().toLowerCase() !== 'dropped') {
      students.push(student);
    }
  }
  
  UtilityScriptLibrary.debugLog("extractRosterDataForAttendance", "INFO", 
                                "Extracted roster data", 
                                students.length + " students found", "");
  return students;
}

function processTeacherForNewAttendance(teacherName, rosterUrl, targetMonthName) {
  try {
    UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "INFO", 
                                  "Processing teacher", teacherName, "");
    
    // Open teacher roster
    var rosterSS = SpreadsheetApp.openByUrl(rosterUrl);
    
    // Find the most recent roster sheet
    var rosterSheet = findMostRecentRosterSheet(rosterSS);
    
    if (!rosterSheet) {
      UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "INFO", 
                                    "No roster sheet found - skipping", 
                                    teacherName, "");
      return { skipped: true, reason: 'no_roster' };
    }
    
    var usedRosterName = rosterSheet.getName();
    UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "INFO", 
                                  "Using roster sheet", 
                                  teacherName + " - " + usedRosterName, "");
    
    // Extract roster data and filter for students with lessons remaining
    var allStudents = extractRosterDataForAttendance(rosterSheet);
    var activeStudents = [];
    
    for (var i = 0; i < allStudents.length; i++) {
      var student = allStudents[i];
      var status = (student.status || '').toString().trim();
      // Include student if they have lessons remaining > 0 AND status is Active or Carryover
      if (student.lessonsRemaining && student.lessonsRemaining > 0 && 
          (status === 'Active' || status === 'Carryover')) {
        activeStudents.push(student);
      }
    }
    
    if (activeStudents.length === 0) {
      UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "INFO", 
                                    "No active/carryover students with lessons remaining - skipping", 
                                    teacherName, "");
      return { skipped: true, reason: 'no_active_students' };
    }
    
    // Alphabetize students by last name, then first name
    activeStudents.sort(function(a, b) {
      var lastNameCompare = (a.lastName || '').localeCompare(b.lastName || '');
      if (lastNameCompare !== 0) {
        return lastNameCompare;
      }
      return (a.firstName || '').localeCompare(b.firstName || '');
    });
    
    // Check if target month sheet already exists
    var existingSheet = rosterSS.getSheetByName(targetMonthName);
    
    if (existingSheet) {
      // Update existing sheet by adding missing students
      var studentsAdded = addMissingStudentsToAttendanceSheet(existingSheet, activeStudents);
      
      UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "SUCCESS", 
                                    "Updated attendance sheet", 
                                    teacherName + " - " + targetMonthName + " (added " + studentsAdded + " students)", "");
      
      return { updated: true, studentsAdded: studentsAdded };
      
    } else {
      // Create new attendance sheet using Utility function
      UtilityScriptLibrary.createMonthlyAttendanceSheet(rosterSS, targetMonthName, activeStudents);
      
      UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "SUCCESS", 
                                    "Created attendance sheet", 
                                    teacherName + " - " + targetMonthName + " (" + activeStudents.length + " students)", "");
      
      return { created: true, studentCount: activeStudents.length };
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("processTeacherForNewAttendance", "ERROR", 
                                  "Error processing teacher", 
                                  teacherName, error.message);
    throw error;
  }
}

function generateDocumentForStudent(studentData, billingData, templateType, deliveryMethod, currentSemester, billingSheetName, billingSheet, headerMap) {
  try {
    var documentTypeDisplay = templateType.charAt(0).toUpperCase() + templateType.slice(1);
    
    UtilityScriptLibrary.debugLog("generateDocumentForStudent", "INFO", "Generating " + documentTypeDisplay, 
                  "Student: " + studentData.firstName + " " + studentData.lastName + ", Type: " + templateType, "");
    
    var templateKey = selectDocumentTemplate(templateType, studentData, deliveryMethod, currentSemester, billingData);
    var variables = buildTemplateVariables(studentData, billingData, templateType);
    var fileName = buildDocumentFileName(studentData, billingData, templateType, deliveryMethod);
    
    UtilityScriptLibrary.debugLog("generateDocumentForStudent", "DEBUG", "Getting destination folder", 
                  "Billing sheet name: " + billingSheetName, "");
    
    var destinationFolder = getStudentDocumentsFolder(billingSheetName);
    
    UtilityScriptLibrary.debugLog("generateDocumentForStudent", "DEBUG", "Template variables built", 
                  "Variables count: " + Object.keys(variables).length + ", Template: " + templateKey, "");
    
    if (UtilityScriptLibrary.documentAlreadyExists(fileName, destinationFolder)) {
      return {
        success: true,
        message: documentTypeDisplay + " already exists",
        fileName: fileName,
        alreadyExists: true,
        templateUsed: templateKey
      };
    }
    
    UtilityScriptLibrary.debugLog("generateDocumentForStudent", "DEBUG", "Calling generateDocumentFromTemplate", 
                  "Template: " + templateKey + ", File: " + fileName, "");
    
    var result = UtilityScriptLibrary.generateDocumentFromTemplate(
      templateKey,
      variables,
      fileName,
      destinationFolder
    );
    
    UtilityScriptLibrary.debugLog("generateDocumentForStudent", "DEBUG", "generateDocumentFromTemplate returned", 
                  "Success: " + result.success + ", Has error: " + (result.error ? "Yes" : "No"), 
                  result.error || "");
    
    if (result.success) {
      UtilityScriptLibrary.debugLog("generateDocumentForStudent", "SUCCESS", documentTypeDisplay + " generated", 
                    "File: " + fileName, "");
      
      updateDocIdInBillingSheet(studentData.studentId, templateType, result.fileId, result.url, billingSheet, headerMap);
      
      return {
        success: true,
        fileId: result.fileId,
        url: result.url,
        fileName: fileName,
        templateUsed: templateKey
      };
    } else {
      UtilityScriptLibrary.debugLog("generateDocumentForStudent", "ERROR", "Document generation failed", 
                    "Template: " + templateKey + ", File: " + fileName, 
                    result.error || "No error message returned");
      
      return {
        success: false,
        error: result.error,
        templateUsed: templateKey
      };
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateDocumentForStudent", "ERROR", "Failed to generate " + templateType, 
                  "Method: " + deliveryMethod, error.message + " | " + error.stack);
    return {
      success: false,
      error: error.message
    };
  }
}

function generateInvoicesForBillingCycle(billingSheetName, options) {
  options = options || {};
  var includeNegativeBalances = options.includeNegativeBalances || false;
  var forceRegenerate = options.forceRegenerate || false;
  var balanceFilter = options.balanceFilter || 'positive'; // 'positive', 'negative', 'all'
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = billingSheetName ? 
      ss.getSheetByName(billingSheetName) : 
      ss.getCurrentBillingSheet();
    
    if (!billingSheet) {
      throw new Error('Billing sheet not found: ' + billingSheetName);
    }
    
    UtilityScriptLibrary.debugLog("generateInvoicesForBillingCycle", "INFO", "Starting batch invoice generation", 
                  "Sheet: " + billingSheet.getName(), "");
    
    var data = billingSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    
    var results = {
      generated: 0,
      refunds: 0,
      skipped: 0,
      errors: 0,
      details: []
    };
    
    // Process each student row (skip header)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentData = extractStudentDataFromBillingRow(row, headerMap);
      
      // Skip rows without student data
      if (!studentData.firstName || !studentData.lastName) {
        continue;
      }
      
      // Apply balance filter
      var currentBalance = parseFloat(row[headerMap[UtilityScriptLibrary.normalizeHeader("Current Balance")] - 1]) || 0;
      var shouldProcess = false;
      
      if (balanceFilter === 'all') {
        shouldProcess = currentBalance !== 0;
      } else if (balanceFilter === 'positive') {
        shouldProcess = currentBalance > 0;
      } else if (balanceFilter === 'negative') {
        shouldProcess = currentBalance < 0;
      }
      
      if (!shouldProcess) {
        continue;
      }
      
      // Generate invoice
      var invoiceOptions = {
        includeNegativeBalances: includeNegativeBalances || (balanceFilter === 'negative'),
        forceRegenerate: forceRegenerate
      };
      
      var result = generateInvoiceForStudent(studentData, row, headerMap, invoiceOptions);
      
      if (result.success) {
        if (result.skipped || result.alreadyExists) {
          results.skipped++;
        } else if (result.isRefund) {
          results.refunds++;
          // Update Invoice URL column in billing sheet
          updateInvoiceUrlInBillingSheet(billingSheet, i + 1, headerMap, result.url);
        } else {
          results.generated++;
          // Update Invoice URL column in billing sheet
          updateInvoiceUrlInBillingSheet(billingSheet, i + 1, headerMap, result.url);
        }
        
        results.details.push({
          student: studentData.firstName + ' ' + studentData.lastName,
          status: result.alreadyExists ? 'existing' : (result.isRefund ? 'refund' : 'generated'),
          balance: result.balance || currentBalance
        });
      } else {
        results.errors++;
        results.details.push({
          student: studentData.firstName + ' ' + studentData.lastName,
          status: 'error',
          error: result.error
        });
      }
    }
    
    UtilityScriptLibrary.debugLog("generateInvoicesForBillingCycle", "INFO", "Batch generation completed", 
                  "Generated: " + results.generated + ", Refunds: " + results.refunds + ", Errors: " + results.errors, "");
    
    return results;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateInvoicesForBillingCycle", "ERROR", "Batch generation failed", 
                  "", error.message + " | " + error.stack);
    throw error;
  }
}

function generateRefundInvoicesForBillingCycle(billingSheetName) {
  return generateInvoicesForBillingCycle(billingSheetName, {
    balanceFilter: 'negative',
    includeNegativeBalances: true
  });
}

function generateRegistrationPacketForStudentWithSelection(studentData, billingData, selectedDocumentTypes, deliveryMethods, currentSemester) {
  try {
    UtilityScriptLibrary.debugLog("generateRegistrationPacketForStudentWithSelection", "INFO", "Starting registration packet generation", 
                  "Student: " + studentData.firstName + " " + studentData.lastName, "");
    
    var results = {
      student: studentData.firstName + " " + studentData.lastName,
      documents: [],
      errors: []
    };
    
    if (!deliveryMethods || deliveryMethods.length === 0) {
      deliveryMethods = ['email'];
    }
    
    for (var i = 0; i < selectedDocumentTypes.length; i++) {
      var docType = selectedDocumentTypes[i];
      
      for (var j = 0; j < deliveryMethods.length; j++) {
        var deliveryMethod = deliveryMethods[j];
        
        try {
          var documentResult = generateDocumentForStudent(
            studentData, 
            billingData, 
            docType, 
            deliveryMethod, 
            currentSemester
          );
          
          if (documentResult.success) {
            results.documents.push({
              type: docType,
              method: deliveryMethod,
              fileName: documentResult.fileName,
              url: documentResult.url,
              alreadyExists: documentResult.alreadyExists || false,
              template: documentResult.templateUsed
            });
          } else {
            results.errors.push({
              type: docType,
              method: deliveryMethod,
              error: documentResult.error,
              template: documentResult.templateUsed
            });
          }
          
        } catch (docError) {
          UtilityScriptLibrary.debugLog("generateRegistrationPacketForStudentWithSelection", "ERROR", "Error generating document", 
                        "Type: " + docType + ", Method: " + deliveryMethod, docError.message);
          results.errors.push({
            type: docType,
            method: deliveryMethod,
            error: docError.message
          });
        }
      }
    }
    
    return results;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateRegistrationPacketForStudentWithSelection", "ERROR", "Error in packet generation", 
                  "Student: " + (studentData.firstName + " " + studentData.lastName), error.message);
    throw error;
  }
}

function generateRegistrationPacketsForBillingCycle() {
  try {
    var billingSheet = SpreadsheetApp.getActiveSheet();
    var billingSheetName = billingSheet.getName();
    
    UtilityScriptLibrary.debugLog("generateRegistrationPacketsForBillingCycle", "INFO", "Starting batch packet generation", 
                  "Sheet: " + billingSheetName, "");
    
    // Store the billing sheet name for later use
    PropertiesService.getScriptProperties().setProperty('currentBillingSheet', billingSheetName);
    
    // Show document selection dialog
    showSimpleDocumentSelectionDialog();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("generateRegistrationPacketsForBillingCycle", "ERROR", "Failed to start generation", "", error.message);
    SpreadsheetApp.getUi().alert('Error starting packet generation: ' + error.message);
  }
}

function getDocIdColumnName(docType) {
  var columnMap = {
    'welcome letter': 'Welcome Letter ID',
    'invoice': 'Invoice ID',
    'agreement': 'Agreement ID',
    'media release': 'Media Release ID'
  };
  return columnMap[docType] || '';
}

function getDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap) {
  var columnName = getDocIdColumnName(docType);
  var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
  var docIdCol = headerMap[UtilityScriptLibrary.normalizeHeader(columnName)];
  
  if (!studentIdCol || !docIdCol) {
    return null;
  }
  
  var data = billingSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][studentIdCol - 1] === studentId) {
      return data[i][docIdCol - 1] || null;
    }
  }
  return null;
}

function getDocumentSelectionHtml() {
  return '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<base target="_top">' +
    '<style>' +
    'body { font-family: Arial, sans-serif; padding: 20px; margin: 0; }' +
    '.header { color: #1f4e79; margin-bottom: 20px; border-bottom: 2px solid #1f4e79; padding-bottom: 10px; }' +
    '.document-list { margin: 20px 0; }' +
    '.document-item { margin: 15px 0; padding: 10px; border: 1px solid #ddd; border-radius: 5px; background-color: #f9f9f9; }' +
    '.document-item label { display: flex; align-items: center; cursor: pointer; font-weight: bold; }' +
    '.document-item input[type="checkbox"] { margin-right: 10px; transform: scale(1.2); }' +
    '.document-description { margin-top: 5px; font-size: 0.9em; color: #666; margin-left: 25px; }' +
    '.button-container { text-align: right; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; }' +
    'button { padding: 10px 20px; margin-left: 10px; border: none; border-radius: 5px; cursor: pointer; font-size: 14px; }' +
    '.btn-secondary { background-color: #6c757d; color: white; }' +
    '.btn-success { background-color: #28a745; color: white; }' +
    'button:hover { opacity: 0.8; }' +
    'button:disabled { opacity: 0.5; cursor: not-allowed; }' +
    '.select-all-container { margin: 15px 0; padding: 10px; background-color: #e9ecef; border-radius: 5px; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '<div class="header"><h2>Select Documents to Generate</h2></div>' +
    '<p>Based on billing records, the following documents are needed. Select which ones to generate:</p>' +
    '<div class="select-all-container"><label><input type="checkbox" id="selectAll"> <strong>Select All</strong></label></div>' +
    '<div class="document-list" id="documentList"></div>' +
    '<div class="button-container">' +
    '<button type="button" class="btn-secondary" onclick="cancelGeneration()">Cancel</button>' +
    '<button type="button" class="btn-success" onclick="generateSelected()">Generate Selected</button>' +
    '</div>' +
    '<script>' +
    'console.log("Script starting - ES5 version");' +
    'var documentsNeededJson = "<?= documentsNeededJson ?>";' +
    'console.log("Raw JSON string:", documentsNeededJson);' +
    'var documentsNeeded = [];' +
    'try {' +
    'documentsNeeded = JSON.parse(documentsNeededJson);' +
    'console.log("Successfully parsed documents:", documentsNeeded);' +
    '} catch (e) {' +
    'console.error("JSON parse error:", e);' +
    'console.error("Raw JSON was:", documentsNeededJson);' +
    '}' +
    'var documentDescriptions = {' +
    '"welcome letter": "Introductory letter with program information",' +
    '"invoice": "Billing statement with payment details",' +
    '"agreement": "Terms and conditions that need to be signed",' +
    '"media release": "Permission form for using photos/videos"' +
    '};' +
    'function populateDocuments() {' +
    'console.log("populateDocuments() called, documents array:", documentsNeeded);' +
    'var container = document.getElementById("documentList");' +
    'if (!container) { console.error("Container not found"); return; }' +
    'if (!documentsNeeded || !Array.isArray(documentsNeeded)) { console.error("documentsNeeded is not valid array:", documentsNeeded); return; }' +
    'console.log("About to create", documentsNeeded.length, "checkboxes");' +
    'for (var i = 0; i < documentsNeeded.length; i++) {' +
    'var doc = documentsNeeded[i];' +
    'console.log("Processing document", i, ":", doc);' +
    'var div = document.createElement("div");' +
    'div.className = "document-item";' +
    'var description = documentDescriptions[doc.type] || "Required document";' +
    'div.innerHTML = "<label><input type=\\"checkbox\\" id=\\"doc_" + i + "\\" value=\\"" + doc.type + "\\" checked>" + doc.name + "</label><div class=\\"document-description\\">" + description + "</div>";' +
    'container.appendChild(div);' +
    'console.log("Added checkbox for:", doc.name);' +
    '}' +
    'console.log("Finished populating documents");' +
    '}' +
    'document.getElementById("selectAll").addEventListener("change", function() {' +
    'var checkboxes = document.querySelectorAll(".document-item input[type=\\"checkbox\\"]");' +
    'for (var i = 0; i < checkboxes.length; i++) {' +
    'checkboxes[i].checked = this.checked;' +
    '}' +
    '});' +
    'function generateSelected() {' +
    'var checkboxes = document.querySelectorAll(".document-item input[type=\\"checkbox\\"]");' +
    'var selectedTypes = [];' +
    'for (var i = 0; i < checkboxes.length; i++) {' +
    'if (checkboxes[i].checked) {' +
    'selectedTypes.push(checkboxes[i].value);' +
    '}' +
    '}' +
    'if (selectedTypes.length === 0) {' +
    'alert("Please select at least one document type to generate.");' +
    'return;' +
    '}' +
    'var button = document.querySelector(".btn-success");' +
    'button.disabled = true;' +
    'button.textContent = "Generating...";' +
    'console.log("Calling processDocumentSelection with:", selectedTypes);' +
    'google.script.run' +
    '.withSuccessHandler(function(response) {' +
    'console.log("Generation completed successfully", response);' +
    'google.script.host.close();' +
    '})' +
    '.withFailureHandler(function(error) {' +
    'console.error("Generation error:", error);' +
    'button.disabled = false;' +
    'button.textContent = "Generate Selected";' +
    'alert("Error: " + error.message);' +
    '})' +
    '.processDocumentSelection(selectedTypes);' +
    '}' +
    'function cancelGeneration() {' +
    'google.script.host.close();' +
    '}' +
    'console.log("About to call populateDocuments");' +
    'populateDocuments();' +
    'console.log("populateDocuments completed");' +
    '</script>' +
    '</body>' +
    '</html>';
}

function processDocumentSelection(selectedTypes) {
  try {
    UtilityScriptLibrary.debugLog("processDocumentSelection", "INFO", "Document selection received", 
                  "Selected: " + selectedTypes.join(', '), "");
    
    // Get and store the current billing sheet name
    var billingSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var billingSheetName = billingSheet.getName();
    
    UtilityScriptLibrary.debugLog("processDocumentSelection", "DEBUG", "Storing billing sheet name", 
                  "Name: '" + billingSheetName + "'", "");
    
    // Store both the selection and sheet name for the background process
    PropertiesService.getScriptProperties().setProperties({
      'selectedDocTypes': JSON.stringify(selectedTypes),
      'currentBillingSheet': billingSheetName
    });
    
    // Create a trigger to run the actual processing in 1 second
    ScriptApp.newTrigger('executeDocumentGeneration')
      .timeBased()
      .after(1000)
      .create();
    
    return { success: true };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("processDocumentSelection", "ERROR", "Error processing selection", "", error.message);
    throw error;
  }
}

function executeDocumentGeneration() {
  try {
    var properties = PropertiesService.getScriptProperties();
    var selectedTypesJson = properties.getProperty('selectedDocTypes');
    var billingSheetName = properties.getProperty('currentBillingSheet');
    
    UtilityScriptLibrary.debugLog("executeDocumentGeneration", "DEBUG", "Retrieved billing sheet name", 
                  "Name: '" + billingSheetName + "'", "");
    
    if (!selectedTypesJson || !billingSheetName) {
      throw new Error('Missing generation data in properties');
    }
    
    var selectedTypes = JSON.parse(selectedTypesJson);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getSheetByName(billingSheetName);
    
    if (!billingSheet) {
      throw new Error('Billing sheet not found: ' + billingSheetName);
    }
    
    UtilityScriptLibrary.debugLog("executeDocumentGeneration", "DEBUG", "Calling getCurrentSemesterFromBillingMetadata", 
                  "With sheet name: '" + billingSheetName + "'", "");
    
    var currentSemester = getCurrentSemesterFromBillingMetadata(billingSheetName);
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    var data = billingSheet.getDataRange().getValues();
    
    var results = {
      generated: 0,
      skipped: 0,
      errors: 0
    };
    
    var studentIdColIndex = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')] - 1;
    
    for (var i = 1; i < data.length; i++) {
      try {
        var studentId = data[i][studentIdColIndex];
        if (!studentId) {
          continue;
        }
        
        var needsAnyDocs = false;
        for (var j = 0; j < selectedTypes.length; j++) {
          var existingDocId = getDocIdFromBillingSheet(studentId, selectedTypes[j], billingSheet, headerMap);
          if (!existingDocId) {
            needsAnyDocs = true;
            break;
          }
        }
        
        if (!needsAnyDocs) {
          results.skipped++;
          continue;
        }
        
        var studentData = extractStudentDataFromBillingRow(data[i], headerMap);
        var billingData = extractBillingDataFromRow(data[i], headerMap, billingSheet);
        
        if (!studentData || !studentData.studentId) {
          continue;
        }
        
        var deliveryMethods = determinePacketVersions(billingData.deliveryPreference);
        
        // Pre-check: Determine if any core documents (invoice, agreement, media release) are needed
        var needsInvoice = shouldGenerateInvoice(billingData);
        var needsAgreement = shouldIncludeAgreement(studentData.studentId, billingSheet, headerMap);
        var needsMediaRelease = shouldIncludeMediaRelease(studentData.studentId, billingSheet, headerMap);
        var needsAnyCoreDoc = needsInvoice || needsAgreement || needsMediaRelease;
        
        for (var j = 0; j < selectedTypes.length; j++) {
          var docType = selectedTypes[j];
          var existingDocId = getDocIdFromBillingSheet(studentData.studentId, docType, billingSheet, headerMap);
          
          if (!existingDocId) {
            // Skip invoice generation if there are no line items
            if (docType === 'invoice' && !needsInvoice) {
              UtilityScriptLibrary.debugLog("executeDocumentGeneration", "INFO", "Skipping invoice - no line items", 
                            "Student: " + studentData.firstName + " " + studentData.lastName + 
                            ", Balance: " + billingData.currentBalance + 
                            ", Programs: " + (billingData.programTotals ? billingData.programTotals.programs.length : 0), "");
              results.skipped++;
              continue;
            }
            
            // Skip agreement if student already has one on file
            if (docType === 'agreement' && !needsAgreement) {
              UtilityScriptLibrary.debugLog("executeDocumentGeneration", "INFO", "Skipping agreement - already on file", 
                            "Student: " + studentData.firstName + " " + studentData.lastName, "");
              results.skipped++;
              continue;
            }
            
            // Skip media release if student already has one on file
            if (docType === 'media release' && !needsMediaRelease) {
              UtilityScriptLibrary.debugLog("executeDocumentGeneration", "INFO", "Skipping media release - already on file", 
                            "Student: " + studentData.firstName + " " + studentData.lastName, "");
              results.skipped++;
              continue;
            }
            
            // Skip welcome letter if no core documents are needed
            if (docType === 'welcome letter' && !needsAnyCoreDoc) {
              UtilityScriptLibrary.debugLog("executeDocumentGeneration", "INFO", "Skipping welcome letter - no core documents needed", 
                            "Student: " + studentData.firstName + " " + studentData.lastName + 
                            ", Invoice: " + needsInvoice + ", Agreement: " + needsAgreement + ", Media: " + needsMediaRelease, "");
              results.skipped++;
              continue;
            }
            
            for (var k = 0; k < deliveryMethods.length; k++) {
              var deliveryMethod = deliveryMethods[k];
              
              var generateResult = generateDocumentForStudent(
                studentData,
                billingData,
                docType,
                deliveryMethod,
                currentSemester,
                billingSheetName,
                billingSheet,
                headerMap
              );
              
              if (generateResult.success && !generateResult.alreadyExists) {
                results.generated++;
                UtilityScriptLibrary.debugLog("executeDocumentGeneration", "SUCCESS", "Generated document", 
                              "Student: " + studentData.firstName + " " + studentData.lastName + ", Type: " + docType, "");
              } else if (generateResult.alreadyExists) {
                results.skipped++;
              } else {
                results.errors++;
                UtilityScriptLibrary.debugLog("executeDocumentGeneration", "ERROR", "Failed to generate", 
                              "Student: " + studentData.firstName + " " + studentData.lastName + ", Type: " + docType, 
                              generateResult.error || "Unknown error");
              }
            }
          }
        }
        
      } catch (studentError) {
        results.errors++;
        UtilityScriptLibrary.debugLog("executeDocumentGeneration", "ERROR", "Error processing student row " + (i + 1), 
                      "", studentError.message);
      }
    }
    
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'executeDocumentGeneration') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    properties.deleteProperty('selectedDocTypes');
    properties.deleteProperty('currentBillingSheet');
    
    SpreadsheetApp.getActive().toast(
      'Generated: ' + results.generated + ' | Skipped: ' + results.skipped + ' | Errors: ' + results.errors,
      'Generation Complete',
      10
    );
    
    UtilityScriptLibrary.debugLog("executeDocumentGeneration", "INFO", "Document generation completed", 
                  "Generated: " + results.generated + ", Skipped: " + results.skipped + ", Errors: " + results.errors);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("executeDocumentGeneration", "ERROR", "Error executing generation", "", error.message);
    SpreadsheetApp.getActive().toast('Error: ' + error.message, 'Generation Failed', 5);
  }
}

function selectDocumentTemplate(templateType, studentData, deliveryMethod, currentSemester, billingData) {
  if (templateType === 'welcome letter') {
    var letterType = billingData && billingData.letterType ? billingData.letterType.toLowerCase().trim() : '';
    
    // If letterType is 'revised', use it exactly - no auto-detection
    if (letterType === 'revised') {
      var suffix = deliveryMethod === 'print' ? 'Print' : 'Email';
      return 'revisedInvoice' + suffix;
    }
    
    // For all other cases, verify conditions and auto-detect
    // Check if missing document letter should be used
    if (shouldUseMissingDocumentLetter(billingData)) {
      letterType = 'missing';
      billingData.letterType = 'missing';  // FIX: Update billingData so buildTemplateVariables knows
      UtilityScriptLibrary.debugLog('selectDocumentTemplate', 'INFO', 'Auto-detected missing document condition', 
                'Student: ' + studentData.studentId, '');
    }
    
    // Now apply the determined letter type
    if (letterType === 'second') {
      var suffix = deliveryMethod === 'print' ? 'Print' : 'Email';
      return 'secondInvoice' + suffix;
    } else if (letterType === 'missing') {
      var suffix = deliveryMethod === 'print' ? 'Print' : 'Email';
      return 'missingDocument' + suffix;
    }
    
    // Handle welcome vs returning letter types
    var isNewStudent = (letterType === 'welcome');
    var isReturningStudent = (letterType === 'returning');
    
    // If no letter type set, fall back to checking enrollment
    if (!letterType || (letterType !== 'welcome' && letterType !== 'returning')) {
      isNewStudent = determineIfNewStudent(studentData.studentId, currentSemester);
      isReturningStudent = !isNewStudent;
    }
    
    var isAdult = UtilityScriptLibrary.determineIfStudentIsAdult(studentData);
    var baseKey;
    
    if (isNewStudent) {
      baseKey = isAdult ? 'newAdult' : 'newFamily';
    } else if (isReturningStudent) {
      baseKey = isAdult ? 'returningAdult' : 'returningFamily';
    } else {
      // Fallback
      baseKey = isAdult ? 'newAdult' : 'newFamily';
    }
    
    var suffix = deliveryMethod === 'print' ? 'Print' : 'Email';
    return baseKey + suffix;
  } else if (templateType === 'agreement') {
    return 'agreement';
  } else if (templateType === 'media release') {
    var isAdult = UtilityScriptLibrary.determineIfStudentIsAdult(studentData);
    return isAdult ? 'mediaReleaseAdult' : 'mediaReleaseChild';
  } else if (templateType === 'invoice') {
    return 'invoice';
  }
  
  throw new Error('Unknown template type: ' + templateType);
}

function shouldGenerateInvoice(billingData) {
  try {
    // Check if current balance is positive (amount owed)
    var balance = parseFloat(billingData.currentBalance) || 0;
    var hasPositiveBalance = balance > 0;
    
    // Check if there are any programs with quantity > 0
    var hasPrograms = false;
    if (billingData.programTotals && billingData.programTotals.programs) {
      for (var i = 0; i < billingData.programTotals.programs.length; i++) {
        var program = billingData.programTotals.programs[i];
        var quantity = parseFloat(program.quantity) || 0;
        if (quantity > 0) {
          hasPrograms = true;
          break;
        }
      }
    }
    
    var shouldGenerate = hasPositiveBalance || hasPrograms;
    
    UtilityScriptLibrary.debugLog("shouldGenerateInvoice", "DEBUG", "Invoice generation check", 
                  "Balance: " + balance + ", Has programs: " + hasPrograms + ", Result: " + shouldGenerate, "");
    
    return shouldGenerate;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("shouldGenerateInvoice", "ERROR", "Error checking invoice generation", 
                  "", error.message);
    return false;
  }
}

function shouldIncludeAgreement(studentId, billingSheet, headerMap) {
  try {
    var agreementFormCol = headerMap[UtilityScriptLibrary.normalizeHeader('Agreement Form')];
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    
    if (!agreementFormCol || !studentIdCol) {
      UtilityScriptLibrary.debugLog("shouldIncludeAgreement", "WARNING", "Required columns not found", 
                    "AgreementForm col: " + agreementFormCol + ", StudentID col: " + studentIdCol, "");
      return true; // Default to needing agreement if we can't check
    }
    
    var data = billingSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol - 1] === studentId) {
        var hasAgreement = data[i][agreementFormCol - 1];
        // If checkbox is TRUE, they have it. If FALSE or empty, they need it.
        var needsAgreement = !hasAgreement || hasAgreement === false;
        
        UtilityScriptLibrary.debugLog("shouldIncludeAgreement", "DEBUG", "Agreement check", 
                      "Student: " + studentId + ", HasAgreement: " + hasAgreement + ", Needs: " + needsAgreement, "");
        
        return needsAgreement;
      }
    }
    
    UtilityScriptLibrary.debugLog("shouldIncludeAgreement", "WARNING", "Student not found in billing sheet", 
                  "Student ID: " + studentId, "");
    return true; // Student not found, assume needs agreement
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("shouldIncludeAgreement", "ERROR", "Error checking agreement status", 
                  "", error.message);
    return true; // Default to needing agreement on error
  }
}

function shouldIncludeMediaRelease(studentId, billingSheet, headerMap) {
  try {
    var mediaReleaseCol = headerMap[UtilityScriptLibrary.normalizeHeader('Media Release')];
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    
    if (!mediaReleaseCol || !studentIdCol) {
      UtilityScriptLibrary.debugLog("shouldIncludeMediaRelease", "WARNING", "Required columns not found", 
                    "MediaRelease col: " + mediaReleaseCol + ", StudentID col: " + studentIdCol, "");
      return true; // Default to needing media release if we can't check
    }
    
    var data = billingSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol - 1] === studentId) {
        var hasMediaRelease = data[i][mediaReleaseCol - 1];
        // If checkbox is TRUE, they have it. If FALSE or empty, they need it.
        var needsMediaRelease = !hasMediaRelease || hasMediaRelease === false;
        
        UtilityScriptLibrary.debugLog("shouldIncludeMediaRelease", "DEBUG", "Media release check", 
                      "Student: " + studentId + ", HasRelease: " + hasMediaRelease + ", Needs: " + needsMediaRelease, "");
        
        return needsMediaRelease;
      }
    }
    
    UtilityScriptLibrary.debugLog("shouldIncludeMediaRelease", "WARNING", "Student not found in billing sheet", 
                  "Student ID: " + studentId, "");
    return true; // Student not found, assume needs media release
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("shouldIncludeMediaRelease", "ERROR", "Error checking media release status", 
                  "", error.message);
    return true; // Default to needing media release on error
  }
}

function shouldIncludeDocument(studentId, docType, billingSheet, headerMap) {
  // Check if student still needs to return this document
  var stillNeeded = checkIfDocumentStillNeeded(studentId, docType, billingSheet, headerMap);
  if (!stillNeeded) return false;
  
  // Check if we already generated one this cycle (has doc ID)
  var existingDocId = getDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap);
  if (existingDocId) {
    try {
      DriveApp.getFileById(existingDocId);
      return false; // Document exists, don't regenerate
    } catch (e) {
      // Document was deleted, clear the ID and regenerate
      clearDocIdFromBillingSheet(studentId, docType, billingSheet, headerMap);
    }
  }
  
  return true; // Generate new document
}

function shouldUseMissingDocumentLetter(billingData) {
  try {
    // Check if invoice is needed (balance > 0 or has programs)
    var invoiceNeeded = shouldGenerateInvoice(billingData);
    
    // If invoice is needed, this is NOT a missing document situation
    if (invoiceNeeded) {
      UtilityScriptLibrary.debugLog('shouldUseMissingDocumentLetter', 'DEBUG', 'Invoice still needed, not missing doc situation', 
                    'Balance: ' + billingData.currentBalance, '');
      return false;
    }
    
    // Invoice not needed (balance = 0), check if any forms are missing
    var missingAgreement = !billingData.agreementForm || billingData.agreementForm === false;
    var missingMediaRelease = !billingData.mediaRelease || billingData.mediaRelease === false;
    
    var hasMissingForms = missingAgreement || missingMediaRelease;
    
    UtilityScriptLibrary.debugLog('shouldUseMissingDocumentLetter', 'DEBUG', 'Checking missing doc conditions', 
                  'Invoice needed: false, Missing agreement: ' + missingAgreement + ', Missing media: ' + missingMediaRelease, '');
    
    return hasMissingForms;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('shouldUseMissingDocumentLetter', 'ERROR', 'Error checking missing document conditions', 
                                  '', error.message);
    return false;
  }
}

function updateDocIdInBillingSheet(studentId, docType, fileId, docUrl, billingSheet, headerMap) {
  try {
    var columnName = getDocIdColumnName(docType);
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var docIdCol = headerMap[UtilityScriptLibrary.normalizeHeader(columnName)];
    
    if (!studentIdCol || !docIdCol) {
      UtilityScriptLibrary.debugLog("updateDocIdInBillingSheet", "ERROR", "Columns not found", 
                    "StudentID col: " + studentIdCol + ", DocID col: " + docIdCol + ", ColumnName: " + columnName, "");
      return;
    }
    
    var data = billingSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][studentIdCol - 1] === studentId) {
        billingSheet.getRange(i + 1, docIdCol).setFormula('=HYPERLINK("' + docUrl + '", "' + fileId + '")');
        UtilityScriptLibrary.debugLog("updateDocIdInBillingSheet", "INFO", "Updated doc ID", 
                      "Student: " + studentId + ", DocType: " + docType + ", FileID: " + fileId, "");
        return;
      }
    }
    
    UtilityScriptLibrary.debugLog("updateDocIdInBillingSheet", "ERROR", "Student not found", 
                  "StudentID: " + studentId, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("updateDocIdInBillingSheet", "ERROR", "Error updating doc ID", 
                  "StudentID: " + studentId + ", DocType: " + docType, error.message);
  }
}

function updateInvoiceUrlInBillingSheet(billingSheet, rowNumber, headerMap, url) {
  var norm = UtilityScriptLibrary.normalizeHeader;
  var urlCol = headerMap[norm("Invoice URL")];
  
  if (urlCol && url) {
    billingSheet.getRange(rowNumber, urlCol).setValue(url);
  }
}

// ============================================================================
// SECTION 5: RECONCILIATION
// ============================================================================

function applyAdminVisualFormatting(billingSheet) {
  try {
    UtilityScriptLibrary.debugLog('🎨 Applying visual formatting to student names...');
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    
    // Get column indices - UPDATED to use new column names
    var lastNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student Last Name')];
    var firstNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student First Name')];
    var currentBalanceCol = headerMap[UtilityScriptLibrary.normalizeHeader('Current Balance')];
    var hoursRemainingCol = headerMap[UtilityScriptLibrary.normalizeHeader('Hours Remaining')];
    var lessonLengthCol = headerMap[UtilityScriptLibrary.normalizeHeader('Lesson Length')]; // ADDED: Need lesson length
    var agreementCol = headerMap[UtilityScriptLibrary.normalizeHeader('Agreement Form')];
    var mediaCol = headerMap[UtilityScriptLibrary.normalizeHeader('Media Release')];
    
    if (!lastNameCol || !firstNameCol || !currentBalanceCol || !hoursRemainingCol || 
        !lessonLengthCol || !agreementCol || !mediaCol) {
      throw new Error('Required columns not found for visual formatting');
    }
    
    // STEP 1: RESET ALL NAME CELLS TO DEFAULT FORMATTING
    var nameRange = billingSheet.getRange(2, lastNameCol, billingSheet.getLastRow() - 1, 2);
    nameRange.clearFormat();
    
    // Apply default formatting to all name cells
    nameRange.setBackground('#ffffff')      // White background (default)
             .setFontColor('#000000')       // Black text (default)
             .setFontWeight('normal');      // Normal weight (default)
    
    UtilityScriptLibrary.debugLog('✅ Reset all name cells to default formatting');
    
    // Define color values
    var colors = {
      greenBackground: '#e8f5e8',  // Light green for credits
      redBackground: '#ffebee',    // Light red for debts
      orangeText: '#ff9800',       // Orange for low lessons
      redText: '#f44336',          // Red for no lessons
      redBorder: '#f44336',        // Red border for missing agreement
      blueBorder: '#2196f3',       // Blue border for missing media
      purpleBorder: '#9c27b0'      // Purple border for missing both
    };
    
    // Get data for processing
    var dataRange = billingSheet.getRange(2, 1, billingSheet.getLastRow() - 1, billingSheet.getLastColumn());
    var data = dataRange.getValues();
    
    // STEP 2: APPLY CONDITIONAL FORMATTING TO STUDENTS WHO NEED IT
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var currentBalance = parseFloat(row[currentBalanceCol - 1]) || 0;
      var hoursRemaining = parseFloat(row[hoursRemainingCol - 1]) || 0;
      var lessonLength = parseFloat(row[lessonLengthCol - 1]) || 30; // Default to 30 minutes
      var hasAgreement = row[agreementCol - 1] === '✅';
      var hasMedia = row[mediaCol - 1] === '✅';
      
      // FIXED: Convert hours back to lessons for warning logic
      var lessonsRemaining = lessonLength > 0 ? hoursRemaining / (lessonLength / 60) : 0;
      
      var nameRange = billingSheet.getRange(i + 2, lastNameCol, 1, 2);
      
      // Background color based on balance
      var backgroundColor = currentBalance >= 0 ? colors.greenBackground : colors.redBackground;
      nameRange.setBackground(backgroundColor);
      
      // Text color based on lessons remaining
      var textColor = '#000000'; // Default black
      if (lessonsRemaining <= 0) {
        textColor = colors.redText;
      } else if (lessonsRemaining <= 3) {
        textColor = colors.orangeText;
      }
      nameRange.setFontColor(textColor);
      
      // Border color based on form status
      var borderColor = null;
      var borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
      
      if (!hasAgreement && !hasMedia) {
        borderColor = colors.purpleBorder; // Purple for both missing
      } else if (!hasAgreement) {
        borderColor = colors.redBorder;    // Red for missing agreement
      } else if (!hasMedia) {
        borderColor = colors.blueBorder;   // Blue for missing media
      }
      
      if (borderColor) {
        nameRange.setBorder(true, true, true, true, true, true, borderColor, borderStyle);
      }
    }
    
    UtilityScriptLibrary.debugLog('✅ Applied conditional formatting to ' + data.length + ' students');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error in applyAdminVisualFormatting: ' + error.message);
    throw error;
  }
}

function applyWarningsToTeacherWorkbook(teacherSS, warningStudents, targetDate) {
  var sheets = teacherSS.getSheets();
  var formattedDate = UtilityScriptLibrary.formatDateFlexible(targetDate, "MM/dd");
  
  // Create lookup map for warning students with all their data
  var warningStudentMap = {};
  for (var i = 0; i < warningStudents.length; i++) {
    var student = warningStudents[i];
    warningStudentMap[student.studentId] = {
      firstName: student.firstName,
      lastName: student.lastName,
      instrument: student.instrument,
      lessonLength: student.lessonLength,
      lessonsRemaining: student.lessonsRemaining
    };
  }
  
  // Process each attendance sheet (monthly sheets)
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Only process attendance sheets (month names)
    if (!UtilityScriptLibrary.isMonthSheet(sheetName)) {
      continue;
    }
    
    // Check if this sheet is current or future month
    if (!UtilityScriptLibrary.isCurrentOrFutureMonth(sheetName, targetDate)) {
      UtilityScriptLibrary.debugLog("applyWarningsToTeacherWorkbook", "DEBUG", 
                                    "Skipping past month sheet", sheetName, "");
      continue;
    }
    
    updateSheetStudentWarnings(sheet, warningStudentMap, formattedDate);
  }
  
  UtilityScriptLibrary.debugLog("applyWarningsToTeacherWorkbook", "SUCCESS", 
                                "Applied warnings to current/future months", 
                                warningStudents.length + " warning students", "");
}

function expandSheetAttendanceRows(sheet) {
  try {
    var sheetName = sheet.getName();
    UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'INFO', 'Analyzing sheet for row expansion', 'Sheet: ' + sheetName, '');
    
    // OPTIMIZATION: Read all data once instead of cell-by-cell
    var dataRange = sheet.getDataRange();
    var allData = dataRange.getValues();
    var numRows = allData.length;
    
    if (numRows < 2) {
      UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'INFO', 'Sheet has insufficient rows', 'Rows: ' + numRows, '');
      return false;
    }
    
    // Find column indices from header row
    var headers = allData[0];
    var studentIdCol = -1;
    var dateCol = -1;
    var adminCommentsCol = -1;
    var lengthCol = -1;
    var studentNameCol = -1;
    
    for (var col = 0; col < headers.length; col++) {
      var header = UtilityScriptLibrary.normalizeHeader(String(headers[col]));
      if (header === 'studentid') studentIdCol = col;
      else if (header === 'date') dateCol = col;
      else if (header === 'admincomments') adminCommentsCol = col;
      else if (header === 'length') lengthCol = col;
      else if (header === 'studentname') studentNameCol = col;
    }
    
    if (studentIdCol === -1 || dateCol === -1) {
      UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'ERROR', 'Required columns not found', '', '');
      return false;
    }
    
    // Get month index from sheet name to calculate days in month
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var monthIndex = -1;
    var sheetNameLower = sheetName.toLowerCase();
    for (var m = 0; m < monthNames.length; m++) {
      if (monthNames[m].toLowerCase() === sheetNameLower) {
        monthIndex = m;
        break;
      }
    }
    
    // Track student data
    var studentData = {};
    var currentStudent = null;
    
    for (var i = 1; i < numRows; i++) {
      var row = allData[i];
      var studentId = row[studentIdCol];
      var dateValue = row[dateCol];
      var adminComment = row[adminCommentsCol];
      var lengthValue = row[lengthCol];
      var studentName = row[studentNameCol];
      
      var hasAdminComment = adminComment && String(adminComment).trim() !== '';
      
      // Check if this is a student header row (has " - " in student name)
      if (studentId && studentName && String(studentName).indexOf(' - ') !== -1) {
        currentStudent = String(studentId).trim();
        
        // Extract numeric lesson length from header (strips " minutes" suffix)
        var numericLessonLength = extractNumericLessonLength(lengthValue);
        
        var nameParts = String(studentName).split(' - ');
        var fullName = nameParts[0] ? String(nameParts[0]).trim() : '';
        var nameSplit = fullName.split(',');
        var lastName = nameSplit[0] ? String(nameSplit[0]).trim() : '';
        var firstName = nameSplit[1] ? String(nameSplit[1]).trim() : '';
        
        if (!studentData[currentStudent]) {
          studentData[currentStudent] = {
            studentId: currentStudent,
            firstName: firstName,
            lastName: lastName,
            lessonLength: numericLessonLength,
            headerRowIndex: i,
            lastRowIndex: i,
            emptyRows: 0
          };
        }
      } else if (currentStudent && studentId === currentStudent) {
        // This is a lesson row for current student
        studentData[currentStudent].lastRowIndex = i;
        
        // Count as empty if no date AND no admin comment
        if ((!dateValue || String(dateValue).trim() === '') && !hasAdminComment) {
          studentData[currentStudent].emptyRows++;
        }
      }
    }
    
    // Determine which students need additional rows
    var studentsNeedingRows = [];
    
    for (var studentId in studentData) {
      if (studentData.hasOwnProperty(studentId)) {
        var student = studentData[studentId];
        
        // Only add rows if student has fewer than 3 empty rows
        if (student.emptyRows < 3) {
          var rowsToAdd = 3 - student.emptyRows;
          
          // Cap at 2 rows maximum
          if (rowsToAdd > 2) {
            rowsToAdd = 2;
          }
          
          // Cap rows based on days remaining in month
          if (monthIndex !== -1) {
            var today = new Date();
            var currentYear = today.getFullYear();
            var currentMonth = today.getMonth();
            
            // If this sheet is for current month, calculate days remaining
            if (monthIndex === currentMonth) {
              var lastDayOfMonth = new Date(currentYear, currentMonth + 1, 0).getDate();
              var currentDay = today.getDate();
              var daysRemaining = lastDayOfMonth - currentDay;
              
              // Cap rowsToAdd to daysRemaining
              if (rowsToAdd > daysRemaining) {
                UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'DEBUG', 'Capping rows by days remaining', 
                              'Student: ' + studentId + ', Was: ' + rowsToAdd + ', Now: ' + daysRemaining, '');
                rowsToAdd = daysRemaining;
              }
            }
          }
          
          if (rowsToAdd > 0) {
            student.rowsToAdd = rowsToAdd;
            studentsNeedingRows.push(student);
            UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'DEBUG', 'Student needs rows', 
                          'Student: ' + studentId + ', Has: ' + student.emptyRows + ', Adding: ' + rowsToAdd, '');
          }
        }
      }
    }
    
    if (studentsNeedingRows.length === 0) {
      UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'INFO', 'Row expansion complete', 'No rows needed', '');
      return false;
    }
    
    // Sort students by lastRowIndex descending (work backwards to maintain indices)
    studentsNeedingRows.sort(function(a, b) {
      return b.lastRowIndex - a.lastRowIndex;
    });
    
    var totalRowsAdded = 0;
    
    // Add rows for each student
    for (var i = 0; i < studentsNeedingRows.length; i++) {
      var student = studentsNeedingRows[i];
      var insertAfterRow = student.lastRowIndex + 1; // Convert to 1-based
      
      // FIXED: Check the background color of the last existing row to determine starting color
      var lastRowBackground = sheet.getRange(insertAfterRow, 1).getBackground();
      var startWithDark;
      
      // Determine which color the last row has, then start with the opposite
      if (lastRowBackground.toLowerCase() === UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background.toLowerCase()) {
        startWithDark = false; // Last row was dark, so start new rows with light
      } else {
        startWithDark = true; // Last row was light (or white), so start new rows with dark
      }
      
      // Insert the rows
      sheet.insertRowsAfter(insertAfterRow, student.rowsToAdd);
      
      // Format the new rows to match lesson rows
      for (var j = 0; j < student.rowsToAdd; j++) {
        var newRowNum = insertAfterRow + j + 1;
        
        // Create lesson row data (11 columns to match createLessonRows)
        var lessonData = [
          student.studentId,                                          // A - Student ID
          student.lastName + ', ' + student.firstName,                // B - Student Name
          '',                                                         // C - Date (empty)
          student.lessonLength,                                       // D - Length (numeric only, no " minutes")
          '',                                                         // E - Status (empty)
          '',                                                         // F - Comments (empty)
          '',                                                         // G - Admin Review Date (empty)
          '',                                                         // H - Invoice Date (empty)
          '',                                                         // I - Payment Date (empty)
          '',                                                         // J - Invoice Number (empty)
          ''                                                          // K - Admin Comments (empty)
        ];
        
        sheet.getRange(newRowNum, 1, 1, lessonData.length).setValues([lessonData]);
        
        // FIXED: Apply alternating row colors based on previous row's color
        var shouldBeDark = (j % 2 === 0) ? startWithDark : !startWithDark;
        var backgroundColor = shouldBeDark ? 
          UtilityScriptLibrary.STYLES.ALTERNATING_DARK.background : 
          UtilityScriptLibrary.STYLES.ALTERNATING_LIGHT.background;
        sheet.getRange(newRowNum, 1, 1, lessonData.length).setBackground(backgroundColor);
        
        // FIXED: Add green border using STYLES constant
        sheet.getRange(newRowNum, 6).setBorder(null, null, null, true, null, null, 
          UtilityScriptLibrary.STYLES.HEADER.background, SpreadsheetApp.BorderStyle.SOLID_THICK);
        
        // Apply status dropdown to column E
        var statusOptions = ['Lesson', 'No Show', 'No Lesson'];
        var statusRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(statusOptions)
          .setAllowInvalid(false)
          .build();
        sheet.getRange(newRowNum, 5).setDataValidation(statusRule);
        
        // Apply date format to column C
        sheet.getRange(newRowNum, 3).setNumberFormat('ddd, M/d');
      }
      
      totalRowsAdded += student.rowsToAdd;
    }
    
    UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'SUCCESS', 'Added rows for ' + studentsNeedingRows.length + ' students', 
                  'Total rows added: ' + totalRowsAdded, '');
    
    return studentsNeedingRows.length > 0;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('expandSheetAttendanceRows', 'ERROR', 'Failed to expand rows', 
                  'Sheet: ' + sheet.getName(), error.message);
    throw error;
  }
}

function extractNumericLessonLength(lengthValue) {
  if (!lengthValue) return 30; // default
  var strValue = String(lengthValue).trim();
  // Extract the number from strings like "30 minutes" or just "30"
  var match = strValue.match(/^(\d+)/);
  return match ? parseInt(match[1], 10) : 30;
}

function findBillingRowByStudentId(billingData, studentId, studentIdColIndex) {
  try {
    if (!billingData || billingData.length < 2 || studentIdColIndex === -1) {
      return null;
    }
    
    // Search through billing data (skip header row at index 0)
    for (var i = 1; i < billingData.length; i++) {
      var row = billingData[i];
      var rowStudentId = row[studentIdColIndex];
      
      if (rowStudentId === studentId) {
        UtilityScriptLibrary.debugLog('findBillingRowByStudentId', 'SUCCESS', 'Student found in billing', 
                     'Student ID: ' + studentId + ', Row: ' + i, '');
        return i;
      }
    }
    
    UtilityScriptLibrary.debugLog('findBillingRowByStudentId', 'WARNING', 'Student not found in billing sheet', 
                 'Student ID: ' + studentId, '');
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('findBillingRowByStudentId', 'ERROR', 'Search failed', 
                 'Student ID: ' + studentId, error.message);
    return null;
  }
}

function findMostRecentRosterSheet(workbook) {
  try {
    // Get all sheets in workbook
    var allSheets = workbook.getSheets();
    var rosterSheets = [];
    
    // Filter for roster sheets and extract seasons
    for (var i = 0; i < allSheets.length; i++) {
      var sheetName = allSheets[i].getName();
      if (sheetName.indexOf(' Roster') !== -1) {
        var season = sheetName.replace(' Roster', '').trim();
        rosterSheets.push({
          sheet: allSheets[i],
          season: season
        });
      }
    }
    
    if (rosterSheets.length === 0) {
      UtilityScriptLibrary.debugLog("findMostRecentRosterSheet", "INFO", 
                                    "No roster sheets found in workbook", "", "");
      return null;
    }
    
    // Get semester metadata to determine chronological order
    var semesterMetadataSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      UtilityScriptLibrary.debugLog("findMostRecentRosterSheet", "WARNING", 
                                    "Semester Metadata sheet not found - using first roster", "", "");
      return rosterSheets[0].sheet;
    }
    
    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find columns
    var nameCol = -1, startCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      }
    }
    
    if (nameCol === -1 || startCol === -1) {
      UtilityScriptLibrary.debugLog("findMostRecentRosterSheet", "WARNING", 
                                    "Required columns not found in Semester Metadata - using first roster", "", "");
      return rosterSheets[0].sheet;
    }
    
    // Build array of semesters with their seasons and start dates
    var semesters = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[startCol]) {
        var semesterName = row[nameCol].toString().trim();
        var season = UtilityScriptLibrary.extractSeasonFromSemester(semesterName);
        if (season) {
          semesters.push({
            season: season,
            semesterName: semesterName,
            startDate: new Date(row[startCol])
          });
        }
      }
    }
    
    // Sort semesters by start date (most recent first)
    semesters.sort(function(a, b) {
      return b.startDate - a.startDate;
    });
    
    // Find the most recent semester that has a matching roster sheet
    for (var i = 0; i < semesters.length; i++) {
      var semester = semesters[i];
      for (var j = 0; j < rosterSheets.length; j++) {
        if (rosterSheets[j].season === semester.season) {
          UtilityScriptLibrary.debugLog("findMostRecentRosterSheet", "SUCCESS", 
                                        "Found most recent roster", 
                                        semester.season + " Roster (" + semester.semesterName + ")", "");
          return rosterSheets[j].sheet;
        }
      }
    }
    
    // Fallback: if no match found, use first roster sheet
    UtilityScriptLibrary.debugLog("findMostRecentRosterSheet", "WARNING", 
                                  "No semester match found - using first available roster", 
                                  rosterSheets[0].season + " Roster", "");
    return rosterSheets[0].sheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("findMostRecentRosterSheet", "ERROR", 
                                  "Error finding roster sheet", "", error.message);
    return null;
  }
}

function findStudentInContacts(contactsData, studentIdCol, targetStudentId) {
  for (var i = 1; i < contactsData.length; i++) {
    if (contactsData[i][studentIdCol - 1] === targetStudentId) {
      return i;
    }
  }
  return -1;
}

function generateReconciliationSummary(results) {
  var lines = [];
  
  // Attendance summary
  if (results.attendance) {
    lines.push('📚 ATTENDANCE: Lessons processed');
  } else if (results.errors.some(function(e) { return e.indexOf('Attendance') !== -1; })) {
    lines.push('❌ ATTENDANCE: Failed to process');
  }
  
  // Payment summary
  if (results.payments) {
    lines.push('💰 PAYMENTS: ' + results.payments.processed + ' payments processed');
    if (results.payments.errors > 0) {
      lines.push('   ⚠️ ' + results.payments.errors + ' payment errors');
    }
  } else if (results.errors.some(function(e) { return e.indexOf('Payment') !== -1; })) {
    lines.push('❌ PAYMENTS: Failed to process');
  }
  
  // Forms summary
  if (results.forms) {
    var formsTotal = results.forms.agreementUpdates + results.forms.mediaUpdates;
    lines.push('📋 FORMS: ' + formsTotal + ' forms updated');
    lines.push('   • ' + results.forms.agreementUpdates + ' agreements');
    lines.push('   • ' + results.forms.mediaUpdates + ' media releases');
    if (results.forms.errors > 0) {
      lines.push('   ⚠️ ' + results.forms.errors + ' form errors');
    }
  } else if (results.errors.some(function(e) { return e.indexOf('Forms') !== -1; })) {
    lines.push('❌ FORMS: Failed to process');
  }
  
  // Error summary
  if (results.errors.length > 0) {
    lines.push('');
    lines.push('ERRORS:');
    for (var i = 0; i < Math.min(results.errors.length, 3); i++) {
      lines.push('• ' + results.errors[i]);
    }
    if (results.errors.length > 3) {
      lines.push('• ... and ' + (results.errors.length - 3) + ' more errors');
    }
  }
  
  return lines.join('\n');
}

function generateReconciliationSummaryUpdated(results) {
  var lines = [];
  
  // Attendance summary
  if (results.attendance) {
    lines.push('📚 ATTENDANCE: Lessons processed');
  } else if (results.errors.some(function(e) { return e.indexOf('Attendance') !== -1; })) {
    lines.push('❌ ATTENDANCE: Failed to process');
  }
  
  // Combined payment and forms summary
  if (results.combined) {
    lines.push('💰 PAYMENTS: ' + results.combined.paymentsProcessed + ' payments processed');
    if (results.combined.paymentsErrors > 0) {
      lines.push('   ⚠️ ' + results.combined.paymentsErrors + ' payment errors');
    }
    
    var formsTotal = results.combined.agreementUpdates + results.combined.mediaUpdates;
    lines.push('📋 FORMS: ' + formsTotal + ' forms updated');
    lines.push('   • ' + results.combined.agreementUpdates + ' agreements');
    lines.push('   • ' + results.combined.mediaUpdates + ' media releases');
    if (results.combined.formsErrors > 0) {
      lines.push('   ⚠️ ' + results.combined.formsErrors + ' form errors');
    }
  } else if (results.errors.some(function(e) { return e.indexOf('Payment') !== -1 || e.indexOf('Forms') !== -1; })) {
    lines.push('❌ PAYMENTS/FORMS: Failed to process');
  }
  
  // Error summary
  if (results.errors.length > 0) {
    lines.push('');
    lines.push('ERRORS:');
    for (var i = 0; i < Math.min(results.errors.length, 3); i++) {
      lines.push('• ' + results.errors[i]);
    }
    if (results.errors.length > 3) {
      lines.push('• ... and ' + (results.errors.length - 3) + ' more errors');
    }
  }
  
  return lines.join('\n');
}

function getBillingSheet(paymentDate, activeSheetName, shouldLog) {
  var monthNames = UtilityScriptLibrary.getMonthNames();
  
  if (shouldLog) {
    UtilityScriptLibrary.debugLog("getBillingSheet", "INFO", "Looking for billing sheet", 
                 "paymentDate: " + paymentDate + ", activeSheetName: " + activeSheetName, 
                 "shouldLog: " + (shouldLog ? "YES" : "NO"), "");
  }
  
  var billingSS = UtilityScriptLibrary.getWorkbook('billing');
  var spreadsheetList = billingSS.getSheetByName("Billing Metadata");
  if (!spreadsheetList) {
    UtilityScriptLibrary.debugLog("getBillingSheet", "ERROR", "'Billing Metadata' sheet not found in billing spreadsheet", "", "");
    return null;
  }

  var data = spreadsheetList.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var billingMonth = row[0];
    
    // Check if it's a Date object (Google Sheets sometimes returns dates that fail instanceof)
    if (Object.prototype.toString.call(billingMonth) === '[object Date]' || typeof billingMonth.getMonth === 'function') {
      billingMonth = monthNames[billingMonth.getMonth()] + " " + billingMonth.getFullYear();
      UtilityScriptLibrary.debugLog("getBillingSheet", "DEBUG", "Converted Date to string", "Result: " + billingMonth, "");
    }
    
    var paymentStartDate = new Date(row[2]);
    var paymentEndDate = new Date(row[3]);
    var semesterName = row[6];
    
    // Debug: Show what we're comparing
    UtilityScriptLibrary.debugLog("getBillingSheet", "DEBUG", "Checking row", 
      "paymentDate: " + paymentDate + " in range [" + paymentStartDate + " to " + paymentEndDate + "]? " + 
      (paymentDate >= paymentStartDate && paymentDate <= paymentEndDate) + 
      ", semesterName: '" + semesterName + "' === '" + activeSheetName + "'? " + (semesterName === activeSheetName), "");
    
    if (paymentDate >= paymentStartDate && paymentDate <= paymentEndDate && 
        semesterName === activeSheetName) {
      
      UtilityScriptLibrary.debugLog("getBillingSheet", "DEBUG", "Date match found, looking for sheet", 
                   "Sheet name to find: " + billingMonth, "");
      
      var billingSheet = billingSS.getSheetByName(billingMonth);
      
      if (!billingSheet) {
        UtilityScriptLibrary.debugLog("getBillingSheet", "WARNING", "Billing sheet not found", 
                     "Sheet name: " + billingMonth, "");
        continue;
      }
      
      return {
        sheet: billingSheet,
        startDate: paymentStartDate,
        endDate: paymentEndDate,
        billingMonth: billingMonth
      };
    }
  }
  
  UtilityScriptLibrary.debugLog("getBillingSheet", "WARNING", "No billing metadata match", 
               "Date: " + paymentDate + ", Payment sheet: " + activeSheetName, "");
  return null;
}

function getInvoiceNumber(billingSheet, billingRowIndex) {
  // 🧾 Dynamically get invoice number, with fallback to "Past Invoice Number"
  var headers = billingSheet.getDataRange().getValues()[0];
  var invoiceColIndex = headers.indexOf("Invoice Number");
  var pastInvoiceColIndex = headers.indexOf("Past Invoice Number");

  if (invoiceColIndex === -1 && pastInvoiceColIndex === -1) {
    UtilityScriptLibrary.debugLog("❌ Neither 'Invoice Number' nor 'Past Invoice Number' columns found. Returning null.");
    return null;
  }

  var rowNumber = billingRowIndex + 1;
  var invoiceValue = null;

  if (invoiceColIndex !== -1) {
    invoiceValue = billingSheet.getRange(rowNumber, invoiceColIndex + 1).getValue();
  }

  if (!invoiceValue && pastInvoiceColIndex !== -1) {
    invoiceValue = billingSheet.getRange(rowNumber, pastInvoiceColIndex + 1).getValue();
  }

  return invoiceValue || null;
}


function getStudentBalancesFromBilling(billingSheet, teacherName) {
  var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
  var dataRange = billingSheet.getRange(2, 1, billingSheet.getLastRow() - 1, billingSheet.getLastColumn());
  var data = dataRange.getValues();
  
  // Get column indices dynamically
  var norm = UtilityScriptLibrary.normalizeHeader;
  var studentIdCol = headerMap[norm('Student ID')] - 1;
  var teacherCol = headerMap[norm('Teacher')] - 1;
  var hoursRemainingCol = headerMap[norm('Hours Remaining')] - 1;
  var lessonLengthCol = headerMap[norm('Lesson Length')] - 1;
  var firstNameCol = headerMap[norm('Student First Name')] - 1;
  var lastNameCol = headerMap[norm('Student Last Name')] - 1;
  var instrumentCol = headerMap[norm('Instrument')] - 1;
  
  if (studentIdCol === undefined || teacherCol === undefined || hoursRemainingCol === undefined || 
      lessonLengthCol === undefined || firstNameCol === undefined || lastNameCol === undefined || 
      instrumentCol === undefined) {
    throw new Error('Required billing columns not found for reading balances');
  }
  
  var studentBalances = [];
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowTeacher = row[teacherCol];
    
    // Only get students for THIS teacher
    if (rowTeacher === teacherName) {
      var studentId = row[studentIdCol];
      var hoursRemaining = parseFloat(row[hoursRemainingCol]) || 0;
      var lessonLength = parseFloat(row[lessonLengthCol]) || 30;
      var firstName = row[firstNameCol] || '';
      var lastName = row[lastNameCol] || '';
      var instrument = row[instrumentCol] || '';
      
      // Calculate lessons remaining
      var lessonsRemaining = lessonLength > 0 ? Math.floor((hoursRemaining * 60) / lessonLength) : 0;
      
      if (studentId) {
        studentBalances.push({
          studentId: studentId,
          firstName: firstName,
          lastName: lastName,
          instrument: instrument,
          hoursRemaining: hoursRemaining,
          lessonsRemaining: lessonsRemaining,
          lessonLength: lessonLength
        });
      }
    }
  }
  
  return studentBalances;
}

function identifyWarningStudents(studentBalances) {
  var warningStudents = [];
  
  for (var i = 0; i < studentBalances.length; i++) {
    var student = studentBalances[i];
    
    // Students with 3 or fewer lessons remaining get warnings
    if (student.lessonsRemaining <= 3) {
      warningStudents.push({
        studentId: student.studentId,
        firstName: student.firstName,
        lastName: student.lastName,
        instrument: student.instrument,
        lessonLength: student.lessonLength,
        lessonsRemaining: student.lessonsRemaining
      });
    }
  }
  
  return warningStudents;
}

function locateStudentRecord(rowData, paymentSheet, billingInfo) {
  var norm = UtilityScriptLibrary.normalizeHeader;
  var headerMap = UtilityScriptLibrary.getHeaderMap(paymentSheet);
  var invoiceValue = rowData[headerMap[norm("Invoice Number")] - 1];
  if (!billingInfo || !billingInfo.sheet) return null;
  var billingSheet = billingInfo.sheet;
  var billingData = billingSheet.getDataRange().getValues();
  
  // Use the utility function for billing sheet header map (returns 1-based indices)
  var billingHeaderMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
  
  var invoiceColIndex = billingHeaderMap[norm("Invoice Number")];
  var pastInvoiceColIndex = billingHeaderMap[norm("Past Invoice Number")];
  var studentIdColIndex = billingHeaderMap[norm("Student ID")];
  var lastNameColIndex = billingHeaderMap[norm("Student Last Name")];
  var firstNameColIndex = billingHeaderMap[norm("Student First Name")];
  var instrumentColIndex = billingHeaderMap[norm("Instrument")];
  var billingRowIndex = null;
  var studentId = null;
  
  // STRATEGY 1: Try invoice number lookup
  if (invoiceValue) {
    for (var i = 1; i < billingData.length; i++) {
      var row = billingData[i];
      if ((invoiceColIndex !== undefined && row[invoiceColIndex - 1] === invoiceValue) ||
          (pastInvoiceColIndex !== undefined && row[pastInvoiceColIndex - 1] === invoiceValue)) {
        billingRowIndex = i;
        studentId = row[studentIdColIndex - 1];
        break;
      }
    }
  }
  
  // STRATEGY 2: Try name (and instrument) lookup in billing sheet
  if (billingRowIndex === null) {
    var lastName = UtilityScriptLibrary.cleanName(rowData[headerMap[norm("Student Last Name")] - 1]);
    var firstName = UtilityScriptLibrary.cleanName(rowData[headerMap[norm("Student First Name")] - 1]);
    var instrument = headerMap[norm("Instrument")] ? rowData[headerMap[norm("Instrument")] - 1] : null;
    
    // First try to match with instrument if available
    if (instrument && instrumentColIndex !== undefined) {
      for (var j = 1; j < billingData.length; j++) {
        var billingRow = billingData[j];
        var billingLastName = UtilityScriptLibrary.cleanName(billingRow[lastNameColIndex - 1]);
        var billingFirstName = UtilityScriptLibrary.cleanName(billingRow[firstNameColIndex - 1]);
        var billingInstrument = billingRow[instrumentColIndex - 1];
        
        if (billingLastName === lastName && 
            billingFirstName === firstName && 
            billingInstrument === instrument) {
          billingRowIndex = j;
          studentId = billingRow[studentIdColIndex - 1];
          break;
        }
      }
    }
    
    // If no match with instrument, try just name
    if (billingRowIndex === null) {
      for (var k = 1; k < billingData.length; k++) {
        var billingRow2 = billingData[k];
        var billingLastName2 = UtilityScriptLibrary.cleanName(billingRow2[lastNameColIndex - 1]);
        var billingFirstName2 = UtilityScriptLibrary.cleanName(billingRow2[firstNameColIndex - 1]);
        
        if (billingLastName2 === lastName && billingFirstName2 === firstName) {
          billingRowIndex = k;
          studentId = billingRow2[studentIdColIndex - 1];
          break;
        }
      }
    }
  }
  
  return billingRowIndex !== null ? {
    billingRowIndex: billingRowIndex,
    studentId: studentId
  } : null;
}

function locateStudentRecordEnhanced(rowData, billingInfo, studentIdInput, invoiceNumberInput, lastNameCol, firstNameCol) {
  try {
    var billingSheet = billingInfo.sheet;
    var billingData = billingSheet.getDataRange().getValues();
    var billingHeaders = billingData[0];
    
    var invoiceColIndex = billingHeaders.indexOf("Invoice Number");
    var pastInvoiceColIndex = billingHeaders.indexOf("Past Invoice Number");
    var studentIdColIndex = billingHeaders.indexOf("Student ID");
    var lastNameColIndex = billingHeaders.indexOf("Student Last Name");
    var firstNameColIndex = billingHeaders.indexOf("Student First Name");
    
    var billingRowIndex = null;
    var studentId = studentIdInput;
    var invoiceValue = invoiceNumberInput;
    
    // STRATEGY 1: Try invoice number lookup (if we have an invoice number)
    if (invoiceValue) {
      for (var i = 1; i < billingData.length; i++) {
        var row = billingData[i];
        if ((invoiceColIndex !== -1 && row[invoiceColIndex] === invoiceValue) ||
            (pastInvoiceColIndex !== -1 && row[pastInvoiceColIndex] === invoiceValue)) {
          billingRowIndex = i;
          studentId = row[studentIdColIndex];
          break;
        }
      }
    }
    
    // STRATEGY 2: Try student ID lookup in billing (if we have student ID but no match yet)
    if (!billingRowIndex && studentId) {
      billingRowIndex = findBillingRowByStudentId(billingData, studentId, studentIdColIndex);
      if (billingRowIndex !== null) {
        // Get invoice number from billing
        var row = billingData[billingRowIndex];
        if (invoiceColIndex !== -1 && row[invoiceColIndex]) {
          invoiceValue = row[invoiceColIndex];
        } else if (pastInvoiceColIndex !== -1 && row[pastInvoiceColIndex]) {
          invoiceValue = row[pastInvoiceColIndex];
        }
      }
    }
    
    // STRATEGY 3: Try name lookup in billing (last resort)
    if (!billingRowIndex) {
      var lastName = UtilityScriptLibrary.cleanName(rowData[lastNameCol - 1]);
      var firstName = UtilityScriptLibrary.cleanName(rowData[firstNameCol - 1]);
      
      // Search billing data by name
      for (var i = 1; i < billingData.length; i++) {
        var row = billingData[i];
        var rowLastName = UtilityScriptLibrary.cleanName(row[lastNameColIndex]);
        var rowFirstName = UtilityScriptLibrary.cleanName(row[firstNameColIndex]);
        
        if (rowLastName === lastName && rowFirstName === firstName) {
          billingRowIndex = i;
          studentId = row[studentIdColIndex];
          
          // Get invoice number
          if (invoiceColIndex !== -1 && row[invoiceColIndex]) {
            invoiceValue = row[invoiceColIndex];
          } else if (pastInvoiceColIndex !== -1 && row[pastInvoiceColIndex]) {
            invoiceValue = row[pastInvoiceColIndex];
          }
          break;
        }
      }
    }
    
    if (billingRowIndex === null) {
      return null;
    }
    
    return {
      billingRowIndex: billingRowIndex,
      studentId: studentId,
      invoiceNumber: invoiceValue
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('locateStudentRecordEnhanced', 'ERROR', 'Failed to locate student', '', error.message);
    return null;
  }
}

function logMysteryStudents(mysteryStudents) {
  try {
    if (mysteryStudents.length === 0) return;
    
    UtilityScriptLibrary.debugLog('🕵️ Logging mystery students...');
    
    var billingSS = SpreadsheetApp.getActiveSpreadsheet();
    var mysterySheet = billingSS.getSheetByName('Mystery Students');
    
    // Create sheet if it doesn't exist
    if (!mysterySheet) {
      mysterySheet = billingSS.insertSheet('Mystery Students');
      mysterySheet.getRange(1, 1, 1, 5).setValues([
        ['Date Found', 'Student ID', 'Student Name', 'Teacher', 'Notes']
      ]);
      mysterySheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    }
    
  for (var i = 0; i < mysteryStudents.length; i++) {
        var student = mysteryStudents[i];
        mysterySheet.appendRow([
          today,
          student.studentId || 'Unknown',
          student.studentName || 'Unknown',
          student.teacherName || 'Unknown',
          'Found in teacher attendance but not in billing - teacher needs to report new student'
        ]);
      }
      
      UtilityScriptLibrary.debugLog('âœ… Logged ' + mysteryStudents.length + ' mystery students');
      
    } catch (error) {
      UtilityScriptLibrary.debugLog('âŒ Error logging mystery students: ' + error.message);
    }
}

function processFormsData(paymentSheet, studentsSheet) {
  // Get payment data
  var paymentHeaderMap = UtilityScriptLibrary.getHeaderMap(paymentSheet);
  var paymentData = paymentSheet.getDataRange().getValues();
  
  // Get contacts data  
  var contactsHeaderMap = UtilityScriptLibrary.getHeaderMap(studentsSheet);
  var contactsData = studentsSheet.getDataRange().getValues();
  
  // Required column indices
  var paymentStudentIdCol = paymentHeaderMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
  var agreementCol = paymentHeaderMap[UtilityScriptLibrary.normalizeHeader('Agreement Form Received')];
  var mediaCol = paymentHeaderMap[UtilityScriptLibrary.normalizeHeader('Media Release Response')];
  
  var contactsStudentIdCol = contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
  var contactsAgreementCol = contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Current Agreement')];
  var contactsMediaCol = contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Current Media')];
  
  if (!paymentStudentIdCol || !agreementCol || !mediaCol) {
    throw new Error('Required payment columns not found');
  }
  
  if (!contactsStudentIdCol || !contactsAgreementCol || !contactsMediaCol) {
    throw new Error('Required contacts columns not found');
  }
  
  var results = {
    agreementUpdates: 0,
    mediaUpdates: 0,
    errors: 0,
    details: []
  };
  
  // Process each payment record (skip header)
  for (var i = 1; i < paymentData.length; i++) {
    var paymentRow = paymentData[i];
    var studentId = paymentRow[paymentStudentIdCol - 1];
    
    if (!studentId) continue;
    
    var agreementReceived = paymentRow[agreementCol - 1] === true;
    var mediaResponse = paymentRow[mediaCol - 1];
    var hasMediaResponse = (mediaResponse && String(mediaResponse).trim() !== '');
    
    // Skip if no forms data to process
    if (!agreementReceived && !hasMediaResponse) continue;
    
    try {
      // Find student in contacts
      var contactsRowIndex = findStudentInContacts(contactsData, contactsStudentIdCol, studentId);
      
      if (contactsRowIndex === -1) {
        results.errors++;
        results.details.push('Student ID not found in contacts: ' + studentId);
        continue;
      }
      
      // Update agreement if received
      if (agreementReceived) {
        studentsSheet.getRange(contactsRowIndex + 1, contactsAgreementCol).setValue(true);
        results.agreementUpdates++;
        results.details.push('Updated agreement for ' + studentId);
      }
      
      // Update media response if provided
      if (hasMediaResponse) {
        studentsSheet.getRange(contactsRowIndex + 1, contactsMediaCol).setValue(String(mediaResponse).trim());
        results.mediaUpdates++;
        results.details.push('Updated media response for ' + studentId + ': ' + mediaResponse);
      }
      
    } catch (error) {
      results.errors++;
      results.details.push('Error updating ' + studentId + ': ' + error.message);
    }
  }
  
  return results;
}

function processFormsReconciliationForRow(rowData, studentsSheet, studentIdInput, hasAgreement, hasMediaResponse, studentIdCol, agreementCol, mediaCol) {
  try {
    // Get student ID
    var studentId = studentIdInput;
    
    if (!studentId) {
      // Try to get from payment row using provided column index
      studentId = rowData[studentIdCol - 1];
    }
    
    if (!studentId) {
      return {
        success: false,
        message: 'Student ID not found'
      };
    }
    
    // NEW: Get current billing sheet to update forms there as well
    var billingSheet = getCurrentBillingSheet();
    var billingUpdated = false;
    
    if (billingSheet) {
      var billingHeaderMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
      var billingData = billingSheet.getDataRange().getValues();
      var billingStudentIdCol = billingHeaderMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
      var billingAgreementCol = billingHeaderMap[UtilityScriptLibrary.normalizeHeader('Agreement Form')];
      var billingMediaCol = billingHeaderMap[UtilityScriptLibrary.normalizeHeader('Media Release')];
      
      if (billingStudentIdCol && (billingAgreementCol || billingMediaCol)) {
        // Find student in billing sheet
        for (var i = 1; i < billingData.length; i++) {
          if (billingData[i][billingStudentIdCol - 1] === studentId) {
            // Update Agreement Form checkbox in billing sheet
            if (hasAgreement && billingAgreementCol) {
              var agreementCell = billingSheet.getRange(i + 1, billingAgreementCol);
              agreementCell.insertCheckboxes();
              agreementCell.setValue(true);
            }
            
            // Update Media Release checkbox in billing sheet
            if (hasMediaResponse && billingMediaCol) {
              var mediaCell = billingSheet.getRange(i + 1, billingMediaCol);
              mediaCell.insertCheckboxes();
              mediaCell.setValue(true);
            }
            
            billingUpdated = true;
            break;
          }
        }
      }
    }
    
    // Find student in contacts
    var contactsHeaderMap = UtilityScriptLibrary.getHeaderMap(studentsSheet);
    var contactsData = studentsSheet.getDataRange().getValues();
    
    var contactsStudentIdCol = contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var contactsAgreementCol = contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Current Agreement')];
    var contactsMediaCol = contactsHeaderMap[UtilityScriptLibrary.normalizeHeader('Current Media')];
    
    if (!contactsStudentIdCol || !contactsAgreementCol || !contactsMediaCol) {
      return {
        success: false,
        message: 'Required columns not found in Students sheet'
      };
    }
    
    var contactsRowIndex = findStudentInContacts(contactsData, contactsStudentIdCol, studentId);
    
    if (contactsRowIndex === -1) {
      return {
        success: false,
        message: 'Student ID not found in contacts: ' + studentId
      };
    }
    
    var agreementUpdated = false;
    var mediaUpdated = false;
    
    // Update agreement if received
    if (hasAgreement) {
      studentsSheet.getRange(contactsRowIndex + 1, contactsAgreementCol).setValue(true);
      agreementUpdated = true;
    }
    
    // Update media response if provided
    if (hasMediaResponse) {
      var mediaResponse = rowData[mediaCol - 1];
      studentsSheet.getRange(contactsRowIndex + 1, contactsMediaCol).setValue(String(mediaResponse).trim());
      mediaUpdated = true;
    }
    
    return {
      success: true,
      message: 'Forms updated for student ' + studentId + (billingUpdated ? ' (Billing & Contacts)' : ' (Contacts only)'),
      agreementUpdated: agreementUpdated,
      mediaUpdated: mediaUpdated
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

function processPaymentReconciliationForRow(rowData, paymentSheet, rowNumber, studentIdInput, invoiceNumberInput, amountCol, dateCol, lastNameCol, firstNameCol) {
  try {
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'INFO', 'Starting', 
                 'Row: ' + rowNumber, '');
    
    // Extract payment data
    var rawAmount = amountCol ? rowData[amountCol - 1] : undefined;
    var rawDate = dateCol ? rowData[dateCol - 1] : undefined;
    
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Raw payment data', 
                 'Amount: ' + rawAmount + ', Date: ' + rawDate, '');
    
    var amountPaid = parseFloat(String(rawAmount).replace(/[^0-9.]/g, '').trim());
    var paymentDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
    
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Parsed payment data', 
                 'Amount: ' + amountPaid + ', Date: ' + paymentDate, '');
    
    // Validate payment data
    if (isNaN(amountPaid) || isNaN(paymentDate.getTime())) {
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'WARNING', 'Invalid payment data', 
                   'Amount valid: ' + !isNaN(amountPaid) + ', Date valid: ' + !isNaN(paymentDate.getTime()), '');
      return { 
        success: false, 
        message: 'Incomplete payment data'
      };
    }
    
    // Get billing sheet for this payment date
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Getting billing sheet', 
                 'Date: ' + paymentDate + ', Payment sheet: ' + paymentSheet.getName(), '');
    
    var billingInfo = getBillingSheet(paymentDate, paymentSheet.getName(), true);
    if (!billingInfo || !billingInfo.sheet) {
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'WARNING', 'No billing sheet found', 
                   'billingInfo null: ' + (billingInfo === null), '');
      return { 
        success: false, 
        message: 'No billing sheet found for payment date'
      };
    }
    
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Found billing sheet', 
                 'Sheet: ' + billingInfo.sheet.getName() + ', Start: ' + billingInfo.startDate + 
                 ', End: ' + billingInfo.endDate, '');
    
    // Locate student record in billing
    var studentRecord = locateStudentRecordEnhanced(
      rowData, 
      billingInfo, 
      studentIdInput, 
      invoiceNumberInput,
      lastNameCol,
      firstNameCol
    );
    
    if (!studentRecord) {
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'WARNING', 'Student not found', 
                   'Could not locate student in billing sheet', '');
      return { 
        success: false, 
        message: 'Student record not found in billing'
      };
    }
    
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Found student record', 
                 'Student ID: ' + studentRecord.studentId + ', Billing row: ' + studentRecord.billingRowIndex + 
                 ', Invoice: ' + studentRecord.invoiceNumber, '');
    
    // CRITICAL FIX: Write Student ID to payment sheet BEFORE calling sumPayments
    var headerMap = UtilityScriptLibrary.getHeaderMap(paymentSheet);
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    
    if (studentIdCol && studentRecord.studentId && !rowData[studentIdCol - 1]) {
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Writing Student ID to payment sheet', 
                   'Row: ' + rowNumber + ', Student ID: ' + studentRecord.studentId, '');
      paymentSheet.getRange(rowNumber, studentIdCol).setValue(studentRecord.studentId);
      SpreadsheetApp.flush(); // Ensure write completes before reading
    }
    
    // Calculate total payments for this student
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Calculating total payments', 
                 'Student: ' + studentRecord.studentId + ', Start: ' + billingInfo.startDate + 
                 ', End: ' + billingInfo.endDate, '');
    
    var totalPayments = sumPayments(
      paymentSheet, 
      studentRecord.studentId, 
      billingInfo.startDate, 
      billingInfo.endDate
    );
    
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Total payments calculated', 
                 'Total: ' + totalPayments + ', Current payment: ' + amountPaid, '');
    
    // Get invoice number from billing if we don't have it
    var invoiceValue = studentRecord.invoiceNumber;
    
    // Update billing sheet with payment total
    var billingHeaders = billingInfo.sheet.getDataRange().getValues()[0];
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Billing sheet headers', 
                 'Headers: ' + billingHeaders.join(', '), '');
    
    var paymentReceivedIndex = billingHeaders.indexOf("Payment Received");
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Payment Received column', 
                 'Index: ' + paymentReceivedIndex + ', Found: ' + (paymentReceivedIndex !== -1), '');
    
    if (paymentReceivedIndex !== -1) {
      var targetRow = studentRecord.billingRowIndex + 1;
      var targetCol = paymentReceivedIndex + 1;
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Writing to billing sheet', 
                   'Sheet: ' + billingInfo.sheet.getName() + ', Row: ' + targetRow + 
                   ', Col: ' + targetCol + ', Value: ' + totalPayments, '');
      
      billingInfo.sheet.getRange(targetRow, targetCol).setValue(totalPayments);
      
      // Verify the write
      var verifyValue = billingInfo.sheet.getRange(targetRow, targetCol).getValue();
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'DEBUG', 'Verified billing sheet write', 
                   'Written: ' + totalPayments + ', Read back: ' + verifyValue + 
                   ', Match: ' + (verifyValue === totalPayments), '');
    } else {
      UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'WARNING', 
                   'Payment Received column not found in billing sheet', 
                   'Cannot update payment amount', '');
    }
    
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'SUCCESS', 'Payment reconciled', 
                 'Student: ' + studentRecord.studentId + ', Amount: $' + amountPaid, '');
    
    return {
      success: true,
      message: 'Payment reconciled for student ' + studentRecord.studentId + ': $' + amountPaid,
      studentId: studentRecord.studentId,
      invoiceNumber: invoiceValue
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('processPaymentReconciliationForRow', 'ERROR', 'Payment reconciliation failed', 
                 'Row: ' + rowNumber, error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

function processPaymentRecord(rowData, paymentSheet, headerMap, rowNumber) {
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Starting', 'Row: ' + rowNumber, '');
  
  var amountPaidCol = headerMap["amountpaid"];
  var dateCol = headerMap["date"];
  var reconciledCol = headerMap["reconciled"];
  var studentIdCol = headerMap["studentid"];
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Column indices', 
               'Amount Paid col: ' + amountPaidCol + ', Date col: ' + dateCol + ', Reconciled col: ' + reconciledCol + ', Student ID col: ' + studentIdCol, '');
  
  var rawAmount = amountPaidCol ? rowData[amountPaidCol - 1] : undefined;
  var rawDate = dateCol ? rowData[dateCol - 1] : undefined;
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Raw data', 
               'Amount: ' + rawAmount + ', Date: ' + rawDate + ', Reconciled col: ' + reconciledCol, '');
  
  // Skip if already reconciled
  if (reconciledCol && rowData[reconciledCol - 1] === true) {
    UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Already reconciled', 'Row: ' + rowNumber, '');
    return { processed: false, message: 'Already reconciled' };
  }
  
  var amountPaid = parseFloat(String(rawAmount).replace(/[^0-9.]/g, '').trim());
  var paymentDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Parsed data', 
               'Amount: ' + amountPaid + ', Date: ' + paymentDate + ', Valid: ' + (!isNaN(amountPaid) && !isNaN(paymentDate.getTime())), '');
  
  // Skip incomplete records
  if (isNaN(amountPaid) || isNaN(paymentDate.getTime())) {
    UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Incomplete data', 'Row: ' + rowNumber, '');
    return { processed: false, message: 'Incomplete data' };
  }
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Calling getBillingSheet', 
               'Date: ' + paymentDate + ', Sheet name: ' + paymentSheet.getName(), '');
  
  var billingInfo = getBillingSheet(paymentDate, paymentSheet.getName(), true);
  
  if (!billingInfo || !billingInfo.sheet) {
    UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'No billing sheet found', 
                 'billingInfo is null: ' + (billingInfo === null), '');
    return { processed: false, message: 'No billing sheet found' };
  }
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Found billing sheet', 
               'Sheet: ' + billingInfo.sheet.getName(), '');
  
  var studentRecord = locateStudentRecord(rowData, paymentSheet, billingInfo);
  if (!studentRecord) {
    UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Student record not found', 'Row: ' + rowNumber, '');
    return { processed: false, message: 'Student record not found' };
  }
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Found student record', 
               'Student ID: ' + studentRecord.studentId, '');
  
  // CRITICAL FIX: Write Student ID to payment sheet BEFORE calling sumPayments
  if (studentIdCol && studentRecord.studentId && !rowData[studentIdCol - 1]) {
    UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Writing Student ID to payment sheet', 
                 'Row: ' + rowNumber + ', Student ID: ' + studentRecord.studentId, '');
    paymentSheet.getRange(rowNumber, studentIdCol).setValue(studentRecord.studentId);
    // Update rowData array so sumPayments will see it
    rowData[studentIdCol - 1] = studentRecord.studentId;
  }
  
  var totalPayments = sumPayments(paymentSheet, studentRecord.studentId, billingInfo.startDate, billingInfo.endDate);
  var invoiceValue = getInvoiceNumber(billingInfo.sheet, studentRecord.billingRowIndex);
  
  reconcilePayment(paymentSheet, billingInfo.sheet, studentRecord.billingRowIndex, rowNumber, invoiceValue, totalPayments);
  
  UtilityScriptLibrary.debugLog('processPaymentRecord', 'DEBUG', 'Successfully reconciled', 
               'Student: ' + studentRecord.studentId + ', Amount: $' + amountPaid, '');
  
  return {
    processed: true,
    message: 'Reconciled payment for student ' + studentRecord.studentId + ': $' + amountPaid
  };
}

function processTeacherAttendanceForBilling(teacherSS, teacherName, targetDate, cycleStartDate) {
  // Process attendance sheets row by row: mark reconcilable lessons AND sum billable hours
  
  // DETERMINE APPROPRIATE START DATE FOR RECONCILIATION
  var effectiveStartDate;
  
  if (targetDate < cycleStartDate) {
    // Retroactive reconciliation - use start of target month
    effectiveStartDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
    UtilityScriptLibrary.debugLog("processTeacherAttendanceForBilling", "INFO", 
                                  "Retroactive reconciliation detected", 
                                  "Target: " + targetDate.toDateString() + ", Effective start: " + effectiveStartDate.toDateString() + 
                                  " (was: " + cycleStartDate.toDateString() + ")", "");
  } else {
    // Normal forward reconciliation - look back 4 months to catch late-added lessons
    effectiveStartDate = new Date(cycleStartDate);
    effectiveStartDate.setMonth(effectiveStartDate.getMonth() - 4);
    UtilityScriptLibrary.debugLog("processTeacherAttendanceForBilling", "INFO", 
                                  "Forward reconciliation with 4-month lookback", 
                                  "Cycle start: " + cycleStartDate.toDateString() + ", Effective start: " + effectiveStartDate.toDateString(), "");
  }
  
  var studentHours = {};
  var stats = {
    lessonsMarked: 0,
    lessonsCounted: 0
  };
  
  var sheets = teacherSS.getSheets();
  
  for (var sheetIdx = 0; sheetIdx < sheets.length; sheetIdx++) {
    var sheet = sheets[sheetIdx];
    var sheetName = sheet.getName();
    
    // Only process attendance sheets
    if (!UtilityScriptLibrary.isMonthSheet(sheetName)) {
      continue;
    }
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      continue;
    }
    
    var headers = data[0];
    
    // Find column indices
    var studentIdCol = -1;
    var dateCol = -1;
    var lengthCol = -1;
    var statusCol = -1;
    var adminReviewCol = -1;
    var invoiceDateCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var normalized = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (normalized === 'studentid') studentIdCol = i;
      else if (normalized === 'date') dateCol = i;
      else if (normalized === 'length') lengthCol = i;
      else if (normalized === 'status') statusCol = i;
      else if (normalized === 'adminreviewdate') adminReviewCol = i;
      else if (normalized === 'invoicedate') invoiceDateCol = i;
    }
    
    if (studentIdCol === -1 || dateCol === -1 || lengthCol === -1 || statusCol === -1 || adminReviewCol === -1) {
      UtilityScriptLibrary.debugLog("processTeacherAttendanceForBilling", "WARNING", 
                                    "Missing required columns", sheetName, "");
      continue;
    }
    
    // Process each row
    for (var rowIdx = 1; rowIdx < data.length; rowIdx++) {
      var row = data[rowIdx];
      
      var studentId = row[studentIdCol];
      var lessonDate = row[dateCol];
      var length = parseFloat(row[lengthCol]) || 0;
      var status = row[statusCol];
      var adminReviewDate = row[adminReviewCol];
      var invoiceDate = invoiceDateCol !== -1 ? row[invoiceDateCol] : '';
      
      // Skip if no student ID or invalid date
      if (!studentId || !lessonDate || !(lessonDate instanceof Date)) {
        continue;
      }
      
      // STEP 1 & 2: Check if reconcilable and mark it
      var isReconcilable = (
        (status === 'Lesson' || status === 'No Show' || status === 'No Lesson') &&
        lessonDate >= effectiveStartDate &&  // Uses effectiveStartDate (4 months back for forward reconciliation)
        lessonDate <= targetDate &&
        (!adminReviewDate || adminReviewDate === '')
      );
      
      if (isReconcilable) {
        sheet.getRange(rowIdx + 1, adminReviewCol + 1).setValue(targetDate);
        stats.lessonsMarked++;
      }
      
      // STEP 3: Check if should be summed for billing
      // After marking above, we need to check if THIS row should count
      // (either just marked OR was already marked in previous reconciliation)
      var shouldSum = (
        (adminReviewDate !== '' || isReconcilable) && // Has admin review date (or just got it)
        (!invoiceDate || invoiceDate === '') &&        // No invoice date
        status !== 'No Lesson' &&                      // Not a "No Lesson"
        length > 0                                     // Has length
      );
      
      if (shouldSum) {
        var hours = length / 60;
        if (!studentHours[studentId]) {
          studentHours[studentId] = 0;
        }
        studentHours[studentId] += hours;
        stats.lessonsCounted++;
      }
    }
  }
  
  UtilityScriptLibrary.debugLog("processTeacherAttendanceForBilling", "INFO", 
                                "Processed attendance", 
                                "Marked: " + stats.lessonsMarked + ", Counted: " + stats.lessonsCounted, "");
  
  return {
    studentHours: studentHours,
    stats: stats
  };
}

function processTeacherReconciliation(teacherName, rosterUrl, billingSheet, targetDate, billingCycleDates, currentSemester) {
  // This function processes ONE teacher completely in a single pass
  
  var stats = {
    lessonsCollected: 0,
    studentsUpdated: 0,
    warningsApplied: 0
  };
  
  // Extract file ID from roster URL
  var fileIdMatch = rosterUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!fileIdMatch) {
    throw new Error('Invalid roster URL format');
  }
  
  // STEP 1: Open teacher workbook (ONCE)
  var teacherSS = SpreadsheetApp.openById(fileIdMatch[1]);
  
  // STEP 2: Process attendance sheets (mark reconcilable lessons AND calculate billing totals)
  var attendanceResult = processTeacherAttendanceForBilling(teacherSS, teacherName, targetDate, billingCycleDates.startDate);
  stats.lessonsCollected = attendanceResult.stats.lessonsMarked;
  
  // STEP 3: Update billing sheet for THIS teacher's students
  updateBillingForTeacherStudents(billingSheet, attendanceResult.studentHours, teacherName);
  
  // STEP 4: Read back Hours Remaining from billing sheet
  var studentBalances = getStudentBalancesFromBilling(billingSheet, teacherName);
  
  // STEP 5: Update roster sheet with balances
  updateTeacherRosterBalances(teacherSS, studentBalances, currentSemester);
  stats.studentsUpdated = studentBalances.length;
  
  // STEP 6: Apply attendance warnings (pink headers for ≤3 lessons)
  var warningStudents = identifyWarningStudents(studentBalances);
  applyWarningsToTeacherWorkbook(teacherSS, warningStudents, targetDate);
  stats.warningsApplied = warningStudents.length;
  
  return stats;
}

function runCombinedReconciliation() {
  try {
    UtilityScriptLibrary.debugLog('runCombinedReconciliation', 'INFO', 'Starting combined payment and forms reconciliation', '', '');
    
    // Get payment workbook and current semester sheet
    var paymentsSS = UtilityScriptLibrary.getWorkbook('payments');
    var currentSemester = getCurrentSemesterName();
    var paymentSheet = paymentsSS.getSheetByName(currentSemester);
    
    if (!paymentSheet) {
      throw new Error('Payment sheet not found for semester: ' + currentSemester);
    }
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(paymentSheet);
    var data = paymentSheet.getDataRange().getValues();
    
    // Get column indices
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var lastNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student Last Name')];
    var firstNameCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student First Name')];
    var amountCol = headerMap[UtilityScriptLibrary.normalizeHeader('Amount Paid')];
    var dateCol = headerMap[UtilityScriptLibrary.normalizeHeader('Date')];
    var invoiceCol = headerMap[UtilityScriptLibrary.normalizeHeader('Invoice Number')];
    var agreementCol = headerMap[UtilityScriptLibrary.normalizeHeader('Agreement Form Received')];
    var mediaCol = headerMap[UtilityScriptLibrary.normalizeHeader('Media Release Response')];
    var reconciledCol = headerMap[UtilityScriptLibrary.normalizeHeader('Reconciled')];
    var commentsCol = headerMap[UtilityScriptLibrary.normalizeHeader('System Comments')];
    
    // Get contacts workbook (for forms reconciliation)
    var contactsSS = UtilityScriptLibrary.getWorkbook('contacts');
    var studentsSheet = contactsSS.getSheetByName('Students');
    
    if (!studentsSheet) {
      throw new Error('Students sheet not found in Contacts workbook');
    }
    
    var results = {
      paymentsProcessed: 0,
      paymentsErrors: 0,
      agreementUpdates: 0,
      mediaUpdates: 0,
      formsErrors: 0,
      details: []
    };
    
    // Process each payment record (skip header row)
    for (var i = 1; i < data.length; i++) {
      var rowData = data[i];
      var rowNumber = i + 1;
      
      // Skip if already reconciled
      if (reconciledCol && rowData[reconciledCol - 1] === true) {
        continue;
      }
      
      // Determine what data exists on this row
      var hasPaymentData = rowData[amountCol - 1] && rowData[dateCol - 1];
      var hasAgreement = rowData[agreementCol - 1] === true;
      var hasMediaResponse = rowData[mediaCol - 1] && String(rowData[mediaCol - 1]).trim() !== '';
      var hasFormsData = hasAgreement || hasMediaResponse;
      
      // Skip if no data to process
      if (!hasPaymentData && !hasFormsData) {
        continue;
      }
      
      var paymentSuccess = false;
      var formsSuccess = false;
      var errorMessages = [];
      var studentIdForRow = rowData[studentIdCol - 1] || null;
      var invoiceNumber = rowData[invoiceCol - 1] || null;
      
      // PROCESS PAYMENT RECONCILIATION
      if (hasPaymentData) {
        try {
          var paymentResult = processPaymentReconciliationForRow(
            rowData, 
            paymentSheet, 
            rowNumber,
            studentIdForRow,
            invoiceNumber,
            amountCol,
            dateCol,
            lastNameCol,
            firstNameCol
          );
          
          if (paymentResult.success) {
            paymentSuccess = true;
            results.paymentsProcessed++;
            results.details.push(paymentResult.message);
            
            // Update studentId and invoiceNumber if they were found during payment processing
            if (paymentResult.studentId && !studentIdForRow) {
              studentIdForRow = paymentResult.studentId;
            }
            if (paymentResult.invoiceNumber && !invoiceNumber) {
              invoiceNumber = paymentResult.invoiceNumber;
            }
          } else {
            errorMessages.push('Payment: ' + paymentResult.message);
            results.paymentsErrors++;
          }
        } catch (error) {
          errorMessages.push('Payment: ' + error.message);
          results.paymentsErrors++;
        }
      } else {
        paymentSuccess = true; // No payment data to process = success
      }
      
      // PROCESS FORMS RECONCILIATION
      if (hasFormsData) {
        try {
          var formsResult = processFormsReconciliationForRow(
            rowData,
            studentsSheet,
            studentIdForRow,
            hasAgreement,
            hasMediaResponse,
            studentIdCol,
            agreementCol,
            mediaCol
          );
          
          if (formsResult.success) {
            formsSuccess = true;
            if (formsResult.agreementUpdated) {
              results.agreementUpdates++;
            }
            if (formsResult.mediaUpdated) {
              results.mediaUpdates++;
            }
            results.details.push(formsResult.message);
          } else {
            errorMessages.push('Forms: ' + formsResult.message);
            results.formsErrors++;
          }
        } catch (error) {
          errorMessages.push('Forms: ' + error.message);
          results.formsErrors++;
        }
      } else {
        formsSuccess = true; // No forms data to process = success
      }
      
      // UPDATE PAYMENT SHEET WITH RESULTS
      // FIXED: Calculate overall success - only TRUE if no errors occurred
      var overallSuccess = (paymentSuccess || !hasPaymentData) && 
                          (formsSuccess || !hasFormsData) && 
                          errorMessages.length === 0;
      
      // Set reconciled checkbox based on overall success
      if (reconciledCol) {
        var reconciledCell = paymentSheet.getRange(rowNumber, reconciledCol);
        reconciledCell.insertCheckboxes();
        reconciledCell.setValue(overallSuccess);
      }
      
      // Write Student ID if we found it
      if (studentIdCol && studentIdForRow && !rowData[studentIdCol - 1]) {
        paymentSheet.getRange(rowNumber, studentIdCol).setValue(studentIdForRow);
      }
      
      // Write Invoice Number if we found it
      if (invoiceCol && invoiceNumber && !rowData[invoiceCol - 1]) {
        paymentSheet.getRange(rowNumber, invoiceCol).setValue(invoiceNumber);
      }
      
      // Write error messages to System Comments ONLY if reconciliation failed
      if (commentsCol) {
        if (!overallSuccess && errorMessages.length > 0) {
          paymentSheet.getRange(rowNumber, commentsCol).setValue(errorMessages.join(' | '));
        } else if (overallSuccess) {
          // Clear any old error messages on successful reconciliation
          paymentSheet.getRange(rowNumber, commentsCol).clearContent();
        }
      }
    }
    
    UtilityScriptLibrary.debugLog('runCombinedReconciliation', 'SUCCESS', 'Combined reconciliation completed', 
                 'Payments: ' + results.paymentsProcessed + ', Agreements: ' + results.agreementUpdates + 
                 ', Media: ' + results.mediaUpdates, '');
    
    return results;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('runCombinedReconciliation', 'ERROR', 'Combined reconciliation failed', '', error.message);
    throw error;
  }
}

function runWeeklyLessonReconciliation(customDate) {
  try {
    var targetDate = customDate || new Date();
    
    UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "INFO", 
                                  "Starting single-pass weekly lesson reconciliation", 
                                  "Target date: " + targetDate, "");
    
    // Get current billing sheet
    var billingSheet = getCurrentBillingSheet();
    
    if (!billingSheet) {
      throw new Error('No active billing sheet found. Please create a billing cycle first.');
    }
    
    var billingCycleDates = getCurrentBillingCycleDates();
    
    // Get current semester name for roster sheet lookup
    var currentSemester = getCurrentSemesterName();
    if (!currentSemester || currentSemester === 'Unknown Semester') {
      throw new Error('Could not determine current semester name');
    }
    
    // Get all active teachers
    var teacherLookupSS = UtilityScriptLibrary.getWorkbook('formResponses');
    var teacherLookupSheet = teacherLookupSS.getSheetByName('Teacher Roster Lookup');
    
    if (!teacherLookupSheet) {
      throw new Error('Teacher Roster Lookup sheet not found');
    }
    
    var teacherData = teacherLookupSheet.getDataRange().getValues();
    var teacherHeaders = teacherData[0];
    var displayNameCol = teacherHeaders.indexOf('Display Name');
    var rosterUrlCol = teacherHeaders.indexOf('Roster URL');
    var statusCol = teacherHeaders.indexOf('Status');
    
    if (displayNameCol === -1 || rosterUrlCol === -1 || statusCol === -1) {
      throw new Error('Required teacher columns not found in lookup sheet');
    }
    
    // Process statistics
    var stats = {
      teachersProcessed: 0,
      teachersSkipped: 0,
      teachersErrored: 0,
      totalLessons: 0,
      totalStudents: 0,
      errors: []
    };
    
    // SINGLE PASS: Process each teacher one at a time
    for (var i = 1; i < teacherData.length; i++) {
      var teacherRow = teacherData[i];
      var teacherName = teacherRow[displayNameCol];
      var rosterUrl = teacherRow[rosterUrlCol];
      var status = teacherRow[statusCol];
      
      if (status !== 'active' || !teacherName || !rosterUrl) {
        stats.teachersSkipped++;
        continue;
      }
      
      try {
        UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "INFO", 
                                      "Processing teacher", teacherName, "");
        
        var teacherStats = processTeacherReconciliation(
          teacherName,
          rosterUrl,
          billingSheet,
          targetDate,
          billingCycleDates,
          currentSemester
        );
        
        stats.teachersProcessed++;
        stats.totalLessons += teacherStats.lessonsCollected;
        stats.totalStudents += teacherStats.studentsUpdated;
        
        UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "SUCCESS", 
                                      "Completed teacher", 
                                      teacherName + " - Lessons: " + teacherStats.lessonsCollected + 
                                      ", Students: " + teacherStats.studentsUpdated, "");
        
      } catch (teacherError) {
        stats.teachersErrored++;
        stats.errors.push(teacherName + ": " + teacherError.message);
        UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "ERROR", 
                                      "Failed to process teacher", 
                                      teacherName, teacherError.message);
        // Continue with next teacher
      }
    }
    
    // Apply visual formatting to billing sheet (once at the end)
    applyAdminVisualFormatting(billingSheet);
    
    // Add lesson rows where needed (once at the end)
    expandTeacherAttendanceSheets();
    
    UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "SUCCESS", 
                                  "Weekly lesson reconciliation completed", 
                                  "Teachers: " + stats.teachersProcessed + "/" + (teacherData.length - 1) + 
                                  ", Lessons: " + stats.totalLessons + ", Students: " + stats.totalStudents +
                                  (stats.errors.length > 0 ? ", Errors: " + stats.errors.length : ""), "");
    
    if (stats.errors.length > 0) {
      UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "WARNING", 
                                    "Errors encountered", stats.errors.join("; "), "");
    }
    
    return stats;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("runWeeklyLessonReconciliation", "ERROR", 
                                  "Reconciliation failed", "", error.message);
    throw error;
  }
}

function reconcilePayment(paymentSheet, billingSheet, billingRowIndex, paymentRowNumber, invoiceValue, totalPayments) {
  UtilityScriptLibrary.debugLog('reconcilePayment', 'INFO', 'Starting reconciliation', 
               'Row: ' + paymentRowNumber + ', Total Payments: ' + totalPayments, '');
  
  var billingHeaders = billingSheet.getDataRange().getValues()[0];
  UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Got billing headers', 
               'Headers: ' + billingHeaders.join(', '), '');
  
  var paymentReceivedIndex = billingHeaders.indexOf("Payment Received");
  UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Looking for Payment Received column', 
               'Index: ' + paymentReceivedIndex + ', Found: ' + (paymentReceivedIndex !== -1), '');

  // Update billing sheet with total payments
  if (paymentReceivedIndex !== -1) {
    var targetRow = billingRowIndex + 1;
    var targetCol = paymentReceivedIndex + 1;
    UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Attempting to write to billing sheet', 
                 'Sheet: ' + billingSheet.getName() + ', Row: ' + targetRow + ', Col: ' + targetCol + 
                 ', Value: ' + totalPayments, '');
    
    billingSheet.getRange(targetRow, targetCol).setValue(totalPayments);
    
    // Verify the write
    var verifyValue = billingSheet.getRange(targetRow, targetCol).getValue();
    UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Verified write to billing sheet', 
                 'Written value: ' + totalPayments + ', Read back value: ' + verifyValue + 
                 ', Match: ' + (verifyValue === totalPayments), '');
  } else {
    UtilityScriptLibrary.debugLog('reconcilePayment', 'WARNING', 'Payment Received column not found', 
                 'Cannot update billing sheet', '');
  }

  // Set invoice number in payment sheet
  var headerMap = UtilityScriptLibrary.getHeaderMap(paymentSheet);
  UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Setting invoice in payment sheet', 
               'Invoice value: ' + invoiceValue + ', Has invoice column: ' + (headerMap["invoicenumber"] !== undefined), '');
  
  if (headerMap["invoicenumber"] && invoiceValue) {
    paymentSheet.getRange(paymentRowNumber, headerMap["invoicenumber"]).setValue(invoiceValue);
    UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Invoice number written', 
                 'Row: ' + paymentRowNumber + ', Value: ' + invoiceValue, '');
  }

  // Set reconciled checkbox
  if (headerMap["reconciled"]) {
    var reconciledCell = paymentSheet.getRange(paymentRowNumber, headerMap["reconciled"]);
    reconciledCell.insertCheckboxes();
    reconciledCell.setValue(true);
    UtilityScriptLibrary.debugLog('reconcilePayment', 'DEBUG', 'Reconciled checkbox set', 
                 'Row: ' + paymentRowNumber, '');
  }
  
  UtilityScriptLibrary.debugLog('reconcilePayment', 'SUCCESS', 'Reconciliation complete', 
               'Payment row: ' + paymentRowNumber, '');
}

function sumPayments(sheet, studentId, startDate, endDate) {
  try {
    UtilityScriptLibrary.debugLog('sumPayments', 'INFO', 'Starting payment sum', 
                 'Student: ' + studentId + ', Start: ' + startDate + ', End: ' + endDate, '');
    
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Got header map', 
                 'Keys: ' + Object.keys(headerMap).join(', '), '');
    
    var data = sheet.getDataRange().getValues();
    UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Got data', 
                 'Total rows: ' + data.length, '');
    
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var amountCol = headerMap[UtilityScriptLibrary.normalizeHeader('Amount Paid')];
    var dateCol = headerMap[UtilityScriptLibrary.normalizeHeader('Date')];
    
    UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Column indices', 
                 'Student ID col: ' + studentIdCol + ', Amount col: ' + amountCol + ', Date col: ' + dateCol, '');
    
    if (!studentIdCol || !amountCol || !dateCol) {
      UtilityScriptLibrary.debugLog('sumPayments', 'ERROR', 'Missing required columns', 
                   'Student ID: ' + studentIdCol + ', Amount: ' + amountCol + ', Date: ' + dateCol, '');
      return 0;
    }
    
    var total = 0;
    var paymentsFound = 0;
    
    // Skip header row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowStudentId = row[studentIdCol - 1];
      var rawAmount = row[amountCol - 1];
      var rawDate = row[dateCol - 1];
      
      // Debug first 3 rows and any matching student ID
      if (i <= 3 || rowStudentId === studentId) {
        UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Checking row ' + (i + 1), 
                     'Student ID: ' + rowStudentId + ', Amount: ' + rawAmount + ', Date: ' + rawDate, '');
      }
      
      // Check if student ID matches
      if (rowStudentId !== studentId) {
        if (i <= 3 || rowStudentId === studentId) {
          UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Student ID mismatch', 
                       'Row ' + (i + 1) + ': Looking for "' + studentId + '", found "' + rowStudentId + '"', '');
        }
        continue;
      }
      
      UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Student ID match found', 
                   'Row ' + (i + 1) + ': ' + rowStudentId, '');
      
      // Parse amount
      var amount = parseFloat(String(rawAmount).replace(/[^0-9.]/g, '').trim());
      if (isNaN(amount) || amount === 0) {
        UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Invalid or zero amount', 
                     'Row ' + (i + 1) + ': ' + rawAmount, '');
        continue;
      }
      
      UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Valid amount found', 
                   'Row ' + (i + 1) + ': $' + amount, '');
      
      // Parse and validate date
      var paymentDate = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (isNaN(paymentDate.getTime())) {
        UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Invalid date', 
                     'Row ' + (i + 1) + ': ' + rawDate, '');
        continue;
      }
      
      UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Valid date found', 
                   'Row ' + (i + 1) + ': ' + paymentDate, '');
      
      // Check if date is within range
      var inRange = paymentDate >= startDate && paymentDate <= endDate;
      UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Date range check', 
                   'Row ' + (i + 1) + ': Payment date ' + paymentDate + 
                   ' in range [' + startDate + ' to ' + endDate + ']? ' + inRange, '');
      
      if (inRange) {
        total += amount;
        paymentsFound++;
        UtilityScriptLibrary.debugLog('sumPayments', 'DEBUG', 'Payment added to total', 
                     'Row ' + (i + 1) + ': $' + amount + ', Running total: $' + total, '');
      }
    }
    
    UtilityScriptLibrary.debugLog('sumPayments', 'INFO', 'Payment sum calculated', 
                 'Student: ' + studentId + ', Payments found: ' + paymentsFound + ', Sum: $' + total, '');
    
    return total;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('sumPayments', 'ERROR', 'Payment sum failed', 
                 'Student: ' + studentId, error.message);
    return 0;
  }
}

function updateBillingForTeacherStudents(billingSheet, studentHours, teacherName) {
  var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
  var lastRow = billingSheet.getLastRow();
  
  if (lastRow < 2) {
    UtilityScriptLibrary.debugLog("updateBillingForTeacherStudents", "INFO", 
                                  "No data rows in billing sheet", "", "");
    return;
  }
  
  // Get column indices
  var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
  var teacherCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher')];
  var hoursThisCycleCol = headerMap[UtilityScriptLibrary.normalizeHeader('Current Hours Taught This Billing Cycle')];
  
  if (!studentIdCol || !teacherCol || !hoursThisCycleCol) {
    throw new Error('Required billing columns not found');
  }
  
  // Read ONLY the columns we need for matching
  var numDataRows = lastRow - 1;
  var studentIdData = billingSheet.getRange(2, studentIdCol, numDataRows, 1).getValues();
  var teacherData = billingSheet.getRange(2, teacherCol, numDataRows, 1).getValues();
  
  // Update ONLY the hours column for matching students, one cell at a time
  var updatedCount = 0;
  for (var i = 0; i < numDataRows; i++) {
    var rowStudentId = studentIdData[i][0];
    var rowTeacher = teacherData[i][0];
    
    // Only update rows for THIS teacher with matching student IDs
    if (rowTeacher === teacherName && studentHours[rowStudentId] !== undefined) {
      var rowNumber = i + 2; // +2 for header row and 0-based index
      billingSheet.getRange(rowNumber, hoursThisCycleCol).setValue(studentHours[rowStudentId]);
      updatedCount++;
    }
  }
  
  // Force recalculation of formulas
  if (updatedCount > 0) {
    SpreadsheetApp.flush();
    UtilityScriptLibrary.debugLog("updateBillingForTeacherStudents", "SUCCESS", 
                                  "Updated billing for teacher", 
                                  teacherName + " - " + updatedCount + " students updated", "");
  } else {
    UtilityScriptLibrary.debugLog("updateBillingForTeacherStudents", "INFO", 
                                  "No students updated for teacher", 
                                  teacherName, "");
  }
}

function updateSheetStudentWarnings(sheet, warningStudentMap, formattedDate) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return; // No data rows
  }
  
  var headers = data[0];
  
  // Find Student ID column
  var studentIdCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (UtilityScriptLibrary.normalizeHeader(headers[i]) === 'studentid') {
      studentIdCol = i;
      break;
    }
  }
  
  if (studentIdCol === -1) {
    return;
  }
  
  // Track which students have sections in this sheet
  var studentSections = {};
  var currentStudentId = null;
  var sectionStartRow = -1;
  
  // Identify student sections (rows between header rows)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var studentId = row[studentIdCol];
    
    // Check if this is a header row (bold, colored background)
    var range = sheet.getRange(i + 1, 1);
    var background = range.getBackground();
    var fontWeight = range.getFontWeight();
    
    // Check for header backgrounds using STYLES
    var isHeaderRow = (fontWeight === 'bold' || 
                       background === UtilityScriptLibrary.STYLES.WARNING.background || 
                       background === UtilityScriptLibrary.STYLES.SUBHEADER.background);
    
    if (isHeaderRow && studentId) {
      // Save previous section if exists
      if (currentStudentId && sectionStartRow > 0) {
        studentSections[currentStudentId] = {
          startRow: sectionStartRow,
          endRow: i
        };
      }
      
      // Start new section
      currentStudentId = studentId;
      sectionStartRow = i + 1; // 1-based row number
    }
  }
  
  // Save last section
  if (currentStudentId && sectionStartRow > 0) {
    studentSections[currentStudentId] = {
      startRow: sectionStartRow,
      endRow: data.length
    };
  }
  
  // Apply or remove warnings for each student section
  for (var studentId in studentSections) {
    var section = studentSections[studentId];
    var headerRow = section.startRow;
    
    var warningData = warningStudentMap[studentId];
    var shouldWarn = warningData !== undefined;
    
    if (shouldWarn) {
      // Apply WARNING style (pink background, dark red text)
      var headerRange = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn());
      headerRange.setBackground(UtilityScriptLibrary.STYLES.WARNING.background)
                 .setFontColor(UtilityScriptLibrary.STYLES.WARNING.text)
                 .setFontWeight('bold');
      
      // Column B: Keep student name format "LastName, FirstName - Instrument"
      var studentName = (warningData.lastName || '') + ', ' + (warningData.firstName || '') + ' - ' + (warningData.instrument || '');
      sheet.getRange(headerRow, 2).setValue(studentName);
      
      // Column D: Update with warning text including date
      var numericLength = UtilityScriptLibrary.extractNumericLessonLength(warningData.lessonLength);
      var warningText = numericLength + ' minutes - WARNING: Only ' + warningData.lessonsRemaining + 
                        ' lesson' + (warningData.lessonsRemaining === 1 ? '' : 's') + 
                        ' left as of ' + formattedDate;
      sheet.getRange(headerRow, 4).setValue(warningText);
      
    } else {
      // Apply SUBHEADER style (light blue background)
      var headerRange = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn());
      headerRange.setBackground(UtilityScriptLibrary.STYLES.SUBHEADER.background)
                 .setFontColor(UtilityScriptLibrary.STYLES.SUBHEADER.text)
                 .setFontWeight('bold');
      
      // Note: Student name should already be correct in Column B
      // Column D should already have the lesson length without warning
      // We don't need to change these unless they were previously warning-formatted
    }
  }
}

function updateTeacherRosterBalances(teacherSS, studentBalances, currentSemester) {
  // Extract season from semester name (e.g., "Spring 2025" -> "Spring")
  var season = UtilityScriptLibrary.extractSeasonFromSemester(currentSemester);
  if (!season) {
    throw new Error('Could not extract season from semester: ' + currentSemester);
  }
  
  var rosterSheetName = season + ' Roster';
  var rosterSheet = teacherSS.getSheetByName(rosterSheetName);
  
  if (!rosterSheet) {
    UtilityScriptLibrary.debugLog("updateTeacherRosterBalances", "WARNING", 
                                  "Roster sheet not found", rosterSheetName, "");
    return;
  }
  
  var headerMap = UtilityScriptLibrary.getHeaderMap(rosterSheet);
  var dataRange = rosterSheet.getRange(2, 1, Math.max(rosterSheet.getLastRow() - 1, 1), rosterSheet.getLastColumn());
  var data = dataRange.getValues();
  
  // Get column indices
  var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')] - 1;
  var hoursRemainingCol = headerMap[UtilityScriptLibrary.normalizeHeader('Hours Remaining')] - 1;
  var lessonsRemainingCol = headerMap[UtilityScriptLibrary.normalizeHeader('Lessons Remaining')] - 1;
  
  if (studentIdCol === undefined || hoursRemainingCol === undefined || lessonsRemainingCol === undefined) {
    throw new Error('Required roster columns not found');
  }
  
  // Create lookup map for quick access
  var balanceMap = {};
  for (var i = 0; i < studentBalances.length; i++) {
    var balance = studentBalances[i];
    balanceMap[balance.studentId] = balance;
  }
  
  // Update roster rows
  var updatedCount = 0;
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var studentId = row[studentIdCol];
    
    if (studentId && balanceMap[studentId]) {
      var balance = balanceMap[studentId];
      data[i][hoursRemainingCol] = balance.hoursRemaining;
      data[i][lessonsRemainingCol] = balance.lessonsRemaining;
      updatedCount++;
    }
  }
  
  // Write updated data back to roster sheet
  if (updatedCount > 0) {
    dataRange.setValues(data);
    UtilityScriptLibrary.debugLog("updateTeacherRosterBalances", "SUCCESS", 
                                  "Updated roster balances", 
                                  rosterSheetName + " - " + updatedCount + " students", "");
  }
}

// ============================================================================
// SECTION 6: GENERAL HELPER FUNCTIONS
// ============================================================================

function calculateLessonEquivalents(studentId, actualLength) {
  try {
    // Parse actual length to ensure it's a number
    var actualMinutes = parseInt(actualLength) || 0;
    if (actualMinutes <= 0) return 0;
    
    // Get student's registered lesson length from current billing sheet
    var registeredLength = getStudentRegisteredLessonLength(studentId);
    if (!registeredLength || registeredLength <= 0) {
      UtilityScriptLibrary.debugLog('⚠️ Could not find registered lesson length for student ' + studentId + ', defaulting to 30 minutes');
      return actualMinutes / 30; // Default to 30-minute lessons
    }
    
    // Calculate equivalents: actual minutes / registered minutes
    var lessonEquivalents = actualMinutes / registeredLength;
    
    UtilityScriptLibrary.debugLog('🔢 Student ' + studentId + ': ' + actualMinutes + 'min lesson = ' + lessonEquivalents + ' equivalents (registered for ' + registeredLength + 'min)');
    
    return Math.round(lessonEquivalents * 100) / 100; // Round to 2 decimal places
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error calculating lesson equivalents for student ' + studentId + ': ' + error.message);
    return 0;
  }
}

function calculateTotalCreditsApplied(billingData) {
  try {
    var totalCredits = 0;
    
    // ONLY include negative past balance as credit
    // Program credits are already shown as line items, so don't count them here
    if (billingData.pastBalance && billingData.pastBalance < 0) {
      var pastBalanceCredit = Math.abs(billingData.pastBalance);
      totalCredits += pastBalanceCredit;
      
      UtilityScriptLibrary.debugLog('calculateTotalCreditsApplied', 'DEBUG', 
                    'Past balance credit included', 
                    'Amount: $' + pastBalanceCredit, '');
    }
    
    UtilityScriptLibrary.debugLog('calculateTotalCreditsApplied', 'INFO', 
                  'Total credits calculated (non-line-item credits only)', 
                  'Total: $' + totalCredits, '');
    
    return totalCredits;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('calculateTotalCreditsApplied', 'ERROR', 
                  'Failed to calculate total credits', '', error.message);
    return 0;
  }
}

function expandTeacherAttendanceRows(teacherSS) {
  try {
    var sheets = teacherSS.getSheets();
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    var expanded = false;
    
    // Process each monthly attendance sheet
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName().toLowerCase();
      
      // Check if sheet name matches any month name (case-insensitive)
      var isMonthSheet = false;
      for (var j = 0; j < monthNames.length; j++) {
        if (sheetName === monthNames[j].toLowerCase()) {
          isMonthSheet = true;
          break;
        }
      }
      
      if (!isMonthSheet) continue;
      
      if (expandSheetAttendanceRows(sheet)) {
        expanded = true;
      }
    }
    
    return expanded;
    
  } catch (error) {
    throw error;
  }
}

function expandTeacherAttendanceSheets() {
  try {
    UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "INFO", "Expanding teacher attendance sheets", "", "");
    
    var teacherLookupSS = UtilityScriptLibrary.getWorkbook('formResponses');
    var teacherLookupSheet = teacherLookupSS.getSheetByName('Teacher Roster Lookup');
    
    if (!teacherLookupSheet) {
      UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "ERROR", "Teacher Roster Lookup sheet not found", "", "");
      return;
    }

    var teacherData = teacherLookupSheet.getDataRange().getValues();
    if (teacherData.length <= 1) {
      UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "WARNING", "No teacher data found", "", "");
      return;
    }

    var headers = teacherData[0];
    
    var teacherNameCol = headers.indexOf('Teacher Name');
    var rosterUrlCol = headers.indexOf('Roster URL');
    var statusCol = headers.indexOf('Status');
    
    if (teacherNameCol === -1 || rosterUrlCol === -1 || statusCol === -1) {
      UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "ERROR", "Required columns not found in Teacher Roster Lookup", "", "");
      return;
    }
    
    var processedCount = 0;
    var expandedCount = 0;
    
    for (var i = 1; i < teacherData.length; i++) {
      var row = teacherData[i];
      var teacherName = row[teacherNameCol];
      var rosterUrl = row[rosterUrlCol];
      var status = row[statusCol];
      
      if (!teacherName || !rosterUrl || status !== 'Active') {
        continue;
      }
      
      UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "DEBUG", "Processing teacher", 
                   "Name: " + teacherName + ", Status: " + status + ", URL: " + (rosterUrl ? "present" : "missing"), "");
      
      try {
        var fileIdMatch = rosterUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
        if (!fileIdMatch) {
          UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "WARNING", "Invalid roster URL", 
                       "Teacher: " + teacherName + ", URL: " + rosterUrl, "");
          continue;
        }
        
        UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "DEBUG", "Opening teacher spreadsheet", 
                     "Teacher: " + teacherName + ", File ID: " + fileIdMatch[1], "");
        
        var teacherSS = SpreadsheetApp.openById(fileIdMatch[1]);
        
        // OPTIMIZATION: Use getMonthSheets to get all month sheets at once
        var monthSheets = UtilityScriptLibrary.getMonthSheets(teacherSS);
        var teacherExpanded = false;
        
        for (var j = 0; j < monthSheets.length; j++) {
          if (expandSheetAttendanceRows(monthSheets[j])) {
            teacherExpanded = true;
          }
        }
        
        if (teacherExpanded) {
          expandedCount++;
        }
        
        UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "DEBUG", "Teacher processing result", 
                     "Teacher: " + teacherName + ", Expanded: " + teacherExpanded, "");
        
        processedCount++;
        
      } catch (teacherError) {
        UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "WARNING", "Failed to process teacher", 
                     "Teacher: " + teacherName, teacherError.message);
      }
    }
    
    UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "SUCCESS", "Expanded attendance sheets", 
                 "Active teachers: " + processedCount + ", Processed: " + processedCount + ", Expanded: " + expandedCount, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("expandTeacherAttendanceSheets", "ERROR", "Failed to expand attendance sheets", 
                 "", error.message);
    throw error;
  }
}

function extractBillingDataFromRow(billingRowData, headerMap) {
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  // Extract basic billing data
  var billingData = {
    parentFirstName: billingRowData[headerMap[norm("Parent First Name")] - 1] || '',
    parentLastName: billingRowData[headerMap[norm("Parent Last Name")] - 1] || '',
    parentAddress: billingRowData[headerMap[norm("Parent Address")] - 1] || '',
    termOfAddress: billingRowData[headerMap[norm("Salutation")] - 1] || '',
    invoiceNumber: billingRowData[headerMap[norm("Invoice Number")] - 1] || '',
    invoiceDate: billingRowData[headerMap[norm("Invoice Date")] - 1] || '',
    dueDate: billingRowData[headerMap[norm("Due Date")] - 1] || '',
    invoiceUrl: billingRowData[headerMap[norm("Invoice URL")] - 1] || '',
    lessonLength: billingRowData[headerMap[norm("Lesson Length")] - 1] || '',
    pastBalance: parseFloat(billingRowData[headerMap[norm("Past Balance")] - 1]) || 0,
    lateFee: parseFloat(billingRowData[headerMap[norm("Late Fee")] - 1]) || 0,
    currentInvoiceTotal: parseFloat(billingRowData[headerMap[norm("Current Invoice Total")] - 1]) || 0,
    paymentReceived: parseFloat(billingRowData[headerMap[norm("Payment Received")] - 1]) || 0,
    credit: parseFloat(billingRowData[headerMap[norm("Credit")] - 1]) || 0,
    currentBalance: parseFloat(billingRowData[headerMap[norm("Current Balance")] - 1]) || 0,
    comments: billingRowData[headerMap[norm("Admin Comments")] - 1] || '',
    letterType: billingRowData[headerMap[norm("Letter Type")] - 1] || '',
    agreementForm: billingRowData[headerMap[norm("Agreement Form")] - 1] || false,
    mediaRelease: billingRowData[headerMap[norm("Media Release")] - 1] || false,
    deliveryPreference: billingRowData[headerMap[norm("Delivery Preference")] - 1] || ''
  };
  
  // CRITICAL FIX: Extract program totals for dynamic invoice generation
  billingData.programTotals = extractProgramTotals(billingRowData, headerMap);
  
  UtilityScriptLibrary.debugLog("extractBillingDataFromRow", "DEBUG", "Extracted billing data with program totals", 
                "Program count: " + billingData.programTotals.programs.length + ", Letter Type: " + billingData.letterType, "");
  
  return billingData;
}

function extractDeliveryPreference(billingRowData, headerMap) {
  try {
    var norm = UtilityScriptLibrary.normalizeHeader;
    var deliveryPrefCol = headerMap[norm("Delivery Preference")];
    
    if (!deliveryPrefCol) {
      UtilityScriptLibrary.debugLog("extractDeliveryPreference", "WARNING", "Delivery Preference column not found", "", "");
      return "Email"; // Default value
    }
    
    var deliveryPreference = billingRowData[deliveryPrefCol - 1];
    
    // Return default if empty or undefined
    if (!deliveryPreference || deliveryPreference.toString().trim() === '') {
      return "Email";
    }
    
    return deliveryPreference.toString().trim();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractDeliveryPreference", "ERROR", "Failed to extract delivery preference", 
                  "", error.message);
    return "Email"; // Default fallback
  }
}

function extractPreviousBillingData(options) {
  var includeAll = options && options.includeAll !== undefined ? options.includeAll : false;
  
  UtilityScriptLibrary.debugLog("extractPreviousBillingData", "INFO", "Extracting previous billing data", 
                "Include all: " + includeAll, "");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = ss.getSheetByName("Billing Metadata");
    
    if (!metadataSheet) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "WARNING", "Billing Metadata sheet not found", "", "");
      return null;
    }
    
    var lastRow = metadataSheet.getLastRow();
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "WARNING", "No existing billing cycles found", 
                    "Rows: " + lastRow, "");
      return null;
    }
    
    // Get current semester from Semester Metadata (most recent semester)
    var semesterSheet = ss.getSheetByName('Semester Metadata');
    if (!semesterSheet) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "ERROR", "Semester Metadata sheet not found", "", "");
      return null;
    }
    
    var semesterData = semesterSheet.getDataRange().getValues();
    var currentSemester = semesterData[semesterData.length - 1][0]; // Most recent semester
    
    // Use dynamic column lookup for Billing Metadata sheet
    var metadataHeaderMap = UtilityScriptLibrary.getHeaderMap(metadataSheet);
    var semesterCol = metadataHeaderMap[UtilityScriptLibrary.normalizeHeader("Semester Name")];
    var billingMonthCol = metadataHeaderMap[UtilityScriptLibrary.normalizeHeader("Billing Month")];
    
    if (!semesterCol) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "ERROR", "Semester Name column not found in Billing Metadata", "", "");
      return null;
    }
    
    if (!billingMonthCol) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "ERROR", "Billing Month column not found in Billing Metadata", "", "");
      return null;
    }
    
    // Get the last existing billing cycle's semester (for logging only)
    var lastExistingRowSemester = metadataSheet.getRange(lastRow, semesterCol).getValue();
    
    UtilityScriptLibrary.debugLog("extractPreviousBillingData", "DEBUG", "Comparing semesters", 
                  "Current: " + currentSemester + ", Last existing: " + lastExistingRowSemester, "");
    
    if (lastExistingRowSemester !== currentSemester) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "INFO", "Different semester detected - carrying over from previous semester", 
                    "Current: " + currentSemester + ", Previous: " + lastExistingRowSemester, "");
    }
    
    // Get the previous billing data regardless of semester
    var previousBillingMonth = metadataSheet.getRange(lastRow, billingMonthCol).getDisplayValue();
    
    UtilityScriptLibrary.debugLog("extractPreviousBillingData", "DEBUG", "Extracting from previous cycle", 
                  "Looking for sheet: " + previousBillingMonth, "");
    
    var billingSheet = ss.getSheetByName(previousBillingMonth);
    if (!billingSheet) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "WARNING", "Previous billing sheet not found", 
                    "Sheet name: " + previousBillingMonth, "");
      return null;
    }
    
    var data = billingSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    
    // Use normalized header lookups instead of hardcoded strings
    var balanceCol = headerMap[UtilityScriptLibrary.normalizeHeader("Current Balance")];
    var idCol = headerMap[UtilityScriptLibrary.normalizeHeader("Student ID")];
    var hoursRemainingCol = headerMap[UtilityScriptLibrary.normalizeHeader("Hours Remaining")];
    
    if (!balanceCol || !idCol || !hoursRemainingCol) {
      UtilityScriptLibrary.debugLog("extractPreviousBillingData", "ERROR", "Required columns not found", 
                    "Balance col: " + balanceCol + ", ID col: " + idCol + ", Hours col: " + hoursRemainingCol, "");
      return null;
    }
    
    var rowsToCarry = {};
    var carriedCount = 0;
    var totalRows = data.length - 1; // Exclude header
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentId = row[idCol - 1];
      var balance = parseFloat(row[balanceCol - 1]) || 0;
      var hoursRemaining = parseFloat(row[hoursRemainingCol - 1]) || 0;
      
      if (studentId && (includeAll || balance !== 0 || hoursRemaining !== 0)) {
        rowsToCarry[studentId] = row;
        carriedCount++;
      }
    }
    
    UtilityScriptLibrary.debugLog("extractPreviousBillingData", "INFO", "Previous billing data extracted successfully", 
                  "Sheet: " + previousBillingMonth + ", Total rows: " + totalRows + ", Carried: " + carriedCount, "");
    
    return {
      previousSheetName: previousBillingMonth,
      prevHeaderMap: headerMap,
      rowsToCarry: rowsToCarry
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("extractPreviousBillingData", "ERROR", "Failed to extract previous billing data", 
                  "Include all: " + includeAll, error.message);
    throw error;
  }
}

function extractProgramTotals(row, headerMap) {
  var norm = UtilityScriptLibrary.normalizeHeader;
  var programs = [];
  var lessonTotal = 0;
  var lessonQuantity = 0;
  
  UtilityScriptLibrary.debugLog("extractProgramTotals", "DEBUG", "Starting program extraction", 
                "Header count: " + Object.keys(headerMap).length, "");
  
  // Look for program columns (Lesson, Group, etc.)
  for (var header in headerMap) {
    if (header.indexOf('total') !== -1 && header !== 'currentinvoicetotal') {
      var colIndex = headerMap[header];
      var total = parseFloat(row[colIndex - 1]) || 0;
      
      // Extract program name from header (e.g., "lessontotal" -> "lesson")
      var programName = header.replace('total', '');
      
      // Get corresponding quantity, credit, and price
      var qtyHeader = programName + 'quantity';
      var creditHeader = programName + 'credit';
      var priceHeader = programName + 'price';
      
      var quantity = headerMap[qtyHeader] ? (parseFloat(row[headerMap[qtyHeader] - 1]) || 0) : 0;
      var credit = headerMap[creditHeader] ? (parseFloat(row[headerMap[creditHeader] - 1]) || 0) : 0;
      var price = headerMap[priceHeader] ? (parseFloat(row[headerMap[priceHeader] - 1]) || 0) : 0;
      
      // Include programs that have ANY activity (total, quantity, or credit != 0)
      if (total !== 0 || quantity !== 0 || credit !== 0) {
        // Capitalize first letter of program name
        var displayName = programName.charAt(0).toUpperCase() + programName.slice(1);
        
        programs.push({
          name: displayName,
          quantity: quantity,
          credit: credit,
          price: price,
          total: total
        });
        
        // Track lesson-specific data
        if (programName.toLowerCase() === 'lesson') {
          lessonTotal = total;
          lessonQuantity = quantity;
        }
        
        UtilityScriptLibrary.debugLog("extractProgramTotals", "DEBUG", "Found program", 
                      "Program: " + displayName + ", Qty: " + quantity + ", Credit: " + credit + ", Price: " + price + ", Total: " + total, "");
      }
    }
  }
  
  UtilityScriptLibrary.debugLog("extractProgramTotals", "INFO", "Program extraction completed", 
                "Total programs found: " + programs.length, "");
  
  return {
    programs: programs,
    lessonTotal: lessonTotal,
    lessonQuantity: lessonQuantity
  };
}

function extractStudentDataFromBillingRow(row, headerMap) {
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  // Add diagnostic logging
  var lessonQtyColIndex = headerMap[norm("Lesson Quantity")];
  var lessonQtyValue = lessonQtyColIndex ? row[lessonQtyColIndex - 1] : undefined;
  
  UtilityScriptLibrary.debugLog("extractStudentDataFromBillingRow", "DEBUG", "Extracting lesson quantity", 
                "Column index: " + lessonQtyColIndex + ", Raw value: '" + lessonQtyValue + "', Type: " + typeof lessonQtyValue, "");
  
  return {
    firstName: row[headerMap[norm("Student First Name")] - 1] || '',
    lastName: row[headerMap[norm("Student Last Name")] - 1] || '',
    studentId: row[headerMap[norm("Student ID")] - 1] || '',
    age: row[headerMap[norm("Age")] - 1] || '',
    instrument: row[headerMap[norm("Instrument")] - 1] || '',
    teacher: row[headerMap[norm("Teacher")] - 1] || '',
    lessonLength: row[headerMap[norm("Lesson Length")] - 1] || '',
    lessonQuantity: lessonQtyValue || ''
  };
}

function getActivePrograms() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Programs List');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var activeIndex = headers.indexOf("Active");
  var nameIndex = headers.indexOf("Program Name");

  var activePrograms = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[activeIndex] === true) {
      activePrograms.push(row[nameIndex]);
    }
  }
  
  return activePrograms;
}

function getCurrentBillingCycleDates() {
  try {
    UtilityScriptLibrary.debugLog("getCurrentBillingCycleDates", "INFO", "Getting current billing cycle dates", "", "");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = ss.getSheetByName('Billing Metadata');
    
    if (!metadataSheet) {
      throw new Error('Billing Metadata sheet not found');
    }
    
    var metadataData = metadataSheet.getDataRange().getValues();
    if (metadataData.length < 2) {
      throw new Error('No billing metadata rows found');
    }
    
    // Get the most recent billing cycle (last row)
    var latestRow = metadataData[metadataData.length - 1];
    
    // Get Payment Starting Date (column 2)
    var paymentStartDate = latestRow[2];
    
    // Convert to proper Date object if it isn't already
    if (!(paymentStartDate instanceof Date)) {
      paymentStartDate = new Date(paymentStartDate);
    }
    
    // End date is always today for current billing cycle
    var paymentEndDate = new Date();
    
    UtilityScriptLibrary.debugLog("getCurrentBillingCycleDates", "SUCCESS", "Retrieved billing cycle dates", 
                 "Start: " + paymentStartDate.toDateString() + ", End: " + paymentEndDate.toDateString(), "");
    
    return {
      startDate: paymentStartDate,
      endDate: paymentEndDate
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getCurrentBillingCycleDates", "ERROR", "Failed to get billing cycle dates", 
                 "", error.message);
    throw error;
  }
}

function getCurrentBillingSheet() {
  try {
    UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "INFO", "Starting billing sheet lookup", "", "");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = ss.getSheetByName('Billing Metadata');
    
    if (!metadataSheet) {
      UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "ERROR", "Billing Metadata sheet not found", "", "");
      return null;
    }
    
    var metadataData = metadataSheet.getDataRange().getValues();
    if (metadataData.length < 2) {
      UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "ERROR", "No billing metadata rows found", "", "");
      return null;
    }
    
    var latestRow = metadataData[metadataData.length - 1];
    
    // FIXED: Convert the date object to the proper sheet name format using utility function
    var billingDate = latestRow[0]; // This is a date object
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var month = monthNames[billingDate.getMonth()];
    var year = billingDate.getFullYear();
    var latestBillingCycle = month + " " + year; // "January 2024"
    
    UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "INFO", "Constructed billing cycle name", 
                 "Date object: " + billingDate + " → Sheet name: " + latestBillingCycle, "");
    
    var billingSheet = ss.getSheetByName(latestBillingCycle);
    
    if (!billingSheet) {
      var allSheets = ss.getSheets().map(function(sheet) { return sheet.getName(); });
      UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "ERROR", "Billing sheet still not found", 
                   "Looking for: '" + latestBillingCycle + "' | Available: " + allSheets.join(', '), "");
      return null;
    }
    
    UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "SUCCESS", "Found billing sheet", latestBillingCycle, "");
    return billingSheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getCurrentBillingSheet", "ERROR", "Exception occurred", "", error.message);
    return null;
  }
}

function getCurrentRateChartName() {
  try {
    var ratesSheet = UtilityScriptLibrary.getSheet('rates');
    if (!ratesSheet) {
      throw new Error('Rates sheet not found.');
    }
    
    var headers = ratesSheet.getRange(1, 1, 1, ratesSheet.getLastColumn()).getValues()[0];
    
    // Column B (index 1) always contains the most recent rate chart
    var currentRateChart = headers[1];
    
    if (!currentRateChart || typeof currentRateChart !== 'string') {
      throw new Error('Rate chart name in column B is invalid or empty.');
    }

    return currentRateChart;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('getCurrentRateChartName', 'ERROR', 'Error getting current rate chart name', '', error.message);
    throw error;
  }
}

function getCurrentSemesterFromBillingMetadata(billingSheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = ss.getSheetByName('Billing Metadata');
    
    if (!metadataSheet) {
      UtilityScriptLibrary.debugLog("getCurrentSemesterFromBillingMetadata", "ERROR", "Billing Metadata sheet not found", "", "");
      return null;
    }
    
    var data = metadataSheet.getDataRange().getValues();
    var headers = data[0];
    
    var billingMonthCol = -1;
    var semesterCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      if (UtilityScriptLibrary.normalizeHeader(headers[i]) === UtilityScriptLibrary.normalizeHeader('Billing Month')) {
        billingMonthCol = i;
      }
      if (UtilityScriptLibrary.normalizeHeader(headers[i]) === UtilityScriptLibrary.normalizeHeader('Semester Name')) {
        semesterCol = i;
      }
    }
    
    if (billingMonthCol === -1 || semesterCol === -1) {
      UtilityScriptLibrary.debugLog("getCurrentSemesterFromBillingMetadata", "ERROR", "Required columns not found in Billing Metadata", "", "");
      return null;
    }
    
    // Get month names array once
    var monthNames = UtilityScriptLibrary.getMonthNames();
    
    // Find the row matching the billing sheet name
    for (var i = 1; i < data.length; i++) {
      var metadataValue = data[i][billingMonthCol];
      
      if (!metadataValue) continue;
      
      // Convert to string for comparison (handles both text and date objects)
      var metadataString = '';
      if (metadataValue instanceof Date) {
        // Format date as "Month YYYY" using the array
        metadataString = monthNames[metadataValue.getMonth()] + ' ' + metadataValue.getFullYear();
      } else {
        metadataString = String(metadataValue).trim();
      }
      
      if (metadataString === billingSheetName) {
        var semester = data[i][semesterCol];
        UtilityScriptLibrary.debugLog("getCurrentSemesterFromBillingMetadata", "INFO", "Found semester", 
                      "Billing Month: " + billingSheetName + ", Semester: " + semester, "");
        return semester;
      }
    }
    
    UtilityScriptLibrary.debugLog("getCurrentSemesterFromBillingMetadata", "WARNING", "No matching billing month found", 
                  "Searched for: " + billingSheetName, "");
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getCurrentSemesterFromBillingMetadata", "ERROR", "Error getting semester", 
                  "", error.message);
    return null;
  }
}

function getCurrentSemesterName() {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    UtilityScriptLibrary.debugLog('📋 getCurrentSemesterName - Getting current semester name');
    
    var semesterSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterSheet) {
      UtilityScriptLibrary.debugLog('📋 getCurrentSemesterName - Semester Metadata sheet not found');
      return 'Unknown Semester';
    }
    
    var data = semesterSheet.getDataRange().getValues();
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog('📋 getCurrentSemesterName - No semester data found in metadata');
      return 'Unknown Semester';
    }
    
    // Get the most recent semester (last row)
    var currentSemester = data[data.length - 1][0] || 'Unknown Semester';
    UtilityScriptLibrary.debugLog('📋 getCurrentSemesterName - Found current semester: ' + currentSemester);
    
    return currentSemester;
    
  }, 'Current semester name retrieved successfully', 'getCurrentSemesterName', {
    showUI: false,
    logLevel: 'INFO'
  }).data;
}

function getCurrentSemesterInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var semesterSheet = ss.getSheetByName('semesterMetadata');
    
    if (semesterSheet) {
      var data = semesterSheet.getDataRange().getValues();
      if (data.length > 1) {
        var latestSemester = data[data.length - 1];
        var semesterName = latestSemester[0];
        var startDate = latestSemester[1];
        
        // Extract academic year from semester name (e.g., "Fall 2024" -> "2024-2025")
        var year = semesterName.match(/\d{4}/);
        if (year) {
          var currentYear = parseInt(year[0]);
          var academicYear = `${currentYear}-${currentYear + 1}`;
          var formattedStartDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), 'MMMM d, yyyy');
          
          return {
            academicYear: academicYear,
            startDate: formattedStartDate
          };
        }
      }
    }
    
    // Fallback
    var currentYear = new Date().getFullYear();
    return {
      academicYear: `${currentYear}-${currentYear + 1}`,
      startDate: 'September 1, ' + currentYear
    };
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getCurrentSemesterInfo", "ERROR", "Error getting semester info", "", error.message);
    var currentYear = new Date().getFullYear();
    return {
      academicYear: `${currentYear}-${currentYear + 1}`,
      startDate: 'September 1, ' + currentYear
    };
  }
}

function getCurrentSemesterRateForLength(lessonLength) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ratesSheet = ss.getSheetByName("Rates");
    var rateHeaders = ratesSheet.getRange(1, 1, 1, ratesSheet.getLastColumn()).getValues()[0];
    var rateYearCol = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
    var rateMap = UtilityScriptLibrary.buildRateMapFromSheet(ratesSheet, rateHeaders, rateYearCol);
    
    var lengthKey = lessonLength + 'Min';
    
    return parseFloat(rateMap[lengthKey]) || parseFloat(rateMap['30Min']) || 12;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error getting rate for ' + lessonLength + ' minute lessons: ' + error.message);
    return 12;
  }
}

function getExpandedPrograms(row, context) {
  var enrollmentRaw = row[context.getColIndex("Enrollment")] || '';
  var enrolled = enrollmentRaw
    .toString()
    .split(',')
    .map(function(p) { return p.trim().toLowerCase(); })
    .filter(function(p) { return p; });

  // ES5: Use object to track unique programs instead of Set
  var expandedObj = {};
  for (var i = 0; i < enrolled.length; i++) {
    expandedObj[enrolled[i]] = true;
  }

  for (var j = 0; j < enrolled.length; j++) {
    var programName = enrolled[j];
    var programRow = null;
    
    // Find the program in programData
    for (var k = 0; k < context.programData.length; k++) {
      var r = context.programData[k];
      var nameValue = r[context.nameCol];
      if (nameValue && nameValue.toString().toLowerCase() === programName) {
        programRow = r;
        break;
      }
    }
    
    if (!programRow) continue;

    var type = programRow[context.typeCol];
    var aliasFor = programRow[context.aliasCol] || '';

    if (type === "Package" && aliasFor) {
      var parts = aliasFor
        .split(',')
        .map(function(p) { return p.trim().toLowerCase(); })
        .filter(function(p) { return p; });
      
      for (var m = 0; m < parts.length; m++) {
        expandedObj[parts[m]] = true;
      }
    }
  }

  // Convert object keys to array for return
  var result = [];
  for (var key in expandedObj) {
    if (expandedObj.hasOwnProperty(key)) {
      result.push(key);
    }
  }
  
  return result;  // Return ARRAY, not Set or Object
}

function getFormsDataFromContacts() {
  try {
    UtilityScriptLibrary.debugLog('📋 Getting forms data from Contacts...');
    
    var contactsSS = UtilityScriptLibrary.getWorkbook('contacts');
    var studentsSheet = contactsSS.getSheetByName('students');
    
    if (!studentsSheet) {
      throw new Error('Students sheet not found in Contacts workbook');
    }
    
    // RANGE VALIDATION: Check if students sheet has data
    var lastRow = studentsSheet.getLastRow();
    if (lastRow < 2) {
      UtilityScriptLibrary.debugLog('No student data found in Contacts workbook - returning empty forms data');
      return {};
    }
    
    var data = studentsSheet.getDataRange().getValues();
    
    // RANGE VALIDATION: Double-check data array
    if (!data || data.length < 2) {
      UtilityScriptLibrary.debugLog('No data rows found in students sheet - returning empty forms data');
      return {};
    }
    
    var headers = data[0];
    
    var studentIdCol = headers.findIndex(function(h) { 
      return UtilityScriptLibrary.normalizeHeader(h) === 'student id'; 
    });
    var agreementCol = headers.findIndex(function(h) { 
      return UtilityScriptLibrary.normalizeHeader(h) === 'current agreement'; 
    });
    var mediaCol = headers.findIndex(function(h) { 
      return UtilityScriptLibrary.normalizeHeader(h) === 'current media'; 
    });
    
    if (studentIdCol === -1 || agreementCol === -1 || mediaCol === -1) {
      throw new Error('Required columns not found in Students sheet');
    }
    
    var formsData = {};
    
    // Process each student row (starting from row 1 to skip headers)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentId = row[studentIdCol];
      
      if (!studentId) continue;
      
      // UPDATED: Handle new data structure
      // Agreement: boolean (checkbox) - true if received
      var hasAgreement = row[agreementCol] === true;
      
      // Media: any non-blank response means form received
      var mediaResponse = row[mediaCol];
      var hasMediaResponse = (mediaResponse && String(mediaResponse).trim() !== '');
      
      formsData[studentId] = {
        agreement: hasAgreement,
        media: hasMediaResponse
      };
    }
    
    UtilityScriptLibrary.debugLog('✅ Retrieved forms data for ' + Object.keys(formsData).length + ' students');
    return formsData;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error getting forms data: ' + error.message);
    throw error;
  }
}

function getLessonLengthFromRow(row, get) {
  var qty60Col = get("Qty60");
  var qty45Col = get("Qty45");
  var qty30Col = get("Qty30");
  
  var qty60 = 0;
  var qty45 = 0;
  var qty30 = 0;
  
  if (qty60Col !== -1 && row[qty60Col]) {
    qty60 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty60Col]) || 0;
  }
  if (qty45Col !== -1 && row[qty45Col]) {
    qty45 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty45Col]) || 0;
  }
  if (qty30Col !== -1 && row[qty30Col]) {
    qty30 = UtilityScriptLibrary.extractLessonQuantityFromPackage(row[qty30Col]) || 0;
  }
  
  if (qty60 > 0) return "60";
  else if (qty45 > 0) return "45";
  else return "30";
}

function getPreviousSemester(currentSemesterName) {
  try {
    var semesterMetadataSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      return null;
    }
    
    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find columns
    var nameCol = -1, startCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      }
    }
    
    if (nameCol === -1 || startCol === -1) {
      return null;
    }
    
    // Build array of semesters with start dates
    var semesters = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[nameCol] && row[startCol]) {
        semesters.push({
          name: row[nameCol],
          startDate: new Date(row[startCol])
        });
      }
    }
    
    // Sort by start date
    semesters.sort(function(a, b) {
      return a.startDate - b.startDate;
    });
    
    // Find current semester and return previous one
    for (var i = 0; i < semesters.length; i++) {
      if (semesters[i].name === currentSemesterName) {
        if (i > 0) {
          UtilityScriptLibrary.debugLog("getPreviousSemester", "INFO", 
                                        "Found previous semester", 
                                        "Current: " + currentSemesterName + ", Previous: " + semesters[i-1].name, "");
          return semesters[i-1].name;
        } else {
          UtilityScriptLibrary.debugLog("getPreviousSemester", "INFO", 
                                        "No previous semester (first semester)", 
                                        currentSemesterName, "");
          return null;
        }
      }
    }
    
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getPreviousSemester", "ERROR", "Error finding previous semester", "", error.message);
    return null;
  }
}

function getPreviousSemesterBalance(studentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var balanceSheet = ss.getSheetByName('Semester Lesson Balances');
    
    if (!balanceSheet) {
      return null;
    }
    
    var data = balanceSheet.getDataRange().getValues();
    
    // Find most recent active balance for this student
    for (var i = data.length - 1; i >= 1; i--) { // Start from bottom, skip header
      var row = data[i];
      if (row[0] === studentId && row[6] === 'active' && row[5] > 0) { // Student ID, Status, Positive Balance
        return {
          studentId: row[0],
          semesterName: row[1],
          lessonLength: row[2],
          registeredLessons: row[3],
          taughtLessons: row[4],
          lessonBalance: row[5]
        };
      }
    }
    
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error getting previous semester balance for student ' + studentId + ': ' + error.message);
    return null;
  }
}

function getRateForSemester(semesterName, rateType) {
  try {
    // Get the rate chart name for this semester
    var rateChartName = getRateChartForSemester(semesterName);
    
    // Get the rate from the Rates sheet
    var billingSS = SpreadsheetApp.getActiveSpreadsheet();
    var ratesSheet = billingSS.getSheetByName('Rates');
    if (!ratesSheet) throw new Error('Rates sheet not found.');

    var data = ratesSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find the rate chart column
    var rateChartColIndex = headers.indexOf(rateChartName);
    if (rateChartColIndex === -1) {
      throw new Error(`Rate chart "${rateChartName}" not found in Rates sheet.`);
    }
    
    // Find the rate type row
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === rateType) { // Column A contains rate titles
        var rate = data[i][rateChartColIndex];
        if (rate === undefined || rate === null || rate === '') {
          throw new Error(`Rate for "${rateType}" not found in chart "${rateChartName}"`);
        }
        return rate;
      }
    }
    
    throw new Error(`Rate type "${rateType}" not found in Rates sheet.`);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog(`❌ Error getting rate: ${error.message}`);
    throw error;
  }
}

function getRateChartForSemester(semesterName) {
  try {
    var billingSS = SpreadsheetApp.getActiveSpreadsheet();
    var metadataSheet = billingSS.getSheetByName('semesterMetadata');
    if (!metadataSheet) throw new Error('Semester Metadata sheet not found.');

    var data = metadataSheet.getDataRange().getValues();
    var headers = data[0];
    
    var semesterCol = 0; // Column A
    var rateChartCol = 3; // Column D (was "Rates Verification", now rate chart name)
    
    // Find the semester row
    for (var i = 1; i < data.length; i++) {
      if (data[i][semesterCol] === semesterName) {
        var rateChart = data[i][rateChartCol];
        if (!rateChart || typeof rateChart !== 'string') {
          throw new Error(`No rate chart found for semester "${semesterName}"`);
        }
        return rateChart;
      }
    }
    
    throw new Error(`Semester "${semesterName}" not found in metadata.`);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog(`❌ Error getting rate chart for semester: ${error.message}`);
    throw error;
  }
}

function getRateColumnFromMetadata(semesterName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var metadataSheet = ss.getSheetByName("Semester Metadata");
  
  if (!metadataSheet) {
    throw new Error("Semester Metadata sheet not found.");
  }
  
  var data = metadataSheet.getDataRange().getValues();
  
  if (data.length < 2) {
    throw new Error("No semester metadata rows found.");
  }
  
  // Get headers from first row
  var headers = data[0];
  
  // Find column indices dynamically
  var semesterNameCol = -1;
  var ratesVerificationCol = -1;
  
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === "Semester Name") semesterNameCol = i;
    if (headers[i] === "Rates Verification") ratesVerificationCol = i;
  }
  
  if (semesterNameCol === -1) {
    throw new Error("'Semester Name' column not found in Semester Metadata.");
  }
  if (ratesVerificationCol === -1) {
    throw new Error("'Rates Verification' column not found in Semester Metadata.");
  }
  
  // Find the row matching the semester name
  for (var i = 1; i < data.length; i++) {
    if (data[i][semesterNameCol] === semesterName) {
      var rateColumnLabel = data[i][ratesVerificationCol];
      
      if (!rateColumnLabel || typeof rateColumnLabel !== 'string') {
        throw new Error("Rate column label not found for semester: " + semesterName);
      }
      
      return rateColumnLabel;
    }
  }
  
  throw new Error("Semester '" + semesterName + "' not found in Semester Metadata.");
}

function getSemesterForDate(targetDate) {
  try {
    var semesterMetadataSheet = UtilityScriptLibrary.getSheet('semesterMetadata');
    if (!semesterMetadataSheet) {
      UtilityScriptLibrary.debugLog("getSemesterForDate", "ERROR", "Semester Metadata sheet not found", "", "");
      return null;
    }
    
    var data = semesterMetadataSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find columns
    var nameCol = -1, startCol = -1, endCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var header = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (header.indexOf('semester') !== -1 || header.indexOf('name') !== -1) {
        nameCol = i;
      } else if (header.indexOf('start') !== -1) {
        startCol = i;
      } else if (header.indexOf('end') !== -1) {
        endCol = i;
      }
    }
    
    if (nameCol === -1 || startCol === -1 || endCol === -1) {
      UtilityScriptLibrary.debugLog("getSemesterForDate", "ERROR", "Required columns not found", "", "");
      return null;
    }
    
    // Find semester where targetDate falls between start and end
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var semesterName = row[nameCol];
      var startDate = new Date(row[startCol]);
      var endDate = new Date(row[endCol]);
      
      if (targetDate >= startDate && targetDate <= endDate) {
        UtilityScriptLibrary.debugLog("getSemesterForDate", "INFO", 
                                      "Found semester for date", 
                                      "Date: " + targetDate.toDateString() + ", Semester: " + semesterName, "");
        return semesterName;
      }
    }
    
    UtilityScriptLibrary.debugLog("getSemesterForDate", "WARNING", 
                                  "No semester found for date", 
                                  targetDate.toDateString(), "");
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getSemesterForDate", "ERROR", "Error finding semester", "", error.message);
    return null;
  }
}

function createNewAttendanceSheets() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    UtilityScriptLibrary.debugLog("createNewAttendanceSheets", "INFO", "Starting attendance sheet creation", "", "");
    
    // Show the month/year selection dialog
    // The dialog will handle the rest via handleMonthYearSelection()
    promptForMonthAndYear();
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("createNewAttendanceSheets", "ERROR", "Fatal error", "", error.message);
    ui.alert('❌ Error', 'An error occurred:\n\n' + error.message, ui.ButtonSet.OK);
  }
}

function continueAttendanceSheetCreation(targetMonthName, targetYear) {
  var ui = SpreadsheetApp.getUi();
  
  try {
    UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "INFO", "Continuing attendance sheet creation", 
                                  targetMonthName + " " + targetYear, "");
    
    // Step 1: Create a date in the target month for semester lookup
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var monthIndex = monthNames.indexOf(targetMonthName);
    var targetDate = new Date(targetYear, monthIndex, 15); // Use middle of month
    
    // Step 2: Get semester for target date
    var semesterName = getSemesterForDate(targetDate);
    if (!semesterName) {
      ui.alert('⚠️ Semester Not Found',
               'Could not find a semester for ' + targetMonthName + ' ' + targetYear + '.\n\n' +
               'Please set up the semester first in Semester Metadata.');
      UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "ERROR", "No semester found for date", 
                                    targetDate.toDateString(), "");
      return;
    }
    
    UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "INFO", "Semester found", semesterName, "");
    
    // Step 3: Prompt for reconciliation confirmation
    var reconcileResponse = ui.alert(
      'Confirm Reconciliation Status',
      'Is reconciliation up to date for all teachers?\n\n' +
      '(Reconciliation updates the "Lessons Remaining" counts that determine which students get added to ' + targetMonthName + ')',
      ui.ButtonSet.YES_NO
    );
    
    if (reconcileResponse !== ui.Button.YES) {
      ui.alert('⚠️ Reconciliation Required',
               'Please reconcile all teacher rosters before creating attendance sheets.\n\n' +
               'Run the reconciliation process from the Billing menu first.');
      UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "INFO", "User indicated reconciliation not complete", "", "");
      return;
    }
    
    // Step 4: Get Active teachers
    var teacherLookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');

    if (!teacherLookupSheet) {
      throw new Error('Teacher Roster Lookup sheet not found');
    }

    var teacherData = teacherLookupSheet.getDataRange().getValues();
    var getCol = UtilityScriptLibrary.createColumnFinder(teacherLookupSheet);
    var nameCol = getCol('Display Name') - 1;
    var urlCol = getCol('Roster URL') - 1;
    var statusCol = getCol('Status') - 1;

    if (nameCol === -1 || urlCol === -1 || statusCol === -1) {
      throw new Error('Required columns not found in Enhanced Teacher Roster Lookup');
    }
    
    // Step 5: Process each Active teacher
    var stats = {
      processed: 0,
      created: 0,
      updated: 0,
      skipped: 0,
      errors: []
    };
    
    for (var i = 1; i < teacherData.length; i++) {
      var row = teacherData[i];
      var teacherName = row[nameCol];
      var rosterUrl = row[urlCol];
      var status = row[statusCol];
      
      if (status !== 'active' || !teacherName || !rosterUrl) {
        continue;
      }
      
      try {
        var result = processTeacherForNewAttendance(teacherName, rosterUrl, targetMonthName);
        stats.processed++;
        
        if (result.created) {
          stats.created++;
        } else if (result.updated) {
          stats.updated++;
        } else if (result.skipped) {
          stats.skipped++;
        }
        
      } catch (error) {
        stats.errors.push({
          teacher: teacherName,
          error: error.message
        });
        UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "ERROR", 
                                      "Error processing teacher", 
                                      teacherName, error.message);
      }
    }
    
    // Step 6: Log summary
    UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "SUCCESS", 
                                  "Attendance sheet creation completed", 
                                  "Processed: " + stats.processed + ", Created: " + stats.created + 
                                  ", Updated: " + stats.updated + ", Skipped: " + stats.skipped + 
                                  ", Errors: " + stats.errors.length, "");
    
    // Step 7: Show summary to user
    var summaryMessage = 'Attendance Sheet Creation Complete\n\n' +
                        'Month: ' + targetMonthName + ' ' + targetYear + '\n' +
                        'Teachers Processed: ' + stats.processed + '\n' +
                        'Sheets Created: ' + stats.created + '\n' +
                        'Sheets Updated: ' + stats.updated + '\n' +
                        'Teachers Skipped: ' + stats.skipped;
    
    if (stats.errors.length > 0) {
      summaryMessage += '\n\nErrors (' + stats.errors.length + '):';
      for (var i = 0; i < Math.min(stats.errors.length, 5); i++) {
        summaryMessage += '\n• ' + stats.errors[i].teacher + ': ' + stats.errors[i].error;
      }
      if (stats.errors.length > 5) {
        summaryMessage += '\n• ... and ' + (stats.errors.length - 5) + ' more';
      }
    }
    
    ui.alert('✅ Complete', summaryMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("continueAttendanceSheetCreation", "ERROR", "Fatal error", "", error.message);
    ui.alert('❌ Error', 'An error occurred:\n\n' + error.message, ui.ButtonSet.OK);
  }
}

function getStudentRegisteredLessonLength(studentId) {
  try {
    // Get current billing context to find the active billing sheet
    var billingSS = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get the most recent billing sheet (assuming it's the current one)
    var metadataSheet = billingSS.getSheetByName('Billing Metadata');
    if (!metadataSheet) {
      throw new Error('Billing Metadata sheet not found');
    }
    
    // RANGE VALIDATION: Check if metadata sheet has data
    var metadataLastRow = metadataSheet.getLastRow();
    if (metadataLastRow < 2) {
      UtilityScriptLibrary.debugLog('No billing cycles found in metadata - student ' + studentId + ' not found');
      return null;
    }
    
    var metadataData = metadataSheet.getDataRange().getValues();
    if (metadataData.length < 2) {
      throw new Error('No billing cycles found');
    }
    
    // Get the most recent billing cycle
    var latestBillingCycle = metadataData[metadataData.length - 1][0]; // Column A = Billing Month
    var currentBillingSheet = billingSS.getSheetByName(latestBillingCycle);
    
    if (!currentBillingSheet) {
      throw new Error('Current billing sheet not found: ' + latestBillingCycle);
    }
    
    // RANGE VALIDATION: Check if billing sheet has data
    var billingLastRow = currentBillingSheet.getLastRow();
    if (billingLastRow < 2) {
      UtilityScriptLibrary.debugLog('No data rows found in billing sheet - student ' + studentId + ' not found');
      return null;
    }
    
    // Find the student in the billing sheet
    var headerMap = UtilityScriptLibrary.getHeaderMap(currentBillingSheet);
    
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    var lessonLengthCol = headerMap[UtilityScriptLibrary.normalizeHeader('Lesson Length')];
    
    if (!studentIdCol || !lessonLengthCol) {
      throw new Error('Required columns not found in billing sheet');
    }
    
    var data = currentBillingSheet.getDataRange().getValues();
    
    // RANGE VALIDATION: Double-check data array
    if (!data || data.length < 2) {
      UtilityScriptLibrary.debugLog('No data rows to search in billing sheet for student ' + studentId);
      return null;
    }
    
    // Find the student row
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[studentIdCol - 1] === studentId) {
        var lessonLength = parseInt(row[lessonLengthCol - 1]) || 0;
        return lessonLength;
      }
    }
    
    // Student not found in current billing sheet
    UtilityScriptLibrary.debugLog('Student ' + studentId + ' not found in current billing sheet ' + latestBillingCycle);
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('Error getting registered lesson length for student ' + studentId + ': ' + error.message);
    return null;
  }
}

function getStudentDocumentsFolder(monthName) {
  try {
    var mainFolder = UtilityScriptLibrary.getGeneratedDocumentsFolder();
    var studentFolders = mainFolder.getFoldersByName("Student Documents");
    
    if (!studentFolders.hasNext()) {
      throw new Error("Student Documents folder not found. Please create it manually in the generated documents folder.");
    }
    
    var studentDocumentsFolder = studentFolders.next();
    
    // Create or get the monthly subfolder
    var monthlyFolders = studentDocumentsFolder.getFoldersByName(monthName);
    
    if (monthlyFolders.hasNext()) {
      UtilityScriptLibrary.debugLog("getStudentDocumentsFolder", "INFO", "Using existing monthly folder", 
                    "Month: " + monthName, "");
      return monthlyFolders.next();
    } else {
      var newMonthFolder = studentDocumentsFolder.createFolder(monthName);
      UtilityScriptLibrary.debugLog("getStudentDocumentsFolder", "INFO", "Created new monthly folder", 
                    "Month: " + monthName, "");
      return newMonthFolder;
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getStudentDocumentsFolder", "ERROR", "Failed to get student documents folder", 
                  "Month: " + monthName, error.message);
    throw error;
  }
}

function getStudentsNeedingPackets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getActiveSheet();
    
    var data = billingSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    
    var studentsNeedingPackets = [];
    
    // Process each student row (skip header)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var studentData = extractStudentDataFromBillingRow(row, headerMap);
      
      // Skip rows without student data
      if (!studentData.firstName || !studentData.lastName) {
        continue;
      }
      
      // Extract billing data for this student
      var billingData = extractBillingDataFromRow(row, headerMap);
      var currentSemester = getCurrentSemesterName();
      
      studentsNeedingPackets.push({
        studentData: studentData,
        billingData: billingData,
        currentSemester: currentSemester
      });
    }
    
    return studentsNeedingPackets;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getStudentsNeedingPackets", "ERROR", "Failed to get students", "", error.message);
    throw error;
  }
}

function isHeaderRow(row, dateIdx, lengthIdx) {
  var dateValue = row[dateIdx];
  var lengthValue = row[lengthIdx];
  
  // Check if date contains a month name instead of MM/DD format
  if (dateValue && typeof dateValue === 'string') {
    var monthNames = UtilityScriptLibrary.getMonthNames();
    var dateLower = dateValue.toLowerCase();
    for (var i = 0; i < monthNames.length; i++) {
      if (dateLower.indexOf(monthNames[i].toLowerCase()) !== -1) {
        return true;
      }
    }
  }
  
  // Check if length contains " minutes" suffix
  if (lengthValue && typeof lengthValue === 'string' && lengthValue.indexOf(' minutes') !== -1) {
    return true;
  }
  
  return false;
}

function promptForBillingCycleName(customToday) {
  var monthNames = UtilityScriptLibrary.getMonthNames();
  
  var defaultCycleName = monthNames[customToday.getMonth()] + ' ' + customToday.getFullYear();
  
  return UtilityScriptLibrary.promptForNameWithDefault({
    defaultValue: defaultCycleName,
    entityType: "billing cycle",
    promptTitle: "Enter Billing Cycle Name",
    promptMessage: "Please enter the name for this billing cycle (e.g., \"Summer 2025\"):",
    minLength: 3
  });
}

function promptForCustomToday() {
  var ui = SpreadsheetApp.getUi();
  var today = new Date();
  var defaultTodayStr = UtilityScriptLibrary.formatDateFlexible(today, "MM/dd/yyyy");

  var confirm = ui.alert(
    'Use today\'s date (' + defaultTodayStr + ') for this billing cycle?',
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.YES) {
    UtilityScriptLibrary.debugLog('Using today\'s date: ' + today);
    return today;
  }

  var response = ui.prompt(
    'Enter Custom Date',
    'Please enter the date to use as "today" (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Billing cycle setup cancelled.');
    throw new Error('User cancelled custom date entry.');
  }

  var parsedDate = UtilityScriptLibrary.parseDateFromString(response.getResponseText().trim());
  if (!parsedDate) {
    ui.alert('Invalid date. Please use MM/DD/YYYY.');
    throw new Error('Invalid date value.');
  }

  UtilityScriptLibrary.debugLog('Parsed custom date: ' + parsedDate);
  return parsedDate;
}

function promptForMonthAndYear() {
  var ui = SpreadsheetApp.getUi();
  var monthNames = UtilityScriptLibrary.getMonthNames();
  
  // Clear any previous selection
  PropertiesService.getScriptProperties().deleteProperty('monthYearSelection');
  
  // Create month dropdown HTML
  var monthOptions = '';
  for (var i = 0; i < monthNames.length; i++) {
    monthOptions += '<option value="' + monthNames[i] + '">' + monthNames[i] + '</option>';
  }
  
  // Create year dropdown HTML (current year and next 2 years)
  var currentYear = new Date().getFullYear();
  var yearOptions = '';
  for (var i = 0; i < 3; i++) {
    var year = currentYear + i;
    yearOptions += '<option value="' + year + '"' + (i === 0 ? ' selected' : '') + '>' + year + '</option>';
  }
  
  var html = '<style>' +
             'body { font-family: Arial, sans-serif; padding: 20px; }' +
             'label { display: block; margin-top: 15px; margin-bottom: 5px; font-weight: bold; }' +
             'select { width: 100%; padding: 8px; font-size: 14px; }' +
             '.button-container { margin-top: 20px; display: flex; gap: 10px; }' +
             'button { flex: 1; padding: 10px 20px; border: none; cursor: pointer; font-size: 14px; border-radius: 4px; }' +
             '#cancelBtn { background-color: #666; color: white; }' +
             '#cancelBtn:hover { background-color: #555; }' +
             '#okBtn { background-color: #4CAF50; color: white; }' +
             '#okBtn:hover { background-color: #45a049; }' +
             'button:disabled { background-color: #ccc; cursor: not-allowed; }' +
             '</style>' +
             '<form>' +
             '<label for="month">Select Month:</label>' +
             '<select id="month" name="month">' + monthOptions + '</select>' +
             '<label for="year">Select Year:</label>' +
             '<select id="year" name="year">' + yearOptions + '</select>' +
             '<div class="button-container">' +
             '<button type="button" id="cancelBtn" onclick="handleCancel()">Cancel</button>' +
             '<button type="button" id="okBtn" onclick="handleSubmit()">OK</button>' +
             '</div>' +
             '</form>' +
             '<script>' +
             'var isSubmitting = false;' +
             'function handleSubmit() {' +
             '  if (isSubmitting) return;' +
             '  isSubmitting = true;' +
             '  var okBtn = document.getElementById("okBtn");' +
             '  var cancelBtn = document.getElementById("cancelBtn");' +
             '  okBtn.disabled = true;' +
             '  cancelBtn.disabled = true;' +
             '  okBtn.textContent = "Processing...";' +
             '  var month = document.getElementById("month").value;' +
             '  var year = document.getElementById("year").value;' +
             '  google.script.run' +
             '    .withSuccessHandler(function() { google.script.host.close(); })' +
             '    .withFailureHandler(function(error) {' +
             '      alert("Error: " + error.message);' +
             '      okBtn.disabled = false;' +
             '      cancelBtn.disabled = false;' +
             '      okBtn.textContent = "OK";' +
             '      isSubmitting = false;' +
             '    })' +
             '    .handleMonthYearSelection(month, year);' +
             '}' +
             'function handleCancel() {' +
             '  google.script.run' +
             '    .withSuccessHandler(function() { google.script.host.close(); })' +
             '    .handleMonthYearCancel();' +
             '}' +
             '</script>';
  
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(350)
    .setHeight(220);
  
  ui.showModalDialog(htmlOutput, 'Select Month and Year for Attendance Sheets');
}

function handleMonthYearSelection(month, year) {
  try {
    UtilityScriptLibrary.debugLog("handleMonthYearSelection", "INFO", 
                                  "Month/year selected", 
                                  month + " " + year, "");
    
    // Continue with the attendance sheet creation process
    continueAttendanceSheetCreation(month, parseInt(year));
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("handleMonthYearSelection", "ERROR", 
                                  "Error in selection handler", 
                                  "", error.message);
    var ui = SpreadsheetApp.getUi();
    ui.alert('❌ Error', 'An error occurred:\n\n' + error.message, ui.ButtonSet.OK);
  }
}

function handleMonthYearCancel() {
  UtilityScriptLibrary.debugLog("handleMonthYearCancel", "INFO", "User cancelled month/year selection", "", "");
}

function promptForSemesterDates() {
  var result = UtilityScriptLibrary.executeWithErrorHandling(function() {
    UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Starting semester date prompts');
    
    var ui = SpreadsheetApp.getUi();
    
    // Prompt for start date
    var startPrompt = ui.prompt(
      'Semester Start Date', 
      'Enter the semester start date (MM/DD/YYYY):\nExample: 01/15/2024', 
      ui.ButtonSet.OK_CANCEL
    );
    
    if (startPrompt.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - User cancelled start date prompt');
      throw new Error('Semester setup cancelled.');
    }
    
    var startDateRaw = startPrompt.getResponseText();
    UtilityScriptLibrary.debugLog('📅 Raw start date input: "' + startDateRaw + '" (type: ' + typeof startDateRaw + ')');
    
    if (!startDateRaw) {
      var emptyStartMsg = 'Start date cannot be empty. Please enter a valid date.';
      ui.alert('Error', emptyStartMsg, ui.ButtonSet.OK);
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Empty start date input');
      throw new Error(emptyStartMsg);
    }
    
    var startDateText = startDateRaw.trim();
    UtilityScriptLibrary.debugLog('📅 Trimmed start date input: "' + startDateText + '" (length: ' + startDateText.length + ')');
    
    var startDate = UtilityScriptLibrary.parseDateFromString(startDateText);
    UtilityScriptLibrary.debugLog('📅 Parsed start date result: ' + startDate + ' (type: ' + typeof startDate + ')');
    
    if (!startDate) {
      var startErrorMsg = 'Invalid start date format. Please use MM/DD/YYYY format (e.g., 01/15/2024).';
      ui.alert('Error', startErrorMsg + '\n\nYou entered: "' + startDateText + '"', ui.ButtonSet.OK);
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Invalid start date: "' + startDateText + '"');
      throw new Error('Invalid startDate returned from prompt.');
    }
    
    UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Valid start date parsed: ' + startDate.toDateString());
    
    // Prompt for end date
    var endPrompt = ui.prompt(
      'Semester End Date', 
      'Enter the semester end date (MM/DD/YYYY):\nExample: 05/15/2024', 
      ui.ButtonSet.OK_CANCEL
    );
    
    if (endPrompt.getSelectedButton() !== ui.Button.OK) {
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - User cancelled end date prompt');
      throw new Error('Semester setup cancelled.');
    }
    
    var endDateRaw = endPrompt.getResponseText();
    UtilityScriptLibrary.debugLog('📅 Raw end date input: "' + endDateRaw + '" (type: ' + typeof endDateRaw + ')');
    
    if (!endDateRaw) {
      var emptyEndMsg = 'End date cannot be empty. Please enter a valid date.';
      ui.alert('Error', emptyEndMsg, ui.ButtonSet.OK);
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Empty end date input');
      throw new Error(emptyEndMsg);
    }
    
    var endDateText = endDateRaw.trim();
    UtilityScriptLibrary.debugLog('📅 Trimmed end date input: "' + endDateText + '" (length: ' + endDateText.length + ')');
    
    var endDate = UtilityScriptLibrary.parseDateFromString(endDateText);
    UtilityScriptLibrary.debugLog('📅 Parsed end date result: ' + endDate + ' (type: ' + typeof endDate + ')');
    
    if (!endDate) {
      var endErrorMsg = 'Invalid end date format. Please use MM/DD/YYYY format (e.g., 05/15/2024).';
      ui.alert('Error', endErrorMsg + '\n\nYou entered: "' + endDateText + '"', ui.ButtonSet.OK);
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Invalid end date: "' + endDateText + '"');
      throw new Error('Invalid endDate returned from prompt.');
    }
    
    UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Valid end date parsed: ' + endDate.toDateString());
    
    // Validate date range
    if (startDate.getTime() >= endDate.getTime()) {
      var rangeErrorMsg = 'Start date must be before end date. Please check your dates.';
      ui.alert('Error', rangeErrorMsg + '\n\nStart: ' + startDate.toDateString() + '\nEnd: ' + endDate.toDateString(), ui.ButtonSet.OK);
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Invalid date range: start=' + startDate.toDateString() + ', end=' + endDate.toDateString());
      throw new Error(rangeErrorMsg);
    }
    
    // Additional validation: check if dates are reasonable for a semester
    var timeDiff = endDate.getTime() - startDate.getTime();
    var daysDiff = Math.round(timeDiff / (1000 * 3600 * 24));
    
    if (daysDiff < 30) {
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Warning: Short semester duration: ' + daysDiff + ' days');
      ui.alert('Warning', 'Semester duration is only ' + daysDiff + ' days. This seems short for a typical semester.', ui.ButtonSet.OK);
    } else if (daysDiff > 365) {
      UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Warning: Long semester duration: ' + daysDiff + ' days');
      ui.alert('Warning', 'Semester duration is ' + daysDiff + ' days. This seems long for a typical semester.', ui.ButtonSet.OK);
    }
    
    UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - Valid date range confirmed: ' + startDate.toDateString() + ' to ' + endDate.toDateString() + ' (' + daysDiff + ' days)');
    
    return { 
      startDate: startDate, 
      endDate: endDate 
    };
    
  }, 'Semester dates collected successfully', 'promptForSemesterDates', {
    showUI: false,
    logLevel: 'INFO'
  });
  
  // Handle the executeWithErrorHandling response properly
  if (!result.success) {
    UtilityScriptLibrary.debugLog('📅 promptForSemesterDates - executeWithErrorHandling failed: ' + result.error);
    throw new Error(result.error);
  }
  
  return result.data;
}

function promptForSemesterName() {
  return UtilityScriptLibrary.executeWithErrorHandling(function() {
    UtilityScriptLibrary.debugLog('promptForSemesterName - Starting semester name prompt');
    
    return UtilityScriptLibrary.promptForNameWithDefault({
      defaultValue: null, // No default for semesters
      entityType: "semester",
      promptTitle: "New Semester Setup",
      promptMessage: "Enter the new semester name (e.g., \"Fall 2025\"):",
      minLength: 3
    });
    
  }, 'Semester name prompt completed', 'promptForSemesterName', {
    showUI: false,
    logLevel: 'DEBUG'
  }).data;
}

function protectBillingSheet(sheet) {
  /**
   * Protects an entire billing sheet with warning-only mode.
   * This prevents accidental manual edits while still allowing script-based updates.
   * Formulas are kept intact (not frozen to values).
   * 
   * @param {Sheet} sheet - The billing sheet to protect
   */
  try {
    // Remove any existing sheet-level protections first
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < protections.length; i++) {
      protections[i].remove();
    }
    
    // Protect the entire sheet
    var protection = sheet.protect();
    protection.setDescription('Protected billing cycle - formulas preserved');
    protection.setWarningOnly(true);
    
    UtilityScriptLibrary.debugLog("protectBillingSheet", "INFO", "Sheet protected successfully", 
                  "Sheet: " + sheet.getName(), "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("protectBillingSheet", "ERROR", "Failed to protect billing sheet", 
                  "Sheet: " + sheet.getName(), error.message);
    throw error;
  }
}

function storeSemesterEndBalances(creditBalances) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var balanceSheet = ss.getSheetByName('Semester Lesson Balances');
    
    // Create sheet if it doesn't exist
    if (!balanceSheet) {
      balanceSheet = ss.insertSheet('Semester Lesson Balances');
      
      // Set up headers
      var headers = [
        'Student ID',
        'Semester Name', 
        'Lesson Length',
        'Lessons Registered',
        'Lessons Taught',
        'Lesson Balance',
        'Status',
        'Date Processed'
      ];
      
      balanceSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      balanceSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // Add balance records
    var rowsToAdd = [];
    for (var i = 0; i < creditBalances.length; i++) {
      var balance = creditBalances[i];
      rowsToAdd.push([
        balance.studentId,
        balance.semesterName,
        balance.lessonLength,
        balance.registeredLessons,
        balance.taughtLessons,
        balance.lessonBalance,
        balance.status,
        new Date()
      ]);
    }
    
    if (rowsToAdd.length > 0) {
      var startRow = balanceSheet.getLastRow() + 1;
      balanceSheet.getRange(startRow, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }
    
    UtilityScriptLibrary.debugLog('✅ Stored ' + creditBalances.length + ' semester end balances');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('❌ Error storing semester end balances: ' + error.message);
    throw error;
  }
}

function verifyRatesEnhanced() {
  UtilityScriptLibrary.debugLog('verifyRatesEnhanced - Starting enhanced rates verification');
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rateSheet = ss.getSheetByName("Rates");
    
    if (!rateSheet) {
      var errorMsg = "Rates sheet not found";
      UtilityScriptLibrary.debugLog('verifyRatesEnhanced - ERROR: ' + errorMsg);
      throw new Error(errorMsg);
    }
    
    // Use UtilityScriptLibrary function instead of local helper
    var allData = rateSheet.getDataRange().getValues();
    var headers = allData[0];
    var rateColIndex = UtilityScriptLibrary.getMostRecentRateColumn(headers);
    
    if (rateColIndex === -1) {
      throw new Error("No valid rate column found");
    }
    
    var typeColIndex = headers.indexOf("Type");
    if (typeColIndex === -1) {
      throw new Error("Type column not found in Rates sheet");
    }
    
    var currentRateChart = headers[rateColIndex];
    UtilityScriptLibrary.debugLog('verifyRatesEnhanced - Using rate chart: ' + currentRateChart);
    
    // Build formatted summary
    var formattedEntries = [];
    
    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      var title = row[0];
      var value = row[rateColIndex];
      var type = row[typeColIndex];
      
      var formattedValue;
      
      if (type === "Currency") {
        var numValue = typeof value === "string" ? parseFloat(value) : value;
        formattedValue = UtilityScriptLibrary.formatCurrency(numValue);
      } else if (type === "Quantity") {
        var numValue = typeof value === "string" ? parseInt(value) : value;
        formattedValue = isNaN(numValue) ? "0" : numValue.toString();
      } else {
        formattedValue = value ? value.toString() : "0";
      }
      
      formattedEntries.push(title + ": " + formattedValue);
    }
    
    var summary = formattedEntries.join("\n");
    
    var ui = SpreadsheetApp.getUi();
    var alertMessage = "Current Rates (" + currentRateChart + "):\n\n" + summary + "\n\nDo you confirm these rates are correct for this billing cycle?";
    var response = ui.alert(alertMessage, ui.ButtonSet.YES_NO);
    
    var confirmed = response === ui.Button.YES;
    var verificationStatus = confirmed ? "confirmed" : "cancelled";
    
    UtilityScriptLibrary.debugLog('verifyRatesEnhanced - Rates verification ' + verificationStatus + ' for ' + currentRateChart);
    
    return confirmed;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('verifyRatesEnhanced - ERROR: ' + error.message);
    throw error;
  }
}

// ============================================================================
// SECTION 7: DEBUGGING AND TESTING FUNCTIONS
// ============================================================================

function testPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Test Prompt', 'Enter something:', ui.ButtonSet.OK_CANCEL);
  UtilityScriptLibrary.debugLog(response.getResponseText());
}

function grantDocumentPermissions() {
  try {
    // This will trigger the permission request for DocumentApp
    var testDoc = DocumentApp.create("Permission Test");
    UtilityScriptLibrary.debugLog("Created test doc: " + testDoc.getId());
    
    // Clean up
    DriveApp.getFileById(testDoc.getId()).setTrashed(true);
    
    SpreadsheetApp.getUi().alert("✅ Document permissions should now be granted!");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("Permission error: " + error.message);
    SpreadsheetApp.getUi().alert("❌ Permission error: " + error.message);
  }
}

function checkRowFormatting() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var lastCol = sheet.getLastColumn();
    
    // Get headers from row 1
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // Get the values and formats from row 2
    var values = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
    var formats = sheet.getRange(2, 1, 1, lastCol).getNumberFormats()[0];
    
    UtilityScriptLibrary.debugLog("checkRowFormatting", "INFO", "Starting column formatting analysis", 
                                  "Sheet: " + sheet.getName() + ", Total columns: " + lastCol, "");
    
    var currencyCount = 0;
    var hoursCount = 0;
    var decimalCount = 0;
    var integerCount = 0;
    var otherCount = 0;
    var problemColumns = [];
    
    for (var col = 0; col < lastCol; col++) {
      var header = headers[col] || "No Header";
      var value = values[col] || "";
      var format = formats[col] || "General";
      var colLetter = UtilityScriptLibrary.columnToLetter(col + 1);
      
      // Categorize the format
      var category = "";
      if (format.indexOf("$") !== -1) {
        category = "CURRENCY";
        currencyCount++;
      } else if (format === "0.00") {
        category = "DECIMAL";
        decimalCount++;
      } else if (format === "0") {
        category = "INTEGER";
        integerCount++;
      } else {
        category = "OTHER";
        otherCount++;
      }
      
      // Check if this is a hours column
      var isHoursColumn = header.toString().indexOf("Hours") !== -1;
      if (isHoursColumn) {
        hoursCount++;
        if (format.indexOf("$") !== -1) {
          category += " (HOURS COLUMN - PROBLEM!)";
          problemColumns.push({
            col: col + 1,
            header: header,
            format: format,
            letter: colLetter
          });
        } else {
          category += " (Hours Column - OK)";
        }
      }
      
      UtilityScriptLibrary.debugLog("checkRowFormatting", "DEBUG", "Column analysis", 
                                    "Col " + colLetter + " (" + (col + 1) + "): " + header, 
                                    "Value: " + value + " | Format: " + format + " | Type: " + category);
    }
    
    UtilityScriptLibrary.debugLog("checkRowFormatting", "INFO", "Formatting summary", 
                                  "Currency: " + currencyCount + ", Hours: " + hoursCount + ", Decimal: " + decimalCount, 
                                  "Integer: " + integerCount + ", Other: " + otherCount);
    
    // Log problematic formatting
    if (problemColumns.length > 0) {
      UtilityScriptLibrary.debugLog("checkRowFormatting", "ERROR", "Found hours columns with currency formatting", 
                                    "Problem columns: " + problemColumns.length, "");
      
      for (var i = 0; i < problemColumns.length; i++) {
        var prob = problemColumns[i];
        UtilityScriptLibrary.debugLog("checkRowFormatting", "ERROR", "Currency-formatted hours column", 
                                      "Col " + prob.letter + ": " + prob.header, 
                                      "Format: " + prob.format);
      }
    } else if (hoursCount > 0) {
      UtilityScriptLibrary.debugLog("checkRowFormatting", "INFO", "All hours columns correctly formatted", 
                                    "Hours columns found: " + hoursCount, "");
    }
    
    var alertMessage = "Formatting check complete. See Debug sheet for details.\n\n" +
                      "Summary:\n" +
                      "Currency: " + currencyCount + "\n" +
                      "Hours columns: " + hoursCount + "\n" +
                      "Decimal: " + decimalCount + "\n" +
                      "Integer: " + integerCount + "\n" +
                      "Other: " + otherCount;
    
    if (problemColumns.length > 0) {
      alertMessage += "\n\n🚨 FOUND " + problemColumns.length + " HOURS COLUMNS WITH CURRENCY FORMATTING!";
    }
    
    SpreadsheetApp.getUi().alert(alertMessage);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("checkRowFormatting", "ERROR", "Error in formatting check", 
                                  "", error.message + " | " + error.stack);
    SpreadsheetApp.getUi().alert("Error: " + error.message);
  }
}

function testFormatDateFunction() {
  try {
    console.log("=== Testing formatDate function ===");
    
    // Test with the same date that's failing in your system
    var testDate1 = new Date("Mon Jan 15 2024 00:00:00 GMT-0500 (Eastern Standard Time)");
    console.log("Test 1 - Date from log:");
    console.log("Input: " + testDate1);
    console.log("Result: '" + UtilityScriptLibrary.formatDateFlexible(testDate1, 'MMMM d, yyyy') + "'");
    console.log("Result (MMM): '" + UtilityScriptLibrary.formatDateFlexible(testDate1, 'MMM d, yyyy') + "'");
    
    // Test with a simple date construction
    var testDate2 = new Date(2024, 0, 15); // January 15, 2024
    console.log("\nTest 2 - Simple date construction:");
    console.log("Input: " + testDate2);
    console.log("Result: '" + UtilityScriptLibrary.formatDateFlexible(testDate2, 'MMMM d, yyyy') + "'");
    
    // Test with current date
    var testDate3 = new Date();
    console.log("\nTest 3 - Current date:");
    console.log("Input: " + testDate3);
    console.log("Result: '" + UtilityScriptLibrary.formatDateFlexible(testDate3, 'MMMM d, yyyy') + "'");
    
    // Test with string conversion
    var dateString = "Mon Jan 15 2024 00:00:00 GMT-0500 (Eastern Standard Time)";
    var testDate4 = new Date(dateString);
    console.log("\nTest 4 - String to Date conversion:");
    console.log("String: " + dateString);
    console.log("Date object: " + testDate4);
    console.log("Result: '" + UtilityScriptLibrary.formatDateFlexible(testDate4, 'MMMM d, yyyy') + "'");
    
    console.log("\n=== Test complete ===");
    
  } catch (error) {
    console.log("Error in test: " + error.message);
    console.log("Stack: " + error.stack);
  }
}

function testDynamicInvoiceFunctions() {
  try {
    // Get the active billing sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getActiveSheet();
    var data = billingSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    
    // Test with first data row (row 1 is header, row 2 is first student)
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog('ERROR: No student data found in billing sheet');
      return;
    }
    
    var testRow = data[1]; // First student row
    UtilityScriptLibrary.debugLog('=== TESTING DYNAMIC FUNCTIONS ===');
    UtilityScriptLibrary.debugLog('Using student row index 2 (first data row)');
    
    // Extract billing data using our fixed function
    var billingData = extractBillingDataFromRow(testRow, headerMap);
    
    UtilityScriptLibrary.debugLog('=== BILLING DATA EXTRACTED ===');
    UtilityScriptLibrary.debugLog('billingData.pastBalance: ' + billingData.pastBalance);
    UtilityScriptLibrary.debugLog('billingData.lateFee: ' + billingData.lateFee);
    UtilityScriptLibrary.debugLog('billingData.lessonLength: ' + billingData.lessonLength);
    UtilityScriptLibrary.debugLog('billingData.programTotals: ' + JSON.stringify(billingData.programTotals));
    
    // Test dynamic line items
    var lineItems = buildDynamicLineItems(billingData);
    UtilityScriptLibrary.debugLog('=== DYNAMIC LINE ITEMS RESULT ===');
    UtilityScriptLibrary.debugLog('"' + lineItems + '"');
    
    // Test dynamic amounts
    var amounts = buildDynamicAmounts(billingData);
    UtilityScriptLibrary.debugLog('=== DYNAMIC AMOUNTS RESULT ===');
    UtilityScriptLibrary.debugLog('"' + amounts + '"');
    
    UtilityScriptLibrary.debugLog('=== TEST COMPLETED ===');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('ERROR in testDynamicInvoiceFunctions: ' + error.message);
    UtilityScriptLibrary.debugLog('Error stack: ' + error.stack);
  }
}

function debugBillingSheetColumns() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getActiveSheet();
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    var data = billingSheet.getDataRange().getValues();
    
    UtilityScriptLibrary.debugLog('=== BILLING SHEET COLUMN ANALYSIS ===');
    UtilityScriptLibrary.debugLog('Total columns: ' + Object.keys(headerMap).length);
    UtilityScriptLibrary.debugLog('');
    
    UtilityScriptLibrary.debugLog('=== ALL NORMALIZED HEADERS ===');
    var sortedHeaders = Object.keys(headerMap).sort();
    for (var i = 0; i < sortedHeaders.length; i++) {
      var header = sortedHeaders[i];
      var colIndex = headerMap[header];
      UtilityScriptLibrary.debugLog('Column ' + colIndex + ': "' + header + '"');
    }
    
    UtilityScriptLibrary.debugLog('');
    UtilityScriptLibrary.debugLog('=== HEADERS CONTAINING "total" ===');
    var totalHeaders = [];
    for (var header in headerMap) {
      if (header.indexOf('total') !== -1) {
        totalHeaders.push(header);
        var colIndex = headerMap[header];
        var sampleValue = data.length > 1 ? data[1][colIndex - 1] : 'N/A';
        UtilityScriptLibrary.debugLog('Found total column: "' + header + '" (Column ' + colIndex + ') Sample value: ' + sampleValue);
      }
    }
    
    if (totalHeaders.length === 0) {
      UtilityScriptLibrary.debugLog('❌ NO TOTAL COLUMNS FOUND! This explains why programTotals.programs is empty.');
    }
    
    UtilityScriptLibrary.debugLog('');
    UtilityScriptLibrary.debugLog('=== SAMPLE DATA ROW VALUES ===');
    if (data.length > 1) {
      var sampleRow = data[1];
      UtilityScriptLibrary.debugLog('Row 2 (first student) has ' + sampleRow.length + ' columns');
      for (var j = 0; j < Math.min(sampleRow.length, 20); j++) {
        UtilityScriptLibrary.debugLog('Column ' + (j + 1) + ': "' + sampleRow[j] + '"');
      }
    }
    
    UtilityScriptLibrary.debugLog('=== ANALYSIS COMPLETED ===');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('ERROR in debugBillingSheetColumns: ' + error.message);
    UtilityScriptLibrary.debugLog('Error stack: ' + error.stack);
  }
}

function testDynamicInvoiceFunctionsOnCorrectSheet() {
  try {
    // Get all sheets and look for the billing sheet with student data
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    
    UtilityScriptLibrary.debugLog('=== SEARCHING FOR CORRECT BILLING SHEET ===');
    UtilityScriptLibrary.debugLog('Total sheets: ' + sheets.length);
    
    var billingSheet = null;
    
    // Look for a sheet with student billing data (should have columns like Student First Name, etc.)
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      
      try {
        var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
        var hasStudentName = false;
        var hasProgramTotals = false;
        
        // Check for expected billing sheet columns
        for (var header in headerMap) {
          if (header.indexOf('studentfirstname') !== -1 || header.indexOf('studentlastname') !== -1) {
            hasStudentName = true;
          }
          if (header.indexOf('total') !== -1 && header !== 'currentinvoicetotal') {
            hasProgramTotals = true;
          }
        }
        
        UtilityScriptLibrary.debugLog('Sheet "' + sheetName + '": hasStudentName=' + hasStudentName + ', hasProgramTotals=' + hasProgramTotals);
        
        if (hasStudentName && hasProgramTotals) {
          billingSheet = sheet;
          UtilityScriptLibrary.debugLog('✅ Found billing sheet: "' + sheetName + '"');
          break;
        }
        
      } catch (error) {
        UtilityScriptLibrary.debugLog('Error checking sheet "' + sheetName + '": ' + error.message);
      }
    }
    
    if (!billingSheet) {
      UtilityScriptLibrary.debugLog('❌ No billing sheet found with student data and program totals');
      UtilityScriptLibrary.debugLog('Available sheets: ' + sheets.map(function(s) { return s.getName(); }).join(', '));
      return;
    }
    
    var data = billingSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    
    // Test with first data row (row 1 is header, row 2 is first student)
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog('ERROR: No student data found in billing sheet');
      return;
    }
    
    var testRow = data[1]; // First student row
    UtilityScriptLibrary.debugLog('=== TESTING DYNAMIC FUNCTIONS ON CORRECT SHEET ===');
    UtilityScriptLibrary.debugLog('Using billing sheet: ' + billingSheet.getName());
    UtilityScriptLibrary.debugLog('Using student row index 2 (first data row)');
    
    // Extract billing data using our fixed function
    var billingData = extractBillingDataFromRow(testRow, headerMap);
    
    UtilityScriptLibrary.debugLog('=== BILLING DATA EXTRACTED ===');
    UtilityScriptLibrary.debugLog('billingData.pastBalance: ' + billingData.pastBalance);
    UtilityScriptLibrary.debugLog('billingData.lateFee: ' + billingData.lateFee);
    UtilityScriptLibrary.debugLog('billingData.lessonLength: ' + billingData.lessonLength);
    UtilityScriptLibrary.debugLog('billingData.programTotals: ' + JSON.stringify(billingData.programTotals));
    
    // Test dynamic line items
    var lineItems = buildDynamicLineItems(billingData);
    UtilityScriptLibrary.debugLog('=== DYNAMIC LINE ITEMS RESULT ===');
    UtilityScriptLibrary.debugLog('"' + lineItems + '"');
    
    // Test dynamic amounts
    var amounts = buildDynamicAmounts(billingData);
    UtilityScriptLibrary.debugLog('=== DYNAMIC AMOUNTS RESULT ===');
    UtilityScriptLibrary.debugLog('"' + amounts + '"');
    
    UtilityScriptLibrary.debugLog('=== TEST COMPLETED ===');
    
  } catch (error) {
    UtilityScriptLibrary.debugLog('ERROR in testDynamicInvoiceFunctionsOnCorrectSheet: ' + error.message);
    UtilityScriptLibrary.debugLog('Error stack: ' + error.stack);
  }
}

function testExtractStudentDataFromBillingRow() {
  try {
    // Get the active billing sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    UtilityScriptLibrary.debugLog("testExtractStudentDataFromBillingRow", "INFO", "Starting test", 
                  "Sheet: " + sheet.getName(), "");
    
    // Get the header map
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    // Log all headers for debugging
    var headerList = [];
    for (var header in headerMap) {
      headerList.push(header + ": col " + headerMap[header]);
    }
    UtilityScriptLibrary.debugLog("testExtractStudentDataFromBillingRow", "DEBUG", "Header map", 
                  headerList.join(", "), "");
    
    // Get the first data row (row 2)
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      UtilityScriptLibrary.debugLog("testExtractStudentDataFromBillingRow", "ERROR", "No data rows found", "", "");
      SpreadsheetApp.getUi().alert("No data rows found in sheet");
      return;
    }
    
    var testRow = data[1]; // First data row (index 1, since 0 is headers)
    
    // Log the raw row data
    UtilityScriptLibrary.debugLog("testExtractStudentDataFromBillingRow", "DEBUG", "Test row data", 
                  "Row length: " + testRow.length + ", First few values: " + testRow.slice(0, 5).join(", "), "");
    
    // Call the function
    var studentData = extractStudentDataFromBillingRow(testRow, headerMap);
    
    // Log the results
    UtilityScriptLibrary.debugLog("testExtractStudentDataFromBillingRow", "INFO", "Extraction complete", 
                  "Student: " + studentData.firstName + " " + studentData.lastName + 
                  ", LsnQty: '" + studentData.lessonQuantity + "', LsnLength: '" + studentData.lessonLength + "'", "");
    
    // Show results in alert
    var message = "Student Data Extracted:\n\n";
    message += "Name: " + studentData.firstName + " " + studentData.lastName + "\n";
    message += "ID: " + studentData.studentId + "\n";
    message += "Instrument: " + studentData.instrument + "\n";
    message += "Teacher: " + studentData.teacher + "\n";
    message += "Lesson Length: " + studentData.lessonLength + "\n";
    message += "Lesson Quantity: " + studentData.lessonQuantity + "\n\n";
    message += "Check Debug sheet for detailed logs.";
    
    SpreadsheetApp.getUi().alert("Test Complete", message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("testExtractStudentDataFromBillingRow", "ERROR", "Test failed", 
                  "", error.message + " | " + error.stack);
    SpreadsheetApp.getUi().alert("Test Error: " + error.message);
  }
}

function testFullAgreementGeneration() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getActiveSheet();
    
    UtilityScriptLibrary.debugLog("testFullAgreementGeneration", "INFO", "Starting full agreement test", 
                  "Sheet: " + billingSheet.getName(), "");
    
    // Get header map and first data row
    var headerMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    var data = billingSheet.getDataRange().getValues();
    var testRow = data[1];
    
    // Extract student data (this works, we confirmed)
    var studentData = extractStudentDataFromBillingRow(testRow, headerMap);
    UtilityScriptLibrary.debugLog("testFullAgreementGeneration", "DEBUG", "Student data extracted", 
                  "LsnQty: '" + studentData.lessonQuantity + "'", "");
    
    // Extract billing data
    var billingData = extractBillingDataFromRow(testRow, headerMap);
    UtilityScriptLibrary.debugLog("testFullAgreementGeneration", "DEBUG", "Billing data extracted", 
                  "Parent: " + billingData.parentFirstName + " " + billingData.parentLastName, "");
    
    // Build template variables (THIS is where we need to check)
    var variables = buildTemplateVariables(studentData, billingData, 'agreement');
    
    // Log ALL variables
    var varList = [];
    for (var varName in variables) {
      varList.push(varName + ": '" + variables[varName] + "'");
    }
    UtilityScriptLibrary.debugLog("testFullAgreementGeneration", "DEBUG", "Template variables built", 
                  varList.join(", "), "");
    
    // Check specifically for LsnQuantity
    UtilityScriptLibrary.debugLog("testFullAgreementGeneration", "INFO", "Final check", 
                  "LsnQuantity in variables object: '" + variables.LsnQuantity + "'", "");
    
    // Show results
    var message = "Template Variables Test:\n\n";
    message += "Student: " + studentData.firstName + " " + studentData.lastName + "\n";
    message += "LsnQuantity in studentData: '" + studentData.lessonQuantity + "'\n";
    message += "LsnQuantity in variables: '" + variables.LsnQuantity + "'\n\n";
    message += "Total variables: " + Object.keys(variables).length + "\n\n";
    message += "Check Debug sheet for full variable list.";
    
    SpreadsheetApp.getUi().alert("Test Complete", message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("testFullAgreementGeneration", "ERROR", "Test failed", 
                  "", error.message + " | " + error.stack);
    SpreadsheetApp.getUi().alert("Test Error: " + error.message);
  }
}

function testTemplateLiteral() {
  const name = "Bob";
  const msg = `Hello, ${name}`;
  UtilityScriptLibrary.debugLog(msg);
}

// ============================================================================
// MAIN BUILD FUNCTIONS
// ============================================================================

function buildBillingRowFromForm(formRow, prevRow, context, rowIndex) {
  var get = context.getColIndex;
  var studentId = formRow[get("Student ID")];
  var newRow = new Array(Object.keys(context.headerMap).length).fill("");
  var quantityCols = [];
  var currencyCols = [];

  UtilityScriptLibrary.debugLog('buildBillingRowFromForm', 'DEBUG', 'Building complete billing row', 
                                'Student ID: ' + studentId + ', Has prev: ' + !!prevRow, '');

  // Copy static fields from form
  copyStaticFieldsToBillingRow(newRow, formRow, context, get);

  // Delivery preference from form
  setDeliveryPreference(newRow, formRow, context, true);

  // Invoice metadata
  populateInvoiceMetadata(newRow, studentId, context, rowIndex);
  
  // Letter Type
  populateLetterType(newRow, context, "form", null);

  // Enrolled programs
  var enrolledPrograms = getExpandedPrograms(formRow, context);

  // Get package quantities using Utility helper
  var qty30Package = get("Qty30") !== -1 ? formRow[get("Qty30")] : "";
  var qty45Package = get("Qty45") !== -1 ? formRow[get("Qty45")] : "";
  var qty60Package = get("Qty60") !== -1 ? formRow[get("Qty60")] : "";
  
  var quantities = UtilityScriptLibrary.parseAllPackageQuantities(qty30Package, qty45Package, qty60Package);
  var lessonLength = UtilityScriptLibrary.getLessonLengthFromPackages(qty30Package, qty45Package, qty60Package);

  // Apply lesson credits if available
  var creditApplication = context.lessonCredits && context.lessonCredits[studentId] 
    ? context.lessonCredits[studentId] : null;
  
  if (creditApplication && creditApplication.totalCredits > 0) {
    var result = buildDynamicProgramColumnsWithCredits(
      newRow, formRow, enrolledPrograms, context,
      quantityCols, currencyCols, rowIndex, creditApplication
    );
    quantityCols = result.quantityCols;
    currencyCols = result.currencyCols;

    // Update Admin Comments
    var adminCommentsCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Admin Comments")];
    if (adminCommentsCol) {
      var existingComments = newRow[adminCommentsCol - 1] || '';
      var creditNote = "Credits applied: " + creditApplication.explanation;
      newRow[adminCommentsCol - 1] = existingComments ? existingComments + "; " + creditNote : creditNote;
    }
  } else {
    var result = buildDynamicProgramColumns(
      newRow, formRow, enrolledPrograms, context,
      quantityCols, currencyCols, rowIndex
    );
    quantityCols = result.quantityCols;
    currencyCols = result.currencyCols;
  }

  // Set lesson length
  var lengthCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Lesson Length")];
  if (lengthCol) {
    newRow[lengthCol - 1] = lessonLength;
  }

  // Timestamp
  var timestampCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Timestamp")];
  var formTimestampCol = get("Timestamp");
  if (timestampCol && formTimestampCol !== -1) {
    newRow[timestampCol - 1] = formRow[formTimestampCol];
  }

  // Payment Received = 0
  var paymentCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Payment Received")];
  if (paymentCol) {
    newRow[paymentCol - 1] = 0;
  }

    // Apply past data if student existed in previous cycle
  if (prevRow) {
    applyPastDataToRow(newRow, prevRow, context, currencyCols);
    applyLateFeeToRow(newRow, context, currencyCols);
  } else {
    // NEW: If no prevRow (student wasn't carried over), check cumulative tracking
    var studentId = get("Student ID") !== -1 ? formRow[get("Student ID")] : null;
    if (studentId) {
      applyCumulativeHistory(newRow, studentId, context, currencyCols);
    }
  }

  // Formulas
  addCumulativeFormulas(newRow, context, rowIndex, quantityCols);
  addInvoiceTotalFormula(newRow, context, rowIndex, currencyCols);
  populateCurrentBalanceFormula(newRow, context, rowIndex);

  return { newRow: newRow, quantityCols: quantityCols, currencyCols: currencyCols };
}

function buildBillingRowFromPrevious(prevRow, context, rowIndex) {
  UtilityScriptLibrary.debugLog("buildBillingRowFromPrevious", "DEBUG", "Building complete row from previous", 
                "Row: " + rowIndex, "");
  
  if (!context || !context.headerMap) {
    throw new Error("Context or headerMap is missing");
  }
  
  var newRow = new Array(Object.keys(context.headerMap).length).fill("");
  var quantityCols = [];
  var currencyCols = [];
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  // Get Student ID
  var studentId = UtilityScriptLibrary.getStudentIdFromRow(prevRow, context.prevHeaderMap);
  if (!studentId) {
    throw new Error("Student ID not found in previous billing row");
  }

  // Copy static fields
  copyStaticFieldsToBillingRow(newRow, prevRow, context);
  
  // Delivery preference from previous
  setDeliveryPreference(newRow, prevRow, context, false);
  
  // Invoice metadata
  populateInvoiceMetadata(newRow, studentId, context, rowIndex);

  // Copy lesson length from previous
  copyPreviousColumnToNew(newRow, prevRow, context, context.headerMap, context.prevHeaderMap, 
                          { newCol: "Lesson Length", prevCol: "Lesson Length" });

  // Set all program quantities to 0 with current rates
  setProgramQuantitiesForCarryover(newRow, context, quantityCols, currencyCols);

  // Generate program formulas
  generateProgramFormulas(newRow, context, rowIndex, quantityCols, currencyCols);

  // Apply past data
  applyPastDataToRow(newRow, prevRow, context, currencyCols);
  applyLateFeeToRow(newRow, context, currencyCols);

  // Formulas
  addInvoiceTotalFormula(newRow, context, rowIndex, currencyCols);
  populateCurrentBalanceFormula(newRow, context, rowIndex);
  
  // Letter Type
  populateLetterType(newRow, context, "previous", prevRow);

  return { 
    newRow: newRow, 
    quantityCols: quantityCols, 
    currencyCols: currencyCols 
  };
}

function populateBillingSheet(context, carryOverData) {
  UtilityScriptLibrary.debugLog("populateBillingSheet", "INFO", "Starting billing sheet population", "", "");
  
  try {
    var formData = context.formSheet.getDataRange().getValues();
    var get = context.getColIndex;
    var timestampCol = get("Timestamp");
    var studentIdCol = get("Student ID");
    var billingSheet = context.billingSheet;
    var existingIds = {};
    var prevDataMap = (carryOverData && carryOverData.rowsToCarry) ? carryOverData.rowsToCarry : {};
    var allStudents = [];

    // Build all form students with correct row indices
    var currentRowIndex = 2; // Start at row 2 (row 1 is header)
    for (var i = 1; i < formData.length; i++) {
      var row = formData[i];
      var studentId = row[studentIdCol];
      var timestamp = new Date(row[timestampCol]);

      if (!studentId || isNaN(timestamp.getTime())) continue;

      var prevRow = prevDataMap[studentId];
      var result = buildBillingRowFromForm(row, prevRow, context, currentRowIndex);
      
      allStudents.push(result);
      existingIds[studentId] = true;
      currentRowIndex++; // Increment for next student
    }

    // Build carryover students with correct row indices
    var carryoverCount = 0;
    for (var studentId in prevDataMap) {
      if (existingIds[studentId]) continue;

      var prevRow = prevDataMap[studentId];
      if (!prevRow || !Array.isArray(prevRow)) continue;

      var result = buildBillingRowFromPrevious(prevRow, context, currentRowIndex);
      
      allStudents.push(result);
      carryoverCount++;
      currentRowIndex++; // Increment for next student
    }

    UtilityScriptLibrary.debugLog("populateBillingSheet", "INFO", "Built all students", 
                  "Form: " + Object.keys(existingIds).length + ", Carryover: " + carryoverCount, "");
    
    // Write and format in batch
    writeAndFormatRows(billingSheet, allStudents);
    
    // Cumulative columns
    populateAllCumulativeColumns();
    
    // Overages
    if (carryOverData && carryOverData.previousSheetName) {
      var overages = detectAndBillOverages(billingSheet, context.billingCycleName);
      if (overages.length > 0) {
        UtilityScriptLibrary.debugLog("populateBillingSheet", "INFO", "Overages detected", 
                      "Count: " + overages.length, "");
      }
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateBillingSheet", "ERROR", "Failed", "", error.message);
    throw error;
  }
}

function populateBillingSheetContinuingSemester(context, billingSheet, existingStudentIds, previousData, previousStartDate) {
  try {
    UtilityScriptLibrary.debugLog("populateBillingSheetContinuingSemester", "INFO", "Starting", "", "");
    
    var filterDate = getFilterDate(context, previousStartDate);
    var formSheet = getFormSheet(context);
    var allStudents = [];
    
    // Build form students starting at row 2
    var formResult = buildFormStudents(formSheet, filterDate, previousData, context, allStudents, 2);
    existingStudentIds.push.apply(existingStudentIds, formResult.formStudentIds);
    
    // Build carryover students continuing from where form students left off
    buildCarryoverStudents(previousData, existingStudentIds, context, allStudents, formResult.nextRowIndex);
    
    UtilityScriptLibrary.debugLog("populateBillingSheetContinuingSemester", "INFO", "Built all students", 
                  "Total: " + allStudents.length, "");
    
    // Write and format in batch
    writeAndFormatRows(billingSheet, allStudents);
    
    // Cumulative columns
    populateAllCumulativeColumns();
    
    // Overages
    if (previousData && Object.keys(previousData).length > 0) {
      var overages = detectAndBillOverages(billingSheet, context.billingCycleName);
      if (overages.length > 0) {
        UtilityScriptLibrary.debugLog("populateBillingSheetContinuingSemester", "INFO", "Overages detected", 
                      "Count: " + overages.length, "");
      }
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("populateBillingSheetContinuingSemester", "ERROR", "Failed", "", error.message);
    throw error;
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function addCumulativeFormulas(newRow, context, rowIndex, quantityCols) {
  var h = context.headerMap;
  var norm = UtilityScriptLibrary.normalizeHeader;
  var colToLetter = UtilityScriptLibrary.columnToLetter;
  
  // Cumulative Billed formula
  var cumBilledCol = h[norm("Current Cumulative Hours Billed")];
  var lessonHoursCol = h[norm("Lesson Hours")];
  
  if (cumBilledCol && lessonHoursCol) {
    newRow[cumBilledCol - 1] = "=" + colToLetter(lessonHoursCol) + rowIndex;
    quantityCols.push(cumBilledCol);
  }

  // Hours Remaining formula
  var hoursRemainingCol = h[norm("Hours Remaining")];
  var currentCumTaughtCol = h[norm("Current Cumulative Hours Taught")];
  
  if (hoursRemainingCol && currentCumTaughtCol && cumBilledCol) {
    var formula = "=" + colToLetter(cumBilledCol) + rowIndex + " - " + colToLetter(currentCumTaughtCol) + rowIndex;
    newRow[hoursRemainingCol - 1] = formula;
    quantityCols.push(hoursRemainingCol);
  }
}

function addInvoiceTotalFormula(newRow, context, rowIndex, currencyCols) {
  var invoiceTotalCol = context.headerMap[UtilityScriptLibrary.normalizeHeader("Current Invoice Total")];
  if (invoiceTotalCol) {
    newRow[invoiceTotalCol - 1] = buildInvoiceTotalFormula(context.headerMap, rowIndex);
    UtilityScriptLibrary.addToCurrencyCols(currencyCols, invoiceTotalCol, "Current Invoice Total");
  }
}

function addMissingStudentsToAttendanceSheet(sheet, activeStudents) {
  try {
    // Get existing data from sheet
    var data = sheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    
    var studentIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Student ID')];
    
    if (!studentIdCol) {
      throw new Error('Student ID column not found in attendance sheet');
    }
    
    // Build set of existing student IDs
    var existingStudentIds = {};
    for (var i = 1; i < data.length; i++) {
      var studentId = data[i][studentIdCol - 1];
      if (studentId) {
        existingStudentIds[studentId] = true;
      }
    }
    
    // Find students who need to be added
    var studentsToAdd = [];
    for (var i = 0; i < activeStudents.length; i++) {
      var student = activeStudents[i];
      if (!existingStudentIds[student.id]) {
        studentsToAdd.push(student);
      }
    }
    
    if (studentsToAdd.length === 0) {
      UtilityScriptLibrary.debugLog("addMissingStudentsToAttendanceSheet", "INFO", 
                                    "No missing students to add", 
                                    sheet.getName(), "");
      return 0;
    }
    
    // Use the Utility library function to add student sections in the proper format
    UtilityScriptLibrary.createStudentSections(sheet, studentsToAdd);
    
    UtilityScriptLibrary.debugLog("addMissingStudentsToAttendanceSheet", "SUCCESS", 
                                  "Added missing students", 
                                  sheet.getName() + " (" + studentsToAdd.length + " students)", "");
    
    return studentsToAdd.length;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("addMissingStudentsToAttendanceSheet", "ERROR", 
                                  "Error adding missing students", 
                                  sheet.getName(), error.message);
    throw error;
  }
}

function applyCumulativeHistory(newRow, studentId, context, currencyCols) {
  var cumulativeHistory = getCumulativeHistory(studentId);
  
  if (!cumulativeHistory) {
    UtilityScriptLibrary.debugLog("applyCumulativeHistory", "DEBUG", "No cumulative history found", 
                  "Student ID: " + studentId, "");
    return;
  }
  
  var h = context.headerMap;
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  var pastCumTaughtCol = h[norm("Past Cumulative Hours Taught")];
  var pastCumBilledCol = h[norm("Past Cumulative Hours Billed")];
  
  if (pastCumTaughtCol) {
    newRow[pastCumTaughtCol - 1] = cumulativeHistory.cumulativeHoursTaught;
  }
  
  if (pastCumBilledCol) {
    newRow[pastCumBilledCol - 1] = cumulativeHistory.cumulativeHoursBilled;
  }
  
  UtilityScriptLibrary.debugLog("applyCumulativeHistory", "INFO", "Applied cumulative history", 
                "Student ID: " + studentId + ", Taught: " + cumulativeHistory.cumulativeHoursTaught + 
                ", Billed: " + cumulativeHistory.cumulativeHoursBilled, "");
}

function applyLateFeeToRow(newRow, context, currencyCols) {
  var h = context.headerMap;
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  var lateFeeCol = h[norm("Late Fee")];
  var pastBalanceCol = h[norm("Past Balance")];
  var paymentReceivedCol = h[norm("Payment Received")];
  var pastInvoiceNumberCol = h[norm("Past Invoice Number")];
  
  if (!lateFeeCol || !pastBalanceCol || !paymentReceivedCol || !pastInvoiceNumberCol) {
    return;
  }
  
  var pastBalance = UtilityScriptLibrary.safeParseFloat(newRow[pastBalanceCol - 1]);
  var paymentReceived = UtilityScriptLibrary.safeParseFloat(newRow[paymentReceivedCol - 1]);
  var pastInvoiceNumber = newRow[pastInvoiceNumberCol - 1] || '';
  
  var rateMap = getRateMap(context);
  var lateFeeAmount = calculateLateFee(pastBalance, paymentReceived, pastInvoiceNumber, rateMap);
  
  if (lateFeeAmount > 0) {
    newRow[lateFeeCol - 1] = lateFeeAmount;
    UtilityScriptLibrary.addToCurrencyCols(currencyCols, lateFeeCol, "Late Fee");
  }
}

function applyPastDataToRow(newRow, prevRow, context, currencyCols) {
  var h = context.headerMap;
  var prevH = context.prevHeaderMap;
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  // Define all past data mappings
  var mappings = [
    { newCol: "Past Invoice Number", prevCol: "Invoice Number", type: "string" },
    { newCol: "Past Cumulative Hours Taught", prevCol: "Current Cumulative Hours Taught", type: "number" },
    { newCol: "Past Cumulative Hours Billed", prevCol: "Current Cumulative Hours Billed", type: "number" },
    { newCol: "Past Hours Remaining", prevCol: "Hours Remaining", type: "number" }
  ];
  
  // Copy all mapped fields
  for (var i = 0; i < mappings.length; i++) {
    var mapping = mappings[i];
    var newCol = h[norm(mapping.newCol)];
    var prevCol = prevH[norm(mapping.prevCol)];
    
    if (newCol && prevCol && prevRow.length >= prevCol) {
      if (mapping.type === "number") {
        newRow[newCol - 1] = UtilityScriptLibrary.safeParseFloat(prevRow[prevCol - 1]);
      } else {
        newRow[newCol - 1] = prevRow[prevCol - 1] || '';
      }
    }
  }
  
  // Handle Past Balance / Credit (special case - can be positive or negative)
  var pastBalanceCol = h[norm("Past Balance")];
  var creditCol = h[norm("Credit")];
  var currentBalanceCol = prevH[norm("Current Balance")];
  
  if (currentBalanceCol && prevRow.length >= currentBalanceCol) {
    var previousBalance = UtilityScriptLibrary.safeParseFloat(prevRow[currentBalanceCol - 1]);
    
    if (previousBalance !== 0 && pastBalanceCol) {
      newRow[pastBalanceCol - 1] = previousBalance; // Keep the sign (+/-)
      UtilityScriptLibrary.addToCurrencyCols(currencyCols, pastBalanceCol, "Past Balance");
    }
  }
  
  // Calculate Lesson Credit if applicable
  var lessonCreditCol = h[norm("Lesson Credit")];
  var lessonQuantityCol = h[norm("Lesson Quantity")];
  var lessonLengthCol = h[norm("Lesson Length")];
  var pastHoursRemainingCol = h[norm("Past Hours Remaining")];
  
  if (lessonCreditCol && lessonQuantityCol && lessonLengthCol && pastHoursRemainingCol) {
    var pastHoursRemaining = UtilityScriptLibrary.safeParseFloat(newRow[pastHoursRemainingCol - 1]);
    var lessonQuantity = UtilityScriptLibrary.safeParseFloat(newRow[lessonQuantityCol - 1]);
    var lessonLength = UtilityScriptLibrary.safeParseFloat(newRow[lessonLengthCol - 1]);
    
    if (pastHoursRemaining > 0 && lessonQuantity > 0 && lessonLength > 0) {
      newRow[lessonCreditCol - 1] = pastHoursRemaining / (lessonLength / 60);
    } else {
      newRow[lessonCreditCol - 1] = 0;
    }
  }
}

function buildCarryoverStudents(previousData, existingStudentIds, context, allStudents, startingRowIndex) {
  var currentRowIndex = startingRowIndex;
  
  for (var studentId in previousData) {
    if (!previousData.hasOwnProperty(studentId)) continue;
    if (existingStudentIds.indexOf(studentId) !== -1) continue;
    
    var prevRow = previousData[studentId];
    var result = buildBillingRowFromPrevious(prevRow, context, currentRowIndex);
    
    allStudents.push(result);
    currentRowIndex++;
  }
}

function buildFormStudents(formSheet, filterDate, previousData, context, allStudents, startingRowIndex) {
  var data = formSheet.getDataRange().getValues();
  var formHeaderMap = UtilityScriptLibrary.getHeaderMap(formSheet);
  var norm = UtilityScriptLibrary.normalizeHeader;
  var timestampCol = formHeaderMap[norm("Timestamp")];
  var studentIdCol = formHeaderMap[norm("Student ID")];
  var formStudentIds = [];
  var currentRowIndex = startingRowIndex;
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = row[timestampCol - 1];
    var studentId = row[studentIdCol - 1];
    
    if (studentId && (!filterDate || timestamp >= filterDate)) {
      var prevRow = previousData[studentId];
      var result = buildBillingRowFromForm(row, prevRow, context, currentRowIndex);
      
      allStudents.push(result);
      formStudentIds.push(studentId);
      currentRowIndex++;
    }
  }
  
  return { formStudentIds: formStudentIds, nextRowIndex: currentRowIndex };
}

function calculateLateFee(pastBalance, paymentReceived, pastInvoiceNumber, rateMap) {
  var lessonRate = parseFloat(rateMap["Lessons"]) || 0;
  var gracePeriod = parseInt(rateMap["Grace Period"]) || 10;
  
  if (pastBalance <= lessonRate || paymentReceived > 0 || !pastInvoiceNumber) {
    return 0;
  }
  
  var dateMatch = pastInvoiceNumber.match(/-(\d{8})$/);
  if (!dateMatch) return 0;
  
  var invoiceDateStr = dateMatch[1];
  var invoiceDate = new Date(
    parseInt(invoiceDateStr.substr(0,4)),
    parseInt(invoiceDateStr.substr(4,2)) - 1,
    parseInt(invoiceDateStr.substr(6,2))
  );
  
  var dueDate = new Date(invoiceDate);
  dueDate.setDate(dueDate.getDate() + 14);
  
  var lateFeeDate = new Date(dueDate);
  lateFeeDate.setDate(lateFeeDate.getDate() + gracePeriod);
  
  var today = new Date();
  if (today > lateFeeDate) {
    return parseFloat(rateMap["Late Fee"]) || 0;
  }
  
  return 0;
}

function copyPreviousColumnToNew(newRow, prevRow, context, currMap, prevMap, mapping) {
  // mapping = { newCol: "New Column Name", prevCol: "Previous Column Name" }
  var newCol = currMap[UtilityScriptLibrary.normalizeHeader(mapping.newCol)];
  var prevCol = prevMap[UtilityScriptLibrary.normalizeHeader(mapping.prevCol)];
  
  if (newCol && prevCol && prevRow.length >= prevCol) {
    newRow[newCol - 1] = prevRow[prevCol - 1];
    return true;
  }
  return false;
}

function getFilterDate(context, previousStartDate) {
  if (previousStartDate) return previousStartDate;
  
  var metadataSheet = context.ss.getSheetByName("Billing Metadata");
  if (metadataSheet) {
    var metadataData = metadataSheet.getDataRange().getValues();
    if (metadataData.length > 1) {
      var lastMetadataRow = metadataData[metadataData.length - 1];
      var retrievedStartDate = lastMetadataRow[2];
      if (retrievedStartDate instanceof Date) {
        return retrievedStartDate;
      }
    }
  }
  return null;
}

function getFormSheet(context) {
  var formResponsesSS = UtilityScriptLibrary.getWorkbook('formResponses');
  var formSheet = formResponsesSS.getSheetByName(context.semesterName);
  
  if (!formSheet) {
    throw new Error("Semester '" + context.semesterName + "' not found in Responses workbook");
  }
  
  return formSheet;
}

function getRateMap(context) {
  // Check if already cached in context
  if (context.rateMap) return context.rateMap;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rateSheet = ss.getSheetByName('Rates');
  var rateData = rateSheet.getDataRange().getValues();
  var rateHeaders = rateData[0];
  var bestColIndex = UtilityScriptLibrary.getMostRecentRateColumn(rateHeaders);
  
  // Use Utility function to build rate map
  context.rateMap = UtilityScriptLibrary.buildRateMapFromSheet(rateSheet, rateHeaders, bestColIndex);
  return context.rateMap;
}

function setDeliveryPreference(newRow, sourceRow, context, sourceIsForm) {
  var h = context.headerMap;
  var deliveryPrefCol = h[UtilityScriptLibrary.normalizeHeader("Delivery Preference")];
  
  if (!deliveryPrefCol) return;
  
  if (sourceIsForm) {
    var get = context.getColIndex;
    var formDeliveryCol = get("Delivery Preference");
    newRow[deliveryPrefCol - 1] = (formDeliveryCol !== -1) ? sourceRow[formDeliveryCol] : "";
  } else {
    var prevH = context.prevHeaderMap;
    var prevDeliveryCol = prevH[UtilityScriptLibrary.normalizeHeader("Delivery Preference")];
    newRow[deliveryPrefCol - 1] = (prevDeliveryCol) ? sourceRow[prevDeliveryCol - 1] : "";
  }
}

function setProgramQuantitiesForCarryover(newRow, context, quantityCols, currencyCols) {
  var rateMap = getRateMap(context);
  var norm = UtilityScriptLibrary.normalizeHeader;
  
  for (var programName in context.programMap) {
    var program = context.programMap[programName];
    var prefix = program.prefix;
    var rateKey = program.rateKey;
    
    var qtyCol = context.headerMap[norm(prefix + " Quantity")];
    var creditCol = context.headerMap[norm(prefix + " Credit")];
    var priceCol = context.headerMap[norm(prefix + " Price")];
    
    if (qtyCol) {
      newRow[qtyCol - 1] = 0;
      quantityCols.push(qtyCol);
    }
    if (creditCol) {
      newRow[creditCol - 1] = 0;
      quantityCols.push(creditCol);
    }
    if (priceCol) {
      newRow[priceCol - 1] = rateMap[rateKey] || 0;
      currencyCols.push(priceCol);
    }
  }
}

function writeAndFormatRows(billingSheet, allStudents) {
  if (allStudents.length === 0) return;
  
  // Write all rows in batch
  var allRows = allStudents.map(function(s) { return s.newRow; });
  billingSheet.getRange(2, 1, allRows.length, allRows[0].length).setValues(allRows);
  
  // Format all rows
  for (var i = 0; i < allStudents.length; i++) {
    formatRow(billingSheet, i + 2, allStudents[i].quantityCols, allStudents[i].currencyCols);
  }
}

function updateCumulativeTracking(billingSheetName) {
  try {
    UtilityScriptLibrary.debugLog("updateCumulativeTracking", "INFO", "Starting cumulative tracking update", 
                  "Billing sheet: " + billingSheetName, "");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var billingSheet = ss.getSheetByName(billingSheetName);
    
    if (!billingSheet) {
      UtilityScriptLibrary.debugLog("updateCumulativeTracking", "WARNING", "Billing sheet not found", 
                    "Sheet name: " + billingSheetName, "");
      return;
    }
    
    var trackingSheet = ss.getSheetByName("Cumulative Tracking");
    if (!trackingSheet) {
      UtilityScriptLibrary.debugLog("updateCumulativeTracking", "ERROR", "Cumulative Tracking sheet not found", "", "");
      return;
    }
    
    var billingData = billingSheet.getDataRange().getValues();
    var billingHeaderMap = UtilityScriptLibrary.getHeaderMap(billingSheet);
    var trackingHeaderMap = UtilityScriptLibrary.getHeaderMap(trackingSheet);
    
    var norm = UtilityScriptLibrary.normalizeHeader;
    var studentIdCol = billingHeaderMap[norm("Student ID")];
    var cumTaughtCol = billingHeaderMap[norm("Current Cumulative Hours Taught")];
    var cumBilledCol = billingHeaderMap[norm("Current Cumulative Hours Billed")];
    
    if (!studentIdCol || !cumTaughtCol || !cumBilledCol) {
      UtilityScriptLibrary.debugLog("updateCumulativeTracking", "ERROR", "Required columns not found in billing sheet", "", "");
      return;
    }
    
    var trackingData = trackingSheet.getDataRange().getValues();
    var trackingIdCol = trackingHeaderMap[norm("Student ID")];
    var trackingCycleCol = trackingHeaderMap[norm("Last Billing Cycle")];
    var trackingTaughtCol = trackingHeaderMap[norm("Cumulative Hours Taught")];
    var trackingBilledCol = trackingHeaderMap[norm("Cumulative Hours Billed")];
    var trackingUpdatedCol = trackingHeaderMap[norm("Last Updated")];
    
    if (!trackingIdCol || !trackingCycleCol || !trackingTaughtCol || !trackingBilledCol || !trackingUpdatedCol) {
      UtilityScriptLibrary.debugLog("updateCumulativeTracking", "ERROR", "Required columns not found in tracking sheet", "", "");
      return;
    }
    
    var now = new Date();
    var updatedCount = 0;
    var addedCount = 0;
    
    // Build tracking lookup map
    var trackingMap = {};
    for (var i = 1; i < trackingData.length; i++) {
      var studentId = trackingData[i][trackingIdCol - 1];
      if (studentId) {
        trackingMap[studentId] = i + 1; // Store actual row number (1-indexed)
      }
    }
    
    // Process each student from billing sheet
    for (var i = 1; i < billingData.length; i++) {
      var row = billingData[i];
      var studentId = row[studentIdCol - 1];
      
      if (!studentId) continue;
      
      var cumTaught = parseFloat(row[cumTaughtCol - 1]) || 0;
      var cumBilled = parseFloat(row[cumBilledCol - 1]) || 0;
      
      if (trackingMap[studentId]) {
        // Update existing row
        var trackingRow = trackingMap[studentId];
        trackingSheet.getRange(trackingRow, trackingCycleCol).setValue(billingSheetName);
        trackingSheet.getRange(trackingRow, trackingTaughtCol).setValue(cumTaught);
        trackingSheet.getRange(trackingRow, trackingBilledCol).setValue(cumBilled);
        trackingSheet.getRange(trackingRow, trackingUpdatedCol).setValue(now);
        updatedCount++;
      } else {
        // Append new row
        trackingSheet.appendRow([studentId, billingSheetName, cumTaught, cumBilled, now]);
        addedCount++;
      }
    }
    
    UtilityScriptLibrary.debugLog("updateCumulativeTracking", "SUCCESS", "Cumulative tracking updated", 
                  "Updated: " + updatedCount + ", Added: " + addedCount, "");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("updateCumulativeTracking", "ERROR", "Failed to update cumulative tracking", 
                  "", error.message);
    throw error;
  }
}

function getCumulativeHistory(studentId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var trackingSheet = ss.getSheetByName("Cumulative Tracking");
    
    if (!trackingSheet) {
      UtilityScriptLibrary.debugLog("getCumulativeHistory", "WARNING", "Cumulative Tracking sheet not found", "", "");
      return null;
    }
    
    var trackingData = trackingSheet.getDataRange().getValues();
    var trackingHeaderMap = UtilityScriptLibrary.getHeaderMap(trackingSheet);
    
    var norm = UtilityScriptLibrary.normalizeHeader;
    var idCol = trackingHeaderMap[norm("Student ID")];
    var cycleCol = trackingHeaderMap[norm("Last Billing Cycle")];
    var taughtCol = trackingHeaderMap[norm("Cumulative Hours Taught")];
    var billedCol = trackingHeaderMap[norm("Cumulative Hours Billed")];
    
    if (!idCol || !cycleCol || !taughtCol || !billedCol) {
      UtilityScriptLibrary.debugLog("getCumulativeHistory", "ERROR", "Required columns not found in tracking sheet", "", "");
      return null;
    }
    
    // Search for student
    for (var i = 1; i < trackingData.length; i++) {
      if (trackingData[i][idCol - 1] === studentId) {
        return {
          lastBillingCycle: trackingData[i][cycleCol - 1],
          cumulativeHoursTaught: parseFloat(trackingData[i][taughtCol - 1]) || 0,
          cumulativeHoursBilled: parseFloat(trackingData[i][billedCol - 1]) || 0
        };
      }
    }
    
    // Student not found
    return null;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("getCumulativeHistory", "ERROR", "Failed to get cumulative history", 
                  "Student ID: " + studentId, error.message);
    return null;
  }
}