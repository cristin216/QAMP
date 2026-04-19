================================================================================
TEACHER RESPONSES FUNCTION DIRECTORY
================================================================================
    Total Functions: 21
    Most Recent version: 15

    This directory provides a quick reference for all functions in TeacherResponses script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), UtilityScriptLibrary.function2()

  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    addOrUpdateTeacherRosterLookup, createTeacherRosterLookupSheet, extractInstrumentsList,
    extractTeacherFormData, findInstrumentRow, findTeacherInRosterLookup, findTeacherRow,
    generateUniqueDisplayNames, handleTeacherFormSubmit, onOpen, processSingleInstrument,
    processTeacher, processTeacherInstruments, processTeacherManually, setInstrumentAvailability,
    trackFutureProspect, updateDisplayNameInRosterLookup, updateInstrumentAvailability,
    updateLocalTeacherTracking, updateTeacherDisplayNames, updateTeacherFields

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addOrUpdateTeacherRosterLookup(formData, teacherId) -> void
        Adds a new teacher or updates an existing teacher in the Teacher Roster Lookup
        sheet in the Form Responses workbook. Creates the lookup sheet if it doesn't exist.
        Category: ROSTER_MANAGEMENT
        Local functions used: createTeacherRosterLookupSheet(), findTeacherInRosterLookup()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog()

    createTeacherRosterLookupSheet(workbook) -> Sheet
        Creates the Teacher Roster Lookup sheet with headers in the specified workbook.
        Returns the newly created sheet.
        Category: SHEET_OPERATIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    extractInstrumentsList(formData) -> Array
        Extracts and parses instrument data from form submission. Handles multiple formats
        (dash notation, parentheses, plain text). Returns array of instrument objects with
        instrument name, level, and availability flags (teachAtOP, summer, schoolYear).
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getFieldMappingFromSheet(), 
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    extractTeacherFormData(e) -> Object
        Extracts form data from either a form submission event or active sheet.
        Returns an object with form field names as keys and their values.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getColumnHeaders()

    findInstrumentRow(sheet, firstName, lastName, instrument) -> Number
        Finds the row number for a specific teacher/instrument combination in the
        Instrument List sheet. Returns -1 if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder()

    findTeacherInRosterLookup(lookupSheet, teacherName) -> Number
        Finds a teacher in the Teacher Roster Lookup sheet by full name.
        Returns the row number or -1 if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder()

    findTeacherRow(sheet, teacherKey) -> Number
        Finds a teacher in the Teachers and Admin sheet using a generated key.
        Returns the row number or -1 if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder()

    generateUniqueDisplayNames(baseLastName, teachers) -> Object
        Generates unique display names for teachers with the same last name.
        Uses first initial(s) to disambiguate when necessary.
        Returns object mapping firstName to displayName.
        Category: NAME_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    handleTeacherFormSubmit(e) -> void
        Main form submission trigger function. Processes teacher interest form submissions,
        checks teaching interest, and routes to appropriate processing functions.
        Uses script lock to prevent concurrent executions.
        Category: FORM_PROCESSING
        Local functions used: extractTeacherFormData(), trackFutureProspect(), processTeacher(),
                              processTeacherInstruments(), addOrUpdateTeacherRosterLookup(),
                              updateLocalTeacherTracking()
        Utility functions used: UtilityScriptLibrary.debugLog(), UtilityScriptLibrary.getWorkbook()

    onOpen() -> void
        Creates custom menu in the spreadsheet UI when the workbook opens.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    processSingleInstrument(instrumentSheet, instrumentData, firstName, lastName, teacherId, getCol) -> void
        Processes a single instrument entry for a teacher. Updates existing entries or
        creates new ones in the Instrument List sheet.
        Category: INSTRUMENT_PROCESSING
        Local functions used: findInstrumentRow(), updateInstrumentAvailability(), setInstrumentAvailability()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processTeacher(formData, teachersSheet) -> String
        Main teacher processing function. Creates or updates teacher records in the
        Teachers and Admin sheet. Handles duplicate checking, display name generation,
        and field mapping. Returns the teacher ID.
        Category: TEACHER_PROCESSING
        Local functions used: findTeacherRow(), updateTeacherFields(), updateTeacherDisplayNames()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.createColumnFinder(),
                                UtilityScriptLibrary.generateKey(), UtilityScriptLibrary.generateNextId(),
                                UtilityScriptLibrary.formatPhoneNumber(), UtilityScriptLibrary.formatAddress(),
                                UtilityScriptLibrary.debugLog()

    processTeacherInstruments(formData, instrumentSheet, teacherId) -> void
        Processes all instruments for a teacher by extracting instrument list from form
        data and calling processSingleInstrument for each.
        Category: INSTRUMENT_PROCESSING
        Local functions used: extractInstrumentsList(), processSingleInstrument()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.createColumnFinder(),
                                UtilityScriptLibrary.debugLog()

    processTeacherManually() -> void
        Manual batch processing function triggered from UI menu. Processes all untracked
        teachers in the responses sheet and shows a summary dialog.
        Category: BATCH_PROCESSING
        Local functions used: trackFutureProspect(), processTeacher(), processTeacherInstruments(),
                              addOrUpdateTeacherRosterLookup(), updateLocalTeacherTracking()
        Utility functions used: UtilityScriptLibrary.getColumnHeaders(), UtilityScriptLibrary.getWorkbook(),
                                UtilityScriptLibrary.debugLog()

    setInstrumentAvailability(row, instrumentData, getCol) -> void
        Sets availability checkboxes (Teach at OP, Summer, School Year) for a new
        instrument row array before it's appended to the sheet.
        Category: INSTRUMENT_PROCESSING
        Local functions used: None
        Utility functions used: None

    trackFutureProspect(formData) -> void
        Adds teachers who are not currently interested but want to be contacted later
        to the Future Prospects sheet.
        Category: PROSPECT_TRACKING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    updateDisplayNameInRosterLookup(teacherFullName, newDisplayName) -> void
        Updates the display name for a teacher in the Teacher Roster Lookup sheet
        after display names are regenerated due to conflicts.
        Category: ROSTER_MANAGEMENT
        Local functions used: findTeacherInRosterLookup()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.debugLog()

    updateInstrumentAvailability(sheet, row, instrumentData, getCol) -> void
        Updates availability checkboxes for an existing instrument row in the sheet.
        Category: INSTRUMENT_PROCESSING
        Local functions used: None
        Utility functions used: None

    updateLocalTeacherTracking(formData, teacherId) -> void
        Updates the local Teacher Tracking sheet in the form responses workbook to
        record which teachers have been processed. Creates the tracking sheet if needed.
        Category: TRACKING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()

    updateTeacherDisplayNames(newTeacherData) -> String
        Checks for display name conflicts when adding a new teacher. Updates existing
        teachers' display names if needed and returns the display name for the new teacher.
        Category: NAME_MANAGEMENT
        Local functions used: generateUniqueDisplayNames(), updateDisplayNameInRosterLookup()
        Utility functions used: UtilityScriptLibrary.getWorkbook(), UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.debugLog()

    updateTeacherFields(sheet, row, fieldMap, get, getCol) -> void
        Updates teacher contact fields (email, phone, address, salutation) in the
        Teachers and Admin sheet using field mapping.
        Category: TEACHER_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.formatPhoneNumber(), UtilityScriptLibrary.formatAddress()
================================================================================
END OF FUNCTION DIRECTORY
================================================================================