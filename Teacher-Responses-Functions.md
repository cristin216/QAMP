================================================================================
TEACHER RESPONSES FUNCTION DIRECTORY
================================================================================
    Total Functions: 16
    Most Recent version: 17

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
    addOrUpdateTeacherRosterLookup
    appendAdminFlag
    buildTeacherDisplayName
    createPartialReturningRecord
    extractFormData
    findInstrumentRow
    findTeacherRow
    handleReturningFormSubmit
    handleTeacherFormSubmit
    logSheetHeaders
    migrateTeacherDisplayNames
    processSingleInstrument
    processTeacher
    setInstrumentAvailability
    updateInstrumentAvailability
    updateTeacherFields

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    addOrUpdateTeacherRosterLookup(formData, teacherId, get) -> void
        Adds a new teacher or updates an existing teacher in the Teacher Roster Lookup sheet.
        Builds display name via buildTeacherDisplayName() and uses the get() accessor for
        form field lookup. Uses UtilityScriptLibrary.findTeacherInRosterLookup() for row lookup.
        Category: ROSTER_MANAGEMENT
        Local functions used: buildTeacherDisplayName()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.createColumnFinder(),
                                UtilityScriptLibrary.findTeacherInRosterLookup(),
                                UtilityScriptLibrary.getColumnHeaders(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    appendAdminFlag(contactsSheet, contactRow, headerMap, message) -> void
        Appends a flag message to the Notes column of a contact row, pipe-separating from
        any existing note text.
        Category: DATA_UPDATE
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()
        Called by: handleReturningFormSubmit()

    buildTeacherDisplayName(firstName, lastName, teacherId) -> String
        Builds a unique display name in the format "LastName FirstInitial-IDnum"
        (e.g. "Smith J-2"). Strips leading zeros from the numeric portion of the teacher ID.
        Category: NAME_MANAGEMENT
        Local functions used: None
        Utility functions used: None
        Called by: addOrUpdateTeacherRosterLookup(), migrateTeacherDisplayNames()

    createPartialReturningRecord(contactsSheet, firstName, lastName, email) -> void
        Creates a minimal Potential teacher record in the Teachers and Admin sheet for a
        returning teacher who submitted the returning form but was not found in Contacts.
        Sets status to 'Potential' and adds a note that a new teacher form is required.
        Category: TEACHER_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.generateNextId(), UtilityScriptLibrary.generateKey(),
                                UtilityScriptLibrary.createColumnFinder(), UtilityScriptLibrary.debugLog()
        Called by: handleReturningFormSubmit()

    extractFormData(e?, sheetKey) -> Object
        Extracts form data from either a form submission event or, if no event is provided,
        reads the last row of the sheet identified by sheetKey. Returns an object with
        form field names as keys and their values.
        Category: DATA_EXTRACTION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getSheet()
        Called by: handleTeacherFormSubmit(), handleReturningFormSubmit()

    findInstrumentRow(sheet, firstName, lastName, instrument) -> Number
        Finds the row number for a specific teacher/instrument combination in the
        Instrument List sheet. Returns -1 if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder()

    findTeacherRow(sheet, teacherKey) -> Number
        Finds a teacher in the Teachers and Admin sheet using a generated key.
        Returns the row number or -1 if not found.
        Category: DATA_LOOKUP
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.createColumnFinder()

    handleReturningFormSubmit(e) -> void
        Event handler for returning teacher form submissions. Acquires a script lock,
        extracts form data, looks up the teacher in Contacts, and either updates their
        record or creates a partial record via createPartialReturningRecord(). Processes
        instruments via processSingleInstrument() and updates the roster lookup.
        Appends admin flags for unresolved or problematic records.
        Category: FORM_PROCESSING
        Local functions used: extractReturningFormData(), processTeacher(),
                              addOrUpdateTeacherRosterLookup(), createPartialReturningRecord(),
                              appendAdminFlag(), processSingleInstrument()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.debugLog()

    handleTeacherFormSubmit(e) -> void
        Main event handler for new teacher interest form submissions. Acquires a script lock,
        extracts form data, checks teaching interest, and routes to processTeacher(),
        processSingleInstrument(), and addOrUpdateTeacherRosterLookup() as appropriate.
        Teachers not interested in teaching are tracked in the Future Prospects sheet inline.
        Category: FORM_PROCESSING
        Local functions used: extractTeacherFormData(), processTeacher(),
                              processSingleInstrument(), addOrUpdateTeacherRosterLookup()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.debugLog()

    logSheetHeaders() -> void
        Logs all sheet names and their header row values to the console for debugging.
        Skips empty sheets.
        Category: TESTING
        Local functions used: None
        Utility functions used: None

    migrateTeacherDisplayNames() -> void
        One-time migration utility. Iterates the Teacher Roster Lookup sheet and rewrites
        each teacher's display name to the new format via buildTeacherDisplayName().
        Logs count of updated rows on completion.
        Category: HELPER_FUNCTIONS
        Local functions used: buildTeacherDisplayName()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.createColumnFinder(),
                                UtilityScriptLibrary.debugLog()

    processSingleInstrument(instrumentSheet, instrumentData, firstName, lastName, teacherId, getCol) -> void
        Processes a single instrument entry for a teacher. Updates existing entries or
        creates new ones in the Instrument List sheet.
        Category: INSTRUMENT_PROCESSING
        Local functions used: findInstrumentRow(), updateInstrumentAvailability(), setInstrumentAvailability()
        Utility functions used: UtilityScriptLibrary.debugLog()

    processTeacher(formData, teachersSheet) -> String
        Main teacher processing function. Creates or updates teacher records in the
        Teachers and Admin sheet. Handles duplicate checking, field mapping, and display
        name generation. Returns the teacher ID.
        Category: TEACHER_PROCESSING
        Local functions used: findTeacherRow(), updateTeacherFields()
        Utility functions used: UtilityScriptLibrary.getSheet(), UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.normalizeHeader(), UtilityScriptLibrary.createColumnFinder(),
                                UtilityScriptLibrary.generateKey(), UtilityScriptLibrary.generateNextId(),
                                UtilityScriptLibrary.formatPhoneNumber(), UtilityScriptLibrary.formatAddress(),
                                UtilityScriptLibrary.debugLog()

    setInstrumentAvailability(row, instrumentData, getCol) -> void
        Sets availability checkboxes (Teach at OP, Summer, School Year) for a new
        instrument row array before it's appended to the sheet.
        Category: INSTRUMENT_PROCESSING
        Local functions used: None
        Utility functions used: None

    updateInstrumentAvailability(sheet, row, instrumentData, getCol) -> void
        Updates availability checkboxes for an existing instrument row in the sheet.
        Category: INSTRUMENT_PROCESSING
        Local functions used: None
        Utility functions used: None

    updateTeacherFields(sheet, row, fieldMap, get, getCol) -> void
        Updates teacher contact fields (email, phone, address, salutation) in the
        Teachers and Admin sheet using field mapping.
        Category: TEACHER_PROCESSING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.formatPhoneNumber(), UtilityScriptLibrary.formatAddress()
================================================================================
END OF FUNCTION DIRECTORY
================================================================================