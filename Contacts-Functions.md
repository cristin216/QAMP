================================================================================
CONTACTS FUNCTION DIRECTORY
================================================================================
    Total Functions: 16
    Most Recent version: 1

    This directory provides a quick reference for all functions in Contacts script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), ...

================================================================================
ALPHABETICAL INDEX:
================================================================================
    addFutureProspect, appendContactRow, buildInstrumentRow, getContactByKey,
    getExistingFutureKeys, getExistingInstrumentKeys, getExistingKeys,
    getNormalizedHeaders, onOpen, parseInstrumentsWithLevels, processTeacherResponses,
    runLogHeaders, setCheckboxColumns, syncContactsAndInstruments, transformToContact,
    updateOrInsertRow

================================================================================
FUNCTION CATEGORIES:
================================================================================

    UI_MENU (2 functions):
      onOpen, syncContactsAndInstruments

    MAIN (1 function):
      processTeacherResponses

    CONTACT_MANAGEMENT (3 functions):
      addFutureProspect, appendContactRow, transformToContact

    INSTRUMENT_MANAGEMENT (2 functions):
      buildInstrumentRow, updateOrInsertRow

    HELPER_FUNCTIONS (8 functions):
      getContactByKey, getExistingFutureKeys, getExistingInstrumentKeys,
      getExistingKeys, getNormalizedHeaders, parseInstrumentsWithLevels,
      runLogHeaders, setCheckboxColumns

================================================================================
FUNCTION DETAILS (ALPHABETICAL):
================================================================================

    addFutureProspect(entry, futureSheet) -> void
        Writes a prospective teacher to the Future Teacher Contacts sheet.
        Skips duplicates based on name key.
        Category: CONTACT_MANAGEMENT
        Local functions used: getExistingFutureKeys()
        Utility functions used: UtilityScriptLibrary.generateKey(),
                                UtilityScriptLibrary.formatPhoneNumber()

    appendContactRow(contactObj, contactsSheet) -> void
        Appends a new contact row to Teachers and Admin sheet.
        Reads headers dynamically. Skips duplicates based on Key column.
        Category: CONTACT_MANAGEMENT
        Local functions used: getExistingKeys()
        Utility functions used: None

    buildInstrumentRow(instrument, level, firstName, lastName, teacherId,
                       teachAtOP, summer, schoolYear) -> Object
        Builds a row object for the Instrument List sheet.
        Returns object keyed to Instrument List column headers.
        Category: INSTRUMENT_MANAGEMENT
        Local functions used: None
        Utility functions used: None

    getContactByKey(contactsSheet, firstName, lastName) -> Object|null
        Finds a contact in Teachers and Admin by name key.
        Returns contact object keyed by lowercase headers, or null if not found.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.generateKey()

    getExistingFutureKeys(futureSheet) -> Array<String>
        Returns array of name keys from Future Teacher Contacts sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.generateKey()

    getExistingInstrumentKeys(sheet) -> Object
        Returns map of "teacherId_instrument" keys to 1-based row indices
        from Instrument List sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    getExistingKeys(contactsSheet) -> Array<String>
        Returns array of key values from Teachers and Admin sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: getNormalizedHeaders()
        Utility functions used: None

    getNormalizedHeaders(sheet) -> Array<String>
        Returns normalized (lowercase, trimmed) header array for a sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.normalizeHeader()

    onOpen() -> void
        Creates Contacts Tools menu on spreadsheet open.
        Category: UI_MENU
        Local functions used: None
        Utility functions used: None

    parseInstrumentsWithLevels(rawString) -> Array<{instrument, levels}>
        Parses raw instrument/level string from form into structured array.
        Handles parentheses, dash, slash, and space delimiters.
        Applies instrument aliases and brass expansion.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: None

    processTeacherResponses() -> void
        Main processing loop for Teacher Responses sheet.
        Routes each row based on Interest and Teach at OP answers:
          - Interest=No → addFutureProspect()
          - Interest=Yes/Maybe, teachAtOP=No → discard
          - Interest=Yes/Maybe, teachAtOP=Yes → appendContactRow() + instrument update
        Category: MAIN
        Local functions used: getExistingInstrumentKeys(), addFutureProspect(),
                              transformToContact(), appendContactRow(),
                              getContactByKey(), parseInstrumentsWithLevels(),
                              buildInstrumentRow(), updateOrInsertRow(),
                              setCheckboxColumns()
        Utility functions used: UtilityScriptLibrary.getSheet(),
                                UtilityScriptLibrary.getFieldMappingFromSheet(),
                                UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader()

    runLogHeaders() -> void
        Diagnostic function. Logs all sheet headers via Utility.
        Run manually from script editor only.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.logAllSheetHeaders()

    setCheckboxColumns(sheet, headerNames) -> void
        Applies checkbox data validation to specified columns in a sheet.
        Category: HELPER_FUNCTIONS
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getColumnHeaders()

    syncContactsAndInstruments() -> void
        Menu entry point. Calls processTeacherResponses() and logs result.
        Category: UI_MENU
        Local functions used: processTeacherResponses()
        Utility functions used: None

    transformToContact(entry, contactsSheet) -> Object
        Transforms a form entry object into a contact row object
        keyed to Teachers and Admin column headers.
        Generates Teacher ID and Key. Formats phone and address.
        Category: CONTACT_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.generateNextId(),
                                UtilityScriptLibrary.generateKey(),
                                UtilityScriptLibrary.formatPhoneNumber(),
                                UtilityScriptLibrary.formatAddress()

    updateOrInsertRow(sheet, rowObj, existingKeys) -> void
        Updates existing Instrument List row if tracked fields changed,
        or appends new row if not found.
        Tracked fields: Level, Teach at OP, Summer, School Year.
        Category: INSTRUMENT_MANAGEMENT
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getColumnHeaders()

================================================================================
END OF FUNCTION DIRECTORY
================================================================================