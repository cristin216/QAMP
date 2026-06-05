================================================================================
CONTACTS FUNCTION DIRECTORY
================================================================================
    Total Functions: 2
    Most Recent version: 5

    This directory provides a quick reference for all functions in Contacts script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), ...

    NOTE: Contact processing, instrument management, and form-response handling
    were moved to the Teacher-Responses script. This script handles only
    spreadsheet-level event triggers for the Contacts workbook.

================================================================================
ALPHABETICAL INDEX:
================================================================================
    logSheetHeaders
    onEditContacts

================================================================================
FUNCTION CATEGORIES:
================================================================================

    EVENT_TRIGGER (1 function):
      onEditContacts

    TESTING (1 function):
      logSheetHeaders

================================================================================
FUNCTION DETAILS (ALPHABETICAL):
================================================================================

    logSheetHeaders() -> void
        Logs all sheet names and their header row values to the console for debugging.
        Skips empty sheets.
        Category: TESTING
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.logAllSheetHeaders()

    onEditContacts(e) -> void
        Handles edit events on the Teachers and Admin sheet.
        Watches for Status column changes to 'former' on teacher rows (Teacher ID
        starting with 'T'). Skips multi-cell edits, non-teacher rows, and edits
        outside the Status column. Calls cascadeFormerStatus() to propagate the
        status change downstream.
        Category: EVENT_TRIGGER
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.getHeaderMap(),
                                UtilityScriptLibrary.normalizeHeader(),
                                UtilityScriptLibrary.cascadeFormerStatus(),
                                UtilityScriptLibrary.debugLog()

================================================================================
END OF FUNCTION DIRECTORY
================================================================================
