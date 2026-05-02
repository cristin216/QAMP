================================================================================
CONTACTS FUNCTION DIRECTORY
================================================================================
    Total Functions: 1
    Most Recent version: 3

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
    onEditContacts

================================================================================
FUNCTION CATEGORIES:
================================================================================

    EVENT_TRIGGER (1 function):
      onEditContacts

================================================================================
FUNCTION DETAILS (ALPHABETICAL):
================================================================================

    onEditContacts(e) -> void
        Handles edit events on the Teachers and Admin sheet.
        Watches for Status column changes to 'Former' on teacher rows (Teacher ID
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