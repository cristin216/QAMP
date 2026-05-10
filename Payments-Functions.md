================================================================================
PAYMENTS FUNCTION DIRECTORY
================================================================================
    Total Functions: 2
    Most Recent version: 2

    This directory provides a quick reference for all functions in Payments script.
    Parameters marked with ? are optional.

    Format: functionName(param1, param2, ...) -> ReturnType
            Brief description of what the function does
            Category: CATEGORY_NAME
            Local functions used: function1(), function2()
            Utility functions used: UtilityScriptLibrary.function1(), UtilityScriptLibrary.function2()

  ================================================================================
  ALPHABETICAL INDEX:
  ================================================================================
    createPaymentReceiptDocument
    onEditInstallable

  ================================================================================
  FUNCTION CATEGORIES:
  ================================================================================

    DOCUMENT_GENERATION (1 function):
      createPaymentReceiptDocument

    EVENT_TRIGGER (1 function):
      onEditInstallable

  ================================================================================
  FUNCTION REFERENCE (Alphabetical)
  ================================================================================

    createPaymentReceiptDocument(paymentData, receiptsFolder) -> Object
        Creates a formatted Google Doc payment receipt, converts it to PDF, saves it
        to the receiptsFolder, and trashes the source Doc. Handles duplicate filenames
        by appending a sequence number. Returns object with success, fileId, url, and
        fileName on success, or success: false and error message on failure.
        Category: DOCUMENT_GENERATION
        Local functions used: None
        Utility functions used: UtilityScriptLibrary.debugLog()
        Called by: onEditInstallable()

    onEditInstallable(e) -> void
        Installable edit trigger. Fires when the Receipt Needed column is set to TRUE
        on any row in the Payments sheet. Skips rows that already have a receipt URL.
        Reads payment data from the row, calls createPaymentReceiptDocument(), and
        writes the resulting PDF URL back to the Receipt URL column.
        Category: EVENT_TRIGGER
        Local functions used: createPaymentReceiptDocument()
        Utility functions used: UtilityScriptLibrary.getHeaderMap(), UtilityScriptLibrary.EnvironmentManager.get(),
                                UtilityScriptLibrary.getConfig(), UtilityScriptLibrary.debugLog()

================================================================================
END OF FUNCTION DIRECTORY
================================================================================