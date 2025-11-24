/**
 * SETUP INSTRUCTIONS:
 * 1. In Google Apps Script Editor, create and run setupScriptProperties() 
 *    with your actual Google IDs (one time only)
 * 2. After running, delete that function
 * 3. This template shows the structure but uses Properties Service for security
 */

// ONE-TIME SETUP FUNCTION - Run in Google Apps Script Editor
function setupScriptProperties() {
  const props = PropertiesService.getScriptProperties();
  
  props.setProperties({
    // Test Environment IDs
    'TEST_PAYMENTS_ID': 'YOUR_TEST_PAYMENTS_ID_HERE',
    'TEST_CONTACTS_ID': 'YOUR_TEST_CONTACTS_ID_HERE',
    'TEST_BILLING_ID': 'YOUR_TEST_BILLING_ID_HERE',
    'TEST_FORM_RESPONSES_ID': 'YOUR_TEST_FORM_RESPONSES_ID_HERE',
    'TEST_ROSTER_FOLDER_ID': 'YOUR_TEST_ROSTER_FOLDER_ID_HERE',
    'TEST_TEMPLATE_FOLDER_ID': 'YOUR_TEST_TEMPLATE_FOLDER_ID_HERE',
    'TEST_METADATA_SHEET_NAME': 'YOUR_TEST_METADATA_SHEET_NAME_HERE',
    'TEST_TEACHER_INVOICES_ID': 'YOUR_TEST_TEACHER_INVOICES_ID_HERE',
    'TEST_TEACHER_INTEREST_ID': 'YOUR_TEST_TEACHER_INTEREST_ID_HERE',
    'TEST_GENERATED_DOCUMENT_FOLDER_ID': 'YOUR_TEST_GENERATED_DOCUMENT_FOLDER_ID_HERE',
    
    // Prod Environment IDs
    'PROD_PAYMENTS_ID': 'YOUR_PROD_PAYMENTS_ID_HERE',
    'PROD_CONTACTS_ID': 'YOUR_PROD_CONTACTS_ID_HERE',
    'PROD_BILLING_ID': 'YOUR_PROD_BILLING_ID_HERE',
    'PROD_FORM_RESPONSES_ID': 'YOUR_PROD_FORM_RESPONSES_ID_HERE',
    'PROD_ROSTER_FOLDER_ID': 'YOUR_PROD_ROSTER_FOLDER_ID_HERE',
    'PROD_TEMPLATE_FOLDER_ID': 'YOUR_PROD_TEMPLATE_FOLDER_ID_HERE',
    'PROD_METADATA_SHEET_NAME': 'YOUR_PROD_METADATA_SHEET_NAME_HERE',
    'PROD_TEACHER_INVOICES_ID': 'YOUR_PROD_TEACHER_INVOICES_ID_HERE',
    'PROD_TEACHER_INTEREST_ID': 'YOUR_PROD_TEACHER_INTEREST_ID_HERE',
    'PROD_GENERATED_DOCUMENT_FOLDER_ID': 'YOUR_PROD_GENERATED_DOCUMENT_FOLDER_ID_HERE'
  });
  
  Logger.log('Properties configured! Delete this function now.');
}

// ===== CONFIGURATION =====
var _executionCache = {};

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  
  return {
    test: {
      paymentsId: props.getProperty('TEST_PAYMENTS_ID'),
      contactsId: props.getProperty('TEST_CONTACTS_ID'),
      billingId: props.getProperty('TEST_BILLING_ID'),
      formResponsesId: props.getProperty('TEST_FORM_RESPONSES_ID'),
      rosterFolderId: props.getProperty('TEST_ROSTER_FOLDER_ID'),
      templateFolderId: props.getProperty('TEST_TEMPLATE_FOLDER_ID'),
      metadataSheetName: props.getProperty('TEST_METADATA_SHEET_NAME'),
      teacherInvoicesId: props.getProperty('TEST_TEACHER_INVOICES_ID'),
      teacherInterestId: props.getProperty('TEST_TEACHER_INTEREST_ID'),
      generatedDocumentFolderId: props.getProperty('TEST_GENERATED_DOCUMENT_FOLDER_ID')
    },
    prod: {
      paymentsId: props.getProperty('PROD_PAYMENTS_ID'),
      contactsId: props.getProperty('PROD_CONTACTS_ID'),
      billingId: props.getProperty('PROD_BILLING_ID'),
      formResponsesId: props.getProperty('PROD_FORM_RESPONSES_ID'),
      rosterFolderId: props.getProperty('PROD_ROSTER_FOLDER_ID'),
      templateFolderId: props.getProperty('PROD_TEMPLATE_FOLDER_ID'),
      metadataSheetName: props.getProperty('PROD_METADATA_SHEET_NAME'),
      teacherInvoicesId: props.getProperty('PROD_TEACHER_INVOICES_ID'),
      teacherInterestId: props.getProperty('PROD_TEACHER_INTEREST_ID'),
      generatedDocumentFolderId: props.getProperty('PROD_GENERATED_DOCUMENT_FOLDER_ID')
    }
  };
}

var CONFIG = getConfig();