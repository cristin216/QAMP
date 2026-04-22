/**
 * SETUP INSTRUCTIONS:
 * 1. Fill in your real IDs in setupScriptProperties() below
 * 2. Run it once in the Apps Script editor
 * 3. Delete the contents of setupScriptProperties() but leave the function shell
 * 4. Never run it again unless doing a full reset
 * 
 * To add/change a single property, use the Apps Script UI:
 * Project Settings → Script Properties → Edit
 * Or use: PropertiesService.getScriptProperties().setProperty('KEY', 'value')
 */

function setupScriptProperties() {
  PropertiesService.getScriptProperties().setProperties({
    // Test Environment
    'TEST_RECEIPTS_FOLDER_ID':              'your-id-here',
    'TEST_CONTACTS_ID':                     'your-id-here',
    'TEST_BILLING_ID':                      'your-id-here',
    'TEST_FORM_RESPONSES_ID':               'your-id-here',
    'TEST_TEACHER_INVOICES_ID':             'your-id-here',
    'TEST_TEACHER_INTEREST_ID':             'your-id-here',
    'TEST_PAYMENTS_ID':                     'your-id-here',
    'TEST_ROSTER_FOLDER_ID':                'your-id-here',
    'TEST_TEMPLATE_FOLDER_ID':              'your-id-here',
    'TEST_GENERATED_DOCUMENTS_FOLDER_ID':   'your-id-here',

    // Prod Environment
    'PROD_RECEIPTS_FOLDER_ID':              'your-id-here',
    'PROD_CONTACTS_ID':                     'your-id-here',
    'PROD_BILLING_ID':                      'your-id-here',
    'PROD_FORM_RESPONSES_ID':               'your-id-here',
    'PROD_TEACHER_INVOICES_ID':             'your-id-here',
    'PROD_TEACHER_INTEREST_ID':             'your-id-here',
    'PROD_PAYMENTS_ID':                     'your-id-here',
    'PROD_ROSTER_FOLDER_ID':                'your-id-here',
    'PROD_TEMPLATE_FOLDER_ID':              'your-id-here',
    'PROD_GENERATED_DOCUMENTS_FOLDER_ID':   'your-id-here'
  });

  Logger.log('Properties set. Delete the contents of this function now.');
}

function getConfig() {
  var props = PropertiesService.getScriptProperties();

  return {
    test: {
      receiptsFolderId:           props.getProperty('TEST_RECEIPTS_FOLDER_ID'),
      contactsId:                 props.getProperty('TEST_CONTACTS_ID'),
      billingId:                  props.getProperty('TEST_BILLING_ID'),
      formResponsesId:            props.getProperty('TEST_FORM_RESPONSES_ID'),
      teacherInvoicesId:          props.getProperty('TEST_TEACHER_INVOICES_ID'),
      teacherInterestId:          props.getProperty('TEST_TEACHER_INTEREST_ID'),
      paymentsId:                 props.getProperty('TEST_PAYMENTS_ID'),
      rosterFolderId:             props.getProperty('TEST_ROSTER_FOLDER_ID'),
      templateFolderId:           props.getProperty('TEST_TEMPLATE_FOLDER_ID'),
      generatedDocumentsFolderId: props.getProperty('TEST_GENERATED_DOCUMENTS_FOLDER_ID')
    },
    prod: {
      receiptsFolderId:           props.getProperty('PROD_RECEIPTS_FOLDER_ID'),
      contactsId:                 props.getProperty('PROD_CONTACTS_ID'),
      billingId:                  props.getProperty('PROD_BILLING_ID'),
      formResponsesId:            props.getProperty('PROD_FORM_RESPONSES_ID'),
      teacherInvoicesId:          props.getProperty('PROD_TEACHER_INVOICES_ID'),
      teacherInterestId:          props.getProperty('PROD_TEACHER_INTEREST_ID'),
      paymentsId:                 props.getProperty('PROD_PAYMENTS_ID'),
      rosterFolderId:             props.getProperty('PROD_ROSTER_FOLDER_ID'),
      templateFolderId:           props.getProperty('PROD_TEMPLATE_FOLDER_ID'),
      generatedDocumentsFolderId: props.getProperty('PROD_GENERATED_DOCUMENTS_FOLDER_ID')
    }
  };
}

var CONFIG = getConfig();