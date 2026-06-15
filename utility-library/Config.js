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
    'TEST_WEBAPP_URL':                      'your-url-here',

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
    'PROD_GENERATED_DOCUMENTS_FOLDER_ID':   'your-id-here',
    'PROD_WEBAPP_URL':                      'your-url-here',

    // Scheduling (environment-independent)
    'SCHEDULING_ID':                        'your-id-here',
    'SCHEDULING_WEBAPP_URL':                'your-url-here',
    'ROOM_1_CALENDAR_ID':                   'your-id-here',
    'ROOM_2_CALENDAR_ID':                   'your-id-here',
    'ROOM_3_CALENDAR_ID':                   'your-id-here'
  });

  Logger.log('Properties set. Delete the contents of this function now.');
}

function getConfig() {
  var props = PropertiesService.getScriptProperties();

  // Scheduling values are environment-independent
  var scheduling = {
    schedulingId:       props.getProperty('SCHEDULING_ID'),
    schedulingWebAppUrl: props.getProperty('SCHEDULING_WEBAPP_URL'),
    room1CalendarId:    props.getProperty('ROOM_1_CALENDAR_ID'),
    room2CalendarId:    props.getProperty('ROOM_2_CALENDAR_ID'),
    room3CalendarId:    props.getProperty('ROOM_3_CALENDAR_ID')
  };

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
      generatedDocumentsFolderId: props.getProperty('TEST_GENERATED_DOCUMENTS_FOLDER_ID'),
      webAppUrl:                  props.getProperty('TEST_WEBAPP_URL'),
      schedulingId:               scheduling.schedulingId,
      schedulingWebAppUrl:        scheduling.schedulingWebAppUrl,
      room1CalendarId:            scheduling.room1CalendarId,
      room2CalendarId:            scheduling.room2CalendarId,
      room3CalendarId:            scheduling.room3CalendarId
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
      generatedDocumentsFolderId: props.getProperty('PROD_GENERATED_DOCUMENTS_FOLDER_ID'),
      webAppUrl:                  props.getProperty('PROD_WEBAPP_URL'),
      schedulingId:               scheduling.schedulingId,
      schedulingWebAppUrl:        scheduling.schedulingWebAppUrl,
      room1CalendarId:            scheduling.room1CalendarId,
      room2CalendarId:            scheduling.room2CalendarId,
      room3CalendarId:            scheduling.room3CalendarId
    }
  };
}

var CONFIG = getConfig();