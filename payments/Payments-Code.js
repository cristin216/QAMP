/*
================================================================================
PAYMENTS CODE
================================================================================
Version: 3
Total Functions: 3
Documentation: See Payments-Functions.md
================================================================================
*/

function onEditInstallable(e) {
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (row === 1 || e.range.getNumRows() > 1) return;

  var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
  var receiptNeededCol = headerMap['receiptneeded'];
  var receiptUrlCol = headerMap['receipturl'];

  if (!receiptNeededCol || !receiptUrlCol) return;
  if (col !== receiptNeededCol) return;
  if (e.value !== 'TRUE') return;

  var existingUrl = sheet.getRange(row, receiptUrlCol).getValue();
  if (existingUrl && String(existingUrl).trim() !== '') return;

  var env = UtilityScriptLibrary.EnvironmentManager.get();
  var config = UtilityScriptLibrary.getConfig();
  if (!config[env].receiptsFolderId) {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'ERROR', 'receiptsFolderId not found in config', '', '');
    return;
  }

  var receiptsFolder;
  try {
    receiptsFolder = DriveApp.getFolderById(config[env].receiptsFolderId);
  } catch (err) {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'ERROR', 'Cannot access receipts folder', config[env].receiptsFolderId, err.message);
    return;
  }

  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var studentLastNameCol = headerMap['studentlastname'];
  var studentFirstNameCol = headerMap['studentfirstname'];
  var amountPaidCol = headerMap['amountpaid'];
  var dateCol = headerMap['date'];
  var commentsCol = headerMap['comments'];
  var studentIdCol = headerMap['studentid'];
  var invoiceNumberCol = headerMap['invoicenumber'];

  if (!studentLastNameCol || !studentFirstNameCol || !amountPaidCol || !dateCol) {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'ERROR', 'Required columns missing in sheet', sheet.getName(), '');
    return;
  }

  var paymentData = {
    studentLastName: data[studentLastNameCol - 1] || '',
    studentFirstName: data[studentFirstNameCol - 1] || '',
    amountPaid: data[amountPaidCol - 1] || 0,
    date: data[dateCol - 1] || new Date(),
    comments: commentsCol ? data[commentsCol - 1] || '' : '',
    studentId: studentIdCol ? data[studentIdCol - 1] || '' : '',
    invoiceNumber: invoiceNumberCol ? data[invoiceNumberCol - 1] || '' : ''
  };

  if (!paymentData.studentLastName || !paymentData.amountPaid) {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'WARNING', 'Missing student name or amount', 'Row: ' + row, '');
    return;
  }

  var result = createPaymentReceiptDocument(paymentData, receiptsFolder);

  if (result.success) {
    sheet.getRange(row, receiptUrlCol).setValue(result.url);
    UtilityScriptLibrary.debugLog('onEditInstallable', 'SUCCESS', 'Receipt generated', 
                                  'Student: ' + paymentData.studentFirstName + ' ' + paymentData.studentLastName, '');
  } else {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'ERROR', 'Receipt generation failed', 
                                  'Row: ' + row, result.error);
  }
}

function createPaymentReceiptDocument(paymentData, receiptsFolder) {
  try {
    var paymentDate = paymentData.date instanceof Date ?
                      paymentData.date :
                      new Date(paymentData.date);
    var formattedDate = UtilityScriptLibrary.formatDateFlexible(paymentDate, 'MMMM d, yyyy');
    var dateForFilename = UtilityScriptLibrary.formatDateFlexible(paymentDate, 'yyyyMMdd');

    var amount = parseFloat(String(paymentData.amountPaid).replace(/[^0-9.-]/g, ''));
    var formattedAmount = '$' + amount.toFixed(2);
    var paymentMethod = /cash/i.test(paymentData.comments) ? 'Cash' : 'Check';

    var lastName = paymentData.studentLastName || 'Payment';
    var baseFileName = 'Receipt - ' + lastName + ' - ' + dateForFilename;

    // Handle duplicate filenames
    var fileName = baseFileName;
    if (receiptsFolder.getFilesByName(baseFileName + '.pdf').hasNext()) {
      var sequenceNumber = 2;
      while (true) {
        fileName = baseFileName + ' - ' + sequenceNumber;
        if (!receiptsFolder.getFilesByName(fileName + '.pdf').hasNext()) break;
        sequenceNumber++;
      }
    }

    var variables = {
      PaymentDate:   formattedDate,
      AmountPaid:    formattedAmount,
      PaymentMethod: paymentMethod,
      InvoiceNumber: paymentData.invoiceNumber || '',
      StudentId:     paymentData.studentId || '',
      StudentName:   ((paymentData.studentFirstName || '') + ' ' + (paymentData.studentLastName || '')).trim()
    };

    var result = UtilityScriptLibrary.generateDocumentFromTemplate('paymentReceipt', variables, fileName, receiptsFolder);

    if (!result.success) {
      throw new Error(result.error);
    }

    // Convert to PDF
    var pdfBlob = DriveApp.getFileById(result.fileId).getAs('application/pdf');
    pdfBlob.setName(fileName + '.pdf');
    var pdfFile = receiptsFolder.createFile(pdfBlob);
    DriveApp.getFileById(result.fileId).setTrashed(true);

    return {
      success: true,
      fileId: pdfFile.getId(),
      url: pdfFile.getUrl(),
      fileName: fileName + '.pdf'
    };

  } catch (error) {
    UtilityScriptLibrary.debugLog('createPaymentReceiptDocument', 'ERROR', 'Failed to create receipt',
      'Student: ' + (paymentData.studentFirstName || 'Unknown') + ' ' + (paymentData.studentLastName || ''),
      error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

function logSheetHeaders() {
  UtilityScriptLibrary.logAllSheetHeaders();
}