/*
================================================================================
PAYMENTS CODE
================================================================================
Version: 1
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

  var config = UtilityScriptLibrary.getConfig();
  if (!config.receiptsFolderId) {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'ERROR', 'receiptsFolderId not found in config', '', '');
    return;
  }

  var receiptsFolder;
  try {
    receiptsFolder = DriveApp.getFolderById(config.receiptsFolderId);
  } catch (err) {
    UtilityScriptLibrary.debugLog('onEditInstallable', 'ERROR', 'Cannot access receipts folder', config.receiptsFolderId, err.message);
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
    var formattedDate = Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'MMMM d, yyyy');
    var dateForFilename = Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), 'yyyyMMdd');

    var amount = parseFloat(String(paymentData.amountPaid).replace(/[^0-9.-]/g, ''));
    var formattedAmount = '$' + amount.toFixed(2);

    var lastName = paymentData.studentLastName || 'Payment';
    var baseFileName = 'Receipt - ' + lastName + ' - ' + dateForFilename;

    // Handle duplicate filenames
    var existingFiles = receiptsFolder.getFilesByName(baseFileName + '.pdf');
    var sequenceNumber = 1;
    var fileName = baseFileName;

    if (existingFiles.hasNext()) {
      while (true) {
        sequenceNumber++;
        fileName = baseFileName + ' - ' + sequenceNumber;
        var testFiles = receiptsFolder.getFilesByName(fileName + '.pdf');
        if (!testFiles.hasNext()) break;
      }
    }

    // Create Google Doc
    var doc = DocumentApp.create(fileName);
    var body = doc.getBody();
    body.clear();

    body.appendParagraph('').setSpacingAfter(10);

    var orgName = body.appendParagraph('QUAKER ARTS MUSIC PROGRAM');
    orgName.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    orgName.setFontSize(16);
    orgName.setBold(true);
    orgName.setForegroundColor('#4a7c59');

    body.appendParagraph('').setSpacingAfter(5);

    var title = body.appendParagraph('PAYMENT RECEIPT');
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.setFontSize(18);
    title.setBold(true);

    body.appendParagraph('').setSpacingAfter(20);
    body.appendHorizontalRule();
    body.appendParagraph('').setSpacingAfter(15);

    var detailsTable = [];
    detailsTable.push(['Payment Date:', formattedDate]);
    detailsTable.push(['Amount Paid:', formattedAmount]);
    detailsTable.push(['Payment Method:', paymentData.paymentMethod || 'Check']);
    if (paymentData.invoiceNumber) detailsTable.push(['Invoice Number:', paymentData.invoiceNumber]);
    if (paymentData.studentId) detailsTable.push(['Student ID:', paymentData.studentId]);

    var studentName = (paymentData.studentFirstName || '') + ' ' + (paymentData.studentLastName || '');
    if (studentName.trim()) detailsTable.push(['Student:', studentName.trim()]);

    var table = body.appendTable(detailsTable);
    table.setBorderWidth(0);

    for (var i = 0; i < table.getNumRows(); i++) {
      var tableRow = table.getRow(i);
      tableRow.getCell(0).setWidth(150).setPaddingBottom(5).setPaddingTop(5)
              .editAsText().setFontSize(11).setBold(true);
      tableRow.getCell(1).setPaddingBottom(5).setPaddingTop(5)
              .editAsText().setFontSize(11);
    }

    body.appendParagraph('').setSpacingAfter(20);
    body.appendHorizontalRule();
    body.appendParagraph('').setSpacingAfter(15);

    var thankYou = body.appendParagraph('Thank you for your payment.');
    thankYou.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    thankYou.setFontSize(11);
    thankYou.setItalic(true);

    body.appendParagraph('').setSpacingAfter(30);
    body.appendHorizontalRule();
    body.appendParagraph('').setSpacingAfter(10);

    var footer1 = body.appendParagraph('Quaker Arts Music Program');
    footer1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer1.setFontSize(10);

    var footer2 = body.appendParagraph('PO Box 372, Orchard Park, NY 14127');
    footer2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer2.setFontSize(9);

    var footer3 = body.appendParagraph('info@quakermusic.org');
    footer3.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    footer3.setFontSize(9);

    doc.saveAndClose();

    var docId = doc.getId();
    var pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
    pdfBlob.setName(fileName + '.pdf');

    var pdfFile = receiptsFolder.createFile(pdfBlob);
    DriveApp.getFileById(docId).setTrashed(true);

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