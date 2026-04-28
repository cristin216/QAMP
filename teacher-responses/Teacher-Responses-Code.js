
function handleReturningFormSubmit(e) {
  if (e && e.range && e.range.getSheet().getName() !== UtilityScriptLibrary.SHEET_MAP.teacherReturningResponses.name) {
    UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'INFO', 'Wrong sheet — skipping', e.range.getSheet().getName(), '');
    return;
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'ERROR', 'Could not obtain lock', '', error.message);
    return;
  }

  try {
    UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'INFO', 'Starting', '', '');

    var formData = extractReturningFormData(e);
    var fieldMapSheet = UtilityScriptLibrary.getSheet('teacherReturningFieldMap');
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);

    var get = function(internalField) {
      var normalizedFormHeader = Object.keys(fieldMap).find(function(key) {
        return fieldMap[key] === internalField;
      });
      if (!normalizedFormHeader) return '';
      var actualKey = Object.keys(formData).find(function(key) {
        return UtilityScriptLibrary.normalizeHeader(key) === normalizedFormHeader;
      });
      return actualKey ? formData[actualKey] : '';
    };

    var firstName = get('First Name');
    var lastName = get('Last Name');
    var email = get('Email');
    var interest = get('Interest').toString().trim().toLowerCase();
    var teachAtOP = get('Teach at OP').toString().trim().toLowerCase();

    UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'INFO', 'Form data extracted',
      firstName + ' ' + lastName + ' | Interest: ' + interest, '');

    var contactsSheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
    var teacherKey = UtilityScriptLibrary.generateKey(firstName, lastName);
    var contactRow = findTeacherRow(contactsSheet, teacherKey);
    var headerMap = UtilityScriptLibrary.getHeaderMap(contactsSheet);
    var statusCol = headerMap[UtilityScriptLibrary.normalizeHeader('Status')];
    var teacherIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    var emailCol = headerMap[UtilityScriptLibrary.normalizeHeader('Email')];
    var currentStatus = contactRow !== -1
      ? String(contactsSheet.getRange(contactRow, statusCol).getValue() || '').trim()
      : '';

    // --- NO: set Unavailable or Former ---
    if (interest === 'no') {
      if (contactRow === -1) {
        UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'WARNING',
          'No-response with no Contacts match — no action', firstName + ' ' + lastName, '');
        return;
      }
      if (currentStatus === 'Former') {
        UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'INFO',
          'Already Former — no action', firstName + ' ' + lastName, '');
        return;
      }

      var future = get('Future').toString().trim().toLowerCase();
      var newStatus = future.startsWith('y') ? 'Unavailable' : 'Former';
      contactsSheet.getRange(contactRow, statusCol).setValue(newStatus);

      if (newStatus === 'Former') {
        var teacherId = String(contactsSheet.getRange(contactRow, teacherIdCol).getValue() || '').trim();
        UtilityScriptLibrary.cascadeFormerStatus(teacherId);
      }

      UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'SUCCESS',
        'Status set to ' + newStatus, firstName + ' ' + lastName, '');
      return;
    }

    // --- YES / MAYBE ---
    if (!teachAtOP.startsWith('y')) {
      UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'WARNING',
        'Unwilling to teach at OP — no action', firstName + ' ' + lastName, '');
      return;
    }

    // Unmatched: create partial record and exit
    if (contactRow === -1) {
      createPartialReturningRecord(contactsSheet, firstName, lastName, email);
      UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'WARNING',
        'No Contacts match — partial record created', firstName + ' ' + lastName, '');
      return;
    }

    // Former submitted — flag for admin, no status change
    if (currentStatus === 'Former') {
      appendAdminFlag(contactsSheet, contactRow, headerMap,
        'Submitted returning form while status is Former — review required');
      UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'WARNING',
        'Former teacher submitted returning form — flagged', firstName + ' ' + lastName, '');
      return;
    }

    // Update email if changed
    if (emailCol && email) {
      var existingEmail = String(contactsSheet.getRange(contactRow, emailCol).getValue() || '').trim();
      if (existingEmail !== email) {
        contactsSheet.getRange(contactRow, emailCol).setValue(email);
        UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'INFO',
          'Email updated', firstName + ' ' + lastName, '');
      }
    }

    // Unavailable → Returning
    if (currentStatus === 'Unavailable') {
      contactsSheet.getRange(contactRow, statusCol).setValue('Returning');
      UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'INFO',
        'Status updated Unavailable → Returning', firstName + ' ' + lastName, '');
    }

    // Parse and update instruments
    var formHeaders = Object.keys(formData);
    var formValues = formHeaders.map(function(h) { return formData[h]; });
    var instruments = UtilityScriptLibrary.parseGridInstruments(formHeaders, formValues);

    if (instruments.length > 0) {
      var teacherId = String(contactsSheet.getRange(contactRow, teacherIdCol).getValue() || '').trim();
      var instrumentSheet = UtilityScriptLibrary.getSheet('instrumentList');
      var getCol = UtilityScriptLibrary.createColumnFinder(instrumentSheet);
      var summer = get('Summer').toString().toLowerCase().startsWith('y');
      var schoolYear = get('School Year').toString().toLowerCase().startsWith('y');

      for (var i = 0; i < instruments.length; i++) {
        processSingleInstrument(instrumentSheet, {
          instrument: instruments[i].instrument,
          level: instruments[i].levels,
          teachAtOP: true,
          summer: summer,
          schoolYear: schoolYear
        }, firstName, lastName, teacherId, getCol);
      }
    }

    UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'SUCCESS',
      'Completed', firstName + ' ' + lastName, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('handleReturningFormSubmit', 'ERROR',
      'Failed', '', error.message);
  } finally {
    lock.releaseLock();
  }
}

function handleTeacherFormSubmit(e) {
  if (e && e.range && e.range.getSheet().getName() !== UtilityScriptLibrary.SHEET_MAP.teacherResponses.name) {
    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'Wrong sheet — skipping', e.range.getSheet().getName(), '');
    return;
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'ERROR', 'Could not obtain lock', '', error.message);
    return;
  }

  try {
    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'Starting', '', '');

    var formData = extractTeacherFormData(e);
    var fieldMapSheet = UtilityScriptLibrary.getSheet('teacherFieldMap');
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);

    var get = function(internalField) {
      var normalizedFormHeader = Object.keys(fieldMap).find(function(key) {
        return fieldMap[key] === internalField;
      });
      if (!normalizedFormHeader) return '';
      var actualKey = Object.keys(formData).find(function(key) {
        return UtilityScriptLibrary.normalizeHeader(key) === normalizedFormHeader;
      });
      return actualKey ? formData[actualKey] : '';
    };

    var firstName = get('First Name');
    var lastName = get('Last Name');
    var interest = get('Interest').toString().trim().toLowerCase();

    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'Form data extracted',
      firstName + ' ' + lastName + ' | Interest: ' + interest, '');

    // --- NO ---
    if (interest === 'no') {
      var future = get('Future').toString().trim().toLowerCase();
      if (future.startsWith('y')) {
        var futureSheet = UtilityScriptLibrary.getSheet('futureTeachers');
        var futureHeaders = UtilityScriptLibrary.getColumnHeaders(futureSheet);
        var futureRow = futureHeaders.map(function(h) {
          var key = h.toLowerCase().trim();
          if (key === 'timestamp') return get('Timestamp') || new Date();
          if (key === 'last name') return lastName;
          if (key === 'first name') return firstName;
          if (key === 'salutation') return get('Salutation');
          if (key === 'email') return get('Future Email');
          if (key === 'contact again') return 'Yes';
          if (key === 'notes') return '';
          return '';
        });
        futureSheet.appendRow(futureRow);
        UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'Added to future contacts',
          firstName + ' ' + lastName, '');
      } else {
        UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'No interest, no future contact — no action',
          firstName + ' ' + lastName, '');
      }
      return;
    }

    // --- YES / MAYBE ---
    var teachAtOP = get('Teach at OP').toString().trim().toLowerCase();
    if (!teachAtOP.startsWith('y')) {
      UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'WARNING',
        'Unwilling to teach at OP — no action', firstName + ' ' + lastName, '');
      return;
    }

    var contactsSheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
    var teacherId = processTeacher(formData, contactsSheet);
    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'Teacher processed', 'ID: ' + teacherId, '');

    var instrumentSheet = UtilityScriptLibrary.getSheet('instrumentList');
    var getCol = UtilityScriptLibrary.createColumnFinder(instrumentSheet);
    var summer = get('Summer').toString().toLowerCase().startsWith('y');
    var schoolYear = get('School Year').toString().toLowerCase().startsWith('y');

    var formHeaders = Object.keys(formData);
    var formValues = formHeaders.map(function(h) { return formData[h]; });
    var instruments = UtilityScriptLibrary.parseGridInstruments(formHeaders, formValues);

    if (instruments.length > 0) {
      for (var i = 0; i < instruments.length; i++) {
        processSingleInstrument(instrumentSheet, {
          instrument: instruments[i].instrument,
          level: instruments[i].levels,
          teachAtOP: true,
          summer: summer,
          schoolYear: schoolYear
        }, firstName, lastName, teacherId, getCol);
      }
      UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'INFO', 'Instruments processed',
        instruments.length + ' instruments', '');
    } else {
      UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'WARNING',
        'No instruments found', firstName + ' ' + lastName, '');
    }

    addOrUpdateTeacherRosterLookup(formData, teacherId, get);

    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'SUCCESS', 'Completed',
      firstName + ' ' + lastName, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('handleTeacherFormSubmit', 'ERROR', 'Failed', '', error.message);
  } finally {
    lock.releaseLock();
  }
}

// === PROCESSING ===
function extractReturningFormData(e) {
  var formData = {};
  if (e && e.namedValues) {
    for (var key in e.namedValues) {
      formData[key] = e.namedValues[key][0];
    }
  } else {
    var sheet = UtilityScriptLibrary.getSheet('teacherReturningResponses');
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      formData[headers[i]] = values[i];
    }
  }
  return formData;
}

function extractTeacherFormData(e) {
  var formData = {};
  if (e && e.namedValues) {
    for (var key in e.namedValues) {
      formData[key] = e.namedValues[key][0];
    }
  } else {
    var sheet = UtilityScriptLibrary.getSheet('teacherResponses');
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      formData[headers[i]] = values[i];
    }
  }
  return formData;
}

function processSingleInstrument(instrumentSheet, instrumentData, firstName, lastName, teacherId, getCol) {
  var instrument = instrumentData.instrument;
  var level = instrumentData.level || 'all';
  var family = UtilityScriptLibrary.getInstrumentFamily(instrument);

  var existingRow = findInstrumentRow(instrumentSheet, firstName, lastName, instrument);

  if (existingRow !== -1) {
    UtilityScriptLibrary.debugLog('processSingleInstrument', 'INFO', 'Updating existing instrument entry', instrument, '');
    instrumentSheet.getRange(existingRow, getCol('Status')).setValue('Potential');
    instrumentSheet.getRange(existingRow, getCol('Level')).setValue(level);
    if (getCol('Family')) {
      instrumentSheet.getRange(existingRow, getCol('Family')).setValue(family);
    }
    updateInstrumentAvailability(instrumentSheet, existingRow, instrumentData, getCol);
  } else {
    UtilityScriptLibrary.debugLog('processSingleInstrument', 'INFO', 'Creating new instrument entry', instrument, '');
    var newRow = new Array(instrumentSheet.getLastColumn()).fill('');

    newRow[getCol('Instrument') - 1] = instrument;
    newRow[getCol('First Name') - 1] = firstName;
    newRow[getCol('Last Name') - 1] = lastName;
    newRow[getCol('Status') - 1] = 'Potential';
    newRow[getCol('Level') - 1] = level;
    newRow[getCol('Teacher ID') - 1] = teacherId;
    if (getCol('Family')) {
      newRow[getCol('Family') - 1] = family;
    }

    setInstrumentAvailability(newRow, instrumentData, getCol);
    instrumentSheet.appendRow(newRow);
  }
}

function processTeacher(formData, teachersSheet) {
  try {
    UtilityScriptLibrary.debugLog('processTeacher', 'INFO', 'Starting', '', '');

    var fieldMapSheet = UtilityScriptLibrary.getSheet('teacherFieldMap');
    if (!fieldMapSheet) {
      throw new Error('Teacher Field Map sheet not found');
    }
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);

    var get = function(internalField) {
      var formHeader = Object.keys(fieldMap).find(function(key) {
        return fieldMap[key] === internalField;
      });
      if (!formHeader) {
        UtilityScriptLibrary.debugLog('processTeacher', 'WARNING', 'No form header found for internal field', internalField, '');
        return '';
      }
      var actualFormKey = Object.keys(formData).find(function(key) {
        return UtilityScriptLibrary.normalizeHeader(key) === formHeader;
      });
      return actualFormKey ? formData[actualFormKey] : '';
    };

    var getCol = UtilityScriptLibrary.createColumnFinder(teachersSheet);

    var firstName = get('First Name');
    var lastName = get('Last Name');
    var teacherKey = UtilityScriptLibrary.generateKey(firstName, lastName);
    var teacherRow = findTeacherRow(teachersSheet, teacherKey);

    UtilityScriptLibrary.debugLog('processTeacher', 'INFO', 'Duplicate check',
      'Key: ' + teacherKey + ' | Row: ' + teacherRow, '');

    var teacherId = '';

    if (teacherRow !== -1) {
      UtilityScriptLibrary.debugLog('processTeacher', 'INFO', 'Updating existing teacher', firstName + ' ' + lastName, '');
      var rowValues = teachersSheet.getRange(teacherRow, 1, 1, teachersSheet.getLastColumn()).getValues()[0];
      teacherId = String(rowValues[getCol('Teacher ID') - 1] || '');
      updateTeacherFields(teachersSheet, teacherRow, fieldMap, get, getCol);

    } else {
      UtilityScriptLibrary.debugLog('processTeacher', 'INFO', 'Creating new teacher', firstName + ' ' + lastName, '');
      teacherId = UtilityScriptLibrary.generateNextId(teachersSheet, 'Teacher ID', 'T');

      var headers = teachersSheet.getRange(1, 1, 1, teachersSheet.getLastColumn()).getValues()[0];
      var newRow = new Array(headers.length).fill('');

      newRow[getCol('Teacher ID') - 1] = teacherId;
      newRow[getCol('First Name') - 1] = firstName;
      newRow[getCol('Last Name') - 1] = lastName;
      newRow[getCol('Key') - 1] = teacherKey;
      newRow[getCol('Role') - 1] = 'Teacher';

      if (getCol('Salutation')) newRow[getCol('Salutation') - 1] = get('Salutation');
      if (getCol('Email')) newRow[getCol('Email') - 1] = get('Email');
      if (getCol('E-mail 2')) newRow[getCol('E-mail 2') - 1] = get('E-mail 2');
      if (getCol('Phone')) newRow[getCol('Phone') - 1] = UtilityScriptLibrary.formatPhoneNumber(get('Phone'));
      if (getCol('Address')) {
        newRow[getCol('Address') - 1] = UtilityScriptLibrary.formatAddress(get('Street'), get('City'), get('Zip'));
      }

      teachersSheet.appendRow(newRow);
    }

    UtilityScriptLibrary.debugLog('processTeacher', 'SUCCESS', 'Completed', 'ID: ' + teacherId, '');
    return teacherId;

  } catch (error) {
    UtilityScriptLibrary.debugLog('processTeacher', 'ERROR', 'Failed', '', error.message);
    throw error;
  }
}

// === HELPER FUNCTIONS ===

function addOrUpdateTeacherRosterLookup(formData, teacherId, get) {
  try {
    var firstName = get('First Name');
    var lastName = get('Last Name');

    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var existingRow = findTeacherInRosterLookup(lookupSheet, teacherId);

    if (existingRow === -1) {
      var headers = UtilityScriptLibrary.getColumnHeaders(lookupSheet);
      var newRow = headers.map(function(h) {
        var key = UtilityScriptLibrary.normalizeHeader(h);
        if (key === UtilityScriptLibrary.normalizeHeader('First Name')) return firstName;
        if (key === UtilityScriptLibrary.normalizeHeader('Last Name')) return lastName;
        if (key === UtilityScriptLibrary.normalizeHeader('Teacher ID')) return teacherId;
        if (key === UtilityScriptLibrary.normalizeHeader('Status')) return 'Potential';
        if (key === UtilityScriptLibrary.normalizeHeader('Last Updated')) return new Date();
        return '';
      });
      lookupSheet.appendRow(newRow);
      UtilityScriptLibrary.debugLog('addOrUpdateTeacherRosterLookup', 'INFO',
        'Added new teacher', firstName + ' ' + lastName, '');
    } else {
      var lastUpdatedCol = getCol('Last Updated');
      if (lastUpdatedCol) lookupSheet.getRange(existingRow, lastUpdatedCol).setValue(new Date());
      UtilityScriptLibrary.debugLog('addOrUpdateTeacherRosterLookup', 'INFO',
        'Updated existing teacher', firstName + ' ' + lastName, '');
    }

  } catch (error) {
    UtilityScriptLibrary.debugLog('addOrUpdateTeacherRosterLookup', 'ERROR', 'Failed', '', error.message);
  }
}

function appendAdminFlag(contactsSheet, contactRow, headerMap, message) {
  var notesCol = headerMap[UtilityScriptLibrary.normalizeHeader('Notes')];
  if (!notesCol) return;
  var existing = String(contactsSheet.getRange(contactRow, notesCol).getValue() || '').trim();
  var updated = existing ? existing + ' | ' + message : message;
  contactsSheet.getRange(contactRow, notesCol).setValue(updated);
}

function createPartialReturningRecord(contactsSheet, firstName, lastName, email) {
  var teacherId = UtilityScriptLibrary.generateNextId(contactsSheet, 'Teacher ID', 'T');
  var key = UtilityScriptLibrary.generateKey(firstName, lastName);
  var headers = contactsSheet.getRange(1, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
  var getCol = UtilityScriptLibrary.createColumnFinder(contactsSheet);
  var newRow = new Array(headers.length).fill('');

  newRow[getCol('Teacher ID') - 1] = teacherId;
  newRow[getCol('First Name') - 1] = firstName;
  newRow[getCol('Last Name') - 1] = lastName;
  newRow[getCol('Key') - 1] = key;
  newRow[getCol('Role') - 1] = 'Teacher';
  newRow[getCol('Email') - 1] = email;
  newRow[getCol('Status') - 1] = 'Potential';
  newRow[getCol('Notes') - 1] = 'Incomplete record — submitted returning form but not found in Contacts. Requires new teacher form (Form B) before going Active.';

  contactsSheet.appendRow(newRow);
}

function findInstrumentRow(sheet, firstName, lastName, instrument) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var instrumentCol = -1;
  var firstNameCol = -1;
  var lastNameCol = -1;
  
  for (var i = 0; i < headers.length; i++) {
    var normalized = UtilityScriptLibrary.normalizeHeader(headers[i]);
    if (normalized === UtilityScriptLibrary.normalizeHeader("Instrument")) instrumentCol = i;
    if (normalized === UtilityScriptLibrary.normalizeHeader("First Name"))  firstNameCol = i;
    if (normalized === UtilityScriptLibrary.normalizeHeader("Last Name"))   lastNameCol = i;
  }
  
  if (instrumentCol === -1 || firstNameCol === -1 || lastNameCol === -1) {
    UtilityScriptLibrary.debugLog("❌ Required columns not found in instrument sheet");
    return -1;
  }
  
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][instrumentCol]) === instrument &&
        String(data[j][firstNameCol]) === firstName &&
        String(data[j][lastNameCol])  === lastName) {
      return j + 1;
    }
  }
  
  return -1;
}

function findTeacherInRosterLookup(lookupSheet, teacherId) {
  try {
    var getCol = UtilityScriptLibrary.createColumnFinder(lookupSheet);
    var teacherIdCol = getCol('Teacher ID');

    if (!teacherIdCol) {
      UtilityScriptLibrary.debugLog('findTeacherInRosterLookup', 'ERROR',
        'Teacher ID column not found', '', '');
      return -1;
    }

    var data = lookupSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][teacherIdCol - 1]).trim() === String(teacherId).trim()) {
        return i + 1;
      }
    }

    return -1;

  } catch (error) {
    UtilityScriptLibrary.debugLog('findTeacherInRosterLookup', 'ERROR', 'Failed', '', error.message);
    return -1;
  }
}

function findTeacherRow(sheet, teacherKey) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var keyCol = -1;
  
  // Find Key column
  for (var i = 0; i < headers.length; i++) {
    if (UtilityScriptLibrary.normalizeHeader(headers[i]) === "key") {
      keyCol = i;
      break;
    }
  }
  
  if (keyCol === -1) return -1;
  
  // Search for matching key
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][keyCol]).toLowerCase() === teacherKey.toLowerCase()) {
      return i + 1; // Return 1-based row number
    }
  }
  
  return -1;
}

function setInstrumentAvailability(row, instrumentData, getCol) {
  if (getCol("Teach at OP")) row[getCol("Teach at OP") - 1] = instrumentData.teachAtOP || false;
  if (getCol("Summer")) row[getCol("Summer") - 1] = instrumentData.summer || false;
  if (getCol("School Year")) row[getCol("School Year") - 1] = instrumentData.schoolYear || false;
}

function updateInstrumentAvailability(sheet, row, instrumentData, getCol) {
  if (getCol("Teach at OP")) sheet.getRange(row, getCol("Teach at OP")).setValue(instrumentData.teachAtOP || false);
  if (getCol("Summer")) sheet.getRange(row, getCol("Summer")).setValue(instrumentData.summer || false);
  if (getCol("School Year")) sheet.getRange(row, getCol("School Year")).setValue(instrumentData.schoolYear || false);
}

function updateTeacherFields(sheet, row, fieldMap, get, getCol) {
  // Update fields that might have changed using field mapping
  var fieldsToUpdate = [
    {internal: "Email", sheet: "Email"},
    {internal: "E-mail 2", sheet: "E-mail 2"},
    {internal: "Phone", sheet: "Phone"},
    {internal: "Salutation", sheet: "Salutation"}
  ];
  
  for (var i = 0; i < fieldsToUpdate.length; i++) {
    var field = fieldsToUpdate[i];
    var col = getCol(field.sheet);
    var value = get(field.internal);
    
    if (col && value) {
      if (field.sheet === "Phone") {
        value = UtilityScriptLibrary.formatPhoneNumber(value);
      }
      sheet.getRange(row, col).setValue(value);
      UtilityScriptLibrary.debugLog("Updated " + field.sheet + " with: " + value);
    }
  }
  
  // Update address
  var addressCol = getCol("Address");
  if (addressCol) {
    var street = get("Street");
    var city = get("City");
    var zip = get("Zip");
    var address = UtilityScriptLibrary.formatAddress(street, city, zip);
    sheet.getRange(row, addressCol).setValue(address);
    UtilityScriptLibrary.debugLog("Updated Address with: " + address);
  }
}
