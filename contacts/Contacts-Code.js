// === MAIN FUNCTIONS ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Contacts Tools')
    .addItem('Sync Contacts & Instruments', 'syncContactsAndInstruments')
    .addToUi();
}

function syncContactsAndInstruments() {
  try {
    processTeacherResponses();
    Logger.log("✅ Sync completed successfully.");
  } catch (e) {
    Logger.log("❌ Error during sync: " + e.message);
  }
}

function addFutureProspect(entry, futureSheet) {
  const existingKeys = getExistingFutureKeys(futureSheet);
  const key = UtilityScriptLibrary.generateKey(entry.FirstName, entry.LastName);

  if (existingKeys.includes(key)) {
    Logger.log(`Skipped duplicate future prospect: ${key}`);
    return;
  }

  const headers = futureSheet.getRange(1, 1, 1, futureSheet.getLastColumn()).getValues()[0];
  const rowObj = {
    'Timestamp': entry.Timestamp || '',
    'Last Name': entry.LastName || '',
    'First Name': entry.FirstName || '',
    'Salutation': entry.Salutation || '',
    'Email': entry.Email || '',
    'Phone': UtilityScriptLibrary.formatPhoneNumber(entry.Phone || ''),
    'Notes': ''
  };

  const row = headers.map(h => rowObj[h] || '');
  futureSheet.appendRow(row);
}

// === HELPER FUNCTIONS ===

function appendContactRow(contactObj, contactsSheet) {
  const headers = contactsSheet.getRange(1, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
  const existingKeys = getExistingKeys(contactsSheet);

  if (existingKeys.includes(contactObj['Key'])) {
    Logger.log(`Skipped duplicate contact with key: ${contactObj['Key']}`);
    return;
  }

  const row = headers.map(header => contactObj[header] || '');
  contactsSheet.appendRow(row);
}

function buildInstrumentRow({
  instrument,
  level,
  firstName,
  lastName,
  teacherId,
  teachAtOP,
  summer,
  schoolYear
  }) {
  return {
    Instrument: instrument,
    Level: level,
    'First Name': firstName,
    'Last Name': lastName,
    'Teacher ID': teacherId,
    Status: 'potential',
    'Teach at OP': !!teachAtOP,
    Summer: !!summer,
    'School Year': !!schoolYear,
    Notes: '',
    Comments: ''
  };
}

function getContactByKey(contactsSheet, firstName, lastName) {
  const key = UtilityScriptLibrary.generateKey(firstName, lastName);
  const data = contactsSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const keyIndex = headers.indexOf('key');

  if (keyIndex === -1) throw new Error("Missing 'Key' column in Contacts sheet");

  for (let i = 1; i < data.length; i++) {
    if ((data[i][keyIndex] || '').toString().toLowerCase().trim() === key) {
      const row = data[i];
      const contact = {};
      headers.forEach((h, j) => contact[h] = row[j]);
      return contact;
    }
  }

  return null;
}

function getExistingFutureKeys(futureSheet) {
  const data = futureSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const firstIndex = headers.indexOf('first name');
  const lastIndex = headers.indexOf('last name');
  if (firstIndex === -1 || lastIndex === -1) return [];

  return data.slice(1).map(row => {
    return UtilityScriptLibrary.generateKey(
      (row[firstIndex] || '').toString(),
      (row[lastIndex] || '').toString()
    );
  }).filter(k => k);
}

function getExistingInstrumentKeys(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const idIndex = headers.indexOf('teacher id');
  const instrumentIndex = headers.indexOf('instrument');

  const keyMap = {};
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][idIndex] || '').toString().trim();
    const instrument = (data[i][instrumentIndex] || '').toString().trim();
    if (id && instrument) {
      keyMap[`${id}_${instrument}`.toLowerCase()] = i + 1; // 1-based row index
    }
  }

  return keyMap;
}

function getExistingKeys(contactsSheet) {
  const keyColumnIndex = getNormalizedHeaders(contactsSheet).indexOf('key');
  if (keyColumnIndex === -1) return [];

  const lastRow = contactsSheet.getLastRow();
  if (lastRow <= 1) return [];  // No data rows

  const keysRange = contactsSheet.getRange(2, keyColumnIndex + 1, lastRow - 1);
  const keys = keysRange.getValues().flat().map(key => key.toString().trim());
  return keys;
}

function getNormalizedHeaders(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.map(header => UtilityScriptLibrary.normalizeHeader(header));
}

function parseInstrumentsWithLevels(rawString) {
  if (!rawString || typeof rawString !== 'string') return [];

  const entries = rawString
    .split(/,|\band\b|\n/i)
    .map(e => e.trim())
    .filter(e => e.length > 0);

  const instrumentAliases = {
    "string bass": "Bass",
    "double bass": "Bass",
    "brass/beg": "Brass",
    "all brass": "Brass",
    "all percussion": "Percussion"
  };

  const brassExpansion = {
    "brass": ["Trumpet", "French Horn", "Trombone", "Euphonium", "Tuba"],
    "high brass": ["Trumpet", "French Horn"],
    "low brass": ["Trombone", "Euphonium", "Tuba"]
  };

  const results = [];

  entries.forEach(entry => {
    let instrument = '';
    let levels = '';

    const parenMatch = entry.match(/^(.+?)\s*\((.+?)\)$/);
    const dashMatch = entry.match(/^(.+?)\s*-\s*(.+)$/);
    const slashMatch = entry.match(/^(.+?)\/(beg|int|adv|all)(?:\/(beg|int|adv|all))?$/i);

    if (parenMatch) {
      instrument = parenMatch[1].trim();
      levels = parenMatch[2].trim().toLowerCase();
    } else if (dashMatch) {
      instrument = dashMatch[1].trim();
      levels = dashMatch[2].trim().toLowerCase();
    } else if (slashMatch) {
      instrument = slashMatch[1].trim();
      levels = [slashMatch[2], slashMatch[3]].filter(Boolean).join('/');
    } else {
      const parts = entry.split(/\s+/);
      if (parts.length > 1 && ['beg', 'int', 'adv', 'all'].includes(parts.at(-1).toLowerCase())) {
        levels = parts.pop().toLowerCase();
        instrument = parts.join(' ');
      } else {
        instrument = entry.trim();
        levels = 'all';
      }
    }

    const aliasKey = instrument.toLowerCase();
    if (aliasKey in instrumentAliases) {
      instrument = instrumentAliases[aliasKey];
    }

    instrument = instrument
      .split(/\s+/)
      .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
      .join(' ');

    const expansionKey = instrument.toLowerCase();
    if (expansionKey in brassExpansion) {
      brassExpansion[expansionKey].forEach(subInstrument => {
        results.push({ instrument: subInstrument, levels });
      });
    } else {
      results.push({ instrument, levels });
    }
  });

  return results;
}

function processTeacherResponses() {
  const responsesSheet = UtilityScriptLibrary.getSheet('teacherResponses');
  const fieldMapSheet = UtilityScriptLibrary.getSheet('teacherFieldMap');
  const contactsSheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
  const instrumentSheet = UtilityScriptLibrary.getSheet('instrumentList');
  const futureSheet = UtilityScriptLibrary.getSheet('futureTeachers');

  const fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
  const headerMap = UtilityScriptLibrary.getHeaderMap(responsesSheet);
  const existingInstrumentKeys = getExistingInstrumentKeys(instrumentSheet);

  const rows = responsesSheet.getRange(2, 1, responsesSheet.getLastRow() - 1, responsesSheet.getLastColumn()).getValues();

  rows.forEach(row => {
    const get = (internalField) => {
      const formHeader = Object.keys(fieldMap).find(key => fieldMap[key] === internalField);
      if (!formHeader) return '';
      const normHeader = UtilityScriptLibrary.normalizeHeader(formHeader);
      const colIndex = headerMap[normHeader] - 1;
      return row[colIndex] ?? '';
    };

    const interest = get('Interest').toString().trim().toLowerCase();
    const teachAtOP = get('Teach at OP').toString().trim().toLowerCase();

    if (interest === 'no') {
      addFutureProspect({
        Timestamp: get('Timestamp'),
        FirstName: get('First Name'),
        LastName: get('Last Name'),
        Salutation: get('Salutation'),
        Email: get('Email'),
        Phone: get('Phone')
      }, futureSheet);
      return;
    }

    // Yes or Maybe
    if (!teachAtOP.startsWith('y')) {
      Logger.log(`Discarded: ${get('First Name')} ${get('Last Name')} — unwilling to teach at OP`);
      return;
    }

    const entry = {
      FirstName: get('First Name'),
      LastName: get('Last Name'),
      Salutation: get('Salutation'),
      Email: get('Email'),
      SecondaryEmail: get('E-mail 2'),
      Phone: get('Phone'),
      Street: get('Street'),
      City: get('City'),
      Zip: get('Zip'),
      Role: 'Teacher'
    };

    const contactObj = transformToContact(entry, contactsSheet);
    appendContactRow(contactObj, contactsSheet);

    const contact = getContactByKey(contactsSheet, entry.FirstName, entry.LastName);
    if (!contact || !contact['teacher id']) return;

    const teacherId = contact['teacher id'];
    const instrumentRaw = get('Instruments');
    const summer = get('Summer').toString().toLowerCase().startsWith('y');
    const schoolYear = get('School Year').toString().toLowerCase().startsWith('y');
    const teachAtOPBool = true; // already filtered above

    const parsedInstruments = parseInstrumentsWithLevels(instrumentRaw);
    parsedInstruments.forEach(({ instrument, levels }) => {
      const rowObj = buildInstrumentRow({
        instrument,
        level: levels,
        firstName: entry.FirstName,
        lastName: entry.LastName,
        teacherId,
        teachAtOP: teachAtOPBool,
        summer,
        schoolYear
      });
      updateOrInsertRow(instrumentSheet, rowObj, existingInstrumentKeys);
    });
  });

  setCheckboxColumns(instrumentSheet, ['Teach at OP', 'Summer', 'School Year']);
}

function runLogHeaders() {
  UtilityScriptLibrary.logAllSheetHeaders();
}

function setCheckboxColumns(sheet, headerNames) {
  const headers = UtilityScriptLibrary.getColumnHeaders(sheet);
  headerNames.forEach(name => {
    const colIndex = headers.indexOf(name);
    if (colIndex !== -1) {
      const range = sheet.getRange(2, colIndex + 1, sheet.getMaxRows() - 1);
      const rule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .setAllowInvalid(false)
        .build();
      range.setDataValidation(rule);
    }
  });
}

function transformToContact(entry, contactsSheet) {
  const role = ['Teacher', 'Both'].includes(entry.Role) ? entry.Role : 'Admin';
  const teacherId = role !== 'Admin'
    ? UtilityScriptLibrary.generateNextId(contactsSheet, 'Teacher ID', 'T')
    : '';
  const contactKey = UtilityScriptLibrary.generateKey(entry.FirstName, entry.LastName);

  return {
    'Teacher ID': teacherId,
    'Last Name': entry.LastName || '',
    'First Name': entry.FirstName || '',
    'Salutation': entry.Salutation || '',
    'Email': entry.Email || '',
    'E-mail 2': entry.SecondaryEmail || '',
    'Phone': UtilityScriptLibrary.formatPhoneNumber(entry.Phone || ''),
    'Address': UtilityScriptLibrary.formatAddress(entry.Street || '', entry.City || '', entry.Zip || ''),
    'Role': role,
    'Key': contactKey
  };
}

function updateOrInsertRow(sheet, rowObj, existingKeys) {
  const headers = UtilityScriptLibrary.getColumnHeaders(sheet);
  const key = `${rowObj['Teacher ID']}_${rowObj['Instrument']}`.toLowerCase();
  const rowIndex = existingKeys[key];

  const values = headers.map(h => rowObj[h] ?? '');

  if (rowIndex) {
    const existingRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const changed = headers.some((h, i) => {
      if (['Level', 'Teach at OP', 'Summer', 'School Year'].includes(h)) {
        return existingRow[i] !== rowObj[h];
      }
      return false;
    });

    if (changed) {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([values]);
    }
  } else {
    sheet.appendRow(values);
  }
}