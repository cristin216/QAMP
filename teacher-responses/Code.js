

// === UI MENU ===
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Teacher Tools')
    .addItem('Process Teacher', 'processTeacherManually')
    .addToUi();
}

function processTeacherManually() {
  try {
    UtilityScriptLibrary.debugLog("=== MANUAL BATCH PROCESSING INITIATED ===");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var responsesSheet = ss.getActiveSheet();
    var trackingSheet = ss.getSheetByName("Teacher Tracking");
    
    // Get headers from responses sheet
    var headers = UtilityScriptLibrary.getColumnHeaders(responsesSheet);
    var lastRow = responsesSheet.getLastRow();
    
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('Info', 'No data to process.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get all response data
    var allData = responsesSheet.getRange(2, 1, lastRow - 1, responsesSheet.getLastColumn()).getValues();
    
    // Get existing tracked teachers
    var trackedNames = {};
    if (trackingSheet) {
      var trackingData = trackingSheet.getDataRange().getValues();
      for (var i = 1; i < trackingData.length; i++) {
        var name = trackingData[i][2]; // Column C: Name
        if (name) {
          trackedNames[String(name).trim()] = true;
        }
      }
    }
    
    UtilityScriptLibrary.debugLog("Found " + Object.keys(trackedNames).length + " already tracked teachers");
    
    // Process each untracked row
    var processed = 0;
    var skipped = 0;
    var errors = [];
    
    for (var i = 0; i < allData.length; i++) {
      var rowData = allData[i];
      var formData = {};
      
      // Build formData object from row
      for (var j = 0; j < headers.length; j++) {
        formData[headers[j]] = rowData[j];
      }
      
      // Build name to check against tracking
      var firstName = formData["First Name"] || '';
      var lastName = formData["Last Name"] || '';
      var fullName = (firstName + ' ' + lastName).trim();
      
      if (!fullName || trackedNames[fullName]) {
        skipped++;
        UtilityScriptLibrary.debugLog("Skipping already tracked: " + fullName);
        continue;
      }
      
      // Process this teacher
      try {
        UtilityScriptLibrary.debugLog("Processing untracked teacher: " + fullName);
        
        // Check if interested in teaching
        var interest = formData["Are you interested in teaching lessons with the Quaker Arts Music Program?"] || '';
        
        if (interest !== "Yes" && interest !== "Maybe") {
          UtilityScriptLibrary.debugLog("Skipping - not interested in teaching: " + fullName);
          
          // Track future prospect if applicable
          var futureInterest = formData["Would you like us to keep your name on file for future opportunities?"] || '';
          if (futureInterest === "Yes") {
            trackFutureProspect(formData);
          }
          
          skipped++;
          continue;
        }
        
        // Process the teacher
        var contactsWorkbook = UtilityScriptLibrary.getWorkbook('contacts');
        if (!contactsWorkbook) {
          throw new Error("Could not access Contacts workbook");
        }
        
        var teachersSheet = contactsWorkbook.getSheetByName("Teachers and Admin");
        if (!teachersSheet) {
          throw new Error("Teachers and Admin sheet not found");
        }
        
        var teacherId = processTeacher(formData, teachersSheet);
        
        var instrumentSheet = contactsWorkbook.getSheetByName("Instrument List");
        if (!instrumentSheet) {
          throw new Error("Instrument List sheet not found");
        }
        
        processTeacherInstruments(formData, instrumentSheet, teacherId);
        addOrUpdateTeacherRosterLookup(formData, teacherId);
        updateLocalTeacherTracking(formData, teacherId);
        
        processed++;
        UtilityScriptLibrary.debugLog("âœ… Processed: " + fullName);
        
      } catch (error) {
        errors.push(fullName + ": " + error.message);
        UtilityScriptLibrary.debugLog("âŒ Error processing " + fullName + ": " + error.message);
      }
    }
    
    // Show summary
    var message = "Processing complete!\n\n" +
                  "Processed: " + processed + "\n" +
                  "Skipped: " + skipped;
    
    if (errors.length > 0) {
      message += "\n\nErrors (" + errors.length + "):\n" + errors.join("\n");
    }
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('Batch Processing Complete', message, ui.ButtonSet.OK);
    UtilityScriptLibrary.debugLog("=== BATCH PROCESSING COMPLETE ===");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("âŒ Error in manual batch processing: " + error.message);
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Failed to process teachers: ' + error.message, ui.ButtonSet.OK);
  }
}

function handleTeacherFormSubmit(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Could not obtain lock - another execution in progress. EXITING.");
    return;
  }

  try {
    UtilityScriptLibrary.debugLog("=== STARTING handleTeacherFormSubmit ===");
    
    var formData = extractTeacherFormData(e);
    UtilityScriptLibrary.debugLog("Extracted teacher form data: " + JSON.stringify(formData));

    // **NEW: Check if this person actually wants to teach**
    var interest = formData["Are you interested in teaching lessons with the Quaker Arts Music Program?"] || '';
    
    if (interest !== "Yes" && interest !== "Maybe") {
      UtilityScriptLibrary.debugLog("⏭️ Skipping processing - respondent not interested in teaching now (Interest: '" + interest + "')");
      
      // Optionally track future prospects separately
      var futureInterest = formData["Would you like us to keep your name on file for future opportunities?"] || '';
      if (futureInterest === "Yes") {
        UtilityScriptLibrary.debugLog("📋 This person wants to be contacted for future opportunities");
        // Optional: Add them to a "Future Prospects" tracking
        trackFutureProspect(formData);
      }
      
      UtilityScriptLibrary.debugLog("✅ Form processing complete (no teacher processing needed)");
      return; // Exit early - don't process as teacher
    }
    
    UtilityScriptLibrary.debugLog("✅ Proceeding with teacher processing (Interest: '" + interest + "')");

    // Get Contacts workbook (external)
    var contactsWorkbook = UtilityScriptLibrary.getWorkbook('contacts');
    if (!contactsWorkbook) {
      throw new Error("Could not access Contacts workbook");
    }

    // Process teacher in Teachers and Admin sheet
    var teachersSheet = contactsWorkbook.getSheetByName("Teachers and Admin");
    if (!teachersSheet) {
      throw new Error("Teachers and Admin sheet not found in Contacts workbook");
    }

    var teacherId = processTeacher(formData, teachersSheet);
    UtilityScriptLibrary.debugLog("✅ Teacher processed - ID: " + teacherId);

    // Process instruments in Instrument List sheet
    var instrumentSheet = contactsWorkbook.getSheetByName("Instrument List");
    if (!instrumentSheet) {
      throw new Error("Instrument List sheet not found in Contacts workbook");
    }

    processTeacherInstruments(formData, instrumentSheet, teacherId);
    UtilityScriptLibrary.debugLog("✅ Teacher instruments processed");

    // NEW: Add/update teacher in Teacher Roster Lookup
    addOrUpdateTeacherRosterLookup(formData, teacherId);
    UtilityScriptLibrary.debugLog("✅ Teacher roster lookup updated");

    // Update local tracking sheet
    updateLocalTeacherTracking(formData, teacherId);
    UtilityScriptLibrary.debugLog("✅ Local tracking updated");

    UtilityScriptLibrary.debugLog("✅ Completed handleTeacherFormSubmit");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in handleTeacherFormSubmit: " + error.message);
    UtilityScriptLibrary.debugLog("❌ Stack trace: " + error.stack);
  } finally {
    lock.releaseLock();
  }
}

// === PROCESSING ===

function extractTeacherFormData(e) {
  var formData = {};
  
  if (e && e.namedValues) {
    // Form submission event
    for (var key in e.namedValues) {
      formData[key] = e.namedValues[key][0]; // Get first value
    }
  } else {
    // Manual processing - get from active sheet
    var sheet = SpreadsheetApp.getActiveSheet();
    var headers = UtilityScriptLibrary.getColumnHeaders(sheet);
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
  var level = instrumentData.level || "all";
  
  // Check if this teacher/instrument combination already exists
  var existingRow = findInstrumentRow(instrumentSheet, firstName, lastName, instrument);
  
  if (existingRow !== -1) {
    UtilityScriptLibrary.debugLog("🔄 Updating existing instrument entry: " + instrument);
    // Update existing entry
    instrumentSheet.getRange(existingRow, getCol("Status")).setValue("Potential");
    instrumentSheet.getRange(existingRow, getCol("Level")).setValue(level);
    updateInstrumentAvailability(instrumentSheet, existingRow, instrumentData, getCol);
  } else {
    UtilityScriptLibrary.debugLog("➕ Creating new instrument entry: " + instrument);
    // Create new entry
    var newRow = new Array(instrumentSheet.getLastColumn()).fill('');
    
    newRow[getCol("Instrument") - 1] = instrument;
    newRow[getCol("First Name") - 1] = firstName;
    newRow[getCol("Last Name") - 1] = lastName;
    newRow[getCol("Status") - 1] = "Potential";
    newRow[getCol("Level") - 1] = level;
    newRow[getCol("Teacher ID") - 1] = teacherId;
    
    // Set availability checkboxes (removed "future")
    setInstrumentAvailability(newRow, instrumentData, getCol);
    
    instrumentSheet.appendRow(newRow);
  }
}

function processTeacher(formData, teachersSheet) {
  try {
    UtilityScriptLibrary.debugLog("👨‍🏫 Starting processTeacher");
    
    // Get the field map for teacher processing
    var fieldMapSheet = UtilityScriptLibrary.getSheet('teacherFieldMap');
    if (!fieldMapSheet) {
      throw new Error("Teacher Field Map sheet not found");
    }
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
    UtilityScriptLibrary.debugLog("Field map loaded: " + JSON.stringify(fieldMap));
    
    // Create helper function to get mapped form values - FIXED CASE SENSITIVITY
    var get = function(internalField) {
      var formHeader = Object.keys(fieldMap).find(key => fieldMap[key] === internalField);
      if (!formHeader) {
        UtilityScriptLibrary.debugLog("⚠️ No form header found for internal field: " + internalField);
        return '';
      }
      
      // Need to find the actual form data key that matches the normalized form header
      var actualFormKey = Object.keys(formData).find(key => 
        UtilityScriptLibrary.normalizeHeader(key) === formHeader
      );
      
      var value = actualFormKey ? formData[actualFormKey] : '';
      UtilityScriptLibrary.debugLog("Mapped " + internalField + " → " + formHeader + " (actual key: " + actualFormKey + ") = '" + value + "'");
      return value;
    };
    
    var getCol = UtilityScriptLibrary.createColumnFinder(teachersSheet);

    // Create teacher lookup key using mapped fields
    var firstName = get("First Name");
    var lastName = get("Last Name");
    var teacherKey = UtilityScriptLibrary.generateKey(firstName, lastName);

    var teacherRow = findTeacherRow(teachersSheet, teacherKey);
    
    UtilityScriptLibrary.debugLog("=== TEACHER DUPLICATE CHECK ===");
    UtilityScriptLibrary.debugLog("Looking for key: '" + teacherKey + "'");
    UtilityScriptLibrary.debugLog("findTeacherRow result: " + teacherRow);
    UtilityScriptLibrary.debugLog("=== END TEACHER DEBUG ===");

    var teacherId = '';

    if (teacherRow !== -1) {
      UtilityScriptLibrary.debugLog("Updating existing teacher");
      var rowValues = teachersSheet.getRange(teacherRow, 1, 1, teachersSheet.getLastColumn()).getValues()[0];
      teacherId = String(rowValues[getCol("Teacher ID") - 1] || '');

      // Update fields that might have changed
      updateTeacherFields(teachersSheet, teacherRow, fieldMap, get, getCol);

    } else {
      UtilityScriptLibrary.debugLog("➕ Creating new teacher");
      teacherId = UtilityScriptLibrary.generateNextId(teachersSheet, 'Teacher ID', 'T');
      
      var headers = teachersSheet.getRange(1, 1, 1, teachersSheet.getLastColumn()).getValues()[0];
      var newRow = new Array(headers.length).fill('');
      
      // Set required fields
      newRow[getCol("Teacher ID") - 1] = teacherId;
      newRow[getCol("First Name") - 1] = firstName;
      newRow[getCol("Last Name") - 1] = lastName;
      newRow[getCol("Key") - 1] = teacherKey;
      newRow[getCol("Role") - 1] = "Teacher";
      
      // Set form fields using field mapping
      if (getCol("Salutation")) {
        newRow[getCol("Salutation") - 1] = get("Salutation");
      }
      
      if (getCol("Email")) {
        newRow[getCol("Email") - 1] = get("Email");
      }
      
      if (getCol("E-mail 2")) {
        newRow[getCol("E-mail 2") - 1] = get("E-mail 2");
      }
      
      if (getCol("Phone")) {
        newRow[getCol("Phone") - 1] = UtilityScriptLibrary.formatPhoneNumber(get("Phone"));
      }
      
      if (getCol("Address")) {
        var street = get("Street");
        var city = get("City");
        var zip = get("Zip");
        newRow[getCol("Address") - 1] = UtilityScriptLibrary.formatAddress(street, city, zip);
      }
      
      // Add Display Name field (clean last name for roster purposes)
      if (getCol("Display Name")) {
        var displayName = UtilityScriptLibrary.createDisplayName(lastName);
        newRow[getCol("Display Name") - 1] = displayName;
      }
      
      teachersSheet.appendRow(newRow);
    }

    UtilityScriptLibrary.debugLog("✅ Completed processTeacher - ID: " + teacherId);
    return teacherId;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in processTeacher: " + error.message);
    throw error;
  }
}


function processTeacherInstruments(formData, instrumentSheet, teacherId) {
  try {
    UtilityScriptLibrary.debugLog("ðŸŽµ Starting processTeacherInstruments");
    
    // Extract instruments from form data
    var instruments = extractInstrumentsList(formData);
    UtilityScriptLibrary.debugLog("Extracted instruments: " + JSON.stringify(instruments));
    
    if (instruments.length === 0) {
      UtilityScriptLibrary.debugLog("âš ï¸ No instruments found in form data");
      return;
    }

    var getCol = UtilityScriptLibrary.createColumnFinder(instrumentSheet);

    // Get field map and create get function for proper field mapping
    var fieldMapSheet = UtilityScriptLibrary.getSheet('teacherFieldMap');
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
    var get = function(internalField) {
      var formHeader = Object.keys(fieldMap).find(key => fieldMap[key] === internalField);
      if (!formHeader) return '';
      var actualFormKey = Object.keys(formData).find(key => 
        UtilityScriptLibrary.normalizeHeader(key) === formHeader
      );
      return actualFormKey ? formData[actualFormKey] : '';
    };

    var firstName = get("First Name");
    var lastName = get("Last Name");

    // Process each instrument
    for (var i = 0; i < instruments.length; i++) {
      var instrumentData = instruments[i];
      processSingleInstrument(instrumentSheet, instrumentData, firstName, lastName, teacherId, getCol);
    }
    
    UtilityScriptLibrary.debugLog("âœ… Completed processTeacherInstruments");
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("âŒ Error in processTeacherInstruments: " + error.message);
  }
}

// === HELPER FUNCTIONS ===

function addOrUpdateTeacherRosterLookup(formData, teacherId) {
  try {
    UtilityScriptLibrary.debugLog("📋 Starting addOrUpdateTeacherRosterLookup");
    
    var firstName = formData["First Name"] || formData["Teacher First Name"] || '';
    var lastName = formData["Last Name"] || formData["Teacher Last Name"] || '';
    var teacherName = firstName + ' ' + lastName;
    var displayName = UtilityScriptLibrary.createDisplayName(lastName);
    
    UtilityScriptLibrary.debugLog("Teacher: " + teacherName + ", Display Name: " + displayName);
    
    var responsesWorkbook = UtilityScriptLibrary.getWorkbook('formResponses');
    var lookupSheet = responsesWorkbook.getSheetByName("Teacher Roster Lookup");
    
    if (!lookupSheet) {
      UtilityScriptLibrary.debugLog("Creating new Teacher Roster Lookup sheet");
      lookupSheet = createTeacherRosterLookupSheet(responsesWorkbook);
    }
    
    var existingRow = findTeacherInRosterLookup(lookupSheet, teacherName);
    
    if (existingRow === -1) {
      // Add new teacher - UPDATED with Group Assignment column and lowercase status
      lookupSheet.appendRow([
        teacherName,     // A: Teacher Name
        '',              // B: Roster URL (empty until roster created)  
        teacherId,       // C: Teacher ID
        displayName,     // D: Display Name (clean last name)
        '',              // E: Group Assignment (empty)
        'potential',     // F: Status (lowercase - potential until they get students)
        new Date()       // G: Last Updated
      ]);
      UtilityScriptLibrary.debugLog("✅ Added new teacher to roster lookup: " + teacherName);
    } else {
      // Update existing teacher - UPDATED column references
      lookupSheet.getRange(existingRow, 3).setValue(teacherId);    // C: Teacher ID
      lookupSheet.getRange(existingRow, 4).setValue(displayName);  // D: Display Name
      lookupSheet.getRange(existingRow, 7).setValue(new Date());   // G: Last Updated
      UtilityScriptLibrary.debugLog("✅ Updated existing teacher in roster lookup: " + teacherName);
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in addOrUpdateTeacherRosterLookup: " + error.message);
    // Don't throw - this shouldn't break the main teacher processing
  }
}

function createTeacherRosterLookupSheet(workbook) {
  try {
    var sheet = workbook.insertSheet("Teacher Roster Lookup");
    
    // Set headers - UPDATED with Group Assignment column
    var headers = [
      "Teacher Name",      // A: Full name (John Smith)
      "Roster URL",        // B: Link to teacher's roster workbook
      "Teacher ID",        // C: System ID (T0015)
      "Display Name",      // D: Clean last name for backend (Smith)
      "Group Assignment",  // E: Group assignment (empty for now)
      "Status",            // F: potential/active/inactive (lowercase)
      "Last Updated"       // G: Timestamp
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply header styling
    try {
      UtilityScriptLibrary.styleHeaderRow(sheet, headers);
    } catch (styleError) {
      // Fallback styling if utility function fails - use STYLES constants
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold")
                 .setBackground(UtilityScriptLibrary.STYLES.HEADER.background)
                 .setFontColor(UtilityScriptLibrary.STYLES.HEADER.text);
    }
    
    UtilityScriptLibrary.debugLog("✅ Created Teacher Roster Lookup sheet with headers");
    return sheet;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error creating Teacher Roster Lookup sheet: " + error.message);
    throw error;
  }
}

function extractInstrumentsList(formData) {
  try {
    var instruments = [];
    
    // Get the field map for teacher processing
    var fieldMapSheet = UtilityScriptLibrary.getSheet('teacherFieldMap');
    if (!fieldMapSheet) {
      UtilityScriptLibrary.debugLog("⚠️ Teacher Field Map sheet not found - using direct field access");
      return instruments;
    }
    var fieldMap = UtilityScriptLibrary.getFieldMappingFromSheet(fieldMapSheet);
    
    // Create helper function to get mapped form values - FIXED CASE SENSITIVITY
    var get = function(internalField) {
      var formHeader = Object.keys(fieldMap).find(key => fieldMap[key] === internalField);
      if (!formHeader) return '';
      
      // Need to find the actual form data key that matches the normalized form header
      var actualFormKey = Object.keys(formData).find(key => 
        UtilityScriptLibrary.normalizeHeader(key) === formHeader
      );
      
      return actualFormKey ? formData[actualFormKey] : '';
    };
    
    // Get instruments using field mapping
    var instrumentsText = get("Instruments");
    UtilityScriptLibrary.debugLog("Looking for instruments in field value: '" + instrumentsText + "'");
    
    if (instrumentsText) {
      // Split by comma to handle multiple instruments
      var instrumentEntries = instrumentsText.split(',');
      
      for (var i = 0; i < instrumentEntries.length; i++) {
        var entry = instrumentEntries[i].trim();
        if (entry) {
          var instrument = '';
          var level = 'all'; // default level
          
          // Parse different formats:
          // violin-beg/int
          // viola-beg
          // piano (adv)
          // cello - all
          
          // Check for dash format: "violin-beg/int"
          if (entry.includes('-')) {
            var parts = entry.split('-');
            instrument = parts[0].trim();
            level = parts[1] ? parts[1].trim() : 'all';
          }
          // Check for parentheses format: "piano (adv)"
          else if (entry.includes('(') && entry.includes(')')) {
            var parenMatch = entry.match(/^(.+?)\s*\((.+?)\)$/);
            if (parenMatch) {
              instrument = parenMatch[1].trim();
              level = parenMatch[2].trim();
            } else {
              instrument = entry;
              level = 'all';
            }
          }
          // Default: assume entire entry is instrument name
          else {
            instrument = entry;
            level = 'all';
          }
          
          // Normalize instrument name (capitalize first letter)
          instrument = instrument.charAt(0).toUpperCase() + instrument.slice(1).toLowerCase();
          
          // Add to instruments array with mapped availability fields
          instruments.push({
            instrument: instrument,
            level: level,
            teachAtOP: get("Teach at OP") === "Yes",
            summer: get("Summer") === "Yes",
            schoolYear: get("School Year") === "Yes",
            future: get("Future") === "Yes"
          });
          
          UtilityScriptLibrary.debugLog("Parsed instrument: " + instrument + " with level: " + level);
        }
      }
    }
    
    UtilityScriptLibrary.debugLog("Final extracted instruments: " + JSON.stringify(instruments));
    return instruments;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in extractInstrumentsList: " + error.message);
    return [];
  }
}

function findInstrumentRow(sheet, firstName, lastName, instrument) {
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === instrument && 
        String(data[i][1]) === firstName && 
        String(data[i][2]) === lastName) {
      return i + 1;
    }
  }
  
  return -1;
}

function findTeacherInRosterLookup(lookupSheet, teacherName) {
  try {
    var data = lookupSheet.getDataRange().getValues();
    
    // Search for teacher by name (column A)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === teacherName.trim()) {
        return i + 1; // Return 1-based row number
      }
    }
    
    return -1; // Not found
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error finding teacher in roster lookup: " + error.message);
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

function generateUniqueDisplayNames(baseLastName, teachers) {
  var displayNames = {};
  var usedNames = {};
  
  // Sort teachers by first name for consistent results
  teachers.sort(function(a, b) {
    return a.firstName.localeCompare(b.firstName);
  });
  
  // First pass: try single initials
  var needsMoreChars = [];
  
  for (var i = 0; i < teachers.length; i++) {
    var teacher = teachers[i];
    var firstName = teacher.firstName;
    var firstInitial = firstName.charAt(0).toUpperCase();
    
    // Count how many teachers start with this initial
    var sameInitialCount = 0;
    for (var j = 0; j < teachers.length; j++) {
      if (teachers[j].firstName.charAt(0).toUpperCase() === firstInitial) {
        sameInitialCount++;
      }
    }
    
    if (sameInitialCount === 1) {
      // Only teacher with this initial - can use just the initial
      var displayName = baseLastName + firstInitial;
      displayNames[firstName] = displayName;
      usedNames[displayName] = true;
      UtilityScriptLibrary.debugLog("🏷️ " + firstName + " " + baseLastName + " → " + displayName + " (unique initial)");
    } else {
      // Multiple teachers with this initial - need more characters
      needsMoreChars.push(teacher);
    }
  }
  
  // Second pass: handle teachers that need more than just initial
  for (var i = 0; i < needsMoreChars.length; i++) {
    var teacher = needsMoreChars[i];
    var firstName = teacher.firstName;
    
    // Start with 2 characters and increase until unique
    var charIndex = 2;
    var displayName = baseLastName + firstName.substring(0, charIndex);
    
    while (usedNames[displayName] && charIndex <= firstName.length) {
      charIndex++;
      displayName = baseLastName + firstName.substring(0, charIndex);
    }
    
    // If we've used the whole first name and it's still not unique, add numbers
    var counter = 2;
    var baseDisplayName = displayName;
    while (usedNames[displayName]) {
      displayName = baseDisplayName + counter;
      counter++;
    }
    
    displayNames[firstName] = displayName;
    usedNames[displayName] = true;
    
    UtilityScriptLibrary.debugLog("🏷️ " + firstName + " " + baseLastName + " → " + displayName + " (extended for uniqueness)");
  }
  
  return displayNames;
}

function setInstrumentAvailability(row, instrumentData, getCol) {
  if (getCol("Teach at OP")) row[getCol("Teach at OP") - 1] = instrumentData.teachAtOP || false;
  if (getCol("Summer")) row[getCol("Summer") - 1] = instrumentData.summer || false;
  if (getCol("School Year")) row[getCol("School Year") - 1] = instrumentData.schoolYear || false;
}

function trackFutureProspect(formData) {
  try {
    UtilityScriptLibrary.debugLog("📋 Tracking future prospect");
    
    // Use UtilityScriptLibrary to get the sheet from Contacts workbook
    var prospectSheet = UtilityScriptLibrary.getSheet('futureTeachers');
    
    if (!prospectSheet) {
      UtilityScriptLibrary.debugLog("⚠️ Future Teacher Contacts sheet not found - skipping tracking");
      return;
    }
    
    // Add the prospect to existing sheet - default to "" if not found
    var timestamp = formData["Timestamp"] || new Date();
    var firstName = formData["First Name"] || "";
    var lastName = formData["Last Name"] || "";
    var email = formData["Which email do you prefer we use to contact you and share with students?"] || "";
    var phone = formData["Phone number"] || "";
    var primaryInstrument = formData["What is your primary instrument?"] || "";
    
    prospectSheet.appendRow([
      timestamp,
      firstName,
      lastName,
      email,
      phone,
      primaryInstrument,
      "Future opportunity - not currently teaching"
    ]);
    
    UtilityScriptLibrary.debugLog("✅ Added future prospect: " + firstName + " " + lastName);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error tracking future prospect: " + error.message);
    // Don't throw - this shouldn't break the main flow
  }
}

function updateDisplayNameInRosterLookup(teacherFullName, newDisplayName) {
  try {
    var responsesWorkbook = UtilityScriptLibrary.getWorkbook('formResponses');
    var lookupSheet = responsesWorkbook.getSheetByName("Teacher Roster Lookup");
    
    if (!lookupSheet) return;
    
    var teacherRow = findTeacherInRosterLookup(lookupSheet, teacherFullName);
    
    if (teacherRow !== -1) {
      lookupSheet.getRange(teacherRow, 4).setValue(newDisplayName); // Column D: Display Name
      lookupSheet.getRange(teacherRow, 6).setValue(new Date());     // Column F: Last Updated
      UtilityScriptLibrary.debugLog("✅ Updated display name in Teacher Roster Lookup: " + newDisplayName);
    }
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error updating display name in roster lookup: " + error.message);
  }
}

function updateInstrumentAvailability(sheet, row, instrumentData, getCol) {
  if (getCol("Teach at OP")) sheet.getRange(row, getCol("Teach at OP")).setValue(instrumentData.teachAtOP || false);
  if (getCol("Summer")) sheet.getRange(row, getCol("Summer")).setValue(instrumentData.summer || false);
  if (getCol("School Year")) sheet.getRange(row, getCol("School Year")).setValue(instrumentData.schoolYear || false);
}

function updateLocalTeacherTracking(formData, teacherId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var trackingSheet = ss.getSheetByName("Teacher Tracking");
    
    if (!trackingSheet) {
      trackingSheet = ss.insertSheet("Teacher Tracking");
      trackingSheet.getRange(1, 1, 1, 4).setValues([["Timestamp", "Teacher ID", "Name", "Status"]]);
    }
    
    var name = (formData["First Name"] || '') + ' ' + (formData["Last Name"] || '');
    trackingSheet.appendRow([new Date(), teacherId, name, "Processed"]);
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("⚠️ Could not update local tracking: " + error.message);
  }
}

function updateTeacherDisplayNames(newTeacherData) {
  try {
    UtilityScriptLibrary.debugLog("🔄 Checking for display name conflicts with new teacher: " + newTeacherData.lastName);
    
    var contactsWorkbook = UtilityScriptLibrary.getWorkbook('contacts');
    var teachersSheet = contactsWorkbook.getSheetByName('Teachers and Admin');
    
    if (!teachersSheet) {
      UtilityScriptLibrary.debugLog("❌ Teachers and Admin sheet not found");
      return;
    }
    
    var data = teachersSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var lastNameCol = -1;
    var firstNameCol = -1;
    var displayNameCol = -1;
    var teacherIdCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var normalized = UtilityScriptLibrary.normalizeHeader(headers[i]);
      if (normalized === 'last name') lastNameCol = i;
      if (normalized === 'first name') firstNameCol = i;
      if (normalized === 'display name') displayNameCol = i;
      if (normalized === 'teacher id') teacherIdCol = i;
    }
    
    if (lastNameCol === -1 || firstNameCol === -1 || displayNameCol === -1) {
      UtilityScriptLibrary.debugLog("❌ Required columns not found");
      return;
    }
    
    var baseLastName = newTeacherData.lastName;
    var newFirstName = newTeacherData.firstName;
    
    // Find all teachers with the same base last name
    var matchingTeachers = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowLastName = (row[lastNameCol] || '').toString().trim();
      var rowFirstName = (row[firstNameCol] || '').toString().trim();
      var rowDisplayName = (row[displayNameCol] || '').toString().trim();
      var rowTeacherId = (row[teacherIdCol] || '').toString().trim();
      
      if (rowLastName === baseLastName && rowTeacherId) { // Only active teachers with IDs
        matchingTeachers.push({
          rowIndex: i + 1, // 1-based for sheet operations
          lastName: rowLastName,
          firstName: rowFirstName,
          currentDisplayName: rowDisplayName,
          teacherId: rowTeacherId
        });
      }
    }
    
    UtilityScriptLibrary.debugLog("📋 Found " + matchingTeachers.length + " teachers with last name '" + baseLastName + "'");
    
    if (matchingTeachers.length === 0) {
      // First teacher with this last name
      UtilityScriptLibrary.debugLog("📝 First teacher with this last name, using: " + baseLastName);
      return baseLastName;
    }
    
    // Generate display names for all teachers with this last name
    var allTeachers = matchingTeachers.slice(); // Copy existing teachers
    allTeachers.push({ firstName: newFirstName, lastName: baseLastName }); // Add new teacher
    
    var displayNames = generateUniqueDisplayNames(baseLastName, allTeachers);
    
    // Update existing teachers if needed
    for (var i = 0; i < matchingTeachers.length; i++) {
      var teacher = matchingTeachers[i];
      var newDisplayName = displayNames[teacher.firstName];
      
      if (newDisplayName !== teacher.currentDisplayName) {
        // Update in Teachers and Admin sheet
        teachersSheet.getRange(teacher.rowIndex, displayNameCol + 1).setValue(newDisplayName);
        UtilityScriptLibrary.debugLog("✅ Updated existing teacher: " + teacher.firstName + " " + teacher.lastName + " → " + newDisplayName);
        
        // Update in Teacher Roster Lookup
        updateDisplayNameInRosterLookup(teacher.firstName + " " + teacher.lastName, newDisplayName);
      }
    }
    
    // Return the display name for the new teacher
    var newDisplayName = displayNames[newFirstName];
    UtilityScriptLibrary.debugLog("📝 New teacher should use display name: " + newDisplayName);
    return newDisplayName;
    
  } catch (error) {
    UtilityScriptLibrary.debugLog("❌ Error in updateTeacherDisplayNames: " + error.message);
    return newTeacherData.lastName; // Fallback to just last name
  }
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


