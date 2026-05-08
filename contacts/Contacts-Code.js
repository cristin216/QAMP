function migrateStudentTeacherDisplayNamesToIds() {
  var norm   = UtilityScriptLibrary.normalizeHeader;
  var stats  = { fixed: 0, skipped: 0, errors: [] };

  try {
    var studentsSheet = UtilityScriptLibrary.getSheet('students');
    var headerMap     = UtilityScriptLibrary.getHeaderMap(studentsSheet);
    var teacherCol    = headerMap[norm('Teacher')];

    if (!teacherCol) throw new Error('Teacher column not found in Students sheet');

    var lastRow = studentsSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No student data to migrate.');
      return;
    }

    var teacherData = studentsSheet.getRange(2, teacherCol, lastRow - 1, 1).getValues();
    var updates     = [];

    for (var r = 0; r < teacherData.length; r++) {
      var current = String(teacherData[r][0] || '').trim();
      if (!current || /^T\d+$/.test(current)) continue;

      var resolved = UtilityScriptLibrary.getTeacherIdByDisplayName(current);
      if (resolved) {
        updates.push({ row: r + 2, id: resolved });
      } else {
        stats.skipped++;
        UtilityScriptLibrary.debugLog('migrateStudentTeacherDisplayNamesToIds', 'WARNING',
          'Could not resolve teacher name', 'Row ' + (r + 2) + ': ' + current, '');
      }
    }

    for (var u = 0; u < updates.length; u++) {
      studentsSheet.getRange(updates[u].row, teacherCol).setValue(updates[u].id);
      stats.fixed++;
    }

    UtilityScriptLibrary.debugLog('migrateStudentTeacherDisplayNamesToIds', 'SUCCESS',
      'Migration complete',
      'Fixed: ' + stats.fixed + ', Unresolved: ' + stats.skipped +
      (stats.errors.length ? ', Errors: ' + stats.errors.join('; ') : ''), '');

    Logger.log('Migration complete. Fixed: ' + stats.fixed + ', Unresolved: ' + stats.skipped);

  } catch (error) {
    UtilityScriptLibrary.debugLog('migrateStudentTeacherDisplayNamesToIds', 'ERROR',
      'Migration failed', '', error.message);
    throw error;
  }
}

function logSheetHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var output = [];
  
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var lastCol = sheet.getLastColumn();
    
    if (lastCol === 0) {
      output.push(name + ': [empty]');
      continue;
    }
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    output.push(name + ': ' + headers.filter(String).join(' | '));
  }
  
  Logger.log(output.join('\n\n'));
}
/*
================================================================================
CONTACTS CODE
================================================================================
Version: 3
Total Functions: 1
Documentation: See Contacts-Functions.md
================================================================================
*/

function onEditContacts(e) {
  try {
    var sheet = e.range.getSheet();

    if (sheet.getName() !== UtilityScriptLibrary.SHEET_MAP.teachersAndAdmin.name) {
      return;
    }

    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      return;
    }

    var newValue = String(e.value || '').trim();
    if (newValue !== 'Former') {
      return;
    }

    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var statusCol = headerMap[UtilityScriptLibrary.normalizeHeader('Status')];
    if (!statusCol || e.range.getColumn() !== statusCol) {
      return;
    }

    var teacherIdCol = headerMap[UtilityScriptLibrary.normalizeHeader('Teacher ID')];
    if (!teacherIdCol) {
      UtilityScriptLibrary.debugLog('onEditContacts', 'ERROR', 'Teacher ID column not found', '', '');
      return;
    }

    var teacherId = String(sheet.getRange(e.range.getRow(), teacherIdCol).getValue() || '').trim();

    if (!teacherId || teacherId.charAt(0) !== 'T') {
      UtilityScriptLibrary.debugLog('onEditContacts', 'INFO', 'Skipping non-teacher row', 'ID: ' + teacherId, '');
      return;
    }

    UtilityScriptLibrary.cascadeFormerStatus(teacherId);

  } catch (error) {
    UtilityScriptLibrary.debugLog('onEditContacts', 'ERROR', 'onEdit failed', '', error.message);
  }
}
