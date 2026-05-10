/*
================================================================================
CONTACTS CODE
================================================================================
Version: 4
Total Functions: 2
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
