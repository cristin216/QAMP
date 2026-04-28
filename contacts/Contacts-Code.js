// === MAIN FUNCTIONS ===
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
