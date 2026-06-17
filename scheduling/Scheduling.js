var DAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

// Room constraint definitions
var ROOM_CONSTRAINTS = {
  'Piano':      'room1',
  'Percussion': 'room1',
  'Harp':       'room2'
};

// Room fallback order for non-constrained instruments
var ROOM_FALLBACK_ORDER = ['room3', 'room2', 'room1'];

// Room display names
var ROOM_NAMES = {
  room1: 'Band Room',
  room2: 'Orchestra Room',
  room3: 'Chorus Room'
};

function assignRoom(teacherId, students, blockStart, blockEnd, calendarEvents) {
  try {
    // Check for hard constraints
    var requiredRoom = getRequiredRoom(students);

    if (requiredRoom) {
      // Hard constraint — only one option
      if (hasConflict(calendarEvents[requiredRoom], blockStart, blockEnd)) {
        throw new Error('Required room (' + requiredRoom + ') is not available at this time');
      }
      saveTeacherPreference(teacherId, requiredRoom);
      return requiredRoom;
    }

    // No hard constraint — check teacher preference first
    var preferredRoom = getTeacherPreference(teacherId);
    if (preferredRoom && !hasConflict(calendarEvents[preferredRoom], blockStart, blockEnd)) {
      return preferredRoom;
    }

    // Preference unavailable or not set — try fallback order
    for (var i = 0; i < ROOM_FALLBACK_ORDER.length; i++) {
      var roomKey = ROOM_FALLBACK_ORDER[i];
      if (!hasConflict(calendarEvents[roomKey], blockStart, blockEnd)) {
        saveTeacherPreference(teacherId, roomKey);
        return roomKey;
      }
    }

    throw new Error('No rooms available for this time slot');

  } catch (error) {
    UtilityScriptLibrary.debugLog('assignRoom', 'ERROR', 'Failed', teacherId, error.message);
    throw error;
  }
}

function buildTeacherWeekSchedule(weekEvents, teacherLastName, weekDates) {
  var schedule   = '';
  var hasLessons = false;

  for (var d = 0; d < weekDates.length; d++) {
    var dayDate    = weekDates[d];
    var dayLessons = [];

    for (var rk in weekEvents) {
      var roomEvents = weekEvents[rk];
      for (var ei = 0; ei < roomEvents.length; ei++) {
        var ev      = roomEvents[ei];
        var evDate  = new Date(ev.start);
        var isMatch = evDate.toDateString() === dayDate.toDateString() &&
                      ev.title.indexOf(teacherLastName) === 0;
        if (isMatch) {
          dayLessons.push({
            title:    ev.title,
            start:    ev.start,
            end:      ev.end,
            roomName: ROOM_NAMES[rk]
          });
        }
      }
    }

    if (!dayLessons.length) continue;

    hasLessons = true;
    dayLessons.sort(function(a, b) { return new Date(a.start) - new Date(b.start); });

    schedule += DAY_NAMES[d] + ' ' + formatDateForEmail(dayDate) + '\n';
    for (var li = 0; li < dayLessons.length; li++) {
      var lesson     = dayLessons[li];
      var titleParts = lesson.title.split(' - ');
      var instrument = titleParts.length >= 2 ? titleParts[1] : '';
      var length     = titleParts.length >= 3 ? titleParts[2] : '';
      schedule += '  ' + formatTimeForEmail(new Date(lesson.start)) +
        ' - ' + instrument + ', ' + length + ', ' + lesson.roomName + '\n';
    }
    schedule += '\n';
  }

  return hasLessons ? schedule : 'No lessons scheduled for this week.';
}

function doGet(e) {
  var teacherId = e.parameter.tid || '';
  var template = HtmlService.createTemplateFromFile('Scheduling');
  template.teacherId = teacherId;
  return template.evaluate()
    .setTitle('QAMP Scheduling')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function findLessonLogRowByEventId(rosterSS, eventId) {
  try {
    var sheets = rosterSS.getSheets();
    var norm   = UtilityScriptLibrary.normalizeHeader;

    for (var i = 0; i < sheets.length; i++) {
      var sheet     = sheets[i];
      var sheetName = sheet.getName();

      // Only search monthly attendance sheets
      if (!UtilityScriptLibrary.isMonthSheet(sheetName)) continue;

      var headerMap  = UtilityScriptLibrary.getHeaderMap(sheet);
      var eventIdCol = headerMap[norm('Calendar Event ID')];
      if (!eventIdCol) continue;

      var data = sheet.getDataRange().getValues();
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][eventIdCol - 1]).trim() === String(eventId).trim()) {
          return { sheet: sheet, row: r + 1 };
        }
      }
    }

    UtilityScriptLibrary.debugLog('findLessonLogRowByEventId', 'WARNING',
      'Event ID not found in any monthly sheet', eventId, '');
    return null;

  } catch (error) {
    UtilityScriptLibrary.debugLog('findLessonLogRowByEventId', 'ERROR', 'Failed', eventId, error.message);
    return null;
  }
}

function findNextBlankLessonRow(monthSheet, studentId) {
  try {
    var data      = monthSheet.getDataRange().getValues();
    var headerMap = UtilityScriptLibrary.getHeaderMap(monthSheet);
    var norm      = UtilityScriptLibrary.normalizeHeader;

    var idCol   = headerMap[norm('Student ID')];
    var dateCol = headerMap[norm('Date')];

    if (!idCol || !dateCol) return null;

    for (var i = 1; i < data.length; i++) {
      var rowStudentId = String(data[i][idCol - 1] || '').trim();
      var rowDate      = data[i][dateCol - 1];
      if (rowStudentId === String(studentId).trim() && !rowDate) {
        return i + 1; // 1-based row number
      }
    }

    return null;

  } catch (error) {
    UtilityScriptLibrary.debugLog('findNextBlankLessonRow', 'ERROR', 'Failed', studentId, error.message);
    return null;
  }
}

function formatDateForEmail(date) {
  return UtilityScriptLibrary.getMonthNames()[date.getMonth()] + ' ' + date.getDate();
}

function formatTimeForEmail(date) {
  var hours   = date.getHours();
  var minutes = date.getMinutes();
  var ampm    = hours >= 12 ? 'pm' : 'am';
  hours       = hours % 12 || 12;
  return hours + (minutes > 0 ? ':' + (minutes < 10 ? '0' : '') + minutes : '') + ampm;
}

function getComingWeekDates() {
  var today  = new Date();
  var day    = today.getDay(); // 0=Sun, 1=Mon...
  var toMon  = (day === 0) ? 1 : (day === 1) ? 7 : (8 - day);
  var monday = new Date(today.getTime() + toMon * 24 * 60 * 60 * 1000);
  monday.setHours(0, 0, 0, 0);

  var dates = [];
  for (var i = 0; i < 5; i++) {
    dates.push(new Date(monday.getTime() + i * 24 * 60 * 60 * 1000));
  }
  return dates;
}

function getRequiredRoom(students) {
  var requiredRoom = null;
  for (var i = 0; i < students.length; i++) {
    var instrument = (students[i].instrument || '').trim();
    var normalizedInstrument = instrument.charAt(0).toUpperCase() + instrument.slice(1).toLowerCase();
    var constraint = ROOM_CONSTRAINTS[normalizedInstrument];
    if (constraint) {
      if (requiredRoom && requiredRoom !== constraint) {
        throw new Error('Conflicting room constraints in block: ' + requiredRoom + ' vs ' + constraint);
      }
      requiredRoom = constraint;
    }
  }
  return requiredRoom;
}

function getSchedulingData(teacherId) {
  try {
    UtilityScriptLibrary.debugLog('getSchedulingData', 'INFO', 'Loading scheduling data', teacherId, '');

    // --- Teacher info ---
    var contactsSheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
    var headerMap = UtilityScriptLibrary.getHeaderMap(contactsSheet);
    var norm = UtilityScriptLibrary.normalizeHeader;

    var idCol        = headerMap[norm('Teacher ID')];
    var firstNameCol = headerMap[norm('First Name')];
    var lastNameCol  = headerMap[norm('Last Name')];
    var emailCol     = headerMap[norm('Email')];

    if (!idCol || !firstNameCol || !lastNameCol || !emailCol) {
      throw new Error('Required columns not found in Teachers and Admin');
    }

    var teacherInfo = null;
    var contactsData = contactsSheet.getDataRange().getValues();
    for (var i = 1; i < contactsData.length; i++) {
      if (String(contactsData[i][idCol - 1]).trim() === String(teacherId).trim()) {
        teacherInfo = {
          teacherId:  teacherId,
          firstName:  String(contactsData[i][firstNameCol - 1] || '').trim(),
          lastName:   String(contactsData[i][lastNameCol - 1]  || '').trim(),
          email:      String(contactsData[i][emailCol - 1]     || '').trim()
        };
        break;
      }
    }

    if (!teacherInfo) {
      throw new Error('Teacher not found: ' + teacherId);
    }

    // --- Roster URL ---
    var lookupSheet     = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    var lookupData      = lookupSheet.getDataRange().getValues();
    var lookupHeaderMap = UtilityScriptLibrary.getHeaderMap(lookupSheet);

    var tidCol    = lookupHeaderMap[norm('Teacher ID')];
    var urlCol    = lookupHeaderMap[norm('Roster URL')];
    var statusCol = lookupHeaderMap[norm('Status')];

    if (!tidCol || !urlCol) {
      throw new Error('Required columns not found in Teacher Roster Lookup');
    }

    var rosterUrl = null;
    for (var j = 1; j < lookupData.length; j++) {
      if (String(lookupData[j][tidCol - 1]).trim() === String(teacherId).trim()) {
        var status = statusCol ? String(lookupData[j][statusCol - 1]).trim().toLowerCase() : '';
        if (status !== 'active') {
          throw new Error('Teacher roster is not active: ' + teacherId);
        }
        rosterUrl = String(lookupData[j][urlCol - 1]).trim();
        break;
      }
    }

    if (!rosterUrl) {
      throw new Error('Roster URL not found for teacher: ' + teacherId);
    }
    teacherInfo.rosterUrl = rosterUrl;

    // --- Students from roster ---
    var rosterSS    = SpreadsheetApp.openByUrl(rosterUrl);
    var rosterSheet = UtilityScriptLibrary.findMostRecentRosterSheet(rosterSS);

    if (!rosterSheet) {
      throw new Error('No roster sheet found for teacher: ' + teacherId);
    }

    var students = UtilityScriptLibrary.extractRosterData(rosterSheet);

    // --- Summer end date from semester metadata ---
    var semesterName  = UtilityScriptLibrary.getCurrentSemesterName(new Date());
    var semesterDates = UtilityScriptLibrary.getSemesterDates(semesterName);
    var today         = new Date();
    var endDate       = semesterDates ? semesterDates.end : new Date(today.getFullYear(), 8, 1);

    // --- Calendar events today → semester end ---
    var env       = UtilityScriptLibrary.EnvironmentManager.get();
    var envConfig = UtilityScriptLibrary.getConfig()[env];

    var rooms = [
      { key: 'room1', name: ROOM_NAMES.room1, calendarId: envConfig.room1CalendarId },
      { key: 'room2', name: ROOM_NAMES.room2, calendarId: envConfig.room2CalendarId },
      { key: 'room3', name: ROOM_NAMES.room3, calendarId: envConfig.room3CalendarId }
    ];

    var calendarEvents = {};
    for (var r = 0; r < rooms.length; r++) {
      var room = rooms[r];
      var cal  = CalendarApp.getCalendarById(room.calendarId);
      if (!cal) {
        UtilityScriptLibrary.debugLog('getSchedulingData', 'WARNING', 'Calendar not found', room.name, '');
        calendarEvents[room.key] = [];
        continue;
      }
      var events    = cal.getEvents(today, endDate);
      var eventList = [];
      for (var ei = 0; ei < events.length; ei++) {
        var ev = events[ei];
        eventList.push({
          eventId: ev.getId(),
          title:   ev.getTitle(),
          start:   ev.getStartTime().toISOString(),
          end:     ev.getEndTime().toISOString()
        });
      }
      calendarEvents[room.key] = eventList;
    }

    UtilityScriptLibrary.debugLog('getSchedulingData', 'SUCCESS', 'Data loaded',
      teacherId + ' - ' + students.length + ' students', '');

    return {
      teacher:        teacherInfo,
      students:       students,
      calendarEvents: calendarEvents,
      rooms:          rooms.map(function(r) { return { key: r.key, name: r.name }; }),
      semesterEnd:    endDate.toISOString()
    };

  } catch (error) {
    UtilityScriptLibrary.debugLog('getSchedulingData', 'ERROR', 'Failed', teacherId, error.message);
    return { error: error.message };
  }
}

function getTeacherPreference(teacherId) {
  try {
    var sheet     = UtilityScriptLibrary.getSheet('teacherPreferences');
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var norm      = UtilityScriptLibrary.normalizeHeader;
    var idCol     = headerMap[norm('Teacher ID')];
    var roomCol   = headerMap[norm('Preferred Room')];

    if (!idCol || !roomCol) return null;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol - 1]).trim() === String(teacherId).trim()) {
        return String(data[i][roomCol - 1]).trim() || null;
      }
    }
    return null;

  } catch (error) {
    UtilityScriptLibrary.debugLog('getTeacherPreference', 'ERROR', 'Failed', teacherId, error.message);
    return null;
  }
}

function hasConflict(roomEvents, blockStart, blockEnd) {
  var BUFFER_MS = 30 * 60 * 1000;
  for (var i = 0; i < roomEvents.length; i++) {
    var ev      = roomEvents[i];
    var evStart = new Date(ev.start).getTime();
    var evEnd   = new Date(ev.end).getTime() + BUFFER_MS;
    var bStart  = blockStart.getTime();
    var bEnd    = blockEnd.getTime() + BUFFER_MS;
    if (bStart < evEnd && bEnd > evStart) {
      return true;
    }
  }
  return false;
}

function loadWeekEvents(roomCalendarIds, weekStart, weekEnd) {
  var endOfFriday = new Date(weekEnd.getTime());
  endOfFriday.setHours(23, 59, 59, 999);

  var weekEvents = {};
  for (var rk in roomCalendarIds) {
    var cal = CalendarApp.getCalendarById(roomCalendarIds[rk]);
    if (!cal) {
      weekEvents[rk] = [];
      continue;
    }
    var events    = cal.getEvents(weekStart, endOfFriday);
    var eventList = [];
    for (var ei = 0; ei < events.length; ei++) {
      var ev = events[ei];
      eventList.push({
        eventId: ev.getId(),
        title:   ev.getTitle(),
        start:   ev.getStartTime().toISOString(),
        end:     ev.getEndTime().toISOString()
      });
    }
    weekEvents[rk] = eventList;
  }
  return weekEvents;
}

function parseDateTime(dateStr, timeStr) {
  // dateStr: 'YYYY-MM-DD', timeStr: 'HH:MM'
  var parts     = dateStr.split('-');
  var timeParts = timeStr.split(':');
  return new Date(
    parseInt(parts[0]),
    parseInt(parts[1]) - 1,
    parseInt(parts[2]),
    parseInt(timeParts[0]),
    parseInt(timeParts[1]),
    0, 0
  );
}

function saveTeacherPreference(teacherId, roomKey) {
  try {
    var sheet     = UtilityScriptLibrary.getSheet('teacherPreferences');
    var headerMap = UtilityScriptLibrary.getHeaderMap(sheet);
    var norm      = UtilityScriptLibrary.normalizeHeader;
    var idCol     = headerMap[norm('Teacher ID')];
    var roomCol   = headerMap[norm('Preferred Room')];

    if (!idCol || !roomCol) return;

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol - 1]).trim() === String(teacherId).trim()) {
        // Already exists — only update if not already set
        if (!data[i][roomCol - 1]) {
          sheet.getRange(i + 1, roomCol).setValue(roomKey);
        }
        return;
      }
    }

    // Teacher not in sheet yet — this shouldn't happen since we pre-populate
    // but handle gracefully
    UtilityScriptLibrary.debugLog('saveTeacherPreference', 'WARNING', 
      'Teacher not found in preferences sheet', teacherId, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('saveTeacherPreference', 'ERROR', 'Failed', teacherId, error.message);
  }
}

function sendMaintenanceSchedule() {
  try {
    UtilityScriptLibrary.debugLog('sendMaintenanceSchedule', 'INFO', 'Starting maintenance schedule', '', '');

    var weekDates = getComingWeekDates();
    var weekStart = weekDates[0];
    var weekEnd   = weekDates[4];
    var weekLabel = formatDateForEmail(weekStart);

    var env       = UtilityScriptLibrary.EnvironmentManager.get();
    var envConfig = UtilityScriptLibrary.getConfig()[env];

    var roomCalendarIds = {
      room1: envConfig.room1CalendarId,
      room2: envConfig.room2CalendarId,
      room3: envConfig.room3CalendarId
    };

    var weekEvents = loadWeekEvents(roomCalendarIds, weekStart, weekEnd);

    // Build summary per room
    var body = 'QAMP Room Usage - Week of ' + weekLabel + '\n\n';

    for (var rk in ROOM_NAMES) {
      body += ROOM_NAMES[rk] + ':\n';
      var roomHasEvents = false;

      for (var d = 0; d < weekDates.length; d++) {
        var dayDate   = weekDates[d];
        var dayEvents = weekEvents[rk].filter(function(ev) {
          var evDate = new Date(ev.start);
          return evDate.toDateString() === dayDate.toDateString();
        });

        if (!dayEvents.length) continue;

        roomHasEvents = true;
        dayEvents.sort(function(a, b) { return new Date(a.start) - new Date(b.start); });

        var firstStart = new Date(dayEvents[0].start);
        var lastEnd    = new Date(dayEvents[dayEvents.length - 1].end);

        body += '  ' + DAY_NAMES[d] + ' ' + formatDateForEmail(dayDate) + ': ' +
          formatTimeForEmail(firstStart) + ' - ' + formatTimeForEmail(lastEnd) + '\n';
      }

      if (!roomHasEvents) {
        body += '  Not in use this week\n';
      }

      body += '\n';
    }

    // Get maintenance email from config
    var maintenanceEmail = envConfig.maintenanceEmail;
    if (!maintenanceEmail) {
      throw new Error('Maintenance email not found in config');
    }

    UtilityScriptLibrary.sendEmail(maintenanceEmail, 'QAMP Room Schedule - Week of ' + weekLabel, body);

    UtilityScriptLibrary.debugLog('sendMaintenanceSchedule', 'SUCCESS', 'Maintenance schedule sent', maintenanceEmail, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('sendMaintenanceSchedule', 'ERROR', 'Failed', '', error.message);
  }
}

function sendTeacherConfirmations() {
  try {
    UtilityScriptLibrary.debugLog('sendTeacherConfirmations', 'INFO', 'Starting weekly confirmations', '', '');

    var weekDates  = getComingWeekDates(); // Monday-Friday
    var weekStart  = weekDates[0];
    var weekEnd    = weekDates[4];
    var weekLabel  = formatDateForEmail(weekStart);

    var env        = UtilityScriptLibrary.EnvironmentManager.get();
    var envConfig  = UtilityScriptLibrary.getConfig()[env];

    var roomCalendarIds = {
      room1: envConfig.room1CalendarId,
      room2: envConfig.room2CalendarId,
      room3: envConfig.room3CalendarId
    };

    // Load all calendar events for the week once
    var weekEvents = loadWeekEvents(roomCalendarIds, weekStart, weekEnd);

    // Get all active teachers
    var lookupSheet = UtilityScriptLibrary.getSheet('teacherRosterLookup');
    var lookupData  = lookupSheet.getDataRange().getValues();
    var headerMap   = UtilityScriptLibrary.getHeaderMap(lookupSheet);
    var norm        = UtilityScriptLibrary.normalizeHeader;

    var tidCol      = headerMap[norm('Teacher ID')];
    var statusCol   = headerMap[norm('Status')];

    if (!tidCol || !statusCol) {
      throw new Error('Required columns not found in Teacher Roster Lookup');
    }

    var contactsSheet = UtilityScriptLibrary.getSheet('teachersAndAdmin');
    var contactsData  = contactsSheet.getDataRange().getValues();
    var contactsMap   = UtilityScriptLibrary.getHeaderMap(contactsSheet);

    var cIdCol        = contactsMap[norm('Teacher ID')];
    var cFirstCol     = contactsMap[norm('First Name')];
    var cLastCol      = contactsMap[norm('Last Name')];
    var cEmailCol     = contactsMap[norm('Email')];

    var sent   = 0;
    var errors = [];

    for (var i = 1; i < lookupData.length; i++) {
      var row    = lookupData[i];
      var status = statusCol ? String(row[statusCol - 1]).trim().toLowerCase() : '';
      if (status !== 'active') continue;

      var teacherId = String(row[tidCol - 1]).trim();
      if (!teacherId) continue;

      // Get teacher contact info
      var teacherContact = null;
      for (var c = 1; c < contactsData.length; c++) {
        if (String(contactsData[c][cIdCol - 1]).trim() === teacherId) {
          teacherContact = {
            firstName: String(contactsData[c][cFirstCol - 1] || '').trim(),
            lastName:  String(contactsData[c][cLastCol  - 1] || '').trim(),
            email:     String(contactsData[c][cEmailCol - 1] || '').trim()
          };
          break;
        }
      }

      if (!teacherContact || !teacherContact.email) {
        errors.push('No contact info for teacher: ' + teacherId);
        continue;
      }

      // Filter events for this teacher across all rooms
      var teacherSchedule = buildTeacherWeekSchedule(weekEvents, teacherContact.lastName, weekDates);

      // Build email body
      var schedulingUrl = envConfig.schedulingWebAppUrl + '?tid=' + teacherId;
      var body = 'Hi ' + teacherContact.firstName + ',\n\n' +
        'Here is your lesson schedule for the week of ' + weekLabel + ':\n\n' +
        teacherSchedule + '\n\n' +
        'If you need to make any changes, please visit your scheduling page:\n' +
        schedulingUrl + '\n\n' +
        'Thank you,\nQAMP';

      var subject = 'QAMP Schedule - Week of ' + weekLabel;

      try {
        UtilityScriptLibrary.sendEmail(teacherContact.email, subject, body);
        sent++;
        UtilityScriptLibrary.debugLog('sendTeacherConfirmations', 'INFO', 'Confirmation sent',
          teacherId + ' - ' + teacherContact.email, '');
      } catch (emailError) {
        errors.push('Failed to send to ' + teacherContact.email + ': ' + emailError.message);
      }
    }

    UtilityScriptLibrary.debugLog('sendTeacherConfirmations', 'SUCCESS',
      'Confirmations complete', 'Sent: ' + sent + ', Errors: ' + errors.length, '');

  } catch (error) {
    UtilityScriptLibrary.debugLog('sendTeacherConfirmations', 'ERROR', 'Failed', '', error.message);
  }
}

function submitSchedule(payload) {
  try {
    UtilityScriptLibrary.debugLog('submitSchedule', 'INFO', 'Processing submission', payload.teacherId, '');

    var teacherId = payload.teacherId;
    var days      = payload.days;

    if (!teacherId || !days || !days.length) {
      throw new Error('Invalid payload — missing teacherId or days');
    }

    // Load calendar events once for conflict checking
    var schedulingData = getSchedulingData(teacherId);
    if (schedulingData.error) {
      throw new Error(schedulingData.error);
    }

    var calendarEvents = schedulingData.calendarEvents;
    var teacher        = schedulingData.teacher;
    var env            = UtilityScriptLibrary.EnvironmentManager.get();
    var envConfig      = UtilityScriptLibrary.getConfig()[env];

    var roomCalendarIds = {
      room1: envConfig.room1CalendarId,
      room2: envConfig.room2CalendarId,
      room3: envConfig.room3CalendarId
    };

    // Open teacher roster once
    var rosterSS = SpreadsheetApp.openByUrl(schedulingData.teacher.rosterUrl);

    var results = [];
    var errors  = [];

    for (var d = 0; d < days.length; d++) {
      var day     = days[d];
      var date    = new Date(day.date);
      var lessons = day.lessons;

      if (!lessons || !lessons.length) continue;

      // Build student objects for room assignment
      var dayStudents = [];
      for (var l = 0; l < lessons.length; l++) {
        var lesson = lessons[l];
        // Find student in schedulingData.students
        var studentInfo = null;
        for (var s = 0; s < schedulingData.students.length; s++) {
          if (String(schedulingData.students[s].id).trim() === String(lesson.studentId).trim()) {
            studentInfo = schedulingData.students[s];
            break;
          }
        }
        if (!studentInfo) {
          errors.push('Student not found: ' + lesson.studentId);
          continue;
        }
        dayStudents.push(studentInfo);
      }

      if (!dayStudents.length) continue;

      // Determine earliest start and latest end for the day to assign room
      var dayStart = null;
      var dayEnd   = null;
      for (var l2 = 0; l2 < lessons.length; l2++) {
        var ls        = lessons[l2];
        var lStart    = parseDateTime(day.date, ls.startTime);
        var lEnd      = new Date(lStart.getTime() + ls.lessonLength * 60 * 1000);
        if (!dayStart || lStart < dayStart) dayStart = lStart;
        if (!dayEnd   || lEnd   > dayEnd)   dayEnd   = lEnd;
      }

      // Assign room for the day
      var roomKey;
      try {
        roomKey = assignRoom(teacherId, dayStudents, dayStart, dayEnd, calendarEvents);
      } catch (roomError) {
        errors.push('Day ' + day.date + ': ' + roomError.message);
        continue;
      }

      var cal      = CalendarApp.getCalendarById(roomCalendarIds[roomKey]);
      var roomName = ROOM_NAMES[roomKey] ;

      // Process each lesson
      for (var l3 = 0; l3 < lessons.length; l3++) {
        var lesson3    = lessons[l3];
        var lessonStart = parseDateTime(day.date, lesson3.startTime);
        var lessonEnd   = new Date(lessonStart.getTime() + lesson3.lessonLength * 60 * 1000);

        // Find student info
        var student3 = null;
        for (var s2 = 0; s2 < schedulingData.students.length; s2++) {
          if (String(schedulingData.students[s2].id).trim() === String(lesson3.studentId).trim()) {
            student3 = schedulingData.students[s2];
            break;
          }
        }
        if (!student3) continue;

        try {
          // Create calendar event
          var eventTitle = teacher.lastName + ' - ' + student3.instrument + ' - ' + lesson3.lessonLength + ' min';
          var event      = cal.createEvent(eventTitle, lessonStart, lessonEnd, { location: roomName });
          var eventId    = event.getId();

          // Write date + event ID to lesson log
          var monthName  = UtilityScriptLibrary.getMonthNameFromDate(date, true);
          var monthSheet = rosterSS.getSheetByName(monthName);

          if (monthSheet) {
            var targetRow = findNextBlankLessonRow(monthSheet, lesson3.studentId);
            if (targetRow) {
              var headerMap  = UtilityScriptLibrary.getHeaderMap(monthSheet);
              var norm       = UtilityScriptLibrary.normalizeHeader;
              var dateCol    = headerMap[norm('Date')];
              var eventIdCol = headerMap[norm('Calendar Event ID')];
              if (dateCol)    monthSheet.getRange(targetRow, dateCol).setValue(date);
              if (eventIdCol) monthSheet.getRange(targetRow, eventIdCol).setValue(eventId);
            } else {
              errors.push('No blank lesson row for student ' + lesson3.studentId + ' on ' + day.date);
            }
          } else {
            errors.push('No roster sheet found for month ' + monthName);
          }

          // Add event to local calendarEvents so same-submission conflicts are caught
          calendarEvents[roomKey].push({
            eventId: eventId,
            title:   eventTitle,
            start:   lessonStart.toISOString(),
            end:     lessonEnd.toISOString()
          });

          results.push({
            studentId: lesson3.studentId,
            date:      day.date,
            startTime: lesson3.startTime,
            room:      roomName,
            eventId:   eventId
          });

          UtilityScriptLibrary.debugLog('submitSchedule', 'INFO', 'Lesson scheduled',
            lesson3.studentId + ' - ' + day.date + ' - ' + roomName, '');

        } catch (lessonError) {
          errors.push('Lesson ' + lesson3.studentId + ' on ' + day.date + ': ' + lessonError.message);
          UtilityScriptLibrary.debugLog('submitSchedule', 'ERROR', 'Lesson failed',
            lesson3.studentId + ' - ' + day.date, lessonError.message);
        }
      }
    }

    UtilityScriptLibrary.debugLog('submitSchedule', 'SUCCESS', 'Submission complete',
      teacherId + ' - ' + results.length + ' scheduled, ' + errors.length + ' errors', '');

    return { success: true, results: results, errors: errors };

  } catch (error) {
    UtilityScriptLibrary.debugLog('submitSchedule', 'ERROR', 'Failed', payload.teacherId, error.message);
    return { success: false, error: error.message };
  }
}

function updateLesson(payload) {
  try {
    UtilityScriptLibrary.debugLog('updateLesson', 'INFO', 'Processing update', payload.teacherId, '');

    var teacherId  = payload.teacherId;
    var eventId    = payload.eventId;
    var action     = payload.action; // 'reschedule', 'cancel', 'update'
    var studentId  = payload.studentId;

    if (!teacherId || !eventId || !action || !studentId) {
      throw new Error('Invalid payload — missing required fields');
    }

    var env            = UtilityScriptLibrary.EnvironmentManager.get();
    var envConfig      = UtilityScriptLibrary.getConfig()[env];

    var roomCalendarIds = {
      room1: envConfig.room1CalendarId,
      room2: envConfig.room2CalendarId,
      room3: envConfig.room3CalendarId
    };

    // Find which calendar contains this event
    var targetCal   = null;
    var targetEvent = null;
    var roomKey     = null;

    for (var key in roomCalendarIds) {
      var cal = CalendarApp.getCalendarById(roomCalendarIds[key]);
      if (!cal) continue;
      try {
        var ev = cal.getEventById(eventId);
        if (ev) {
          targetCal   = cal;
          targetEvent = ev;
          roomKey     = key;
          break;
        }
      } catch (calError) {
        continue;
      }
    }

    if (!targetEvent) {
      throw new Error('Event not found in any room calendar: ' + eventId);
    }

    // Find the lesson log row with this event ID
    var schedulingData = getSchedulingData(teacherId);
    if (schedulingData.error) throw new Error(schedulingData.error);

    var rosterSS  = SpreadsheetApp.openByUrl(schedulingData.teacher.rosterUrl);
    var logRow    = findLessonLogRowByEventId(rosterSS, eventId);

    if (action === 'cancel') {
      // Delete calendar event
      targetEvent.deleteEvent();

      // Clear date and event ID from lesson log
      if (logRow) {
        var norm       = UtilityScriptLibrary.normalizeHeader;
        var headerMap  = UtilityScriptLibrary.getHeaderMap(logRow.sheet);
        var dateCol    = headerMap[norm('Date')];
        var eventIdCol = headerMap[norm('Calendar Event ID')];
        if (dateCol)    logRow.sheet.getRange(logRow.row, dateCol).clearContent();
        if (eventIdCol) logRow.sheet.getRange(logRow.row, eventIdCol).clearContent();
      }

      UtilityScriptLibrary.debugLog('updateLesson', 'SUCCESS', 'Lesson cancelled', eventId, '');
      return { success: true, action: 'cancel', eventId: eventId };

    } else if (action === 'reschedule') {
      // Validate new date/time
      if (!payload.date || !payload.startTime || !payload.lessonLength) {
        throw new Error('Reschedule requires date, startTime, and lessonLength');
      }

      var newStart  = parseDateTime(payload.date, payload.startTime);
      var newEnd    = new Date(newStart.getTime() + payload.lessonLength * 60 * 1000);

      // Check conflicts in the same room, excluding this event
      var calendarEvents = schedulingData.calendarEvents;
      var roomEvents     = calendarEvents[roomKey].filter(function(ev) {
        return ev.eventId !== eventId;
      });

      if (hasConflict(roomEvents, newStart, newEnd)) {
        // Try to find another available room
        var student = null;
        for (var si = 0; si < schedulingData.students.length; si++) {
          if (String(schedulingData.students[si].id).trim() === String(studentId).trim()) {
            student = schedulingData.students[si];
            break;
          }
        }

        var allEventsExcluding = {};
        for (var rk in calendarEvents) {
          allEventsExcluding[rk] = calendarEvents[rk].filter(function(ev) {
            return ev.eventId !== eventId;
          });
        }

        var newRoomKey;
        try {
          newRoomKey = assignRoom(teacherId, [student], newStart, newEnd, allEventsExcluding);
        } catch (roomError) {
          throw new Error('No rooms available for rescheduled time: ' + roomError.message);
        }

        // Capture title before deleting event
        var existingTitle = targetEvent.getTitle();
        var titleParts    = existingTitle.split(' - ');
        var newEventTitle = (titleParts.length === 3 && payload.lessonLength)
          ? titleParts[0] + ' - ' + titleParts[1] + ' - ' + payload.lessonLength + ' min'
          : existingTitle;

        // Move to new room
        targetEvent.deleteEvent();
        var newCal      = CalendarApp.getCalendarById(roomCalendarIds[newRoomKey]);
        var newRoomName = ROOM_NAMES[newRoomKey];
        var newEvent    = newCal.createEvent(newEventTitle, newStart, newEnd, { location: newRoomName });
        var newEventId = newEvent.getId();

        // Update lesson log
        if (logRow) {
          var norm2       = UtilityScriptLibrary.normalizeHeader;
          var headerMap2  = UtilityScriptLibrary.getHeaderMap(logRow.sheet);
          var dateCol2    = headerMap2[norm2('Date')];
          var eventIdCol2 = headerMap2[norm2('Calendar Event ID')];
          if (dateCol2)    logRow.sheet.getRange(logRow.row, dateCol2).setValue(new Date(payload.date));
          if (eventIdCol2) logRow.sheet.getRange(logRow.row, eventIdCol2).setValue(newEventId);
        }

        UtilityScriptLibrary.debugLog('updateLesson', 'SUCCESS', 'Lesson rescheduled to new room',
          eventId + ' -> ' + newEventId + ' in ' + newRoomKey, '');
        return { success: true, action: 'reschedule', eventId: newEventId, room: newRoomName, roomChanged: true };

      } else {
        // Same room, just update time
        targetEvent.setTime(newStart, newEnd);

        // Update title if lesson length changed
        if (payload.lessonLength) {
          var oldTitle   = targetEvent.getTitle();
          var titleParts = oldTitle.split(' - ');
          if (titleParts.length === 3) {
            targetEvent.setTitle(titleParts[0] + ' - ' + titleParts[1] + ' - ' + payload.lessonLength + ' min');
          }
        }

        // Update lesson log date
        if (logRow) {
          var norm3       = UtilityScriptLibrary.normalizeHeader;
          var headerMap3  = UtilityScriptLibrary.getHeaderMap(logRow.sheet);
          var dateCol3    = headerMap3[norm3('Date')];
          var eventIdCol3 = headerMap3[norm3('Calendar Event ID')];
          if (dateCol3)    logRow.sheet.getRange(logRow.row, dateCol3).setValue(new Date(payload.date));
          if (eventIdCol3) logRow.sheet.getRange(logRow.row, eventIdCol3).setValue(eventId);
        }

        UtilityScriptLibrary.debugLog('updateLesson', 'SUCCESS', 'Lesson rescheduled same room', eventId, '');
        return { success: true, action: 'reschedule', eventId: eventId, room: ROOM_NAMES[roomKey], roomChanged: false };
      }

    } else if (action === 'update') {
      // Length change only — same date/time
      if (!payload.lessonLength) {
        throw new Error('Update requires lessonLength');
      }

      var currentStart = targetEvent.getStartTime();
      var updatedEnd   = new Date(currentStart.getTime() + payload.lessonLength * 60 * 1000);

      // Check conflicts with new end time
      var roomEventsExcluding = schedulingData.calendarEvents[roomKey].filter(function(ev) {
        return ev.eventId !== eventId;
      });

      if (hasConflict(roomEventsExcluding, currentStart, updatedEnd)) {
        throw new Error('Updated lesson length causes a conflict in ' + ROOM_NAMES[roomKey]);
      }

      targetEvent.setTime(currentStart, updatedEnd);

      // Update title
      var oldTitle2   = targetEvent.getTitle();
      var titleParts2 = oldTitle2.split(' - ');
      if (titleParts2.length === 3) {
        targetEvent.setTitle(titleParts2[0] + ' - ' + titleParts2[1] + ' - ' + payload.lessonLength + ' min');
      }

      // Update lesson log length
      if (logRow) {
        var norm4      = UtilityScriptLibrary.normalizeHeader;
        var headerMap4 = UtilityScriptLibrary.getHeaderMap(logRow.sheet);
        var lengthCol  = headerMap4[norm4('Length')];
        if (lengthCol) logRow.sheet.getRange(logRow.row, lengthCol).setValue(payload.lessonLength);
      }

      UtilityScriptLibrary.debugLog('updateLesson', 'SUCCESS', 'Lesson length updated', eventId, '');
      return { success: true, action: 'update', eventId: eventId };

    } else {
      throw new Error('Unknown action: ' + action);
    }

  } catch (error) {
    UtilityScriptLibrary.debugLog('updateLesson', 'ERROR', 'Failed', payload.teacherId, error.message);
    return { success: false, error: error.message };
  }
}
