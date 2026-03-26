// ============================================================
// EVC Class Roster Generator — Add-on Module (v2)
// ============================================================
// HR Program Coordinator: Kyle Mahoney
// Emory Valley Center
// ============================================================
// PASTE THIS INTO: Extensions > Apps Script as a NEW file
//   (click the + next to Files, name it "ClassRosters")
//   Do NOT replace Code.gs — this runs alongside it.
//
// ALSO: Replace writeRosterSheet() in Code.gs with the
//   updated version at the bottom of this file. The new
//   version scans class roster tabs and marks scheduled
//   people on the Training Rosters output.
//
// After pasting, reload the spreadsheet. Add this line to
// your createMenu() in Code.gs:
//   .addItem("Generate Class Rosters", "generateClassRosters")
//
// ============================================================
// HOW IT WORKS:
//
// 1. "Generate Class Rosters" pulls from the same needs data
//    as Training Rosters, assigns people to upcoming class
//    dates by priority (Expired > Never Completed > Expiring
//    Soon), and creates one tab per class date.
//
// 2. When Training Rosters refreshes (auto or manual), it
//    scans all existing class roster tabs (any tab whose name
//    starts with a known training name). If a person on the
//    needs list appears on a class roster tab, their entry
//    shows "Scheduled (M/D/YYYY)" instead of the normal
//    priority label.
//
// 3. Scheduled people are NOT removed — they stay visible
//    so you see the full picture, but you know they're
//    already handled.
// ============================================================

// ============================================================
// CLASS ROSTER CONFIGURATION
// ============================================================
//
//   name:          Display name (must match TRAINING_CONFIG "name")
//   classCapacity: Default max seats per class (overridable)
//   schedule:      How upcoming class dates are generated
//     recurring:   Array of recurrence rules. Each rule has:
//       weekday:   Day of week (e.g. "Thursday")
//       nthWeek:   (optional) Array of week-of-month [1-5]
//                  e.g. [2,4] = 2nd and 4th week only
//     dates:       (optional) Array of specific "M/D/YYYY" strings
//                  added on top of any recurring rules
//   weeksOut:      How many weeks ahead to generate (default 4)
// ============================================================

var CLASS_ROSTER_CONFIG = [
  {
    name: "CPR/FA",
    classCapacity: 10,
    schedule: {
      recurring: [
        { weekday: "Thursday" }          // Every Thursday
      ]
    },
    weeksOut: 4
  },
  {
    name: "Ukeru",
    classCapacity: 12,
    schedule: {
      recurring: [
        { weekday: "Monday", nthWeek: [2] },  // Monday of 2nd week
        { weekday: "Friday", nthWeek: [4] }    // Friday of 4th week
      ],
      dates: []                                // Add one-off overrides here
    },
    weeksOut: 6
  },
  {
    name: "Mealtime",
    classCapacity: 15,
    schedule: {
      recurring: [
        { weekday: "Wednesday", nthWeek: [3] } // 3rd Wednesday monthly
      ]
    },
    weeksOut: 8
  },
  {
    name: "Med Recert",
    classCapacity: 4,
    schedule: { dates: [] },
    weeksOut: 4
  },
  {
    name: "Post Med",
    classCapacity: 8,
    schedule: { dates: [] },
    weeksOut: 4
  },
  {
    name: "POMs",
    classCapacity: 15,
    schedule: { dates: [] },
    weeksOut: 4
  },
  {
    name: "Person Centered",
    classCapacity: 15,
    schedule: { dates: [] },
    weeksOut: 4
  },
  {
    name: "Van/Lift Training",
    classCapacity: 10,
    schedule: { dates: [] },
    weeksOut: 4
  }
];

// ============================================================
// PRIORITY ORDER for filling seats
// 1 = first pick. Change numbers to reorder.
// ============================================================
var SEAT_PRIORITY = [
  { bucket: "expired",      priority: 1 },
  { bucket: "needed",       priority: 2 },
  { bucket: "expiringSoon", priority: 3 }
];
SEAT_PRIORITY.sort(function(a, b) { return a.priority - b.priority; });

// ============================================================
// All known training name prefixes used for class roster tabs.
// Used by scanClassRosterTabs() to identify which tabs are
// class rosters. Built automatically from CLASS_ROSTER_CONFIG.
// ============================================================
function getClassRosterPrefixes() {
  var prefixes = [];
  for (var i = 0; i < CLASS_ROSTER_CONFIG.length; i++) {
    prefixes.push(CLASS_ROSTER_CONFIG[i].name);
  }
  return prefixes;
}


// ************************************************************
//
//   CLASS ROSTER GENERATION
//
// ************************************************************

// ============================================================
// MAIN ENTRY: Generate Class Rosters
// ============================================================
function generateClassRosters() {
  var ui = SpreadsheetApp.getUi();

  // Step 1: Pick training type(s)
  var msg = "Which training do you want to generate class rosters for?\n\n";
  msg += "0.  ALL trainings (generate everything)\n";
  for (var i = 0; i < CLASS_ROSTER_CONFIG.length; i++) {
    msg += (i + 1) + ".  " + CLASS_ROSTER_CONFIG[i].name +
           " (capacity: " + CLASS_ROSTER_CONFIG[i].classCapacity + ")\n";
  }
  msg += "\nEnter a number, or comma-separated numbers (e.g. 1,3,5):";

  var choice = ui.prompt("Generate Class Rosters", msg, ui.ButtonSet.OK_CANCEL);
  if (choice.getSelectedButton() !== ui.Button.OK) return;

  var input = choice.getResponseText().trim();
  var selectedConfigs = [];

  if (input === "0") {
    selectedConfigs = CLASS_ROSTER_CONFIG.slice();
  } else {
    var parts = input.split(",");
    for (var p = 0; p < parts.length; p++) {
      var idx = parseInt(parts[p].trim()) - 1;
      if (!isNaN(idx) && idx >= 0 && idx < CLASS_ROSTER_CONFIG.length) {
        selectedConfigs.push(CLASS_ROSTER_CONFIG[idx]);
      }
    }
  }

  if (selectedConfigs.length === 0) {
    ui.alert("No valid selection. Please try again.");
    return;
  }

  // Step 2: Build needs list
  var rosterResult = buildRosterData(true);
  if (!rosterResult) {
    ui.alert("Could not read Training sheet data. Check that your Training sheet exists.");
    return;
  }

  var ss = rosterResult.ss;
  var today = rosterResult.today;
  var allRosters = rosterResult.allRosters;

  // Scan existing class roster tabs so we don't double-assign
  var alreadyScheduled = scanClassRosterTabs(ss);

  var tabsCreated = 0;
  var summaryLines = [];

  // Step 3: Process each selected training
  for (var s = 0; s < selectedConfigs.length; s++) {
    var config = selectedConfigs[s];
    var lookupName = config.name;

    // Find matching roster data
    var rosterData = null;
    for (var r = 0; r < allRosters.length; r++) {
      if (allRosters[r].name === lookupName) {
        rosterData = allRosters[r];
        break;
      }
    }

    if (!rosterData || rosterData.error) {
      summaryLines.push(config.name + ": Skipped (no roster data or column error)");
      continue;
    }

    // Build prioritized pool, excluding already-scheduled people
    var pool = buildPriorityPool(rosterData, alreadyScheduled, lookupName);

    if (pool.length === 0) {
      summaryLines.push(config.name + ": Everyone is current or already scheduled!");
      continue;
    }

    // Generate upcoming dates
    var dates = generateUpcomingDates(config, today);

    // Let user review/edit dates and capacity
    var dateStr = "";
    for (var d = 0; d < dates.length; d++) {
      dateStr += formatClassDate(dates[d]) + "\n";
    }
    if (!dateStr) dateStr = "(none generated from config)\n";

    var datePrompt = ui.prompt(
      config.name + " — Class Dates",
      "Upcoming dates for " + config.name + ":\n\n" +
      dateStr +
      "\nPeople needing this training: " + pool.length +
      " (excludes already-scheduled)" +
      "\nDefault capacity: " + config.classCapacity + " per class" +
      "\n\nEdit the dates below (one per line, M/D/YYYY).\n" +
      "Add or remove lines as needed.\n" +
      "Leave blank to skip this training.",
      ui.ButtonSet.OK_CANCEL
    );
    if (datePrompt.getSelectedButton() !== ui.Button.OK) continue;

    var dateInput = datePrompt.getResponseText().trim();
    if (!dateInput) {
      summaryLines.push(config.name + ": Skipped (no dates entered)");
      continue;
    }

    var finalDates = parseDateList(dateInput);
    if (finalDates.length === 0) {
      summaryLines.push(config.name + ": Skipped (no valid dates parsed)");
      continue;
    }

    // Capacity override
    var capPrompt = ui.prompt(
      config.name + " — Class Capacity",
      "Max seats per class for " + config.name + "?\n\n" +
      "Default: " + config.classCapacity + "\n" +
      "People in pool: " + pool.length + "\n" +
      "Classes: " + finalDates.length + "\n" +
      "Total seats: " + (config.classCapacity * finalDates.length) +
      "\n\nPress OK to use default, or enter a new number:",
      ui.ButtonSet.OK_CANCEL
    );
    if (capPrompt.getSelectedButton() !== ui.Button.OK) continue;

    var capInput = capPrompt.getResponseText().trim();
    var capacity = config.classCapacity;
    if (capInput && !isNaN(parseInt(capInput))) {
      capacity = parseInt(capInput);
    }

    // Assign people to classes
    var assignments = assignToClasses(pool, finalDates, capacity);

    // Create tabs
    for (var a = 0; a < assignments.length; a++) {
      var classInfo = assignments[a];
      var tabName = buildTabName(config.name, classInfo.date);

      var existingTab = ss.getSheetByName(tabName);
      if (existingTab) ss.deleteSheet(existingTab);

      var tab = ss.insertSheet(tabName);
      writeClassRosterTab(tab, config.name, classInfo, capacity, today);
      tabsCreated++;

      // Track these assignments so if the same training appears
      // again later in the run it won't double-book anyone
      for (var ap = 0; ap < classInfo.people.length; ap++) {
        var personName = classInfo.people[ap].name.toLowerCase().trim();
        var trainKey = lookupName.toLowerCase();
        if (!alreadyScheduled[trainKey]) alreadyScheduled[trainKey] = {};
        alreadyScheduled[trainKey][personName] = formatClassDate(classInfo.date);
      }
    }

    var totalAssigned = 0;
    for (var aa = 0; aa < assignments.length; aa++) {
      totalAssigned += assignments[aa].people.length;
    }
    var leftover = pool.length - totalAssigned;

    var line = config.name + ": " + totalAssigned + " assigned across " +
               finalDates.length + " class(es)";
    if (leftover > 0) line += " | " + leftover + " still unassigned (need more classes)";
    summaryLines.push(line);
  }

  // Refresh Training Rosters so it picks up scheduled people
  generateRostersSilent();

  // Sort class roster tabs: grouped by training type, then by date
  orderClassRosterTabs(ss);

  // Summary
  var summary = "Class Roster Generation Complete!\n\n";
  summary += "Tabs created: " + tabsCreated + "\n\n";
  for (var sl = 0; sl < summaryLines.length; sl++) {
    summary += summaryLines[sl] + "\n";
  }
  summary += "\nTraining Rosters tab refreshed to show scheduled staff.";
  summary += "\nClass roster tabs sorted by training type and date.";

  ui.alert(summary);
}


// ************************************************************
//
//   SCANNING CLASS ROSTER TABS
//
// ************************************************************

// ============================================================
// scanClassRosterTabs — reads all existing class roster tabs
// Returns: { "training_lower": { "name_lower": "M/D/YYYY" } }
// ============================================================
function scanClassRosterTabs(ss) {
  var scheduled = {};
  var prefixes = getClassRosterPrefixes();
  var sheets = ss.getSheets();

  for (var s = 0; s < sheets.length; s++) {
    var sheetName = sheets[s].getName();
    var matchedTraining = null;

    // Match longest prefix first for accuracy
    for (var p = 0; p < prefixes.length; p++) {
      if (sheetName.indexOf(prefixes[p] + " ") === 0) {
        if (!matchedTraining || prefixes[p].length > matchedTraining.length) {
          matchedTraining = prefixes[p];
        }
      }
    }

    if (!matchedTraining) continue;

    // Extract date from tab name
    var datePart = sheetName.substring(matchedTraining.length + 1).trim();
    var classDate = parseClassDate(datePart);
    var classDateStr = classDate ? formatClassDate(classDate) : datePart;

    // Use matched training name directly as the key
    var trainKey = matchedTraining.toLowerCase();

    if (!scheduled[trainKey]) scheduled[trainKey] = {};

    // Read tab data. Names are in column 2 (index 1), starting row 8
    // (rows 1-6 = header info, row 7 = column headers)
    var data = sheets[s].getDataRange().getValues();
    for (var r = 7; r < data.length; r++) {
      var nameVal = data[r][1] ? data[r][1].toString().trim() : "";
      if (!nameVal) continue;
      if (nameVal.toLowerCase().indexOf("open seat") > -1) continue;

      var nameLower = nameVal.toLowerCase();
      // Keep earliest scheduled date if on multiple rosters
      if (!scheduled[trainKey][nameLower]) {
        scheduled[trainKey][nameLower] = classDateStr;
      }
    }
  }

  return scheduled;
}


// ************************************************************
//
//   POOL BUILDING & ASSIGNMENT
//
// ************************************************************

// ============================================================
// Build prioritized pool, excluding already-scheduled people
// ============================================================
function buildPriorityPool(rosterData, alreadyScheduled, trainingName) {
  var pool = [];
  var trainKey = trainingName.toLowerCase();
  var scheduledMap = alreadyScheduled[trainKey] || {};

  for (var p = 0; p < SEAT_PRIORITY.length; p++) {
    var bucketName = SEAT_PRIORITY[p].bucket;
    var bucketData = rosterData[bucketName] || [];

    for (var i = 0; i < bucketData.length; i++) {
      var personNameLower = bucketData[i].name.toLowerCase().trim();
      if (scheduledMap[personNameLower]) continue;

      pool.push({
        name: bucketData[i].name,
        status: bucketData[i].status,
        bucket: bucketName,
        lastDate: bucketData[i].lastDate || "",
        expDate: bucketData[i].expDate || ""
      });
    }
  }

  return pool;
}

// ============================================================
// Assign people to classes with per-date priority awareness
//
// For each class date, anyone whose expiration falls BEFORE
// that class date gets treated as effectively expired for
// that class, regardless of their current bucket. This means
// an "Expiring Soon" person whose cert runs out before the
// next class gets bumped to top priority for that class.
//
// Effective priority per class date:
//   1. Already expired (bucket "expired")
//   2. Will be expired by this class date (bucket "expiringSoon"
//      but expDate < classDate)
//   3. Never completed (bucket "needed")
//   4. Expiring soon but NOT before this class date
// ============================================================
function assignToClasses(pool, dates, capacity) {
  var assignments = [];
  var assigned = {}; // track who's been assigned (by name lowercase)

  for (var d = 0; d < dates.length; d++) {
    var classDate = dates[d];
    var classGroup = {
      date: classDate,
      people: []
    };

    // Build a scored list of remaining unassigned people
    var candidates = [];
    for (var p = 0; p < pool.length; p++) {
      var person = pool[p];
      if (assigned[person.name.toLowerCase()]) continue;

      var effectivePriority = getEffectivePriority(person, classDate);
      candidates.push({
        person: person,
        effectivePriority: effectivePriority,
        originalIndex: p
      });
    }

    // Sort by effective priority (lowest number = first pick)
    candidates.sort(function(a, b) {
      if (a.effectivePriority !== b.effectivePriority) {
        return a.effectivePriority - b.effectivePriority;
      }
      // Tie-break: keep original pool order
      return a.originalIndex - b.originalIndex;
    });

    // Fill seats up to capacity
    for (var c = 0; c < candidates.length && classGroup.people.length < capacity; c++) {
      var candidate = candidates[c];
      var personCopy = {
        name: candidate.person.name,
        status: candidate.person.status,
        bucket: candidate.person.bucket,
        lastDate: candidate.person.lastDate,
        expDate: candidate.person.expDate,
        effectiveBucket: candidate.effectivePriority <= 2 ? "expired" : candidate.person.bucket
      };

      // Update status text if they'll be expired by class date
      // but aren't currently in the expired bucket
      if (candidate.effectivePriority === 2 && candidate.person.bucket === "expiringSoon") {
        personCopy.status = "Will expire before class (" + candidate.person.expDate + ")";
        personCopy.effectiveBucket = "expired";
      }

      classGroup.people.push(personCopy);
      assigned[candidate.person.name.toLowerCase()] = true;
    }

    assignments.push(classGroup);
  }

  return assignments;
}

// ============================================================
// getEffectivePriority — determines a person's priority for
// a specific class date
//
// Returns:
//   1 = Already expired
//   2 = Will be expired by this class date
//   3 = Never completed
//   4 = Expiring soon but still valid on class date
// ============================================================
function getEffectivePriority(person, classDate) {
  // Already expired
  if (person.bucket === "expired") return 1;

  // Expiring soon — check if they'll actually be expired by class date
  if (person.bucket === "expiringSoon" && person.expDate) {
    var exp = parseClassDate(person.expDate);
    if (exp && exp.getTime() < classDate.getTime()) {
      return 2; // Will be expired by class date — treat as priority
    }
    return 4; // Still valid on class date
  }

  // Never completed
  if (person.bucket === "needed") return 3;

  // Fallback
  return 5;
}


// ************************************************************
//
//   DATE GENERATION
//
// ************************************************************

function generateUpcomingDates(config, today) {
  var dates = [];
  var schedule = config.schedule || {};
  var weeksOut = config.weeksOut || 4;

  // Add any hardcoded specific dates
  if (schedule.dates && schedule.dates.length > 0) {
    for (var d = 0; d < schedule.dates.length; d++) {
      var parsed = parseClassDate(schedule.dates[d]);
      if (parsed && parsed >= today) {
        dates.push(parsed);
      }
    }
  }

  // Process each recurring rule
  var recurring = schedule.recurring || [];
  var dayMap = {
    "sunday": 0, "monday": 1, "tuesday": 2, "wednesday": 3,
    "thursday": 4, "friday": 5, "saturday": 6
  };

  for (var rc = 0; rc < recurring.length; rc++) {
    var rule = recurring[rc];
    if (!rule.weekday) continue;

    var targetDay = dayMap[rule.weekday.toLowerCase()];
    if (targetDay === undefined) continue;

    var endDate = new Date(today);
    endDate.setDate(endDate.getDate() + (weeksOut * 7));

    var cursor = new Date(today);
    while (cursor.getDay() !== targetDay) {
      cursor.setDate(cursor.getDate() + 1);
    }

    while (cursor <= endDate) {
      var skip = false;

      if (rule.nthWeek && rule.nthWeek.length > 0) {
        var weekNum = getWeekOfMonth(cursor);
        var inScope = false;
        for (var w = 0; w < rule.nthWeek.length; w++) {
          if (rule.nthWeek[w] === weekNum) { inScope = true; break; }
        }
        if (!inScope) skip = true;
      }

      if (!skip) {
        var isDupe = false;
        for (var dd = 0; dd < dates.length; dd++) {
          if (dates[dd].getTime() === cursor.getTime()) { isDupe = true; break; }
        }
        if (!isDupe) dates.push(new Date(cursor));
      }

      cursor.setDate(cursor.getDate() + 7);
    }
  }

  dates.sort(function(a, b) { return a - b; });
  return dates;
}

function getWeekOfMonth(date) {
  return Math.ceil(date.getDate() / 7);
}


// ************************************************************
//
//   CLASS ROSTER TAB OUTPUT
//
// ************************************************************

function writeClassRosterTab(sheet, trainingName, classInfo, capacity, today) {
  var NAVY = "#1F3864";
  var RED = "#C00000";
  var ORANGE = "#E65100";
  var GREEN = "#2E7D32";
  var LIGHT_GRAY = "#F2F2F2";
  var WHITE = "#FFFFFF";

  var row = 1;

  sheet.getRange(row, 1, 1, 5).merge();
  sheet.getRange(row, 1).setValue(trainingName + " — Class Roster");
  sheet.getRange(row, 1).setFontSize(14).setFontWeight("bold")
    .setFontColor(WHITE).setBackground(NAVY).setFontFamily("Arial");
  row++;

  sheet.getRange(row, 1).setValue("Class Date:");
  sheet.getRange(row, 1).setFontWeight("bold").setFontFamily("Arial");
  sheet.getRange(row, 2).setValue(formatClassDate(classInfo.date));
  sheet.getRange(row, 2).setFontFamily("Arial");
  row++;

  sheet.getRange(row, 1).setValue("Generated:");
  sheet.getRange(row, 1).setFontWeight("bold").setFontFamily("Arial");
  sheet.getRange(row, 2).setValue(
    Utilities.formatDate(today, Session.getScriptTimeZone(), "M/d/yyyy h:mm a")
  );
  sheet.getRange(row, 2).setFontSize(9).setFontColor("#666666").setFontFamily("Arial");
  row++;

  sheet.getRange(row, 1).setValue("Capacity:");
  sheet.getRange(row, 1).setFontWeight("bold").setFontFamily("Arial");
  sheet.getRange(row, 2).setValue(classInfo.people.length + " / " + capacity);
  sheet.getRange(row, 2).setFontFamily("Arial");
  row++;

  sheet.getRange(row, 1).setValue("HR Program Coordinator: Kyle Mahoney");
  sheet.getRange(row, 1).setFontSize(9).setFontColor("#666666").setFontFamily("Arial");
  row++;
  row++;

  var headers = ["#", "Name", "Status", "Last Completed", "Priority"];
  sheet.getRange(row, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(row, 1, 1, headers.length)
    .setFontWeight("bold").setBackground(LIGHT_GRAY)
    .setFontFamily("Arial").setFontSize(10);
  row++;

  for (var i = 0; i < classInfo.people.length; i++) {
    var person = classInfo.people[i];
    var priorityLabel = "";
    var labelColor = NAVY;

    // Use effectiveBucket if available (accounts for will-expire-by-class-date)
    var displayBucket = person.effectiveBucket || person.bucket;

    if (displayBucket === "expired") {
      priorityLabel = person.bucket === "expired" ? "EXPIRED" : "EXPIRES BEFORE CLASS";
      labelColor = RED;
    } else if (displayBucket === "needed") {
      priorityLabel = "NEVER COMPLETED";
      labelColor = NAVY;
    } else if (displayBucket === "expiringSoon") {
      priorityLabel = "EXPIRING SOON";
      labelColor = ORANGE;
    }

    var vals = [i + 1, person.name, person.status, person.lastDate, priorityLabel];
    sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
    sheet.getRange(row, 1, 1, vals.length).setFontFamily("Arial").setFontSize(10);
    sheet.getRange(row, 5).setFontColor(WHITE).setBackground(labelColor).setFontWeight("bold");
    sheet.getRange(row, 3).setFontColor(labelColor);
    row++;
  }

  var emptySeats = capacity - classInfo.people.length;
  if (emptySeats > 0) {
    row++;
    sheet.getRange(row, 1, 1, 5).merge();
    sheet.getRange(row, 1).setValue(emptySeats + " open seat(s) remaining");
    sheet.getRange(row, 1).setFontColor(GREEN).setFontStyle("italic").setFontFamily("Arial");
  }

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 160);

  sheet.setFrozenRows(7);
}

function buildTabName(trainingName, date) {
  var d = formatClassDate(date);
  var name = trainingName + " " + d;
  name = name.replace(/[:\\\/?*\[\]]/g, "");
  if (name.length > 100) name = name.substring(0, 100);
  return name;
}


// ************************************************************
//
//   DATE HELPERS
//
// ************************************************************

function formatClassDate(d) {
  if (!d) return "";
  return (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear();
}

function parseClassDate(str) {
  if (!str) return null;
  str = str.toString().trim();

  var parts = str.split("/");
  if (parts.length === 3) {
    var mo = parseInt(parts[0]) - 1;
    var da = parseInt(parts[1]);
    var yr = parseInt(parts[2]);
    if (yr < 100) yr += 2000;
    var d = new Date(yr, mo, da);
    d.setHours(0, 0, 0, 0);
    if (!isNaN(d.getTime())) return d;
  }

  var d2 = new Date(str);
  if (!isNaN(d2.getTime())) {
    d2.setHours(0, 0, 0, 0);
    return d2;
  }

  return null;
}

function parseDateList(input) {
  var lines = input.split(/[\n,]+/);
  var dates = [];
  for (var i = 0; i < lines.length; i++) {
    var d = parseClassDate(lines[i].trim());
    if (d) dates.push(d);
  }
  dates.sort(function(a, b) { return a - b; });
  return dates;
}


// ************************************************************
//
//   UPDATED writeRosterSheet — REPLACE IN Code.gs
//
// ************************************************************
// This replaces the existing writeRosterSheet function.
// It adds a 6th "Scheduled" column and scans class roster
// tabs to mark anyone already assigned to an upcoming class.
//
// IMPORTANT: Delete or comment out the OLD writeRosterSheet
// in Code.gs, then this version (defined here in ClassRosters.gs)
// will be the one that runs. Apps Script uses the last-defined
// version of a function across all .gs files.
//
// OR: Copy this function into Code.gs replacing the old one.
// Either approach works.
// ************************************************************

function writeRosterSheet(ss, allRosters, today) {
  var existing = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(ROSTER_SHEET_NAME);

  var NAVY = "#1F3864";
  var RED = "#C00000";
  var ORANGE = "#E65100";
  var GREEN = "#2E7D32";
  var BLUE = "#1565C0";
  var LIGHT_GRAY = "#F2F2F2";
  var WHITE = "#FFFFFF";

  // Scan class roster tabs for scheduled people
  var alreadyScheduled = {};
  try {
    alreadyScheduled = scanClassRosterTabs(ss);
  } catch (e) {
    Logger.log("scanClassRosterTabs not available: " + e.toString());
  }

  var row = 1;

  sheet.getRange(row, 1).setValue("EVC Training Rosters");
  sheet.getRange(row, 1).setFontSize(16).setFontWeight("bold").setFontColor(NAVY).setFontFamily("Arial");
  row++;

  sheet.getRange(row, 1).setValue("Generated: " + Utilities.formatDate(today, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy 'at' h:mm a"));
  sheet.getRange(row, 1).setFontSize(10).setFontColor("#666666").setFontFamily("Arial");
  row++;

  sheet.getRange(row, 1).setValue("HR Program Coordinator: Kyle Mahoney");
  sheet.getRange(row, 1).setFontSize(10).setFontColor("#666666").setFontFamily("Arial");
  row++;
  row++;

  for (var t = 0; t < allRosters.length; t++) {
    var roster = allRosters[t];

    var renewalLabel = roster.renewalYears === 0 ? "Indefinite" : roster.renewalYears + " Year Renewal";
    if (roster.error) renewalLabel = "ERROR";
    var reqLabel = roster.required ? " | REQUIRED FOR ALL" : " | Excusable";

    sheet.getRange(row, 1, 1, 6).merge();
    sheet.getRange(row, 1).setValue(roster.name + "  (" + renewalLabel + reqLabel + ")");
    sheet.getRange(row, 1).setFontSize(13).setFontWeight("bold").setFontColor(WHITE).setBackground(NAVY).setFontFamily("Arial");
    row++;

    if (roster.error) {
      sheet.getRange(row, 1).setValue(roster.error);
      sheet.getRange(row, 1).setFontColor(RED).setFontStyle("italic");
      row += 2;
      continue;
    }

    var totalFlagged = roster.expired.length + roster.expiringSoon.length + roster.needed.length;

    if (totalFlagged === 0) {
      sheet.getRange(row, 1).setValue("All staff are current. No action needed.");
      sheet.getRange(row, 1).setFontColor(GREEN).setFontWeight("bold").setFontFamily("Arial");
      row += 2;
      continue;
    }

    // Column headers — now includes Scheduled
    var colHeaders = ["Name", "Status", "Last Completed", "Expiration Date", "Priority", "Scheduled"];
    sheet.getRange(row, 1, 1, colHeaders.length).setValues([colHeaders]);
    sheet.getRange(row, 1, 1, colHeaders.length).setFontWeight("bold").setBackground(LIGHT_GRAY).setFontFamily("Arial");
    row++;

    // Get scheduled map for this training
    var trainKey = roster.name.toLowerCase();
    var scheduledMap = alreadyScheduled[trainKey] || {};

    // Write expired rows
    for (var e = 0; e < roster.expired.length; e++) {
      var emp = roster.expired[e];
      var personLower = emp.name.toLowerCase().trim();
      var schedDate = scheduledMap[personLower] || "";
      var schedLabel = schedDate ? "Scheduled (" + schedDate + ")" : "";

      var vals = [emp.name, emp.status, emp.lastDate, emp.expDate, "EXPIRED", schedLabel];
      sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
      sheet.getRange(row, 5).setFontColor(WHITE).setBackground(RED).setFontWeight("bold");
      sheet.getRange(row, 1, 1, 4).setFontColor(RED);
      if (schedDate) {
        sheet.getRange(row, 6).setFontColor(WHITE).setBackground(BLUE).setFontWeight("bold");
      }
      row++;
    }

    // Write expiring soon rows
    for (var s = 0; s < roster.expiringSoon.length; s++) {
      var emp = roster.expiringSoon[s];
      var personLower = emp.name.toLowerCase().trim();
      var schedDate = scheduledMap[personLower] || "";
      var schedLabel = schedDate ? "Scheduled (" + schedDate + ")" : "";

      var vals = [emp.name, emp.status, emp.lastDate, emp.expDate, "EXPIRING SOON", schedLabel];
      sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
      sheet.getRange(row, 5).setFontColor(WHITE).setBackground(ORANGE).setFontWeight("bold");
      sheet.getRange(row, 1, 1, 4).setFontColor(ORANGE);
      if (schedDate) {
        sheet.getRange(row, 6).setFontColor(WHITE).setBackground(BLUE).setFontWeight("bold");
      }
      row++;
    }

    // Write needs training rows
    for (var n = 0; n < roster.needed.length; n++) {
      var emp = roster.needed[n];
      var personLower = emp.name.toLowerCase().trim();
      var schedDate = scheduledMap[personLower] || "";
      var schedLabel = schedDate ? "Scheduled (" + schedDate + ")" : "";

      var vals = [emp.name, emp.status, "", "", "NEEDS TRAINING", schedLabel];
      sheet.getRange(row, 1, 1, vals.length).setValues([vals]);
      sheet.getRange(row, 5).setFontColor(WHITE).setBackground(NAVY).setFontWeight("bold");
      if (schedDate) {
        sheet.getRange(row, 6).setFontColor(WHITE).setBackground(BLUE).setFontWeight("bold");
      }
      row++;
    }

    row++;
  }

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 200);

  sheet.setFrozenRows(4);
}


// ************************************************************
//
//   TAB ORDERING
//
// ************************************************************

// ============================================================
// orderClassRosterTabs — sorts class roster tabs by training
// type (in the order they appear in CLASS_ROSTER_CONFIG) then
// by date within each group. Places them after the core sheets.
// ============================================================
function orderClassRosterTabs(ss) {
  var sheets = ss.getSheets();
  var prefixes = getClassRosterPrefixes();

  // Identify which sheets are class roster tabs
  var classSheets = [];
  var coreSheetCount = 0;

  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    var isClassRoster = false;

    for (var p = 0; p < prefixes.length; p++) {
      if (name.indexOf(prefixes[p] + " ") === 0) {
        isClassRoster = true;

        // Extract date for sorting
        var datePart = name.substring(prefixes[p].length + 1).trim();
        var parsedDate = parseClassDate(datePart);

        // Find config index for group ordering
        var configIdx = -1;
        for (var ci = 0; ci < CLASS_ROSTER_CONFIG.length; ci++) {
          if (CLASS_ROSTER_CONFIG[ci].name === prefixes[p]) {
            configIdx = ci;
            break;
          }
        }

        classSheets.push({
          sheet: sheets[s],
          name: name,
          prefix: prefixes[p],
          configIdx: configIdx >= 0 ? configIdx : 999,
          date: parsedDate,
          dateMs: parsedDate ? parsedDate.getTime() : 0
        });
        break;
      }
    }

    if (!isClassRoster) {
      coreSheetCount++;
    }
  }

  if (classSheets.length === 0) return;

  // Sort: first by config index (training group), then by date
  classSheets.sort(function(a, b) {
    if (a.configIdx !== b.configIdx) return a.configIdx - b.configIdx;
    return a.dateMs - b.dateMs;
  });

  // Move each class roster tab into position after the core sheets
  for (var i = 0; i < classSheets.length; i++) {
    var targetPosition = coreSheetCount + i;
    ss.setActiveSheet(classSheets[i].sheet);
    ss.moveActiveSheet(targetPosition + 1); // 1-indexed
  }
}


// ************************************************************
//
//   SHEET DELETE DETECTION
//
// ************************************************************

// ============================================================
// monitorClassRosterTabs — called by a time-driven trigger
// every minute. Tracks which class roster tabs exist. If any
// disappear (deleted by user), refreshes Training Rosters.
// ============================================================
function monitorClassRosterTabs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getScriptProperties();
    var prefixes = getClassRosterPrefixes();

    // Build current list of class roster tab names
    var sheets = ss.getSheets();
    var currentTabs = [];
    for (var s = 0; s < sheets.length; s++) {
      var name = sheets[s].getName();
      for (var p = 0; p < prefixes.length; p++) {
        if (name.indexOf(prefixes[p] + " ") === 0) {
          currentTabs.push(name);
          break;
        }
      }
    }

    var currentKey = currentTabs.sort().join("|");
    var previousKey = props.getProperty("classRosterTabs") || "";

    if (previousKey && currentKey !== previousKey) {
      // Tabs changed — something was added or deleted
      // Check if any were removed (deletion)
      var prevTabs = previousKey.split("|");
      var removed = false;
      for (var pt = 0; pt < prevTabs.length; pt++) {
        if (prevTabs[pt] && currentTabs.indexOf(prevTabs[pt]) === -1) {
          removed = true;
          break;
        }
      }

      if (removed) {
        Logger.log("Class roster tab deleted — refreshing Training Rosters");
        generateRostersSilent();
        // Re-order remaining tabs
        orderClassRosterTabs(ss);
      }
    }

    // Save current state
    props.setProperty("classRosterTabs", currentKey);

  } catch (err) {
    Logger.log("monitorClassRosterTabs error: " + err.toString());
  }
}

// ============================================================
// installClassRosterTriggers — RUN ONCE to enable:
//   1. Tab-delete monitoring (checks every minute)
//
// Go to EVC Tools > Install Class Roster Triggers
// ============================================================
function installClassRosterTriggers() {
  var ui = SpreadsheetApp.getUi();

  // Remove any existing monitor triggers
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "monitorClassRosterTabs") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  // Create a time-driven trigger that runs every minute
  ScriptApp.newTrigger("monitorClassRosterTabs")
    .timeBased()
    .everyMinutes(1)
    .create();

  // Snapshot current tabs so the first run has a baseline
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prefixes = getClassRosterPrefixes();
    var sheets = ss.getSheets();
    var currentTabs = [];
    for (var s = 0; s < sheets.length; s++) {
      var name = sheets[s].getName();
      for (var p = 0; p < prefixes.length; p++) {
        if (name.indexOf(prefixes[p] + " ") === 0) {
          currentTabs.push(name);
          break;
        }
      }
    }
    PropertiesService.getScriptProperties().setProperty(
      "classRosterTabs", currentTabs.sort().join("|")
    );
  } catch (e) {
    Logger.log("Baseline snapshot error: " + e.toString());
  }

  ui.alert(
    "Class Roster triggers installed!\n\n" +
    (removed > 0 ? "Removed " + removed + " old trigger(s) first.\n\n" : "") +
    "What this does:\n" +
    "  Checks every minute if any class roster tabs were deleted.\n" +
    "  If so, automatically refreshes the Training Rosters tab\n" +
    "  and re-orders remaining class roster tabs.\n\n" +
    "You only need to run this once. It persists across sessions."
  );
}


// ************************************************************
//
//   UPDATED createMenu — REPLACE IN Code.gs
//
// ************************************************************
// Uncomment this and delete/comment the one in Code.gs,
// or just add the new lines to your existing menu.
// ************************************************************

// function createMenu() {
//   SpreadsheetApp.getUi().createMenu("EVC Tools")
//     .addItem("Set Session End Time & Length (batch)", "batchSetSessionInfo")
//     .addItem("Batch Pass/Fail for a session", "batchPassFail")
//     .addSeparator()
//     .addItem("Flag late arrivals", "flagLateArrivals")
//     .addItem("Flag early departures", "flagEarlyDepartures")
//     .addSeparator()
//     .addItem("Generate Training Rosters", "generateRosters")
//     .addItem("Generate Roster for One Training", "generateSingleRoster")
//     .addItem("Generate Class Rosters", "generateClassRosters")
//     .addItem("View Roster Config", "showConfig")
//     .addSeparator()
//     .addItem("Backfill Training Access from Records", "backfillTrainingAccess")
//     .addSeparator()
//     .addItem("Install Auto-Refresh Trigger (run once)", "installEditTrigger")
//     .addItem("Install Class Roster Triggers (run once)", "installClassRosterTriggers")
//     .addSeparator()
//     .addItem("Test Training Access connection", "testTrainingAccessConnection")
//     .addItem("Test Name Check", "testCheckName")
//     .addItem("How to export to Excel", "exportReminder")
//     .addToUi();
// }
