/**
 * Debate Tools — Classroom Debate Topic Selector & Team Matcher
 *
 * Workflow (also available any time via Debate Tools > "What do I do next?"):
 *   0. Start new debate cycle   (before sharing the Round 1 form)
 *   1. Run Borda Count          (after Round 1 votes are in)
 *   2. Generate Round 2 Form    (share the link it gives you)
 *   3. Match Teams              (after Round 2 responses are in)
 */

var PROP_CYCLE_START = 'CYCLE_START';
var PROP_ROUND1_FORM = 'ROUND1_FORM_ID';
var PROP_ROUND2_FORM = 'ROUND2_FORM_ID';
var PROP_TOPIC_A = 'DEBATE_TOPIC_A';
var PROP_TOPIC_B = 'DEBATE_TOPIC_B';

var STANCE_CHOICES = [
  "Strongly Prefer Pro",
  "Prefer Pro",
  "No Preference",
  "Prefer Con",
  "Strongly Prefer Con"
];

/**
 * Creates the Debate Tools menu
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Debate Tools')
      .addItem('❓ What do I do next?', 'showStatus')
      .addSeparator()
      .addItem('0. Start new debate cycle', 'startNewCycle')
      .addItem('1. Run Borda Count (Round 1)', 'runBordaCount')
      .addItem('2. Generate Round 2 Form', 'createRound2Form')
      .addItem('3. Match Teams (Round 2)', 'matchTeams')
      .addToUi();
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function props_() {
  return PropertiesService.getDocumentProperties();
}

/** Timestamp for result tab names. Includes seconds so quick re-runs don't collide. */
function stamp_(ss) {
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM dd hh:mm:ss a");
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Returns the form with this ID, or null if it was deleted or can't be opened. */
function openFormSafe_(formId) {
  if (!formId) return null;
  try { return FormApp.openById(formId); } catch (e) { return null; }
}

/** Finds the response tab linked to a given form ID, or null. */
function findSheetForForm_(formId) {
  if (!formId) return null;
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var url = sheets[i].getFormUrl();
    if (!url) continue;
    try {
      if (FormApp.openByUrl(url).getId() === formId) return sheets[i];
    } catch (e) { /* linked form was deleted; skip */ }
  }
  return null;
}

/**
 * Returns the Round 1 response tab. The first time, it learns it from the
 * active tab (after confirming) and remembers it from then on.
 */
function getRound1Sheet_() {
  var props = props_();
  var known = findSheetForForm_(props.getProperty(PROP_ROUND1_FORM));
  if (known) return known;

  var sheet = SpreadsheetApp.getActiveSheet();
  var url = sheet.getFormUrl();
  if (!url) return null;

  var formId = FormApp.openByUrl(url).getId();
  if (formId === props.getProperty(PROP_ROUND2_FORM)) return null;

  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    'Use "' + sheet.getName() + '" as your Round 1 responses tab from now on?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return null;

  props.setProperty(PROP_ROUND1_FORM, formId);
  return sheet;
}

/** Pulls the topic name out of a grid header like "Question title [Topic]". */
function gridRowFromHeader_(header) {
  var match = String(header).match(/\[([^\[\]]*)\]\s*$/);
  return match ? match[1].trim() : null;
}

/* ------------------------------------------------------------------ */
/* Status / next step                                                  */
/* ------------------------------------------------------------------ */

/**
 * Shows where you are in the workflow, links to both forms, response counts,
 * and what to do next. Optional headline is shown at the top.
 */
function showStatus(headline) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = props_();
  var tz = ss.getSpreadsheetTimeZone();
  var start = Number(props.getProperty(PROP_CYCLE_START) || 0);
  var topicA = props.getProperty(PROP_TOPIC_A);
  var topicB = props.getProperty(PROP_TOPIC_B);
  var r1 = openFormSafe_(props.getProperty(PROP_ROUND1_FORM));
  var r2 = openFormSafe_(props.getProperty(PROP_ROUND2_FORM));

  var html = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6">';
  if (typeof headline === 'string' && headline) {
    html += '<div style="background:#d9ead3;padding:6px 10px;margin-bottom:8px"><b>' +
            escapeHtml_(headline) + '</b></div>';
  }

  html += '<b>Cycle started:</b> ' +
          (start ? Utilities.formatDate(new Date(start), tz, 'MMM d, h:mm a') : 'not yet') + '<br>';

  if (r1) {
    var r1Count = start ? r1.getResponses(new Date(start)).length : r1.getResponses().length;
    html += '<b>Round 1 ballots this cycle:</b> ' + r1Count +
            ' · <a href="' + r1.getPublishedUrl() + '" target="_blank">student link</a>' +
            (r1.isAcceptingResponses() ? '' : ' <i>(closed)</i>') + '<br>';
  } else {
    html += '<b>Round 1 tab:</b> not set yet (remembered the first time you run step 1)<br>';
  }

  if (topicA && topicB) {
    html += '<b>Winning topics:</b> ' + escapeHtml_(topicA) + ' / ' + escapeHtml_(topicB) + '<br>';
  }

  if (r2) {
    html += '<b>Round 2 responses:</b> ' + r2.getResponses().length +
            ' · <a href="' + r2.getPublishedUrl() + '" target="_blank">student link</a>' +
            ' · <a href="' + r2.getEditUrl() + '" target="_blank">edit form</a>' +
            (r2.isAcceptingResponses() ? '' : ' <i>(closed)</i>') + '<br>';
  }

  var next;
  if (!start) {
    next = 'Choose <b>0. Start new debate cycle</b>, then share your Round 1 form with students.';
  } else if (!topicA || !topicB) {
    next = 'When Round 1 votes are in, choose <b>1. Run Borda Count</b>.';
  } else if (!r2) {
    next = 'Choose <b>2. Generate Round 2 Form</b>.';
  } else if (r2.isAcceptingResponses()) {
    next = 'Share the Round 2 student link. When everyone has submitted, choose <b>3. Match Teams</b>.';
  } else {
    next = 'Teams are done. Next time, start with <b>0. Start new debate cycle</b>.';
  }
  html += '<hr><b>Next step:</b> ' + next + '</div>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(300),
    'Debate Tools: Status');
}

/* ------------------------------------------------------------------ */
/* Step 0: Start a new cycle                                           */
/* ------------------------------------------------------------------ */

function startNewCycle() {
  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    'Start a new debate cycle?',
    'Round 1 will only count ballots submitted from now on. Old result tabs are kept.',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  var props = props_();
  props.deleteProperty(PROP_TOPIC_A);
  props.deleteProperty(PROP_TOPIC_B);
  props.deleteProperty(PROP_ROUND2_FORM);
  props.setProperty(PROP_CYCLE_START, String(Date.now()));

  // Reopen the Round 1 form if we know which one it is
  var r1 = openFormSafe_(props.getProperty(PROP_ROUND1_FORM));
  if (r1 && !r1.isAcceptingResponses()) r1.setAcceptingResponses(true);

  showStatus('New cycle started.');
}

/* ------------------------------------------------------------------ */
/* Step 1: Borda Count                                                 */
/* ------------------------------------------------------------------ */

function runBordaCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var props = props_();

  var sheet = getRound1Sheet_();
  if (!sheet) {
    ui.alert("Click your Round 1 'Form Responses' tab and run this again. (You only have to do this once.)");
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("Not enough data to calculate. Wait for student submissions.");
    return;
  }

  // Find the grid columns
  var headers = data[0];
  var topicCols = [];
  for (var i = 0; i < headers.length; i++) {
    var topic = gridRowFromHeader_(headers[i]);
    if (topic) topicCols.push({ index: i, topic: topic });
  }

  // Ignore columns for topics that are no longer on the form (Sheets keeps old columns)
  var formUrl = sheet.getFormUrl();
  if (formUrl) {
    var current = {};
    FormApp.openByUrl(formUrl).getItems(FormApp.ItemType.GRID).forEach(function(item) {
      item.asGridItem().getRows().forEach(function(r) { current[String(r).trim()] = true; });
    });
    topicCols = topicCols.filter(function(tc) { return current[tc.topic]; });
  }

  if (topicCols.length < 2) {
    ui.alert("Error: Could not find at least two ranking columns. Make sure your form uses a Multiple Choice Grid.");
    return;
  }

  var numTopics = topicCols.length;
  var scores = {};
  topicCols.forEach(function(tc) { scores[tc.topic] = 0; });

  var cycleStart = Number(props.getProperty(PROP_CYCLE_START) || 0);
  var validVotes = 0;
  var skippedOld = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];

    // Skip ballots from before this cycle started (column A is the form Timestamp)
    if (cycleStart && row[0] && new Date(row[0]).getTime() < cycleStart) {
      skippedOld++;
      continue;
    }

    var rowHasData = false;
    for (var c = 0; c < topicCols.length; c++) {
      var rank = parseInt(row[topicCols[c].index], 10);
      if (!isNaN(rank)) {
        scores[topicCols[c].topic] += numTopics - rank + 1;
        rowHasData = true;
      }
    }
    if (rowHasData) validVotes++;
  }

  if (validVotes === 0) {
    ui.alert("Error: No valid rankings found" +
      (skippedOld ? " this cycle (" + skippedOld + " older ballots were ignored)." : "."));
    return;
  }

  var results = [];
  for (var topicName in scores) {
    results.push({ topic: topicName, score: scores[topicName] });
  }

  // Documented tie-breaker: higher score first, then alphabetical
  results.sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return a.topic.localeCompare(b.topic);
  });

  // Only a tie at the 2nd/3rd boundary changes which topics win
  var tieWarning = "";
  if (results.length > 2 && results[1].score === results[2].score) {
    var tied = results.filter(function(x) { return x.score === results[1].score; })
                      .map(function(x) { return "'" + x.topic + "'"; });
    tieWarning = "\n\nWARNING: " + tied.join(", ") + " tied at " + results[1].score +
      " points for the last winning spot. The tie was broken alphabetically. Review the Borda Results tab.";
  }

  // --- AUDIT TRAIL SHEET ---
  var resultSheet = ss.insertSheet("Borda Results (" + stamp_(ss) + ")");
  resultSheet.getRange(1, 1, 1, 3).setValues([["Rank", "Debate Topic", "Total Points"]]);

  var outputData = results.map(function(x, j) { return [j + 1, x.topic, x.score]; });
  resultSheet.getRange(2, 1, outputData.length, 3).setValues(outputData);

  resultSheet.getRange("A1:C1").setFontWeight("bold");
  resultSheet.getRange(1, 1, outputData.length + 1, 1).setHorizontalAlignment("center");
  resultSheet.getRange(1, 3, outputData.length + 1, 1).setHorizontalAlignment("center");
  resultSheet.getRange(2, 1, Math.min(2, results.length), 3).setBackground("#d9ead3");

  resultSheet.getRange("E1").setValue("Winner Summary").setFontWeight("bold");
  resultSheet.getRange("E2").setValue("Topic A: " + results[0].topic);
  resultSheet.getRange("E3").setValue("Topic B: " + results[1].topic);
  resultSheet.getRange("E5").setValue("Ballots counted: " + validVotes);
  if (skippedOld) resultSheet.getRange("E6").setValue("Older ballots ignored: " + skippedOld);
  resultSheet.autoResizeColumns(1, 5);

  // If winners changed after a Round 2 form was already made, that form is stale
  var staleNote = "";
  var oldA = props.getProperty(PROP_TOPIC_A);
  var oldB = props.getProperty(PROP_TOPIC_B);
  var winnersChanged = oldA && oldB &&
      [oldA, oldB].sort().join('|') !== [results[0].topic, results[1].topic].sort().join('|');
  if (winnersChanged && props.getProperty(PROP_ROUND2_FORM)) {
    props.deleteProperty(PROP_ROUND2_FORM);
    staleNote = "\n\nNOTE: The winners changed since your Round 2 form was created. " +
      "Run step 2 again to make a new form, and stop sharing the old one.";
  }

  props.setProperty(PROP_TOPIC_A, results[0].topic);
  props.setProperty(PROP_TOPIC_B, results[1].topic);

  var tip = cycleStart ? "" :
    "\n\nTip: Next time, choose '0. Start new debate cycle' before sharing the Round 1 form so old ballots are ignored.";

  resultSheet.activate();
  ui.alert("Calculation complete!\n\nCounted " + validVotes + " ballots" +
    (skippedOld ? " (ignored " + skippedOld + " from before this cycle)" : "") + "." +
    tieWarning + staleNote + tip +
    "\n\nNext step: Debate Tools > 2. Generate Round 2 Form.");
}

/* ------------------------------------------------------------------ */
/* Step 2: Generate Round 2 Form                                       */
/* ------------------------------------------------------------------ */

function createRound2Form() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var props = props_();
  var topicA = props.getProperty(PROP_TOPIC_A);
  var topicB = props.getProperty(PROP_TOPIC_B);

  if (!topicA || !topicB) {
    ui.alert("Missing Round 1 winners. Run '1. Run Borda Count' first.");
    return;
  }

  // Guard against making a second form for the same cycle by accident
  if (openFormSafe_(props.getProperty(PROP_ROUND2_FORM))) {
    var answer = ui.alert(
      'A Round 2 form already exists for this cycle.',
      'Create another one anyway? (Choose No and use "What do I do next?" to get the existing link.)',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
  }

  var dateLabel = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM d, yyyy");
  var form = FormApp.create('Debate Round 2: ' + topicA + ' vs. ' + topicB + ' (' + dateLabel + ')');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty(PROP_ROUND2_FORM, form.getId());

  form.addTextItem()
      .setTitle("Student Name")
      .setHelpText("First and last name")
      .setRequired(true);

  var grid = form.addGridItem()
      .setTitle("Rank the Debate Topics (1 = First Choice, 2 = Second Choice)")
      .setRows([topicA, topicB])
      .setColumns(["1", "2"])
      .setRequired(true);
  grid.setValidation(FormApp.createGridValidation()
      .setHelpText("Use each rank only once.")
      .requireLimitOneResponsePerColumn()
      .build());

  form.addMultipleChoiceItem()
      .setTitle("Preferred Debate Side (Applies to whichever topic you are assigned)")
      .setChoiceValues(STANCE_CHOICES)
      .setRequired(true);

  showStatus('Round 2 form created. Share the student link below.');
}

/* ------------------------------------------------------------------ */
/* Step 3: Match Teams                                                 */
/* ------------------------------------------------------------------ */

function matchTeams() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var props = props_();

  var formId = props.getProperty(PROP_ROUND2_FORM);
  var sheet = findSheetForForm_(formId);
  if (!sheet) {
    ui.alert("Couldn't find this cycle's Round 2 responses. Run '2. Generate Round 2 Form' first.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    ui.alert("Not enough data. Wait for students to submit Round 2.");
    return;
  }

  // --- FIND COLUMNS ---
  var headers = data[0];
  var nameCol = -1, stanceCol = -1;
  var topicCols = [];

  for (var i = 0; i < headers.length; i++) {
    var topic = gridRowFromHeader_(headers[i]);
    if (topic) {
      topicCols.push({ index: i, topic: topic });
      continue;
    }
    var header = String(headers[i]).toLowerCase();
    if (nameCol === -1 && header.indexOf('name') !== -1) {
      nameCol = i;
    } else if (stanceCol === -1 && (header.indexOf('side') !== -1 || header.indexOf('preference') !== -1)) {
      stanceCol = i;
    }
  }

  if (nameCol === -1 || stanceCol === -1 || topicCols.length < 2) {
    ui.alert("Could not identify the name, topic ranking, and side columns on the Round 2 responses tab.");
    return;
  }

  var topics = topicCols.map(function(t) { return t.topic; });

  // --- READ STUDENTS (latest submission per name wins) ---
  var byName = {};
  var dupes = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name = String(row[nameCol] || '').trim().replace(/\s+/g, ' ');
    if (!name) continue;

    var prefs = [];
    for (var c = 0; c < topicCols.length; c++) {
      var rank = parseInt(row[topicCols[c].index], 10);
      if (!isNaN(rank)) prefs.push({ topic: topicCols[c].topic, rank: rank });
    }
    prefs.sort(function(a, b) { return a.rank - b.rank; });
    var topicPrefs = prefs.map(function(p) { return p.topic; });
    // Fill in any unranked topics so every student has a complete list
    topics.forEach(function(t) { if (topicPrefs.indexOf(t) === -1) topicPrefs.push(t); });

    var key = name.toLowerCase();
    if (byName[key]) dupes++;
    byName[key] = {
      name: name,
      topicPrefs: topicPrefs,
      stancePref: String(row[stanceCol] || '').trim(),
      assignedTopic: null,
      assignedStance: null,
      topicGranted: null,
      stanceGranted: null
    };
  }

  var students = Object.keys(byName).map(function(k) { return byName[k]; });
  if (students.length === 0) {
    ui.alert("No student names found on the Round 2 responses tab.");
    return;
  }

  function shuffleArray(array) {
    for (var i = array.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var temp = array[i]; array[i] = array[j]; array[j] = temp;
    }
  }

  function getStanceScore(pref) {
    if (pref === "Strongly Prefer Pro") return 2;
    if (pref === "Prefer Pro") return 1;
    if (pref === "Strongly Prefer Con") return -2;
    if (pref === "Prefer Con") return -1;
    return 0;
  }

  // --- PHASE 1: TOPIC ASSIGNMENT ---
  shuffleArray(students);
  var topicCapacities = {};
  var baseCap = Math.floor(students.length / topics.length);
  topics.forEach(function(t) { topicCapacities[t] = baseCap; });

  var remainder = students.length % topics.length;
  if (remainder > 0) {
    var firstChoiceCounts = {};
    topics.forEach(function(t) { firstChoiceCounts[t] = 0; });
    students.forEach(function(s) { firstChoiceCounts[s.topicPrefs[0]]++; });
    var popSorted = topics.slice().sort(function(a, b) { return firstChoiceCounts[b] - firstChoiceCounts[a]; });
    for (var p = 0; p < remainder; p++) topicCapacities[popSorted[p]]++;
  }

  var topicTeams = {};
  topics.forEach(function(t) { topicTeams[t] = []; });

  students.forEach(function(s) {
    var choice = s.topicPrefs[0];
    if (topicTeams[choice].length < topicCapacities[choice]) {
      s.assignedTopic = choice; s.topicGranted = 1; topicTeams[choice].push(s);
    }
  });

  students.filter(function(s) { return !s.assignedTopic; }).forEach(function(s) {
    var choice = s.topicPrefs[1];
    if (topicTeams[choice].length < topicCapacities[choice]) {
      s.assignedTopic = choice; s.topicGranted = 2; topicTeams[choice].push(s);
    }
  });

  // --- PHASE 2: PRO/CON BALANCING ---
  topics.forEach(function(t) {
    var team = topicTeams[t];
    shuffleArray(team);  // randomizes order within each score tier

    team.sort(function(a, b) {
      return getStanceScore(b.stancePref) - getStanceScore(a.stancePref);
    });

    var capPro = Math.floor(team.length / 2);
    if (team.length % 2 !== 0) {
      var proDemand = team.filter(function(s) { return getStanceScore(s.stancePref) > 0; }).length;
      var conDemand = team.filter(function(s) { return getStanceScore(s.stancePref) < 0; }).length;
      if (conDemand <= proDemand) capPro++;
    }

    for (var i = 0; i < team.length; i++) {
      var s = team[i];
      var score = getStanceScore(s.stancePref);
      if (i < capPro) {
        s.assignedStance = 'Pro';
        s.stanceGranted = score > 0 ? 'Yes' : (score === 0 ? 'N/A' : 'No');
      } else {
        s.assignedStance = 'Con';
        s.stanceGranted = score < 0 ? 'Yes' : (score === 0 ? 'N/A' : 'No');
      }
    }
  });

  // Safety check for unassigned students
  var failed = students.filter(function(s) { return !s.assignedTopic || !s.assignedStance; });
  if (failed.length > 0) {
    ui.alert("Error: Algorithm failed to assign " + failed.length + " student(s). Check capacities.");
    return;
  }

  var stats = { topic1: 0, topic2: 0, stanceYes: 0, stanceNo: 0, stanceNA: 0 };
  students.forEach(function(s) {
    if (s.topicGranted === 1) stats.topic1++;
    if (s.topicGranted === 2) stats.topic2++;
    if (s.stanceGranted === 'Yes') stats.stanceYes++;
    if (s.stanceGranted === 'No') stats.stanceNo++;
    if (s.stanceGranted === 'N/A') stats.stanceNA++;
  });

  // Roster order: topic, then Pro before Con, then name
  students.sort(function(a, b) {
    if (a.assignedTopic !== b.assignedTopic) return a.assignedTopic.localeCompare(b.assignedTopic);
    if (a.assignedStance !== b.assignedStance) return b.assignedStance.localeCompare(a.assignedStance);
    return a.name.localeCompare(b.name);
  });

  // --- OUTPUT ---
  var rosterSheet = ss.insertSheet("Final Teams (" + stamp_(ss) + ")");

  var report = [
    ["Match Satisfaction Report", "Count"],
    ["Topic: 1st Choice", stats.topic1],
    ["Topic: 2nd Choice", stats.topic2],
    ["Stance: Preferred (Strongly or Normal)", stats.stanceYes],
    ["Stance: Forced to Switch", stats.stanceNo],
    ["Stance: Had No Preference", stats.stanceNA]
  ];
  rosterSheet.getRange(1, 1, report.length, 2).setValues(report);
  rosterSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#cfe2f3");

  rosterSheet.getRange("A8:E8")
      .setValues([["Debate Topic", "Side", "Student Name", "Topic Choice", "Stance Granted"]])
      .setFontWeight("bold").setBackground("#f3f3f3");

  var outputData = students.map(function(s) {
    return [s.assignedTopic, s.assignedStance, s.name, s.topicGranted, s.stanceGranted];
  });
  rosterSheet.getRange(9, 1, outputData.length, 5).setValues(outputData);
  rosterSheet.autoResizeColumns(1, 5);
  rosterSheet.getRange(9, 4, outputData.length, 2).setHorizontalAlignment("center");

  // Close Round 2 so late submissions don't arrive after teams are posted
  var r2 = openFormSafe_(formId);
  if (r2) r2.setAcceptingResponses(false);

  rosterSheet.activate();
  ui.alert("Assignments complete! " + students.length + " students placed." +
    (dupes ? "\n\n" + dupes + " duplicate submission(s) found; each student's latest one was used." : "") +
    "\n\nThe Round 2 form is now closed. To accept late responses, reopen it from the form's Responses tab and run step 3 again.");
}
