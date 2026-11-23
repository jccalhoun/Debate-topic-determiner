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
var PROP_NUM_DEBATES = 'NUM_DEBATES';      // '1' or '2' (missing = 2, for older cycles)
var PROP_ROUND2_TOPICS = 'ROUND2_TOPICS';  // JSON list of topics on the current Round 2 form
var PROP_ROUND2_SLOTS = 'ROUND2_SLOTS';    // JSON list of {label, topic, side} choices on a two-debate Round 2 form

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

/** True if this tab's linked form is the given form. */
function sheetIsLinkedTo_(sheet, formId) {
  var url = sheet.getFormUrl();
  if (!url) return false;
  try { return FormApp.openByUrl(url).getId() === formId; } catch (e) { return false; }
}

/**
 * Finds the response tab linked to a given form ID, or null.
 * Checks the tab cached from last time first (one form lookup instead of one
 * per linked tab), then scans newest tabs first and caches what it finds.
 */
function findSheetForForm_(formId) {
  if (!formId) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = props_();
  var cacheKey = 'SHEET_FOR_' + formId;
  var sheets = ss.getSheets();

  var cachedId = props.getProperty(cacheKey);
  if (cachedId) {
    for (var i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getSheetId()) === cachedId) {
        if (sheetIsLinkedTo_(sheets[i], formId)) return sheets[i];
        break;
      }
    }
  }

  for (var j = sheets.length - 1; j >= 0; j--) {
    if (String(sheets[j].getSheetId()) === cachedId) continue;
    if (sheetIsLinkedTo_(sheets[j], formId)) {
      props.setProperty(cacheKey, String(sheets[j].getSheetId()));
      return sheets[j];
    }
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

  var formId;
  try { formId = FormApp.openByUrl(url).getId(); } catch (e) { return null; }
  if (formId === props.getProperty(PROP_ROUND2_FORM)) return null;

  var ui = SpreadsheetApp.getUi();
  var answer = ui.alert(
    'Use "' + sheet.getName() + '" as your Round 1 responses tab from now on?',
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return null;

  props.setProperty(PROP_ROUND1_FORM, formId);
  return sheet;
}

/**
 * Moves a file into the same Drive folder as the Round 1 form (or, if Round 1
 * isn't known, the spreadsheet's folder). Returns the folder name, or null if
 * it couldn't be moved. Never throws: the form still works wherever it is.
 */
function moveNextToRound1_(fileId) {
  try {
    var anchorId = props_().getProperty(PROP_ROUND1_FORM) ||
                   SpreadsheetApp.getActiveSpreadsheet().getId();
    var parents = DriveApp.getFileById(anchorId).getParents();
    if (!parents.hasNext()) return null;
    var folder = parents.next();
    DriveApp.getFileById(fileId).moveTo(folder);
    return folder.getName();
  } catch (e) {
    return null;
  }
}

/** This cycle's Round 1 winners: one topic or two. Empty if step 1 hasn't run. */
function savedWinners_() {
  var props = props_();
  var a = props.getProperty(PROP_TOPIC_A);
  var b = props.getProperty(PROP_TOPIC_B);
  if (!a) return [];
  if (props.getProperty(PROP_NUM_DEBATES) === '1') return [a];
  return b ? [a, b] : [];
}

/** The topic/side choices on this cycle's two-debate Round 2 form, or []. */
function savedSlots_() {
  try { return JSON.parse(props_().getProperty(PROP_ROUND2_SLOTS) || '[]'); } catch (e) { return []; }
}

/** Row names and column count of every grid question on a form. */
function gridInfo_(form) {
  var rows = [], columns = 0;
  form.getItems(FormApp.ItemType.GRID).forEach(function(item) {
    var g = item.asGridItem();
    g.getRows().forEach(function(r) { if (rows.indexOf(r) === -1) rows.push(r); });
    columns = Math.max(columns, g.getColumns().length);
  });
  return { rows: rows, columns: columns };
}

/**
 * Finds response columns for known topic names. A grid header looks like
 * "Question title [Topic]", so we match the exact "[Topic]" suffix. This works
 * even when a topic name contains brackets, e.g. "TikTok Ban [US]".
 * A topic can match more than one column (deleted and re-added on the form).
 */
function findTopicCols_(headers, topicNames) {
  var cols = [];
  topicNames.forEach(function(topic) {
    var suffix = '[' + topic + ']';
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim();
      if (h.length > suffix.length && h.slice(-suffix.length) === suffix) {
        cols.push({ index: i, topic: topic });
      }
    }
  });
  return cols;
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
  var winners = savedWinners_();
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

  if (winners.length === 1) {
    html += '<b>Winning topic (one debate):</b> ' + escapeHtml_(winners[0]) + '<br>';
  } else if (winners.length === 2) {
    html += '<b>Winning topics:</b> ' + escapeHtml_(winners[0]) + ' / ' + escapeHtml_(winners[1]) + '<br>';
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
  } else if (winners.length === 0) {
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

  // If the previous cycle was abandoned before step 3, close its Round 2 form
  var oldR2 = openFormSafe_(props.getProperty(PROP_ROUND2_FORM));
  if (oldR2 && oldR2.isAcceptingResponses()) oldR2.setAcceptingResponses(false);

  props.deleteProperty(PROP_NUM_DEBATES);
  props.deleteProperty(PROP_ROUND2_TOPICS);
  props.deleteProperty(PROP_ROUND2_SLOTS);
  props.deleteProperty(PROP_TOPIC_A);
  props.deleteProperty(PROP_TOPIC_B);
  props.deleteProperty(PROP_ROUND2_FORM);
  props.setProperty(PROP_CYCLE_START, String(Date.now()));

  // Learn the Round 1 tab now if we don't know it yet, so the status
  // window can show the Round 1 student link on the very first cycle.
  getRound1Sheet_();

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

  // The form's current grid rows are the topics. Columns for topics that were
  // removed from the form are ignored (Sheets keeps old columns around).
  var grid = gridInfo_(FormApp.openByUrl(sheet.getFormUrl()));
  var topicNames = grid.rows;
  var numTopics = topicNames.length;
  var topicCols = findTopicCols_(data[0], topicNames);

  if (numTopics < 2 || topicCols.length < 2) {
    ui.alert("Error: Could not find at least two topic ranking columns. Make sure your Round 1 form uses a Multiple Choice Grid.");
    return;
  }

  var setupWarning = "";
  if (grid.columns !== numTopics) {
    setupWarning = "\n\nCHECK YOUR FORM: the Round 1 grid has " + numTopics + " topics but " +
      grid.columns + " rank columns. These should match (ranks 1 to " + numTopics + ").";
  }

  var scores = {};
  topicNames.forEach(function(t) { scores[t] = 0; });

  var cycleStart = Number(props.getProperty(PROP_CYCLE_START) || 0);
  var validVotes = 0;
  var skippedOld = 0;
  var outOfRange = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];

    // Skip ballots from before this cycle started (column A is the form Timestamp)
    if (cycleStart && row[0] && new Date(row[0]).getTime() < cycleStart) {
      skippedOld++;
      continue;
    }

    var rowHasData = false;
    for (var c = 0; c < topicCols.length; c++) {
      var cell = row[topicCols[c].index];
      if (cell === '' || cell === null) continue;
      var rank = parseInt(cell, 10);
      if (isNaN(rank) || rank < 1 || rank > numTopics) {
        outOfRange++;
        continue;
      }
      scores[topicCols[c].topic] += numTopics - rank + 1;
      rowHasData = true;
    }
    if (rowHasData) validVotes++;
  }

  if (outOfRange) {
    setupWarning += "\n\nIgnored " + outOfRange + " rank(s) outside 1 to " + numTopics +
      ". Check that the Round 1 grid's rank columns match the number of topics.";
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

  // One debate (small class) or two?
  var howMany = ui.alert(
    'How many debates will this class have?',
    'YES = two debates (the top 2 topics).\nNO = one debate (only the top topic), for a small class.',
    ui.ButtonSet.YES_NO);
  if (howMany !== ui.Button.YES && howMany !== ui.Button.NO) return;  // closed the dialog
  var numDebates = (howMany === ui.Button.NO) ? 1 : 2;
  var winners = results.slice(0, numDebates).map(function(x) { return x.topic; });

  // Only a tie at the last winning spot changes which topics win
  // (2nd vs 3rd for two debates, 1st vs 2nd for one)
  var tieWarning = "";
  var cut = numDebates - 1;
  if (results.length > numDebates && results[cut].score === results[numDebates].score) {
    var tied = results.filter(function(x) { return x.score === results[cut].score; })
                      .map(function(x) { return "'" + x.topic + "'"; });
    tieWarning = "\n\nWARNING: " + tied.join(", ") + " tied at " + results[cut].score +
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
  resultSheet.getRange(2, 1, numDebates, 3).setBackground("#d9ead3");

  resultSheet.getRange("E1").setValue("Winner Summary").setFontWeight("bold");
  if (numDebates === 1) {
    resultSheet.getRange("E2").setValue("Topic: " + winners[0]);
    resultSheet.getRange("E3").setValue("(One debate for this class)");
  } else {
    resultSheet.getRange("E2").setValue("Topic A: " + winners[0]);
    resultSheet.getRange("E3").setValue("Topic B: " + winners[1]);
  }
  resultSheet.getRange("E5").setValue("Ballots counted: " + validVotes);
  if (skippedOld) resultSheet.getRange("E6").setValue("Older ballots ignored: " + skippedOld);
  resultSheet.autoResizeColumns(1, 5);

  // If winners changed after a Round 2 form was already made, that form is stale
  var staleNote = "";
  var oldWinners = savedWinners_();
  var winnersChanged = oldWinners.length > 0 &&
      oldWinners.slice().sort().join('|') !== winners.slice().sort().join('|');
  if (winnersChanged && props.getProperty(PROP_ROUND2_FORM)) {
    var staleForm = openFormSafe_(props.getProperty(PROP_ROUND2_FORM));
    if (staleForm) staleForm.setAcceptingResponses(false);
    props.deleteProperty(PROP_ROUND2_FORM);
    staleNote = "\n\nNOTE: The winners changed since your Round 2 form was created, so that form " +
      "has been closed. Run step 2 again to make a new one.";
  }

  props.setProperty(PROP_NUM_DEBATES, String(numDebates));
  props.setProperty(PROP_TOPIC_A, winners[0]);
  if (numDebates === 2) props.setProperty(PROP_TOPIC_B, winners[1]);
  else props.deleteProperty(PROP_TOPIC_B);

  var tip = cycleStart ? "" :
    "\n\nTip: Next time, choose '0. Start new debate cycle' before sharing the Round 1 form so old ballots are ignored.";

  resultSheet.activate();
  ui.alert("Calculation complete!\n\nCounted " + validVotes + " ballots" +
    (skippedOld ? " (ignored " + skippedOld + " from before this cycle)" : "") + "." +
    tieWarning + setupWarning + staleNote + tip +
    "\n\nNext step: Debate Tools > 2. Generate Round 2 Form.");
}

/* ------------------------------------------------------------------ */
/* Step 2: Generate Round 2 Form                                       */
/* ------------------------------------------------------------------ */

function createRound2Form() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var props = props_();
  var winners = savedWinners_();

  if (winners.length === 0) {
    ui.alert("Missing Round 1 winners. Run '1. Run Borda Count' first.");
    return;
  }
  var single = winners.length === 1;

  // Guard against making a second form for the same cycle by accident
  var existing = openFormSafe_(props.getProperty(PROP_ROUND2_FORM));
  if (existing) {
    var answer = ui.alert(
      'A Round 2 form already exists for this cycle.',
      'Create a new one and close the old one? (Choose No and use "What do I do next?" to get the existing link.)',
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
    existing.setAcceptingResponses(false);
  }

  var dateLabel = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMM d, yyyy");
  var form = FormApp.create('Debate Round 2: ' + winners.join(' vs. ') + ' (' + dateLabel + ')');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  props.setProperty(PROP_ROUND2_FORM, form.getId());
  props.setProperty(PROP_ROUND2_TOPICS, JSON.stringify(winners));
  if (single) form.setDescription('Debate topic: ' + winners[0]);

  form.addTextItem()
      .setTitle("Student Name")
      .setHelpText("First and last name")
      .setRequired(true);

  if (single) {
    // One debate: nothing to rank, just which side they'd rather argue
    props.deleteProperty(PROP_ROUND2_SLOTS);
    form.addMultipleChoiceItem()
        .setTitle("Preferred Debate Side")
        .setHelpText("Topic: " + winners[0])
        .setChoiceValues(STANCE_CHOICES)
        .setRequired(true);
  } else {
    // Two debates: rank every topic + side combination together
    var slots = [];
    winners.forEach(function(t) {
      slots.push({ label: t + ' — Pro (for)', topic: t, side: 'Pro' });
      slots.push({ label: t + ' — Con (against)', topic: t, side: 'Con' });
    });
    props.setProperty(PROP_ROUND2_SLOTS, JSON.stringify(slots));

    var columns = slots.map(function(s, k) { return String(k + 1); });
    var grid = form.addGridItem()
        .setTitle("Rank your choices (1 = most wanted, " + slots.length + " = least wanted)")
        .setHelpText("Each choice is a topic plus the side you would argue. Use each rank once.")
        .setRows(slots.map(function(s) { return s.label; }))
        .setColumns(columns)
        .setRequired(true);
    grid.setValidation(FormApp.createGridValidation()
        .setHelpText("Use each rank only once.")
        .requireLimitOneResponsePerColumn()
        .build());
  }

  // Keep each class's forms together: put Round 2 next to the Round 1 form
  var folderName = moveNextToRound1_(form.getId());
  showStatus('Round 2 form created' + (folderName ? ' in folder "' + folderName + '"' : '') +
    '. Share the student link below.');
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

  // --- WHICH KIND OF ROUND 2 FORM? ---
  // Two debates: one grid ranking the topic/side choices ("slots").
  // One debate: no grid, just the 5-point side question.
  var r2Form = openFormSafe_(formId);
  var gridRows = r2Form ? gridInfo_(r2Form).rows : [];
  var slots = savedSlots_();
  var single = gridRows.length === 0;

  if (!single && !(slots.length && slots.every(function(s) { return gridRows.indexOf(s.label) !== -1; }))) {
    ui.alert("This Round 2 form was made by an older version of Debate Tools. " +
      "Run '2. Generate Round 2 Form' again to make a new one.");
    return;
  }

  var topic = null;
  if (single) {
    var saved = [];
    try { saved = JSON.parse(props.getProperty(PROP_ROUND2_TOPICS) || '[]'); } catch (e) {}
    topic = saved[0] || savedWinners_()[0];
  }

  // --- FIND COLUMNS ---
  var headers = data[0];
  var slotCols = single ? [] : findTopicCols_(headers, slots.map(function(s) { return s.label; }));
  var slotColIdx = slotCols.map(function(c) { return c.index; });
  var nameCol = -1, stanceCol = -1;

  for (var i = 0; i < headers.length; i++) {
    if (slotColIdx.indexOf(i) !== -1) continue;
    var header = String(headers[i]).toLowerCase();
    if (nameCol === -1 && header.indexOf('name') !== -1) {
      nameCol = i;
    } else if (stanceCol === -1 && (header.indexOf('side') !== -1 || header.indexOf('preference') !== -1)) {
      stanceCol = i;
    }
  }

  var missingSlotCol = !single && slots.some(function(s) {
    return !slotCols.some(function(c) { return c.topic === s.label; });
  });
  if (nameCol === -1 || (single && (stanceCol === -1 || !topic)) || missingSlotCol) {
    ui.alert("Could not identify the name and preference columns on the Round 2 responses tab.");
    return;
  }

  // --- READ STUDENTS (latest submission per name wins) ---
  var byName = {};
  var dupes = 0;

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name = String(row[nameCol] || '').trim().replace(/\s+/g, ' ');
    if (!name) continue;

    var student = { name: name, assignedTopic: null, assignedStance: null };

    if (single) {
      student.stancePref = String(row[stanceCol] || '').trim();
    } else {
      var ranked = [];
      slotCols.forEach(function(c) {
        var rank = parseInt(row[c.index], 10);
        if (!isNaN(rank)) ranked.push({ label: c.topic, rank: rank });
      });
      ranked.sort(function(a, b) { return a.rank - b.rank; });
      var slotPrefs = [];
      ranked.forEach(function(p) { if (slotPrefs.indexOf(p.label) === -1) slotPrefs.push(p.label); });
      // Anything left unranked goes at the end so every student has a complete list
      slots.forEach(function(s) { if (slotPrefs.indexOf(s.label) === -1) slotPrefs.push(s.label); });
      student.slotPrefs = slotPrefs;
    }

    var key = name.toLowerCase();
    if (byName[key]) dupes++;
    byName[key] = student;
  }

  var students = Object.keys(byName).map(function(k) { return byName[k]; });
  if (students.length === 0) {
    ui.alert("No student names found on the Round 2 responses tab.");
    return;
  }

  // --- ASSIGN ---
  shuffle_(students);  // removes any advantage from submitting early
  if (single) {
    students.forEach(function(s) { s.assignedTopic = topic; });
    splitBySideGradient_(students);
  } else {
    assignSlots_(students, slots);
  }

  // Safety check for unassigned students
  var failed = students.filter(function(s) { return !s.assignedTopic || !s.assignedStance; });
  if (failed.length > 0) {
    ui.alert("Error: Algorithm failed to assign " + failed.length + " student(s). Check capacities.");
    return;
  }

  // --- REPORT + ROSTER ---
  var report = [["Match Satisfaction Report", "Count"]];
  var rosterHeader, outputData;

  if (single) {
    var yes = 0, no = 0, na = 0;
    students.forEach(function(s) {
      if (s.stanceGranted === 'Yes') yes++;
      else if (s.stanceGranted === 'No') no++;
      else na++;
    });
    report.push(
      ["Topic (one debate): " + topic, students.length],
      ["Stance: Preferred (Strongly or Normal)", yes],
      ["Stance: Forced to Switch", no],
      ["Stance: Had No Preference", na]);
    rosterHeader = ["Debate Topic", "Side", "Student Name", "Side Preference", "Stance Granted"];
    outputData = function(s) { return [s.assignedTopic, s.assignedStance, s.name, s.stancePref, s.stanceGranted]; };
  } else {
    var bySlot = {};
    slots.forEach(function(s) { bySlot[s.label] = s; });
    var choiceCounts = slots.map(function() { return 0; });
    var gotTopic = 0, gotSide = 0;
    students.forEach(function(s) {
      choiceCounts[s.choiceGranted - 1]++;
      if (bySlot[s.slotPrefs[0]].topic === s.assignedTopic) gotTopic++;
      // On their assigned topic, did they get the side they ranked higher?
      var onTopic = s.slotPrefs.filter(function(l) { return bySlot[l].topic === s.assignedTopic; });
      if (bySlot[onTopic[0]].side === s.assignedStance) gotSide++;
    });
    choiceCounts.forEach(function(n, k) { report.push(["Got their " + ordinal_(k + 1) + " choice", n]); });
    report.push(
      ["Got their 1st-choice topic", gotTopic],
      ["Got the side they ranked higher on their topic", gotSide]);
    rosterHeader = ["Debate Topic", "Side", "Student Name", "Choice Received", "Their 1st Choice"];
    outputData = function(s) { return [s.assignedTopic, s.assignedStance, s.name, s.choiceGranted, s.slotPrefs[0]]; };
  }

  // Roster order: topic, then Pro before Con, then name
  students.sort(function(a, b) {
    if (a.assignedTopic !== b.assignedTopic) return a.assignedTopic.localeCompare(b.assignedTopic);
    if (a.assignedStance !== b.assignedStance) return b.assignedStance.localeCompare(a.assignedStance);
    return a.name.localeCompare(b.name);
  });

  var rosterSheet = ss.insertSheet("Final Teams (" + stamp_(ss) + ")");
  rosterSheet.getRange(1, 1, report.length, 2).setValues(report);
  rosterSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#cfe2f3");

  // Roster starts one blank row below the report
  var headerRow = report.length + 2;
  rosterSheet.getRange(headerRow, 1, 1, 5).setValues([rosterHeader])
      .setFontWeight("bold").setBackground("#f3f3f3");
  var rows = students.map(outputData);
  rosterSheet.getRange(headerRow + 1, 1, rows.length, 5).setValues(rows);
  rosterSheet.autoResizeColumns(1, 5);
  rosterSheet.getRange(headerRow + 1, 4, rows.length, 2).setHorizontalAlignment("center");

  // Close Round 2 so late submissions don't arrive after teams are posted
  if (r2Form) r2Form.setAcceptingResponses(false);

  rosterSheet.activate();
  ui.alert("Assignments complete! " + students.length + " students placed." +
    (dupes ? "\n\n" + dupes + " duplicate submission(s) found; each student's latest one was used." : "") +
    "\n\nThe Round 2 form is now closed. To accept late responses, reopen it from the form's Responses tab and run step 3 again.");
}

/* ------------------------------------------------------------------ */
/* Matching helpers                                                    */
/* ------------------------------------------------------------------ */

function shuffle_(array) {
  for (var i = array.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = array[i]; array[i] = array[j]; array[j] = temp;
  }
}

function ordinal_(n) {
  return n + (n === 1 ? 'st' : n === 2 ? 'nd' : n === 3 ? 'rd' : 'th');
}

function stanceScore_(pref) {
  if (pref === "Strongly Prefer Pro") return 2;
  if (pref === "Prefer Pro") return 1;
  if (pref === "Strongly Prefer Con") return -2;
  if (pref === "Prefer Con") return -1;
  return 0;
}

/**
 * ONE DEBATE: split a (shuffled) team into Pro/Con with the 5-point gradient.
 * Sort from +2 to -2 and cut in the middle, so No Preference and mild
 * preferences fill out team sizes before any Strongly Prefer student moves.
 */
function splitBySideGradient_(team) {
  team.sort(function(a, b) { return stanceScore_(b.stancePref) - stanceScore_(a.stancePref); });

  var capPro = Math.floor(team.length / 2);
  if (team.length % 2 !== 0) {
    var proDemand = team.filter(function(s) { return stanceScore_(s.stancePref) > 0; }).length;
    var conDemand = team.filter(function(s) { return stanceScore_(s.stancePref) < 0; }).length;
    if (conDemand <= proDemand) capPro++;
  }

  team.forEach(function(s, i) {
    var score = stanceScore_(s.stancePref);
    if (i < capPro) {
      s.assignedStance = 'Pro';
      s.stanceGranted = score > 0 ? 'Yes' : (score === 0 ? 'N/A' : 'No');
    } else {
      s.assignedStance = 'Con';
      s.stanceGranted = score < 0 ? 'Yes' : (score === 0 ? 'N/A' : 'No');
    }
  });
}

/**
 * TWO DEBATES: assign (shuffled) students to topic/side slots from their rankings.
 * Each topic gets half the class and each side half its topic; odd extra seats
 * go to whichever topic or side the most students ranked 1st. Then pass 1 gives
 * everyone their 1st choice while seats last, pass 2 places the rest in their
 * 2nd choice, and so on. Every student ranks every slot and seats = students,
 * so everyone is placed by the last pass.
 */
function assignSlots_(students, slots) {
  var topics = [];
  slots.forEach(function(s) { if (topics.indexOf(s.topic) === -1) topics.push(s.topic); });

  var firstChoice = {};
  slots.forEach(function(s) { firstChoice[s.label] = 0; });
  students.forEach(function(st) { firstChoice[st.slotPrefs[0]]++; });

  var topicDemand = {};
  topics.forEach(function(t) { topicDemand[t] = 0; });
  slots.forEach(function(s) { topicDemand[s.topic] += firstChoice[s.label]; });

  // Topic sizes
  var topicCap = {};
  var base = Math.floor(students.length / topics.length);
  topics.forEach(function(t) { topicCap[t] = base; });
  topics.slice().sort(function(a, b) { return topicDemand[b] - topicDemand[a]; })
      .slice(0, students.length % topics.length)
      .forEach(function(t) { topicCap[t]++; });

  // Side sizes within each topic (ties go to Pro, which is listed first)
  var cap = {};
  topics.forEach(function(t) {
    var sides = slots.filter(function(s) { return s.topic === t; });
    var per = Math.floor(topicCap[t] / sides.length);
    sides.forEach(function(s) { cap[s.label] = per; });
    sides.slice().sort(function(a, b) { return firstChoice[b.label] - firstChoice[a.label]; })
        .slice(0, topicCap[t] % sides.length)
        .forEach(function(s) { cap[s.label]++; });
  });

  var bySlot = {}, filled = {};
  slots.forEach(function(s) { bySlot[s.label] = s; filled[s.label] = 0; });

  for (var pass = 0; pass < slots.length; pass++) {
    students.forEach(function(st) {
      if (st.assignedTopic) return;
      var label = st.slotPrefs[pass];
      if (filled[label] < cap[label]) {
        filled[label]++;
        st.assignedTopic = bySlot[label].topic;
        st.assignedStance = bySlot[label].side;
        st.choiceGranted = pass + 1;
      }
    });
  }
}
