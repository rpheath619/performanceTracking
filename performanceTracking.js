var sheet = SpreadsheetApp.openById("1vcEs9uALlHWG3uKs0Wyq-4qfBwUHpO2ir_yiP-J3J8Y");
var issueSheet = sheet.getSheetByName("Issue Information");
var occurrenceSheet = sheet.getSheetByName("Issue Occurrences");
var issueData = issueSheet.getDataRange();
var issueDataValues = issueData.getValues();
var occurrenceData = occurrenceSheet.getDataRange();
var occurrenceDataValues = occurrenceData.getValues();

var issueHeadingIndices = [];
var occurrenceHeadingIndices = [];

function getId() {
  Browser.msgBox('Spreadsheet key: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
}

function createScriptTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length - 1; i >= 0; i--) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger("formSubmitted")
    .forSpreadsheet(sheet)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("edited")
    .forSpreadsheet(sheet)
    .onEdit()
    .create();

  ScriptApp.newTrigger("opened")
    .forSpreadsheet(sheet)
    .onOpen()
    .create();
}

function formSubmitted(formData) {
  identifyIssueHeadingIndices();
  createIssueEntry(formData.values);
}

function buttonPressed() {
  identifyIssueHeadingIndices();
  identifyOccurrenceHeadingIndices();
  var matchId = Browser.inputBox("Enter the Match ID:",Browser.Buttons.OK_CANCEL);
  if (matchId == "cancel") return;
  var allPlayers = Browser.msgBox("Are all players participating in the match?",Browser.Buttons.YES_NO_CANCEL);
  if (allPlayers == "cancel") return;
  if (allPlayers == "yes") {
    var matchParticipants = [];
    for (var i = issueDataValues.length - 1; i >= 1; i--) {
      var nextIssuePlayer = issueDataValues[i][issueHeadingIndices["Player"]];
      if (matchParticipants.indexOf(nextIssuePlayer) == -1) {
        matchParticipants.push(nextIssuePlayer);
      }
    }
  } else {
    var matchParticipants = Browser.inputBox("Enter the names of all players participating in the match, separated only by commas:",Browser.Buttons.OK_CANCEL);
    if (matchParticipants == "cancel") return;
    matchParticipants = matchParticipants.split(",");
  }
  createBlankOccurrenceEntries(matchId,matchParticipants);
}

function edited(eventData) {
  identifyIssueHeadingIndices();
  identifyOccurrenceHeadingIndices();
  var eventRange = eventData.range;
  var eventColumnIndex = eventRange.getColumn() - 1;
  var eventRowIndex = eventRange.getRow() - 1;
  if (eventRange.getSheet().getSheetId() == issueSheet.getSheetId() && eventRowIndex > 0) {
    var eventIssueId = parseInt(issueDataValues[eventRowIndex][issueHeadingIndices["Issue ID"]]);
    if (eventColumnIndex == issueHeadingIndices["Impact Rating"] || issueHeadingIndices["Recurrence Rate"]) {
      calculateIssuePriority(eventIssueId);
      writeIssueData();
    }
  } else if (eventRange.getSheet().getSheetId() == occurrenceSheet.getSheetId() && eventRowIndex > 0) {
    var eventIssueId = parseInt(occurrenceDataValues[eventRowIndex][occurrenceHeadingIndices["Issue ID"]]);
    if (eventColumnIndex == occurrenceHeadingIndices["Success Count"] || occurrenceHeadingIndices["Fail Count"] || occurrenceHeadingIndices["Issue ID"]) {
      calculateIssueOccurrenceData(eventIssueId);
      calculateIssueProgress(eventIssueId);
      calculateIssueStatus(eventIssueId);
      calculateIssuePriority(eventIssueId);
      writeIssueData();
    }
  }
}


function opened() {
  identifyIssueHeadingIndices();
  identifyOccurrenceHeadingIndices();
  var uniqueIssueIds = [];
  for (var i = occurrenceDataValues.length - 1; i >= 1; i--) {
    var eventIssueId = parseInt(occurrenceDataValues[i][occurrenceHeadingIndices["Issue ID"]]);
    if (uniqueIssueIds.indexOf(eventIssueId) == -1) {
      uniqueIssueIds.push(parseInt(eventIssueId));
    }
  }
  for (var i = uniqueIssueIds.length - 1; i >= 0; i--) {
    calculateIssueOccurrenceData(uniqueIssueIds[i]);
    calculateIssueProgress(uniqueIssueIds[i]);
    calculateIssueStatus(uniqueIssueIds[i]);
    calculateIssuePriority(uniqueIssueIds[i]);
    writeIssueData();
  }
}

function calculateIssuePriority(issueId) {
  var issueRow = getIssueRowById(issueId);
  var issueImpactRating = issueDataValues[issueRow][issueHeadingIndices["Impact Rating"]];
  var issueRecurrenceRate = issueDataValues[issueRow][issueHeadingIndices["Recurrence Rate"]];
  if (issueImpactRating == "" || issueRecurrenceRate == "" || issueRecurrenceRate == "TBD") {
    setIssueRangeByHeaderAndId(issueId,"Priority","TBD");
  } else {
    var issuePriority = (parseFloat(issueImpactRating) + parseFloat(issueRecurrenceRate))/2;
    setIssueRangeByHeaderAndId(issueId,"Priority",issuePriority.toString());
  }
}

function calculateIssueProgress(issueId) {
  var issueRow = getIssueRowById(issueId);
  var issueActualOccurrences = parseFloat(issueDataValues[issueRow][issueHeadingIndices["Actual Occurrences"]]);
  var issueTargetOccurrences = parseFloat(issueDataValues[issueRow][issueHeadingIndices["Target Occurrences"]]);
  var issueActualSuccessRate = parseFloat(issueDataValues[issueRow][issueHeadingIndices["Actual Success Rate"]]);
  var issueTargetSuccessRate = parseFloat(issueDataValues[issueRow][issueHeadingIndices["Target Success Rate"]]);
  var issueProgress = parseInt(issueActualOccurrences) == 0 ? "TBD" : (Math.min(issueTargetOccurrences,issueActualOccurrences)/issueTargetOccurrences*Math.min(issueTargetSuccessRate,issueActualSuccessRate)/issueTargetSuccessRate*100).toString()+"%";
  setIssueRangeByHeaderAndId(issueId,"Progress",issueProgress);
}

function calculateIssueStatus(issueId) {
  var issueRow = getIssueRowById(issueId);
  var issueProgress = parseFloat(issueDataValues[issueRow][issueHeadingIndices["Progress"]]);
  var issueRecurrenceRate = issueDataValues[issueRow][issueHeadingIndices["Recurrence Rate"]];
  var issueStatus;
  if (issueRecurrenceRate == "TBD") {
    issueStatus = "TBD";
  } else if (issueProgress < 100) {
    issueStatus = "Improve";
  } else if (issueStatus == "") {
  	// decide on a suspension phrase and leave the cell alone if it's present
  } else {
    issueStatus = "Maintain";
  }
  setIssueRangeByHeaderAndId(issueId,"Status",issueStatus);
}

function calculateIssueOccurrenceData(issueId) {
  var issueOccurrences = filterOccurrenceData(occurrenceDataValues,"Issue ID",issueId);
  var totalIssueSuccessCount = 0;
  var totalIssueFailCount = 0;
  var issueMatches = issueOccurrences.length;
  for (var i = issueOccurrences.length - 1; i >= 0; i--) {
    totalIssueSuccessCount += (parseInt(issueOccurrences[i][occurrenceHeadingIndices["Success Count"]]) || 0);
    totalIssueFailCount += (parseInt(issueOccurrences[i][occurrenceHeadingIndices["Fail Count"]]) || 0);
  }
  var issueActualOccurrences = totalIssueSuccessCount + totalIssueFailCount;
  var issueOccurrencesPerMatch = issueActualOccurrences/issueMatches
  var issueRecurrenceRateRaw = issueOccurrencesPerMatch*(7 - 1)/(3 - 0.166666667) + 1;
  var issueRecurrenceRate = issueActualOccurrences == 0 ? "TBD" : Math.min(Math.max(issueRecurrenceRateRaw,1),7);
  var issueActualSuccessRate = issueActualOccurrences == 0 ? "TBD" : (totalIssueSuccessCount/issueActualOccurrences*100).toString()+"%";
  setIssueRangeByHeaderAndId(issueId,"Actual Occurrences",issueActualOccurrences);
  setIssueRangeByHeaderAndId(issueId,"Recurrence Rate",issueRecurrenceRate);
  setIssueRangeByHeaderAndId(issueId,"Actual Success Rate",issueActualSuccessRate);
}

function getIssueRowById(issueId) {
  for (var i = issueDataValues.length - 1; i >= 1; i--) {
    if (issueDataValues[i][issueHeadingIndices["Issue ID"]] == issueId) {
      return i;
    }
  }
}

function setIssueRangeByHeaderAndId(issueId,columnHeader,valueToWrite) {
  var issueRow = getIssueRowById(issueId);
  issueDataValues[issueRow][issueHeadingIndices[columnHeader]] = valueToWrite;
}

function writeIssueData() {
  issueData.setValues(issueDataValues);
}

function createIssueEntry(formValues) {
  identifyIssueHeadingIndices();
  var issueDataHeight = issueDataValues.length;
  var formColumnHeaders = ["Submission Timestamp","Player","Description","Impact Rating"]
  var maxId = 0;
  for (var i = issueDataHeight - 1; i >= 1; i--) {
    var nextId = issueDataValues[i][issueHeadingIndices["Issue ID"]];
    maxId = nextId > maxId ? nextId : maxId;
  }
  maxId++;
  maxId = maxId.toString()
  for (var i = maxId.length; i < 10; i++) {
    maxId = "0" + maxId;
  }
  var newIssueRow = [[]];
  for (var i = issueDataValues[0].length - 1; i >= 0; i--) {
    newIssueRow[0].push("");
  }
  for (var i = formValues.length - 1; i >= 0; i--) {
    newIssueRow[0][issueHeadingIndices[formColumnHeaders[i]]] = formValues[i];
  }
  newIssueRow[0][issueHeadingIndices["Target Occurrences"]] = "10";
  newIssueRow[0][issueHeadingIndices["Target Success Rate"]] = "80%";
  newIssueRow[0][issueHeadingIndices["Issue ID"]] = maxId;
  for (var i = newIssueRow[0].length - 1; i >= 9; i--) {
    newIssueRow[0][i] = newIssueRow[0][i] == "" ? "TBD" : newIssueRow[0][i];
  }
  var newIssueRange = issueSheet.getRange(issueDataHeight + 1,1,1,issueDataValues[0].length);
  newIssueRange.setValues(newIssueRow);
}

function createBlankOccurrenceEntries(matchId,matchParticipants) {
  var newOccurrenceRows = [];
  var playerIssueData = [];
  var playerImproveIssues = [];
  var playerTbdIssues = [];
  var applicableIssueData = [];
  for (var i = matchParticipants.length - 1; i >= 0; i--) {
    playerIssueData = playerIssueData.concat(filterIssueData(issueDataValues,"Player",matchParticipants[i]));
  }
  if (playerIssueData.length > 0) {
    playerImproveIssues = filterIssueData(playerIssueData,"Status","Improve");
    playerTbdIssues = filterIssueData(playerIssueData,"Status","TBD");
  }
  applicableIssueData = applicableIssueData.concat(playerImproveIssues);
  applicableIssueData = applicableIssueData.concat(playerTbdIssues);
  for (var i = applicableIssueData.length - 1; i >= 0; i--) {
    newOccurrenceRows[i] = [];
    newOccurrenceRows[i][0] = matchId;
    var timestamp = new Date();
    newOccurrenceRows[i][1] = timestamp.toString()
    for (var n = occurrenceDataValues[0].length - 1; n >= 2; n--) {
      if (issueDataValues[0].indexOf(occurrenceDataValues[0][n]) > -1) {
        newOccurrenceRows[i][n] = applicableIssueData[i][issueHeadingIndices[occurrenceDataValues[0][n]]];
      } else {
        newOccurrenceRows[i][n] = "";
      }
    }
  }
  if (applicableIssueData.length > 0) {
    var newDataRange = occurrenceSheet.insertRowsBefore(2,newOccurrenceRows.length).getRange(2,1,newOccurrenceRows.length,newOccurrenceRows[0].length);
    newDataRange.setValues(newOccurrenceRows);
  }
}

function identifyIssueHeadingIndices() {
  for (var i = issueDataValues[0].length - 1; i >= 0; i--) {
    issueHeadingIndices[issueDataValues[0][i]] = i;
  }
}

function identifyOccurrenceHeadingIndices() {
  for (var i = occurrenceDataValues[0].length - 1; i >= 0; i--) {
    occurrenceHeadingIndices[occurrenceDataValues[0][i]] = i;
  }
}

function filterIssueData(issueData,columnHeader,filterValue) {
  identifyIssueHeadingIndices();
  var filteredArray = [];
  for (var i = issueData.length - 1; i >= 0; i--) {
    if (issueData[i][issueHeadingIndices[columnHeader]] == filterValue) {
      filteredArray.push(issueData[i]);
    }
  }
  return filteredArray;
}

function filterOccurrenceData(occurrenceData,columnHeader,filterValue) {
  identifyOccurrenceHeadingIndices();
  var filteredArray = [];
  for (var i = occurrenceData.length - 1; i >= 0; i--) {
    if (occurrenceData[i][occurrenceHeadingIndices[columnHeader]] == filterValue) {
      filteredArray.push(occurrenceData[i]);
    }
  }
  return filteredArray;
}

var sortTargetColumnIndex;

function test() {
  identifyOccurrenceHeadingIndices();
  sortTargetColumnIndex = occurrenceHeadingIndices["Issue ID"];
  var sortedArray = occurrenceDataValues.slice(1,occurrenceDataValues.length).sort(sortRowsByColumn);
}

function sortRowsByColumn(a,b) {
  if ((a[sortTargetColumnIndex] - b[sortTargetColumnIndex]) > 0) return 1;
  if ((a[sortTargetColumnIndex] - b[sortTargetColumnIndex]) < 0) return -1;
  return 0;
}