// Google Apps Script code
function weeklyCaseDistribution() {
  var ss = SpreadsheetApp.openById('your-spreadsheet-id');
  var sheet1 = ss.getSheetByName('Copy of Apigee Q\'s salesforce');
  var sheet2 = ss.getSheetByName('Copy of WHP Apigee Queue');
  var rawSheet = ss.getSheetByName('Raw Data');

  var agents = rawSheet.getRange('G13:G42').getValues().flat();
  var smes = rawSheet.getRange('I17:I21').getValues().flat();
  var workingStatus = rawSheet.getRange('L17:L21').getValues().flat();
  var loadDistribution = rawSheet.getRange('M17:M21').getValues().flat();

  Logger.log('Agents: ' + agents);
  Logger.log('SMEs: ' + smes);
  Logger.log('Working Status: ' + workingStatus);
  Logger.log('Load Distribution: ' + loadDistribution);

  var cases1 = sheet1.getRange('A2:A' + sheet1.getLastRow()).getValues().flat();
  var cases2 = sheet2.getRange('B2:B' + sheet2.getLastRow()).getValues().flat().map(function(value) {
    var match = value.match(/"(\d+)"\)$/);
    return match ? match[1] : null;
  }).filter(function(value) { return value !== null; });

  Logger.log('Cases from sheet1: ' + cases1);
  Logger.log('Cases from sheet2: ' + cases2);

  var cases = cases1.concat(cases2).filter(function(caseNumber, index) {
    var agent = index < cases1.length ? sheet1.getRange(index + 2, 'B').getValue() : sheet2.getRange(index - cases1.length + 2, 'A').getValue();
    return agents.includes(agent);
  });

  Logger.log('Filtered cases: ' + cases);

  var totalCases = cases.length;
  var smeCases = {};

  smes.forEach(function(sme, index) {
    if (workingStatus[index]) {
      var load = loadDistribution[index] / 100;
      var numCases = Math.round(totalCases * load);
      smeCases[sme] = cases.splice(0, numCases);
    }
  });

  Logger.log('SME Cases: ' + JSON.stringify(smeCases));

  var overlap = Math.round(totalCases * 0.1);
  var overlapCases = cases.splice(0, overlap);

  Logger.log('Overlap cases: ' + overlapCases);

  Object.keys(smeCases).forEach(function(sme) {
    smeCases[sme] = smeCases[sme].concat(overlapCases);
    var email = sme + '@google.com';
    var subject = 'Weekly Case Review';
    var body = 'Here are the cases for you to review this week:\n\n' + smeCases[sme].join('\n');
    MailApp.sendEmail(email, subject, body);
  });
}

// Set up a weekly trigger
function createTrigger() {
  ScriptApp.newTrigger('weeklyCaseDistribution')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
}
