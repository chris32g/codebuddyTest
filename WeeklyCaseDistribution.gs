// Google Apps Script code
function weeklyCaseDistribution() {
  var ss = SpreadsheetApp.openById('1rYjlZrB1MI4mMHqwOiPecvv0caum-BjOQ-EBAKBEW8s');
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
  var cases2 = sheet2.getRange('B2:B' + sheet2.getLastRow()).getValues().flat();

  Logger.log('Cases from sheet1: ' + cases1);
  Logger.log('Cases from sheet2: ' + cases2);

 var cases = [];
 for (var i = 2; i <= sheet1.getLastRow(); i++) {
   if (agents.includes(sheet1.getRange(i, 2).getValue())) {
     cases.push(sheet1.getRange(i, 1).getValue());
   }
 }
 for (var i = 2; i <= sheet2.getLastRow(); i++) {
   if (agents.includes(sheet2.getRange(i, 1).getValue())) {
     var value = sheet2.getRange(i, 2).getValue();
     var match = value.match(/"(\d+)"\)$/);
     if (match) {
       cases.push(match[1]);
     }
   }
 }

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
    var email = 'gonchristian' + '@google.com';
    var subject = 'Weekly Case Review ' + sme;
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
