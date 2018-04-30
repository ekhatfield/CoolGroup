/**
 * Lizzy Hatfield
 * CoolGroup Newsletter Generator
 */


/*
issues: 
-change layout of final html doc to test.html
--read announcement from doc file
--read comic url from doc file
-manage data in the spreadsheet
*/

//main function to execute script
function main() {
var SPREADSHEET_ID = '1nFi6f98GuG_46Xmg3V8WPoY0rtMK4LykOIW5QilFRiY'
var RANGE = 'Form Responses 1!B2:G20'

var WeekNumber = getWeek()
  
var data = getSpreadsheetData(RANGE, SPREADSHEET_ID)

var groupTables = []
for (i in data) {
  var row = data[i]
  if (row[0])
     {groupTables.push(buildEmailRow(row))}
}

var email = buildEmail(groupTables)

saveEmail(email)

GmailApp.sendEmail("email@student.tudelft.nl", "Test Email to Outlook", "testing", {htmlBody: email});
  
}


//helper functions
function getWeek() {
  var d = new Date();
  d.setHours(0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  return Math.ceil((((d - new Date(d.getFullYear(), 0, 1)) / 8.64e7) + 1) / 7);//-1
}

function getSpreadsheetData(range, id) {
  var sheet = SpreadsheetApp.openById(id)//here

  var dataRange = sheet.getRange(range)

  return dataRange.getValues()
}

function buildEmailRow(row) {
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th>' + row[0] + ', ' + row[1] + ' Student</th>',
    '<th>Project Name: ' + row[2] + '</th>',
    '</tr>',
    '<tr style = "background-color: #eee;">',
    '<td><b>This Week\'s Achievement</b></td>',
    '<td width=\"67%\">' + row[3] + '</td>',
    '</tr>',
    '<tr style = "background-color: #fff;">',
    '<td><b>This Week\'s Problem(s)</b></td>',
    '<td>' + row[4] + '</td>',
    '</tr>',
    '<tr style = "background-color: #eee;">',
    '<td><b>Next Week\'s Plan</b></td>',
    '<td>' + row[5] + '</td>',
    '</tr>',
    '</table>'
  ].join('\n')
}

function buildEmailHead() {
  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '<style>',
    'table {width: 800px;margin: 20px;}',
    'table, th, td {border-collapse: collapse;}',
    'th, td {padding: 12px; text-align: left;}',
    'table#t01 tr:nth-child(even) {background-color: #eee;}',
    'table#t01 tr:nth-child(odd) {background-color: #fff;}',
    'table#t01 th {background-color: #00A6D6; color: white;}',
    '</style>',
    '</head>'
  ].join('\n')
}

function buildEmail(groupTables) {
  return [
    buildEmailHead(),
    '<body>',
    '<h2 style="font-family:arial;">CoolGroup Newsletter Week ' + getWeek() + '</h2>',
    groupTables.join('\n'),
    '</body>',
    '</html>'
  ].join('\n')
}

function saveEmail(email) {
  //var folder = DriveApp.getRootFolder()
  var CGfolderID = '16bmSg6V75RwXc6dWlQpWBCxgBWoJ92Fp'
  folder = DriveApp.getFolderById(CGfolderID)
  
  folder.createFile('newsletter.html', email, MimeType.HTML)
}
