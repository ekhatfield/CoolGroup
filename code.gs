/************************************************************************************************************
 * Engineer: Lizzy Hatfield 
 * Name: CoolGroup Newsletter Generator
 * Description: Script that generates and sends weekly newsletter emails via Google Forms, Sheets, and Docs
 * Ver: 1.1
 * © Lizzy Hatfield
************************************************************************************************************/


/* ----------------issues: 
-html input in announcement doc?
-check for end of entries
-error/malice-checking form sheet
-clean up and comment code
----------------------- */


/*-----------------Documentation
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                          .".
                         /  |
                        /  /
                       / ,"
           .-------.--- /
          "._ __.-/ o. o\  
             "   (    Y  )
                  )     /
                 /     (
                /       Y
            .-"         |
           /  _     \    \ 
          /    `. ". ) /' )
         Y       )( / /(,/
        ,|      /     )
       ( |     /     /
        " \_  (__   (__        [nabis]
            "-._,)--._,)

© Chris Johnson, 
http://www.chris.com/ascii/index.php?art=animals/rabbits
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


New in Version 1.1:
++sort data before creating newsletter
++function to define the depth of the ranges
++add copyright
++make announcement/comic optional tables?
++save HTML docs in a separate folder
-+clearData range as input

for the future when you want to read announcements.doc as html:
   https://stackoverflow.com/questions/47299478/google-apps-script-docs-convert-selected-element-to-html/47313357#47313357

-a big THANK YOU to Jarva for translating the alpha version of this algorithm from Python to JavaScript/Google Script (and writing the getWeek() function!)
------------------------------*/



/*==============================================Main Function 

This function exists to execute the functionality of the script by calling the helper functions defined below.
Essentially it ties everything together.

===========================================================*/
function main() {
//start timing
var d0 = new Date();
var t0 = d0.getTime().toFixed(0);
//define IDs and ranges for the Google Docs and Sheets
var form_sheet_id = '1nFi6f98GuG_46Xmg3V8WPoY0rtMK4LykOIW5QilFRiY';
var archive_sheet_id = '1PkE5F7MWh-nS67w-PKAxgaYXsdleA_mdD_QpZUI8aAQ';
var announcement_doc_id = '12qAcj4a-SQ-1Nwd9QyQeCYToXA-objeBrvE58iUg8j8';
var comic_doc_id = '1JF-EUh_yNrbaTyKfwY-1t12MvqKgtjEu8sjq8ZHkS2o';
var work_folder = '16bmSg6V75RwXc6dWlQpWBCxgBWoJ92Fp';
var HTML_folder = '1ewwc7cRepBUsKh-DGxD9rIwir5LC1WzM';
var range_depth = 20;
var newsletter_range = 'Form Responses 1!B2:G' + range_depth;
var full_data_range = 'Form Responses 1!A2:G' + range_depth;
//set flags to turn on and off optional features
var AddAnn = true;
var AddComic = true;
var doSort = false;
//pull data needed to create newlsetter
var WeekNumber = getWeek();
var ann = retrieveAnnouncement(announcement_doc_id);
var com = retrieveComic(comic_doc_id);
if (doSort)
{
 alphabetizeSheet(form_sheet_id, newsletter_range);
}
var data = getSpreadsheetData(newsletter_range, form_sheet_id);
//create an array with the tables of the weelky reports
var groupTables = [];
for (i in data) {
  var row = data[i];
  if (row[0])//check that there is contents in this row via the first element, which is 'Name'
     {
       groupTables.push(buildEmailRow(row));
     }
}
//create the HTML file and send the email
var email = buildEmail(AddAnn, AddComic, groupTables, announcement_doc_id, comic_doc_id);
saveEmail(email, HTML_folder);//comment this if you dont want the output HTML file
GmailApp.sendEmail("email@gmail.com", "Newsletter Week " + getWeek(), "testing", {htmlBody: email});
//copy the data from the Forms Sheet to the Archive Sheet
updateArchive(full_data_range, form_sheet_id, archive_sheet_id);
//remove the data from the Form Sheet
clearData(form_sheet_id, range_depth);
//end timing and print to log
var d1 = new Date();
var t1 = d1.getTime().toFixed(0);
Logger.log("Script took " + (t1 - t0)/1000 + " seconds to execute.");
}//end main



/*====================================Helper Functions

These functions perform the heavy-lifting of the script such as reading from the Google Sheets/Docs, compiling the newsletter HTML file, etc

====================================================*/


/*------------------------------getWeek()

getWeek() calculates the week number using Date() and Math functions, which is returned as an integer

-credits to Jarva for writing this function
------------------------------*/
function getWeek() {
  var d = new Date();
  d.setHours(0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  return Math.ceil((((d - new Date(d.getFullYear(), 0, 1)) / 8.64e7) + 1) / 7);
}


/*-----------------------------------------------------------retrieveAnnouncement()

Grabs the announcement from the Announcement Google Doc as a string

-DocumentApp.openById(id) does not allow passing id as an argument, thus global variables are used
-------------------------------------------------------------*/
function retrieveAnnouncement(doc_id){
  var announcement = DocumentApp.openById(doc_id).getBody().getText();
  return announcement
}


/*-----------------------------------------------------------retrieveComic()

Grabs the comic URL from the Comic Google Doc as a string

-DocumentApp.openById(id) does not allow passing id as an argument, thus global variables are used
-------------------------------------------------------------*/
function retrieveComic(doc_id){
  var comic = DocumentApp.openById(doc_id).getBody().getText();
  return comic
}


/*--------------------------------getSpreadsheetData() 

getSpreadsheetData() returns the values from the specified range of the specified sheet as a [string?] array

----------------------------------------------------*/
function getSpreadsheetData(range, sheet_id) {
  var sheet = SpreadsheetApp.openById(sheet_id);
  var data_range = sheet.getRange(range);
  return data_range.getValues()
}


/*-------------------------------------------------------buildEmailRow()

buildEmailRow() writes the tables that contain the weekly report data in HTML as a string array

----------------------------------------------------------------------*/
function buildEmailRow(row) {
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th>' + row[0] + ', ' + row[1] + '</th>',
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


/*---------------------------------------------------------buildEmailHead()

Writes the header for the HTML file, which includes the CSS definitions for the tables/overall layout of the file, as a string array

-add "'th {font-size:125%}'," if destination is gmail
------------------------------------------------------------------------*/
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


/*---------------------------------------------------------buildEmailAnn()

Writes the header for the HTML file, which includes the CSS definitions for the tables/overall layout of the file, as a string array

-add "'th {font-size:125%}'," if destination is gmail
------------------------------------------------------------------------*/
function buildEmailAnn(ann_doc_id) {
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th style="text-align:center">Announcements</th>',
    '</tr>',
    '<tr style = "background-color: #eee;">',
    '<td>' + retrieveAnnouncement(ann_doc_id) + '</td>',
    '</tr>',
    '</table>'
  ].join('\n')
}

/*---------------------------------------------------------buildEmailComic()

Writes the header for the HTML file, which includes the CSS definitions for the tables/overall layout of the file, as a string array

-add "'th {font-size:125%}'," if destination is gmail
------------------------------------------------------------------------*/
function buildEmailComic(com_doc_id) {
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th style="text-align:center">Comic of the Week</th>',
    '</tr>',
    '<tr style = "background-color: #eee;">',
    '<td style="text-align:center"><img src="' + retrieveComic(com_doc_id) + '" alt="You win this round, [email client name]" width="776"></td>',
    '</tr>',
    '</table>'
  ].join('\n')
}

/*---------------------------------------------------------buildEmail()

Writes all of the HTML for the email as a string array

---------------------------------------------------------------------*/
function buildEmail(ann, comic, groupTables, ann_doc_id, com_doc_id) {
  return [
    buildEmailHead(),
    '<body>',
    '<h2 style="font-family:arial;">CoolGroup Newsletter Week ' + getWeek() + '</h2>',
    ann ? buildEmailAnn(ann_doc_id) : ' ',
    groupTables.join('\n'),    
    comic ? buildEmailComic(com_doc_id) : ' ',
    '<h5 style="font-family:arial;">Software Copyright Lizzy Hatfield 2018 </h5>',
    '</body>',
    '</html>'
  ].join('\n')
}


/*-------------------------------------------------------------------------saveEmail()

Creates the physical HTML file out of the HTML code that was generated before

------------------------------------------------------------------------------------*/
function saveEmail(email, dest_id) {
  var folder = DriveApp.getFolderById(dest_id) 
  var d = new Date();     
  folder.createFile('newsletter-wk' + getWeek() + '-' + d.getFullYear() +'.html', email, MimeType.HTML)
}


/*-----------------------------------------------------------updateArchive()

Appends a new row of data to the Archive Sheet

-------------------------------------------------------------*/
function updateArchive(source_range, source_id, dest_id) {
  var dest_sheet = SpreadsheetApp.openById(dest_id);
  var source_data = getSpreadsheetData(source_range, source_id);
  for (i in source_data) {
    var row = source_data[i];
    if (row[0])//if row[0] == anything
    {
      dest_sheet.appendRow(row) 
    }
  }   
}


/*--------------------------------------------------------------------clearData()

Deletes the rows of the Form Sheet in ascending order

-if the function is reversed, only every other row is deleted, as the N+1 row fills the Nth row once it is deleted, then the algorithm deletes the contents of row N+1, etc
-------------------------------------------------------------------------------*/
function clearData(id, depth) {
  var sheet = SpreadsheetApp.openById(id)
  var i;
  for (i = depth ; i > 1 ; i--) 
  {
    sheet.deleteRow(i);
  }
}

/*--------------------------------------------------------------------alphabetizeSheet()

Sorts the rows of the Form Sheet alpabetically by name

-------------------------------------------------------------------------------*/
function alphabetizeSheet(sheet_id, sort_range){
  var data = SpreadsheetApp.openById(sheet_id).getRange(sort_range);
  data.sort(2);  
}
