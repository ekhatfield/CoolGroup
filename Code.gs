//https://script.google.com/macros/s/AKfycbwRYsn2YNF_HyH0vGm1QlIHE-0Fnw1I278f2qVlxTs/dev
//https://script.google.com/macros/s/AKfycbwHahwRXvXbtNFujh45oSsdyqw7YyImQ1BcYVxYae4WKtaUj90/exec


/************************************************************************************************************
 * Engineer: Lizzy Hatfield (4625129)
 * Name: CoolGroup Newsletter Generator
 * Description: Script that generates and sends weekly newsletter emails via Google Forms, Sheets, and Docs
 * Ver: 3.0
 * © Lizzy Hatfield
************************************************************************************************************/


/* ----------------------issues/new ideas:
-triggered reminder email
-error-checking form sheet
->multiple entries from one person
->joke entries
-comic sheet management
-report sheet management
-set up pinakin with the gui
--------------------------------------- */


/*---------------------------------------------------------------------------Documentation-----------------------------------------------------------------------------
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
++reports can be sorted alphabetically by name
++depth of the ranges can be programmicably defined
++added copyright
++announcement/comic are now optional
++HTML saved in a separate folder
++clearData takes range as an input

New in Version 1.2:
++function testEmail() to test the output email
++function testAnnouncement() to test announcement for correct formatting added from Test-Funcs
++function testThisHTML() to test new HTML formatting file for email compatability added from Test-Funcs
++function sendReminder() to send a reminder email

New in Version 2.0:
++added GUI for running the script to control all the bells and whistles more cleanly
++main renamed to sendNewsletter

New in Version 3.0:
++bug in clearing old submission data from form spreadsheet fixed
++email sent to developer on error in GUI
++announcements and reminder read from Docs text to HTML directly
++comics can be added by anyone and are chosen at random by the script
++malice checking of progress submissions (</td>)
++sendNewsletter and testEmail called functions bundled to keep consistency between the two
++code is now fully portable to any group/lab/university/etc


for the future when you want to read announcements.doc as html:
   https://stackoverflow.com/questions/47299478/google-apps-script-docs-convert-selected-element-to-html/47313357#47313357

-a big THANK YOU to Jarva for translating the alpha version of this algorithm from Python to JavaScript/Google Script (and writing the getWeek() function!)
-------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

//front end code


function doGet() {
  var html = HtmlService.createTemplateFromFile('Webpage').evaluate();
  html.setTitle("coolGroup Newsletter");
  return html; 
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getContent();
}


//backend code


/*==================================================================Global Variables==================================================================================

All of the global variables are defined here. These are the pieces of information that are the same for every function, although most functions are written to use
these values as if they were local variables declared in the main functions.

====================================================================================================================================================================*/
var form_sheet_id = '1nFi6f98GuG_46Xmg3V8WPoY0rtMK4LykOIW5QilFRiY';
var archive_sheet_id = '1PkE5F7MWh-nS67w-PKAxgaYXsdleA_mdD_QpZUI8aAQ';
var announcement_doc_id = '12qAcj4a-SQ-1Nwd9QyQeCYToXA-objeBrvE58iUg8j8';
var work_folder = '16bmSg6V75RwXc6dWlQpWBCxgBWoJ92Fp';
var HTML_folder = '1ewwc7cRepBUsKh-DGxD9rIwir5LC1WzM';
var comic_sheet_id = '1xt5fVg-G1on9X6igIvaILBIdB_wKg7GFd9Bp0aAURMc';
var reminder_doc_id = '1R4l1rokHQLdFt2LmVRIlty-Tew9W6OOPptG3vgVuQcg';
var archiveEmails = true;

var devEmail = 'yoyoultimate@gmail.com';
var userEmail = 'fabio.sebastiano@gmail.com';
var labEmail = 'coolgroup-ewi@tudelft.nl';

var labName = 'CoolGroup';
var headerBGColor = '#00A6D6'; //cyan
var headerTextColor = 'white';
var accentColor1 = '#eee'; //light gray
var accentColor2 = '#fff'; //white



/*==================================================================Main Functions====================================================================================

These are the top-level functions that execute all the tasks needed respectively. They are directly called by the GUI

-rewrite with functions bundled to keep consistency between testEmail and sendNewsletter
====================================================================================================================================================================*/
function sendNewsletter(doSort, AddAnn, AddComic) {

var subject = "Newsletter Week " + getWeek();
var body = createEmail(true, true, doSort);
var recipient = labEmail;

sendAndSaveEmail(subject, recipient, body, archiveEmails);
cleanup();
}//end 


/*-------------------------------------------------------------------------testEmail()------------------------------------------------------------------------------------

Description: Sends a test email to a single recipient of the generated newsletter

Arguments:
-none
Returns:
-none
Notes:
-allows for the creation and sending of an email in the same way as main(), but without performing the cleanup functions main() does such as saving items to the Archive
-should be rewritten in the future to run the same code as main() so that the functions stay coherent (and leave main() to run some cleanup() procedures)
------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function testEmail(doSort, AddAnn, AddComic) {

var subject = "Test Newsletter Week " + getWeek();
var body = createEmail(AddAnn, AddComic, doSort);
var recipient = userEmail;

sendAndSaveEmail(subject, recipient, body, false);
}//end


/*-------------------------------------------------------------------------sendReminder()----------------------------------------------------------------------------------

Description: Sends a reminder email to the CoolGroup to fill in the Google Form

Arguments:
-none
Returns:
-none
Notes:
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function sendReminder() {
//create the HTML file and send the email
var email = retrieveAnnouncement(reminder_doc_id);
//send email
GmailApp.sendEmail(labEmail, "Newsletter Submission Reminder", "Error Sending Email", {htmlBody: email});
}



/*======================================================================Utility Functions================================================================================

These functions are here to test functionalities without sending the newsletter to the entire coolgroup. They are not called anywhere, but rather are meant to be called
by the user from the Google Script Development environment.

=======================================================================================================================================================================*/


/*-----------------------------------------------------------------------testAnnouncement()------------------------------------------------------------------------------

Description: Sends an email with only the Announcement box from the Newsletter to test the HTML formatting of the announcement itself

Arguments:
-none
Returns:
-none
Notes:
-needs to be rewritten so it calls the same functions used 
-should no longer be self-contained
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function testAnnouncement(){
 //open announcement doc
 var docID = '12b9LNsjeh6ac3qlhRHTjvDQeNBgbsBVlhkbHlIk5NCU'; 
 var announcement = [];
 announcement = convertToHTML(docID);

 //create email
 var email = [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '<style>',
    'table {width: 800px;margin: 20px;}',
    'table, th, td {border-collapse: collapse;}',
    'th, td {padding: 12px; text-align: left;}',
    'table#t01 tr:nth-child(even) {background-color: ' + accentColor1 + ';}',
    'table#t01 tr:nth-child(odd) {background-color: ' + accentColor2 + ';}',
    'table#t01 th {background-color: #00A6D6; color: white;}',
    '</style>',
    '</head>',
    '<body>',
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th style="text-align:center">Announcements</th>',
    '</tr>',
    '<tr style = "background-color: ' + accentColor1 + ';">',
    '<td>' + announcement + '</td>',
    '</tr>',
    '</table>',
    '</body>',
    '</html>'
  ].join('\n') 
 //send email
 GmailApp.sendEmail(userEmail, "Announcement Test", "error", {htmlBody: email});
}


/*-------------------------------------------------------------------------testThisHTML()----------------------------------------------------------------------------------

Description: Reads from a test format HTML file and sends it as an email to the recipient to test if the format is email-compatible

Arguments:
-none
Returns:
-none
Notes:
-could be linearized (one line)
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
/* use this guy to test a new html style file for email compatability */
function testThisHTML(){
 var files = DriveApp.getFilesByName('new-style.html');
 var file = files.next();
 var html = file.getAs('text/html');
 var email = html.getDataAsString()
 GmailApp.sendEmail(userEmail, "HTML Test", "error", {htmlBody: email});
}


/*=======================================================================Bundle Functions=================================================================================

These functions execute the functions to perform general groups of tasks. This aliasing is done to keep testEmail and sendAnnouncement congruent

========================================================================================================================================================================*/


function createEmail(AddAnn, AddComic, doSort){
 //get data to build email with
 var ann = retrieveAnnouncement(announcement_doc_id);
 var com = retrieveComic(comic_sheet_id);
 //check for image file in URL
 if ((String(com[0]).search(".png") == -1) && (String(com[0]).search(".jpg") == -1) && (String(com[0]).search(".jpeg") == -1) && (String(com[0]).search(".gif") == -1))//search for image file extension
   throw "Error! Comic URL does not contain an image file. Please fix this URL or pick another comic.";
 if (doSort)
 {
  alphabetizeSheet(form_sheet_id, full_data_range);
 }
 var data = SpreadsheetApp.openById(form_sheet_id).getSheets()[0].getDataRange().offset(1, 1).getValues();//open sheet, grab table as far as valid data extends, remove top row and leftmost column

 //create an array with the tables of the weelky reports
 var groupTables = [];
 for (i in data) {
  var row = data[i];
  for (j in row)//malice check every entry in this submission
  {
    if (String(row[j]).search("</td>") != -1)
    {
      var num = Number(j) + 2;//unexpected strings are unexpected
      var col = (num + 9).toString(16).toUpperCase();
      var error_message = "Error! Someone is trying to break the newsletter HTML! Check the Newsletter Submission spreadsheet, row " + (Number(i) + 2) + ", column " + col + ", for the tag \"&#60;/td&#62;\".";
      throw error_message;
    }
    if (String(row[j]).search("</tr>") != -1)
    {
      var num = Number(j) + 2;//unexpected strings are unexpected
      var col = (num + 9).toString(16).toUpperCase();
      var error_message = "Error! Someone is trying to break the newsletter HTML! Check the Newsletter Submission spreadsheet, row " + (Number(i) + 2) + ", column " + col + ", for the tag \"&#60;/tr&#62;\".";
      throw error_message;
    }
    if (String(row[j]).search("</table>") != -1)
    {
      var num = Number(j) + 2;//unexpected strings are unexpected
      var col = (num + 9).toString(16).toUpperCase();
      var error_message = "Error! Someone is trying to break the newsletter HTML! Check the Newsletter Submission spreadsheet, row " + (Number(i) + 2) + ", column " + col + ", for the tag \"&#60;/table&#62;\".";
      throw error_message;
    }
  }
  if (row[0])//check that there is contents in this row via the first element, which is 'Name'
     {
       groupTables.push(buildEmailRow(row));
     }
 }
 
 //make html email
 var email = buildEmail(AddAnn, AddComic, groupTables, announcement_doc_id, comic_sheet_id);
 return email
}



function sendAndSaveEmail(subject, recipient, html, save){
 //send email
 GmailApp.sendEmail(recipient, subject, "testing", {htmlBody: html});
 
 //save email if thats what we want to do
 if(save){
  saveEmail(html, HTML_folder);
 } 
}



function cleanup(){
//copy the data from the Forms Sheet to the Archive Sheet
updateArchive(form_sheet_id, archive_sheet_id);
//remove the data from the Form Sheet
clearData(form_sheet_id);
//update comic sheet
updateComics(comic_sheet_id);
}


/*=======================================================================Helper Functions=================================================================================

These functions perform the heavy-lifting of the script such as reading from the Google Sheets/Docs, compiling the newsletter HTML file, etc

========================================================================================================================================================================*/


/*-------------------------------------------------------------------------getWeek()----------------------------------------------------------------------------------

Description: Calculates the week number using Date() and Math functions, which is returned as an integer

Arguments:
-none
Returns:
-([week number] integer)
Notes:
-credits to Jarva for writing this function
---------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function getWeek() {
  var d = new Date();
  d.setHours(0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  return Math.ceil((((d - new Date(d.getFullYear(), 0, 1)) / 8.64e7) + 1) / 7);
}


/*-----------------------------------------------------------retrieveAnnouncement()-------------------------------------------------------------------------------------

Description: Grabs the announcement from the Announcement Google Doc as a string

Arguments:
-doc_id([GoogleDocId] string)
Returns:
-announcement(string)
Notes:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function retrieveAnnouncement(doc_id){
  var announcement = convertToHTML(doc_id);
  return announcement
  
}


/*--------------------------------------------------------------------retrieveComic()-------------------------------------------------------------------------------------

Description: Grabs the comic URL from the Comic Google Doc as a string

Arguments:
-doc_id([GoogleDocId] string)
Returns:
-comic(string array)
Notes:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function retrieveComic(doc_id){
  var comic_data = [];
  var comic = SpreadsheetApp.openById(doc_id).getRange('B2').getValues();
  if(comic == null)
    throw "Error! No URL found for the comic. Please check the comic spreadsheet."
  var name = SpreadsheetApp.openById(doc_id).getRange('C2').getValues();
  if(name == null || name == '')
    name = 'Anonymous';
  comic_data[0] = comic;
  comic_data[1] = name;
  return comic_data
}


/*------------------------------------------------------------------getSpreadsheetData()--------------------------------------------------------------------------------- 

Description: getSpreadsheetData() returns the values from the specified range of the specified sheet as a [string?] array

Arguments:
-range([GoogeSheetRange] string)
-sheet_id([GoogleSheetId] string)
Returns:
-([array] string)
Notes:
---------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function getSpreadsheetData(range, sheet_id) {
  var data_range = SpreadsheetApp.openById(sheet_id).getRange(range);
  return data_range.getValues()
}


/*-------------------------------------------------------buildEmailRow()-----------------------------------------------------------------------------------------------

Description: buildEmailRow() writes the tables that contain the weekly report data in HTML as a string array

Arguments:
-row([weekly report data item] string)
Returns:
-([HTML] string)
Notes:
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function buildEmailRow(row) {
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th>' + row[0] + ', ' + row[1] + '</th>',
    '<th>Project Name: ' + row[2] + '</th>',
    '</tr>',
    '<tr style = "background-color: ' + accentColor1 + ';">',
    '<td><b>This Week\'s Achievement</b></td>',
    '<td width=\"67%\">' + row[3] + '</td>',
    '</tr>',
    '<tr style = "background-color: ' + accentColor2 + ';">',
    '<td><b>This Week\'s Problem(s)</b></td>',
    '<td>' + row[4] + '</td>',
    '</tr>',
    '<tr style = "background-color: ' + accentColor1 + ';">',
    '<td><b>Next Week\'s Plan</b></td>',
    '<td>' + row[5] + '</td>',
    '</tr>',
    '</table>'
  ].join('\n')
}


/*---------------------------------------------------------buildEmailHead()-------------------------------------------------------------------------------------------

Description: Writes the header for the HTML file, which includes the CSS definitions for the tables/overall layout of the file, as a string array

Arguments:
-none
Returns:
-([HTML] string)
Notes:
-uncomment "'th {font-size:125%}'," if destination is gmail
---------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function buildEmailHead() {
  return [
    '<!DOCTYPE html>',
    '<html>',
    '<head>',
    '<style>',
    'table {width: 800px;margin: 20px;}',
    'table, th, td {border-collapse: collapse;}',
    'th, td {padding: 12px; text-align: left;}',
    /*'th {font-size:125%}',*/
    'table#t01 tr:nth-child(even) {background-color: ' + accentColor1 + ';}',
    'table#t01 tr:nth-child(odd) {background-color: ' + accentColor2 + ';}',
    'table#t01 th {background-color: ' + headerBGColor + '; color: ' + headerTextColor + ';}',
    '</style>',
    '</head>'
  ].join('\n')
}


/*---------------------------------------------------------buildEmailAnn()-------------------------------------------------------------------------------------------

Description: Writes the header for the HTML file, which includes the CSS definitions for the tables/overall layout of the file, as a string array

Arguments:
-ann_doc_id([GoogleDocId] string)
Returns:
-([HTML] string)
Notes:
-add "font-size:125%;" to 'th style' if destination is gmail
---------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function buildEmailAnn(ann_doc_id) {
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th style="text-align:center">Announcements</th>',
    '</tr>',
    '<tr style = "background-color: ' + accentColor1 + ';">',
    '<td>' + retrieveAnnouncement(ann_doc_id) + '</td>',
    '</tr>',
    '</table>'
  ].join('\n')
}


/*---------------------------------------------------------buildEmailComic()--------------------------------------------------------------------------------------------

Description: Writes the header for the HTML file, which includes the CSS definitions for the tables/overall layout of the file, as a string array

Arguments:
-comic_doc_id([GoogleDocId] string)
Returns:
-([HTML] string)
Notes:
-add "font-size:125%;" to 'th style' if destination is gmail
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function buildEmailComic(com_doc_id) {
  var com_data = retrieveComic(com_doc_id);
  return [
    '<table style="font-family:arial;" id="t01">',
    '<tr>',
    '<th style="text-align:center">Comic of the Week</th>',
    '</tr>',
    '<tr style = "background-color: ' + accentColor1 + ';">',
    '<td style="text-align:center"><img src="' + com_data[0] + '" alt="You win this round, [email client name]" width="776">Comic submitted by: ' + com_data[1] + '</td>',
    '</tr>',
    '</table>'
  ].join('\n')
}


/*---------------------------------------------------------buildEmail()--------------------------------------------------------------------------------------------------

Description: Writes all of the HTML for the email as a string array

Arguments:
-ann(boolean)
-comic(boolean)
-groupTables([Google Sheet rows] string array)
-ann_doc_id([GoogleDocId] string)
-com_doc_id([GoogleDocId] string)
Returns:
-([HTML] string)
Notes:
-uses the [A?B:C] conditional to enable or disable the tables correspoding to the announcement and comic
------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function buildEmail(ann, comic, groupTables, ann_doc_id, com_doc_id) {
  return [
    buildEmailHead(),
    '<body>',
    '<h2 style="font-family:arial;">'+ labName + ' Newsletter Week ' + getWeek() + '</h2>',
    ann ? buildEmailAnn(ann_doc_id) : ' ',
    groupTables.join('\n'),    
    comic ? buildEmailComic(com_doc_id) : ' ',
    '<h5 style="font-family:arial;"> Software Copyright © Lizzy Hatfield 2018 </h5>',
    '</body>',
    '</html>'
  ].join('\n')
}


/*-------------------------------------------------------------------------saveEmail()-------------------------------------------------------------------------------------

Description: Creates the physical HTML file out of the HTML code that was generated before

Arguments:
-email([HTML] string)
-dest_id([GoogleFolderId] string)
Returns:
-nothing
Notes:
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function saveEmail(email, dest_id) {
  var folder = DriveApp.getFolderById(dest_id); 
  var date = new Date();     
  folder.createFile('newsletter-wk' + getWeek() + '-' + date.getFullYear() +'.html', email, MimeType.HTML);
}


/*-----------------------------------------------------------updateArchive()------------------------------------------------------------------------------------------------

Description: Appends a new row of data to the Archive Sheet

Arguments:
-source_range([GoogeSheetRange] string)
-source_id([GoogeSheetId] string)
-dest_id([GoogeSheetId] string)
Returns:
-nothing
Notes:
-check for row[0] not being null, as the name field will be filled for every valid entry into the Form Sheet
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function updateArchive(source_id, dest_id) {
  var dest_sheet = SpreadsheetApp.openById(dest_id);
  var source_data = SpreadsheetApp.openById(source_id).getSheets()[0].getDataRange().offset(1, 0).getValues();
  for (i in source_data) {
    var row = source_data[i];
    if (row[0])//if row[0] == anything
    {
      dest_sheet.appendRow(row);
    }
  }   
}


/*--------------------------------------------------------------------clearData()------------------------------------------------------------------------------------------

Description: Deletes the rows of the Form Sheet in ascending order

Arguments:
-id([GoogeSheetId] string)
-depth([unsigned] integer)
Returns:
-nothing
Notes:
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function clearData(id) {
  var range = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().offset(1,0);  
  range.clear();
}


/*--------------------------------------------------------------------alphabetizeSheet()------------------------------------------------------------------------------------

Description: Sorts the rows of the Form Sheet alpabetically by name

Arguments:
-sheet_id([GoogeSheetId] string)
-sort_range([unsigned] integer)
Returns:
-nothing
Notes:
-two is for the second column which holds the 'Name' field of the Form data
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function alphabetizeSheet(sheet_id, sort_range){
  var data = SpreadsheetApp.openById(sheet_id).getSheets()[0].getDataRange().offset(1, 0);
  data.sort(2);  
}


/*-----------------------------------------------------------errorEmail()-------------------------------------------------------------------------------------

Description: Sends an email to the developer when there is an error from the GUI with the error message

Arguments:
-error_message(error.message, string)
Returns:
-nothing
Notes: emails developer on error
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function errorEmail(error_message){
GmailApp.sendEmail(devEmail, "Newsletter Generator Error!", "Error Message: " + error_message);
}


/*-----------------------------------------------------------shuffleSheet()-------------------------------------------------------------------------------------

Description: 

Arguments:
-none
Returns:
-nothing (randomizes rows of the comic sheet)
Notes: only works on the comic sheet
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function shuffleSheet(){
  var sheet = SpreadsheetApp.openById(comic_sheet_id).getSheets()[0];
  var full_range = sheet.getDataRange();
  var range = full_range.offset(1, 0);  
  range.setValues(range.randomize());
}//need to check for bad links, and an empty comic sheet


/*-----------------------------------------------------------updateComics()-------------------------------------------------------------------------------------

Description:

Arguments:
-none
Returns:
-nothing 
Notes: 
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function updateComics(comic_id){
  var d = new Date();
  //grab row 2
  var row = SpreadsheetApp.openById(comic_id).getSheets()[0].getDataRange().getValues()[1];
  //place todays date in row[0]
  row[0] = String((d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getYear());
  //copy to sheet #2
  SpreadsheetApp.openById(comic_id).getSheets()[1].appendRow(row);
  //delete row 2 in sheet #1
  SpreadsheetApp.openById(comic_id).getSheets()[0].getRange("A2:C2").clear();
}


/*-----------------------------------------------------------convertToHTML()-------------------------------------------------------------------------------------

Description: Sends an email to the developer when there is an error from the GUI with the error message

Arguments:
-GoogleDocID(string)
Returns:
-output(string array)
Notes: This code is adapted from the work by Omar Al Zabir, source at: https://github.com/oazabir/GoogleDoc2Html
----------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
function convertToHTML(doc_id){
  var body = DocumentApp.openById(doc_id).getBody();
  var numChildren = body.getNumChildren();
  var output = [];
  var images = [];
  var listCounters = {};

  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    output.push(processItem(child, listCounters, images));
  }

  return output.join('\n')
}  
 

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "", suffix = "";

  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING6: 
        prefix = "<h6>", suffix = "</h6>"; break;
      case DocumentApp.ParagraphHeading.HEADING5: 
        prefix = "<h5>", suffix = "</h5>"; break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>", suffix = "</h4>"; break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>", suffix = "</h3>"; break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>", suffix = "</h2>"; break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>", suffix = "</h1>"; break;
      default: 
        prefix = "<p>", suffix = "</p>";
    }

    if (item.getNumChildren() == 0)
      return "";
  }
  //else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE)
  //{
    //processImage(item, images, output);
  //}
  else if (item.getType()===DocumentApp.ElementType.LIST_ITEM) {
    var listItem = item;
    var gt = listItem.getGlyphType();
    var key = listItem.getListId() + '.' + listItem.getNestingLevel();
    var counter = listCounters[key] || 0;

    // First list item
    if ( counter == 0 ) {
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix = '<ul class="small"><li>', suffix = "</li>";

          suffix += "</ul>";
        }
      else {
        // Ordered list (<ol>):
        prefix = "<ol><li>", suffix = "</li>";
      }
    }
    else {
      prefix = "<li>";
      suffix = "</li>";
    }

    if (item.isAtDocumentEnd() || item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM) {
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        suffix += "</ul>";
      }
      else {
        // Ordered list (<ol>):
        suffix += "</ol>";
      }

    }

    counter++;
    listCounters[key] = counter;
  }

  output.push(prefix);

  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output);
  }
  else {


    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();

      // Walk through all the child elements of the doc.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child, listCounters, images));
      }
    }

  }

  output.push(suffix);
  return output.join('');
}


function processText(item, output) {
  var text = item.getText();
  var indices = item.getTextAttributeIndices();

  if (indices.length <= 1) {
    // Assuming that a whole para fully italic is a quote
    if(item.isBold()) {
      output.push('<b>' + text + '</b>');
    }
    else if(item.isItalic()) {
      output.push('<blockquote>' + text + '</blockquote>');
    }
    else if (text.trim().indexOf('http://') == 0) {
      output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
    }
    else {
      output.push(text);
    }
  }
  else {

    for (var i=0; i < indices.length; i ++) {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i+1 < indices.length ? indices[i+1]: text.length;
      var partText = text.substring(startPos, endPos);

      Logger.log(partText);

      if (partAtts.ITALIC) {
        output.push('<i>');
      }
      if (partAtts.BOLD) {
        output.push('<b>');
      }
      if (partAtts.UNDERLINE) {
        output.push('<u>');
      }

      // If someone has written [xxx] and made this whole text some special font, like superscript
      // then treat it as a reference and make it superscript.
      // Unfortunately in Google Docs, there's no way to detect superscript
      if (partText.indexOf('[')==0 && partText[partText.length-1] == ']') {
        output.push('<sup>' + partText + '</sup>');
      }
      else if (partText.trim().indexOf('http://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      }
      else {
        output.push(partText);
      }

      if (partAtts.ITALIC) {
        output.push('</i>');
      }
      if (partAtts.BOLD) {
        output.push('</b>');
      }
      if (partAtts.UNDERLINE) {
        output.push('</u>');
      }

    }
  }
}

