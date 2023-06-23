//global variables
var stuff = SpreadsheetApp.getActiveSheet();


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Who is Teaching this Week')
      .addItem('If it asks for authorization, it didn\'t run, run it again','NothingHere')
      .addItem('Run', 'triggerMe')
      .addItem('Always automatically trigger Wednsday, 5-6 (can\'t be undone here- you\'ll have to go to the script editor)', 'createTimeTriggerEveryNWeeks')
     
      .addToUi();
}




function createTimeTriggerEveryNWeeks() {
 ScriptApp.newTrigger("triggerMe")
   .timeBased()
   .atHour(17)
   .everyWeeks(7)
   .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
   .create();
}

function triggerMe() {
  var value = stuff.getRange("A13").getValue();
  Logger.log(value);
  if (value == "Yes") {
    Logger.log("Running everything now....");
    decideWho();
  } else {
    Logger.log("We'll run next week");
  }

}

//look for who has the least amount of lessons taught under their belt
function decideWho() {
     Logger.log("deicde Called");

  //how many cells to loop through
  var AmountPeople = stuff.getRange("D2").getValue();

  //start counter high
  var min = 100; 

  //loop through everyone until we get to the lowest amount of lessons taught
  for (var row = 2; row < AmountPeople+2; row++) {
    var currentValue = stuff.getRange(row,2).getValue();

    if (min > currentValue) {
      min = currentValue;
      personIs = stuff.getRange(row,1).getValue();
      rowIs = row;
    }

  }
 //text them and then update the spreadsheet
  notifyThem(rowIs);
  UpdateWhoTaught(rowIs);
  notifyLeaders(rowIs);
}

function decideWhoButNot(notRow) {
     Logger.log("deicde but not Called");

  //how many cells to loop through
  var AmountPeople = stuff.getRange("D2").getValue();

  //start counter high
  var min = 100; 

  //loop through everyone until we get to the lowest amount of lessons taught
  for (var row = 2; row < AmountPeople+2; row++) {
    var currentValue = stuff.getRange(row,2).getValue();

    if (min > currentValue && notRow != row) {
      min = currentValue;
      personIs = stuff.getRange(row,1).getValue();
      rowIs = row;
    }

  }
 //text them and then update the spreadsheet
  notifyThem(rowIs);
  UpdateWhoTaught(rowIs);
  notifyLeaders(rowIs);

}


//send them a little reminder
 function notifyThem(row) {
      Logger.log("Notify Called");


   //get their email/text thing
   var sendTo = stuff.getRange(row, 3).getValue();

   //get name of person teaching
   var theirName = stuff.getRange(row, 1).getValue(); 

   //get message being sent
   var message = stuff.getRange("J2").getValue();
  

  MailApp.sendEmail({
    name: "Teaching Lesson Reminder",
    to: sendTo,
    subject: "",
    body: theirName+message+row+" "
  });
 }

 function UpdateWhoTaught(row) {
   Logger.log("Update Called");
   var old = stuff.getRange(row,2).getValue();
   stuff.getRange(row,2).setValue(old+1);

 }



function doGet(e) {
  var row = e.queryString;
  var html = HtmlService.createTemplateFromFile('Index');
  html.row = row;
  return html.evaluate();
}

function doSomething(row) {   
  //re-run the functions to get a different person
  Logger.log(row);
  decideWhoButNot(row);
  Logger.log("editting continued");
  old = stuff.getRange(row, 2).getValue();
  stuff.getRange(row, 2).setValue(old-1);
}

function notifyLeaders(youthRow) {

   //how many cells to loop through
  var AmountPeople = stuff.getRange("H2").getValue();

//name of the youth teaching
var youthTeaching = stuff.getRange(youthRow, 1).getValue();

  //send them all a message
  for (var row = 2; row < AmountPeople+2; row++) {
  var email = stuff.getRange(row,7).getValue();
 MailApp.sendEmail({
    name: "Teaching Lesson Reminder",
    to: email,
    subject: "",
    body: youthTeaching+" should have the lesson this week."
  });


  }
}



function clearOverride() {
  stuff.getRange("C16").setValue("FALSE");
  stuff.getRange("C17").setValue("FALSE");
}
