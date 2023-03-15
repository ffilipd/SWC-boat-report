/************************************************************************
 * The script is made for HSS SWC equipment report system to            *
 * automatically send out notifications and reminders about equipment   *
 * damage                                                               *
 * For support, contact Filip Dahlskog –  filipdahlskog@gmail.com       *
 ************************************************************************
*/

/******** CHECK WEEKLY FOR UNFIXED DAMAGES ********/
function weeklyDamageCheck() {
  var event = 'weekly';                                                                       // NAME OF EVENT
  var [taskSheets,spreadSheet] = getSheets('tasks');                                          // GET MAINTENANCE TASK LISTS
  var [damagedBoats, allDamagesList] = getDamagedBoats(taskSheets, spreadSheet, event);       // GET UNFIXED DAMAGES BY EVENT
  
  if (damagedBoats.length > 0) {                                                              // IF UNFIXED DAMAGES EXISTS
    var mailList = getMailList(spreadSheet);                                                  // GET EMAIL ADDRESSES
    remindCrewLeaders(damagedBoats, mailList, spreadSheet, event);                            // SEND REMINDER EMAIL TO CREWLEADERS
  
    sendCopies(mailList, allDamagesList, spreadSheet, event);                                        // SEND MAIL COPIES
  }
}

/******** GET MESSAGE TEMPLATE FROM SHEET ********/
function getMessageTemplate(spreadSheet, event) {
  var messageSheet = getSheets(event)[0];
  var messageArray = spreadSheet.getSheetByName(messageSheet).getRange(2,1,6,2).getValues(); // GET THE MESSAGE TEMPLATE FROM SPREADSHEET
  var messageObject = {
    subject:messageArray[0][1],
    header:messageArray[1][1],
    subHeader:messageArray[2][1],
    footer:messageArray[3][1],
    signature:messageArray[4][1],
  }
  return messageObject;
}


/******* CHECK FOR MAJOR DAMAGES - EVERY TIME A BOAT REPORT OCCURS *******/
function majorDamageCheck() {
  var event = 'majorDamage';                                                                // NAME OF EVENT

  var [taskSheets,spreadSheet] = getSheets('tasks');                                           // GET MAINTENANCE TASK LIST
  var [damagedBoats, allDamagesList] = getDamagedBoats(taskSheets, spreadSheet, event);     // GET UNFIXED DAMAGES BY EVENT
  
  if (damagedBoats.length > 0) {                                            // IF UNFIXED DAMAGES EXISTS
    var mailList = getMailList(spreadSheet);                                // GET EMAIL ADDRESSES

    remindCrewLeaders(damagedBoats, mailList, spreadSheet, event);                       // SEND EMAIL TO CREWLEADER ABOUT MAJOR DAMAGE
    damagedBoats.forEach(damage => damage.damages[3]);                      // MARK MAJOR DAMAGE AS "NOTIFIED" IN MAINT. TASK LIST
    sendCopies(mailList, allDamagesList, spreadSheet, event);                                   // SEND MAIL COPIES
  }
}




/******* GET NAME SPECIFIC SHEETS *******/
function getSheets(sheetName) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();          // GET ALL SHEETS
  var sheets = spreadSheet.getSheets();                             // ASSIGN SHEETS TO A VARIABLE
  var ret = [];                                                     // DECLARE ARRAY FOR PROPERTIES TO RETURN
  var regexSearchString = new RegExp(sheetName , 'i');              // CREATE A REGEX STRING

  for (i = 0; i < sheets.length; i++) {                             // TAKE ONE SHEET A TIME
    var sheetName = sheets[i].getName();                            // GET THE NAME OF THE CURRENT SHEET
    if (sheetName.match(regexSearchString)) {                       // USE REGEX TO FIND SHEET WITH PROVIDED STRING IN THE NAME
      ret.push(sheets[i].getName());                                // IF MATCH, ADD TO 'ret'
    }
  }
  return [ret, spreadSheet];                                        // RETURN THE ARRAY AND THE WHOLE SPREADSHEET
}



/******* GET DAMAGES FROM MAINTENANCE TASK LISTS *******/
function getDamagedBoats(taskSheets, spreadSheet, event) {
  var damagedBoats = [];                                            // DECLARE VARIABLE FOR DAMAGED BOATS
  var allDamagesList = [];                                          // DECLARE VARIABLE FOR ALL DAMAGES (TO BE USED IN EMAIL COPIES)

  for (i = 0; i < taskSheets.length; i++) {                         // TAKE ONE TASK SHEET A TIME
    var boatNameCell = "A1";                                        // WE KNOW THAT BOAT NAME IS IN THE FIRST CELL OF TASK SHEET
    var currentSheet = spreadSheet.getSheetByName(taskSheets[i]);   // TAKE THE CURRENT SHEET
    var currentBoatName = currentSheet.getRange(boatNameCell).getValue();   // GET THE NAME OF THE BOAT

    var taskList = currentSheet.getRange(3,1,999,6);                // GET THE RANGE OF TASKS
    var values = taskList.getValues();                              // GET THE VALUES FROM THE RANGE
    var damagedBoatObj = {                                          // PREPARE AN OBJECT FOR STORING TASKS
      name:     currentBoatName,
      id:       currentSheet.getSheetId(),
      damages:  []
    };

    for (let j = 0;j < values.length; j++) {                        // TAKE ONE TASK AT A TIME
      var statusCell = 0;                                           // FIRST CELL IS STATUS OF DAMAGE (FIXED OR NOT)
      var dateCell = 2;                                             // CELL #2 IS DATE WHEN DAMAGE WAS REPORTED
      var damageTypeCell = 3;                                       // CELL #3 IS TYPE OF DAMAGE (MAJOR OR MINOR)
      var descriptionCell = 4;                                      // CELL #4 IS DESCRIPTION OF DAMAGE
      var notifiedCell = 5;                                         // CELL #5 TELLS US IF CREWLEADER HAVE BEEN NOTIFIED ABOUT MAJOR DAMAGES
      var fixedStatus = values[j][statusCell];                      // GET STATUS OF DAMAGE
      var taskDescription = values[j][descriptionCell];             // GET DESCRIPTION
      var damageDate = values[j][dateCell];                         // GET THE DATE WHEN DAMAGE WAS REPORTED

      if (fixedStatus == true) continue;                            // IF DAMAGE IS MARKED AS FIXED, TAKE NEXT DAMAGE
      if (damageDate.length < 2) {                                  // IF DATE CELL IS EMPTY, WE HAVE REACHED THE END OF LIST
        break;                                                      // BREAK THE LOOP AND TAKE NEXT TASK SHEET
      }

      var dateFormatted = damageDate.toDateString();                // FORMAT THE DATE TO MORE READABLE VALUE
      var damageType = values[j][damageTypeCell];                   // TAKE THE DAMAGE TYPE
      var majorDamage = /major/i.test(damageType);                  // CHECK IF DAMAGE TYPE IS 'MAJOR'
      var notified = values[j][notifiedCell] ?? false;              // CHECK IF DAMAGE HAVE BEEN NOTIFIED

      if (                                                          // IF
        majorDamage &&                                              // THE DAMAGE IS MAJOR 
        event == 'majorDamage' &&                                         // THE EVENT IS 'majorDamage'
        notified == false)                                          // THE CREWLEADER HAS NOT BEEN NOTIFIED
      {
        var taskRow = j + 3;                                        // OFFSET ROW BY +3 TO GET THE CURRENT TASK ROW IN SHEET
        var notifiedColumn = notifiedCell + 1;                      // OFFSET THE COLUMN BY +1 TO GET THE NOTIFIED COLUMN IN SHEET
        var setNotified = currentSheet.getRange(taskRow ,notifiedColumn).setValue('true'); // ASSIGN THE STATUS CHANGE TO A VARIABLE
        appendDamages();                                            // APPEND DAMAGES
      }

      if (event == 'weekly') {                                      // IF THE EVENT IS 'weekly'
        appendDamages();                                            // APPEND DAMAGES
      }

      /** THIS FUNCTION IS COMMON FOR WEEKLY DAMAGE CHECK AND MAJOR DAMAGE CHECK */
      function appendDamages() {                                    
        damagedBoatObj.damages.push([                               // ADD THE TASK TO THE DAMAGED BOAT OBJECT
          dateFormatted,
          damageType,
          taskDescription,
          setNotified = notified ? undefined : setNotified          // THIS IS ONLY ASSIGNED WHEN MAJOR DAMAGES ARE CHECKED
          ]);
        allDamagesList.push([                                       // ADD THE TASK TO THE LIST WITH ALL DAMAGES (FOR COPY)
          currentBoatName,
          damageType,
          taskDescription
          ]);
      }
    }
    if (damagedBoatObj.damages.length > 0) damagedBoats.push(damagedBoatObj); // IF DAMAGES WERE FOUND AND APPENDED
  }

  return [damagedBoats, allDamagesList]                             // RETURN THE LISTS
}




/******* SEND EMAIL TO CREW LEADERS ABOUT DAMAGES *******/
function remindCrewLeaders(damagedBoats, mailList, spreadSheet, event) {
  var messageTemplate = getMessageTemplate(spreadSheet, event);

  for (let i = 0; i < damagedBoats.length; i++) {                     // TAKE ONE DAMAGE A TIME
    var damagedBoat = damagedBoats[i];                                // GET THE CURRENT DAMAGED BOAT
    var damagedBoatName = damagedBoat.name;                           // TAKE THE NAME OF THE BOAT
    var firstLetterInDamagedBoatName = damagedBoatName[0];            // TAKE THE FIRST CHARACTER OF BOAT NAME
    var damagedBoatNumber = damagedBoatName[damagedBoatName.length - 1];  // AND THE LAST CHARACTER

    /** 
     * now we have the first and last character of the boat name, which should be e.g. "E" and "5" or "J" and "0"
     * we use the the character to match with boat names assigned to mailList. By matching first and last character (also
     * case-insensitive) misspelling is omitted. 
     */
    var regexSearchString = new RegExp(                               // CREATE A REGEX STRING TO SEARCH FOR THE CREWLEADER
      '^' + firstLetterInDamagedBoatName +
      '.*' + damagedBoatNumber + 
      '$', 'i'
      );
    var boatNameOfCrewLeaderInfo = 0;                                 // BOAT NAME IS IN CELL[0]
    var crewleaderInfo = mailList.filter(leader =>                    // USE THE REGEX STRING TO FILTER OUT THE CREWLEADER FOR THIS BOAT
    leader[boatNameOfCrewLeaderInfo].match(regexSearchString))[0];    // Array.filter() RETURNS AN ARRAY, SO PICK INDEX 0

    var crewLeaderEmailIndex = 1;                                       // INDEX FOR EMAIL ADDRESS
    var crewLeaderEmailAddress = crewleaderInfo[crewLeaderEmailIndex];  // TAKE THE EMAIL ADDRESS OF CREWLEADER

    var mailTemplate = 'mailbase';
    var templ = HtmlService.createTemplateFromFile(mailTemplate);       // CREATE A TEMPLATE
    var spreadSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var taskListUrl = spreadSheetUrl + '#gid=' + damagedBoat.id;

    var messageProperties = {                                                        // CREATE AN OBJECT FOR THE TEMPLATE TO USE
      header: messageTemplate.header,
      subHeader: messageTemplate.subHeader,
      footer: messageTemplate.footer,
      signature: messageTemplate.signature,
      name: damagedBoatName,
      damages: damagedBoat.damages,
      url: taskListUrl
    }

    templ.props = messageProperties;                                            // GIVE THE PROPERTIES OBJECT TO THE TEMPLATE
    var message = templ.evaluate().getContent();                                // CREATE A MESSAGE FROM THE TEMPLATE

    if (crewLeaderEmailAddress.length > 1) {                                    // IF THE CREWLEADER EMAIL ADDRESS EXISTS
        MailApp.sendEmail({                                                     // SEND THE EMAIL TO THE CREWLEADER
        to: crewLeaderEmailAddress,                                             // EMAIL ADDRESS  
        subject: `${messageTemplate.subject} – ${messageProperties.name}`,      // SUBJECT
        htmlBody: message                                                       // MESSAGE (FROM TEMPLATE)
      });

    }
  }
}



/******* SEND COPIES TO MAIL ADRESSES MARKED WITH 'COPY' *******/
function sendCopies(mailList, damages, spreadSheet, event) {
  var eventType = event != 'weekly' ? 'MAJOR' : 'Weekly';
  var type = 'copy'
  var copyMailList = mailList.filter(name => name[0] == type)         // FILTER OUT EMAIL ADDRESSES MARKED 'copy' IN THE LIST
  var templ = HtmlService.createTemplateFromFile('copymailbase');     // CREATE A TEMPLATE FOR THE MESSAGE
  var spreadSheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var messageTemplate = getMessageTemplate(spreadSheet, type);
  var messageProperties = {                                           // CREATE AN OBJECT WITH PROPERTIES FOR THE TEMPLATE
    header: messageTemplate.header,
    subHeader: messageTemplate.subHeader,
    footer: messageTemplate.footer,
    signature: messageTemplate.signature,
    damages: damages,
    url: spreadSheetUrl,
    bgColor1: '#d6d6d6',                                            // COLOR 1 FOR ALTERNATING BACKGROUD COLOR IN TABLE
    bgColor2: '#fff'                                                // COLOR 2 FOR ALTERNATING BACKGROUD COLOR IN TABLE
  }

  templ.props = messageProperties;                                                // GIVE THE OBJECT TO THE TEMPLATE
  var message = templ.evaluate().getContent();                      // CREATE A MESSAGE FROM THE TEMPLATE

  copyMailList.forEach(copy =>{                                     // FOR EACH EMAIL ADDRESS MARKED 'copy'
    var emailAddr = copy[1];                                        // INDEX OF EMAIL ADDRESS IS 1
    if (emailAddr.length > 1){                                      // IF THERE IS AN EMAIL ADDRESS
      MailApp.sendEmail({                                           // SEND THE EMAIL COPY
        to: emailAddr,
        subject: `${eventType} ${messageTemplate.subject} - copy`,
        htmlBody: message
      });
    };
  })
}


/******* GET MAIL LIST *******/
function getMailList(spreadSheet) {
  return spreadSheet.getSheetByName("Mail List").getRange(2,1,20,2).getValues(); // GET THE MAIL LIST FROM SPREADSHEET
}
