var SOLA_VISTA_SPREADSHEET = "1-XAiw9RgU4Mye5E8HSQxKUNzL20na3RMWQKZPTn6dFU";
var TESTING = false;

// Create dropdown menu and options
function onOpen()
{
  var menu = SpreadsheetApp.getUi().createMenu("SYMPRO");
  menu.addItem('Check new row to process', 'checkNewRowToProcess');
  menu.addToUi();
} // onOpen()

function checkNewRowToProcess() {
  try {
    var lastOnGoingCheckTimestamp = doProjectProperty_("LastOnGoingCheckTimestamp");
    var now = new Date();
    if (true || lastOnGoingCheckTimestamp == "" || (now.getTime() - lastOnGoingCheckTimestamp) > 7*60*1000) {
      doProjectProperty_("LastOnGoingCheckTimestamp",now.getTime());
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sola Color Selection");
      var lastColumnIndex = sheet.getLastColumn();
      sheet.sort(1,false);
      var formValues = sheet.getDataRange().getValues();
      for (var i = (formValues.length-1); i >= 1; i--) {
        if (formValues[i][lastColumnIndex-1] != "X" && formValues[i][lastColumnIndex-1] != "on-going") {
          sheet.getRange(i+1,lastColumnIndex).setValue("on-going");
          if (processRow(i))
            sheet.getRange(i+1,lastColumnIndex).setValue("X");
        }
      }      
      doProjectProperty_("LastOnGoingCheckTimestamp","");
    }
  }
  catch (e) {
    Logger.log(e);
    sendErrorLog("Sola Color Selection - New row processing");
    doProjectProperty_("LastOnGoingCheckTimestamp","");
  }
}

function processRow(row) {
  var formValues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sola Color Selection").getDataRange().getValues();
  var storeID = "";
  var fixturePackage = "";
  var colorScheme = "";
  var shampooBowlColor = "";
  for (var i = 1; i < (formValues[row].length-1); i++) {
    formValues[row][i] = (formValues[row][i] instanceof String)?(formValues[row][i].trim()):formValues[row][i];
    switch (formValues[0][i]) {
      case "Location Name?":
        storeID = formValues[row][i];
        break;
      case "Fixture Package?":
        fixturePackage = formValues[row][i];
        break;
      case "Select your cabinet & counter-top color combination:":
        if (formValues[row][i] != "")
          colorScheme = formValues[row][i].split(" ")[0];
        break;
      case "Select your Shampoo Bowl Color:":
        shampooBowlColor = formValues[row][i];
      default:
        break;
    }
  }
  Logger.log("storeID: "+storeID);
  var statusUpdateStoreInSolaVista = updateStoreInSolaVista(storeID, fixturePackage, colorScheme, shampooBowlColor);
  Logger.log("statusUpdateStoreInSolaVista: "+statusUpdateStoreInSolaVista);
  var statusUpdateAsanaTasks = updateAsanaTasks(storeID, fixturePackage, colorScheme, shampooBowlColor);
  Logger.log("statusUpdateAsanaTasks: "+statusUpdateAsanaTasks);
  
  if (!statusUpdateStoreInSolaVista.success || !statusUpdateAsanaTasks.successFixturePackageTask || !statusUpdateAsanaTasks.successColorSchemeReceivedTask) {
    var mailSubject = storeID+": ERROR on Sympro Vista/Asana Sync **AC ACTION REQUIRED**";
    var mailBody = "";
    if (statusUpdateStoreInSolaVista.ac != "")
      mailBody += statusUpdateStoreInSolaVista.ac+",";
    else
      mailBody += "Sola Team,";
    
    if (!statusUpdateStoreInSolaVista.success) {
      mailBody += "<br/><br/>"+storeID+" not found in Vista. Please confirm project name in both files match. Contact the Sympro manager to re-process the Color Selection form response.";
    }
    if (!statusUpdateStoreInSolaVista.isProjectFoundInAsana) {
      if (statusUpdateStoreInSolaVista.success)
      storeID+" not found in Asana. Please confirm project name in both files match. Contact the Sympro manager to re-process the Color Selection form response.";
    }
    if (!statusUpdateAsanaTasks.successFixturePackageTask)
      mailBody += "<br/><br/>Asana \"Fixture Package\" task not found. Please copy the missing task from the Asana template to your project. Then contact the Sympro manager to re-process the Color Selection form response.";
    if (!statusUpdateAsanaTasks.successColorSchemeReceivedTask)
      mailBody += "<br/><br/>Asana \"Color Scheme Received\" task not found. Please copy the missing task from the Asana template to your project. Then contact the Sympro manager to re-process the Color Selection form response.";
    
    var mailTo = (TESTING)?Session.getEffectiveUser().getEmail():"solateam@sympatecoinc.com";    
    var options = {};
    options['htmlBody'] = mailBody;
    if (!TESTING) {
      options['bcc'] = "sl.sympro@sympatecoinc.com, gillianm@sympatecoinc.com";
    }                    
    MailApp.sendEmail(mailTo, mailSubject, "", options);
  }
  return true;
}

function updateStoreInSolaVista(storeID, fixturePackage, colorScheme, shampooBowlColor) {
  var spreadsheet = SpreadsheetApp.openById(SOLA_VISTA_SPREADSHEET);
  var now = new Date();
  var currentYear = now.getFullYear();
  var sheet = spreadsheet.getSheetByName("2020 Asana test");
  var storeList = sheet.getRange(4,1, sheet.getLastRow()-3,4).getValues();
  var rowIndex = -1;
  for (var i=0;i<storeList.length;i++) {
    var currentStoreID = storeList[i][0];
    if (storeID == currentStoreID) {
      rowIndex = i+4;
      var ac = storeList[i][1];
      sheet.getRange(rowIndex,26).setValue(colorScheme);
      sheet.getRange(rowIndex,27).setValue((shampooBowlColor == "White")?"WHT":"BLK");
      sheet.getRange(rowIndex,29).setValue(fixturePackage);
      return {success: true, ac:ac}
    }
  }
  if (rowIndex == -1)
    return {success: false, ac: ""};
}

function updateAsanaTasks(storeID, fixturePackage, colorScheme, shampooBowlColor) {
  var tasks = getAsanaTasks(storeID);
  var successFixturePackageTask = false;
  var successColorSchemeReceivedTask = false;
  if (tasks.length == 0)
    return {successFixturePackageTask: false, successColorSchemeReceivedTask:false, isProjectFoundInAsana: false};
  for (var i=0;i<tasks.length;i++) {
    var id = tasks[i].gid;
    if (tasks[i].name == "Fixture Package") {
      updateAsanaTaskName(tasks[i].gid, "Fixture Package "+fixturePackage);
      markAsanaTaskAsCompleted(tasks[i].gid);
      successFixturePackageTask = true;
    }
    else if (tasks[i].name == "Color Scheme Received") {
      updateAsanaTaskName(tasks[i].gid, "Color Scheme Received "+colorScheme+"/"+shampooBowlColor.toLowerCase());
      markAsanaTaskAsCompleted(tasks[i].gid);
      successColorSchemeReceivedTask = true;
    }
    
  }
  return {successFixturePackageTask: successFixturePackageTask, successColorSchemeReceivedTask: successColorSchemeReceivedTask, isProjectFoundInAsana: true};  
}

// PRIVATE Getter-Setter:
// Sets the script property if value parameter exists
// Gets the script property if value parameter does not exist
function doProjectProperty_(key, value)
{ 
  // Record all read/write actions for measurement
  //var propertiesRecordSheet = SpreadsheetApp.openById("1n6JN1h6uqdZjmM0leoAAIU3yK9JF6puARYhA6YyHyCI").getSheetByName("Sheet1");  
  //propertiesRecordSheet.getRange(propertiesRecordSheet.getLastRow()+1,1,1,4).setValues([[new Date(),value?"W":"R",key,SpreadsheetApp.getActiveSpreadsheet().getName()]]);
    
  var properties = PropertiesService.getDocumentProperties();
  // Check if second parameter exists
  if (value)
  {
    // Add or set the given key
    properties.setProperty(key, JSON.stringify(value));
  }
  else
  {
    // Return the value of the given key
    return JSON.parse(properties.getProperty(key)) || "";
  } // if value is not empty
} // doProjectProperty_()
