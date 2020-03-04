var SOLA_VISTA_SPREADSHEET = "1-XAiw9RgU4Mye5E8HSQxKUNzL20na3RMWQKZPTn6dFU";

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
  
  for (var i = 1; i < (formValues[row].length-1); i++) {
    formValues[row][i] = (formValues[row][i] instanceof String)?(formValues[row][i].trim()):formValues[row][i];
    switch (formValues[0][i]) {
      case "Location Name?":
        storeID = formValues[row][i];
        break;
      case "Style?":
        fixturePackage = formValues[row][i];
        break;
      case "Color Combination Selection":
        colorScheme = formValues[row][i].split(" ")[0];
        break;
      default:
        break;
    }
  }
  var isSuccessUpdateStoreInSolaVista = updateStoreInSolaVista(storeID, fixturePackage, colorScheme);
  var isUpdateAsanaFixturePackageTask = updateAsanaFixturePackageTask(storeID, fixturePackage);
  return true;
}

function updateStoreInSolaVista(storeID, fixturePackage, colorScheme) {
  var spreadsheet = SpreadsheetApp.openById(SOLA_VISTA_SPREADSHEET);
  var now = new Date();
  var currentYear = now.getFullYear();
  var sheet = spreadsheet.getSheetByName(currentYear);
  var storeList = sheet.getRange(4,1, sheet.getLastRow()-3,4).getValues();
  var rowIndex = -1;
  for (var i=0;i<storeList.length;i++) {
    var currentStoreID = storeList[i][0];
    if (storeID == currentStoreID) {
      rowIndex = i+4;
      break;
    }
  }
  if (rowIndex == -1)
    return false;
  else {
    sheet.getRange(rowIndex,26).setValue(colorScheme);
    sheet.getRange(rowIndex,29).setValue(fixturePackage);
    return true;
  }
}

function updateAsanaFixturePackageTask(storeID, fixturePackage) {
  var tasks = getAsanaTasks(storeID);
  for (var i=0;i<tasks.length;i++) {
    var id = tasks[i].gid;
    if (tasks[i].name == "Fixture Package") {
      updateAsanaTaskName(tasks[i].gid, "Fixture Package "+fixturePackage);
      markAsanaTaskAsCompleted(tasks[i].gid);
      return true;
    }
  }
  return false;
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
