/*************************
 * Asana     Functions    *
 *************************/

// first Global constants ... Key Ids / tokens etc.
ASANA_PERSONAL_ACCESS_TOKEN = "0/69309846243066eb845b3a14851cae7c"; // Put your unique Personal access token here
ASANA_WORKSPACE_ID = "730301893411225"; // Put in the main workspace key you want to access (you can copy from asana web address)

function getAsanaTasks(storeID) {
  var projectID = getAsanaProjectID(storeID);
  if (projectID != "") {
    var bearerToken = "Bearer " + ASANA_PERSONAL_ACCESS_TOKEN;
    var options = {
      "method" : "GET",
      "headers" : {"Authorization": bearerToken}, 
      "contentType": 'application/json'
    };
    try {
      var url = "https://app.asana.com/api/1.0/tasks?project="+projectID;
      var result = UrlFetchApp.fetch(url, options);
      var reqReturn = result.getContentText();
      var tasks = JSON.parse(reqReturn).data;
      return tasks;
    } 
    catch (e) {
      Logger.log(e);
    }
  }
  return [];
}

function getAsanaProjectID(storeID) {
  var bearerToken = "Bearer " + ASANA_PERSONAL_ACCESS_TOKEN;
  var options = {
    "method" : "GET",
    "headers" : {"Authorization": bearerToken}, 
    "contentType": 'application/json'
  };
  try {
    var url = "https://app.asana.com/api/1.0/projects?workspace="+ASANA_WORKSPACE_ID;
    var result = UrlFetchApp.fetch(url, options);
    var reqReturn = result.getContentText();
    var projects = JSON.parse(reqReturn).data;
    for (var i=0;i<projects.length;i++) {
      if (projects[i].name.toUpperCase() == storeID.trim().toUpperCase())
        return projects[i].gid;
    }
  } 
  catch (e) {
    Logger.log(e);
  }
  return "";
}

function markAsanaTaskAsCompleted(taskId) {
  var bearerToken = "Bearer " + ASANA_PERSONAL_ACCESS_TOKEN;
  var task = {
    data: {
      "completed": true
    }
  };
  var options = {
    "method" : "PUT",
    "headers" : {"Authorization": bearerToken}, 
    "contentType": 'application/json',
    "payload" : JSON.stringify(task) 
  };
  try {
    var url = "https://app.asana.com/api/1.0/tasks/"+taskId;
    var result = UrlFetchApp.fetch(url, options);
    var reqReturn = result.getContentText();
    //Logger.log(reqReturn);
    var isTaskCompleted = JSON.parse(reqReturn).data.completed;
    //Logger.log(isTaskCompleted);
    return isTaskCompleted;
  } 
  catch (e) {
    Logger.log(e);
  }
  return false;
}

function updateAsanaTaskName(taskId, name) {
  var bearerToken = "Bearer " + ASANA_PERSONAL_ACCESS_TOKEN;
  var task = {
    data: {
      "name": name
    }
  };
  var options = {
    "method" : "PUT",
    "headers" : {"Authorization": bearerToken}, 
    "contentType": 'application/json',
    "payload" : JSON.stringify(task) 
  };
  try {
    var url = "https://app.asana.com/api/1.0/tasks/"+taskId;
    var result = UrlFetchApp.fetch(url, options);
    var reqReturn = result.getContentText();
    //Logger.log(reqReturn);
    return true;
  } 
  catch (e) {
    Logger.log(e);
  }
  return false;
}