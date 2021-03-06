// ---------------------------------------------------------------------------------------------------------------------------------------------------
// The MIT License (MIT)
// 
// Copyright (c) 2014 Iain Brown - http://www.littlebluemonkey.com/blog/automatically-import-jira-backlog-into-google-spreadsheet
//
// Inspired by http://gmailblog.blogspot.co.nz/2011/07/gmail-snooze-with-apps-script.html
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

var C_MAX_RESULTS = 1000;

function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Configure Jira", functionName: "jiraConfigure"},{name: "Refresh Data Now", functionName: "jiraPullManual"}/*,{name: "Schedule 4 Hourly Automatic Refresh", functionName: "scheduleRefresh"},{name: "Stop Automatic Refresh", functionName: "removeTriggers"}*/]; 
  ss.addMenu("Jira", menuEntries);
                     
  menuEntries = [ {name: "Create cards", functionName: "createCardsFromBacklog"}, {name: "Create cards from selected rows", functionName: "createCardsFromSelectedRowsInBacklog"} ];
  ss.addMenu("Story Cards", menuEntries);
  
 }

function jiraConfigure() {

  var ss = getJiraConfigSheet();
  var host = ss.getRange("B4").getValue(); 

  PropertiesService.getUserProperties().setProperty("host", host);
  
  var userAndPassword = Browser.inputBox("Enter your Jira On Demand User id and Password in the form User:Password. e.g. Tommy.Smith:ilovejira (Note: This will be base64 Encoded and saved as a property on the spreadsheet)", "Userid:Password", Browser.Buttons.OK_CANCEL);
  var x = Utilities.base64Encode(userAndPassword);

  PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
  
  Browser.msgBox("Jira configuration saved successfully.");
}  


function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  Browser.msgBox("Spreadsheet will no longer refresh automatically.");
}  

function scheduleRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  ScriptApp.newTrigger("jiraPull").timeBased().everyHours(4).create();
  
  Browser.msgBox("Spreadsheet will refresh automatically every 4 hours.");
}  

function jiraPullManual() {
  jiraPull();
  Browser.msgBox("Jira backlog successfully imported");
}  

function getFields() {
  return JSON.parse(getDataForAPI("field"));
}  

function getStories() {
	var ss = getJiraConfigSheet();
  var jql = ss.getRange("B5").getValue();
  jql.replace("\"", "'");
  
	var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
  
  while (data.startAt + data.maxResults < data.total) {
    Logger.log("Making request for %s entries", C_MAX_RESULTS);
    data =  JSON.parse(getDataForAPI("search?jql=" + encodeURIComponent(jql) + "&maxResults=" + C_MAX_RESULTS + "&startAt=" + startAt + "&expand=status"));  
    
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
  }  
  
  return allData;
}  

function getDataForAPI(path) {
   var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
   var digestfull = PropertiesService.getUserProperties().getProperty("digest");
  
   var headers = { 
              "Accept":"application/json", 
              "Content-Type":"application/json", 
              "method": "GET",
              "headers": {"Authorization": digestfull},
              "muteHttpExceptions": true
             };
  
  var resp = UrlFetchApp.fetch(url,headers );
  if (resp.getResponseCode() != 200) {
    Browser.msgBox("Error retrieving data for url" + url + ":" + resp.getContentText());
    return "";
  }  
  else {
    return resp.getContentText();
  }  
}  

function jiraPull() {
  
  var allFields = getAllFields();
  var data = getStories();
  
  if (allFields === "" || data === "") {
    Browser.msgBox("Error pulling data from Jira - aborting now.");
    return;
  }  
  
  var ss = getBacklogSheet();
  //var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  var headings = getHeadersRange(ss).getValues()[0];
  
  var y = new Array();
  for (i=0;i<data.issues.length;i++) {
    var d=data.issues[i];
    y.push(getStory(d,headings,allFields,i+1));
  }  
  
  ss = getBacklogSheet();
  var last = ss.getLastRow();
  if (last >= 2) {
    ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();  
  }  
  
  if (y.length > 0) {
    ss.getRange(2, 1, data.issues.length, y[0].length).setValues(y);
  }
}

function getAllFields() {
  
  var theFields = getFields();
  var allFields = new Object();
  allFields.ids = new Array();
  allFields.names = new Array();
  
  for (var i = 0; i < theFields.length; i++) {
      allFields.ids.push(theFields[i].id);
      allFields.names.push(theFields[i].name.toLowerCase());
  }  
  
  return allFields;
}  

function getStory(data,headings,fields,rank) {
  
  var story = [];
  for (var i = 0;i < headings.length;i++) {
    if (headings[i] !== "") {
      if (headings[i].toLowerCase() == "rank") // hack for the black voodoo Jira ranks
        data[headings[i].toLowerCase()] = rank;
      story.push(getDataForHeading(data,headings[i].toLowerCase(),fields));
    }  
  }        
  
  return story;
}  

function getDataForHeading(data,heading,fields) {

  if (isObject(data.fields[heading])) {
    if (data.fields[heading].hasOwnProperty("name")) {
      return data.fields[heading].name;
    } 
    else if (data.fields[heading].hasOwnProperty("value")) {
      return data.fields[heading].value;
    } 
  }
  else if (data.hasOwnProperty(heading)) {
    return data[heading];
  }  
  else if (data.fields.hasOwnProperty(heading)) {
    return data.fields[heading];
  }  
  
  var indirectFieldName = getFieldName(heading,fields);
  return getDataFromIndirectFieldName(indirectFieldName, data);
  
  var splitName = heading.split(" ");
  
  if (splitName.length == 2) {
    if (data.fields.hasOwnProperty(splitName[0])) {
      if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
        return data.fields[splitName[0]][splitName[1]];
      }
      return "";
    }  
  }  
  
  return "Could not find value for " + heading;
}  

function getDataFromIndirectFieldName(indirectFieldName, data) {
  
  if (indirectFieldName !== "") {
    if (isObject(data.fields[indirectFieldName])) {
      if (data.fields[indirectFieldName].hasOwnProperty("name")) {
        return data.fields[indirectFieldName].name;
      } 
      else if (data.fields[indirectFieldName].hasOwnProperty("value")) {
        return data.fields[indirectFieldName].value;
      }         
      else if (data.fields.hasOwnProperty(indirectFieldName)) {
        var arrayToString = "";
        var prevFieldValueFound = false;
        if (Array.isArray(data.fields[indirectFieldName])) {
          for (var i = 0; i < data.fields[indirectFieldName].length; i++) {              
            
            var fieldValue = "";
            // Attention!  name only!!  (No full coverage but working)
            if (data.fields[indirectFieldName][i].indexOf("name=") != -1) {
              fieldValue = getPropertyValueFromCsvString(data.fields[indirectFieldName][i], "name")
              if (fieldValue != "") {
                if ( prevFieldValueFound ) {
                  arrayToString += ",";
                }
                prevFieldValueFound = true;
                arrayToString += fieldValue;
              } 
            }
            
            if (fieldValue == "") {
              if ( i>0 ) {
                arrayToString += ",";
              }
              arrayToString += data.fields[indirectFieldName][i];
            }
          }
          return arrayToString;
        }
      }
    } 
    else if (data.hasOwnProperty(indirectFieldName)) {
      return data[indirectFieldName];
    }  
    else if (data.fields.hasOwnProperty(indirectFieldName)) {
      return data.fields[indirectFieldName];
    }  
  }
}

function getPropertyValueFromCsvString(csvString, propertyName) {
  
  var fieldprops = csvString.split(",");
  for (var j = 0; j < fieldprops.length; j++) {
    if (fieldprops[j].indexOf(propertyName) != -1){ 
      var prop = fieldprops[j].split("=");
      return prop[1];
    } 
  }
}

function getFieldName(heading,fields) {
  var index = fields.names.indexOf(heading);
  if ( index > -1) {
     return fields.ids[index]; 
  }
  return "";
}  

function isObject(val) {
    if (val === null) { return false;}
    return ( (typeof val === 'function') || (typeof val === 'object') );
}