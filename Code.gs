// Declare global variables

var today = new Date();
var currentYear = today.getFullYear();
var currentMonth = today.getMonth(); //0-11
var monthDates = [31, currentYear % 4 ==0 ? 29:28,31,30,31,30,31,31,30,31,30,31];
var monthName = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]; 

var raEvents = ["In", "Duty Phone", "Duty Round"];


//Spread Sheet

var scriptProperties = PropertiesService.getScriptProperties();
var userID = Session.getActiveUser().getEmail();
scriptProperties.setProperty(userID+'userID', userID);

if(scriptProperties.getProperty(userID+'selectedNotif') === null){
  scriptProperties.setProperty(userID+'selectedNotif', "None");
}

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet;
var sheets;
var  lastRow;
var  lastColumn;



//Calender
var calendars = CalendarApp.getAllOwnedCalendars();
// -------- Sidebar HTML Setup Functions --------
  
function onInstall(e) {
   onOpen(e);
}
// As the spreadsheet opens add a menu
function onOpen() {
    var ui = SpreadsheetApp.getUi(); 
    ui.createMenu('Ins and Outs Scheduler')
        .addItem('Update Calender', 'showFormSidebar')
        .addToUi();
    
}

function showFormSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('Form')
        .setTitle('Ins and Outs Scheduler')
        .setWidth(300);
     SpreadsheetApp.getUi()
        .showSidebar(html);
    getDate();
    getSheet();
  //scriptProperties.deleteAllProperties();
     
}

function getAllHeaders(){
    var allItems = sheet.getRange(1, 1, 1, lastColumn).getValues();
    return createDropMenu(allItems);
}

// Get first row of spreadsheet, return to html for drop menu
function createDropMenu(allItems){
    var drop = "";
    for (let i = 0; i < lastColumn; i++){
        drop += '<option value="' + i + '">'  + getLetter(i) + allItems[0][i] + '</option>';
    }
    return drop;
}


// Get all of your account calendars, return to html for drop menu
function getCalendars(){
    var drop = "";
    for (var i = 0; i < calendars.length; i++){
        drop += '<option value="' + i + '">'  + calendars[i].getName() + '</option>';
        if(calendars[i].getId() == scriptProperties.getProperty(userID+'calId')){
        scriptProperties.setProperty(userID+'calendarIndex', i);
        }
    }
    return drop;
}

// Returns array of saved properties
// This is used to populate the sidebar with previously saved data
function getSavedPropsForSidebar() {
    var propertiesAndKeys = {}
    var data = scriptProperties.getProperties();
    for (var key in data) {
        propertiesAndKeys[key] = scriptProperties.getProperty(key);
    }
    return propertiesAndKeys;
}



function selectCurrentCell(){
    getDate();
    getSheet();
   
    var cellRange = sheet.getCurrentCell();
    var name = cellRange.getValue();
    var location = cellRange.getA1Notation();
    
    scriptProperties.setProperty(userID+'raName', name);
  scriptProperties.setProperty('raName',  scriptProperties.getProperty(userID+"raName"));
    scriptProperties.setProperty(userID+'nameLocation', location);
  

    //go up until reach one cell above top of list of names 
    var currentCell = sheet.getRange(cellRange.getRow()-1, cellRange.getColumn());
    while(!currentCell.isBlank()){
        currentCell = sheet.getRange(currentCell.getRow()-1, currentCell.getColumn());
    }
    //go left until reach date or number
    var nameDateGap =1;
    currentCell = sheet.getRange(currentCell.getRow(), currentCell.getColumn()+1);
    while(!checkIfDate(currentCell.getValue())){
        nameDateGap = nameDateGap+1; 
        currentCell = sheet.getRange(currentCell.getRow(), currentCell.getColumn()+1);
     }
    var firstDate = currentCell.getValue();
    //convert to Integer if Date Object
    if(checkIfDateObject(firstDate)){
        scriptProperties.setProperty(userID+'isDateOBject', "yes");
        firstDate = firstDate.getDate();
    }
    //convert to number if in the form "6th"
    else if (String(firstDate).length > 2){
        var tmpVal = Number(String(firstDate).slice(0,-2));
        if(!isNaN(tmpVal)){
            firstDate = parseInt(tmpVal);
        }
    }
    
    scriptProperties.setProperty(userID+'nameToDateGap', nameDateGap);
    scriptProperties.setProperty(userID+'firstDate', firstDate);
    
    return  getSavedPropsForSidebar();
}





// -------- Save Submitted Sidebar Data --------

// User clicked submit, save all info to properties
function saveSidebar(sideData) {
   //scriptProperties.deleteAllProperties();
   scriptProperties.setProperty(userID+'calendarIndex', sideData.calIndex);
   scriptProperties.setProperty(userID+'calId', calendars[sideData.calIndex].getId());
   scriptProperties.setProperty(userID+'inNightSymbol', sideData.inNightSymbol);
   scriptProperties.setProperty(userID+'dutyPhoneSymbol', sideData.dutyPhoneSymbol);
   scriptProperties.setProperty(userID+'dutyRoundSymbol', sideData.dutyRoundSymbol);
  updateDisplayProperties();
  
}


function updateDisplayProperties(){
  
              scriptProperties.setProperty('calendarIndex',  scriptProperties.getProperty(userID+"calendarIndex"));
            scriptProperties.setProperty('raName',  scriptProperties.getProperty(userID+"raName"));
            scriptProperties.setProperty('inNightSymbol',  scriptProperties.getProperty(userID+"inNightSymbol"));
            scriptProperties.setProperty('dutyPhoneSymbol',  scriptProperties.getProperty(userID+"dutyPhoneSymbol"));
             scriptProperties.setProperty('dutyRoundSymbol',  scriptProperties.getProperty(userID+"dutyRoundSymbol"));
           scriptProperties.setProperty('selectedNotif',  scriptProperties.getProperty(userID+"selectedNotif"));
              scriptProperties.setProperty('sheetNameofCurrentMonth',  scriptProperties.getProperty(userID+"sheetNameofCurrentMonth"));
            
}

function saveNotifSettings(notifNames){
  scriptProperties.setProperty(userID+'selectedNotif', notifNames);
}

// --------Copy Sheets Schedule to Calender---
//first date is already a number (not Date object)
function createCalender(firstDate){
    var monthIndex = currentMonth;
    var yearIndex= currentYear;
    var isLastMonth = false;
    var currentDate=firstDate;
  
    if(firstDate >= 25){
        monthIndex = currentMonth == 0 ? 11 : currentMonth - 1; 
        yearIndex = monthIndex == 11 ? currentYear -1 : currentYear;
        isLastMonth = true;
    }

  
    var calenderArr = [];
        for(let i = 0; i < 6; i++) {
            calenderArr[i] = [];
            for(let j = 0; j < 7; j++) {
                calenderArr[i][j] = new Date(yearIndex, monthIndex, currentDate, 22,50,0);//19= 7pm
                
                //check if reached last date in month
                if(currentDate+1 > monthDates[monthIndex] ){
                    if(isLastMonth ){
                        monthIndex = monthIndex == 11 ? 0 : monthIndex+1;
                        yearIndex =monthIndex ==0 ? yearIndex + 1 : yearIndex;
                        currentDate=1;
                        isLastMonth = false;
                    }
                    else{
                        return calenderArr;
                    }
                }
                
                else{
                    currentDate++;
                }
            }
        }
    return calenderArr;
}


function recordInNights(startCell){
    var currentCell = sheet.getRange(startCell.getRow(), startCell.getColumn() +  Number(scriptProperties.getProperty(userID+'nameToDateGap')));
    var inNightsArr = [];
    var cellValue;
    for(let i =0; i<6; i++){
        inNightsArr[i]=[];
        for(let j =0; j<7; j++){
            cellValue = String(currentCell.getValue());
  
            if(cellValue == scriptProperties.getProperty(userID+'inNightSymbol') && !currentCell.isBlank()){
                inNightsArr[i][j] = raEvents[0];// "In";

            }
            else if(cellValue == scriptProperties.getProperty(userID+'dutyPhoneSymbol') && !currentCell.isBlank()){
                inNightsArr[i][j] = raEvents[1]; //"Duty Phone";
            }
            else if (cellValue == scriptProperties.getProperty(userID+'dutyRoundSymbol') && !currentCell.isBlank()){
                inNightsArr[i][j] = raEvents[2]; //"Duty Round";
            }
            else{
                inNightsArr[i][j] = "out";
             }
            currentCell = sheet.getRange(currentCell.getRow(), currentCell.getColumn() + 1);
        }
        
        
        
        //check if already at last row in sheet
        if( currentCell.getRow()+1 > lastRow){
            return inNightsArr;
        }
        
        //move down to next week 
        currentCell = sheet.getRange(currentCell.getRow()+1, startCell.getColumn());
        while(String(currentCell.getValue()) != scriptProperties.getProperty(userID+'raName') && currentCell.getRow() <= lastRow){
            currentCell = sheet.getRange(currentCell.getRow()+1, currentCell.getColumn());
        }
        //check if already at last week in sheet
        if(currentCell.getRow() > lastRow &&  String(currentCell.getValue()) != scriptProperties.getProperty(userID+'raName')){
            return inNightsArr;
        }
        currentCell = sheet.getRange(currentCell.getRow(), startCell.getColumn() + Number(scriptProperties.getProperty(userID+'nameToDateGap')));

    }
    return inNightsArr;
}


function updateCalender(){
    getDate();
    getSheet();
    deleteRAevents();
    
    var userCalender = CalendarApp.getCalendarById(scriptProperties.getProperty(userID+'calId'));
    var newEvent;
    var firstDate = scriptProperties.getProperty(userID+'firstDate');
    var calender = createCalender(Number(firstDate));
    
    var nameCell = sheet.getRange(String(scriptProperties.getProperty(userID+'nameLocation')));
    var inNights = recordInNights(nameCell);
    
    for(let i = 0; i < calender.length; i++) {
        for(let j = 0; j < calender[i].length; j++) {
            //only create events for the present or future 
            if(calender[i][j] >= today){
                //only create event if RA is "in" or "duty" or "D.R."
                if(inNights[i][j] != "out"){
                    newEvent = userCalender.createEvent(inNights[i][j], calender[i][j], new Date(calender[i][j].getFullYear(), calender[i][j].getMonth(), calender[i][j].getDate(), 23,0,0 ));
                    addNotif(newEvent);
                }
            }
        }
    }
}


function addNotif( myEvent ){
  var selectedNotifTimes = scriptProperties.getProperty(userID+'selectedNotif').split(',');
  if(selectedNotifTimes[0] != "None"){
    for(var i=0; i< selectedNotifTimes.length;i++){
      myEvent.addPopupReminder(Number(selectedNotifTimes[i]));
    }
    
  }

  
    
  
  
}


// -------- Helper Functions --------
  
function checkIfDateObject(value){
    var trueTypeOf = (obj) => Object.prototype.toString.call(obj).slice(8, -1).toLowerCase();
    if(trueTypeOf(value) == 'date'){
        return true;
    }
    return false;
}

function checkIfDate(value){
    var trueTypeOf = (obj) => Object.prototype.toString.call(obj).slice(8, -1).toLowerCase();
    if(trueTypeOf(value) == 'date' || typeof value === 'number'){
        return true;
    }
  
    else if (String(value).length > 2){
        var tmpVal = Number(String(value).slice(0,-2));
        if(!isNaN(tmpVal)){
            return true;
        }
    }
    return false;
}


function printArray(arr){
    var list ="start";
        for(let i = 0; i < arr.length; i++) {
            for(let j = 0; j < arr[i].length; j++) {
                list = list.concat(", "+ arr[i][j]);
            }
            list = list.concat(", Break");
        }
    return list;
}


function deleteRAevents(){
    // for month 0 = Jan, 1 = Feb etc
    // below delete from now to May 1st 2022
    var now = new Date(); 
    var toDate = new Date(2022,4,1,0,0,0);
    var calendar = CalendarApp.getCalendarById(scriptProperties.getProperty(userID+'calId'))
    var events = calendar.getEvents(now, toDate);
    for(var i=0; i<events.length;i++){
        var event = events[i];
        for(var j =0; j<raEvents.length; j++){
            if(event.getTitle() == raEvents[j]){
                event.deleteEvent();
            }
        }
    }
}

function getDate(){
  today = new Date();
  currentYear = today.getFullYear();
  currentMonth = today.getMonth(); //0-11
  monthDates = [31, currentYear % 4 ==0 ? 29:28,31,30,31,30,31,31,30,31,30,31];
}

function getSheet(){
  sheets = ss.getSheets(); //name of months
  var currentSheetName;
  var sheetNameofCurrentMonth = "test";
  

  for(var i=0; i<sheets.length; i++){
    currentSheetName = sheets[i].getName();
    scriptProperties.setProperty(userID+'currentSheetName',currentSheetName.split(" ")[0].toLowerCase());
  
    currentMonthName = monthName[currentMonth].substring(0, currentSheetName.split(" ")[0].length);
    scriptProperties.setProperty(userID+'currentMonthName',currentMonthName);
  
    if(scriptProperties.getProperty(userID+'currentSheetName').toString() == (scriptProperties.getProperty(userID+'currentMonthName').toString())) {
      sheetNameofCurrentMonth = currentSheetName;
    }
    
  }


  scriptProperties.setProperty(userID+'sheetNameofCurrentMonth', sheetNameofCurrentMonth);
  scriptProperties.setProperty('sheetNameofCurrentMonth',  scriptProperties.getProperty(userID+"sheetNameofCurrentMonth"));
  sheet = ss.getSheetByName(sheetNameofCurrentMonth);
  lastRow = sheet.getLastRow();
  lastColumn = sheet.getLastColumn();
}

function deleteProporties(){
  scriptProperties.deleteAllProperties();
  getDate();
  getSheet();
  scriptProperties.setProperty(userID+'selectedNotif', "None");
}

