
// You do NOT need to update these functions if the template is changed.
// Only the headers of the Backlog need to correspond with Jira fields

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet(); 
}

function getBacklogSheet() {
  return getSpreadsheet().getSheetByName("Backlog");
}

function getTemplateSheet() {
  return getSpreadsheet().getSheetByName("Card Template");
}

function getCardSheet() {
  return getSpreadsheet().getSheetByName("Generated Cards");
}

function getJiraConfigSheet() {
  return getSpreadsheet().getSheetByName("Jira Config");
}
//END: Get sheets

// START: Get range within sheets
function getRangeHeight(range) {
  var sheetRowCount = range.getHeight();
  var pixelSize = 0;
  for (var i = 1; i <= sheetRowCount; i++){
  	var rowHeight = range.getSheet().getRowHeight(i);
    pixelSize += rowHeight;
  };
  return pixelSize;
}

function getHeadersRange(backlog) {
  return backlog.getRange(1, 1, 1, backlog.getLastColumn());
}

function getItemsRange(backlog) {
  var numRows = backlog.getLastRow() - 1;
  return backlog.getRange(2, 1, numRows, backlog.getLastColumn());
}

function getSelectedItemsRange(backlog) {
  var range = getSpreadsheet().getActiveRange();
  var startRow = range.getRowIndex();
  var rows = range.getNumRows();
  
  if (startRow < 2 ) { 
    startRow = 2; 
    rows = (rows > 1 ? rows-1 : rows);
  }
  
  return backlog.getRange(startRow, 1, rows, backlog.getLastColumn());
}
// END: Get range within sheets

// START: Set dimensions columns in sheet
function setColumnWidthTo(cardSheet, templateRange) {
  var templateSheet = getTemplateSheet();
  var max = templateRange.getLastColumn() + 1;
  
  for (var i = 1; i < max; i++) {
    var currentWidth = templateSheet.getColumnWidth(i);
    cardSheet.setColumnWidth(i, currentWidth);
  }
}
// END: Set dimensions columns in sheet

/* Get backlog items as objects with property name and values from the backlog. */
function getBacklogItems(selectedOnly) {
  var backlog = getBacklogSheet();
  var rowsRange = (selectedOnly ? getSelectedItemsRange(backlog) : getItemsRange(backlog));
  var rows = rowsRange.getValues();
  var headers = getHeadersRange(backlog).getValues()[0];
  
  var backlogItems = [];
  
  for (var i = 0; i < rows.length; i++) {
    var backlogItem = {};
    
    for (var j = 0; j < rows[i].length; j++) {
      backlogItem[headers[j]] = rows[i][j];
    }
    backlogItems.push(backlogItem);
  }
  return backlogItems;
}

function assertCardSheetExists() {
  if (getCardSheet() == null) {
    getSpreadsheet().insertSheet("Generated Cards", 2);
    Browser.msgBox("The 'Cards' sheet was missing and has now been added. Please try again.");
    return false;
  }
  return true;
}

function createCardsFromBacklog() {
  if (!assertCardSheetExists()) {
    return;
  }
  var backlogItems = getBacklogItems(false);
  createCards(backlogItems);
}

function createCardsFromSelectedRowsInBacklog() {
  if (!assertCardSheetExists()) {
    return;
  }
  if (getBacklogSheet().getName() != SpreadsheetApp.getActiveSheet().getName()) {
    Browser.msgBox("The Backlog sheet need to be active when creating cards from selected rows. Please try again.");
    return;
  }
  var backlogItems = getBacklogItems(true);
  createCards(backlogItems);
}

function getHeadingsFromBacklogSheet(){
  var ss = getBacklogSheet();
  var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  return headings;
}

function createCards(backlogItems) {
  var headings = getHeadingsFromBacklogSheet();
  var templateVariableMap = scanCardTemplateForHeadings(headings);
  var startRow = 1;
  var startColumn = 1;
  var numberOfTemplateRows = getTemplateSheet().getLastRow();
  var numberOfTemplateCols = getTemplateSheet().getLastColumn();
  var template = getTemplateSheet().getRange(1, 1, numberOfTemplateRows, numberOfTemplateCols);  
  var templateSize = getRangeHeight(template); 
  var printPageHeight = getJiraConfigSheet().getRange("B6").getValue();
  var cardCountOnPage = Math.floor(printPageHeight/templateSize);
  var remainderPageSize = printPageHeight%templateSize; 
  var remainderCellsNeeded = 0;
  var cardSheet = getCardSheet();
  var endOfPage = false;

  if (remainderPageSize>0){
    remainderCellsNeeded = Math.floor(backlogItems.length/cardCountOnPage);
  } else {
    remainderCellsNeeded = 0;
  }
  
  initializeCardSheet(cardSheet, template);

  for (var i = 0; i < backlogItems.length; i++) {
    cardSheet.insertRows(startRow, numberOfTemplateRows);
    
    for (var currentRow = 1; currentRow <= numberOfTemplateRows; currentRow++) {
	  	if ( (i == 0) && (currentRow == 1) ){
	        cardSheet.setRowHeight(startRow, 1); //spacer = 0 this time!
	  	} else if (endOfPage && currentRow == 1){
        cardSheet.setRowHeight(startRow, remainderPageSize);
      } else {
        var currentHeight = getTemplateSheet().getRowHeight(currentRow);
        cardSheet.setRowHeight(startRow + currentRow - 1, currentHeight);
      }
    }

    var card = cardSheet.getRange(startRow, startColumn, numberOfTemplateRows, numberOfTemplateCols);
    template.copyTo(card);
    populateCard(card, headings, templateVariableMap, backlogItems[i]);
    
    if (((i+1)%cardCountOnPage)==0) {
      endOfPage = true;      
    } else {
      endOfPage = false;      
    }

    startRow += numberOfTemplateRows;
  }
  Browser.msgBox("Done!");
}

function initializeCardSheet(cardSheet, template){
  cardSheet.clear();
  if (cardSheet.getMaxRows() > 1)
    cardSheet.deleteRows(1, cardSheet.getMaxRows()-1);
  
  setColumnWidthTo(cardSheet, template);	
}

function setCardRowHeightsTo(cardSheet, numberOfTemplateRows) {
  for (var currentRow = 1; currentRow <= numberOfTemplateRows; currentRow++) {
    var currentHeight = getTemplateSheet().getRowHeight(currentRow);
    cardSheet.setRowHeight(currentRow, currentHeight);
  }
}

function populateCard(card, headings, templateVariableMap, backlogItem){
	for (var x = 0; x < headings.length; x++) {
		for (var z = 0; z < templateVariableMap[headings[x]].length; z++){
		  var col = templateVariableMap[headings[x]][z][0];
		  var row = templateVariableMap[headings[x]][z][1];
		  var val = backlogItem[headings[x]];
		  card.getCell(col, row).setValue(val);
		}
	}	
}

function scanCardTemplateForHeadings(headings){
  var headingCoords = [];

  for (var i = 0;i < headings.length;i++) {
    if (headings[i] !== "") {      
      var foundCells = find("<"+headings[i]+">", getTemplateSheet());
      headingCoords[headings[i]] = foundCells;
    }
  }
  return headingCoords; 
}

function find(value, templateSheet) {
  var range = templateSheet.getRange(1, 1, templateSheet.getLastRow(), templateSheet.getLastColumn());

  var data = range.getValues();
  var retVal = [];
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (data[i][j] == value) {
        retVal.push([i+1, j+1]); 
      }
    }
  }
  return retVal;
}
