/**
 * This script's function is to format a spreadsheet of companies
 * 
 * Each company may have many rows depending on how many positions are available 
 *
 * This custom formatting will change the background color of each row depending on what company it belongs to
 *
 * The user will not have to manually change the color, this script will change the colors automatically
 * 
 * 
 * @author Fowad Sohail
 * @version 1/2/19
 */

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]; //get active spreadsheet and the sheet from index in array
var lastRow = sheet.getLastRow();

var range = sheet.getRange('A3:A'); //range -> A column excluding header row


//"#a2c4c9" ---- green/blue
//"#f6b26b" ----- orange


/**
 * This function calls clearColors() and formatRows()
 * It is the function used in this project's triggers
 */
function clearAndFormat() {
  clearColors();
  formatCompanies();
  colorRows();
}

/** This function formats the first column's font, font size and bold
*/
function formatCompanies() {
 range.setFontFamily("Times New Roman");
 range.setFontWeight("bold");
 range.setFontSize(12);
}


/**
 * This function will format the rows in the columns depending on the cell in the first column right above 
 */
function colorRows() {

  sheet.getRange(sheet.getRange(1,1).getRow(), 1, 1, sheet.getLastColumn()).setBackground("#999999"); //header rows color grey
  sheet.getRange(sheet.getRange(2,1).getRow(), 1, 1, sheet.getLastColumn()).setBackground("#999999");

  for(var i = 1; i < lastRow-1; i++) { //iterate through the rows
  
      var currentCell = range.getCell(i, 1); //get the cell (in the first column) for the current row
      var currentRow = sheet.getRange(currentCell.getRow(),1,1,sheet.getLastColumn()); //current row
      
      var prevCell = currentCell; //boundary case --> if i=1, cant get prev cell
      if(i > 1) {
         var prevCell = range.getCell(i-1, 1);
      }
      
      var prevRow = sheet.getRange(prevCell.getRow(),1,1,sheet.getLastColumn());//previous row
      
      
      
      if(!(currentCell.isBlank())) { //current cell is not blank
        //get above row/cell color
        //set row color opposite to that
          var aboveColor = prevRow.getBackground();
        

        if(aboveColor == "#ffffff" || aboveColor == "#b7b7b7" || aboveColor == "#f6b26b") { //above color white, grey or orange
           currentRow.setBackground("#a2c4c9"); //set blue
        }
        
        if(aboveColor == "#a2c4c9") {
         currentRow.setBackground("#f6b26b");
        }
      }
      
      
      if(currentCell.isBlank()) { //empty cells same color as above cells
        currentRow.setBackground(prevRow.getBackground());
      }
    
    
  }
}

/**
* This function sets the background color of the entire spreadsheet to white
*/
function clearColors() {
  var allRange = sheet.getRange('A:Z');
  allRange.setBackground('#ffffff');
}


/**
* This function determines the time needed to travel from here to there
* @param here Address as a string 
* @param there Address as a string
* @return the amount of time it takes to get from here to there, in minutes
* @customfunction
*/
function DISTANCE(here, there) {
 
  var mapObj = Maps.newDirectionFinder();
  mapObj.setOrigin(here);
  mapObj.setDestination(there);
  
  var directions = mapObj.getDirections();
  var getTheLeg = directions["routes"][0]["legs"][0];
  var mins = getTheLeg["duration"]["text"];
  
  return mins;
}
