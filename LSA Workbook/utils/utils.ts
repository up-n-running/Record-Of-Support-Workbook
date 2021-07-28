function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

//ADD A SHARED DRIVE PARAM IN CASE ADDEDITOR FAILS BELOW
function createOrGetChildFolder(parentFolderID, childFolderName, addEditorEmailAddress){

  //parse parameter 3 if none was passed in set to null
  addEditorEmailAddress = (addEditorEmailAddress) ? addEditorEmailAddress : null;

  var parentFolder = DriveApp.getFolderById(parentFolderID);
  var subFolders = parentFolder.getFolders();
  var doesntExists = true;
  var newFolder = null;
  
  // Check if folder already exists.
  while(subFolders.hasNext()){
    var folder = subFolders.next();
    
    //If the name exists return the id of the folder
    if(folder.getName() === childFolderName){
      doesntExists = false;
      newFolder = folder;
      return newFolder;
    }
  }
  
  //If the name doesn't exists, then create a new folder
  if(doesntExists == true){
    //If the file doesn't exists
    newFolder = parentFolder.createFolder(childFolderName);

    //now we have created the folder if they passed in an addEditorEmailAddress param then set their permission to be an editor
    if( addEditorEmailAddress != null ) {
      newFolder.addEditor( addEditorEmailAddress );
    }

    return newFolder;
  }
}


function getFileLastModifiedDate(fileId) {
  if(fileId) {
    return DriveApp.getFileById(fileId).getLastUpdated();
  }
  else {
    return null;
  }
}

function catchErrorToString (err) {
  var errInfo = ""; 
  for (var prop in err)  {  
    errInfo += "  error property: "+ prop+ "\n    value: ["+ err[prop]+ "]\n"; 
  } 
  errInfo += "  err.toString(): [" + err.toString() + "]"; 
  return errInfo;
}

function deduplicateArray(array, removeBlanks) {
  var outArray = [];
  array.sort();
  outArray.push(array[0]);
  for(var n in array){
    if(( !removeBlanks || array[n] != "" ) && outArray[outArray.length-1]!=array[n]){
      outArray.push(array[n]);
    }
  }
  return outArray;
}

function removeElementsFromArray(masterArray, elementsToRemove) { 
  return masterArray.filter(function(ele){ 
    return !elementsToRemove.includes( ele ); 
  });
}

function CloneArray_ShallowCopy( sourceArray ) {
  var destArray = new Array();
  for (let i = 0; i < sourceArray.length; i++) {
    destArray[i] = sourceArray[i];
  }
  return destArray;
}

function isValidInteger( str: any ) {
  var n = Math.floor(Number(str));
  return n !== Infinity && String(n) === str && n >= 0;
}

function isValidDate( potentialDateCellValue: any ) {
  try{
    let newDate = new Date(potentialDateCellValue);
    let timeCheck = newDate.getTime();
    return (timeCheck === timeCheck);  //if not date timeCheck is NaN which doesnt equal itself!
  }
  catch( e ) {
    return false;
  }
}

//mon = 1, sun = 7
function dayNameToNumber( dayName: string, offset?: number ): number {
  offset = offset ? offset : 0;
  let dayNumber = -1;
  switch (dayName) {
    case "Monday":    dayNumber = 1; break;
    case "Tuesday":   dayNumber = 2; break;
    case "Wednesday": dayNumber = 3; break;
    case "Thursday":  dayNumber = 4; break;
    case "Friday":    dayNumber = 5; break;
    case "Saturday":  dayNumber = 6; break;
    case "Sunday":    dayNumber = 7; break;
  }
  dayNumber = dayNumber < 0 ? dayNumber : ( ( ( dayNumber - 1 ) + offset + 7 ) % 7 ) + 1;
  return dayNumber;
}

//mon = 1, sun = 7
function dayNumberToName( dayNumber: number, offset?: number ): string {
Logger.log( "B4 dayNumber = " + dayNumber );
Logger.log( "B4 offset = " + offset );
  offset = offset ? offset : 0;
  let dayName: string|null = null;
  dayNumber = dayNumber < 0 ? dayNumber : ( ( ( dayNumber - 1 ) - offset + 7 ) % 7 ) + 1;
Logger.log( "AFTA dayNumber = " + dayNumber );
Logger.log( "AFTA offset = " + offset );
  switch (dayNumber) {
    case 1: dayName = "Monday"; break;
    case 2: dayName = "Tuesday"; break;
    case 3: dayName = "Wednesday"; break;
    case 4: dayName = "Thursday"; break;
    case 5: dayName = "Friday"; break;
    case 6: dayName = "Saturday"; break;
    case 7: dayName = "Sunday"; break;
  }
  return dayName;
}

function arrayToString(arr) {
let str = "";
  for (let item of arr) {
    if (Array.isArray(item)) str += arrayToString(item);
    else str += item + ", ";
  }
  return str;
}

function create2DPrePopulatedArray( rows: number, cols: number, value: string): any[][] {
  var arr = Array(rows);
  for (let ri = 0; ri < rows; ri++) {
      arr[ri] = Array(cols).fill(value);
  }
  return arr;
}