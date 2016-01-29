/*
##################################################################################
##	Author			:	Miztiik
##	Date   			:	29Jan2016
##	Version			:	0.3
##	Description		:	This script is to used read media year and title from google spreadsheet and fetch the metadata from OMDB through their API and update the sheet
##	Assumptions		:	There is a header row(@Row 1) and the year and title are in columns 4 & 5 respectively
##################################################################################
*/
// Global Variables & Customizable variables
//var targetSheetName = "Yet2Get";
//var searchSheet = Browser.inputBox("Enter the name of the sheet to update:");

// Get the sheet from which we need to read the media title and year to get metadata
var targetSheet = sheetName();
var todaysDate = Utilities.formatDate(new Date(), 'GMT+5:30', 'dd/MM/yyyy HH:mm'),
  mFoundColor = 'Black',
  mDuplicates = 'Blue',
  mNotFoundColor = 'Red',
  mMultipleColor = 'Green',
  mMultipleTitles;

// Measure execution time
var startTime = new Date().getTime();

setHeaderColumn(targetSheet);
readYearTitle();

Logger.log("Fetching Metadata from OMDB & Post Processing took " + (new Date().getTime() - startTime) / 1000 + " Seconds.");

/*
Function to read the movie title and year from the sheet , assuming the previous two cells on the same row contain this data
It checks for last update column and only updates metadata when last updated is empty
*/

function readYearTitle() {
  //The last row in the "sheet" with content (excluding the header row)
  var lastRow = targetSheet.getLastRow() + 1,
    //The last column in the sheet with content
    lastColumn = targetSheet.getLastColumn(),
    //Get the media year and title (assuming they are starting at 2nd Row and 3rd Column, till the end of 'lastRow')
   mTY = targetSheet.getRange(2, getColIndexByName(targetSheet, "Year"), lastRow - 1, 13).getValues();
  
  var mCount = 0;
  
  // Lets fetch metadata for each of the entries one by one
  for (var d = 0; d < mTY.length; d++) {    
    // Continue with search for metadata only if the last update date is empty. Probably in future this has to be checked for old updates.    
    if (mTY[d][12]) continue;    
    // Lets do the metadata search only for items which has value for Title column.
    // http://goo.gl/5sEYSk http://goo.gl/i9wdMK
    if (mTY[d][1]) {
      // Preparing the 'mMultipleTitles' global variable so each time it has the proper prefix, after clearing the cache from previous iteration
      mMultipleTitles = "Multiple Titles Found:";
      searchOmdb(mTY[d][0], mTY[d][1], d + 2);
      mCount ++;
    }
  }
  Logger.log("Total of " + mCount + " rows to be updated!");
  Browser.msgBox("Total of " + mCount + " rows updated in " + (new Date().getTime() - startTime) / 1000 + " Seconds.");
}

/*
Function to search for all media for the given title in OMDB
@param {Number} - Year of release of the movie, ex 1999,1920 etc
@param {String} The title of movie, ex The Matrix, Predestination etc
@param {Number} The row number in the sheet ex 2, 3
*/
function searchOmdb(mYear, mTitle, rID) {
  var data,
  mTitle,
  mYear,
  mfound = 0,
    rows = [];
  
  // Convert the cell value to string and trim the string to remove empty sapces
  mTitle = mTitle.toString().trim();

  // Error Codes for media found variable 'mfound'
  // 1 or higher - Found, 2 - Not able to connect to omDB API, 0 = Default state

  //Hard setting the values for testing 
  var isTest = "n";

  if (isTest == "y") {
    mTitle = "point break",
    //mTitle = "Kicking and Screaming",
    //mTitle = "Furrrrr";
    mYear = "";
  }
  myUrl = "http://www.omdbapi.com/?s=" + mTitle + "&type=movie&y=" + mYear + "&plot=short&r=json";

  var jsonData = UrlFetchApp.fetch(myUrl),
    jsonArray = JSON.parse(jsonData.getContentText());

  if (jsonArray.hasOwnProperty("Search")) {
    // Lets check if there are multiple media with the same title
    // getMatchingTitles returns '0' length object when no exact matches found

    var mArr = getMatchingTitles(jsonArray, 'Title', mTitle);

    /* Five possible cases here
           # 0 - 'No Exact Matches' found although there are results from OMDB (Array Length from 'getMatchingTitles' will be 0)
           # 1 - 'Exact Match' of Title or both Title & Year (Array Length from 'getMatchingTitles' will be 1)
           # 2 & > - 'Multiple exact matches' found in the result (Array Length from 'getMatchingTitles' will be >2)
           # 'No media found' for the given title
           # 'Not able to connect' - Yet to be written      
        */

    if (mArr.length == 0) {
      // When there is no 'exact match' of title with the media found in OMDB
      rows.push([mMultipleTitles]);
      pushData("M", targetSheet, rID, rows);
    } else if (mArr.length == 1) {
      getOmdbMetaData(mArr, rID);
      mfound++;
    } else if (mArr.length > 1) {
      var dupTitles = 'Duplicate Titles Found:';
      for (var obj in mArr) {
        if (mArr.hasOwnProperty(obj)) {
          if (mArr[obj].hasOwnProperty('Year') && mArr[obj].hasOwnProperty('Title')) {
            dupTitles = dupTitles.concat(' "' + mArr[obj]['Year'] + ':' + mArr[obj]['Title'] + '"');
          }
        }
      }
      rows.push([dupTitles]);
      mfound++;
      pushData("D", targetSheet, rID, rows);
    }
  } else if (jsonArray.Response == "False") {
    // When no media matching the title in OMDB are found
    rows.push(["Error:Movie not found!"]);
    pushData("N", targetSheet, rID, rows);
  }

  // Another case will need to be written for not able to connect scenario
}

/* Function to get the meta data for the exact matches through the OMDB API using the IMDB ID flag
@param {Array} - An array of movie titles
@param {Number} The row to which the data needs to be written
*/
function getOmdbMetaData(mArr, rID) {
  for (i = 0; i < mArr.length; i++) {
    // Check if the IMDB ID is present in the array
    if (mArr.hasOwnProperty('Year') && mArr.hasOwnProperty('Title')) continue;
    // (mArr.hasOwnProperty(mArr[i].imdbID)) continue;
    var mUrl = "http://www.omdbapi.com/?i=" + mArr[i].imdbID + "&type=movie&y=" + mArr[i].Year + "&plot=short&tomatoes=true&r=json";
    var mData = UrlFetchApp.fetch(mUrl),
      mArray = JSON.parse(mData.getContentText());

    var rows = [];
    // imdbRating	tomatoRating	Metascore	Runtime	Genre	Director	Actors	Language	Plot	imdbID	Timestamp
    rows.push([mArray.Year,
    mArray.Title,
    mArray.imdbRating,
    mArray.tomatoRating,
    mArray.Metascore,
    mArray.Runtime.replace(" min", ""),
    mArray.Genre,
    mArray.Director,
    mArray.Actors,
    mArray.Language,
    mArray.Plot,
      "www.imdb.com/title/" + mArray.imdbID,
    todaysDate]);
    pushData("Y", targetSheet, rID, rows);
  }
}

/*
Function to find the active spreadsheet in which the updates are to be made
*/
function sheetName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssActiveSheet = SpreadsheetApp.getActiveSheet().getName();
    if (ssActiveSheet == "Yet2Get" || ssActiveSheet == "Pokran") {
      return ss.getSheetByName(ssActiveSheet);
    } else {
      return ss.getSheetByName("Yet2Get");      
    }
}

/* 
Function to set the headers to the given sheet
@param {String} Name of the sheet where the headers are to be set.
*/
function setHeaderColumn(mSheet) {
  //mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).setFontWeight('bold');
  mSheet.getRange(1, 1, 1, 16).setFontWeight('bold');
  mSheet.getRange(1, 1, 1, 16).setBackground('#cfe2f3');
  mSheet.getRange(1, 1, 1, 16).setFontColor('Black');
  var mHeaders = [];
  mHeaders = [
    ["Status", "Storage", "MyRating", "Year", "Title", "imdbRating", "tomatoRating", "Metascore", "Runtime", "Genre", "Director", "Actors", "Language", "Plot", "imdbID", "Timestamp"]
  ];
  mSheet.getRange(1, 1, 1, 16).setValues(mHeaders);
}

/*
Function to flush the metadata to the spreadsheets
@param {String} Single character string ex "Y" - Media Found, "N" - Media Not Found , "D" - Duplicates Found
@param {String} Name of the spreadsheet to which the updates are to be written, Ex., "MySheet", "MyMedia"
@param {Number} The row to which the data needs to be written
@param {Array} 1D Array of media Meta Data
*/
function pushData(foundMedia, targetSheet, rID, rows) {
  if (foundMedia == "Y") {
    // Set the color, Update the message
    // Update the Title & Year as well
    targetSheet.getRange(rID, 4, 1, rows[0].length).setFontColor(mFoundColor);
    targetSheet.getRange(rID, 4, 1, rows[0].length).setValues(rows);
  } else if (foundMedia == "N") {
    // Set the color, Clear the cells,Update the 'Error' message
    targetSheet.getRange(rID, 6, 1, rows[0].length).setFontColor(mNotFoundColor);
    targetSheet.getRange(rID, 6, 1, targetSheet.getLastColumn()).clear({contentsOnly: true});
    targetSheet.getRange(rID, 6, 1, rows[0].length).setValue(rows);
  } else if (foundMedia == "D") {
    // Set the color, Clear the cells,Update the 'Error' message
    targetSheet.getRange(rID, 6, 1, rows[0].length).setFontColor(mDuplicates);
    targetSheet.getRange(rID, 6, 1, targetSheet.getLastColumn()).clear({contentsOnly: true});
    targetSheet.getRange(rID, 6, 1, rows[0].length).setValues(rows);
  } else if (foundMedia == "M") {
    // Set the color, Clear the cells,Update the 'Error' message
    targetSheet.getRange(rID, 6, 1, rows[0].length).setFontColor(mMultipleColor);
    targetSheet.getRange(rID, 6, 1, targetSheet.getLastColumn()).clear({contentsOnly: true});
    targetSheet.getRange(rID, 6, 1, rows[0].length).setValues(rows);
  }
}

/* Find the column index using the header name
   @param {String} column - The column header name to get ('Year', 'Title', etc)
   @param {Number} index -  The column number (1, 5, 15)
*/
function getColIndexByName(mSheet, colName) {
  var headers = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0];
  for (i in headers) {
    if (headers[i] == colName) {
      return parseInt(i) + 1;
    }
  }
  return -1;
}


/*
Plaigiarised from @ http://techslides.com/how-to-parse-and-search-json-in-javascript
Function to search an json array and return an array of objects according to key, value, or key and value matching
@param {object} Object of array in which you want to search
@param {key} Key within array
@param {value} Any value the key might contain
*/
function getMatchingTitles(obj, key, val) {
  var objects = [];
  // Need to convert to UpperCase only for 'Alphanumeric' Movie Titles, Movie titles with just numbers like '300' fails.
  var isNum = /^\d+$/.test(val);
  if (!isNum) {
    val = val.toUpperCase();
  }
  for (var i in obj) {
    if (!obj.hasOwnProperty(i)) continue;
    if (typeof obj[i] == 'object') {
      objects = objects.concat(getMatchingTitles(obj[i], key, val));
    } else
    // if key matches and value matches or if key matches and value is not passed (eliminating the case where key matches but passed value does not)
    if (i == key && obj[i].toUpperCase() == val || i == key && val == '') {
      objects.push(obj);
    } else if (obj[i].toUpperCase() == val && key == '') {
      // only add if the object is not already in the array
      if (objects.lastIndexOf(obj) == -1) {
        objects.push(obj);
      }
    } else if (i == key && val != '') {
      // Collect the 'Title's' if the key and value are not matching, so we can show them to user
      mMultipleTitles = mMultipleTitles.concat(' ' + '"' + obj[i] + ':');
    } else if (i == 'Year' && val != '') {
      // Collect the 'Year' if the key and value are not matching, so we can show them to user
      mMultipleTitles = mMultipleTitles.concat(obj[i] + '"');
    }
  }
  return objects;
}
