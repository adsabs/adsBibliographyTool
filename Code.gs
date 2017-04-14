//runs when the spreadsheet is open in order to create the menu items
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ADS Bibliography Maker')
      .addItem('Initialize', 'initialize')
      .addItem('Affiliation Query', 'getAffResults')
      .addItem('Author Query', 'getAuthorResults')
      .addItem('Make Library', 'makeLibrary')
      .addToUi();
}


//create input sheets for the user
function initialize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.insertSheet("Affiliations");
  var affiliationsheet = SpreadsheetApp.getActiveSheet();
  var affHeaders = [["Affiliation","Ambiguity","Exclude","Start Year","Start Month","End Year","End Month", "API Key"]]
  datarange = affiliationsheet.getRange(1,1,1,affHeaders[0].length);
  datarange.setValues(affHeaders);
  
  ss.insertSheet("AuthorNames");
  var authorsheet = SpreadsheetApp.getActiveSheet();
  var authorHeaders = [["Identifier", "Identifier Type", "Author Names"]];
  datarange = authorsheet.getRange(1,1,1,authorHeaders[0].length);
  datarange.setValues(authorHeaders);
  
  ss.insertSheet("BibcodesForNewLibrary");
  var newlibrarysheet = SpreadsheetApp.getActiveSheet();
  var newlibraryHeaders = [["Bibcode List"]];
  datarange = newlibrarysheet.getRange(1,1,1,newlibraryHeaders[0].length);
  datarange.setValues(newlibraryHeaders);
}


//create date strings
//used in getResults
function makeDate() {
  var d = new Date();
  var day = d.getDate();
  var year = d.getFullYear();
  var hours = d.getHours();
  var minutes = d.getMinutes();
  if (minutes < 10) {
    minutes = "0" + minutes.toString();
  }
  else {
    mintues = minutes.toString();
  }
  var month = d.getMonth() + 1;
  var fulldate = year.toString() + "-" + month.toString() + "-" 
                   + day.toString() + " " + hours.toString() + ":" + minutes;
  return fulldate;
}
  

//define API variables
//used in getAffResults
function APIQuery(affquery, daterange, token) {
  var api_url = 'https://api.adsabs.harvard.edu/v1/' ;
  var headers = {
    "Authorization" : 'Bearer ' + token
  }
  var options = {
    "headers" : headers
  }
  var query = 'q=aff:%22' + affquery + '%22&fq=pubdate:%5B' + daterange
      +  '%5D&fl=bibcode,title,author,aff' + '&rows=2000' ;
  var query_url = api_url + 'search/query?' + query;
  Logger.log(query_url);
  var requeststring = UrlFetchApp.getRequest(query_url, options);
  Logger.log("request string: " + requeststring);
  var response = UrlFetchApp.fetch(query_url, options);
  var response_json = JSON.parse(response.getContentText());
  return [response_json.response.docs, query_url]
  }


//define the dialog that appears when the script is run
//used in getAffResults()
function infoDisplayDialog(authors, affiliations, daterange) {
  var numberofauthors = authors.length - 1;
  var numberofaffiliations = affiliations.length - 1;
  SpreadsheetApp.getUi()
  .alert('You are running the query with: \n ' + numberofauthors.toString()    + ' authors, \n '
                                        + numberofaffiliations.toString()      + ' affiliation strings, \n and \n'
                                        + daterange + ' as your date range \n' + 'at ' + makeDate());
}


//main affiliation search function
function getAffResults() {
//select the spreadsheet, and import its sheet data as arrays
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var affiliationsheet = ss.getSheetByName("Affiliations");
  var affiliationdata = affiliationsheet.getDataRange().getValues();
  var authorsheet = ss.getSheetByName("AuthorNames");
  var authordata = authorsheet.getDataRange().getValues();

//get daterange and API token as variables  
  var querydaterange = affiliationdata[1][3].toString() + '-' + affiliationdata[1][4].toString() + ' TO ' 
                       + affiliationdata[1][5].toString() + '-' + affiliationdata[1][6].toString();
  var APItoken = affiliationdata[1][7];

//populate the authorlist array with author names and their variants by separating the NameString column
  var verifiedauthorlist = []
  for (var i = 1; i < authordata.length; i++) {
    var authornamevariants = authordata[i][2].split("|");
    verifiedauthorlist = verifiedauthorlist.concat(authornamevariants);
  }
  
//populate the exclusion list  
  var excludedaffiliationlist = []
  for (var i = 1; i < affiliationdata.length; i++) {
    var affiliationexcludevariants = affiliationdata[i][2].split("|");
    excludedaffiliationlist = excludedaffiliationlist.concat([affiliationexcludevariants]);
  }

//run dialog function 
  infoDisplayDialog(authordata,affiliationdata,querydaterange);
  
//add new output sheets, and create variables for them
  ss.insertSheet(makeDate() + ' Affiliation Search',0);
  var outputsheet = SpreadsheetApp.getActiveSheet();

 //define the arrays that will be used to construct the output sheets
    var outputrows = [];
  
//loop through affiliation strings and perform ADS search on each 
  for (i=1; i < affiliationdata.length; i++) {
    var currentaffiliationquery = affiliationdata[i][0]
    var currentexclusionset = excludedaffiliationlist[i-1]
    var currentaffiliationuncertainty = affiliationdata[i][1]

//use the APIQuery function defined above to get response data
    var APIQueryOutput = APIQuery(currentaffiliationquery, querydaterange, APItoken);
    var BBBdata = APIQueryOutput[0];
    var currentaffiliationqueryURL = APIQueryOutput[1];
 //loop through the rows in the result set    
    for (j = 0; j < BBBdata.length; j++) {
      var currentaffiliationresults = BBBdata[j].aff;
      var matchingaffiliations = [];
      var pairedauthors = [];
      var verifiedauthor = "not verified";
      var excludedaffiliation = "not excluded";
      
  //loop through the individual affiliations in each affiliation result set     
      for (k = 0; k < currentaffiliationresults.length; k++) {
        var singleaffiliationresult = currentaffiliationresults[k].toLowerCase();
        var matchaff = currentaffiliationquery.toLowerCase();

   //add the affiliation that matched and the paired author      
        if (singleaffiliationresult.indexOf(matchaff) != -1) {
          matchingaffiliations = matchingaffiliations.concat(singleaffiliationresult);
          pairedauthors = pairedauthors.concat(BBBdata[j].author[k]);
          for (m = 0; m < currentexclusionset.length; m++){  
            if (singleaffiliationresult.indexOf(currentexclusionset[m].toLowerCase()) != -1) {
              if (currentexclusionset[m] != "") {
                excludedaffiliation = "excluded"
              }
            }
          }  
        }
      }
      
      for (k = 0; k < pairedauthors.length; k++) {
        for (m = 0; m < verifiedauthorlist.length; m++) {
          var singleverifiedauthor = verifiedauthorlist[m].toLowerCase();
          var matchauthor = pairedauthors[k].toLowerCase();
          if (singleverifiedauthor.indexOf(matchauthor) != -1) {
            verifiedauthor = "verified"
          }
        }  
      }
      
   //add result to the output array
      
      outputrows.push([BBBdata[j].bibcode, "https://ui.adsabs.harvard.edu/#abs/" + BBBdata[j].bibcode + "/abstract", BBBdata[j].title[0], currentaffiliationquery, 
                       currentaffiliationuncertainty, currentaffiliationqueryURL, matchingaffiliations.join("|"), 
                       excludedaffiliation, pairedauthors.join("|"), verifiedauthor, BBBdata[j].author.length]);
    }
  }
    
  var colHeaders = [["bibcode", "itemLink", "title", "affiliationQuery", "affiliationUncertainty", "affiliationQueryURL", "affiliationMatch", "affiliationExclusion", "pairedAuthors", "verifiedAuthor", "numberOfAuthors"]];

  outdatarange = outputsheet.getRange(1,1,1,colHeaders[0].length);
  outdatarange.setValues(colHeaders);
  outdatarange = outputsheet.getRange(2,1,outputrows.length,colHeaders[0].length);
  outdatarange.setValues(outputrows);
  
  SpreadsheetApp.getUi()
  .alert('Affiliation Query Complete');
}


//main author search function
function getAuthorResults() {
//select the spreadsheet, and import its sheet data as arrays
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var affiliationsheet = ss.getSheetByName("Affiliations");
  var affiliationdata = affiliationsheet.getDataRange().getValues();
  var authorsheet = ss.getSheetByName("AuthorNames");
  var authordata = authorsheet.getDataRange().getValues();

//get daterange and API token as variables  
  var querydaterange = affiliationdata[1][3].toString() + '-' + affiliationdata[1][4].toString() + ' TO ' 
                       + affiliationdata[1][5].toString() + '-' + affiliationdata[1][6].toString();
  var APItoken = affiliationdata[1][7];

//populate the authorlist array with author names
  var ADSauthorlist = [];
  for (var i = 1; i < authordata.length; i++) {
    var authornamevariants = authordata[i][2].split("|");
//take just the first version of the author name
    // first version should be most complete to leverage ADS search
    ADSauthorlist = ADSauthorlist.concat(authornamevariants[0]);
  }
  
//create author last name list
  var ADSauthorsurnames = [];
  for (var i = 0; i < ADSauthorlist.length; i++) {
    var surname = ADSauthorlist[i].split(",")[0];
    ADSauthorsurnames = ADSauthorsurnames.concat(surname + ", " + ADSauthorlist[i].split(",")[1].charAt(1));
  }
  
//  logger.Log(ADSauthorsurnames)
  
  var numberofauthors = authordata.length - 1;
                   
//display basic search info
  var numberofaffiliations = affiliationdata.length - 1;
  SpreadsheetApp.getUi()
  .alert('You are running the query with: \n ' + numberofauthors.toString()         + ' authors, \n '
                                        + numberofaffiliations.toString()           + ' affiliation strings, \n and \n'
                                        + querydaterange + ' as your date range \n' + 'at ' + makeDate());
  
//add new output sheets, and create variables for them
  ss.insertSheet(makeDate() + ' Authors Search',0);
  var authorsheetout = SpreadsheetApp.getActiveSheet();

//define API variables
  var api_url = 'https://api.adsabs.harvard.edu/v1/' ;
  var headers = {
    "Authorization" : 'Bearer ' + APItoken
  };
  var options = {
    "headers" : headers
  };

 //define the arrays that will be used to construct the output sheets
    var outputrows = [];
  
//loop through author strings and perform ADS search on each 
  for (i=0; i < ADSauthorlist.length; i++) {
    var currentauthorquery = ADSauthorlist[i];
    var currentauthorsurname = ADSauthorsurnames[i].toLowerCase();
    var othercurrentauthorsurname = currentauthorquery.split(",")[0].toLowerCase() + ", " + currentauthorquery.split(",")[1].charAt(1).toLowerCase();
    if (currentauthorsurname != othercurrentauthorsurname) {
      Logger.log("surname mismatch: " + currentauthorsurname + " != " + othercurrentauthorsurname);
    }
    var query = 'q=author:%22' + currentauthorquery + '%22&fq=pubdate:%5B' + querydaterange + '%5D'
      +  '&fl=bibcode,title,author,aff' + '&rows=2000' ;
    var query_url = api_url + 'search/query?' + query;
    Logger.log("query url")
    Logger.log(query_url);
    var response = UrlFetchApp.fetch(query_url, options);
    var response_json = JSON.parse(response.getContentText());
    var BBBdata = response_json.response.docs;
      
 //loop through the rows in the result set    
    for (j = 0; j < BBBdata.length; j++) {
      var currentauthorresults = BBBdata[j].author
      var matchingauthors = [];
      var pairedaffiliations = [];
      
  //loop through the individual authors in each author result set     
      for (k = 0; k < currentauthorresults.length; k++) {
        var singleauthorresult = currentauthorresults[k].toLowerCase();

   //add the affiliation that matched and the paired author      
        if (singleauthorresult.indexOf(currentauthorsurname) != -1) {
          matchingauthors = matchingauthors.concat(singleauthorresult);
          pairedaffiliations = pairedaffiliations.concat(BBBdata[j].aff[k]);
        }
      }
    
   //add each row to the array
    if (BBBdata == []) {
      outputrows.push(["", "", currentauthorquery, "", "", ""]);
    }
    else {
      outputrows.push([BBBdata[j].bibcode, "https://ui.adsabs.harvard.edu/#abs/" + BBBdata[j].bibcode + "/abstract", BBBdata[j].title[0], currentauthorquery, matchingauthors.join("|"), pairedaffiliations.join("|"), BBBdata[j].author.length]);
     }
    }
  }
  var colHeaders = [["bibcode", "itemLink", "title", "authorQuery", "authorMatch", "pairedAffiliaions", "numberOfArticleAuthors"]];
  
  datarange = authorsheetout.getRange(1,1,1,colHeaders[0].length);
  datarange.setValues(colHeaders);
  datarange = authorsheetout.getRange(2,1,outputrows.length,colHeaders[0].length);
  datarange.setValues(outputrows);
  
  SpreadsheetApp.getUi()
  .alert('Author Query Complete');
}


//define API variables
function APIPush(bibcodes, token) {
  var api_library_url = 'https://api.adsabs.harvard.edu/v1/biblib/libraries' ;
  
  var headers = {
    'Authorization': 'Bearer ' + token
  };
  
  var payload = {
    'name': makeDate() + " Bibliography",
    'description': "Bibliography uploaded to ADS on " + makeDate(),
    'public': false,
    'action': 'add',
    'bibcode': bibcodes
  };
  
  var options = {
    'contentType': 'application/json',
    'Accept': 'text/plain',
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(payload)
//    "payload" : payload
  };
  
  var response = UrlFetchApp.fetch(api_library_url, options);
//  var response_json = JSON.parse(response.getContentText());
//  return response_json.response.docs;
  }


// create a library in ADS
function makeLibrary() {
//select the spreadsheet, and import its sheet data as arrays
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var affiliationsheet = ss.getSheetByName("Affiliations");
  var affiliationdata = affiliationsheet.getDataRange().getValues();
  var bibcodesheet = ss.getSheetByName("BibcodesForNewLibrary");
  var bibcodedata = bibcodesheet.getDataRange().getValues();

  var bibcodelist = []
  for (var i = 1; i < bibcodedata.length; i++) {
    bibcodelist = bibcodelist.concat(bibcodedata[i][0]);
  }
  
//get API token as a variable 
  var APItoken = affiliationdata[1][7];
  
  APIPush(bibcodelist, APItoken);

}