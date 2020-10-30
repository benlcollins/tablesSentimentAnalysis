/**
 * Get data from Google Tables
 * Determine rows to test for sentiment analysis
 * Send data to NL API for testing
 * Paste returned values back into Google Tables
 */

/**
 * global variables
 */
const API_KEY = ''; // <-- enter your google cloud project API key here

/**
 * Get each new row of form data and retrieve the sentiment 
 * scores from the NL API for text in the feedback column.
 */
function analyzeFeedback() {
  
  // get data from the Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sheet1');
  const allRange = sheet.getDataRange();
  const allData = allRange.getValues();
  
  // remove the header row
  allData.shift();
  //console.log(allData);

  allData.forEach(row => {
    const rowId = row[0];
    const description = row[6];
    console.log(rowId + ' ' + description);

    const nlData = retrieveSentiment(description);

    console.log(nlData);
  });

}


/**
 * Calls Google Cloud Natural Language API with cell string from my Sheet
 * @param {String} cell The string from a cell in my Sheet
 * @return {Object} the entities and related sentiment present in my string
 */
function retrieveSentiment(cell) {
  
  //console.log(cell);
  
  const apiEndpoint = 'https://language.googleapis.com/v1/documents:analyzeEntitySentiment?key=' + API_KEY;
  
  // Create our json request, w/ text, language, type & encoding
  const nlData = {
    document: {
      language: 'en-us',
      type: 'PLAIN_TEXT',
      content: cell
    },
    encodingType: 'UTF8'
  };
  
  //  Package all of the options and the data together for the call
  const nlOptions = {
    method : 'post',
    contentType: 'application/json',  
    payload : JSON.stringify(nlData)
  };
  
  //  Try fetching the natural language api
  try {
    
    // return the parsed JSON data if successful
    const response = UrlFetchApp.fetch(apiEndpoint, nlOptions);
    return JSON.parse(response);
    
  } catch(e) {
    
    // log the error message and return null if not successful
    console.log("Error fetching the Natural Language API: " + e);
    return null;
  } 
}
