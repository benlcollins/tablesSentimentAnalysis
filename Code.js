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
const TABLE_NAME = ''; // <-- enter your google tables table ID here

/**
 * doPost webhook to catch data from Google Tables
 */
function doPost(e) {
  
  if (typeof e !== 'undefined') {

    // parse data
    const data = JSON.parse(e.postData.contents);

    // get the id and description
    const rowId = data.id
    const description = data.description;

    // analyze sentiment
    const sentiment = analyzeFeedback(description); // [nlScore,nlMagnitude,emotion]
    
    // combine arrays
    const sentimentArray = [rowId,description].concat(sentiment);

    // send score back to Google Tables
    const rowName = 'tables/' + TABLE_NAME + '/rows/' + rowId;
    const sentimentValues = {
      'Sentiment Score': sentiment[0],
      'Sentiment Magnitude': sentiment[1],
      'Sentiment Tag': sentiment[2]
    };
    Area120Tables.Tables.Rows.patch({values: sentimentValues}, rowName);

    return null;
  }
}

/**
 * Get each new row of form data and retrieve the sentiment 
 * scores from the NL API for text in the feedback column.
 */
function analyzeFeedback(description) {

  if (description !== '') {
      
      // call the NL API
      const nlData = retrieveSentiment(description);
      nlMagnitude = nlData.documentSentiment.magnitude ? nlData.documentSentiment.magnitude : 0; // set to 0 if nothing returned by api
      nlScore = nlData.documentSentiment.score ? nlData.documentSentiment.score : 0; // set to 0 if nothing returned by api
      //console.log(nlMagnitude);
      //console.log(nlScore);
    }
    else {
      
      // set to zero if the description is blank
      nlMagnitude = 0;
      nlScore = 0;

    }

    // turn sentiment numbers into tags
    let emotion = '';
    
    // happy
    if (nlScore > 0.5) { 
      
      if (nlMagnitude > 2) { emotion = 'Super happy!'; } // higher magnitude gets higher emotion tag
      else { emotion = 'Happy'; }

    }
    
    // satisfied
    else if (nlScore > 0) {  emotion = 'Satisfied'; }

    // frustrated
    else if (nlScore < 0 && nlScore >= -0.5) { emotion = 'Frustrated'; }

    // angry
    else if (nlScore < -0.5) { 
      
      if (nlMagnitude > 2) { emotion = 'Super angry!'; } // higher magnitude gets higher emotion tag
      else { emotion = 'Angry'; }
    
    }

    // if score is 0
    else { emotion = 'No opinion' }

    return [nlScore,nlMagnitude,emotion];
}

/**
 * Calls Google Cloud Natural Language API with string from Tables
 * @param {String} description The string from row in Tables
 * @return {Object} the entities and related sentiment present in my string
 */
function retrieveSentiment(description) {
  
  //console.log(description);

  const apiEndpoint = 'https://language.googleapis.com/v1/documents:analyzeSentiment?key=' + API_KEY;
  
  // Create our json request, w/ text, language, type & encoding
  const nlData = {
    document: {
      language: 'en-us',
      type: 'PLAIN_TEXT',
      content: description
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