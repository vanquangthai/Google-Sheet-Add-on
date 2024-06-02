// Function to creat menu Hunter.io in Google Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hunter.io')
    .addItem('Set API Key', 'showApiKeyPrompt')
    .addItem('Show API Key', 'showApiKey')
    .addToUi();
}

// Function to creat HTML dialog in Google Sheet
function showApiKeyPrompt() {
  var html = HtmlService.createHtmlOutputFromFile('dialog')
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'API Key Dialog');
}
// Function to check and set the API key
function checkApiKey(apiKey) {
  if (isValidApiKey(apiKey)) {
    PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
    return 'API key saved successfully!';
  } else {
    return 'Invalid API key. Please try again.';
  }
}

// Function to validate the API key
function isValidApiKey(apiKey) {
  
  return apiKey.length >30; 
}

// Function to show the apiKey in Google Sheet
function showApiKey() {
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  if (apiKey) {
    SpreadsheetApp.getUi().alert('Hunter.io API Key: ' + apiKey);
  } else {
    SpreadsheetApp.getUi().alert('No API Key set. Please set the API Key first.');
  }
}

function setApiKey() {
  return PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  
}


function getSheetValuesAsString(range) {
  // Open the spreadsheet by its ID or URL
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeString = range.toString()    
  // Get the values from the specified range
  var rangeString = sheet.getRange(rangeString);
  var values = rangeString.getValues();
  
  // Convert each value to a string (if not already a string)
  var firstName = values[0][0].toString();
  var lastName = values[0][1].toString();
  var domain = values[0][2].toString();
  
  return { firstName: firstName, lastName: lastName, domain: domain };
}


function FindEmailwithRange(range) {
  
  var { firstName, lastName, domain } = getSheetValuesAsString(range);
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  
  if (!apiKey) {
    return 'API key not set. Please set the API key first.';
  }
  
  var url = `https://api.hunter.io/v2/email-finder?domain=${domain}&first_name=${firstName}&last_name=${lastName}&api_key=${apiKey}`;
  var response = UrlFetchApp.fetch(url);
  
  if (response.getResponseCode() === 200) {
    var data = JSON.parse(response.getContentText());
    
    if (data.data && data.data.email) {
      return data.data.email;
    } else {
      return 'Email not found.';
    }
  } else {
    return 'Failed to fetch data. Response code: ' + response.getResponseCode();
  }
}

function FindEmail(firstName,lastName,domain) {
  var args = Array.prototype.slice.call(arguments); // Convert arguments to an array
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var argStrings = [];

  for (var i = 0; i < args.length; i++) {
    argStrings[i] = args[i].toString();
  }

  // Retrieve the values of the specified cells
  var firstName = sheet.getRange(argStrings[0]).getDisplayValue();
  var lastName = sheet.getRange(argStrings[1]).getDisplayValue();
  var domain = sheet.getRange(argStrings[2]).getDisplayValue();

  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  
  if (!apiKey) {
    return 'API key not set. Please set the API key first.';
  }
  
  var url = `https://api.hunter.io/v2/email-finder?domain=${domain}&first_name=${firstName}&last_name=${lastName}&api_key=${apiKey}`;

  var response = UrlFetchApp.fetch(url);

  if (response.getResponseCode() === 200) {
    var data = JSON.parse(response.getContentText());
    
    if (data.data && data.data.email) {
      return data.data.email;
    } else {
      return 'Email not found.';
    }
  } else {
    return 'Failed to fetch data. Response code: ' + response.getResponseCode();
  }
}
