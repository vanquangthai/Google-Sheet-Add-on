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
  SpreadsheetApp.getUi().showModalDialog(html, 'Enter Hunter.io API Key');
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





function FindEmail(firstNameCell, lastNameCell, domainCell) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // Convert arguments to strings if not already
  firstNameCell = String(firstNameCell);
  lastNameCell = String(lastNameCell);
  domainCell = String(domainCell);
  
  // Parse cell references
  var firstName = String(sheet.getRange(firstNameCell).getValue());
  var lastName = String(sheet.getRange(lastNameCell).getValue());
  var domain = String(sheet.getRange(domainCell).getValue());
  
  
  var firstName = String(sheet.getRange(firstNameCell).getDisplayValue());
  var lastName = String(sheet.getRange(lastNameCell).getDisplayValue());
  var domain = String(sheet.getRange(domainCell).getDisplayValue());
  

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


function testFindEmailRange() {
  
  var result =findEmail("A2:C2");
  Logger.log(result);
}

function testFindEmail() {
  
  var result = FindEmail("A2", "B2","C2");
  Logger.log(result);
}
