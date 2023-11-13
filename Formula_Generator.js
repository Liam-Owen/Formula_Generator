
function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setWidth(300)
      .setTitle('Formula Generator');
  SpreadsheetApp.getUi().showSidebar(html);
}

function generateFormula(userInput) {
  // Use the GPT-3 API to generate a formula based on the user's input.
  var generatedFormula = generateFormulaUsingGPT3(userInput);

  return generatedFormula;
}

function saveUserInput(userInput) {
  // Save the user input in the script-level variable (if needed).
  userInput = userInput;
}

function getUserInput() {
  // Retrieve the most recent user input (if needed).
  return userInput;
}


function generateFormulaUsingGPT3(userInput) {
  // Set up the API endpoint URL and your API key.
  var apiUrl = 'https://api.openai.com/v1/chat/completions';
  var apiKey = '*****';

  // Set up the request payload.
  var requestData = {
    model: "gpt-3.5-turbo",
  messages: [
    {"role": "system", "content": "You provide the user with google sheets formulas based on their request. You are to only return the formulas and no other text"},
    {"role": "user", "content": userInput}
  ]

  };

  // Set up the HTTP headers.
  var headers = {
    'Authorization': 'Bearer ' + apiKey,
    'Content-Type': 'application/json',
  };

  // Make a POST request to the API.
  var options = {
    'method' : 'post',
    'headers' : headers,
    'payload' : JSON.stringify(requestData)
  };

  var response = UrlFetchApp.fetch(apiUrl, options);

  // Parse the response to extract the generated formula.
  var data = JSON.parse(response.getContentText());
  var generatedFormula = data.choices[0].message.content; // Adjust this based on your API response structure.

  return generatedFormula;
}



function applyFormulaToSheet(formula) {
  // Apply the generated formula to a specific cell in your Google Sheets.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell = sheet.getRange("A1");
  cell.setValue(formula);
}







