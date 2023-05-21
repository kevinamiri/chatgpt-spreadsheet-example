
// Predefined API Key and Prompt
const API_KEY = "OPENAI_API_KEY";
const PROMPT = "Provide a single-sentence summary of the given text.";
// Spreadsheet ID
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Function to add a menu item for translation
function onOpen() {
  const spreadsheet = SpreadsheetApp.getActive();
  const menuItems = [
    { name: "Translate Text", functionName: "performTranslation" },
  ];
  spreadsheet.addMenu("Custom Functions", menuItems);
}

// Function to get settings
function getSettings() {
  return {
    apiKey: API_KEY,
    prompt: PROMPT,
  };
}

// Function to perform translation using OpenAI
async function performTranslation() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const firstRow = range.getRow();
  const lastRow = range.getLastRow();
  const columnIndex = range.getColumn();
  sheet.insertColumnAfter(columnIndex);

  let emptyRowCount = 0;

  for (let i = firstRow; i <= lastRow; i += 5) {
    for (let j = i; j < i + 5 && j <= lastRow; j++) {
      const cell = sheet.getRange(j, columnIndex);
      const text = cell.getValue();

      if (text === "") {
        emptyRowCount++;
        if (emptyRowCount >= 5) {
          break; // Stop translating when encountering five consecutive empty rows
        }
      } else {
        emptyRowCount = 0; // Reset the empty row count when encountering a non-empty row
        await translateRow(sheet, columnIndex, j);
      }
    }
  }
}

// Function to translate a single row
function translateRow(sheet, columnIndex, rowIndex) {
  return new Promise((resolve, reject) => {
    const cell = sheet.getRange(rowIndex, columnIndex);
    const text = cell.getValue();

    if (text === "") {
      resolve(null); // No translation needed
    } else {
      const translatedText = fetchTranslation(text);

      if (translatedText) {
        sheet.getRange(rowIndex, columnIndex + 1).setValue(translatedText);
        resolve(translatedText);
      } else {
        Logger.log("Error: Translation failed at row " + rowIndex + ".");
        reject(new Error("Translation failed at row " + rowIndex + "."));
      }
    }
  });
}

// Function to fetch translation from OpenAI
function fetchTranslation(text) {
  const settings = getSettings();
  if (!settings.apiKey) {
    Logger.log(
      "Error: API key not set. Set the API key in the 'API_KEY' constant."
    );
    return null;
  }

  const url = "https://api.openai.com/v1/chat/completions";
  const bodyObj = {
    max_tokens: 200,
    model: "gpt-3.5-turbo",
    messages: [
      { role: "system", content: settings.prompt },
      { role: "user", content: text },
    ],
    temperature: 0.5,
    top_p: 1,
    presence_penalty: 0.0,
    frequency_penalty: 0.0,
    n: 1,
    stop: ["---"],
  };

  const options = {
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + settings.apiKey,
    },
    payload: JSON.stringify(bodyObj),
    method: "POST",
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() === 200) {
    const body = JSON.parse(response.getContentText());
    const choices = body.choices;
    const contents = choices.map(choice => choice.message.content);
    return contents[0];
  } else {
    Logger.log("Error: " + response.getContentText());
    return null;
  }
}
