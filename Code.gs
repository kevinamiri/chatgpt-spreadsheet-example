
// To display the input form dialog
function showSettingsDialog() {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('settings.html')
        .setWidth(300)
        .setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Settings');
}


// To process the input form data
function saveSettings(apiKey, prompt) {
    var scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('API_KEY', apiKey);
    scriptProperties.setProperty('PROMPT', prompt);
}

function getSettings() {
    var scriptProperties = PropertiesService.getScriptProperties();
    return {
        apiKey: scriptProperties.getProperty('API_KEY'),
        prompt: scriptProperties.getProperty('PROMPT')
    };
}


// Modify the onOpen() function to include a menu item for showing the input form
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActive();
    var menuItems = [
        { name: 'Translate Text', functionName: 'translateTextOpenAI' },
        { name: 'Settings', functionName: 'showSettingsDialog' }
    ];
    spreadsheet.addMenu('Custom Functions', menuItems);
}

function translateTextOpenAI() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getActiveRange();
    var firstRow = range.getRow();
    var lastRow = range.getLastRow();
    var columnIndex = range.getColumn();
    sheet.insertColumnAfter(columnIndex);

    var emptyRowCount = 0;

    for (var i = firstRow; i <= lastRow; i++) {
        var cell = sheet.getRange(i, columnIndex);
        var text = cell.getValue();

        if (text === '') {
            emptyRowCount++;
            if (emptyRowCount >= 5) {
                break; // Stop translating when encountering five consecutive empty rows
            }
        } else {
            emptyRowCount = 0; // Reset the empty row count when encountering a non-empty row
            var translatedText = fetchChatGPT(text);

            if (translatedText) {
                sheet.getRange(i, columnIndex + 1).setValue(translatedText);
            } else {
                Logger.log("Error: Translation failed at row " + i + ".");
            }
        }
    }
}

function fetchChatGPT(text) {
    var settings = getSettings();
    if (!settings.apiKey) {
        Logger.log('Error: API key not set. Set the API key in the Settings menu.');
        return null;
    }

    var url = 'https://api.openai.com/v1/chat/completions';
    var bodyObj = {
        max_tokens: 200,
        model: 'gpt-3.5-turbo',
        messages: [
            { role: 'system', content: settings.prompt },
            { role: 'user', content: text }
        ],
        temperature: 0.5,
        top_p: 1,
        presence_penalty: 0.0,
        frequency_penalty: 0.0,
        n: 1,
        stop: ['---']
    };

    var options = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + settings.apiKey
        },
        payload: JSON.stringify(bodyObj),
        method: 'POST',
        muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
        var body = JSON.parse(response.getContentText());
        var choices = body.choices;
        var contents = choices.map(function (choice) {
            return choice.message.content;
        });
        return contents[0];
    } else {
        Logger.log('Error: ' + response.getContentText());
        return null;
    }
}


