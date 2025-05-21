const OPENAI_API_KEY = "your-openai-api-key-here"

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('AI Calculator')
      .addItem('Open Sidebar', 'showSidebar')
      .addToUi();
  } catch (e) {
    console.warn("UI functions can only be run in certain contexts:", e);
  }
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AI-Powered Calculator');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Custom function for use in spreadsheet cells
 * @param {string} text The text to summarize or process
 * @param {string} format The format instruction for the output
 * @param {number} temperature The temperature setting (0-1)
 * @param {string} model The OpenAI model to use
 * @return {string} The AI-generated result
 * @customfunction
 */
function GPT_SUMMARIZE(text, format = "Summarize this text", temperature = 0.7, model) {
  if (!text) return "No input text provided";
  
  try {
    const prompt = `${format}:\n\n${text}`;
    const result = callOpenAI(prompt, model, temperature);
    return result;
  } catch (error) {
    return "Error: " + error.toString();
  }
}

/**
 * Process data using the sidebar settings
 */
function processDataWithAI(config) {
  const { headerRow, prompt, model, resultColumn, startMode, startRange, numberOfRows, allRows } = config;
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    // Extract headers (assuming first row contains headers)
    const headers = data[headerRow - 1]; // Convert to 0-based index
    const resultColIndex = headers.indexOf(resultColumn);
    
    if (resultColIndex === -1) {
      return { success: false, message: `Column "${resultColumn}" not found in headers` };
    }
    
    // Determine rows to process
    let startRowIndex, endRowIndex;
    if (startMode === "auto") {
      startRowIndex = headerRow; // Start from the row after headers
      endRowIndex = allRows ? data.length : startRowIndex + Number(numberOfRows);
    } else { // Fixed mode
      startRowIndex = Number(startRange.split("to")[0].trim()) - 1; // Convert to 0-based
      endRowIndex = Number(startRange.split("to")[1].trim()); // Already converted to 0-based (end is exclusive)
    }
    
    // Make sure we don't exceed data bounds
    endRowIndex = Math.min(endRowIndex, data.length);
    
    // Process each row
    let processedCount = 0;
    for (let i = startRowIndex; i < endRowIndex; i++) {
      const row = data[i];
      
      // Replace column placeholders in the prompt with actual cell values
      let rowPrompt = prompt;
      headers.forEach((header, index) => {
        const placeholder = `{{${header}}}`;
        if (rowPrompt.includes(placeholder)) {
          rowPrompt = rowPrompt.replace(new RegExp(placeholder, 'g'), row[index]);
        }
        console.log(rowPrompt);
      });
      
      // Call OpenAI with the constructed prompt
      const result = callOpenAI(rowPrompt, model);
      
      // Write result back to the sheet
      sheet.getRange(i + 1, resultColIndex + 1).setValue(result); // Convert to 1-based for Range
      processedCount++;
    }
    
    return { success: true, message: `Successfully processed ${processedCount} rows` };
  } catch (error) {
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Helper function to call OpenAI API
 */
function callOpenAI(prompt, model, temperature = 0.7) {
  const url = "https://api.openai.com/v1/chat/completions";
  
  const payload = {
    model: model,
    messages: [
      { role: "system", content: "You are a helpful assistant that provides accurate and concise responses." },
      { role: "user", content: prompt }
    ],
    temperature: temperature
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + OPENAI_API_KEY
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseData = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() !== 200) {
    throw new Error(responseData.error.message || "API Error");
  }
  
  return responseData.choices[0].message.content.trim();
}

/**
 * Get all headers from the active sheet
 */
function getSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.filter(header => header !== '');
}