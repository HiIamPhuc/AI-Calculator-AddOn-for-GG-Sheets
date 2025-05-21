# AI-Calculator-AddOn-for-GG-Sheets

This add-on integrates AI capabilities directly into Google Sheets, allowing users to run GPT-powered calculations and custom prompts on their spreadsheet data without leaving the sheet.

## Features

- **Seamless Google Sheets Integration**: Adds a custom menu item and sidebar for easy access
- **Batch Processing Support**: Generate outputs for multiple rows based on column values and a shared prompt
- **AI-Powered**: Uses OpenAI's Chat API (`gpt-4o`, `gpt-4o-mini`) for intelligent, context-aware responses
- **Customizable Options**:
    - Adjustable temperature setting for output creativity
    - Prompt templating using `{{columnName}}` placeholders
    - Flexible model selection and control over result placement

## Installation Instructions

Follow these steps to test the add-on in your Google Sheets environment:

1. **Create a new Google Apps Script project**:
    - Open any Google Sheet and go to **Extensions > Apps Script**
    - This will open the Google Apps Script editor in a new tab

2. **Set up the project files**:
    - Replace the default content in the editor with the code from `Code.gs`
    - Click the **+** button next to "Files" and select **HTML**
    - Name the new file `Sidebar.html`
    - Paste the HTML content from `Sidebar.html` into this file

3. **Save the project**:
    - Click on the unnamed project title and rename it to "AI-Powered Calculator"
    - Press **Save** (Ctrl+S or ⌘+S)

4. **Configure the API key**:
    - Edit this line in `Code.gs`:
    
        ```js
        const OPENAI_API_KEY = "your-openai-api-key-here";
        ```

5. **Run the project directly**:
    - Select the `onOpen` function from the dropdown next to the "Debug" button
    - Click the "Run" button (►)
    - Authorize the script when prompted

6. **Access the add-on in Google Sheets**:
    - Refresh your Google Sheet
    - You should now see a new menu item called "AI Calculator"
    - Click **AI Calculator > Open Sidebar** to launch the interface

## Using the Add-on

1. The sidebar will appear on the right side of your spreadsheet with a simple prompt interface.

2. **Running a Batch Calculation**:
    - **Enter Prompt**: Write a prompt using column headers as variables (e.g., `Summarize {{Review}}`)
    - **Select Settings**:
        - Starting row
        - Number of rows (or process entire sheet)
        - Column name to store results (will auto-create if needed)
        - AI model and temperature
    - **Click "Process Data"** to let the AI process each row

3. **Prompt Templating**:
   - Use `{{columnName}}` syntax to insert cell values from a given row
   - Example: `What are the pros and cons of {{Product Description}}?`

4. **Tips for Best Results**:
   - Double-check your column names (case-sensitive) and that your prompt uses them correctly
   - Use lower temperatures for factual/consistent answers
   - Use higher temperatures for creative or abstract outputs

## Using the GPT_SUMMARIZE() Function

In addition to the sidebar interface, this add-on provides a custom Google Sheets formula called `GPT_SUMMARIZE()` for quick in-cell AI summarization.

### Syntax

```
=GPT_SUMMARIZE(text, [prompt], [model], [temperature])
```

### Parameters

- `text` *(required)*: The text you want to summarize or process.
- `prompt` *(optional)*: A custom instruction for GPT. Defaults to `"Summarize this text:"`.
- `model` *(optional)*: Specify `"gpt-4o-mini"` or `"gpt-4o"`. Defaults to `"gpt-4o-mini"`.
- `temperature` *(optional)*: A number between 0 and 1 controlling creativity. Defaults to `0.7`.

### Example Usage

```
=GPT_summarize(A2)
=GPT_summarize(A2, "Extract the key issues mentioned in this review")
=GPT_summarize(A2, "Summarize formally", 0.3, "gpt-4o")
```

### Notes

- The function calls the OpenAI API and requires that your script project is configured with your API key.
- Due to Google Sheets quotas and API limits, excessive or frequent use may result in rate limiting.
- Output is returned directly into the cell.

## Important Notes

- You must have an active internet connection and a valid OpenAI API key to use this tool
- The OpenAI API may incur charges based on your usage; monitor your token usage in your OpenAI account

## Troubleshooting

If you encounter any issues:
- Make sure you've authorized the script to run
- Confirm your API key is correctly added in script
- Check your internet connection
- Reload the spreadsheet and try again
