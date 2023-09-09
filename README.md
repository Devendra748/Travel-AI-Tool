# Travel-AI-Tool

This Google Apps Script automates the analysis of emails and extracts specific information from them. It utilizes the OpenAI API to assist in processing and extracting data from emails. The extracted data is then organized and stored in Google Sheets.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Functions](#functions)
- [License](#license)

## Prerequisites

Before using this script, make sure you have the following:

1. A Google account.
2. A Google Sheets document where you want to run the script.
3. Access to the Google Apps Script Editor.

## Installation

1. Open your Google Sheets document.
2. Click on `Extensions` -> `Apps Script` to open the Google Apps Script Editor.
3. Copy and paste the code from this repository into the script editor.
4. Save the script.

## Usage

1. Make sure you have set up the necessary configuration details in your Google Sheets document. You should have values for the API key, Azure OpenAPI, and other required settings.
2. Create a new sheet in your Google Sheets document where you want to run the analysis.
3. In the Google Apps Script Editor, run the `iterateArray()` function to start the email analysis process. This function iterates through the email list and applies the analysis to each email.
4. Once the analysis is complete, the results will be stored in your Google Sheets document.

## Functions

Here are some key functions in the script:

- `gettingEmailAnalysisData(email)`: This function analyzes the provided email and extracts specific data.
- `applyFormulas()`: This function applies formulas to calculate and compare various data points in the generated sheets.
- Other utility functions that retrieve configuration details and format data as needed.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
