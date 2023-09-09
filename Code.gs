
function gettingEmailAnalysisData(email) {
  let newSheetName = newCreatedSheetName() + " Report";
  newReportSheet(newSheetName)
  let startTime = Date.now();
  let openApiAzure = azureOpenApi();
  let apiUrl = "https://" + openApiAzure + ".openai.azure.com/openai/deployments/GPT4-8k/chat/completions?api-version=2023-05-15";
  let apiKey = chatGpt4apiKey();
  let todayDate = new Date();
  let prompt = `You are an experienced Ticket agent. Your task is to extract answers from questions:
                Question: Provide the answer for the following `+ email + ` as shown in the examples.
                Example:
                - Start Address: HOME.
                - Meeting Address: St. Louis.
                - Trip: Extracted from the email (e.g., ONE_WAY_OUTGOING, TWO_WAY).
                - Start Day: If a date is present in the email, use that date in YYYY-MM-DD format. If the email mentions a week days, predict the date according to ${todayDate.toDateString()} and format it as YYYY-MM-DD.
                - Start Time 1: Determine based on conditions or extract from the email if specified. If the email mentions "early morning," provide time 1 as 4:00. If it mentions "morning," provide time 1 as 4:00. If it mentions "afternoon," provide time 1 as 11:00. If it mentions "evening," provide time 1 as 18:00. If it mentions "night," provide time 1 as 23:00 (like 2 pm = 14:00).
                - Start Time 2: Determine based on conditions or extract from the email if specified. If the email mentions "early morning," provide time 2 as 8:00. If it mentions "morning," provide time 2 as 11:00 (like 2 pm = 14:00).
                - End Day: If a date is present in the email, use that date in YYYY-MM-DD format. If the email mentions a week days, predict the date according to ${todayDate.toDateString()} and format it as YYYY-MM-DD.
                - End Time 1: Determine based on conditions or extract from the email if specified. If the email mentions "early morning," provide time 1 as 4:00. If it mentions "morning," provide time 1 as 4:00. If it mentions "afternoon," provide time 1 as 11:00. If it mentions "evening," provide time 1 as 18:00. If it mentions "night," provide time 1 as 23:00 (like 2 pm = 14:00).
                - End Time 2: Determine based on conditions or extract from the email if specified. If the email mentions "early morning," provide time 2 as 8:00. If it mentions "morning," provide time 2 as 11:00 (like 2 pm = 14:00).`

  let outputString = prompt.replace(/\n/g, ' ');
  let payload = {
    "model": chatGptModelName(),
    "messages": [
      { "role": "user", "content": outputString }
    ],
    "temperature": 0.0
  };


  let headers = {
    "Content-Type": "application/json",
    "api-key": apiKey,
    "Helicone-Auth": heliconeAuth(),
    "Helicone-User-Id": heliconeId()

  };

  let options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(payload)
  };

  let response = UrlFetchApp.fetch(apiUrl, options);
  let responseData = response.getContentText();
  let parsedResponse = JSON.parse(responseData);
  let assistantResponse = parsedResponse.choices[0].message.content;
  let bulletPointList = assistantResponse.split('\n');
  bulletPointList = bulletPointList.map(point => point.replace(/^- /, ''));
  let completion_tokens = 0 + +parsedResponse['usage']['completion_tokens']
  let prompt_tokens = 0 + +parsedResponse['usage']['prompt_tokens']
  let total_tokens = 0 + +parsedResponse['usage']['total_tokens']
  let endTime = Date.now();
  let runtime = (endTime - startTime) / 1000;
  let cost = (completion_tokens / 1000) * 0.06 + (prompt_tokens / 1000) * 0.03
  let tokenLatencyArray = [completion_tokens, prompt_tokens, total_tokens, runtime, cost]
  console.log(tokenLatencyArray)
  createSubsheetWithColumnsAndValues(newSheetName, tokenLatencyArray)
  console.log("Email:", bulletPointList)
  parseAndInsertData(getHeadings(), bulletPointList, email)
  return bulletPointList
}
function chatGptModelName() {
  let sheetName = "SettingTab";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let cellValue = sheet.getRange("B4").getValue();
  return cellValue;

}
function heliconeId() {
  let sheetName = "SettingTab";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let cellValue = sheet.getRange("B6").getValue();
  return cellValue;

}
function heliconeAuth() {
  let sheetName = "SettingTab";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let cellValue = sheet.getRange("B5").getValue();
  return cellValue;

}


function newCreatedSheetName() {
  let sheetName = "SettingTab";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let cellValue = sheet.getRange("B3").getValue();
  return cellValue;

}
function newReportSheet(newSheetName) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(newSheetName);

  if (!sheet) {
    let newSheet = spreadsheet.insertSheet(newSheetName);
    let headers = ["Completion Tokens", "Prompt Tokens", "Total Tokens", "Runtime For Each Resume", "Cost in Dollar"];
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    newSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    newSheet.getRange(1, 1, 1, headers.length).setWrap(true).setBorder(true, true, true, true, true, true)
  }
}
function createSubsheetWithColumnsAndValues(newSheetName, values) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(newSheetName);

  if (!sheet) {
    let newSheet = spreadsheet.insertSheet(newSheetName);
    let headers = ["Completion Tokens", "Prompt Tokens", "Total Tokens", "Runtime For Each Resume", "Cost in Dollar"];
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    newSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    newSheet.getRange(1, 1, 1, headers.length).setWrap(true).setBorder(true, true, true, true, true, true)
  } else {
    let lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, values.length).setWrap(true).setBorder(true, true, true, true, true, true)
    sheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]);
  }
}
function chatGpt4apiKey() {
  let sheetName = "SettingTab";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let cellValue = sheet.getRange("B1").getValue();
  return cellValue;

}
function azureOpenApi() {
  let sheetName = "SettingTab";
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let cellValue = sheet.getRange("B2").getValue();
  return cellValue;

}

function createSubSheetWithColumns() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let subsheetName = newCreatedSheetName();
  if (spreadsheet.getSheetByName(subsheetName)) {
    Logger.log("Subsheet already exists.");
    return;
  }
  let newSheet = spreadsheet.insertSheet(subsheetName);
  let columns = ['Email', 'Start Address', 'Meeting Address', 'Trip', 'Start Day', 'Start Time 1', 'Start Time 2', 'End Day', 'End Time 1', 'End Time 2', 'Manual?', 'END'];
  let headerRow = newSheet.getRange(1, 1, 1, columns.length);
  headerRow.setFontWeight("bold");
  headerRow.setWrap(true);
  headerRow.setBorder(true, true, true, true, true, true);
  headerRow.setValues([columns]);
}



function getHeadingsInSubSheetByName() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(newCreatedSheetName());
  if (!sheet) {
    Logger.log("Subsheet not found.");
    return [];
  }
  let headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log(headings)
  return headings;
}

function getColumnValues() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newCreatedSheetName());
  let column = sheet.getRange('A2:A');
  let values = column.getValues();
  let flattenedValues = [];
  for (let i = 0; i < values.length; i++) {
    let value = values[i][0];
    if (value === "" || value === null) {
      break;
    }
    flattenedValues.push(value);
  }
  console.log('flattenedValues', flattenedValues)

  return flattenedValues
}

function iterateArray() {
  for (let element of getColumnValues()) {
    console.log(element)
    gettingEmailAnalysisData(element)
  }
  createThirdSubsheet()
  applyFormulas()
}
function parseAndInsertData(columnHeaders, valuesList, email) {
  const sheetName = newCreatedSheetName();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log("Sheet not found");
    return;
  }

  // const lastRow = 1 + 1; 
  const row = [];
  columnHeaders.forEach(header => {
    const dataValue = findValueForHeader(header, valuesList);
    row.push(dataValue);
  });
  let columnBData = sheet.getRange("B:B").getValues().flat();
  let lastRow = columnBData.filter(String).length;
  row.shift()
  row.unshift(email)
  let dataArray = row;
  sheet.getRange(lastRow + 1, 1, 1, dataArray.length).setWrap(true).setBorder(true, true, true, true, true, true)
  sheet.getRange(lastRow + 1, 1, 1, dataArray.length).setValues([dataArray]);
  // sheet.getRange(lastRow, 1, 1, row.length).setValues([row]); // Insert data in the next available row
  // console.log("Row", row);
}

function findValueForHeader(header, valuesList) {
  for (const value of valuesList) {
    const valuePairs = value.split(':');
    const valueHeader = valuePairs[0].trim();
    const dataValue = extractFirstHalf(valuePairs.slice(1).join(':').trim()); // Join the remaining parts after the colon
    if (valueHeader === header) {
      return dataValue;
    }
  }
  return '';
}

function getColumnHeadings() {
  let sheetName = newCreatedSheetName();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    let headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    return headings;
  } else {
    console.log("Sheet with name '" + sheetName + "' not found.");
    return [];
  }
}
function getHeadings() {
  let headings = getColumnHeadings();

  if (headings.length > 0) {
    console.log("Column headings: ", headings);
  }
  console.log("Column headings: ", headings)
  return headings;
}
function extractFirstHalf(inputString) {
  // Split the input string by the opening parenthesis "("
  let parts = inputString.split('(');

  // If there's at least one part after splitting, return the first part (before "(")
  if (parts.length >= 1) {
    return parts[0].trim(); // Trim any leading or trailing spaces
  } else {
    return inputString; // Return the original string if there's no "("
  }
}

function createThirdSubsheet() {
  let comparisonSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(newCreatedSheetName() + " Comparison Sheet");
  let headers = [
    "Email",
    "Start Address (Original)",
    "Start Address (OpenAI)",
    "Comparison of Start Address",
    "Meeting Address (Original)",
    "Meeting Address (OpenAI)",
    "Comparison of Meeting Address",
    "Trip (Original)",
    "Trip (OpenAI)",
    "Comparison of Trip",
    "Start Day (Original)",
    "Start Day (OpenAI)",
    "Comparison of Start Day",
    "Start Time 1 (Original)",
    "Start Time 1 (OpenAI)",
    "Comparison of Start Time 1",
    "Start Time 2 (Original)",
    "Start Time 2 (OpenAI)",
    "Comparison of Start Time 2",
    "End Day (Original)",
    "End Day (OpenAI)",
    "Comparison of End Day",
    "End Time 1 (Original)",
    "End Time 1 (OpenAI)",
    "Comparison of End Time 1",
    "End Time 2 (Original)",
    "End Time 2 (OpenAI)",
    "Comparison of End Time 2",
  ];

  // Append headers to the comparison sheet
  comparisonSheet.appendRow(headers);

  // Get data from two source sheets
  let originalData = getSheetValuesByName("Email Previously provided output");
  let openAiData = getSheetValuesByName(newCreatedSheetName());

  // Merge data from both sheets into the comparison sheet
  for (let i = 0; i < originalData.length+1; i++) {
    let rowDataOriginal = originalData[i];
    let rowDataOpenAi = openAiData[i];

    let comparisonData = [
      rowDataOriginal[0], // Email
      rowDataOriginal[1], // Start Address (Original)
      rowDataOpenAi[1],   // Start Address (OpenAI)
      "=IF(C" + (i + 2) + "=B" + (i + 2) + ", 1, 0)", // Comparison of Start Address
      rowDataOriginal[2], // Meeting Address (Original)
      rowDataOpenAi[2],   // Meeting Address (OpenAI)
      "=IF(F" + (i + 2) + "=E" + (i + 2) + ", 1, 0)", // Comparison of Meeting Address
      rowDataOriginal[3], // Trip (Original)
      rowDataOpenAi[3],   // Trip (OpenAI)
      "=IF(I" + (i + 2) + "=H" + (i + 2) + ", 1, 0)", // Comparison of Trip
      rowDataOriginal[4], // Start Day (Original)
      rowDataOpenAi[4],   // Start Day (OpenAI)
      dayOfWeekToDate(rowDataOpenAi[4], rowDataOriginal[4]), // Comparison of Start Day
      rowDataOriginal[5], // Start Time 1 (Original)
      rowDataOpenAi[5],   // Start Time 1 (OpenAI)
      "=IF(O" + (i + 2) + "=N" + (i + 2) + ", 1, 0)", // Comparison of Start Time 1
      rowDataOriginal[6], // Start Time 2 (Original)
      rowDataOpenAi[6],   // Start Time 2 (OpenAI)
      "=IF(R" + (i + 2) + "=Q" + (i + 2) + ", 1, 0)", // Comparison of Start Time 2
      rowDataOriginal[7], // End Day (Original)
      rowDataOpenAi[7],   // End Day (OpenAI)
      dayOfWeekToDate(rowDataOpenAi[7], rowDataOriginal[7]), // Comparison of End Day
      rowDataOriginal[8], // End Time 1 (Original)
      rowDataOpenAi[8],   // End Time 1 (OpenAI)
      "=IF(X" + (i + 2) + "=W" + (i + 2) + ", 1, 0)", // Comparison of End Time 1
      rowDataOriginal[9], // End Time 2 (Original)
      rowDataOpenAi[9],   // End Time 2 (OpenAI)
      "=IF(AA" + (i + 2) + "=Z" + (i + 2) + ", 1, 0)"  // Comparison of End Time 2
    ];
    console.log(rowDataOriginal[7])
    console.log(rowDataOpenAi[7])
    // Append the comparison data to the comparison sheet
    comparisonSheet.appendRow(comparisonData);
  }
}

function getSheetValuesByName(sheetName) {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  data.splice(0, 1); // Remove headers
  return data;
}

function dayOfWeekToDate(dayOfWeek, input) {
  if (input.length > 20) {
    input = formatDateNew(input)
  }
  dayOfWeek = formatDateNew(dayOfWeek)
  let result = 0;
  let currentDate = new Date();
  if (dayOfWeek !== "" && input !== "" && !isDateFormatValid(input) && dayOfWeek.length < 20 && input.length < 20) {
    let daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

    let daysToAdd = 0;

    if (input.includes("Today")) {
      daysToAdd = 0;
    } else if (input.includes("TOMORROW")) {
      daysToAdd = 1;
    } else if (input.includes("Day after tomorrow")) {
      daysToAdd = 2;
    } else {
      // If none of the keywords are found, assume it's based on the day of the week
      let specifiedDayOfWeek = input;
      let dayIndex = daysOfWeek.indexOf(specifiedDayOfWeek);
      if (dayIndex !== -1) {
        daysToAdd = dayIndex - currentDate.getDay();
        if (daysToAdd <= 0) {
          daysToAdd += 7;
        }
      } else {
        return "Invalid day of the week";
      }
    }

    currentDate.setDate(currentDate.getDate() + daysToAdd);
    let formattedDate = Utilities.formatDate(currentDate, "GMT", "dd/MM/yyyy");

    let dateArray = [formattedDate, formatDateNewStyle(formattedDate)];

    result = checkStrings(dateArray, dayOfWeek);
  }
  else if (isDateFormatValid(input) && dayOfWeek !== "") {
    let formattedDate = input;
    let dateArray = [formattedDate, formatDateNewStyle(formattedDate)];
    result = checkStrings(dateArray, dayOfWeek);
  }
  else {
    result = 0;
  }
  console.log(result)
  return result;
}

function formatDateNewStyle(inputDate) {
  let parts = inputDate.split('/');
  let formattedDate = parts[2] + '-' + parts[1].padStart(2, '0') + '-' + parts[0].padStart(2, '0');
  return formattedDate;
}

function checkStrings(dateArray, dayOfWeek) {
  let results = 0;
  for (let i = 0; i < dateArray.length; i++) {
    if (dayOfWeek.indexOf(dateArray[i]) !== -1) {
      results = 1;
    }
  }
  return results;
}

function isDateFormatValid(dateString) {
  const datePattern = /^(0[1-9]|[12][0-9]|3[01])\/(0[1-9]|1[0-2])\/\d{4}$/;
  return datePattern.test(dateString);
}

function formatDateNew(inputDate) {
  // Create a JavaScript Date object from the input string
  // let inputDate="Sat Oct 21 2023 03:00:00 GMT-0400 (Eastern Daylight Time)"
  let dateObj = new Date(inputDate);

  // Extract day, month, and year components from the date object
  let day = dateObj.getDate();
  let month = dateObj.getMonth() + 1; // Months are 0-based, so add 1
  let year = dateObj.getFullYear();

  // Format day, month, and year as dd/mm/yyyy
  let formattedDate = padZero(day) + "/" + padZero(month) + "/" + year;
  console.log('formattedDate', formattedDate)
  return formattedDate;
}

function padZero(number) {
  return number < 10 ? "0" + number : number.toString();
}
function applyFormulaComparisonofMeetingAddress() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(G2:G${lastRow-1})/COUNT(G2:G${lastRow-1})`;
  const cell = sheet.getRange(`G${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofStartAddress() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(D2:D${lastRow})/COUNT(D2:D${lastRow})`;
  const cell = sheet.getRange(`D${lastRow+1}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofTrip() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(J2:J${lastRow-1})/COUNT(J2:J${lastRow-1})`;
  const cell = sheet.getRange(`J${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofStartDay() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(M2:M${lastRow-1})/COUNT(M2:M${lastRow-1})`;
  const cell = sheet.getRange(`M${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofStartTime1() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(P2:P${lastRow-1})/COUNT(P2:P${lastRow-1})`;
  const cell = sheet.getRange(`P${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofStartTime2() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(S2:S${lastRow-1})/COUNT(S2:S${lastRow-1})`;
  const cell = sheet.getRange(`S${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofEndDay() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(V2:V${lastRow-1})/COUNT(V2:V${lastRow-1})`;
  const cell = sheet.getRange(`V${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofMeetingAddressEndTime2() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(Y2:Y${lastRow-1})/COUNT(Y2:Y${lastRow-1})`;
  const cell = sheet.getRange(`Y${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}
function applyFormulaComparisonofEndTime2() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(newCreatedSheetName() + " Comparison Sheet"); // Replace with the actual sheet name
  const lastRow = sheet.getLastRow();
  const formula = `=SUM(AB2:AB${lastRow-1})/COUNT(AB2:AB${lastRow-1})`;
  const cell = sheet.getRange(`AB${lastRow}`); // Replace with the actual cell reference
  cell.setFormula(formula+"%");
}

function applyFormulas() {
  
  applyFormulaComparisonofStartAddress()
  applyFormulaComparisonofMeetingAddress()
  applyFormulaComparisonofTrip()
  applyFormulaComparisonofStartDay()
  applyFormulaComparisonofStartTime1()
  applyFormulaComparisonofStartTime2()
  applyFormulaComparisonofEndDay()
  applyFormulaComparisonofMeetingAddressEndTime2()
  applyFormulaComparisonofEndTime2()
}
