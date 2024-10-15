function fetchAndInsertData() {

  const startDate = '2024-09-21';
  const endDate = '2024-09-30';
  const url = `url/api/1/exportData.csv?_allProjects=true&allUsers=true&startDate=${startDate}&endDate=${endDate}&moreFields=labels&moreFields=timeoriginalestimate&Apikey=token`;

  // Fetch the CSV data from the API
  const response = UrlFetchApp.fetch(url);
  const csvData = response.getContentText();

  // Parse the CSV data
  const rows = Utilities.parseCsv(csvData);

  // Extract the header and desired columns
  const desiredColumns = ["Project", "Issue Type", "Key", "Summary", "Priority", "Labels", "Timeoriginalestimate", "Date Started", "Display Name", "Time Spent (h)", "Work Description"];
  const header = rows[0];

  // Get indexes of desired columns
  const columnIndexes = desiredColumns.map(column => header.indexOf(column));

  // Prepare data to insert and filter out the header
  const dataToInsert = rows.slice(1).map(row => {
    return columnIndexes.map(index => row[index]);
  }).filter(row => row.length > 0);

  // Access the spreadsheet and the desired sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let newSheet = spreadsheet.getSheetByName('Timesheet Data');

  if (!newSheet) {
    // If the sheet doesn't exist, create it and set the header row
    newSheet = spreadsheet.insertSheet('Timesheet Data');
    newSheet.appendRow(desiredColumns);
    // Insert all new data for the first time
    newSheet.getRange(2, 1, dataToInsert.length, desiredColumns.length).setValues(dataToInsert);
    return; // Exit function after first time insert
  }

  // Find the last row with "Total" in the sheet
  const dataRange = newSheet.getDataRange();
  const dataValues = dataRange.getValues();
  let totalRow = -1; // Initialize totalRow to -1 (not found)

  // Look for the row with "Total"
  for (let i = 0; i < dataValues.length; i++) {
    if (dataValues[i].some(cell => typeof cell === 'string' && cell.toLowerCase().includes('total'))) {
      totalRow = i + 1; // 1-based index
      break;
    }
  }

  // If a Total row is found, replace it with new data
  if (totalRow > 0) {
    const numRowsToReplace = Math.min(dataToInsert.length, dataValues.length - totalRow);
    // Check if there are rows to replace
    if (numRowsToReplace > 0) {
      // Replace the "Total" row and any subsequent rows
      newSheet.getRange(totalRow, 1, numRowsToReplace, desiredColumns.length).setValues(dataToInsert.slice(0, numRowsToReplace));
    }

    // If new data has more rows than existing data, add remaining data to new rows
    if (dataToInsert.length > numRowsToReplace) {
      newSheet.getRange(totalRow + numRowsToReplace, 1, dataToInsert.length - numRowsToReplace, desiredColumns.length).setValues(dataToInsert.slice(numRowsToReplace));
    }
  } else {
    // If no "Total" row is found, append new data at the end
    newSheet.getRange(dataValues.length + 1, 1, dataToInsert.length, desiredColumns.length).setValues(dataToInsert);
  }
}
