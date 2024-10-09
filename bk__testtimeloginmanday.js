function readSheet5DataAndWriteToSheet6() {
  // Open the active spreadsheet
  var sheet5 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet5");
  var sheet6 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet6");
  
  if (!sheet5) {
    Logger.log("Sheet5 not found!");
    return;
  }
  
  if (!sheet6) {
    // Create "Sheet6" if it does not exist
    sheet6 = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Sheet6");
  } else {
    // Clear the existing data in "TestATOM" before adding new data
    sheet6.clear();
  }

  // Get all data from Sheet5
  var data = sheet5.getDataRange().getValues();
  
  // Assuming first row contains headers, skip it
  var headers = data[0];
  Logger.log("Headers: " + headers);

  var userData = {}; // Object to store user-wise data

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var key = row[0]; // 'Key' column
    var issueType = row[1]; // 'Issue Type' column
    var labels = row[2].toLowerCase(); // 'Labels' column (convert to lower case for easier matching)
    var author = row[3]; // 'Author Display Name' column
    var timeSpentHours = row[6]; // 'Time Spent (Seconds)' column
    
    // var timeSpentHours = timeSpentSeconds / 3600; // Convert time spent to hours

    // Initialize user data if not existing
    if (!userData[author]) {
      userData[author] = {
        bugLogHours: 0,
        nonDevLogHours: 0,
        devLogHours: 0,
        bugCount: 0,
        totalLogHours: 0
      };
    }

    // Process time spent based on issue type and labels
    if (issueType === "Bug" || labels.includes("bug")) {
      // Bug-related log hours
      userData[author].bugLogHours += timeSpentHours;
    } else if (labels.includes("non-functional")) {
      // Non-dev related log hours
      userData[author].nonDevLogHours += timeSpentHours;
    }

    // Calculate dev log hours as total time minus bug and non-dev hours
    var otherHours = timeSpentHours - (userData[author].bugLogHours + userData[author].nonDevLogHours);

    // Increment bug count if the issue type is "Bug"
    if (issueType === "Bug") {
      userData[author].bugCount++;
    }

    // Add time spent to total log hours
    userData[author].totalLogHours += timeSpentHours;
    
    // Calculate dev log hours, ensuring non-negative values
    var devHour = userData[author].totalLogHours - (userData[author].bugLogHours + userData[author].nonDevLogHours);
    userData[author].devLogHours = devHour > 0 ? devHour : 0; 
  }

  // Prepare data for writing to Sheet6
  var output = [];
  output.push(["Author", "Bug Log Hours (Man Days)", "Non-Dev Log Hours (Man Days)", "Dev Log Hours (Man Days)", "Bug Count", "Total Log Hours (Man Days)"]); // Headers

  for (var user in userData) {
    // Convert each hour metric to man days (hours/6.5)
    var bugLogManDays = (userData[user].bugLogHours / 6.5).toFixed(2);
    var nonDevLogManDays = (userData[user].nonDevLogHours / 6.5).toFixed(2);
    var devLogManDays = (userData[user].devLogHours / 6.5).toFixed(2);
    var totalLogManDays = (userData[user].totalLogHours / 6.5).toFixed(2);

    output.push([
      user,
      bugLogManDays,
      nonDevLogManDays,
      devLogManDays,
      userData[user].bugCount,
      totalLogManDays
    ]);
  }

  // Write headers to the first row and data from the second row in Sheet6
  sheet6.getRange(1, 1, 1, output[0].length).setValues([output[0]]); // Write headers to row 1
  sheet6.getRange(2, 1, output.length - 1, output[0].length).setValues(output.slice(1)); // Write data from row 2

  Logger.log("Data written to Sheet6 starting from row 2");
}
