function getJiraData() {
  // Set your Jira base URL and credentials
  const baseUrl = 'url'; // Replace with your Jira base URL
  const email = 'email'; // Replace with your Jira account email
  const apiToken = 'API Token';

Logger.log(Utilities.base64Encode(email + ':' + apiToken))

  const projectName = searchValues('B1');
  const monthMMM = searchValues('D1');
  const valueYYYY = searchValues('F1');

  
  const monthMM = convertMMMtoMM(monthMMM);
  const lastDay = getLastDayOfMonth(monthMM);

  // Set your date range
  const startDate = ''+valueYYYY+'-'+monthMM+'-20'; // Start date in YYYY-MM-DD format
  const endDate = ''+valueYYYY+'-'+monthMM+'-20'; // End date in YYYY-MM-DD format
  // const endDate = ''+valueYYYY+'-'+monthMM+'-'+lastDay+''; // End date in YYYY-MM-DD format

  // The API endpoint to fetch issues
  const endpoint = '/rest/api/3/search';
  const projectCode = searchProjectCode(projectName);

  // Construct the request URL with JQL to filter issues by date range
  const jqlQuery = `project= "${projectCode}" AND worklogDate >= "${startDate}" AND worklogDate <= "${endDate}"`;

  // Set up the headers for authorization
  const headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(email + ':' + apiToken),
    'Accept': 'application/json'
  };

  // Initialize the spreadsheet and sheets
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'Jira Data'; // Raw data sheet
  // const summarySheetName = 'User Worklog Summary'; // Summary sheet
  const summarySheetName = 'worklog'; // Summary sheet
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    // Create the raw data sheet if it doesn't exist
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(['Project Name', 'Summary', 'Jira Ticket', 'Priority', 'Labels', 'Issue Type', 'Assignee', 'Worklogs', 'Total Worklog (Hours)']);

  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
  }

  // Create or clear the summary sheet for storing aggregated worklog data
  let summarySheet = spreadsheet.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet(summarySheetName);
    summarySheet.getRange(2, 1, 1, 6).setValues([['Project Name', 'User', 'Dev Log Hours', 'Bug Worklog (Hours)', 'Non-Dev Worklog (Hours)', 'Bug Count']]);
  } else {
    summarySheet.getRange(3, 1, summarySheet.getMaxRows() - 2, summarySheet.getMaxColumns()).clearContent();
  }

  // Add headers to the raw data sheet


  // Pagination variables
  let startAt = 0;
  const maxResults = 50;
  let total = 1; // Initialize with a value greater than startAt to enter the loop

  // Object to store user worklogs by project
  const worklogSummary = {};

  // Loop to fetch all issues using pagination
  while (startAt < total) {
    // Construct the URL with pagination parameters
    const url = `${baseUrl}${endpoint}?jql=${encodeURIComponent(jqlQuery)}&startAt=${startAt}&maxResults=${maxResults}`;
    Logger.log(url);

    try {
      // Fetch the response
      const response = UrlFetchApp.fetch(url, {
        method: 'GET',
        headers: headers,
        muteHttpExceptions: true
      });

      // Check if the response is successful (HTTP status 200)
      if (response.getResponseCode() !== 200) {
        Logger.log(`Error: Received response code ${response.getResponseCode()}`);
        Logger.log(`Response: ${response.getContentText()}`);
        break; // Exit the loop if the response is not successful
      }

      const data = JSON.parse(response.getContentText());

      // Check if the data object contains the issues array
      if (!data || !data.issues || !Array.isArray(data.issues)) {
        Logger.log('Error: Issues data is missing or not an array.');
        Logger.log(`Received data: ${JSON.stringify(data)}`);
        break; // Exit the loop if data is not as expected
      }

      // Update the total to know when to stop
      total = data.total;

      // Parse and add the relevant fields to the raw data sheet
      data.issues.forEach(issue => {
        const projectName = issue.fields.project ? issue.fields.project.name : 'Unknown Project';
        const summary = issue.fields.summary || 'No Summary';
        const assignee = issue.fields.assignee ? issue.fields.assignee.displayName : 'Unassigned';
        const ticket = issue.key || 'No Ticket Key';
        const priority = issue.fields.priority ? issue.fields.priority.name : 'N/A';
        const labels = issue.fields.labels && issue.fields.labels.length > 0 
          ? issue.fields.labels.join(', ') 
          : 'N/A';
        const issueType = issue.fields.issuetype ? issue.fields.issuetype.name : 'N/A';

        // Determine if the issue is a bug or non-functional
        const isBug = issueType.toLowerCase() === 'bug';
        const isNonFunctional = labels.includes('non-functional');

        // Get worklogs (requires a separate API call for each issue)
        const worklogUrl = `${baseUrl}/rest/api/3/issue/${issue.key}/worklog`;
        const worklogResponse = UrlFetchApp.fetch(worklogUrl, {
          method: 'GET',
          headers: headers,
          muteHttpExceptions: true
        });

        // Check if the worklog response is successful
        if (worklogResponse.getResponseCode() !== 200) {
          Logger.log(`Error fetching worklogs for issue ${ticket}: ${worklogResponse.getContentText()}`);
          return; // Skip this issue if the worklog fetch fails
        }

        const worklogData = JSON.parse(worklogResponse.getContentText());

        // Retrieve and convert worklogs into hours
        let issueTotalHours = 0;
        let issueBugHours = 0;
        let issueNonDevHours = 0;

        const worklogs = (worklogData.worklogs || []).map(wl => {
          const timeSpent = wl.timeSpent || '0m';
          const hours = convertTimeSpentToHours(timeSpent);
          issueTotalHours += hours;

          if (isBug) {
            issueBugHours += hours;
          }

          if (isNonFunctional) {
            issueNonDevHours += hours;
          }

          return timeSpent;
        }).join(', ');

        // Append the data as a new row in the raw data sheet
        sheet.appendRow([projectName, summary, ticket, priority, labels, issueType, assignee, worklogs, issueTotalHours]);

        // Aggregate worklog data for summary sheet
        if (!worklogSummary[projectName]) {
          worklogSummary[projectName] = {};
        }

        if (!worklogSummary[projectName][assignee]) {
          worklogSummary[projectName][assignee] = { 
            totalHours: 0, 
            bugHours: 0, 
            bugCount: 0, 
            nonDevHours: 0 
          };
        }

        worklogSummary[projectName][assignee].totalHours += issueTotalHours;
        worklogSummary[projectName][assignee].bugHours += issueBugHours;
        if (isBug) {
          worklogSummary[projectName][assignee].bugCount += 1;
        }
        worklogSummary[projectName][assignee].nonDevHours += issueNonDevHours;
      });

      // Update startAt to fetch the next set of results
      startAt += maxResults;

    } catch (error) {
      Logger.log(`Error: ${error.message}`);
      break; // Exit the loop if an error occurs
    }
  }

  // Add the summarized worklog data to the summary sheet
  for (const project in worklogSummary) {
    for (const user in worklogSummary[project]) {
      const { totalHours, bugHours, bugCount, nonDevHours } = worklogSummary[project][user];
      const devHours = totalHours - (bugHours + nonDevHours); // Calculate Dev Log Hours
      summarySheet.appendRow([project, user, devHours, bugHours, nonDevHours, bugCount]);
    }
  }

  Logger.log('All data has been successfully added to the Google Sheet.');
}

/**
 * Convert Jira worklog time spent format (e.g., "2d 5h 30m") to total hours.
 * @param {string} timeSpent - The time spent string from Jira worklog.
 * @return {number} Total hours.
 */
function convertTimeSpentToHours(timeSpent) {
  let totalHours = 0;

  // Match days, hours, and minutes from the timeSpent string
  const daysMatch = timeSpent.match(/(\d+)d/);
  const hoursMatch = timeSpent.match(/(\d+)h/);
  const minutesMatch = timeSpent.match(/(\d+)m/);

  if (daysMatch) {
    totalHours += parseInt(daysMatch[1], 10) * 24; // Convert days to hours
  }

  if (hoursMatch) {
    totalHours += parseInt(hoursMatch[1], 10); // Add hours
  }

  if (minutesMatch) {
    totalHours += parseInt(minutesMatch[1], 10) / 60; // Convert minutes to hours
  }

  return totalHours;
}

function searchValues(value) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
  var range = sheet.getRange(value);
  
  // Get the data validation rule for that range
  var dataValidation = range.getDataValidations()[0][0];
  
  // Check if data validation exists and it's a dropdown
  if (dataValidation != null && dataValidation.getCriteriaType() == SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    
    
    // Get the selected value in the cell
    var selectedValue = range.getValue();
    
    // Log the selected value
    Logger.log("Selected value: " + selectedValue);
    
    // Return both the dropdown values and selected value
    return selectedValue
    
  } else {
    Logger.log('No dropdown found in the specified range.');
  }
}

function searchProjectCode(projectNameValue){
  var jiraDomain = 'Jira url';  // Replace with your Jira domain
  var email = 'EMAIL';  // Replace with your Jira email
  var apiToken = 'API token;  // Replace with your Jira API token
  
  // API endpoint to get all projects
  var url = `https://${jiraDomain}/rest/api/3/project`;
  
  // Set up headers for authentication
  var headers = {
    'Authorization': 'Basic ' + Utilities.base64Encode(email + ':' + apiToken),
    'Accept': 'application/json'
  };
  
  // Make the GET request to Jira API to fetch projects
  var options = {
    'method': 'GET',
    'headers': headers,
    'muteHttpExceptions': true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var jsonResponse = JSON.parse(response.getContentText());
  
  // Loop through the projects to find the project with the matching name
  var projectCode = null;
  for (var i = 0; i < jsonResponse.length; i++) {
    if (jsonResponse[i].name == projectNameValue) {
      projectCode = jsonResponse[i].key;  // The project code (key) you are looking for
      Logger.log('Project Code: ' + projectCode);
      break;
    }
  }
  
  // If no project found, log a message
  if (!projectCode) {
    Logger.log('No project found with the name: ' + projectName);
    projectCode = '';
  }
  return projectCode;
}

function convertMMMtoMM(value) {
 // Create a map for the conversion from "MMM" to "mm"
  var monthMap = {
    'Jan': '01',
    'Feb': '02',
    'Mar': '03',
    'Apr': '04',
    'May': '05',
    'Jun': '06',
    'Jul': '07',
    'Aug': '08',
    'Sep': '09',
    'Oct': '10',
    'Nov': '11',
    'Dec': '12'
  };
  
  // Convert the "MMM" to "mm" using the map
  var monthMM = monthMap[value];
  
  // If the month is valid, place the result in cell B1
  if (monthMM) {
    Logger.log("Converted month: " + monthMM);
  } else {
    Logger.log("Invalid month: " + value);
  }
  return monthMM;
}

function getLastDayOfMonth( month) {
  var year = new Date().getFullYear();  // Get the current year
  var monthNum = parseInt(month);       // Convert two-digit month to a number
  
  // Create a new date for the first day of the next month, then go back one day
  var lastDay = new Date(year, monthNum, 0);  // Month is 0-indexed, so we use monthNum + 1

  // Get the last day of the month
  Logger.log("Last day of the month: " + lastDay.getDate());
  
  return lastDay.getDate();  // Return the last day
}
