function generateWeeklyResourceUsageReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resourcesSheet = ss.getSheetByName('Resources');
  const reportSheet = ss.getSheetByName('Weekly Report') || ss.insertSheet('Weekly Report');

  const dataRange = resourcesSheet.getDataRange();
  const data = dataRange.getValues();

  const today = new Date();
  const oneWeekAgo = new Date(today);
  oneWeekAgo.setDate(today.getDate() - 7);

  let report = [];
  report.push(['Resource ID', 'Resource Name', 'Total Usage Hours']);

  for (let i = 1; i < data.length; i++) {
    const usageEndDate = new Date(data[i][13]); 
    if (usageEndDate >= oneWeekAgo && usageEndDate <= today) {
      report.push([data[i][0], data[i][1], data[i][3]]); 
    }
  }

  reportSheet.clear();
  reportSheet.getRange(1, 1, report.length, report[0].length).setValues(report);

  // Update existing chart or create a new one if it doesn't exist
  const charts = reportSheet.getCharts();
  if (charts.length > 0) {
    let chart = charts[0];
    chart = chart.modify()
                 .clearRanges()
                 .addRange(reportSheet.getRange('A2:C' + report.length))
                 .setOption('title', 'Weekly Resource Usage')
                 .build();
    reportSheet.updateChart(chart);
  } else {
    const chart = reportSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(reportSheet.getRange('A2:C' + report.length))
      .setPosition(1, 4, 0, 0)
      .setOption('title', 'Weekly Resource Usage')
      .build();
    reportSheet.insertChart(chart);
  }

  // Create a temporary spreadsheet to include only the relevant sheets
  const tempSpreadsheet = SpreadsheetApp.create('TempSpreadsheet');
  const tempResourcesSheet = resourcesSheet.copyTo(tempSpreadsheet);
  tempResourcesSheet.setName('Resources');
  const tempReportSheet = reportSheet.copyTo(tempSpreadsheet);
  tempReportSheet.setName('Weekly Report');

  // Remove default sheet in temp spreadsheet
  tempSpreadsheet.deleteSheet(tempSpreadsheet.getSheets()[0]);

  // Generate PDF
  const pdfUrl = 'https://docs.google.com/spreadsheets/d/' + tempSpreadsheet.getId() + '/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=true&printtitle=false&pagenumbers=false&gridlines=false&fzr=true';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(pdfUrl, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  const pdf = response.getBlob().setName('Weekly_Resource_Usage_Report.pdf');

  // Send email with PDF attachment
  MailApp.sendEmail({
    to: 'piopreety2807@gmail.com',
    subject: 'Weekly Resource Usage Report',
    body: 'Please find the attached weekly resource usage report.',
    attachments: [pdf]
  });

  // Clean up by deleting the temporary spreadsheet
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
}

function checkResourceUtilization() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Resources');
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    Logger.log('Data: ' + JSON.stringify(data));

    let underUtilizedResources = [];
    let overUtilizedResources = [];

    const underUtilizationThreshold = 10; // Define thresholds as needed
    const overUtilizationThreshold = 100;

    for (let i = 1; i < data.length; i++) {
      const usageHours = data[i][3]; 
      Logger.log('Row ' + i + ': ' + data[i][1] + ', Usage Hours: ' + usageHours);
      if (usageHours < underUtilizationThreshold) {
        underUtilizedResources.push(data[i][1]); 
      } else if (usageHours > overUtilizationThreshold) {
        overUtilizedResources.push(data[i][1]);
      }
    }

    Logger.log('Under-utilized Resources: ' + underUtilizedResources);
    Logger.log('Over-utilized Resources: ' + overUtilizedResources);

    let alertMessage = 'Resource Utilization Alert:\n\n';
    let reportData = [['Resource Name', 'Utilization Status', 'Usage Hours']];
    
    if (underUtilizedResources.length > 0) {
      alertMessage += 'Under-utilized Resources:\n' + underUtilizedResources.join(', ') + '\n\n';
      underUtilizedResources.forEach(resource => {
        const resourceData = data.find(row => row[1] === resource);
        reportData.push([resource, 'Under-utilized', resourceData[3]]);
      });
    }
    
    if (overUtilizedResources.length > 0) {
      alertMessage += 'Over-utilized Resources:\n' + overUtilizedResources.join(', ') + '\n\n';
      overUtilizedResources.forEach(resource => {
        const resourceData = data.find(row => row[1] === resource);
        reportData.push([resource, 'Over-utilized', resourceData[3]]);
      });
    }

    if (underUtilizedResources.length > 0 || overUtilizedResources.length > 0) {
      // Create or get the sheet for the report
      const reportSheet = spreadsheet.getSheetByName('Resource Utilization Report') || 
                          spreadsheet.insertSheet('Resource Utilization Report');
      reportSheet.clear();
      reportSheet.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);
      
      // Update existing chart or create a new one if it doesn't exist
      const charts = reportSheet.getCharts();
      if (charts.length > 0) {
        let chart = charts[0];
        chart = chart.modify()
                     .clearRanges()
                     .addRange(reportSheet.getRange('A2:C' + reportData.length))
                     .setOption('title', 'Resource Utilization Status')
                     .build();
        reportSheet.updateChart(chart);
      } else {
        const chart = reportSheet.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(reportSheet.getRange('A2:C' + reportData.length))
          .setPosition(1, 4, 0, 0)
          .setOption('title', 'Resource Utilization Status')
          .build();
        reportSheet.insertChart(chart);
      }

      // Create a temporary spreadsheet to include only the relevant sheets
      const tempSpreadsheet = SpreadsheetApp.create('TempSpreadsheet');
      const tempResourcesSheet = sheet.copyTo(tempSpreadsheet);
      tempResourcesSheet.setName('Resources');
      const tempReportSheet = reportSheet.copyTo(tempSpreadsheet);
      tempReportSheet.setName('Resource Utilization Report');

      // Remove default sheet in temp spreadsheet
      tempSpreadsheet.deleteSheet(tempSpreadsheet.getSheets()[0]);

      // Generate PDF
      const pdfUrl = 'https://docs.google.com/spreadsheets/d/' + tempSpreadsheet.getId() + '/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=true&printtitle=false&pagenumbers=false&gridlines=false&fzr=true';
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(pdfUrl, {
        headers: {
          'Authorization': 'Bearer ' + token
        }
      });
      const pdf = response.getBlob().setName('Resource_Utilization_Report.pdf');

      // Send email with PDF attachment
      MailApp.sendEmail({
        to: 'piopreety2807@gmail.com', // Replace with recipient email address
        subject: 'Resource Utilization Alert',
        body: alertMessage,
        attachments: [pdf]
      });

      // Clean up by deleting the temporary spreadsheet
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);

      Logger.log('Email sent successfully with PDF attachment.');
    } else {
      Logger.log('No under-utilized or over-utilized resources found.');
    }
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}

function sendTaskReminders() {
  var sheetName = "Tasks"; // Ensure this matches exactly your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Sheet not found: " + sheetName);
    return;
  }

  var dataRange = sheet.getDataRange();
  if (!dataRange) {
    Logger.log("Data range not found");
    return;
  }

  var data = dataRange.getValues();
  var today = new Date();
  var emailSent = 0;

  for (var i = 1; i < data.length; i++) {
    var dueDate = new Date(data[i][6]); 
    var email = data[i][4]; 
    var taskName = data[i][1]; 

    // Calculate the difference in days between the due date and today
    var diffDays = (dueDate - today) / (1000 * 60 * 60 * 24);
    
    if (diffDays <= 1 && data[i][6] != "Completed") {
      var emailBody = `Dear Employee,
      
      This is a reminder that the task '${taskName}' is due tomorrow. Please ensure it is completed on time.

      Best Regards,
      Task Management System`;

      MailApp.sendEmail({
        to: email,
        subject: "Task Reminder: " + taskName,
        body: emailBody
      });

      emailSent++;
    }
  }

  Logger.log("Total emails sent: " + emailSent);
}

function updateWeeklyReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resources');
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  const today = new Date();
  const oneWeekAgo = new Date(today);
  oneWeekAgo.setDate(today.getDate() - 7);

  let report = [];
  report.push(['Resource ID', 'Resource Name', 'Total Usage Hours']);

  for (let i = 1; i < data.length; i++) {
    const usageEndDate = new Date(data[i][12]); 
    if (usageEndDate >= oneWeekAgo && usageEndDate <= today) {
      report.push([data[i][0], data[i][1], data[i][3]]); 
    }
  }

  const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Weekly Report') || 
                      SpreadsheetApp.getActiveSpreadsheet().insertSheet('Weekly Report');
  reportSheet.clear();
  reportSheet.getRange(1, 1, report.length, report[0].length).setValues(report);

  // Update existing chart or create a new one if it doesn't exist
  const charts = reportSheet.getCharts();
  if (charts.length > 0) {
    let chart = charts[0];
    chart = chart.modify()
                 .clearRanges()
                 .addRange(reportSheet.getRange('A2:C' + report.length))
                 .setOption('title', 'Weekly Resource Usage')
                 .build();
    reportSheet.updateChart(chart);
  } else {
    const chart = reportSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(reportSheet.getRange('A2:C' + report.length))
      .setPosition(1, 4, 0, 0)
      .setOption('title', 'Weekly Resource Usage')
      .build();
    reportSheet.insertChart(chart);
  }
}

function updateResourceUtilizationReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resources');
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  let underUtilizedResources = [];
  let overUtilizedResources = [];

  const underUtilizationThreshold = 10;
  const overUtilizationThreshold = 100;

  for (let i = 1; i < data.length; i++) {
    const usageHours = data[i][3]; 
    if (usageHours < underUtilizationThreshold) {
      underUtilizedResources.push(data[i][1]); 
    } else if (usageHours > overUtilizationThreshold) {
      overUtilizedResources.push(data[i][1]);
    }
  }

  let reportData = [['Resource Name', 'Utilization Status', 'Usage Hours']];
  
  underUtilizedResources.forEach(resource => {
    const resourceData = data.find(row => row[1] === resource);
    reportData.push([resource, 'Under-utilized', resourceData[3]]);
  });

  overUtilizedResources.forEach(resource => {
    const resourceData = data.find(row => row[1] === resource);
    reportData.push([resource, 'Over-utilized', resourceData[3]]);
  });

  const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resource Utilization Report') || 
                      SpreadsheetApp.getActiveSpreadsheet().insertSheet('Resource Utilization Report');
  reportSheet.clear();
  reportSheet.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);

  // Update existing chart or create a new one if it doesn't exist
  const charts = reportSheet.getCharts();
  if (charts.length > 0) {
    let chart = charts[0];
    chart = chart.modify()
                 .clearRanges()
                 .addRange(reportSheet.getRange('A2:C' + reportData.length))
                 .setOption('title', 'Resource Utilization Status')
                 .build();
    reportSheet.updateChart(chart);
  } else {
    const chart = reportSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(reportSheet.getRange('A2:C' + reportData.length))
      .setPosition(1, 4, 0, 0)
      .setOption('title', 'Resource Utilization Status')
      .build();
    reportSheet.insertChart(chart);
  }
}

function generateProjectReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tasks');
  const data = sheet.getDataRange().getValues();
  
  let reportData = [['Task ID', 'Task Name', 'Client Name', 'Description', 'Assigned To', 'Start Date', 'Due Date', 'Budget(RM)', 'Priority', 'Status', 'Progress']];
  
  for (let i = 1; i < data.length; i++) {
    reportData.push([data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][5], data[i][6], data[i][7], data[i][8], data[i][9], data[i][10]]);
  }
  
  const reportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Project Report') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Project Report');
  reportSheet.clear();

  // Remove existing charts
  const charts = reportSheet.getCharts();
  charts.forEach(chart => reportSheet.removeChart(chart));

  reportSheet.getRange(1, 1, reportData.length, reportData[0].length).setValues(reportData);
  
  // Create a chart in the 'Project Report' sheet
  const chart = reportSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(reportSheet.getRange('A1:K' + reportData.length))
    .setPosition(5, 5, 0, 0)
    .build();
  reportSheet.insertChart(chart);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheetId = reportSheet.getSheetId();
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=pdf&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=true&gid=${reportSheetId}`;
  
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    Logger.log('Error: ' + response.getContentText());
    throw new Error('Failed to generate PDF. ' + response.getContentText());
  }

  const pdf = response.getBlob().setName('Project_Report.pdf');
  
  MailApp.sendEmail({
    to: 'piopreety2807@gmail.com',
    subject: 'Project Report',
    body: 'Please find the attached project report.',
    attachments: [pdf]
  });
}

function sendPolicyAcknowledgmentReminder(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Policy');
  const data = sheet.getDataRange().getValues();
  
  let docId;
  
  for (let i = 1; i < data.length; i++) {
    const currentStatus = data[i][1]; // Assuming column B contains the acknowledgment status
    const lastStatus = data[i][3]; // Assuming column D contains the last acknowledgment status
    
    if (currentStatus !== lastStatus && currentStatus !== 'Acknowledged') {
      // Create a temporary document for the policy acknowledgment reminder
      const doc = DocumentApp.create('Policy Acknowledgment Reminder');
      const body = doc.getBody();
      body.appendParagraph('Policy Acknowledgment Reminder');
      body.appendParagraph('Please acknowledge the following policy:');
      body.appendParagraph('\n');
      body.appendParagraph('Policy Details:\n' + data[i][0]); // Assuming column A contains the policy details
      docId = doc.getId();
      
      // Save and close the document
      doc.saveAndClose();
      
      // Generate PDF from the document
      const pdfUrl = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(pdfUrl, {
        headers: {
          'Authorization': `Bearer ${token}`
        },
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        Logger.log('Error: ' + response.getContentText());
        throw new Error('Failed to generate PDF. ' + response.getContentText());
      }

      const pdf = response.getBlob().setName('Policy_Acknowledgment_Reminder.pdf');
      
      // Send email with PDF attachment
      MailApp.sendEmail({
        to: 'piopreety2807@gmail.com',
        subject: 'Policy Acknowledgment Reminder',
        body: 'Please find the attached policy acknowledgment reminder.',
        attachments: [pdf]
      });
      
      // Clean up by deleting the temporary document
      DriveApp.getFileById(docId).setTrashed(true);

      // Update the last acknowledgment status in the sheet
      sheet.getRange(i + 1, 4).setValue(currentStatus); // Update column D
    }
  }
}

// Function to send email notifications
function sendEmailNotification(e) {
  // Get the submitted data from the "Contact Form Responses" sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Contact Form Responses");
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Extract form data based on your column arrangement
  var userName = data[1]; 
  var userEmail = data[2]; 
  var userMessage = data[3]; 
  var recipientEmail = data[4]; 
  
  // Log data for debugging
  Logger.log("User Name: " + userName);
  Logger.log("User Email: " + userEmail);
  Logger.log("User Message: " + userMessage);
  Logger.log("Recipient Email: " + recipientEmail);

  // Validate email addresses
  if (!validateEmail(userEmail)) {
    Logger.log("Invalid user email: " + userEmail);
    return;
  }

  if (!validateEmail(recipientEmail)) {
    Logger.log("Invalid recipient email: " + recipientEmail);
    return;
  }
  
  // Set the email subject and body
  var subject = "New Message from " + userName;
  var body = "You have received a new message from " + userName + " (" + userEmail + ").\n\n" +
             "Message: " + userMessage + "\n\n" +
             "Please respond to the message at your earliest convenience.";
  
  // Send the email
  MailApp.sendEmail(recipientEmail, subject, body);
}

// Function to validate email addresses
function validateEmail(email) {
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

// Trigger the function on form submission
function createTrigger() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('sendEmailNotification')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();
}

// Triggers
function onEdit(e) {
  // Check if the event object is defined
  if (!e) {
    return;
  }
  
  const sheetName = e.source.getActiveSheet().getName();
  if (sheetName === 'Resources') {
    updateWeeklyReport();
    updateResourceUtilizationReport();
  }
}

function createDailyTriggers() {
  // Delete all existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

  // Set up daily trigger for sendTaskReminders
  ScriptApp.newTrigger("sendTaskReminders")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  // Set up weekly trigger for generateWeeklyResourceUsageReport
  ScriptApp.newTrigger("generateWeeklyResourceUsageReport")
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();

  // Set up daily trigger for checkResourceUtilization
  ScriptApp.newTrigger("checkResourceUtilization")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // Set up weekly trigger for generateProjectReport
  ScriptApp.newTrigger("generateProjectReport")
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();

  // Set up weekly trigger for sendPolicyAcknowledgmentReminder
  ScriptApp.newTrigger("sendPolicyAcknowledgmentReminder")
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(10)
    .create();

  Logger.log("Daily and weekly triggers created successfully.");
}
