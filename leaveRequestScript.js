function onFormSubmit(e) {
  var responses = e.values;
  var leaveDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leave Data");
  var employeeDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employee Data");
  // or if the employee data is on another shhet, then use the code below
  // var employeeDataSheet = SpreadsheetApp.openById("Employee Data").getSheetByName("Employee Data");
  
  // Check if the header row exists, if not, add it
  if (leaveDataSheet.getLastRow() === 0) {
    leaveDataSheet.appendRow([
      "Employee Name",
      "Employee ID",
      "Employee Email",
      "Leave Type",
      "Start Date",
      "End Date",
      "Proof of Leave",
      "Additional Note",
      "Status"
    ]);
  }

  // Check if the dates are valid
  if (!isValidDate(new Date(responses[5]), new Date(responses[6]))) {
    // Set status as "Rejected" due to invalid dates
    leaveDataSheet.appendRow([
      responses[1], // Employee Name
      responses[2], // Employee ID
      responses[3], // Employee Email
      responses[4], // Leave Type
      responses[5], // Start Date
      responses[6], // End Date
      responses[7], // Proof of Leave
      responses[8], // Additional Note
      "Rejected, Invalid Dates" // Status
    ]);
    sendEmailNotification(responses[3], "rejected due to invalid dates", responses[4], responses[5], responses[6]);
    return;
  }
  
  // Append new leave request to Leave Data sheet
  leaveDataSheet.appendRow([
    responses[1], // Employee Name
    responses[2], // Employee ID
    responses[3], // Employee Email
    responses[4], // Leave Type
    responses[5], // Start Date
    responses[6], // End Date
    responses[7], // Proof of Leave
    responses[8], // Additional Note
    "Pending"     // Status
  ]); 

  // Get the row number of the newly appended row
  var row = leaveDataSheet.getLastRow();
  
  // Auto approve or reject the leave if it's sick leave or annual leave
  // Else send an email to HR to let them decide whether to approve or reject
  if (responses[4] != 'Sick Leave' && responses[4] != 'Annual Leave') {
    sendHRNotification(row, leaveDataSheet);
  } else {
    autoApproveOrRejectLeave(row, responses[1], responses[2], employeeDataSheet, leaveDataSheet);
  }
}


// Function to check if start date and end date are valid
function isValidDate(startDate, endDate) {
  var currentDate = new Date();
  // Check if start date is in the future
  if (startDate < currentDate) {
    return false;
  }
  // Check if end date is before start date
  if (endDate < startDate) {
    return false;
  }
  return true;
}

// Function to auto-approve or reject leave based on available balance
function autoApproveOrRejectLeave(row, employeeName, employeeId, employeeDataSheet, leaveDataSheet) {
  var employeeRange = employeeDataSheet.getDataRange().getValues();
  var leaveStatusCell = leaveDataSheet.getRange(row, 9); 
  var startDate = new Date(leaveDataSheet.getRange(row, 5).getValue());
  var endDate = new Date(leaveDataSheet.getRange(row, 6).getValue());

  // Calculate the number of leave days
  var leaveDays = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24));

  // Log details for debugging
  Logger.log("Employee ID from form: " + employeeId);
  Logger.log("Leave Days: " + leaveDays);

  for (var i = 1; i < employeeRange.length; i++) { // Skip header row
    var currentEmployeeId = employeeRange[i][0].toString().trim();
    var currentEmployeeName = employeeRange[i][1].toUpperCase().trim();
    Logger.log("Employee ID in sheet: " + currentEmployeeId + currentEmployeeName + employeeName + employeeName.toUpperCase());

    if (currentEmployeeId == employeeId.toString().trim() && currentEmployeeName == employeeName.toUpperCase().trim()) {
      var leaveType = leaveDataSheet.getRange(row, 4).getValue();
      var availableLeave = parseInt(employeeRange[i][3]); 
      if (leaveType == 'Sick Leave') {
        availableLeave = parseInt(employeeRange[i][4]); 
      }
      Logger.log("Available Leave: " + availableLeave);

      if (availableLeave >= leaveDays) {
        leaveStatusCell.setValue("Approved");

        // Update available leave in Employee Data sheet
        employeeDataSheet.getRange(i + 1, 4).setValue(availableLeave - leaveDays);

        sendEmailNotification(
          leaveDataSheet.getRange(row, 3).getValue(), // Employee Email
          "approved",
          leaveType, 
          leaveDataSheet.getRange(row, 5).getValue(), // Start Date
          leaveDataSheet.getRange(row, 6).getValue()  // End Date
        );
      } else {
        leaveStatusCell.setValue("Rejected, Insufficient Leave Balance");
        sendEmailNotification(
          leaveDataSheet.getRange(row, 3).getValue(), // Employee Email
          "rejected due to insufficient leave balance",
          leaveType, // Leave Type
          leaveDataSheet.getRange(row, 5).getValue(), // Start Date
          leaveDataSheet.getRange(row, 6).getValue()  // End Date
        );
      }
      return;
    }
  }

  // If employee ID or name not found, set status as "Rejected"
  leaveStatusCell.setValue("Rejected, Invalid employee ID or name");
  sendEmailNotification(
    leaveDataSheet.getRange(row, 3).getValue(), // Employee Email
    "rejected due to invalid employee ID or name",
    leaveDataSheet.getRange(row, 4).getValue(), // Leave Type
    leaveDataSheet.getRange(row, 5).getValue(), // Start Date
    leaveDataSheet.getRange(row, 6).getValue()  // End Date
  );
}

// Function to send email notification
function sendEmailNotification(email, status, leaveType, startDate, endDate) {
  var subject = "Leave Request " + status;
  var message = "Your " + leaveType + " leave request from " + startDate + " to " + endDate + " has been " + status + ".";
  MailApp.sendEmail(email, subject, message);
}

// Function to notify HR for "Other" leave types
function sendHRNotification(row, leaveDataSheet) {
  var hrEmail = "hr@example.gmail.com"; // Replace with HR's email address
  var employeeName = leaveDataSheet.getRange(row, 1).getValue();
  var employeeEmail = leaveDataSheet.getRange(row, 3).getValue();
  var leaveType = leaveDataSheet.getRange(row, 4).getValue();
  var startDate = leaveDataSheet.getRange(row, 5).getValue();
  var endDate = leaveDataSheet.getRange(row, 6).getValue();
  var additionalNote = leaveDataSheet.getRange(row, 7).getValue();
  var evidence = leaveDataSheet.getRange(row, 8).getValue();
  
  var subject = "Leave Request for Approval - " + leaveType;
  var message = "A new leave request has been submitted by " + employeeName + " " + employeeEmail + ".\n\n" +
                "Leave Type: " + leaveType + "\n" +
                "Start Date: " + startDate + "\n" +
                "End Date: " + endDate + "\n" +
                "Additional Note: " + additionalNote + "\n" +
                "Supported Evidence: " + evidence + "\n\n" +
                "Please review and approve or reject the request.";
  
  MailApp.sendEmail(hrEmail, subject, message);
  leaveDataSheet.getRange(row, 9).setValue("Pending HR Approval");
}

