function doGet(e) {
    if (e.parameter.action && e.parameter.row) {
      var template = HtmlService.createTemplateFromFile('Index1');
      template.action = e.parameter.action;
      template.row = e.parameter.row;
      return template.evaluate();
    }
    return HtmlService.createHtmlOutputFromFile('Index');
  }
  
  function handleFormSubmit(formData) {
    Logger.log("Form Data: " + JSON.stringify(formData));
  
    // Validate Email
    if (!formData.email) {
      Logger.log("Email not found in formData: " + JSON.stringify(formData));
      return "Email is required!";
    }
  
    // Regular expression to allow only letters, numbers, spaces, dots, commas, and dashes
    var validPattern = /^[A-Za-z0-9 .,-]+$/;
  
    // Validate other fields
    var fieldsToValidate = ['name', 'department', 'leavingFrom', 'destination', 'event', 'empID'];
    for (var i = 0; i < fieldsToValidate.length; i++) {
      if (!validPattern.test(formData[fieldsToValidate[i]])) {
        return "Invalid characters in " + fieldsToValidate[i].charAt(0).toUpperCase() + fieldsToValidate[i].slice(1) +
               ". Only letters, numbers, spaces, dots (.), commas (,), and dashes (-) are allowed.";
      }
    }
  
    // Validate Date Requested is not in the past
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Set time to midnight
    var dateRequested = new Date(formData.dateRequested);
    if (dateRequested < today) {
      return "Date Requested cannot be earlier than today's date.";
    }
  
    // Validate Leaving From and Destination are not the same
    var leavingFrom = formData.leavingFrom.trim().replace(/\s+/g, '').toLowerCase();
    var destination = formData.destination.trim().replace(/\s+/g, '').toLowerCase();
    if (leavingFrom === destination) {
      return "Leaving From and Destination cannot be the same.";
    }
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var userIP = getUserIP();
  
    sheet.appendRow([
      new Date(),
      formData.name,
      formData.department,
      formData.dateRequested,
      formData.leavingFrom,
      formData.timeLeaving,
      formData.returnDate,
      formData.returnTime,
      formData.destination,
      formData.event,
      formData.occupants,
      formData.extraInfo,
      formData.email,
      formData.empID,
      formData.requestedBy,
      "Pending",
      "",
      "",
      userIP,
      "" // New column for approver's email
    ]);
  
    sendConfirmationEmail(formData);
    sendAdminNotification(formData);
    
    return "Form submitted successfully!";
  }
  
  function getUserIP() {
    var response = UrlFetchApp.fetch("https://api.ipify.org?format=json");
    var json = JSON.parse(response.getContentText());
    return json.ip;
  }
  
  function sendConfirmationEmail(formData) {
    Logger.log("Sending email to: " + formData.email);
  
    var requesterSubject = "Vehicle Request Confirmation";
    var requesterBody = "Dear " + formData.name + ",\n\n" +
                        "Your vehicle request has been successfully submitted. Here are the details:\n\n" +
                        "Emp. ID: " + formData.empID + "\n" +
                        "Date Requested: " + formData.dateRequested + "\n" +
                        "Leaving From: " + formData.leavingFrom + "\n" +
                        "Leaving Time: " + formData.timeLeaving + "\n" +
                        "Purpose: " + formData.event + "\n" +
                        "Return Date: " + formData.returnDate + "\n\n" +
                        "Requested By: " + formData.requestedBy + "\n" +
                        "Your IP Address: " + getUserIP() + "\n\n" +
                        "Thank you,\n" +
                        "Vehicle Management Team";
    
    GmailApp.sendEmail(formData.email, requesterSubject, requesterBody);
    Logger.log("Sent confirmation email to requester.");
  }
  
  function sendAdminNotification(formData) {
    var adminEmail = "anubhavaman@ghcl.co.in";
    var allRecipients = [adminEmail];
  
    Logger.log("Sending email to admin: " + allRecipients.join(","));
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
  
    var adminSubject = "New Vehicle Request Submitted";
  
    var adminBody = "<html><body>" +
                    "<p>A new vehicle request has been submitted. Here are the details:</p>" +
                    "<p>Name: " + formData.name + "<br>" +
                    "Email: " + formData.email + "<br>" +
                    "Department: " + formData.department + "<br>" +
                    "Emp. ID: " + formData.empID + "<br>" +
                    "Date Requested: " + formData.dateRequested + "<br>" +
                    "Leaving From: " + formData.leavingFrom + "<br>" +
                    "Leaving Time: " + formData.timeLeaving + "<br>" +
                    "Purpose: " + formData.event + "<br>" +
                    "Return Date: " + formData.returnDate + "<br>" +
                    "Requested By: " + formData.requestedBy + "<br>" +
                    "User IP Address: " + getUserIP() + "</p>" +
                    "<p><b>Action Required:</b></p>" +
                    "<p><a href='" + ScriptApp.getService().getUrl() + "?action=approve&row=" + lastRow + "' " +
                    "style='color: white; background-color: green; padding: 10px 20px; text-decoration: none; border-radius: 5px;'>Approve</a> " +
                    "<a href='" + ScriptApp.getService().getUrl() + "?action=deny&row=" + lastRow + "' " +
                    "style='color: white; background-color: red; padding: 10px 20px; text-decoration: none; border-radius: 5px;'>Reject</a></p>" +
                    "<p>Note: This email can be forwarded to other authorized personnel for approval or rejection.</p>" +
                    "</body></html>";
  
    GmailApp.sendEmail(allRecipients.join(","), adminSubject, "", {
      htmlBody: adminBody
    });
  
    Logger.log("Sent notification email to admin.");
  }
  
  function handleAdminFormSubmit(action, row, comments) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
    var requesterEmail = data[12];
    var requesterName = data[1];
    var adminEmail = "anubhavaman@ghcl.co.in";
  
    var subject = "Vehicle Request " + (action === "approve" ? "Approved" : "Denied");
    var body = "Dear " + requesterName + ",\n\n" +
               "Your vehicle request has been " + (action === "approve" ? "approved" : "denied") + ".\n\n" +
               "Comments: " + comments + "\n\n";
  
    var timestamp = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd HH:mm:ss");
    sheet.getRange(row, 16).setValue(action === "approve" ? "Approved" : "Denied");
    sheet.getRange(row, 17).setValue(comments);
    sheet.getRange(row, 18).setValue(timestamp);
    
    var approverEmail = Session.getActiveUser().getEmail();
    sheet.getRange(row, 20).setValue(approverEmail);
  
    var pdfLink = "";
    if (action === "approve") {
      var pdfBlob = generatePdf(data, action, comments, data[18], timestamp, approverEmail);
      var pdfFile = savePdfToDrive(pdfBlob, data[13], timestamp);
      pdfLink = pdfFile.getUrl();
      body += "You can access the approved request PDF here: " + pdfLink + "\n\n";
    }
  
    body += "Thank you,\n" +
            "Vehicle Management Team";
  
    GmailApp.sendEmail(requesterEmail, subject, body);
  
    // Send confirmation to the approver
    var approverSubject = "Confirmation: Vehicle Request " + (action === "approve" ? "Approved" : "Denied");
    var approverBody = "You have " + (action === "approve" ? "approved" : "denied") + " a vehicle request.\n\n" +
                       "Requester: " + requesterName + "\n" +
                       "Emp. ID: " + data[13] + "\n" +
                       "Date Requested: " + Utilities.formatDate(new Date(data[3]), "GMT+5:30", "yyyy-MM-dd") + "\n" +
                       "Comments: " + comments + "\n\n";
    
    if (pdfLink) {
      approverBody += "You can access the approved request PDF here: " + pdfLink + "\n\n";
    }
    
    approverBody += "Thank you for your action.";
    
    GmailApp.sendEmail(approverEmail, approverSubject, approverBody);
  
    // Send notification to admin
    var adminSubject = "Vehicle Request " + (action === "approve" ? "Approved" : "Denied") + " by " + approverEmail;
    var adminBody = "A vehicle request has been " + (action === "approve" ? "approved" : "denied") + " by " + approverEmail + ".\n\n" +
                    "Requester: " + requesterName + "\n" +
                    "Emp. ID: " + data[13] + "\n" +
                    "Date Requested: " + Utilities.formatDate(new Date(data[3]), "GMT+5:30", "yyyy-MM-dd") + "\n" +
                    "Comments: " + comments + "\n\n";
    
    if (pdfLink) {
      adminBody += "You can access the approved request PDF here: " + pdfLink + "\n\n";
    }
    
    GmailApp.sendEmail(adminEmail, adminSubject, adminBody);
  
    return "Action recorded. Emails have been sent to the requester, approver, and admin.";
  }
  
  function generatePdf(data, action, comments, userIP, approvalTime, approverEmail) {
    var formatDate = function(dateString) {
      var date = new Date(dateString);
      return Utilities.formatDate(date, "GMT+5:30", "yyyy-MM-dd");
    };
  
    var formatTime = function(timeString) {
      var date = new Date(timeString);
      return Utilities.formatDate(date, "GMT+5:30", "HH:mm:ss");
    };
  
    var htmlContent = "<html><body>";
    htmlContent += "<h2>Vehicle Request Details</h2>";
    htmlContent += "<p><b>Emp ID:</b> " + data[13] + "</p>";
    htmlContent += "<p><b>Name:</b> " + data[1] + "</p>";
    htmlContent += "<p><b>Department:</b> " + data[2] + "</p>";
    htmlContent += "<p><b>Email:</b> " + data[12] + "</p>";
    htmlContent += "<p><b>Date Requested:</b> " + formatDate(data[3]) + "</p>";
    htmlContent += "<p><b>Leaving From:</b> " + data[4] + "</p>";
    htmlContent += "<p><b>Leaving Time:</b> " + formatTime(data[5]) + "</p>";
    htmlContent += "<p><b>Return Date:</b> " + formatDate(data[6]) + "</p>";
    htmlContent += "<p><b>Return Time:</b> " + formatTime(data[7]) + "</p>";
    htmlContent += "<p><b>Destination:</b> " + data[8] + "</p>";
    htmlContent += "<p><b>Purpose:</b> " + data[9] + "</p>";
    htmlContent += "<p><b>Occupants:</b> " + data[10] + "</p>";
    htmlContent += "<p><b>Extra Information:</b> " + data[11] + "</p>";
    htmlContent += "<p><b>Requested By:</b> " + data[14] + "</p>";
    htmlContent += "<p><b>Action:</b> " + (action === "approve" ? "Approved" : "Denied") + "</p>";
    htmlContent += "<p><b>Admin Comments:</b> " + comments + "</p>";
    htmlContent += "<p><b>User IP Address:</b> " + userIP + "</p>";
    htmlContent += "<p><b>Approval Time:</b> " + approvalTime + "</p>";
    htmlContent += "<p><b>Approved By:</b> " + approverEmail + "</p>";
    htmlContent += "</body></html>";
  
    return HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
  }
  
  function savePdfToDrive(pdfBlob, empId, timestamp) {
    var folder = DriveApp.getFolderById("1I_pEwlAKaAwdB8MuRI_LbNSI-Lj1axWs");
    var fileName = empId + "_" + Utilities.formatDate(new Date(timestamp), "GMT+5:30", "yyyy-MM-dd_HH-mm-ss") + ".pdf";
    var file = folder.createFile(pdfBlob).setName(fileName);
    Logger.log("PDF saved to Drive: " + fileName);
    return file;
  }
  
  function sendAccumulatedRequests() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    var adminEmail = "anubhavaman@ghcl.co.in";
    var allRecipients = [adminEmail];
  
    var subject = "Daily Vehicle Requests Summary";
  
    var body = "<html><body>";
    body += "<h3>Here are the vehicle requests for today:</h3>";
    body += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"; 
    body += "<tr><th>Name</th><th>Email</th><th>Department</th><th>Emp. ID</th><th>Date Requested</th>" +
            "<th>Leaving From</th><th>Leaving Time</th><th>Purpose</th><th>Return Date</th><th>Requested By</th><th>Action Status</th><th>Approval Time</th><th>User IP Address</th><th>Approver Email</th></tr>";
  
    var requestsToday = false;
    
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var timestamp = new Date(row[0]);
        timestamp.setHours(0, 0, 0, 0);
        
        if (timestamp.getTime() === today.getTime()) {
          requestsToday = true;
          body += "<tr>";
          body += "<td>" + row[1] + "</td>";
            body += "<td>" + row[12] + "</td>";
            body += "<td>" + row[2] + "</td>";
            body += "<td>" + row[13] + "</td>";
            body += "<td>" + Utilities.formatDate(new Date(row[3]), "GMT+5:30", "yyyy-MM-dd") + "</td>";
            body += "<td>" + row[4] + "</td>";
            body += "<td>" + Utilities.formatDate(new Date(row[5]), "GMT+5:30", "HH:mm:ss") + "</td>";
            body += "<td>" + row[9] + "</td>";
            body += "<td>" + Utilities.formatDate(new Date(row[6]), "GMT+5:30", "yyyy-MM-dd") + "</td>";
            body += "<td>" + row[14] + "</td>";
            body += "<td>" + (row[15] || "Pending") + "</td>";
            body += "<td>" + (row[17] ? Utilities.formatDate(new Date(row[17]), "GMT+5:30", "yyyy-MM-dd HH:mm:ss") : "N/A") + "</td>";
            body += "<td>" + row[18] + "</td>";
            body += "</tr>";
          }
        }
      
        body += "</table>";
        body += "</body></html>";
      
        if (requestsToday) {
          GmailApp.sendEmail(allRecipients.join(","), subject, "", {
            htmlBody: body
          });
          Logger.log("Sent daily summary email to admin.");
        } else {
          Logger.log("No requests made today.");
        }
      }
      
      function createTimeDrivenTrigger() {
        ScriptApp.newTrigger("sendAccumulatedRequests")
                 .timeBased()
                 .everyDays(1)
                 .atHour(23)
                 .nearMinute(10)
                 .create();
        }
        
        function getUserEmail() {
          return Session.getActiveUser().getEmail();
        }