function createJobApplicationSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = ["Name", "Company", "Contact Email", "Job Title", "Application Date", "Contact Number", "Status", "Response Date", "Notes"];
  
  // Check if headers already exist
  var range = sheet.getRange(1, 1, 1, headers.length);
  var values = range.getValues();
  
  var headerExists = values[0].every(function(cell, index) {
    return cell === headers[index];
  });
  
  if (!headerExists) {
    // Set the headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}


function sendJobApplications() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  createJobApplicationSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  var resumeFileId = '1L96OfiVjiCPia_xfpUuVBaaLo8MDe--W'; // Replace with your file ID
  var resumeFile;
  
  try {
    resumeFile = DriveApp.getFileById(resumeFileId);
  } catch (e) {
    Logger.log('Error accessing file: ' + e.message);
    return; // Exit function if file cannot be accessed
  }

  for (var i = 1; i < values.length; i++) {
    var status = values[i][6]; // Status (column 7)
    var email = values[i][2]; // Contact Email (column 3)
    var name = values[i][0]; // Name (column 1)
    var company = values[i][1]; // Company (column 2)
    var jobTitle = values[i][3]; // Job Title (column 4)

    if (status !== "Sent" && email) {
      var subject = "Excited to Explore Growth Opportunities at  "+ company +" for "+ jobTitle ;
      var body;

      if (name) {
        // If Name is present, use personalized greeting
        body = "Hi " + name + ",\n\nHope you're doing well!\n\nI’m excited to apply for the " + jobTitle + " role at " + company + ". With hands-on experience in Java, Spring Boot, React.js, and SQL, I’m ready to bring my skills to your team.\n\nCheck out my resume and cover letter attached. I’d love to chat about how I can contribute to " + company + ".\n\nThanks for considering my application!\n\nBest regards,\nRasmiRanjan Parida\n6371153178\nrasmiranjanparida77@gmail.com";
      } else {
        // If Name is not present, use generic greeting
        body = "Hi " + company + " Team,\n\nHope you're doing well!\n\nI’m excited to apply for the " + jobTitle + " role at " + company + ". With hands-on experience in Java, Spring Boot, React.js, and SQL, I’m ready to bring my skills to your team.\n\nCheck out my resume and cover letter attached. I’d love to chat about how I can contribute to " + company + ".\n\nThanks for considering my application!\n\nBest regards,\nRasmiRanjan Parida\n6371153178\nrasmiranjanparida77@gmail.com";
      }

      try {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body,
          attachments: [resumeFile.getAs(MimeType.PDF)], // Attach the resume
        });
        Logger.log('Email sent to: ' + email);
        
        sheet.getRange(i + 1, 7).setValue("Sent"); // Update Status to "Sent" (column 7)
        sheet.getRange(i + 1, 5).setValue(new Date()); // Update Application Date (column 5)
        sheet.getRange(i + 1, 8).setValue(formatDate(new Date())); // Set formatted timestamp (column 8)
        
        var sentLabel = GmailApp.getUserLabelByName("Job Applications Sent");
        if (!sentLabel) {
          sentLabel = GmailApp.createLabel("Job Applications Sent");
        }
        var threads = GmailApp.search('to:' + email + ' subject:"' + subject + '"');
        for (var j = 0; j < threads.length; j++) {
          threads[j].addLabel(sentLabel);
        }
      } catch (e) {
        Logger.log('Error sending email: ' + e.message);
      }
    } else {
      Logger.log('Skipping row ' + (i + 1) + ': Email already sent or no email provided');
    }
  }
}








function processResponses() {
  var sentLabel = GmailApp.getUserLabelByName("Job Applications Sent");
  var responseLabel = GmailApp.getUserLabelByName("Job Applications Responses");
  
  if (!responseLabel) {
    responseLabel = GmailApp.createLabel("Job Applications Responses");
  }
  
  var threads = sentLabel.getThreads();
  
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var from = message.getFrom();
      var subject = message.getSubject();
      
      // Check if the response is related to a sent job application
      if (from.includes("company") && subject.startsWith("Re:")) {
        // Move thread to "Job Applications Responses"
        thread.addLabel(responseLabel);
        thread.removeLabel(sentLabel);
        
        // Optional: Add response information to the sheet
        var email = message.getTo(); // Get the email address the response was sent to
        var rowIndex = findRowIndexByEmailAndSubject(email, subject, SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
        if (rowIndex !== -1) {
          SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(rowIndex + 1, 7).setValue(formatDate(new Date())); // Update Response Date
        }
      }
    }
  }
}

// Find the row index based on email and subject
function findRowIndexByEmailAndSubject(email, subject, sheet) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === email && data[i][3] === subject) {
      return i;
    }
  }
  return -1;
}

// Format the date to a readable string
function formatDate(date) {
  var options = { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
  return date.toLocaleDateString("en-US", options);
}
