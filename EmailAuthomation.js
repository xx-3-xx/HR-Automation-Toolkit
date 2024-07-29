function sendInterviewEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interview Schedule');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
   Assuming headers are in the first row
  var headers = data[0];
  
   Get the column indexes
  var fullNameIdx = headers.indexOf('Candidate Full Name');
  var positionIdx = headers.indexOf('Position Applied For');
  var emailIdx = headers.indexOf('Email Address');
  var interviewDateIdx = headers.indexOf('Interview Date');
  var interviewTimeIdx = headers.indexOf('Interview Time');
  var interviewerIdx = headers.indexOf('Interviewer(s)');
  var interviewTypeIdx = headers.indexOf('Interview Type');
  var interviewLocationIdx = headers.indexOf('Interview Location');
  var confirmationSentIdx = headers.indexOf('Confirmation sent');
  var confirmationSentDateIdx = headers.indexOf('Confirmation Sent Date');
  
   Get the time zone of the spreadsheet
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
   Loop through the data rows
  for (var i = 1; i  data.length; i++) {
    var row = data[i];
    var confirmationSent = row[confirmationSentIdx];
    
     If confirmation email not sent
    if (confirmationSent !== 'Yes') {
      var fullName = row[fullNameIdx];
      var position = row[positionIdx];
      var email = row[emailIdx];
      var interviewDate = new Date(row[interviewDateIdx]);
      var interviewTime = new Date(row[interviewTimeIdx]);
      var interviewer = row[interviewerIdx];
      var interviewType = row[interviewTypeIdx];
      var interviewLocation = row[interviewLocationIdx];
      
       Format date and time using the spreadsheet's time zone
      var formattedDate = Utilities.formatDate(interviewDate, timeZone, 'yyyy-MM-dd');
      var formattedTime = Utilities.formatDate(interviewTime, timeZone, 'HHmm');
      
       Compose the email
      var subject = 'Interview Confirmation';
      var body = 'Dear ' + fullName + ',nn' +
        'We are pleased to invite you to an interview for the position of ' + position + '. nn' +
        'Details of your interview are as followsn' +
        'Date ' + formattedDate + 'n' +
        'Time ' + formattedTime + 'n' +
        'Interviewer(s) ' + interviewer + 'n' +
        'Type ' + interviewType + 'n' +
        'Location ' + interviewLocation + 'nn' +
        'Please confirm your availability by replying to this email within 3 days.nn' +
        'Best regards,nHR Team';
      
       Send the email
      MailApp.sendEmail(email, subject, body);
      
       Mark confirmation sent and record the date
      sheet.getRange(i + 1, confirmationSentIdx + 1).setValue('Yes');
      sheet.getRange(i + 1, confirmationSentDateIdx + 1).setValue(new Date());
    }
  }
}

function sendFollowUpEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interview Schedule');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
   Assuming headers are in the first row
  var headers = data[0];
  
   Get the column indexes
  var fullNameIdx = headers.indexOf('Candidate Full Name');
  var positionIdx = headers.indexOf('Position Applied For');
  var emailIdx = headers.indexOf('Email Address');
  var confirmationSentIdx = headers.indexOf('Confirmation sent');
  var confirmationSentDateIdx = headers.indexOf('Confirmation Sent Date');
  var followUpRequiredIdx = headers.indexOf('Follow-up required');
  var followUpSentIdx = headers.indexOf('Follow-up sent');
  var followUpSentDateIdx = headers.indexOf('Follow-up Sent Date');
  
   Get the time zone of the spreadsheet
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
   Loop through the data rows
  for (var i = 1; i  data.length; i++) {
    var row = data[i];
    var confirmationSent = row[confirmationSentIdx];
    var followUpRequired = row[followUpRequiredIdx];
    var followUpSent = row[followUpSentIdx];
    var confirmationSentDate = new Date(row[confirmationSentDateIdx]);
    
     Calculate the date difference
    var currentDate = new Date();
    var dateDifference = Math.floor((currentDate - confirmationSentDate)  (1000  60  60  24));
    
     If confirmation email sent and 3 days have passed
    if (confirmationSent === 'Yes' && followUpSent !== 'Yes' && dateDifference = 3) {
      var fullName = row[fullNameIdx];
      var position = row[positionIdx];
      var email = row[emailIdx];
      
       Mark follow-up as required
      sheet.getRange(i + 1, followUpRequiredIdx + 1).setValue('Yes');
      
       Compose the follow-up email
      var subject = 'Follow-up Interview Confirmation Needed';
      var body = 'Dear ' + fullName + ',nn' +
        'This is a friendly follow-up to remind you to confirm your interview for the position of ' + position + '. nn' +
        'Please check your previous email for the interview details and confirm your availability as soon as possible.nn' +
        'If you have any questions or need to reschedule, please let us know.nn' +
        'Thank you and looking forward to your confirmation.nn' +
        'Best regards,nHR Team';
      
       Send the follow-up email
      MailApp.sendEmail(email, subject, body);
      
       Mark follow-up as sent and record the date
      sheet.getRange(i + 1, followUpSentIdx + 1).setValue('Yes');
      sheet.getRange(i + 1, followUpSentDateIdx + 1).setValue(new Date());
    }
  }
}

function sendDisqualificationEmails() {
  var interviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Interview Schedule'); 
  var applicationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job Application Form Responses'); 

  var interviewDataRange = interviewSheet.getDataRange();
  var interviewData = interviewDataRange.getValues();
  
  var applicationDataRange = applicationSheet.getDataRange();
  var applicationData = applicationDataRange.getValues();
  
   Assuming headers are in the first row
  var interviewHeaders = interviewData[0];
  var applicationHeaders = applicationData[0];
  
   Get the column indexes for Interview Schedule
  var fullNameIdx = interviewHeaders.indexOf('Candidate Full Name');
  var positionIdx = interviewHeaders.indexOf('Position Applied For');
  var emailIdx = interviewHeaders.indexOf('Email Address');
  var followUpSentIdx = interviewHeaders.indexOf('Follow-up sent');
  var followUpSentDateIdx = interviewHeaders.indexOf('Follow-up Sent Date');
  var disqualifiedIdx = interviewHeaders.indexOf('Disqualified');
  var disqualificationSentDateIdx = interviewHeaders.indexOf('Disqualification Sent Date');
  
   Get the column indexes for Job Application Form Responses
  var appEmailIdx = applicationHeaders.indexOf('Email Address');
  var appInterviewStatusIdx = applicationHeaders.indexOf('Interview Status');
  
   Get the time zone of the spreadsheet
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
   Loop through the Interview Schedule data rows
  for (var i = 1; i  interviewData.length; i++) {
    var interviewRow = interviewData[i];
    var followUpSent = interviewRow[followUpSentIdx];
    var followUpSentDate = new Date(interviewRow[followUpSentDateIdx]);
    var disqualified = interviewRow[disqualifiedIdx];
    
     Calculate the date difference
    var currentDate = new Date();
    var dateDifference = Math.floor((currentDate - followUpSentDate)  (1000  60  60  24));
    
     If follow-up email sent and 3 days have passed without response
    if (followUpSent === 'Yes' && dateDifference = 3 && disqualified !== 'Yes') {
      var fullName = interviewRow[fullNameIdx];
      var position = interviewRow[positionIdx];
      var email = interviewRow[emailIdx];
      
       Compose the disqualification email
      var subject = 'Interview Disqualification';
      var body = 'Dear ' + fullName + ',nn' +
        'We regret to inform you that we have not received a response to our follow-up email regarding your interview for the position of ' + position + '. nn' +
        'As a result, we are disqualifying you from the interview process.nn' +
        'Thank you for your interest in the position and we wish you all the best in your future endeavors.nn' +
        'Best regards,nHR Team';
      
       Send the disqualification email
      MailApp.sendEmail(email, subject, body);
      
       Mark the candidate as disqualified and record the date in the Interview Schedule sheet
      interviewSheet.getRange(i + 1, disqualifiedIdx + 1).setValue('Yes');
      interviewSheet.getRange(i + 1, disqualificationSentDateIdx + 1).setValue(new Date());
      
       Find the corresponding row in the Job Application Form Responses sheet and update the Interview Status
      for (var j = 1; j  applicationData.length; j++) {
        var applicationRow = applicationData[j];
        var applicationEmail = applicationRow[appEmailIdx];
        
        if (applicationEmail === email) {
          applicationSheet.getRange(j + 1, appInterviewStatusIdx + 1).setValue('Disqualified');
          break;
        }
      }
    }
  }
}

function processInterviewResults() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Job Application Form Responses');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
   Assuming headers are in the first row
  var headers = data[0];
  
   Get the column indexes
  var fullNameIdx = headers.indexOf('Full Name (As per IC)');
  var emailIdx = headers.indexOf('Email Address');
  var positionIdx = headers.indexOf('Position Applied For');
  var interviewResultIdx = headers.indexOf('Interview Result');
  var rejectionSentDateIdx = headers.indexOf('Rejection Letter Sent Date');
  var offerSentDateIdx = headers.indexOf('Offer Letter Sent Date');
  
   Get the time zone of the spreadsheet
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  
   Loop through the data rows
  for (var i = 1; i  data.length; i++) {
    var row = data[i];
    var interviewResult = row[interviewResultIdx];
    var rejectionSentDate = row[rejectionSentDateIdx];
    var offerSentDate = row[offerSentDateIdx];
    
    var fullName = row[fullNameIdx];
    var email = row[emailIdx];
    var position = row[positionIdx];
    
    if (interviewResult === 'Fail' && !rejectionSentDate) {
       Compose the rejection email
      var subject = 'Interview Result - ' + position;
      var body = 'Dear ' + fullName + ',nn' +
        'Thank you for your interest in the ' + position + ' position at our company. We appreciate the time you spent with us during the interview process.nn' +
        'After careful consideration, we regret to inform you that we have decided not to proceed with your application at this time.nn' +
        'We encourage you to apply for future openings that match your skills and experience.nn' +
        'Thank you again for your interest in our company, and we wish you all the best in your future endeavors.nn' +
        'Best regards,nHR Team';
      
       Send the rejection email
      MailApp.sendEmail(email, subject, body);
      
       Record the rejection sent date
      sheet.getRange(i + 1, rejectionSentDateIdx + 1).setValue(new Date());
    } else if (interviewResult === 'Pass' && !offerSentDate) {
       Compose the offer email
      var subject = 'Job Offer - ' + position;
      var body = 'Dear ' + fullName + ',nn' +
        'Congratulations! We are pleased to offer you the position of ' + position + ' at our company.nn' +
        'Please review the attached offer letter and confirm your acceptance by replying to this email.nn' +
        'We look forward to welcoming you to our team and working with you.nn' +
        'Best regards,nHR Team';
      
       Send the offer email
      MailApp.sendEmail(email, subject, body);
      
       Record the offer sent date
      sheet.getRange(i + 1, offerSentDateIdx + 1).setValue(new Date());
    }
  }
}
