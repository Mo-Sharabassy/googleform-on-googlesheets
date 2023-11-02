// Function to clear specific cells in the "Form" sheet
function clearCells() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("Form");
    var cellsToClear = ["G8", "E11", "E14", "E17", "E20", "I11", "I14", "I17", "I20"];
  
    for (var i = 0; i < cellsToClear.length; i++) {
      formSheet.getRange(cellsToClear[i]).clearContent();
    }
  }
  
  // Function to submit form data
  function submitForm() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("Form");
  
    // Get values from the form
    var emailCell = formSheet.getRange(8, 7).getValue();
    var tripID = formSheet.getRange(11, 5).getValue();
    var clientCell = formSheet.getRange(14, 5).getValue();
    var supplierCell = formSheet.getRange(17, 5).getValue();
    var driverCell = formSheet.getRange(20, 5).getValue();
    var problemCause = formSheet.getRange(11, 9).getValue();
    var theProblem = formSheet.getRange(14, 9).getValue();
    var furtherInfo = formSheet.getRange(17, 9).getValue();
    var compDirectedTo = formSheet.getRange(20, 9).getValue();
  
    // Validate form inputs
    if (
      emailCell === '' ||
      tripID === '' ||
      clientCell === '' ||
      supplierCell === '' ||
      driverCell === '' ||
      problemCause === '' ||
      theProblem === '' ||
      compDirectedTo === ''
    ) {
      SpreadsheetApp.getUi().alert('Please fill in all required fields.');
    } else {
      // Prepare the data to be saved
      var values = [[
        emailCell,
        tripID,
        clientCell,
        supplierCell,
        driverCell,
        problemCause,
        theProblem,
        furtherInfo,
        compDirectedTo,
      ]];
  
      var depName = compDirectedTo;
      var sentSheet = ss.getSheetByName('Sent');
      var aw = SpreadsheetApp.openByUrl(
        'RECIPIENT_SHEET_URL'
      );
      var receivedSheet = aw.getSheetByName('Received');
  
      // Save data to "Sent" and "Received" sheets
      var assignedTimestamp = sentSheet.getRange(sentSheet.getLastRow() + 1, 10);
      sentSheet.getRange(sentSheet.getLastRow() + 1, 1, 1, 9).setValues(values);
      receivedSheet.getRange(receivedSheet.getLastRow() + 1, 1, 1, 9).setValues(values);
      assignedTimestamp.setValue(new Date());
  
      // Prepare and send an email
      var emailSubject = `Complaint Regarding Client ${clientCell}`;
      var bodyText = `${emailCell} has a complaint regarding\nTrip ID : ${tripID}\nClient : ${clientCell}\nSupplier : ${supplierCell}\nDriver : ${driverCell}\nThe Problem is related to "${problemCause}" and the problem is "${theProblem}"\nFurther Details: ${furtherInfo}`;
  
      sendEmail(depName, emailSubject, bodyText);
  
      clearCells();
      SpreadsheetApp.getUi().alert('Your Complaint Has Been Submitted Successfully');
    }
  }
  
  // Function to send an email
  function sendEmail(email, subject, body) {
    MailApp.sendEmail(email, subject, body);
  }
  