function createMergedDocsFromTemplate() {
  // Open the Google Sheets document and select the data range
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Open the Google Docs template
  var templateId = '###################'; // Replace with your template ID
  var templateFile = DriveApp.getFileById(templateId);
  var folder = DriveApp.getFolderById('###############'); // Replace with the target folder ID

  // Get header row to map columns
  var headers = data[0];
  
  // Loop through each row of data, starting from the second row (index 1)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var copyFile = templateFile.makeCopy('Requirement - ' + row[0] + ' v1', folder); // Change this to the desired naming format (In this case it will be 'Requirement - Lemon Co v1')
    var copyDoc = DocumentApp.openById(copyFile.getId());
    var body = copyDoc.getBody();
    var header = copyDoc.getHeader();
    
    // Replace placeholders in the document with actual values (Placeholder example: {{ Organization }})
    for (var j = 0; j < headers.length; j++) {
      var placeholder = '{{' + headers[j] + '}}';
      var value = row[j];

        body.replaceText(placeholder, value);

        if (header){
          header.replaceText(placeholder, value);
        }

    }
    Logger.log(row[0] +' v1 done cloned');
    // Save and close the document
    copyDoc.saveAndClose();
  }

  Logger.log('Documents have been created and saved in the specified folder.');
}
