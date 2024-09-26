function OliviasubmitWeeklyPlan() {
    // Access the active spreadsheet and the two relevant sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Olivia_Weekly Plan");
    var weeklyReportSheet = ss.getSheetByName("Olivia_Weekly Report");
  
    // Retrieve key information from the Weekly Plan header
    var staffName = weeklyPlanSheet.getRange("B6").getValue();
    var reportingDate = weeklyPlanSheet.getRange("B7").getValue();
    var reportingMonth = weeklyPlanSheet.getRange("B8").getValue();
    var reportingWeek = weeklyPlanSheet.getRange("B9").getValue();
    
    // Get the range of project data from Weekly Plan (A33:F1002)
    var planDataRange = weeklyPlanSheet.getRange(33, 1, weeklyPlanSheet.getLastRow() - 32, 6);
    var planData = planDataRange.getValues();
  
    // Find the first empty row in the Weekly Report (checking column C - Reporting Date)
    var reportDataRange = weeklyReportSheet.getRange(3, 3, weeklyReportSheet.getLastRow() - 2, 1);
    var reportData = reportDataRange.getValues();
    var firstEmptyRow = -1;
    
    for (var i = 0; i < reportData.length; i++) {
      if (!reportData[i][0]) {
        firstEmptyRow = i + 3; // Adjusting for row offset (row 3 onwards)
        break;
      }
    }
  
    if (firstEmptyRow === -1) {
      SpreadsheetApp.getUi().alert("No empty row found in the Olivia_Weekly Report sheet.");
      return;
    }
  
    // Loop through each row in the Weekly Plan and map to the Weekly Report, skipping Column G
    for (var j = 0; j < planData.length; j++) {
      if (planData[j][0]) { // Ensure there is a project name before proceeding
        weeklyReportSheet.getRange(firstEmptyRow + j, 3).setValue(reportingDate); // Reporting Date (Column C)
        weeklyReportSheet.getRange(firstEmptyRow + j, 4).setValue(reportingMonth); // Reporting Month (Column D)
        weeklyReportSheet.getRange(firstEmptyRow + j, 5).setValue(reportingWeek); // Reporting Week (Column E)
        weeklyReportSheet.getRange(firstEmptyRow + j, 6).setValue(planData[j][0]); // Project Name (Column F)
        // Skipping Column G (Project Code) because it has a formula
        weeklyReportSheet.getRange(firstEmptyRow + j, 8).setValue(planData[j][1]); // Planned Activities (Column H)
      }
    }
  
    // Generate PDF for the Olivia_Weekly Plan
    var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
    OliviageneratePDF(weeklyPlanSheet, pdfName);
    
    // Notify the user that the data has been submitted
    SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
  }
  
  function OliviageneratePDF() {
    try {
      // Get the active spreadsheet
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var weeklyPlanSheet = ss.getSheetByName("Olivia_Weekly Plan");
      
      // Check if the sheet exists
      if (!weeklyPlanSheet) {
        throw new Error("Sheet with name 'Olivia_Weekly Plan' not found.");
      }
      
      // Retrieve key information for the PDF name
      var staffName = weeklyPlanSheet.getRange("B6").getValue().replace(/\s+/g, '_');
      var reportingDate = weeklyPlanSheet.getRange("B7").getValue();
      var reportingYear = reportingDate.getFullYear();
      var reportingWeek = weeklyPlanSheet.getRange("B9").getValue();
      
      // Construct the PDF name
      var pdfName = staffName + "_" + reportingYear + "_Wk" + reportingWeek + ".pdf";
      
      // Define the URL for exporting the sheet as PDF
      var sheetId = weeklyPlanSheet.getSheetId();
      var url = ss.getUrl().replace(/edit$/, '') + 'export?format=pdf&gid=' + sheetId + '&range=A1:F50';
      
      // Fetch the PDF as a blob
      var response = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
      });
      
      // Create a PDF file from the blob
      var pdfBlob = response.getBlob().setName(pdfName);
      
      // Get the folder by its ID and save the PDF
      var folderId = "128NMK42iUUIToTgsCXE21x6pgnZ9-e0R"; // Replace with your folder ID
      var folder = DriveApp.getFolderById(folderId);
      folder.createFile(pdfBlob);
      
      Logger.log("PDF generated and saved successfully: " + pdfName);
    } catch (e) {
      Logger.log("Error generating PDF: " + e.toString());
      SpreadsheetApp.getUi().alert("An error occurred while generating the PDF: " + e.message);
    }
  }