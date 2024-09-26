//..............Blaise.........................................................

function blaisesubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Blaise_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Blaise_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Blaise_Weekly Report sheet.");
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

  // Generate PDF for the Blaise_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  blaisegeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function blaisegeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Blaise_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Blaise_Weekly Plan' not found.");
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

//.............Daniel.........................................

function danielsubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Daniel_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Daniel_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Daniel_Weekly Report sheet.");
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

  // Generate PDF for the Daniel_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  danielgeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function danielgeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Daniel_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Daniel_Weekly Plan' not found.");
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

//.............Margo........................................

function MargosubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Margo_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Margo_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Margo_Weekly Report sheet.");
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

  // Generate PDF for the Margo_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  MargogeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function MargogeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Margo_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Margo_Weekly Plan' not found.");
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

//.............Amedee........................................

function AmedeesubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Amedee_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Amedee_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Amedee_Weekly Report sheet.");
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

  // Generate PDF for the Amedee_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  AmedeegeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function AmedeegeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Amedee_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Amedee_Weekly Plan' not found.");
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

//.............Irene........................................

function IrenesubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Irene_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Irene_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Irene_Weekly Report sheet.");
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

  // Generate PDF for the Irene_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  IrenegeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function IrenegeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Irene_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Irene_Weekly Plan' not found.");
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

//.............Gaston........................................

function GastonsubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Gaston_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Gaston_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Gaston_Weekly Report sheet.");
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

  // Generate PDF for the Gaston_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  GastongeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function GastongeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Gaston_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Gaston_Weekly Plan' not found.");
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

//.............Gisele........................................

function GiselesubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Gisele_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Gisele_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Gisele_Weekly Report sheet.");
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

  // Generate PDF for the Gisele_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  GiselegeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function GiselegeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Gisele_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Gisele_Weekly Plan' not found.");
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

//.............Augustin........................................

function AugustinsubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Augustin_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Augustin_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Augustin_Weekly Report sheet.");
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

  // Generate PDF for the Augustin_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  AugustingeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function AugustingeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Augustin_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Augustin_Weekly Plan' not found.");
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

//.............Charlse........................................

function CharlsesubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("Charlse_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("Charlse_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the Charlse_Weekly Report sheet.");
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

  // Generate PDF for the Charlse_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  CharlsegeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function CharlsegeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("Charlse_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'Charlse_Weekly Plan' not found.");
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

//.............JPN........................................

function JPNsubmitWeeklyPlan() {
  // Access the active spreadsheet and the two relevant sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var weeklyPlanSheet = ss.getSheetByName("JPN_Weekly Plan");
  var weeklyReportSheet = ss.getSheetByName("JPN_Weekly Report");

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
    SpreadsheetApp.getUi().alert("No empty row found in the JPN_Weekly Report sheet.");
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

  // Generate PDF for the JPN_Weekly Plan
  var pdfName = staffName + "_" + reportingDate.getFullYear() + "Wk " + reportingWeek;
  JPNgeneratePDF(weeklyPlanSheet, pdfName);
  
  // Notify the user that the data has been submitted
  SpreadsheetApp.getUi().alert("Weekly Plan data has been successfully submitted to the Weekly Report and saved as a PDF.");
}

function JPNgeneratePDF() {
  try {
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var weeklyPlanSheet = ss.getSheetByName("JPN_Weekly Plan");
    
    // Check if the sheet exists
    if (!weeklyPlanSheet) {
      throw new Error("Sheet with name 'JPN_Weekly Plan' not found.");
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

//.............Olivia........................................

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
// Weekly Activity Unique ID

function SHEETNAME() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}


// Send Weekly Report

function consolidateWeeklyReports() {
  try {
    // Target Spreadsheet & Sheet
    var targetSS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1HtRFdjuOh7dRLTK_pKVeuBJ_xD8Qfc9ll6VRp3-X7GY/edit?usp=sharing");
    var targetSheet = targetSS.getSheetByName("Staff Weekly Report Submissions");

    // Source Spreadsheet
    var sourceSS = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1q-5mbgMjj7t0jrlQdOw8K8xMLfYpTikIt9ZtAk1Ce5g/edit?usp=sharing");

    // Array of Staff Names 
    var staffNames = [
      "Blaise", "Daniel", "Margo", "Amedee", "Irene", "Gaston", "Gisele", 
      "Augustin", "Charlse", "JPN", "Olivia"
    ]; 

    var staffWithSubmittedReports = []; 

    // Find the first empty row in the target sheet
    var targetLastRow = targetSheet.getLastRow();
    var firstEmptyRow = targetLastRow + 1; // Start populating from the next row

    for (var i = 0; i < staffNames.length; i++) {
      var staffName = staffNames[i];

      // Source Sheet
      var sourceSheet = sourceSS.getSheetByName(staffName + "_Weekly Report");
      if (!sourceSheet) {  
        Logger.log("Error: Sheet not found for " + staffName);
        continue; 
      }

      // Get Data Range (A3:N1000, data starts from row 3) 
      var dataRange = sourceSheet.getRange("A3:N1000"); 
      var data = dataRange.getValues();

      // Filter out empty rows (where column A is empty)
      var filteredData = data.filter(function(row) {
        return row[0] != ""; 
      });

      // Add Staff Name to Column F in filteredData
      for (var j = 0; j < filteredData.length; j++) {
        filteredData[j].splice(5, 0, staffName); 
      }

      // Append Filtered Data to Target Sheet, starting from the first empty row
      if (filteredData.length > 0) {
        targetSheet.getRange(firstEmptyRow, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
        staffWithSubmittedReports.push(staffName); 
        firstEmptyRow += filteredData.length; // Update firstEmptyRow for the next staff member
      } else {
        Logger.log("No data found for " + staffName); 
      }
    }

    // Log staff with submitted reports
    if (staffWithSubmittedReports.length > 0) {
      Logger.log("Reports submitted for: " + staffWithSubmittedReports.join(", "));
    } else {
      Logger.log("No reports were submitted.");
    }

    Logger.log("Consolidation completed!");

  } catch (error) {
    Logger.log("An error occurred during consolidation: " + error);
  }
}

// Time-Based Trigger 
function createTrigger() {
  // Delete existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "consolidateWeeklyReports") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create new trigger for every Thursday at 10 PM CAT
  ScriptApp.newTrigger("consolidateWeeklyReports")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.THURSDAY)
    .atHour(22) // 10 PM CAT is 22:00 in GMT
    .create();
}

