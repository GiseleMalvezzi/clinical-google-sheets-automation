/**
 * Clinical Report Automation for Google Sheets
 * Automates the generation of clinical reports from patient data
 * 
 * @author GiseleMalvezzi
 * @version 1.0
 */

/**
 * Function to create menu on spreadsheet open
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Clinical Reports')
    .addItem('Generate Report', 'generateClinicalReport')
    .addItem('Export to PDF', 'exportToPDF')
    .addItem('Clear Data', 'clearData')
    .addToUi();
}

/**
 * Main function to generate clinical report
 */
function generateClinicalReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  // Get data range
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  
  if (data.length <= 1) {
    ui.alert('No Data', 'Please import patient data first.', ui.ButtonSet.OK);
    return;
  }
  
  // Process each row of patient data
  var headers = data[0];
  var reportSheet = getOrCreateReportSheet(ss);
  
  // Clear previous reports
  reportSheet.clear();
  
  // Add headers to report
  reportSheet.appendRow(['Patient ID', 'Name', 'Report Date', 'Summary']);
  
  // Generate reports for each patient
  for (var i = 1; i < data.length; i++) {
    var patientData = data[i];
    var reportRow = processPatientData(patientData, headers);
    reportSheet.appendRow(reportRow);
  }
  
  // Format report
  formatReportSheet(reportSheet);
  
  ui.alert('Success', 'Clinical reports generated successfully!', ui.ButtonSet.OK);
}

/**
 * Process individual patient data
 */
function processPatientData(patientData, headers) {
  var patientId = patientData[0] || 'N/A';
  var patientName = patientData[1] || 'Unknown';
  var reportDate = new Date().toLocaleDateString();
  
  // Generate summary based on available data
  var summary = 'Clinical report for ' + patientName + ' - ';
  summary += 'Data collected on ' + reportDate;
  
  return [patientId, patientName, reportDate, summary];
}

/**
 * Get or create report sheet
 */
function getOrCreateReportSheet(ss) {
  var reportSheet = ss.getSheetByName('Clinical Reports');
  
  if (!reportSheet) {
    reportSheet = ss.insertSheet('Clinical Reports');
  }
  
  return reportSheet;
}

/**
 * Format the report sheet
 */
function formatReportSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastRow > 0) {
    // Format header row
    var headerRange = sheet.getRange(1, 1, 1, lastCol);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    // Auto-resize columns
    for (var i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
    }
  }
}

/**
 * Export report to PDF
 */
function exportToPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName('Clinical Reports');
  var ui = SpreadsheetApp.getUi();
  
  if (!reportSheet) {
    ui.alert('No Report', 'Please generate a clinical report first.', ui.ButtonSet.OK);
    return;
  }
  
  ui.alert('PDF Export', 'Use File > Download > PDF to export the report.', ui.ButtonSet.OK);
}

/**
 * Clear all data from active sheet
 */
function clearData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm', 'Are you sure you want to clear all data?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    ui.alert('Success', 'Data cleared successfully!', ui.ButtonSet.OK);
  }
}
