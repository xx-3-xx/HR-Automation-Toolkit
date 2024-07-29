// Code.gs

function code(){
  initialize();
  loadData();
}

/**
 * Initialize global settings and configurations
 */
function initialize() {
  // Set up global constants or variables
  const TEMPLATE_ID = '1HD5J9iQTiPvHgDfegs-1dG8acxA1AcC8knNh9dcYd3E'; // google form id as template id
  PropertiesService.getScriptProperties().setProperty('1HD5J9iQTiPvHgDfegs-1dG8acxA1AcC8knNh9dcYd3E', TEMPLATE_ID);
  
  Logger.log('Initialization complete');
}


function loadData() {
  const spreadsheetId = '1_ZLCXK1IRc37XVveeIR4pLPInNLBucXe_fkK4l_r8CA'; //spreadsheet ID
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  
  const formResponsesSheet = spreadsheet.getSheetByName('Job Application Form Responses');
  const interviewScheduleSheet = spreadsheet.getSheetByName('Interview Schedule');
  
  if (!formResponsesSheet || !interviewScheduleSheet) {
    Logger.log('One or more sheets are missing.');
    return;
  }
  
  const formResponsesData = formResponsesSheet.getDataRange().getValues();
  const interviewScheduleData = interviewScheduleSheet.getDataRange().getValues();
  
  Logger.log('Form Responses Data Length: ' + formResponsesData.length);
  Logger.log('Interview Schedule Data Length: ' + interviewScheduleData.length);
  
  // Log actual data
  Logger.log('Form Responses Data: ' + JSON.stringify(formResponsesData));
  Logger.log('Interview Schedule Data: ' + JSON.stringify(interviewScheduleData));
}