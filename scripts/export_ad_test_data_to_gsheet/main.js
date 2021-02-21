
function main() {

  // EDIT ME -- Google Sheet ID for Template
  const gsheetId = 'XXXXXXXXX__XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX';

  // EDIT ME -- General settings for the script
  const emailSettings = {
    send_summary_email: true,
    recipients: [
      'someone@domain.com',
      'someonelse@domain.com'
    ]
  };

  // Read all test configurations from GSheet
  const testConfigurations = loadTestConfigsFromSheet(gsheetId);

  // Loop through each test
  for (var i = 0; i < testConfigurations.length; i++) {

    // Start process
    var config = testConfigurations[i];
    Logger.log('Info: Starting export for test: ' + config.name);

    // Validate the test config
    try {
      validateConfiguration(config);
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log('Info: Skipping export for test: ' + config.name);
      testConfigurations[i]['success'] = false;
      continue;
    }

    // Create a query to pull test data
    var awqlQuery = buildQuery(config);

    // Query data for each ad in the test and aggregate
    var aggTestData = queryAndAggregateData(config, awqlQuery);
    if (Object.keys(aggTestData).length === 0) {
      Logger.log('Error: No data observed for test: ' + config.name);
      Logger.log('Info: Skipping export for test: ' + config.name);
      testConfigurations[i]['success'] = false;
      continue;
    }

    // Format test data for export to GSheet
    var formattedTestData = formatTestDataForExport(config, aggTestData);

    // Export test data to Google Sheet
    try {
      exportDataToSheet(gsheetId, config, formattedTestData)
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log('Info: Skipping export for test: ' + config.name);
      testConfigurations[i]['success'] = false;
      continue;
    }

    // Mark data export for test as success
    testConfigurations[i]['success'] = true;
  }

  // If enabled, send a summary email
  if (emailSettings.send_summary_email) {
      sendSummaryEmail(gsheetId, emailSettings.recipients, testConfigurations);
  }
}
