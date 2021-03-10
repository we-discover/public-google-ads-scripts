
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

  // Determine runtime environment
  var executionContext = 'client_account';
  if (typeof AdsManagerApp != "undefined") {
    executionContext = 'manager_account';
  }

  // If MCC, run process on a loop through all accounts
  if (executionContext === 'manager_account') {
    var managerAccount = AdsApp.currentAccount();
    var accountIterator = AdsManagerApp.accounts().get();
    while (accountIterator.hasNext()) {
      var account = accountIterator.next();
      AdsManagerApp.select(account);
      Logger.log('Info: Processing account ' + AdsApp.currentAccount().getName());
      testConfigurations = runExportsForAccount(testConfigurations, gsheetId);
    }
    AdsManagerApp.select(managerAccount);
  }

  // If client account, run on that account only
  if (executionContext === 'client_account') {
    testConfigurations = runExportsForAccount(testConfigurations, gsheetId);
  }

  // If enabled, send a summary email
  if (emailSettings.send_summary_email) {
      sendSummaryEmail(gsheetId, emailSettings.recipients, testConfigurations);
  }

}


function runExportsForAccount(testConfigurations, gsheetId) {

  // Loop through each test
  for (var i = 0; i < testConfigurations.length; i++) {

    // Start process
    var config = testConfigurations[i];
    var accountTestMessage = (
      ' test: ' + config.name +
      ' in account: ' + AdsApp.currentAccount().getName()
    );
    Logger.log('Info: Starting export for' + accountTestMessage);

    // Validate the test config
    try {
      validateConfiguration(config);
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      testConfigurations[i]['success'] = false;
      continue;
    }

    // Check if test exists in current account
    var testLabelIds = getTestLabelIds(config);
    if (testLabelIds.length < 1) {
      Logger.log('Info: No matching labels found in account');
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      testConfigurations[i]['success'] = false;
      continue;
    }

    try {
      // Create a query to pull test data
      var awqlQuery = buildQuery(config, testLabelIds);
      // Query data for each ad in the test and aggregate
      var aggTestData = queryAndAggregateData(config, awqlQuery);
      if (Object.keys(aggTestData).length === 0) {
        Logger.log('Error: No data observed for test: ' + config.name);
        Logger.log('Info: Skipping export for:' + accountTestMessage);
        testConfigurations[i]['success'] = false;
        continue;
      }
      // Format test data for export to GSheet
      var formattedTestData = formatTestDataForExport(config, aggTestData);
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log(awqlQuery);
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      testConfigurations[i]['success'] = false;
      continue;
    }

    // Export test data to Google Sheet
    try {
      exportDataToSheet(gsheetId, config, formattedTestData)
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      testConfigurations[i]['success'] = false;
      continue;
    }

    // Mark data export for test as success
    testConfigurations[i]['success'] = true;
  }

  return testConfigurations;
}
