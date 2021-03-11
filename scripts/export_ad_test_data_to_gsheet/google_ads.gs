
// Entrypoint for script
function main() {

  // EDIT ME -- Google Sheet ID for Template
  const gsheetId = 'XXXXXXX';

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

  // Log test configurations with end state
  Logger.log(testConfigurations);

}

// Process that is run on each account
function runExportsForAccount(testConfigurations, gsheetId) {

  // Loop through each test
  for (var i = 0; i < testConfigurations.length; i++) {

    // Load test configuration
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

    // Skip if not set to update
    if (!config.update) {
      Logger.log('Info: Update set to FALSE');
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


// ========= UTILITY FUNCTIONS ==============================================================


// Function to load test configurations from GSheet
function loadTestConfigsFromSheet(gsheetId) {
    var testConfigurations = [];

    try {
      var spreadsheet = SpreadsheetApp.openById(gsheetId);
      Logger.log('Info: Sucessfully connected to gsheet.');
    } catch (e) {
      throw Error('Failed to connect to gsheet.')
    }

    try {
      var testConfigSheet = spreadsheet.getSheetByName('Google Ads - Test Details');

      var [rows, columns] = [testConfigSheet.getLastRow(), testConfigSheet.getLastColumn()];
      var data = testConfigSheet.getRange(1, 1, rows, columns).getValues();
      const header = data[0];

      data.shift();
      data.map(function(row) {
        var empty = row[0] === '';
        if (!empty) {
          var config = header.reduce(function(o, h, i) {
            o[h] = row[i];
            return o;
          }, {});
          testConfigurations.push(config);
        }
      });
    } catch (e) {
      throw Error('Failed to load test configurations from gsheet.')
    }

    return testConfigurations;
}


// Runs some basic alidation on a single test config
function validateConfiguration(config) {

  function assert(check, condition) {
    if (!condition) {
      throw new Error('Validation failed: ' + check);
    }
  }

  assert("start date formatted correctly", config.start_date.match('[0-9]{8}'));
  assert("end date formatted correctly",  config.end_date.match('[0-9]{8}'));
  assert("start date before end date", Number(config.start_date) < Number(config.end_date));
}


// Query the current account to get label IDs
function getTestLabelIds(config) {
    var labelIds = [];
    var labelIterator = AdsApp.labels()
      .withCondition("Name CONTAINS '" + config.mvt_label + "'")
      .get();
    while (labelIterator.hasNext()) {
      labelIds.push(labelIterator.next().getId());
    }
    return labelIds;
}


// Builds a query to pull raw data for a single test
function buildQuery(config, labelIds) {
    return (" \
      SELECT \
          CustomerDescriptiveName \
        , Labels \
        , Date \
        , Cost \
        , Impressions \
        , Clicks \
        , Conversions \
        , ConversionValue \
      FROM \
        AD_PERFORMANCE_REPORT \
      WHERE \
        Labels CONTAINS_ANY [" + labelIds.join(',') + "] \
        AND Impressions > 0 \
      DURING " +
        config.start_date + "," + config.end_date
    ).replace(/ +(?= )/g, '');
}


// Extracts variant_id from labels
function extractVariantIdFromLabels(config, labels) {
  var mvtLabel = config.mvt_label;
  var matches = labels.match('(' + mvtLabel + ';var_id:[0-9]{1,3})');
  if (matches) {
    var variantPart = matches[0].split(';')[1];
    return variantPart.replace('var_id:', '');
  }
}


// Runs AWQL query and aggregates data on a daily variant level
function queryAndAggregateData(config, awqlQuery) {
    var resultIterator = AdsApp.report(awqlQuery).rows();

    var dataObj = {};

    while (resultIterator.hasNext()) {
      var result = resultIterator.next();

      var varId = extractVariantIdFromLabels(config, result["Labels"]);
      var date = result["Date"];

      if (!dataObj.hasOwnProperty(varId)) {
        dataObj[varId] = {};
      }

      if (!dataObj[varId].hasOwnProperty(date)) {
        dataObj[varId][date] = {
          'cost': 0,
          'impressions': 0,
          'clicks': 0,
          'conversions': 0,
          'conversion_value': 0
        };
      }

      dataObj[varId][date]['cost'] += Number(result["Cost"]) || 0;
      dataObj[varId][date]['impressions'] += Number(result["Impressions"]) || 0;
      dataObj[varId][date]['clicks'] += Number(result["Clicks"]) || 0;
      dataObj[varId][date]['conversions'] += Number(result["Conversions"]) || 0;
      dataObj[varId][date]['conversion_value'] += Number(result["ConversionValue"]) || 0;
    }

    return dataObj;
}


// Converts aggregated test data into array based GSheet rows
function formatTestDataForExport(config, data) {
  var output = [[
    'account_name',
    'currency',
    'test_name',
    'mvt_label',
    'variant_id',
    'date',
    'cost',
    'impressions',
    'clicks',
    'conversions',
    'conversion_value'
  ]];

  var accountName = AdsApp.currentAccount().getName();
  var currency = AdsApp.currentAccount().getCurrencyCode();

  for (var variantId in data) {
    for (var date in data[variantId]) {
      output.push([
        accountName,
        currency,
        config.name,
        config.mvt_label,
        variantId,
        date,
        data[variantId][date]["cost"],
        data[variantId][date]["impressions"],
        data[variantId][date]["clicks"],
        data[variantId][date]["conversions"],
        data[variantId][date]["conversion_value"]
      ]);
    }
  }

  return output;
}


// Connects to a Google Sheet and writes data for a single test
function exportDataToSheet(gsheetId, config, data) {
    try {
      var spreadsheet = SpreadsheetApp.openById(gsheetId);
      Logger.log('Info: Sucessfully connected to sheet for test: ' + config.name);
    } catch (e) {
      throw Error('Connection to sheet failed for test: ' + config.name)
    }

    var importSheetName = "Data Import: " + config.name;
    var importSheet = spreadsheet.getSheetByName(importSheetName);
    if (importSheet === null) {
      importSheet = spreadsheet.insertSheet(importSheetName, 99);
    }
    importSheet.clear();
    Logger.log('Info: Sucessfully loaded data import sheet for test: ' + config.name);

    var importRange = importSheet.getRange(1, 1, data.length, data[0].length);
    importRange.setValues(data);
    Logger.log('Info: Sucessfully exported test data for test: ' + config.name);
}
