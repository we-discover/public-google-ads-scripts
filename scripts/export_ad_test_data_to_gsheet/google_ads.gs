
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
      testConfigurations = extractDataForTestConfig(testConfigurations, gsheetId);
    }
    AdsManagerApp.select(managerAccount);
  }

  // If client account, run on that account only
  if (executionContext === 'client_account') {
    testConfigurations = extractDataForTestConfig(testConfigurations, gsheetId);
  }

  // Loop through each test and export extracted data
  for (var i = 0; i < testConfigurations.length; i++) {
    // Export test data to Google Sheet
    try {
      if (testConfigurations[i].update) {
        exportDataToSheet(gsheetId, testConfigurations[i])
      }
    } catch (anyErrors) {
      Logger.log(anyErrors);
      // Mark data export for test as failure
      testConfigurations[i]['success'] = false;
      continue;
    }
    // Mark data export for test as success
    testConfigurations[i]['success'] = true;
  }

  // Log test configurations with end state
  Logger.log(testConfigurations);
  
  // Reset Test Name, Variant 1 and Variant 2
  resetTestName(gsheetId)
  
  // run experiments data script
  getDailyExperimentData(gsheetId)
}

// Process that is run on each account
function extractDataForTestConfig(testConfigurations, gsheetId) {

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
      continue;
    }

    // Skip if not set to update
    if (!config.update) {
      Logger.log('Info: Update set to FALSE');
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      continue;
    }

    // Check if test exists in current account
    var testLabelIds = getTestLabelIds(config);
    if (testLabelIds.length < 1) {
      Logger.log('Info: No matching labels found in account');
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      continue;
    }

    try {
      // Create a query to pull test data
      var awqlQuery = buildQuery(config, testLabelIds);
      // Query data for each ad in the test and aggregate
      var aggTestData = queryAndAggregateData(config, awqlQuery);
      config['data'] = aggTestData;
      if (Object.keys(aggTestData).length === 0) {
        Logger.log('Error: No data observed for test: ' + accountTestMessage);
        continue;
      }
    } catch (anyErrors) {
      Logger.log(anyErrors);
      Logger.log(awqlQuery);
      Logger.log('Info: Skipping export for:' + accountTestMessage);
      continue;
    }
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
          config['data'] = {};
          config['success'] = false;
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
  var matches = labels.match('(' + mvtLabel + '\\$var_id:[0-9]{1,3})');
  if (matches) {
    var variantPart = matches[0].split('$')[1];
    return variantPart.replace('var_id:', '');
  }
}


// Runs AWQL query and aggregates data on a daily variant level
function queryAndAggregateData(config, awqlQuery) {
    var resultIterator = AdsApp.report(awqlQuery).rows();

    var dataObj = config['data'];

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

      dataObj[varId][date]['cost'] += Number(result["Cost"].replace(',', '')) || 0;
      dataObj[varId][date]['impressions'] += Number(result["Impressions"].replace(',', '')) || 0;
      dataObj[varId][date]['clicks'] += Number(result["Clicks"].replace(',', '')) || 0;
      dataObj[varId][date]['conversions'] += Number(result["Conversions"].replace(',', '')) || 0;
      dataObj[varId][date]['conversion_value'] += Number(result["ConversionValue"].replace(',', '')) || 0;
    }

    return dataObj;
}


// Converts aggregated test data into array based GSheet rows
function formatTestDataForExport(config) {
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

  var data = config['data'];
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
function exportDataToSheet(gsheetId, config) {

    var data = formatTestDataForExport(config);

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
    importSheet.hideSheet();
    Logger.log('Info: Sucessfully exported test data for test: ' + config.name);
}


// Resets 'Test Name' dropdown on 'Test Evaluation' sheet so data displayed upon entry
function resetTestName(gSheetId) {
    
  // Sheets
  var spreadsheet = SpreadsheetApp.openById(gSheetId);
  var testEvalSheet = spreadsheet.getSheetByName('Test Evaluation: Overview');
  var testDetailsSheet = spreadsheet.getSheetByName('Test details');
  
  // Ranges
  var mainControlsRange = testEvalSheet.getRange(3, 11, 2, 1)
  var abControlRange = testEvalSheet.getRange(40, 11, 1, 4)
  var detailsTestRange = testDetailsSheet.getRange(2, 2, 1, 1)
  var detailsVariantRange = testDetailsSheet.getRange(2, 4, 2, 1)
  
  // Cells
  var mainTestCell = mainControlsRange.getCell(1, 1)
  var abtestVariant1Cell = abControlRange.getCell(1, 1)
  var abtestVariant2Cell = abControlRange.getCell(1, 4)
  
  //Values
  var test1Cell = detailsTestRange.getCell(1, 1).getValue()
  var variant1Cell = detailsVariantRange.getCell(1, 1).getValue()
  var variant2Cell = detailsVariantRange.getCell(2, 1).getValue()
  
  // Set all to empty
  mainTestCell.setValue('')
  abtestVariant1Cell.setValue('')
  abtestVariant2Cell.setValue('')
  
  // Set all to first valid entries
  mainTestCell.setValue(test1Cell)
  abtestVariant1Cell.setValue(variant1Cell)
  abtestVariant2Cell.setValue(variant2Cell)
  
}
  
  
//--------------------------------------------------------------

function getDailyExperimentData(gsheetId) {
  
  // config variables from Experiment Details sheet
  try {
    var spreadsheet = SpreadsheetApp.openById(gsheetId);
    Logger.log('Info: Sucessfully connected to sheet');
  } catch (e) {
    throw Error('Connection to sheet failed')
  }
  var experimentSheet = spreadsheet.getSheetByName('Experiment Details');
  var experimentLastRow = experimentSheet.getLastRow();
  var experimentSearchRange = experimentSheet.getRange(2, 1, experimentLastRow, 4);
  
  // loop through rows in config sheet
  for (i = 1; i < experimentLastRow; i++) {
    
    // manager Accounts
    var managerAccount = AdsApp.currentAccount();
    var accountIterator = AdsManagerApp.accounts().get();

    // define variables for experiment
    var testName = experimentSearchRange.getCell(i, 1).getValue()
    var startDate = experimentSearchRange.getCell(i, 2).getValue()
    var endDate = experimentSearchRange.getCell(i, 3).getValue()
    var experimentUpdate = experimentSearchRange.getCell(i, 4).getValue()
    
    // create sheet for each experiment in MVT Testing gsheet
    var exportSheetName = "Data Import: " + testName;
    var exportSheet = spreadsheet.getSheetByName(exportSheetName);
    if (exportSheet === null) {
      exportSheet = spreadsheet.insertSheet(exportSheetName, 99);
    }
    exportSheet.clear();
    Logger.log('Info: Sucessfully loaded data import sheet for test: ' + testName);
    
    // create export array for sending data to gsheet
    var outputEntities = [[
        'account_name',
        'currency',
        'test_name',
        'experiment_label',
        'variant_id',
        'date',
        'cost',
        'impressions',
        'clicks',
        'conversions',
        'conversion_value'
      ]]
    
    // create data object for aggregated test data
    dataObj = {}
    dataObj['data'] = {}
    // add testName layer to dataObj
    dataObj['data'][testName] = {}
    // add control and variant layers to dataObj
    dataObj['data'][testName]['control'] = {};
    dataObj['data'][testName]['variant'] = {};
    Logger.log(dataObj)
    
    // Iterate through the list of accounts
    while (accountIterator.hasNext()) {
      var account = accountIterator.next();

      // Select the client account and get currency and timezone
      AdsManagerApp.select(account);
      var accountName = account.getName()
      var accountCurrency = account.getCurrencyCode()
      var timeZone = account.getTimeZone()
      
      // while start date is less than or equal to end date
      while (startDate <= endDate) {
        
        Logger.log(Utilities.formatDate(startDate, timeZone, "dd/MM/yyyy"))
        // Select campaigns under the client account
        var campaignSelector = AdsApp.campaigns()
        var campaignIterator = campaignSelector.get()
        
        
        // iterate through campaigns
        while (campaignIterator.hasNext()) {

          var campaign = campaignIterator.next();
          var campaignString = campaign.toString()
          var campaignName = campaign.getName()

          
          // check if campaign is Experiment, same as the config sheet, and update set to True
          if (campaign.isExperimentCampaign() && campaignString.indexOf(testName) != -1 && experimentUpdate) {
            
            // get experiment and base campaign names, define date as date string
            var expCampaign = campaign
            var baseCampaign = campaign.getBaseCampaign()
            var date = Utilities.formatDate(startDate, timeZone, "dd/MM/yyyy")
            
            // create empty dicts for base campaigns
            dataObj['data'][testName]['control'][date] = {
              'account_name': accountName,
              'currency': accountCurrency,
              'cost': 0,
              'impressions': 0,
              'clicks': 0,
              'conversions': 0,
              'conversion_value': 0
            }
            
            // get and add data for base campaign
            var baseStats = baseCampaign.getStatsFor(Utilities.formatDate(startDate, timeZone, "yyyyMMdd"),
                                                     Utilities.formatDate(startDate, timeZone, "yyyyMMdd"))
            var cost = baseStats.getCost()
            var impressions = baseStats.getImpressions()
            var clicks = baseStats.getClicks()
            var conversions = baseStats.getConversions()

            dataObj['data'][testName]['control'][date]['cost'] += cost;
            dataObj['data'][testName]['control'][date]['impressions'] += impressions;
            dataObj['data'][testName]['control'][date]['clicks'] += clicks;
            dataObj['data'][testName]['control'][date]['conversions'] += conversions;
            dataObj['data'][testName]['control'][date]['conversion_value'] += 0;
            
            
            // create empty dicts for experiment campaigns
            dataObj['data'][testName]['variant'][date] = {
              'account_name': accountName,
              'currency': accountCurrency,
              'cost': 0,
              'impressions': 0,
              'clicks': 0,
              'conversions': 0,
              'conversion_value': 0
            }
            
            // get data for experiment campaign
            var expStats = expCampaign.getStatsFor(Utilities.formatDate(startDate, timeZone, "yyyyMMdd"),
                                                   Utilities.formatDate(startDate, timeZone, "yyyyMMdd"))
            var cost = expStats.getCost()
            var impressions = expStats.getImpressions()
            var clicks = expStats.getClicks()
            var conversions = expStats.getConversions()

            dataObj['data'][testName]['variant'][date]['cost'] += cost;
            dataObj['data'][testName]['variant'][date]['impressions'] += impressions;
            dataObj['data'][testName]['variant'][date]['clicks'] += clicks;
            dataObj['data'][testName]['variant'][date]['conversions'] += conversions;
            dataObj['data'][testName]['variant'][date]['conversion_value'] += 0;
          }
        }
        
        // increment date by one day
        startDate.setDate(startDate.getDate() + 1)
        
      }
    }
    
    Logger.log(dataObj)
    
    for (var testName in dataObj['data']) {
      for (var date in dataObj['data'][testName]['control']) {
        Logger.log(date)
        outputEntities.push([
          dataObj['data'][testName]['control'][date]['account_name'],
          dataObj['data'][testName]['control'][date]['currency'],
          testName,
          'control',
          1,
          date,
          dataObj['data'][testName]['control'][date]['cost'],
          dataObj['data'][testName]['control'][date]['impressions'],
          dataObj['data'][testName]['control'][date]['clicks'],
          dataObj['data'][testName]['control'][date]['conversions'],
          dataObj['data'][testName]['control'][date]['conversion_value']
        ])
      }

      for (var date in dataObj['data'][testName]['variant']) {
        outputEntities.push([
          dataObj['data'][testName]['control'][date]['account_name'],
          dataObj['data'][testName]['control'][date]['currency'],
          testName,
          'variant',
          2,
          date,
          dataObj['data'][testName]['variant'][date]['cost'],
          dataObj['data'][testName]['variant'][date]['impressions'],
          dataObj['data'][testName]['variant'][date]['clicks'],
          dataObj['data'][testName]['variant'][date]['conversions'],
          dataObj['data'][testName]['variant'][date]['conversion_value']
        ])
      }
    }

    var exportRange = exportSheet.getRange(1, 1, outputEntities.length, outputEntities[0].length);
    exportRange.setValues(outputEntities);
    exportSheet.hideSheet();
    Logger.log('Info: Sucessfully exported test data for test: ' + testName);
   
  }
}
