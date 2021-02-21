
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

// Basic implementation of assertion
function assert(check, condition) {
  if (!condition) {
    throw new Error('Validation failed: ' + check);
  }
}


// Runs some basic alidation on a single test config
function validateConfiguration(config) {
  assert("start date formatted correctly", config.start_date.match('[0-9]{8}'));
  assert("end date formatted correctly",  config.end_date.match('[0-9]{8}'));
  assert("start date before end date", Number(config.start_date) < Number(config.end_date));
}


// Builds a query to pull raw data for a single test
function buildQuery(config) {
    var labelIds = [];
    var labelIterator = AdsApp.labels()
      .withCondition("Name CONTAINS '" + config.mvt_label + "'")
      .get();
    while (labelIterator.hasNext()) {
      labelIds.push(labelIterator.next().getId());
    }

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

  const accountName = AdsApp.currentAccount().getName();
  const currency = AdsApp.currentAccount().getCurrencyCode();

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


// Function to send a summary email on run
function sendSummaryEmail(gsheetId, recipients, configs) {
  const today = new Date();
  var timeZone = AdsApp.currentAccount().getTimeZone();
  const sendTime = Utilities.formatDate(today, timeZone, "yyyy-MM-dd HH:mm");

  MailApp.sendEmail(
    recipients.join(','),
    'Data export summary: ' + sendTime,
    '',
    {
      name: 'Google Ads MVT Data Exporter',
      noReply: true,
      htmlBody: generateEmailBody(gsheetId, configs),
    }
  );
}


// Function to generate the body of the summary email
function generateEmailBody(gsheetId, configs) {
  var failedItems = [];
  var succesfulItems = [];
  for (var i = 0; i < configs.length; i++) {
    if (!configs[i].success) {
      failedItems.push('<li>' + configs[i].name + '</li>');
      continue;
    }
    var testGsheet = DriveApp.getFileById(gsheetId);
    succesfulItems.push(
      '<a href="' + testGsheet.getUrl() + '"><li>' + configs[i].name + '</li></a>'
    );
  }

  var emailBody = "<h1>MVT Monitor Summary</h1>";

  if (succesfulItems.length > 0) {
    emailBody += (
      '<h2>Successful exports: ' +  succesfulItems.length + '/' + configs.length + '</h2>' +
      '<ul>' + succesfulItems.join('') + '</ul>'
    );
  }

  if (failedItems.length > 0) {
    emailBody += (
      '<h2>Successful exports: ' +  failedItems.length + '/' + configs.length + '</h2>' +
      '<ul>' + failedItems.join('') + '</ul>'
    );
  }

  const account = AdsApp.currentAccount();
  emailBody += (
    '<p style="margin-top:50px;">You are receving this receiving this email from a script enabled on  ' +
    'the following Google Ads account: ' + account.getName() + ' (' + account.getCustomerId() + ').</p>' +
    '<p><i>Orginally created by WeDiscover Digital LTD<i></p>'
  )

  return emailBody;
}
