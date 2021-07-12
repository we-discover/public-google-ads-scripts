// *** Globals *** //

  // Email Notification Sheet
  var emailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Notifications');
  var emailLastRow = emailSheet.getLastRow();
  var emailLastCol = emailSheet.getLastColumn();
  var emailSearchRange = emailSheet.getRange(2, 1, emailLastRow, emailLastCol);

  // Test Evaluation Sheet
  var testevalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview');
  var controlsSearchRange = testevalSheet.getRange(4, 10, 2, 6);
  var testCell = controlsSearchRange.getCell(1, 2);
  var metricCell = controlsSearchRange.getCell(2, 2);
  var confidenceCell = controlsSearchRange.getCell(1, 6).getValue();
  var powerCell = controlsSearchRange.getCell(2, 6).getValue();

  // Test Type
  var testTypeRange = testevalSheet.getRange(3, 11, 1, 1);
  var testTypeCell = testTypeRange.getCell(1, 1)
  
  // MVT Test Results
  var mvtSearchRange = testevalSheet.getRange(31, 10, 6, 3)
  var mvtLastRow = 6
  var mvtSent = false


// *** Functions *** //

function sendEmailNotifications() {

  /**
   * This function sends emails to all users listed in the "Email Notifications" tab,
   * in relation to the respective "Test Name"/"Metric Name" combinations. Once an email has beeen
   * sent, indicating a significant result in relation to a specfic test, further emails will no
   * longer be sent until the user resets the email notification functionality.
   */

  // loop through all rows in the Email Notifications tab
  for (i = 1; i < emailLastRow; i++) {

    // set email sent status
    mvtSent = false

    // get respective values from Email Notifications sheet
    var emailTestCell = emailSearchRange.getCell(i, 1).getValue()
    var emailMetricCell = emailSearchRange.getCell(i, 2).getValue()
    var emailUserCell = emailSearchRange.getCell(i, 3).getValue()
    var emailAddressCell = emailSearchRange.getCell(i, 4).getValue()
    var emailSentCell = emailSearchRange.getCell(i, 5).getValue()
    var emailTestTypeCell = emailSearchRange.getCell(i, 6).getValue()

    // set dropdown menu values for "Test Name" and "Objective Metric" in main Controls
    testTypeCell.setValue(emailTestTypeCell)
    Logger.log(emailTestTypeCell);
    testCell.setValue(emailTestCell)
    Logger.log(emailTestCell);
    metricCell.setValue(emailMetricCell)
    Logger.log(emailMetricCell);

    // loop through rows in the MVT "Statistical Test Outcome" section
    for (j = 1; j < mvtLastRow; j++) {

        // get mvt test result
        var mvtVariantCell = mvtSearchRange.getCell(j, 1).getValue();
        var mvtTestCell = mvtSearchRange.getCell(j, 2).getValue();
        var mvtPowerCell = mvtSearchRange.getCell(j, 3).getValue();

        

        // check for significant result
        if (mvtTestCell.includes("significance") && mvtPowerCell >= powerCell && emailSentCell == false) {

            // set email variables
            var emailSubject = "WeDiscover Experimentation - " + emailTestTypeCell + " - Significant Result"
            var emailMessage = ""

            // template for email body
            var emailBodyTemplate = "<div style=\"font-size:16px\">\
	<h1>Significant result achieved!</h1>\
	<p>\
		Hi [name],\
		<br><br>\
		This is an automated email to let you know that your <b>[test type] -</b> <b>[test name]</b> has returned a statistically significant result:\
		<br><br>\
		The <b>[variant name]</b> variant has performed [better/worse]\
		<br><br>\
		This result has been obtained with the test evaluation conditions below:\
	</p>\
	<ul>\
		<li>Objective metric: X</li>\
		<li>Confidence: >Y%</li>\
		<li>Power: Z%</li>	\
	</ul>\
	<p>\
		Happy experimenting!\
		<br><br>\
		The WeDiscover team\
	</p>\
	<hr>\
	<p style=\"font-size:12px\"><em>\
		Find more details in the <a href=\"[spreadsheet_link]\">test evaluation sheet</a>.\
		<br><br>\
		This tool was made open source by WeDiscover. For further details on sheet set up, please see <a href=\"[link]\">this article</a>. \
		<br>\
		For any questions relating to the script or test evaulation, please email:                 \
		<a\
        	href=\"mailto:scripts@we-discover.com?subject=MVT Questions\"\
            target=\"_blank\"\
            rel=\"noopener noreferrer\"\
        >\
    		scripts@we-discover.com\
    	</a>\
    	<br><br>\
    	If you are not expecting to receive this email, please contact the owner of the sheet listed above.\
	</em></p>\
	<hr>\
</div>'"
            // replace variables in html email body above
            emailBodyTemplate = emailBodyTemplate.replace("[name]", emailUserCell)
            emailBodyTemplate = emailBodyTemplate.replace("[test name]", emailTestCell)
            emailBodyTemplate = emailBodyTemplate.replace("[test type]", emailTestTypeCell)
            emailBodyTemplate = emailBodyTemplate.replace("[variant name]", mvtVariantCell)
            emailBodyTemplate = emailBodyTemplate.replace("X", emailMetricCell)
            emailBodyTemplate = emailBodyTemplate.replace("Y", (confidenceCell*100).toFixed(0))
            emailBodyTemplate = emailBodyTemplate.replace("Z", (mvtPowerCell*100).toFixed(0))
            emailBodyTemplate = emailBodyTemplate.replace("[better/worse]", mvtTestCell.toLowerCase())

            // construct link to current sheet and insert into email
            var SS = SpreadsheetApp.getActiveSpreadsheet();
            var ss = SS.getActiveSheet();
            var url = '';
            url += SS.getUrl();
            url += '#gid=';
            url += ss.getSheetId();
            emailBodyTemplate = emailBodyTemplate.replace("[spreadsheet_link]", url)

            // send email if email address present
            if (emailAddressCell.length > 0) {
              MailApp.sendEmail(emailAddressCell, emailSubject, emailMessage, {
                htmlBody: emailBodyTemplate
              })
              mvtSent = true
            }
            else {
              Logger.log("Please check the configuraton of your email addresses.")
            }
        }
    }

  // change the Sent colum to True if an email has been sent
  if (mvtSent == true) {

    Logger.log("Email sent successfully!")
    emailSearchRange.getCell(i, 5).setValue(true)
  }
  } 
}

function resetEmailNotifications() {

  /**  
   * This function resets the email notification functionality. 
   * All rows in the "Sent" column will be reset to TRUE if currently FALSE
   */ 

  // loop through all rows in "Sent" column and reset to FALSE
  for (i = 1; i < emailLastRow; i++){

    var cell = emailSearchRange.getCell(i, 5).getValue();

    if (cell = true){

      emailSearchRange.getCell(i, 5).setValue(false)
    }
  }
}

function onEdit(e) {

  /**  
   * This function executes only when the "Test Name" control on the mail "Test Evaluation" sheet is edited
   * If this cells is edited, the script will amend the Variants in the A/B Test control section to the first two valid Variant names.
   * If this cell is NOT edited, no changes will take place. Finally all cells in the sheet are re-evaluated.
   */ 

  // Set edit variables
  var range = e.range;
  var spreadSheet = e.source;
  var sheetName = spreadSheet.getActiveSheet().getName();
  var column = range.getColumn();
  var row = range.getRow();

  // Check if the cell has been edited
  if(sheetName == 'Test Evaluation: Overview' && column == 11 && row == 3) {

    // set variables for logic
    var testTypeCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview').getRange("K3")
    var testTypeValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview').getRange("K3").getValue();
    var testInputCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview').getRange("K4")
    var testInputValue = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview').getRange("K4").getValue();
    var aInputCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview').getRange("K41");
    var bInputCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test Evaluation: Overview').getRange("N41");

    if (testTypeValue == 'Ads Test') {
      var testDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ad Test details');
    }
    else if (testTypeValue == 'D&E Test') {
      var testDetailsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('D&E Test details');
    }

    var testDetailsLastRow = testDetailsSheet.getLastRow();
    var testDetailsRange = testDetailsSheet.getRange(2, 2, testDetailsLastRow - 1, 3);

    // Loop through all test names in "Details" sheet
    for (i = 1; i < testDetailsLastRow; i++) {

      testNameValue = testDetailsRange.getCell(i, 1).getValue()

      // If chosen test name equal to test name in "Details" sheet
      if (testTypeValue == 'Ads Test') {
        
        // Set Test Input Cell
        testInputCell.setValue(testNameValue)
        // Set Variants 1 and 2 to valid entries from "Details" sheet
        aInputCell.setValue(testDetailsRange.getCell(i, 3).getValue())
        bInputCell.setValue(testDetailsRange.getCell(i + 1, 3).getValue())

        break
      }

      if (testTypeValue == 'D&E Test') {
        
        // Set Test Input Cell
        testInputCell.setValue(testNameValue)
        // Set Variants 1 and 2 to valid entries from "Details" sheet
        aInputCell.setValue('Control')
        bInputCell.setValue('Treatment')

        break
      }
    
  SpreadsheetApp.flush();

    }
  }
}
