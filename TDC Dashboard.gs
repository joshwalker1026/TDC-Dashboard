var ss = SpreadsheetApp.getActiveSpreadsheet();

var weeklySheet;
var mailingSheet;
var scheduleSheet;
var dashboardSheet;

var releaseSheet;
var bugzillaURL = "https://bugzilla.mozilla.org";
var currentVersion;
var releaseDate;

var countP1 = 0;
var countP2 = 0;
var countP3 = 0;
var countP4 = 0;
var countP5 = 0;
var countPN = 0;

var FFversion = ['52', '51'];

var d = new Date();

function onOpen() {
  ss = SpreadsheetApp.getActiveSpreadsheet();

  
  
  var menuItems = [
    {name: 'Update', functionName: 'overallCountBug'}
  ];
  ss.addMenu('[Caculate Bugs]', menuItems);
}


function overallCountBug() {
  
  var startRow;
  var RESTQuery;
  var link;
  var response;
  var totalBugs;
  var firefoxBugs;
  var platformBugs;
  var startColumn = 3;
  
  weeklySheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[1]);
  mailingSheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[2]);
  scheduleSheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[3]);
  dashboardSheet = SpreadsheetApp.setActiveSheet(ss.getSheets()[0]);
  
  currentVersion = dashboardSheet.getRange('B3').getValue().substring(0,2);
  releaseDate = scheduleSheet.getRange('B6').getValue();
  
  for (var j=0; j < FFversion.length; j++){
    startRow=j+1; 
    countP1 = 0;
    countP2 = 0;
    countP3 = 0;
    countP4 = 0;
    countP5 = 0;
    countPN = 0;

    
    // Query all bugs
    RESTQuery = bugzillaURL + "/rest/bug?include_fields=id,priority&bug_status=RESOLVED&f1=cf_status_firefox" 
    + FFversion[j]
    + "&o1=equals&resolution=FIXED&v1=fixed";
    
    link = bugzillaURL + "/buglist.cgi?o1=equals&v1=fixed&f1=cf_status_firefox" 
    + FFversion[j]
    + "&resolution=FIXED&query_format=advanced&bug_status=RESOLVED";
   
    totalBugs = sendRequest(RESTQuery);
    prioritizeBugs (totalBugs);
    dashboardSheet.getRange(startRow+2,  startColumn).setFormula("=hyperlink(\"" + link + "\";\"" + totalBugs.bugs.length + "\")");
   
    dashboardSheet.getRange(startRow+2,  startColumn+6).setValue(countP1);
    dashboardSheet.getRange(startRow+2,  startColumn+7).setValue(countP2);
    dashboardSheet.getRange(startRow+2,  startColumn+8).setValue(countP3);
    dashboardSheet.getRange(startRow+2,  startColumn+9).setValue(countP4);
    dashboardSheet.getRange(startRow+2,  startColumn+10).setValue(countP5);
    dashboardSheet.getRange(startRow+2,  startColumn+11).setValue(countPN);
    
    countP1 = 0;
    countP2 = 0;
    countP3 = 0;
    countP4 = 0;
    countP5 = 0;
    countPN = 0;
    
    // Query TDC Firefox and platform bugs
    RESTQuery = bugzillaURL + "/rest/bug?include_fields=id,priority&bug_status=RESOLVED&f1=cf_status_firefox"
    +FFversion[j]
    +"&f2=assigned_to&o1=equals&o2=anywordssubstr&resolution=FIXED&v1=fixed&v2=gasolin%20timdream%20tchien%20rchien%20schung%20ralin%20flin%20etseng%20scwwu%20ehung%20lchang%20dhuang%20kmlee%20selee%20yliao%20fliu%20jyeh%20vchen%20hhsu%20tchen%20pchen%20mochen%20thsieh%20hhuang%20jhuang%20chuang%20mliang%20jalin%20bmao%20fshih%20atsai%20gchang%20whsu%20ashiue%20ctang%20wiwang%20ywu%20brsun%20echuang%20alchen%20lochang";
    
    link = bugzillaURL + "/buglist.cgi?f1=cf_status_firefox"
    +FFversion[j]
    +"&o1=equals&resolution=FIXED&o2=anywordssubstr&query_format=advanced&f2=assigned_to&bug_status=RESOLVED&v1=fixed&v2=gasolin%2C%20timdream%2C%20tchien%2C%20rchien%2C%20schung%2C%20ralin%2C%20flin%2C%20etseng%2C%20scwwu%2C%20ehung%2C%20lchang%2C%20dhuang%2C%20kmlee%2C%20selee%2C%20yliao%2C%20fliu%2C%20jyeh%2C%20vchen%2C%20hhsu%2C%20tchen%2C%20pchen%2C%20mochen%2C%20thsieh%2C%20hhuang%2C%20jhuang%2C%20chuang%2C%20mliang%2C%20jalin%2C%20bmao%2C%20fshih%2C%20atsai%2C%20gchang%2C%20whsu%2C%20ashiue%2C%20ctang%2C%20wiwang%2C%20ywu%2C%20brsun%2C%20echuang%2C%20alchen%2C%20lochang";
       
    firefoxBugs = sendRequest(RESTQuery);
    prioritizeBugs (firefoxBugs);
    dashboardSheet.getRange(startRow+2,  startColumn+2).setFormula("=hyperlink(\"" + link + "\";\"" + firefoxBugs.bugs.length + "\")");
    
    
    // Query TDC Platform bugs 
    RESTQuery = bugzillaURL + "/rest/bug?include_fields=id,priority&bug_status=RESOLVED&f1=cf_status_firefox"
    +FFversion[j]
    +"&f2=assigned_to&o1=equals&o2=anywordssubstr&resolution=FIXED&v1=fixed&v2=shuang%20ttung%20joliu%20tlee%20kchen%20tchou%20cyu%20wpan%20cku%20fatseng%20pchang%20boris.chiou%20tlin%20hshih%20mtseng%20vliu%20ethlin%20aschen%20echen%20btseng%20jhao%20jjong%20htsai%20bhsu%20jdai%20sawang%20sshih%20hchang%20allstars.chh%20dlee%20ettseng%20tnguyen%20tihuang%20gweng%20bechen%20jolin%20ctai%20jwwang%20ayang%20alwu%20tkuo%20mchiang%20bwu%20tkuo%20howareyou322%20kaku%2C%20xeonchen%2C%20amchung%2C%20juhsu%2C%20swu%2C%20kechen%2C%20jacheng%2C%20kuoe0%2C%20dmu%2C%20cleu%2C%20etsai%2C%20kikuo%2C%20kechang%2C%20cchang%2C%20schien";
        
    link = bugzillaURL + "/buglist.cgi?f1=cf_status_firefox"
    +FFversion[j]
    +"&o1=equals&resolution=FIXED&o2=anywordssubstr&query_format=advanced&f2=assigned_to&bug_status=RESOLVED&v1=fixed&v2=shuang%2C%20ttung%2C%20joliu%2C%20tlee%2C%20kchen%2C%20tchou%2C%20cyu%2C%20%20wpan%2C%20cku%2C%20%20fatseng%2C%20%20pchang%2C%20%20boris.chiou%2C%20tlin%2C%20%20hshih%2C%20%20mtseng%2C%20%20%20vliu%2C%20%20ethlin%2C%20%20aschen%2C%20%20echen%2C%20%20btseng%2C%20jhao%2C%20%20jjong%2C%20%20htsai%2C%20%20bhsu%2C%20%20jdai%2C%20%20sawang%2C%20%20sshih%2C%20%20hchang%2C%20%20allstars.chh%2C%20%20dlee%2C%20%20ettseng%2C%20%20tnguyen%2C%20%20tihuang%2C%20%20gweng%2C%20%20bechen%2C%20%20jolin%2C%20%20ctai%2C%20%20jwwang%2C%20%20ayang%2C%20%20alwu%2C%20%20tkuo%2C%20%20mchiang%2C%20bwu%2C%20tkuo%2C%20howareyou322%2C%20kaku%2C%20xeonchen%2C%20amchung%2C%20juhsu%2C%20swu%2C%20kechen%2C%20jacheng%2C%20kuoe0%2C%20dmu%2C%20cleu%2C%20etsai%2C%20kikuo%2C%20kechang%2C%20cchang%2C%20schien"
    
    platformBugs = sendRequest(RESTQuery);
    prioritizeBugs (platformBugs);
    
    dashboardSheet.getRange(startRow+2,  startColumn+3).setFormula("=hyperlink(\""+link+"\";\"" + platformBugs.bugs.length + "\")");
    dashboardSheet.getRange(startRow+2,  startColumn+12).setValue(countP1);
    dashboardSheet.getRange(startRow+2,  startColumn+13).setValue(countP2);
    dashboardSheet.getRange(startRow+2,  startColumn+14).setValue(countP3);
    dashboardSheet.getRange(startRow+2,  startColumn+15).setValue(countP4);
    dashboardSheet.getRange(startRow+2,  startColumn+16).setValue(countP5);
    dashboardSheet.getRange(startRow+2,  startColumn+17).setValue(countPN);
  } 
  
  // Insert update time
  dashboardSheet.getRange('AA2').setValue(d.toString()).setFontColor('#ffffff');
}


function prioritizeBugs (resultBugs){
  for (var i=0; i < resultBugs.bugs.length; i++)
    {
      switch (resultBugs.bugs[i].priority) 
      {
        case '--':
          countPN+=1;
          break;
        case 'P1':
          countP1+=1;
          break;
        case 'P2':
          countP2+=1; 
          break;
        case 'P3':
          countP3+=1; 
          break;
        case 'P4':
          countP4+=1;   
          break;
        case 'P5':
          countP5+=1;   
          break;
      }     
    } 
}

// Send mail notification
function sendStatusEmail() {  
  var startRow = 2;  // First row of data to process
  var numRows = 100;   // Number of rows to process
  
  // Fetch the range of cells B2
  var dataRange = mailingSheet.getRange(startRow, 1, numRows, 2)
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  // Get all bug counts
  var overallCell = dashboardSheet.getRange('C3');
  var overallFormula = overallCell.getFormula();
  var overallLink = overallFormula.substring(12, overallFormula.indexOf("\","));
  
  // Get TDC firefox bug counts
  var tdcFirefoxCell = dashboardSheet.getRange('E3');
  var tdcFirefoxFormula = tdcFirefoxCell.getFormula();
  var tdcFirefoxlLink = tdcFirefoxFormula.substring(12, tdcFirefoxFormula.indexOf("\","));
  
  // Get TDC platform bug counts
  var tdcPlatformCell = dashboardSheet.getRange('F3');
  var tdcPlatformFormula = tdcPlatformCell.getFormula();
  var tdcPlatformlLink = tdcPlatformFormula.substring(12, tdcPlatformFormula.indexOf("\","));
  
  // Get TDC total
  var tdcTotalCell = dashboardSheet.getRange('G3')
  
  // Get TDC percentage
  var tdcPercentageCell = dashboardSheet.getRange('H3')
  
  // Send mails
  for (i in data) {
    var row = data[i];
    var name = row[0];
    var email = row[1];
    
    if (email == '')
      break;
    
    var subject = "TDC Weekly Bug Status Update, "+ d.toDateString();
    var emailBody = "Dear all,<br><br>"
    + "  Please check Taipei weekly bug count update:<br>"
    + "  Detail: <a href=\"https://goo.gl/gKXPv6\"> TDC Bug Status </a><br><br>"
    + "<table style=border-style:solid;border-width:1px;border-collapse:collapse;border-spacing:1;border:1>"
    + "<tr>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Current Version</th>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Release Date</th>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Mozilla Total</th>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Taipei Firefox</th>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Taipei Platform</th>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Taipei Total</th>"
    + "<th style=background-color:#2b90d9;color:#ffffff;text-align:center;padding:10px 5px>Percentage</th>"
    + "</tr>"
    + "<tr>"
    + "<td style=text-align:center>"+currentVersion+"</td>"
    + "<td style=text-align:center>"+releaseDate+"</td>"
    + "<td style=text-align:center;background-color:#D2E4FC><a href=\"" + overallLink + "\">" + overallCell.getValue() + "</a></td>"
    + "<td style=text-align:center><a href=\"" + tdcFirefoxlLink + "\">" + tdcFirefoxCell.getValue() + "</a></td>"
    + "<td style=text-align:center;background-color:#D2E4FC><a href=\"" + tdcPlatformlLink + "\">" + tdcPlatformCell.getValue() + "</a></td>"
    + "<td style=text-align:center>"+tdcTotalCell.getValue() +"</td>"
    + "<td style=text-align:center;background-color:#D2E4FC>"+ (tdcPercentageCell.getValue() * 100).toFixed(2)  +"%</td>"
    + "</tr>"
    + "</table>"
    // Attach charts
    + "<table style=border-width:0;border-collapse:collapse;border-spacing:1;border:0>"
    + "<tr>"
    + "<p align='left'>"
    + "<img src='https://docs.google.com/spreadsheets/d/1s2LCo4Raba0dni--pNi5m-MegGsihvPPhQzlI80U08g/pubchart?oid=1859949787&format=image'></p>"
    + "</tr>"
    + "<tr>"
    + "<p align='left'>"
    + "<img src='https://docs.google.com/spreadsheets/d/1s2LCo4Raba0dni--pNi5m-MegGsihvPPhQzlI80U08g/pubchart?oid=1380858210&format=image'>"
    + "<img src='https://docs.google.com/spreadsheets/d/1s2LCo4Raba0dni--pNi5m-MegGsihvPPhQzlI80U08g/pubchart?oid=1653645995&format=image'></p>"
    + "</tr>"
    + "Josh Cheng<br>"
    + "Engineering Program Manager, Mozilla<br>"
    + "✉ joshcheng@mozilla.com";
    
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: emailBody,
    });
  }
}



// Send REST request to url 
function sendRequest (url)
{
  var response = UrlFetchApp.fetch(
    url,
    {
      method: "GET",
      contentType: "application/json",
      muteHttpExceptions: true,
    }
  );
  var responseCode = response.getResponseCode()
  var responseBody = response.getContentText()
  
  if (responseCode === 200) {
    responseJson = JSON.parse(responseBody)
    return responseJson;
  } else {
    Logger.log(Utilities.formatString("Request failed. Expected 200, got %d: %s", responseCode, responseBody))
  }
}

// create weekly records
function saveWeeklyRecords() {
  // insert new row
  weeklySheet.insertRows(3);
  
  // Get latest records from dashboard sheet
  var dataRange = dashboardSheet.getRange(3, 2, 1, 19);
  var data = dataRange.getValues();
  data[0][0] = currentVersion;
    
  weeklySheet.getRange(3, 2, 1, 19).setValues(data);

  // Set Weekly date
  weeklySheet.getRange(3, 1).setValue(d.toDateString().substring(4, 15));

}


function doGet(request) {
  dashboardSheet = SpreadsheetApp.openById("1s2LCo4Raba0dni--pNi5m-MegGsihvPPhQzlI80U08g").getSheetByName("Overall Dashboard");
  var overAllData = dashboardSheet.getRange(3,  2, 100, 26).getValues();
  
  var jsonResult = {
    Versions: []
  };
 
  for (var i=0; i < overAllData.length; i++)
  {
    if (overAllData[i][0] == "")
      break;
    
    
    jsonResult.Versions.push({ 
      "Version":            overAllData[i][0],
      "Mozilla_Total":      overAllData[i][1],
      "Mozilla_Total_Link": /"(.*?)"/.exec(dashboardSheet.getRange(i+3, 3).getFormula())[1],
      "TotalwoTDC":         overAllData[i][2],
      "TDC_Firefox":        overAllData[i][3],
      "TDC_Firefox_Link":   /"(.*?)"/.exec(dashboardSheet.getRange(i+3, 5).getFormula())[1],
      "TDC_Platform":       overAllData[i][4],
      "TDC_Platform_Link":  /"(.*?)"/.exec(dashboardSheet.getRange(i+3, 6).getFormula())[1],
      "TDC_Total":          overAllData[i][5],
      "TDC_Percentage":     overAllData[i][6],
      "Overall_P1":         overAllData[i][7],
      "Overall_P2":         overAllData[i][8],
      "Overall_P3":         overAllData[i][9],
      "Overall_P4":         overAllData[i][10],
      "Overall_P5":         overAllData[i][11],
      "Overall_PN":         overAllData[i][12],
      "TDC_P1":             overAllData[i][13],
      "TDC_P2":             overAllData[i][14],
      "TDC_P3":             overAllData[i][15],
      "TDC_P4":             overAllData[i][16],
      "TDC_P5":             overAllData[i][17],
      "TDC_PN":             overAllData[i][18],
      "P1_Percentage":      overAllData[i][19],
      "P2_Percentage":      overAllData[i][20],
      "P3_Percentage":      overAllData[i][21],
      "P4_Percentage":      overAllData[i][22],
      "P5_Percentage":      overAllData[i][23],
      "PN_Percentage":      overAllData[i][24],
    });
  } 
    
  return ContentService.createTextOutput(request.parameters.callback + "(" + JSON.stringify(jsonResult)+ ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  
}
