function generateAccessToken(refresh_token) {
  var url = 'https://login.questrade.com/oauth2/token?grant_type=refresh_token&refresh_token=' + refresh_token;
  var response = UrlFetchApp.fetch(url);
  //Logger.log(response);
  //sample response from Questrade
  /*
{"access_token":"the_actual_access_token_key","api_server":"https:\/\/api01.iq.questrade.com\/", "expires_in":1800, "refresh_token":"the_actual_refresh_token_key","token_type":"Bearer"}
  */
  return response;  
}

function appendToSheet(today, pl_cad, pl_usd, pl_all, pl_percent, total_investment) {
  let values = [
      [today, pl_cad, pl_usd, pl_all, pl_percent, total_investment]
      // Additional rows ...
    ];
    try {
      let valueRange = Sheets.newRowData();
      valueRange.values = values;
      let appendRequest = Sheets.newAppendCellsRequest();
      appendRequest.sheetId = PropertiesService.getScriptProperties().getProperty('spreadsheetId');
      appendRequest.rows = [valueRange];
      const result = Sheets.Spreadsheets.Values.append(valueRange, PropertiesService.getScriptProperties().getProperty('spreadsheetId'), PropertiesService.getScriptProperties().getProperty('range'), {valueInputOption: 'USER_ENTERED'});
      return result;
    } catch (err) {
      Logger.log('Failed with error %s', err.message);
    }
}

function getPortfolioBalances() {
  var response = JSON.parse(generateAccessToken(PropertiesService.getScriptProperties().getProperty('refresh_token')).getContentText());
  var api_server = response.api_server;
  //Logger.log(api_server);
  var access_token = response.access_token;
  var refresh_token = response.refresh_token;
  Logger.log(refresh_token);
  var url =  api_server + 'v1/accounts/' + PropertiesService.getScriptProperties().getProperty('account_no') + '/balances';
  var balances = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + access_token
    }
  });
  Logger.log(balances);
  //Sample JSON response for balances
  /*
{
     "perCurrencyBalances": [
         {
              "currency": "CAD",
               "cash": 243971.7,
               "marketValue":  6017,
               "totalEquity":  249988.7,
               "buyingPower": 496367.2,
               "maintenanceExcess": 248183.6,
              "isRealTime": false
         },
         {
              "currency": "USD",
              "cash": 198259.05,
              "marketValue":  53745,
              "totalEquity":  252004.05,
              "buyingPower":  461013.3,
              "maintenanceExcess":  230506.65,
             "isRealTime": false
          }
     ],
         "combinedBalances": [
                 ...
     ],
         "sodPerCurrencyBalances": [
                 ...
     ],
        "sodCombinedBalances": [
                 ...
     ]
 }
  */
  
  //update refresh_token to use for next time
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('refresh_token', refresh_token);

  var today = Utilities.formatDate(new Date(), "GMT-6", "MM/dd/yyyy")
  //Logger.log(today);
  balances = JSON.parse(balances);
  var pl_cad = balances.perCurrencyBalances[0].totalEquity;
  //Logger.log(pl_cad);
  var pl_usd = balances.perCurrencyBalances[1].totalEquity;
  //Logger.log(pl_usd);
  var pl_all = balances.combinedBalances[0].totalEquity;
  //Logger.log(pl_all);
  
  //calculate profit/loss in percentage based on last pl_all value
  var pl_percent = ((pl_all - PropertiesService.getScriptProperties().getProperty('pl_all_yesterday'))/pl_all) * 100;
  //Logger.log(pl_percent);
  
  //update today's pl_all value to calculate pl_percent for next execution
  scriptProperties.setProperty('pl_all_yesterday', pl_all);

  appendToSheet(today, pl_cad, pl_usd, pl_all, pl_percent, PropertiesService.getScriptProperties().getProperty('total_investment'));
}
