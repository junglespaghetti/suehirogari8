const accountKeys = "guaranteedStopLossOrderMode,hedgingEnabled,createdTime,currency,alias,marginRate,lastTransactionID,balance,openTradeCount,openPositionCount,pendingOrderCount,pl,resettablePL,resettablePLTime,financing,commission,dividendAdjustment,guaranteedExecutionFees,unrealizedPL,NAV,marginUsed,marginAvailable,positionValue,marginCloseoutUnrealizedPL,marginCloseoutNAV,marginCloseoutMarginUsed,marginCloseoutPositionValue,marginCloseoutPercent,withdrawalLimit,marginCallEnterTime"
const positionskeys = "instrument,pl,resettablePL,financing,commission,dividendAdjustment,guaranteedExecutionFees,unrealizedPL,marginUsed"
const longShortKeys = "units,averagePrice,pl,resettablePL,financing,dividendAdjustment,guaranteedExecutionFees,tradeIDs,unrealizedPL"
const tradesKey = "Date,id,instrument,price,openTime,initialUnits,initialMarginRequired,state,currentUnits,realizedPL,financing,dividendAdjustment,unrealizedPL,marginUsed"

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var days = PropertiesService.getUserProperties().getProperty("atr_days") || 20;
  var accountType = PropertiesService.getUserProperties().getProperty("accountType") || "demo";
  ui.createMenu(multiLang('Trade'))
  .addItem(multiLang('Import'), 'openDialog')
  .addItem(multiLang('Add current sheet to list'), 'setAccountID')
  .addItem(multiLang('Oanda order'), 'onClickItem2')
  .addItem(multiLang('ATR '+days+' days'), "setATRdays")
    .addSeparator()
    .addSubMenu(
      ui.createMenu(multiLang('Oanda settings'))
      .addItem(multiLang('Account ID setting'), "setAccountID")
      .addItem(multiLang('Authentication key setting'), "setAuth")
      .addItem(multiLang('Account type ('+accountType+')'), "setAccountType")
    )
    .addToUi();
  createWorkSheets()
}

function updateSheets(){
  addSummary();
  addPosition();
}

// menu function

function openDialog() {
  var html = HtmlService.createTemplateFromFile("html/load_dialog.html");
  html.fName = multiLang('File Name:')
  html.buttonTitle = multiLang('Import')
  SpreadsheetApp.getUi().showModelessDialog(html.evaluate(),multiLang('Load of a local file'));
}

function writeSheetupload(formObject) {
 
  // フォームで指定したテキストファイルを読み込む
  var fileBlob = formObject.myFile;
  
  // テキストとして取得（Windowsの場合、文字コードに Shift_JIS を指定）
  var text = fileBlob.getDataAsString("sjis");  
  
  // 改行コードで分割し配列に格納する
  var textLines = text.split(/[\s]+/);
  
  // 書き込むシートを取得
  var sheet = SpreadsheetApp.getActiveSheet();

  // テキストファイルをシートに展開する
  for (var i = 0; i < textLines.length; i++) {
    sheet.getRange(i + 1, 1).setValue(textLines[i]);
  }
  
  // 処理終了のメッセージボックスを出力
  Browser.msgBox("ローカルファイルを読み込みました");
}

function setATRdays(){
  var days = PropertiesService.getUserProperties().getProperty("atr_days") || 20;
  var nawDays = Browser.inputBox(multiLang('Change ATR '+days+' days'),Browser.Buttons.OK_CANCEL);
  if(nawDays == "cancel" || nawDays == ""){
    Browser.msgBox(multiLang('ATR '+days+' days'));
    return;
  } else {
    PropertiesService.getUserProperties().setProperty("atr_days", nawDays);
    Browser.msgBox(multiLang('ATR '+ nawDays +' days'));
    onOpen();
  }
}

function setAccountID(){
  var Account = PropertiesService.getUserProperties().getProperty("account");
  var newAccount = Browser.inputBox(multiLang('Account ID setting\n') + Account,Browser.Buttons.OK_CANCEL);
  if(newAccount == "cancel" || newAccount == ""){
    Browser.msgBox(multiLang('Current Account ID ') + Account);
    return;
  } else {
    PropertiesService.getUserProperties().setProperty("account", newAccount);
    Browser.msgBox(multiLang('Set to  '+ newAccount));
    onOpen();
  }
}

function setAuth(){
  var auth = PropertiesService.getUserProperties().getProperty("auth");
  var nawAuth = Browser.inputBox(multiLang('Authentication key setting\n')+auth,Browser.Buttons.OK_CANCEL);
  if(nawAuth == "cancel" || nawAuth == ""){
    Browser.msgBox(multiLang('Current authentication key ')+ auth);
    return;
  } else {
    PropertiesService.getUserProperties().setProperty("auth", nawAuth);
    Browser.msgBox(multiLang('Set to  '+ nawAuth));
    onOpen();
  }
}

function setAccountType(){
  var accountType = PropertiesService.getUserProperties().getProperty("accountType") || "demo";
  var newType = accountType;
  if(accountType == "demo"){
    newType = "live";
  }else if(accountType == "live"){
    newType = "demo";
  }
  var result=Browser.msgBox(multiLang("Change account type to "+newType+"?"),Browser.Buttons.OK_CANCEL);
  if(result=="ok"){
    PropertiesService.getUserProperties().setProperty("accountType", newType);
    Browser.msgBox(multiLang("Changed account type to "+newType+"."));
    onOpen();
  }
}

// create sheets

function createWorkSheets(){
  if(!isSheetNameIncluded(multiLang("WatchList"))){
    var watchList = SpreadsheetApp.getActiveSpreadsheet().insertSheet(multiLang("WatchList"),0);
    watchList.appendRow(getWatchListHeader());
  }
  if(!isSheetNameIncluded("OandaPositions")){
   var summarys = SpreadsheetApp.getActiveSpreadsheet().insertSheet("OandaPositions",2);
    summarys.appendRow(getPositionHeader());
  } 
  if(!isSheetNameIncluded("OandaSummarys")){
   var summarys = SpreadsheetApp.getActiveSpreadsheet().insertSheet("OandaSummarys",3);
    summarys.appendRow(getSummarySheetHeader());
  }
  if(!isSheetNameIncluded(multiLang("ErrorLog"))){
    var ErrorLog = SpreadsheetApp.getActiveSpreadsheet().insertSheet(multiLang("ErrorLog"),4);
    ErrorLog.appendRow(getErrorLogHeader());
  }  
  if(!isSheetNameIncluded(multiLang("OandaTrades"))){
    var ErrorLog = SpreadsheetApp.getActiveSpreadsheet().insertSheet("OandaTrades",4);
  //  ErrorLog.appendRow(getErrorLogHeader());
  }  
}

function getWatchListHeader(){
  var head = [];
  head.push(multiLang('Date'));
  head.push(multiLang('Broker'));
  head.push(multiLang('Symbol'));
  head.push(multiLang('Interval'));
  head.push(multiLang('ATR'));
  head.push(multiLang('Max Size'));
  return head
}

function getSummarySheetHeader(){
  var keys = accountKeys.split(",");
  var head = [];
  head.push('Date');
  keys.forEach(function(item){
    head.push(item);
  });
  return head
}

function getPositionHeader(){
  var keys = positionskeys.split(",");
  var head = [];
  head.push('Date');
  keys.forEach(function(item){
    head.push(item);
  });
  keys = longShortKeys.split(",");
  keys.forEach(function(item){
    head.push("L-"+item);
  });
  keys.forEach(function(item){
    head.push("S-"+item);
  });
  return head
}

function getErrorLogHeader(){
    var head = [];
  head.push(multiLang('Date'));
  head.push(multiLang('Function'));
  head.push(multiLang('Error'));
  head.push("");
  head.push("");
  head.push("");
  head.push("");
  head.push("");
  head.push("");
  head.push(multiLang('JsonString'));
  return head
}

// Alert

function doPost(e) {
  Logger.log(e);
  var params = JSON.parse(e.postData.getDataAsString());
  
var response = {
    data: responseList,
    meta: { status: 'success' }
  };
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}


function doGet(e) {
  Logger.log("get");
  Logger.log(e);
  var param = e.parameter.param;
  Logger.log(param);
  return ContentService.createTextOutput(param);
}

//  Watch list

function addWatchSheet(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = multiLang("WatchList");
  var sName1 = sheet.getSheetName().split(",");
  var sName2 = sName1[0].split("_");
  if(sName1.length == 2 && sName2.length == 2){
  var watchList = SpreadsheetApp.getActive().getSheetByName(sheetName); 
    if(watchList == null){
     createWorkSheets();
      watchList = SpreadsheetApp.getActive().getSheetByName(sheetName); 
    }
    var days = PropertiesService.getUserProperties().getProperty("atr_days") || 20;
    var arr = [];
    var data = sheet.getRange(sheet.getLastRow()-Number(days),1,Number(days)+1,5).getValues();
    arr.push(data[data.length-1][0]);
    arr.push(sName2[0]);
    arr.push(sName2[1]);
    arr.push(sName1[1]);
    arr.push(ArrToATR(data,2,3,4));
//    Logger.log(arr);
    watchList.appendRow(arr);
    
  }
}

function ArrToATR(data,high,low,close){ //An array of days, plus one day's worth of the previous day's closing price.
  var atr = 0;
  for(var i=1;i<data.length;i++){
    if(data[i][high]-data[i][low] >= data[i][high]-data[i-1][close] && data[i][high]-data[i][low] >= data[i][low]-data[i-1][close]){
      atr += data[i][high]-data[i][low];
    }else if(data[i][high]-data[i][low] <= data[i][high]-data[i-1][close] && data[i][high]-data[i-1][close] >= data[i][low]-data[i-1][close]){
      atr += data[i][high]-data[i-1][close];
    }else if(data[i][high]-data[i][low] <= data[i][low]-data[i-1][close] && data[i][high]-data[i-1][close] <= data[i][low]-data[i-1][close]){
      atr += data[i][3]-data[i-1][close];
    }
  }
  return atr/(data.length-1); 
}

// Summary



// acount position

function addPosition(){
  var sheetName = multiLang("OandaPositions");
  var PositionsSheet = SpreadsheetApp.getActive().getSheetByName(sheetName); 
    if(PositionsSheet == null){
      createWorkSheets();
      PositionsSheet = SpreadsheetApp.getActive().getSheetByName(sheetName); 
    }
  var Positions = getPosition();
  if(Positions && Positions.account){
    var keys = positionskeys.split(",");
    var keys2 = longShortKeys.split(",");
    var arr = [];
    
    PositionsSheet.clear();
    PositionsSheet.appendRow(getPositionHeader());
    
    arr.push(new Date);
    for(var i=0 ;i <= Positions.account.positions.length-1 ; i++){
      var arr = [];
      arr.push(new Date);
      keys.forEach(function(key){
        arr.push(Positions.account.positions[i][key]);
      });
      keys2.forEach(function(key){
        if(key == "tradeIDs" && Positions.account.positions[i].long[key]){
          arr.push(Positions.account.positions[i].long[key].join(","));
        }else{
          arr.push(Positions.account.positions[i].long[key]);
        }
      });
      keys2.forEach(function(key){
        if(key == "tradeIDs" && Positions.account.positions[i].short[key]){
           arr.push(Positions.account.positions[i].short[key]);
        }else{
          arr.push(Positions.account.positions[i].short[key]);
        }
        });
      PositionsSheet.appendRow(arr);
    }
  }
}

function getPosition(){

  var Authorization = 'Bearer '+ PropertiesService.getUserProperties().getProperty("auth");
  var accountID = PropertiesService.getUserProperties().getProperty("account");
  var accountType = PropertiesService.getUserProperties().getProperty("accountType");
  var url = changeFQN(accountType) + "/v3/accounts/"+ accountID;
  
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': Authorization
  };  

  var options = {
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = JSON.parse(UrlFetchApp.fetch(url, options));
  
  if(!response.errorMessage){
    
    return response;
    
  }else{
    
    var errorLog = SpreadsheetApp.getActive().getSheetByName(multiLang("ErrorLog"));
    var arr =[];
    arr.push(new Date);
    arr.push("position");
    arr.push(response.errorMessage);
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push(JSON.stringify(response));
    errorLog.appendRow(arr);    
    return 
    
  }
}

function addSummary(){
  var sheetName = multiLang("OandaSummarys");
  var summarySheet = SpreadsheetApp.getActive().getSheetByName(sheetName); 
    if(summarySheet == null){
      createWorkSheets();
      summarySheet = SpreadsheetApp.getActive().getSheetByName(sheetName); 
    }
  var summary = getSummary();
  if(summary && summary.account){
    var keys = accountKeys.split(",");
    var arr = [];
    arr.push(new Date);
    keys.forEach(function(item){
      arr.push(summary.account[item]);
    });
    Logger.log(arr);
   summarySheet.appendRow(arr);
  }

}

function getSummary(){

  var Authorization = 'Bearer '+ PropertiesService.getUserProperties().getProperty("auth");
  var accountID = PropertiesService.getUserProperties().getProperty("account");
  var accountType = PropertiesService.getUserProperties().getProperty("accountType");
  var url = changeFQN(accountType) + "/v3/accounts/"+ accountID+"/summary";
  
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': Authorization
  };  

  var options = {
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = JSON.parse(UrlFetchApp.fetch(url, options));
  
  if(!response.errorMessage){
    
    return response;
    
  }else{
    
    var errorLog = SpreadsheetApp.getActive().getSheetByName(multiLang("ErrorLog"));
    var arr =[];
    arr.push(new Date);
    arr.push("summary");
    arr.push(response.errorMessage);
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push(JSON.stringify(response));
    errorLog.appendRow(arr);    
    return 
    
  }
}


// instruments list

function addTrades(){
  var sheetName = multiLang("OandaTrades");
  var tradesSheet = SpreadsheetApp.getActive().getSheetByName(sheetName); 
  if(tradesSheet == null){
    createWorkSheets();
    tradesSheet = SpreadsheetApp.getActive().getSheetByName(sheetName); 
  }
  var trades = getTrades();
  Logger.log(trades.trades[0]);
  if(trades && trades.trades){
    var keys = tradesKey.split(",");
  
    for(var i=0 ;i <= trades.trades.length-1 ; i++){
      var arr =[];
      arr.push(new Date)
      keys.forEach(function(key){
        arr.push(trades.trades[i][key]);
      });
      Logger.log(arr);
      tradesSheet.appendRow(arr);
  //Logger.log(instruments.instruments[i]);
    }
  }
}

function getTrades(){

  var Authorization = 'Bearer '+ PropertiesService.getUserProperties().getProperty("auth");
  var accountID = PropertiesService.getUserProperties().getProperty("account");
  var accountType = PropertiesService.getUserProperties().getProperty("accountType");
  var url = changeFQN(accountType) + "/v3/accounts/"+ accountID+"/trades";
  
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': Authorization
  };  

  var options = {
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = JSON.parse(UrlFetchApp.fetch(url, options));
  
  if(!response.errorMessage){
    
    return response;
    
  }else{
    
    var errorLog = SpreadsheetApp.getActive().getSheetByName(multiLang("ErrorLog"));
    var arr =[];
    arr.push(new Date);
    arr.push("trades");
    arr.push(response.errorMessage);
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push(JSON.stringify(response));
    errorLog.appendRow(arr);    
    return 
    
  }
}

function makeATR(){
  
  var data = {
    
      "granularity": "D",
    "count":20
  };
  var response = getCandles(data);
  
  var arr = [];
  
  Object.keys(response.candles).forEach(function (key){
    
    
    
  })
  
  Logger.log(response);
  
}

function getCandles(obj){

  var Authorization = 'Bearer '+ PropertiesService.getUserProperties().getProperty("auth");
  var accountID = PropertiesService.getUserProperties().getProperty("account");
  var instrument = "USD_JPY";
  var accountType = PropertiesService.getUserProperties().getProperty("accountType");
  var url = changeFQN(accountType) + "/v3/accounts/"+ accountID +"/instruments/"+ instrument + "/candles"+objToParameter(obj);

  var headers = {
    'Content-Type': 'application/json',
    'Authorization': Authorization
  };
  
  var options = {
    "method": "get",
    "headers": headers,
    "muteHttpExceptions": true
  };

  var response = JSON.parse(UrlFetchApp.fetch(url, options));

  if(!response.errorMessage){
    
    return response;
    
  }else{
    
    var errorLog = SpreadsheetApp.getActive().getSheetByName(multiLang("ErrorLog"));
    var arr =[];
    arr.push(new Date);
    arr.push("candles");
    arr.push(response.errorMessage);
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push("");
    arr.push(JSON.stringify(response));
    errorLog.appendRow(arr);    
    return 
    
    }
}

function objToParameter(obj){
  if(obj instanceof Object && !(obj instanceof Array)){
    var arr = [];
    Object.keys(obj).forEach(function (key){
      arr.push(key+"="+obj[key]);
    })
    Logger.log("aaa")
    return "?"+arr.join("&");
  }
  return "";
}

function makeOrder(){
    //APIパラメータ
  var instrument   = "USD_JPY" ; //通貨ペア
  var units        = "10" ; //枚数(-は売り、なしは買い)
  var type         = "MARKET" ; //注文条件(成行・指値等)
  var positionFill = "DEFAULT" ; //
  var timeInForce  = "FOK" ; //

  /** order情報 **/
  var data = {
    "order": {
      "instrument": instrument,
      "units": units,
      "type": type,
      "positionFill": positionFill,
      "timeInForce": timeInForce
    }
  };
  
  var response = postOanda(data);
  
  Logger.log(response);
  
}

function changeFQN(accountType){
  if(accountType == "demo"){
    return 'https://api-fxpractice.oanda.com';
  }else if(accountType == "live"){
    return 'https://api-fxtrade.oanda.com';
  }
}

// Order

function postOanda(order){

  var Authorization = 'Bearer '+ PropertiesService.getUserProperties().getProperty("auth");
  var accountID = PropertiesService.getUserProperties().getProperty("account");
  var accountType = PropertiesService.getUserProperties().getProperty("accountType");
  var FQDN = 'api-fxpractice.oanda.com'; //デモFQDN
  var url = "https://" + FQDN + "/v3/accounts/"+ accountID +"/orders";
  
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': Authorization
  };  

  var options = {
    "method": "POST",
    "payload": JSON.stringify(order),
    "headers": headers,
    "muteHttpExceptions": true
  };

  return UrlFetchApp.fetch(url, options);
}

 /*
 * postOrders
 */
function postOrders(order){

  var Authorization = 'Bearer '+ PropertiesService.getUserProperties().getProperty("auth");
  var accountID = PropertiesService.getUserProperties().getProperty("account");
  var accountType = PropertiesService.getUserProperties().getProperty("accountType");
  var FQDN = 'api-fxpractice.oanda.com'; //デモFQDN
  var url = "https://" + FQDN + "/v3/accounts/"+ accountID +"/orders";
  
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': Authorization
  };  

  var options = {
    "method": "POST",
    "payload": JSON.stringify(order),
    "headers": headers,
    "muteHttpExceptions": true
  };

  return UrlFetchApp.fetch(url, options);
  
//  var response = UrlFetchApp.fetch(url, options);
//  var responseCode = response.getResponseCode();
//  var responseText = response.getContentText();
//  var responseData = JSON.parse(responseText);
//  Logger.log("responseCode:"+responseCode);//正常:201 異常:400等
//  Logger.log("responseText:"+responseText);//
/*
  //HTTPステータスコード：201（リクエストは成功し、その結果新たなリソースが作成された。POSTのレスポンス）
  if(responseCode == 201){
    Logger.log("【正常終了】")
  }else{
    Logger.log("【異常終了】" + responseText);
  }*/
}

// Multi language

function multiLang(str){
  var lang = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale().substr(0,2);
  return LanguageApp.translate(str,"",lang)
}

// etc

function isSheetNameIncluded(name){
 var flag = SpreadsheetApp.getActive().getSheetByName(name);
  if(flag == null){
    return false;
  }else{
    return true;
  }
}
