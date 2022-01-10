//=====================================================================================================================
//
// Name: TweetFetcher
//
// Desc:
//
// Author:  Mune
//
// History:
//  2021-12-22 : Initial version
//
//=====================================================================================================================
// ID of Target Google Spreadsheet (Book)
let VAL_ID_TARGET_BOOK           = '1PQdZeCzHgk_pktFmzCK-dbiCipgDOzESjDoU5yi51xY';
// ID of the Google Drive where the images will be placed
let VAL_ID_GDRIVE_FOLDER_MEDIA   = '1g8lf_LcSf9zMoB-GAChlciC6JovjYqTY';
// ID of Google Drive to place backup Spreadsheet
let VAL_ID_GDRIVE_FOLDER_BACKUP  = '1qFPnaZeljn8Gr4twG6484FbAI5qJP0Ki';
// Key and Secret to access Twitter APIs
let VAL_CONSUMER_API_KEY         = 'vN1to7V6dL1HqXkmg07Owuejr';
let VAL_CONSUMER_API_SECRET      = 'fR3UCVqYlSa6q4j4V2HUkFe3rwXLr58CMrQpLDmgA9AaNxRwDR';

//=====================================================================================================================
// DEFINES
//=====================================================================================================================
let VERSION                      = 1.0;
let TIME_LOCALE                  = "JST";
let FORMAT_DATETIME_ISO8601_DATE = "yyyy-MM-dd";
let FORMAT_DATETIME_ISO8601_TIME = "HH:mm:ss";
let FORMAT_DATETIME_DATE_NUM     = "yyyyMMdd";
let FORMAT_DATETIME              = "yyyy-MM-dd HH:mm:ss";
let FORMAT_TIMESTAMP             = "yyyyMMddHHmmss";
let NAME_SHEET_USAGE             = "!USAGE";
let NAME_SHEET_LOG               = "!LOG";
let NAME_SHEET_ERROR             = "!ERROR";
let SHEET_NAME_COMMON_SETTINGS   = "%Screen Names%";

class HeaderRow {
  id_str                         : any;
  link                           : any;
  created_at                     : any;
  text                           : any;
  in_reply_to_screen_name        : any;
  retweet_count                  : any;
  favorite_count                 : any;
  media                          : any;
}
let HEADER_TITLES:HeaderRow = {
  id_str                         : "Tweet Id",
  link                           : "link",
  created_at                     : "Created at",
  text                           : "Tweet",
  in_reply_to_screen_name        : "Reply to",
  retweet_count                  : "Retweet Count",
  favorite_count                 : "Favorite Count",
  media                          : "Media"
}
class HeaderInfo {
  screenName                     : string;
  rowHeader                      : number;
  headerCols                     : HeaderRow;
}

let MAX_ROW_SEEK_HEADER          = 20;
let MAX_COL_SEEK_HEADER          = Object.keys(HEADER_TITLES).length;
let DEFAULT_ROW_HEADER           = 4;

let OFFSET_ROW_SCREENNAME_LIST   = 0;

//=====================================================================================================================
// GLOBALS
//=====================================================================================================================
let g_isDebugMode                = true;
let g_isEnabledLogging           = true;
let g_isDownlodingMedia          = true;
let g_datetime                   = new Date();
let g_timestamp                  = TIME_LOCALE + ": " + Utilities.formatDate(g_datetime, TIME_LOCALE, FORMAT_DATETIME);
let g_folderMedia                = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_MEDIA);
let g_folderBackup               = DriveApp.getFolderById(VAL_ID_GDRIVE_FOLDER_BACKUP);
let g_book                       = SpreadsheetApp.openById(VAL_ID_TARGET_BOOK);

//=====================================================================================================================
// CODE for General
//=====================================================================================================================

//
// Name: gsGetValuesWithFormulas
// Desc:
//  return 2D array with formulas
//
function gsGetValuesWithFormulas(range){
  var valsRng = range.getValues();
  //Logger.log("valuesAndFomulas:%s", valuesAndFomulas);
  var formulasRng = range.getFormulas();
  //Logger.log("tempFormulas:%s", tempFormulas);

  for(let r = 0; r < valsRng.length; r++){
    for(let c = 0; c < valsRng[0].length; c++){
      if(formulasRng[r][c].length != 0){
        valsRng[r][c] = formulasRng[r][c];
      }else{
        ;
      }
    }
  }
  return valsRng;
}

//
// Name: gsAddLineAtLast
// Desc:
//  Add the specified text at the bootom of the specified sheet.
//
function gsAddLineAtBottom(sheetName, text) {
  try {
    let sheet = g_book.getSheetByName(sheetName);
    if (!sheet) {
      sheet = g_book.insertSheet(sheetName, g_book.getNumSheets());
    }
    let range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2);
    if (range) {
      let valsRng = range.getValues();
      let row = valsRng[0];
      row[0] = g_timestamp;
      row[1] = String(text);
      range.setValues(valsRng);
    }
  }
  catch (e) {
    Logger.log("EXCEPTION: gsAddLineAtBottom: " + e.message);
  }
}

//
// Name: logOut
// Desc:
//
function logOut(text) {
  text = g_timestamp + "\t" + text;
  if (!g_isEnabledLogging) {
    return;
  }
  gsAddLineAtBottom(NAME_SHEET_LOG, text);
}
//
// Name: errOut
// Desc:
//
function errOut(text) {
  text = g_timestamp + "\t" + text;
  gsAddLineAtBottom(NAME_SHEET_ERROR, text);
}

//=====================================================================================================================
// CODE for TwitterAPI
//=====================================================================================================================
function logOAuthURL() {
  let twitterService = getTwitterService();
  Logger.log(twitterService.authorize());
}
function getTwitterService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth1.createService('twitter')
    // Set the endpoint URLs.
    .setAccessTokenUrl('https://api.twitter.com/oauth/access_token')
    .setRequestTokenUrl('https://api.twitter.com/oauth/request_token')
    .setAuthorizationUrl('https://api.twitter.com/oauth/authorize')
    // Set the consumer key and secret.
    .setConsumerKey(VAL_CONSUMER_API_KEY)
    .setConsumerSecret(VAL_CONSUMER_API_SECRET)
    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction('authCallback')
    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties());
}
function resetTwitterService() {
  let twitterService = getTwitterService();
  twitterService.reset();
}
function authCallback(request) {
  let twitterService = getTwitterService();
  let isAuthorized = twitterService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  }
  else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}
function twitterSearch(keyword, maxCount) {
  let encodedKeyword = encodeURIComponent(keyword);
  try {
    let twitterService = getTwitterService();
    if (!twitterService.hasAccess()) {
      //Logger.log(twitterService.getLastError());
      return null;
    }
    let url = 'https://api.twitter.com/1.1/search/tweets.json?q='
      + encodedKeyword
      + '&result_type=recent&lang=ja&locale=ja&count='
      + maxCount;
    let response = twitterService.fetch(url, { method: "GET" });
    let tweets = JSON.parse(response);
    //Logger.log(tweets);
    return tweets;
  }
  catch (ex) {
    Logger.log(ex);
    return null;
  }
}
function twitterUserTimeline(screenName, maxCount, trimUser, excludeReplies, includeRts) {
  try {
    let twitterService = getTwitterService();
    if (!twitterService.hasAccess()) {
      //Logger.log(twitterService.getLastError());
      return null;
    }
    let url = 'https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name='
      + screenName
      + '&trim_user='
      + trimUser
      + '&exclude_replies='
      + excludeReplies
      + '&include_rts='
      + includeRts
      + '&count='
      + maxCount;
    let response = twitterService.fetch(url, { method: "GET" });
    let tweets = JSON.parse(response);
    // Logger.log(tweets);
    return tweets;
  }
  catch (ex) {
    Logger.log(ex);
    return null;
  }
}

//=====================================================================================================================
// CODE
//=====================================================================================================================

//
// Name: getHeaderInfo
// Desc: Seek the header info from the specified sheet.
//
function getHeaderInfo(sheet, headerTitles:HeaderRow):HeaderInfo {
  let range = sheet.getRange(1, 1, MAX_ROW_SEEK_HEADER, MAX_COL_SEEK_HEADER);
  if (range == null) {
    return null;
  }
  let valsRng = range.getValues();
  if (!valsRng) {
    return null;
  }
  // assuming that the screen_name exist at {row=1, col=1}
  let screenName = String(valsRng[0][0]);
  if (screenName) {
    screenName = screenName.trim();
  }
  if (!screenName) {
    return null;
  }
  let r = 1;
  for (; r < valsRng.length; r++) {
    let objRow = valsRng[r];
    let headerCols = new HeaderRow();
    for (let c = 0; c < objRow.length; c++) {
      let txtCell = String(objRow[c]);
      if (!txtCell) {
        continue;
      }
      for (let i = 0; i < Object.values(headerTitles).length; i++) {
        if (headerCols[Object.keys(headerTitles)[i]]) {
          continue;
        }
        if (Object.values(headerTitles)[i].toLowerCase().trim() == txtCell.toLowerCase().trim()) {
          headerCols[Object.keys(headerTitles)[i]] = c;
          break;
        }
      }
    }
    let i = 0;
    for (i = 0; i < Object.values(headerTitles).length; i++) {
      if (null == headerCols[Object.keys(headerTitles)[i]]) {
        break;
      }
    }
    if (i == Object.values(headerTitles).length) {
      return { screenName: screenName, rowHeader: r, headerCols: headerCols };
    }
  }
  return { screenName: screenName, rowHeader: null, headerCols: null };
}
//
// Name: generateHeader
// Desc:
//
function generateHeader(sheet, screenName:string, headerTitles:HeaderRow):HeaderInfo {
  if (sheet.getMaxRows() > 1) {
    sheet.deleteRows(2, sheet.getMaxRows() - 1);
  }
  if (sheet.getMaxColumns() > Object.values(headerTitles).length) {
    sheet.deleteColumns(Object.values(headerTitles).length + 1, sheet.getMaxColumns() - Object.values(headerTitles).length );
  }
  let range = sheet.getRange(1, 1, (DEFAULT_ROW_HEADER + 1), MAX_COL_SEEK_HEADER);
  if (range == null) {
    throw new Error("generateHeader: range wasn't able to acquired.");
  }
  let valsRng = range.getValues();
  let headerCols = new HeaderRow();
  let objRow = valsRng[DEFAULT_ROW_HEADER];
  for (let c = 0; c < Object.values(headerTitles).length; c++) {
    objRow[c] = Object.values(headerTitles)[c];
    headerCols[Object.keys(headerTitles)[c]] = c;
  }
  valsRng[0][0] = screenName;
  range.setValues(valsRng);
  return { screenName: screenName, rowHeader: DEFAULT_ROW_HEADER, headerCols: headerCols };
}

//
// Name: updateStoredTweets
// Desc:
//
function updateStoredTweets(tweets, sheet, headerInfo):number[] {
  if (0 == sheet.getLastRow() - (headerInfo.rowHeader + 1)) {
    return [];
  }
  let range = sheet.getRange(headerInfo.rowHeader + 2, 1, sheet.getLastRow() - (headerInfo.rowHeader + 1), sheet.getLastColumn());
  if (range == null) {
    throw new Error("updateStoredTweets: range wasn't able to acquired.");
  }
  let idxsUpdatedTweets = []; // array of indexes of handled tweets
  let valsRng = gsGetValuesWithFormulas(range);
  for (let t = 0; t < tweets.length && valsRng.length > idxsUpdatedTweets.length; t++) {
    for (let i = 0; i < valsRng.length; i++) {
      let objRow = valsRng[i];
      if (tweets[t].id_str == objRow[headerInfo.headerCols.id_str] ) {
        objRow[headerInfo.headerCols.favorite_count] = tweets[t].favorite_count;
        objRow[headerInfo.headerCols.retweet_count] = tweets[t].retweet_count;
        idxsUpdatedTweets.push(t);
        break;
      }
    }
  }
  range.setValues(valsRng);
  return idxsUpdatedTweets;
}

//
// Name: downloadMedia
// Desc:
//  Download media used in a tweet in the date folder.
//
function downloadMedia(tweet, dateCreatedAt) {
  let folderMedia = null;
  if (tweet.entities.media != undefined && tweet.entities.media[0].type == 'photo') {
    let strDate = Utilities.formatDate(dateCreatedAt, TIME_LOCALE, FORMAT_DATETIME_ISO8601_DATE);
    let foldersOfDate = g_folderMedia.getFoldersByName(strDate);
    let folderDate = null;
    if (foldersOfDate.hasNext()) {
      folderDate = foldersOfDate.next();
    }
    else {
      folderDate = g_folderMedia.createFolder(strDate);
    }
    folderMedia = folderDate.createFolder(tweet.id_str);
    if ( g_isDownlodingMedia ){
      for (let i = 0; i < tweet.extended_entities.media.length; i++) {
        let imageBlob = UrlFetchApp.fetch(tweet.extended_entities.media[i].media_url).getBlob();
        folderMedia.createFile(imageBlob);
      }
    }
  }
  return folderMedia;
}
//
// Name: addNewTweets
// Desc:
//
function addNewTweets(sheet, headerInfo, tweets, idxesUpdatedTweets:number[]) {
  if (0 == tweets.length - idxesUpdatedTweets.length) {
    return;
  }
  sheet.insertRowsAfter(headerInfo.rowHeader + 1, (tweets.length - idxesUpdatedTweets.length));
  // if new rows need be added under the bottom row...
  let range = sheet.getRange(headerInfo.rowHeader + 2, 1, (tweets.length - idxesUpdatedTweets.length), sheet.getLastColumn());
  if (range == null) {
    throw new Error("updateStoredTweets: range wasn't able to acquired.");
  }
  let valsRng = range.getValues();
  for (let t = 0, r = 0; t < tweets.length; t++) {
    if (-1 != idxesUpdatedTweets.indexOf(t)) {
      continue;
    }
    valsRng[r][headerInfo.headerCols.id_str         ] = tweets[t].id_str;
    valsRng[r][headerInfo.headerCols.link           ] = '=HYPERLINK("https://twitter.com/' + tweets[t].screen_name + '/status/' + tweets[t].id_str + '", "link")';
    valsRng[r][headerInfo.headerCols.created_at     ] = tweets[t].created_at;
    valsRng[r][headerInfo.headerCols.text           ] = tweets[t].text;
    //valsRng[r][headerInfo.headerCols.user_id_str    ] = tweets[t].user.id_str;
    if (tweets[t].in_reply_to_screen_name) {
      valsRng[r][headerInfo.headerCols.in_reply_to_screen_name] = '=HYPERLINK("https://twitter.com/' + tweets[t].in_reply_to_screen_name + '", "' + tweets[t].in_reply_to_screen_name + '")';
    }
    valsRng[r][headerInfo.headerCols.retweet_count  ] = tweets[t].retweet_count;
    valsRng[r][headerInfo.headerCols.favorite_count ] = tweets[t].favorite_count;
    if ( g_isDownlodingMedia ){
      let folderMedia = downloadMedia(tweets[t], new Date(tweets[t].created_at));
      if (folderMedia) {
        valsRng[r][headerInfo.headerCols.media] = '=HYPERLINK("' + folderMedia.getUrl() + '", "media")';
      }
    }
    r++;
  }
  range.setValues(valsRng);
}

//
// Name: getScreenNames
// Desc:
//
function getScreenNames( sheet ):string[] {
  let range = sheet.getRange(1 + OFFSET_ROW_SCREENNAME_LIST, 1, sheet.getLastRow()-OFFSET_ROW_SCREENNAME_LIST, 1); // Screen names need to be placed at the first column
  let valsRng = range.getValues();
  let listScreenNames:string[] = [];
  for ( let i=0; i<valsRng.length; i++) {
    if (valsRng[i][0]) {
      listScreenNames.push((valsRng[i][0]).trim());
    } else {
      listScreenNames.push(null);
    }
  }
  return listScreenNames;
}

//
// Name: updateSheet
// Desc:
//
function updateSheet( sheet, screenName:string ):boolean {
  let headerInfo = getHeaderInfo(sheet, HEADER_TITLES);
  if (!headerInfo) {
    headerInfo = generateHeader(sheet, screenName, HEADER_TITLES);
    sheet.setName(headerInfo.screenName);
  }
  let tweets = twitterUserTimeline(headerInfo.screenName, 50, true, false, true);
  if (tweets) {
    let idxUpdated = updateStoredTweets(tweets, sheet, headerInfo);
    if (tweets.length > idxUpdated.length) {
      addNewTweets(sheet, headerInfo, tweets, idxUpdated);
    }
  }
  return true;
}

//
// Name: main
// Desc:
//  Entry point of this program.
//
function main() {
  try {
    let sheetConfig = g_book.getSheetByName(SHEET_NAME_COMMON_SETTINGS);
    if (!sheetConfig) {
      throw new Error("The sheet \"" + SHEET_NAME_COMMON_SETTINGS + "\" was not found.");
    }
    let screenNames: string[] = getScreenNames(sheetConfig);
    for (let i = 0; i < screenNames.length; i++) {
      if (!screenNames[i]) {
        continue;
      }
      let sheet = g_book.getSheetByName(screenNames[i]);
      if (!sheet) {
        sheet = g_book.insertSheet(screenNames[i], g_book.getNumSheets());
      }
      let range = sheetConfig.getRange(i + 1 + OFFSET_ROW_SCREENNAME_LIST, 2, 1, 2);
      if ( updateSheet(sheet, screenNames[i] ) ) {
        range.setValues([["OK", g_timestamp]]);
      } else {
        range.setValues([["ERR", g_timestamp]]);
      }
    }
  }
  catch (ex) {
    errOut(ex.message);
  }
}
