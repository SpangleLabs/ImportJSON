/*====================================================================================================================================*
  ImportJSON by Roselyn Coales, Brad Jasper and Trevor Lohrbeer
  ====================================================================================================================================
  Version:      1.9.0
  Project Page: https://github.com/joshcoales/ImportJSON
  Copyright:    (c) 2021 by Roselyn Coales
                (c) 2017-2019 by Brad Jasper
                (c) 2012-2017 by Trevor Lohrbeer
  License:      GNU General Public License, version 3 (GPL-3.0) 
                http://www.opensource.org/licenses/gpl-3.0.html
  ------------------------------------------------------------------------------------------------------------------------------------
  A library for importing JSON feeds into Google spreadsheets. Functions include:

     ImportJSON            For use by end users to import a JSON feed from a URL 
     ImportJSONWithPost    For use by end users to import a JSON feed from a URL using POST parameters
     ImportJSONAdvanced    For use by script developers to easily extend the functionality of this library
     ParseJSON             For use by end users to import JSON data from string
     ParseJSONFromSheet    For use by end users to import JSON from one of the Sheets in the current spreadsheet
     ParseJSONAdvanced     For use by script developers to easily extend the json parsing functionality of this library
  
  It also adds an "Update JSON cache" button to the scripts menu, which allows caching JSON results in the spreadsheet

  For future enhancements see https://github.com/joshcoales/ImportJSON/issues?q=is%3Aissue+is%3Aopen+label%3Aenhancement
  
  For bug reports see https://github.com/joshcoales/ImportJSON/issues

  ------------------------------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.9.0  (October 6, 2021) Various xpath fixes, and rawJson output fix
  1.8.3  (October 6, 2021) Adding =ParseJSON() and ParseJSONAdvanced() helper method
  1.8.2  (October 6, 2021) Rename ImportJSONViaPost to ImportJSONWithPost
  1.8.1  (October 6, 2021) Remove ImportJSONBasicAuth, as URL can have basic auth info in it, and then use ImportJSON()
  1.8.0  (October 6, 2021) Adding "Update JSON cache" button to menu
  1.7.4  (October 6, 2021) Adding rawJson option
  1.7.3  (October 6, 2021) Return floats as floats
  1.7.2  (October 6, 2021) Ensuring xpath does not treat key as prefix
  1.7.1  (October 6, 2021) Adding retryFetch option
  1.7.0  (October 6, 2021) Renaming ImportJsonFromSheet to ParseJsonFromSheet
  1.6.0  (June 2, 2019) Fixed null values (thanks @gdesmedt1)
  1.5.0  (January 11, 2019) Adds ability to include all headers in a fixed order even when no data is present for a given header in some or all rows.
  1.4.0  (July 23, 2017) Transfer project to Brad Jasper. Fixed off-by-one array bug. Fixed previous value bug. Added custom annotations. Added ImportJSONFromSheet and ImportJSONBasicAuth.
  1.3.0  Adds ability to import the text from a set of rows containing the text to parse. All cells are concatenated
  1.2.1  Fixed a bug with how nested arrays are handled. The rowIndex counter wasn't incrementing properly when parsing.
  1.2.0  Added ImportJSONViaPost and support for fetchOptions to ImportJSONAdvanced
  1.1.1  Added a version number using Google Scripts Versioning so other developers can use the library
  1.1.0  Added support for the noHeaders option
  1.0.0  Initial release
 *====================================================================================================================================*/

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the updateCachedJson() function.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Update Cached JSON",
    functionName : "updateCachedJson"
  }];
  sheet.addMenu("Script Center Menu", entries);
};

function getCacheHeaderPath_(value) {
  var cacheNames = new Array("cached json", "json cache");
  var valueLower = value.toString().toLowerCase();
  if (cacheNames.indexOf(valueLower) >= 0) {
    return "/";
  }
  for (var cacheName of cacheNames) {
    if (valueLower.indexOf(cacheName) == 0) {
      var path = value.substring(cacheName.length).replace(/^[: ]*/, "");
      return path;
    }
  }
  return null;
}

function isCacheHeader_(value) {
  return (getCacheHeaderPath_(value) !== null)
}

/**
 * Finds all columns titled "cached json", then fetches and writes the json from the API URL in the cell to the left of them.
 * Header title is case insensitive. Possible values are specified in the getCacheHeaderPath_ function.
 * The header title can be suffixed with an xpath, and then only the specified json paths are saved.
 */
function updateCachedJson() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetRange = sheet.getDataRange();
  var numRows = sheetRange.getNumRows();
  var numColumns = sheetRange.getNumColumns();
  var values = sheetRange.getValues();
  var cacheColumns = {};  // Map of column index to json query

  for (var rowNum = 0; rowNum <= numRows - 1; rowNum++) {
    for (var colNum = 0; colNum <= numColumns - 1; colNum++) {
      var value = values[rowNum][colNum];
      // Check for new cache columns
      var cachePath = getCacheHeaderPath_(value);
      if (cachePath !== null) {
        cacheColumns[colNum] = cachePath;
        Logger.log("Found header column: "+colNum.toString()+" with path: "+cachePath);
        continue;
      }

      // If this is a cache column, update it
      if (cacheColumns[colNum] !== undefined) {
        // Get URL, cell to the left
        var urlValue = values[rowNum][colNum - 1];
        // Don't update empty URLs
        if (!urlValue) {
          continue;
        }
        // Get JSON and set it
        Logger.log("Updating cell: "+colNum.toString()+","+rowNum.toString()+" with URL: "+urlValue);
        var jsondata = ImportJSON(urlValue, cacheColumns[colNum], "rawJson,retryFetch,noHeaders,noTruncate");
        sheet.getRange(rowNum + 1, colNum + 1).setValue(jsondata);
        Logger.log("Updated cell: "+jsondata);
      }
    }
  }
}

/**
 * Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
 * a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
 * the JSON feed. The remaining rows contain the data. 
 * 
 * By default, data gets transformed so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *      of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    noHeaders:     Don't include headers, only the data
 *    debugLocation: Prepend each value with the row & column it belongs in
 *    retryFetch:    Retries fetching data from the URL up to ten times
 *    rawJson:       Returns the raw json data, which can then be parsed with =ParseJSON()
 *
 * For example:
 *
 *   =ImportJSON("http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", "/feed/entry/title,/feed/entry/content",
 *               "noInherit,noTruncate,rawHeaders")
 * 
 * @param {url}          the URL to a public JSON feed
 * @param {query}        a comma-separated list of paths to import. Any path starting with one of these paths gets imported.
 * @param {options}      a comma-separated list of options that alter fetching and processing of the data (optional)
 * @customfunction
 *
 * @return a two-dimensional array containing the data, with the first row containing headers
 **/
function ImportJSON(url, query, options) {
  return ImportJSONAdvanced(url, null, query, options, includeXPath_, defaultTransform_);
}

/**
 * Imports a JSON feed with a POST request and returns the results to be inserted into a Google Spreadsheet. The JSON feed is 
 * flattened to create a two-dimensional array. The first row contains the headers, with each column header indicating the path to 
 * that data in the JSON feed. The remaining rows contain the data.
 *
 * To retrieve the JSON, a POST request is sent to the URL and the payload is passed as the content of the request using the content 
 * type "application/x-www-form-urlencoded". If the fetchOptions define a value for "method", "payload" or "contentType", these 
 * values will take precedent. For example, advanced users can use this to make this function pass XML as the payload using a GET 
 * request and a content type of "application/xml; charset=utf-8". For more information on the available fetch options, see
 * https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app . At this time the "headers" option is not supported.
 * 
 * By default, the returned data gets transformed so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *     of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    noHeaders:     Don't include headers, only the data
 *    debugLocation: Prepend each value with the row & column it belongs in
 *    retryFetch:    Retries fetching data from the URL up to ten times
 *    rawJson:       Returns the raw json data, which can then be parsed with =ParseJSON()
 *
 * For example:
 *
 *   =ImportJSONWithPost("http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", "user=bob&apikey=xxxx", 
 *               "validateHttpsCertificates=false", "/feed/entry/title,/feed/entry/content", "noInherit,noTruncate,rawHeaders")
 * 
 * @param {url}          the URL to a public JSON feed
 * @param {payload}      the content to pass with the POST request; usually a URL encoded list of parameters separated by ampersands
 * @param {fetchOptions} a comma-separated list of URLFetchApp key=value options used to retrieve the JSON feed from the URL. Can be used to set content type, (which default to form data)
 * @param {query}        a comma-separated list of paths to import. Any path starting with one of these paths gets imported.
 * @param {options}      a comma-separated list of options that alter fetching and processing of the data (optional)
 * @customfunction
 *
 * @return a two-dimensional array containing the data, with the first row containing headers
 **/
function ImportJSONWithPost(url, payload, fetchOptions, query, options) {
  var postOptions = parseToObject_(fetchOptions);
  
  if (postOptions["method"] == null) {
    postOptions["method"] = "POST";
  }

  if (postOptions["payload"] == null) {
    postOptions["payload"] = payload;
  }

  if (postOptions["contentType"] == null) {
    postOptions["contentType"] = "application/x-www-form-urlencoded";
  }

  convertToBool_(postOptions, "validateHttpsCertificates");
  convertToBool_(postOptions, "useIntranet");
  convertToBool_(postOptions, "followRedirects");
  convertToBool_(postOptions, "muteHttpExceptions");
  
  return ImportJSONAdvanced(url, postOptions, query, options, includeXPath_, defaultTransform_);
}


/**
 * An advanced version of ImportJSON designed to be easily extended by a script. This version cannot be called from within a 
 * spreadsheet.
 *
 * Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
 * a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
 * the JSON feed. The remaining rows contain the data. 
 *
 * The fetchOptions can be used to change how the JSON feed is retrieved. For instance, the "method" and "payload" options can be 
 * set to pass a POST request with post parameters. For more information on the available parameters, see 
 * https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app .
 *
 * Use the include and transformation functions to determine what to include in the import and how to transform the data after it is
 * imported. 
 *
 * For example:
 *
 *   ImportJSONAdvanced(
 *       "http://gdata.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json",
 *       new Object() { "method" : "post", "payload" : "user=bob&apikey=xxxx" },
 *       "/feed/entry",
 *       "",
 *       function (query, path) { return path.indexOf(query) == 0; },
 *       function (data, row, column) { data[row][column] = data[row][column].toString().substr(0, 100); }
 *   )
 *
 * In this example, the import function checks to see if the path to the data being imported starts with the query. The transform 
 * function takes the data and truncates it. For more robust versions of these functions, see the internal code of this library.
 *
 * @param {url}           the URL to a public JSON feed
 * @param {fetchOptions}  an object whose properties are URLFetchApp options used to retrieve the JSON feed from the URL
 * @param {query}         the query passed to the include function
 * @param {options}       a comma-separated list of options that may alter fetching and processing of the data
 * @param {includeFunc}   a function with the signature func(query, path, options) that returns true if the data element at the given path
 *                        should be included or false otherwise. 
 * @param {transformFunc} a function with the signature func(data, row, column, options) where data is a 2-dimensional array of the data 
 *                        and row & column are the current row and column being processed. Any return value is ignored. Note that row 0 
 *                        contains the headers for the data, so test for row==0 to process headers only.
 *
 * @return a two-dimensional array containing the data, with the first row containing headers
 **/
function ImportJSONAdvanced(url, fetchOptions, query, options, includeFunc, transformFunc) {
  var optionsList = new Array();
  if (options) {
    optionsList = options.toString().split(",");
  }
  var resp = FetchUrl_(url, optionsList, fetchOptions);
  var jsondata = resp.getContentText();
  return ParseJSONAdvanced(jsondata, query, options, includeFunc, transformFunc);
}

/**
 * Parses JSON data which is already inside the spreadsheet, uses the same syntax and options as IMPORTJSON
 *
 * By default, data gets transformed so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *      of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    noHeaders:     Don't include headers, only the data
 *    debugLocation: Prepend each value with the row & column it belongs in
 *
 * For example:
 *
 *   =ParseJSON(C9, "/feed/entry/title,/feed/entry/content",
 *               "noInherit,noTruncate,rawHeaders")
 * 
 * @param {jsondata} The json data as a string
 * @param {query}    a comma-separated lists of paths to import. Any path starting with one of these paths gets imported.
 * @param {options}  a comma-separated list of options that alter processing of the data
 * 
 * @return a two-dimensional array containing the data, with the first row containing headers
 * @customfunction
 */
function ParseJSON(jsondata, query, options) {
  return ParseJSONAdvanced(jsondata, query, options, includeXPath_, defaultTransform_);
}

/**
 * Parses a JSON text from a named Sheet and returns the results to be inserted into a Google Spreadsheet.
 * The entire target sheet is concatenated into one json string, which is then parsed and handled as normal.
 * 
 * By default, data gets transformed so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *      of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    noHeaders:     Don't include headers, only the data
 *    debugLocation: Prepend each value with the row & column it belongs in
 *
 * For example:
 *
 *   =ParseJSONFromSheet("Source", "/feed/entry/title,/feed/entry/content",
 *               "noInherit,noTruncate,rawHeaders")
 * 
 * @param {sheetName} the name of the sheet containg the text for the JSON
 * @param {query}     a comma-separated list of paths to import. Any path starting with one of these paths gets imported.
 * @param {options}   a comma-separated list of options that alter processing of the data
 *
 * @return a two-dimensional array containing the data, with the first row containing headers
 * @customfunction
 **/
function ParseJSONFromSheet(sheetName, query, options) {
  var jsondata = getDataFromNamedSheet_(sheetName);
  return ParseJSONAdvanced(jsondata, query, options, includeXPath_, defaultTransform_);
}

function ParseJSONAdvanced(jsondata, query, options, includeFunc, transformFunc) {
  if (options) {
    options = options.toString().split(",");
  }
  var object   = JSON.parse(jsondata);
  
  return parseJSONObject_(object, query, options, includeFunc, transformFunc);
}

/**
 * Gets the json data, handling options and basic auth, and such
 * 
 * @param {url}          The url to fetch
 * @param {options}      List of options for the function
 * @param {fetchOptions} URLFetchApp options (optional)
 * 
 * @return the json data from the URL
 */
function FetchUrl_(url, options, fetchOptions) {
  Logger.log("Fetching url: "+url);
  if (fetchOptions === null) {
    fetchOptions = {};
  }
  var basicAuth = url.match(/^https?:\/\/(.*?:.*?)@/);
  if (basicAuth) {
    fetchOptions.headers = {
      "Authorization": "Basic " + Utilities.base64Encode(basicAuth[1], Utilities.Charset.UTF_8)
    };
    url = url.replace(/^(https?):\/\/.*?:.*?@/, "$1://");
  }
  
  var retries = 0;
  if (hasOption_(options, "retryFetch")) {
    retries = 10;
    fetchOptions.muteHttpExceptions = true;
  }
  var resp = UrlFetchApp.fetch(url, fetchOptions);
  for (var i = 0; i <= retries; i = i + 1) {
    if (resp.getResponseCode() == 200) {
      return resp
    }
    Utilities.sleep(1000);
    Logger.log("Retry URL: "+url);
    resp = UrlFetchApp.fetch(url, fetchOptions);
  }
  return resp;
}

/** 
 * Encodes the given value to use within a URL.
 *
 * @param {value} the value to be encoded
 * 
 * @return the value encoded using URL percent-encoding
 */
function URLEncode(value) {
  return encodeURIComponent(value.toString());  
}

/**
 * Adds an oAuth service using the given name and the list of properties.
 *
 * @note This method is an experiment in trying to figure out how to add an oAuth service without having to specify it on each 
 *       ImportJSON call. The idea was to call this method in the first cell of a spreadsheet, and then use ImportJSON in other
 *       cells. This didn't work, but leaving this in here for further experimentation later. 
 *
 *       The test I did was to add the following into the A1:
 *  
 *           =AddOAuthService("twitter", "https://api.twitter.com/oauth/access_token", 
 *                            "https://api.twitter.com/oauth/request_token", "https://api.twitter.com/oauth/authorize", 
 *                            "<my consumer key>", "<my consumer secret>", "", "")
 *
 *       Information on obtaining a consumer key & secret for Twitter can be found at https://dev.twitter.com/docs/auth/using-oauth
 *
 *       Then I added the following into A2:
 *
 *           =ImportJSONViaPost("https://api.twitter.com/1.1/statuses/user_timeline.json?screen_name=fastfedora&count=2", "",
 *                              "oAuthServiceName=twitter,oAuthUseToken=always", "/", "")
 *
 *       I received an error that the "oAuthServiceName" was not a valid value. [twl 18.Apr.13]
 */
function AddOAuthService__(name, accessTokenUrl, requestTokenUrl, authorizationUrl, consumerKey, consumerSecret, method, paramLocation) {
  var oAuthConfig = UrlFetchApp.addOAuthService(name);

  if (accessTokenUrl != null && accessTokenUrl.length > 0) {
    oAuthConfig.setAccessTokenUrl(accessTokenUrl);
  }
  
  if (requestTokenUrl != null && requestTokenUrl.length > 0) {
    oAuthConfig.setRequestTokenUrl(requestTokenUrl);
  }
  
  if (authorizationUrl != null && authorizationUrl.length > 0) {
    oAuthConfig.setAuthorizationUrl(authorizationUrl);
  }
  
  if (consumerKey != null && consumerKey.length > 0) {
    oAuthConfig.setConsumerKey(consumerKey);
  }
  
  if (consumerSecret != null && consumerSecret.length > 0) {
    oAuthConfig.setConsumerSecret(consumerSecret);
  }
  
  if (method != null && method.length > 0) {
    oAuthConfig.setMethod(method);
  }
  
  if (paramLocation != null && paramLocation.length > 0) {
    oAuthConfig.setParamLocation(paramLocation);
  }
}

/** 
 * Parses a JSON object and returns a two-dimensional array containing the data of that object.
 */
function parseJSONObject_(object, query, options, includeFunc, transformFunc) {
  var headers = {};  // A map of query path to column index
  var data    = new Array();
  
  if (!query) {
    query = "/";
  }
  
  if (query && !Array.isArray(query) && query.toString().indexOf(",") != -1) {
    query = query.toString().split(",");
  }
  
  // Prepopulate the headers map and data headers to lock in their order
  if (Array.isArray(query)) {
    data[0] = new Array();
    for (var i in query) {
      headers[query[i]] = i;
      data[0][i] = query[i];
    }
  } else {
    headers[query] = 0;
    data[0] = new Array(query);
  }
  
  parseData_(headers, data, "", object, query, options, includeFunc);
  transformData_(data, options, transformFunc);

  if (hasOption_(options, "noHeaders") || hasOption_(options, "rawJson")) {
    data = (data.length > 1 ? data.slice(1) : new Array());
  }
  
  return data;
}

/** 
 * Parses the data contained within the given value and inserts it into the data two-dimensional array starting at the rowIndex. 
 * If the data is to be inserted into a new column, a new header is added to the headers array. The value can be an object, 
 * array or scalar value.
 *
 * If the value is an object, it's properties are iterated through and passed back into this function with the name of each 
 * property extending the path. For instance, if the object contains the property "entry" and the path passed in was "/feed",
 * this function is called with the value of the entry property and the path "/feed/entry".
 *
 * If the value is an array containing only scalar values, those values are joined together and inserted into the data array as 
 * a single value.
 *
 * If the value is a scalar, the value is inserted directly into the data array.
 */
function parseData_(headers, data, path, value, query, options, includeFunc) {
  if (!includeFunc || includeFunc(query, path, options) !== false) {
    var colIndex = includeFunc(query, path, options);
    insertValue_(data, colIndex, value);
  }
  
  if (isObject_(value)) {
    for (key in value) {
      var subPath = path + "/" + key;
      parseData_(headers, data, subPath, value[key], query, options, includeFunc);
    }
  } else if (Array.isArray(value)) {
    for (var i = 0; i < value.length; i++) {
      var subPath = path + "[" + i.toString() + "]"
      parseData_(headers, data, subPath, value[i], query, options, includeFunc);
    }
  }
}

function insertValue_(data, colIndex, value) {
  var rowIndex = 1;
  while (true) {
    // Insert new row if one doesn't already exist
    if (!data[rowIndex]) {
      data[rowIndex] = new Array(data[0].length);
    }
    
    if (data[rowIndex][colIndex] === undefined) {
      data[rowIndex][colIndex] = value;
      return;
    } else {
      rowIndex += 1;
    }
  }
}

/** 
 * Applies the transform function for each element in the data array, going through each column of each row.
 */
function transformData_(data, options, transformFunc) {
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[0].length; j++) {
      transformFunc(data, i, j, options);
    }
  }
}

/** 
 * Returns true if the given test value is an object; false otherwise.
 */
function isObject_(test) {
  return Object.prototype.toString.call(test) === '[object Object]';
}

/** 
 * Returns true if the given test value is an array containing at least one object; false otherwise.
 */
function isObjectArray_(test) {
  for (var i = 0; i < test.length; i++) {
    if (isObject_(test[i])) {
      return true; 
    }
  }  

  return false;
}

/** 
 * Returns index of the query if given path matches, otherwise false.
 */
function includeXPath_(query, path, options) {
  if (!query) {
    return true; 
  } else if (Array.isArray(query)) {
    for (var i = 0; i < query.length; i++) {
      if (applyXPathRule_(query[i], path, options)) {
        return i; 
      }
    }  
  } else {
    if (applyXPathRule_(query, path, options)) {
      return 0;
    }
  }
  
  return false; 
}

/**
 * Returns true if the rule applies to the given path.
 */
function applyXPathRule_(rule, path, options) {
  rule = rule.replace(/\/$/, "");
  if (rule == path) {
    return true;
  }
  // Handle list indexing TODO: Don't remove all list indexes, try a couple
  if(path.match(/\[[0-9]+\]/g)) {
    path = path.replace(/\[[0-9]+\]/g, "");
  }
  return rule == path;
}

/** 
 * By default, this function transforms the value at the given row & column so it looks more like a normal data import. Specifically:
 *
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *     of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Values in row 0 (headers) have slashes converted to spaces, common prefixes removed and the resulting text converted to title 
 *     case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from previous element results
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    debugLocation: Prepend each value with the row & column it belongs in
 *    rawJson:       Dump each value as JSON
 */
function defaultTransform_(data, row, column, options) {
  if (data[row][column] == null) {
    if (row < 2 || hasOption_(options, "noInherit")) {
      data[row][column] = "";
    } else {
      data[row][column] = data[row-1][column];
    }
  }

  if (!hasOption_(options, "rawHeaders") && row == 0) {
    if (column == 0 && data[row].length > 1) {
      removeCommonPrefixes_(data, row);  
    }
    data[row][column] = toTitleCase_(data[row][column].toString().replace(/[\/\_]/g, " "));
  }

  if (hasOption_(options, "rawJson") && row > 0 && data[row][column]) {
    data[row][column] = JSON.stringify(data[row][column]);
  }
  
  if (!hasOption_(options, "noTruncate") && data[row][column]) {
    data[row][column] = data[row][column].toString().substr(0, 256);
  }

  if (hasOption_(options, "debugLocation")) {
    data[row][column] = "[" + row + "," + column + "]" + data[row][column];
  }

  if (!isNaN(parseFloat(data[row][column])) && isFinite(data[row][column])) {
    data[row][column] = parseFloat(data[row][column]);
  }
}

/** 
 * If all the values in the given row share the same prefix, remove that prefix.
 */
function removeCommonPrefixes_(data, row) {
  var matchIndex = data[row][0].length;

  for (var i = 1; i < data[row].length; i++) {
    matchIndex = findEqualityEndpoint_(data[row][i-1], data[row][i], matchIndex);

    if (matchIndex == 0) {
      return;
    }
  }
  
  for (var i = 0; i < data[row].length; i++) {
    data[row][i] = data[row][i].substring(matchIndex, data[row][i].length);
  }
}

/** 
 * Locates the index where the two strings values stop being equal, stopping automatically at the stopAt index.
 */
function findEqualityEndpoint_(string1, string2, stopAt) {
  if (!string1 || !string2) {
    return -1; 
  }
  
  var maxEndpoint = Math.min(stopAt, string1.length, string2.length);
  
  for (var i = 0; i < maxEndpoint; i++) {
    if (string1.charAt(i) != string2.charAt(i)) {
      return i;
    }
  }
  
  return maxEndpoint;
}
  

/** 
 * Converts the text to title case.
 */
function toTitleCase_(text) {
  if (text == null) {
    return null;
  }
  
  return text.replace(/\w\S*/g, function(word) { return word.charAt(0).toUpperCase() + word.substr(1).toLowerCase(); });
}

/** 
 * Returns true if the given set of options contains the given option.
 */
function hasOption_(options, option) {
  return options && options.indexOf(option) >= 0;
}

/** 
 * Parses the given string into an object, trimming any leading or trailing spaces from the keys.
 */
function parseToObject_(text) {
  var map     = new Object();
  var entries = (text != null && text.trim().length > 0) ? text.toString().split(",") : new Array();
  
  for (entry of entries) {
    addToMap_(map, entry);
  }
  
  return map;
}

/** 
 * Parses the given entry and adds it to the given map, trimming any leading or trailing spaces from the key.
 */
function addToMap_(map, entry) {
  var equalsIndex = entry.indexOf("=");
  var key         = (equalsIndex != -1) ? entry.substring(0, equalsIndex) : entry;
  var value       = (key.length + 1 < entry.length) ? entry.substring(key.length + 1) : "";
  
  map[key.trim()] = value;
}

/** 
 * Returns the given value as a boolean.
 */
function toBool_(value) {
  if (value === null) {
    return false;
  }
  if (value === "") {  // Support passing flags by name alone
    return true;
  }
  return (value.toString().toLowerCase() == "true");
}

/**
 * Converts the value for the given key in the given map to a bool.
 */
function convertToBool_(map, key) {
  if (map[key] !== null) {
    map[key] = toBool_(map[key]);
  }
}

function getDataFromNamedSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  var jsonRange = sheet.getDataRange();
  var jsonValues = jsonRange.getValues();
  
  var jsonText = "";
  for (var rowNum in jsonValues) {
    for (var colNum in jsonValues[rowNum]) {
      jsonText += jsonValues[rowNum][colNum];
    }
  }
  Logger.log("Constructed json from sheet: "+jsonText);
  return jsonText;
}
