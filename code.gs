/**
 * Facebook Ads Report Generator for Google Sheets
 * 
 * SETUP INSTRUCTIONS:
 * 1. Go to Extensions > Apps Script in your Google Sheet.
 * 2. In the "Libraries" section (left sidebar), add the OAuth2 library:
 *    Script ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 * 3. Replace the Client IDs and Secrets below.
 * 4. Ensure you run logRedirectUri() for BOTH apps to hook them up to FB.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('FB Reports')
    .addItem('1. Authorize Facebook', 'showAuthSidebar')
    .addItem('2. Generate Ads Report', 'showReportSidebar')
    .addItem('3. Logout', 'clearService')
    .addSeparator()
    .addItem('Setup Automation Config', 'setupConfigSheet')
    .addItem('Enable Daily Schedule (6AM)', 'createDailyTrigger')
    .addToUi();
}
/**
 * Configure the OAuth2 Service based on the active user's email domain
 */
function getService() {
  var email = Session.getActiveUser().getEmail();
  var isCorporate = (email.indexOf('@publicisgroupe.net') > -1);
  
  var clientId = isCorporate ? CLIENT_ID_CORPORATE : CLIENT_ID_PUBLIC;
  var clientSecret = isCorporate ? CLIENT_SECRET_CORPORATE : CLIENT_SECRET_PÜBLIC;
  
  // Naming the service slightly differently based on the app to avoid token conflicts
  var serviceName = isCorporate ? 'Facebook_Corporate' : 'Facebook_Public';
  return OAuth2.createService(serviceName)
    .setAuthorizationBaseUrl('https://www.facebook.com/dialog/oauth')
    .setTokenUrl('https://graph.facebook.com/v25.0/oauth/access_token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('ads_management,ads_read'); 
}
function showAuthSidebar() {
  var service = getService();
  
  // Show in the UI which app we are using
  var email = Session.getActiveUser().getEmail();
  var isCorporate = (email.indexOf('@publicisgroupe.net') > -1);
  var appName = isCorporate ? "Corporate App (@publicisgroupe.net)" : "Public Testing App";
  var color = isCorporate ? "green" : "orange";
  
  if (!service.hasAccess()) {
    var authorizationUrl = service.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">' +
        '<div style="padding: 15px;">' +
        '<h4>Facebook Authorization</h4>' +
        '<p style="font-size: 13px; color: #555;">Routing via: <strong style="color: ' + color + '">' + appName + '</strong><br>Detected Email: ' + email + '</p>' +
        '<p>Please authorize this script to access your Facebook Ads data.</p>' +
        '<a class="button blue" href="<?= authorizationUrl ?>" target="_blank">Authorize Facebook</a>' +
        '<p style="margin-top: 10px; color: gray; font-size: 12px;">Reopen this sidebar to check your status after authorizing.</p>' +
        '</div>');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate().setTitle('Facebook Status');
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    SpreadsheetApp.getUi().alert('You are already authorized using the ' + appName + ' Route!');
  }
}
function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab and return to Sheets.');
  } else {
    return HtmlService.createHtmlOutput('Access Denied. You can close this tab.');
  }
}
function clearService() {
  getService().reset();
  SpreadsheetApp.getUi().alert('Logged out from Facebook successfully.');
}
/**
 * Show the main configuration Sidebar for running a report
 */
function showReportSidebar() {
  var service = getService();
  if (!service.hasAccess()) {
    SpreadsheetApp.getUi().alert('Please setup authorization first via: FB Reports -> 1. Authorize Facebook');
    return;
  }
  
  var htmlTemplate = HtmlService.createTemplateFromFile('Sidebar');
  var html = htmlTemplate.evaluate().setTitle('Ads Report Generator');
  SpreadsheetApp.getUi().showSidebar(html);
}
/**
 * Core function that executes the report download from Facebook and pastes it to a sheet
 */
function fetchAndPasteFacebookData(config) {
  var service = getService();
  if (!service.hasAccess()) {
    throw new Error('Not authorized. Please login via: FB Reports -> 1. Authorize Facebook');
  }
  // Formatting Account ID appropriately
  var accountId = config.accountId.trim();
  if (!accountId.startsWith('act_')) {
    accountId = 'act_' + accountId;
  }
  // Requested specific fields
  var fields = config.fields || [
    'account_currency', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
    'adset_id', 'adset_name', 'ad_id', 'ad_name', 'date_start', 'date_stop', 'spend', 'impressions',
    'inline_link_clicks', 'video_p25_watched_actions', 'video_p50_watched_actions',
    'video_p75_watched_actions', 'video_p95_watched_actions', 'video_p100_watched_actions',
    'video_play_actions', 'purchase_roas', 'converted_product_omni_purchase',
    'converted_product_omni_purchase_values', 'converted_product_website_pixel_purchase',
    'converted_product_website_pixel_purchase_value', 'actions'
  ].join(',');
  var limit = 1000;
  var timeParam = (config.timeIncrement === '1') ? '&time_increment=1' : '';
  
  // Conditionally build Date Range string
  var dateStr = '';
  if (config.customTimeRange) {
    dateStr = '&time_range=' + encodeURIComponent(JSON.stringify(config.customTimeRange));
  } else if (config.datePreset) {
    dateStr = '&date_preset=' + config.datePreset;
  }
  
  var url = 'https://graph.facebook.com/V25.0/' + accountId + '/insights' +
            '?level=ad' +
            dateStr +
            timeParam +
            '&fields=' + fields +
            '&limit=' + limit +
            '&access_token=' + service.getAccessToken();
  Logger.log('Fetching Report...');
  Logger.log('URL: ' + url.replace(/access_token=[^&]+/, "access_token=HIDDEN"));
  var allData = [];
  var headersExtracted = false;
  var headers = [];
  // Specific action types to extract
  var actionKeysStr = config.actionBreakdowns || [
    'comment',
    'post_reaction',
    'post_engagement',
    'landing_page_view',
    'add_to_cart',
    'lead',
    'purchase',
    'omni_purchase',
    'offsite_conversion.fb_pixel_custom'
  ].join(',');
  var actionKeys = actionKeysStr ? actionKeysStr.split(',').map(function(s){return s.trim();}) : [];
  var currentUrl = url;
  while (currentUrl) {
    var response = UrlFetchApp.fetch(currentUrl, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    var responseCode = response.getResponseCode();
    var contentText = response.getContentText();
    
    if (responseCode !== 200) {
      try {
        var errorJson = JSON.parse(contentText);
        throw new Error('Graph API Error: ' + (errorJson.error.message || contentText));
      } catch (e) {
        throw new Error('API Error: ' + contentText.substring(0, 250));
      }
    }
    
    var json = JSON.parse(contentText);
    
    if (json.data && json.data.length > 0) {
      if (!headersExtracted) {
        // Build base headers. 
        var baseFields = fields.split(',').filter(function(f) { return f !== 'actions'; });
        headers = baseFields.concat(actionKeys);
        allData.push(headers);
        headersExtracted = true;
        
        // Save the first row's account info so the user has an easier time selecting it later
        var firstRow = json.data[0];
        if (firstRow && firstRow.account_id && firstRow.account_name) {
          saveAccountToProperties(firstRow.account_id, firstRow.account_name);
        }
      }
      
      // Process each row to flatten the API response into columns matching `headers`
      for (var i = 0; i < json.data.length; i++) {
        var row = json.data[i];
        var rowArray = [];
        var baseFields = fields.split(',').filter(function(f) { return f !== 'actions'; });
        
        // Map out the actions array for this row into a dictionary for fast lookup
        var actionDict = {};
        if (row.actions && Array.isArray(row.actions)) {
          for (var k = 0; k < row.actions.length; k++) {
            actionDict[row.actions[k].action_type] = row.actions[k].value;
          }
        }
        // 1. Process base fields (Metrics)
        for (var j = 0; j < baseFields.length; j++) {
          var fieldName = baseFields[j];
          var val = row[fieldName];
          
          if (val === undefined || val === null) {
            rowArray.push('');
          } else if (Array.isArray(val)) {
            // Complex objects like `purchase_roas` or `video_p25_watched_actions` arrive as arrays of dictionaries
            if (val.length === 1 && val[0].value) {
                rowArray.push(val[0].value);
            } else {
                rowArray.push(JSON.stringify(val));
            }
          } else if (typeof val === 'object') {
            rowArray.push(JSON.stringify(val));
          } else {
            rowArray.push(val);
          }
        }
        
        // 2. Process specific action columns extracted from the `actions` array
        for (var l = 0; l < actionKeys.length; l++) {
           var actVal = actionDict[actionKeys[l]];
           rowArray.push(actVal !== undefined ? actVal : '');
        }
        allData.push(rowArray);
      }
    }
    
    // Check for next page of data
    if (json.paging && json.paging.next) {
      currentUrl = json.paging.next;
    } else {
      currentUrl = null; 
    }
  }
  if (allData.length === 0) {
    return { success: true, message: 'Report completed, but no data was returned for this date preset.' };
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = config.targetSheetName || 'Data';
  var sheet = ss.getSheetByName(sheetName);
  
  // Create the target sheet if it doesn't exist yet
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  var numRows = allData.length;
  var numCols = headers.length; // Ensure matrix exactly matches headers length
  
  var range = sheet.getRange(1, 1, numRows, numCols);
  range.setValues(allData);
  
  return { success: true, message: 'Report imported successfully! (' + (numRows - 1) + ' rows)' };
}
/**
 * Called by the Sidebar UI to execute a manual report download
 */
function runReport(config) {
  // Hardcode the manual target sheet name for Sidebar users
  config.targetSheetName = 'Data'; 
  return fetchAndPasteFacebookData(config);
}
/**
 * Returns saved account history from user properties
 */
function getSavedAccounts() {
  var props = PropertiesService.getUserProperties();
  var saved = props.getProperty('savedFBAccounts');
  return saved ? JSON.parse(saved) : [];
}
/**
 * Save an account to user properties so it appears in the dropdown
 */
function saveAccountToProperties(id, name) {
  var props = PropertiesService.getUserProperties();
  var saved = props.getProperty('savedFBAccounts');
  var accounts = saved ? JSON.parse(saved) : [];
  
  // Format to standard ID
  var cleanId = id.replace('act_', '');
  
  var exists = false;
  for (var i = 0; i < accounts.length; i++) {
    if (accounts[i].id === cleanId) {
      accounts[i].name = name; // Update name just in case it mutated
      exists = true;
      break;
    }
  }
  
  if (!exists) {
    accounts.push({id: cleanId, name: name});
    props.setProperty('savedFBAccounts', JSON.stringify(accounts));
  }
}
function logRedirectUri() {
  Logger.log('Your Redirect URI is: ' + OAuth2.getRedirectUri());
}
/**
 * Creates or formats the FB_Config tab used for automated scheduled runs
 */
function setupConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'FB_Config';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  // Headers for multi-row
  var headers = [
    'Enabled', 'Target Sheet Name', 'Ad Account ID', 'Time Breakdown', 
    'Date Mode', 'Date Preset', 'Custom Start Date', 'Custom End Date',
    'Fields (Comma Separated)', 'Action Breakdowns (Comma Separated)'
  ];
  
  var defaultFields = [
    'account_currency', 'account_name', 'account_id', 'campaign_id', 'campaign_name',
    'adset_id', 'adset_name', 'ad_id', 'ad_name', 'date_start', 'date_stop', 'spend', 'impressions',
    'inline_link_clicks', 'video_p25_watched_actions', 'video_p50_watched_actions',
    'video_p75_watched_actions', 'video_p95_watched_actions', 'video_p100_watched_actions',
    'video_play_actions', 'purchase_roas', 'converted_product_omni_purchase',
    'converted_product_omni_purchase_values', 'converted_product_website_pixel_purchase',
    'converted_product_website_pixel_purchase_value', 'actions'
  ].join(',');
  
  var defaultActions = [
    'comment', 'post_reaction', 'post_engagement', 'landing_page_view', 
    'add_to_cart', 'lead', 'purchase', 'omni_purchase', 'offsite_conversion.fb_pixel_custom'
  ].join(',');
  
  var defaultRow = [
    true, 'Data', '123456789', 'Lifetime', 'Preset', 'last_30d', '=TODAY()-30', '=TODAY()', defaultFields, defaultActions
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, 1, defaultRow.length).setValues([defaultRow]);
  
  // Insert Inserted Checkboxes
  sheet.getRange('A2:A100').insertCheckboxes();
  
  // Formatting
  sheet.getRange('A1:J1').setFontWeight('bold').setBackground('#d9ead3');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(9, 300);
  sheet.setColumnWidth(10, 300);
  
  // Data Validation for Dropdowns
  var breakdownRule = SpreadsheetApp.newDataValidation().requireValueInList(['Lifetime', 'Daily']).build();
  sheet.getRange('D2:D100').setDataValidation(breakdownRule);
  
  var modeRule = SpreadsheetApp.newDataValidation().requireValueInList(['Preset', 'Custom']).build();
  sheet.getRange('E2:E100').setDataValidation(modeRule);
  
  SpreadsheetApp.getUi().alert('FB_Config tab created! Add one row per report you want to schedule.');
}
/**
 * Creates a daily Time-Driven trigger to run the report at 6 AM
 */
function createDailyTrigger() {
  // First, clear any existing triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runScheduledReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger
  ScriptApp.newTrigger('runScheduledReport')
           .timeBased()
           .everyDays(1)
           .atHour(6) // Runs around 6 AM to 7 AM depending on timezone
           .create();
           
  SpreadsheetApp.getUi().alert('Success! The report will now automatically read from FB_Config and run every day around 6:00 AM.');
}
/**
 * The function called by the automated trigger. Reads the config and fetches the report.
 */
function runScheduledReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName('FB_Config');
  
  if (!configSheet) {
    throw new Error('Automated run failed: FB_Config tab not found.');
  }
  
  var data = configSheet.getDataRange().getValues();
  // Skip header row
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var isEnabled = row[0];
    
    // Only run if the checkbox is checked and there's a Target Sheet Name
    if (isEnabled === true && row[1]) {
      var config = {
        targetSheetName: row[1].toString(),
        accountId: row[2].toString(),
        timeIncrement: row[3] === 'Daily' ? '1' : 'none',
      };
      
      var dateMode = row[4];
      if (dateMode === 'Preset') {
        config.datePreset = row[5].toString();
      } else {
        // Handling Custom cell ranges (which Google parses as full Date objects if using formulas like =TODAY())
        config.customTimeRange = {
          since: _formatDateForFB(row[6]),
          until: _formatDateForFB(row[7])
        };
      }
      
      // Override default fields/actions if provided
      if (row[8]) config.fields = row[8].toString();
      if (row[9]) config.actionBreakdowns = row[9].toString();
      
      try {
        fetchAndPasteFacebookData(config);
      } catch (e) {
        Logger.log('Error running row ' + (i+1) + ': ' + e.message);
      }
    }
  }
}
/**
 * Helper to parse Google Sheets dates into Facebook 'YYYY-MM-DD' strings
 */
function _formatDateForFB(rawVal) {
  if (rawVal instanceof Date) {
    // It's a real Date object (e.g. they typed =TODAY())
    var y = rawVal.getFullYear();
    var m = ('0' + (rawVal.getMonth() + 1)).slice(-2);
    var d = ('0' + rawVal.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  // Otherwise assume it's already a text string like '2023-01-01'
  return rawVal.toString();
}
