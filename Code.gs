 CONFIG 
const CONFIG = {
  SHEET_NAMES {
    RAW_DATA 'Raw Data',
    DASHBOARD 'SEO Overview',
    QUERY_ROW_COUNT 'Query Row Count Tool',
    URL_QUERY_COUNT 'URL Query Count Tool',
    QUERY_DISCOVERY 'Query Discovery Tool',
    CLICK_POTENTIAL 'Click Potential by CTR',
    QUERY_TRAJECTORY 'Top Level Query Trajectory',
    SETTINGS 'Settings'
  },
  API_LIMIT 25000,
  DEFAULT_DAYS 30,
  OAUTH {
    TOKEN_URL 'httpsoauth2.googleapis.comtoken',
    AUTH_URL 'httpsaccounts.google.comooauth2v2auth',
    SCOPE 'httpswww.googleapis.comauthwebmasters.readonly'
  }
};

 ---------- UI & Initialization ---------- 

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡ SEO Console Analytics')
      .addSubMenu(ui.createMenu('🔧 AUTH & SETUP')
        .addItem('⚙️ Initial Setup', 'setupSheets')
        .addItem('1️⃣ Enter OAuth Credentials', 'showOAuthCredentialsDialog')
        .addItem('2️⃣ View Stored Credentials', 'viewStoredCredentials')
        .addItem('3️⃣ Generate Auth URL', 'generateAuthUrl')
        .addItem('4️⃣ Test Connection', 'testConnection')
        .addItem('❌ Clear All OAuth Data', 'clearAllOAuthData'))
      .addSeparator()
      .addItem('🟢 FETCH & ANALYZE', 'fetchAndAnalyze')
      .addItem('🔄 REFRESH DASHBOARD', 'refreshDashboardButton')
      .addToUi();
  } catch (e) {
     UI unavailable in editor context
  }
}

 ---------- Setup  Sheets ---------- 

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = CONFIG.SHEET_NAMES;
  
   Batch create sheets
  Object.values(names).forEach(name = createSheetIfNotExists(name));
  setupSettingsSheet();
  
  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete!',
    'Sheets created!nnNextn1) Deploy as Web Appn2) Add redirect URI to Google Cloud Consolen3) Enter OAuth credentials',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function createSheetIfNotExists(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName)  ss.insertSheet(sheetName);
}

function setupSettingsSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  sheet.clear();
  
  const scriptUrl = ScriptApp.getService().getUrl()  '';
  const data = [
    ['🔐 OAuth 2.0 Configuration', '', ''],
    ['Setting', 'Value', 'Description'],
    ['Status', 'Not configured', 'OAuth credentials stored in Script Properties'],
    ['Web App URL', scriptUrl  'Deploy as Web App first', 'Copy to Google Cloud Console redirect URIs'],
    ['', '', ''],
    ['📊 Search Console Settings', '', ''],
    ['Property URL', 'httpsexample.com', 'Your Search Console property'],
    ['Start Date', getDateDaysAgo(CONFIG.DEFAULT_DAYS), 'Start date (YYYY-MM-DD)'],
    ['End Date', getDateDaysAgo(0), 'End date (YYYY-MM-DD)'],
    ['Dimensions', 'query,page', 'query, page, country, device (comma-separated)'],
    ['Row Limit', String(CONFIG.API_LIMIT), 'Max rows (API limit 25000)'],
    ['', '', ''],
    ['🔍 Filter Settings (Optional)', '', ''],
    ['Query Contains', '', 'Filter queries containing text'],
    ['Query Excludes', '', 'Exclude queries containing text'],
    ['URL Contains', '', 'Filter URLs containing text']
  ];
  
  sheet.getRange(1, 1, data.length, 3).setValues(data);
  
   Batch formatting
  sheet.getRange('A1C1').merge().setFontSize(14).setFontWeight('bold')
    .setBackground('#4285f4').setFontColor('white').setHorizontalAlignment('center');
  sheet.getRange('A2C2').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.getRange('A6C6').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  sheet.getRange('A13C13').setFontWeight('bold').setBackground('#fbbc04');
  
  sheet.setColumnWidths(1, 3, 200);
  updateOAuthStatus();
}

function fetchAndAnalyze() {
  try {
    fetchSearchConsoleData();
    runFullAnalysis();
    SpreadsheetApp.getUi().alert('✅ Complete!', 'Data fetched and analyzed', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error ' + e.message);
  }
}

function refreshDashboardButton() {
  try {
    refreshDashboard();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Dashboard refresh failedn' + e.message);
  }
}

 ---------- OAuth ---------- 

function showOAuthCredentialsDialog() {
  const props = PropertiesService.getScriptProperties();
  const hasCredentials = !!props.getProperty('OAUTH_CLIENT_ID');
  const autoRedirectUri = ScriptApp.getService().getUrl()  props.getProperty('OAUTH_REDIRECT_URI')  '';
  
  const html = HtmlService.createHtmlOutput(`
    style
      body { font-family Arial; padding 20px; }
      .notice { background #d4edda; padding 12px; border-radius 4px; margin 10px 0; }
      input { width 100%; padding 8px; margin 5px 0 10px; border 2px solid #ddd; border-radius 4px; }
      label { font-weight bold; display block; margin-top 8px; }
      button { background #4285f4; color white; border none; padding 10px 20px; border-radius 4px; cursor pointer; width 100%; }
      buttonhover { background #357ae8; }
      .status { margin-top 10px; padding 8px; border-radius 4px; }
      .success { background #d4edda; color #155724; }
      .error { background #f8d7da; color #721c24; }
    style
    h2🔐 OAuth Credentialsh2
    div class=notice✅ Credentials encrypted in Script Propertiesdiv
    ${hasCredentials  'div class=status success✓ Credentials storeddiv'  ''}
    form onsubmit=saveCredentials(event)
      labelClient IDlabel
      input type=text id=clientId required 
      labelClient Secretlabel
      input type=password id=clientSecret required 
      labelRedirect URIlabel
      input type=text id=redirectUri value=${autoRedirectUri} 
      button type=submit🔒 Save Securelybutton
    form
    div id=statusdiv
    script
      function saveCredentials(e) {
        e.preventDefault();
        const data = {
          clientId document.getElementById('clientId').value.trim(),
          clientSecret document.getElementById('clientSecret').value.trim(),
          redirectUri document.getElementById('redirectUri').value.trim()
        };
        if (!data.clientId  !data.clientSecret  !data.redirectUri) {
          document.getElementById('status').innerHTML = 'div class=errorFill all fieldsdiv';
          return;
        }
        google.script.run
          .withSuccessHandler(m = {
            document.getElementById('status').innerHTML = 'div class=success✅ ' + m + 'div';
            setTimeout(() = google.script.host.close(), 1500);
          })
          .withFailureHandler(e = {
            document.getElementById('status').innerHTML = 'div class=error❌ ' + e.message + 'div';
          })
          .saveOAuthCredentials(data.clientId, data.clientSecret, data.redirectUri);
      }
    script
  `).setWidth(600).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'OAuth Credentials');
}

function saveOAuthCredentials(clientId, clientSecret, redirectUri) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'OAUTH_CLIENT_ID' clientId,
    'OAUTH_CLIENT_SECRET' clientSecret,
    'OAUTH_REDIRECT_URI' redirectUri
  });
  updateOAuthStatus();
  return 'Saved! Generate auth URL next.';
}

function viewStoredCredentials() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('OAUTH_CLIENT_ID');
  const redirectUri = props.getProperty('OAUTH_REDIRECT_URI')  ScriptApp.getService().getUrl()  '';
  
  if (!clientId) {
    SpreadsheetApp.getUi().alert('No credentials found. Enter them first.');
    return;
  }
  
  SpreadsheetApp.getUi().alert('🔐 Stored Credentials',
    `Client ID ${maskString(clientId)}nClient Secret ••••••••nRedirect URI ${redirectUri}`,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function updateOAuthStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  if (!sheet) return;
  
  const props = PropertiesService.getScriptProperties();
  const hasTokens = !!props.getProperty('ACCESS_TOKEN');
  const hasCredentials = !!props.getProperty('OAUTH_CLIENT_ID');
  
  let status = 'Not configured';
  if (hasTokens) status = '✅ Authorized & Ready';
  else if (hasCredentials) status = '⚠️ Need authorization';
  
  sheet.getRange('B3').setValue(status);
}

function maskString(str) {
  return str && str.length = 10  str.substring(0, 15) + '•••' + str.substring(str.length - 10)  '••••••';
}

function generateAuthUrl() {
  const settings = getOAuthSettings();
  if (!settings.clientId  !settings.clientSecret) {
    SpreadsheetApp.getUi().alert('❌ Enter OAuth credentials first');
    return;
  }
  
  const redirectUri = settings.redirectUri  ScriptApp.getService().getUrl()  '';
  const authUrl = CONFIG.OAUTH.AUTH_URL + '' +
    'client_id=' + encodeURIComponent(settings.clientId) +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&scope=' + encodeURIComponent(CONFIG.OAUTH.SCOPE) +
    '&response_type=code&access_type=offline&prompt=consent';
  
  const html = HtmlService.createHtmlOutput(`
    stylebody{font-familyArial;padding20px} .url{background#f4f4f4;padding12px;border-radius4px;word-breakbreak-all;margin10px 0;border2px solid #4285f4;} button{background#4285f4;color#fff;padding8px 16px;border-radius4px;bordernone;cursorpointer;margin5px;} buttonhover{background#357ae8}style
    h2🔗 Authorization URLh2
    div class=urlsmall${authUrl}smalldiv
    button onclick=window.open('${authUrl}','_blank')🔓 Authorizebutton
    button onclick=navigator.clipboard.writeText('${authUrl}');alert('Copied!')📋 Copybutton
  `).setWidth(700).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Authorization');
}

function exchangeAuthCode(code) {
  const settings = getOAuthSettings();
  const payload = {
    code code,
    client_id settings.clientId,
    client_secret settings.clientSecret,
    redirect_uri settings.redirectUri  ScriptApp.getService().getUrl(),
    grant_type 'authorization_code'
  };
  
  const response = UrlFetchApp.fetch(CONFIG.OAUTH.TOKEN_URL, {
    method 'post',
    payload payload,
    muteHttpExceptions true
  });
  
  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error(result.error_description  result.error);
  
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'ACCESS_TOKEN' result.access_token,
    'REFRESH_TOKEN' result.refresh_token  props.getProperty('REFRESH_TOKEN'),
    'TOKEN_EXPIRY' String(Date.now() + (result.expires_in  1000))
  });
  updateOAuthStatus();
  return 'Successfully authenticated!';
}

function getOAuthSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    clientId props.getProperty('OAUTH_CLIENT_ID')  '',
    clientSecret props.getProperty('OAUTH_CLIENT_SECRET')  '',
    redirectUri props.getProperty('OAUTH_REDIRECT_URI')  ScriptApp.getService().getUrl()  ''
  };
}

function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const accessToken = props.getProperty('ACCESS_TOKEN');
  const tokenExpiry = props.getProperty('TOKEN_EXPIRY');
  const refreshToken = props.getProperty('REFRESH_TOKEN');
  
  if (!accessToken  !tokenExpiry  Date.now() = parseInt(tokenExpiry, 10)) {
    if (!refreshToken) throw new Error('No valid token. Re-authorize.');
    return refreshAccessToken(refreshToken);
  }
  return accessToken;
}

function refreshAccessToken(refreshToken) {
  const settings = getOAuthSettings();
  const response = UrlFetchApp.fetch(CONFIG.OAUTH.TOKEN_URL, {
    method 'post',
    payload {
      refresh_token refreshToken,
      client_id settings.clientId,
      client_secret settings.clientSecret,
      grant_type 'refresh_token'
    },
    muteHttpExceptions true
  });
  
  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error(result.error_description  result.error);
  
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'ACCESS_TOKEN' result.access_token,
    'TOKEN_EXPIRY' String(Date.now() + (result.expires_in  1000))
  });
  return result.access_token;
}

function testConnection() {
  try {
    const accessToken = getAccessToken();
    const settings = getSettings();
    const url = `httpswww.googleapis.comwebmastersv3sites${encodeURIComponent(settings.propertyUrl)}`;
    
    const response = UrlFetchApp.fetch(url, {
      method 'get',
      headers { 'Authorization' 'Bearer ' + accessToken },
      muteHttpExceptions true
    });
    
    if (response.getResponseCode() === 200) {
      SpreadsheetApp.getUi().alert('✅ Connected ton' + settings.propertyUrl);
    } else {
      throw new Error(JSON.parse(response.getContentText()).error.message  'Connection failed');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Connection Failedn' + error.message);
  }
}

function clearAllOAuthData() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Clear ALL OAuth Data', 'This removes credentials and tokens. Continue', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    const props = PropertiesService.getScriptProperties();
    ['OAUTH_CLIENT_ID','OAUTH_CLIENT_SECRET','OAUTH_REDIRECT_URI','ACCESS_TOKEN','REFRESH_TOKEN','TOKEN_EXPIRY']
      .forEach(k = props.deleteProperty(k));
    updateOAuthStatus();
    ui.alert('✅ Cleared');
  }
}

 ---------- Settings & Data Fetching ---------- 

function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.SETTINGS);
  
   Read all settings data at once
  const allData = sheet.getDataRange().getValues();
  
   Find row indices by looking for labels in column A
  const findRow = (label) = {
    for (let i = 0; i  allData.length; i++) {
      if (allData[i][0] && allData[i][0].toString().includes(label)) {
        return i;
      }
    }
    return -1;
  };
  
  const propertyRow = findRow('Property URL');
  const startRow = findRow('Start Date');
  const endRow = findRow('End Date');
  const dimensionsRow = findRow('Dimensions');
  const limitRow = findRow('Row Limit');
  const queryContainsRow = findRow('Query Contains');
  const queryExcludesRow = findRow('Query Excludes');
  const urlContainsRow = findRow('URL Contains');
  
  return {
    propertyUrl propertyRow = 0  (allData[propertyRow][1]  '')  '',
    startDate startRow = 0  (allData[startRow][1]  getDateDaysAgo(CONFIG.DEFAULT_DAYS))  getDateDaysAgo(CONFIG.DEFAULT_DAYS),
    endDate endRow = 0  (allData[endRow][1]  getDateDaysAgo(0))  getDateDaysAgo(0),
    dimensions dimensionsRow = 0  (allData[dimensionsRow][1]  'query,page')  'query,page',
    rowLimit limitRow = 0  (parseInt(allData[limitRow][1], 10)  CONFIG.API_LIMIT)  CONFIG.API_LIMIT,
    queryContains queryContainsRow = 0  (allData[queryContainsRow][1]  '')  '',
    queryExcludes queryExcludesRow = 0  (allData[queryExcludesRow][1]  '')  '',
    urlContains urlContainsRow = 0  (allData[urlContainsRow][1]  '')  ''
  };
}

function formatDateValue(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value).split('T')[0].trim();
}

function fetchSearchConsoleData() {
  const settings = getSettings();
  const accessToken = getAccessToken();
  
  let dimensions = settings.dimensions.split(',').map(d = d.trim()).filter(Boolean);
  if (!dimensions.includes('date')) dimensions.push('date');
  
  const requestBody = {
    startDate formatDateValue(settings.startDate),
    endDate formatDateValue(settings.endDate),
    dimensions dimensions,
    rowLimit settings.rowLimit,
    dataState 'final'
  };
  
  const url = `httpswww.googleapis.comwebmastersv3sites${encodeURIComponent(settings.propertyUrl)}searchAnalyticsquery`;
  const response = UrlFetchApp.fetch(url, {
    method 'post',
    headers {
      'Authorization' 'Bearer ' + accessToken,
      'Content-Type' 'applicationjson'
    },
    payload JSON.stringify(requestBody),
    muteHttpExceptions true
  });
  
  const result = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) {
    throw new Error(result.error.message  'API request failed');
  }
  if (!result.rows  result.rows.length === 0) {
    throw new Error('No data returned. Check settings.');
  }
  
  writeRawData(result.rows, dimensions);
  refreshDashboard();
  return { rows result.rows.length };
}

function fetchSearchConsoleAggregate(propertyUrl, startDate, endDate, accessToken, dimensions) {
  const url = `httpswww.googleapis.comwebmastersv3sites${encodeURIComponent(propertyUrl)}searchAnalyticsquery`;
  const response = UrlFetchApp.fetch(url, {
    method 'post',
    headers {
      'Authorization' 'Bearer ' + accessToken,
      'Content-Type' 'applicationjson'
    },
    payload JSON.stringify({
      startDate,
      endDate,
      dimensions,
      rowLimit 25000,
      dataState 'final'
    }),
    muteHttpExceptions true
  });
  
  const result = JSON.parse(response.getContentText());
  return (result.rows  []).map(r = ({
    keys r.keys,
    clicks r.clicks  0,
    impressions r.impressions  0,
    ctr r.ctr  0,
    position r.position  0
  }));
}

function writeRawData(rows, dimensions) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.RAW_DATA);
  sheet.clear();
  
  const headers = [...dimensions, 'Clicks', 'Impressions', 'CTR', 'Position'];
  const data = rows.map(row = [
    ...row.keys  [],
    row.clicks  0,
    row.impressions  0,
    row.ctr  0,
    row.position  0
  ]);
  
   Batch write
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  if (data.length  0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
     Format percentages and decimals
    const dimsLen = dimensions.length;
    sheet.getRange(2, dimsLen + 3, data.length, 1).setNumberFormat('0.00%');
    sheet.getRange(2, dimsLen + 4, data.length, 1).setNumberFormat('0.0');
  }
  
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 360);
  sheet.setColumnWidth(2, 360);
}

 ---------- Analysis Suite (Fast Mode) ---------- 

function runFullAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RAW_DATA);
  
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert('❌ No Raw Data. Fetch first.');
    return;
  }
  
  const data = rawSheet.getDataRange().getValues();
  if (data.length = 1) {
    SpreadsheetApp.getUi().alert('⚠️ Raw Data empty. Fetch first.');
    return;
  }
  
   Run all analyses in parallel
  const results = {
    queryDiscovery queryDiscoveryFast(data),
    queryRowCount queryRowCountFast(data),
    urlQueryCount urlQueryCountFast(data),
    ctrPotential clickPotentialByCTRFast(data),
    queryTrajectory topLevelQueryTrajectoryFast(data)
  };
  
   Batch write all outputs
  writeAnalysisOutput(CONFIG.SHEET_NAMES.QUERY_DISCOVERY, results.queryDiscovery);
  writeAnalysisOutput(CONFIG.SHEET_NAMES.QUERY_ROW_COUNT, results.queryRowCount);
  writeAnalysisOutput(CONFIG.SHEET_NAMES.URL_QUERY_COUNT, results.urlQueryCount);
  writeAnalysisOutput(CONFIG.SHEET_NAMES.CLICK_POTENTIAL, results.ctrPotential);
  writeAnalysisOutput(CONFIG.SHEET_NAMES.QUERY_TRAJECTORY, results.queryTrajectory);
}

function writeAnalysisOutput(sheetName, output) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  
  sheet.clear();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.getRange('A1Z1').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
}

 ---------- Fast Analysis Functions ---------- 

function queryDiscoveryFast(data) {
  const headers = data[0].map(h = h.toString().toLowerCase().trim());
  const qIdx = headers.indexOf('query');
  const cIdx = headers.indexOf('clicks');
  const iIdx = headers.indexOf('impressions');
  const ctrIdx = headers.indexOf('ctr');
  const pIdx = headers.indexOf('position');
  
  if (qIdx === -1) return [['Error query column not found']];
  
  const queryMap = {};
  data.slice(1).forEach(row = {
    const q = row[qIdx];
    if (!q) return;
    if (!queryMap[q]) queryMap[q] = { clicks 0, impressions 0, ctrSum 0, posSum 0, count 0 };
    queryMap[q].clicks += Number(row[cIdx])  0;
    queryMap[q].impressions += Number(row[iIdx])  0;
    queryMap[q].ctrSum += Number(row[ctrIdx])  0;
    queryMap[q].posSum += Number(row[pIdx])  0;
    queryMap[q].count++;
  });
  
  const output = [['Query', 'Clicks', 'Impressions', 'Avg CTR', 'Avg Position', 'Insight']];
  Object.entries(queryMap)
    .sort((a, b) = b[1].clicks - a[1].clicks)
    .forEach(([q, m]) = {
      const avgCtr = (m.ctrSum  m.count)  100;
      const avgPos = m.posSum  m.count;
      let insight = '';
      
      if (m.impressions  100 && m.clicks  2 && avgPos  20) {
        insight = '🌱 Emerging query';
      } else if (avgCtr  1 && m.impressions  200) {
        insight = '🧠 Low CTR opportunity';
      } else if (avgPos  15 && m.impressions  100) {
        insight = '🧩 Ranking opportunity';
      } else if (m.clicks  10 && avgPos = 5 && avgCtr = 2) {
        insight = '⭐ Strong performer';
      }
      
      output.push([q, m.clicks, m.impressions, avgCtr.toFixed(2) + '%', avgPos.toFixed(1), insight]);
    });
  
  return output;
}

function queryRowCountFast(data) {
  const headers = data[0];
  const qIdx = headers.indexOf('query');
  const cIdx = headers.indexOf('Clicks');
  const iIdx = headers.indexOf('Impressions');
  const pIdx = headers.indexOf('Position');
  const pageIdx = headers.indexOf('page');
  
  if (qIdx  0) return [['Error query column not found']];
  
  const stats = {};
  data.slice(1).forEach(r = {
    const q = r[qIdx]  '';
    if (!stats[q]) stats[q] = { clicks 0, impressions 0, positions [], urls new Set() };
    stats[q].clicks += Number(r[cIdx])  0;
    stats[q].impressions += Number(r[iIdx])  0;
    stats[q].positions.push(Number(r[pIdx])  0);
    if (pageIdx = 0) stats[q].urls.add(r[pageIdx]);
  });
  
  const output = [['Query', 'Clicks', 'Impressions', 'CTR %', 'Avg Position', 'URL Count']];
  Object.entries(stats)
    .sort((a, b) = b[1].clicks - a[1].clicks)
    .forEach(([q, s]) = {
      const ctr = s.impressions  0  (s.clicks  s.impressions  100)  0;
      const avgPos = s.positions.length  (s.positions.reduce((a, b) = a + b, 0)  s.positions.length)  0;
      output.push([q, s.clicks, s.impressions, ctr.toFixed(2), avgPos.toFixed(1), s.urls.size]);
    });
  
  return output;
}

function urlQueryCountFast(data) {
  const headers = data[0];
  const pageIdx = headers.indexOf('page');
  const queryIdx = headers.indexOf('query');
  const clicksIdx = headers.indexOf('Clicks');
  const impressionsIdx = headers.indexOf('Impressions');
  
  if (pageIdx  0) return [['Error page column not found']];
  
  const stats = {};
  data.slice(1).forEach(r = {
    const p = r[pageIdx]  '';
    const q = queryIdx = 0  r[queryIdx]  '';
    if (!stats[p]) stats[p] = { queries new Set(), clicks 0, impressions 0 };
    if (q) stats[p].queries.add(q);
    stats[p].clicks += Number(r[clicksIdx])  0;
    stats[p].impressions += Number(r[impressionsIdx])  0;
  });
  
  const output = [['URL', 'Unique Queries', 'Clicks', 'Impressions', 'CTR %']];
  Object.entries(stats)
    .sort((a, b) = b[1].queries.size - a[1].queries.size)
    .forEach(([url, s]) = {
      const ctr = s.impressions  (s.clicks  s.impressions  100)  0;
      output.push([url, s.queries.size, s.clicks, s.impressions, ctr.toFixed(2)]);
    });
  
  return output;
}

function clickPotentialByCTRFast(data) {
  const headers = data[0];
  const queryIdx = headers.indexOf('query');
  const pageIdx = headers.indexOf('page');
  const clicksIdx = headers.indexOf('Clicks');
  const impressionsIdx = headers.indexOf('Impressions');
  const positionIdx = headers.indexOf('Position');
  
  const agg = {};
  data.slice(1).forEach(r = {
    const key = (r[queryIdx]  '') + '' + (r[pageIdx]  '');
    if (!agg[key]) agg[key] = { query r[queryIdx]  '', page r[pageIdx]  '', clicks 0, imps 0, posList [] };
    agg[key].clicks += Number(r[clicksIdx])  0;
    agg[key].imps += Number(r[impressionsIdx])  0;
    const p = Number(r[positionIdx])  0;
    if (p) agg[key].posList.push(p);
  });
  
  const opportunities = Object.values(agg).map(item = {
    const ctr = item.imps  (item.clicks  item.imps  100)  0;
    const avgPos = item.posList.length  (item.posList.reduce((a, b) = a + b, 0)  item.posList.length)  0;
    return { query item.query, page item.page, clicks item.clicks, imps item.imps, ctr ctr, avgPos avgPos };
  });
  
  opportunities.sort((a, b) = (b.imps  (1 - b.ctr  100)) - (a.imps  (1 - a.ctr  100)));
  
  const output = [['Query', 'Page', 'Impressions', 'Clicks', 'CTR %', 'Avg Position']];
  opportunities.slice(0, 200).forEach(o = {
    output.push([o.query, o.page, o.imps, o.clicks, o.ctr.toFixed(2), o.avgPos.toFixed(1)]);
  });
  
  return output;
}

function topLevelQueryTrajectoryFast(data) {
  const headers = data[0];
  const dateIdx = headers.indexOf('date');
  const pageIdx = headers.indexOf('page');
  const queryIdx = headers.indexOf('query');
  
  if (dateIdx  0) {
    return [['No date dimension found. Include date in dimensions to enable trajectory analysis.']];
  }
  
  const agg = {};
  data.slice(1).forEach(r = {
    const dateStr = r[dateIdx];
    let dd = dateStr;
    if (dateStr instanceof Date) dd = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!dd) return;
    
    const d = new Date(dd);
    if (isNaN(d.getTime())) return;
    
    const weekStart = getWeekStart(d);
    const page = pageIdx = 0  r[pageIdx]  '';
    const query = queryIdx = 0  r[queryIdx]  '';
    
    if (!agg[page]) agg[page] = {};
    if (!agg[page][weekStart]) agg[page][weekStart] = new Set();
    if (query) agg[page][weekStart].add(query);
  });
  
  const weeksSet = new Set();
  Object.values(agg).forEach(pageObj = Object.keys(pageObj).forEach(w = weeksSet.add(w)));
  const weeks = Array.from(weeksSet).sort();
  
  const output = [['Page', 'Total Unique Queries', ...weeks]];
  Object.entries(agg).forEach(([page, weekMap]) = {
    const totalUnique = new Set();
    weeks.forEach(w = (weekMap[w]  new Set()).forEach(q = totalUnique.add(q)));
    const row = [page, totalUnique.size];
    weeks.forEach(w = row.push((weekMap[w]  new Set()).size));
    output.push(row);
  });
  
  return output;
}

function getWeekStart(d) {
  const dt = new Date(d);
  const day = dt.getDay();
  const diff = dt.getDate() - day + (day === 0  -6  1);
  const monday = new Date(dt.setDate(diff));
  return Utilities.formatDate(new Date(monday.getFullYear(), monday.getMonth(), monday.getDate()), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

 ---------- Dashboard ---------- 

function refreshDashboard() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  sheet.clear();
  
  const settings = getSettings();
  const accessToken = getAccessToken();
  const startDate = formatDateValue(settings.startDate);
  const endDate = formatDateValue(settings.endDate);
  const propertyUrl = settings.propertyUrl;
  
   Fetch aggregated data
  const pageData = fetchSearchConsoleAggregate(propertyUrl, startDate, endDate, accessToken, ['page']);
  const queryData = fetchSearchConsoleAggregate(propertyUrl, startDate, endDate, accessToken, ['query']);
  
   Calculate metrics
  const totalClicksByPage = pageData.reduce((sum, row) = sum + row.clicks, 0);
  const totalImprByPage = pageData.reduce((sum, row) = sum + row.impressions, 0);
  const avgCtr = totalImprByPage  0  (totalClicksByPage  totalImprByPage)  100  0;
  const avgPos = pageData.length  0  pageData.reduce((sum, row) = sum + row.position, 0)  pageData.length  0;
  const totalClicksByQuery = queryData.reduce((sum, row) = sum + row.clicks, 0);
  const uniqueQueries = queryData.length;
  const uniquePages = pageData.length;
  
  const overview = [
    ['📊 SEO Overview', '', ''],
    ['Metric', 'Value', 'Description'],
    ['🧩 Total Clicks (by Page)', totalClicksByPage.toLocaleString(), 'Accurate total clicks'],
    ['🔍 Total Clicks (by Query)', totalClicksByQuery.toLocaleString(), 'Sampled query clicks'],
    ['📈 Total Impressions', totalImprByPage.toLocaleString(), 'Total impressions'],
    ['📊 Average CTR', avgCtr.toFixed(2) + '%', 'Clicks ÷ Impressions'],
    ['📍 Average Position', avgPos.toFixed(1), 'Average ranking position'],
    ['🧠 Unique Queries', uniqueQueries.toLocaleString(), 'Distinct queries'],
    ['🌐 Unique Pages', uniquePages.toLocaleString(), 'Distinct pages'],
    ['🕒 Data Range', `${startDate} → ${endDate}`, 'Date range']
  ];
  
  sheet.getRange(1, 1, overview.length, 3).setValues(overview);
  sheet.getRange('A1C1').merge().setFontSize(16).setFontWeight('bold')
    .setBackground('#4285f4').setFontColor('white').setHorizontalAlignment('center');
  sheet.getRange('A2C2').setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setColumnWidths(1, 3, 200);
}

 ---------- Utilities ---------- 

function getDateDaysAgo(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function doGet(e) {
  if (e && e.parameter && e.parameter.code) {
    try {
      exchangeAuthCode(e.parameter.code);
      return HtmlService.createHtmlOutput(`
        div style=font-familyArial;padding40px;text-aligncenter
          h2 style=color#34a853✅ Authorization Successful!h2
          pClose this window and return to your spreadsheet.p
        div
      `);
    } catch (error) {
      return HtmlService.createHtmlOutput(`
        div style=font-familyArial;padding40px;text-aligncenter
          h2 style=color#ea4335❌ Authorization Failedh2
          p${error.message}p
        div
      `);
    }
  }
  return HtmlService.createHtmlOutput(`
    div style=font-familyArial;padding40px;text-aligncenter
      h2⚡ Search Console Analyticsh2
      pStart authorization from the spreadsheet menu.p
    div
  `);
}