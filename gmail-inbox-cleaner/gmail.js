/**
 * Resumable Gmail inbox scan with frequent checkpointing.
 * 
 * Key design:
 * - Checkpoints every CHECKPOINT_THREADS or CHECKPOINT_INTERVAL_MS (whichever first)
 * - Each checkpoint: merge to sheet, update index, save cursor
 * - Cursor always reflects what's persisted — no lost work on timeout
 * - Uses hidden index sheet for O(1) sender->row lookups
 */

const CONFIG = {
  SHEET_NAME: 'Sender Report',
  INDEX_SHEET_NAME: '_SenderIndex',   // hidden helper sheet
  INBOX_QUERY: 'in:inbox',
  
  // Fetching
  PAGE_SIZE_THREADS: 50,              // threads per Gmail API call
  
  // Checkpointing (save progress frequently)
  CHECKPOINT_THREADS: 100,            // checkpoint after this many threads
  CHECKPOINT_INTERVAL_MS: 60 * 1000,  // or after this much time (1 min)
  
  // Hard stop before Apps Script kills us
  RUNTIME_HARD_LIMIT_MS: 5 * 60 * 1000,  // stop at 5 min (Apps Script limit is 6)
  
  DRY_RUN: false,
  TIMEZONE: Session.getScriptTimeZone() || 'Etc/UTC',
  
  // For web app deployment: set this to your spreadsheet ID
  // You can find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  // Leave empty to auto-detect (only works in bound scripts, not deployed web apps)
  SPREADSHEET_ID: '',
};

/**
 * Get the spreadsheet - works in both bound script and web app contexts
 * Creates a new spreadsheet if none exists
 * Uses UserProperties so each user gets their own spreadsheet
 */
function getSpreadsheet() {
  const props = getUserProps();
  const logs = [];
  let ss = null;
  
  // In web app mode, we should ONLY use per-user spreadsheets stored in UserProperties
  // Do NOT use getActiveSpreadsheet() as it returns the bound spreadsheet (owner's)
  // which other users can't access
  
  // Check if user has a saved spreadsheet ID
  const savedId = props.getProperty(KEYS.SPREADSHEET_ID);
  logs.push('UserProps savedId: ' + (savedId || 'null'));
  
  if (savedId) {
    try {
      ss = SpreadsheetApp.openById(savedId);
      // Verify we can actually access it
      const name = ss.getName();
      logs.push('Opened saved spreadsheet: ' + name);
      console.log('getSpreadsheet logs: ' + logs.join(' | '));
      return ss;
    } catch (e) {
      logs.push('Could not open saved spreadsheet: ' + e.message);
      // Clear invalid ID
      props.deleteProperty(KEYS.SPREADSHEET_ID);
      ss = null;
    }
  }
  
  // No saved spreadsheet - create a new one for this user
  logs.push('Creating new spreadsheet...');
  try {
    const userEmail = Session.getEffectiveUser().getEmail() || 'User';
    const sheetName = 'Gmail Inbox Cleaner - ' + userEmail.split('@')[0] + ' - ' + new Date().toLocaleDateString();
    
    logs.push('Creating for user: ' + userEmail);
    ss = SpreadsheetApp.create(sheetName);
    
    if (!ss) {
      throw new Error('SpreadsheetApp.create returned null');
    }
    
    const newId = ss.getId();
    logs.push('Created spreadsheet ID: ' + newId);
    
    // Verify we can access the newly created spreadsheet
    const verifyName = ss.getName();
    logs.push('Verified access: ' + verifyName);
    
    // Save the ID to UserProperties
    props.setProperty(KEYS.SPREADSHEET_ID, newId);
    logs.push('Saved ID to UserProperties');
    
    // Set up the initial sheets
    const sheet = ss.getActiveSheet();
    sheet.setName(CONFIG.SHEET_NAME);
    setupHeaderIfEmpty(sheet);
    
    let indexSheet = ss.insertSheet(CONFIG.INDEX_SHEET_NAME);
    indexSheet.hideSheet();
    resetIndexSheetIfEmpty(indexSheet);
    
    logs.push('Spreadsheet setup complete. URL: ' + ss.getUrl());
    console.log('getSpreadsheet logs: ' + logs.join(' | '));
    return ss;
    
  } catch (e) {
    logs.push('CREATE FAILED: ' + e.message);
    console.error('getSpreadsheet logs: ' + logs.join(' | '));
    throw new Error('Could not create spreadsheet: ' + e.message + ' [Logs: ' + logs.join('; ') + ']');
  }
}

/**
 * Set up header only if sheet is empty (safe version)
 */
function setupHeaderIfEmpty(sheet) {
  if (sheet.getLastRow() > 0) return;
  
  const headers = [
    'Act?', 'Sender Email', 'Sender Name', 'Count', 'Example Subject',
    'Last Seen (ISO)', 'Unsubscribe URL', 'Unsubscribe Mailto', 'Notes / Status', 'Processed'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange('D:D').setNumberFormat('0');
  sheet.getRange('F:F').setNumberFormat('@');
  sheet.getRange('J:J').setNumberFormat('@');
}

/**
 * Set up index sheet only if empty
 */
function resetIndexSheetIfEmpty(indexSheet) {
  if (indexSheet.getLastRow() > 0) return;
  
  indexSheet.getRange(1, 1, 1, 2).setValues([['email_lc', 'row']]);
  indexSheet.setFrozenRows(1);
}

/**
 * Save the spreadsheet ID for web app use (run this once from the spreadsheet)
 */
function saveSpreadsheetId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    getUserProps().setProperty(KEYS.SPREADSHEET_ID, ss.getId());
    safeAlert('Spreadsheet ID saved: ' + ss.getId() + '\n\nThe web app will now work correctly.');
  } else {
    safeAlert('Error: No active spreadsheet. Run this from a spreadsheet.');
  }
}

/**
 * Get the URL to the spreadsheet (for UI)
 */
function getSpreadsheetUrl() {
  const ss = getSpreadsheet();
  return ss ? ss.getUrl() : null;
}

const KEYS = {
  CURSOR: 'cursor',           // thread offset
  INITIALIZED: 'initialized',
  TOTAL_PROCESSED: 'totalProcessed',  // running total across all runs
  AUTO_SCAN_ENABLED: 'autoScanEnabled',
  SPREADSHEET_ID: 'spreadsheetId',    // per-user spreadsheet
};

/**
 * Get properties service - uses UserProperties for per-user data in multi-user mode
 * This ensures each user gets their own spreadsheet and scan state
 */
function getUserProps() {
  return PropertiesService.getUserProperties();
}

/**
 * Get script-level properties (for app-wide settings only)
 */
function getScriptProps() {
  return PropertiesService.getScriptProperties();
}

/**
 * Check if the user has authorized the app
 */
function checkAuthStatus() {
  let email = '';
  let gmailOk = false;
  let errors = [];
  
  try {
    email = Session.getEffectiveUser().getEmail() || '';
  } catch (e) {
    errors.push('getEffectiveUser: ' + e.message);
  }
  
  try {
    GmailApp.getAliases();
    gmailOk = true;
  } catch (e) {
    errors.push('Gmail: ' + e.message);
  }
  
  // Consider authorized if we can get email and Gmail works
  const authorized = !!email && gmailOk;
  
  return {
    authorized: authorized,
    email: email || 'Unknown',
    gmailOk: gmailOk,
    errors: errors,
  };
}

/**
 * Detailed debug info - call this to troubleshoot
 */
function getDebugInfo() {
  const result = {
    timestamp: new Date().toISOString(),
    auth: {},
    session: {},
    properties: {},
    spreadsheet: {},
    getSpreadsheetTest: {},
  };
  
  // Session info
  try {
    result.session.effectiveUser = Session.getEffectiveUser().getEmail() || '(empty)';
  } catch (e) {
    result.session.effectiveUser = 'ERROR: ' + e.message;
  }
  
  try {
    result.session.activeUser = Session.getActiveUser().getEmail() || '(empty)';
  } catch (e) {
    result.session.activeUser = 'ERROR: ' + e.message;
  }
  
  // Simple auth check (don't create test spreadsheet)
  try {
    GmailApp.getAliases();
    result.auth.gmailOk = true;
  } catch (e) {
    result.auth.gmailOk = false;
    result.auth.gmailError = e.message;
  }
  
  // Properties
  try {
    result.properties.user = getUserProps().getProperties();
  } catch (e) {
    result.properties.user = 'ERROR: ' + e.message;
  }
  
  try {
    result.properties.script = getScriptProps().getProperties();
  } catch (e) {
    result.properties.script = 'ERROR: ' + e.message;
  }
  
  // Spreadsheet ID from various sources
  const userSsId = (typeof result.properties.user === 'object') ? result.properties.user[KEYS.SPREADSHEET_ID] : null;
  const scriptSsId = (typeof result.properties.script === 'object') ? result.properties.script['SPREADSHEET_ID'] : null;
  const configSsId = CONFIG.SPREADSHEET_ID;
  
  result.spreadsheet.userPropsId = userSsId || '(none)';
  result.spreadsheet.scriptPropsId = scriptSsId || '(none)';
  result.spreadsheet.configId = configSsId || '(none)';
  
  // Try to actually call getSpreadsheet
  try {
    const ss = getSpreadsheet();
    result.getSpreadsheetTest.success = true;
    result.getSpreadsheetTest.id = ss.getId();
    result.getSpreadsheetTest.name = ss.getName();
    result.getSpreadsheetTest.url = ss.getUrl();
  } catch (e) {
    result.getSpreadsheetTest.success = false;
    result.getSpreadsheetTest.error = e.message;
  }
  
  return result;
}

/**
 * Get the authorization URL for the user to grant permissions
 */
function getAuthUrl() {
  // This triggers the OAuth flow when called
  return ScriptApp.getService().getUrl();
}

/**
 * Force trigger authorization by accessing protected services
 * Call this to initiate the OAuth prompt
 */
function triggerAuth() {
  // These calls will trigger OAuth consent if not already authorized
  const email = Session.getActiveUser().getEmail();
  GmailApp.getAliases(); // Triggers Gmail scope
  SpreadsheetApp.create('Auth Test - Delete Me').setTrashed(true); // Triggers Sheets scope, then delete
  
  return { success: true, email: email };
}

/**
 * Debug function - returns current user state
 */
function debugUserState() {
  const userProps = getUserProps().getProperties();
  const scriptProps = getScriptProps().getProperties();
  
  let ssInfo = 'No spreadsheet ID saved';
  const ssId = userProps[KEYS.SPREADSHEET_ID] || scriptProps['SPREADSHEET_ID'];
  if (ssId) {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      ssInfo = 'OK: ' + ss.getUrl();
    } catch (e) {
      ssInfo = 'ERROR: ' + e.message;
    }
  }
  
  return {
    userProps: userProps,
    scriptProps: scriptProps,
    spreadsheetStatus: ssInfo,
    effectiveUser: Session.getEffectiveUser().getEmail() || 'unknown',
  };
}

/**
 * Clear all cached user data and start fresh
 * Call this if you're getting permission errors
 */
function clearUserDataApi() {
  // Clear user properties
  getUserProps().deleteAllProperties();
  
  // Also clear any old script properties that might be cached
  // (from before we switched to user properties)
  try {
    getScriptProps().deleteProperty('SPREADSHEET_ID');
  } catch (e) {
    // Ignore - might not have permission
  }
  
  return { success: true, message: 'User data cleared. A new spreadsheet will be created on next action.' };
}

/* ================== Auto-Scan Trigger Functions ================== */

/**
 * Start automatic background scanning
 * Creates a time-based trigger that runs every 5 minutes until complete
 */
function startAutoScan() {
  // Remove any existing triggers first
  stopAutoScan();
  
  // Create a trigger to run every 5 minutes
  ScriptApp.newTrigger('autoScanBatch')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  getUserProps().setProperty(KEYS.AUTO_SCAN_ENABLED, 'true');
  
  // Run first batch immediately
  buildSenderReportBatch();
  
  return { success: true, message: 'Auto-scan started. Will run every 5 minutes until complete.' };
}

/**
 * Stop automatic background scanning
 */
function stopAutoScan() {
  // Delete all triggers for autoScanBatch
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'autoScanBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  getUserProps().deleteProperty(KEYS.AUTO_SCAN_ENABLED);
  
  return { success: true, message: 'Auto-scan stopped.' };
}

/**
 * Check if auto-scan is enabled
 */
function isAutoScanEnabled() {
  return getUserProps().getProperty(KEYS.AUTO_SCAN_ENABLED) === 'true';
}

/**
 * Called by the trigger - runs a batch and stops if complete
 */
function autoScanBatch() {
  console.log('Auto-scan batch starting...');
  
  try {
    buildSenderReportBatch();
    
    // Check if we're done
    const props = getUserProps();
    const cursor = parseInt(props.getProperty(KEYS.CURSOR) || '0', 10);
    const remaining = GmailApp.search(CONFIG.INBOX_QUERY, cursor, 1);
    
    if (remaining.length === 0) {
      // We're done! Stop the trigger
      console.log('Auto-scan complete! Stopping trigger.');
      stopAutoScan();
      
      // Send email notification
      try {
        const ss = getSpreadsheet();
        const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
        const stats = getSheetStats(sheet);
        
        GmailApp.sendEmail(
          Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(),
          '✅ Gmail Inbox Scan Complete',
          `Your inbox scan is complete!\n\n` +
          `Unique senders: ${stats.senderCount}\n` +
          `Total emails: ${stats.totalEmails}\n\n` +
          `View results: ${ss.getUrl()}`
        );
      } catch (e) {
        console.log('Could not send notification email: ' + e.message);
      }
    } else {
      console.log('Auto-scan batch complete. More to scan. Cursor at: ' + cursor);
    }
  } catch (e) {
    console.error('Auto-scan error: ' + e.message);
    // Don't stop on error - will retry next interval
  }
}

/* ================== HTTP Web App Endpoints ================== */

/**
 * Handle GET requests
 * Endpoints:
 *   ?action=status     - Get scan status and stats
 *   ?action=senders    - Get list of senders (with optional filters)
 *   ?action=ui         - Serve HTML UI (default)
 */
function doGet(e) {
  const action = (e.parameter.action || 'ui').toLowerCase();
  
  try {
    switch (action) {
      case 'status':
        return jsonResponse(getStatusApi());
      
      case 'senders':
        return jsonResponse(getSendersApi(e.parameter));
      
      case 'ui':
      default:
        return HtmlService.createHtmlOutput(getHtmlUi())
          .setTitle('Gmail Inbox Cleaner')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) }, 500);
  }
}

/**
 * Handle POST requests
 * Endpoints:
 *   action=scan.start    - Start or resume scan
 *   action=scan.reset    - Reset scan (requires confirm=true)
 *   action=do.unsubscribe - Unsubscribe from specified senders
 *   action=do.delete      - Delete emails from specified senders
 *   action=do.both        - Unsubscribe and delete
 *   action=check          - Check/uncheck senders for action
 */
function doPost(e) {
  try {
    const body = e.postData ? JSON.parse(e.postData.contents) : {};
    const action = (body.action || e.parameter.action || '').toLowerCase();
    
    switch (action) {
      case 'scan.start':
        return jsonResponse(startScanApi());
      
      case 'scan.reset':
        if (!body.confirm) {
          return jsonResponse({ error: 'Reset requires confirm=true' }, 400);
        }
        return jsonResponse(resetScanApi());
      
      case 'do.unsubscribe':
        return jsonResponse(doActionApi(body.emails || [], { unsubscribe: true, delete: false }));
      
      case 'do.delete':
        return jsonResponse(doActionApi(body.emails || [], { unsubscribe: false, delete: true }));
      
      case 'do.both':
        return jsonResponse(doActionApi(body.emails || [], { unsubscribe: true, delete: true }));
      
      case 'check':
        return jsonResponse(setCheckedApi(body.emails || [], body.checked !== false));
      
      default:
        return jsonResponse({ error: `Unknown action: ${action}` }, 400);
    }
  } catch (err) {
    return jsonResponse({ error: err.message || String(err) }, 500);
  }
}

/** Helper: JSON response */
function jsonResponse(data, status = 200) {
  const output = ContentService.createTextOutput(JSON.stringify(data, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/* ================== API Implementations ================== */

/** GET ?action=status */
function getStatusApi() {
  // First check if user is authorized
  const authStatus = checkAuthStatus();
  if (!authStatus.authorized) {
    return {
      error: 'Please authorize the app to access your Gmail and Google Sheets.',
      needsAuth: true,
      status: 'not_authorized',
      user: { email: 'Not authorized', isLoggedIn: false },
    };
  }
  
  const props = getUserProps();
  
  let ss;
  try {
    ss = getSpreadsheet();
  } catch (e) {
    return { 
      error: e.message, 
      status: 'error',
      user: getUserInfo(),
    };
  }
  
  if (!ss) {
    return { 
      error: 'Spreadsheet not found. Try clicking "Reset Data" to create a new one.', 
      status: 'error',
      user: getUserInfo(),
    };
  }
  
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  const cursor = parseInt(props.getProperty(KEYS.CURSOR) || '0', 10);
  const totalProcessed = parseInt(props.getProperty(KEYS.TOTAL_PROCESSED) || '0', 10);
  const initialized = !!props.getProperty(KEYS.INITIALIZED);
  const autoScanEnabled = isAutoScanEnabled();
  
  let stats = { senderCount: 0, totalEmails: 0 };
  if (sheet && sheet.getLastRow() > 1) {
    stats = getSheetStats(sheet);
  }
  
  // Check if there are more messages to scan
  let hasMore = false;
  if (initialized) {
    const remaining = GmailApp.search(CONFIG.INBOX_QUERY, cursor, 1);
    hasMore = remaining.length > 0;
  }
  
  // Get user info
  const userInfo = getUserInfo();
  
  return {
    status: initialized ? (hasMore ? 'in_progress' : 'complete') : 'not_started',
    cursor,
    totalProcessed,
    senderCount: stats.senderCount,
    totalEmails: stats.totalEmails,
    hasMore,
    autoScanEnabled,
    spreadsheetUrl: ss.getUrl(),
    user: userInfo,
  };
}

/**
 * Get current user info
 */
function getUserInfo() {
  let email = '';
  
  try {
    // In deployed web apps, getEffectiveUser() usually works
    // getActiveUser() often returns empty unless deployed with "Execute as: User"
    const effectiveUser = Session.getEffectiveUser();
    if (effectiveUser) {
      email = effectiveUser.getEmail();
    }
    
    // If still empty, try active user
    if (!email) {
      const activeUser = Session.getActiveUser();
      if (activeUser) {
        email = activeUser.getEmail();
      }
    }
  } catch (e) {
    console.log('Could not get user info: ' + e.message);
  }
  
  return {
    email: email || 'Unknown',
    isLoggedIn: !!email && email !== 'Unknown',
  };
}

/**
 * Get the deployed script URL (for redirects)
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Clear all user data and start fresh - useful for fixing broken state
 */
function clearUserDataApi() {
  try {
    getUserProps().deleteAllProperties();
    return { success: true, message: 'Data cleared! Refresh the page to create a new spreadsheet.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Debug function - returns current user state
 */
function debugStateApi() {
  const props = getUserProps();
  const allProps = props.getProperties();
  
  let ssInfo = { id: null, url: null, name: null, error: null };
  const ssId = allProps[KEYS.SPREADSHEET_ID];
  
  if (ssId) {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      ssInfo = { id: ssId, url: ss.getUrl(), name: ss.getName(), error: null };
    } catch (e) {
      ssInfo = { id: ssId, url: null, name: null, error: e.message };
    }
  }
  
  let userEmail = 'unknown';
  try {
    userEmail = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || 'unknown';
  } catch (e) {}
  
  return {
    properties: allProps,
    spreadsheet: ssInfo,
    user: userEmail,
  };
}

/** GET ?action=senders */
function getSendersApi(params) {
  params = params || {};
  
  let ss;
  try {
    ss = getSpreadsheet();
  } catch (e) {
    return { error: 'Could not access spreadsheet: ' + e.message, senders: [], total: 0 };
  }
  
  if (!ss) {
    return { error: 'Spreadsheet not found. Click "Reset Data" to create one.', senders: [], total: 0 };
  }
  
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { senders: [], total: 0 };
  }
  
  const limit = Math.min(parseInt(params.limit || '100', 10), 500);
  const offset = parseInt(params.offset || '0', 10);
  const minCount = parseInt(params.minCount || '0', 10);
  const onlyUnchecked = params.onlyUnchecked === 'true';
  const onlyUnprocessed = params.onlyUnprocessed === 'true';
  
  const data = sheet.getDataRange().getValues();
  const senders = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const count = Number(row[3] || 0);
    const isChecked = !!row[0];
    const isProcessed = !!(row[9] || '').toString().trim();
    
    // Apply filters
    if (count < minCount) continue;
    if (onlyUnchecked && isChecked) continue;
    if (onlyUnprocessed && isProcessed) continue;
    
    senders.push({
      row: i + 1,
      checked: isChecked,
      email: row[1] || '',
      name: row[2] || '',
      count: count,
      subject: row[4] || '',
      lastSeen: row[5] || '',
      unsubUrl: row[6] || '',
      unsubMailto: row[7] || '',
      notes: row[8] || '',
      processed: row[9] || '',
    });
  }
  
  // Sort by count desc
  senders.sort((a, b) => b.count - a.count);
  
  return {
    senders: senders.slice(offset, offset + limit),
    total: senders.length,
    offset,
    limit,
  };
}

/** POST action=scan.start */
function startScanApi() {
  const startTime = Date.now();
  
  // Run one batch
  buildSenderReportBatch();
  
  const elapsed = Date.now() - startTime;
  const status = getStatusApi();
  
  return {
    ...status,
    elapsedMs: elapsed,
    message: status.hasMore ? 'Batch complete, more to scan' : 'Scan complete',
  };
}

/** POST action=scan.reset */
function resetScanApi() {
  const ss = getSpreadsheet();
  
  if (!ss) {
    return { error: 'Spreadsheet not found. Run saveSpreadsheetId from the sheet first.' };
  }
  
  // Delete and recreate sheets
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  let indexSheet = ss.getSheetByName(CONFIG.INDEX_SHEET_NAME);
  if (indexSheet) ss.deleteSheet(indexSheet);
  indexSheet = ss.insertSheet(CONFIG.INDEX_SHEET_NAME);
  indexSheet.hideSheet();

  // Set up headers
  const headers = [
    'Act?', 'Sender Email', 'Sender Name', 'Count', 'Example Subject',
    'Last Seen (ISO)', 'Unsubscribe URL', 'Unsubscribe Mailto', 'Notes / Status', 'Processed'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange('D:D').setNumberFormat('0');
  sheet.getRange('F:F').setNumberFormat('@');
  sheet.getRange('J:J').setNumberFormat('@');
  if (!sheet.getFilter()) sheet.getDataRange().createFilter();
  
  indexSheet.getRange(1, 1, 1, 2).setValues([['email_lc', 'row']]);
  indexSheet.setFrozenRows(1);

  getUserProps().deleteAllProperties();
  
  return { success: true, message: 'Scan reset complete' };
}

/** POST action=do.* */
function doActionApi(emails, options) {
  if (!emails || !emails.length) {
    return { error: 'No emails provided', processed: 0 };
  }
  
  const ss = getSpreadsheet();
  
  if (!ss) {
    return { error: 'Spreadsheet not found. Run saveSpreadsheetId from the sheet first.', processed: 0 };
  }
  
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    return { error: 'Sheet not found', processed: 0 };
  }
  
  migrateAddProcessedColumn(sheet);
  
  const data = sheet.getDataRange().getValues();
  const emailSet = new Set(emails.map(e => e.toLowerCase()));
  
  const results = [];
  let totalDeleted = 0;
  let unsubSuccess = 0;
  let unsubFailed = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = (row[1] || '').toString().trim().toLowerCase();
    
    if (!emailSet.has(email)) continue;
    
    const unsubUrl = (row[6] || '').toString().trim();
    const unsubMailto = (row[7] || '').toString().trim();
    
    const result = { email, actions: [] };
    
    try {
      if (options.unsubscribe) {
        const unsubResult = unsubscribeIfPossible(unsubUrl, unsubMailto);
        result.actions.push({ type: 'unsubscribe', result: unsubResult });
        
        if (unsubResult.includes('OK') || unsubResult.includes('sent')) {
          unsubSuccess++;
        } else {
          unsubFailed++;
        }
      }
      
      if (options.delete) {
        const deleteResult = deleteAllFromSender(email);
        result.actions.push({ type: 'delete', result: deleteResult });
        
        const countMatch = deleteResult.match(/(\d+)\s+thread/);
        const deletedCount = countMatch ? parseInt(countMatch[1], 10) : 0;
        totalDeleted += deletedCount;
        result.deletedThreads = deletedCount;
      }
      
      // Update sheet: notes, uncheck, set processed
      const notes = result.actions.map(a => a.result).join(' | ');
      const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
      
      sheet.getRange(i + 1, 1).setValue(false);  // Uncheck
      sheet.getRange(i + 1, 9).setValue(notes);  // Notes
      sheet.getRange(i + 1, 10).setValue(timestamp);  // Processed
      
      result.success = true;
    } catch (err) {
      result.success = false;
      result.error = err.message || String(err);
      sheet.getRange(i + 1, 9).setValue(`ERROR: ${result.error}`);
    }
    
    results.push(result);
  }
  
  return {
    processed: results.length,
    totalDeleted,
    unsubSuccess,
    unsubFailed,
    results,
  };
}

/** POST action=check */
function setCheckedApi(emails, checked) {
  const ss = getSpreadsheet();
  
  if (!ss) {
    return { error: 'Spreadsheet not found. Run saveSpreadsheetId from the sheet first.', updated: 0 };
  }
  
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    return { error: 'Sheet not found', updated: 0 };
  }
  
  const data = sheet.getDataRange().getValues();
  const emailSet = new Set(emails.map(e => e.toLowerCase()));
  
  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const email = (data[i][1] || '').toString().trim().toLowerCase();
    if (emailSet.has(email)) {
      sheet.getRange(i + 1, 1).setValue(checked);
      updated++;
    }
  }
  
  return { updated, checked };
}

/* ================== HTML UI ================== */

function getHtmlUi() {
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Gmail Inbox Cleaner</title>
  <style>
    * { box-sizing: border-box; }
    body { 
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      max-width: 1200px; 
      margin: 0 auto; 
      padding: 20px;
      background: #f5f5f5;
    }
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 20px;
    }
    h1 { color: #333; margin: 0; }
    .user-info {
      display: flex;
      align-items: center;
      gap: 12px;
      background: #fff;
      padding: 8px 16px;
      border-radius: 24px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .user-avatar {
      width: 32px;
      height: 32px;
      border-radius: 50%;
      background: #9e9e9e;
      display: flex;
      align-items: center;
      justify-content: center;
      color: white;
      font-weight: 600;
      font-size: 14px;
    }
    .user-email {
      font-size: 14px;
      color: #333;
      max-width: 200px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .user-actions {
      display: flex;
      gap: 8px;
    }
    .user-actions button {
      padding: 6px 12px;
      font-size: 12px;
      border-radius: 4px;
    }
    .description {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      border-radius: 8px;
      margin-bottom: 20px;
      overflow: hidden;
    }
    .description-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 15px 20px;
      cursor: pointer;
    }
    .description-header:hover {
      background: rgba(255,255,255,0.1);
    }
    .description-header h2 { margin: 0; font-size: 18px; }
    .description-toggle {
      font-size: 20px;
      transition: transform 0.2s;
    }
    .description-toggle.collapsed {
      transform: rotate(-90deg);
    }
    .description-content {
      padding: 0 20px 20px 20px;
      transition: max-height 0.3s ease;
      overflow: hidden;
    }
    .description-content.collapsed {
      max-height: 0 !important;
      padding-bottom: 0;
    }
    .description p { margin: 5px 0; font-size: 14px; opacity: 0.95; }
    .description ul { margin: 10px 0; padding-left: 20px; }
    .description li { margin: 5px 0; font-size: 13px; }
    .description a { color: #fff; text-decoration: underline; }
    .status-bar {
      background: #fff;
      padding: 15px;
      border-radius: 8px;
      margin-bottom: 20px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .status-bar .stats { display: flex; gap: 30px; flex-wrap: wrap; }
    .status-bar .stat { }
    .status-bar .stat-value { font-size: 24px; font-weight: bold; color: #1a73e8; }
    .status-bar .stat-label { font-size: 12px; color: #666; }
    .status-bar .stat-value.running { color: #f9ab00; }
    .status-bar .stat-value.complete { color: #34a853; }
    .spreadsheet-link { margin-top: 10px; font-size: 13px; }
    .spreadsheet-link a { color: #1a73e8; }
    .actions { margin: 20px 0; display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
    button {
      padding: 10px 20px;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 14px;
      font-weight: 500;
      transition: background 0.2s;
    }
    button:disabled { opacity: 0.5; cursor: not-allowed; }
    .btn-primary { background: #1a73e8; color: white; }
    .btn-primary:hover:not(:disabled) { background: #1557b0; }
    .btn-success { background: #34a853; color: white; }
    .btn-success:hover:not(:disabled) { background: #2d8e47; }
    .btn-warning { background: #f9ab00; color: white; }
    .btn-warning:hover:not(:disabled) { background: #e09d00; }
    .btn-danger { background: #d93025; color: white; }
    .btn-danger:hover:not(:disabled) { background: #b3261e; }
    .btn-secondary { background: #e8eaed; color: #333; }
    .btn-secondary:hover:not(:disabled) { background: #d2d5db; }
    .btn-ai { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; }
    .btn-ai:hover:not(:disabled) { background: linear-gradient(135deg, #5a6fd6 0%, #6a4190 100%); }
    .btn-text { background: transparent; color: #1a73e8; padding: 6px 12px; }
    .btn-text:hover { background: #e8f0fe; }
    .auto-scan-status {
      padding: 8px 12px;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 500;
    }
    .auto-scan-status.active { background: #e6f4ea; color: #137333; }
    .auto-scan-status.inactive { background: #f1f3f4; color: #5f6368; }
    .sender-list {
      background: #fff;
      border-radius: 8px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    .sender-header {
      display: grid;
      grid-template-columns: 40px 1fr 80px 200px 150px;
      padding: 12px 15px;
      background: #f8f9fa;
      font-weight: 600;
      font-size: 12px;
      color: #666;
      border-bottom: 1px solid #e0e0e0;
    }
    .sender-row {
      display: grid;
      grid-template-columns: 40px 1fr 80px 200px 150px;
      padding: 12px 15px;
      border-bottom: 1px solid #f0f0f0;
      align-items: center;
    }
    .sender-row:hover { background: #f8f9fa; }
    .sender-row.processed { opacity: 0.6; background: #f0fff0; }
    .sender-email { font-weight: 500; word-break: break-all; }
    .sender-name { font-size: 12px; color: #666; }
    .sender-count { font-weight: bold; color: #1a73e8; text-align: center; }
    .sender-subject { font-size: 12px; color: #888; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .sender-processed { font-size: 11px; color: #34a853; }
    .loading { text-align: center; padding: 40px; color: #666; }
    .checkbox { width: 18px; height: 18px; cursor: pointer; }
    .select-actions { display: flex; gap: 10px; align-items: center; margin-bottom: 10px; flex-wrap: wrap; }
    .filter-row { display: flex; gap: 15px; align-items: center; margin-bottom: 15px; flex-wrap: wrap; }
    .filter-row label { display: flex; align-items: center; gap: 5px; font-size: 14px; color: #555; cursor: pointer; }
    .log { 
      background: #1e1e1e; 
      color: #d4d4d4; 
      padding: 10px 15px; 
      border-radius: 8px; 
      font-family: monospace; 
      font-size: 12px;
      height: 72px;
      overflow-y: auto;
      margin-bottom: 20px;
    }
    .log-entry { margin: 2px 0; white-space: nowrap; }
    .log-entry.error { color: #f48771; }
    .log-entry.success { color: #89d185; }
    .log-entry.info { color: #6cb6ff; }
    .polling-indicator {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: #666;
      margin-left: 10px;
    }
    .polling-dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #34a853;
      animation: pulse 1s infinite;
    }
    @keyframes pulse {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.4; }
    }
    @media (max-width: 768px) {
      .header { flex-direction: column; gap: 15px; align-items: flex-start; }
      .user-info { width: 100%; justify-content: space-between; }
      .sender-header, .sender-row {
        grid-template-columns: 30px 1fr 60px;
      }
      .sender-header > div:nth-child(4),
      .sender-header > div:nth-child(5),
      .sender-row > div:nth-child(4),
      .sender-row > div:nth-child(5) {
        display: none;
      }
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>📧 Gmail Inbox Cleaner</h1>
    <div class="user-info">
      <div class="user-avatar" id="userAvatar">?</div>
      <div class="user-email" id="userEmail">Loading...</div>
      <div class="user-actions">
        <button class="btn-text" onclick="openGoogleAccount()" title="Google Account">Account</button>
        <button class="btn-text" onclick="clearData()" title="Clear all data and start fresh">Reset Data</button>
        <button class="btn-text" onclick="showDebug()" title="Debug info">Debug</button>
      </div>
    </div>
  </div>
  
  <div class="description">
    <div class="description-header" onclick="toggleDescription()">
      <h2>ℹ️ How it works</h2>
      <span class="description-toggle" id="descToggle">▼</span>
    </div>
    <div class="description-content" id="descContent">
      <p>This tool scans your Gmail inbox and groups emails by sender, making it easy to bulk unsubscribe and delete.</p>
      <ul>
        <li><strong>Scan:</strong> Analyzes your inbox in batches (Gmail API has rate limits). Large inboxes may need multiple runs.</li>
        <li><strong>Auto-Scan:</strong> Runs automatically every 5 minutes in the background until complete. You'll get an email when done.</li>
        <li><strong>Unsubscribe:</strong> Automatically hits unsubscribe links or sends unsubscribe emails.</li>
        <li><strong>Delete:</strong> Moves ALL emails from selected senders to trash (not just inbox).</li>
      </ul>
      <p>⚠️ Gmail API limits: ~250 operations/second, daily quotas vary. If you hit limits, wait and try again later.</p>
      <p>📊 Data is stored in a Google Sheet: <a href="#" id="sheetLink" target="_blank">Open Spreadsheet</a></p>
    </div>
  </div>
  
  <div class="status-bar">
    <div class="stats">
      <div class="stat">
        <div class="stat-value" id="senderCount">-</div>
        <div class="stat-label">Unique Senders</div>
      </div>
      <div class="stat">
        <div class="stat-value" id="totalEmails">-</div>
        <div class="stat-label">Total Emails</div>
      </div>
      <div class="stat">
        <div class="stat-value" id="scanStatus">-</div>
        <div class="stat-label">Scan Status</div>
      </div>
      <div class="stat">
        <div class="stat-value" id="cursor">-</div>
        <div class="stat-label">Threads Scanned</div>
      </div>
    </div>
  </div>
  
  <div id="authBtnContainer" style="display: none; margin-bottom: 20px;">
    <div style="background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 15px; display: flex; align-items: center; gap: 15px;">
      <span style="font-size: 24px;">🔐</span>
      <div style="flex: 1;">
        <strong>Authorization Required</strong>
        <p style="margin: 5px 0 0 0; font-size: 14px; color: #666;">This app needs permission to access your Gmail and Google Sheets.</p>
      </div>
      <button class="btn-primary" onclick="authorizeApp()" style="white-space: nowrap;">🔓 Authorize App</button>
    </div>
  </div>
  
  <div class="actions">
    <button class="btn-primary" id="scanBtn" onclick="startScan()">▶️ Run One Batch</button>
    <button class="btn-success" id="autoScanBtn" onclick="toggleAutoScan()">🔄 Start Auto-Scan</button>
    <button class="btn-danger" id="resetBtn" onclick="resetScan()">🗑️ Reset Scan</button>
    <button class="btn-secondary" onclick="refreshData()">🔃 Refresh</button>
    <span class="auto-scan-status inactive" id="autoScanStatus">Auto-scan: Off</span>
    <span class="polling-indicator" id="pollingIndicator" style="display: none;">
      <span class="polling-dot"></span>
      Updating...
    </span>
  </div>
  
  <div class="log" id="log"></div>
  
  <div class="filter-row">
    <label>
      <input type="checkbox" id="hideProcessed" checked onchange="renderSenders()">
      Hide processed senders
    </label>
    <span style="color: #888; font-size: 13px;" id="filterInfo"></span>
  </div>
  
  <div class="select-actions">
    <button class="btn-secondary" onclick="selectAll()">Select All</button>
    <button class="btn-secondary" onclick="selectNone()">Select None</button>
    <button class="btn-secondary" onclick="selectTop(10)">Top 10</button>
    <button class="btn-secondary" onclick="selectTop(25)">Top 25</button>
    <button class="btn-ai" onclick="autoselectWithAI()" title="Use AI to suggest senders to unsubscribe">🤖 Autoselect with AI</button>
    <span style="margin-left: auto; display: flex; gap: 10px;">
      <button class="btn-secondary" onclick="doAction('unsubscribe')">📧 Unsubscribe</button>
      <button class="btn-danger" onclick="doAction('delete')">🗑️ Delete</button>
      <button class="btn-danger" onclick="doAction('both')">⚡ Both</button>
    </span>
  </div>
  
  <div class="sender-list">
    <div class="sender-header">
      <div>✓</div>
      <div>Sender</div>
      <div>Count</div>
      <div>Example Subject</div>
      <div>Processed</div>
    </div>
    <div id="senderRows">
      <div class="loading">Loading...</div>
    </div>
  </div>

  <script>
    let senders = [];
    let autoScanEnabled = false;
    let currentUser = null;
    let pollingInterval = null;
    let isActionInProgress = false;
    let descriptionCollapsed = false;
    
    // AI classification endpoint (Vercel-deployed Gemini API)
    const AI_ENDPOINT = 'https://gmail-inbox-cleaner-ai.vercel.app/api/classify';
    
    function log(msg, type = '') {
      const logEl = document.getElementById('log');
      const entry = document.createElement('div');
      entry.className = 'log-entry ' + type;
      entry.textContent = new Date().toLocaleTimeString() + ' - ' + msg;
      logEl.insertBefore(entry, logEl.firstChild);
      // Keep only last 100 entries
      while (logEl.children.length > 100) {
        logEl.removeChild(logEl.lastChild);
      }
    }
    
    function toggleDescription() {
      descriptionCollapsed = !descriptionCollapsed;
      const content = document.getElementById('descContent');
      const toggle = document.getElementById('descToggle');
      if (descriptionCollapsed) {
        content.classList.add('collapsed');
        toggle.classList.add('collapsed');
      } else {
        content.classList.remove('collapsed');
        toggle.classList.remove('collapsed');
      }
      localStorage.setItem('descCollapsed', descriptionCollapsed);
    }
    
    // Restore description state
    if (localStorage.getItem('descCollapsed') === 'true') {
      toggleDescription();
    }
    
    // Promisified wrapper for google.script.run
    function callServer(fnName, ...args) {
      return new Promise((resolve, reject) => {
        google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler(reject)
          [fnName](...args);
      });
    }
    
    function startPolling() {
      if (pollingInterval) return;
      document.getElementById('pollingIndicator').style.display = 'inline-flex';
      pollingInterval = setInterval(async () => {
        await refreshData(true); // silent refresh
      }, 10000);
    }
    
    function stopPolling() {
      if (pollingInterval) {
        clearInterval(pollingInterval);
        pollingInterval = null;
      }
      document.getElementById('pollingIndicator').style.display = 'none';
    }
    
    function updatePollingState(status) {
      if (status === 'in_progress' || isActionInProgress || autoScanEnabled) {
        startPolling();
      } else {
        stopPolling();
      }
    }
    
    async function refreshStatus(silent = false) {
      try {
        const data = await callServer('getStatusApi');
        
        // Handle authorization needed
        if (data.needsAuth) {
          if (!silent) log('Authorization required. Please click "Authorize" to grant permissions.', 'error');
          document.getElementById('senderCount').textContent = '-';
          document.getElementById('totalEmails').textContent = '-';
          document.getElementById('scanStatus').textContent = 'Not Authorized';
          document.getElementById('cursor').textContent = '-';
          
          const statusEl = document.getElementById('scanStatus');
          statusEl.className = 'stat-value';
          statusEl.style.color = '#f9ab00';
          
          showAuthButton(true);
          currentUser = { email: 'Not authorized', isLoggedIn: false };
          updateUserUI();
          return;
        }
        
        showAuthButton(false);
        
        if (data.error) {
          if (!silent) log('Error: ' + data.error, 'error');
          document.getElementById('senderCount').textContent = '-';
          document.getElementById('totalEmails').textContent = '-';
          document.getElementById('scanStatus').textContent = 'Error';
          document.getElementById('cursor').textContent = '-';
          
          const statusEl = document.getElementById('scanStatus');
          statusEl.className = 'stat-value';
          statusEl.style.color = '#d93025';
          statusEl.title = data.error;
        } else {
          document.getElementById('senderCount').textContent = data.senderCount || 0;
          document.getElementById('totalEmails').textContent = data.totalEmails || 0;
          document.getElementById('cursor').textContent = data.totalProcessed || 0;
          
          const statusEl = document.getElementById('scanStatus');
          statusEl.textContent = data.status || 'unknown';
          statusEl.style.color = '';
          statusEl.title = '';
          statusEl.className = 'stat-value ' + (data.status === 'complete' ? 'complete' : data.status === 'in_progress' ? 'running' : '');
          
          autoScanEnabled = data.autoScanEnabled;
          updateAutoScanUI();
          
          if (data.spreadsheetUrl) {
            document.getElementById('sheetLink').href = data.spreadsheetUrl;
          }
          
          // Update polling based on status
          updatePollingState(data.status);
        }
        
        if (data.user) {
          currentUser = data.user;
        } else {
          currentUser = { email: 'Unknown', isLoggedIn: false };
        }
        updateUserUI();
        
      } catch (err) {
        if (!silent) {
          log('Error getting status: ' + (err.message || err), 'error');
        }
        currentUser = { email: 'Error', isLoggedIn: false };
        updateUserUI();
      }
    }
    
    function showAuthButton(show) {
      const authBtnContainer = document.getElementById('authBtnContainer');
      if (authBtnContainer) {
        authBtnContainer.style.display = show ? 'block' : 'none';
      }
    }
    
    async function authorizeApp() {
      log('Opening authorization...', 'success');
      try {
        const scriptUrl = await callServer('getScriptUrl');
        log('Opening authorization page. Please authorize in the new window, then return here and click Refresh.', 'success');
        window.open(scriptUrl, '_blank');
      } catch (err) {
        log('Reloading page to trigger authorization...', 'success');
        window.top.location.reload();
      }
    }
    
    function updateUserUI() {
      const avatarEl = document.getElementById('userAvatar');
      const emailEl = document.getElementById('userEmail');
      
      if (currentUser && currentUser.isLoggedIn && currentUser.email) {
        const initial = currentUser.email.charAt(0).toUpperCase();
        avatarEl.textContent = initial;
        avatarEl.style.background = 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)';
        emailEl.textContent = currentUser.email;
        emailEl.title = currentUser.email;
      } else if (currentUser && currentUser.email === 'Unknown') {
        avatarEl.textContent = '👤';
        avatarEl.style.background = '#9e9e9e';
        emailEl.textContent = 'Signed in (Execute as owner)';
        emailEl.title = 'The script runs with the owner\\'s permissions';
      } else {
        avatarEl.textContent = '?';
        avatarEl.style.background = '#9e9e9e';
        emailEl.textContent = 'Not signed in';
        emailEl.title = '';
      }
    }
    
    function updateAutoScanUI() {
      const statusEl = document.getElementById('autoScanStatus');
      const btnEl = document.getElementById('autoScanBtn');
      
      if (autoScanEnabled) {
        statusEl.textContent = 'Auto-scan: Running';
        statusEl.className = 'auto-scan-status active';
        btnEl.textContent = '⏹️ Stop Auto-Scan';
        btnEl.className = 'btn-warning';
      } else {
        statusEl.textContent = 'Auto-scan: Off';
        statusEl.className = 'auto-scan-status inactive';
        btnEl.textContent = '🔄 Start Auto-Scan';
        btnEl.className = 'btn-success';
      }
    }
    
    async function refreshSenders(silent = false) {
      try {
        const data = await callServer('getSendersApi', { limit: '500', offset: '0' });
        if (data.error) {
          if (!silent) log('Error: ' + data.error, 'error');
          senders = [];
        } else {
          senders = data.senders || [];
        }
        renderSenders();
      } catch (err) {
        if (!silent) log('Error getting senders: ' + (err.message || err), 'error');
      }
    }
    
    function escapeHtml(str) {
      const div = document.createElement('div');
      div.textContent = str || '';
      return div.innerHTML;
    }
    
    function renderSenders() {
      const container = document.getElementById('senderRows');
      const hideProcessed = document.getElementById('hideProcessed').checked;
      
      let filtered = senders;
      if (hideProcessed) {
        filtered = senders.filter(s => !s.processed);
      }
      
      // Update filter info
      const hiddenCount = senders.length - filtered.length;
      document.getElementById('filterInfo').textContent = 
        hiddenCount > 0 ? '(' + hiddenCount + ' processed senders hidden)' : '';
      
      if (!filtered.length) {
        container.innerHTML = '<div class="loading">' + 
          (senders.length > 0 ? 'All senders have been processed. Uncheck "Hide processed" to see them.' : 'No senders found. Run a scan first.') + 
          '</div>';
        return;
      }
      
      container.innerHTML = filtered.map((s, i) => 
        '<div class="sender-row ' + (s.processed ? 'processed' : '') + '">' +
          '<div><input type="checkbox" class="checkbox" data-email="' + escapeHtml(s.email) + '" ' + (s.checked ? 'checked' : '') + '></div>' +
          '<div>' +
            '<div class="sender-email">' + escapeHtml(s.email) + '</div>' +
            '<div class="sender-name">' + escapeHtml(s.name || '') + '</div>' +
          '</div>' +
          '<div class="sender-count">' + s.count + '</div>' +
          '<div class="sender-subject" title="' + escapeHtml(s.subject) + '">' + escapeHtml(s.subject) + '</div>' +
          '<div class="sender-processed">' + escapeHtml(s.processed || '') + '</div>' +
        '</div>'
      ).join('');
    }
    
    async function refreshData(silent = false) {
      if (!silent) log('Refreshing data...', 'info');
      await Promise.all([refreshStatus(silent), refreshSenders(silent)]);
      if (!silent) log('Data refreshed', 'success');
    }
    
    async function startScan() {
      const btn = document.getElementById('scanBtn');
      btn.disabled = true;
      btn.textContent = '⏳ Scanning...';
      log('Starting scan batch...', 'info');
      
      try {
        const result = await callServer('startScanApi');
        log(result.message + ' - ' + result.senderCount + ' senders, ' + result.totalEmails + ' emails', 'success');
        await refreshData();
        
        if (result.hasMore) {
          log('More messages to scan. Click again or use Auto-Scan.', 'info');
        }
      } catch (err) {
        log('Scan error: ' + (err.message || err), 'error');
      }
      
      btn.disabled = false;
      btn.textContent = '▶️ Run One Batch';
    }
    
    async function toggleAutoScan() {
      const btn = document.getElementById('autoScanBtn');
      btn.disabled = true;
      
      try {
        if (autoScanEnabled) {
          log('Stopping auto-scan...', 'info');
          const result = await callServer('stopAutoScan');
          log(result.message, 'success');
        } else {
          log('Starting auto-scan (runs every 5 minutes)...', 'info');
          const result = await callServer('startAutoScan');
          log(result.message, 'success');
        }
        await refreshData();
      } catch (err) {
        log('Auto-scan error: ' + (err.message || err), 'error');
      }
      
      btn.disabled = false;
    }
    
    async function resetScan() {
      if (!confirm('Are you sure? This will delete all scan data.')) return;
      
      log('Resetting scan...', 'info');
      try {
        if (autoScanEnabled) {
          await callServer('stopAutoScan');
        }
        const result = await callServer('resetScanApi');
        log(result.message, 'success');
        await refreshData();
      } catch (err) {
        log('Reset error: ' + (err.message || err), 'error');
      }
    }
    
    function getSelectedEmails() {
      return Array.from(document.querySelectorAll('.checkbox:checked')).map(cb => cb.dataset.email);
    }
    
    function selectAll() {
      document.querySelectorAll('.checkbox').forEach(cb => cb.checked = true);
    }
    
    function selectNone() {
      document.querySelectorAll('.checkbox').forEach(cb => cb.checked = false);
    }
    
    function selectTop(n) {
      selectNone();
      document.querySelectorAll('.checkbox').forEach((cb, i) => { if (i < n) cb.checked = true; });
    }
    
    async function autoselectWithAI() {
      const hideProcessed = document.getElementById('hideProcessed').checked;
      let sendersForAI = hideProcessed ? senders.filter(s => !s.processed) : senders;
      
      if (sendersForAI.length === 0) {
        alert('No senders to analyze. Run a scan first.');
        return;
      }
      
      log('Sending ' + sendersForAI.length + ' senders to AI for analysis...', 'info');
      
      // Prepare data for AI
      const aiPayload = {
        senders: sendersForAI.map(s => ({
          email: s.email,
          name: s.name,
          count: s.count,
          subject: s.subject,
        }))
      };
      
      try {
        const response = await fetch(AI_ENDPOINT, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(aiPayload)
        });

        if (!response.ok) {
          const errBody = await response.json().catch(() => ({}));
          throw new Error(errBody.error || 'AI service returned ' + response.status);
        }

        const result = await response.json();
        const suggestedEmails = Array.isArray(result.delete) ? result.delete : [];

        // Select the suggested emails
        selectNone();
        let selectedCount = 0;
        const suggestedSet = new Set(suggestedEmails.map(e => e.toLowerCase()));
        document.querySelectorAll('.checkbox').forEach(cb => {
          if (suggestedSet.has((cb.dataset.email || '').toLowerCase())) {
            cb.checked = true;
            selectedCount++;
          }
        });

        log('AI suggested ' + selectedCount + ' senders for unsubscribe/delete', 'success');
        if (selectedCount === 0) {
          log('No obvious newsletters/marketing emails detected. Try selecting manually.', 'info');
        }

      } catch (err) {
        log('AI error: ' + (err.message || err), 'error');
      }
    }
    
    async function doAction(type) {
      const emails = getSelectedEmails();
      if (!emails.length) {
        alert('No senders selected');
        return;
      }
      
      const actionName = { unsubscribe: 'Unsubscribe', delete: 'Delete', both: 'Unsubscribe & Delete' }[type];
      if (!confirm(actionName + ' for ' + emails.length + ' sender(s)?')) return;
      
      log('Running ' + actionName + ' for ' + emails.length + ' senders...', 'info');
      isActionInProgress = true;
      updatePollingState();
      
      try {
        const options = {
          unsubscribe: type === 'unsubscribe' || type === 'both',
          delete: type === 'delete' || type === 'both'
        };
        const result = await callServer('doActionApi', emails, options);
        
        if (result.error) {
          log('Error: ' + result.error, 'error');
        } else {
          let msg = 'Processed ' + result.processed + ' senders.';
          if (result.totalDeleted) msg += ' Deleted ' + result.totalDeleted + ' threads.';
          if (result.unsubSuccess) msg += ' Unsub success: ' + result.unsubSuccess + '.';
          log(msg, 'success');
        }
      } catch (err) {
        log('Action error: ' + (err.message || err), 'error');
      }
      
      isActionInProgress = false;
      updatePollingState();
      await refreshData();
    }
    
    function openGoogleAccount() {
      window.open('https://myaccount.google.com/', '_blank');
    }
    
    async function clearData() {
      if (!confirm('This will clear all your saved data (spreadsheet ID, scan progress) and create a fresh spreadsheet.\\n\\nContinue?')) {
        return;
      }
      
      log('Clearing user data...', 'info');
      try {
        const result = await callServer('clearUserDataApi');
        if (result.success) {
          log(result.message, 'success');
          alert('Data cleared! The page will now refresh.');
          window.location.reload();
        } else {
          log('Error: ' + result.error, 'error');
        }
      } catch (err) {
        log('Error clearing data: ' + (err.message || err), 'error');
      }
    }
    
    async function showDebug() {
      log('Fetching debug info...', 'info');
      try {
        const debug = await callServer('getDebugInfo');
        console.log('Debug info:', debug);
        
        const lines = [
          '=== DEBUG INFO ===',
          'Timestamp: ' + debug.timestamp,
          '',
          '--- Session ---',
          'Effective User: ' + debug.session.effectiveUser,
          'Active User: ' + debug.session.activeUser,
          '',
          '--- Gmail Auth ---',
          'Gmail OK: ' + debug.auth.gmailOk,
          debug.auth.gmailError ? 'Gmail Error: ' + debug.auth.gmailError : '',
          '',
          '--- Spreadsheet IDs ---',
          'From UserProps: ' + debug.spreadsheet.userPropsId,
          'From ScriptProps: ' + debug.spreadsheet.scriptPropsId,
          'From CONFIG: ' + debug.spreadsheet.configId,
          '',
          '--- getSpreadsheet() Test ---',
          'Success: ' + debug.getSpreadsheetTest.success,
          debug.getSpreadsheetTest.success ? 
            'URL: ' + debug.getSpreadsheetTest.url :
            'Error: ' + debug.getSpreadsheetTest.error,
          '',
          '--- User Properties ---',
          JSON.stringify(debug.properties.user, null, 2),
        ].filter(line => line !== '');
        
        alert(lines.join('\\n'));
        log('Debug info logged to console (F12)', 'success');
      } catch (err) {
        log('Debug error: ' + (err.message || err), 'error');
        alert('Error getting debug info: ' + (err.message || err));
      }
    }
    
    // Initial load
    refreshData();
  </script>
</body>
</html>`;
}

/* ================== UI & Alerts ================== */

function safeAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (_) {
    console.log(msg);
    const ss = getSpreadsheet();
    if (!ss) return;
    let logSheet = ss.getSheetByName('Log');
    if (!logSheet) logSheet = ss.insertSheet('Log');
    logSheet.appendRow([new Date(), msg]);
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Inbox Tools')
    .addSubMenu(
      ui.createMenu('Scan Inbox')
        .addItem('Start/Resume Scan (one batch)', 'buildSenderReportBatch')
        .addItem('Start Auto-Scan (background)', 'startAutoScan')
        .addItem('Stop Auto-Scan', 'stopAutoScan')
        .addSeparator()
        .addItem('Reset Scan (fresh start)', 'resetScan')
    )
    .addSubMenu(
      ui.createMenu('Actions (on checked rows)')
        .addItem('Unsubscribe Only', 'actionUnsubscribeOnly')
        .addItem('Delete All Messages', 'actionDeleteOnly')
        .addSeparator()
        .addItem('Unsubscribe & Delete', 'actionUnsubscribeAndDelete')
    )
    .addSubMenu(
      ui.createMenu('Settings')
        .addItem('Save Spreadsheet ID (for web app)', 'saveSpreadsheetId')
    )
    .addToUi();
}

/* ================== Main Scan Function ================== */

function buildSenderReportBatch() {
  const runStart = Date.now();
  const props = getUserProps();
  const ss = getSpreadsheet();

  // Ensure sheets exist
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  let indexSheet = ss.getSheetByName(CONFIG.INDEX_SHEET_NAME);
  if (!indexSheet) {
    indexSheet = ss.insertSheet(CONFIG.INDEX_SHEET_NAME);
    indexSheet.hideSheet();
  }

  // Migrate: add Processed column if missing (for existing sheets)
  if (sheet.getLastRow() > 0) {
    migrateAddProcessedColumn(sheet);
  }

  // First-time initialization — but NEVER wipe existing data
  if (!props.getProperty(KEYS.INITIALIZED)) {
    // Check if sheet already has data (header + at least one row)
    const existingRows = sheet.getLastRow();
    const indexRows = indexSheet.getLastRow();
    
    if (existingRows > 1 || indexRows > 1) {
      // Data exists! Don't wipe it. Just resume from the beginning.
      console.log(`Found existing data (${existingRows} rows in sheet, ${indexRows} in index). Resuming scan, not wiping.`);
      props.setProperty(KEYS.CURSOR, '0');
      props.setProperty(KEYS.TOTAL_PROCESSED, '0');
      props.setProperty(KEYS.INITIALIZED, '1');
    } else {
      // Truly fresh start — safe to set up headers
      setupHeader(sheet);
      resetIndexSheet(indexSheet);
      props.setProperty(KEYS.CURSOR, '0');
      props.setProperty(KEYS.TOTAL_PROCESSED, '0');
      props.setProperty(KEYS.INITIALIZED, '1');
    }
  }

  let cursor = parseInt(props.getProperty(KEYS.CURSOR) || '0', 10);
  let totalProcessed = parseInt(props.getProperty(KEYS.TOTAL_PROCESSED) || '0', 10);
  
  /** @type {Record<string, SenderBucket>} */
  let runMap = {};
  let threadsSinceCheckpoint = 0;
  let lastCheckpointTime = Date.now();
  let checkpointCount = 0;

  console.log(`Starting scan at cursor=${cursor}, totalProcessed=${totalProcessed}`);

  // Early exit if inbox is empty or we're already past all messages
  const initialCheck = GmailApp.search(CONFIG.INBOX_QUERY, cursor, 1);
  if (initialCheck.length === 0) {
    const finalStats = getSheetStats(sheet);
    if (finalStats.senderCount > 0) {
      safeAlert(`✅ Nothing new to scan.\n\nYour existing data:\n• ${finalStats.senderCount} senders\n• ${finalStats.totalEmails} emails\n\nUse "Reset Scan" if you want to start fresh.`);
    } else {
      safeAlert(`📭 No messages found in inbox matching query: "${CONFIG.INBOX_QUERY}"`);
    }
    return;
  }

  // Main loop
  while (true) {
    // Hard time limit check
    if (Date.now() - runStart > CONFIG.RUNTIME_HARD_LIMIT_MS) {
      console.log('Approaching time limit, forcing final checkpoint');
      break;
    }

    // Fetch next page of threads
    const threads = GmailApp.search(CONFIG.INBOX_QUERY, cursor, CONFIG.PAGE_SIZE_THREADS);
    
    if (!threads.length) {
      console.log('No more threads found');
      break;
    }

    // Batch fetch all messages (much faster than per-thread)
    const threadMsgs = GmailApp.getMessagesForThreads(threads);

    // Process each thread
    for (let tIdx = 0; tIdx < threads.length; tIdx++) {
      // Time check inside loop
      if (Date.now() - runStart > CONFIG.RUNTIME_HARD_LIMIT_MS) {
        break;
      }

      const msgs = threadMsgs[tIdx];
      if (!msgs || !msgs.length) continue;

      processMessages(msgs, runMap);
    }

    cursor += threads.length;
    threadsSinceCheckpoint += threads.length;
    totalProcessed += threads.length;

    // Checkpoint if thresholds reached
    const timeSinceCheckpoint = Date.now() - lastCheckpointTime;
    const shouldCheckpoint = 
      threadsSinceCheckpoint >= CONFIG.CHECKPOINT_THREADS ||
      timeSinceCheckpoint >= CONFIG.CHECKPOINT_INTERVAL_MS;

    if (shouldCheckpoint && Object.keys(runMap).length > 0) {
      // Get stats before checkpoint to calculate new vs updated
      const statsBefore = getSheetStats(sheet);
      
      checkpoint(sheet, indexSheet, props, runMap, cursor, totalProcessed);
      checkpointCount++;
      
      // Sanity check stats after
      const statsAfter = getSheetStats(sheet);
      const newSenders = statsAfter.senderCount - statsBefore.senderCount;
      const updatedSenders = Object.keys(runMap).length - newSenders;
      const newEmails = statsAfter.totalEmails - statsBefore.totalEmails;
      
      console.log(`Checkpoint #${checkpointCount}: cursor=${cursor} | batch: +${newSenders} new, ${updatedSenders} updated, +${newEmails} emails | TOTALS: ${statsAfter.senderCount} senders, ${statsAfter.totalEmails} emails`);
      
      // Reset for next chunk
      runMap = {};
      threadsSinceCheckpoint = 0;
      lastCheckpointTime = Date.now();
    }
  }

  // Final checkpoint for any remaining data
  if (Object.keys(runMap).length > 0) {
    const statsBefore = getSheetStats(sheet);
    checkpoint(sheet, indexSheet, props, runMap, cursor, totalProcessed);
    checkpointCount++;
    const statsAfter = getSheetStats(sheet);
    const newSenders = statsAfter.senderCount - statsBefore.senderCount;
    const updatedSenders = Object.keys(runMap).length - newSenders;
    const newEmails = statsAfter.totalEmails - statsBefore.totalEmails;
    console.log(`Final checkpoint #${checkpointCount}: cursor=${cursor} | batch: +${newSenders} new, ${updatedSenders} updated, +${newEmails} emails | TOTALS: ${statsAfter.senderCount} senders, ${statsAfter.totalEmails} emails`);
  }

  // Check if we're done
  const remaining = GmailApp.search(CONFIG.INBOX_QUERY, cursor, 1);
  const done = remaining.length === 0;
  
  // Get final stats for alert
  const finalStats = getSheetStats(sheet);

  if (done) {
    props.deleteProperty(KEYS.CURSOR);
    props.deleteProperty(KEYS.INITIALIZED);
    props.deleteProperty(KEYS.TOTAL_PROCESSED);
    safeAlert(`✅ Scan complete!\n\nThreads processed: ${totalProcessed}\nUnique senders: ${finalStats.senderCount}\nTotal emails counted: ${finalStats.totalEmails}\n\nYou can now sort/filter the sheet or run actions.`);
  } else {
    props.setProperty(KEYS.CURSOR, String(cursor));
    props.setProperty(KEYS.TOTAL_PROCESSED, String(totalProcessed));
    safeAlert(`⏸️ Progress saved.\n\nThreads processed: ${totalProcessed}\nUnique senders: ${finalStats.senderCount}\nTotal emails counted: ${finalStats.totalEmails}\nCursor at offset: ${cursor}\n\nRun "Start/Resume Scan" again to continue.`);
  }
}

/**
 * Process messages from a thread into the runMap
 * @param {GoogleAppsScript.Gmail.GmailMessage[]} msgs 
 * @param {Record<string, SenderBucket>} runMap 
 */
function processMessages(msgs, runMap) {
  const tz = CONFIG.TIMEZONE;

  for (const m of msgs) {
    // Only count messages actually in inbox
    if (typeof m.isInInbox === 'function' && !m.isInInbox()) continue;

    const rawFrom = m.getFrom() || '';
    const { email, name } = parseFrom(rawFrom);
    if (!email) continue;

    const subject = m.getSubject() || '';
    const dateIso = Utilities.formatDate(m.getDate(), tz, "yyyy-MM-dd'T'HH:mm:ssXXX");
    const unsubHeader = (m.getHeader && typeof m.getHeader === 'function') 
      ? (m.getHeader('List-Unsubscribe') || '') 
      : '';
    const unsub = parseListUnsubscribeHeader(unsubHeader);

    const key = email.toLowerCase();
    if (!runMap[key]) {
      runMap[key] = {
        email,
        name: name || null,
        count: 0,
        sampleSubject: null,
        lastSeen: null,
        unsubUrl: null,
        unsubMailto: null,
      };
    }

    const bucket = runMap[key];
    bucket.count += 1;
    if (!bucket.sampleSubject && subject) bucket.sampleSubject = subject;
    if (!bucket.lastSeen || dateIso > bucket.lastSeen) bucket.lastSeen = dateIso;
    if (unsub) {
      if (unsub.url && !bucket.unsubUrl) bucket.unsubUrl = unsub.url;
      if (unsub.mailto && !bucket.unsubMailto) bucket.unsubMailto = unsub.mailto;
    }
  }
}

/**
 * Checkpoint: merge runMap to sheet, update index, save cursor
 */
function checkpoint(sheet, indexSheet, props, runMap, cursor, totalProcessed) {
  // Merge data
  mergeRunMapUsingIndex(sheet, indexSheet, runMap);
  
  // Save cursor (this is the critical part - cursor reflects what's persisted)
  props.setProperty(KEYS.CURSOR, String(cursor));
  props.setProperty(KEYS.TOTAL_PROCESSED, String(totalProcessed));
  
  // Force flush (Apps Script batches writes, this ensures they're committed)
  SpreadsheetApp.flush();
}

/**
 * Get sanity check stats from the sheet
 * @returns {{senderCount: number, totalEmails: number}}
 */
function getSheetStats(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { senderCount: 0, totalEmails: 0 };
  }
  
  // Count column is D (column 4)
  const countRange = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  let totalEmails = 0;
  let senderCount = 0;
  
  for (const row of countRange) {
    const count = Number(row[0] || 0);
    if (count > 0) {
      senderCount++;
      totalEmails += count;
    }
  }
  
  return { senderCount, totalEmails };
}

/**
 * Migrate existing sheet to add Processed column if missing
 */
function migrateAddProcessedColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Check if Processed column already exists
  if (headers.includes('Processed')) {
    return; // Already migrated
  }
  
  // Add Processed header in column J (10)
  const nextCol = headers.length + 1;
  sheet.getRange(1, nextCol).setValue('Processed');
  sheet.getRange(nextCol + ':' + nextCol).setNumberFormat('@');
  
  console.log('Migrated: Added Processed column');
}

/* ================== Reset ================== */

function resetScan() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  // Check if there's data to lose
  if (sheet && sheet.getLastRow() > 1) {
    const stats = getSheetStats(sheet);
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚠️ Confirm Reset',
      `This will DELETE all existing data:\n\n• ${stats.senderCount} senders\n• ${stats.totalEmails} emails\n\nAre you sure?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      safeAlert('Reset cancelled.');
      return;
    }
  }

  if (sheet) ss.deleteSheet(sheet);
  const newSheet = ss.insertSheet(CONFIG.SHEET_NAME);

  let indexSheet = ss.getSheetByName(CONFIG.INDEX_SHEET_NAME);
  if (indexSheet) ss.deleteSheet(indexSheet);
  indexSheet = ss.insertSheet(CONFIG.INDEX_SHEET_NAME);
  indexSheet.hideSheet();

  // Force setup (these functions have safety checks, but we just deleted the sheets so they're empty)
  const headers = [
    'Act?', 'Sender Email', 'Sender Name', 'Count', 'Example Subject',
    'Last Seen (ISO)', 'Unsubscribe URL', 'Unsubscribe Mailto', 'Notes / Status', 'Processed'
  ];
  newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  newSheet.setFrozenRows(1);
  newSheet.autoResizeColumns(1, headers.length);
  newSheet.getRange('D:D').setNumberFormat('0');
  newSheet.getRange('F:F').setNumberFormat('@');
  newSheet.getRange('J:J').setNumberFormat('@'); // Processed timestamp
  if (!newSheet.getFilter()) newSheet.getDataRange().createFilter();
  
  indexSheet.getRange(1, 1, 1, 2).setValues([['email_lc', 'row']]);
  indexSheet.setFrozenRows(1);

  getUserProps().deleteAllProperties();
  safeAlert('🔄 Scan reset. Use "Start/Resume Scan" to begin again.');
}

/* ================== Sheet Helpers ================== */

function setupHeader(sheet) {
  // Safety check: never wipe existing data
  if (sheet.getLastRow() > 1) {
    console.log('setupHeader called but sheet has data — skipping to preserve data');
    return;
  }
  
  const headers = [
    'Act?', 'Sender Email', 'Sender Name', 'Count', 'Example Subject',
    'Last Seen (ISO)', 'Unsubscribe URL', 'Unsubscribe Mailto', 'Notes / Status', 'Processed'
  ];
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  // NOTE: Checkboxes are added per-row in mergeRunMapUsingIndex, not here
  sheet.autoResizeColumns(1, headers.length);
  sheet.getRange('D:D').setNumberFormat('0');
  sheet.getRange('F:F').setNumberFormat('@');
  sheet.getRange('J:J').setNumberFormat('@'); // Processed timestamp as text
  if (!sheet.getFilter()) sheet.getDataRange().createFilter();
}

function resetIndexSheet(indexSheet) {
  // Safety check: never wipe existing data
  if (indexSheet.getLastRow() > 1) {
    console.log('resetIndexSheet called but index has data — skipping to preserve data');
    return;
  }
  
  indexSheet.clear();
  indexSheet.getRange(1, 1, 1, 2).setValues([['email_lc', 'row']]);
  indexSheet.setFrozenRows(1);
}

/**
 * Read index into memory
 * @returns {{map: Record<string,number>}}
 */
function readIndex(indexSheet) {
  const values = indexSheet.getDataRange().getValues();
  /** @type {Record<string,number>} */
  const m = {};
  for (let i = 1; i < values.length; i++) {
    const emailLc = String(values[i][0] || '').trim().toLowerCase();
    const row = Number(values[i][1] || 0);
    if (emailLc && row > 1) m[emailLc] = row;
  }
  return { map: m };
}

/**
 * Merge runMap into sheet using index for O(1) lookups
 */
function mergeRunMapUsingIndex(sheet, indexSheet, runMap) {
  if (!runMap || Object.keys(runMap).length === 0) return;

  const idx = readIndex(indexSheet);
  const indexMap = idx.map;

  const newRows = [];
  const newIndexRows = [];
  const existingUpdates = []; // {row, values}

  // Separate new vs existing
  for (const key of Object.keys(runMap)) {
    const bucket = runMap[key];
    
    if (indexMap.hasOwnProperty(key)) {
      // Existing sender - queue for update
      existingUpdates.push({ key, row: indexMap[key], bucket });
    } else {
      // New sender - queue for append
      newRows.push({ key, row: buildRowFromBucket(bucket) });
    }
  }

  // Batch append new rows
  if (newRows.length > 0) {
    const firstNewRow = sheet.getLastRow() + 1;
    const rowData = newRows.map(r => r.row);
    sheet.getRange(firstNewRow, 1, rowData.length, rowData[0].length).setValues(rowData);

    // Add checkboxes only for the newly added rows
    sheet.getRange(firstNewRow, 1, rowData.length, 1).insertCheckboxes();

    // Build index entries
    for (let i = 0; i < newRows.length; i++) {
      const rowNum = firstNewRow + i;
      newIndexRows.push([newRows[i].key, rowNum]);
      indexMap[newRows[i].key] = rowNum; // Update local cache too
    }

    // Append to index sheet
    const indexStart = indexSheet.getLastRow() + 1;
    indexSheet.getRange(indexStart, 1, newIndexRows.length, 2).setValues(newIndexRows);
  }

  // Update existing rows (batch read, then batch write)
  if (existingUpdates.length > 0) {
    for (const { key, row, bucket } of existingUpdates) {
      if (row < 2) continue;

      const vals = sheet.getRange(row, 1, 1, 10).getValues()[0];
      // Cols: 0=Act?, 1=Email, 2=Name, 3=Count, 4=Subject, 5=LastSeen, 6=URL, 7=Mailto, 8=Notes, 9=Processed

      const newCount = Number(vals[3] || 0) + (bucket.count || 0);
      const subject = String(vals[4] || '') || (bucket.sampleSubject || '');
      const existingLast = String(vals[5] || '');
      const lastSeen = (!existingLast || (bucket.lastSeen && bucket.lastSeen > existingLast))
        ? (bucket.lastSeen || existingLast)
        : existingLast;
      const unsubUrl = String(vals[6] || '') || (bucket.unsubUrl || '');
      const unsubMailto = String(vals[7] || '') || (bucket.unsubMailto || '');
      const name = String(vals[2] || '') || (bucket.name || '');

      const newRowVals = [
        !!vals[0],                          // Keep Act? checkbox state
        String(vals[1] || bucket.email),    // Keep existing email
        name,
        newCount,
        subject,
        lastSeen,
        unsubUrl,
        unsubMailto,
        vals[8] || '',                      // Keep existing notes
        vals[9] || '',                      // Keep existing processed timestamp
      ];

      sheet.getRange(row, 1, 1, 10).setValues([newRowVals]);
    }
  }

  // Re-sort by Count desc
  try {
    const filter = sheet.getFilter() || sheet.getDataRange().createFilter();
    filter.sort(4, false);
  } catch (_) {}
}

function buildRowFromBucket(b) {
  return [
    false,                    // A: Act?
    b.email,                  // B: Sender Email
    b.name || '',             // C: Sender Name
    b.count || 0,             // D: Count
    b.sampleSubject || '',    // E: Example Subject
    b.lastSeen || '',         // F: Last Seen (ISO)
    b.unsubUrl || '',         // G: Unsubscribe URL
    b.unsubMailto || '',      // H: Unsubscribe Mailto
    '',                       // I: Notes / Status
    ''                        // J: Processed
  ];
}

/* ================== Actions ================== */

/**
 * Unsubscribe only (no delete) for checked rows
 */
function actionUnsubscribeOnly() {
  processCheckedRows({ unsubscribe: true, delete: false });
}

/**
 * Delete all messages from sender for checked rows (no unsubscribe)
 */
function actionDeleteOnly() {
  processCheckedRows({ unsubscribe: false, delete: true });
}

/**
 * Unsubscribe AND delete for checked rows
 */
function actionUnsubscribeAndDelete() {
  processCheckedRows({ unsubscribe: true, delete: true });
}

/**
 * Core action processor
 * @param {{unsubscribe: boolean, delete: boolean}} options
 */
function processCheckedRows(options) {
  const sheet = getSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) { 
    safeAlert('Sheet not found. Run "Start/Resume Scan" first.'); 
    return; 
  }
  
  // Migrate: add Processed column if missing
  migrateAddProcessedColumn(sheet);
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) { 
    safeAlert('No rows to process.'); 
    return; 
  }

  // Count checked rows first
  let checkedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (!!data[i][0] && (data[i][1] || '').toString().trim()) checkedCount++;
  }

  if (checkedCount === 0) {
    safeAlert('No rows checked. Check the "Act?" box for rows you want to process.');
    return;
  }

  // Build action description
  const actions = [];
  if (options.unsubscribe) actions.push('unsubscribe');
  if (options.delete) actions.push('delete all messages');
  const actionDesc = actions.join(' & ');

  console.log(`=== Starting ${actionDesc.toUpperCase()} for ${checkedCount} checked rows ===`);

  let acted = 0;
  let errors = 0;
  let totalEmailsDeleted = 0;
  let unsubSuccess = 0;
  let unsubFailed = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const act = !!row[0];
    const email = (row[1] || '').toString().trim();
    const senderName = (row[2] || '').toString().trim();
    const unsubUrl = (row[6] || '').toString().trim();
    const unsubMailto = (row[7] || '').toString().trim();
    const statusCol = 9;     // Column I: Notes / Status
    const processedCol = 10; // Column J: Processed
    const actCol = 1;        // Column A: Act?

    if (!act || !email) continue;

    const notes = [];
    const logParts = [];
    
    try {
      if (CONFIG.DRY_RUN) {
        if (options.unsubscribe) {
          const method = unsubUrl ? 'URL' : unsubMailto ? 'mailto' : 'none';
          notes.push(`[DRY] Would unsub via ${method}`);
          logParts.push(`unsub: DRY (${method})`);
        }
        if (options.delete) {
          notes.push(`[DRY] Would delete all`);
          logParts.push(`delete: DRY`);
        }
      } else {
        // Unsubscribe
        if (options.unsubscribe) {
          const unsubResult = unsubscribeIfPossible(unsubUrl, unsubMailto);
          notes.push(unsubResult);
          
          const isSuccess = unsubResult.includes('OK') || unsubResult.includes('sent');
          if (isSuccess) {
            unsubSuccess++;
            logParts.push(`unsub: ✓`);
          } else {
            unsubFailed++;
            logParts.push(`unsub: ✗ (${unsubResult})`);
          }
        }
        
        // Delete
        if (options.delete) {
          const deleteResult = deleteAllFromSender(email);
          notes.push(deleteResult);
          
          // Extract count from result like "Trashed 42 thread(s)"
          const countMatch = deleteResult.match(/(\d+)\s+thread/);
          const deletedCount = countMatch ? parseInt(countMatch[1], 10) : 0;
          totalEmailsDeleted += deletedCount;
          logParts.push(`deleted: ${deletedCount} threads`);
        }
      }
      
      // Update Notes column
      sheet.getRange(i + 1, statusCol).setValue(notes.join(' | '));
      
      // Uncheck Act? and set Processed timestamp (skip for dry run)
      if (!CONFIG.DRY_RUN) {
        const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
        sheet.getRange(i + 1, actCol).setValue(false);
        sheet.getRange(i + 1, processedCol).setValue(timestamp);
      }
      
      acted++;
      
      // Log this entry
      const displayName = senderName ? `${senderName} <${email}>` : email;
      console.log(`[${acted}/${checkedCount}] ${displayName} → ${logParts.join(', ')}`);
      
    } catch (e) {
      sheet.getRange(i + 1, statusCol).setValue(`ERROR: ${e?.message || e}`);
      errors++;
      console.log(`[${acted + errors}/${checkedCount}] ${email} → ERROR: ${e?.message || e}`);
    }
  }

  // Summary log
  console.log(`=== COMPLETE ===`);
  console.log(`Processed: ${acted}, Errors: ${errors}`);
  if (options.unsubscribe) {
    console.log(`Unsubscribe: ${unsubSuccess} success, ${unsubFailed} failed/none`);
  }
  if (options.delete) {
    console.log(`Deleted: ${totalEmailsDeleted} total threads`);
  }

  // Summary alert
  const dryPrefix = CONFIG.DRY_RUN ? '[DRY RUN] ' : '';
  let alertMsg = `${dryPrefix}${actionDesc.toUpperCase()}\n\nProcessed: ${acted} row(s)\nErrors: ${errors}`;
  if (options.unsubscribe && !CONFIG.DRY_RUN) {
    alertMsg += `\n\nUnsubscribe:\n• Success: ${unsubSuccess}\n• Failed/None: ${unsubFailed}`;
  }
  if (options.delete && !CONFIG.DRY_RUN) {
    alertMsg += `\n\nDeleted: ${totalEmailsDeleted} threads`;
  }
  
  safeAlert(alertMsg);
}

/* ================== Utilities ================== */

function parseFrom(rawFrom) {
  let email = '';
  let name = '';

  const angle = rawFrom.match(/<([^>]+)>/);
  if (angle) {
    email = angle[1].trim();
    name = rawFrom.replace(angle[0], '').trim().replace(/^"|"$/g, '');
  } else {
    const bare = rawFrom.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    if (bare) {
      email = bare[0].trim();
      name = rawFrom.replace(bare[0], '').trim().replace(/^[("'\s]+|[)"'\s]+$/g, '');
    }
  }

  if (email) email = email.toLowerCase();
  if (name) name = name.replace(/^"|"$/g, '');
  return { email, name: name || null };
}

function parseListUnsubscribeHeader(headerValue) {
  if (!headerValue) return null;
  const parts = headerValue.split(',').map(s => s.trim().replace(/^<|>$/g, ''));
  let url = null;
  let mailto = null;
  for (const p of parts) {
    if (/^https?:\/\//i.test(p)) url = url || p;
    else if (/^mailto:/i.test(p)) mailto = mailto || p;
  }
  return (url || mailto) ? { url, mailto } : null;
}

function unsubscribeIfPossible(unsubUrl, unsubMailto) {
  if (unsubUrl) {
    try {
      const resp = UrlFetchApp.fetch(unsubUrl, { 
        muteHttpExceptions: true, 
        followRedirects: true 
      });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) return `Unsub URL OK (${code})`;
      
      const resp2 = UrlFetchApp.fetch(unsubUrl, {
        method: 'post',
        payload: {},
        muteHttpExceptions: true,
        followRedirects: true,
      });
      const code2 = resp2.getResponseCode();
      if (code2 >= 200 && code2 < 300) return `Unsub URL POST OK (${code2})`;
      return `Unsub URL non-2xx (${code}/${code2})`;
    } catch (_) {
      // Fall through to mailto
    }
  }

  if (unsubMailto) {
    try {
      const m = unsubMailto.match(/^mailto:([^?]+)(\?(.*))?$/i);
      if (!m) return 'Invalid mailto';
      const toAddr = decodeURIComponent(m[1]);
      const params = new URLSearchParams(m[3] || '');
      const subject = params.get('subject') || 'unsubscribe';
      const body = params.get('body') || 'Please unsubscribe me.';
      GmailApp.sendEmail(toAddr, subject, body);
      return `Unsub mailto sent to ${toAddr}`;
    } catch (e) {
      return `Mailto failed: ${e?.message || e}`;
    }
  }

  return 'No unsub info';
}

function deleteAllFromSender(senderEmail) {
  // Search ALL mail from this sender (not just inbox)
  const query = `from:"${senderEmail}"`;
  let start = 0;
  const page = 100;
  let total = 0;

  while (true) {
    const threads = GmailApp.search(query, start, page);
    if (!threads.length) break;
    if (!CONFIG.DRY_RUN) {
      threads.forEach(th => th.moveToTrash());
    }
    total += threads.length;
    start += page;
  }
  return `${CONFIG.DRY_RUN ? '[DRY] ' : ''}Trashed ${total} thread(s)`;
}