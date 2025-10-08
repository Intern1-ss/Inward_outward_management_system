// Configuration
const CONFIG = {
  INWARD_SHEET: 'Inward',
  OUTWARD_SHEET: 'Outward',
  CONFIRMATION_SHEET: 'Confirmations',
  LINKS_SHEET: 'Entry_Links',
  HEADER_ROW: 1,
  BOSS_EMAIL:'sathyajain9@gmail.com',
  NOTIFICATION_SUBJECT:"Inward/Outward Pending Report",
  TRIGGER_FUNCTION_NAME:'sendWeeklyPendingReport'
};


// =====================================================
// WEB APP ENTRY POINT
// =====================================================

function doGet(e) {
  try {
    setupSheets();
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Document Management System')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput(`
      <html><body>
        <h1>Error Loading Application</h1>
        <p>Error: ${error.toString()}</p>
        <p>Please contact the administrator.</p>
        <button onclick="window.location.reload()">Try Again</button>
      </body></html>
    `);
  }
}

// =====================================================
// ENHANCED SHEET SETUP FUNCTIONS
// =====================================================

function setupSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create Inward sheet if it doesn't exist
    if (!ss.getSheetByName(CONFIG.INWARD_SHEET)) {
      const inwardSheet = ss.insertSheet(CONFIG.INWARD_SHEET);
      const inwardHeaders = [
        'Sl. No', 'Means', 'Inward No', 'From Whom', 'Subject', 
        'Taken By', 'Date & Time', 'Action Taken', 'File Reference', 'Postal Tariff'
      ];
      inwardSheet.getRange(1, 1, 1, inwardHeaders.length).setValues([inwardHeaders]);
      Logger.log('Created Inward sheet');
    }
    
    // Create Outward sheet if it doesn't exist
    if (!ss.getSheetByName(CONFIG.OUTWARD_SHEET)) {
      const outwardSheet = ss.insertSheet(CONFIG.OUTWARD_SHEET);
      const outwardHeaders = [
      'Sl. No', 'Means', 'Outward No', 'To Whom', 'Subject', 
      'Sent By', 'Date & Time', 'Case Closed', 'File Reference', 'Postal Tariff', 'Due Date'];
      outwardSheet.getRange(1, 1, 1, outwardHeaders.length).setValues([outwardHeaders]);
      Logger.log('Created Outward sheet');
    }
    
    // Create Confirmations sheet if it doesn't exist
    if (!ss.getSheetByName(CONFIG.CONFIRMATION_SHEET)) {
      const confirmationSheet = ss.insertSheet(CONFIG.CONFIRMATION_SHEET);
      const confirmHeaders = [
        'Date', 'User Email', 'Sheet Name', 'Row Number', 'Entry ID', 'Status', 'Confirmation Note', 'Action Type'
      ];
      confirmationSheet.getRange(1, 1, 1, confirmHeaders.length).setValues([confirmHeaders]);
      Logger.log('Created Confirmations sheet');
    }
    
    // Create Entry Links sheet with UUID support
    if (!ss.getSheetByName(CONFIG.LINKS_SHEET)) {
      const linksSheet = ss.insertSheet(CONFIG.LINKS_SHEET);
      const linkHeaders = [
        'Link ID', 'Primary Entry ID', 'Linked Entry ID', 'Link Type', 'Created Date', 'Created By', 'Notes', 'Link Group UUID'
      ];
      linksSheet.getRange(1, 1, 1, linkHeaders.length).setValues([linkHeaders]);
      Logger.log('Created Entry Links sheet with UUID support');
    } else {
      // Check if UUID column exists, if not add it
      const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
      const lastColumn = linksSheet.getLastColumn();
      const headers = linksSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      
      Logger.log(`Links sheet headers: ${JSON.stringify(headers)}`);
      
      if (!headers.includes('Link Group UUID')) {
        const newColumn = lastColumn + 1;
        linksSheet.getRange(1, newColumn).setValue('Link Group UUID');
        Logger.log(`Added Link Group UUID column at position ${newColumn}`);
      }
    }
    
    return true;
  } catch (error) {
    Logger.log('Setup error: ' + error.toString());
    return false;
  }
}

// =====================================================
// USER FUNCTIONS
// =====================================================

function getCurrentUser() {
  try {
    let userEmail = '';
    let errorDetails = '';
    
    try {
      userEmail = Session.getActiveUser().getEmail();
      Logger.log(`Got active user email: "${userEmail}"`);
    } catch (e) {
      errorDetails += `Active user failed: ${e.toString()}; `;
      try { 
        userEmail = Session.getEffectiveUser().getEmail(); 
        Logger.log(`Got effective user email: "${userEmail}"`);
      } catch (e2) { 
        errorDetails += `Effective user failed: ${e2.toString()}; `;
        userEmail = ''; 
      }
    }

    userEmail = (userEmail || '').toString().trim();
    const isAdmin = userEmail ? CONFIG.ADMIN_USERS.includes(userEmail.toLowerCase()) : false;
    
    Logger.log(`Final user info - Email: "${userEmail}", IsAdmin: ${isAdmin}`);
    if (errorDetails) {
      Logger.log(`User detection errors: ${errorDetails}`);
    }

    return {
      success: true,
      userEmail: userEmail,
      isAdmin: isAdmin,
      message: userEmail ? 'User info retrieved' : 'No user email detected'
    };

  } catch (error) {
    Logger.log('Error in getCurrentUser: ' + error.toString());
    return { 
      success: true,
      userEmail: '',
      isAdmin: false,
      message: 'Using default user due to error'
    };
  }
}

function isUserAdmin(userEmail = null) {
  try {
    let emailToCheck = userEmail;
    
    if (!emailToCheck) {
      const userInfo = getCurrentUser();
      emailToCheck = userInfo.userEmail;
    }
    
    if (!emailToCheck) {
      Logger.log('isUserAdmin: No email to check');
      return false;
    }
    
    const isAdmin = CONFIG.ADMIN_USERS.includes(emailToCheck.toLowerCase());
    Logger.log(`isUserAdmin: "${emailToCheck}" is admin: ${isAdmin}`);
    
    return isAdmin;
  } catch (error) {
    Logger.log(`Error checking admin status: ${error.toString()}`);
    return false;
  }
}

// =====================================================
// INITIAL DATA LOADING
// =====================================================

function getInitialData() {
  try {
    const userInfo = getCurrentUser();
    
    Logger.log(`Getting initial data for: ${userInfo.userEmail}, Admin: ${userInfo.isAdmin}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets();

    const stats = calculateUserStats(ss, userInfo.userEmail, userInfo.isAdmin);

    return {
      success: true,
      user: {
        email: userInfo.userEmail,
        isAdmin: userInfo.isAdmin
      },
      stats: stats,
      message: 'Initial data loaded successfully'
    };

  } catch (error) {
    Logger.log('Error in getInitialData: ' + error.toString());
    return {
      success: false,
      message: 'Error loading initial data: ' + error.toString(),
      user: {
        isAdmin: false
      },
      stats: { ready: 0, confirmed: 0 }
    };
  }
}

// =====================================================
// ENHANCED ENTRIES FUNCTIONS
// =====================================================

function getSheetEntriesWithDetailsDebug(ss, sheetName, userEmail, isAdmin, confirmations, links) {
  const entries = [];
  
  try {
    Logger.log(`--- Processing ${sheetName} Sheet ---`);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`ERROR: ${sheetName} sheet not found`);
      return entries;
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    Logger.log(`${sheetName}: ${lastRow} rows, ${lastColumn} columns`);
    
    if (lastRow <= 1) {
      Logger.log(`${sheetName}: No data rows found`);
      return entries;
    }
    
    // Read all data at once
    let data;
    try {
      data = sheet.getDataRange().getValues();
      Logger.log(`${sheetName}: Successfully read ${data.length} rows of data`);
    } catch (readError) {
      Logger.log(`${sheetName}: Error reading data - ${readError.toString()}`);
      return entries;
    }
    
    // Log headers for verification
    if (data.length > 0) {
      Logger.log(`${sheetName} Headers: [${data[0].slice(0, 10).map(h => `"${h}"`).join(', ')}]`);
    }
    
    const userEmailLower = (userEmail || '').toLowerCase();
    Logger.log(`${sheetName}: Looking for user "${userEmailLower}", Admin: ${isAdmin}`);
    
    let processedRows = 0;
    let validRows = 0;
    let skippedRows = 0;
    
    for (let i = 1; i < data.length; i++) {
      processedRows++;
      const row = data[i];
      
      // Log first few rows for debugging
      if (i <= 3) {
        Logger.log(`${sheetName} Row ${i + 1}: [${row.slice(0, 7).map(cell => `"${cell}"`).join(', ')}]`);
      }
      
      // Check if row has any data at all
      const hasAnyData = row.some(cell => cell !== null && cell !== undefined && cell !== '');
      if (!hasAnyData) {
        if (i <= 3) Logger.log(`${sheetName} Row ${i + 1}: Skipping empty row`);
        continue;
      }
      
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      if (i <= 3) {
        Logger.log(`${sheetName} Row ${i + 1}: User email = "${rowUserEmail}"`);
      }
      
      // More permissive permission check
      let shouldInclude = false;
      if (isAdmin) {
        shouldInclude = true;
      } else if (!userEmailLower) {
        // No user email detected - include all entries for debugging
        shouldInclude = true;
      } else if (rowUserEmail === userEmailLower) {
        shouldInclude = true;
      } else if (!rowUserEmail) {
        // Include entries with no user email (legacy data)
        shouldInclude = true;
      }
      
      if (!shouldInclude) {
        skippedRows++;
        if (i <= 3) Logger.log(`${sheetName} Row ${i + 1}: Skipped due to permissions`);
        continue;
      }
      
      validRows++;
      
      const entryId = `${sheetName}-${i + 1}`;
      
      // Check if entry has basic required data
      const hasBasicData = !!(row[1] || row[3] || row[4]); // Means, Person, Subject
      const isComplete = !!(row[1] && row[3] && row[4] && row[6]); // + Date Time
      
      if (i <= 3) {
        Logger.log(`${sheetName} Row ${i + 1}: Basic data = ${hasBasicData}, Complete = ${isComplete}`);
      }
      
      // Skip rows with no basic data
      if (!hasBasicData) {
        if (i <= 3) Logger.log(`${sheetName} Row ${i + 1}: Skipping - no basic data`);
        continue;
      }
      
      // Check if entry is confirmed
      const confirmKey = `${sheetName}-${i + 1}`;
      const isConfirmed = confirmations.has(confirmKey);
      
      // Get linked entries
      let linkedEntries = [];
      try {
        linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
      } catch (linkError) {
        Logger.log(`${sheetName} Row ${i + 1}: Error getting links - ${linkError.toString()}`);
      }
      
      // Format date safely
      let formattedDateTime = '';
      try {
        if (row[6]) {
          formattedDateTime = formatDateTime(row[6]);
        }
      } catch (dateError) {
        Logger.log(`${sheetName} Row ${i + 1}: Date format error - ${dateError.toString()}`);
        formattedDateTime = row[6] ? row[6].toString() : '';
      }
      
      const entry = {
        id: entryId,
        type: sheetName,
        subject: (row[4] || '').toString(),
        person: (row[3] || '').toString(),
        user: (row[5] || '').toString(),
        dateTime: formattedDateTime,
        means: (row[1] || '').toString(),
        fileReference: (row[8] || '').toString(),
        postalTariff: row[9] || '',
        complete: isComplete,
        confirmed: isConfirmed,
        linkedEntries: linkedEntries,
        // Additional type-specific fields
        ...(sheetName === 'Inward' ? {
          inwardNo: (row[2] || '').toString(),
          fromWhom: (row[3] || '').toString(),
          actionTaken: (row[7] || '').toString()
        } : {
          outwardNo: (row[2] || '').toString(),
          toWhom: (row[3] || '').toString(),
          caseClosed: (row[7] || 'No').toString()
        })
      };
      
      entries.push(entry);
      
      if (i <= 3) {
        Logger.log(`${sheetName} Row ${i + 1}: Entry created - "${entry.subject}"`);
      }
    }
    
    Logger.log(`${sheetName} SUMMARY:`);
    Logger.log(`- Processed rows: ${processedRows}`);
    Logger.log(`- Valid rows: ${validRows}`);
    Logger.log(`- Skipped rows: ${skippedRows}`);
    Logger.log(`- Final entries: ${entries.length}`);
    
  } catch (error) {
    Logger.log(`Error getting entries from ${sheetName}: ${error.toString()}`);
    Logger.log(`Error stack: ${error.stack}`);
  }
  
  return entries;
}

function loadEntriesFromSheet(ss, sheetName, userEmail, isAdmin) {
  const entries = [];
  
  try {
    Logger.log(`--- Processing ${sheetName} Sheet ---`);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`ERROR: ${sheetName} sheet not found`);
      return entries;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`${sheetName}: ${lastRow} rows`);
    
    if (lastRow <= 1) {
      Logger.log(`${sheetName}: No data rows found`);
      return entries;
    }
    
    // Read all data
    const data = sheet.getDataRange().getValues();
    Logger.log(`${sheetName}: Successfully read ${data.length} rows of data`);
    
    // Log headers
    if (data.length > 0) {
      Logger.log(`${sheetName} Headers: [${data[0].slice(0, 10).map(h => `"${h}"`).join(', ')}]`);
    }
    
    const userEmailLower = (userEmail || '').toLowerCase();
    Logger.log(`${sheetName}: Looking for user "${userEmailLower}", Admin: ${isAdmin}`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip completely empty rows
      const hasAnyData = row.some(cell => cell !== null && cell !== undefined && cell !== '');
      if (!hasAnyData) {
        continue;
      }
      
      // Log first few rows
      if (i <= 3) {
        Logger.log(`${sheetName} Row ${i + 1}: [${row.slice(0, 7).map(cell => `"${cell}"`).join(', ')}]`);
      }
      
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // SIMPLIFIED PERMISSION CHECK - be more permissive
      let shouldInclude = true; // Default to include
      
      // Only filter if we have both user emails and user is not admin
      if (!isAdmin && userEmailLower && rowUserEmail && rowUserEmail !== userEmailLower) {
        shouldInclude = false;
      }
      
      if (!shouldInclude) {
        continue;
      }
      
      // Must have at least subject to be valid
      if (!row[4]) {
        continue;
      }
      
      const entryId = `${sheetName}-${i + 1}`;
      
      // Check completeness
      const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
      
      // Format date safely
      let formattedDateTime = '';
      try {
        if (row[6]) {
          if (row[6] instanceof Date) {
            formattedDateTime = row[6].toLocaleString('en-US', {
              year: 'numeric',
              month: 'short',
              day: '2-digit',
              hour: '2-digit',
              minute: '2-digit'
            });
          } else {
            formattedDateTime = new Date(row[6]).toLocaleString('en-US', {
              year: 'numeric',
              month: 'short',
              day: '2-digit',
              hour: '2-digit',
              minute: '2-digit'
            });
          }
        }
      } catch (dateError) {
        Logger.log(`${sheetName} Row ${i + 1}: Date format error - ${dateError.toString()}`);
        formattedDateTime = row[6] ? row[6].toString() : '';
      }
      
      const entry = {
        id: entryId,
        type: sheetName,
        subject: (row[4] || '').toString(),
        person: (row[3] || '').toString(),
        user: (row[5] || '').toString(),
        dateTime: formattedDateTime,
        means: (row[1] || '').toString(),
        fileReference: (row[8] || '').toString(),
        postalTariff: row[9] || '',
        complete: isComplete,
        confirmed: false, // Simplified - skip confirmation checking for now
        linkedEntries: [], // Simplified - skip link checking for now
        // Type-specific fields
        ...(sheetName === 'Inward' ? {
          inwardNo: (row[2] || '').toString(),
          fromWhom: (row[3] || '').toString(),
          actionTaken: (row[7] || '').toString()
        } : {
          outwardNo: (row[2] || '').toString(),
          toWhom: (row[3] || '').toString(),
          caseClosed: (row[7] || 'No').toString()
        })
      };
      
      entries.push(entry);
      
      if (i <= 3) {
        Logger.log(`${sheetName} Row ${i + 1}: Entry created - "${entry.subject}"`);
      }
    }
    
    Logger.log(`${sheetName} Final entries: ${entries.length}`);
    
  } catch (error) {
    Logger.log(`Error loading entries from ${sheetName}: ${error.toString()}`);
    Logger.log(`Error stack: ${error.stack}`);
  }
  
  return entries;
}

function loadEntriesWithStatus(ss, sheetName, entryType, userEmail, isAdmin, confirmations, links) {
  const entries = [];
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      return entries;
    }
    
    const data = sheet.getDataRange().getValues();
    const userEmailLower = (userEmail || '').toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Must have at least subject to be valid
      if (!row[4]) continue;
      
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // Simplified permission - show entries to everyone (remove admin filtering)
      // You can add back permission filtering here if needed later
      
      const entryId = `${sheetName}-${i + 1}`;
      
      // FIXED: Strict completeness check
      const hasRequiredFields = !!(row[1] && row[3] && row[4]); // Means, Person, Subject
      const hasDateTime = !!row[6]; // Date & Time
      const hasUserInfo = !!row[5]; // Taken By / Sent By
      
      // Entry is complete ONLY if ALL required fields are present
      const isComplete = hasRequiredFields && hasDateTime && hasUserInfo;
      
      // Check if entry is confirmed using simplified key
      const confirmKey = `${sheetName}-${i + 1}`;
      const isConfirmed = confirmations.has(confirmKey);
      
      // Get linked entries
      const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
      const hasLinks = linkedEntries.length > 0;
      
      // Format date
      let formattedDateTime = '';
      try {
        if (row[6]) {
          if (row[6] instanceof Date) {
            formattedDateTime = row[6].toLocaleString('en-US', {
              year: 'numeric',
              month: 'short',
              day: '2-digit',
              hour: '2-digit',
              minute: '2-digit'
            });
          } else {
            formattedDateTime = new Date(row[6]).toLocaleString();
          }
        }
      } catch (dateError) {
        formattedDateTime = row[6] ? row[6].toString() : '';
      }
      
      const entry = {
        id: entryId,
        type: entryType,
        subject: (row[4] || '').toString(),
        person: (row[3] || '').toString(),
        user: (row[5] || '').toString(),
        dateTime: formattedDateTime,
        means: (row[1] || '').toString(),
        fileReference: (row[8] || '').toString(),
        postalTariff: row[9] || '',
        complete: isComplete, // This will now be more strict
        confirmed: isConfirmed,
        linkedEntries: linkedEntries,
        hasLinks: hasLinks,
        // Type-specific fields
        ...(entryType === 'Inward' ? {
          inwardNo: (row[2] || '').toString(),
          fromWhom: (row[3] || '').toString(),
          actionTaken: (row[7] || '').toString()
        } : {
          outwardNo: (row[2] || '').toString(),
          toWhom: (row[3] || '').toString(),
          caseClosed: (row[7] || 'No').toString()
        })
      };
      
      entries.push(entry);
    }
    
  } catch (error) {
    Logger.log(`Error loading entries from ${sheetName}: ${error.toString()}`);
  }
  
  return entries;
}

function calculateUserStatsSimplified(ss, userEmail, isAdmin) {
  let pending = 0;  // Changed from 'ready' to 'pending' for consistency
  let confirmed = 0;
  
  try {
    const confirmations = getConfirmationsSimplified(ss);
    
    // Calculate stats for both sheets
    const sheets = [CONFIG.INWARD_SHEET, CONFIG.OUTWARD_SHEET];
    
    for (const sheetName of sheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) continue;
      
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Skip entries without subject
        if (!row[4]) continue;
        
        // Check if complete (has all required fields)
        const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
        
        // Check if confirmed (physically processed)
        const confirmKey = `${sheetName}-${i + 1}`;
        const isConfirmed = confirmations.has(confirmKey);
        
        // CORRECTED LOGIC:
        if (isConfirmed) {
          // Entry is confirmed = Work Complete
          confirmed++;
        } else if (isComplete) {
          // Entry is complete but not confirmed = Pending Work
          pending++;
        }
        // Incomplete entries are not counted in either category
      }
    }
    
  } catch (error) {
    Logger.log('Error calculating stats: ' + error.toString());
  }
  
  return { pending, confirmed };  // Changed from { ready, confirmed }
}

function getEntriesWithDetails() {
  try {
    Logger.log('=== LOADING ENTRIES WITH ENHANCED STATUS ===');
    
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const confirmations = getConfirmationsSimplified(ss);
    const links = getEntryLinks(ss);
    
    Logger.log(`Loaded ${confirmations.size} confirmations and ${links.size} link entries`);
    
    let allEntries = [];
    
    // Load from both sheets
    const sheetsToProcess = [
      { name: CONFIG.INWARD_SHEET, type: 'Inward' },
      { name: CONFIG.OUTWARD_SHEET, type: 'Outward' }
    ];
    
    for (const sheetConfig of sheetsToProcess) {
      try {
        const entries = loadEntriesWithStatus(ss, sheetConfig.name, sheetConfig.type, currentUser, isAdmin, confirmations, links);
        allEntries = allEntries.concat(entries);
        Logger.log(`Loaded ${entries.length} entries from ${sheetConfig.name}`);
      } catch (error) {
        Logger.log(`Error loading ${sheetConfig.name}: ${error.toString()}`);
      }
    }
    
    // Sort by date (newest first)
    allEntries.sort((a, b) => {
      const dateA = new Date(a.dateTime || 0);
      const dateB = new Date(b.dateTime || 0);
      return dateB - dateA;
    });
    
    const stats = calculateUserStatsSimplified(ss, currentUser, isAdmin);
    
    Logger.log(`Total entries loaded: ${allEntries.length}`);
    
    return {
      success: true,
      entries: allEntries,
      stats: stats,
      message: `Loaded ${allEntries.length} entries successfully`
    };
    
  } catch (error) {
    Logger.log(`Error in getEntriesWithDetails: ${error.toString()}`);
    return {
      success: false,
      message: 'Error loading entries: ' + error.toString(),
      entries: [],
      stats: { ready: 0, confirmed: 0 }
    };
  }
}


function getSheetEntriesWithDetails(ss, sheetName, userEmail, isAdmin, confirmations, links) {
  const entries = [];
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      return entries;
    }
    
    const data = sheet.getDataRange().getValues();
    const userEmailLower = userEmail.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // Skip if not admin and entry doesn't belong to current user
      if (!isAdmin && rowUserEmail !== userEmailLower) {
        continue;
      }
      
      const entryId = `${sheetName}-${i + 1}`;
      
      // Check if entry is complete
      const isComplete = !!(row[1] && row[3] && row[4] && row[6]); // Means, From/To, Subject, Date
      
      // Check if entry is confirmed
      const confirmKey = `${sheetName}-${i + 1}`;
      const isConfirmed = confirmations.has(confirmKey);
      
      // Get linked entries
      const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
      
      const entry = {
        id: entryId,
        type: sheetName,
        subject: row[4] || '',
        person: row[3] || '', // From/To field
        user: row[5] || '',
        dateTime: row[6] ? formatDateTime(row[6]) : '',
        means: row[1] || '',
        fileReference: row[8] || '',
        postalTariff: row[9] || '',
        complete: isComplete,
        confirmed: isConfirmed,
        linkedEntries: linkedEntries,
        // Additional type-specific fields
        ...(sheetName === 'Inward' ? {
          inwardNo: row[2] || '',
          fromWhom: row[3] || '',
          actionTaken: row[7] || ''
        } : {
          outwardNo: row[2] || '',
          toWhom: row[3] || '',
          caseClosed: row[7] || 'No'
        })
      };
      
      entries.push(entry);
    }
    
  } catch (error) {
    Logger.log(`Error getting entries from ${sheetName}: ${error.toString()}`);
  }
  
  return entries;
}

function formatDateTime(date) {
  try {
    if (!date) return '';
    
    let dateObj;
    if (date instanceof Date) {
      dateObj = date;
    } else {
      dateObj = new Date(date);
    }
    
    if (isNaN(dateObj.getTime())) {
      // If date parsing fails, return as string
      return date.toString();
    }
    
    return dateObj.toLocaleString('en-US', {
      year: 'numeric',
      month: 'short',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
  } catch (error) {
    Logger.log(`Date formatting error: ${error.toString()}`);
    return date ? date.toString() : '';
  }
}

// =====================================================
// CONFIRMATION FUNCTIONS
// =====================================================

function getConfirmationsSimplified(ss) {
  const confirmations = new Map();
  
  try {
    const confirmSheet = ss.getSheetByName(CONFIG.CONFIRMATION_SHEET);
    if (!confirmSheet || confirmSheet.getLastRow() <= 1) {
      return confirmations;
    }
    
    const data = confirmSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[2] && row[3]) { // sheetName and rowNumber exist
        const key = `${row[2]}-${row[3]}`; // Format: "Inward-2"
        confirmations.set(key, {
          date: row[0],
          userEmail: row[1],
          sheetName: row[2],
          rowNumber: row[3],
          entryId: row[4],
          status: row[5],
          note: row[6],
          actionType: row[7]
        });
      }
    }
    
    Logger.log(`Loaded ${confirmations.size} confirmations (simplified)`);
    
  } catch (error) {
    Logger.log('Error loading confirmations: ' + error.toString());
  }
  
  return confirmations;
}

function confirmEntry(entryId, confirmationNote = '') {
  try {
    Logger.log(`Confirming entry: ${entryId}`);
    
    if (!entryId) {
      return { success: false, message: 'Entry ID is required' };
    }
    
    // Parse entry ID
    const parts = entryId.split('-');
    if (parts.length < 2) {
      return { success: false, message: 'Invalid entry ID format' };
    }
    
    const sheetName = parts[0];
    const rowNumber = parseInt(parts[1]);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || rowNumber > sheet.getLastRow()) {
      return { success: false, message: 'Entry not found' };
    }
    
    // Get entry data
    const entryData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    
    // Check if entry is complete
    if (!(entryData[1] && entryData[3] && entryData[4] && entryData[6])) {
      return { success: false, message: 'Entry is not complete and cannot be confirmed' };
    }
    
    // Check if already confirmed using simplified key
    const confirmKey = `${sheetName}-${rowNumber}`;
    const confirmations = getConfirmationsSimplified(ss);
    
    if (confirmations.has(confirmKey)) {
      return { success: false, message: 'Entry is already confirmed' };
    }
    
    // Get current user (with fallback)
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail || 'system-user';
    
    // Record the confirmation
    const confirmSheet = ss.getSheetByName(CONFIG.CONFIRMATION_SHEET);
    if (!confirmSheet) {
      return { success: false, message: 'Confirmation sheet not found' };
    }
    
    const confirmationRow = [
      new Date(), // Date
      currentUser, // User Email
      sheetName, // Sheet Name
      rowNumber, // Row Number
      entryId, // Entry ID
      'Confirmed', // Status
      confirmationNote || '', // Confirmation Note
      'User Confirmation' // Action Type
    ];
    
    const nextRow = confirmSheet.getLastRow() + 1;
    confirmSheet.getRange(nextRow, 1, 1, confirmationRow.length).setValues([confirmationRow]);
    
    Logger.log(`Confirmation saved for ${entryId}`);
    
    return {
      success: true,
      message: 'Entry confirmed successfully!',
      entryId: entryId,
      confirmedBy: currentUser,
      confirmedAt: new Date().toLocaleString()
    };
    
  } catch (error) {
    Logger.log(`Error confirming entry: ${error.toString()}`);
    return { success: false, message: `Error confirming entry: ${error.toString()}` };
  }
}



function getConfirmations(ss) {
  const confirmations = new Map();
  
  try {
    const confirmSheet = ss.getSheetByName(CONFIG.CONFIRMATION_SHEET);
    if (!confirmSheet || confirmSheet.getLastRow() <= 1) {
      return confirmations;
    }
    
    const data = confirmSheet.getDataRange().getValues();
    
    // Key is now sheetName-rowNumber (confirmation is for the entry itself)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // row indexes: [0]=Date, [1]=User Email, [2]=Sheet Name, [3]=Row Number, [4]=Entry ID, ...
      if (row[2] && row[3]) { // sheetName, rowNumber must exist
        const key = `${row[2]}-${row[3]}`;
        confirmations.set(key, {
          date: row[0],
          userEmail: row[1] || '',
          sheetName: row[2],
          rowNumber: row[3],
          entryId: row[4],
          status: row[5],
          note: row[6],
          actionType: row[7]
        });
      }
    }
    
  } catch (error) {
    Logger.log('Error loading confirmations: ' + error.toString());
  }
  
  return confirmations;
}

// =====================================================
// ENTRY LINKING FUNCTIONS
// =====================================================
function validateEntriesForLinkingSimplified(ss, primaryEntryId, linkedEntryIds) {
  try {
    const allEntries = [primaryEntryId, ...linkedEntryIds];
    
    Logger.log(`Validating entries: ${JSON.stringify(allEntries)}`);
    
    for (const entryId of allEntries) {
      const parts = entryId.split('-');
      if (parts.length < 2) {
        return { success: false, message: `Invalid entry ID format: ${entryId}` };
      }
      
      const sheetName = parts[0];
      const rowNumber = parseInt(parts[1]);
      
      if (!sheetName || isNaN(rowNumber)) {
        return { success: false, message: `Invalid entry ID: ${entryId}` };
      }
      
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return { success: false, message: `Sheet not found for entry: ${entryId}` };
      }
      
      if (rowNumber > sheet.getLastRow() || rowNumber < 2) {
        return { success: false, message: `Entry not found: ${entryId} (row ${rowNumber})` };
      }
      
      Logger.log(`✅ Entry ${entryId} validated successfully`);
    }
    
    return { success: true };
    
  } catch (error) {
    Logger.log(`Validation error: ${error.toString()}`);
    return { success: false, message: `Validation error: ${error.toString()}` };
  }
}
function linkEntries(primaryEntryId, linkedEntryIds) {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    
    Logger.log(`=== LINKING DEBUG ===`);
    Logger.log(`Primary Entry: ${primaryEntryId}`);
    Logger.log(`Linked Entries: ${JSON.stringify(linkedEntryIds)}`);
    Logger.log(`User: ${currentUser}`);
    
    if (!primaryEntryId || !linkedEntryIds || linkedEntryIds.length === 0) {
      return { success: false, message: 'Primary entry and linked entries are required' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ensure sheets are set up correctly
    setupSheets();
    
    const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
    if (!linksSheet) {
      return { success: false, message: 'Links sheet not found' };
    }
    
    // Validate entries exist (simplified validation)
    const validationResult = validateEntriesForLinkingSimplified(ss, primaryEntryId, linkedEntryIds);
    if (!validationResult.success) {
      Logger.log(`Validation failed: ${validationResult.message}`);
      return validationResult;
    }
    
    // Generate a unique UUID for this linking group
    const linkGroupUUID = generateUUID();
    Logger.log(`Generated UUID: ${linkGroupUUID}`);
    
    // Create links
    const linkResults = [];
    const linkDate = new Date();
    const linksToAdd = [];
    
    for (const linkedEntryId of linkedEntryIds) {
      // Create bidirectional links with UUID
      const link1Id = `LINK-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
      const link2Id = `LINK-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
      
      // Primary -> Linked
      const linkRow1 = [
        link1Id, // Link ID
        primaryEntryId, // Primary Entry ID
        linkedEntryId, // Linked Entry ID
        'Manual Link', // Link Type
        linkDate, // Created Date
        currentUser, // Created By
        `Linked by ${currentUser} - UUID: ${linkGroupUUID}`, // Notes with UUID
        linkGroupUUID // Link Group UUID
      ];
      
      // Linked -> Primary (bidirectional)
      const linkRow2 = [
        link2Id, // Link ID
        linkedEntryId, // Primary Entry ID
        primaryEntryId, // Linked Entry ID
        'Manual Link', // Link Type
        linkDate, // Created Date
        currentUser, // Created By
        `Linked by ${currentUser} - UUID: ${linkGroupUUID}`, // Notes with UUID
        linkGroupUUID // Link Group UUID
      ];
      
      linksToAdd.push(linkRow1, linkRow2);
      
      linkResults.push({
        primaryEntry: primaryEntryId,
        linkedEntry: linkedEntryId,
        linkIds: [link1Id, link2Id],
        uuid: linkGroupUUID
      });
    }
    
    // Add all links at once for better performance
    const nextRow = linksSheet.getLastRow() + 1;
    const numColumns = Math.max(8, linksSheet.getLastColumn()); // Ensure we have at least 8 columns
    
    Logger.log(`Adding ${linksToAdd.length} link rows starting at row ${nextRow}`);
    Logger.log(`Sample link row: ${JSON.stringify(linksToAdd[0])}`);
    
    // Write all links at once
    if (linksToAdd.length > 0) {
      linksSheet.getRange(nextRow, 1, linksToAdd.length, numColumns).setValues(linksToAdd);
    }
    
    Logger.log(`Successfully created ${linkResults.length * 2} links with UUID: ${linkGroupUUID}`);
    Logger.log(`=== LINKING COMPLETE ===`);
    
    return {
      success: true,
      message: `Successfully linked ${linkedEntryIds.length} entries!`,
      links: linkResults,
      linkedBy: currentUser,
      linkedAt: linkDate.toLocaleString(),
      uuid: linkGroupUUID
    };
    
  } catch (error) {
    Logger.log(`Error linking entries: ${error.toString()}`);
    Logger.log(`Error stack: ${error.stack}`);
    return { success: false, message: `Error linking entries: ${error.toString()}` };
  }
}


function validateEntriesForLinking(ss, primaryEntryId, linkedEntryIds, currentUser) {
  try {
    const allEntries = [primaryEntryId, ...linkedEntryIds];
    const currentUserLower = currentUser.toLowerCase();
    const isAdmin = isUserAdmin(currentUser);
    
    for (const entryId of allEntries) {
      const parts = entryId.split('-');
      if (parts.length < 2) {
        return { success: false, message: `Invalid entry ID format: ${entryId}` };
      }
      
      const sheetName = parts[0];
      const rowNumber = parseInt(parts[1]);
      
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        return { success: false, message: `Sheet not found for entry: ${entryId}` };
      }
      
      if (rowNumber > sheet.getLastRow()) {
        return { success: false, message: `Entry not found: ${entryId}` };
      }
      
      // Check permission
      const entryData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
      const entryUserEmail = entryData[5] ? entryData[5].toString().toLowerCase() : '';
      
      if (!isAdmin && entryUserEmail !== currentUserLower) {
        return { success: false, message: `You can only link your own entries. No permission for: ${entryId}` };
      }
    }
    
    return { success: true };
    
  } catch (error) {
    return { success: false, message: `Validation error: ${error.toString()}` };
  }
}

function getSheetEntriesForLinking(ss, sheetName, userEmail, isAdmin, confirmations, links) {
  const entries = [];
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      return entries;
    }
    
    const data = sheet.getDataRange().getValues();
    const userEmailLower = userEmail.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // More permissive permission check for linking
      // Allow users to link with any entry (not just their own)
      // Comment out the permission check for now
      // if (!isAdmin && rowUserEmail !== userEmailLower) {
      //   continue;
      // }
      
      const entryId = `${sheetName}-${i + 1}`;
      
      // Check if entry has basic data
      const isComplete = !!(row[1] && row[3] && row[4]); // Means, From/To, Subject
      
      // Skip entries with no basic data
      if (!isComplete) {
        continue;
      }
      
      // Check if entry is confirmed
      const confirmKey = `${sheetName}-${i + 1}`;
      const isConfirmed = confirmations.has(confirmKey);
      
      // Get linked entries
      const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
      
      const entry = {
        id: entryId,
        type: sheetName,
        subject: row[4] || '',
        person: row[3] || '', // From/To field
        user: row[5] || '',
        dateTime: row[6] ? formatDateTime(row[6]) : '',
        means: row[1] || '',
        fileReference: row[8] || '',
        postalTariff: row[9] || '',
        complete: isComplete,
        confirmed: isConfirmed,
        linkedEntries: linkedEntries,
        // Additional type-specific fields
        ...(sheetName === 'Inward' ? {
          inwardNo: row[2] || '',
          fromWhom: row[3] || '',
          actionTaken: row[7] || ''
        } : {
          outwardNo: row[2] || '',
          toWhom: row[3] || '',
          caseClosed: row[7] || 'No'
        })
      };
      
      entries.push(entry);
    }
    
  } catch (error) {
    Logger.log(`Error getting entries from ${sheetName} for linking: ${error.toString()}`);
  }
  
  return entries;
}

function getLinkableEntries(currentEntryId) {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`=== GET LINKABLE ENTRIES DEBUG ===`);
    Logger.log(`Current Entry: ${currentEntryId}`);
    Logger.log(`User: ${currentUser}, Admin: ${isAdmin}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const confirmations = getConfirmations(ss);
    const links = getEntryLinks(ss);
    
    // Get existing links for current entry to exclude them
    const existingLinks = getLinkedEntriesForEntry(currentEntryId, links, ss);
    const excludeIds = new Set([currentEntryId, ...existingLinks.map(e => e.id)]);
    
    Logger.log(`Excluding IDs: ${JSON.stringify([...excludeIds])}`);
    
    let allEntries = [];
    
    // Get Inward entries (more permissive - show more entries for linking)
    const inwardEntries = getSheetEntriesForLinking(ss, CONFIG.INWARD_SHEET, currentUser, isAdmin, confirmations, links);
    allEntries = allEntries.concat(inwardEntries);
    Logger.log(`Found ${inwardEntries.length} inward entries`);
    
    // Get Outward entries
    const outwardEntries = getSheetEntriesForLinking(ss, CONFIG.OUTWARD_SHEET, currentUser, isAdmin, confirmations, links);
    allEntries = allEntries.concat(outwardEntries);
    Logger.log(`Found ${outwardEntries.length} outward entries`);
    
    // Filter out current entry and already linked entries
    const linkableEntries = allEntries.filter(entry => !excludeIds.has(entry.id));
    Logger.log(`Linkable entries after filtering: ${linkableEntries.length}`);
    
    // Sort by date (newest first)
    linkableEntries.sort((a, b) => new Date(b.dateTime) - new Date(a.dateTime));
    
    Logger.log(`=== GET LINKABLE ENTRIES COMPLETE ===`);
    
    return {
      success: true,
      entries: linkableEntries,
      message: `Found ${linkableEntries.length} linkable entries`
    };
    
  } catch (error) {
    Logger.log('Error getting linkable entries: ' + error.toString());
    return {
      success: false,
      message: 'Error loading linkable entries: ' + error.toString(),
      entries: []
    };
  }
}

function getEntryLinks(ss) {
  const links = new Map();
  
  try {
    const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
    if (!linksSheet || linksSheet.getLastRow() <= 1) {
      Logger.log('No links sheet or no data in links sheet');
      return links;
    }
    
    const data = linksSheet.getDataRange().getValues();
    Logger.log(`Processing ${data.length - 1} link rows`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] && row[2]) { // Primary Entry ID and Linked Entry ID
        const primaryId = row[1];
        const linkedId = row[2];
        
        if (!links.has(primaryId)) {
          links.set(primaryId, []);
        }
        
        links.get(primaryId).push({
          linkId: row[0],
          linkedEntryId: linkedId,
          linkType: row[3] || 'Manual Link',
          createdDate: row[4],
          createdBy: row[5],
          notes: row[6],
          uuid: row[7] || '' // Link Group UUID
        });
      }
    }
    
    Logger.log(`Loaded links for ${links.size} entries`);
    
  } catch (error) {
    Logger.log('Error loading entry links: ' + error.toString());
  }
  
  return links;
}

function getLinkedEntriesForEntry(entryId, links, ss) {
  const linkedEntries = [];
  
  try {
    const entryLinks = links.get(entryId) || [];
    
    for (const link of entryLinks) {
      const linkedEntryId = link.linkedEntryId;
      const parts = linkedEntryId.split('-');
      
      if (parts.length >= 2) {
        const sheetName = parts[0];
        const rowNumber = parseInt(parts[1]);
        
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && rowNumber <= sheet.getLastRow()) {
          const entryData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
          
          linkedEntries.push({
            id: linkedEntryId,
            type: sheetName,
            subject: entryData[4] || 'No Subject',
            person: entryData[3] || 'Unknown',
            linkType: link.linkType,
            linkedBy: link.createdBy
          });
        }
      }
    }
    
  } catch (error) {
    Logger.log(`Error getting linked entries for ${entryId}: ${error.toString()}`);
  }
  
  return linkedEntries;
}

// =====================================================
// DASHBOARD DATA FUNCTIONS (Updated)
// =====================================================

function getDashboardData() {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`Getting dashboard data for: ${currentUser}, Admin: ${isAdmin}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets();
    
    const stats = calculateUserStats(ss, currentUser, isAdmin);
    
    return {
      success: true,
      userEmail: currentUser,
      isAdmin: isAdmin,
      stats: stats,
      message: 'Dashboard data loaded successfully'
    };
    
  } catch (error) {
    Logger.log('Error in getDashboardData: ' + error.toString());
    return {
      success: false,
      message: 'Error loading dashboard data: ' + error.toString(),
      stats: { ready: 0, confirmed: 0 }
    };
  }
}

function calculateUserStats(ss, userEmail, isAdmin) {
  let pending = 0;
  let confirmed = 0;
  
  try {
    Logger.log(`=== CALCULATING STATS DEBUG ===`);
    Logger.log(`User: "${userEmail}", Admin: ${isAdmin}`);
    
    const confirmations = getConfirmations(ss);
    Logger.log(`Loaded ${confirmations.size} confirmations`);
    
    // Calculate stats for Inward sheet
    const inwardStats = calculateSheetStats(ss, CONFIG.INWARD_SHEET, userEmail, isAdmin, confirmations);
    Logger.log(`Inward stats: ${JSON.stringify(inwardStats)}`);
    
    // Calculate stats for Outward sheet
    const outwardStats = calculateSheetStats(ss, CONFIG.OUTWARD_SHEET, userEmail, isAdmin, confirmations);
    Logger.log(`Outward stats: ${JSON.stringify(outwardStats)}`);
    
    pending = inwardStats.pending + outwardStats.pending;
    confirmed = inwardStats.confirmed + outwardStats.confirmed;
    
    Logger.log(`TOTAL STATS - Pending: ${pending}, Confirmed: ${confirmed}`);
    Logger.log(`=== STATS CALCULATION COMPLETE ===`);
    
  } catch (error) {
    Logger.log('Error calculating stats: ' + error.toString());
  }
  
  return { pending, confirmed };
}

function calculateSheetStats(ss, sheetName, userEmail, isAdmin, confirmations) {
  let pending = 0;
  let confirmed = 0;
  
  try {
    Logger.log(`--- Processing ${sheetName} sheet ---`);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`${sheetName} sheet not found!`);
      return { pending, confirmed };
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`${sheetName} last row: ${lastRow}`);
    
    if (lastRow <= 1) {
      Logger.log(`${sheetName} has no data rows`);
      return { pending, confirmed };
    }
    
    const data = sheet.getDataRange().getValues();
    const userEmailLower = (userEmail || '').toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip entries without subject
      if (!row[4]) continue;
      
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // Permission check
      let shouldProcess = false;
      if (isAdmin) {
        shouldProcess = true;
      } else if (!userEmailLower) {
        shouldProcess = true;
      } else if (rowUserEmail === userEmailLower) {
        shouldProcess = true;
      } else if (!rowUserEmail) {
        shouldProcess = true;
      }
      
      if (!shouldProcess) {
        continue;
      }
      
      // Check if entry is complete (has all required fields)
      const isComplete = !!(row[1] && row[3] && row[4] && row[6]); // Means, Person, Subject, DateTime
      
      // Check if entry is confirmed (physically processed)
      const confirmKey = `${sheetName}-${i + 1}`;
      const isConfirmed = confirmations.has(confirmKey);
      
      // CORRECTED LOGIC:
      if (isConfirmed) {
        // Entry is confirmed = Work Complete
        confirmed++;
      } else if (isComplete) {
        // Entry is complete but not confirmed = Pending Work (ready for physical action)
        pending++;
      }
      // Note: Incomplete entries (missing required fields) are not counted in either category
      // This matches the display logic where incomplete entries show "Incomplete Data" status
    }
    
    Logger.log(`${sheetName} CORRECTED SUMMARY:`);
    Logger.log(`- Pending Work (complete but not confirmed): ${pending}`);
    Logger.log(`- Work Complete (confirmed): ${confirmed}`);
    
  } catch (error) {
    Logger.log(`Error calculating stats for ${sheetName}: ${error.toString()}`);
  }
  
  return { pending, confirmed };
}


// =====================================================
// ENTRY CREATION FUNCTIONS (Same as before)
// =====================================================

function createNewEntry(entryType, entryData) {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    
    Logger.log(`Creating ${entryType} entry for user: ${currentUser}`);
    
    if (!entryType || !entryData) {
      return { success: false, message: 'Invalid parameters provided' };
    }
    
    // Validate entry data
    const validation = validateEntryData(entryData, entryType);
    if (!validation.isValid) {
      return { success: false, message: validation.message };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(entryType === 'Inward' ? CONFIG.INWARD_SHEET : CONFIG.OUTWARD_SHEET);
    
    if (!targetSheet) {
      return { success: false, message: `${entryType} sheet not found` };
    }
    
    const lastRow = targetSheet.getLastRow();
    const newRowNumber = lastRow + 1;
    
    // Calculate proper serial number (always starts from 1)
    const serialNumber = lastRow > 1 ? lastRow : 1; // If lastRow is 1 (header only), start with 1
    
    // Generate automatic inward/outward codes
    const currentYear = new Date().getFullYear();
    const entryNumber = serialNumber; // Use serial number as entry number
    
    const autoInwardNo = entryType === 'Inward' ? `INW/${currentYear}/${entryNumber.toString().padStart(3, '0')}` : '';
    const autoOutwardNo = entryType === 'Outward' ? `OTW/${currentYear}/${entryNumber.toString().padStart(3, '0')}` : '';
    
    Logger.log(`Serial: ${serialNumber}, Auto Code: ${autoInwardNo || autoOutwardNo}`);
    
    let rowData;
    
    if (entryType === 'Inward') {
      rowData = [
        serialNumber, // Sl. No (starts from 1)
        entryData.means || '',
        autoInwardNo, // Auto-generated Inward No
        entryData.fromWhom || '',
        entryData.subject || '',
        entryData.takenBy || '',
        entryData.receiptDateTime ? new Date(entryData.receiptDateTime) : new Date(),
        entryData.actionTaken || '',
        entryData.fileReference || '',
        parseFloat(entryData.postalTariff) || ''
      ];
    } else { // Outward
      rowData = [
        serialNumber, // Sl. No (starts from 1)
        entryData.means || '',
        autoOutwardNo, // Auto-generated Outward No
        entryData.toWhom || '',
        entryData.subject || '',
        entryData.sentBy || '',
        entryData.receiptDateTime ? new Date(entryData.receiptDateTime) : new Date(),
        entryData.caseClosed || 'No',
        entryData.fileReference || '',
        parseFloat(entryData.postalTariff) || '',
        entryData.dueDate ? new Date(entryData.dueDate) : ''  // Added Due Date
      ];
    }
    
    // Write the data
    targetSheet.getRange(newRowNumber, 1, 1, rowData.length).setValues([rowData]);
    
    Logger.log(`Successfully created ${entryType} entry at row ${newRowNumber} with serial ${serialNumber}`);
    
    return {
      success: true,
      message: `${entryType} entry created successfully! (Serial No: ${serialNumber}, Code: ${autoInwardNo || autoOutwardNo})`,
      rowNumber: newRowNumber,
      entryId: `${entryType}-${newRowNumber}`,
      serialNumber: serialNumber,
      generatedCode: autoInwardNo || autoOutwardNo
    };
    
  } catch (error) {
    Logger.log(`Error creating entry: ${error.toString()}`);
    return { success: false, message: `Error creating entry: ${error.toString()}` };
  }
}

function validateEntryData(entryData, entryType) {
  // Check basic required fields
  if (!entryData.receiptDateTime || !entryData.means || !entryData.subject) {
    return { isValid: false, message: 'Date/Time, Means, and Subject are required' };
  }
  
  // Check entry-type specific fields
  if (entryType === 'Inward') {
    if (!entryData.fromWhom) {
      return { isValid: false, message: 'From Whom field is required for inward entries' };
    }
    if (!entryData.takenBy) {
      return { isValid: false, message: 'Taken By field is required for inward entries' };
    }
  }
  
  if (entryType === 'Outward') {
    if (!entryData.toWhom) {
      return { isValid: false, message: 'To Whom field is required for outward entries' };
    }
    if (!entryData.sentBy) {
      return { isValid: false, message: 'Sent By field is required for outward entries' };
    }
  }
  
  // Validate date
  try {
    const testDate = new Date(entryData.receiptDateTime);
    if (isNaN(testDate.getTime())) {
      return { isValid: false, message: 'Invalid date format' };
    }
  } catch (error) {
    return { isValid: false, message: 'Invalid date format' };
  }
  
  return { isValid: true };
}

// =====================================================
// ADMIN FUNCTIONS (Same as before)
// =====================================================

function openAdminPanel() {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    if (!isAdmin) {
      return { success: false, message: 'Admin access required' };
    }
    
    Logger.log(`Opening admin panel for: ${currentUser}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminStats = getAdminStatistics(ss);
    
    return {
      success: true,
      message: 'Admin panel opened',
      stats: adminStats,
      userEmail: currentUser
    };
    
  } catch (error) {
    Logger.log(`Error opening admin panel: ${error.toString()}`);
    return { success: false, message: `Error opening admin panel: ${error.toString()}` };
  }
}

function getAdminStatistics(ss) {
  try {
    const inwardSheet = ss.getSheetByName(CONFIG.INWARD_SHEET);
    const outwardSheet = ss.getSheetByName(CONFIG.OUTWARD_SHEET);
    const confirmSheet = ss.getSheetByName(CONFIG.CONFIRMATION_SHEET);
    const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
    
    const totalInward = inwardSheet ? inwardSheet.getLastRow() - 1 : 0;
    const totalOutward = outwardSheet ? outwardSheet.getLastRow() - 1 : 0;
    const totalConfirmations = confirmSheet ? confirmSheet.getLastRow() - 1 : 0;
    const totalLinks = linksSheet ? (linksSheet.getLastRow() - 1) / 2 : 0; // Divided by 2 because links are bidirectional
    
    const userStats = getUserBreakdown(ss);
    
    return {
      totalInwardEntries: totalInward,
      totalOutwardEntries: totalOutward,
      totalConfirmations: totalConfirmations,
      totalLinks: totalLinks,
      totalEntries: totalInward + totalOutward,
      activeUsers: userStats.length,
      userBreakdown: userStats,
      systemHealth: totalInward + totalOutward > 100 ? 'Busy' : 'Good',
      lastUpdated: new Date().toLocaleString()
    };
    
  } catch (error) {
    Logger.log(`Error getting admin statistics: ${error.toString()}`);
    return {
      totalInwardEntries: 0,
      totalOutwardEntries: 0,
      totalConfirmations: 0,
      totalLinks: 0,
      totalEntries: 0,
      activeUsers: 0,
      userBreakdown: [],
      systemHealth: 'Error',
      lastUpdated: new Date().toLocaleString()
    };
  }
}

function getUserBreakdown(ss) {
  const userMap = new Map();
  
  try {
    // Process Inward sheet
    const inwardSheet = ss.getSheetByName(CONFIG.INWARD_SHEET);
    if (inwardSheet && inwardSheet.getLastRow() > 1) {
      const data = inwardSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const userEmail = data[i][5]; // Taken by column
        if (userEmail) {
          const email = userEmail.toString().toLowerCase();
          if (!userMap.has(email)) {
            userMap.set(email, { inward: 0, outward: 0, confirmed: 0 });
          }
          userMap.get(email).inward++;
        }
      }
    }
    
    // Process Outward sheet
    const outwardSheet = ss.getSheetByName(CONFIG.OUTWARD_SHEET);
    if (outwardSheet && outwardSheet.getLastRow() > 1) {
      const data = outwardSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const userEmail = data[i][5]; // Sent by column
        if (userEmail) {
          const email = userEmail.toString().toLowerCase();
          if (!userMap.has(email)) {
            userMap.set(email, { inward: 0, outward: 0, confirmed: 0 });
          }
          userMap.get(email).outward++;
        }
      }
    }
    
    // Process Confirmations
    const confirmSheet = ss.getSheetByName(CONFIG.CONFIRMATION_SHEET);
    if (confirmSheet && confirmSheet.getLastRow() > 1) {
      const data = confirmSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const userEmail = data[i][1]; // User Email column
        if (userEmail) {
          const email = userEmail.toString().toLowerCase();
          if (userMap.has(email)) {
            userMap.get(email).confirmed++;
          }
        }
      }
    }
    
    // Convert to array
    const userArray = [];
    for (const [email, stats] of userMap.entries()) {
      userArray.push({
        email: email,
        inwardEntries: stats.inward,
        outwardEntries: stats.outward,
        confirmedEntries: stats.confirmed,
        totalEntries: stats.inward + stats.outward,
        isAdmin: CONFIG.ADMIN_USERS.includes(email)
      });
    }
    
    return userArray.sort((a, b) => b.totalEntries - a.totalEntries);
    
  } catch (error) {
    Logger.log(`Error getting user breakdown: ${error.toString()}`);
    return [];
  }
}

// =====================================================
// REPORT FUNCTIONS (Same as before)
// =====================================================

function generateSystemReport() {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`Generating system report for: ${currentUser}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stats = getDashboardData();
    const adminStats = isAdmin ? getAdminStatistics(ss) : null;
    
    const report = {
      generatedBy: currentUser,
      generatedAt: new Date().toLocaleString(),
      userStats: stats.stats,
      adminStats: adminStats,
      summary: {
        message: 'Report generated successfully',
        userEmail: currentUser,
        isAdmin: isAdmin
      }
    };
    
    Logger.log('System report generated successfully');
    return {
      success: true,
      message: 'Report generated successfully',
      report: report
    };
    
  } catch (error) {
    Logger.log(`Error generating report: ${error.toString()}`);
    return { success: false, message: `Error generating report: ${error.toString()}` };
  }
}

function fixSerialNumbers(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, message: 'No data to fix' };
    }
    
    Logger.log(`Fixing serial numbers and codes in ${sheetName} sheet`);
    
    const currentYear = new Date().getFullYear();
    const isInward = sheetName === CONFIG.INWARD_SHEET;
    
    // Update serial numbers and codes starting from row 2 (first data row)
    for (let i = 2; i <= lastRow; i++) {
      const serialNumber = i - 1; // Row 2 gets serial 1, row 3 gets serial 2, etc.
      const entryNumber = serialNumber;
      
      // Generate the appropriate code
      const autoCode = isInward 
        ? `INW/${currentYear}/${entryNumber.toString().padStart(3, '0')}`
        : `OTW/${currentYear}/${entryNumber.toString().padStart(3, '0')}`;
      
      // Update both serial number (column 1) and code (column 2)
      sheet.getRange(i, 1).setValue(serialNumber);
      sheet.getRange(i, 3).setValue(autoCode); // Inward/Outward No is column 3
      
      Logger.log(`Row ${i}: Serial ${serialNumber}, Code ${autoCode}`);
    }
    
    Logger.log(`Fixed serial numbers and codes for ${lastRow - 1} entries in ${sheetName}`);
    
    return {
      success: true,
      message: `Fixed serial numbers and codes for ${lastRow - 1} entries in ${sheetName}`,
      entriesFixed: lastRow - 1
    };
    
  } catch (error) {
    Logger.log(`Error fixing serial numbers: ${error.toString()}`);
    return { success: false, message: `Error fixing serial numbers: ${error.toString()}` };
  }
}


// Function to fix both sheets
function fixAllSerialNumbersAndCodes() {
  try {
    const inwardResult = fixSerialNumbers(CONFIG.INWARD_SHEET);
    const outwardResult = fixSerialNumbers(CONFIG.OUTWARD_SHEET);
    
    return {
      success: true,
      message: 'Serial numbers and codes fixed for both sheets',
      inward: inwardResult,
      outward: outwardResult
    };
  } catch (error) {
    return { success: false, message: `Error fixing serial numbers and codes: ${error.toString()}` };
  }
}


function searchInSheetSimplified(ss, sheetName, query, userEmail, isAdmin) {
  const results = [];
  
  try {
    Logger.log(`--- Searching in ${sheetName} ---`);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`Sheet ${sheetName} not found`);
      return results;
    }
    
    const lastRow = sheet.getLastRow();
    Logger.log(`${sheetName} last row: ${lastRow}`);
    
    if (lastRow <= 1) {
      Logger.log(`${sheetName} has no data rows`);
      return results;
    }
    
    // Get all data at once for efficiency
    const data = sheet.getDataRange().getValues();
    Logger.log(`Retrieved data array with ${data.length} rows`);
    
    // Log first few rows for debugging
    if (data.length > 1) {
      Logger.log(`Sample data row 1 (header): ${JSON.stringify(data[0])}`);
      if (data.length > 2) {
        Logger.log(`Sample data row 2: ${JSON.stringify(data[1])}`);
      }
    }
    
    const userEmailLower = userEmail.toLowerCase();
    let processedRows = 0;
    let matchedRows = 0;
    
    for (let i = 1; i < data.length; i++) { // Skip header row
      processedRows++;
      const row = data[i];
      
      // Check if row has basic required data
      if (!row || row.length < 6) {
        Logger.log(`Row ${i + 1} has insufficient data: ${row.length} columns`);
        continue;
      }
      
      // Get the user who created this entry (column index 5 = "Taken By" or "Sent By")
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // For debugging - log a few entries
      if (i <= 3) {
        Logger.log(`Row ${i + 1} user: "${rowUserEmail}", current user: "${userEmailLower}"`);
      }
      
      // Permission check - skip if not admin and entry doesn't belong to current user
      // BUT: Let's temporarily allow all entries for debugging
      // if (!isAdmin && rowUserEmail !== userEmailLower) {
      //   continue;
      // }
      
      // Create searchable text from all relevant fields
      const searchableText = [
        row[1] || '', // Means
        row[2] || '', // Inward/Outward No
        row[3] || '', // From/To
        row[4] || '', // Subject
        row[5] || '', // Taken/Sent By
        row[7] || '', // Action Taken/Case Closed
        row[8] || '', // File Reference
      ].join(' ').toLowerCase();
      
      // Check if query matches
      if (searchableText.includes(query)) {
        matchedRows++;
        
        // Log the match for debugging
        if (matchedRows <= 3) {
          Logger.log(`Match found in row ${i + 1}: "${searchableText.substring(0, 100)}..."`);
        }
        
        const entryId = `${sheetName}-${i + 1}`;
        
        const entry = {
          id: entryId,
          type: sheetName,
          subject: row[4] || '',
          person: row[3] || '',
          user: row[5] || '',
          dateTime: row[6] ? formatDateTime(row[6]) : '',
          means: row[1] || '',
          fileReference: row[8] || '',
          postalTariff: row[9] || '',
          complete: !!(row[1] && row[3] && row[4] && row[6]), // Basic completeness check
          confirmed: false, // We'll determine this later if needed
          linkedEntries: [], // We'll populate this later if needed
          relevanceScore: 100, // Simple scoring for now
          // Additional fields for display
          ...(sheetName === 'Inward' ? {
            inwardNo: row[2] || '',
            fromWhom: row[3] || '',
            actionTaken: row[7] || ''
          } : {
            outwardNo: row[2] || '',
            toWhom: row[3] || '',
            caseClosed: row[7] || 'No'
          })
        };
        
        results.push(entry);
      }
    }
    
    Logger.log(`${sheetName} - Processed: ${processedRows} rows, Matched: ${matchedRows} rows`);
    
  } catch (error) {
    Logger.log(`Error searching in ${sheetName}: ${error.toString()}`);
    Logger.log(`Error stack: ${error.stack}`);
  }
  
  return results;
}

function searchInSheetEnhanced(ss, sheetName, query, userEmail, isAdmin, confirmations, links) {
  const results = [];
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      return results;
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Must have subject to be searchable
      if (!row[4]) continue;
      
      // Create searchable text from all relevant fields
      const searchableText = [
        row[1] || '', // Means
        row[2] || '', // Inward/Outward No
        row[3] || '', // From/To
        row[4] || '', // Subject
        row[5] || '', // Taken/Sent By
        row[7] || '', // Action Taken/Case Closed
        row[8] || '', // File Reference
      ].join(' ').toLowerCase();
      
      if (searchableText.includes(query)) {
        const entryId = `${sheetName}-${i + 1}`;
        const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
        const confirmKey = `${sheetName}-${i + 1}`;
        const isConfirmed = confirmations.has(confirmKey);
        const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
        
        let formattedDateTime = '';
        try {
          if (row[6]) {
            formattedDateTime = row[6] instanceof Date ? 
              row[6].toLocaleString() : new Date(row[6]).toLocaleString();
          }
        } catch (dateError) {
          formattedDateTime = row[6] ? row[6].toString() : '';
        }
        
        const entry = {
          id: entryId,
          type: sheetName,
          subject: (row[4] || '').toString(),
          person: (row[3] || '').toString(),
          user: (row[5] || '').toString(),
          dateTime: formattedDateTime,
          means: (row[1] || '').toString(),
          fileReference: (row[8] || '').toString(),
          complete: isComplete,
          confirmed: isConfirmed,
          linkedEntries: linkedEntries,
          hasLinks: linkedEntries.length > 0,
          relevanceScore: calculateRelevanceScore(searchableText, query),
          // Type-specific fields
          ...(sheetName === 'Inward' ? {
            inwardNo: (row[2] || '').toString(),
            fromWhom: (row[3] || '').toString(),
            actionTaken: (row[7] || '').toString()
          } : {
            outwardNo: (row[2] || '').toString(),
            toWhom: (row[3] || '').toString(),
            caseClosed: (row[7] || 'No').toString()
          })
        };
        
        results.push(entry);
      }
    }
    
  } catch (error) {
    Logger.log(`Error searching in ${sheetName}: ${error.toString()}`);
  }
  
  return results;
}

function getEntryFullDetails(ss, entryId, confirmations, links) {
  try {
    const parts = entryId.split('-');
    if (parts.length < 2) return null;
    
    const sheetName = parts[0];
    const rowNumber = parseInt(parts[1]);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || rowNumber > sheet.getLastRow()) return null;
    
    const row = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    if (!row[4]) return null; // No subject
    
    const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
    const confirmKey = `${sheetName}-${rowNumber}`;
    const isConfirmed = confirmations.has(confirmKey);
    const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
    
    let formattedDateTime = '';
    try {
      if (row[6]) {
        formattedDateTime = row[6] instanceof Date ? 
          row[6].toLocaleString() : new Date(row[6]).toLocaleString();
      }
    } catch (dateError) {
      formattedDateTime = row[6] ? row[6].toString() : '';
    }
    
    return {
      id: entryId,
      type: sheetName,
      subject: (row[4] || '').toString(),
      person: (row[3] || '').toString(),
      user: (row[5] || '').toString(),
      dateTime: formattedDateTime,
      means: (row[1] || '').toString(),
      fileReference: (row[8] || '').toString(),
      complete: isComplete,
      confirmed: isConfirmed,
      linkedEntries: linkedEntries,
      hasLinks: linkedEntries.length > 0,
      // Type-specific fields
      ...(sheetName === 'Inward' ? {
        inwardNo: (row[2] || '').toString(),
        fromWhom: (row[3] || '').toString(),
        actionTaken: (row[7] || '').toString()
      } : {
        outwardNo: (row[2] || '').toString(),
        toWhom: (row[3] || '').toString(),
        caseClosed: (row[7] || 'No').toString()
      })
    };
    
  } catch (error) {
    Logger.log(`Error getting entry details for ${entryId}: ${error.toString()}`);
    return null;
  }
}

function sortSearchResultsEnhanced(results, query) {
  return results.sort((a, b) => {
    // Direct search results first
    if (a.isDirectResult && !b.isDirectResult) return -1;
    if (!a.isDirectResult && b.isDirectResult) return 1;
    
    // Then by relevance score (if both are direct results)
    if (a.isDirectResult && b.isDirectResult) {
      if (b.relevanceScore !== a.relevanceScore) {
        return b.relevanceScore - a.relevanceScore;
      }
    }
    
    // Finally by date (newest first)
    const dateA = new Date(a.dateTime || 0);
    const dateB = new Date(b.dateTime || 0);
    return dateB - dateA;
  });
}

function searchEntriesWithLinkFilter(searchQuery, searchType = 'all', linkFilter = 'all') {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`=== ENHANCED SEARCH WITH LINK FILTER ===`);
    Logger.log(`Query: "${searchQuery}", Type: ${searchType}, Link Filter: ${linkFilter}, User: ${currentUser}`);
    
    if (!searchQuery || searchQuery.trim().length < 1) {
      return { success: false, message: 'Search query must be at least 1 character long' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const confirmations = getConfirmationsSimplified(ss);
    const links = getEntryLinks(ss);
    
    let directResults = [];
    const query = searchQuery.toLowerCase().trim();
    
    // Search in sheets directly
    if (searchType === 'all' || searchType === 'inward') {
      const inwardResults = searchInSheetEnhanced(ss, CONFIG.INWARD_SHEET, query, currentUser, isAdmin, confirmations, links);
      directResults = directResults.concat(inwardResults);
    }
    
    if (searchType === 'all' || searchType === 'outward') {
      const outwardResults = searchInSheetEnhanced(ss, CONFIG.OUTWARD_SHEET, query, currentUser, isAdmin, confirmations, links);
      directResults = directResults.concat(outwardResults);
    }
    
    Logger.log(`Direct search results: ${directResults.length}`);
    
    // Apply link filtering
    let filteredResults = [];
    
    switch(linkFilter) {
      case 'linked-only':
        filteredResults = directResults.filter(entry => 
          entry.linkedEntries && entry.linkedEntries.length > 0
        );
        break;
      case 'no-links':
        filteredResults = directResults.filter(entry => 
          !entry.linkedEntries || entry.linkedEntries.length === 0
        );
        break;
      case 'by-uuid':
        // For UUID search, use the searchByUUID function
        return searchByUUID(searchQuery);
      default:
        // 'all' - include linked entries as before
        let linkedResults = [];
        const processedIds = new Set();
        
        for (const result of directResults) {
          processedIds.add(result.id);
          
          const linkedEntries = getLinkedEntriesForEntry(result.id, links, ss);
          
          for (const linkedEntry of linkedEntries) {
            if (!processedIds.has(linkedEntry.id)) {
              const linkedEntryFull = getEntryFullDetails(ss, linkedEntry.id, confirmations, links);
              if (linkedEntryFull) {
                linkedEntryFull.linkedToSearchResult = true;
                linkedEntryFull.linkedToEntry = {
                  id: result.id,
                  type: result.type,
                  subject: result.subject
                };
                linkedResults.push(linkedEntryFull);
                processedIds.add(linkedEntry.id);
              }
            }
          }
        }
        
        filteredResults = [
          ...directResults.map(r => ({ ...r, isDirectResult: true })),
          ...linkedResults
        ];
    }
    
    // Remove duplicates and sort
    const uniqueResults = removeDuplicateEntries(filteredResults);
    const sortedResults = sortSearchResultsEnhanced(uniqueResults, query);
    
    Logger.log(`Final filtered results: ${uniqueResults.length}`);
    
    let filterDescription = '';
    switch(linkFilter) {
      case 'linked-only':
        filterDescription = ' (linked entries only)';
        break;
      case 'no-links':
        filterDescription = ' (entries without links)';
        break;
      case 'by-uuid':
        filterDescription = ' (UUID search)';
        break;
    }
    
    return {
      success: true,
      results: sortedResults,
      searchQuery: searchQuery,
      searchType: searchType,
      linkFilter: linkFilter,
      totalResults: sortedResults.length,
      directResults: directResults.length,
      linkedResults: filteredResults.length - directResults.length,
      message: `Found ${sortedResults.length} results for "${searchQuery}"${filterDescription}`
    };
    
  } catch (error) {
    Logger.log(`Error in enhanced search with link filter: ${error.toString()}`);
    return { success: false, message: `Search error: ${error.toString()}` };
  }
}

function searchInSheet(ss, sheetName, query, userEmail, isAdmin, confirmations, links) {
  const results = [];
  
  try {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      return results;
    }
    
    const data = sheet.getDataRange().getValues();
    const userEmailLower = userEmail.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // Skip if not admin and entry doesn't belong to current user
      if (!isAdmin && rowUserEmail !== userEmailLower) {
        continue;
      }
      
      // Search in relevant fields
      const searchableText = [
        row[1], // Means
        row[2], // Inward/Outward No
        row[3], // From/To
        row[4], // Subject
        row[5], // Taken/Sent By
        row[7], // Action Taken/Case Closed
        row[8], // File Reference
      ].join(' ').toLowerCase();
      
      if (searchableText.includes(query)) {
        const entryId = `${sheetName}-${i + 1}`;
        
        // Check if entry is complete and confirmed
        const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
        const confirmKey = `${sheetName}-${i + 1}`;
        const isConfirmed = confirmations.has(confirmKey);
        
        // Get linked entries
        const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
        
        const entry = {
          id: entryId,
          type: sheetName,
          subject: row[4] || '',
          person: row[3] || '',
          user: row[5] || '',
          dateTime: row[6] ? formatDateTime(row[6]) : '',
          means: row[1] || '',
          fileReference: row[8] || '',
          complete: isComplete,
          confirmed: isConfirmed,
          linkedEntries: linkedEntries,
          relevanceScore: calculateRelevanceScore(searchableText, query),
          // Additional fields for display
          ...(sheetName === 'Inward' ? {
            inwardNo: row[2] || '',
            fromWhom: row[3] || '',
            actionTaken: row[7] || ''
          } : {
            outwardNo: row[2] || '',
            toWhom: row[3] || '',
            caseClosed: row[7] || 'No'
          })
        };
        
        results.push(entry);
      }
    }
    
  } catch (error) {
    Logger.log(`Error searching in ${sheetName}: ${error.toString()}`);
  }
  
  return results;
}

function searchByUUID(uuidQuery) {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`=== UUID SEARCH ===`);
    Logger.log(`UUID Query: "${uuidQuery}", User: ${currentUser}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const confirmations = getConfirmationsSimplified(ss);
    const links = getEntryLinks(ss);
    const results = [];
    
    const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
    if (!linksSheet || linksSheet.getLastRow() <= 1) {
      return {
        success: true,
        results: [],
        message: 'No UUID links found'
      };
    }
    
    const data = linksSheet.getDataRange().getValues();
    const matchedEntryIds = new Set();
    
    // Search for UUID in links
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const linkUUID = row[7] ? row[7].toString() : ''; // UUID column
      const notes = row[6] ? row[6].toString().toLowerCase() : ''; // Notes column
      
      if (linkUUID.includes(uuidQuery) || notes.includes(uuidQuery.toLowerCase())) {
        matchedEntryIds.add(row[1]); // Primary Entry ID
        matchedEntryIds.add(row[2]); // Linked Entry ID
      }
    }
    
    Logger.log(`Found ${matchedEntryIds.size} entries with UUID: ${uuidQuery}`);
    
    // Get details for matched entries
    for (const entryId of matchedEntryIds) {
      const parts = entryId.split('-');
      if (parts.length >= 2) {
        const sheetName = parts[0];
        const rowNumber = parseInt(parts[1]);
        
        const sheet = ss.getSheetByName(sheetName);
        if (sheet && rowNumber <= sheet.getLastRow()) {
          const entryData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
          const entryUserEmail = entryData[5] ? entryData[5].toString().toLowerCase() : '';
          
          // Check permissions (simplified for UUID search)
          if (!isAdmin && currentUser && entryUserEmail !== currentUser.toLowerCase()) {
            continue;
          }
          
          const isComplete = !!(entryData[1] && entryData[3] && entryData[4] && entryData[6]);
          const confirmKey = `${sheetName}-${rowNumber}`;
          const isConfirmed = confirmations.has(confirmKey);
          const linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
          
          const entry = {
            id: entryId,
            type: sheetName,
            subject: entryData[4] || '',
            person: entryData[3] || '',
            user: entryData[5] || '',
            dateTime: entryData[6] ? formatDateTime(entryData[6]) : '',
            means: entryData[1] || '',
            fileReference: entryData[8] || '',
            complete: isComplete,
            confirmed: isConfirmed,
            linkedEntries: linkedEntries,
            relevanceScore: 100, // High relevance for UUID matches
            matchedByUUID: true,
            uuidQuery: uuidQuery,
            // Type-specific fields
            ...(sheetName === 'Inward' ? {
              inwardNo: entryData[2] || '',
              fromWhom: entryData[3] || '',
              actionTaken: entryData[7] || ''
            } : {
              outwardNo: entryData[2] || '',
              toWhom: entryData[3] || '',
              caseClosed: entryData[7] || 'No'
            })
          };
          
          results.push(entry);
        }
      }
    }
    
    // Sort results by date (newest first)
    results.sort((a, b) => new Date(b.dateTime) - new Date(a.dateTime));
    
    Logger.log(`UUID search complete: ${results.length} results`);
    
    return {
      success: true,
      results: results,
      searchQuery: uuidQuery,
      searchType: 'uuid',
      totalResults: results.length,
      message: results.length > 0 
        ? `Found ${results.length} entries with UUID "${uuidQuery}"`
        : `No entries found with UUID "${uuidQuery}"`
    };
    
  } catch (error) {
    Logger.log(`Error in UUID search: ${error.toString()}`);
    return { success: false, message: `UUID search error: ${error.toString()}` };
  }
}

// =====================================================
// UTILITY FUNCTIONS FOR SEARCH
// =====================================================

function calculateRelevanceScore(text, query) {
  let score = 0;
  const queryWords = query.split(' ').filter(word => word.length > 0);
  
  for (const word of queryWords) {
    const wordCount = (text.match(new RegExp(word, 'g')) || []).length;
    score += wordCount * 10;
    
    // Bonus for exact matches
    if (text.includes(query)) {
      score += 50;
    }
  }
  
  return score;
}

function sortSearchResults(results, query) {
  return results.sort((a, b) => {
    // UUID matches first
    if (a.matchedByUUID && !b.matchedByUUID) return -1;
    if (!a.matchedByUUID && b.matchedByUUID) return 1;
    
    // Then by relevance score
    if (b.relevanceScore !== a.relevanceScore) {
      return b.relevanceScore - a.relevanceScore;
    }
    
    // Then by date (newest first)
    return new Date(b.dateTime) - new Date(a.dateTime);
  });
}

function removeDuplicateEntries(results) {
  const seen = new Set();
  return results.filter(entry => {
    if (seen.has(entry.id)) {
      return false;
    }
    seen.add(entry.id);
    return true;
  });
}

function generateUUID() {
  // Generate a simple UUID-like string
  return 'uuid-' + Date.now().toString(36) + '-' + Math.random().toString(36).substr(2, 9);
}



// =====================================================
// UTILITY FUNCTIONS
// =====================================================

function testBackendFunctions() {
  Logger.log('=== TESTING ENHANCED BACKEND FUNCTIONS ===');
  
  const userInfo = getCurrentUser();
  Logger.log(`Current user: ${JSON.stringify(userInfo)}`);
  
  // Test dashboard data
  const dashboardResult = getDashboardData();
  Logger.log(`Dashboard result: ${JSON.stringify(dashboardResult)}`);
  
  // Test entries loading
  const entriesResult = getEntriesWithDetails();
  Logger.log(`Entries result: ${JSON.stringify(entriesResult)}`);
  
  // Test admin functions if user is admin
  if (userInfo.isAdmin) {
    const adminResult = openAdminPanel();
    Logger.log(`Admin result: ${JSON.stringify(adminResult)}`);
  }
  
  Logger.log('=== TESTING COMPLETE ===');
  return { userInfo, dashboardResult, entriesResult };
}

function getAllEntriesForSearch() {
  try {
    Logger.log('=== LOADING ENTRIES FOR SEARCH DROPDOWN ===');
    
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`User: ${currentUser}, Admin: ${isAdmin}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets(); // Ensure sheets exist
    
    let allEntries = [];
    
    // Load from both sheets with simplified logic
    const sheets = [CONFIG.INWARD_SHEET, CONFIG.OUTWARD_SHEET];
    
    for (const sheetName of sheets) {
      try {
        Logger.log(`Processing ${sheetName} sheet...`);
        
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() <= 1) {
          Logger.log(`${sheetName}: No data found`);
          continue;
        }
        
        const data = sheet.getDataRange().getValues();
        Logger.log(`${sheetName}: Found ${data.length - 1} rows`);
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          
          // Must have subject to be included in search
          if (!row[4] || row[4].toString().trim() === '') {
            continue;
          }
          
          // More permissive user filtering for search dropdown
          const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
          const currentUserLower = (currentUser || '').toLowerCase();
          
          // Include entry if: admin, no user restriction, user matches, or no user in entry
          const shouldInclude = isAdmin || 
                               !currentUserLower || 
                               rowUserEmail === currentUserLower || 
                               !rowUserEmail;
          
          if (!shouldInclude) {
            continue;
          }
          
          const entryId = `${sheetName}-${i + 1}`;
          
          // Format date safely
          let formattedDateTime = '';
          try {
            if (row[6]) {
              if (row[6] instanceof Date) {
                formattedDateTime = row[6].toLocaleString('en-US', {
                  year: 'numeric',
                  month: 'short',
                  day: '2-digit',
                  hour: '2-digit',
                  minute: '2-digit'
                });
              } else {
                formattedDateTime = new Date(row[6]).toLocaleString('en-US', {
                  year: 'numeric',
                  month: 'short',
                  day: '2-digit',
                  hour: '2-digit',
                  minute: '2-digit'
                });
              }
            }
          } catch (dateError) {
            formattedDateTime = row[6] ? row[6].toString() : 'Unknown Date';
          }
          
          // Check basic completeness
          const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
          
          const entry = {
            id: entryId,
            type: sheetName,
            subject: (row[4] || '').toString(),
            person: (row[3] || '').toString(),
            user: (row[5] || '').toString(),
            dateTime: formattedDateTime,
            means: (row[1] || '').toString(),
            complete: isComplete,
            confirmed: false, // Simplified for dropdown
            // Type-specific fields
            inwardNo: sheetName === 'Inward' ? (row[2] || '').toString() : '',
            outwardNo: sheetName === 'Outward' ? (row[2] || '').toString() : '',
            actionTaken: sheetName === 'Inward' ? (row[7] || '').toString() : ''
          };
          
          allEntries.push(entry);
        }
        
        Logger.log(`${sheetName}: Added ${allEntries.filter(e => e.type === sheetName).length} entries`);
        
      } catch (sheetError) {
        Logger.log(`Error processing ${sheetName}: ${sheetError.toString()}`);
        // Continue with other sheets
      }
    }
    
    // Sort by date (newest first)
    allEntries.sort((a, b) => {
      const dateA = new Date(a.dateTime || 0);
      const dateB = new Date(b.dateTime || 0);
      return dateB - dateA;
    });
    
    Logger.log(`TOTAL ENTRIES FOR SEARCH: ${allEntries.length}`);
    
    return {
      success: true,
      entries: allEntries,
      count: allEntries.length,
      message: `Found ${allEntries.length} entries for search`
    };
    
  } catch (error) {
    Logger.log(`Error in getAllEntriesForSearch: ${error.toString()}`);
    Logger.log(`Error stack: ${error.stack}`);
    
    return {
      success: false,
      message: 'Error loading search entries: ' + error.toString(),
      entries: [],
      count: 0
    };
  }
}

function getSheetEntriesForSearchDropdownDebug(ss, sheetName, userEmail, isAdmin, confirmations, links) {
  const entries = [];
  
  try {
    Logger.log(`--- Processing ${sheetName} for search dropdown ---`);
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      Logger.log(`${sheetName}: No data for search dropdown`);
      return entries;
    }
    
    const data = sheet.getDataRange().getValues();
    const userEmailLower = (userEmail || '').toLowerCase();
    
    Logger.log(`${sheetName} search: Processing ${data.length - 1} rows`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Must have at least subject to be searchable
      if (!row[4]) { // No subject
        continue;
      }
      
      const rowUserEmail = row[5] ? row[5].toString().toLowerCase() : '';
      
      // More permissive filtering for search - show more entries
      let shouldInclude = isAdmin || !userEmailLower || rowUserEmail === userEmailLower || !rowUserEmail;
      
      if (!shouldInclude) {
        continue;
      }
      
      const entryId = `${sheetName}-${i + 1}`;
      const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
      const confirmKey = `${sheetName}-${i + 1}`;
      const isConfirmed = confirmations.has(confirmKey);
      
      let linkedEntries = [];
      try {
        linkedEntries = getLinkedEntriesForEntry(entryId, links, ss);
      } catch (linkError) {
        // Ignore link errors for search dropdown
      }
      
      let formattedDateTime = '';
      try {
        if (row[6]) {
          formattedDateTime = formatDateTime(row[6]);
        }
      } catch (dateError) {
        formattedDateTime = row[6] ? row[6].toString() : '';
      }
      
      const entry = {
        id: entryId,
        type: sheetName,
        subject: (row[4] || '').toString(),
        person: (row[3] || '').toString(),
        user: (row[5] || '').toString(),
        dateTime: formattedDateTime,
        means: (row[1] || '').toString(),
        fileReference: (row[8] || '').toString(),
        complete: isComplete,
        confirmed: isConfirmed,
        linkedEntries: linkedEntries
      };
      
      entries.push(entry);
    }
    
    Logger.log(`${sheetName} search entries: ${entries.length}`);
    
  } catch (error) {
    Logger.log(`Error getting ${sheetName} entries for search: ${error.toString()}`);
  }
  
  return entries;
}

function testStatsCalculation() {
  Logger.log('=== MANUAL STATS TEST ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setupSheets();
    
    // Test with empty user (should show all entries)
    const userInfo = getCurrentUser();
    Logger.log(`User info: ${JSON.stringify(userInfo)}`);
    
    const stats1 = calculateUserStats(ss, userInfo.userEmail, userInfo.isAdmin);
    Logger.log(`Stats with current user: ${JSON.stringify(stats1)}`);
    
    // Test with admin privileges (should show all entries)
    const stats2 = calculateUserStats(ss, userInfo.userEmail, true);
    Logger.log(`Stats with admin override: ${JSON.stringify(stats2)}`);
    
    // Test with no user filter (should show all entries)
    const stats3 = calculateUserStats(ss, '', true);
    Logger.log(`Stats with no user filter: ${JSON.stringify(stats3)}`);
    
    return { userInfo, stats1, stats2, stats3 };
    
  } catch (error) {
    Logger.log(`Test error: ${error.toString()}`);
    return { error: error.toString() };
  }
}

function getAllLinkedEntries() {
  try {
    const userInfo = getCurrentUser();
    const currentUser = userInfo.userEmail;
    const isAdmin = userInfo.isAdmin;
    
    Logger.log(`=== GET ALL LINKED ENTRIES ===`);
    Logger.log(`User: ${currentUser}, Admin: ${isAdmin}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const confirmations = getConfirmationsSimplified(ss);
    const links = getEntryLinks(ss);
    
    let allEntries = [];
    
    // Load from both sheets
    const sheets = [CONFIG.INWARD_SHEET, CONFIG.OUTWARD_SHEET];
    
    for (const sheetName of sheets) {
      try {
        const entries = loadEntriesWithStatus(ss, sheetName, sheetName, currentUser, isAdmin, confirmations, links);
        allEntries = allEntries.concat(entries);
        Logger.log(`Loaded ${entries.length} entries from ${sheetName}`);
      } catch (error) {
        Logger.log(`Error loading ${sheetName}: ${error.toString()}`);
      }
    }
    
    // Filter to only entries with links
    const linkedEntries = allEntries.filter(entry => 
      entry.linkedEntries && entry.linkedEntries.length > 0
    );
    
    // Sort by date (newest first)
    linkedEntries.sort((a, b) => {
      const dateA = new Date(a.dateTime || 0);
      const dateB = new Date(b.dateTime || 0);
      return dateB - dateA;
    });
    
    Logger.log(`Found ${linkedEntries.length} entries with links`);
    
    return {
      success: true,
      entries: linkedEntries,
      totalEntries: allEntries.length,
      linkedEntries: linkedEntries.length,
      message: `Found ${linkedEntries.length} entries with links out of ${allEntries.length} total entries`
    };
    
  } catch (error) {
    Logger.log(`Error getting all linked entries: ${error.toString()}`);
    return {
      success: false,
      message: 'Error loading linked entries: ' + error.toString(),
      entries: []
    };
  }
}
function getLinkStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
    
    if (!linksSheet || linksSheet.getLastRow() <= 1) {
      return {
        totalLinks: 0,
        uniqueUUIDs: 0,
        entriesWithLinks: 0
      };
    }
    
    const data = linksSheet.getDataRange().getValues();
    const uuids = new Set();
    const linkedEntryIds = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[7]) { // UUID column
        uuids.add(row[7]);
      }
      if (row[1]) { // Primary Entry ID
        linkedEntryIds.add(row[1]);
      }
      if (row[2]) { // Linked Entry ID
        linkedEntryIds.add(row[2]);
      }
    }
    
    return {
      totalLinks: data.length - 1, // Exclude header
      uniqueUUIDs: uuids.size,
      entriesWithLinks: linkedEntryIds.size
    };
    
  } catch (error) {
    Logger.log(`Error getting link statistics: ${error.toString()}`);
    return {
      totalLinks: 0,
      uniqueUUIDs: 0,
      entriesWithLinks: 0
    };
  }
}

function updateEntryActionTaken(entryId, actionTaken) {
  try {
    const parts = entryId.split('-');
    const sheetName = parts[0];
    const rowNumber = parseInt(parts[1]);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || rowNumber > sheet.getLastRow()) {
      return { success: false, message: 'Entry not found' };
    }
    
    // Update Action Taken field (column 8 for Inward entries)
    sheet.getRange(rowNumber, 8).setValue(actionTaken);
    
    return { 
      success: true, 
      message: 'Action Taken updated successfully',
      entryId: entryId
    };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
function updateEntry(updatedData) {
  try {
    Logger.log(`Updating entry: ${updatedData.id}`);
    
    if (!updatedData || !updatedData.id) {
      return { success: false, message: 'Entry ID is required' };
    }
    
    // Parse entry ID
    const parts = updatedData.id.split('-');
    if (parts.length < 2) {
      return { success: false, message: 'Invalid entry ID format' };
    }
    
    const sheetName = parts[0];
    const rowNumber = parseInt(parts[1]);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || rowNumber > sheet.getLastRow() || rowNumber < 2) {
      return { success: false, message: 'Entry not found' };
    }
    
    // Validate required fields based on entry type
    const validation = validateUpdateData(updatedData);
    if (!validation.isValid) {
      return { success: false, message: validation.message };
    }
    
    // Get current data to preserve serial number and auto-generated codes
    const currentData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    const serialNumber = currentData[0]; // Preserve existing serial number
    
    // Prepare updated row data
    let rowData;
    
    if (updatedData.type === 'Inward') {
      rowData = [
        serialNumber, // Keep existing serial number
        updatedData.means || '',
        updatedData.inwardNo || currentData[2], // Keep existing inward number
        updatedData.fromWhom || '',
        updatedData.subject || '',
        updatedData.takenBy || '',
        updatedData.receiptDateTime ? new Date(updatedData.receiptDateTime) : currentData[6],
        updatedData.actionTaken || '',
        updatedData.fileReference || '',
        parseFloat(updatedData.postalTariff) || ''
      ];
    } else { // Outward
      rowData = [
        serialNumber, // Keep existing serial number
        updatedData.means || '',
        updatedData.outwardNo || currentData[2], // Keep existing outward number
        updatedData.toWhom || '',
        updatedData.subject || '',
        updatedData.sentBy || '',
        updatedData.receiptDateTime ? new Date(updatedData.receiptDateTime) : currentData[6],
        updatedData.caseClosed || 'No',
        updatedData.fileReference || '',
        parseFloat(updatedData.postalTariff) || '',
        updatedData.dueDate ? new Date(updatedData.dueDate) : (currentData[10] || '')  // Added Due Date
      ];
    }
    
    // Update the row
    sheet.getRange(rowNumber, 1, 1, rowData.length).setValues([rowData]);
    
    Logger.log(`Successfully updated ${updatedData.type} entry at row ${rowNumber}`);
    
    return {
      success: true,
      message: `${updatedData.type} entry updated successfully!`,
      entryId: updatedData.id,
      rowNumber: rowNumber
    };
    
  } catch (error) {
    Logger.log(`Error updating entry: ${error.toString()}`);
    return { success: false, message: `Error updating entry: ${error.toString()}` };
  }
}

function validateUpdateData(updatedData) {
  // Check basic required fields
  if (!updatedData.receiptDateTime || !updatedData.means || !updatedData.subject) {
    return { isValid: false, message: 'Date/Time, Means, and Subject are required' };
  }
  
  // Check entry-type specific fields
  if (updatedData.type === 'Inward') {
    if (!updatedData.fromWhom) {
      return { isValid: false, message: 'From Whom field is required for inward entries' };
    }
    if (!updatedData.takenBy) {
      return { isValid: false, message: 'Taken By field is required for inward entries' };
    }
  }
  
  if (updatedData.type === 'Outward') {
    if (!updatedData.toWhom) {
      return { isValid: false, message: 'To Whom field is required for outward entries' };
    }
    if (!updatedData.sentBy) {
      return { isValid: false, message: 'Sent By field is required for outward entries' };
    }
  }
  
  // Validate date
  try {
    const testDate = new Date(updatedData.receiptDateTime);
    if (isNaN(testDate.getTime())) {
      return { isValid: false, message: 'Invalid date format' };
    }
  } catch (error) {
    return { isValid: false, message: 'Invalid date format' };
  }
  
  return { isValid: true };
}
function debugSearchDropdown() {
  try {
    Logger.log('=== DEBUG SEARCH DROPDOWN ===');
    
    // Test user info
    const userInfo = getCurrentUser();
    Logger.log(`User Info: ${JSON.stringify(userInfo)}`);
    
    // Test spreadsheet access
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log(`Spreadsheet ID: ${ss.getId()}`);
    
    // Test sheet setup
    const setupResult = setupSheets();
    Logger.log(`Setup Result: ${setupResult}`);
    
    // Test sheet access
    const inwardSheet = ss.getSheetByName(CONFIG.INWARD_SHEET);
    const outwardSheet = ss.getSheetByName(CONFIG.OUTWARD_SHEET);
    
    Logger.log(`Inward Sheet: ${inwardSheet ? 'EXISTS' : 'MISSING'} - Rows: ${inwardSheet ? inwardSheet.getLastRow() : 'N/A'}`);
    Logger.log(`Outward Sheet: ${outwardSheet ? 'EXISTS' : 'MISSING'} - Rows: ${outwardSheet ? outwardSheet.getLastRow() : 'N/A'}`);
    
    // Test data retrieval
    const searchResult = getAllEntriesForSearch();
    Logger.log(`Search Result: ${JSON.stringify(searchResult)}`);
    
    return {
      success: true,
      userInfo: userInfo,
      sheetAccess: {
        inward: inwardSheet ? inwardSheet.getLastRow() : 0,
        outward: outwardSheet ? outwardSheet.getLastRow() : 0
      },
      searchResult: searchResult
    };
    
  } catch (error) {
    Logger.log(`Debug Error: ${error.toString()}`);
    return {
      success: false,
      error: error.toString()
    };
  }
}
// =====================================================
// WEEKLY EMAIL NOTIFICATION FUNCTIONS
// =====================================================

function setupWeeklyEmailTrigger() {
  try {
    // Delete existing triggers for this function
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === CONFIG.TRIGGER_FUNCTION_NAME) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new weekly trigger for Saturday at 11 AM
    ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTION_NAME)
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.SATURDAY)
      .atHour(11)
      .create();
    
    Logger.log('Weekly email trigger set up successfully for Saturday 11 AM');
    return { success: true, message: 'Weekly email trigger created successfully' };
    
  } catch (error) {
    Logger.log(`Error setting up weekly trigger: ${error.toString()}`);
    return { success: false, message: error.toString() };
  }
}

function sendWeeklyPendingReport() {
  try {
    Logger.log('=== WEEKLY PENDING REPORT ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pendingEntries = getPendingEntriesForEmail(ss);
    
    if (pendingEntries.length === 0) {
      Logger.log('No pending entries found - no email sent');
      return { success: true, message: 'No pending entries found' };
    }
    
    const emailBody = generatePendingReportEmail(pendingEntries);
    const subject = `${CONFIG.NOTIFICATION_SUBJECT} - ${pendingEntries.length} Pending Entries`;
    
    // Send email to boss
    GmailApp.sendEmail(
      CONFIG.BOSS_EMAIL,
      subject,
      emailBody,
      {
        htmlBody: emailBody,
        name: 'Document Management System'
      }
    );
    
    Logger.log(`Weekly pending report sent to ${CONFIG.BOSS_EMAIL} with ${pendingEntries.length} entries`);
    
    return {
      success: true,
      message: `Report sent successfully to ${CONFIG.BOSS_EMAIL}`,
      pendingCount: pendingEntries.length
    };
    
  } catch (error) {
    Logger.log(`Error sending weekly report: ${error.toString()}`);
    return { success: false, message: error.toString() };
  }
}

function getPendingEntriesForEmail(ss) {
  const pendingEntries = [];
  
  try {
    const confirmations = getConfirmationsSimplified(ss);
    const sheets = [CONFIG.INWARD_SHEET, CONFIG.OUTWARD_SHEET];
    
    for (const sheetName of sheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() <= 1) continue;
      
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Skip entries without subject
        if (!row[4]) continue;
        
        // Check if complete (has all required fields)
        const isComplete = !!(row[1] && row[3] && row[4] && row[6]);
        
        // Check if confirmed
        const confirmKey = `${sheetName}-${i + 1}`;
        const isConfirmed = confirmations.has(confirmKey);
        
        // Entry is pending if complete but not confirmed
        if (isComplete && !isConfirmed) {
          // For Inward entries, also check if action is taken
          if (sheetName === 'Inward' && !row[7]) {
            // Action Required - include this
            pendingEntries.push({
              type: sheetName,
              entryNumber: row[2] || `${sheetName}-${i + 1}`,
              subject: row[4],
              person: row[3],
              dateTime: row[6] ? formatDateTime(row[6]) : 'Unknown Date',
              user: row[5] || 'Unknown User',
              status: 'Action Required',
              actionTaken: row[7] || 'Not specified'
            });
          } else if (sheetName === 'Outward' || (sheetName === 'Inward' && row[7])) {
            // Ready for physical work
            pendingEntries.push({
              type: sheetName,
              entryNumber: row[2] || `${sheetName}-${i + 1}`,
              subject: row[4],
              person: row[3],
              dateTime: row[6] ? formatDateTime(row[6]) : 'Unknown Date',
              user: row[5] || 'Unknown User',
              status: 'Pending Physical Work',
              actionTaken: row[7] || 'N/A'
            });
          }
        }
      }
    }
    
  } catch (error) {
    Logger.log(`Error getting pending entries: ${error.toString()}`);
  }
  
  return pendingEntries;
}

function generatePendingReportEmail(pendingEntries) {
  const currentDate = new Date().toLocaleString('en-US', {
    weekday: 'long',
    year: 'numeric',
    month: 'long',
    day: 'numeric'
  });
  
  let emailBody = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
          .header { background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
          .summary { background-color: #fff3cd; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #ffc107; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
          th, td { padding: 10px; text-align: left; border-bottom: 1px solid #ddd; }
          th { background-color: #f8f9fa; font-weight: bold; }
          .action-required { background-color: #f8d7da; }
          .pending-work { background-color: #fff3cd; }
          .footer { font-size: 12px; color: #666; margin-top: 20px; }
        </style>
      </head>
      <body>
        <div class="header">
          <h2>📋 Weekly Pending Entries Report</h2>
          <p><strong>Generated:</strong> ${currentDate}</p>
          <p><strong>Total Pending Entries:</strong> ${pendingEntries.length}</p>
        </div>
        
        <div class="summary">
          <h3>⚠️ Summary</h3>
          <p>The following entries require attention and are pending completion:</p>
          <ul>
            <li><strong>Action Required:</strong> ${pendingEntries.filter(e => e.status === 'Action Required').length} entries</li>
            <li><strong>Pending Physical Work:</strong> ${pendingEntries.filter(e => e.status === 'Pending Physical Work').length} entries</li>
          </ul>
        </div>
        
        <h3>📋 Detailed Report</h3>
        <table>
          <thead>
            <tr>
              <th>Type</th>
              <th>Entry Number</th>
              <th>Subject</th>
              <th>Person</th>
              <th>Date</th>
              <th>User</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
  `;
  
  pendingEntries.forEach(entry => {
    const rowClass = entry.status === 'Action Required' ? 'action-required' : 'pending-work';
    emailBody += `
      <tr class="${rowClass}">
        <td><strong>${entry.type}</strong></td>
        <td>${entry.entryNumber}</td>
        <td>${entry.subject}</td>
        <td>${entry.person}</td>
        <td>${entry.dateTime}</td>
        <td>${entry.user}</td>
        <td><strong>${entry.status}</strong></td>
      </tr>
    `;
  });
  
  emailBody += `
          </tbody>
        </table>
        
        <div class="footer">
          <p><em>This is an automated report from the Document Management System.</em></p>
          <p><em>Please log in to the system to take appropriate action on these pending entries.</em></p>
        </div>
      </body>
    </html>
  `;
  
  return emailBody;
}

// Function to manually test the email system
function testWeeklyEmail() {
  const result = sendWeeklyPendingReport();
  Logger.log(`Test result: ${JSON.stringify(result)}`);
  return result;
}

// =====================================================
// FINANCIAL REPORT FUNCTION
// =====================================================

function generateFinancialReport(dateFrom, dateTo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const outwardSheet = ss.getSheetByName(CONFIG.OUTWARD_SHEET);
    const linksSheet = ss.getSheetByName(CONFIG.LINKS_SHEET);
    const inwardSheet = ss.getSheetByName(CONFIG.INWARD_SHEET);
    
    if (!outwardSheet) {
      return { success: false, message: 'Outward sheet not found' };
    }
    
    const outwardData = outwardSheet.getDataRange().getValues();
    const links = getEntryLinks(ss);
    const reportData = [];
    
    // Process each outward entry
    for (let i = 1; i < outwardData.length; i++) {
      const row = outwardData[i];
      
      // Skip empty rows
      if (!row[4]) continue; // No subject
      
      // Filter by date range if specified
      if (dateFrom || dateTo) {
        const entryDate = new Date(row[6]); // Date & Time column
        if (dateFrom && entryDate < new Date(dateFrom)) continue;
        if (dateTo && entryDate > new Date(dateTo + 'T23:59:59')) continue;
      }
      
      // Find linked Inward entries
      const outwardEntryId = `Outward-${i + 1}`;
      const linkedEntries = links.get(outwardEntryId) || [];
      let crossNo = '';
      
      // Find the linked Inward entry number
      for (const link of linkedEntries) {
        if (link.linkedEntryId.startsWith('Inward-')) {
          const parts = link.linkedEntryId.split('-');
          const inwardRowNum = parseInt(parts[1]);
          if (inwardSheet && inwardRowNum <= inwardSheet.getLastRow()) {
            const inwardRow = inwardSheet.getRange(inwardRowNum, 1, 1, 10).getValues()[0];
            crossNo = inwardRow[2] || ''; // Inward No column
            break; // Use first linked Inward entry
          }
        }
      }
      
      // Format the report row
      const reportRow = {
        serialNo: reportData.length + 1,
        ackRec: row[6] ? formatDateTime(row[6]) : '', // Date & Time when entered
        crossNo: crossNo, // Linked Inward entry number
        date: row[6] ? formatDateTime(row[6]) : '', // Date & Time
        fileReference: row[8] || '', // File Reference
        address: row[3] || '', // To Whom
        particular: row[4] || '', // Subject
        dueDate: row[10] ? formatDateTime(row[10]) : 'Not Set', // Due Date (new field)
        receiptNumber: row[2] || '', // Outward No
        postalAmount: row[9] || '0' // Postal Tariff
      };
      
      reportData.push(reportRow);
    }
    
    // Calculate total expenditure
    const totalExpenditure = reportData.reduce((sum, row) => {
      return sum + (parseFloat(row.postalAmount) || 0);
    }, 0);
    
    return {
      success: true,
      data: reportData,
      totalExpenditure: totalExpenditure,
      dateRange: {
        from: dateFrom || 'All dates',
        to: dateTo || 'All dates'
      },
      generatedAt: new Date().toLocaleString()
    };
    
  } catch (error) {
    Logger.log(`Error generating financial report: ${error.toString()}`);
    return { success: false, message: error.toString() };
  }
}


