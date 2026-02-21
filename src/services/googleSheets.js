import { google } from 'googleapis';
import { config } from '../lib/config.js';

export async function analyzeSheetColumns(sheetId, accessToken) {
  const auth = new google.auth.OAuth2();
  if (accessToken) {
    auth.setCredentials({ access_token: accessToken });
  }
  
  const sheets = google.sheets({ 
    version: 'v4', 
    auth: accessToken ? auth : undefined 
  });

  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      key: !accessToken ? config.googlelibrarykey : undefined
    });
    
    const sheet = response.data.sheets[0];
    const range = `${sheet.properties.title}!1:1`;
    
    const rows = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range,
      key: !accessToken ? config.googlelibrarykey : undefined
    });
    
    const columns = rows.data.values?.[0] || [];
    return { columns, sheetTitle: sheet.properties.title };
  } catch (error) {
    console.error('[GOOGLE] Error analyzing columns:', error.message);
    throw error;
  }
}

export async function getSheetRowCount(sheetId, accessToken) {
  const auth = new google.auth.OAuth2();
  if (accessToken) {
    auth.setCredentials({ access_token: accessToken });
  }

  const sheets = google.sheets({ 
    version: 'v4', 
    auth: accessToken ? auth : undefined 
  });

  try {
    const response = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      key: !accessToken ? config.googlelibrarykey : undefined
    });
    const sheet = response.data.sheets[0];
    return sheet.properties.gridProperties.rowCount || 0;
  } catch (error) {
    console.error('Error getting Google Sheet row count:', error);
    throw error;
  }
}

export function extractGoogleSheetId(url) {
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * Converts a 0-based index to an Excel/Google column address (A, B, C...)
 */
function indexToColumn(index) {
  let column = '';
  while (index >= 0) {
    column = String.fromCharCode((index % 26) + 65) + column;
    index = Math.floor(index / 26) - 1;
  }
  return column;
}

export async function fetchGoogleColumnData(sheetId, columnName, accessToken, prefetchedColumns = null) {
  const auth = new google.auth.OAuth2();
  if (accessToken) {
    auth.setCredentials({ access_token: accessToken });
  }

  const sheets = google.sheets({ 
    version: 'v4', 
    auth: accessToken ? auth : undefined 
  });
  
  try {
    // 1. Find the column index
    const meta = (prefetchedColumns && typeof prefetchedColumns === 'object' && !Array.isArray(prefetchedColumns))
      ? prefetchedColumns
      : { columns: (Array.isArray(prefetchedColumns) ? prefetchedColumns : (await analyzeSheetColumns(sheetId, accessToken)).columns) };
    
    const columns = meta.columns;
    const sheetTitle = meta.sheetTitle || 'Sheet1';
    const colIndex = columns.indexOf(columnName);
    
    if (colIndex === -1) {
      console.warn(`[GOOGLE] Column "${columnName}" not found in sheet ${sheetId}. Available: ${columns.join(', ')}`);
      return [];
    }
    
    const colAddr = indexToColumn(colIndex);
    const range = `${sheetTitle}!${colAddr}2:${colAddr}100`; 
    console.log(`[GOOGLE-v3] Fetching data for "${columnName}" at ${range}`);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: range,
      key: !accessToken ? config.googlelibrarykey : undefined
    });
    
    // Flatten values (skip header)
    const values = response.data.values || [];
    return values.slice(1).map(row => row[0]);
  } catch (error) {
    console.error(`Error fetching column ${columnName} data:`, error);
    throw error;
  }
}

export async function setupGoogleWebhook(sheetId, callbackUrl, accessToken) {
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  
  const script = google.script({ version: 'v1', auth });

  try {
    // 1. Create a script project bound to the spreadsheet
    const request = {
      resource: {
        title: `ExthingsSync_${sheetId}`,
        parentId: sheetId
      }
    };
    const project = await script.projects.create(request);
    const scriptId = project.data.scriptId;

    // 2. Upload the source code
    const scriptCode = `
      function onOpen() {
        sendNotification('view');
      }

      function onEdit(e) {
        sendNotification('edit', e);
      }

      function sendNotification(action, e) {
        try {
          var payload = {
            action: action,
            sheetId: "${sheetId}",
            userEmail: Session.getActiveUser().getEmail(),
            timestamp: new Date().toISOString()
          };
          
          if (e && e.range) {
            payload.cell = e.range.getA1Notation();
            payload.oldValue = e.oldValue;
            payload.newValue = e.value;
          }

          UrlFetchApp.fetch("${callbackUrl}", {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
          });
        } catch (err) {
          console.error('Error sending LiveSync notification:', err);
        }
      }
    `;

    await script.projects.updateContent({
      scriptId,
      resource: {
        files: [
          {
            name: 'Code',
            type: 'SERVER_JS',
            source: scriptCode
          },
          {
            name: 'appsscript',
            type: 'JSON',
            source: JSON.stringify({
              timeZone: 'GMT',
              exceptionLogging: 'STACKDRIVER',
              runtimeVersion: 'V8'
            })
          }
        ]
      }
    });

    // 3. Create a version and deployment
    const versionResponse = await script.projects.versions.create({
      scriptId,
      resource: { description: 'LiveSync Deployment' }
    });

    await script.projects.deployments.create({
      scriptId,
      resource: {
        versionNumber: versionResponse.data.versionNumber,
        description: 'LiveSync Deployment'
      }
    });

    console.log(`Successfully deployed Apps Script ${scriptId} for sheet ${sheetId}`);
    return scriptId;
  } catch (error) {
    console.error('Error deploying Apps Script:', error);
    throw error;
  }
}

export async function deleteGoogleWebhook(scriptId, accessToken) {
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  const script = google.script({ version: 'v1', auth });

  try {
    console.log(`[GOOGLE] Deleting Apps Script project: ${scriptId}`);
    await script.projects.delete({ scriptId });
    return true;
  } catch (error) {
    console.warn(`[GOOGLE] Failed to delete Apps Script project ${scriptId}:`, error.message);
    return false;
  }
}

export async function fetchGoogleSnapshot(sheetId, accessToken) {
  const auth = new google.auth.OAuth2();
  if (accessToken) {
    auth.setCredentials({ access_token: accessToken });
  }

  const sheets = google.sheets({ 
    version: 'v4', 
    auth: accessToken ? auth : undefined 
  });

  try {
    console.log(`[ULTRA] Fetching full Google Snapshot for: ${sheetId}`);
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: sheetId,
      key: !accessToken ? config.googlelibrarykey : undefined
    });

    const sheetTitle = meta.data.sheets?.[0]?.properties?.title || 'Sheet1';
    
    // FETCH BROAD RANGE: Ensure we get up to ZZ columns and 5000 rows and avoid A:Z restriction
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: sheetId,
      range: `${sheetTitle}!A1:ZZ5000`, 
      key: !accessToken ? config.googlelibrarykey : undefined
    });

    const rows = response.data.values || [];
    const headers = rows[0] || [];
    const dataRows = rows.slice(1);

    console.log(`[ULTRA] Snapshot captured: ${headers.length} columns, ${dataRows.length} rows.`);
    return { headers, data: dataRows, sheetTitle };
  } catch (error) {
    console.error(`[ULTRA] Google Snapshot failed:`, error.message);
    throw error;
  }
}

export async function getLastModifiedUser(sheetId, accessToken) {
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: accessToken });
  const drive = google.drive({ version: 'v3', auth });

  try {
    const response = await drive.files.get({
      fileId: sheetId,
      fields: 'lastModifyingUser'
    });
    return response.data.lastModifyingUser;
  } catch (error) {
    console.error('Error fetching Google last modifying user:', error);
    return null;
  }
}
