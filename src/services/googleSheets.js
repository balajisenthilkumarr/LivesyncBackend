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
    
    return rows.data.values?.[0] || [];
  } catch (error) {
    console.error('Error analyzing Google Sheet columns:', error);
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

export async function fetchGoogleColumnData(sheetId, columnName, accessToken) {
  const auth = new google.auth.OAuth2();
  if (accessToken) {
    auth.setCredentials({ access_token: accessToken });
  }

  const sheets = google.sheets({ 
    version: 'v4', 
    auth: accessToken ? auth : undefined 
  });
  
  // Convert column name to A1 notation if needed, but usually we just use the column letter
  // For simplicity, we'll assume columnName passed is like 'A', 'B', etc.
  const range = `${columnName}:${columnName}`; 

  try {
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
