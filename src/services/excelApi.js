import { Client } from '@microsoft/microsoft-graph-client';
import { config } from '../lib/config.js';

function getAuthenticatedClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

/**
 * Encodes a URL for the Microsoft Graph /shares API
 */
function encodeShareId(url) {
  const base64 = Buffer.from(url).toString('base64');
  return 'u!' + base64.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
}

export async function analyzeExcelColumns(url, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  const fileId = extractExcelFileId(url);
  
  if (!fileId) {
    throw new Error('Could not extract docId from the provided URL');
  }

  try {
    console.log(`[EXCEL-v3] Analyzing columns using direct docId: ${fileId}`);
    const workbookBase = `/me/drive/items/${fileId}/workbook`;
    
    // Fetch worksheets to get the first sheet name, then fetch headers
    const sheetsResult = await client.api(`${workbookBase}/worksheets`).get();
    const sheetName = sheetsResult.value?.[0]?.name || 'Sheet1';
    
    const headersResult = await client.api(`${workbookBase}/worksheets/${sheetName}/range(address='A1:Z1')`).get();
    const columns = headersResult.values?.[0] || [];
    
    console.log(`[EXCEL-v3] Successfully fetched ${columns.length} columns from ${sheetName}.`);
    return { columns, workbookBase, sheetName };
  } catch (error) {
    console.error('[EXCEL-v3] Error analyzing columns:', error.statusCode, error.message);
    throw error;
  }
}

export async function getExcelRowCount(url, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  const fileId = extractExcelFileId(url);
  
  if (!fileId) throw new Error('Invalid Excel URL');

  try {
    const workbookBase = `/me/drive/items/${fileId}/workbook`;
    const sheetsResult = await client.api(`${workbookBase}/worksheets`).get();
    const sheetName = sheetsResult.value?.[0]?.name || 'Sheet1';
    const result = await client.api(`${workbookBase}/worksheets/${sheetName}/usedRange`).get();
    return result.rowCount || 0;
  } catch (error) {
    console.error('[EXCEL-v3] Error getting row count:', error.statusCode, error.message);
    throw error;
  }
}

export function extractExcelFileId(url) {
  console.log(`[EXCEL] Extracting ID from URL: ${url}`);
  
  // Pattern 1: docId parameter (excel.cloud.microsoft)
  try {
    const urlObj = new URL(url);
    const docId = urlObj.searchParams.get('docId');
    if (docId) {
      console.log(`[EXCEL] Found docId in query: ${docId}`);
      return docId;
    }
  } catch (e) {}

  // Pattern 2: items/ID
  const itemsMatch = url.match(/items\/([a-zA-Z0-9!_-]+)/);
  if (itemsMatch) {
    console.log(`[EXCEL] Found ID in items path: ${itemsMatch[1]}`);
    return itemsMatch[1];
  }

  // Pattern 3: s!ID (OneDrive short links)
  const shortMatch = url.match(/[?&]s!([a-zA-Z0-9_-]+)/) || url.match(/\/s!([a-zA-Z0-9_-]+)/);
  if (shortMatch) {
    console.log(`[EXCEL] Found ID in short link pattern: ${shortMatch[1]}`);
    return `s!${shortMatch[1]}`;
  }

  // Pattern 4: resid parameter
  try {
    const urlObj = new URL(url);
    const resid = urlObj.searchParams.get('resid');
    if (resid) {
      console.log(`[EXCEL] Found resid: ${resid}`);
      return resid;
    }
  } catch (e) {}

  console.warn(`[EXCEL] Could not extract ID from URL: ${url}`);
  return null;
}

export async function setupExcelSubscription(fileId, callbackUrl, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  
  const subscription = {
    changeType: 'updated',
    notificationUrl: callbackUrl,
    resource: `/me/drive/items/${fileId}/workbook`,
    expirationDateTime: new Date(Date.now() + 86400000).toISOString(),
    clientState: config.webhook.secret
  };
  
  try {
    console.log(`[EXCEL] Setting up subscription for ${fileId}`);
    return await client.api('/subscriptions').post(subscription);
  } catch (error) {
    if (error.statusCode === 400 && error.message?.includes('MSA requests')) {
      console.warn('[EXCEL] NOTE: Live Webhooks (Subscriptions) are not supported for personal (MSA) accounts. The sheet is connected, but updates will only sync on refresh.');
      logToFile('[EXCEL] Subscription skipped: Not supported for MSA accounts.');
    } else {
      console.error('[EXCEL] Error setting up subscription:', error.statusCode, error.code, error.message);
    }
    // Subscription errors shouldn't block the whole connection for demo
    return null; 
  }
}

/**
 * Converts a 0-based index to an Excel column address (A, B, C... Z, AA, AB...)
 */
function indexToColumn(index) {
  let column = '';
  while (index >= 0) {
    column = String.fromCharCode((index % 26) + 65) + column;
    index = Math.floor(index / 26) - 1;
  }
  return column;
}

export async function fetchExcelColumnData(url, columnName, accessToken, prefetchedColumns = null, workbookBase = null) {
  const client = getAuthenticatedClient(accessToken);
  const fileId = extractExcelFileId(url);
  
  try {
    const columnsObj = (prefetchedColumns && typeof prefetchedColumns === 'object' && !Array.isArray(prefetchedColumns))
      ? prefetchedColumns 
      : (await analyzeExcelColumns(url, accessToken));
    
    const columns = columnsObj.columns;
    const base = workbookBase || columnsObj.workbookBase || `/me/drive/items/${fileId}/workbook`;
    
    const colIndex = columns.findIndex(c => c === columnName);
    if (colIndex === -1) {
      console.warn(`[EXCEL-v3] Column "${columnName}" not found. Available: ${columns.join(', ')}`);
      throw new Error(`Column ${columnName} not found`);
    }
    
    const colAddr = indexToColumn(colIndex);
    console.log(`[EXCEL-v3] Fetching "${columnName}" at ${colAddr} using base: ${base}`);

    // Optimization: Skip worksheets lookup if sheetName is already in metadata
    let sheetName = columnsObj.sheetName;
    if (!sheetName) {
      const sheetsResult = await client.api(`${base}/worksheets`).get();
      sheetName = sheetsResult.value?.[0]?.name || 'Sheet1';
    }

    const result = await client.api(`${base}/worksheets/${sheetName}/range(address='${colAddr}2:${colAddr}100')`).get();
    const rows = result.values || [];
    
    console.log(`[EXCEL-v3] Fetched ${rows.length} rows for ${columnName}.`);
    return rows.map(row => row[0]);
  } catch (error) {
    console.error(`[EXCEL-v3] Error fetching column ${columnName} data:`, error.statusCode, error.message);
    throw error;
  }
}


export async function deleteExcelSubscription(subscriptionId, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  try {
    console.log(`[EXCEL] Deleting subscription: ${subscriptionId}`);
    return await client.api(`/subscriptions/${subscriptionId}`).delete();
  } catch (error) {
    console.warn(`[EXCEL] Failed to delete subscription ${subscriptionId}:`, error.message);
    return null;
  }
}

export async function fetchExcelSnapshot(url, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  const fileId = extractExcelFileId(url);
  
  try {
    const workbookBase = `/me/drive/items/${fileId}/workbook`;
    const sheetsResult = await client.api(`${workbookBase}/worksheets`).get();
    const sheetName = sheetsResult.value?.[0]?.name || 'Sheet1';
    
    // DYNAMIC USED RANGE: Detects the exact data bounds (e.g., A1:K50) 
    // This is faster for sparse sheets and ensures 100% coverage for dense ones.
    console.log(`[ULTRA] Detecting used range for: ${sheetName}`);
    const usedRange = await client.api(`${workbookBase}/worksheets/${sheetName}/usedRange`).get();
    const address = usedRange.address || 'A1:AZ1000'; // Bulletproof fallback
    
    console.log(`[ULTRA] Smart snapshot using range: ${address}`);
    const result = await client.api(`${workbookBase}/worksheets/${sheetName}/range(address='${address}')`).get();
    const rows = result.values || [];
    
    const headers = rows[0] || [];
    const dataRows = rows.slice(1);
    
    console.log(`[ULTRA] Snapshot captured: ${headers.length} columns, ${dataRows.length} rows.`);
    return { 
      headers, 
      data: dataRows, 
      sheetName,
      range: address,
      rowCount: rows.length
    };
  } catch (error) {
    console.error(`[ULTRA] Smart Excel Snapshot failed:`, error.statusCode, error.message);
    throw error;
  }
}
