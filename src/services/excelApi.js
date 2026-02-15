import { Client } from '@microsoft/microsoft-graph-client';
import { config } from '../lib/config.js';

function getAuthenticatedClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

export async function analyzeExcelColumns(fileId, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  
  try {
    const result = await client.api(`/me/drive/items/${fileId}/workbook/worksheets/Sheet1/range(address='A1:Z1')`)
      .get();
    
    return result.values?.[0] || [];
  } catch (error) {
    console.error('Error analyzing Excel columns:', error);
    throw error;
  }
}

export async function getExcelRowCount(fileId, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  
  try {
    const result = await client.api(`/me/drive/items/${fileId}/workbook/worksheets/Sheet1/usedRange`)
      .get();
    
    return result.rowCount || 0;
  } catch (error) {
    console.error('Error getting Excel row count:', error);
    throw error;
  }
}

export function extractExcelFileId(url) {
  // Common OneDrive/Excel Online URL patterns
  const match = url.match(/items\/([a-zA-Z0-9!]+)/) || url.match(/s!([a-zA-Z0-9]+)/);
  return match ? match[1] : null;
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
    return await client.api('/subscriptions').post(subscription);
  } catch (error) {
    console.error('Error setting up Excel subscription:', error);
    throw error;
  }
}

export async function fetchExcelColumnData(fileId, columnName, accessToken) {
  const client = getAuthenticatedClient(accessToken);
  
  try {
    const result = await client.api(`/me/drive/items/${fileId}/workbook/worksheets/Sheet1/range(address='${columnName}:${columnName}')`)
      .get();
    
    const values = result.values || [];
    return values.slice(1).map(row => row[0]);
  } catch (error) {
    console.error(`Error fetching Excel column ${columnName} data:`, error);
    throw error;
  }
}


export async function getLastModifiedUser(fileId, accessToken) {
  const client = getAuthenticatedClient(accessToken);

  try {
    const result = await client.api(`/me/drive/items/${fileId}`)
      .select('lastModifiedBy')
      .get();
    
    return result.lastModifiedBy?.user;
  } catch (error) {
    console.error('Error fetching Excel last modified user:', error);
    return null;
  }
}
