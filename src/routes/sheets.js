import { analyzeSheetColumns, getSheetRowCount, extractGoogleSheetId, setupGoogleWebhook, deleteGoogleWebhook, fetchGoogleColumnData } from '../services/googleSheets.js';
import { analyzeExcelColumns, getExcelRowCount, extractExcelFileId, setupExcelSubscription, deleteExcelSubscription, fetchExcelColumnData } from '../services/excelApi.js';
import { fetchSheetData } from '../services/dataService.js';
import { tokenService } from '../services/tokenService.js';
import { cacheService } from '../services/cacheService.js';
import { config } from '../lib/config.js';

export default async function (fastify, opts) {
  const db = fastify.mongo.db;
  const tokens = tokenService(db);

  fastify.post('/connect', {
    schema: {
      body: {
        type: 'object',
        required: ['url'],
        properties: { url: { type: 'string' } }
      }
    }
  }, async (request, reply) => {
    const { url } = request.body;
    const userId = request.user?.id || 'demo-user';
    
    console.log(`[CONNECT] Attempting to connect sheet. URL: ${url}, User: ${userId}`);
    
    let type, id, name, columns, rowCount, accessToken;

    if (url.includes('docs.google.com/spreadsheets')) {
      type = 'google';
      id = extractGoogleSheetId(url);
      console.log(`[CONNECT] Detected Google Sheet. ID: ${id}`);
      
      const accessToken = await tokens.getValidAccessToken(userId, 'google');
      if (!accessToken) {
        return reply.status(401).send({ error: 'OAuth session missing', type: 'google' });
      }

      const meta = await analyzeSheetColumns(id, accessToken);
      columns = meta.columns;
      name = meta.sheetTitle || 'Google Sheet';
      rowCount = await getSheetRowCount(id, accessToken);
    } else if (
      url.includes('1drv.ms') || 
      url.includes('office.com') || 
      url.includes('sharepoint.com') || 
      url.includes('excel.cloud.microsoft')
    ) {
      type = 'excel';
      id = extractExcelFileId(url);
      console.log(`[CONNECT] Detected Excel/Microsoft. ID: ${id}`);
      
      const accessToken = await tokens.getValidAccessToken(userId, 'excel');
      if (!accessToken) {
        return reply.status(401).send({ error: 'OAuth session missing', type: 'excel' });
      }

      const meta = await analyzeExcelColumns(url, accessToken);
      columns = meta.columns;
      name = meta.sheetName || 'Excel Workbook';
      rowCount = await getExcelRowCount(url, accessToken);
    } else {
      console.error(`[CONNECT] ERROR: Unsupported URL format: ${url}`);
      return reply.status(400).send({ error: 'Unsupported URL format' });
    }

    const webhookUrl = `${config.webhook.url}/api/webhooks/${type}/${id}`;
    let webhookId = null;

    try {
      if (type === 'google') {
        webhookId = await setupGoogleWebhook(id, webhookUrl, accessToken);
      } else {
        const sub = await setupExcelSubscription(id, webhookUrl, accessToken);
        webhookId = sub?.id;
      }
    } catch (e) {
      console.warn(`[CONNECT] Webhook setup failed (not critical):`, e.message);
    }

    const sheetData = {
      id,
      userId,
      type,
      name,
      url,
      columns,
      rowCount,
      webhookId,
      customColumns: [],
      customColumnCount: 0,
      mergedColumns: [],
      mergedColumnCount: 0,
      status: 'live',
      lastSync: new Date()
    };

    await db.collection('sheets').updateOne(
      { id, userId },
      { $set: sheetData },
      { upsert: true }
    );

    // PRE-WARM CACHE: Trigger an immediate background sync so the first preview is instant
    (async () => {
      try {
        const colArray = columns || [];
        await fetchSheetData(sheetData, colArray.slice(0, 5), userId, tokens, true);
        console.log(`[ULTRA-WARM] Pre-warmed cache for new sheet: ${id}`);
      } catch (e) {
        console.warn(`[ULTRA-WARM] Failed to pre-warm cache: ${e.message}`);
      }
    })();

    return reply.status(201).send(sheetData);
  });

  fastify.get('/:sheetId', async (request, reply) => {
    const { sheetId } = request.params;
    const userId = request.user?.id || 'demo-user';
    const { analyzeExcelColumns, getExcelRowCount, fetchExcelColumnData } = await import('../services/excelApi.js');
    const { analyzeSheetColumns, getSheetRowCount } = await import('../services/googleSheets.js');
    
    let sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    
    if (!sheet) return reply.status(404).send({ error: 'Sheet not found' });

    try {
      let updated = false;
      const accessToken = await tokens.getValidAccessToken(userId, sheet.type);

      // 1. Lazy Discover columns/rows if empty
      const existingColumns = Array.isArray(sheet.columns) ? sheet.columns : (sheet.columns?.columns || []);
      
      if (existingColumns.length === 0 && (accessToken || sheet.type === 'google')) {
        console.log(`[GET] Performing lazy discovery for sheet: ${sheetId} (${sheet.type})`);
        if (sheet.type === 'google') {
          const meta = await analyzeSheetColumns(sheetId, accessToken);
          sheet.columns = meta.columns;
          sheet.rowCount = await getSheetRowCount(sheetId, accessToken);
        } else { // Excel
          const meta = await analyzeExcelColumns(sheet.url, accessToken);
          sheet.columns = meta.columns;
          sheet.rowCount = await getExcelRowCount(sheet.url, accessToken);
        }
        updated = true;
      }

      // 2. Fetch Sample Data using the Centralized Optimized Engine
      const actualColumns = Array.isArray(sheet.columns) ? sheet.columns : (sheet.columns?.columns || []);
      const columnsToFetch = sheet.customColumns && sheet.customColumns.length > 0 
        ? sheet.customColumns 
        : actualColumns.slice(0, 5);

      let data = {};
      if (columnsToFetch.length > 0) {
        try {
          data = await fetchSheetData(sheet, columnsToFetch, userId, tokens);
        } catch (fetchErr) {
          console.warn(`[GET] Sample data fetch failed for dashboard:`, fetchErr.message);
        }
      }

      if (updated) {
        await db.collection('sheets').updateOne(
          { id: sheetId, userId },
          { $set: { columns: sheet.columns, rowCount: sheet.rowCount, lastSync: new Date() } }
        );
      }

      // Broadcast presence
      fastify.broadcastPresence(sheetId);

      // 3. Webhook Health Indicator
      const isHealthy = sheet.status === 'updated' || sheet.status === 'synced' || (new Date() - new Date(sheet.lastSync) < 10 * 60 * 1000);

      return {
        ...sheet,
        dashboardData: data,
        webhookHealth: {
          active: !!sheet.webhookId,
          status: isHealthy ? 'healthy' : 'idle',
          lastSync: sheet.lastSync
        },
        customizationKeyCount: sheet.customColumnCount || 0,
        mergedColumnCount: sheet.mergedColumnCount || 0
      };
    } catch (error) {
      console.error(`[GET] Error loading sheet details for ${sheetId}:`, error);
      // Return existing metadata even if data fetch fails
      return sheet;
    }
  });

  fastify.get('/:sheetId/live', async (request, reply) => {
    const { sheetId } = request.params;
    const userId = request.user?.id || 'demo-user';

    // Verify user owns/connected this sheet
    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    if (!sheet) return reply.status(404).send({ error: 'Sheet not found' });

    const liveUsers = fastify.getPresence(sheetId);
    return {
      sheetId,
      liveCount: liveUsers.length,
      users: liveUsers
    };
  });

  fastify.get('/:sheetId/data', async (request, reply) => {
    const { sheetId } = request.params;
    const userId = request.user?.id || 'demo-user';
    const { fetchExcelColumnData, analyzeExcelColumns } = await import('../services/excelApi.js');

    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    if (!sheet || sheet.type !== 'excel') {
      return reply.status(404).send({ error: 'Excel sheet not found' });
    }

    const accessToken = await tokens.getValidAccessToken(userId, 'excel');
    if (!accessToken) return reply.status(401).send({ error: 'Microsoft account not connected' });

    try {
      // For verification, we'll fetch the first few columns identified in the sheet
      let columns = sheet.columns;
      if (columns && !Array.isArray(columns) && columns.columns) {
        columns = columns.columns;
      } else if (!columns || (Array.isArray(columns) && columns.length === 0)) {
        const meta = await analyzeExcelColumns(sheet.url, accessToken);
        columns = meta.columns;
      }
      
      const limit = 5; // Fetch first 5 columns for sample
      const data = {};

      for (const col of columns.slice(0, limit)) {
        // Use sheet.url to support Shares API resolution in data fetching too
        data[col] = await fetchExcelColumnData(sheet.url, col, accessToken);
      }

      return {
        sheetId,
        name: sheet.name,
        data
      };
    } catch (error) {
      console.error('Error fetching Excel data:', error);
      return reply.status(500).send({ error: 'Failed to fetch Excel data' });
    }
  });

  fastify.get('/', async (request, reply) => {
    const userId = request.user?.id || 'demo-user';
    console.log(`[SHEETS] Fetching all sheets for user: ${userId}`);
    const sheets = await db.collection('sheets').find({ userId }).toArray();
    console.log(`[SHEETS] Found ${sheets.length} sheets`);
    return sheets;
  });

  fastify.delete('/:sheetId', async (request, reply) => {
    const { sheetId } = request.params;
    const userId = request.user?.id || 'demo-user';

    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    if (!sheet) return reply.status(404).send({ error: 'Sheet not found' });

    console.log(`[DISCONNECT] Removing sheet: ${sheetId} for user: ${userId}`);

    // 1. Unsubscribe Webhook in background
    if (sheet.webhookId) {
      (async () => {
        try {
          const accessToken = await tokens.getValidAccessToken(userId, sheet.type);
          if (sheet.type === 'google') {
            await deleteGoogleWebhook(sheet.webhookId, accessToken);
          } else {
            await deleteExcelSubscription(sheet.webhookId, accessToken);
          }
        } catch (e) {
          console.warn(`[DISCONNECT] Webhook cleanup failed:`, e.message);
        }
      })();
    }

    // 2. Clear Caches
    cacheService.invalidateBySheetId(sheetId);

    // 3. Delete from DB
    await db.collection('sheets').deleteOne({ id: sheetId, userId });

    return { status: 'disconnected', sheetId };
  });
}
