import { aggregateMergedMetrics, calculateColumnMetrics } from '../services/aggregator.js';
import { fetchGoogleColumnData, analyzeSheetColumns } from '../services/googleSheets.js';
import { fetchExcelColumnData, analyzeExcelColumns } from '../services/excelApi.js';
import { tokenService } from '../services/tokenService.js';
import { cacheService } from '../services/cacheService.js';
import { fetchSheetData } from '../services/dataService.js';

export default async function (fastify, opts) {
  console.log('[COMPOSER] Registering routes...');
  const db = fastify.mongo.db;
  const tokens = tokenService(db);

  fastify.get('/ping', async () => ({ status: 'composer active (GET)', timestamp: new Date() }));
  fastify.post('/ping', async () => ({ status: 'composer active (POST)', timestamp: new Date() }));

  // CUSTOMIZE (GET: View all available columns | POST: Save customized keys)
  fastify.get('/:sheetId/customize', async (request, reply) => {
    const { sheetId } = request.params;
    const userId = request.user?.id || 'demo-user';
    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    if (!sheet) return reply.status(404).send({ error: 'Sheet not found' });
    
    return {
      sheetId: sheet.id,
      availableColumns: sheet.columns,
      customColumns: sheet.customColumns || []
    };
  });

  // PREVIEW (Accepts 'columns' in body for live reorganization preview)
  fastify.post('/:sheetId/preview', {
    schema: {
      params: { sheetId: { type: 'string' } },
      body: {
        type: 'object',
        properties: { 
          columns: { type: 'array', items: { type: 'string' } } 
        }
      }
    }
  }, async (request, reply) => {
    const { sheetId } = request.params;
    const bodyColumns = request.body?.columns;
    const userId = request.user?.id || 'demo-user';
    const forceRefresh = request.query?.refresh === 'true';
    
    console.log(`[PREVIEW] Incoming request for sheetId: ${sheetId} (Refresh: ${forceRefresh})`);
    
    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    if (!sheet) {
      return reply.status(404).send({ error: 'Sheet not found' });
    }

    // Safely extract columns array
    const colArray = Array.isArray(sheet.columns) ? sheet.columns : (sheet.columns?.columns || []);
    
    const columnsToFetch = bodyColumns 
      ? bodyColumns
      : (sheet.customColumns?.length > 0 ? sheet.customColumns : colArray.slice(0, 5));

    try {
      const metrics = await fetchSheetData(sheet, columnsToFetch, userId, tokens, forceRefresh);
      return {
        sheetId,
        metrics,
        columns: columnsToFetch
      };
    } catch (error) {
      console.error(`[PREVIEW] Fatal error fetching metrics for ${sheetId}:`, error);
      return reply.status(500).send({ error: 'Failed to fetch dashboard metrics' });
    }
  });

  // DEBUG: Simplified preview test
  fastify.post('/:sheetId/preview-test', async (request) => {
    return { 
      msg: 'Preview test route reached', 
      sheetId: request.params.sheetId,
      originalUrl: request.url 
    };
  });

  fastify.post('/:sheetId/customize', {
    schema: {
      params: { sheetId: { type: 'string' } },
      body: {
        type: 'object',
        required: ['columns'],
        properties: { 
          columns: { type: 'array', items: { type: 'string' } } 
        }
      }
    }
  }, async (request, reply) => {
    const { sheetId } = request.params;
    const { columns } = request.body;
    const userId = request.user?.id || 'demo-user';
    
    // 1. Fetch sheet metadata
    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    if (!sheet) return reply.status(404).send({ error: 'Sheet not found' });

    // 2. Persistent Save
    await db.collection('sheets').updateOne(
      { id: sheetId, userId },
      { $set: { customColumns: columns, customColumnCount: columns.length, updatedAt: new Date() } }
    );

    // 3. Fetch LIVE data
    const liveMetrics = await fetchSheetData(sheet, columns, userId, tokens);

    return {
      sheetId,
      customizedKeys: columns,
      previewMetrics: liveMetrics,
      status: 'saved'
    };
  });

  // MERGE (Aggregate multiple sheets & Persist merge keys)
  fastify.post('/:sheetId/merge', {
    schema: {
      params: { sheetId: { type: 'string' } },
      body: {
        type: 'object',
        required: ['otherSheets', 'columns'],
        properties: { 
          otherSheets: { type: 'array', items: { type: 'string' } },
          columns: { type: 'array', items: { type: 'string' } }
        }
      }
    }
  }, async (request, reply) => {
    const { sheetId } = request.params;
    const { otherSheets, columns } = request.body;
    const allSheetIds = [sheetId, ...otherSheets];
    const userId = request.user?.id || 'demo-user';
    
    // 1. Fetch metadata for ALL sheets (Isolated by userId)
    const sheetsMetadata = await db.collection('sheets')
      .find({ id: { $in: allSheetIds }, userId }).toArray();

    if (sheetsMetadata.length !== allSheetIds.length) {
      return reply.status(400).send({ error: 'One or more sheets not found or unauthorized' });
    }

    // 2. Persistent Save
    // Store columns as objects { name, type } for math power
    const columnConfig = columns.map(col => ({
      name: typeof col === 'string' ? col : col.name,
      type: typeof col === 'string' ? 'add' : (col.type || 'add')
    }));

    await db.collection('sheets').updateOne(
      { id: sheetId, userId },
      { $set: { 
        mergedSheets: otherSheets, 
        mergedColumns: columnConfig, 
        mergedColumnCount: columnConfig.length, 
        updatedAt: new Date() 
      } }
    );

    // 3. Fetch LIVE metrics
    // We pass the full config to the aggregator
    const allSheetMetrics = await Promise.all(
      sheetsMetadata.map(sheet => fetchSheetData(sheet, columnConfig.map(c => c.name), userId, tokens))
    );

    // 4. Aggregate across all sheets (Math logic applied here)
    const mergedMetrics = await aggregateMergedMetrics(allSheetMetrics, columnConfig);
    
    return {
      primarySheet: sheetId,
      mergedSheets: otherSheets,
      mergedKeys: columnConfig,
      previewMetrics: mergedMetrics,
      sheetCount: allSheetIds.length,
      timestamp: new Date().toISOString(),
      status: 'merged_saved'
    };
  });

  // SAVE DASHBOARD (Isolated by userId)
  fastify.post('/save', async (request, reply) => {
    const userId = request.user?.id || 'demo-user';
    const dashboardId = `dash_${Date.now()}`;
    const dashboardConfig = {
      ...request.body,
      id: dashboardId,
      userId,
      createdAt: new Date()
    };

    await db.collection('dashboards').insertOne(dashboardConfig);
    return { dashboardId, userId, url: `/dashboard/${dashboardId}` };
  });
}
