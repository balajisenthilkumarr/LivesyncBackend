import { analyzeSheetColumns, getSheetRowCount, extractGoogleSheetId, setupGoogleWebhook } from '../services/googleSheets.js';
import { analyzeExcelColumns, getExcelRowCount, extractExcelFileId, setupExcelSubscription } from '../services/excelApi.js';
import { tokenService } from '../services/tokenService.js';
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
    
    let type, id, name, columns, rowCount, accessToken;

    if (url.includes('docs.google.com/spreadsheets')) {
      type = 'google';
      id = extractGoogleSheetId(url);
      accessToken = await tokens.getValidAccessToken(userId, 'google');
      
      // Fallback: If no account connected, we still try using the Library/API Key
      // This works for "Anyone with link" spreadsheets
      name = 'Google Sheet'; 
      columns = await analyzeSheetColumns(id, accessToken);
      rowCount = await getSheetRowCount(id, accessToken);
      
      // Webhooks/Presence still require OAuth permission for automated deployment
      if (accessToken) {
        await setupGoogleWebhook(id, `${config.webhook.url}/google/${id}`, accessToken);
      }
    } else if (url.includes('1drv.ms') || url.includes('office.com') || url.includes('sharepoint.com')) {
      type = 'excel';
      id = extractExcelFileId(url);
      accessToken = await tokens.getValidAccessToken(userId, 'excel');
      if (!accessToken) return reply.status(401).send({ error: 'Microsoft account not connected' });

      name = 'Excel Workbook';
      columns = await analyzeExcelColumns(id, accessToken);
      rowCount = await getExcelRowCount(id, accessToken);
      await setupExcelSubscription(id, `${config.webhook.url}/excel/${id}`, accessToken);
    } else {
      return reply.status(400).send({ error: 'Unsupported URL format' });
    }

    const sheetData = {
      id,
      userId,
      type,
      name,
      url,
      columns,
      rowCount,
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

    return sheetData;
  });

  fastify.get('/:sheetId', async (request, reply) => {
    const { sheetId } = request.params;
    const userId = request.user?.id || 'demo-user';
    const sheet = await db.collection('sheets').findOne({ id: sheetId, userId });
    
    if (sheet) {
      // Broadcast presence when sheet is accessed via API
      // This serves as an early signal of activity
      fastify.broadcastPresence(sheetId);

      return {
        ...sheet,
        customizationKeyCount: sheet.customColumnCount || 0,
        mergedColumnCount: sheet.mergedColumnCount || 0
      };
    }
    
    return reply.status(404).send({ error: 'Sheet not found' });
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

  fastify.get('/', async (request, reply) => {
    const userId = request.user?.id || 'demo-user';
    const sheets = await db.collection('sheets').find({ userId }).toArray();
    return sheets;
  });
}
