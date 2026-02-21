import { cacheService } from '../services/cacheService.js';
import { fetchSheetData } from '../services/dataService.js';
import { tokenService } from '../services/tokenService.js';

export default async function (fastify, opts) {
  const { io } = fastify;
  const db = fastify.mongo.db;
  const tokens = tokenService(db);

  // Google Sheets Webhook
  fastify.post('/google/:sheetId', async (request, reply) => {
    const { sheetId } = request.params;
    const { cell, oldValue, newValue, userEmail, action } = request.body; // 'action' can be 'view', 'open', or 'edit'
    
    if (action === 'edit') {
      fastify.io.to(`sheet:${sheetId}`).emit('sheetUpdate', {
        type: 'google',
        sheetId, cell, oldValue, newValue, user: userEmail,
        timestamp: new Date().toISOString()
      });
      
      // Automatic Cache Invalidation: Real-time data sync!
      cacheService.invalidateBySheetId(sheetId);

      // IMMEDIATE FETCH OPTIMIZATION: Pro-actively refresh data in background
      (async () => {
        try {
          const sheet = await db.collection('sheets').findOne({ id: sheetId });
          if (sheet) {
            console.log(`[ULTRA-SYNC] Pro-actively capturing fresh snapshot for Google sheet: ${sheetId}`);
            // Extract column array safely
            const colArray = Array.isArray(sheet.columns) ? sheet.columns : (sheet.columns?.columns || []);
            const columnsToSync = sheet.customColumns && sheet.customColumns.length > 0 ? sheet.customColumns : colArray;
            
            await fetchSheetData(sheet, columnsToSync, sheet.userId, tokens, true);
          }
        } catch (syncErr) {
          console.error(`[ULTRA-SYNC] Background sync failed for ${sheetId}:`, syncErr.message);
        }
      })();

      // Signal that the cache should be considered stale or the UI should refresh,
      await db.collection('sheets').updateOne(
        { id: sheetId },
        { $set: { lastSync: new Date(), status: 'updated' } }
      );
    }

    // Notify Presence System of External Activity (Works for both Views and Edits)
    if (userEmail) {
      await fastify.broadcastPresence(sheetId, { 
        email: userEmail, 
        name: userEmail.split('@')[0],
        status: action === 'edit' ? 'editing' : 'viewing'
      });
    }

    return { status: 'acknowledged' };
  });

  // Microsoft Excel Webhook
  fastify.post('/excel/:fileId', async (request, reply) => {
    const { fileId } = request.params;

    if (request.query.validationToken) {
      console.log(`[WEBHOOK-VERIFY] Validating Excel subscription for: ${fileId}`);
      return reply.type('text/plain').send(request.query.validationToken);
    }

    console.log(`[WEBHOOK-HIT] Change detected in Excel: ${fileId}`);

    io.to(`sheet:${fileId}`).emit('sheetUpdate', {
      type: 'excel',
      sheetId: fileId,
      timestamp: new Date().toISOString()
    });
    
    // 1. CLEAR MEMORY CACHE
    cacheService.invalidateBySheetId(fileId);
    
    // 2. DEEP INVALIDATION: Mark DB record as dirty so it ignores the persistent cache
    await db.collection('sheets').updateOne(
      { id: fileId },
      { $set: { lastMetrics: {}, status: 'dirty', lastSync: new Date() } }
    );
    
    // 3. IMMEDIATE RE-SYNC: Pro-actively refresh data (Excel)
    (async () => {
      try {
        const sheet = await db.collection('sheets').findOne({ id: fileId });
        if (sheet) {
           console.log(`[WEBHOOK-SYNC] Triggering background refresh for Excel: ${fileId}`);
           const colArray = Array.isArray(sheet.columns) ? sheet.columns : (sheet.columns?.columns || []);
           const columnsToSync = (sheet.customColumns?.length > 0) ? sheet.customColumns : colArray;
           
           // This will also detect new columns because of our new logic in dataService.js
           await fetchSheetData(sheet, columnsToSync, sheet.userId, tokens, true);
           console.log(`[WEBHOOK-SYNC] Success! Data is now hot for: ${fileId}`);
        }
      } catch (syncErr) {
        console.error(`[WEBHOOK-SYNC] Background sync failed:`, syncErr.message);
      }
    })();

    return { status: 'acknowledged' };
  });
}
