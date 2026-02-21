import { aggregateMergedMetrics, calculateColumnMetrics } from './aggregator.js';
import { fetchGoogleColumnData, analyzeSheetColumns, fetchGoogleSnapshot } from './googleSheets.js';
import { fetchExcelColumnData, analyzeExcelColumns, fetchExcelSnapshot } from './excelApi.js';
import { cacheService } from './cacheService.js';

export async function fetchSheetData(sheet, columns, userId, tokens, forceRefresh = false) {
  const startTime = Date.now();
  
  // 1. FAST PATH: Check for Materialized Metrics (fully pre-calculated)
  const cacheKey = cacheService.generateKey(sheet.id, columns);
  const cachedMetrics = forceRefresh ? null : cacheService.get(cacheKey);

  if (cachedMetrics) {
    // console.log(`[ULTRA] Materialized cache hit for ${sheet.id}. Speed: <1ms`);
    return cachedMetrics;
  }

  // 1.5 TIER-3 PATH: Check for DB-level Persistent Metrics
  if (!forceRefresh && sheet.lastMetrics && Object.keys(sheet.lastMetrics).length > 0) {
    // Re-verify if all requested columns are in lastMetrics
    const hasAll = columns.every(col => sheet.lastMetrics.hasOwnProperty(col));
    if (hasAll) {
       console.log(`[ULTRA] Persistent DB cache hit for ${sheet.id}. Speed: ~5ms`);
       // Re-hydrate the memory cache
       cacheService.set(cacheKey, sheet.lastMetrics, sheet.id);
       return sheet.lastMetrics;
    }
  }

  // 2. ULTRA FAST PATH: Use a "Full Snapshot" cache to avoid network calls
  let snapshot = forceRefresh ? null : cacheService.getSnapshot(sheet.id);
  const accessToken = await tokens.getValidAccessToken(userId, sheet.type === 'google' ? 'google' : 'excel');

  if (!snapshot) {
    // console.log(`[ULTRA] Snapshot miss. Fetching for ${sheet.id}...`);
    try {
      if (sheet.type === 'google') {
        snapshot = await fetchGoogleSnapshot(sheet.id, accessToken);
      } else if (sheet.type === 'excel') {
        snapshot = await fetchExcelSnapshot(sheet.url, accessToken);
      }
      if (snapshot) cacheService.setSnapshot(sheet.id, snapshot);
    } catch (snapErr) {
      console.warn(`[ULTRA] Snapshot approach failed:`, snapErr.message);
    }
  }

  const results = {};

  if (snapshot) {
    const { headers, data } = snapshot;
    const rowCount = data.length;
    
    // DYNAMIC SYNC: If headers in snapshot differ from sheet.columns, update the DB
    const existingCols = Array.isArray(sheet.columns) ? sheet.columns : (sheet.columns?.columns || []);
    const headersChanged = JSON.stringify(headers) !== JSON.stringify(existingCols);

    if (headersChanged && headers.length > 0) {
      // console.log(`[ULTRA] Detected column mismatch. Syncing DB for ${sheet.id}...`);
      (async () => {
        try {
          const db = await import('../server.js').then(m => m.fastify?.mongo?.db);
          if (db) {
            await db.collection('sheets').updateOne(
              { id: sheet.id },
              { $set: { columns: headers, lastDiscovery: new Date() } }
            );
          }
        } catch (e) {}
      })();
    }

    // Optimized extraction: Find all indices first
    const colMap = columns.map(col => ({ name: col, index: headers.indexOf(col) }));
    
    for (const { name, index } of colMap) {
      if (index === -1) {
        results[name] = { sum: 0, count: 0, avg: 0, error: "Not found" };
        continue;
      }
      
      // Map-less extraction (minor perf gain for huge sheets)
      const columnValues = new Array(rowCount);
      for (let r = 0; r < rowCount; r++) {
        columnValues[r] = data[r][index];
      }
      
      results[name] = calculateColumnMetrics(columnValues, name);
    }
  } else {
    // 3. FALLBACK: Direct Parallel Fetch (Worst case)
    const prefetchedMeta = (sheet.type === 'google') 
      ? await analyzeSheetColumns(sheet.id, accessToken) 
      : await analyzeExcelColumns(sheet.url, accessToken);

    const targetColumns = Array.isArray(prefetchedMeta.columns) ? prefetchedMeta.columns : [];

    const fetchPromises = columns.map(async (col) => {
      try {
        let data = [];
        if (sheet.type === 'google') {
          data = await fetchGoogleColumnData(sheet.id, col, accessToken, prefetchedMeta);
        } else if (sheet.type === 'excel') {
          data = await fetchExcelColumnData(sheet.url, col, accessToken, prefetchedMeta);
        }
        results[col] = calculateColumnMetrics(data, col);
      } catch (colError) {
        results[col] = { sum: 0, count: 0, avg: 0, error: colError.message };
      }
    });
    await Promise.all(fetchPromises);
  }
  
  // Materialize for next time
  cacheService.set(cacheKey, results, sheet.id); 
  
  // Persistence Tier: Save to DB in background
  (async () => {
    try {
      const db = await import('../server.js').then(m => m.fastify?.mongo?.db);
      if (db) {
        await db.collection('sheets').updateOne(
          { id: sheet.id },
          { $set: { materializedAt: new Date(), lastMetrics: results, status: 'synced' } }
        );
      }
    } catch (dbErr) {}
  })();

  console.log(`[ULTRA] Response completed for ${sheet.id} in ${Date.now() - startTime}ms`);
  return results;
}
