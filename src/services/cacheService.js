/**
 * Simple In-Memory Cache Service
 * Used to store expensive API results (Google/Excel) for a short period.
 */

class CacheService {
  constructor() {
    this.cache = new Map();
    this.sheetToKeys = new Map(); // Index: sheetId -> Set(keys)
    this.DEFAULT_TTL = 30 * 60 * 1000; // Increased to 30 mins for Ultra tier
  }

  set(key, value, sheetId = null, ttl = this.DEFAULT_TTL) {
    const expiresAt = Date.now() + ttl;
    this.cache.set(key, { value, expiresAt, sheetId });
    
    // Auto-index if sheetId provided
    if (sheetId) {
      if (!this.sheetToKeys.has(sheetId)) {
        this.sheetToKeys.set(sheetId, new Set());
      }
      this.sheetToKeys.get(sheetId).add(key);
    }
  }

  setSnapshot(sheetId, snapshot) {
    const key = `snap:${sheetId}`;
    this.set(key, snapshot, sheetId, 20 * 60 * 1000); // 20 mins
  }

  getSnapshot(sheetId) {
    return this.get(`snap:${sheetId}`);
  }

  get(key) {
    const entry = this.cache.get(key);
    if (!entry) return null;

    if (Date.now() > entry.expiresAt) {
      this.delete(key);
      return null;
    }
    return entry.value;
  }

  delete(key) {
    const entry = this.cache.get(key);
    if (entry && entry.sheetId) {
      this.sheetToKeys.get(entry.sheetId)?.delete(key);
    }
    this.cache.delete(key);
  }

  generateKey(sheetId, columns) {
    // Fast key generation: Avoid sorting if one column, otherwise light sort
    if (columns.length === 1) return `m:${sheetId}:${columns[0]}`;
    return `m:${sheetId}:${columns.sort().join('|')}`;
  }

  invalidateBySheetId(sheetId) {
    const keys = this.sheetToKeys.get(sheetId);
    if (keys) {
      console.log(`[ULTRA-CACHE] O(1) Clear for ${sheetId}: ${keys.size} entries.`);
      for (const key of keys) {
        this.cache.delete(key);
      }
      this.sheetToKeys.delete(sheetId);
    }
  }

  flush() {
    this.cache.clear();
    this.sheetToKeys.clear();
  }
}

export const cacheService = new CacheService();
