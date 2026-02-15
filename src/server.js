import Fastify from 'fastify';
import fastifySocketIO from 'fastify-socket.io';
import fastifyCors from '@fastify/cors';
import fastifyRateLimit from '@fastify/rate-limit';
import fastifyMongo from '@fastify/mongodb';
import fastifyJwt from '@fastify/jwt';
import { config } from './lib/config.js';

const fastify = Fastify({
  logger: true
});

// MongoDB setup
await fastify.register(fastifyMongo, {
  url: config.mongoUri
});

// JWT setup
await fastify.register(fastifyJwt, {
  secret: config.jwtSecret
});

// TTL Indexes setup
const setupTTLCaches = async (db) => {
  // Sheet metadata cache (60s TTL)
  await db.collection('sheets').createIndex(
    { lastSync: 1 }, 
    { expireAfterSeconds: 60 }
  );
  
  // Preview metrics (30s TTL)
  await db.collection('previews').createIndex(
    { createdAt: 1 }, 
    { expireAfterSeconds: 30 }
  );

  // Sessions / JWT Blacklist (24h TTL)
  await db.collection('sessions').createIndex(
    { expires: 1 },
    { expireAfterSeconds: 0 }
  );

  // OAuth Tokens (Unique index for lookups)
  await db.collection('oauth_tokens').createIndex(
    { userId: 1, provider: 1 },
    { unique: true }
  );

  // Sheets (Deep Multi-Tenancy: Unique per user)
  await db.collection('sheets').createIndex(
    { id: 1, userId: 1 },
    { unique: true }
  );
};

// Register plugins
await fastify.register(fastifyCors, { origin: true });
await fastify.register(fastifyRateLimit, { max: 100, timeWindow: '1 minute' });
await fastify.register(fastifySocketIO, { cors: { origin: "*" } });

// Socket.io room & presence management
fastify.ready(err => {
  if (err) return;
  const db = fastify.mongo.db;
  const presence = new Map(); // room -> Map(socketId/externalId -> profile)

  // Expose helpers
  fastify.decorate('getPresence', (sheetId) => {
    if (!presence.has(sheetId)) return [];
    
    // Auto-cleanup on get as well
    const roomMap = presence.get(sheetId);
    for (const [id, prof] of roomMap.entries()) {
      if (prof.type === 'external' && prof.expiresAt < Date.now()) {
        roomMap.delete(id);
      }
    }
    return Array.from(roomMap.values());
  });

  // Expose a helper to broadcast presence from elsewhere (e.g. webhooks)
  fastify.decorate('broadcastPresence', async (sheetId, externalUser = null) => {
    if (!presence.has(sheetId)) presence.set(sheetId, new Map());
    
    if (externalUser) {
      // Add temporary external user (expires in 5 mins)
      const extId = `ext:${externalUser.email}`;
      presence.get(sheetId).set(extId, {
        userId: externalUser.email,
        name: externalUser.name || externalUser.email,
        type: 'external',
        status: externalUser.status || 'viewing', // 'viewing' or 'editing'
        expiresAt: Date.now() + 5 * 60 * 1000
      });
    }

    // Clean up expired external users
    const roomMap = presence.get(sheetId);
    for (const [id, prof] of roomMap.entries()) {
      if (prof.type === 'external' && prof.expiresAt < Date.now()) {
        roomMap.delete(id);
      }
    }

    const viewers = Array.from(roomMap.values());
    fastify.io.to(`sheet:${sheetId}`).emit('presence-update', viewers);
  });

  fastify.io.on('connection', async (socket) => {
    const token = socket.handshake.auth?.token;
    if (!token) return;

    let user;
    try {
      user = fastify.jwt.verify(token);
    } catch (e) {
      return;
    }

    socket.on('join-sheet', async (sheetId) => {
      socket.join(`sheet:${sheetId}`);
      const profile = await db.collection('persons').findOne({ userId: user.id });
      
      if (!presence.has(sheetId)) presence.set(sheetId, new Map());
      presence.get(sheetId).set(socket.id, {
        userId: user.id,
        name: profile?.name || 'Anonymous',
        type: 'dashboard'
      });

      fastify.broadcastPresence(sheetId);
    });
    
    socket.on('disconnecting', () => {
      for (const room of socket.rooms) {
        if (room.startsWith('sheet:')) {
          const sheetId = room.split(':')[1];
          if (presence.has(sheetId)) {
            presence.get(sheetId).delete(socket.id);
            fastify.broadcastPresence(sheetId);
          }
        }
      }
    });
  });
});

// Routes registration
await fastify.register(import('./routes/auth.js'), { prefix: '/api/auth' });
await fastify.register(import('./routes/sheets.js'), { prefix: '/api/sheets' });
await fastify.register(import('./routes/composer.js'), { prefix: '/api/composer' });
await fastify.register(import('./routes/webhooks.js'), { prefix: '/api/webhooks' });

// Health check
fastify.get('/health', async () => ({ status: 'ok', database: 'mongodb' }));

// Start server
const start = async () => {
  try {
    await fastify.ready();
    await setupTTLCaches(fastify.mongo.db);
    
    await fastify.listen({ port: config.port, host: '0.0.0.0' });
    fastify.log.info(`Server listening on ${fastify.server.address().port}`);
  } catch (err) {
    fastify.log.error(err);
    process.exit(1);
  }
};

start();

export { fastify };
