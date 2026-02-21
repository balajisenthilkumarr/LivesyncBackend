import Fastify from 'fastify';
import fastifySocketIO from 'fastify-socket.io';
import fastifyCors from '@fastify/cors';
import fastifyRateLimit from '@fastify/rate-limit';
import fastifyMongo from '@fastify/mongodb';
import fastifyJwt from '@fastify/jwt';
import authRoutes from './routes/auth.js';
import sheetRoutes from './routes/sheets.js';
import composerRoutes from './routes/composer.js';
import webhookRoutes from './routes/webhooks.js';
import { config } from './lib/config.js';

const log = (msg) => console.log(`[${new Date().toISOString()}] ${msg}`);

const fastify = Fastify({
  logger: true
});

// Global request logger for debugging
fastify.addHook('onRequest', async (request, reply) => {
  console.log(`[HTTP] ${request.method} ${request.url}`);
});

// MongoDB setup
log(`Registering MongoDB with URI: ${config.mongoUri}`);
await fastify.register(fastifyMongo, {
  url: config.mongoUri
});
log('MongoDB registered successfully');

// JWT setup
log('Registering JWT');
await fastify.register(fastifyJwt, {
  secret: config.jwtSecret
});
log('JWT registered successfully');

// TTL Indexes setup
const setupTTLCaches = async (db) => {
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

// Routes registration
await fastify.register(authRoutes, { prefix: '/api/auth' });
await fastify.register(sheetRoutes, { prefix: '/api/sheets' });
await fastify.register(composerRoutes, { prefix: '/api/composer' });
await fastify.register(webhookRoutes, { prefix: '/api/webhooks' });

// Health check
fastify.get('/health', async () => ({ status: 'ok', database: 'mongodb' }));

const presence = new Map(); // room -> Map(socketId/externalId -> profile)

// Expose helpers
fastify.decorate('getPresence', (sheetId) => {
  if (!presence.has(sheetId)) return [];
  const roomMap = presence.get(sheetId);
  for (const [id, prof] of roomMap.entries()) {
    if (prof.type === 'external' && prof.expiresAt < Date.now()) {
      roomMap.delete(id);
    }
  }
  return Array.from(roomMap.values());
});

fastify.decorate('broadcastPresence', async (sheetId, externalUser = null) => {
  if (!presence.has(sheetId)) presence.set(sheetId, new Map());
  if (externalUser) {
    const extId = `ext:${externalUser.email}`;
    presence.get(sheetId).set(extId, {
      userId: externalUser.email,
      name: externalUser.name || externalUser.email,
      type: 'external',
      status: externalUser.status || 'viewing',
      expiresAt: Date.now() + 5 * 60 * 1000
    });
  }
  const roomMap = presence.get(sheetId);
  for (const [id, prof] of roomMap.entries()) {
    if (prof.type === 'external' && prof.expiresAt < Date.now()) {
      roomMap.delete(id);
    }
  }
  const viewers = Array.from(roomMap.values());
  fastify.io.to(`sheet:${sheetId}`).emit('presence-update', viewers);
});

// Start server
const start = async () => {
  try {
    log('Waiting for Fastify to be ready...');
    await fastify.ready();
    log('Fastify is ready, setting up socket IO handlers...');
    
    const db = fastify.mongo.db;

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

    log('Setting up TTL caches...');
    if (db) {
      await setupTTLCaches(db);
      log('TTL caches set up');
    }
    
    console.log('[SERVER] Registered Routes:\n', fastify.printRoutes());

    log('Starting to listen...');
    await fastify.listen({ port: config.port, host: '0.0.0.0' });
    log(`Server is LIVE on port ${config.port}`);
  } catch (err) {
    log(`SERVER START ERROR: ${err.message}`);
    console.error(err);
    process.exit(1);
  }
};

start();
export { fastify };
