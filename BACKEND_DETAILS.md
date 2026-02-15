# 🚀 Universal Sheet Connector - Backend Architecture (MongoDB Edition)

This backend is optimized for **zero-cost production deployment** by replacing Redis with MongoDB TTL (Time-To-Live) indexes.

## 🏗️ Technical Stack

- **Runtime**: Node.js (v20+)
- **Framework**: Fastify (High Performance)
- **Database**: MongoDB (Metadata & Configuration Only)
- **Auth**: Stateless JWT (`@fastify/jwt`)
- **Real-time**: Socket.io

## 💎 Zero-Storage Policy

To minimize database costs and maximize privacy:
- **No Raw Data Storage**: Actual spreadsheet content is **NEVER** stored in MongoDB.
- **On-Demand Fetching**: The backend fetches specific columns from Google/Excel APIs only when a preview or merge is requested.
- **Ephemeral Metrics**: Calculated metrics (Sum, Avg) are returned to the frontend and briefly cached in MongoDB (30s) to avoid redundant API calls.

## 📊 Database Collections & TTL Logic

We use MongoDB for both persistent storage and as a high-speed cache:

1.  **`users`**: Persistent user profiles and credentials.
2.  **`sheets`**: Cache for sheet metadata.
    - **TTL**: 60 seconds (`lastSync` index).
    - **Benefit**: Reduces API hits to Google/Microsoft.
3.  **`previews`**: Cache for customized metric previews.
    - **TTL**: 30 seconds (`createdAt` index).
    - **Benefit**: Instant dashboard previews for shared configurations.
4.  **`dashboards`**: Persistent user-saved dashboard configurations.
5.  **`sessions`**: JWT Blacklist for secure logout.
    - **TTL**: 24 hours (`expires` index).

## 🔒 Security Features

- **Stateless Auth**: JWT-based authentication means no server-side sessions.
- **Rate Limiting**: Protected against brute force and DDoS.
- **Webhook Validation**: 
    - Google: Verification tokens in payload.
    - Microsoft: `clientState` and validation tokens.
- **Input Validation**: Strict JSON Schema validation for every endpoint.

## 🌐 API Routes

### Sheet Management
- `POST /api/sheets/connect`: Connect Google/Excel URL.
- `GET /api/sheets/:id`: Fetch sheet details (from cache/API).
- `GET /api/sheets`: List all connected sheets for the user.

### Composer & Metrics
- `POST /api/composer/:id/customize`: Pivot columns and preview metrics.
- `POST /api/composer/:id/merge`: Aggregate data from multiple sources.
- `POST /api/composer/save`: Persist a dashboard configuration.

### Real-time Webhooks
- `POST /api/webhooks/google/:sheetId`: Callback for Apps Script.
- `POST /api/webhooks/excel/:fileId`: Callback for MS Graph.

## 🚀 Deployment Guide

1.  **MongoDB Atlas**: Create a free M0 cluster.
2.  **Environment Variables**: Set `MONGO_URI`, `JWT_SECRET`, and API keys in `.env`.
3.  **Deploy**: Compatible with Vercel, Render, or Railway.

```bash
# Start Production
npm install
npm start
```
