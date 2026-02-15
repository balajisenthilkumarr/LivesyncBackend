export default async function (fastify, opts) {
  const { io } = fastify;
  const db = fastify.mongo.db;

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

      // Invalidate Mongo cache only on actual EDITS
      await db.collection('sheets').deleteMany({ id: sheetId });
      await db.collection('previews').deleteMany({ sheetId });
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
      return reply.type('text/plain').send(request.query.validationToken);
    }

    io.to(`sheet:${fileId}`).emit('sheetUpdate', {
      type: 'excel',
      sheetId: fileId,
      timestamp: new Date().toISOString()
    });

    // Notify Presence System of External Activity (Excel)
    // Note: Excel webhooks often require an extra API call to get the specific editor, 
    // so we signal an "Anonymous" external edit for now.
    await fastify.broadcastPresence(fileId, { 
      email: 'external@excel.com', 
      name: 'Excel User',
      status: 'editing'
    });

    // Invalidate Mongo cache for ALL users of this sheet
    await db.collection('sheets').deleteMany({ id: fileId });
    await db.collection('previews').deleteMany({ sheetId: fileId });

    return { status: 'acknowledged' };
  });
}
