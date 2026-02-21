import bcrypt from 'bcryptjs';

export default async function (fastify, opts) {
  console.log('Entering Auth routes registration');
  const db = fastify.mongo.db;

  // SIGNUP (Dual-collection write)
  fastify.post('/signup', {
    schema: {
      body: {
        type: 'object',
        required: ['email', 'password', 'name'],
        properties: {
          email: { type: 'string', format: 'email' },
          password: { type: 'string', minLength: 6 },
          name: { type: 'string' },
          fullName: { type: 'string' }
        }
      }
    }
  }, async (request, reply) => {
    const { email, password, name, fullName } = request.body;

    // 1. Check if user exists
    const existingUser = await db.collection('users').findOne({ email });
    if (existingUser) {
      return reply.status(400).send({ error: 'User already exists' });
    }

    // 2. Hash password
    const hashedPassword = await bcrypt.hash(password, 10);
    const userId = `user_${Date.now()}`;

    // 3. SECURE AUTH: Save to 'users'
    await db.collection('users').insertOne({
      id: userId,
      email,
      password: hashedPassword,
      createdAt: new Date()
    });

    // 4. PROFILE DATA: Save to 'persons'
    await db.collection('persons').insertOne({
      userId: userId,
      name,
      fullName: fullName || name,
      avatar: null,
      bio: '',
      updatedAt: new Date()
    });

    // 5. Generate JWT
    const token = fastify.jwt.sign({ id: userId, email });

    const userProfile = {
      userId,
      email,
      name,
      fullName: fullName || name,
      createdAt: new Date()
    };

    return { 
      status: 'registered', 
      token, 
      user: userProfile 
    };
  });

  // SIGNIN
  fastify.post('/signin', {
    schema: {
      body: {
        type: 'object',
        required: ['email', 'password'],
        properties: {
          email: { type: 'string', format: 'email' },
          password: { type: 'string' }
        }
      }
    }
  }, async (request, reply) => {
    const { email, password } = request.body;

    // 1. Find user
    const user = await db.collection('users').findOne({ email });
    if (!user) {
      return reply.status(401).send({ error: 'Invalid email or password' });
    }

    // 2. Verify password
    const isMatch = await bcrypt.compare(password, user.password);
    if (!isMatch) {
      return reply.status(401).send({ error: 'Invalid email or password' });
    }

    // 3. Get profile
    const person = await db.collection('persons').findOne({ userId: user.id });

    // 4. Generate JWT
    const token = fastify.jwt.sign({ id: user.id, email });

    return { 
      status: 'authenticated', 
      token, 
      user: {
        userId: user.id,
        email: user.email,
        name: person?.name,
        fullName: person?.fullName,
        avatar: person?.avatar
      } 
    };
  });

  // ME / PROFILE (Decrypts token to get user details)
  fastify.get('/me', async (request, reply) => {
    try {
      await request.jwtVerify();
      const userId = request.user.id;
      
      const user = await db.collection('users').findOne({ id: userId });
      const person = await db.collection('persons').findOne({ userId });

      if (!user || !person) {
        return reply.status(404).send({ error: 'User details not found' });
      }

      return {
        userId,
        email: user.email,
        name: person.name,
        fullName: person.fullName,
        avatar: person.avatar,
        bio: person.bio,
        createdAt: user.createdAt
      };
    } catch (err) {
      reply.status(401).send({ error: 'Invalid or expired token' });
    }
  });

  // Microsoft OAuth initiation
  fastify.get('/ms', async (request, reply) => {
    const { config } = await import('../lib/config.js');
    const params = new URLSearchParams({
      client_id: config.microsoft.clientId,
      response_type: 'code',
      redirect_uri: config.microsoft.redirectUri,
      scope: config.microsoft.scopes.join(' '),
      response_mode: 'query'
    });
    const url = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`;
    return reply.redirect(url);
  });

  // Microsoft OAuth callback
  fastify.get('/ms/callback', async (request, reply) => {
    const { code } = request.query;
    if (!code) return reply.status(400).send({ error: 'No code provided' });

    const { config } = await import('../lib/config.js');
    const { tokenService } = await import('../services/tokenService.js');
    const tokens = tokenService(db);

    try {
      const params = new URLSearchParams({
        client_id: config.microsoft.clientId,
        client_secret: config.microsoft.clientSecret,
        code,
        grant_type: 'authorization_code',
        redirect_uri: config.microsoft.redirectUri
      });

      const response = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString()
      });

      const data = await response.json();
      if (!response.ok) throw new Error(data.error_description || 'Microsoft token exchange failed');

      // For MVP/Demo, we'll use a fixed userId or extract from profile if needed
      // Ideally, we'd verify the user via JWT before this, but for the "Connect" flow:
      const userId = 'demo-user'; 

      await tokens.saveTokens(userId, 'excel', {
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        expires_in: data.expires_in
      });

      // Generate a session token for the frontend
      const sessionToken = fastify.jwt.sign({ id: userId, email: 'demo@example.com' });

      // Redirect back to frontend with token in fragment
      const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
      const redirectUrl = `${frontendUrl}/dashboard#token=${sessionToken}&connection=success&provider=microsoft`;
      
      console.log(`[AUTH] MS Callback Success. Redirecting to: ${redirectUrl}`);
      return reply.redirect(redirectUrl);
    } catch (error) {
      console.error('[AUTH] Microsoft OAuth Error:', error);
      return reply.status(500).send({ error: error.message });
    }
  });

  // Google OAuth initiation
  fastify.get('/google', async (request, reply) => {
    const { config } = await import('../lib/config.js');
    const params = new URLSearchParams({
      client_id: config.google.clientId,
      redirect_uri: config.google.redirectUri,
      response_type: 'code',
      scope: config.google.scopes.join(' '),
      access_type: 'offline',
      prompt: 'consent'
    });
    const url = `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
    return reply.redirect(url);
  });

  // Google OAuth callback
  fastify.get('/google/callback', async (request, reply) => {
    const { code } = request.query;
    if (!code) return reply.status(400).send({ error: 'No code provided' });

    const { config } = await import('../lib/config.js');
    const { tokenService } = await import('../services/tokenService.js');
    const tokens = tokenService(db);

    try {
      const params = new URLSearchParams({
        client_id: config.google.clientId,
        client_secret: config.google.clientSecret,
        code,
        grant_type: 'authorization_code',
        redirect_uri: config.google.redirectUri
      });

      const response = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString()
      });

      const data = await response.json();
      if (!response.ok) throw new Error(data.error_description || 'Google token exchange failed');

      // Use JWT user if available, fallback to demo for standalone auth testing
      let userId = 'demo-user';
      try {
         await request.jwtVerify();
         userId = request.user.id;
      } catch (e) {
         console.warn('[AUTH] Google callback using demo-user (no JWT found)');
      }

      await tokens.saveTokens(userId, 'google', {
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        expires_in: data.expires_in
      });

      const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
      const redirectUrl = `${frontendUrl}/dashboard#connection=success&provider=google`;
      
      console.log(`[AUTH] Google Callback Success for user ${userId}. Redirecting.`);
      return reply.redirect(redirectUrl);
    } catch (error) {
      console.error('[AUTH] Google OAuth Error:', error);
      return reply.status(500).send({ error: error.message });
    }
  });

  // DISCONNECT GOOGLE
  fastify.delete('/google/disconnect', async (request, reply) => {
    let userId = 'demo-user';
    try {
      await request.jwtVerify();
      userId = request.user.id;
    } catch (e) {}

    const { tokenService } = await import('../services/tokenService.js');
    const tokens = tokenService(db);
    
    console.log(`[AUTH] Disconnecting Google account for user: ${userId}`);
    await tokens.deleteTokens(userId, 'google');
    
    return { status: 'disconnected', provider: 'google', userId };
  });

  // DISCONNECT MICROSOFT
  fastify.delete('/ms/disconnect', async (request, reply) => {
    let userId = 'demo-user';
    try {
      await request.jwtVerify();
      userId = request.user.id;
    } catch (e) {}

    const { tokenService } = await import('../services/tokenService.js');
    const tokens = tokenService(db);
    
    console.log(`[AUTH] Disconnecting Microsoft account for user: ${userId}`);
    await tokens.deleteTokens(userId, 'excel');
    
    return { status: 'disconnected', provider: 'microsoft', userId };
  });
}
