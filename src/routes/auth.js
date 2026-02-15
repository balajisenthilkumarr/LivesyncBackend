import bcrypt from 'bcryptjs';

export default async function (fastify, opts) {
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
}
