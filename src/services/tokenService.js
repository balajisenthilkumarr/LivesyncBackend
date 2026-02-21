import { config } from '../lib/config.js';

export const tokenService = (db) => ({
  async getValidAccessToken(userId, provider) {
    const tokenDoc = await db.collection('oauth_tokens').findOne({ userId, provider });
    if (!tokenDoc) return null;

    const { access_token, refresh_token, updatedAt, expires_in } = tokenDoc;
    
    // Check if token is expired or expiring in 5 minutes
    const now = new Date();
    const expiryTime = new Date(updatedAt.getTime() + (expires_in * 1000));
    const isBufferExpired = now.getTime() > (expiryTime.getTime() - 300000);

    if (!isBufferExpired) return access_token;

    // Refresh Token Logic
    if (!refresh_token) {
      console.warn(`No refresh token found for ${provider} user ${userId}`);
      return access_token; // Try with existing token as fallback
    }

    try {
      let newTokens;
      console.log(`[TOKEN-SERVICE] Token for ${provider} is expiring or expired. Attempting refresh for user ${userId}...`);
      if (provider === 'google') {
        newTokens = await this.refreshGoogleToken(refresh_token);
      } else if (provider === 'excel') {
        newTokens = await this.refreshMicrosoftToken(refresh_token);
      }

      if (newTokens) {
        console.log(`[TOKEN-SERVICE] Successfully refreshed ${provider} token for user ${userId}.`);
        await this.saveTokens(userId, provider, {
          ...newTokens,
          refresh_token: newTokens.refresh_token || refresh_token // Keep old if not provided
        });
        return newTokens.access_token;
      }
    } catch (error) {
      console.error(`[TOKEN-SERVICE] Failed to refresh ${provider} token for user ${userId}:`, error.message);
    }

    return access_token;
  },

  async refreshGoogleToken(refreshToken) {
    const params = new URLSearchParams({
      client_id: config.google.clientId,
      client_secret: config.google.clientSecret,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    });

    const response = await fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params
    });

    const data = await response.json();
    if (!response.ok) throw new Error(data.error_description || 'Google refresh failed');
    return {
      access_token: data.access_token,
      expires_in: data.expires_in,
      refresh_token: data.refresh_token // Google might send a new one
    };
  },

  async refreshMicrosoftToken(refreshToken) {
    const params = new URLSearchParams({
      client_id: config.microsoft.clientId,
      client_secret: config.microsoft.clientSecret,
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
      scope: config.microsoft.scopes.join(' ')
    });

    const response = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params
    });

    const data = await response.json();
    if (!response.ok) throw new Error(data.error_description || 'Microsoft refresh failed');
    return {
      access_token: data.access_token,
      expires_in: data.expires_in,
      refresh_token: data.refresh_token
    };
  },

  async getAccessToken(userId, provider) {
    // Deprecated: use getValidAccessToken instead
    return this.getValidAccessToken(userId, provider);
  },

  async saveTokens(userId, provider, tokens) {
    const { access_token, refresh_token, expires_in } = tokens;
    
    await db.collection('oauth_tokens').updateOne(
      { userId, provider },
      { 
        $set: { 
          access_token,
          refresh_token,
          expires_in,
          updatedAt: new Date()
        } 
      },
      { upsert: true }
    );
  },

  async deleteTokens(userId, provider) {
    await db.collection('oauth_tokens').deleteOne({ userId, provider });
  }
});
