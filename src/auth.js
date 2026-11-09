import { readFile } from 'node:fs/promises';
export class AuthRequiredError extends Error {
  constructor(message) { super(message); this.name = 'AuthRequiredError'; }
}
// Read on every turn so replacing the token file takes effect without a restart.
export function createTokenProvider({ token = process.env.M365_ACCESS_TOKEN, file = process.env.M365_TOKEN_FILE } = {}) {
  return async () => {
    let value = token;
    if (file) {
      try { value = (await readFile(file, 'utf8')).trim(); }
      catch { throw new AuthRequiredError('Cannot read M365_TOKEN_FILE'); }
      if (value.startsWith('{')) {
        try { value = JSON.parse(value).access_token; }
        catch { throw new AuthRequiredError('Invalid token JSON'); }
      }
    }
    if (typeof value !== 'string' || !value.trim()) throw new AuthRequiredError('Set M365_TOKEN_FILE or M365_ACCESS_TOKEN to a valid Sydney access token');
    value = value.trim();
    let claims;
    try { claims = JSON.parse(Buffer.from(value.split('.')[1], 'base64url').toString()); }
    catch { throw new AuthRequiredError('Invalid access token'); }
    if (!claims.oid || !claims.tid) throw new AuthRequiredError('Access token requires oid and tid claims');
    if (typeof claims.exp !== 'number' || claims.exp * 1000 <= Date.now() + 30000) {
      throw new AuthRequiredError('Access token expired or expires within 30 seconds; replace the token file');
    }
    return value;
  };
}
