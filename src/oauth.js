import { createHash, randomBytes } from 'node:crypto';
import { mkdir, readFile, writeFile, rename, chmod } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join, dirname } from 'node:path';
import { AuthRequiredError } from './auth.js';
export const CLIENT_ID = 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1';
export const AUTHORITY = 'https://login.microsoftonline.com/common/oauth2/v2.0';
export const REDIRECT_URI = 'https://login.microsoftonline.com/common/oauth2/nativeclient';
export const SCOPES = 'openid profile offline_access https://substrate.office.com/sydney/M365Chat.Read https://substrate.office.com/sydney/sydney.readwrite';
export const CONFIG_DIR = process.env.M365_CONFIG_DIR || join(homedir(), '.config', 'm365-serve');
export const CACHE_FILE = join(CONFIG_DIR, 'tokens.json');
export function authorization() {
  const verifier = randomBytes(32).toString('base64url');
  const state = randomBytes(32).toString('base64url');
  const params = new URLSearchParams({ client_id: CLIENT_ID, response_type: 'code', redirect_uri: REDIRECT_URI,
    scope: SCOPES, code_challenge: createHash('sha256').update(verifier).digest('base64url'),
    code_challenge_method: 'S256', state, prompt: 'select_account' });
  return { verifier, state, url: `${AUTHORITY}/authorize?${params}` };
}
export function codeFromUrl(raw, state) {
  let url; try { url = new URL(raw); } catch { return; }
  if (url.origin + url.pathname !== REDIRECT_URI) return;
  if (url.searchParams.get('state') !== state) throw new Error('OAuth state mismatch');
  if (url.searchParams.has('error')) throw new Error(`Microsoft login failed: ${url.searchParams.get('error')}`);
  return url.searchParams.get('code') || undefined;
}
export async function saveTokens(tokens, file = CACHE_FILE) {
  await mkdir(dirname(file), { recursive: true, mode: 0o700 });
  const tmp = `${file}.${randomBytes(8).toString('hex')}.tmp`;
  await writeFile(tmp, JSON.stringify(tokens), { mode: 0o600, flag: 'wx' });
  await rename(tmp, file);
  if (process.platform !== 'win32') await chmod(file, 0o600);
}
export async function exchange(parameters, { fetchImpl = fetch, previous = {} } = {}) {
  const response = await fetchImpl(`${AUTHORITY}/token`, {
    method: 'POST', signal: AbortSignal.timeout(30000),
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CLIENT_ID, scope: SCOPES, ...parameters }),
  });
  const data = await response.json();
  if (!response.ok || !data.access_token) {
    throw new AuthRequiredError(`Microsoft token request failed (${response.status}, ${data.error || 'invalid_response'}). Run node src/cli.js login`);
  }
  return { access_token: data.access_token, refresh_token: data.refresh_token || previous.refresh_token,
    expires_at: Date.now() + Number(data.expires_in || 3600) * 1000 };
}
export function cachedTokenProvider({ file = CACHE_FILE, fetchImpl = fetch } = {}) {
  let pending;
  async function acquire() {
    let cache;
    try { cache = JSON.parse(await readFile(file, 'utf8')); }
    catch { throw new AuthRequiredError('No valid login cache. Run node src/cli.js login first'); }
    if (cache.access_token && cache.expires_at > Date.now() + 120000) return cache.access_token;
    if (!cache.refresh_token) throw new AuthRequiredError('No refresh token. Run node src/cli.js login');
    const tokens = await exchange({ grant_type: 'refresh_token', refresh_token: cache.refresh_token }, { previous: cache, fetchImpl });
    await saveTokens(tokens, file);
    return tokens.access_token;
  }
  return () => pending ??= acquire().finally(() => { pending = undefined; });
}
