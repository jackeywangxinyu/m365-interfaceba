// Interactive OAuth PKCE through Chrome DevTools Protocol, using only Node built-ins.
import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { mkdir, readFile, unlink } from 'node:fs/promises';
import { join, basename } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { authorization, codeFromUrl, exchange, saveTokens, REDIRECT_URI, CONFIG_DIR } from './oauth.js';

export function browserPath() {
  if (process.env.EDGE_PATH) {
    const executable = process.env.EDGE_PATH;
    if (!/^(msedge(?:\.exe)?|microsoft-edge(?:-stable)?|Microsoft Edge)$/i.test(basename(executable))) {
      throw new Error('EDGE_PATH must point to Microsoft Edge (msedge.exe / microsoft-edge)');
    }
    if (!existsSync(executable)) throw new Error('Microsoft Edge was not found at EDGE_PATH');
    return executable;
  }
  const paths = process.platform === 'win32' ? [
    join(process.env['PROGRAMFILES(X86)'] || 'C:\\Program Files (x86)', 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    join(process.env.PROGRAMFILES || 'C:\\Program Files', 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    ...(process.env.LOCALAPPDATA ? [join(process.env.LOCALAPPDATA, 'Microsoft', 'Edge', 'Application', 'msedge.exe')] : []),
  ] : process.platform === 'darwin' ? [
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
  ] : ['/usr/bin/microsoft-edge', '/usr/bin/microsoft-edge-stable', '/opt/microsoft/msedge/msedge'];
  const found = paths.find(p => existsSync(p));
  if (!found) throw new Error('Microsoft Edge not found. Install Edge or set EDGE_PATH to msedge.exe');
  return found;
}

export async function login({ authorize = authorization, exchangeCode = exchange, persist = saveTokens } = {}) {
  const auth = authorize();
  const profile = join(CONFIG_DIR, 'browser-profile');
  await mkdir(profile, { recursive: true, mode: 0o700 });
  const portFile = join(profile, 'DevToolsActivePort');
  await unlink(portFile).catch(() => {});
  const child = spawn(browserPath(), ['--remote-debugging-port=0', '--remote-debugging-address=127.0.0.1',
    `--user-data-dir=${profile}`, '--no-first-run', '--no-default-browser-check', 'about:blank'], { stdio: 'ignore', windowsHide: false });
  let launchError, exited = false, ws;
  child.on('error', error => { launchError = error; });
  child.on('exit', () => { exited = true; });
  try {
    let port;
    for (let i = 0; i < 150; i++) {
      if (launchError) throw new Error(`Browser launch failed: ${launchError.message}`);
      try { port = Number((await readFile(portFile, 'utf8')).split('\n')[0]); } catch {}
      if (port > 0 && port <= 65535) break;
      if (exited) throw new Error('Browser exited. Close the m365-serve login browser window and retry');
      await delay(200);
    }
    if (!port) throw new Error('Browser debugging port did not become ready');
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`, { signal: AbortSignal.timeout(5000) })).json();
    const target = targets.find(t => t.type === 'page');
    if (!target) throw new Error('Browser has no page target');
    const debugUrl = new URL(target.webSocketDebuggerUrl);
    if (!['127.0.0.1', 'localhost', '[::1]'].includes(debugUrl.hostname)) throw new Error('Nonlocal browser debugging URL refused');
    ws = new WebSocket(debugUrl);
    await new Promise((resolve, reject) => {
      ws.addEventListener('open', resolve, { once: true });
      ws.addEventListener('error', () => reject(new Error('Cannot connect to browser debugger')), { once: true });
    });
    let id = 0;
    const pending = new Map();
    const call = (method, params = {}) => new Promise((resolve, reject) => {
      const current = ++id;
      const timer = setTimeout(() => { pending.delete(current); reject(new Error(`Browser command timed out: ${method}`)); }, 10000);
      pending.set(current, { resolve, reject, timer });
      ws.send(JSON.stringify({ id: current, method, params }));
    });
    let finishCode, failCode;
    const codePromise = new Promise((resolve, reject) => { finishCode = resolve; failCode = reject; });
    // Install a rejection handler immediately while CDP setup is in flight.
    codePromise.catch(() => {});
    const timer = setTimeout(() => failCode(new Error('Login timed out after 10 minutes')), 600000);
    ws.addEventListener('close', () => failCode(new Error('Login browser disconnected')));
    ws.addEventListener('message', event => {
      let message; try { message = JSON.parse(event.data); } catch { return; }
      const request = pending.get(message.id);
      if (request) {
        clearTimeout(request.timer); pending.delete(message.id);
        if (message.error) request.reject(new Error(`Browser command failed: ${message.error.message}`));
        else request.resolve(message.result);
      }
      if (message.method === 'Fetch.requestPaused') {
        const { requestId, request: intercepted } = message.params;
        try {
          const code = codeFromUrl(intercepted.url, auth.state);
          if (code) {
            // Stop the redirect before it leaves the browser; the code is consumed here.
            call('Fetch.fulfillRequest', { requestId, responseCode: 200,
              responseHeaders: [{ name: 'Content-Type', value: 'text/plain; charset=utf-8' }],
              body: Buffer.from('Login received. Return to the terminal.').toString('base64') }).catch(() => {});
            finishCode(code);
          } else call('Fetch.continueRequest', { requestId }).catch(failCode);
        } catch (error) { failCode(error); }
      }
    });
    try {
      await call('Fetch.enable', { patterns: [{ urlPattern: `${REDIRECT_URI}*`, requestStage: 'Request' }] });
      console.log('请在打开的 Microsoft Edge 中完成微软登录和 MFA。');
      await call('Page.navigate', { url: auth.url });
      const code = await codePromise;
      const tokens = await exchangeCode({ grant_type: 'authorization_code', code, redirect_uri: REDIRECT_URI, code_verifier: auth.verifier });
      await persist(tokens);
      console.log('登录成功，令牌已缓存。可运行 node src/cli.js serve');
    } finally {
      clearTimeout(timer);
      for (const item of pending.values()) { clearTimeout(item.timer); item.reject(new Error('Login finished')); }
    }
  } finally {
    ws?.close();
    child.kill();
  }
}
