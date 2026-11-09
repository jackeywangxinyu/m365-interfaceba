import { parseArgs } from 'node:util';
import { startServer } from './server.js';
import { createTokenProvider } from './auth.js';
import { cachedTokenProvider } from './oauth.js';
import { login } from './login.js';

try {
  const { values, positionals } = parseArgs({ allowPositionals: true, options: {
    port: { type: 'string', default: '4141' }, help: { type: 'boolean', short: 'h' },
    'no-lean': { type: 'boolean' }, 'multi-tool': { type: 'boolean' },
  }});
  if (positionals[0] === 'login' && !values.help) {
    if (positionals.length !== 1) throw new Error('Usage: node src/cli.js login');
    await login();
  } else if (values.help) {
    console.log('node src/cli.js [login|serve] [--port 4141] [--no-lean] [--multi-tool]\nSet M365_TOKEN_FILE (or M365_ACCESS_TOKEN), optionally M365_PROXY_KEY and M365_AGENT_ID.');
  } else {
    if (positionals.length > 1 || (positionals[0] && positionals[0] !== 'serve')) throw new Error('Use login or serve');
    const port = Number(values.port);
    if (!/^\d+$/.test(values.port) || port < 0 || port > 65535) throw new Error('Port must be 0..65535');
    const proxy = await startServer({
      getToken: (process.env.M365_TOKEN_FILE || process.env.M365_ACCESS_TOKEN) ? createTokenProvider() : cachedTokenProvider(), port, apiKey: process.env.M365_PROXY_KEY,
      lean: !values['no-lean'], allowMultiTool: Boolean(values['multi-tool']),
      resolveAgent: async () => process.env.M365_AGENT_ID || null,
    });
    console.log(`Base URL: ${proxy.url}/v1\nAPI key: ${proxy.apiKey}`);
    const stop = async () => { await proxy.close(); process.exit(0); };
    process.once('SIGINT', stop); process.once('SIGTERM', stop);
  }
} catch (error) { console.error(error.message); process.exitCode = 1; }
