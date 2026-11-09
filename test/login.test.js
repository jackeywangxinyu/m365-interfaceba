import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

test('real Chromium CDP captures nativeclient redirect before navigation, exchanges PKCE code and persists',
  {skip: process.platform !== 'linux' || !existsSync('/usr/bin/chromium'), timeout:30000}, async t=>{
    const dir=await mkdtemp(join(tmpdir(),'m365-browser-test-'));
    t.after(()=>rm(dir,{recursive:true,force:true}));
    const wrapper=join(dir,'msedge');
    await writeFile(wrapper,'#!/bin/sh\nexec /usr/bin/chromium --headless --no-sandbox --disable-dev-shm-usage "$@"\n',{mode:0o700});
    const code=`
      import assert from 'node:assert/strict';
      import { login } from './src/login.js';
      import { REDIRECT_URI } from './src/oauth.js';
      let exchanged=false, persisted=false;
      await login({
        authorize:()=>({state:'test-state',verifier:'test-verifier',url:REDIRECT_URI+'?code=test-code&state=test-state'}),
        exchangeCode: async args=>{assert.equal(args.code,'test-code');assert.equal(args.code_verifier,'test-verifier');exchanged=true;return {access_token:'fake'};},
        persist:async tokens=>{assert.equal(tokens.access_token,'fake');persisted=true;}
      });
      assert.ok(exchanged && persisted);
    `;
    const result=await promisify(execFile)(process.execPath,['--input-type=module','-e',code],{
      cwd:new URL('..',import.meta.url),timeout:25000,
      env:{...process.env,EDGE_PATH:wrapper,M365_CONFIG_DIR:join(dir,'config')}
    });
    assert.match(result.stdout,/登录成功/);
});
