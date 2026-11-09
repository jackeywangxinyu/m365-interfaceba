import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createServer, request } from 'node:http';
import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { startServer } from '../src/server.js';
import { encodeFrame, WebSocket } from '../src/websocket.js';
import { TurnText, cursorMessageId } from '../src/session.js';
import { authorization, codeFromUrl, REDIRECT_URI, exchange, saveTokens, cachedTokenProvider } from '../src/oauth.js';
import { createTokenProvider } from '../src/auth.js';

const token = 'a.' + Buffer.from(JSON.stringify({ oid: 'user', tid: 'tenant', exp: Math.floor(Date.now()/1000)+3600 })).toString('base64url') + '.b';
async function stub(t, answer = '你好') {
  const batches = [], urls = [], sockets = new Set();
  const server = createServer();
  server.on('upgrade', (req, socket) => {
    sockets.add(socket); socket.on('close', () => sockets.delete(socket)); urls.push(req.url);
    assert.equal(req.headers.origin, 'https://m365.cloud.microsoft');
    assert.match(req.headers['user-agent'], /Mozilla/);
    const accept = createHash('sha1').update(req.headers['sec-websocket-key']+'258EAFA5-E914-47DA-95CA-C5AB0DC85B11').digest('base64');
    socket.write(`HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Accept: ${accept}\r\n\r\n`);
    let buffer = Buffer.alloc(0);
    socket.on('data', chunk => {
      buffer = Buffer.concat([buffer, chunk]);
      while (buffer.length >= 2) {
        let n = buffer[1] & 127, off = 2;
        if (n === 126) { if(buffer.length<4) return; n=buffer.readUInt16BE(2); off=4; }
        else if (n === 127) { if(buffer.length<10) return; n=Number(buffer.readBigUInt64BE(2)); off=10; }
        if(buffer.length<off+4+n) return;
        const mask=buffer.subarray(off,off+4), data=Buffer.from(buffer.subarray(off+4,off+4+n)), op=buffer[0]&15;
        for(let i=0;i<n;i++) data[i]^=mask[i%4];
        buffer=buffer.subarray(off+4+n);
        if(op===8) { socket.end(encodeFrame(data,8,false)); continue; }
        if(op!==1) continue;
        const frames=data.toString().split('\x1e').filter(Boolean).map(JSON.parse); batches.push(frames);
        if(frames[0].protocol) socket.write(encodeFrame('{}\x1e',1,false));
        if(frames[0].target==='chat') {
          assert.equal(frames[1]?.target, 'Metrics', 'Metrics must share the chat WebSocket message');
          // `answer` is either plain text or the exact list of update arguments to send.
          const updates = Array.isArray(answer) ? answer
            : [answer === 'DISENGAGED' ? { messages:[{author:'bot', messageType:'Disengaged'}] } : {writeAtCursor:answer}];
          const payload=updates.map(u=>JSON.stringify({type:1,target:'update',arguments:[u]})+'\x1e').join('')+JSON.stringify({type:3,invocationId:'0'})+'\x1e';
          // Exercise fragmented messages and TCP frames split in the middle of a header.
          const a=encodeFrame(payload.slice(0,8),1,false); a[0]&=127;
          const b=encodeFrame(payload.slice(8),0,false);
          socket.write(a.subarray(0,1)); socket.write(Buffer.concat([a.subarray(1),b]));
        }
      }
    });
  });
  await new Promise(r=>server.listen(0,'127.0.0.1',r));
  t.after(async()=>{ for(const s of sockets)s.destroy(); await new Promise(r=>server.close(r)); });
  const proxy=await startServer({getToken:async()=>token,endpoint:`ws://127.0.0.1:${server.address().port}`,apiKey:'test-key'});
  t.after(()=>proxy.close());
  const post=async body=>fetch(proxy.url+'/v1/chat/completions',{method:'POST',headers:{Authorization:'Bearer test-key','Content-Type':'application/json','x-m365-session-id':'test'},body:JSON.stringify(body)});
  return {proxy,post,batches,urls};
}

test('HTTP security, validation and routes',async t=>{
  const {proxy,post}=await stub(t);
  assert.equal((await fetch(proxy.url+'/health')).status,200);
  assert.equal((await fetch(proxy.url+'/v1/models')).status,401);
  assert.equal((await fetch(proxy.url+'/health',{headers:{Origin:'https://example.com'}})).status,403);
  assert.equal(await new Promise((resolve, reject) => { const req=request(proxy.url+'/health', {headers:{Host:'evil.example'}},res=>{res.resume();resolve(res.statusCode)}); req.on('error',reject);req.end(); }),403);
  const headers={Authorization:'Bearer test-key'};
  const models=await (await fetch(proxy.url+'/v1/models',{headers})).json();
  assert.ok(models.data.some(m=>m.id==='m365-copilot'));
  assert.equal((await post({messages:[]})).status,400);
  assert.equal((await post({messages:[{role:'user',content:123}]})).status,400);
  assert.equal((await fetch(proxy.url+'/missing',{headers})).status,404);
});
test('real HTTP -> WebSocket -> SignalR round trip, streaming, session reuse',async t=>{
  const {post,batches,urls}=await stub(t);
  const first=await (await post({model:'m365-copilot',messages:[{role:'user',content:'hello'}]})).json();
  assert.equal(first.choices[0].message.content,'你好');
  const second=await post({model:'m365-copilot',stream:true,messages:[{role:'user',content:'hello'},{role:'assistant',content:'你好'},{role:'user',content:'next'}]});
  const text=await second.text(); assert.match(text,/你好/); assert.match(text,/\[DONE\]/);
  assert.equal(new URL(urls[0],'http://x').searchParams.get('ConversationId'),new URL(urls[1],'http://x').searchParams.get('ConversationId'));
  const chats=batches.filter(f=>f[0].target==='chat');
  assert.equal(chats[0][0].arguments[0].isStartOfSession,true);
  assert.equal(chats[1][0].arguments[0].isStartOfSession,false);
  assert.match(chats[1][0].arguments[0].message.text,/next/);
  assert.doesNotMatch(chats[1][0].arguments[0].message.text,/hello/);
});
test('tool response converted to OpenAI call',async t=>{
  const {post}=await stub(t,'```read\nfilePath: /tmp/demo\n```');
  const response=await (await post({messages:[{role:'user',content:'read file'}],tools:[{type:'function',function:{name:'read',parameters:{type:'object',properties:{filePath:{type:'string'}}}}}]})).json();
  assert.equal(response.choices[0].finish_reason,'tool_calls');
  assert.equal(response.choices[0].message.tool_calls[0].function.name,'read');
});
test('Disengaged is surfaced as error',async t=>{
  const {post}=await stub(t,'DISENGAGED');
  const response=await post({messages:[{role:'user',content:'hi'}]});
  assert.equal(response.status,502); assert.equal((await response.json()).error.code,'m365_disengaged');
});
// M365's real stream shape: each bot message's first token arrives ONLY in a snapshot
// that carries a cursor; writeAtCursor deltas (no message id) extend that message.
const cur=id=>({j:`$['${id}'].adaptiveCards[0].body[0].text`,p:-1});
const snap=(id,text,cursor=true)=>({...(cursor?{cursor:cur(id)}:{}),messages:[{author:'bot',messageId:id,text}]});
const ONE_MESSAGE=[snap('m1','re'),{writeAtCursor:'vert'},{writeAtCursor:' the last commit'}];
const TWO_MESSAGES=[
  snap('m1','Let'),{writeAtCursor:' me check.'},snap('m1','Let me check.',false),
  snap('m2','re'),{writeAtCursor:'vert it'},snap('m2','revert it',false),
];
test('the head token that arrives only in a snapshot is kept (revert, not vert)',async t=>{
  const {post}=await stub(t,ONE_MESSAGE);
  const response=await (await post({messages:[{role:'user',content:'undo'}]})).json();
  assert.equal(response.choices[0].message.content,'revert the last commit');
});
test('a multi-message turn keeps every message head, in order',async t=>{
  const {post}=await stub(t,TWO_MESSAGES);
  const response=await (await post({messages:[{role:'user',content:'undo'}]})).json();
  assert.equal(response.choices[0].message.content,'Let me check.\n\nrevert it');
});
test('streaming sends the head token too, exactly once',async t=>{
  const {post}=await stub(t,TWO_MESSAGES);
  const sse=await (await post({stream:true,messages:[{role:'user',content:'undo'}]})).text();
  const content=sse.split('\n').filter(l=>l.startsWith('data: {')).map(l=>JSON.parse(l.slice(6)).choices[0].delta.content??'').join('');
  assert.equal(content,'Let me check.\n\nrevert it');
});
test('TurnText routes deltas by the last cursor, not the last snapshot',()=>{
  const text=new TurnText();
  text.cursor('m1'); text.snapshot('m1','one');
  text.cursor('m2'); text.snapshot('m2','two');
  text.snapshot('m1','one!'); text.delta(' more');
  assert.equal(text.text,'one!\n\ntwo more');
  const legacy=new TurnText(); legacy.delta('plain'); legacy.delta(' deltas');
  assert.equal(legacy.text,'plain deltas');
  assert.equal(cursorMessageId("$['abc-1'].adaptiveCards[0].body[0].text"),'abc-1');
  assert.equal(cursorMessageId(undefined),null);
});
test('WebSocket accepts binary frames like the ws package does',()=>{
  const ws=Object.create(WebSocket.prototype);
  Object.assign(ws,{buffer:Buffer.alloc(0),fragments:[],fragmentBytes:0,fragmented:false,binary:false,maxBytes:1e6});
  const got=[]; ws.emit=(event,message)=>{ if(event==='message') got.push(message); };
  ws.consume(Buffer.concat([encodeFrame(Buffer.from('{"type":6}\x1e'),2,false),encodeFrame('text',1,false)]));
  assert.deepEqual(got,['{"type":6}\x1e','text']);
});
test('WebSocket reads every payload length across the 7/16/64-bit boundaries',async t=>{
  // A 127-byte payload is sent with the 16-bit length 127; it used to be re-read as
  // the 64-bit marker and fail with "WebSocket frame too large".
  const sizes=[0,125,126,127,128,65535,65536];
  const server=createServer(), sockets=new Set();
  server.on('upgrade',(req,socket)=>{
    sockets.add(socket);
    const accept=createHash('sha1').update(req.headers['sec-websocket-key']+'258EAFA5-E914-47DA-95CA-C5AB0DC85B11').digest('base64');
    socket.write(`HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Accept: ${accept}\r\n\r\n`);
    socket.write(Buffer.concat(sizes.map(n=>encodeFrame('x'.repeat(n),1,false))));
  });
  await new Promise(r=>server.listen(0,'127.0.0.1',r));
  t.after(async()=>{ for(const s of sockets)s.destroy(); await new Promise(r=>server.close(r)); });
  const ws=new WebSocket(`ws://127.0.0.1:${server.address().port}`);
  const got=await new Promise((resolve,reject)=>{
    const lengths=[];
    ws.on('error',reject);
    ws.on('message',m=>{ lengths.push(m.length); if(lengths.length===sizes.length) resolve(lengths); });
  });
  ws.terminate();
  assert.deepEqual(got,sizes);
});
test('PKCE, state validation and exact redirect matching',()=>{
  const auth=authorization(); const u=new URL(auth.url);
  assert.equal(u.searchParams.get('code_challenge'),createHash('sha256').update(auth.verifier).digest('base64url'));
  assert.equal(codeFromUrl(`${REDIRECT_URI}?code=ok&state=${auth.state}`,auth.state),'ok');
  assert.throws(()=>codeFromUrl(`${REDIRECT_URI}?code=x&state=wrong`,auth.state),/state/);
  assert.equal(codeFromUrl('https://evil.example/oauth2/nativeclient?code=x',auth.state),undefined);
});
test('refresh token rotation, caching, concurrent refresh and restart',async t=>{
  const dir=await mkdtemp(join(tmpdir(),'m365-test-')); t.after(()=>rm(dir,{recursive:true,force:true}));
  const file=join(dir,'tokens.json'); await saveTokens({access_token:'old',refresh_token:'refresh1',expires_at:0},file);
  let calls=0;
  const provider=cachedTokenProvider({file,fetchImpl:async(url,options)=>{
    calls++; assert.equal(options.body.get('grant_type'),'refresh_token'); assert.equal(options.body.get('refresh_token'),'refresh1');
    return {ok:true,json:async()=>({access_token:token,refresh_token:'refresh2',expires_in:3600})};
  }});
  assert.deepEqual(await Promise.all([provider(),provider()]),[token,token]); assert.equal(calls,1);
  assert.equal(JSON.parse(await readFile(file,'utf8')).refresh_token,'refresh2');
  assert.equal(await cachedTokenProvider({file,fetchImpl:()=>assert.fail('must use cache')})(),token);
  await assert.rejects(exchange({},{fetchImpl:async()=>({ok:false,status:400,json:async()=>({error:'invalid_grant'})})}),/login/);
});
test('manual token validation',async()=>{
  assert.equal(await createTokenProvider({token})(),token);
  await assert.rejects(createTokenProvider({token:'invalid'}),/Invalid/);
});
