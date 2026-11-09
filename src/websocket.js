// Minimal RFC 6455 client for the M365 text protocol. No compression/extensions.
import http from 'node:http';
import https from 'node:https';
import { EventEmitter } from 'node:events';
import { createHash, randomBytes } from 'node:crypto';

export function encodeFrame(payload, opcode = 1, masked = true) {
  const data = Buffer.from(payload);
  const extra = data.length < 126 ? 0 : data.length <= 65535 ? 2 : 8;
  const header = Buffer.alloc(2 + extra + (masked ? 4 : 0));
  header[0] = 0x80 | opcode;
  header[1] = (masked ? 0x80 : 0) | (extra === 0 ? data.length : extra === 2 ? 126 : 127);
  if (extra === 2) header.writeUInt16BE(data.length, 2);
  if (extra === 8) header.writeBigUInt64BE(BigInt(data.length), 2);
  const body = Buffer.from(data);
  if (masked) {
    const mask = randomBytes(4);
    mask.copy(header, 2 + extra);
    for (let i = 0; i < body.length; i++) body[i] ^= mask[i % 4];
  }
  return Buffer.concat([header, body]);
}

export class WebSocket extends EventEmitter {
  static CONNECTING = 0; static OPEN = 1; static CLOSING = 2; static CLOSED = 3;
  readyState = 0;
  buffer = Buffer.alloc(0);
  fragments = [];
  fragmentBytes = 0;
  fragmented = false;
  maxBytes = 32 * 1024 * 1024;
  constructor(address, { headers = {} } = {}) {
    super();
    const url = new URL(address);
    if (!['ws:', 'wss:'].includes(url.protocol)) throw new Error('Expected ws: or wss:');
    const key = randomBytes(16).toString('base64');
    const expected = createHash('sha1').update(key + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11').digest('base64');
    const transport = url.protocol === 'wss:' ? https : http;
    url.protocol = url.protocol === 'wss:' ? 'https:' : 'http:';
    this.request = transport.request(url, { headers: {
      ...headers, Connection: 'Upgrade', Upgrade: 'websocket',
      'Sec-WebSocket-Version': '13', 'Sec-WebSocket-Key': key,
    }});
    this.timer = setTimeout(() => this.fail(new Error('WebSocket upgrade timed out')), 30000);
    this.timer.unref();
    this.request.on('error', () => this.fail(new Error('WebSocket connection failed')));
    this.request.on('response', response => {
      response.resume();
      this.fail(new Error(`WebSocket upgrade rejected: HTTP ${response.statusCode}`));
    });
    this.request.on('upgrade', (response, socket, head) => {
      clearTimeout(this.timer);
      if (this.readyState !== 0) { socket.destroy(); return; }
      if (response.headers['sec-websocket-accept'] !== expected ||
          response.headers.upgrade?.toLowerCase() !== 'websocket' ||
          response.headers['sec-websocket-extensions']) {
        socket.destroy(); this.fail(new Error('Invalid WebSocket upgrade')); return;
      }
      this.socket = socket;
      this.readyState = 1;
      socket.on('data', chunk => {
        try { this.consume(chunk); } catch (error) { this.fail(error); }
      });
      socket.on('error', () => this.fail(new Error('WebSocket transport failed')));
      socket.on('close', () => this.closed());
      this.emit('open');
      if (head.length) {
        try { this.consume(head); } catch (error) { this.fail(error); }
      }
    });
    this.request.end();
  }
  fail(error) {
    if (this.readyState === 3) return;
    this.emit('error', error);
    this.terminate();
  }
  closed() {
    if (this.readyState === 3) return;
    clearTimeout(this.timer);
    this.readyState = 3;
    this.emit('close', this.closeCode ?? 1006, this.closeReason ?? Buffer.alloc(0));
  }
  send(text) {
    if (this.readyState !== 1) throw new Error('WebSocket is not open');
    this.socket.write(encodeFrame(text));
  }
  close() {
    if (this.readyState === 0) { this.terminate(); return; }
    if (this.readyState !== 1) return;
    this.readyState = 2;
    this.socket.write(encodeFrame(Buffer.from([3, 232]), 8));
  }
  terminate() {
    clearTimeout(this.timer);
    this.request.destroy();
    this.socket?.destroy();
    this.closed();
  }
  consume(chunk) {
    this.buffer = Buffer.concat([this.buffer, chunk]);
    while (this.buffer.length >= 2) {
      const b = this.buffer, fin = Boolean(b[0] & 128), opcode = b[0] & 15;
      if ((b[0] & 112) || (b[1] & 128)) throw new Error('Invalid server frame flags');
      let length = b[1] & 127, offset = 2;
      if (length === 126) { if (b.length < 4) return; length = b.readUInt16BE(2); offset = 4; }
      if (length === 127) {
        if (b.length < 10) return;
        const big = b.readBigUInt64BE(2);
        if (big > BigInt(this.maxBytes)) throw new Error('WebSocket frame too large');
        length = Number(big); offset = 10;
      }
      if (length > this.maxBytes) throw new Error('WebSocket frame too large');
      if (opcode >= 8 && (!fin || length > 125)) throw new Error('Invalid control frame');
      if (b.length < offset + length) return;
      const data = b.subarray(offset, offset + length);
      this.buffer = b.subarray(offset + length);
      if (opcode === 8) {
        if (length === 1) throw new Error('Invalid close frame');
        this.closeCode = length >= 2 ? data.readUInt16BE(0) : 1005;
        this.closeReason = data.subarray(2);
        if (this.readyState === 1) this.socket.write(encodeFrame(data, 8));
        this.readyState = 2;
        this.socket.end(); return;
      }
      if (opcode === 9) { this.socket.write(encodeFrame(data, 10)); continue; }
      if (opcode === 10) continue;
      if (![0, 1].includes(opcode) || (opcode === 0 ? !this.fragmented : this.fragmented)) {
        throw new Error('Invalid text continuation');
      }
      this.fragments.push(data); this.fragmentBytes += length;
      if (this.fragmentBytes > this.maxBytes) throw new Error('WebSocket message too large');
      this.fragmented = !fin;
      if (fin) {
        const message = Buffer.concat(this.fragments);
        this.fragments = []; this.fragmentBytes = 0;
        // Reject invalid UTF-8 instead of silently replacing it.
        this.emit('message', new TextDecoder('utf-8', { fatal: true }).decode(message));
      }
    }
  }
}
