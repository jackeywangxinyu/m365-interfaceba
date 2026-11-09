function check(ok, message) { if (!ok) throw new Error(message); }
const object = v => v !== null && typeof v === 'object' && !Array.isArray(v);
function optional(o, key, type) { check(o[key] === undefined || typeof o[key] === type, `${key} must be ${type}`); }
export const ChatCompletionRequest = { parse(value) {
  check(object(value), 'Request must be an object');
  optional(value, 'model', 'string'); optional(value, 'stream', 'boolean');
  for (const key of ['temperature', 'top_p', 'max_tokens']) {
    optional(value, key, 'number');
    check(value[key] === undefined || Number.isFinite(value[key]), `${key} must be finite`);
  }
  check(Array.isArray(value.messages) && value.messages.length > 0, 'messages must be a nonempty array');
  for (const m of value.messages) {
    check(object(m) && typeof m.role === 'string', 'Each message requires a role');
    optional(m, 'name', 'string'); optional(m, 'tool_call_id', 'string');
    check(m.content == null || typeof m.content === 'string' || Array.isArray(m.content), 'Invalid message content');
    if (Array.isArray(m.content)) for (const p of m.content) {
      check(object(p) && typeof p.type === 'string', 'Invalid content part'); optional(p, 'text', 'string');
    }
    if (m.tool_calls !== undefined) {
      check(Array.isArray(m.tool_calls), 'tool_calls must be an array');
      for (const t of m.tool_calls) {
        check(object(t) && typeof t.id === 'string' && object(t.function), 'Invalid tool call');
        check(t.type === undefined || t.type === 'function', 'Invalid tool call type');
        check(typeof t.function.name === 'string' && typeof t.function.arguments === 'string', 'Invalid tool call function');
      }
    }
  }
  if (value.tools !== undefined) {
    check(Array.isArray(value.tools), 'tools must be an array');
    for (const t of value.tools) {
      check(object(t) && t.type === 'function' && object(t.function) && typeof t.function.name === 'string', 'Invalid tool definition');
      optional(t.function, 'description', 'string');
    }
  }
  return value;
}};
