/**
 * OpenAI ⇄ M365 translation.
 *
 * Two asymmetries drive the design:
 *
 * - **M365 is stateful, OpenAI is not.** An OpenAI client resends the whole history
 *   every turn; M365 keeps its own server-side context and charges 600 user messages
 *   per *conversation*. So we key conversations on the first user message and send
 *   only what is new.
 * - **M365 has no tool calls.** Tools go out as a fenced contract in the prompt and
 *   come back as fenced blocks, which `fenced.ts` converts in both directions.
 */

import { createHash, randomUUID } from "node:crypto";
import { ChatCompletionRequest } from "./validation.js";
export { ChatCompletionRequest };
import { buildToolPrompt, parseToolCalls,                                   } from "./fenced.js";


// --- request schema -------------------------------------------------------------

// --- conversation pooling -------------------------------------------------------



















/**
 * Maps an OpenAI client's stateless history onto stateful M365 conversations.
 *
 * Identity is the harness's session id when it names one (the opencode 2 plugin
 * does). Otherwise it is a sha256 of the system prompt plus the first user message:
 * an agent loop keeps resending both, so together they are a stable handle for "this
 * is still the same task" — and the system prompt carries the working directory, so
 * the same opener in two repos stays apart.
 */
export class ConversationPool {
                   conversations = new Map                           ();
                   maxIdleMs        ;

  constructor(options              = {}) {
    this.maxIdleMs = options.maxIdleMs ?? 60 * 60 * 1000;
  }

  resolve(messages               , options                 = {})                    {
    this.evictStale();

    const key = options.sessionId ? `s:${createHash("sha256").update(options.sessionId).digest("hex")}` : fingerprint(messages);
    const existing = this.conversations.get(key);
    if (existing) {
      // A shorter history than we have already sent means the client restarted the
      // task; the server-side context no longer matches.
      if (messages.length < existing.sentMessageCount) existing.sentMessageCount = 0;
      existing.lastAccessedAt = Date.now();
      return existing;
    }

    const state                    = { key, sentMessageCount: 0, lastAccessedAt: Date.now() };
    this.conversations.set(key, state);
    return state;
  }

  get size()         {
    return this.conversations.size;
  }

          evictStale()       {
    const now = Date.now();
    for (const [key, state] of this.conversations) {
      if (now - state.lastAccessedAt >= this.maxIdleMs) this.conversations.delete(key);
    }
  }
}

function fingerprint(messages               )         {
  const system = messages
    .filter((message) => message.role === "system" || message.role === "developer")
    .map(contentOf)
    .join("\n");
  const firstUser = messages.find((message) => message.role === "user");
  const hash = createHash("sha256");
  // Length-prefixed, so no split of the same bytes between the two fields collides.
  for (const part of [system, firstUser ? contentOf(firstUser) : ""]) hash.update(`${part.length}:${part}`);
  return `c:${hash.digest("hex")}`;
}

/** Message content as plain text, flattening the array-of-parts form. */
export function contentOf(message             )         {
  const { content } = message;
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map((part) => (typeof part.text === "string" ? part.text : ""))
      .filter(Boolean)
      .join("\n");
  }
  return "";
}

function hasImagePart(message             )          {
  return Array.isArray(message.content) && message.content.some((part) => part.type === "image_url");
}

// --- prompt construction --------------------------------------------------------

/**
 * Capability catalogues a harness injects into its system prompt, and the tool each
 * one needs in order to be actionable.
 *
 * Measured against opencode 1.18: `<available_skills>` alone was 34k of a 53k-char
 * system prompt — and lean mode disables the `skill` tool, so every entry in it named
 * something the model had no way to invoke. Advertising unusable capabilities is
 * worse than saying nothing: it invites the model to claim it did something it cannot.
 */
const CAPABILITY_CATALOGUES                                                              = [
  { tag: "available_skills", requires: ["skill"] },
  { tag: "available_agents", requires: ["task", "agent"] },
];














/**
 * Trim a harness system prompt down to what this toolset can actually act on.
 *
 * This happens here rather than in the plugin because opencode's own hooks cannot do
 * it: `experimental.chat.system.transform` fires and accepts a new `system` array,
 * but the request still carries the original prompt (verified against 1.18.18), and
 * `agent.<name>.prompt` set from the config hook is ignored the same way. The proxy
 * sees the final request, so it is the one place the change reliably lands.
 */
export function condenseSystemPrompt(
  text        ,
  tools                    ,
  options                  = {},
)         {
  const toolNames = new Set(tools.map((tool) => tool.function.name.toLowerCase()));
  let result = text;

  for (const { tag, requires } of CAPABILITY_CATALOGUES) {
    if (requires.some((name) => toolNames.has(name))) continue;
    result = result.replace(new RegExp(`<${tag}>[\\s\\S]*?</${tag}>`, "g"), "");
  }

  if (options.replaceProseWith !== undefined) {
    // Drop only the harness's own persona prose. Keep the user's instructions —
    // opencode injects project and global rules as PROSE under an
    // `Instructions from: <path>` line — and keep the trailing structured blocks,
    // which carry context the model needs (`<env>` and friends).
    const { leadingBlocks, instructions, blocks } = splitSystemPrompt(result);
    result = [options.replaceProseWith, ...leadingBlocks, instructions, blocks].filter(Boolean).join("\n\n");
  }

  return result.replace(/\n{3,}/g, "\n\n").trim();
}

const BLOCK_PATTERN = /<([a-z_][a-z0-9_-]*)>[\s\S]*?<\/\1>/gi;

/**
 * Split a harness system prompt into: leading prose, the user's own instructions,
 * and the structured blocks the harness appends at the end.
 *
 * The boundary that matters is the **trailing** run of structured blocks, not the
 * first tag-like line. A rules file is Markdown and routinely contains HTML or JSX —
 * cutting at the first `<something>` would drop every rule after it, which is exactly
 * the silent deletion this logic exists to prevent.
 *
 * Everything from the first `Instructions from:` line up to that trailing run is
 * treated as the user's, verbatim, whatever it contains.
 */
function splitSystemPrompt(text        )




  {
  const tailStart = trailingBlockStart(text);
  const head = text.slice(0, tailStart);
  const blocks = text.slice(tailStart).trim();

  const marker = head.search(/^Instructions from:[ \t]*\S/m);
  const proseEnd = marker === -1 ? head.length : marker;

  // Harvest blocks only from the region we are about to discard. Taking them from
  // the whole prompt would duplicate any HTML the user has in their own rules.
  const leadingBlocks = [...head.slice(0, proseEnd).matchAll(BLOCK_PATTERN)].map((match) => match[0]);
  const instructions = marker === -1 ? "" : head.slice(marker).trim();

  return { leadingBlocks, instructions, blocks };
}

/**
 * Offset at which the closing run of structured blocks begins.
 *
 * Walks matched `<tag>…</tag>` pairs from the end, extending backwards while nothing
 * but whitespace separates them, so only blocks that genuinely trail the prompt are
 * treated as the harness's own. Returns `text.length` when there are none.
 */
function trailingBlockStart(text        )         {
  const matches = [...text.matchAll(BLOCK_PATTERN)];
  let start = text.length;

  for (let i = matches.length - 1; i >= 0; i -= 1) {
    const match = matches[i] ;
    const end = match.index + match[0].length;
    if (text.slice(end, start).trim() !== "") break;
    start = match.index;
  }

  return start;
}






/**
 * Build the single prompt string for this turn.
 *
 * Only messages beyond `sentMessageCount` are included — the server already has the
 * rest. The tool block is repeated every turn: it is the contract, and the model
 * drifts away from it if it falls out of view.
 */
export function buildTurnPrompt(
  messages               ,
  tools                    ,
  sentMessageCount        ,
  options                = {},
)         {
  const parts           = [];
  const isFirstTurn = sentMessageCount === 0;

  if (isFirstTurn) {
    const systemText = messages
      .filter((message) => message.role === "system" || message.role === "developer")
      .map(contentOf)
      .filter(Boolean)
      .join("\n\n");
    if (systemText) {
      parts.push(
        condenseSystemPrompt(systemText, tools, {
          ...(options.leanSystemPrompt !== undefined ? { replaceProseWith: options.leanSystemPrompt } : {}),
        }),
      );
    }
  }

  if (tools.length > 0) parts.push(buildToolPrompt(tools));

  // Correlate tool results back to the call that produced them, so the label can say
  // what the output actually is.
  const callsById = new Map                                             ();
  for (const message of messages) {
    for (const call of message.tool_calls ?? []) callsById.set(call.id, call.function);
  }

  for (const message of messages.slice(sentMessageCount)) {
    if (message.role === "system" || message.role === "developer") continue;

    if (message.role === "tool") {
      parts.push(renderToolResult(message, callsById));
      continue;
    }

    if (message.role === "assistant") {
      // The server already has its own replies; only surface an assistant turn the
      // client synthesised.
      const text = contentOf(message);
      if (text) parts.push(text);
      continue;
    }

    // A pasted image with nothing typed: the image itself travels as an upload, but
    // the turn still needs words.
    const text = contentOf(message) || (hasImagePart(message) ? "(See the attached image.)" : "");
    if (text) parts.push(text);
  }

  return parts.filter(Boolean).join("\n\n");
}

/**
 * Wrap a tool result so the model can tell what it is looking at.
 *
 * An unlabelled blob is routinely misread — an `ls` listing taken for file contents,
 * an empty stdout taken for an empty file.
 */
function renderToolResult(
  message             ,
  callsById                                                  ,
)         {
  const call = message.tool_call_id ? callsById.get(message.tool_call_id) : undefined;
  const attributes           = [];

  if (call?.name) attributes.push(`tool="${escapeAttribute(call.name)}"`);
  const summary = summariseArguments(call?.arguments);
  if (summary) attributes.push(`command="${escapeAttribute(summary)}"`);

  const open = attributes.length > 0 ? `<tool_response ${attributes.join(" ")}>` : "<tool_response>";
  return `${open}\n${contentOf(message)}\n</tool_response>`;
}

/** The most identifying argument value, for the label. */
function summariseArguments(raw                    )                     {
  if (!raw) return undefined;
  try {
    const parsed = JSON.parse(raw);
    const value = parsed.command ?? parsed.filePath ?? parsed.path ?? parsed.pattern ?? parsed.query;
    if (typeof value !== "string") return undefined;
    return value.length > 120 ? `${value.slice(0, 117)}…` : value;
  } catch {
    return undefined;
  }
}

function escapeAttribute(value        )         {
  return value.replace(/"/g, "&quot;").replace(/\n/g, " ");
}

// --- response construction ------------------------------------------------------
























export function buildCompletion(result            , options                   )                 {
  const { toolCalls, text } = parseToolCalls(result.text, options.tools, {
    allowMultiple: options.allowMultiple ?? false,
  });

  const hasCalls = toolCalls.length > 0;
  return {
    id: `chatcmpl-${randomUUID()}`,
    object: "chat.completion",
    created: Math.floor(Date.now() / 1000),
    model: options.model,
    choices: [
      {
        index: 0,
        message: {
          role: "assistant",
          content: hasCalls ? null : text,
          ...(hasCalls ? { tool_calls: toolCalls.map(toOpenAiToolCall) } : {}),
        },
        finish_reason: hasCalls ? "tool_calls" : "stop",
      },
    ],
    usage: buildUsage(result),
  };
}

export function toOpenAiToolCall(call                ) {
  return {
    id: `call_${randomUUID().replace(/-/g, "").slice(0, 24)}`,
    type: "function"         ,
    function: { name: call.name, arguments: JSON.stringify(call.arguments) },
  };
}

/**
 * Usage, plus M365-specific extensions.
 *
 * M365 never reports token counts, so the standard fields are zero. The extensions
 * carry what it *does* report: the per-conversation message quota — the closest thing
 * to context-window utilisation available — and the classifier score that rises
 * before the Disengaged filter fires. Clients that ignore unknown fields are fine.
 */
export function buildUsage(result            ) {
  const usage                      = {
    prompt_tokens: 0,
    completion_tokens: 0,
    total_tokens: 0,
  };

  if (result.throttle) {
    const { current, max } = result.throttle;
    usage.x_m365_conversation_messages = current;
    usage.x_m365_conversation_max = max;
    usage.x_m365_conversation_remaining = Math.max(0, max - current);
    usage.x_m365_conversation_pct = max > 0 ? Math.round((current / max) * 100) : 0;
  }
  if (result.contentOrigin) usage.x_m365_content_origin = result.contentOrigin;
  if (typeof result.turnCount === "number") usage.x_m365_turn_count = result.turnCount;
  if (result.scores) {
    usage.x_m365_classifier_scores = result.scores;
    if (typeof result.scores.dea_violation === "number") usage.x_m365_dea_score = result.scores.dea_violation;
  }

  return usage;
}

// --- streaming ------------------------------------------------------------------







/** One `chat.completion.chunk` as a server-sent event frame. */
export function streamChunk(id        , model        , delta            )         {
  const payload = {
    id,
    object: "chat.completion.chunk",
    created: Math.floor(Date.now() / 1000),
    model,
    choices: [
      {
        index: 0,
        delta: {
          ...(delta.content !== undefined ? { content: delta.content } : {}),
          ...(delta.toolCalls ? { tool_calls: delta.toolCalls.map((call, index) => ({ index, ...call })) } : {}),
        },
        finish_reason: delta.finishReason ?? null,
      },
    ],
  };
  return `data: ${JSON.stringify(payload)}\n\n`;
}

export function streamDone()         {
  return "data: [DONE]\n\n";
}

// --- local title generation -----------------------------------------------------

/** Instruction wrappers a harness puts around a title request, which are not the subject. */
const TITLE_INSTRUCTION = /generate|summar|title|respond with|in a few words|short phrase/i;

/**
 * Name a session without spending an M365 conversation on it.
 *
 * A harness titles each session on the main model by default, which would open a second
 * conversation for every session — and the throttle counts conversations started, not
 * messages. This is not as good as a model-written title; it is free, and it keeps
 * the account out of the throttle.
 */
export function generateTitle(messages               )         {
  const subject = messages
    .filter((message) => message.role === "user")
    .map(contentOf)
    .filter((text) => text.trim() !== "" && !TITLE_INSTRUCTION.test(firstLine(text)))
    .join("\n");

  const cleaned = (subject || messages.map(contentOf).join("\n"))
    // Drop fenced code, which is never a good title.
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/`[^`]*`/g, " ")
    .replace(/<tool_response[\s\S]*?<\/tool_response>/g, " ")
    .replace(/[#*_>]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (cleaned === "") return "New session";

  const words = cleaned.split(" ").slice(0, 8).join(" ");
  const title = words.length > 50 ? `${words.slice(0, 47).trimEnd()}…` : words;
  return title.charAt(0).toUpperCase() + title.slice(1);
}

function firstLine(text        )         {
  return text.split("\n", 1)[0] ?? "";
}
