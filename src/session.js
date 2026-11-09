/**
 * The M365 Copilot transport: SignalR over WebSocket.
 *
 * Things that are load-bearing and non-obvious, all of which fail *silently* when
 * you get them wrong:
 *
 * - The access token goes in the WebSocket URL **query string**, not a header.
 * - A custom-header WebSocket transport plus a browser-shaped
 *   `Origin`/`User-Agent` is what the server accepts.
 * - Frames are JSON terminated by `0x1E` (record separator), and one WebSocket
 *   message can carry several of them.
 * - A `Metrics` frame must ride in the **same send** as the chat frame. Omit it and
 *   the turn simply never produces output — no error, just silence.
 * - Only bot messages with **no** `messageType` carry the answer. Everything with one
 *   is control traffic.
 * - `messageType: "Disengaged"` is a refusal with empty content. It looks exactly
 *   like rate limiting and must not be retried into.
 *
 * Each turn opens a fresh socket but reuses the conversation id, so the server keeps
 * threading context. The 600-message cap is per conversation, which is why we reuse.
 */

import { WebSocket } from "./websocket.js";
import { randomUUID } from "node:crypto";
import { createLogger, trunc } from "./log.js";
import { uploadImage,                      } from "./image.js";

const log = createLogger("session");

/** Record separator: every SignalR frame ends with this byte. */
const RS = "\x1e";

const DEFAULT_ENDPOINT = "wss://substrate.office.com";

/** A turn fails only after this long with no frame at all (pings included). */
const IDLE_TIMEOUT_MS = 180000;

/** The upgrade is rejected without browser-shaped headers. */
const BROWSER_HEADERS = {
  Origin: "https://m365.cloud.microsoft",
  "User-Agent":
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:148.0) Gecko/20100101 Firefox/148.0",
};

/**
 * Feature flags copied from a captured first-party session.
 *
 * Cargo-culted on purpose: these are the values the real Office web client sends, and
 * which subset actually matters has never been established. Removing them is untested.
 */
const VARIANTS = [
  "feature.EnableWebSearch",
  "feature.EnableCodeInterpreter",
  "feature.EnableSydneyClientCompliance",
  "feature.EnableStreaming",
].join(",");

/** The message types we tell the server we are able to render. */
const ALLOWED_MESSAGE_TYPES = [
  "Chat",
  "Suggestion",
  "Progress",
  "EndOfRequest",
  "Disengaged",
  "InternalSearchQuery",
  "InternalSearchResult",
  "RenderCardRequest",
  "ReferencesListComplete",
  "GeneratedCode",
];







/** Read the object id and tenant id out of a Sydney token — they address the hub. */
export function decodeJwt(token        )            {
  const parts = token.split(".");
  if (parts.length < 2) throw new Error("Malformed JWT: expected three dot-separated segments");
  let claims     ;
  try {
    claims = JSON.parse(Buffer.from(parts[1] , "base64url").toString("utf8"));
  } catch {
    throw new Error("Malformed JWT: payload is not valid base64url JSON");
  }
  if (!claims || typeof claims !== "object") throw new Error("Malformed JWT: payload is not an object");
  return claims             ;
}

/** The message id a streaming cursor names: `$['<messageId>'].adaptiveCards[0].body[0].text`. */
export function cursorMessageId(path) {
  const match = typeof path === "string" ? /^\$\['([^']+)'\]/.exec(path) : null;
  return match ? match[1] : null;
}

/**
 * Rebuild one turn's answer from M365's per-message stream.
 *
 * Each bot message streams the same way: its FIRST token arrives only inside a full
 * snapshot (`messages[]`, with a `cursor` naming the message), then `writeAtCursor`
 * deltas extend whatever message the last cursor named (they carry no message id),
 * then a final snapshot. A turn can hold several messages — narration in one, a tool
 * fence in the next.
 *
 * Concatenating only the deltas therefore drops the head of every message
 * ("revert" arrives as snapshot "re" + delta "vert"), and comparing one global delta
 * string against one snapshot cannot recover it once a turn has two messages. So text
 * is kept per message and joined in order of first appearance.
 */
export class TurnText {
  messages = new Map();
  target = null;

  /** Subsequent deltas extend message `id`. */
  cursor(id) {
    this.target = id;
  }

  /** A full-text snapshot of one message; the longer of it and what we hold wins. */
  snapshot(id, text) {
    const key = id ?? this.target ?? "";
    if (this.target === null) this.target = key;
    const current = this.messages.get(key) ?? "";
    this.messages.set(key, text.length > current.length ? text : current);
  }

  /** A token delta for the message the cursor last named. */
  delta(text) {
    const key = this.target ?? "";
    this.messages.set(key, (this.messages.get(key) ?? "") + text);
  }

  get text() {
    return [...this.messages.values()].filter((text) => text.length > 0).join("\n\n");
  }
}

/** M365 refused to engage with the prompt. Not a rate limit — do not retry into it. */
export class DisengagedError extends Error {
           disengaged = true;
  constructor(message = "M365 Copilot disengaged from this conversation") {
    super(message);
    this.name = "DisengagedError";
  }
}









































/**
 * One M365 conversation.
 *
 * Stateful across turns: the conversation and session ids persist so the server keeps
 * its own context, and follow-up turns only need to carry new messages.
 */
export class CopilotSession {
  conversationId        ;
  sessionId        ;
  turnIndex = 0;

                   options                       ;

  constructor(options                       ) {
    this.options = options;
    this.conversationId = options.conversationId ?? randomUUID();
    this.sessionId = options.sessionId ?? randomUUID();
  }

  /** Abandon the server-side context and start a fresh conversation. */
  reset()       {
    this.conversationId = randomUUID();
    this.sessionId = randomUUID();
    this.turnIndex = 0;
  }

  async run(text        , turn              = {})                      {
    if (turn.signal?.aborted) throw abortError();

    const token = await this.options.getToken();
    const { oid, tid } = decodeJwt(token);
    const url = this.buildUrl(token, oid, tid);

    const { annotations, note } = await this.uploadImages(token, turn);
    const message                  = { text: note ? `${text}\n\n${note}` : text, annotations };

    if (turn.signal?.aborted) throw abortError();
    const socket = new WebSocket(url, { headers: BROWSER_HEADERS });
    try {
      const result = await this.pump(socket, message, turn);
      this.turnIndex += 1;
      return result;
    } finally {
      closeSocket(socket);
    }
  }

  /**
   * Upload the turn's images. A failure costs that image, not the turn: the model is
   * told what did not arrive, so it says so instead of describing a picture it never
   * received.
   */
          async uploadImages(
    token        ,
    turn             ,
  )                                                             {
    const images = turn.images ?? [];
    const annotations                    = [];
    const origin = (this.options.endpoint ?? DEFAULT_ENDPOINT).replace(/^ws/, "http");

    for (const image of images) {
      try {
        annotations.push(
          await uploadImage(image, {
            token,
            conversationId: this.conversationId,
            origin,
            ...(turn.signal ? { signal: turn.signal } : {}),
          }),
        );
      } catch (error) {
        if (turn.signal?.aborted) throw abortError();
        log.warn("image not attached", error instanceof Error ? error.message : String(error));
      }
    }

    const missing = images.length - annotations.length;
    if (missing === 0) return { annotations };
    const note =
      images.length === 1
        ? "(The user attached an image, but it could not be uploaded, so it is not visible to you. If it matters, let them know.)"
        : `(${missing} of the ${images.length} attached images could not be uploaded, so they are not visible to you. If they matter, let the user know.)`;
    return { annotations, note };
  }

          buildUrl(token        , oid        , tid        )         {
    const origin = this.options.endpoint ?? DEFAULT_ENDPOINT;
    const params = new URLSearchParams({
      access_token: token,
      ConversationId: this.conversationId,
      chatsessionid: this.sessionId,
      "X-SessionId": this.sessionId,
      clientrequestid: randomUUID(),
      source: '"officeweb"',
      product: "Office",
      agentHost: "Bizchat.FullScreen",
      scenario: "OfficeWebIncludedCopilot",
      variants: VARIANTS,
    });
    return `${origin}/m365Copilot/Chathub/${oid}@${tid}?${params.toString()}`;
  }

  /** Drive one turn to completion. */
          pump(socket           , outgoing                 , turn             )                      {
    return new Promise            ((resolve, reject) => {
      const text = new TurnText();
      // The completion item's `result.message`: a fallback when no update carried text.
      let finalMessage = "";
      // What has gone out through `onDelta`. Streamed bytes cannot be retracted, so
      // only a suffix that extends this exactly is ever emitted.
      let streamed = "";
      let handshakeDone = false;
      let settled = false;
      // An inactivity limit, not a turn limit: the server pings every few seconds, so
      // silence this long means a dead connection, while a long answer that is still
      // streaming is never cut off.
      let idle;
      const touch = () => {
        clearTimeout(idle);
        idle = setTimeout(() => finish(new Error(`M365 sent nothing for ${IDLE_TIMEOUT_MS / 1000} seconds`)), IDLE_TIMEOUT_MS);
        idle.unref();
      };
      touch();
      const result             = { text: "" };

      const advance = () => {
        if (!turn.onDelta) return;
        const now = text.text;
        if (now.length > streamed.length && now.startsWith(streamed)) {
          turn.onDelta(now.slice(streamed.length));
          streamed = now;
        }
      };

      const finish = (error        ) => {
        if (settled) return;
        settled = true;
        clearTimeout(idle);
        turn.signal?.removeEventListener("abort", onAbort);
        if (error) {
          reject(error);
          return;
        }
        const composed = text.text;
        result.text = composed.length >= finalMessage.length ? composed : finalMessage;
        result.streamed = streamed;
        resolve(result);
      };

      const onAbort = () => {
        // The real client's Stop button: an invocation on the same socket, with a
        // different invocationId from the chat frame, then close.
        try {
          socket.send(`${JSON.stringify({ arguments: [{}], invocationId: "1", target: "stop", type: 1 })}${RS}`);
        } catch {
          /* the socket may already be gone */
        }
        finish(abortError());
      };
      turn.signal?.addEventListener("abort", onAbort, { once: true });

      socket.on("error", (error) => finish(error instanceof Error ? error : new Error(String(error))));
      // Only a completion frame (type 3, or a clean type 7) ends a turn. A close
      // before one is a dropped transport: resolving with whatever arrived would hand
      // back a truncated — or empty — answer that looks like a dead agent.
      socket.on("close", (code, reason) => {
        const why = reason.toString();
        finish(new Error(`M365 connection closed before the turn completed (code ${code}${why ? `: ${why}` : ""})`));
      });

      socket.on("open", () => {
        socket.send(`${JSON.stringify({ protocol: "json", version: 1 })}${RS}`);
      });

      socket.on("message", (raw) => {
        touch();
        for (const frame of splitFrames(raw.toString())) {
          // The handshake ack is `{}`; the chat frame goes out immediately after.
          if (!handshakeDone) {
            handshakeDone = true;
            this.sendChat(socket, outgoing, turn);
            // An ack carries nothing else, but a server that batches may.
            if (Object.keys(frame).length === 0) continue;
          }

          switch (frame.type) {
            case 1:
              if (frame.target === "update") {
                for (const argument of frame.arguments ?? []) {
                  if (this.applyUpdate(argument, result, text)) {
                    finish(new DisengagedError());
                    return;
                  }
                }
                advance();
              }
              break;

            case 2:
              this.applyStreamItem(frame.item, result);
              if (typeof frame.item?.result?.message === "string" && frame.item.result.message.length > finalMessage.length) {
                finalMessage = frame.item.result.message;
              }
              break;

            case 3:
              if (frame.error) {
                finish(new Error(`M365 rejected the turn: ${frame.error}`));
                return;
              }
              finish();
              return;

            case 6:
              // Keepalive. Silence here gets the connection dropped mid-answer.
              socket.send(`${JSON.stringify({ type: 6 })}${RS}`);
              break;

            case 7:
              if (frame.error) {
                finish(new Error(`M365 closed the connection: ${frame.error}`));
                return;
              }
              finish();
              return;

            default:
              break;
          }
        }
      });
    });
  }

  /** The chat invocation and its mandatory Metrics companion, in one send. */
          sendChat(socket           , outgoing                 , turn             )       {
    const now = new Date().toISOString();
    const { text } = outgoing;

    const message                          = {
      text,
      author: "user",
      inputMethod: "Keyboard",
      messageType: "Chat",
      locale: "en-US",
      timestamp: now,
      messageId: randomUUID(),
    };
    if (outgoing.annotations.length > 0) message.messageAnnotations = outgoing.annotations;

    const argument                          = {
      message,
      tone: turn.tone ?? "magic",
      source: "officeweb",
      streamingMode: "ConciseWithPadding",
      isStartOfSession: this.turnIndex === 0,
      conversationId: this.conversationId,
      allowedMessageTypes: ALLOWED_MESSAGE_TYPES,
      optionsSets: [],
      clientInfo: {
        clientPlatform: "mcmcopilot-web",
        clientAppName: "Office",
// eslint-disable-next-line @typescript-eslint/naming-convention
        clientAppVersion: "1.0.0",
      },
      traceId: randomUUID(),
    };

    if (turn.agentId) {
      // With an agent attached, the declarative-agent fields replace `plugins`.
      argument.threadLevelGptId = { id: turn.agentId, source: "MOS3" };
      argument.gpts = [
        {
          id: turn.agentId,
          source: "MOS3",
          version: "1.0.0",
          clientOverrides: { capabilities: [], "deepResearchModels@odata.type": "Collection(String)" },
        },
      ];
    } else {
      argument.plugins = [{ Id: "BingWebSearch", Source: "BuiltIn" }];
    }

    const chat = { arguments: [argument], invocationId: "0", target: "chat", type: 4 };
    const metrics = {
      arguments: [
        {
          Timestamps: {
            ConnectionStart: now,
            UserInputStart: now,
            ConnectionEstablished: now,
            UserInputSubmit: now,
          },
        },
      ],
      target: "Metrics",
      type: 1,
    };

    log.info(
      "sending turn",
      trunc({ tone: argument.tone, agent: Boolean(turn.agentId), images: outgoing.annotations.length, text }),
    );
    // Both frames, one send. This is the part that silently breaks everything.
    socket.send(`${JSON.stringify(chat)}${RS}${JSON.stringify(metrics)}${RS}`);
  }

  /** Fold one `update` argument into the turn's text. Returns true on Disengaged. */
          applyUpdate(
    argument     ,
    result            ,
    text          ,
  )           {
    // The cursor first: it names the message this argument's text belongs to.
    const cursor = cursorMessageId(argument?.cursor?.j);
    if (cursor) text.cursor(cursor);

    if (typeof argument?.writeAtCursor === "string") text.delta(argument.writeAtCursor);

    if (argument?.throttling) {
      result.throttle = {
        current: argument.throttling.numUserMessagesInConversation,
        max: argument.throttling.maxNumUserMessagesInConversation,
      };
    }

    for (const message of argument?.messages ?? []) {
      if (message?.author !== "bot") continue;
      if (message.messageType === "Disengaged") return true;
      // A messageType means control traffic, never the answer.
      if (message.messageType) continue;
      if (typeof message.text === "string") text.snapshot(message.messageId, message.text);
      absorbMetadata(message, result);
    }

    return false;
  }

          applyStreamItem(item     , result            )       {
    if (!item) return;
    if (item.throttling) {
      result.throttle = {
        current: item.throttling.numUserMessagesInConversation,
        max: item.throttling.maxNumUserMessagesInConversation,
      };
    }
    for (const message of item.messages ?? []) {
      if (message?.author === "bot" && !message.messageType) absorbMetadata(message, result);
    }
  }
}

/** What the chat frame carries from the caller: the prompt and its uploaded images. */





/** Pull the diagnostic fields off a final bot message. */
function absorbMetadata(message     , result            )       {
  if (typeof message.contentOrigin === "string") result.contentOrigin = message.contentOrigin;
  if (typeof message.turnCount === "number") result.turnCount = message.turnCount;
  if (Array.isArray(message.scores)) {
    result.scores ??= {};
    for (const score of message.scores) {
      if (typeof score?.component === "string") result.scores[score.component] = score.score;
    }
  }
}

/** One WebSocket message can carry several `0x1E`-terminated frames. */
function splitFrames(raw        )        {
  const frames        = [];
  for (const chunk of raw.split(RS)) {
    if (chunk.trim() === "") continue;
    try {
      frames.push(JSON.parse(chunk));
    } catch {
      log.warn("unparseable frame", trunc(chunk));
    }
  }
  return frames;
}

/**
 * Shut a turn's socket down without losing anything still buffered.
 *
 * `terminate()` destroys the socket immediately and drops queued frames — which is
 * how a pong or the stop frame can be sent and still never arrive. `close()` flushes
 * first; the timer is the backstop for a server that never completes the handshake.
 */
function closeSocket(socket           )       {
  if (socket.readyState === WebSocket.CLOSED || socket.readyState === WebSocket.CLOSING) return;
  try {
    socket.close();
  } catch {
    socket.terminate();
    return;
  }
  const timer = setTimeout(() => socket.terminate(), 1000);
  timer.unref?.();
  socket.once("close", () => clearTimeout(timer));
}

function abortError()        {
  const error = new Error("The M365 turn was aborted");
  error.name = "AbortError";
  return error;
}
