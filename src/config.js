/**
 * opencode configuration: the provider entry, and the lean tool profile that makes
 * M365 willing to engage at all.
 *
 * Everything here is pure — it takes a base URL and returns/mutates plain objects —
 * so the interesting decisions are testable without opencode, a network, or a tenant.
 */

import { MODELS, isLocalModel } from "./models.js";
import { findShellTool,              } from "./fenced.js";

/** The provider id opencode addresses models by: `m365/<model>`. */
export const PROVIDER_ID = "m365";

/**
 * The plugin's own id.
 *
 * opencode 2 requires one, scopes the plugin's storage by it, and names the plugin
 * by it in `plugin list` and diagnostics. Changing it is not cosmetic: it orphans
 * anything stored under the old one.
 */
export const PLUGIN_ID = "m365-copilot";

/**
 * The measured point at which M365's Disengaged filter starts firing.
 *
 * From the protocol notes: ~1 tool is fine, ~12 is borderline (disengages once,
 * recovers on retry), and a full coding-agent toolset (~15+, opencode's is named
 * explicitly) disengages *persistently*. We stay comfortably below it.
 *
 * Note the trigger is prompt *shape*, not byte count — 500k tokens of benign filler
 * never disengaged. It is the number of tool blocks that matters.
 */
export const DISENGAGE_TOOL_BUDGET = 12;

/**
 * The tools we leave enabled.
 *
 * `bash` is not optional. M365's chat-tuned model refuses to "act as an agent" on
 * demand, but reflexively writes a ```bash block — routing that block to the shell
 * tool is the single lever that took the reference implementation's benchmark from
 * 0/5 to real multi-turn loops. Remove bash and tool calling largely stops working.
 *
 * The rest is the minimum needed to read and change code.
 */
export const LEAN_TOOLS = ["bash", "read", "edit", "write", "apply_patch", "grep"]         ;

/**
 * How many tools to keep when we recognise none of the harness's names.
 *
 * This is the path a future opencode rename drops us on — `apply_patch` replaced
 * `edit`/`write` once already — so it must not sit at the edge of the threshold.
 * ~12 is where the filter "disengages once and recovers on retry"; landing there by
 * accident would look like flakiness rather than a drifted allowlist.
 */
export const LEAN_FALLBACK_LIMIT = 5;

/**
 * Cut a harness's toolset down to something M365 will engage with.
 *
 * This is the enforcement point. opencode 1's own `tools` config was resolved into the
 * config but **not applied to the request** (verified against 1.18.18: a config
 * disabling eight tools still produced a request offering all nine). opencode 2's
 * `tool.transform` has not been measured against a live tenant. The proxy sees the
 * final request whatever the harness does, so the trim lives here.
 *
 * Two rules, in order:
 *   1. Keep the tools we recognise as load-bearing for a coding loop.
 *   2. If we recognise none of them — a harness with its own vocabulary — fall back
 *      to a hard cap, keeping any shell tool first.
 */
export function selectLeanTools(tools                    )            {
  if (tools.length === 0) return [...tools];

  const wanted = new Set        (LEAN_TOOLS);
  const kept = tools.filter((tool) => wanted.has(tool.function.name.toLowerCase()));
  if (kept.length > 0) return kept;

  // Unknown vocabulary. Keep the shell tool if we can spot one, then fill up to the
  // cap in the order the harness offered them.
  const shell = findShellTool(tools);
  const rest = tools.filter((tool) => tool !== shell);
  const remaining = shell ? LEAN_FALLBACK_LIMIT - 1 : LEAN_FALLBACK_LIMIT;
  return [...(shell ? [shell] : []), ...rest.slice(0, remaining)];
}

/** Tool names that mean "change a file", across opencode versions and harnesses. */
const EDITING_TOOLS = ["edit", "write", "apply_patch", "patch", "str_replace", "create"];

/**
 * Describe anything worrying about the trimmed toolset, or undefined if it is fine.
 *
 * The allowlist in `LEAN_TOOLS` is coupled to what the harness calls its tools, and
 * that has already changed once. When the trim silently removes the ability to edit
 * or to run a command, the model degrades quietly — worth a line in the log rather
 * than a mystery.
 */
export function describeToolSelection(requested                    )                     {
  if (requested.length === 0) return undefined;
  const kept = selectLeanTools(requested);
  const names = kept.map((tool) => tool.function.name.toLowerCase());

  if (!findShellTool(kept)) {
    return "the trimmed toolset has no shell tool, so shell-routing is unavailable and tool calling will be markedly less reliable";
  }
  if (!names.some((name) => EDITING_TOOLS.includes(name))) {
    return "the trimmed toolset has no editing tool — the model can only change files through the shell. The harness may have renamed its tools; check LEAN_TOOLS.";
  }
  return undefined;
}
