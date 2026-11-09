// Never log prompts, images, credentials or upstream URLs.
export function trunc() { return '[payload omitted]'; }
export function createLogger(scope) {
  const write = level => (...parts) => {
    if (process.env.M365_DEBUG === '1') console.error(`[${scope}] ${level}: ${String(parts[0])}`);
  };
  return { info: write('info'), warn: write('warn'), error: write('error') };
}
