/**
 * Test doubles for everything a command touches outside the process:
 * the Graph HTTP API (globalThis.fetch), inquirer prompts, process.exit and
 * terminal output.
 */
import inquirer from 'inquirer';

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

function jsonResponse(status, body) {
  if (status === 204 || body === undefined) return new Response(null, { status });
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

/**
 * Replace globalThis.fetch with `handler`. Every request is recorded in the
 * returned `calls` array as { method, url, path, body }.
 *
 * `handler(call)` returns { status?, body? } (status defaults to 200), or
 * nothing for a 404. Avoid 429/503/504: the Graph client retries those with
 * real delays.
 *
 * @param {object} t - ava execution context (restores fetch on teardown)
 */
export function mockGraph(t, handler) {
  const original = globalThis.fetch;
  const calls = [];

  globalThis.fetch = async (input, init = {}) => {
    const url = decodeURIComponent(typeof input === 'string' ? input : input.url);
    const call = {
      method: (init.method || 'GET').toUpperCase(),
      url,
      path: url.startsWith(GRAPH_ROOT) ? url.slice(GRAPH_ROOT.length) : url,
      body: typeof init.body === 'string' ? JSON.parse(init.body) : null,
    };
    calls.push(call);
    const { status = 200, body } = (await handler(call)) ?? {
      status: 404,
      body: { error: { code: 'Request_ResourceNotFound', message: `not mocked: ${call.method} ${call.path}` } },
    };
    return jsonResponse(status, body);
  };

  t.teardown(() => { globalThis.fetch = original; });
  return calls;
}

/** Build the response for a Graph $batch POST from a per-request resolver. */
export function batchResponse(call, resolve) {
  return {
    body: {
      responses: call.body.requests.map((r) => ({ id: r.id, status: 200, ...resolve(r) })),
    },
  };
}

/** Requests that change tenant state. */
export function writes(calls) {
  return calls.filter((c) => c.method !== 'GET' && !c.path.endsWith('/$batch'));
}

/** Answer inquirer prompts from a queue of answer objects, one per prompt() call. */
export function mockPrompts(t, answers) {
  const original = inquirer.prompt;
  const queue = [...answers];
  inquirer.prompt = async () => {
    if (queue.length === 0) throw new Error('Unexpected inquirer prompt');
    return queue.shift();
  };
  t.teardown(() => { inquirer.prompt = original; });
}

class ExitSignal extends Error {
  constructor(code) {
    super(`process.exit(${code})`);
    this.code = code;
  }
}

/**
 * Run a command with terminal output captured and process.exit intercepted.
 * @returns {Promise<{ exitCode: number|null, output: string }>} exitCode is
 *   null when the command returned without calling process.exit.
 */
export async function runCommand(fn) {
  const original = {
    exit: process.exit,
    out: process.stdout.write,
    err: process.stderr.write,
  };
  let output = '';
  const capture = (chunk) => { output += String(chunk); return true; };

  process.exit = (code = 0) => { throw new ExitSignal(code); };
  process.stdout.write = capture;
  process.stderr.write = capture;

  let exitCode = null;
  try {
    await fn();
  } catch (err) {
    if (!(err instanceof ExitSignal)) throw err;
    exitCode = err.code;
  } finally {
    process.exit = original.exit;
    process.stdout.write = original.out;
    process.stderr.write = original.err;
  }
  // Strip ANSI colour codes so assertions can match plain text
  return { exitCode, output: output.replace(/\x1B\[[\x20-\x3F]*[\x40-\x7E]/g, '') };
}
