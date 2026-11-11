// Tests voor fail-closed auth (F6-1) en stateless server per request (F6-2).
// Draaien tegen de gecompileerde code: `npm test` (tsc && node --test test/).
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createApp } from "../dist/app.js";
import { createMs365Server as createServer } from "../dist/server.js";

// Env die de server nodig heeft naast MCP_AUTH_TOKEN (dummywaarden).
const EXTRA_ENV = {
  MS365_CLIENT_ID: "00000000-0000-0000-0000-000000000001",
  MS365_TENANT_ID: "00000000-0000-0000-0000-000000000002",
  DATA_DIR: mkdtempSync(join(tmpdir(), "ms365-test-")),
};
const TOKEN = "test-token-0123456789abcdef0123456789abcdef";

let server;
let base;

before(async () => {
  const app = createApp({ authToken: TOKEN, createServer });
  await new Promise((resolve) => {
    server = app.listen(0, "127.0.0.1", resolve);
  });
  base = `http://127.0.0.1:${server.address().port}`;
});

after(() => new Promise((resolve) => server.close(resolve)));

function rpc(body, token = TOKEN) {
  const headers = {
    "Content-Type": "application/json",
    Accept: "application/json, text/event-stream",
  };
  if (token !== null) headers.Authorization = `Bearer ${token}`;
  return fetch(`${base}/mcp`, {
    method: "POST",
    headers,
    body: typeof body === "string" ? body : JSON.stringify(body),
  });
}

const initialize = (id) => ({
  jsonrpc: "2.0",
  id,
  method: "initialize",
  params: {
    protocolVersion: "2025-06-18",
    capabilities: {},
    clientInfo: { name: "test", version: "1.0.0" },
  },
});
const toolsList = (id) => ({ jsonrpc: "2.0", id, method: "tools/list", params: {} });

test("health is open", async () => {
  const res = await fetch(`${base}/health`);
  assert.equal(res.status, 200);
});

test("zonder token: 401", async () => {
  const res = await rpc(initialize(1), null);
  assert.equal(res.status, 401);
});

test("fout token: 401", async () => {
  const res = await rpc(initialize(1), "fout");
  assert.equal(res.status, 401);
  const res2 = await rpc(initialize(1), TOKEN + "x");
  assert.equal(res2.status, 401);
});

test("geen Bearer-schema: 401", async () => {
  const res = await fetch(`${base}/mcp`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: TOKEN },
    body: JSON.stringify(initialize(1)),
  });
  assert.equal(res.status, 401);
});

test("body wordt niet geparsed vóór de auth", async () => {
  // Kapotte JSON zonder token: 401 (auth eerst), niet 400 (parser eerst).
  const res = await rpc("{kapot", null);
  assert.equal(res.status, 401);
});

test("twee opeenvolgende sessies werken", async () => {
  for (let i = 1; i <= 2; i++) {
    const init = await rpc(initialize(i));
    assert.equal(init.status, 200);
    const initBody = await init.json();
    assert.ok(initBody.result?.serverInfo, "initialize geeft serverInfo");
    const list = await rpc(toolsList(i + 10));
    assert.equal(list.status, 200);
    const listBody = await list.json();
    assert.ok(listBody.result.tools.length > 0, "tools/list geeft tools");
  }
});

test("gelijktijdige requests: geen 'Already connected'", async () => {
  // Server met een trage tool, zodat de requests echt tegelijk lopen.
  const slowServer = () => {
    const s = new McpServer({ name: "slow", version: "1.0.0" });
    s.tool("slow", "wacht even", async () => {
      await new Promise((r) => setTimeout(r, 200));
      return { content: [{ type: "text", text: "klaar" }] };
    });
    return s;
  };
  const app = createApp({ authToken: TOKEN, createServer: slowServer });
  const srv = await new Promise((resolve) => {
    const s = app.listen(0, "127.0.0.1", () => resolve(s));
  });
  try {
    const url = `http://127.0.0.1:${srv.address().port}/mcp`;
    const call = (id) =>
      fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Accept: "application/json, text/event-stream",
          Authorization: `Bearer ${TOKEN}`,
        },
        body: JSON.stringify({
          jsonrpc: "2.0",
          id,
          method: "tools/call",
          params: { name: "slow", arguments: {} },
        }),
        signal: AbortSignal.timeout(5000),
      }).then((r) => r.json());
    const results = await Promise.all(Array.from({ length: 5 }, (_, i) => call(100 + i)));
    for (const [i, body] of results.entries()) {
      assert.equal(body.id, 100 + i);
      assert.equal(body.result?.content?.[0]?.text, "klaar");
    }
  } finally {
    await new Promise((resolve) => srv.close(resolve));
  }
});

function runIndex(env) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, ["dist/index.js"], {
      env: { PATH: process.env.PATH, PORT: "0", ...env },
    });
    let output = "";
    const onData = (d) => {
      output += d;
      // Server luistert: genoeg gezien, stoppen.
      if (/running/i.test(output)) child.kill("SIGTERM");
    };
    child.stdout.on("data", onData);
    child.stderr.on("data", onData);
    const timer = setTimeout(() => child.kill("SIGTERM"), 30000);
    child.on("exit", (code, signal) => {
      clearTimeout(timer);
      resolve({ code, signal, output });
    });
  });
}

test("start geweigerd zonder MCP_AUTH_TOKEN", async () => {
  const { code, output } = await runIndex({ ...EXTRA_ENV });
  assert.equal(code, 1);
  assert.match(output, /MCP_AUTH_TOKEN/);
});

test("start geweigerd met lege MCP_AUTH_TOKEN", async () => {
  const { code, output } = await runIndex({ ...EXTRA_ENV, MCP_AUTH_TOKEN: "" });
  assert.equal(code, 1);
  assert.match(output, /MCP_AUTH_TOKEN/);
});

test("korte token: start wel, met waarschuwing", async () => {
  const { code, signal, output } = await runIndex({ ...EXTRA_ENV, MCP_AUTH_TOKEN: "kort" });
  assert.equal(code, null);
  assert.equal(signal, "SIGTERM");
  assert.match(output, /WAARSCHUWING: MCP_AUTH_TOKEN is korter/);
});
