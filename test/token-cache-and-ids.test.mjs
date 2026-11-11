// Tests voor de versleutelde tokencache met migratie (F6-4) en Graph-ID-encoding.
import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, readFileSync, writeFileSync, statSync, readdirSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { TokenCacheStore, isEncrypted } from "../dist/auth/token-cache.js";
import { gid } from "../dist/shared/graph-id.js";

const SAMPLE = JSON.stringify({
  Account: { a: { username: "test@example.org" } },
  RefreshToken: { r: { secret: "refresh-geheim-123" } },
});
const KEY = "sleutel-0123456789abcdef0123456789abcdef";

const tmpFile = () => join(mkdtempSync(join(tmpdir(), "cache-")), "token-cache.json");

test("zonder sleutel: onversleuteld zoals voorheen, bestand 0600", async () => {
  const path = tmpFile();
  const store = new TokenCacheStore(path, "");
  await store.save(SAMPLE);
  assert.equal(readFileSync(path, "utf8"), SAMPLE);
  assert.equal(statSync(path).mode & 0o777, 0o600);
  const { content, needsMigration } = await store.load();
  assert.equal(content, SAMPLE);
  assert.equal(needsMigration, false);
});

test("met sleutel: refresh-token niet leesbaar op schijf, wel terug te lezen", async () => {
  const path = tmpFile();
  const store = new TokenCacheStore(path, KEY);
  await store.save(SAMPLE);
  const raw = readFileSync(path, "utf8");
  assert.ok(isEncrypted(raw));
  assert.ok(!raw.includes("refresh-geheim-123"));
  assert.equal(statSync(path).mode & 0o777, 0o600);
  const { content } = await new TokenCacheStore(path, KEY).load();
  assert.equal(content, SAMPLE);
  // geen achtergebleven tijdelijke bestanden
  assert.deepEqual(readdirSync(join(path, "..")), ["token-cache.json"]);
});

test("migratie: bestaande onversleutelde cache wordt geladen en daarna versleuteld", async () => {
  const path = tmpFile();
  writeFileSync(path, SAMPLE, { mode: 0o644 });
  const store = new TokenCacheStore(path, KEY);
  const { content, needsMigration } = await store.load();
  assert.equal(content, SAMPLE);
  assert.equal(needsMigration, true);
  await store.save(content);
  assert.ok(isEncrypted(readFileSync(path, "utf8")));
  const again = await new TokenCacheStore(path, KEY).load();
  assert.equal(again.content, SAMPLE);
  assert.equal(again.needsMigration, false);
});

test("verkeerde of ontbrekende sleutel: niets geladen, bestand niet aangeraakt", async () => {
  const path = tmpFile();
  await new TokenCacheStore(path, KEY).save(SAMPLE);
  const before = readFileSync(path, "utf8");
  assert.equal((await new TokenCacheStore(path, "andere-sleutel-xxxxxxxxxxxxxxxxxxxxxxxx").load()).content, null);
  assert.equal((await new TokenCacheStore(path, "").load()).content, null);
  assert.equal(readFileSync(path, "utf8"), before);
});

test("geen cachebestand: leeg resultaat", async () => {
  const { content, needsMigration } = await new TokenCacheStore(tmpFile(), KEY).load();
  assert.equal(content, null);
  assert.equal(needsMigration, false);
});

test("gid: ID's worden één padsegment", () => {
  assert.equal(gid("AAMkAGI2THVSAAA="), "AAMkAGI2THVSAAA%3D");
  assert.equal(gid("19:abc@thread.v2"), "19%3Aabc%40thread.v2");
  assert.equal(gid("../messages/123"), "..%2Fmessages%2F123");
  assert.equal(gid("x?$select=a#b"), "x%3F%24select%3Da%23b");
  assert.throws(() => gid(""));
  assert.throws(() => gid(".."));
  assert.throws(() => gid("."));
});
