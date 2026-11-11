import {
  createCipheriv,
  createDecipheriv,
  hkdfSync,
  randomBytes,
} from "node:crypto";
import { chmod, mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { dirname } from "node:path";

/**
 * Opslag van de MSAL-tokencache (met refresh-token) op het volume.
 *
 * - Met MS365_CACHE_KEY: AES-256-GCM, sleutel via HKDF uit de env-waarde.
 *   Een bestaande onversleutelde cache wordt bij het opstarten ingelezen en
 *   meteen versleuteld teruggeschreven (migratie zonder opnieuw inloggen).
 * - Zonder MS365_CACHE_KEY: onversleuteld zoals voorheen, met een waarschuwing.
 * Altijd: bestand 0600, atomisch geschreven (tijdelijk bestand + rename).
 */

const FORMAT = "mcp-ms365-token-cache";
const ALG = "aes-256-gcm";

interface Envelope {
  format: typeof FORMAT;
  v: 1;
  alg: typeof ALG;
  iv: string;
  tag: string;
  data: string;
}

export function deriveKey(secret: string): Buffer {
  return Buffer.from(hkdfSync("sha256", secret, "mcp-ms365", "token-cache v1", 32));
}

export function isEncrypted(text: string): boolean {
  try {
    const parsed = JSON.parse(text) as Partial<Envelope>;
    return parsed?.format === FORMAT;
  } catch {
    return false;
  }
}

export function encryptCache(plain: string, key: Buffer): string {
  const iv = randomBytes(12);
  const cipher = createCipheriv(ALG, key, iv);
  const data = Buffer.concat([cipher.update(plain, "utf8"), cipher.final()]);
  const envelope: Envelope = {
    format: FORMAT,
    v: 1,
    alg: ALG,
    iv: iv.toString("base64"),
    tag: cipher.getAuthTag().toString("base64"),
    data: data.toString("base64"),
  };
  return JSON.stringify(envelope);
}

export function decryptCache(text: string, key: Buffer): string {
  const env = JSON.parse(text) as Envelope;
  if (env.format !== FORMAT || env.v !== 1 || env.alg !== ALG) {
    throw new Error("Onbekend formaat tokencache");
  }
  const decipher = createDecipheriv(ALG, key, Buffer.from(env.iv, "base64"));
  decipher.setAuthTag(Buffer.from(env.tag, "base64"));
  return Buffer.concat([
    decipher.update(Buffer.from(env.data, "base64")),
    decipher.final(),
  ]).toString("utf8");
}

export interface LoadResult {
  /** De MSAL-cache als JSON-tekst, of null als er niets bruikbaars is. */
  content: string | null;
  /** True als de cache onversleuteld op schijf stond en er een sleutel is. */
  needsMigration: boolean;
}

export class TokenCacheStore {
  private readonly key: Buffer | null;

  constructor(private readonly path: string, secret: string) {
    this.key = secret ? deriveKey(secret) : null;
  }

  get encrypted(): boolean {
    return this.key !== null;
  }

  async load(): Promise<LoadResult> {
    let text: string;
    try {
      text = await readFile(this.path, "utf-8");
    } catch {
      return { content: null, needsMigration: false }; // eerste start
    }

    if (isEncrypted(text)) {
      if (!this.key) {
        console.error(
          "Tokencache is versleuteld, maar MS365_CACHE_KEY ontbreekt: cache niet geladen. Zet de sleutel terug of log opnieuw in.",
        );
        return { content: null, needsMigration: false };
      }
      try {
        return { content: decryptCache(text, this.key), needsMigration: false };
      } catch {
        console.error(
          "Tokencache kon niet worden ontsleuteld (verkeerde MS365_CACHE_KEY?): cache niet geladen.",
        );
        return { content: null, needsMigration: false };
      }
    }

    if (!this.key) {
      console.warn(
        "WAARSCHUWING: MS365_CACHE_KEY niet gezet: de tokencache (met refresh-token) staat onversleuteld op het volume.",
      );
    }
    return { content: text, needsMigration: this.key !== null };
  }

  async save(plain: string): Promise<void> {
    const out = this.key ? encryptCache(plain, this.key) : plain;
    if (this.key && decryptCache(out, this.key) !== plain) {
      throw new Error("Controle van de versleutelde tokencache mislukt; niet weggeschreven");
    }
    await mkdir(dirname(this.path), { recursive: true });
    const tmp = `${this.path}.${process.pid}.tmp`;
    await writeFile(tmp, out, { encoding: "utf-8", mode: 0o600 });
    await rename(tmp, this.path);
    await chmod(this.path, 0o600);
  }
}
