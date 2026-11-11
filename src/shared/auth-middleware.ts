import { createHash, timingSafeEqual } from "node:crypto";
import type { Request, Response, NextFunction, RequestHandler } from "express";

// Tokens korter dan dit krijgen een waarschuwing in de log. Niet weigeren:
// de relay en andere clients gebruiken de bestaande token.
export const MIN_TOKEN_BYTES = 32;

/**
 * Fail-closed: zonder (of met lege) MCP_AUTH_TOKEN start de server niet.
 * Een korte token levert alleen een waarschuwing op.
 */
export function requireAuthToken(token: string | undefined): string {
  if (!token) {
    throw new Error(
      "MCP_AUTH_TOKEN ontbreekt of is leeg; de server start niet zonder token",
    );
  }
  if (Buffer.byteLength(token, "utf8") < MIN_TOKEN_BYTES) {
    console.warn(
      `WAARSCHUWING: MCP_AUTH_TOKEN is korter dan ${MIN_TOKEN_BYTES} bytes; vervang hem door een lange willekeurige token`,
    );
  }
  return token;
}

function sha256(value: string): Buffer {
  return createHash("sha256").update(value, "utf8").digest();
}

// Timing-safe vergelijking: beide kanten eerst hashen, zodat de lengte gelijk is.
export function createAuthMiddleware(authToken: string | undefined): RequestHandler {
  const expectedHash = sha256(requireAuthToken(authToken));
  return (req: Request, res: Response, next: NextFunction): void => {
    const header = req.headers.authorization;
    const match = typeof header === "string" ? /^Bearer (.+)$/i.exec(header) : null;
    if (!match || !timingSafeEqual(sha256(match[1]), expectedHash)) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
    next();
  };
}
