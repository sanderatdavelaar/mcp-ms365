import type { Request, Response, NextFunction } from "express";

export function createAuthMiddleware(authToken: string) {
  return (req: Request, res: Response, next: NextFunction): void => {
    if (!authToken) {
      next();
      return;
    }
    const authHeader = req.headers.authorization;
    if (
      !authHeader ||
      !authHeader.startsWith("Bearer ") ||
      authHeader.slice(7) !== authToken
    ) {
      res.status(401).json({ error: "Unauthorized" });
      return;
    }
    next();
  };
}
