import { getAccessToken } from "./msal.js";

const GRAPH_ORIGIN = "https://graph.microsoft.com";
const GRAPH_BASE = `${GRAPH_ORIGIN}/v1.0`;
const MAX_RETRIES = 3;
const MAX_PAGES = 10;

interface GraphErrorBody {
  error?: {
    code?: string;
    message?: string;
  };
}

interface GraphPagedResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.count"?: number;
}

/**
 * Graph-fout met de HTTP-status erbij, zodat een aanroeper op bv. een 400 kan
 * terugvallen op een eenvoudiger variant. De melding is dezelfde als voorheen.
 */
export class GraphError extends Error {
  constructor(
    public readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

/**
 * Een `@odata.nextLink` is een volledige URL; die mag alleen naar Graph zelf
 * wijzen (het token gaat mee). Al het andere is een pad onder /v1.0.
 */
function graphUrl(endpoint: string): string {
  if (endpoint.startsWith(`${GRAPH_ORIGIN}/`)) return endpoint;
  if (/^[a-z][a-z0-9+.-]*:/i.test(endpoint)) {
    throw new Error("Ongeldige Graph-URL");
  }
  return `${GRAPH_BASE}${endpoint}`;
}

async function parseErrorBody(response: Response): Promise<GraphErrorBody> {
  try {
    return (await response.json()) as GraphErrorBody;
  } catch {
    return {};
  }
}

function getErrorMessage(errorBody: GraphErrorBody, fallback: string): string {
  return errorBody.error?.message ?? fallback;
}

async function graphRequest<T>(
  method: string,
  endpoint: string,
  body?: unknown,
  options?: { allPages?: boolean; headers?: Record<string, string> },
): Promise<T> {
  const token = await getAccessToken();

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
    ...options?.headers,
  };

  const fetchOptions: RequestInit = {
    method,
    headers,
  };

  if (body !== undefined) {
    fetchOptions.body = JSON.stringify(body);
  }

  const url = graphUrl(endpoint);
  let retries = 0;

  // eslint-disable-next-line no-constant-condition
  while (true) {
    const response = await fetch(url, fetchOptions);

    if (response.status === 429) {
      if (retries >= MAX_RETRIES) {
        throw new Error(`Rate limited after ${MAX_RETRIES} retries: ${endpoint}`);
      }
      const retryAfter = response.headers.get("Retry-After");
      const waitSeconds = retryAfter
        ? parseInt(retryAfter, 10)
        : Math.pow(2, retries); // 1s, 2s, 4s
      await new Promise((resolve) => setTimeout(resolve, waitSeconds * 1000));
      retries++;
      continue;
    }

    // Korte storing aan de kant van Microsoft (bad gateway, niet beschikbaar): bij lezen één keer opnieuw proberen.
    // Een 504 niet: daar heeft Graph al een halve minuut op gewacht, en nog eens wachten laat de aanroeper verlopen.
    if (method === "GET" && [502, 503].includes(response.status) && retries < 1) {
      await new Promise((resolve) => setTimeout(resolve, 2000));
      retries++;
      continue;
    }

    if (response.status === 401) {
      throw new Error("Session expired. Please re-authenticate with ms365_login.");
    }

    if (response.status === 403) {
      const errorBody = await parseErrorBody(response);
      throw new Error(
        `Insufficient permissions: ${getErrorMessage(errorBody, "Access denied")}`,
      );
    }

    if (response.status === 404) {
      throw new Error(`Not found: ${endpoint}`);
    }

    if (!response.ok) {
      const errorBody = await parseErrorBody(response);
      throw new GraphError(
        response.status,
        `Graph API error ${response.status}: ${getErrorMessage(errorBody, response.statusText)}`,
      );
    }

    // 202 Accepted / 204 No Content — no body to parse
    if (response.status === 202 || response.status === 204) {
      return undefined as T;
    }

    const json = await response.json();

    // Handle pagination when allPages is requested
    if (options?.allPages && isPagedResponse(json)) {
      return await fetchAllPages<T>(json, headers);
    }

    return json as T;
  }
}

function isPagedResponse(json: unknown): json is GraphPagedResponse<unknown> {
  return (
    typeof json === "object" &&
    json !== null &&
    "value" in json &&
    Array.isArray((json as GraphPagedResponse<unknown>).value)
  );
}

async function fetchAllPages<T>(
  firstPage: GraphPagedResponse<unknown>,
  headers: Record<string, string>,
): Promise<T> {
  const allValues = [...firstPage.value];
  let nextLink = firstPage["@odata.nextLink"];
  let pageCount = 1;

  while (nextLink && pageCount < MAX_PAGES) {
    const response = await fetch(nextLink, { method: "GET", headers });

    if (!response.ok) {
      const errorBody = await parseErrorBody(response);
      throw new Error(
        `Graph API error ${response.status} during pagination: ${getErrorMessage(errorBody, response.statusText)}`,
      );
    }

    const json = (await response.json()) as GraphPagedResponse<unknown>;
    allValues.push(...json.value);
    nextLink = json["@odata.nextLink"];
    pageCount++;
  }

  return {
    ...firstPage,
    value: allValues,
    "@odata.nextLink": undefined,
  } as T;
}

export async function graphGet<T>(
  endpoint: string,
  allPages?: boolean,
  headers?: Record<string, string>,
): Promise<T> {
  return graphRequest<T>("GET", endpoint, undefined, { allPages, headers });
}

export type GraphGet = <T>(
  endpoint: string,
  allPages?: boolean,
  headers?: Record<string, string>,
) => Promise<T>;

export interface GraphPagesOptions<T> {
  /** Eigen bovengrens voor het aantal pagina's. */
  maxPages: number;
  headers?: Record<string, string>;
  /** Na elke pagina: `false` stopt het bladeren (de rest is niet nodig). */
  more?: (page: T[]) => boolean;
  /** Alleen voor tests: vervangt `graphGet`. */
  get?: GraphGet;
}

export interface GraphPages<T> {
  value: T[];
  pages: number;
  /** `true` als de bovengrens is geraakt terwijl er nog een volgende pagina was. */
  truncated: boolean;
}

/**
 * Volgt `@odata.nextLink` met een eigen bovengrens en een stopconditie.
 * Anders dan `allPages` (vaste MAX_PAGES, kapt stil af) gaat elke pagina door
 * `graphGet` (zelfde afhandeling van 429/401/403) en meldt het resultaat of
 * er is afgekapt.
 */
export async function graphGetPages<T>(
  endpoint: string,
  options: GraphPagesOptions<T>,
): Promise<GraphPages<T>> {
  const get = options.get ?? graphGet;
  const value: T[] = [];
  let next: string | undefined = endpoint;
  let pages = 0;

  while (next && pages < options.maxPages) {
    const page: GraphPagedResponse<T> = await get<GraphPagedResponse<T>>(
      next,
      false,
      options.headers,
    );
    const items = page.value ?? [];
    value.push(...items);
    pages++;
    next = page["@odata.nextLink"];
    if (next && options.more && !options.more(items)) {
      next = undefined;
    }
  }

  return { value, pages, truncated: next !== undefined };
}

export async function graphPost<T>(
  endpoint: string,
  body: unknown,
): Promise<T> {
  return graphRequest<T>("POST", endpoint, body);
}

export async function graphPatch<T>(
  endpoint: string,
  body: unknown,
): Promise<T> {
  return graphRequest<T>("PATCH", endpoint, body);
}

export async function graphDelete(endpoint: string): Promise<void> {
  await graphRequest<void>("DELETE", endpoint);
}
