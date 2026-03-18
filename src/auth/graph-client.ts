import { getAccessToken } from "./msal.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
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

  let url = `${GRAPH_BASE}${endpoint}`;
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
      throw new Error(
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
