/**
 * Centralized API Fetcher for F2 Fintech
 *
 * Guarantees that responses from /api endpoints are valid JSON.
 * If HTML is returned (e.g. from Vite SPA catch-all, proxy errors, or server crashes),
 * it throws a human-readable, descriptive error rather than letting JSON.parse() crash
 * with "Unexpected token '<'".
 */

export interface ApiFetchOptions extends RequestInit {
  timeoutMs?: number;
}

export async function apiFetch<T = any>(url: string, options: ApiFetchOptions = {}): Promise<T> {
  const { timeoutMs = 60000, ...fetchOpts } = options;

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, {
      ...fetchOpts,
      signal: fetchOpts.signal || controller.signal,
      headers: {
        Accept: 'application/json',
        ...(fetchOpts.headers || {}),
      },
    });

    clearTimeout(timeoutId);

    const contentType = res.headers.get('content-type') || '';
    const isJson = contentType.includes('application/json');

    if (!isJson) {
      const text = await res.text();
      const preview = text.trim().slice(0, 100);
      if (preview.startsWith('<!doctype') || preview.startsWith('<html') || preview.startsWith('<')) {
        throw new Error(
          `Server returned HTML instead of JSON for ${url} (HTTP ${res.status}). The requested route may be missing, or the server returned an error page.`
        );
      }
      throw new Error(`Server returned unexpected Content-Type "${contentType}" for ${url}: ${preview}`);
    }

    const data = await res.json();

    if (!res.ok) {
      const errMsg = (data && data.error) || (data && data.message) || `Request failed with HTTP status ${res.status}`;
      const err = new Error(errMsg);
      (err as any).status = res.status;
      (err as any).data = data;
      throw err;
    }

    return data as T;
  } catch (err: any) {
    clearTimeout(timeoutId);
    if (err.name === 'AbortError') {
      throw new Error(`Request timed out after ${timeoutMs / 1000}s: ${url}`);
    }
    throw err;
  }
}
