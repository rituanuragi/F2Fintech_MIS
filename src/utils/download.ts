/** POST JSON to an export route and save the returned file in the browser. */
export async function downloadFromPost(url: string, body: unknown, filename: string): Promise<void> {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.json().catch(() => null);
    throw new Error((err && err.error) || `Export failed (HTTP ${res.status})`);
  }
  const blob = await res.blob();
  const href = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = href;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  window.URL.revokeObjectURL(href);
}
