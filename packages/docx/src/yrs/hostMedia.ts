/**
 * Replaces the `media:{n}` image tokens under `comments` with the `data:` URLs
 * `dataUrl` gives: comments reach the host as parsed, so every reader must be
 * able to show their images.
 */
export function resolveCommentMedia(
  comments: unknown,
  dataUrl: (token: string) => string | null
): void {
  if (Array.isArray(comments)) {
    for (const item of comments) resolveCommentMedia(item, dataUrl);
  } else if (comments && typeof comments === 'object') {
    const record = comments as Record<string, unknown>;
    for (const [key, field] of Object.entries(record)) {
      if (key === 'src' && typeof field === 'string') record[key] = dataUrl(field) ?? field;
      else resolveCommentMedia(field, dataUrl);
    }
  }
}

/** `resolveCommentMedia` over the comments of an `open_docx` host JSON reply. */
export function resolveHostJsonCommentMedia(
  json: string,
  dataUrl: (token: string) => string | null
): string {
  if (!json.includes('"media:')) return json;
  const host = JSON.parse(json) as {
    envelope?: { document?: { package?: { document?: { comments?: unknown } } } };
  };
  const comments = host.envelope?.document?.package?.document?.comments;
  if (comments === undefined) return json;
  resolveCommentMedia(comments, dataUrl);
  return JSON.stringify(host);
}
