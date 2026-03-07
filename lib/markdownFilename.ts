const INVALID_FILE_CHARS_RE = /[/\\?%*:|"<>]/g;

export function getFirstH1Title(markdown: string): string {
  const lines = markdown.split(/\r?\n/);
  for (const line of lines) {
    if (!line.startsWith("# ")) {
      continue;
    }

    const title = line.slice(2).trim();
    if (title.length > 0) {
      return title;
    }
  }

  return "";
}

export function sanitizeFilenameBase(name: string, maxLength = 80): string {
  const cleaned = name
    .replace(INVALID_FILE_CHARS_RE, "_")
    .replace(/\s+/g, " ")
    .trim();

  return cleaned.length > maxLength ? cleaned.slice(0, maxLength).trim() : cleaned;
}

export function normalizeDocxFilename(input: string | undefined, fallback = "md2doc.docx"): string {
  const safeFallback = fallback.toLowerCase().endsWith(".docx") ? fallback : `${fallback}.docx`;

  const base = sanitizeFilenameBase((input ?? "").trim(), 120);
  if (!base) {
    return safeFallback;
  }

  return base.toLowerCase().endsWith(".docx") ? base : `${base}.docx`;
}

export function buildFilenameFromMarkdown(markdown: string, fallback = "md2doc.docx"): string {
  const title = sanitizeFilenameBase(getFirstH1Title(markdown));
  if (!title) {
    return normalizeDocxFilename(fallback, fallback);
  }

  return normalizeDocxFilename(`${title}.docx`, fallback);
}
