/**
 * Local models answer with fenced blocks, chat preamble, trailing prose, and trailing commas.
 * These helpers pull a usable value out of that text without inventing content: a failure
 * returns null or an empty value so each caller can own its own error sentence.
 */

function findFirstOpen(text: string): number {
  const brace = text.indexOf("{");
  const bracket = text.indexOf("[");
  if (brace < 0) return bracket;
  if (bracket < 0) return brace;
  return Math.min(brace, bracket);
}

/** Slice from the first opening bracket to the matching close, ignoring string contents. */
function balancedSlice(text: string): string | null {
  const start = findFirstOpen(text);
  if (start < 0) return null;

  const open = text[start];
  const close = open === "{" ? "}" : "]";
  let depth = 0;
  let inString = false;
  let escaped = false;

  for (let index = start; index < text.length; index += 1) {
    const char = text[index];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (char === "\\") {
      if (inString) escaped = true;
      continue;
    }
    if (char === '"') {
      inString = !inString;
      continue;
    }
    if (inString) continue;

    if (char === open) {
      depth += 1;
    } else if (char === close) {
      depth -= 1;
      if (depth === 0) return text.slice(start, index + 1);
    }
  }
  return null;
}

/** Drop commas that sit directly before a closing bracket. Never touches string contents. */
function dropTrailingCommas(text: string): string {
  let out = "";
  let inString = false;
  let escaped = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (escaped) {
      out += char;
      escaped = false;
      continue;
    }
    if (char === "\\") {
      out += char;
      if (inString) escaped = true;
      continue;
    }
    if (char === '"') {
      inString = !inString;
      out += char;
      continue;
    }
    if (!inString && char === "," && /^\s*[}\]]/.test(text.slice(index + 1))) {
      continue;
    }
    out += char;
  }
  return out;
}

function tryParse(text: string): { value: unknown } | undefined {
  try {
    return { value: JSON.parse(text) };
  } catch {
    return undefined;
  }
}

/**
 * Pull the first JSON object or array out of sloppy model text. Returns null rather than
 * throwing so the caller decides what the user is told.
 */
export function extractJson(raw: string): unknown | null {
  if (typeof raw !== "string" || raw.trim().length === 0) return null;

  const candidates: string[] = [];
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced && fenced[1]) candidates.push(fenced[1]);
  candidates.push(raw);

  for (const candidate of candidates) {
    const sliced = balancedSlice(candidate);
    if (!sliced) continue;
    const parsed = tryParse(sliced) ?? tryParse(dropTrailingCommas(sliced));
    if (parsed) return parsed.value;
  }
  return null;
}

/** Trim, collapse inner whitespace, clamp length. Anything non-string becomes "". */
export function cleanString(value: unknown, maxChars: number): string {
  if (typeof value !== "string") return "";
  const collapsed = value.replace(/\s+/g, " ").trim();
  return collapsed.length > maxChars ? collapsed.slice(0, maxChars).trim() : collapsed;
}

/** Clean every entry, drop empties, de-dupe case-insensitively, clamp the count. */
export function cleanStringList(value: unknown, maxItems: number, maxChars: number): string[] {
  if (!Array.isArray(value)) return [];

  const seen = new Set<string>();
  const out: string[] = [];
  for (const item of value) {
    const cleaned = cleanString(item, maxChars);
    if (cleaned.length === 0) continue;
    const key = cleaned.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(cleaned);
    if (out.length >= maxItems) break;
  }
  return out;
}
