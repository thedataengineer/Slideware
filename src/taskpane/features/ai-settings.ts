export type AiProvider = "claude" | "ollama" | "openai";

export interface AiSettings {
  provider: AiProvider;
  apiKey: string;
  claudeModel: string;
  ollamaUrl: string;
  ollamaModel: string;
  openaiUrl: string;
  openaiKey: string;
  openaiModel: string;
}

export function defaultAiSettings(): AiSettings {
  return {
    provider: "claude",
    apiKey: "",
    claudeModel: "claude-opus-5",
    ollamaUrl: "http://localhost:11434",
    ollamaModel: "llama3.2",
    openaiUrl: "https://api.openai.com/v1",
    openaiKey: "",
    openaiModel: "gpt-4o-mini",
  };
}

export function serializeAiSettings(settings: AiSettings): string {
  return JSON.stringify(settings);
}

function cleanUrl(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim().length > 0
    ? value.trim().replace(/\/+$/, "")
    : fallback;
}

function cleanText(value: unknown, fallback: string): string {
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : fallback;
}

export function parseAiSettings(raw: string | null): AiSettings {
  const defaults = defaultAiSettings();
  if (!raw) return defaults;
  try {
    const parsed = JSON.parse(raw) as Partial<AiSettings>;
    if (typeof parsed !== "object" || parsed === null) return defaults;
    const provider: AiProvider =
      parsed.provider === "claude" || parsed.provider === "ollama" || parsed.provider === "openai"
        ? parsed.provider
        : defaults.provider;
    return {
      provider,
      apiKey: typeof parsed.apiKey === "string" ? parsed.apiKey : defaults.apiKey,
      claudeModel: cleanText(parsed.claudeModel, defaults.claudeModel),
      ollamaUrl: cleanUrl(parsed.ollamaUrl, defaults.ollamaUrl),
      ollamaModel: cleanText(parsed.ollamaModel, defaults.ollamaModel),
      openaiUrl: cleanUrl(parsed.openaiUrl, defaults.openaiUrl),
      openaiKey: typeof parsed.openaiKey === "string" ? parsed.openaiKey : defaults.openaiKey,
      openaiModel: cleanText(parsed.openaiModel, defaults.openaiModel),
    };
  } catch {
    return defaults;
  }
}
