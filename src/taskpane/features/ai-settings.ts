export type AiProvider = "claude" | "ollama";

export interface AiSettings {
  provider: AiProvider;
  apiKey: string;
  ollamaUrl: string;
  ollamaModel: string;
}

export function defaultAiSettings(): AiSettings {
  return {
    provider: "claude",
    apiKey: "",
    ollamaUrl: "http://localhost:11434",
    ollamaModel: "llama3.2",
  };
}

export function serializeAiSettings(settings: AiSettings): string {
  return JSON.stringify(settings);
}

export function parseAiSettings(raw: string | null): AiSettings {
  const defaults = defaultAiSettings();
  if (!raw) return defaults;
  try {
    const parsed = JSON.parse(raw) as Partial<AiSettings>;
    if (typeof parsed !== "object" || parsed === null) return defaults;
    const provider: AiProvider =
      parsed.provider === "claude" || parsed.provider === "ollama"
        ? parsed.provider
        : defaults.provider;
    const ollamaUrl =
      typeof parsed.ollamaUrl === "string" && parsed.ollamaUrl.trim().length > 0
        ? parsed.ollamaUrl.trim().replace(/\/+$/, "")
        : defaults.ollamaUrl;
    return {
      provider,
      apiKey: typeof parsed.apiKey === "string" ? parsed.apiKey : defaults.apiKey,
      ollamaUrl,
      ollamaModel:
        typeof parsed.ollamaModel === "string" && parsed.ollamaModel.trim().length > 0
          ? parsed.ollamaModel.trim()
          : defaults.ollamaModel,
    };
  } catch {
    return defaults;
  }
}
