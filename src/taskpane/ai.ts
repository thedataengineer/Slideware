import Anthropic from "@anthropic-ai/sdk";
import {
  AiSettings,
  defaultAiSettings,
  parseAiSettings,
  serializeAiSettings,
} from "./features/ai-settings";

/* global localStorage, fetch, Response */

export const AI_MODEL = "claude-opus-5";
const AI_SETTINGS_STORAGE_KEY = "slideware.ai";
const LEGACY_API_KEY_STORAGE_KEY = "slideware.apiKey";

export interface AiRequest {
  system: string;
  messages: { role: "user" | "assistant"; content: string }[];
  maxTokens?: number;
}

export function loadAiSettings(): AiSettings {
  try {
    const settings = parseAiSettings(localStorage.getItem(AI_SETTINGS_STORAGE_KEY));
    if (!settings.apiKey) {
      const legacyKey = localStorage.getItem(LEGACY_API_KEY_STORAGE_KEY);
      if (legacyKey) settings.apiKey = legacyKey;
    }
    return settings;
  } catch {
    return defaultAiSettings();
  }
}

export function saveAiSettings(settings: AiSettings): void {
  localStorage.setItem(AI_SETTINGS_STORAGE_KEY, serializeAiSettings(settings));
}

async function callClaude(settings: AiSettings, request: AiRequest): Promise<string> {
  if (!settings.apiKey) {
    throw new Error("Add your Anthropic API key in the Gen AI tab first.");
  }

  const client = new Anthropic({ apiKey: settings.apiKey, dangerouslyAllowBrowser: true });

  try {
    const response = await client.messages.create({
      model: AI_MODEL,
      max_tokens: request.maxTokens ?? 16000,
      system: request.system,
      messages: request.messages,
    });

    if (response.stop_reason === "refusal") {
      throw new Error("Claude declined this request.");
    }

    const text = response.content
      .filter((block): block is Anthropic.TextBlock => block.type === "text")
      .map((block) => block.text)
      .join("\n")
      .trim();

    if (!text) {
      throw new Error("Claude returned no text. Try again.");
    }
    return text;
  } catch (error) {
    if (error instanceof Anthropic.AuthenticationError) {
      throw new Error("The Claude API rejected the key. Check it in the Gen AI tab.");
    }
    if (error instanceof Anthropic.RateLimitError) {
      throw new Error("The Claude API rate-limited this request. Try again shortly.");
    }
    if (error instanceof Anthropic.APIError) {
      throw new Error(
        `Claude API error${error.status ? ` ${error.status}` : ""}: ${error.message}`
      );
    }
    throw error;
  }
}

interface OllamaChatResponse {
  message?: { content?: string };
  error?: string;
}

async function callOllama(settings: AiSettings, request: AiRequest): Promise<string> {
  let response: Response;
  try {
    response = await fetch(`${settings.ollamaUrl}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: settings.ollamaModel,
        stream: false,
        messages: [
          { role: "system", content: request.system },
          ...request.messages.map((message) => ({ role: message.role, content: message.content })),
        ],
      }),
    });
  } catch {
    throw new Error(
      `Could not reach Ollama at ${settings.ollamaUrl}. Is it running? Start it with "ollama serve", and if the request is blocked set OLLAMA_ORIGINS="*".`
    );
  }

  if (response.status === 404) {
    throw new Error(
      `Ollama does not have the model "${settings.ollamaModel}". Pull it with "ollama pull ${settings.ollamaModel}".`
    );
  }
  if (!response.ok) {
    throw new Error(`Ollama error ${response.status}: ${await response.text()}`);
  }

  const data = (await response.json()) as OllamaChatResponse;
  if (data.error) {
    throw new Error(`Ollama error: ${data.error}`);
  }
  const text = (data.message?.content ?? "").trim();
  if (!text) {
    throw new Error("Ollama returned no text. Try again or switch models.");
  }
  return text;
}

export async function callAi(request: AiRequest): Promise<string> {
  const settings = loadAiSettings();
  return settings.provider === "ollama"
    ? callOllama(settings, request)
    : callClaude(settings, request);
}
