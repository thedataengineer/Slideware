import Anthropic from "@anthropic-ai/sdk";
import {
  AiSettings,
  defaultAiSettings,
  parseAiSettings,
  serializeAiSettings,
} from "./features/ai-settings";

/* global localStorage, fetch, Response */

export const AI_MODEL = "claude-opus-5";

export const CLAUDE_MODELS = [
  "claude-opus-5",
  "claude-fable-5",
  "claude-sonnet-5",
  "claude-opus-4-8",
  "claude-opus-4-7",
  "claude-opus-4-6",
  "claude-sonnet-4-6",
  "claude-haiku-4-5",
];
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
    // Office webviews block direct http:// requests from the https:// pane (mixed
    // content), so the old direct default is migrated to the dev-server proxy path.
    if (settings.ollamaUrl === "http://localhost:11434") {
      settings.ollamaUrl = "/ollama";
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
      model: settings.claudeModel || AI_MODEL,
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
      `Could not reach Ollama at ${settings.ollamaUrl}. Is it running? Start it with "ollama serve". Use the "/ollama" URL inside PowerPoint; direct http URLs are blocked by the Office webview.`
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

export async function listClaudeModels(apiKey: string): Promise<string[]> {
  if (!apiKey) return CLAUDE_MODELS;
  try {
    const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
    const models: string[] = [];
    for await (const model of client.models.list()) {
      models.push(model.id);
    }
    return models.length > 0 ? models : CLAUDE_MODELS;
  } catch {
    return CLAUDE_MODELS;
  }
}

interface OpenAiChatResponse {
  choices?: { message?: { content?: string } }[];
  error?: { message?: string };
}

async function callOpenAi(settings: AiSettings, request: AiRequest): Promise<string> {
  let response: Response;
  try {
    response = await fetch(`${settings.openaiUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...(settings.openaiKey ? { Authorization: `Bearer ${settings.openaiKey}` } : {}),
      },
      body: JSON.stringify({
        model: settings.openaiModel,
        messages: [
          { role: "system", content: request.system },
          ...request.messages.map((message) => ({ role: message.role, content: message.content })),
        ],
      }),
    });
  } catch {
    throw new Error(
      `Could not reach ${settings.openaiUrl}. Check the base URL in the Gen AI tab; the server must allow browser requests (CORS).`
    );
  }

  if (response.status === 401 || response.status === 403) {
    throw new Error("The provider rejected the API key. Check it in the Gen AI tab.");
  }
  if (!response.ok) {
    const body = (await response.json().catch(() => ({}))) as OpenAiChatResponse;
    throw new Error(
      `Provider error ${response.status}: ${body.error?.message ?? "request failed"}`
    );
  }

  const data = (await response.json()) as OpenAiChatResponse;
  const text = (data.choices?.[0]?.message?.content ?? "").trim();
  if (!text) {
    throw new Error("The provider returned no text. Try again or switch models.");
  }
  return text;
}

interface OpenAiModelsResponse {
  data?: { id?: string }[];
}

export async function listOpenAiModels(url: string, key: string): Promise<string[]> {
  try {
    const response = await fetch(`${url}/models`, {
      headers: key ? { Authorization: `Bearer ${key}` } : {},
    });
    if (!response.ok) return [];
    const data = (await response.json()) as OpenAiModelsResponse;
    return (data.data ?? []).map((model) => model.id ?? "").filter((id) => id.length > 0);
  } catch {
    return [];
  }
}

interface OllamaTagsResponse {
  models?: { name?: string; capabilities?: string[] }[];
}

export async function listOllamaModels(url: string): Promise<string[]> {
  const response = await fetch(`${url}/api/tags`);
  if (!response.ok) return [];
  const data = (await response.json()) as OllamaTagsResponse;
  return (data.models ?? [])
    .filter((model) => !model.capabilities || model.capabilities.includes("completion"))
    .map((model) => model.name ?? "")
    .filter((name) => name.length > 0);
}

export async function callAi(request: AiRequest): Promise<string> {
  const settings = loadAiSettings();
  if (settings.provider === "ollama") return callOllama(settings, request);
  if (settings.provider === "openai") return callOpenAi(settings, request);
  return callClaude(settings, request);
}
