import Anthropic from "@anthropic-ai/sdk";

export const AI_MODEL = "claude-opus-5";

export interface AiRequest {
  apiKey: string;
  system: string;
  messages: Anthropic.MessageParam[];
  maxTokens?: number;
}

export async function callClaude(request: AiRequest): Promise<string> {
  if (!request.apiKey) {
    throw new Error("Add your Anthropic API key in the Gen AI tab first.");
  }

  const client = new Anthropic({ apiKey: request.apiKey, dangerouslyAllowBrowser: true });

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
