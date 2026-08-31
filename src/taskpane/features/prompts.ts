export interface PromptPair {
  system: string;
  user: string;
}

export type EditPreset = "proofread" | "shorten" | "expand" | "clarify";

const SLIDE_TEXT_SYSTEM =
  "You edit PowerPoint slide text. Return only the final slide text with no preamble, quotes, or commentary. Keep line breaks that make sense on a slide.";

const PRESET_INSTRUCTIONS: Record<EditPreset, string> = {
  proofread: "Fix typos, grammar, and punctuation. Change nothing else.",
  shorten: "Rewrite this noticeably shorter while keeping the meaning. Aim for half the length.",
  expand: "Expand this with one or two supporting phrases while keeping the tone.",
  clarify: "Rewrite this so a first-time reader understands it immediately. Remove jargon.",
};

export function presetPrompt(text: string, preset: string): PromptPair {
  const instruction = PRESET_INSTRUCTIONS[preset as EditPreset];
  if (!instruction) throw new Error(`Unknown preset: ${preset}.`);
  return {
    system: SLIDE_TEXT_SYSTEM,
    user: `${instruction}\n\nSlide text:\n${text}`,
  };
}

export function editPrompt(text: string, instruction: string): PromptPair {
  return {
    system: SLIDE_TEXT_SYSTEM,
    user: `Apply this instruction to the slide text: ${instruction}\n\nSlide text:\n${text}`,
  };
}

export interface CreatedContent {
  title: string;
  bullets: string[];
}

export function createPrompt(topic: string): PromptPair {
  return {
    system:
      'You draft PowerPoint slide content. Respond with strict JSON only, shaped exactly as {"title": string, "bullets": string[]} with 3 to 5 concise bullets. No markdown fences, no commentary.',
    user: `Draft slide content for: ${topic}`,
  };
}

export function parseCreateResponse(raw: string): CreatedContent {
  const stripped = raw
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/```\s*$/, "")
    .trim();
  try {
    const parsed = JSON.parse(stripped) as Partial<CreatedContent>;
    if (
      typeof parsed.title !== "string" ||
      !Array.isArray(parsed.bullets) ||
      parsed.bullets.some((bullet) => typeof bullet !== "string")
    ) {
      throw new Error("bad shape");
    }
    return { title: parsed.title, bullets: parsed.bullets };
  } catch {
    throw new Error("Claude's response could not parse as slide content. Try again.");
  }
}

export function translatePrompt(text: string, language: string): PromptPair {
  return {
    system: SLIDE_TEXT_SYSTEM,
    user: `Translate this slide text into ${language}. Keep names and numbers as they are.\n\nSlide text:\n${text}`,
  };
}

const MAX_OUTLINE_CHARS = 8000;

export function darwinSystem(outline: string): string {
  const truncated =
    outline.length > MAX_OUTLINE_CHARS
      ? `${outline.slice(0, MAX_OUTLINE_CHARS)}\n[outline truncated]`
      : outline;
  return `You are Darwin, a presentation coach inside the Slideware PowerPoint add-in. Give direct, specific advice about this deck: structure, story, clarity, and design. Keep answers short and skimmable.\n\nCurrent deck outline:\n${truncated}`;
}
