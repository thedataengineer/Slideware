import { cleanString, cleanStringList, extractJson } from "./model-json";

export const MAX_QUESTIONS = 4;
export const MIN_OPTIONS = 2;
export const MAX_OPTIONS = 4;
export const MAX_QUESTION_CHARS = 200;
export const MAX_OPTION_CHARS = 60;
export const MAX_DETECTED = 5;
export const MAX_DETECTED_CHARS = 200;
export const MAX_SUMMARY_CHARS = 300;

const UNREADABLE =
  "The model's analysis could not be read as JSON. Try again, or pick a stronger model in the Gen AI tab.";

/**
 * v1 asks only for single-select questions. The union exists so that adding a second kind
 * is a type error everywhere it matters rather than a silent behaviour change.
 */
export type QuestionKind = "choice";

export interface DeckQuestion {
  id: string;
  question: string;
  kind: QuestionKind;
  options: string[];
  /** Always equal to options[0], so a select element defaults to it with no extra code. */
  recommended: string;
}

export interface DeckAnalysis {
  summary: string;
  detected: string[];
  questions: DeckQuestion[];
}

export interface DeckAnswer {
  questionId: string;
  /** Carried so a prompt can be built from the answers alone. */
  question: string;
  answer: string;
}

function looseKey(value: string): string {
  return value
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[.!?,;:]+$/, "")
    .trim();
}

function questionsArray(root: unknown): unknown[] | undefined {
  if (Array.isArray(root)) return root;
  if (root && typeof root === "object") {
    const candidate = (root as { questions?: unknown }).questions;
    if (Array.isArray(candidate)) return candidate;
  }
  return undefined;
}

function toQuestion(raw: unknown, position: number): DeckQuestion | undefined {
  if (!raw || typeof raw !== "object") return undefined;
  const source = raw as { question?: unknown; options?: unknown; recommended?: unknown };

  const question = cleanString(source.question, MAX_QUESTION_CHARS);
  if (question.length === 0) return undefined;

  // Extra headroom: de-duping and adopting a missing recommendation happen before the hard cap.
  const options = cleanStringList(source.options, MAX_OPTIONS + 4, MAX_OPTION_CHARS);
  if (options.length < MIN_OPTIONS) return undefined;

  const recommended = cleanString(source.recommended, MAX_OPTION_CHARS);
  if (recommended.length === 0) return undefined;

  let ordered = options;
  const exact = options.indexOf(recommended);
  const loose =
    exact >= 0 ? exact : options.findIndex((option) => looseKey(option) === looseKey(recommended));
  if (loose >= 0) {
    ordered = [options[loose], ...options.filter((_, index) => index !== loose)];
  } else {
    // The model had an opinion it forgot to list. Keep the opinion and its alternatives.
    ordered = [recommended, ...options];
  }

  const capped = ordered.slice(0, MAX_OPTIONS);
  return {
    id: `q${position + 1}`,
    question,
    kind: "choice",
    options: capped,
    recommended: capped[0],
  };
}

/**
 * Throws when the model gave nothing. Clamps or drops when it gave too much or gave
 * individually bad items. Never invents content.
 */
export function parseDeckAnalysis(raw: string): DeckAnalysis {
  const root = extractJson(raw);
  const rawQuestions = questionsArray(root);
  if (rawQuestions === undefined) throw new Error(UNREADABLE);

  const container = Array.isArray(root) ? {} : (root as { summary?: unknown; detected?: unknown });
  const questions: DeckQuestion[] = [];
  rawQuestions.forEach((entry) => {
    if (questions.length >= MAX_QUESTIONS) return;
    const question = toQuestion(entry, questions.length);
    if (question) questions.push(question);
  });

  return {
    summary: cleanString(container.summary, MAX_SUMMARY_CHARS),
    detected: cleanStringList(container.detected, MAX_DETECTED, MAX_DETECTED_CHARS),
    questions,
  };
}

/** Used when the model returns no usable question, so the feature still works on a small model. */
export function standardQuestions(): DeckQuestion[] {
  return [
    {
      id: "q1",
      question: "Who is the audience?",
      kind: "choice",
      options: ["Internal exec review", "Client or prospect", "Engineering team"],
      recommended: "Internal exec review",
    },
    {
      id: "q2",
      question: "How many slides?",
      kind: "choice",
      options: ["About 7", "About 5", "About 12"],
      recommended: "About 7",
    },
    {
      id: "q3",
      question: "How should the slides look?",
      kind: "choice",
      options: ["Bullets throughout", "Section dividers between topics", "Few words per slide"],
      recommended: "Bullets throughout",
    },
    {
      id: "q4",
      question: "End with a next steps slide?",
      kind: "choice",
      options: ["Yes", "No"],
      recommended: "Yes",
    },
  ];
}

export function withFallbackQuestions(analysis: DeckAnalysis): DeckAnalysis {
  if (analysis.questions.length > 0) return analysis;
  return { ...analysis, questions: standardQuestions() };
}

/** The click-straight-through path: accept every recommendation. */
export function defaultAnswers(questions: DeckQuestion[]): DeckAnswer[] {
  return questions.map((question) => ({
    questionId: question.id,
    question: question.question,
    answer: question.recommended,
  }));
}

export function setAnswer(answers: DeckAnswer[], questionId: string, value: string): DeckAnswer[] {
  if (!answers.some((answer) => answer.questionId === questionId)) {
    throw new Error(`Unknown question: ${questionId}.`);
  }
  return answers.map((answer) =>
    answer.questionId === questionId ? { ...answer, answer: value } : answer
  );
}
