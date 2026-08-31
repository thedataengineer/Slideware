export interface Agenda {
  lines: string[];
  text: string;
}

export function buildAgenda(titles: (string | undefined)[]): Agenda {
  const cleaned = titles.map((title) => (title ?? "").trim()).filter((title) => title.length > 0);

  if (cleaned.length === 0) {
    throw new Error("No slide titles found.");
  }

  const lines = cleaned.map((title, index) => `${index + 1}. ${title}`);
  return { lines, text: lines.join("\n") };
}
