import type { Section } from './parser.js';
export function quoteSection(path: string, section: Section): string {
  const body = section.rows
    ? section.rows.map((r) => r.join("\t")).join("\n")
    : (section.paragraphs ?? []).join("\n");
  return `${path} — ${section.title}\n${body}`.slice(0, 12000);
}
