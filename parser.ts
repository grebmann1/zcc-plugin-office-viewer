import { unzipSync, strFromU8 } from "fflate";
import { SaxesParser } from "saxes";
export type Section = {
  title: string;
  rows?: string[][];
  rowNumbers?: number[];
  paragraphs?: string[];
};
export type Document = {
  kind: "Word" | "Excel" | "PowerPoint";
  sections: Section[];
  limited: boolean;
};
type Node = {
  name: string;
  text: string;
  attrs: Record<string, string>;
  children: Node[];
};
const MAX_BYTES = 8 * 1024 * 1024,
  MAX_XML = 32 * 1024 * 1024;
function xml(text: string): Node {
  const root: Node = { name: "root", text: "", attrs: {}, children: [] };
  const stack = [root];
  let count = 0;
  const parser = new SaxesParser({ xmlns: false });
  parser.on("doctype", () => {
    throw new Error("Document type declarations are unsupported");
  });
  parser.on("opentag", (tag) => {
    if (stack.length > 100 || ++count > 150000)
      throw new Error("Document XML is too complex");
    const node: Node = {
      name: tag.name.split(":").pop()!,
      text: "",
      attrs: tag.attributes as Record<string, string>,
      children: [],
    };
    stack[stack.length - 1].children.push(node);
    stack.push(node);
  });
  parser.on("text", (s) => {
    stack[stack.length - 1].text += s;
  });
  parser.on("cdata", (s) => {
    stack[stack.length - 1].text += s;
  });
  parser.on("closetag", () => {
    stack.pop();
  });
  parser.write(text).close();
  return root;
}
const find = (node: Node, name: string): Node[] =>
  node.children.flatMap((c) => (c.name === name ? [c] : find(c, name)));
const allText = (node: Node): string =>
  node.text + node.children.map(allText).join("");
const textRuns = (node: Node): string => find(node, "t").map(allText).join("");
function column(ref: string): number {
  const m = /^([A-Z]{1,3})[1-9][0-9]*$/.exec(ref);
  if (!m) return -1;
  return [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
}
export function parseOffice(bytes: Uint8Array, extension: string): Document {
  if (bytes.length > MAX_BYTES)
    throw new Error("Office preview limit is 8 MiB");
  if (!["docx", "xlsx", "pptx"].includes(extension))
    throw new Error("Supported formats: docx, xlsx, pptx");
  let expanded = 0,
    entries = 0;
  const files = unzipSync(bytes, {
    filter(file) {
      if (++entries > 3000) throw new Error("Too many archive entries");
      expanded += file.originalSize;
      if (expanded > MAX_XML)
        throw new Error("Expanded document exceeds 32 MiB");
      return /^(word\/document\.xml|ppt\/slides\/slide[0-9]+\.xml|xl\/(sharedStrings\.xml|workbook\.xml|_rels\/workbook\.xml.rels|worksheets\/sheet[0-9]+\.xml))$/.test(
        file.name,
      );
    },
  });
  const part = (name: string) => {
    const b = files[name];
    if (!b) throw new Error("Missing Office document part: " + name);
    return xml(strFromU8(b));
  };
  if (extension === "docx") {
    const paragraphs = find(part("word/document.xml"), "p")
      .map(textRuns)
      .filter(Boolean);
    return {
      kind: "Word",
      sections: [
        {
          title: "Document",
          paragraphs: paragraphs.slice(0, 1000).map((p) => p.slice(0, 10000)),
        },
      ],
      limited:
        paragraphs.length > 1000 || paragraphs.some((p) => p.length > 10000),
    };
  }
  if (extension === "pptx") {
    const paths = Object.keys(files)
      .filter((p) => /^ppt\/slides\/slide[0-9]+\.xml$/.test(p))
      .sort(
        (a, b) =>
          Number(a.match(/slide([0-9]+)/)![1]) -
          Number(b.match(/slide([0-9]+)/)![1]),
      );
    if (!paths.length) throw new Error("No readable slides");
    let limited = paths.length > 100;
    const sections = paths.slice(0, 100).map((p, i) => {
      const paragraphs = find(part(p), "p").map(textRuns).filter(Boolean);
      if (paragraphs.length > 200 || paragraphs.some((p) => p.length > 10000))
        limited = true;
      return {
        title: `Slide ${i + 1}`,
        paragraphs: paragraphs.slice(0, 200).map((p) => p.slice(0, 10000)),
      };
    });
    return { kind: "PowerPoint", sections, limited };
  }
  const shared = files["xl/sharedStrings.xml"]
    ? find(part("xl/sharedStrings.xml"), "si").map(textRuns)
    : [];
  const names = new Map<string, string>();
  if (files["xl/workbook.xml"] && files["xl/_rels/workbook.xml.rels"]) {
    const rels = new Map(
      find(part("xl/_rels/workbook.xml.rels"), "Relationship")
        .filter((r) => r.attrs.TargetMode !== "External")
        .map((r) => [r.attrs.Id, r.attrs.Target]),
    );
    for (const s of find(part("xl/workbook.xml"), "sheet")) {
      const target = rels.get(s.attrs["r:id"]);
      if (target) {
        const path = target.startsWith("/xl/")
          ? target.slice(1)
          : "xl/" + target.replace(/^\.\//, "");
        names.set(path, s.attrs.name);
      }
    }
  }
  const paths = Object.keys(files)
    .filter((p) => /^xl\/worksheets\/sheet[0-9]+\.xml$/.test(p))
    .sort(
      (a, b) =>
        Number(a.match(/sheet([0-9]+)/)![1]) -
        Number(b.match(/sheet([0-9]+)/)![1]),
    );
  if (!paths.length) throw new Error("No readable worksheets");
  let limited = paths.length > 30;
  const sections = paths.slice(0, 30).map((p, i) => {
    const all = find(part(p), "row");
    if (all.length > 200) limited = true;
    const rows = all.slice(0, 200).map((row) => {
      const cells: string[] = [];
      for (const cell of row.children.filter((c) => c.name === "c")) {
        const index = column(cell.attrs.r ?? "");
        if (index < 0) continue;
        if (index >= 50) {
          limited = true;
          continue;
        }
        const value = cell.children.find((c) => c.name === "v");
        const raw = value ? allText(value) : "";
        const formula = cell.children.find((c) => c.name === "f");
        const output =
          cell.attrs.t === "s"
            ? (shared[Number(raw)] ?? "")
            : cell.attrs.t === "inlineStr"
              ? textRuns(cell)
              : raw || (formula ? "=" + allText(formula) : "");
        if (output.length > 2000) limited = true;
        cells[index] = output.slice(0, 2000);
      }
      return Array.from({ length: cells.length }, (_, n) => cells[n] ?? "");
    });
    return { title: names.get(p) || `Sheet ${i + 1}`, rows, rowNumbers:all.slice(0,200).map((row,i)=>{const n=Number(row.attrs.r);return Number.isInteger(n)&&n>=1&&n<=1048576?n:i+1}) };
  });
  return { kind: "Excel", sections, limited };
}
export { quoteSection } from './quote.js';
