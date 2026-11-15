import { expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { parseOffice, quoteSection } from "./parser.js";
const archive = (parts: Record<string, string>) =>
  zipSync(
    Object.fromEntries(Object.entries(parts).map(([k, v]) => [k, strToU8(v)])),
  );
const doc = (content: string) =>
  archive({
    "word/document.xml": `<w:document xmlns:w="word"><w:body>${content}</w:body></w:document>`,
  });
it("extracts Word paragraphs and quotes bounded text without rendering markup", () => {
  const result = parseOffice(
    doc(
      "<w:p><w:r><w:t>One &amp; two</w:t></w:r><w:r><w:t><![CDATA[ <script> ]]></w:t></w:r></w:p><w:p/>",
    ),
    "docx",
  );
  expect(result).toEqual({
    kind: "Word",
    sections: [{ title: "Document", paragraphs: ["One & two <script> "] }],
    limited: false,
  });
  expect(quoteSection("report.docx", result.sections[0])).toContain(
    "report.docx — Document\nOne & two",
  );
  expect(quoteSection("x", { title: "Empty" })).toBe("x — Empty\n");
  expect(
    quoteSection("x", { title: "Long", paragraphs: ["a".repeat(13000)] }),
  ).toHaveLength(12000);
});
it("orders numeric slides and limits text-heavy documents", () => {
  const result = parseOffice(
    archive({
      "ppt/slides/slide10.xml": "<s><p><t>Ten</t></p></s>",
      "ppt/slides/slide2.xml": "<s><p><t>Two</t></p></s>",
    }),
    "pptx",
  );
  expect(result.sections.map((s) => s.paragraphs)).toEqual([["Two"], ["Ten"]]);
  expect(
    parseOffice(doc("<w:p><w:t>" + "x".repeat(10001) + "</w:t></w:p>"), "docx")
      .limited,
  ).toBe(true);
  expect(
    parseOffice(doc("<w:p><w:t>x</w:t></w:p>".repeat(1001)), "docx").sections[0]
      .paragraphs,
  ).toHaveLength(1000);
  const slides = Object.fromEntries(
    Array.from({ length: 101 }, (_, i) => [
      `ppt/slides/slide${i + 1}.xml`,
      "<s><p><t>x</t></p></s>",
    ]),
  );
  expect(parseOffice(archive(slides), "pptx")).toMatchObject({ limited: true });
  expect(
    parseOffice(
      archive({
        "ppt/slides/slide1.xml": "<s>" + "<p><t>x</t></p>".repeat(201) + "</s>",
      }),
      "pptx",
    ).limited,
  ).toBe(true);
});
it("reads shared, inline, cached formula and blank sparse cells, with workbook names", () => {
  const result = parseOffice(
    archive({
      "xl/workbook.xml":
        '<workbook><sheets><sheet name="Forecast" r:id="r1"/></sheets></workbook>',
      "xl/_rels/workbook.xml.rels":
        '<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>',
      "xl/sharedStrings.xml": "<sst><si><t>Quarter</t></si></sst>",
      "xl/worksheets/sheet1.xml":
        '<worksheet><row><c r="A1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>Forecast</t></is></c><c r="D1"><f>1+2</f><v>3</v></c><c r="E1"><f>NOW()</f></c><c r="F1" t="s"><v>90</v></c><c><v>bad</v></c><c r="ZZZ1"><v>outside</v></c></row></worksheet>',
    }),
    "xlsx",
  );
  expect(result.sections[0]).toEqual({
    title: "Forecast",
    rowNumbers: [1],
    rows: [["Quarter", "", "Forecast", "3", "=NOW()", ""]],
  });
  expect(result.limited).toBe(true);
  expect(quoteSection("a.xlsx", result.sections[0])).toContain(
    "Quarter\t\tForecast\t3",
  );
});
it("limits rows/sheets and clips oversized cell text", () => {
  const parts = Object.fromEntries(
    Array.from({ length: 31 }, (_, i) => [
      `xl/worksheets/sheet${i + 1}.xml`,
      "<worksheet>" +
        '<row><c r="A1"><v>1</v></c></row>'.repeat(201) +
        "</worksheet>",
    ]),
  );
  expect(parseOffice(archive(parts), "xlsx")).toMatchObject({ limited: true });
  const large = parseOffice(
    archive({
      "xl/worksheets/sheet1.xml":
        '<w><row><c r="A1" t="inlineStr"><t>' +
        "x".repeat(2001) +
        "</t></c></row></w>",
    }),
    "xlsx",
  );
  expect(large.sections[0].rows![0][0]).toHaveLength(2000);
  expect(large.limited).toBe(true);
});
it("handles absolute relationship targets, missing names and empty cell values", () => {
  const result = parseOffice(
    archive({
      "xl/workbook.xml": '<w><sheet name="Numbers" r:id="r1"/></w>',
      "xl/_rels/workbook.xml.rels":
        '<r><Relationship Id="r1" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="ext" Target="https://invalid" TargetMode="External"/></r>',
      "xl/worksheets/sheet1.xml": '<w><row r="12"><c r="A12"/></row></w>',
    }),
    "xlsx",
  );
  expect(result.sections[0]).toEqual({ title: "Numbers", rowNumbers: [12], rows: [[""]] });
});
it("rejects wrong formats, corrupt/encrypted archives, missing parts, DTDs and deep XML", () => {
  expect(() =>
    parseOffice(new Uint8Array(8 * 1024 * 1024 + 1), "docx"),
  ).toThrow("8 MiB");
  expect(() => parseOffice(new Uint8Array(), "doc")).toThrow("Supported");
  expect(() => parseOffice(new Uint8Array([1, 2]), "docx")).toThrow();
  expect(() => parseOffice(archive({}), "docx")).toThrow("Missing");
  for (const ext of ["xlsx", "pptx"])
    expect(() => parseOffice(archive({}), ext)).toThrow("No readable");
  expect(() =>
    parseOffice(
      archive({
        "word/document.xml":
          '<!DOCTYPE d [<!ENTITY x SYSTEM "file:///etc/passwd">]><d>&x;</d>',
      }),
      "docx",
    ),
  ).toThrow("declarations");
  expect(() =>
    parseOffice(doc("<a>".repeat(102) + "</a>".repeat(102)), "docx"),
  ).toThrow("complex");
});
it("rejects zip expansion and excessive entry counts before extraction", () => {
  const bomb = archive({ "unused.bin": "x" });
  const view = new DataView(bomb.buffer, bomb.byteOffset, bomb.byteLength);
  for (let i = 0; i < bomb.length - 4; i++)
    if (view.getUint32(i, true) === 0x02014b50)
      view.setUint32(i + 24, 32 * 1024 * 1024 + 1, true);
  expect(() => parseOffice(bomb, "docx")).toThrow("32 MiB");
  expect(() =>
    parseOffice(
      archive(
        Object.fromEntries(
          Array.from({ length: 3001 }, (_, i) => [String(i), ""]),
        ),
      ),
      "docx",
    ),
  ).toThrow("entries");
});
