import { useState } from "react";
import { definePluginApp, useComposer } from "@zana-ai/zcc-plugin-sdk/app";
import type { PluginFileOpenerProps } from "@zana-ai/zcc-plugin-sdk/app";
import { Frame, Notice, control, useData, usePage } from "./ui.js";
import { quoteSection } from "./quote.js";
import type { Document } from "./parser.js";
export function Viewer({ path, source }: PluginFileOpenerProps) {
  const { data, error, pending, reload } = useData<Document>("preview", {
    path,
    source,
  });
  const [index, setIndex] = useState(0);
  const [message, setMessage] = useState("");
  const composer = useComposer();
  const section = data?.sections[index];
  usePage("Office Viewer", path);
  return (
    <Frame title={path.split(/[\\/]/).pop() || "Office Viewer"}>
      <Notice pending={pending} error={error} />
      <button style={control} onClick={reload}>
        Reload document
      </button>
      {data && (
        <>
          <p>{data.kind} · content preview</p>
          <p style={{ opacity: 0.7 }}>
            Text and cell values. Images, layout, charts and formula calculation
            are unavailable.
          </p>
          {data.limited && (
            <p role="status">
              This preview is limited. Showing up to 1,000 paragraphs, 100
              slides, or 30 sheets (200 rows × 50 columns).
            </p>
          )}
          <label>
            Section{" "}
            <select
              style={control}
              aria-label="Section"
              value={index}
              onChange={(e) => {
                setIndex(Number(e.target.value));
                setMessage("");
              }}
            >
              {data.sections.map((s, i) => (
                <option key={i} value={i}>
                  {s.title}
                </option>
              ))}
            </select>
          </label>
          {section && (
            <>
              <h3>{section.title}</h3>
              {section.rows ? (
                <div style={{ overflowX: "auto" }}>
                  <table style={{ borderCollapse: "collapse", fontSize: 13 }}>
                    <tbody>
                      {section.rows.map((row, i) => (
                        <tr key={i}>
                          <th style={{ padding: 8, border: "1px solid #666" }}>
                            {section.rowNumbers?.[i] ?? i + 1}
                          </th>
                          {row.map((cell, j) => (
                            <td
                              key={j}
                              style={{
                                padding: 8,
                                border: "1px solid #666",
                                maxWidth: 420,
                                whiteSpace: "pre-wrap",
                                overflowWrap: "anywhere",
                              }}
                            >
                              {cell}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : (
                section.paragraphs?.map((p, i) => (
                  <p
                    key={i}
                    style={{ whiteSpace: "pre-wrap", overflowWrap: "anywhere" }}
                  >
                    {p}
                  </p>
                ))
              )}
              {(section.rows?.length ?? section.paragraphs?.length ?? 0) ===
                0 && <p>No readable text in this section.</p>}
              <button
                style={{ ...control, marginTop: 16 }}
                onClick={() => {
                  try {
                    composer.addQuote(quoteSection(path, section));
                    composer.focus();
                    setMessage("Quote added to the conversation draft");
                  } catch (cause) {
                    setMessage(
                      String(cause instanceof Error ? cause.message : cause),
                    );
                  }
                }}
              >
                Quote section
              </button>
              {message && <p role="status">{message}</p>}
            </>
          )}
        </>
      )}
    </Frame>
  );
}
export function Overview() {
  usePage("Office Viewer");
  return (
    <Frame title="Office Viewer">
      <p>
        Open a .docx, .xlsx or .pptx file from a project or thread attachment to
        read it beside the conversation.
      </p>
      <p>
        Choose a sheet or slide and use “Quote section” to add its text to your
        draft.
      </p>
      <p>
        Read-only content previews support files up to 8 MiB.
        Password-protected, macro and legacy binary formats are unsupported.
      </p>
    </Frame>
  );
}
export default definePluginApp((app) => {
  app.slots.fileOpener({
    id: "office",
    title: "Office Viewer",
    extensions: ["docx", "xlsx", "pptx"],
    component: Viewer,
  });
  app.slots.navPanel({
    id: "main",
    title: "Office Viewer",
    icon: "FileText",
    placement: "extensions",
    component: Overview,
  });
});
