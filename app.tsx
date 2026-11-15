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
function SettingInfo({ label, description, value }: {
  label: string;
  description: string;
  value: string;
}) {
  return (
    <div className="settings-field" style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 24 }}>
      <dt className="settings-toggle-copy">
        <span className="settings-label">{label}</span>
        <p className="settings-help">{description}</p>
      </dt>
      <dd style={{ margin: 0, flexShrink: 0, color: "var(--text-muted)", fontSize: 12 }}>{value}</dd>
    </div>
  );
}
export function OfficeViewerSettings() {
  usePage("Office Viewer");
  return (
    <div aria-label="Office Viewer preferences">
      <section className="settings-section" aria-labelledby="office-viewer-documents">
        <h3 id="office-viewer-documents">Supported documents</h3>
        <dl className="settings-section-body" style={{ margin: 0 }}>
          <SettingInfo label="Word documents" description="Preview document paragraphs." value=".docx" />
          <SettingInfo label="Excel workbooks" description="Preview sheets with stored cell values or formula text." value=".xlsx" />
          <SettingInfo label="PowerPoint presentations" description="Preview slide text." value=".pptx" />
        </dl>
      </section>
      <section className="settings-section" aria-labelledby="office-viewer-behavior">
        <h3 id="office-viewer-behavior">Preview behavior</h3>
        <dl className="settings-section-body" style={{ margin: 0 }}>
          <SettingInfo label="Preview mode" description="Text and cell values. Images, layout, charts and formula calculation are unavailable." value="Read-only" />
          <SettingInfo label="File size limit" description="Password-protected, macro and legacy binary formats are unsupported." value="8 MiB" />
        </dl>
      </section>
      <section className="settings-section" aria-labelledby="office-viewer-usage">
        <h3 id="office-viewer-usage">Using Office Viewer</h3>
        <div className="settings-section-body">
          <details>
            <summary style={{ cursor: "pointer" }}>How to preview and quote a document</summary>
            <p className="settings-help">
              Open a .docx, .xlsx or .pptx file from a project or thread attachment.
              Choose a sheet or slide, then select “Quote section” to add its text
              to the conversation draft.
            </p>
          </details>
        </div>
        <p className="settings-help">Supported files open automatically. No additional configuration is required.</p>
      </section>
    </div>
  );
}
export default definePluginApp((app) => {
  app.slots.fileOpener({
    id: "office",
    title: "Office Viewer",
    extensions: ["docx", "xlsx", "pptx"],
    component: Viewer,
  });
  app.slots.settingsSection({
    id: "preview",
    component: OfficeViewerSettings,
  });
});
