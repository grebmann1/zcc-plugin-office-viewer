// host:react
var m = globalThis.__bbPluginRuntime?.react;
if (!m) throw new Error("Open this plugin in Zana");
var { useState, useEffect, useCallback, useRef, createElement, Fragment } = m;

// host:@zana-ai/zcc-plugin-sdk/app
var m2 = globalThis.__bbPluginRuntime?.pluginSdkApp;
if (!m2) throw new Error("Open this plugin in Zana");
var { definePluginApp, useRpc, useComposer, useZccNavigate } = m2;

// host:react/jsx-runtime
var m3 = globalThis.__bbPluginRuntime?.jsxRuntime;
if (!m3) throw new Error("Open this plugin in Zana");
var { jsx, jsxs, Fragment: Fragment2 } = m3;

// ui.tsx
var box = {
  height: "100%",
  overflow: "auto",
  padding: 24,
  boxSizing: "border-box",
  color: "var(--text-primary,inherit)",
  fontFamily: "inherit"
};
var control = {
  background: "var(--bg-elevated,#27292d)",
  color: "inherit",
  border: "1px solid var(--border-default,#555)",
  borderRadius: 7,
  padding: "8px 12px",
  font: "inherit"
};
function Frame({
  title,
  children
}) {
  return /* @__PURE__ */ jsxs("main", { style: box, children: [
    /* @__PURE__ */ jsx("h2", { style: { marginTop: 0 }, children: title }),
    children
  ] });
}
function usePage(title, resourceId) {
  const nav = useZccNavigate();
  useEffect(() => {
    nav.setPageContext?.({
      title,
      resourceId,
      view: resourceId ? "detail" : "list"
    });
    return () => nav.setPageContext?.(null);
  }, [nav.setPageContext, title, resourceId]);
}
function useData(method, args) {
  const rpc = useRpc();
  const key = JSON.stringify(args);
  const [data, setData] = useState(null);
  const [error, setError] = useState("");
  const [pending, setPending] = useState(false);
  const [revision, refresh] = useState(0);
  useEffect(() => {
    let alive = true;
    setPending(true);
    setData(null);
    setError("");
    void rpc.call(method, JSON.parse(key)).then(
      (value) => {
        if (alive) setData(value);
      },
      (cause) => {
        if (alive)
          setError(String(cause instanceof Error ? cause.message : cause));
      }
    ).finally(() => {
      if (alive) setPending(false);
    });
    return () => {
      alive = false;
    };
  }, [rpc, method, key, revision]);
  const reload = useCallback(() => refresh((x) => x + 1), []);
  return { data, error, pending, reload };
}
function Notice({
  pending,
  error
}) {
  return /* @__PURE__ */ jsxs(Fragment2, { children: [
    pending && /* @__PURE__ */ jsx("p", { role: "status", children: "Loading\u2026" }),
    error && /* @__PURE__ */ jsx("p", { role: "alert", children: error })
  ] });
}

// quote.ts
function quoteSection(path, section) {
  const body = section.rows ? section.rows.map((r) => r.join("	")).join("\n") : (section.paragraphs ?? []).join("\n");
  return `${path} \u2014 ${section.title}
${body}`.slice(0, 12e3);
}

// app.tsx
function Viewer({ path, source }) {
  const { data, error, pending, reload } = useData("preview", {
    path,
    source
  });
  const [index, setIndex] = useState(0);
  const [message, setMessage] = useState("");
  const composer = useComposer();
  const section = data?.sections[index];
  usePage("Office Viewer", path);
  return /* @__PURE__ */ jsxs(Frame, { title: path.split(/[\\/]/).pop() || "Office Viewer", children: [
    /* @__PURE__ */ jsx(Notice, { pending, error }),
    /* @__PURE__ */ jsx("button", { style: control, onClick: reload, children: "Reload document" }),
    data && /* @__PURE__ */ jsxs(Fragment2, { children: [
      /* @__PURE__ */ jsxs("p", { children: [
        data.kind,
        " \xB7 content preview"
      ] }),
      /* @__PURE__ */ jsx("p", { style: { opacity: 0.7 }, children: "Text and cell values. Images, layout, charts and formula calculation are unavailable." }),
      data.limited && /* @__PURE__ */ jsx("p", { role: "status", children: "This preview is limited. Showing up to 1,000 paragraphs, 100 slides, or 30 sheets (200 rows \xD7 50 columns)." }),
      /* @__PURE__ */ jsxs("label", { children: [
        "Section",
        " ",
        /* @__PURE__ */ jsx(
          "select",
          {
            style: control,
            "aria-label": "Section",
            value: index,
            onChange: (e) => {
              setIndex(Number(e.target.value));
              setMessage("");
            },
            children: data.sections.map((s, i) => /* @__PURE__ */ jsx("option", { value: i, children: s.title }, i))
          }
        )
      ] }),
      section && /* @__PURE__ */ jsxs(Fragment2, { children: [
        /* @__PURE__ */ jsx("h3", { children: section.title }),
        section.rows ? /* @__PURE__ */ jsx("div", { style: { overflowX: "auto" }, children: /* @__PURE__ */ jsx("table", { style: { borderCollapse: "collapse", fontSize: 13 }, children: /* @__PURE__ */ jsx("tbody", { children: section.rows.map((row, i) => /* @__PURE__ */ jsxs("tr", { children: [
          /* @__PURE__ */ jsx("th", { style: { padding: 8, border: "1px solid #666" }, children: section.rowNumbers?.[i] ?? i + 1 }),
          row.map((cell, j) => /* @__PURE__ */ jsx(
            "td",
            {
              style: {
                padding: 8,
                border: "1px solid #666",
                maxWidth: 420,
                whiteSpace: "pre-wrap",
                overflowWrap: "anywhere"
              },
              children: cell
            },
            j
          ))
        ] }, i)) }) }) }) : section.paragraphs?.map((p, i) => /* @__PURE__ */ jsx(
          "p",
          {
            style: { whiteSpace: "pre-wrap", overflowWrap: "anywhere" },
            children: p
          },
          i
        )),
        (section.rows?.length ?? section.paragraphs?.length ?? 0) === 0 && /* @__PURE__ */ jsx("p", { children: "No readable text in this section." }),
        /* @__PURE__ */ jsx(
          "button",
          {
            style: { ...control, marginTop: 16 },
            onClick: () => {
              try {
                composer.addQuote(quoteSection(path, section));
                composer.focus();
                setMessage("Quote added to the conversation draft");
              } catch (cause) {
                setMessage(
                  String(cause instanceof Error ? cause.message : cause)
                );
              }
            },
            children: "Quote section"
          }
        ),
        message && /* @__PURE__ */ jsx("p", { role: "status", children: message })
      ] })
    ] })
  ] });
}
function Overview() {
  usePage("Office Viewer");
  return /* @__PURE__ */ jsxs(Frame, { title: "Office Viewer", children: [
    /* @__PURE__ */ jsx("p", { children: "Open a .docx, .xlsx or .pptx file from a project or thread attachment to read it beside the conversation." }),
    /* @__PURE__ */ jsx("p", { children: "Choose a sheet or slide and use \u201CQuote section\u201D to add its text to your draft." }),
    /* @__PURE__ */ jsx("p", { children: "Read-only content previews support files up to 8 MiB. Password-protected, macro and legacy binary formats are unsupported." })
  ] });
}
var app_default = definePluginApp((app) => {
  app.slots.fileOpener({
    id: "office",
    title: "Office Viewer",
    extensions: ["docx", "xlsx", "pptx"],
    component: Viewer
  });
  app.slots.navPanel({
    id: "main",
    title: "Office Viewer",
    icon: "FileText",
    placement: "extensions",
    component: Overview
  });
});
export {
  Overview,
  Viewer,
  app_default as default
};
