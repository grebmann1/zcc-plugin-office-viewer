import { afterEach, beforeEach, expect, it } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import { Overview, Viewer } from "./app.tsx";
import { rpc, composer, slots } from "./test-runtime.js";
const props = {
  pluginId: "office-viewer",
  path: "Forecast.xlsx",
  source: {
    kind: "workspace" as const,
    projectId: "p",
    threadId: null,
    environmentId: null,
  },
};
afterEach(cleanup);
beforeEach(() => {
  rpc.call.mockReset();
  composer.addQuote.mockReset();
});
it("registers the Office opener without replacing PDF/media ownership", () => {
  expect(slots.fileOpener.mock.calls[0][0].extensions).toEqual([
    "docx",
    "xlsx",
    "pptx",
  ]);
  render(<Overview />);
  expect(screen.getByText(/Open a .docx/)).toBeTruthy();
});
it("renders sheets safely, switches sections and quotes to the current draft", async () => {
  rpc.call.mockResolvedValue({
    kind: "Excel",
    limited: true,
    sections: [
      { title: "Forecast", rows: [["<script>alert(1)</script>", "42"]] },
      { title: "Empty", rows: [] },
    ],
  });
  render(<Viewer {...props} />);
  await screen.findByText("<script>alert(1)</script>");
  expect(document.querySelector("script")).toBeNull();
  fireEvent.click(screen.getByText("Quote section"));
  expect(composer.addQuote).toHaveBeenCalledWith(
    "Forecast.xlsx — Forecast\n<script>alert(1)</script>\t42",
  );
  await screen.findByText("Quote added to the conversation draft");
  fireEvent.change(screen.getByLabelText("Section"), {
    target: { value: "1" },
  });
  await screen.findByText("No readable text in this section.");
  fireEvent.click(screen.getByText("Reload document"));
  expect(rpc.call).toHaveBeenCalledWith("preview", {
    path: props.path,
    source: props.source,
  });
});
it("renders paragraphs and displays quote errors", async () => {
  rpc.call.mockResolvedValue({
    kind: "Word",
    limited: false,
    sections: [{ title: "Document", paragraphs: ["Project proposal"] }],
  });
  composer.addQuote.mockImplementationOnce(() => {
    throw new Error("No draft selected");
  });
  render(<Viewer {...props} path="report.docx" />);
  await screen.findByText("Project proposal");
  fireEvent.click(screen.getByText("Quote section"));
  await screen.findByText("No draft selected");
  composer.addQuote.mockImplementationOnce(() => {
    throw "Unavailable";
  });
  fireEvent.click(screen.getByText("Quote section"));
  await screen.findByText("Unavailable");
});
it("handles no selected section and a loading failure", async () => {
  rpc.call.mockResolvedValue({
    kind: "PowerPoint",
    limited: false,
    sections: [],
  });
  const view = render(<Viewer {...props} path="" />);
  await screen.findByText("PowerPoint · content preview");
  view.unmount();
  rpc.call.mockRejectedValue(new Error("Access denied"));
  render(<Viewer {...props} />);
  expect((await screen.findByRole("alert")).textContent).toBe("Access denied");
});
