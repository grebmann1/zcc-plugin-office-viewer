// @vitest-environment node
import { expect, it, vi } from "vitest";
const state = vi.hoisted(() => ({
  postMessage: vi.fn(),
  parse: vi.fn(),
  data: { bytes: [1, 2], extension: "docx" },
}));
vi.mock("node:worker_threads", () => {
  const fake = {
    parentPort: { postMessage: state.postMessage },
    workerData: state.data,
  };
  return { ...fake, default: fake };
});
vi.mock("./parser.js", () => ({ parseOffice: state.parse }));
it("returns parsed content or a bounded error to the owning process", async () => {
  state.parse.mockReturnValueOnce({ kind: "Word" });
  await import("./worker.js");
  expect(state.postMessage).toHaveBeenLastCalledWith({
    ok: true,
    value: { kind: "Word" },
  });
  vi.resetModules();
  state.parse.mockImplementationOnce(() => {
    throw new Error("Invalid");
  });
  await import("./worker.js");
  expect(state.postMessage).toHaveBeenLastCalledWith({
    ok: false,
    error: "Invalid",
  });
  vi.resetModules();
  state.parse.mockImplementationOnce(() => {
    throw "Broken";
  });
  await import("./worker.js");
  expect(state.postMessage).toHaveBeenLastCalledWith({
    ok: false,
    error: "Broken",
  });
});
