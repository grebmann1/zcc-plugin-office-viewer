// @vitest-environment node
import { EventEmitter } from "node:events";
import { afterEach, expect, it, vi } from "vitest";
const state = vi.hoisted(() => ({
  workers: [] as Array<EventEmitter & { terminate: ReturnType<typeof vi.fn> }>,
}));
vi.mock("node:worker_threads", async () => {
  const { EventEmitter } = await import("node:events");
  class Worker extends EventEmitter {
    terminate = vi.fn(async () => 0);
    constructor() {
      super();
      state.workers.push(this);
    }
  }
  return { Worker, default: { Worker } };
});
import plugin from "./server.js";
import type { HostApi } from "./host.js";
const source = {
  kind: "workspace",
  projectId: "p",
  environmentId: null,
  threadId: null,
};
function harness() {
  const methods = new Map<string, (v: unknown) => Promise<unknown>>();
  let dispose!: () => void;
  const read = vi.fn(async () => ({
    content: "UEs=",
    contentEncoding: "base64",
    sizeBytes: 2,
  }));
  plugin({
    rpc: {
      method: (n: string, h: (v: unknown) => Promise<unknown>) =>
        methods.set(n, h),
    },
    onDispose: (h: () => void) => {
      dispose = h;
    },
    sdk: { files: { readProject: read } },
  } as unknown as HostApi);
  return { call: methods.get("preview")!, read, dispose: () => dispose() };
}
afterEach(() => {
  vi.useRealTimers();
  state.workers.length = 0;
});
it("reads the host-authorized source and cleans up a successful worker", async () => {
  const h = harness();
  const pending = h.call({ path: "report.docx", source });
  await vi.waitFor(() => expect(state.workers).toHaveLength(1));
  const worker = state.workers[0];
  worker.emit("message", {
    ok: true,
    value: { kind: "Word", sections: [], limited: false },
  });
  await expect(pending).resolves.toMatchObject({ kind: "Word" });
  expect(h.read).toHaveBeenCalledWith({ path: "report.docx", source });
  expect(worker.terminate).toHaveBeenCalledOnce();
  worker.emit("exit", 0);
  h.dispose();
});
it("rejects invalid file identities and unsupported formats before reading", async () => {
  const h = harness();
  for (const arg of [
    null,
    { path: "" },
    { path: "x".repeat(4097) },
    { path: "x.pdf", source },
    { path: "x.docx", source: { kind: "arbitrary" } },
    { path: "x.docx", source: { ...source, projectId: 2 } },
  ])
    await expect(h.call(arg)).rejects.toThrow();
  expect(h.read).not.toHaveBeenCalled();
  h.dispose();
  await expect(h.call({ path: "a.docx", source })).rejects.toThrow("stopped");
});
it("rejects oversize/text content and handles host read failures", async () => {
  const h = harness();
  h.read.mockResolvedValueOnce({
    content: "",
    contentEncoding: "base64",
    sizeBytes: 9 * 1024 * 1024,
  });
  await expect(h.call({ path: "a.docx", source })).rejects.toThrow("8 MiB");
  h.read.mockResolvedValueOnce({
    content: "a".repeat(12 * 1024 * 1024 + 1),
    contentEncoding: "base64",
    sizeBytes: 1,
  });
  await expect(h.call({ path: "a.docx", source })).rejects.toThrow("8 MiB");
  h.read.mockResolvedValueOnce({
    content: "not binary",
    contentEncoding: "utf8",
    sizeBytes: 10,
  });
  await expect(h.call({ path: "a.docx", source })).rejects.toThrow("binary");
  h.read.mockRejectedValueOnce(new Error("Access denied"));
  await expect(h.call({ path: "a.docx", source })).rejects.toThrow(
    "Access denied",
  );
});
it("handles worker parser failure, crash and early exit", async () => {
  for (const event of ["message", "error", "exit"]) {
    const h = harness();
    const p = h.call({ path: "a.docx", source });
    const rejection = expect(p).rejects.toThrow();
    await vi.waitFor(() => expect(state.workers.length).toBeGreaterThan(0));
    const worker = state.workers.pop()!;
    if (event === "message")
      worker.emit(event, { ok: false, error: "Corrupt document" });
    else if (event === "error") worker.emit(event, new Error("Worker crashed"));
    else worker.emit(event, 1);
    await rejection;
    expect(worker.terminate).toHaveBeenCalledOnce();
    h.dispose();
  }
});
it("bounds concurrent reads, times out work and cancels on dispose", async () => {
  vi.useFakeTimers();
  const h = harness();
  const p1 = h.call({ path: "a.docx", source }),
    p2 = h.call({ path: "b.xlsx", source });
  const rejected1 = expect(p1).rejects.toThrow("timed out");
  const rejected2 = expect(p2).rejects.toThrow("timed out");
  await expect(h.call({ path: "c.pptx", source })).rejects.toThrow(
    "Two Office",
  );
  await vi.advanceTimersByTimeAsync(8001);
  await rejected1;
  await rejected2;
  vi.useRealTimers();
  const p = h.call({ path: "a.docx", source });
  const rejected = expect(p).rejects.toThrow("exited");
  await vi.waitFor(() => expect(state.workers.length).toBe(3));
  h.dispose();
  expect(state.workers[2].terminate).toHaveBeenCalled();
  state.workers[2].emit("exit", 1);
  await rejected;
});
it("cancels a host read that completes after plugin disposal", async () => {
  const h = harness();
  let resolve!: (x: unknown) => void;
  h.read.mockImplementationOnce(
    () =>
      new Promise((r) => {
        resolve = r;
      }) as never,
  );
  const p = h.call({ path: "x.docx", source });
  const rejected = expect(p).rejects.toThrow("stopped");
  h.dispose();
  resolve({ content: "UEs=", contentEncoding: "base64", sizeBytes: 2 });
  await rejected;
});
