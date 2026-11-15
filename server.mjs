// server.ts
import { Worker } from "node:worker_threads";
var record = (x) => x !== null && typeof x === "object" && !Array.isArray(x) ? x : {};
function plugin(zcc) {
  const workers = /* @__PURE__ */ new Set();
  let disposed = false;
  let active = 0;
  zcc.onDispose(() => {
    disposed = true;
    for (const worker of workers) void worker.terminate();
    workers.clear();
  });
  zcc.rpc.method("preview", async (input) => {
    if (disposed) throw new Error("Plugin has stopped");
    if (active >= 2)
      throw new Error(
        "Two Office previews are already loading. Try again shortly."
      );
    const args = record(input);
    const path = args.path;
    const source = record(args.source);
    if (typeof path !== "string" || path.length > 4096 || !path.trim())
      throw new Error("A document path is required");
    const extension = path.split(".").pop().toLowerCase();
    if (!["docx", "xlsx", "pptx"].includes(extension))
      throw new Error("Supported formats: docx, xlsx, pptx");
    if (!["workspace", "host", "thread-storage"].includes(String(source.kind)))
      throw new Error("A project-backed file source is required");
    for (const key of ["threadId", "environmentId", "projectId"])
      if (source[key] !== null && typeof source[key] !== "string")
        throw new Error("Invalid file source");
    active++;
    try {
      const file = await zcc.sdk.files.readProject({
        path,
        source
      });
      if (disposed) throw new Error("Plugin has stopped");
      if (file.sizeBytes > 8 * 1024 * 1024 || file.content.length > 12 * 1024 * 1024)
        throw new Error("Office preview limit is 8 MiB");
      if (file.contentEncoding !== "base64")
        throw new Error("Expected a binary Office document");
      const bytes = Buffer.from(file.content, "base64");
      return await new Promise((resolve, reject) => {
        const worker = new Worker(new URL("./worker.mjs", import.meta.url), {
          workerData: { bytes, extension },
          resourceLimits: { maxOldGenerationSizeMb: 96 }
        });
        workers.add(worker);
        let settled = false;
        const finish = (error, value) => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          workers.delete(worker);
          void worker.terminate();
          if (error) reject(error);
          else resolve(value);
        };
        const timer = setTimeout(
          () => finish(new Error("Document parsing timed out")),
          8e3
        );
        worker.once(
          "message",
          (message) => message.ok ? finish(void 0, message.value) : finish(new Error(message.error || "Cannot parse document"))
        );
        worker.once("error", (error) => finish(error));
        worker.once(
          "exit",
          () => finish(
            new Error("Document worker exited before returning a preview")
          )
        );
      });
    } finally {
      active--;
    }
  });
}
export {
  plugin as default
};
