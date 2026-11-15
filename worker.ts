import { parentPort, workerData } from "node:worker_threads";
import { parseOffice } from "./parser.js";
try {
  parentPort!.postMessage({
    ok: true,
    value: parseOffice(new Uint8Array(workerData.bytes), workerData.extension),
  });
} catch (cause) {
  parentPort!.postMessage({
    ok: false,
    error: cause instanceof Error ? cause.message : String(cause),
  });
}
