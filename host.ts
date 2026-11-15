import type { ZccPluginApi } from "@zana-ai/zcc-plugin-sdk/server";
import type { PluginFileOpenerSource } from "@zana-ai/zcc-plugin-sdk/app";
type Sdk = ZccPluginApi["sdk"];
type Thread = NonNullable<Awaited<ReturnType<Sdk["threads"]["get"]>>>;
/** Structurally checked additions present in the 2.6.1 host source SDK. */
export type HostApi = Omit<ZccPluginApi, "sdk"> & {
  sdk: Sdk & {
    threads: Sdk["threads"] & {
      list(args: { limit: number; archived: boolean }): Promise<Thread[]>;
    };
    files: Sdk["files"] & {
      readProject(args: {
        path: string;
        source: PluginFileOpenerSource;
      }): Promise<{
        content: string;
        contentEncoding: "utf8" | "base64";
        sizeBytes: number;
        sha256: string;
      }>;
    };
  };
};
