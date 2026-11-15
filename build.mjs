import { build } from "esbuild";
import { readFile, writeFile } from "node:fs/promises";
const pkg = JSON.parse(await readFile("package.json", "utf8"));
const shims = {
  react: [
    "react",
    [
      "useState",
      "useEffect",
      "useCallback",
      "useRef",
      "createElement",
      "Fragment",
    ],
  ],
  "react/jsx-runtime": ["jsxRuntime", ["jsx", "jsxs", "Fragment"]],
  "@zana-ai/zcc-plugin-sdk/app": [
    "pluginSdkApp",
    ["definePluginApp", "useRpc", "useComposer", "useZccNavigate"],
  ],
};
const runtime = {
  name: "host-runtime",
  setup(b) {
    b.onResolve(
      { filter: /^(react|react\/jsx-runtime|@zana-ai\/zcc-plugin-sdk\/app)$/ },
      (a) => ({ path: a.path, namespace: "host" }),
    );
    b.onLoad({ filter: /.*/, namespace: "host" }, (a) => {
      const [slot, names] = shims[a.path];
      return {
        contents: `const m=globalThis.__bbPluginRuntime?.${slot};if(!m)throw new Error('Open this plugin in Zana');export default m;export const {${names.join(",")}}=m;`,
      };
    });
  },
};
await build({
  entryPoints: ["app.tsx"],
  outfile: "app.js",
  bundle: true,
  format: "esm",
  platform: "browser",
  target: "es2022",
  jsx: "automatic",
  plugins: [runtime],
  define: { "process.env.NODE_ENV": '"production"' },
  legalComments: "eof",
});
await build({
  entryPoints: ["server.ts"],
  outfile: "server.mjs",
  bundle: true,
  format: "esm",
  platform: "node",
  target: "node22",
  legalComments: "eof",
});
if (pkg.name.endsWith("office-viewer"))
  await build({
    entryPoints: ["worker.ts"],
    outfile: "worker.mjs",
    bundle: true,
    format: "esm",
    platform: "node",
    target: "node22",
    legalComments: "eof",
  });
const meta = {
  sdkMajor: 1,
  sdkVersion: "0.1.0",
  artifactFormatVersion: 1,
  pluginId: pkg.name.split("/").at(-1),
  pluginVersion: pkg.version,
  builtWith: { zccVersion: "2.6.1", pluginSdkVersion: "0.1.0" },
};
for (const name of ["app.meta.json", "server.meta.json"])
  await writeFile(name, JSON.stringify(meta, null, 2) + "\n");
