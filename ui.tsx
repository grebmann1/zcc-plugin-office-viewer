import { useCallback, useEffect, useState } from "react";
import { useRpc, useZccNavigate } from "@zana-ai/zcc-plugin-sdk/app";
import type { CSSProperties, ReactNode } from "react";
export const box: CSSProperties = {
  height: "100%",
  overflow: "auto",
  padding: 24,
  boxSizing: "border-box",
  color: "var(--text-primary,inherit)",
  fontFamily: "inherit",
};
export const control: CSSProperties = {
  background: "var(--bg-elevated,#27292d)",
  color: "inherit",
  border: "1px solid var(--border-default,#555)",
  borderRadius: 7,
  padding: "8px 12px",
  font: "inherit",
};
export function Frame({
  title,
  children,
}: {
  title: string;
  children: ReactNode;
}) {
  return (
    <main style={box}>
      <h2 style={{ marginTop: 0 }}>{title}</h2>
      {children}
    </main>
  );
}
export function usePage(title: string, resourceId?: string) {
  const nav = useZccNavigate();
  useEffect(() => {
    nav.setPageContext?.({
      title,
      resourceId,
      view: resourceId ? "detail" : "list",
    });
    return () => nav.setPageContext?.(null);
  }, [nav.setPageContext, title, resourceId]);
}
export function useData<T>(method: string, args: unknown) {
  const rpc = useRpc();
  const key = JSON.stringify(args);
  const [data, setData] = useState<T | null>(null);
  const [error, setError] = useState("");
  const [pending, setPending] = useState(false);
  const [revision, refresh] = useState(0);
  useEffect(() => {
    let alive = true;
    setPending(true);
    setData(null);
    setError("");
    void rpc
      .call(method, JSON.parse(key))
      .then(
        (value) => {
          if (alive) setData(value as T);
        },
        (cause) => {
          if (alive)
            setError(String(cause instanceof Error ? cause.message : cause));
        },
      )
      .finally(() => {
        if (alive) setPending(false);
      });
    return () => {
      alive = false;
    };
  }, [rpc, method, key, revision]);
  const reload = useCallback(() => refresh((x) => x + 1), []);
  return { data, error, pending, reload };
}
export function Notice({
  pending,
  error,
}: {
  pending: boolean;
  error: string;
}) {
  return (
    <>
      {pending && <p role="status">Loading…</p>}
      {error && <p role="alert">{error}</p>}
    </>
  );
}
