import { vi } from "vitest";
export const rpc = { call: vi.fn() };
export const composer = { addQuote: vi.fn(), focus: vi.fn() };
export const navigation = { setPageContext: vi.fn() };
export const slots = {
  navPanel: vi.fn(),
  threadPanelAction: vi.fn(),
  fileOpener: vi.fn(),
};
export const useRpc = () => rpc;
export const useComposer = () => composer;
export const useZccNavigate = () => navigation;
export const definePluginApp = (
  setup: (app: { slots: typeof slots }) => void,
) => {
  setup({ slots });
  return slots;
};
