import { afterEach, beforeEach, expect, it } from "vitest";
import {
  cleanup,
  fireEvent,
  render,
  screen,
  waitFor,
} from "@testing-library/react";
import { Frame, Notice, useData, usePage } from "./ui.js";
import { rpc, navigation } from "./test-runtime.js";
afterEach(cleanup);
beforeEach(() => {
  rpc.call.mockReset();
});
function Probe() {
  usePage("Resource", "one");
  const state = useData<string>("load", {});
  return (
    <Frame title="Probe">
      <Notice {...state} />
      <p>{state.data}</p>
      <button onClick={state.reload}>Refresh</button>
    </Frame>
  );
}
it("loads source data, refreshes it and clears current-page identity on unmount", async () => {
  rpc.call.mockResolvedValueOnce("First").mockResolvedValueOnce("Second");
  const view = render(<Probe />);
  await screen.findByText("First");
  fireEvent.click(screen.getByText("Refresh"));
  await screen.findByText("Second");
  expect(navigation.setPageContext).toHaveBeenCalledWith({
    title: "Resource",
    resourceId: "one",
    view: "detail",
  });
  view.unmount();
  expect(navigation.setPageContext).toHaveBeenLastCalledWith(null);
});
it("shows failures and recovers through a human refresh", async () => {
  rpc.call
    .mockRejectedValueOnce(new Error("Disconnected"))
    .mockResolvedValueOnce("Recovered");
  render(<Probe />);
  await screen.findByRole("alert");
  expect(screen.getByRole("alert").textContent).toBe("Disconnected");
  fireEvent.click(screen.getByText("Refresh"));
  await screen.findByText("Recovered");
  expect(screen.queryByRole("alert")).toBeNull();
});
it("handles non-Error rejection values", async () => {
  rpc.call.mockRejectedValue("Unavailable");
  render(<Probe />);
  expect((await screen.findByRole("alert")).textContent).toBe("Unavailable");
});
it("ignores replies arriving after navigation/unmount", async () => {
  let finish!: (value: string) => void;
  rpc.call.mockImplementation(
    () =>
      new Promise<string>((r) => {
        finish = r;
      }),
  );
  const view = render(<Probe />);
  expect(screen.getByRole("status").textContent).toBe("Loading…");
  view.unmount();
  finish("Old page");
  await waitFor(() => expect(screen.queryByText("Old page")).toBeNull());
});
