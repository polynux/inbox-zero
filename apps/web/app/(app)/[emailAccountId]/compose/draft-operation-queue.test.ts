import { describe, expect, it } from "vitest";
import {
  enqueueDraftOperation,
  waitForDraftOperations,
} from "./draft-operation-queue";

describe("draft-operation-queue send ordering", () => {
  it("drains an in-flight autosave before the send reads the draft id", async () => {
    const key = `send-order-${Date.now()}`;
    const order: string[] = [];
    const { promise: autosaveGate, resolve: releaseAutosave } =
      Promise.withResolvers<void>();
    const autosave = enqueueDraftOperation(key, async () => {
      await autosaveGate;
      order.push("autosave");
    });
    const send = enqueueDraftOperation(key, async () => {
      order.push("send");
    });
    releaseAutosave();
    await autosave;
    await send;
    await waitForDraftOperations(key);
    expect(order).toEqual(["autosave", "send"]);
  });
});
