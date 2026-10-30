import { NextRequest } from "next/server";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { pollAllImapAccounts } from "@/utils/imap/poll";

vi.mock("@/utils/middleware", async () => {
  const { createWithErrorTestMiddleware } = await vi.importActual<
    typeof import("@/__tests__/helpers")
  >("@/__tests__/helpers");

  return createWithErrorTestMiddleware();
});
vi.mock("@/utils/imap/poll", () => ({
  pollAllImapAccounts: vi.fn(),
}));

const mockEnv = vi.hoisted(() => ({
  CRON_SECRET: undefined as string | undefined,
}));

vi.mock("@/env", () => ({
  env: mockEnv,
}));

import { POST } from "./route";

function postPoll(authorization?: string) {
  return POST(
    new NextRequest("http://localhost/api/imap/poll", {
      method: "POST",
      headers: authorization ? { authorization } : {},
    }),
  );
}

describe("POST /api/imap/poll", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockEnv.CRON_SECRET = "cron-secret";
    vi.mocked(pollAllImapAccounts).mockResolvedValue([]);
  });

  it("rejects requests when CRON_SECRET is not configured", async () => {
    mockEnv.CRON_SECRET = undefined;

    const response = await postPoll();

    expect(response.status).toBe(401);
    expect(pollAllImapAccounts).not.toHaveBeenCalled();
  });

  it("rejects requests with the wrong secret", async () => {
    const response = await postPoll("Bearer wrong-secret");

    expect(response.status).toBe(401);
    expect(pollAllImapAccounts).not.toHaveBeenCalled();
  });

  it("polls accounts when the secret matches", async () => {
    const response = await postPoll("Bearer cron-secret");

    expect(response.status).toBe(200);
    expect(await response.json()).toEqual({
      polled: 0,
      newMessages: 0,
      results: [],
    });
    expect(pollAllImapAccounts).toHaveBeenCalledOnce();
  });
});
