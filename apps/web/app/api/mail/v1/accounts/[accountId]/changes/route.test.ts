import { NextRequest } from "next/server";
import { describe, expect, it, vi } from "vitest";
import { POST } from "./route";

vi.mock("server-only", () => ({}));

vi.mock("@/utils/middleware", () => ({
  withEmailProvider:
    (
      _name: string,
      handler: (
        request: NextRequest & Record<string, unknown>,
        context: { params: Promise<{ accountId: string }> },
      ) => Promise<Response>,
    ) =>
    (
      request: NextRequest,
      context: { params: Promise<{ accountId: string }> },
    ) =>
      handler(
        Object.assign(request, {
          auth: { emailAccountId: "acc-1", userId: "user-1" },
          emailProvider: { name: "imap" },
        }),
        context,
      ),
}));

describe("POST /changes", () => {
  it("answers an empty body with 400 instead of throwing", async () => {
    const response = await POST(
      new NextRequest("http://localhost/api/mail/v1/accounts/acc-1/changes", {
        method: "POST",
      }),
      { params: Promise.resolve({ accountId: "acc-1" }) } as never,
    );

    expect(response.status).toBe(400);
    await expect(response.json()).resolves.toMatchObject({
      error: { code: "invalid" },
    });
  });
});
