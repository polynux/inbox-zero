import { NextRequest } from "next/server";
import { beforeEach, describe, expect, it, vi } from "vitest";
import prisma from "@/utils/__mocks__/prisma";

vi.mock("@/utils/prisma");
vi.mock("@/utils/middleware", async () => {
  const { createWithAuthTestMiddleware } = await vi.importActual<
    typeof import("@/__tests__/helpers")
  >("@/__tests__/helpers");

  return createWithAuthTestMiddleware();
});

import { DELETE } from "./route";

function deleteImapAccount(emailAccountId = "email-account-1") {
  return DELETE(
    new NextRequest("http://localhost/api/imap/linking/delete", {
      method: "DELETE",
      body: JSON.stringify({ emailAccountId }),
    }),
  );
}

describe("DELETE /api/imap/linking/delete", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("only looks up IMAP email accounts", async () => {
    prisma.emailAccount.findFirst.mockResolvedValue({
      accountId: "account-1",
    } as never);
    prisma.account.delete.mockResolvedValue({ id: "account-1" } as never);

    const response = await deleteImapAccount();

    expect(prisma.emailAccount.findFirst).toHaveBeenCalledWith({
      where: {
        id: "email-account-1",
        userId: "user-1",
        account: { provider: "imap" },
      },
      select: { accountId: true },
    });
    expect(prisma.account.delete).toHaveBeenCalledWith({
      where: { id: "account-1" },
    });
    expect(response.status).toBe(200);
  });

  it("reports 404 and deletes nothing for a non-IMAP account", async () => {
    prisma.emailAccount.findFirst.mockResolvedValue(null);

    const response = await deleteImapAccount();

    expect(response.status).toBe(404);
    expect(prisma.account.delete).not.toHaveBeenCalled();
  });
});
