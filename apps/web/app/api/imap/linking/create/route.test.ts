import { NextRequest } from "next/server";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { Prisma } from "@/generated/prisma/client";
import prisma from "@/utils/__mocks__/prisma";

vi.mock("@/utils/prisma");
vi.mock("@/utils/middleware", async () => {
  const { createWithAuthTestMiddleware } = await vi.importActual<
    typeof import("@/__tests__/helpers")
  >("@/__tests__/helpers");

  return createWithAuthTestMiddleware();
});
vi.mock("@/utils/encryption", () => ({
  encryptToken: vi.fn(() => "encrypted-password"),
}));
vi.mock("@/utils/imap/client", () => ({
  createImapConnection: vi.fn(() => ({
    connect: vi.fn(async () => undefined),
    mailboxOpen: vi.fn(async () => ({ uidNext: 42 })),
    logout: vi.fn(async () => undefined),
  })),
}));

const mockEnv = vi.hoisted(() => ({
  IMAP_ALLOW_PRIVATE_HOSTS: false,
}));

vi.mock("@/env", () => ({
  env: mockEnv,
}));

import { POST } from "./route";

const validBody = {
  email: "user@example.com",
  imapHost: "1.1.1.1",
  imapPort: 993,
  imapSecurity: "tls",
  smtpHost: "1.1.1.1",
  smtpPort: 587,
  smtpSecurity: "starttls",
  username: "user@example.com",
  password: "password",
};

function postLinkingCreate(body: Record<string, unknown>) {
  return POST(
    new NextRequest("http://localhost/api/imap/linking/create", {
      method: "POST",
      body: JSON.stringify(body),
    }),
  );
}

describe("POST /api/imap/linking/create", () => {
  beforeEach(() => {
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = false;
    vi.clearAllMocks();
  });

  it("links an account for public hosts", async () => {
    prisma.account.create.mockResolvedValue({
      emailAccount: { id: "email-account-1", email: "user@example.com" },
    } as never);

    const response = await postLinkingCreate(validBody);

    expect(response.status).toBe(200);
    expect(await response.json()).toEqual({
      emailAccountId: "email-account-1",
      email: "user@example.com",
    });
    expect(prisma.account.create).toHaveBeenCalledWith(
      expect.objectContaining({
        data: expect.objectContaining({
          imapCredential: expect.objectContaining({
            create: expect.objectContaining({ lastSeenUid: 41 }),
          }),
        }),
      }),
    );
  });

  it("links without a cursor seed when the probe cannot connect", async () => {
    const { createImapConnection } = await import("@/utils/imap/client");
    vi.mocked(createImapConnection).mockImplementationOnce(
      () =>
        ({
          connect: vi.fn(async () => {
            throw new Error("connection refused");
          }),
          logout: vi.fn(async () => undefined),
        }) as never,
    );
    prisma.account.create.mockResolvedValue({
      emailAccount: { id: "email-account-1", email: "user@example.com" },
    } as never);

    const response = await postLinkingCreate(validBody);

    expect(response.status).toBe(200);
    expect(prisma.account.create).toHaveBeenCalledWith(
      expect.objectContaining({
        data: expect.objectContaining({
          imapCredential: expect.objectContaining({
            create: expect.not.objectContaining({
              lastSeenUid: expect.anything(),
            }),
          }),
        }),
      }),
    );
  });

  it("rejects a metadata IMAP host without storing credentials", async () => {
    const response = await postLinkingCreate({
      ...validBody,
      imapHost: "169.254.169.254",
    });

    expect(response.status).toBe(400);
    expect(prisma.account.create).not.toHaveBeenCalled();
  });

  it("rejects a loopback SMTP host without storing credentials", async () => {
    const response = await postLinkingCreate({
      ...validBody,
      smtpHost: "localhost",
    });

    expect(response.status).toBe(400);
    expect(prisma.account.create).not.toHaveBeenCalled();
  });

  it("stores a loopback host when IMAP_ALLOW_PRIVATE_HOSTS is enabled", async () => {
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = true;
    prisma.account.create.mockResolvedValue({
      emailAccount: { id: "email-account-1", email: "user@example.com" },
    } as never);

    const response = await postLinkingCreate({
      ...validBody,
      imapHost: "127.0.0.1",
      smtpHost: "127.0.0.1",
      imapPort: 1143,
      smtpPort: 1025,
      imapSecurity: "none",
      smtpSecurity: "none",
    });

    expect(response.status).toBe(200);
    expect(prisma.account.create).toHaveBeenCalledOnce();
  });

  it("returns 409 when a concurrent request links the same email first", async () => {
    prisma.account.create.mockRejectedValue(
      new Prisma.PrismaClientKnownRequestError(
        "Unique constraint failed on the fields: (`providerAccountId`)",
        { code: "P2002", clientVersion: "test" },
      ),
    );

    const response = await postLinkingCreate(validBody);

    expect(response.status).toBe(409);
    expect(await response.json()).toEqual({
      error: "This email is already linked to an account",
    });
  });
});
