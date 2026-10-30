import { NextRequest } from "next/server";
import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { createImapConnection } from "@/utils/imap/client";
import { testSmtpConnection } from "@/utils/imap/mail";

vi.mock("@/utils/middleware", async () => {
  const { createWithAuthTestMiddleware } = await vi.importActual<
    typeof import("@/__tests__/helpers")
  >("@/__tests__/helpers");

  return createWithAuthTestMiddleware();
});
vi.mock("@/utils/imap/client", () => ({
  createImapConnection: vi.fn(),
}));
vi.mock("@/utils/imap/mail", () => ({
  testSmtpConnection: vi.fn(),
}));

const mockEnv = vi.hoisted(() => ({
  IMAP_ALLOW_PRIVATE_HOSTS: false,
}));

vi.mock("@/env", () => ({
  env: mockEnv,
}));

import { POST } from "./route";

const validBody = {
  imapHost: "1.1.1.1",
  imapPort: 993,
  imapSecurity: "tls",
  smtpHost: "1.1.1.1",
  smtpPort: 587,
  smtpSecurity: "starttls",
  username: "user@example.com",
  password: "password",
};

function postTestConnection(body: Record<string, unknown>) {
  return POST(
    new NextRequest("http://localhost/api/imap/test-connection", {
      method: "POST",
      body: JSON.stringify(body),
    }),
  );
}

describe("POST /api/imap/test-connection", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = false;
    vi.mocked(createImapConnection).mockReturnValue({
      connect: vi.fn().mockResolvedValue(undefined),
      logout: vi.fn().mockResolvedValue(undefined),
    } as unknown as ImapFlow);
    vi.mocked(testSmtpConnection).mockResolvedValue(true);
  });

  it("rejects a loopback IMAP host without connecting", async () => {
    const response = await postTestConnection({
      ...validBody,
      imapHost: "127.0.0.1",
    });

    expect(response.status).toBe(400);
    expect(await response.json()).toEqual({
      imap: false,
      smtp: false,
      error: "Host is not allowed",
    });
    expect(createImapConnection).not.toHaveBeenCalled();
    expect(testSmtpConnection).not.toHaveBeenCalled();
  });

  it("rejects a metadata SMTP host without connecting", async () => {
    const response = await postTestConnection({
      ...validBody,
      smtpHost: "169.254.169.254",
    });

    expect(response.status).toBe(400);
    expect(createImapConnection).not.toHaveBeenCalled();
    expect(testSmtpConnection).not.toHaveBeenCalled();
  });

  it("allows loopback hosts when IMAP_ALLOW_PRIVATE_HOSTS is enabled", async () => {
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = true;

    const response = await postTestConnection({
      ...validBody,
      imapHost: "127.0.0.1",
      smtpHost: "127.0.0.1",
      imapPort: 1143,
      smtpPort: 1025,
      imapSecurity: "none",
      smtpSecurity: "none",
    });

    expect(response.status).toBe(200);
    expect(await response.json()).toEqual({
      imap: true,
      smtp: true,
      error: "",
    });
    expect(createImapConnection).toHaveBeenCalledWith(
      expect.objectContaining({ imapHost: "127.0.0.1" }),
    );
  });
});
