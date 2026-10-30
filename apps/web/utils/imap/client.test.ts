import { beforeEach, describe, expect, it, vi } from "vitest";

const mockEnv = vi.hoisted(() => ({ IMAP_ALLOW_PRIVATE_HOSTS: false }));

vi.mock("@/env", () => ({ env: mockEnv }));

const ImapFlowMock = vi.hoisted(() => vi.fn());

vi.mock("imapflow", () => ({ ImapFlow: ImapFlowMock }));

import { createImapConnection } from "@/utils/imap/client";

function baseConfig(overrides: Record<string, unknown> = {}) {
  return {
    email: "me@example.com",
    emailAccountId: "account-1",
    imapHost: "127.0.0.1",
    imapPort: 1143,
    imapSecurity: "starttls",
    password: "secret",
    smtpHost: "127.0.0.1",
    smtpPort: 1025,
    smtpSecurity: "starttls",
    username: "me",
    ...overrides,
  } as Parameters<typeof createImapConnection>[0];
}

describe("createImapConnection TLS verification", () => {
  beforeEach(() => {
    ImapFlowMock.mockClear();
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = false;
  });

  it("skips verification for a loopback host when private hosts are allowed", () => {
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = true;

    createImapConnection(baseConfig());

    expect(ImapFlowMock).toHaveBeenCalledWith(
      expect.objectContaining({ tls: { rejectUnauthorized: false } }),
    );
  });

  it("verifies a loopback host when private hosts are not allowed", () => {
    createImapConnection(baseConfig());

    expect(ImapFlowMock).toHaveBeenCalledWith(
      expect.objectContaining({ tls: { rejectUnauthorized: true } }),
    );
  });

  it("verifies a remote host even when private hosts are allowed", () => {
    mockEnv.IMAP_ALLOW_PRIVATE_HOSTS = true;

    createImapConnection(baseConfig({ imapHost: "imap.fastmail.com" }));

    expect(ImapFlowMock).toHaveBeenCalledWith(
      expect.objectContaining({ tls: { rejectUnauthorized: true } }),
    );
  });
});
