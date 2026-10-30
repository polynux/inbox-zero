import { beforeEach, describe, expect, it, vi } from "vitest";

const mockEnv = vi.hoisted(() => ({ IMAP_ALLOW_PRIVATE_HOSTS: false }));

vi.mock("@/env", () => ({ env: mockEnv }));

const ImapFlowMock = vi.hoisted(() => vi.fn());

vi.mock("imapflow", () => ({ ImapFlow: ImapFlowMock }));

import { createImapConnection, withImapConnection } from "@/utils/imap/client";

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

describe("createImapConnection STARTTLS", () => {
  beforeEach(() => {
    ImapFlowMock.mockClear();
  });

  it("requires STARTTLS when the IMAP security is starttls", () => {
    createImapConnection(baseConfig({ imapSecurity: "starttls" }));

    expect(ImapFlowMock).toHaveBeenCalledWith(
      expect.objectContaining({ secure: false, doSTARTTLS: true }),
    );
  });

  it("does not opportunistically upgrade a plaintext connection", () => {
    createImapConnection(baseConfig({ imapSecurity: "none" }));

    expect(ImapFlowMock).toHaveBeenCalledWith(
      expect.objectContaining({ secure: false, doSTARTTLS: false }),
    );
  });

  it("uses direct TLS without STARTTLS", () => {
    createImapConnection(baseConfig({ imapSecurity: "tls" }));

    expect(ImapFlowMock).toHaveBeenCalledWith(
      expect.objectContaining({ secure: true, doSTARTTLS: false }),
    );
  });
});

describe("withImapConnection", () => {
  function mockClient(overrides: Record<string, unknown> = {}) {
    const client = {
      connect: vi.fn(async () => {}),
      logout: vi.fn(async () => {}),
      ...overrides,
    };
    ImapFlowMock.mockImplementation(function (this: unknown) {
      Object.assign(this as object, client);
    });
    return client;
  }

  beforeEach(() => {
    ImapFlowMock.mockClear();
    vi.useRealTimers();
  });

  it("leaves no connect timeout pending after a successful call", async () => {
    vi.useFakeTimers();
    const client = mockClient();

    await withImapConnection(baseConfig(), async () => {});

    expect(client.logout).toHaveBeenCalledOnce();
    expect(vi.getTimerCount()).toBe(0);
  });

  it("clears the connect timeout after a failed call", async () => {
    vi.useFakeTimers();
    mockClient({
      connect: vi.fn(async () => {
        throw new Error("auth failed");
      }),
    });

    await expect(
      withImapConnection(baseConfig(), async () => {}),
    ).rejects.toThrow("auth failed");

    expect(vi.getTimerCount()).toBe(0);
  });

  it("rejects when the connection takes too long", async () => {
    vi.useFakeTimers();
    mockClient({ connect: vi.fn(() => new Promise(() => {})) });

    const pending = withImapConnection(baseConfig(), async () => {});
    const assertion = expect(pending).rejects.toThrow(
      "IMAP connection timeout",
    );
    await vi.advanceTimersByTimeAsync(30_001);

    await assertion;
  });
});
