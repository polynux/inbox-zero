import { beforeEach, describe, expect, it, vi } from "vitest";

const mockEnv = vi.hoisted(() => ({ IMAP_ALLOW_PRIVATE_HOSTS: false }));
vi.mock("@/env", () => ({ env: mockEnv }));

const createTransport = vi.hoisted(() =>
  vi.fn(() => ({ verify: vi.fn(async () => true), close: vi.fn() })),
);
vi.mock("nodemailer", () => ({ default: { createTransport } }));

const validateHost = vi.hoisted(() => vi.fn(async () => ({ valid: true })));
vi.mock("@/utils/imap/host-validation", () => ({
  assertImapHostAllowed: validateHost,
}));

import { sendSmtpEmail, testSmtpConnection } from "@/utils/imap/mail";

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
  } as Parameters<typeof testSmtpConnection>[0];
}

describe("testSmtpConnection", () => {
  beforeEach(() => {
    createTransport.mockClear();
    validateHost.mockClear();
    validateHost.mockResolvedValue({ valid: true });
  });

  it("re-validates the stored SMTP host before connecting", async () => {
    await testSmtpConnection(baseConfig());

    expect(validateHost).toHaveBeenCalledWith(
      "127.0.0.1",
      expect.objectContaining({ allowPrivateHosts: false }),
    );
  });

  it("refuses to send when the stored SMTP host no longer validates", async () => {
    validateHost.mockRejectedValue(new Error("Host is not allowed") as never);

    await expect(
      sendSmtpEmail(baseConfig(), {
        to: "you@example.com",
        subject: "Hi",
        text: "Hello",
      }),
    ).rejects.toThrow("Host is not allowed");
    expect(createTransport).not.toHaveBeenCalled();
  });

  it("requires STARTTLS when the SMTP security is starttls", async () => {
    await expect(testSmtpConnection(baseConfig())).resolves.toBe(true);

    expect(createTransport).toHaveBeenCalledWith(
      expect.objectContaining({ secure: false, requireTLS: true }),
    );
  });

  it("uses direct TLS without STARTTLS", async () => {
    await testSmtpConnection(baseConfig({ smtpSecurity: "tls" }));

    expect(createTransport).toHaveBeenCalledWith(
      expect.objectContaining({ secure: true, requireTLS: false }),
    );
  });
});
