import { describe, expect, it } from "vitest";
import {
  imapCredentialSchema,
  testImapConnectionSchema,
} from "@/utils/actions/imap.validation";

const baseBody = {
  email: "user@example.com",
  imapHost: "127.0.0.1",
  imapPort: 1143,
  imapSecurity: "none",
  smtpHost: "127.0.0.1",
  smtpPort: 1025,
  smtpSecurity: "none",
  username: "user@example.com",
  password: "password",
};

describe("imapCredentialSchema plaintext security", () => {
  it.each([
    ["127.0.0.1"],
    ["localhost"],
    ["[::1]"],
  ])("accepts plaintext for the loopback host %s", (host) => {
    const parsed = imapCredentialSchema.safeParse({
      ...baseBody,
      imapHost: host,
      smtpHost: host,
    });

    expect(parsed.success).toBe(true);
  });

  it.each([
    ["imapSecurity", { imapHost: "imap.fastmail.com" }],
    ["smtpSecurity", { smtpHost: "smtp.fastmail.com" }],
  ])("rejects plaintext %s on a remote host", (field, override) => {
    const parsed = imapCredentialSchema.safeParse({ ...baseBody, ...override });

    expect(parsed.success).toBe(false);
    expect(parsed.error?.issues).toEqual([
      expect.objectContaining({ path: [field] }),
    ]);
  });

  it("accepts STARTTLS on a remote host", () => {
    const parsed = imapCredentialSchema.safeParse({
      ...baseBody,
      imapHost: "imap.fastmail.com",
      imapSecurity: "starttls",
      smtpHost: "smtp.fastmail.com",
      smtpSecurity: "tls",
    });

    expect(parsed.success).toBe(true);
  });
});

describe("testImapConnectionSchema plaintext security", () => {
  it("rejects plaintext on a remote host", () => {
    const { email: _email, ...connection } = baseBody;
    const parsed = testImapConnectionSchema.safeParse({
      ...connection,
      smtpHost: "smtp.fastmail.com",
    });

    expect(parsed.success).toBe(false);
  });
});
