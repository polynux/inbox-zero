import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { saveDraft } from "@/utils/imap/draft";

vi.mock("@/utils/imap/folder", () => ({
  findDraftsFolder: vi.fn(async () => "Drafts"),
}));

function createClient() {
  const append = vi.fn(async () => ({ uid: 42 }));
  return { client: { append } as unknown as ImapFlow, append };
}

function appendedHeaderLines(
  append: ReturnType<typeof createClient>["append"],
) {
  const raw = String(append.mock.calls[0][1]);
  return raw.split("\r\n\r\n")[0].split("\r\n");
}

describe("saveDraft", () => {
  let client: ImapFlow;
  let append: ReturnType<typeof createClient>["append"];

  beforeEach(() => {
    ({ client, append } = createClient());
  });

  it("keeps CR, LF and NUL out of header values so they cannot add headers", async () => {
    await saveDraft(client, {
      from: "me@example.com\r\nBcc: attacker@example.com",
      to: "to@example.com\nBcc: attacker@example.com",
      subject: "Hi\r\nBcc: attacker@example.com\0",
      inReplyTo: "<root@example.com>\r\nBcc: attacker@example.com",
      references: "<root@example.com>\r\nCc: attacker@example.com",
      text: "body",
    });

    const headers = appendedHeaderLines(append);
    expect(headers.some((line) => /^(Bcc|Cc):/.test(line))).toBe(false);
    expect(headers.filter((line) => line.startsWith("To:"))).toHaveLength(1);
    expect(headers).toContain("Subject: HiBcc: attacker@example.com");
  });

  it("writes headers and body unchanged for clean input", async () => {
    await saveDraft(client, {
      from: "me@example.com",
      to: "to@example.com",
      subject: "Weekly update",
      inReplyTo: "<root@example.com>",
      references: "<root@example.com>",
      text: "Hello",
    });

    const headers = appendedHeaderLines(append);
    expect(headers).toEqual(
      expect.arrayContaining([
        "From: me@example.com",
        "To: to@example.com",
        "Subject: Weekly update",
        "In-Reply-To: <root@example.com>",
        "References: <root@example.com>",
      ]),
    );
  });
});
