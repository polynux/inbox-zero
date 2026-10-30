import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { findDraftsFolder } from "@/utils/imap/folder";
import { saveDraft } from "@/utils/imap/draft";

vi.mock("@/utils/imap/folder", () => ({
  findDraftsFolder: vi.fn(async () => "Drafts"),
}));

function createClient() {
  const append = vi.fn(async () => ({ uid: 42 }));
  const mailboxCreate = vi.fn(async () => undefined);
  const client = { append, mailboxCreate } as unknown as ImapFlow;
  return { client, append, mailboxCreate };
}

function appendedHeaderLines(append: {
  mock: { calls: Array<[string, Buffer]> };
}) {
  const raw = String(append.mock.calls[0][1]);
  return raw.split("\r\n\r\n")[0].split("\r\n");
}

describe("saveDraft", () => {
  let client: ImapFlow;
  let append: ReturnType<typeof createClient>["append"];
  let mailboxCreate: ReturnType<typeof createClient>["mailboxCreate"];

  beforeEach(() => {
    ({ client, append, mailboxCreate } = createClient());
    vi.mocked(findDraftsFolder).mockResolvedValue("Drafts");
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

  it("creates the Drafts folder when the server lists none", async () => {
    vi.mocked(findDraftsFolder)
      .mockResolvedValueOnce(null)
      .mockResolvedValueOnce("Drafts");

    await saveDraft(client, {
      from: "me@example.com",
      to: "to@example.com",
      subject: "Hello",
      text: "body",
    });

    expect(mailboxCreate).toHaveBeenCalledWith("Drafts");
    expect(append).toHaveBeenCalled();
  });

  it("throws a clear error when the Drafts folder cannot be created", async () => {
    vi.mocked(findDraftsFolder).mockResolvedValue(null);
    mailboxCreate.mockRejectedValueOnce(new Error("NO cannot create"));

    await expect(
      saveDraft(client, {
        from: "me@example.com",
        to: "to@example.com",
        subject: "Hello",
        text: "body",
      }),
    ).rejects.toThrow("No Drafts folder found");
    expect(append).not.toHaveBeenCalled();
  });
});
