import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { findDraftsFolder } from "@/utils/imap/folder";
import {
  encodeImapDraftId,
  encodeReplyDraftId,
  parseImapDraftId,
  parseImapDraftIds,
  saveDraft,
} from "@/utils/imap/draft";

vi.mock("@/utils/imap/folder", () => ({
  findDraftsFolder: vi.fn(async () => "Drafts"),
}));

function createClient() {
  const append = vi.fn(async () => ({ uid: 42, uidValidity: 1 }));
  const mailboxCreate = vi.fn(async () => undefined);
  const mailboxOpen = vi.fn(async () => ({}));
  const search = vi.fn(async () => [] as number[]);
  const client = {
    append,
    mailboxCreate,
    mailboxOpen,
    search,
  } as unknown as ImapFlow;
  return { client, append, mailboxCreate, mailboxOpen, search };
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
  let search: ReturnType<typeof createClient>["search"];

  beforeEach(() => {
    ({ client, append, mailboxCreate, search } = createClient());
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
    const id = await saveDraft(client, {
      from: "me@example.com",
      to: "to@example.com",
      subject: "Weekly update",
      inReplyTo: "<root@example.com>",
      references: "<root@example.com>",
      text: "Hello",
    });

    // Draft ids are stable across updates: they name the Message-ID, not the
    // UID, so the id resolves after the same Message-ID is re-APPENDed.
    expect(id.startsWith("d.")).toBe(true);
    expect(parseImapDraftId(id)).toMatch(/@inbox-zero\.local$/);

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

  it("recovers the APPENDed UID by Message-ID when the server returns none", async () => {
    append.mockResolvedValueOnce({ uidValidity: 123 });
    search.mockImplementationOnce(async () => [77]);

    const id = await saveDraft(client, {
      from: "me@example.com",
      to: "to@example.com",
      subject: "Hello",
      text: "body",
    });

    const raw = String(append.mock.calls[0][1]);
    expect(raw).toMatch(/^Message-ID: <[^>]+@inbox-zero\.local>$/m);
    expect(search).toHaveBeenCalledOnce();
    // The id still names the generated Message-ID even though the UID came
    // from a follow-up SEARCH.
    expect(id.startsWith("d.")).toBe(true);
  });

  it("throws a clear error when APPEND returns no UID (no UIDPLUS)", async () => {
    append.mockResolvedValueOnce({ uidValidity: 123 });

    await expect(
      saveDraft(client, {
        from: "me@example.com",
        to: "to@example.com",
        subject: "Hello",
        text: "body",
      }),
    ).rejects.toThrow(/did not return a UID/i);
  });
});

describe("reply draft ids", () => {
  it("round-trips the draft and parent Message-IDs", () => {
    const id = encodeImapDraftId("<draft@example.com>", "<parent@example.com>");
    expect(parseImapDraftIds(id)).toEqual({
      messageId: "draft@example.com",
      parentId: "parent@example.com",
    });
    expect(parseImapDraftId(id)).toBe("draft@example.com");
  });

  it("keeps accepting the old single-part form", () => {
    const id = encodeImapDraftId("<draft@example.com>");
    expect(parseImapDraftIds(id)).toEqual({ messageId: "draft@example.com" });
  });

  it("attaches a parent id to an existing stable id", () => {
    const id = encodeReplyDraftId(
      encodeImapDraftId("<draft@example.com>"),
      "<parent@example.com>",
    );
    expect(parseImapDraftIds(id).parentId).toBe("parent@example.com");
  });

  it("rejects a three-part id", () => {
    expect(() => parseImapDraftIds("d.a.b.c")).toThrow(/undecodable/);
  });
});
