import { describe, expect, it } from "vitest";
import { convertImapMessage, parseSearchQuery } from "@/utils/imap/message";
describe("parseSearchQuery", () => {
  it("parses from: query", () => {
    const result = parseSearchQuery("from:test@example.com");
    expect(result).toEqual({ from: "test@example.com" });
  });

  it("parses to: query", () => {
    const result = parseSearchQuery("to:user@example.com");
    expect(result).toEqual({ to: "user@example.com" });
  });

  it("parses subject: query", () => {
    const result = parseSearchQuery("subject:meeting");
    expect(result).toEqual({ subject: "meeting" });
  });

  it("parses is:unread", () => {
    const result = parseSearchQuery("is:unread");
    expect(result).toEqual({ unseen: true });
  });

  it("parses is:read", () => {
    const result = parseSearchQuery("is:read");
    expect(result).toEqual({ seen: true });
  });

  it("parses combined queries", () => {
    const result = parseSearchQuery("from:test@ex.com is:unread");
    expect(result).toEqual({
      and: [{ from: "test@ex.com" }, { unseen: true }],
    });
  });

  it("returns all:true for empty query", () => {
    const result = parseSearchQuery("");
    expect(result).toEqual({ all: true });
  });

  it("treats bare text as body search", () => {
    const result = parseSearchQuery("important");
    expect(result).toEqual({ body: "important" });
  });

  it("keeps a quoted field value as one term", () => {
    const result = parseSearchQuery('subject:"weekly update"');
    expect(result).toEqual({ subject: "weekly update" });
  });

  it("keeps a bare quoted phrase as one body term", () => {
    const result = parseSearchQuery('"exact phrase"');
    expect(result).toEqual({ body: "exact phrase" });
  });

  it("mixes quoted values with unquoted terms", () => {
    const result = parseSearchQuery(
      'from:alice@example.com subject:"weekly update" important',
    );
    expect(result).toEqual({
      and: [
        { from: "alice@example.com" },
        { subject: "weekly update" },
        { body: "important" },
      ],
    });
  });

  it("treats an unterminated quote as a single term", () => {
    expect(parseSearchQuery('subject:"weekly update')).toEqual({
      subject: "weekly update",
    });
    expect(parseSearchQuery('"exact phrase')).toEqual({
      body: "exact phrase",
    });
  });
});

describe("convertImapMessage attachments", () => {
  const baseMessage = {
    uid: 42,
    envelope: {
      date: new Date("2026-01-01"),
      subject: "Files",
      from: [{ address: "sender@example.com" }],
      to: [],
    },
    flags: new Set<string>(),
  };

  it("populates attachment metadata from bodyStructure without fetching content", async () => {
    const parsed = await convertImapMessage(
      {
        ...baseMessage,
        bodyStructure: {
          type: "multipart/mixed",
          childNodes: [
            { type: "text/plain", part: "1", size: 100 },
            {
              type: "application/pdf",
              part: "2",
              size: 12_345,
              disposition: "attachment",
              dispositionParameters: { filename: "report.pdf" },
            },
          ],
        },
      } as never,
      undefined,
      "INBOX",
    );

    expect(parsed?.attachments).toEqual([
      {
        attachmentId: "2",
        filename: "report.pdf",
        mimeType: "application/pdf",
        size: 12_345,
        headers: {
          "content-description": "",
          "content-disposition": "attachment",
          "content-id": "",
          "content-transfer-encoding": "",
          "content-type": "application/pdf",
        },
      },
    ]);
    expect(parsed?.hasAttachment).toBe(true);
  });

  it("treats a named part as an attachment and an inline part as inline", async () => {
    const parsed = await convertImapMessage(
      {
        ...baseMessage,
        bodyStructure: {
          type: "multipart/related",
          childNodes: [
            {
              type: "image/png",
              part: "2",
              size: 500,
              disposition: "inline",
              id: "<logo@example.com>",
              parameters: { name: "logo.png" },
            },
            {
              type: "application/zip",
              part: "3",
              size: 999,
              parameters: { name: "files.zip" },
            },
          ],
        },
      } as never,
      undefined,
      "INBOX",
    );

    expect(parsed?.attachments?.map((a) => a.filename)).toEqual(["files.zip"]);
    expect(parsed?.inline.map((a) => a.attachmentId)).toEqual(["2"]);
  });

  it("leaves messages without a bodyStructure attachment-free", async () => {
    const parsed = await convertImapMessage(
      { ...baseMessage } as never,
      undefined,
      "INBOX",
    );

    expect(parsed?.attachments).toBeUndefined();
    expect(parsed?.hasAttachment).toBeUndefined();
    expect(parsed?.inline).toEqual([]);
  });
});

describe("convertImapMessage label ids", () => {
  const baseMessage = {
    uid: 7,
    envelope: {
      date: new Date("2026-01-01"),
      subject: "Hello",
      from: [{ address: "sender@example.com" }],
      to: [],
    },
    flags: new Set<string>(),
  };

  it("maps an unseen INBOX message to UNREAD + INBOX", async () => {
    const parsed = await convertImapMessage(
      { ...baseMessage } as never,
      undefined,
      "INBOX",
    );
    expect(parsed?.labelIds).toContain("UNREAD");
    expect(parsed?.labelIds).toContain("INBOX");
    expect(parsed?.labelIds).not.toContain("\\Seen");
  });

  it("maps flags and the Sent folder to STARRED + SENT without UNREAD", async () => {
    const parsed = await convertImapMessage(
      { ...baseMessage, flags: new Set(["\\Seen", "\\Flagged"]) } as never,
      undefined,
      "Sent",
    );
    expect(parsed?.labelIds).toContain("STARRED");
    expect(parsed?.labelIds).toContain("SENT");
    expect(parsed?.labelIds).not.toContain("UNREAD");
    expect(parsed?.labelIds).not.toContain("INBOX");
  });

  it("keeps custom keywords", async () => {
    const parsed = await convertImapMessage(
      { ...baseMessage, flags: new Set(["$MailFlagBit0"]) } as never,
      undefined,
      "INBOX",
    );
    expect(parsed?.labelIds).toContain("$MailFlagBit0");
  });

  it("does not read 'Presentations' as SENT", async () => {
    const parsed = await convertImapMessage(
      { ...baseMessage } as never,
      undefined,
      "Presentations",
    );
    expect(parsed?.labelIds).not.toContain("SENT");
  });

  it("maps special-use flags exactly (no substring matching)", async () => {
    const sent = await convertImapMessage(
      { ...baseMessage } as never,
      undefined,
      "Some Folder",
      "\\Sent",
    );
    expect(sent?.labelIds).toContain("SENT");

    const plain = await convertImapMessage(
      { ...baseMessage } as never,
      undefined,
      "Some Folder",
    );
    expect(plain?.labelIds).not.toContain("SENT");
    expect(plain?.labelIds).not.toContain("TRASH");
    expect(plain?.labelIds).not.toContain("SPAM");
  });
});
