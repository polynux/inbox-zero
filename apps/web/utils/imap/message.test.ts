import { describe, expect, it, vi } from "vitest";
import {
  convertImapMessage,
  parseSearchQuery,
  searchImapMessages,
} from "@/utils/imap/message";
describe("parseSearchQuery", () => {
  it("parses from: query", () => {
    expect(parseSearchQuery("from:test@example.com")).toEqual({
      criteria: { from: "test@example.com" },
      hasAttachment: false,
    });
  });

  it("parses to: query", () => {
    expect(parseSearchQuery("to:user@example.com")).toEqual({
      criteria: { to: "user@example.com" },
      hasAttachment: false,
    });
  });

  it("parses subject: query", () => {
    expect(parseSearchQuery("subject:meeting")).toEqual({
      criteria: { subject: "meeting" },
      hasAttachment: false,
    });
  });

  it("parses is:unread", () => {
    expect(parseSearchQuery("is:unread")).toEqual({
      criteria: { unseen: true },
      hasAttachment: false,
    });
  });

  it("parses is:read", () => {
    expect(parseSearchQuery("is:read")).toEqual({
      criteria: { seen: true },
      hasAttachment: false,
    });
  });

  it("maps is:seen, is:starred and is:flagged to flags", () => {
    expect(parseSearchQuery("is:seen").criteria).toEqual({ seen: true });
    expect(parseSearchQuery("is:starred").criteria).toEqual({
      flagged: true,
    });
    expect(parseSearchQuery("is:flagged").criteria).toEqual({
      flagged: true,
    });
  });

  // imapflow has no `and` operator (`{and: [...]}` compiles to nothing, which
  // Bridge answers `false` to): distinct terms merge into one natively-ANDed
  // object.
  it("merges combined queries into one natively-ANDed object", () => {
    expect(parseSearchQuery("from:test@ex.com is:unread")).toEqual({
      criteria: { from: "test@ex.com", unseen: true },
      hasAttachment: false,
    });
  });

  // A repeated key cannot merge into one object, so AND without `and`: NOT
  // (NOT a OR NOT b).
  it("expresses repeated keys as De Morgan NOT/OR", () => {
    expect(parseSearchQuery("invoice receipt").criteria).toEqual({
      not: { or: [{ not: { body: "invoice" } }, { not: { body: "receipt" } }] },
    });
  });

  it("returns all:true for empty query", () => {
    expect(parseSearchQuery("")).toEqual({
      criteria: { all: true },
      hasAttachment: false,
    });
  });

  it("treats bare text as body search", () => {
    expect(parseSearchQuery("important")).toEqual({
      criteria: { body: "important" },
      hasAttachment: false,
    });
  });

  it("keeps a quoted field value as one term", () => {
    expect(parseSearchQuery('subject:"weekly update"')).toEqual({
      criteria: { subject: "weekly update" },
      hasAttachment: false,
    });
  });

  it("keeps a bare quoted phrase as one body term", () => {
    expect(parseSearchQuery('"exact phrase"')).toEqual({
      criteria: { body: "exact phrase" },
      hasAttachment: false,
    });
  });

  it("mixes quoted values with unquoted terms in one object", () => {
    expect(
      parseSearchQuery('from:alice@example.com subject:"weekly update" important')
        .criteria,
    ).toEqual({
      from: "alice@example.com",
      subject: "weekly update",
      body: "important",
    });
  });

  it("treats an unterminated quote as a single term", () => {
    expect(parseSearchQuery('subject:"weekly update').criteria).toEqual({
      subject: "weekly update",
    });
    expect(parseSearchQuery('"exact phrase').criteria).toEqual({
      body: "exact phrase",
    });
  });

  it("returns has:attachment as a client-side filter flag, never BODY text", () => {
    expect(parseSearchQuery("has:attachment invoice")).toEqual({
      criteria: { body: "invoice" },
      hasAttachment: true,
    });
    expect(parseSearchQuery("has:attachment").criteria).toEqual({
      all: true,
    });
  });

  it("drops unknown is:/has: tokens instead of searching them as body text", () => {
    expect(parseSearchQuery("is:snoozed invoice").criteria).toEqual({
      body: "invoice",
    });
    expect(parseSearchQuery("has:click invoice").criteria).toEqual({
      body: "invoice",
    });
    expect(parseSearchQuery("is:snoozed").criteria).toEqual({ all: true });
  });
});

describe("searchImapMessages", () => {
  // Live: Bridge answers `false` to queries it cannot run. That is not "no
  // hits": warn so it never reads as an empty mailbox.
  it("warns and returns [] when the server returns a non-array result", async () => {
    const warn = vi.fn();
    const client = {
      search: async () => false,
      mailbox: { path: "INBOX" },
    };

    const result = await searchImapMessages(
      client as never,
      { and: [{ from: "x" }] },
      undefined,
      { warn },
    );

    expect(result).toEqual([]);
    expect(warn).toHaveBeenCalledOnce();
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
