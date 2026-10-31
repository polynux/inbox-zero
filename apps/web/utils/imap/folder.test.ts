import type { ImapFlow, ListResponse } from "imapflow";
import { describe, expect, it, vi } from "vitest";
import {
  findAllMailFolder,
  findArchiveFolder,
  findDraftsFolder,
  findJunkFolder,
  findSentFolder,
} from "@/utils/imap/folder";

function mailbox(path: string, specialUse?: string): ListResponse {
  return { path, name: path, specialUse } as unknown as ListResponse;
}

function mockClient(mailboxes: ListResponse[]) {
  const mailboxCreate = vi.fn().mockResolvedValue(undefined);
  const client = {
    list: vi.fn().mockResolvedValue(mailboxes),
    mailboxCreate,
  } as unknown as ImapFlow;
  return { client, mailboxCreate };
}

describe("findAllMailFolder", () => {
  it("returns the special-use all-mail mailbox", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("All Mail", "\\All"),
    ]);
    await expect(findAllMailFolder(client)).resolves.toBe("All Mail");
  });

  it("detects Bridge's All Mail by exact name when no mailbox has \\All", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("All Mail")]);
    await expect(findAllMailFolder(client)).resolves.toBe("All Mail");
  });

  it("accepts the Gmail-style All Mail path without a \\All flag", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("[Gmail]/All Mail"),
    ]);
    await expect(findAllMailFolder(client)).resolves.toBe("[Gmail]/All Mail");
  });

  it("does not match an All Mail name when a \\All mailbox exists", async () => {
    const { client } = mockClient([
      mailbox("Tout", "\\All"),
      mailbox("All Mail"),
    ]);
    await expect(findAllMailFolder(client)).resolves.toBe("Tout");
  });

  it("returns null when no all-mail mailbox exists", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Archive")]);
    await expect(findAllMailFolder(client)).resolves.toBeNull();
  });

  it("ignores a Labels child that only mentions All Mail", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("Labels/All Mail backup"),
    ]);
    await expect(findAllMailFolder(client)).resolves.toBeNull();
  });
});

describe("findSentFolder", () => {
  it("returns the special-use sent mailbox", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("Gesendet", "\\Sent"),
    ]);
    await expect(findSentFolder(client)).resolves.toBe("Gesendet");
  });

  it("returns null when the LIST result has no sent mailbox", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Drafts")]);
    await expect(findSentFolder(client)).resolves.toBeNull();
  });
});

describe("findDraftsFolder", () => {
  it("returns the special-use drafts mailbox", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("Entwürfe", "\\Drafts"),
    ]);
    await expect(findDraftsFolder(client)).resolves.toBe("Entwürfe");
  });

  it("returns null when the LIST result has no drafts mailbox", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Sent")]);
    await expect(findDraftsFolder(client)).resolves.toBeNull();
  });
});

describe("findArchiveFolder", () => {
  it("prefers the special-use archive mailbox", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("All Mail", "\\All"),
      mailbox("Archives", "\\Archive"),
    ]);
    await expect(findArchiveFolder(client)).resolves.toBe("Archives");
  });

  it("creates Archive when the LIST result lists none", async () => {
    const { client, mailboxCreate } = mockClient([mailbox("INBOX")]);
    await expect(findArchiveFolder(client)).resolves.toBe("Archive");
    expect(mailboxCreate).toHaveBeenCalledWith("Archive");
  });
});

describe("findJunkFolder", () => {
  it("returns the special-use junk mailbox", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("Junk Mail", "\\Junk"),
    ]);
    await expect(findJunkFolder(client)).resolves.toBe("Junk Mail");
  });

  it("returns a folder whose name matches a known junk name", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Spam")]);
    await expect(findJunkFolder(client)).resolves.toBe("Spam");
  });

  it("returns null when the LIST result has no junk mailbox", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Archive")]);
    await expect(findJunkFolder(client)).resolves.toBeNull();
  });
});
