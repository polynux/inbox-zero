import type { ImapFlow, ListResponse } from "imapflow";
import { describe, expect, it, vi } from "vitest";
import {
  findArchiveFolder,
  findDraftsFolder,
  findSentFolder,
  findTrashFolder,
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

describe("findTrashFolder", () => {
  it("returns the special-use trash mailbox", async () => {
    const { client } = mockClient([
      mailbox("INBOX"),
      mailbox("[Gmail]/Bin", "\\Trash"),
    ]);
    await expect(findTrashFolder(client)).resolves.toBe("[Gmail]/Bin");
  });

  it("returns a folder whose name matches a known trash name", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Deleted Items")]);
    await expect(findTrashFolder(client)).resolves.toBe("Deleted Items");
  });

  it("returns null when the LIST result has no trash mailbox", async () => {
    const { client } = mockClient([mailbox("INBOX"), mailbox("Archive")]);
    await expect(findTrashFolder(client)).resolves.toBeNull();
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
