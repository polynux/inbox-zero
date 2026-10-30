import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ImapProvider } from "@/utils/email/imap";
import * as mailModule from "@/utils/imap/mail";
const connection = vi.hoisted(() => ({ client: null as unknown }));

vi.mock("@/utils/imap/client", () => ({
  withImapConnection: async <T>(
    _config: unknown,
    fn: (client: never) => Promise<T>,
  ): Promise<T> => fn(connection.client as never),
}));

type FakeMessage = {
  uid: number;
  date: string;
  from?: string;
  to?: string;
  messageId?: string;
  inReplyTo?: string;
};

type FakeMailbox = {
  path: string;
  messages: FakeMessage[];
  specialUse?: string;
  flags?: string[];
  listed?: boolean;
};

const config = {
  email: "me@example.com",
  emailAccountId: "account-1",
  imapHost: "127.0.0.1",
  imapPort: 1143,
  imapSecurity: "none" as const,
  password: "secret",
  smtpHost: "127.0.0.1",
  smtpPort: 1025,
  smtpSecurity: "none" as const,
  username: "me",
};

function message(
  uid: number,
  date: string,
  from?: string,
  to?: string,
  headers?: { messageId?: string; inReplyTo?: string },
): FakeMessage {
  return { uid, date, from, to, ...headers };
}

function qualifiedId(folder: string, uid: number): string {
  return `${uid}@${Buffer.from(folder, "utf8").toString("base64url")}`;
}

function mailboxesOf(fake: { client: unknown }): Record<string, FakeMessage[]> {
  const list = (
    fake.client as {
      __mailboxes?: Record<string, FakeMessage[]>;
    }
  ).__mailboxes;
  if (!list) throw new Error("fake has no mailbox store");
  return list;
}

function createFakeClient(mailboxes: FakeMailbox[]) {
  const opened: string[] = [];
  let current = mailboxes[0];
  const fetchedUids: number[] = [];
  const fetchQueries: Array<Record<string, unknown> | undefined> = [];
  const sequenceOf = (message: FakeMessage) =>
    current.messages.indexOf(message) + 1;

  function fetched(message: FakeMessage) {
    return {
      uid: message.uid,
      envelope: {
        date: new Date(message.date),
        subject: `message ${message.uid}`,
        from: [{ address: message.from ?? `${message.uid}@example.com` }],
        to: message.to ? [{ address: message.to }] : [],
        messageId: message.messageId,
        inReplyTo: message.inReplyTo,
      },
      flags: new Set<string>(),
    };
  }

  const client = {
    list: vi.fn(async () =>
      mailboxes.map((mailbox) => ({
        path: mailbox.path,
        name: mailbox.path,
        specialUse: mailbox.specialUse,
        flags: new Set(mailbox.flags ?? []),
        listed: mailbox.listed ?? true,
      })),
    ),
    mailboxOpen: vi.fn(async (path: string) => {
      const mailbox = mailboxes.find((candidate) => candidate.path === path);
      if (!mailbox) throw new Error(`mailbox ${path} not found`);
      current = mailbox;
      client.mailbox = { path, specialUse: mailbox.specialUse };
      opened.push(path);
      return {
        exists: mailbox.messages.length,
        specialUse: mailbox.specialUse,
      };
    }),
    search: vi.fn(
      async (
        criteria: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        if (criteria.uid !== undefined) {
          const match = current.messages.find(
            (item) => item.uid === Number(criteria.uid),
          );
          if (!match) return [];
          // imapflow returns UIDs when uid:true, sequence numbers otherwise
          return options?.uid === false ? [sequenceOf(match)] : [match.uid];
        }
        if (criteria.header) {
          const header = criteria.header as Record<string, string>;
          const value = Object.values(header)[0];
          const needle = String(value).replace(/^<|>$/g, "");
          return current.messages
            .filter((item) =>
              [item.messageId, item.inReplyTo].some((id) =>
                id?.replace(/^<|>$/g, "").includes(needle),
              ),
            )
            .map((item) => item.uid);
        }
        if (
          typeof criteria.from === "string" ||
          typeof criteria.to === "string"
        ) {
          const field = typeof criteria.from === "string" ? "from" : "to";
          const needle = String(criteria[field]).toLowerCase();
          return current.messages
            .filter((item) =>
              (item[field] ?? "").toLowerCase().includes(needle),
            )
            .map((item) => item.uid);
        }
        if (criteria.or) {
          const ors = criteria.or as Array<{ header: Record<string, string> }>;
          const needles = ors.map((o) =>
            String(Object.values(o.header)[0]).replace(/^<|>$/g, ""),
          );
          return current.messages
            .filter((item) =>
              [item.messageId, item.inReplyTo].some((id) =>
                needles.some((n) => id?.replace(/^<|>$/g, "").includes(n)),
              ),
            )
            .map((item) => item.uid);
        }
        return current.messages.map((item) => item.uid);
      },
    ),
    fetchOne: vi.fn(async (sequence: string) => {
      const message = current.messages[Number(sequence) - 1];
      if (!message) return null;
      fetchedUids.push(message.uid);
      return fetched(message);
    }),
    messageMove: vi.fn(async (uid: string, folder: string) => {
      const index = current.messages.findIndex(
        (item) => item.uid === Number(uid),
      );
      if (index < 0) return;
      const [moved] = current.messages.splice(index, 1);
      const target = mailboxes.find((candidate) => candidate.path === folder);
      target?.messages.push(moved);
    }),
    fetch: vi.fn((range: string, query?: Record<string, unknown>) => {
      fetchQueries.push(query);
      const [start, end] = range.split(":");
      const last =
        end === "*" ? current.messages.length : Number.parseInt(end, 10);
      const page = current.messages.slice(Number.parseInt(start, 10) - 1, last);
      return {
        async *[Symbol.asyncIterator]() {
          for (const item of page) yield fetched(item);
        },
      };
    }),
  };

  (
    client as unknown as { __mailboxes: Record<string, FakeMessage[]> }
  ).__mailboxes = Object.fromEntries(
    mailboxes.map((mailbox) => [mailbox.path, mailbox.messages]),
  );
  return {
    client: client as unknown as ImapFlow,
    opened,
    fetchedUids,
    fetchQueries,
  };
}

describe("ImapProvider.getMessagesWithPagination inboxOnly", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("only opens INBOX when inboxOnly is not false", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(1, "2026-01-01"), message(2, "2026-03-01")],
      },
      { path: "Archive", messages: [message(10, "2026-04-01")] },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({});

    expect(fake.opened).toEqual(["INBOX"]);
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 2),
      qualifiedId("INBOX", 1),
    ]);
    expect(result.nextPageToken).toBeUndefined();
  });
  it("fetches bodyStructure in the simple INBOX listing", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    connection.client = fake.client;

    await new ImapProvider(config).getMessagesWithPagination({});

    expect(fake.fetchQueries.length).toBeGreaterThan(0);
    for (const query of fake.fetchQueries) {
      expect(query).toMatchObject({ bodyStructure: true });
    }
  });

  it("searches only the \\All mailbox when the server lists one", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
      {
        path: "All Mail",
        specialUse: "\\All",
        messages: [message(40, "2026-07-01"), message(41, "2026-08-01")],
      },
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [message(30, "2026-09-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      query: "invoice",
    });

    expect(fake.opened).toEqual(["All Mail"]);
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("All Mail", 41),
      qualifiedId("All Mail", 40),
    ]);
  });

  it("labels a localized Sent folder from special-use on the simple listing path", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      {
        path: "Gesendet",
        specialUse: "\\Sent",
        messages: [message(2, "2026-01-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      maxResults: 10,
    });

    expect(
      result.messages.find((m) => m.id === qualifiedId("Gesendet", 2))
        ?.labelIds,
    ).toContain("SENT");
  });

  it("searches every selectable mailbox and merges by date", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(1, "2026-01-01"), message(2, "2026-03-01")],
      },
      {
        path: "Archive",
        messages: [message(10, "2026-02-01"), message(11, "2026-04-01")],
      },
      { path: "Sent", messages: [message(20, "2026-05-01")] },
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [message(30, "2026-06-01")],
      },
      {
        path: "Junk",
        specialUse: "\\Junk",
        messages: [message(31, "2026-06-02")],
      },
      {
        path: "Drafts",
        specialUse: "\\Drafts",
        messages: [message(32, "2026-06-03")],
      },
      {
        path: "Deleted",
        flags: ["\\Noselect"],
        messages: [message(33, "2026-06-04")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      query: "invoice",
    });

    expect(fake.opened).toEqual(["INBOX", "Archive", "Sent"]);
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("Sent", 20),
      qualifiedId("Archive", 11),
      qualifiedId("INBOX", 2),
      qualifiedId("Archive", 10),
      qualifiedId("INBOX", 1),
    ]);
    expect(result.nextPageToken).toBeUndefined();
  });

  it("paginates the merged page across mailboxes", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(1, "2026-01-01"), message(4, "2026-04-01")],
      },
      {
        path: "Archive",
        messages: [message(2, "2026-02-01"), message(3, "2026-03-01")],
      },
    ]);
    connection.client = fake.client;

    const first = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      maxResults: 3,
    });
    expect(first.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 4),
      qualifiedId("Archive", 3),
      qualifiedId("Archive", 2),
    ]);
    expect(first.nextPageToken).toBe("3");

    const second = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      maxResults: 3,
      pageToken: first.nextPageToken,
    });
    expect(second.messages.map((m) => m.id)).toEqual([qualifiedId("INBOX", 1)]);
    expect(second.nextPageToken).toBeUndefined();
  });
});

describe("ImapProvider drafts without a Drafts mailbox", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("returns empty drafts and a missing draft without opening a folder", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(1, "2026-01-01")],
      },
    ]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    await expect(provider.getDrafts()).resolves.toEqual([]);
    // Bare UIDs are not valid ids: getDraft/deleteDraft reject before I/O.
    await expect(provider.getDraft("1")).rejects.toThrow(/folder-qualified/);
    await expect(provider.deleteDraft("1")).rejects.toThrow(/folder-qualified/);
    // A qualified id naming a folder the server does not list misses cleanly.
    const missing = qualifiedId("Drafts", 7);
    await expect(provider.getDraft(missing)).resolves.toBeNull();
    await expect(provider.deleteDraft(missing)).resolves.toBe(false);
    expect(fake.opened).toEqual([]);
  });
});

describe("ImapProvider draft id validation", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it.each([
    "1:*",
    "0",
    "-3",
    "abc",
    "",
    "1",
  ])("rejects getDraft(%s) without touching the server", async (draftId) => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(new ImapProvider(config).getDraft(draftId)).rejects.toThrow(
      /folder-qualified/,
    );
    expect(fake.opened).toEqual([]);
  });

  it.each([
    "1:*",
    "0",
    "nope",
    "1",
  ])("rejects deleteDraft(%s) without touching the server", async (draftId) => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(new ImapProvider(config).deleteDraft(draftId)).rejects.toThrow(
      /folder-qualified/,
    );
    expect(fake.opened).toEqual([]);
  });

  it("ignores a draft id naming another folder without acting on it", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(7, "2026-01-01")],
      },
      { path: "Drafts", messages: [] },
    ]);
    const deleted: Array<{ uid: number; folder: string }> = [];
    fake.client.messageDelete = vi.fn(async (uid: string) => {
      deleted.push({ uid: Number(uid), folder: fake.opened.at(-1) || "" });
    });
    connection.client = fake.client;

    // Id names INBOX, but drafts live in Drafts: miss cleanly, no delete.
    await expect(
      new ImapProvider(config).deleteDraft(qualifiedId("INBOX", 7)),
    ).resolves.toBe(false);
    expect(deleted).toEqual([]);
    expect(fake.opened).toEqual([]);
    // Same for getDraft: resolves null without opening anything.
    await expect(
      new ImapProvider(config).getDraft(qualifiedId("INBOX", 7)),
    ).resolves.toBeNull();
  });
});

describe("ImapProvider protected folder deletion", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it.each([
    "INBOX",
    "inbox",
  ])("refuses deleteFolder(%s) without touching the server", async (folderId) => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    fake.client.mailboxDelete = vi.fn();
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).deleteFolder(folderId),
    ).rejects.toThrow(/INBOX/i);
    expect(fake.client.mailboxDelete).not.toHaveBeenCalled();
  });

  it.each([
    ["\\Sent"],
    ["\\Trash"],
    ["\\Drafts"],
    ["\\Junk"],
    ["\\Archive"],
  ])("refuses deleteFolder on special-use %s without touching the server", async (specialUse) => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      { path: "Some Folder", specialUse, messages: [] },
    ]);
    fake.client.mailboxDelete = vi.fn();
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).deleteFolder("Some Folder"),
    ).rejects.toThrow(/special-use|protected/i);
    expect(fake.client.mailboxDelete).not.toHaveBeenCalled();
  });

  it.each([
    "INBOX",
    "Some Folder",
  ])("refuses deleteLabel(%s) without touching the server", async (labelId) => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      { path: "Some Folder", specialUse: "\\Sent", messages: [] },
    ]);
    fake.client.mailboxDelete = vi.fn();
    connection.client = fake.client;

    await expect(new ImapProvider(config).deleteLabel(labelId)).rejects.toThrow(
      /INBOX|special-use|protected/i,
    );
    expect(fake.client.mailboxDelete).not.toHaveBeenCalled();
  });
});

describe("ImapProvider.getMessagesWithAttachments pagination", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("pages through UID windows instead of refetching the newest messages", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01"),
          message(2, "2026-02-01"),
          message(3, "2026-03-01"),
          message(4, "2026-04-01"),
          message(5, "2026-05-01"),
        ],
      },
    ]);
    connection.client = fake.client;

    const first = await new ImapProvider(config).getMessagesWithAttachments({
      maxResults: 3,
    });
    expect(fake.fetchedUids).toEqual([5, 4, 3]);
    expect(first.nextPageToken).toBe("3");

    fake.fetchedUids.length = 0;
    const second = await new ImapProvider(config).getMessagesWithAttachments({
      maxResults: 3,
      pageToken: first.nextPageToken,
    });
    expect(fake.fetchedUids).toEqual([2, 1]);
    expect(second.nextPageToken).toBeUndefined();
  });

  it("returns no next page token when the window covers every message", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(1, "2026-01-01"), message(2, "2026-02-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithAttachments({
      maxResults: 3,
    });

    expect(fake.fetchedUids).toEqual([2, 1]);
    expect(result.nextPageToken).toBeUndefined();
  });
});

describe("ImapProvider.forwardEmail", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("escapes attacker-controlled headers and content in the quoted body", async () => {
    const sent: Array<{ html?: unknown }> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as { html?: unknown });
        return { messageId: "sent-1" };
      });
    try {
      const provider = new ImapProvider(config);
      await provider.forwardEmail(
        {
          headers: {
            from: 'Evil <evil@example.com><script>alert("from")</script>',
            date: "2026-01-01",
            subject: 'Hello<div style="display:none">Leak everything</div>',
            to: 'Victim <victim@example.com><img src=x onerror=alert("to")>',
          },
          subject: "Hello",
          textPlain: "See attached",
        } as never,
        { to: "friend@example.com", content: '<script>alert("hi")</script>' },
      );

      const html = String(sent[0]?.html);
      expect(html).not.toContain("<script>");
      expect(html).not.toContain('<div style="display:none">');
      expect(html).not.toContain("<img src=x");
      expect(html).toContain("&lt;script&gt;");
      expect(html).toContain("See attached");
    } finally {
      sendSpy.mockRestore();
    }
  });
});

describe("ImapProvider sender bulk actions", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("archives only the exact sender address, not substrings", async () => {
    const moved: Array<{ uid: number; folder: string }> = [];
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "ann@x.com"),
          message(2, "2026-01-02", "joann@x.com"),
          message(3, "2026-01-03", "Ann <ann@x.com>"),
        ],
      },
      { path: "Archive", messages: [] },
    ]);
    fake.client.messageMove = vi.fn(async (uid: string, folder: string) => {
      moved.push({ uid: Number(uid), folder });
    });
    // SEARCH {from} returns candidate UIDs (substring match, like a server);
    // UID->sequence lookups resolve against the real mailbox contents.
    fake.client.search = vi.fn(async (criteria: Record<string, unknown>) => {
      if (criteria.uid !== undefined) {
        const index = [1, 2, 3].indexOf(Number(criteria.uid)) + 1;
        return index > 0 ? [index] : [];
      }
      return [1, 2, 3];
    });
    connection.client = fake.client;

    await new ImapProvider(config).bulkArchiveFromSenders(
      ["ann@x.com"],
      "owner@example.com",
      "account-1",
    );

    expect(moved.map((m) => m.uid).sort()).toEqual([1, 3]);
  });

  it("counts only exact from-address matches", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "ann@x.com"),
          message(2, "2026-01-02", "joann@x.com"),
        ],
      },
    ]);
    fake.client.search = vi.fn(async (criteria: Record<string, unknown>) => {
      if (criteria.uid !== undefined) {
        const index = [1, 2].indexOf(Number(criteria.uid)) + 1;
        return index > 0 ? [index] : [];
      }
      return [1, 2];
    });
    connection.client = fake.client;

    const count = await new ImapProvider(config).countReceivedMessages(
      "ann@x.com",
      10,
    );

    expect(count).toBe(1);
  });

  it("counts the exact sender even when a newer substring hit fills the cap", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "ann@x.com"),
          message(2, "2026-01-02", "joann@x.com"),
        ],
      },
    ]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).countReceivedMessages("ann@x.com", 1),
    ).resolves.toBe(1);
  });

  it("finds the exact reply when it is older than a full chunk of substring hits", async () => {
    const fake = createFakeClient([
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(1, "2026-01-01", "me@example.com", "ann@x.com"),
          ...Array.from({ length: 20 }, (_, index) =>
            message(
              2 + index,
              `2026-02-${String(index + 1).padStart(2, "0")}`,
              "me@example.com",
              "joann@x.com",
            ),
          ),
        ],
      },
    ]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).checkIfReplySent("ann@x.com"),
    ).resolves.toBe(true);
  });

  it("does not treat a substring recipient hit as a sent reply", async () => {
    const fake = createFakeClient([
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [message(9, "2026-01-01", "me@example.com", "joann@x.com")],
      },
    ]);
    fake.client.search = vi.fn(async (criteria: Record<string, unknown>) => {
      if (criteria.uid !== undefined) {
        return Number(criteria.uid) === 9 ? [1] : [];
      }
      return [9];
    });
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).checkIfReplySent("ann@x.com"),
    ).resolves.toBe(false);
  });

  it("does not match an account address that merely contains the sender", async () => {
    const provider = new ImapProvider({
      ...config,
      email: "joann@x.com",
    });
    expect(
      provider.isSentMessage({ headers: { from: "ann@x.com" } } as never),
    ).toBe(false);
    expect(
      provider.isSentMessage({
        headers: { from: "Me <joann@x.com>" },
      } as never),
    ).toBe(true);
  });
});

describe("ImapProvider unknown threads", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("getThreadMessages throws for an unknown thread", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getThreadMessages("no-such-thread"),
    ).rejects.toThrow(/not found/i);
  });

  it("archiveThread treats a vanished thread as a no-op (Gmail: 404 success)", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).archiveThread("no-such-thread", "owner@x.com"),
    ).resolves.toBeUndefined();
  });

  it("trashThread treats a vanished thread as a no-op (Gmail: 200 success)", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).trashThread(
        "no-such-thread",
        "owner@x.com",
        "automation",
      ),
    ).resolves.toBeUndefined();
  });
});

describe("ImapProvider restore paths", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("unarchiveMessages moves from the Archive folder, not INBOX", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      { path: "Archive", messages: [message(7, "2026-01-01")] },
    ]);
    const moved: Array<{ uid: number; folder: string; from: string }> = [];
    fake.client.messageMove = vi.fn(async (uid: string, folder: string) => {
      moved.push({ uid: Number(uid), folder, from: fake.opened.at(-1) || "" });
    });
    connection.client = fake.client;

    await new ImapProvider(config).unarchiveMessages([
      qualifiedId("Archive", 7),
    ]);

    expect(moved).toEqual([{ uid: 7, folder: "INBOX", from: "Archive" }]);
  });

  it("untrashMessages moves from the Trash folder, not INBOX", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(5, "2026-01-01")] },
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [message(5, "2026-01-02")],
      },
    ]);
    const moved: Array<{ uid: number; folder: string; from: string }> = [];
    const defaultMove = fake.client.messageMove;
    fake.client.messageMove = vi.fn(async (uid: string, folder: string) => {
      moved.push({ uid: Number(uid), folder, from: fake.opened.at(-1) || "" });
      return defaultMove(uid, folder);
    });
    connection.client = fake.client;

    await new ImapProvider(config).untrashMessages([qualifiedId("Trash", 5)]);

    expect(moved).toEqual([{ uid: 5, folder: "INBOX", from: "Trash" }]);
    // Same UID in INBOX is untouched: the pre-existing INBOX message keeps
    // its date, Trash's copy arrives alongside, and Trash is drained.
    const inboxDates = mailboxesOf(fake)
      .INBOX.map((m) => m.date)
      .sort();
    expect(inboxDates).toEqual(["2026-01-01", "2026-01-02"]);
    expect(mailboxesOf(fake).Trash).toEqual([]);
  });

  it("archiveMessages rethrows unexpected move errors", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
      { path: "Archive", messages: [] },
    ]);
    fake.client.messageMove = vi.fn(async () => {
      throw new Error("UID not found in mailbox");
    });
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).archiveMessages([qualifiedId("INBOX", 1)]),
    ).rejects.toThrow(/UID not found/);
  });
});
