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

type FakeMessage = { uid: number; date: string; from?: string; to?: string };

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
): FakeMessage {
  return { uid, date, from, to };
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
      opened.push(path);
      return { exists: mailbox.messages.length };
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
        return current.messages.map((item) => item.uid);
      },
    ),
    fetchOne: vi.fn(async (sequence: string) => {
      const message = current.messages[Number(sequence) - 1];
      fetchedUids.push(message.uid);
      return fetched(message);
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
    expect(result.messages.map((m) => m.id)).toEqual(["2", "1"]);
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
    expect(result.messages.map((m) => m.id)).toEqual(["41", "40"]);
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
      "20",
      "11",
      "2",
      "10",
      "1",
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
    expect(first.messages.map((m) => m.id)).toEqual(["4", "3", "2"]);
    expect(first.nextPageToken).toBe("3");

    const second = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      maxResults: 3,
      pageToken: first.nextPageToken,
    });
    expect(second.messages.map((m) => m.id)).toEqual(["1"]);
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
    await expect(provider.getDraft("1")).resolves.toBeNull();
    await expect(provider.deleteDraft("1")).resolves.toBe(false);
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
  ])("rejects getDraft(%s) without touching the server", async (draftId) => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(new ImapProvider(config).getDraft(draftId)).rejects.toThrow(
      "positive integer",
    );
    expect(fake.opened).toEqual([]);
  });

  it.each([
    "1:*",
    "0",
    "nope",
  ])("rejects deleteDraft(%s) without touching the server", async (draftId) => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(new ImapProvider(config).deleteDraft(draftId)).rejects.toThrow(
      "positive integer",
    );
    expect(fake.opened).toEqual([]);
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
    fake.client.search = vi.fn(
      async (
        criteria: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        if (criteria.uid !== undefined) {
          const index = [1, 2, 3].indexOf(Number(criteria.uid)) + 1;
          return index > 0 ? [index] : [];
        }
        return [1, 2, 3];
      },
    );
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
    fake.client.search = vi.fn(
      async (
        criteria: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        if (criteria.uid !== undefined) {
          const index = [1, 2].indexOf(Number(criteria.uid)) + 1;
          return index > 0 ? [index] : [];
        }
        return [1, 2];
      },
    );
    connection.client = fake.client;

    const count = await new ImapProvider(config).countReceivedMessages(
      "ann@x.com",
      10,
    );

    expect(count).toBe(1);
  });

  it("does not treat a substring recipient hit as a sent reply", async () => {
    const fake = createFakeClient([
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [message(9, "2026-01-01", "me@example.com", "joann@x.com")],
      },
    ]);
    fake.client.search = vi.fn(
      async (
        criteria: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        if (criteria.uid !== undefined) {
          return Number(criteria.uid) === 9 ? [1] : [];
        }
        return [9];
      },
    );
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
