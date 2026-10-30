import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ImapProvider } from "@/utils/email/imap";

const connection = vi.hoisted(() => ({ client: null as unknown }));

vi.mock("@/utils/imap/client", () => ({
  withImapConnection: async <T>(
    _config: unknown,
    fn: (client: never) => Promise<T>,
  ): Promise<T> => fn(connection.client as never),
}));

type FakeMessage = { uid: number; date: string };

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

function message(uid: number, date: string): FakeMessage {
  return { uid, date };
}

function createFakeClient(mailboxes: FakeMailbox[]) {
  const opened: string[] = [];
  const fetchedUids: number[] = [];
  let current = mailboxes[0];

  function fetched(message: FakeMessage) {
    return {
      uid: message.uid,
      envelope: {
        date: new Date(message.date),
        subject: `message ${message.uid}`,
        from: [{ address: `${message.uid}@example.com` }],
        to: [],
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
        if (options?.uid) return current.messages.map((item) => item.uid);
        const sequence =
          current.messages.findIndex(
            (item) => item.uid === Number(criteria.uid),
          ) + 1;
        return sequence > 0 ? [sequence] : [];
      },
    ),
    fetchOne: vi.fn(async (sequence: string) => {
      const message = current.messages[Number(sequence) - 1];
      fetchedUids.push(message.uid);
      return fetched(message);
    }),
    fetch: vi.fn((range: string) => {
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
