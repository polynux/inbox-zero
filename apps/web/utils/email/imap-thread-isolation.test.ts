import { beforeEach, describe, expect, it, vi } from "vitest";
import { ImapProvider } from "@/utils/email/imap";
import { buildThreadId } from "@/utils/imap/thread";
import { encodeImapMessageId } from "@/utils/imap/message-id";

const connection = vi.hoisted(() => ({ client: null as unknown }));

vi.mock("@/utils/imap/client", () => ({
  withImapConnection: async <T>(
    _config: unknown,
    fn: (client: never) => Promise<T>,
  ): Promise<T> => fn(connection.client as never),
}));

type Msg = {
  uid: number;
  subject?: string;
  from?: string;
  to?: string;
  messageId?: string;
  inReplyTo?: string;
};

const config = {
  email: "owner@example.com",
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

function subjectOf(m: Msg): string {
  return m.subject ?? `message ${m.uid}`;
}

// Fake models RFC 3501 per-mailbox UIDs: SEARCH {uid} resolves in the opened
// folder, substring SEARCH returns every message as a candidate (like a
// server), and MOVE for a UID the opened folder lacks completes with OK
// without touching anything — the bug class under test.
function build(
  mailboxes: Record<string, Msg[]>,
  specialUse: Record<string, string>,
) {
  const state = {
    opened: "INBOX",
    moves: [] as Array<{ uid: number; from: string; to: string }>,
    flags: [] as Array<{ uid: number; folder: string; flag: string }>,
  };
  const envelopeOf = (m: Msg) => ({
    uid: m.uid,
    flags: new Set<string>(),
    envelope: {
      date: new Date("2026-01-01"),
      subject: subjectOf(m),
      from: [{ address: m.from ?? "someone@example.com" }],
      to: m.to ? [{ address: m.to }] : [],
      messageId: m.messageId,
      inReplyTo: m.inReplyTo,
    },
  });
  const client = {
    mailbox: undefined as unknown,
    list: async () =>
      Object.keys(mailboxes).map((path) => ({
        path,
        name: path,
        listed: true,
        flags: new Set<string>(),
        specialUse: specialUse[path],
      })),
    mailboxOpen: async (path: string) => {
      state.opened = path;
      client.mailbox = {
        path,
        exists: mailboxes[path].length,
        uidValidity: 1,
        uidNext: 100,
      };
      return client.mailbox;
    },
    fetch: (range: string) =>
      (async function* () {
        const list = mailboxes[state.opened];
        // Honor sequence ranges (`start:*`) so tests can place threads
        // outside a listing window like the old newest-200 scan had.
        const start = Number(range.split(":")[0]) || 1;
        for (const m of list.slice(start - 1)) yield envelopeOf(m);
      })(),
    fetchOne: async (seq: string) => {
      const m = mailboxes[state.opened][Number(seq) - 1];
      return m ? envelopeOf(m) : null;
    },
    search: async (criteria: Record<string, unknown>) => {
      const list = mailboxes[state.opened];
      if (criteria.uid !== undefined) {
        const index = list.findIndex((m) => m.uid === Number(criteria.uid));
        return index >= 0 ? [index + 1] : [];
      }
      if (typeof criteria.from === "string") {
        const needle = criteria.from.toLowerCase();
        return list
          .filter((m) => (m.from ?? "").toLowerCase().includes(needle))
          .map((m) => m.uid);
      }
      if (typeof criteria.to === "string") {
        const needle = criteria.to.toLowerCase();
        return list
          .filter((m) => (m.to ?? "").toLowerCase().includes(needle))
          .map((m) => m.uid);
      }
      return list.map((m) => m.uid);
    },
    messageMove: async (uid: string, to: string) => {
      const list = mailboxes[state.opened];
      const index = list.findIndex((m) => m.uid === Number(uid));
      if (index < 0) return;
      state.moves.push({ uid: Number(uid), from: state.opened, to });
      const [moved] = list.splice(index, 1);
      mailboxes[to]?.push(moved);
    },
    messageFlagsAdd: vi.fn(async (uid: string, flags: string[]) => {
      state.flags.push({
        uid: Number(uid),
        folder: state.opened,
        flag: flags[0],
      });
    }),
    messageFlagsRemove: vi.fn(),
    logout: async () => undefined,
  };
  return { client, state, mailboxes };
}

beforeEach(() => {
  connection.client = null;
});

describe("thread ops act on the right message in the right mailbox", () => {
  it("archives only the INBOX member, not an unrelated same-UID message in Archive", async () => {
    const { client, state, mailboxes } = build(
      {
        INBOX: [{ uid: 3, subject: "thread root", messageId: "<root@x>" }],
        Archive: [
          { uid: 3, subject: "unrelated archived", messageId: "<news@x>" },
        ],
      },
      { Archive: "\\Archive" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    await new ImapProvider(config).archiveThread(threadId, "owner@example.com");

    expect(state.moves).toEqual([{ uid: 3, from: "INBOX", to: "Archive" }]);
    expect(mailboxes.Archive.map((m) => m.subject).sort()).toEqual([
      "thread root",
      "unrelated archived",
    ]);
  });

  it("restores only the Archive member and never touches a same-UID message in Sent", async () => {
    const { client, state, mailboxes } = build(
      {
        INBOX: [],
        Archive: [{ uid: 5, subject: "archived root", messageId: "<root@x>" }],
        Sent: [
          {
            uid: 5,
            subject: "unrelated sent",
            messageId: "<other@x>",
          },
        ],
      },
      { Sent: "\\Sent", Archive: "\\Archive" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    await new ImapProvider(config).unarchiveThread(threadId);

    expect(state.moves).toEqual([{ uid: 5, from: "Archive", to: "INBOX" }]);
    // Same UID in Sent is untouched: still there with its own subject.
    expect(mailboxes.Sent.map((m) => m.subject)).toEqual(["unrelated sent"]);
  });
  it("returns same-UID members in two folders instead of collapsing them", async () => {
    const { client } = build(
      {
        INBOX: [{ uid: 5, subject: "root", messageId: "<root@x>" }],
        Sent: [
          {
            uid: 5,
            subject: "reply",
            messageId: "<reply@x>",
            inReplyTo: "<root@x>",
          },
        ],
      },
      { Sent: "\\Sent" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    const messages = await new ImapProvider(config).getThreadMessages(threadId);

    expect(messages.map((m) => m.subject).sort()).toEqual(["reply", "root"]);
    const ids = new Set(messages.map((m) => m.id));
    expect(ids.size).toBe(2);
    expect(ids).toContain(encodeImapMessageId("INBOX", 5));
    expect(ids).toContain(encodeImapMessageId("Sent", 5));
  });

  it("marks thread members read in their own mailboxes", async () => {
    const { client, state } = build(
      {
        INBOX: [{ uid: 5, subject: "root", messageId: "<root@x>" }],
        Sent: [
          {
            uid: 5,
            subject: "reply",
            messageId: "<reply@x>",
            inReplyTo: "<root@x>",
          },
        ],
      },
      { Sent: "\\Sent" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    await new ImapProvider(config).markReadThread(threadId, true);

    expect(state.flags).toEqual([
      { uid: 5, folder: "INBOX", flag: "\\Seen" },
      { uid: 5, folder: "Sent", flag: "\\Seen" },
    ]);
  });
});

describe("thread resolution reaches Trash and Junk", () => {
  it("untrashThread finds a thread living only in Trash", async () => {
    const { client, state, mailboxes } = build(
      {
        INBOX: [],
        Trash: [{ uid: 4, subject: "trashed", messageId: "<t@x>" }],
      },
      { Trash: "\\Trash" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<t@x>");
    await new ImapProvider(config).untrashThread(threadId);

    expect(state.moves).toEqual([{ uid: 4, from: "Trash", to: "INBOX" }]);
    expect(mailboxes.Trash).toEqual([]);
  });

  it("markNotSpam finds a thread living only in Junk", async () => {
    const { client, state, mailboxes } = build(
      {
        INBOX: [],
        Junk: [{ uid: 7, subject: "newsletter", messageId: "<j@x>" }],
      },
      { Junk: "\\Junk" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<j@x>");
    await new ImapProvider(config).markNotSpam(threadId);

    expect(state.moves).toEqual([{ uid: 7, from: "Junk", to: "INBOX" }]);
    expect(mailboxes.Junk).toEqual([]);
  });
});

describe("thread ops other than untrash/markNotSpam ignore Trash and Junk", () => {
  it("archiveThread leaves a trashed reply in Trash", async () => {
    const { client, state } = build(
      {
        INBOX: [{ uid: 1, subject: "root", messageId: "<root@x>" }],
        Trash: [
          {
            uid: 2,
            subject: "deleted reply",
            messageId: "<reply@x>",
            inReplyTo: "<root@x>",
          },
        ],
        Archive: [],
      },
      { Trash: "\\Trash", Archive: "\\Archive" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    await new ImapProvider(config).archiveThread(threadId, "owner@example.com");

    expect(state.moves).toEqual([{ uid: 1, from: "INBOX", to: "Archive" }]);
  });

  it("unarchiveThread (snooze wake-up) does not pull a trashed reply into INBOX", async () => {
    const { client, state } = build(
      {
        INBOX: [],
        Archive: [{ uid: 1, subject: "root", messageId: "<root@x>" }],
        Trash: [
          {
            uid: 2,
            subject: "deleted reply",
            messageId: "<reply@x>",
            inReplyTo: "<root@x>",
          },
        ],
      },
      { Trash: "\\Trash", Archive: "\\Archive" },
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    await new ImapProvider(config).unarchiveThread(threadId);

    expect(state.moves).toEqual([{ uid: 1, from: "Archive", to: "INBOX" }]);
  });
});

describe("header-less threads resolve without a full scan", () => {
  it("fetches a header-less thread directly when only \\All is scanned", async () => {
    const { client } = build(
      {
        INBOX: [{ uid: 9, subject: "no headers" }],
        "All Mail": [{ uid: 9, subject: "no headers" }],
      },
      { "All Mail": "\\All" },
    );
    connection.client = client;

    // The thread id carries the listing key, so resolution opens that folder
    // directly instead of scanning only \All.
    const listedThreadId = buildThreadId(
      undefined,
      undefined,
      undefined,
      encodeImapMessageId("INBOX", 9),
    );
    const messages = await new ImapProvider(config).getThreadMessages(
      listedThreadId,
    );

    expect(messages).toHaveLength(1);
    expect(messages[0].id).toBe(encodeImapMessageId("INBOX", 9));
  });
});

describe("thread lookup via HEADER search", () => {
  it("finds a thread whose members are all older than any listing window", async () => {
    // 210 filler messages bury the two thread members past message 200, so
    // a newest-N scan alone never sees the thread; HEADER search must.
    const fillers = Array.from({ length: 210 }, (_, index) => ({
      uid: 100 + index,
      subject: `filler ${100 + index}`,
      messageId: `<filler${100 + index}@x>`,
    }));
    const { client } = build(
      {
        INBOX: [
          { uid: 1, subject: "ancient root", messageId: "<ancient@x>" },
          {
            uid: 2,
            subject: "ancient reply",
            messageId: "<ancient-reply@x>",
            inReplyTo: "<ancient@x>",
          },
          ...fillers,
        ],
      },
      {},
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<ancient@x>");
    const messages = await new ImapProvider(config).getThreadMessages(threadId);

    expect(messages.map((m) => m.subject).sort()).toEqual([
      "ancient reply",
      "ancient root",
    ]);
  });

  it("filters substring false-positives from HEADER search", async () => {
    const { client } = build(
      {
        INBOX: [
          { uid: 1, subject: "exact root", messageId: "<root@x>" },
          { uid: 2, subject: "lookalike", messageId: "<root@x-body>" },
        ],
      },
      {},
    );
    connection.client = client;

    const threadId = buildThreadId(undefined, undefined, "<root@x>");
    const messages = await new ImapProvider(config).getThreadMessages(threadId);

    expect(messages.map((m) => m.subject)).toEqual(["exact root"]);
  });
});

describe("exact address matching is not hidden by substring hits", () => {
  it("checkIfReplySent is true when a newer substring hit exists", async () => {
    const { client } = build(
      {
        INBOX: [],
        Sent: [
          { uid: 1, subject: "re: hi", to: "ann@x.com" },
          { uid: 2, subject: "hello", to: "joann@x.com" },
        ],
      },
      { Sent: "\\Sent" },
    );
    connection.client = client;

    await expect(
      new ImapProvider(config).checkIfReplySent("ann@x.com"),
    ).resolves.toBe(true);
  });
});
