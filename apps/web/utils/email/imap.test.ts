import type { ImapFlow } from "imapflow";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ImapMessageNotFoundError, ImapProvider } from "@/utils/email/imap";
import * as mailModule from "@/utils/imap/mail";
import { encodeImapDraftId } from "@/utils/imap/draft";
import { encodeImapMessageId } from "@/utils/imap/message-id";
import { buildThreadId } from "@/utils/imap/thread";
import { DraftNotFoundError } from "@/utils/error";
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
  cc?: string;
  bcc?: string;
  subject?: string;
  html?: string;
  messageId?: string;
  inReplyTo?: string;
  references?: string;
  hasAttachment?: boolean;
  /** Full MIME source the fake download yields instead of rawOf. */
  raw?: string;
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
  headers?: {
    messageId?: string;
    inReplyTo?: string;
    references?: string;
    cc?: string;
    bcc?: string;
    subject?: string;
    html?: string;
    hasAttachment?: boolean;
  },
): FakeMessage {
  return {
    uid,
    date,
    from,
    to,
    ...headers,
  };
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

// Bridge's APPEND rewrite on the stored bytes: In-Reply-To is dropped and
// References is replaced by a per-copy internal id, so reads that go through
// the raw MIME (draft attachment parts, getMessage downloads) see the same
// headers a live Bridge serves.
function rewriteRawThreading(raw: string, uid: number): string {
  const [head, ...body] = raw.split("\r\n\r\n");
  const kept = (head ?? "")
    .split("\r\n")
    .filter((line) => !/^(in-reply-to|references):/i.test(line))
    .join("\r\n");
  const rewritten = `${kept}\r\nReferences: <fake-internal-${uid}@protonmail.internalid>`;
  return body.length > 0
    ? `${rewritten}\r\n\r\n${body.join("\r\n\r\n")}`
    : rewritten;
}

function stripAttachmentIds(raw: string): string {
  return raw
    .split("\r\n")
    .filter((line) => !/^x-attachment-id:/i.test(line))
    .join("\r\n");
}

type FakeImapTestClient = {
  client: ImapFlow;
  opened: string[];
  fetchedUids: number[];
  fetchQueries: Array<Record<string, unknown> | undefined>;
  appended: Array<{ path: string; flags: string[] }>;
};

function createFakeClient(
  mailboxes: FakeMailbox[],
  options?: { uidPlus?: boolean; rewriteDraftHeaders?: boolean },
): FakeImapTestClient {
  const mailboxSupportsUidPlus = options?.uidPlus ?? true;
  // Bridge rewrites a draft's threading headers on APPEND: In-Reply-To is
  // dropped and References is replaced by its own internal id. Tests that
  // need that behavior opt in so other suites keep exact headers.
  const rewriteDraftHeaders = options?.rewriteDraftHeaders ?? false;
  const appended: Array<{ path: string; flags: string[] }> = [];
  const opened: string[] = [];
  let current = mailboxes[0];
  const fetchedUids: number[] = [];
  const fetchQueries: Array<Record<string, unknown> | undefined> = [];
  const sequenceOf = (message: FakeMessage) =>
    current.messages.indexOf(message) + 1;

  // The real ENVELOPE never carries References: it arrives via a
  // HEADER.FIELDS fetch (listings, thread resolution) or the full download
  // (getMessage). The fake surfaces it as a raw header buffer, but only when
  // the FETCH query asked for headers — otherwise this class of bug (a fetch
  // path that forgets `headers: ["references"]`) stays invisible.
  function headersOf(
    message: FakeMessage,
    query?: Record<string, unknown>,
  ): Buffer | undefined {
    if (!message.references) return;
    const wanted =
      Array.isArray(query?.headers) &&
      query.headers.some(
        (header) => String(header).toLowerCase() === "references",
      );
    if (!wanted) return;
    return Buffer.from(`References: ${message.references}\r\n`);
  }

  function fetched(message: FakeMessage, query?: Record<string, unknown>) {
    const wantsInternalDate = Boolean(
      (query as { internalDate?: unknown } | undefined)?.internalDate,
    );
    const bodyStructure = (query as { bodyStructure?: unknown } | undefined)
      ?.bodyStructure
      ? {
          part: "",
          childNodes: message.hasAttachment
            ? [
                { part: "1", type: "text/html" },
                {
                  part: "2",
                  type: "application/octet-stream",
                  disposition: "attachment",
                  dispositionParameters: { filename: "file.bin" },
                  size: 4,
                },
              ]
            : [{ part: "2", type: "application/octet-stream" }],
        }
      : undefined;
    return {
      uid: message.uid,
      ...(wantsInternalDate
        ? { internalDate: new Date(message.date) }
        : {
            envelope: {
              date: new Date(message.date),
              subject: message.subject ?? `message ${message.uid}`,
              from: [{ address: message.from ?? `${message.uid}@example.com` }],
              to: message.to ? [{ address: message.to }] : [],
              cc: message.cc ? [{ address: message.cc }] : [],
              bcc: message.bcc ? [{ address: message.bcc }] : [],
              messageId: message.messageId,
              inReplyTo: message.inReplyTo,
            },
            flags: new Set<string>(),
            headers: headersOf(message, query),
            bodyStructure,
          }),
    };
  }
  function rawOf(message: FakeMessage): string {
    const lines = [
      `From: ${message.from ?? `${message.uid}@example.com`}`,
      `To: ${message.to ?? "owner@example.com"}`,
      ...(message.cc ? [`Cc: ${message.cc}`] : []),
      ...(message.bcc ? [`Bcc: ${message.bcc}`] : []),
      `Subject: ${message.subject ?? `message ${message.uid}`}`,
      `Date: ${new Date(message.date).toUTCString()}`,
      message.messageId ? `Message-ID: ${message.messageId}` : "",
      message.inReplyTo ? `In-Reply-To: ${message.inReplyTo}` : "",
      message.references ? `References: ${message.references}` : "",
      "MIME-Version: 1.0",
      "Content-Type: text/html; charset=utf-8",
    ].filter(Boolean);
    return `${lines.join("\r\n")}\r\n\r\n${message.html ?? `body of message ${message.uid}`}\r\n`;
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
    status: vi.fn(async (path: string) => {
      const mailbox = mailboxes.find((candidate) => candidate.path === path);
      if (!mailbox) throw new Error(`mailbox ${path} not found`);
      return { messages: mailbox.messages.length, unseen: 0 };
    }),
    mailboxOpen: vi.fn(async (path: string) => {
      const mailbox = mailboxes.find((candidate) => candidate.path === path);
      if (!mailbox) throw new Error(`mailbox ${path} not found`);
      current = mailbox;
      client.mailbox = {
        path,
        specialUse: mailbox.specialUse,
        exists: mailbox.messages.length,
      };
      opened.push(path);
      return {
        path,
        exists: mailbox.messages.length,
        specialUse: mailbox.specialUse,
      };
    }),
    search: vi.fn(
      async (
        criteria: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        const matches = (
          item: FakeMessage,
          query: Record<string, unknown>,
        ): boolean =>
          Object.entries(query).every(([key, expected]) => {
            if (key === "uid" || key === "all") return true;
            if (key === "header") {
              const value = Object.values(
                expected as Record<string, string>,
              )[0];
              const needle = String(value).replace(/^<|>$/g, "");
              return [item.messageId, item.inReplyTo, item.references].some(
                (id) => id?.replace(/^<|>$/g, "").includes(needle),
              );
            }
            if (key === "from" || key === "to") {
              return (item[key] ?? "")
                .toLowerCase()
                .includes(String(expected).toLowerCase());
            }
            if (key === "subject") {
              return (item.subject ?? "")
                .toLowerCase()
                .includes(String(expected).toLowerCase());
            }
            if (key === "body") {
              const haystack =
                `${item.subject ?? ""} ${item.html ?? ""}`.toLowerCase();
              return haystack.includes(String(expected).toLowerCase());
            }
            if (key === "text") {
              const haystack =
                `${item.subject ?? ""} ${item.html ?? ""} ${item.from ?? ""} ${item.to ?? ""}`.toLowerCase();
              return haystack.includes(String(expected).toLowerCase());
            }
            if (key === "or") {
              return (expected as Array<Record<string, unknown>>).some((part) =>
                matches(item, part),
              );
            }
            if (key === "not") {
              return !matches(item, expected as Record<string, unknown>);
            }
            return true;
          });
        if (criteria.uid !== undefined) {
          const match = current.messages.find(
            (item) => item.uid === Number(criteria.uid),
          );
          if (!match) return [];
          // imapflow returns UIDs when uid:true, sequence numbers otherwise
          return options?.uid === false ? [sequenceOf(match)] : [match.uid];
        }
        if (criteria.or && !criteria.not) {
          const ors = criteria.or as Array<{ header: Record<string, string> }>;
          const headersOnly = ors.every(
            (part) => Object.keys(part).length === 1 && "header" in part,
          );
          if (headersOnly) {
            const needles = ors.map((o) =>
              String(Object.values(o.header)[0]).replace(/^<|>$/g, ""),
            );
            return current.messages
              .filter((item) =>
                [item.messageId, item.inReplyTo, item.references].some((id) =>
                  needles.some((n) => id?.replace(/^<|>$/g, "").includes(n)),
                ),
              )
              .map((item) => item.uid);
          }
        }
        return current.messages
          .filter((item) => matches(item, criteria))
          .map((item) => item.uid);
      },
    ),
    fetchOne: vi.fn(
      async (
        range: string,
        query?: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        // Like download: a UID range resolves by UID, otherwise by position.
        fetchQueries.push(query);
        const message = options?.uid
          ? current.messages.find((item) => item.uid === Number(range))
          : current.messages[Number(range) - 1];
        if (!message) return null;
        fetchedUids.push(message.uid);
        return fetched(message, query);
      },
    ),
    messageMove: vi.fn(async (uid: string, folder: string) => {
      const index = current.messages.findIndex(
        (item) => item.uid === Number(uid),
      );
      if (index < 0) return;
      const [moved] = current.messages.splice(index, 1);
      const target = mailboxes.find((candidate) => candidate.path === folder);
      target?.messages.push(moved);
    }),
    messageFlagsAdd: vi.fn(async () => true),
    messageFlagsRemove: vi.fn(async () => true),
    messageDelete: vi.fn(async (range: string, options?: { uid?: boolean }) => {
      // UID DELETE flags \Deleted; the fake expunges immediately.
      const ids = range
        .split(",")
        .map((part) => part.split(":").map(Number))
        .flatMap(([start, end]) =>
          end === undefined || Number.isNaN(end)
            ? [start]
            : Array.from(
                { length: end - start + 1 },
                (_, index) => start + index,
              ),
        );
      const keep = options?.uid
        ? current.messages.filter((item) => !ids.includes(item.uid))
        : current.messages.filter((_, index) => !ids.includes(index + 1));
      current.messages.splice(0, current.messages.length, ...keep);
      return true;
    }),
    mailboxCreate: vi.fn(async (path: string) => {
      if (!mailboxes.some((mailbox) => mailbox.path === path)) {
        mailboxes.push({ path, messages: [] });
      }
    }),
    append: vi.fn(
      async (path: string, content: Buffer, flags?: string[], _date?: Date) => {
        const mailbox = mailboxes.find((candidate) => candidate.path === path);
        if (!mailbox) throw new Error(`mailbox ${path} not found`);
        const nextUid = Math.max(0, ...mailbox.messages.map((m) => m.uid)) + 1;
        const raw = content.toString("utf8");
        const header = (name: string) => {
          const match = raw.match(new RegExp(`^${name}: (.*)$`, "im"));
          return match?.[1].trim();
        };
        const rewritten = rewriteDraftHeaders && path === "Drafts";
        // Bridge strips the stamped per-file X-Attachment-Id from draft
        // parts, so files list under derived ids there.
        const stripped = path === "Drafts" ? stripAttachmentIds(raw) : raw;
        const storedRaw = rewritten
          ? rewriteRawThreading(stripped, nextUid)
          : stripped;
        mailbox.messages.push({
          uid: nextUid,
          date: new Date().toISOString(),
          from: header("From"),
          to: header("To"),
          cc: header("Cc"),
          bcc: header("Bcc"),
          subject: header("Subject"),
          html: storedRaw.split("\r\n\r\n").slice(1).join("\r\n\r\n"),
          messageId: header("Message-ID"),
          inReplyTo: rewritten ? undefined : header("In-Reply-To"),
          references: rewritten
            ? `<fake-internal-${nextUid}@protonmail.internalid>`
            : header("References"),
          raw: storedRaw,
        });
        appended.push({ path, flags: flags ?? [] });
        return mailboxSupportsUidPlus ? { uid: nextUid } : undefined;
      },
    ),
    fetch: vi.fn((range: string, query?: Record<string, unknown>) => {
      fetchQueries.push(query);
      const [start, end] = range.split(":");
      const last =
        end === "*" ? current.messages.length : Number.parseInt(end, 10);
      const page = current.messages.slice(Number.parseInt(start, 10) - 1, last);
      return {
        async *[Symbol.asyncIterator]() {
          for (const item of page) yield fetched(item, query);
        },
      };
    }),
    fetchAll: vi.fn(
      async (
        range: number[] | string,
        query?: Record<string, unknown>,
        options?: { uid?: boolean },
      ) => {
        fetchQueries.push(query);
        const uids = Array.isArray(range)
          ? range
          : String(range)
              .split(",")
              .flatMap((part) => {
                const [start, end] = part.split(":").map(Number);
                if (Number.isNaN(end)) return [start];
                const last = part.endsWith("*")
                  ? Math.max(...current.messages.map((m) => m.uid), 0)
                  : end;
                return Array.from(
                  { length: last - start + 1 },
                  (_, index) => start + index,
                );
              });
        return uids.flatMap((uid) => {
          const found = options?.uid
            ? current.messages.find((item) => item.uid === uid)
            : current.messages[uid - 1];
          if (!found) return [];
          fetchedUids.push(found.uid);
          return [fetched(found, query)];
        });
      },
    ),
    download: vi.fn(
      async (range: string, _part?: string, options?: { uid?: boolean }) => {
        const pick = options?.uid
          ? current.messages.find((item) => item.uid === Number(range))
          : current.messages[Number(range) - 1];
        const content = (async function* () {
          if (pick) yield Buffer.from(pick.raw ?? rawOf(pick));
        })();
        return { content };
      },
    ),
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
    appended,
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

    expect(fake.opened).not.toHaveLength(0);
    for (const folder of fake.opened) expect(folder).toBe("INBOX");
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

    const envelopeQueries = fake.fetchQueries.filter(
      (query) => query && "envelope" in query,
    );
    expect(envelopeQueries.length).toBeGreaterThan(0);
    for (const query of envelopeQueries) {
      expect(query).toMatchObject({ bodyStructure: true });
    }
  });

  it("fetches References in the simple listing so a depth-2 reply keeps the root thread id", async () => {
    const root = "<r@x>";
    const rootThreadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, { messageId: root }),
          message(2, "2026-01-02", "owner@example.com", undefined, {
            messageId: "<m1@x>",
            inReplyTo: root,
            references: root,
          }),
          message(3, "2026-01-03", "alice@x", undefined, {
            messageId: "<m2@x>",
            inReplyTo: "<m1@x>",
            references: `${root} <m1@x>`,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const listing = await new ImapProvider(config).getMessagesWithPagination(
      {},
    );
    for (const query of fake.fetchQueries) {
      if (!query || !("envelope" in query)) continue;
      expect(query).toMatchObject({ headers: ["references"] });
    }
    const listedM2 = listing.messages.find((m) => m.subject === "message 3");
    expect(listedM2?.threadId).toBe(rootThreadId);

    const thread = await new ImapProvider(config).getThread(rootThreadId);
    expect(thread.messages).toHaveLength(3);
  });

  it("searches only the \\All mailbox when the server lists one", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
      {
        path: "All Mail",
        specialUse: "\\All",
        messages: [
          message(40, "2026-07-01", undefined, undefined, {
            subject: "invoice July",
          }),
          message(41, "2026-08-01", undefined, undefined, {
            subject: "invoice August",
          }),
        ],
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

    expect([...new Set(fake.opened)]).toEqual(["All Mail"]);
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("All Mail", 41),
      qualifiedId("All Mail", 40),
    ]);
  });
  it("serves a folder-scoped page from the folder-delta sync", async () => {
    const fake = createFakeClient([
      {
        path: "Archive",
        messages: [message(1, "2026-01-01"), message(2, "2026-02-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      folderId: "Archive",
      maxResults: 20,
    });

    expect(fake.opened).toEqual(["Archive"]);
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("Archive", 1),
      qualifiedId("Archive", 2),
    ]);
    expect(result.nextPageToken).toBeUndefined();
  });

  it("routes the archive view to the Archive folder sync page", async () => {
    const fake = createFakeClient([
      {
        path: "Archive",
        messages: [message(1, "2026-01-01"), message(2, "2026-02-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getThreadsWithQuery({
      query: { type: "archive" },
      maxResults: 20,
    });

    expect(fake.opened).toEqual(["Archive"]);
    expect(result.threads.map((t) => t.id)).toHaveLength(2);
    expect(result.nextPageToken).toBeUndefined();
  });

  it("routes the trash view to the Trash folder sync page", async () => {
    const fake = createFakeClient([
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [message(1, "2026-01-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getThreadsWithQuery({
      query: { type: "trash" },
      maxResults: 20,
    });

    expect(fake.opened).toEqual(["Trash"]);
    expect(result.threads).toHaveLength(1);
  });

  it("untrashThread on a flagged-\\All server also searches Trash", async () => {
    const root = "<trashed@x>";
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, {
            messageId: "<other@x>",
          }),
        ],
      },
      {
        path: "All Mail",
        specialUse: "\\All",
        messages: [
          message(30, "2026-09-01", "alice@x", undefined, {
            messageId: root,
          }),
        ],
      },
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [
          message(30, "2026-09-01", "alice@x", undefined, {
            messageId: root,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const messages = await new ImapProvider(config).getThreadMessages(
      buildThreadId(undefined, undefined, root),
      { includeTrashAndJunk: true },
    );

    // The global mailbox omits deleted mail, so the thread resolves from
    // Trash even on a flagged-\All server.
    expect(fake.opened).toEqual(expect.arrayContaining(["All Mail", "Trash"]));
    expect(messages.map((m) => m.headers["message-id"])).toContain(root);
  });

  it("fans out over real folders on Bridge, never returning All Mail view ids", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            subject: "invoice January",
          }),
          message(2, "2026-03-01", undefined, undefined, {
            subject: "invoice March",
          }),
        ],
      },
      { path: "Sent", specialUse: "\\Sent", messages: [] },
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      {
        path: "All Mail",
        messages: [message(40, "2026-07-01"), message(41, "2026-08-01")],
      },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      {
        path: "Labels/Work",
        messages: [message(41, "2026-08-01")],
      },
      {
        path: "Starred",
        specialUse: "\\Flagged",
        messages: [message(41, "2026-08-01")],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      query: "invoice",
    });

    expect(fake.opened).not.toContain("All Mail");
    expect(fake.opened).not.toContain("Labels/Work");
    expect(fake.opened).not.toContain("Starred");
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 2),
      qualifiedId("INBOX", 1),
    ]);
  });

  it("skips Noselect, Labels, and Starred views in fan-out without an all-mail mailbox", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            subject: "invoice January",
          }),
        ],
      },
      {
        path: "Labels",
        flags: ["\\Noselect"],
        messages: [message(2, "2026-02-01")],
      },
      {
        path: "Labels/Work",
        messages: [message(1, "2026-01-01")],
      },
      {
        path: "Starred",
        specialUse: "\\Flagged",
        messages: [message(1, "2026-01-01")],
      },
      {
        path: "Archive",
        messages: [
          message(10, "2026-04-01", undefined, undefined, {
            subject: "invoice April",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      query: "invoice",
    });

    expect([...new Set(fake.opened)]).toEqual(["INBOX", "Archive"]);
    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("Archive", 10),
      qualifiedId("INBOX", 1),
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
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            subject: "invoice January",
          }),
          message(2, "2026-03-01", undefined, undefined, {
            subject: "invoice March",
          }),
        ],
      },
      {
        path: "Archive",
        messages: [
          message(10, "2026-02-01", undefined, undefined, {
            subject: "invoice February",
          }),
          message(11, "2026-04-01", undefined, undefined, {
            subject: "invoice April",
          }),
        ],
      },
      {
        path: "Sent",
        messages: [
          message(20, "2026-05-01", undefined, undefined, {
            subject: "invoice May",
          }),
        ],
      },
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

    expect([...new Set(fake.opened)]).toEqual(["INBOX", "Archive", "Sent"]);
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

  // Live: Bridge answers `false` to the old `{and: [...]}` form, which read
  // as an empty mailbox. Multi-term queries now use imapflow's native AND.
  it("intersects multi-term queries instead of returning nothing", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-03-01", "alice@x", undefined, {
            subject: "invoice March",
          }),
          message(2, "2026-02-01", "bob@x", undefined, {
            subject: "invoice February",
          }),
          message(3, "2026-01-01", "alice@x", undefined, {
            subject: "receipt January",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    expect(fake.client.search).not.toHaveBeenCalled();
    const result = await new ImapProvider(config).getMessagesWithPagination({
      query: "from:alice@x invoice",
    });

    expect(result.messages.map((m) => m.id)).toEqual([qualifiedId("INBOX", 1)]);
    const sent = (
      fake.client.search as unknown as { mock: { calls: unknown[][] } }
    ).mock.calls.map((call) => call[0] as Record<string, unknown>);
    expect(sent).not.toHaveLength(0);
    for (const criteria of sent) expect(criteria).not.toHaveProperty("and");
  });

  it("ANDs two body phrases without the unsupported and operator", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-03-01", undefined, undefined, {
            subject: "invoice receipt combined",
          }),
          message(2, "2026-02-01", undefined, undefined, {
            subject: "invoice only",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      query: "invoice receipt",
    });

    expect(result.messages.map((m) => m.id)).toEqual([qualifiedId("INBOX", 1)]);
  });

  it("filters has:attachment client-side on populated attachment metadata", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-03-01", undefined, undefined, {
            subject: "invoice with file",
            hasAttachment: true,
          }),
          message(2, "2026-02-01", undefined, undefined, {
            subject: "invoice without file",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      query: "has:attachment invoice",
    });

    expect(result.messages.map((m) => m.id)).toEqual([qualifiedId("INBOX", 1)]);
  });
  // Live: Bridge UID order is append order, not date order, so a sequence
  // window repeated pages 2-4 and never returned the oldest UIDs.
  it("pages disjointly when UIDs are out of date order", async () => {
    const dates = [
      "2026-09-01",
      "2026-01-05",
      "2026-08-01",
      "2026-01-10",
      "2026-07-01",
      "2026-02-01",
      "2026-06-01",
      "2026-03-01",
      "2026-05-01",
      "2026-04-01",
    ];
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: dates.map((date, index) => message(index + 1, date)),
      },
    ]);
    connection.client = fake.client;

    const seen: string[] = [];
    let pageToken: string | undefined;
    do {
      const page = await new ImapProvider(config).getMessagesWithPagination({
        maxResults: 3,
        ...(pageToken ? { pageToken } : {}),
      });
      seen.push(...page.messages.map((m) => m.id));
      pageToken = page.nextPageToken;
    } while (pageToken);

    // Every UID exactly once, newest-date-first overall.
    expect(seen).toHaveLength(10);
    expect(new Set(seen).size).toBe(10);
    expect(seen).toEqual(
      [1, 3, 5, 7, 9, 10, 8, 6, 4, 2].map((uid) => qualifiedId("INBOX", uid)),
    );
  });

  // Live: dating every match with per-UID SEARCH + FETCH round trips took
  // ~190 s on a 3k-message INBOX. One UID FETCH per listing step instead.
  it("dates and envelopes each page with one batched fetch", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: Array.from({ length: 10 }, (_, index) =>
          message(index + 1, `2026-01-${String(index + 1).padStart(2, "0")}`),
        ),
      },
    ]);
    connection.client = fake.client;

    const page = await new ImapProvider(config).getMessagesWithPagination({
      maxResults: 3,
    });

    expect(page.messages).toHaveLength(3);
    const fetchAll = fake.client.fetchAll as unknown as {
      mock: { calls: unknown[][] };
    };
    const fetchOne = fake.client.fetchOne as unknown as {
      mock: { calls: unknown[][] };
    };
    expect(fetchAll.mock.calls).toHaveLength(2);
    expect(fetchOne.mock.calls).toHaveLength(0);
    const search = fake.client.search as unknown as {
      mock: { calls: unknown[][] };
    };
    expect(
      search.mock.calls.filter(
        ([criteria]) =>
          typeof criteria === "object" &&
          criteria !== null &&
          "uid" in criteria,
      ),
    ).toHaveLength(0);
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

describe("ImapProvider stable draft ids", () => {
  beforeEach(() => {
    connection.client = null;
  });

  function draftsMailbox(): {
    path: string;
    specialUse: string;
    messages: FakeMessage[];
  } {
    return { path: "Drafts", specialUse: "\\Drafts", messages: [] };
  }

  it("createDraft returns a stable id that resolves the draft", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "Hello",
      messageHtml: "<p>Hi</p>",
    });

    expect(id.startsWith("d.")).toBe(true);
    const draft = await provider.getDraft(id);
    expect(draft?.subject).toBe("Hello");
    expect(draft?.textHtml).toContain("Hi");
    expect(draft?.headers["message-id"]).toBeTruthy();
    expect(fake.appended).toEqual([
      { path: "Drafts", flags: ["\\Draft", "\\Seen"] },
    ]);
  });

  it("keeps the same draft id across updates with one copy in Drafts", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "v1",
      messageHtml: "<p>one</p>",
    });
    const first = await provider.getDraft(id);
    const firstMessageId = first?.headers["message-id"];
    expect(firstMessageId).toBeTruthy();

    await provider.updateDraft(id, {
      subject: "v2",
      messageHtml: "<p>two</p>",
    });
    await provider.updateDraft(id, {
      subject: "v3",
      messageHtml: "<p>three</p>",
    });

    const updated = await provider.getDraft(id);
    expect(updated?.subject).toBe("v3");
    expect(updated?.textHtml).toContain("three");
    expect(updated?.headers["message-id"]).toBe(firstMessageId);
    expect(mailboxesOf(fake).Drafts).toHaveLength(1);
  });

  it("ignores a substring Message-ID match when resolving a draft", async () => {
    const fake = createFakeClient([
      {
        path: "Drafts",
        specialUse: "\\Drafts",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            messageId: "<xabc123@example.com>",
            subject: "decoy",
            html: "decoy",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    // The SEARCH needle matches both; only the exact Message-ID resolves.
    await expect(
      provider.getDraft(encodeImapDraftId("<abc123@example.com>")),
    ).resolves.toBeNull();
  });

  it("updateDraft keeps attachments instead of dropping them", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "with file",
      messageHtml: "<p>body</p>",
    });
    await provider.addDraftAttachment(id, {
      id: "file-1",
      filename: "notes.txt",
      mimeType: "text/plain",
      size: 11,
      disposition: "attachment",
      content: Buffer.from("hello world"),
    });

    await provider.updateDraft(id, { subject: "edited" });

    const listed = await provider.getDraftAttachments(id);
    expect(listed?.attachments.map((file) => file.filename)).toEqual([
      "notes.txt",
    ]);
    expect(mailboxesOf(fake).Drafts).toHaveLength(1);
  });

  it("rejects updating or deleting a missing draft", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    const missing = encodeImapDraftId("<gone@example.com>");
    await expect(
      provider.updateDraft(missing, { subject: "x" }),
    ).rejects.toThrow(DraftNotFoundError);
    await expect(provider.deleteDraft(missing)).resolves.toBe(false);
  });

  it("keeps reply threading headers through updates", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "Re: thread",
      messageHtml: "<p>reply</p>",
      replyToMessageId: "<parent@example.com>",
    });

    await provider.updateDraft(id, { subject: "Re: thread (edited)" });

    const updated = await provider.getDraft(id);
    expect(updated?.headers["in-reply-to"]).toBe("<parent@example.com>");
    expect(updated?.threadId).toBe(
      buildThreadId(
        undefined,
        "<parent@example.com>",
        updated?.headers["message-id"],
      ),
    );
  });

  it("createDraft as reply threads the sent message with its parent", async () => {
    const parentMessageId = "<parent-1@example.com>";
    const parentReferences = "<root@example.com> <parent-1@example.com>";
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(7, "2026-01-01", "alice@example.com", "me@example.com", {
            messageId: parentMessageId,
            references: parentReferences,
            subject: "thread",
          }),
        ],
      },
      draftsMailbox(),
    ]);
    connection.client = fake.client;
    const sent: Array<Record<string, unknown>> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as Record<string, unknown>);
        return { messageId: "<smtp-sent@example.com>" };
      });
    try {
      const provider = new ImapProvider(config);
      const parent = await provider.getMessage(qualifiedId("INBOX", 7));
      const { id } = await provider.createDraft({
        to: "alice@example.com",
        subject: "Re: thread",
        messageHtml: "<p>reply</p>",
        replyToMessageId: qualifiedId("INBOX", 7),
      });

      const draft = await provider.getDraft(id);
      expect(draft?.headers["in-reply-to"]).toBe(parentMessageId);
      expect(draft?.headers.references).toBe(
        `${parentReferences} ${parentMessageId}`,
      );
      const { threadId } = await provider.sendDraft(id);
      expect(threadId).toBe(parent.threadId);
      expect(sent[0]).toMatchObject({
        inReplyTo: parentMessageId,
        references: `${parentReferences} ${parentMessageId}`,
      });
    } finally {
      sendSpy.mockRestore();
    }
  });

  it("createDraft reply ignores a proton-internal parent id", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(7, "2026-01-01", "alice@example.com", "me@example.com", {
            messageId: "<abc@protonmail.internalid>",
            subject: "thread",
          }),
        ],
      },
      draftsMailbox(),
    ]);
    connection.client = fake.client;
    const provider = new ImapProvider(config);
    const replyTo = qualifiedId("INBOX", 7);
    const { id } = await provider.createDraft({
      to: "alice@example.com",
      subject: "Re: thread",
      messageHtml: "<p>reply</p>",
      replyToMessageId: replyTo,
    });

    const draft = await provider.getDraft(id);
    expect(draft?.headers["in-reply-to"]).toBe(replyTo);
    expect(draft?.headers.references).toBeUndefined();
  });

  it("sendDraft sends the latest content via SMTP and removes the draft", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const sent: Array<Record<string, unknown>> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as Record<string, unknown>);
        return { messageId: "<smtp-sent@example.com>" };
      });
    try {
      const provider = new ImapProvider(config);
      const { id } = await provider.createDraft({
        to: "friend@example.com",
        subject: "send me",
        messageHtml: "<p>final</p>",
      });
      await provider.updateDraft(id, { subject: "send me (final)" });

      const result = await provider.sendDraft(id);

      expect(sent).toHaveLength(1);
      expect(sent[0]).toMatchObject({
        to: "friend@example.com",
        subject: "send me (final)",
      });
      expect(String(sent[0]?.html)).toContain("final");
      expect(result.messageId).toBe("<smtp-sent@example.com>");
      expect(typeof result.threadId).toBe("string");
      expect(mailboxesOf(fake).Drafts).toHaveLength(0);
      await expect(provider.getDraft(id)).resolves.toBeNull();
    } finally {
      sendSpy.mockRestore();
    }
  });

  it("sendDraft refuses to send a missing draft", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async () => ({ messageId: "x" }));
    try {
      await expect(
        new ImapProvider(config).sendDraft(
          encodeImapDraftId("<gone@example.com>"),
        ),
      ).rejects.toThrow(DraftNotFoundError);
      expect(sendSpy).not.toHaveBeenCalled();
    } finally {
      sendSpy.mockRestore();
    }
  });

  it("deleteDraft removes the draft by stable id", async () => {
    const fake = createFakeClient([
      {
        path: "Drafts",
        specialUse: "\\Drafts",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            messageId: "<bye@example.com>",
            subject: "bye",
            html: "bye",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).deleteDraft(
        encodeImapDraftId("<bye@example.com>"),
      ),
    ).resolves.toBe(true);
    expect(mailboxesOf(fake).Drafts).toHaveLength(0);
  });

  it("getDraftReferenceForMessage finds the stable id by message id", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "Hello",
      messageHtml: "<p>Hi</p>",
    });
    const draft = await provider.getDraft(id);

    await expect(
      provider.getDraftReferenceForMessage(draft?.id ?? "missing"),
    ).resolves.toEqual({ id });
    await expect(
      provider.getDraftReferenceForMessage(qualifiedId("Drafts", 999)),
    ).resolves.toBeNull();
  });

  it("getDrafts lists stable ids that deleteDraft accepts", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;

    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "Hello",
      messageHtml: "<p>Hi</p>",
    });

    const listed = await provider.getDrafts();
    expect(listed.map((draft) => draft.id)).toEqual([id]);
    await expect(provider.deleteDraft(listed[0].id)).resolves.toBe(true);
    expect(mailboxesOf(fake).Drafts).toHaveLength(0);
  });

  it("addDraftAttachment stores the file on the draft", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "with file",
      messageHtml: "<p>body</p>",
    });

    const added = await provider.addDraftAttachment(id, {
      id: "file-1",
      filename: "notes.txt",
      mimeType: "text/plain",
      size: 11,
      disposition: "attachment",
      content: Buffer.from("hello world"),
    });

    // The fake strips X-Attachment-Id like Bridge, so the returned id is the
    // derived id the draft lists — the one removal matches on.
    expect(added.attachmentId).toBe("notes.txt:text/plain:11");
    expect(added.attachments).toHaveLength(1);
    expect(added.attachments[0]).toMatchObject({
      id: "notes.txt:text/plain:11",
      filename: "notes.txt",
      mimeType: "text/plain",
    });
    const draft = await provider.getDraft(id);
    expect(draft?.subject).toBe("with file");
    expect(mailboxesOf(fake).Drafts).toHaveLength(1);
  });

  it("removeDraftAttachment drops only that file", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "two files",
      messageHtml: "<p>body</p>",
    });
    const first = await provider.addDraftAttachment(id, {
      id: "file-1",
      filename: "a.txt",
      mimeType: "text/plain",
      size: 1,
      disposition: "attachment",
      content: Buffer.from("a"),
    });
    const second = await provider.addDraftAttachment(id, {
      id: "file-2",
      filename: "b.txt",
      mimeType: "text/plain",
      size: 1,
      disposition: "attachment",
      content: Buffer.from("b"),
    });

    const removed = await provider.removeDraftAttachment(
      id,
      first.attachmentId,
    );

    expect(removed.attachments.map((file) => file.id)).toEqual([
      second.attachmentId,
    ]);
    expect(mailboxesOf(fake).Drafts).toHaveLength(1);
    await expect(
      provider.removeDraftAttachment(id, first.attachmentId),
    ).rejects.toThrow(/no longer on the draft/);
  });

  it("sendDraft delivers attachments through SMTP", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const sent: Array<Record<string, unknown>> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as Record<string, unknown>);
        return { messageId: "<smtp-sent@example.com>" };
      });
    try {
      const provider = new ImapProvider(config);
      const { id } = await provider.createDraft({
        to: "friend@example.com",
        subject: "send file",
        messageHtml: "<p>body</p>",
      });
      await provider.addDraftAttachment(id, {
        id: "file-1",
        filename: "notes.txt",
        mimeType: "text/plain",
        size: 11,
        disposition: "attachment",
        content: Buffer.from("hello world"),
      });

      await provider.sendDraft(id);

      const attachments = sent[0]?.attachments as Array<{
        filename: string;
        content: Buffer;
      }>;
      expect(attachments).toHaveLength(1);
      expect(attachments[0].filename).toBe("notes.txt");
      expect(attachments[0].content.toString()).toBe("hello world");
      expect(mailboxesOf(fake).Drafts).toHaveLength(0);
    } finally {
      sendSpy.mockRestore();
    }
  });

  it("startDraftAttachmentUpload templates the whole message for chunking", async () => {
    const fake = createFakeClient([draftsMailbox()]);
    connection.client = fake.client;
    const provider = new ImapProvider(config);
    const { id } = await provider.createDraft({
      to: "friend@example.com",
      subject: "big file",
      messageHtml: "<p>body</p>",
    });
    await provider.addDraftAttachment(id, {
      id: "file-1",
      filename: "a.txt",
      mimeType: "text/plain",
      size: 1,
      disposition: "attachment",
      content: Buffer.from("a"),
    });

    const upload = await provider.startDraftAttachmentUpload(id, {
      id: "file-2",
      filename: "big.bin",
      mimeType: "application/octet-stream",
      size: 4 * 1024 * 1024,
      disposition: "attachment",
    });

    expect(upload.type).toBe("gmail-message");
    if (upload.type !== "gmail-message") return;
    expect(upload.sessionUri.startsWith("imap:")).toBe(true);
    const slots = upload.parts.filter((part) => part.type === "attachment");
    expect(slots.map((slot) => slot.attachmentId).sort()).toEqual([
      "a.txt:text/plain:1",
      "file-2",
    ]);
    expect(upload.totalBytes).toBeGreaterThan(4 * 1024 * 1024);
  });

  it("sendDraft threads on the parent when Bridge rewrote the stored draft", async () => {
    const parentMessageId = "<parent-bridge@example.com>";
    const parentReferences = "<root@example.com>";
    const fake = createFakeClient(
      [
        {
          path: "INBOX",
          messages: [
            message(7, "2026-01-01", "alice@example.com", "me@example.com", {
              messageId: parentMessageId,
              references: parentReferences,
              subject: "thread",
            }),
          ],
        },
        draftsMailbox(),
      ],
      { rewriteDraftHeaders: true },
    );
    connection.client = fake.client;
    const sent: Array<Record<string, unknown>> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as Record<string, unknown>);
        return { messageId: "<smtp-sent@example.com>" };
      });
    try {
      const provider = new ImapProvider(config);
      const parent = await provider.getMessage(qualifiedId("INBOX", 7));
      const { id } = await provider.createDraft({
        to: "alice@example.com",
        subject: "Re: thread",
        messageHtml: "<p>reply</p>",
        replyToMessageId: qualifiedId("INBOX", 7),
      });
      // The id carries the parent Message-ID past Bridge's rewrite.
      expect(id.split(".")).toHaveLength(3);
      await provider.updateDraft(id, { subject: "Re: thread (edited)" });
      const stored = await provider.getDraft(id);
      expect(stored?.headers["in-reply-to"]).toBeUndefined();

      const { threadId } = await provider.sendDraft(id);

      expect(threadId).toBe(parent.threadId);
      expect(sent[0]).toMatchObject({
        inReplyTo: parentMessageId,
        references: `${parentReferences} ${parentMessageId}`,
      });
    } finally {
      sendSpy.mockRestore();
    }
  });

  it("sendDraft falls back to stored headers when the parent is gone", async () => {
    const fake = createFakeClient(
      [{ path: "INBOX", messages: [] }, draftsMailbox()],
      { rewriteDraftHeaders: true },
    );
    connection.client = fake.client;
    const sent: Array<Record<string, unknown>> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as Record<string, unknown>);
        return { messageId: "<smtp-sent@example.com>" };
      });
    try {
      const provider = new ImapProvider(config);
      const { id } = await provider.createDraft({
        to: "friend@example.com",
        subject: "Hello",
        messageHtml: "<p>Hi</p>",
      });
      const { threadId } = await provider.sendDraft(id);
      expect(typeof threadId).toBe("string");
      expect(sent).toHaveLength(1);
    } finally {
      sendSpy.mockRestore();
    }
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

  it("pages disjointly by date offset instead of refetching the newest messages", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-05-01", undefined, undefined, {
            hasAttachment: true,
          }),
          message(2, "2026-01-01", undefined, undefined, {
            hasAttachment: true,
          }),
          message(3, "2026-04-01", undefined, undefined, {
            hasAttachment: true,
          }),
          message(4, "2026-02-01", undefined, undefined, {
            hasAttachment: true,
          }),
          message(5, "2026-03-01", undefined, undefined, {
            hasAttachment: true,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const first = await new ImapProvider(config).getMessagesWithAttachments({
      maxResults: 3,
    });
    expect(first.nextPageToken).toBe("3");
    expect(first.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 1),
      qualifiedId("INBOX", 3),
      qualifiedId("INBOX", 5),
    ]);

    const second = await new ImapProvider(config).getMessagesWithAttachments({
      maxResults: 3,
      pageToken: first.nextPageToken,
    });
    expect(second.nextPageToken).toBeUndefined();
    expect(second.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 4),
      qualifiedId("INBOX", 2),
    ]);
  });

  it("returns no next page token when the window covers every message", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            hasAttachment: true,
          }),
          message(2, "2026-02-01", undefined, undefined, {
            hasAttachment: true,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithAttachments({
      maxResults: 3,
    });

    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 2),
      qualifiedId("INBOX", 1),
    ]);
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

  it("rethrows an unrelated error that merely mentions 'not found'", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    fake.client.mailboxOpen = vi.fn(async () => {
      throw new Error("mailbox INBOX not found");
    });
    connection.client = fake.client;
    const provider = new ImapProvider(config);
    // Well-formed but unknown: resolution reaches the failing mailboxOpen.
    const ghost = buildThreadId(undefined, undefined, "<ghost@x>");

    await expect(provider.archiveThread(ghost, "owner@x.com")).rejects.toThrow(
      /not found/i,
    );
    await expect(
      provider.trashThread(ghost, "owner@x.com", "user"),
    ).rejects.toThrow(/not found/i);
  });

  it("getThreadMessages throws for an unknown thread", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getThreadMessages("no-such-thread"),
    ).rejects.toThrow(/not found/i);
  });
  it("getMessage throws ImapMessageNotFoundError for a vanished message", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getMessage(qualifiedId("INBOX", 999)),
    ).rejects.toBeInstanceOf(ImapMessageNotFoundError);
  });

  it("getMessagesBatch fetches each folder's UIDs", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            subject: "one",
          }),
          message(2, "2026-01-02", undefined, undefined, {
            subject: "two",
          }),
        ],
      },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(9, "2026-01-03", undefined, undefined, {
            subject: "sent",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const messages = await new ImapProvider(config).getMessagesBatch([
      qualifiedId("INBOX", 2),
      qualifiedId("Sent", 9),
    ]);

    expect(messages.map((item) => item.subject).sort()).toEqual([
      "sent",
      "two",
    ]);
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

  it("moveThreadToFolder treats a vanished thread as a no-op", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).moveThreadToFolder(
        "no-such-thread",
        "owner@x.com",
        "Trash",
      ),
    ).resolves.toBeUndefined();
  });
});

describe("ImapProvider attachments", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("getAttachment downloads the MIME part and returns base64 bytes", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    const content = Buffer.from("attachment bytes here");
    fake.client.download = vi.fn(async () => ({
      content: (async function* () {
        yield content.subarray(0, 10);
        yield content.subarray(10);
      })(),
    }));
    connection.client = fake.client;

    const result = await new ImapProvider(config).getAttachment(
      qualifiedId("INBOX", 1),
      "2",
    );

    expect(fake.client.download).toHaveBeenCalledWith("1", "2", {
      uid: true,
    });
    expect(result).toEqual({
      data: content.toString("base64"),
      size: content.length,
    });
    expect(fake.opened).toEqual(["INBOX"]);
  });

  it("getAttachmentStream streams the MIME part bytes", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    const content = Buffer.from("streamed bytes");
    fake.client.download = vi.fn(async () => ({
      content: (async function* () {
        yield content;
      })(),
    }));
    connection.client = fake.client;

    const stream = await new ImapProvider(config).getAttachmentStream(
      qualifiedId("INBOX", 1),
      "2",
    );
    const bytes = Buffer.from(await new Response(stream).arrayBuffer());

    expect(bytes.equals(content)).toBe(true);
    expect(fake.client.download).toHaveBeenCalledWith("1", "2", {
      uid: true,
    });
  });

  it("getAttachmentStream rejects on an aborted signal without downloading", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    fake.client.download = vi.fn();
    connection.client = fake.client;

    const controller = new AbortController();
    controller.abort();
    await expect(
      new ImapProvider(config).getAttachmentStream(
        qualifiedId("INBOX", 1),
        "2",
        controller.signal,
      ),
    ).rejects.toThrow();
    expect(fake.client.download).not.toHaveBeenCalled();
  });

  it("sendEmailWithHtml marks base64 attachments so bytes survive SMTP", async () => {
    const fake = createFakeClient([{ path: "INBOX", messages: [] }]);
    connection.client = fake.client;
    // Bytes that are not valid UTF-8: sent as text they corrupt, decoded
    // from base64 they round-trip.
    const original = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a]);
    const sent: Array<Record<string, unknown>> = [];
    const sendSpy = vi
      .spyOn(mailModule, "sendSmtpEmail")
      .mockImplementation(async (_config, options) => {
        sent.push(options as Record<string, unknown>);
        return { messageId: "<smtp-sent@example.com>" };
      });
    try {
      await new ImapProvider(config).sendEmailWithHtml({
        to: "friend@example.com",
        subject: "photo",
        messageHtml: "<p>see attached</p>",
        attachments: [
          {
            filename: "photo.jpg",
            content: original.toString("base64"),
            contentType: "image/jpeg",
          },
        ],
      });

      expect(sent).toHaveLength(1);
      const attachments = sent[0]?.attachments as Array<{
        encoding: string;
        content: string;
      }>;
      expect(attachments).toHaveLength(1);
      expect(attachments[0]?.encoding).toBe("base64");
      expect(
        Buffer.from(attachments[0]?.content ?? "", "base64").equals(original),
      ).toBe(true);
    } finally {
      sendSpy.mockRestore();
    }
  });
});

describe("ImapProvider starring", () => {
  beforeEach(() => {
    connection.client = null;
  });
  it("starMessage adds \\Flagged in the message's own folder without moving it", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(1, "2026-01-01"), message(2, "2026-02-01")],
      },
      { path: "Starred", messages: [] },
    ]);
    const flagged: Array<{ uid: number; folder: string; flag: string }> = [];
    fake.client.messageFlagsAdd = vi.fn(
      async (uid: string, flags: string[]) => {
        flagged.push({
          uid: Number(uid),
          folder: fake.opened.at(-1) || "",
          flag: flags[0],
        });
      },
    );
    const moved: unknown[] = [];
    fake.client.messageMove = vi.fn(async (...args: unknown[]) => {
      moved.push(args);
    });
    fake.client.messageCopy = vi.fn(async (...args: unknown[]) => {
      moved.push(args);
    });
    connection.client = fake.client;

    await new ImapProvider(config).starMessage(qualifiedId("INBOX", 2));

    expect(flagged).toEqual([{ uid: 2, folder: "INBOX", flag: "\\Flagged" }]);
    expect(moved).toEqual([]);
    expect(fake.opened).toEqual(["INBOX"]);
  });

  it("markMessagesStarredState removes \\Flagged in each message's own folder", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
      { path: "Archive", messages: [message(1, "2026-01-02")] },
    ]);
    const removed: Array<{ uid: number; folder: string; flag: string }> = [];
    fake.client.messageFlagsRemove = vi.fn(
      async (uid: string, flags: string[]) => {
        removed.push({
          uid: Number(uid),
          folder: fake.opened.at(-1) || "",
          flag: flags[0],
        });
      },
    );
    connection.client = fake.client;

    await new ImapProvider(config).markMessagesStarredState(
      [qualifiedId("INBOX", 1), qualifiedId("Archive", 1)],
      false,
    );

    expect(removed).toEqual([
      { uid: 1, folder: "INBOX", flag: "\\Flagged" },
      { uid: 1, folder: "Archive", flag: "\\Flagged" },
    ]);
  });
  it("markMessagesStarredState surfaces a flag failure instead of swallowing it", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    fake.client.messageFlagsAdd = vi.fn(async () => {
      throw new Error("mailbox is read-only");
    });
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).markMessagesStarredState(
        [qualifiedId("INBOX", 1)],
        true,
      ),
    ).rejects.toThrow(/read-only/);
  });

  it("markMessagesReadState surfaces a flag failure instead of swallowing it", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    fake.client.messageFlagsAdd = vi.fn(async () => {
      throw new Error("mailbox is read-only");
    });
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).markMessagesReadState(
        [qualifiedId("INBOX", 1)],
        true,
      ),
    ).rejects.toThrow(/read-only/);
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

describe("ImapProvider Proton labels (Labels/* namespace)", () => {
  beforeEach(() => {
    connection.client = null;
  });

  function protonMailboxes() {
    return [
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "a@x.com", undefined, {
            messageId: "<root@x>",
          }),
        ],
      },
      { path: "Sent", specialUse: "\\Sent", messages: [] },
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      { path: "All Mail", messages: [] },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      { path: "Labels/Work", messages: [] },
      { path: "Starred", specialUse: "\\Flagged", messages: [] },
    ];
  }

  it("lists only Labels/* children as labels", async () => {
    const fake = createFakeClient(protonMailboxes());
    connection.client = fake.client;

    const labels = await new ImapProvider(config).getLabels();

    expect(labels).toEqual([{ id: "Labels/Work", name: "Work", type: "user" }]);
  });

  it("applies a label by COPY into Labels/X without removing INBOX", async () => {
    const fake = createFakeClient(protonMailboxes());
    const copied: Array<{ uid: number; to: string; from: string }> = [];
    fake.client.messageCopy = vi.fn(async (uid: string, to: string) => {
      copied.push({ uid: Number(uid), to, from: fake.opened.at(-1) || "" });
      const inbox = mailboxesOf(fake).INBOX;
      const found = inbox.find((m) => m.uid === Number(uid));
      if (found) mailboxesOf(fake)["Labels/Work"].push(found);
    });
    connection.client = fake.client;

    const result = await new ImapProvider(config).labelMessage({
      messageId: qualifiedId("INBOX", 1),
      labelId: "Labels/Work",
      labelName: "Work",
    });

    expect(result).toEqual({ actualLabelId: "Labels/Work" });
    expect(copied).toEqual([{ uid: 1, to: "Labels/Work", from: "INBOX" }]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(1);
    expect(mailboxesOf(fake)["Labels/Work"]).toHaveLength(1);
  });

  it("creates Labels/X on apply when it does not exist", async () => {
    const fake = createFakeClient(
      protonMailboxes().filter((mailbox) => mailbox.path !== "Labels/Work"),
    );
    fake.client.mailboxCreate = vi.fn(async (path: string) => {
      mailboxesOf(fake)[path] = [];
    });
    const copied: Array<{ uid: number; to: string }> = [];
    fake.client.messageCopy = vi.fn(async (uid: string, to: string) => {
      copied.push({ uid: Number(uid), to });
    });
    connection.client = fake.client;

    await new ImapProvider(config).labelMessage({
      messageId: qualifiedId("INBOX", 1),
      labelId: "Labels/New",
      labelName: "New",
    });

    expect(fake.client.mailboxCreate).toHaveBeenCalledWith("Labels/New");
    expect(copied).toEqual([{ uid: 1, to: "Labels/New" }]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(1);
  });
  it("reports label support on the Bridge layout only", async () => {
    const bridge = createFakeClient(protonMailboxes());
    connection.client = bridge.client;
    await expect(new ImapProvider(config).supportsLabels()).resolves.toBe(true);

    const plain = createFakeClient([
      { path: "INBOX", messages: [] },
      { path: "Archive", messages: [] },
    ]);
    connection.client = plain.client;
    await expect(new ImapProvider(config).supportsLabels()).resolves.toBe(
      false,
    );
  });

  it("removes a label by expunging only from Labels/X, keeping INBOX", async () => {
    const root = message(1, "2026-01-01", "a@x.com", undefined, {
      messageId: "<root@x>",
    });
    const labeled = message(9, "2026-01-01", "a@x.com", undefined, {
      messageId: "<root@x>",
    });
    const fake = createFakeClient(
      protonMailboxes().map((mailbox) =>
        mailbox.path === "INBOX"
          ? { ...mailbox, messages: [root] }
          : mailbox.path === "Labels/Work"
            ? { ...mailbox, messages: [labeled] }
            : mailbox,
      ),
    );
    const deleted: Array<{ uid: number; folder: string }> = [];
    fake.client.messageDelete = vi.fn(async (uid: string) => {
      deleted.push({ uid: Number(uid), folder: fake.opened.at(-1) || "" });
      const list = mailboxesOf(fake)["Labels/Work"];
      const index = list.findIndex((m) => m.uid === Number(uid));
      if (index >= 0) list.splice(index, 1);
    });
    connection.client = fake.client;

    await new ImapProvider(config).removeThreadLabel(
      buildThreadId(undefined, undefined, "<root@x>"),
      "Labels/Work",
    );

    expect(deleted).toEqual([{ uid: 9, folder: "Labels/Work" }]);
    expect(mailboxesOf(fake)["Labels/Work"]).toEqual([]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(1);
  });

  it("keeps MOVE semantics for labelMessage on servers without a Labels namespace", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
      { path: "Projects", messages: [] },
    ]);
    const moved: Array<{ uid: number; to: string }> = [];
    fake.client.messageMove = vi.fn(async (uid: string, to: string) => {
      moved.push({ uid: Number(uid), to });
    });
    connection.client = fake.client;

    await new ImapProvider(config).labelMessage({
      messageId: qualifiedId("INBOX", 1),
      labelId: "Projects",
      labelName: null,
    });

    expect(moved).toEqual([{ uid: 1, to: "Projects" }]);
  });

  it("resolves labels by id and by name inside the namespace", async () => {
    const fake = createFakeClient(protonMailboxes());
    connection.client = fake.client;
    const provider = new ImapProvider(config);

    await expect(provider.getLabelById("Labels/Work")).resolves.toEqual({
      id: "Labels/Work",
      name: "Work",
      type: "user",
    });
    await expect(provider.getLabelById("Labels/Missing")).resolves.toBeNull();
    await expect(provider.getLabelByName("Work")).resolves.toEqual({
      id: "Labels/Work",
      name: "Work",
      type: "user",
    });
    await expect(provider.getLabelByName("Missing")).resolves.toBeNull();
  });

  it("createLabel creates Labels/<name> and deleteLabel removes only Labels/*", async () => {
    const fake = createFakeClient(protonMailboxes());
    fake.client.mailboxCreate = vi.fn(async (path: string) => {
      mailboxesOf(fake)[path] = [];
    });
    fake.client.mailboxDelete = vi.fn(async (path: string) => {
      delete mailboxesOf(fake)[path];
    });
    connection.client = fake.client;
    const provider = new ImapProvider(config);

    await expect(provider.createLabel("New")).resolves.toEqual({
      id: "Labels/New",
      name: "New",
      type: "user",
    });
    expect(fake.client.mailboxCreate).toHaveBeenCalledWith("Labels/New");
    await provider.deleteLabel("Labels/Work");
    expect(fake.client.mailboxDelete).toHaveBeenCalledWith("Labels/Work");
    await expect(provider.deleteLabel("INBOX")).rejects.toThrow(/only Labels/i);
  });

  it("getThreadsWithLabel reads the label mailbox", async () => {
    const labeled = message(9, "2026-01-01", "a@x.com", undefined, {
      messageId: "<root@x>",
    });
    const fake = createFakeClient(
      protonMailboxes().map((mailbox) =>
        mailbox.path === "Labels/Work"
          ? { ...mailbox, messages: [labeled] }
          : mailbox,
      ),
    );
    connection.client = fake.client;

    const threads = await new ImapProvider(config).getThreadsWithLabel({
      labelId: "Labels/Work",
    });

    expect(threads).toHaveLength(1);
    expect(threads[0].messages.map((m) => m.id)).toEqual([
      qualifiedId("Labels/Work", 9),
    ]);
  });

  it("removeThreadLabels removes every named label", async () => {
    const labeledWork = message(9, "2026-01-01", "a@x.com", undefined, {
      messageId: "<root@x>",
    });
    const labeledBills = message(11, "2026-01-01", "a@x.com", undefined, {
      messageId: "<root@x>",
    });
    const labeledRoot = message(1, "2026-01-01", "a@x.com", undefined, {
      messageId: "<root@x>",
    });
    const fake = createFakeClient([
      ...protonMailboxes().map((mailbox) =>
        mailbox.path === "INBOX"
          ? { ...mailbox, messages: [labeledRoot] }
          : mailbox.path === "Labels/Work"
            ? { ...mailbox, messages: [labeledWork] }
            : mailbox,
      ),
      { path: "Labels/Bills", messages: [labeledBills] },
    ]);
    const deleted: Array<{ uid: number; folder: string }> = [];
    fake.client.messageDelete = vi.fn(async (uid: string) => {
      const folder = fake.opened.at(-1) || "";
      deleted.push({ uid: Number(uid), folder });
      const list = mailboxesOf(fake)[folder];
      const index = list.findIndex((m) => m.uid === Number(uid));
      if (index >= 0) list.splice(index, 1);
    });
    connection.client = fake.client;

    await new ImapProvider(config).removeThreadLabels(
      buildThreadId(undefined, undefined, "<root@x>"),
      ["Labels/Work", "Labels/Bills"],
    );

    expect(deleted).toEqual([
      { uid: 9, folder: "Labels/Work" },
      { uid: 11, folder: "Labels/Bills" },
    ]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(1);
  });

  it("archiveThreadWithLabel keeps MOVE semantics for folders on Proton", async () => {
    const fake = createFakeClient([
      ...protonMailboxes(),
      { path: "Folders/Archive", messages: [] },
    ]);
    const moved: Array<{ uid: number; to: string; from: string }> = [];
    fake.client.messageMove = vi.fn(async (uid: string, to: string) => {
      moved.push({ uid: Number(uid), to, from: fake.opened.at(-1) || "" });
    });
    connection.client = fake.client;

    await new ImapProvider(config).archiveThreadWithLabel(
      buildThreadId(undefined, undefined, "<root@x>"),
      "owner@example.com",
      "Folders/Archive",
    );

    expect(moved).toEqual([{ uid: 1, to: "Folders/Archive", from: "INBOX" }]);
  });

  it("archiveThreadWithLabel into a Proton label COPYs the label then archives, never MOVEs into Labels/*", async () => {
    const fake = createFakeClient(protonMailboxes());
    const copied: Array<{ uid: number; to: string; from: string }> = [];
    const moved: Array<{ uid: number; to: string; from: string }> = [];
    fake.client.messageCopy = vi.fn(async (uid: string, to: string) => {
      copied.push({ uid: Number(uid), to, from: fake.opened.at(-1) || "" });
      const inbox = mailboxesOf(fake).INBOX;
      const found = inbox.find((m) => m.uid === Number(uid));
      if (found) mailboxesOf(fake)["Labels/Work"].push(found);
    });
    fake.client.messageMove = vi.fn(async (uid: string, to: string) => {
      moved.push({ uid: Number(uid), to, from: fake.opened.at(-1) || "" });
      const inbox = mailboxesOf(fake).INBOX;
      const index = inbox.findIndex((m) => m.uid === Number(uid));
      if (index >= 0) {
        const [item] = inbox.splice(index, 1);
        mailboxesOf(fake)[to].push(item);
      }
    });
    connection.client = fake.client;

    await new ImapProvider(config).archiveThreadWithLabel(
      buildThreadId(undefined, undefined, "<root@x>"),
      "owner@example.com",
      "Labels/Work",
    );

    expect(copied).toEqual([{ uid: 1, to: "Labels/Work", from: "INBOX" }]);
    expect(moved).toEqual([{ uid: 1, to: "Archive", from: "INBOX" }]);
    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Archive).toHaveLength(1);
    expect(mailboxesOf(fake)["Labels/Work"]).toHaveLength(1);
  });

  it("stale view ids on Bridge resolve as not found, so archive is a no-op", async () => {
    const fake = createFakeClient(protonMailboxes());
    connection.client = fake.client;

    const stale = `fallback:${encodeImapMessageId("All Mail", 40)}`;
    await new ImapProvider(config).archiveThread(stale, "owner@example.com");

    expect(fake.client.messageMove).not.toHaveBeenCalled();
    expect(fake.opened).not.toContain("All Mail");
  });

  it("never SELECTs the top-level Folders/Labels containers on Bridge", async () => {
    const fake = createFakeClient(protonMailboxes());
    connection.client = fake.client;

    await new ImapProvider(config).getThreadMessages(
      buildThreadId(undefined, undefined, "<root@x>"),
    );

    expect(fake.opened).not.toContain("Folders");
    expect(fake.opened).not.toContain("Labels");
  });

  it("skips one failing folder and still resolves the thread from the rest", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, {
            messageId: "<root@x>",
          }),
          message(2, "2026-01-02", "owner@example.com", undefined, {
            messageId: "<m1@x>",
            inReplyTo: "<root@x>",
            references: "<root@x>",
          }),
        ],
      },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
    ]);
    const open = fake.client.mailboxOpen;
    fake.client.mailboxOpen = vi.fn(async (path: string) => {
      if (path === "Archive")
        throw Object.assign(new Error("Mailbox Archive gone"), {
          responseStatus: "NO",
        });
      return open(path);
    });
    connection.client = fake.client;

    const messages = await new ImapProvider(config).getThreadMessages(
      buildThreadId(undefined, undefined, "<root@x>"),
    );

    expect(messages).toHaveLength(2);
  });

  it("getMessagesWithPagination skips a NO folder instead of failing", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", undefined, undefined, {
            subject: "invoice January",
          }),
          message(2, "2026-03-01", undefined, undefined, {
            subject: "invoice March",
          }),
        ],
      },
      { path: "Archive", messages: [message(10, "2026-02-01")] },
    ]);
    const open = fake.client.mailboxOpen;
    fake.client.mailboxOpen = vi.fn(async (path: string) => {
      if (path === "Archive")
        throw Object.assign(new Error("Mailbox Archive gone"), {
          responseStatus: "NO",
        });
      return open(path);
    });
    connection.client = fake.client;

    const result = await new ImapProvider(config).getMessagesWithPagination({
      inboxOnly: false,
      query: "invoice",
    });

    expect(result.messages.map((m) => m.id)).toEqual([
      qualifiedId("INBOX", 2),
      qualifiedId("INBOX", 1),
    ]);
  });

  it("getFolders on Bridge hides All Mail, Starred, Labels/* and containers", async () => {
    const fake = createFakeClient(protonMailboxes());
    connection.client = fake.client;

    const folders = await new ImapProvider(config).getFolders();
    const ids = folders.map((folder) => folder.id);

    expect(ids).toEqual(
      expect.arrayContaining(["INBOX", "Sent", "Drafts", "Archive"]),
    );
    for (const id of ids) {
      expect(id).not.toBe("All Mail");
      expect(id).not.toBe("Starred");
      expect(id).not.toBe("Folders");
      expect(id).not.toBe("Labels");
      expect(id.startsWith("Labels/")).toBe(false);
    }
    await expect(
      new ImapProvider(config).moveThreadToFolder(
        buildThreadId(undefined, undefined, "<root@x>"),
        "owner@example.com",
        "All Mail",
      ),
    ).rejects.toThrow(/view mailbox/);
  });
});

describe("ImapProvider thread id agreement across listing, getMessage and thread ops", () => {
  beforeEach(() => {
    connection.client = null;
  });

  const root = "<r@x>";
  const rootThreadId = buildThreadId(undefined, undefined, root);

  // Thread: R (root) <- M1 (replies to R) <- M2 (replies to M1, cites R M1).
  function threadMailboxes() {
    return [
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", "owner@example.com", {
            messageId: root,
          }),
          message(2, "2026-01-02", "owner@example.com", "alice@x", {
            messageId: "<m1@x>",
            inReplyTo: root,
            references: root,
          }),
          message(3, "2026-01-03", "alice@x", "owner@example.com", {
            messageId: "<m2@x>",
            inReplyTo: "<m1@x>",
            references: `${root} <m1@x>`,
          }),
        ],
      },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
      { path: "Sent", specialUse: "\\Sent", messages: [] },
    ];
  }

  function bridgeThreadMailboxes() {
    const plain = threadMailboxes();
    return [
      ...plain,
      { path: "Sent", specialUse: "\\Sent", messages: [] },
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      {
        path: "All Mail",
        messages: plain[0].messages.map((m, index) => ({
          ...m,
          uid: 11 + index,
        })),
      },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      { path: "Labels/Work", messages: [] },
      { path: "Starred", specialUse: "\\Flagged", messages: [] },
    ];
  }

  it("getThreads rejects a view folder id on Bridge", async () => {
    const fake = createFakeClient(bridgeThreadMailboxes());
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getThreads("All Mail"),
    ).rejects.toThrow(/view mailbox/i);
    expect(fake.opened).not.toContain("All Mail");
  });

  function spreadBridgeMailboxes() {
    return [
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, {
            messageId: "<r@x>",
          }),
        ],
      },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(2, "2026-01-02", "owner@example.com", undefined, {
            messageId: "<m1@x>",
            inReplyTo: "<r@x>",
            references: "<r@x>",
          }),
        ],
      },
      {
        path: "Folders/Projets/AFPS",
        messages: [
          message(3, "2026-01-03", "alice@x", undefined, {
            messageId: "<m2@x>",
            inReplyTo: "<m1@x>",
            references: "<r@x> <m1@x>",
          }),
        ],
      },
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      {
        path: "All Mail",
        messages: [
          message(11, "2026-01-01", "alice@x", undefined, {
            messageId: "<r@x>",
          }),
          message(12, "2026-01-02", "owner@example.com", undefined, {
            messageId: "<m1@x>",
            inReplyTo: "<r@x>",
            references: "<r@x>",
          }),
          message(13, "2026-01-03", "alice@x", undefined, {
            messageId: "<m2@x>",
            inReplyTo: "<m1@x>",
            references: "<r@x> <m1@x>",
          }),
        ],
      },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      { path: "Labels/Work", messages: [] },
      { path: "Starred", specialUse: "\\Flagged", messages: [] },
    ];
  }

  it("gives a depth-2 reply the same thread id in the listing and in getMessage", async () => {
    const fake = createFakeClient(threadMailboxes());
    connection.client = fake.client;

    const threads = await new ImapProvider(config).getThreads();
    const listedM2 = threads
      .flatMap((thread) => thread.messages)
      .find((m) => m.subject === "message 3")?.threadId;
    const fullM2 = (
      await new ImapProvider(config).getMessage(qualifiedId("INBOX", 3))
    ).threadId;

    expect(listedM2).toBe(fullM2);
    expect(fullM2).toBe(rootThreadId);
  });

  it("keeps one thread for R-M1-M2 in the listing", async () => {
    const fake = createFakeClient(threadMailboxes());
    connection.client = fake.client;

    const threads = await new ImapProvider(config).getThreads();

    expect(threads).toHaveLength(1);
    expect(threads[0].messages).toHaveLength(3);
  });

  it("resolves the whole thread from any member id", async () => {
    const fake = createFakeClient(threadMailboxes());
    connection.client = fake.client;
    const listedM2 = buildThreadId(`${root} <m1@x>`, "<m1@x>", "<m2@x>");

    const fromM2 = await new ImapProvider(config).getThreadMessages(listedM2);
    const fromRoot = await new ImapProvider(config).getThreadMessages(
      rootThreadId,
    );

    expect(fromM2.map((m) => m.subject).sort()).toEqual([
      "message 1",
      "message 2",
      "message 3",
    ]);
    expect(fromRoot.map((m) => m.subject).sort()).toEqual([
      "message 1",
      "message 2",
      "message 3",
    ]);
  });

  it("archiveThread on the root moves every member out of INBOX", async () => {
    const fake = createFakeClient(threadMailboxes());
    connection.client = fake.client;

    await new ImapProvider(config).archiveThread(
      rootThreadId,
      "owner@example.com",
    );

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Archive).toHaveLength(3);
  });

  it("archiveThread on a Bridge layout moves real-folder copies, never All Mail views", async () => {
    const fake = createFakeClient(bridgeThreadMailboxes());
    connection.client = fake.client;

    await new ImapProvider(config).archiveThread(
      rootThreadId,
      "owner@example.com",
    );

    const moves = (
      fake.client.messageMove as unknown as { mock: { calls: unknown[][] } }
    ).mock.calls;
    expect(moves.length).toBeGreaterThan(0);
    for (const [, to] of moves) expect(to).toBe("Archive");
    expect(fake.opened).not.toContain("All Mail");
    expect(fake.opened).not.toContain("Labels/Work");
    expect(fake.opened).not.toContain("Starred");
    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake)["All Mail"]).toHaveLength(3);
  });

  // Gmail semantics: archiving removes the INBOX placement only. The Sent
  // copy never leaves Sent; restoring moves only the archived copy back.
  it("archiveThread moves only the INBOX member of an INBOX+Sent thread", async () => {
    const root = "<sent-copy@x>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", "me@example.com", {
            messageId: root,
          }),
        ],
      },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(2, "2026-01-01", "me@example.com", "alice@x", {
            messageId: "<sent@x>",
            inReplyTo: root,
            references: root,
          }),
        ],
      },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
    ]);
    connection.client = fake.client;

    await new ImapProvider(config).archiveThread(threadId, "me@example.com");

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Archive).toHaveLength(1);
    expect(mailboxesOf(fake).Sent).toHaveLength(1);
  });

  it("unarchiveThread moves only the Archive member back to INBOX", async () => {
    const root = "<archived@x>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(2, "2026-01-01", "me@example.com", "alice@x", {
            messageId: "<sent@x>",
            inReplyTo: root,
            references: root,
          }),
        ],
      },
      {
        path: "Archive",
        specialUse: "\\Archive",
        messages: [
          message(7, "2026-01-01", "alice@x", "me@example.com", {
            messageId: root,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    await new ImapProvider(config).unarchiveThread(threadId);

    expect(mailboxesOf(fake).Archive).toEqual([]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(1);
    expect(mailboxesOf(fake).Sent).toHaveLength(1);
  });

  it("markSpam moves only received members, never Sent or Drafts", async () => {
    const root = "<spam-target@x>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", "me@example.com", {
            messageId: root,
          }),
        ],
      },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(2, "2026-01-01", "me@example.com", "alice@x", {
            messageId: "<sent@x>",
            inReplyTo: root,
            references: root,
          }),
        ],
      },
      {
        path: "Drafts",
        specialUse: "\\Drafts",
        messages: [
          message(3, "2026-01-01", "me@example.com", "alice@x", {
            messageId: "<draft@x>",
            inReplyTo: root,
            references: root,
          }),
        ],
      },
      { path: "Junk", specialUse: "\\Junk", messages: [] },
    ]);
    connection.client = fake.client;

    await new ImapProvider(config).markSpam(threadId);

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Junk).toHaveLength(1);
    expect(mailboxesOf(fake).Sent).toHaveLength(1);
    expect(mailboxesOf(fake).Drafts).toHaveLength(1);
  });

  it("markNotSpam moves only Spam members back to INBOX", async () => {
    const root = "<not-spam@x>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(2, "2026-01-01", "me@example.com", "alice@x", {
            messageId: "<sent@x>",
            inReplyTo: root,
            references: root,
          }),
        ],
      },
      {
        path: "Junk",
        specialUse: "\\Junk",
        messages: [
          message(9, "2026-01-01", "alice@x", "me@example.com", {
            messageId: root,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    await new ImapProvider(config).markNotSpam(threadId);

    expect(mailboxesOf(fake).Junk).toEqual([]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(1);
    expect(mailboxesOf(fake).Sent).toHaveLength(1);
  });

  it("untrashThread restores the sent copy to Sent and received mail to INBOX", async () => {
    const root = "<trashed-thread@x>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient([
      { path: "INBOX", messages: [] },
      { path: "Sent", specialUse: "\\Sent", messages: [] },
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [
          message(5, "2026-01-01", "me@example.com", "alice@x", {
            messageId: "<sent@x>",
            inReplyTo: root,
            references: root,
          }),
          message(6, "2026-01-01", "alice@x", "me@example.com", {
            messageId: root,
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    await new ImapProvider(config).untrashThread(threadId);

    expect(mailboxesOf(fake).Trash).toEqual([]);
    expect(mailboxesOf(fake).Sent.map((m) => m.uid)).toEqual([5]);
    expect(mailboxesOf(fake).INBOX.map((m) => m.uid)).toEqual([6]);
  });
  it("a non-Proton server with a user folder named All Mail still fans out to INBOX", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", "owner@example.com", {
            messageId: root,
          }),
          message(2, "2026-01-02", "owner@example.com", "alice@x", {
            messageId: "<m1@x>",
            inReplyTo: root,
            references: root,
          }),
        ],
      },
      { path: "All Mail", messages: [] },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
    ]);
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getThreadMessages(rootThreadId),
    ).resolves.toHaveLength(2);
  });

  it("a lookalike header that merely contains the root is not in the thread", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, {
            messageId: root,
          }),
          message(2, "2026-01-02", "bob@x", undefined, {
            messageId: "<xr@x>",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const messages = await new ImapProvider(config).getThreadMessages(
      rootThreadId,
    );

    expect(messages.map((m) => m.subject)).toEqual(["message 1"]);
  });

  it("rejects a short crafted root before any SEARCH", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    connection.client = fake.client;
    const shortId = `t.${Buffer.from("a", "utf8").toString("base64url")}`;

    await expect(
      new ImapProvider(config).getThreadMessages(shortId),
    ).rejects.toThrow(/implausible/);
    expect(fake.client.search).not.toHaveBeenCalled();
  });

  it("resolves a thread with one SEARCH per real folder on Bridge, never All Mail", async () => {
    const plain = threadMailboxes();
    const fake = createFakeClient([
      ...plain,
      { path: "Sent", specialUse: "\\Sent", messages: [] },
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      {
        path: "All Mail",
        messages: plain[0].messages.map((m, index) => ({
          ...m,
          uid: 11 + index,
        })),
      },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      { path: "Labels/Work", messages: [] },
      { path: "Starred", specialUse: "\\Flagged", messages: [] },
    ]);
    connection.client = fake.client;

    const messages = await new ImapProvider(config).getThreadMessages(
      rootThreadId,
    );

    expect(messages).toHaveLength(3);
    expect(fake.opened).not.toContain("All Mail");
    expect(fake.opened).not.toContain("Labels/Work");
    expect(fake.opened).not.toContain("Starred");
    for (const message of messages) {
      const folder = Buffer.from(
        message.id.slice(message.id.indexOf("@") + 1),
        "base64url",
      ).toString("utf8");
      expect(folder).not.toBe("All Mail");
      expect(folder.startsWith("Labels/")).toBe(false);
      expect(folder).not.toBe("Starred");
    }
    const headerSearches = (
      fake.client.search as unknown as { mock: { calls: unknown[][] } }
    ).mock.calls.filter((call) =>
      ["header", "or"].some(
        (key) => key in (call[0] as Record<string, unknown>),
      ),
    );
    expect(headerSearches).toHaveLength(fake.opened.length);
  });

  it("resolves a thread spread over INBOX, Sent and Folders/* with real-folder ids only", async () => {
    const fake = createFakeClient(spreadBridgeMailboxes());
    connection.client = fake.client;

    const thread = await new ImapProvider(config).getThread(rootThreadId);

    expect(thread.messages).toHaveLength(3);
    expect(fake.opened).not.toContain("All Mail");
    expect(
      thread.messages
        .map((m) =>
          Buffer.from(m.id.slice(m.id.indexOf("@") + 1), "base64url").toString(
            "utf8",
          ),
        )
        .sort(),
    ).toEqual(["Folders/Projets/AFPS", "INBOX", "Sent"]);
  });

  it("archives a message id obtained from getThread out of its real folder", async () => {
    const fake = createFakeClient(spreadBridgeMailboxes());
    const moved: Array<{ uid: number; to: string; from: string }> = [];
    fake.client.messageMove = vi.fn(async (uid: string, to: string) => {
      moved.push({ uid: Number(uid), to, from: fake.opened.at(-1) || "" });
      const current = mailboxesOf(fake)[fake.opened.at(-1) || ""];
      const index = current.findIndex((m) => m.uid === Number(uid));
      if (index >= 0) {
        const [item] = current.splice(index, 1);
        mailboxesOf(fake)[to].push(item);
      }
    });
    connection.client = fake.client;

    const thread = await new ImapProvider(config).getThread(rootThreadId);
    const sentCopy = thread.messages.find((m) => m.id.startsWith("2@"));
    expect(sentCopy).toBeDefined();
    await new ImapProvider(config).archiveMessage(sentCopy!.id);

    expect(moved).toEqual([{ uid: 2, to: "Archive", from: "Sent" }]);
    expect(mailboxesOf(fake).Sent).toEqual([]);
  });

  it("keeps one thread for Proton's split copies: parent INBOX+Sent, reply Sent+INBOX", async () => {
    const parentId = "<33bc@proton.me>";
    const parentThreadId = buildThreadId(undefined, undefined, parentId);
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(7, "2026-10-30", "me@proton.me", "them@x", {
            messageId: parentId,
            references: "<1ehgou==@protonmail.internalid>",
          }),
          message(9, "2026-10-31", "them@x", "me@proton.me", {
            messageId: "<reply@proton.me>",
            inReplyTo: parentId,
            references: `${parentId} <96tJ@protonmail.internalid>`,
          }),
        ],
      },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(5, "2026-10-30", "me@proton.me", "them@x", {
            messageId: parentId,
          }),
          message(6, "2026-10-31", "me@proton.me", "them@x", {
            messageId: "<reply@proton.me>",
            inReplyTo: parentId,
            references: `${parentId} <96tJ@protonmail.internalid>`,
          }),
        ],
      },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
    ]);
    connection.client = fake.client;

    const threads = await new ImapProvider(config).getMessagesWithPagination(
      {},
    );
    const ids = new Set(threads.messages.map((m) => m.threadId));
    expect(ids).toEqual(new Set([parentThreadId]));

    const resolved = await new ImapProvider(config).getThread(parentThreadId);
    expect(resolved.messages).toHaveLength(4);
  });
});

describe("ImapProvider Proton label edge cases", () => {
  beforeEach(() => {
    connection.client = null;
  });

  const root = "<r@x>";

  function bridgeWithLabels(
    labelMailboxes: Array<{ path: string; messages: FakeMessage[] }>,
  ) {
    return [
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, { messageId: root }),
          message(2, "2026-01-02", "owner@example.com", undefined, {
            messageId: "<m1@x>",
            inReplyTo: root,
            references: root,
          }),
          message(3, "2026-01-03", "alice@x", undefined, {
            messageId: "<m2@x>",
            inReplyTo: "<m1@x>",
            references: `${root} <m1@x>`,
          }),
        ],
      },
      { path: "Sent", specialUse: "\\Sent", messages: [] },
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      { path: "All Mail", messages: [] },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      ...labelMailboxes,
      { path: "Starred", specialUse: "\\Flagged", messages: [] },
    ];
  }

  it("removing a label that was never created is a no-op", async () => {
    const fake = createFakeClient(bridgeWithLabels([]));
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).removeThreadLabel(
        buildThreadId(undefined, undefined, root),
        "Labels/Ghost",
      ),
    ).resolves.toBeUndefined();
    expect(fake.opened).not.toContain("Labels/Ghost");
  });

  it("removal matches Message-ID exactly, not by substring", async () => {
    const fake = createFakeClient(
      bridgeWithLabels([
        {
          path: "Labels/Work",
          messages: [
            message(21, "2026-01-01", "alice@x", undefined, {
              messageId: root,
            }),
            message(22, "2026-01-02", "bob@x", undefined, {
              messageId: "<xr@x>",
            }),
          ],
        },
      ]),
    );
    const deleted: Array<{ uid: number; folder: string }> = [];
    fake.client.messageDelete = vi.fn(async (uid: string) => {
      deleted.push({ uid: Number(uid), folder: fake.opened.at(-1) || "" });
      const list = mailboxesOf(fake)["Labels/Work"];
      const index = list.findIndex((m) => m.uid === Number(uid));
      if (index >= 0) list.splice(index, 1);
    });
    connection.client = fake.client;

    await new ImapProvider(config).removeThreadLabel(
      buildThreadId(undefined, undefined, root),
      "Labels/Work",
    );

    expect(deleted).toEqual([{ uid: 21, folder: "Labels/Work" }]);
    expect(mailboxesOf(fake)["Labels/Work"].map((m) => m.uid)).toEqual([22]);
    expect(mailboxesOf(fake).INBOX).toHaveLength(3);
  });

  it("a plain folder named Labels keeps MOVE semantics", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, { messageId: root }),
        ],
      },
      { path: "Labels", messages: [] },
      { path: "Projects", messages: [] },
    ]);
    const moved: Array<{ uid: number; to: string }> = [];
    fake.client.messageMove = vi.fn(async (uid: string, to: string) => {
      moved.push({ uid: Number(uid), to });
    });
    connection.client = fake.client;

    await new ImapProvider(config).labelMessage({
      messageId: qualifiedId("INBOX", 1),
      labelId: "Projects",
      labelName: null,
    });

    expect(moved).toEqual([{ uid: 1, to: "Projects" }]);
  });

  it("createLabel with different case returns the existing mailbox id", async () => {
    const fake = createFakeClient(
      bridgeWithLabels([{ path: "Labels/Work", messages: [] }]),
    );
    fake.client.mailboxCreate = vi.fn(async () => undefined);
    connection.client = fake.client;

    await expect(new ImapProvider(config).createLabel("work")).resolves.toEqual(
      {
        id: "Labels/Work",
        name: "Work",
        type: "user",
      },
    );
    expect(fake.client.mailboxCreate).not.toHaveBeenCalled();
  });

  it("message-id mutators skip Bridge view ids and reads throw", async () => {
    const viewId = qualifiedId("All Mail", 40);
    const cases: Array<{
      name: string;
      run: (provider: ImapProvider) => Promise<unknown>;
      throws: boolean;
    }> = [
      {
        name: "archiveMessages",
        run: (p) => p.archiveMessages([viewId]),
        throws: false,
      },
      {
        name: "unarchiveMessages",
        run: (p) => p.unarchiveMessages([viewId]),
        throws: false,
      },
      {
        name: "trashMessages",
        run: (p) => p.trashMessages([viewId]),
        throws: false,
      },
      {
        name: "untrashMessages",
        run: (p) => p.untrashMessages([viewId]),
        throws: false,
      },
      {
        name: "markMessagesReadState",
        run: (p) => p.markMessagesReadState([viewId], true),
        throws: false,
      },
      {
        name: "markMessagesStarredState",
        run: (p) => p.markMessagesStarredState([viewId], true),
        throws: false,
      },
      {
        name: "archiveMessage",
        run: (p) => p.archiveMessage(viewId),
        throws: true,
      },
      {
        name: "starMessage",
        run: (p) => p.starMessage(viewId),
        throws: true,
      },
      {
        name: "getMessageSource",
        run: (p) => p.getMessageSource(viewId),
        throws: true,
      },
      {
        name: "getAttachment",
        run: (p) => p.getAttachment(viewId, "2"),
        throws: true,
      },
    ];

    for (const { name, run, throws } of cases) {
      const fake = createFakeClient(
        bridgeWithLabels([{ path: "Labels/Work", messages: [] }]).map(
          (mailbox) =>
            mailbox.path === "All Mail"
              ? {
                  ...mailbox,
                  messages: [message(40, "2026-07-01", "alice@x")],
                }
              : mailbox,
        ),
      );
      connection.client = fake.client;
      const provider = new ImapProvider(config);

      if (throws) {
        await expect(run(provider), name).rejects.toThrow(/view mailbox/i);
      } else {
        await expect(run(provider), name).resolves.toBeUndefined();
      }
      expect(fake.opened, name).not.toContain("All Mail");
      expect(fake.client.messageMove, name).not.toHaveBeenCalled();
      expect(fake.client.messageFlagsAdd, name).not.toHaveBeenCalled();
      expect(fake.client.messageFlagsRemove, name).not.toHaveBeenCalled();
    }
  });
});

describe("ImapProvider attachment edge cases", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("an absent part id throws instead of returning an empty attachment", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    fake.client.download = vi.fn(async () => ({
      content: (async function* () {})(),
    }));
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getAttachment(qualifiedId("INBOX", 1), "9"),
    ).rejects.toThrow(/not found/i);
    await expect(
      new ImapProvider(config)
        .getAttachmentStream(qualifiedId("INBOX", 1), "9")
        .then(
          async (stream) => {
            for await (const _chunk of stream as unknown as AsyncIterable<Uint8Array>) {
              // drain
            }
            return "resolved";
          },
          (error: Error) => error.message,
        ),
    ).resolves.toMatch(/not found/i);
  });

  it("finds the attachment by UID when uid differs from sequence number", async () => {
    // UID 7 sits at sequence 2; a sequence-number fetch would read the
    // wrong message (or none) for the BODYSTRUCTURE check.
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [message(3, "2026-01-01"), message(7, "2026-01-02")],
      },
    ]);
    connection.client = fake.client;

    const attachment = await new ImapProvider(config).getAttachment(
      qualifiedId("INBOX", 7),
      "2",
    );
    expect(Buffer.from(attachment.data, "base64").toString("utf8")).toContain(
      "body of message 7",
    );
    expect(fake.fetchedUids).toEqual([7]);
    expect(fake.client.download).toHaveBeenCalledWith("7", "2", { uid: true });
  });

  it("a legitimately empty part downloads as empty bytes without throwing", async () => {
    const fake = createFakeClient([
      { path: "INBOX", messages: [message(1, "2026-01-01")] },
    ]);
    fake.client.fetchOne = vi.fn(async () => ({
      bodyStructure: { part: "2", type: "text", subtype: "plain", size: 0 },
    }));
    fake.client.download = vi.fn(async () => ({
      content: (async function* () {})(),
    }));
    connection.client = fake.client;

    await expect(
      new ImapProvider(config).getAttachment(qualifiedId("INBOX", 1), "2"),
    ).resolves.toEqual({ data: "", size: 0 });
  });
});
describe("ImapProvider Bridge shared-item placement (self-sent mail)", () => {
  beforeEach(() => {
    connection.client = null;
  });

  function bridgeLayout(extra: FakeMailbox[]) {
    return [
      ...extra,
      { path: "Drafts", specialUse: "\\Drafts", messages: [] },
      { path: "Archive", specialUse: "\\Archive", messages: [] },
      { path: "Spam", specialUse: "\\Junk", messages: [] },
      { path: "Trash", specialUse: "\\Trash", messages: [] },
      { path: "All Mail", messages: [] },
      { path: "Folders", flags: ["\\Noselect"], messages: [] },
      { path: "Labels", flags: ["\\Noselect"], messages: [] },
      { path: "Starred", specialUse: "\\Flagged", messages: [] },
    ];
  }

  // Bridge models one Proton item with several folder placements: moving any
  // placement relocates the whole item, and moving a Trash copy into Trash
  // deletes it permanently (live-verified on Bridge v3.27.1). A UID keeps
  // naming its item after a move, so a stale placement snapshot still
  // resolves — that is the data-loss path under test.
  function sharedItemMoves(fake: FakeImapTestClient, trashPath: string) {
    const uidToItem = new Map<number, string>();
    const rememberLive = () => {
      for (const messages of Object.values(mailboxesOf(fake))) {
        for (const m of messages) {
          if (m.messageId) uidToItem.set(m.uid, m.messageId);
        }
      }
    };
    rememberLive();
    fake.client.messageMove = vi.fn(async (range: string, to: string) => {
      const from = fake.opened.at(-1) || "";
      const store = mailboxesOf(fake);
      rememberLive();
      const itemId =
        store[from]?.find((m) => m.uid === Number(range))?.messageId ??
        uidToItem.get(Number(range));
      if (!itemId) return;
      const live = Object.values(store).flatMap((messages) =>
        messages.filter((m) => m.messageId === itemId),
      );
      if (live.length === 0) return;
      const inTarget = Object.entries(store).find(
        ([path, messages]) =>
          path === to && messages.some((m) => m.messageId === itemId),
      );
      // The item already lives in the target: same-folder MOVE is a no-op
      // everywhere except Trash, where Bridge deletes it permanently.
      if (inTarget) {
        if (to === trashPath) {
          for (const messages of Object.values(store)) {
            const index = messages.findIndex((m) => m.messageId === itemId);
            if (index >= 0) messages.splice(index, 1);
          }
        }
        return;
      }
      const moved = store[from]?.find((m) => m.messageId === itemId) ?? live[0];
      for (const messages of Object.values(store)) {
        const index = messages.findIndex((m) => m.messageId === itemId);
        if (index >= 0) messages.splice(index, 1);
      }
      store[to]?.push(moved);
    });
  }

  function selfSentThread() {
    const root = "<self-sent@proton.me>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient(
      bridgeLayout([
        {
          path: "INBOX",
          messages: [
            message(2957, "2026-10-31", "me@proton.me", "me@proton.me", {
              messageId: root,
            }),
          ],
        },
        {
          path: "Sent",
          specialUse: "\\Sent",
          messages: [
            message(654, "2026-10-31", "me@proton.me", "me@proton.me", {
              messageId: root,
            }),
          ],
        },
      ]),
    );
    sharedItemMoves(fake, "Trash");
    return { fake, threadId, root };
  }

  it("trashThread moves a self-sent item to Trash once, never deleting it", async () => {
    const { fake, threadId, root } = selfSentThread();
    connection.client = fake.client;

    await new ImapProvider({
      ...config,
      email: "me@proton.me",
    }).trashThread(threadId, "me@proton.me", "user");

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Sent).toEqual([]);
    expect(mailboxesOf(fake).Trash.map((m) => m.messageId)).toEqual([root]);
  });

  it("trashThread on a 3-message self-sent thread leaves every item in Trash", async () => {
    const root = "<self-3@proton.me>";
    const threadId = buildThreadId(undefined, undefined, root);
    const fake = createFakeClient(
      bridgeLayout([
        {
          path: "INBOX",
          messages: [0, 1, 2].map((index) =>
            message(100 + index, "2026-10-31", "me@proton.me", "me@proton.me", {
              messageId: index === 0 ? root : `<self-3-r${index}@proton.me>`,
              ...(index === 0
                ? {}
                : {
                    inReplyTo: root,
                    references: root,
                  }),
            }),
          ),
        },
        {
          path: "Sent",
          specialUse: "\\Sent",
          messages: [0, 1, 2].map((index) =>
            message(200 + index, "2026-10-31", "me@proton.me", "me@proton.me", {
              messageId: index === 0 ? root : `<self-3-r${index}@proton.me>`,
              ...(index === 0
                ? {}
                : {
                    inReplyTo: root,
                    references: root,
                  }),
            }),
          ),
        },
      ]),
    );
    sharedItemMoves(fake, "Trash");
    connection.client = fake.client;

    await new ImapProvider({
      ...config,
      email: "me@proton.me",
    }).trashThread(threadId, "me@proton.me", "user");

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Sent).toEqual([]);
    expect(mailboxesOf(fake).Trash).toHaveLength(3);
  });

  it("untrashThread restores a trashed self-sent item so it resolves again", async () => {
    const { fake, threadId, root } = selfSentThread();
    const provider = new ImapProvider({ ...config, email: "me@proton.me" });
    connection.client = fake.client;
    await provider.trashThread(threadId, "me@proton.me", "user");
    expect(mailboxesOf(fake).Trash.map((m) => m.messageId)).toEqual([root]);

    await provider.untrashThread(threadId);

    expect(mailboxesOf(fake).Trash).toEqual([]);
    const inbox = mailboxesOf(fake).INBOX.map((m) => m.messageId);
    const sent = mailboxesOf(fake).Sent.map((m) => m.messageId);
    expect([...inbox, ...sent]).toEqual([root]);
    const restored = await provider.getThread(threadId);
    expect(restored.messages).toHaveLength(1);
  });

  it("trashing an already-trashed item is a no-op success", async () => {
    const root = "<already-trashed@proton.me>";
    const fake = createFakeClient(
      bridgeLayout([
        { path: "INBOX", messages: [] },
        { path: "Sent", specialUse: "\\Sent", messages: [] },
      ]),
    );
    mailboxesOf(fake).Trash.push(
      message(302, "2026-10-31", "me@proton.me", "me@proton.me", {
        messageId: root,
      }),
    );
    sharedItemMoves(fake, "Trash");
    connection.client = fake.client;
    const provider = new ImapProvider({ ...config, email: "me@proton.me" });

    await provider.trashMessages([qualifiedId("Trash", 302)]);

    expect(mailboxesOf(fake).Trash.map((m) => m.messageId)).toEqual([root]);
  });

  it("archive relocates a self-sent item and unarchive restores it", async () => {
    const { fake, threadId, root } = selfSentThread();
    const provider = new ImapProvider({ ...config, email: "me@proton.me" });
    connection.client = fake.client;

    await provider.archiveThread(threadId, "me@proton.me");

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Sent).toEqual([]);
    expect(mailboxesOf(fake).Archive.map((m) => m.messageId)).toEqual([root]);

    await provider.unarchiveThread(threadId);

    expect(mailboxesOf(fake).Archive).toEqual([]);
    expect(mailboxesOf(fake).INBOX.map((m) => m.messageId)).toEqual([root]);
  });

  it("spam relocates a self-sent item and not-spam restores it", async () => {
    const { fake, threadId, root } = selfSentThread();
    const provider = new ImapProvider({ ...config, email: "me@proton.me" });
    connection.client = fake.client;

    await provider.markSpam(threadId);

    expect(mailboxesOf(fake).INBOX).toEqual([]);
    expect(mailboxesOf(fake).Sent).toEqual([]);
    expect(mailboxesOf(fake).Spam.map((m) => m.messageId)).toEqual([root]);

    await provider.markNotSpam(threadId);

    expect(mailboxesOf(fake).Spam).toEqual([]);
    const inbox = mailboxesOf(fake).INBOX.map((m) => m.messageId);
    const sent = mailboxesOf(fake).Sent.map((m) => m.messageId);
    expect([...inbox, ...sent]).toEqual([root]);
  });
});

describe("ImapProvider scoped provider search", () => {
  beforeEach(() => {
    connection.client = null;
  });

  function scopedMailboxes() {
    return createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, {
            subject: "invoice January",
          }),
        ],
      },
      {
        path: "Trash",
        specialUse: "\\Trash",
        messages: [
          message(7, "2026-02-01", "bob@x", undefined, {
            subject: "invoice February",
          }),
        ],
      },
    ]);
  }

  it("searches only the Trash folder for a trash mailbox scope", async () => {
    const fake = scopedMailboxes();
    connection.client = fake.client;

    const result = await new ImapProvider(config).searchMessages({
      query: "",
      mailboxSearch: {
        mailbox: "trash",
        text: { kind: "text", field: "any", value: "invoice", match: "term" },
      },
    });

    expect([...new Set(fake.opened)]).toEqual(["Trash"]);
    expect(result.messages.map((m) => m.id)).toEqual([qualifiedId("Trash", 7)]);
  });

  it("searches only the Sent folder for a sent mailbox scope", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        messages: [
          message(1, "2026-01-01", "alice@x", undefined, {
            subject: "invoice January",
          }),
        ],
      },
      {
        path: "Sent",
        specialUse: "\\Sent",
        messages: [
          message(3, "2026-03-01", "me@x", undefined, {
            subject: "invoice March",
          }),
        ],
      },
    ]);
    connection.client = fake.client;

    const result = await new ImapProvider(config).searchMessages({
      query: "",
      mailboxSearch: {
        mailbox: "sent",
        text: { kind: "text", field: "any", value: "invoice", match: "term" },
      },
    });

    expect([...new Set(fake.opened)]).toEqual(["Sent"]);
    expect(result.messages.map((m) => m.id)).toEqual([qualifiedId("Sent", 3)]);
  });

  it("uses TEXT (not BODY) for field-any so subject matches hit", async () => {
    const fake = scopedMailboxes();
    connection.client = fake.client;

    await new ImapProvider(config).searchMessages({
      query: "",
      mailboxSearch: {
        mailbox: "trash",
        text: { kind: "text", field: "any", value: "invoice", match: "term" },
      },
    });

    const search = fake.client.search as unknown as {
      mock: { calls: Array<[Record<string, unknown>]> };
    };
    const criteria = search.mock.calls.map(([c]) => c);
    expect(criteria.some((c) => "text" in c)).toBe(true);
    expect(criteria.some((c) => "body" in c)).toBe(false);
  });

  it("scopes legacy searchThreads to the trash folder", async () => {
    const fake = scopedMailboxes();
    connection.client = fake.client;

    const result = await new ImapProvider(config).searchThreads({
      query: "invoice",
      folder: "trash",
    });

    expect([...new Set(fake.opened)]).toEqual(["Trash"]);
    expect(result.threads.flatMap((t) => t.messages.map((m) => m.id))).toEqual([
      qualifiedId("Trash", 7),
    ]);
  });

  it("includes trash and junk in the fan-out when asked", async () => {
    const fake = scopedMailboxes();
    connection.client = fake.client;

    const result = await new ImapProvider(config).searchThreads({
      query: "invoice",
      includeSpamTrash: true,
    });

    expect(new Set(fake.opened)).toEqual(new Set(["INBOX", "Trash"]));
    expect(
      result.threads.flatMap((t) => t.messages.map((m) => m.id)).sort(),
    ).toEqual([qualifiedId("INBOX", 1), qualifiedId("Trash", 7)].sort());
  });
});
describe("ImapProvider.getFolderCounts", () => {
  beforeEach(() => {
    connection.client = null;
  });

  it("tags system folders so count consumers can match INBOX", async () => {
    const fake = createFakeClient([
      {
        path: "INBOX",
        specialUse: "\\Inbox",
        messages: [message(1, "2026-01-01")],
      },
      {
        path: "Work",
        messages: [message(1, "2026-01-01")],
      },
    ]);
    connection.client = fake.client;

    const counts = await new ImapProvider(config).getFolderCounts();

    expect(counts).toEqual([
      { id: "INBOX", name: "INBOX", total: 1, unread: 0, systemType: "INBOX" },
      { id: "Work", name: "Work", total: 1, unread: 0 },
    ]);
  });
});
