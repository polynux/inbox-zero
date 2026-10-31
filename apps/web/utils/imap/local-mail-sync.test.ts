import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  decodeMailboxSyncCursor,
  encodeMailboxSyncCursor,
} from "@/utils/email/mailbox-sync";
import { getImapMailboxSyncPage } from "@/utils/imap/local-mail-sync";

vi.mock("@/utils/imap/client", () => ({
  withImapConnection: vi.fn(
    async (_config: unknown, fn: (client: unknown) => Promise<unknown>) =>
      fn(stubClient),
  ),
}));

const MESSAGE_DATE = new Date("2026-01-01T00:00:00Z");

let mailboxUids: number[];
let uidValidity: number | undefined;

const stubClient: { mailbox: { path?: string } } & Record<string, unknown> = {
  mailbox: {},
  mailboxOpen: vi.fn(async (folder: string, _options?: unknown) => {
    stubClient.mailbox = { path: folder };
    return {
      path: folder,
      exists: mailboxUids.length,
      uidNext: mailboxUids.length > 0 ? Math.max(...mailboxUids) + 1 : 1,
      uidValidity,
    };
  }),
  search: vi.fn(async (criteria: Record<string, unknown>) => {
    if (typeof criteria.uid === "string") {
      const from = Number(criteria.uid.split(":")[0]);
      const matched = mailboxUids.filter((uid) => uid >= from);
      // RFC 3501: `n:*` matches the highest UID even when n is above it.
      return matched.length > 0 ? matched : mailboxUids.slice(-1);
    }
    if (criteria.since instanceof Date) {
      return MESSAGE_DATE >= criteria.since ? [...mailboxUids] : [];
    }
    return [...mailboxUids];
  }),
  fetchOne: vi.fn(async (seq: string) => {
    const uid = mailboxUids[Number(seq) - 1];
    if (uid === undefined) return false;
    return {
      uid,
      envelope: {
        from: [{ name: "Sender", address: "sender@example.com" }],
        to: [{ name: "Me", address: "me@example.com" }],
        subject: `Subject ${uid}`,
        date: MESSAGE_DATE,
        messageId: `<${uid}@example.com>`,
      },
      flags: new Set(["\\Seen"]),
    };
  }),
};

const baseInput = {
  config: {
    email: "me@example.com",
    emailAccountId: "account-1",
    imapHost: "localhost",
    imapPort: 993,
    imapSecurity: "tls" as const,
    password: "secret",
    smtpHost: "localhost",
    smtpPort: 587,
    smtpSecurity: "starttls" as const,
    username: "me@example.com",
  },
  logger: { info: vi.fn(), warn: vi.fn(), error: vi.fn() } as never,
  emailAccountId: "account-1",
  limit: 2,
};

function qualifiedId(uid: number): string {
  return `${uid}@${Buffer.from("INBOX", "utf8").toString("base64url")}`;
}

function lastUidOf(cursor: string): number {
  const decoded = decodeMailboxSyncCursor(cursor, "imap");
  return decoded.state.lastUid;
}

function cursorAt(lastUid: number): string {
  return encodeMailboxSyncCursor({
    version: 1,
    provider: "imap",
    state: {
      emailAccountId: "account-1",
      folder: "INBOX",
      uidValidity,
      lastUid,
    },
  });
}

beforeEach(() => {
  vi.clearAllMocks();
  mailboxUids = [1, 2, 3, 4, 5];
  uidValidity = 42;
});

describe("getImapMailboxSyncPage", () => {
  it("snapshots oldest-first with reset on the initial page", async () => {
    const page = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2025-01-01T00:00:00Z"),
    });

    expect(page.reset).toBe(true);
    expect(page.hasMore).toBe(true);
    expect(page.upsertedMessages.map((m) => m.id)).toEqual([
      qualifiedId(1),
      qualifiedId(2),
    ]);
    expect(lastUidOf(page.cursor)).toBe(2);
  });

  it("pages without gaps or duplicates until the mailbox is drained", async () => {
    const seen: string[] = [];
    let cursor: string | undefined;
    let resetCount = 0;
    for (let i = 0; i < 5; i++) {
      const page = await getImapMailboxSyncPage({
        ...baseInput,
        after: new Date("2025-01-01T00:00:00Z"),
        cursor,
      });
      if (page.reset) resetCount++;
      seen.push(...page.upsertedMessages.map((m) => m.id));
      cursor = page.cursor;
      if (!page.hasMore) break;
    }

    expect(seen).toEqual([
      qualifiedId(1),
      qualifiedId(2),
      qualifiedId(3),
      qualifiedId(4),
      qualifiedId(5),
    ]);
    expect(resetCount).toBe(1);
    expect(cursor).toBeDefined();
    const drained = await getImapMailboxSyncPage({
      ...baseInput,
      cursor,
    });
    expect(drained.upsertedMessages).toEqual([]);
    expect(drained.hasMore).toBe(false);
    expect(drained.reset).toBe(false);
  });

  it("resets when UIDVALIDITY changes", async () => {
    const first = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2025-01-01T00:00:00Z"),
    });
    uidValidity = 43;
    mailboxUids = [1, 2, 3];

    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: first.cursor,
    });
    expect(page.reset).toBe(true);
    expect(page.upsertedMessages.map((m) => m.id)).toEqual([
      qualifiedId(1),
      qualifiedId(2),
    ]);
  });

  it("reports an expunged UID as a deleted message id", async () => {
    const first = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2025-01-01T00:00:00Z"),
    });

    mailboxUids = [1, 2, 4, 5];

    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: first.cursor,
    });

    expect(page.deletedMessageIds).toEqual([qualifiedId(3)]);
    expect(lastUidOf(page.cursor)).toBe(5);
  });

  it("reports deletions even when no new mail arrived", async () => {
    const first = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2025-01-01T00:00:00Z"),
      limit: 5,
    });

    mailboxUids = [1, 2, 3, 4];

    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: first.cursor,
    });

    expect(page.upsertedMessages).toEqual([]);
    expect(page.deletedMessageIds).toEqual([qualifiedId(5)]);
    expect(lastUidOf(page.cursor)).toBe(5);
  });

  it("reports no deletions when the mailbox is unchanged", async () => {
    const first = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2025-01-01T00:00:00Z"),
    });

    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: first.cursor,
    });

    expect(page.deletedMessageIds).toEqual([]);
  });

  it("rejects cursors from another account or folder", async () => {
    const first = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2025-01-01T00:00:00Z"),
    });

    await expect(
      getImapMailboxSyncPage({
        ...baseInput,
        cursor: first.cursor,
        folderId: "Archive",
      }),
    ).rejects.toThrow();
  });

  it("returns only UIDs above the cursor", async () => {
    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: cursorAt(3),
    });

    expect(page.reset).toBe(false);
    expect(page.upsertedMessages.map((m) => m.id)).toEqual([
      qualifiedId(4),
      qualifiedId(5),
    ]);
    expect(page.hasMore).toBe(false);
    expect(lastUidOf(page.cursor)).toBe(5);
  });

  it("reports nothing for an idle mailbox instead of its newest message", async () => {
    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: cursorAt(5),
    });

    expect(page.upsertedMessages).toEqual([]);
    expect(page.hasMore).toBe(false);
    expect(lastUidOf(page.cursor)).toBe(5);
  });

  it("never moves the cursor backward", async () => {
    const page = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: cursorAt(10),
    });

    expect(page.upsertedMessages).toEqual([]);
    expect(lastUidOf(page.cursor)).toBe(10);
  });

  it("does not replay the mailbox when the initial window is empty", async () => {
    const page = await getImapMailboxSyncPage({
      ...baseInput,
      after: new Date("2026-06-01T00:00:00Z"),
    });

    expect(page.upsertedMessages).toEqual([]);
    expect(lastUidOf(page.cursor)).toBe(5);

    const next = await getImapMailboxSyncPage({
      ...baseInput,
      cursor: page.cursor,
    });
    expect(next.upsertedMessages).toEqual([]);
    expect(next.hasMore).toBe(false);
  });
});
