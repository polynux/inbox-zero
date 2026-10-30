import { beforeEach, describe, expect, it, vi } from "vitest";
import { decodeMailboxSyncCursor } from "@/utils/email/mailbox-sync";
import { getImapMailboxSyncPage } from "@/utils/imap/local-mail-sync";

vi.mock("@/utils/imap/client", () => ({
  withImapConnection: vi.fn(
    async (_config: unknown, fn: (client: unknown) => Promise<unknown>) =>
      fn(stubClient),
  ),
}));

let mailboxUids: number[];
let uidValidity: number | undefined;

const stubClient = {
  mailboxOpen: vi.fn(async (_folder: string, _options?: unknown) => ({
    exists: mailboxUids.length,
    uidNext: mailboxUids.length > 0 ? Math.max(...mailboxUids) + 1 : 1,
    uidValidity,
  })),
  search: vi.fn(async (criteria: Record<string, unknown>) => {
    if (typeof criteria.uid === "string") {
      const from = Number(criteria.uid.split(":")[0]);
      return mailboxUids.filter((uid) => uid >= from);
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
        date: new Date("2026-01-01T00:00:00Z"),
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

function lastUidOf(cursor: string): number {
  const decoded = decodeMailboxSyncCursor(cursor, "imap");
  return decoded.state.lastUid;
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
    expect(page.upsertedMessages.map((m) => m.id)).toEqual(["1", "2"]);
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

    expect(seen).toEqual(["1", "2", "3", "4", "5"]);
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
    expect(page.upsertedMessages.map((m) => m.id)).toEqual(["1", "2"]);
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
});
