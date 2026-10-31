import { beforeEach, describe, expect, it, vi } from "vitest";
import { encodeImapMessageId } from "@/utils/imap/message-id";

type Credential = {
  id: string;
  lastSeenUid: number | null;
  uidValidity: bigint | null;
  pollFailureUid: number | null;
  pollFailureCount: number;
  pollLockedUntil: Date | null;
};

const { state, prismaFake } = vi.hoisted(() => {
  const state = {
    credential: null as Credential | null,
    accounts: [{ id: "email-account-1" }],
    mailboxUids: [] as number[],
    mailboxOverride: null as { uidNext?: number; uidValidity?: bigint } | null,
    processed: [] as string[],
    seen: [] as Array<{
      messageId: string;
      message: { labelIds?: string[]; threadId: string };
      context: unknown;
    }>,
    processHook: null as null | ((messageId: string) => Promise<void>),
    fromAddress: "news@example.com",
    labelIds: ["INBOX", "UNREAD"] as string[],
    validationFails: false,
    redisConfigured: false,
    lockSkipped: false,
  };

  const prismaFake = {
    imapCredential: {
      findFirst: async () =>
        state.credential ? { ...state.credential } : null,
      update: async ({ data }: { data: Partial<Credential> }) => {
        Object.assign(state.credential!, data);
        return state.credential;
      },
      updateMany: async ({
        where,
        data,
      }: {
        where: {
          lastSeenUid?: { lt: number };
          pollLockedUntil?: Date;
          OR?: unknown[];
        };
        data: Partial<Credential>;
      }) => {
        const credential = state.credential!;
        // Claim: either no lock or an expired one.
        if (where.OR) {
          if (
            credential.pollLockedUntil &&
            credential.pollLockedUntil > new Date()
          )
            return { count: 0 };
          credential.pollLockedUntil = data.pollLockedUntil ?? null;
          return { count: 1 };
        }
        // Release: only the run that still owns the lock clears it.
        if (where.pollLockedUntil) {
          if (
            credential.pollLockedUntil?.getTime() !==
            where.pollLockedUntil.getTime()
          )
            return { count: 0 };
          credential.pollLockedUntil = null;
          return { count: 1 };
        }
        // Cursor advance, guarded so it never moves backward.
        if (where.lastSeenUid) {
          if (
            credential.lastSeenUid === null ||
            credential.lastSeenUid >= where.lastSeenUid.lt
          )
            return { count: 0 };
          Object.assign(credential, data);
          return { count: 1 };
        }
        throw new Error("unexpected updateMany");
      },
    },
    emailAccount: { findMany: async () => state.accounts },
  };

  return { state, prismaFake };
});

const client = {
  mailboxOpen: async () =>
    state.mailboxOverride ?? {
      uidNext: Math.max(0, ...state.mailboxUids) + 1,
      uidValidity: 1n,
    },
};

vi.mock("@/utils/prisma", () => ({ default: prismaFake }));
vi.mock("@/utils/imap/client", () => ({
  withImapConnection: async (
    _config: unknown,
    fn: (client: unknown) => Promise<unknown>,
  ) => fn(client),
}));
vi.mock("@/utils/imap/credential", () => ({
  getImapCredentials: async () => ({
    email: "me@example.com",
    emailAccountId: "email-account-1",
    imapHost: "localhost",
    imapPort: 993,
    imapSecurity: "tls",
    smtpHost: "localhost",
    smtpPort: 587,
    smtpSecurity: "starttls",
    username: "me@example.com",
    password: "secret",
  }),
}));
vi.mock("@/utils/imap/message", () => ({
  searchImapMessages: async (
    _client: unknown,
    criteria: { uid?: string; all?: boolean },
  ) => {
    if (criteria.all) return [...state.mailboxUids].reverse();
    const uids = state.mailboxUids;
    if (!criteria.uid) return [...uids].reverse();
    const from = Number(criteria.uid.split(":")[0]);
    const matched = uids.filter((uid) => uid >= from);
    // RFC 3501: `n:*` matches the highest UID even when n is above it.
    return (matched.length > 0 ? matched : uids.slice(-1)).reverse();
  },
  fetchMessageByUid: async (_client: unknown, uid: number) => ({
    id: `${uid}@${Buffer.from("INBOX").toString("base64url")}`,
    labelIds: state.labelIds,
    threadId: "t.thread",
    headers: { from: state.fromAddress },
  }),
}));
vi.mock("@/utils/email/provider", () => ({
  createEmailProvider: async () => ({}),
}));
vi.mock("@/utils/redis/message-processing", () => ({
  withMessageProcessingLock: async (
    _key: unknown,
    fn: () => Promise<unknown>,
  ) => {
    if (state.lockSkipped) return { skipped: true };
    await fn();
    return { skipped: false };
  },
}));
vi.mock("@/utils/webhook/process-history-item", () => ({
  processHistoryItem: async (
    item: {
      messageId: string;
      message: { labelIds?: string[]; threadId: string };
    },
    context: unknown,
  ) => {
    state.processed.push(item.messageId);
    state.seen.push({
      messageId: item.messageId,
      message: {
        labelIds: item.message.labelIds,
        threadId: item.message.threadId,
      },
      context,
    });
    if (state.processHook) await state.processHook(item.messageId);
  },
}));
vi.mock("@/utils/webhook/validate-webhook-account", () => ({
  getWebhookEmailAccount: async () => ({ id: "email-account-1" }),
  validateWebhookAccount: async () =>
    state.validationFails
      ? { success: false, response: new Response() }
      : {
          success: true,
          data: {
            emailAccount: { id: "email-account-1", rules: [] },
            hasAutomationRules: true,
            hasAiAccess: true,
          },
        },
}));
vi.mock("@/env", () => ({
  env: {
    get REDIS_HTTP_URL() {
      return state.redisConfigured ? "http://redis.test" : undefined;
    },
    get REDIS_HTTP_TOKEN() {
      return state.redisConfigured ? "token" : undefined;
    },
  },
}));

import {
  POLL_TIME_BUDGET_MS,
  pollAllImapAccounts,
  pollImapAccount,
} from "@/utils/imap/poll";

function setCursor(
  lastSeenUid: number | null,
  uidValidity: bigint | null = null,
) {
  state.credential = {
    id: "credential-1",
    lastSeenUid,
    uidValidity,
    pollFailureUid: null,
    pollFailureCount: 0,
    pollLockedUntil: null,
  };
}

beforeEach(() => {
  state.credential = null;
  state.accounts = [{ id: "email-account-1" }];
  state.mailboxUids = [1, 2, 3, 4, 5];
  state.mailboxOverride = null;
  state.processed = [];
  state.seen = [];
  state.processHook = null;
  state.fromAddress = "news@example.com";
  state.labelIds = ["INBOX", "UNREAD"];
  state.validationFails = false;
  state.redisConfigured = false;
  state.lockSkipped = false;
  setCursor(null);
});

describe("pollImapAccount", () => {
  it("reports nothing for an idle mailbox instead of its newest message", async () => {
    setCursor(5);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 0,
    });
    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.processed).toEqual([]);
  });

  it("reports and records new messages above the stored UID", async () => {
    setCursor(3);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 2,
    });
    expect(state.credential?.lastSeenUid).toBe(5);
  });

  it("feeds new inbox messages to the shared processor with qualified ids", async () => {
    setCursor(4);

    await pollImapAccount("email-account-1");

    expect(state.seen).toEqual([
      expect.objectContaining({
        messageId: encodeImapMessageId("INBOX", 5),
        message: expect.objectContaining({
          labelIds: ["INBOX", "UNREAD"],
          threadId: "t.thread",
        }),
        context: expect.objectContaining({
          hasAutomationRules: true,
          hasAiAccess: true,
          emailAccount: expect.objectContaining({
            id: "email-account-1",
            account: { provider: "imap" },
          }),
        }),
      }),
    ]);
  });

  it("does not advance the cursor past a failed message", async () => {
    setCursor(3);
    state.processHook = async () => {
      throw new Error("processing failed");
    };

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(state.credential?.lastSeenUid).toBe(3);
    expect(state.credential?.pollFailureUid).toBe(4);
    expect(state.credential?.pollFailureCount).toBe(1);
  });

  it("skips a message that keeps failing so newer mail is not blocked", async () => {
    setCursor(4);
    state.processHook = async () => {
      throw new Error("processing failed");
    };

    await pollImapAccount("email-account-1");
    expect(state.credential?.lastSeenUid).toBe(4);

    await pollImapAccount("email-account-1");
    expect(state.credential?.lastSeenUid).toBe(4);

    await pollImapAccount("email-account-1");

    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.credential?.pollFailureUid).toBeNull();
    expect(state.credential?.pollFailureCount).toBe(0);
  });

  it("keeps the failure count across invocations, not in module memory", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];
    state.processHook = async () => {
      throw new Error("boom");
    };

    const cursors: Array<number | null> = [];
    for (let i = 0; i < 5; i++) {
      // A poll is a fresh invocation in production; the strike counter must
      // live in the database to survive it.
      vi.resetModules();
      const fresh = await import("@/utils/imap/poll");
      await fresh.pollImapAccount("email-account-1");
      cursors.push(state.credential!.lastSeenUid);
    }

    expect(cursors).toEqual([5, 5, 6, 6, 6]);
  });

  it("clears failure strikes when UIDVALIDITY resets the cursor", async () => {
    setCursor(3, 42n);
    state.mailboxOverride = { uidNext: 6, uidValidity: 42n };
    state.processHook = async (messageId) => {
      if (messageId === encodeImapMessageId("INBOX", 4))
        throw new Error("processing failed");
    };

    await pollImapAccount("email-account-1");
    await pollImapAccount("email-account-1");
    expect(state.credential?.pollFailureUid).toBe(4);
    expect(state.credential?.pollFailureCount).toBe(2);

    state.mailboxOverride = { uidNext: 6, uidValidity: 43n };
    await pollImapAccount("email-account-1");
    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.credential?.pollFailureUid).toBeNull();
    expect(state.credential?.pollFailureCount).toBe(0);

    // The renumbered mailbox delivers a fresh UID 4; one failure must not be
    // the third strike that skips it.
    state.credential!.lastSeenUid = 3;
    await pollImapAccount("email-account-1");

    expect(state.credential?.pollFailureUid).toBe(4);
    expect(state.credential?.pollFailureCount).toBe(1);
    expect(state.credential?.lastSeenUid).toBe(3);
  });

  it("advances the cursor per handled message so a stopped run is not replayed", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6, 7, 8];
    state.processHook = async (messageId) => {
      if (messageId === encodeImapMessageId("INBOX", 8))
        throw new Error("boom");
    };

    await pollImapAccount("email-account-1");
    expect(state.credential?.lastSeenUid).toBe(7);

    state.processHook = null;
    await pollImapAccount("email-account-1");

    expect(state.processed.filter((id) => id === "6@SU5CT1g")).toHaveLength(1);
    expect(state.credential?.lastSeenUid).toBe(8);
  });

  it("processes a shared UID once when polls overlap", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];
    const processing = Promise.withResolvers<void>();
    state.processHook = async () => {
      await processing.promise;
    };

    const firstPoll = pollImapAccount("email-account-1");
    const secondPoll = pollImapAccount("email-account-1");
    processing.resolve();
    const [first, second] = await Promise.all([firstPoll, secondPoll]);

    expect(state.processed.filter((id) => id === "6@SU5CT1g")).toHaveLength(1);
    expect(first.newMessages + second.newMessages).toBe(1);
    expect(state.credential?.pollLockedUntil).toBeNull();
  });

  it("does not count a message another run holds the lock for", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];
    state.redisConfigured = true;
    state.lockSkipped = true;

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(state.processed).toEqual([]);
    expect(state.credential?.lastSeenUid).toBe(6);
  });

  it("resumes after an expired claim", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];
    state.credential!.pollLockedUntil = new Date(Date.now() - 60_000);

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(1);
    expect(state.credential?.lastSeenUid).toBe(6);
  });

  it("skips a message the account sent itself but still advances", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];
    state.fromAddress = "Me <Me@Example.com>";

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(state.processed).toEqual([]);
    expect(state.credential?.lastSeenUid).toBe(6);
  });

  it("skips a message labelled SENT like Outlook's sent-items rule", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];
    state.labelIds = ["SENT"];

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(state.processed).toEqual([]);
    expect(state.credential?.lastSeenUid).toBe(6);
  });

  it("stops starting messages once the deadline has passed", async () => {
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6, 7];

    const result = await pollImapAccount("email-account-1", undefined, {
      deadline: Date.now() - 1,
    });

    expect(result.newMessages).toBe(0);
    expect(state.processed).toEqual([]);
    expect(state.credential?.lastSeenUid).toBe(5);
  });

  it("skips accounts without automation rules or AI access", async () => {
    setCursor(3);
    state.validationFails = true;

    const result = await pollImapAccount("email-account-no-rules");

    expect(result).toEqual({
      emailAccountId: "email-account-no-rules",
      newMessages: 0,
    });
    expect(state.processed).toEqual([]);
    expect(state.credential?.lastSeenUid).toBe(3);
  });

  it("baselines a null cursor: first poll reports 0 new", async () => {
    setCursor(null);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 0,
    });
    expect(state.credential?.lastSeenUid).toBe(5);
  });

  it("clears failure strikes when a first poll baselines the cursor", async () => {
    setCursor(null, 42n);
    state.credential!.pollFailureUid = 4;
    state.credential!.pollFailureCount = 2;
    state.mailboxOverride = { uidNext: 6, uidValidity: 42n };

    await pollImapAccount("email-account-1");

    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.credential?.pollFailureUid).toBeNull();
    expect(state.credential?.pollFailureCount).toBe(0);
  });

  it("keeps the baseline unset when the server omits UIDNEXT", async () => {
    setCursor(null);
    state.mailboxOverride = {};

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 0,
    });
    expect(state.credential?.lastSeenUid).toBeNull();
  });

  it("resets the cursor to the baseline when UIDVALIDITY changes", async () => {
    setCursor(3, 42n);
    state.mailboxOverride = { uidNext: 6, uidValidity: 43n };

    const result = await pollImapAccount("email-account-uidvalidity");

    expect(result.newMessages).toBe(0);
    expect(state.processed).toEqual([]);
    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.credential?.uidValidity).toBe(43n);
  });

  it("records UIDVALIDITY on the first poll", async () => {
    setCursor(null);
    state.mailboxOverride = { uidNext: 6, uidValidity: 42n };

    await pollImapAccount("email-account-first-uidvalidity");

    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.credential?.uidValidity).toBe(42n);
  });

  it("records UIDVALIDITY for an existing cursor that has none", async () => {
    setCursor(5);
    state.mailboxOverride = { uidNext: 6, uidValidity: 42n };

    await pollImapAccount("email-account-learn-uidvalidity");

    expect(state.credential?.lastSeenUid).toBe(5);
    expect(state.credential?.uidValidity).toBe(42n);
  });

  it("never moves the stored UID backward", async () => {
    setCursor(10);

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(state.credential?.lastSeenUid).toBe(10);
  });
});

describe("pollAllImapAccounts", () => {
  it("stops starting accounts once the run budget is spent", async () => {
    state.accounts = [{ id: "email-account-1" }, { id: "email-account-2" }];
    setCursor(5);
    state.mailboxUids = [1, 2, 3, 4, 5, 6];

    vi.useFakeTimers();
    try {
      state.processHook = async () => {
        vi.setSystemTime(Date.now() + POLL_TIME_BUDGET_MS + 1000);
      };

      const results = await pollAllImapAccounts();

      expect(results.map((result) => result.emailAccountId)).toEqual([
        "email-account-1",
      ]);
      expect(state.processed).toEqual([encodeImapMessageId("INBOX", 6)]);
    } finally {
      vi.useRealTimers();
    }
  });
});
