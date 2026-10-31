import { beforeEach, describe, expect, it, vi } from "vitest";
import { getImapCredentials } from "@/utils/imap/credential";
import { encodeImapMessageId } from "@/utils/imap/message-id";
import { fetchMessageByUid, searchImapMessages } from "@/utils/imap/message";
import { pollImapAccount } from "@/utils/imap/poll";
import { createEmailProvider } from "@/utils/email/provider";
import { processHistoryItem } from "@/utils/webhook/process-history-item";
import { validateWebhookAccount } from "@/utils/webhook/validate-webhook-account";
import prisma from "@/utils/__mocks__/prisma";

vi.mock("@/utils/prisma");
vi.mock("@/utils/imap/credential", () => ({
  getImapCredentials: vi.fn(),
}));
vi.mock("@/utils/imap/client", () => ({
  withImapConnection: vi.fn(
    async (_config: unknown, fn: (client: unknown) => Promise<unknown>) =>
      fn(stubClient),
  ),
}));
vi.mock("@/utils/imap/message", () => ({
  fetchMessageByUid: vi.fn(),
  searchImapMessages: vi.fn(),
}));
vi.mock("@/utils/email/provider", () => ({
  createEmailProvider: vi.fn(),
}));
vi.mock("@/utils/redis/message-processing", () => ({
  withMessageProcessingLock: vi.fn(
    async (_key: unknown, fn: () => Promise<unknown>) => {
      await fn();
      return { skipped: false };
    },
  ),
}));
vi.mock("@/utils/webhook/process-history-item", () => ({
  processHistoryItem: vi.fn(),
}));
vi.mock("@/utils/webhook/validate-webhook-account", () => ({
  getWebhookEmailAccount: vi.fn(),
  validateWebhookAccount: vi.fn(),
}));

let mailboxUids: number[];

const stubClient = {
  mailboxOpen: vi.fn(async () => ({
    uidNext: mailboxUids.length > 0 ? Math.max(...mailboxUids) + 1 : 1,
  })),
};

const validatedAccount = {
  id: "email-account-1",
  email: "me@example.com",
  userId: "user-1",
  account: { provider: "imap" },
  rules: [],
};

function mockLastSeenUid(lastSeenUid: number | null) {
  prisma.imapCredential.findFirst.mockResolvedValue({
    id: "credential-1",
    lastSeenUid,
  } as never);
}

beforeEach(() => {
  vi.clearAllMocks();
  mailboxUids = [1, 2, 3, 4, 5];
  vi.mocked(getImapCredentials).mockResolvedValue({
    email: "me@example.com",
  } as never);
  prisma.imapCredential.update.mockResolvedValue({
    id: "credential-1",
  } as never);
  vi.mocked(searchImapMessages).mockImplementation(
    async (_client, criteria) => {
      if (typeof criteria.uid !== "string") return [...mailboxUids];
      const from = Number(criteria.uid.split(":")[0]);
      const matched = mailboxUids.filter((uid) => uid >= from);
      // RFC 3501: `n:*` matches the highest UID even when n is above it.
      return (matched.length > 0 ? matched : mailboxUids.slice(-1)).reverse();
    },
  );
  vi.mocked(fetchMessageByUid).mockImplementation(
    async (_client, uid) =>
      ({ id: encodeImapMessageId("INBOX", uid) }) as never,
  );
  vi.mocked(createEmailProvider).mockResolvedValue({} as never);
  vi.mocked(validateWebhookAccount).mockResolvedValue({
    success: true,
    data: {
      emailAccount: validatedAccount,
      hasAutomationRules: true,
      hasAiAccess: true,
    },
  } as never);
  vi.mocked(processHistoryItem).mockResolvedValue(undefined);
});

describe("pollImapAccount", () => {
  it("reports nothing for an idle mailbox instead of its newest message", async () => {
    mockLastSeenUid(5);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 0,
    });
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();
    expect(processHistoryItem).not.toHaveBeenCalled();
  });

  it("reports and records new messages above the stored UID", async () => {
    mockLastSeenUid(3);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 2,
    });
    expect(prisma.imapCredential.update).toHaveBeenCalledWith(
      expect.objectContaining({
        where: { id: "credential-1" },
        data: expect.objectContaining({ lastSeenUid: 5 }),
      }),
    );
  });

  it("feeds new inbox messages to the shared processor with qualified ids", async () => {
    mockLastSeenUid(4);

    await pollImapAccount("email-account-1");

    const messageId = encodeImapMessageId("INBOX", 5);
    expect(processHistoryItem).toHaveBeenCalledTimes(1);
    expect(processHistoryItem).toHaveBeenCalledWith(
      { messageId, message: { id: messageId } },
      expect.objectContaining({
        provider: expect.anything(),
        hasAutomationRules: true,
        hasAiAccess: true,
        emailAccount: expect.objectContaining({
          id: "email-account-1",
          account: { provider: "imap" },
        }),
      }),
    );
  });

  it("does not advance the cursor past a failed message", async () => {
    mockLastSeenUid(3);
    vi.mocked(processHistoryItem).mockRejectedValueOnce(
      new Error("processing failed"),
    );

    const result = await pollImapAccount("email-account-failure");

    expect(result.newMessages).toBe(0);
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();
  });

  it("skips a message that keeps failing so newer mail is not blocked", async () => {
    vi.mocked(processHistoryItem).mockRejectedValue(
      new Error("processing failed"),
    );
    mailboxUids = [1, 2, 3, 4, 5];
    prisma.imapCredential.findFirst.mockResolvedValue({
      id: "credential-1",
      lastSeenUid: 4,
    } as never);

    await pollImapAccount("email-account-permanent-failure");
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();

    await pollImapAccount("email-account-permanent-failure");
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();

    vi.mocked(processHistoryItem).mockResolvedValue(undefined);
    await pollImapAccount("email-account-permanent-failure");

    expect(prisma.imapCredential.update).toHaveBeenCalledWith(
      expect.objectContaining({
        data: expect.objectContaining({ lastSeenUid: 5 }),
      }),
    );
  });

  it("skips accounts without automation rules or AI access", async () => {
    mockLastSeenUid(3);
    vi.mocked(validateWebhookAccount).mockResolvedValue({
      success: false,
      response: new Response(),
    } as never);

    const result = await pollImapAccount("email-account-no-rules");

    expect(result).toEqual({
      emailAccountId: "email-account-no-rules",
      newMessages: 0,
    });
    expect(processHistoryItem).not.toHaveBeenCalled();
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();
  });

  it("baselines a null cursor: first poll reports 0 new", async () => {
    mockLastSeenUid(null);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 0,
    });
    expect(prisma.imapCredential.update).toHaveBeenCalledWith(
      expect.objectContaining({
        where: { id: "credential-1" },
        data: expect.objectContaining({ lastSeenUid: 5 }),
      }),
    );
  });

  it("keeps the baseline unset when the server omits UIDNEXT", async () => {
    mockLastSeenUid(null);
    vi.mocked(stubClient.mailboxOpen).mockResolvedValueOnce({} as never);

    const result = await pollImapAccount("email-account-1");

    expect(result).toEqual({
      emailAccountId: "email-account-1",
      newMessages: 0,
    });
    expect(prisma.imapCredential.update).toHaveBeenCalledWith({
      where: { id: "credential-1" },
      data: { lastPolledAt: expect.any(Date) },
    });
  });

  it("never moves the stored UID backward", async () => {
    mockLastSeenUid(10);

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();
  });
});
