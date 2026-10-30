import { beforeEach, describe, expect, it, vi } from "vitest";
import { getImapCredentials } from "@/utils/imap/credential";
import { searchImapMessages } from "@/utils/imap/message";
import { pollImapAccount } from "@/utils/imap/poll";
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
  searchImapMessages: vi.fn(),
}));

let mailboxUids: number[];

const stubClient = {
  mailboxOpen: vi.fn(async () => ({
    uidNext: mailboxUids.length > 0 ? Math.max(...mailboxUids) + 1 : 1,
  })),
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
  vi.mocked(getImapCredentials).mockResolvedValue({} as never);
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

  it("never moves the stored UID backward", async () => {
    mockLastSeenUid(10);

    const result = await pollImapAccount("email-account-1");

    expect(result.newMessages).toBe(0);
    expect(prisma.imapCredential.update).not.toHaveBeenCalled();
  });
});
