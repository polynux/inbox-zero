import { describe, expect, it } from "vitest";
import { encodeImapMessageId } from "@/utils/imap/message-id";
import {
  buildThreadId,
  getAllThreadMessageIds,
  getRootMessageId,
  parseFallbackThreadId,
  parseMessageIdList,
  parseThreadRootMessageId,
} from "@/utils/imap/thread";

describe("parseMessageIdList", () => {
  it("parses angle-bracket message IDs", () => {
    const result = parseMessageIdList(
      "<abc@example.com> <def@example.com> <ghi@example.com>",
    );
    expect(result).toEqual([
      "abc@example.com",
      "def@example.com",
      "ghi@example.com",
    ]);
  });

  it("handles single message ID", () => {
    const result = parseMessageIdList("<single@example.com>");
    expect(result).toEqual(["single@example.com"]);
  });

  it("falls back to whitespace splitting for bare IDs", () => {
    const result = parseMessageIdList("abc@example.com def@example.com");
    expect(result).toEqual(["abc@example.com", "def@example.com"]);
  });

  it("returns empty array for empty string", () => {
    const result = parseMessageIdList("");
    expect(result).toEqual([]);
  });
});

describe("getRootMessageId", () => {
  it("returns first reference as root", () => {
    const result = getRootMessageId(
      "<root@ex.com> <reply1@ex.com> <reply2@ex.com>",
      "<reply2@ex.com>",
      "<reply3@ex.com>",
    );
    expect(result).toBe("root@ex.com");
  });

  it("falls back to inReplyTo when no references", () => {
    const result = getRootMessageId(
      undefined,
      "<parent@ex.com>",
      "<child@ex.com>",
    );
    expect(result).toBe("parent@ex.com");
  });

  it("strips angle brackets from the messageId fallback", () => {
    const result = getRootMessageId(undefined, undefined, " <only@ex.com> ");
    expect(result).toBe("only@ex.com");
  });

  it("returns 'fallback' when all headers are missing", () => {
    const result = getRootMessageId(undefined, undefined, undefined);
    expect(result).toBe("fallback");
  });
});
describe("buildThreadId with a fallback key", () => {
  const inboxOne = encodeImapMessageId("INBOX", 1);
  const inboxTwo = encodeImapMessageId("INBOX", 2);

  it("gives header-less messages distinct thread ids per key", () => {
    const first = buildThreadId(undefined, undefined, undefined, inboxOne);
    const second = buildThreadId(undefined, undefined, undefined, inboxTwo);
    expect(first).not.toBe(second);
  });

  it("is stable for the same fallback key", () => {
    expect(buildThreadId(undefined, undefined, undefined, inboxOne)).toBe(
      buildThreadId(undefined, undefined, undefined, inboxOne),
    );
  });

  it("keeps the listing key verbatim in the thread id", () => {
    expect(buildThreadId(undefined, undefined, undefined, inboxOne)).toBe(
      `fallback:${inboxOne}`,
    );
  });

  it("requires a fallback key for header-less messages", () => {
    expect(() => buildThreadId(undefined, undefined, undefined)).toThrow(
      /fallback key/,
    );
  });
});

describe("parseFallbackThreadId", () => {
  it("recovers folder and uid from a header-less thread id", () => {
    expect(
      parseFallbackThreadId(`fallback:${encodeImapMessageId("INBOX", 9)}`),
    ).toEqual({ folder: "INBOX", uid: 9 });
    expect(
      parseFallbackThreadId(`fallback:${encodeImapMessageId("All Mail", 12)}`),
    ).toEqual({ folder: "All Mail", uid: 12 });
  });

  it("round-trips a folder that contains a slash", () => {
    const threadId = buildThreadId(
      undefined,
      undefined,
      undefined,
      encodeImapMessageId("Projects/Acme", 4),
    );

    expect(parseFallbackThreadId(threadId)).toEqual({
      folder: "Projects/Acme",
      uid: 4,
    });
  });

  it("returns null for hashed (header-derived) thread ids", () => {
    expect(
      parseFallbackThreadId(buildThreadId(undefined, undefined, "<a@x>")),
    ).toBeNull();
  });

  it.each([
    "fallback:",
    "fallback:INBOX",
    "fallback::9",
    "fallback:INBOX:0",
    "fallback:INBOX:x",
    "fallback:0@SU5CT1g",
    "fallback:9@",
    "fallback:9@YQ@Yg",
  ])("rejects malformed key %s", (threadId) => {
    expect(() => parseFallbackThreadId(threadId)).toThrow();
  });
});

describe("buildThreadId", () => {
  it("encodes the root Message-ID recoverably with a t. prefix", () => {
    const id = buildThreadId(
      "<root@ex.com> <r1@ex.com>",
      "<r1@ex.com>",
      "<r2@ex.com>",
    );
    expect(id).toBe(
      `t.${Buffer.from("root@ex.com", "utf8").toString("base64url")}`,
    );
    expect(parseThreadRootMessageId(id)).toBe("root@ex.com");
  });

  it("hashes a root message and its reply to the same thread", () => {
    const root = buildThreadId(undefined, undefined, "<abc@x>");
    const reply = buildThreadId("<abc@x>", "<abc@x>", "<def@x>");
    expect(root).toBe(reply);
  });

  it("returns different ids for different threads", () => {
    const id1 = buildThreadId("<thread1@ex.com>", undefined, undefined);
    const id2 = buildThreadId("<thread2@ex.com>", undefined, undefined);
    expect(id1).not.toBe(id2);
  });

  it("normalizes angle brackets and case before encoding", () => {
    expect(buildThreadId(undefined, undefined, " <ABC@X> ")).toBe(
      buildThreadId(undefined, undefined, "<abc@x>"),
    );
  });

  it("parseThreadRootMessageId returns null for fallback ids", () => {
    expect(
      parseThreadRootMessageId(
        buildThreadId(undefined, undefined, undefined, "9@SU5CT1g"),
      ),
    ).toBeNull();
  });

  it("parseThreadRootMessageId rejects malformed ids", () => {
    expect(() => parseThreadRootMessageId("t.!!!")).toThrow();
    expect(parseThreadRootMessageId("nope")).toBeNull();
  });

  it("parseThreadRootMessageId rejects roots that cannot be a Message-ID", () => {
    const shortRoot = `t.${Buffer.from("a", "utf8").toString("base64url")}`;
    expect(() => parseThreadRootMessageId(shortRoot)).toThrow(/implausible/);
    const noAt = `t.${Buffer.from("no-at-sign-here", "utf8").toString("base64url")}`;
    expect(() => parseThreadRootMessageId(noAt)).toThrow(/implausible/);
  });
});

describe("getAllThreadMessageIds", () => {
  it("collects all unique message IDs from headers", () => {
    const ids = getAllThreadMessageIds(
      "<a@ex.com> <b@ex.com>",
      "<b@ex.com>",
      "<c@ex.com>",
    );
    expect(ids).toContain("a@ex.com");
    expect(ids).toContain("b@ex.com");
    expect(ids).toContain("c@ex.com");
    expect(ids).toHaveLength(3);
  });

  it("deduplicates IDs across headers", () => {
    const ids = getAllThreadMessageIds(
      "<same@ex.com>",
      "<same@ex.com>",
      "<same@ex.com>",
    );
    expect(ids).toHaveLength(1);
  });

  it("returns empty array when all headers are missing", () => {
    const ids = getAllThreadMessageIds(undefined, undefined, undefined);
    expect(ids).toHaveLength(0);
  });
});
