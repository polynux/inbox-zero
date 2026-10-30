import { describe, expect, it } from "vitest";
import {
  encodeImapMessageId,
  parseImapMessageId,
} from "@/utils/imap/message-id";

describe("encodeImapMessageId/parseImapMessageId", () => {
  it("round-trips a simple folder", () => {
    const id = encodeImapMessageId("INBOX", 42);
    expect(parseImapMessageId(id)).toEqual({ folder: "INBOX", uid: 42 });
  });

  it.each([
    "All Mail",
    "INBOX/Archive:2024",
    "Gesendet & Empfangen",
    "日本語フォルダ",
    "[Gmail]/All Mail",
  ])("round-trips tricky folder %s", (folder) => {
    const id = encodeImapMessageId(folder, 7);
    expect(parseImapMessageId(id)).toEqual({ folder, uid: 7 });
    // URL-safe: no +/= and a single @ separator.
    expect(id).not.toMatch(/[+/=]/);
    expect(id.split("@")).toHaveLength(2);
  });

  it.each([
    "",
    "1",
    "0@e30=",
    "5@!!!",
    "5@",
    "@e30=",
    "5@e30=@e30=",
    "0@e30",
  ])("rejects garbage id %s", (id) => {
    expect(() => parseImapMessageId(id)).toThrow();
  });
});
