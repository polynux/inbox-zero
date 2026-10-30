import { describe, expect, it } from "vitest";
import { parseSearchQuery } from "@/utils/imap/message";

describe("parseSearchQuery", () => {
  it("parses from: query", () => {
    const result = parseSearchQuery("from:test@example.com");
    expect(result).toEqual({ from: "test@example.com" });
  });

  it("parses to: query", () => {
    const result = parseSearchQuery("to:user@example.com");
    expect(result).toEqual({ to: "user@example.com" });
  });

  it("parses subject: query", () => {
    const result = parseSearchQuery("subject:meeting");
    expect(result).toEqual({ subject: "meeting" });
  });

  it("parses is:unread", () => {
    const result = parseSearchQuery("is:unread");
    expect(result).toEqual({ unseen: true });
  });

  it("parses is:read", () => {
    const result = parseSearchQuery("is:read");
    expect(result).toEqual({ seen: true });
  });

  it("parses combined queries", () => {
    const result = parseSearchQuery("from:test@ex.com is:unread");
    expect(result).toEqual({
      and: [{ from: "test@ex.com" }, { unseen: true }],
    });
  });

  it("returns all:true for empty query", () => {
    const result = parseSearchQuery("");
    expect(result).toEqual({ all: true });
  });

  it("treats bare text as body search", () => {
    const result = parseSearchQuery("important");
    expect(result).toEqual({ body: "important" });
  });

  it("keeps a quoted field value as one term", () => {
    const result = parseSearchQuery('subject:"weekly update"');
    expect(result).toEqual({ subject: "weekly update" });
  });

  it("keeps a bare quoted phrase as one body term", () => {
    const result = parseSearchQuery('"exact phrase"');
    expect(result).toEqual({ body: "exact phrase" });
  });

  it("mixes quoted values with unquoted terms", () => {
    const result = parseSearchQuery(
      'from:alice@example.com subject:"weekly update" important',
    );
    expect(result).toEqual({
      and: [
        { from: "alice@example.com" },
        { subject: "weekly update" },
        { body: "important" },
      ],
    });
  });

  it("treats an unterminated quote as a single term", () => {
    expect(parseSearchQuery('subject:"weekly update')).toEqual({
      subject: "weekly update",
    });
    expect(parseSearchQuery('"exact phrase')).toEqual({
      body: "exact phrase",
    });
  });
});
