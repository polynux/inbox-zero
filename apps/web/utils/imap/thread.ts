import { createHash } from "node:crypto";
import { parseImapMessageId } from "@/utils/imap/message-id";

// Prefix for thread ids of header-less messages. The id carries the
// folder-qualified listing key (`uid@base64url(folder)`, the same key as a
// message id) so the owning message can be derived back from the thread id
// without scanning every mailbox.
const FALLBACK_THREAD_PREFIX = "fallback:";

/**
 * Build a deterministic thread ID from email threading headers.
 * Uses the root Message-ID from the References chain (first entry),
 * falling back to In-Reply-To, then the message's own Message-ID.
 */
export function buildThreadId(
  references: string | undefined,
  inReplyTo: string | undefined,
  messageId: string | undefined,
  fallbackKey?: string,
): string {
  if (!references && !inReplyTo && !messageId) {
    // Header-less message: keep the per-message key in the thread id so the
    // folder + UID stay derivable without a mailbox scan. Without a key
    // there is nothing unique to hash, so keep the legacy constant.
    if (fallbackKey) return `${FALLBACK_THREAD_PREFIX}${fallbackKey}`;
    return hashToThreadId("fallback");
  }
  const rootId = getRootMessageId(references, inReplyTo, messageId);
  return hashToThreadId(rootId);
}

/**
 * Extract the root message ID from threading headers.
 * The References header lists Message-IDs oldest-first,
 * so the first entry is the original message that started the thread.
 */
export function getRootMessageId(
  references: string | undefined,
  inReplyTo: string | undefined,
  messageId: string | undefined,
): string {
  if (references) {
    const ids = parseMessageIdList(references);
    if (ids.length > 0) return ids[0];
  }

  if (inReplyTo) {
    const ids = parseMessageIdList(inReplyTo);
    if (ids.length > 0) return ids[0];
  }

  if (messageId) {
    const ids = parseMessageIdList(messageId);
    if (ids.length > 0) return ids[0];
  }

  return "fallback";
}

/**
 * Parse a space/comma-separated list of Message-IDs (RFC 5322 format).
 * Handles both `<id@domain>` and bare `id@domain` formats.
 */
export function parseMessageIdList(header: string): string[] {
  const ids: string[] = [];
  const regex = /<([^>]+)>/g;
  let match: RegExpExecArray | null;

  match = regex.exec(header);
  while (match !== null) {
    ids.push(match[1]);
    match = regex.exec(header);
  }

  // If no angle-bracket IDs found, try splitting by whitespace
  if (ids.length === 0) {
    return header
      .split(/[\s,]+/)
      .map((s) => s.trim())
      .filter(Boolean);
  }

  return ids;
}

/**
 * Hash a message ID to a fixed-length thread ID.
 */
function hashToThreadId(messageId: string): string {
  return createHash("sha256").update(messageId).digest("hex").slice(0, 24);
}

/**
 * Extract all Message-IDs in a thread's reference chain,
 * useful for searching related messages.
 */
export function getAllThreadMessageIds(
  references: string | undefined,
  inReplyTo: string | undefined,
  messageId: string | undefined,
): string[] {
  const ids = new Set<string>();

  if (references) {
    for (const id of parseMessageIdList(references)) {
      ids.add(id);
    }
  }

  if (inReplyTo) {
    for (const id of parseMessageIdList(inReplyTo)) {
      ids.add(id);
    }
  }

  if (messageId) {
    ids.add(messageId.replace(/^<|>$/g, ""));
  }

  return [...ids];
}

/**
 * Recover the folder-qualified listing key carried by a header-less thread
 * id. Returns null for header-derived (hashed) thread ids. Rejects malformed
 * keys so callers never fetch or act on a fabricated folder.
 */
export function parseFallbackThreadId(threadId: string): {
  folder: string;
  uid: number;
} | null {
  if (!threadId.startsWith(FALLBACK_THREAD_PREFIX)) return null;
  return parseImapMessageId(threadId.slice(FALLBACK_THREAD_PREFIX.length));
}
