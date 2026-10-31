import { SafeError } from "@/utils/error";
import { parseImapMessageId } from "@/utils/imap/message-id";

// Thread ids for messages with threading headers. The id is `t.` plus the
// base64url-encoded normalized root Message-ID, so the root stays
// recoverable and `fetchThreadMessages` can SEARCH for it directly instead
// of scanning a newest-N window.
const THREAD_ID_PREFIX = "t.";

// Prefix for thread ids of header-less messages. The id carries the
// folder-qualified listing key (`uid@base64url(folder)`, the same key as a
// message id) so the owning message can be derived back from the thread id
// without scanning every mailbox.
const FALLBACK_THREAD_PREFIX = "fallback:";

/**
 * Build a deterministic thread ID from email threading headers.
 * Uses the root Message-ID from the References chain (first entry),
 * falling back to In-Reply-To, then the message's own Message-ID.
 * The id encodes the normalized root, so callers can recover it with
 * `parseThreadRootMessageId` and SEARCH for members directly.
 */
export function buildThreadId(
  references: string | undefined,
  inReplyTo: string | undefined,
  messageId: string | undefined,
  fallbackKey?: string,
): string {
  if (!references && !inReplyTo && !messageId) {
    // Header-less message: keep the per-message key in the thread id so the
    // folder + UID stay derivable without a mailbox scan. The key is
    // required; callers always have the listing key.
    if (!fallbackKey) throw new SafeError("Thread id needs a fallback key");
    return `${FALLBACK_THREAD_PREFIX}${fallbackKey}`;
  }
  // Normalize (strip angle brackets/whitespace, lowercase): servers vary the
  // header's case and spacing, but the thread must stay one thread.
  const rootId = getRootMessageId(references, inReplyTo, messageId)
    .replace(/^<|>$/g, "")
    .trim()
    .toLowerCase();
  return `${THREAD_ID_PREFIX}${Buffer.from(rootId, "utf8").toString("base64url")}`;
}

/**
 * Recover the normalized root Message-ID carried by a header-derived
 * thread id. Returns null for header-less (`fallback:`) ids. Rejects
 * malformed ids and implausible roots so callers never SEARCH for a
 * fabricated or unbounded-matching value.
 */
export function parseThreadRootMessageId(threadId: string): string | null {
  if (threadId.startsWith(FALLBACK_THREAD_PREFIX)) return null;
  if (!threadId.startsWith(THREAD_ID_PREFIX)) return null;
  const root = Buffer.from(
    threadId.slice(THREAD_ID_PREFIX.length),
    "base64url",
  ).toString("utf8");
  if (!root) throw new SafeError("Thread id holds an empty root Message-ID");
  // The decoded root becomes a SEARCH term, and gluon matches substrings:
  // a one-character root fans out over the whole mailbox. `Buffer.from(x,
  // "base64url")` never throws on bad input, so decodability alone proves
  // nothing — require something shaped like a Message-ID.
  if (root.length < 3 || !root.includes("@") || /\s/.test(root))
    throw new SafeError("Thread id holds an implausible root Message-ID");
  return root;
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
