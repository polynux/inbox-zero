import type { ImapFlow } from "imapflow";
import { randomUUID } from "node:crypto";
import { DraftNotFoundError, SafeError } from "@/utils/error";
import { fetchMessageByUid, searchImapMessages } from "@/utils/imap/message";
import { findDraftsFolder } from "@/utils/imap/folder";
import { parseImapMessageId } from "@/utils/imap/message-id";
import type { ParsedMessage } from "@/utils/types";

export type SaveDraftOptions = {
  from: string;
  to: string;
  cc?: string;
  bcc?: string;
  subject: string;
  html?: string;
  text?: string;
  inReplyTo?: string;
  references?: string;
  messageId?: string;
};

export type ResolvedDraft = {
  folder: string;
  uid: number;
  message: ParsedMessage;
};

export async function saveDraft(
  client: ImapFlow,
  options: SaveDraftOptions,
): Promise<string> {
  const draftsFolder = await requireDraftsFolder(client);

  // Secure, unique Message-ID: the stable draft id derives from it, so an
  // update (APPEND + expunge) keeps the same id while the UID changes. It
  // also lets a server without UIDPLUS find the draft this APPEND created.
  const messageId = options.messageId ?? `<${randomUUID()}@inbox-zero.local>`;

  const { id } = await appendDraftMessage(client, draftsFolder, {
    ...options,
    messageId,
  });
  return id;
}

// The stable draft id names the draft's own Message-ID, not its UID: IMAP
// can't edit a message, so an update APPENDs a new copy (new UID) and
// expunges the old one. Folder-qualified ids would break on the second save.
export function encodeImapDraftId(messageId: string): string {
  const normalized = normalizeMessageId(messageId);
  if (!normalized) throw new SafeError("Draft Message-ID is missing");
  return `d.${Buffer.from(normalized, "utf8").toString("base64url")}`;
}

export function parseImapDraftId(draftId: string): string {
  if (!draftId.startsWith("d.") || draftId.length <= 2)
    throw new SafeError(
      "Draft id is not a stable IMAP draft id (expected d.<message-id>)",
    );
  let messageId: string;
  try {
    messageId = Buffer.from(draftId.slice(2), "base64url").toString("utf8");
  } catch {
    throw new SafeError("Draft id holds an undecodable Message-ID");
  }
  const normalized = normalizeMessageId(messageId);
  if (!normalized) throw new SafeError("Draft id holds an empty Message-ID");
  return normalized;
}

export function isImapDraftId(draftId: string): boolean {
  return draftId.startsWith("d.");
}

// APPEND the raw draft and return its stable id. Callers that replace a copy
// (update, send) expunge the old UID themselves; append-first keeps a failure
// from losing the draft.
export async function appendDraftMessage(
  client: ImapFlow,
  draftsFolder: string,
  options: SaveDraftOptions & { messageId: string },
): Promise<{ id: string; uid: number }> {
  const normalizedMessageId = normalizeMessageId(options.messageId);
  if (!normalizedMessageId) throw new SafeError("Draft Message-ID is missing");
  const messageId = `<${normalizedMessageId}>`;

  const messageParts: string[] = [
    `From: ${stripHeaderBreaks(options.from)}`,
    `To: ${stripHeaderBreaks(options.to)}`,
    ...(options.cc ? [`Cc: ${stripHeaderBreaks(options.cc)}`] : []),
    ...(options.bcc ? [`Bcc: ${stripHeaderBreaks(options.bcc)}`] : []),
    `Subject: ${stripHeaderBreaks(options.subject)}`,
    `Date: ${new Date().toUTCString()}`,
    `Message-ID: ${messageId}`,
    "MIME-Version: 1.0",
  ];

  if (options.inReplyTo) {
    messageParts.push(`In-Reply-To: ${stripHeaderBreaks(options.inReplyTo)}`);
  }
  if (options.references) {
    messageParts.push(`References: ${stripHeaderBreaks(options.references)}`);
  }

  if (options.html) {
    messageParts.push("Content-Type: text/html; charset=utf-8");
    messageParts.push("");
    messageParts.push(options.html);
  } else {
    messageParts.push("Content-Type: text/plain; charset=utf-8");
    messageParts.push("");
    messageParts.push(options.text || "");
  }

  const rawMessage = messageParts.join("\r\n");

  const result = await client.append(
    draftsFolder,
    Buffer.from(rawMessage),
    ["\\Draft", "\\Seen"],
    new Date(),
  );

  if (!result) throw new Error("Failed to save draft: APPEND rejected");
  // APPEND without UIDPLUS reports no UID. The draft exists, so recover it by
  // searching the Message-ID written above; failing that, throw instead of
  // returning a fabricated id (a retry would duplicate the draft).
  let uid = result.uid;
  if (!uid) {
    try {
      await client.mailboxOpen(draftsFolder, { readOnly: true });
      [uid] = await searchImapMessages(
        client,
        { header: { "Message-ID": normalizedMessageId } },
        1,
      );
    } catch {
      // Fall through to the explicit error below.
    }
  }
  if (!uid)
    throw new Error(
      "Failed to save draft: server did not return a UID (no UIDPLUS support)",
    );
  return { id: encodeImapDraftId(normalizedMessageId), uid };
}

// Find the draft carrying this Message-ID in the open Drafts folder. Bridge's
// HEADER search is a substring match (`<xr@x>` also matches `<r@x>`), so the
// fetched Message-ID is verified exactly.
export async function findDraftByMessageId(
  client: ImapFlow,
  messageId: string,
): Promise<{ uid: number; message: ParsedMessage } | null> {
  const wanted = normalizeMessageId(messageId);
  if (!wanted) return null;
  const candidates = await searchImapMessages(client, {
    header: { "Message-ID": wanted },
  });
  for (const uid of candidates) {
    const message = await fetchMessageByUid(client, uid);
    if (
      message &&
      normalizeMessageId(message.headers["message-id"] ?? "") === wanted
    )
      return { uid, message };
  }
  return null;
}

// Resolve any draft id form to its current copy in the Drafts folder. Stable
// ids (`d.` + Message-ID) survive updates; the legacy folder-qualified form
// names one UID and goes stale after an update. Malformed ids throw
// SafeError; a well-formed id with no live copy resolves null.
export async function resolveDraftUid(
  client: ImapFlow,
  draftId: string,
  options?: { readOnly?: boolean },
): Promise<ResolvedDraft | null> {
  if (isImapDraftId(draftId)) {
    const wanted = parseImapDraftId(draftId);
    const draftsFolder = await findDraftsFolder(client);
    if (!draftsFolder) return null;
    await client.mailboxOpen(draftsFolder, {
      readOnly: options?.readOnly,
    });
    const found = await findDraftByMessageId(client, wanted);
    return found ? { folder: draftsFolder, ...found } : null;
  }
  const { folder, uid } = parseImapMessageId(draftId);
  const draftsFolder = await findDraftsFolder(client);
  if (!draftsFolder || folder !== draftsFolder) return null;
  await client.mailboxOpen(draftsFolder, { readOnly: options?.readOnly });
  const message = await fetchMessageByUid(client, uid);
  return message ? { folder: draftsFolder, uid, message } : null;
}

export function draftNotFoundError(operation: string): DraftNotFoundError {
  return new DraftNotFoundError(`Could not find this draft to ${operation}.`);
}

// Rebuilding a draft from headers + HTML body drops MIME parts, so refuse
// instead of silently losing attachments.
export function assertDraftHasNoAttachments(
  draft: Pick<ParsedMessage, "attachments" | "inline" | "hasAttachment">,
  operation: string,
): void {
  if (
    (draft.attachments?.length ?? 0) > 0 ||
    (draft.inline?.length ?? 0) > 0 ||
    draft.hasAttachment
  )
    throw new SafeError(
      `Cannot ${operation} this draft: it has attachments, which IMAP drafts do not support yet.`,
    );
}

export async function requireDraftsFolder(client: ImapFlow): Promise<string> {
  const draftsFolder = await findDraftsFolder(client);
  if (draftsFolder) return draftsFolder;
  try {
    await client.mailboxCreate("Drafts");
  } catch {
    throw new Error(
      "No Drafts folder found on the IMAP server and creating one failed",
    );
  }
  const created = await findDraftsFolder(client);
  if (!created) {
    throw new Error(
      "No Drafts folder found on the IMAP server and creating one failed",
    );
  }
  return created;
}

// A raw CR, LF or NUL in a header value would start a new header line.
function stripHeaderBreaks(value: string) {
  return value.replace(/[\r\n\0]/g, "");
}

function normalizeMessageId(value: string): string {
  return value.replace(/^<|>$/g, "").trim().toLowerCase();
}
