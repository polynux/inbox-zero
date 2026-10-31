import type { ImapFlow } from "imapflow";
import MailComposer from "nodemailer/lib/mail-composer";
import type { Attachment as MailAttachment } from "nodemailer/lib/mailer";
import { randomUUID } from "node:crypto";
import { simpleParser } from "mailparser";
import { DraftNotFoundError, SafeError } from "@/utils/error";
import { fetchMessageByUid, searchImapMessages } from "@/utils/imap/message";
import { findDraftsFolder } from "@/utils/imap/folder";
import { parseImapMessageId } from "@/utils/imap/message-id";
import type { ParsedMessage } from "@/utils/types";

// Stable per-file id across draft rebuilds, same header Gmail stamps so the
// composer can remove a file it attached earlier.
export const DRAFT_ATTACHMENT_ID_HEADER = "X-Attachment-Id";

export type DraftAttachmentPart = {
  id: string;
  filename: string;
  mimeType: string;
  size: number;
  disposition: "attachment" | "inline";
  contentId?: string;
  content: Buffer;
};

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
  attachments?: MailAttachment[];
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

  // Attachment drafts go through the MIME builder so files and inline images
  // keep their parts; plain drafts keep the hand-built single-part body.
  const raw = await buildDraftMime({ ...options, messageId });

  const result = await client.append(
    draftsFolder,
    raw,
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

export async function buildDraftMime(
  options: SaveDraftOptions & { messageId: string },
): Promise<Buffer> {
  if (!options.attachments?.length)
    return Buffer.from(buildPlainDraftMessage(options));
  const messageIdHeader = stripHeaderBreaks(options.messageId);
  const compiled = new MailComposer({
    from: stripHeaderBreaks(options.from),
    to: stripHeaderBreaks(options.to),
    cc: options.cc ? stripHeaderBreaks(options.cc) : undefined,
    bcc: options.bcc ? stripHeaderBreaks(options.bcc) : undefined,
    subject: stripHeaderBreaks(options.subject),
    messageId: messageIdHeader,
    date: new Date().toUTCString(),
    inReplyTo: options.inReplyTo
      ? stripHeaderBreaks(options.inReplyTo)
      : undefined,
    references: options.references
      ? stripHeaderBreaks(options.references)
      : undefined,
    ...(options.html ? { html: options.html } : { text: options.text || "" }),
    attachments: options.attachments,
  }).compile();
  compiled.keepBcc = true;
  return compiled.build();
}

function buildPlainDraftMessage(
  options: SaveDraftOptions & { messageId: string },
): string {
  const messageParts: string[] = [
    `From: ${stripHeaderBreaks(options.from)}`,
    `To: ${stripHeaderBreaks(options.to)}`,
    ...(options.cc ? [`Cc: ${stripHeaderBreaks(options.cc)}`] : []),
    ...(options.bcc ? [`Bcc: ${stripHeaderBreaks(options.bcc)}`] : []),
    `Subject: ${stripHeaderBreaks(options.subject)}`,
    `Date: ${new Date().toUTCString()}`,
    `Message-ID: ${options.messageId}`,
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

  return messageParts.join("\r\n");
}

// Read a draft's files back from its MIME so a rebuild can carry them: the
// raw source is the only store, since APPEND keeps no attachment metadata.
export async function readDraftAttachmentParts(
  source: Buffer,
): Promise<DraftAttachmentPart[]> {
  const parsed = await simpleParser(source);
  const seen = new Map<string, number>();
  return parsed.attachments.map((part) => {
    const stamped = part.headers
      .get(DRAFT_ATTACHMENT_ID_HEADER.toLowerCase())
      ?.toString();
    const filename = part.filename || "attachment";
    const mimeType = part.contentType || "application/octet-stream";
    return {
      id:
        stamped && /^[\w-]{1,64}$/u.test(stamped)
          ? stamped
          : derivedDraftAttachmentId(seen, filename, mimeType, part.size),
      filename,
      mimeType,
      size: part.size,
      disposition:
        part.contentDisposition === "inline" ? "inline" : "attachment",
      ...(part.contentId
        ? { contentId: part.contentId.replace(/^<|>$/g, "") }
        : {}),
      content: part.content,
    };
  });
}

export function toDraftMailAttachment(
  part: Pick<
    DraftAttachmentPart,
    "id" | "filename" | "mimeType" | "disposition" | "contentId"
  > & { content: Buffer },
): MailAttachment {
  const inline = part.disposition === "inline" && part.contentId;
  return {
    filename: part.filename,
    contentType: part.mimeType,
    content: part.content,
    // An attached email must keep a 7bit or 8bit encoding (RFC 2046), so the
    // builder picks it; every other file is base64.
    ...(part.mimeType.startsWith("message/")
      ? {}
      : { contentTransferEncoding: "base64" }),
    contentDisposition: inline ? "inline" : "attachment",
    ...(inline ? { cid: part.contentId } : {}),
    headers: { [DRAFT_ATTACHMENT_ID_HEADER]: part.id },
  };
}

// Files the composer didn't stamp (forwards, other clients) still need a
// stable id within the draft; filename + type + size disambiguates repeats.
function derivedDraftAttachmentId(
  seen: Map<string, number>,
  filename: string,
  mimeType: string,
  size: number,
): string {
  const base = `${filename}:${mimeType}:${size}`;
  const count = seen.get(base) ?? 0;
  seen.set(base, count + 1);
  return count === 0 ? base : `${base}:${count}`;
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
