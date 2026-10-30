import type {
  ImapFlow,
  FetchMessageObject,
  MailboxObject,
  MessageStructureObject,
} from "imapflow";
import { simpleParser } from "mailparser";
import type { ParsedMessage, ParsedMessageHeaders } from "@/utils/types";
import { encodeImapMessageId } from "@/utils/imap/message-id";
import { buildThreadId } from "@/utils/imap/thread";

// Gmail-style system label ids that downstream consumers (sync observations,
// stats, assistant tools) read off ParsedMessage.labelIds.
const IMAP_LABEL = {
  INBOX: "INBOX",
  SENT: "SENT",
  UNREAD: "UNREAD",
  STARRED: "STARRED",
  SPAM: "SPAM",
  TRASH: "TRASH",
  DRAFT: "DRAFT",
} as const;

/**
 * Fetch a single message by sequence number with full body content.
 */
export async function fetchMessageBySeq(
  client: ImapFlow,
  seq: number,
): Promise<ParsedMessage | null> {
  const msg = await client.fetchOne(String(seq), {
    uid: true,
    envelope: true,
    flags: true,
    bodyStructure: true,
  });

  if (!msg) return null;

  const body = await downloadMessageBody(client, seq);
  return convertImapMessage(
    msg,
    body,
    client.mailbox?.path,
    client.mailbox?.specialUse,
  );
}

/**
 * Fetch a single message by UID with full body content.
 * Uses SEARCH to find the sequence number first (WorkMail-compatible).
 */
export async function fetchMessageByUid(
  client: ImapFlow,
  uid: number,
): Promise<ParsedMessage | null> {
  try {
    // Find the sequence number for this UID
    const seqNums =
      (await client.search({ uid: `${uid}` }, { uid: false })) || [];
    if (seqNums.length === 0) return null;

    const seq = seqNums[0];
    const msg = await client.fetchOne(String(seq), {
      uid: true,
      envelope: true,
      flags: true,
      bodyStructure: true,
    });
    if (!msg) return null;

    const body = await downloadMessageBody(client, seq);
    return convertImapMessage(
      msg,
      body,
      client.mailbox?.path,
      client.mailbox?.specialUse,
    );
  } catch {
    return null;
  }
}

/**
 * Fetch the most recent N messages from the currently selected mailbox.
 * Uses sequence numbers (most reliable across IMAP servers).
 */
export async function fetchRecentMessages(
  client: ImapFlow,
  mailbox: MailboxObject,
  maxResults: number,
): Promise<ParsedMessage[]> {
  const total = mailbox.exists || 0;
  if (total === 0) return [];

  const start = Math.max(1, total - maxResults + 1);
  const range = `${start}:*`;

  const messages: ParsedMessage[] = [];
  for await (const msg of client.fetch(range, {
    uid: true,
    envelope: true,
    flags: true,
    bodyStructure: true,
  })) {
    const parsed = await convertImapMessage(
      msg,
      undefined,
      mailbox.path,
      mailbox.specialUse,
    );
    if (parsed) messages.push(parsed);
  }

  // Return newest first
  messages.reverse();
  return messages;
}

/**
 * Fetch multiple messages by UIDs - envelope only (no body).
 * Uses UID-based SEARCH to find each message's sequence number,
 * then fetches by sequence range (WorkMail-compatible).
 */
export async function fetchMessagesByUids(
  client: ImapFlow,
  uids: number[],
  log?: { warn: (message: string, context?: Record<string, unknown>) => void },
): Promise<ParsedMessage[]> {
  if (uids.length === 0) return [];

  // Find the sequence numbers for these UIDs via SEARCH
  // Then fetch by sequence range which works on all servers
  const messages: ParsedMessage[] = [];

  for (const uid of uids) {
    try {
      // SEARCH UID <uid> returns matching sequence numbers
      const seqNums =
        (await client.search({ uid: `${uid}` }, { uid: false })) || [];
      if (seqNums.length === 0) continue;

      const seq = seqNums[0];
      const msg = await client.fetchOne(String(seq), {
        uid: true,
        envelope: true,
        flags: true,
        bodyStructure: true,
      });
      if (msg) {
        const parsed = await convertImapMessage(
          msg,
          undefined,
          client.mailbox?.path,
          client.mailbox?.specialUse,
        );
        if (parsed) messages.push(parsed);
      }
    } catch (error) {
      // Skip messages that fail to fetch, but log at warn (not silently):
      // callers move/flag the returned set, so a silent drop reads as done.
      log?.warn("IMAP fetch by UID failed", {
        folder: client.mailbox?.path,
        uid,
        error,
      });
    }
  }

  return messages;
}

/**
 * Download the full body of a message by sequence number.
 */
async function downloadMessageBody(
  client: ImapFlow,
  seq: number,
): Promise<{ textHtml?: string; textPlain?: string; references?: string }> {
  try {
    const downloaded = await client.download(String(seq));
    const chunks: Buffer[] = [];
    for await (const chunk of downloaded.content) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    const raw = Buffer.concat(chunks);
    const parsed = await simpleParser(raw);

    let references: string | undefined;
    if (parsed.references) {
      references = Array.isArray(parsed.references)
        ? parsed.references.join(" ")
        : parsed.references;
    }

    return {
      textHtml: parsed.html || undefined,
      textPlain: parsed.text || undefined,
      references,
    };
  } catch {
    return {};
  }
}

/**
 * Search messages in the currently selected mailbox.
 */
export async function searchImapMessages(
  client: ImapFlow,
  criteria: Record<string, unknown>,
  maxResults?: number,
): Promise<number[]> {
  const uids = (await client.search(criteria, { uid: true })) || [];

  // UIDs are returned in ascending order; reverse for newest-first
  uids.reverse();

  if (maxResults && uids.length > maxResults) {
    return uids.slice(0, maxResults);
  }

  return uids;
}

/**
 * Convert an imapflow FetchMessageObject to our ParsedMessage format.
 */
export async function convertImapMessage(
  msg: FetchMessageObject,
  body:
    | { textHtml?: string; textPlain?: string; references?: string }
    | undefined,
  folderPath: string,
  specialUse?: string,
): Promise<ParsedMessage | null> {
  try {
    const envelope = msg.envelope;
    if (!envelope) return null;

    const textHtml = body?.textHtml;
    const textPlain = body?.textPlain;
    const references = body?.references;

    const fromAddr = envelope.from?.[0];
    const fromStr = fromAddr
      ? formatAddress(fromAddr.name, fromAddr.address)
      : "";

    const toAddrs = envelope.to || [];
    const toStr = toAddrs
      .map((a) => formatAddress(a.name, a.address))
      .join(", ");

    const ccAddrs = envelope.cc || [];
    const ccStr = ccAddrs
      .map((a) => formatAddress(a.name, a.address))
      .join(", ");

    const bccAddrs = envelope.bcc || [];
    const bccStr = bccAddrs
      .map((a) => formatAddress(a.name, a.address))
      .join(", ");

    const messageId = envelope.messageId || undefined;
    const inReplyTo = envelope.inReplyTo || undefined;

    const date = envelope.date
      ? new Date(envelope.date).toISOString()
      : new Date().toISOString();

    const fallbackKey = msg.uid
      ? encodeImapMessageId(folderPath, msg.uid)
      : undefined;
    const threadId = buildThreadId(
      references,
      inReplyTo,
      messageId,
      fallbackKey,
    );

    const flags = msg.flags ? [...msg.flags] : [];
    const labelIds = buildImapLabelIds({ flags, folderPath, specialUse });

    const headers: ParsedMessageHeaders = {
      from: fromStr,
      to: toStr,
      subject: envelope.subject || "(no subject)",
      date,
      ...(ccStr && { cc: ccStr }),
      ...(bccStr && { bcc: bccStr }),
      ...(messageId && { "message-id": messageId }),
      ...(inReplyTo && { "in-reply-to": inReplyTo }),
      ...(references && { references }),
    };

    const snippet = textPlain
      ? textPlain.slice(0, 200).replace(/\n/g, " ")
      : envelope.subject || "";

    const { attachments, inline } = extractAttachments(msg.bodyStructure);

    return {
      id: msg.uid ? encodeImapMessageId(folderPath, msg.uid) : String(msg.uid),
      threadId,
      historyId: msg.uid ? `${folderPath}:${msg.uid}` : String(msg.uid),
      date,
      headers,
      subject: envelope.subject || "(no subject)",
      snippet,
      textHtml,
      textPlain,
      labelIds,
      attachments: attachments.length > 0 ? attachments : undefined,
      hasAttachment: attachments.length > 0 ? true : undefined,
      inline,
      ...(inReplyTo && { internalDate: date }),
    };
  } catch {
    return null;
  }
}

// Walk an imapflow BODYSTRUCTURE tree and collect attachment metadata
// (filename, MIME type, size, part id). Content bytes stay on the server:
// callers that need them use download/getAttachment on the part id.
function extractAttachments(structure?: MessageStructureObject): {
  attachments: ParsedMessage["attachments"];
  inline: ParsedMessage["inline"];
} {
  const attachments: NonNullable<ParsedMessage["attachments"]> = [];
  const inline: NonNullable<ParsedMessage["inline"]> = [];

  if (!structure) return { attachments: [], inline: [] };

  const visit = (node: MessageStructureObject) => {
    for (const child of node.childNodes || []) visit(child);

    if (node.childNodes?.length) return;

    const disposition = node.disposition?.toLowerCase();
    const filename =
      node.dispositionParameters?.filename || node.parameters?.name;
    const isInline = disposition === "inline";
    const isAttachment =
      disposition === "attachment" || (!!filename && !isInline);
    if (!isAttachment && !isInline) return;

    const entry = {
      attachmentId: node.part || "",
      filename: filename || `part-${node.part || "1"}`,
      mimeType: node.type || "application/octet-stream",
      size: node.size || 0,
      headers: {
        "content-description": node.description || "",
        "content-disposition": node.disposition || "",
        "content-id": node.id || "",
        "content-transfer-encoding": node.encoding || "",
        "content-type": node.type || "",
      },
    };

    if (isInline) inline.push(entry);
    else attachments.push(entry);
  };

  visit(structure);
  return { attachments, inline };
}

function formatAddress(
  name: string | undefined,
  address: string | undefined,
): string {
  if (!address) return name || "";
  if (name) return `${name} <${address}>`;
  return address;
}

/**
 * Build an IMAP SEARCH criteria object from a simple query string.
 */
export function parseSearchQuery(query: string): Record<string, unknown> {
  const criteria: Record<string, unknown>[] = [];

  // Quoted values (`subject:"weekly update"`) and bare phrases keep their
  // spaces; an unterminated quote swallows the rest of the query.
  const parts = query.match(/(\w+:"[^"]*"?|"[^"]*"?|\S+)/g) || [];

  for (const part of parts) {
    const fieldMatch = part.match(/^(\w+):([\s\S]*)$/);
    const field = fieldMatch?.[1];
    const value = fieldMatch ? unquote(fieldMatch[2]) : part;

    if (part === "is:unread") {
      criteria.push({ unseen: true });
    } else if (part === "is:read") {
      criteria.push({ seen: true });
    } else if (part === "has:attachment") {
      criteria.push({ header: { "Content-Type": "multipart/mixed" } });
    } else if (field === "from") {
      criteria.push({ from: value });
    } else if (field === "to") {
      criteria.push({ to: value });
    } else if (field === "subject") {
      criteria.push({ subject: value });
    } else if (field === "since" || field === "after") {
      criteria.push({ since: new Date(value) });
    } else if (field === "before") {
      criteria.push({ before: new Date(value) });
    } else {
      criteria.push({ body: unquote(part) });
    }
  }

  if (criteria.length === 0) return { all: true };
  if (criteria.length === 1) return criteria[0];
  return { and: criteria };
}

function unquote(value: string): string {
  if (!value.startsWith('"')) return value;
  const end =
    value.length > 1 && value.endsWith('"') ? value.length - 1 : value.length;
  return value.slice(1, end);
}

export function buildImapLabelIds(options: {
  flags?: Set<string> | string[];
  folderPath?: string;
  specialUse?: string;
}): string[] {
  const flags = options.flags ? [...options.flags] : [];
  const labelIds: string[] = [];

  if (!flags.includes("\\Seen")) labelIds.push(IMAP_LABEL.UNREAD);
  if (flags.includes("\\Flagged")) labelIds.push(IMAP_LABEL.STARRED);
  if (flags.includes("\\Draft")) labelIds.push(IMAP_LABEL.DRAFT);

  const role = imapRoleLabel(options.folderPath, options.specialUse);
  if (role) labelIds.push(role);

  for (const flag of flags) {
    if (flag.startsWith("\\") || labelIds.includes(flag)) continue;
    labelIds.push(flag);
  }

  return labelIds;
}

function imapRoleLabel(
  folderPath?: string,
  specialUse?: string,
): string | undefined {
  // Special-use flags are authoritative; name matching below is exact-only
  // so user folders like "Presentations" never read as SENT.
  const use = specialUse?.toLowerCase();
  if (use === "\\trash") return IMAP_LABEL.TRASH;
  if (use === "\\junk") return IMAP_LABEL.SPAM;
  if (use === "\\sent") return IMAP_LABEL.SENT;
  if (use === "\\drafts") return IMAP_LABEL.DRAFT;
  if (use === "\\inbox") return IMAP_LABEL.INBOX;

  const folder = folderPath?.toLowerCase();
  if (!folder) return;
  if (folder === "inbox") return IMAP_LABEL.INBOX;
  // Exact well-known names only (plus the leaf segment of a hierarchy, so
  // `Mail/Trash` still resolves but `Presentations` does not).
  const leaf = folder.split(/[/.]/).pop() || folder;
  if (leaf === "trash" || leaf === "deleted items" || leaf === "deleted")
    return IMAP_LABEL.TRASH;
  if (
    leaf === "junk" ||
    leaf === "spam" ||
    leaf === "junk mail" ||
    leaf === "junk e-mail"
  )
    return IMAP_LABEL.SPAM;
  if (
    leaf === "sent" ||
    leaf === "sent items" ||
    leaf === "sent messages" ||
    leaf === "sent mail"
  )
    return IMAP_LABEL.SENT;
  if (leaf === "drafts" || leaf === "draft") return IMAP_LABEL.DRAFT;
  return;
}
