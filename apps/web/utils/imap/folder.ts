import type { ImapFlow, ListResponse } from "imapflow";
import type { EmailLabel } from "@/utils/email/types";
import type { OutlookFolder } from "@/utils/outlook/folders";

// Well-known IMAP special-use folders (RFC 6154)
const SPECIAL_USE_MAP: Record<string, string> = {
  "\\Inbox": "INBOX",
  "\\Sent": "SENT",
  "\\Drafts": "DRAFTS",
  "\\Trash": "TRASH",
  "\\Junk": "SPAM",
  "\\Archive": "ARCHIVE",
  "\\All": "ALL",
  "\\Flagged": "STARRED",
};

/**
 * List all IMAP mailbox folders and convert them to EmailLabel format.
 */
export async function listFolders(client: ImapFlow): Promise<EmailLabel[]> {
  const mailboxes = await client.list();
  return mailboxes.map(convertMailboxToLabel);
}

/**
 * List IMAP mailboxes in OutlookFolder format (for getFolders() compatibility).
 */
export async function listFoldersAsOutlookFolders(
  client: ImapFlow,
): Promise<OutlookFolder[]> {
  const mailboxes = await client.list();
  return mailboxes.map((mb) => ({
    id: mb.path,
    displayName: mb.name,
    childFolders: [],
    childFolderCount: 0,
    unreadItemCount: 0,
    totalItemCount: 0,
  }));
}

function convertMailboxToLabel(mailbox: ListResponse): EmailLabel {
  const specialUse = mailbox.specialUse
    ? SPECIAL_USE_MAP[mailbox.specialUse]
    : undefined;

  return {
    id: mailbox.path,
    name: mailbox.name,
    type: specialUse ? "system" : "user",
    labelListVisibility: mailbox.listed ? "labelShow" : "labelHide",
    messageListVisibility: "show",
  };
}

/**
 * Find or create a folder by name.
 */
export async function getOrCreateFolder(
  client: ImapFlow,
  folderName: string,
): Promise<string> {
  const mailboxes = await client.list();
  const existing = mailboxes.find(
    (mb) =>
      mb.path.toLowerCase() === folderName.toLowerCase() ||
      mb.name.toLowerCase() === folderName.toLowerCase(),
  );

  if (existing) return existing.path;

  await client.mailboxCreate(folderName);
  return folderName;
}

/**
 * Move a message to a different folder. imapflow resolves a missing UID to
 * an empty range and reports success, so this returns nothing; callers that
 * need certainty must verify presence first.
 */
export async function moveMessageToFolder(
  client: ImapFlow,
  uid: number,
  targetFolder: string,
): Promise<void> {
  await client.messageMove(String(uid), targetFolder, { uid: true });
}

/**
 * Find the archive folder path. IMAP servers may use different names.
 */
export async function findArchiveFolder(client: ImapFlow): Promise<string> {
  const mailboxes = await client.list();

  // Look for special-use Archive flag first
  const archive = mailboxes.find((mb) => mb.specialUse === "\\Archive");
  if (archive) return archive.path;

  // Fall back to common archive folder names
  const archiveNames = ["Archive", "All Mail", "All"];
  for (const name of archiveNames) {
    const found = mailboxes.find(
      (mb) => mb.name.toLowerCase() === name.toLowerCase(),
    );
    if (found) return found.path;
  }

  // Create an Archive folder if none exists
  await client.mailboxCreate("Archive");
  return "Archive";
}

/**
 * Find the trash folder path.
 * Returns null when the server lists no trash mailbox.
 */
export async function findTrashFolder(
  client: ImapFlow,
): Promise<string | null> {
  const mailboxes = await client.list();

  const trash = mailboxes.find((mb) => mb.specialUse === "\\Trash");
  if (trash) return trash.path;

  const trashNames = ["Trash", "Deleted Items", "Deleted"];
  for (const name of trashNames) {
    const found = mailboxes.find(
      (mb) => mb.name.toLowerCase() === name.toLowerCase(),
    );
    if (found) return found.path;
  }

  return null;
}

/**
 * Find the sent folder path.
 * Returns null when the server lists no sent mailbox.
 */
export async function findSentFolder(client: ImapFlow): Promise<string | null> {
  const mailboxes = await client.list();

  const sent = mailboxes.find((mb) => mb.specialUse === "\\Sent");
  if (sent) return sent.path;

  const sentNames = ["Sent", "Sent Items", "Sent Messages"];
  for (const name of sentNames) {
    const found = mailboxes.find(
      (mb) => mb.name.toLowerCase() === name.toLowerCase(),
    );
    if (found) return found.path;
  }

  return null;
}

/**
 * Find the drafts folder path.
 * Returns null when the server lists no drafts mailbox.
 */
export async function findDraftsFolder(
  client: ImapFlow,
): Promise<string | null> {
  const mailboxes = await client.list();

  const drafts = mailboxes.find((mb) => mb.specialUse === "\\Drafts");
  if (drafts) return drafts.path;

  const draftNames = ["Drafts", "Draft"];
  for (const name of draftNames) {
    const found = mailboxes.find(
      (mb) => mb.name.toLowerCase() === name.toLowerCase(),
    );
    if (found) return found.path;
  }

  return null;
}

/**
 * Find the all-mail mailbox, listing first.
 */
export async function findAllMailFolder(
  client: ImapFlow,
): Promise<string | null> {
  return findAllMailInList(await client.list());
}

/**
 * Pure scan over one LIST result, so callers that already hold the listing
 * (e.g. resolveSearchFolders) don't LIST twice. Proton Bridge's "All Mail"
 * carries no \All flag, so fall back to an exact path match ("All Mail"
 * itself or a path ending in "/All Mail", e.g. "[Gmail]/All Mail") only
 * when no mailbox advertises \All.
 */
export function findAllMailInList(mailboxes: ListResponse[]): string | null {
  const flagged = mailboxes.find((mb) => mb.specialUse === "\\All");
  if (flagged) return flagged.path;

  const named = mailboxes.find(
    (mb) => mb.path === "All Mail" || mb.path.endsWith("/All Mail"),
  );
  if (named) return named.path;

  return null;
}
/**
 * Find the junk folder path.
 * Returns null when the server lists no junk mailbox.
 */
export async function findJunkFolder(client: ImapFlow): Promise<string | null> {
  const mailboxes = await client.list();

  const junk = mailboxes.find((mb) => mb.specialUse === "\\Junk");
  if (junk) return junk.path;

  const junkNames = ["Junk", "Spam", "Junk Mail", "Junk E-Mail"];
  for (const name of junkNames) {
    const found = mailboxes.find(
      (mb) => mb.name.toLowerCase() === name.toLowerCase(),
    );
    if (found) return found.path;
  }

  return null;
}

// Proton label namespace is really the whole Bridge layout: top-level
// `Folders`, `Labels` and `All Mail` with no `\All` flag anywhere (live
// Bridge probe 2026-10-31). A message lives in exactly one real folder and
// additionally appears in Labels/* mailboxes; COPY into Labels/X applies a
// label, expunging from Labels/X removes it. One stray user folder named
// `Labels` or `All Mail` must not flip label or thread-resolution behavior.
export async function isProtonBridgeLayout(client: ImapFlow): Promise<boolean> {
  return isProtonBridgeLayoutInList(await client.list());
}

// Single predicate for the Proton Bridge mailbox layout. `delimiter` splits
// nested paths, so `[Gmail]/All Mail` style subfolders never count as
// top-level.
export function isProtonBridgeLayoutInList(mailboxes: ListResponse[]): boolean {
  const topLevel = new Set(
    mailboxes
      .filter((mb) => !mb.path.includes(mb.delimiter || "/"))
      .map((mb) => mb.path),
  );
  if (
    !topLevel.has("INBOX") ||
    !topLevel.has("Folders") ||
    !topLevel.has("Labels") ||
    !topLevel.has("All Mail")
  )
    return false;
  return !mailboxes.some((mb) => mb.specialUse === "\\All");
}

export async function listLabelMailboxes(client: ImapFlow): Promise<string[]> {
  const mailboxes = await client.list();
  if (!mailboxes.some((mb) => mb.path === "Labels")) return [];
  return mailboxes
    .filter((mb) => mb.path.startsWith("Labels/"))
    .map((mb) => mb.path);
}

export function labelNameFromPath(path: string): string {
  return path.startsWith("Labels/") ? path.slice("Labels/".length) : path;
}

export function labelPathFromName(name: string): string {
  return name.startsWith("Labels/") ? name : `Labels/${name}`;
}
