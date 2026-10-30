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
 * Move a message to a different folder.
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
 * Find the junk/spam folder path.
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
