import { SafeError } from "@/utils/error";

// Folder-qualified IMAP message ids. UIDs are unique only per mailbox
// (RFC 3501 section 2.3.1.1), so a bare UID cannot name a message
// account-wide. Encoding: `<uid>@<base64url(folder)>`. base64url carries any
// mailbox name (colons, slashes, unicode) without a separator clash, and `@`
// never appears in a decimal UID, so the split is unambiguous.
export function encodeImapMessageId(folder: string, uid: number): string {
  if (!folder)
    throw new SafeError("IMAP folder is required to build a message id");
  if (!Number.isInteger(uid) || uid <= 0)
    throw new SafeError("IMAP UID must be a positive integer");
  const encodedFolder = Buffer.from(folder, "utf8").toString("base64url");
  return `${uid}@${encodedFolder}`;
}

export function parseImapMessageId(id: string): {
  folder: string;
  uid: number;
} {
  const at = id.indexOf("@");
  if (at <= 0 || at === id.length - 1 || id.indexOf("@", at + 1) !== -1)
    throw new SafeError(
      "Message id is not a folder-qualified IMAP id (expected <uid>@<folder>)",
    );
  const uidText = id.slice(0, at);
  const folderText = id.slice(at + 1);
  if (!/^[1-9]\d*$/.test(uidText))
    throw new SafeError("Message id holds an invalid IMAP UID");
  let folder: string;
  try {
    folder = Buffer.from(folderText, "base64url").toString("utf8");
  } catch {
    throw new SafeError("Message id holds an undecodable IMAP folder");
  }
  if (!folder) throw new SafeError("Message id holds an empty IMAP folder");
  return { folder, uid: Number(uidText) };
}
