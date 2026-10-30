import type { ImapFlow } from "imapflow";
import { findDraftsFolder } from "@/utils/imap/folder";
import { encodeImapMessageId } from "@/utils/imap/message-id";
export async function saveDraft(
  client: ImapFlow,
  options: {
    from: string;
    to: string;
    subject: string;
    html?: string;
    text?: string;
    inReplyTo?: string;
    references?: string;
  },
): Promise<string> {
  let draftsFolder = await findDraftsFolder(client);

  if (!draftsFolder) {
    try {
      await client.mailboxCreate("Drafts");
    } catch {
      throw new Error(
        "No Drafts folder found on the IMAP server and creating one failed",
      );
    }
    draftsFolder = await findDraftsFolder(client);
    if (!draftsFolder) {
      throw new Error(
        "No Drafts folder found on the IMAP server and creating one failed",
      );
    }
  }

  const messageParts: string[] = [
    `From: ${stripHeaderBreaks(options.from)}`,
    `To: ${stripHeaderBreaks(options.to)}`,
    `Subject: ${stripHeaderBreaks(options.subject)}`,
    `Date: ${new Date().toUTCString()}`,
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
    ["\\Draft"],
    new Date(),
  );

  if (!result) throw new Error("Failed to save draft: APPEND rejected");
  // APPEND without UIDPLUS reports no UID: the draft exists but its id is
  // unknowable, so fail loudly instead of returning a UIDVALIDITY-shaped
  // or "unknown" id that later acts on the wrong message.
  if (!result.uid)
    throw new Error(
      "Failed to save draft: server did not return a UID (no UIDPLUS support)",
    );
  return encodeImapMessageId(draftsFolder, result.uid);
}

// A raw CR, LF or NUL in a header value would start a new header line.
function stripHeaderBreaks(value: string) {
  return value.replace(/[\r\n\0]/g, "");
}
