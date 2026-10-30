import type { ImapFlow } from "imapflow";
import { findDraftsFolder } from "@/utils/imap/folder";

/**
 * Save a draft message to the IMAP Drafts folder.
 * Returns the UID of the appended message.
 */
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
  const draftsFolder = await findDraftsFolder(client);

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
  return result.uid
    ? String(result.uid)
    : result.uidValidity
      ? String(result.uidValidity)
      : "unknown";
}

// A raw CR, LF or NUL in a header value would start a new header line.
function stripHeaderBreaks(value: string) {
  return value.replace(/[\r\n\0]/g, "");
}
