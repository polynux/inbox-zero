import {
  canonicalizeEmailAddress,
  extractNameFromEmail,
  splitRecipientList,
} from "@/utils/email";
import { GmailLabel } from "@/utils/gmail/label";
import type { ParsedMessageHeaders } from "@/utils/types";

type ParticipantMessage = {
  headers: Pick<ParsedMessageHeaders, "from" | "to">;
  labelIds?: string[] | null;
};

export function getThreadParticipantNames(
  messages: ParticipantMessage[],
  userEmail: string,
) {
  const normalizedUserEmail = canonicalizeEmailAddress(userEmail);
  const senders = new Map<string, string>();
  const nonDraftMessages = messages.filter(
    (message) => !message.labelIds?.includes(GmailLabel.DRAFT),
  );

  for (const message of nonDraftMessages) {
    addParticipant(senders, message.headers.from, normalizedUserEmail);
  }

  // A sent-only thread names its recipients however the From reads: the
  // sender is one of the account's own addresses (aliases, plus-addressing,
  // a renamed mailbox), which never equals the login email exactly. List rows
  // share union labels across their synthesized messages, so SENT on every
  // message only means sent-only when no other mailbox role is present.
  const hasOnlySentMessages =
    nonDraftMessages.length > 0 &&
    nonDraftMessages.every(
      (message) =>
        message.labelIds?.includes(GmailLabel.SENT) &&
        !message.labelIds?.some((labelId) =>
          [GmailLabel.INBOX, GmailLabel.TRASH, GmailLabel.SPAM].includes(
            labelId,
          ),
        ),
    );
  const hasOnlyAccountOwnerAsSender =
    senders.size === 1 && senders.has(normalizedUserEmail);
  if (
    senders.size > 0 &&
    !hasOnlyAccountOwnerAsSender &&
    !hasOnlySentMessages
  ) {
    return [...senders.values()];
  }

  const recipients = new Map<string, string>();
  const recipientMessages = senders.size > 0 ? nonDraftMessages : messages;
  for (const message of recipientMessages) {
    for (const recipient of splitRecipientList(message.headers.to)) {
      addParticipant(recipients, recipient, normalizedUserEmail);
    }
  }
  // An outgoing-only thread should name the other side, not "me".
  recipients.delete(normalizedUserEmail);

  return recipients.size ? [...recipients.values()] : [...senders.values()];
}

function addParticipant(
  participants: Map<string, string>,
  header: string,
  normalizedUserEmail: string,
) {
  const email = canonicalizeEmailAddress(header);
  const key = email || header.trim().toLowerCase();
  if (!key || participants.has(key)) return;

  const name =
    email && email === normalizedUserEmail
      ? "me"
      : extractNameFromEmail(header) || email;
  if (name) participants.set(key, name);
}
