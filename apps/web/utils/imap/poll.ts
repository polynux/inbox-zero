import { env } from "@/env";
import { createEmailProvider } from "@/utils/email/provider";
import { withImapConnection } from "@/utils/imap/client";
import { getImapCredentials } from "@/utils/imap/credential";
import { encodeImapMessageId } from "@/utils/imap/message-id";
import { fetchMessageByUid, searchImapMessages } from "@/utils/imap/message";
import type { Logger } from "@/utils/logger";
import { createScopedLogger } from "@/utils/logger";
import prisma from "@/utils/prisma";
import { withMessageProcessingLock } from "@/utils/redis/message-processing";
import { processHistoryItem } from "@/utils/webhook/process-history-item";
import {
  getWebhookEmailAccount,
  validateWebhookAccount,
} from "@/utils/webhook/validate-webhook-account";

const defaultLogger = createScopedLogger("imap/poll");

// A poll feeds one page of new mail per account so a long-idle mailbox cannot
// turn a single cron tick into a mailbox-wide run.
const MAX_MESSAGES_PER_POLL = 50;
// One bad message must not pin the cursor forever; after this many polls it is
// skipped (and logged) so later mail still reaches automation.
const MAX_PROCESSING_FAILURES = 3;

const processingFailures = new Map<string, number>();

interface PollResult {
  emailAccountId: string;
  error?: string;
  newMessages: number;
}

/**
 * Poll a single IMAP account for new messages and hand each new INBOX message
 * to the shared history processor (same rules/AI path as Gmail and Outlook).
 * Compares the current UIDNEXT with the stored lastSeenUid.
 */
export async function pollImapAccount(
  emailAccountId: string,
  logger?: Logger,
): Promise<PollResult> {
  const log = logger || defaultLogger;

  try {
    const credentials = await getImapCredentials(emailAccountId);

    // Same gate Gmail/Outlook webhooks use: disconnected, non-premium or
    // no-AI accounts and accounts without rules/filing have nothing to feed.
    const emailAccount = await getWebhookEmailAccount(
      { email: credentials.email.toLowerCase() },
      log,
    );
    const validation = await validateWebhookAccount(emailAccount, log);
    if (!validation.success) {
      return { emailAccountId, newMessages: 0 };
    }

    const result = await withImapConnection(credentials, async (client) => {
      const mailbox = await client.mailboxOpen("INBOX", { readOnly: true });

      // Get the stored last seen UID
      const credential = await prisma.imapCredential.findFirst({
        where: {
          account: { emailAccount: { id: emailAccountId } },
        },
        select: { id: true, lastSeenUid: true, uidValidity: true },
      });

      if (!credential) {
        throw new Error("IMAP credential not found");
      }

      const uidNext =
        mailbox.uidNext === undefined ? null : Number(mailbox.uidNext);
      const uidValidity =
        mailbox.uidValidity === undefined ? null : BigInt(mailbox.uidValidity);
      const uidValidityUpdate = uidValidity == null ? {} : { uidValidity };

      // First poll (null cursor): baseline only, report 0 new. The whole
      // mailbox predates linking, so reporting it all as new would archive
      // or process years of history on first run. A server that omits
      // UIDNEXT leaves the baseline unset instead of storing -1.
      if (credential.lastSeenUid == null) {
        await prisma.imapCredential.update({
          where: { id: credential.id },
          data: {
            ...(uidNext == null ? {} : { lastSeenUid: uidNext - 1 }),
            ...uidValidityUpdate,
            lastPolledAt: new Date(),
          },
        });
        return { newMessages: 0 };
      }

      // A different UIDVALIDITY means the server renumbered UIDs (RFC 3501
      // 2.3.1.1): every stored UID is meaningless, so re-baseline instead of
      // reporting the whole mailbox as new.
      if (
        credential.uidValidity != null &&
        uidValidity != null &&
        credential.uidValidity !== uidValidity
      ) {
        log.warn("IMAP UIDVALIDITY changed; resetting poll cursor", {
          emailAccountId,
          previousUidValidity: credential.uidValidity.toString(),
          uidValidity: uidValidity.toString(),
        });
        await prisma.imapCredential.update({
          where: { id: credential.id },
          data: {
            ...(uidNext == null ? {} : { lastSeenUid: uidNext - 1 }),
            ...uidValidityUpdate,
            lastPolledAt: new Date(),
          },
        });
        return { newMessages: 0 };
      }

      if (uidNext == null) return { newMessages: 0 };

      const lastSeenUid = credential.lastSeenUid;
      if (uidNext <= lastSeenUid + 1) {
        // No new messages
        if (uidValidity != null && credential.uidValidity == null) {
          await prisma.imapCredential.update({
            where: { id: credential.id },
            data: { ...uidValidityUpdate, lastPolledAt: new Date() },
          });
        }
        return { newMessages: 0 };
      }

      // Search for messages with UID > lastSeenUid. RFC 3501: `UID n:*` also
      // matches the highest existing UID when n is above it, so drop anything
      // below the cursor instead of trusting the search result.
      const newUids = (
        await searchImapMessages(client, {
          uid: `${lastSeenUid + 1}:*`,
        })
      )
        .filter((uid) => uid > lastSeenUid)
        .reverse()
        .slice(0, MAX_MESSAGES_PER_POLL);

      if (newUids.length === 0) {
        return { newMessages: 0 };
      }

      log.info("Found new IMAP messages", {
        emailAccountId,
        count: newUids.length,
      });

      const provider = await createEmailProvider({
        emailAccountId,
        provider: "imap",
        logger: log,
      });
      const processWithLock = Boolean(
        env.REDIS_HTTP_URL && env.REDIS_HTTP_TOKEN,
      );

      // Oldest first, and the cursor only moves past UIDs that were handled.
      let cursor = lastSeenUid;
      let newMessages = 0;
      for (const uid of newUids) {
        const messageId = encodeImapMessageId("INBOX", uid);
        const failureKey = `${emailAccountId}:${messageId}`;

        try {
          const message = await fetchMessageByUid(client, uid);
          if (!message) throw new Error("IMAP message could not be fetched");

          const processMessage = async () => {
            await processHistoryItem(
              { messageId, message },
              {
                provider,
                isNewMessage: true,
                emailAccount: {
                  ...validation.data.emailAccount,
                  account: { provider: "imap" },
                },
                hasAutomationRules: validation.data.hasAutomationRules,
                hasAiAccess: validation.data.hasAiAccess,
                rules: validation.data.emailAccount.rules,
                logger: log,
              },
            );
          };

          let skipped = false;
          if (processWithLock) {
            ({ skipped } = await withMessageProcessingLock(
              { userEmail: credentials.email, messageId },
              processMessage,
            ));
          } else {
            await processMessage();
          }

          if (skipped) {
            log.info("Skipping. Message already being processed.", {
              emailAccountId,
              messageId,
            });
          }
        } catch (error) {
          const failures = (processingFailures.get(failureKey) ?? 0) + 1;
          processingFailures.set(failureKey, failures);
          const errorMessage =
            error instanceof Error ? error.message : "Unknown error";

          if (failures < MAX_PROCESSING_FAILURES) {
            log.warn("Stopping IMAP poll before a message that failed", {
              emailAccountId,
              uid,
              failures,
              error: errorMessage,
            });
            break;
          }

          log.warn("Skipping IMAP message after repeated failures", {
            emailAccountId,
            uid,
            failures,
            error: errorMessage,
          });
          processingFailures.delete(failureKey);
          cursor = uid;
          continue;
        }

        processingFailures.delete(failureKey);
        cursor = uid;
        newMessages++;
      }

      if (cursor !== lastSeenUid) {
        await prisma.imapCredential.update({
          where: { id: credential.id },
          data: {
            lastSeenUid: cursor,
            ...uidValidityUpdate,
            lastPolledAt: new Date(),
          },
        });
      }

      return { newMessages };
    });

    return {
      emailAccountId,
      ...result,
    };
  } catch (error) {
    const errorMessage =
      error instanceof Error ? error.message : "Unknown error";
    log.error("Error polling IMAP account", {
      emailAccountId,
      error: errorMessage,
    });
    return {
      emailAccountId,
      newMessages: 0,
      error: errorMessage,
    };
  }
}

/**
 * Poll all active IMAP accounts.
 */
export async function pollAllImapAccounts(
  logger?: Logger,
): Promise<PollResult[]> {
  const log = logger || defaultLogger;

  const imapAccounts = await prisma.emailAccount.findMany({
    where: {
      account: {
        provider: "imap",
        disconnectedAt: null,
      },
    },
    select: { id: true },
  });

  log.info("Polling IMAP accounts", { count: imapAccounts.length });

  const results: PollResult[] = [];
  for (const account of imapAccounts) {
    const result = await pollImapAccount(account.id, log);
    results.push(result);
  }

  return results;
}
