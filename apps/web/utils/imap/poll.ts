import { env } from "@/env";
import { extractEmailAddress } from "@/utils/email";
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
// A whole run must finish inside the scheduler's request timeout; whatever is
// left is picked up by the next tick, because the cursor advances per message.
export const POLL_TIME_BUDGET_MS = 50_000;
// Overlapping ticks (a retry, a manual call) must not process the same mail.
// The claim expires so a killed run cannot block the account forever.
const POLL_LOCK_MS = 5 * 60_000;

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
  options?: { deadline?: number },
): Promise<PollResult> {
  const log = logger || defaultLogger;
  const deadline = options?.deadline;

  const lockedUntil = await claimImapPoll(emailAccountId);
  if (!lockedUntil) {
    log.info("Skipping IMAP poll; another poll holds the account lock", {
      emailAccountId,
    });
    return { emailAccountId, newMessages: 0 };
  }

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
        select: {
          id: true,
          lastSeenUid: true,
          uidValidity: true,
          pollFailureUid: true,
          pollFailureCount: true,
        },
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
            pollFailureUid: null,
            pollFailureCount: 0,
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
            pollFailureUid: null,
            pollFailureCount: 0,
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
      const ownAddress = credentials.email.toLowerCase();

      // Oldest first, and the cursor moves once per handled UID so a run that
      // is killed mid-batch never replays the mail it already processed.
      let newMessages = 0;
      for (const uid of newUids) {
        if (deadline !== undefined && Date.now() >= deadline) {
          log.warn("Stopping IMAP poll at the time budget", {
            emailAccountId,
            uid,
          });
          break;
        }

        const messageId = encodeImapMessageId("INBOX", uid);

        try {
          const message = await fetchMessageByUid(client, uid);
          if (!message) throw new Error("IMAP message could not be fetched");

          // Same rule Outlook applies: only inbox and sent items reach
          // automation, and a copy of the account's own mail is not new mail
          // for it to act on.
          const fromAddress = extractEmailAddress(message.headers?.from ?? "");
          const isOwnSent =
            message.labelIds?.includes("SENT") ||
            (fromAddress !== "" && fromAddress.toLowerCase() === ownAddress);

          if (isOwnSent) {
            log.info("Skipping IMAP message the account sent itself", {
              emailAccountId,
              messageId,
            });
          } else {
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
            } else {
              newMessages++;
            }
          }
        } catch (error) {
          // Strikes are persisted: a poll is a fresh invocation (serverless or
          // cron container), so an in-memory counter would never reach three.
          const failures =
            credential.pollFailureUid === uid
              ? credential.pollFailureCount + 1
              : 1;
          await prisma.imapCredential.update({
            where: { id: credential.id },
            data: { pollFailureUid: uid, pollFailureCount: failures },
          });
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
        }

        await advanceImapPollCursor(credential.id, uid, uidValidityUpdate);
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
  } finally {
    await releaseImapPoll(emailAccountId, lockedUntil);
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

  const deadline = Date.now() + POLL_TIME_BUDGET_MS;
  const results: PollResult[] = [];
  for (const account of imapAccounts) {
    if (Date.now() >= deadline) {
      log.warn("Stopping IMAP poll run at the time budget", {
        remainingAccounts: imapAccounts.length - results.length,
      });
      break;
    }
    results.push(await pollImapAccount(account.id, log, { deadline }));
  }

  return results;
}

/**
 * Move the poll cursor past one handled UID and clear its strikes. The
 * `lastSeenUid < uid` guard keeps the cursor monotonic if two runs ever race.
 */
async function advanceImapPollCursor(
  credentialId: string,
  uid: number,
  uidValidityUpdate: { uidValidity?: bigint },
): Promise<void> {
  await prisma.imapCredential.updateMany({
    where: { id: credentialId, lastSeenUid: { lt: uid } },
    data: {
      lastSeenUid: uid,
      ...uidValidityUpdate,
      lastPolledAt: new Date(),
      pollFailureUid: null,
      pollFailureCount: 0,
    },
  });
}

/**
 * Claim the account's poll lock. Returns the expiry to release with, or null
 * when another poll already holds it.
 */
async function claimImapPoll(emailAccountId: string): Promise<Date | null> {
  const now = new Date();
  const lockedUntil = new Date(now.getTime() + POLL_LOCK_MS);
  const { count } = await prisma.imapCredential.updateMany({
    where: {
      account: { emailAccount: { id: emailAccountId } },
      OR: [{ pollLockedUntil: null }, { pollLockedUntil: { lt: now } }],
    },
    data: { pollLockedUntil: lockedUntil },
  });
  return count === 0 ? null : lockedUntil;
}

/**
 * Release this run's claim. A claim another run took over after ours expired
 * is left alone.
 */
async function releaseImapPoll(
  emailAccountId: string,
  lockedUntil: Date,
): Promise<void> {
  await prisma.imapCredential.updateMany({
    where: {
      account: { emailAccount: { id: emailAccountId } },
      pollLockedUntil: lockedUntil,
    },
    data: { pollLockedUntil: null },
  });
}
