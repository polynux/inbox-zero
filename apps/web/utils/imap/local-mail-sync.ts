import { z } from "zod";
import type { ImapFlow } from "imapflow";
import {
  InvalidMailboxSyncCursorError,
  decodeMailboxSyncCursor,
  encodeMailboxSyncCursor,
} from "@/utils/email/mailbox-sync";
import type { MailboxSyncPage } from "@/utils/email/types";
import { withLocalMailSyncBudget } from "@/utils/email/local-mail-sync-budget";
import { withImapConnection } from "@/utils/imap/client";
import { listFoldersAsOutlookFolders } from "@/utils/imap/folder";
import { encodeImapMessageId } from "@/utils/imap/message-id";
import { fetchMessagesByUids, searchImapMessages } from "@/utils/imap/message";
import type { ImapCredentialConfig } from "@/utils/imap/types";
import type { Logger } from "@/utils/logger";

const cursorPayloadSchema = z.object({
  emailAccountId: z.string().min(1),
  folder: z.string().min(1),
  uidValidity: z.number().int().optional(),
  lastUid: z.number().int().nonnegative(),
  knownUids: z.string().optional(),
});

function encodeImapCursor(
  emailAccountId: string,
  folder: string,
  uidValidity: number | undefined,
  lastUid: number,
  knownUids: string | undefined,
): string {
  return encodeMailboxSyncCursor({
    version: 1,
    provider: "imap",
    state: { emailAccountId, folder, uidValidity, lastUid, knownUids },
  });
}

type SyncInput = {
  config: ImapCredentialConfig;
  logger: Logger;
  emailAccountId: string;
  limit: number;
  priority?: "backfill" | "current";
};

/**
 * UID-based folder-delta sync for IMAP. The cursor tracks the highest seen
 * UID per folder plus a compact range-set of the UIDs the local mailbox holds;
 * UIDVALIDITY changes invalidate it (server-side UID reset).
 *
 * IMAP has no tombstone feed (Bridge advertises neither CONDSTORE nor QRESYNC),
 * so a delta page compares `UID SEARCH ALL` against that range-set and reports
 * the missing UIDs as deleted.
 */
export async function getImapMailboxSyncPage(
  input: SyncInput & {
    cursor?: string;
    after?: Date;
    folderId?: string;
  },
): Promise<MailboxSyncPage> {
  const folder = input.folderId || "INBOX";
  return withImapConnection(input.config, async (client) => {
    const mailbox = await client.mailboxOpen(folder, { readOnly: true });
    const uidValidity =
      mailbox.uidValidity === undefined
        ? undefined
        : Number(mailbox.uidValidity);
    const uidNext = mailbox.uidNext === undefined ? 0 : Number(mailbox.uidNext);

    if (input.cursor) {
      const decoded = decodeMailboxSyncCursor(input.cursor, "imap");
      const payload = cursorPayloadSchema.parse(decoded.state);
      if (
        payload.emailAccountId !== input.emailAccountId ||
        payload.folder !== folder
      )
        throw new InvalidMailboxSyncCursorError();
      // A UIDVALIDITY change means the server renumbered UIDs; the local
      // mailbox must drop everything and resnapshot.
      if (
        payload.uidValidity !== undefined &&
        uidValidity !== undefined &&
        payload.uidValidity !== uidValidity
      )
        return getImapSnapshotPage(client, {
          folder,
          after: input.after,
          limit: input.limit,
          uidValidity,
          uidNext,
          emailAccountId: input.emailAccountId,
        });
      return getImapDeltaPage(client, {
        folder,
        limit: input.limit,
        uidValidity,
        uidNext,
        lastUid: payload.lastUid,
        knownUids: payload.knownUids,
        emailAccountId: input.emailAccountId,
        logger: input.logger,
      });
    }

    if (!input.after)
      throw new Error("after is required for initial mailbox sync");
    return getImapSnapshotPage(client, {
      folder,
      after: input.after,
      limit: input.limit,
      uidValidity,
      uidNext,
      emailAccountId: input.emailAccountId,
    });
  });
}

async function getImapDeltaPage(
  client: ImapFlow,
  options: {
    folder: string;
    limit: number;
    uidValidity: number | undefined;
    uidNext: number;
    lastUid: number;
    knownUids: string | undefined;
    emailAccountId: string;
    logger: Logger;
  },
): Promise<MailboxSyncPage> {
  // UID-only search, so the diff below costs one round trip and no fetches.
  const serverUids = await searchImapMessages(client, { all: true });
  const serverUidSet = new Set(serverUids);
  // A cursor written before knownUids existed bootstraps from the server's
  // current UIDs: reporting a deletion the local mailbox never received is a
  // harmless no-op, while missing one would leave stale mail behind.
  const knownUidSet =
    options.knownUids === undefined
      ? new Set(serverUids)
      : parseUidRanges(options.knownUids);
  const deletedMessageIds = [...knownUidSet]
    .filter((uid) => !serverUidSet.has(uid))
    .sort((a, b) => a - b)
    .map((uid) => encodeImapMessageId(options.folder, uid));

  if (deletedMessageIds.length > 0) {
    options.logger.info("IMAP folder delta reports deleted messages", {
      emailAccountId: options.emailAccountId,
      folder: options.folder,
      count: deletedMessageIds.length,
    });
  }

  const knownUids = formatUidRanges(serverUids);

  if (options.uidNext <= options.lastUid + 1) {
    return {
      cursor: encodeImapCursor(
        options.emailAccountId,
        options.folder,
        options.uidValidity,
        options.lastUid,
        knownUids,
      ),
      deletedMessageIds,
      hasMore: false,
      reset: false,
      upsertedMessages: [],
    };
  }

  // searchImapMessages returns newest-first; page oldest-first so the cursor
  // always advances past everything returned (no gaps, no duplicates).
  // RFC 3501: `UID n:*` also matches the highest existing UID when n is above
  // it, so an idle mailbox re-reports its newest message; drop anything the
  // cursor already covered instead of trusting the search result.
  const newUids = (
    await searchImapMessages(client, {
      uid: `${options.lastUid + 1}:*`,
    })
  )
    .filter((uid) => uid > options.lastUid)
    .reverse();
  const pageUids = newUids.slice(0, options.limit);
  const messages = await fetchMessagesByUids(client, pageUids);
  const maxUid = pageUids.reduce(
    (highest, uid) => (uid > highest ? uid : highest),
    options.lastUid,
  );

  return {
    cursor: encodeImapCursor(
      options.emailAccountId,
      options.folder,
      options.uidValidity,
      maxUid,
      knownUids,
    ),
    deletedMessageIds,
    hasMore: newUids.length > pageUids.length,
    reset: false,
    upsertedMessages: messages,
  };
}

async function getImapSnapshotPage(
  client: ImapFlow,
  options: {
    folder: string;
    after?: Date;
    limit: number;
    uidValidity: number | undefined;
    uidNext: number;
    emailAccountId: string;
  },
): Promise<MailboxSyncPage> {
  const criteria: Record<string, unknown> = options.after
    ? { since: options.after }
    : { all: true };
  const allUids = (await searchImapMessages(client, criteria)).reverse();
  const pageUids = allUids.slice(0, options.limit);
  const messages = await fetchMessagesByUids(client, pageUids);
  // An empty window still has to advance the cursor past everything that
  // already exists, otherwise the next delta lists the whole mailbox.
  const maxUid =
    pageUids.length > 0
      ? pageUids.reduce((highest, uid) => (uid > highest ? uid : highest), 0)
      : Math.max(options.uidNext - 1, 0);
  return {
    cursor: encodeImapCursor(
      options.emailAccountId,
      options.folder,
      options.uidValidity,
      maxUid,
      formatUidRanges(allUids),
    ),
    deletedMessageIds: [],
    hasMore: allUids.length > pageUids.length,
    reset: true,
    upsertedMessages: messages,
  };
}

export async function getImapMailFoldersPage(
  input: SyncInput & { parentFolderId?: string; cursor?: string },
) {
  if (!input.emailAccountId) throw new Error("Invalid local mail sync input");
  return withLocalMailSyncBudget(
    {
      emailAccountId: input.emailAccountId,
      provider: "imap",
      cost: 1,
      priority: input.priority ?? "backfill",
    },
    () =>
      withImapConnection(input.config, async (client) => {
        const folders = await listFoldersAsOutlookFolders(client);
        return {
          folders: folders.map((folder) => ({
            id: folder.id,
            // IMAP hierarchy is flattened: OutlookFolder carries no parent.
            parentFolderId: undefined as string | undefined,
            displayName: folder.displayName,
            childFolderCount: folder.childFolderCount ?? 0,
            isHidden: false,
            totalItemCount: folder.totalItemCount ?? 0,
            unreadItemCount: folder.unreadItemCount ?? 0,
          })),
          nextCursor: undefined as string | undefined,
        };
      }),
  );
}

// Compact range-set for the cursor: ascending UIDs grouped into "a-b" runs.
function formatUidRanges(uids: number[]): string {
  const sorted = [...uids].sort((a, b) => a - b);
  const ranges: string[] = [];
  let start: number | undefined;
  let end: number | undefined;

  for (const uid of sorted) {
    if (start === undefined || end === undefined) {
      start = uid;
      end = uid;
    } else if (uid === end + 1) {
      end = uid;
    } else {
      ranges.push(start === end ? `${start}` : `${start}-${end}`);
      start = uid;
      end = uid;
    }
  }
  if (start !== undefined && end !== undefined) {
    ranges.push(start === end ? `${start}` : `${start}-${end}`);
  }

  return ranges.join(",");
}

function parseUidRanges(ranges: string): Set<number> {
  const uids = new Set<number>();

  for (const part of ranges.split(",")) {
    if (!part) continue;
    const [startText, endText] = part.split("-");
    const start = Number(startText);
    const end = endText === undefined ? start : Number(endText);
    if (
      !Number.isInteger(start) ||
      !Number.isInteger(end) ||
      start < 1 ||
      end < start
    )
      continue;
    for (let uid = start; uid <= end; uid++) uids.add(uid);
  }

  return uids;
}
