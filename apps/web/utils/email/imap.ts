import type { Attachment as MailAttachment } from "nodemailer/lib/mailer";
import type { ImapFlow, MailboxObject } from "imapflow";
import { randomBytes } from "node:crypto";
import type { ThreadsQuery } from "@/utils/threads/validation";
import type { LocalMailSyncRequest } from "@/utils/actions/local-mail-sync.validation";
import type { EmailContact } from "@/utils/email/contact";
import type { SendAsAddress } from "@/utils/email/send-as";
import type {
  BulkArchiveResult,
  BulkArchiveThread,
  DraftAttachmentUploadStart,
  DraftAttachmentsResult,
  EmailFilter,
  EmailFolderCount,
  EmailLabel,
  EmailLabelUpdate,
  EmailProvider,
  EmailSignature,
  EmailThread,
  GetThreadOptions,
  MailboxSyncPage,
  ProviderMailboxSearch,
  SentMessagePage,
} from "@/utils/email/types";
import type { ParsedMessage } from "@/utils/types";
import type { OutlookFolder } from "@/utils/outlook/folders";
import type { InboxZeroLabel } from "@/utils/label";
import type { Logger } from "@/utils/logger";
import type { ImapCredentialConfig } from "@/utils/imap/types";
import { UnsupportedImapOperationError } from "@/utils/imap/types";
import { DraftNotFoundError, SafeError } from "@/utils/error";
import { withImapConnection } from "@/utils/imap/client";
import {
  fetchInternalDatesByUids,
  fetchMessageByUid,
  fetchMessagesByUids,
  fetchRecentMessages,
  parseSearchQuery,
  searchImapMessages,
} from "@/utils/imap/message";
import {
  getImapMailFoldersPage,
  getImapMailboxSyncPage,
} from "@/utils/imap/local-mail-sync";
import type { DraftAttachmentMetadata } from "@/utils/actions/draft-attachments.validation";
import {
  findAllMailInList,
  findArchiveFolder,
  findDraftsFolder,
  findJunkFolder,
  findSentFolder,
  findTrashFolder,
  getOrCreateFolder,
  isProtonBridgeLayout,
  isProtonBridgeLayoutInList,
  labelNameFromPath,
  labelPathFromName,
  listFolders,
  listFoldersAsOutlookFolders,
  listLabelMailboxes,
  moveMessageToFolder,
} from "@/utils/imap/folder";
import {
  encodeImapMessageId,
  parseImapMessageId,
} from "@/utils/imap/message-id";
import {
  appendDraftMessage,
  buildDraftMime,
  draftNotFoundError,
  encodeImapDraftId,
  findDraftByMessageId,
  readDraftAttachmentParts,
  resolveDraftUid,
  saveDraft,
  toDraftMailAttachment,
  type DraftAttachmentPart,
  type ResolvedDraft,
} from "@/utils/imap/draft";
import {
  assertDraftAttachmentLimits,
  type DraftMessageUploadPart,
  getDraftMessageUploadLength,
} from "@/utils/email/draft-attachment-upload";
import { sendSmtpEmail } from "@/utils/imap/mail";
import { toMailerAttachments } from "@/utils/types/mail";
import {
  buildThreadId,
  isProtonInternalId,
  parseFallbackThreadId,
  parseThreadRootMessageId,
} from "@/utils/imap/thread";
import { buildThreadingHeaders } from "@/utils/email/threading";
import { extractEmailAddress, splitRecipientList } from "@/utils/email";
import { createScopedLogger } from "@/utils/logger";
import { escapeHtml, textToHtmlParagraphs } from "@/utils/string";

const defaultLogger = createScopedLogger("imap-provider");

export class ImapProvider implements EmailProvider {
  readonly name = "imap" as const;
  private readonly config: ImapCredentialConfig;
  private readonly logger: Logger;

  constructor(config: ImapCredentialConfig, logger?: Logger) {
    this.config = config;
    this.logger = logger || defaultLogger;
  }

  private async withConnection<T>(
    fn: (client: ImapFlow) => Promise<T>,
  ): Promise<T> {
    return withImapConnection(this.config, fn);
  }

  toJSON() {
    return { name: "imap", type: "imap" };
  }

  getAccessToken(): string {
    return "";
  }

  // --- Message Operations ---

  async getMessage(messageId: string): Promise<ParsedMessage> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      await client.mailboxOpen(folder, { readOnly: true });
      const msg = await fetchMessageByUid(client, uid);
      // A moved message keeps its old folder-qualified id, so like a
      // vanished thread it reads as missing (Gmail/Outlook 404) rather
      // than a 503 that retries forever.
      if (!msg) throw new ImapMessageNotFoundError(messageId);
      return msg;
    });
  }

  async getMessageByRfc822MessageId(
    rfc822MessageId: string,
  ): Promise<ParsedMessage | null> {
    return this.withConnection(async (client) => {
      const folders = await resolveSearchFolders(client);
      for (const folder of folders) {
        if (!(await openFolderSkippingNo(client, folder))) continue;
        const uids = await searchImapMessages(
          client,
          { header: { "Message-ID": rfc822MessageId } },
          1,
        );
        if (uids.length === 0) continue;
        return fetchMessageByUid(client, uids[0]);
      }
      return null;
    });
  }

  async getMessagesBatch(messageIds: string[]): Promise<ParsedMessage[]> {
    return this.withConnection(async (client) => {
      const grouped = groupMessageIdsByFolder(messageIds);
      const messages: ParsedMessage[] = [];
      for (const [folder, uids] of grouped) {
        await client.mailboxOpen(folder, { readOnly: true });
        messages.push(...(await fetchMessagesByUids(client, uids)));
      }
      return messages;
    });
  }

  async getInboxMessages(maxResults?: number): Promise<ParsedMessage[]> {
    return this.withConnection(async (client) => {
      const mailbox = await client.mailboxOpen("INBOX", { readOnly: true });
      return fetchRecentMessages(client, mailbox, maxResults || 50);
    });
  }

  async getInboxStats(): Promise<{ total: number; unread: number }> {
    return this.withConnection(async (client) => {
      const mailbox = await client.mailboxOpen("INBOX", { readOnly: true });
      const unreadUids = await searchImapMessages(client, { unseen: true });
      return {
        total: mailbox.exists || 0,
        unread: unreadUids.length,
      };
    });
  }

  async getMessagesWithPagination(options: {
    query?: string;
    maxResults?: number;
    pageToken?: string;
    folderId?: string;
    before?: Date;
    after?: Date;
    inboxOnly?: boolean;
    unreadOnly?: boolean;
    includeTrashAndJunk?: boolean;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    // Folder-scoped bootstrap pages (Archived/Trash views) read from the
    // same folder-delta sync as catch-up, so the views list what local
    // mail holds. Query pagination keeps the old fan-out search.
    if (options.folderId && !options.query && !options.unreadOnly) {
      const page = await getImapMailboxSyncPage({
        config: this.config,
        logger: this.logger,
        emailAccountId: this.config.emailAccountId,
        folderId: options.folderId,
        after: options.after ?? new Date(0),
        cursor: options.pageToken,
        limit: options.maxResults || 20,
      });
      return {
        messages: page.upsertedMessages,
        nextPageToken: page.hasMore ? page.cursor : undefined,
      };
    }
    return this.withConnection(async (client) => {
      const folders =
        options.inboxOnly === false
          ? await resolveSearchFolders(client, {
              includeTrashAndJunk: options.includeTrashAndJunk,
            })
          : ["INBOX"];

      const maxResults = options.maxResults || 20;
      const offset = options.pageToken ? Number(options.pageToken) : 0;

      const criteria: Record<string, unknown> = {};
      let attachmentOnly = false;
      if (options.query) {
        const parsed = parseSearchQuery(options.query);
        Object.assign(criteria, parsed.criteria);
        attachmentOnly = parsed.hasAttachment;
      } else {
        criteria.all = true;
      }
      if (options.before) criteria.before = options.before;
      if (options.after) criteria.since = options.after;
      if (options.unreadOnly) criteria.unseen = true;

      // UID order is append order, which Bridge does not keep in date order:
      // a sequence window (or slicing UIDs) repeats and skips pages. SEARCH
      // the full match set, rank it by INTERNALDATE desc (UID desc tie-break)
      // with one cheap fetch per match, then fetch envelopes for the page.
      const ranked: Array<{ folder: string; uid: number; time: number }> = [];
      for (const folder of folders) {
        if (!(await openFolderSkippingNo(client, folder))) continue;
        const uids = await searchImapMessages(client, criteria);
        const dated = await fetchInternalDatesByUids(client, uids);
        for (const entry of dated) ranked.push({ folder, ...entry });
      }
      ranked.sort((a, b) => b.time - a.time || b.uid - a.uid);

      // has:attachment has no SEARCH key: fetch envelopes for the ranked
      // matches, keep the ones bodyStructure flags as attachments, page that.
      let candidates = ranked;
      if (attachmentOnly) {
        const byFolder = new Map<string, number[]>();
        for (const entry of ranked) {
          const uids = byFolder.get(entry.folder) || [];
          uids.push(entry.uid);
          byFolder.set(entry.folder, uids);
        }
        const fetched: ParsedMessage[] = [];
        for (const [folder, uids] of byFolder) {
          await client.mailboxOpen(folder, { readOnly: true });
          fetched.push(...(await fetchMessagesByUids(client, uids)));
        }
        const keep = new Set(
          fetched
            .filter((m) => m.attachments && m.attachments.length > 0)
            .map((m) => m.id),
        );
        candidates = ranked.filter((entry) =>
          keep.has(encodeImapMessageId(entry.folder, entry.uid)),
        );
      }

      const page = candidates.slice(offset, offset + maxResults);
      const collected: ParsedMessage[] = [];
      // One folder open per folder on the page, not one per message.
      const pageByFolder = new Map<string, number[]>();
      for (const entry of page) {
        const uids = pageByFolder.get(entry.folder) || [];
        uids.push(entry.uid);
        pageByFolder.set(entry.folder, uids);
      }
      // Preserve the page order: fetch per folder, then re-sort by rank.
      const rank = new Map(
        page.map((entry, index) => [`${entry.folder}:${entry.uid}`, index]),
      );
      for (const [folder, uids] of pageByFolder) {
        await client.mailboxOpen(folder, { readOnly: true });
        collected.push(...(await fetchMessagesByUids(client, uids)));
      }
      collected.sort(
        (a, b) =>
          (rank.get(byFolderKeyOf(a.id)) ?? 0) -
          (rank.get(byFolderKeyOf(b.id)) ?? 0),
      );
      const nextPageToken =
        offset + maxResults < candidates.length
          ? String(offset + maxResults)
          : undefined;

      return { messages: collected, nextPageToken };
    });
  }

  async searchMessages(options: {
    query: string;
    mailboxSearch?: ProviderMailboxSearch;
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    const search = options.mailboxSearch;
    if (!search) {
      return this.getMessagesWithPagination({
        query: options.query,
        maxResults: options.maxResults,
        pageToken: options.pageToken,
      });
    }
    return this.withConnection(async (client) => {
      const folder = await mailboxSearchFolder(client, search.mailbox);
      // Like Outlook's well-known-folder mapping: a scoped mailbox searches
      // that folder only; unmappable scopes fall back to the whole mailbox.
      if (!folder) {
        return this.getMessagesWithPagination({
          query: mailboxSearchQuery(search),
          maxResults: options.maxResults,
          pageToken: options.pageToken,
          inboxOnly: false,
        });
      }
      return searchImapFolder(client, {
        folder,
        search,
        maxResults: options.maxResults,
        pageToken: options.pageToken,
      });
    });
  }

  async getMessagesFromSender(options: {
    senderEmail: string;
    maxResults?: number;
    pageToken?: string;
    before?: Date;
    after?: Date;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    return this.getMessagesWithPagination({
      query: `from:${options.senderEmail}`,
      maxResults: options.maxResults,
      pageToken: options.pageToken,
      before: options.before,
      after: options.after,
    });
  }

  async getMessagesWithAttachments(options: {
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    return this.withConnection(async (client) => {
      await client.mailboxOpen("INBOX", { readOnly: true });
      // Search for messages with attachments is not universally supported
      // Fall back to text search. UID order is append order, not date order:
      // rank the full match set by INTERNALDATE desc, then page by offset so
      // a filter miss on one page cannot strand later pages.
      const allUids = await searchImapMessages(client, { all: true });
      const dated = await fetchInternalDatesByUids(client, allUids);
      dated.sort((a, b) => b.time - a.time || b.uid - a.uid);
      const offset = options.pageToken ? Number(options.pageToken) : 0;
      const pageUids = dated
        .slice(offset, offset + (options.maxResults || 20))
        .map((entry) => entry.uid);

      const messages = await fetchMessagesByUids(client, pageUids);
      // Filter client-side for messages with attachments
      const withAttachments = messages.filter(
        (m) => m.attachments && m.attachments.length > 0,
      );

      return {
        messages: withAttachments,
        nextPageToken:
          offset + (options.maxResults || 20) < dated.length
            ? String(offset + (options.maxResults || 20))
            : undefined,
      };
    });
  }

  // --- Thread Operations ---

  async getThread(
    threadId: string,
    _options?: GetThreadOptions,
  ): Promise<EmailThread> {
    const messages = await this.getThreadMessages(threadId);
    return {
      id: threadId,
      messages,
      snippet: messages[0].snippet,
    };
  }

  async getThreadMessages(
    threadId: string,
    options?: { includeTrashAndJunk?: boolean; forMutation?: boolean },
  ): Promise<ParsedMessage[]> {
    const messages = await this.withConnection((client) =>
      fetchThreadMessages(client, threadId, options),
    );
    if (messages.length === 0) throw new ImapThreadNotFoundError(threadId);
    return messages.sort(
      (a, b) => new Date(a.date).getTime() - new Date(b.date).getTime(),
    );
  }

  async getThreadMessagesInInbox(threadId: string): Promise<ParsedMessage[]> {
    return this.getThreadMessages(threadId);
  }

  async getThreads(_folderId?: string): Promise<EmailThread[]> {
    return this.withConnection(async (client) => {
      const folder = _folderId || "INBOX";
      if ((await isProtonBridgeLayout(client)) && isFallbackViewFolder(folder))
        throw new SafeError(
          `Cannot list threads in a Proton view mailbox (${folder})`,
        );
      const mailbox = await client.mailboxOpen(folder, { readOnly: true });
      const messages = await fetchRecentMessages(client, mailbox, 50);

      // Group messages by threadId
      const threadMap = new Map<string, ParsedMessage[]>();
      for (const msg of messages) {
        const existing = threadMap.get(msg.threadId) || [];
        existing.push(msg);
        threadMap.set(msg.threadId, existing);
      }

      return [...threadMap.entries()].map(([id, msgs]) => ({
        id,
        messages: msgs.sort(
          (a, b) => new Date(a.date).getTime() - new Date(b.date).getTime(),
        ),
        snippet: msgs[msgs.length - 1].snippet,
      }));
    });
  }

  async getThreadsWithQuery(options: {
    query?: ThreadsQuery;
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ threads: EmailThread[]; nextPageToken?: string }> {
    const type =
      typeof options.query === "object" ? options.query?.type : undefined;
    const folderForType =
      type === "archive"
        ? "Archive"
        : type === "trash"
          ? "Trash"
          : type === "spam"
            ? "Junk"
            : type === "sent"
              ? "Sent"
              : undefined;
    // The Archived/Trash/Spam/Sent views list one special-use folder, so a
    // folder-qualified sync page serves them without a cross-folder search.
    if (folderForType) {
      const { messages, nextPageToken } = await this.getMessagesWithPagination({
        folderId: folderForType,
        maxResults: options.maxResults,
        pageToken: options.pageToken,
      });
      const threadMap = new Map<string, ParsedMessage[]>();
      for (const msg of messages) {
        const existing = threadMap.get(msg.threadId) || [];
        existing.push(msg);
        threadMap.set(msg.threadId, existing);
      }
      return {
        threads: [...threadMap.entries()].map(([id, msgs]) => ({
          id,
          messages: msgs,
          snippet: msgs[msgs.length - 1].snippet,
        })),
        nextPageToken,
      };
    }
    const queryStr = options.query
      ? typeof options.query === "string"
        ? options.query
        : ""
      : "";
    const { messages, nextPageToken } = await this.getMessagesWithPagination({
      query: queryStr,
      maxResults: options.maxResults,
      pageToken: options.pageToken,
    });

    // Group into threads
    const threadMap = new Map<string, ParsedMessage[]>();
    for (const msg of messages) {
      const existing = threadMap.get(msg.threadId) || [];
      existing.push(msg);
      threadMap.set(msg.threadId, existing);
    }

    const threads = [...threadMap.entries()].map(([id, msgs]) => ({
      id,
      messages: msgs,
      snippet: msgs[msgs.length - 1].snippet,
    }));

    return { threads, nextPageToken };
  }

  async getThreadsWithLabel(options: {
    labelId: string;
    maxResults?: number;
  }): Promise<EmailThread[]> {
    return this.withConnection(async (client) => {
      await client.mailboxOpen(options.labelId, { readOnly: true });
      // searchImapMessages caps newest UIDs, and UID order is append order:
      // rank the full match set by INTERNALDATE desc, then fetch the page.
      const uids = await searchImapMessages(client, { all: true });
      const dated = await fetchInternalDatesByUids(client, uids);
      dated.sort((a, b) => b.time - a.time || b.uid - a.uid);
      const pageUids = dated
        .slice(0, options.maxResults || 20)
        .map((entry) => entry.uid);
      const messages = await fetchMessagesByUids(client, pageUids);
      const threadMap = new Map<string, ParsedMessage[]>();
      for (const msg of messages) {
        const existing = threadMap.get(msg.threadId) || [];
        existing.push(msg);
        threadMap.set(msg.threadId, existing);
      }

      return [...threadMap.entries()].map(([id, msgs]) => ({
        id,
        messages: msgs,
        snippet: msgs[msgs.length - 1].snippet,
      }));
    });
  }

  async getThreadsWithParticipant(options: {
    participantEmail: string;
    maxThreads?: number;
  }): Promise<EmailThread[]> {
    const { messages } = await this.getMessagesFromSender({
      senderEmail: options.participantEmail,
      maxResults: options.maxThreads || 20,
    });

    const threadMap = new Map<string, ParsedMessage[]>();
    for (const msg of messages) {
      const existing = threadMap.get(msg.threadId) || [];
      existing.push(msg);
      threadMap.set(msg.threadId, existing);
    }

    return [...threadMap.entries()].map(([id, msgs]) => ({
      id,
      messages: msgs,
      snippet: msgs[msgs.length - 1].snippet,
    }));
  }

  async getThreadsFromSenderWithSubject(
    sender: string,
    limit: number,
  ): Promise<Array<{ id: string; snippet: string; subject: string }>> {
    const { messages } = await this.getMessagesFromSender({
      senderEmail: sender,
      maxResults: limit,
    });

    const threadMap = new Map<
      string,
      { id: string; snippet: string; subject: string }
    >();
    for (const msg of messages) {
      if (!threadMap.has(msg.threadId)) {
        threadMap.set(msg.threadId, {
          id: msg.threadId,
          snippet: msg.snippet,
          subject: msg.subject,
        });
      }
    }

    return [...threadMap.values()];
  }

  async getLatestMessageInThread(
    threadId: string,
  ): Promise<ParsedMessage | null> {
    const messages = await this.getThreadMessages(threadId);
    return messages.length > 0 ? messages[messages.length - 1] : null;
  }

  async getLatestMessageFromThreadSnapshot(
    thread: Pick<EmailThread, "id" | "messages">,
  ): Promise<ParsedMessage | null> {
    if (thread.messages.length > 0) {
      return thread.messages[thread.messages.length - 1];
    }
    return this.getLatestMessageInThread(thread.id);
  }

  async getPreviousConversationMessages(
    messageIds: string[],
  ): Promise<ParsedMessage[]> {
    return this.getMessagesBatch(messageIds);
  }

  // --- Label/Folder Operations ---

  async getLabels(_options?: {
    includeHidden?: boolean;
  }): Promise<EmailLabel[]> {
    return this.withConnection(async (client) => {
      if (await isProtonBridgeLayout(client)) {
        const paths = await listLabelMailboxes(client);
        return paths.map((path) => ({
          id: path,
          name: labelNameFromPath(path),
          type: "user",
        }));
      }
      return listFolders(client);
    });
  }

  // The UI cannot LIST from the browser, so it asks the provider whether
  // "label" detaches (Proton Bridge Labels/*) or moves (plain IMAP folders).
  async supportsLabels(): Promise<boolean> {
    return this.withConnection((client) => isProtonBridgeLayout(client));
  }

  async getLabelById(labelId: string): Promise<EmailLabel | null> {
    const labels = await this.getLabels();
    return labels.find((l) => l.id === labelId) || null;
  }

  async getLabelByName(name: string): Promise<EmailLabel | null> {
    return this.withConnection(async (client) => {
      if (await isProtonBridgeLayout(client)) {
        const path = labelPathFromName(name);
        const paths = await listLabelMailboxes(client);
        const found = paths.find(
          (candidate) =>
            candidate.toLowerCase() === path.toLowerCase() ||
            labelNameFromPath(candidate).toLowerCase() === name.toLowerCase(),
        );
        return found
          ? { id: found, name: labelNameFromPath(found), type: "user" }
          : null;
      }
      const labels = await listFolders(client);
      return (
        labels.find((l) => l.name.toLowerCase() === name.toLowerCase()) || null
      );
    });
  }

  async getOrCreateInboxZeroLabel(_key: InboxZeroLabel): Promise<EmailLabel> {
    const folderName = `InboxZero/${_key}`;
    return this.withConnection(async (client) => {
      const path = await getOrCreateFolder(client, folderName);
      return {
        id: path,
        name: folderName,
        type: "user",
      };
    });
  }

  async createLabel(name: string, _description?: string): Promise<EmailLabel> {
    return this.withConnection(async (client) => {
      if (await isProtonBridgeLayout(client)) {
        const path = labelPathFromName(name);
        const existing = await listLabelMailboxes(client);
        const canonical = existing.find(
          (p) => p.toLowerCase() === path.toLowerCase(),
        );
        if (!canonical) {
          await client.mailboxCreate(path);
        }
        const id = canonical ?? path;
        return { id, name: labelNameFromPath(id), type: "user" };
      }
      await client.mailboxCreate(name);
      return {
        id: name,
        name,
        type: "user",
      };
    });
  }

  async deleteLabel(labelId: string): Promise<void> {
    return this.withConnection(async (client) => {
      if (await isProtonBridgeLayout(client)) {
        if (!labelId.startsWith("Labels/"))
          throw new SafeError(
            `Only Labels/* mailboxes can be deleted as labels (${labelId})`,
          );
        await client.mailboxDelete(labelId);
        return;
      }
      await assertDeletableFolder(client, labelId);
      await client.mailboxDelete(labelId);
    });
  }

  async labelMessage(options: {
    messageId: string;
    labelId: string;
    labelName: string | null;
  }): Promise<{ usedFallback?: boolean; actualLabelId?: string }> {
    return this.withConnection(async (client) => {
      const { folder, uid } = parseImapMessageId(options.messageId);
      if (await isProtonBridgeLayout(client)) {
        const target =
          options.labelId.startsWith("Labels/") && options.labelId !== "Labels/"
            ? options.labelId
            : labelPathFromName(options.labelName || options.labelId);
        const existing = await listLabelMailboxes(client);
        if (!existing.some((p) => p.toLowerCase() === target.toLowerCase())) {
          await client.mailboxCreate(target);
        }
        if (folder === target) return { actualLabelId: target };
        await client.mailboxOpen(folder);
        await client.messageCopy(String(uid), target, { uid: true });
        return { actualLabelId: target };
      }
      // No label namespace: "labeling" means moving to a folder.
      if (folder === options.labelId) return { actualLabelId: options.labelId };
      await client.mailboxOpen(folder);
      await moveMessageToFolder(client, uid, options.labelId, {
        sourceFolder: folder,
      });
      return { actualLabelId: options.labelId };
    });
  }

  async removeThreadLabel(threadId: string, labelId: string): Promise<void> {
    return this.withConnection(async (client) => {
      if (!(await isProtonBridgeLayout(client))) {
        // No label namespace: a message lives in one folder, so there is
        // nothing to detach without moving it.
        this.logger.warn("removeThreadLabel is a no-op for IMAP", {
          threadId,
        });
        return;
      }
      await removeProtonThreadLabel(client, threadId, labelId);
    });
  }

  async removeThreadLabels(
    threadId: string,
    labelIds: string[],
  ): Promise<void> {
    return this.withConnection(async (client) => {
      if (!(await isProtonBridgeLayout(client))) {
        this.logger.warn("removeThreadLabels is a no-op for IMAP", {
          threadId,
        });
        return;
      }
      for (const labelId of labelIds) {
        await removeProtonThreadLabel(client, threadId, labelId);
      }
    });
  }

  // --- Thread Actions ---

  async archiveThread(threadId: string, _ownerEmail: string): Promise<void> {
    let messages: ParsedMessage[];
    try {
      messages = await this.getThreadMessages(threadId, { forMutation: true });
    } catch (error) {
      // Gmail treats archiving a vanished thread as success (404 → no-op);
      // IMAP already throws ImapThreadNotFoundError from getThreadMessages,
      // so match only that instead of failing automations that race a
      // concurrent move (and never swallow an unrelated "not found").
      if (error instanceof ImapThreadNotFoundError) {
        this.logger.warn("IMAP archiveThread on missing thread", { threadId });
        return;
      }
      throw error;
    }

    return this.withConnection(async (client) => {
      const archiveFolder = await findArchiveFolder(client);
      const targets = (await isProtonBridgeLayout(client))
        ? selectOnePlacementPerItem(messages)
        : messages;
      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        // Gmail semantics: archiving removes the INBOX placement only. A
        // Sent copy is not an inbox copy, so it stays in Sent; on Bridge the
        // move relocates the whole shared item (Sent placement hidden until
        // unarchive, live-verified).
        if (folder !== "INBOX") continue;
        if (folder === archiveFolder) continue;
        try {
          await movePlacementToFolder(client, msg, archiveFolder);
        } catch (error) {
          if (isAlreadyInTargetError(error)) continue;
          this.logger.warn("IMAP archiveThread move failed", {
            threadId,
            messageId: msg.id,
            error,
          });
          throw error;
        }
      }
    });
  }

  async archiveThreadWithLabel(
    threadId: string,
    ownerEmail: string,
    _labelId?: string,
  ): Promise<void> {
    // For IMAP, archive with label = move to the label folder (or archive)
    if (_labelId) {
      const messages = await this.getThreadMessages(threadId, {
        forMutation: true,
      });
      return this.withConnection(async (client) => {
        const onBridge = await isProtonBridgeLayout(client);
        const archiveFolder = onBridge ? await findArchiveFolder(client) : null;
        const isLabelMove =
          onBridge &&
          archiveFolder &&
          (_labelId === "Labels" || _labelId.startsWith("Labels/"));
        const targets = onBridge
          ? selectOnePlacementPerItem(messages)
          : messages;
        for (const msg of targets) {
          const { folder, uid } = parseImapMessageId(msg.id);
          if (isLabelMove) {
            // A Proton label is a view, not a location: COPY into it, then
            // archive from the real folder — never MOVE into Labels/*.
            if (folder !== _labelId) {
              await client.mailboxOpen(folder);
              await client.messageCopy(String(uid), _labelId, { uid: true });
            }
            if (folder === archiveFolder || !archiveFolder) continue;
            await movePlacementToFolder(client, msg, archiveFolder);
            continue;
          }
          if (folder === _labelId) continue;
          await movePlacementToFolder(client, msg, _labelId);
        }
      });
    }
    return this.archiveThread(threadId, ownerEmail);
  }

  async archiveMessage(messageId: string): Promise<void> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      const onBridge = await isProtonBridgeLayout(client);
      if (onBridge && isFallbackViewFolder(folder))
        throw new SafeError(
          `Cannot archive a message in a Proton view mailbox (${folder})`,
        );
      const archiveFolder = await findArchiveFolder(client);
      if (folder === archiveFolder) return;
      await client.mailboxOpen(folder);
      const live = await fetchMessageByUid(client, uid);
      if (!live) return;
      await movePlacementToFolder(client, live, archiveFolder);
    });
  }

  async trashThread(
    threadId: string,
    _ownerEmail: string,
    _actionSource: "user" | "automation",
  ): Promise<void> {
    let messages: ParsedMessage[];
    try {
      messages = await this.getThreadMessages(threadId, { forMutation: true });
    } catch (error) {
      // Same no-op-on-missing rule as archiveThread: Gmail trash returns 200
      // for an already-deleted thread, so a raced automation must not fail.
      // Match only the thread-not-found error, never any "/not found/".
      if (error instanceof ImapThreadNotFoundError) {
        this.logger.warn("IMAP trashThread on missing thread", { threadId });
        return;
      }
      throw error;
    }

    return this.withConnection(async (client) => {
      const trashFolder = await findTrashFolder(client);
      if (!trashFolder) return;
      const targets = (await isProtonBridgeLayout(client))
        ? selectOnePlacementPerItem(messages)
        : messages;
      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        if (folder === trashFolder) continue;
        try {
          await movePlacementToFolder(client, msg, trashFolder);
        } catch (error) {
          if (isAlreadyInTargetError(error)) continue;
          this.logger.warn("IMAP trashThread move failed", {
            threadId,
            messageId: msg.id,
            error,
          });
          throw error;
        }
      }
    });
  }

  async markReadThread(threadId: string, read: boolean): Promise<void> {
    const messages = await this.getThreadMessages(threadId, {
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        await client.mailboxOpen(folder);
        if (read) {
          await client.messageFlagsAdd(String(uid), ["\\Seen"], {
            uid: true,
          });
        } else {
          await client.messageFlagsRemove(String(uid), ["\\Seen"], {
            uid: true,
          });
        }
      }
    });
  }

  async markRead(threadId: string): Promise<void> {
    return this.markReadThread(threadId, true);
  }

  async markSpam(threadId: string): Promise<void> {
    // Move to Junk/Spam folder. Only received mail is spam-reportable: the
    // account's own Sent and Draft copies stay where they are.
    const messages = await this.getThreadMessages(threadId, {
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      const junkFolder = await findJunkFolder(client);
      const targetFolder = junkFolder || "Junk";
      const sentFolder = await findSentFolder(client);
      const draftsFolder = await findDraftsFolder(client);
      const skipped = new Set(
        [sentFolder, draftsFolder].filter((folder) => folder !== null),
      );
      const targets = (await isProtonBridgeLayout(client))
        ? selectOnePlacementPerItem(messages)
        : messages;

      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        if (folder === targetFolder || skipped.has(folder)) continue;
        await movePlacementToFolder(client, msg, targetFolder);
      }
    });
  }
  async blockUnsubscribedEmail(_messageId: string): Promise<void> {
    throw new UnsupportedImapOperationError("blockUnsubscribedEmail");
  }

  async moveThreadToFolder(
    threadId: string,
    _ownerEmail: string,
    folderName: string,
  ): Promise<void> {
    let messages: ParsedMessage[];
    try {
      messages = await this.getThreadMessages(threadId, {
        forMutation: true,
      });
    } catch (error) {
      // Same no-op-on-missing rule as trashThread: a second move finds
      // nothing to move, so it succeeds instead of throwing.
      if (error instanceof ImapThreadNotFoundError) {
        this.logger.warn("IMAP moveThreadToFolder on missing thread", {
          threadId,
        });
        return;
      }
      throw error;
    }
    return this.withConnection(async (client) => {
      const onBridge = await isProtonBridgeLayout(client);
      if (onBridge && isFallbackViewFolder(folderName))
        throw new SafeError(
          `Cannot move a thread into a Proton view mailbox (${folderName})`,
        );
      const targets = onBridge ? selectOnePlacementPerItem(messages) : messages;
      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        if (folder === folderName) continue;
        await movePlacementToFolder(client, msg, folderName);
      }
    });
  }

  // --- Bulk Operations ---

  async bulkArchiveFromSenders(
    fromEmails: string[],
    _ownerEmail: string,
    _emailAccountId: string,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      await client.mailboxOpen("INBOX");
      const archiveFolder = await findArchiveFolder(client);

      for (const email of fromEmails) {
        const candidates = await searchImapMessages(client, { from: email });
        const uids = await filterUidsByExactAddress(
          client,
          candidates,
          "from",
          email,
        );
        for (const uid of uids) {
          await moveMessageToFolder(client, uid, archiveFolder, {
            sourceFolder: "INBOX",
          });
        }
      }
    });
  }

  async bulkTrashFromSenders(
    fromEmails: string[],
    _ownerEmail: string,
    _emailAccountId: string,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      await client.mailboxOpen("INBOX");
      const trashFolder = await findTrashFolder(client);
      if (!trashFolder) return;

      for (const email of fromEmails) {
        const candidates = await searchImapMessages(client, { from: email });
        const uids = await filterUidsByExactAddress(
          client,
          candidates,
          "from",
          email,
        );
        for (const uid of uids) {
          await moveMessageToFolder(client, uid, trashFolder, {
            sourceFolder: "INBOX",
          });
        }
      }
    });
  }

  // --- Sent Messages ---

  async getSentMessages(maxResults?: number): Promise<ParsedMessage[]> {
    return this.withConnection(async (client) => {
      const sentFolder = await findSentFolder(client);
      if (!sentFolder) return [];
      await client.mailboxOpen(sentFolder, { readOnly: true });
      // UIDs are append order, not date order: rank the full match set by
      // INTERNALDATE desc (UID desc tie-break), then fetch the newest page.
      const uids = await searchImapMessages(client, { all: true });
      const dated = await fetchInternalDatesByUids(client, uids);
      dated.sort((a, b) => b.time - a.time || b.uid - a.uid);
      const pageUids = dated
        .slice(0, maxResults || 20)
        .map((entry) => entry.uid);
      const messages = await fetchMessagesByUids(client, pageUids);
      const order = new Map(pageUids.map((uid, index) => [uid, index]));
      messages.sort(
        (a, b) =>
          (order.get(parseImapMessageId(a.id).uid) ?? 0) -
          (order.get(parseImapMessageId(b.id).uid) ?? 0),
      );
      return messages;
    });
  }

  async getSentMessageIds(options: {
    maxResults: number;
    after?: Date;
    before?: Date;
    pageToken?: string;
  }): Promise<SentMessagePage> {
    return this.withConnection(async (client) => {
      const sentFolder = await findSentFolder(client);
      if (!sentFolder) return { messages: [] };
      await client.mailboxOpen(sentFolder, { readOnly: true });

      const criteria: Record<string, unknown> = { all: true };
      if (options.after) criteria.since = options.after;
      if (options.before) criteria.before = options.before;

      const offset = options.pageToken ? Number(options.pageToken) : 0;
      const allUids = await searchImapMessages(client, criteria);
      // UID order is append order, not date order: rank by INTERNALDATE desc
      // (UID desc tie-break) and offset into that ranking, like the listing.
      const dated = await fetchInternalDatesByUids(client, allUids);
      dated.sort((a, b) => b.time - a.time || b.uid - a.uid);
      const pageUids = dated
        .slice(offset, offset + options.maxResults)
        .map((entry) => entry.uid);
      const messages = await fetchMessagesByUids(client, pageUids);
      const order = new Map(pageUids.map((uid, index) => [uid, index]));
      messages.sort(
        (a, b) =>
          (order.get(parseImapMessageId(a.id).uid) ?? 0) -
          (order.get(parseImapMessageId(b.id).uid) ?? 0),
      );
      const nextOffset = offset + options.maxResults;
      return {
        messages: messages.map((m) => ({ id: m.id, threadId: m.threadId })),
        nextPageToken:
          nextOffset < dated.length ? String(nextOffset) : undefined,
      };
    });
  }

  async getSentThreadsExcluding(options: {
    excludeToEmails?: string[];
    excludeFromEmails?: string[];
    maxResults?: number;
  }): Promise<EmailThread[]> {
    const messages = await this.getSentMessages(options.maxResults || 20);

    const filtered = messages.filter((m) => {
      if (
        options.excludeToEmails?.some((e) =>
          splitRecipientList(m.headers.to).some(
            (entry) =>
              extractEmailAddress(entry).toLowerCase() === e.toLowerCase(),
          ),
        )
      )
        return false;
      if (
        options.excludeFromEmails?.some(
          (e) =>
            extractEmailAddress(m.headers.from).toLowerCase() ===
            e.toLowerCase(),
        )
      )
        return false;
      return true;
    });

    const threadMap = new Map<string, ParsedMessage[]>();
    for (const msg of filtered) {
      const existing = threadMap.get(msg.threadId) || [];
      existing.push(msg);
      threadMap.set(msg.threadId, existing);
    }

    return [...threadMap.entries()].map(([id, msgs]) => ({
      id,
      messages: msgs,
      snippet: msgs[msgs.length - 1].snippet,
    }));
  }

  // --- Email Sending ---

  async sendEmail(args: {
    to: string;
    cc?: string;
    bcc?: string;
    subject: string;
    messageText: string;
    attachments?: MailAttachment[];
  }): Promise<{ messageId: string }> {
    return sendSmtpEmail(this.config, {
      to: args.to,
      cc: args.cc,
      bcc: args.bcc,
      subject: args.subject,
      text: args.messageText,
      attachments: args.attachments,
    });
  }

  async sendEmailWithHtml(body: {
    replyToEmail?: {
      threadId: string;
      headerMessageId: string;
      references?: string;
      messageId?: string;
    };
    to: string;
    from?: string;
    cc?: string;
    bcc?: string;
    replyTo?: string;
    subject: string;
    messageHtml: string;
    attachments?: Array<{
      filename: string;
      content: string;
      contentType: string;
    }>;
  }): Promise<{ messageId: string; threadId: string }> {
    const result = await sendSmtpEmail(this.config, {
      to: body.to,
      from: body.from,
      cc: body.cc,
      bcc: body.bcc,
      subject: body.subject,
      html: body.messageHtml,
      replyTo: body.replyTo,
      inReplyTo: body.replyToEmail?.headerMessageId,
      references: body.replyToEmail?.references,
      // Base64 content strings need `encoding: "base64"` or nodemailer sends
      // them as UTF-8 text; the Gmail/Microsoft providers use the same helper.
      attachments: toMailerAttachments(body.attachments),
    });

    const threadId =
      body.replyToEmail?.threadId ||
      buildThreadId(
        body.replyToEmail?.references,
        body.replyToEmail?.headerMessageId,
        result.messageId,
      );

    return { messageId: result.messageId, threadId };
  }

  async replyToEmail(
    email: ParsedMessage,
    content: string,
    options?: {
      replyTo?: string;
      from?: string;
      cc?: string;
      bcc?: string;
      attachments?: MailAttachment[];
    },
  ): Promise<{ messageId: string }> {
    const replySubject = email.subject.startsWith("Re:")
      ? email.subject
      : `Re: ${email.subject}`;

    return sendSmtpEmail(this.config, {
      to: email.headers.from,
      from: options?.from,
      cc: options?.cc,
      bcc: options?.bcc,
      subject: replySubject,
      html: content,
      replyTo: options?.replyTo,
      inReplyTo: email.headers["message-id"],
      references: email.headers.references
        ? `${email.headers.references} ${email.headers["message-id"]}`
        : email.headers["message-id"],
      attachments: options?.attachments,
    });
  }

  async forwardEmail(
    email: ParsedMessage,
    args: { to: string; cc?: string; bcc?: string; content?: string },
  ): Promise<{ messageId: string }> {
    const fwdSubject = email.subject.startsWith("Fwd:")
      ? email.subject
      : `Fwd: ${email.subject}`;

    const messageContent =
      email.textHtml || textToHtmlParagraphs(email.textPlain);

    const forwardedBody = `
<div dir="ltr">${escapeHtml(args.content || "")}<br><br>
---------- Forwarded message ----------<br/>
From: ${escapeHtml(email.headers.from)}<br/>
Date: ${escapeHtml(email.headers.date)}<br/>
Subject: ${escapeHtml(email.headers.subject)}<br/>
To: ${escapeHtml(email.headers.to)}<br/>
<br/>
${messageContent}
</div>`.trim();

    return sendSmtpEmail(this.config, {
      to: args.to,
      cc: args.cc,
      bcc: args.bcc,
      subject: fwdSubject,
      html: forwardedBody,
    });
  }

  // --- Drafts ---

  // Draft listings carry the stable draft id (not the folder-qualified
  // message id) so stored draft ids match listed ones.
  async getDrafts(options?: { maxResults?: number }): Promise<ParsedMessage[]> {
    return this.withConnection(async (client) => {
      const draftsFolder = await findDraftsFolder(client);
      if (!draftsFolder) return [];
      await client.mailboxOpen(draftsFolder, { readOnly: true });
      // searchImapMessages caps newest UIDs, and UID order is append order:
      // rank the full match set by INTERNALDATE desc, then fetch the page.
      const uids = await searchImapMessages(client, { all: true });
      const dated = await fetchInternalDatesByUids(client, uids);
      dated.sort((a, b) => b.time - a.time || b.uid - a.uid);
      const pageUids = dated
        .slice(0, options?.maxResults || 20)
        .map((entry) => entry.uid);
      const drafts = await fetchMessagesByUids(client, pageUids);
      return drafts.map((draft) => {
        const headerId = draft.headers["message-id"];
        if (!headerId) return draft;
        try {
          return { ...draft, id: encodeImapDraftId(headerId) };
        } catch {
          return draft;
        }
      });
    });
  }

  async getDraft(draftId: string): Promise<ParsedMessage | null> {
    return this.withConnection(async (client) => {
      const resolved = await resolveDraftUid(client, draftId, {
        readOnly: true,
      });
      return resolved?.message ?? null;
    });
  }

  async createDraft(params: {
    to: string;
    from?: string;
    cc?: string;
    bcc?: string;
    subject: string;
    messageHtml: string;
    replyToMessageId?: string;
    forwardedMessageId?: string;
  }): Promise<{ id: string }> {
    return this.withConnection(async (client) => {
      // Thread a reply on the parent's headers: its Message-ID becomes
      // In-Reply-To and extends References, like replyToEmail does when
      // sending directly. Proton-internal ids never anchor a thread.
      let inReplyTo = params.forwardedMessageId;
      let references: string | undefined;
      if (params.replyToMessageId) {
        const parent = await this.getMessage(params.replyToMessageId).catch(
          () => null,
        );
        const parentId = parent?.headers["message-id"];
        if (parentId && !isProtonInternalId(parentId)) {
          const threading = buildThreadingHeaders({
            headerMessageId: parentId,
            references: parent?.headers.references,
          });
          inReplyTo = threading.inReplyTo;
          references = threading.references || undefined;
        } else {
          inReplyTo = params.replyToMessageId;
        }
      }
      const id = await saveDraft(client, {
        from: params.from || this.config.email,
        to: params.to,
        cc: params.cc,
        bcc: params.bcc,
        subject: params.subject,
        html: params.messageHtml,
        inReplyTo,
        references,
      });
      return { id };
    });
  }

  async draftEmail(
    email: ParsedMessage,
    args: {
      to?: string;
      subject?: string;
      content: string;
      cc?: string;
      bcc?: string;
      attachments?: MailAttachment[];
    },
    userEmail: string,
    _executedRule?: { id: string; threadId: string; emailAccountId: string },
  ): Promise<{ draftId: string }> {
    return this.withConnection(async (client) => {
      const draftId = await saveDraft(client, {
        from: userEmail,
        to: args.to || email.headers.from,
        subject:
          args.subject ||
          (email.subject.startsWith("Re:")
            ? email.subject
            : `Re: ${email.subject}`),
        html: args.content,
        inReplyTo: email.headers["message-id"],
        references: email.headers.references,
      });
      return { draftId };
    });
  }

  // IMAP can't edit a message: APPEND the updated copy with the same
  // Message-ID (so the stable draft id survives), then expunge the old UID.
  // Append-first keeps a failure from losing the draft. The rebuild carries
  // the draft's files, so a text save never drops them.
  async updateDraft(
    draftId: string,
    params: {
      messageHtml?: string;
      subject?: string;
      to?: string;
      from?: string;
      cc?: string;
      bcc?: string;
    },
  ): Promise<void> {
    await this.withConnection(async (client) => {
      const resolved = await resolveDraftUid(client, draftId);
      if (!resolved) throw draftNotFoundError("update");
      const headers = resolved.message.headers;
      const messageId = headers["message-id"];
      // Without the current Message-ID there is no id to keep stable: a
      // fresh id would orphan the caller's draft.
      if (!messageId) throw draftNotFoundError("update");
      const parts = await this.readDraftParts(client, resolved);
      await appendDraftMessage(client, resolved.folder, {
        from: params.from ?? headers.from,
        to: params.to ?? headers.to ?? "",
        cc: params.cc ?? headers.cc,
        bcc: params.bcc ?? headers.bcc,
        subject: params.subject ?? resolved.message.subject,
        html: params.messageHtml ?? resolved.message.textHtml ?? "",
        inReplyTo: headers["in-reply-to"],
        references: headers.references,
        messageId,
        ...(parts.length
          ? { attachments: parts.map((part) => toDraftMailAttachment(part)) }
          : {}),
      });
      await client.messageDelete(String(resolved.uid), { uid: true });
    });
  }
  async deleteDraft(draftId: string, _version?: string): Promise<boolean> {
    return this.withConnection(async (client) => {
      const resolved = await resolveDraftUid(client, draftId).catch(
        (error: unknown) => {
          if (error instanceof DraftNotFoundError) return null;
          throw error;
        },
      );
      if (!resolved) return false;
      await client.messageDelete(String(resolved.uid), { uid: true });
      return true;
    });
  }

  // Send the draft's current content via SMTP, then remove it from Drafts.
  // Bridge files sent mail into Sent itself; the provider never APPENDs.
  async sendDraft(
    draftId: string,
  ): Promise<{ messageId: string; threadId: string }> {
    return this.withConnection(async (client) => {
      const resolved = await resolveDraftUid(client, draftId);
      if (!resolved) throw draftNotFoundError("send");
      const headers = resolved.message.headers;
      const parts = await this.readDraftParts(client, resolved);
      const { messageId } = await sendSmtpEmail(this.config, {
        to: headers.to ?? "",
        from: headers.from || undefined,
        cc: headers.cc,
        bcc: headers.bcc,
        subject: resolved.message.subject,
        html: resolved.message.textHtml ?? "",
        inReplyTo: headers["in-reply-to"],
        references: headers.references,
        ...(parts.length
          ? { attachments: parts.map((part) => toDraftMailAttachment(part)) }
          : {}),
      });
      await client.messageDelete(String(resolved.uid), { uid: true });
      return {
        messageId,
        threadId: buildThreadId(
          headers.references,
          headers["in-reply-to"],
          headers["message-id"],
        ),
      };
    });
  }

  // --- Folders ---

  async getFolders(): Promise<OutlookFolder[]> {
    return this.withConnection(async (client) => {
      const folders = await listFoldersAsOutlookFolders(client);
      // Views are not move targets: on the Bridge layout All Mail, Starred,
      // Labels/* and the Folders/Labels containers apply a label or show a
      // view — a MOVE there strands or duplicates the message.
      if (!(await isProtonBridgeLayout(client))) return folders;
      return folders.filter(
        (folder) =>
          !isFallbackViewFolder(folder.id) &&
          folder.id !== "Folders" &&
          folder.id !== "Labels",
      );
    });
  }

  async getOrCreateFolderIdByName(folderName: string): Promise<string> {
    return this.withConnection(async (client) =>
      getOrCreateFolder(client, folderName),
    );
  }

  // --- Filters (not supported) ---

  async getFiltersList(): Promise<EmailFilter[]> {
    return [];
  }

  async createFilter(_options: {
    from: string;
    addLabelIds?: string[];
    removeLabelIds?: string[];
  }): Promise<{ status: number }> {
    throw new UnsupportedImapOperationError("createFilter");
  }

  async createAutoArchiveFilter(_options: {
    from: string;
    gmailLabelId?: string;
    labelName?: string;
  }): Promise<{ status: number }> {
    throw new UnsupportedImapOperationError("createAutoArchiveFilter");
  }

  async deleteFilter(_id: string): Promise<{ status: number }> {
    throw new UnsupportedImapOperationError("deleteFilter");
  }

  // --- Utility Methods ---

  async getSignatures(): Promise<EmailSignature[]> {
    return [];
  }

  async getAttachment(
    messageId: string,
    attachmentId: string,
  ): Promise<{ data: string; size: number }> {
    const bytes = await this.downloadAttachmentPart(messageId, attachmentId);
    return { data: bytes.toString("base64"), size: bytes.length };
  }

  async getOriginalMessage(
    originalMessageId: string | undefined,
  ): Promise<ParsedMessage | null> {
    if (!originalMessageId) return null;
    try {
      return await this.getMessage(originalMessageId);
    } catch {
      return null;
    }
  }

  async checkIfReplySent(senderEmail: string): Promise<boolean> {
    return this.withConnection(async (client) => {
      const sentFolder = await findSentFolder(client);
      if (!sentFolder) return false;
      await client.mailboxOpen(sentFolder, { readOnly: true });
      // Exact-match, newest first in chunks: substring SEARCH can hide the
      // exact reply behind many newer substring hits, and fetching every
      // candidate at once is wasteful when the reply is recent.
      const candidates = await searchImapMessages(client, {
        to: senderEmail,
      });
      for (let i = 0; i < candidates.length; i += 20) {
        const uids = await filterUidsByExactAddress(
          client,
          candidates.slice(i, i + 20),
          "to",
          senderEmail,
        );
        if (uids.length > 0) return true;
      }
      return false;
    });
  }

  async countReceivedMessages(
    senderEmail: string,
    threshold: number,
  ): Promise<number> {
    return this.withConnection(async (client) => {
      await client.mailboxOpen("INBOX", { readOnly: true });
      const candidates = await searchImapMessages(client, {
        from: senderEmail,
      });
      const uids = await filterUidsByExactAddress(
        client,
        candidates,
        "from",
        senderEmail,
      );
      // Cap after the exact filter: capping first lets newer substring hits
      // (joann@x.com for ann@x.com) fill the window and hide exact matches.
      return Math.min(uids.length, threshold);
    });
  }

  async hasPreviousCommunicationsWithSenderOrDomain(options: {
    from: string;
    date: Date;
    messageId: string;
  }): Promise<boolean> {
    return this.withConnection(async (client) => {
      const sentFolder = await findSentFolder(client);
      if (!sentFolder) return false;
      await client.mailboxOpen(sentFolder, { readOnly: true });

      // Check if we've sent to this address before
      const email =
        extractEmailAddress(options.from) ||
        options.from.match(/<([^>]+)>/)?.[1] ||
        options.from;
      const candidates = await searchImapMessages(client, {
        to: email,
        before: options.date,
      });
      const uids = await filterUidsByExactAddress(
        client,
        candidates,
        "to",
        email,
      );
      return uids.length > 0;
    });
  }

  isSentMessage(message: ParsedMessage): boolean {
    const from = extractEmailAddress(message.headers.from).toLowerCase();
    return from.length > 0 && from === this.config.email.toLowerCase();
  }

  isReplyInThread(message: ParsedMessage): boolean {
    return !!(message.headers["in-reply-to"] || message.headers.references);
  }

  async archiveMessages(
    messageIds: string[],
    _labelId?: string,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      const archiveFolder = await findArchiveFolder(client);
      const onBridge = await isProtonBridgeLayout(client);
      await moveMessageSnapshots(client, messageIds, archiveFolder, {
        onBridge,
        logSkippedView: (id) =>
          this.logger.warn("IMAP archiveMessages skipping Bridge view id", {
            messageId: id,
          }),
        onError: (id, error) =>
          this.logger.warn("IMAP archiveMessages move failed", {
            messageId: id,
            error,
          }),
      });
    });
  }

  async unarchiveMessages(messageIds: string[]): Promise<void> {
    return this.withConnection(async (client) => {
      const onBridge = await isProtonBridgeLayout(client);
      await moveMessageSnapshots(client, messageIds, "INBOX", {
        onBridge,
        logSkippedView: (id) =>
          this.logger.warn("IMAP unarchiveMessages skipping Bridge view id", {
            messageId: id,
          }),
      });
    });
  }

  async trashMessages(messageIds: string[]): Promise<void> {
    return this.withConnection(async (client) => {
      const trashFolder = await findTrashFolder(client);
      if (!trashFolder) return;
      const onBridge = await isProtonBridgeLayout(client);
      await moveMessageSnapshots(client, messageIds, trashFolder, {
        onBridge,
        logSkippedView: (id) =>
          this.logger.warn("IMAP trashMessages skipping Bridge view id", {
            messageId: id,
          }),
        onError: (id, error) =>
          this.logger.warn("IMAP trashMessages move failed", {
            messageId: id,
            error,
          }),
      });
    });
  }

  async untrashMessages(messageIds: string[]): Promise<void> {
    return this.withConnection(async (client) => {
      const onBridge = await isProtonBridgeLayout(client);
      await moveMessageSnapshots(client, messageIds, "INBOX", {
        onBridge,
        logSkippedView: (id) =>
          this.logger.warn("IMAP untrashMessages skipping Bridge view id", {
            messageId: id,
          }),
      });
    });
  }

  async markMessagesReadState(
    messageIds: string[],
    read: boolean,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      const onBridge = await isProtonBridgeLayout(client);
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        if (onBridge && isFallbackViewFolder(folder)) {
          this.logger.warn("IMAP markMessagesReadState skipping view id", {
            messageId,
          });
          continue;
        }
        await client.mailboxOpen(folder);
        if (read) {
          await client.messageFlagsAdd(String(uid), ["\\Seen"], {
            uid: true,
          });
        } else {
          await client.messageFlagsRemove(String(uid), ["\\Seen"], {
            uid: true,
          });
        }
      }
    });
  }

  async markMessagesStarredState(
    messageIds: string[],
    starred: boolean,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      const onBridge = await isProtonBridgeLayout(client);
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        if (onBridge && isFallbackViewFolder(folder)) {
          this.logger.warn("IMAP markMessagesStarredState skipping view id", {
            messageId,
          });
          continue;
        }
        await client.mailboxOpen(folder);
        if (starred) {
          await client.messageFlagsAdd(String(uid), ["\\Flagged"], {
            uid: true,
          });
        } else {
          await client.messageFlagsRemove(String(uid), ["\\Flagged"], {
            uid: true,
          });
        }
      }
    });
  }

  async starMessage(messageId: string): Promise<void> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      if ((await isProtonBridgeLayout(client)) && isFallbackViewFolder(folder))
        throw new SafeError(
          `Cannot star a message in a Proton view mailbox (${folder})`,
        );
      await client.mailboxOpen(folder);
      await client.messageFlagsAdd(String(uid), ["\\Flagged"], {
        uid: true,
      });
    });
  }

  async markNotSpam(threadId: string): Promise<void> {
    const messages = await this.getThreadMessages(threadId, {
      includeTrashAndJunk: true,
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      const junkFolder = await findJunkFolder(client);
      const spamFolder = junkFolder || "Junk";
      const targets = (await isProtonBridgeLayout(client))
        ? selectOnePlacementPerItem(messages, spamFolder)
        : messages;
      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        // Not-spam returns spam mail to INBOX; copies elsewhere keep theirs.
        if (folder !== spamFolder) continue;
        await movePlacementToFolder(client, msg, "INBOX");
      }
    });
  }

  async unarchiveThread(threadId: string): Promise<void> {
    const messages = await this.getThreadMessages(threadId, {
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      const archiveFolder = await findArchiveFolder(client);
      const targets = (await isProtonBridgeLayout(client))
        ? selectOnePlacementPerItem(messages, archiveFolder)
        : messages;
      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        // Unarchiving returns archived mail to INBOX; anything elsewhere
        // (notably Sent) keeps its placement.
        if (folder !== archiveFolder) continue;
        await movePlacementToFolder(client, msg, "INBOX");
      }
    });
  }

  async untrashThread(threadId: string): Promise<void> {
    // A thread whose only copy is in Trash is invisible to unarchiveThread,
    // which now scans the same folders as every other thread operation.
    const messages = await this.getThreadMessages(threadId, {
      includeTrashAndJunk: true,
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      const trashFolder = await findTrashFolder(client);
      if (!trashFolder) return;
      const sentFolder = await findSentFolder(client);
      const ownAddress = this.config.email.toLowerCase();
      const targets = (await isProtonBridgeLayout(client))
        ? selectOnePlacementPerItem(messages, trashFolder)
        : messages;
      for (const msg of targets) {
        const { folder } = parseImapMessageId(msg.id);
        // Only trashed copies move back. A copy the account sent restores to
        // Sent; received mail restores to INBOX.
        if (folder !== trashFolder) continue;
        const from = extractEmailAddress(msg.headers.from).toLowerCase();
        const target =
          sentFolder && from !== "" && from === ownAddress
            ? sentFolder
            : "INBOX";
        await movePlacementToFolder(client, msg, target);
      }
    });
  }

  async bulkArchiveSenderOrThrow(
    fromEmail: string,
    _ownerEmail: string,
    _emailAccountId: string,
  ): Promise<number> {
    return this.withConnection(async (client) => {
      await client.mailboxOpen("INBOX");
      const archiveFolder = await findArchiveFolder(client);
      const candidates = await searchImapMessages(client, { from: fromEmail });
      const uids = await filterUidsByExactAddress(
        client,
        candidates,
        "from",
        fromEmail,
      );
      for (const uid of uids) {
        await moveMessageToFolder(client, uid, archiveFolder, {
          sourceFolder: "INBOX",
        });
      }
      return uids.length;
    });
  }

  async bulkArchiveThreads(
    threads: BulkArchiveThread[],
    _ownerEmail: string,
  ): Promise<BulkArchiveResult> {
    const succeededThreadIds: string[] = [];
    const failedThreadIds: string[] = [];
    await this.withConnection(async (client) => {
      const archiveFolder = await findArchiveFolder(client);
      const onBridge = await isProtonBridgeLayout(client);
      for (const thread of threads) {
        try {
          await moveMessageSnapshots(client, thread.messageIds, archiveFolder, {
            onBridge,
            logSkippedView: () => {},
          });
          succeededThreadIds.push(thread.threadId);
        } catch {
          failedThreadIds.push(thread.threadId);
        }
      }
    });
    return { succeededThreadIds, failedThreadIds };
  }

  async countMessages(options: {
    folderId?: string;
    labelId?: string;
  }): Promise<number> {
    return this.withConnection(async (client) => {
      const folder = options.folderId || options.labelId || "INBOX";
      const mailbox = await client.mailboxOpen(folder, { readOnly: true });
      return mailbox.exists || 0;
    });
  }

  async searchThreads(options: {
    query: string;
    maxResults?: number;
    pageToken?: string;
    messageFormat?: "full" | "metadata";
    includeSpamTrash?: boolean;
    folder?: "spam" | "trash";
    labelIds?: string[];
  }): Promise<{ threads: EmailThread[]; nextPageToken?: string }> {
    if (options.folder) {
      const { messages, nextPageToken } = await this.withConnection(
        async (client) => {
          const folder =
            options.folder === "spam"
              ? await findJunkFolder(client)
              : await findTrashFolder(client);
          if (!folder) return { messages: [], nextPageToken: undefined };
          return searchImapFolder(client, {
            folder,
            search: {
              mailbox: options.folder as "spam" | "trash",
              ...(options.query
                ? {
                    text: {
                      kind: "text",
                      field: "any",
                      value: options.query,
                      match: "term",
                    } as const,
                  }
                : {}),
            },
            maxResults: options.maxResults,
            pageToken: options.pageToken,
          });
        },
      );
      return {
        threads: groupImapSearchThreads(messages),
        nextPageToken,
      };
    }
    const { messages, nextPageToken } = await this.getMessagesWithPagination({
      query: options.query,
      maxResults: options.maxResults,
      pageToken: options.pageToken,
      ...(options.includeSpamTrash
        ? { inboxOnly: false, includeTrashAndJunk: true }
        : {}),
    });
    return {
      threads: groupImapSearchThreads(messages),
      nextPageToken,
    };
  }

  async searchContacts(_query: string): Promise<EmailContact[]> {
    // IMAP has no contacts API
    return [];
  }

  async getContactPhotos(): Promise<Record<string, string>> {
    // IMAP has no contact photos API
    return {};
  }

  async getForwardingAddresses(): Promise<string[]> {
    // IMAP has no equivalent of Gmail's verified forwarding addresses list
    return [];
  }

  async getSendAsAddresses(): Promise<SendAsAddress[]> {
    // IMAP sends from the account address
    return [];
  }

  async getFolderCounts(): Promise<EmailFolderCount[]> {
    return this.withConnection(async (client) => {
      const folders = await listFoldersAsOutlookFolders(client);
      const counts: EmailFolderCount[] = [];
      for (const folder of folders) {
        try {
          const status = await client.status(folder.id, {
            messages: true,
            unseen: true,
          });
          counts.push({
            id: folder.id,
            name: folder.displayName,
            total: status.messages || 0,
            unread: status.unseen || 0,
            ...(folder.systemType ? { systemType: folder.systemType } : {}),
          });
        } catch {
          counts.push({
            id: folder.id,
            name: folder.displayName,
            total: 0,
            unread: 0,
            ...(folder.systemType ? { systemType: folder.systemType } : {}),
          });
        }
      }
      return counts;
    });
  }

  async renameFolder(folderId: string, name: string): Promise<void> {
    return this.withConnection(async (client) => {
      await client.mailboxRename(folderId, name);
    });
  }

  async deleteFolder(folderId: string): Promise<void> {
    return this.withConnection(async (client) => {
      await assertDeletableFolder(client, folderId);
      await client.mailboxDelete(folderId);
    });
  }

  async updateLabel(labelId: string, update: EmailLabelUpdate): Promise<void> {
    if (
      update.color ||
      update.labelListVisibility ||
      update.messageListVisibility
    )
      throw new Error("IMAP folders have no color or visibility settings");
    if (!update.name) throw new Error("IMAP folder rename requires a name");
    return this.renameFolder(labelId, update.name);
  }

  async getMessageSource(messageId: string): Promise<Buffer> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      if ((await isProtonBridgeLayout(client)) && isFallbackViewFolder(folder))
        throw new SafeError(
          `Cannot read a message in a Proton view mailbox (${folder})`,
        );
      await client.mailboxOpen(folder, { readOnly: true });
      const downloaded = await client.download(String(uid), undefined, {
        uid: true,
      });
      const chunks: Buffer[] = [];
      for await (const chunk of downloaded.content) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      }
      return Buffer.concat(chunks);
    });
  }

  async getAttachmentStream(
    messageId: string,
    attachmentId: string,
    signal?: AbortSignal,
  ): Promise<ReadableStream<Uint8Array>> {
    signal?.throwIfAborted();
    const bytes = await this.downloadAttachmentPart(messageId, attachmentId);
    signal?.throwIfAborted();
    return new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array(bytes));
        controller.close();
      },
    });
  }

  // Attachment ids are MIME part ids from bodyStructure. Open the mailbox
  // named by the folder-qualified message id, verify the part exists, then
  // download it — a legitimately empty part returns empty bytes, while an
  // absent part throws.
  private async downloadAttachmentPart(
    messageId: string,
    attachmentId: string,
  ): Promise<Buffer> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      if ((await isProtonBridgeLayout(client)) && isFallbackViewFolder(folder))
        throw new SafeError(
          `Cannot read a message in a Proton view mailbox (${folder})`,
        );
      await client.mailboxOpen(folder, { readOnly: true });
      const listed = await client.fetchOne(
        String(uid),
        { bodyStructure: true },
        { uid: true },
      );
      if (!findBodyPart(listed?.bodyStructure, attachmentId))
        throw new SafeError(
          `Attachment part ${attachmentId} not found on message ${messageId}`,
        );
      const downloaded = await client.download(String(uid), attachmentId, {
        uid: true,
      });
      const chunks: Buffer[] = [];
      for await (const chunk of downloaded.content) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      }
      return Buffer.concat(chunks);
    });
  }

  // The compose flow hands `getDraft`'s message id back in; resolve it to
  // the draft's stable id when the message still lives in Drafts.
  async getDraftReferenceForMessage(
    messageId: string,
  ): Promise<{ id: string; version?: string } | null> {
    return this.withConnection(async (client) => {
      const draftsFolder = await findDraftsFolder(client);
      if (!draftsFolder) return null;
      try {
        const { folder, uid } = parseImapMessageId(messageId);
        if (folder !== draftsFolder) return null;
        await client.mailboxOpen(draftsFolder, { readOnly: true });
        const message = await fetchMessageByUid(client, uid);
        const headerId = message?.headers["message-id"];
        if (!headerId) return null;
        const found = await findDraftByMessageId(
          client,
          headerId.replace(/^<|>$/g, ""),
        );
        return found
          ? {
              id: encodeImapDraftId(
                found.message.headers["message-id"] ?? headerId,
              ),
            }
          : null;
      } catch {
        return null;
      }
    });
  }

  async getDraftAttachments(
    draftId: string,
  ): Promise<DraftAttachmentsResult | null> {
    return this.withConnection(async (client) => {
      const resolved = await resolveDraftUid(client, draftId, {
        readOnly: true,
      });
      if (!resolved) return null;
      const parts = await this.readDraftParts(client, resolved);
      const messageId = encodeImapMessageId(resolved.folder, resolved.uid);
      return {
        messageId,
        attachments: parts.map((part) => ({
          id: part.id,
          filename: part.filename,
          mimeType: part.mimeType,
          size: part.size,
          disposition: part.disposition,
          ...(part.contentId ? { contentId: part.contentId } : {}),
          messageId,
          // The composer re-upload path reads bytes through the attachment
          // route, which needs a provider part id; IMAP has none per part,
          // so it re-reads the whole draft.
          providerAttachmentId: part.id,
        })),
      };
    });
  }

  // IMAP has no per-attachment draft API: like Gmail, every change rebuilds
  // the whole draft MIME (APPEND with the same Message-ID, then expunge the
  // old copy) so the stable draft id survives.
  async addDraftAttachment(
    draftId: string,
    attachment: DraftAttachmentMetadata & { content: Buffer },
  ): Promise<DraftAttachmentsResult & { attachmentId: string }> {
    return this.withConnection(async (client) => {
      const resolved = await this.requireDraftForAttachment(client, draftId);
      const parts = await this.readDraftParts(client, resolved);
      const kept = parts.filter((part) => part.id !== attachment.id);
      assertDraftAttachmentLimits(kept, attachment.size);
      await this.replaceDraftKeepingId(client, resolved, [
        ...kept.map((part) => toDraftMailAttachment(part)),
        toDraftMailAttachment({
          id: attachment.id,
          filename: attachment.filename,
          mimeType: attachment.mimeType,
          disposition: attachment.disposition,
          ...(attachment.contentId ? { contentId: attachment.contentId } : {}),
          content: attachment.content,
        }),
      ]);
      const result = await this.getDraftAttachments(draftId);
      if (!result) throw draftNotFoundError("attach to");
      return { ...result, attachmentId: attachment.id };
    });
  }

  async removeDraftAttachment(
    draftId: string,
    attachmentId: string,
  ): Promise<DraftAttachmentsResult> {
    return this.withConnection(async (client) => {
      const resolved = await this.requireDraftForAttachment(client, draftId);
      const parts = await this.readDraftParts(client, resolved);
      if (!parts.some((part) => part.id === attachmentId))
        throw new SafeError("This file is no longer on the draft.");
      await this.replaceDraftKeepingId(
        client,
        resolved,
        parts
          .filter((part) => part.id !== attachmentId)
          .map((part) => toDraftMailAttachment(part)),
      );
      const result = await this.getDraftAttachments(draftId);
      if (!result) throw draftNotFoundError("remove an attachment from");
      return result;
    });
  }

  async startDraftAttachmentUpload(
    draftId: string,
    attachment: DraftAttachmentMetadata,
  ): Promise<DraftAttachmentUploadStart> {
    return this.withConnection(async (client) => {
      const resolved = await this.requireDraftForAttachment(client, draftId);
      const parts = await this.readDraftParts(client, resolved);
      const existing = parts.filter((part) => part.id !== attachment.id);
      assertDraftAttachmentLimits(existing, attachment.size);
      if (
        [...existing, attachment].some((item) =>
          item.mimeType.startsWith("message/"),
        )
      )
        throw new SafeError(
          "Large files can't be added to a draft that has an attached email. Attach a file under 3 MB, or remove the attached email.",
        );
      const slots = [
        ...existing.map((part) => ({
          attachment: part,
          part: {
            type: "attachment" as const,
            attachmentId: part.id,
            size: part.size,
            source: {
              messageId: encodeImapMessageId(resolved.folder, resolved.uid),
              providerAttachmentId: part.id,
            },
          },
        })),
        {
          attachment,
          part: {
            type: "attachment" as const,
            attachmentId: attachment.id,
            size: attachment.size,
          },
        },
      ].map((slot) => {
        // A short placeholder whose base64 fits one line, so it can be found
        // in the built message and swapped for the real file.
        const content = Buffer.from(randomBytes(15).toString("hex"));
        return { ...slot, content, encoded: content.toString("base64") };
      });
      const template = await this.buildDraftUploadTemplate(
        resolved,
        slots.map((slot) =>
          toDraftMailAttachment({ ...slot.attachment, content: slot.content }),
        ),
      );
      return templateToUploadStart(template, slots);
    });
  }

  // Reads a draft's files from its raw MIME: APPEND keeps no attachment
  // metadata, so the source is the only store. The Drafts folder stays open
  // on the resolved folder throughout.
  private async readDraftParts(
    client: ImapFlow,
    resolved: ResolvedDraft,
  ): Promise<DraftAttachmentPart[]> {
    await client.mailboxOpen(resolved.folder, { readOnly: true });
    const downloaded = await client.download(String(resolved.uid), undefined, {
      uid: true,
    });
    const chunks: Buffer[] = [];
    for await (const chunk of downloaded.content) {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    return readDraftAttachmentParts(Buffer.concat(chunks));
  }

  private async requireDraftForAttachment(
    client: ImapFlow,
    draftId: string,
  ): Promise<ResolvedDraft> {
    const resolved = await resolveDraftUid(client, draftId);
    if (!resolved) throw draftNotFoundError("attach to");
    if (!resolved.message.headers["message-id"])
      throw draftNotFoundError("attach to");
    return resolved;
  }

  private async replaceDraftKeepingId(
    client: ImapFlow,
    resolved: ResolvedDraft,
    attachments: MailAttachment[],
  ): Promise<void> {
    const headers = resolved.message.headers;
    const messageId = headers["message-id"];
    if (!messageId) throw draftNotFoundError("update");
    await appendDraftMessage(client, resolved.folder, {
      from: headers.from,
      to: headers.to ?? "",
      cc: headers.cc,
      bcc: headers.bcc,
      subject: resolved.message.subject,
      html: resolved.message.textHtml ?? "",
      inReplyTo: headers["in-reply-to"],
      references: headers.references,
      messageId,
      ...(attachments.length ? { attachments } : {}),
    });
    await client.messageDelete(String(resolved.uid), { uid: true });
  }

  private async buildDraftUploadTemplate(
    resolved: ResolvedDraft,
    attachments: MailAttachment[],
  ): Promise<Buffer> {
    const headers = resolved.message.headers;
    const messageId = headers["message-id"];
    if (!messageId) throw draftNotFoundError("attach to");
    return buildDraftMime({
      from: headers.from,
      to: headers.to ?? "",
      cc: headers.cc,
      bcc: headers.bcc,
      subject: resolved.message.subject,
      html: resolved.message.textHtml ?? "",
      inReplyTo: headers["in-reply-to"],
      references: headers.references,
      messageId,
      attachments,
    });
  }

  // --- Local mail sync (UID-based folder-delta) ---

  readonly localMailSyncStrategy = "folder-delta" as const;

  async syncLocalMail(
    request: LocalMailSyncRequest,
    context: { emailAccountId: string },
  ): Promise<LocalMailSyncResponse> {
    if (!context.emailAccountId)
      throw new Error("Local mail account context is required");
    if (request.phase === "capabilities")
      return {
        status: "ok",
        phase: request.phase,
        result: {
          strategy: this.localMailSyncStrategy,
          excludedFolderIds: [],
          maxHydrationMessages: 1,
        },
      };
    if (request.phase === "folders")
      return {
        status: "ok",
        phase: request.phase,
        result: await getImapMailFoldersPage({
          config: this.config,
          logger: this.logger,
          emailAccountId: context.emailAccountId,
          limit: request.limit,
          parentFolderId: request.parentFolderId,
          cursor: request.cursor,
        }),
      };
    return { status: "unsupported", strategy: this.localMailSyncStrategy };
  }

  async getMailboxSyncPage(options: {
    after?: Date;
    cursor?: string;
    folderId?: string;
    limit: number;
  }): Promise<MailboxSyncPage> {
    return getImapMailboxSyncPage({
      config: this.config,
      logger: this.logger,
      emailAccountId: this.config.emailAccountId,
      ...options,
    });
  }

  // --- Watch (polling-based, no push) ---

  async watchEmails(): Promise<{
    expirationDate: Date;
    subscriptionId?: string;
  } | null> {
    // IMAP doesn't support push notifications via webhooks.
    // Polling is handled separately via /api/imap/poll
    return null;
  }

  async unwatchEmails(_subscriptionId?: string): Promise<void> {
    // No-op for IMAP
  }
}

/**
 * Mailboxes a whole-account IMAP search should cover.
 * A flagged `\All` mailbox (Gmail) is canonical: reads and mutations both
 * use it. On the Bridge layout (top-level `Folders`, `Labels`, `All Mail`,
 * no `\All` anywhere) "All Mail" is a view, so every path — reads and
 * mutations alike — fans out over real folders (INBOX, Archive, Sent,
 * Drafts, Spam, Trash, `Folders/*`), skipping views (`All Mail`,
 * `Labels/*`, `Starred`, `\Noselect`, the `Folders`/`Labels` containers)
 * never point into a view. A stray user folder named `All Mail` on another
 * server never narrows a search. Otherwise fan out over every selectable
 * real mailbox, skipping views whose messages also live in a real folder.
 * Mutation callers (archive, trash, restore, move, mark, star) pass
 * `forMutation: true`; Trash and Junk are excluded by default (listing and
 * search scope) but included when `includeTrashAndJunk` is set, so
 * untrash/markNotSpam can find threads living only there — including on a
 * flagged `\All` server, where the single global mailbox otherwise suffices.
 */
async function resolveSearchFolders(
  client: ImapFlow,
  options?: { includeTrashAndJunk?: boolean; forMutation?: boolean },
): Promise<string[]> {
  const mailboxes = await client.list();

  const allMail = findAllMailInList(mailboxes);
  const flaggedAll = mailboxes.some((mb) => mb.specialUse === "\\All");
  if (allMail && flaggedAll)
    return options?.includeTrashAndJunk
      ? [allMail, ...trashAndJunkPaths(mailboxes)]
      : [allMail];
  const onBridge = !flaggedAll && isProtonBridgeLayoutInList(mailboxes);

  const skippedSpecialUse = new Set(
    options?.includeTrashAndJunk
      ? ["\\Drafts", "\\Flagged"]
      : ["\\Trash", "\\Junk", "\\Drafts", "\\Flagged"],
  );
  const viewPaths = (path: string) =>
    path === "Labels" ||
    path.startsWith("Labels/") ||
    path === "Starred" ||
    path === "All Mail" ||
    path.endsWith("/All Mail") ||
    (onBridge && (path === "Folders" || path === "Labels"));
  const paths = mailboxes
    .filter((mb) => !mb.flags.has("\\Noselect"))
    .filter((mb) => mb.listed || onBridge)
    .filter((mb) => !mb.specialUse || !skippedSpecialUse.has(mb.specialUse))
    .filter((mb) => !viewPaths(mb.path))
    .map((mb) => mb.path);

  return paths.length > 0 ? paths : ["INBOX"];
}

// Selectable Trash/Junk mailbox paths from one LIST result. Gmail's `\All`
// omits deleted mail, so thread resolution that must reach Trash (untrash)
// searches them explicitly alongside the global mailbox.
function trashAndJunkPaths(
  mailboxes: { path: string; flags: Set<string>; specialUse?: string }[],
): string[] {
  return mailboxes
    .filter((mb) => !mb.flags.has("\\Noselect"))
    .filter((mb) => mb.specialUse === "\\Trash" || mb.specialUse === "\\Junk")
    .map((mb) => mb.path);
}

/**
 * Group qualified message ids by folder so each member's own mailbox is
 * opened exactly once. Malformed ids throw (fail closed, never act on the
 * wrong message).
 */
function groupMessageIdsByFolder(
  messageIds: string[],
): Map<string, Array<{ folder: string; uid: number }>> {
  const grouped = new Map<string, Array<{ folder: string; uid: number }>>();
  for (const messageId of new Set(messageIds)) {
    const parsed = parseImapMessageId(messageId);
    const existing = grouped.get(parsed.folder);
    if (existing) existing.push(parsed);
    else grouped.set(parsed.folder, [parsed]);
  }
  return grouped;
}

// Bridge exposes one Proton item in several real folders at once (a
// self-sent mail is ONE item with INBOX and Sent placements, each with its
// own UID but the same Message-ID). Moving one placement relocates the item,
// so a thread op must act on one placement per item: moving every snapshot
// placement re-moves an already-relocated item, and Bridge permanently
// deletes a message moved to Trash while already in Trash.
function normalizeItemMessageId(value: string | undefined): string | null {
  if (!value) return null;
  const normalized = value.replace(/^<|>$/g, "").trim().toLowerCase();
  return normalized ? normalized : null;
}

// One placement per Proton item: prefer the acted-on folder (archive/trash
// act on the INBOX placement, restores on the source-folder one), else keep
// the snapshot's first placement for determinism; members without a
// Message-ID are distinct items (cannot be grouped) and always stay.
function selectOnePlacementPerItem(
  messages: ParsedMessage[],
  preferredFolder = "INBOX",
): ParsedMessage[] {
  const byItem = new Map<string, ParsedMessage>();
  const headerless: ParsedMessage[] = [];
  for (const message of messages) {
    const itemId = normalizeItemMessageId(message.headers["message-id"]);
    if (!itemId) {
      headerless.push(message);
      continue;
    }
    const existing = byItem.get(itemId);
    if (!existing) {
      byItem.set(itemId, message);
      continue;
    }
    const prefersIncoming =
      parseImapMessageId(message.id).folder === preferredFolder &&
      parseImapMessageId(existing.id).folder !== preferredFolder;
    if (prefersIncoming) byItem.set(itemId, message);
  }
  const selected = [...byItem.values(), ...headerless];
  const order = new Map(messages.map((message, index) => [message, index]));
  selected.sort((a, b) => (order.get(a) ?? 0) - (order.get(b) ?? 0));
  return selected;
}

// Move one placement to the target, verifying the member is still where the
// snapshot said and still outside the target: an earlier move in the same op
// may have relocated the shared item (Bridge) or a concurrent client may have
// moved it. A same-folder or already-relocated member is a no-op success —
// never issue a MOVE into the folder the item already lives in.
async function movePlacementToFolder(
  client: ImapFlow,
  message: ParsedMessage,
  targetFolder: string,
): Promise<void> {
  const { folder, uid } = parseImapMessageId(message.id);
  await client.mailboxOpen(folder);
  const stillThere = await searchImapMessages(client, { uid: `${uid}` });
  if (stillThere.length === 0) return;
  const itemId = normalizeItemMessageId(message.headers["message-id"]);
  if (itemId && folder !== targetFolder) {
    await client.mailboxOpen(targetFolder);
    const inTarget = await searchImapMessages(client, {
      header: { "Message-ID": itemId },
    });
    if (inTarget.length > 0) {
      const exact = await fetchMessagesByUids(client, inTarget);
      if (
        exact.some(
          (hit) => normalizeItemMessageId(hit.headers["message-id"]) === itemId,
        )
      )
        return;
    }
    await client.mailboxOpen(folder);
  }
  await moveMessageToFolder(client, uid, targetFolder, {
    sourceFolder: folder,
  });
}

// Move id-named members one item at a time: resolve each id to its live
// copy, skip items already handled in this op, and re-check the live folder
// before every move. `onError` logs a per-member failure before it rethrows;
// restores pass a per-message target instead of a fixed folder.
async function moveMessageSnapshots(
  client: ImapFlow,
  messageIds: string[],
  resolveTarget: string | ((message: ParsedMessage) => string),
  options: {
    onBridge: boolean;
    logSkippedView: (messageId: string) => void;
    onError?: (messageId: string, error: unknown) => void;
  },
): Promise<void> {
  const handledItems = new Set<string>();
  for (const messageId of new Set(messageIds)) {
    const { folder, uid } = parseImapMessageId(messageId);
    if (options.onBridge && isFallbackViewFolder(folder)) {
      options.logSkippedView(messageId);
      continue;
    }
    await client.mailboxOpen(folder);
    const live = await fetchMessageByUid(client, uid);
    if (!live) continue;
    const target =
      typeof resolveTarget === "string" ? resolveTarget : resolveTarget(live);
    if (folder === target) continue;
    const itemId = options.onBridge
      ? normalizeItemMessageId(live.headers["message-id"])
      : null;
    if (itemId) {
      if (handledItems.has(itemId)) continue;
      handledItems.add(itemId);
    }
    try {
      await movePlacementToFolder(client, live, target);
    } catch (error) {
      if (isAlreadyInTargetError(error)) continue;
      options.onError?.(messageId, error);
      throw error;
    }
  }
}

function isNoResponse(error: unknown): boolean {
  return (
    !!error &&
    typeof error === "object" &&
    "responseStatus" in error &&
    error.responseStatus === "NO"
  );
}

// A container or a concurrently deleted folder answers SELECT with NO.
// Fan-out loops skip it and keep the rest instead of failing the whole op.
async function openFolderSkippingNo(
  client: ImapFlow,
  folder: string,
): Promise<MailboxObject | null> {
  try {
    return await client.mailboxOpen(folder, { readOnly: true });
  } catch (error) {
    if (!isNoResponse(error)) throw error;
    return null;
  }
}
// Provider-search scope (mail/v1 `search`) to the one IMAP folder that holds
// it, reusing the special-use finders so servers with renamed folders still
// resolve. Scopes with no single folder (all/archive/starred) return null
// for the whole-mailbox fallback.
async function mailboxSearchFolder(
  client: ImapFlow,
  mailbox: ProviderMailboxSearch["mailbox"],
): Promise<string | null> {
  switch (mailbox) {
    case "inbox":
      return "INBOX";
    case "sent":
      return findSentFolder(client);
    case "drafts":
      return findDraftsFolder(client);
    case "spam":
      return findJunkFolder(client);
    case "trash":
      return findTrashFolder(client);
    default:
      return null;
  }
}

// Provider-search scope to a query string for the whole-mailbox fallback.
// Folder scoping has no query equivalent here (an `in:` token would be
// dropped by parseSearchQuery), so only the text and flag filters carry
// over; the candidates_require_local_filter contract re-applies the mailbox
// scope locally.
function mailboxSearchQuery(search: ProviderMailboxSearch): string {
  const parts: string[] = [];
  if (search.text) parts.push(mailboxSearchText(search));
  if (search.read !== undefined)
    parts.push(search.read ? "is:read" : "is:unread");
  if (search.starred !== undefined)
    parts.push(search.starred ? "is:starred" : "-is:starred");
  if (search.hasAttachment) parts.push("has:attachment");
  return parts.join(" ");
}

// Scoped provider search: SEARCH text + flag criteria in one folder, ranked
// newest-first by INTERNALDATE like the listing, then sliced by offset.
async function searchImapFolder(
  client: ImapFlow,
  options: {
    folder: string;
    search: ProviderMailboxSearch;
    maxResults?: number;
    pageToken?: string;
  },
): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
  if (!(await openFolderSkippingNo(client, options.folder))) {
    return { messages: [] };
  }
  const { criteria, hasAttachment } = parseSearchQuery(
    mailboxSearchText(options.search),
  );
  if (options.search.read !== undefined)
    criteria[options.search.read ? "seen" : "unseen"] = true;
  if (options.search.starred) criteria.flagged = true;
  // IMAP TEXT covers headers + body; the local filter re-applies the scope.
  const { body, ...restCriteria } = criteria;
  const searchCriteria =
    options.search.text?.field === "any" && body !== undefined
      ? { ...restCriteria, text: body }
      : criteria;
  const uids = await searchImapMessages(client, searchCriteria);
  const dated = await fetchInternalDatesByUids(client, uids);
  dated.sort((a, b) => b.time - a.time || b.uid - a.uid);
  let candidates = dated;
  const maxResults = options.maxResults || 20;
  const offset = options.pageToken ? Number(options.pageToken) : 0;
  if (hasAttachment || options.search.hasAttachment) {
    await client.mailboxOpen(options.folder, { readOnly: true });
    const fetched = await fetchMessagesByUids(
      client,
      dated.map((entry) => entry.uid),
    );
    const keep = new Set(
      fetched
        .filter((m) => m.attachments && m.attachments.length > 0)
        .map((m) => m.id),
    );
    candidates = dated.filter((entry) =>
      keep.has(encodeImapMessageId(options.folder, entry.uid)),
    );
  }
  const pageUids = candidates
    .slice(offset, offset + maxResults)
    .map((entry) => entry.uid);
  await client.mailboxOpen(options.folder, { readOnly: true });
  const messages = await fetchMessagesByUids(client, pageUids);
  const order = new Map(pageUids.map((uid, index) => [uid, index]));
  messages.sort(
    (a, b) =>
      (order.get(parseImapMessageId(a.id).uid) ?? 0) -
      (order.get(parseImapMessageId(b.id).uid) ?? 0),
  );
  return {
    messages,
    nextPageToken:
      offset + maxResults < candidates.length
        ? String(offset + maxResults)
        : undefined,
  };
}

// The free-text part of a provider search as a query string.
function mailboxSearchText(search: ProviderMailboxSearch): string {
  if (!search.text) return "";
  const value = search.text.value.trim();
  if (!value) return "";
  if (search.text.field === "subject") return `subject:${value}`;
  if (search.text.field === "body") return value;
  return search.text.match === "phrase" ? `"${value}"` : value;
}

function groupImapSearchThreads(messages: ParsedMessage[]): EmailThread[] {
  const threadMap = new Map<string, ParsedMessage[]>();
  for (const msg of messages) {
    const existing = threadMap.get(msg.threadId) || [];
    existing.push(msg);
    threadMap.set(msg.threadId, existing);
  }
  return [...threadMap.entries()].map(([id, msgs]) => ({
    id,
    messages: msgs,
    snippet: msgs[msgs.length - 1].snippet,
  }));
}

// Rank key for a fetched page member: the provider ids are folder-qualified
// (`uid@base64url(folder)`), so decode back to the `folder:uid` rank key.
function byFolderKeyOf(messageId: string): string {
  const parsed = parseImapMessageId(messageId);
  return `${parsed.folder}:${parsed.uid}`;
}

// MIME part ids in `convertImapMessage` come from the bodyStructure tree, so
// a download for a part absent from that tree is a wrong part id, not an
// empty file. Empty parts legitimately download as zero bytes.
function findBodyPart(
  structure: { part?: string; childNodes?: unknown[] } | undefined | null,
  attachmentId: string,
): boolean {
  if (!structure || typeof structure !== "object") return false;
  if (structure.part === attachmentId) return true;
  return (
    Array.isArray(structure.childNodes) &&
    structure.childNodes.some(
      (child) =>
        !!child &&
        typeof child === "object" &&
        findBodyPart(
          child as { part?: string; childNodes?: unknown[] },
          attachmentId,
        ),
    )
  );
}

// View mailboxes on the Bridge layout never hold a message's only copy, so
// a stale id pointing there must not be mutated.
function isFallbackViewFolder(folder: string): boolean {
  return (
    folder === "All Mail" ||
    folder.endsWith("/All Mail") ||
    folder === "Labels" ||
    folder.startsWith("Labels/") ||
    folder === "Starred"
  );
}

/**
 * Resolve every message of a thread across the whole account. A thread id
 * carries its normalized root Message-ID (`t.` + base64url), so SEARCH the
 * Message-ID, References and In-Reply-To headers for the root directly —
 * no newest-N scan. Members key by full id (folder+UID): UIDs repeat
 * across folders. HEADER SEARCH is substring-based, so candidates are
 * re-verified by recomputing the thread id exactly. Trash and Junk are
 * scanned only when the caller asks for them (untrash/markNotSpam); every
 * other thread operation must not act on a deleted reply.
 */
async function fetchThreadMessages(
  client: ImapFlow,
  threadId: string,
  options?: { includeTrashAndJunk?: boolean; forMutation?: boolean },
): Promise<ParsedMessage[]> {
  // Header-less messages keep their folder-qualified listing key
  // (`uid@base64url(folder)`) in the thread id, so the owning message is
  // fetchable directly — no search needed. A key naming a view mailbox on
  // the Bridge layout (e.g. an older All Mail id) resolves there no longer:
  // like a vanished thread, treat it as not found so mutations no-op.
  const fallback = parseFallbackThreadId(threadId);
  if (fallback) {
    const onBridge = isProtonBridgeLayoutInList(await client.list());
    if (onBridge && isFallbackViewFolder(fallback.folder)) return [];
    await client.mailboxOpen(fallback.folder, { readOnly: true });
    const direct = await fetchMessageByUid(client, fallback.uid);
    return direct ? [direct] : [];
  }

  const root = parseThreadRootMessageId(threadId);
  if (!root) return [];

  const folders = await resolveSearchFolders(client, options);
  const byUid = new Map<string, ParsedMessage>();

  for (const folder of folders) {
    if (!(await openFolderSkippingNo(client, folder))) continue;
    const candidates = await searchImapMessages(client, {
      or: [
        { header: { "Message-ID": root } },
        { header: { References: root } },
        { header: { "In-Reply-To": root } },
      ],
    });
    const verified = await filterThreadCandidates(client, candidates, threadId);
    for (const message of verified) byUid.set(message.id, message);
  }

  return [...byUid.values()];
}

async function filterThreadCandidates(
  client: ImapFlow,
  uids: number[],
  threadId: string,
): Promise<ParsedMessage[]> {
  const messages = await fetchMessagesByUids(client, uids);
  return messages.filter((message) => message.threadId === threadId);
}

const PROTECTED_SPECIAL_USE = new Set([
  "\\Sent",
  "\\Trash",
  "\\Drafts",
  "\\Junk",
  "\\Archive",
  "\\All",
]);

// Remove a Proton label from every message of a thread: open Labels/<name>,
// find each member by its Message-ID header, flag \Deleted and expunge
// there ONLY. The message's real folder copy is never touched, and like
// Gmail's 200-success for a vanished thread, removing an absent label is a
// no-op rather than an error.
async function removeProtonThreadLabel(
  client: ImapFlow,
  threadId: string,
  labelId: string,
): Promise<void> {
  const labelFolder = labelId.startsWith("Labels/")
    ? labelId
    : labelPathFromName(labelId);
  const present = (await listLabelMailboxes(client)).some(
    (path) => path.toLowerCase() === labelFolder.toLowerCase(),
  );
  if (!present) return;
  // Mutating path: resolve members in their real folders, never in a view.
  const messages = await fetchThreadMessages(client, threadId, {
    forMutation: true,
  });
  const wanted = new Set(
    messages.flatMap((message) =>
      message.headers["message-id"]
        ? [message.headers["message-id"].replace(/^<|>$/g, "").toLowerCase()]
        : [],
    ),
  );
  if (wanted.size === 0) return;
  await client.mailboxOpen(labelFolder);
  for (const messageId of wanted) {
    const uids = await searchImapMessages(client, {
      header: { "Message-ID": messageId },
    });
    const hits = await fetchMessagesByUids(client, uids);
    for (const hit of hits) {
      // gluon HEADER search is a substring match: `<xr@x>` also matches
      // `<r@x>`, so only expunge the exact Message-ID.
      if (
        hit.headers["message-id"]?.replace(/^<|>$/g, "").toLowerCase() ===
        messageId
      ) {
        await client.messageDelete(String(parseImapMessageId(hit.id).uid), {
          uid: true,
        });
      }
    }
  }
}

async function assertDeletableFolder(
  client: ImapFlow,
  folderId: string,
): Promise<void> {
  if (folderId.toLowerCase() === "inbox")
    throw new SafeError("INBOX cannot be deleted");
  const mailboxes = await client.list();
  const mailbox = mailboxes.find(
    (mb) => mb.path.toLowerCase() === folderId.toLowerCase(),
  );
  if (mailbox?.specialUse && PROTECTED_SPECIAL_USE.has(mailbox.specialUse))
    throw new SafeError(
      `Special-use mailbox ${mailbox.specialUse} cannot be deleted`,
    );
}

/**
 * IMAP FROM/TO SEARCH is substring-based: `ann@x.com` also matches
 * `joann@x.com`. Fetch the candidates and keep only the UIDs whose
 * envelope address equals the target (case-insensitive).
 */
async function filterUidsByExactAddress(
  client: ImapFlow,
  uids: number[],
  header: "from" | "to",
  email: string,
): Promise<number[]> {
  const target = email.toLowerCase();
  const messages = await fetchMessagesByUids(client, uids);
  const matched: number[] = [];
  for (const message of messages) {
    const headerValue =
      header === "from" ? message.headers.from : message.headers.to;
    const exact = splitRecipientList(headerValue).some(
      (entry) => extractEmailAddress(entry).toLowerCase() === target,
    );
    if (exact) matched.push(parseImapMessageId(message.id).uid);
  }
  return matched;
}
function isAlreadyInTargetError(error: unknown): boolean {
  const text = error instanceof Error ? error.message : String(error);
  return /already (exists|in|there)|ALREADYEXISTS/i.test(text);
}

export class ImapThreadNotFoundError extends Error {
  constructor(threadId: string) {
    super(`Thread ${threadId} not found`);
    this.name = "ImapThreadNotFoundError";
  }
}

export class ImapMessageNotFoundError extends Error {
  constructor(messageId: string) {
    super(`Message ${messageId} not found`);
    this.name = "ImapMessageNotFoundError";
  }
}

// Split a MIME template with placeholder bytes into text parts and attachment
// slots: the browser fills each slot with the file's base64 and uploads the
// assembled message in chunks through the upload's PUT route.
function templateToUploadStart(
  template: Buffer,
  slots: Array<{
    part: DraftMessageUploadPart;
    encoded: string;
  }>,
): DraftAttachmentUploadStart {
  // The browser re-encodes the text parts, so a byte outside ASCII would
  // change their length and break the declared upload size.
  if (template.some((byte) => byte > 0x7f))
    throw new Error("Could not prepare the draft upload.");
  const message = template.toString("latin1");

  // Inline images sit in a related part ahead of regular attachments, so the
  // message order can differ from the slot order.
  const located = slots
    .map((slot) => ({ ...slot, index: message.indexOf(slot.encoded) }))
    .sort((a, b) => a.index - b.index);
  const parts: DraftMessageUploadPart[] = [];
  let cursor = 0;
  for (const { part, encoded, index } of located) {
    if (index < 0 || index < cursor)
      throw new Error("Could not prepare the draft upload.");
    parts.push({ type: "text", text: message.slice(cursor, index) }, part);
    cursor = index + encoded.length;
  }
  parts.push({ type: "text", text: message.slice(cursor) });

  return {
    type: "gmail-message",
    sessionUri: `imap:draft-upload:${randomBytes(16).toString("hex")}`,
    parts,
    totalBytes: getDraftMessageUploadLength(parts),
  };
}
