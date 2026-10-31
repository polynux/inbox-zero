import type { Attachment as MailAttachment } from "nodemailer/lib/mailer";
import type { ImapFlow } from "imapflow";
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
  SentMessagePage,
} from "@/utils/email/types";
import type { LocalMailSyncResponse } from "@/utils/email/local-mail-sync-types";
import type { OutlookFolder } from "@/utils/outlook/folders";
import type { InboxZeroLabel } from "@/utils/label";
import type { Logger } from "@/utils/logger";
import type { ParsedMessage } from "@/utils/types";
import type { ImapCredentialConfig } from "@/utils/imap/types";
import { UnsupportedImapOperationError } from "@/utils/imap/types";
import { SafeError } from "@/utils/error";
import { withImapConnection } from "@/utils/imap/client";
import {
  convertImapMessage,
  fetchMessageByUid,
  fetchMessagesByUids,
  fetchRecentMessages,
  parseReferencesHeader,
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
import { parseImapMessageId } from "@/utils/imap/message-id";
import { saveDraft } from "@/utils/imap/draft";
import { sendSmtpEmail } from "@/utils/imap/mail";
import {
  buildThreadId,
  parseFallbackThreadId,
  parseThreadRootMessageId,
} from "@/utils/imap/thread";
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
      if (!msg) throw new Error(`Message ${messageId} not found`);
      return msg;
    });
  }

  async getMessageByRfc822MessageId(
    rfc822MessageId: string,
  ): Promise<ParsedMessage | null> {
    return this.withConnection(async (client) => {
      const folders = await resolveSearchFolders(client);
      for (const folder of folders) {
        await client.mailboxOpen(folder, { readOnly: true });
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
    before?: Date;
    after?: Date;
    inboxOnly?: boolean;
    unreadOnly?: boolean;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    return this.withConnection(async (client) => {
      const folders =
        options.inboxOnly === false
          ? await resolveSearchFolders(client)
          : ["INBOX"];

      const maxResults = options.maxResults || 20;
      const offset = options.pageToken ? Number(options.pageToken) : 0;
      // Fetch the newest window from every folder so the merged page is
      // identical to the page a single global mailbox would return.
      const window = offset + maxResults;
      const simpleListing =
        !options.query &&
        !options.before &&
        !options.after &&
        !options.unreadOnly;

      const criteria: Record<string, unknown> = {};
      if (options.query) {
        Object.assign(criteria, parseSearchQuery(options.query));
      } else {
        criteria.all = true;
      }
      if (options.before) criteria.before = options.before;
      if (options.after) criteria.since = options.after;
      if (options.unreadOnly) criteria.unseen = true;

      const collected: ParsedMessage[] = [];
      let hasMore = false;

      for (const folder of folders) {
        const mailbox = await client.mailboxOpen(folder, { readOnly: true });

        // For simple listing without query, use sequence numbers (most reliable)
        if (simpleListing) {
          const total = mailbox.exists || 0;
          if (total > window) hasMore = true;
          if (total === 0) continue;

          const start = Math.max(1, total - window + 1);
          for await (const msg of client.fetch(`${start}:${total}`, {
            uid: true,
            envelope: true,
            flags: true,
            bodyStructure: true,
            headers: ["references"],
          })) {
            const parsed = await convertImapMessage(
              msg,
              msg.headers?.length
                ? parseReferencesHeader(msg.headers)
                : undefined,
              folder,
              mailbox.specialUse,
            );
            if (parsed) collected.push(parsed);
          }
          continue;
        }

        // For queries, use IMAP SEARCH then fetch by UID
        const allUids = await searchImapMessages(client, criteria);
        if (allUids.length > window) hasMore = true;
        const messages = await fetchMessagesByUids(
          client,
          allUids.slice(0, window),
        );
        collected.push(...messages);
      }

      collected.sort(
        (a, b) => new Date(b.date).getTime() - new Date(a.date).getTime(),
      );
      const messages = collected.slice(offset, offset + maxResults);
      const nextPageToken =
        hasMore || collected.length > window ? String(window) : undefined;

      return { messages, nextPageToken };
    });
  }

  async searchMessages(options: {
    query: string;
    maxResults?: number;
    pageToken?: string;
  }): Promise<{ messages: ParsedMessage[]; nextPageToken?: string }> {
    return this.getMessagesWithPagination({
      query: options.query,
      maxResults: options.maxResults,
      pageToken: options.pageToken,
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
      // Fall back to text search
      const allUids = await searchImapMessages(client, { all: true });
      const lastSeenUid = options.pageToken
        ? Number(options.pageToken)
        : undefined;
      const remainingUids = lastSeenUid
        ? allUids.filter((uid) => uid < lastSeenUid)
        : allUids;
      const pageUids = remainingUids.slice(0, options.maxResults || 20);

      const messages = await fetchMessagesByUids(client, pageUids);
      // Filter client-side for messages with attachments
      const withAttachments = messages.filter(
        (m) => m.attachments && m.attachments.length > 0,
      );

      return {
        messages: withAttachments,
        nextPageToken:
          remainingUids.length > pageUids.length
            ? String(pageUids[pageUids.length - 1])
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
      const uids = await searchImapMessages(
        client,
        { all: true },
        options.maxResults || 20,
      );
      const messages = await fetchMessagesByUids(client, uids);

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
      await moveMessageToFolder(client, uid, options.labelId);
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
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === archiveFolder) continue;
        await client.mailboxOpen(folder);
        try {
          await moveMessageToFolder(client, uid, archiveFolder);
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
        for (const msg of messages) {
          const { folder, uid } = parseImapMessageId(msg.id);
          if (folder === _labelId) continue;
          await client.mailboxOpen(folder);
          await moveMessageToFolder(client, uid, _labelId);
        }
      });
    }
    return this.archiveThread(threadId, ownerEmail);
  }

  async archiveMessage(messageId: string): Promise<void> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      const archiveFolder = await findArchiveFolder(client);
      if (folder === archiveFolder) return;
      await client.mailboxOpen(folder);
      await moveMessageToFolder(client, uid, archiveFolder);
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
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === trashFolder) continue;
        await client.mailboxOpen(folder);
        try {
          await moveMessageToFolder(client, uid, trashFolder);
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
    // Move to Junk/Spam folder
    const messages = await this.getThreadMessages(threadId, {
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      const junkFolder = await findJunkFolder(client);
      const targetFolder = junkFolder || "Junk";

      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === targetFolder) continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, targetFolder);
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
    const messages = await this.getThreadMessages(threadId, {
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === folderName) continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, folderName);
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
          await moveMessageToFolder(client, uid, archiveFolder);
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
          await moveMessageToFolder(client, uid, trashFolder);
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
      const uids = await searchImapMessages(
        client,
        { all: true },
        maxResults || 20,
      );
      return fetchMessagesByUids(client, uids);
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
      const pageUids = allUids.slice(offset, offset + options.maxResults);
      const messages = await fetchMessagesByUids(client, pageUids);
      const nextOffset = offset + options.maxResults;
      return {
        messages: messages.map((m) => ({ id: m.id, threadId: m.threadId })),
        nextPageToken:
          nextOffset < allUids.length ? String(nextOffset) : undefined,
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
      attachments: body.attachments,
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

  async getDrafts(options?: { maxResults?: number }): Promise<ParsedMessage[]> {
    return this.withConnection(async (client) => {
      const draftsFolder = await findDraftsFolder(client);
      if (!draftsFolder) return [];
      await client.mailboxOpen(draftsFolder, { readOnly: true });
      const uids = await searchImapMessages(
        client,
        { all: true },
        options?.maxResults || 20,
      );
      return fetchMessagesByUids(client, uids);
    });
  }

  async getDraft(draftId: string): Promise<ParsedMessage | null> {
    const { folder, uid } = parseImapMessageId(draftId);
    return this.withConnection(async (client) => {
      const draftsFolder = await findDraftsFolder(client);
      if (!draftsFolder) return null;
      if (folder !== draftsFolder) return null;
      await client.mailboxOpen(draftsFolder, { readOnly: true });
      return fetchMessageByUid(client, uid);
    });
  }

  async createDraft(params: {
    to: string;
    subject: string;
    messageHtml: string;
    replyToMessageId?: string;
  }): Promise<{ id: string }> {
    return this.withConnection(async (client) => {
      const id = await saveDraft(client, {
        from: this.config.email,
        to: params.to,
        subject: params.subject,
        html: params.messageHtml,
        inReplyTo: params.replyToMessageId,
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

  async updateDraft(
    _draftId: string,
    _params: { messageHtml?: string; subject?: string },
  ): Promise<void> {
    throw new UnsupportedImapOperationError("updateDraft");
  }

  async deleteDraft(draftId: string, _version?: string): Promise<boolean> {
    const { folder, uid } = parseImapMessageId(draftId);
    return this.withConnection(async (client) => {
      const draftsFolder = await findDraftsFolder(client);
      if (!draftsFolder) return false;
      if (folder !== draftsFolder) return false;
      await client.mailboxOpen(draftsFolder);
      await client.messageDelete(String(uid), { uid: true });
      return true;
    });
  }

  async sendDraft(
    _draftId: string,
  ): Promise<{ messageId: string; threadId: string }> {
    throw new UnsupportedImapOperationError("sendDraft");
  }

  // --- Folders ---

  async getFolders(): Promise<OutlookFolder[]> {
    return this.withConnection(async (client) =>
      listFoldersAsOutlookFolders(client),
    );
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
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        if (folder === archiveFolder) continue;
        await client.mailboxOpen(folder);
        try {
          await moveMessageToFolder(client, uid, archiveFolder);
        } catch (error) {
          if (isAlreadyInTargetError(error)) continue;
          this.logger.warn("IMAP archiveMessages move failed", {
            messageId,
            error,
          });
          throw error;
        }
      }
    });
  }

  async unarchiveMessages(messageIds: string[]): Promise<void> {
    return this.withConnection(async (client) => {
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        if (folder === "INBOX") continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, "INBOX");
      }
    });
  }

  async trashMessages(messageIds: string[]): Promise<void> {
    return this.withConnection(async (client) => {
      const trashFolder = await findTrashFolder(client);
      if (!trashFolder) return;
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        if (folder === trashFolder) continue;
        await client.mailboxOpen(folder);
        try {
          await moveMessageToFolder(client, uid, trashFolder);
        } catch {
          this.logger.warn("IMAP trashMessages move failed", { messageId });
        }
      }
    });
  }

  async untrashMessages(messageIds: string[]): Promise<void> {
    return this.withConnection(async (client) => {
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        if (folder === "INBOX") continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, "INBOX");
      }
    });
  }

  async markMessagesReadState(
    messageIds: string[],
    read: boolean,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        await client.mailboxOpen(folder);
        try {
          if (read) {
            await client.messageFlagsAdd(String(uid), ["\\Seen"], {
              uid: true,
            });
          } else {
            await client.messageFlagsRemove(String(uid), ["\\Seen"], {
              uid: true,
            });
          }
        } catch {
          this.logger.warn("IMAP markMessagesReadState failed", { messageId });
        }
      }
    });
  }

  async markMessagesStarredState(
    messageIds: string[],
    starred: boolean,
  ): Promise<void> {
    return this.withConnection(async (client) => {
      for (const messageId of new Set(messageIds)) {
        const { folder, uid } = parseImapMessageId(messageId);
        await client.mailboxOpen(folder);
        try {
          if (starred) {
            await client.messageFlagsAdd(String(uid), ["\\Flagged"], {
              uid: true,
            });
          } else {
            await client.messageFlagsRemove(String(uid), ["\\Flagged"], {
              uid: true,
            });
          }
        } catch {
          this.logger.warn("IMAP markMessagesStarredState failed", {
            messageId,
          });
        }
      }
    });
  }

  async starMessage(messageId: string): Promise<void> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
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
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === "INBOX") continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, "INBOX");
      }
    });
  }

  async unarchiveThread(threadId: string): Promise<void> {
    const messages = await this.getThreadMessages(threadId, {
      forMutation: true,
    });
    return this.withConnection(async (client) => {
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === "INBOX") continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, "INBOX");
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
      for (const msg of messages) {
        const { folder, uid } = parseImapMessageId(msg.id);
        if (folder === "INBOX") continue;
        await client.mailboxOpen(folder);
        await moveMessageToFolder(client, uid, "INBOX");
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
        await moveMessageToFolder(client, uid, archiveFolder);
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
      for (const thread of threads) {
        try {
          for (const messageId of thread.messageIds) {
            const { folder, uid } = parseImapMessageId(messageId);
            if (folder === archiveFolder) continue;
            await client.mailboxOpen(folder);
            await moveMessageToFolder(client, uid, archiveFolder);
          }
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
    const { messages, nextPageToken } = await this.getMessagesWithPagination({
      query: options.query,
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
      const labels = await listFolders(client);
      const counts: EmailFolderCount[] = [];
      for (const label of labels) {
        try {
          const status = await client.status(label.id, {
            messages: true,
            unseen: true,
          });
          counts.push({
            id: label.id,
            name: label.name,
            total: status.messages || 0,
            unread: status.unseen || 0,
          });
        } catch {
          counts.push({ id: label.id, name: label.name, total: 0, unread: 0 });
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
  // named by the folder-qualified message id and download that part.
  private async downloadAttachmentPart(
    messageId: string,
    attachmentId: string,
  ): Promise<Buffer> {
    const { folder, uid } = parseImapMessageId(messageId);
    return this.withConnection(async (client) => {
      await client.mailboxOpen(folder, { readOnly: true });
      const downloaded = await client.download(String(uid), attachmentId, {
        uid: true,
      });
      const chunks: Buffer[] = [];
      for await (const chunk of downloaded.content) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      }
      if (chunks.length === 0)
        throw new SafeError(
          `Attachment part ${attachmentId} not found on message ${messageId}`,
        );
      return Buffer.concat(chunks);
    });
  }

  async getDraftReferenceForMessage(
    _messageId: string,
  ): Promise<{ id: string; version?: string } | null> {
    // IMAP drafts are raw APPENDs with no link back to the source message
    return null;
  }

  async getDraftAttachments(
    _draftId: string,
  ): Promise<DraftAttachmentsResult | null> {
    // IMAP drafts are raw APPENDs with no attachment metadata tracking
    return null;
  }

  async addDraftAttachment(
    _draftId: string,
    _attachment: DraftAttachmentMetadata & { content: Buffer },
  ): Promise<DraftAttachmentsResult & { attachmentId: string }> {
    throw new UnsupportedImapOperationError("addDraftAttachment");
  }

  async removeDraftAttachment(
    _draftId: string,
    _attachmentId: string,
  ): Promise<DraftAttachmentsResult> {
    throw new UnsupportedImapOperationError("removeDraftAttachment");
  }

  async startDraftAttachmentUpload(
    _draftId: string,
    _attachment: DraftAttachmentMetadata,
  ): Promise<DraftAttachmentUploadStart> {
    throw new UnsupportedImapOperationError("startDraftAttachmentUpload");
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
 * use it. Bridge's unflagged "All Mail" is a view, so it serves read-only
 * lookups only, and only on the full Bridge layout (top-level `Folders`,
 * `Labels`, `All Mail`, no `\All` anywhere) — a stray user folder named
 * `All Mail` never narrows a search. Otherwise fan out over every
 * selectable real mailbox, skipping views whose messages also live in a
 * real folder (Labels mailboxes, Starred/Flagged, All Mail) so the same
 * message is not returned twice. Mutation callers (archive, trash,
 * restore, move, mark, star) pass `forMutation: true`: on Bridge layouts
 * the members then resolve in their real folders, so each member's own
 * mailbox is opened exactly once and views are never mutated. Message ids
 * in thread results then point at real folders. Trash and Junk are
 * excluded by default (listing/search scope) but included for thread
 * resolution, so untrash/markNotSpam can find threads living only there.
 */
async function resolveSearchFolders(
  client: ImapFlow,
  options?: { includeTrashAndJunk?: boolean; forMutation?: boolean },
): Promise<string[]> {
  const mailboxes = await client.list();

  const allMail = findAllMailInList(mailboxes);
  const flaggedAll = mailboxes.some((mb) => mb.specialUse === "\\All");
  const useAllMail =
    allMail &&
    (flaggedAll ||
      (!options?.forMutation && isProtonBridgeLayoutInList(mailboxes)));
  if (useAllMail) {
    if (!options?.includeTrashAndJunk) return [allMail];
    const trashAndJunk = mailboxes
      .filter(
        (mb) =>
          (mb.specialUse === "\\Trash" || mb.specialUse === "\\Junk") &&
          mb.path !== allMail,
      )
      .map((mb) => mb.path);
    return [allMail, ...trashAndJunk];
  }

  const skippedSpecialUse = new Set(
    options?.includeTrashAndJunk
      ? ["\\Drafts", "\\Flagged"]
      : ["\\Trash", "\\Junk", "\\Drafts", "\\Flagged"],
  );
  const paths = mailboxes
    .filter((mb) => mb.listed && !mb.flags.has("\\Noselect"))
    .filter((mb) => !mb.specialUse || !skippedSpecialUse.has(mb.specialUse))
    .filter((mb) => !isViewMailbox(mb.path))
    .map((mb) => mb.path);

  return paths.length > 0 ? paths : ["INBOX"];
}

// Proton view mailboxes whose messages also live in a real folder. Fanning
// out over them returns every labeled or starred message twice.
function isViewMailbox(path: string): boolean {
  if (path === "Labels" || path.startsWith("Labels/")) return true;
  if (path === "Starred") return true;
  if (path === "All Mail" || path.endsWith("/All Mail")) return true;
  return false;
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
  // fetchable directly — no search needed.
  const fallback = parseFallbackThreadId(threadId);
  if (fallback) {
    await client.mailboxOpen(fallback.folder, { readOnly: true });
    const direct = await fetchMessageByUid(client, fallback.uid);
    return direct ? [direct] : [];
  }

  const root = parseThreadRootMessageId(threadId);
  if (!root) return [];

  const folders = await resolveSearchFolders(client, options);
  const byUid = new Map<string, ParsedMessage>();

  for (const folder of folders) {
    await client.mailboxOpen(folder, { readOnly: true });
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

class ImapThreadNotFoundError extends Error {
  constructor(threadId: string) {
    super(`Thread ${threadId} not found`);
    this.name = "ImapThreadNotFoundError";
  }
}
