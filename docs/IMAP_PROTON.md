# Proton Mail via Bridge (IMAP/SMTP)

Local-dev-only path for running Inbox Zero against Proton Mail through
[Proton Mail Bridge](https://proton.me/mail/bridge), which exposes the
account as a local IMAP/SMTP server. Fork-owned; no live Bridge was
available on this machine, so everything below is verified from code and
Bridge docs unless marked **needs live check**.

## Verified Bridge defaults

| Setting | Value | Source |
|---|---|---|
| IMAP host/port | `127.0.0.1:1143` | [Port error guide](https://proton.me/support/port-already-occupied-error): "Bridge uses ports 1143 and 1025 by default" |
| SMTP host/port | `127.0.0.1:1025` | Same guide |
| Connection mode | STARTTLS by default, switchable to SSL in Bridge settings | [Settings guide](https://proton.me/support/comprehensive-guide-to-bridge-settings): Connection mode "allows you to change the protocol between STARTTLS and SSL" |
| TLS certificate | Self-signed, generated on first setup; clients must confirm a certificate exception manually | [SSL connection issues](https://proton.me/support/bridge-ssl-connection-issue): "self-signed TLS/SSL certificate that Bridge generates when you first set up the app"; [Thunderbird guide](https://proton.me/support/protonmail-bridge-clients-windows-thunderbird): "this can't be signed by a certificate authority and requires you to confirm it manually" |
| Password | Bridge-generated password from Mailbox details, NOT the Proton account password | [Thunderbird guide](https://proton.me/support/protonmail-bridge-clients-windows-thunderbird): "Make sure you enter the password displayed in the Bridge Mailbox details rather than your regular account password" |
| Port conflicts | Anything else on 1025/1143 (e.g. Mailpit on 1025) must move; Bridge also lets you change its ports (Settings → Advanced → Default ports → Change) | [Port error guide](https://proton.me/support/port-already-occupied-error) |

The preset in `apps/web/utils/imap/types.ts` (`ProtonMail Bridge`,
`127.0.0.1:1143` / `127.0.0.1:1025`, STARTTLS on both) matches these
defaults.

## TLS decision

`rejectUnauthorized` was `true` for every non-`none` connection, so
connecting to Bridge always failed certificate validation. The app now
skips verification **only when both hold**:

- the host is loopback (`localhost`, `127.0.0.1`, `::1`), and
- `IMAP_ALLOW_PRIVATE_HOSTS=true` (the same flag that gates the
  host allow-list at link time).

Remote hosts always verify. See `apps/web/utils/imap/tls.ts` and its
tests. The shipped default stays deny-all.

## Setup

1. Install Bridge, add your account (paid Proton plan required), and copy
   the Bridge password from Mailbox details.
2. Set `IMAP_ALLOW_PRIVATE_HOSTS=true` (server env, e.g. `.env.local`).
   Also documented in `docs/hosting/environment-variables.mdx`.
3. Set `CRON_SECRET` to a generated value (`openssl rand -hex 32`).
4. Add the account in Settings → Accounts → Add IMAP Account → select the
   `ProtonMail Bridge` preset → use the Bridge password → Test Connection.
5. Trigger polling locally with the cron secret:
   `curl -X POST -H "Authorization: Bearer $CRON_SECRET" http://localhost:3000/api/imap/poll`
   (see `apps/web/app/api/imap/poll/route.ts`; fails closed with 401 when
   `CRON_SECRET` is unset). Production uses the Docker Compose `cron`
   container; any external scheduler hitting the endpoint every 2–5 minutes
   works.
6. If port 1025/1143 is taken locally, move the other service (commonly
   Mailpit on 1025) or bump Bridge's ports in Settings → Advanced settings
   → Default ports → Change, then match the ports in the form.

## Probe script

`apps/web/scripts/imap-probe.ts` (run with `cd apps/web && pnpm exec tsx
scripts/imap-probe.ts`) reads `IMAP_HOST/PORT/USER/PASS`,
`SMTP_HOST/PORT/USER/PASS` (plus optional `IMAP_SECURE`, `SMTP_SECURE`,
`SMTP_USER`, `SMTP_PASS`) and prints: CAPABILITY list with
IDLE/MOVE/UIDPLUS/CONDSTORE/SPECIAL-USE presence, LIST with special-use
flags, INBOX PERMANENTFLAGS (custom-keyword support), and an SMTP
EHLO/auth check. Read-only; user runs it against their own Bridge.

## Capability matrix

| Feature | Status | Basis |
|---|---|---|
| Inbox read | code-verified | `getInboxMessages`/`getMessagesWithPagination` open INBOX read-only |
| Archive | code-verified | `archiveThread` moves to `findArchiveFolder` (special-use `\Archive`, name fallback, creates `Archive`) |
| Trash | code-verified | `trashThread` moves to `findTrashFolder` (special-use `\Trash`, name fallback; no-op when absent) |
| Labels | code-verified with folder fallback | Labels are folders (`getOrCreateFolder`); `removeThreadLabel(s)` are documented no-ops since a message lives in one folder. Custom keywords are NOT written: no `messageFlagsAdd` of `$`-labels anywhere outside `\Seen`/`\Flagged`. Whether Bridge advertises `PERMANENTFLAGS \*` is a **needs live check** (probe prints it) |
| Send | code-verified | `sendSmtpEmail` via nodemailer STARTTLS to Bridge SMTP; Bridge saves to Sent automatically per Bridge docs |
| Drafts | code-verified | `saveDraft` APPENDs with `\Draft`; creates the Drafts folder when missing, else a clear error. `sendDraft`/`updateDraft` explicitly unsupported |
| Search | code-verified, INBOX-only by default | `parseSearchQuery` covers from/to/subject/body/phrases/dates; `inboxOnly: false` searches Bridge's `\All` mailbox when listed, else every selectable mailbox except trash/junk/drafts |
| Local mail sync | code-verified | `getImapMailboxSyncPage` delta/snapshot with UIDVALIDITY reset handling |
| AI assistant | code-verified (provider surface) | `ImapProvider` implements the full `EmailProvider` interface and is constructed by the provider factory, so assistant flows run against it; answer quality on real mail is a **needs live check** |
| Attachments (list) | code-verified | `convertImapMessage` populates filename/mimeType/size/part-id metadata from `bodyStructure`; `attachments` drives rules/AI, `inline` holds inline parts |
| Attachment download | unsupported by design | `getAttachment`/`getAttachmentStream` throw `UnsupportedImapOperationError`; metadata (incl. part id) is recorded, bytes stay on the server |

## Known limitations

- **DNS rebinding**: hosts are validated only at link time (`validateImapHosts` on test-connection and create). Acceptable for trusted single-tenant local use.
- **Polling instead of IDLE**: `/api/imap/poll` compares UIDNEXT against `lastSeenUid`. No IDLE listener.
- **Deletions invisible**: expunged UIDs are not reported as deletions (`deletedMessageIds` always `[]`); removals only surface as a full reset when UIDVALIDITY changes.
- **`getAttachmentStream` unsupported**: attachment viewers/download flows that need bytes throw.
- **`sendDraft`/`updateDraft` unsupported**: drafts can be created/listed/deleted, not edited or sent from the draft store.
- **`removeThreadLabel(s)` are no-ops**: folder-based labels can't be detached without moving the message.
- **No filters/contacts/signatures/send-as**: `createFilter`, contacts, signatures, and send-as return empty or throw.
- **Bridge offline = connection refused**: the app surfaces the IMAP/SMTP error; nothing queues locally.
