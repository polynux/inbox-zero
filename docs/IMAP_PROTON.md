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

STARTTLS is used when selected, never opportunistically: imapflow runs with
`doSTARTTLS: true` and nodemailer with `requireTLS: true`, so a server (or a
MITM) that hides the STARTTLS capability fails the connection instead of
continuing in cleartext.

`security: none` is only accepted for loopback hosts (`localhost`,
`127.0.0.1`, `::1`); both IMAP credential schemas reject it for any other
host, since the password would cross the network in the clear.

## Setup

1. Install Bridge, add your account (paid Proton plan required), and copy
   the Bridge password from Mailbox details.
2. Sign in to Inbox Zero with Google, Microsoft, or SSO first. Password
   sign-up is disabled (`emailAndPassword: { enabled: false }` in
   `apps/web/utils/auth.ts`); the IMAP account is linked to an existing user,
   it never creates one.
3. Set `IMAP_ALLOW_PRIVATE_HOSTS=true` (server env, e.g. `.env.local`).
   Also documented in `docs/hosting/environment-variables.mdx`.
4. Set `CRON_SECRET` to a generated value (`openssl rand -hex 32`).
5. Add the account in Settings → Accounts → Add IMAP Account → select the
   `ProtonMail Bridge` preset → use the Bridge password → Test Connection.
6. Trigger polling locally with the cron secret:
   `curl -X POST -H "Authorization: Bearer $CRON_SECRET" http://localhost:3000/api/imap/poll`
   (see `apps/web/app/api/imap/poll/route.ts`; fails closed with 401 when
   `CRON_SECRET` is unset). Production uses the Docker Compose `cron`
   container; any external scheduler hitting the endpoint every 2–5 minutes
   works.
   **What polling does**: each poll takes the oldest ≤50 new INBOX UIDs above
   `lastSeenUid` and hands each to the shared `processHistoryItem` (rules, AI,
   drafts), gated by `validateWebhookAccount` like Gmail/Outlook webhooks and
   the Redis message lock when Redis is configured. Messages labelled SENT or
   sent from the account's own address are skipped (Outlook's rule). The
   cursor advances after **each** handled UID, so a killed run resumes at the
   next message. A message failing 3 consecutive polls is logged and skipped
   (strikes persist in `ImapCredential`). A per-account DB claim
   (`pollLockedUntil`, 5 min TTL) stops overlapping polls even without Redis,
   and one poll run stops starting work after 50 s; the next tick continues.
   `newMessages` in the response counts handled messages. For local
   dev, `CRON_SECRET=... docker compose -f docker/proton-bridge/poll-cron.yml up -d`
   (only `CRON_SECRET` enters the container) polls `host.docker.internal:3000` every 60 s.
7. If port 1025/1143 is taken locally, move the other service (commonly
   Mailpit on 1025) or bump Bridge's ports in Settings → Advanced settings
   → Default ports → Change, then match the ports in the form.

## Running Bridge in Docker

`docker/proton-bridge/docker-compose.yml` builds Bridge `v3.27.1` (the latest
release) from the official
[ProtonMail/proton-bridge](https://github.com/ProtonMail/proton-bridge) source
with `docker/proton-bridge/Dockerfile`, adapted from
[`shenxn/protonmail-bridge-docker`](https://github.com/shenxn/protonmail-bridge-docker)
(entrypoint vendored from commit `5874fde`). That recipe no longer builds
current Bridge (missing `libfido2`), which is also why its Docker Hub images
stopped at 3.19.0; the fork's Dockerfile adds the FIDO2/CBOR/OpenSSL deps.
Building locally also keeps a third-party prebuilt image away from your Proton
login. The first build compiles Go and takes a few minutes.
Ports are bound to loopback only: IMAP `127.0.0.1:1143`, SMTP
`127.0.0.1:1026` (not 1025, which Mailpit commonly holds).

```sh
docker compose -f docker/proton-bridge/docker-compose.yml build
# one-time, interactive (supports 2FA)
docker compose -f docker/proton-bridge/docker-compose.yml run --rm proton-bridge init
#   bridge> login      # Proton credentials + 2FA
#   bridge> info       # copy the Bridge username/password
#   bridge> exit       # then Ctrl+C if the container doesn't stop
docker compose -f docker/proton-bridge/docker-compose.yml up -d
```

In the Add IMAP Account form, pick the `ProtonMail Bridge` preset and change
the SMTP port to `1026`. Bridge's self-signed certificate is accepted because
the host is loopback and `IMAP_ALLOW_PRIVATE_HOSTS=true`.

Probe it before linking:

```sh
cd apps/web
IMAP_ALLOW_PRIVATE_HOSTS=true \
IMAP_HOST=127.0.0.1 IMAP_PORT=1143 IMAP_USER=<bridge user> IMAP_PASS=<bridge password> \
SMTP_HOST=127.0.0.1 SMTP_PORT=1026 \
pnpm exec tsx scripts/imap-probe.ts
```

## Live Bridge probe (v3.27.1, 2026-10-31)

Results from a real account through `docker/proton-bridge`:

- CAPABILITY: `IMAP4rev1 IDLE MOVE UIDPLUS UNSELECT STARTTLS ID AUTH=PLAIN`.
  No CONDSTORE/QRESYNC, no SPECIAL-USE capability (special-use flags are still
  sent in LIST).
- LIST: `INBOX (\Inbox)`, `Sent (\Sent)`, `Drafts (\Drafts)`,
  `Archive (\Archive)`, `Spam (\Junk)`, `Trash (\Trash)`, `All Mail` with
  **no `\All` flag**, `Folders/*` (real folders), `Labels/*` (Proton labels),
  `Starred` (view of `\Flagged`).
- INBOX PERMANENTFLAGS: `$Forwarded Forwarded \Deleted \Flagged \Seen`, so
  **custom keywords are not allowed**.
- SMTP on 1026: STARTTLS, `AUTH PLAIN LOGIN`, auth OK.
- Bridge's IMAP engine (gluon) implements `SEARCH HEADER` as a
  case-insensitive substring match.

How the provider uses this:

- The Bridge layout is detected only when top-level `INBOX`, `Folders`,
  `Labels` and `All Mail` all exist and no mailbox has `\All`; a lone user
  folder named `Labels` or `All Mail` on another server changes nothing.
- On that layout, `All Mail` serves read-only searches. Every mutation
  (archive, trash, restore, spam, move, mark read/star) resolves thread
  members over **real folders only** (INBOX, Archive, Sent, Drafts, Spam,
  Trash, `Folders/*`) and acts on the copy there; All Mail, `Labels/*` and
  `Starred` are never mutated, except removing a label inside
  `Labels/<name>`. Each Proton message lives in exactly one real folder, so
  this yields one copy per message (one SEARCH per real folder per op).
- Proton labels are `Labels/<name>` mailboxes: applying a label COPYs the
  message there, removing it expunges the copy in `Labels/<name>` only, and
only hits whose Message-ID matches exactly (gluon's HEADER search is a
substring match). Removing an absent label is a no-op. The
  message stays in its real folder. Servers without a `Labels` namespace keep
  folder (MOVE) semantics.

## Upgrading from bare-UID ids

Older builds used a bare UID as an IMAP message id, a 24-hex hash as the
thread id, and `fallback:<folder>:<uid>` for header-less thread ids. Thread
ids are now `t.` + base64url(normalized root Message-ID). Ids are now
folder-qualified (`<uid>@base64url(folder)`), and header-less threads carry
that same key, so rows persisted against the old ids no longer resolve. On a
dev database, clear the account's cached mail rows (at least `EmailMessage`
and `ExecutedRule`) and reset `ImapCredential.lastSeenUid`/`uidValidity`, or delete and
re-link the account, so the next sync repopulates. Schema changes ship as
migrations `20261030194839_add_imap_credential` and
`20261031120000_add_imap_poll_state`.

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
| Inbox read | code-verified | `getInboxMessages`/`getMessagesWithPagination` open INBOX read-only; ids are folder-qualified (`uid@base64url(folder)`) so same-UID messages in different folders never collide |
| Archive | code-verified | Per-member: each thread member moves from its own mailbox to `findArchiveFolder` (special-use `\Archive`, name fallback, creates `Archive`); members already there are skipped; vanished thread is a no-op like Gmail's 404-success |
| Trash | code-verified | Per-member: each thread member moves from its own mailbox to `findTrashFolder` (special-use `\Trash`, name fallback; no-op when absent); vanished thread is a no-op like Gmail's 200-success |
| Unarchive/untrash/restore | code-verified | Per-member restore from the member's own mailbox to INBOX; the id names the source folder, so a same-UID message elsewhere is never touched |
| Thread resolution | code-verified | Thread id = `t.` + base64url(root Message-ID). Every fetch path requests the References header, so listing, search, thread ops, `getMessage` and poll compute the same id. Lookup runs `SEARCH OR HEADER Message-ID/References/In-Reply-To <root>` per folder (All Mail for reads on Bridge, real folders for mutations) and re-verifies the thread id exactly (Bridge's HEADER search is substring); malformed roots (no `@`, whitespace, <3 chars) are rejected. No window limit. Members key by full id (folder+UID). Trash and Junk are searched **only** for `untrashThread`/`markNotSpam`. Header-less threads carry the folder-qualified key and fetch directly |
| Labels | live-probed + code-verified | On Bridge, labels are `Labels/<name>` mailboxes: apply = COPY (creates the mailbox), remove = expunge the copy in `Labels/<name>` only, found by Message-ID; the message stays in its real folder. Bridge allows no custom keywords (PERMANENTFLAGS has no `\*`). Servers without a `Labels` namespace keep folder semantics |
| Send | code-verified | `sendSmtpEmail` via nodemailer STARTTLS to Bridge SMTP; Bridge saves to Sent automatically per Bridge docs |
| Drafts | code-verified | `saveDraft` APPENDs with `\Draft` and returns a folder-qualified id; creates the Drafts folder when missing, else a clear error. Every draft carries a generated `Message-ID`, so when APPEND returns no UID (no UIDPLUS) the draft is found again by `HEADER Message-ID` in Drafts; only if that search misses does it throw (retrying would duplicate the draft). `sendDraft`/`updateDraft` explicitly unsupported |
| Role labels | code-verified | `convertImapMessage` maps roles from `specialUse` flags + exact well-known names only; user folders like `Presentations` never read as SENT |
| Search | code-verified | `parseSearchQuery` covers from/to/subject/body/phrases/dates; `inboxOnly: false` searches the all-mail mailbox (`\All`, or `All Mail` by name on Bridge), else every selectable non-view mailbox except trash/junk/drafts; exact-address filters run before slicing |
| Local mail sync | code-verified | `getImapMailboxSyncPage` delta/snapshot with UIDVALIDITY reset. Deletions: one `UID SEARCH ALL` per delta run (on its last page) reports UIDs missing since the last run (`knownUids` range-set in the cursor) as `deletedMessageIds`; Bridge has no CONDSTORE/QRESYNC |
| Automation on new mail (rules, AI) | code-verified | `/api/imap/poll` feeds the oldest ≤50 new INBOX messages per poll into `processHistoryItem`, gated by `validateWebhookAccount`; cursor advances only past handled UIDs |
| AI assistant | code-verified (provider surface) | `ImapProvider` implements the full `EmailProvider` interface and is constructed by the provider factory, so assistant flows run against it; answer quality on real mail is a **needs live check** |
| Attachments (list) | code-verified | `convertImapMessage` populates filename/mimeType/size/part-id metadata from `bodyStructure`; `attachments` drives rules/AI, `inline` holds inline parts |
| Attachment download | code-verified | `getAttachment`/`getAttachmentStream` download the MIME part by UID from the message's own folder (`client.download`), returning base64 data like Gmail/Outlook; a missing part throws |

## Known limitations

- **Polling instead of IDLE**: Bridge supports IDLE, but the app polls (cron every 1–2 min); at most 50 new messages per account and 50 s per run.
- **First poll baselines**: linking seeds `lastSeenUid` from `uidNext - 1` and `uidValidity`; a null cursor baselines instead of reporting history as new. A UIDVALIDITY change re-baselines (messages that arrived during the reset are not processed).
- **No UIDNEXT, no baseline**: a server that omits UIDNEXT leaves `lastSeenUid` unset; polling reports 0 until the server provides one.
- **Deletion detection costs one `UID SEARCH ALL` per delta run**; deletions made before a cursor gained `knownUids` are not reported.
- **Thread ops cost one SEARCH per real folder** on Bridge (mutations avoid All Mail), so accounts with many `Folders/*` pay more round trips.
- **Drafts need a recoverable UID**: APPEND without UIDPLUS is recovered by searching the draft's generated Message-ID (Bridge has UIDPLUS).
- **`sendDraft`/`updateDraft` unsupported**: drafts can be created/listed/deleted, not edited or sent from the draft store.
- **Label removal on non-Proton servers is a no-op**: folder-based labels can't be detached without moving the message.
- **No filters/contacts/signatures/send-as**: `createFilter`, contacts, signatures, and send-as return empty or throw.
- **Bridge offline = connection refused**: the app surfaces the IMAP/SMTP error; nothing queues locally.
- **DNS rebinding**: hosts are re-validated on every IMAP connection and SMTP send, but the client still does its own lookup (no IP pinning), so a sub-second rebind race remains. Irrelevant for loopback Bridge.
