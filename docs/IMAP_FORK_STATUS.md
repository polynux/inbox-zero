# IMAP / Proton Bridge fork: status and remaining work

Snapshot: 2026-11-01, branch `imap` @ `b5e9e74d9` (plus this doc). Base: upstream
`elie222/inbox-zero` main @ `56303d122`. Setup and behaviour: `docs/IMAP_PROTON.md`.

## What works (verified live against a real Proton Bridge v3.27.1 account)

- Linking an IMAP account (Bridge preset, loopback TLS, SSRF guard with
  `IMAP_ALLOW_PRIVATE_HOSTS`), inbox listing (< 1 s per page), pagination.
- Threading on Proton (internal `@protonmail.internalid` references ignored),
  folder-qualified message ids, real-folder resolution (All Mail / Labels /
  Starred never mutated).
- Archive (+ "Move to inbox"), trash (no data loss: Trash→Trash guard,
  one placement per Proton item), spam, star, read/unread.
- Proton labels as `Labels/<name>` (COPY to apply, expunge copy to remove) at
  provider level.
- Compose: drafts with stable ids, attachments (incl. chunked > 3 MB), send,
  draft removed after send, Sent observation; reply threading when the reply
  draft is created and sent in the same session.
- Search (folder-scoped incl. `?type=trash&q=`), attachments download,
  polling every 60 s feeding rules/AI, UIDVALIDITY, deletion sync.

## Must-fix

1. **Thread briefly empties / shows duplicate members after a move.**
   `apps/web/utils/imap/local-mail-sync.ts` (~205-235) tombstones the old
   folder-qualified id in the source folder's stream; the moved copy arrives in
   the destination folder's own stream page, which can apply later (≤ 60 s for
   system folders, up to 10 min for custom `Folders/*` scopes —
   `LOW_PRIORITY_CATCH_UP_INTERVAL_MS` in `packages/mail-core/src/engine.ts`).
   Final state is correct, but the UI can show the thread gone (old page first)
   or with 4 members for 2 messages (new page first). Fix: defer the tombstone
   until the destination page lands (pending-delete with timeout), or apply
   destination pages for moved ids first. Remove the false "same batch"
   comment.
2. **Reply threading lost when a reply draft is reopened.**
   `getDrafts` (`apps/web/utils/email/imap.ts` ~1344) and
   `getDraftReferenceForMessage` (~2281) rebuild draft ids without the parent
   part (`d.<draft>` instead of `d.<draft>.<parent>`). Bridge strips
   `In-Reply-To` and rewrites `References` on APPEND, so after a reload the
   reply is sent into a new thread. Persist the parent per draft (record keyed
   by the draft Message-ID, or a header Bridge preserves) and return the
   3-part id from both listing paths.
3. **Send silently does nothing with a restored stale draft** (UI round 5, B1).
   A composer restored onto a draft whose attachment 404s closes in ~2 s, sends
   nothing, shows no error. Must either send or show an error.
4. **Label views and label persistence in the UI** (UI round 5, B3).
   `?type=label&labelId=Labels/<name>` shows 0 rows while a thread is
   labelled; the chip disappears after reload; chip × removal untestable.
   Root: label placements live in `Labels/*` mailboxes and never enter the
   cached row's `label_ids_json`; only in-session applies are synthesized, and
   externally applied Proton labels never appear (no `Labels/*` streams).

## Should-fix

- **Reply draft left behind after a UI reply send** (UI round 5, B2) — fresh
  compose sends now remove their draft; replies still don't.
- **Draft "Delete" in the thread toolbar fires no request** (B4);
  composer "Discard draft" works.
- **Hydration slow**: 88 of 94 slow requests in round 5 were
  `POST …/hydration` (max 16.4 s); one row-click hung > 90 s on
  "Loading message…" (B5).
- **Out-of-range UID pauses hydration**: an id above UIDNEXT throws imapflow
  `Command failed` instead of not-found (`imap.ts` ~130-141,
  `utils/mail-api/source.ts` ~577).
- **Gmail/Outlook behaviour changes to confirm or gate** (round-5 review):
  `withAppliedLabel` in `utils/mail-api/operations.ts` (~123/179/867) now
  self-attests label applies for all providers — gate on IMAP; search in
  `MailShell.tsx` (~393) now ANDs the sidebar scope for every provider
  (product decision).
- Noisy 404s for stale draft ids around compose/send (B6).

## Nits

- Miss classification by English regex on `SafeError` messages
  (`source.ts` ~592-600) — prefer error classes/codes.
- `apps/web/app/api/mail/v1/devices/route.ts` (~112) still has an unguarded
  `await request.json()`.
- Identical files on one draft are indistinguishable for removal (derived
  `filename:mime:size` ids, since Bridge strips `X-Attachment-Id`).
- `imapRoleLabel` (`utils/imap/message.ts` ~624) treats any leaf named `Sent`
  as SENT.
- Type-only assertions in `local-mail-sync.test.ts`; `withAppliedLabel`
  placed mid-file in `operations.ts`.
- Not started: unarchive/restore buttons don't exist upstream either; IDLE
  push (polling only); DNS pinning for remote IMAP hosts (re-validation only);
  draft `Importance`/`X-*` headers; Reply-To on drafts (dropped commit kept on
  branch `imap-r5-drafts-replyto-dropped` — it touched Gmail/Outlook and
  Bridge overwrites Reply-To anyway).

## Local environment left running

- systemd user units: `iz-dev` (Next dev on :3000, log `/tmp/inbox-zero-dev.log`),
  `iz-google` (Google sign-in emulator :4002), `iz-llm` (LLM emulator :4006).
  Stop with `systemctl --user stop iz-dev iz-google iz-llm`.
- Docker: dev Postgres/Redis (`docker-compose.dev.yml`), Proton Bridge
  (`docker/proton-bridge/docker-compose.yml`), poll cron
  (`docker/proton-bridge/poll-cron.yml`).
- `apps/web/.env` is generated (gitignored); `DEFAULT_LLMS` points at the LLM
  emulator — set a real provider + `LLM_API_KEY` to test AI features.
- Worktrees under `~/.t3/worktrees/inbox-zero/` (`imap-fix-correctness`,
  `imap-fix-security`, `imap-rebase`) are merged into `imap` and can be
  removed.
- All agent test mail (`IZTEST*`, `IZPROV*`, `IZR*`, `IZVERIFY*`, `IZTRASH*`)
  was cleaned up; no test labels remain.
