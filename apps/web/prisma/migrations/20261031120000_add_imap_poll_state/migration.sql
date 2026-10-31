-- UIDVALIDITY was added to the already-shipped create-table migration; this
-- migration carries it for databases that ran the original one.
ALTER TABLE "ImapCredential" ADD COLUMN "uidValidity" BIGINT;

-- Poll strikes must survive a restart: serverless invocations share no memory.
ALTER TABLE "ImapCredential" ADD COLUMN "pollFailureUid" INTEGER;
ALTER TABLE "ImapCredential" ADD COLUMN "pollFailureCount" INTEGER NOT NULL DEFAULT 0;

-- Claim so overlapping ticks do not process the same mail twice.
ALTER TABLE "ImapCredential" ADD COLUMN "pollLockedUntil" TIMESTAMP(3);
