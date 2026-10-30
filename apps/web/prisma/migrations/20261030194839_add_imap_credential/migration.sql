-- CreateTable
CREATE TABLE "ImapCredential" (
    "id" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,
    "imapHost" TEXT NOT NULL,
    "imapPort" INTEGER NOT NULL DEFAULT 993,
    "imapSecurity" TEXT NOT NULL DEFAULT 'tls',
    "smtpHost" TEXT NOT NULL,
    "smtpPort" INTEGER NOT NULL DEFAULT 587,
    "smtpSecurity" TEXT NOT NULL DEFAULT 'starttls',
    "username" TEXT NOT NULL,
    "password" TEXT NOT NULL,
    "lastPolledAt" TIMESTAMP(3),
    "lastSeenUid" INTEGER,
    "accountId" TEXT NOT NULL,

    CONSTRAINT "ImapCredential_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "ImapCredential_accountId_key" ON "ImapCredential"("accountId");

-- AddForeignKey
ALTER TABLE "ImapCredential" ADD CONSTRAINT "ImapCredential_accountId_fkey" FOREIGN KEY ("accountId") REFERENCES "Account"("id") ON DELETE CASCADE ON UPDATE CASCADE;
