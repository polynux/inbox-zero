import { NextResponse } from "next/server";
import { env } from "@/env";
import { withAuth } from "@/utils/middleware";
import { imapCredentialSchema } from "@/utils/actions/imap.validation";
import { encryptToken } from "@/utils/encryption";
import { createImapConnection } from "@/utils/imap/client";
import { validateImapHosts } from "@/utils/imap/host-validation";
import { isDuplicateError } from "@/utils/prisma-helpers";
import prisma from "@/utils/prisma";

export const POST = withAuth("imap/linking/create", async (request) => {
  const userId = request.auth.userId;
  const body = await request.json();
  const parsed = imapCredentialSchema.safeParse(body);

  if (!parsed.success) {
    return NextResponse.json(
      { error: "Invalid input", details: parsed.error.flatten() },
      { status: 400 },
    );
  }

  const data = parsed.data;

  const hostValidation = await validateImapHosts(
    [data.imapHost, data.smtpHost],
    {
      allowPrivateHosts: env.IMAP_ALLOW_PRIVATE_HOSTS,
    },
  );
  if (!hostValidation.valid) {
    return NextResponse.json({ error: hostValidation.error }, { status: 400 });
  }

  const normalizedEmail = data.email.trim().toLowerCase();

  // Check if email is already linked
  const existingEmailAccount = await prisma.emailAccount.findUnique({
    where: { email: normalizedEmail },
    select: { id: true, userId: true },
  });

  if (existingEmailAccount) {
    if (existingEmailAccount.userId !== userId) {
      return NextResponse.json(
        { error: "This email is already linked to another account" },
        { status: 409 },
      );
    }
    return NextResponse.json(
      { error: "This email is already linked to your account" },
      { status: 409 },
    );
  }

  // Encrypt the password
  const encryptedPassword = encryptToken(data.password);
  if (!encryptedPassword) {
    return NextResponse.json(
      { error: "Failed to encrypt credentials" },
      { status: 500 },
    );
  }

  // Seed the poll cursor from the mailbox's uidNext/uidValidity so the first
  // poll is a baseline (0 new). Without this the first poll would report the
  // whole pre-link mailbox as new. A failed probe leaves the cursor null and
  // the first poll baselines instead — same 0-new outcome.
  let lastSeenUid: number | null = null;
  let uidValidity: bigint | null = null;
  try {
    const probe = createImapConnection({
      email: normalizedEmail,
      emailAccountId: "",
      imapHost: data.imapHost,
      imapPort: data.imapPort,
      imapSecurity: data.imapSecurity,
      password: data.password,
      smtpHost: data.smtpHost,
      smtpPort: data.smtpPort,
      smtpSecurity: data.smtpSecurity,
      username: data.username,
    });
    await probe.connect();
    try {
      const mailbox = await probe.mailboxOpen("INBOX", { readOnly: true });
      lastSeenUid =
        mailbox.uidNext === undefined ? null : Number(mailbox.uidNext) - 1;
      uidValidity =
        mailbox.uidValidity === undefined ? null : BigInt(mailbox.uidValidity);
    } finally {
      await probe.logout().catch(() => {});
    }
  } catch {
    lastSeenUid = null;
    uidValidity = null;
  }

  // Create Account + ImapCredential + EmailAccount in a single transaction
  try {
    const result = await prisma.account.create({
      data: {
        userId,
        provider: "imap",
        providerAccountId: normalizedEmail,
        type: "credential",
        imapCredential: {
          create: {
            imapHost: data.imapHost,
            imapPort: data.imapPort,
            imapSecurity: data.imapSecurity,
            smtpHost: data.smtpHost,
            smtpPort: data.smtpPort,
            smtpSecurity: data.smtpSecurity,
            username: data.username,
            password: encryptedPassword,
            ...(lastSeenUid != null ? { lastSeenUid } : {}),
            ...(uidValidity != null ? { uidValidity } : {}),
          },
        },
        emailAccount: {
          create: {
            email: normalizedEmail,
            name: data.name || normalizedEmail.split("@")[0],
            userId,
          },
        },
      },
      include: {
        emailAccount: { select: { id: true, email: true } },
      },
    });

    return NextResponse.json({
      emailAccountId: result.emailAccount?.id,
      email: result.emailAccount?.email,
    });
  } catch (error) {
    // A concurrent request can pass the lookup above and then hit the unique index.
    if (isDuplicateError(error)) {
      return NextResponse.json(
        { error: "This email is already linked to an account" },
        { status: 409 },
      );
    }
    throw error;
  }
});
