// Run with: `cd apps/web && pnpm exec tsx scripts/imap-probe.ts`
// Reads connection settings from env vars, connects to an IMAP server
// (e.g. Proton Mail Bridge on 127.0.0.1), and prints capabilities,
// special-use folders, INBOX permanent flags, and an SMTP EHLO/auth check.
// Read-only on the mailbox: never appends, moves, or deletes.
import "dotenv/config";
import { createConnection } from "node:net";
import type { Socket } from "node:net";
import { connect } from "node:tls";
import { ImapFlow } from "imapflow";
import nodemailer from "nodemailer";
import { shouldBypassTlsVerification } from "@/utils/imap/tls";

const REQUIRED_ENV = [
  "IMAP_HOST",
  "IMAP_PORT",
  "IMAP_USER",
  "IMAP_PASS",
  "SMTP_HOST",
  "SMTP_PORT",
] as const;

type ProbeConfig = {
  imapHost: string;
  imapPort: number;
  imapSecure: boolean;
  imapUser: string;
  imapPass: string;
  smtpHost: string;
  smtpPort: number;
  smtpSecure: boolean;
  smtpUser: string;
  smtpPass: string;
};

function readConfig(): ProbeConfig {
  for (const name of REQUIRED_ENV) {
    if (!process.env[name]) throw new Error(`Missing ${name} env var`);
  }

  return {
    imapHost: process.env.IMAP_HOST as string,
    imapPort: Number(process.env.IMAP_PORT),
    imapSecure: (process.env.IMAP_SECURE || "starttls").toLowerCase() === "tls",
    imapUser: process.env.IMAP_USER as string,
    imapPass: process.env.IMAP_PASS as string,
    smtpHost: process.env.SMTP_HOST as string,
    smtpPort: Number(process.env.SMTP_PORT),
    smtpSecure: (process.env.SMTP_SECURE || "starttls").toLowerCase() === "tls",
    smtpUser: (process.env.SMTP_USER || process.env.IMAP_USER) as string,
    smtpPass: (process.env.SMTP_PASS || process.env.IMAP_PASS) as string,
  };
}

async function probeImap(config: ProbeConfig) {
  const client = new ImapFlow({
    host: config.imapHost,
    port: config.imapPort,
    secure: config.imapSecure,
    auth: { user: config.imapUser, pass: config.imapPass },
    logger: false,
    // Bridge's self-signed certificate is only accepted on loopback, same rule
    // as the app; a remote host must present a valid certificate.
    tls: { rejectUnauthorized: !bypassTlsVerification(config.imapHost) },
  });

  await client.connect();

  try {
    console.log("\n## CAPABILITY");
    const capabilities = [...client.capabilities.keys()].sort();
    for (const capability of capabilities) console.log(`- ${capability}`);
    const wanted = ["IDLE", "MOVE", "UIDPLUS", "CONDSTORE", "SPECIAL-USE"];
    console.log(
      `wanted: ${wanted.map((w) => `${w}=${capabilities.includes(w)}`).join(" ")}`,
    );

    console.log("\n## LIST with special-use flags");
    for (const mailbox of await client.list()) {
      console.log(
        `- ${mailbox.path}${mailbox.specialUse ? ` (${mailbox.specialUse})` : ""}`,
      );
    }

    console.log("\n## INBOX");
    const inbox = await client.mailboxOpen("INBOX", { readOnly: true });
    const permanentFlags = inbox.permanentFlags
      ? [...inbox.permanentFlags].sort()
      : [];
    console.log(`messages: ${inbox.exists ?? 0}`);
    console.log(`permanentFlags: ${permanentFlags.join(" ") || "(none)"}`);
    console.log(
      `custom keywords allowed: ${permanentFlags.includes("\\*") || inbox.permanentFlags === undefined}`,
    );
  } finally {
    await client.logout();
  }
}

// Raw EHLO so we can print the server's capability list and auth mechanisms.
// nodemailer's verify() only reports success/failure.
async function readEhloReply(
  host: string,
  port: number,
  secure: boolean,
  rejectUnauthorized: boolean,
): Promise<string[]> {
  const { promise, resolve, reject } = Promise.withResolvers<string[]>();
  const socket: Socket = secure
    ? connect({ host, port, rejectUnauthorized })
    : createConnection({ host, port });

  let greeted = false;
  let buffer = "";

  const fail = (error: Error) => {
    socket.destroy();
    reject(error);
  };

  socket.setTimeout(15_000, () => fail(new Error("SMTP probe timed out")));
  socket.on("error", fail);
  socket.on("data", (chunk: Buffer) => {
    buffer += chunk.toString("utf8");
    if (!greeted) {
      if (!buffer.includes("\r\n")) return;
      greeted = true;
      buffer = "";
      socket.write("EHLO imap-probe\r\n");
      return;
    }
    if (/\r\n\d{3} /.test(buffer)) {
      socket.end();
      resolve(
        buffer
          .split("\r\n")
          .filter(Boolean)
          .map((line) => line.replace(/^\d{3}[- ]/, "")),
      );
    }
  });

  return promise;
}

async function probeSmtp(config: ProbeConfig) {
  console.log("\n## SMTP EHLO");
  for (const line of await readEhloReply(
    config.smtpHost,
    config.smtpPort,
    config.smtpSecure,
    !bypassTlsVerification(config.smtpHost),
  )) {
    console.log(`- ${line}`);
  }

  console.log("\n## SMTP auth");
  const transport = nodemailer.createTransport({
    host: config.smtpHost,
    port: config.smtpPort,
    secure: config.smtpSecure,
    auth: { user: config.smtpUser, pass: config.smtpPass },
    tls: { rejectUnauthorized: !bypassTlsVerification(config.smtpHost) },
  });

  try {
    await transport.verify();
    console.log("auth: ok");
  } finally {
    transport.close();
  }
}

async function main() {
  const config = readConfig();
  await probeImap(config);
  await probeSmtp(config);
}

main().catch((error) => {
  console.error(
    `probe failed: ${error instanceof Error ? error.message : error}`,
  );
  process.exit(1);
});

// Same rule as the app: skip verification only for a loopback host, and only
// when the operator opted in (Bridge's certificate is self-signed).
function bypassTlsVerification(host: string): boolean {
  return shouldBypassTlsVerification(host, {
    allowPrivateHosts: process.env.IMAP_ALLOW_PRIVATE_HOSTS === "true",
  });
}
