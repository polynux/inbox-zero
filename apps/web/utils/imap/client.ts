import { ImapFlow } from "imapflow";
import { env } from "@/env";
import { shouldBypassTlsVerification } from "@/utils/imap/tls";
import type { ImapCredentialConfig } from "@/utils/imap/types";
import { createScopedLogger } from "@/utils/logger";

const logger = createScopedLogger("imap/client");

const CONNECT_TIMEOUT_MS = 30_000;
export function createImapConnection(config: ImapCredentialConfig): ImapFlow {
  const tlsBypass =
    config.imapSecurity !== "none" &&
    shouldBypassTlsVerification(config.imapHost, {
      allowPrivateHosts: env.IMAP_ALLOW_PRIVATE_HOSTS,
    });

  return new ImapFlow({
    host: config.imapHost,
    port: config.imapPort,
    secure: config.imapSecurity === "tls",
    auth: {
      user: config.username,
      pass: config.password,
    },
    logger: false,
    tls: {
      rejectUnauthorized: !tlsBypass,
    },
  });
}

export async function withImapConnection<T>(
  config: ImapCredentialConfig,
  fn: (client: ImapFlow) => Promise<T>,
): Promise<T> {
  const client = createImapConnection(config);

  try {
    await Promise.race([
      client.connect(),
      new Promise<never>((_, reject) =>
        setTimeout(
          () => reject(new Error("IMAP connection timeout")),
          CONNECT_TIMEOUT_MS,
        ),
      ),
    ]);

    return await fn(client);
  } catch (error) {
    logger.error("IMAP connection error", {
      host: config.imapHost,
      error,
    });
    throw error;
  } finally {
    try {
      await client.logout();
    } catch {
      // Ignore logout errors
    }
  }
}
