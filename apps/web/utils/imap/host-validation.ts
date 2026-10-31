import {
  isBlockedHostname,
  isBlockedIpAddress,
  normalizeHostname,
} from "@inboxzero/network/host-policy";
import { resolveHostAddresses } from "@inboxzero/network/pinned-dns";
import { SafeError } from "@/utils/error";

type ImapHostValidationResult =
  | { valid: true }
  | { valid: false; error: string };

// Stays blocked even when private hosts are allowed.
const METADATA_HOSTS = new Set([
  "169.254.169.254",
  "fd00:ec2::254",
  "metadata.google.internal",
  "metadata.gcp.internal",
]);

/**
 * Validates user-supplied IMAP/SMTP hosts before connecting to them. Each host
 * is checked by name and by every DNS-resolved address, which also catches
 * decimal, octal and hex IPv4 forms.
 *
 * SECURITY: `allowPrivateHosts` (IMAP_ALLOW_PRIVATE_HOSTS) lets loopback and
 * private-network hosts through so a local mail bridge can be used. Metadata
 * and link-local addresses stay blocked either way. Only enable on a trusted,
 * single-tenant self-hosted deployment.
 */
export async function validateImapHosts(
  hosts: string[],
  { allowPrivateHosts }: { allowPrivateHosts: boolean },
): Promise<ImapHostValidationResult> {
  for (const host of hosts) {
    const result = await validateImapHost(host, { allowPrivateHosts });
    if (!result.valid) return result;
  }

  return { valid: true };
}

/**
 * Re-checks a stored host immediately before use. Linking validates the host
 * once, but DNS can be re-pointed afterwards (rebinding), so every connection
 * resolves and checks again instead of trusting the earlier result.
 */
export async function assertImapHostAllowed(
  host: string,
  { allowPrivateHosts }: { allowPrivateHosts: boolean },
): Promise<void> {
  const result = await validateImapHost(host, { allowPrivateHosts });
  if (!result.valid) throw new SafeError(result.error);
}

async function validateImapHost(
  host: string,
  { allowPrivateHosts }: { allowPrivateHosts: boolean },
): Promise<ImapHostValidationResult> {
  const hostname = normalizeHostname(host);
  if (!hostname || METADATA_HOSTS.has(hostname)) return hostNotAllowed();
  if (!allowPrivateHosts && isBlockedHostname(hostname)) {
    return hostNotAllowed();
  }

  let addresses: { address: string }[];
  try {
    addresses = await resolveHostAddresses(hostname);
  } catch {
    return { valid: false, error: "Host could not be resolved" };
  }

  const hasBlockedAddress = addresses.some(
    ({ address }) =>
      METADATA_HOSTS.has(address) ||
      isBlockedIpAddress(address, { allowPrivateRanges: allowPrivateHosts }),
  );
  return hasBlockedAddress ? hostNotAllowed() : { valid: true };
}

function hostNotAllowed(): ImapHostValidationResult {
  return { valid: false, error: "Host is not allowed" };
}
