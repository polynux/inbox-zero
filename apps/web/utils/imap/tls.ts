import { normalizeHostname } from "@inboxzero/network/host-policy";

// Hosts that always resolve to this machine. Bridge binds 127.0.0.1 by
// default, so these cover every loopback spelling without a DNS lookup.
const LOOPBACK_HOSTS: Record<string, true> = {
  localhost: true,
  "127.0.0.1": true,
  "::1": true,
};

export function isLoopbackHost(host: string): boolean {
  return LOOPBACK_HOSTS[normalizeHostname(host)] === true;
}

/**
 * Whether to skip TLS certificate verification for an IMAP/SMTP connection.
 *
 * Proton Mail Bridge secures its local listener with a self-signed
 * certificate that no CA will ever sign, so connecting to Bridge with
 * verification on always fails. Skipping verification is only safe on
 * loopback, where no other machine can intercept the traffic, and only when
 * the operator explicitly opted in with IMAP_ALLOW_PRIVATE_HOSTS (which also
 * gates the host allow-list at link time). Remote hosts always verify.
 */
export function shouldBypassTlsVerification(
  host: string,
  { allowPrivateHosts }: { allowPrivateHosts: boolean },
): boolean {
  return allowPrivateHosts && isLoopbackHost(host);
}
