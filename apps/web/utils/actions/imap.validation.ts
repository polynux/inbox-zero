import { z } from "zod";
import { isLoopbackHost } from "@/utils/imap/tls";

export const imapSecuritySchema = z.enum(["tls", "starttls", "none"]);

export const imapCredentialSchema = z
  .object({
    email: z.string().email(),
    name: z.string().optional(),
    imapHost: z.string().min(1),
    imapPort: z.number().int().min(1).max(65_535).default(993),
    imapSecurity: imapSecuritySchema.default("tls"),
    smtpHost: z.string().min(1),
    smtpPort: z.number().int().min(1).max(65_535).default(587),
    smtpSecurity: imapSecuritySchema.default("starttls"),
    username: z.string().min(1),
    password: z.string().min(1),
  })
  .superRefine(rejectPlaintextForRemoteHosts);

export type ImapCredentialInput = z.infer<typeof imapCredentialSchema>;

export const testImapConnectionSchema = z
  .object({
    imapHost: z.string().min(1),
    imapPort: z.number().int().min(1).max(65_535),
    imapSecurity: imapSecuritySchema,
    smtpHost: z.string().min(1),
    smtpPort: z.number().int().min(1).max(65_535),
    smtpSecurity: imapSecuritySchema,
    username: z.string().min(1),
    password: z.string().min(1),
  })
  .superRefine(rejectPlaintextForRemoteHosts);

// Plaintext only makes sense where no network hop exists (a local Bridge).
// A remote host with `none` sends credentials in the clear.
function rejectPlaintextForRemoteHosts(
  value: {
    imapHost: string;
    imapSecurity: z.infer<typeof imapSecuritySchema>;
    smtpHost: string;
    smtpSecurity: z.infer<typeof imapSecuritySchema>;
  },
  ctx: z.RefinementCtx,
) {
  const pairs = [
    {
      path: "imapSecurity",
      host: value.imapHost,
      security: value.imapSecurity,
    },
    {
      path: "smtpSecurity",
      host: value.smtpHost,
      security: value.smtpSecurity,
    },
  ] as const;

  for (const { path, host, security } of pairs) {
    if (security === "none" && !isLoopbackHost(host)) {
      ctx.addIssue({
        code: "custom",
        message: "Plaintext (none) is only allowed for a loopback host",
        path: [path],
      });
    }
  }
}
