import { NextResponse } from "next/server";
import { withError } from "@/utils/middleware";
import { hasCronSecret } from "@/utils/cron";
import { pollAllImapAccounts } from "@/utils/imap/poll";

/**
 * Cron endpoint to poll all IMAP accounts for new messages.
 * Should be called by an external scheduler (e.g., every 2-5 minutes).
 */
export const POST = withError("imap/poll", async (request) => {
  if (!hasCronSecret(request)) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const results = await pollAllImapAccounts();

  return NextResponse.json({
    polled: results.length,
    newMessages: results.reduce((sum, r) => sum + r.newMessages, 0),
    results,
  });
});
