import { NextResponse } from "next/server";
import { z } from "zod";
import {
  GMAIL_UPLOAD_CHUNK_BYTES,
  parseContentRange,
  validateGmailUploadChunk,
} from "@/utils/email/draft-attachment-upload";
import type { EmailProvider } from "@/utils/email/types";
import { SafeError } from "@/utils/error";
import { uploadGmailDraftChunk } from "@/utils/gmail/resumable-draft-upload";
import { withEmailProvider } from "@/utils/middleware";
import {
  appendDraftAttachmentUploadChunk,
  deleteDraftAttachmentUpload,
  getDraftAttachmentUpload,
  saveDraftAttachmentUpload,
  takeDraftAttachmentUploadChunks,
} from "@/utils/redis/draft-attachment-upload";
import { readRequestBytes } from "@/utils/request-body";

export const maxDuration = 120;

export type UploadDraftMessageChunkResponse =
  | { status: "incomplete"; nextOffset: number }
  | ({ status: "complete"; attachmentId: string } & DraftAttachmentsResult);

const paramsSchema = z.object({
  draftId: z.string().min(1),
  uploadId: z.uuid(),
});

/**
 * Forwards one chunk of a Gmail draft upload, or accumulates one chunk of
 * an IMAP draft upload until the last one arrives. The session URI and the
 * token stay here; the browser only ever names the upload it was given, and
 * the key includes the account so another account's upload id finds nothing.
 */
export const PUT = withEmailProvider(
  "user/drafts/attachments/upload-chunk",
  async (request, context) => {
    const { draftId, uploadId } = paramsSchema.parse(await context.params);
    const { emailAccountId } = request.auth;
    const upload = await getDraftAttachmentUpload(emailAccountId, uploadId);
    if (!upload || upload.draftId !== draftId)
      throw new SafeError("This upload expired. Attach the file again.", 404);

    const range = parseContentRange(request.headers.get("content-range"));
    if (!range || range.total !== upload.totalBytes)
      throw new SafeError("Upload chunk range is invalid.");
    const validation = validateGmailUploadChunk({
      start: range.start,
      length: range.length,
      totalBytes: upload.totalBytes,
      expectedStart: upload.nextOffset,
    });
    if (!validation.valid) throw new SafeError(validation.error, 409);

    const bytes = await readRequestBytes(request, GMAIL_UPLOAD_CHUNK_BYTES);
    if (!bytes || bytes.byteLength !== range.length)
      throw new SafeError("Upload chunk does not match its range.");

    // IMAP uploads have no provider session: the session URI is ours, and
    // chunks accumulate server-side until the last one completes the draft.
    if (upload.sessionUri.startsWith("imap:draft-upload:")) {
      if (!request.emailProvider.completeDraftAttachmentUpload)
        throw new SafeError("This upload expired. Attach the file again.", 404);
      await appendDraftAttachmentUploadChunk(
        emailAccountId,
        uploadId,
        Buffer.from(bytes).toString("base64"),
      );
      const nextOffset = range.start + range.length;
      if (nextOffset < upload.totalBytes) {
        await saveDraftAttachmentUpload(emailAccountId, uploadId, {
          ...upload,
          nextOffset,
        });
        return NextResponse.json({
          status: "incomplete",
          nextOffset,
        } satisfies UploadDraftMessageChunkResponse);
      }
      const chunks = await takeDraftAttachmentUploadChunks(
        emailAccountId,
        uploadId,
      );
      await request.emailProvider.completeDraftAttachmentUpload(
        draftId,
        Buffer.concat(chunks.map((chunk) => Buffer.from(chunk, "base64"))),
      );
      await deleteDraftAttachmentUpload(emailAccountId, uploadId);
      return completeUploadResponse(request.emailProvider, draftId, upload);
    }

    const result = await uploadGmailDraftChunk({
      accessToken: request.emailProvider.getAccessToken(),
      sessionUri: upload.sessionUri,
      start: range.start,
      bytes,
      totalBytes: upload.totalBytes,
    });
    if (result.status === "incomplete") {
      await saveDraftAttachmentUpload(emailAccountId, uploadId, {
        ...upload,
        nextOffset: result.nextOffset,
      });
      return NextResponse.json({
        status: "incomplete",
        nextOffset: result.nextOffset,
      } satisfies UploadDraftMessageChunkResponse);
    }

    await deleteDraftAttachmentUpload(emailAccountId, uploadId);
    return completeUploadResponse(request.emailProvider, draftId, upload);
  },
);

async function completeUploadResponse(
  emailProvider: EmailProvider,
  draftId: string,
  upload: {
    attachmentId: string;
    filename?: string;
    mimeType?: string;
    size?: number;
  },
) {
  const attachments = await emailProvider.getDraftAttachments(draftId);
  if (!attachments)
    throw new SafeError(
      "This draft is no longer available. Check Sent before trying again.",
      404,
    );
  // IMAP servers can strip the stamped attachment id (Bridge drops
  // X-Attachment-Id), so the file lists under a derived id: return the id
  // the draft actually lists, which is what removal matches on.
  const listed =
    attachments.attachments.find((item) => item.id === upload.attachmentId) ??
    (upload.filename === undefined ||
    upload.mimeType === undefined ||
    upload.size === undefined
      ? undefined
      : attachments.attachments.find(
          (item) =>
            item.filename === upload.filename &&
            item.mimeType === upload.mimeType &&
            item.size === upload.size,
        ));
  return NextResponse.json({
    status: "complete",
    attachmentId: listed?.id ?? upload.attachmentId,
    ...attachments,
  } satisfies UploadDraftMessageChunkResponse);
}
