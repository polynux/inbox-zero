import { redis } from "@/utils/redis";

// Long enough to upload 15 MB on a slow connection; Gmail keeps the session
// for about a week, so ours always expires first.
const UPLOAD_TTL_SECONDS = 60 * 60;

type DraftAttachmentUpload = {
  draftId: string;
  attachmentId: string;
  sessionUri: string;
  totalBytes: number;
  nextOffset: number;
};

// IMAP has no resumable-upload endpoint: chunks accumulate here until the
// last one arrives, then the provider APPENDs the assembled message. Stored
// beside the upload record under the same TTL.
type DraftAttachmentUploadChunks = {
  chunks: string[];
};

export async function saveDraftAttachmentUpload(
  emailAccountId: string,
  uploadId: string,
  upload: DraftAttachmentUpload,
) {
  await redis.set(uploadKey(emailAccountId, uploadId), upload, {
    ex: UPLOAD_TTL_SECONDS,
  });
}

export async function getDraftAttachmentUpload(
  emailAccountId: string,
  uploadId: string,
) {
  return redis.get<DraftAttachmentUpload>(uploadKey(emailAccountId, uploadId));
}

export async function deleteDraftAttachmentUpload(
  emailAccountId: string,
  uploadId: string,
) {
  await redis.del(uploadKey(emailAccountId, uploadId));
}

// The account is part of the key, so another account's upload id finds nothing.
function uploadKey(emailAccountId: string, uploadId: string) {
  return `draft-attachment-upload:${emailAccountId}:${uploadId}`;
}

export async function appendDraftAttachmentUploadChunk(
  emailAccountId: string,
  uploadId: string,
  chunkBase64: string,
) {
  const key = uploadChunksKey(emailAccountId, uploadId);
  const stored = await redis.get<DraftAttachmentUploadChunks>(key);
  await redis.set(
    key,
    { chunks: [...(stored?.chunks ?? []), chunkBase64] },
    { ex: UPLOAD_TTL_SECONDS },
  );
}

export async function takeDraftAttachmentUploadChunks(
  emailAccountId: string,
  uploadId: string,
) {
  const key = uploadChunksKey(emailAccountId, uploadId);
  const stored = await redis.get<DraftAttachmentUploadChunks>(key);
  await redis.del(key);
  return stored?.chunks ?? [];
}

function uploadChunksKey(emailAccountId: string, uploadId: string) {
  return `draft-attachment-upload:${emailAccountId}:${uploadId}:chunks`;
}
