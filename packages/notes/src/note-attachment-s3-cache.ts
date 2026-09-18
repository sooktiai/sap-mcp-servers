import { S3Client, GetObjectCommand, PutObjectCommand } from '@aws-sdk/client-s3';
import { OfficeParser, type SupportedFileType } from 'officeparser';
import { logger } from './logger.js';

export interface CachedAttachment {
  buffer: Buffer;
  contentType?: string;
  /** Markdown produced by officeparser, when the extension is one it supports; null otherwise. */
  markdown: string | null;
}

/**
 * Extensions officeparser can actually parse (its SupportedFileType union).
 * Notably excludes raster images (png/jpg/...) — its OCR mode only reads
 * images embedded inside one of these document types, or scanned PDF pages,
 * never a standalone image file.
 */
const SUPPORTED_EXTENSIONS: ReadonlySet<SupportedFileType> = new Set([
  'docx', 'pptx', 'xlsx', 'odt', 'odp', 'ods', 'odg', 'pdf', 'rtf', 'md', 'html', 'csv', 'epub'
]);

/**
 * S3-caches an attachment's raw bytes (sibling to SapNoteS3Cache, same
 * conventions) and owns its Markdown conversion, so a cache hit and a
 * fresh download both convert through the same convertToMarkdown() call.
 * Only the raw bytes are cached — Markdown is always redone on-the-fly.
 */
export class NoteAttachmentS3Cache {
  private readonly bucket: string | undefined;
  private client: S3Client | null;
  private readonly clientOverride?: Pick<S3Client, 'send'>;

  /** @param clientOverride - Injectable for tests; defaults to a real S3Client. */
  constructor(clientOverride?: Pick<S3Client, 'send'>) {
    this.bucket = process.env.SAP_NOTES_CACHE_BUCKET;
    this.client = null;
    this.clientOverride = clientOverride;
  }

  private getClient(): Pick<S3Client, 'send'> {
    if (this.clientOverride) return this.clientOverride;
    if (!this.client) this.client = new S3Client();
    return this.client;
  }

  private keyFor(noteId: string, filename: string): string {
    return `notes/${noteId}/${filename}`;
  }

  private extensionOf(filename: string): string {
    const dot = filename.lastIndexOf('.');
    return dot === -1 ? '' : filename.slice(dot + 1).toLowerCase();
  }

  /**
   * Converts an attachment's raw bytes to Markdown via officeparser, based
   * on its filename's extension. Returns null (never throws) when the
   * extension is unsupported or the conversion itself fails — the caller
   * falls back to returning the raw bytes untouched.
   *
   * The extension is passed to officeparser as an explicit `fileType` hint,
   * not left to its Buffer magic-byte auto-detection: csv/md/html have no
   * magic bytes at all, so auto-detection fails outright for exactly those
   * formats without this hint.
   */
  async convertToMarkdown(filename: string, buffer: Buffer): Promise<string | null> {
    const extension = this.extensionOf(filename);
    if (!SUPPORTED_EXTENSIONS.has(extension as SupportedFileType)) return null;

    try {
      const ast = await OfficeParser.parseOffice(buffer, { fileType: extension as SupportedFileType });
      const { value } = await ast.to('md');
      return value;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      logger.warn(`⚠️ officeparser failed to convert attachment "${filename}": ${message}`);
      return null;
    }
  }

  /**
   * Returns the cached attachment — raw bytes plus its on-the-fly Markdown
   * conversion — or null on a cache miss (including when caching is
   * disabled or S3 is unreachable — a cache-read failure must never block
   * a real download).
   */
  async get(noteId: string, filename: string): Promise<CachedAttachment | null> {
    if (!this.bucket) return null;

    try {
      const response = await this.getClient().send(
        new GetObjectCommand({ Bucket: this.bucket, Key: this.keyFor(noteId, filename) })
      );
      const bytes = await response.Body?.transformToByteArray();
      if (!bytes) return null;

      logger.info(`✅ Retrieved attachment "${filename}" for SAP Note ${noteId} from S3 cache`);
      const buffer = Buffer.from(bytes);
      const markdown = await this.convertToMarkdown(filename, buffer);
      return { buffer, contentType: response.ContentType, markdown };
    } catch (error) {
      const name = (error as { name?: string })?.name;
      if (name !== 'NoSuchKey' && name !== 'NotFound') {
        const message = error instanceof Error ? error.message : String(error);
        logger.warn(`⚠️ Attachment cache read failed for ${noteId}/${filename}, fetching normally: ${message}`);
      }
      return null;
    }
  }

  /**
   * Stores the attachment's raw bytes for future reads. Best-effort — a
   * write failure is logged and swallowed so it never fails the caller.
   *
   * Uses a conditional write (IfNoneMatch: '*') so a concurrent fetch of the
   * same not-yet-cached attachment can never overwrite whichever write
   * lands first — the loser's PreconditionFailed is treated as a normal
   * no-op, not an error.
   */
  async set(noteId: string, filename: string, buffer: Buffer, contentType?: string): Promise<void> {
    if (!this.bucket) return;

    try {
      await this.getClient().send(
        new PutObjectCommand({
          Bucket: this.bucket,
          Key: this.keyFor(noteId, filename),
          Body: buffer,
          ...(contentType ? { ContentType: contentType } : {}),
          IfNoneMatch: '*',
        })
      );
    } catch (error) {
      const name = (error as { name?: string })?.name;
      if (name === 'PreconditionFailed') return; // already cached — not an error

      const message = error instanceof Error ? error.message : String(error);
      logger.warn(`⚠️ Attachment cache write failed for ${noteId}/${filename}: ${message}`);
    }
  }
}
