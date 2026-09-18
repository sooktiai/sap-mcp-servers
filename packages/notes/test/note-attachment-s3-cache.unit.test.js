import assert from 'node:assert/strict';
import test from 'node:test';
import { NoteAttachmentS3Cache } from '../dist/note-attachment-s3-cache.js';

const NOTE_ID = '3096734';
const FILENAME = 'Pasted image.png';
const BYTES = Buffer.from('fake-png-bytes');

const ORIGINAL_BUCKET = process.env.SAP_NOTES_CACHE_BUCKET;

function withBucket(value, fn) {
  process.env.SAP_NOTES_CACHE_BUCKET = value;
  return fn().finally(() => {
    if (ORIGINAL_BUCKET === undefined) delete process.env.SAP_NOTES_CACHE_BUCKET;
    else process.env.SAP_NOTES_CACHE_BUCKET = ORIGINAL_BUCKET;
  });
}

test('get() and set() are no-ops when SAP_NOTES_CACHE_BUCKET is unset', async () => {
  delete process.env.SAP_NOTES_CACHE_BUCKET;
  const cache = new NoteAttachmentS3Cache({ send: async () => { throw new Error('S3 must not be called when caching is disabled'); } });

  assert.equal(await cache.get(NOTE_ID, FILENAME), null);
  await cache.set(NOTE_ID, FILENAME, BYTES, 'image/png');

  if (ORIGINAL_BUCKET === undefined) delete process.env.SAP_NOTES_CACHE_BUCKET;
  else process.env.SAP_NOTES_CACHE_BUCKET = ORIGINAL_BUCKET;
});

test('get() returns the cached bytes on a hit, keyed by notes/{noteId}/{filename}, with markdown null for an unsupported extension', () => withBucket('test-bucket', async () => {
  let capturedInput;
  const cache = new NoteAttachmentS3Cache({
    send: async (command) => {
      capturedInput = command.input;
      return {
        Body: { transformToByteArray: async () => new Uint8Array(BYTES) },
        ContentType: 'image/png'
      };
    }
  });

  const result = await cache.get(NOTE_ID, FILENAME);

  assert.ok(Buffer.isBuffer(result.buffer));
  assert.equal(result.buffer.toString(), BYTES.toString());
  assert.equal(result.contentType, 'image/png');
  assert.equal(result.markdown, null);
  assert.equal(capturedInput.Bucket, 'test-bucket');
  assert.equal(capturedInput.Key, 'notes/3096734/Pasted image.png');
}));

test('get() returns real Markdown on a hit for a supported extension', () => withBucket('test-bucket', async () => {
  const csv = Buffer.from('name,value\nfoo,1\nbar,2\n');
  const cache = new NoteAttachmentS3Cache({
    send: async () => ({
      Body: { transformToByteArray: async () => new Uint8Array(csv) },
      ContentType: 'text/csv'
    })
  });

  const result = await cache.get(NOTE_ID, 'data.csv');

  assert.ok(result.markdown, 'expected non-null markdown for a supported extension');
  assert.match(result.markdown, /foo/);
  assert.match(result.markdown, /bar/);
}));

test('get() returns null on NoSuchKey without throwing (cache miss)', () => withBucket('test-bucket', async () => {
  const cache = new NoteAttachmentS3Cache({
    send: async () => {
      const error = new Error('not found');
      error.name = 'NoSuchKey';
      throw error;
    }
  });

  assert.equal(await cache.get(NOTE_ID, FILENAME), null);
}));

test('get() treats an unexpected S3 error as a miss rather than failing the caller', () => withBucket('test-bucket', async () => {
  const cache = new NoteAttachmentS3Cache({ send: async () => { throw new Error('S3 unreachable'); } });

  assert.equal(await cache.get(NOTE_ID, FILENAME), null);
}));

test('set() writes a conditional PutObject (IfNoneMatch) at notes/{noteId}/{filename}', () => withBucket('test-bucket', async () => {
  let capturedInput;
  const cache = new NoteAttachmentS3Cache({
    send: async (command) => {
      capturedInput = command.input;
      return {};
    }
  });

  await cache.set(NOTE_ID, FILENAME, BYTES, 'image/png');

  assert.equal(capturedInput.Bucket, 'test-bucket');
  assert.equal(capturedInput.Key, 'notes/3096734/Pasted image.png');
  assert.equal(capturedInput.IfNoneMatch, '*');
  assert.equal(capturedInput.Body, BYTES);
  assert.equal(capturedInput.ContentType, 'image/png');
}));

test('set() omits ContentType when none is given', () => withBucket('test-bucket', async () => {
  let capturedInput;
  const cache = new NoteAttachmentS3Cache({
    send: async (command) => {
      capturedInput = command.input;
      return {};
    }
  });

  await cache.set(NOTE_ID, FILENAME, BYTES);

  assert.equal('ContentType' in capturedInput, false);
}));

test('set() silently no-ops when the attachment is already cached (PreconditionFailed)', () => withBucket('test-bucket', async () => {
  const cache = new NoteAttachmentS3Cache({
    send: async () => {
      const error = new Error('At least one of the pre-conditions you specified did not hold');
      error.name = 'PreconditionFailed';
      throw error;
    }
  });

  await assert.doesNotReject(cache.set(NOTE_ID, FILENAME, BYTES));
}));

test('set() swallows an unexpected write failure rather than throwing', () => withBucket('test-bucket', async () => {
  const cache = new NoteAttachmentS3Cache({ send: async () => { throw new Error('S3 unreachable'); } });

  await assert.doesNotReject(cache.set(NOTE_ID, FILENAME, BYTES));
}));

// ─── convertToMarkdown() ────────────────────────────────────────────────
// No mocking needed: these feed real tiny buffers through the actual
// officeparser call (fast, deterministic).

test('convertToMarkdown() converts a supported extension (csv) to Markdown', async () => {
  const cache = new NoteAttachmentS3Cache();
  const csv = Buffer.from('name,value\nfoo,1\nbar,2\n');

  const markdown = await cache.convertToMarkdown('data.csv', csv);

  assert.ok(markdown, 'expected non-null markdown for a supported extension');
  assert.match(markdown, /foo/);
  assert.match(markdown, /bar/);
});

test('convertToMarkdown() converts a supported extension (html) to Markdown', async () => {
  const cache = new NoteAttachmentS3Cache();
  const html = Buffer.from('<html><body><h1>Title</h1><p>Hello world</p></body></html>');

  const markdown = await cache.convertToMarkdown('page.html', html);

  assert.ok(markdown, 'expected non-null markdown for a supported extension');
  assert.match(markdown, /Title/);
  assert.match(markdown, /Hello world/);
});

test('convertToMarkdown() returns null for an unsupported extension (png) without attempting to parse', async () => {
  const cache = new NoteAttachmentS3Cache();
  // Not valid PNG bytes — if this ever reached officeparser it would throw,
  // not return markdown. A null result here proves the extension check
  // short-circuited before any parse attempt.
  const notActuallyAPng = Buffer.from('definitely not png bytes');

  const markdown = await cache.convertToMarkdown('Pasted image.png', notActuallyAPng);

  assert.equal(markdown, null);
});

test('convertToMarkdown() returns null for an extension officeparser has no route for at all (zip)', async () => {
  const cache = new NoteAttachmentS3Cache();
  const bogus = Buffer.from('PK\x03\x04-fake-zip-bytes');

  const markdown = await cache.convertToMarkdown('archive.zip', bogus);

  assert.equal(markdown, null);
});

test('convertToMarkdown() returns null (rather than throwing) when the bytes are corrupt for a supported extension', async () => {
  const cache = new NoteAttachmentS3Cache();
  const corruptPdf = Buffer.from('this is not a real pdf');

  const markdown = await cache.convertToMarkdown('steps.pdf', corruptPdf);

  assert.equal(markdown, null);
});
