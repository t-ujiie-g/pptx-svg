/**
 * ZIP extraction and building for PPTX files.
 *
 * Uses browser-native DecompressionStream/CompressionStream for DEFLATE.
 */

import { crc32 } from './utils.js';

/** Return true if a ZIP entry should be decoded as UTF-8 text. */
function isTextEntry(name: string): boolean {
  const lower = name.toLowerCase();
  return lower.endsWith('.xml')  ||
         lower.endsWith('.rels') ||
         lower.endsWith('.txt')  ||
         lower.endsWith('.json') ||
         lower.endsWith('.html') ||
         lower.endsWith('.css')  ||
         lower === '[content_types].xml';
}

/** Decompress raw DEFLATE bytes using the browser's DecompressionStream. */
async function inflate(compressed: Uint8Array): Promise<Uint8Array> {
  const stream = new DecompressionStream('deflate-raw');
  const writer = stream.writable.getWriter();
  const reader = stream.readable.getReader();

  writer.write(compressed as BufferSource);
  writer.close();

  const chunks: Uint8Array[] = [];
  let totalLen = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    totalLen += value.length;
  }

  const result = new Uint8Array(totalLen);
  let pos = 0;
  for (const chunk of chunks) { result.set(chunk, pos); pos += chunk.length; }
  return result;
}

/** Compress bytes using DEFLATE-raw via CompressionStream. */
async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new CompressionStream('deflate-raw');
  const writer = stream.writable.getWriter();
  const reader = stream.readable.getReader();

  writer.write(data as BufferSource);
  writer.close();

  const chunks: Uint8Array[] = [];
  let totalLen = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    totalLen += value.length;
  }

  const result = new Uint8Array(totalLen);
  let pos = 0;
  for (const chunk of chunks) { result.set(chunk, pos); pos += chunk.length; }
  return result;
}

/** Result of extracting a ZIP archive. */
export interface ZipContents {
  textFiles: Map<string, string>;
  binaryFiles: Map<string, Uint8Array>;
}

/** Find the End-of-Central-Directory record. Scans backwards from the end. */
function findEocd(view: DataView): number {
  // EOCD is at least 22 bytes; comment can be up to 65535 bytes.
  const minPos = Math.max(0, view.byteLength - 22 - 65535);
  for (let i = view.byteLength - 22; i >= minPos; i--) {
    if (view.getUint32(i, true) === 0x06054b50) return i;
  }
  return -1;
}

/** Central Directory entry info (sizes and CRC are always reliable here). */
interface CdEntry {
  method: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

/** Parse the Central Directory to get reliable sizes for each entry. */
function parseCentralDirectory(view: DataView, bytes: Uint8Array): Map<string, CdEntry> {
  const decoder = new TextDecoder('utf-8');
  const entries = new Map<string, CdEntry>();
  const eocdOff = findEocd(view);
  if (eocdOff < 0) return entries;

  const cdOffset = view.getUint32(eocdOff + 16, true);
  const cdSize = view.getUint32(eocdOff + 12, true);
  let off = cdOffset;
  const cdEnd = cdOffset + cdSize;

  while (off < cdEnd && off + 46 <= view.byteLength) {
    if (view.getUint32(off, true) !== 0x02014b50) break;
    const method = view.getUint16(off + 10, true);
    const entryCrc32 = view.getUint32(off + 16, true);
    const compressedSize = view.getUint32(off + 20, true);
    const uncompressedSize = view.getUint32(off + 24, true);
    const nameLen = view.getUint16(off + 28, true);
    const extraLen = view.getUint16(off + 30, true);
    const commentLen = view.getUint16(off + 32, true);
    const localHeaderOffset = view.getUint32(off + 42, true);
    const name = decoder.decode(bytes.slice(off + 46, off + 46 + nameLen));
    entries.set(name, { method, crc32: entryCrc32, compressedSize, uncompressedSize, localHeaderOffset });
    off += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/**
 * Extract all entries from a ZIP archive.
 *
 * Uses the Central Directory for reliable entry sizes, which handles
 * ZIP files with data descriptors (flag bit 3) where local header
 * sizes are set to 0 (e.g. files produced by Google Slides).
 */
export async function extractZip(
  buffer: ArrayBuffer,
  log?: { warn(...args: unknown[]): void },
): Promise<ZipContents> {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(buffer);
  const textFiles = new Map<string, string>();
  const binaryFiles = new Map<string, Uint8Array>();
  const decoder = new TextDecoder('utf-8');

  // Parse Central Directory for reliable sizes
  const cdEntries = parseCentralDirectory(view, bytes);

  // Walk local file headers using CD info for sizes
  for (const [name, cd] of cdEntries) {
    const offset = cd.localHeaderOffset;
    if (offset + 30 > bytes.length) continue;
    if (view.getUint32(offset, true) !== 0x04034b50) continue;

    const fileNameLen = view.getUint16(offset + 26, true);
    const extraLen    = view.getUint16(offset + 28, true);
    const dataOffset  = offset + 30 + fileNameLen + extraLen;

    // Use sizes from Central Directory (always reliable)
    const compressed = bytes.slice(dataOffset, dataOffset + cd.compressedSize);

    let decompressed: Uint8Array;
    if (cd.method === 0) {
      decompressed = compressed;
    } else if (cd.method === 8) {
      decompressed = await inflate(compressed);
    } else {
      log?.warn(`Unsupported compression method ${cd.method} for ${name}, skipping`);
      continue;
    }

    if (isTextEntry(name)) {
      textFiles.set(name, decoder.decode(decompressed));
    } else {
      binaryFiles.set(name, decompressed);
    }
  }

  return { textFiles, binaryFiles };
}

interface ZipEntry {
  name: string;
  method: number;
  flags: number;
  time: number;
  date: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  compressedData: Uint8Array;
  extra: Uint8Array;
}

/**
 * Build a new ZIP by iterating the original ZIP entries and replacing
 * modified text entries with new content.
 *
 * @param originalBuffer - The original PPTX bytes
 * @param modifications - path → new XML content
 * @returns Rebuilt ZIP as ArrayBuffer
 */
export async function buildZip(
  originalBuffer: ArrayBuffer,
  modifications: Map<string, string>,
): Promise<ArrayBuffer> {
  const origBytes = new Uint8Array(originalBuffer);
  const origView = new DataView(originalBuffer);
  const decoder = new TextDecoder('utf-8');
  const encoder = new TextEncoder();

  // Parse Central Directory for reliable sizes (handles data descriptor flag)
  const cdEntries = parseCentralDirectory(origView, origBytes);

  // Collect all entries using CD info
  const entries: ZipEntry[] = [];
  for (const [name, cd] of cdEntries) {
    const offset = cd.localHeaderOffset;
    if (offset + 30 > origBytes.length) continue;
    if (origView.getUint32(offset, true) !== 0x04034b50) continue;

    const fileNameLen = origView.getUint16(offset + 26, true);
    const extraLen    = origView.getUint16(offset + 28, true);
    const dataOffset  = offset + 30 + fileNameLen + extraLen;

    const flags = origView.getUint16(offset + 6, true) & ~0x08; // clear DD flag for rebuilt ZIP
    const time  = origView.getUint16(offset + 10, true);
    const date  = origView.getUint16(offset + 12, true);

    entries.push({
      name, method: cd.method, flags, time, date,
      crc32: cd.crc32,
      compressedSize: cd.compressedSize,
      uncompressedSize: cd.uncompressedSize,
      compressedData: origBytes.slice(dataOffset, dataOffset + cd.compressedSize),
      extra: origBytes.slice(offset + 30 + fileNameLen, dataOffset),
    });
  }

  // Process modifications: replace entry data
  for (const entry of entries) {
    if (modifications.has(entry.name)) {
      const newContent = encoder.encode(modifications.get(entry.name)!);
      const compressed = await deflate(newContent);
      entry.method = 8;
      entry.compressedData = compressed;
      entry.compressedSize = compressed.length;
      entry.uncompressedSize = newContent.length;
      entry.crc32 = crc32(newContent);
      entry.extra = new Uint8Array(0);
    }
  }

  // Build the new ZIP
  const parts: Uint8Array[] = [];
  const centralDir: Uint8Array[] = [];
  let localOffset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);

    // Local file header (30 bytes + name + extra + data)
    const localHeader = new ArrayBuffer(30);
    const lhView = new DataView(localHeader);
    lhView.setUint32(0, 0x04034b50, true);   // signature
    lhView.setUint16(4, 20, true);            // version needed
    lhView.setUint16(6, entry.flags, true);   // flags
    lhView.setUint16(8, entry.method, true);  // method
    lhView.setUint16(10, entry.time, true);   // time
    lhView.setUint16(12, entry.date, true);   // date
    lhView.setUint32(14, entry.crc32, true);  // crc32
    lhView.setUint32(18, entry.compressedSize, true);
    lhView.setUint32(22, entry.uncompressedSize, true);
    lhView.setUint16(26, nameBytes.length, true);
    lhView.setUint16(28, entry.extra.length, true);

    parts.push(new Uint8Array(localHeader));
    parts.push(nameBytes);
    parts.push(entry.extra);
    parts.push(entry.compressedData);

    // Central directory entry (46 bytes + name)
    const cdHeader = new ArrayBuffer(46);
    const cdView = new DataView(cdHeader);
    cdView.setUint32(0, 0x02014b50, true);    // signature
    cdView.setUint16(4, 20, true);             // version made by
    cdView.setUint16(6, 20, true);             // version needed
    cdView.setUint16(8, entry.flags, true);
    cdView.setUint16(10, entry.method, true);
    cdView.setUint16(12, entry.time, true);
    cdView.setUint16(14, entry.date, true);
    cdView.setUint32(16, entry.crc32, true);
    cdView.setUint32(20, entry.compressedSize, true);
    cdView.setUint32(24, entry.uncompressedSize, true);
    cdView.setUint16(28, nameBytes.length, true);
    cdView.setUint16(30, 0, true);             // extra length
    cdView.setUint16(32, 0, true);             // comment length
    cdView.setUint16(34, 0, true);             // disk number
    cdView.setUint16(36, 0, true);             // internal attrs
    cdView.setUint32(38, 0, true);             // external attrs
    cdView.setUint32(42, localOffset, true);   // local header offset

    centralDir.push(new Uint8Array(cdHeader));
    centralDir.push(nameBytes);

    localOffset += 30 + nameBytes.length + entry.extra.length + entry.compressedData.length;
  }

  const cdOffset = localOffset;
  let cdSize = 0;
  for (const part of centralDir) cdSize += part.length;

  // End of central directory (22 bytes)
  const eocd = new ArrayBuffer(22);
  const eocdView = new DataView(eocd);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  eocdView.setUint16(8, entries.length, true);
  eocdView.setUint16(10, entries.length, true);
  eocdView.setUint32(12, cdSize, true);
  eocdView.setUint32(16, cdOffset, true);
  eocdView.setUint16(20, 0, true);

  // Combine all parts
  let totalSize = 0;
  for (const p of parts) totalSize += p.length;
  totalSize += cdSize + 22;

  const result = new Uint8Array(totalSize);
  let pos = 0;
  for (const p of parts) { result.set(p, pos); pos += p.length; }
  for (const p of centralDir) { result.set(p, pos); pos += p.length; }
  result.set(new Uint8Array(eocd), pos);

  return result.buffer;
}
