/**
 * Utility functions for pptx-render.
 */

/** Encode a Uint8Array as a base64 string. */
export function bytesToBase64(bytes: Uint8Array | undefined): string {
  if (!bytes) return '';
  let binary = '';
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}

/** Compute CRC-32 for a Uint8Array. */
export function crc32(data: Uint8Array): number {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
