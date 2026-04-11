import { test } from 'node:test';
import { expect, assert, section, resetAssertions, finishAssertions } from './_helpers.mjs';

test('utility functions (bytesToBase64, crc32)', async () => {
  resetAssertions();
  const { bytesToBase64, crc32 } = await import('../../dist/utils.js');

  section('Utility functions — bytesToBase64');
  assert('undefined → ""', bytesToBase64(undefined) === '');
  assert('empty → ""', bytesToBase64(new Uint8Array(0)) === '');
  assert('Hello → SGVsbG8=', bytesToBase64(new TextEncoder().encode('Hello')) === 'SGVsbG8=');
  assert('single byte 0xFF', bytesToBase64(new Uint8Array([0xFF])) === '/w==');

  section('Utility functions — crc32');
  const crc1 = crc32(new Uint8Array(0));
  assert('empty data CRC', crc1 === 0x00000000);
  const crc2 = crc32(new TextEncoder().encode('123456789'));
  assert('CRC-32 of "123456789" = 0xCBF43926', crc2 === 0xCBF43926);
  finishAssertions();
});
