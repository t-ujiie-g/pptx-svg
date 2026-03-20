/**
 * Unit tests for editing-helpers.ts (unit conversion functions).
 *
 * Run: node test_fixtures/test_editing_helpers.mjs
 */

import {
  EMU_PER_PT, EMU_PER_PX_96DPI,
  pxToEmu, emuToPx, ptToHundredths, hundredthsToPt,
  degreesToOoxml, ooxmlToDegrees,
} from '../dist/editing-helpers.js';

let passed = 0;
let failed = 0;

function assert(name, condition) {
  if (condition) {
    console.log(`  ✓ ${name}`);
    passed++;
  } else {
    console.log(`  ✗ ${name}`);
    failed++;
  }
}

console.log('=== Editing Helpers Tests ===\n');

// ── Constants ──
console.log('── Constants ──');
assert('EMU_PER_PT = 12700', EMU_PER_PT === 12700);
assert('EMU_PER_PX_96DPI = 9525', EMU_PER_PX_96DPI === 9525);

// ── pxToEmu / emuToPx ──
console.log('\n── pxToEmu / emuToPx ──');
assert('pxToEmu(1) = 9525', pxToEmu(1) === 9525);
assert('pxToEmu(0) = 0', pxToEmu(0) === 0);
assert('pxToEmu(96) = 914400', pxToEmu(96) === 914400);
assert('emuToPx(914400) = 96', emuToPx(914400) === 96);
assert('emuToPx(9525) = 1', emuToPx(9525) === 1);
assert('emuToPx(0) = 0', emuToPx(0) === 0);
assert('pxToEmu roundtrip: 100px', emuToPx(pxToEmu(100)) === 100);
assert('pxToEmu roundtrip: 960px', emuToPx(pxToEmu(960)) === 960);

// Custom DPI
assert('pxToEmu(1, 72) = 12700', pxToEmu(1, 72) === 12700);
assert('emuToPx(914400, 72) = 72', emuToPx(914400, 72) === 72);

// ── ptToHundredths / hundredthsToPt ──
console.log('\n── ptToHundredths / hundredthsToPt ──');
assert('ptToHundredths(18) = 1800', ptToHundredths(18) === 1800);
assert('ptToHundredths(0) = 0', ptToHundredths(0) === 0);
assert('ptToHundredths(10.5) = 1050', ptToHundredths(10.5) === 1050);
assert('hundredthsToPt(1800) = 18', hundredthsToPt(1800) === 18);
assert('hundredthsToPt(0) = 0', hundredthsToPt(0) === 0);
assert('hundredthsToPt(1050) = 10.5', hundredthsToPt(1050) === 10.5);

// ── degreesToOoxml / ooxmlToDegrees ──
console.log('\n── degreesToOoxml / ooxmlToDegrees ──');
assert('degreesToOoxml(90) = 5400000', degreesToOoxml(90) === 5400000);
assert('degreesToOoxml(0) = 0', degreesToOoxml(0) === 0);
assert('degreesToOoxml(360) = 21600000', degreesToOoxml(360) === 21600000);
assert('degreesToOoxml(45) = 2700000', degreesToOoxml(45) === 2700000);
assert('ooxmlToDegrees(5400000) = 90', ooxmlToDegrees(5400000) === 90);
assert('ooxmlToDegrees(0) = 0', ooxmlToDegrees(0) === 0);
assert('ooxmlToDegrees(21600000) = 360', ooxmlToDegrees(21600000) === 360);

// ── Summary ──
console.log(`\nTotal: ${passed + failed}  Passed: ${passed}  Failed: ${failed}`);
if (failed > 0) {
  console.log('SOME TESTS FAILED');
  process.exit(1);
} else {
  console.log('All tests passed.');
}
