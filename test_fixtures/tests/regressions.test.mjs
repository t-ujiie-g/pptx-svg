import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  assert, hasTag,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
  FIXTURES_DIR, DIST_DIR,
} from './_helpers.mjs';

test("regressions (slides 96-98)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

  // ── Slide 96: empty custGeom + fontRef text color ──────────────────────
  {
    console.log('\n── test_features.pptx — Slide 96: regressions ──');
    const slide96 = textFiles.get('ppt/slides/slide96.xml');
    assert('slide96 exists', !!slide96);
    if (!slide96) {
      finishAssertions();
      return;
    }

    // Reference rect: must carry <p:style>/<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>
    // so the renderer's font_ref_color fallback gives white text.
    assert('slide96 has <p:style> on the reference rect', hasTag(slide96, 'p:style'));
    assert('slide96 has <a:fontRef> inside <p:style>', hasTag(slide96, 'a:fontRef'));
    assert(
      'slide96 fontRef targets lt1 (white text)',
      slide96.includes('<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>'),
    );
    assert(
      'slide96 reference run has no explicit color',
      slide96.includes('<a:r><a:t>Reference rect (prstGeom)</a:t></a:r>'),
    );

    // Empty custGeom: the issue #39 reproduction. The shape must be present
    // in OOXML (so the renderer exercises the suppression path) but have an
    // empty path with no draw commands.
    assert('slide96 has EmptyCustGeom shape', slide96.includes('EmptyCustGeom'));
    assert('slide96 EmptyCustGeom has solidFill 00B4D8', slide96.includes('00B4D8'));
    assert(
      'slide96 EmptyCustGeom path has no moveTo / lnTo / cubicBezTo',
      slide96.includes('<a:pathLst><a:path w="508000" h="508000"/></a:pathLst>'),
    );

    // Valid custGeom: control case — confirms the fix doesn't break normal
    // custGeom rendering.
    assert('slide96 has ValidCustGeom shape', slide96.includes('ValidCustGeom'));
    assert('slide96 ValidCustGeom has fill 06A77D', slide96.includes('06A77D'));
    assert('slide96 ValidCustGeom has a:moveTo', hasTag(slide96, 'a:moveTo'));
    assert('slide96 ValidCustGeom has a:lnTo', hasTag(slide96, 'a:lnTo'));
  }

  // ── Slide 97: header/footer field placeholders (date / footer / slide num) ──
  {
    console.log('\n── test_features.pptx — Slide 97: header/footer fields ──');
    const slide97 = textFiles.get('ppt/slides/slide97.xml');
    assert('slide97 exists', !!slide97);
    if (slide97) {
      // Date + slide-number fields and footer placeholder must be preserved in
      // OOXML (round-trip); the renderer fills the actual values at render time.
      assert('slide97 has date field', slide97.includes('type="datetime1"'));
      assert('slide97 has slide-number field', slide97.includes('type="slidenum"'));
      assert('slide97 has dt placeholder', slide97.includes('type="dt"'));
      assert('slide97 has ftr placeholder', slide97.includes('type="ftr"'));
      assert('slide97 has sldNum placeholder', slide97.includes('type="sldNum"'));
      assert('slide97 footer text preserved', slide97.includes('moon-pptx footer'));
    }
  }

  // ── Slide 98: underline styles + underline color (a:uFill) + double strike ──
  {
    console.log('\n── test_features.pptx — Slide 98: text decoration fidelity ──');
    const slide98 = textFiles.get('ppt/slides/slide98.xml');
    assert('slide98 exists', !!slide98);
    if (slide98) {
      assert('slide98 has u="dbl"', slide98.includes('u="dbl"'));
      assert('slide98 has u="wavy"', slide98.includes('u="wavy"'));
      assert('slide98 has u="dotted"', slide98.includes('u="dotted"'));
      assert('slide98 has underline color a:uFill', slide98.includes('<a:uFill>'));
      assert('slide98 has dblStrike', slide98.includes('strike="dblStrike"'));
    }
  }

  // ── Slide 99: flipH/flipV mirroring (issue #55) ─────────────────────────────
  {
    console.log('\n── test_features.pptx — Slide 99: flipH/flipV mirroring ──');
    const slide99 = textFiles.get('ppt/slides/slide99.xml');
    assert('slide99 exists', !!slide99);
    if (slide99) {
      // OOXML round-trip: the flip attributes survive in the slide XML.
      assert('slide99 has flipH', slide99.includes('flipH="1"'));
      assert('slide99 has flipV', slide99.includes('flipV="1"'));
    }
    // Rendering: the renderer must emit a negative-scale mirror for the flips
    // (the actual issue #55 bug — flips were dropped, only rotate() was kept).
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const svg99 = r.renderSlideSvg(98); // 0-based → slide 99
    assert('slide99 SVG mirrors flipH (scale(-1,1))', svg99.includes('scale(-1,1)'));
    assert('slide99 SVG mirrors flipV (scale(1,-1))', svg99.includes('scale(1,-1)'));
  }

  // ── Slide 100: preset geometries from the ECMA-376 definitions (issue #62) ─
  {
    console.log('\n── test_features.pptx — Slide 100: preset geometry (issue #62) ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const svg = r.renderSlideSvg(99);

    // Shape order on the slide: 12 presets, then 4 wedgeRectCallout variants.
    const presets = ['rect', 'donut', 'cloudCallout', 'star12', 'circularArrow',
      'uturnArrow', 'quadArrow', 'lightningBolt', 'sun', 'moon', 'cube',
      'flowChartDocument'];
    const pathOf = (idx) => {
      const start = svg.indexOf(`data-ooxml-shape-idx="${idx}"`);
      const end = svg.indexOf(`data-ooxml-shape-idx="${idx + 1}"`);
      const seg = svg.slice(start, end < 0 ? undefined : end);
      const m = seg.match(/<path d="([^"]+)"/);
      return start < 0 || !m ? '' : m[1];
    };
    // Parse M/L/A/C/Q/Z into commands with absolute end points.
    const parse = (d) => {
      const cmds = [];
      const re = /([MLACQZ])([^MLACQZ]*)/g;
      let m;
      while ((m = re.exec(d))) {
        const n = m[2].trim().split(/[\s,]+/).filter(Boolean).map(Number);
        cmds.push({ c: m[1], x: n[n.length - 2], y: n[n.length - 1] });
      }
      return cmds;
    };

    presets.forEach((prst, idx) => {
      if (prst === 'rect') return; // control shape: rendered as <rect>, not <path>
      const d = pathOf(idx);
      assert(`slide100 ${prst} renders a path`, d.length > 0);
      // An SVG arc whose end point equals its start point draws nothing
      // (the donut / cloudCallout blank-shape bug).
      let cx = 0, cy = 0, sx = 0, sy = 0, degenerate = 0;
      for (const k of parse(d)) {
        if (k.c === 'Z') { cx = sx; cy = sy; continue; }
        if (k.c === 'A' && Math.abs(k.x - cx) < 0.05 && Math.abs(k.y - cy) < 0.05) degenerate++;
        if (k.c === 'M') { sx = k.x; sy = k.y; }
        cx = k.x; cy = k.y;
      }
      assert(`slide100 ${prst} has no zero-length arcs`, degenerate === 0, `${degenerate} found`);
    });

    // star12: 24 vertices (M + 23 L).
    const star = parse(pathOf(3));
    assert('slide100 star12 has 24 vertices',
      star.filter((k) => k.c === 'M' || k.c === 'L').length === 24);

    // moon: every point stays inside its box, which starts at M r b.
    const moon = parse(pathOf(9));
    const [mr, mb] = [moon[0].x, moon[0].y];
    assert('slide100 moon stays within its box',
      moon.every((k) => k.c === 'Z' || (k.x <= mr + 0.5 && k.y <= mb + 0.5)));

    // wedgeRectCallout: vertex 0 = (l,t), vertex 8 = (r,b); the tip is the one
    // vertex outside the box and must sit on the edge nearest to it.
    const tipSlot = { W_below: 10, W_above: 2, W_left: 14, W_right: 6 };
    Object.entries(tipSlot).forEach(([name, slot], i) => {
      const v = parse(pathOf(presets.length + i)).filter((k) => k.c !== 'Z');
      if (v.length !== 16) {
        assert(`slide100 ${name} has 16 vertices`, false, `got ${v.length}`);
        return;
      }
      const [l, t, rr, b] = [v[0].x, v[0].y, v[8].x, v[8].y];
      const outside = v.findIndex((p) => p.x < l - 0.5 || p.x > rr + 0.5 || p.y < t - 0.5 || p.y > b + 0.5);
      assert(`slide100 ${name} tail leaves the matching edge`, outside === slot, `tip at vertex ${outside}`);
    });
  }

  // ── Slide 101: presets generated from presetShapeDefinitions.xml ──────────
  {
    console.log('\n── test_features.pptx — Slide 101: generated preset geometry ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const svg = r.renderSlideSvg(100);
    const shapeSvg = (idx) => {
      const start = svg.indexOf(`data-ooxml-shape-idx="${idx}"`);
      const end = svg.indexOf(`data-ooxml-shape-idx="${idx + 1}"`);
      return start < 0 ? '' : svg.slice(start, end < 0 ? undefined : end);
    };
    const presets = ['arc', 'cube', 'chartPlus', 'bentConnector3', 'actionButtonHome', 'star7'];
    presets.forEach((prst, idx) => {
      const s = shapeSvg(idx);
      // An unknown preset falls back to <rect>; every generated one is a <path>.
      assert(`slide101 ${prst} renders as a path, not the rect fallback`,
        s.includes('<path d=') && !s.includes('<rect'));
    });
    assert('slide101 arc strokes only the arc (a fill="none" path)', shapeSvg(0).includes('fill="none" stroke="rgb('));
    assert('slide101 arc wedge is filled without stroke', /fill="rgb\([^"]+\)" stroke="none"/.test(shapeSvg(0)));
    assert('slide101 cube faces are shaded', shapeSvg(1).includes('fill-opacity="0.'));
    const plus = shapeSvg(2);
    assert('slide101 chartPlus lines are drawn after the box',
      plus.lastIndexOf('fill="none"') > plus.indexOf('stroke="none"'));
    assert('slide101 bentConnector3 is stroke-only', shapeSvg(3).includes('fill="none" stroke="rgb('));
    assert('slide101 actionButtonHome icon is shaded', shapeSvg(4).includes('fill-opacity="0.'));
  }

  finishAssertions();
});
