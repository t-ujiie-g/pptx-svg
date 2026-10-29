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

    // Empty custGeom (#39): present in the OOXML, with a path that has no commands.
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

  {
    console.log('\n── test_features.pptx — Slide 102: custGeom per-path styles ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const svg = r.renderSlideSvg(101);
    const paths = [...svg.matchAll(/<path d="([^"]+)"([^>]*)\/>/g)].map((m) => ({ d: m[1], attrs: m[2] }));
    const fillOnly = paths.filter((p) => /fill="rgb\(91,155,213\)" stroke="none"/.test(p.attrs));
    const strokeOnly = paths.filter((p) => /fill="none" stroke="rgb\(255,192,0\)"/.test(p.attrs));
    const shade = paths.filter((p) => p.attrs.includes('fill-opacity="0.4"'));
    // The square, and the darken path's base fill (its overlay is separate).
    assert('slide102 fill-only paths are drawn without stroke', fillOnly.length === 2, `got ${fillOnly.length}`);
    assert('slide102 diagonal is stroke-only', strokeOnly.length === 1);
    assert('slide102 darken path gets a shade overlay', shade.length === 1);
    // The diagonal lives in a 50×50 space: it must span the whole shape (as wide
    // as the 100×100 square), not half of it as the first path's space would give.
    if (strokeOnly.length === 1 && fillOnly.length > 0) {
      const nums = (d) => d.match(/-?[\d.]+/g).map(Number);
      const sq = nums(fillOnly[0].d); // M x0 y0 L x1 y0 ...
      const dg = nums(strokeOnly[0].d);
      assert('slide102 diagonal uses its own coordinate space',
        Math.abs((dg[2] - dg[0]) - (sq[2] - sq[0])) < 1, `${strokeOnly[0].d} vs ${fillOnly[0].d}`);
    }
    // Round-trip SVG → SlideData → OOXML keeps the per-path attributes. (An
    // untouched slide exports its original XML, so go through the SVG.)
    r.updateSlideFromSvg(101, svg);
    const xml = r.getSlideOoxml(101);
    assert('slide102 export keeps fill="none"', xml.includes('<a:path w="50" h="50" fill="none">'));
    assert('slide102 export keeps fill="darken" + stroke', xml.includes('fill="darken" stroke="0"'));
    assert('slide102 export keeps stroke="0"', xml.includes('<a:path w="100" h="100" stroke="0">'));
  }

  {
    console.log('\n── test_features.pptx — layout blipFill background inheritance ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const { buildZip } = await import(join(DIST_DIR, 'zip.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    // Slide 1 has no <p:bg> and uses slideLayout7; give that layout an image
    // background whose rId exists only in the layout's .rels.
    const layoutPath = 'ppt/slideLayouts/slideLayout7.xml';
    const layoutRelsPath = 'ppt/slideLayouts/_rels/slideLayout7.xml.rels';
    const layoutXml = textFiles.get(layoutPath).replace(/<p:cSld([^>]*)>/,
      '<p:cSld$1><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdLayoutBg"/>' +
      '<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>');
    const layoutRels = textFiles.get(layoutRelsPath).replace('</Relationships>',
      '<Relationship Id="rIdLayoutBg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
      'Target="../media/image1.png"/></Relationships>');
    assert('slide1 has no own background', !textFiles.get('ppt/slides/slide1.xml').includes('<p:bg>'));
    const patched = await buildZip(pptxAb, new Map([[layoutPath, layoutXml], [layoutRelsPath, layoutRels]]));
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(patched);
    const hasBgImage = (svg) => /<image x="0" y="0"[^>]*href="data:image\/png;base64,/.test(svg);
    assert('inherited layout image background is rendered', hasBgImage(r.renderSlideSvg(0)));
    // An edited slide renders from the SlideData cache; the background rels must survive.
    r.updateShapeFill(0, 0, 255, 0, 0);
    assert('inherited layout image background survives an edit', hasBgImage(r.renderSlideSvg(0)));
    // Export must not bake the inherited background (and the layout's rId) into the slide.
    assert('edited slide export has no <p:bg>', !r.getSlideOoxml(0).includes('<p:bg>'));
    // Undo/Redo re-parses the serialized slide XML, which re-inherits the background.
    r.undo();
    r.redo();
    assert('inherited layout image background survives undo/redo', hasBgImage(r.renderSlideSvg(0)));
    // SVG round-trip carries the inherited flag (data-ooxml-bg-inherited).
    const svg = r.renderSlideSvg(0);
    assert('SVG marks the background as inherited', svg.includes('data-ooxml-bg-inherited="1"'));
    r.updateSlideFromSvg(0, svg);
    assert('SVG round-trip export has no <p:bg>', !r.getSlideOoxml(0).includes('<p:bg>'));
  }

  {
    console.log('\n── test_features.pptx — Slides 103-105: background SVG round-trip ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const bgOf = (xml) => (xml.match(/<p:bg>.*?<\/p:bg>/s) || [''])[0];
    const cases = [
      { idx: 102, name: 'gradient', attr: 'data-ooxml-bg-grad-stops="0,1F4E79;100000,9DC3E6"',
        svgMark: 'url(#grad-bg-s103)', xml: ['<a:gradFill', '1F4E79', '9DC3E6'] },
      { idx: 103, name: 'pattern', attr: 'data-ooxml-bg-patt-prst="dkDnDiag"',
        svgMark: 'patt-bg-s104', xml: ['<a:pattFill prst="dkDnDiag"', 'C00000', 'FFF2CC'] },
      { idx: 104, name: 'image', attr: 'data-ooxml-bg-blip-rid="rId',
        svgMark: 'href="data:image/png;base64,', xml: ['<a:blipFill', 'r:embed="rId'] },
    ];
    for (const c of cases) {
      const svg = r.renderSlideSvg(c.idx);
      assert(`slide${c.idx + 1} SVG carries the ${c.name} background attrs`, svg.includes(c.attr));
      assert(`slide${c.idx + 1} renders the ${c.name} background`, svg.includes(c.svgMark));
      r.updateSlideFromSvg(c.idx, svg);
      const bg = bgOf(r.getSlideOoxml(c.idx));
      for (const frag of c.xml) {
        assert(`slide${c.idx + 1} round-trip export keeps ${c.name} <p:bg> (${frag})`, bg.includes(frag), bg);
      }
      assert(`slide${c.idx + 1} re-renders the ${c.name} background after round-trip`,
        r.renderSlideSvg(c.idx).includes(c.svgMark));
    }
  }

  {
    console.log('\n── test_features.pptx — chart / OLE graphicFrame survives editing ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const frameOf = (xml) => (xml.match(/<p:graphicFrame>.*?<\/p:graphicFrame>/s) || [''])[0];
    const xfrmOf = (frame) => (frame.match(/<p:xfrm[^>]*>.*?<\/p:xfrm>/s) || [''])[0];
    const chartBars = (svg) => (svg.match(/data-ooxml-shape-type="chart"[\s\S]*?<\/g><\/g>/) || [''])[0]
      .split('<rect').length - 1;

    // Slide 52: a classic chart (shape 0) next to a text box.
    const idx = 51;
    const barsBefore = chartBars(r.renderSlideSvg(idx));
    assert('slide52 chart renders bars', barsBefore > 2, `${barsBefore}`);
    // Editing only the other shape must not drop the chart.
    r.updateShapeTransform(idx, 1, 914400, 5029200, 3657600, 457200, 0);
    assert('slide52 export keeps the chart after editing another shape',
      frameOf(r.getSlideOoxml(idx)).includes('<c:chart '));
    // Moving the chart itself is written into its p:xfrm.
    r.updateShapeTransform(idx, 0, 457200, 228600, 4572000, 2743200, 0);
    const moved = xfrmOf(frameOf(r.getSlideOoxml(idx)));
    assert('slide52 moved chart exports its new offset', moved.includes('<a:off x="457200" y="228600"/>'), moved);
    assert('slide52 moved chart exports its new size', moved.includes('<a:ext cx="4572000" cy="2743200"/>'), moved);
    // Undo/Redo re-parses the serialized slide: the chart must come back with its data.
    r.undo();
    r.redo();
    assert('slide52 chart renders after undo/redo', chartBars(r.renderSlideSvg(idx)) === barsBefore);
    // SVG round-trip: the frame travels as data-ooxml-ole and the chart is re-resolved.
    r.updateSlideFromSvg(idx, r.renderSlideSvg(idx));
    assert('slide52 SVG round-trip export keeps the chart', frameOf(r.getSlideOoxml(idx)).includes('<c:chart '));
    assert('slide52 chart renders after SVG round-trip', chartBars(r.renderSlideSvg(idx)) === barsBefore);
    // Charts stay non-copyable: their chart part cannot be carried by a paste.
    assert('chart copy is refused', r.getShapeSpec(idx, 0).startsWith('ERROR'));

    // Slide 90: an Office 2016+ chart (cx:chart).
    r.renderSlideSvg(89);
    r.updateShapeTransform(89, 0, 914400, 914400, 4572000, 2743200, 0);
    assert('slide90 export keeps the ChartEx frame', r.getSlideOoxml(89).includes('<cx:chart '));

    // Slide 77: an OLE object — its move is exported too (the frame used to be written verbatim).
    r.renderSlideSvg(76);
    r.updateShapeTransform(76, 0, 1828800, 914400, 2743200, 1828800, 0);
    const ole = frameOf(r.getSlideOoxml(76));
    assert('slide77 export keeps the OLE object', ole.includes('<p:oleObj'));
    assert('slide77 moved OLE exports its new offset', xfrmOf(ole).includes('<a:off x="1828800" y="914400"/>'), xfrmOf(ole));
  }

  // ── Slide 106: shape ids, names and alt text survive editing ─────────────
  {
    console.log('\n── test_features.pptx — Slide 106: shape ids / names / alt text ──');
    const { PptxRenderer } = await import(join(DIST_DIR, 'index.js'));
    const wasmBuf = readFileSync(join(DIST_DIR, 'main.wasm'));
    const buf = readFileSync(join(FIXTURES_DIR, 'test_features.pptx'));
    const pptxAb = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const r = new PptxRenderer({ logLevel: 'silent' });
    await r.init(wasmBuf);
    await r.loadPptx(pptxAb);
    const idx = 105;
    const cnvPrs = (xml) => [...xml.matchAll(/<p:cNvPr ([^>]*?)\/?>/g)].map((m) => m[1]);
    const ids = (xml) => cnvPrs(xml).map((a) => a.match(/id="(\d+)"/)[1]);
    const spid = r.getSlideOoxml(idx).match(/spid="(\d+)"/)[1];

    const svg = r.renderSlideSvg(idx);
    assert('slide106 SVG carries the alt text as <title>',
      svg.includes('<title>Quarterly revenue, up 12% &amp; rising</title>') && svg.includes('<title>Company logo</title>'));
    assert('slide106 SVG carries the shape name', svg.includes('data-ooxml-nv-name="Revenue Box"'));

    r.updateShapeFill(idx, 0, 255, 0, 0);
    const edited = r.getSlideOoxml(idx);
    const attrs = cnvPrs(edited).join(' | ');
    assert('slide106 export keeps the animation', edited.includes(`spid="${spid}"`));
    assert('slide106 export keeps the animated shape\'s id', ids(edited).includes(spid), attrs);
    assert('slide106 export keeps names and alt text',
      attrs.includes('name="Revenue Box" descr="Quarterly revenue, up 12% &amp; rising" title="Revenue"') &&
      attrs.includes('name="Logo" descr="Company logo"'), attrs);

    // A duplicate gets a fresh id; the original keeps the one the animation targets.
    r.duplicateShape(idx, 0, 91440, 91440);
    const dup = r.getSlideOoxml(idx);
    assert('slide106 ids stay unique after duplicating', new Set(ids(dup)).size === ids(dup).length, ids(dup).join(','));
    assert('slide106 original keeps its id after duplicating', cnvPrs(dup)[1].includes(`id="${spid}"`), cnvPrs(dup)[1]);

    // SVG round-trip keeps them too.
    r.updateSlideFromSvg(idx, r.renderSlideSvg(idx));
    assert('slide106 SVG round-trip keeps ids, names and alt text',
      cnvPrs(r.getSlideOoxml(idx)).join(' | ').includes(`id="${spid}" name="Revenue Box" descr=`));
  }

  finishAssertions();
});
