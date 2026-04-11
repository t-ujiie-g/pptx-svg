import { test } from 'node:test';
import {
  expect, assert, section, hasTag, findRelTarget, countSlideIds,
  loadFeatures, pptxExists, resetAssertions, finishAssertions,
} from './_helpers.mjs';

test("charts (slides 52-69, 73, 89-94)", async () => {
  resetAssertions();
  if (!pptxExists('test_features.pptx')) {
    console.log('  SKIPPED: test_features.pptx not found');
    return;
  }
  const { textFiles } = await loadFeatures();

  // ── Slide 52: Column chart ──
  {
    const slide52 = textFiles.get('ppt/slides/slide52.xml') || '';
    const rels52 = textFiles.get('ppt/slides/_rels/slide52.xml.rels') || '';
    section('test_features.pptx — Slide 52: column chart');
    assert('slide52 has p:graphicFrame', slide52.includes('p:graphicFrame'));
    assert('slide52 has a:graphicData', slide52.includes('a:graphicData'));
    assert('slide52 has c:chart', slide52.includes('c:chart'));
    assert('slide52 rels has chart ref', rels52.includes('/chart'));
    // Check chart XML exists
    const chartTarget52 = rels52.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget52) {
      const chartPath = 'ppt/charts/' + chartTarget52[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide52 chart XML exists', chartXml.length > 0);
      assert('slide52 has c:chartSpace', chartXml.includes('c:chartSpace'));
      assert('slide52 has c:barChart', chartXml.includes('c:barChart'));
      assert('slide52 has barDir col', chartXml.includes('val="col"'));
      assert('slide52 has c:ser', chartXml.includes('c:ser'));
      assert('slide52 has c:numCache', chartXml.includes('c:numCache'));
      assert('slide52 has c:strCache', chartXml.includes('c:strCache'));
      assert('slide52 has Q1 category', chartXml.includes('Q1'));
      assert('slide52 has Sales 2024', chartXml.includes('Sales 2024'));
    }
  }

  // ── Slide 53: Line chart ──
  {
    const slide53 = textFiles.get('ppt/slides/slide53.xml') || '';
    section('test_features.pptx — Slide 53: line chart');
    assert('slide53 has p:graphicFrame', slide53.includes('p:graphicFrame'));
    assert('slide53 has c:chart', slide53.includes('c:chart'));
    const rels53 = textFiles.get('ppt/slides/_rels/slide53.xml.rels') || '';
    const chartTarget53 = rels53.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget53) {
      const chartPath = 'ppt/charts/' + chartTarget53[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide53 chart has c:lineChart', chartXml.includes('c:lineChart'));
      assert('slide53 chart has Revenue', chartXml.includes('Revenue'));
      assert('slide53 chart has Jan', chartXml.includes('Jan'));
    }
  }

  // ── Slide 54: Pie chart ──
  {
    const slide54 = textFiles.get('ppt/slides/slide54.xml') || '';
    section('test_features.pptx — Slide 54: pie chart');
    assert('slide54 has p:graphicFrame', slide54.includes('p:graphicFrame'));
    const rels54 = textFiles.get('ppt/slides/_rels/slide54.xml.rels') || '';
    const chartTarget54 = rels54.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget54) {
      const chartPath = 'ppt/charts/' + chartTarget54[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide54 chart has c:pieChart', chartXml.includes('c:pieChart'));
      assert('slide54 chart has Desktop', chartXml.includes('Desktop'));
      assert('slide54 chart has Market Share', chartXml.includes('Market Share'));
    }
  }

  // ── Slide 55: Bar + Donut chart ──
  {
    const slide55 = textFiles.get('ppt/slides/slide55.xml') || '';
    section('test_features.pptx — Slide 55: bar + donut chart');
    assert('slide55 has p:graphicFrame', slide55.includes('p:graphicFrame'));
    const rels55 = textFiles.get('ppt/slides/_rels/slide55.xml.rels') || '';
    assert('slide55 has chart refs', rels55.includes('/chart'));
    // Check for both chart files
    const chartTargets55 = [...rels55.matchAll(/Target="([^"]*chart[^"]*)"/g)];
    assert('slide55 has 2 chart refs', chartTargets55.length >= 2);
  }

  // ── Slide 56: Column chart with data labels ──
  {
    const slide56 = textFiles.get('ppt/slides/slide56.xml') || '';
    section('test_features.pptx — Slide 56: column chart with data labels');
    assert('slide56 has p:graphicFrame', slide56.includes('p:graphicFrame'));
    const rels56 = textFiles.get('ppt/slides/_rels/slide56.xml.rels') || '';
    const chartTarget56 = rels56.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget56) {
      const chartPath = 'ppt/charts/' + chartTarget56[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide56 has c:barChart', chartXml.includes('c:barChart'));
      assert('slide56 has c:dLbls', chartXml.includes('c:dLbls'));
      assert('slide56 has showVal', chartXml.includes('showVal'));
      assert('slide56 has North category', chartXml.includes('North'));
    }
  }

  // ── Slide 57: Pie chart with dPt colors + % labels ──
  {
    const slide57 = textFiles.get('ppt/slides/slide57.xml') || '';
    section('test_features.pptx — Slide 57: pie chart with dPt + % labels');
    assert('slide57 has p:graphicFrame', slide57.includes('p:graphicFrame'));
    const rels57 = textFiles.get('ppt/slides/_rels/slide57.xml.rels') || '';
    const chartTarget57 = rels57.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget57) {
      const chartPath = 'ppt/charts/' + chartTarget57[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide57 has c:pieChart', chartXml.includes('c:pieChart'));
      assert('slide57 has c:dPt', chartXml.includes('c:dPt'));
      assert('slide57 has srgbClr', chartXml.includes('srgbClr'));
      assert('slide57 has showPercent', chartXml.includes('showPercent'));
      assert('slide57 has Chrome category', chartXml.includes('Chrome'));
    }
  }

  // ── Slide 58: Line chart with series colors + data labels ──
  {
    const slide58 = textFiles.get('ppt/slides/slide58.xml') || '';
    section('test_features.pptx — Slide 58: line chart with series colors');
    assert('slide58 has p:graphicFrame', slide58.includes('p:graphicFrame'));
    const rels58 = textFiles.get('ppt/slides/_rels/slide58.xml.rels') || '';
    const chartTarget58 = rels58.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget58) {
      const chartPath = 'ppt/charts/' + chartTarget58[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide58 has c:lineChart', chartXml.includes('c:lineChart'));
      assert('slide58 has c:dLbls', chartXml.includes('c:dLbls'));
      assert('slide58 has spPr with color', chartXml.includes('srgbClr'));
      assert('slide58 has Week 1', chartXml.includes('Week 1'));
    }
  }

  // ── Slide 59: Scatter chart + Area chart ──
  {
    const slide59 = textFiles.get('ppt/slides/slide59.xml') || '';
    section('test_features.pptx — Slide 59: scatter + area');
    assert('slide59 has p:graphicFrame', slide59.includes('p:graphicFrame'));
  }

  // ── Slide 60: Radar chart ──
  {
    const slide60 = textFiles.get('ppt/slides/slide60.xml') || '';
    section('test_features.pptx — Slide 60: radar chart');
    assert('slide60 has p:graphicFrame', slide60.includes('p:graphicFrame'));
  }

  // ── Slide 61: Bubble chart ──
  {
    const slide61 = textFiles.get('ppt/slides/slide61.xml') || '';
    section('test_features.pptx — Slide 61: bubble chart');
    assert('slide61 has p:graphicFrame', slide61.includes('p:graphicFrame'));
    const rels61 = textFiles.get('ppt/slides/_rels/slide61.xml.rels') || '';
    const chartTarget61 = rels61.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget61) {
      const chartPath = 'ppt/charts/' + chartTarget61[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide61 has c:bubbleChart', chartXml.includes('c:bubbleChart'));
      assert('slide61 has c:bubbleSize', chartXml.includes('c:bubbleSize'));
      assert('slide61 has xVal', chartXml.includes('c:xVal'));
      assert('slide61 has yVal', chartXml.includes('c:yVal'));
    }
  }

  // ── Slide 62: Stock chart ──
  {
    const slide62 = textFiles.get('ppt/slides/slide62.xml') || '';
    section('test_features.pptx — Slide 62: stock chart');
    assert('slide62 has p:graphicFrame', slide62.includes('p:graphicFrame'));
    const rels62 = textFiles.get('ppt/slides/_rels/slide62.xml.rels') || '';
    const chartTarget62 = rels62.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget62) {
      const chartPath = 'ppt/charts/' + chartTarget62[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide62 has c:stockChart', chartXml.includes('c:stockChart'));
      assert('slide62 has Open series', chartXml.includes('Open'));
      assert('slide62 has High series', chartXml.includes('High'));
      assert('slide62 has Low series', chartXml.includes('Low'));
      assert('slide62 has Close series', chartXml.includes('Close'));
    }
  }

  // ── Slide 63: All chart types overview ──
  {
    const slide63 = textFiles.get('ppt/slides/slide63.xml') || '';
    section('test_features.pptx — Slide 63: all chart types overview');
    assert('slide63 has p:graphicFrame', slide63.includes('p:graphicFrame'));
  }

  // ── Slide 64: Line chart with trendline ──
  {
    const slide64 = textFiles.get('ppt/slides/slide64.xml') || '';
    section('test_features.pptx — Slide 64: line chart with trendline');
    assert('slide64 has p:graphicFrame', slide64.includes('p:graphicFrame'));
    const rels64 = textFiles.get('ppt/slides/_rels/slide64.xml.rels') || '';
    const chartTarget64 = rels64.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget64) {
      const chartPath = 'ppt/charts/' + chartTarget64[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide64 has c:lineChart', chartXml.includes('c:lineChart'));
      assert('slide64 has c:trendline', chartXml.includes('c:trendline'));
      assert('slide64 has trendlineType linear', chartXml.includes('trendlineType') && chartXml.includes('linear'));
    }
  }

  // ── Slide 65: Column chart with error bars ──
  {
    const slide65 = textFiles.get('ppt/slides/slide65.xml') || '';
    section('test_features.pptx — Slide 65: column chart with error bars');
    assert('slide65 has p:graphicFrame', slide65.includes('p:graphicFrame'));
    const rels65 = textFiles.get('ppt/slides/_rels/slide65.xml.rels') || '';
    const chartTarget65 = rels65.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget65) {
      const chartPath = 'ppt/charts/' + chartTarget65[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide65 has c:barChart', chartXml.includes('c:barChart'));
      assert('slide65 has c:errBars', chartXml.includes('c:errBars'));
      assert('slide65 has fixedVal', chartXml.includes('fixedVal'));
    }
  }

  // ── Slide 66: Composite chart (column + line) ──
  {
    const slide66 = textFiles.get('ppt/slides/slide66.xml') || '';
    section('test_features.pptx — Slide 66: composite chart');
    assert('slide66 has p:graphicFrame', slide66.includes('p:graphicFrame'));
    const rels66 = textFiles.get('ppt/slides/_rels/slide66.xml.rels') || '';
    const chartTarget66 = rels66.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget66) {
      const chartPath = 'ppt/charts/' + chartTarget66[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide66 has c:barChart', chartXml.includes('c:barChart'));
      assert('slide66 has c:lineChart', chartXml.includes('c:lineChart'));
      assert('slide66 has Profit series', chartXml.includes('Profit'));
    }
  }

  // ── Slide 67: Surface chart ──
  {
    const slide67 = textFiles.get('ppt/slides/slide67.xml') || '';
    section('test_features.pptx — Slide 67: surface chart');
    assert('slide67 has p:graphicFrame', slide67.includes('p:graphicFrame'));
    const rels67 = textFiles.get('ppt/slides/_rels/slide67.xml.rels') || '';
    const chartTarget67 = rels67.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget67) {
      const chartPath = 'ppt/charts/' + chartTarget67[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide67 has c:surfaceChart', chartXml.includes('c:surfaceChart'));
      assert('slide67 has c:view3D', chartXml.includes('c:view3D'));
      assert('slide67 has Row 1 series', chartXml.includes('Row 1'));
      assert('slide67 has c:serAx', chartXml.includes('c:serAx'));
    }
  }

  // ── Slide 68: Pie of pie chart ──
  {
    const slide68 = textFiles.get('ppt/slides/slide68.xml') || '';
    section('test_features.pptx — Slide 68: ofPieChart');
    assert('slide68 has p:graphicFrame', slide68.includes('p:graphicFrame'));
    const rels68 = textFiles.get('ppt/slides/_rels/slide68.xml.rels') || '';
    const chartTarget68 = rels68.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget68) {
      const chartPath = 'ppt/charts/' + chartTarget68[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide68 has c:ofPieChart', chartXml.includes('c:ofPieChart'));
      assert('slide68 has c:ofPieType pie', chartXml.includes('ofPieType') && chartXml.includes('"pie"'));
      assert('slide68 has c:splitPos', chartXml.includes('c:splitPos'));
      assert('slide68 has Product A', chartXml.includes('Product A'));
    }
  }

  // ── Slide 69: 3D bar chart with view3D ──
  {
    const slide69 = textFiles.get('ppt/slides/slide69.xml') || '';
    section('test_features.pptx — Slide 69: 3D bar chart');
    assert('slide69 has p:graphicFrame', slide69.includes('p:graphicFrame'));
    const rels69 = textFiles.get('ppt/slides/_rels/slide69.xml.rels') || '';
    const chartTarget69 = rels69.match(/Target="([^"]*chart[^"]*)"/);
    if (chartTarget69) {
      const chartPath = 'ppt/charts/' + chartTarget69[1].replace(/^.*\//, '');
      const chartXml = textFiles.get(chartPath) || '';
      assert('slide69 has c:bar3DChart', chartXml.includes('c:bar3DChart'));
      assert('slide69 has c:view3D', chartXml.includes('c:view3D'));
      assert('slide69 has rotX', chartXml.includes('rotX'));
      assert('slide69 has depthPercent', chartXml.includes('depthPercent'));
      assert('slide69 has Revenue series', chartXml.includes('Revenue'));
    }
  }


  // ── Slide 73: Stacked / Percent-stacked bar charts ─────────────────────────
  {
    section('test_features.pptx — Slide 73: Stacked / Percent-stacked bar charts');
    // Should have chart references
    const rels73 = textFiles.get('ppt/slides/_rels/slide74.xml.rels') || '';
    const chartRefs73 = findRelTarget(rels73, 'chart');
    assert('slide73 has chart references', chartRefs73.length >= 2, `got ${chartRefs73.length}`);
    // Check the chart XMLs for grouping types
    let foundPercentStacked = false;
    let foundStacked = false;
    for (const [path, content] of textFiles) {
      if (path.startsWith('ppt/charts/') && path.endsWith('.xml')) {
        if (content.includes('percentStacked')) foundPercentStacked = true;
        if (content.includes('<c:grouping val="stacked"')) foundStacked = true;
      }
    }
    assert('has percentStacked grouping', foundPercentStacked);
    assert('has stacked grouping', foundStacked);
  }

  // ── Slides 89-94: ChartEx (cx: namespace) ──────────────────────────────────
  {
    section('test_features.pptx — Slides 89-94: ChartEx (cx: namespace)');
    const chartExTypes = ['waterfall', 'treemap', 'sunburst', 'clusteredColumn', 'boxWhisker', 'funnel'];
    const chartExNames = ['Waterfall', 'Treemap', 'Sunburst', 'Histogram', 'BoxWhisker', 'Funnel'];
    for (let i = 0; i < 6; i++) {
      const slideNum = 89 + i;
      const slideFile = `ppt/slides/slide${slideNum + 1}.xml`;
      const slideXml = textFiles.get(slideFile) || '';
      assert(`slide${slideNum} exists`, slideXml.length > 0);
      assert(`slide${slideNum} has cx:chart graphicFrame`,
        slideXml.includes('cx:chart') && slideXml.includes('a:graphicData'));

      // Check chartex file exists
      const chartExFile = `ppt/charts/chartEx${i + 1}.xml`;
      const chartExXml = textFiles.get(chartExFile) || '';
      assert(`chartEx${i + 1}.xml exists`, chartExXml.length > 0);
      assert(`chartEx${i + 1} has cx:chartSpace`, chartExXml.includes('cx:chartSpace'));
      assert(`chartEx${i + 1} has layoutId="${chartExTypes[i]}"`,
        chartExXml.includes(`layoutId="${chartExTypes[i]}"`));

      // Check rels
      const relsFile = `ppt/slides/_rels/slide${slideNum + 1}.xml.rels`;
      const relsXml = textFiles.get(relsFile) || '';
      assert(`slide${slideNum} rels has chartEx relationship`,
        relsXml.includes('chartEx'));
    }

    // Check Content_Types
    const contentTypes = textFiles.get('[Content_Types].xml') || '';
    assert('Content_Types has chartex content type',
      contentTypes.includes('chartex+xml'));
  }

  finishAssertions();
});
