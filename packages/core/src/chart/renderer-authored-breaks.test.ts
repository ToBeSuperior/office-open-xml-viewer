import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartRect, ChartSeries, SecondaryValueAxis } from '../types/chart';
import { renderChart } from './renderer.js';
import { renderChartExChart } from './chart-ex-renderer.js';
import { renderSimpleThreeDChart } from './three-d-renderer.js';
import { renderRegionMapChart } from './region-map-renderer.js';
import { drawSecondaryCategoryAxis, secondaryCategoryLabelExtraH } from './shared/secondary-axis.js';

// Authored line breaks in chart label text must start new lines in every
// label path (PowerPoint draws the waterfall category "FY2024\n1Q営業利益" on
// two lines). Each case renders the same chart with and without the break.

interface Label { text: string; x: number; y: number; align: string; baseline: string; font: string }
interface Segment { x1: number; y1: number; x2: number; y2: number }

interface FilledRect { x: number; y: number; w: number; h: number }

function recordingContext(): { ctx: CanvasRenderingContext2D; labels: Label[]; segments: Segment[]; rects: FilledRect[] } {
  const rects: FilledRect[] = [];
  const labels: Label[] = [];
  const segments: Segment[] = [];
  const path: Array<{ x: number; y: number }> = [];
  const state: Record<string, unknown> = {
    font: '10px sans-serif', fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1, lineCap: 'butt', lineJoin: 'miter',
  };
  const handler: ProxyHandler<Record<string, unknown>> = {
    get(_target, prop: string) {
      if (prop in state && typeof state[prop] !== 'function') return state[prop];
      switch (prop) {
        case 'measureText': return (text: string) => ({ width: Array.from(String(text)).length * 7 });
        case 'beginPath': return () => { path.length = 0; };
        case 'moveTo': return (x: number, y: number) => { path.length = 0; path.push({ x, y }); };
        case 'lineTo': return (x: number, y: number) => { path.push({ x, y }); };
        case 'stroke': return () => {
          if (path.length === 2) segments.push({ x1: path[0].x, y1: path[0].y, x2: path[1].x, y2: path[1].y });
        };
        case 'fillText': return (text: string, x: number, y: number) => labels.push({
          text, x, y, align: String(state.textAlign), baseline: String(state.textBaseline), font: String(state.font),
        });
        case 'getLineDash': return () => [];
        case 'fillRect': return (x: number, y: number, w: number, h: number) => { rects.push({ x, y, w, h }); };
        case 'createLinearGradient':
        case 'createRadialGradient': return () => ({ addColorStop() {} });
        case 'createPattern': return () => null;
        default: return () => undefined;
      }
    },
    set(_target, prop: string, value) { state[prop] = value; return true; },
  };
  return { ctx: new Proxy(state, handler) as unknown as CanvasRenderingContext2D, labels, segments, rects };
}

function series(over: Partial<ChartSeries> = {}): ChartSeries {
  return { name: 'S1', color: '4472C4', values: [10, 20], ...over };
}

function model(over: Partial<ChartModel>): ChartModel {
  return {
    chartType: 'clusteredColumn',
    title: null,
    categories: ['Alpha', 'Beta'],
    series: [series()],
    showDataLabels: false,
    valMin: 0,
    valMax: 20,
    catAxisTitle: null,
    valAxisTitle: null,
    catAxisHidden: false,
    valAxisHidden: false,
    catAxisLineHidden: false,
    valAxisLineHidden: true,
    plotAreaBg: null,
    chartBg: null,
    showLegend: false,
    legendPos: null,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'none',
    catAxisMajorTickMark: 'none',
    titleFontSizeHpt: null,
    titleFontColor: null,
    titleFontFace: null,
    catAxisFontSizeHpt: 1000,
    valAxisFontSizeHpt: 1000,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisMajorGridlines: false,
    ...over,
  };
}

const RECT: ChartRect = { x: 0, y: 0, w: 640, h: 360 };
// 10pt at ptToPx 1 → 10 px font; stacked axis lines are font + 2 px apart.
const LINE = 12;

function render(chart: ChartModel): ReturnType<typeof recordingContext> {
  const rec = recordingContext();
  renderChart(
    rec.ctx, chart, RECT, 1, 0,
    { render: renderSimpleThreeDChart }, undefined, undefined,
    { render: renderChartExChart },
  );
  return rec;
}

function label(rec: ReturnType<typeof recordingContext>, text: string): Label {
  const found = rec.labels.find(entry => entry.text === text);
  if (!found) throw new Error(`no fillText(${JSON.stringify(text)}) in ${JSON.stringify(rec.labels.map(entry => entry.text))}`);
  return found;
}

function expectNoJoinedText(rec: ReturnType<typeof recordingContext>): void {
  expect(rec.labels.filter(entry => /[\r\n]/.test(entry.text))).toEqual([]);
}

describe('authored breaks in category-axis labels', () => {
  for (const chartType of ['clusteredColumn', 'line', 'area'] as const) {
    it(`${chartType}: draws each line and lifts the plot bottom by the extra line`, () => {
      const broken = render(model({ chartType, categories: ['FY2024\n1Q', 'Beta'] }));
      const plain = render(model({ chartType, categories: ['FY2024', 'Beta'] }));
      expectNoJoinedText(broken);
      const first = label(broken, 'FY2024');
      const second = label(broken, '1Q');
      expect(first.baseline).toBe('top');
      expect(second.x).toBe(first.x);
      expect(second.y - first.y).toBeCloseTo(LINE, 6);
      expect(label(plain, 'FY2024').y - first.y).toBeCloseTo(LINE, 6);
    });
  }

  it('stock: draws each category line and lifts the plot bottom', () => {
    const stock = (categories: string[]) => model({
      chartType: 'stock',
      categories,
      valMax: 60,
      series: [
        series({ name: 'High', values: [55, 57] }),
        series({ name: 'Low', values: [11, 12] }),
        series({ name: 'Close', values: [40, 30] }),
      ],
    });
    const broken = render(stock(['FY2024\n1Q', 'Beta']));
    const plain = render(stock(['FY2024', 'Beta']));
    expectNoJoinedText(broken);
    expect(label(broken, '1Q').y - label(broken, 'FY2024').y).toBeCloseTo(LINE, 6);
    expect(label(plain, 'FY2024').y - label(broken, 'FY2024').y).toBeCloseTo(LINE, 6);
  });

  it('rotated column labels stack lines perpendicular to the text and reserve their vertical extent', () => {
    const rot = -2_700_000; // -45°
    const broken = render(model({ categories: ['FY2024\n1Q', 'Beta'], catAxisLabelRotation: rot }));
    const plain = render(model({ categories: ['FY2024', 'Beta'], catAxisLabelRotation: rot }));
    expectNoJoinedText(broken);
    // Drawn in the rotated frame around the pivot, centred as a block.
    expect(label(broken, 'FY2024').y).toBeCloseTo(-LINE / 2, 6);
    expect(label(broken, '1Q').y).toBeCloseTo(LINE / 2, 6);
    expect(label(plain, 'FY2024').y).toBe(0);
    // Bars stand on a plot bottom lifted by |cos 45°| of the extra line.
    const barBottom = (rec: ReturnType<typeof recordingContext>) =>
      Math.max(...rec.segments.map(segment => segment.y1).filter(y => y < RECT.h));
    expect(barBottom(plain) - barBottom(broken)).toBeCloseTo(LINE * Math.SQRT1_2, 6);
  });

  it('horizontal bars centre a multi-line category on its slot', () => {
    const broken = render(model({ chartType: 'clusteredBarH', categories: ['FY2024\n1Q', 'Beta'] }));
    const plain = render(model({ chartType: 'clusteredBarH', categories: ['FY2024', 'Beta'] }));
    expectNoJoinedText(broken);
    const first = label(broken, 'FY2024');
    const second = label(broken, '1Q');
    expect(first.baseline).toBe('middle');
    expect(first.align).toBe('right');
    expect(second.y - first.y).toBeCloseTo(LINE, 6);
    expect((first.y + second.y) / 2).toBeCloseTo(label(plain, 'FY2024').y, 6);
  });

  it('multi-level columns start outer rows and brackets below the inner labels\' real line count', () => {
    const multi = (inner: string[], outer: string) => model({
      categories: inner,
      categoryLevels: [inner, [outer, '']],
      catAxisLineColor: '000000',
    });
    const plain = render(multi(['Alpha', 'Beta'], 'Group'));
    const broken = render(multi(['Alpha\nA2', 'Beta'], 'Group'));
    const outerBroken = render(multi(['Alpha', 'Beta'], 'Group\nG2'));
    expectNoJoinedText(broken);
    expectNoJoinedText(outerBroken);
    const offset = (rec: ReturnType<typeof recordingContext>) => label(rec, 'Group').y - label(rec, 'Alpha').y;
    expect(offset(broken) - offset(plain)).toBeCloseTo(LINE, 6);
    // A broken outer label draws both lines; the outer rows/brackets for it grow.
    expect(label(outerBroken, 'G2').y - label(outerBroken, 'Group').y).toBeCloseTo(LINE, 6);
    const deepest = (rec: ReturnType<typeof recordingContext>) => {
      const axisY = label(rec, 'Alpha').y;
      return Math.max(...rec.segments.filter(segment => segment.x1 === segment.x2 && segment.y2 > axisY)
        .map(segment => segment.y2 - axisY));
    };
    expect(deepest(broken) - deepest(plain)).toBeCloseTo(LINE, 6);
    expect(deepest(outerBroken) - deepest(plain)).toBeCloseTo(LINE, 6);
  });
});

describe('authored breaks in value-axis labels', () => {
  // A number format whose literal contains a line feed.
  const code = '0"\nunits"';
  for (const chartType of ['clusteredColumn', 'line', 'area', 'scatter'] as const) {
    it(`${chartType}: stacks each tick label on its tick and keeps the top label inside the reserve`, () => {
      const scatterOver = chartType === 'scatter'
        ? { categories: ['0', '10'], series: [series({ categories: ['0', '10'], values: [0, 20] })], scatterStyle: 'marker' as const }
        : {};
      const broken = render(model({ chartType, valAxisFormatCode: code, ...scatterOver }));
      const plain = render(model({ chartType, valAxisFormatCode: '0', ...scatterOver }));
      expectNoJoinedText(broken);
      const tops = broken.labels.filter(entry => entry.text === '20');
      const unitLines = broken.labels.filter(entry => entry.text === 'units');
      expect(tops).toHaveLength(1);
      expect(unitLines.length).toBeGreaterThan(1);
      const top = tops[0];
      const topUnits = unitLines.find(entry => Math.abs(entry.y - top.y - LINE) < 1e-6);
      expect(topUnits?.x).toBe(top.x);
      expect(top.baseline).toBe('middle');
      // The plot top moves down by half the extra line, so the block's first
      // line sits where the single-line label sat.
      expect(top.y).toBeCloseTo(label(plain, '20').y, 6);
    });
  }

  it('horizontal-bar value labels hang their extra lines below the axis', () => {
    const broken = render(model({ chartType: 'clusteredBarH', valAxisFormatCode: code }));
    const plain = render(model({ chartType: 'clusteredBarH', valAxisFormatCode: '0' }));
    expectNoJoinedText(broken);
    const zero = label(broken, '0');
    const units = broken.labels.filter(entry => entry.text === 'units' && entry.x === zero.x);
    expect(units).toHaveLength(1);
    expect(units[0].y - zero.y).toBeCloseTo(LINE, 6);
    expect(units[0].y).toBeCloseTo(label(plain, '0').y, 6);
  });

  it('radar and stock value labels draw each line', () => {
    const radar = render(model({
      chartType: 'radar', categories: ['A', 'B', 'C'],
      series: [series({ values: [5, 10, 20] })], valAxisFormatCode: code,
    }));
    expectNoJoinedText(radar);
    expect(radar.labels.some(entry => entry.text === 'units')).toBe(true);
    const stock = render(model({
      chartType: 'stock', valAxisFormatCode: code, valMax: 60,
      series: [series({ values: [55, 57] }), series({ values: [11, 12] }), series({ values: [40, 30] })],
    }));
    expectNoJoinedText(stock);
    expect(stock.labels.some(entry => entry.text === 'units')).toBe(true);
  });
});

describe('authored breaks in legends, data tables, secondary and ChartEx/3-D labels', () => {
  it('side legend keeps every authored line, including an interior blank line', () => {
    const rec = render(model({ showLegend: true, legendPos: 'r', series: [series({ name: 'A\n\nB' })] }));
    expectNoJoinedText(rec);
    const a = label(rec, 'A');
    const b = label(rec, 'B');
    expect(b.x).toBe(a.x);
    expect(b.y - a.y).toBeCloseTo(20, 6); // two font-size intervals
  });

  it('top/bottom legend draws each line inside a row tall enough for the block', () => {
    const rec = render(model({ showLegend: true, legendPos: 'b', series: [series({ name: 'North\nEast' })] }));
    expectNoJoinedText(rec);
    expect(label(rec, 'East').y - label(rec, 'North').y).toBeCloseTo(10, 6);
  });

  it('data table rows grow for a multi-line series name', () => {
    const table = { showHorizontalBorder: true, showVerticalBorder: true, showOutline: true, showKeys: false };
    const rec = render(model({ dataTable: table, series: [series({ name: 'Net\nSales' }), series({ name: 'Cost', values: [1, 2] })] }));
    const plain = render(model({ dataTable: table, series: [series({ name: 'Net' }), series({ name: 'Cost', values: [1, 2] })] }));
    expectNoJoinedText(rec);
    // Default 9pt table text: lines 10.8 px apart; the next row starts one
    // table line lower relative to the first.
    const net = label(rec, 'Net');
    expect(label(rec, 'Sales').y - net.y).toBeCloseTo(10.8, 6);
    expect((label(rec, 'Cost').y - net.y) - (label(plain, 'Cost').y - label(plain, 'Net').y))
      .toBeCloseTo(10.8, 6);
  });

  it('secondary (top) category labels grow upward and widen their band', () => {
    const rec = recordingContext();
    const axis = { hidden: false, tickLabelPos: 'nextTo', lineHidden: true, majorTickMark: 'none', fontSizeHpt: 1000 } as unknown as SecondaryValueAxis;
    const chart = model({});
    drawSecondaryCategoryAxis(rec.ctx, chart, axis, ['FY2024\n1Q', 'Beta'], RECT, 0, 100, 200, 1);
    expect(label(rec, 'FY2024').baseline).toBe('bottom');
    expect(label(rec, '1Q').y - label(rec, 'FY2024').y).toBeCloseTo(LINE, 6);
    expect(secondaryCategoryLabelExtraH(['FY2024\n1Q', 'Beta'], axis, chart, 10)).toBe(LINE);
    expect(secondaryCategoryLabelExtraH(['FY2024', 'Beta'], axis, chart, 10)).toBe(0);
  });

  it('ChartEx funnel and box-whisker categories draw each line', () => {
    const funnel = render(model({ chartType: 'funnel', categories: ['FY2024\n1Q', 'Beta'] }));
    expectNoJoinedText(funnel);
    expect(label(funnel, '1Q').y - label(funnel, 'FY2024').y).toBeCloseTo(LINE, 6);

    const box = (category: string) => model({
      chartType: 'boxWhisker',
      categories: [category],
      series: [series({ values: [] })],
      chartexBox: {
        categories: [category],
        series: [{
          name: 'S1', color: '4472C4', valuesByCategory: [[1, 2, 3, 4, 5]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    });
    const broken = render(box('FY2024\n1Q'));
    const plain = render(box('FY2024'));
    expectNoJoinedText(broken);
    expect(label(broken, '1Q').y - label(broken, 'FY2024').y).toBeCloseTo(LINE, 6);
    expect(label(plain, 'FY2024').y - label(broken, 'FY2024').y).toBeCloseTo(LINE, 6);
  });

  it('3-D category and legend labels draw each line', () => {
    const rec = render(model({
      categories: ['FY2024\n1Q', 'Beta'],
      chartType: 'clusteredBar',
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      plotGroups: [{ kind: 'bar3D', seriesStart: 0, seriesCount: 1, categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none', varyColors: false, barDirection: 'col', grouping: 'clustered' }],
      showLegend: true,
      legendPos: 'b',
      series: [series({ name: 'North\nEast' })],
    }));
    expectNoJoinedText(rec);
    expect(label(rec, '1Q').y).toBeGreaterThan(label(rec, 'FY2024').y);
    expect(label(rec, 'East').y).toBeGreaterThan(label(rec, 'North').y);
  });

  it('surface category, series and value labels draw each line', () => {
    const rec = render(model({
      chartType: 'surface', categories: ['FY2024\n1Q', 'Beta', 'Gamma'],
      valMin: 0, valMax: 10, valAxisMajorUnit: 5, valAxisFormatCode: '0"\nunits"',
      series: [series({ name: 'North\nEast', values: [1, 9, 2] }), series({ name: 'South', values: [9, 1, 8] })],
      threeD: { rotationX: 30, rotationY: 20, perspective: 30 },
    }));
    expectNoJoinedText(rec);
    expect(label(rec, '1Q').y - label(rec, 'FY2024').y).toBeCloseTo(LINE, 6);
    expect(rec.labels.some(entry => entry.text === 'units')).toBe(true);
    expect(rec.labels.some(entry => entry.text === 'East')).toBe(true);
  });

  it('display-unit labels draw each line', () => {
    const rec = render(model({
      valMax: 20_000,
      series: [series({ values: [10_000, 20_000] })],
      valAxisDisplayUnits: { builtInUnit: 'thousands', divisor: 1000, label: { text: 'Thousands\nUSD' } },
    }));
    expectNoJoinedText(rec);
    expect(label(rec, 'USD').y - label(rec, 'Thousands').y).toBeCloseTo(LINE, 6);
  });

  it('region-map legend name, scale values and region labels draw each line', () => {
    const rec = recordingContext();
    const handled = renderRegionMapChart(rec.ctx, {
      chartType: 'regionMap',
      categories: ['United\nStates', 'Canada'],
      series: [{ name: 'Net\nSales', values: [850, 420], valFormatCode: '0"\nunits"' }],
      showLegend: true,
      chartexRegionMap: {
        rows: [
          { label: 'United\nStates', entityId: 'USA', value: 850 },
          { label: 'Canada', value: 420 },
        ],
        regionLabelLayout: 'showAll',
        geography: { projectionType: 'robinson', cachePresent: false },
      },
    } as ChartModel, RECT, 1);
    expect(handled).toBe(true);
    expectNoJoinedText(rec);
    expect(label(rec, 'Sales').y - label(rec, 'Net').y).toBeCloseTo(11, 6);
    expect(label(rec, 'States').y - label(rec, 'United').y).toBeCloseTo(10, 6); // 8 px region labels
    expect(rec.labels.filter(entry => entry.text === 'units')).toHaveLength(2);
  });

});

describe('authored-break reservations stay on the painted side and inside the chart', () => {
  it('scatter X labels at the high edge reserve their extra lines at the top, inside the chart', () => {
    const scatter = (labelPos: 'high' | 'nextTo') => model({
      chartType: 'scatter', categories: ['0', '10'], scatterStyle: 'marker',
      series: [series({ categories: ['0', '10'], values: [0, 20] })],
      catAxisTickLabelPos: labelPos,
      catAxisFormatCode: '0"\na\nb\nc"',
    });
    const rec = render(scatter('high'));
    expectNoJoinedText(rec);
    const xLines = rec.labels.filter(entry => ['a', 'b', 'c'].includes(entry.text));
    expect(xLines.length).toBeGreaterThan(0);
    for (const line of xLines) {
      expect(line.baseline).toBe('bottom');
      expect(line.y - 10).toBeGreaterThanOrEqual(RECT.y);
    }
  });

  it('3-D side legend keeps the last authored line after a long wrapped first line', () => {
    const rec = render(model({
      chartType: 'clusteredBar',
      threeD: { rotationX: 15, rotationY: 20, perspective: 30 },
      plotGroups: [{ kind: 'bar3D', seriesStart: 0, seriesCount: 2, categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none', varyColors: false, barDirection: 'col', grouping: 'clustered' }],
      showLegend: true,
      legendPos: 'r',
      series: [
        series({ name: `${'word '.repeat(20)}\nLast` }),
        series({ name: 'Other', values: [5, 6] }),
      ],
    }));
    expectNoJoinedText(rec);
    expect(rec.labels.some(entry => entry.text === 'Last')).toBe(true);
  });

  it('region-map legend grows so multi-line names and values stay above the map', () => {
    const rec = recordingContext();
    renderRegionMapChart(rec.ctx, {
      chartType: 'regionMap',
      categories: ['Canada'],
      series: [{ name: 'A\nB\nC\nD', values: [420, 850], valFormatCode: '0"\nunits"' }],
      showLegend: true,
      plotAreaBg: 'FFFFFF',
      chartexRegionMap: {
        rows: [{ label: 'Canada', value: 420 }, { label: 'United States', value: 850 }],
        geography: { projectionType: 'robinson', cachePresent: false },
      },
    } as ChartModel, RECT, 1);
    const legendBottom = Math.max(...rec.labels.filter(entry => entry.text === 'units').map(entry => entry.y + 9));
    const mapTop = Math.min(...rec.rects.filter(r => r.w > RECT.w * 0.9).map(r => r.y));
    expect(legendBottom).toBeLessThanOrEqual(mapTop);
  });

  it('box-whisker never paints a negative-size plot when category lines exhaust the frame', () => {
    const category = Array.from({ length: 32 }, (_, index) => `L${index}`).join('\n');
    const rec = render(model({
      chartType: 'boxWhisker',
      plotAreaBg: 'FFFFFF',
      categories: [category],
      series: [series({ values: [] })],
      chartexBox: {
        categories: [category],
        series: [{
          name: 'S1', color: '4472C4', valuesByCategory: [[1, 2, 3, 4, 5]],
          meanMarker: false, meanLine: false, showOutliers: false, showNonoutliers: false,
          quartileMethod: 'inclusive',
        }],
      },
    }));
    expect(rec.rects.filter(r => r.w < 0 || r.h < 0)).toEqual([]);
  });
});
