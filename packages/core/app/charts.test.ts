import { describe, expect, it } from 'vitest';
import { BUILTIN_STYLES, chartData, chartFile, chartSvg, cleanStyle, findStyle, niceTicks } from './charts.ts';

describe('chart axes', () => {
  it('reach past the largest value and below the smallest', () => {
    expect(niceTicks(0, 192000)).toEqual([0, 50000, 100000, 150000, 200000]);
    expect(niceTicks(-12, 40)).toEqual([-20, 0, 20, 40]);
    expect(niceTicks(0, 20)).toEqual([0, 5, 10, 15, 20]);
  });
});

describe('chart data', () => {
  it('reads categories from the first column and series names from the first row', () => {
    expect(
      chartData([
        ['Rep', 'Q1', 'Q2'],
        ['Lin', 10, 12],
        ['May', 8, null],
      ]),
    ).toEqual({ categories: ['Lin', 'May'], series: [{ name: 'Q1', values: [10, 8] }, { name: 'Q2', values: [12, null] }] });
  });
  it('numbers the categories when there is no label column', () => {
    expect(chartData([['Qty'], [3], [5]])).toEqual({ categories: ['1', '2'], series: [{ name: 'Qty', values: [3, 5] }] });
  });
});

describe('every chart type', () => {
  const cells = [
    ['月', '營收', '成長率'],
    ['1月', 120, 0.05],
    ['2月', 150, 0.25],
    ['3月', 90, -0.4],
  ];
  const draw = (type: Parameters<typeof chartSvg>[0]['type'], data = chartData(cells as never)) =>
    chartSvg({ id: 'c', type, range: 'A1:C4', title: 'T', x: 0, y: 0, w: 1, h: 1 }, data, 720, 405);

  it('draws each type without a broken number', () => {
    for (const t of ['bar', 'barH', 'stacked', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo'] as const) {
      const svg = draw(t);
      expect(svg, t).toMatch(/^<svg /);
      expect(svg, t).not.toMatch(/NaN|Infinity|undefined/);
    }
  });

  it('gives each its shape', () => {
    // horizontal: 3 months × 2 series, each bar with its tooltip (the legend swatches have none)
    expect((draw('barH').match(/<rect [^>]*><title>/g) ?? []).length).toBe(6);
    expect(draw('doughnut')).toContain('class="c-total"');
    expect(draw('area')).toContain('fill-opacity="0.16"');
    // stacked: one column per month, its parts on top of each other
    const xs = [...draw('stacked').matchAll(/<rect x="([\d.]+)" y="[\d.]+" width="[\d.]+" height="[\d.]+" fill/g)].map((m) => m[1]);
    expect(new Set(xs).size).toBe(3);
    // combo: the rate is on its own right-hand axis
    expect(draw('combo')).toMatch(/<text x="(6[89]\d|7\d\d)[\d.]*" y="[\d.]+" class="c-tick">/);
    // scatter: the first numeric column is x
    const pts = chartData([
      ['單價', '數量'],
      [100, 3],
      [200, 5],
      [300, 2],
    ] as never);
    expect((draw('scatter', pts).match(/<circle /g) ?? []).length).toBe(3);
  });
});

it('a chart file takes its style: colours, background, type and grid', () => {
  const data = chartData([['', 'A', 'B'], ['x', 1, 2], ['y', 3, 4]] as never);
  const spec = { id: 'c', type: 'bar' as const, range: 'A1:C3', x: 0, y: 0, w: 1, h: 1 };
  const mono = findStyle('mono');
  const svg = chartFile(spec, data, 720, 405, undefined, mono);
  expect(svg).toContain(`fill="${mono.colors[0]}"`);
  expect(svg).toContain('.c-grid{stroke:none');
  expect(svg).toContain('monospace');
  expect(chartFile(spec, data, 720, 405, undefined, findStyle('night'))).toContain('<rect width="100%" height="100%" fill="#15181e"/>');
  // a style that is gone falls back to Default
  expect(findStyle('gone', [])).toBe(BUILTIN_STYLES[0]);
});

it('a style read from a file is checked field by field', () => {
  expect(cleanStyle({ id: 'Bad Id' })).toBeNull();
  const s = cleanStyle({ id: 'brand', colors: ['#ABCDEF', 'red', '#123'], radius: 99, font: 'comic', background: 'url(x)' })!;
  expect(s.colors).toEqual(['#abcdef']);
  expect(s.radius).toBe(8);
  expect(s.font).toBe('sans');
  expect(s.background).toBe('#ffffff');
});
