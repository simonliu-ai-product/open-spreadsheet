import { describe, expect, it } from 'vitest';
import type { Scalar } from '../../engine/src/index.ts';
import { conditionalLook, iconCount, iconGlyph } from './features.ts';
import type { ConditionalRule, SheetData } from './types.ts';

// A1:A5 = 10, 20, 30, 40, 50
const values: Scalar[] = [10, 20, 30, 40, 50];
const lookFor = (rule: ConditionalRule) => {
  const sheet: SheetData = { name: 'S', cells: {}, conditional: [rule] };
  const look = conditionalLook(sheet, (row) => values[row] ?? null, { rows: 5, cols: 1 });
  return values.map((_, row) => look(row, 0));
};

describe('Excel conditional formats from xlsx', () => {
  it('Top and Bottom N, by count and by percent', () => {
    expect(lookFor({ range: 'A1:A5', type: 'top', rank: 2 }).map(Boolean)).toEqual([false, false, false, true, true]);
    expect(lookFor({ range: 'A1:A5', type: 'top', rank: 2, bottom: true }).map(Boolean)).toEqual([true, true, false, false, false]);
    expect(lookFor({ range: 'A1:A5', type: 'top', rank: 40, percent: true }).map(Boolean)).toEqual([false, false, false, true, true]);
  });

  it('above and below the average, with or without it', () => {
    expect(lookFor({ range: 'A1:A5', type: 'average' }).map(Boolean)).toEqual([false, false, false, true, true]);
    expect(lookFor({ range: 'A1:A5', type: 'average', below: true, equal: true }).map(Boolean)).toEqual([true, true, true, false, false]);
  });

  it('icon sets pick a level by percent of the range, reversible', () => {
    const arrows = lookFor({ range: 'A1:A5', type: 'iconSet', set: '3Arrows' }).map((l) => l?.icon?.glyph);
    // 10..50 sit at 0, 25, 50, 75 and 100 percent; Excel's 3-icon cuts are 33 and 67.
    expect(arrows).toEqual(['↓', '↓', '→', '↑', '↑']);
    const reversed = lookFor({ range: 'A1:A5', type: 'iconSet', set: '3Arrows', reverse: true }).map((l) => l?.icon?.glyph);
    expect(reversed).toEqual(['↑', '↑', '→', '↓', '↓']);
    expect(lookFor({ range: 'A1:A5', type: 'iconSet', set: '3TrafficLights1', thresholds: [50, 90] }).map((l) => l?.icon?.color)).toEqual([
      '#d93025',
      '#d93025',
      '#f9ab00',
      '#f9ab00',
      '#1e8e3e',
    ]);
    expect(iconCount('5Rating')).toBe(5);
    expect(iconGlyph('4Arrows', 3, 4).glyph).toBe('↑');
  });
});
