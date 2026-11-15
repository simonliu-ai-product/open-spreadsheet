import { describe, expect, it } from 'vitest';
import { chartData, niceTicks } from './charts.ts';

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
