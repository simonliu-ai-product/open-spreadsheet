/** Signatures and one-line descriptions for formula autocomplete. */
import { FUNCTIONS } from '@open-spreadsheet/engine';
import { isZh } from './i18n.ts';

type Help = [signature: string, en: string, zh: string];

const HELP: Record<string, Help> = {
  SUM: ['SUM(number1, [number2], …)', 'Adds numbers', '加總數值'],
  AVERAGE: ['AVERAGE(number1, [number2], …)', 'Average of numbers', '平均值'],
  COUNT: ['COUNT(value1, [value2], …)', 'Counts numbers', '計算數值個數'],
  COUNTA: ['COUNTA(value1, [value2], …)', 'Counts non-empty cells', '計算非空白個數'],
  COUNTBLANK: ['COUNTBLANK(range)', 'Counts empty cells', '計算空白個數'],
  MAX: ['MAX(number1, [number2], …)', 'Largest value', '最大值'],
  MIN: ['MIN(number1, [number2], …)', 'Smallest value', '最小值'],
  MEDIAN: ['MEDIAN(number1, [number2], …)', 'Median', '中位數'],
  LARGE: ['LARGE(array, k)', 'k-th largest', '第 k 大的值'],
  SMALL: ['SMALL(array, k)', 'k-th smallest', '第 k 小的值'],
  'STDEV.S': ['STDEV.S(number1, [number2], …)', 'Sample standard deviation', '樣本標準差'],
  PRODUCT: ['PRODUCT(number1, [number2], …)', 'Multiplies numbers', '乘積'],
  SUMPRODUCT: ['SUMPRODUCT(array1, [array2], …)', 'Sum of products', '乘積的總和'],
  SUMIF: ['SUMIF(range, criteria, [sum_range])', 'Sums cells that meet a condition', '依條件加總'],
  SUMIFS: ['SUMIFS(sum_range, range1, criteria1, …)', 'Sums with several conditions', '依多個條件加總'],
  COUNTIF: ['COUNTIF(range, criteria)', 'Counts cells that meet a condition', '依條件計數'],
  COUNTIFS: ['COUNTIFS(range1, criteria1, …)', 'Counts with several conditions', '依多個條件計數'],
  AVERAGEIF: ['AVERAGEIF(range, criteria, [average_range])', 'Averages cells that meet a condition', '依條件平均'],
  ROUND: ['ROUND(number, num_digits)', 'Rounds to digits', '四捨五入'],
  ROUNDUP: ['ROUNDUP(number, num_digits)', 'Rounds away from zero', '無條件進位'],
  ROUNDDOWN: ['ROUNDDOWN(number, num_digits)', 'Rounds toward zero', '無條件捨去'],
  INT: ['INT(number)', 'Rounds down to an integer', '取整數（向下）'],
  TRUNC: ['TRUNC(number, [num_digits])', 'Truncates', '截去小數'],
  ABS: ['ABS(number)', 'Absolute value', '絕對值'],
  MOD: ['MOD(number, divisor)', 'Remainder', '餘數'],
  POWER: ['POWER(number, power)', 'Raises to a power', '次方'],
  SQRT: ['SQRT(number)', 'Square root', '平方根'],
  PI: ['PI()', 'π', '圓周率'],
  SIGN: ['SIGN(number)', 'Sign of a number', '正負號'],
  IF: ['IF(logical_test, value_if_true, [value_if_false])', 'Chooses by a condition', '依條件傳回不同值'],
  IFS: ['IFS(test1, value1, [test2, value2], …)', 'First condition that holds', '多重條件'],
  IFERROR: ['IFERROR(value, value_if_error)', 'Replaces an error', '錯誤時改傳回指定值'],
  IFNA: ['IFNA(value, value_if_na)', 'Replaces #N/A', '#N/A 時改傳回指定值'],
  SWITCH: ['SWITCH(expression, value1, result1, …, [default])', 'Matches a value', '依值對應結果'],
  AND: ['AND(logical1, [logical2], …)', 'All true', '全部成立'],
  OR: ['OR(logical1, [logical2], …)', 'Any true', '任一成立'],
  NOT: ['NOT(logical)', 'Reverses', '反轉'],
  XOR: ['XOR(logical1, [logical2], …)', 'Exclusive or', '互斥或'],
  VLOOKUP: ['VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', 'Looks down the first column', '垂直查表'],
  HLOOKUP: ['HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])', 'Looks across the first row', '水平查表'],
  XLOOKUP: ['XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])', 'Looks up in any direction', '萬用查表'],
  INDEX: ['INDEX(array, row_num, [column_num])', 'Value at a position', '取得指定位置的值'],
  MATCH: ['MATCH(lookup_value, lookup_array, [match_type])', 'Position of a value', '找出值的位置'],
  CHOOSE: ['CHOOSE(index_num, value1, [value2], …)', 'Picks by index', '依序號選擇'],
  ROW: ['ROW([reference])', 'Row number', '列號'],
  COLUMN: ['COLUMN([reference])', 'Column number', '欄號'],
  ROWS: ['ROWS(array)', 'Number of rows', '列數'],
  COLUMNS: ['COLUMNS(array)', 'Number of columns', '欄數'],
  LEFT: ['LEFT(text, [num_chars])', 'Leftmost characters', '取左邊字元'],
  RIGHT: ['RIGHT(text, [num_chars])', 'Rightmost characters', '取右邊字元'],
  MID: ['MID(text, start_num, num_chars)', 'Characters from the middle', '取中間字元'],
  LEN: ['LEN(text)', 'Length of text', '字元數'],
  FIND: ['FIND(find_text, within_text, [start_num])', 'Finds text (case-sensitive)', '尋找文字（分大小寫）'],
  SEARCH: ['SEARCH(find_text, within_text, [start_num])', 'Finds text, wildcards allowed', '尋找文字（可萬用字元）'],
  SUBSTITUTE: ['SUBSTITUTE(text, old_text, new_text, [instance_num])', 'Replaces text', '取代文字'],
  REPLACE: ['REPLACE(old_text, start_num, num_chars, new_text)', 'Replaces by position', '依位置取代'],
  UPPER: ['UPPER(text)', 'Upper case', '轉大寫'],
  LOWER: ['LOWER(text)', 'Lower case', '轉小寫'],
  PROPER: ['PROPER(text)', 'Capitalises words', '字首大寫'],
  TRIM: ['TRIM(text)', 'Removes extra spaces', '移除多餘空白'],
  CONCAT: ['CONCAT(text1, [text2], …)', 'Joins text', '串接文字'],
  CONCATENATE: ['CONCATENATE(text1, [text2], …)', 'Joins text', '串接文字'],
  TEXTJOIN: ['TEXTJOIN(delimiter, ignore_empty, text1, …)', 'Joins with a delimiter', '以分隔符號串接'],
  TEXT: ['TEXT(value, format_text)', 'Formats a number as text', '依格式轉成文字'],
  HYPERLINK: ['HYPERLINK(link_location, [friendly_name])', 'A link: shows the name, ⌘-click opens it', '超連結：顯示名稱，⌘＋點擊開啟'],
  VALUE: ['VALUE(text)', 'Text to number', '文字轉數值'],
  REPT: ['REPT(text, number_times)', 'Repeats text', '重複文字'],
  EXACT: ['EXACT(text1, text2)', 'Exact comparison', '完全相同'],
  DATE: ['DATE(year, month, day)', 'Builds a date', '建立日期'],
  TODAY: ['TODAY()', "Today's date", '今天日期'],
  NOW: ['NOW()', 'Current date and time', '現在時間'],
  YEAR: ['YEAR(serial_number)', 'Year of a date', '年'],
  MONTH: ['MONTH(serial_number)', 'Month of a date', '月'],
  DAY: ['DAY(serial_number)', 'Day of a date', '日'],
  WEEKDAY: ['WEEKDAY(serial_number, [return_type])', 'Day of the week', '星期幾'],
  DAYS: ['DAYS(end_date, start_date)', 'Days between dates', '相差天數'],
  EDATE: ['EDATE(start_date, months)', 'Date months away', '幾個月後的日期'],
  EOMONTH: ['EOMONTH(start_date, months)', 'Last day of a month', '月底日期'],
  PMT: ['PMT(rate, nper, pv, [fv], [type])', 'Loan payment', '貸款每期付款'],
  PV: ['PV(rate, nper, pmt, [fv], [type])', 'Present value', '現值'],
  FV: ['FV(rate, nper, pmt, [pv], [type])', 'Future value', '未來值'],
  NPV: ['NPV(rate, value1, [value2], …)', 'Net present value', '淨現值'],
  AVERAGEIFS: ['AVERAGEIFS(average_range, range1, criteria1, …)', 'Average with several conditions', '依多個條件平均'],
  MAXIFS: ['MAXIFS(max_range, range1, criteria1, …)', 'Largest value meeting conditions', '依條件取最大值'],
  MINIFS: ['MINIFS(min_range, range1, criteria1, …)', 'Smallest value meeting conditions', '依條件取最小值'],
  SUBTOTAL: ['SUBTOTAL(function_num, ref1, …)', 'Subtotal by function code (9 = SUM)', '依函數代碼小計（9 為加總）'],
  CEILING: ['CEILING(number, significance)', 'Rounds up to a multiple', '無條件進位到倍數'],
  FLOOR: ['FLOOR(number, significance)', 'Rounds down to a multiple', '無條件捨去到倍數'],
  MROUND: ['MROUND(number, multiple)', 'Rounds to the nearest multiple', '四捨五入到倍數'],
  'RANK.EQ': ['RANK.EQ(number, ref, [order])', 'Rank in a list', '排名'],
  'PERCENTILE.INC': ['PERCENTILE.INC(array, k)', 'k-th percentile', '百分位數'],
  'QUARTILE.INC': ['QUARTILE.INC(array, quart)', 'Quartile', '四分位數'],
  CORREL: ['CORREL(array1, array2)', 'Correlation coefficient', '相關係數'],
  'FORECAST.LINEAR': ['FORECAST.LINEAR(x, known_ys, known_xs)', 'Linear forecast', '線性預測'],
  NETWORKDAYS: ['NETWORKDAYS(start_date, end_date, [holidays])', 'Working days between dates', '兩日期間的工作天數'],
  WORKDAY: ['WORKDAY(start_date, days, [holidays])', 'Date after working days', '幾個工作天後的日期'],
  DATEDIF: ['DATEDIF(start_date, end_date, unit)', 'Difference in "Y", "M", "D", "YM", "MD" or "YD"', '日期差（單位 Y、M、D、YM、MD、YD）'],
  YEARFRAC: ['YEARFRAC(start_date, end_date, [basis])', 'Fraction of a year', '佔一年的比例'],
  WEEKNUM: ['WEEKNUM(serial_number, [return_type])', 'Week of the year', '一年中的第幾週'],
  TIME: ['TIME(hour, minute, second)', 'Builds a time', '建立時間'],
  HOUR: ['HOUR(serial_number)', 'Hour of a time', '時'],
  TEXTBEFORE: ['TEXTBEFORE(text, delimiter, [instance_num])', 'Text before a delimiter', '分隔符號之前的文字'],
  TEXTAFTER: ['TEXTAFTER(text, delimiter, [instance_num])', 'Text after a delimiter', '分隔符號之後的文字'],
  XMATCH: ['XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])', 'Position of a value, any direction', '找出值的位置（萬用）'],
  IRR: ['IRR(values, [guess])', 'Internal rate of return', '內部報酬率'],
  RATE: ['RATE(nper, pmt, pv, [fv], [type], [guess])', 'Interest rate per period', '每期利率'],
  NPER: ['NPER(rate, pmt, pv, [fv], [type])', 'Number of periods', '期數'],
  DSUM: ['DSUM(database, field, criteria)', 'Sum of matching records', '資料庫條件加總'],
  ISBLANK: ['ISBLANK(value)', 'Is it empty?', '是否空白'],
  ISNUMBER: ['ISNUMBER(value)', 'Is it a number?', '是否為數值'],
  ISTEXT: ['ISTEXT(value)', 'Is it text?', '是否為文字'],
  ISERROR: ['ISERROR(value)', 'Is it an error?', '是否為錯誤'],
  ISNA: ['ISNA(value)', 'Is it #N/A?', '是否為 #N/A'],
  ISLOGICAL: ['ISLOGICAL(value)', 'Is it TRUE/FALSE?', '是否為邏輯值'],
  NA: ['NA()', 'The #N/A error', '傳回 #N/A'],
  TRUE: ['TRUE()', 'TRUE', '邏輯值 TRUE'],
  FALSE: ['FALSE()', 'FALSE', '邏輯值 FALSE'],
};

export interface FunctionHelp {
  name: string;
  signature: string;
  description: string;
}

export function helpFor(name: string): FunctionHelp | null {
  const n = name.toUpperCase();
  if (!(n in FUNCTIONS)) return null;
  const h = HELP[n];
  return { name: n, signature: h?.[0] ?? `${n}(…)`, description: h ? (isZh() ? h[2] : h[1]) : '' };
}

/** Functions whose name starts with the prefix, the common ones first. */
export function suggest(prefix: string, limit = 8): FunctionHelp[] {
  const p = prefix.toUpperCase();
  return Object.keys(FUNCTIONS)
    .filter((n) => n.startsWith(p))
    .sort((a, b) => Number(!HELP[a]) - Number(!HELP[b]) || a.length - b.length || a.localeCompare(b))
    .slice(0, limit)
    .map((n) => helpFor(n)!);
}

export const functionCount = () => Object.keys(FUNCTIONS).length;
