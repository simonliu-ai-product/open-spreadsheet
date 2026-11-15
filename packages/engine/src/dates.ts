/**
 * Excel's 1900 date system. Serial 1 is 1900-01-01, and serial 60 is the
 * 29 February 1900 that never existed: Excel kept Lotus 1-2-3's leap-year bug,
 * so every serial from 61 on is one day later than a plain day count.
 */

const DAY_MS = 86_400_000;
const EPOCH = Date.UTC(1899, 11, 31); // serial 0

export function dateToSerial(year: number, month: number, day: number): number {
  // Excel lets month and day overflow: DATE(2026,14,1) is 2027-02-01.
  const ms = Date.UTC(year, month - 1, day);
  let serial = Math.round((ms - EPOCH) / DAY_MS);
  if (serial >= 60) serial += 1;
  return serial;
}

export interface YMD {
  year: number;
  month: number;
  day: number;
}

export function serialToDate(serial: number): YMD {
  const whole = Math.floor(serial);
  if (whole === 60) return { year: 1900, month: 2, day: 29 };
  const days = whole > 60 ? whole - 1 : whole;
  const d = new Date(EPOCH + days * DAY_MS);
  return { year: d.getUTCFullYear(), month: d.getUTCMonth() + 1, day: d.getUTCDate() };
}

/** Serial for a JavaScript time, in the local time zone, with the time of day as a fraction. */
export function serialFromJsDate(date: Date): number {
  const whole = dateToSerial(date.getFullYear(), date.getMonth() + 1, date.getDate());
  const secs = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
  return whole + secs / 86_400;
}

/** Excel's WEEKDAY, return type 1: Sunday is 1. */
export function weekday(serial: number): number {
  // Serial 1 (1900-01-01) is a Sunday in Excel's calendar, bug included.
  return ((Math.floor(serial) % 7) + 7) % 7 || 7;
}
