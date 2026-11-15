/** Starting points offered on the home page. Each is an ordinary workbook. */
import type { WorkbookData } from './types.ts';

export interface Template {
  id: string;
  title: { en: string; zh: string };
  description: { en: string; zh: string };
  build: (zh: boolean) => WorkbookData;
}

const header = { b: true, bg: '#e8f0e9' };

export const TEMPLATES: Template[] = [
  {
    id: 'budget',
    title: { en: 'Monthly budget', zh: '每月預算' },
    description: { en: 'Income, spending by category, and what is left', zh: '收入、各類支出與結餘' },
    build: (zh) => ({
      title: zh ? '每月預算' : 'Monthly budget',
      sheets: [
        {
          name: zh ? '預算' : 'Budget',
          cells: {
            A1: zh ? '每月預算' : 'Monthly budget',
            A3: zh ? '項目' : 'Item', B3: zh ? '預算' : 'Budget', C3: zh ? '實際' : 'Actual', D3: zh ? '差額' : 'Difference',
            A4: zh ? '房租' : 'Rent', B4: '22000', C4: '22000',
            A5: zh ? '餐飲' : 'Food', B5: '10000', C5: '11250',
            A6: zh ? '交通' : 'Transport', B6: '2000', C6: '1680',
            A7: zh ? '娛樂' : 'Fun', B7: '3000', C7: '3600',
            A8: zh ? '儲蓄' : 'Savings', B8: '8000', C8: '8000',
            A9: zh ? '合計' : 'Total', B9: '=SUM(B4:B8)', C9: '=SUM(C4:C8)', D9: '=SUM(D4:D8)',
            A11: zh ? '收入' : 'Income', B11: '50000',
            A12: zh ? '結餘' : 'Left over', B12: '=B11-C9',
            D4: '=B4-C4', D5: '=B5-C5', D6: '=B6-C6', D7: '=B7-C7', D8: '=B8-C8',
          },
          cols: { A: 140 },
          styles: {
            A1: { b: true },
            A3: header, B3: header, C3: header, D3: header,
            ...Object.fromEntries(['B', 'C', 'D'].flatMap((c) => [4, 5, 6, 7, 8, 9].map((r) => [`${c}${r}`, { fmt: '#,##0;[Red]-#,##0' }]))),
            A9: { b: true }, B9: { b: true, fmt: '#,##0' }, C9: { b: true, fmt: '#,##0' }, D9: { b: true, fmt: '#,##0;[Red]-#,##0' },
            B11: { fmt: '#,##0' }, A12: { b: true }, B12: { b: true, fmt: '#,##0;[Red]-#,##0' },
          },
          freeze: { rows: 3, cols: 0 },
        },
      ],
    }),
  },
  {
    id: 'sales',
    title: { en: 'Sales tracker', zh: '銷售紀錄' },
    description: { en: 'Orders with totals per person and product', zh: '訂單明細，依業務與產品統計' },
    build: (zh) => {
      const people = zh ? ['小林', '阿美', '大偉'] : ['Lin', 'May', 'Dave'];
      const products = zh ? ['筆電', '螢幕', '鍵盤'] : ['Laptop', 'Monitor', 'Keyboard'];
      const prices = [32000, 6800, 1500];
      const cells: Record<string, string> = {
        A1: zh ? '日期' : 'Date', B1: zh ? '業務' : 'Rep', C1: zh ? '產品' : 'Product', D1: zh ? '數量' : 'Qty', E1: zh ? '單價' : 'Price', F1: zh ? '金額' : 'Amount',
        H1: zh ? '業務' : 'Rep', I1: zh ? '業績' : 'Sales', J1: zh ? '占比' : 'Share',
      };
      for (let i = 0; i < 12; i++) {
        const r = i + 2;
        const p = (i * 7) % 3;
        cells[`A${r}`] = `2026-${String(1 + Math.floor(i / 4)).padStart(2, '0')}-${String(3 + (i % 4) * 7).padStart(2, '0')}`;
        cells[`B${r}`] = people[i % 3]!;
        cells[`C${r}`] = products[p]!;
        cells[`D${r}`] = String(1 + ((i * 5) % 9));
        cells[`E${r}`] = String(prices[p]);
        cells[`F${r}`] = `=D${r}*E${r}`;
      }
      people.forEach((n, i) => {
        cells[`H${i + 2}`] = n;
        cells[`I${i + 2}`] = `=SUMIF($B$2:$B$13,H${i + 2},$F$2:$F$13)`;
        cells[`J${i + 2}`] = `=I${i + 2}/SUM($I$2:$I$4)`;
      });
      cells.H6 = zh ? '總計' : 'Total';
      cells.I6 = '=SUM(F2:F13)';
      cells.H7 = zh ? '冠軍' : 'Top rep';
      cells.I7 = '=INDEX(H2:H4,MATCH(MAX(I2:I4),I2:I4,0))';
      const styles: Record<string, object> = {};
      for (const c of ['A', 'B', 'C', 'D', 'E', 'F', 'H', 'I', 'J']) styles[`${c}1`] = header;
      for (let r = 2; r <= 13; r++) {
        styles[`E${r}`] = { fmt: '#,##0' };
        styles[`F${r}`] = { fmt: '#,##0' };
      }
      for (let r = 2; r <= 6; r++) styles[`I${r}`] = { fmt: '#,##0' };
      for (let r = 2; r <= 4; r++) styles[`J${r}`] = { fmt: '0.0%' };
      styles.H6 = { b: true };
      styles.I6 = { b: true, fmt: '#,##0' };
      return { title: zh ? '銷售紀錄' : 'Sales tracker', sheets: [{ name: zh ? '訂單' : 'Orders', cells, styles, freeze: { rows: 1, cols: 0 }, cols: { A: 110 } }] };
    },
  },
  {
    id: 'project',
    title: { en: 'Project plan', zh: '專案進度' },
    description: { en: 'Tasks, owners, dates and progress', zh: '工作項目、負責人、日期與進度' },
    build: (zh) => {
      const tasks = zh ? ['需求訪談', '規格撰寫', '介面設計', '開發', '測試', '上線'] : ['Interviews', 'Spec', 'Design', 'Build', 'Test', 'Launch'];
      const cells: Record<string, string> = {
        A1: zh ? '工作項目' : 'Task', B1: zh ? '負責人' : 'Owner', C1: zh ? '開始' : 'Start', D1: zh ? '天數' : 'Days', E1: zh ? '結束' : 'End', F1: zh ? '進度' : 'Progress', G1: zh ? '狀態' : 'Status',
      };
      tasks.forEach((t, i) => {
        const r = i + 2;
        cells[`A${r}`] = t;
        cells[`B${r}`] = (zh ? ['Simon', 'Amy', 'Ken'] : ['Simon', 'Amy', 'Ken'])[i % 3]!;
        cells[`C${r}`] = i === 0 ? '2026-12-01' : `=E${r - 1}+1`;
        cells[`D${r}`] = String([5, 7, 10, 20, 8, 2][i]);
        cells[`E${r}`] = `=C${r}+D${r}-1`;
        cells[`F${r}`] = String([1, 1, 0.6, 0.2, 0, 0][i]);
        cells[`G${r}`] = `=IF(F${r}>=1,"${zh ? '完成' : 'Done'}",IF(F${r}>0,"${zh ? '進行中' : 'In progress'}","${zh ? '未開始' : 'Not started'}"))`;
      });
      cells.A9 = zh ? '整體進度' : 'Overall';
      cells.F9 = '=SUMPRODUCT(D2:D7,F2:F7)/SUM(D2:D7)';
      const styles: Record<string, object> = {};
      for (const c of 'ABCDEFG') styles[`${c}1`] = header;
      for (let r = 2; r <= 7; r++) {
        styles[`C${r}`] = { fmt: 'yyyy-mm-dd' };
        styles[`E${r}`] = { fmt: 'yyyy-mm-dd' };
        styles[`F${r}`] = { fmt: '0%' };
      }
      styles.A9 = { b: true };
      styles.F9 = { b: true, fmt: '0%' };
      return { title: zh ? '專案進度' : 'Project plan', sheets: [{ name: zh ? '時程' : 'Schedule', cells, styles, cols: { A: 120, C: 100, E: 100 }, freeze: { rows: 1, cols: 1 } }] };
    },
  },
  {
    id: 'invoice',
    title: { en: 'Invoice', zh: '報價單' },
    description: { en: 'Line items with tax and total', zh: '品項、稅額與總計' },
    build: (zh) => ({
      title: zh ? '報價單' : 'Invoice',
      sheets: [
        {
          name: zh ? '報價單' : 'Invoice',
          cells: {
            A1: zh ? '報價單' : 'INVOICE',
            A3: zh ? '客戶' : 'Client', B3: zh ? '範例公司' : 'Example Co.',
            A4: zh ? '日期' : 'Date', B4: '=TODAY()',
            A6: zh ? '品項' : 'Item', B6: zh ? '數量' : 'Qty', C6: zh ? '單價' : 'Unit price', D6: zh ? '小計' : 'Line total',
            A7: zh ? '網站設計' : 'Website design', B7: '1', C7: '60000', D7: '=B7*C7',
            A8: zh ? '主機一年' : 'Hosting (1 year)', B8: '1', C8: '12000', D8: '=B8*C8',
            A9: zh ? '教育訓練（小時）' : 'Training (hours)', B9: '6', C9: '2500', D9: '=B9*C9',
            C11: zh ? '未稅合計' : 'Subtotal', D11: '=SUM(D7:D9)',
            C12: zh ? '營業稅 5%' : 'Tax 5%', D12: '=ROUND(D11*5%,0)',
            C13: zh ? '總計' : 'Total', D13: '=D11+D12',
          },
          cols: { A: 180, C: 110, D: 120 },
          merges: ['A1:D1'],
          styles: {
            A1: { b: true, align: 'center', bg: '#217346', color: '#ffffff' },
            B4: { fmt: 'yyyy-mm-dd', align: 'left' },
            A6: header, B6: header, C6: header, D6: header,
            C7: { fmt: '#,##0' }, C8: { fmt: '#,##0' }, C9: { fmt: '#,##0' },
            D7: { fmt: '#,##0' }, D8: { fmt: '#,##0' }, D9: { fmt: '#,##0' },
            D11: { fmt: '#,##0' }, D12: { fmt: '#,##0' },
            C13: { b: true }, D13: { b: true, fmt: '"NT$"#,##0' },
          },
        },
      ],
    }),
  },
];
