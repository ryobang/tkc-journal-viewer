import * as ExcelJS from 'exceljs';

export const TKC_HEADERS_44 = [
  '月日', '伝票番号', '証憑番号',
  '借方科目コード', '借方科目名', '借方補助コード', '借方口座名', '借方部門コード', '借方部門名',
  '借方課税区分', '借方事業区分', '借方消費税額自動計算か否か', '借方軽減税率か否か', '借方税率', '借方控除割合',
  '借方取引金額', '借方消費税等', '借方税抜き金額',
  '貸方科目コード', '貸方科目名', '貸方補助コード', '貸方口座名', '貸方部門コード', '貸方部門名',
  '貸方課税区分', '貸方事業区分', '貸方消費税額自動計算か否か', '貸方軽減税率か否か', '貸方税率', '貸方控除割合',
  '貸方取引金額', '貸方消費税等', '貸方税抜き金額',
  '取引先コード', '取引先名', '取引先の事業者登録番号', '元帳摘要',
  '実際の仕入れ年月日表示区分', '実際の仕入れ年月日１', '実際の仕入れ年月日２',
  '収支区分コード', '収支区分名', '内訳区分コード', '内訳区分名',
];

export interface AccountSide {
  code: string;
  name: string;
  subCode: string;
  subName: string;
  deptCode: string;
  deptName: string;
  taxKbn: string;
  bizKbn: string;
  taxAuto: string;
  taxReduced: string;
  taxRate: string;
  taxControlRate: string;
  amount: number;
  taxAmount: number;
  netAmount: number;
}

export interface JournalEntry {
  rowIndex: number;
  date: string;            // ISO YYYY-MM-DD or empty
  dateRaw: string;
  voucherNo: string;
  evidenceNo: string;
  debit: AccountSide;
  credit: AccountSide;
  partnerCode: string;
  partnerName: string;
  partnerTNumber: string;
  memo: string;
  realPurchaseDate: string;
  incomeKbnCode: string;
  incomeKbnName: string;
  breakdownKbnCode: string;
  breakdownKbnName: string;
}

export interface ParseResult {
  ok: boolean;
  isTkcFormat: boolean;
  sheetName: string;
  headers: string[];
  entries: JournalEntry[];
  totalDebit: number;
  totalCredit: number;
  message?: string;
}

function cellToString(v: ExcelJS.CellValue): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return formatDate(v);
  if (typeof v === 'object') {
    const obj = v as { result?: unknown; text?: unknown; richText?: Array<{ text?: string }>; hyperlink?: string; formula?: string };
    if (obj.result !== undefined && obj.result !== null) return String(obj.result);
    if (typeof obj.text === 'string') return obj.text;
    if (Array.isArray(obj.richText)) return obj.richText.map(r => r.text ?? '').join('');
    if (typeof obj.hyperlink === 'string') return obj.hyperlink;
    if (typeof obj.formula === 'string') return obj.formula;
    return '';
  }
  return String(v).trim();
}

function cellToNumber(v: ExcelJS.CellValue): number {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const cleaned = v.replace(/[, ¥]/g, '');
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }
  if (typeof v === 'object') {
    const obj = v as { result?: unknown };
    if (typeof obj.result === 'number') return obj.result;
    if (typeof obj.result === 'string') {
      const n = Number(obj.result);
      return Number.isFinite(n) ? n : 0;
    }
  }
  return 0;
}

function cellToDateIso(v: ExcelJS.CellValue): { iso: string; raw: string } {
  if (v === null || v === undefined || v === '') return { iso: '', raw: '' };
  if (v instanceof Date) {
    return { iso: toIso(v), raw: formatDate(v) };
  }
  if (typeof v === 'number') {
    // Excel serial date
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(epoch.getTime() + v * 24 * 60 * 60 * 1000);
    return { iso: toIso(d), raw: formatDate(d) };
  }
  if (typeof v === 'string') {
    const t = v.trim();
    const m = t.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (m) {
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      return { iso: toIso(d), raw: t };
    }
    return { iso: '', raw: t };
  }
  if (typeof v === 'object') {
    const obj = v as { result?: unknown; text?: unknown };
    if (obj.result instanceof Date) return cellToDateIso(obj.result);
    if (typeof obj.result === 'string' || typeof obj.result === 'number') return cellToDateIso(obj.result as ExcelJS.CellValue);
    if (typeof obj.text === 'string') return cellToDateIso(obj.text);
  }
  return { iso: '', raw: String(v) };
}

function toIso(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function formatDate(d: Date): string {
  return toIso(d);
}

function headersMatchTkc(headers: string[]): boolean {
  // Check at least the critical headers exist
  const required = ['月日', '借方科目コード', '借方科目名', '貸方科目コード', '貸方科目名', '元帳摘要'];
  for (const r of required) {
    if (!headers.includes(r)) return false;
  }
  let matched = 0;
  for (const h of TKC_HEADERS_44) {
    if (headers.includes(h)) matched++;
  }
  return matched >= 25; // tolerate column renames between xlsx(44列) and CSV(仕訳帳データ)
}

/**
 * 抽出済みのヘッダ＋行データ(セル値の二次元配列)から仕訳を組み立てる。
 * xlsx(44列)・CSV(仕訳帳データ49列) 双方で共通利用する。
 */
function buildResult(
  sheetName: string,
  headers: string[],
  rows: ExcelJS.CellValue[][],
): ParseResult {
  if (!headersMatchTkc(headers)) {
    return {
      ok: true, isTkcFormat: false, sheetName, headers, entries: [],
      totalDebit: 0, totalCredit: 0,
      message: 'TKC仕訳フォーマット(44列xlsx / 仕訳帳データCSV)ではありません',
    };
  }

  const colIndex = new Map<string, number>();
  headers.forEach((h, i) => {
    if (!colIndex.has(h)) colIndex.set(h, i);
  });
  // 先に挙げた列名から順に探し、最初に見つかった列の値を返す(別名フォールバック)
  const get = (row: ExcelJS.CellValue[], ...names: string[]): ExcelJS.CellValue => {
    for (const name of names) {
      const i = colIndex.get(name);
      if (i !== undefined) return row[i] ?? null;
    }
    return null;
  };

  const entries: JournalEntry[] = [];
  let totalDebit = 0;
  let totalCredit = 0;

  rows.forEach((row, idx) => {
    const dateCell = get(row, '月日');
    const debitCode = cellToString(get(row, '借方科目コード'));
    const creditCode = cellToString(get(row, '貸方科目コード'));
    if (!dateCell && !debitCode && !creditCode) return;

    const { iso, raw } = cellToDateIso(dateCell);

    const debit: AccountSide = {
      code: debitCode,
      name: cellToString(get(row, '借方科目名')),
      subCode: cellToString(get(row, '借方補助コード')),
      subName: cellToString(get(row, '借方口座名')),
      deptCode: cellToString(get(row, '借方部門コード')),
      deptName: cellToString(get(row, '借方部門名')),
      taxKbn: cellToString(get(row, '借方課税区分')),
      bizKbn: cellToString(get(row, '借方事業区分')),
      taxAuto: cellToString(get(row, '借方消費税額自動計算か否か')),
      taxReduced: cellToString(get(row, '借方軽減税率か否か')),
      taxRate: cellToString(get(row, '借方税率')),
      taxControlRate: cellToString(get(row, '借方控除割合')),
      amount: cellToNumber(get(row, '借方取引金額')),
      taxAmount: cellToNumber(get(row, '借方消費税等')),
      netAmount: cellToNumber(get(row, '借方税抜き金額')),
    };
    const credit: AccountSide = {
      code: creditCode,
      name: cellToString(get(row, '貸方科目名')),
      subCode: cellToString(get(row, '貸方補助コード')),
      subName: cellToString(get(row, '貸方口座名')),
      deptCode: cellToString(get(row, '貸方部門コード')),
      deptName: cellToString(get(row, '貸方部門名')),
      taxKbn: cellToString(get(row, '貸方課税区分')),
      bizKbn: cellToString(get(row, '貸方事業区分')),
      taxAuto: cellToString(get(row, '貸方消費税額自動計算か否か')),
      taxReduced: cellToString(get(row, '貸方軽減税率か否か')),
      taxRate: cellToString(get(row, '貸方税率')),
      taxControlRate: cellToString(get(row, '貸方控除割合')),
      amount: cellToNumber(get(row, '貸方取引金額')),
      taxAmount: cellToNumber(get(row, '貸方消費税等')),
      netAmount: cellToNumber(get(row, '貸方税抜き金額')),
    };

    const realPurchase = cellToDateIso(get(row, '実際の仕入れ年月日１', '実際の仕入れ年月日')).iso;
    // 伝票番号が無いCSV(仕訳帳データ)では入力番号で複合仕訳をまとめる
    const voucherNo = cellToString(get(row, '伝票番号')) || cellToString(get(row, '入力番号'));

    entries.push({
      rowIndex: idx + 2,
      date: iso,
      dateRaw: raw,
      voucherNo,
      evidenceNo: cellToString(get(row, '証憑番号')).trim(),
      debit,
      credit,
      partnerCode: cellToString(get(row, '取引先コード')),
      partnerName: cellToString(get(row, '取引先名')),
      // インボイス登録番号(T番号)。様式ごとに列名が異なるため別名で吸収する
      // xlsx読込: 適格請求書発行事業者の登録番号 / CSV仕訳帳: 適格請求書発行事業者 / 44列定数: 取引先の事業者登録番号
      partnerTNumber: cellToString(get(row, '取引先の事業者登録番号', '適格請求書発行事業者の登録番号', '適格請求書発行事業者')),
      memo: cellToString(get(row, '元帳摘要')),
      realPurchaseDate: realPurchase,
      incomeKbnCode: cellToString(get(row, '収支区分コード')),
      incomeKbnName: cellToString(get(row, '収支区分名')),
      breakdownKbnCode: cellToString(get(row, '内訳区分コード')),
      breakdownKbnName: cellToString(get(row, '内訳区分名')),
    });
    totalDebit += debit.amount;
    totalCredit += credit.amount;
  });

  return {
    ok: true,
    isTkcFormat: true,
    sheetName,
    headers,
    entries,
    totalDebit,
    totalCredit,
  };
}

/** 拡張子に応じて xlsx / csv を判別して解析する。 */
export async function parseTkc(buffer: Buffer, fileName: string): Promise<ParseResult> {
  if (/\.csv$/i.test(fileName)) {
    return parseTkcCsv(buffer);
  }
  return parseTkcXlsx(buffer);
}

export async function parseTkcXlsx(buffer: Buffer): Promise<ParseResult> {
  const wb = new ExcelJS.Workbook();
  try {
    const ab = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    await wb.xlsx.load(ab as ArrayBuffer);
  } catch (e) {
    return {
      ok: false,
      isTkcFormat: false,
      sheetName: '',
      headers: [],
      entries: [],
      totalDebit: 0,
      totalCredit: 0,
      message: `Excelファイルを開けませんでした: ${(e as Error).message}`,
    };
  }

  const ws = wb.worksheets[0];
  if (!ws) {
    return {
      ok: false, isTkcFormat: false, sheetName: '', headers: [], entries: [],
      totalDebit: 0, totalCredit: 0, message: 'シートが見つかりません',
    };
  }

  const lastCol = ws.actualColumnCount || ws.columnCount;
  const headerRow = ws.getRow(1);
  const headers: string[] = [];
  for (let c = 1; c <= lastCol; c++) {
    headers.push(cellToString(headerRow.getCell(c).value).trim());
  }

  const rows: ExcelJS.CellValue[][] = [];
  const lastRow = ws.actualRowCount || ws.rowCount;
  for (let r = 2; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const values: ExcelJS.CellValue[] = [];
    for (let c = 1; c <= lastCol; c++) {
      values.push(row.getCell(c).value);
    }
    rows.push(values);
  }

  return buildResult(ws.name, headers, rows);
}

export function parseTkcCsv(buffer: Buffer): ParseResult {
  let table: string[][];
  try {
    table = parseCsv(decodeText(buffer));
  } catch (e) {
    return {
      ok: false, isTkcFormat: false, sheetName: '', headers: [], entries: [],
      totalDebit: 0, totalCredit: 0,
      message: `CSVファイルを開けませんでした: ${(e as Error).message}`,
    };
  }
  if (table.length === 0) {
    return {
      ok: true, isTkcFormat: false, sheetName: 'CSV', headers: [], entries: [],
      totalDebit: 0, totalCredit: 0, message: '空のCSVファイルです',
    };
  }
  const headers = table[0].map(h => h.trim());
  const rows: ExcelJS.CellValue[][] = table.slice(1);
  return buildResult('仕訳帳データ', headers, rows);
}

/** UTF-8(BOM有無)優先、不可なら Shift_JIS としてデコードする。 */
function decodeText(buffer: Buffer): string {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return buffer.toString('utf8', 3);
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buffer);
  } catch {
    try {
      return new TextDecoder('shift_jis').decode(buffer);
    } catch {
      return buffer.toString('utf8');
    }
  }
}

/** RFC4180 準拠の最小CSVパーサ(ダブルクォート・改行埋め込み対応)。 */
function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else { inQuotes = false; }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"') { inQuotes = true; }
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\r') { /* skip */ }
    else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
    else { field += ch; }
  }
  if (field !== '' || row.length > 0) { row.push(field); rows.push(row); }
  // 完全に空の行は除外する
  return rows.filter(r => r.some(c => c.trim() !== ''));
}
