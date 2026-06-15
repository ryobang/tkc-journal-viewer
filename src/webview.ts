import { JournalEntry, ParseResult } from './parser';

export interface RenderOptions {
  fileName: string;
  result: ParseResult;
  cspSource: string;
  config: { dateFormat: string; zeroAmount: string };
  /** ポップアップ(別ウィンドウ)で開いたときに引き継ぐ表示状態 */
  initial?: { filter?: string; month?: string; fontPx?: number };
}

export function renderHtml(opts: RenderOptions): string {
  const { fileName, result, cspSource, config } = opts;

  if (!result.ok) {
    return wrapError(cspSource, fileName, result.message ?? '読み込みエラー');
  }
  if (!result.isTkcFormat) {
    return wrapNotice(cspSource, fileName, result);
  }

  const data = JSON.stringify({
    fileName,
    sheetName: result.sheetName,
    entries: result.entries,
    totalDebit: result.totalDebit,
    totalCredit: result.totalCredit,
    config,
    initial: opts.initial ?? {},
  });

  const nonce = makeNonce();
  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy"
  content="default-src 'none'; style-src ${cspSource} 'unsafe-inline'; script-src 'nonce-${nonce}';">
<title>${escapeHtml(fileName)} - TKC仕訳帳</title>
<style>${baseCss()}</style>
</head>
<body>
<div id="app">
  <header class="bar">
    <div class="title">
      <span class="file">${escapeHtml(fileName)}</span>
      <span class="sheet">[${escapeHtml(result.sheetName)}]</span>
    </div>
    <div class="actions">
      <input id="filter" type="text" placeholder="検索 (科目/取引先/摘要/伝票番号)" />
      <select id="monthFilter"><option value="">全月</option></select>
      <button id="reload" title="ファイルを再読込">再読込</button>
      <button id="openExcel" title="Excelでこのファイルを開く">Excelで開く</button>
      <button id="popout" title="表示中の内容を別ウィンドウで開く(別モニターへ移動可能)">別ウィンドウ</button>
      <span class="fontsize">
        <button id="fontDec" title="文字を小さく">A−</button>
        <span class="lvl" id="fontLvl">13px</span>
        <button id="fontInc" title="文字を大きく">A＋</button>
        <button id="fontReset" title="文字サイズを既定に戻す">既定</button>
      </span>
    </div>
  </header>
  <div class="summary" id="summary"></div>
  <div class="balance-warn" id="balanceWarn" hidden></div>
  <div class="tableWrap">
    <table id="journal">
      <thead>
        <tr>
          <th class="c-no">No</th>
          <th class="c-date">月日</th>
          <th class="c-vno">伝票/証憑</th>
          <th class="c-acct">借方科目</th>
          <th class="c-amt">借方金額</th>
          <th class="c-tax">課税</th>
          <th class="c-acct">貸方科目</th>
          <th class="c-amt">貸方金額</th>
          <th class="c-tax">課税</th>
          <th class="c-partner">取引先</th>
          <th class="c-memo">摘要</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
      <tfoot>
        <tr>
          <td colspan="4" class="t-label">合計</td>
          <td class="amt" id="ftDebit">0</td>
          <td></td>
          <td></td>
          <td class="amt" id="ftCredit">0</td>
          <td colspan="3"></td>
        </tr>
      </tfoot>
    </table>
  </div>
  <div class="empty" id="empty" hidden>表示できる仕訳がありません</div>
</div>
<script nonce="${nonce}">
window.__TKC__ = ${data};
${clientScript()}
</script>
</body>
</html>`;
}

function wrapError(cspSource: string, fileName: string, msg: string): string {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'unsafe-inline';">
<style>${baseCss()}</style></head>
<body><div class="notice err"><h2>${escapeHtml(fileName)}</h2><p>${escapeHtml(msg)}</p></div></body></html>`;
}

function wrapNotice(cspSource: string, fileName: string, result: ParseResult): string {
  const headersList = result.headers.slice(0, 50).map(h => `<li>${escapeHtml(h)}</li>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src ${cspSource} 'unsafe-inline';">
<style>${baseCss()}</style></head>
<body><div class="notice">
  <h2>${escapeHtml(fileName)}</h2>
  <p>${escapeHtml(result.message ?? '対応していないフォーマットです')}</p>
  <p>このファイルは TKC 読み込み用 Excel仕訳 (44列の固定ヘッダ) ではないため、仕訳帳ビューでは表示できません。</p>
  <details><summary>検出されたヘッダ (${result.headers.length}列)</summary>
    <ol>${headersList}</ol>
  </details>
</div></body></html>`;
}

function baseCss(): string {
  return `
:root { color-scheme: light dark; }
* { box-sizing: border-box; }
html,body { margin:0; padding:0; height:100%; font-family: "Segoe UI", "Yu Gothic UI", "Meiryo", system-ui, sans-serif; color: var(--vscode-foreground); background: var(--vscode-editor-background); }
body { font-size: var(--tkc-font, 13px); }
.fontsize { display:flex; gap:2px; align-items:center; }
.fontsize button { padding:4px 8px; background: var(--vscode-button-secondaryBackground, var(--vscode-button-background)); color: var(--vscode-button-secondaryForeground, var(--vscode-button-foreground)); border:none; border-radius:3px; cursor:pointer; font-size:.9em; line-height:1; }
.fontsize .lvl { min-width:42px; text-align:center; opacity:.7; font-size:.8em; font-variant-numeric: tabular-nums; }
#app { display:flex; flex-direction:column; height:100vh; }
.bar { display:flex; align-items:center; gap:12px; padding:8px 12px; border-bottom:1px solid var(--vscode-panel-border); flex-wrap:wrap; }
.title { font-weight:600; flex:1 1 auto; min-width:200px; }
.title .sheet { font-weight:400; opacity:.6; margin-left:8px; font-size:.85em; }
.actions { display:flex; gap:6px; align-items:center; flex-wrap:wrap; }
.actions input[type=text], .actions select { padding:4px 8px; background: var(--vscode-input-background); color: var(--vscode-input-foreground); border:1px solid var(--vscode-input-border, transparent); border-radius:3px; min-width:200px; font-size:.9em; }
.actions select { min-width:auto; }
.actions button { padding:4px 10px; background: var(--vscode-button-background); color: var(--vscode-button-foreground); border:none; border-radius:3px; cursor:pointer; font-size:.9em; }
.actions button:hover { background: var(--vscode-button-hoverBackground); }
.chk { display:flex; align-items:center; gap:4px; font-size:.9em; cursor:pointer; user-select:none; }
.summary { padding:6px 12px; font-size:.85em; opacity:.85; border-bottom:1px solid var(--vscode-panel-border); display:flex; gap:16px; flex-wrap:wrap; }
.summary span.k { opacity:.6; margin-right:4px; }
.balance-warn { padding:6px 12px; background: var(--vscode-inputValidation-warningBackground, #5a4900); color: var(--vscode-inputValidation-warningForeground, inherit); border-bottom:1px solid var(--vscode-inputValidation-warningBorder, #b89500); font-size:.9em; }
.tableWrap { flex:1 1 auto; overflow:auto; }
table { border-collapse: collapse; width:100%; font-size:.9em; }
thead th { position:sticky; top:0; background: var(--vscode-editorWidget-background, var(--vscode-editor-background)); border-bottom:2px solid var(--vscode-panel-border); padding:6px 8px; text-align:left; font-weight:600; white-space:nowrap; z-index:1; }
tbody td, tfoot td { padding:5px 8px; border-bottom:1px solid var(--vscode-panel-border); vertical-align: top; }
tbody tr:hover { background: var(--vscode-list-hoverBackground); }
tbody tr.month-break td { border-top:2px solid var(--vscode-panel-border); }
.c-no { width:48px; text-align:right; opacity:.6; font-variant-numeric: tabular-nums; }
.c-date { width:80px; white-space:nowrap; font-variant-numeric: tabular-nums; }
.c-vno { width:54px; font-size:.85em; opacity:.85; font-variant-numeric: tabular-nums; text-align:right; padding-right:6px; }
.c-vno .ev { display:block; opacity:.55; font-size:.85em; }
.c-acct { width:200px; }
.c-amt { width:110px; text-align:right; font-variant-numeric: tabular-nums; white-space:nowrap; }
.c-tax { width:50px; text-align:center; font-size:.85em; opacity:.85; }
.c-partner { width:180px; }
.c-memo { min-width:200px; }
td.amt { text-align:right; font-variant-numeric: tabular-nums; white-space:nowrap; }
td.acct .code { opacity:.55; font-size:.85em; margin-right:4px; font-variant-numeric: tabular-nums; }
td.acct .sub { display:block; opacity:.7; font-size:.85em; margin-top:1px; }
td.acct .dept { display:block; opacity:.6; font-size:.8em; margin-top:1px; }
td.tax .auto { font-size:.7em; opacity:.6; display:block; }
td.tax .rate8 { display:block; font-size:.85em; opacity:1; font-weight:600; color: var(--vscode-charts-orange, var(--vscode-foreground)); }
.tnum { display:block; opacity:.55; font-size:.75em; font-variant-numeric: tabular-nums; }
.empty { padding:24px; text-align:center; opacity:.6; }
.notice { padding:24px; max-width:760px; }
.notice.err { color: var(--vscode-errorForeground); }
.notice h2 { margin-top:0; }
tfoot td.t-label { text-align:right; font-weight:600; }
tfoot td.amt { font-weight:600; border-top:2px solid var(--vscode-panel-border); }
.zero { opacity:.35; }
tbody tr.cmpx-first td { border-top: 2px solid color-mix(in srgb, var(--vscode-textLink-foreground) 55%, transparent); }
tbody tr.cmpx-last td { border-bottom: 2px solid color-mix(in srgb, var(--vscode-textLink-foreground) 55%, transparent); }
tbody tr.cmpx-mid td.c-date, tbody tr.cmpx-mid td.c-vno, tbody tr.cmpx-last td.c-date, tbody tr.cmpx-last td.c-vno { color: transparent; }
.cmpx-badge { display:inline-block; margin-left:4px; padding:0 5px; font-size:.7em; background: var(--vscode-textLink-foreground); color: var(--vscode-editor-background); border-radius:3px; vertical-align:middle; font-weight:600; opacity:.85; }
`;
}

function clientScript(): string {
  // This runs in the webview, not in Node.
  return `
(function(){
  const vscode = acquireVsCodeApi();
  const data = window.__TKC__;
  const cfg = data.config || { dateFormat: 'MM/DD', zeroAmount: 'blank' };
  const entries = data.entries || [];

  const $ = (id) => document.getElementById(id);
  const filterInput = $('filter');
  const monthSel = $('monthFilter');
  const tbody = $('tbody');
  const summary = $('summary');
  const balanceWarn = $('balanceWarn');
  const empty = $('empty');
  const ftDebit = $('ftDebit');
  const ftCredit = $('ftCredit');

  function fmtAmount(n) {
    if (!n) {
      if (cfg.zeroAmount === 'blank') return '';
      if (cfg.zeroAmount === 'dash') return '—';
      return '0';
    }
    return Number(n).toLocaleString('ja-JP');
  }
  function fmtDate(iso, raw) {
    if (!iso) return raw || '';
    const m = iso.match(/^(\\d{4})-(\\d{2})-(\\d{2})/);
    if (!m) return raw || iso;
    if (cfg.dateFormat === 'YYYY-MM-DD') return iso;
    if (cfg.dateFormat === 'M/D') return Number(m[2]) + '/' + Number(m[3]);
    return m[2] + '/' + m[3];
  }
  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>\"']/g, (c) => ({
      '&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'
    }[c]));
  }

  // Populate month filter
  const months = Array.from(new Set(entries.map(e => e.date ? e.date.slice(0,7) : ''))).filter(Boolean).sort();
  for (const m of months) {
    const opt = document.createElement('option');
    opt.value = m; opt.textContent = m;
    monthSel.appendChild(opt);
  }

  function matchEntry(e, q, month) {
    if (month && (!e.date || !e.date.startsWith(month))) return false;
    if (!q) return true;
    const ql = q.toLowerCase();
    const blob = [
      e.voucherNo, e.evidenceNo,
      e.debit.code, e.debit.name, e.debit.subName, e.debit.deptName,
      e.credit.code, e.credit.name, e.credit.subName, e.credit.deptName,
      e.partnerCode, e.partnerName, e.partnerTNumber, e.memo,
    ].join(' ').toLowerCase();
    return blob.indexOf(ql) >= 0;
  }

  function renderRows() {
    const q = filterInput.value.trim();
    const month = monthSel.value;
    const filtered = entries.filter(e => matchEntry(e, q, month));

    // 複合仕訳グルーピング (date + voucherNo で連続する行)
    const grpKey = (e) => (e.date || e.dateRaw || '') + '|' + (e.voucherNo || '');
    const grpInfo = filtered.map(() => ({ size: 1, pos: 0, debitSum: 0, creditSum: 0 }));
    let i = 0;
    while (i < filtered.length) {
      const k = grpKey(filtered[i]);
      let j = i;
      let dSum = 0, cSum = 0;
      while (j < filtered.length && grpKey(filtered[j]) === k && filtered[j].voucherNo) {
        dSum += filtered[j].debit.amount || 0;
        cSum += filtered[j].credit.amount || 0;
        j++;
      }
      const size = j - i;
      const stop = size >= 1 ? j : i + 1;
      const realSize = Math.max(size, 1);
      for (let p = 0; p < realSize; p++) {
        grpInfo[i + p] = { size: realSize, pos: p, debitSum: dSum, creditSum: cSum };
      }
      i = stop;
    }

    let html = '';
    let totalDebit = 0, totalCredit = 0;
    let prevMonth = '';
    filtered.forEach((e, idx) => {
      const m = e.date ? e.date.slice(0,7) : '';
      const isBreak = idx > 0 && m !== prevMonth;
      prevMonth = m;
      totalDebit += e.debit.amount || 0;
      totalCredit += e.credit.amount || 0;
      const gi = grpInfo[idx];
      const isCmpx = gi.size > 1;
      const cmpxClass = isCmpx
        ? ' cmpx' + (gi.pos === 0 ? ' cmpx-first' : (gi.pos === gi.size - 1 ? ' cmpx-last' : ' cmpx-mid'))
        : '';

      const debitAcct = formatAcct(e.debit);
      const creditAcct = formatAcct(e.credit);
      const pCode = (e.partnerCode || '').trim();
      const pTNum = (e.partnerTNumber || '').trim();
      const pSub = pCode || pTNum;  // 取引先コードがあればそれ、なければインボイス番号(T番号)
      const partnerHtml = (e.partnerName || pSub)
        ? escapeHtml(e.partnerName || '') + (pSub ? '<span class=\"tnum\">' + escapeHtml(pSub) + '</span>' : '')
        : '';
      const vno = e.voucherNo ? escapeHtml(String(e.voucherNo).trim()) : '';
      const evRaw = (e.evidenceNo || '').trim();
      const ev = evRaw ? '<span class=\"ev\">' + escapeHtml(evRaw) + '</span>' : '';

      const trClass = [isBreak ? 'month-break' : '', cmpxClass.trim()].filter(Boolean).join(' ');
      const cmpxBadge = (isCmpx && gi.pos === 0)
        ? '<span class=\"cmpx-badge\" title=\"複合仕訳 ' + gi.size + '行 / 借 ' + gi.debitSum.toLocaleString('ja-JP') + ' 貸 ' + gi.creditSum.toLocaleString('ja-JP') + '\">複合 ' + gi.size + '</span>'
        : '';

      html += '<tr' + (trClass ? ' class=\"' + trClass + '\"' : '') + '>'
        + '<td class=\"c-no\">' + (idx + 1) + cmpxBadge + '</td>'
        + '<td class=\"c-date\">' + escapeHtml(fmtDate(e.date, e.dateRaw)) + '</td>'
        + '<td class=\"c-vno\">' + vno + ev + '</td>'
        + '<td class=\"c-acct acct\">' + debitAcct + '</td>'
        + '<td class=\"c-amt amt' + (e.debit.amount ? '' : ' zero') + '\">' + escapeHtml(fmtAmount(e.debit.amount)) + '</td>'
        + '<td class=\"c-tax tax\">' + escapeHtml(e.debit.taxKbn || '')
          + (function(){ const l = fmtTaxRateLabel(e.debit); return l ? '<span class=\"rate8\">' + escapeHtml(l) + '</span>' : ''; })()
          + (e.debit.taxAmount ? '<span class=\"auto\">税 ' + Number(e.debit.taxAmount).toLocaleString('ja-JP') + '</span>' : '')
          + '</td>'
        + '<td class=\"c-acct acct\">' + creditAcct + '</td>'
        + '<td class=\"c-amt amt' + (e.credit.amount ? '' : ' zero') + '\">' + escapeHtml(fmtAmount(e.credit.amount)) + '</td>'
        + '<td class=\"c-tax tax\">' + escapeHtml(e.credit.taxKbn || '')
          + (function(){ const l = fmtTaxRateLabel(e.credit); return l ? '<span class=\"rate8\">' + escapeHtml(l) + '</span>' : ''; })()
          + (e.credit.taxAmount ? '<span class=\"auto\">税 ' + Number(e.credit.taxAmount).toLocaleString('ja-JP') + '</span>' : '')
          + '</td>'
        + '<td class=\"c-partner\">' + partnerHtml + '</td>'
        + '<td class=\"c-memo\">' + escapeHtml(e.memo || '') + '</td>'
        + '</tr>';

    });

    tbody.innerHTML = html;
    ftDebit.textContent = totalDebit.toLocaleString('ja-JP');
    ftCredit.textContent = totalCredit.toLocaleString('ja-JP');
    empty.hidden = filtered.length > 0;

    // Summary
    summary.innerHTML =
      '<div><span class=\"k\">表示</span>' + filtered.length + ' / ' + entries.length + ' 件</div>'
      + '<div><span class=\"k\">借方計</span>' + totalDebit.toLocaleString('ja-JP') + '</div>'
      + '<div><span class=\"k\">貸方計</span>' + totalCredit.toLocaleString('ja-JP') + '</div>'
      + (months.length ? '<div><span class=\"k\">期間</span>' + months[0] + ' 〜 ' + months[months.length-1] + '</div>' : '');

    if (totalDebit !== totalCredit) {
      const diff = totalDebit - totalCredit;
      balanceWarn.textContent = '⚠ 貸借不一致: 差額 ' + diff.toLocaleString('ja-JP') + ' 円 (借方 ' + totalDebit.toLocaleString('ja-JP') + ' / 貸方 ' + totalCredit.toLocaleString('ja-JP') + ')';
      balanceWarn.hidden = false;
    } else {
      balanceWarn.hidden = true;
    }
  }

  function fmtTaxRateLabel(side) {
    const kbn = String(side.taxKbn || '').trim();
    if (!/^[1567]/.test(kbn)) return '';
    const reduced = String(side.taxReduced == null ? '' : side.taxReduced).trim();
    const rate = String(side.taxRate == null ? '' : side.taxRate).trim().replace(/%/g, '');
    const isReduced = reduced === '1' || reduced === '○' || reduced === 'は' || reduced.toLowerCase() === 'true';
    const is8 = rate === '8' || rate === '8.0' || rate === '08';
    if (isReduced && is8) return '軽8%';
    return '';
  }

  function formatAcct(side) {
    if (!side.name && !side.code) return '';
    const code = side.code || '';
    const subCode = (side.subCode || '').trim();
    const codeText = subCode ? code + ':' + subCode : code;
    let html = '<span class=\"code\">' + escapeHtml(codeText) + '</span>' + escapeHtml(side.name || '');
    if (side.subName) html += '<span class=\"sub\">' + escapeHtml(side.subName) + '</span>';
    if (side.deptName) html += '<span class=\"dept\">部門: ' + escapeHtml(side.deptName) + '</span>';
    return html;
  }

  filterInput.addEventListener('input', renderRows);
  monthSel.addEventListener('change', renderRows);
  $('reload').addEventListener('click', () => vscode.postMessage({ type:'reload' }));
  $('openExcel').addEventListener('click', () => vscode.postMessage({ type:'openInExcel' }));

  // 別ウィンドウ(フローティング)で開く。現在の検索/月フィルタを引き継ぐ。
  // 実体はVSCodeの独立Webviewパネルを開き、拡張側で別ウィンドウへ移動させる。
  // ----- 文字サイズ -----
  const FONT_MIN = 9, FONT_MAX = 28, FONT_DEFAULT = 13;
  const fontLvl = $('fontLvl');
  function clampFont(px) { return Math.min(FONT_MAX, Math.max(FONT_MIN, Math.round(px))); }
  let fontPx = FONT_DEFAULT;
  function applyFont(px, persist) {
    fontPx = clampFont(px);
    document.documentElement.style.setProperty('--tkc-font', fontPx + 'px');
    if (fontLvl) fontLvl.textContent = fontPx + 'px';
    if (persist !== false) {
      const st = vscode.getState() || {};
      st.fontPx = fontPx;
      vscode.setState(st);
    }
  }
  $('fontInc').addEventListener('click', () => applyFont(fontPx + 1));
  $('fontDec').addEventListener('click', () => applyFont(fontPx - 1));
  $('fontReset').addEventListener('click', () => applyFont(FONT_DEFAULT));

  $('popout').addEventListener('click', () => vscode.postMessage({
    type: 'popout',
    filter: filterInput.value,
    month: monthSel.value,
    fontPx: fontPx,
  }));

  // ポップアップで開かれた場合、呼び出し元の表示状態を復元
  const initial = data.initial || {};
  if (initial.filter) filterInput.value = initial.filter;
  if (initial.month) monthSel.value = initial.month;

  // 文字サイズ復元: 保存状態 > ポップアップ引き継ぎ > 既定
  const savedState = vscode.getState() || {};
  applyFont(savedState.fontPx || initial.fontPx || FONT_DEFAULT, false);

  renderRows();
})();
`;
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!));
}

function makeNonce(): string {
  let s = '';
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 24; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

export type { JournalEntry };
