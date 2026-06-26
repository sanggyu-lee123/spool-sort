/* SPOOL 서명지 정렬·집계 변환 엔진
 * - 오전/오후 그룹별로 시트를 모두 합산
 * - 신청업체 1차, 구분 2차 기준 정렬
 * - 10건 단위로 시트 재분할
 * - 갑지(집계) 시트 생성 (값만 계산해서 기록, 수식 없음)
 */

const PER_SHEET = 10;
const DATA_START_ROW = 6; // 1-indexed, matches template row 6
const DATA_ROWS = 10; // rows 6..15
const DATA_COLS = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'K'];

const VENDOR_ORDER = [
  '최형재팀', '박인기팀', '유성완팀', '연호', '하나기술',
  '엘테크', '엘테크(SGAS)', '연호(사외)', '하나기술(사외)',
];
const CATEGORY_ORDER = ['GAS', 'PCW', '구조관', '코팅배관', 'SPIRAL', 'PVC', '서포트'];

// 갑지 행 라벨과, 그 라벨로 합산할 실제 업체명(별칭) 목록
const COVER_VENDOR_ROWS = ['최형재팀', '박인기팀', '유성완팀', '연호', '하나기술', '엘테크', '연호(사외)', '하나기술(사외)'];
const COVER_VENDOR_ALIASES = {
  '엘테크': ['엘테크', '엘테크(SGAS)'],
};
const COVER_CATEGORY_COLS = ['GAS', 'PCW', '구조관', '코팅배관', 'SPIRAL', 'PVC', '서포트'];

function norm(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/\n/g, '').replace(/\s/g, '').trim();
}

function vendorKey(v) {
  const nv = norm(v);
  const idx = VENDOR_ORDER.findIndex((name) => norm(name) === nv);
  return idx === -1 ? VENDOR_ORDER.length : idx;
}

function categoryKey(c) {
  const nc = norm(c);
  const idx = CATEGORY_ORDER.findIndex((name) => norm(name) === nc);
  return idx === -1 ? CATEGORY_ORDER.length : idx;
}

function periodOf(sheetName) {
  return sheetName.replace(/\s*\(\d+\)\s*$/, '').trim();
}

/**
 * @param {object} XLSX - the SheetJS module
 * @param {ArrayBuffer} arrayBuffer - uploaded file contents
 * @returns {{workbook: object, meta: object}} processed workbook (SheetJS book) + summary info
 */
function normalizeWorkbookNewlines(wb) {
  wb.SheetNames.forEach((sn) => {
    const ws = wb.Sheets[sn];
    Object.keys(ws).forEach((addr) => {
      if (addr[0] === '!') return;
      const cell = ws[addr];
      if (cell && cell.t === 's' && typeof cell.v === 'string' && cell.v.indexOf('\r') !== -1) {
        cell.v = cell.v.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        if (typeof cell.w === 'string') {
          cell.w = cell.w.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        }
      }
    });
  });
}

function processWorkbook(XLSX, arrayBuffer) {
  const bytes = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
  const wb = XLSX.read(bytes, { type: 'array', cellNF: true, cellStyles: true, cellText: false });
  normalizeWorkbookNewlines(wb);

  // ---- 1. group sheets by period (오전 / 오후 / ...) ----
  const groups = {}; // period -> [sheetNames]
  const periodOrder = [];
  wb.SheetNames.forEach((sn) => {
    const p = periodOf(sn);
    if (!groups[p]) {
      groups[p] = [];
      periodOrder.push(p);
    }
    groups[p].push(sn);
  });

  // ---- 2. collect template (header/style reference) + all data rows per period ----
  const periodRows = {}; // period -> array of row objects {C,D,E,F,G,H,I,K}
  const templateInfo = {}; // period -> {sheetName, headerDate}

  periodOrder.forEach((p) => {
    const rows = [];
    groups[p].forEach((sn) => {
      const ws = wb.Sheets[sn];
      for (let r = DATA_START_ROW; r < DATA_START_ROW + DATA_ROWS; r++) {
        const rowVals = {};
        let empty = true;
        DATA_COLS.forEach((col) => {
          const addr = `${col}${r}`;
          const cell = ws[addr];
          const val = cell ? cell.v : undefined;
          rowVals[col] = val === undefined ? null : val;
          if (val !== undefined && val !== null && val !== '') empty = false;
        });
        if (!empty) rows.push(rowVals);
      }
    });
    rows.sort((a, b) => {
      const va = vendorKey(a.F);
      const vb = vendorKey(b.F);
      if (va !== vb) return va - vb;
      return categoryKey(a.C) - categoryKey(b.C);
    });
    periodRows[p] = rows;

    const firstSheet = groups[p][0];
    const ws0 = wb.Sheets[firstSheet];
    const dateCell = ws0['C2'];
    templateInfo[p] = {
      sheetName: firstSheet,
      headerDate: dateCell ? dateCell.v : null,
      dateCellRaw: dateCell || null,
    };
  });

  // ---- 3. build new workbook reusing the first sheet of each period as a style template ----
  const newWb = XLSX.utils.book_new();
  newWb.Props = wb.Props;

  const sheetMeta = []; // for later use (cover sheet references etc.)

  periodOrder.forEach((p) => {
    const rows = periodRows[p];
    const sheetCount = Math.max(1, Math.ceil(rows.length / PER_SHEET));
    const templateSheetName = templateInfo[p].sheetName;
    const templateWs = wb.Sheets[templateSheetName];

    for (let i = 0; i < sheetCount; i++) {
      const newName = sheetCount === 1 || i === 0 ? p : `${p} (${i + 1})`;
      const newWs = cloneSheetTemplate(XLSX, templateWs);
      const chunk = rows.slice(i * PER_SHEET, (i + 1) * PER_SHEET);
      writeDataRows(newWs, chunk);
      XLSX.utils.book_append_sheet(newWb, newWs, newName);
      sheetMeta.push({ period: p, name: newName });
    }
  });

  // ---- 4. build cover ("갑지") sheet with computed (static) values ----
  const coverWs = buildCoverSheet(XLSX, periodOrder, periodRows, templateInfo);
  XLSX.utils.book_append_sheet(newWb, coverWs, '갑지');

  // move 갑지 to front
  const idx = newWb.SheetNames.indexOf('갑지');
  if (idx > 0) {
    newWb.SheetNames.splice(idx, 1);
    newWb.SheetNames.unshift('갑지');
  }

  return {
    workbook: newWb,
    meta: { periodOrder, sheetMeta, periodRows },
  };
}

/** Trim a SheetJS '!cols' array down to the columns actually in use.
 *  SheetJS expands a single "remaining columns" range (e.g. col 15~16384)
 *  read from the source file into one array entry PER COLUMN, which then
 *  gets written back out as one <col> XML tag per entry — ballooning the
 *  file to hundreds of KB per sheet for no visual benefit. We only need
 *  the handful of columns the template actually uses (well beyond
 *  DATA_COLS/K to be safe), so anything past that is dropped. */
const MAX_TEMPLATE_COLS = 20;
function trimCols(cols) {
  if (!cols) return cols;
  return cols.slice(0, MAX_TEMPLATE_COLS).map((c) => Object.assign({}, c));
}

/** Clone a worksheet's structure (merges, col widths, row heights, header cells, styles)
 *  but leave data rows (6..15) empty, ready to be filled. */
function cloneSheetTemplate(XLSX, templateWs) {
  const newWs = {};
  // copy every cell (deep-ish copy), we'll overwrite data rows after
  Object.keys(templateWs).forEach((addr) => {
    if (addr[0] === '!') return;
    newWs[addr] = Object.assign({}, templateWs[addr]);
  });
  if (templateWs['!ref']) newWs['!ref'] = templateWs['!ref'];
  if (templateWs['!merges']) newWs['!merges'] = templateWs['!merges'].map((m) => Object.assign({}, m));
  if (templateWs['!cols']) newWs['!cols'] = trimCols(templateWs['!cols']);
  if (templateWs['!rows']) newWs['!rows'] = templateWs['!rows'].map((r) => (r ? Object.assign({}, r) : r));

  // clear data rows 6..15 across the relevant columns (including No. column B)
  for (let r = DATA_START_ROW; r < DATA_START_ROW + DATA_ROWS; r++) {
    [...DATA_COLS, 'B'].forEach((col) => {
      const addr = `${col}${r}`;
      if (newWs[addr]) {
        delete newWs[addr].v;
        delete newWs[addr].w;
        delete newWs[addr].f;
        // keep .s (style) and .t will be reset on write
      }
    });
  }
  return newWs;
}

function writeDataRows(ws, rows) {
  for (let offset = 0; offset < DATA_ROWS; offset++) {
    const r = DATA_START_ROW + offset;
    const rowVals = rows[offset]; // undefined if beyond actual data -> blank row
    DATA_COLS.forEach((col) => {
      const addr = `${col}${r}`;
      const existing = ws[addr] || {};
      const val = rowVals ? rowVals[col] : null;
      if (val === null || val === undefined || val === '') {
        ws[addr] = Object.assign({}, existing);
        delete ws[addr].v;
        delete ws[addr].w;
        delete ws[addr].f;
        return;
      }
      const cell = Object.assign({}, existing);
      if (typeof val === 'number') {
        cell.t = 'n';
        cell.v = val;
      } else {
        cell.t = 's';
        cell.v = String(val);
      }
      delete cell.f;
      ws[addr] = cell;
    });
    // No. column (B): sequential number only for actual data rows; blank otherwise
    const bAddr = `B${r}`;
    const bExisting = ws[bAddr] || {};
    if (rowVals) {
      ws[bAddr] = Object.assign({}, bExisting, { t: 'n', v: offset + 1 });
      delete ws[bAddr].f;
    } else {
      ws[bAddr] = Object.assign({}, bExisting);
      delete ws[bAddr].v;
      delete ws[bAddr].w;
      delete ws[bAddr].f;
    }
  }
}

function buildCoverSheet(XLSX, periodOrder, periodRows, templateInfo) {
  const NCOLS = 1 + 1 + COVER_CATEGORY_COLS.length + 1; // label + 총수량 + categories + 기타
  const aoa = [];

  aoa.push(['SPOOL 불출 수령 포장개수 집계 갑지']);
  aoa.push(['안성공장 (통합 외부제작 Shop)  |  업체별·구분별 포장 수량 집계']);
  aoa.push([]);

  const merges = [];
  let r = 0; // 0-indexed row for aoa
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: NCOLS - 1 } });
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: NCOLS - 1 } });
  r = 3;

  periodOrder.forEach((p) => {
    const rows = periodRows[p];
    const dateVal = templateInfo[p].headerDate;
    let dateStr = '';
    if (dateVal instanceof Date) {
      dateStr = `${dateVal.getFullYear()}년 ${dateVal.getMonth() + 1}월 ${dateVal.getDate()}일`;
    } else if (typeof dateVal === 'number') {
      // Excel serial date
      const d = XLSX.SSF.parse_date_code(dateVal);
      if (d) dateStr = `${d.y}년 ${d.m}월 ${d.d}일`;
    } else if (dateVal) {
      dateStr = String(dateVal);
    }

    // title row: date | period label
    const titleRow = new Array(NCOLS).fill('');
    titleRow[0] = dateStr;
    titleRow[2] = `${p}  SPOOL 불출 현황`;
    aoa.push(titleRow);
    merges.push({ s: { r, c: 0 }, e: { r, c: 1 } });
    merges.push({ s: { r, c: 2 }, e: { r, c: NCOLS - 1 } });
    r += 1;

    // header row
    const headerRow = ['신청업체', '총수량', ...COVER_CATEGORY_COLS, '기타'];
    aoa.push(headerRow);
    r += 1;

    // compute totals
    const vendorTotals = {}; // label -> {total, byCategory:{}, }
    COVER_VENDOR_ROWS.forEach((v) => {
      vendorTotals[v] = { total: 0, byCategory: {} };
      COVER_CATEGORY_COLS.forEach((c) => { vendorTotals[v].byCategory[c] = 0; });
    });
    let etcTotal = 0;
    const etcByCategory = {};
    COVER_CATEGORY_COLS.forEach((c) => { etcByCategory[c] = 0; });

    // build reverse alias lookup: normalized alias -> row label
    const aliasLookup = {};
    COVER_VENDOR_ROWS.forEach((v) => {
      const aliases = COVER_VENDOR_ALIASES[v] || [v];
      aliases.forEach((a) => { aliasLookup[norm(a)] = v; });
    });

    rows.forEach((row) => {
      const qty = typeof row.K === 'number' ? row.K : 0;
      const vendorNorm = norm(row.F);
      const catNorm = norm(row.C);
      const matchedVendor = aliasLookup[vendorNorm];
      const matchedCat = COVER_CATEGORY_COLS.find((c) => norm(c) === catNorm);

      if (matchedVendor) {
        vendorTotals[matchedVendor].total += qty;
        if (matchedCat) {
          vendorTotals[matchedVendor].byCategory[matchedCat] += qty;
        }
      } else {
        etcTotal += qty;
        if (matchedCat) etcByCategory[matchedCat] += qty;
      }
    });

    let grandTotal = 0;
    const catGrandTotal = {};
    COVER_CATEGORY_COLS.forEach((c) => { catGrandTotal[c] = 0; });

    COVER_VENDOR_ROWS.forEach((v) => {
      const t = vendorTotals[v];
      const dataRow = [v, t.total];
      COVER_CATEGORY_COLS.forEach((c) => {
        dataRow.push(t.byCategory[c]);
        catGrandTotal[c] += t.byCategory[c];
      });
      const namedSum = COVER_CATEGORY_COLS.reduce((acc, c) => acc + t.byCategory[c], 0);
      dataRow.push(t.total - namedSum); // 기타 column for this vendor row
      aoa.push(dataRow);
      grandTotal += t.total;
    });
    r += COVER_VENDOR_ROWS.length;

    // 기타 vendor row
    const etcNamedSum = COVER_CATEGORY_COLS.reduce((acc, c) => acc + etcByCategory[c], 0);
    const etcRow = ['기타', etcTotal, ...COVER_CATEGORY_COLS.map((c) => etcByCategory[c]), etcTotal - etcNamedSum];
    aoa.push(etcRow);
    r += 1;
    grandTotal += etcTotal;
    COVER_CATEGORY_COLS.forEach((c) => { catGrandTotal[c] += etcByCategory[c]; });

    // total row
    const catGrandSum = COVER_CATEGORY_COLS.reduce((acc, c) => acc + catGrandTotal[c], 0);
    const totalRow = ['합계', grandTotal, ...COVER_CATEGORY_COLS.map((c) => catGrandTotal[c]), grandTotal - catGrandSum];
    aoa.push(totalRow);
    r += 1;

    aoa.push([]); // spacer
    r += 1;
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  ws['!cols'] = [
    { wch: 16 }, { wch: 11 },
    ...COVER_CATEGORY_COLS.map(() => ({ wch: 10.5 })),
    { wch: 10.5 },
  ];

  const lastRow = aoa.length; // 1-indexed last row used
  const lastColLetter = XLSX.utils.encode_col(NCOLS - 1);
  ws['!printArea'] = `A1:${lastColLetter}${lastRow}`; // custom marker consumed by xlsx-print-patch.js

  return ws;
}

if (typeof module !== 'undefined') {
  module.exports = { processWorkbook, periodOf, norm, vendorKey, categoryKey };
}
