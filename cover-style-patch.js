/* xlsx 내부 styles.xml과 워크시트 XML을 직접 패치해서, SheetJS 무료판이 쓰지 못하는
 * 셀 서식(글꼴 색, 배경색, 테두리, 굵게)을 갑지 시트에 적용한다.
 *
 * SheetJS community 빌드는 cellStyles:true로 읽을 수는 있어도 쓰지는 못하므로,
 * 인쇄설정(xlsx-print-patch.js)과 동일한 방식 — zip 내부 XML 직접 수정 — 으로 구현한다.
 *
 * 사용법:
 *   const patched = await applyCoverStyles(xlsxArrayBuffer, '갑지', rowRoles, ncols);
 */

const COVER_PALETTE = {
  fonts: {
    normal: { sz: 11, name: '맑은 고딕' },
    title: { sz: 20, bold: true, color: 'FF1F3864', name: '맑은 고딕' },
    subtitle: { sz: 10.5, italic: true, color: 'FF595959', name: '맑은 고딕' },
    whiteBold: { sz: 11, bold: true, color: 'FFFFFFFF', name: '맑은 고딕' },
    navyBold: { sz: 11, bold: true, color: 'FF1F3864', name: '맑은 고딕' },
    labelBold: { sz: 11, bold: true, name: '맑은 고딕' },
    etc: { sz: 11, italic: true, color: 'FF7F7F7F', name: '맑은 고딕' },
  },
  fills: {
    navy: 'FF1F3864',
    blue: 'FF4472C4',
    band: 'FFF2F2F2',
    etc: 'FFFCE9E9',
    total: 'FFFFF2CC',
  },
};

// numFmtId 0 = 기본(General). 갑지의 "0 -> -" 표시 형식은 convert-engine.js가 이미
// SheetJS를 통해 numFmts에 등록해두므로(보통 id 60 근방), 패치 시점에 styles.xml을
// 읽어서 formatCode 문자열로 실제 id를 찾아 재사용한다 — 하드코딩하면 사용 환경마다
// id가 달라질 수 있어 어긋난다.
const DASH_FORMAT_CODE = '#,##0;-#,##0;"-"';

function buildFontXml(f) {
  const parts = [];
  if (f.bold) parts.push('<b/>');
  if (f.italic) parts.push('<i/>');
  parts.push(`<sz val="${f.sz}"/>`);
  if (f.color) parts.push(`<color rgb="${f.color}"/>`);
  else parts.push('<color theme="1"/>');
  parts.push(`<name val="${f.name}"/>`);
  return `<font>${parts.join('')}</font>`;
}

function buildFillXml(rgb) {
  return `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor rgb="${rgb}"/></patternFill></fill>`;
}

async function applyCoverStyles(xlsxArrayBuffer, sheetName, rowRoles, ncols) {
  if (!rowRoles || rowRoles.length === 0) return xlsxArrayBuffer;

  const zip = await JSZip.loadAsync(xlsxArrayBuffer);

  const workbookXml = await zip.file('xl/workbook.xml').async('string');
  const sheetNameToIndex = {};
  const sheetTagRegex = /<sheet[^>]*name="([^"]*)"[^>]*\/>/g;
  let m;
  let order = 0;
  while ((m = sheetTagRegex.exec(workbookXml)) !== null) {
    sheetNameToIndex[decodeXmlEntities(m[1])] = order;
    order += 1;
  }
  const idx = sheetNameToIndex[sheetName];
  if (idx === undefined) return xlsxArrayBuffer;
  const sheetPath = `xl/worksheets/sheet${idx + 1}.xml`;
  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) return xlsxArrayBuffer;

  // ---- 1. patch styles.xml: append fonts/fills/borders/cellXfs ----
  let stylesXml = await zip.file('xl/styles.xml').async('string');

  // find the dash-format numFmtId already registered by SheetJS (if any)
  let dashNumFmtId = 0;
  const numFmtRegex = /<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\/>/g;
  let nm;
  while ((nm = numFmtRegex.exec(stylesXml)) !== null) {
    const code = decodeXmlEntities(nm[2]);
    if (code === DASH_FORMAT_CODE) {
      dashNumFmtId = parseInt(nm[1], 10);
      break;
    }
  }

  // current counts (to know the starting index for each new entry)
  const fontCountMatch = stylesXml.match(/<fonts count="(\d+)"/);
  const fillCountMatch = stylesXml.match(/<fills count="(\d+)"/);
  const borderCountMatch = stylesXml.match(/<borders count="(\d+)"/);
  const cellXfsCountMatch = stylesXml.match(/<cellXfs count="(\d+)"/);
  const fontStart = fontCountMatch ? parseInt(fontCountMatch[1], 10) : 1;
  const fillStart = fillCountMatch ? parseInt(fillCountMatch[1], 10) : 2;
  const borderStart = borderCountMatch ? parseInt(borderCountMatch[1], 10) : 1;
  const cellXfStart = cellXfsCountMatch ? parseInt(cellXfsCountMatch[1], 10) : 1;

  // -- fonts --
  const fontKeys = Object.keys(COVER_PALETTE.fonts);
  const fontIdOf = {};
  fontKeys.forEach((k, i) => { fontIdOf[k] = fontStart + i; });
  const newFontsXml = fontKeys.map((k) => buildFontXml(COVER_PALETTE.fonts[k])).join('');
  stylesXml = stylesXml.replace(/<fonts count="(\d+)">/, (full, n) => `<fonts count="${parseInt(n, 10) + fontKeys.length}">`);
  stylesXml = stylesXml.replace('</fonts>', `${newFontsXml}</fonts>`);

  // -- fills --
  const fillKeys = Object.keys(COVER_PALETTE.fills);
  const fillIdOf = {};
  fillKeys.forEach((k, i) => { fillIdOf[k] = fillStart + i; });
  const newFillsXml = fillKeys.map((k) => buildFillXml(COVER_PALETTE.fills[k])).join('');
  stylesXml = stylesXml.replace(/<fills count="(\d+)">/, (full, n) => `<fills count="${parseInt(n, 10) + fillKeys.length}">`);
  stylesXml = stylesXml.replace('</fills>', `${newFillsXml}</fills>`);

  // -- borders (one extra: medium top+bottom for the total row) --
  const totalBorderId = borderStart;
  const newBordersXml = '<border><left/><right/><top style="medium"><color rgb="FF1F3864"/></top><bottom style="medium"><color rgb="FF1F3864"/></bottom><diagonal/></border>';
  stylesXml = stylesXml.replace(/<borders count="(\d+)">/, (full, n) => `<borders count="${parseInt(n, 10) + 1}">`);
  stylesXml = stylesXml.replace('</borders>', `${newBordersXml}</borders>`);

  // -- cellXfs: one combination per (role, column-kind) pair we need --
  // each entry: [numFmtId, fontId, fillId, borderId, horizontalAlign]
  const xfDefs = [
    ['title', 0, fontIdOf.title, 0, 0, 'center'],
    ['subtitle', 0, fontIdOf.subtitle, 0, 0, 'center'],
    ['periodBar', 0, fontIdOf.whiteBold, fillIdOf.navy, 0, 'center'],
    ['colHeader', 0, fontIdOf.whiteBold, fillIdOf.blue, 0, 'center'],
    ['vendorLabel', 0, fontIdOf.labelBold, 0, 0, 'left'],
    ['vendorLabelBand', 0, fontIdOf.labelBold, fillIdOf.band, 0, 'left'],
    ['vendorQty', dashNumFmtId, fontIdOf.navyBold, 0, 0, 'center'],
    ['vendorQtyBand', dashNumFmtId, fontIdOf.navyBold, fillIdOf.band, 0, 'center'],
    ['vendorCat', dashNumFmtId, fontIdOf.normal, 0, 0, 'center'],
    ['vendorCatBand', dashNumFmtId, fontIdOf.normal, fillIdOf.band, 0, 'center'],
    ['etcLabel', 0, fontIdOf.etc, fillIdOf.etc, 0, 'left'],
    ['etcQty', dashNumFmtId, fontIdOf.etc, fillIdOf.etc, 0, 'center'],
    ['totalLabel', 0, fontIdOf.navyBold, fillIdOf.total, totalBorderId, 'left'],
    ['totalQty', dashNumFmtId, fontIdOf.navyBold, fillIdOf.total, totalBorderId, 'center'],
  ];
  const xfIndexOf = {};
  xfDefs.forEach((d, i) => { xfIndexOf[d[0]] = cellXfStart + i; });
  const newCellXfsXml = xfDefs.map(([, numFmtId, fontId, fillId, borderId, align]) => {
    const applyNum = numFmtId ? ' applyNumberFormat="1"' : '';
    const applyBorder = borderId ? ' applyBorder="1"' : '';
    return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyFont="1" applyFill="1"${applyBorder}${applyNum}><alignment horizontal="${align}" vertical="center"/></xf>`;
  }).join('');
  stylesXml = stylesXml.replace(/<cellXfs count="(\d+)">/, (full, n) => `<cellXfs count="${parseInt(n, 10) + xfDefs.length}">`);
  stylesXml = stylesXml.replace('</cellXfs>', `${newCellXfsXml}</cellXfs>`);

  zip.file('xl/styles.xml', stylesXml);

  // ---- 2. patch the cover sheet's cells: assign s="xfIndex" per role ----
  let sheetXml = await sheetFile.async('string');

  const roleByRow = {};
  rowRoles.forEach((rr) => { roleByRow[rr.row] = rr.role; });

  function xfFor(role, col) {
    // col: 1-indexed column number; col 1 = label column, others = numeric columns
    switch (role) {
      case 'title': return xfIndexOf.title;
      case 'subtitle': return xfIndexOf.subtitle;
      case 'periodBar': return xfIndexOf.periodBar;
      case 'colHeader': return xfIndexOf.colHeader;
      case 'vendor': return col === 1 ? xfIndexOf.vendorLabel : (col === 2 ? xfIndexOf.vendorQty : xfIndexOf.vendorCat);
      case 'vendorBand': return col === 1 ? xfIndexOf.vendorLabelBand : (col === 2 ? xfIndexOf.vendorQtyBand : xfIndexOf.vendorCatBand);
      case 'etc': return col === 1 ? xfIndexOf.etcLabel : xfIndexOf.etcQty;
      case 'total': return col === 1 ? xfIndexOf.totalLabel : xfIndexOf.totalQty;
      default: return null;
    }
  }

  // rewrite <row r="N">...</row> blocks for rows we have a role for, setting/replacing
  // the s="" attribute on every <c> within that row (creating cells for blank/missing
  // ones isn't necessary since aoa_to_sheet already emits a <c> for every used column).
  sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g, (full, rowNumStr, rowAttrs, inner) => {
    const rowNum = parseInt(rowNumStr, 10);
    const role = roleByRow[rowNum];
    if (!role) return full;
    const newInner = inner.replace(/<c r="([A-Z]+)(\d+)"([^>]*)>/g, (cfull, colLetters, rRow, cAttrs) => {
      const colNum = colLettersToNumber(colLetters);
      const xfIdx = xfFor(role, colNum);
      if (xfIdx === null || xfIdx === undefined) return cfull;
      // strip any existing s="..." attribute, then add ours
      const cleanedAttrs = cAttrs.replace(/\s*s="\d+"/, '');
      return `<c r="${colLetters}${rRow}"${cleanedAttrs} s="${xfIdx}">`;
    });
    return `<row r="${rowNumStr}"${rowAttrs}>${newInner}</row>`;
  });

  zip.file(sheetPath, sheetXml);

  return zip.generateAsync({ type: 'uint8array' });
}

function colLettersToNumber(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n;
}

function decodeXmlEntities(s) {
  return String(s)
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

if (typeof module !== 'undefined') {
  module.exports = { applyCoverStyles };
}
