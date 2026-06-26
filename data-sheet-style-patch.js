/* 오전/오후 등 데이터 시트(서명지)에 원본 양식의 모양 — 큰 제목, 회색 헤더 바,
 * 데이터 영역 테두리 — 을 적용한다. 갑지(cover-style-patch.js)와 같은 방식으로
 * zip 내부 styles.xml / worksheet XML을 직접 패치한다.
 *
 * 시트 구조는 회사 템플릿 고정 구조를 가정한다 (DATA_START_ROW=6, DATA_ROWS=10):
 *   row 1: 제목 (병합 B1:K1)
 *   row 2: 오전/오후 + 날짜
 *   row 4: 불출/수령 라벨 바 (병합 B4:I4, K4)
 *   row 5: 열 헤더 (B5:K5)
 *   row 6~15: 데이터 (테두리 + 가운데 정렬)
 *   row 16: 푸터 (물류 담당자)
 *
 * 사용법:
 *   const patched = await applyDataSheetStyles(xlsxArrayBuffer, ['오전','오후','오후 (2)']);
 */

async function applyDataSheetStyles(xlsxArrayBuffer, sheetNames) {
  if (!sheetNames || sheetNames.length === 0) return xlsxArrayBuffer;

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

  // ---- 1. patch styles.xml once: append the handful of styles every data sheet needs ----
  let stylesXml = await zip.file('xl/styles.xml').async('string');

  const fontCountMatch = stylesXml.match(/<fonts count="(\d+)"/);
  const fillCountMatch = stylesXml.match(/<fills count="(\d+)"/);
  const borderCountMatch = stylesXml.match(/<borders count="(\d+)"/);
  const cellXfsCountMatch = stylesXml.match(/<cellXfs count="(\d+)"/);
  const fontStart = fontCountMatch ? parseInt(fontCountMatch[1], 10) : 1;
  const fillStart = fillCountMatch ? parseInt(fillCountMatch[1], 10) : 2;
  const borderStart = borderCountMatch ? parseInt(borderCountMatch[1], 10) : 1;
  const cellXfStart = cellXfsCountMatch ? parseInt(cellXfsCountMatch[1], 10) : 1;

  // -- fonts --
  const fonts = {
    title: '<font><b/><sz val="24"/><color theme="1"/><name val="맑은 고딕"/></font>',
    sub: '<font><b/><sz val="12"/><color theme="1"/><name val="맑은 고딕"/></font>',
    headerBold: '<font><b/><sz val="11"/><color theme="1"/><name val="맑은 고딕"/></font>',
    bodyNormal: '<font><sz val="11"/><color theme="1"/><name val="맑은 고딕"/></font>',
    footerBold: '<font><b/><sz val="11"/><color theme="1"/><name val="맑은 고딕"/></font>',
  };
  const fontKeys = Object.keys(fonts);
  const fontIdOf = {};
  fontKeys.forEach((k, i) => { fontIdOf[k] = fontStart + i; });
  stylesXml = stylesXml.replace(/<fonts count="(\d+)">/, (full, n) => `<fonts count="${parseInt(n, 10) + fontKeys.length}">`);
  stylesXml = stylesXml.replace('</fonts>', `${fontKeys.map((k) => fonts[k]).join('')}</fonts>`);

  // -- fills: light gray header bar (Excel "White, Background 1, Darker 25%" ≈ #BFBFBF) --
  const grayFillId = fillStart;
  const newFillsXml = '<fill><patternFill patternType="solid"><fgColor rgb="FFBFBFBF"/><bgColor rgb="FFBFBFBF"/></patternFill></fill>';
  stylesXml = stylesXml.replace(/<fills count="(\d+)">/, (full, n) => `<fills count="${parseInt(n, 10) + 1}">`);
  stylesXml = stylesXml.replace('</fills>', `${newFillsXml}</fills>`);

  // -- borders: thin grid for data cells, medium divider under the 불출/수령 label bar --
  const thinBorderId = borderStart;
  const newThinBorderXml = '<border><left style="thin"><color rgb="FF808080"/></left><right style="thin"><color rgb="FF808080"/></right><top style="thin"><color rgb="FF808080"/></top><bottom style="thin"><color rgb="FF808080"/></bottom><diagonal/></border>';
  const dividerBorderId = borderStart + 1;
  const newDividerBorderXml = '<border><left/><right/><top/><bottom style="medium"><color rgb="FF595959"/></bottom><diagonal/></border>';
  stylesXml = stylesXml.replace(/<borders count="(\d+)">/, (full, n) => `<borders count="${parseInt(n, 10) + 2}">`);
  stylesXml = stylesXml.replace('</borders>', `${newThinBorderXml}${newDividerBorderXml}</borders>`);

  // -- cellXfs combos --
  const xfDefs = [
    // [key, fontId, fillId, borderId, align, wrap]
    ['title', fontIdOf.title, 0, 0, 'center', false],
    ['sub', fontIdOf.sub, 0, 0, 'left', false],
    ['headerBarTop', fontIdOf.headerBold, grayFillId, dividerBorderId, 'center', true], // row 4 — adds the divider line under it
    ['headerBarBottom', fontIdOf.headerBold, grayFillId, 0, 'center', true], // row 5 — no extra border (divider already drawn above by row 4)
    ['dataCell', fontIdOf.bodyNormal, 0, thinBorderId, 'center', false],
    ['footer', fontIdOf.footerBold, 0, 0, 'left', false],
  ];
  const xfIndexOf = {};
  xfDefs.forEach((d, i) => { xfIndexOf[d[0]] = cellXfStart + i; });
  const newCellXfsXml = xfDefs.map(([, fontId, fillId, borderId, align, wrap]) => {
    const applyBorder = borderId ? ' applyBorder="1"' : '';
    const wrapAttr = wrap ? ' wrapText="1"' : '';
    return `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0" applyFont="1" applyFill="1"${applyBorder}><alignment horizontal="${align}" vertical="center"${wrapAttr}/></xf>`;
  }).join('');
  stylesXml = stylesXml.replace(/<cellXfs count="(\d+)">/, (full, n) => `<cellXfs count="${parseInt(n, 10) + xfDefs.length}">`);
  stylesXml = stylesXml.replace('</cellXfs>', `${newCellXfsXml}</cellXfs>`);

  zip.file('xl/styles.xml', stylesXml);

  // row role map shared by every data sheet (fixed company template layout)
  const ROLE_BY_ROW = { 1: 'title', 2: 'sub', 4: 'headerBarTop', 5: 'headerBarBottom', 16: 'footer' };
  for (let r = 6; r <= 15; r++) ROLE_BY_ROW[r] = 'dataCell';

  // row 2 also contains the date cell (보통 column C) which already carries its own
  // date number format (z) baked into its existing style index — restyling it with our
  // generic "sub" xf (numFmtId=0) would clobber that and show the raw serial number
  // instead of a date. Skip any column in row 2 other than the label column (B).
  const ROW2_ONLY_COLS = new Set(['B']);

  // ---- 2. patch each requested sheet's worksheet XML ----
  for (const sheetName of sheetNames) {
    const idx = sheetNameToIndex[sheetName];
    if (idx === undefined) continue;
    const sheetPath = `xl/worksheets/sheet${idx + 1}.xml`;
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    let sheetXml = await sheetFile.async('string');

    sheetXml = sheetXml.replace(/<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g, (full, rowNumStr, rowAttrs, inner) => {
      const rowNum = parseInt(rowNumStr, 10);
      const role = ROLE_BY_ROW[rowNum];
      if (!role) return full;
      const xfIdx = xfIndexOf[role];
      const newInner = inner.replace(/<c r="([A-Z]+)(\d+)"([^>]*)>/g, (cfull, colLetters, rRow, cAttrs) => {
        if (rowNum === 2 && !ROW2_ONLY_COLS.has(colLetters)) return cfull; // keep date cell's own format
        const cleanedAttrs = cAttrs.replace(/\s*s="\d+"/, '');
        return `<c r="${colLetters}${rRow}"${cleanedAttrs} s="${xfIdx}">`;
      });
      return `<row r="${rowNumStr}"${rowAttrs}>${newInner}</row>`;
    });

    zip.file(sheetPath, sheetXml);
  }

  return zip.generateAsync({ type: 'uint8array' });
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
  module.exports = { applyDataSheetStyles };
}
