/* xlsx 파일(zip) 내부의 worksheet XML / workbook XML을 직접 패치해서
 * SheetJS만으로는 신뢰할 수 없는 인쇄설정(가로방향, 한 페이지 맞춤, 인쇄영역)을
 * 정확하게 적용한다.
 *
 * 사용법:
 *   const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
 *   const patched = await applyPrintSettings(buf, {
 *     '갑지': { orientation: 'landscape', fitToPage: true, printArea: 'A1:J29' },
 *   });
 *   // patched: Uint8Array, ready to download
 */

async function applyPrintSettings(xlsxArrayBuffer, sheetSettingsByName) {
  const zip = await JSZip.loadAsync(xlsxArrayBuffer);

  // 1. read workbook.xml to map sheet name -> sheet index (r:id / order)
  const workbookXmlPath = 'xl/workbook.xml';
  let workbookXml = await zip.file(workbookXmlPath).async('string');

  const sheetNameToIndex = {};
  const sheetTagRegex = /<sheet[^>]*name="([^"]*)"[^>]*\/>/g;
  let m;
  let order = 0;
  while ((m = sheetTagRegex.exec(workbookXml)) !== null) {
    const decodedName = decodeXmlEntities(m[1]);
    sheetNameToIndex[decodedName] = order;
    order += 1;
  }

  // 2. for each requested sheet, patch xl/worksheets/sheetN.xml
  const printAreaDefs = [];

  for (const sheetName of Object.keys(sheetSettingsByName)) {
    const settings = sheetSettingsByName[sheetName];
    const idx = sheetNameToIndex[sheetName];
    if (idx === undefined) continue; // sheet not found, skip safely

    const sheetPath = `xl/worksheets/sheet${idx + 1}.xml`;
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) continue;
    let xml = await sheetFile.async('string');

    if (settings.fitToPage) {
      if (!/<sheetPr/.test(xml)) {
        xml = xml.replace(/(<worksheet[^>]*>)/, `$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`);
      } else {
        xml = xml.replace(/<sheetPr([^>]*)\/>/, `<sheetPr$1><pageSetUpPr fitToPage="1"/></sheetPr>`);
      }
    }

    // remove any existing pageMargins/pageSetup so we don't duplicate
    xml = xml.replace(/<pageMargins[^>]*\/>/, '');
    xml = xml.replace(/<pageSetup[^>]*\/>/, '');

    const margins = settings.margins || { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 };
    const marginsXml = `<pageMargins left="${margins.left}" right="${margins.right}" top="${margins.top}" bottom="${margins.bottom}" header="${margins.header}" footer="${margins.footer}"/>`;

    const orientation = settings.orientation || 'portrait';
    const fitW = settings.fitToPage ? ' fitToWidth="1" fitToHeight="1"' : '';
    const pageSetupXml = `<pageSetup orientation="${orientation}"${fitW}/>`;
    const insertion = `${marginsXml}${pageSetupXml}`;

    // OOXML requires a strict child element order inside <worksheet>. pageMargins/pageSetup
    // must come BEFORE elements like headerFooter, rowBreaks, colBreaks, customProperties,
    // cellWatches, ignoredErrors, smartTags, drawing, tableParts, extLst — inserting blindly
    // right before </worksheet> can land AFTER one of those (most commonly <ignoredErrors>,
    // which SheetJS adds whenever a numeric-looking string is stored as text) and produce a
    // file that lenient readers (LibreOffice, openpyxl) tolerate but real Excel rejects with
    // a "needs repair" prompt. Insert before the first such trailing element if present.
    const trailingTagPattern = /<(ignoredErrors|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[ >]/;
    const trailingMatch = xml.match(trailingTagPattern);
    if (trailingMatch) {
      xml = xml.slice(0, trailingMatch.index) + insertion + xml.slice(trailingMatch.index);
    } else if (/<\/worksheet>/.test(xml)) {
      xml = xml.replace('</worksheet>', `${insertion}</worksheet>`);
    }

    zip.file(sheetPath, xml);

    if (settings.printArea) {
      printAreaDefs.push({ idx, sheetName, range: settings.printArea });
    }
  }

  // 3. add/merge definedNames for print areas in workbook.xml
  if (printAreaDefs.length > 0) {
    const newDefs = printAreaDefs
      .map(({ idx, sheetName, range }) => {
        const safeRange = range.replace(/([A-Za-z가-힣0-9_]+)!/, ''); // strip any sheet prefix if present
        return `<definedName name="_xlnm.Print_Area" localSheetId="${idx}">'${escapeXmlAttr(sheetName)}'!${range}</definedName>`;
      })
      .join('');

    if (/<definedNames>/.test(workbookXml)) {
      workbookXml = workbookXml.replace('</definedNames>', `${newDefs}</definedNames>`);
    } else if (/<\/sheets>/.test(workbookXml)) {
      workbookXml = workbookXml.replace('</sheets>', `</sheets><definedNames>${newDefs}</definedNames>`);
    }
    zip.file(workbookXmlPath, workbookXml);
  }

  return zip.generateAsync({ type: 'uint8array' });
}

function escapeXmlAttr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
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
  module.exports = { applyPrintSettings };
}
