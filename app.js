(function () {
  'use strict';

  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const fileChipHolder = document.getElementById('fileChipHolder');
  const fileNameEl = document.getElementById('fileName');
  const removeFileBtn = document.getElementById('removeFile');
  const convertBtn = document.getElementById('convertBtn');
  const statusBox = document.getElementById('statusBox');
  const statusTxt = document.getElementById('statusTxt');
  const errorBox = document.getElementById('errorBox');
  const uploadCard = document.getElementById('uploadCard');
  const resultCard = document.getElementById('resultCard');
  const summaryGrid = document.getElementById('summaryGrid');
  const sheetList = document.getElementById('sheetList');
  const downloadBtn = document.getElementById('downloadBtn');
  const resetBtn = document.getElementById('resetBtn');
  const rulesToggleBtn = document.getElementById('rulesToggleBtn');
  const rulesPanel = document.getElementById('rulesPanel');
  const stepDots = document.querySelectorAll('.steps .dot');

  let selectedFile = null;
  let outputBlob = null;
  let outputFileName = '';

  function setStep(n) {
    stepDots.forEach((d) => {
      const s = Number(d.dataset.step);
      d.classList.remove('active', 'done');
      if (s < n) d.classList.add('done');
      if (s === n) d.classList.add('active');
    });
  }
  setStep(1);

  function showError(msg) {
    errorBox.style.display = 'block';
    errorBox.textContent = msg;
  }
  function clearError() {
    errorBox.style.display = 'none';
    errorBox.textContent = '';
  }

  function pickFile(file) {
    if (!file) return;
    const lower = file.name.toLowerCase();
    if (!lower.endsWith('.xlsx')) {
      showError('.xlsx 형식의 엑셀 파일만 지원합니다.');
      return;
    }
    clearError();
    selectedFile = file;
    fileNameEl.textContent = file.name;
    fileChipHolder.style.display = 'block';
    dropzone.style.display = 'none';
    convertBtn.disabled = false;
  }

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) pickFile(e.target.files[0]);
  });

  ['dragenter', 'dragover'].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.add('drag');
    });
  });
  ['dragleave', 'drop'].forEach((evt) => {
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.remove('drag');
    });
  });
  dropzone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files && e.dataTransfer.files[0]) pickFile(e.dataTransfer.files[0]);
  });

  removeFileBtn.addEventListener('click', () => {
    selectedFile = null;
    fileInput.value = '';
    fileChipHolder.style.display = 'none';
    dropzone.style.display = 'block';
    convertBtn.disabled = true;
    clearError();
  });

  rulesToggleBtn.addEventListener('click', () => {
    rulesPanel.classList.toggle('open');
    rulesToggleBtn.textContent = rulesPanel.classList.contains('open')
      ? '정렬·집계 기준 닫기'
      : '정렬·집계 기준 보기';
  });

  function fmtDateForFilename(d) {
    const yy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yy}${mm}${dd}`;
  }

  convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    clearError();
    convertBtn.disabled = true;
    convertBtn.classList.add('loading');
    statusBox.style.display = 'flex';
    statusTxt.textContent = '파일을 읽는 중…';
    setStep(2);

    try {
      const arrayBuffer = await selectedFile.arrayBuffer();

      statusTxt.textContent = '정렬하고 집계하는 중…';
      await nextFrame();
      const { workbook, meta } = processWorkbook(XLSX, arrayBuffer);

      statusTxt.textContent = '엑셀 파일로 저장하는 중…';
      await nextFrame();
      const rawOut = XLSX.write(workbook, { type: 'array', bookType: 'xlsx', cellStyles: true, bookSST: true });

      statusTxt.textContent = '인쇄 설정을 적용하는 중…';
      await nextFrame();
      const coverWs = workbook.Sheets['갑지'];
      const printArea = coverWs['!ref'] || 'A1:J28';
      const printSettings = {
        '갑지': {
          orientation: 'landscape',
          fitToPage: true,
          printArea,
        },
      };
      // 오전/오후 등 데이터 시트도 가로 + 한 페이지 맞춤으로 동일하게 인쇄 설정
      meta.sheetMeta.forEach((s) => {
        printSettings[s.name] = {
          orientation: 'landscape',
          fitToPage: true,
          printArea: s.printArea || undefined,
        };
      });
      const patched = await applyPrintSettings(rawOut, printSettings);

      statusTxt.textContent = '갑지에 서식을 입히는 중…';
      await nextFrame();
      const styled = await applyCoverStyles(patched, '갑지', meta.coverRowRoles, meta.coverNCols);

      outputBlob = new Blob([styled], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      const base = selectedFile.name.replace(/\.xlsx$/i, '');
      outputFileName = `${base}_정렬_갑지.xlsx`;

      renderResult(meta);
      setStep(3);
      uploadCard.style.display = 'none';
      resultCard.style.display = 'block';
    } catch (err) {
      console.error(err);
      showError('변환 중 문제가 발생했습니다: ' + (err && err.message ? err.message : String(err)) +
        '\n파일 형식이 기존 서명지 양식과 같은지 확인해 주세요.');
      setStep(1);
    } finally {
      convertBtn.disabled = false;
      convertBtn.classList.remove('loading');
      statusBox.style.display = 'none';
    }
  });

  function nextFrame() {
    return new Promise((resolve) => setTimeout(resolve, 30));
  }

  function renderResult(meta) {
    summaryGrid.innerHTML = '';
    sheetList.innerHTML = '';

    meta.periodOrder.forEach((p) => {
      const rows = meta.periodRows[p];
      const sheetsForPeriod = meta.sheetMeta.filter((s) => s.period === p);
      const total = rows.reduce((acc, r) => acc + (typeof r.K === 'number' ? r.K : 0), 0);

      const box = document.createElement('div');
      box.className = 'summary-box';
      box.innerHTML = `
        <div class="lbl">${escapeHtml(p)}</div>
        <div class="val">${rows.length}건</div>
        <div class="sub">포장 ${total}개 · 시트 ${sheetsForPeriod.length}개</div>
      `;
      summaryGrid.appendChild(box);
    });

    meta.sheetMeta.forEach((s) => {
      const row = document.createElement('div');
      row.className = 'sheet-row';
      row.innerHTML = `
        <span><span class="tag">${escapeHtml(s.period)}</span>${escapeHtml(s.name)}</span>
        <span class="cnt"></span>
      `;
      sheetList.appendChild(row);
    });
    const coverRow = document.createElement('div');
    coverRow.className = 'sheet-row';
    coverRow.innerHTML = `<span><span class="tag" style="background:#fcefe3;color:#b25a12;">집계</span>갑지</span><span class="cnt"></span>`;
    sheetList.insertBefore(coverRow, sheetList.firstChild);
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
  }

  downloadBtn.addEventListener('click', () => {
    if (!outputBlob) return;
    const url = URL.createObjectURL(outputBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = outputFileName || 'spool_sorted.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  });

  resetBtn.addEventListener('click', () => {
    selectedFile = null;
    outputBlob = null;
    fileInput.value = '';
    fileChipHolder.style.display = 'none';
    dropzone.style.display = 'block';
    convertBtn.disabled = true;
    uploadCard.style.display = 'block';
    resultCard.style.display = 'none';
    clearError();
    setStep(1);
  });
})();
