# SPOOL 서명지 정렬 · 갑지 자동생성 (spool-sort-app)

오전·오후 SPOOL 불출 수령 서명지 엑셀 파일을 업로드하면, 브라우저 안에서:

1. 오전 계열 시트(오전, 오전 (2)…)와 오후 계열 시트를 각각 하나로 합산
2. 신청업체 1차 → 구분 2차 기준으로 정렬
3. 10건 단위로 새 시트에 재분할 (오전, 오전 (2), … / 오후, 오후 (2), …)
4. 업체×구분 집계 "갑지" 시트를 자동 생성 (가로 인쇄, 한 페이지 맞춤 적용)

까지 전부 처리해서 .xlsx 파일로 다운로드해 줍니다. **서버로 파일이 전송되지 않고, 기기 안에서만 처리됩니다.**

## 기존 GitHub Pages 저장소에 추가하는 방법

1. 이 폴더(`spool-sort-app`)를 저장소 루트에 `spool-sort/` 같은 이름으로 복사합니다.
   ```
   sanggyu-lee123.github.io/
     ├── label-app/        (기존 라벨 앱)
     └── spool-sort/       (이 폴더 내용)
   ```
2. 저장소에 커밋 & 푸시하면 `https://sanggyu-lee123.github.io/spool-sort/`에서 바로 열립니다.
3. 메인 페이지(루트 `index.html`)가 있다면 거기에 새 도구로 링크를 하나 추가해 주세요.

## 폴더 구조

```
index.html              메인 화면 (업로드 → 변환 → 다운로드 UI)
app.js                   UI 동작 로직 (드래그앤드롭, 진행 표시, 다운로드)
convert-engine.js        정렬·합산·분할·갑지 계산 (핵심 로직, 순수 JS)
xlsx-print-patch.js      가로방향·한페이지맞춤·인쇄영역을 xlsx 내부 XML에 직접 적용
manifest.json / sw.js    PWA 설정 (홈 화면 추가, 오프라인 캐싱)
vendor/xlsx.full.min.js  SheetJS (엑셀 읽기/쓰기)
vendor/jszip.min.js      JSZip (xlsx 내부 zip 패치용)
```

## 정렬·집계 기준 바꾸고 싶을 때

`convert-engine.js` 상단의 다음 목록만 고치면 됩니다 (다른 코드는 안 건드려도 됩니다):

```js
const VENDOR_ORDER = [...]       // 정렬 1차 기준 (신청업체)
const CATEGORY_ORDER = [...]     // 정렬 2차 기준 (구분)
const COVER_VENDOR_ROWS = [...]  // 갑지에 표시할 업체 행
const COVER_VENDOR_ALIASES = {   // 여러 표기를 한 행으로 합산 (예: 엘테크 + 엘테크(SGAS))
  '엘테크': ['엘테크', '엘테크(SGAS)'],
};
const COVER_CATEGORY_COLS = [...] // 갑지에 표시할 구분 열
const PER_SHEET = 10;             // 시트당 건수
```

## 원본 양식이 바뀌면?

원본 엑셀의 데이터 영역(행 6~15, 열 C/D/E/F/G/H/I/K)이나 날짜 셀 위치(C2)가 바뀌면
`convert-engine.js`의 `DATA_START_ROW`, `DATA_ROWS`, `DATA_COLS` 값을 맞춰 수정해야 합니다.

## 동작 확인 완료 사항

- 실제 운영 데이터(오전 3건 / 오후 18건) 기준 정렬·집계 결과가 수식 기반 버전과 100% 일치
- 오전 10시트·오후 10시트(각 100건)까지 늘려도 합계·분할·갑지 모두 정상
- 결과 파일은 LibreOffice 기준 수식 오류 0건, Excel에서 복구 경고 없이 열림
- 가로방향·한 페이지 맞춤·인쇄영역이 갑지 시트에 정확히 적용됨
