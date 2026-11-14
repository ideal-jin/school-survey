/**
 * 취학관리 협력체계 개선(안) 전문가 의견 조사 — Google Apps Script 백엔드
 *
 * 이 스크립트는 응답을 저장할 구글 스프레드시트에 "확장 프로그램 > Apps Script"로
 * 만든 프로젝트에 붙여 넣어 사용합니다. (설치 방법은 SETUP.md 참고)
 *
 *  - doGet          : 설문 페이지(index.html)를 보여 줍니다.
 *  - submitResponse : 응답 1건을 '응답' 시트에 한 줄로 추가합니다.
 *  - doPost         : (선택) 설문 페이지를 다른 곳에 올렸을 때 쓰는 전송 주소입니다.
 *  - 메뉴 [설문 집계] : 문항별 평균·표준편차·CVR 집계표와 의견 모음표를 만듭니다.
 */

const RESPONSE_SHEET = '응답';
const SUMMARY_SHEET = '집계';
const COMMENT_SHEET = '의견 모음';
const MAX_TEXT = 5000;

/* ---------------- 설문 페이지 ---------------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('취학관리 협력체계 개선(안) 전문가 의견 조사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- 응답 저장 ---------------- */
function submitResponse(payload) {
  try {
    if (!payload || !Array.isArray(payload.headers) || !Array.isArray(payload.values) ||
        payload.headers.length !== payload.values.length || payload.headers.length > 400) {
      return { ok: false, error: '응답 형식이 올바르지 않습니다.' };
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sheet = getResponseSheet_();
      const headers = ['제출일시'].concat(payload.headers.map(String));
      const values = [new Date()].concat(payload.values.map(clean_));

      // 머리글: 비어 있으면 새로 쓰고, 새 문항이 생기면 오른쪽에 열을 추가
      let current = sheet.getLastColumn() > 0
        ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String)
        : [];
      if (current.length === 0 || current.every(h => h === '')) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        formatHeader_(sheet, headers.length);
        current = headers;
      } else {
        const missing = headers.filter(h => current.indexOf(h) === -1);
        if (missing.length) {
          sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
          current = current.concat(missing);
          formatHeader_(sheet, current.length);
        }
      }

      // 중복 제출 방지(같은 응답ID)
      const idCol = current.indexOf('응답ID') + 1;
      const id = payload.values[payload.headers.indexOf('응답ID')];
      if (idCol > 0 && sheet.getLastRow() > 1 && id) {
        const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getValues().flat();
        if (ids.indexOf(id) !== -1) return { ok: true, duplicate: true };
      }

      const row = current.map(h => {
        const i = headers.indexOf(h);
        return i === -1 ? '' : values[i];
      });
      sheet.appendRow(row);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function doPost(e) {
  let result;
  try {
    result = submitResponse(JSON.parse(e.postData.contents));
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getResponseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(RESPONSE_SHEET) || ss.insertSheet(RESPONSE_SHEET, 0);
}

function formatHeader_(sheet, n) {
  sheet.getRange(1, 1, 1, n)
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff')
    .setWrap(true).setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

// 숫자는 그대로, 글자는 길이 제한 + 수식으로 해석되지 않도록 처리
function clean_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return v;
  let s = String(v).slice(0, MAX_TEXT);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

/* ---------------- 집계 메뉴 ---------------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('설문 집계')
    .addItem('집계표·의견 모음 만들기(갱신)', 'buildReports')
    .addToUi();
}

function buildReports() {
  buildSummary_();
  buildComments_();
  SpreadsheetApp.getActiveSpreadsheet().toast('집계표와 의견 모음을 갱신했습니다.', '설문 집계', 5);
}

/**
 * '집계' 시트: [평정] 열마다 응답 수·평균·표준편차·점수 분포·4~5점 비율·CVR을
 * '응답' 시트를 참조하는 수식으로 만듭니다(응답이 늘면 자동 반영).
 * CVR(내용타당도 비율) = (4·5점 응답 수 − N/2) ÷ (N/2)
 */
function buildSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(RESPONSE_SHEET);
  if (!src || src.getLastColumn() === 0) throw new Error('아직 응답이 없습니다.');
  const headers = src.getRange(1, 1, 1, src.getLastColumn()).getValues()[0].map(String);

  const sh = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  sh.clear();

  // 입력값: Lawshe(1975) CVR 최소값 표
  sh.getRange('A1').setValue('문항별 평정 집계').setFontWeight('bold').setFontSize(13);
  sh.getRange('A2').setValue('응답자 수(N)');
  sh.getRange('B2').setFormula(`=COUNTA('${RESPONSE_SHEET}'!A2:A)`);
  sh.getRange('C2').setValue('CVR 최소값(N 기준)');
  sh.getRange('D2').setFormula('=IFERROR(VLOOKUP(B2,$P$5:$Q$20,2,TRUE),"")');
  sh.getRange('A3').setValue('※ CVR이 최소값 이상이면 내용타당도가 확보된 것으로 봅니다. 평균은 5점 만점.')
    .setFontColor('#666666');

  const head = ['문항', '응답 수', '평균', '표준편차', '1점', '2점', '3점', '4점', '5점', '4~5점 비율', 'CVR', 'CVR 기준 충족'];
  sh.getRange(4, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');

  const rows = [];
  headers.forEach((h, i) => {
    if (!/\[평정\]$/.test(h)) return;
    const col = columnLetter_(i + 1);
    const rng = `'${RESPONSE_SHEET}'!${col}2:${col}`;
    const r = 5 + rows.length;
    rows.push([
      h.replace(/\s*\[평정\]$/, ''),
      `=COUNT(${rng})`,
      `=IF(B${r}=0,"",AVERAGE(${rng}))`,
      `=IF(B${r}<2,"",STDEV(${rng}))`,
      `=COUNTIF(${rng},1)`, `=COUNTIF(${rng},2)`, `=COUNTIF(${rng},3)`,
      `=COUNTIF(${rng},4)`, `=COUNTIF(${rng},5)`,
      `=IF(B${r}=0,"",(H${r}+I${r})/B${r})`,
      `=IF(B${r}=0,"",((H${r}+I${r})-B${r}/2)/(B${r}/2))`,
      `=IF(OR(B${r}=0,$D$2=""),"",IF(K${r}>=$D$2,"충족","미달"))`
    ]);
  });
  if (rows.length) {
    sh.getRange(5, 1, rows.length, head.length).setValues(rows);
    sh.getRange(5, 3, rows.length, 2).setNumberFormat('0.00');
    sh.getRange(5, 10, rows.length, 1).setNumberFormat('0%');
    sh.getRange(5, 11, rows.length, 1).setNumberFormat('0.00');
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('미달').setBackground('#fdecea').setFontColor('#c0392b')
      .setRanges([sh.getRange(5, 12, rows.length, 1)]).build();
    const scale = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#fdecea', SpreadsheetApp.InterpolationType.NUMBER, '2.5')
      .setGradientMaxpointWithValue('#dff1ec', SpreadsheetApp.InterpolationType.NUMBER, '5')
      .setRanges([sh.getRange(5, 3, rows.length, 1)]).build();
    sh.setConditionalFormatRules([rule, scale]);
  }

  // Lawshe CVR 최소값 표(유의수준 .05)
  const lawshe = [[5,0.99],[6,0.99],[7,0.99],[8,0.75],[9,0.78],[10,0.62],[11,0.59],[12,0.56],
                  [13,0.54],[14,0.51],[15,0.49],[20,0.42],[25,0.37],[30,0.33],[35,0.31],[40,0.29]];
  sh.getRange('P4:Q4').setValues([['N', 'CVR 최소값']]).setFontWeight('bold');
  sh.getRange(5, 16, lawshe.length, 2).setValues(lawshe);
  sh.getRange('P21').setValue('출처: Lawshe(1975), 유의수준 .05').setFontColor('#666666');

  sh.setFrozenRows(4);
  sh.setColumnWidth(1, 380);
  sh.getRange(5, 1, Math.max(rows.length, 1), 1).setWrap(true);
}

/**
 * '의견 모음' 시트: [의견] 열에 적힌 글을 문항별로 모아 한 줄에 하나씩 보여 줍니다.
 * (회의 자료용. 응답이 추가되면 메뉴에서 다시 실행하세요.)
 */
function buildComments_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(RESPONSE_SHEET);
  const data = src.getDataRange().getValues();
  const headers = data[0].map(String);
  const orgCol = headers.indexOf('기본정보: 소속 기관 유형');

  const sh = ss.getSheetByName(COMMENT_SHEET) || ss.insertSheet(COMMENT_SHEET);
  sh.clear();
  const out = [['문항', '응답자 소속', '평정', '의견']];
  headers.forEach((h, c) => {
    if (!/\[의견\]$/.test(h)) return;
    const base = h.replace(/\s*\[의견\]$/, '');
    const rc = headers.indexOf(base + ' [평정]');
    for (let r = 1; r < data.length; r++) {
      const v = String(data[r][c] || '').trim();
      if (!v) continue;
      out.push([base, orgCol >= 0 ? data[r][orgCol] : '', rc >= 0 ? data[r][rc] : '', v.replace(/^'/, '')]);
    }
  });
  sh.getRange(1, 1, out.length, 4).setValues(out);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 150); sh.setColumnWidth(3, 50); sh.setColumnWidth(4, 600);
  sh.getRange(1, 1, out.length, 4).setWrap(true).setVerticalAlignment('top');
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
