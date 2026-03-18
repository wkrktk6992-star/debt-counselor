// ============================================
// 채무조정 상담도구 - Google Apps Script v5
// 선택제도 앞으로 배치 + 깔끔한 시트 포맷
// ============================================

const SHEET_NAME = '상담이력';

// 컬럼 순서 (34개) — 선택제도·결과가 앞쪽
const HEADERS = [
  '상담ID',                          // A
  '이름', '연락처',                    // B-C
  '선택제도', '예상감면율', '월변제액', '변제기간',  // D-G  ★ 핵심결과
  '추천1위', '추천2위', '추천3위', '추천4위',     // H-K
  '출생연도', '거주지', '가구원수',        // L-N
  '유입경로', '가족관계',               // O-P
  '기존진행', '완료시기',               // Q-R
  '월소득', '배우자소득', '소득유형', '소득상세',  // S-V
  '총채무', '담보채무', '세금체납', '개인채무',   // W-Z
  '협약외채무', '최근6개월신규', '연체기간',     // AA-AC
  '재산총액', '재산상세', '증여상속',       // AD-AF
  '뱃지(플래그)', '기타메모'              // AG-AH
];

// 날짜+순번 ID 생성 (20260313-001)
function generateId(sheet) {
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  let seq = 1;
  if (sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const todayIds = ids.filter(id => String(id).startsWith(today));
    if (todayIds.length > 0) {
      const maxSeq = Math.max(...todayIds.map(id => parseInt(String(id).split('-')[1]) || 0));
      seq = maxSeq + 1;
    }
  }
  return today + '-' + String(seq).padStart(3, '0');
}

// ── 시트 초기 세팅 ──
function setupSheet(sheet) {
  // 헤더
  sheet.appendRow(HEADERS);
  const hRange = sheet.getRange(1, 1, 1, HEADERS.length);
  hRange.setFontWeight('bold')
        .setFontColor('#FFFFFF')
        .setBackground('#1B2A4A')
        .setFontSize(9)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);
  sheet.setFrozenRows(1);

  // 열 너비
  const widths = {
    1:110, 2:70, 3:100,           // ID, 이름, 연락처
    4:90, 5:75, 6:85, 7:75,       // 선택제도, 감면율, 월변제액, 변제기간
    8:80, 9:80, 10:80, 11:80,     // 추천1~4
    12:65, 13:80, 14:55,           // 출생, 거주지, 가구원수
    15:65, 16:80,                  // 유입경로, 가족관계
    17:65, 18:75,                  // 기존진행, 완료시기
    19:85, 20:85, 21:65, 22:120,   // 소득 관련
    23:90, 24:85, 25:80, 26:85,    // 채무 관련
    27:85, 28:90, 29:70,           // 협약외, 최근6개월, 연체
    30:85, 31:120, 32:80,          // 재산 관련
    33:110, 34:160                 // 뱃지, 메모
  };
  Object.entries(widths).forEach(([col, w]) => sheet.setColumnWidth(Number(col), w));

  // 금액 컬럼 숫자 포맷 (#,##0)
  const moneyFormat = '#,##0';
  // 이 컬럼들은 데이터가 들어올 때 적용 (6=월변제액, 19=월소득, 20=배우자소득, 23~27=채무, 28=최근6개월, 30=재산총액)

  // 선택제도 조건부 서식 (D열 = 4번)
  const rules = sheet.getConditionalFormatRules();
  const col4 = sheet.getRange('D2:D1000');
  const colorMap = [
    { text: '개인회생',       bg: '#D4EDDA', fg: '#155724' },  // 초록
    { text: '개인파산·면책',   bg: '#F8D7DA', fg: '#721C24' },  // 빨강
    { text: '새출발기금',     bg: '#D1ECF1', fg: '#0C5460' },  // 파랑
    { text: '신용회복위원회',  bg: '#FFF3CD', fg: '#856404' }   // 노랑
  ];
  colorMap.forEach(c => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(c.text)
        .setBackground(c.bg)
        .setFontColor(c.fg)
        .setBold(true)
        .setRanges([col4])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);

  // 필터
  sheet.getRange(1, 1, 1, HEADERS.length).createFilter();
}

// ── 새 행 서식 적용 ──
function formatNewRow(sheet, rowNum) {
  const numCols = HEADERS.length;
  const rowRange = sheet.getRange(rowNum, 1, 1, numCols);

  // 교대 배경
  const bg = rowNum % 2 === 0 ? '#F8F9FC' : '#FFFFFF';
  rowRange.setBackground(bg)
          .setFontSize(9)
          .setVerticalAlignment('middle')
          .setBorder(true, true, true, true, true, true,
                     '#E0E0E0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(rowNum, 28);

  // 금액 포맷
  const moneyCols = [6, 19, 20, 23, 24, 25, 26, 27, 28, 30]; // 월변제액, 소득들, 채무들, 재산
  moneyCols.forEach(c => {
    sheet.getRange(rowNum, c).setNumberFormat('#,##0');
  });

  // 선택제도 · 이름 · ID 볼드
  sheet.getRange(rowNum, 1).setFontWeight('bold'); // ID
  sheet.getRange(rowNum, 2).setFontWeight('bold'); // 이름
  sheet.getRange(rowNum, 4).setFontWeight('bold'); // 선택제도
  sheet.getRange(rowNum, 4).setHorizontalAlignment('center');
}

// ── POST: 상담 저장 ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      setupSheet(sheet);
    }

    const d = data;
    const id = generateId(sheet);

    // 컬럼 순서에 맞게 행 추가
    sheet.appendRow([
      id,
      d.name || '',
      d.phone || '',
      d.selectedProgram || '',        // 선택제도 ★
      d.reductionRate || '',
      d.monthlyPayment || '',
      d.paymentPeriod || '',
      d.rank1 || '',
      d.rank2 || '',
      d.rank3 || '',
      d.rank4 || '',
      d.birthYear || '',
      d.address || '',
      d.household || '',
      d.channel || '',
      d.family || '',
      d.prevCase || '',
      d.prevCaseDate || '',
      d.income || 0,
      d.spouseIncome || 0,
      d.incomeType || '',
      d.incomeDetail || '',
      d.totalDebt || 0,
      d.securedDebt || 0,
      d.taxDebt || 0,
      d.personalDebt || 0,
      d.nonPartnerDebt || 0,
      d.recentDebt || 0,
      d.overdueDays || 0,
      d.assets || 0,
      d.assetsDetail || '',
      d.inheritance || '',
      d.badges || '',
      d.memo || ''
    ]);

    formatNewRow(sheet, sheet.getLastRow());

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, id: id })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET: 이력 조회 ──
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const numCols = HEADERS.length;

    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: [] })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const action = e.parameter.action || 'list';
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

    if (action === 'list') {
      const lastRow = sheet.getLastRow();
      const startRow = Math.max(2, lastRow - 99);
      const numRows = lastRow - startRow + 1;
      const values = sheet.getRange(startRow, 1, numRows, numCols).getValues();

      const results = values.map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      }).reverse();

      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: results })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'search') {
      const keyword = (e.parameter.q || '').toLowerCase();
      const lastRow = sheet.getLastRow();
      const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

      const results = values.filter(row =>
        row.some(cell => String(cell).toLowerCase().includes(keyword))
      ).map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      }).reverse();

      return ContentService.createTextOutput(
        JSON.stringify({ success: true, data: results })
      ).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 시트 리셋 (기존 데이터 삭제 + 재생성) ──
function reformatSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName(SHEET_NAME);
  if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(SHEET_NAME);
  setupSheet(sheet);
  SpreadsheetApp.getUi().alert('✅ 시트가 초기화되었습니다.');
}
