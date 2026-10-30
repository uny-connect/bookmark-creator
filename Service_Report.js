/** 업체별 월간 보고: Master_Log 집계 → Report_Data(누적) → Report(확인용) → 업체별 PDF(Drive) */

const REPORT_HEADERS = ['月', '店舗ID', '店舗名(JP)', '提出件数', '提出日', 'コンテンツURL', 'Googleマップ口コミURL', 'JP一言レビュー', '状態', '訪問日'];
const REPORT_DEFAULT_FOLDER = '월간보고';

/** 📅 'yyyy-MM' 형식의 전월 */
function _prevMonth_(tz, now) {
  const d = new Date(now || new Date());
  d.setDate(1);
  d.setMonth(d.getMonth() - 1);
  return Utilities.formatDate(d, tz, 'yyyy-MM');
}

function _isValidMonth_(m) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(m || '')); }

/** Settings 시트 A키/B값 조회 (없으면 '') */
function _getSetting_(key) {
  const rows = _getSheetsData([SHEET.SETTINGS])[SHEET.SETTINGS];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][IDX_SETTINGS.KEY] || '').trim().toUpperCase() === key) return String(rows[i][IDX_SETTINGS.VALUE] || '').trim();
  }
  return '';
}

/** Drive 문서/폴더 ID. URL(.../d/ID/edit, .../folders/ID?...)을 그대로 붙여넣어도 ID만 추출 */
function _extractDriveId_(v) {
  const s = String(v || '').trim();
  const m = s.match(/\/(?:d|folders)\/([\w-]+)/) || s.match(/[?&]id=([\w-]+)/);
  return m ? m[1] : s;
}

/** 날짜/문자 셀 → 문자열 (Date 는 yyyy-MM-dd) */
function _cellText_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v === undefined || v === null ? '' : v).trim();
}

/**
 * 순수 집계: Master_Log/Restaurant_List 2차원 배열 → Report_Data 행 배열(IDX_REPORT 순서)
 * - 대상: H열 방문월 == month, L열 상태가 방문완료(제출 대기) / 제출완료
 * - 업체당 제출완료 건은 제출일 순, 제출 대기 건은 뒤에 한 행씩. 건수 = 제출완료 건 수
 */
function buildReportRows_(masterData, restData, month, tz) {
  const names = {};
  for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
    const id = String(restData[k][IDX_REST.ID] || '').trim().toUpperCase();
    if (!id || names[id]) continue;
    names[id] = String(restData[k][IDX_REST.NAME_JP] || '').trim() || String(restData[k][IDX_REST.NAME_KO] || '').trim();
  }

  const stores = {}; // storeId -> { name, done:[], pending:[] }
  for (let i = MASTER_LOG_START_ROW - 1; i < masterData.length; i++) {
    const r = masterData[i];
    const status = String(r[IDX_MASTER.STATUS] || '').trim();
    if (status !== '제출완료' && status !== '방문완료') continue;
    if (_splitVisit_(r[IDX_MASTER.VISIT], tz).dateStr.substring(0, 7) !== month) continue;

    const id = String(r[IDX_MASTER.STORE_ID] || '').trim().toUpperCase();
    if (!stores[id]) {
      stores[id] = { name: id ? (names[id] || '(店舗ID不明: ' + id + ')') : '(店舗ID未入力)', done: [], pending: [] };
      if (id && !names[id]) Logger.log('리포트: Restaurant_List 에 없는 점포ID ' + id);
    }
    const visitDate = _splitVisit_(r[IDX_MASTER.VISIT], tz).dateStr;
    if (status === '제출완료') {
      stores[id].done.push([_cellText_(r[IDX_MASTER.SUBMIT_DATE], tz), _cellText_(r[IDX_MASTER.REVIEW_URL], tz), _cellText_(r[IDX_MASTER.GOOGLE_URL], tz), _cellText_(r[IDX_MASTER.SHORT_REVIEW_JP], tz), visitDate]);
    } else {
      stores[id].pending.push(visitDate);
    }
  }

  const out = [];
  Object.keys(stores).sort().forEach(id => {
    const s = stores[id];
    s.done.sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0));
    s.done.forEach(d => out.push([month, id, s.name, s.done.length, d[0], d[1], d[2], d[3], REPORT_STATE.DONE, d[4]]));
    s.pending.forEach(v => out.push([month, id, s.name, s.done.length, '', '', '', '', REPORT_STATE.PENDING, v]));
  });
  return out;
}

/** Report_Data 시트 확보 (없으면 헤더와 함께 생성, A열 텍스트 서식으로 'yyyy-MM' 날짜 자동변환 방지) */
function _ensureReportDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET.REPORT_DATA);
  if (!sh) {
    sh = ss.insertSheet(SHEET.REPORT_DATA);
    sh.setFrozenRows(1);
  }
  sh.getRange(1, COL_REPORT.MONTH, 1, REPORT_HEADERS.length).setValues([REPORT_HEADERS]).setFontWeight('bold'); // 열 추가 시 기존 시트도 헤더 보정
  sh.getRange(1, COL_REPORT.MONTH, sh.getMaxRows(), 1).setNumberFormat('@');
  return sh;
}

/** 해당 월 집계 → Report_Data 에서 그 월 행만 교체 (멱등). 반환: { month, rows, stores } */
function generateMonthlyReport(targetMonth) {
  const tz = Session.getScriptTimeZone();
  const month = targetMonth || _prevMonth_(tz);
  if (!_isValidMonth_(month)) throw new Error('월 형식 오류(yyyy-MM): ' + month);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = _getSheetsData([SHEET.MASTER, SHEET.REST]);
    const rows = buildReportRows_(data[SHEET.MASTER], data[SHEET.REST], month, tz);

    const sh = _ensureReportDataSheet_();
    const last = sh.getLastRow();
    const kept = last > 1
      ? sh.getRange(2, COL_REPORT.MONTH, last - 1, REPORT_HEADERS.length).getValues().filter(r => _cellText_(r[IDX_REPORT.MONTH], tz).substring(0, 7) !== month && String(r[IDX_REPORT.MONTH]) !== '')
      : [];
    kept.forEach(r => { r[IDX_REPORT.MONTH] = _cellText_(r[IDX_REPORT.MONTH], tz).substring(0, 7); });
    const all = kept.concat(rows);

    if (last > 1) sh.getRange(2, COL_REPORT.MONTH, last - 1, REPORT_HEADERS.length).clearContent();
    if (all.length) sh.getRange(2, COL_REPORT.MONTH, all.length, REPORT_HEADERS.length).setValues(all);

    _refreshReportSheet_(month);
    return { month: month, rows: rows.length, stores: new Set(rows.map(r => r[IDX_REPORT.STORE_ID])).size };
  } finally {
    lock.releaseLock();
  }
}

/** Report 시트: A1 월 드롭다운 + FILTER (내부 점검용) */
function _refreshReportSheet_(month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  let sh = ss.getSheetByName(SHEET.REPORT);
  if (!sh) sh = ss.insertSheet(SHEET.REPORT);

  const dataSh = ss.getSheetByName(SHEET.REPORT_DATA);
  const last = dataSh.getLastRow();
  const months = last > 1
    ? Array.from(new Set(dataSh.getRange(2, COL_REPORT.MONTH, last - 1, 1).getValues().map(r => _cellText_(r[0], tz).substring(0, 7)).filter(Boolean))).sort().reverse()
    : [month];

  const monthCell = sh.getRange('A1');
  monthCell.setNumberFormat('@');
  monthCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(months, true).build());
  monthCell.setValue(month);

  sh.getRange(2, COL_REPORT.MONTH, 1, REPORT_HEADERS.length - 1).setValues([REPORT_HEADERS.slice(1)]).setFontWeight('bold');
  sh.getRange('A3').setFormula('=IFERROR(FILTER(' + SHEET.REPORT_DATA + '!B2:J,' + SHEET.REPORT_DATA + '!A2:A=$A$1),"データなし")');
}

function _getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 보고서 루트 폴더: Settings REPORT_ROOT_FOLDER_ID 우선, 없으면 내 드라이브/월간보고 */
function _getReportRootFolder_() {
  const id = _extractDriveId_(_getSetting_('REPORT_ROOT_FOLDER_ID'));
  return id ? DriveApp.getFolderById(id) : _getOrCreateFolder_(DriveApp.getRootFolder(), REPORT_DEFAULT_FOLDER);
}

function _safeFileName_(s) { return String(s).replace(/[\\\/:*?"<>|]/g, '_').trim(); }
/** 본문·머리글·바닥글의 {{key}} 치환 */
function _docReplace_(doc, key, val) {
  const rep = String(val).replace(/\$/g, '$$$$');
  [doc.getBody(), doc.getHeader(), doc.getFooter()].forEach(part => { if (part) part.replaceText('\\{\\{' + key + '\\}\\}', rep); });
}
/** '2026-09' → '2026年9月' */
function _monthLabelJp_(month) { return month.substring(0, 4) + '年' + Number(month.substring(5, 7)) + '月'; }

/** 업체별 PDF 생성 (Docs 템플릿 복사 → 치환 → 표 채움 → PDF). 반환: { month, folderUrl, files } */
function exportVendorPdfs(targetMonth) {
  const tz = Session.getScriptTimeZone();
  const month = targetMonth || _prevMonth_(tz);
  if (!_isValidMonth_(month)) throw new Error('월 형식 오류(yyyy-MM): ' + month);
  const templateId = _extractDriveId_(_getSetting_('REPORT_TEMPLATE_DOC_ID'));
  if (!templateId) throw new Error('Settings 에 REPORT_TEMPLATE_DOC_ID 가 없습니다.');

  const dataSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.REPORT_DATA);
  const last = dataSh ? dataSh.getLastRow() : 0;
  const rows = last > 1
    ? dataSh.getRange(2, COL_REPORT.MONTH, last - 1, REPORT_HEADERS.length).getValues().filter(r => _cellText_(r[IDX_REPORT.MONTH], tz).substring(0, 7) === month)
    : [];
  if (!rows.length) throw new Error(month + ' 데이터가 Report_Data 에 없습니다. 먼저 리포트를 생성하세요.');

  const byStore = {};
  rows.forEach(r => { (byStore[r[IDX_REPORT.STORE_ID]] = byStore[r[IDX_REPORT.STORE_ID]] || []).push(r); });

  const folder = _getOrCreateFolder_(_getReportRootFolder_(), month);
  const files = [];
  Object.keys(byStore).forEach(id => {
    const list = byStore[id];
    const name = list[0][IDX_REPORT.STORE_NAME_JP];
    const pdfName = _safeFileName_(name + '_' + month) + '.pdf';
    const copy = DriveApp.getFileById(templateId).makeCopy('tmp_' + pdfName, folder);
    try {
      const doc = DocumentApp.openById(copy.getId());
      const body = doc.getBody();
      _docReplace_(doc, '업체명', name);
      _docReplace_(doc, '월', _monthLabelJp_(month));
      _docReplace_(doc, '제출 건수', list[0][IDX_REPORT.COUNT]);

      const table = body.getTables()[0];
      if (table) {
        const sample = table.getRow(1);
        list.forEach((r, n) => {
          const tr = table.appendTableRow(sample.copy());
          const pending = r[IDX_REPORT.STATE] === REPORT_STATE.PENDING;
          // 템플릿 열: No. / 訪問日 / 提出日 / コンテンツリンク / GoogleMapリンク / 一言レビュー
          const texts = [
            ('0' + (n + 1)).slice(-2),
            _cellText_(r[IDX_REPORT.VISIT_DATE], tz) || '-',
            pending ? '提出待ち' : _cellText_(r[IDX_REPORT.SUBMIT_DATE], tz),
            r[IDX_REPORT.REVIEW_URL] ? 'リンク' : '-',
            r[IDX_REPORT.GOOGLE_URL] ? 'リンク' : '-',
            r[IDX_REPORT.REVIEW_JP] ? String(r[IDX_REPORT.REVIEW_JP]) : '-',
          ];
          const urls = [null, null, null, r[IDX_REPORT.REVIEW_URL], r[IDX_REPORT.GOOGLE_URL], null];
          texts.forEach((t, c) => {
            const cell = tr.getCell(c);
            cell.setText(t);
            if (urls[c]) cell.editAsText().setLinkUrl(String(urls[c]));
          });
        });
        table.removeRow(1);
      }
      doc.saveAndClose();

      const it = folder.getFilesByName(pdfName);
      while (it.hasNext()) it.next().setTrashed(true); // 재실행 시 중복 PDF 방지
      files.push(folder.createFile(copy.getAs('application/pdf').setName(pdfName)).getUrl());
    } finally {
      copy.setTrashed(true);
    }
  });
  return { month: month, folderUrl: folder.getUrl(), files: files };
}

/** 📋 시트 메뉴: 월 선택 → 집계 재생성 */
function runMonthlyReportMenu() {
  const ui = SpreadsheetApp.getUi();
  const month = _askReportMonth_(ui, '리포트 데이터 재생성');
  if (!month) return;
  try {
    const r = generateMonthlyReport(month);
    ui.alert('✅ ' + r.month + ' 집계 완료\n업체 ' + r.stores + '곳 / ' + r.rows + '행\n(Report 시트에서 확인)');
  } catch (e) { ui.alert('❌ 실패: ' + e.message); }
}

/** 📄 시트 메뉴: 월 선택 → 업체별 PDF 생성 */
function runMonthlyReportPdfMenu() {
  const ui = SpreadsheetApp.getUi();
  const month = _askReportMonth_(ui, '업체별 PDF 생성');
  if (!month) return;
  try {
    const r = exportVendorPdfs(month);
    ui.alert('✅ PDF ' + r.files.length + '개 저장\n' + r.folderUrl);
  } catch (e) { ui.alert('❌ 실패: ' + e.message); }
}

function _askReportMonth_(ui, title) {
  const def = _prevMonth_(Session.getScriptTimeZone());
  const res = ui.prompt(title, '대상 월 (yyyy-MM). 비우면 ' + def, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return '';
  const m = String(res.getResponseText() || '').trim() || def;
  if (!_isValidMonth_(m)) { ui.alert('월 형식이 올바르지 않습니다: ' + m); return ''; }
  return m;
}

/** ⏰ 트리거 본체: 전월 집계 + PDF (UI 호출 없음) */
function runMonthlyReportTrigger() {
  try {
    const g = generateMonthlyReport();
    const p = exportVendorPdfs(g.month);
    Logger.log('월간 리포트 완료 ' + g.month + ' 업체 ' + g.stores + '곳 / PDF ' + p.files.length + '개 / ' + p.folderUrl);
  } catch (e) {
    Logger.log('월간 리포트 실패: ' + e.toString());
  }
}

/** ⏰ 매월 15일 9시 트리거 설치 (기존 동일 트리거는 교체) */
function installMonthlyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'runMonthlyReportTrigger') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runMonthlyReportTrigger').timeBased().onMonthDay(15).atHour(9).create();
  SpreadsheetApp.getUi().alert('✅ 매월 15일 9시 월간 리포트 트리거를 설치했습니다.');
}
