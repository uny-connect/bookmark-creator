/** 시트 이벤트: onOpen 메뉴, onEdit 자동 보정 */

/*******************************************************************************************************
 * [2] 시트 어시스턴트 유틸리티 매뉴얼 (onOpen / onEdit)
 *******************************************************************************************************/
function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ BOOKMARK CREATORS 관리')
    .addItem('✨ 빈칸 자동 채우기 (점포ID/이름/마감일/보증금)', 'fillMissingData')
    .addItem('📩 [재발송]선택한 예약 건 점주 메일', 'resendStoreBookingEmail')
    .addItem('❌ [취소메일]선택한 예약 건 점주 취소 통지 메일', 'sendStoreCancellationEmail') // 👈 추가
    .addItem('🔑 매장별 고유 PIN 6자리 생성', 'generateStorePins') 
    .addItem('🚨 자동 노쇼 일괄 처리 (과거 날짜)', 'checkAndMarkNoShow')
    .addItem('👥 아임웹 신규 회원 동기화', 'syncImwebUsers')
    .addItem('📧 매장 메일 수신 테스트 발송', 'sendTestEmailToStore')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return; 
  const sheet = e.range.getSheet(); 
  if (sheet.getName() !== SHEET.MASTER) return;
  
  const row = e.range.getRow(); 
  const col = e.range.getColumn();
  const val = String(e.range.getValue()).trim();

  // 1️⃣ E열(5번째 열) 점포 ID 입력 시 ➔ F열(6번째 열) 점포명 실시간 자동 기입
  if (col === COL_MASTER.STORE_ID && row >= 3) { 
    const storeId = val.toUpperCase();
    const storeNameCell = sheet.getRange(row, COL_MASTER.STORE_NAME);
    const statusCell = sheet.getRange(row, COL_MASTER.STATUS);
    
    if (storeId === "") {
      storeNameCell.clearContent();
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const restSheet = ss.getSheetByName(SHEET.REST);
      if (restSheet) {
        const restData = restSheet.getDataRange().getValues();
        let matchedName = "식당명 없음";
        
        for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
          const sheetStoreId = String(restData[k][IDX_REST.ID] || '').trim().toUpperCase();
          if (sheetStoreId === storeId) {
            matchedName = String(restData[k][IDX_REST.NAME_KO] || '').trim();
            break;
          }
        }
        storeNameCell.setValue(matchedName);
      }
      
      if (String(statusCell.getValue()).trim() === "") {
        statusCell.setValue("예약대기");
      }
    }
  }

  // 2️⃣ 🎯 F열(6번째 열) 점포명 직접 입력 시 ➔ E열(5번째 열) 점포 ID 실시간 역방향 기입 + F열 공식 명칭 보정
  if (col === COL_MASTER.STORE_NAME && row >= 3) {
    const cleanStoreName = val.replace(/\s+/g, '');
    const storeIdCell = sheet.getRange(row, COL_MASTER.STORE_ID);
    const storeNameCell = sheet.getRange(row, COL_MASTER.STORE_NAME);
    const statusCell = sheet.getRange(row, COL_MASTER.STATUS);

    if (cleanStoreName === "") {
      storeIdCell.clearContent();
    } else {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const restSheet = ss.getSheetByName(SHEET.REST);
      if (restSheet) {
        const restData = restSheet.getDataRange().getValues();
        let matchedId = "";
        let matchedOfficialName = "";

        for (let k = REST_DATA_START_INDEX; k < restData.length; k++) {
          const sheetStoreId = String(restData[k][IDX_REST.ID] || '').trim().toUpperCase();
          const sheetNameKo = String(restData[k][IDX_REST.NAME_KO] || '').trim();
          const sheetNameJp = String(restData[k][IDX_REST.NAME_JP] || '').trim();
          const cleanKo = sheetNameKo.replace(/\s+/g, '');
          const cleanJp = sheetNameJp.replace(/\s+/g, '');

          if ((cleanKo && (cleanKo === cleanStoreName || cleanStoreName.includes(cleanKo) || cleanKo.includes(cleanStoreName))) ||
              (cleanJp && (cleanJp === cleanStoreName || cleanStoreName.includes(cleanJp)))) {
            matchedId = sheetStoreId;
            matchedOfficialName = sheetNameKo;
            break;
          }
        }

        if (matchedId) {
          storeIdCell.setValue(matchedId);
          if (matchedOfficialName && matchedOfficialName !== val) {
            storeNameCell.setValue(matchedOfficialName); // 공식 명칭으로 자동 보정
          }
        }
      }

      if (String(statusCell.getValue()).trim() === "") {
        statusCell.setValue("예약대기");
      }
    }
  }

  // 3️⃣ H열(8번째 열) 방문예정일시 편집 시 J열(마감일), K열(보증금) 자동 연산
  if (col === COL_MASTER.VISIT && row >= 3) { 
    const deadlineCell = sheet.getRange(row, COL_MASTER.DEADLINE);
    const depositCell = sheet.getRange(row, COL_MASTER.DEPOSIT);
    const cellValue = e.range.getValue();
    
    if (cellValue instanceof Date) { 
      const d = new Date(cellValue); 
      d.setDate(d.getDate() + 10); 
      deadlineCell.setValue(d); 
      if (!depositCell.getValue()) depositCell.setValue(10000); // 🎯 10,000원 보정
    } else if (!cellValue) { 
      deadlineCell.clearContent(); 
    } 
  }

  // 4️⃣ L열(12번째 열) 진행 상태 변경 시 슬롯 리셋
  if (col === COL_MASTER.STATUS && row >= 3) { 
    if (val === '예약대기' || val === '일정조율필요') {
      sheet.getRange(row, COL_MASTER.VISIT, 1, MASTER_SLOT_COLS).clearContent();
    }
  }
}
