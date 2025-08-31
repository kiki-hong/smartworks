/**
 * JHT 스마트 업무일지 - 통합 서버 코드 (견적서 기능 제외, 일지보기/삭제 기능 추가)
 */

/***************************************
 * JHT 스마트 업무일지 - 통합 서버 코드
 *  - '설정' 시트 기반 동작
 *  - 일지 ID: DIARY_COUNTER_CELL 사용(락 적용)
 *  - RichText 하이퍼링크(전화/이메일) 지원
 *  - 이미지 업<ctrl61>pload(UPLOAD_FOLDER_ID)
 *  - 로깅(LOG_SHEET, LOG_RETENTION_DAYS)
 ***************************************/

/* ----------------------------------------------------------------------
   설정 읽기 및 캐시
   ---------------------------------------------------------------------- */
const CFG_SHEET_NAME   = '설정';
const CFG_CACHE_TTL_SEC = 60; // 캐시 1분

function getConfig_() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('CONFIG_JSON');
  if (cached) return JSON.parse(cached);

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const cfgSheet = ss.getSheetByName(CFG_SHEET_NAME);
  if (!cfgSheet) throw new Error(`'${CFG_SHEET_NAME}' 시트를 찾을 수 없습니다.`);

  const last  = cfgSheet.getLastRow();
  const range = cfgSheet.getRange(2, 1, Math.max(0, last - 1), 2); // A:키, B:값
  const rows  = range.getValues();

  const cfg = {};
  rows.forEach(([k, v]) => {
    const key = (k || '').toString().trim();
    if (!key) return;
    cfg[key] = (v == null) ? '' : v.toString();
  });

  // 기본값 설정
  if (!cfg.TIMEZONE)        cfg.TIMEZONE        = 'Asia/Seoul';
  if (!cfg.DATE_FMT)        cfg.DATE_FMT        = 'yyyy-MM-dd';
  if (!cfg.SHEET_DIARY)     cfg.SHEET_DIARY     = '일지';
  if (!cfg.SHEET_TEAM)      cfg.SHEET_TEAM      = '팀원';
  if (!cfg.SHEET_ROOM)      cfg.SHEET_ROOM      = '호실현황';
  if (!cfg.ENABLE_TABS)     cfg.ENABLE_TABS     = '업무일지,호실현황,일지보기';
  if (!cfg.IMAGE_MAX_MB)    cfg.IMAGE_MAX_MB    = '5';
  if (!cfg.IMAGE_TYPES)     cfg.IMAGE_TYPES     = 'image/png,image/jpeg';
  if (!cfg.DIARY_COUNTER_CELL) cfg.DIARY_COUNTER_CELL = 'B20';
  if (!cfg.LOG_RETENTION_DAYS) cfg.LOG_RETENTION_DAYS = '90';
  if (!cfg.ENABLE_MAIL_CONFIRM)     cfg.ENABLE_MAIL_CONFIRM = 'false'; // 메일 알림 기본 OFF
  if (!cfg.MAIL_SUBJECT_TEMPLATE)   cfg.MAIL_SUBJECT_TEMPLATE = '[JHT] 업무일지 등록 완료 - ID: {{id}}';
  if (!cfg.MAIL_SENDER_NAME)        cfg.MAIL_SENDER_NAME = 'JHT 업무일지';
  if (!cfg.MAIL_BCC)                cfg.MAIL_BCC = ''; // 선택

  cache.put('CONFIG_JSON', JSON.stringify(cfg), CFG_CACHE_TTL_SEC);
  return cfg;
}

function cfg(key, defVal) {
  const c = getConfig_();
  const v = c[key];
  return (v === undefined || v === '') ? defVal : v;
}
function cfgBool(key, defVal = false) {
  const v = cfg(key, '');
  if (v === '') return defVal;
  return String(v).trim().toLowerCase() === 'true';
}
function cfgNumber(key, defVal = 0) {
  const n = Number(cfg(key, ''));
  return isNaN(n) ? defVal : n;
}

/* ----------------------------------------------------------------------
   웹앱 엔트리
   ---------------------------------------------------------------------- */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(cfg('APP_TITLE','JHT 스마트 업무일지'))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ----------------------------------------------------------------------
   사용자 인증 (팀원 시트 기반)
   ---------------------------------------------------------------------- */
function validateUser(name, email) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const team  = ss.getSheetByName(cfg('SHEET_TEAM', '팀원'));
    if (!team) return { success: false, message:'팀원 시트를 찾을 수 없습니다.' };

    const rows  = team.getDataRange().getValues();
    const inName  = (name || '').trim();
    const inEmail = (email || '').trim().toLowerCase();

    const adminList = cfg('ADMIN_EMAILS','').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    if (adminList.includes(inEmail)) {
      return { success: true, message:'인증 성공(관리자)', user:{ name: inName, email: inEmail } };
    }

    for (let i = 1; i < rows.length; i++) {
      const rowName  = (rows[i][0] || '').toString().trim();
      const rowEmail = (rows[i][1] || '').toString().trim().toLowerCase();
      if (rowName === inName && rowEmail === inEmail) {
        return { success: true, message:'인증 성공', user:{ name: rowName, email: rowEmail } };
      }
    }
    return { success: false, message:'등록되지 않은 사용자입니다. 이름과 이메일을 확인해주세요.' };
  } catch (err) {
    return { success: false, message:'인증 처리 중 오류가 발생했습니다.' };
  }
}


/* ----------------------------------------------------------------------
   일지 저장
   ---------------------------------------------------------------------- */
function submitWorkDiary(data) {
  try {
    const diary = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_DIARY', '일지'));
    if (!diary) return { success:false, message:'일지 시트를 찾을 수 없습니다.' };

    const id  = nextDiaryId_(data?.date);
    const row = diary.getLastRow() + 1;

    diary.getRange(row, 1, 1, 9).setValues([[ 
      String(id),
      data.date       || '',
      data.author     || '',
      data.category   || '',
      data.customerInfo || '',
      data.workContent || '',
      data.etc || '',
      '', // Image placeholder
      data.timestamp || new Date()
    ]]);
    diary.getRange(row, 1).setNumberFormat('@');

    let emailTried = false, emailOk = false;
    if (cfgBool('ENABLE_MAIL_CONFIRM', false) && data?.authorEmail) {
      emailTried = true;
      emailOk = sendConfirmEmail_(data.authorEmail, data, id, null);
    }

    return { success:true, message:'업무일지가 성공적으로 저장되었습니다.', id, emailTried, emailOk };
  } catch (err) {
    return { success:false, message:'업무일지 저장 중 오류가 발생했습니다.' };
  }
}

function submitWorkDiaryWithImage(data, imageData) {
  try {
    const diary = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_DIARY','일지'));
    if (!diary) return { success:false, message:'일지 시트를 찾을 수 없습니다.' };

    const id  = nextDiaryId_(data?.date);
    const row = diary.getLastRow() + 1;

    let imageInfo = null;
    if (imageData && imageData.base64 && imageData.fileName) {
      const blob  = Utilities.newBlob(Utilities.base64Decode(imageData.base64.split(',')[1]), 'image/jpeg', imageData.fileName);
      const folderId = cfg('UPLOAD_FOLDER_ID','');
      const file  = folderId ? DriveApp.getFolderById(folderId).createFile(blob) : DriveApp.createFile(blob);
      imageInfo = { id:file.getId(), name:file.getName(), url:file.getUrl() };
    }

    const imageUrl = imageInfo ? `=HYPERLINK("${imageInfo.url}","이미지 보기")` : '';

    diary.getRange(row, 1, 1, 9).setValues([[ 
      String(id),
      data.date      || '',
      data.author    || '',
      data.category  || '',
      data.customerInfo || '',
      data.workContent || '',
      data.etc || '',
      imageUrl,
      data.timestamp || new Date()
    ]]);
    diary.getRange(row, 1).setNumberFormat('@');

    let emailTried = false, emailOk = false;
    if (cfgBool('ENABLE_MAIL_CONFIRM', false) && data?.authorEmail) {
      emailTried = true;
      emailOk = sendConfirmEmail_(data.authorEmail, data, id, imageInfo);
    }

    return { success:true, message:'업무일지가 성공적으로 저장되었습니다.', id, imageInfo, emailTried, emailOk };
  } catch (err) {
    return { success:false, message:'업무일지 저장 중 오류가 발생했습니다.' };
  }
}

/* ----------------------------------------------------------------------
   호실현황 데이터 반환
   ---------------------------------------------------------------------- */
function getRoomStatus() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_ROOM','호실현황'));
    if (!sheet) return { success:false, message:'호실현황 시트를 찾을 수 없습니다.' };

    const range = sheet.getDataRange();
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const fontColors = range.getFontColors();

    const data = values.map((row, i) => 
      row.map((cell, j) => ({
        value: cell,
        backgroundColor: backgrounds[i][j],
        fontColor: fontColors[i][j]
      }))
    );
    return { success:true, data, lastUpdate: new Date().toLocaleString('ko-KR') };
  } catch (err) {
    return { success:false, message:'호실현황을 가져오는 중 오류가 발생했습니다.' };
  }
}

/* ----------------------------------------------------------------------
   일지보기 기능
   ---------------------------------------------------------------------- */

function getFilterOptions() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_DIARY', '일지'));
    if (!sheet) return { success: false, message: '일지 시트를 찾을 수 없습니다.' };
    if (sheet.getLastRow() < 2) return { success: true, data: { dates: ['전체'], authors: ['전체'], categories: ['전체'] } };
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    
    const dates = new Set();
    const authors = new Set();
    const categories = new Set();

    for (let i = 0; i < data.length; i++) {
      if(data[i][1]) dates.add(Utilities.formatDate(new Date(data[i][1]), cfg('TIMEZONE', 'Asia/Seoul'), 'yyyy-MM-dd'));
      if(data[i][2]) authors.add(data[i][2]);
      if(data[i][3]) categories.add(data[i][3]);
    }

    return {
      success: true,
      data: {
        dates: ['전체', ...Array.from(dates).sort().reverse()],
        authors: ['전체', ...Array.from(authors).sort()],
        categories: ['전체', ...Array.from(categories).sort()]
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getDiaryEntries(filters) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_DIARY', '일지'));
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8);
    const values = range.getValues();
    const richTextValues = range.getRichTextValues();

    const allEntries = values.map((row, i) => {
      const richTextRow = richTextValues[i];
      return {
        id: row[0],
        date: row[1] ? Utilities.formatDate(new Date(row[1]), cfg('TIMEZONE', 'Asia/Seoul'), 'yyyy-MM-dd') : '',
        author: row[2],
        category: row[3],
        customerInfo: row[4],
        workContent: row[5],
        etc: row[6],
        image: richTextRow[7] ? richTextRow[7].getLinkUrl() : null
      };
    });

    const filteredData = allEntries.filter(entry => {
      if (filters.date && filters.date !== '전체' && entry.date !== filters.date) return false;
      if (filters.author && filters.author !== '전체' && entry.author !== filters.author) return false;
      if (filters.category && filters.category !== '전체' && entry.category !== filters.category) return false;
      if (filters.customerInfo && !entry.customerInfo.includes(filters.customerInfo)) return false;
      return true;
    });

    return { success: true, data: filteredData };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteDiaryEntry(id) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { success: false, message: '사용자 정보를 확인할 수 없습니다.' };

    const teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_TEAM', '팀원'));
    if (!teamSheet) return { success: false, message: '팀원 시트를 찾을 수 없습니다.' };
    const teamData = teamSheet.getDataRange().getValues();
    let userName = '';
    for (let i = 1; i < teamData.length; i++) {
      if (teamData[i][1].toLowerCase() === userEmail.toLowerCase()) {
        userName = teamData[i][0];
        break;
      }
    }

    if (!userName) return { success: false, message: '팀원 목록에서 사용자를 찾을 수 없습니다.' };

    const diarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg('SHEET_DIARY', '일지'));
    if (!diarySheet) return { success: false, message: '일지 시트를 찾을 수 없습니다.' };
    const diaryData = diarySheet.getDataRange().getValues();

    for (let i = diaryData.length - 1; i >= 1; i--) {
      if (String(diaryData[i][0]) === String(id)) {
        if (diaryData[i][2] === userName) {
          diarySheet.deleteRow(i + 1);
          return { success: true, message: '일지가 삭제되었습니다.' };
        } else {
          return { success: false, message: '본인이 작성한 일지만 삭제할 수 있습니다.' };
        }
      }
    }
    return { success: false, message: '삭제할 일지를 찾을 수 없습니다.' };
  } catch (e) {
    return { success: false, message: '삭제 중 오류가 발생했습니다: ' + e.message };
  }
}


/* ----------------------------------------------------------------------
   기타 유틸리티
   ---------------------------------------------------------------------- */

function nextDiaryId_(dateStr) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setSheet = ss.getSheetByName(cfg('CFG_SHEET_NAME', '설정'));
    let cellA1 = cfg('DIARY_COUNTER_CELL', 'B20');
    if (/^\d+$/.test(String(cellA1).trim())) cellA1 = 'B' + String(cellA1).trim();
    const cell = setSheet.getRange(cellA1);
    let cur = Number(cell.getValue());
    if (isNaN(cur) || cur < 0) cur = 0;
    const next = cur + 1;
    cell.setValue(next);
    return next;
  } finally {
    lock.releaseLock();
  }
}

function _tpl_(tpl, vars) {
  return String(tpl).replace(/\{\{\w+\}\}/g, (_, k) => (vars[k] != null ? String(vars[k]) : ''));
}

function sendConfirmEmail_(to, data, id, imageInfo) {
  if (!to) return false;
  try {
    const subject = _tpl_(cfg('MAIL_SUBJECT_TEMPLATE','[JHT] 업무일지 등록 완료 - ID: {{id}}'), {
      id, date: data.date, author: data.author, category: data.category
    });

    const senderName = cfg('MAIL_SENDER_NAME','JHT 업무일지');
    const bcc = (cfg('MAIL_BCC','') || '').trim();

    const plain = [
      `안녕하세요, ${data.author}님.`, 
      ``,
      `업무일지 등록이 완료되었습니다.`, 
      ``,
      `ID: ${id}`, 
      `날짜: ${data.date}`, 
      `구분: ${data.category}`, 
      `고객정보: ${data.customerInfo || '-'}`, 
      `업무내용: ${data.workContent || '-'}`, 
      `기타: ${data.etc || '-'}`, 
      `타임스탬프: ${data.timestamp}`, 
      `이미지: ${imageInfo?.url || '없음'}`, 
      ``,
      `감사합니다.`, 
      `- ${senderName}`
    ].join('\n');

    const html = `
      <div style="font-family:Segoe UI,Apple SD Gothic Neo,Malgun Gothic,Arial; line-height:1.6; color:#222">
        <h2 style="margin:0 0 12px;">업무일지 등록 완료</h2>
        <table style="border-collapse:collapse; width:100%; max-width:640px;">
          <tbody>
            <tr><td style="padding:6px 8px; background:#f5f7fb; width:140px;">ID</td><td style="padding:6px 8px;">${id}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">날짜</td><td style="padding:6px 8px;">${data.date}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">구분</td><td style="padding:6px 8px;">${data.category}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">작성자</td><td style="padding:6px 8px;">${data.author} &lt;${data.authorEmail || ''}&gt;</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">고객정보</td><td style="padding:6px 8px; white-space:pre-wrap;">${(data.customerInfo||'-')}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">업무내용</td><td style="padding:6px 8px; white-space:pre-wrap;">${(data.workContent||'-')}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">기타</td><td style="padding:6px 8px; white-space:pre-wrap;">${(data.etc||'-')}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">타임스탬프</td><td style="padding:6px 8px;">${data.timestamp}</td></tr>
            <tr><td style="padding:6px 8px; background:#f5f7fb;">이미지</td><td style="padding:6px 8px;">${imageInfo?.url ? `<a href="${imageInfo.url}" target="_blank">이미지 보기</a>` : '없음'}</td></tr>
          </tbody>
        </table>
        <p style="color:#666; margin-top:12px;">이 메일은 시스템에서 자동 발송되었습니다.</p>
      </div>`.trim();

    MailApp.sendEmail({
      to,
      subject,
      name: senderName,
      htmlBody: html,
      body: plain,
      noReply: true,
      ...(bcc ? { bcc } : {})
    });
    return true;
  } catch (err) {
    return false;
  }
}
