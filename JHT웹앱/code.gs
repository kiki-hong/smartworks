/**
 * JHT 스마트 업무일지 - 통합 서버 코드 (수정본)
 *
 * 본 코드는 기존 통합 서버 코드에 "견적 조회" 기능을 추가한 버전입니다.
 * 변경 사항:
 *   1) 설정값 기본에 SHEET_QUOTE_WORK 키를 추가하여 견적 계산에 사용할 시트 이름을
 *      지정할 수 있도록 했습니다. 기본값은 '견적서'입니다.
 *   2) processQuoteRoom(roomNo) 함수를 추가하여, 클라이언트에서 전송된 호실 번호를
 *      견적 시트에 기록하고, 계산된 결과(B~K열) 값을 반환합니다. 이 함수는
 *      MOBILE→클라우드→프리젠테이션(MCP) 패턴의 간단한 API로 사용할 수 있습니다.
 */

/***************************************
 * JHT 스마트 업무일지 - 통합 서버 코드
 *  - '설정' 시트 기반 동작
 *  - 일지 ID: DIARY_COUNTER_CELL 사용(락 적용)
 *  - RichText 하이퍼링크(전화/이메일) 지원
 *  - 이미지 업로드(UPLOAD_FOLDER_ID)
 *  - 로깅(LOG_SHEET, LOG_RETENTION_DAYS)
 ***************************************/

/* ----------------------------------------------------------------------
   설정 읽기 및 캐시
   - 설정 시트는 기본 이름이 '설정'이며, A열에 키, B열에 값을 입력합니다.
   - getConfig_()는 스크립트 캐시를 활용해 1분 동안 값을 저장합니다.
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
  if (!cfg.ENABLE_TABS)     cfg.ENABLE_TABS     = '업무일지,호실현황';
  if (!cfg.IMAGE_MAX_MB)    cfg.IMAGE_MAX_MB    = '5';
  if (!cfg.IMAGE_TYPES)     cfg.IMAGE_TYPES     = 'image/png,image/jpeg';
  if (!cfg.DIARY_COUNTER_CELL) cfg.DIARY_COUNTER_CELL = 'B20';
  if (!cfg.LOG_RETENTION_DAYS) cfg.LOG_RETENTION_DAYS = '90';
  if (!cfg.ENABLE_MAIL_CONFIRM)     cfg.ENABLE_MAIL_CONFIRM = 'false'; // 메일 알림 기본 OFF
  if (!cfg.MAIL_SUBJECT_TEMPLATE)   cfg.MAIL_SUBJECT_TEMPLATE = '[JHT] 업무일지 등록 완료 - ID: {{id}}';
  if (!cfg.MAIL_SENDER_NAME)        cfg.MAIL_SENDER_NAME = 'JHT 업무일지';
  if (!cfg.MAIL_BCC)                cfg.MAIL_BCC = ''; // 선택

  // 견적서 PDF 생성/메일 전송과 관련된 기본 설정
  if (!cfg.QUOTE_COMMENT_CELL)    cfg.QUOTE_COMMENT_CELL = 'B4';
  if (!cfg.QUOTE_MAIL_SUBJECT_TEMPLATE) cfg.QUOTE_MAIL_SUBJECT_TEMPLATE = '[JHT] 견적서 - ID: {{id}}';
  // 견적 작업에 사용할 시트 이름. 지정하지 않으면 기본 '견적서'
  if (!cfg.SHEET_QUOTE_WORK) cfg.SHEET_QUOTE_WORK = '견적서';

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
  return String(v).toLowerCase() === 'true';
}
function cfgNumber(key, defVal = 0) {
  const n = Number(cfg(key, ''));
  return isNaN(n) ? defVal : n;
}

/* 레거시 호환 래퍼: 기존 코드에서 _cfg() 등을 사용하고 있다면 그대로 호출 가능 */
function _cfg(key, defVal)   { return cfg(key, defVal); }
function _cfgBool(key, defVal) { return cfgBool(key, defVal); }
function _cfgNum(key, defVal)  { return cfgNumber(key, defVal); }

/* ----------------------------------------------------------------------
   시트 핸들러
   ---------------------------------------------------------------------- */
function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}
function _ss() {
  return getSpreadsheet_();
}
function _sheet(nameKey, defName) {
  const name = cfg(nameKey, defName);
  return _ss().getSheetByName(name);
}

/* ----------------------------------------------------------------------
   클라이언트에 전달할 기본 설정 (앱 제목, 타임존 등)
   ---------------------------------------------------------------------- */
function getClientConfig() {
  return {
    APP_TITLE:    cfg('APP_TITLE', 'JHT 스마트 업무일지'),
    TIMEZONE:     cfg('TIMEZONE', 'Asia/Seoul'),
    DATE_FMT:     cfg('DATE_FMT', 'yyyy-MM-dd'),
    ENABLE_TABS:  cfg('ENABLE_TABS', '업무일지,호실현황'),
    IMAGE_MAX_MB: cfgNumber('IMAGE_MAX_MB', 5),
    IMAGE_TYPES:  cfg('IMAGE_TYPES', 'image/png,image/jpeg, image/jpg'),
    ENABLE_SW:    cfgBool('ENABLE_SW', true),
    CACHE_VERSION: cfg('CACHE_VERSION', 'v1'),
  };
}

/* ----------------------------------------------------------------------
   로그 기록
   ---------------------------------------------------------------------- */
function _appendLog_(level, action, payload) {
  try {
    const logSheetName = cfg('LOG_SHEET', 'log');
    let sh = _ss().getSheetByName(logSheetName);
    if (!sh) { sh = _ss().insertSheet(logSheetName); sh.appendRow(['ts','level','action','payload']); }

    const ts  = new Date();
    const row = [ts, level, action, (typeof payload === 'string' ? payload : JSON.stringify(payload || {}))];
    sh.appendRow(row);

    // 오래된 로그 정리
    const days = cfgNumber('LOG_RETENTION_DAYS', 90);
    if (days > 0) {
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        const values = sh.getRange(2, 1, lastRow - 1, 1).getValues(); // ts만
        const cutoff = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
        // 오래된 행 삭제(아래에서 위로)
        for (let i = values.length; i >= 1; i--) {
          const d = values[i - 1][0];
          if (d && d < cutoff) {
            sh.deleteRow(i + 1);
          }
        }
      }
    }
  } catch (err) {
    // 로그 기록 중 에러가 발생해도 서비스 전체에 영향은 주지 않음
    console.error('log write error:', err);
  }
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
    const ss    = getSpreadsheet_();
    const team  = ss.getSheetByName(cfg('SHEET_TEAM', '팀원'));
    if (!team) return { success: false, message:'팀원 시트를 찾을 수 없습니다.' };

    const rows  = team.getDataRange().getValues();
    const inName  = (name || '').trim();
    const inEmail = (email || '').trim().toLowerCase();

    // 관리자 이메일은 즉시 통과
    const adminList = cfg('ADMIN_EMAILS','').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    if (adminList.includes(inEmail)) {
      _appendLog_('INFO','login_admin_bypass',{ email: inEmail });
      return { success: true, message:'인증 성공(관리자)', user:{ name: inName, email: inEmail } };
    }

    for (let i = 1; i < rows.length; i++) {
      const rowName  = (rows[i][0] || '').toString().trim();
      const rowEmail = (rows[i][1] || '').toString().trim().toLowerCase();
      if (rowName === inName && rowEmail === inEmail) {
        _appendLog_('INFO','login_ok',{ name: inName, email: inEmail });
        return { success: true, message:'인증 성공', user:{ name: rowName, email: rowEmail } };
      }
    }
    _appendLog_('WARN','login_fail',{ name: inName, email: inEmail });
    return { success: false, message:'등록되지 않은 사용자입니다. 이름과 이메일을 확인해주세요.' };
  } catch (err) {
    _appendLog_('ERROR','login_error', String(err));
    return { success: false, message:'인증 처리 중 오류가 발생했습니다.' };
  }
}

/* ----------------------------------------------------------------------
   RichText 링크 변환 (전화번호는 tel:, 이메일은 mailto:)
   ---------------------------------------------------------------------- */
function buildRichTextWithLinks_(text) {
  // 설정값이 비었거나 "(기본값)"이면 내장 기본 정규식 사용
  const phoneFromCfg = (cfg('PHONE_REGEX', '') || '').trim();
  const emailFromCfg = (cfg('EMAIL_REGEX', '') || '').trim();

  const defaultPhone = /(01[0-9][-\.\s]?\d{3,4}[-\.\s]?\d{4})/g;
  const defaultEmail = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;

  function toRegex(src, def) {
    if (!src || src === '(기본값)') return def;
    try {
      // /.../flags 형식이 아니라면 전체를 패턴 문자열로 간주
      if (src.startsWith('/') && src.lastIndexOf('/') > 0) {
        const last = src.lastIndexOf('/');
        const pat = src.slice(1, last);
        const flg = src.slice(last + 1) || 'g';
        return new RegExp(pat, flg.includes('g') ? flg : (flg + 'g'));
      }
      return new RegExp(src, 'g');
    } catch (e) {
      return def; // 설정 오류 시 안전하게 기본값
    }
  }

  const phonePattern = toRegex(phoneFromCfg, defaultPhone);
  const emailPattern = toRegex(emailFromCfg, defaultEmail);

  const str = (text || '').toString();
  const builder = SpreadsheetApp.newRichTextValue().setText(str);
  const matches = [];

  function collect(regex, kind) {
    const re = new RegExp(regex.source, regex.flags); // 확실한 RegExp
    let m;
    while ((m = re.exec(str)) !== null) {
      matches.push({ s: m.index, e: m.index + m[0].length, val: m[0], kind });
    }
  }
  collect(phonePattern, 'phone');
  collect(emailPattern, 'email');

  matches.sort((a, b) => a.s - b.s);
  matches.forEach(({ s, e, val, kind }) => {
    if (kind === 'phone') {
      const clean = val.replace(/[-\.\s]/g, '');
      builder.setLinkUrl(s, e, 'tel:' + clean);
    } else {
      builder.setLinkUrl(s, e, 'mailto:' + val);
    }
  });

  return builder.build();
}


/* ----------------------------------------------------------------------
   일지 ID 생성 (설정 시트의 카운터 셀을 사용)
   ---------------------------------------------------------------------- */
function nextDiaryId_(dateStr) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    // ✅ 설정 캐시 무효화(즉시 최신값 반영)
    CacheService.getScriptCache().remove('CONFIG_JSON');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setSheet = ss.getSheetByName(CFG_SHEET_NAME);

    const dStr = (dateStr || '').toString().trim();
    const baseDate = dStr ? new Date(dStr) : new Date();

    const startStr  = cfg('DIARY_PREFIX_START', '');
    const usePrefix = startStr ? (baseDate >= new Date(startStr)) : false;

    if (usePrefix) {
      // ✅ 숫자만 적어도 안전하게 보정 (예: "21" → "B21")
      let cellA1 = cfg('DIARY_PREFIX_COUNTER_CELL', 'B21');
      if (/^\d+$/.test(String(cellA1).trim())) cellA1 = 'B' + String(cellA1).trim();

      const cell = setSheet.getRange(cellA1);
      let cur = Number(cell.getValue());
      if (isNaN(cur) || cur < 0) cur = 0;

      const next = cur + 1;
      cell.setValue(next);

      const prefix = String(cfg('DIARY_PREFIX', 'h'));
      return prefix + String(next);   // 예: "2509-123"
    } else {
      // ✅ 숫자만 적어도 안전하게 보정 (예: "20" → "B20")
      let cellA1 = cfg('DIARY_COUNTER_CELL', 'B20');
      if (/^\d+$/.test(String(cellA1).trim())) cellA1 = 'B' + String(cellA1).trim();

      const cell = setSheet.getRange(cellA1);
      let cur = Number(cell.getValue());
      if (isNaN(cur) || cur < 0) cur = 0;

      const next = cur + 1;
      cell.setValue(next);

      return next; // 숫자 ID (예: 123)
    }
  } finally {
    lock.releaseLock();
  }
}



/* ----------------------------------------------------------------------
   이미지 데이터(Data URL) → Blob 변환
   ---------------------------------------------------------------------- */
function _dataUrlToBlob_(dataUrl, fallbackName) {
  let mime = MimeType.PNG, bytes;
  try {
    if (String(dataUrl).indexOf('base64,') >= 0) {
      const header = dataUrl.substring(0, dataUrl.indexOf('base64,'));
      const base64 = dataUrl.split('base64,')[1];
      bytes = Utilities.base64Decode(base64);
      const m = header.match(/^data:(.*?);base64,$/i);
      if (m && m[1]) mime = m[1];
    } else {
      bytes = Utilities.base64Decode(dataUrl);
    }
  } catch (e) {
    throw new Error('이미지 디코드 실패');
  }
  return Utilities.newBlob(bytes, mime, fallbackName || 'upload');
}

/* ----------------------------------------------------------------------
   일지 저장 (이미지 없음)
   ---------------------------------------------------------------------- */
function submitWorkDiary(data) {
  try {
    const diary = _sheet('SHEET_DIARY', '일지');
    if (!diary) return { success:false, message:'일지 시트를 찾을 수 없습니다.' };

    const id  = nextDiaryId_(data?.date);
    const row = diary.getLastRow() + 1;

    diary.getRange(row, 1, 1, 9).setValues([[
      String(id),
      data.date       || '',
      data.author     || '',
      data.category   || '',
      '', '', // 고객/업무 RichText는 아래서
      (data.etc || '').toString().replace(/[\r\n]+/g, ' '),
      '',
      data.timestamp || Utilities.formatDate(new Date(), cfg('TIMEZONE','Asia/Seoul'), 'yyyy-MM-dd HH:mm:ss')
    ]]);
    diary.getRange(row, 1).setNumberFormat('@');

    diary.getRange(row, 5).setRichTextValue(buildRichTextWithLinks_(data.customerInfo || ''));
    diary.getRange(row, 6).setRichTextValue(buildRichTextWithLinks_(data.workContent || ''));

    diary.getRange(row, 1, 1, 9).setBorder(null,null,true,null,null,null,'#969696',SpreadsheetApp.BorderStyle.SOLID);

    // 메일 전송
    let emailTried = false, emailOk = false;
    if (cfgBool('ENABLE_MAIL_CONFIRM', false) && data?.authorEmail) {
      emailTried = true;
      emailOk = sendConfirmEmail_(data.authorEmail, data, id, null);
    }

    _appendLog_('INFO','diary_submit',{ id, author:data.author, hasImage:false, emailTried, emailOk });
    return { success:true, message:'업무일지가 성공적으로 저장되었습니다.', id, emailTried, emailOk };
  } catch (err) {
    _appendLog_('ERROR','diary_submit_error', String(err));
    return { success:false, message:'업무일지 저장 중 오류가 발생했습니다.' };
  }
}

function submitWorkDiaryWithImage(data, imageData) {
  try {
    const diary = _sheet('SHEET_DIARY','일지');
    if (!diary) return { success:false, message:'일지 시트를 찾을 수 없습니다.' };

    const id  = nextDiaryId_(data?.date);
    const row = diary.getLastRow() + 1;

    let imageInfo = null;
    if (imageData && imageData.base64 && imageData.fileName) {
      const maxMb  = cfgNumber('IMAGE_MAX_MB', 5);
      let mime     = 'application/octet-stream';
      let raw      = imageData.base64;

      if (raw.startsWith('data:')) {
        const i    = raw.indexOf(',');
        const head = raw.substring(0, i);
        raw        = raw.substring(i + 1);
        const mt   = /data:(.*?);base64/.exec(head);
        if (mt && mt[1]) mime = mt[1];
      } else {
        mime = 'image/png';
      }

      const allowed  = cfg('IMAGE_TYPES','image/png,image/jpeg').split(',').map(s => s.trim());
      if (allowed.length && !allowed.includes(mime)) {
        return { success:false, message:`허용되지 않은 이미지 형식입니다: ${mime}` };
      }
      const bytes   = Utilities.base64Decode(raw);
      const sizeMb  = bytes.length / (1024 * 1024);
      if (sizeMb > maxMb) {
        return { success:false, message:`이미지 용량 초과(${sizeMb.toFixed(2)}MB > ${maxMb}MB)` };
      }

      const blob  = Utilities.newBlob(bytes, mime, imageData.fileName);
      const folderId = cfg('UPLOAD_FOLDER_ID','');
      const file  = folderId ? DriveApp.getFolderById(folderId).createFile(blob)
                             : DriveApp.createFile(blob);
      imageInfo = { id:file.getId(), name:file.getName(), url:file.getUrl(), mime };
    }

    const cleanEtc = (data.etc || '').toString().replace(/[\r\n]+/g, ' ');
    diary.getRange(row, 1, 1, 9).setValues([[
      String(id),
      data.date      || '',
      data.author    || '',
      data.category  || '',
      '', '', // 고객/업무 RichText는 아래서
      cleanEtc,
      '',
      data.timestamp || Utilities.formatDate(new Date(), cfg('TIMEZONE','Asia/Seoul'), 'yyyy-MM-dd HH:mm:ss')
    ]]);
    diary.getRange(row, 1).setNumberFormat('@');

    diary.getRange(row, 5).setRichTextValue(buildRichTextWithLinks_(data.customerInfo || ''));
    diary.getRange(row, 6).setRichTextValue(buildRichTextWithLinks_(data.workContent || ''));
    if (imageInfo) {
      diary.getRange(row, 8).setFormula(`=HYPERLINK("${imageInfo.url}","이미지 보기")`);
    }
    diary.getRange(row, 1, 1, 9).setBorder(null,null,true,null,null,null,'#969696',SpreadsheetApp.BorderStyle.SOLID);

    // 메일 전송
    let emailTried = false, emailOk = false;
    if (cfgBool('ENABLE_MAIL_CONFIRM', false) && data?.authorEmail) {
      emailTried = true;
      emailOk = sendConfirmEmail_(data.authorEmail, data, id, imageInfo);
    }

    _appendLog_('INFO','diary_submit',{ id, author:data.author, hasImage:Boolean(imageInfo), emailTried, emailOk });
    return { success:true, message:'업무일지가 성공적으로 저장되었습니다.', id, imageInfo, emailTried, emailOk };
  } catch (err) {
  _appendLog_('ERROR', 'diary_submit_error', {
    message: err.message,
    stack: err.stack
  });
  return { success:false, message:'업무일지 저장 중 오류가 발생했습니다.' };
}
}


/* ----------------------------------------------------------------------
   기존 데이터에 링크 일괄 적용 (옵션)
   ---------------------------------------------------------------------- */
function applyHyperlinksToExistingData() {
  try {
    const diary = _sheet('SHEET_DIARY','일지');
    if (!diary) { console.log('일지 시트를 찾을 수 없습니다.'); return; }
    const last = diary.getLastRow();
    if (last < 2) { console.log('적용할 데이터가 없습니다.'); return; }

    const vals = diary.getRange(2, 1, last - 1, 8).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = 2 + i;
      diary.getRange(r, 5).setRichTextValue(buildRichTextWithLinks_(vals[i][4]));
      diary.getRange(r, 6).setRichTextValue(buildRichTextWithLinks_(vals[i][5]));
    }
    console.log('기존 데이터 하이퍼링크 적용 완료');
  } catch (err) {
    console.log('applyHyperlinksToExistingData error:', err);
  }
}

/* ----------------------------------------------------------------------
   호실현황 데이터 반환
   ---------------------------------------------------------------------- */
function getRoomStatus() {
  try {
    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(cfg('SHEET_ROOM','호실현황'));
    if (!sheet) return { success:false, message:'호실현황 시트를 찾을 수 없습니다.' };

    const a1 = cfg('ROOM_RANGE','');
    const rangeObj = a1 ? sheet.getRange(a1) : sheet.getDataRange();

    const values      = rangeObj.getValues();
    const backgrounds = rangeObj.getBackgrounds();
    const fontColors  = rangeObj.getFontColors();
    if (!values.length) return { success:false, message:'호실현황 데이터가 없습니다.' };

    const data = values.map((row, i) =>
      row.map((cell, j) => ({
        value: cell,
        backgroundColor: backgrounds[i][j],
        fontColor: fontColors[i][j],
        isEmpty: !cell || String(cell).trim() === ''
      }))
    );
    return { success:true, data, lastUpdate: new Date().toLocaleString('ko-KR') };
  } catch (err) {
    console.error(err);
    return { success:false, message:'호실현황을 가져오는 중 오류가 발생했습니다.' };
  }
}

/* ----------------------------------------------------------------------
   최근 일지 목록 가져오기 (옵션)
   ---------------------------------------------------------------------- */
function getUserDiaries(userName, limit) {
  try {
    limit = limit || 10;
    const diary = _sheet('SHEET_DIARY','일지');
    if (!diary) return { success:false, message:'일지 시트를 찾을 수 없습니다.' };

    const data = diary.getDataRange().getValues();
    const out  = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][2] === userName) {
        out.push({
          id: data[i][0], date: data[i][1], author: data[i][2], category: data[i][3],
          customerInfo: data[i][4], workContent: data[i][5], etc: data[i][6],
          image: data[i][7], // 필요하면 포함
          timestamp: data[i][8] // ← 여기!
        });
        if (out.length >= limit) break;
      }
    }
    return { success:true, diaries: out };
  } catch (err) {
    console.error(err);
    return { success:false, message:'일지 목록을 가져오는 중 오류가 발생했습니다.' };
  }
}


function initializeSheets() {
  try {
    const ss = getSpreadsheet_();

    // 일지 시트
    let diary = ss.getSheetByName(cfg('SHEET_DIARY','일지'));
    if (!diary) {
      diary = ss.insertSheet(cfg('SHEET_DIARY','일지'));
      // 헤더를 9열로 정의: 이미지 열 포함
      diary.getRange(1, 1, 1, 9).setValues([[
        'ID','날짜','작성자','구분','고객정보','업무내용','기타','이미지','타임스탬프'
      ]]);
    }

    // 팀원 시트
    let team = ss.getSheetByName(cfg('SHEET_TEAM','팀원'));
    if (!team) {
      team = ss.insertSheet(cfg('SHEET_TEAM','팀원'));
      team.getRange(1, 1, 1, 2).setValues([['이름','이메일']]);
    }

    // 로그 시트
    const logName = cfg('LOG_SHEET','log');
    if (!ss.getSheetByName(logName)) {
      const lg = ss.insertSheet(logName);
      lg.appendRow(['ts','level','action','payload']);
    }

    return { success:true, message:'시트 초기화가 완료되었습니다.' };
  } catch (err) {
    _appendLog_('ERROR','init_sheets_error', String(err));
    return { success:false, message:'시트 초기화 중 오류가 발생했습니다.' };
  }
}



/* ----------------------------------------------------------------------
   테스트용 팀원 데이터 생성 (옵션)
   ---------------------------------------------------------------------- */
function createTestData() {
  try {
    const team = _sheet('SHEET_TEAM','팀원');
    if (team && team.getLastRow() <= 1) {
      const rows = [
        ['김봉기','kimbg0033@gmail.com'],
        ['이상율','eureka120333@gmail.com'],
        ['김태욱','oldandnew143@gmail.com'],
        ['박승원','qkrsw6666@gmail.com'],
        ['장경화','mickey2820@gmail.com'],
        ['홍덕기','bintan21@gmail.com'],
        ['백종근','dyc8822@daum.net'],
      ];
      team.getRange(2,1,rows.length,2).setValues(rows);
    }
    return { success:true, message:'테스트 데이터가 생성되었습니다.' };
  } catch (err) {
    _appendLog_('ERROR','create_test_data_error', String(err));
    return { success:false, message:'테스트 데이터 생성 중 오류가 발생했습니다.' };
  }
}

/* ----------------------------------------------------------------------
   스프레드시트 편집 트리거
   - 일지 시트의 A~I 열 중 하나가 수정되면 그 행 하단에 테두리를 설정/해제합니다.
   ---------------------------------------------------------------------- */
function onEdit(e) {
  if (!e) return;
  const range = e.range, sheet = range.getSheet();
  const row   = range.getRow(), col = range.getColumn();

  // 일지 시트의 A~I 열(1~9) 수정 시
  if (sheet.getName() === cfg('SHEET_DIARY','일지') && row > 1 && col >= 1 && col <= 9) {
    const r    = sheet.getRange(row, 1, 1, 9);
    const vals = r.getValues()[0];
    const has  = vals.some(v => v !== null && v !== undefined && String(v).trim() !== '');
    r.setBorder(null, null, has, null, null, null, '#969696', SpreadsheetApp.BorderStyle.SOLID);
  }
}

/* ----------------------------------------------------------------------
   관리자 전용 메뉴 (관리자만 보이도록)
   ---------------------------------------------------------------------- */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const adminList = cfg('ADMIN_EMAILS','').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);

  if (adminList.includes(userEmail)) {
    ui.createMenu('관리자 도구')
      .addItem('ID 카운터 리셋', 'menuResetDiaryCounter')
      .addItem('접두사 설정', 'menuSetDiaryPrefix')
      .addItem('접두사 시작일 설정', 'menuSetDiaryPrefixStart')
      .addSeparator()
      .addItem('현재 설정 보기', 'menuShowCurrentConfig')
      .addSeparator()
      .addItem('설정 백업(JSON)', 'menuBackupConfig')
      .addItem('설정 복원(JSON)', 'menuRestoreConfig')
      .addSeparator()
      // ▼▼▼ 여기 3줄 추가 ▼▼▼
      .addItem('로그 정리: 날짜 이전 삭제', 'menuPurgeLogsByDate')
      .addItem('로그 정리: N일 초과 삭제', 'menuPurgeLogsByDays')
      .addItem('로그 정리: 모두 삭제', 'menuClearAllLogs')
      .addToUi();
  }
}


/* ----------------------------------------------------------------------
   ID 카운터 리셋
   ---------------------------------------------------------------------- */
function menuResetDiaryCounter() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('ID 리셋', '새로운 시작 번호를 입력하세요 (예: 0)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const newVal = Number(response.getResponseText()) || 0;
    resetDiaryCounter(newVal);
    ui.alert(`ID 카운터가 ${newVal}으로 리셋되었습니다.`);
  }
}

function resetDiaryCounter(newValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName(CFG_SHEET_NAME);
  let cellA1 = cfg('DIARY_PREFIX_COUNTER_CELL', 'B21');
  if (/^\d+$/.test(String(cellA1).trim())) cellA1 = 'B' + String(cellA1).trim();
  const cell = setSheet.getRange(cellA1);
  cell.setValue(Number(newValue) || 0);
  CacheService.getScriptCache().remove('CONFIG_JSON');
  _appendLog_('INFO','resetDiaryCounter',{ by:Session.getActiveUser().getEmail(), newValue });
  return newValue;
}

/* ----------------------------------------------------------------------
   접두사 설정
   ---------------------------------------------------------------------- */
function menuSetDiaryPrefix() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('접두사 변경', '새 접두사를 입력하세요 (예: 2508-)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const prefix = response.getResponseText().trim();
    updateConfigValue_('DIARY_PREFIX', prefix);
    ui.alert(`접두사가 '${prefix}' 로 변경되었습니다.`);
  }
}

/* ----------------------------------------------------------------------
   접두사 시작일 설정
   ---------------------------------------------------------------------- */
function menuSetDiaryPrefixStart() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('접두사 시작일 변경', 'YYYY-MM-DD 형식으로 입력하세요 (예: 2025-09-01)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const dateStr = response.getResponseText().trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      ui.alert('날짜 형식이 올바르지 않습니다. (예: 2025-09-01)');
      return;
    }
    updateConfigValue_('DIARY_PREFIX_START', dateStr);
    ui.alert(`접두사 시작일이 '${dateStr}' 로 변경되었습니다.`);
  }
}

/* ----------------------------------------------------------------------
   현재 설정 보기
   ---------------------------------------------------------------------- */
function menuShowCurrentConfig() {
  const ui = SpreadsheetApp.getUi();
  const prefix   = cfg('DIARY_PREFIX','(없음)');
  const start    = cfg('DIARY_PREFIX_START','(없음)');
  const counter  = (() => {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const setSheet = ss.getSheetByName(CFG_SHEET_NAME);
      let cellA1 = cfg('DIARY_PREFIX_COUNTER_CELL','B21');
      if (/^\d+$/.test(String(cellA1).trim())) cellA1 = 'B' + String(cellA1).trim();
      return setSheet.getRange(cellA1).getValue();
    } catch(e) {
      return '(읽기 오류)';
    }
  })();

  const msg = 
    `📌 현재 일지 ID 설정\n\n` +
    `- 접두사(DIARY_PREFIX): ${prefix}\n` +
    `- 접두사 시작일(DIARY_PREFIX_START): ${start}\n` +
    `- 현재 카운터(DIARY_PREFIX_COUNTER_CELL): ${counter}`;
  ui.alert(msg);
}

/* ----------------------------------------------------------------------
   설정 백업(JSON)
   ---------------------------------------------------------------------- */
function menuBackupConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setSheet = ss.getSheetByName(CFG_SHEET_NAME);
    if (!setSheet) throw new Error("'설정' 시트를 찾을 수 없습니다.");

    const last = setSheet.getLastRow();
    if (last < 2) throw new Error("설정 데이터가 없습니다.");

    const rows = setSheet.getRange(2,1,last-1,2).getValues();
    const cfgObj = {};
    rows.forEach(([k,v]) => { if (k) cfgObj[k] = v; });

    const json = JSON.stringify(cfgObj, null, 2);
    const file = DriveApp.createFile(`config_backup_${new Date().toISOString()}.json`, json, MimeType.JSON);

    SpreadsheetApp.getUi().alert(`✅ 설정이 JSON 파일로 백업되었습니다.\n\nDrive 파일: ${file.getName()}`);
    _appendLog_('INFO','config_backup',{ by:Session.getActiveUser().getEmail(), file:file.getUrl() });
  } catch (err) {
    SpreadsheetApp.getUi().alert('백업 중 오류: ' + err.message);
  }
}

/* ----------------------------------------------------------------------
   설정 복원(JSON)
   ---------------------------------------------------------------------- */
function menuRestoreConfig() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("⚠️ 설정 복원은 보안상 자동 파일선택이 불가합니다.\n\nJSON 파일 내용을 복사해서 입력창에 붙여넣으세요.");

  const response = ui.prompt('설정 복원', 'JSON 내용을 붙여넣으세요', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    try {
      const jsonText = response.getResponseText();
      const obj = JSON.parse(jsonText);
      for (const [k,v] of Object.entries(obj)) {
        updateConfigValue_(k, v);
      }
      ui.alert('✅ 설정이 JSON으로부터 복원되었습니다.');
      _appendLog_('INFO','config_restore',{ by:Session.getActiveUser().getEmail() });
    } catch(e) {
      ui.alert('❌ JSON 파싱 오류: ' + e.message);
    }
  }
}

/* ----------------------------------------------------------------------
   설정 시트 값 갱신 유틸
   ---------------------------------------------------------------------- */
function updateConfigValue_(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setSheet = ss.getSheetByName(CFG_SHEET_NAME);
  const last = setSheet.getLastRow();
  const range = setSheet.getRange(2,1,last-1,2).getValues();
  for (let i=0; i<range.length; i++) {
    if (range[i][0] === key) {
      setSheet.getRange(i+2,2).setValue(value);
      CacheService.getScriptCache().remove('CONFIG_JSON'); // 캐시 무효화
      _appendLog_('INFO','config_update',{ key, value, by:Session.getActiveUser().getEmail() });
      return;
    }
  }
  // 키가 없으면 새로 추가
  setSheet.appendRow([key, value]);
  CacheService.getScriptCache().remove('CONFIG_JSON');
  _appendLog_('INFO','config_insert',{ key, value, by:Session.getActiveUser().getEmail() });
}

/* =========================
   로그 정리 유틸
   - 로그 시트: cfg('LOG_SHEET','log')
   - 컬럼: ts | level | action | payload
========================= */
function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = cfg('LOG_SHEET','log');
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`'${name}' 시트를 찾을 수 없습니다.`);
  return sh;
}

// YYYY-MM-DD → Date(로컬 00:00:00) 파서
function parseDateYMD_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  const [y,m,d] = s.split('-').map(Number);
  const dt = new Date(y, m-1, d, 0,0,0,0);
  return isNaN(dt.getTime()) ? null : dt;
}

// 핵심 삭제 로직: predicate(tsDate) === true 인 행을 삭제
function purgeLogs_(predicate) {
  const sh = getLogSheet_();
  const last = sh.getLastRow();
  if (last <= 1) return { deleted: 0 };

  // ts만 먼저 읽어서 대상 행 번호 수집
  const tsValues = sh.getRange(2, 1, last - 1, 1).getValues(); // [ [ts], [ts], ... ]
  const toDelete = [];
  for (let i = 0; i < tsValues.length; i++) {
    const ts = tsValues[i][0];
    if (ts && predicate(new Date(ts))) {
      // 실제 행 번호는 헤더를 고려해 +2
      toDelete.push(i + 2);
    }
  }

  // 아래에서 위로 삭제(행번호 변동 방지)
  for (let i = toDelete.length - 1; i >= 0; i--) {
    sh.deleteRow(toDelete[i]);
  }
  return { deleted: toDelete.length };
}

/* =========================
   관리자 메뉴: 로그 정리
========================= */

// 1) 입력한 'YYYY-MM-DD' **이전** 로그 모두 삭제
function menuPurgeLogsByDate() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('로그 정리 - 날짜 이전 삭제',
    '기준 날짜를 YYYY-MM-DD 형식으로 입력하세요. (예: 2025-08-01)\n※ 입력한 날짜 **이전** 로그가 삭제됩니다.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const s = (res.getResponseText() || '').trim();
  const cutoff = parseDateYMD_(s);
  if (!cutoff) {
    ui.alert('날짜 형식이 올바르지 않습니다. 예: 2025-08-01');
    return;
  }
  // cutoff 이전(< cutoff 00:00:00) 삭제
  const { deleted } = purgeLogs_(ts => ts < cutoff);
  _appendLog_('INFO', 'logs_purge_by_date', { cutoff: s, deleted });
  ui.alert(`완료: ${s} 이전 로그 ${deleted}건 삭제되었습니다.`);
}

// 2) 입력한 'N일' **초과** 로그 삭제 (오늘 기준)
function menuPurgeLogsByDays() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('로그 정리 - N일 초과 삭제',
    '삭제 기준 일수를 입력하세요. (예: 90)\n※ 오늘로부터 N일을 초과한 로그가 삭제됩니다.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const n = Number((res.getResponseText() || '').trim());
  if (!Number.isFinite(n) || n < 0) {
    ui.alert('유효한 정수 일수를 입력하세요. 예: 90');
    return;
  }
  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // 오늘 00:00
  cutoff.setDate(cutoff.getDate() - n); // N일 전 00:00
  const { deleted } = purgeLogs_(ts => ts < cutoff);
  _appendLog_('INFO', 'logs_purge_by_days', { days: n, cutoff: cutoff.toISOString().slice(0,10), deleted });
  ui.alert(`완료: 오늘 기준 ${n}일 초과 로그 ${deleted}건 삭제되었습니다.`);
}

// 3) **전체 삭제** (헤더 제외)
function menuClearAllLogs() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('⚠️ 로그 전체 삭제',
    '정말 모든 로그를 삭제하시겠습니까? 계속하려면 대문자로 "DELETE" 를 입력하세요.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  if ((res.getResponseText() || '').trim() !== 'DELETE') {
    ui.alert('취소되었습니다.');
    return;
  }

  const sh = getLogSheet_();
  const last = sh.getLastRow();
  let deleted = 0;
  if (last > 1) {
    sh.deleteRows(2, last - 1);
    deleted = last - 1;
  }
  _appendLog_('WARN', 'logs_clear_all', { deleted });
  ui.alert(`완료: 로그 ${deleted}건이 삭제되었습니다.`);
}

// 간단한 템플릿 치환: {{키}} -> 값
function _tpl_(tpl, vars) {
  return String(tpl).replace(/\{\{(\w+)\}\}/g, (_, k) => (vars[k] != null ? String(vars[k]) : ''));
}

// 확인 메일 보내기 (성공/실패만 반환)
function sendConfirmEmail_(to, data, id, imageInfo) {
  if (!to) return false;
  try {
    const subject = _tpl_(cfg('MAIL_SUBJECT_TEMPLATE','업무일지 등록 완료 - ID: {{id}}'), {
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
    _appendLog_('INFO','mail_sent',{ to, id });
    return true;
  } catch (err) {
    _appendLog_('ERROR','mail_fail', String(err));
    return false;
  }
}

/*******************************************************
 * [견적서 1단계] 템플릿 채우고 지정 범위를 PDF로 생성
 *  - createQuotePdf_(data): 프론트에서 호출
 *  - dataUrl(브라우저 열기/다운로드), Drive 파일 URL 동시 반환
 *******************************************************/

/** 견적 ID 증가(Q101, Q102 ...) */
function nextQuoteId_() {
  const lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setSheet = ss.getSheetByName(CFG_SHEET_NAME);
    let cellA1 = cfg('QUOTE_COUNTER_CELL', 'B30');
    if (/^\d+$/.test(String(cellA1).trim())) cellA1 = 'B' + String(cellA1).trim();

    const cell = setSheet.getRange(cellA1);
    let cur = Number(cell.getValue()); if (isNaN(cur) || cur < 0) cur = 0;
    const next = cur + 1; cell.setValue(next);

    const prefix = String(cfg('QUOTE_PREFIX','Q'));
    return prefix + next;
  } finally { lock.releaseLock(); }
}

/** 시트를 PDF Blob으로 내보내기(범위 지정 지원: &range=A1:J38) */
function exportSheetToPdfBlob_(ssId, gid, filename, rangeA1) {
  const token = ScriptApp.getOAuthToken();
  const params = {
    format: 'pdf', size: 'A4', portrait: 'true',
    fitw: 'true', gridlines: 'false', printtitle: 'false',
    sheetnames: 'false', pagenum: 'UNDEFINED',
    top_margin: '0.50', bottom_margin: '0.50', left_margin: '0.50', right_margin: '0.50',
    gid: gid
  };
  const qs = Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&')
           + (rangeA1 ? '&range=' + encodeURIComponent(rangeA1) : '');
  const url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' + qs;

  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  return res.getBlob().setName(filename || 'quote.pdf');
}

/**
 * 템플릿 채우고 PDF 생성(다운로드용)
 * @param data { customer:{name,email}, issuedDate, items[], author:{name,email}, comment }
 * @return { success, id, totals:{supply,vat,total}, pdf:{id,url,name}, dataUrl }
 */
function createQuotePdf_(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tplName = cfg('SHEET_QUOTE_TEMPLATE', '견적서');
    const tpl = ss.getSheetByName(tplName);
    if (!tpl) return { success:false, message:'견적서 템플릿 시트를 찾을 수 없습니다.' };
    if (!data || !data.items || data.items.length === 0) {
      return { success:false, message:'항목이 없습니다.' };
    }

    // 1) 템플릿 복사
    const id = nextQuoteId_();
    const copy = tpl.copyTo(ss).setName('견적서_' + id);

    // 2) 위치/설정
    const startRow = cfgNumber('QUOTE_TABLE_START_ROW', 6);
    const issuedCell = cfg('QUOTE_ISSUED_AT_CELL', 'J3');

    // 3) 발행일 및 메모
    const issue = data.issuedDate || Utilities.formatDate(new Date(), cfg('TIMEZONE','Asia/Seoul'),'yyyy-MM-dd');
    try { copy.getRange(issuedCell).setValue(issue); } catch (_) {}
    // 메모(comment) 셀 작성: 설정에 따라 셀 위치 지정
    const commentCell = cfg('QUOTE_COMMENT_CELL', 'B4');
    if (data.comment) {
      try { copy.getRange(commentCell).setValue(String(data.comment)); } catch (_) {}
    }

    // 4) 표 영역 클리어
    copy.getRange(startRow, 1, Math.max(1, copy.getMaxRows()-startRow), copy.getLastColumn()).clearContent();

    // 5) 항목 계산/기입
    const rows = data.items;
    const toRows = [];
    let sumSupply = 0, sumVat = 0, sumTotal = 0;

    // 유틸 함수 정의: 평으로 변환, 숫자 변환
    function toPyeongFunc(m2) {
      if (m2) {
        return (Number(m2) / 3.305785).toFixed(2);
      }
      return '';
    }
    function toNumberFunc(v) {
      var x = Number(v);
      return isNaN(x) ? 0 : x;
    }
    // 항목 순회 및 합계 계산
    for (var idx = 0; idx < rows.length; idx++) {
      var it = rows[idx];
      var byP = Number(toPyeongFunc(toNumberFunc(it.by_m2) || 0));   // 분양면적(평)
      var unit = toNumberFunc(it.unitPrice);                   // 평당가(부가세 제외)
      var supply = Math.round(byP * unit);          // 공급금액
      var vat = Math.round(supply * 0.1);           // 부가세
      var total = supply + vat;                     // 분양금액
      sumSupply += supply;
      sumVat += vat;
      sumTotal += total;
      toRows.push([
        it.roomNo || '',
        toNumberFunc(it.jy_m2) || '',
        toPyeongFunc(toNumberFunc(it.jy_m2)) || '',
        toNumberFunc(it.by_m2) || '',
        toPyeongFunc(toNumberFunc(it.by_m2)) || '',
        unit || '',
        supply || '',
        vat || '',
        total || '',
        it.note || ''
      ]);
    }

    if (toRows.length) {
      copy.getRange(startRow, 1, toRows.length, 10).setValues(toRows);        // A~J
      copy.getRange(startRow, 2, toRows.length, 1).setNumberFormat('#,##0.00'); // 전용 m²
      copy.getRange(startRow, 3, toRows.length, 1).setNumberFormat('#,##0.00'); // 전용 평
      copy.getRange(startRow, 4, toRows.length, 1).setNumberFormat('#,##0.00'); // 분양 m²
      copy.getRange(startRow, 5, toRows.length, 1).setNumberFormat('#,##0.00'); // 분양 평
      copy.getRange(startRow, 6, toRows.length, 1).setNumberFormat('#,##0');    // 평당가
      copy.getRange(startRow, 7, toRows.length, 3).setNumberFormat('#,##0');    // 공급/부가/분양
    }

    // 6) 합계 행
    const sumRow = startRow + Math.max(1, toRows.length);
    copy.getRange(sumRow, 1, 1, 10).clearContent();
    copy.getRange(sumRow, 1).setValue('합계');
    copy.getRange(sumRow, 7).setValue(sumSupply).setNumberFormat('#,##0');
    copy.getRange(sumRow, 8).setValue(sumVat).setNumberFormat('#,##0');
    copy.getRange(sumRow, 9).setValue(sumTotal).setNumberFormat('#,##0');

    // 7) 납부조건 테이블 작성 (계약금/잔금) - 합계 다음 두 행 아래부터 시작
    const depStartRow = sumRow + 2;
    // 헤더 및 데이터 행 구성
    const depHeader = ['구분','납부비율','공급금액','부가세','납부금액','납부일자'];
    const depRates = [10, 90];
    const depNames = ['계약금','잔금'];
    const depDates = ['계약시','입주시'];
    const depRows = [];
    for (let i = 0; i < depRates.length; i++) {
      const rate = depRates[i];
      const sPart = Math.round(sumSupply * rate / 100);
      const vPart = Math.round(sumVat * rate / 100);
      const pay = sPart + vPart;
      depRows.push([ depNames[i], rate + '%', sPart, vPart, pay, depDates[i] ]);
    }
    const depTotalRow = ['합계','100%', sumSupply, sumVat, sumTotal, ''];
    const depTable = [depHeader, ...depRows, depTotalRow];
    // 납부조건 테이블을 B열(2열)부터 채움 (헤더 포함)
    copy.getRange(depStartRow, 2, depTable.length, depHeader.length).setValues(depTable);
    // 공급금액/부가세/납부금액 형식 지정: depStartRow+1 행부터 depRows.length+1 행까지, 컬럼 D(E): 4,5,6 relative to sheet
    copy.getRange(depStartRow + 1, 4, depRows.length + 1, 3).setNumberFormat('#,##0');

    // 새 PDF 범위 결정: 납부조건 테이블 끝까지 포함
    const pdfEndRow = depStartRow + depTable.length - 1;
    let pdfRange = cfg('QUOTE_PDF_RANGE', 'A1:J38');
    if (pdfRange.indexOf('{{end}}') >= 0) pdfRange = pdfRange.replace('{{end}}', String(pdfEndRow));

    // 8) PDF 생성
    var custNamePdf = (data && data.customer && data.customer.name) ? String(data.customer.name) : '';
    const pdfName = '견적서_' + custNamePdf + '_' + id + '.pdf';
    const blob = exportSheetToPdfBlob_(ss.getId(), copy.getSheetId(), pdfName, pdfRange);

    // 9) Drive 저장
    const folderId = (cfg('QUOTE_PDF_FOLDER_ID','') || '').trim();
    const file = folderId ? DriveApp.getFolderById(folderId).createFile(blob)
                          : DriveApp.createFile(blob);

    // 10) 임시 시트 삭제(원치 않으면 QUOTE_KEEP_SHEET=true)
    if (!cfgBool('QUOTE_KEEP_SHEET', false)) ss.deleteSheet(copy);

    // 11) 프론트에서 즉시 열/다운로드 가능한 DataURL 포함 반환
    const dataUrl = 'data:application/pdf;base64,' + Utilities.base64Encode(blob.getBytes());

    _appendLog_('INFO','quote_pdf_created',{ id, file:file.getUrl(), total:sumTotal });

    // 12) 이메일 발송: 고객 및 작성자에게 PDF 보내기 (선택)
    try {
      const mailSubject = _tpl_(cfg('QUOTE_MAIL_SUBJECT_TEMPLATE','[JHT] 견적서 - ID: {{id}}'), { id });
      const senderName = cfg('MAIL_SENDER_NAME','JHT 업무일지');
      // 고객명 및 메모 준비
      const custName = (data && data.customer && data.customer.name) ? String(data.customer.name) : '';
      const commentVal = data && data.comment ? String(data.comment) : '';
      // HTML 본문
      let htmlBodyBase = '<p>안녕하세요';
      if (custName) htmlBodyBase += ', ' + custName + '님';
      htmlBodyBase += '.</p><p>첨부된 견적서를 확인해주세요.</p>';
      if (commentVal) {
        const commentHtml = commentVal.replace(/\n/g, '<br>');
        htmlBodyBase += '<p><strong>메모:</strong> ' + commentHtml + '</p>';
      }
      htmlBodyBase += '<p>감사합니다.</p>';
      // Plain 본문
      let plainBodyBase = '안녕하세요';
      if (custName) plainBodyBase += ', ' + custName + '님';
      plainBodyBase += '.\n\n첨부된 견적서를 확인해주세요.';
      if (commentVal) plainBodyBase += '\n\n메모: ' + commentVal;
      plainBodyBase += '\n\n감사합니다.';
      // 메일 내용 옵션 생성
      // 고객에게 메일 전송
      if (data && data.customer && data.customer.email) {
        try {
          var custOpts = {
            to: data.customer.email,
            subject: mailSubject,
            htmlBody: htmlBodyBase,
            body: plainBodyBase,
            attachments: [blob],
            name: senderName,
            noReply: true
          };
          MailApp.sendEmail(custOpts);
          _appendLog_('INFO','quote_mail_sent',{ to:data.customer.email, id });
        } catch (merr) {
          _appendLog_('ERROR','quote_mail_error',{ to:data.customer.email, error: String(merr) });
        }
      }
      // 작성자에게 메일 전송 (고객과 동일하지 않을 때)
      if (data && data.author && data.author.email) {
        var authEmail = String(data.author.email).trim().toLowerCase();
        var custEmail = (data && data.customer && data.customer.email) ? String(data.customer.email).trim().toLowerCase() : '';
        if (!custEmail || authEmail !== custEmail) {
          try {
            var authorOpts = {
              to: data.author.email,
              subject: mailSubject,
              htmlBody: htmlBodyBase,
              body: plainBodyBase,
              attachments: [blob],
              name: senderName,
              noReply: true
            };
            MailApp.sendEmail(authorOpts);
            _appendLog_('INFO','quote_mail_sent',{ to:data.author.email, id });
          } catch (merr) {
            _appendLog_('ERROR','quote_mail_error',{ to:data.author.email, error: String(merr) });
          }
        }
      }
    } catch (mailErr) {
      _appendLog_('ERROR','quote_mail_fail', String(mailErr));
    }

    return {
      success:true,
      id,
      totals:{ supply:sumSupply, vat:sumVat, total:sumTotal },
      pdf:{ id:file.getId(), url:file.getUrl(), name:file.getName() },
      dataUrl
    };
  } catch (err) {
    _appendLog_('ERROR','quote_pdf_error', String(err));
    return { success:false, message:'견적서 PDF 생성 중 오류가 발생했습니다.' };
  }
}

/**
 * DATA 시트에서 호실 정보를 찾아 견적에 필요한 값을 반환
 * @param {string} roomNo 호실 번호
 * @returns {Object} { success, data:{ jy_m2, jy_pyeong, by_m2, by_pyeong,
 *                                unitPrice, supply, vat, total, usage, status } }
 */
function getRoomQuoteInfo_(roomNo) {
  try {
    const sheetName = cfg('SHEET_QUOTE_DATA', 'DATA');  // 설정에 맞춰 DATA 시트명 사용
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      return { success: false, message: 'DATA 시트를 찾을 수 없습니다.' };
    }

    const values = sheet.getDataRange().getValues();
    // 헤더 행 기준으로 컬럼 인덱스 찾기
    const header = values[0];
    const idxRoom  = header.indexOf('호실');
    const idxJyM2  = header.indexOf('전용면적(㎡)');
    const idxJyP   = header.indexOf('전용면적(평)');
    const idxByM2  = header.indexOf('분양면적(㎡)');
    const idxByP   = header.indexOf('분양면적(평)');
    const idxUnit  = header.indexOf('평당가(부가세제외)');
    const idxSup   = header.indexOf('공급금액');
    const idxVat   = header.indexOf('부가세');
    const idxTotal = header.indexOf('분양금액');
    const idxUsage = header.indexOf('용도');
    const idxStat  = header.indexOf('상태');

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (String(row[idxRoom]).trim() === String(roomNo).trim()) {
        return {
          success: true,
          data: {
            jy_m2:    row[idxJyM2],
            jy_pyeong:row[idxJyP],
            by_m2:    row[idxByM2],
            by_pyeong:row[idxByP],
            unitPrice:row[idxUnit],
            supply:   row[idxSup],
            vat:      row[idxVat],
            total:    row[idxTotal],
            usage:    row[idxUsage],
            status:   row[idxStat]
          }
        };
      }
    }
    return { success: false, message: '해당 호실 정보를 찾을 수 없습니다.' };
  } catch (e) {
    return { success: false, message: '오류: ' + e.message };
  }
}

/**
 * 견적 시트에 호실 번호를 기록하고 계산 결과를 반환하는 함수
 *
 * 모바일 앱에서 `google.script.run.processQuoteRoom` 으로 호출하여,
 * 사용자가 입력한 호실 번호를 견적 시트(A열)에 기록하고, 자동으로 계산된
 * B~K열 값을 배열로 돌려줍니다. 견적 시트에는 이미 VLOOKUP이나 INDEX/MATCH
 * 수식이 설정되어 있어야 합니다.
 *
 * @param {string|number} roomNo  호실 번호
 * @returns {Object} { success, row: number, values: Array }
 */
function processQuoteRoom(roomNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // 견적 계산에 사용할 시트 이름을 설정에서 읽거나 기본 '견적서'로 사용
    const sheetName = cfg('SHEET_QUOTE_WORK', cfg('SHEET_QUOTE_TEMPLATE','견적서'));
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success:false, message: `${sheetName} 시트를 찾을 수 없습니다.` };
    // 시작 행: 견적 수식이 시작되는 행. 기본값은 6
    const startRow = cfgNumber('QUOTE_TABLE_START_ROW', 6);
    let writeRow = startRow;
    const lastRow = sheet.getLastRow();
    // A열의 빈 행 찾기: startRow부터 내려가며 첫 번째 빈 셀을 찾음
    for (let r = startRow; r <= lastRow; r++) {
      const val = sheet.getRange(r, 1).getValue();
      if (!val) { writeRow = r; break; }
    }
    // 마지막 행까지 꽉 찬 경우에는 다음 행에 작성
    if (writeRow <= 0 || writeRow > lastRow) {
      writeRow = lastRow + 1;
    }
    // A열에 호실 번호 기록
    sheet.getRange(writeRow, 1).setValue(roomNo);
    // 수식 즉시 계산
    SpreadsheetApp.flush();
    // B~K열 값 읽기 (총 10개)
    const rowVals = sheet.getRange(writeRow, 2, 1, 10).getValues()[0];
    return { success:true, row: writeRow, values: rowVals };
  } catch (err) {
    return { success:false, message: '견적 처리 오류: ' + err.message };
  }
}
