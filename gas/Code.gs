/**
 * ==========================================================================
 * 강동어울림복지관 홍보 통합실적 관리 앱 - Google Apps Script (GAS) 백엔드 Code.gs
 * ==========================================================================
 * - 역할: 
 *   1. 복지관 7개 게시판 정기 스크래핑 (GNUBoard Table & Gallery 스킨 완벽 대응)
 *   2. post_id 기반 중복 제거 및 Google Sheets 적재
 *   3. KST 기준 일별/카테고리별 실적 인덱스 기반 인코딩 안전 카운팅
 *   4. 프론트엔드(GitHub Pages) 연결용 Web App JSON API (doGet)
 * ==========================================================================
 */

// 1. 대상 게시판 스펙 정의 (인덱스 기반 매핑)
const TARGET_CONFIGS = [
  { index: 0, category: '공지사항', mainCategory: '어울림 소식', subCategory: '공지사항', url: 'https://gde.or.kr/notice', boTable: 'notice' },
  { index: 1, category: '공시자료', mainCategory: '어울림 소식', subCategory: '공지사항', url: 'https://gde.or.kr/infoopen', boTable: 'infoopen' },
  { index: 2, category: '인재채용', mainCategory: '어울림 소식', subCategory: '공지사항', url: 'https://gde.or.kr/recruitment', boTable: 'recruitment' },
  { index: 3, category: '이용인모집', mainCategory: '어울림 소식', subCategory: '이용인 모집', url: 'https://gde.or.kr/program', boTable: 'program' },
  { index: 4, category: '정보안내', mainCategory: '어울림 소식', subCategory: '정보안내', url: 'https://gde.or.kr/information', boTable: 'information' },
  { index: 5, category: '갤러리(전체)', mainCategory: '어울림 갤러리', subCategory: '전체', url: 'https://gde.or.kr/gallery', boTable: 'gallery' },
  { index: 6, category: '이용상담문의', mainCategory: '참여하기', subCategory: '이용상담문의', url: 'https://gde.or.kr/counseling', boTable: 'counseling' }
];

const CATEGORIES = [
  '공지사항', '공시자료', '인재채용', '이용인모집', '정보안내', '갤러리(전체)', '이용상담문의'
];

/**
 * 2. 초기 구글 스프레드시트 구조 세팅
 * (최초 1회 스크립트 편집기에서 initSheets() 함수 실행)
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 수집이력 시트
  let historySheet = ss.getSheetByName('수집이력');
  if (!historySheet) {
    historySheet = ss.insertSheet('수집이력');
  }
  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow(['post_id', '대분류', '중분류', '세부항목', '제목', '게시일(KST)', '수집시각']);
    historySheet.getRange('A1:G1').setBackground('#E2E8F0').setFontWeight('bold');
  }

  // ② 일별집계 시트
  let dailySheet = ss.getSheetByName('일별집계');
  if (!dailySheet) {
    dailySheet = ss.insertSheet('일별집계');
  }
  if (dailySheet.getLastRow() === 0) {
    dailySheet.appendRow(['날짜', '공지사항', '공시자료', '인재채용', '이용인모집', '정보안내', '갤러리(전체)', '이용상담문의', '합계']);
    dailySheet.getRange('A1:I1').setBackground('#CBD5E1').setFontWeight('bold');
  }

  // ③ config 시트
  let configSheet = ss.getSheetByName('config');
  if (!configSheet) {
    configSheet = ss.insertSheet('config');
  }
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['세부항목', '대분류', 'URL', 'boTable', '비고']);
    TARGET_CONFIGS.forEach(conf => {
      configSheet.appendRow([conf.category, conf.mainCategory, conf.url, conf.boTable, '그누보드']);
    });
    configSheet.getRange('A1:E1').setBackground('#E2E8F0').setFontWeight('bold');
  }

  Logger.log('시트 초기화 세팅 완료!');
}

/**
 * 3. 메인 게시판 스크래핑 및 카운팅 엔진
 * (15분 주기 시간 기반 트리거로 자동 실행)
 */
function runScraper() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName('수집이력') || ss.insertSheet('수집이력');
  const dailySheet = ss.getSheetByName('일별집계') || ss.insertSheet('일별집계');

  // 기존 수집된 post_id 목록 (중복 제거용 Set)
  const existingPostIds = new Set();
  if (historySheet.getLastRow() > 1) {
    const historyData = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, 1).getValues();
    historyData.forEach(row => {
      if (row[0]) existingPostIds.add(String(row[0]).trim());
    });
  }

  const kstNow = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  let newCollectedCount = 0;

  // 7개 대상 URL 스크래핑
  TARGET_CONFIGS.forEach(conf => {
    try {
      const response = UrlFetchApp.fetch(conf.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
      });

      if (response.getResponseCode() !== 200) {
        Logger.log(`HTTP 경고 (${conf.category}): ${response.getResponseCode()}`);
        return;
      }

      const html = response.getContentText('UTF-8');
      const posts = parseGnuboardHtml(html, conf);

      posts.forEach(post => {
        const postId = `${conf.boTable}_${post.wrId}`;
        
        // 중복검사
        if (!existingPostIds.has(postId)) {
          existingPostIds.add(postId);
          newCollectedCount++;

          // 수집이력 행 추가
          historySheet.appendRow([
            postId,
            conf.mainCategory,
            conf.subCategory,
            conf.category,
            post.title,
            post.date,
            kstNow
          ]);

          // 인덱스 기반 일별 집계 반영 (인코딩 안전 방식)
          updateDailyAggregation(dailySheet, post.date, conf.index);
        }
      });
    } catch (e) {
      Logger.log(`스크래핑 에러 [${conf.category}]: ${e.toString()}`);
    }
  });

  Logger.log(`스크래핑 작업 완료! 신규 수집건수: ${newCollectedCount}건`);
}

/**
 * 4. GNUBoard HTML 파싱 함수 (Table 리스트 & Gallery 카드 스킨 완벽 대응)
 */
function parseGnuboardHtml(html, conf) {
  const posts = [];
  
  // 주석 제거 (주석 내부 작성일/링크 오탐 방지)
  const cleanHtml = html.replace(/<!--[\s\S]*?-->/g, '');
  const seenIds = {};

  // ① Strategy 1: <tr>...</tr> 테이블 행 파싱 (공지사항, 공시자료, 인재채용, 이용상담문의)
  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let trMatch;
  while ((trMatch = trRegex.exec(cleanHtml)) !== null) {
    const tr = trMatch[1];
    const wrIdMatch = tr.match(/wr_id=(\d+)/i);
    if (!wrIdMatch) continue;
    
    const wrId = wrIdMatch[1];
    if (seenIds[wrId]) continue;

    let title = "";
    const boTitMatch = tr.match(/class=["'][^"']*bo_tit[^"']*["'][^>]*>([\s\S]*?)<\/a>/i);
    if (boTitMatch) {
      title = boTitMatch[1].replace(/<[^>]+>/g, '').trim();
    }
    if (!title) {
      const titleAttrMatch = tr.match(/title=["']([^"']+)["']/i);
      if (titleAttrMatch) title = titleAttrMatch[1].trim();
    }
    if (!title) {
      const linkMatch = tr.match(new RegExp(`href=["'][^"']*wr_id=${wrId}[^"']*["'][^>]*>([\\s\\S]*?)<\\/a>`, 'i'));
      if (linkMatch) title = linkMatch[1].replace(/<[^>]+>/g, '').trim();
    }

    title = title.replace(/\s*N\s*새글\s*$/i, '').replace(/\s*새글\s*$/i, '').trim();
    if (!title || title.length < 2 || title.indexOf('답변') !== -1 || title.indexOf('제목') !== -1) continue;

    let dateStr = "";
    const tdDateMatch = tr.match(/class=["'][^"']*td_datetime[^"']*["'][^>]*>([\s\S]*?)<\/td>/i);
    if (tdDateMatch) {
      const textDate = tdDateMatch[1].replace(/<[^>]+>/g, '').trim();
      const dm = textDate.match(/\b(20\d{2}[-.\/]\d{2}[-.\/]\d{2})\b/);
      if (dm) dateStr = dm[1].replace(/[\/.]/g, '-');
    }
    if (!dateStr) {
      const textOnly = tr.replace(/<[^>]+>/g, ' ');
      const dm = textOnly.match(/\b(20\d{2}[-.\/]\d{2}[-.\/]\d{2})\b/);
      if (dm) dateStr = dm[1].replace(/[\/.]/g, '-');
    }

    if (dateStr) {
      seenIds[wrId] = true;
      posts.push({ wrId: wrId, title: title, date: dateStr });
    }
  }

  // ② Strategy 2: <li class="gall_li">...</li> 카드 파싱 (갤러리, 이용인모집, 정보안내 카드 스킨)
  const liRegex = /<li[^>]*class=["'][^"']*gall_li[^"']*["'][^>]*>([\s\S]*?)<\/li>/gi;
  let liMatch;
  while ((liMatch = liRegex.exec(cleanHtml)) !== null) {
    const li = liMatch[1];
    const wrIdMatch = li.match(/wr_id=(\d+)/i);
    if (!wrIdMatch) continue;

    const wrId = wrIdMatch[1];
    if (seenIds[wrId]) continue;

    let title = "";
    const boTitMatch = li.match(/class=["'][^"']*bo_tit[^"']*["'][^>]*>([\s\S]*?)<\/a>/i);
    if (boTitMatch) {
      title = boTitMatch[1].replace(/<[^>]+>/g, '').trim();
    }
    if (!title) {
      const titleAttrMatch = li.match(/title=["']([^"']+)["']/i);
      if (titleAttrMatch) title = titleAttrMatch[1].trim();
    }
    if (!title) {
      const linkMatch = li.match(new RegExp(`href=["'][^"']*wr_id=${wrId}[^"']*["'][^>]*>([\\s\\S]*?)<\\/a>`, 'i'));
      if (linkMatch) title = linkMatch[1].replace(/<[^>]+>/g, '').trim();
    }

    title = title.replace(/\s*N\s*새글\s*$/i, '').replace(/\s*새글\s*$/i, '').trim();
    if (!title || title.length < 2 || title.indexOf('답변') !== -1) continue;

    const textOnly = li.replace(/<[^>]+>/g, ' ');
    const dm = textOnly.match(/\b(20\d{2}[-.\/]\d{2}[-.\/]\d{2})\b/);
    const dateStr = dm ? dm[1].replace(/[\/.]/g, '-') : "";

    if (dateStr) {
      seenIds[wrId] = true;
      posts.push({ wrId: wrId, title: title, date: dateStr });
    }
  }

  return posts;
}

/**
 * 5. 일별 집계 시트 업데이트 함수 (인덱스 기반)
 */
function updateDailyAggregation(dailySheet, dateStr, confIndex) {
  if (confIndex < 0 || confIndex > 6) return;
  const colIndex = confIndex + 2; // Col B (2) to Col H (8)

  const lastRow = dailySheet.getLastRow();
  let targetRowIndex = -1;

  if (lastRow > 1) {
    const dates = dailySheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < dates.length; i++) {
      let dStr = "";
      if (dates[i][0] instanceof Date) {
        dStr = Utilities.formatDate(dates[i][0], "Asia/Seoul", "yyyy-MM-dd");
      } else {
        dStr = String(dates[i][0]).trim().substring(0, 10);
      }
      if (dStr === dateStr) {
        targetRowIndex = i + 2;
        break;
      }
    }
  }

  if (targetRowIndex !== -1) {
    const currentVal = Number(dailySheet.getRange(targetRowIndex, colIndex).getValue()) || 0;
    dailySheet.getRange(targetRowIndex, colIndex).setValue(currentVal + 1);
    
    // Recalculate total sum in col I (9)
    const rowVals = dailySheet.getRange(targetRowIndex, 2, 1, 7).getValues()[0];
    const rowTotal = rowVals.reduce((acc, v) => acc + (Number(v) || 0), 0);
    dailySheet.getRange(targetRowIndex, 9).setValue(rowTotal);
  } else {
    const newRow = [dateStr, 0, 0, 0, 0, 0, 0, 0, 0];
    newRow[colIndex - 1] = 1;
    newRow[8] = 1; // Total
    dailySheet.appendRow(newRow);
  }
}

/**
 * 6. Web App JSON API Endpoint (doGet)
 * ?type=reset : 시트 초기화 후 스크래핑 재실행
 * ?type=scrape : 즉시 스크래퍼 실행
 */
function doGet(e) {
  if (e && e.parameter) {
    if (e.parameter.type === 'reset') {
      resetAndScrape();
    } else if (e.parameter.type === 'scrape') {
      runScraper();
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getSheetByName('일별집계');
  const historySheet = ss.getSheetByName('수집이력');

  const dailyData = [];
  if (dailySheet && dailySheet.getLastRow() > 1) {
    const rows = dailySheet.getRange(2, 1, dailySheet.getLastRow() - 1, 9).getValues();
    rows.forEach(r => {
      const dateStr = r[0] instanceof Date ? 
        Utilities.formatDate(r[0], "Asia/Seoul", "yyyy-MM-dd") : String(r[0]).trim().substring(0, 10);
      
      const cats = {};
      CATEGORIES.forEach((c, idx) => {
        cats[c] = Number(r[idx + 1]) || 0;
      });

      dailyData.push({
        date: dateStr,
        categories: cats,
        total: Number(r[8]) || 0
      });
    });
  }

  const historyData = [];
  if (historySheet && historySheet.getLastRow() > 1) {
    const totalRows = historySheet.getLastRow() - 1;
    const readRows = Math.min(totalRows, 500);
    // Read most recent rows first
    const startRow = Math.max(2, historySheet.getLastRow() - readRows + 1);
    const rows = historySheet.getRange(startRow, 1, readRows, 7).getValues();
    
    rows.reverse().forEach(r => {
      historyData.push({
        post_id: String(r[0]),
        main_category: String(r[1]),
        sub_category: String(r[2]),
        category: String(r[3]),
        title: String(r[4]),
        date: r[5] instanceof Date ? Utilities.formatDate(r[5], "Asia/Seoul", "yyyy-MM-dd") : String(r[5]).substring(0, 10),
        collected_at: r[6] instanceof Date ? Utilities.formatDate(r[6], "Asia/Seoul", "yyyy-MM-dd HH:mm:ss") : String(r[6])
      });
    });
  }

  const result = {
    status: 'success',
    timestamp: Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"),
    daily: dailyData,
    history: historyData
  };

  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * 7. 시간 기반 자동 트리거 생성 함수
 * (최초 1회 실행하면 15분마다 runScraper 자동 실행)
 */
function createTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'runScraper') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runScraper')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('15분 주기 자동 스크래핑 트리거 설정 완료!');
}

/**
 * 8. 데이터 완전 초기화 후 재스크래핑 (클린업용)
 */
function resetAndScrape() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName('수집이력');
  let dailySheet = ss.getSheetByName('일별집계');
  
  if (historySheet) historySheet.clear();
  if (dailySheet) dailySheet.clear();
  
  initSheets();
  runScraper();
  Logger.log('데이터 초기화 및 최신 스크래핑 재실행 완료!');
}
