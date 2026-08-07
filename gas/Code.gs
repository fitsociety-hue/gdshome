/**
 * ==========================================================================
 * 강동어울림복지관 홍보 통합실적 관리 앱 - Google Apps Script (GAS) 백엔드 Code.gs
 * ==========================================================================
 * - 역할: 
 *   1. 복지관 7개 게시판 정기 스크래핑 (gde.or.kr - GNUBoard 스킨 대응)
 *   2. 유튜브 채널 신규 업로드 동영상 감지 (YouTube Channel RSS / Atom XML)
 *   3. 네이버 블로그 카테고리별 신규 포스팅 감지 (Naver Blog RSS XML)
 *   4. post_id 기반 중복 제거 및 Google Sheets 적재
 *   5. KST 기준 일별/채널별/카테고리별 실적 인덱스 기반 인코딩 안전 카운팅
 *   6. 프론트엔드(GitHub Pages) 연결용 Web App JSON API (doGet)
 * ==========================================================================
 */

// 1. 대상 채널 및 카테고리 스펙 정의 (총 15개 항목)
const TARGET_CONFIGS = [
  // [홈페이지 게시판 7개]
  { index: 0, channel: '홈페이지', category: '공지사항', mainCategory: '어울림 소식', subCategory: '공지사항', url: 'https://gde.or.kr/notice', boTable: 'notice', type: 'gnuboard' },
  { index: 1, channel: '홈페이지', category: '공시자료', mainCategory: '어울림 소식', subCategory: '공지사항', url: 'https://gde.or.kr/infoopen', boTable: 'infoopen', type: 'gnuboard' },
  { index: 2, channel: '홈페이지', category: '인재채용', mainCategory: '어울림 소식', subCategory: '공지사항', url: 'https://gde.or.kr/recruitment', boTable: 'recruitment', type: 'gnuboard' },
  { index: 3, channel: '홈페이지', category: '이용인모집', mainCategory: '어울림 소식', subCategory: '이용인 모집', url: 'https://gde.or.kr/program', boTable: 'program', type: 'gnuboard' },
  { index: 4, channel: '홈페이지', category: '정보안내', mainCategory: '어울림 소식', subCategory: '정보안내', url: 'https://gde.or.kr/information', boTable: 'information', type: 'gnuboard' },
  { index: 5, channel: '홈페이지', category: '갤러리(전체)', mainCategory: '어울림 갤러리', subCategory: '전체', url: 'https://gde.or.kr/gallery', boTable: 'gallery', type: 'gnuboard' },
  { index: 6, channel: '홈페이지', category: '이용상담문의', mainCategory: '참여하기', subCategory: '이용상담문의', url: 'https://gde.or.kr/counseling', boTable: 'counseling', type: 'gnuboard' },

  // [유튜브 1개]
  { index: 7, channel: '유튜브', category: '유튜브(동영상)', mainCategory: '영상 미디어', subCategory: '신규 동영상', url: 'https://www.youtube.com/@%EA%B0%95%EB%8F%99%EC%96%B4%EC%9A%B8%EB%A6%BC%EB%B3%B5%EC%A7%80%EA%B4%80', boTable: 'youtube_video', type: 'youtube', channelId: 'UCgAO5d2OyVURs3e2F6tphVw' },

  // [네이버 블로그 7개 카테고리]
  { index: 8, channel: '네이버 블로그', category: '블로그-전체', mainCategory: '네이버 블로그', subCategory: '전체보기', url: 'https://blog.naver.com/gds0741', boTable: 'blog_all', type: 'blog' },
  { index: 9, channel: '네이버 블로그', category: '블로그-공지사항', mainCategory: '네이버 블로그', subCategory: '공지사항', url: 'https://blog.naver.com/gds0741', boTable: 'blog_notice', type: 'blog' },
  { index: 10, channel: '네이버 블로그', category: '블로그-복지관소식', mainCategory: '네이버 블로그', subCategory: '복지관소식', url: 'https://blog.naver.com/gds0741', boTable: 'blog_news', type: 'blog' },
  { index: 11, channel: '네이버 블로그', category: '블로그-복지정보', mainCategory: '네이버 블로그', subCategory: '복지정보', url: 'https://blog.naver.com/gds0741', boTable: 'blog_info', type: 'blog' },
  { index: 12, channel: '네이버 블로그', category: '블로그-도서추천', mainCategory: '네이버 블로그', subCategory: '도서 추천', url: 'https://blog.naver.com/gds0741', boTable: 'blog_book', type: 'blog' },
  { index: 13, channel: '네이버 블로그', category: '블로그-프로그램모집', mainCategory: '네이버 블로그', subCategory: '프로그램모집', url: 'https://blog.naver.com/gds0741', boTable: 'blog_program', type: 'blog' },
  { index: 14, channel: '네이버 블로그', category: '블로그-배너', mainCategory: '네이버 블로그', subCategory: '배너', url: 'https://blog.naver.com/gds0741', boTable: 'blog_banner', type: 'blog' }
];

const CATEGORIES = TARGET_CONFIGS.map(c => c.category);

/**
 * 2. 초기 구글 스프레드시트 구조 세팅
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 수집이력 시트
  let historySheet = ss.getSheetByName('수집이력');
  if (!historySheet) {
    historySheet = ss.insertSheet('수집이력');
  }
  if (historySheet.getLastRow() === 0) {
    historySheet.appendRow(['post_id', '채널', '대분류', '중분류', '세부항목', '제목', '게시일(KST)', '수집시각', '원문URL']);
    historySheet.getRange('A1:I1').setBackground('#E2E8F0').setFontWeight('bold');
  }

  // ② 일별집계 시트
  let dailySheet = ss.getSheetByName('일별집계');
  if (!dailySheet) {
    dailySheet = ss.insertSheet('일별집계');
  }
  if (dailySheet.getLastRow() === 0) {
    const headers = ['날짜', ...CATEGORIES, '합계'];
    dailySheet.appendRow(headers);
    dailySheet.getRange(1, 1, 1, headers.length).setBackground('#CBD5E1').setFontWeight('bold');
  }

  // ③ config 시트
  let configSheet = ss.getSheetByName('config');
  if (!configSheet) {
    configSheet = ss.insertSheet('config');
  }
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(['세부항목', '채널', '대분류', 'URL', 'boTable', '수집방식']);
    TARGET_CONFIGS.forEach(conf => {
      configSheet.appendRow([conf.category, conf.channel, conf.mainCategory, conf.url, conf.boTable, conf.type]);
    });
    configSheet.getRange('A1:F1').setBackground('#E2E8F0').setFontWeight('bold');
  }

  Logger.log('시트 초기화 세팅 완료! (15개 카테고리 지원)');
}

/**
 * 3. 메인 스크래핑 및 카운팅 엔진
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

  // ① 7개 홈페이지 게시판 스크래핑
  const websiteConfigs = TARGET_CONFIGS.filter(c => c.type === 'gnuboard');
  websiteConfigs.forEach(conf => {
    try {
      const response = UrlFetchApp.fetch(conf.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
      });

      if (response.getResponseCode() === 200) {
        const html = response.getContentText('UTF-8');
        const posts = parseGnuboardHtml(html, conf);

        posts.forEach(post => {
          const postId = `${conf.boTable}_${post.wrId}`;
          if (!existingPostIds.has(postId)) {
            existingPostIds.add(postId);
            newCollectedCount++;

            historySheet.appendRow([
              postId,
              conf.channel,
              conf.mainCategory,
              conf.subCategory,
              conf.category,
              post.title,
              post.date,
              kstNow,
              post.url || conf.url
            ]);

            updateDailyAggregation(dailySheet, post.date, conf.index);
          }
        });
      }
    } catch (e) {
      Logger.log(`홈페이지 스크래핑 에러 [${conf.category}]: ${e.toString()}`);
    }
  });

  // ② 유튜브 채널 신규 업로드 동영상 수집
  const ytConfig = TARGET_CONFIGS.find(c => c.type === 'youtube');
  if (ytConfig) {
    try {
      const ytRssUrl = `https://www.youtube.com/feeds/videos.xml?channel_id=${ytConfig.channelId}`;
      const response = UrlFetchApp.fetch(ytRssUrl, { muteHttpExceptions: true });

      if (response.getResponseCode() === 200) {
        const xml = response.getContentText('UTF-8');
        const ytPosts = parseYouTubeRss(xml, ytConfig);

        ytPosts.forEach(post => {
          const postId = `yt_${post.videoId}`;
          if (!existingPostIds.has(postId)) {
            existingPostIds.add(postId);
            newCollectedCount++;

            historySheet.appendRow([
              postId,
              ytConfig.channel,
              ytConfig.mainCategory,
              ytConfig.subCategory,
              ytConfig.category,
              post.title,
              post.date,
              kstNow,
              post.url
            ]);

            updateDailyAggregation(dailySheet, post.date, ytConfig.index);
          }
        });
      }
    } catch (e) {
      Logger.log(`유튜브 스크래핑 에러: ${e.toString()}`);
    }
  }

  // ③ 네이버 블로그 카테고리별 신규 포스팅 수집
  try {
    const blogRssUrl = 'https://rss.blog.naver.com/gds0741.xml';
    const response = UrlFetchApp.fetch(blogRssUrl, { muteHttpExceptions: true });

    if (response.getResponseCode() === 200) {
      const xml = response.getContentText('UTF-8');
      const blogPosts = parseNaverBlogRss(xml);

      blogPosts.forEach(post => {
        // 블로그 포스팅은 전체보기 카테고리와 해당 세부 카테고리에 동시 적재
        const matchedConfigs = getBlogMatchedConfigs(post.rawCategory);

        matchedConfigs.forEach(conf => {
          const postId = `blog_${post.logNo}_${conf.boTable}`;
          if (!existingPostIds.has(postId)) {
            existingPostIds.add(postId);
            newCollectedCount++;

            historySheet.appendRow([
              postId,
              conf.channel,
              conf.mainCategory,
              conf.subCategory,
              conf.category,
              post.title,
              post.date,
              kstNow,
              post.url
            ]);

            updateDailyAggregation(dailySheet, post.date, conf.index);
          }
        });
      });
    }
  } catch (e) {
    Logger.log(`네이버 블로그 스크래핑 에러: ${e.toString()}`);
  }

  Logger.log(`스크래핑 작업 완료! 신규 수집건수: ${newCollectedCount}건`);
}

/**
 * 4. GNUBoard HTML 파싱 함수
 */
function parseGnuboardHtml(html, conf) {
  const posts = [];
  const cleanHtml = html.replace(/<!--[\s\S]*?-->/g, '');
  const seenIds = {};

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
    if (boTitMatch) title = boTitMatch[1].replace(/<[^>]+>/g, '').trim();
    if (!title) {
      const titleAttrMatch = tr.match(/title=["']([^"']+)["']/i);
      if (titleAttrMatch) title = titleAttrMatch[1].trim();
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
      posts.push({ wrId: wrId, title: title, date: dateStr, url: `${conf.url}?wr_id=${wrId}` });
    }
  }

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
    if (boTitMatch) title = boTitMatch[1].replace(/<[^>]+>/g, '').trim();
    if (!title) {
      const titleAttrMatch = li.match(/title=["']([^"']+)["']/i);
      if (titleAttrMatch) title = titleAttrMatch[1].trim();
    }

    title = title.replace(/\s*N\s*새글\s*$/i, '').replace(/\s*새글\s*$/i, '').trim();
    if (!title || title.length < 2) continue;

    const textOnly = li.replace(/<[^>]+>/g, ' ');
    const dm = textOnly.match(/\b(20\d{2}[-.\/]\d{2}[-.\/]\d{2})\b/);
    const dateStr = dm ? dm[1].replace(/[\/.]/g, '-') : "";

    if (dateStr) {
      seenIds[wrId] = true;
      posts.push({ wrId: wrId, title: title, date: dateStr, url: `${conf.url}?wr_id=${wrId}` });
    }
  }

  return posts;
}

/**
 * 5. 유튜브 Atom/RSS XML 파싱 함수
 */
function parseYouTubeRss(xml, conf) {
  const posts = [];
  const entryRegex = /<entry>([\s\S]*?)<\/entry>/gi;
  let match;

  while ((match = entryRegex.exec(xml)) !== null) {
    const entry = match[1];
    const videoIdMatch = entry.match(/<yt:videoId>([^<]+)<\/yt:videoId>/i);
    const titleMatch = entry.match(/<title>([^<]+)<\/title>/i);
    const pubMatch = entry.match(/<published>([^<]+)<\/published>/i);
    const linkMatch = entry.match(/<link[^>]*href=["']([^"']+)["']/i);

    if (videoIdMatch && titleMatch && pubMatch) {
      const videoId = videoIdMatch[1].trim();
      const title = titleMatch[1].replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
      const rawPub = pubMatch[1].trim();
      
      // Parse ISO date string to KST date 'yyyy-MM-dd'
      const pubDate = new Date(rawPub);
      const dateStr = Utilities.formatDate(pubDate, "Asia/Seoul", "yyyy-MM-dd");
      const videoUrl = linkMatch ? linkMatch[1] : `https://www.youtube.com/watch?v=${videoId}`;

      posts.push({
        videoId: videoId,
        title: title,
        date: dateStr,
        url: videoUrl
      });
    }
  }

  return posts;
}

/**
 * 6. 네이버 블로그 RSS XML 파싱 함수
 */
function parseNaverBlogRss(xml) {
  const posts = [];
  const itemRegex = /<item>([\s\S]*?)<\/item>/gi;
  let match;

  while ((match = itemRegex.exec(xml)) !== null) {
    const item = match[1];
    const titleMatch = item.match(/<title>([\s\S]*?)<\/title>/i);
    const linkMatch = item.match(/<link>([\s\S]*?)<\/link>/i);
    const catMatch = item.match(/<category>([\s\S]*?)<\/category>/i);
    const pubDateMatch = item.match(/<pubDate>([\s\S]*?)<\/pubDate>/i);

    if (titleMatch && linkMatch && pubDateMatch) {
      const rawTitle = titleMatch[1].replace(/<!\[CDATA\[([\s\S]*?)\]\]>/gi, '$1').replace(/<[^>]+>/g, '').trim();
      const link = linkMatch[1].trim();
      const rawCategory = catMatch ? catMatch[1].replace(/<!\[CDATA\[([\s\S]*?)\]\]>/gi, '$1').replace(/<[^>]+>/g, '').trim() : '소식';
      
      // Extract logNo from link (e.g. https://blog.naver.com/gds0741/223537000000)
      const logNoMatch = link.match(/\/(\d+)\s*$/) || link.match(/logNo=(\d+)/i);
      const logNo = logNoMatch ? logNoMatch[1] : String(Math.abs(hashCode(link)));

      // Parse pubDate (RFC 822 format: e.g. "Wed, 05 Aug 2026 13:36:01 +0900")
      const pubDateText = pubDateMatch[1].trim();
      let dateStr = "";
      try {
        const d = new Date(pubDateText);
        dateStr = Utilities.formatDate(d, "Asia/Seoul", "yyyy-MM-dd");
      } catch (e) {
        dateStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
      }

      posts.push({
        logNo: logNo,
        title: rawTitle,
        rawCategory: rawCategory,
        date: dateStr,
        url: link
      });
    }
  }

  return posts;
}

/**
 * 네이버 블로그 카테고리 매핑 규칙 (전체보기 1개 + 세부 1개)
 */
function getBlogMatchedConfigs(rawCat) {
  const blogConfigs = TARGET_CONFIGS.filter(c => c.type === 'blog');
  const matched = [];

  // ① 항상 '블로그-전체' 매핑 포함
  const allConf = blogConfigs.find(c => c.category === '블로그-전체');
  if (allConf) matched.push(allConf);

  // ② 세부 카테고리 매핑
  const catClean = rawCat.replace(/\s+/g, '');

  if (catClean.indexOf('공지') !== -1) {
    const c = blogConfigs.find(conf => conf.category === '블로그-공지사항');
    if (c) matched.push(c);
  } else if (catClean.indexOf('복지관소식') !== -1 || catClean.indexOf('소식') !== -1) {
    const c = blogConfigs.find(conf => conf.category === '블로그-복지관소식');
    if (c) matched.push(c);
  } else if (catClean.indexOf('복지정보') !== -1 || catClean.indexOf('정보') !== -1) {
    const c = blogConfigs.find(conf => conf.category === '블로그-복지정보');
    if (c) matched.push(c);
  } else if (catClean.indexOf('도서') !== -1 || catClean.indexOf('추천') !== -1) {
    const c = blogConfigs.find(conf => conf.category === '블로그-도서추천');
    if (c) matched.push(c);
  } else if (catClean.indexOf('프로그램') !== -1 || catClean.indexOf('모집') !== -1) {
    const c = blogConfigs.find(conf => conf.category === '블로그-프로그램모집');
    if (c) matched.push(c);
  } else if (catClean.indexOf('배너') !== -1) {
    const c = blogConfigs.find(conf => conf.category === '블로그-배너');
    if (c) matched.push(c);
  } else {
    // 디폴트: 복지관소식
    const c = blogConfigs.find(conf => conf.category === '블로그-복지관소식');
    if (c && !matched.includes(c)) matched.push(c);
  }

  return matched;
}

function hashCode(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = (hash << 5) - hash + str.charCodeAt(i);
    hash |= 0;
  }
  return hash;
}

/**
 * 7. 일별 집계 시트 업데이트 함수 (인덱스 기반 15개 항목)
 */
function updateDailyAggregation(dailySheet, dateStr, confIndex) {
  if (confIndex < 0 || confIndex >= TARGET_CONFIGS.length) return;
  const colIndex = confIndex + 2; // Col B (2) to Col P (16)

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
    
    // Recalculate total sum in last col (17)
    const rowVals = dailySheet.getRange(targetRowIndex, 2, 1, TARGET_CONFIGS.length).getValues()[0];
    const rowTotal = rowVals.reduce((acc, v) => acc + (Number(v) || 0), 0);
    dailySheet.getRange(targetRowIndex, TARGET_CONFIGS.length + 2).setValue(rowTotal);
  } else {
    const newRow = [dateStr, ...new Array(TARGET_CONFIGS.length).fill(0), 0];
    newRow[colIndex - 1] = 1;
    newRow[TARGET_CONFIGS.length + 1] = 1; // Total
    dailySheet.appendRow(newRow);
  }
}

/**
 * 8. Web App JSON API Endpoint (doGet)
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
    const colCount = TARGET_CONFIGS.length + 2;
    const rows = dailySheet.getRange(2, 1, dailySheet.getLastRow() - 1, colCount).getValues();
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
        total: Number(r[colCount - 1]) || 0
      });
    });
  }

  const historyData = [];
  if (historySheet && historySheet.getLastRow() > 1) {
    const totalRows = historySheet.getLastRow() - 1;
    const readRows = Math.min(totalRows, 600);
    const startRow = Math.max(2, historySheet.getLastRow() - readRows + 1);
    const rows = historySheet.getRange(startRow, 1, readRows, 9).getValues();
    
    rows.reverse().forEach(r => {
      historyData.push({
        post_id: String(r[0]),
        channel: String(r[1] || '홈페이지'),
        main_category: String(r[2]),
        sub_category: String(r[3]),
        category: String(r[4]),
        title: String(r[5]),
        date: r[6] instanceof Date ? Utilities.formatDate(r[6], "Asia/Seoul", "yyyy-MM-dd") : String(r[6]).substring(0, 10),
        collected_at: r[7] instanceof Date ? Utilities.formatDate(r[7], "Asia/Seoul", "yyyy-MM-dd HH:mm:ss") : String(r[7]),
        url: String(r[8] || '')
      });
    });
  }

  const result = {
    status: 'success',
    timestamp: Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss"),
    categories: CATEGORIES,
    configs: TARGET_CONFIGS,
    daily: dailyData,
    history: historyData
  };

  const output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * 9. 시간 기반 자동 트리거 생성 함수
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
 * 10. 데이터 완전 초기화 후 재스크래핑
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
