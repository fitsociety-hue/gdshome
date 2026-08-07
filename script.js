/**
 * 강동어울림복지관 홍보 통합실적 관리 시스템 - 프론트엔드 로직 (script.js v2.0)
 * 홈페이지(7) + 네이버 블로그(7) + 유튜브(1) Multi-Channel 통합 대시보드
 */

// 1. App Configuration & Constants
const DEFAULT_GAS_URL = 'https://script.google.com/macros/s/AKfycbypHZAg1bQRguKDvXGJpkRdfkTY7Aqy9K1LdL5TwLuzkBjcA8QfqPrvjzc9JqxrCuJI/exec';

const CATEGORIES = [
  // 홈페이지 (7)
  '공지사항',
  '공시자료',
  '인재채용',
  '이용인모집',
  '정보안내',
  '갤러리(전체)',
  '이용상담문의',
  // 유튜브 (1)
  '유튜브(동영상)',
  // 네이버 블로그 (7)
  '블로그-전체',
  '블로그-공지사항',
  '블로그-복지관소식',
  '블로그-복지정보',
  '블로그-도서추천',
  '블로그-프로그램모집',
  '블로그-배너'
];

const CATEGORY_CHANNELS = {
  '공지사항': '홈페이지',
  '공시자료': '홈페이지',
  '인재채용': '홈페이지',
  '이용인모집': '홈페이지',
  '정보안내': '홈페이지',
  '갤러리(전체)': '홈페이지',
  '이용상담문의': '홈페이지',
  '유튜브(동영상)': '유튜브',
  '블로그-전체': '네이버 블로그',
  '블로그-공지사항': '네이버 블로그',
  '블로그-복지관소식': '네이버 블로그',
  '블로그-복지정보': '네이버 블로그',
  '블로그-도서추천': '네이버 블로그',
  '블로그-프로그램모집': '네이버 블로그',
  '블로그-배너': '네이버 블로그'
};

const CATEGORY_COLORS = {
  // 홈페이지 (Blue/Cyan/Indigo/Purple/Emerald)
  '공지사항': '#3B82F6',
  '공시자료': '#2563EB',
  '인재채용': '#1D4ED8',
  '이용인모집': '#06B6D4',
  '정보안내': '#6366F1',
  '갤러리(전체)': '#8B5CF6',
  '이용상담문의': '#10B981',
  // 유튜브 (Red)
  '유튜브(동영상)': '#EF4444',
  // 네이버 블로그 (Green Variants)
  '블로그-전체': '#03C75A',
  '블로그-공지사항': '#059669',
  '블로그-복지관소식': '#10B981',
  '블로그-복지정보': '#34D399',
  '블로그-도서추천': '#0284C7',
  '블로그-프로그램모집': '#14B8A6',
  '블로그-배너': '#84CC16'
};

const CATEGORY_URLS = {
  '공지사항': 'https://gde.or.kr/notice',
  '공시자료': 'https://gde.or.kr/infoopen',
  '인재채용': 'https://gde.or.kr/recruitment',
  '이용인모집': 'https://gde.or.kr/program',
  '정보안내': 'https://gde.or.kr/information',
  '갤러리(전체)': 'https://gde.or.kr/gallery',
  '이용상담문의': 'https://gde.or.kr/counseling',
  '유튜브(동영상)': 'https://www.youtube.com/@%EA%B0%95%EB%8F%99%EC%96%B4%EC%9A%B8%EB%A6%BC%EB%B3%B5%EC%A7%80%EA%B4%80',
  '블로그-전체': 'https://blog.naver.com/gds0741',
  '블로그-공지사항': 'https://blog.naver.com/gds0741',
  '블로그-복지관소식': 'https://blog.naver.com/gds0741',
  '블로그-복지정보': 'https://blog.naver.com/gds0741',
  '블로그-도서추천': 'https://blog.naver.com/gds0741',
  '블로그-프로그램모집': 'https://blog.naver.com/gds0741',
  '블로그-배너': 'https://blog.naver.com/gds0741'
};

// 2. Application State
const state = {
  gasUrl: localStorage.getItem('gds_gas_url') || DEFAULT_GAS_URL,
  isDemoMode: localStorage.getItem('gds_demo_mode') === 'true',
  selectedChannelTab: 'ALL', // 'ALL', '홈페이지', '네이버 블로그', '유튜브'
  period: 'today', // 'today', 'week', 'month', 'all', 'custom'
  startDate: '',
  endDate: '',
  rawDailyData: [],
  rawHistoryData: [],
  chartType: 'line', // 'line' or 'bar'
  trendChartInstance: null,
  doughnutChartInstance: null,
  currentPage: 1,
  pageSize: 10,
  searchQuery: '',
  selectedTableChannelFilter: 'ALL',
  selectedCategoryFilter: 'ALL'
};

// 3. Helper Functions
function getTodayKST() {
  const now = new Date();
  const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
  const kst = new Date(utc + (9 * 3600000));
  return kst.toISOString().split('T')[0];
}

function formatDate(dateObj) {
  const year = dateObj.getFullYear();
  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
  const day = String(dateObj.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function updateKSTClock() {
  const clockEl = document.getElementById('kstClock');
  if (!clockEl) return;
  const now = new Date();
  const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
  const kst = new Date(utc + (9 * 3600000));
  const timeStr = kst.toLocaleTimeString('ko-KR', { hour12: false });
  const dateStr = `${kst.getFullYear()}.${String(kst.getMonth() + 1).padStart(2, '0')}.${String(kst.getDate()).padStart(2, '0')}`;
  clockEl.textContent = `${dateStr} ${timeStr} KST`;
}

// 4. Sample Mock Data Generator (Demo Mode)
function generateMockData() {
  const todayStr = getTodayKST();
  const today = new Date(todayStr);
  const daily = [];
  const history = [];

  const blogSampleTitles = {
    '블로그-공지사항': '★장애가족운동회 \'같이 어울린데이\' 참가자 모집 안내 ★',
    '블로그-복지관소식': '[맞춤지원팀] 자조모임 \'라온\' 7월 활동 소식',
    '블로그-복지정보': '알기 쉬운 2026년 장애인 복지 혜택 종합 안내',
    '블로그-도서추천': '어울림 서재 추천 도서: <다함께 따뜻한 세상 만들기>',
    '블로그-프로그램모집': '2026년 하반기 정보화 교실 신규 수강생 모집',
    '블로그-배너': '강동어울림복지관 홈페이지 개편 안내 배너 포스팅'
  };

  const ytSampleTitles = [
    '2026년 서울시장애인돌봄가족휴가제 2차 대기자 추첨 (10가족)',
    '2026년 서울시장애인돌봄가족휴가제 실시간 추첨 (31가족)',
    '강동어울림복지관 미래경영팀 2026년 온라인 사업설명회',
    '강동어울림복지관 성장지원팀 2026년 온라인 사업설명회',
    '강동어울림복지관 건강문화팀 2026년 온라인 사업설명회',
    '강동어울림복지관 전략기획팀 2026년 온라인 사업설명회'
  ];

  // Generate last 30 days of data
  for (let i = 29; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dateStr = formatDate(d);
    
    const cats = {};
    let totalDay = 0;
    const isWeekend = (d.getDay() === 0 || d.getDay() === 6);
    const multiplier = isWeekend ? 0.3 : 1.0;

    CATEGORIES.forEach((cat, idx) => {
      let count = 0;

      if (cat === '유튜브(동영상)') {
        count = (!isWeekend && (i % 4 === 0)) ? 1 : 0;
      } else if (cat.startsWith('블로그-')) {
        if (cat === '블로그-전체') {
          // Handled separately after detail counts
          count = 0;
        } else {
          const base = (idx % 2 === 0) ? 1 : 0;
          count = Math.floor((Math.random() * 2 + base) * multiplier);
        }
      } else {
        // 홈페이지
        const base = (idx % 3 === 0) ? 2 : 1;
        count = Math.floor((Math.random() * 2 + base) * multiplier);
      }

      cats[cat] = count;
      totalDay += count;
    });

    // Calculate 블로그-전체 sum
    let blogSum = 0;
    ['블로그-공지사항', '블로그-복지관소식', '블로그-복지정보', '블로그-도서추천', '블로그-프로그램모집', '블로그-배너'].forEach(bc => {
      blogSum += cats[bc] || 0;
    });
    if (blogSum === 0 && !isWeekend) blogSum = 1;
    cats['블로그-전체'] = blogSum;

    // Generate history records
    CATEGORIES.forEach(cat => {
      const cnt = cats[cat];
      if (cat === '블로그-전체') return; // Don't duplicate history for blog-all

      for (let k = 0; k < cnt; k++) {
        const ch = CATEGORY_CHANNELS[cat];
        let title = `[${cat}] 2026년 ${d.getMonth() + 1}월 ${cat} 안내 #${k + 1}`;
        let url = CATEGORY_URLS[cat];
        let postId = `${cat.substring(0, 4)}_${dateStr.replace(/-/g, '')}_${k + 1}`;

        if (ch === '네이버 블로그') {
          title = blogSampleTitles[cat] || `[네이버 블로그] ${cat} 포스팅 #${k + 1}`;
          url = `https://blog.naver.com/gds0741/2235370${String(i * 10 + k).padStart(4, '0')}`;
          postId = `blog_2235370${String(i * 10 + k).padStart(4, '0')}_${cat}`;
        } else if (ch === '유튜브') {
          title = ytSampleTitles[k % ytSampleTitles.length];
          url = `https://www.youtube.com/watch?v=sampleYt${i}${k}`;
          postId = `yt_sampleYt${i}${k}`;
        }

        history.push({
          post_id: postId,
          channel: ch,
          main_category: cat.startsWith('블로그') ? '네이버 블로그' : (ch === '유튜브' ? '영상 미디어' : '어울림 소식'),
          sub_category: cat,
          category: cat,
          title: title,
          date: dateStr,
          collected_at: `${dateStr} ${String(9 + (k * 2)).padStart(2, '0')}:15:00`,
          url: url
        });

        // Also add a blog-all history record for blog posts
        if (ch === '네이버 블로그') {
          history.push({
            post_id: `blog_all_${postId}`,
            channel: '네이버 블로그',
            main_category: '네이버 블로그',
            sub_category: '전체보기',
            category: '블로그-전체',
            title: title,
            date: dateStr,
            collected_at: `${dateStr} ${String(9 + (k * 2)).padStart(2, '0')}:15:00`,
            url: url
          });
        }
      }
    });

    daily.push({
      date: dateStr,
      categories: cats,
      total: totalDay
    });
  }

  history.sort((a, b) => new Date(b.date + ' ' + b.collected_at) - new Date(a.date + ' ' + a.collected_at));

  return { daily, history };
}

// 5. Data Fetching Logic
async function loadData(forceScrape = false) {
  setSyncStatus('syncing', forceScrape ? '실시간 스크래핑 및 동기화 중...' : '데이터 수집 중...');
  const refreshIcon = document.getElementById('refreshIcon');
  if (refreshIcon) refreshIcon.classList.add('fa-spin');

  try {
    if (state.isDemoMode) {
      await new Promise(res => setTimeout(res, 400));
      const mock = generateMockData();
      state.rawDailyData = mock.daily;
      state.rawHistoryData = mock.history;
      setSyncStatus('online', '시연용 샘플 데이터');
    } else {
      const apiUrl = forceScrape ? `${state.gasUrl}?type=scrape` : `${state.gasUrl}?type=all`;
      const response = await fetch(apiUrl, { method: 'GET', mode: 'cors' });

      if (!response.ok) throw new Error(`GAS API HTTP 에러: ${response.status}`);

      const json = await response.json();
      if (json.status === 'success' && json.daily && json.history) {
        state.rawDailyData = json.daily;
        state.rawHistoryData = json.history;
        rebuildDailyFromHistoryIfNeeded();
        setSyncStatus('online', '실시간 동기화 완료');
      } else {
        console.warn('GAS 응답 데이터 미비, 샘플 모드로 표시합니다.', json);
        const mock = generateMockData();
        state.rawDailyData = mock.daily;
        state.rawHistoryData = mock.history;
        setSyncStatus('online', 'GAS 연동 (샘플 데이터 표출)');
      }
    }
  } catch (err) {
    console.error('데이터 로드 오류:', err);
    const mock = generateMockData();
    state.rawDailyData = mock.daily;
    state.rawHistoryData = mock.history;
    setSyncStatus('offline', '연동 오프라인 (샘플 모드 전환)');
  } finally {
    if (refreshIcon) refreshIcon.classList.remove('fa-spin');
    renderDashboard();
  }
}

function rebuildDailyFromHistoryIfNeeded() {
  if (!state.rawHistoryData || state.rawHistoryData.length === 0) return;

  const historyByDate = {};
  state.rawHistoryData.forEach(item => {
    const d = item.date;
    const cat = item.category;
    if (!d || !cat) return;

    if (!historyByDate[d]) {
      historyByDate[d] = { date: d, categories: {}, total: 0 };
      CATEGORIES.forEach(c => historyByDate[d].categories[c] = 0);
    }
    historyByDate[d].categories[cat] = (historyByDate[d].categories[cat] || 0) + 1;
    
    // Don't double count blog-all in total
    if (cat !== '블로그-전체') {
      historyByDate[d].total += 1;
    }
  });

  Object.keys(historyByDate).forEach(d => {
    const existingIndex = state.rawDailyData.findIndex(item => item.date === d);
    if (existingIndex !== -1) {
      if (historyByDate[d].total > (state.rawDailyData[existingIndex].total || 0)) {
        state.rawDailyData[existingIndex] = historyByDate[d];
      }
    } else {
      state.rawDailyData.push(historyByDate[d]);
    }
  });

  state.rawDailyData.sort((a, b) => a.date.localeCompare(b.date));
}

function setSyncStatus(type, message) {
  const statusDot = document.querySelector('.status-dot');
  const statusText = document.getElementById('statusText');
  if (!statusDot || !statusText) return;

  statusDot.className = `status-dot ${type}`;
  statusText.textContent = message;
}

// 6. Filter & Period Logic
function getFilteredDailyData() {
  const todayStr = getTodayKST();
  const today = new Date(todayStr);

  let filtered = [];

  if (state.period === 'today') {
    filtered = state.rawDailyData.filter(d => d.date === todayStr);
  } else if (state.period === 'week') {
    const weekAgo = new Date(today);
    weekAgo.setDate(weekAgo.getDate() - 6);
    const weekAgoStr = formatDate(weekAgo);
    filtered = state.rawDailyData.filter(d => d.date >= weekAgoStr && d.date <= todayStr);
  } else if (state.period === 'month') {
    const startOfMonthStr = `${todayStr.substring(0, 7)}-01`;
    filtered = state.rawDailyData.filter(d => d.date >= startOfMonthStr && d.date <= todayStr);
  } else if (state.period === 'custom' && state.startDate && state.endDate) {
    filtered = state.rawDailyData.filter(d => d.date >= state.startDate && d.date <= state.endDate);
  } else {
    filtered = [...state.rawDailyData];
  }

  return filtered;
}

// 7. Dashboard Rendering Core
function renderDashboard() {
  const filteredDaily = getFilteredDailyData();

  // 1. Calculate Aggregations per Category & Channel
  const categoryCounts = {};
  CATEGORIES.forEach(cat => { categoryCounts[cat] = 0; });

  let grandTotal = 0;
  let websiteTotal = 0;
  let blogTotal = 0;
  let ytTotal = 0;

  filteredDaily.forEach(day => {
    if (day.categories) {
      CATEGORIES.forEach(cat => {
        const cnt = day.categories[cat] || 0;
        categoryCounts[cat] += cnt;

        const ch = CATEGORY_CHANNELS[cat];
        if (ch === '홈페이지') websiteTotal += cnt;
        else if (ch === '유튜브') ytTotal += cnt;
      });
    }
  });

  // Calculate blog total from blog-all
  blogTotal = categoryCounts['블로그-전체'] || 0;

  // Calculate active channel total
  let activeTotal = 0;
  if (state.selectedChannelTab === '홈페이지') activeTotal = websiteTotal;
  else if (state.selectedChannelTab === '네이버 블로그') activeTotal = blogTotal;
  else if (state.selectedChannelTab === '유튜브') activeTotal = ytTotal;
  else activeTotal = websiteTotal + blogTotal + ytTotal;

  // 2. Update Total Card & Title
  const totalCardTitle = document.getElementById('totalCardTitle');
  const totalIconWrap = document.getElementById('totalIconWrap');
  if (totalCardTitle) {
    const titleMap = {
      ALL: '전체 홍보 통합 실적',
      '홈페이지': '홈페이지 실적 요약',
      '네이버 블로그': '네이버 블로그 실적 요약',
      '유튜브': '유튜브 실적 요약'
    };
    totalCardTitle.textContent = titleMap[state.selectedChannelTab] || '홍보 통합 실적';
  }

  if (totalIconWrap) {
    totalIconWrap.className = 'kpi-icon-wrap';
    if (state.selectedChannelTab === '네이버 블로그') totalIconWrap.classList.add('green');
    else if (state.selectedChannelTab === '유튜브') totalIconWrap.classList.add('red');
    else totalIconWrap.classList.add('blue');
  }

  const totalCountEl = document.getElementById('totalCount');
  if (totalCountEl) animateValue(totalCountEl, parseInt(totalCountEl.textContent) || 0, activeTotal, 400);

  const periodSubtextEl = document.getElementById('periodSubtext');
  if (periodSubtextEl) {
    const labels = {
      today: '오늘 기준 실적',
      week: '최근 7일 누적 실적',
      month: '이번 달 누적 실적',
      all: '전체 수집 누적 실적',
      custom: `${state.startDate} ~ ${state.endDate} 실적`
    };
    periodSubtextEl.textContent = labels[state.period] || '누적 실적';
  }

  // 3. Render 15 Category KPI Cards
  CATEGORIES.forEach(cat => {
    const catEl = document.getElementById(`cat-${cat}`);
    if (catEl) {
      animateValue(catEl, parseInt(catEl.textContent) || 0, categoryCounts[cat], 400);
    }
  });

  // 4. Update Channel Group Containers Visibility based on selected Tab
  document.querySelectorAll('.channel-group-container').forEach(container => {
    const grp = container.dataset.channelGroup;
    if (state.selectedChannelTab === 'ALL' || state.selectedChannelTab === grp) {
      container.style.display = 'flex';
    } else {
      container.style.display = 'none';
    }
  });

  // 5. Render Charts
  renderTrendChart(filteredDaily);
  renderDoughnutChart(categoryCounts);

  // 6. Render Activity History Table
  renderHistoryTable();
}

function animateValue(obj, start, end, duration) {
  if (start === end) {
    obj.textContent = end;
    return;
  }
  let startTimestamp = null;
  const step = (timestamp) => {
    if (!startTimestamp) startTimestamp = timestamp;
    const progress = Math.min((timestamp - startTimestamp) / duration, 1);
    obj.textContent = Math.floor(progress * (end - start) + start);
    if (progress < 1) {
      window.requestAnimationFrame(step);
    }
  };
  window.requestAnimationFrame(step);
}

// 8. Chart Rendering Functions (Chart.js)
function renderTrendChart(dailyData) {
  const ctx = document.getElementById('trendChart')?.getContext('2d');
  if (!ctx) return;

  const sorted = [...dailyData].sort((a, b) => a.date.localeCompare(b.date));
  const labels = sorted.map(d => d.date.substring(5)); // 'MM-DD'

  // Calculate values based on selected channel tab
  const dataValues = sorted.map(d => {
    if (state.selectedChannelTab === 'ALL') {
      let sum = 0;
      CATEGORIES.forEach(c => {
        if (c !== '블로그-전체') sum += (d.categories?.[c] || 0);
      });
      return sum;
    } else {
      let sum = 0;
      CATEGORIES.forEach(c => {
        if (CATEGORY_CHANNELS[c] === state.selectedChannelTab) {
          if (state.selectedChannelTab === '네이버 블로그') {
            if (c === '블로그-전체') sum += (d.categories?.[c] || 0);
          } else {
            sum += (d.categories?.[c] || 0);
          }
        }
      });
      return sum;
    }
  });

  if (state.trendChartInstance) {
    state.trendChartInstance.destroy();
  }

  let strokeColor = '#3B82F6';
  let gradientStart = 'rgba(59, 130, 246, 0.35)';

  if (state.selectedChannelTab === '네이버 블로그') {
    strokeColor = '#03C75A';
    gradientStart = 'rgba(3, 199, 90, 0.35)';
  } else if (state.selectedChannelTab === '유튜브') {
    strokeColor = '#EF4444';
    gradientStart = 'rgba(239, 68, 68, 0.35)';
  }

  const gradient = ctx.createLinearGradient(0, 0, 0, 260);
  gradient.addColorStop(0, gradientStart);
  gradient.addColorStop(1, 'rgba(255, 255, 255, 0.0)');

  const config = {
    type: state.chartType,
    data: {
      labels: labels,
      datasets: [{
        label: '신규 콘텐츠 수 (건)',
        data: dataValues,
        borderColor: strokeColor,
        borderWidth: 2.5,
        backgroundColor: state.chartType === 'line' ? gradient : strokeColor,
        fill: state.chartType === 'line',
        tension: 0.35,
        pointBackgroundColor: '#FFFFFF',
        pointBorderColor: strokeColor,
        pointBorderWidth: 2,
        pointRadius: 4,
        pointHoverRadius: 6,
        borderRadius: state.chartType === 'bar' ? 6 : 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          titleFont: { family: 'Paperlogy', size: 13 },
          bodyFont: { family: 'Paperlogy', size: 12 },
          padding: 10,
          cornerRadius: 8
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { font: { family: 'Paperlogy', size: 11 }, color: '#64748B' }
        },
        y: {
          beginAtZero: true,
          grid: { color: 'rgba(226, 232, 240, 0.6)' },
          ticks: {
            stepSize: 1,
            font: { family: 'Paperlogy', size: 11 },
            color: '#64748B'
          }
        }
      }
    }
  };

  state.trendChartInstance = new Chart(ctx, config);
}

function renderDoughnutChart(categoryCounts) {
  const ctx = document.getElementById('doughnutChart')?.getContext('2d');
  if (!ctx) return;

  let targetCats = CATEGORIES;
  if (state.selectedChannelTab !== 'ALL') {
    targetCats = CATEGORIES.filter(c => CATEGORY_CHANNELS[c] === state.selectedChannelTab);
  } else {
    // In ALL tab, exclude blog-all to prevent double doughnut slice
    targetCats = CATEGORIES.filter(c => c !== '블로그-전체');
  }

  const labels = targetCats;
  const dataValues = targetCats.map(cat => categoryCounts[cat] || 0);
  const colors = targetCats.map(cat => CATEGORY_COLORS[cat]);

  if (state.doughnutChartInstance) {
    state.doughnutChartInstance.destroy();
  }

  const config = {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: dataValues,
        backgroundColor: colors,
        borderWidth: 2,
        borderColor: '#FFFFFF',
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '65%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            font: { family: 'Paperlogy', size: 10 },
            boxWidth: 10,
            padding: 8,
            color: '#334155'
          }
        },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.85)',
          titleFont: { family: 'Paperlogy', size: 13 },
          bodyFont: { family: 'Paperlogy', size: 12 },
          padding: 10,
          cornerRadius: 8
        }
      }
    }
  };

  state.doughnutChartInstance = new Chart(ctx, config);
}

// 9. Recent Activity History Table Logic
function renderHistoryTable() {
  const tbody = document.getElementById('activityTableBody');
  const countEl = document.getElementById('tableRecordCount');
  if (!tbody) return;

  let filtered = state.rawHistoryData.filter(item => {
    // Exclude blog-all duplicate items from table unless explicitly filtering blog-all
    if (item.category === '블로그-전체' && state.selectedCategoryFilter !== '블로그-전체') {
      return false;
    }

    const matchesSearch = state.searchQuery === '' ||
      item.title.toLowerCase().includes(state.searchQuery.toLowerCase()) ||
      item.post_id.toLowerCase().includes(state.searchQuery.toLowerCase());

    const matchesChannelTab = state.selectedChannelTab === 'ALL' ||
      item.channel === state.selectedChannelTab;

    const matchesTableChannel = state.selectedTableChannelFilter === 'ALL' ||
      item.channel === state.selectedTableChannelFilter;

    const matchesCategory = state.selectedCategoryFilter === 'ALL' ||
      item.category === state.selectedCategoryFilter;

    return matchesSearch && matchesChannelTab && matchesTableChannel && matchesCategory;
  });

  if (countEl) countEl.textContent = `총 ${filtered.length}건 수집됨`;

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="loading-cell">
          <i class="fa-solid fa-inbox"></i> 조건에 일치하는 수집 콘텐츠가 없습니다.
        </td>
      </tr>
    `;
    renderPagination(0);
    return;
  }

  const totalPages = Math.ceil(filtered.length / state.pageSize);
  if (state.currentPage > totalPages) state.currentPage = 1;
  const startIdx = (state.currentPage - 1) * state.pageSize;
  const pageItems = filtered.slice(startIdx, startIdx + state.pageSize);

  tbody.innerHTML = pageItems.map(item => {
    const ch = item.channel || CATEGORY_CHANNELS[item.category] || '홈페이지';
    let tagClass = 'tag-blue';
    let btnClass = '';
    
    if (ch === '네이버 블로그') {
      tagClass = 'tag-green';
      btnClass = 'btn-blog';
    } else if (ch === '유튜브') {
      tagClass = 'tag-red';
      btnClass = 'btn-yt';
    }

    const postUrl = item.url || CATEGORY_URLS[item.category] || 'https://gde.or.kr';

    return `
      <tr>
        <td><span class="post-id-badge">${item.post_id}</span></td>
        <td><span class="cat-tag ${tagClass}">${ch}</span></td>
        <td><span class="cat-tag ${tagClass}">${item.category}</span></td>
        <td class="post-title-cell">${item.title}</td>
        <td>${item.date}</td>
        <td style="font-size:0.8rem; color:#64748B;">${item.collected_at || item.date}</td>
        <td>
          <a href="${postUrl}" target="_blank" class="post-link-btn ${btnClass}" title="콘텐츠 원문 이동">
            <i class="fa-solid fa-arrow-up-right-from-square"></i>
          </a>
        </td>
      </tr>
    `;
  }).join('');

  renderPagination(totalPages);
}

function renderPagination(totalPages) {
  const container = document.getElementById('paginationControls');
  if (!container) return;

  if (totalPages <= 1) {
    container.innerHTML = '';
    return;
  }

  let html = '';
  for (let p = 1; p <= totalPages; p++) {
    html += `<button class="page-btn ${p === state.currentPage ? 'active' : ''}" data-page="${p}">${p}</button>`;
  }
  container.innerHTML = html;

  container.querySelectorAll('.page-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      state.currentPage = parseInt(e.target.dataset.page);
      renderHistoryTable();
    });
  });
}

// 10. Event Listeners Setup
function setupEventListeners() {
  // Channel Tab Selector
  document.querySelectorAll('.channel-tab').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.channel-tab').forEach(b => b.classList.remove('active'));
      const targetBtn = e.target.closest('.channel-tab');
      targetBtn.classList.add('active');
      state.selectedChannelTab = targetBtn.dataset.channel;
      state.currentPage = 1;
      renderDashboard();
    });
  });

  // Date Period Filters
  document.querySelectorAll('.filter-pill').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.filter-pill').forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      state.period = e.target.dataset.period;
      renderDashboard();
    });
  });

  const applyDateBtn = document.getElementById('applyDateBtn');
  if (applyDateBtn) {
    applyDateBtn.addEventListener('click', () => {
      const start = document.getElementById('startDateInput').value;
      const end = document.getElementById('endDateInput').value;
      if (start && end) {
        state.startDate = start;
        state.endDate = end;
        state.period = 'custom';
        document.querySelectorAll('.filter-pill').forEach(b => b.classList.remove('active'));
        renderDashboard();
      } else {
        alert('시작일과 종료일을 모두 선택해 주세요.');
      }
    });
  }

  const refreshBtn = document.getElementById('refreshBtn');
  if (refreshBtn) {
    refreshBtn.addEventListener('click', () => {
      loadData(true);
    });
  }

  const chartTypeLine = document.getElementById('chartTypeLine');
  const chartTypeBar = document.getElementById('chartTypeBar');
  if (chartTypeLine && chartTypeBar) {
    chartTypeLine.addEventListener('click', () => {
      chartTypeLine.classList.add('active');
      chartTypeBar.classList.remove('active');
      state.chartType = 'line';
      renderTrendChart(getFilteredDailyData());
    });
    chartTypeBar.addEventListener('click', () => {
      chartTypeBar.classList.add('active');
      chartTypeLine.classList.remove('active');
      state.chartType = 'bar';
      renderTrendChart(getFilteredDailyData());
    });
  }

  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      state.currentPage = 1;
      renderHistoryTable();
    });
  }

  const channelTableFilter = document.getElementById('channelTableFilter');
  if (channelTableFilter) {
    channelTableFilter.addEventListener('change', (e) => {
      state.selectedTableChannelFilter = e.target.value;
      state.currentPage = 1;
      renderHistoryTable();
    });
  }

  const categoryFilter = document.getElementById('categoryFilter');
  if (categoryFilter) {
    categoryFilter.addEventListener('change', (e) => {
      state.selectedCategoryFilter = e.target.value;
      state.currentPage = 1;
      renderHistoryTable();
    });
  }

  // Modal Dialog Listeners
  const apiModal = document.getElementById('apiModal');
  const apiConfigBtn = document.getElementById('apiConfigBtn');
  const modalCloseBtn = document.getElementById('modalCloseBtn');
  const saveGasUrlBtn = document.getElementById('saveGasUrlBtn');
  const resetGasUrlBtn = document.getElementById('resetGasUrlBtn');
  const gasUrlInput = document.getElementById('gasUrlInput');
  const demoModeToggle = document.getElementById('demoModeToggle');
  const currentGasUrlDisplay = document.getElementById('currentGasUrlDisplay');

  if (apiConfigBtn && apiModal) {
    apiConfigBtn.addEventListener('click', () => {
      if (gasUrlInput) gasUrlInput.value = state.gasUrl;
      if (demoModeToggle) demoModeToggle.checked = state.isDemoMode;
      if (currentGasUrlDisplay) currentGasUrlDisplay.textContent = state.gasUrl;
      apiModal.classList.add('open');
    });
  }

  if (modalCloseBtn && apiModal) {
    modalCloseBtn.addEventListener('click', () => {
      apiModal.classList.remove('open');
    });
  }

  if (saveGasUrlBtn && apiModal) {
    saveGasUrlBtn.addEventListener('click', () => {
      const newUrl = gasUrlInput.value.trim();
      state.gasUrl = newUrl || DEFAULT_GAS_URL;
      state.isDemoMode = demoModeToggle.checked;
      
      localStorage.setItem('gds_gas_url', state.gasUrl);
      localStorage.setItem('gds_demo_mode', state.isDemoMode ? 'true' : 'false');
      
      apiModal.classList.remove('open');
      loadData();
    });
  }

  if (resetGasUrlBtn) {
    resetGasUrlBtn.addEventListener('click', () => {
      state.gasUrl = DEFAULT_GAS_URL;
      state.isDemoMode = false;
      localStorage.removeItem('gds_gas_url');
      localStorage.removeItem('gds_demo_mode');
      if (gasUrlInput) gasUrlInput.value = DEFAULT_GAS_URL;
      if (demoModeToggle) demoModeToggle.checked = false;
      if (currentGasUrlDisplay) currentGasUrlDisplay.textContent = DEFAULT_GAS_URL;
    });
  }
}

// 11. App Initialization
document.addEventListener('DOMContentLoaded', () => {
  const todayStr = getTodayKST();
  const startDateInput = document.getElementById('startDateInput');
  const endDateInput = document.getElementById('endDateInput');
  if (startDateInput) startDateInput.value = todayStr;
  if (endDateInput) endDateInput.value = todayStr;

  updateKSTClock();
  setInterval(updateKSTClock, 1000);

  setupEventListeners();

  loadData();

  setInterval(() => {
    loadData();
  }, 5 * 60 * 1000);
});
