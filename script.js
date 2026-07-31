/**
 * 강동어울림복지관 홍보 통합실적 관리 시스템 - 프론트엔드 로직 (script.js)
 * GAS Web App API 연동 & Chart.js 시각화 & 실시간 필터링
 */

// 1. App Configuration & Constants
const DEFAULT_GAS_URL = 'https://script.google.com/macros/s/AKfycbxI-bCJ5RPaniBenLll1YRHLeySk2YQ_Cg3H_jJILcYP9Je7JoL2rpq7qSXthikAQcw/exec';

const CATEGORIES = [
  '공지사항',
  '공시자료',
  '인재채용',
  '이용인모집',
  '정보안내',
  '갤러리(전체)',
  '이용상담문의'
];

const CATEGORY_COLORS = {
  '공지사항': '#3B82F6',    // Blue
  '공시자료': '#2563EB',    // Darker Blue
  '인재채용': '#1D4ED8',    // Indigo Blue
  '이용인모집': '#06B6D4',  // Cyan
  '정보안내': '#6366F1',    // Indigo
  '갤러리(전체)': '#8B5CF6',// Purple
  '이용상담문의': '#10B981' // Emerald
};

const CATEGORY_URLS = {
  '공지사항': 'https://gde.or.kr/notice',
  '공시자료': 'https://gde.or.kr/infoopen',
  '인재채용': 'https://gde.or.kr/recruitment',
  '이용인모집': 'https://gde.or.kr/program',
  '정보안내': 'https://gde.or.kr/information',
  '갤러리(전체)': 'https://gde.or.kr/gallery',
  '이용상담문의': 'https://gde.or.kr/counseling'
};

// 2. Application State
const state = {
  gasUrl: localStorage.getItem('gds_gas_url') || DEFAULT_GAS_URL,
  isDemoMode: localStorage.getItem('gds_demo_mode') === 'true',
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

  // Generate last 30 days of data
  for (let i = 29; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const dateStr = formatDate(d);
    
    // Seeded random counts
    const cats = {};
    let totalDay = 0;
    
    // Higher activity on weekdays
    const isWeekend = (d.getDay() === 0 || d.getDay() === 6);
    const multiplier = isWeekend ? 0.3 : 1.0;

    CATEGORIES.forEach((cat, idx) => {
      const base = (idx % 3 === 0) ? 2 : 1;
      const count = Math.floor((Math.random() * 3 + base) * multiplier);
      cats[cat] = count;
      totalDay += count;

      // Generate history records
      for (let k = 0; k < count; k++) {
        const postId = `${cat.substring(0, 3)}_${dateStr.replace(/-/g, '')}_${k + 1}`;
        history.push({
          post_id: postId,
          main_category: cat === '갤러리(전체)' ? '어울림 갤러리' : (cat === '이용상담문의' ? '참여하기' : '어울림 소식'),
          sub_category: cat,
          category: cat,
          title: `[${cat}] 2026년 ${d.getMonth() + 1}월 ${cat} 소식지 및 사업 안내 #${k + 1}`,
          date: dateStr,
          collected_at: `${dateStr} ${String(9 + k).padStart(2, '0')}:15:00`
        });
      }
    });

    daily.push({
      date: dateStr,
      categories: cats,
      total: totalDay
    });
  }

  // Sort history descending by date
  history.sort((a, b) => new Date(b.date + ' ' + b.collected_at) - new Date(a.date + ' ' + a.collected_at));

  return { daily, history };
}

// 5. Data Fetching Logic (GAS API vs Demo Mode)
async function loadData(forceScrape = false) {
  setSyncStatus('syncing', forceScrape ? '실시간 스크래핑 및 동기화 중...' : '데이터 수집 중...');
  const refreshIcon = document.getElementById('refreshIcon');
  if (refreshIcon) refreshIcon.classList.add('fa-spin');

  try {
    if (state.isDemoMode) {
      // Demo mode fallback
      await new Promise(res => setTimeout(res, 400));
      const mock = generateMockData();
      state.rawDailyData = mock.daily;
      state.rawHistoryData = mock.history;
      setSyncStatus('online', '시연용 샘플 데이터');
    } else {
      // Fetch from GAS Web App API (trigger immediate scrape if forceScrape = true)
      const apiUrl = forceScrape ? `${state.gasUrl}?type=scrape` : `${state.gasUrl}?type=all`;
      const response = await fetch(apiUrl, {
        method: 'GET',
        mode: 'cors'
      });

      if (!response.ok) {
        throw new Error(`GAS API HTTP 에러: ${response.status}`);
      }

      const json = await response.json();
      if (json.status === 'success' && json.daily && json.history) {
        state.rawDailyData = json.daily;
        state.rawHistoryData = json.history;
        
        // Build hybrid daily aggregation from history if daily is missing today's row
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

/**
 * If history has items for a date that daily sheet lacks or has 0,
 * recalculate daily map directly from history items for full accuracy.
 */
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
    historyByDate[d].total += 1;
  });

  // Merge historyByDate into rawDailyData
  Object.keys(historyByDate).forEach(d => {
    const existingIndex = state.rawDailyData.findIndex(item => item.date === d);
    if (existingIndex !== -1) {
      // Replace if history has more comprehensive data
      if (historyByDate[d].total > (state.rawDailyData[existingIndex].total || 0)) {
        state.rawDailyData[existingIndex] = historyByDate[d];
      }
    } else {
      state.rawDailyData.push(historyByDate[d]);
    }
  });

  // Sort daily data ascending by date
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
    // 'all'
    filtered = [...state.rawDailyData];
  }

  return filtered;
}

// 7. Dashboard Rendering Core
function renderDashboard() {
  const filteredDaily = getFilteredDailyData();

  // 1. Calculate Aggregations
  let totalCount = 0;
  const categoryCounts = {};
  CATEGORIES.forEach(cat => { categoryCounts[cat] = 0; });

  filteredDaily.forEach(day => {
    totalCount += (day.total || 0);
    if (day.categories) {
      CATEGORIES.forEach(cat => {
        categoryCounts[cat] += (day.categories[cat] || 0);
      });
    }
  });

  // 2. Render Total KPI Card & Subtext
  const totalCountEl = document.getElementById('totalCount');
  if (totalCountEl) animateValue(totalCountEl, parseInt(totalCountEl.textContent) || 0, totalCount, 400);

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

  // 3. Render 7 Category Cards
  CATEGORIES.forEach(cat => {
    const catEl = document.getElementById(`cat-${cat}`);
    if (catEl) {
      animateValue(catEl, parseInt(catEl.textContent) || 0, categoryCounts[cat], 400);
    }
  });

  // 4. Render Charts
  renderTrendChart(filteredDaily);
  renderDoughnutChart(categoryCounts);

  // 5. Render History Table
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
  const dataValues = sorted.map(d => d.total);

  if (state.trendChartInstance) {
    state.trendChartInstance.destroy();
  }

  const gradient = ctx.createLinearGradient(0, 0, 0, 260);
  gradient.addColorStop(0, 'rgba(59, 130, 246, 0.35)');
  gradient.addColorStop(1, 'rgba(59, 130, 246, 0.0)');

  const config = {
    type: state.chartType,
    data: {
      labels: labels,
      datasets: [{
        label: '신규 게시글 수 (건)',
        data: dataValues,
        borderColor: '#3B82F6',
        borderWidth: 2.5,
        backgroundColor: state.chartType === 'line' ? gradient : 'rgba(59, 130, 246, 0.75)',
        fill: state.chartType === 'line',
        tension: 0.35,
        pointBackgroundColor: '#FFFFFF',
        pointBorderColor: '#1D4ED8',
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

  const labels = CATEGORIES;
  const dataValues = CATEGORIES.map(cat => categoryCounts[cat] || 0);
  const colors = CATEGORIES.map(cat => CATEGORY_COLORS[cat]);

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
      cutout: '68%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            font: { family: 'Paperlogy', size: 11 },
            boxWidth: 10,
            padding: 12,
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
    const matchesSearch = state.searchQuery === '' ||
      item.title.toLowerCase().includes(state.searchQuery.toLowerCase()) ||
      item.post_id.toLowerCase().includes(state.searchQuery.toLowerCase());

    const matchesCategory = state.selectedCategoryFilter === 'ALL' ||
      item.category === state.selectedCategoryFilter;

    return matchesSearch && matchesCategory;
  });

  if (countEl) countEl.textContent = `총 ${filtered.length}건 수집됨`;

  if (filtered.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" class="loading-cell">
          <i class="fa-solid fa-inbox"></i> 조건에 일치하는 수집 게시글이 없습니다.
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
    const postUrl = CATEGORY_URLS[item.category] || 'https://gde.or.kr';
    return `
      <tr>
        <td><span class="post-id-badge">${item.post_id}</span></td>
        <td>${item.main_category || '어울림 소식'}</td>
        <td><span class="cat-tag tag-blue">${item.category}</span></td>
        <td class="post-title-cell">${item.title}</td>
        <td>${item.date}</td>
        <td style="font-size:0.8rem; color:#64748B;">${item.collected_at || item.date}</td>
        <td>
          <a href="${postUrl}" target="_blank" class="post-link-btn" title="게시글 원문 이동">
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
      loadData(true); // Force immediate scrape on GAS
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

  const categoryFilter = document.getElementById('categoryFilter');
  if (categoryFilter) {
    categoryFilter.addEventListener('change', (e) => {
      state.selectedCategoryFilter = e.target.value;
      state.currentPage = 1;
      renderHistoryTable();
    });
  }

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
