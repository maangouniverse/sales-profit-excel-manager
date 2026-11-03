// 공통비용 8개 고정 항목 (드롭다운 선택용)
const COMMON_COST_CATEGORIES = [
  '종묘비',
  '무기질비료비',
  '유기질비료비',
  '농약비',
  '수도광열비',
  '기타재료비',
  '고용노동비',
  '기타'
];

// 초기 기본 데이터 (예시 수치를 모두 비운 실사용 초기 상태: 2026.11 ~ 2027.10 기준)
const INITIAL_STATE = {
  channels: ['농라', '스마트스토어', '카톡', '고흥몰', '남도장터', '로컬푸드', '공판장', '기타'],
  commonCostCategories: [...COMMON_COST_CATEGORIES],
  targetMonths: [
    { idx: 1, ym: '2026-11', label: '26년 11월', start: '2026-11-01', end: '2026-11-30', days: 30, target: 0 },
    { idx: 2, ym: '2026-12', label: '26년 12월', start: '2026-12-01', end: '2026-12-31', days: 31, target: 0 },
    { idx: 3, ym: '2027-01', label: '27년 1월',  start: '2027-01-01', end: '2027-01-31', days: 31, target: 0 },
    { idx: 4, ym: '2027-02', label: '27년 2월',  start: '2027-02-01', end: '2027-02-28', days: 28, target: 0 },
    { idx: 5, ym: '2027-03', label: '27년 3월',  start: '2027-03-01', end: '2027-03-31', days: 31, target: 0 },
    { idx: 6, ym: '2027-04', label: '27년 4월',  start: '2027-04-01', end: '2027-04-30', days: 30, target: 0 },
    { idx: 7, ym: '2027-05', label: '27년 5월',  start: '2027-05-01', end: '2027-05-31', days: 31, target: 0 },
    { idx: 8, ym: '2027-06', label: '27년 6월',  start: '2027-06-01', end: '2027-06-30', days: 30, target: 0 },
    { idx: 9, ym: '2027-07', label: '27년 7월',  start: '2027-07-01', end: '2027-07-31', days: 31, target: 0 },
    { idx: 10, ym: '2027-08', label: '27년 8월', start: '2027-08-01', end: '2027-08-31', days: 31, target: 0 },
    { idx: 11, ym: '2027-09', label: '27년 9월', start: '2027-09-01', end: '2027-09-30', days: 30, target: 0 },
    { idx: 12, ym: '2027-10', label: '27년 10월', start: '2027-10-01', end: '2027-10-31', days: 31, target: 0 }
  ],
  channelRates: [
    { id: 1, channel: '농라', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 2, channel: '스마트스토어', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 3, channel: '카톡', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 4, channel: '고흥몰', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 5, channel: '남도장터', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 6, channel: '로컬푸드', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 7, channel: '공판장', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
    { id: 8, channel: '기타', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' }
  ],
  dailySales: [],
  commonCosts: []
};

const STORAGE_KEY = 'sales_profit_manager_state_v4';
let state = loadState();
let charts = {};

// 도넛 차트 조각 내부에 흰색 퍼센트(26%, 24% 등)를 표시하는 플러그인 (매출성과.png 스타일)
const doughnutPercentPlugin = {
  id: 'doughnutPercentPlugin',
  afterDatasetsDraw(chart) {
    if (chart.config.type !== 'doughnut') return;
    const { ctx, data } = chart;
    const dataset = data.datasets[0];
    if (!dataset) return;
    const total = dataset.data.reduce((a, b) => a + Math.max(0, Number(b) || 0), 0);
    if (total <= 0) return;

    const meta = chart.getDatasetMeta(0);
    ctx.save();
    ctx.font = 'bold 12px Pretendard, Malgun Gothic, sans-serif';
    ctx.fillStyle = '#ffffff';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    meta.data.forEach((arc, i) => {
      const val = Math.max(0, Number(dataset.data[i]) || 0);
      const pct = Math.round((val / total) * 100);
      if (pct >= 5) {
        const pos = arc.tooltipPosition();
        ctx.fillText(`${pct}%`, pos.x, pos.y);
      }
    });
    ctx.restore();
  }
};
Chart.register(doughnutPercentPlugin);

function loadState() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      const parsed = JSON.parse(saved);
      parsed.commonCostCategories = [...COMMON_COST_CATEGORIES];
      if (Array.isArray(parsed.commonCosts)) {
        parsed.commonCosts = parsed.commonCosts.map(c => {
          const isValidCat = COMMON_COST_CATEGORIES.includes(c.category);
          return {
            ...c,
            category: isValidCat ? c.category : '기타',
            detail: c.detail !== undefined ? c.detail : (!isValidCat && c.category ? c.category : ''),
            amount: Number(c.amount) || 0,
            note: c.note || ''
          };
        });
      } else {
        parsed.commonCosts = [];
      }
      return parsed;
    }
  } catch (e) {
    console.warn(e);
  }
  return JSON.parse(JSON.stringify(INITIAL_STATE));
}

let toastTimer = null;

function showSaveToast(message) {
  const toast = document.getElementById('saveToast');
  const msgEl = document.getElementById('saveToastMsg');
  if (!toast || !msgEl) return;
  msgEl.textContent = message || '입력·수정하신 모든 내용이 안전하게 저장되었습니다!';
  toast.classList.add('show');
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove('show');
  }, 2600);
}

function updateSavedBadgeTimestamp() {
  const badge = document.getElementById('lastSavedBadge');
  if (!badge) return;
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  badge.textContent = `✅ 저장 완료 (${hh}:${mm}:${ss})`;
}

function saveState(toastMessage) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    updateSavedBadgeTimestamp();
    if (toastMessage) {
      showSaveToast(toastMessage);
    }
  } catch (e) {
    console.warn(e);
  }
}

window.triggerManualSave = function(customMsg) {
  // 현재 포커스된 입력칸이 있으면 변경사항 즉시 반영 후 저장
  if (document.activeElement && typeof document.activeElement.blur === 'function') {
    document.activeElement.blur();
  }
  saveState(customMsg || '💾 현재까지 입력·수정하신 모든 내용이 안전하게 저장되었습니다!');
  renderAll();
};

function fmtWon(n) {
  if (typeof n !== 'number' || isNaN(n)) return n;
  return Math.round(n).toLocaleString('ko-KR') + '원';
}

function fmtNum(n) {
  if (typeof n !== 'number' || isNaN(n)) return n;
  return Math.round(n).toLocaleString('ko-KR');
}

function fmtPct(r, digits = 1) {
  if (typeof r !== 'number' || isNaN(r)) return '0.0%';
  return (r * 100).toFixed(digits) + '%';
}

function escapeHtmlAttr(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

// 실적자동계산1.png 스타일 날짜 표시: "26년 11월 2일 (월)" + 토/일 요일 클래스
function formatKoreanDateWithDay(yyyyMmDd) {
  if (!yyyyMmDd) return { text: '', dayChar: '', cls: '' };
  const [y, m, d] = yyyyMmDd.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const dayIdx = dt.getUTCDay();
  const shortY = String(y).slice(2);
  const dayChar = days[dayIdx] || '';
  const text = `${shortY}년 ${m}월 ${d}일 (${dayChar})`;
  const cls = dayIdx === 6 ? 'date-sat' : dayIdx === 0 ? 'date-sun' : '';
  return { text, dayChar, cls };
}

// 현재 적용된 조회 시작일 ~ 종료일 계산
function getActivePeriod() {
  const mode = document.getElementById('filterMode').value;
  if (mode === 'custom') {
    const s = document.getElementById('filterStartDate').value || '2026-11-01';
    const e = document.getElementById('filterEndDate').value || '2027-10-31';
    return { start: s, end: e };
  } else {
    const ymVal = document.getElementById('filterYm').value;
    if (ymVal === 'all') {
      return { start: '2026-11-01', end: '2027-10-31' };
    }
    if (ymVal === 'year-2026') {
      return { start: '2026-01-01', end: '2026-12-31' };
    }
    if (ymVal === 'year-2027') {
      return { start: '2027-01-01', end: '2027-12-31' };
    }
    const found = state.targetMonths.find(t => t.ym === ymVal);
    if (found) {
      return { start: found.start, end: found.end };
    }
    return { start: '2026-11-01', end: '2026-11-30' };
  }
}

// 채널별 기간 설정 중복 검증
function checkRateOverlap(rate, allRates) {
  const s1 = rate.start;
  const e1 = rate.end || '2099-12-31';
  if (s1 > e1) return '⚠️ 시작일>종료일 오류';

  const overlaps = allRates.filter(other => {
    if (other.channel !== rate.channel) return false;
    const s2 = other.start;
    const e2 = other.end || '2099-12-31';
    return s2 <= e1 && e2 >= s1;
  });
  return overlaps.length > 1 ? '⚠️ 기간 중복' : '정상';
}

// 일일 매출 행 비용 및 순이익 계산 (10kg 미만/이상 구분 + 박스/포장비닐/리플렛 합산)
function evaluateDailyRow(row, allRates) {
  const u10 = Number(row.ordersUnder10) || 0;
  const o10 = Number(row.ordersOver10) || 0;
  const totalOrders = u10 + o10;

  const matched = allRates.filter(r => {
    if (r.channel !== row.channel) return false;
    const s = r.start;
    const e = r.end || '2099-12-31';
    return row.date >= s && row.date <= e;
  });

  if (matched.length === 0) {
    return { ...row, totalOrders, status: '비용 설정 필요', isValid: false };
  }
  if (matched.length > 1) {
    return { ...row, totalOrders, status: '기간 중복 확인', isValid: false };
  }

  const rate = matched[0];
  const feeRate = Number(rate.feeRate) || 0;
  const feeAmt = Math.round(row.sales * feeRate);
  const packSum = (Number(rate.boxCost) || 0) + (Number(rate.vinylCost) || 0) + (Number(rate.leafletCost) || 0);
  const unitUnder10 = (Number(rate.shipUnder10) || 0) + packSum;
  const unitOver10 = (Number(rate.shipOver10) || 0) + packSum;
  const orderCost = (u10 * unitUnder10) + (o10 * unitOver10);
  const totalCost = feeAmt + orderCost + (Number(row.extraCost) || 0);
  const netProfit = row.sales - totalCost;
  const margin = row.sales > 0 ? netProfit / row.sales : 0;

  return {
    ...row,
    totalOrders,
    status: '정상',
    isValid: true,
    feeRate,
    feeAmt,
    unitUnder10,
    unitOver10,
    orderCost,
    totalCost,
    netProfit,
    margin
  };
}

// 임의 조회기간에 대한 월별 일할 목표매출 계산
function calcProRatedTargetForMonth(tm, periodStart, periodEnd) {
  if (periodEnd < tm.start || periodStart > tm.end) return 0;
  if (periodStart <= tm.start && periodEnd >= tm.end) return tm.target;

  const overlapStart = periodStart > tm.start ? periodStart : tm.start;
  const overlapEnd = periodEnd < tm.end ? periodEnd : tm.end;
  const d1 = new Date(overlapStart + 'T00:00:00Z');
  const d2 = new Date(overlapEnd + 'T00:00:00Z');
  const overlapDays = Math.round((d2 - d1) / 86400000) + 1;
  return Math.round(tm.target * (overlapDays / tm.days));
}

// 전체 화면 렌더링
function renderAll() {
  const { start, end } = getActivePeriod();
  const d1 = new Date(start + 'T00:00:00Z');
  const d2 = new Date(end + 'T00:00:00Z');
  const diffDays = Math.max(1, Math.round((d2 - d1) / 86400000) + 1);
  document.getElementById('appliedPeriodBadge').textContent = `${start} ~ ${end} (${diffDays}일간)`;

  const totalAnnualTarget = state.targetMonths.reduce((a, b) => a + (Number(b.target) || 0), 0);
  document.getElementById('totalAnnualTargetBadge').textContent = fmtWon(totalAnnualTarget);

  // 1. 모든 일일 매출 행 평가
  const evaluatedSales = state.dailySales.map(r => evaluateDailyRow(r, state.channelRates));
  const missingRateCount = evaluatedSales.filter(r => r.status === '비용 설정 필요').length;
  const overlapRateCount = state.channelRates.filter(r => checkRateOverlap(r, state.channelRates) !== '정상').length;

  const banner = document.getElementById('healthStatusBanner');
  const icon = document.getElementById('healthStatusIcon');
  const title = document.getElementById('healthStatusTitle');
  const desc = document.getElementById('healthStatusDesc');
  if (missingRateCount === 0 && overlapRateCount === 0) {
    banner.className = 'health-banner ok';
    icon.textContent = '✅';
    title.textContent = '모든 채널 비용 설정 및 적용 기간 정상';
    desc.textContent = '비용 설정 누락 0건 · 채널 기간 중복 0건';
  } else {
    banner.className = 'health-banner warn';
    icon.textContent = '⚠️';
    title.textContent = '비용 설정 점검이 필요합니다!';
    desc.textContent = `비용 설정 누락 ${missingRateCount}건 · 채널 기간 중복/오류 ${overlapRateCount}건`;
  }

  // 2. 조회기간 내 유효 매출 필터링
  const periodRows = evaluatedSales.filter(r => r.isValid && r.date >= start && r.date <= end);

  let periodTarget = 0;
  state.targetMonths.forEach(tm => {
    periodTarget += calcProRatedTargetForMonth(tm, start, end);
  });

  const actualSales = periodRows.reduce((a, b) => a + b.sales, 0);
  const totalUnder10 = periodRows.reduce((a, b) => a + (Number(b.ordersUnder10) || 0), 0);
  const totalOver10 = periodRows.reduce((a, b) => a + (Number(b.ordersOver10) || 0), 0);
  const totalOrders = totalUnder10 + totalOver10;
  const totalVarCost = periodRows.reduce((a, b) => a + b.totalCost, 0);
  const netProfit = actualSales - totalVarCost;
  const netMargin = actualSales > 0 ? netProfit / actualSales : 0;
  const achieveRate = periodTarget > 0 ? actualSales / periodTarget : 0;

  const periodCommonCost = state.commonCosts
    .filter(c => c.date >= start && c.date <= end)
    .reduce((a, b) => a + b.amount, 0);

  const finalProfit = netProfit - periodCommonCost;

  // 대시보드 상단 요약 카드 업데이트 (매출성과.png 스타일 숫자 표기)
  document.getElementById('kpiActualSales').textContent = fmtNum(actualSales);
  document.getElementById('kpiNetProfit').textContent = fmtNum(netProfit);
  document.getElementById('kpiNetProfitCard').className = `perf-stat-card ${netProfit < 0 ? 'bar-red' : 'bar-olive'}`;
  document.getElementById('kpiTargetSales').textContent = fmtWon(periodTarget);
  document.getElementById('kpiAchieveRate').textContent = `목표 달성률 ${fmtPct(achieveRate, 1)}`;
  document.getElementById('kpiOrders').textContent = `${fmtNum(totalOrders)}건`;
  document.getElementById('kpiOrdersSplitSub').textContent = `10kg미만 ${fmtNum(totalUnder10)}건 · 10kg이상 ${fmtNum(totalOver10)}건`;
  document.getElementById('kpiTotalVarCost').textContent = fmtWon(totalVarCost);
  document.getElementById('kpiNetMargin').textContent = `순이익률 ${fmtPct(netMargin, 1)} (합계 기준)`;
  document.getElementById('kpiFinalProfit').textContent = fmtWon(finalProfit);
  document.getElementById('kpiCommonCostSub').textContent = `공통비용 ${fmtWon(periodCommonCost)} 차감 후`;

  // 대시보드 표 및 차트 업데이트
  renderDashboard(evaluatedSales, periodRows, periodTarget, actualSales, totalVarCost, netProfit, finalProfit, start, end, periodCommonCost);

  // 시트2: 일일 매출 입력 테이블 렌더링
  renderDailySalesTable(evaluatedSales, start, end);

  // 시트3: 판매채널 비용 설정 테이블 렌더링
  renderChannelRatesTable();

  // 시트4: 판매채널별 손익 분석 렌더링
  renderAnalysisTab(periodRows, actualSales, totalOrders, periodCommonCost);

  // 시트5: 목표매출(26.11~27.10) 및 공통비용 설정 렌더링
  renderTargetsAndCommonTab(start, end, totalAnnualTarget, periodTarget, periodCommonCost);

  // 시트6: 판매채널 마스터 목록 렌더링
  renderMasterSelectsAndLists();

  // 모든 날짜 입력칸에 클릭 시 달력 즉시 열기 연결
  attachCalendarPickers();
}

// 공통비용 8개 항목별 집계 헬퍼 (조회기간 / 연도별 / 전체누적)
function getCommonCostCategoryStats(start, end) {
  const categories = [...COMMON_COST_CATEGORIES];

  const totalPeriodCommon = state.commonCosts
    .filter(c => c.date >= start && c.date <= end)
    .reduce((a, b) => a + (Number(b.amount) || 0), 0);
  const totalAllCommon = state.commonCosts
    .reduce((a, b) => a + (Number(b.amount) || 0), 0);

  return categories.map(catName => {
    const catRows = state.commonCosts.filter(c => c.category === catName);
    const periodRows = catRows.filter(c => c.date >= start && c.date <= end);
    const periodCount = periodRows.length;
    const periodAmt = periodRows.reduce((a, b) => a + (Number(b.amount) || 0), 0);
    const year2026Amt = catRows
      .filter(c => String(c.date || '').startsWith('2026-'))
      .reduce((a, b) => a + (Number(b.amount) || 0), 0);
    const year2027Amt = catRows
      .filter(c => String(c.date || '').startsWith('2027-'))
      .reduce((a, b) => a + (Number(b.amount) || 0), 0);
    const allCount = catRows.length;
    const allAmt = catRows.reduce((a, b) => a + (Number(b.amount) || 0), 0);
    const periodShare = totalPeriodCommon > 0 ? periodAmt / totalPeriodCommon : 0;
    const allShare = totalAllCommon > 0 ? allAmt / totalAllCommon : 0;
    return { catName, periodCount, periodAmt, periodShare, year2026Amt, year2027Amt, allCount, allAmt, allShare };
  });
}

function renderDashboard(evaluatedSales, periodRows, periodTarget, actualSales, totalVarCost, netProfit, finalProfit, start, end, periodCommonCost) {
  const isMobile = window.innerWidth <= 768;
  const doughnutLegendPos = isMobile ? 'bottom' : 'right';

  // 1) 상단 히어로 막대 차트 (매출성과.png 우측 차트 스타일: 단위 천원)
  updateChart('chartTargetVsActual', {
    type: 'bar',
    data: {
      labels: isMobile ? ['목표매출', '총 매출액', '총 순이익', '최종순이익'] : ['목표매출', '총 매출액', '총 순이익', '최종순이익(공통비후)'],
      datasets: [{
        label: '단위: 천원',
        data: [
          Math.round(periodTarget / 1000),
          Math.round(actualSales / 1000),
          Math.round(netProfit / 1000),
          Math.round(finalProfit / 1000)
        ],
        backgroundColor: ['#92add0', '#2b4cbe', '#c5d185', '#5c7cfa'],
        barThickness: isMobile ? 26 : 38,
        borderRadius: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, grid: { color: '#f0f0f0' } }
      }
    }
  });

  // 2) 채널별 집계 후 매출액 TOP 5 & 순이익 TOP 5 (매출성과.png & 매출성과1.png 스타일)
  const chStats = state.channels.map(chName => {
    const rows = periodRows.filter(r => r.channel === chName);
    const s = rows.reduce((a, b) => a + b.sales, 0);
    const c = rows.reduce((a, b) => a + b.totalCost, 0);
    const p = s - c;
    const share = actualSales > 0 ? s / actualSales : 0;
    const mar = s > 0 ? p / s : 0;
    return { chName, sales: s, cost: c, profit: p, share, mar };
  });

  const topSales = [...chStats].sort((a, b) => b.sales - a.sales);
  const topProfit = [...chStats].sort((a, b) => b.profit - a.profit);

  const salesTbody = document.querySelector('#topSalesChannelTable tbody');
  salesTbody.innerHTML = '';
  topSales.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${idx + 1}</strong></td>
      <td class="name-cell">${item.chName}</td>
      <td class="num">${fmtNum(item.sales)}</td>
      <td class="num" style="color:#666;font-weight:600;">${fmtPct(item.share, 1)}</td>
    `;
    salesTbody.appendChild(tr);
  });

  const profitTbody = document.querySelector('#topProfitChannelTable tbody');
  profitTbody.innerHTML = '';
  topProfit.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${idx + 1}</strong></td>
      <td class="name-cell">${item.chName}</td>
      <td class="num ${item.profit < 0 ? 'cell-negative' : ''}">${fmtNum(item.profit)}</td>
      <td class="num" style="color:#666;font-weight:600;">${fmtPct(item.mar, 1)}</td>
    `;
    profitTbody.appendChild(tr);
  });

  const palette = ['#92add0', '#c1d4f5', '#112240', '#e6cc55', '#5c7cfa', '#c5d185', '#f2a68d', '#d0d7de'];

  // 매출액 비중 도넛 차트
  updateChart('chartChannelShare', {
    type: 'doughnut',
    data: {
      labels: topSales.map(x => `${x.chName} (${fmtPct(x.share, 1)})`),
      datasets: [{
        data: topSales.map(x => x.sales),
        backgroundColor: palette,
        borderWidth: 1,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '55%',
      plugins: {
        legend: { position: doughnutLegendPos, labels: { boxWidth: 10, font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const val = ctx.parsed || 0;
              const pct = actualSales > 0 ? ((val / actualSales) * 100).toFixed(1) : '0.0';
              return ` ${fmtWon(val)} (${pct}%)`;
            }
          }
        }
      }
    }
  });

  // 순이익 비중 도넛 차트 (양수 순이익 채널 비중)
  const posProfitTotal = topProfit.reduce((a, b) => a + Math.max(0, b.profit), 0);
  updateChart('chartChannelProfitShare', {
    type: 'doughnut',
    data: {
      labels: topProfit.map(x => x.chName),
      datasets: [{
        data: topProfit.map(x => Math.max(0, x.profit)),
        backgroundColor: palette,
        borderWidth: 1,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '55%',
      plugins: {
        legend: { position: doughnutLegendPos, labels: { boxWidth: 10, font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const val = ctx.parsed || 0;
              const pct = posProfitTotal > 0 ? ((val / posProfitTotal) * 100).toFixed(1) : '0.0';
              return ` ${fmtWon(val)} (${pct}%)`;
            }
          }
        }
      }
    }
  });

  // 3) 월별·연도별 손익현황 (26년 11월 ~ 27년 10월, 채널비용과 공통비용 별도 관리 및 손익 반영)
  const mLabels = [];
  const mSales = [];
  const mVarCosts = [];
  const mCommonCosts = [];
  const mFinalProfits = [];

  const mTbody = document.getElementById('dashboardMonthlyTbody');
  const mTfoot = document.getElementById('dashboardMonthlyTfoot');
  mTbody.innerHTML = '';

  let sumTarget = 0, sumSales = 0, sumCost = 0, sumProfit = 0, sumCommon = 0, sumFinalProfit = 0, sumOrd = 0;
  let y26 = { target: 0, sales: 0, cost: 0, profit: 0, common: 0, finalP: 0, ord: 0 };
  let y27 = { target: 0, sales: 0, cost: 0, profit: 0, common: 0, finalP: 0, ord: 0 };

  state.targetMonths.forEach(tm => {
    const rows = evaluatedSales.filter(r => r.isValid && r.date >= tm.start && r.date <= tm.end);
    const s = rows.reduce((a, b) => a + b.sales, 0);
    const c = rows.reduce((a, b) => a + b.totalCost, 0);
    const p = s - c;
    const cc = state.commonCosts
      .filter(x => x.date >= tm.start && x.date <= tm.end)
      .reduce((a, b) => a + (Number(b.amount) || 0), 0);
    const finalP = p - cc;
    const ord = rows.reduce((a, b) => a + b.totalOrders, 0);
    const ach = tm.target > 0 ? s / tm.target : 0;
    const finalMar = s > 0 ? finalP / s : 0;

    mLabels.push(tm.label);
    mSales.push(s);
    mVarCosts.push(c);
    mCommonCosts.push(cc);
    mFinalProfits.push(finalP);

    sumTarget += tm.target;
    sumSales += s;
    sumCost += c;
    sumProfit += p;
    sumCommon += cc;
    sumFinalProfit += finalP;
    sumOrd += ord;

    const yBucket = tm.ym.startsWith('2026') ? y26 : y27;
    yBucket.target += tm.target;
    yBucket.sales += s;
    yBucket.cost += c;
    yBucket.profit += p;
    yBucket.common += cc;
    yBucket.finalP += finalP;
    yBucket.ord += ord;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="연월" class="card-title-cell"><strong>${tm.label}</strong></td>
      <td data-label="목표매출" class="num">${fmtNum(tm.target)}</td>
      <td data-label="총 매출" class="num primary-metric"><strong>${s > 0 ? fmtNum(s) : '0'}</strong></td>
      <td data-label="달성률" class="num">${s > 0 ? fmtPct(ach, 1) : '0.0%'}</td>
      <td data-label="채널 비용(건당·수수료)" class="num">${c > 0 ? fmtNum(c) : '0'}</td>
      <td data-label="순이익(공통비 전)" class="num ${p < 0 ? 'cell-negative' : p > 0 ? 'profit-pos' : ''}">${s > 0 || c > 0 ? fmtNum(p) : '0'}</td>
      <td data-label="공통비용(8개 항목)" class="num">${cc > 0 ? fmtNum(cc) : '0'}</td>
      <td data-label="최종 순이익(공통비 후)" class="num primary-metric ${finalP < 0 ? 'cell-negative' : finalP > 0 ? 'profit-pos' : ''}"><strong>${s > 0 || c > 0 || cc > 0 ? fmtNum(finalP) : '0'}</strong></td>
      <td data-label="최종 이익률" class="num">${s > 0 ? fmtPct(finalMar, 2) : '0.00%'}</td>
      <td data-label="주문건수" class="num">${ord > 0 ? fmtNum(ord) + '건' : '0건'}</td>
    `;
    mTbody.appendChild(tr);
  });

  mTfoot.innerHTML = `
    <tr class="subtotal-row">
      <td data-label="구분" class="card-title-cell">2026년 합계 (11~12월)</td>
      <td data-label="목표매출" class="num">${fmtNum(y26.target)}</td>
      <td data-label="총 매출" class="num">${fmtNum(y26.sales)}</td>
      <td data-label="달성률" class="num">${fmtPct(y26.target > 0 ? y26.sales / y26.target : 0, 1)}</td>
      <td data-label="채널 비용" class="num">${fmtNum(y26.cost)}</td>
      <td data-label="순이익(공통비 전)" class="num ${y26.profit < 0 ? 'cell-negative' : ''}">${fmtNum(y26.profit)}</td>
      <td data-label="공통비용" class="num">${fmtNum(y26.common)}</td>
      <td data-label="최종 순이익" class="num ${y26.finalP < 0 ? 'cell-negative' : ''}">${fmtNum(y26.finalP)}</td>
      <td data-label="최종 이익률" class="num">${fmtPct(y26.sales > 0 ? y26.finalP / y26.sales : 0, 2)}</td>
      <td data-label="주문건수" class="num">${fmtNum(y26.ord)}건</td>
    </tr>
    <tr class="subtotal-row">
      <td data-label="구분" class="card-title-cell">2027년 합계 (1~10월)</td>
      <td data-label="목표매출" class="num">${fmtNum(y27.target)}</td>
      <td data-label="총 매출" class="num">${fmtNum(y27.sales)}</td>
      <td data-label="달성률" class="num">${fmtPct(y27.target > 0 ? y27.sales / y27.target : 0, 1)}</td>
      <td data-label="채널 비용" class="num">${fmtNum(y27.cost)}</td>
      <td data-label="순이익(공통비 전)" class="num ${y27.profit < 0 ? 'cell-negative' : ''}">${fmtNum(y27.profit)}</td>
      <td data-label="공통비용" class="num">${fmtNum(y27.common)}</td>
      <td data-label="최종 순이익" class="num ${y27.finalP < 0 ? 'cell-negative' : ''}">${fmtNum(y27.finalP)}</td>
      <td data-label="최종 이익률" class="num">${fmtPct(y27.sales > 0 ? y27.finalP / y27.sales : 0, 2)}</td>
      <td data-label="주문건수" class="num">${fmtNum(y27.ord)}건</td>
    </tr>
    <tr>
      <td data-label="구분" class="card-title-cell">12개월 총합계</td>
      <td data-label="목표매출 합계" class="num">${fmtNum(sumTarget)}</td>
      <td data-label="총 매출 합계" class="num">${fmtNum(sumSales)}</td>
      <td data-label="달성률" class="num">${fmtPct(sumTarget > 0 ? sumSales / sumTarget : 0, 1)}</td>
      <td data-label="채널 비용 합계" class="num">${fmtNum(sumCost)}</td>
      <td data-label="순이익(공통비 전)" class="num ${sumProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sumProfit)}</td>
      <td data-label="공통비용 총합계" class="num">${fmtNum(sumCommon)}</td>
      <td data-label="최종 순이익 합계" class="num ${sumFinalProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sumFinalProfit)}</td>
      <td data-label="최종 이익률" class="num">${fmtPct(sumSales > 0 ? sumFinalProfit / sumSales : 0, 2)}</td>
      <td data-label="주문건수 합계" class="num">${fmtNum(sumOrd)}건</td>
    </tr>
  `;

  updateChart('chartMonthlyTrend', {
    type: 'bar',
    data: {
      labels: mLabels,
      datasets: [
        { label: '총 매출', data: mSales, backgroundColor: '#2b4cbe', borderRadius: 3 },
        { label: '채널 비용(주문건당+수수료)', data: mVarCosts, backgroundColor: '#92add0', borderRadius: 3 },
        { label: '공통비용(8개 항목)', data: mCommonCosts, backgroundColor: '#f2a68d', borderRadius: 3 },
        { label: '최종 순이익(공통비 후)', data: mFinalProfits, backgroundColor: '#c5d185', borderRadius: 3 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } },
      scales: { x: { grid: { display: false } } }
    }
  });

  // 4) 공통비용 8개 항목별 지출 비중 표 & 도넛/막대 그래프
  const ccStats = getCommonCostCategoryStats(start, end);
  const ccTbody = document.querySelector('#dashboardCommonCostTable tbody');
  const ccTfoot = document.querySelector('#dashboardCommonCostTable tfoot');
  if (ccTbody && ccTfoot) {
    ccTbody.innerHTML = '';
    ccStats.forEach((item, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${idx + 1}</strong></td>
        <td class="name-cell">${item.catName}</td>
        <td class="num">${fmtNum(item.periodAmt)}</td>
        <td class="num" style="color:#666;font-weight:600;">${fmtPct(item.periodShare, 1)}</td>
      `;
      ccTbody.appendChild(tr);
    });
    ccTfoot.innerHTML = `
      <tr>
        <td colspan="2"><strong>공통비용 총합계 (조회기간)</strong></td>
        <td class="num"><strong>${fmtNum(periodCommonCost)}</strong></td>
        <td class="num"><strong>${periodCommonCost > 0 ? '100.0%' : '0.0%'}</strong></td>
      </tr>
    `;
  }

  updateChart('chartCommonCostShare', {
    type: 'doughnut',
    data: {
      labels: ccStats.map(x => `${x.catName} (${fmtPct(x.periodShare, 1)})`),
      datasets: [{
        data: ccStats.map(x => x.periodAmt),
        backgroundColor: palette,
        borderWidth: 1,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '55%',
      plugins: {
        legend: { position: doughnutLegendPos, labels: { boxWidth: 10, font: { size: 12 } } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const val = ctx.parsed || 0;
              const pct = periodCommonCost > 0 ? ((val / periodCommonCost) * 100).toFixed(1) : '0.0';
              return ` ${fmtWon(val)} (${pct}%)`;
            }
          }
        }
      }
    }
  });

  updateChart('chartCommonCostBar', {
    type: 'bar',
    data: {
      labels: ccStats.map(x => x.catName),
      datasets: [
        { label: '선택기간 지출액 (원)', data: ccStats.map(x => x.periodAmt), backgroundColor: '#2b4cbe', borderRadius: 3 },
        { label: '전체 누적 지출액 (원)', data: ccStats.map(x => x.allAmt), backgroundColor: '#92add0', borderRadius: 3 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, grid: { color: '#f0f0f0' } }
      }
    }
  });
}

function renderDailySalesTable(evaluatedSales, start, end) {
  const onlyPeriod = document.getElementById('chkFilterDailyByPeriod').checked;
  const rows = onlyPeriod
    ? evaluatedSales.filter(r => r.date >= start && r.date <= end)
    : evaluatedSales;

  const tbody = document.querySelector('#dailySalesTable tbody');
  const tfoot = document.querySelector('#dailySalesTable tfoot');
  tbody.innerHTML = '';

  let sumU10 = 0, sumO10 = 0, sumOrd = 0, sumSales = 0, sumExtra = 0, sumFee = 0, sumOrdCost = 0, sumCost = 0, sumProfit = 0;

  if (rows.length === 0) {
    const emptyTr = document.createElement('tr');
    emptyTr.className = 'empty-table-row';
    emptyTr.innerHTML = `<td colspan="18" class="empty-row-cell" style="text-align:center; padding:24px; color:#64748b;">등록된 일일 매출 내역이 없습니다. 상단 입력 칸에서 새로운 매출·주문 내역을 추가해 주세요.</td>`;
    tbody.appendChild(emptyTr);
  }

  rows.forEach(r => {
    sumU10 += Number(r.ordersUnder10) || 0;
    sumO10 += Number(r.ordersOver10) || 0;
    sumOrd += r.totalOrders;
    sumSales += r.sales;
    sumExtra += r.extraCost;
    if (r.isValid) {
      sumFee += r.feeAmt;
      sumOrdCost += r.orderCost;
      sumCost += r.totalCost;
      sumProfit += r.netProfit;
    }

    const dFormatted = formatKoreanDateWithDay(r.date);
    const errCls = !r.isValid ? 'cell-error' : '';
    const profitCls = !r.isValid ? 'cell-error' : r.netProfit < 0 ? 'cell-negative' : 'profit-pos';
    const chOpts = state.channels
      .map(ch => `<option value="${escapeHtmlAttr(ch)}" ${ch === r.channel ? 'selected' : ''}>${escapeHtmlAttr(ch)}</option>`)
      .join('');

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="일자 (달력 수정)" class="card-title-cell">
        <span class="cell-value-group">
          <input type="date" class="table-inline-date date-clickable" value="${r.date}" onchange="updateDailyField(${r.id}, 'date', this.value)" />
          <span class="date-day-badge ${dFormatted.cls}">(${dFormatted.dayChar})</span>
        </span>
      </td>
      <td data-label="판매채널" class="primary-metric">
        <select class="table-inline-select" onchange="updateDailyField(${r.id}, 'channel', this.value)">
          ${chOpts}
        </select>
      </td>
      <td data-label="10kg 미만 건수" class="num">
        <input type="number" class="table-inline-input" style="width:74px" min="0" step="1" value="${r.ordersUnder10}" onchange="updateDailyField(${r.id}, 'ordersUnder10', Math.max(0, Math.round(Number(this.value)||0)))" />
      </td>
      <td data-label="10kg 이상 건수" class="num">
        <input type="number" class="table-inline-input" style="width:74px" min="0" step="1" value="${r.ordersOver10}" onchange="updateDailyField(${r.id}, 'ordersOver10', Math.max(0, Math.round(Number(this.value)||0)))" />
      </td>
      <td data-label="총 주문건수" class="num primary-metric"><strong>${fmtNum(r.totalOrders)}</strong></td>
      <td data-label="총 매출 (일일 매출액)" class="num primary-metric">
        <input type="number" class="table-inline-input" style="width:110px" min="0" step="100" value="${r.sales}" onchange="updateDailyField(${r.id}, 'sales', Math.max(0, Number(this.value)||0))" />
      </td>
      <td data-label="추가 비용" class="num">
        <input type="number" class="table-inline-input" style="width:88px" min="0" step="100" value="${r.extraCost || 0}" onchange="updateDailyField(${r.id}, 'extraCost', Math.max(0, Number(this.value)||0))" />
      </td>
      <td data-label="비고">
        <input type="text" class="table-inline-text" value="${escapeHtmlAttr(r.note)}" onchange="updateDailyField(${r.id}, 'note', this.value)" placeholder="비고 입력" />
      </td>
      <td data-label="설정상태" class="${errCls}">${r.status}</td>
      <td data-label="적용 수수료율" class="num ${errCls}">${r.isValid ? fmtPct(r.feeRate, 2) : '비용 설정 필요'}</td>
      <td data-label="매출액 수수료" class="num ${errCls}">${r.isValid ? fmtNum(r.feeAmt) : '비용 설정 필요'}</td>
      <td data-label="10kg미만 단가" class="num ${errCls}">${r.isValid ? fmtNum(r.unitUnder10) : '비용 설정 필요'}</td>
      <td data-label="10kg이상 단가" class="num ${errCls}">${r.isValid ? fmtNum(r.unitOver10) : '비용 설정 필요'}</td>
      <td data-label="건수비용 합계" class="num ${errCls}">${r.isValid ? fmtNum(r.orderCost) : '비용 설정 필요'}</td>
      <td data-label="총 지출 (총비용)" class="num ${errCls}"><strong>${r.isValid ? fmtNum(r.totalCost) : '비용 설정 필요'}</strong></td>
      <td data-label="순이익" class="num primary-metric ${profitCls}">${r.isValid ? fmtNum(r.netProfit) : '비용 설정 필요'}</td>
      <td data-label="순이익률" class="num ${errCls}">${r.isValid ? fmtPct(r.margin, 1) : '비용 설정 필요'}</td>
      <td data-label="관리">
        <div class="row-action-group">
          <button type="button" class="btn-save-sm" onclick="triggerManualSave('해당 일일 매출·주문 내역이 저장되었습니다!')">💾 저장</button>
          <button type="button" class="btn-danger-sm" onclick="deleteDailyRow(${r.id})">삭제</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="2" data-label="구분" class="card-title-cell">합계 (${rows.length}건)</td>
      <td data-label="10kg 미만 합계" class="num">${fmtNum(sumU10)}</td>
      <td data-label="10kg 이상 합계" class="num">${fmtNum(sumO10)}</td>
      <td data-label="총 주문건수" class="num">${fmtNum(sumOrd)}</td>
      <td data-label="총 매출 합계" class="num">${fmtNum(sumSales)}</td>
      <td data-label="추가 비용 합계" class="num">${fmtNum(sumExtra)}</td>
      <td class="mobile-hide-empty"></td>
      <td data-label="집계 상태">정상 합산</td>
      <td class="num mobile-hide-empty">-</td>
      <td data-label="매출액 수수료 합계" class="num">${fmtNum(sumFee)}</td>
      <td class="num mobile-hide-empty">-</td>
      <td class="num mobile-hide-empty">-</td>
      <td data-label="건수비용 합계" class="num">${fmtNum(sumOrdCost)}</td>
      <td data-label="총 지출 합계" class="num">${fmtNum(sumCost)}</td>
      <td data-label="순이익 합계" class="num ${sumProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sumProfit)}</td>
      <td data-label="순이익률" class="num">${fmtPct(sumSales > 0 ? sumProfit / sumSales : 0, 1)}</td>
      <td class="mobile-hide-empty"></td>
    </tr>
  `;
}

function renderChannelRatesTable() {
  const tbody = document.querySelector('#channelRatesTable tbody');
  tbody.innerHTML = '';

  state.channelRates.forEach((r, idx) => {
    const packSum = (Number(r.boxCost) || 0) + (Number(r.vinylCost) || 0) + (Number(r.leafletCost) || 0);
    const sumUnder10 = (Number(r.shipUnder10) || 0) + packSum;
    const sumOver10 = (Number(r.shipOver10) || 0) + packSum;
    const status = checkRateOverlap(r, state.channelRates);
    const isErr = status !== '정상';
    const chOpts = state.channels
      .map(ch => `<option value="${escapeHtmlAttr(ch)}" ${ch === r.channel ? 'selected' : ''}>${escapeHtmlAttr(ch)}</option>`)
      .join('');

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="No" class="card-title-cell"><strong>#${idx + 1}</strong></td>
      <td data-label="판매채널" class="primary-metric">
        <select class="table-inline-select" onchange="updateRateField(${r.id}, 'channel', this.value)">
          ${chOpts}
        </select>
      </td>
      <td data-label="📅 적용 시작일"><input type="date" class="table-inline-date date-clickable" value="${r.start}" onchange="updateRateField(${r.id}, 'start', this.value)" /></td>
      <td data-label="📅 적용 종료일"><input type="date" class="table-inline-date date-clickable" value="${r.end || ''}" onchange="updateRateField(${r.id}, 'end', this.value)" /></td>
      <td data-label="수수료율 (%)" class="num">
        <span class="cell-value-group">
          <input type="number" class="table-inline-input" step="0.1" value="${(r.feeRate * 100).toFixed(2)}" onchange="updateRateField(${r.id}, 'feeRate', Number(this.value)/100)" />
          <span>%</span>
        </span>
      </td>
      <td data-label="10kg 미만 배송비" class="num"><input type="number" class="table-inline-input" step="100" value="${r.shipUnder10}" onchange="updateRateField(${r.id}, 'shipUnder10', Number(this.value))" /></td>
      <td data-label="10kg 이상 배송비" class="num"><input type="number" class="table-inline-input" step="100" value="${r.shipOver10}" onchange="updateRateField(${r.id}, 'shipOver10', Number(this.value))" /></td>
      <td data-label="박스 비용" class="num"><input type="number" class="table-inline-input" step="10" value="${r.boxCost}" onchange="updateRateField(${r.id}, 'boxCost', Number(this.value))" /></td>
      <td data-label="포장비닐 비용" class="num"><input type="number" class="table-inline-input" step="10" value="${r.vinylCost}" onchange="updateRateField(${r.id}, 'vinylCost', Number(this.value))" /></td>
      <td data-label="리플렛 비용" class="num"><input type="number" class="table-inline-input" step="10" value="${r.leafletCost}" onchange="updateRateField(${r.id}, 'leafletCost', Number(this.value))" /></td>
      <td data-label="10kg 미만 합계" class="num primary-metric"><strong>${fmtWon(sumUnder10)}</strong></td>
      <td data-label="10kg 이상 합계" class="num primary-metric"><strong>${fmtWon(sumOver10)}</strong></td>
      <td data-label="기간 중복 검증" class="${isErr ? 'cell-error' : ''}">${status}</td>
      <td data-label="비고">
        <input type="text" class="table-inline-text" value="${escapeHtmlAttr(r.note)}" onchange="updateRateField(${r.id}, 'note', this.value)" placeholder="비고 입력" />
      </td>
      <td data-label="관리">
        <div class="row-action-group">
          <button type="button" class="btn-save-sm" onclick="triggerManualSave('해당 채널 비용 설정이 저장되었습니다!')">💾 저장</button>
          <button type="button" class="btn-danger-sm" onclick="deleteRateRow(${r.id})">삭제</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function renderAnalysisTab(periodRows, totalSales, totalOrders, periodCommonCost) {
  const sortBy = document.getElementById('analysisSortBy').value;
  const allocRule = document.getElementById('analysisAllocRule').value;

  const channelStats = state.channels.map((name, idx) => {
    const rows = periodRows.filter(r => r.channel === name);
    const u10 = rows.reduce((a, b) => a + (Number(b.ordersUnder10) || 0), 0);
    const o10 = rows.reduce((a, b) => a + (Number(b.ordersOver10) || 0), 0);
    const orders = u10 + o10;
    const sales = rows.reduce((a, b) => a + b.sales, 0);
    const share = totalSales > 0 ? sales / totalSales : 0;
    const fee = rows.reduce((a, b) => a + b.feeAmt, 0);
    const orderCost = rows.reduce((a, b) => a + b.orderCost, 0);
    const extra = rows.reduce((a, b) => a + b.extraCost, 0);
    const totalCost = fee + orderCost + extra;
    const netProfit = sales - totalCost;
    const margin = sales > 0 ? netProfit / sales : 0;

    let allocCommon = 0;
    if (allocRule === 'sales') {
      allocCommon = Math.round(periodCommonCost * share);
    } else if (allocRule === 'orders') {
      allocCommon = Math.round(periodCommonCost * (totalOrders > 0 ? orders / totalOrders : 0));
    }
    const finalProfit = netProfit - allocCommon;
    const finalMargin = sales > 0 ? finalProfit / sales : 0;

    return { idx, name, u10, o10, orders, sales, share, fee, orderCost, extra, totalCost, netProfit, margin, allocCommon, finalProfit, finalMargin };
  });

  if (sortBy === 'profit') {
    channelStats.sort((a, b) => b.netProfit - a.netProfit);
  } else if (sortBy === 'sales') {
    channelStats.sort((a, b) => b.sales - a.sales);
  }

  const tbody = document.querySelector('#channelAnalysisTable tbody');
  const tfoot = document.querySelector('#channelAnalysisTable tfoot');
  tbody.innerHTML = '';

  let sU10 = 0, sO10 = 0, sOrd = 0, sSales = 0, sFee = 0, sOrdCost = 0, sExtra = 0, sCost = 0, sProfit = 0, sAlloc = 0, sFinal = 0;

  channelStats.forEach((st, i) => {
    sU10 += st.u10;
    sO10 += st.o10;
    sOrd += st.orders;
    sSales += st.sales;
    sFee += st.fee;
    sOrdCost += st.orderCost;
    sExtra += st.extra;
    sCost += st.totalCost;
    sProfit += st.netProfit;
    sAlloc += st.allocCommon;
    sFinal += st.finalProfit;

    const neg1 = st.netProfit < 0 ? 'cell-negative' : 'profit-pos';
    const neg2 = st.finalProfit < 0 ? 'cell-negative' : '';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="순위">${i + 1}위</td>
      <td data-label="판매채널명" class="card-title-cell"><strong>${st.name}</strong></td>
      <td data-label="10kg 미만 건수" class="num">${fmtNum(st.u10)}</td>
      <td data-label="10kg 이상 건수" class="num">${fmtNum(st.o10)}</td>
      <td data-label="총 주문건수" class="num primary-metric"><strong>${fmtNum(st.orders)}</strong></td>
      <td data-label="매출액 (원)" class="num primary-metric"><strong>${fmtNum(st.sales)}</strong></td>
      <td data-label="매출 비중" class="num">${fmtPct(st.share, 1)}</td>
      <td data-label="매출액 수수료" class="num">${fmtNum(st.fee)}</td>
      <td data-label="택배·포장비 합계" class="num">${fmtNum(st.orderCost)}</td>
      <td data-label="추가 비용" class="num">${fmtNum(st.extra)}</td>
      <td data-label="총비용(변동)" class="num">${fmtNum(st.totalCost)}</td>
      <td data-label="순이익 (공통비 전)" class="num primary-metric ${neg1}"><strong>${fmtNum(st.netProfit)}</strong></td>
      <td data-label="순이익률 (공통비 전)" class="num ${st.netProfit < 0 ? 'cell-negative' : ''}">${fmtPct(st.margin, 1)}</td>
      <td data-label="공통비용 배분액" class="num">${fmtNum(st.allocCommon)}</td>
      <td data-label="최종 순이익 (공통비 후)" class="num primary-metric ${neg2}"><strong>${fmtNum(st.finalProfit)}</strong></td>
      <td data-label="최종 이익률 (공통비 후)" class="num ${neg2}">${fmtPct(st.finalMargin, 1)}</td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="2" data-label="구분" class="card-title-cell">전체 ${channelStats.length}개 판매채널 합계</td>
      <td data-label="10kg 미만 합계" class="num">${fmtNum(sU10)}</td>
      <td data-label="10kg 이상 합계" class="num">${fmtNum(sO10)}</td>
      <td data-label="총 주문건수" class="num">${fmtNum(sOrd)}</td>
      <td data-label="매출액 합계" class="num">${fmtNum(sSales)}</td>
      <td data-label="매출 비중" class="num">100.0%</td>
      <td data-label="매출액 수수료 합계" class="num">${fmtNum(sFee)}</td>
      <td data-label="택배·포장비 합계" class="num">${fmtNum(sOrdCost)}</td>
      <td data-label="추가 비용 합계" class="num">${fmtNum(sExtra)}</td>
      <td data-label="총비용(변동) 합계" class="num">${fmtNum(sCost)}</td>
      <td data-label="순이익 합계 (공통비 전)" class="num ${sProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sProfit)}</td>
      <td data-label="순이익률 (공통비 전)" class="num">${fmtPct(sSales > 0 ? sProfit / sSales : 0, 1)}</td>
      <td data-label="공통비용 배분액 합계" class="num">${fmtNum(sAlloc)}</td>
      <td data-label="최종 순이익 합계" class="num ${sFinal < 0 ? 'cell-negative' : ''}">${fmtNum(sFinal)}</td>
      <td data-label="최종 이익률" class="num">${fmtPct(sSales > 0 ? sFinal / sSales : 0, 1)}</td>
    </tr>
  `;

  updateChart('chartChannelCompare', {
    type: 'bar',
    data: {
      labels: channelStats.map(c => c.name),
      datasets: [
        { label: '매출액 (원)', data: channelStats.map(c => c.sales), backgroundColor: '#2b4cbe', borderRadius: 3 },
        { label: '순이익 (공통비 전)', data: channelStats.map(c => c.netProfit), backgroundColor: '#c5d185', borderRadius: 3 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } }
    }
  });

  // 실적자동계산3.png 스타일: 일자별 순이익 변동 현황 (양수는 코랄 #c86464, 적자는 블루 #5271b5)
  const sortedByDate = [...periodRows].sort((a, b) => a.date.localeCompare(b.date));
  updateChart('chartDailyProfitFluctuation', {
    type: 'bar',
    data: {
      labels: sortedByDate.map(r => `${r.date.slice(5)} (${r.channel})`),
      datasets: [{
        label: '일별 순이익 (원)',
        data: sortedByDate.map(r => r.netProfit),
        backgroundColor: sortedByDate.map(r => r.netProfit >= 0 ? '#c86464' : '#5271b5'),
        barThickness: 12,
        borderRadius: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false } },
        y: { grid: { color: '#f0f0f0' } }
      }
    }
  });
}

function syncCommonEtcGuide() {
  const catSel = document.getElementById('inCommonCat');
  const detailIn = document.getElementById('inCommonDetail');
  const detailLabel = document.getElementById('inCommonDetailLabel');
  const guideBanner = document.getElementById('commonEtcGuideBanner');
  if (!catSel || !detailIn || !detailLabel || !guideBanner) return;

  const isEtc = catSel.value === '기타';
  const hasDetail = detailIn.value.trim().length > 0;

  if (isEtc) {
    guideBanner.style.display = 'block';
    detailLabel.innerHTML = `상세 내역 <span style="color:#dc2626;">* (‘기타’ 선택 시 입력 필수)</span>`;
    detailIn.placeholder = "⚠️ '기타' 선택 시 상세 내역을 꼭 입력해 주세요";
    if (!hasDetail) {
      detailIn.classList.add('etc-input-highlight');
    } else {
      detailIn.classList.remove('etc-input-highlight');
    }
  } else {
    guideBanner.style.display = 'none';
    detailLabel.textContent = '상세 내역';
    detailIn.placeholder = '구체적인 품목이나 지출 내역 입력';
    detailIn.classList.remove('etc-input-highlight');
  }
}

function renderTargetsAndCommonTab(start, end, totalAnnualTarget, periodTarget, periodCommonCost) {
  const sumAllCommon = state.commonCosts.reduce((a, b) => a + (Number(b.amount) || 0), 0);
  const sum2026Common = state.commonCosts
    .filter(c => String(c.date || '').startsWith('2026-'))
    .reduce((a, b) => a + (Number(b.amount) || 0), 0);
  const sum2027Common = state.commonCosts
    .filter(c => String(c.date || '').startsWith('2027-'))
    .reduce((a, b) => a + (Number(b.amount) || 0), 0);
  const periodCommonCount = state.commonCosts.filter(c => c.date >= start && c.date <= end).length;

  document.getElementById('tab5TotalTargetSum').textContent = fmtWon(totalAnnualTarget);
  document.getElementById('tab5AppliedTargetSum').textContent = fmtWon(periodTarget);
  document.getElementById('tab5PeriodCommonSum').textContent = fmtWon(periodCommonCost);

  const periodCountEl = document.getElementById('tab5PeriodCommonCountSub');
  if (periodCountEl) {
    periodCountEl.textContent = `조회기간 ${fmtNum(periodCommonCount)}건 (주문 1건당 비용과 별도)`;
  }
  const yearlyCommonEl = document.getElementById('tab5YearlyCommonSum');
  if (yearlyCommonEl) {
    yearlyCommonEl.textContent = `26년 ${fmtWon(sum2026Common)} / 27년 ${fmtWon(sum2027Common)}`;
  }
  const totalCommonEl = document.getElementById('tab5TotalCommonSum');
  if (totalCommonEl) {
    totalCommonEl.textContent = fmtWon(sumAllCommon);
  }
  const totalCommonCountEl = document.getElementById('tab5TotalCommonCountSub');
  if (totalCommonCountEl) {
    totalCommonCountEl.textContent = `8개 항목 총 ${fmtNum(state.commonCosts.length)}건 입력됨`;
  }

  // 월별 목표매출 표
  const tTbody = document.querySelector('#monthlyTargetsTable tbody');
  const tTfoot = document.querySelector('#monthlyTargetsTable tfoot');
  tTbody.innerHTML = '';

  state.targetMonths.forEach(tm => {
    const daily = Math.round(tm.target / tm.days);
    const applied = calcProRatedTargetForMonth(tm, start, end);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="대상 연월" class="card-title-cell"><strong>${tm.label}</strong></td>
      <td data-label="적용 기간" style="color:#666;">${tm.start} ~ ${tm.end}</td>
      <td data-label="총 일수">${tm.days}일</td>
      <td data-label="월별 목표매출 (원)" class="num primary-metric">
        <input type="number" class="table-inline-input" style="width:130px" step="100000" value="${tm.target}" onchange="updateMonthlyTarget('${tm.ym}', Number(this.value))" />
      </td>
      <td data-label="일할 목표매출 (원/일)" class="num">${fmtWon(daily)}</td>
      <td data-label="조회기간 반영 목표매출" class="num primary-metric"><strong>${fmtWon(applied)}</strong></td>
    `;
    tTbody.appendChild(tr);
  });

  tTfoot.innerHTML = `
    <tr>
      <td colspan="2" data-label="구분" class="card-title-cell">12개월 목표매출 총합계 (2026.11 ~ 2027.10)</td>
      <td data-label="총 일수">365일</td>
      <td data-label="목표매출 총합계" class="num">${fmtWon(totalAnnualTarget)}</td>
      <td data-label="평균 일할 목표" class="num">${fmtWon(Math.round(totalAnnualTarget / 365))}/일</td>
      <td data-label="조회기간 반영 합계" class="num">${fmtWon(periodTarget)}</td>
    </tr>
  `;

  // 1) 공통비용 입력 시트 테이블 (발생일자 · 비용 항목 · 상세 내역 · 금액 · 비고)
  const cTbody = document.querySelector('#commonCostsTable tbody');
  const cTfoot = document.querySelector('#commonCostsTable tfoot');
  cTbody.innerHTML = '';

  const filterChk = document.getElementById('chkFilterCommonByPeriod');
  const onlyPeriodCommon = filterChk ? filterChk.checked : false;
  const visibleCommonRows = onlyPeriodCommon
    ? state.commonCosts.filter(c => c.date >= start && c.date <= end)
    : state.commonCosts;

  if (visibleCommonRows.length === 0) {
    const emptyTr = document.createElement('tr');
    emptyTr.className = 'empty-table-row';
    emptyTr.innerHTML = `<td colspan="8" class="empty-row-cell" style="text-align:center; padding:24px; color:#64748b;">등록된 공통비용 내역이 없습니다. 상단 입력 칸에서 [발생일자 · 비용 항목 · 상세 내역 · 금액 · 비고]를 입력해 추가해 주세요.</td>`;
    cTbody.appendChild(emptyTr);
  }

  let visibleSum = 0;
  visibleCommonRows.forEach((c, idx) => {
    const inPeriod = c.date >= start && c.date <= end;
    const amt = Number(c.amount) || 0;
    visibleSum += amt;
    const dFormatted = formatKoreanDateWithDay(c.date);
    const ymStr = c.date ? c.date.slice(0, 7) : '-';
    const isEtcEmpty = c.category === '기타' && !(c.detail || '').trim();

    const catOpts = COMMON_COST_CATEGORIES
      .map(cat => `<option value="${escapeHtmlAttr(cat)}" ${cat === c.category ? 'selected' : ''}>${escapeHtmlAttr(cat)}</option>`)
      .join('');

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td data-label="No" class="card-title-cell"><strong>#${idx + 1}</strong></td>
      <td data-label="발생일자">
        <span class="cell-value-group">
          <input type="date" class="table-inline-date date-clickable" value="${c.date}" onchange="updateCommonField(${c.id}, 'date', this.value)" />
          <span class="date-day-badge ${dFormatted.cls}">(${dFormatted.dayChar})</span>
        </span>
      </td>
      <td data-label="비용 항목" class="primary-metric">
        <select class="table-inline-select" style="width:130px;" onchange="updateCommonField(${c.id}, 'category', this.value)">
          ${catOpts}
        </select>
      </td>
      <td data-label="상세 내역" style="text-align:left;">
        <input type="text" class="table-inline-text ${isEtcEmpty ? 'etc-input-highlight' : ''}" style="width:210px;" value="${escapeHtmlAttr(c.detail || '')}" onchange="updateCommonField(${c.id}, 'detail', this.value)" placeholder="${c.category === '기타' ? "⚠️ '기타' 상세 내역을 입력해 주세요" : '상세 내역 입력'}" />
        ${isEtcEmpty ? `<span class="etc-inline-hint">⚠️ ‘기타’를 선택한 경우 상세 내역을 입력해 주세요</span>` : ''}
      </td>
      <td data-label="금액 (원)" class="num primary-metric">
        <input type="number" class="table-inline-input" style="width:115px;" min="0" step="100" value="${amt}" onchange="updateCommonField(${c.id}, 'amount', Math.max(0, Number(this.value)||0))" />
      </td>
      <td data-label="비고" style="text-align:left;">
        <input type="text" class="table-inline-text" style="width:170px;" value="${escapeHtmlAttr(c.note || '')}" onchange="updateCommonField(${c.id}, 'note', this.value)" placeholder="비고 입력" />
      </td>
      <td data-label="귀속연월 · 조회기간 반영">
        <span style="font-weight:700; color:#334155;">${ymStr}</span>
        <span style="margin-left:4px; color:${inPeriod ? '#0f683c' : '#94a3b8'}; font-weight:700;">(${inPeriod ? '✅ 조회기간 포함' : '기간 외'})</span>
      </td>
      <td data-label="관리">
        <div class="row-action-group">
          <button type="button" class="btn-save-sm" onclick="triggerManualSave('해당 공통비용 내역이 저장되었습니다!')">💾 저장</button>
          <button type="button" class="btn-danger-sm" onclick="deleteCommonCost(${c.id})">삭제</button>
        </div>
      </td>
    `;
    cTbody.appendChild(tr);
  });

  cTfoot.innerHTML = `
    <tr>
      <td colspan="4" data-label="구분" class="card-title-cell">공통비용 총합계 (표시 ${visibleCommonRows.length}건 / 전체 ${state.commonCosts.length}건)</td>
      <td data-label="표시 내역 합계" class="num">${fmtWon(visibleSum)}</td>
      <td data-label="연도별 합계">2026년: ${fmtWon(sum2026Common)} / 2027년: ${fmtWon(sum2027Common)}</td>
      <td colspan="2" data-label="조회기간 및 전체 총합계">조회기간: <strong>${fmtWon(periodCommonCost)}</strong> | 전체 총합계: <strong>${fmtWon(sumAllCommon)}</strong></td>
    </tr>
  `;

  // 2) 공통비용 8개 항목별 합계 및 총합계 집계표 & 그래프 업데이트
  const ccStats = getCommonCostCategoryStats(start, end);
  const catTbody = document.querySelector('#tab5CommonCategoryTable tbody');
  const catTfoot = document.querySelector('#tab5CommonCategoryTable tfoot');
  if (catTbody && catTfoot) {
    catTbody.innerHTML = '';
    ccStats.forEach((st, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td data-label="No">${idx + 1}</td>
        <td data-label="비용 항목 (8개 구분)" class="card-title-cell"><strong>${st.catName}</strong></td>
        <td data-label="조회기간 건수" class="num">${fmtNum(st.periodCount)}건</td>
        <td data-label="조회기간 합계 (원)" class="num primary-metric"><strong>${fmtWon(st.periodAmt)}</strong></td>
        <td data-label="조회기간 비중" class="num">${fmtPct(st.periodShare, 1)}</td>
        <td data-label="2026년 합계 (원)" class="num">${fmtWon(st.year2026Amt)}</td>
        <td data-label="2027년 합계 (원)" class="num">${fmtWon(st.year2027Amt)}</td>
        <td data-label="공통비용 총합계 (전체)" class="num primary-metric"><strong>${fmtWon(st.allAmt)}</strong></td>
        <td data-label="전체 비중" class="num">${fmtPct(st.allShare, 1)}</td>
      `;
      catTbody.appendChild(tr);
    });
    catTfoot.innerHTML = `
      <tr>
        <td colspan="2" data-label="구분" class="card-title-cell">공통비용 총합계 (8개 항목 합계)</td>
        <td data-label="조회기간 총건수" class="num">${fmtNum(periodCommonCount)}건</td>
        <td data-label="조회기간 총합계" class="num">${fmtWon(periodCommonCost)}</td>
        <td data-label="조회기간 비중" class="num">${periodCommonCost > 0 ? '100.0%' : '0.0%'}</td>
        <td data-label="2026년 총합계" class="num">${fmtWon(sum2026Common)}</td>
        <td data-label="2027년 총합계" class="num">${fmtWon(sum2027Common)}</td>
        <td data-label="공통비용 총합계(전체)" class="num">${fmtWon(sumAllCommon)}</td>
        <td data-label="전체 비중" class="num">${sumAllCommon > 0 ? '100.0%' : '0.0%'}</td>
      </tr>
    `;
  }

  // 3) 발생일자 기준 월별 · 연도별 · 조회기간별 공통비용 자동 집계표
  const monthlyTbody = document.getElementById('tab5CommonMonthlyTbody');
  const monthlyTfoot = document.getElementById('tab5CommonMonthlyTfoot');
  if (monthlyTbody && monthlyTfoot) {
    monthlyTbody.innerHTML = '';

    // 기본 12개월(2026-11 ~ 2027-10) + 그 외 발생월이 있을 경우 자동 추가
    const monthMap = new Map();
    state.targetMonths.forEach(tm => {
      monthMap.set(tm.ym, tm.label);
    });
    state.commonCosts.forEach(c => {
      if (c.date && c.date.length >= 7) {
        const ym = c.date.slice(0, 7);
        if (!monthMap.has(ym)) {
          const [y, m] = ym.split('-');
          monthMap.set(ym, `${String(y).slice(2)}년 ${Number(m)}월`);
        }
      }
    });

    const sortedYmList = Array.from(monthMap.keys()).sort();

    const buildCatCells = (rowsForGroup) => {
      let rowTotal = 0;
      const cellsHtml = COMMON_COST_CATEGORIES.map(catName => {
        const amt = rowsForGroup
          .filter(c => c.category === catName)
          .reduce((a, b) => a + (Number(b.amount) || 0), 0);
        rowTotal += amt;
        return `<td data-label="${catName}" class="num">${amt > 0 ? fmtNum(amt) : '0'}</td>`;
      }).join('');
      return { cellsHtml, rowTotal, count: rowsForGroup.length };
    };

    sortedYmList.forEach(ym => {
      const label = monthMap.get(ym);
      const mRows = state.commonCosts.filter(c => String(c.date || '').slice(0, 7) === ym);
      const { cellsHtml, rowTotal, count } = buildCatCells(mRows);

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td data-label="구분 (월별)" class="card-title-cell"><strong>${label}</strong> <span style="color:#888; font-size:11.5px;">(${ym})</span></td>
        ${cellsHtml}
        <td data-label="공통비용 총합계 (원)" class="num primary-metric"><strong>${fmtWon(rowTotal)}</strong></td>
        <td data-label="입력 건수" class="num">${fmtNum(count)}건</td>
      `;
      monthlyTbody.appendChild(tr);
    });

    const rows2026 = state.commonCosts.filter(c => String(c.date || '').startsWith('2026-'));
    const rows2027 = state.commonCosts.filter(c => String(c.date || '').startsWith('2027-'));
    const rowsPeriod = state.commonCosts.filter(c => c.date >= start && c.date <= end);
    const rowsAll = state.commonCosts;

    const sum26 = buildCatCells(rows2026);
    const sum27 = buildCatCells(rows2027);
    const sumPer = buildCatCells(rowsPeriod);
    const sumAll = buildCatCells(rowsAll);

    monthlyTfoot.innerHTML = `
      <tr class="subtotal-row">
        <td data-label="구분" class="card-title-cell">📆 2026년 연도별 합계</td>
        ${sum26.cellsHtml}
        <td data-label="2026년 총합계" class="num"><strong>${fmtWon(sum26.rowTotal)}</strong></td>
        <td data-label="2026년 건수" class="num">${fmtNum(sum26.count)}건</td>
      </tr>
      <tr class="subtotal-row">
        <td data-label="구분" class="card-title-cell">📆 2027년 연도별 합계</td>
        ${sum27.cellsHtml}
        <td data-label="2027년 총합계" class="num"><strong>${fmtWon(sum27.rowTotal)}</strong></td>
        <td data-label="2027년 건수" class="num">${fmtNum(sum27.count)}건</td>
      </tr>
      <tr class="period-highlight-row">
        <td data-label="구분" class="card-title-cell">📅 현재 조회기간 합계 (${start} ~ ${end})</td>
        ${sumPer.cellsHtml}
        <td data-label="조회기간 총합계" class="num"><strong>${fmtWon(sumPer.rowTotal)}</strong></td>
        <td data-label="조회기간 건수" class="num">${fmtNum(sumPer.count)}건</td>
      </tr>
      <tr>
        <td data-label="구분" class="card-title-cell">🏢 공통비용 총합계 (전체 누적)</td>
        ${sumAll.cellsHtml}
        <td data-label="공통비용 총합계" class="num"><strong>${fmtWon(sumAll.rowTotal)}</strong></td>
        <td data-label="총 입력 건수" class="num">${fmtNum(sumAll.count)}건</td>
      </tr>
    `;
  }

  const palette = ['#92add0', '#c1d4f5', '#112240', '#e6cc55', '#5c7cfa', '#c5d185', '#f2a68d', '#d0d7de'];
  updateChart('chartTab5CommonDoughnut', {
    type: 'doughnut',
    data: {
      labels: ccStats.map(x => `${x.catName} (${fmtPct(x.periodShare, 1)})`),
      datasets: [{
        data: ccStats.map(x => x.periodAmt),
        backgroundColor: palette,
        borderWidth: 1,
        borderColor: '#ffffff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '55%',
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 11 } } }
      }
    }
  });

  updateChart('chartTab5CommonBar', {
    type: 'bar',
    data: {
      labels: ccStats.map(x => x.catName),
      datasets: [
        { label: '조회기간 합계', data: ccStats.map(x => x.periodAmt), backgroundColor: '#2b4cbe', borderRadius: 3 },
        { label: '전체 총합계', data: ccStats.map(x => x.allAmt), backgroundColor: '#92add0', borderRadius: 3 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } },
      scales: {
        x: { grid: { display: false } },
        y: { beginAtZero: true, grid: { color: '#f0f0f0' } }
      }
    }
  });

  syncCommonEtcGuide();
}

function renderMasterSelectsAndLists() {
  const chSel1 = document.getElementById('inSaleChannel');
  const chSel2 = document.getElementById('inRateChannel');
  const prevCh1 = chSel1.value;
  const prevCh2 = chSel2.value;

  const chOptions = state.channels.map(c => `<option value="${c}">${c}</option>`).join('');
  chSel1.innerHTML = chOptions;
  chSel2.innerHTML = chOptions;

  if (state.channels.includes(prevCh1)) chSel1.value = prevCh1;
  if (state.channels.includes(prevCh2)) chSel2.value = prevCh2;

  document.getElementById('channelMasterList').innerHTML = state.channels.map(c => `<li>🛒 ${c}</li>`).join('');
}

function updateChart(canvasId, config) {
  const el = document.getElementById(canvasId);
  if (!el) return;
  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }
  const ctx = el.getContext('2d');
  charts[canvasId] = new Chart(ctx, config);
}

// 날짜 입력창 클릭 시 브라우저 달력(showPicker) 즉시 호출
function attachCalendarPickers() {
  document.querySelectorAll('input[type="date"]').forEach(inp => {
    if (inp.dataset.pickerAttached) return;
    inp.dataset.pickerAttached = '1';
    inp.addEventListener('click', () => {
      if (typeof inp.showPicker === 'function') {
        try { inp.showPicker(); } catch (e) {}
      }
    });
  });
}

// ============================================================================
// 전역 액션 함수
// ============================================================================
window.deleteDailyRow = function(id) {
  state.dailySales = state.dailySales.filter(r => r.id !== id);
  saveState('🗑️ 일일 매출 내역이 삭제 및 저장되었습니다.');
  renderAll();
};

window.updateDailyField = function(id, field, val) {
  const target = state.dailySales.find(r => r.id === id);
  if (target) {
    target[field] = val;
    saveState('✅ 일일 매출·주문 수정사항이 저장되었습니다.');
    renderAll();
  }
};

window.deleteRateRow = function(id) {
  state.channelRates = state.channelRates.filter(r => r.id !== id);
  saveState('🗑️ 판매채널 비용 설정이 삭제 및 저장되었습니다.');
  renderAll();
};

window.updateRateField = function(id, field, val) {
  const target = state.channelRates.find(r => r.id === id);
  if (target) {
    target[field] = val;
    saveState('✅ 판매채널 비용 설정 수정사항이 저장되었습니다.');
    renderAll();
  }
};

window.updateMonthlyTarget = function(ym, val) {
  const t = state.targetMonths.find(x => x.ym === ym);
  if (t) {
    t.target = Math.max(0, val || 0);
    saveState('✅ 월별 목표매출이 저장되었습니다.');
    renderAll();
  }
};

window.deleteCommonCost = function(id) {
  state.commonCosts = state.commonCosts.filter(c => c.id !== id);
  saveState('🗑️ 공통비용 내역이 삭제 및 저장되었습니다.');
  renderAll();
};

window.updateCommonField = function(id, field, val) {
  const target = state.commonCosts.find(c => c.id === id);
  if (target) {
    target[field] = val;
    if (field === 'category' && val === '기타' && !(target.detail || '').trim()) {
      saveState("💡 '기타' 항목을 선택하셨습니다. [상세 내역]을 입력해 주세요.");
    } else {
      saveState('✅ 공통비용 수정사항이 저장되었습니다.');
    }
    renderAll();
  }
};

document.addEventListener('DOMContentLoaded', () => {
  // 상단 '현재 내용 저장하기' 버튼 연결
  const btnManualSave = document.getElementById('btnManualSave');
  if (btnManualSave) {
    btnManualSave.addEventListener('click', () => {
      window.triggerManualSave('💾 현재까지 입력·수정하신 모든 내용이 안전하게 저장되었습니다!');
    });
  }

  // 탭 전환 및 상단 부제목 업데이트
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
      if (btn.dataset.subtitle) {
        document.getElementById('currentSheetSubtitle').textContent = btn.dataset.subtitle;
      }
      renderAll();
    });
  });

  // 조회 방식 전환
  document.getElementById('filterMode').addEventListener('change', e => {
    const isCustom = e.target.value === 'custom';
    document.getElementById('yearMonthGroup').style.display = isCustom ? 'none' : 'block';
    document.getElementById('customDateGroup').style.display = isCustom ? 'block' : 'none';
    renderAll();
  });

  ['filterYm', 'filterStartDate', 'filterEndDate', 'chkFilterDailyByPeriod', 'chkFilterCommonByPeriod', 'analysisSortBy', 'analysisAllocRule'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderAll);
  });

  // 공통비용 '기타' 선택 시 상세 내역 입력 안내 연동
  const inCommonCatEl = document.getElementById('inCommonCat');
  const inCommonDetailEl = document.getElementById('inCommonDetail');
  if (inCommonCatEl) {
    inCommonCatEl.addEventListener('change', () => {
      syncCommonEtcGuide();
      if (inCommonCatEl.value === '기타' && inCommonDetailEl) {
        inCommonDetailEl.focus();
      }
    });
  }
  if (inCommonDetailEl) {
    inCommonDetailEl.addEventListener('input', syncCommonEtcGuide);
  }

  // 전체 데이터 초기화
  document.getElementById('btnResetSample').addEventListener('click', () => {
    if (!confirm('입력하신 모든 데이터를 비우고 초기 상태(0원)로 초기화하시겠습니까?')) return;
    state = JSON.parse(JSON.stringify(INITIAL_STATE));
    saveState('🧹 모든 데이터가 초기 상태(0원)로 비워졌습니다.');
    renderAll();
  });

  // 일일 매출 추가 폼 (10kg 미만 / 10kg 이상 구분)
  document.getElementById('addDailySaleForm').addEventListener('submit', e => {
    e.preventDefault();
    const u10 = Number(document.getElementById('inSaleUnder10').value);
    const o10 = Number(document.getElementById('inSaleOver10').value);
    if (!Number.isInteger(u10) || u10 < 0 || !Number.isInteger(o10) || o10 < 0) {
      alert('10kg 미만 및 10kg 이상 주문건수는 0 이상의 정수만 입력할 수 있습니다.');
      return;
    }
    state.dailySales.push({
      id: Date.now(),
      date: document.getElementById('inSaleDate').value,
      channel: document.getElementById('inSaleChannel').value,
      ordersUnder10: u10,
      ordersOver10: o10,
      sales: Number(document.getElementById('inSaleAmount').value) || 0,
      extraCost: Number(document.getElementById('inSaleExtra').value) || 0,
      note: document.getElementById('inSaleNote').value
    });
    saveState('➕ 새로운 일일 매출·주문 내역이 추가 및 저장되었습니다!');
    renderAll();
    document.getElementById('inSaleNote').value = '';
  });

  // 채널 기간별 비용 설정 추가 폼 (박스, 포장비닐, 리플렛 포함)
  document.getElementById('addChannelRateForm').addEventListener('submit', e => {
    e.preventDefault();
    state.channelRates.push({
      id: Date.now(),
      channel: document.getElementById('inRateChannel').value,
      start: document.getElementById('inRateStart').value,
      end: document.getElementById('inRateEnd').value || '',
      feeRate: (Number(document.getElementById('inRateFeePct').value) || 0) / 100,
      shipUnder10: Number(document.getElementById('inRateShipUnder10').value) || 0,
      shipOver10: Number(document.getElementById('inRateShipOver10').value) || 0,
      boxCost: Number(document.getElementById('inRateBox').value) || 0,
      vinylCost: Number(document.getElementById('inRateVinyl').value) || 0,
      leafletCost: Number(document.getElementById('inRateLeaflet').value) || 0,
      note: document.getElementById('inRateNote').value
    });
    saveState('➕ 새로운 판매채널 비용 설정이 추가 및 저장되었습니다!');
    renderAll();
    document.getElementById('inRateNote').value = '';
  });

  // 공통비용 추가 폼 (8개 항목 드롭다운 + 같은 날짜/항목 복수 입력 + '기타' 상세 내역 안내)
  document.getElementById('addCommonCostForm').addEventListener('submit', e => {
    e.preventDefault();
    const dateVal = document.getElementById('inCommonDate').value;
    const catVal = document.getElementById('inCommonCat').value;
    const detailEl = document.getElementById('inCommonDetail');
    const detailVal = (detailEl ? detailEl.value : '').trim();
    const amtVal = Number(document.getElementById('inCommonAmt').value) || 0;
    const noteVal = document.getElementById('inCommonNote').value.trim();

    if (catVal === '기타' && !detailVal) {
      syncCommonEtcGuide();
      alert("비용 항목을 ‘기타’로 선택한 경우 [상세 내역]을 입력해 주세요.");
      if (detailEl) detailEl.focus();
      return;
    }

    state.commonCosts.push({
      id: Date.now() + Math.floor(Math.random() * 10000),
      date: dateVal,
      category: catVal,
      detail: detailVal,
      amount: amtVal,
      note: noteVal
    });
    saveState('➕ 새로운 공통비용 내역이 추가 및 저장되었습니다!');
    if (detailEl) detailEl.value = '';
    document.getElementById('inCommonAmt').value = '';
    document.getElementById('inCommonNote').value = '';
    renderAll();
  });

  // 판매채널 마스터 추가
  document.getElementById('addChannelMasterForm').addEventListener('submit', e => {
    e.preventDefault();
    const name = document.getElementById('inNewChannelName').value.trim();
    if (name && !state.channels.includes(name)) {
      state.channels.push(name);
      saveState(`➕ 새로운 판매채널 [${name}]이(가) 추가 및 저장되었습니다!`);
      renderAll();
    }
    document.getElementById('inNewChannelName').value = '';
  });

  // 화면 너비 변경 시 차트 및 레이아웃 반응형 갱신
  let lastViewportWidth = window.innerWidth;
  let resizeTimer = null;
  window.addEventListener('resize', () => {
    if (window.innerWidth === lastViewportWidth) return;
    lastViewportWidth = window.innerWidth;
    if (resizeTimer) clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      renderAll();
    }, 120);
  });

  renderAll();
});
