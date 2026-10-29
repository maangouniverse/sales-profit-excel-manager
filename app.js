// 초기 기본 데이터 (엑셀 파일과 100% 동일한 데이터셋: 2026.11 ~ 2027.10 기준)
const INITIAL_STATE = {
  channels: ['농라', '스마트스토어', '카톡', '고흥몰', '남도장터', '로컬푸드', '기타'],
  targetMonths: [
    { idx: 1, ym: '2026-11', label: '26년 11월', start: '2026-11-01', end: '2026-11-30', days: 30, target: 12000000 },
    { idx: 2, ym: '2026-12', label: '26년 12월', start: '2026-12-01', end: '2026-12-31', days: 31, target: 15000000 },
    { idx: 3, ym: '2027-01', label: '27년 1월',  start: '2027-01-01', end: '2027-01-31', days: 31, target: 18000000 },
    { idx: 4, ym: '2027-02', label: '27년 2월',  start: '2027-02-01', end: '2027-02-28', days: 28, target: 14000000 },
    { idx: 5, ym: '2027-03', label: '27년 3월',  start: '2027-03-01', end: '2027-03-31', days: 31, target: 10000000 },
    { idx: 6, ym: '2027-04', label: '27년 4월',  start: '2027-04-01', end: '2027-04-30', days: 30, target: 9000000 },
    { idx: 7, ym: '2027-05', label: '27년 5월',  start: '2027-05-01', end: '2027-05-31', days: 31, target: 11000000 },
    { idx: 8, ym: '2027-06', label: '27년 6월',  start: '2027-06-01', end: '2027-06-30', days: 30, target: 8000000 },
    { idx: 9, ym: '2027-07', label: '27년 7월',  start: '2027-07-01', end: '2027-07-31', days: 31, target: 7000000 },
    { idx: 10, ym: '2027-08', label: '27년 8월', start: '2027-08-01', end: '2027-08-31', days: 31, target: 9000000 },
    { idx: 11, ym: '2027-09', label: '27년 9월', start: '2027-09-01', end: '2027-09-30', days: 30, target: 16000000 },
    { idx: 12, ym: '2027-10', label: '27년 10월', start: '2027-10-01', end: '2027-10-31', days: 31, target: 11000000 }
  ],
  channelRates: [
    { id: 1, channel: '농라', start: '2026-11-01', end: '', feeRate: 0.00, shipUnder10: 3500, shipOver10: 4500, boxCost: 650, vinylCost: 150, leafletCost: 100, note: '[예시] 농라 카페 직거래 (수수료 0%, 10kg미만 3,500원 / 이상 4,500원)' },
    { id: 2, channel: '스마트스토어', start: '2026-11-01', end: '2026-11-15', feeRate: 0.055, shipUnder10: 3200, shipOver10: 4200, boxCost: 650, vinylCost: 150, leafletCost: 100, note: '[단가변경 전] 11/15까지 적용 단가 (수수료 5.5%, 미만택배 3,200원/이상 4,200원)' },
    { id: 3, channel: '스마트스토어', start: '2026-11-16', end: '', feeRate: 0.058, shipUnder10: 3500, shipOver10: 4500, boxCost: 750, vinylCost: 150, leafletCost: 100, note: '[단가변경 후] 11/16부터 택배비·박스비 인상 적용 (과거 매출은 이전 단가 유지)' },
    { id: 4, channel: '카톡', start: '2026-11-01', end: '', feeRate: 0.033, shipUnder10: 3500, shipOver10: 4500, boxCost: 650, vinylCost: 150, leafletCost: 100, note: '[예시] 카카오톡 톡채널/선물 주문' },
    { id: 5, channel: '고흥몰', start: '2026-11-01', end: '', feeRate: 0.040, shipUnder10: 3000, shipOver10: 4000, boxCost: 650, vinylCost: 150, leafletCost: 100, note: '[예시] 고흥몰 입점 특산물 판매' },
    { id: 6, channel: '남도장터', start: '2026-11-01', end: '', feeRate: 0.050, shipUnder10: 3200, shipOver10: 4200, boxCost: 650, vinylCost: 150, leafletCost: 100, note: '[예시] 남도장터 온라인몰 판매' },
    { id: 7, channel: '로컬푸드', start: '2026-11-01', end: '', feeRate: 0.100, shipUnder10: 0, shipOver10: 0, boxCost: 400, vinylCost: 120, leafletCost: 80, note: '[예시] 로컬푸드 직매장 납품 (택배비 없음, 소포장 박스·비닐·리플렛만 적용)' },
    { id: 8, channel: '기타', start: '2026-11-01', end: '', feeRate: 0.030, shipUnder10: 3500, shipOver10: 4500, boxCost: 650, vinylCost: 150, leafletCost: 100, note: '[예시] 전화/방문 및 기타 외부채널 기본값' }
  ],
  dailySales: [
    { id: 1, date: '2026-11-02', channel: '농라', ordersUnder10: 24, ordersOver10: 12, sales: 1480000, extraCost: 0, note: '[예시] 11월 첫째주 햇유자·특산물 공동구매' },
    { id: 2, date: '2026-11-04', channel: '고흥몰', ordersUnder10: 30, ordersOver10: 10, sales: 1620000, extraCost: 0, note: '[예시] 고흥몰 가을 수확 기획전' },
    { id: 3, date: '2026-11-06', channel: '카톡', ordersUnder10: 18, ordersOver10: 7, sales: 980000, extraCost: 0, note: '[예시] 단골 고객 카톡 알림톡 주문' },
    { id: 4, date: '2026-11-08', channel: '남도장터', ordersUnder10: 22, ordersOver10: 14, sales: 1540000, extraCost: 0, note: '[예시] 남도장터 주말 할인전' },
    { id: 5, date: '2026-11-11', channel: '로컬푸드', ordersUnder10: 45, ordersOver10: 5, sales: 1150000, extraCost: 0, note: '[예시] 로컬푸드 직매장 출하' },
    { id: 6, date: '2026-11-14', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0, note: '[단가변경 전 확인] 11/15 이전: 수수료 5.5%, 미만 4,100원 / 이상 5,100원 적용' },
    { id: 7, date: '2026-11-17', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0, note: '[단가변경 후 확인] 11/16 이후: 수수료 5.8%, 미만 4,500원 / 이상 5,500원 자동 적용!' },
    { id: 8, date: '2026-11-20', channel: '농라', ordersUnder10: 28, ordersOver10: 15, sales: 1850000, extraCost: 15000, note: '[예시] 대용량 10kg 이상 주문 급증 (완충재 추가비용 반영)' },
    { id: 9, date: '2026-11-22', channel: '기타', ordersUnder10: 12, ordersOver10: 8, sales: 750000, extraCost: 0, note: '[예시] 전화 주문 단체 발송' },
    { id: 10, date: '2026-11-25', channel: '기타', ordersUnder10: 15, ordersOver10: 10, sales: 95000, extraCost: 35000, note: '[적자 구분 예시] 파격 샘플행사 및 파손 재발송비로 인한 일시 적자' },
    { id: 11, date: '2026-11-27', channel: '고흥몰', ordersUnder10: 25, ordersOver10: 12, sales: 1520000, extraCost: 0, note: '[예시] 고흥몰 베스트 상품전' },
    { id: 12, date: '2026-11-29', channel: '카톡', ordersUnder10: 20, ordersOver10: 8, sales: 1120000, extraCost: 0, note: '[예시] 월말 카톡 재구매 주문' },
    { id: 13, date: '2026-12-05', channel: '스마트스토어', ordersUnder10: 45, ordersOver10: 20, sales: 2850000, extraCost: 0, note: '[예시] 연말 선물세트 주문' },
    { id: 14, date: '2026-12-12', channel: '농라', ordersUnder10: 50, ordersOver10: 25, sales: 3400000, extraCost: 0, note: '[예시] 겨울철 인기 특산물 공구' },
    { id: 15, date: '2026-12-20', channel: '남도장터', ordersUnder10: 38, ordersOver10: 18, sales: 2420000, extraCost: 0, note: '[예시] 연말 감사 기획전' },
    { id: 16, date: '2026-12-28', channel: '고흥몰', ordersUnder10: 40, ordersOver10: 15, sales: 2380000, extraCost: 0, note: '[예시] 해맞이 특산품 행사' },
    { id: 17, date: '2027-01-08', channel: '스마트스토어', ordersUnder10: 55, ordersOver10: 30, sales: 3950000, extraCost: 0, note: '[예시] 설 명절 사전예약 1차' },
    { id: 18, date: '2027-01-16', channel: '농라', ordersUnder10: 60, ordersOver10: 35, sales: 4500000, extraCost: 20000, note: '[예시] 설 명절 대량주문' },
    { id: 19, date: '2027-01-24', channel: '카톡', ordersUnder10: 42, ordersOver10: 22, sales: 2980000, extraCost: 0, note: '[예시] 카톡 명절 선물세트' },
    { id: 20, date: '2027-01-29', channel: '로컬푸드', ordersUnder10: 70, ordersOver10: 15, sales: 2150000, extraCost: 0, note: '[예시] 명절 직매장 특판' },
    { id: 21, date: '2027-02-06', channel: '고흥몰', ordersUnder10: 48, ordersOver10: 20, sales: 3100000, extraCost: 0, note: '[예시] 정월대보름 기획전' },
    { id: 22, date: '2027-02-15', channel: '남도장터', ordersUnder10: 35, ordersOver10: 15, sales: 2250000, extraCost: 0, note: '[예시] 2월 정기 프로모션' },
    { id: 23, date: '2027-02-23', channel: '스마트스토어', ordersUnder10: 40, ordersOver10: 12, sales: 2320000, extraCost: 0, note: '[예시] 봄맞이 사전 특가' }
  ],
  commonCosts: [
    { id: 1, date: '2026-11-25', category: '월 광고선전비', amount: 300000, note: '[예시] 11월 키워드 및 배너 광고비' },
    { id: 2, date: '2026-11-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 11월 저온창고 임대료' },
    { id: 3, date: '2026-11-28', category: '공과금/통신/프로그램', amount: 100000, note: '[예시] 11월 전기료 및 송장프로그램' },
    { id: 4, date: '2026-12-25', category: '월 광고선전비', amount: 350000, note: '[예시] 12월 연말 선물세트 광고비' },
    { id: 5, date: '2026-12-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 12월 저온창고 임대료' },
    { id: 6, date: '2027-01-25', category: '월 광고선전비', amount: 400000, note: '[예시] 1월 설 명절 집중 광고비' },
    { id: 7, date: '2027-01-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 1월 저온창고 임대료' },
    { id: 8, date: '2027-02-25', category: '월 광고선전비', amount: 250000, note: '[예시] 2월 기본 검색광고비' },
    { id: 9, date: '2027-02-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 2월 저온창고 임대료' }
  ]
};

const STORAGE_KEY = 'sales_profit_manager_state_v2';
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
    if (saved) return JSON.parse(saved);
  } catch (e) {
    console.warn(e);
  }
  return JSON.parse(JSON.stringify(INITIAL_STATE));
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (e) {
    console.warn(e);
  }
}

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

// 실적자동계산1.png 스타일 날짜 표시: "26년 11월 2일 (월)" + 토/일 요일 클래스
function formatKoreanDateWithDay(yyyyMmDd) {
  if (!yyyyMmDd) return { text: '', cls: '' };
  const [y, m, d] = yyyyMmDd.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const dayIdx = dt.getUTCDay();
  const shortY = String(y).slice(2);
  const text = `${shortY}년 ${m}월 ${d}일 (${days[dayIdx]})`;
  const cls = dayIdx === 6 ? 'date-sat' : dayIdx === 0 ? 'date-sun' : '';
  return { text, cls };
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
  renderDashboard(evaluatedSales, periodRows, periodTarget, actualSales, totalVarCost, netProfit, finalProfit);

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

function renderDashboard(evaluatedSales, periodRows, periodTarget, actualSales, totalVarCost, netProfit, finalProfit) {
  // 1) 상단 히어로 막대 차트 (매출성과.png 우측 차트 스타일: 단위 천원)
  updateChart('chartTargetVsActual', {
    type: 'bar',
    data: {
      labels: ['목표매출', '총 매출액', '총 순이익', '최종순이익(공통비후)'],
      datasets: [{
        label: '단위: 천원',
        data: [
          Math.round(periodTarget / 1000),
          Math.round(actualSales / 1000),
          Math.round(netProfit / 1000),
          Math.round(finalProfit / 1000)
        ],
        backgroundColor: ['#92add0', '#2b4cbe', '#c5d185', '#5c7cfa'],
        barThickness: 38,
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

  const palette = ['#92add0', '#c1d4f5', '#112240', '#e6cc55', '#5c7cfa', '#c5d185', '#d0d7de'];

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
        legend: { position: 'right', labels: { boxWidth: 10, font: { size: 12 } } },
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
        legend: { position: 'right', labels: { boxWidth: 10, font: { size: 12 } } },
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

  // 3) 월별 손익현황 (26년 11월 ~ 27년 10월, 실적자동계산2.png 스타일)
  const mLabels = [];
  const mSales = [];
  const mCosts = [];
  const mProfits = [];

  const mTbody = document.getElementById('dashboardMonthlyTbody');
  const mTfoot = document.getElementById('dashboardMonthlyTfoot');
  mTbody.innerHTML = '';

  let sumTarget = 0, sumSales = 0, sumCost = 0, sumProfit = 0, sumOrd = 0;

  state.targetMonths.forEach(tm => {
    const rows = evaluatedSales.filter(r => r.isValid && r.date >= tm.start && r.date <= tm.end);
    const s = rows.reduce((a, b) => a + b.sales, 0);
    const c = rows.reduce((a, b) => a + b.totalCost, 0);
    const p = s - c;
    const ord = rows.reduce((a, b) => a + b.totalOrders, 0);
    const ach = tm.target > 0 ? s / tm.target : 0;
    const mar = s > 0 ? p / s : 0;

    mLabels.push(tm.label);
    mSales.push(s);
    mCosts.push(c);
    mProfits.push(p);

    sumTarget += tm.target;
    sumSales += s;
    sumCost += c;
    sumProfit += p;
    sumOrd += ord;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${tm.label}</strong></td>
      <td class="num">${fmtNum(tm.target)}</td>
      <td class="num"><strong>${s > 0 ? fmtNum(s) : ''}</strong></td>
      <td class="num">${s > 0 ? fmtPct(ach, 1) : '0.0%'}</td>
      <td class="num">${c > 0 ? fmtNum(c) : ''}</td>
      <td class="num ${p < 0 ? 'cell-negative' : p > 0 ? 'profit-pos' : ''}"><strong>${s > 0 || c > 0 ? fmtNum(p) : ''}</strong></td>
      <td class="num">${s > 0 ? fmtPct(mar, 2) : '0.00%'}</td>
      <td class="num">${ord > 0 ? fmtNum(ord) + '건' : ''}</td>
    `;
    mTbody.appendChild(tr);
  });

  mTfoot.innerHTML = `
    <tr>
      <td>12개월 합계</td>
      <td class="num">${fmtNum(sumTarget)}</td>
      <td class="num">${fmtNum(sumSales)}</td>
      <td class="num">${fmtPct(sumTarget > 0 ? sumSales / sumTarget : 0, 1)}</td>
      <td class="num">${fmtNum(sumCost)}</td>
      <td class="num ${sumProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sumProfit)}</td>
      <td class="num">${fmtPct(sumSales > 0 ? sumProfit / sumSales : 0, 2)}</td>
      <td class="num">${fmtNum(sumOrd)}건</td>
    </tr>
  `;

  updateChart('chartMonthlyTrend', {
    type: 'bar',
    data: {
      labels: mLabels,
      datasets: [
        { label: '총 매출', data: mSales, backgroundColor: '#2b4cbe', borderRadius: 3 },
        { label: '총 지출(비용)', data: mCosts, backgroundColor: '#92add0', borderRadius: 3 },
        { label: '순이익', data: mProfits, backgroundColor: '#c5d185', borderRadius: 3 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } },
      scales: { x: { grid: { display: false } } }
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

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="${dFormatted.cls}">${dFormatted.text}</td>
      <td><strong>${r.channel}</strong></td>
      <td class="num">${fmtNum(r.ordersUnder10)}</td>
      <td class="num">${fmtNum(r.ordersOver10)}</td>
      <td class="num"><strong>${fmtNum(r.totalOrders)}</strong></td>
      <td class="num"><strong>${fmtNum(r.sales)}</strong></td>
      <td class="num">${r.extraCost > 0 ? fmtNum(r.extraCost) : ''}</td>
      <td style="text-align:left;color:#555;">${r.note || ''}</td>
      <td class="${errCls}">${r.status}</td>
      <td class="num ${errCls}">${r.isValid ? fmtPct(r.feeRate, 2) : '비용 설정 필요'}</td>
      <td class="num ${errCls}">${r.isValid ? fmtNum(r.feeAmt) : '비용 설정 필요'}</td>
      <td class="num ${errCls}">${r.isValid ? fmtNum(r.unitUnder10) : '비용 설정 필요'}</td>
      <td class="num ${errCls}">${r.isValid ? fmtNum(r.unitOver10) : '비용 설정 필요'}</td>
      <td class="num ${errCls}">${r.isValid ? fmtNum(r.orderCost) : '비용 설정 필요'}</td>
      <td class="num ${errCls}"><strong>${r.isValid ? fmtNum(r.totalCost) : '비용 설정 필요'}</strong></td>
      <td class="num ${profitCls}">${r.isValid ? fmtNum(r.netProfit) : '비용 설정 필요'}</td>
      <td class="num ${errCls}">${r.isValid ? fmtPct(r.margin, 1) : '비용 설정 필요'}</td>
      <td><button class="btn-danger-sm" onclick="deleteDailyRow(${r.id})">삭제</button></td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="2">합계 (${rows.length}건)</td>
      <td class="num">${fmtNum(sumU10)}</td>
      <td class="num">${fmtNum(sumO10)}</td>
      <td class="num">${fmtNum(sumOrd)}</td>
      <td class="num">${fmtNum(sumSales)}</td>
      <td class="num">${fmtNum(sumExtra)}</td>
      <td></td>
      <td>정상 합산</td>
      <td class="num">-</td>
      <td class="num">${fmtNum(sumFee)}</td>
      <td class="num">-</td>
      <td class="num">-</td>
      <td class="num">${fmtNum(sumOrdCost)}</td>
      <td class="num">${fmtNum(sumCost)}</td>
      <td class="num ${sumProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sumProfit)}</td>
      <td class="num">${fmtPct(sumSales > 0 ? sumProfit / sumSales : 0, 1)}</td>
      <td></td>
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

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><strong>${r.channel}</strong></td>
      <td><input type="date" class="table-inline-date date-clickable" value="${r.start}" onchange="updateRateField(${r.id}, 'start', this.value)" /></td>
      <td><input type="date" class="table-inline-date date-clickable" value="${r.end || ''}" onchange="updateRateField(${r.id}, 'end', this.value)" /></td>
      <td class="num"><input type="number" class="table-inline-input" step="0.1" value="${(r.feeRate * 100).toFixed(2)}" onchange="updateRateField(${r.id}, 'feeRate', Number(this.value)/100)" />%</td>
      <td class="num"><input type="number" class="table-inline-input" step="100" value="${r.shipUnder10}" onchange="updateRateField(${r.id}, 'shipUnder10', Number(this.value))" /></td>
      <td class="num"><input type="number" class="table-inline-input" step="100" value="${r.shipOver10}" onchange="updateRateField(${r.id}, 'shipOver10', Number(this.value))" /></td>
      <td class="num"><input type="number" class="table-inline-input" step="10" value="${r.boxCost}" onchange="updateRateField(${r.id}, 'boxCost', Number(this.value))" /></td>
      <td class="num"><input type="number" class="table-inline-input" step="10" value="${r.vinylCost}" onchange="updateRateField(${r.id}, 'vinylCost', Number(this.value))" /></td>
      <td class="num"><input type="number" class="table-inline-input" step="10" value="${r.leafletCost}" onchange="updateRateField(${r.id}, 'leafletCost', Number(this.value))" /></td>
      <td class="num"><strong>${fmtWon(sumUnder10)}</strong></td>
      <td class="num"><strong>${fmtWon(sumOver10)}</strong></td>
      <td class="${isErr ? 'cell-error' : ''}">${status}</td>
      <td style="text-align:left;color:#555;">${r.note || ''}</td>
      <td><button class="btn-danger-sm" onclick="deleteRateRow(${r.id})">삭제</button></td>
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
      <td>${i + 1}</td>
      <td><strong>${st.name}</strong></td>
      <td class="num">${fmtNum(st.u10)}</td>
      <td class="num">${fmtNum(st.o10)}</td>
      <td class="num"><strong>${fmtNum(st.orders)}</strong></td>
      <td class="num"><strong>${fmtNum(st.sales)}</strong></td>
      <td class="num">${fmtPct(st.share, 1)}</td>
      <td class="num">${fmtNum(st.fee)}</td>
      <td class="num">${fmtNum(st.orderCost)}</td>
      <td class="num">${fmtNum(st.extra)}</td>
      <td class="num">${fmtNum(st.totalCost)}</td>
      <td class="num ${neg1}"><strong>${fmtNum(st.netProfit)}</strong></td>
      <td class="num ${st.netProfit < 0 ? 'cell-negative' : ''}">${fmtPct(st.margin, 1)}</td>
      <td class="num">${fmtNum(st.allocCommon)}</td>
      <td class="num ${neg2}"><strong>${fmtNum(st.finalProfit)}</strong></td>
      <td class="num ${neg2}">${fmtPct(st.finalMargin, 1)}</td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="2">전체 ${channelStats.length}개 판매채널 합계</td>
      <td class="num">${fmtNum(sU10)}</td>
      <td class="num">${fmtNum(sO10)}</td>
      <td class="num">${fmtNum(sOrd)}</td>
      <td class="num">${fmtNum(sSales)}</td>
      <td class="num">100.0%</td>
      <td class="num">${fmtNum(sFee)}</td>
      <td class="num">${fmtNum(sOrdCost)}</td>
      <td class="num">${fmtNum(sExtra)}</td>
      <td class="num">${fmtNum(sCost)}</td>
      <td class="num ${sProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sProfit)}</td>
      <td class="num">${fmtPct(sSales > 0 ? sProfit / sSales : 0, 1)}</td>
      <td class="num">${fmtNum(sAlloc)}</td>
      <td class="num ${sFinal < 0 ? 'cell-negative' : ''}">${fmtNum(sFinal)}</td>
      <td class="num">${fmtPct(sSales > 0 ? sFinal / sSales : 0, 1)}</td>
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

function renderTargetsAndCommonTab(start, end, totalAnnualTarget, periodTarget, periodCommonCost) {
  document.getElementById('tab5TotalTargetSum').textContent = fmtWon(totalAnnualTarget);
  document.getElementById('tab5AppliedTargetSum').textContent = fmtWon(periodTarget);
  document.getElementById('tab5PeriodCommonSum').textContent = fmtWon(periodCommonCost);

  const tTbody = document.querySelector('#monthlyTargetsTable tbody');
  const tTfoot = document.querySelector('#monthlyTargetsTable tfoot');
  tTbody.innerHTML = '';

  state.targetMonths.forEach(tm => {
    const daily = Math.round(tm.target / tm.days);
    const applied = calcProRatedTargetForMonth(tm, start, end);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${tm.label}</strong></td>
      <td style="color:#666;">${tm.start} ~ ${tm.end}</td>
      <td>${tm.days}일</td>
      <td class="num">
        <input type="number" class="table-inline-input" style="width:130px" step="100000" value="${tm.target}" onchange="updateMonthlyTarget('${tm.ym}', Number(this.value))" />
      </td>
      <td class="num">${fmtWon(daily)}</td>
      <td class="num"><strong>${fmtWon(applied)}</strong></td>
    `;
    tTbody.appendChild(tr);
  });

  tTfoot.innerHTML = `
    <tr>
      <td colspan="2">12개월 목표매출 총합계 (2026.11 ~ 2027.10)</td>
      <td>365일</td>
      <td class="num">${fmtWon(totalAnnualTarget)}</td>
      <td class="num">${fmtWon(Math.round(totalAnnualTarget / 365))}/일</td>
      <td class="num">${fmtWon(periodTarget)}</td>
    </tr>
  `;

  // 공통비용 표
  const cTbody = document.querySelector('#commonCostsTable tbody');
  const cTfoot = document.querySelector('#commonCostsTable tfoot');
  cTbody.innerHTML = '';

  let sumAllCommon = 0;
  state.commonCosts.forEach(c => {
    const inPeriod = c.date >= start && c.date <= end;
    sumAllCommon += c.amount;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${c.date}</td>
      <td><strong>${c.category}</strong></td>
      <td class="num">${fmtWon(c.amount)}</td>
      <td>${inPeriod ? '✅ 포함' : '제외'}</td>
      <td style="text-align:left;color:#555;">${c.note || ''}</td>
      <td><button class="btn-danger-sm" onclick="deleteCommonCost(${c.id})">삭제</button></td>
    `;
    cTbody.appendChild(tr);
  });

  cTfoot.innerHTML = `
    <tr>
      <td colspan="2">공통비용 합계</td>
      <td class="num">전체: ${fmtWon(sumAllCommon)}</td>
      <td colspan="3">현재 조회기간 반영액: ${fmtWon(periodCommonCost)}</td>
    </tr>
  `;
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
  saveState();
  renderAll();
};

window.deleteRateRow = function(id) {
  state.channelRates = state.channelRates.filter(r => r.id !== id);
  saveState();
  renderAll();
};

window.updateRateField = function(id, field, val) {
  const target = state.channelRates.find(r => r.id === id);
  if (target) {
    target[field] = val;
    saveState();
    renderAll();
  }
};

window.updateMonthlyTarget = function(ym, val) {
  const t = state.targetMonths.find(x => x.ym === ym);
  if (t) {
    t.target = Math.max(0, val || 0);
    saveState();
    renderAll();
  }
};

window.deleteCommonCost = function(id) {
  state.commonCosts = state.commonCosts.filter(c => c.id !== id);
  saveState();
  renderAll();
};

document.addEventListener('DOMContentLoaded', () => {
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

  ['filterYm', 'filterStartDate', 'filterEndDate', 'chkFilterDailyByPeriod', 'analysisSortBy', 'analysisAllocRule'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', renderAll);
  });

  // 예시 데이터 초기화
  document.getElementById('btnResetSample').addEventListener('click', () => {
    state = JSON.parse(JSON.stringify(INITIAL_STATE));
    saveState();
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
    saveState();
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
    saveState();
    renderAll();
    document.getElementById('inRateNote').value = '';
  });

  // 공통비용 추가 폼
  document.getElementById('addCommonCostForm').addEventListener('submit', e => {
    e.preventDefault();
    state.commonCosts.push({
      id: Date.now(),
      date: document.getElementById('inCommonDate').value,
      category: document.getElementById('inCommonCat').value,
      amount: Number(document.getElementById('inCommonAmt').value) || 0,
      note: document.getElementById('inCommonNote').value
    });
    saveState();
    renderAll();
    document.getElementById('inCommonCat').value = '';
    document.getElementById('inCommonAmt').value = '';
    document.getElementById('inCommonNote').value = '';
  });

  // 판매채널 마스터 추가
  document.getElementById('addChannelMasterForm').addEventListener('submit', e => {
    e.preventDefault();
    const name = document.getElementById('inNewChannelName').value.trim();
    if (name && !state.channels.includes(name)) {
      state.channels.push(name);
      saveState();
      renderAll();
    }
    document.getElementById('inNewChannelName').value = '';
  });

  renderAll();
});
