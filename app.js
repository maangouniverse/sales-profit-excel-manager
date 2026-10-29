// 초기 기본 데이터 (엑셀 파일과 100% 동일한 데이터셋)
const INITIAL_STATE = {
  channels: ['농라', '스마트스토어', '카톡', '고흥몰', '남도장터', '로컬푸드', '기타'],
  clients: ['청정영농조합', '고흥유자농원', '남도바다수산', '해피푸드유통', '초록들녘농산', '바른먹거리장터', '우리동네직판장', '자체직영(소비자)'],
  channelRates: [
    { id: 1, channel: '농라', start: '2026-01-01', end: '', feeRate: 0.00, shipFee: 3500, packFee: 800, etcFee: 200, note: '[예시] 농라 카페 직거래 (수수료 0%, 택배·포장비 적용)' },
    { id: 2, channel: '스마트스토어', start: '2026-01-01', end: '2026-04-15', feeRate: 0.055, shipFee: 3000, packFee: 1000, etcFee: 110, note: '[예시] 4/15까지 적용 단가 (기타비용에 배송비수수료 110원 반영)' },
    { id: 3, channel: '스마트스토어', start: '2026-04-16', end: '', feeRate: 0.058, shipFee: 3500, packFee: 1200, etcFee: 130, note: '[예시] 4/16부터 택배·포장비 인상 및 수수료율 변경 적용' },
    { id: 4, channel: '카톡', start: '2026-01-01', end: '', feeRate: 0.033, shipFee: 3500, packFee: 900, etcFee: 100, note: '[예시] 카카오톡 스토어/톡채널 주문' },
    { id: 5, channel: '고흥몰', start: '2026-01-01', end: '', feeRate: 0.040, shipFee: 3000, packFee: 1000, etcFee: 0, note: '[예시] 고흥몰 입점 특산물 판매' },
    { id: 6, channel: '남도장터', start: '2026-01-01', end: '', feeRate: 0.050, shipFee: 3200, packFee: 1000, etcFee: 0, note: '[예시] 남도장터 온라인몰 판매' },
    { id: 7, channel: '로컬푸드', start: '2026-01-01', end: '', feeRate: 0.100, shipFee: 0, packFee: 500, etcFee: 200, note: '[예시] 로컬푸드 직매장 납품 (개별 택배비 없음)' },
    { id: 8, channel: '기타', start: '2026-01-01', end: '', feeRate: 0.030, shipFee: 3500, packFee: 800, etcFee: 0, note: '[예시] 전화주문/외부 제휴몰 기본 단가' }
  ],
  dailySales: [
    { id: 1, date: '2026-01-08', client: '청정영농조합', channel: '스마트스토어', orders: 25, sales: 1125000, cogs: 580000, extraCost: 0, note: '[예시] 신년 선물세트 기획전' },
    { id: 2, date: '2026-01-15', client: '고흥유자농원', channel: '고흥몰', orders: 30, sales: 1350000, cogs: 690000, extraCost: 0, note: '[예시] 유자청 베스트 상품' },
    { id: 3, date: '2026-01-22', client: '남도바다수산', channel: '농라', orders: 40, sales: 1920000, cogs: 1150000, extraCost: 15000, note: '[예시] 제철 수산물 공동구매 (아이스팩 추가)' },
    { id: 4, date: '2026-01-28', client: '우리동네직판장', channel: '로컬푸드', orders: 55, sales: 990000, cogs: 540000, extraCost: 0, note: '[예시] 직매장 주말 납품' },
    { id: 5, date: '2026-02-05', client: '해피푸드유통', channel: '남도장터', orders: 35, sales: 1680000, cogs: 920000, extraCost: 0, note: '[예시] 설 명절 특선 세트' },
    { id: 6, date: '2026-02-11', client: '고흥유자농원', channel: '카톡', orders: 28, sales: 1120000, cogs: 560000, extraCost: 0, note: '[예시] 카톡 선물하기 주문' },
    { id: 7, date: '2026-02-19', client: '초록들녘농산', channel: '스마트스토어', orders: 22, sales: 880000, cogs: 460000, extraCost: 6000, note: '[예시] 단순변심 반품 왕복택배비 1건 반영' },
    { id: 8, date: '2026-02-25', client: '자체직영(소비자)', channel: '기타', orders: 18, sales: 720000, cogs: 350000, extraCost: 0, note: '[예시] 전화 주문 단골 고객' },
    { id: 9, date: '2026-03-06', client: '청정영농조합', channel: '농라', orders: 32, sales: 1440000, cogs: 780000, extraCost: 0, note: '[예시] 봄맞이 산지직송' },
    { id: 10, date: '2026-03-14', client: '남도바다수산', channel: '남도장터', orders: 26, sales: 1300000, cogs: 740000, extraCost: 0, note: '[예시] 건어물 기획전' },
    { id: 11, date: '2026-03-21', client: '바른먹거리장터', channel: '고흥몰', orders: 20, sales: 760000, cogs: 490000, extraCost: 0, note: '[예시] 친환경 곡물 세트' },
    { id: 12, date: '2026-03-29', client: '해피푸드유통', channel: '스마트스토어', orders: 38, sales: 1710000, cogs: 890000, extraCost: 0, note: '[예시] 월말 타임특가' },
    { id: 13, date: '2026-04-02', client: '청정영농조합', channel: '농라', orders: 30, sales: 1500000, cogs: 820000, extraCost: 0, note: '[예시] 4월 첫째주 공동구매' },
    { id: 14, date: '2026-04-02', client: '남도바다수산', channel: '농라', orders: 20, sales: 1100000, cogs: 640000, extraCost: 0, note: '[예시] 같은 날짜·같은 채널 다른 거래처 분리 입력 예시' },
    { id: 15, date: '2026-04-05', client: '고흥유자농원', channel: '고흥몰', orders: 35, sales: 1575000, cogs: 790000, extraCost: 0, note: '[예시] 고흥몰 봄 정기행사' },
    { id: 16, date: '2026-04-08', client: '해피푸드유통', channel: '카톡', orders: 25, sales: 1050000, cogs: 530000, extraCost: 0, note: '[예시] 톡딜 프로모션' },
    { id: 17, date: '2026-04-11', client: '초록들녘농산', channel: '남도장터', orders: 28, sales: 1260000, cogs: 680000, extraCost: 0, note: '[예시] 남도장터 주말특가' },
    { id: 18, date: '2026-04-14', client: '고흥유자농원', channel: '스마트스토어', orders: 20, sales: 800000, cogs: 400000, extraCost: 0, note: '[단가변경 전 확인] 4/15 이전: 수수료 5.5%, 건당비용 4,110원 적용' },
    { id: 19, date: '2026-04-17', client: '고흥유자농원', channel: '스마트스토어', orders: 20, sales: 800000, cogs: 400000, extraCost: 0, note: '[단가변경 후 확인] 4/16 이후: 수수료 5.8%, 건당비용 4,830원 자동 적용!' },
    { id: 20, date: '2026-04-19', client: '우리동네직판장', channel: '로컬푸드', orders: 60, sales: 1080000, cogs: 590000, extraCost: 0, note: '[예시] 로컬푸드 직매장 주말 완판' },
    { id: 21, date: '2026-04-22', client: '바른먹거리장터', channel: '스마트스토어', orders: 25, sales: 450000, cogs: 360000, extraCost: 30000, note: '[적자 구분 예시] 파격 할인특가 및 반품비로 인한 적자 발생' },
    { id: 22, date: '2026-04-25', client: '자체직영(소비자)', channel: '기타', orders: 15, sales: 675000, cogs: 320000, extraCost: 0, note: '[예시] 기업 단체선물 주문' },
    { id: 23, date: '2026-04-27', client: '청정영농조합', channel: '카톡', orders: 22, sales: 990000, cogs: 510000, extraCost: 0, note: '[예시] 재구매 알림톡 주문' },
    { id: 24, date: '2026-04-29', client: '남도바다수산', channel: '고흥몰', orders: 18, sales: 900000, cogs: 490000, extraCost: 7000, note: '[예시] 냉동포장 드라이아이스 추가비용 반영' }
  ],
  monthlyTargets: [
    { month: 1, target: 5000000 },
    { month: 2, target: 4500000 },
    { month: 3, target: 5000000 },
    { month: 4, target: 11000000 },
    { month: 5, target: 12000000 },
    { month: 6, target: 10000000 },
    { month: 7, target: 9000000 },
    { month: 8, target: 10000000 },
    { month: 9, target: 15000000 },
    { month: 10, target: 11000000 },
    { month: 11, target: 12000000 },
    { month: 12, target: 14000000 }
  ],
  commonCosts: [
    { id: 1, date: '2026-01-25', category: '월 광고선전비', amount: 250000, note: '[예시] 1월 검색광고 및 카페 배너비' },
    { id: 2, date: '2026-01-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 1월 저온창고 월세' },
    { id: 3, date: '2026-02-25', category: '월 광고선전비', amount: 300000, note: '[예시] 2월 설 명절 키워드 광고비' },
    { id: 4, date: '2026-02-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 2월 저온창고 월세' },
    { id: 5, date: '2026-03-25', category: '월 광고선전비', amount: 250000, note: '[예시] 3월 쇼핑검색 광고비' },
    { id: 6, date: '2026-03-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 3월 저온창고 월세' },
    { id: 7, date: '2026-04-10', category: '월 광고선전비', amount: 350000, note: '[예시] 4월 봄 프로모션 광고비' },
    { id: 8, date: '2026-04-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 4월 저온창고 월세' },
    { id: 9, date: '2026-04-28', category: '프로그램/통신/공과금', amount: 100000, note: '[예시] 4월 송장출력기 및 인터넷/전기료' }
  ]
};

const STORAGE_KEY = 'sales_profit_manager_state_v1';
let state = loadState();
let charts = {};

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

// 현재 적용된 시작일/종료일 계산
function getActivePeriod() {
  const mode = document.getElementById('filterMode').value;
  if (mode === 'custom') {
    const s = document.getElementById('filterStartDate').value || '2026-01-01';
    const e = document.getElementById('filterEndDate').value || '2026-12-31';
    return { start: s, end: e, year: Number(s.slice(0, 4)) };
  } else {
    const year = Number(document.getElementById('filterYear').value) || 2026;
    const mVal = document.getElementById('filterMonth').value;
    if (mVal === 'all') {
      return { start: `${year}-01-01`, end: `${year}-12-31`, year };
    }
    const m = Number(mVal);
    const days = new Date(Date.UTC(year, m, 0)).getUTCDate();
    const start = `${year}-${String(m).padStart(2, '0')}-01`;
    const end = `${year}-${String(m).padStart(2, '0')}-${String(days).padStart(2, '0')}`;
    return { start, end, year };
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

// 일일 매출 행 비용 및 순이익 계산
function evaluateDailyRow(row, allRates) {
  const matched = allRates.filter(r => {
    if (r.channel !== row.channel) return false;
    const s = r.start;
    const e = r.end || '2099-12-31';
    return row.date >= s && row.date <= e;
  });

  if (matched.length === 0) {
    return {
      ...row,
      status: '비용 설정 필요',
      isValid: false
    };
  }
  if (matched.length > 1) {
    return {
      ...row,
      status: '기간 중복 확인',
      isValid: false
    };
  }

  const rate = matched[0];
  const feeRate = Number(rate.feeRate) || 0;
  const feeAmt = Math.round(row.sales * feeRate);
  const unitCost = (Number(rate.shipFee) || 0) + (Number(rate.packFee) || 0) + (Number(rate.etcFee) || 0);
  const orderCost = row.orders * unitCost;
  const totalCost = row.cogs + feeAmt + orderCost + row.extraCost;
  const netProfit = row.sales - totalCost;
  const margin = row.sales > 0 ? netProfit / row.sales : 0;

  return {
    ...row,
    status: '정상',
    isValid: true,
    feeRate,
    feeAmt,
    unitCost,
    orderCost,
    totalCost,
    netProfit,
    margin
  };
}

// 임의 조회기간에 대한 월별 일할 목표매출 계산
function calcProRatedTarget(year, month, periodStart, periodEnd, monthlyTarget) {
  const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
  const mStart = `${year}-${String(month).padStart(2, '0')}-01`;
  const mEnd = `${year}-${String(month).padStart(2, '0')}-${String(daysInMonth).padStart(2, '0')}`;

  if (periodEnd < mStart || periodStart > mEnd) return 0;
  if (periodStart <= mStart && periodEnd >= mEnd) return monthlyTarget;

  const overlapStart = periodStart > mStart ? periodStart : mStart;
  const overlapEnd = periodEnd < mEnd ? periodEnd : mEnd;
  const d1 = new Date(overlapStart + 'T00:00:00Z');
  const d2 = new Date(overlapEnd + 'T00:00:00Z');
  const overlapDays = Math.round((d2 - d1) / 86400000) + 1;
  return Math.round(monthlyTarget * (overlapDays / daysInMonth));
}

// 전체 UI 렌더링
function renderAll() {
  const { start, end, year } = getActivePeriod();
  const d1 = new Date(start + 'T00:00:00Z');
  const d2 = new Date(end + 'T00:00:00Z');
  const diffDays = Math.max(1, Math.round((d2 - d1) / 86400000) + 1);
  document.getElementById('appliedPeriodBadge').textContent = `${start} ~ ${end} (${diffDays}일간)`;

  // 1. 모든 일일 매출 행 평가
  const evaluatedSales = state.dailySales.map(r => evaluateDailyRow(r, state.channelRates));
  const missingRateCount = evaluatedSales.filter(r => r.status === '비용 설정 필요').length;
  const overlapRateCount = state.channelRates.filter(r => checkRateOverlap(r, state.channelRates) !== '정상').length;

  // 상단 데이터 점검 배너 업데이트
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

  // 조회기간 목표매출 합산
  let periodTarget = 0;
  state.monthlyTargets.forEach(t => {
    periodTarget += calcProRatedTarget(year, t.month, start, end, t.target);
  });

  const actualSales = periodRows.reduce((a, b) => a + b.sales, 0);
  const totalOrders = periodRows.reduce((a, b) => a + b.orders, 0);
  const totalCogs = periodRows.reduce((a, b) => a + b.cogs, 0);
  const totalFee = periodRows.reduce((a, b) => a + b.feeAmt, 0);
  const totalOrderCost = periodRows.reduce((a, b) => a + b.orderCost, 0);
  const totalExtra = periodRows.reduce((a, b) => a + b.extraCost, 0);
  const totalVarCost = periodRows.reduce((a, b) => a + b.totalCost, 0);
  const netProfit = actualSales - totalVarCost;
  const netMargin = actualSales > 0 ? netProfit / actualSales : 0;
  const achieveRate = periodTarget > 0 ? actualSales / periodTarget : 0;

  const periodCommonCost = state.commonCosts
    .filter(c => c.date >= start && c.date <= end)
    .reduce((a, b) => a + b.amount, 0);

  const finalProfit = netProfit - periodCommonCost;
  const finalMargin = actualSales > 0 ? finalProfit / actualSales : 0;

  // KPI 카드 업데이트
  document.getElementById('kpiTargetSales').textContent = fmtWon(periodTarget);
  document.getElementById('kpiActualSales').textContent = fmtWon(actualSales);
  document.getElementById('kpiOrderCountSub').textContent = `총 주문 ${fmtNum(totalOrders)}건`;
  document.getElementById('kpiAchieveRate').textContent = fmtPct(achieveRate, 1);
  document.getElementById('kpiAchieveBar').style.width = `${Math.min(100, Math.max(0, achieveRate * 100))}%`;
  document.getElementById('kpiOrders').textContent = `${fmtNum(totalOrders)}건`;
  document.getElementById('kpiAvgOrderVal').textContent = `건당 평균 매출 ${totalOrders > 0 ? fmtWon(actualSales / totalOrders) : '0원'}`;
  document.getElementById('kpiTotalVarCost').textContent = fmtWon(totalVarCost);
  document.getElementById('kpiCostBreakdownSub').textContent = `원가 ${fmtNum(totalCogs)} + 수수료 ${fmtNum(totalFee)} + 건수비 ${fmtNum(totalOrderCost)}`;
  document.getElementById('kpiNetProfit').textContent = fmtWon(netProfit);
  document.getElementById('kpiNetMargin').textContent = `순이익률 ${fmtPct(netMargin, 1)} (합계 기준)`;
  document.getElementById('kpiNetProfitCard').className = `kpi-card ${netProfit < 0 ? 'accent-red' : 'accent-green'}`;
  document.getElementById('kpiCommonCost').textContent = fmtWon(periodCommonCost);
  document.getElementById('kpiFinalProfit').textContent = fmtWon(finalProfit);
  document.getElementById('kpiFinalMargin').textContent = `최종 순이익률 ${fmtPct(finalMargin, 1)}`;
  document.getElementById('kpiFinalProfitCard').className = `kpi-card ${finalProfit < 0 ? 'accent-red' : 'accent-indigo'}`;

  // 대시보드 표 및 차트 업데이트
  renderDashboardTablesAndCharts(year, start, end, evaluatedSales, periodRows, periodTarget, actualSales, totalVarCost, netProfit, finalProfit, totalOrders);

  // 시트2: 일일 매출 입력 테이블 렌더링
  renderDailySalesTable(evaluatedSales, start, end);

  // 시트3: 판매채널 비용 설정 테이블 렌더링
  renderChannelRatesTable();

  // 시트4: 거래처별·채널별 손익 분석 렌더링
  renderAnalysisTab(periodRows, actualSales, totalOrders, periodCommonCost);

  // 시트5: 목표매출 및 공통비용 설정 렌더링
  renderTargetsAndCommonTab(year, start, end);

  // 시트6: 드롭다운 및 마스터 목록 렌더링
  renderMasterSelectsAndLists();
}

function renderDashboardTablesAndCharts(year, start, end, evaluatedSales, periodRows, periodTarget, actualSales, totalVarCost, netProfit, finalProfit, totalOrders) {
  // 1) 월별 추이 데이터
  const mLabels = [];
  const mTargets = [];
  const mSales = [];
  const mCosts = [];
  const mProfits = [];

  const mTbody = document.querySelector('#dashboardMonthlyTable tbody');
  const mTfoot = document.querySelector('#dashboardMonthlyTable tfoot');
  mTbody.innerHTML = '';

  let sumTarget = 0, sumSales = 0, sumCost = 0, sumProfit = 0;

  for (let m = 1; m <= 12; m++) {
    const label = `${m}월`;
    const target = (state.monthlyTargets.find(t => t.month === m) || { target: 0 }).target;
    const prefix = `${year}-${String(m).padStart(2, '0')}`;
    const rows = evaluatedSales.filter(r => r.isValid && r.date.startsWith(prefix));
    const s = rows.reduce((a, b) => a + b.sales, 0);
    const c = rows.reduce((a, b) => a + b.totalCost, 0);
    const p = s - c;
    const ach = target > 0 ? s / target : 0;
    const mar = s > 0 ? p / s : 0;

    mLabels.push(label);
    mTargets.push(target);
    mSales.push(s);
    mCosts.push(c);
    mProfits.push(p);

    sumTarget += target;
    sumSales += s;
    sumCost += c;
    sumProfit += p;

    const tr = document.createElement('tr');
    if (p < 0) tr.classList.add('row-negative');
    tr.innerHTML = `
      <td><strong>${label}</strong></td>
      <td class="num">${fmtNum(target)}</td>
      <td class="num"><strong>${fmtNum(s)}</strong></td>
      <td class="num">${fmtPct(ach, 1)}</td>
      <td class="num">${fmtNum(c)}</td>
      <td class="num profit-col ${p < 0 ? 'cell-negative' : ''}"><strong>${fmtNum(p)}</strong></td>
      <td class="num ${p < 0 ? 'cell-negative' : ''}">${fmtPct(mar, 1)}</td>
    `;
    mTbody.appendChild(tr);
  }

  mTfoot.innerHTML = `
    <tr>
      <td>연간 합계</td>
      <td class="num">${fmtNum(sumTarget)}</td>
      <td class="num">${fmtNum(sumSales)}</td>
      <td class="num">${fmtPct(sumTarget > 0 ? sumSales / sumTarget : 0, 1)}</td>
      <td class="num">${fmtNum(sumCost)}</td>
      <td class="num ${sumProfit < 0 ? 'cell-negative' : ''}">${fmtNum(sumProfit)}</td>
      <td class="num">${fmtPct(sumSales > 0 ? sumProfit / sumSales : 0, 1)}</td>
    </tr>
  `;

  // 2) 판매채널별 매출 비중 표
  const chTbody = document.querySelector('#dashboardChannelTable tbody');
  const chTfoot = document.querySelector('#dashboardChannelTable tfoot');
  chTbody.innerHTML = '';

  const chLabels = [];
  const chSalesArr = [];

  state.channels.forEach(chName => {
    const rows = periodRows.filter(r => r.channel === chName);
    const ord = rows.reduce((a, b) => a + b.orders, 0);
    const s = rows.reduce((a, b) => a + b.sales, 0);
    const c = rows.reduce((a, b) => a + b.totalCost, 0);
    const p = s - c;
    const share = actualSales > 0 ? s / actualSales : 0;
    const mar = s > 0 ? p / s : 0;

    chLabels.push(`${chName} (${fmtPct(share, 1)})`);
    chSalesArr.push(s);

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${chName}</strong></td>
      <td class="num">${fmtNum(ord)}건</td>
      <td class="num"><strong>${fmtNum(s)}</strong></td>
      <td class="num"><strong>${fmtPct(share, 1)}</strong></td>
      <td class="num">${fmtNum(c)}</td>
      <td class="num ${p < 0 ? 'cell-negative' : ''}"><strong>${fmtNum(p)}</strong></td>
      <td class="num ${p < 0 ? 'cell-negative' : ''}">${fmtPct(mar, 1)}</td>
    `;
    chTbody.appendChild(tr);
  });

  chTfoot.innerHTML = `
    <tr>
      <td>전체 합계</td>
      <td class="num">${fmtNum(totalOrders)}건</td>
      <td class="num">${fmtNum(actualSales)}</td>
      <td class="num">100.0%</td>
      <td class="num">${fmtNum(totalVarCost)}</td>
      <td class="num ${netProfit < 0 ? 'cell-negative' : ''}">${fmtNum(netProfit)}</td>
      <td class="num">${fmtPct(actualSales > 0 ? netProfit / actualSales : 0, 1)}</td>
    </tr>
  `;

  // 차트 1: 목표매출 vs 실제매출 비교
  updateChart('chartTargetVsActual', {
    type: 'bar',
    data: {
      labels: ['목표매출', '실제매출', '순이익(공통비전)', '최종순이익(공통비후)'],
      datasets: [{
        label: '선택기간 금액 (원)',
        data: [periodTarget, actualSales, netProfit, finalProfit],
        backgroundColor: ['#94a3b8', '#2563eb', '#10b981', '#6366f1'],
        borderRadius: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } }
    }
  });

  // 차트 2: 월별 매출·비용·순이익 추이
  updateChart('chartMonthlyTrend', {
    type: 'bar',
    data: {
      labels: mLabels,
      datasets: [
        { label: '실제매출', data: mSales, backgroundColor: '#2563eb', borderRadius: 4 },
        { label: '총비용(변동)', data: mCosts, backgroundColor: '#f59e0b', borderRadius: 4 },
        { label: '순이익(공통비전)', data: mProfits, backgroundColor: '#10b981', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } }
    }
  });

  // 차트 3: 판매채널별 매출 비중 도넛 그래프
  updateChart('chartChannelShare', {
    type: 'doughnut',
    data: {
      labels: chLabels,
      datasets: [{
        data: chSalesArr,
        backgroundColor: ['#2563eb', '#10b981', '#f59e0b', '#6366f1', '#ec4899', '#14b8a6', '#64748b']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '55%',
      plugins: {
        legend: { position: 'right' },
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
}

function renderDailySalesTable(evaluatedSales, start, end) {
  const onlyPeriod = document.getElementById('chkFilterDailyByPeriod').checked;
  const rows = onlyPeriod
    ? evaluatedSales.filter(r => r.date >= start && r.date <= end)
    : evaluatedSales;

  const tbody = document.querySelector('#dailySalesTable tbody');
  const tfoot = document.querySelector('#dailySalesTable tfoot');
  tbody.innerHTML = '';

  let sumOrd = 0, sumSales = 0, sumCogs = 0, sumExtra = 0, sumFee = 0, sumOrdCost = 0, sumCost = 0, sumProfit = 0;

  rows.forEach(r => {
    sumOrd += r.orders;
    sumSales += r.sales;
    sumCogs += r.cogs;
    sumExtra += r.extraCost;
    if (r.isValid) {
      sumFee += r.feeAmt;
      sumOrdCost += r.orderCost;
      sumCost += r.totalCost;
      sumProfit += r.netProfit;
    }

    const tr = document.createElement('tr');
    const errCls = !r.isValid ? 'cell-error' : '';
    const negCls = r.isValid && r.netProfit < 0 ? 'cell-negative' : '';

    tr.innerHTML = `
      <td class="cell-input">${r.date}</td>
      <td class="cell-input"><strong>${r.client}</strong></td>
      <td class="cell-input">${r.channel}</td>
      <td class="cell-input num">${fmtNum(r.orders)}</td>
      <td class="cell-input num"><strong>${fmtNum(r.sales)}</strong></td>
      <td class="cell-input num">${fmtNum(r.cogs)}</td>
      <td class="cell-input num">${fmtNum(r.extraCost)}</td>
      <td class="cell-input" style="text-align:left;">${r.note || ''}</td>
      <td class="cell-calc ${errCls}">${r.status}</td>
      <td class="cell-calc num ${errCls}">${r.isValid ? fmtPct(r.feeRate, 2) : '비용 설정 필요'}</td>
      <td class="cell-calc num ${errCls}">${r.isValid ? fmtNum(r.feeAmt) : '비용 설정 필요'}</td>
      <td class="cell-calc num ${errCls}">${r.isValid ? fmtNum(r.unitCost) : '비용 설정 필요'}</td>
      <td class="cell-calc num ${errCls}">${r.isValid ? fmtNum(r.orderCost) : '비용 설정 필요'}</td>
      <td class="cell-calc num ${errCls}">${r.isValid ? fmtNum(r.totalCost) : '비용 설정 필요'}</td>
      <td class="cell-calc num ${errCls} ${negCls}"><strong>${r.isValid ? fmtNum(r.netProfit) : '비용 설정 필요'}</strong></td>
      <td class="cell-calc num ${errCls} ${negCls}">${r.isValid ? fmtPct(r.margin, 1) : '비용 설정 필요'}</td>
      <td class="cell-calc"><button class="btn-danger-sm" onclick="deleteDailyRow(${r.id})">삭제</button></td>
    `;
    tbody.appendChild(tr);
  });

  tfoot.innerHTML = `
    <tr>
      <td colspan="3">표시된 내역 합계 (${rows.length}건)</td>
      <td class="num">${fmtNum(sumOrd)}</td>
      <td class="num">${fmtNum(sumSales)}</td>
      <td class="num">${fmtNum(sumCogs)}</td>
      <td class="num">${fmtNum(sumExtra)}</td>
      <td></td>
      <td>정상 합산</td>
      <td class="num">-</td>
      <td class="num">${fmtNum(sumFee)}</td>
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
    const unitSum = (Number(r.shipFee) || 0) + (Number(r.packFee) || 0) + (Number(r.etcFee) || 0);
    const status = checkRateOverlap(r, state.channelRates);
    const isErr = status !== '정상';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td class="cell-input"><strong>${r.channel}</strong></td>
      <td class="cell-input"><input type="date" class="table-inline-date" value="${r.start}" onchange="updateRateField(${r.id}, 'start', this.value)" /></td>
      <td class="cell-input"><input type="date" class="table-inline-date" value="${r.end || ''}" onchange="updateRateField(${r.id}, 'end', this.value)" /></td>
      <td class="cell-input num"><input type="number" class="table-inline-input" step="0.1" value="${(r.feeRate * 100).toFixed(2)}" onchange="updateRateField(${r.id}, 'feeRate', Number(this.value)/100)" />%</td>
      <td class="cell-input num"><input type="number" class="table-inline-input" step="100" value="${r.shipFee}" onchange="updateRateField(${r.id}, 'shipFee', Number(this.value))" /></td>
      <td class="cell-input num"><input type="number" class="table-inline-input" step="100" value="${r.packFee}" onchange="updateRateField(${r.id}, 'packFee', Number(this.value))" /></td>
      <td class="cell-input num"><input type="number" class="table-inline-input" step="10" value="${r.etcFee}" onchange="updateRateField(${r.id}, 'etcFee', Number(this.value))" /></td>
      <td class="cell-calc num"><strong>${fmtWon(unitSum)}</strong></td>
      <td class="cell-calc ${isErr ? 'cell-error' : ''}">${status}</td>
      <td class="cell-input" style="text-align:left;">${r.note || ''}</td>
      <td class="cell-calc"><button class="btn-danger-sm" onclick="deleteRateRow(${r.id})">삭제</button></td>
    `;
    tbody.appendChild(tr);
  });
}

function renderAnalysisTab(periodRows, totalSales, totalOrders, periodCommonCost) {
  const sortBy = document.getElementById('analysisSortBy').value;
  const allocRule = document.getElementById('analysisAllocRule').value;

  function buildGroupStats(names, keyField) {
    const list = names.map((name, idx) => {
      const rows = periodRows.filter(r => r[keyField] === name);
      const orders = rows.reduce((a, b) => a + b.orders, 0);
      const sales = rows.reduce((a, b) => a + b.sales, 0);
      const share = totalSales > 0 ? sales / totalSales : 0;
      const cogs = rows.reduce((a, b) => a + b.cogs, 0);
      const fee = rows.reduce((a, b) => a + b.feeAmt, 0);
      const orderCost = rows.reduce((a, b) => a + b.orderCost, 0);
      const extra = rows.reduce((a, b) => a + b.extraCost, 0);
      const totalCost = cogs + fee + orderCost + extra;
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

      return { idx, name, orders, sales, share, cogs, fee, orderCost, extra, totalCost, netProfit, margin, allocCommon, finalProfit, finalMargin };
    });

    if (sortBy === 'profit') {
      list.sort((a, b) => b.netProfit - a.netProfit);
    } else if (sortBy === 'sales') {
      list.sort((a, b) => b.sales - a.sales);
    }
    return list;
  }

  const clientStats = buildGroupStats(state.clients, 'client');
  const channelStats = buildGroupStats(state.channels, 'channel');

  function fillAnalysisTable(tableId, stats, totalLabel) {
    const tbody = document.querySelector(`#${tableId} tbody`);
    const tfoot = document.querySelector(`#${tableId} tfoot`);
    tbody.innerHTML = '';

    let sOrd = 0, sSales = 0, sCogs = 0, sFee = 0, sOrdCost = 0, sExtra = 0, sCost = 0, sProfit = 0, sAlloc = 0, sFinal = 0;

    stats.forEach((st, i) => {
      sOrd += st.orders;
      sSales += st.sales;
      sCogs += st.cogs;
      sFee += st.fee;
      sOrdCost += st.orderCost;
      sExtra += st.extra;
      sCost += st.totalCost;
      sProfit += st.netProfit;
      sAlloc += st.allocCommon;
      sFinal += st.finalProfit;

      const neg1 = st.netProfit < 0 ? 'cell-negative' : '';
      const neg2 = st.finalProfit < 0 ? 'cell-negative' : '';

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${i + 1}</td>
        <td><strong>${st.name}</strong></td>
        <td class="num">${fmtNum(st.orders)}</td>
        <td class="num"><strong>${fmtNum(st.sales)}</strong></td>
        <td class="num">${fmtPct(st.share, 1)}</td>
        <td class="num">${fmtNum(st.cogs)}</td>
        <td class="num">${fmtNum(st.fee)}</td>
        <td class="num">${fmtNum(st.orderCost)}</td>
        <td class="num">${fmtNum(st.extra)}</td>
        <td class="num">${fmtNum(st.totalCost)}</td>
        <td class="num ${neg1}"><strong>${fmtNum(st.netProfit)}</strong></td>
        <td class="num ${neg1}">${fmtPct(st.margin, 1)}</td>
        <td class="num">${fmtNum(st.allocCommon)}</td>
        <td class="num ${neg2}"><strong>${fmtNum(st.finalProfit)}</strong></td>
        <td class="num ${neg2}">${fmtPct(st.finalMargin, 1)}</td>
      `;
      tbody.appendChild(tr);
    });

    tfoot.innerHTML = `
      <tr>
        <td colspan="2">${totalLabel}</td>
        <td class="num">${fmtNum(sOrd)}</td>
        <td class="num">${fmtNum(sSales)}</td>
        <td class="num">100.0%</td>
        <td class="num">${fmtNum(sCogs)}</td>
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
  }

  fillAnalysisTable('clientAnalysisTable', clientStats, `전체 ${clientStats.length}개 거래처 합계`);
  fillAnalysisTable('channelAnalysisTable', channelStats, `전체 ${channelStats.length}개 판매채널 합계`);

  // 차트 4: 거래처별 순이익 비교
  updateChart('chartClientProfit', {
    type: 'bar',
    data: {
      labels: clientStats.map(c => c.name),
      datasets: [
        { label: '순이익 (공통비 전)', data: clientStats.map(c => c.netProfit), backgroundColor: '#2563eb', borderRadius: 4 },
        { label: '최종 순이익 (공통비 후)', data: clientStats.map(c => c.finalProfit), backgroundColor: '#10b981', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } }
    }
  });

  // 차트 5: 판매채널별 매출·순이익 비교
  updateChart('chartChannelCompare', {
    type: 'bar',
    data: {
      labels: channelStats.map(c => c.name),
      datasets: [
        { label: '매출액 (원)', data: channelStats.map(c => c.sales), backgroundColor: '#3b82f6', borderRadius: 4 },
        { label: '순이익 (공통비 전)', data: channelStats.map(c => c.netProfit), backgroundColor: '#10b981', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } }
    }
  });
}

function renderTargetsAndCommonTab(year, start, end) {
  const tTbody = document.querySelector('#monthlyTargetsTable tbody');
  const tTfoot = document.querySelector('#monthlyTargetsTable tfoot');
  tTbody.innerHTML = '';

  let sumAnnual = 0, sumApplied = 0;

  state.monthlyTargets.forEach(t => {
    const days = new Date(Date.UTC(year, t.month, 0)).getUTCDate();
    const daily = Math.round(t.target / days);
    const applied = calcProRatedTarget(year, t.month, start, end, t.target);
    sumAnnual += t.target;
    sumApplied += applied;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${t.month}월</strong></td>
      <td>${days}일</td>
      <td class="cell-input num">
        <input type="number" class="table-inline-input" style="width:130px" step="100000" value="${t.target}" onchange="updateMonthlyTarget(${t.month}, Number(this.value))" />
      </td>
      <td class="cell-calc num">${fmtWon(daily)}</td>
      <td class="cell-calc num"><strong>${fmtWon(applied)}</strong></td>
    `;
    tTbody.appendChild(tr);
  });

  tTfoot.innerHTML = `
    <tr>
      <td colspan="2">연간 합계</td>
      <td class="num">${fmtWon(sumAnnual)}</td>
      <td class="num">${fmtWon(Math.round(sumAnnual / 365))}/일</td>
      <td class="num">${fmtWon(sumApplied)}</td>
    </tr>
  `;

  // 공통비용 표
  const cTbody = document.querySelector('#commonCostsTable tbody');
  const cTfoot = document.querySelector('#commonCostsTable tfoot');
  cTbody.innerHTML = '';

  let sumAllCommon = 0, sumPeriodCommon = 0;
  state.commonCosts.forEach(c => {
    const inPeriod = c.date >= start && c.date <= end;
    sumAllCommon += c.amount;
    if (inPeriod) sumPeriodCommon += c.amount;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="cell-input">${c.date}</td>
      <td class="cell-input"><strong>${c.category}</strong></td>
      <td class="cell-input num">${fmtWon(c.amount)}</td>
      <td class="cell-calc">${inPeriod ? '✅ 포함' : '제외'}</td>
      <td class="cell-input" style="text-align:left;">${c.note || ''}</td>
      <td class="cell-calc"><button class="btn-danger-sm" onclick="deleteCommonCost(${c.id})">삭제</button></td>
    `;
    cTbody.appendChild(tr);
  });

  cTfoot.innerHTML = `
    <tr>
      <td colspan="2">공통비용 합계</td>
      <td class="num">전체: ${fmtWon(sumAllCommon)}</td>
      <td colspan="3">현재 조회기간 반영액: ${fmtWon(sumPeriodCommon)}</td>
    </tr>
  `;
}

function renderMasterSelectsAndLists() {
  const clientSel = document.getElementById('inSaleClient');
  const chSel1 = document.getElementById('inSaleChannel');
  const chSel2 = document.getElementById('inRateChannel');

  const prevClient = clientSel.value;
  const prevCh1 = chSel1.value;
  const prevCh2 = chSel2.value;

  clientSel.innerHTML = state.clients.map(c => `<option value="${c}">${c}</option>`).join('');
  const chOptions = state.channels.map(c => `<option value="${c}">${c}</option>`).join('');
  chSel1.innerHTML = chOptions;
  chSel2.innerHTML = chOptions;

  if (state.clients.includes(prevClient)) clientSel.value = prevClient;
  if (state.channels.includes(prevCh1)) chSel1.value = prevCh1;
  if (state.channels.includes(prevCh2)) chSel2.value = prevCh2;

  document.getElementById('channelMasterList').innerHTML = state.channels.map(c => `<li>🛒 ${c}</li>`).join('');
  document.getElementById('clientMasterList').innerHTML = state.clients.map(c => `<li>🏢 ${c}</li>`).join('');
}

function updateChart(canvasId, config) {
  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }
  const ctx = document.getElementById(canvasId).getContext('2d');
  charts[canvasId] = new Chart(ctx, config);
}

// ============================================================================
// 이벤트 핸들러 및 사용자 액션
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

window.updateMonthlyTarget = function(month, val) {
  const t = state.monthlyTargets.find(x => x.month === month);
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
  // 탭 전환
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
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

  ['filterYear', 'filterMonth', 'filterStartDate', 'filterEndDate', 'chkFilterDailyByPeriod', 'analysisSortBy', 'analysisAllocRule'].forEach(id => {
    document.getElementById(id).addEventListener('change', renderAll);
  });

  // 예시 데이터 초기화
  document.getElementById('btnResetSample').addEventListener('click', () => {
    state = JSON.parse(JSON.stringify(INITIAL_STATE));
    saveState();
    renderAll();
  });

  // 일일 매출 추가 폼
  document.getElementById('addDailySaleForm').addEventListener('submit', e => {
    e.preventDefault();
    const ordersVal = Number(document.getElementById('inSaleOrders').value);
    if (!Number.isInteger(ordersVal) || ordersVal < 0) {
      alert('주문건수는 0 이상의 정수만 입력할 수 있습니다.');
      return;
    }
    const newRow = {
      id: Date.now(),
      date: document.getElementById('inSaleDate').value,
      client: document.getElementById('inSaleClient').value,
      channel: document.getElementById('inSaleChannel').value,
      orders: ordersVal,
      sales: Number(document.getElementById('inSaleAmount').value) || 0,
      cogs: Number(document.getElementById('inSaleCogs').value) || 0,
      extraCost: Number(document.getElementById('inSaleExtra').value) || 0,
      note: document.getElementById('inSaleNote').value
    };
    state.dailySales.push(newRow);
    saveState();
    renderAll();
    document.getElementById('inSaleNote').value = '';
  });

  // 채널 기간별 비용 설정 추가 폼
  document.getElementById('addChannelRateForm').addEventListener('submit', e => {
    e.preventDefault();
    const newRate = {
      id: Date.now(),
      channel: document.getElementById('inRateChannel').value,
      start: document.getElementById('inRateStart').value,
      end: document.getElementById('inRateEnd').value || '',
      feeRate: (Number(document.getElementById('inRateFeePct').value) || 0) / 100,
      shipFee: Number(document.getElementById('inRateShip').value) || 0,
      packFee: Number(document.getElementById('inRatePack').value) || 0,
      etcFee: Number(document.getElementById('inRateEtc').value) || 0,
      note: document.getElementById('inRateNote').value
    };
    state.channelRates.push(newRate);
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

  // 거래처 마스터 추가
  document.getElementById('addClientMasterForm').addEventListener('submit', e => {
    e.preventDefault();
    const name = document.getElementById('inNewClientName').value.trim();
    if (name && !state.clients.includes(name)) {
      state.clients.push(name);
      saveState();
      renderAll();
    }
    document.getElementById('inNewClientName').value = '';
  });

  renderAll();
});
