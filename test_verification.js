const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { CHANNELS, COMMON_COST_CATEGORIES, TARGET_MONTHS, CHANNEL_RATES, DAILY_SALES, COMMON_COSTS, computeRowMetrics } = require('./excel_data_styles');

console.log('=== [수정사항 반영 매출관리 프로그램 자동 검증 테스트] ===');

// 1. 엑셀 파일 확인
const xlsxPath = path.join(__dirname, '거래처별_매출순이익_관리프로그램.xlsx');
assert(fs.existsSync(xlsxPath), '엑셀 파일이 존재해야 합니다.');
console.log(`✅ 1. 엑셀 파일 생성 확인 완료 (${fs.statSync(xlsxPath).size.toLocaleString()} bytes)`);

// 2. 목표매출 기간: 2026년 11월 ~ 2027년 10월 (12개월) 및 예시 수치 0원 초기화 확인
assert.strictEqual(TARGET_MONTHS.length, 12);
assert.strictEqual(TARGET_MONTHS[0].ym, '2026-11');
assert.strictEqual(TARGET_MONTHS[11].ym, '2027-10');
const totalTargetSum = TARGET_MONTHS.reduce((a, b) => a + b.target, 0);
assert.strictEqual(totalTargetSum, 0, '실사용을 위해 초기 목표매출 합계는 0원이어야 합니다.');
assert.strictEqual(DAILY_SALES.length, 0, '실사용을 위해 일일 매출 예시 데이터는 비워져 있어야 합니다.');
assert.strictEqual(COMMON_COSTS.length, 0, '실사용을 위해 공통비용 예시 데이터는 비워져 있어야 합니다.');
assert.strictEqual(CHANNEL_RATES.length, 8, '8개 기본 판매채널 설정 행이 준비되어 있어야 합니다.');
assert(CHANNEL_RATES.every(r => r.feeRate === 0 && r.shipUnder10 === 0 && r.shipOver10 === 0 && r.boxCost === 0), '채널별 초기 단가는 0으로 비워져 있어야 합니다.');
console.log('✅ 2. 모든 예시 수치 삭제 및 실사용 초기 상태(0원/빈 내역) 확인 완료');

// 3. 10kg 미만 / 10kg 이상 주문건수 구분 및 박스·포장비닐·리플렛 비용 합산 로직 검증 (가상 테스트 데이터)
const mockRates = [
  { channel: '스마트스토어', start: '2026-11-01', end: '2026-11-15', feeRate: 0.055, shipUnder10: 3200, shipOver10: 4200, boxCost: 650, vinylCost: 150, leafletCost: 100, etcCost: 0 },
  { channel: '스마트스토어', start: '2026-11-16', end: '', feeRate: 0.058, shipUnder10: 3500, shipOver10: 4500, boxCost: 750, vinylCost: 150, leafletCost: 100, etcCost: 0 },
  { channel: '공판장', start: '2026-11-01', end: '', feeRate: 0.070, shipUnder10: 1000, shipOver10: 1500, boxCost: 650, vinylCost: 100, leafletCost: 0, etcCost: 0 }
];
const beforeSale = { date: '2026-11-14', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0 };
const afterSale = { date: '2026-11-17', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0 };
const mBefore = computeRowMetrics(beforeSale, mockRates);
const mAfter = computeRowMetrics(afterSale, mockRates);

assert.strictEqual(mBefore.status, '정상');
assert.strictEqual(mBefore.totalOrders, 30);
assert.strictEqual(mBefore.feeAmt, 66000);
assert.strictEqual(mBefore.unitUnder10, 4100);
assert.strictEqual(mBefore.unitOver10, 5100);
assert.strictEqual(mBefore.orderCost, 133000);
assert.strictEqual(mBefore.totalCost, 199000);
assert.strictEqual(mBefore.netProfit, 1001000);

assert.strictEqual(mAfter.status, '정상');
assert.strictEqual(mAfter.totalOrders, 30);
assert.strictEqual(mAfter.feeAmt, 69600);
assert.strictEqual(mAfter.unitUnder10, 4500);
assert.strictEqual(mAfter.unitOver10, 5500);
assert.strictEqual(mAfter.orderCost, 145000);
assert.strictEqual(mAfter.totalCost, 214600);
assert.strictEqual(mAfter.netProfit, 985400);

// 10kg 이상 주문건수가 0건일 때 10kg 이상 단가에 비용이 0원으로 표시되고, 10kg 미만 주문건수가 0건일 때 10kg 미만 단가에 0원으로 표시되는지 검증
const zeroOver10Sale = { date: '2026-11-17', channel: '스마트스토어', ordersUnder10: 15, ordersOver10: 0, sales: 450000, extraCost: 0 };
const mZeroOver10 = computeRowMetrics(zeroOver10Sale, mockRates);
assert.strictEqual(mZeroOver10.unitUnder10, 4500);
assert.strictEqual(mZeroOver10.unitOver10, 0, '10kg 이상 주문건수가 0건인 경우 10kg 이상 단가에 비용이 잡히지 않고 0이어야 합니다.');
assert.strictEqual(mZeroOver10.orderCost, 15 * 4500);

const zeroUnder10Sale = { date: '2026-11-17', channel: '스마트스토어', ordersUnder10: 0, ordersOver10: 8, sales: 400000, extraCost: 0 };
const mZeroUnder10 = computeRowMetrics(zeroUnder10Sale, mockRates);
assert.strictEqual(mZeroUnder10.unitUnder10, 0, '10kg 미만 주문건수가 0건인 경우 10kg 미만 단가에 비용이 잡히지 않고 0이어야 합니다.');
assert.strictEqual(mZeroUnder10.unitOver10, 5500);
assert.strictEqual(mZeroUnder10.orderCost, 8 * 5500);
console.log('✅ 3. 10kg 미만/이상 택배비 구분(주문 0건 시 단가 0원 처리 포함) 및 박스·포장비닐·리플렛 비용 합산, 단가 변경 전후 수식 검증 완료');

// 4. 판매채널 '공판장' 추가 및 계산 검증
assert(CHANNELS.includes('공판장'), '판매채널 목록에 공판장이 포함되어야 합니다.');
assert.strictEqual(CHANNELS.length, 8);
const gongSale = { date: '2026-11-28', channel: '공판장', ordersUnder10: 30, ordersOver10: 40, sales: 2100000, extraCost: 0 };
const mGong = computeRowMetrics(gongSale, mockRates);
assert.strictEqual(mGong.status, '정상');
assert.strictEqual(mGong.totalOrders, 70);
assert.strictEqual(mGong.feeAmt, 147000);
assert.strictEqual(mGong.unitUnder10, 1750);
assert.strictEqual(mGong.unitOver10, 2250);
assert.strictEqual(mGong.orderCost, 142500);
assert.strictEqual(mGong.totalCost, 289500);
assert.strictEqual(mGong.netProfit, 1810500);
console.log('✅ 4. 판매채널 [공판장] 등록 및 무게별 비용·순이익 자동계산 검증 완료');

// 5. 공통비용 8개 항목 확인
assert.strictEqual(COMMON_COST_CATEGORIES.length, 8);
assert.deepStrictEqual(COMMON_COST_CATEGORIES, [
  '종묘비',
  '무기질비료비',
  '유기질비료비',
  '농약비',
  '수도광열비',
  '기타재료비',
  '고용노동비',
  '기타'
]);
console.log('✅ 5. 공통비용 8대 항목(종묘비, 무기질비료비, 유기질비료비, 농약비, 수도광열비, 기타재료비, 고용노동비, 기타) 설정 확인 완료');

// 6. 모바일 화면 구성(반복 공통 영역 제거, 메뉴별 즉시 표시, 목표매출 총합계 분리, 스크롤 초기화, 조회 실패/재시도 UI) 검증
const htmlContent = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const cssContent = fs.readFileSync(path.join(__dirname, 'styles.css'), 'utf8');
const jsContent = fs.readFileSync(path.join(__dirname, 'app.js'), 'utf8');

const mainIdx = htmlContent.indexOf('<main');
const firstTabIdx = htmlContent.indexOf('<section id="tab-dashboard"');
const beforeFirstTabHtml = htmlContent.slice(mainIdx, firstTabIdx);
assert(!beforeFirstTabHtml.includes('id="filterMode"'), '모든 메뉴 상단에 반복 표시되던 공통 조회방식 영역이 제거되어야 합니다.');
assert(!beforeFirstTabHtml.includes('id="totalAnnualTargetBadge"'), '12개월 목표매출총합계가 공통 상단 영역에 반복 표시되지 않아야 합니다.');

const tabTargetsIdx = htmlContent.indexOf('<section id="tab-targets"');
const tabGuideIdx = htmlContent.indexOf('<section id="tab-guide"');
const tabTargetsHtml = htmlContent.slice(tabTargetsIdx, tabGuideIdx);
assert(tabTargetsHtml.includes('id="tab5TotalTargetSum"') && tabTargetsHtml.includes('12개월 목표매출 총합계'), '12개월 목표매출 총합계는 6. 목표매출 메뉴에 표시되어야 합니다.');

// 각 메뉴 선택 시 핵심 입력 폼/표가 안내문보다 먼저 나오는지 검증
const tabDailyHtml = htmlContent.slice(htmlContent.indexOf('<section id="tab-daily"'), htmlContent.indexOf('<section id="tab-rates"'));
assert(tabDailyHtml.indexOf('id="addDailySaleForm"') < tabDailyHtml.indexOf('info-alert-details'), '2. 일일 매출 메뉴는 입력 폼과 내역 표가 상단에 바로 표시되어야 합니다.');

const tabRatesHtml = htmlContent.slice(htmlContent.indexOf('<section id="tab-rates"'), htmlContent.indexOf('<section id="tab-analysis"'));
assert(tabRatesHtml.indexOf('id="addChannelRateForm"') < tabRatesHtml.indexOf('info-alert-details'), '3. 판매채널 비용 설정 메뉴는 설정 폼과 표가 상단에 바로 표시되어야 합니다.');

const tabCommonHtml = htmlContent.slice(htmlContent.indexOf('<section id="tab-common"'), htmlContent.indexOf('<section id="tab-targets"'));
assert(tabCommonHtml.indexOf('id="addCommonCostForm"') < tabCommonHtml.indexOf('info-alert-details'), '5. 공통비용 메뉴는 입력 폼과 저장 내역이 상단에 바로 표시되어야 합니다.');

assert(cssContent.includes('flex-direction: row') && cssContent.includes('.desktop-only-desc'), '모바일에서 탭 메뉴가 가로 1줄 스크롤로 배치되고 불필요한 상단 설명이 숨겨져야 합니다.');
assert(jsContent.includes('window.scrollTo') && jsContent.includes('retryCloudFetch') && jsContent.includes('loadErrorState'), '메뉴 전환 시 스크롤 최상단 이동 및 데이터 조회 실패 시 안내/재시도 기능이 포함되어야 합니다.');

// 추가 수정 사항 5가지 검증:
// 1) 전체데이터 초기화 버튼 삭제
assert(!htmlContent.includes('id="btnResetSample"') && !htmlContent.includes('전체 데이터 초기화'), '전체데이터 초기화 버튼이 삭제되어야 합니다.');
// 2) 일자별 순이익 변동 현황 (선택 기간) 삭제
assert(!htmlContent.includes('chartDailyProfitFluctuation') && !htmlContent.includes('일자별 순이익 변동 현황 (선택 기간)'), '일자별 순이익 변동 현황 (선택 기간)이 삭제되어야 합니다.');
// 3) 매출·지출내역 입력 및 자동계산 시트 날짜순 정렬
assert(jsContent.includes('String(a.date || \'\').localeCompare(String(b.date || \'\'))'), '매출·지출내역 입력 및 자동계산 시트가 날짜순으로 정렬되어야 합니다.');
// 4) 일일매출액 10원 단위까지 입력 가능 (step="10")
assert(htmlContent.includes('id="inSaleAmount" class="input-field yellow-input" min="0" step="10"'), '일일매출액 입력란이 10원 단위(step="10")로 설정되어야 합니다.');
// 5) 시트의 비고를 판매무게(kg)으로 수정하고 하단 합계에 무게 총합 표시
assert(htmlContent.includes('판매무게(kg)') && jsContent.includes('판매무게 총합(kg)') && jsContent.includes('fmtWeightKg(sumWeight)'), '비고가 판매무게(kg)으로 수정되고 하단 합계에 무게 총합이 표시되어야 합니다.');
// 6) 달력이 매출·지출내역 입력 및 자동계산 시트보다 위에 배치되고 날짜별 매출총액·수수료총액·순수익총액을 표시하는지 검증
const calCardIdx = tabDailyHtml.indexOf('id="dailySummaryCalendarCard"');
const sheetTableIdx = tabDailyHtml.indexOf('id="dailySalesTable"');
assert(calCardIdx !== -1 && sheetTableIdx !== -1 && calCardIdx < sheetTableIdx, '달력은 매출·지출내역 입력 및 자동계산 시트보다 위에 디스플레이되어야 합니다.');
assert(jsContent.includes('renderDailySummaryCalendar') && jsContent.includes('salesTotal') && jsContent.includes('feeTotal') && jsContent.includes('netProfitTotal'), '달력에 일일 매출총액, 수수료총액, 순수익총액이 날짜별로 집계·표시되어야 합니다.');
// 7) 매출·지출내역 입력 및 자동계산 시트 메인화면 10개 기본 표시 및 나머지 내역 펼치기 기능 검증
assert(jsContent.includes('DAILY_TABLE_DEFAULT_LIMIT = 10') && jsContent.includes('toggleDailySalesExpand') && htmlContent.includes('id="btnToggleDailyExpand"'), '매출·지출내역 입력 및 자동계산 시트는 메인화면에서 10개만 기본 표시하고 펼치기 버튼으로 나머지를 볼 수 있어야 합니다.');

console.log('✅ 6. 모바일 반복 공통 영역 제거, 메뉴 선택 시 내용 즉시 표시, 스크롤 초기화, 조회 실패/재시도 UI, 상단 달력(매출/수수료/순수익) 및 시트 10개 표시·펼치기 검증 완료');

// 7. PC ↔ 모바일 계정 동기화 API (/api/state) 및 기존 데이터 보존(merge)·비공개 인증 검증
async function verifyCloudSyncApi() {
  process.env.SPM_LOCAL_FALLBACK = '1';
  const apiHandler = require('./api/state');

  function callApi({ method = 'GET', auth = null, body = null }) {
    return new Promise((resolve) => {
      const headers = {};
      if (auth) {
        const b64 = Buffer.from(JSON.stringify(auth), 'utf8').toString('base64');
        headers.authorization = `Bearer ${b64}`;
      }
      const req = { method, headers, body };
      const res = {
        statusCode: 200,
        headers: {},
        setHeader(k, v) { this.headers[k] = v; },
        end(payload) {
          resolve({ statusCode: this.statusCode, data: JSON.parse(payload) });
        }
      };
      apiHandler(req, res);
    });
  }

  const testUser = { userId: `test_sync_${Date.now()}`, password: 'correct_password_123' };

  // 7-1. 미인증 접근 차단 확인
  const unauthRes = await callApi({ method: 'GET' });
  assert.strictEqual(unauthRes.statusCode, 401);
  assert.strictEqual(unauthRes.data.code, 'AUTH_REQUIRED');

  // 7-2. PC에서 기존 로컬 데이터를 가진 상태로 첫 로그인 및 보존(merge) 저장
  const pcInitialState = {
    channels: CHANNELS,
    targetMonths: TARGET_MONTHS.map(t => t.ym === '2026-11' ? { ...t, target: 5000000 } : t),
    channelRates: mockRates,
    dailySales: [
      { id: 101, date: '2026-11-14', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0, note: 'PC저장분' }
    ],
    commonCosts: [
      { id: 201, date: '2026-11-15', category: '종묘비', detail: '배추 모종', amount: 150000, note: 'PC공통비' }
    ]
  };
  const pcSaveRes = await callApi({
    method: 'POST',
    auth: testUser,
    body: { action: 'merge', state: pcInitialState }
  });
  assert.strictEqual(pcSaveRes.statusCode, 200);
  assert.strictEqual(pcSaveRes.data.ok, true);
  assert.strictEqual(pcSaveRes.data.state.dailySales.length, 1);
  assert.strictEqual(pcSaveRes.data.state.commonCosts.length, 1);

  // 7-3. 잘못된 비밀번호로 모바일에서 접근 시 차단 확인
  const wrongPwRes = await callApi({
    method: 'GET',
    auth: { userId: testUser.userId, password: 'wrong_password' }
  });
  assert.strictEqual(wrongPwRes.statusCode, 401);
  assert.strictEqual(wrongPwRes.data.code, 'INVALID_CREDENTIALS');

  // 7-4. 모바일에서 같은 계정으로 접속하여 PC 저장 내역 조회 확인
  const mobileFetchRes = await callApi({
    method: 'GET',
    auth: testUser
  });
  assert.strictEqual(mobileFetchRes.statusCode, 200);
  assert.strictEqual(mobileFetchRes.data.state.dailySales.length, 1);
  assert.strictEqual(mobileFetchRes.data.state.dailySales[0].note, 'PC저장분');
  assert.strictEqual(mobileFetchRes.data.state.commonCosts[0].detail, '배추 모종');

  // 7-5. 모바일에서 새로운 매출·공통비용 추가 후 저장 -> PC에서 재조회 시 반영 확인
  const mobileUpdatedState = {
    ...mobileFetchRes.data.state,
    dailySales: [
      ...mobileFetchRes.data.state.dailySales,
      { id: 102, date: '2026-11-28', channel: '공판장', ordersUnder10: 30, ordersOver10: 40, sales: 2100000, extraCost: 0, note: '모바일추가분' }
    ],
    commonCosts: [
      ...mobileFetchRes.data.state.commonCosts,
      { id: 202, date: '2026-11-28', category: '기타', detail: '작업복 및 장갑 구입', amount: 85000, note: '모바일공통비' }
    ]
  };
  const mobileSaveRes = await callApi({
    method: 'POST',
    auth: testUser,
    body: { action: 'save', state: mobileUpdatedState }
  });
  assert.strictEqual(mobileSaveRes.statusCode, 200);

  const pcRefetchRes = await callApi({
    method: 'GET',
    auth: testUser
  });
  assert.strictEqual(pcRefetchRes.statusCode, 200);
  assert.strictEqual(pcRefetchRes.data.state.dailySales.length, 2);
  assert.strictEqual(pcRefetchRes.data.state.commonCosts.length, 2);
  assert.strictEqual(pcRefetchRes.data.state.dailySales[1].note, '모바일추가분');
  assert.strictEqual(pcRefetchRes.data.state.commonCosts[1].detail, '작업복 및 장갑 구입');

  // 7-6. 별도 로그인 없이도 기본 자동동기화 계정(DEFAULT_CLOUD_CREDS)으로 PC 저장분과 모바일 저장분이 서로 덮어쓰지 않고 합쳐지는지 검증
  assert(jsContent.includes('DEFAULT_CLOUD_CREDS') && jsContent.includes('hasLocalUnpushedDifference'), '별도 로그인 없이도 PC·모바일 자동동기화 및 기존 데이터 병합이 작동해야 합니다.');
  const defaultAuth = { userId: 'shared_workspace_v1', password: 'shared_workspace_pw_v1' };
  await callApi({
    method: 'POST',
    auth: defaultAuth,
    body: {
      action: 'merge',
      state: {
        channels: CHANNELS,
        targetMonths: TARGET_MONTHS.map(t => t.ym === '2026-11' ? { ...t, target: 3000000 } : t),
        channelRates: mockRates,
        dailySales: [{ id: 501, date: '2026-11-10', channel: '농라', ordersUnder10: 5, ordersOver10: 2, sales: 250000, extraCost: 0, note: 'PC단독저장' }],
        commonCosts: []
      }
    }
  });
  const mergedDefaultRes = await callApi({
    method: 'POST',
    auth: defaultAuth,
    body: {
      action: 'merge',
      state: {
        channels: CHANNELS,
        targetMonths: TARGET_MONTHS.map(t => t.ym === '2026-12' ? { ...t, target: 4000000 } : t),
        channelRates: mockRates,
        dailySales: [{ id: 502, date: '2026-11-11', channel: '카톡', ordersUnder10: 3, ordersOver10: 1, sales: 140000, extraCost: 0, note: '핸드폰단독저장' }],
        commonCosts: [{ id: 601, date: '2026-11-11', category: '농약비', detail: '친환경 방제제', amount: 45000, note: '핸드폰공통비' }]
      }
    }
  });
  assert.strictEqual(mergedDefaultRes.statusCode, 200);
  assert.strictEqual(mergedDefaultRes.data.state.dailySales.length, 2, 'PC에서 저장한 매출과 핸드폰에서 저장한 매출이 모두 합쳐져 보여야 합니다.');
  assert.strictEqual(mergedDefaultRes.data.state.commonCosts.length, 1, '핸드폰에서 저장한 공통비용이 보존되어야 합니다.');
  assert.strictEqual(mergedDefaultRes.data.state.targetMonths.find(t => t.ym === '2026-11').target, 3000000);
  assert.strictEqual(mergedDefaultRes.data.state.targetMonths.find(t => t.ym === '2026-12').target, 4000000);

  // 7-7. 판매채널 등록관리 채널 삭제 기능 및 3. 판매채널 기간별비용설정 / 4. 판매채널별손익분석 즉시 연동 검증
  assert(jsContent.includes('window.deleteChannelMaster') && jsContent.includes('btn-channel-del'), '판매채널 등록관리에서 채널 삭제 버튼과 삭제 함수가 존재해야 합니다.');
  assert(jsContent.includes('state.channelRates = state.channelRates.filter(r => r.channel !== targetName)'), '채널 삭제 시 3. 판매채널 기간별비용설정에 바로 반영되어야 합니다.');
  assert(jsContent.includes('!state.channelRates.some(r => r.channel === name)'), '채널 등록 시 3. 판매채널 기간별비용설정에 기본 설정 행이 바로 추가되어야 합니다.');

  const channelDelTestUser = { userId: `test_ch_del_${Date.now()}`, password: 'correct_password_123' };
  const afterDeleteState = {
    channels: ['농라', '스마트스토어', '쿠팡'],
    deletedChannels: ['카톡', '고흥몰', '남도장터', '로컬푸드', '공판장', '기타'],
    targetMonths: TARGET_MONTHS,
    channelRates: [
      { id: 1, channel: '농라', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
      { id: 2, channel: '스마트스토어', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' },
      { id: 9, channel: '쿠팡', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, note: '' }
    ],
    dailySales: [],
    commonCosts: []
  };
  const chSaveRes = await callApi({
    method: 'POST',
    auth: channelDelTestUser,
    body: { action: 'save', state: afterDeleteState }
  });
  assert.strictEqual(chSaveRes.statusCode, 200);
  assert.deepStrictEqual(chSaveRes.data.state.channels, ['농라', '스마트스토어', '쿠팡'], '삭제한 기본 채널이 다시 복원되지 않고 등록/남아있는 채널만 정확히 유지되어야 합니다.');
  assert.strictEqual(chSaveRes.data.state.channelRates.length, 3, '비용 설정에도 등록된 3개 채널만 정확히 반영되어야 합니다.');

  // 테스트용 임시 로컬 파일 정리
  const localStoreDir = path.join(__dirname, '.local_cloud_store');
  if (fs.existsSync(localStoreDir)) {
    fs.rmSync(localStoreDir, { recursive: true, force: true });
  }

  console.log('✅ 7. PC ↔ 모바일 동일 계정 실시간 동기화, 비공개 인증 보호, 기존 데이터 병합 보존 검증 완료');
  console.log('=== [모든 검증 테스트 100% 통과!] ===');
}

verifyCloudSyncApi().catch(err => {
  console.error(err);
  process.exit(1);
});
