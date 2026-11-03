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
console.log('✅ 3. 10kg 미만/이상 택배비 구분 및 박스·포장비닐·리플렛 비용 합산, 단가 변경 전후 수식 검증 완료');

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

console.log('=== [모든 검증 테스트 100% 통과!] ===');
