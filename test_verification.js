const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { CHANNELS, COMMON_COST_CATEGORIES, TARGET_MONTHS, CHANNEL_RATES, DAILY_SALES, COMMON_COSTS, computeRowMetrics } = require('./excel_data_styles');

console.log('=== [수정사항 반영 매출관리 프로그램 자동 검증 테스트] ===');

// 1. 엑셀 파일 확인
const xlsxPath = path.join(__dirname, '거래처별_매출순이익_관리프로그램.xlsx');
assert(fs.existsSync(xlsxPath), '엑셀 파일이 존재해야 합니다.');
console.log(`✅ 1. 엑셀 파일 생성 확인 완료 (${fs.statSync(xlsxPath).size.toLocaleString()} bytes)`);

// 2. 목표매출 기간: 2026년 11월 ~ 2027년 10월 (12개월) 및 총합계 확인
assert.strictEqual(TARGET_MONTHS.length, 12);
assert.strictEqual(TARGET_MONTHS[0].ym, '2026-11');
assert.strictEqual(TARGET_MONTHS[11].ym, '2027-10');
const totalTargetSum = TARGET_MONTHS.reduce((a, b) => a + b.target, 0);
assert.strictEqual(totalTargetSum, 140000000, '12개월 목표매출 총합계는 140,000,000원이어야 합니다.');
console.log(`✅ 2. 목표설정 기간(2026년 11월 ~ 2027년 10월) 및 목표매출 총합계(${totalTargetSum.toLocaleString()}원) 확인 완료`);

// 3. 10kg 미만 / 10kg 이상 주문건수 구분 및 박스·포장비닐·리플렛 비용 합산 검증
// 스마트스토어 단가 변경일 전(2026-11-14) vs 후(2026-11-17) 비교
const beforeSale = DAILY_SALES.find(r => r.date === '2026-11-14' && r.channel === '스마트스토어');
const afterSale = DAILY_SALES.find(r => r.date === '2026-11-17' && r.channel === '스마트스토어');
const mBefore = computeRowMetrics(beforeSale, CHANNEL_RATES);
const mAfter = computeRowMetrics(afterSale, CHANNEL_RATES);

// 11/14 (단가 변경 전):
// 수수료율 5.5% -> 1,200,000 * 5.5% = 66,000원
// 포장부자재 = 박스(650) + 비닐(150) + 리플렛(100) = 900원
// 10kg 미만 건당 = 3,200 + 900 = 4,100원 (20건 = 82,000원)
// 10kg 이상 건당 = 4,200 + 900 = 5,100원 (10건 = 51,000원)
// 주문건수 비용 합계 = 82,000 + 51,000 = 133,000원
// 총비용 = 66,000 + 133,000 = 199,000원 -> 순이익 = 1,200,000 - 199,000 = 1,001,000원
assert.strictEqual(mBefore.status, '정상');
assert.strictEqual(mBefore.totalOrders, 30);
assert.strictEqual(mBefore.feeAmt, 66000);
assert.strictEqual(mBefore.unitUnder10, 4100);
assert.strictEqual(mBefore.unitOver10, 5100);
assert.strictEqual(mBefore.orderCost, 133000);
assert.strictEqual(mBefore.totalCost, 199000);
assert.strictEqual(mBefore.netProfit, 1001000);

// 11/17 (단가 변경 후):
// 수수료율 5.8% -> 1,200,000 * 5.8% = 69,600원
// 포장부자재 = 박스(750) + 비닐(150) + 리플렛(100) = 1,000원
// 10kg 미만 건당 = 3,500 + 1,000 = 4,500원 (20건 = 90,000원)
// 10kg 이상 건당 = 4,500 + 1,000 = 5,500원 (10건 = 55,000원)
// 주문건수 비용 합계 = 90,000 + 55,000 = 145,000원
// 총비용 = 69,600 + 145,000 = 214,600원 -> 순이익 = 1,200,000 - 214,600 = 985,400원
assert.strictEqual(mAfter.status, '정상');
assert.strictEqual(mAfter.totalOrders, 30);
assert.strictEqual(mAfter.feeAmt, 69600);
assert.strictEqual(mAfter.unitUnder10, 4500);
assert.strictEqual(mAfter.unitOver10, 5500);
assert.strictEqual(mAfter.orderCost, 145000);
assert.strictEqual(mAfter.totalCost, 214600);
assert.strictEqual(mAfter.netProfit, 985400);
console.log('✅ 3. 10kg 미만/이상 택배비 구분 및 박스·포장비닐·리플렛 비용 합산, 단가 변경 전후 검증 완료');

// 4. 판매채널 '공판장' 추가 및 계산 검증
assert(CHANNELS.includes('공판장'), '판매채널 목록에 공판장이 포함되어야 합니다.');
assert.strictEqual(CHANNELS.length, 8);
const gongSale = DAILY_SALES.find(r => r.date === '2026-11-28' && r.channel === '공판장');
const mGong = computeRowMetrics(gongSale, CHANNEL_RATES);
// 공판장: 수수료 7% (2,100,000 * 7% = 147,000원)
// 10kg미만 단가 = 1,000 + 650 + 100 + 0 = 1,750원 (30건 = 52,500원)
// 10kg이상 단가 = 1,500 + 650 + 100 + 0 = 2,250원 (40건 = 90,000원)
// 건수비용 합계 = 142,500원, 총비용 = 289,500원, 순이익 = 1,810,500원
assert.strictEqual(mGong.status, '정상');
assert.strictEqual(mGong.totalOrders, 70);
assert.strictEqual(mGong.feeAmt, 147000);
assert.strictEqual(mGong.unitUnder10, 1750);
assert.strictEqual(mGong.unitOver10, 2250);
assert.strictEqual(mGong.orderCost, 142500);
assert.strictEqual(mGong.totalCost, 289500);
assert.strictEqual(mGong.netProfit, 1810500);
console.log('✅ 4. 판매채널 [공판장] 등록 및 무게별 비용·순이익 자동계산 검증 완료');

// 5. 공통비용 항목별 집계 및 그래프 데이터 검증
assert.strictEqual(COMMON_COST_CATEGORIES.length, 5);
const novCommonSum = COMMON_COSTS.filter(c => c.date >= '2026-11-01' && c.date <= '2026-11-30').reduce((a, b) => a + b.amount, 0);
assert.strictEqual(novCommonSum, 1250000);
console.log(`✅ 5. 공통비용 5대 항목 및 11월 공통비용 집계(${novCommonSum.toLocaleString()}원) 검증 완료`);

console.log('=== [모든 검증 테스트 100% 통과!] ===');
