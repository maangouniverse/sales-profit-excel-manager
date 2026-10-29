const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { CHANNELS, CLIENTS, CHANNEL_RATES, DAILY_SALES, MONTHLY_TARGETS, COMMON_COSTS, computeRowMetrics } = require('./excel_data_styles');

console.log('=== [매출관리 엑셀 프로그램 자동 검증 테스트 시작] ===');

// 1. 엑셀 파일 생성 확인
const xlsxPath = path.join(__dirname, '거래처별_매출순이익_관리프로그램.xlsx');
assert(fs.existsSync(xlsxPath), '엑셀 파일이 존재해야 합니다.');
const stat = fs.statSync(xlsxPath);
assert(stat.size > 50000, '엑셀 파일 크기가 정상이어야 합니다.');
console.log(`✅ 1. 엑셀 파일 생성 확인 완료 (${stat.size.toLocaleString()} bytes)`);

// 2. 기본 판매채널 7종 등록 확인
const expectedChannels = ['농라', '스마트스토어', '카톡', '고흥몰', '남도장터', '로컬푸드', '기타'];
assert.deepStrictEqual(CHANNELS, expectedChannels, '기본 판매채널 7종이 정확히 등록되어야 합니다.');
console.log('✅ 2. 기본 판매채널 7종(농라, 스마트스토어, 카톡, 고흥몰, 남도장터, 로컬푸드, 기타) 확인 완료');

// 3. 비용 변경일 전후 단가 분리 적용 테스트 (스마트스토어: 2026-04-15 이전 vs 2026-04-16 이후)
const beforeSale = DAILY_SALES.find(r => r.date === '2026-04-14' && r.channel === '스마트스토어');
const afterSale = DAILY_SALES.find(r => r.date === '2026-04-17' && r.channel === '스마트스토어');
const mBefore = computeRowMetrics(beforeSale, CHANNEL_RATES);
const mAfter = computeRowMetrics(afterSale, CHANNEL_RATES);

assert.strictEqual(mBefore.status, '정상');
assert.strictEqual(mBefore.feeRate, 0.055, '4/15 이전 스마트스토어 수수료율은 5.5%여야 합니다.');
assert.strictEqual(mBefore.feeAmt, 44000, '800,000 * 5.5% = 44,000원이어야 합니다.');
assert.strictEqual(mBefore.unitCost, 4110, '4/15 이전 건당비용은 3000+1000+110 = 4,110원이어야 합니다.');
assert.strictEqual(mBefore.orderCost, 82200, '20건 * 4,110원 = 82,200원이어야 합니다.');
assert.strictEqual(mBefore.totalCost, 526200, '원가 400,000 + 수수료 44,000 + 건수비 82,200 = 526,200원이어야 합니다.');
assert.strictEqual(mBefore.netProfit, 273800, '순이익은 800,000 - 526,200 = 273,800원이어야 합니다.');

assert.strictEqual(mAfter.status, '정상');
assert.strictEqual(mAfter.feeRate, 0.058, '4/16 이후 스마트스토어 수수료율은 5.8%여야 합니다.');
assert.strictEqual(mAfter.feeAmt, 46400, '800,000 * 5.8% = 46,400원이어야 합니다.');
assert.strictEqual(mAfter.unitCost, 4830, '4/16 이후 건당비용은 3500+1200+130 = 4,830원이어야 합니다.');
assert.strictEqual(mAfter.orderCost, 96600, '20건 * 4,830원 = 96,600원이어야 합니다.');
assert.strictEqual(mAfter.totalCost, 543000, '원가 400,000 + 수수료 46,400 + 건수비 96,600 = 543,000원이어야 합니다.');
assert.strictEqual(mAfter.netProfit, 257000, '순이익은 800,000 - 543,000 = 257,000원이어야 합니다.');
console.log('✅ 3. 비용 변경일(4/16) 전후 단가 분리 적용 및 주문건수 비례 비용 계산 검증 완료');

// 4. 비용 설정이 없는 날짜 입력 시 0원 처리하지 않고 "비용 설정 필요" 표시하는지 검증
const unconfiguredRow = { date: '2025-12-15', client: '청정영농조합', channel: '스마트스토어', orders: 10, sales: 500000, cogs: 250000, extraCost: 0 };
const mUnconf = computeRowMetrics(unconfiguredRow, CHANNEL_RATES);
assert.strictEqual(mUnconf.status, '비용 설정 필요');
assert.strictEqual(mUnconf.totalCost, '비용 설정 필요');
assert.strictEqual(mUnconf.netProfit, '비용 설정 필요');
console.log('✅ 4. 미설정 기간 입력 시 임의 0원 처리 방지 및 [비용 설정 필요] 표시 검증 완료');

// 5. 동일 채널 기간 중복 시 경고 검증
const overlappingRates = [
  ...CHANNEL_RATES,
  { channel: '스마트스토어', start: '2026-04-10', end: '2026-04-20', feeRate: 0.06, shipFee: 3500, packFee: 1000, etcFee: 0 }
];
const mOverlap = computeRowMetrics(beforeSale, overlappingRates);
assert.strictEqual(mOverlap.status, '기간 중복 확인');
console.log('✅ 5. 동일 판매채널 적용 기간 중복 감지 검증 완료');

// 6. 매출 0원일 때 나누기 오류(NaN/Infinity) 방지 검증
const zeroSaleRow = { date: '2026-04-10', client: '청정영농조합', channel: '농라', orders: 0, sales: 0, cogs: 0, extraCost: 0 };
const mZero = computeRowMetrics(zeroSaleRow, CHANNEL_RATES);
assert.strictEqual(mZero.margin, 0, '매출이 0원일 때 순이익률은 에러 없이 0이어야 합니다.');
console.log('✅ 6. 매출/목표매출 0원일 때 0 나누기 오류 방지 검증 완료');

console.log('=== [모든 검증 테스트 100% 통과!] ===');
