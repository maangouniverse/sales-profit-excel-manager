const { dateToSerial } = require('./xlsx_writer');

// 1. 기본 판매채널 7종
const CHANNELS = [
  '농라',
  '스마트스토어',
  '카톡',
  '고흥몰',
  '남도장터',
  '로컬푸드',
  '기타'
];

// 2. 목표설정 기간: 2026년 11월 ~ 2027년 10월 (12개월)
const TARGET_MONTHS = [
  { idx: 1, ym: '2026-11', label: '26년 11월', year: 2026, month: 11, start: '2026-11-01', end: '2026-11-30', days: 30, target: 12000000 },
  { idx: 2, ym: '2026-12', label: '26년 12월', year: 2026, month: 12, start: '2026-12-01', end: '2026-12-31', days: 31, target: 15000000 },
  { idx: 3, ym: '2027-01', label: '27년 1월',  year: 2027, month: 1,  start: '2027-01-01', end: '2027-01-31', days: 31, target: 18000000 },
  { idx: 4, ym: '2027-02', label: '27년 2월',  year: 2027, month: 2,  start: '2027-02-01', end: '2027-02-28', days: 28, target: 14000000 },
  { idx: 5, ym: '2027-03', label: '27년 3월',  year: 2027, month: 3,  start: '2027-03-01', end: '2027-03-31', days: 31, target: 10000000 },
  { idx: 6, ym: '2027-04', label: '27년 4월',  year: 2027, month: 4,  start: '2027-04-01', end: '2027-04-30', days: 30, target: 9000000 },
  { idx: 7, ym: '2027-05', label: '27년 5월',  year: 2027, month: 5,  start: '2027-05-01', end: '2027-05-31', days: 31, target: 11000000 },
  { idx: 8, ym: '2027-06', label: '27년 6월',  year: 2027, month: 6,  start: '2027-06-01', end: '2027-06-30', days: 30, target: 8000000 },
  { idx: 9, ym: '2027-07', label: '27년 7월',  year: 2027, month: 7,  start: '2027-07-01', end: '2027-07-31', days: 31, target: 7000000 },
  { idx: 10, ym: '2027-08', label: '27년 8월', year: 2027, month: 8,  start: '2027-08-01', end: '2027-08-31', days: 31, target: 9000000 },
  { idx: 11, ym: '2027-09', label: '27년 9월', year: 2027, month: 9,  start: '2027-09-01', end: '2027-09-30', days: 30, target: 16000000 },
  { idx: 12, ym: '2027-10', label: '27년 10월', year: 2027, month: 10, start: '2027-10-01', end: '2027-10-31', days: 31, target: 11000000 }
];

// 3. 판매채널·기간별 비용 설정 (10kg 미만 배송비 / 10kg 이상 배송비 구분 + 박스, 포장비닐, 리플렛, 기타비용)
// 스마트스토어 단가 변경 이력 예시 포함 (2026-11-15 이전 vs 2026-11-16 이후)
const CHANNEL_RATES = [
  {
    channel: '농라',
    start: '2026-11-01',
    end: '',
    feeRate: 0.00,
    shipUnder10: 3500,
    shipOver10: 4500,
    boxCost: 650,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[예시] 농라 카페 직거래 (수수료 0%, 10kg미만 3,500원 / 10kg이상 4,500원)'
  },
  {
    channel: '스마트스토어',
    start: '2026-11-01',
    end: '2026-11-15',
    feeRate: 0.055,
    shipUnder10: 3200,
    shipOver10: 4200,
    boxCost: 650,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[단가변경 전] 11/15까지 적용 단가 (수수료 5.5%, 미만택배 3,200원/이상 4,200원)'
  },
  {
    channel: '스마트스토어',
    start: '2026-11-16',
    end: '',
    feeRate: 0.058,
    shipUnder10: 3500,
    shipOver10: 4500,
    boxCost: 750,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[단가변경 후] 11/16부터 택배비·박스비 인상 적용 (과거 매출은 이전 단가 유지)'
  },
  {
    channel: '카톡',
    start: '2026-11-01',
    end: '',
    feeRate: 0.033,
    shipUnder10: 3500,
    shipOver10: 4500,
    boxCost: 650,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[예시] 카카오톡 톡채널/선물 주문'
  },
  {
    channel: '고흥몰',
    start: '2026-11-01',
    end: '',
    feeRate: 0.040,
    shipUnder10: 3000,
    shipOver10: 4000,
    boxCost: 650,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[예시] 고흥몰 입점 특산물 판매'
  },
  {
    channel: '남도장터',
    start: '2026-11-01',
    end: '',
    feeRate: 0.050,
    shipUnder10: 3200,
    shipOver10: 4200,
    boxCost: 650,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[예시] 남도장터 온라인몰 판매'
  },
  {
    channel: '로컬푸드',
    start: '2026-11-01',
    end: '',
    feeRate: 0.100,
    shipUnder10: 0,
    shipOver10: 0,
    boxCost: 400,
    vinylCost: 120,
    leafletCost: 80,
    etcCost: 0,
    note: '[예시] 로컬푸드 직매장 납품 (택배비 없음, 소포장 박스·비닐·리플렛만 적용)'
  },
  {
    channel: '기타',
    start: '2026-11-01',
    end: '',
    feeRate: 0.030,
    shipUnder10: 3500,
    shipOver10: 4500,
    boxCost: 650,
    vinylCost: 150,
    leafletCost: 100,
    etcCost: 0,
    note: '[예시] 전화/방문 및 기타 외부채널 기본값'
  }
];

// 4. 일일 매출 입력 예시 데이터 (2026년 11월 ~ 2027년 2월, 거래처/상품원가 삭제 & 10kg 미만/이상 주문건수 구분)
const DAILY_SALES = [
  // 2026년 11월 (집중 분석 월: 12건, 단가 변경일 11/15 전후 비교 및 적자 예시 포함)
  { date: '2026-11-02', channel: '농라', ordersUnder10: 24, ordersOver10: 12, sales: 1480000, extraCost: 0, note: '[예시] 11월 첫째주 햇유자·특산물 공동구매' },
  { date: '2026-11-04', channel: '고흥몰', ordersUnder10: 30, ordersOver10: 10, sales: 1620000, extraCost: 0, note: '[예시] 고흥몰 가을 수확 기획전' },
  { date: '2026-11-06', channel: '카톡', ordersUnder10: 18, ordersOver10: 7, sales: 980000, extraCost: 0, note: '[예시] 단골 고객 카톡 알림톡 주문' },
  { date: '2026-11-08', channel: '남도장터', ordersUnder10: 22, ordersOver10: 14, sales: 1540000, extraCost: 0, note: '[예시] 남도장터 주말 할인전' },
  { date: '2026-11-11', channel: '로컬푸드', ordersUnder10: 45, ordersOver10: 5, sales: 1150000, extraCost: 0, note: '[예시] 로컬푸드 직매장 출하' },
  { date: '2026-11-14', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0, note: '[단가변경 전 확인] 11/15 이전: 수수료 5.5%, 미만 4,200원 / 이상 5,200원 적용' },
  { date: '2026-11-17', channel: '스마트스토어', ordersUnder10: 20, ordersOver10: 10, sales: 1200000, extraCost: 0, note: '[단가변경 후 확인] 11/16 이후: 수수료 5.8%, 미만 4,600원 / 이상 5,600원 자동 적용!' },
  { date: '2026-11-20', channel: '농라', ordersUnder10: 28, ordersOver10: 15, sales: 1850000, extraCost: 15000, note: '[예시] 대용량 10kg 이상 주문 급증 (완충재 추가비용 반영)' },
  { date: '2026-11-22', channel: '기타', ordersUnder10: 12, ordersOver10: 8, sales: 750000, extraCost: 0, note: '[예시] 전화 주문 단체 발송' },
  { date: '2026-11-25', channel: '기타', ordersUnder10: 15, ordersOver10: 10, sales: 95000, extraCost: 35000, note: '[적자 구분 예시] 파격 샘플행사 및 파손 재발송비로 인한 일시 적자' },
  { date: '2026-11-27', channel: '고흥몰', ordersUnder10: 25, ordersOver10: 12, sales: 1520000, extraCost: 0, note: '[예시] 고흥몰 베스트 상품전' },
  { date: '2026-11-29', channel: '카톡', ordersUnder10: 20, ordersOver10: 8, sales: 1120000, extraCost: 0, note: '[예시] 월말 카톡 재구매 주문' },

  // 2026년 12월
  { date: '2026-12-05', channel: '스마트스토어', ordersUnder10: 45, ordersOver10: 20, sales: 2850000, extraCost: 0, note: '[예시] 연말 선물세트 주문' },
  { date: '2026-12-12', channel: '농라', ordersUnder10: 50, ordersOver10: 25, sales: 3400000, extraCost: 0, note: '[예시] 겨울철 인기 특산물 공구' },
  { date: '2026-12-20', channel: '남도장터', ordersUnder10: 38, ordersOver10: 18, sales: 2420000, extraCost: 0, note: '[예시] 연말 감사 기획전' },
  { date: '2026-12-28', channel: '고흥몰', ordersUnder10: 40, ordersOver10: 15, sales: 2380000, extraCost: 0, note: '[예시] 해맞이 특산품 행사' },

  // 2027년 1월
  { date: '2027-01-08', channel: '스마트스토어', ordersUnder10: 55, ordersOver10: 30, sales: 3950000, extraCost: 0, note: '[예시] 설 명절 사전예약 1차' },
  { date: '2027-01-16', channel: '농라', ordersUnder10: 60, ordersOver10: 35, sales: 4500000, extraCost: 20000, note: '[예시] 설 명절 대량주문' },
  { date: '2027-01-24', channel: '카톡', ordersUnder10: 42, ordersOver10: 22, sales: 2980000, extraCost: 0, note: '[예시] 카톡 명절 선물세트' },
  { date: '2027-01-29', channel: '로컬푸드', ordersUnder10: 70, ordersOver10: 15, sales: 2150000, extraCost: 0, note: '[예시] 명절 직매장 특판' },

  // 2027년 2월
  { date: '2027-02-06', channel: '고흥몰', ordersUnder10: 48, ordersOver10: 20, sales: 3100000, extraCost: 0, note: '[예시] 정월대보름 기획전' },
  { date: '2027-02-15', channel: '남도장터', ordersUnder10: 35, ordersOver10: 15, sales: 2250000, extraCost: 0, note: '[예시] 2월 정기 프로모션' },
  { date: '2027-02-23', channel: '스마트스토어', ordersUnder10: 40, ordersOver10: 12, sales: 2320000, extraCost: 0, note: '[예시] 봄맞이 사전 특가' }
];

// 5. 월별 공통비용 (2026년 11월 ~ 2027년 2월 예시)
const COMMON_COSTS = [
  { date: '2026-11-25', category: '월 광고선전비', amount: 300000, note: '[예시] 11월 키워드 및 배너 광고비' },
  { date: '2026-11-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 11월 저온창고 임대료' },
  { date: '2026-11-28', category: '공과금/통신/프로그램', amount: 100000, note: '[예시] 11월 전기료 및 송장프로그램' },
  { date: '2026-12-25', category: '월 광고선전비', amount: 350000, note: '[예시] 12월 연말 선물세트 광고비' },
  { date: '2026-12-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 12월 저온창고 임대료' },
  { date: '2027-01-25', category: '월 광고선전비', amount: 400000, note: '[예시] 1월 설 명절 집중 광고비' },
  { date: '2027-01-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 1월 저온창고 임대료' },
  { date: '2027-02-25', category: '월 광고선전비', amount: 250000, note: '[예시] 2월 기본 검색광고비' },
  { date: '2027-02-25', category: '작업장/창고 임대료', amount: 400000, note: '[예시] 2월 저온창고 임대료' }
];

// 일일 매출 행 계산 함수 (10kg 미만 / 10kg 이상 구분 + 박스/비닐/리플렛 반영)
function computeRowMetrics(row, rates) {
  const dSerial = dateToSerial(row.date);
  const matched = rates.filter(r => {
    if (r.channel !== row.channel) return false;
    const s = dateToSerial(r.start);
    const e = r.end ? dateToSerial(r.end) : dateToSerial('2099-12-31');
    return dSerial >= s && dSerial <= e;
  });

  const totalOrders = (Number(row.ordersUnder10) || 0) + (Number(row.ordersOver10) || 0);

  if (matched.length === 0) {
    return {
      totalOrders,
      status: '비용 설정 필요',
      feeRate: '비용 설정 필요',
      feeAmt: '비용 설정 필요',
      unitUnder10: '비용 설정 필요',
      unitOver10: '비용 설정 필요',
      orderCost: '비용 설정 필요',
      totalCost: '비용 설정 필요',
      netProfit: '비용 설정 필요',
      margin: '비용 설정 필요'
    };
  }
  if (matched.length > 1) {
    return {
      totalOrders,
      status: '기간 중복 확인',
      feeRate: '비용 설정 필요',
      feeAmt: '비용 설정 필요',
      unitUnder10: '비용 설정 필요',
      unitOver10: '비용 설정 필요',
      orderCost: '비용 설정 필요',
      totalCost: '비용 설정 필요',
      netProfit: '비용 설정 필요',
      margin: '비용 설정 필요'
    };
  }

  const rate = matched[0];
  const feeRate = rate.feeRate;
  const feeAmt = Math.round(row.sales * feeRate);
  const packSum = (rate.boxCost || 0) + (rate.vinylCost || 0) + (rate.leafletCost || 0) + (rate.etcCost || 0);
  const unitUnder10 = (rate.shipUnder10 || 0) + packSum;
  const unitOver10 = (rate.shipOver10 || 0) + packSum;
  const orderCost = (row.ordersUnder10 * unitUnder10) + (row.ordersOver10 * unitOver10);
  const totalCost = feeAmt + orderCost + (row.extraCost || 0);
  const netProfit = row.sales - totalCost;
  const margin = row.sales > 0 ? netProfit / row.sales : 0;

  return {
    totalOrders,
    status: '정상',
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

// 첨부 이미지(매출성과.png, 실적자동계산1~3.png) 스타일을 반영한 styles.xml
// 차콜 그레이 헤더(#4A4E54 / #383B40), 깔끔한 연회색 가로선, 파스텔 블루/올리브 포인트 반영
function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="0.0%"/>
    <numFmt numFmtId="166" formatCode="0.00%"/>
    <numFmt numFmtId="167" formatCode="yyyy-mm-dd (aaa)"/>
  </numFmts>
  <fonts count="9">
    <font><sz val="10"/><color rgb="FF222222"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="15"/><color rgb="FF111111"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="11"/><color rgb="FF222222"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><color rgb="FF111111"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="9"/><color rgb="FF666666"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="14"/><color rgb="FF111111"/><name val="맑은 고딕"/></font>
    <font><sz val="9.5"/><color rgb="FF333333"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><color rgb="FF1D4ED8"/><name val="맑은 고딕"/></font>
  </fonts>
  <fills count="12">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <!-- 2: Charcoal Dark Header (#4A4E54 from reference image) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF4A4E54"/></patternFill></fill>
    <!-- 3: Deep Charcoal Subheader (#383B40 from reference image) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FF383B40"/></patternFill></fill>
    <!-- 4: Soft Yellow Input (#FFFBEB) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFBEB"/></patternFill></fill>
    <!-- 5: Clean Light Gray Calc (#F8F9FA) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8F9FA"/></patternFill></fill>
    <!-- 6: Soft Blue Point Calc (#EEF4FF) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFEEF4FF"/></patternFill></fill>
    <!-- 7: Summary Row Gray (#F1F3F5) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F3F5"/></patternFill></fill>
    <!-- 8: Card Label Fill (#F8F9FA) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8F9FA"/></patternFill></fill>
    <!-- 9: Pure White (#FFFFFF) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
    <!-- 10: Guide Box Soft Gray/Blue (#F4F6F9) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFF4F6F9"/></patternFill></fill>
    <!-- 11: Status OK Soft Green (#ECFDF5) -->
    <fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFE5E7EB"/></left>
      <right style="thin"><color rgb="FFE5E7EB"/></right>
      <top style="thin"><color rgb="FFE5E7EB"/></top>
      <bottom style="thin"><color rgb="FFE5E7EB"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFD1D5DB"/></left>
      <right style="thin"><color rgb="FFD1D5DB"/></right>
      <top style="medium"><color rgb="FF383B40"/></top>
      <bottom style="double"><color rgb="FF383B40"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="28">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="167" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="167" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="4" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="166" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="7" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="4" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="4" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="5" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="6" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="6" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="7" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="8" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <dxfs count="2">
    <dxf>
      <font><b/><color rgb="FFDC2626"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
    </dxf>
    <dxf>
      <font><b/><color rgb="FFB91C1C"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FFFECACA"/></patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>`;
}

module.exports = {
  CHANNELS,
  TARGET_MONTHS,
  CHANNEL_RATES,
  DAILY_SALES,
  COMMON_COSTS,
  computeRowMetrics,
  buildStylesXml
};
