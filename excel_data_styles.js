const { dateToSerial } = require('./xlsx_writer');

// 1. 기본 판매채널 7종 (요구사항 3번)
const CHANNELS = [
  '농라',
  '스마트스토어',
  '카톡',
  '고흥몰',
  '남도장터',
  '로컬푸드',
  '기타'
];

// 2. 기본 거래처 목록 8곳 (예시)
const CLIENTS = [
  '청정영농조합',
  '고흥유자농원',
  '남도바다수산',
  '해피푸드유통',
  '초록들녘농산',
  '바른먹거리장터',
  '우리동네직판장',
  '자체직영(소비자)'
];

// 3. 판매채널·기간별 비용 설정 초기 데이터 (시트3)
// 스마트스토어에 2026-04-15 이전 / 2026-04-16 이후 기간 이력 분리 적용 예시 포함!
const CHANNEL_RATES = [
  { channel: '농라', start: '2026-01-01', end: '', feeRate: 0.00, shipFee: 3500, packFee: 800, etcFee: 200, note: '[예시] 농라 카페 직거래 (수수료 0%, 택배·포장비 적용)' },
  { channel: '스마트스토어', start: '2026-01-01', end: '2026-04-15', feeRate: 0.055, shipFee: 3000, packFee: 1000, etcFee: 110, note: '[예시] 4/15까지 적용 단가 (기타비용에 배송비수수료 110원 반영)' },
  { channel: '스마트스토어', start: '2026-04-16', end: '', feeRate: 0.058, shipFee: 3500, packFee: 1200, etcFee: 130, note: '[예시] 4/16부터 택배·포장비 인상 및 수수료율 변경 적용' },
  { channel: '카톡', start: '2026-01-01', end: '', feeRate: 0.033, shipFee: 3500, packFee: 900, etcFee: 100, note: '[예시] 카카오톡 스토어/톡채널 주문' },
  { channel: '고흥몰', start: '2026-01-01', end: '', feeRate: 0.040, shipFee: 3000, packFee: 1000, etcFee: 0, note: '[예시] 고흥몰 입점 특산물 판매' },
  { channel: '남도장터', start: '2026-01-01', end: '', feeRate: 0.050, shipFee: 3200, packFee: 1000, etcFee: 0, note: '[예시] 남도장터 온라인몰 판매' },
  { channel: '로컬푸드', start: '2026-01-01', end: '', feeRate: 0.100, shipFee: 0, packFee: 500, etcFee: 200, note: '[예시] 로컬푸드 직매장 납품 (개별 택배비 없음, 소포장/라벨비만 발생)' },
  { channel: '기타', start: '2026-01-01', end: '', feeRate: 0.030, shipFee: 3500, packFee: 800, etcFee: 0, note: '[예시] 전화주문/외부 제휴몰 기본 단가' }
];

// 4. 일일 매출 입력 예시 데이터 (시트2: 2026년 1월 ~ 4월, 총 24건)
const DAILY_SALES = [
  // 1월 실적
  { date: '2026-01-08', client: '청정영농조합', channel: '스마트스토어', orders: 25, sales: 1125000, cogs: 580000, extraCost: 0, note: '[예시] 신년 선물세트 기획전' },
  { date: '2026-01-15', client: '고흥유자농원', channel: '고흥몰', orders: 30, sales: 1350000, cogs: 690000, extraCost: 0, note: '[예시] 유자청 베스트 상품' },
  { date: '2026-01-22', client: '남도바다수산', channel: '농라', orders: 40, sales: 1920000, cogs: 1150000, extraCost: 15000, note: '[예시] 제철 수산물 공동구매 (아이스팩 추가)' },
  { date: '2026-01-28', client: '우리동네직판장', channel: '로컬푸드', orders: 55, sales: 990000, cogs: 540000, extraCost: 0, note: '[예시] 직매장 주말 납품' },

  // 2월 실적
  { date: '2026-02-05', client: '해피푸드유통', channel: '남도장터', orders: 35, sales: 1680000, cogs: 920000, extraCost: 0, note: '[예시] 설 명절 특선 세트' },
  { date: '2026-02-11', client: '고흥유자농원', channel: '카톡', orders: 28, sales: 1120000, cogs: 560000, extraCost: 0, note: '[예시] 카톡 선물하기 주문' },
  { date: '2026-02-19', client: '초록들녘농산', channel: '스마트스토어', orders: 22, sales: 880000, cogs: 460000, extraCost: 6000, note: '[예시] 단순변심 반품 왕복택배비 1건 반영' },
  { date: '2026-02-25', client: '자체직영(소비자)', channel: '기타', orders: 18, sales: 720000, cogs: 350000, extraCost: 0, note: '[예시] 전화 주문 단골 고객' },

  // 3월 실적
  { date: '2026-03-06', client: '청정영농조합', channel: '농라', orders: 32, sales: 1440000, cogs: 780000, extraCost: 0, note: '[예시] 봄맞이 산지직송' },
  { date: '2026-03-14', client: '남도바다수산', channel: '남도장터', orders: 26, sales: 1300000, cogs: 740000, extraCost: 0, note: '[예시] 건어물 기획전' },
  { date: '2026-03-21', client: '바른먹거리장터', channel: '고흥몰', orders: 20, sales: 760000, cogs: 490000, extraCost: 0, note: '[예시] 친환경 곡물 세트' },
  { date: '2026-03-29', client: '해피푸드유통', channel: '스마트스토어', orders: 38, sales: 1710000, cogs: 890000, extraCost: 0, note: '[예시] 월말 타임특가' },

  // 4월 실적 (집중 분석 월: 총 12건, 단가 변경 전후 비교 + 동일일자 다중 거래처 + 적자 거래처 예시 포함)
  { date: '2026-04-02', client: '청정영농조합', channel: '농라', orders: 30, sales: 1500000, cogs: 820000, extraCost: 0, note: '[예시] 4월 첫째주 공동구매' },
  { date: '2026-04-02', client: '남도바다수산', channel: '농라', orders: 20, sales: 1100000, cogs: 640000, extraCost: 0, note: '[예시] 같은 날짜·같은 채널 다른 거래처 분리 입력 예시' },
  { date: '2026-04-05', client: '고흥유자농원', channel: '고흥몰', orders: 35, sales: 1575000, cogs: 790000, extraCost: 0, note: '[예시] 고흥몰 봄 정기행사' },
  { date: '2026-04-08', client: '해피푸드유통', channel: '카톡', orders: 25, sales: 1050000, cogs: 530000, extraCost: 0, note: '[예시] 톡딜 프로모션' },
  { date: '2026-04-11', client: '초록들녘농산', channel: '남도장터', orders: 28, sales: 1260000, cogs: 680000, extraCost: 0, note: '[예시] 남도장터 주말특가' },
  { date: '2026-04-14', client: '고흥유자농원', channel: '스마트스토어', orders: 20, sales: 800000, cogs: 400000, extraCost: 0, note: '[단가변경 전 확인] 4/15 이전: 수수료 5.5%, 건당비용 4,110원 적용' },
  { date: '2026-04-17', client: '고흥유자농원', channel: '스마트스토어', orders: 20, sales: 800000, cogs: 400000, extraCost: 0, note: '[단가변경 후 확인] 4/16 이후: 수수료 5.8%, 건당비용 4,830원 자동 적용!' },
  { date: '2026-04-19', client: '우리동네직판장', channel: '로컬푸드', orders: 60, sales: 1080000, cogs: 590000, extraCost: 0, note: '[예시] 로컬푸드 직매장 주말 완판' },
  { date: '2026-04-22', client: '바른먹거리장터', channel: '스마트스토어', orders: 25, sales: 450000, cogs: 360000, extraCost: 30000, note: '[적자 구분 예시] 파격 할인특가 및 반품비로 인한 적자 발생' },
  { date: '2026-04-25', client: '자체직영(소비자)', channel: '기타', orders: 15, sales: 675000, cogs: 320000, extraCost: 0, note: '[예시] 기업 단체선물 주문' },
  { date: '2026-04-27', client: '청정영농조합', channel: '카톡', orders: 22, sales: 990000, cogs: 510000, extraCost: 0, note: '[예시] 재구매 알림톡 주문' },
  { date: '2026-04-29', client: '남도바다수산', channel: '고흥몰', orders: 18, sales: 900000, cogs: 490000, extraCost: 7000, note: '[예시] 냉동포장 드라이아이스 추가비용 반영' }
];

// 5. 2026년 월별 목표매출 (시트5)
const MONTHLY_TARGETS = [
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
];

// 6. 월별 공통비용 (광고비, 임대료 등 주문건수와 무관한 고정비 - 시트5)
const COMMON_COSTS = [
  { date: '2026-01-25', category: '월 광고선전비', amount: 250000, note: '[예시] 1월 검색광고 및 카페 배너비' },
  { date: '2026-01-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 1월 저온창고 월세' },
  { date: '2026-02-25', category: '월 광고선전비', amount: 300000, note: '[예시] 2월 설 명절 키워드 광고비' },
  { date: '2026-02-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 2월 저온창고 월세' },
  { date: '2026-03-25', category: '월 광고선전비', amount: 250000, note: '[예시] 3월 쇼핑검색 광고비' },
  { date: '2026-03-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 3월 저온창고 월세' },
  { date: '2026-04-10', category: '월 광고선전비', amount: 350000, note: '[예시] 4월 봄 프로모션 광고비' },
  { date: '2026-04-25', category: '창고/작업장 임대료', amount: 400000, note: '[예시] 4월 저온창고 월세' },
  { date: '2026-04-28', category: '프로그램/통신/공과금', amount: 100000, note: '[예시] 4월 송장출력기 및 인터넷/전기료' }
];

// 일일 매출 행에 대한 사전 계산 함수 (엑셀 캐시 값 v 생성 및 웹 대시보드 겸용)
function computeRowMetrics(row, rates) {
  const dSerial = dateToSerial(row.date);
  const matched = rates.filter(r => {
    if (r.channel !== row.channel) return false;
    const s = dateToSerial(r.start);
    const e = r.end ? dateToSerial(r.end) : dateToSerial('2099-12-31');
    return dSerial >= s && dSerial <= e;
  });

  if (matched.length === 0) {
    return {
      status: '비용 설정 필요',
      feeRate: '비용 설정 필요',
      feeAmt: '비용 설정 필요',
      unitCost: '비용 설정 필요',
      orderCost: '비용 설정 필요',
      totalCost: '비용 설정 필요',
      netProfit: '비용 설정 필요',
      margin: '비용 설정 필요'
    };
  }
  if (matched.length > 1) {
    return {
      status: '기간 중복 확인',
      feeRate: '비용 설정 필요',
      feeAmt: '비용 설정 필요',
      unitCost: '비용 설정 필요',
      orderCost: '비용 설정 필요',
      totalCost: '비용 설정 필요',
      netProfit: '비용 설정 필요',
      margin: '비용 설정 필요'
    };
  }

  const rate = matched[0];
  const feeRate = rate.feeRate;
  const feeAmt = Math.round(row.sales * feeRate);
  const unitCost = rate.shipFee + rate.packFee + rate.etcFee;
  const orderCost = row.orders * unitCost;
  const totalCost = row.cogs + feeAmt + orderCost + row.extraCost;
  const netProfit = row.sales - totalCost;
  const margin = row.sales > 0 ? netProfit / row.sales : 0;

  return {
    status: '정상',
    feeRate,
    feeAmt,
    unitCost,
    orderCost,
    totalCost,
    netProfit,
    margin
  };
}

// 스타일 인덱스(s) 정의 및 styles.xml 생성
// s=0: 기본
// s=1: 대제목 (16pt 굵게 네이비)
// s=2: 소제목/섹션타이틀 (11pt 굵게 네이비)
// s=3: 표 헤더 - 직접입력용 (진한 차콜/네이비 배경 #1E293B, 흰색 굵은 글씨, 가운데 정렬)
// s=4: 표 헤더 - 자동계산용 (진한 블루네이비 배경 #1E3A8A, 흰색 굵은 글씨, 가운데 정렬)
// s=5: 직접 입력 셀 - 텍스트 가운데 (연노랑 배경 #FEF9C3, 테두리)
// s=6: 직접 입력 셀 - 텍스트 왼쪽 (연노랑 배경 #FEF9C3, 테두리)
// s=7: 직접 입력 셀 - 날짜 YYYY-MM-DD (연노랑 배경 #FEF9C3, 가운데 정렬)
// s=8: 직접 입력 셀 - 정수/금액 #,##0 (연노랑 배경 #FEF9C3, 오른쪽 정렬)
// s=9: 직접 입력 셀 - 백분율 0.00% (연노랑 배경 #FEF9C3, 오른쪽 정렬)
// s=10: 자동 계산 셀 - 텍스트 가운데 (연파랑/회색 배경 #F1F5F9, 테두리)
// s=11: 자동 계산 셀 - 날짜 YYYY-MM-DD (연파랑/회색 배경 #F1F5F9, 가운데 정렬)
// s=12: 자동 계산 셀 - 금액/정수 #,##0 (연파랑/회색 배경 #F1F5F9, 오른쪽 정렬)
// s=13: 자동 계산 셀 - 금액 굵게 #,##0 (연파랑 배경 #E0F2FE, 오른쪽 정렬, 굵게)
// s=14: 자동 계산 셀 - 백분율 0.0% (연파랑/회색 배경 #F1F5F9, 오른쪽 정렬)
// s=15: 자동 계산 셀 - 백분율 0.00% (연파랑/회색 배경 #F1F5F9, 오른쪽 정렬)
// s=16: 합계 행 - 텍스트 가운데 (진한 연회색 #E2E8F0, 굵게)
// s=17: 합계 행 - 금액 #,##0 (진한 연회색 #E2E8F0, 오른쪽 정렬, 굵게)
// s=18: 합계 행 - 백분율 0.0% (진한 연회색 #E2E8F0, 오른쪽 정렬, 굵게)
// s=19: KPI 카드 라벨 (상단 작은 회색/네이비 텍스트, #F8FAFC 배경)
// s=20: KPI 카드 값 - 금액 #,##0 (14pt 굵은 네이비 텍스트, #FFFFFF 배경)
// s=21: KPI 카드 값 - 백분율 0.0% (14pt 굵은 블루 텍스트, #FFFFFF 배경)
// s=22: 안내 배너 박스 (연한 블루/민트 배경 #EFF6FF, 왼쪽 정렬, 줄바꿈)
// s=23: 일반 흰색 테두리 셀 - 텍스트 가운데
// s=24: 일반 흰색 테두리 셀 - 텍스트 왼쪽
// s=25: 일반 흰색 테두리 셀 - 금액 #,##0
// s=26: 일반 흰색 테두리 셀 - 백분율 0.0%
// s=27: 점검 배너 정상/경고 표시 셀 (굵게 가운데 정렬 #ECFDF5)

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="4">
    <numFmt numFmtId="164" formatCode="#,##0"/>
    <numFmt numFmtId="165" formatCode="0.0%"/>
    <numFmt numFmtId="166" formatCode="0.00%"/>
    <numFmt numFmtId="167" formatCode="yyyy-mm-dd"/>
  </numFmts>
  <fonts count="9">
    <font><sz val="10"/><color rgb="FF1E293B"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="16"/><color rgb="FF0F172A"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="11"/><color rgb="FF1E3A8A"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><color rgb="FF0F172A"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="9"/><color rgb="FF475569"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="14"/><color rgb="FF1E3A8A"/><name val="맑은 고딕"/></font>
    <font><sz val="9.5"/><color rgb="FF1E3A8A"/><name val="맑은 고딕"/></font>
    <font><b/><sz val="10"/><color rgb="FF047857"/><name val="맑은 고딕"/></font>
  </fonts>
  <fills count="12">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A8A"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF9C3"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF1F5F9"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE0F2FE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2E8F0"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEFF6FF"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFECFDF5"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCBD5E1"/></left>
      <right style="thin"><color rgb="FFCBD5E1"/></right>
      <top style="thin"><color rgb="FFCBD5E1"/></top>
      <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
      <diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FF94A3B8"/></left>
      <right style="thin"><color rgb="FF94A3B8"/></right>
      <top style="medium"><color rgb="FF1E293B"/></top>
      <bottom style="double"><color rgb="FF1E293B"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="28">
    <!-- 0: default -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <!-- 1: Title -->
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <!-- 2: Section Header -->
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <!-- 3: Table Header Input (Charcoal) -->
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <!-- 4: Table Header Calc (Navy Blue) -->
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <!-- 5: Input Center -->
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 6: Input Left -->
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <!-- 7: Input Date -->
    <xf numFmtId="167" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 8: Input Currency/Int -->
    <xf numFmtId="164" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 9: Input Percent 0.00% -->
    <xf numFmtId="166" fontId="0" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 10: Calc Center -->
    <xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 11: Calc Date -->
    <xf numFmtId="167" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 12: Calc Currency -->
    <xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 13: Calc Currency Highlight Bold -->
    <xf numFmtId="164" fontId="4" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 14: Calc Percent 0.0% -->
    <xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 15: Calc Percent 0.00% -->
    <xf numFmtId="166" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 16: Total Row Center -->
    <xf numFmtId="0" fontId="4" fillId="7" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 17: Total Row Currency -->
    <xf numFmtId="164" fontId="4" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 18: Total Row Percent -->
    <xf numFmtId="165" fontId="4" fillId="7" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 19: KPI Card Label -->
    <xf numFmtId="0" fontId="5" fillId="8" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 20: KPI Card Currency Value -->
    <xf numFmtId="164" fontId="6" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 21: KPI Card Percent Value -->
    <xf numFmtId="165" fontId="6" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 22: Guide Box -->
    <xf numFmtId="0" fontId="7" fillId="10" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
    <!-- 23: White Cell Center -->
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <!-- 24: White Cell Left -->
    <xf numFmtId="0" fontId="0" fillId="9" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <!-- 25: White Cell Currency -->
    <xf numFmtId="164" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 26: White Cell Percent -->
    <xf numFmtId="165" fontId="0" fillId="9" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <!-- 27: Status Badge OK -->
    <xf numFmtId="0" fontId="8" fillId="11" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
  <dxfs count="2">
    <!-- dxfId=0: Negative profit or error warning (Red text + Light red background) -->
    <dxf>
      <font><b/><color rgb="FFDC2626"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
    </dxf>
    <!-- dxfId=1: Warning count > 0 (Red text + Light red background) -->
    <dxf>
      <font><b/><color rgb="FFB91C1C"/></font>
      <fill><patternFill patternType="solid"><fgColor rgb="FFFECACA"/></patternFill></fill>
    </dxf>
  </dxfs>
</styleSheet>`;
}

module.exports = {
  CHANNELS,
  CLIENTS,
  CHANNEL_RATES,
  DAILY_SALES,
  MONTHLY_TARGETS,
  COMMON_COSTS,
  computeRowMetrics,
  buildStylesXml
};
