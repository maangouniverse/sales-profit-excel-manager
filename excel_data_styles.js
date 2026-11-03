const { dateToSerial } = require('./xlsx_writer');

// 1. 기본 판매채널 8종 (공판장 추가)
const CHANNELS = [
  '농라',
  '스마트스토어',
  '카톡',
  '고흥몰',
  '남도장터',
  '로컬푸드',
  '공판장',
  '기타'
];

// 공통비용 기본 항목 목록 (공통비용 항목별 그래프 및 집계용)
const COMMON_COST_CATEGORIES = [
  '월 광고선전비',
  '작업장/창고 임대료',
  '인건비/작업비',
  '공과금/통신/프로그램',
  '기타 공통고정비'
];

// 2. 목표설정 기간: 2026년 11월 ~ 2027년 10월 (12개월, 실사용을 위해 예시 수치 0원 초기화)
const TARGET_MONTHS = [
  { idx: 1, ym: '2026-11', label: '26년 11월', year: 2026, month: 11, start: '2026-11-01', end: '2026-11-30', days: 30, target: 0 },
  { idx: 2, ym: '2026-12', label: '26년 12월', year: 2026, month: 12, start: '2026-12-01', end: '2026-12-31', days: 31, target: 0 },
  { idx: 3, ym: '2027-01', label: '27년 1월',  year: 2027, month: 1,  start: '2027-01-01', end: '2027-01-31', days: 31, target: 0 },
  { idx: 4, ym: '2027-02', label: '27년 2월',  year: 2027, month: 2,  start: '2027-02-01', end: '2027-02-28', days: 28, target: 0 },
  { idx: 5, ym: '2027-03', label: '27년 3월',  year: 2027, month: 3,  start: '2027-03-01', end: '2027-03-31', days: 31, target: 0 },
  { idx: 6, ym: '2027-04', label: '27년 4월',  year: 2027, month: 4,  start: '2027-04-01', end: '2027-04-30', days: 30, target: 0 },
  { idx: 7, ym: '2027-05', label: '27년 5월',  year: 2027, month: 5,  start: '2027-05-01', end: '2027-05-31', days: 31, target: 0 },
  { idx: 8, ym: '2027-06', label: '27년 6월',  year: 2027, month: 6,  start: '2027-06-01', end: '2027-06-30', days: 30, target: 0 },
  { idx: 9, ym: '2027-07', label: '27년 7월',  year: 2027, month: 7,  start: '2027-07-01', end: '2027-07-31', days: 31, target: 0 },
  { idx: 10, ym: '2027-08', label: '27년 8월', year: 2027, month: 8,  start: '2027-08-01', end: '2027-08-31', days: 31, target: 0 },
  { idx: 11, ym: '2027-09', label: '27년 9월', year: 2027, month: 9,  start: '2027-09-01', end: '2027-09-30', days: 30, target: 0 },
  { idx: 12, ym: '2027-10', label: '27년 10월', year: 2027, month: 10, start: '2027-10-01', end: '2027-10-31', days: 31, target: 0 }
];

// 3. 판매채널·기간별 비용 설정 (기본 8개 채널 틀만 등록, 예시 단가는 모두 0으로 비워둠)
const CHANNEL_RATES = [
  { channel: '농라', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '스마트스토어', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '카톡', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '고흥몰', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '남도장터', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '로컬푸드', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '공판장', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' },
  { channel: '기타', start: '2026-11-01', end: '', feeRate: 0, shipUnder10: 0, shipOver10: 0, boxCost: 0, vinylCost: 0, leafletCost: 0, etcCost: 0, note: '' }
];

// 4. 일일 매출 입력 데이터 (예시 데이터 모두 삭제 완료)
const DAILY_SALES = [];

// 5. 월별 공통비용 데이터 (예시 데이터 모두 삭제 완료)
const COMMON_COSTS = [];

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
  COMMON_COST_CATEGORIES,
  TARGET_MONTHS,
  CHANNEL_RATES,
  DAILY_SALES,
  COMMON_COSTS,
  computeRowMetrics,
  buildStylesXml
};
