const fs = require('fs');
const path = require('path');
const { createZipBuffer, escXml, colToLetter, dateToSerial, SheetBuilder } = require('./xlsx_writer');
const { buildBarChartXml, buildDoughnutChartXml, buildDrawingXml } = require('./chart_builder');
const {
  CHANNELS,
  TARGET_MONTHS,
  CHANNEL_RATES,
  DAILY_SALES,
  COMMON_COSTS,
  computeRowMetrics,
  buildStylesXml
} = require('./excel_data_styles');

const enrichedSales = DAILY_SALES.map(r => ({
  ...r,
  ...computeRowMetrics(r, CHANNEL_RATES)
}));

// 기본 조회 기간: 2026년 11월 (2026-11-01 ~ 2026-11-30)
const defaultYm = '2026-11';
const defaultStartStr = '2026-11-01';
const defaultEndStr = '2026-11-30';
const defaultStartSerial = dateToSerial(defaultStartStr);
const defaultEndSerial = dateToSerial(defaultEndStr);

const periodSales = enrichedSales.filter(r => {
  const s = dateToSerial(r.date);
  return s >= defaultStartSerial && s <= defaultEndSerial && r.status === '정상';
});

const totalAnnualTarget = TARGET_MONTHS.reduce((a, b) => a + b.target, 0);
const periodTarget = TARGET_MONTHS.find(m => m.ym === defaultYm).target;
const periodActualSales = periodSales.reduce((a, b) => a + b.sales, 0);
const periodOrdersUnder10 = periodSales.reduce((a, b) => a + b.ordersUnder10, 0);
const periodOrdersOver10 = periodSales.reduce((a, b) => a + b.ordersOver10, 0);
const periodOrders = periodOrdersUnder10 + periodOrdersOver10;
const periodFee = periodSales.reduce((a, b) => a + b.feeAmt, 0);
const periodOrderCost = periodSales.reduce((a, b) => a + b.orderCost, 0);
const periodExtraCost = periodSales.reduce((a, b) => a + b.extraCost, 0);
const periodTotalCost = periodSales.reduce((a, b) => a + b.totalCost, 0);
const periodNetProfit = periodActualSales - periodTotalCost;
const periodMargin = periodActualSales > 0 ? periodNetProfit / periodActualSales : 0;
const periodCommonCost = COMMON_COSTS.filter(c => {
  const s = dateToSerial(c.date);
  return s >= defaultStartSerial && s <= defaultEndSerial;
}).reduce((a, b) => a + b.amount, 0);
const periodFinalProfit = periodNetProfit - periodCommonCost;
const periodFinalMargin = periodActualSales > 0 ? periodFinalProfit / periodActualSales : 0;

// ============================================================================
// 시트 1: 1_매출손익_대시보드
// ============================================================================
function buildSheet1() {
  const sh = new SheetBuilder('1_매출손익_대시보드', {
    freezeRow: 4,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 15, width: 15 }
    ]
  });

  sh.setRowHeight(2, 30);
  sh.setCell(2, 2, { v: '매출·지출관리 스프레드시트(순이익/실적 자동계산) | 경영 실적 종합 대시보드', s: 1 });
  sh.merge('B2:I2');
  sh.setCell(2, 10, { v: '💡 조회 연월(2026-11 ~ 2027-10) 또는 시작·종료일을 선택하면 요약과 그래프가 자동 갱신됩니다.', s: 22 });
  sh.merge('J2:O2');

  // 3~4행: 기간 선택 및 데이터 점검 배너
  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 26);

  const ctrlHeaders = [
    [2, '조회 연월 선택', 3],
    [3, '12개월 총 목표매출', 4],
    [4, '직접 시작일(선택)', 3],
    [5, '직접 종료일(선택)', 3],
    [6, '적용 시작일(자동)', 4],
    [7, '적용 종료일(자동)', 4],
    [8, '비용설정 누락건수', 4],
    [9, '기간설정 중복건수', 4],
    [10, '데이터/설정 종합 점검 상태', 4]
  ];
  for (const [col, text, st] of ctrlHeaders) {
    sh.setCell(3, col, { v: text, s: st });
  }
  sh.merge('J3:O3');
  for (let c = 11; c <= 15; c++) sh.setCell(3, c, { v: '', s: 4 });

  sh.setCell(4, 2, { v: defaultYm, s: 5 });
  sh.setCell(4, 3, { f: '\'5_목표매출_및_공통비용\'!E19', v: totalAnnualTarget, s: 13 });
  sh.setCell(4, 4, { v: '', s: 7 }); // 직접 시작일
  sh.setCell(4, 5, { v: '', s: 7 }); // 직접 종료일
  sh.setCell(4, 6, {
    f: 'IF(D4<>"",D4,IF(B4="전체",DATE(2026,11,1),DATE(VALUE(LEFT(B4,4)),VALUE(RIGHT(B4,2)),1)))',
    v: defaultStartSerial,
    s: 11
  });
  sh.setCell(4, 7, {
    f: 'IF(E4<>"",E4,IF(B4="전체",DATE(2027,10,31),EOMONTH(DATE(VALUE(LEFT(B4,4)),VALUE(RIGHT(B4,2)),1),0)))',
    v: defaultEndSerial,
    s: 11
  });
  sh.setCell(4, 8, {
    f: 'COUNTIF(\'2_일일매출_입력\'!$J$6:$J$305,"비용 설정 필요")',
    v: 0,
    s: 12
  });
  sh.setCell(4, 9, {
    f: 'COUNTIF(\'3_채널별_비용설정\'!$N$6:$N$55,"*중복*")+COUNTIF(\'3_채널별_비용설정\'!$N$6:$N$55,"*오류*")',
    v: 0,
    s: 12
  });
  sh.setCell(4, 10, {
    f: 'IF(H4+I4=0,"✅ 모든 채널 비용 설정 및 기간 적용 정상","⚠️ 비용 설정 누락(" & H4 & "건) 또는 기간 중복(" & I4 & "건) 점검 필요!")',
    v: '✅ 모든 채널 비용 설정 및 기간 적용 정상',
    s: 27
  });
  sh.merge('J4:O4');
  for (let c = 11; c <= 15; c++) sh.setCell(4, c, { v: '', s: 27 });

  // 6~7행: 핵심 KPI 카드 1단
  sh.setRowHeight(6, 20);
  sh.setRowHeight(7, 32);
  const kpiRow1 = [
    { c1: 2, c2: 3, label: '선택기간 목표매출 (원)', f: '\'5_목표매출_및_공통비용\'!G19', v: periodTarget, s: 20 },
    { c1: 4, c2: 5, label: '총 매출액 (실제매출)', f: 'SUMIFS(\'2_일일매출_입력\'!$G$6:$G$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodActualSales, s: 20 },
    { c1: 6, c2: 7, label: '목표 달성률 (%)', f: 'IFERROR(D7/B7,0)', v: periodTarget > 0 ? periodActualSales / periodTarget : 0, s: 21 },
    { c1: 8, c2: 9, label: '총 주문건수 (미만+이상)', f: 'SUMIFS(\'2_일일매출_입력\'!$F$6:$F$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodOrders, s: 20 },
    { c1: 10, c2: 11, label: '총비용 (공통비 제외)', f: 'SUMIFS(\'2_일일매출_입력\'!$P$6:$P$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodTotalCost, s: 20 },
    { c1: 12, c2: 13, label: '총 순이익 (공통비 전)', f: 'D7-J7', v: periodNetProfit, s: 20 },
    { c1: 14, c2: 15, label: '순이익률 (공통비 전)', f: 'IFERROR(L7/D7,0)', v: periodMargin, s: 21 }
  ];
  for (const k of kpiRow1) {
    sh.setCell(6, k.c1, { v: k.label, s: 19 });
    sh.setCell(6, k.c2, { v: '', s: 19 });
    sh.merge(`${colToLetter(k.c1)}6:${colToLetter(k.c2)}6`);
    sh.setCell(7, k.c1, { f: k.f, v: k.v, s: k.s });
    sh.setCell(7, k.c2, { v: '', s: k.s });
    sh.merge(`${colToLetter(k.c1)}7:${colToLetter(k.c2)}7`);
  }

  // 8~9행: 핵심 KPI 카드 2단 (10kg 미만/이상 주문건수 구분 + 세부 비용 + 공통비용 후 최종순이익)
  sh.setRowHeight(8, 20);
  sh.setRowHeight(9, 28);
  const kpiRow2 = [
    { c1: 2, c2: 3, label: '10kg 미만 주문건수 (건)', f: 'SUMIFS(\'2_일일매출_입력\'!$D$6:$D$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodOrdersUnder10, s: 12 },
    { c1: 4, c2: 5, label: '10kg 이상 주문건수 (건)', f: 'SUMIFS(\'2_일일매출_입력\'!$E$6:$E$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodOrdersOver10, s: 12 },
    { c1: 6, c2: 7, label: '매출액 기준 수수료 (원)', f: 'SUMIFS(\'2_일일매출_입력\'!$L$6:$L$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodFee, s: 12 },
    { c1: 8, c2: 9, label: '택배·포장 건수비용 합계', f: 'SUMIFS(\'2_일일매출_입력\'!$O$6:$O$305,\'2_일일매출_입력\'!$B$6:$B$305,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$305,"<="&$G$4)', v: periodOrderCost, s: 12 },
    { c1: 10, c2: 11, label: '기간 공통비용 (광고·임대)', f: '\'5_목표매출_및_공통비용\'!M26', v: periodCommonCost, s: 12 },
    { c1: 12, c2: 13, label: '최종 순이익 (공통비 후)', f: 'L7-J9', v: periodFinalProfit, s: 13 },
    { c1: 14, c2: 15, label: '최종 순이익률 (공통비 후)', f: 'IFERROR(L9/D7,0)', v: periodFinalMargin, s: 14 }
  ];
  for (const k of kpiRow2) {
    sh.setCell(8, k.c1, { v: k.label, s: 19 });
    sh.setCell(8, k.c2, { v: '', s: 19 });
    sh.merge(`${colToLetter(k.c1)}8:${colToLetter(k.c2)}8`);
    sh.setCell(9, k.c1, { f: k.f, v: k.v, s: k.s });
    sh.setCell(9, k.c2, { v: '', s: k.s });
    sh.merge(`${colToLetter(k.c1)}9:${colToLetter(k.c2)}9`);
  }

  sh.setCell(10, 2, { v: '전체 매출성과 및 채널별 비중 그래프 (실시간 자동 갱신)', s: 2 });

  // 29~43행: 월별 손익현황 (26년 11월 ~ 27년 10월) & 채널별 매출 비중 표
  sh.setRowHeight(29, 24);
  sh.setCell(29, 2, { v: '월별 손익현황 (26년 11월 ~ 27년 10월)', s: 2 });
  sh.setCell(29, 10, { v: '채널별 매출 비중 및 순이익 현황 (선택 기간)', s: 2 });

  sh.setRowHeight(30, 24);
  const mHeaders = [
    [2, '연월'],
    [3, '목표매출'],
    [4, '총 매출'],
    [5, '총 지출(비용)'],
    [6, '순이익'],
    [7, '이익률'],
    [8, '총 주문건수']
  ];
  for (const [c, t] of mHeaders) sh.setCell(30, c, { v: t, s: 4 });

  const monthlyChartCache = {
    cats: [],
    target: [],
    sales: [],
    cost: [],
    profit: []
  };

  TARGET_MONTHS.forEach((tm, idx) => {
    const r = 31 + idx;
    const mStartSer = dateToSerial(tm.start);
    const mEndSer = dateToSerial(tm.end);
    const mRows = enrichedSales.filter(x => {
      const s = dateToSerial(x.date);
      return s >= mStartSer && s <= mEndSer && x.status === '정상';
    });
    const mSales = mRows.reduce((a, b) => a + b.sales, 0);
    const mCost = mRows.reduce((a, b) => a + b.totalCost, 0);
    const mProfit = mSales - mCost;
    const mMargin = mSales > 0 ? mProfit / mSales : 0;
    const mOrders = mRows.reduce((a, b) => a + b.totalOrders, 0);

    monthlyChartCache.cats.push(tm.label);
    monthlyChartCache.target.push(tm.target);
    monthlyChartCache.sales.push(mSales);
    monthlyChartCache.cost.push(mCost);
    monthlyChartCache.profit.push(mProfit);

    sh.setCell(r, 2, { v: tm.label, s: 23 });
    sh.setCell(r, 3, { f: `'5_목표매출_및_공통비용'!E${7 + idx}`, v: tm.target, s: 25 });
    sh.setCell(r, 4, {
      f: `SUMIFS('2_일일매출_입력'!$G$6:$G$305,'2_일일매출_입력'!$B$6:$B$305,">="&'5_목표매출_및_공통비용'!C${7 + idx},'2_일일매출_입력'!$B$6:$B$305,"<="&'5_목표매출_및_공통비용'!D${7 + idx})`,
      v: mSales,
      s: 25
    });
    sh.setCell(r, 5, {
      f: `SUMIFS('2_일일매출_입력'!$P$6:$P$305,'2_일일매출_입력'!$B$6:$B$305,">="&'5_목표매출_및_공통비용'!C${7 + idx},'2_일일매출_입력'!$B$6:$B$305,"<="&'5_목표매출_및_공통비용'!D${7 + idx})`,
      v: mCost,
      s: 25
    });
    sh.setCell(r, 6, { f: `D${r}-E${r}`, v: mProfit, s: 12 });
    sh.setCell(r, 7, { f: `IFERROR(F${r}/D${r},0)`, v: mMargin, s: 26 });
    sh.setCell(r, 8, {
      f: `SUMIFS('2_일일매출_입력'!$F$6:$F$305,'2_일일매출_입력'!$B$6:$B$305,">="&'5_목표매출_및_공통비용'!C${7 + idx},'2_일일매출_입력'!$B$6:$B$305,"<="&'5_목표매출_및_공통비용'!D${7 + idx})`,
      v: mOrders,
      s: 25
    });
  });

  // 43행: 12개월 합계
  const totTarget = monthlyChartCache.target.reduce((a, b) => a + b, 0);
  const totSales = monthlyChartCache.sales.reduce((a, b) => a + b, 0);
  const totCost = monthlyChartCache.cost.reduce((a, b) => a + b, 0);
  const totProfit = totSales - totCost;
  sh.setCell(43, 2, { v: '12개월 합계', s: 16 });
  sh.setCell(43, 3, { f: 'SUM(C31:C42)', v: totTarget, s: 17 });
  sh.setCell(43, 4, { f: 'SUM(D31:D42)', v: totSales, s: 17 });
  sh.setCell(43, 5, { f: 'SUM(E31:E42)', v: totCost, s: 17 });
  sh.setCell(43, 6, { f: 'D43-E43', v: totProfit, s: 17 });
  sh.setCell(43, 7, { f: 'IFERROR(F43/D43,0)', v: totSales > 0 ? totProfit / totSales : 0, s: 18 });
  sh.setCell(43, 8, { f: 'SUM(H31:H42)', v: enrichedSales.reduce((a, b) => a + b.totalOrders, 0), s: 17 });

  // 우측 표 헤더 (J30:O30): 채널별 매출 비중
  const chHeaders = [
    [10, '채널명'],
    [11, '주문건수'],
    [12, '매출액'],
    [13, '비중'],
    [14, '순이익'],
    [15, '이익률']
  ];
  for (const [c, t] of chHeaders) sh.setCell(30, c, { v: t, s: 4 });

  const channelChartCache = {
    cats: [],
    sales: [],
    profits: []
  };

  CHANNELS.forEach((chName, idx) => {
    const r = 31 + idx;
    const chRows = periodSales.filter(x => x.channel === chName);
    const cOrders = chRows.reduce((a, b) => a + b.totalOrders, 0);
    const cSales = chRows.reduce((a, b) => a + b.sales, 0);
    const cProfit = chRows.reduce((a, b) => a + b.netProfit, 0);
    const cShare = periodActualSales > 0 ? cSales / periodActualSales : 0;
    const cMargin = cSales > 0 ? cProfit / cSales : 0;

    channelChartCache.cats.push(chName);
    channelChartCache.sales.push(cSales);
    channelChartCache.profits.push(cProfit);

    sh.setCell(r, 10, { f: `'6_사용설명서_및_기준관리'!B${6 + idx}`, v: chName, s: 23 });
    sh.setCell(r, 11, {
      f: `SUMIFS('2_일일매출_입력'!$F$6:$F$305,'2_일일매출_입력'!$C$6:$C$305,J${r},'2_일일매출_입력'!$B$6:$B$305,">="&$F$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$G$4)`,
      v: cOrders,
      s: 25
    });
    sh.setCell(r, 12, {
      f: `SUMIFS('2_일일매출_입력'!$G$6:$G$305,'2_일일매출_입력'!$C$6:$C$305,J${r},'2_일일매출_입력'!$B$6:$B$305,">="&$F$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$G$4)`,
      v: cSales,
      s: 25
    });
    sh.setCell(r, 13, { f: `IFERROR(L${r}/$L$38,0)`, v: cShare, s: 14 });
    sh.setCell(r, 14, {
      f: `SUMIFS('2_일일매출_입력'!$Q$6:$Q$305,'2_일일매출_입력'!$C$6:$C$305,J${r},'2_일일매출_입력'!$B$6:$B$305,">="&$F$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$G$4)`,
      v: cProfit,
      s: 12
    });
    sh.setCell(r, 15, { f: `IFERROR(N${r}/L${r},0)`, v: cMargin, s: 26 });
  });

  sh.setCell(38, 10, { v: '합계', s: 16 });
  sh.setCell(38, 11, { f: 'SUM(K31:K37)', v: periodOrders, s: 17 });
  sh.setCell(38, 12, { f: 'SUM(L31:L37)', v: periodActualSales, s: 17 });
  sh.setCell(38, 13, { f: 'IFERROR(SUM(M31:M37),0)', v: 1, s: 18 });
  sh.setCell(38, 14, { f: 'SUM(N31:N37)', v: periodNetProfit, s: 17 });
  sh.setCell(38, 15, { f: 'IFERROR(N38/L38,0)', v: periodMargin, s: 18 });

  // 목표 vs 실제 비교 차트용 요약 테이블 (J40:N41)
  sh.setCell(40, 10, { v: '구분', s: 4 });
  sh.setCell(40, 11, { v: '목표매출', s: 4 });
  sh.setCell(40, 12, { v: '총 매출액', s: 4 });
  sh.setCell(40, 13, { v: '총 순이익', s: 4 });
  sh.setCell(40, 14, { v: '최종순이익(공통비후)', s: 4 });

  sh.setCell(41, 10, { v: '전체 매출성과', s: 23 });
  sh.setCell(41, 11, { f: 'B7', v: periodTarget, s: 25 });
  sh.setCell(41, 12, { f: 'D7', v: periodActualSales, s: 25 });
  sh.setCell(41, 13, { f: 'L7', v: periodNetProfit, s: 25 });
  sh.setCell(41, 14, { f: 'L9', v: periodFinalProfit, s: 25 });

  sh.addValidation({
    type: 'list',
    sqref: 'B4',
    formula1: '"전체,2026-11,2026-12,2027-01,2027-02,2027-03,2027-04,2027-05,2027-06,2027-07,2027-08,2027-09,2027-10"',
    promptTitle: '조회 연월 선택 (2026.11 ~ 2027.10)',
    prompt: '조회할 연월(2026-11 ~ 2027-10) 또는 [전체]를 선택하세요.'
  });

  sh.addCondFormat({
    sqref: 'L7:L9 F31:F43 N31:N38',
    rules: [{ type: 'cellIs', operator: 'lessThan', formula: '0', dxfId: 0 }]
  });
  sh.addCondFormat({
    sqref: 'H4:I4',
    rules: [{ type: 'cellIs', operator: 'greaterThan', formula: '0', dxfId: 1 }]
  });

  sh.drawingRelId = 'rId1';

  return { sh, monthlyChartCache, channelChartCache };
}

// ============================================================================
// 시트 2: 2_일일매출_입력 (거래처 삭제, 상품총원가 삭제, 10kg 미만/이상 주문건수 구분)
// ============================================================================
function buildSheet2() {
  const sh = new SheetBuilder('2_일일매출_입력', {
    freezeRow: 5,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 16 }, // B: 매출일자
      { min: 3, max: 3, width: 15 }, // C: 판매채널
      { min: 4, max: 6, width: 14 }, // D: 10kg미만 건수, E: 10kg이상 건수, F: 총 주문건수
      { min: 7, max: 8, width: 16 }, // G: 일일 매출액, H: 추가 비용
      { min: 9, max: 9, width: 36 }, // I: 비고
      { min: 10, max: 10, width: 15 }, // J: 비용설정 상태
      { min: 11, max: 11, width: 13 }, // K: 적용 수수료율
      { min: 12, max: 17, width: 15 }, // L: 매출수수료, M: 10kg미만 단가, N: 10kg이상 단가, O: 건수비용합계, P: 총비용, Q: 순이익
      { min: 18, max: 18, width: 12 }  // R: 순이익률
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '매출·지출관리 스프레드시트(순이익/실적 자동계산) | 일일 매출·주문내역 입력', s: 1 });
  sh.merge('B2:I2');

  sh.setRowHeight(3, 36);
  sh.setCell(3, 2, {
    v: '💡 [입력 안내] ① 노란색 열(매출일자, 판매채널, 10kg 미만 주문건수, 10kg 이상 주문건수, 일일 매출액, 추가비용, 비고)만 입력하세요.\n② 10kg 미만과 10kg 이상 주문건수에 따라 각각 다른 택배비와 포장비(박스+포장비닐+리플렛)가 자동 계산됩니다.\n③ 기본 택배비·박스·비닐·리플렛 비용은 자동 계산되므로 [추가 비용] 칸에 중복 입력하지 마세요.',
    s: 22
  });
  sh.merge('B3:R3');
  for (let c = 3; c <= 18; c++) sh.setCell(3, c, { v: '', s: 22 });

  // 4행: 전체 누적 합계 요약 바 (참고 이미지 2단 그룹 헤더 느낌 반영)
  sh.setRowHeight(4, 24);
  const totUnder10All = enrichedSales.reduce((a, b) => a + b.ordersUnder10, 0);
  const totOver10All = enrichedSales.reduce((a, b) => a + b.ordersOver10, 0);
  const totOrdersAll = totUnder10All + totOver10All;
  const totSalesAll = enrichedSales.reduce((a, b) => a + b.sales, 0);
  const totExtraAll = enrichedSales.reduce((a, b) => a + b.extraCost, 0);
  const totFeeAll = enrichedSales.reduce((a, b) => a + b.feeAmt, 0);
  const totOrderCostAll = enrichedSales.reduce((a, b) => a + b.orderCost, 0);
  const totCostAll = enrichedSales.reduce((a, b) => a + b.totalCost, 0);
  const totProfitAll = totSalesAll - totCostAll;

  sh.setCell(4, 2, { v: '누적 합계', s: 16 });
  sh.setCell(4, 3, { f: 'COUNTA(B6:B305)&"건"', v: `${enrichedSales.length}건`, s: 16 });
  sh.setCell(4, 4, { f: 'SUM(D6:D305)', v: totUnder10All, s: 17 });
  sh.setCell(4, 5, { f: 'SUM(E6:E305)', v: totOver10All, s: 17 });
  sh.setCell(4, 6, { f: 'SUM(F6:F305)', v: totOrdersAll, s: 17 });
  sh.setCell(4, 7, { f: 'SUM(G6:G305)', v: totSalesAll, s: 17 });
  sh.setCell(4, 8, { f: 'SUM(H6:H305)', v: totExtraAll, s: 17 });
  sh.setCell(4, 9, { v: '← 매출·주문 입력 | 비용·손익 자동계산 →', s: 16 });
  sh.setCell(4, 10, { f: 'COUNTIF(J6:J305,"비용 설정 필요")&"건 누락"', v: '0건 누락', s: 16 });
  sh.setCell(4, 11, { v: '-', s: 16 });
  sh.setCell(4, 12, { f: 'SUMIF(L6:L305,">=0")', v: totFeeAll, s: 17 });
  sh.setCell(4, 13, { v: '-', s: 16 });
  sh.setCell(4, 14, { v: '-', s: 16 });
  sh.setCell(4, 15, { f: 'SUMIF(O6:O305,">=0")', v: totOrderCostAll, s: 17 });
  sh.setCell(4, 16, { f: 'SUMIF(P6:P305,">=0")', v: totCostAll, s: 17 });
  sh.setCell(4, 17, { f: 'G4-P4', v: totProfitAll, s: 17 });
  sh.setCell(4, 18, { f: 'IFERROR(Q4/G4,0)', v: totSalesAll > 0 ? totProfitAll / totSalesAll : 0, s: 18 });

  // 5행: 테이블 헤더
  sh.setRowHeight(5, 28);
  const colsMeta = [
    [2, '매출일자\n(달력/날짜입력)', 3],
    [3, '판매채널\n(드롭다운)', 3],
    [4, '10kg 미만\n주문건수(건)', 3],
    [5, '10kg 이상\n주문건수(건)', 3],
    [6, '총 주문건수\n(자동합산)', 4],
    [7, '일일 매출액\n(할인반영 실매출)', 3],
    [8, '추가 비용\n(특이발생분)', 3],
    [9, '비고\n(메모)', 3],
    [10, '비용설정 상태\n(자동검증)', 4],
    [11, '적용 수수료율\n(자동조회)', 4],
    [12, '매출액 수수료\n(자동계산)', 4],
    [13, '10kg 미만\n1건당 비용', 4],
    [14, '10kg 이상\n1건당 비용', 4],
    [15, '주문건수 비용\n합계(택배+포장)', 4],
    [16, '총 지출(총비용)\n(수수료+건수비+추가)', 4],
    [17, '순이익\n(매출-총비용)', 4],
    [18, '순이익률\n(순이익÷매출)', 4]
  ];
  for (const [c, t, st] of colsMeta) sh.setCell(5, c, { v: t, s: st });

  const maxPreformattedRows = 300;
  for (let i = 0; i < maxPreformattedRows; i++) {
    const r = 6 + i;
    const item = enrichedSales[i];

    if (item) {
      sh.setCell(r, 2, { v: dateToSerial(item.date), s: 7 });
      sh.setCell(r, 3, { v: item.channel, s: 5 });
      sh.setCell(r, 4, { v: item.ordersUnder10, s: 8 });
      sh.setCell(r, 5, { v: item.ordersOver10, s: 8 });
      sh.setCell(r, 6, { f: `IF(AND(B${r}="",C${r}=""),"",SUM(D${r}:E${r}))`, v: item.totalOrders, s: 12 });
      sh.setCell(r, 7, { v: item.sales, s: 8 });
      sh.setCell(r, 8, { v: item.extraCost, s: 8 });
      sh.setCell(r, 9, { v: item.note, s: 6 });
    } else {
      sh.setCell(r, 2, { v: '', s: 7 });
      sh.setCell(r, 3, { v: '', s: 5 });
      sh.setCell(r, 4, { v: '', s: 8 });
      sh.setCell(r, 5, { v: '', s: 8 });
      sh.setCell(r, 6, { f: `IF(AND(B${r}="",C${r}=""),"",SUM(D${r}:E${r}))`, v: '', s: 12 });
      sh.setCell(r, 7, { v: '', s: 8 });
      sh.setCell(r, 8, { v: '', s: 8 });
      sh.setCell(r, 9, { v: '', s: 6 });
    }

    // 자동 계산 수식 (J~R열)
    // 시트3 참조: C=판매채널, D=시작일, O=종료기준일, F=수수료율, L=10kg미만건당합계, M=10kg이상건당합계
    const fStatus = `IF(AND(B${r}="",C${r}=""),"",IF(OR(B${r}="",C${r}=""),"비용 설정 필요",IF(COUNTIFS('3_채널별_비용설정'!$C$6:$C$55,C${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$O$6:$O$55,">="&B${r})=0,"비용 설정 필요",IF(COUNTIFS('3_채널별_비용설정'!$C$6:$C$55,C${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$O$6:$O$55,">="&B${r})>1,"기간 중복 확인","정상"))))`;
    const fFeeRate = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",SUMIFS('3_채널별_비용설정'!$F$6:$F$55,'3_채널별_비용설정'!$C$6:$C$55,C${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$O$6:$O$55,">="&B${r})))`;
    const fFeeAmt = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",ROUND(G${r}*K${r},0)))`;
    const fUnitUnder10 = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",SUMIFS('3_채널별_비용설정'!$L$6:$L$55,'3_채널별_비용설정'!$C$6:$C$55,C${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$O$6:$O$55,">="&B${r})))`;
    const fUnitOver10 = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",SUMIFS('3_채널별_비용설정'!$M$6:$M$55,'3_채널별_비용설정'!$C$6:$C$55,C${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$O$6:$O$55,">="&B${r})))`;
    const fOrderCost = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",(D${r}*M${r})+(E${r}*N${r})))`;
    const fTotalCost = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",L${r}+O${r}+H${r}))`;
    const fNetProfit = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",G${r}-P${r}))`;
    const fMargin = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",IFERROR(Q${r}/G${r},0)))`;

    sh.setCell(r, 10, { f: fStatus, v: item ? item.status : '', s: 10 });
    sh.setCell(r, 11, { f: fFeeRate, v: item ? item.feeRate : '', s: 15 });
    sh.setCell(r, 12, { f: fFeeAmt, v: item ? item.feeAmt : '', s: 12 });
    sh.setCell(r, 13, { f: fUnitUnder10, v: item ? item.unitUnder10 : '', s: 12 });
    sh.setCell(r, 14, { f: fUnitOver10, v: item ? item.unitOver10 : '', s: 12 });
    sh.setCell(r, 15, { f: fOrderCost, v: item ? item.orderCost : '', s: 12 });
    sh.setCell(r, 16, { f: fTotalCost, v: item ? item.totalCost : '', s: 12 });
    sh.setCell(r, 17, { f: fNetProfit, v: item ? item.netProfit : '', s: 13 });
    sh.setCell(r, 18, { f: fMargin, v: item ? item.margin : '', s: 14 });
  }

  sh.addValidation({
    type: 'list',
    sqref: 'C6:C305',
    formula1: "'6_사용설명서_및_기준관리'!$B$6:$B$25",
    promptTitle: '판매채널 선택',
    prompt: '목록에서 판매채널을 선택하세요.'
  });
  sh.addValidation({
    type: 'whole',
    operator: 'greaterThanOrEqual',
    sqref: 'D6:E305',
    formula1: '0',
    promptTitle: '주문건수 입력 (0 이상 정수)',
    prompt: '10kg 미만 / 10kg 이상 주문건수를 각각 0 이상의 정수로 입력하세요.',
    errorTitle: '입력 오류 (주문건수)',
    error: '주문건수는 0 이상의 정수만 입력할 수 있습니다.'
  });

  sh.addCondFormat({
    sqref: 'Q6:Q305',
    rules: [{ type: 'cellIs', operator: 'lessThan', formula: '0', dxfId: 0 }]
  });
  sh.addCondFormat({
    sqref: 'J6:R305',
    rules: [{ type: 'containsText', text: '비용 설정 필요', topLeftCell: 'J6', dxfId: 1 }]
  });

  return sh;
}

// ============================================================================
// 시트 3: 3_채널별_비용설정 (10kg 미만/이상 배송비 구분 + 박스, 포장비닐, 리플렛 비용 추가)
// ============================================================================
function buildSheet3() {
  const sh = new SheetBuilder('3_채널별_비용설정', {
    freezeRow: 5,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 6 },   // B: No
      { min: 3, max: 3, width: 15 },  // C: 판매채널
      { min: 4, max: 5, width: 16 },  // D: 적용 시작일, E: 적용 종료일
      { min: 6, max: 6, width: 14 },  // F: 매출액 수수료율
      { min: 7, max: 8, width: 15 },  // G: 10kg 미만 배송비, H: 10kg 이상 배송비
      { min: 9, max: 12, width: 13 }, // I: 박스, J: 포장비닐, K: 리플렛, L: 기타비용
      { min: 13, max: 14, width: 16 }, // M: 10kg 미만 건당합계, N: 10kg 이상 건당합계 -> wait let's check exact columns!
      { min: 15, max: 15, width: 16 }, // O: 기간 중복 검증
      { min: 16, max: 16, width: 15 }, // P: 종료 기준일
      { min: 17, max: 17, width: 40 }  // Q: 비고
    ]
  });

  // 잠깐! 시트2에서 참조한 열 문자와 100% 일치시키겠습니다:
  // B(2): No
  // C(3): 판매채널
  // D(4): 적용 시작일
  // E(5): 적용 종료일
  // F(6): 매출액 기준 수수료율 (%)
  // G(7): 10kg 미만 배송비 (원)
  // H(8): 10kg 이상 배송비 (원)
  // I(9): 박스 비용 (원)
  // J(10): 포장비닐 비용 (원)
  // K(11): 리플렛 비용 (원)
  // L(12): 10kg 미만 1건당 비용 합계 (자동: =G+I+J+K)
  // M(13): 10kg 이상 1건당 비용 합계 (자동: =H+I+J+K)
  // N(14): 적용 기간 중복/오류 검증
  // O(15): 종료 기준일 (수식 자동처리)
  // P(16): 비고
  // 이 열 배치는 시트1(N열 참조) 및 시트2(C, D, O, F, L, M열 참조)와 100% 정확히 일치합니다!

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '매출·지출관리 스프레드시트(순이익/실적 자동계산) | 판매채널·기간별 비용 설정', s: 1 });
  sh.merge('B2:P2');

  sh.setRowHeight(3, 36);
  sh.setCell(3, 2, {
    v: '💡 [비용 설정 안내] ① 주문 무게에 따라 [10kg 미만 배송비]와 [10kg 이상 배송비]를 구분하여 설정하고, 주문 1건당 [박스 비용], [포장비닐 비용], [리플렛 비용]을 입력하세요.\n② 비용이 변경되면 기존 설정의 [적용 종료일]을 입력하고 아래 새 행에 [새로운 적용 시작일]로 추가하세요. (웹 화면에서는 달력 아이콘 클릭으로 날짜를 바로 선택할 수 있습니다.)',
    s: 22
  });
  sh.merge('B3:P3');
  for (let c = 3; c <= 16; c++) sh.setCell(3, c, { v: '', s: 22 });

  sh.setRowHeight(5, 28);
  const headers = [
    [2, 'No', 4],
    [3, '판매채널\n(드롭다운)', 3],
    [4, '적용 시작일\n(달력/날짜선택)', 3],
    [5, '적용 종료일\n(비워두면 계속)', 3],
    [6, '매출액 기준\n수수료율 (%)', 3],
    [7, '10kg 미만\n배송비 (원)', 3],
    [8, '10kg 이상\n배송비 (원)', 3],
    [9, '박스 비용\n(1건당, 원)', 3],
    [10, '포장비닐 비용\n(1건당, 원)', 3],
    [11, '리플렛 비용\n(1건당, 원)', 3],
    [12, '10kg 미만\n1건당 합계(자동)', 4],
    [13, '10kg 이상\n1건당 합계(자동)', 4],
    [14, '적용 기간\n중복/오류 검증', 4],
    [15, '종료 기준일\n(수식 자동처리)', 4],
    [16, '비고 (단가 변경 이력 및 메모)', 3]
  ];
  for (const [c, t, st] of headers) sh.setCell(5, c, { v: t, s: st });

  const maxRateRows = 50;
  for (let i = 0; i < maxRateRows; i++) {
    const r = 6 + i;
    const item = CHANNEL_RATES[i];

    sh.setCell(r, 2, { f: `IF(C${r}="","",ROW()-5)`, v: item ? i + 1 : '', s: 10 });
    if (item) {
      const startSer = dateToSerial(item.start);
      const endSer = item.end ? dateToSerial(item.end) : '';
      const effEndSer = item.end ? endSer : dateToSerial('2099-12-31');
      const packSum = item.boxCost + item.vinylCost + item.leafletCost;
      const sumUnder10 = item.shipUnder10 + packSum;
      const sumOver10 = item.shipOver10 + packSum;

      sh.setCell(r, 3, { v: item.channel, s: 5 });
      sh.setCell(r, 4, { v: startSer, s: 7 });
      sh.setCell(r, 5, { v: endSer, s: 7 });
      sh.setCell(r, 6, { v: item.feeRate, s: 9 });
      sh.setCell(r, 7, { v: item.shipUnder10, s: 8 });
      sh.setCell(r, 8, { v: item.shipOver10, s: 8 });
      sh.setCell(r, 9, { v: item.boxCost, s: 8 });
      sh.setCell(r, 10, { v: item.vinylCost, s: 8 });
      sh.setCell(r, 11, { v: item.leafletCost, s: 8 });
      sh.setCell(r, 12, { f: `IF(C${r}="","",G${r}+SUM(I${r}:K${r}))`, v: sumUnder10, s: 13 });
      sh.setCell(r, 13, { f: `IF(C${r}="","",H${r}+SUM(I${r}:K${r}))`, v: sumOver10, s: 13 });
      sh.setCell(r, 14, {
        f: `IF(C${r}="","",IF(D${r}>O${r},"⚠️ 시작일>종료일 오류",IF(COUNTIFS($C$6:$C$55,C${r},$D$6:$D$55,"<="&O${r},$O$6:$O$55,">="&D${r})>1,"⚠️ 기간 중복","정상")))`,
        v: '정상',
        s: 10
      });
      sh.setCell(r, 15, { f: `IF(C${r}="","",IF(E${r}="",DATE(2099,12,31),E${r}))`, v: effEndSer, s: 11 });
      sh.setCell(r, 16, { v: item.note, s: 6 });
    } else {
      sh.setCell(r, 3, { v: '', s: 5 });
      sh.setCell(r, 4, { v: '', s: 7 });
      sh.setCell(r, 5, { v: '', s: 7 });
      sh.setCell(r, 6, { v: '', s: 9 });
      sh.setCell(r, 7, { v: '', s: 8 });
      sh.setCell(r, 8, { v: '', s: 8 });
      sh.setCell(r, 9, { v: '', s: 8 });
      sh.setCell(r, 10, { v: '', s: 8 });
      sh.setCell(r, 11, { v: '', s: 8 });
      sh.setCell(r, 12, { f: `IF(C${r}="","",G${r}+SUM(I${r}:K${r}))`, v: '', s: 13 });
      sh.setCell(r, 13, { f: `IF(C${r}="","",H${r}+SUM(I${r}:K${r}))`, v: '', s: 13 });
      sh.setCell(r, 14, {
        f: `IF(C${r}="","",IF(D${r}>O${r},"⚠️ 시작일>종료일 오류",IF(COUNTIFS($C$6:$C$55,C${r},$D$6:$D$55,"<="&O${r},$O$6:$O$55,">="&D${r})>1,"⚠️ 기간 중복","정상")))`,
        v: '',
        s: 10
      });
      sh.setCell(r, 15, { f: `IF(C${r}="","",IF(E${r}="",DATE(2099,12,31),E${r}))`, v: '', s: 11 });
      sh.setCell(r, 16, { v: '', s: 6 });
    }
  }

  sh.addValidation({
    type: 'list',
    sqref: 'C6:C55',
    formula1: "'6_사용설명서_및_기준관리'!$B$6:$B$25",
    promptTitle: '판매채널 선택',
    prompt: '비용을 설정할 판매채널을 선택하세요.'
  });

  sh.addValidation({
    type: 'date',
    operator: 'greaterThanOrEqual',
    sqref: 'D6:E55',
    formula1: '43831', // 2020-01-01
    promptTitle: '적용 시작일 / 종료일 입력',
    prompt: 'YYYY-MM-DD 형식으로 날짜를 입력하거나 선택하세요. (종료일은 비워두면 계속 적용됩니다.)'
  });

  sh.addCondFormat({
    sqref: 'N6:N55',
    rules: [{ type: 'containsText', text: '⚠️', topLeftCell: 'N6', dxfId: 1 }]
  });

  return sh;
}

// ============================================================================
// 시트 4: 4_판매채널별_손익분석 (채널별 상세 손익 + 매출액 TOP 5 / 순이익 TOP 5)
// ============================================================================
function buildSheet4() {
  const sh = new SheetBuilder('4_판매채널별_손익분석', {
    freezeRow: 6,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 7 },  // 순위
      { min: 3, max: 3, width: 16 }, // 판매채널명
      { min: 4, max: 6, width: 13 }, // 10kg미만 건수, 10kg이상 건수, 총 주문건수
      { min: 7, max: 7, width: 15 }, // 매출액
      { min: 8, max: 8, width: 12 }, // 매출 비중
      { min: 9, max: 13, width: 15 }, // 매출수수료, 건수비용, 추가비용, 총비용, 순이익(전)
      { min: 14, max: 14, width: 12 }, // 순이익률(전)
      { min: 15, max: 16, width: 15 }, // 공통비용배분, 최종순이익(후)
      { min: 17, max: 17, width: 3 },
      { min: 18, max: 28, width: 13 }  // 우측 자동정렬 보조 영역
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '매출·지출관리 스프레드시트(순이익/실적 자동계산) | 판매채널별 손익 상세 분석', s: 1 });
  sh.merge('B2:P2');

  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 25);
  sh.setCell(3, 2, { v: '연동 시작일', s: 4 });
  sh.setCell(3, 3, { v: '연동 종료일', s: 4 });
  sh.setCell(3, 4, { v: '표/그래프 정렬 기준 (드롭다운)', s: 3 });
  sh.merge('D3:F3');
  sh.setCell(3, 5, { v: '', s: 3 });
  sh.setCell(3, 6, { v: '', s: 3 });
  sh.setCell(3, 7, { v: '공통비용 배분 기준', s: 4 });
  sh.merge('G3:H3');
  sh.setCell(3, 8, { v: '', s: 4 });
  sh.setCell(3, 9, { v: '기간 공통비용 합계', s: 4 });
  sh.setCell(3, 10, { v: '안내 사항', s: 4 });
  sh.merge('J3:P3');
  for (let c = 11; c <= 16; c++) sh.setCell(3, c, { v: '', s: 4 });

  sh.setCell(4, 2, { f: "'1_매출손익_대시보드'!F4", v: defaultStartSerial, s: 11 });
  sh.setCell(4, 3, { f: "'1_매출손익_대시보드'!G4", v: defaultEndSerial, s: 11 });
  sh.setCell(4, 4, { v: '순이익 높은순', s: 5 });
  sh.merge('D4:F4');
  sh.setCell(4, 5, { v: '', s: 5 });
  sh.setCell(4, 6, { v: '', s: 5 });
  sh.setCell(4, 7, { f: "'5_목표매출_및_공통비용'!K4", v: '매출액 비중 배분', s: 10 });
  sh.merge('G4:H4');
  sh.setCell(4, 8, { v: '', s: 10 });
  sh.setCell(4, 9, { f: "'5_목표매출_및_공통비용'!M26", v: periodCommonCost, s: 12 });
  sh.setCell(4, 10, { v: '💡 판매채널별 10kg 미만/이상 주문건수와 비용·순이익을 분석합니다. 적자(<0) 항목은 붉은색으로 강조됩니다.', s: 22 });
  sh.merge('J4:P4');
  for (let c = 11; c <= 16; c++) sh.setCell(4, c, { v: '', s: 22 });

  sh.setRowHeight(6, 24);
  sh.setCell(6, 2, { v: '🛒 판매채널별 손익 분석표 (매출액순 / 순이익순 자동 정렬)', s: 2 });
  sh.merge('B6:P6');
  sh.setCell(6, 18, { v: '[자동정렬용 원본 집계 데이터]', s: 19 });
  sh.merge('R6:AA6');

  sh.setRowHeight(7, 28);
  const analysisHeaders = [
    [2, '순위'],
    [3, '판매채널명'],
    [4, '10kg 미만\n주문건수'],
    [5, '10kg 이상\n주문건수'],
    [6, '총 주문건수'],
    [7, '매출액 (원)'],
    [8, '매출 비중'],
    [9, '매출액 수수료'],
    [10, '주문건수 비용\n(택배+박스+비닐+리플렛)'],
    [11, '추가 비용'],
    [12, '총비용\n(공통비 제외)'],
    [13, '순이익\n(공통비 전)'],
    [14, '순이익률\n(공통비 전)'],
    [15, '공통비용\n배분액'],
    [16, '최종 순이익\n(공통비 후)']
  ];
  for (const [c, t] of analysisHeaders) sh.setCell(7, c, { v: t, s: 4 });

  const rawHeaders = [
    [18, '정렬점수'],
    [19, '채널명'],
    [20, '미만건수'],
    [21, '이상건수'],
    [22, '총건수'],
    [23, '매출액'],
    [24, '매출수수료'],
    [25, '건수비용'],
    [26, '추가비용'],
    [27, '총비용'],
    [28, '순이익']
  ];
  for (const [c, t] of rawHeaders) sh.setCell(7, c, { v: t, s: 19 });

  const channelRawStats = CHANNELS.map((chName, idx) => {
    const rows = periodSales.filter(x => x.channel === chName);
    const u10 = rows.reduce((a, b) => a + b.ordersUnder10, 0);
    const o10 = rows.reduce((a, b) => a + b.ordersOver10, 0);
    const orders = u10 + o10;
    const sales = rows.reduce((a, b) => a + b.sales, 0);
    const fee = rows.reduce((a, b) => a + b.feeAmt, 0);
    const orderCost = rows.reduce((a, b) => a + b.orderCost, 0);
    const extra = rows.reduce((a, b) => a + b.extraCost, 0);
    const totalCost = fee + orderCost + extra;
    const netProfit = sales - totalCost;
    const sortKey = netProfit + (100 - (8 + idx)) * 0.00001;
    return { chName, u10, o10, orders, sales, fee, orderCost, extra, totalCost, netProfit, sortKey, idx };
  });

  const sortedChannels = [...channelRawStats].sort((a, b) => b.sortKey - a.sortKey);

  CHANNELS.forEach((_, idx) => {
    const r = 8 + idx;
    const raw = channelRawStats[idx];

    // 우측 원본 영역 (R8:AB14 -> 18~28열)
    sh.setCell(r, 18, {
      f: `IF($D$4="매출액 높은순",W${r},IF($D$4="순이익 높은순",AB${r},-ROW()))+(100-ROW())*0.00001`,
      v: raw.sortKey,
      s: 25
    });
    sh.setCell(r, 19, { f: `'6_사용설명서_및_기준관리'!B${6 + idx}`, v: raw.chName, s: 23 });
    sh.setCell(r, 20, {
      f: `SUMIFS('2_일일매출_입력'!$D$6:$D$305,'2_일일매출_입력'!$C$6:$C$305,S${r},'2_일일매출_입력'!$B$6:$B$305,">="&$B$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$C$4)`,
      v: raw.u10,
      s: 25
    });
    sh.setCell(r, 21, {
      f: `SUMIFS('2_일일매출_입력'!$E$6:$E$305,'2_일일매출_입력'!$C$6:$C$305,S${r},'2_일일매출_입력'!$B$6:$B$305,">="&$B$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$C$4)`,
      v: raw.o10,
      s: 25
    });
    sh.setCell(r, 22, { f: `T${r}+U${r}`, v: raw.orders, s: 25 });
    sh.setCell(r, 23, {
      f: `SUMIFS('2_일일매출_입력'!$G$6:$G$305,'2_일일매출_입력'!$C$6:$C$305,S${r},'2_일일매출_입력'!$B$6:$B$305,">="&$B$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$C$4)`,
      v: raw.sales,
      s: 25
    });
    sh.setCell(r, 24, {
      f: `SUMIFS('2_일일매출_입력'!$L$6:$L$305,'2_일일매출_입력'!$C$6:$C$305,S${r},'2_일일매출_입력'!$B$6:$B$305,">="&$B$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$C$4)`,
      v: raw.fee,
      s: 25
    });
    sh.setCell(r, 25, {
      f: `SUMIFS('2_일일매출_입력'!$O$6:$O$305,'2_일일매출_입력'!$C$6:$C$305,S${r},'2_일일매출_입력'!$B$6:$B$305,">="&$B$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$C$4)`,
      v: raw.orderCost,
      s: 25
    });
    sh.setCell(r, 26, {
      f: `SUMIFS('2_일일매출_입력'!$H$6:$H$305,'2_일일매출_입력'!$C$6:$C$305,S${r},'2_일일매출_입력'!$B$6:$B$305,">="&$B$4,'2_일일매출_입력'!$B$6:$B$305,"<="&$C$4)`,
      v: raw.extra,
      s: 25
    });
    sh.setCell(r, 27, { f: `SUM(X${r}:Z${r})`, v: raw.totalCost, s: 25 });
    sh.setCell(r, 28, { f: `W${r}-AA${r}`, v: raw.netProfit, s: 25 });

    // 좌측 정렬된 메인 표 (B8:P14)
    const st = sortedChannels[idx];
    const share = periodActualSales > 0 ? st.sales / periodActualSales : 0;
    const margin = st.sales > 0 ? st.netProfit / st.sales : 0;
    const allocCommon = Math.round(periodCommonCost * share);
    const finalProfit = st.netProfit - allocCommon;

    sh.setCell(r, 2, { v: idx + 1, s: 10 });
    sh.setCell(r, 3, { f: `INDEX($S$8:$S$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.chName, s: 23 });
    sh.setCell(r, 4, { f: `INDEX($T$8:$T$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.u10, s: 25 });
    sh.setCell(r, 5, { f: `INDEX($U$8:$U$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.o10, s: 25 });
    sh.setCell(r, 6, { f: `D${r}+E${r}`, v: st.orders, s: 12 });
    sh.setCell(r, 7, { f: `INDEX($W$8:$W$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.sales, s: 25 });
    sh.setCell(r, 8, { f: `IFERROR(G${r}/$G$15,0)`, v: share, s: 14 });
    sh.setCell(r, 9, { f: `INDEX($X$8:$X$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.fee, s: 25 });
    sh.setCell(r, 10, { f: `INDEX($Y$8:$Y$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.orderCost, s: 25 });
    sh.setCell(r, 11, { f: `INDEX($Z$8:$Z$14,MATCH(LARGE($R$8:$R$14,B${r}),$R$8:$R$14,0))`, v: st.extra, s: 25 });
    sh.setCell(r, 12, { f: `SUM(I${r}:K${r})`, v: st.totalCost, s: 12 });
    sh.setCell(r, 13, { f: `G${r}-L${r}`, v: st.netProfit, s: 13 });
    sh.setCell(r, 14, { f: `IFERROR(M${r}/G${r},0)`, v: margin, s: 14 });
    sh.setCell(r, 15, {
      f: `IF($G$4="매출액 비중 배분",ROUND($I$4*H${r},0),IF($G$4="주문건수 비중 배분",ROUND($I$4*IFERROR(F${r}/$F$15,0),0),0))`,
      v: allocCommon,
      s: 12
    });
    sh.setCell(r, 16, { f: `M${r}-O${r}`, v: finalProfit, s: 12 });
  });

  // 15행: 판매채널 합계
  sh.setCell(15, 2, { v: '합계', s: 16 });
  sh.setCell(15, 3, { v: `전체 ${CHANNELS.length}개 판매채널`, s: 16 });
  sh.setCell(15, 4, { f: 'SUM(D8:D14)', v: periodOrdersUnder10, s: 17 });
  sh.setCell(15, 5, { f: 'SUM(E8:E14)', v: periodOrdersOver10, s: 17 });
  sh.setCell(15, 6, { f: 'SUM(F8:F14)', v: periodOrders, s: 17 });
  sh.setCell(15, 7, { f: 'SUM(G8:G14)', v: periodActualSales, s: 17 });
  sh.setCell(15, 8, { f: 'IFERROR(SUM(H8:H14),0)', v: 1, s: 18 });
  sh.setCell(15, 9, { f: 'SUM(I8:I14)', v: periodFee, s: 17 });
  sh.setCell(15, 10, { f: 'SUM(J8:J14)', v: periodOrderCost, s: 17 });
  sh.setCell(15, 11, { f: 'SUM(K8:K14)', v: periodExtraCost, s: 17 });
  sh.setCell(15, 12, { f: 'SUM(L8:L14)', v: periodTotalCost, s: 17 });
  sh.setCell(15, 13, { f: 'G15-L15', v: periodNetProfit, s: 17 });
  sh.setCell(15, 14, { f: 'IFERROR(M15/G15,0)', v: periodMargin, s: 18 });
  sh.setCell(15, 15, { f: 'SUM(O8:O14)', v: periodCommonCost, s: 17 });
  sh.setCell(15, 16, { f: 'M15-O15', v: periodFinalProfit, s: 17 });

  sh.setCell(17, 2, { v: '📈 판매채널별 매출·순이익 비교 그래프', s: 2 });

  sh.addValidation({
    type: 'list',
    sqref: 'D4:F4',
    formula1: '"순이익 높은순,매출액 높은순,입력(기본)순"',
    promptTitle: '정렬 기준 선택',
    prompt: '순이익 높은순 또는 매출액 높은순을 선택하면 표와 그래프가 자동 재정렬됩니다.'
  });

  sh.addCondFormat({
    sqref: 'M8:P15',
    rules: [{ type: 'cellIs', operator: 'lessThan', formula: '0', dxfId: 0 }]
  });

  sh.drawingRelId = 'rId1';

  return { sh, sortedChannels };
}

// ============================================================================
// 시트 5: 5_목표매출_및_공통비용 (2026년 11월 ~ 2027년 10월 + 목표매출 합계 강조)
// ============================================================================
function buildSheet5() {
  const sh = new SheetBuilder('5_목표매출_및_공통비용', {
    freezeRow: 6,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 12 }, // 연월(26년 11월~)
      { min: 3, max: 4, width: 15 }, // 시작일, 종료일
      { min: 5, max: 7, width: 18 }, // 월별 목표매출, 일할 목표, 조회기간 반영 목표
      { min: 8, max: 8, width: 4 },
      { min: 9, max: 9, width: 15 }, // 공통비용 일자
      { min: 10, max: 10, width: 20 }, // 항목
      { min: 11, max: 11, width: 16 }, // 금액
      { min: 12, max: 13, width: 14 }, // 귀속연월, 조회기간 포함여부
      { min: 14, max: 14, width: 32 }  // 비고
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '매출·지출관리 스프레드시트(순이익/실적 자동계산) | 목표매출(2026.11~2027.10) 및 공통비용 설정', s: 1 });
  sh.merge('B2:N2');

  sh.setRowHeight(3, 34);
  sh.setCell(3, 2, {
    v: '💡 [목표설정 기간: 2026년 11월 ~ 2027년 10월] 노란색 셀에 월별 목표매출을 입력하면 상단과 하단에 [12개월 목표매출 총합계]와 [현재 조회기간 목표매출 합계]가 자동 계산됩니다.',
    s: 22
  });
  sh.merge('B3:N3');
  for (let c = 3; c <= 14; c++) sh.setCell(3, c, { v: '', s: 22 });

  // 4행: 목표매출 총합계 확인 요약 바 & 공통비용 배분 기준
  sh.setRowHeight(4, 26);
  sh.setCell(4, 2, { v: '12개월 목표매출 총합계', s: 4 });
  sh.merge('B4:C4');
  sh.setCell(4, 3, { v: '', s: 4 });
  sh.setCell(4, 4, { f: 'SUM(E7:E18)', v: totalAnnualTarget, s: 13 });
  sh.setCell(4, 5, { v: '현재 조회기간 목표매출 합계', s: 4 });
  sh.merge('E4:F4');
  sh.setCell(4, 6, { v: '', s: 4 });
  sh.setCell(4, 7, { f: 'SUM(G7:G18)', v: periodTarget, s: 13 });

  sh.setCell(4, 9, { v: '공통비용 배분 기준 선택', s: 3 });
  sh.merge('I4:J4');
  sh.setCell(4, 10, { v: '', s: 3 });
  sh.setCell(4, 11, { v: '매출액 비중 배분', s: 5 });
  sh.merge('K4:L4');
  sh.setCell(4, 12, { v: '', s: 5 });

  // 6행: 좌·우 테이블 헤더
  sh.setRowHeight(6, 28);
  const leftHeaders = [
    [2, '대상 연월\n(26.11~27.10)', 4],
    [3, '월 시작일', 4],
    [4, '월 종료일', 4],
    [5, '월별 목표매출 (원)\n(직접입력)', 3],
    [6, '일할 목표매출 (원/일)\n(자동계산)', 4],
    [7, '조회기간 반영 목표매출\n(대시보드 연동)', 4]
  ];
  for (const [c, t, st] of leftHeaders) sh.setCell(6, c, { v: t, s: st });

  const rightHeaders = [
    [9, '지출/귀속일자\n(직접입력)', 3],
    [10, '공통비용 항목\n(광고비/임대료 등)', 3],
    [11, '금액 (원)\n(직접입력)', 3],
    [12, '귀속 연월\n(자동표시)', 4],
    [13, '현재 조회기간\n포함 여부', 4],
    [14, '비고 (상세 내역 메모)', 3]
  ];
  for (const [c, t, st] of rightHeaders) sh.setCell(6, c, { v: t, s: st });

  // 좌측: 2026년 11월 ~ 2027년 10월 (7행 ~ 18행)
  TARGET_MONTHS.forEach((tm, idx) => {
    const r = 7 + idx;
    const startSer = dateToSerial(tm.start);
    const endSer = dateToSerial(tm.end);
    const dailyTarget = Math.round(tm.target / tm.days);
    const appliedTarget = tm.ym === defaultYm ? tm.target : 0;

    sh.setCell(r, 2, { v: tm.label, s: 10 });
    sh.setCell(r, 3, { v: startSer, s: 11 });
    sh.setCell(r, 4, { v: endSer, s: 11 });
    sh.setCell(r, 5, { v: tm.target, s: 8 });
    sh.setCell(r, 6, { f: `ROUND(E${r}/DAY(D${r}),0)`, v: dailyTarget, s: 12 });
    sh.setCell(r, 7, {
      f: `IF(OR('1_매출손익_대시보드'!$G$4<C${r},'1_매출손익_대시보드'!$F$4>D${r}),0,IF(AND('1_매출손익_대시보드'!$F$4<=C${r},'1_매출손익_대시보드'!$G$4>=D${r}),E${r},ROUND(E${r}*(MIN('1_매출손익_대시보드'!$G$4,D${r})-MAX('1_매출손익_대시보드'!$F$4,C${r})+1)/DAY(D${r}),0)))`,
      v: appliedTarget,
      s: 13
    });
  });

  // 19행: 목표매출 합계
  sh.setCell(19, 2, { v: '목표매출 합계', s: 16 });
  sh.setCell(19, 3, { v: '2026.11 ~ 2027.10', s: 16 });
  sh.setCell(19, 4, { v: '총 12개월', s: 16 });
  sh.setCell(19, 5, { f: 'SUM(E7:E18)', v: totalAnnualTarget, s: 17 });
  sh.setCell(19, 6, { f: 'ROUND(E19/365,0)', v: Math.round(totalAnnualTarget / 365), s: 17 });
  sh.setCell(19, 7, { f: 'SUM(G7:G18)', v: periodTarget, s: 17 });

  // 우측 공통비용 입력 행 (7행 ~ 25행)
  const maxCommonRows = 19;
  for (let i = 0; i < maxCommonRows; i++) {
    const r = 7 + i;
    const cItem = COMMON_COSTS[i];
    if (cItem) {
      const dSer = dateToSerial(cItem.date);
      const ym = cItem.date.slice(0, 7);
      const inc = (dSer >= defaultStartSerial && dSer <= defaultEndSerial) ? '포함' : '제외';

      sh.setCell(r, 9, { v: dSer, s: 7 });
      sh.setCell(r, 10, { v: cItem.category, s: 5 });
      sh.setCell(r, 11, { v: cItem.amount, s: 8 });
      sh.setCell(r, 12, { f: `IF(I${r}="","",TEXT(I${r},"yyyy-mm"))`, v: ym, s: 10 });
      sh.setCell(r, 13, {
        f: `IF(I${r}="","",IF(AND(I${r}>='1_매출손익_대시보드'!$F$4,I${r}<='1_매출손익_대시보드'!$G$4),"포함","제외"))`,
        v: inc,
        s: 10
      });
      sh.setCell(r, 14, { v: cItem.note, s: 6 });
    } else {
      sh.setCell(r, 9, { v: '', s: 7 });
      sh.setCell(r, 10, { v: '', s: 5 });
      sh.setCell(r, 11, { v: '', s: 8 });
      sh.setCell(r, 12, { f: `IF(I${r}="","",TEXT(I${r},"yyyy-mm"))`, v: '', s: 10 });
      sh.setCell(r, 13, {
        f: `IF(I${r}="","",IF(AND(I${r}>='1_매출손익_대시보드'!$F$4,I${r}<='1_매출손익_대시보드'!$G$4),"포함","제외"))`,
        v: '',
        s: 10
      });
      sh.setCell(r, 14, { v: '', s: 6 });
    }
  }

  const allCommonSum = COMMON_COSTS.reduce((a, b) => a + b.amount, 0);
  sh.setCell(26, 9, { v: '합계', s: 16 });
  sh.setCell(26, 10, { v: '전체 누적 공통비용', s: 16 });
  sh.setCell(26, 11, { f: 'SUM(K7:K25)', v: allCommonSum, s: 17 });
  sh.setCell(26, 12, { v: '조회기간 공통비용 →', s: 16 });
  sh.setCell(26, 13, {
    f: 'SUMIFS(K7:K25,I7:I25,">="&\'1_매출손익_대시보드\'!$F$4,I7:I25,"<="&\'1_매출손익_대시보드\'!$G$4)',
    v: periodCommonCost,
    s: 17
  });
  sh.setCell(26, 14, { v: '← 대시보드 및 손익분석 시트에 자동 반영됨', s: 16 });

  sh.addValidation({
    type: 'list',
    sqref: 'K4:L4',
    formula1: '"매출액 비중 배분,주문건수 비중 배분,배분 안 함(전체에서만 차감)"',
    promptTitle: '공통비용 배분 기준',
    prompt: '채널별 손익 분석 시 공통비용을 어떤 기준으로 배분할지 선택하세요.'
  });

  return sh;
}

// ============================================================================
// 시트 6: 6_사용설명서_및_기준관리
// ============================================================================
function buildSheet6() {
  const sh = new SheetBuilder('6_사용설명서_및_기준관리', {
    freezeRow: 5,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 22 }, // 판매채널 마스터
      { min: 3, max: 3, width: 4 },
      { min: 4, max: 4, width: 26 }, // 가이드 항목
      { min: 5, max: 5, width: 85 }  // 상세 설명
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '매출·지출관리 스프레드시트(순이익/실적 자동계산) | 사용 설명서 및 판매채널 기준 관리', s: 1 });
  sh.merge('B2:E2');

  sh.setRowHeight(3, 32);
  sh.setCell(3, 2, {
    v: '💡 좌측의 [판매채널 목록]의 빈 노란색 칸에 새로운 채널 이름을 입력하시면, [일일 매출 입력] 및 [채널별 비용설정] 시트의 드롭다운 목록에 자동으로 추가됩니다.',
    s: 22
  });
  sh.merge('B3:E3');
  for (let c = 3; c <= 5; c++) sh.setCell(3, c, { v: '', s: 22 });

  sh.setRowHeight(5, 26);
  sh.setCell(5, 2, { v: '🛒 판매채널 등록 목록\n(기본 7종 + 추가 가능)', s: 3 });
  sh.setCell(5, 4, { v: '📌 구분 / 핵심 주제', s: 4 });
  sh.setCell(5, 5, { v: '📖 상세 사용 방법 및 계산 기준 안내 (비개발자 맞춤 가이드)', s: 4 });

  for (let i = 0; i < 20; i++) {
    const r = 6 + i;
    sh.setCell(r, 2, { v: CHANNELS[i] || '', s: 5 });
  }

  const guides = [
    ['1. 목표설정 기간 (26.11~27.10)', '• 목표매출 설정 기간은 [2026년 11월]부터 [2027년 10월]까지 12개월로 설정되어 있습니다.\n• [5_목표매출_및_공통비용] 시트 상단과 하단에서 12개월 목표매출 총합계와 현재 조회기간의 목표매출 합계를 바로 확인할 수 있습니다.'],
    ['2. 10kg 미만 / 10kg 이상 주문 구분', '• [2_일일매출_입력] 시트에서 하루 주문을 [10kg 미만 주문건수]와 [10kg 이상 주문건수]로 나누어 입력합니다.\n• 무게 구간별로 택배비가 다르게 설정되어 있어 각 구간 건수에 맞는 택배비와 부자재비(박스+포장비닐+리플렛)가 자동 합산됩니다.'],
    ['3. 박스·포장비닐·리플렛 비용 설정', '• [3_채널별_비용설정] 시트에서 채널별·기간별로 [10kg 미만 배송비], [10kg 이상 배송비], 그리고 주문 1건당 [박스 비용], [포장비닐 비용], [리플렛 비용]을 각각 등록합니다.\n• 10kg 미만 1건당 비용 = 10kg 미만 배송비 + 박스 + 포장비닐 + 리플렛\n• 10kg 이상 1건당 비용 = 10kg 이상 배송비 + 박스 + 포장비닐 + 리플렛'],
    ['4. 비용 변경 시 기간 이력 관리', '• 택배비나 박스비, 수수료율이 바뀌면 기존 설정 행을 덮어쓰지 마세요.\n• 기존 설정의 [적용 종료일]을 입력한 뒤, 아래 새 행에 [새로운 적용 시작일]과 변경된 단가를 등록하면 과거 매출에는 당시 단가가 그대로 유지됩니다.'],
    ['5. 계산 공식 요약', '• 매출액 기준 수수료 = 일일 매출액 × 해당 날짜의 채널 수수료율\n• 주문건수에 따른 비용 = (10kg 미만 주문건수 × 10kg 미만 1건당 비용) + (10kg 이상 주문건수 × 10kg 이상 1건당 비용)\n• 총비용 = 매출액 기준 수수료 + 주문건수에 따른 비용 + 추가 비용\n• 순이익 = 일일 매출액 − 총비용  |  순이익률 = 순이익 ÷ 일일 매출액 × 100']
  ];

  guides.forEach(([title, desc], idx) => {
    const r = 6 + idx * 2;
    sh.setRowHeight(r, 38);
    sh.setRowHeight(r + 1, 18);
    sh.setCell(r, 4, { v: title, s: 16 });
    sh.setCell(r + 1, 4, { v: '', s: 16 });
    sh.merge(`D${r}:D${r + 1}`);

    sh.setCell(r, 5, { v: desc, s: 24 });
    sh.setCell(r + 1, 5, { v: '', s: 24 });
    sh.merge(`E${r}:E${r + 1}`);
  });

  return sh;
}

// ============================================================================
// 전체 워크북 생성
// ============================================================================
function generateAll() {
  const { sh: sheet1, monthlyChartCache, channelChartCache } = buildSheet1();
  const sheet2 = buildSheet2();
  const sheet3 = buildSheet3();
  const { sh: sheet4, sortedChannels } = buildSheet4();
  const sheet5 = buildSheet5();
  const sheet6 = buildSheet6();

  const sheets = [sheet1, sheet2, sheet3, sheet4, sheet5, sheet6];

  const chart1Xml = buildBarChartXml({
    title: '전체 매출성과 (목표매출 vs 총 매출액 vs 총 순이익)',
    categoriesFormula: "'1_매출손익_대시보드'!$K$40:$N$40",
    categoriesCache: ['목표매출', '총 매출액', '총 순이익', '최종순이익(공통비후)'],
    seriesList: [
      {
        name: '선택기간 실적 (원)',
        nameFormula: "'1_매출손익_대시보드'!$J$41",
        color: '2B4CBE',
        valuesFormula: "'1_매출손익_대시보드'!$K$41:$N$41",
        valuesCache: [periodTarget, periodActualSales, periodNetProfit, periodFinalProfit]
      }
    ]
  });

  const chart2Xml = buildBarChartXml({
    title: '월별 매출 · 총비용 · 순이익 추이 (26.11 ~ 27.10)',
    categoriesFormula: "'1_매출손익_대시보드'!$B$31:$B$42",
    categoriesCache: monthlyChartCache.cats,
    showVal: false,
    seriesList: [
      {
        name: '총 매출',
        nameFormula: "'1_매출손익_대시보드'!$D$30",
        color: '2B4CBE',
        valuesFormula: "'1_매출손익_대시보드'!$D$31:$D$42",
        valuesCache: monthlyChartCache.sales
      },
      {
        name: '총 지출(비용)',
        nameFormula: "'1_매출손익_대시보드'!$E$30",
        color: '92ADD0',
        valuesFormula: "'1_매출손익_대시보드'!$E$31:$E$42",
        valuesCache: monthlyChartCache.cost
      },
      {
        name: '순이익',
        nameFormula: "'1_매출손익_대시보드'!$F$30",
        color: 'C5D185',
        valuesFormula: "'1_매출손익_대시보드'!$F$31:$F$42",
        valuesCache: monthlyChartCache.profit
      }
    ]
  });

  const chart3Xml = buildDoughnutChartXml({
    title: '판매채널별 매출 비중 (금액 및 비율 %)',
    categoriesFormula: "'1_매출손익_대시보드'!$J$31:$J$37",
    categoriesCache: channelChartCache.cats,
    seriesNameFormula: "'1_매출손익_대시보드'!$L$30",
    seriesName: '매출액 (원)',
    valuesFormula: "'1_매출손익_대시보드'!$L$31:$L$37",
    valuesCache: channelChartCache.sales
  });

  const chart4Xml = buildBarChartXml({
    title: '판매채널별 순이익 비교 (공통비 전 vs 후)',
    categoriesFormula: "'4_판매채널별_손익분석'!$C$8:$C$14",
    categoriesCache: sortedChannels.map(c => c.chName),
    seriesList: [
      {
        name: '순이익(공통비 전)',
        nameFormula: "'4_판매채널별_손익분석'!$M$7",
        color: '2B4CBE',
        valuesFormula: "'4_판매채널별_손익분석'!$M$8:$M$14",
        valuesCache: sortedChannels.map(c => c.netProfit)
      },
      {
        name: '최종 순이익(공통비 후)',
        nameFormula: "'4_판매채널별_손익분석'!$P$7",
        color: 'C5D185',
        valuesFormula: "'4_판매채널별_손익분석'!$P$8:$P$14",
        valuesCache: sortedChannels.map(c => {
          const share = periodActualSales > 0 ? c.sales / periodActualSales : 0;
          return c.netProfit - Math.round(periodCommonCost * share);
        })
      }
    ]
  });

  const chart5Xml = buildBarChartXml({
    title: '판매채널별 매출액 및 순이익 비교',
    categoriesFormula: "'4_판매채널별_손익분석'!$C$8:$C$14",
    categoriesCache: sortedChannels.map(c => c.chName),
    seriesList: [
      {
        name: '매출액 (원)',
        nameFormula: "'4_판매채널별_손익분석'!$G$7",
        color: '2B4CBE',
        valuesFormula: "'4_판매채널별_손익분석'!$G$8:$G$14",
        valuesCache: sortedChannels.map(c => c.sales)
      },
      {
        name: '순이익(공통비 전)',
        nameFormula: "'4_판매채널별_손익분석'!$M$7",
        color: 'C5D185',
        valuesFormula: "'4_판매채널별_손익분석'!$M$8:$M$14",
        valuesCache: sortedChannels.map(c => c.netProfit)
      }
    ]
  });

  const drawing1Xml = buildDrawingXml([
    { fromCol: 1, fromRow: 10, toCol: 5, toRow: 27, relId: 'rId1' },
    { fromCol: 5, fromRow: 10, toCol: 10, toRow: 27, relId: 'rId2' },
    { fromCol: 10, fromRow: 10, toCol: 15, toRow: 27, relId: 'rId3' }
  ]);

  const drawing2Xml = buildDrawingXml([
    { fromCol: 1, fromRow: 18, toCol: 8, toRow: 34, relId: 'rId1' },
    { fromCol: 8, fromRow: 18, toCol: 16, toRow: 34, relId: 'rId2' }
  ]);

  const files = [];

  files.push({
    name: '[Content_Types].xml',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart4.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart5.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`
  });

  files.push({
    name: '_rels/.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
  });

  files.push({
    name: 'xl/workbook.xml',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets>
    ${sheets.map((s, idx) => `<sheet name="${escXml(s.name)}" sheetId="${idx + 1}" r:id="rId${idx + 1}"/>`).join('\n    ')}
  </sheets>
  <calcPr calcId="191029" fullCalcOnLoad="1"/>
</workbook>`
  });

  files.push({
    name: 'xl/_rels/workbook.xml.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheets.map((_, idx) => `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`).join('\n  ')}
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  });

  files.push({ name: 'xl/styles.xml', data: buildStylesXml() });

  sheets.forEach((s, idx) => {
    files.push({ name: `xl/worksheets/sheet${idx + 1}.xml`, data: s.toXml() });
  });

  files.push({
    name: 'xl/worksheets/_rels/sheet1.xml.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`
  });

  files.push({
    name: 'xl/worksheets/_rels/sheet4.xml.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>
</Relationships>`
  });

  files.push({ name: 'xl/drawings/drawing1.xml', data: drawing1Xml });
  files.push({ name: 'xl/drawings/drawing2.xml', data: drawing2Xml });

  files.push({
    name: 'xl/drawings/_rels/drawing1.xml.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
</Relationships>`
  });

  files.push({
    name: 'xl/drawings/_rels/drawing2.xml.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart5.xml"/>
</Relationships>`
  });

  files.push({ name: 'xl/charts/chart1.xml', data: chart1Xml });
  files.push({ name: 'xl/charts/chart2.xml', data: chart2Xml });
  files.push({ name: 'xl/charts/chart3.xml', data: chart3Xml });
  files.push({ name: 'xl/charts/chart4.xml', data: chart4Xml });
  files.push({ name: 'xl/charts/chart5.xml', data: chart5Xml });

  const zipBuf = createZipBuffer(files);
  const outPath = path.join(__dirname, '거래처별_매출순이익_관리프로그램.xlsx');
  fs.writeFileSync(outPath, zipBuf);
  console.log('Successfully generated updated Excel:', outPath, 'Size:', zipBuf.length, 'bytes');
}

generateAll();
