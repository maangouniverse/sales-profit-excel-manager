const fs = require('fs');
const path = require('path');
const { createZipBuffer, escXml, colToLetter, dateToSerial, SheetBuilder } = require('./xlsx_writer');
const { buildBarChartXml, buildDoughnutChartXml, buildDrawingXml } = require('./chart_builder');
const {
  CHANNELS,
  CLIENTS,
  CHANNEL_RATES,
  DAILY_SALES,
  MONTHLY_TARGETS,
  COMMON_COSTS,
  computeRowMetrics,
  buildStylesXml
} = require('./excel_data_styles');

// 사전 계산: 전체 일일 매출 데이터
const enrichedSales = DAILY_SALES.map(r => ({
  ...r,
  ...computeRowMetrics(r, CHANNEL_RATES)
}));

// 기본 조회 기간: 2026년 4월 (2026-04-01 ~ 2026-04-30)
const defaultYear = 2026;
const defaultMonth = 4;
const defaultStartStr = '2026-04-01';
const defaultEndStr = '2026-04-30';
const defaultStartSerial = dateToSerial(defaultStartStr);
const defaultEndSerial = dateToSerial(defaultEndStr);

const periodSales = enrichedSales.filter(r => {
  const s = dateToSerial(r.date);
  return s >= defaultStartSerial && s <= defaultEndSerial && r.status === '정상';
});

const periodTarget = MONTHLY_TARGETS.find(m => m.month === defaultMonth).target;
const periodActualSales = periodSales.reduce((a, b) => a + b.sales, 0);
const periodOrders = periodSales.reduce((a, b) => a + b.orders, 0);
const periodCogs = periodSales.reduce((a, b) => a + b.cogs, 0);
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
  sh.setCell(2, 2, { v: '📊 거래처별·채널별 매출 및 손익 종합 대시보드', s: 1 });
  sh.merge('B2:H2');
  sh.setCell(2, 9, { v: '💡 노란색 셀(조회 연·월 또는 직접입력 시작·종료일)을 변경하면 모든 실적과 그래프가 즉시 갱신됩니다.', s: 22 });
  sh.merge('I2:O2');

  // 3~4행: 기간 선택 및 데이터 점검 배너
  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 26);

  const ctrlHeaders = [
    [2, '조회 연도(선택)', 3],
    [3, '조회 월(1~12/전체)', 3],
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

  sh.setCell(4, 2, { v: defaultYear, s: 5 });
  sh.setCell(4, 3, { v: defaultMonth, s: 5 });
  sh.setCell(4, 4, { v: '', s: 7 }); // 직접 시작일 (비워두면 연·월 기준 자동)
  sh.setCell(4, 5, { v: '', s: 7 }); // 직접 종료일 (비워두면 연·월 기준 자동)
  sh.setCell(4, 6, {
    f: 'IF(D4<>"",D4,IF(C4="전체",DATE(B4,1,1),DATE(B4,C4,1)))',
    v: defaultStartSerial,
    s: 11
  });
  sh.setCell(4, 7, {
    f: 'IF(E4<>"",E4,IF(C4="전체",DATE(B4,12,31),EOMONTH(DATE(B4,C4,1),0)))',
    v: defaultEndSerial,
    s: 11
  });
  sh.setCell(4, 8, {
    f: 'COUNTIF(\'2_일일매출_입력\'!$J$6:$J$505,"비용 설정 필요")',
    v: 0,
    s: 12
  });
  sh.setCell(4, 9, {
    f: 'COUNTIF(\'3_채널별_비용설정\'!$K$6:$K$55,"*중복*")+COUNTIF(\'3_채널별_비용설정\'!$K$6:$K$55,"*오류*")',
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

  // 6~7행: 핵심 KPI 카드 1단 (요구사항 1번 필수 지표 7종)
  sh.setRowHeight(6, 20);
  sh.setRowHeight(7, 32);
  const kpiRow1 = [
    { c1: 2, c2: 3, label: '🎯 목표매출 (원)', f: '\'5_목표매출_및_공통비용\'!G19', v: periodTarget, s: 20 },
    { c1: 4, c2: 5, label: '💰 실제매출 (원)', f: 'SUMIFS(\'2_일일매출_입력\'!$F$6:$F$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodActualSales, s: 20 },
    { c1: 6, c2: 7, label: '📈 목표 달성률 (%)', f: 'IFERROR(D7/B7,0)', v: periodTarget > 0 ? periodActualSales / periodTarget : 0, s: 21 },
    { c1: 8, c2: 9, label: '📦 총 주문건수 (건)', f: 'SUMIFS(\'2_일일매출_입력\'!$E$6:$E$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodOrders, s: 20 },
    { c1: 10, c2: 11, label: '🧾 총비용 (공통비 제외)', f: 'SUMIFS(\'2_일일매출_입력\'!$O$6:$O$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodTotalCost, s: 20 },
    { c1: 12, c2: 13, label: '💎 순이익 (공통비 반영 전)', f: 'D7-J7', v: periodNetProfit, s: 20 },
    { c1: 14, c2: 15, label: '📊 순이익률 (공통비 전)', f: 'IFERROR(L7/D7,0)', v: periodMargin, s: 21 }
  ];
  for (const k of kpiRow1) {
    sh.setCell(6, k.c1, { v: k.label, s: 19 });
    sh.setCell(6, k.c2, { v: '', s: 19 });
    sh.merge(`${colToLetter(k.c1)}6:${colToLetter(k.c2)}6`);
    sh.setCell(7, k.c1, { f: k.f, v: k.v, s: k.s });
    sh.setCell(7, k.c2, { v: '', s: k.s });
    sh.merge(`${colToLetter(k.c1)}7:${colToLetter(k.c2)}7`);
  }

  // 8~9행: 핵심 KPI 카드 2단 (비용 세부 구성 & 공통비용 차감 후 최종 순이익)
  sh.setRowHeight(8, 20);
  sh.setRowHeight(9, 28);
  const kpiRow2 = [
    { c1: 2, c2: 3, label: '상품 총원가 (원)', f: 'SUMIFS(\'2_일일매출_입력\'!$G$6:$G$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodCogs, s: 12 },
    { c1: 4, c2: 5, label: '매출액 기준 수수료 (원)', f: 'SUMIFS(\'2_일일매출_입력\'!$L$6:$L$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodFee, s: 12 },
    { c1: 6, c2: 7, label: '주문건수 비례 비용 (원)', f: 'SUMIFS(\'2_일일매출_입력\'!$N$6:$N$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodOrderCost, s: 12 },
    { c1: 8, c2: 9, label: '추가 비용 합계 (원)', f: 'SUMIFS(\'2_일일매출_입력\'!$H$6:$H$505,\'2_일일매출_입력\'!$B$6:$B$505,">="&$F$4,\'2_일일매출_입력\'!$B$6:$B$505,"<="&$G$4)', v: periodExtraCost, s: 12 },
    { c1: 10, c2: 11, label: '🏢 기간 공통비용 (광고/임대 등)', f: '\'5_목표매출_및_공통비용\'!M26', v: periodCommonCost, s: 12 },
    { c1: 12, c2: 13, label: '🏆 최종 순이익 (공통비 차감 후)', f: 'L7-J9', v: periodFinalProfit, s: 13 },
    { c1: 14, c2: 15, label: '🏆 최종 순이익률 (공통비 후)', f: 'IFERROR(L9/D7,0)', v: periodFinalMargin, s: 14 }
  ];
  for (const k of kpiRow2) {
    sh.setCell(8, k.c1, { v: k.label, s: 19 });
    sh.setCell(8, k.c2, { v: '', s: 19 });
    sh.merge(`${colToLetter(k.c1)}8:${colToLetter(k.c2)}8`);
    sh.setCell(9, k.c1, { f: k.f, v: k.v, s: k.s });
    sh.setCell(9, k.c2, { v: '', s: k.s });
    sh.merge(`${colToLetter(k.c1)}9:${colToLetter(k.c2)}9`);
  }

  // 11~27행은 차트 3개가 올라가는 시각화 영역
  sh.setCell(10, 2, { v: '📈 핵심 성과 시각화 차트 (일일 매출 입력/수정 시 자동 갱신)', s: 2 });

  // 29~42행: 하단 집계 데이터 표 (차트 원본 데이터 및 상세 요약표)
  sh.setRowHeight(29, 24);
  sh.setCell(29, 2, { v: '📅 월별 매출·비용·순이익 추이 요약 (조회 연도 기준)', s: 2 });
  sh.setCell(29, 10, { v: '🛒 판매채널별 매출 비중 및 손익 요약 (조회 기간 기준)', s: 2 });

  // 좌측 표 헤더 (B30:H30): 월별 추이
  sh.setRowHeight(30, 24);
  const mHeaders = [
    [2, '월'],
    [3, '목표매출'],
    [4, '실제매출'],
    [5, '총비용(변동)'],
    [6, '순이익(공통비전)'],
    [7, '순이익률'],
    [8, '주문건수']
  ];
  for (const [c, t] of mHeaders) sh.setCell(30, c, { v: t, s: 4 });

  const monthlyChartCache = {
    cats: [],
    target: [],
    sales: [],
    cost: [],
    profit: []
  };

  for (let m = 1; m <= 12; m++) {
    const r = 30 + m;
    const mLabel = `${m}월`;
    const mTarget = MONTHLY_TARGETS[m - 1].target;
    const mStart = dateToSerial(`2026-${String(m).padStart(2, '0')}-01`);
    const mEnd = dateToSerial(new Date(Date.UTC(2026, m, 0)).toISOString().slice(0, 10));
    const mRows = enrichedSales.filter(x => {
      const s = dateToSerial(x.date);
      return s >= mStart && s <= mEnd && x.status === '정상';
    });
    const mSales = mRows.reduce((a, b) => a + b.sales, 0);
    const mCost = mRows.reduce((a, b) => a + b.totalCost, 0);
    const mProfit = mSales - mCost;
    const mMargin = mSales > 0 ? mProfit / mSales : 0;
    const mOrders = mRows.reduce((a, b) => a + b.orders, 0);

    monthlyChartCache.cats.push(mLabel);
    monthlyChartCache.target.push(mTarget);
    monthlyChartCache.sales.push(mSales);
    monthlyChartCache.cost.push(mCost);
    monthlyChartCache.profit.push(mProfit);

    sh.setCell(r, 2, { v: mLabel, s: 23 });
    sh.setCell(r, 3, { f: `'5_목표매출_및_공통비용'!E${6 + m}`, v: mTarget, s: 25 });
    sh.setCell(r, 4, {
      f: `SUMIFS('2_일일매출_입력'!$F$6:$F$505,'2_일일매출_입력'!$B$6:$B$505,">="&DATE($B$4,${m},1),'2_일일매출_입력'!$B$6:$B$505,"<="&EOMONTH(DATE($B$4,${m},1),0))`,
      v: mSales,
      s: 25
    });
    sh.setCell(r, 5, {
      f: `SUMIFS('2_일일매출_입력'!$O$6:$O$505,'2_일일매출_입력'!$B$6:$B$505,">="&DATE($B$4,${m},1),'2_일일매출_입력'!$B$6:$B$505,"<="&EOMONTH(DATE($B$4,${m},1),0))`,
      v: mCost,
      s: 25
    });
    sh.setCell(r, 6, { f: `D${r}-E${r}`, v: mProfit, s: 12 });
    sh.setCell(r, 7, { f: `IFERROR(F${r}/D${r},0)`, v: mMargin, s: 26 });
    sh.setCell(r, 8, {
      f: `SUMIFS('2_일일매출_입력'!$E$6:$E$505,'2_일일매출_입력'!$B$6:$B$505,">="&DATE($B$4,${m},1),'2_일일매출_입력'!$B$6:$B$505,"<="&EOMONTH(DATE($B$4,${m},1),0))`,
      v: mOrders,
      s: 25
    });
  }

  // 43행: 월별 합계
  const totTarget = monthlyChartCache.target.reduce((a, b) => a + b, 0);
  const totSales = monthlyChartCache.sales.reduce((a, b) => a + b, 0);
  const totCost = monthlyChartCache.cost.reduce((a, b) => a + b, 0);
  const totProfit = totSales - totCost;
  sh.setCell(43, 2, { v: '연간 합계', s: 16 });
  sh.setCell(43, 3, { f: 'SUM(C31:C42)', v: totTarget, s: 17 });
  sh.setCell(43, 4, { f: 'SUM(D31:D42)', v: totSales, s: 17 });
  sh.setCell(43, 5, { f: 'SUM(E31:E42)', v: totCost, s: 17 });
  sh.setCell(43, 6, { f: 'D43-E43', v: totProfit, s: 17 });
  sh.setCell(43, 7, { f: 'IFERROR(F43/D43,0)', v: totSales > 0 ? totProfit / totSales : 0, s: 18 });
  sh.setCell(43, 8, { f: 'SUM(H31:H42)', v: enrichedSales.reduce((a, b) => a + b.orders, 0), s: 17 });

  // 우측 표 헤더 (J30:O30): 판매채널별 매출 비중 및 손익 요약
  const chHeaders = [
    [10, '판매채널'],
    [11, '주문건수'],
    [12, '매출액 (원)'],
    [13, '매출 비중 (%)'],
    [14, '순이익 (원)'],
    [15, '순이익률 (%)']
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
    const cOrders = chRows.reduce((a, b) => a + b.orders, 0);
    const cSales = chRows.reduce((a, b) => a + b.sales, 0);
    const cProfit = chRows.reduce((a, b) => a + b.netProfit, 0);
    const cShare = periodActualSales > 0 ? cSales / periodActualSales : 0;
    const cMargin = cSales > 0 ? cProfit / cSales : 0;

    channelChartCache.cats.push(chName);
    channelChartCache.sales.push(cSales);
    channelChartCache.profits.push(cProfit);

    sh.setCell(r, 10, { f: `'6_사용설명서_및_기준관리'!B${6 + idx}`, v: chName, s: 23 });
    sh.setCell(r, 11, {
      f: `SUMIFS('2_일일매출_입력'!$E$6:$E$505,'2_일일매출_입력'!$D$6:$D$505,J${r},'2_일일매출_입력'!$B$6:$B$505,">="&$F$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$G$4)`,
      v: cOrders,
      s: 25
    });
    sh.setCell(r, 12, {
      f: `SUMIFS('2_일일매출_입력'!$F$6:$F$505,'2_일일매출_입력'!$D$6:$D$505,J${r},'2_일일매출_입력'!$B$6:$B$505,">="&$F$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$G$4)`,
      v: cSales,
      s: 25
    });
    sh.setCell(r, 13, { f: `IFERROR(L${r}/$L$38,0)`, v: cShare, s: 14 });
    sh.setCell(r, 14, {
      f: `SUMIFS('2_일일매출_입력'!$P$6:$P$505,'2_일일매출_입력'!$D$6:$D$505,J${r},'2_일일매출_입력'!$B$6:$B$505,">="&$F$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$G$4)`,
      v: cProfit,
      s: 12
    });
    sh.setCell(r, 15, { f: `IFERROR(N${r}/L${r},0)`, v: cMargin, s: 26 });
  });

  // 38행: 판매채널 합계
  sh.setCell(38, 10, { v: '합계', s: 16 });
  sh.setCell(38, 11, { f: 'SUM(K31:K37)', v: periodOrders, s: 17 });
  sh.setCell(38, 12, { f: 'SUM(L31:L37)', v: periodActualSales, s: 17 });
  sh.setCell(38, 13, { f: 'IFERROR(SUM(M31:M37),0)', v: 1, s: 18 });
  sh.setCell(38, 14, { f: 'SUM(N31:N37)', v: periodNetProfit, s: 17 });
  sh.setCell(38, 15, { f: 'IFERROR(N38/L38,0)', v: periodMargin, s: 18 });

  // 목표 vs 실제 비교 차트용 요약 테이블 (J40:M41)
  sh.setCell(40, 10, { v: '구분', s: 4 });
  sh.setCell(40, 11, { v: '목표매출', s: 4 });
  sh.setCell(40, 12, { v: '실제매출', s: 4 });
  sh.setCell(40, 13, { v: '순이익(공통비전)', s: 4 });
  sh.setCell(40, 14, { v: '최종순이익(공통비후)', s: 4 });

  sh.setCell(41, 10, { v: '선택기간 실적', s: 23 });
  sh.setCell(41, 11, { f: 'B7', v: periodTarget, s: 25 });
  sh.setCell(41, 12, { f: 'D7', v: periodActualSales, s: 25 });
  sh.setCell(41, 13, { f: 'L7', v: periodNetProfit, s: 25 });
  sh.setCell(41, 14, { f: 'L9', v: periodFinalProfit, s: 25 });

  // 대시보드 데이터 유효성 검사 & 조건부 서식
  sh.addValidation({
    type: 'list',
    sqref: 'C4',
    formula1: '"전체,1,2,3,4,5,6,7,8,9,10,11,12"',
    promptTitle: '조회 월 선택',
    prompt: '1~12월을 선택하거나 연간 전체를 보려면 [전체]를 선택하세요.'
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
// 시트 2: 2_일일매출_입력
// ============================================================================
function buildSheet2() {
  const sh = new SheetBuilder('2_일일매출_입력', {
    freezeRow: 5,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 14 }, // 매출일자
      { min: 3, max: 3, width: 18 }, // 거래처
      { min: 4, max: 4, width: 15 }, // 판매채널
      { min: 5, max: 5, width: 11 }, // 주문건수
      { min: 6, max: 8, width: 16 }, // 매출액, 원가, 추가비용
      { min: 9, max: 9, width: 34 }, // 비고
      { min: 10, max: 10, width: 16 }, // 비용설정 상태
      { min: 11, max: 11, width: 14 }, // 적용 수수료율
      { min: 12, max: 16, width: 16 }, // 수수료, 건당비용, 건수비용합계, 총비용, 순이익
      { min: 17, max: 17, width: 13 }  // 순이익률
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '📝 일일 매출 및 주문건수 입력 (날짜 × 거래처 × 판매채널별 일일 집계)', s: 1 });
  sh.merge('B2:I2');

  sh.setRowHeight(3, 38);
  sh.setCell(3, 2, {
    v: '💡 [입력 안내] ① 노란색 열(B~I열)만 직접 입력하세요. 파란색 열(J~Q열)은 해당 날짜의 판매채널 비용 설정을 찾아 자동 계산됩니다.\n② 일일 매출액은 할인 등을 반영한 실결제 금액으로 입력하세요(별도 할인 재차감 없음). 기본 배송비·포장비는 자동 계산되므로 [추가 비용]에 중복 입력하지 마세요.\n③ 같은 날짜·채널에 거래처가 여러 곳이면 거래처별로 행을 나누어 입력하고, 주문건수와 매출액이 중복 집계되지 않도록 주의하세요.',
    s: 22
  });
  sh.merge('B3:Q3');
  for (let c = 3; c <= 17; c++) sh.setCell(3, c, { v: '', s: 22 });

  // 4행: 전체 누적 합계 요약 바
  sh.setRowHeight(4, 24);
  const totOrdersAll = enrichedSales.reduce((a, b) => a + b.orders, 0);
  const totSalesAll = enrichedSales.reduce((a, b) => a + b.sales, 0);
  const totCogsAll = enrichedSales.reduce((a, b) => a + b.cogs, 0);
  const totExtraAll = enrichedSales.reduce((a, b) => a + b.extraCost, 0);
  const totFeeAll = enrichedSales.reduce((a, b) => a + b.feeAmt, 0);
  const totOrderCostAll = enrichedSales.reduce((a, b) => a + b.orderCost, 0);
  const totCostAll = enrichedSales.reduce((a, b) => a + b.totalCost, 0);
  const totProfitAll = totSalesAll - totCostAll;

  sh.setCell(4, 2, { v: '전체 누적 합계', s: 16 });
  sh.setCell(4, 3, { f: 'COUNTA(B6:B505)&"건 입력됨"', v: `${enrichedSales.length}건 입력됨`, s: 16 });
  sh.setCell(4, 4, { v: '-', s: 16 });
  sh.setCell(4, 5, { f: 'SUM(E6:E505)', v: totOrdersAll, s: 17 });
  sh.setCell(4, 6, { f: 'SUM(F6:F505)', v: totSalesAll, s: 17 });
  sh.setCell(4, 7, { f: 'SUM(G6:G505)', v: totCogsAll, s: 17 });
  sh.setCell(4, 8, { f: 'SUM(H6:H505)', v: totExtraAll, s: 17 });
  sh.setCell(4, 9, { v: '← 직접 입력 (연노랑) | 자동 계산 (연파랑) →', s: 16 });
  sh.setCell(4, 10, { f: 'COUNTIF(J6:J505,"비용 설정 필요")&"건 누락"', v: '0건 누락', s: 16 });
  sh.setCell(4, 11, { v: '-', s: 16 });
  sh.setCell(4, 12, { f: 'SUMIF(L6:L505,">=0")', v: totFeeAll, s: 17 });
  sh.setCell(4, 13, { v: '-', s: 16 });
  sh.setCell(4, 14, { f: 'SUMIF(N6:N505,">=0")', v: totOrderCostAll, s: 17 });
  sh.setCell(4, 15, { f: 'SUMIF(O6:O505,">=0")', v: totCostAll, s: 17 });
  sh.setCell(4, 16, { f: 'F4-O4', v: totProfitAll, s: 17 });
  sh.setCell(4, 17, { f: 'IFERROR(P4/F4,0)', v: totSalesAll > 0 ? totProfitAll / totSalesAll : 0, s: 18 });

  // 5행: 테이블 헤더
  sh.setRowHeight(5, 28);
  const inputCols = [
    [2, '매출일자\n(필수입력)'],
    [3, '거래처\n(드롭다운)'],
    [4, '판매채널\n(드롭다운)'],
    [5, '주문건수\n(0 이상 정수)'],
    [6, '일일 매출액\n(할인반영 실매출)'],
    [7, '일일 상품 총원가\n(직접입력)'],
    [8, '추가 비용\n(특이발생분만)'],
    [9, '비고\n(메모/예시구분)']
  ];
  for (const [c, t] of inputCols) sh.setCell(5, c, { v: t, s: 3 });

  const calcCols = [
    [10, '비용설정 상태\n(자동검증)'],
    [11, '적용 수수료율\n(자동조회)'],
    [12, '매출액 기준 수수료\n(자동계산)'],
    [13, '주문 1건당 비용\n(자동조회)'],
    [14, '주문건수 비용 합계\n(자동계산)'],
    [15, '총비용\n(자동계산)'],
    [16, '순이익\n(자동계산)'],
    [17, '순이익률\n(자동계산)']
  ];
  for (const [c, t] of calcCols) sh.setCell(5, c, { v: t, s: 4 });

  // 6행 ~ 305행: 예시 데이터 24건 + 향후 추가 입력을 위한 빈 수식 행 300행 사전 세팅
  const maxPreformattedRows = 300;
  for (let i = 0; i < maxPreformattedRows; i++) {
    const r = 6 + i;
    const item = enrichedSales[i];

    if (item) {
      sh.setCell(r, 2, { v: dateToSerial(item.date), s: 7 });
      sh.setCell(r, 3, { v: item.client, s: 5 });
      sh.setCell(r, 4, { v: item.channel, s: 5 });
      sh.setCell(r, 5, { v: item.orders, s: 8 });
      sh.setCell(r, 6, { v: item.sales, s: 8 });
      sh.setCell(r, 7, { v: item.cogs, s: 8 });
      sh.setCell(r, 8, { v: item.extraCost, s: 8 });
      sh.setCell(r, 9, { v: item.note, s: 6 });
    } else {
      sh.setCell(r, 2, { v: '', s: 7 });
      sh.setCell(r, 3, { v: '', s: 5 });
      sh.setCell(r, 4, { v: '', s: 5 });
      sh.setCell(r, 5, { v: '', s: 8 });
      sh.setCell(r, 6, { v: '', s: 8 });
      sh.setCell(r, 7, { v: '', s: 8 });
      sh.setCell(r, 8, { v: '', s: 8 });
      sh.setCell(r, 9, { v: '', s: 6 });
    }

    // 자동 계산 수식 (J~Q열)
    const fStatus = `IF(AND(B${r}="",C${r}="",D${r}=""),"",IF(OR(B${r}="",D${r}=""),"비용 설정 필요",IF(COUNTIFS('3_채널별_비용설정'!$C$6:$C$55,D${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$L$6:$L$55,">="&B${r})=0,"비용 설정 필요",IF(COUNTIFS('3_채널별_비용설정'!$C$6:$C$55,D${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$L$6:$L$55,">="&B${r})>1,"기간 중복 확인","정상"))))`;
    const fFeeRate = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",SUMIFS('3_채널별_비용설정'!$F$6:$F$55,'3_채널별_비용설정'!$C$6:$C$55,D${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$L$6:$L$55,">="&B${r})))`;
    const fFeeAmt = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",ROUND(F${r}*K${r},0)))`;
    const fUnitCost = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",SUMIFS('3_채널별_비용설정'!$J$6:$J$55,'3_채널별_비용설정'!$C$6:$C$55,D${r},'3_채널별_비용설정'!$D$6:$D$55,"<="&B${r},'3_채널별_비용설정'!$L$6:$L$55,">="&B${r})))`;
    const fOrderCost = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",E${r}*M${r}))`;
    const fTotalCost = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",G${r}+L${r}+N${r}+H${r}))`;
    const fNetProfit = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",F${r}-O${r}))`;
    const fMargin = `IF(J${r}="","",IF(J${r}<>"정상","비용 설정 필요",IFERROR(P${r}/F${r},0)))`;

    sh.setCell(r, 10, { f: fStatus, v: item ? item.status : '', s: 10 });
    sh.setCell(r, 11, { f: fFeeRate, v: item ? item.feeRate : '', s: 15 });
    sh.setCell(r, 12, { f: fFeeAmt, v: item ? item.feeAmt : '', s: 12 });
    sh.setCell(r, 13, { f: fUnitCost, v: item ? item.unitCost : '', s: 12 });
    sh.setCell(r, 14, { f: fOrderCost, v: item ? item.orderCost : '', s: 12 });
    sh.setCell(r, 15, { f: fTotalCost, v: item ? item.totalCost : '', s: 12 });
    sh.setCell(r, 16, { f: fNetProfit, v: item ? item.netProfit : '', s: 13 });
    sh.setCell(r, 17, { f: fMargin, v: item ? item.margin : '', s: 14 });
  }

  // 데이터 유효성 검사: 거래처 드롭다운, 판매채널 드롭다운, 주문건수 0 이상 정수
  sh.addValidation({
    type: 'list',
    sqref: 'C6:C505',
    formula1: "'6_사용설명서_및_기준관리'!$D$6:$D$35",
    promptTitle: '거래처 선택',
    prompt: '목록에서 거래처를 선택하세요. 새 거래처는 [6_사용설명서_및_기준관리] 시트에 추가하면 자동 반영됩니다.'
  });
  sh.addValidation({
    type: 'list',
    sqref: 'D6:D505',
    formula1: "'6_사용설명서_및_기준관리'!$B$6:$B$25",
    promptTitle: '판매채널 선택',
    prompt: '목록에서 판매채널을 선택하세요.'
  });
  sh.addValidation({
    type: 'whole',
    operator: 'greaterThanOrEqual',
    sqref: 'E6:E505',
    formula1: '0',
    promptTitle: '주문건수 입력 (0 이상 정수)',
    prompt: '하루 총 주문건수를 0 이상의 정수로 입력하세요.',
    errorTitle: '입력 오류 (주문건수)',
    error: '주문건수는 0 이상의 정수(자연수 또는 0)만 입력할 수 있습니다.'
  });

  // 조건부 서식: 순이익 적자(<0) 빨간색 강조, '비용 설정 필요' 경고 강조
  sh.addCondFormat({
    sqref: 'P6:P505',
    rules: [{ type: 'cellIs', operator: 'lessThan', formula: '0', dxfId: 0 }]
  });
  sh.addCondFormat({
    sqref: 'J6:Q505',
    rules: [{ type: 'containsText', text: '비용 설정 필요', topLeftCell: 'J6', dxfId: 1 }]
  });

  return sh;
}

// ============================================================================
// 시트 3: 3_채널별_비용설정
// ============================================================================
function buildSheet3() {
  const sh = new SheetBuilder('3_채널별_비용설정', {
    freezeRow: 5,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 7 },  // No
      { min: 3, max: 3, width: 16 }, // 판매채널
      { min: 4, max: 5, width: 15 }, // 시작일, 종료일
      { min: 6, max: 6, width: 16 }, // 수수료율
      { min: 7, max: 10, width: 16 }, // 배송비, 포장비, 기타비용, 건당합계
      { min: 11, max: 11, width: 18 }, // 기간 중복 검증
      { min: 12, max: 12, width: 15 }, // 유효종료일(자동기준)
      { min: 13, max: 13, width: 42 }  // 비고
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '⚙️ 판매채널·기간별 수수료 및 건당 비용 설정 시트', s: 1 });
  sh.merge('B2:M2');

  sh.setRowHeight(3, 38);
  sh.setCell(3, 2, {
    v: '💡 [기간별 이력 관리 규칙] ① 비용이나 수수료가 바뀌면 기존 행을 덮어쓰지 말고, 기존 행의 [적용 종료일]을 입력한 뒤 아래 새 행에 [새로운 적용 시작일]로 추가하세요.\n② [적용 종료일]을 비워두면 시작일부터 기한 제한 없이 계속 적용됩니다. 과거 매출에는 당시 적용되던 비용이 그대로 유지됩니다.\n③ 여기서 설정한 배송비·포장비는 일일 매출 시트에서 주문건수에 곱해 자동 계산되므로, 일일 매출 시트의 [추가 비용]에 중복 입력하지 마세요.',
    s: 22
  });
  sh.merge('B3:M3');
  for (let c = 3; c <= 13; c++) sh.setCell(3, c, { v: '', s: 22 });

  sh.setRowHeight(5, 28);
  const headers = [
    [2, 'No', 4],
    [3, '판매채널\n(드롭다운)', 3],
    [4, '적용 시작일\n(포함 >=)', 3],
    [5, '적용 종료일\n(비워두면 계속)', 3],
    [6, '매출액 기준\n수수료율 (%)', 3],
    [7, '주문 1건당\n배송비 (원)', 3],
    [8, '주문 1건당\n포장비 (원)', 3],
    [9, '주문 1건당\n기타 비용 (원)', 3],
    [10, '주문 1건당\n비용 합계 (자동)', 4],
    [11, '적용 기간\n중복/오류 검증', 4],
    [12, '종료 기준일\n(수식 자동처리)', 4],
    [13, '비고 (변경 사유 및 채널별 수수료 특이사항 메모)', 3]
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
      const unitSum = item.shipFee + item.packFee + item.etcFee;

      sh.setCell(r, 3, { v: item.channel, s: 5 });
      sh.setCell(r, 4, { v: startSer, s: 7 });
      sh.setCell(r, 5, { v: endSer, s: 7 });
      sh.setCell(r, 6, { v: item.feeRate, s: 9 });
      sh.setCell(r, 7, { v: item.shipFee, s: 8 });
      sh.setCell(r, 8, { v: item.packFee, s: 8 });
      sh.setCell(r, 9, { v: item.etcFee, s: 8 });
      sh.setCell(r, 10, { f: `IF(C${r}="","",SUM(G${r}:I${r}))`, v: unitSum, s: 13 });
      sh.setCell(r, 11, {
        f: `IF(C${r}="","",IF(D${r}>L${r},"⚠️ 시작일>종료일 오류",IF(COUNTIFS($C$6:$C$55,C${r},$D$6:$D$55,"<="&L${r},$L$6:$L$55,">="&D${r})>1,"⚠️ 기간 중복","정상")))`,
        v: '정상',
        s: 10
      });
      sh.setCell(r, 12, { f: `IF(C${r}="","",IF(E${r}="",DATE(2099,12,31),E${r}))`, v: effEndSer, s: 11 });
      sh.setCell(r, 13, { v: item.note, s: 6 });
    } else {
      sh.setCell(r, 3, { v: '', s: 5 });
      sh.setCell(r, 4, { v: '', s: 7 });
      sh.setCell(r, 5, { v: '', s: 7 });
      sh.setCell(r, 6, { v: '', s: 9 });
      sh.setCell(r, 7, { v: '', s: 8 });
      sh.setCell(r, 8, { v: '', s: 8 });
      sh.setCell(r, 9, { v: '', s: 8 });
      sh.setCell(r, 10, { f: `IF(C${r}="","",SUM(G${r}:I${r}))`, v: '', s: 13 });
      sh.setCell(r, 11, {
        f: `IF(C${r}="","",IF(D${r}>L${r},"⚠️ 시작일>종료일 오류",IF(COUNTIFS($C$6:$C$55,C${r},$D$6:$D$55,"<="&L${r},$L$6:$L$55,">="&D${r})>1,"⚠️ 기간 중복","정상")))`,
        v: '',
        s: 10
      });
      sh.setCell(r, 12, { f: `IF(C${r}="","",IF(E${r}="",DATE(2099,12,31),E${r}))`, v: '', s: 11 });
      sh.setCell(r, 13, { v: '', s: 6 });
    }
  }

  sh.addValidation({
    type: 'list',
    sqref: 'C6:C55',
    formula1: "'6_사용설명서_및_기준관리'!$B$6:$B$25",
    promptTitle: '판매채널 선택',
    prompt: '비용을 설정할 판매채널을 선택하세요.'
  });

  sh.addCondFormat({
    sqref: 'K6:K55',
    rules: [{ type: 'containsText', text: '⚠️', topLeftCell: 'K6', dxfId: 1 }]
  });

  return sh;
}

// ============================================================================
// 시트 4: 4_거래처_채널별_손익분석
// ============================================================================
function buildSheet4() {
  const sh = new SheetBuilder('4_거래처_채널별_손익분석', {
    freezeRow: 6,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 7 },  // 순위
      { min: 3, max: 3, width: 18 }, // 거래처/채널명
      { min: 4, max: 4, width: 11 }, // 주문건수
      { min: 5, max: 5, width: 15 }, // 매출액
      { min: 6, max: 6, width: 12 }, // 매출비중
      { min: 7, max: 12, width: 15 }, // 원가, 수수료, 건수비용, 추가비용, 총비용, 순이익(전)
      { min: 13, max: 13, width: 13 }, // 순이익률(전)
      { min: 14, max: 15, width: 15 }, // 공통비용배분, 최종순이익(후)
      { min: 16, max: 16, width: 13 }, // 최종순이익률(후)
      { min: 17, max: 17, width: 3 },  // 구분 여백
      { min: 18, max: 30, width: 13 }  // 우측 자동정렬 원본 계산 영역
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '📊 거래처별 · 판매채널별 손익 상세 분석 (선택 기간 연동 및 자동 정렬)', s: 1 });
  sh.merge('B2:P2');

  // 3~4행 컨트롤 바
  sh.setRowHeight(3, 22);
  sh.setRowHeight(4, 25);
  sh.setCell(3, 2, { v: '연동 시작일', s: 4 });
  sh.setCell(3, 3, { v: '연동 종료일', s: 4 });
  sh.setCell(3, 4, { v: '표/그래프 정렬 기준 (드롭다운 선택)', s: 3 });
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
  sh.setCell(4, 10, { v: '💡 적자(<0) 항목은 붉은색으로 강조됩니다. 순이익률은 일별 평균이 아닌 [기간 순이익 합계 ÷ 매출액 합계]로 계산됩니다.', s: 22 });
  sh.merge('J4:P4');
  for (let c = 11; c <= 16; c++) sh.setCell(4, c, { v: '', s: 22 });

  // -------------------------------------------------------------------------
  // 섹션 1: 거래처별 손익 분석 표 (6행 ~ 15행) + 우측 원본 집계 영역 (R~AC열)
  // -------------------------------------------------------------------------
  sh.setRowHeight(6, 24);
  sh.setCell(6, 2, { v: '🏢 1. 거래처별 손익 분석표 (같은 거래처의 여러 판매채널 실적 자동 합산)', s: 2 });
  sh.merge('B6:P6');
  sh.setCell(6, 18, { v: '[자동정렬용 원본 집계 데이터 - 수정하지 마세요]', s: 19 });
  sh.merge('R6:AC6');

  sh.setRowHeight(7, 28);
  const analysisHeaders = [
    [2, '순위'],
    [3, '거래처명'],
    [4, '주문건수'],
    [5, '매출액 (원)'],
    [6, '매출 비중'],
    [7, '상품 총원가'],
    [8, '매출액 수수료'],
    [9, '주문건수 비용'],
    [10, '추가 비용'],
    [11, '총비용\n(공통비 제외)'],
    [12, '순이익\n(공통비 전)'],
    [13, '순이익률\n(공통비 전)'],
    [14, '공통비용\n배분액'],
    [15, '최종 순이익\n(공통비 후)'],
    [16, '최종 이익률\n(공통비 후)']
  ];
  for (const [c, t] of analysisHeaders) sh.setCell(7, c, { v: t, s: 4 });

  const rawHeaders = [
    [18, '정렬점수'],
    [19, '항목명'],
    [20, '주문건수'],
    [21, '매출액'],
    [22, '상품원가'],
    [23, '매출수수료'],
    [24, '건수비용'],
    [25, '추가비용'],
    [26, '총비용'],
    [27, '순이익']
  ];
  for (const [c, t] of rawHeaders) sh.setCell(7, c, { v: t, s: 19 });

  // 거래처별 사전 계산 및 정렬
  const clientRawStats = CLIENTS.map((clName, idx) => {
    const rows = periodSales.filter(x => x.client === clName);
    const orders = rows.reduce((a, b) => a + b.orders, 0);
    const sales = rows.reduce((a, b) => a + b.sales, 0);
    const cogs = rows.reduce((a, b) => a + b.cogs, 0);
    const fee = rows.reduce((a, b) => a + b.feeAmt, 0);
    const orderCost = rows.reduce((a, b) => a + b.orderCost, 0);
    const extra = rows.reduce((a, b) => a + b.extraCost, 0);
    const totalCost = cogs + fee + orderCost + extra;
    const netProfit = sales - totalCost;
    const sortKey = netProfit + (100 - (8 + idx)) * 0.00001;
    return { clName, orders, sales, cogs, fee, orderCost, extra, totalCost, netProfit, sortKey, idx };
  });

  const sortedClients = [...clientRawStats].sort((a, b) => b.sortKey - a.sortKey);

  // 우측 원본 데이터 (R8:AA15) 및 좌측 정렬된 테이블 (B8:P15) 세팅
  CLIENTS.forEach((_, idx) => {
    const r = 8 + idx;
    const raw = clientRawStats[idx];

    // 우측 원본 영역 (R~AA열: 18~27열)
    sh.setCell(r, 18, {
      f: `IF($D$4="매출액 높은순",U${r},IF($D$4="순이익 높은순",AA${r},-ROW()))+(100-ROW())*0.00001`,
      v: raw.sortKey,
      s: 25
    });
    sh.setCell(r, 19, { f: `'6_사용설명서_및_기준관리'!D${6 + idx}`, v: raw.clName, s: 23 });
    sh.setCell(r, 20, {
      f: `SUMIFS('2_일일매출_입력'!$E$6:$E$505,'2_일일매출_입력'!$C$6:$C$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.orders,
      s: 25
    });
    sh.setCell(r, 21, {
      f: `SUMIFS('2_일일매출_입력'!$F$6:$F$505,'2_일일매출_입력'!$C$6:$C$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.sales,
      s: 25
    });
    sh.setCell(r, 22, {
      f: `SUMIFS('2_일일매출_입력'!$G$6:$G$505,'2_일일매출_입력'!$C$6:$C$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.cogs,
      s: 25
    });
    sh.setCell(r, 23, {
      f: `SUMIFS('2_일일매출_입력'!$L$6:$L$505,'2_일일매출_입력'!$C$6:$C$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.fee,
      s: 25
    });
    sh.setCell(r, 24, {
      f: `SUMIFS('2_일일매출_입력'!$N$6:$N$505,'2_일일매출_입력'!$C$6:$C$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.orderCost,
      s: 25
    });
    sh.setCell(r, 25, {
      f: `SUMIFS('2_일일매출_입력'!$H$6:$H$505,'2_일일매출_입력'!$C$6:$C$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.extra,
      s: 25
    });
    sh.setCell(r, 26, { f: `SUM(V${r}:Y${r})`, v: raw.totalCost, s: 25 });
    sh.setCell(r, 27, { f: `U${r}-Z${r}`, v: raw.netProfit, s: 25 });

    // 좌측 정렬된 메인 표 (B8:P15)
    const st = sortedClients[idx];
    const share = periodActualSales > 0 ? st.sales / periodActualSales : 0;
    const margin = st.sales > 0 ? st.netProfit / st.sales : 0;
    const allocCommon = Math.round(periodCommonCost * share);
    const finalProfit = st.netProfit - allocCommon;
    const finalMargin = st.sales > 0 ? finalProfit / st.sales : 0;

    sh.setCell(r, 2, { v: idx + 1, s: 10 });
    sh.setCell(r, 3, { f: `INDEX($S$8:$S$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.clName, s: 23 });
    sh.setCell(r, 4, { f: `INDEX($T$8:$T$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.orders, s: 25 });
    sh.setCell(r, 5, { f: `INDEX($U$8:$U$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.sales, s: 25 });
    sh.setCell(r, 6, { f: `IFERROR(E${r}/$E$16,0)`, v: share, s: 14 });
    sh.setCell(r, 7, { f: `INDEX($V$8:$V$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.cogs, s: 25 });
    sh.setCell(r, 8, { f: `INDEX($W$8:$W$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.fee, s: 25 });
    sh.setCell(r, 9, { f: `INDEX($X$8:$X$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.orderCost, s: 25 });
    sh.setCell(r, 10, { f: `INDEX($Y$8:$Y$15,MATCH(LARGE($R$8:$R$15,B${r}),$R$8:$R$15,0))`, v: st.extra, s: 25 });
    sh.setCell(r, 11, { f: `SUM(G${r}:J${r})`, v: st.totalCost, s: 12 });
    sh.setCell(r, 12, { f: `E${r}-K${r}`, v: st.netProfit, s: 13 });
    sh.setCell(r, 13, { f: `IFERROR(L${r}/E${r},0)`, v: margin, s: 14 });
    sh.setCell(r, 14, {
      f: `IF($G$4="매출액 비중 배분",ROUND($I$4*F${r},0),IF($G$4="주문건수 비중 배분",ROUND($I$4*IFERROR(D${r}/$D$16,0),0),0))`,
      v: allocCommon,
      s: 12
    });
    sh.setCell(r, 15, { f: `L${r}-N${r}`, v: finalProfit, s: 12 });
    sh.setCell(r, 16, { f: `IFERROR(O${r}/E${r},0)`, v: finalMargin, s: 14 });
  });

  // 16행: 거래처별 합계
  sh.setCell(16, 2, { v: '합계', s: 16 });
  sh.setCell(16, 3, { v: `전체 ${CLIENTS.length}개 거래처`, s: 16 });
  sh.setCell(16, 4, { f: 'SUM(D8:D15)', v: periodOrders, s: 17 });
  sh.setCell(16, 5, { f: 'SUM(E8:E15)', v: periodActualSales, s: 17 });
  sh.setCell(16, 6, { f: 'IFERROR(SUM(F8:F15),0)', v: 1, s: 18 });
  sh.setCell(16, 7, { f: 'SUM(G8:G15)', v: periodCogs, s: 17 });
  sh.setCell(16, 8, { f: 'SUM(H8:H15)', v: periodFee, s: 17 });
  sh.setCell(16, 9, { f: 'SUM(I8:I15)', v: periodOrderCost, s: 17 });
  sh.setCell(16, 10, { f: 'SUM(J8:J15)', v: periodExtraCost, s: 17 });
  sh.setCell(16, 11, { f: 'SUM(K8:K15)', v: periodTotalCost, s: 17 });
  sh.setCell(16, 12, { f: 'E16-K16', v: periodNetProfit, s: 17 });
  sh.setCell(16, 13, { f: 'IFERROR(L16/E16,0)', v: periodMargin, s: 18 });
  sh.setCell(16, 14, { f: 'SUM(N8:N15)', v: periodCommonCost, s: 17 });
  sh.setCell(16, 15, { f: 'L16-N16', v: periodFinalProfit, s: 17 });
  sh.setCell(16, 16, { f: 'IFERROR(O16/E16,0)', v: periodFinalMargin, s: 18 });

  // -------------------------------------------------------------------------
  // 섹션 2: 판매채널별 손익 분석 표 (18행 ~ 26행) + 우측 원본 집계 영역 (R20:AA26)
  // -------------------------------------------------------------------------
  sh.setRowHeight(18, 24);
  sh.setCell(18, 2, { v: '🛒 2. 판매채널별 손익 분석표 (동일 기간 채널별 매출·비용·순이익 비교)', s: 2 });
  sh.merge('B18:P18');

  sh.setRowHeight(19, 28);
  const chAnalysisHeaders = [...analysisHeaders];
  chAnalysisHeaders[1] = [3, '판매채널명'];
  for (const [c, t] of chAnalysisHeaders) sh.setCell(19, c, { v: t, s: 4 });
  for (const [c, t] of rawHeaders) sh.setCell(19, c, { v: t, s: 19 });

  const channelRawStats = CHANNELS.map((chName, idx) => {
    const rows = periodSales.filter(x => x.channel === chName);
    const orders = rows.reduce((a, b) => a + b.orders, 0);
    const sales = rows.reduce((a, b) => a + b.sales, 0);
    const cogs = rows.reduce((a, b) => a + b.cogs, 0);
    const fee = rows.reduce((a, b) => a + b.feeAmt, 0);
    const orderCost = rows.reduce((a, b) => a + b.orderCost, 0);
    const extra = rows.reduce((a, b) => a + b.extraCost, 0);
    const totalCost = cogs + fee + orderCost + extra;
    const netProfit = sales - totalCost;
    const sortKey = netProfit + (100 - (20 + idx)) * 0.00001;
    return { chName, orders, sales, cogs, fee, orderCost, extra, totalCost, netProfit, sortKey, idx };
  });

  const sortedChannels = [...channelRawStats].sort((a, b) => b.sortKey - a.sortKey);

  CHANNELS.forEach((_, idx) => {
    const r = 20 + idx;
    const raw = channelRawStats[idx];

    sh.setCell(r, 18, {
      f: `IF($D$4="매출액 높은순",U${r},IF($D$4="순이익 높은순",AA${r},-ROW()))+(100-ROW())*0.00001`,
      v: raw.sortKey,
      s: 25
    });
    sh.setCell(r, 19, { f: `'6_사용설명서_및_기준관리'!B${6 + idx}`, v: raw.chName, s: 23 });
    sh.setCell(r, 20, {
      f: `SUMIFS('2_일일매출_입력'!$E$6:$E$505,'2_일일매출_입력'!$D$6:$D$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.orders,
      s: 25
    });
    sh.setCell(r, 21, {
      f: `SUMIFS('2_일일매출_입력'!$F$6:$F$505,'2_일일매출_입력'!$D$6:$D$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.sales,
      s: 25
    });
    sh.setCell(r, 22, {
      f: `SUMIFS('2_일일매출_입력'!$G$6:$G$505,'2_일일매출_입력'!$D$6:$D$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.cogs,
      s: 25
    });
    sh.setCell(r, 23, {
      f: `SUMIFS('2_일일매출_입력'!$L$6:$L$505,'2_일일매출_입력'!$D$6:$D$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.fee,
      s: 25
    });
    sh.setCell(r, 24, {
      f: `SUMIFS('2_일일매출_입력'!$N$6:$N$505,'2_일일매출_입력'!$D$6:$D$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.orderCost,
      s: 25
    });
    sh.setCell(r, 25, {
      f: `SUMIFS('2_일일매출_입력'!$H$6:$H$505,'2_일일매출_입력'!$D$6:$D$505,S${r},'2_일일매출_입력'!$B$6:$B$505,">="&$B$4,'2_일일매출_입력'!$B$6:$B$505,"<="&$C$4)`,
      v: raw.extra,
      s: 25
    });
    sh.setCell(r, 26, { f: `SUM(V${r}:Y${r})`, v: raw.totalCost, s: 25 });
    sh.setCell(r, 27, { f: `U${r}-Z${r}`, v: raw.netProfit, s: 25 });

    // 좌측 정렬된 메인 표 (B20:P26)
    const st = sortedChannels[idx];
    const share = periodActualSales > 0 ? st.sales / periodActualSales : 0;
    const margin = st.sales > 0 ? st.netProfit / st.sales : 0;
    const allocCommon = Math.round(periodCommonCost * share);
    const finalProfit = st.netProfit - allocCommon;
    const finalMargin = st.sales > 0 ? finalProfit / st.sales : 0;

    sh.setCell(r, 2, { v: idx + 1, s: 10 });
    sh.setCell(r, 3, { f: `INDEX($S$20:$S$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.chName, s: 23 });
    sh.setCell(r, 4, { f: `INDEX($T$20:$T$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.orders, s: 25 });
    sh.setCell(r, 5, { f: `INDEX($U$20:$U$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.sales, s: 25 });
    sh.setCell(r, 6, { f: `IFERROR(E${r}/$E$27,0)`, v: share, s: 14 });
    sh.setCell(r, 7, { f: `INDEX($V$20:$V$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.cogs, s: 25 });
    sh.setCell(r, 8, { f: `INDEX($W$20:$W$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.fee, s: 25 });
    sh.setCell(r, 9, { f: `INDEX($X$20:$X$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.orderCost, s: 25 });
    sh.setCell(r, 10, { f: `INDEX($Y$20:$Y$26,MATCH(LARGE($R$20:$R$26,B${r}),$R$20:$R$26,0))`, v: st.extra, s: 25 });
    sh.setCell(r, 11, { f: `SUM(G${r}:J${r})`, v: st.totalCost, s: 12 });
    sh.setCell(r, 12, { f: `E${r}-K${r}`, v: st.netProfit, s: 13 });
    sh.setCell(r, 13, { f: `IFERROR(L${r}/E${r},0)`, v: margin, s: 14 });
    sh.setCell(r, 14, {
      f: `IF($G$4="매출액 비중 배분",ROUND($I$4*F${r},0),IF($G$4="주문건수 비중 배분",ROUND($I$4*IFERROR(D${r}/$D$27,0),0),0))`,
      v: allocCommon,
      s: 12
    });
    sh.setCell(r, 15, { f: `L${r}-N${r}`, v: finalProfit, s: 12 });
    sh.setCell(r, 16, { f: `IFERROR(O${r}/E${r},0)`, v: finalMargin, s: 14 });
  });

  // 27행: 판매채널별 합계
  sh.setCell(27, 2, { v: '합계', s: 16 });
  sh.setCell(27, 3, { v: `전체 ${CHANNELS.length}개 판매채널`, s: 16 });
  sh.setCell(27, 4, { f: 'SUM(D20:D26)', v: periodOrders, s: 17 });
  sh.setCell(27, 5, { f: 'SUM(E20:E26)', v: periodActualSales, s: 17 });
  sh.setCell(27, 6, { f: 'IFERROR(SUM(F20:F26),0)', v: 1, s: 18 });
  sh.setCell(27, 7, { f: 'SUM(G20:G26)', v: periodCogs, s: 17 });
  sh.setCell(27, 8, { f: 'SUM(H20:H26)', v: periodFee, s: 17 });
  sh.setCell(27, 9, { f: 'SUM(I20:I26)', v: periodOrderCost, s: 17 });
  sh.setCell(27, 10, { f: 'SUM(J20:J26)', v: periodExtraCost, s: 17 });
  sh.setCell(27, 11, { f: 'SUM(K20:K26)', v: periodTotalCost, s: 17 });
  sh.setCell(27, 12, { f: 'E27-K27', v: periodNetProfit, s: 17 });
  sh.setCell(27, 13, { f: 'IFERROR(L27/E27,0)', v: periodMargin, s: 18 });
  sh.setCell(27, 14, { f: 'SUM(N20:N26)', v: periodCommonCost, s: 17 });
  sh.setCell(27, 15, { f: 'L27-N27', v: periodFinalProfit, s: 17 });
  sh.setCell(27, 16, { f: 'IFERROR(O27/E27,0)', v: periodFinalMargin, s: 18 });

  // 29~46행: 비교 차트 2종 배치 영역
  sh.setCell(29, 2, { v: '📈 거래처별 순이익 비교 그래프 & 판매채널별 매출·순이익 비교 그래프', s: 2 });

  sh.addValidation({
    type: 'list',
    sqref: 'D4:F4',
    formula1: '"순이익 높은순,매출액 높은순,입력(기본)순"',
    promptTitle: '정렬 기준 선택',
    prompt: '순이익 높은순 또는 매출액 높은순을 선택하면 표와 그래프가 자동 재정렬됩니다.'
  });

  // 적자 발생 항목 조건부 서식 (연분홍 배경 + 빨간 굵은 글씨)
  sh.addCondFormat({
    sqref: 'L8:P16 L20:P27',
    rules: [{ type: 'cellIs', operator: 'lessThan', formula: '0', dxfId: 0 }]
  });

  sh.drawingRelId = 'rId1';

  return { sh, sortedClients, sortedChannels };
}

// ============================================================================
// 시트 5: 5_목표매출_및_공통비용
// ============================================================================
function buildSheet5() {
  const sh = new SheetBuilder('5_목표매출_및_공통비용', {
    freezeRow: 6,
    cols: [
      { min: 1, max: 1, width: 3 },
      { min: 2, max: 2, width: 9 },  // 월
      { min: 3, max: 4, width: 14 }, // 시작일, 종료일
      { min: 5, max: 7, width: 18 }, // 월 목표매출, 일할 목표, 조회기간 반영 목표
      { min: 8, max: 8, width: 4 },  // 여백
      { min: 9, max: 9, width: 14 }, // 공통비용 일자
      { min: 10, max: 10, width: 20 }, // 항목
      { min: 11, max: 11, width: 16 }, // 금액
      { min: 12, max: 13, width: 14 }, // 귀속연월, 조회기간 포함여부
      { min: 14, max: 14, width: 32 }  // 비고
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '🎯 연·월별 목표매출 설정 및 월별 공통비용(고정비) 관리 시트', s: 1 });
  sh.merge('B2:N2');

  sh.setRowHeight(3, 38);
  sh.setCell(3, 2, {
    v: '💡 [목표매출 일할 계산 기준] 대시보드에서 한 달 전체를 조회하면 해당 월 목표매출이 100% 반영되며, 임의의 조회 기간(시작일~종료일)을 선택하면 각 월의 목표매출을 그 달의 총 일수로 나눈 [일할 목표매출 × 조회기간 포함 일수]로 정확히 비례 합산됩니다.\n💡 [공통비용 배분 기준] 광고비·임대료 등 주문건수와 무관한 공통비용은 우측 표에 입력하며, 대시보드 전체 총비용에서는 단 한 번만 차감되어 중복 차감되지 않습니다.',
    s: 22
  });
  sh.merge('B3:N3');
  for (let c = 3; c <= 14; c++) sh.setCell(3, c, { v: '', s: 22 });

  sh.setRowHeight(4, 25);
  sh.setCell(4, 2, { v: '목표 기준연도', s: 3 });
  sh.setCell(4, 3, { v: defaultYear, s: 5 });
  sh.setCell(4, 5, { v: '📌 좌측: 월별 목표매출 입력 (연노랑)', s: 2 });
  sh.merge('E4:G4');

  sh.setCell(4, 9, { v: '공통비용 배분 기준 선택', s: 3 });
  sh.merge('I4:J4');
  sh.setCell(4, 10, { v: '', s: 3 });
  sh.setCell(4, 11, { v: '매출액 비중 배분', s: 5 });
  sh.merge('K4:L4');
  sh.setCell(4, 12, { v: '', s: 5 });

  // 6행: 좌·우 테이블 헤더
  sh.setRowHeight(6, 28);
  const leftHeaders = [
    [2, '월', 4],
    [3, '월 시작일', 4],
    [4, '월 종료일', 4],
    [5, '월별 목표매출 (원)\n(직접입력)', 3],
    [6, '일할 목표매출 (원/일)\n(자동계산)', 4],
    [7, '조회기간 반영 목표매출\n(대시보드 자동연동)', 4]
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

  // 좌측 1~12월 목표매출 세팅 (7행 ~ 18행)
  for (let m = 1; m <= 12; m++) {
    const r = 6 + m;
    const tObj = MONTHLY_TARGETS[m - 1];
    const startSer = dateToSerial(`2026-${String(m).padStart(2, '0')}-01`);
    const daysInMonth = new Date(Date.UTC(2026, m, 0)).getUTCDate();
    const endSer = dateToSerial(`2026-${String(m).padStart(2, '0')}-${String(daysInMonth).padStart(2, '0')}`);
    const dailyTarget = Math.round(tObj.target / daysInMonth);
    const appliedTarget = m === defaultMonth ? tObj.target : 0;

    sh.setCell(r, 2, { v: m, s: 10 });
    sh.setCell(r, 3, { f: `DATE($C$4,B${r},1)`, v: startSer, s: 11 });
    sh.setCell(r, 4, { f: `EOMONTH(C${r},0)`, v: endSer, s: 11 });
    sh.setCell(r, 5, { v: tObj.target, s: 8 });
    sh.setCell(r, 6, { f: `ROUND(E${r}/DAY(D${r}),0)`, v: dailyTarget, s: 12 });
    sh.setCell(r, 7, {
      f: `IF(OR('1_매출손익_대시보드'!$G$4<C${r},'1_매출손익_대시보드'!$F$4>D${r}),0,IF(AND('1_매출손익_대시보드'!$F$4<=C${r},'1_매출손익_대시보드'!$G$4>=D${r}),E${r},ROUND(E${r}*(MIN('1_매출손익_대시보드'!$G$4,D${r})-MAX('1_매출손익_대시보드'!$F$4,C${r})+1)/DAY(D${r}),0)))`,
      v: appliedTarget,
      s: 13
    });
  }

  // 19행: 목표매출 합계
  const annualTargetSum = MONTHLY_TARGETS.reduce((a, b) => a + b.target, 0);
  sh.setCell(19, 2, { v: '합계', s: 16 });
  sh.setCell(19, 3, { v: '연간 총 목표', s: 16 });
  sh.setCell(19, 4, { v: '-', s: 16 });
  sh.setCell(19, 5, { f: 'SUM(E7:E18)', v: annualTargetSum, s: 17 });
  sh.setCell(19, 6, { f: 'ROUND(E19/365,0)', v: Math.round(annualTargetSum / 365), s: 17 });
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

  // 26행: 공통비용 합계
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
    prompt: '거래처별·채널별 손익 분석 시 공통비용을 어떤 기준으로 배분할지 선택하세요.'
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
      { min: 2, max: 2, width: 20 }, // 판매채널 마스터
      { min: 3, max: 3, width: 4 },
      { min: 4, max: 4, width: 22 }, // 거래처 마스터
      { min: 5, max: 5, width: 4 },
      { min: 6, max: 6, width: 24 }, // 가이드 항목
      { min: 7, max: 7, width: 75 }  // 상세 설명
    ]
  });

  sh.setRowHeight(2, 28);
  sh.setCell(2, 2, { v: '📖 프로그램 사용 설명서 및 판매채널·거래처 기준 목록 관리', s: 1 });
  sh.merge('B2:G2');

  sh.setRowHeight(3, 32);
  sh.setCell(3, 2, {
    v: '💡 좌측의 [판매채널 목록]과 [거래처 목록]의 빈 노란색 칸에 새로운 채널이나 거래처 이름을 입력하시면, [일일 매출 입력] 시트의 드롭다운 선택 목록에 자동으로 추가됩니다.',
    s: 22
  });
  sh.merge('B3:G3');
  for (let c = 3; c <= 7; c++) sh.setCell(3, c, { v: '', s: 22 });

  sh.setRowHeight(5, 26);
  sh.setCell(5, 2, { v: '🛒 판매채널 등록 목록\n(기본 7종 + 추가 가능)', s: 3 });
  sh.setCell(5, 4, { v: '🏢 거래처 등록 목록\n(드롭다운 연동 + 추가 가능)', s: 3 });
  sh.setCell(5, 6, { v: '📌 구분 / 핵심 주제', s: 4 });
  sh.setCell(5, 7, { v: '📖 상세 사용 방법 및 계산 기준 안내 (비개발자 맞춤 가이드)', s: 4 });

  for (let i = 0; i < 20; i++) {
    const r = 6 + i;
    sh.setCell(r, 2, { v: CHANNELS[i] || '', s: 5 });
  }
  for (let i = 0; i < 30; i++) {
    const r = 6 + i;
    sh.setCell(r, 4, { v: CLIENTS[i] || '', s: 5 });
  }

  const guides = [
    ['1. 셀 색상 구분 규칙', '🟨 연노랑색 셀: 사용자가 직접 입력하거나 드롭다운으로 선택하는 셀입니다.\n🟦 연파랑/연회색 셀: 수식이 들어있어 자동 계산되는 셀입니다. 수식이 지워지지 않도록 직접 수정하지 마세요.\n🟥 연분홍 배경 + 빨간 글씨: 적자(순이익 < 0)가 발생했거나, 비용 설정 누락/기간 중복 오류가 있을 때 자동으로 켜집니다.'],
    ['2. 일일 매출 입력 방법', '주문 1건씩 일일이 입력할 필요 없이, 하루 단위로 [매출일자 × 거래처 × 판매채널]별 주문건수와 일일 매출액 합계만 입력하면 됩니다.\n일일 매출액은 쿠폰·할인 등을 모두 반영한 실제 결제 금액을 입력하며, 별도의 할인액을 다시 차감하지 않습니다.'],
    ['3. 별도 입력이 필요한 항목', '① [일일 상품 총원가]: 해당 날짜에 판매된 상품들의 원가 합계를 직접 입력합니다.\n② [추가 비용]: 기본 택배비·포장비 외에 당일 특별히 발생한 비용(예: 반품/교환 택배비, 아이스팩/드라이아이스 추가분 등)만 입력합니다.\n※ 주의: [3_채널별_비용설정] 시트에 등록한 건당 배송비와 포장비는 주문건수에 곱해 자동 계산되므로 [추가 비용] 칸에 중복 입력하지 마세요!'],
    ['4. 거래처 및 판매채널 구분 관리', '• 같은 거래처가 여러 판매채널(예: 스마트스토어, 고흥몰, 농라)을 이용하더라도 채널별로 행을 나누어 입력해 주세요. [4_거래처_채널별_손익분석] 시트에서 거래처별 전체 손익이 자동 합산됩니다.\n• 같은 날짜·같은 채널에 거래처가 여러 곳이라면 거래처별로 행을 나누어 입력하되, 주문건수와 매출액이 중복 입력되지 않도록 주의해 주세요.'],
    ['5. 판매채널 비용 이력 관리', '• 택배비나 수수료율이 인상되면 기존 설정 행의 숫자를 덮어쓰지 마세요!\n• 기존 설정 행의 [적용 종료일]에 변경 전날(예: 2026-04-15)을 적고, 바로 아래 새 행에 같은 채널명으로 [적용 시작일](예: 2026-04-16)과 새 단가를 등록하세요.\n• 이렇게 하면 과거 매출에는 과거 단가가 그대로 유지되고, 변경일 이후 매출에만 새 단가가 정확히 적용됩니다.'],
    ['6. 오류 및 누락 자동 감지', '• 매출일자에 해당하는 채널 비용 설정이 없으면 임의로 0원 처리하지 않고 [비용 설정 필요]라고 즉시 경고하며, 첫 번째 대시보드 상단에도 누락 건수가 표시됩니다.\n• 같은 채널의 적용 기간이 하루라도 겹치면 [⚠️ 기간 중복] 경고가 표시됩니다.'],
    ['7. 순이익률 및 계산 공식', '• 매출액 기준 수수료 = 일일 매출액 × 해당 날짜의 채널 수수료율\n• 주문 1건당 비용 = 해당 날짜의 건당 배송비 + 포장비 + 기타 건당 비용\n• 주문건수에 따른 비용 = 주문건수 × 주문 1건당 비용\n• 총비용 = 상품 총원가 + 매출액 기준 수수료 + 주문건수에 따른 비용 + 추가 비용\n• 순이익 = 일일 매출액 - 총비용  |  순이익률 = 순이익 ÷ 일일 매출액 × 100\n※ 월별·거래처별·채널별 순이익률은 일별 이익률의 단순 평균이 아니라 [기간 순이익 합계 ÷ 기간 매출액 합계]로 정확히 계산됩니다.'],
    ['8. 부가세 및 반품·취소 처리', '• 모든 금액(매출, 원가, 수수료, 배송·포장비, 공통비용)은 부가세(VAT) 포함 실결제/실지출 금액 기준으로 통일하여 입력합니다.\n• 발송 전 취소는 당일 주문건수와 매출액에서 제외하고 입력합니다.\n• 발송 후 반품/교환이 발생한 경우 매출액에서 환불 금액을 차감하고, 이미 지출된 반품 회수 택배비 등은 [추가 비용] 칸에 입력한 뒤 비고란에 사유를 적어주세요.'],
    ['9. 예시 데이터 안내', '현재 파일에는 수식과 그래프 작동을 바로 확인하실 수 있도록 비고란에 [예시]가 표시된 샘플 데이터가 들어 있습니다.\n실제 업무에 사용하실 때는 [2_일일매출_입력] 시트의 예시 행(B6:I29)을 지우시고 실제 매출 데이터를 입력해 사용하세요.']
  ];

  guides.forEach(([title, desc], idx) => {
    const r = 6 + idx * 2;
    sh.setRowHeight(r, 38);
    sh.setRowHeight(r + 1, 18);
    sh.setCell(r, 6, { v: title, s: 16 });
    sh.setCell(r + 1, 6, { v: '', s: 16 });
    sh.merge(`F${r}:F${r + 1}`);

    sh.setCell(r, 7, { v: desc, s: 24 });
    sh.setCell(r + 1, 7, { v: '', s: 24 });
    sh.merge(`G${r}:G${r + 1}`);
  });

  return sh;
}

// ============================================================================
// 전체 워크북 패키징 및 파일 저장
// ============================================================================
function generateAll() {
  const { sh: sheet1, monthlyChartCache, channelChartCache } = buildSheet1();
  const sheet2 = buildSheet2();
  const sheet3 = buildSheet3();
  const { sh: sheet4, sortedClients, sortedChannels } = buildSheet4();
  const sheet5 = buildSheet5();
  const sheet6 = buildSheet6();

  const sheets = [sheet1, sheet2, sheet3, sheet4, sheet5, sheet6];

  // 차트 1: 목표매출 vs 실제매출 비교 막대 차트 (시트1 좌측)
  const chart1Xml = buildBarChartXml({
    title: '목표매출 vs 실제매출 및 순이익 비교 (선택 기간)',
    categoriesFormula: "'1_매출손익_대시보드'!$K$40:$N$40",
    categoriesCache: ['목표매출', '실제매출', '순이익(공통비전)', '최종순이익(공통비후)'],
    seriesList: [
      {
        name: '선택기간 실적 (원)',
        nameFormula: "'1_매출손익_대시보드'!$J$41",
        color: '2563EB',
        valuesFormula: "'1_매출손익_대시보드'!$K$41:$N$41",
        valuesCache: [periodTarget, periodActualSales, periodNetProfit, periodFinalProfit]
      }
    ]
  });

  // 차트 2: 월별 매출·비용·순이익 추이 그래프 (시트1 중앙)
  const chart2Xml = buildBarChartXml({
    title: '월별 매출 · 총비용 · 순이익 추이 (연간)',
    categoriesFormula: "'1_매출손익_대시보드'!$B$31:$B$42",
    categoriesCache: monthlyChartCache.cats,
    showVal: false,
    seriesList: [
      {
        name: '실제매출',
        nameFormula: "'1_매출손익_대시보드'!$D$30",
        color: '2563EB',
        valuesFormula: "'1_매출손익_대시보드'!$D$31:$D$42",
        valuesCache: monthlyChartCache.sales
      },
      {
        name: '총비용(변동)',
        nameFormula: "'1_매출손익_대시보드'!$E$30",
        color: 'F59E0B',
        valuesFormula: "'1_매출손익_대시보드'!$E$31:$E$42",
        valuesCache: monthlyChartCache.cost
      },
      {
        name: '순이익(공통비전)',
        nameFormula: "'1_매출손익_대시보드'!$F$30",
        color: '10B981',
        valuesFormula: "'1_매출손익_대시보드'!$F$31:$F$42",
        valuesCache: monthlyChartCache.profit
      }
    ]
  });

  // 차트 3: 판매채널별 매출 비중 도넛 그래프 (금액 + 비율 동시 표시, 시트1 우측)
  const chart3Xml = buildDoughnutChartXml({
    title: '판매채널별 매출 비중 (금액 및 비율 %)',
    categoriesFormula: "'1_매출손익_대시보드'!$J$31:$J$37",
    categoriesCache: channelChartCache.cats,
    seriesNameFormula: "'1_매출손익_대시보드'!$L$30",
    seriesName: '매출액 (원)',
    valuesFormula: "'1_매출손익_대시보드'!$L$31:$L$37",
    valuesCache: channelChartCache.sales
  });

  // 차트 4: 거래처별 순이익 비교 막대 차트 (시트4 좌측)
  const chart4Xml = buildBarChartXml({
    title: '거래처별 순이익 비교 (공통비 반영 전 vs 후)',
    categoriesFormula: "'4_거래처_채널별_손익분석'!$C$8:$C$15",
    categoriesCache: sortedClients.map(c => c.clName),
    seriesList: [
      {
        name: '순이익(공통비 전)',
        nameFormula: "'4_거래처_채널별_손익분석'!$L$7",
        color: '2563EB',
        valuesFormula: "'4_거래처_채널별_손익분석'!$L$8:$L$15",
        valuesCache: sortedClients.map(c => c.netProfit)
      },
      {
        name: '최종 순이익(공통비 후)',
        nameFormula: "'4_거래처_채널별_손익분석'!$O$7",
        color: '10B981',
        valuesFormula: "'4_거래처_채널별_손익분석'!$O$8:$O$15",
        valuesCache: sortedClients.map(c => {
          const share = periodActualSales > 0 ? c.sales / periodActualSales : 0;
          return c.netProfit - Math.round(periodCommonCost * share);
        })
      }
    ]
  });

  // 차트 5: 판매채널별 매출·순이익 비교 막대 차트 (시트4 우측)
  const chart5Xml = buildBarChartXml({
    title: '판매채널별 매출액 및 순이익 비교',
    categoriesFormula: "'4_거래처_채널별_손익분석'!$C$20:$C$26",
    categoriesCache: sortedChannels.map(c => c.chName),
    seriesList: [
      {
        name: '매출액 (원)',
        nameFormula: "'4_거래처_채널별_손익분석'!$E$19",
        color: '3B82F6',
        valuesFormula: "'4_거래처_채널별_손익분석'!$E$20:$E$26",
        valuesCache: sortedChannels.map(c => c.sales)
      },
      {
        name: '순이익(공통비 전)',
        nameFormula: "'4_거래처_채널별_손익분석'!$L$19",
        color: '10B981',
        valuesFormula: "'4_거래처_채널별_손익분석'!$L$20:$L$26",
        valuesCache: sortedChannels.map(c => c.netProfit)
      }
    ]
  });

  // Drawing 1 (시트1의 차트 3개 배치: 10행~27행)
  const drawing1Xml = buildDrawingXml([
    { fromCol: 1, fromRow: 10, toCol: 5, toRow: 27, relId: 'rId1' },
    { fromCol: 5, fromRow: 10, toCol: 10, toRow: 27, relId: 'rId2' },
    { fromCol: 10, fromRow: 10, toCol: 15, toRow: 27, relId: 'rId3' }
  ]);

  // Drawing 2 (시트4의 차트 2개 배치: 30행~46행)
  const drawing2Xml = buildDrawingXml([
    { fromCol: 1, fromRow: 30, toCol: 8, toRow: 46, relId: 'rId1' },
    { fromCol: 8, fromRow: 30, toCol: 16, toRow: 46, relId: 'rId2' }
  ]);

  const files = [];

  // [Content_Types].xml
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

  // _rels/.rels
  files.push({
    name: '_rels/.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
  });

  // xl/workbook.xml (fullCalcOnLoad="1" 설정으로 파일 열 때 엑셀이 즉시 최신 값 재계산)
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

  // xl/_rels/workbook.xml.rels
  files.push({
    name: 'xl/_rels/workbook.xml.rels',
    data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheets.map((_, idx) => `<Relationship Id="rId${idx + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${idx + 1}.xml"/>`).join('\n  ')}
  <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  });

  // xl/styles.xml
  files.push({
    name: 'xl/styles.xml',
    data: buildStylesXml()
  });

  // Worksheets
  sheets.forEach((s, idx) => {
    files.push({
      name: `xl/worksheets/sheet${idx + 1}.xml`,
      data: s.toXml()
    });
  });

  // Worksheet relationships for drawings
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

  // Drawings & relationships
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

  // Charts
  files.push({ name: 'xl/charts/chart1.xml', data: chart1Xml });
  files.push({ name: 'xl/charts/chart2.xml', data: chart2Xml });
  files.push({ name: 'xl/charts/chart3.xml', data: chart3Xml });
  files.push({ name: 'xl/charts/chart4.xml', data: chart4Xml });
  files.push({ name: 'xl/charts/chart5.xml', data: chart5Xml });

  const zipBuf = createZipBuffer(files);
  const outPath = path.join(__dirname, '거래처별_매출순이익_관리프로그램.xlsx');
  fs.writeFileSync(outPath, zipBuf);
  console.log('Successfully generated:', outPath, 'Size:', zipBuf.length, 'bytes');

  // 검증 요약 출력
  console.log('--- [검증 결과 요약] ---');
  const beforeRow = enrichedSales.find(r => r.date === '2026-04-14' && r.channel === '스마트스토어');
  const afterRow = enrichedSales.find(r => r.date === '2026-04-17' && r.channel === '스마트스토어');
  console.log('1) 단가 변경일 전(2026-04-14 스마트스토어 20건, 매출 800,000원):',
    `수수료율=${(beforeRow.feeRate * 100).toFixed(1)}%, 수수료=${beforeRow.feeAmt}원, 건당비용=${beforeRow.unitCost}원, 건수비용=${beforeRow.orderCost}원, 순이익=${beforeRow.netProfit}원`);
  console.log('2) 단가 변경일 후(2026-04-17 스마트스토어 20건, 매출 800,000원):',
    `수수료율=${(afterRow.feeRate * 100).toFixed(1)}%, 수수료=${afterRow.feeAmt}원, 건당비용=${afterRow.unitCost}원, 건수비용=${afterRow.orderCost}원, 순이익=${afterRow.netProfit}원`);
  console.log('3) 4월 조회기간 요약:',
    `목표매출=${periodTarget.toLocaleString()}원, 실제매출=${periodActualSales.toLocaleString()}원, 달성률=${(periodActualSales / periodTarget * 100).toFixed(1)}%, 총비용=${periodTotalCost.toLocaleString()}원, 순이익(공통비전)=${periodNetProfit.toLocaleString()}원, 공통비용=${periodCommonCost.toLocaleString()}원, 최종순이익=${periodFinalProfit.toLocaleString()}원`);
}

generateAll();
