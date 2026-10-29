const { escXml } = require('./xlsx_writer');

function buildBarChartXml({ title, categoriesFormula, categoriesCache, seriesList, barDir = 'col', grouping = 'clustered', showVal = true }) {
  let serXml = '';
  seriesList.forEach((s, idx) => {
    let catPts = '';
    categoriesCache.forEach((c, i) => {
      catPts += `<c:pt idx="${i}"><c:v>${escXml(c)}</c:v></c:pt>`;
    });
    let valPts = '';
    s.valuesCache.forEach((v, i) => {
      valPts += `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`;
    });

    serXml += `
      <c:ser>
        <c:idx val="${idx}"/>
        <c:order val="${idx}"/>
        <c:tx>
          <c:strRef>
            <c:f>${escXml(s.nameFormula)}</c:f>
            <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escXml(s.name)}</c:v></c:pt></c:strCache>
          </c:strRef>
        </c:tx>
        <c:spPr>
          <a:solidFill><a:srgbClr val="${s.color}"/></a:solidFill>
        </c:spPr>
        <c:invertIfNegative val="0"/>
        <c:cat>
          <c:strRef>
            <c:f>${escXml(categoriesFormula)}</c:f>
            <c:strCache><c:ptCount val="${categoriesCache.length}"/>${catPts}</c:strCache>
          </c:strRef>
        </c:cat>
        <c:val>
          <c:numRef>
            <c:f>${escXml(s.valuesFormula)}</c:f>
            <c:numCache>
              <c:formatCode>#,##0</c:formatCode>
              <c:ptCount val="${s.valuesCache.length}"/>${valPts}
            </c:numCache>
          </c:numRef>
        </c:val>
      </c:ser>`;
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill></a:defRPr></a:pPr>
            <a:r><a:rPr lang="ko-KR" sz="1100" b="1"/><a:t>${escXml(title)}</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="${barDir}"/>
        <c:grouping val="${grouping}"/>
        <c:varyColors val="0"/>
        ${serXml}
        <c:dLbls>
          <c:showLegendKey val="0"/>
          <c:showVal val="${showVal ? '1' : '0'}"/>
          <c:showCatName val="0"/>
          <c:showSerName val="0"/>
          <c:showPercent val="0"/>
          <c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:gapWidth val="150"/>
        <c:overlap val="-15"/>
        <c:axId val="101"/>
        <c:axId val="102"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="101"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="102"/>
        <c:crosses val="autoZero"/>
        <c:auto val="1"/>
        <c:lblAlgn val="ctr"/>
        <c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="102"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines>
          <c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E2E8F0"/></a:solidFill></a:ln></c:spPr>
        </c:majorGridlines>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:majorTickMark val="none"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:crossAx val="101"/>
        <c:crosses val="autoZero"/>
        <c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>`;
}

function buildDoughnutChartXml({ title, categoriesFormula, categoriesCache, seriesNameFormula, seriesName, valuesFormula, valuesCache }) {
  let catPts = '';
  categoriesCache.forEach((c, i) => {
    catPts += `<c:pt idx="${i}"><c:v>${escXml(c)}</c:v></c:pt>`;
  });
  let valPts = '';
  valuesCache.forEach((v, i) => {
    valPts += `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`;
  });

  const sliceColors = ['2563EB', '10B981', 'F59E0B', '6366F1', 'EC4899', '14B8A6', '64748B'];
  let dPtXml = '';
  categoriesCache.forEach((_, i) => {
    const clr = sliceColors[i % sliceColors.length];
    dPtXml += `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${clr}"/></a:solidFill></c:spPr></c:dPt>`;
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:roundedCorners val="1"/>
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1100" b="1"><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill></a:defRPr></a:pPr>
            <a:r><a:rPr lang="ko-KR" sz="1100" b="1"/><a:t>${escXml(title)}</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>${escXml(seriesNameFormula)}</c:f>
              <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escXml(seriesName)}</c:v></c:pt></c:strCache>
            </c:strRef>
          </c:tx>
          ${dPtXml}
          <c:dLbls>
            <c:numFmt formatCode="#,##0" sourceLinked="0"/>
            <c:showLegendKey val="0"/>
            <c:showVal val="1"/>
            <c:showCatName val="1"/>
            <c:showSerName val="0"/>
            <c:showPercent val="1"/>
            <c:showBubbleSize val="0"/>
            <c:separator>, </c:separator>
          </c:dLbls>
          <c:cat>
            <c:strRef>
              <c:f>${escXml(categoriesFormula)}</c:f>
              <c:strCache><c:ptCount val="${categoriesCache.length}"/>${catPts}</c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>${escXml(valuesFormula)}</c:f>
              <c:numCache>
                <c:formatCode>#,##0</c:formatCode>
                <c:ptCount val="${valuesCache.length}"/>${valPts}
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:firstSliceAng val="0"/>
        <c:holeSize val="55"/>
      </c:doughnutChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
    <a:ln w="9525"><a:solidFill><a:srgbClr val="CBD5E1"/></a:solidFill></a:ln>
  </c:spPr>
</c:chartSpace>`;
}

function buildDrawingXml(anchors) {
  let anchorXml = '';
  anchors.forEach((a, idx) => {
    anchorXml += `
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>${a.fromCol}</xdr:col>
      <xdr:colOff>95250</xdr:colOff>
      <xdr:row>${a.fromRow}</xdr:row>
      <xdr:rowOff>47625</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>${a.toCol}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>${a.toRow}</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="${idx + 2}" name="Chart ${idx + 1}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${a.relId}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>`;
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
${anchorXml}
</xdr:wsDr>`;
}

module.exports = {
  buildBarChartXml,
  buildDoughnutChartXml,
  buildDrawingXml
};
