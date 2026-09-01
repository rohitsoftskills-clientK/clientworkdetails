import JSZip from 'jszip';
import { ChartType } from '../types/schema';

/**
 * ExcelJS has no API for writing native Excel chart objects — see the
 * long-standing, unresolved upstream issues exceljs/exceljs#141 and #1569.
 * A .xlsx file is just a zip of OOXML parts, though, so this module builds
 * the workbook with ExcelJS as usual and then adds the chart/drawing XML
 * parts directly, the same way Excel itself would write them. The result is
 * a real, editable chart object — not a static image — that references the
 * live cell ranges ExcelProductEngine already wrote, so it stays in sync if
 * someone edits the data in Excel afterwards.
 */

export interface ChartSeries {
  name: string;
  nameRef: string;
  valuesRef: string;
}

export interface ChartPlacement {
  /** Must exactly match the sheet name as written into the workbook. */
  sheetName: string;
  title: string;
  type: ChartType;
  categoryRef: string;
  series: ChartSeries[];
  /** 0-based cell coordinates, per the OOXML spreadsheetDrawing anchor model. */
  anchor: { fromCol: number; fromRow: number; toCol: number; toRow: number };
}

export function quoteSheetRef(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function escapeXml(value: string): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildSeriesXml(ser: ChartSeries, idx: number, categoryRef: string): string {
  return (
    `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/>` +
    `<c:tx><c:strRef><c:f>${escapeXml(ser.nameRef)}</c:f>` +
    `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(ser.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    `<c:cat><c:strRef><c:f>${escapeXml(categoryRef)}</c:f></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>${escapeXml(ser.valuesRef)}</c:f></c:numRef></c:val></c:ser>`
  );
}

function buildChartXml(placement: ChartPlacement, axisIdBase: number): string {
  const catAxId = axisIdBase;
  const valAxId = axisIdBase + 1;
  const seriesXml = placement.series.map((s, i) => buildSeriesXml(s, i, placement.categoryRef)).join('');

  const titleXml =
    `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>` +
    `<a:p><a:r><a:t>${escapeXml(placement.title)}</a:t></a:r></a:p>` +
    `</c:rich></c:tx><c:overlay val="0"/></c:title>`;

  let plotXml: string;
  let axesXml = '';

  if (placement.type === 'pie') {
    plotXml = `<c:pieChart><c:varyColors val="1"/>${seriesXml}<c:firstSliceAng val="0"/></c:pieChart>`;
  } else if (placement.type === 'line') {
    plotXml =
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml}` +
      `<c:marker val="1"/><c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:lineChart>`;
    axesXml =
      `<c:catAx><c:axId val="${catAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="${valAxId}"/></c:catAx>` +
      `<c:valAx><c:axId val="${valAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="${catAxId}"/></c:valAx>`;
  } else {
    // 'column' (vertical bars) or 'bar' (horizontal bars)
    const barDir = placement.type === 'bar' ? 'bar' : 'col';
    const catAxPos = placement.type === 'bar' ? 'l' : 'b';
    const valAxPos = placement.type === 'bar' ? 'b' : 'l';
    plotXml =
      `<c:barChart><c:barDir val="${barDir}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}` +
      `<c:axId val="${catAxId}"/><c:axId val="${valAxId}"/></c:barChart>`;
    axesXml =
      `<c:catAx><c:axId val="${catAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catAxPos}"/><c:crossAx val="${valAxId}"/></c:catAx>` +
      `<c:valAx><c:axId val="${valAxId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valAxPos}"/><c:crossAx val="${catAxId}"/></c:valAx>`;
  }

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:chart>${titleXml}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${plotXml}${axesXml}</c:plotArea>` +
    `<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart>` +
    `</c:chartSpace>`
  );
}

function buildDrawingXml(entries: { rId: string; anchor: ChartPlacement['anchor'] }[]): string {
  const anchorsXml = entries
    .map((entry, i) => {
      const shapeId = 1000 + i;
      const { fromCol, fromRow, toCol, toRow } = entry.anchor;
      return (
        `<xdr:twoCellAnchor>` +
        `<xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
        `<xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
        `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="Chart ${shapeId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
        `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
        `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
        `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${entry.rId}"/>` +
        `</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
      );
    })
    .join('');

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchorsXml}</xdr:wsDr>`
  );
}

export async function injectCharts(buffer: ArrayBuffer | Buffer, placements: ChartPlacement[]): Promise<Buffer> {
  const original = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
  if (!placements || placements.length === 0) {
    return original;
  }

  const zip = await JSZip.loadAsync(original);

  const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
  const workbookRelsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');

  if (!workbookXml || !workbookRelsXml || !contentTypesXml) {
    // Unexpected package shape (e.g. a future ExcelJS version restructures
    // parts). Fail safe: ship the chart-less workbook rather than throwing
    // and losing the whole export over a cosmetic feature.
    console.error('[chart-injector] Unexpected workbook structure — skipping chart injection.');
    return original;
  }

  const sheetNameToRid = new Map<string, string>();
  for (const m of workbookXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const tag = m[0];
    const name = tag.match(/name="([^"]*)"/)?.[1];
    const rid = tag.match(/r:id="([^"]*)"/)?.[1];
    if (name && rid) sheetNameToRid.set(name, rid);
  }
  const ridToTarget = new Map<string, string>();
  for (const m of workbookRelsXml.matchAll(/<Relationship\b[^>]*\/>/g)) {
    const tag = m[0];
    const id = tag.match(/Id="([^"]*)"/)?.[1];
    const target = tag.match(/Target="([^"]*)"/)?.[1];
    if (id && target) ridToTarget.set(id, target);
  }

  const bySheet = new Map<string, ChartPlacement[]>();
  for (const p of placements) {
    const list = bySheet.get(p.sheetName) || [];
    list.push(p);
    bySheet.set(p.sheetName, list);
  }

  let globalChartCounter = 0;
  let globalDrawingCounter = 0;
  let axisIdCounter = 100000000;
  const newContentTypeOverrides: string[] = [];

  for (const [sheetName, chartsForSheet] of bySheet) {
    const rid = sheetNameToRid.get(sheetName);
    const target = rid ? ridToTarget.get(rid) : undefined;
    if (!target) {
      console.error(`[chart-injector] Could not resolve worksheet file for sheet "${sheetName}" — skipping its charts.`);
      continue;
    }

    const sheetPath = `xl/${target}`;
    const sheetXml = await zip.file(sheetPath)?.async('string');
    if (!sheetXml) {
      console.error(`[chart-injector] Worksheet part not found at ${sheetPath} — skipping its charts.`);
      continue;
    }

    const sheetFileName = sheetPath.split('/').pop()!;
    const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;

    let sheetRelsXml = await zip.file(sheetRelsPath)?.async('string');
    let nextSheetRelId = 1;
    if (sheetRelsXml) {
      const existingIds = [...sheetRelsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
      if (existingIds.length > 0) nextSheetRelId = Math.max(...existingIds) + 1;
    } else {
      sheetRelsXml =
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
    }

    globalDrawingCounter += 1;
    const drawingIndex = globalDrawingCounter;
    const drawingPath = `xl/drawings/drawing${drawingIndex}.xml`;
    const drawingRelsPath = `xl/drawings/_rels/drawing${drawingIndex}.xml.rels`;

    const drawingRelEntries: string[] = [];
    const drawingAnchorEntries: { rId: string; anchor: ChartPlacement['anchor'] }[] = [];

    chartsForSheet.forEach((placement, i) => {
      globalChartCounter += 1;
      const chartIndex = globalChartCounter;
      const chartPath = `xl/charts/chart${chartIndex}.xml`;
      zip.file(chartPath, buildChartXml(placement, axisIdCounter));
      axisIdCounter += 2;
      newContentTypeOverrides.push(
        `<Override PartName="/${chartPath}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
      );

      const localRid = `rId${i + 1}`;
      drawingRelEntries.push(
        `<Relationship Id="${localRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartIndex}.xml"/>`
      );
      drawingAnchorEntries.push({ rId: localRid, anchor: placement.anchor });
    });

    zip.file(drawingPath, buildDrawingXml(drawingAnchorEntries));
    zip.file(
      drawingRelsPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRelEntries.join('')}</Relationships>`
    );
    newContentTypeOverrides.push(
      `<Override PartName="/${drawingPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
    );

    const sheetDrawingRid = `rId${nextSheetRelId}`;
    const updatedSheetRelsXml = sheetRelsXml.replace(
      '</Relationships>',
      `<Relationship Id="${sheetDrawingRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingIndex}.xml"/></Relationships>`
    );
    zip.file(sheetRelsPath, updatedSheetRelsXml);

    // <drawing> must precede <extLst> per the CT_Worksheet schema order; if
    // there's no extLst, the end of the element is a safe insertion point.
    const drawingTag = `<drawing r:id="${sheetDrawingRid}"/>`;
    const updatedSheetXml = sheetXml.includes('<extLst>')
      ? sheetXml.replace('<extLst>', `${drawingTag}<extLst>`)
      : sheetXml.replace('</worksheet>', `${drawingTag}</worksheet>`);
    zip.file(sheetPath, updatedSheetXml);
  }

  if (newContentTypeOverrides.length > 0) {
    zip.file('[Content_Types].xml', contentTypesXml.replace('</Types>', `${newContentTypeOverrides.join('')}</Types>`));
  }

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
