import { NextRequest, NextResponse } from 'next/server';
import { ExcelProductEngine } from '../../../lib/excel-engine';
import { Product } from '../../../types/schema';

export async function POST(req: NextRequest) {
  let body: any;
  try {
    body = await req.json();
  } catch (err) {
    console.error('[API/Generate] Invalid JSON body:', err);
    return NextResponse.json({ error: 'Request body must be valid JSON.' }, { status: 400 });
  }

  const product: Product = body?.product || body;

  if (!product || !product.sheets || !Array.isArray(product.sheets)) {
    return NextResponse.json(
      { error: 'Invalid product specification. Sheets array is required.' },
      { status: 400 }
    );
  }

  try {
    const engine = new ExcelProductEngine(product);
    const buffer = await engine.exportAsBuffer();

    const fileName = `${(product.name || 'workbook').toLowerCase().replace(/[^a-z0-9]/g, '_')}_v${product.version || '1.0.0'}.xlsx`;

    // Return the binary file stream
    return new NextResponse(buffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${fileName}"`,
        'X-File-Name': fileName,
        'X-Sheet-Count': String(product.sheets.length),
      },
    });
  } catch (err: any) {
    // Log full detail server-side only. Never echo internal error/stack text
    // back to the client — it can leak implementation details for no benefit
    // to the caller.
    console.error('[API/Generate] Workbook compilation failed:', err);
    return NextResponse.json(
      { error: 'Failed to compile the Excel workbook. Please check your sheet, column, and KPI configuration.' },
      { status: 500 }
    );
  }
}
