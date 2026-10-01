import { Request, Response } from 'express';
import { PrismaClient } from '@prisma/client';
import ExcelJS from 'exceljs';
import { executeMr11Pipeline } from './mr11.engine';
import * as mr11ConfigModule from '../../config/mr11.config';

const prisma = new PrismaClient();

const ORDERED_HEADER_LIST: string[] =
  (mr11ConfigModule as any).ORDERED_HEADER_LIST ||
  ((mr11ConfigModule as any).MR11_ORDERED_COLUMNS || []).map((col: any) => col.target || col.header || String(col));

export async function getLatestMr11(req: Request, res: Response) {
  try {
    const { getLatestMr11FromDatabase } = await import('../../db/supabase');
    let latestRun: any = await getLatestMr11FromDatabase();

    if (!latestRun || !latestRun.records || latestRun.records.length === 0) {
      latestRun = await prisma.mr11Run.findFirst({
        orderBy: { generatedAt: 'desc' },
      });
    }

    const config = await prisma.mr11Config.findUnique({ where: { id: 'singleton' } });

    return res.json({
      success: true,
      data: {
        run: latestRun,
        visibleColumns: (config?.visibleColumns as string[]) || [],
        orderedHeaders: ORDERED_HEADER_LIST,
      },
    });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}

export async function triggerMr11Regenerate(req: Request, res: Response) {
  try {
    // Before rebuilding MR11, refresh other departments from database
    const { refreshLocalPrismaFromDatabase } = await import('../../db/supabase');
    await refreshLocalPrismaFromDatabase(prisma);
    const runId = await executeMr11Pipeline(prisma);
    return res.json({ success: true, message: 'MR11 regenerated', runId });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}

export async function getPlanningSeriesHistory(req: Request, res: Response) {
  try {
    const data = await prisma.planningSeriesHistory.findMany({
      orderBy: [
        { projectNo: 'asc' },
        { stream: 'asc' },
        { seriesNumber: 'asc' },
      ],
    });
    return res.json({ success: true, data });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}

export async function getProductionSeriesHistory(req: Request, res: Response) {
  try {
    const data = await prisma.productionSeriesHistory.findMany({
      orderBy: [
        { projectShortname: 'asc' },
        { stream: 'asc' },
        { seriesNumber: 'asc' },
      ],
    });
    return res.json({ success: true, data });
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message } });
  }
}

export async function exportMr11ToExcel(req: Request, res: Response) {
  try {
    const { getLatestMr11FromDatabase } = await import('../../db/supabase');
    let latestRun: any = await getLatestMr11FromDatabase();

    if (!latestRun || !Array.isArray(latestRun.records) || latestRun.records.length === 0) {
      latestRun = await prisma.mr11Run.findFirst({
        orderBy: { generatedAt: 'desc' },
      });
    }

    if (!latestRun || !Array.isArray(latestRun.records) || latestRun.records.length === 0) {
      return res.status(400).json({ success: false, error: { message: 'No MR11 records to export' } });
    }

    const records = latestRun.records as Record<string, any>[];
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('MR11 Master');

    const headers = ORDERED_HEADER_LIST;

    worksheet.columns = headers.map((header) => ({
      key: header,
      width: Math.max(header.length + 4, 16),
    }));

    // Header strictly covers 1 single row (Row 1)
    const r1 = worksheet.getRow(1);
    r1.height = 28;

    headers.forEach((h, idx) => {
      const colNumber = idx + 1;
      r1.getCell(colNumber).value = h;
    });

    r1.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    r1.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    r1.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' },
    };

    records.forEach((row) => {
      const orderedRowData: Record<string, any> = {};
      headers.forEach((h) => {
        orderedRowData[h] = row[h] ?? '';
      });
      const addedRow = worksheet.addRow(orderedRowData);

      if (row['_fontColor']) {
        const hex = String(row['_fontColor']).replace('#', '');
        const projCell = addedRow.getCell(1);
        projCell.font = {
          color: { argb: `FF${hex}` },
          bold: true,
        };
      }

      if (row['_fillColor']) {
        const hex = String(row['_fillColor']).replace('#', '');
        addedRow.eachCell({ includeEmpty: true }, (cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: `FF${hex}` },
          };
        });
      }
    });

    // Merged ONLY for ShellPlan and Design across the [Project, Stream] span
    const STREAM_MERGE_COLS = [
      'Shell Plan Status - Pending Consultant Drawings', // Col AJ
      'Shell Plan Approved Date',                         // Col AK
      'Formwork Design Status',                          // Col AL
      'Actual Formwork Order Completion Date',           // Col AM
      'Total Quantity Ordered m2',                       // Col AN
    ];

    STREAM_MERGE_COLS.forEach((colName) => {
      const colIdx = headers.findIndex(
        (h) => h === colName || h.toLowerCase().trim() === colName.toLowerCase().trim()
      ) + 1;

      if (colIdx > 0) {
        let r = 2; // Excel row index starts at 2 (Row 1 is the single header row)
        for (let i = 0; i < records.length; ) {
          const span = records[i]['_streamSpan'] || 1;
          if (span > 1) {
            worksheet.mergeCells(r, colIdx, r + span - 1, colIdx);
            const mergedCell = worksheet.getCell(r, colIdx);
            mergedCell.alignment = { vertical: 'middle', horizontal: 'center' };
          }
          r += span;
          i += span;
        }
      }
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=MR11_Master_Order_${Date.now()}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err: any) {
    return res.status(500).json({ success: false, error: { message: err.message || 'Export failed' } });
  }
}