import { Request, Response } from 'express';
import { PrismaClient } from '@prisma/client';

const prisma = new PrismaClient();

export async function getVisualizationData(req: Request, res: Response) {
  const { getLatestMr11FromDatabase } = await import('../../db/supabase');
  let latestRun: any = await getLatestMr11FromDatabase();

  if (!latestRun || !Array.isArray(latestRun.records) || latestRun.records.length === 0) {
    latestRun = await prisma.mr11Run.findFirst({
      orderBy: { generatedAt: 'desc' },
    });
  }

  if (!latestRun || !Array.isArray(latestRun.records)) {
    return res.json({
      success: true,
      data: {
        kpi: { totalProjects: 0, totalSellingUSD: 0, totalOrderedM2: 0, totalProduced: 0 },
        designStatusCounts: {},
        streamDistribution: {},
      },
    });
  }

  const records = latestRun.records as Record<string, any>[];

  let totalSellingUSD = 0;
  let totalOrderedM2 = 0;
  let totalProduced = 0;
  const designStatusCounts: Record<string, number> = {};
  const streamDistribution: Record<string, number> = {};

  for (const r of records) {
    const price = Number(r['Final Selling Price (USD)']) || Number(r['Selling Price (USD)']) || 0;
    totalSellingUSD += price;

    const ordered =
      Number(r['Total Quantity Ordered m2']) ||
      Number(r['Total Quantity Ordered (m2)']) ||
      Number(r['Total Quantity Ordered']) ||
      Number(r['total quantity ordered m2']) ||
      0;
    totalOrderedM2 += ordered;

    const produced =
      Number(r['Total Produced Quantity']) ||
      Number(r['Total Produced']) ||
      Number(r['produced qty']) ||
      0;
    totalProduced += produced;

    const rawDs =
      r['Formwork Design Status'] ||
      r['design status'] ||
      r['Formwork Design Status - To Start ; Ongoing ; Completed (Dropdown)'] ||
      'Unassigned';
    const ds = String(rawDs).trim();
    const normalizedDs =
      ds.toLowerCase() === 'ongoing'
        ? 'Ongoing'
        : ds.toLowerCase() === 'to start'
        ? 'To Start'
        : ds.toLowerCase() === 'completed'
        ? 'Completed'
        : ds;
    designStatusCounts[normalizedDs] = (designStatusCounts[normalizedDs] || 0) + 1;

    const rawStream = r['Stream'] || '1';
    const stream = String(rawStream).trim().startsWith('Stream')
      ? String(rawStream).trim()
      : `Stream ${String(rawStream).trim()}`;
    streamDistribution[stream] = (streamDistribution[stream] || 0) + 1;
  }

  return res.json({
    success: true,
    data: {
      kpi: {
        totalProjects: records.length,
        totalSellingUSD: Math.round(totalSellingUSD),
        totalOrderedM2: Math.round(totalOrderedM2),
        totalProduced: Math.round(totalProduced),
      },
      designStatusCounts,
      streamDistribution,
    },
  });
}