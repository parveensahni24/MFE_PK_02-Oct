import React, { useState } from 'react';

interface HighFidelityViewerProps {
  data: any[];
}

export const HighFidelityViewer: React.FC<HighFidelityViewerProps> = ({ data }) => {
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);

  if (!data || data.length === 0) {
    return (
      <div className="flex items-center justify-center h-full text-slate-400">
        No workbook sheets parsed.
      </div>
    );
  }

  const activeSheet = data[activeSheetIdx] || data[0];
  const celldata = activeSheet.celldata || [];
  const merges = activeSheet.config?.merge || {};

  // Detect and collapse 2-row header into 1 single header row
  let cleanMerges: Record<string, any> = { ...merges };
  let cleanCelldata: any[] = celldata;

  let r0MergeCount = 0;
  for (const m of Object.values(cleanMerges) as any[]) {
    if (m.r === 0 && m.rs === 2) r0MergeCount++;
  }

  const r0Cells = celldata.filter((c: any) => c.r === 0);
  const r1Cells = celldata.filter((c: any) => c.r === 1);
  let duplicateCount = 0;
  if (r0Cells.length > 0 && r1Cells.length > 0) {
    const r0Map = new Map(r0Cells.map((c: any) => [c.c, String(c.v?.m || c.v?.v || '').trim().toLowerCase()]));
    for (const c1 of r1Cells) {
      const v1 = String(c1.v?.m || c1.v?.v || '').trim().toLowerCase();
      if (v1 && r0Map.get(c1.c) === v1) duplicateCount++;
    }
  }

  const shouldCollapseHeader = r0MergeCount >= 3 || duplicateCount >= 3;

  if (shouldCollapseHeader) {
    const updatedMerges: Record<string, any> = {};
    for (const m of Object.values(cleanMerges) as any[]) {
      if (m.r === 0) {
        // Drop vertical merges and horizontal group merges on row 0 so every sub-column is distinct
      } else if (m.r === 1) {
        // Drop merges originating in redundant header row 1
      } else if (m.r >= 2) {
        updatedMerges[`${m.r - 1}_${m.c}`] = { ...m, r: m.r - 1 };
      }
    }
    cleanMerges = updatedMerges;

    const r0CellMap = new Map<number, any>();
    const r1CellMap = new Map<number, any>();
    for (const cell of celldata) {
      if (cell.r === 0) r0CellMap.set(cell.c, cell);
      else if (cell.r === 1) r1CellMap.set(cell.c, cell);
    }

    cleanCelldata = [];
    const allHeaderCols = Array.from(new Set([...r0CellMap.keys(), ...r1CellMap.keys()])).sort((a, b) => a - b);

    for (const c of allHeaderCols) {
      const c0 = r0CellMap.get(c);
      const c1 = r1CellMap.get(c);
      const val0 = String(c0?.v?.m || c0?.v?.v || '').trim();
      const val1 = String(c1?.v?.m || c1?.v?.v || '').trim();

      const finalCell = c0 ? { ...c0 } : { r: 0, c, v: {} };

      // If row 1 has a sub-header that differs from row 0 or if row 0 was empty
      if (val1 && (!val0 || val0.toLowerCase() !== val1.toLowerCase())) {
        finalCell.v = {
          ...(c1?.v || c0?.v || {}),
          v: val1,
          m: val1,
        };
      } else if (c0) {
        finalCell.v = {
          ...c0.v,
          v: c0.v?.v ?? val0,
          m: c0.v?.m ?? val0,
        };
      }

      finalCell.r = 0;
      finalCell.c = c;
      if (finalCell.v) {
        finalCell.v.rowspan = 1;
        finalCell.v.colspan = 1;
      }
      cleanCelldata.push(finalCell);
    }

    // Number duplicate header names on row 0 (e.g. Type, Type 2, Type 3)
    const headerCounts: Record<string, number> = {};
    for (const cell of cleanCelldata) {
      if (cell.r === 0) {
        const text = String(cell.v?.m || cell.v?.v || '').trim();
        if (text) {
          if (!headerCounts[text]) {
            headerCounts[text] = 1;
          } else {
            headerCounts[text]++;
            const renamed = `${text} ${headerCounts[text]}`;
            cell.v = { ...cell.v, v: renamed, m: renamed };
          }
        }
      }
    }

    // Shift all data rows (r >= 2) up by 1
    for (const cell of celldata) {
      if (cell.r >= 2) {
        cleanCelldata.push({
          ...cell,
          r: cell.r - 1,
        });
      }
    }
  }

  let maxR = 0;
  let maxC = 0;
  const gridMap: Record<string, any> = {};
  const coveredCells = new Set<string>();

  Object.values(cleanMerges).forEach((m: any) => {
    for (let r = m.r; r < m.r + m.rs; r++) {
      for (let c = m.c; c < m.c + m.cs; c++) {
        if (r !== m.r || c !== m.c) {
          coveredCells.add(`${r}_${c}`);
        }
      }
    }
  });

  cleanCelldata.forEach((cell: any) => {
    if (cell.r > maxR) maxR = cell.r;
    if (cell.c > maxC) maxC = cell.c;
    gridMap[`${cell.r}_${cell.c}`] = cell.v;
  });

  const rowIndices = Array.from({ length: Math.min(maxR + 1, 300) }, (_, i) => i);
  const colIndices = Array.from({ length: Math.min(maxC + 1, 60) }, (_, i) => i);

  return (
    <div className="flex flex-col h-full bg-white rounded-lg border border-slate-300 overflow-hidden shadow-inner">
      <div className="flex-1 overflow-auto bg-slate-100">
        <table className="border-collapse bg-white text-xs select-none">
          <thead>
            <tr className="bg-slate-200 sticky top-0 z-20">
              <th className="w-12 min-w-[48px] p-1.5 border border-slate-300 text-center font-bold text-slate-500 bg-slate-200">
                #
              </th>
              {colIndices.map((colIdx) => {
                const colLetter = String.fromCharCode(65 + (colIdx % 26));
                return (
                  <th
                    key={colIdx}
                    className="min-w-[120px] px-2 py-1.5 border border-slate-300 text-center font-semibold text-slate-600 uppercase"
                  >
                    {colLetter}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {rowIndices.map((rowIdx) => (
              <tr key={rowIdx} className="hover:bg-blue-50/20">
                <td className="w-12 min-w-[48px] p-1 border border-slate-300 text-center font-mono text-[11px] text-slate-400 bg-slate-100 sticky left-0 z-10">
                  {rowIdx + 1}
                </td>
                {colIndices.map((colIdx) => {
                  const coordKey = `${rowIdx}_${colIdx}`;

                  if (coveredCells.has(coordKey)) {
                    return null;
                  }

                  const cell = gridMap[coordKey];
                  const mergeInfo = cleanMerges[coordKey];
                  const value = cell?.m || cell?.v;
                  const bg = cell?.bg;
                  const fc = cell?.fc;
                  const isBold = cell?.bl === 1 || rowIdx === 0;

                  return (
                    <td
                      key={colIdx}
                      rowSpan={mergeInfo?.rs || 1}
                      colSpan={mergeInfo?.cs || 1}
                      style={{
                        backgroundColor: bg || (rowIdx === 0 ? '#F1F5F9' : undefined),
                        color: fc || (rowIdx === 0 ? '#0F172A' : undefined),
                      }}
                      className={`p-1.5 border border-slate-300 whitespace-nowrap text-slate-800 ${
                        isBold ? 'font-bold' : ''
                      } ${rowIdx === 0 ? 'bg-slate-100 text-slate-900 font-bold sticky top-6 z-10' : mergeInfo ? 'bg-slate-50/60 font-semibold' : ''}`}
                    >
                      {value !== undefined && value !== null ? String(value) : ''}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 border-t border-slate-300 overflow-x-auto select-none">
        <span className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mr-2">
          Sheets:
        </span>
        {data.map((sheet, idx) => (
          <button
            key={sheet.name || idx}
            onClick={() => setActiveSheetIdx(idx)}
            className={`px-4 py-1.5 rounded-t text-xs font-semibold border-t border-x transition ${
              activeSheetIdx === idx
                ? 'bg-white border-slate-300 text-blue-600 shadow-sm'
                : 'bg-slate-200 border-transparent text-slate-600 hover:bg-slate-300'
            }`}
          >
            {sheet.name || `Sheet${idx + 1}`}
          </button>
        ))}
      </div>
    </div>
  );
};