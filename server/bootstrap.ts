import fs from 'fs';
import path from 'path';
import { prisma, RoleCode } from './db/prisma';
import { processAtomicWorkbookUpload } from './modules/departments/department.service';
import { executeMr11Pipeline } from './modules/mr11/mr11.engine';

const deptKeywords: Record<RoleCode, string> = {
  [RoleCode.BD]: 'bd.xlsx',
  [RoleCode.FINANCE]: 'finance.xlsx',
  [RoleCode.SHELLPLAN]: 'shellplan.xlsx',
  [RoleCode.DESIGN]: 'design.xlsx',
  [RoleCode.PLANNING]: 'planning.xlsx',
  [RoleCode.PRODUCTION]: 'production.xlsx',
  [RoleCode.DISPATCH]: 'dispatch.xlsx',
  [RoleCode.ADMIN]: '',
  [RoleCode.CEO]: '',
};

export async function bootstrapSystem() {
  try {
    const admin = await prisma.user.findFirst({ where: { email: 'admin@mfeformwork.com' } });
    const adminId = admin?.id || 'user-admin-1';

    // 1. On start, load active workbooks and the latest MR11 from the database into local Prisma/memory
    const { getActiveDepartmentsFromDatabase, refreshLocalPrismaFromDatabase } = await import('./db/supabase');
    await refreshLocalPrismaFromDatabase(prisma);

    // 2. Identify which departments have active workbooks in the database
    const dbDepts = await getActiveDepartmentsFromDatabase();
    const initializedDepts = new Set<string>();
    for (const d of dbDepts) {
      if (d.activeVersionId && d.activeVersion) {
        initializedDepts.add(d.code);
      }
    }

    const requiredDepts: RoleCode[] = [
      RoleCode.BD,
      RoleCode.FINANCE,
      RoleCode.SHELLPLAN,
      RoleCode.DESIGN,
      RoleCode.PLANNING,
      RoleCode.PRODUCTION,
      RoleCode.DISPATCH,
    ];

    const missingDepts = requiredDepts.filter((code) => !initializedDepts.has(code));

    if (missingDepts.length === 0) {
      console.log(
        '[MFE Formwork MR11] Loaded active workbooks and latest MR11 from database. All 7 departments active. No bootstrap upload needed.'
      );
      return;
    }

    console.log(
      `[MFE Formwork MR11] Sample files filling ${missingDepts.length} missing department(s): ${missingDepts.join(', ')} (in-memory only, never written back)...`
    );

    const uploadsStorageDir = path.resolve(process.cwd(), 'uploads_storage');
    const uploadsDir = path.resolve(process.cwd(), 'uploads');
    const searchDirs = [uploadsStorageDir, uploadsDir].filter((d) => fs.existsSync(d));

    if (searchDirs.length === 0) {
      console.log('[MFE Formwork MR11] No local upload directories found (relying on Supabase cloud data)');
      return;
    }

    let inMemoryLoadedCount = 0;

    // Sample files ONLY fill departments the database has nothing for, and are NEVER written back (persist: false)
    for (const code of missingDepts) {
      const keyword = deptKeywords[code];
      if (!keyword) continue;

      let targetPath: string | null = null;
      let targetFilename: string | null = null;

      for (const dir of searchDirs) {
        try {
          const files = fs.readdirSync(dir);
          const matching = files
            .filter((f) => f.toLowerCase().endsWith(keyword))
            .sort((a, b) => b.localeCompare(a));

          if (matching.length > 0) {
            targetFilename = matching[0];
            targetPath = path.join(dir, targetFilename);
            break;
          }
        } catch {}
      }

      if (targetPath && targetFilename && fs.existsSync(targetPath)) {
        try {
          const stat = fs.statSync(targetPath);
          await processAtomicWorkbookUpload(
            prisma as any,
            code,
            targetPath,
            targetFilename,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            stat.size,
            adminId,
            { persist: false, regenerate: false } // Never written back to DB
          );
          inMemoryLoadedCount++;
          console.log(`[MFE Formwork MR11] Filled missing department ${code} from sample ${targetFilename} (in-memory only)`);
        } catch (e: any) {
          console.warn(`[MFE Formwork MR11] Note: Could not load sample workbook for ${code}:`, e.message);
        }
      }
    }

    // Build in-memory MR11 calculation if needed without saving to database
    if (inMemoryLoadedCount > 0) {
      try {
        await executeMr11Pipeline(prisma as any, { persist: false });
        console.log('[MFE Formwork MR11] In-memory MR11 Master generation complete');
      } catch (e: any) {
        console.warn('[MFE Formwork MR11] In-memory MR11 pipeline notice:', e.message);
      }
    }
  } catch (err: any) {
    console.error('[MFE Formwork MR11] Bootstrap notice:', err.message);
  }
}
