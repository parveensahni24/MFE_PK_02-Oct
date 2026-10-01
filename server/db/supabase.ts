import { Pool } from 'pg';
import { createPgPool } from '../config/database.config';
import { RoleCode } from './prisma';

export interface DatabaseDepartment {
  id: string;
  code: RoleCode;
  name: string;
  description?: string | null;
  activeVersionId?: string | null;
  activeVersion?: DatabaseFileVersion | null;
}

export interface DatabaseFileVersion {
  id: string;
  departmentId: string;
  originalFilename: string;
  storageKey?: string | null;
  storagePath?: string | null;
  fileSize?: number | null;
  mimeType?: string | null;
  status: string;
  parsedWorkbook: any;
  rawDataJson?: any;
  uploadedById?: string | null;
  uploadedAt?: Date | string | null;
  processedAt?: Date | string | null;
  createdAt?: Date | string | null;
}

export interface DatabaseMr11Run {
  id: string;
  generatedAt: Date | string;
  status: string;
  sourceSnapshot: any;
  recordCount: number;
  records: any[];
}

/**
 * Fetch all active departments along with their active FileVersion directly from Supabase.
 */
export async function getActiveDepartmentsFromDatabase(): Promise<DatabaseDepartment[]> {
  const pool = createPgPool();
  try {
    const deptRes = await pool.query(`
      SELECT 
        d.id, 
        d.code, 
        d.name, 
        d.description, 
        d."activeVersionId",
        v.id as "version_id",
        v."originalFilename",
        v."storageKey",
        v."storagePath",
        v."fileSize",
        v."mimeType",
        v.status as "version_status",
        v."parsedWorkbook",
        v."rawDataJson",
        v."uploadedById",
        v."uploadedAt",
        v."processedAt",
        v."createdAt" as "version_createdAt"
      FROM "Department" d
      LEFT JOIN "FileVersion" v ON d."activeVersionId" = v.id
      ORDER BY d.code ASC;
    `);

    const departments: DatabaseDepartment[] = deptRes.rows.map((r: any) => {
      let activeVersion: DatabaseFileVersion | null = null;
      if (r.activeVersionId && r.version_id) {
        activeVersion = {
          id: r.version_id,
          departmentId: r.id,
          originalFilename: r.originalFilename,
          storageKey: r.storageKey,
          storagePath: r.storagePath,
          fileSize: r.fileSize,
          mimeType: r.mimeType,
          status: r.version_status || 'READY',
          parsedWorkbook: r.parsedWorkbook,
          rawDataJson: r.rawDataJson,
          uploadedById: r.uploadedById,
          uploadedAt: r.uploadedAt,
          processedAt: r.processedAt,
          createdAt: r.version_createdAt,
        };
      }

      return {
        id: r.id,
        code: r.code as RoleCode,
        name: r.name,
        description: r.description,
        activeVersionId: r.activeVersionId,
        activeVersion,
      };
    });

    return departments;
  } finally {
    await pool.end();
  }
}

/**
 * Fetch the latest MR11 Run directly from Supabase.
 */
export async function getLatestMr11FromDatabase(): Promise<DatabaseMr11Run | null> {
  const pool = createPgPool();
  try {
    const res = await pool.query(
      'SELECT id, "generatedAt", status, "sourceSnapshot", "recordCount", "calculatedFields" FROM "Mr11Run" ORDER BY "generatedAt" DESC LIMIT 1;'
    );

    if (res.rows.length === 0) return null;
    const r = res.rows[0];

    return {
      id: r.id,
      generatedAt: r.generatedAt,
      status: r.status,
      sourceSnapshot: r.sourceSnapshot,
      recordCount: r.recordCount || (Array.isArray(r.calculatedFields) ? r.calculatedFields.length : 0),
      records: Array.isArray(r.calculatedFields) ? r.calculatedFields : [],
    };
  } finally {
    await pool.end();
  }
}

/**
 * Synchronize all active departments and the latest MR11 from Supabase into local Prisma memory.
 * Call this before executing local pipelines or when spinning up an instance.
 */
export async function refreshLocalPrismaFromDatabase(prismaClient: any): Promise<void> {
  try {
    const dbDepts = await getActiveDepartmentsFromDatabase();
    for (const d of dbDepts) {
      if (d.activeVersion) {
        // Upsert FileVersion into local Prisma
        try {
          const existingV = await prismaClient.fileVersion.findUnique({
            where: { id: d.activeVersion.id },
          });
          if (!existingV) {
            await prismaClient.fileVersion.create({
              data: {
                id: d.activeVersion.id,
                departmentId: d.id,
                originalFilename: d.activeVersion.originalFilename,
                storageKey: d.activeVersion.storageKey || `buffer://${d.activeVersion.originalFilename}`,
                storagePath: d.activeVersion.storagePath || d.activeVersion.storageKey,
                fileSize: d.activeVersion.fileSize || 0,
                mimeType: d.activeVersion.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                status: (d.activeVersion.status || 'READY') as any,
                parsedWorkbook: d.activeVersion.parsedWorkbook,
                rawDataJson: d.activeVersion.rawDataJson || null,
                uploadedById: d.activeVersion.uploadedById,
                uploadedAt: d.activeVersion.uploadedAt ? new Date(d.activeVersion.uploadedAt) : new Date(),
                processedAt: d.activeVersion.processedAt ? new Date(d.activeVersion.processedAt) : new Date(),
              },
            });
          }
        } catch (vErr: any) {
          console.warn(`[SUPABASE SYNC] Note caching version ${d.activeVersion.id} into Prisma:`, vErr?.message);
        }

        // Update Department activeVersionId in local Prisma
        try {
          await prismaClient.department.update({
            where: { id: d.id },
            data: { activeVersionId: d.activeVersion.id },
          });
        } catch (dErr: any) {
          console.warn(`[SUPABASE SYNC] Note updating department ${d.code} in Prisma:`, dErr?.message);
        }
      }
    }

    // Also sync latest MR11 run into local Prisma if present
    const latestMr11 = await getLatestMr11FromDatabase();
    if (latestMr11 && latestMr11.records && latestMr11.records.length > 0) {
      try {
        const existingRun = await prismaClient.mr11Run.findUnique({
          where: { id: latestMr11.id },
        });
        if (!existingRun) {
          await prismaClient.mr11Run.create({
            data: {
              id: latestMr11.id,
              status: latestMr11.status as any,
              generatedAt: new Date(latestMr11.generatedAt),
              recordCount: latestMr11.recordCount,
              records: latestMr11.records as any,
              sourceSnapshot: latestMr11.sourceSnapshot,
            },
          });
        }
      } catch (rErr: any) {
        console.warn('[SUPABASE SYNC] Note caching latest MR11 into Prisma:', rErr?.message);
      }
    }
  } catch (err: any) {
    console.warn('[SUPABASE SYNC] General notice refreshing Prisma from Supabase:', err?.message || err);
  }
}

/**
 * Save a FileVersion and update the department's activeVersionId in Supabase.
 */
export async function saveFileVersionToDatabase(versionData: {
  id: string;
  departmentId: string;
  originalFilename: string;
  storageKey: string;
  storagePath: string;
  fileSize: number;
  mimeType: string;
  status: string;
  parsedWorkbook: any;
  rawDataJson: any;
  uploadedById?: string | null;
}): Promise<void> {
  const pool = createPgPool();
  try {
    await pool.query(
      `INSERT INTO "FileVersion" (
        "id", "departmentId", "originalFilename", "storageKey", "storagePath", 
        "fileSize", "mimeType", "status", "parsedWorkbook", "rawDataJson", 
        "uploadedById", "uploadedAt", "processedAt", "createdAt"
      ) VALUES (
        $1, $2, $3, $4, $5, 
        $6, $7, $8, $9, $10, 
        $11, NOW(), NOW(), NOW()
      )
      ON CONFLICT ("id") DO UPDATE SET 
        "parsedWorkbook" = EXCLUDED."parsedWorkbook",
        "rawDataJson" = EXCLUDED."rawDataJson",
        "storagePath" = EXCLUDED."storagePath",
        "status" = 'READY';`,
      [
        versionData.id,
        versionData.departmentId,
        versionData.originalFilename,
        versionData.storageKey,
        versionData.storagePath,
        versionData.fileSize || 0,
        versionData.mimeType,
        versionData.status || 'READY',
        JSON.stringify(versionData.parsedWorkbook),
        JSON.stringify(versionData.rawDataJson || {}),
        versionData.uploadedById || null,
      ]
    );

    await pool.query(
      'UPDATE "Department" SET "activeVersionId" = $1, "updatedAt" = NOW() WHERE id = $2;',
      [versionData.id, versionData.departmentId]
    );
  } finally {
    await pool.end();
  }
}

/**
 * Save an MR11 Run into Supabase.
 */
export async function saveMr11RunToDatabase(runData: {
  id: string;
  status: string;
  sourceSnapshot: any;
  recordCount: number;
  records: any[];
}): Promise<void> {
  const pool = createPgPool();
  try {
    await pool.query(
      `INSERT INTO "Mr11Run" ("id", "generatedAt", "status", "sourceSnapshot", "recordCount", "calculatedFields")
       VALUES ($1, NOW(), $2, $3, $4, $5)
       ON CONFLICT ("id") DO NOTHING;`,
      [
        runData.id,
        runData.status || 'READY',
        JSON.stringify(runData.sourceSnapshot || {}),
        runData.recordCount,
        JSON.stringify(runData.records),
      ]
    );
  } finally {
    await pool.end();
  }
}

export interface DatabaseUser {
  id: string;
  email: string;
  fullName: string;
  passwordHash: string;
  status: string;
  isActive: boolean;
  createdAt: Date | string;
  updatedAt: Date | string;
  roles: Array<{
    id: string;
    roleId: string;
    roleCode: string;
    roleName: string;
  }>;
}

/**
 * Fetch all users along with their roles directly from Supabase.
 */
export async function getUsersFromDatabase(): Promise<DatabaseUser[]> {
  const pool = createPgPool();
  try {
    const res = await pool.query(`
      SELECT 
        u.id, 
        u.email, 
        u."fullName", 
        u."passwordHash", 
        u.status, 
        u."isActive", 
        u."createdAt", 
        u."updatedAt",
        COALESCE(
          json_agg(
            json_build_object(
              'id', ur.id, 
              'roleId', ur."roleId", 
              'roleCode', r.code, 
              'roleName', r.name
            )
          ) FILTER (WHERE ur.id IS NOT NULL), 
          '[]'
        ) as roles
      FROM "User" u
      LEFT JOIN "UserRole" ur ON u.id = ur."userId"
      LEFT JOIN "Role" r ON ur."roleId" = r.id
      GROUP BY u.id
      ORDER BY u."createdAt" ASC;
    `);

    return res.rows;
  } finally {
    await pool.end();
  }
}

/**
 * Synchronize all users and their role assignments from Supabase into local Prisma memory.
 */
export async function syncUsersFromDatabase(prismaClient: any): Promise<void> {
  try {
    const dbUsers = await getUsersFromDatabase();
    for (const u of dbUsers) {
      let localUser = await prismaClient.user.findUnique({ where: { id: u.id } });
      if (!localUser) {
        localUser = await prismaClient.user.findFirst({ where: { email: u.email } });
      }

      if (!localUser) {
        localUser = await prismaClient.user.create({
          data: {
            id: u.id,
            email: u.email,
            fullName: u.fullName,
            passwordHash: u.passwordHash,
            status: u.status || 'ACTIVE',
            isActive: u.isActive !== false,
            createdAt: new Date(u.createdAt),
            updatedAt: new Date(u.updatedAt),
          },
        });
      } else {
        await prismaClient.user.update({
          where: { id: localUser.id },
          data: {
            fullName: u.fullName,
            passwordHash: u.passwordHash,
            status: u.status || 'ACTIVE',
            isActive: u.isActive !== false,
          },
        });
      }

      if (Array.isArray(u.roles)) {
        for (const r of u.roles) {
          if (!r.roleCode) continue;
          let roleRec = await prismaClient.role.findFirst({ where: { code: r.roleCode } });
          if (!roleRec) {
            roleRec = await prismaClient.role.create({
              data: { id: r.roleId || `role-${r.roleCode.toLowerCase()}`, code: r.roleCode, name: r.roleName || r.roleCode },
            });
          }
          const existingUr = await prismaClient.userRole.findFirst({
            where: { userId: localUser.id, roleId: roleRec.id },
          });
          if (!existingUr) {
            await prismaClient.userRole.create({
              data: {
                id: r.id || `ur-${localUser.id}-${roleRec.id}`,
                userId: localUser.id,
                roleId: roleRec.id,
              },
            });
          }
        }
      }
    }
  } catch (err: any) {
    console.warn('[SUPABASE USERS SYNC] Notice:', err?.message || err);
  }
}

/**
 * Persist a newly created user and roles directly to Supabase.
 */
export async function saveUserToDatabase(user: {
  id: string;
  email: string;
  fullName: string;
  passwordHash: string;
  status: string;
  isActive: boolean;
  roleCodes: string[];
}): Promise<void> {
  const pool = createPgPool();
  try {
    await pool.query(
      `INSERT INTO "User" ("id", "email", "fullName", "passwordHash", "status", "isActive", "createdAt", "updatedAt")
       VALUES ($1, $2, $3, $4, $5, $6, NOW(), NOW())
       ON CONFLICT ("id") DO UPDATE SET
         "fullName" = EXCLUDED."fullName",
         "passwordHash" = EXCLUDED."passwordHash",
         "status" = EXCLUDED."status",
         "isActive" = EXCLUDED."isActive",
         "updatedAt" = NOW();`,
      [user.id, user.email, user.fullName, user.status || 'ACTIVE', user.isActive !== false]
    );

    for (const code of user.roleCodes) {
      const roleRes = await pool.query('SELECT id FROM "Role" WHERE code = $1 LIMIT 1;', [code]);
      let roleId = roleRes.rows[0]?.id;
      if (!roleId) {
        roleId = `role-${code.toLowerCase()}`;
        await pool.query(
          'INSERT INTO "Role" ("id", "code", "name", "createdAt") VALUES ($1, $2, $3, NOW()) ON CONFLICT ("id") DO NOTHING;',
          [roleId, code, code]
        );
      }
      const urId = `ur-${user.id}-${roleId}`;
      await pool.query(
        'INSERT INTO "UserRole" ("id", "userId", "roleId", "createdAt") VALUES ($1, $2, $3, NOW()) ON CONFLICT ("id") DO NOTHING;',
        [urId, user.id, roleId]
      );
    }
  } finally {
    await pool.end();
  }
}

/**
 * Update user roles in Supabase.
 */
export async function updateUserRolesInDatabase(userId: string, roleCodes: string[]): Promise<void> {
  const pool = createPgPool();
  try {
    await pool.query('DELETE FROM "UserRole" WHERE "userId" = $1;', [userId]);

    for (const code of roleCodes) {
      const roleRes = await pool.query('SELECT id FROM "Role" WHERE code = $1 LIMIT 1;', [code]);
      let roleId = roleRes.rows[0]?.id;
      if (!roleId) {
        roleId = `role-${code.toLowerCase()}`;
        await pool.query(
          'INSERT INTO "Role" ("id", "code", "name", "createdAt") VALUES ($1, $2, $3, NOW()) ON CONFLICT ("id") DO NOTHING;',
          [roleId, code, code]
        );
      }
      const urId = `ur-${userId}-${roleId}-${Date.now()}`;
      await pool.query(
        'INSERT INTO "UserRole" ("id", "userId", "roleId", "createdAt") VALUES ($1, $2, $3, NOW()) ON CONFLICT ("id") DO NOTHING;',
        [urId, userId, roleId]
      );
    }
  } finally {
    await pool.end();
  }
}
