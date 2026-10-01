import { Pool } from 'pg';
import { normalizeDatabaseUrl } from '../db/prisma';

/**
 * Supabase PostgreSQL Database Configuration
 * Project: qbjoclyhqctgrvlrqvgf
 * Region: ap-southeast-1 (Singapore)
 */

// Supabase Transaction Pooler (Recommended for Vercel Serverless & Cloud runtimes)
export const SUPABASE_POOLER_URL =
  'postgresql://postgres.qbjoclyhqctgrvlrqvgf:Irely19612026@aws-0-ap-southeast-1.pooler.supabase.com:6543/postgres?pgbouncer=true';

// Supabase Direct PostgreSQL Connection
export const SUPABASE_DIRECT_URL =
  'postgresql://postgres:Irely19612026@db.qbjoclyhqctgrvlrqvgf.supabase.co:5432/postgres';

export function getDatabaseUrl(): string {
  const raw = process.env.DATABASE_URL || SUPABASE_POOLER_URL;
  return normalizeDatabaseUrl(raw) || raw;
}

export function createPgPool(connectionUrl?: string): Pool {
  const url = connectionUrl ? (normalizeDatabaseUrl(connectionUrl) || connectionUrl) : getDatabaseUrl();
  return new Pool({
    connectionString: url,
    ssl: { rejectUnauthorized: false },
    connectionTimeoutMillis: 5000,
  });
}

