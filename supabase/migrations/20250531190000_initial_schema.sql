-- Migration: Initial schema
-- This captures the full database schema from SUPABASE_SETUP.sql

-- 1. Comments table
CREATE TABLE IF NOT EXISTS public.comments (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  nickname TEXT NOT NULL,
  ip TEXT DEFAULT '',
  content TEXT NOT NULL,
  parent_id UUID DEFAULT NULL,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  deleted BOOLEAN DEFAULT FALSE,
  deleted_at TIMESTAMPTZ
);

CREATE INDEX IF NOT EXISTS idx_comments_created_at ON public.comments (created_at DESC);
CREATE INDEX IF NOT EXISTS idx_comments_parent_id ON public.comments (parent_id);

-- 2. Admin settings
CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS public.admin_settings (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL
);

ALTER TABLE public.admin_settings ENABLE ROW LEVEL SECURITY;
REVOKE ALL ON TABLE public.admin_settings FROM anon, authenticated;

INSERT INTO public.admin_settings (key, value)
VALUES ('admin_pass_hash', '8995c1676864ffb8f048919db374e4da951028c5f7768c931669acb0873b31cf')
ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;

CREATE OR REPLACE FUNCTION public.verify_admin_password(password text)
RETURNS boolean
LANGUAGE sql
SECURITY DEFINER
SET search_path = public, extensions
AS $$
  SELECT EXISTS (
    SELECT 1
    FROM public.admin_settings
    WHERE key = 'admin_pass_hash'
      AND value <> ''
      AND value <> 'CHANGE_ME_SHA256_HASH'
      AND value = encode(extensions.digest(coalesce(password, ''), 'sha256'), 'hex')
  );
$$;

CREATE OR REPLACE FUNCTION public.is_admin_request()
RETURNS boolean
LANGUAGE sql
SECURITY DEFINER
SET search_path = public
AS $$
  SELECT
    lower(coalesce((nullif(current_setting('request.headers', true), '')::jsonb ->> 'x-admin'), '')) = 'true'
    AND public.verify_admin_password(nullif(current_setting('request.headers', true), '')::jsonb ->> 'x-admin-password');
$$;

GRANT EXECUTE ON FUNCTION public.verify_admin_password(text) TO anon;
GRANT EXECUTE ON FUNCTION public.is_admin_request() TO anon;

-- 3. Row Level Security
ALTER TABLE public.comments ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anon_select" ON public.comments;
CREATE POLICY "anon_select" ON public.comments
  FOR SELECT USING (true);

DROP POLICY IF EXISTS "anon_insert" ON public.comments;
CREATE POLICY "anon_insert" ON public.comments
  FOR INSERT WITH CHECK (true);

DROP POLICY IF EXISTS "anon_update" ON public.comments;
DROP POLICY IF EXISTS "admin_update" ON public.comments;
CREATE POLICY "admin_update" ON public.comments
  FOR UPDATE USING (public.is_admin_request())
  WITH CHECK (public.is_admin_request());

DROP POLICY IF EXISTS "anon_delete" ON public.comments;
DROP POLICY IF EXISTS "admin_delete" ON public.comments;
CREATE POLICY "admin_delete" ON public.comments
  FOR DELETE USING (public.is_admin_request());

GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.comments TO anon;

-- 4. Status and moderation fields
ALTER TABLE public.comments ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'active';
UPDATE public.comments SET status = 'active' WHERE status IS NULL;
ALTER TABLE public.comments ADD COLUMN IF NOT EXISTS flagged_word TEXT DEFAULT NULL;

-- 5. Banned devices table
CREATE TABLE IF NOT EXISTS public.banned_devices (
  ip TEXT PRIMARY KEY,
  banned_at TIMESTAMPTZ DEFAULT NOW(),
  reason TEXT DEFAULT ''
);

ALTER TABLE public.banned_devices ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "anyone_check_ban" ON public.banned_devices;
CREATE POLICY "anyone_check_ban" ON public.banned_devices
  FOR SELECT USING (true);

DROP POLICY IF EXISTS "admin_manage_ban" ON public.banned_devices;
CREATE POLICY "admin_manage_ban" ON public.banned_devices
  FOR ALL USING (public.is_admin_request());

GRANT SELECT ON TABLE public.banned_devices TO anon;
