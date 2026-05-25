-- ================================================
-- HNUST 评论区 — Supabase 数据库初始化脚本
-- 使用方式：Supabase Dashboard → SQL Editor → 粘贴执行
-- ================================================

-- 1. 创建评论表
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

-- 2. 创建索引
CREATE INDEX IF NOT EXISTS idx_comments_created_at ON public.comments (created_at DESC);
CREATE INDEX IF NOT EXISTS idx_comments_parent_id ON public.comments (parent_id);

-- 3. Admin password verification
CREATE EXTENSION IF NOT EXISTS pgcrypto;

-- Replace CHANGE_ME_SHA256_HASH with the SHA-256 hash of your admin password before running.
-- Example in PowerShell:
--   [BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes("your-password"))).Replace("-","").ToLower()
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

-- 4. 开启 Row Level Security
ALTER TABLE public.comments ENABLE ROW LEVEL SECURITY;

-- 5. 匿名访问策略（客户端 anon key 可用）

-- 任何人都可读取（包括已删除的 —— admin 需要看到）
DROP POLICY IF EXISTS "anon_select" ON public.comments;
CREATE POLICY "anon_select" ON public.comments
  FOR SELECT USING (true);

-- 任何人都可插入评论
DROP POLICY IF EXISTS "anon_insert" ON public.comments;
CREATE POLICY "anon_insert" ON public.comments
  FOR INSERT WITH CHECK (true);

-- 只有通过管理员校验的请求可更新（软删除 / 恢复）
DROP POLICY IF EXISTS "anon_update" ON public.comments;
DROP POLICY IF EXISTS "admin_update" ON public.comments;
CREATE POLICY "admin_update" ON public.comments
  FOR UPDATE USING (public.is_admin_request())
  WITH CHECK (public.is_admin_request());

-- 只有通过管理员校验的请求可删除（永久删除）
DROP POLICY IF EXISTS "anon_delete" ON public.comments;
DROP POLICY IF EXISTS "admin_delete" ON public.comments;
CREATE POLICY "admin_delete" ON public.comments
  FOR DELETE USING (public.is_admin_request());

-- 6. 授予匿名角色权限；RLS 负责限制 update/delete 的真实可用范围
GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.comments TO anon;

-- ================================================
-- 增量迁移（已有数据库升级用）
-- ================================================

-- 7. 增加审核状态字段
ALTER TABLE public.comments ADD COLUMN IF NOT EXISTS status TEXT DEFAULT 'active';
-- 修复旧数据
UPDATE public.comments SET status = 'active' WHERE status IS NULL;
-- 增加违禁词记录字段
ALTER TABLE public.comments ADD COLUMN IF NOT EXISTS flagged_word TEXT DEFAULT NULL;

-- 8. 创建设备封禁表
CREATE TABLE IF NOT EXISTS public.banned_devices (
  ip TEXT PRIMARY KEY,
  banned_at TIMESTAMPTZ DEFAULT NOW(),
  reason TEXT DEFAULT ''
);

ALTER TABLE public.banned_devices ENABLE ROW LEVEL SECURITY;

-- 任何人都可以查询封禁列表（用于前端检查自己的 IP）
DROP POLICY IF EXISTS "anyone_check_ban" ON public.banned_devices;
CREATE POLICY "anyone_check_ban" ON public.banned_devices
  FOR SELECT USING (true);

-- 只有管理员可以管理封禁
DROP POLICY IF EXISTS "admin_manage_ban" ON public.banned_devices;
CREATE POLICY "admin_manage_ban" ON public.banned_devices
  FOR ALL USING (public.is_admin_request());

GRANT SELECT ON TABLE public.banned_devices TO anon;
