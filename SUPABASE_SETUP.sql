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

-- 3. 开启 Row Level Security
ALTER TABLE public.comments ENABLE ROW LEVEL SECURITY;

-- 4. 匿名访问策略（客户端 anon key 可用）

-- 任何人都可读取（包括已删除的 —— admin 需要看到）
DROP POLICY IF EXISTS "anon_select" ON public.comments;
CREATE POLICY "anon_select" ON public.comments
  FOR SELECT USING (true);

-- 任何人都可插入评论
DROP POLICY IF EXISTS "anon_insert" ON public.comments;
CREATE POLICY "anon_insert" ON public.comments
  FOR INSERT WITH CHECK (true);

-- 任何人都可更新（软删除 / 恢复 —— 客户端控制权限）
DROP POLICY IF EXISTS "anon_update" ON public.comments;
CREATE POLICY "anon_update" ON public.comments
  FOR UPDATE USING (true)
  WITH CHECK (true);

-- 任何人都可删除（永久删除 —— 客户端控制权限）
DROP POLICY IF EXISTS "anon_delete" ON public.comments;
CREATE POLICY "anon_delete" ON public.comments
  FOR DELETE USING (true);

-- 5. 授予匿名角色权限
GRANT ALL ON TABLE public.comments TO anon;

-- ================================================
-- 已有数据库迁移（仅需执行一次）
-- 如果是从旧版升级，取消注释并执行：
-- ================================================
-- ALTER TABLE public.comments ADD COLUMN IF NOT EXISTS parent_id UUID DEFAULT NULL;
-- CREATE INDEX IF NOT EXISTS idx_comments_parent_id ON public.comments (parent_id);
