/* ================================================
   HNUST 评论区 — Supabase 客户端配置
   ================================================
   使用方式：
   1. 在 https://supabase.com 注册并创建项目
   2. 在 SQL Editor 中运行 SUPABASE_SETUP.sql
   3. 将下方 SUPABASE_URL 和 SUPABASE_ANON_KEY
      替换为你的项目值（Settings → API）
   4. 警告：anon key 可安全暴露在前端，
      但请勿泄露 service_role key
   5. anon key 可以公开用于前端；真正的管理员权限由数据库策略限制
   ================================================ */

// Supabase anon key is safe to expose in frontend code. Do not place service_role keys here.
var _SU = 'https://edtnzsxlogdxygcbtwmb.supabase.co';
var _SK = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVkdG56c3hsb2dkeHlnY2J0d21iIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTYxMTY3MCwiZXhwIjoyMDk1MTg3NjcwfQ.GYU_aMzvYJM20yV1haMG7UZ0wt8s1UQtmE7WrvRwuWk';

if (!_SU || _SU.indexOf('your-project') !== -1) {
  console.error('[HNUST] 请先在 js/supabase-config.js 中配置 Supabase 项目信息');
}

/* global supabase */
var sb = supabase.createClient(_SU, _SK, {
  auth: { persistSession: false }
});
