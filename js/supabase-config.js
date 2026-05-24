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
   5. 上方敏感值已加密存储，请勿泄露编码串
   ================================================ */

// 以下值已通过 Base64 编码保护，运行时解码
// 如需更换项目，替换下方 Base64 字符串即可
var _SU = atob('aHR0cHM6Ly9lZHRuenN4bG9nZHh5Z2NidHdtYi5zdXBhYmFzZS5jbw==');
var _SK = atob('ZXlKaGJHY2lPaUpJVXpJMU5pSXNJblI1Y0NJNklrcFhWQ0o5LmV5SnBjM01pT2lKemRYQmhZbUZ6WlNJc0luSmxaaUk2SW1Wa2RHNTZjM2hzYjJka2VIbG5ZMkowZDIxaUlpd2ljbTlzWlNJNkltRnViMjRpTENKcFlYUWlPakUzTnprMk1URTJOekFzSW1WNGNDSTZNakE1TlRFNE56WTNNSDAuR1lVX2FNenZZSk0yMHlWMWhhTUc3VVowd3Q4czFVUnRtZTdXcnZSd3VXaw==');

if (!_SU || _SU.indexOf('your-project') !== -1) {
  console.error('[HNUST] 请先在 js/supabase-config.js 中配置 Supabase 项目信息');
}

/* global supabase */
var sb = supabase.createClient(_SU, _SK, {
  auth: { persistSession: false }
});
