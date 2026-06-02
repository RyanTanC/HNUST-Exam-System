// HNUST 考试仿真平台 - 使用统计 Worker (v2)
// 完全避免 KV list() 操作，通过计数器 + 预聚合方式统计

// ── 简单内存限流（Worker 实例级别）────────────────────
const rateLimits = new Map();
const RATE_WINDOW_MS = 60_000; // 1 分钟
const RATE_MAX = 30;           // 每分钟最多 30 次

function isRateLimited(ip) {
  const now = Date.now();
  const entry = rateLimits.get(ip);
  if (!entry || now - entry.start > RATE_WINDOW_MS) {
    rateLimits.set(ip, { start: now, count: 1 });
    return false;
  }
  entry.count++;
  return entry.count > RATE_MAX;
}

// ── 路由 ────────────────────────────────────────────────
export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    try {
      if (request.method === "POST" && url.pathname === "/api/heartbeat") {
        return handleHeartbeat(request, env);
      }
      if (request.method === "POST" && url.pathname === "/api/submit-score") {
        return handleSubmitScore(request, env);
      }
      if (request.method === "GET" && url.pathname === "/api/stats") {
        return handleStatsAPI(request, env);
      }
      if (request.method === "GET" && url.pathname === "/api/export") {
        return handleExport(request, env);
      }
      if (request.method === "POST" && url.pathname === "/api/import") {
        return handleImport(request, env);
      }
      if (request.method === "GET" && url.pathname === "/admin") {
        return handleAdmin(request, env);
      }
      return new Response("Not Found", { status: 404 });
    } catch (e) {
      console.error(JSON.stringify({
        event: "telemetry_worker_error",
        path: url.pathname,
        message: e instanceof Error ? e.message : String(e),
      }));
      return json({ error: "internal error" }, 500);
    }
  },
};

// ── 工具函数 ────────────────────────────────────────────
function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", "Access-Control-Allow-Origin": "*" },
  });
}

function getClientIP(request) {
  return request.headers.get("CF-Connecting-IP") || "unknown";
}

async function readJson(request) {
  const contentLength = Number(request.headers.get("content-length") || "0");
  if (contentLength > 16_384) {
    throw new Error("request body too large");
  }
  try {
    return await request.json();
  } catch {
    throw new Error("invalid json");
  }
}

function todayKey() {
  return new Date().toISOString().slice(0, 10); // YYYY-MM-DD
}

// KV 模拟原子自增：get → +n → put
async function incr(env, key, n = 1) {
  const raw = await env.HNUST_TELEMETRY.get(key);
  const val = (parseInt(raw, 10) || 0) + n;
  await env.HNUST_TELEMETRY.put(key, String(val));
  return val;
}

// ── 收集统计数据（无 list 操作）─────────────────────────
async function collectStats(env) {
  const today = todayKey();

  // 直接读取预聚合的计数器
  const totalDevices = parseInt(await env.HNUST_TELEMETRY.get("total_devices"), 10) || 0;
  const totalSubmissions = parseInt(await env.HNUST_TELEMETRY.get("total_submissions"), 10) || 0;
  const todayDAU = parseInt(await env.HNUST_TELEMETRY.get(`daily_count:${today}`), 10) || 0;

  // 30 天每日活跃：逐键读取（30 次 get，不是 list）
  const dailyCounts = [];
  const now = Date.now();
  for (let i = 29; i >= 0; i--) {
    const d = new Date(now - i * 86400_000).toISOString().slice(0, 10);
    const count = parseInt(await env.HNUST_TELEMETRY.get(`daily_count:${d}`), 10) || 0;
    dailyCounts.push({ date: d, count });
  }

  // 试卷统计 & 最近提交
  const examStatsRaw = JSON.parse(await env.HNUST_TELEMETRY.get("exam_stats_list") || "[]");
  const examStats = examStatsRaw.map(e => ({
    name: e.name,
    count: e.count,
    avg_score: e.score_count ? +(e.total_score / e.score_count).toFixed(1) : null,
    avg_duration: e.duration_count ? Math.round(e.total_duration / e.duration_count / 60) : null,
  }));
  const recentScores = JSON.parse(await env.HNUST_TELEMETRY.get("recent_scores") || "[]");

  return { todayDAU, totalDevices, totalSubmissions, dailyCounts, examStats, recentScores };
}

// ── GET /api/stats?token=xxx ────────────────────────────
async function handleStatsAPI(request, env) {
  const url = new URL(request.url);
  const token = url.searchParams.get("token");
  if (!token || token !== env.ADMIN_TOKEN) {
    return json({ error: "unauthorized" }, 401);
  }
  return json(await collectStats(env));
}

// ── GET /api/export?token=xxx ───────────────────────────
async function handleExport(request, env) {
  const url = new URL(request.url);
  const token = url.searchParams.get("token");
  if (!token || token !== env.ADMIN_TOKEN) {
    return json({ error: "unauthorized" }, 401);
  }

  // 收集所有 daily_count 键（最近 32 天，由 TTL 控制）
  const dailyCounts = {};
  const now = Date.now();
  for (let i = 0; i < 32; i++) {
    const d = new Date(now - i * 86400_000).toISOString().slice(0, 10);
    const val = parseInt(await env.HNUST_TELEMETRY.get(`daily_count:${d}`), 10) || 0;
    if (val > 0) dailyCounts[d] = val;
  }

  const backup = {
    version: "1.0",
    export_time: new Date().toISOString(),
    total_devices: parseInt(await env.HNUST_TELEMETRY.get("total_devices"), 10) || 0,
    total_submissions: parseInt(await env.HNUST_TELEMETRY.get("total_submissions"), 10) || 0,
    daily_counts: dailyCounts,
    exam_stats_list: JSON.parse(await env.HNUST_TELEMETRY.get("exam_stats_list") || "[]"),
    recent_scores: JSON.parse(await env.HNUST_TELEMETRY.get("recent_scores") || "[]"),
  };

  return new Response(JSON.stringify(backup, null, 2), {
    headers: {
      "Content-Type": "application/json",
      "Content-Disposition": 'attachment; filename="hnust_telemetry_backup.json"',
    },
  });
}

// ── POST /api/import?token=xxx ──────────────────────────
async function handleImport(request, env) {
  const url = new URL(request.url);
  const token = url.searchParams.get("token");
  if (!token || token !== env.ADMIN_TOKEN) {
    return json({ error: "unauthorized" }, 401);
  }

  let body;
  try {
    body = await request.json();
  } catch {
    return json({ error: "无效的 JSON" }, 400);
  }

  if (body.version !== "1.0") {
    return json({ error: "版本不兼容" }, 400);
  }

  // 覆盖写入计数器
  if (body.total_devices != null) await env.HNUST_TELEMETRY.put("total_devices", String(body.total_devices));
  if (body.total_submissions != null) await env.HNUST_TELEMETRY.put("total_submissions", String(body.total_submissions));

  // 覆盖写入每日计数
  if (body.daily_counts && typeof body.daily_counts === "object") {
    for (const [date, count] of Object.entries(body.daily_counts)) {
      await env.HNUST_TELEMETRY.put(`daily_count:${date}`, String(count));
    }
  }

  // 覆盖写入试卷统计 & 最近提交
  if (body.exam_stats_list) await env.HNUST_TELEMETRY.put("exam_stats_list", JSON.stringify(body.exam_stats_list));
  if (body.recent_scores) await env.HNUST_TELEMETRY.put("recent_scores", JSON.stringify(body.recent_scores));

  return json({ status: "ok" });
}

// ── POST /api/heartbeat ─────────────────────────────────
async function handleHeartbeat(request, env) {
  const ip = getClientIP(request);
  if (isRateLimited(ip)) return json({ error: "rate limited" }, 429);

  const body = await readJson(request);
  const { device_id, version, os } = body;
  if (!device_id) return json({ error: "missing device_id" }, 400);

  const now = Date.now();
  const date = todayKey();

  // 检查设备是否首次出现
  const existing = await env.HNUST_TELEMETRY.get(`device:${device_id}`);
  if (!existing) {
    await incr(env, "total_devices");
  }

  // 记录日活（带 TTL 32 天自动过期）。同一天重复启动不再重复写 KV，避免免费额度被心跳耗尽。
  const dailyKey = `daily_active:${date}:${device_id}`;
  const alreadyActive = await env.HNUST_TELEMETRY.get(dailyKey);
  if (!alreadyActive) {
    await env.HNUST_TELEMETRY.put(dailyKey, "1", {
      expirationTtl: 86400 * 32,
    });
  }

  // 仅设备当日首次活跃时才增加计数（避免重复心跳膨胀）
  if (!alreadyActive) {
    await incr(env, `daily_count:${date}`);
  }

  // 只在首次设备或当天首次活跃时更新设备记录，减少无统计价值的写入。
  if (!existing || !alreadyActive) {
    await env.HNUST_TELEMETRY.put(`device:${device_id}`, JSON.stringify({
      last_seen: now, version: version || "", os: os || "",
    }));
  }

  return json({ status: "ok" });
}

// ── POST /api/submit-score ──────────────────────────────
async function handleSubmitScore(request, env) {
  const ip = getClientIP(request);
  if (isRateLimited(ip)) return json({ error: "rate limited" }, 429);

  const body = await readJson(request);
  const { device_id, exam_name, score_pct, duration_seconds, question_types, version, os } = body;
  if (!device_id || !exam_name) return json({ error: "missing fields" }, 400);

  const now = Date.now();
  const date = todayKey();

  // 增加累计提交次数
  await incr(env, "total_submissions");

  // 确保设备被记录 & 日活
  const existing = await env.HNUST_TELEMETRY.get(`device:${device_id}`);
  if (!existing) {
    await incr(env, "total_devices");
  }
  await env.HNUST_TELEMETRY.put(`device:${device_id}`, JSON.stringify({
    last_seen: now, version: version || "", os: os || "",
  }));

  const dailyKey = `daily_active:${date}:${device_id}`;
  const alreadyActive = await env.HNUST_TELEMETRY.get(dailyKey);
  if (!alreadyActive) {
    await env.HNUST_TELEMETRY.put(dailyKey, "1", {
      expirationTtl: 86400 * 32,
    });
  }

  // 仅设备当日首次活跃时才增加计数
  if (!alreadyActive) {
    await incr(env, `daily_count:${date}`);
  }

  // 更新试卷统计（读取 → 修改 → 写回）
  const statsRaw = await env.HNUST_TELEMETRY.get("exam_stats_list");
  const statsList = statsRaw ? JSON.parse(statsRaw) : [];
  let exam = statsList.find(e => e.name === exam_name);
  if (!exam) {
    exam = { name: exam_name, count: 0, total_score: 0, score_count: 0, total_duration: 0, duration_count: 0 };
    statsList.push(exam);
  }
  exam.count++;
  if (score_pct != null) { exam.total_score += score_pct; exam.score_count++; }
  if (duration_seconds != null) { exam.total_duration += duration_seconds; exam.duration_count++; }
  await env.HNUST_TELEMETRY.put("exam_stats_list", JSON.stringify(statsList));

  // 更新最近提交记录（unshift，保留最新 20 条）
  const recentRaw = await env.HNUST_TELEMETRY.get("recent_scores");
  const recent = recentRaw ? JSON.parse(recentRaw) : [];
  recent.unshift({
    device_id: device_id.slice(0, 8),
    exam_name,
    score_pct: score_pct ?? null,
    duration_seconds: duration_seconds ?? null,
    version: version || "",
    created_at: new Date(now).toISOString(),
  });
  if (recent.length > 20) recent.length = 20;
  await env.HNUST_TELEMETRY.put("recent_scores", JSON.stringify(recent));

  return json({ status: "ok" });
}

// ── GET /admin?token=xxx ────────────────────────────────
async function handleAdmin(request, env) {
  const url = new URL(request.url);
  const token = url.searchParams.get("token");
  if (!token || token !== env.ADMIN_TOKEN) {
    return new Response("Unauthorized", { status: 401 });
  }

  return new Response(buildDashboardHTML(token), {
    headers: { "Content-Type": "text/html; charset=utf-8" },
  });
}

// ── 仪表盘 HTML ─────────────────────────────────────────
function buildDashboardHTML(token) {
  return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>HNUST 考试平台 - 数据看板</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:system-ui,-apple-system,sans-serif;background:#f0f2f5;color:#1a1a2e;padding:20px}
h1{text-align:center;margin-bottom:8px;font-size:1.5rem}
.sub{text-align:center;color:#888;font-size:.85rem;margin-bottom:20px}
.toolbar{text-align:center;margin-bottom:20px}
.btn{display:inline-block;padding:10px 28px;background:#4361ee;color:#fff;border:none;border-radius:6px;font-size:.9rem;cursor:pointer}
.btn:hover{background:#3a56d4}
.btn:disabled{opacity:.6;cursor:wait}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:16px;margin-bottom:24px}
.card{background:#fff;border-radius:10px;padding:20px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.card .num{font-size:2rem;font-weight:700;color:#4361ee}
.card .label{color:#888;font-size:.85rem;margin-top:4px}
.charts{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:24px}
.chart-box{background:#fff;border-radius:10px;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.chart-box h3{margin-bottom:12px;font-size:.95rem}
table{width:100%;border-collapse:collapse;font-size:.85rem}
th,td{padding:8px 12px;text-align:left;border-bottom:1px solid #eee}
th{background:#f8f9fa;font-weight:600;color:#555}
tr:hover{background:#f0f4ff}
.section{background:#fff;border-radius:10px;padding:16px;margin-bottom:24px;box-shadow:0 1px 3px rgba(0,0,0,.1)}
.section h3{margin-bottom:12px;font-size:.95rem}
.empty{text-align:center;color:#aaa;padding:20px}
.msg{text-align:center;padding:40px;color:#888;font-size:.95rem}
@media(max-width:768px){.charts{grid-template-columns:1fr}}
</style></head><body>
<h1>HNUST 考试仿真平台 - 数据看板</h1>
<p class="sub">数据每 30 天自动清理，点击下方按钮手动刷新</p>
<div class="toolbar">
<button class="btn" id="refreshBtn" onclick="loadData()">刷新数据</button>
<button class="btn" onclick="exportData()">导出数据</button>
<button class="btn" onclick="importData()">导入数据</button>
<input type="file" id="importFile" accept=".json" style="display:none">
</div>
<div id="content"><p class="msg">正在加载数据…</p></div>
<script>
const TOKEN=${JSON.stringify(token)};
let chart7=null,chart30=null;

async function loadData(){
  const btn=document.getElementById('refreshBtn');
  btn.textContent='加载中…';btn.disabled=true;
  try{
    const r=await fetch('/api/stats?token='+TOKEN);
    if(!r.ok)throw new Error('HTTP '+r.status);
    const d=await r.json();
    render(d);
    btn.textContent='刷新数据';
  }catch(e){
    btn.textContent='刷新数据';
    document.getElementById('content').innerHTML='<p class="msg">加载失败: '+esc(String(e))+'，请重试</p>';
  }
  btn.disabled=false;
}

// 导出：直接触发浏览器下载
function exportData(){
  window.location.href='/api/export?token='+TOKEN;
}

// 导入：文件选择 → 上传 → 恢复
function importData(){
  document.getElementById('importFile').click();
}
document.getElementById('importFile').addEventListener('change',async function(e){
  const file=e.target.files[0];
  if(!file)return;
  try{
    const text=await file.text();
    const r=await fetch('/api/import?token='+TOKEN,{method:'POST',headers:{'Content-Type':'application/json'},body:text});
    const d=await r.json();
    if(d.error){alert('导入失败: '+d.error);return;}
    alert('数据恢复成功，页面即将刷新');
    loadData();
  }catch(err){
    alert('导入失败: '+String(err));
  }
  e.target.value='';
});

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function fmtDur(s){return s!=null?Math.round(s/60)+'分钟':'—';}
function fmtScore(s){return s!=null?s.toFixed(1)+'%':'—';}
function fmtTime(t){return(t||'').slice(0,16).replace('T',' ');}

function render(d){
  const examRows=d.examStats.map(e=>'<tr><td>'+esc(e.name)+'</td><td>'+e.count+'</td><td>'+(e.avg_score!=null?e.avg_score+'%':'—')+'</td><td>'+(e.avg_duration!=null?e.avg_duration+'分钟':'—')+'</td></tr>').join('\\n');
  const scoreRows=d.recentScores.map(s=>'<tr><td>'+fmtTime(s.created_at)+'</td><td>'+esc(s.device_id||'')+'…</td><td>'+esc(s.exam_name||'')+'</td><td>'+fmtScore(s.score_pct)+'</td><td>'+fmtDur(s.duration_seconds)+'</td><td>'+(s.version?esc(s.version):'—')+'</td></tr>').join('\\n');
  const now=new Date().toLocaleString('zh-CN');

  document.getElementById('content').innerHTML='\\n'+
  '<div class="cards">\\n'+
  '  <div class="card"><div class="num">'+d.todayDAU+'</div><div class="label">今日活跃设备</div></div>\\n'+
  '  <div class="card"><div class="num">'+d.totalDevices+'</div><div class="label">累计设备总数</div></div>\\n'+
  '  <div class="card"><div class="num">'+d.totalSubmissions+'</div><div class="label">累计提交次数</div></div>\\n'+
  '</div>\\n'+
  '<div class="charts">\\n'+
  '  <div class="chart-box"><h3>近 7 天活跃趋势</h3><canvas id="chart7"></canvas></div>\\n'+
  '  <div class="chart-box"><h3>近 30 天活跃趋势</h3><canvas id="chart30"></canvas></div>\\n'+
  '</div>\\n'+
  '<div class="section"><h3>试卷使用统计</h3>\\n'+
  '<table><thead><tr><th>试卷名称</th><th>提交次数</th><th>平均分</th><th>平均用时</th></tr></thead>\\n'+
  '<tbody>'+(examRows||'<tr><td colspan="4" class="empty">暂无数据</td></tr>')+'</tbody></table>\\n'+
  '</div>\\n'+
  '<div class="section"><h3>最近提交记录（'+d.recentScores.length+' 条）</h3>\\n'+
  '<table><thead><tr><th>时间</th><th>设备</th><th>试卷</th><th>成绩</th><th>用时</th><th>版本</th></tr></thead>\\n'+
  '<tbody>'+(scoreRows||'<tr><td colspan="6" class="empty">暂无数据</td></tr>')+'</tbody></table>\\n'+
  '</div>\\n'+
  '<p style="text-align:center;color:#aaa;font-size:.8rem;margin-top:8px">最后更新: '+now+'</p>';

  // 渲染图表
  if(chart7)chart7.destroy();
  if(chart30)chart30.destroy();
  const co={responsive:true,maintainAspectRatio:true,scales:{y:{beginAtZero:true,ticks:{stepSize:1}}}};
  const labels7=d.dailyCounts.slice(-7).map(c=>c.date.slice(5));
  const values7=d.dailyCounts.slice(-7).map(c=>c.count);
  const labels30=d.dailyCounts.map(c=>c.date.slice(5));
  const values30=d.dailyCounts.map(c=>c.count);
  chart7=new Chart(document.getElementById('chart7'),{type:'line',data:{labels:labels7,datasets:[{label:'活跃设备',data:values7,borderColor:'#4361ee',backgroundColor:'rgba(67,97,238,.1)',fill:true,tension:.3}]},options:co});
  chart30=new Chart(document.getElementById('chart30'),{type:'line',data:{labels:labels30,datasets:[{label:'活跃设备',data:values30,borderColor:'#4361ee',backgroundColor:'rgba(67,97,238,.1)',fill:true,tension:.3}]},options:co});
}

loadData();
</script></body></html>`;
}
