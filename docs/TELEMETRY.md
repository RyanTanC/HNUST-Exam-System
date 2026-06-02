# 匿名使用统计部署指南

本项目使用 Cloudflare Workers + KV 实现完全匿名的使用数据统计，免费且无需服务器。

## 数据收集说明

- **完全匿名**：不收集姓名、学号等个人信息，仅使用随机生成的 16 字符设备 ID
- **收集内容**：设备 ID（匿名）、软件版本、操作系统、试卷名称、得分百分比、用时
- **静默上报**：所有数据上报在后台线程执行，失败不影响正常使用

## 部署步骤

### 1. 注册 Cloudflare 账号

访问 https://www.cloudflare.com 注册免费账号。

### 2. 安装 Wrangler CLI

需要先安装 Node.js，然后执行：

```bash
npm install -g wrangler
```

### 3. 登录 Cloudflare

```bash
wrangler login
```

浏览器会打开授权页面，点击授权即可。

### 4. 创建 KV 存储

```bash
wrangler kv:namespace create "HNUST_TELEMETRY"
```

命令会输出类似：

```
🌀 Creating namespace (id: "abc123def456...")
✨ Success!
Add the following to your configuration file in your kv_namespaces array:
{ binding = "HNUST_TELEMETRY", id = "abc123def456..." }
```

记下这个 `id` 值。

### 5. 配置 wrangler.toml

编辑 `cloudflare-worker/wrangler.toml`，把上一步的 `id` 填入：

```toml
kv_namespaces = [
  { binding = "HNUST_TELEMETRY", id = "abc123def456..." }
]
```

### 6. 设置管理员访问码

```bash
cd cloudflare-worker
wrangler secret put ADMIN_TOKEN
```

输入你想用的管理看板访问码（例如一串随机密码），回车确认。

### 7. 部署

```bash
wrangler deploy
```

部署成功后会显示 Worker 的访问地址，类似：

```
Published hnust-exam-telemetry (x.xx sec)
  https://hnust-exam-telemetry.xxx.workers.dev
```

### 8. 配置客户端

编辑 `hnust_exam/utils/constants.py`，将上一步获得的地址填入：

```python
TELEMETRY_BASE_URL = "https://hnust-exam-telemetry.xxx.workers.dev"
```

### 9. 重新打包发布

使用 PyInstaller 重新打包，分发给学生使用。

## 查看数据看板

浏览器访问：

```
https://你的worker地址/admin?token=你设置的ADMIN_TOKEN
```

看板每 30 秒自动刷新，包含：
- 今日活跃设备数、累计设备数、累计提交次数
- 近 7 天 / 30 天活跃趋势图
- 各试卷使用次数、平均分、平均用时
- 最近 50 条提交记录

## 注意事项

- Cloudflare Workers 免费版每天 10 万次请求，对于学生使用量完全足够
- KV 存储免费版每天 10 万次读取、1000 次写入
- 日活数据设置 32 天自动过期，不会无限增长
- Worker 内置了每分钟 30 次的 IP 限流，防止恶意刷请求
