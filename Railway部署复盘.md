# Railway 部署复盘

日期：2026-05-13

## 结论

Railway 比 Vercel Serverless 和阿里云 FC 默认域名更适合当前这个短期 demo。

当前 Railway 版本已部署成功：

```text
https://talent-ninebox-web-production.up.railway.app
```

访问密码：

```text
TalentNinebox@2026
```

验证结果：

- 登录页正常返回。
- 密码登录成功。
- 首页正常返回。
- 部署状态为 `SUCCESS`。
- 不会像阿里云 FC 默认域名那样强制下载 HTML。

## 为什么 Railway 更适合当前项目

本项目是：

- Python FastAPI Web 应用
- openpyxl 处理 Excel
- 上传 zip / xlsx
- 处理后返回 Excel 文件

Railway 是常驻 Web 服务，更接近一台轻量服务器。相比 Vercel Serverless，它更适合这种文件上传和 Excel 处理任务。

但需要明确：

- Railway 不等于中国大陆稳定访问。
- Railway 仍是海外平台，本次实际 edge 显示为新加坡区域。
- 它解决的是“应用运行形态更合适”，不是“国内网络一定稳定”。

## 本次部署步骤

### 1. 增加 Railway 配置

新增文件：

```text
railway.json
```

内容核心是指定启动命令：

```json
{
  "build": {
    "builder": "NIXPACKS"
  },
  "deploy": {
    "startCommand": "python -m uvicorn talent_ninebox.web.app:app --host 0.0.0.0 --port $PORT",
    "restartPolicyType": "ON_FAILURE",
    "restartPolicyMaxRetries": 3
  }
}
```

### 2. 适配 Railway 临时目录

在 `talent_ninebox/web/app.py` 中加入 Railway 环境识别：

```python
IS_SERVERLESS = bool(
    os.environ.get("VERCEL")
    or os.environ.get("FC_FUNCTION_NAME")
    or os.environ.get("RAILWAY_ENVIRONMENT")
)
```

这样 Railway 环境下结果文件会写入：

```text
/tmp/talent-ninebox-web
```

避免写部署目录。

### 3. 安装并登录 Railway CLI

安装：

```bash
npm install -g @railway/cli
```

登录：

```bash
railway login
```

登录账号：

```text
evafuxiaoyi@gmail.com
```

### 4. 创建项目和服务

项目：

```text
talent-ninebox-web
```

项目 ID：

```text
5c3b704f-5b9b-447c-8168-4129ed23e72d
```

服务：

```text
talent-ninebox-web
```

服务 ID：

```text
5698ad32-a8ba-421b-8031-349843d1f081
```

### 5. 设置环境变量

```bash
railway variable set \
  APP_ACCESS_PASSWORD='TalentNinebox@2026' \
  APP_SESSION_SECRET='talent-ninebox-railway-20260513-session-secret' \
  --service talent-ninebox-web \
  --skip-deploys
```

### 6. 部署

直接从完整项目目录上传时失败，原因是目录过大，包含：

- `.venv`
- `fc-deploy`
- 字体文件
- 缓存目录

第一次直接上传报错：

```text
operation timed out
```

最终采用最小部署目录：

```text
/tmp/talent-ninebox-railway
```

只包含：

```text
talent_ninebox/
requirements.txt
railway.json
```

并排除：

- `__pycache__/`
- `*.pyc`
- `*.bak`
- `static/fonts/`

部署命令：

```bash
railway up /tmp/talent-ninebox-railway \
  --path-as-root \
  --service talent-ninebox-web \
  --message 'Restore working talent ninebox demo' \
  --detach \
  --json
```

最终成功 deployment：

```text
8e62c1eb-7fad-476f-9406-776db08181db
```

状态：

```text
SUCCESS
```

## 遇到的问题

### 1. 当前项目目录不是独立 Git 仓库

`/Users/evafu/Projects/talent-ninebox-web` 没有自己的 `.git`，父目录 `/Users/evafu` 是 Git 仓库。

这会导致：

- `git status` 扫到大量用户目录文件。
- 不适合直接作为 GitHub 自动部署源。

短期处理方式：

- 使用 Railway CLI 从最小目录上传部署。

长期建议：

- 将项目整理成独立 Git 仓库。
- 再接 GitHub -> Railway 自动部署。

### 2. 完整项目上传超时

完整目录约 186MB，其中：

```text
.venv       约 59MB
fc-deploy   约 99MB
```

即使加入 `.railwayignore`，本地 CLI 上传仍然超时。

处理方式：

- 构建 `/tmp/talent-ninebox-railway` 最小上传目录。

### 3. 字体文件导致上传不稳定

本地字体目录约 28MB：

```text
talent_ninebox/web/static/fonts
```

带字体上传 Railway 时仍然超时。

处理方式：

- 当前 Railway 版本不上传字体文件。
- 浏览器回退系统字体。

实际观感：

- 用户反馈字体不影响观感。
- 功能和布局正常。

后续如果要完全一致：

- 走 GitHub 部署，避免 CLI 大文件上传超时。
- 或将字体放 CDN。
- 或压缩 / 子集化字体。

## 当前 Railway 版本状态

访问地址：

```text
https://talent-ninebox-web-production.up.railway.app
```

密码：

```text
TalentNinebox@2026
```

验证命令：

```bash
curl -i -L https://talent-ninebox-web-production.up.railway.app
```

登录验证已通过：

```bash
curl -i -X POST \
  -d 'password=TalentNinebox%402026' \
  https://talent-ninebox-web-production.up.railway.app/login
```

## 后续建议

短期：

- Railway 作为主 demo 链接。
- Vercel 作为备用链接。
- 阿里云 FC 默认域名暂停使用。

中期：

- 将项目整理成独立 Git 仓库。
- 用 GitHub 连接 Railway 自动部署。
- 保留 `railway.json`。

长期：

- 如果需要国内稳定访问，再考虑：
  - 阿里云 OSS 前端 + FC 后端 API
  - 或阿里云 FC 绑定已备案自定义域名

## 操作备忘

查看部署：

```bash
railway deployment list --service talent-ninebox-web --json
```

查看日志：

```bash
railway logs --service talent-ninebox-web --lines 50
```

重新部署最小目录：

```bash
rm -rf /tmp/talent-ninebox-railway
mkdir -p /tmp/talent-ninebox-railway
rsync -a \
  --exclude 'static/fonts/' \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  --exclude '*.bak' \
  talent_ninebox requirements.txt railway.json \
  /tmp/talent-ninebox-railway/

railway up /tmp/talent-ninebox-railway \
  --path-as-root \
  --service talent-ninebox-web \
  --message 'Deploy talent ninebox demo' \
  --detach \
  --json
```
