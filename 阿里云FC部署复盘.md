# 阿里云 FC 部署复盘

日期：2026-05-13

## 结论

阿里云函数计算 FC 可以运行本项目的 FastAPI 后端，但不适合作为“直接给用户打开的 Web 页面入口”，至少不能直接使用 FC 默认的 `fcapp.run` 公网域名。

本次验证结果：

- 函数代码可以成功部署并启动。
- HTTP 触发器需要改为“无需认证”，否则外部访问会返回缺少签名请求头。
- FC 默认域名会对 HTML 响应加 `Content-Disposition: attachment`，Chrome / Atlas 会强制下载 HTML 文件。
- 下载后的 HTML 再本地打开会丢失 `/static/style.css` 等静态资源，UI 和真实线上页面不一致。
- 因此，FC 默认公网域名不能作为本项目的 demo 链接。

## 实际部署过程

创建函数：

- 地域：华北 2 北京
- 函数名：`talent-ninebox-web`
- 运行环境：自定义运行时 Debian 10 / Python 3.10
- 启动命令：`/code/bootstrap`
- 监听端口：`9000`
- HTTP 触发器：`defaultTrigger`
- 公网地址：`https://talent-ebox-web-tlrvwjllcm.cn-beijing.fcapp.run`

环境变量：

```text
APP_ACCESS_PASSWORD=TalentNinebox@2026
APP_SESSION_SECRET=talent-ninebox-fc-secret-20260513
```

## 遇到的问题

### 1. HTTP 触发器默认签名认证

最初访问公网地址返回：

```json
{
  "Code": "MissingRequiredHeader",
  "Message": "required HTTP header Date was not specified"
}
```

原因是 HTTP 触发器使用了“签名认证”。这适合 API 调用，不适合普通浏览器直接访问。

处理方式：

- 进入 FC 函数触发器配置。
- 将认证方式从“签名认证”改为“无需认证”。

### 2. ZIP 依赖缺失

函数启动失败，日志显示：

```text
ModuleNotFoundError: No module named 'exceptiongroup'
```

原因是本地打包环境和 FC Python 3.10 运行环境的依赖解析不完全一致。

处理方式：

- 在 `requirements.txt` 显式增加：

```text
exceptiongroup>=1.2
```

- 使用 Linux Python 3.10 容器重新构建 ZIP。

### 3. macOS 直接打包不可靠

项目包含 `pydantic-core`、`uvloop` 等二进制依赖。macOS 本地构建出的 wheel 不适合直接上传到 Linux FC 环境。

处理方式：

```bash
docker run --rm --platform linux/amd64 \
  -v /Users/evafu/Projects/talent-ninebox-web:/workspace \
  -w /workspace \
  python:3.10-slim \
  bash -lc "apt-get update && apt-get install -y rsync zip && bash scripts/build_fc.sh"
```

生成文件：

```text
/Users/evafu/Projects/talent-ninebox-web/fc-deploy/talent-ninebox-fc.zip
```

### 4. 默认域名强制下载 HTML

即使函数启动成功，访问 FC 默认公网域名仍然会被浏览器下载 HTML，而不是渲染页面。

响应头中出现：

```text
Content-Disposition: attachment
```

这是 FC 默认域名策略，不是业务代码问题。

## 做过的代码修正

为兼容 FC，做过两类修正：

1. 去掉登录流程中对 303 跳转的强依赖，改为直接渲染页面，避免 FC 默认域名下 `ExternalRedirectForbidden`。
2. 将 FC 环境的临时文件目录切到 `/tmp/talent-ninebox-web`。

相关代码位置：

```text
talent_ninebox/web/app.py
```

## 最终判断

FC 适合继续作为后端执行环境，但不适合直接承载当前一体化 Web 页面。

如果未来继续走阿里云，推荐两条可行路线：

### 方案 A：FC 绑定自定义域名

前提：

- 需要已有可用域名。
- 大陆地域通常需要 ICP 备案或接入备案。

优点：

- 改造量较小。
- 当前 FastAPI 一体化应用可以继续使用。

缺点：

- 域名和备案会拖慢短期项目节奏。

### 方案 B：OSS 前端 + FC 后端 API

结构：

```text
浏览器 -> OSS 静态页面 -> FC API -> 返回 Excel
```

优点：

- 不依赖 FC 默认域名渲染 HTML。
- 更符合阿里云轻量部署方式。

缺点：

- 需要拆前后端。
- 要处理 CORS、登录态、跨域下载文件。
- 改造量明显大于 Railway。

## 建议

短期 demo 不继续投入 FC 默认域名。

当前项目更适合：

1. Vercel 继续作为备用演示链接。
2. Railway 作为更适合 Python Excel 工具的主 demo 环境。
3. 如果后续明确要求国内稳定访问，再做“OSS 前端 + FC 后端”或“FC + 已备案自定义域名”。
