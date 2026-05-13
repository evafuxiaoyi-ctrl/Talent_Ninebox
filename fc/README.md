# 阿里云函数计算 FC 部署说明

这是一套独立于 Vercel 的阿里云 FC Custom Runtime 部署方案。保留现有 Vercel 线上版本不变。

## 推荐方案：Custom Container

更推荐使用 FC Custom Container，而不是在本机直接打 zip。原因是本项目依赖 `openpyxl`、`fastapi`、`pydantic-core` 等包，在 macOS 上构建 zip 可能混入 macOS wheel，上传到 FC Linux 环境后会有兼容风险。

```text
函数计算 FC / Web 函数 / Custom Container
```

运行方式：

```text
容器启动 uvicorn
FastAPI 监听 0.0.0.0:9000
FC HTTP 触发器转发请求
```

## 文件说明

```text
fc/Dockerfile                 FC Custom Container 镜像
scripts/build_fc_container.sh  构建容器镜像
scripts/push_fc_acr.sh        推送镜像到阿里云 ACR
fc/bootstrap                  Custom Runtime zip 方案启动脚本
scripts/build_fc.sh            构建 Custom Runtime zip 包
fc-deploy/                    zip 构建输出目录，不提交
```

## 方案 A：Custom Container

本地构建镜像：

```bash
bash scripts/build_fc_container.sh talent-ninebox-fc:latest
```

构建脚本已固定使用 `linux/amd64`，适配函数计算 Custom Container 的运行架构。

本地测试：

```bash
docker run --rm -p 9000:9000 \
  -e APP_ACCESS_PASSWORD='your-password' \
  -e APP_SESSION_SECRET='your-session-secret' \
  talent-ninebox-fc:latest
```

访问：

```text
http://127.0.0.1:9000
```

推送到阿里云 ACR：

```bash
docker login registry.cn-hangzhou.aliyuncs.com
bash scripts/push_fc_acr.sh talent-ninebox-fc:latest registry.cn-hangzhou.aliyuncs.com/<namespace>/<repo>:latest
```

部署到阿里云时，需要：

1. 将镜像推送到阿里云容器镜像服务 ACR。
2. 在函数计算 FC 中创建 Custom Container Web 函数。
3. 镜像地址选择 ACR 中的镜像。
4. 配置监听端口 `9000`。
5. 配置 HTTP 触发器。

建议配置：

```text
函数类型：Web 函数
运行环境：Custom Container
监听端口：9000
内存：1024MB 或 2048MB
超时时间：60s 或 120s
```

环境变量：

```text
APP_ACCESS_PASSWORD=你的访问密码
APP_SESSION_SECRET=随机长字符串
```

## 方案 B：Custom Runtime zip 包

如果不想使用容器，也可以使用 zip 包。但必须在 Linux 环境中构建，不能直接使用 macOS 构建出的依赖包。

在 Linux 环境执行：

```bash
bash scripts/build_fc.sh
```

生成：

```text
fc-deploy/talent-ninebox-fc.zip
```

FC 配置：

```text
函数类型：Web 函数
运行环境：Custom Runtime / custom.debian11 或 custom.debian12
请求处理程序类型：HTTP
监听端口：9000
启动命令：/code/bootstrap
内存：1024MB 或 2048MB
超时时间：60s 或 120s
临时磁盘：默认即可
```

环境变量：

```text
APP_ACCESS_PASSWORD=你的访问密码
APP_SESSION_SECRET=随机长字符串
```

上传：

```text
fc-deploy/talent-ninebox-fc.zip
```

部署后创建 HTTP 触发器，并使用触发器地址访问。

## 注意事项

- 这是短期试用部署方案，不影响 Vercel。
- 结果文件仍保存在函数实例临时目录中，建议处理完成后立即下载。
- 如果出现上传失败，优先检查 FC 的请求体大小、函数超时时间和内存配置。
- 真实人才数据建议使用国内地域，例如华东、华北或华南。
- 免费额度、试用期和按量计费以阿里云控制台当前页面为准。
- 如果使用 zip 包方案，请务必在 Linux 中构建依赖，避免 macOS wheel 不兼容。

## 回滚

Vercel 版本不受影响，回滚只需停止或删除 FC 函数。
