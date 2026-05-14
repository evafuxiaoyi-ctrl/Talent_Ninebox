# Talent Ninebox Web

人才九宫格：内部人才盘点 Excel 整合与九宫格生成工具。

## 功能

- 访问密码保护
- 上传 zip，递归识别 `.xlsx`
- 合并统一模板的人才盘点表
- 保留字段顺序、列宽、基础样式和公式
- 生成可编辑 Excel 九宫格
- 输出处理摘要、异常报告、文件来源映射

## 快速启动

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export APP_ACCESS_PASSWORD='change-me'
uvicorn talent_ninebox.web.app:app --reload --host 127.0.0.1 --port 8000
```

访问：http://127.0.0.1:8000

## 命令行处理

```bash
python -m talent_ninebox.cli input.zip ./output
```

## 安全说明

- 密码通过 `APP_ACCESS_PASSWORD` 环境变量配置。
- 仅支持 `.zip` 输入。
- zip 解压做路径穿越检查。
- 临时文件默认保留 1 小时。
- 不上传 Excel 内容到外部服务。
