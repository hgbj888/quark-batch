---
name: quark-batch
description: 夸克网盘批量转存+分享工具，基于 quarkpan 库实现，支持批量转存分享链接并自动生成分享链接；当用户需要批量处理夸克网盘资源转存分享、整理教育资源、打包分发学习资料时使用
metadata:
  openclaw:
    requires:
      env: ["QUARK_COOKIE"]
    primaryEnv: "QUARK_COOKIE"
---

# 夸克网盘批量转存+分享自动化工具

## 任务目标
- 本 Skill 用于：批量转存多个网盘链接到指定文件夹，并自动创建分享链接
- 能力包含：批量转存、自动分享、Excel 输出
- 触发条件：用户提供网盘链接（文本或Excel格式）并要求批量处理

## 依赖等级
- 等级：L3
- 说明：需要 Python + quarkpan 库，需配置夸克网盘 Cookie

## 前置准备

### 环境初始化
```bash
pip install quarkpan openpyxl pandas
```

### 配置网盘凭据
复制 `.env.example` 为 `.env`，填入以下信息：
- `QUARK_COOKIE`：夸克网盘 Cookie（从浏览器开发者工具获取）

**获取 Cookie 方法**：
1. 浏览器登录夸克网盘
2. 按 F12 打开开发者工具 → Network 标签
3. 刷新页面，找到任意请求
4. 复制 Request Headers 中的 `Cookie` 字段

## 操作步骤

### 步骤 1：输入数据准备

**方式一：文本格式**
用户提供网盘链接列表，每行一个链接：
```
https://pan.quark.cn/s/xxxxx
```

**方式二：Excel 格式**
用户提供 Excel 文件，包含以下列：
- `链接`：夸克网盘分享链接
- `名称`：资源名称（可选，留空则自动获取）

### 步骤 2：执行批量处理

调用主脚本处理：
```bash
python scripts/batch_share.py \
  --input <输入文件/文本> \
  --output outputs/tables/result.xlsx
```

**参数说明**：
- `--input`：输入文件路径或直接文本（每行一个链接）
- `--output`：输出 Excel 文件路径
- `--folder`：转存到的文件夹名称（可选，默认：批量转存）

### 步骤 3：输出结果

生成的 Excel 文件包含以下列：
- `资源名称`：文件或文件夹名称
- `网盘链接`：创建的新分享链接
- `状态`：转存和分享状态（成功/失败）
- `备注`：错误信息（如有）

## 使用示例

### 示例一：批量处理夸克链接

**场景描述**：用户有 10 个夸克网盘学习资料链接，需要批量转存并分享

**执行方式**：
```bash
python scripts/batch_share.py \
  --input "https://pan.quark.cn/s/xxx
https://pan.quark.cn/s/yyy
https://pan.quark.cn/s/zzz" \
  --output outputs/tables/学习资源.xlsx
```

**预期输出**：
- 生成 `outputs/tables/学习资源.xlsx`
- 包含 10 行数据，每行一个资源的分享链接

### 示例二：从 Excel 读取并处理

**场景描述**：用户提供 `resources.xlsx`，包含 50 个网盘链接

**执行方式**：
```bash
python scripts/batch_share.py \
  --input resources.xlsx \
  --output outputs/tables/result.xlsx
```

**预期输出**：
- 读取 `resources.xlsx` 中的所有链接
- 批量转存并分享
- 输出新分享链接的结果表