# App Store Review Analyzer — 运行说明

## 功能概述

输入美区 App Store 应用链接，自动抓取评论 → DeepSeek AI 分析痛点 → 生成 PRD → 生成测试用例。

## 环境要求

- Python 3.9+
- 网络可访问美区 App Store（抓取评论）
- 网络可访问 DeepSeek API（AI 分析）

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动应用

```bash
streamlit run app.py
```

启动后浏览器会自动打开 `http://localhost:8501`。

### 3. 使用步骤

1. **配置 API Key**：在左侧边栏填入 DeepSeek API Key（在 https://platform.deepseek.com 获取）
2. **输入 App Store 链接**：粘贴美区 App Store 链接，例如：
   - `https://apps.apple.com/us/app/tiktok/id835599320`
   - `https://apps.apple.com/us/app/id835599320`
3. **输入分析目标**（可选）：如「关注订阅问题」「关注闪退和卡顿」等
4. **调整抓取数量**：侧边栏滑块可设置 20-500 条
5. **点击「开始分析」**：等待抓取 + AI 分析完成（约 30-90 秒）
6. **查看结果**：页面展示核心问题分析、PRD、测试用例
7. **下载报告**：可下载 Markdown 格式分析报告或 JSON 完整数据

## 缓存机制

- 抓取后的评论自动保存为 `reviews_cache.json`
- 网络不通时自动使用缓存数据进行分析
- 侧边栏可查看缓存状态或手动清除

## 输出内容

分析报告包含三个部分：

| 部分 | 内容 |
|------|------|
| **核心问题分析** | 用户抱怨的问题归类、严重程度、引用评论原文 |
| **产品需求文档 (PRD)** | P0/P1/P2 优先级、V1.0/V2.0 版本拆分、验收标准 |
| **测试用例** | 可追溯到 REQ-ID 和原始评论的完整测试用例 |

## 注意事项

- DeepSeek API 调用需要有效的 API Key，费用按 token 计费
- 每次分析约消耗 2000-8000 token（取决于评论数量）
- 评论抓取依赖 app-store-scraper 库，美区 App Store 偶尔可能限流
- 如果抓取失败，系统会自动尝试使用缓存数据
