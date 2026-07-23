# App Store Review Analyzer — 运行说明

## 功能概述

输入美区 App Store 应用链接，自动抓取评论 → DeepSeek AI 分析痛点 → 生成 PRD → 生成测试用例，并支持导出 **真正的 Word(.docx) 文档**。

---

## 一、环境要求

- Python 3.9+
- 网络可访问美区 App Store（抓取评论）
- 网络可访问 DeepSeek API（AI 分析）
- 一个 DeepSeek API Key（https://platform.deepseek.com 获取）

---

## 二、在 macOS（Warp 终端）上安装与启动

### 1. 安装依赖（一条命令搞定）

```bash
pip install -r requirements.txt
```

> ⚠️ 如果你用的是 Python 3 自带的 pip（系统保护），可能提示权限错误，
> 建议先创建虚拟环境，或改用下面任一方式：
> ```bash
> # 方式 A：用 python3 -m pip（推荐）
> python3 -m pip install -r requirements.txt
>
> # 方式 B：用 Homebrew 装的 Python
> pip3 install -r requirements.txt
> ```
>
> 如果提示 `zsh: command not found: pip`，说明只有 `pip3`：
> ```bash
> pip3 install -r requirements.txt
> ```

### 2. 启动应用

```bash
streamlit run app.py
```

启动后 Warp 会自动打印一个本地地址，**浏览器打开 `http://localhost:8501`** 即可使用。

### 3. 使用步骤

1. **填 API Key**：左侧边栏填入 DeepSeek API Key
2. **贴链接**：把美区 App Store 链接粘进输入框（支持多种格式，见下方「验证 Bug 1」）
3. **写目标（可选）**：如「关注订阅问题」「关注闪退卡顿」
4. **调数量**：侧边栏滑块设置抓取条数（20–500）
5. **点「开始分析」**：等待抓取 + AI 分析（约 30–90 秒）
6. **看结果**：页面展示核心问题分析 / PRD / 测试用例
7. **下载**：可下载 Markdown、JSON，以及 **Word 报告（PRD + 测试用例 .docx）**

---

## 三、验证两个 Bug 已被修复

### ✅ Bug 1：链接解析（在界面「侧边栏 → 🧪 链接解析自测 → 运行自测」点一下即可）

你也完全可以手动在输入框里粘贴下面这些链接，观察界面是否正确显示 App ID、错误时是否给出中文报错：

| 验证场景 | 测试链接 | 预期结果 |
|---------|---------|---------|
| 标准格式（应用名+id） | `https://apps.apple.com/us/app/tiktok/id835599320` | 显示 App ID：**835599320** |
| 无应用名（直接 id） | `https://apps.apple.com/us/app/id123456789` | 显示 App ID：**123456789** |
| iTunes 老链接+参数 | `https://itunes.apple.com/us/app/spotify/id324684580?mt=8` | 显示 App ID：**324684580** |
| 末尾带斜杠 | `https://apps.apple.com/us/app/some-app/id987654321/` | 显示 App ID：**987654321** |
| 带其他查询参数 | `https://apps.apple.com/us/app/app-name/id111222333?l=zh-Hans-CN` | 显示 App ID：**111222333** |
| ❌ 非美区（应报错） | `https://apps.apple.com/cn/app/app-name/id445678123` | 红字报错：检测到国家码为「cn」…请使用 /us/ 链接 |
| ❌ 无数字 ID（应报错） | `https://example.com/abcdefg` | 红字报错：未识别到有效的国家码… |

> 解析逻辑：用正则 `r'id(\d+)'` 从整条链接里抠纯数字 ID，不依赖分段位置；
> 国家码不是 `us` 或找不到 ID 时，用 `st.error()` 给出明确中文提示，绝不静默失败。

### ✅ Bug 2：Word 文档（.docx）能被 Word 正常打开

1. 跑一次完整分析（或点「开始分析」用已有缓存）
2. 在下载区点 **「📄 Word 报告（PRD+测试用例）」**
3. 下载后用 Word / WPS / Pages 打开，应能看到：
   - 带样式的标题（核心问题分析、PRD、测试用例）
   - **真正的表格**（需求表、测试用例表，边框清晰可编辑）
4. 文件是 python-docx 生成的真·Word 格式，不再是「伪 .docx 文本文件」
5. 生成后程序会**自动自检**（用 python-docx 重新读回一遍），自检不过会直接红字报错，不让你拿到坏文件

> 关键修复点：
> - 用 `from docx import Document` 生成，**绝不**用 `open().write(字符串)` 再改后缀
> - 下载时传入 `io.BytesIO` 里的二进制数据，`mime` 填
>   `application/vnd.openxmlformats-officedocument.wordprocessingml.document`

---

## 四、缓存机制

- 抓取后的评论自动保存为 `reviews_cache.json`
- 网络不通时自动使用缓存数据进行分析
- 侧边栏可查看缓存状态或手动清除
- 想重新抓取，把输入框里的链接换掉再点「开始分析」即可

---

## 五、输出内容

| 部分 | 内容 |
|------|------|
| **核心问题分析** | 问题归类、严重程度、引用评论原文 |
| **产品需求文档 (PRD)** | P0/P1/P2 优先级、V1.0/V2.0 版本拆分、验收标准 |
| **测试用例** | 可追溯到 需求编号(REQ-ID) 和 评论ID 的完整用例，Word 表格呈现 |

---

## 六、关于 app-store-scraper（可选）

本应用**主抓取路径用 Apple 官方 RSS Feed**，稳定且无需额外库。
`app-store-scraper` 仅作为「备用」抓取方式，已改为**按需懒加载**：

```bash
# 想启用备用抓取时单独安装（--no-deps 跳过它上游对 requests 的死锁约束）
pip install --no-deps app-store-scraper
```

> 不装它也完全不影响主流程；装了它且网络异常时，RSS 失败会自动尝试它。

---

## 六·补、评论抓取优化：XML + JSON 双通道

### 为什么从「XML 优先 / JSON 兜底」升级为「双通道取优」

最初按需求实现了 Apple RSS 的 **XML 格式**抓取（`/xml`，用标准库 `xml.etree.ElementTree` 解析），
并保留 JSON 作为兜底。但实测多个 App 后发现一个**反直觉的事实**：

| App | JSON 条数 | XML 条数 | 说明 |
|-----|----------|----------|------|
| Netflix | 49 | 50 | 两种都有 |
| TikTok | 50 | 0（或间歇 100） | **只有 JSON 稳定有数据** |
| Spotify | 0 | 50 | **只有 XML 有数据** |
| Instagram / YouTube / WhatsApp | 0 | 0 | 两种都空（Apple 接口限制） |

> 结论：**Apple 对「某个 App 能用哪种格式返回数据」没有规律**，且同一 App 同一格式也会间歇性返回空
> （短时内密集请求会触发限流，表现为 0 条）。因此「XML 优先」并不比「JSON 优先」更稳。

### 现在的做法（`fetch_reviews_via_rss`）

```python
xml_reviews  = _fetch_rss_xml(app_id, country, max_count)   # XML 通道
json_reviews = _fetch_rss_json(app_id, country, max_count)  # JSON 通道
# 两边都为空 -> 返回 []，由界面提示「手动粘贴评论」
# 否则取「数量多」的那一份；数量相同时优先 XML
```

- 两条通道**都抓**，最大化拿到真实评论的概率
- 只保留评分在 **1~5** 之间的条目，自动过滤 Apple 把 App 自身信息混在第 1 条的情况
- 界面「数据来源」会明确标出实际用的是 **XML** 还是 **JSON**，方便你验证接口是否生效
- 两条通道都为 0 时，**不崩溃**，给出中文提示 + 手动粘贴评论的替代入口

### 终端对比两种格式（你自己也能验证）

```bash
python3 -c "import app; app.compare_json_vs_xml('835599320')"
```

会在终端打印：

```
App ID 835599320 抓取对比：
  JSON 方式抓到: 50 条
  XML  方式抓到: 0 条
  => JSON 比 XML 多 50 条
```


---

## 七、零基础补充：python-docx 入门

代码里提供了 `docx_beginner_example()` 函数，逐行中文注释。
在终端运行即可生成一个真实 Word 文档：

```bash
python3 -c "import app; app.docx_beginner_example()"
```

界面「📘 python-docx 零基础入门示例」折叠区也能看到这份逐行注释的代码。
