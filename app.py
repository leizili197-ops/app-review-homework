#!/usr/bin/env python3
"""App Store Review Analyzer — 美区 App Store 评论抓取 + DeepSeek AI 分析 + PRD/测试用例生成"""

import streamlit as st
import json
import re
import os
import time
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Optional

import requests
from app_store_scraper import AppStore

# ============================================================================
# 页面配置
# ============================================================================
st.set_page_config(
    page_title="App Review Analyzer",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================================
# 常量
# ============================================================================
CACHE_FILE = Path(__file__).parent / "reviews_cache.json"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"

# ============================================================================
# 工具函数
# ============================================================================

def extract_app_id(url: str) -> Optional[str]:
    """从美区 App Store 链接中提取 App ID 和名称"""
    # 匹配多种链接格式
    patterns = [
        r"/id(\d+)",                          # /id123456789
        r"/app/[^/]+/id(\d+)",               # /app/xxx/id123456789
        r"apps\.apple\.com/.*?/id(\d+)",      # 完整 URL
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def extract_app_name(url: str) -> str:
    """从 URL 中提取 App 名称"""
    m = re.search(r"/app/([^/]+)/", url)
    if m:
        return m.group(1).replace("-", " ").title()
    return "Unknown App"


def fetch_reviews_via_rss(app_id: str, max_count: int = 200) -> list:
    """通过 Apple RSS Feed 抓取评论（比 app-store-scraper 更稳定）"""
    all_reviews = []
    seen_ids = set()

    for page in range(1, 11):  # 最多 10 页
        if len(all_reviews) >= max_count:
            break

        if page == 1:
            url = f"https://itunes.apple.com/us/rss/customerreviews/id={app_id}/sortBy=mostRecent/json"
        else:
            url = f"https://itunes.apple.com/us/rss/customerreviews/page={page}/id={app_id}/sortby=mostrecent/json"

        try:
            resp = requests.get(url, timeout=30)
            if resp.status_code != 200:
                break

            data = resp.json()
            entries = data.get("feed", {}).get("entry", [])

            if not entries or len(entries) <= 1:
                break

            for entry in entries:
                rid = entry.get("id", {}).get("label", "")
                if rid in seen_ids:
                    continue
                seen_ids.add(rid)

                all_reviews.append({
                    "id": hashlib.md5(
                        f"{entry.get('author', {}).get('name', {}).get('label', '')}"
                        f"{entry.get('updated', {}).get('label', '')}"
                        f"{entry.get('content', {}).get('label', '')}".encode()
                    ).hexdigest()[:8],
                    "userName": entry.get("author", {}).get("name", {}).get("label", "Anonymous"),
                    "title": entry.get("title", {}).get("label", ""),
                    "review": entry.get("content", {}).get("label", ""),
                    "rating": int(entry.get("im:rating", {}).get("label", "0")),
                    "date": entry.get("updated", {}).get("label", ""),
                    "version": entry.get("im:version", {}).get("label", ""),
                })

            if len(entries) < 50:  # 最后一页不足 50 条
                break

            time.sleep(0.3)  # 礼貌延迟

        except Exception:
            break

    return all_reviews


def load_cache() -> Optional[list]:
    """加载缓存的评论数据"""
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return None
    return None


def save_cache(reviews: list, app_name: str, app_id: str):
    """保存评论到缓存文件"""
    cache_data = {
        "app_name": app_name,
        "app_id": app_id,
        "fetched_at": datetime.now().isoformat(),
        "total_reviews": len(reviews),
        "reviews": reviews,
    }
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, ensure_ascii=False, indent=2)


def call_deepseek(prompt: str, api_key: str, max_tokens: int = 4096) -> str:
    """调用 DeepSeek API"""
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": DEEPSEEK_MODEL,
        "messages": [
            {
                "role": "system",
                "content": (
                    "你是一位资深的产品分析师，擅长分析用户评论、撰写产品需求文档（PRD）"
                    "和生成测试用例。你必须严格基于提供的评论数据进行分析，每条结论都要"
                    "引用具体的评论ID或原文。回复使用中文，结构清晰，使用 Markdown 格式。"
                ),
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.3,
        "max_tokens": max_tokens,
    }

    resp = requests.post(DEEPSEEK_API_URL, headers=headers, json=payload, timeout=120)
    if resp.status_code != 200:
        raise RuntimeError(f"DeepSeek API 错误 ({resp.status_code}): {resp.text}")

    data = resp.json()
    return data["choices"][0]["message"]["content"]


def build_analysis_prompt(reviews: list, user_goal: str) -> str:
    """构建分析 prompt"""
    # 为了不超出 token 限制，取前 200 条评论（如果超过的话）
    reviews_for_analysis = reviews[:200]

    reviews_text = json.dumps(reviews_for_analysis, ensure_ascii=False, indent=2)

    prompt = f"""请分析以下 App Store 用户评论数据，完成以下三个任务。

用户特别关注的方向：{user_goal}

评论数据（JSON 格式，每条包含 id、userName、title、review、rating、date）：
```json
{reviews_text}
```

---

### 任务一：评论核心问题分析

请深入分析评论内容，找出用户抱怨的核心问题。要求：
1. **不能只靠关键词匹配**，必须理解语义和上下文
2. 每条结论必须**引用至少 2 条具体评论的 ID 和原文**作为证据
3. 按问题严重程度排序（抱怨人数从多到少）
4. 对每个问题给出影响评估（影响用户数估算、严重程度：高/中/低）
5. 将问题归类为：Bug类、体验类、付费/订阅类、功能缺失类、性能类、内容类等

输出格式：
```
## 核心问题分析

### 问题1：[问题标题]
- **类别**：[类别]
- **严重程度**：高/中/低
- **影响用户数估算**：约 X 人
- **问题描述**：[详细描述]
- **证据评论**：
  > [评论ID: xxx] 用户 xxx (评分 x/5): "[原文]"
  > [评论ID: xxx] 用户 xxx (评分 x/5): "[原文]"
```

---

### 任务二：产品需求文档（PRD）

基于上述分析的问题，生成产品需求文档。要求：
1. 每个需求必须**明确关联到任务一中的具体问题**
2. 按优先级划分 P0（必须修复）/ P1（重要优化）/ P2（锦上添花）
3. 按版本拆分 V1.0（紧急修复）/ V2.0（体验升级）
4. 每个需求包含：需求ID、标题、描述、优先级、目标版本、关联问题、验收标准

输出格式：
```
## 产品需求文档（PRD）

### V1.0 — 紧急修复版本

| 需求ID | 标题 | 优先级 | 描述 | 关联问题 | 验收标准 |
|--------|------|--------|------|----------|----------|
| REQ-001 | xxx  | P0     | xxx  | 问题1    | xxx      |

### V2.0 — 体验升级版本
（同上格式）
```

---

### 任务三：测试用例

基于 PRD 中的每个需求，生成对应的测试用例。要求：
1. 每个测试用例**必须能追溯到对应的需求（REQ-ID）和原始评论**
2. 包含：用例ID、关联需求ID、测试标题、前置条件、测试步骤、预期结果、优先级
3. 覆盖正向场景、边界场景和异常场景

输出格式：
```
## 测试用例

| 用例ID | 关联需求 | 测试标题 | 前置条件 | 测试步骤 | 预期结果 | 优先级 |
|--------|----------|----------|----------|----------|----------|--------|
| TC-001 | REQ-001  | xxx      | xxx      | 1. xxx 2. xxx | xxx | P0     |
```
"""

    return prompt


def render_progress_bar(label: str, progress: float, status_text: str = ""):
    """渲染一个进度条"""
    st.text(f"{label}: {status_text}")
    st.progress(min(progress, 1.0))


# ============================================================================
# 主界面
# ============================================================================

def main():
    st.title("📱 App Store Review Analyzer")
    st.markdown("输入美区 App Store 链接，AI 自动分析用户评论 → 生成 PRD → 生成测试用例")

    # ---- 侧边栏 ----
    with st.sidebar:
        st.header("⚙️ 配置")

        api_key = st.text_input(
            "🔑 DeepSeek API Key",
            type="password",
            placeholder="sk-xxxxxxxxxxxxxxxx",
            help="在 https://platform.deepseek.com 获取",
        )

        st.divider()

        st.markdown("### 📊 抓取设置")
        max_reviews = st.slider(
            "最多抓取评论数",
            min_value=20,
            max_value=500,
            value=200,
            step=10,
            help="抓取太多可能耗时较长",
        )

        st.divider()

        st.markdown("### 📁 缓存状态")
        cache = load_cache()
        if cache:
            st.success(f"✅ 缓存可用 — {cache['total_reviews']} 条评论")
            st.caption(f"App: {cache.get('app_name', 'N/A')}")
            st.caption(f"抓取时间: {cache.get('fetched_at', 'N/A')[:19]}")
        else:
            st.info("📭 暂无缓存")

        if st.button("🗑️ 清除缓存"):
            if CACHE_FILE.exists():
                CACHE_FILE.unlink()
            st.rerun()

    # ---- 主区域 ----
    col1, col2 = st.columns([3, 1])
    with col1:
        app_url = st.text_input(
            "🔗 美区 App Store 链接",
            placeholder="https://apps.apple.com/us/app/xxx/id123456789",
        )
    with col2:
        user_goal = st.text_input(
            "🎯 分析目标",
            placeholder="如：关注订阅问题",
            value="",
        )

    if not user_goal:
        user_goal = "全面分析用户抱怨和痛点"

    # ---- 缓存数据预览 ----
    cache_data = load_cache()
    if cache_data and not app_url:
        st.info(f"💡 检测到缓存数据（{cache_data['total_reviews']} 条评论），输入新链接将重新抓取，或直接点击分析按钮使用缓存。")

    # ---- 操作按钮 ----
    btn_col1, btn_col2 = st.columns([1, 1])

    with btn_col1:
        start_btn = st.button(
            "🚀 开始分析",
            type="primary",
            use_container_width=True,
            disabled=not api_key,
        )
    with btn_col2:
        fetch_only_btn = st.button(
            "📥 仅抓取评论",
            use_container_width=True,
        )

    if not api_key:
        st.warning("⚠️ 请先在侧边栏填写 DeepSeek API Key")

    # ---- 结果展示区 ----
    if "results" not in st.session_state:
        st.session_state.results = None

    # ========== 仅抓取评论 ==========
    if fetch_only_btn and app_url:
        app_id = extract_app_id(app_url)
        if not app_id:
            st.error("❌ 无法从链接中提取 App ID，请检查链接格式")
            return

        app_name = extract_app_name(app_url)

        with st.status("正在抓取评论...", expanded=True) as status:
            st.write(f"App: **{app_name}** (ID: {app_id})")

            try:
                # 优先使用 RSS feed 方式抓取
                formatted_reviews = fetch_reviews_via_rss(app_id, max_reviews)

                # 如果 RSS 方式没抓到，回退到 app-store-scraper
                if not formatted_reviews:
                    st.write("RSS 方式未获取到数据，尝试备用方式...")
                    try:
                        scraper = AppStore(country="us", app_name=app_name, app_id=app_id)
                        scraper.review(how_many=max_reviews)
                        reviews = scraper.reviews
                        if reviews:
                            for r in reviews:
                                formatted_reviews.append({
                                    "id": hashlib.md5(f"{r['userName']}{r['date']}{r['review']}".encode()).hexdigest()[:8],
                                    "userName": r.get("userName", "Anonymous"),
                                    "title": r.get("title", ""),
                                    "review": r.get("review", ""),
                                    "rating": r.get("rating", 0),
                                    "date": str(r.get("date", "")),
                                })
                    except Exception:
                        pass

                if not formatted_reviews:
                    status.update(label="未获取到评论", state="error")
                    st.error("未获取到任何评论，请检查 App ID 是否正确")
                    return

                save_cache(formatted_reviews, app_name, app_id)
                status.update(
                    label=f"✅ 抓取完成！共 {len(formatted_reviews)} 条评论已缓存",
                    state="complete",
                )
                st.success(f"已保存到 `{CACHE_FILE.name}`")

            except Exception as e:
                status.update(label="抓取失败", state="error")
                st.error(f"抓取出错: {e}")
                st.info("💡 如果之前有缓存文件，可以尝试直接点「开始分析」使用缓存数据")

    # ========== 开始分析 ==========
    if start_btn:
        app_id = None
        app_name = "Unknown App"
        reviews_data = None
        from_cache = False

        # Step 1: 获取评论数据
        if app_url:
            app_id = extract_app_id(app_url)
            app_name = extract_app_name(app_url)

            if not app_id:
                st.error("❌ 无法从链接中提取 App ID，请检查链接格式")
                return

            # 尝试抓取
            progress_bar = st.progress(0)
            st.info(f"🔄 正在从美区 App Store 抓取 **{app_name}** 的评论...")

            try:
                # 使用 RSS feed 方式抓取
                for pct in [0.05, 0.1, 0.15]:
                    time.sleep(0.2)
                    progress_bar.progress(pct, text=f"抓取评论中... {int(pct*100)}%")

                formatted_reviews = fetch_reviews_via_rss(app_id, max_reviews)

                progress_bar.progress(0.2, text="格式化数据...")

                # RSS 失败时回退到 app-store-scraper
                if not formatted_reviews:
                    try:
                        scraper = AppStore(country="us", app_name=app_name, app_id=app_id)
                        scraper.review(how_many=max_reviews)
                        reviews = scraper.reviews
                        if reviews:
                            for r in reviews:
                                formatted_reviews.append({
                                    "id": hashlib.md5(
                                        f"{r['userName']}{r['date']}{r['review']}".encode()
                                    ).hexdigest()[:8],
                                    "userName": r.get("userName", "Anonymous"),
                                    "title": r.get("title", ""),
                                    "review": r.get("review", ""),
                                    "rating": r.get("rating", 0),
                                    "date": str(r.get("date", "")),
                                })
                    except Exception:
                        pass

                progress_bar.progress(0.5, text="格式化数据...")

                if not formatted_reviews:
                    st.warning("⚠️ 未抓取到评论，尝试使用缓存...")
                    cache = load_cache()
                    if cache:
                        reviews_data = cache["reviews"]
                        app_name = cache["app_name"]
                        app_id = cache["app_id"]
                        from_cache = True
                        st.info(f"✅ 已从缓存加载 {len(reviews_data)} 条评论")
                    else:
                        st.error("❌ 无评论数据，也无缓存可用")
                        return
                else:
                    reviews_data = formatted_reviews
                    save_cache(reviews_data, app_name, app_id)
                    progress_bar.progress(0.6, text="评论已缓存")

            except Exception as e:
                st.warning(f"⚠️ 抓取失败: {e}")
                st.info("🔄 尝试加载缓存数据...")
                cache = load_cache()
                if cache:
                    reviews_data = cache["reviews"]
                    app_name = cache["app_name"]
                    app_id = cache["app_id"]
                    from_cache = True
                    st.info(f"✅ 已从缓存加载 {len(reviews_data)} 条评论")
                else:
                    st.error("❌ 网络不通且无缓存可用，请检查网络后重试")
                    return

        else:
            # 没有输入链接，尝试使用缓存
            cache = load_cache()
            if cache:
                reviews_data = cache["reviews"]
                app_name = cache["app_name"]
                app_id = cache["app_id"]
                from_cache = True
                st.info(f"📦 使用缓存数据：**{app_name}** — {len(reviews_data)} 条评论")
            else:
                st.error("❌ 请输入 App Store 链接，且无缓存可用")
                return

        # 确保 reviews_data 不为空
        if not reviews_data:
            st.error("❌ 没有可分析的评论数据")
            return

        total = len(reviews_data)

        # ---- Step 2: 数据概览 ----
        st.divider()
        st.subheader("📊 数据概览")

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.metric("评论总数", total)
        with col_b:
            avg_rating = sum(r["rating"] for r in reviews_data) / total if total else 0
            st.metric("平均评分", f"{avg_rating:.1f} ⭐")
        with col_c:
            source_tag = "📦 缓存" if from_cache else "🌐 实时抓取"
            st.metric("数据来源", source_tag)

        # 评分分布
        rating_dist = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0}
        for r in reviews_data:
            rt = int(r["rating"])
            if rt in rating_dist:
                rating_dist[rt] += 1

        st.caption("评分分布")
        dist_cols = st.columns(5)
        for i, (star, count) in enumerate(rating_dist.items()):
            with dist_cols[i]:
                pct = (count / total * 100) if total else 0
                st.metric(f"{star}⭐", count, f"{pct:.0f}%")

        # 低分评论占比
        low_rating = sum(1 for r in reviews_data if r["rating"] <= 2)
        if low_rating / total > 0.3:
            st.warning(f"⚠️ 低分评论（1-2星）占比 {low_rating/total*100:.0f}%，用户满意度较低")

        # ---- Step 3: AI 分析 ----
        st.divider()
        st.subheader("🤖 AI 分析中...")

        analysis_progress = st.progress(0, text="准备分析...")

        try:
            analysis_progress.progress(0.1, text="构建分析 Prompt...")
            prompt = build_analysis_prompt(reviews_data, user_goal)

            analysis_progress.progress(0.2, text="调用 DeepSeek API（可能需要 30-60 秒）...")
            st.caption(f"分析 {min(total, 200)} 条评论...")

            result_text = call_deepseek(prompt, api_key)

            analysis_progress.progress(1.0, text="分析完成！")

            st.session_state.results = {
                "app_name": app_name,
                "app_id": app_id,
                "total_reviews": total,
                "avg_rating": avg_rating,
                "from_cache": from_cache,
                "rating_dist": rating_dist,
                "analysis": result_text,
            }

        except Exception as e:
            st.error(f"❌ AI 分析失败: {e}")
            return

    # ---- 展示分析结果 ----
    if st.session_state.results:
        res = st.session_state.results

        st.divider()
        st.header(f"📋 分析结果 — {res['app_name']}")

        # 元信息
        st.caption(
            f"App ID: {res['app_id']} | 评论数: {res['total_reviews']} | "
            f"平均评分: {res['avg_rating']:.1f} | 数据源: {'缓存' if res['from_cache'] else '实时'}"
        )

        # 渲染分析结果
        st.markdown(res["analysis"])

        # 下载按钮
        st.divider()
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                "📥 下载分析报告 (Markdown)",
                data=res["analysis"],
                file_name=f"analysis_{res['app_name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md",
                mime="text/markdown",
            )
        with col_dl2:
            # 完整结果 JSON
            full_json = json.dumps(res, ensure_ascii=False, indent=2, default=str)
            st.download_button(
                "📥 下载完整结果 (JSON)",
                data=full_json,
                file_name=f"full_result_{res['app_name']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
            )


if __name__ == "__main__":
    main()
