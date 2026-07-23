from playwright.sync_api import sync_playwright
import time, sys

URL = "http://localhost:8501"
API_KEY = "sk-1578e42d45a94d5abb4a8e8be06b9824"
APP_URL = "https://apps.apple.com/us/app/tiktok/id835599320"

log = []
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    errors = []
    page.on("console", lambda m: errors.append(f"[console:{m.type}] {m.text}"))
    page.on("pageerror", lambda e: errors.append(f"[pageerror] {e}"))

    log.append("打开页面...")
    page.goto(URL, wait_until="load", timeout=30000)
    page.wait_for_timeout(3000)

    # 填 API Key（侧边栏密码框）
    try:
        page.fill('input[type="password"]', API_KEY)
        log.append("✅ 已填 API Key")
    except Exception as e:
        log.append(f"❌ 填 API Key 失败: {e}")

    # 填 App 链接
    try:
        page.fill('input[placeholder*="apps.apple.com"]', APP_URL)
        log.append("✅ 已填 App 链接")
    except Exception as e:
        log.append(f"❌ 填 App 链接失败: {e}")

    page.wait_for_timeout(800)

    # 点“开始分析”
    try:
        page.get_by_role("button", name="🚀 开始分析").click(timeout=5000)
        log.append("✅ 已点击 开始分析")
    except Exception as e:
        log.append(f"⚠️ 点 开始分析 失败(可能按钮被禁用): {e}")
        # 备用：按文本找
        try:
            page.get_by_text("开始分析", exact=False).click(timeout=5000)
            log.append("✅ 已通过文本点击 开始分析")
        except Exception as e2:
            log.append(f"❌ 文本点击也失败: {e2}")

    # 等待分析完成（最多 120 秒），轮询结果标题
    done = False
    for i in range(60):
        page.wait_for_timeout(2000)
        try:
            if page.get_by_text("分析结果", exact=False).count() > 0:
                done = True
                log.append(f"✅ 检测到『分析结果』标题（约 {i*2}s）")
                break
        except Exception:
            pass
        # 抓取进度/错误
        try:
            body = page.inner_text("body")
            if "AI 分析失败" in body:
                log.append("❌ 页面显示 AI 分析失败")
                break
            if "手动粘贴评论" in body:
                log.append("⚠️ 页面进入『手动粘贴评论』模式（抓取为 0）")
                break
        except Exception:
            pass

    page.screenshot(path="/workspace/e2e.png", full_page=True)
    try:
        body_text = page.inner_text("body")
    except Exception:
        body_text = "(无法读取页面文本)"

    print("\n".join(log))
    print("\n=== 页面文本(前 3500 字) ===")
    print(body_text[:3500])
    print("\n=== 控制台/页面错误 ===")
    print("\n".join(errors[:40]) or "无")
    browser.close()
