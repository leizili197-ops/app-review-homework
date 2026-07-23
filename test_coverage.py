"""测试：美区(us) 是否每个 App 都能抓到评论。

用 app.py 里的双通道抓取函数 fetch_reviews_via_rss（XML + JSON，country 固定 us），
对一批美区热门 App 逐个测试，统计"能抓到" vs "抓到 0 条"的比例。
"""
import app
import time

# (App 名称, App Store ID) — 清一色美区热门 App
US_APPS = [
    ("TikTok", "835599320"),
    ("Instagram", "389801252"),
    ("YouTube", "544007664"),
    ("Spotify", "324684580"),
    ("Netflix", "363590051"),
    ("WhatsApp", "310633997"),
    ("Facebook", "284882215"),
    ("X (Twitter)", "333903271"),
    ("Snapchat", "447188370"),
    ("Pinterest", "429047995"),
    ("Discord", "863450654"),
    ("Telegram", "686449807"),
    ("Uber", "368677368"),
    ("Airbnb", "401626263"),
    ("Amazon", "297606951"),
    ("Google Maps", "585027354"),
    ("Gmail", "422689480"),
    ("Cash App", "711923939"),
    ("Robinhood", "938320880"),
    ("Venmo", "351727704"),
    ("Lyft", "729693728"),
    ("DoorDash", "674033150"),
    ("McDonald's", "922956432"),
    ("Starbucks", "331177150"),
    ("Bank of America", "492960317"),
    ("Chase", "298867247"),
    ("PayPal", "283646709"),
    ("Reddit", "1064216828"),
    ("Threads", "6446901002"),
    ("CapCut", "1511850313"),
]

print(f"{'App':16} {'ID':14} {'条数':>6} {'来源':>6}")
print("-" * 50)

ok, zero = [], []
for name, aid in US_APPS:
    app.LAST_FETCH_SOURCE = ""
    try:
        n = len(app.fetch_reviews_via_rss(aid, "us", 200))
    except Exception as e:
        n, src = -1, f"ERR:{e}"
        n_str = "ERR"
    else:
        src = app.LAST_FETCH_SOURCE or "-"
        n_str = str(n)
    print(f"{name:16} {aid:14} {n_str:>6} {src:>6}")
    (ok if n > 0 else zero).append((name, n))
    time.sleep(3)  # 礼貌间隔，避免触发 Apple 限流

print("-" * 50)
print(f"总计测试: {len(US_APPS)} 个")
print(f"能抓到(>0): {len(ok)} 个 -> {[n for n,_ in ok]}")
print(f"抓到 0 条 : {len(zero)} 个 -> {[n for n,_ in zero]}")
rate = len(ok) / len(US_APPS) * 100
print(f"抓取成功率: {rate:.0f}%")
