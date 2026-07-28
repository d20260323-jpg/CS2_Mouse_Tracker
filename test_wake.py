from playwright.sync_api import sync_playwright
import time

URL = "https://cs2mousetracker-dd6zri2zn7hatssysq28q5.streamlit.app/"

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto(URL, wait_until="networkidle", timeout=60000)

    try:
        btn = page.get_by_text("get this app back up", exact=False)
        if btn.count() > 0:
            btn.first.click()
            print("检测到休眠页,已点击唤醒按钮,等待启动...")
            time.sleep(30)
        else:
            print("应用已是醒着的状态 ✅")
    except Exception as e:
        print(f"未找到唤醒按钮,可能应用本来就是醒着的: {e}")

    browser.close()