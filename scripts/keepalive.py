"""
Streamlit Cloud 앱 슬립 방지 스크립트
실제 브라우저(headless Chromium)로 앱에 접속해서
"Yes, get this app back up!" 버튼이 있으면 클릭해 깨운다.
"""
import sys
from playwright.sync_api import sync_playwright

APP_URL = "https://mumaev40-ab5vqedjqksvitijg7qilg.streamlit.app/"

def keep_awake(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            page.goto(url, timeout=30000, wait_until="networkidle")
            page.wait_for_timeout(3000)

            # 슬립 상태면 "Yes, get this app back up!" 버튼이 있음
            wake_button = page.get_by_text("Yes, get this app back up", exact=False)
            if wake_button.count() > 0:
                print(f"WAKE: {url} was sleeping, clicking wake button")
                wake_button.first.click()
                page.wait_for_timeout(15000)  # 앱 재시작 대기
            else:
                print(f"OK: {url} already awake")
        except Exception as e:
            print(f"ERROR visiting {url}: {e}")
        finally:
            browser.close()

if __name__ == "__main__":
    keep_awake(APP_URL)
