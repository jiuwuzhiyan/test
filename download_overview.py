# download_overview.py
import os
import sys
from datetime import datetime
from playwright.sync_api import sync_playwright

BASE_URL = "http://gcwy.sce-icm.com/saasFrame/frame/"
USERNAME = "15027614729"
PASSWORD = "123456"

def main(target_path):
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()

            # 登录
            page.goto(BASE_URL, wait_until="networkidle")
            page.get_by_role("textbox", name="用户名").fill(USERNAME)
            page.get_by_role("textbox", name="密码").fill(PASSWORD)
            with page.expect_navigation(wait_until="networkidle"):
                page.get_by_role("button", name="登录").click()

            # 进入租户管理
            page.locator("svg path").first.click(timeout=5000)
            page.get_by_text("能源管理").click(timeout=5000)
            with page.expect_popup() as page1_info:
                page.get_by_text("租户管理").click(timeout=5000)
            page1 = page1_info.value
            page1.wait_for_load_state("networkidle", timeout=10000)

            # 导出
            iframe = page1.frame_locator("iframe[name=\"iframeSystem\"]")
            export_btn = iframe.locator(".p-button")
            export_btn.wait_for(timeout=15000)
            with page1.expect_download() as download_info:
                export_btn.click()
            download = download_info.value

            if download.failure():
                raise Exception(f"下载失败: {download.failure()}")

            download.save_as(target_path)

            if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
                print("SUCCESS")
                sys.exit(0)
            else:
                print("FAIL: 文件为空")
                sys.exit(1)

    except Exception as e:
        print(f"FAIL: {e}")
        sys.exit(1)
    finally:
        try:
            context.close()
            browser.close()
        except:
            pass

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python download_overview.py <target_path>")
        sys.exit(1)
    main(sys.argv[1])