import re
import os
import time
from playwright.sync_api import Playwright, sync_playwright, expect

# 强烈建议用环境变量替代默认值
USERNAME = os.environ.get("ERP_USER", "668702")
PASSWORD = os.environ.get("ERP_PWD", "Xin911005")
BASE_URL = "https://bgerp.sce-re.com"

def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(
        headless=False,
        args=['--start-maximized']          # 启动时最大化
    )
    context = browser.new_context(no_viewport=True)   # 禁用固定视口，页面可撑满窗口
    page = context.new_page()
    # ---------- 1. 登录 ----------
    page.goto(BASE_URL)
    page.get_by_placeholder("请输入用户名").fill(USERNAME)
    page.get_by_placeholder("请输入密码").fill(PASSWORD)
    page.get_by_role("button", name="登录").click()
    # 等待登录成功跳转到主系统
    page.wait_for_url("**/home*", timeout=10000)

    # ---------- 2. 导航到报表 ----------
    page.locator('//*[@id="app"]/section/section/div/div/ul/div[8]/li/div').click()
    page.locator('//*[@id="app"]/section/section/div/div/ul/div[8]/li/ul/div[1]/li/div').click()
    page.locator('//*[@id="app"]/section/section/div/div/ul/div[8]/li/ul/div[1]/li/ul/li[1]').click()

    # ---------- 3. 查询报表 ----------
    page.locator("#main_iframe").content_frame.locator(".x-icon").first.click()
    page.locator("#main_iframe").content_frame.locator(".checkbox-content").first.click()
    page.frame_locator("#main_iframe").locator('//*[@id="fr-btn-FORMSUBMIT0"]').click(timeout=15000)
    time.sleep(5)  # 等待报表加载完成

    # ---------- 4. 导出 Excel ----------
    page.frame_locator("#main_iframe").locator('//*[@id="fr-btn-Export"]/div[1]/div[2]').click(timeout=15000)
    page.locator("#main_iframe").content_frame.get_by_text("Excel").click()

    # 优化：用 expect_download 包裹“原样导出”点击，确保一次触发并捕获下载
    with page.expect_download() as download_info:
        page.locator("#main_iframe").content_frame.get_by_text("原样导出").click()
    download = download_info.value
    save_path = f"./{download.suggested_filename}"
    download.save_as(save_path)
    print(f"✅ 文件已保存至：{save_path}")

    # ---------- 5. 清理 ----------
    page.wait_for_timeout(1000)
    context.close()
    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)