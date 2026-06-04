import re
import os  # 新增
import time

from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.goto("http://gcwy.sce-icm.com/saasFrame/frame/")
    page.get_by_placeholder("用户名").click()
    page.get_by_placeholder("用户名").fill("15027614729")
    page.get_by_placeholder("密码").click()
    page.get_by_placeholder("密码").fill("123456")
    page.get_by_role("button", name="登录").click()
    page.get_by_role("img").first.click()
    page.get_by_text("能源管理").click()
    with page.expect_popup() as page1_info:
        page.get_by_text("报表管理").click()
    page1 = page1_info.value
    page1.locator("iframe[name=\"iframeSystem\"]").content_frame.get_by_text("租户逐时表底数(剩余金额)").click()
    time.sleep(1)
    with page1.expect_download() as download_info:
        page1.locator("iframe[name=\"iframeSystem\"]").content_frame.locator(".p-button").click()
    download = download_info.value

    # --- 完善下载保存部分 ---
    save_dir = r"C:\Users\lenovo\Desktop\财务共享资料\2-能耗系统数据"
    os.makedirs(save_dir, exist_ok=True)          # 自动创建目录（如已存在则跳过）
    download.save_as(os.path.join(save_dir, "能耗余额.xls"))
    print(f"文件已保存至：{save_dir}\\能耗余额.xls")
    # ---------------------

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)