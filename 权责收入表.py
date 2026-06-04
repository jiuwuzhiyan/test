import re
import os  # 新增：用于获取桌面路径
import time

from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context(viewport={"width": 1920, "height": 1080})
    page = context.new_page()
    page.goto("https://bgerp.sce-re.com/")
    page.get_by_role("textbox", name="请输入用户名").click()
    page.get_by_role("textbox", name="请输入用户名").fill("668702")
    page.get_by_role("textbox", name="请输入密码").click()
    page.get_by_role("textbox", name="请输入密码").fill("Xin911005")
    page.get_by_role("button", name="登录").click()
    page.locator("div").filter(has_text=re.compile(r"^报表$")).click()
    page.locator("div").filter(has_text=re.compile(r"^租赁财务报表$")).click()
    page.get_by_role("menuitem", name="权责收入表(新)").click()
    time.sleep(10)
    page.locator("#main_iframe").content_frame.locator(
        "div:nth-child(14) > .fr-trigger-text > .fr-trigger-texteditor").click()
    page.locator("#main_iframe").content_frame.locator(
        ".fr-trigger-texteditor.fr-widget-background.fr-widget-font.fr-trigger-texteditor-focus").fill("唐山中骏世界城")
    page.locator("#main_iframe").content_frame.locator("#fr-btn-FORMSUBMIT0 div").filter(
        has_text=re.compile(r"^查询$")).click()
    time.sleep(10)

    # ========== 仅修改这部分：下载并保存到桌面 ==========
    with page.expect_download() as download_info:
        page.locator("#main_iframe").content_frame.get_by_role("button", name="原样导出").click()
    download = download_info.value

    # 获取系统桌面路径（Windows/Mac/Linux 通用）
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    # 获取下载文件的原始文件名
    file_name = download.suggested_filename
    # 拼接完整保存路径
    save_path = os.path.join(desktop_path, file_name)
    # 保存文件到桌面
    download.save_as(save_path)
    print(f"文件已成功保存到桌面：{save_path}")
    # ====================================================

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)