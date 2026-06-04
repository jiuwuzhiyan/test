import re
import os
import time
from playwright.sync_api import Playwright, sync_playwright, expect

USERNAME = os.environ.get("ERP_USER", "668702")
PASSWORD = os.environ.get("ERP_PWD", "Xin911005")
BASE_URL = "https://bgerp.sce-re.com"


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(
        headless=False,
        slow_mo=500,  # 慢动作执行，方便调试观察
        args=['--start-maximized']
    )
    context = browser.new_context(no_viewport=True)
    page = context.new_page()

    # ---------- 1. 登录 ----------
    page.goto(BASE_URL)
    page.get_by_placeholder("请输入用户名").wait_for(state="visible", timeout=10000)
    page.get_by_placeholder("请输入用户名").fill(USERNAME)
    page.get_by_placeholder("请输入密码").fill(PASSWORD)
    page.get_by_role("button", name="登录").click()
    page.wait_for_url("**/home*", timeout=15000)

    # ---------- 2. 导航到报表 ----------
    menu1 = page.locator('//*[@id="app"]/section/section/div/div/ul/div[8]/li/div')
    menu1.wait_for(state="visible")
    menu1.click()

    menu2 = page.locator('//*[@id="app"]/section/section/div/div/ul/div[8]/li/ul/div[1]/li/div')
    menu2.wait_for(state="visible")
    menu2.click()

    menu3 = page.locator('//*[@id="app"]/section/section/div/div/ul/div[8]/li/ul/div[1]/li/ul/li[6]')
    menu3.wait_for(state="visible")
    menu3.click()

    # ---------- 3. 查询报表 ----------
    # 报表页面的iframe（#main_iframe）
    report_frame = page.frame_locator("#main_iframe").first
    report_frame.locator('//*[@id="app"]/div/form/div/div[5]/button[1]').wait_for(state="visible", timeout=15000)
    report_frame.locator('//*[@id="app"]/div/form/div/div[5]/button[1]').click(timeout=20000)
    report_frame.locator('//*[@id="app"]/div/div[1]/button').wait_for(state="visible", timeout=15000)
    report_frame.locator('//*[@id="app"]/div/div[1]/button').click(timeout=20000)

    # 等待报表生成完成，然后点击进入下载中心
    page.wait_for_timeout(5000)
    page.locator('//*[@id="app"]/section/section/div/div/ul/div[9]/li').wait_for(state="visible", timeout=10000)
    page.locator('//*[@id="app"]/section/section/div/div/ul/div[9]/li').click()

    # ---------- 4. 核心修改：使用你捕获的XPath定位下载按钮 ----------
    # 第一步：定位下载中心专属iframe（替换原有的#main_iframe）
    # 对应你第一个XPath：//iframe[contains(@src, "https://bgerp.sce-re.com/pay_user/#/download")]
    download_frame = page.frame_locator('//iframe[contains(@src, "pay_user/#/download")]')


    # 第二步：在下载iframe内定位第一个下载按钮
    # 对应你第二个XPath：//table//tr[1]/td//button/span[normalize-space(.)="下载"]
    download_btn = download_frame.locator('//table//tr[1]/td//button/span[normalize-space(.)="下载"]')

    # 等待按钮可见（Playwright仅支持attached/detached/visible/hidden四种状态）
    download_btn.wait_for(state="visible", timeout=20000)
    download_btn.scroll_into_view_if_needed()

    # 捕获下载事件并保存文件
    with page.expect_download(timeout=30000) as download_info:
        download_btn.click()

    download = download_info.value
    save_path = f"./{download.suggested_filename}"
    download.save_as(save_path)
    print(f"✅ 文件下载成功，保存路径：{os.path.abspath(save_path)}")

    # ---------- 5. 清理资源 ----------
    page.wait_for_timeout(1000)
    context.close()
    browser.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)