import re
import os
import time
import subprocess
import shutil
import tempfile

from playwright.sync_api import Playwright, sync_playwright, expect


def is_file_in_use(filepath):
    """检查文件是否被占用"""
    if not os.path.exists(filepath):
        return False
    try:
        with open(filepath, 'a'):
            return False
    except IOError:
        return True


def close_excel_if_open():
    """尝试关闭占用文件的Excel进程"""
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'],
                      capture_output=True, text=True, timeout=10)
        time.sleep(1)
        return True
    except:
        return False


def force_delete_file(filepath):
    """强制删除文件，即使被占用"""
    if not os.path.exists(filepath):
        return True
    try:
        os.remove(filepath)
        return True
    except PermissionError:
        pass
    try:
        close_excel_if_open()
        time.sleep(2)
        os.remove(filepath)
        return True
    except:
        pass
    return False


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context(viewport={"width":1920,"height":1080}, accept_downloads=True)
    page = context.new_page()
    page.goto("http://gcwy.sce-icm.com/saasFrame/frame/")
    page.get_by_placeholder("用户名").click()
    page.get_by_placeholder("用户名").fill("15027614729")
    page.get_by_placeholder("密码").click()
    page.get_by_placeholder("密码").fill("123456")
    page.get_by_role("button", name="登录").click()
    page.locator("path").first.click()
    page.get_by_text("能源管理").click()
    with page.expect_popup() as page1_info:
        page.get_by_text("订单管理").click()
    page1 = page1_info.value
    page1.locator("iframe[name=\"iframeSystem\"]").content_frame.get_by_text("其他记录").click()
    time.sleep(1)

    save_dir = os.path.join(os.path.dirname(__file__), "财务共享资料", "2-能耗系统数据")
    save_path = os.path.join(save_dir, "财务订单.xls")
    os.makedirs(save_dir, exist_ok=True)

    with page1.expect_download() as download_info:
        page1.locator("iframe[name=\"iframeSystem\"]").content_frame.get_by_text("下载").click()
    download = download_info.value

    try:
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, f"财务订单_{int(time.time())}.xls")
        download.save_as(temp_path)

        if os.path.exists(save_path):
            if not force_delete_file(save_path):
                print(f"ERROR:无法删除旧文件 '财务订单.xls'，文件可能被其他程序占用。")
                print(f"提示：请检查是否有其他程序正在使用该文件（如杀毒软件、文件预览等）")
                os.remove(temp_path)
                raise PermissionError("无法删除已存在的目标文件")
        else:
            time.sleep(1)

        os.replace(temp_path, save_path)
        print(f"SUCCESS:文件已保存至：{save_path}")
    except PermissionError as e:
        print(f"ERROR:文件保存失败！\n原因：文件 '财务订单.xls' 正在被其他程序占用。\n解决方案：请关闭所有 Excel 窗口或其他可能访问该文件的程序。")
        raise
    except Exception as e:
        print(f"ERROR:文件保存失败：{str(e)}")
        raise

    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)