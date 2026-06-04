from playwright.sync_api import sync_playwright
import os
import datetime
import time

# 配置参数
WEBSITE_URL = "https://bk.sce-icm.com/"
DOWNLOAD_FOLDER = r"d:\BaiduSyncdisk\PYTHON"
LOGIN_USERNAME = "668702"  # 需要填写: 登录用户名
LOGIN_PASSWORD = "Xin911005"  # 需要填写: 登录密码
TIMEOUT = 30000  # 30秒超时

def download_points_file(username: str, password: str, month: str = None):
    """
    使用Playwright下载积分列表文件

    Args:
        username: 登录用户名
        password: 登录密码
        month: 指定月份 (格式: "YYYY-MM"), 默认为当前月份
    """
    if month is None:
        month = datetime.datetime.now().strftime("%Y-%m")

    with sync_playwright() as p:
        # 启动浏览器
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            viewport={'width': 1920, 'height': 1080}
        )
        page = context.new_page()

        try:
            print(f"正在访问网站: {WEBSITE_URL}")
            page.goto(WEBSITE_URL, timeout=TIMEOUT)

            # 等待页面加载
            page.wait_for_load_state("networkidle")
            time.sleep(2)

            # 检查是否需要登录
            if page.url.find("login") != -1 or page.locator("input[type='text'], input[name='username']").count() > 0:
                print("检测到登录页面，正在登录...")
                login(page, username, password)

            # 等待登录完成并跳转到数据页面
            page.wait_for_load_state("networkidle")
            time.sleep(3)

            print("正在查找数据下载入口...")

            # 查找积分列表相关入口
            # 根据网站结构调整选择器
            download_link = find_download_link(page, month)

            if download_link:
                print(f"找到下载链接: {download_link}")
                # 点击下载
                with page.expect_download() as download_info:
                    page.click(download_link)
                    time.sleep(2)

                download = download_info.value
                filename = f"积分列表_{month}.xlsx"
                save_path = os.path.join(DOWNLOAD_FOLDER, filename)

                # 保存文件
                download.save_as(save_path)
                print(f"✅ 文件已下载保存至: {save_path}")
                return save_path
            else:
                print("❌ 未找到下载链接，请手动检查网站结构")

        except Exception as e:
            print(f"❌ 操作失败: {str(e)}")
            # 截图保存错误状态
            error_screenshot = os.path.join(DOWNLOAD_FOLDER, "error_screenshot.png")
            page.screenshot(path=error_screenshot)
            print(f"已保存错误截图: {error_screenshot}")
            return None

        finally:
            browser.close()


def login(page, username: str, password: str):
    """处理登录流程"""
    try:
        # 查找用户名输入框
        username_input = page.locator("input[type='text'], input[name='username'], input[placeholder*='账号'], input[placeholder*='用户名']").first
        username_input.fill(username)
        time.sleep(0.5)

        # 查找密码输入框
        password_input = page.locator("input[type='password'], input[name='password']").first
        password_input.fill(password)
        time.sleep(0.5)

        # 查找登录按钮
        login_button = page.locator("button[type='submit'], button:has-text('登录'), button:has-text('登 录')").first
        login_button.click()
        print("登录请求已发送...")

        # 等待登录完成
        time.sleep(3)

    except Exception as e:
        print(f"登录过程出错: {str(e)}")
        raise


def find_download_link(page, month: str):
    """
    查找下载链接
    根据实际网站结构调整此函数
    """
    # 方法1: 查找包含"积分"和"下载"的链接
    selectors = [
        f"a:has-text('积分列表'), a:has-text('积分')",
        f"button:has-text('下载'), button:has-text('导出')",
        f"a:has-text('{month}')",
        "table tr td a[href*='xlsx'], table tr td a[href*='download']",
    ]

    for selector in selectors:
        try:
            elements = page.locator(selector)
            if elements.count() > 0:
                print(f"找到可能的选择器: {selector}, 数量: {elements.count()}")
                return selector
        except:
            continue

    # 方法2: 打印页面所有链接用于调试
    print("\n当前页面所有链接:")
    links = page.locator("a").all()
    for i, link in enumerate(links):
        try:
            text = link.text_content()
            href = link.get_attribute("href")
            if text and href:
                print(f"{i+1}. {text.strip()} -> {href}")
        except:
            pass

    # 方法3: 打印页面所有按钮用于调试
    print("\n当前页面所有按钮:")
    buttons = page.locator("button").all()
    for i, btn in enumerate(buttons):
        try:
            text = btn.text_content()
            if text:
                print(f"{i+1}. {text.strip()}")
        except:
            pass

    return None


def append_data_to_existing(original_file: str, new_file: str):
    """
    将新下载的数据追加到原有数据中

    Args:
        original_file: 原始数据文件路径
        new_file: 新下载的数据文件路径
    """
    import pandas as pd

    try:
        print(f"正在读取原始数据: {original_file}")
        df_original = pd.read_excel(original_file, dtype=str)

        print(f"正在读取新数据: {new_file}")
        df_new = pd.read_excel(new_file, dtype=str)

        # 追加数据
        df_combined = pd.concat([df_original, df_new], ignore_index=True)

        # 去重（基于关键列）
        key_columns = ['会员号', '积分时间', '消费金额(元)', '积分数']  # 根据实际列名调整
        df_combined = df_combined.drop_duplicates(subset=key_columns, keep='last')

        # 保存合并后的数据
        backup_file = original_file.replace('.xlsx', f'_backup_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
        df_original.to_excel(backup_file, index=False)
        print(f"已备份原始文件至: {backup_file}")

        df_combined.to_excel(original_file, index=False)
        print(f"✅ 数据已成功追加！总行数: {len(df_combined)}")

        return df_combined

    except Exception as e:
        print(f"❌ 数据追加失败: {str(e)}")
        return None


if __name__ == "__main__":
    print("=" * 60)
    print("积分数据下载自动化工具")
    print("=" * 60)

    # 获取用户输入
    username = input("请输入登录用户名: ").strip()
    if not username:
        print("用户名不能为空")
        exit(1)

    password = input("请输入登录密码: ").strip()
    if not password:
        print("密码不能为空")
        exit(1)

    # 设置当前月份
    current_month = datetime.datetime.now().strftime("%Y-%m")
    month_input = input(f"请输入要下载的月份 (格式: YYYY-MM, 默认: {current_month}): ").strip()
    month = month_input if month_input else current_month

    print(f"\n开始下载 {month} 月份的数据...")
    print("-" * 60)

    # 执行下载
    downloaded_file = download_points_file(username, password, month)

    if downloaded_file:
        print("\n" + "=" * 60)
        print("下载成功！")
        print("=" * 60)

        # 询问是否追加到现有数据
        original_file = os.path.join(DOWNLOAD_FOLDER, "积分列表.xlsx")
        if os.path.exists(original_file):
            append_choice = input(f"\n是否将新数据追加到现有文件 ({original_file})? (y/n): ").strip().lower()
            if append_choice == 'y':
                append_data_to_existing(original_file, downloaded_file)
        else:
            print(f"\n未找到原始数据文件，将下载的文件重命名为: {original_file}")
            import shutil
            shutil.copy(downloaded_file, original_file)
