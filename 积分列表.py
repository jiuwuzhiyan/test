import re
import time
import os
import sys
import pandas as pd
from datetime import datetime, timedelta, date
from playwright.sync_api import Playwright, sync_playwright, expect

BASE_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "财务共享资料")
POINTS_FILE = os.path.join(BASE_DIR, "4-报表中心", "积分列表.xlsx")
POINTS_DB = os.path.join(BASE_DIR, "4-报表中心", "积分列表.db")
TEMP_DOWNLOAD = os.path.join(BASE_DIR, "4-报表中心", "积分列表_temp.xlsx")

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def get_max_points_date():
    """获取最大积分时间，优先使用数据库"""
    log(f"读取原文件: {POINTS_FILE}")
    
    # 优先使用数据库
    if os.path.exists(POINTS_DB):
        try:
            from db_utils import get_points_db
            db = get_points_db(POINTS_DB)
            max_date = db.get_max_date()
            
            if max_date:
                start_date = max_date + timedelta(days=1)
                end_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
                log(f"从数据库获取最大积分时间: {max_date}, 开始日期: {start_date}, 结束日期: {end_date}")
                return start_date.strftime('%Y-%m-%d'), end_date
            log("数据库中没有有效数据")
        except Exception as e:
            log(f"从数据库读取失败，尝试Excel: {e}")
    
    # 回退到Excel
    if not os.path.exists(POINTS_FILE):
        log("原文件不存在，使用默认日期范围")
        return None, None
    
    try:
        df = pd.read_excel(POINTS_FILE, dtype=str, engine='openpyxl')
        if '积分时间' not in df.columns:
            log("文件中没有积分时间列")
            return None, None
        df['积分时间'] = pd.to_datetime(df['积分时间'], errors='coerce')
        df = df.dropna(subset=['积分时间'])
        if df.empty:
            log("文件中没有有效的积分时间数据")
            return None, None
        max_date = df['积分时间'].max()
        start_date = max_date + timedelta(days=1)
        end_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        log(f"最大积分时间: {max_date}, 开始日期: {start_date}, 结束日期: {end_date}")
        return start_date.strftime('%Y-%m-%d'), end_date
    except Exception as e:
        log(f"读取文件出错: {e}")
        return None, None

def append_points_data():
    """追加数据到数据库"""
    log(f"检查临时文件: {TEMP_DOWNLOAD}")
    if not os.path.exists(TEMP_DOWNLOAD):
        log("临时文件不存在，下载可能失败")
        return False
    
    try:
        os.makedirs(os.path.dirname(POINTS_DB), exist_ok=True)
        
        # 读取临时文件
        temp_df = pd.read_excel(TEMP_DOWNLOAD, dtype=str, engine='openpyxl')
        
        # 如果没有表头，尝试第一行作为数据
        if len(temp_df.columns) == 1 and temp_df.iloc[0, 0] == temp_df.columns[0]:
            temp_df = pd.read_excel(TEMP_DOWNLOAD, dtype=str, header=None, skiprows=1, engine='openpyxl')
            # 设置列名
            if len(temp_df.columns) >= 17:
                temp_df.columns = [
                    '积分时间', '活动项目', '业态', '租户', '会员卡号', '手机号', '性别',
                    '等级', '销售单号', '原订单号', '关联券订单号', '消费金额', '消费时间',
                    '积分数', '积分方式', '积分类型', '渠道'
                ]
        
        if temp_df.empty:
            log("下载的文件没有数据需要追加")
            return True
        
        # 统一列名 - 确保消费金额 (元) 列名保持不变，匹配数据库字段
        # 如果 Excel 中列名为"消费金额"，则重命名为"消费金额 (元)"以匹配数据库
        if '消费金额' in temp_df.columns and '消费金额 (元)' not in temp_df.columns:
            temp_df = temp_df.rename(columns={'消费金额': '消费金额 (元)'})
        
        # 将积分类型为"退款"的记录修改为"消费"类型
        if '积分类型' in temp_df.columns:
            refund_count = len(temp_df[temp_df['积分类型'] == '退款'])
            temp_df['积分类型'] = temp_df['积分类型'].replace('退款', '消费')
            if refund_count > 0:
                log(f"已将 {refund_count} 条退款类型记录修改为消费类型")
        
        # 删除消费时间字段中的 '--' 值
        if '消费时间' in temp_df.columns:
            temp_df['消费时间'] = temp_df['消费时间'].replace('--', '')
            log("已删除消费时间字段中的 '--' 值")
        
        # 更新数据库
        try:
            from db_utils import get_points_db
            db = get_points_db(POINTS_DB)
            
            # 数据类型处理
            if '积分数' in temp_df.columns:
                temp_df['积分数'] = pd.to_numeric(temp_df['积分数'], errors='coerce')
            
            inserted = db.insert_data(temp_df)
            log(f"数据库更新成功，插入 {inserted} 条记录")
        except Exception as e:
            log(f"数据库更新失败: {e}")
            raise
        
        return True
    except Exception as e:
        log(f"追加数据出错: {e}")
        import traceback
        log(traceback.format_exc())
        return False

def run(playwright: Playwright) -> None:
    start_date, end_date = get_max_points_date()
    if not start_date or not end_date:
        yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        log(f"使用默认日期范围: {yesterday} 至 {yesterday}")
        start_date = yesterday
        end_date = yesterday
    else:
        log(f"积分日期范围: {start_date} 至 {end_date}")

    browser = None
    context = None
    page = None
    page1 = None
    
    try:
        log("启动浏览器...")
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width":1920,"height":1080})
        page = context.new_page()

        log("打开登录页面...")
        page.goto("https://bk.sce-icm.com/", timeout=30000)
        
        log("输入用户名...")
        page.get_by_placeholder("请输入用户名").click(timeout=10000)
        page.get_by_placeholder("请输入用户名").fill("668702")
        
        log("输入密码...")
        page.get_by_placeholder("请输入密码").click(timeout=10000)
        page.get_by_placeholder("请输入密码").fill("Xin911005")
        
        log("点击登录...")
        page.get_by_role("button", name="登录").click(timeout=10000)
        page.wait_for_load_state("networkidle", timeout=30000)

        log("导航到客户菜单...")
        page.get_by_role("link", name="客户").click(timeout=10000)
        
        log("导航到积分列表...")
        page.get_by_role("link", name="积分列表").click(timeout=10000)
        page.wait_for_load_state("networkidle", timeout=30000)

        log("填写开始日期...")
        # Ant Design日期选择器需要点击展开后操作
        start_picker = page.locator("#createdAtStart")
        start_picker.click(timeout=10000)
        # 选择非readonly的input（第二个input是可编辑的）
        start_input = start_picker.locator("input").nth(1)
        start_input.fill(start_date)
        # 按Enter确认输入
        page.keyboard.press("Enter")
        
        log("填写结束日期...")
        end_picker = page.locator("#createdAtEnd")
        end_picker.click(timeout=10000)
        end_input = end_picker.locator("input").nth(1)
        end_input.fill(end_date)
        # 按Enter确认输入
        page.keyboard.press("Enter")
        
        log("点击筛选...")
        page.get_by_role("button", name="筛选").click(timeout=10000)
        page.wait_for_load_state("networkidle", timeout=10000)

        log("触发导出...")
        page.get_by_role("button", name="导出").click(timeout=10000)

        log("等待弹出页面...")
        with page.expect_popup(timeout=30000) as page1_info:
            page.get_by_text("去查看").click(timeout=10000)
        page1 = page1_info.value
        time.sleep(10)
        log("等待下载链接...")
        download_link = page1.locator("role=row[name=/积分列表.*/i] >> a").first
        download_link.wait_for(state="visible", timeout=30000)

        log("开始下载...")
        with page1.expect_download(timeout=30000) as download_info:
            download_link.click(timeout=10000)
        download = download_info.value

        log(f"保存文件到: {TEMP_DOWNLOAD}")
        download.save_as(TEMP_DOWNLOAD)

        log("关闭弹出页面...")
        page1.close()

        log("追加数据到原文件...")
        if append_points_data():
            log("数据追加成功")
            if os.path.exists(TEMP_DOWNLOAD):
                os.remove(TEMP_DOWNLOAD)
                log(f"已删除临时文件: {TEMP_DOWNLOAD}")
            print("SUCCESS")
            sys.exit(0)
        else:
            log("追加数据失败")
            print("FAIL: 追加数据失败")
            sys.exit(1)
            
    except Exception as e:
        log(f"执行出错: {e}")
        import traceback
        log(traceback.format_exc())
        
        # 清理资源
        if page1:
            try:
                page1.close()
            except:
                pass
        if page:
            try:
                page.close()
            except:
                pass
        if context:
            try:
                context.close()
            except:
                pass
        if browser:
            try:
                browser.close()
            except:
                pass
        
        print(f"FAIL: {e}")
        sys.exit(1)

with sync_playwright() as playwright:
    run(playwright)
