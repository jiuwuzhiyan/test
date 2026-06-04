import multiprocessing
multiprocessing.freeze_support()

import sys
import os
import streamlit as st
from datetime import date
from openpyxl import load_workbook
import pandas as pd
from openpyxl import Workbook
from datetime import datetime
from playwright.sync_api import sync_playwright
import subprocess
from io import BytesIO
import time
import shutil
from datetime import date, datetime, timedelta

# ---------- 全局文件基础目录 ----------
import platform

# 尝试本地路径，如果不存在则使用 UNC 共享路径（给同事用）
_LOCAL_DIR = r"C:\Users\lenovo\Desktop\财务共享资料"
_UNC_DIR = r"\\172.16.103.130\Users\lenovo\Desktop\财务共享资料"

if os.path.exists(_LOCAL_DIR):
    BASE_DIR = _LOCAL_DIR
elif os.path.exists(_UNC_DIR):
    BASE_DIR = _UNC_DIR
else:
    BASE_DIR = _LOCAL_DIR

# ---------- 所有文件路径（集中管理） ----------
excel_path = os.path.join(BASE_DIR, "9-数据库", "能耗_电费数据.xlsx")
water_excel_path = os.path.join(BASE_DIR, "9-数据库","能耗_水费数据.xlsx")
bai_gou_template_path = os.path.join(BASE_DIR, "0-模板", "应收帐单导入.xls")
kai_piao_template_path = os.path.join(BASE_DIR, "0-模板", "批量开票-导入开票模板.xlsx")
customer_info_path = os.path.join(BASE_DIR, "1-金蝶系统数据", "客户资料.xlsx")
zhengpu_contract_path = os.path.join(BASE_DIR, "3-百购系统数据", "租赁合同台账_正铺.xlsx")
duojing_contract_path = os.path.join(BASE_DIR, "3-百购系统数据", "租赁合同台账_多经.xlsx")
invoice_info_path = os.path.join(BASE_DIR, "0-模板", "开票信息.xlsx")
tenant_overview_path = os.path.join(BASE_DIR, "2-能耗系统数据", "租户概览.xls")

# 其他本地路径
archive_path = r"D:\BaiduSyncdisk\7-能源系统\结转留存.xlsx"          # 电费结转存档
balance_path = os.path.join(BASE_DIR, "1-金蝶系统数据", "辅助核算项目余额表.xlsx")
energy_balance_path = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗余额.xls")
cwj_order_path = os.path.join(BASE_DIR, "2-能耗系统数据", "财务订单.xls")

# ---------- 工具函数 ----------
def download_tenant_overview(target_path):
    """
    通过独立脚本自动下载租户概览文件。
    返回 (成功标志, 消息)
    """
    script_path = os.path.join(os.path.dirname(__file__), "download_overview.py")
    try:
        # 使用与当前 Python 解释器相同的解释器运行脚本
        result = subprocess.run(
            [sys.executable, script_path, target_path],
            capture_output=True,
            text=True,
            timeout=60  # 设置超时
        )
        stdout = result.stdout.strip()
        if result.returncode == 0 and "SUCCESS" in stdout:
            return True, f"文件已更新至 {target_path}"
        else:
            error_msg = stdout if stdout else result.stderr.strip()
            return False, f"下载失败：{error_msg}"
    except subprocess.TimeoutExpired:
        return False, "下载超时（60秒）"
    except Exception as e:
        return False, f"调用下载脚本出错：{e}"

def process_date_column(df):
    if '申请日期' in df.columns:
        df['申请日期'] = pd.to_datetime(df['申请日期'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['申请日期'] = df['申请日期'].fillna('')
    return df

@st.cache_data(ttl=10)
def load_excel(path):
    if os.path.exists(path):
        df = pd.read_excel(path, dtype=str)
        df = process_date_column(df)
        return df
    else:
        return pd.DataFrame(columns=["表号", "品牌", "客户编号", "申请日期", "表底", "清表类型", "自建ID"])

def clear_form():
    st.session_state.table_no = ""
    st.session_state.brand = ""
    st.session_state.customer_id = ""
    st.session_state.apply_date = date.today()
    st.session_state.meter_bottom = 0.0
    st.session_state.clear_type = "进场清表"
    st.session_state.id_input = ""

def get_file_mod_time(file_path):
    """获取文件修改时间戳"""
    if os.path.exists(file_path):
        return os.path.getmtime(file_path)
    return 0

@st.cache_data(ttl=60)
def load_points_data(_file_path, _db_mtime, _excel_mtime):
    """
    加载积分数据，支持Excel和SQLite数据库两种格式
    优先使用SQLite数据库，性能更高
    缓存策略：文件修改时间变化时自动刷新
    """
    # 检查是否存在SQLite数据库文件
    db_path = _file_path.replace('.xlsx', '.db')
    
    # 优先使用数据库
    if os.path.exists(db_path):
        try:
            from db_utils import get_points_db
            db = get_points_db(db_path)
            df = db.query_data()
            
            if not df.empty:
                # 数据类型转换
                if '积分数' in df.columns:
                    df['积分数'] = pd.to_numeric(df['积分数'], errors='coerce').fillna(0)
                if '消费金额(元)' in df.columns:
                    df['消费金额(元)'] = pd.to_numeric(df['消费金额(元)'], errors='coerce').fillna(0)
                if '积分时间' in df.columns:
                    df['积分时间'] = pd.to_datetime(df['积分时间'], errors='coerce')
                return df
        except Exception as e:
            st.warning(f"数据库读取失败，尝试使用Excel: {str(e)}")
    
    # 如果数据库不存在或读取失败，回退到Excel
    if not os.path.exists(_file_path):
        return None
    
    # 尝试使用parquet缓存，如果不存在或过期则重新读取Excel
    parquet_path = _file_path.replace('.xlsx', '.parquet')
    use_parquet = False
    
    if os.path.exists(parquet_path):
        # 检查parquet文件是否比Excel文件新
        excel_mtime = os.path.getmtime(_file_path)
        parquet_mtime = os.path.getmtime(parquet_path)
        if parquet_mtime > excel_mtime:
            use_parquet = True
    
    if use_parquet:
        try:
            df = pd.read_parquet(parquet_path)
            # 验证数据完整性
            if '积分时间' in df.columns and '积分数' in df.columns:
                return df
        except Exception as e:
            # 如果parquet读取失败，回退到Excel
            pass
    
    # 读取Excel文件
    try:
        df = pd.read_excel(_file_path, dtype=str, engine='openpyxl')
        
        # 数据类型转换
        if '积分数' in df.columns:
            df['积分数'] = pd.to_numeric(df['积分数'], errors='coerce').fillna(0)
        if '消费金额(元)' in df.columns:
            df['消费金额(元)'] = pd.to_numeric(df['消费金额(元)'], errors='coerce').fillna(0)
        if '积分时间' in df.columns:
            df['积分时间'] = pd.to_datetime(df['积分时间'], errors='coerce')
        
        # 尝试保存parquet缓存，但不阻塞主流程
        try:
            df.to_parquet(parquet_path, index=False)
        except Exception:
            pass
            
        return df
    except Exception as e:
        return None

# ==================== 积分列表更新优化函数 ====================
def get_points_max_date(points_file):
    """
    获取积分文件中最大的积分时间，返回 (最大日期+1天, 昨天日期)
    支持Excel和SQLite数据库两种格式
    """
    # 检查是否存在SQLite数据库文件
    db_path = points_file.replace('.xlsx', '.db')
    
    # 优先使用数据库
    if os.path.exists(db_path):
        try:
            from db_utils import get_points_db
            db = get_points_db(db_path)
            max_date = db.get_max_date()
            
            if max_date:
                start_date = max_date + timedelta(days=1)
                end_date = datetime.now() - timedelta(days=1)
                return start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
            return None, None
        except Exception as e:
            st.warning(f"数据库查询失败，尝试使用Excel: {str(e)}")
    
    # 如果数据库不存在或读取失败，回退到Excel
    if not os.path.exists(points_file):
        return None, None
    
    try:
        # 使用更高效的方式读取，只获取积分时间列
        df = pd.read_excel(points_file, dtype=str, engine='openpyxl', usecols=['积分时间'])
        if '积分时间' not in df.columns:
            return None, None
        df['积分时间'] = pd.to_datetime(df['积分时间'], errors='coerce')
        df = df.dropna(subset=['积分时间'])
        if df.empty:
            return None, None
        max_date = df['积分时间'].max()
        start_date = max_date + timedelta(days=1)
        end_date = datetime.now() - timedelta(days=1)
        return start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')
    except Exception as e:
        return None, None

def safe_file_operation(source_path, target_path, operation="copy"):
    """安全的文件操作，确保数据完整性"""
    try:
        if operation == "copy":
            shutil.copy2(source_path, target_path)
        elif operation == "move":
            shutil.move(source_path, target_path)
        return True
    except Exception as e:
        return False

def create_backup_version(file_path, max_versions=5):
    """创建带版本号的备份，保留最近的几个版本"""
    if not os.path.exists(file_path):
        return None
    
    try:
        base_dir = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        name, ext = os.path.splitext(base_name)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f"{name}_backup_{timestamp}{ext}"
        backup_path = os.path.join(base_dir, backup_name)
        
        # 创建备份
        shutil.copy2(file_path, backup_path)
        
        # 清理旧备份，只保留最新的max_versions个
        backup_files = []
        for f in os.listdir(base_dir):
            if f.startswith(f"{name}_backup_") and f.endswith(ext):
                backup_files.append(os.path.join(base_dir, f))
        
        # 按修改时间排序，删除旧的
        backup_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        for old_backup in backup_files[max_versions:]:
            try:
                os.remove(old_backup)
            except:
                pass
        
        return backup_path
    except Exception as e:
        return None

def validate_dataframe(df):
    """验证DataFrame数据完整性"""
    if df is None or df.empty:
        return False, "数据为空"
    
    # 检查必要的列是否存在
    required_cols = ['积分时间', '积分数']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        return False, f"缺少必要列: {', '.join(missing_cols)}"
    
    return True, "数据验证通过"

def update_points_list(base_dir):
    """
    优化版积分列表更新函数
    - 事务管理：先写入临时文件，成功后再覆盖原文件
    - 错误处理：完整的异常捕获和日志记录
    - 性能优化：快速获取最大日期，只读取必要列
    - 数据安全：多版本备份，增量更新
    - 支持SQLite数据库和Excel双格式存储
    """
    points_file = os.path.join(base_dir, "4-报表中心", "积分列表.xlsx")
    db_file = os.path.join(base_dir, "4-报表中心", "积分列表.db")
    temp_file = os.path.join(base_dir, "4-报表中心", "积分列表_temp.xlsx")
    working_temp = os.path.join(base_dir, "4-报表中心", "积分列表_working.xlsx")
    script_path = os.path.join(os.path.dirname(__file__), "积分列表.py")
    
    # 确保目录存在
    os.makedirs(os.path.dirname(points_file), exist_ok=True)
    
    try:
        # 步骤1：获取日期范围（快速方式）
        start_date, end_date = get_points_max_date(points_file)
        if not start_date or not end_date:
            yesterday = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
            start_date = yesterday
            end_date = yesterday
        
    # 步骤2：执行下载脚本，增加超时时间
        result = subprocess.run(
            [sys.executable, script_path, base_dir],
            capture_output=True,
            text=True,
            timeout=600  # 延长到10分钟
        )
        
        stdout = result.stdout.strip()
        stderr = result.stderr.strip()
        
        if result.returncode != 0 or "SUCCESS" not in stdout:
            error_msg = stdout if stdout else stderr
            return False, f"下载失败: {error_msg}"
        
        # 步骤3：脚本已成功执行，数据已写入数据库和Excel，直接返回成功
        # 解析输出获取新增记录数
        added_rows = 0
        for line in stdout.split('\n'):
            if '插入' in line and '条记录' in line:
                import re
                match = re.search(r'插入\s*(\d+)\s*条记录', line)
                if match:
                    added_rows = int(match.group(1))
                    break
        
        if added_rows > 0:
            return True, f"成功更新积分列表，新增 {added_rows} 条记录"
        else:
            return True, "积分列表更新完成，无新增数据"
    
    except subprocess.TimeoutExpired:
        return False, "操作超时（10分钟）"
    except Exception as e:
        return False, f"更新出错: {str(e)}"

def add_total_row_to_export(df):
    """
    在 DataFrame 末尾添加一行汇总行，对指定数字字段求和。
    返回新的 DataFrame（用于导出）。
    """
    if df.empty:
        return df

    # 需要求和的列（业务上有意义的度量字段）
    sum_cols = [
         "用电量",
        "收入_含税", "收入_不含税",
        "成本_含税", "成本_不含税"
    ]
    # 单价字段不求和（在汇总行中置为空字符串）
    skip_cols = ["开始表底", "结束表底","收入单价", "成本单价"]

    total_row = {}
    for col in df.columns:
        if col in sum_cols:
            # 确保转为数字再求和
            total_row[col] = pd.to_numeric(df[col], errors='coerce').sum()
        elif col in skip_cols:
            total_row[col] = ""   # 单价不汇总
        else:
            # 文本类字段（如电表状态、表号、品牌等）在汇总行填写 "合计"
            total_row[col] = "合计" if col == df.columns[0] else ""

    # 将汇总行追加到原 DataFrame 末尾
    total_df = pd.DataFrame([total_row])
    return pd.concat([df, total_df], ignore_index=True)

def generate_voucher_from_carryover(carryover_df, carryover_date, template_path):
    """
    根据结转明细（不含税金额）和模板文件，生成凭证模板 Excel 文件（内存中）
    参数：
        carryover_df: DataFrame，必须包含以下列：表号、品牌、客户编号、收入_不含税、成本_不含税
        carryover_date: datetime.date 对象，结账日期
        template_path: str，模板文件完整路径
    返回：
        BytesIO 对象，可直接用于 st.download_button
    """
    from openpyxl import load_workbook
    from io import BytesIO
    from datetime import datetime

    if carryover_df.empty:
        return None

    # 加载模板
    wb = load_workbook(template_path)
    target_sheet = "凭证模板"
    if target_sheet not in wb.sheetnames:
        ws = wb.create_sheet(target_sheet)
    else:
        ws = wb[target_sheet]

    # --- 清空第2行及以下的所有数据（保留第1行表头）---
    max_row = ws.max_row
    if max_row >= 2:
        for row in range(2, max_row + 1):
            for col in range(1, 39):  # 最多到第38列，足够覆盖VBA中的列
                ws.cell(row=row, column=col).value = None

    # 准备写入，从第2行开始
    next_row = 2

    # 准备日期对象（datetime 类型，时间部分为 00:00:00）
    dt_date = datetime.combine(carryover_date, datetime.min.time())
    month_num = carryover_date.month

    # 固定值
    company_code = "TZ001"
    voucher_type = "转"
    income_voucher_no = "转0001"
    cost_voucher_no = "转0002"
    account_receivable = "2241060301"   # 代收电费（其他应付款）
    account_income = "6051050101"       # 主营业务收入-电费
    account_cost = "6402020101"         # 主营业务成本-电费
    account_payable = "220209"          # 应付账款-电费
    dept_code = "BB01"
    project_code = "01.TZ001"
    contract_code = "05.TZ001.0"
    supplier_code = "ZJ020100029542"

    # ---- 第一组：代收电费科目分录（每个租户一行，金额使用收入_不含税）----
    for idx, row in carryover_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        revenue_notax = float(row["收入_不含税"])

        summary = f"{month_num}月租户电费收入结转"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=dt_date)               # 记账日期
        ws.cell(row=next_row, column=3, value=dt_date)               # 业务日期
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=income_voucher_no)
        ws.cell(row=next_row, column=7, value=1)
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_receivable)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=11, value=1)
        ws.cell(row=next_row, column=12, value=1)
        ws.cell(row=next_row, column=13, value=revenue_notax)        # 借方
        ws.cell(row=next_row, column=14, value=0)
        ws.cell(row=next_row, column=15, value=0)
        ws.cell(row=next_row, column=16, value=revenue_notax)        # 贷方
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)

        next_row += 1

    # ---- 第二组：电费租户收入科目分录（每个租户一行，金额使用收入_不含税）----
    for idx, row in carryover_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        revenue_notax = float(row["收入_不含税"])
        summary = f"{month_num}月租户电费收入结转"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=dt_date)               # 记账日期
        ws.cell(row=next_row, column=3, value=dt_date)               # 业务日期
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=income_voucher_no)
        ws.cell(row=next_row, column=7, value=2)                     # 分录序号2
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_income)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=11, value=1)
        ws.cell(row=next_row, column=12, value=0)
        ws.cell(row=next_row, column=13, value=revenue_notax)        # 借方
        ws.cell(row=next_row, column=14, value=0)
        ws.cell(row=next_row, column=15, value=0)
        ws.cell(row=next_row, column=17, value=revenue_notax)        # 贷方（第17列）
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)
        ws.cell(row=next_row, column=37, value="项目")
        ws.cell(row=next_row, column=38, value=project_code)

        next_row += 1

    # ---- 第三组：电费租户成本科目分录（每个租户一行，金额使用成本_不含税）----
    total_cost_notax = 0.0
    for idx, row in carryover_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        cost_notax = float(row["成本_不含税"])
        total_cost_notax += cost_notax
        summary = f"{month_num}月租户电费成本结转"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=dt_date)               # 记账日期
        ws.cell(row=next_row, column=3, value=dt_date)               # 业务日期
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=cost_voucher_no)
        ws.cell(row=next_row, column=7, value=1)                     # 分录序号1
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_cost)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=11, value=1)
        ws.cell(row=next_row, column=12, value=1)
        ws.cell(row=next_row, column=13, value=cost_notax)           # 借方
        ws.cell(row=next_row, column=14, value=0)
        ws.cell(row=next_row, column=15, value=0)
        ws.cell(row=next_row, column=16, value=cost_notax)           # 贷方
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)
        ws.cell(row=next_row, column=37, value="项目")
        ws.cell(row=next_row, column=38, value=project_code)

        next_row += 1

    # ---- 第四组：应付供应商科目分录（一行汇总，金额使用成本_不含税合计）----
    summary = f"{month_num}月租户电费成本结转"
    ws.cell(row=next_row, column=1, value=company_code)
    ws.cell(row=next_row, column=2, value=dt_date)                   # 记账日期
    ws.cell(row=next_row, column=3, value=dt_date)                   # 业务日期
    ws.cell(row=next_row, column=4, value=month_num)
    ws.cell(row=next_row, column=5, value=voucher_type)
    ws.cell(row=next_row, column=6, value=cost_voucher_no)
    ws.cell(row=next_row, column=7, value=2)                         # 分录序号2
    ws.cell(row=next_row, column=8, value=summary)
    ws.cell(row=next_row, column=9, value=account_payable)
    ws.cell(row=next_row, column=10, value=dept_code)
    ws.cell(row=next_row, column=11, value=1)
    ws.cell(row=next_row, column=12, value=0)
    ws.cell(row=next_row, column=13, value=total_cost_notax)         # 借方
    ws.cell(row=next_row, column=14, value=0)
    ws.cell(row=next_row, column=15, value=0)
    ws.cell(row=next_row, column=17, value=total_cost_notax)         # 贷方（第17列）
    ws.cell(row=next_row, column=33, value=summary)
    ws.cell(row=next_row, column=34, value="供应商")
    ws.cell(row=next_row, column=35, value=supplier_code)
    ws.cell(row=next_row, column=37, value="合同号")
    ws.cell(row=next_row, column=38, value=contract_code)

    # 保存到内存
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
def generate_adjust_voucher(adjust_df, adjust_period, template_path):
    """
    根据单价调整后的数据生成凭证模板（电费成本调整）
    凭证日期使用调整期间次月的最后一天
    """
    from openpyxl import load_workbook
    from io import BytesIO
    from datetime import datetime
    from calendar import monthrange

    if adjust_df.empty:
        return None

    # 加载模板
    wb = load_workbook(template_path)
    target_sheet = "凭证模板"
    if target_sheet not in wb.sheetnames:
        ws = wb.create_sheet(target_sheet)
    else:
        ws = wb[target_sheet]

    # 清空第2行及以下的所有数据
    max_row = ws.max_row
    if max_row >= 2:
        for row in range(2, max_row + 1):
            for col in range(1, 39):
                ws.cell(row=row, column=col).value = None

    next_row = 2

    # 解析调整期间，计算次月最后一天
    year, month = map(int, adjust_period.split('-'))
    if month == 12:
        next_year, next_month = year + 1, 1
    else:
        next_year, next_month = year, month + 1
    last_day = monthrange(next_year, next_month)[1]
    voucher_date = datetime(next_year, next_month, last_day)
    month_num = next_month   # 使用次月月份，与凭证日期(B/C列)的月份保持一致

    # 固定值
    company_code = "TZ001"
    voucher_type = "转"
    voucher_no = "转0002"
    account_cost = "6402020101"
    account_payable = "220209"
    dept_code = "BB01"
    project_code = "01.TZ001"
    contract_code = "05.TZ001.0"
    supplier_code = "ZJ020100029542"

    total_cost = 0.0

    # 循环生成租户成本分录
    for _, row in adjust_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        cost_notax = float(row["新成本_不含税"])
        total_cost += cost_notax

        summary = f"{month_num}月租户电费成本结转调整"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=voucher_date)
        ws.cell(row=next_row, column=3, value=voucher_date)
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=voucher_no)
        ws.cell(row=next_row, column=7, value=1)
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_cost)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=11, value=1)
        ws.cell(row=next_row, column=12, value=1)
        ws.cell(row=next_row, column=13, value=cost_notax)
        ws.cell(row=next_row, column=14, value=0)
        ws.cell(row=next_row, column=15, value=0)
        ws.cell(row=next_row, column=16, value=cost_notax)
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}-调整")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)
        ws.cell(row=next_row, column=37, value="项目")
        ws.cell(row=next_row, column=38, value=project_code)

        next_row += 1

    # 应付供应商汇总分录
    summary_total = f"{month_num}月租户电费成本结转调整"
    ws.cell(row=next_row, column=1, value=company_code)
    ws.cell(row=next_row, column=2, value=voucher_date)
    ws.cell(row=next_row, column=3, value=voucher_date)
    ws.cell(row=next_row, column=4, value=month_num)
    ws.cell(row=next_row, column=5, value=voucher_type)
    ws.cell(row=next_row, column=6, value=voucher_no)
    ws.cell(row=next_row, column=7, value=2)
    ws.cell(row=next_row, column=8, value=summary_total)
    ws.cell(row=next_row, column=9, value=account_payable)
    ws.cell(row=next_row, column=10, value=dept_code)
    ws.cell(row=next_row, column=11, value=1)
    ws.cell(row=next_row, column=12, value=0)
    ws.cell(row=next_row, column=13, value=total_cost)
    ws.cell(row=next_row, column=14, value=0)
    ws.cell(row=next_row, column=15, value=0)
    ws.cell(row=next_row, column=17, value=total_cost)
    ws.cell(row=next_row, column=33, value=summary_total)
    ws.cell(row=next_row, column=34, value="供应商")
    ws.cell(row=next_row, column=35, value=supplier_code)
    ws.cell(row=next_row, column=37, value="合同号")
    ws.cell(row=next_row, column=38, value=contract_code)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============== 新增：辅助函数（下载表底后追加数据） ==============
def append_meter_readings_from_downloaded_file(excel_main_path, downloaded_xls_path):
    """
    读取下载的能耗表底.xls，提取A列(表号)和D列(表底)，
    追加到能耗_电费数据.xlsx的“表底”工作表中，
    B列填充当前日期（Excel日期格式），显示格式为 yyyy/m/d。
    """
    if not os.path.exists(downloaded_xls_path):
        raise FileNotFoundError(f"未找到下载的文件：{downloaded_xls_path}")

    # 读取下载的xls文件（注意：可能有表头）
    df_download = pd.read_excel(downloaded_xls_path, header=None, dtype=str)
    if df_download.shape[1] < 4:
        raise ValueError("下载的表底文件列数不足4列，无法提取A、D列")

    # 提取A列（索引0）和 D列（索引3）
    meter_no = df_download.iloc[:, 0].astype(str).str.strip()
    meter_value = df_download.iloc[:, 3].astype(str).str.replace(',', '', regex=False).str.strip()
    meter_value = pd.to_numeric(meter_value, errors='coerce')

    # 去掉空行或无效值
    valid = meter_no.notna() & (meter_no != '') & (meter_value.notna())
    meter_no = meter_no[valid]
    meter_value = meter_value[valid]

    if meter_no.empty:
        raise ValueError("下载的表底文件中没有有效的表号或表底数据")

    # 当前日期对象（Excel 日期值）
    from datetime import date
    today_date = date.today()

    # 准备写入的数据：表号, 日期对象, 表底
    new_rows = []
    for no, val in zip(meter_no, meter_value):
        new_rows.append([no, today_date, val])

    # 使用 openpyxl 追加到主文件的“表底”工作表
    from openpyxl import load_workbook
    from openpyxl.styles import numbers

    if not os.path.exists(excel_main_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "表底"
    else:
        wb = load_workbook(excel_main_path)

    if "表底" not in wb.sheetnames:
        ws = wb.create_sheet("表底")
        # 写入表头
        ws.cell(row=1, column=1, value="表号")
        ws.cell(row=1, column=2, value="日期")
        ws.cell(row=1, column=3, value="表底")
    else:
        ws = wb["表底"]

    # 找到最后一行（从第2行开始已有数据）
    max_row = ws.max_row
    start_row = max_row + 1 if max_row >= 1 else 2

    # 写入数据，并设置日期列的显示格式
    for i, row_data in enumerate(new_rows, start=start_row):
        ws.cell(row=i, column=1, value=row_data[0])          # 表号
        date_cell = ws.cell(row=i, column=2, value=row_data[1])  # 日期对象
        date_cell.number_format = 'yyyy/m/d'                 # 设置显示格式
        ws.cell(row=i, column=3, value=row_data[2])          # 表底

    wb.save(excel_main_path)
    wb.close()
    return len(new_rows)
def append_single_meter_reading(excel_path, table_no, reading_date, meter_value):
    """
    将单条表底记录追加到能耗_电费数据.xlsx的“表底”工作表。
    参数：
        excel_path: 主数据文件路径
        table_no: 表号（字符串）
        reading_date: datetime.date 对象（申请日期）
        meter_value: 表底数值（float）
    """
    from openpyxl import load_workbook
    from openpyxl.styles import numbers

    # 打开或创建工作簿
    if not os.path.exists(excel_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "表底"
    else:
        wb = load_workbook(excel_path)

    # 确保“表底”工作表存在
    if "表底" not in wb.sheetnames:
        ws = wb.create_sheet("表底")
        ws.cell(row=1, column=1, value="表号")
        ws.cell(row=1, column=2, value="日期")
        ws.cell(row=1, column=3, value="表底")
    else:
        ws = wb["表底"]

    # 追加新行
    max_row = ws.max_row
    next_row = max_row + 1 if max_row >= 1 else 2

    ws.cell(row=next_row, column=1, value=str(table_no))
    date_cell = ws.cell(row=next_row, column=2, value=reading_date)
    date_cell.number_format = 'yyyy/m/d'   # 显示格式
    ws.cell(row=next_row, column=3, value=float(meter_value))

    wb.save(excel_path)
    wb.close()

def run_download_script(script_name, timeout_sec=120):
    """通用函数：执行同目录下的指定Python脚本"""
    script_path = os.path.join(os.path.dirname(__file__), script_name)
    if not os.path.exists(script_path):
        return False, f"脚本不存在：{script_path}"
    try:
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=timeout_sec
        )
        if result.returncode == 0:
            return True, "执行成功"
        else:
            return False, f"脚本执行失败，返回码{result.returncode}，错误：{result.stderr.strip()}"
    except subprocess.TimeoutExpired:
        return False, f"执行超时（{timeout_sec}秒）"
    except Exception as e:
        return False, f"调用脚本出错：{e}"
def generate_water_voucher_from_carryover(carryover_df, carryover_date, template_path):
    """
    根据水费结转明细生成凭证模板 Excel 文件（税率 9%）
    """
    from openpyxl import load_workbook
    from io import BytesIO
    from datetime import datetime

    if carryover_df.empty:
        return None

    wb = load_workbook(template_path)
    target_sheet = "凭证模板"
    if target_sheet not in wb.sheetnames:
        ws = wb.create_sheet(target_sheet)
    else:
        ws = wb[target_sheet]

    # 清空第2行及以下的所有数据
    max_row = ws.max_row
    if max_row >= 2:
        for row in range(2, max_row + 1):
            for col in range(1, 39):
                ws.cell(row=row, column=col).value = None

    next_row = 2
    dt_date = datetime.combine(carryover_date, datetime.min.time())
    month_num = carryover_date.month

    # 固定值（水费科目建议与电费区分，此处暂时沿用，用户可自行修改模板）
    company_code = "TZ001"
    voucher_type = "转"
    income_voucher_no = "转0001"
    cost_voucher_no = "转0002"
    account_receivable = "2241060302"   # 代收水费（其他应付款）
    account_income = "6051050102"       # 主营业务收入-水费（示例，可改）
    account_cost = "6402020102"         # 主营业务成本-水费
    account_payable = "220209"          # 应付账款-水费
    dept_code = "BB01"
    project_code = "01.TZ001"
    contract_code = "05.TZ001.0"
    supplier_code = "ZJ020100029542"    # 水费供应商

    # 第一组：代收水费科目分录
    for _, row in carryover_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        revenue_notax = float(row["收入_不含税"])
        summary = f"{month_num}月租户水费收入结转"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=dt_date)
        ws.cell(row=next_row, column=3, value=dt_date)
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=income_voucher_no)
        ws.cell(row=next_row, column=7, value=1)
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_receivable)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=11, value=1)
        ws.cell(row=next_row, column=12, value=1)
        ws.cell(row=next_row, column=13, value=revenue_notax)   # 借方
        ws.cell(row=next_row, column=16, value=revenue_notax)   # 贷方
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)
        next_row += 1

    # 第二组：水费收入科目分录
    for _, row in carryover_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        revenue_notax = float(row["收入_不含税"])
        summary = f"{month_num}月租户水费收入结转"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=dt_date)
        ws.cell(row=next_row, column=3, value=dt_date)
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=income_voucher_no)
        ws.cell(row=next_row, column=7, value=2)
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_income)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=13, value=revenue_notax)   # 借方
        ws.cell(row=next_row, column=17, value=revenue_notax)   # 贷方
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)
        ws.cell(row=next_row, column=37, value="项目")
        ws.cell(row=next_row, column=38, value=project_code)
        next_row += 1

    # 第三组：水费成本科目分录（每个租户）
    total_cost_notax = 0.0
    for _, row in carryover_df.iterrows():
        table_no = str(row["表号"])
        brand = str(row["品牌"])
        customer_id = str(row["客户编号"])
        cost_notax = float(row["成本_不含税"])
        total_cost_notax += cost_notax
        summary = f"{month_num}月租户水费成本结转"

        ws.cell(row=next_row, column=1, value=company_code)
        ws.cell(row=next_row, column=2, value=dt_date)
        ws.cell(row=next_row, column=3, value=dt_date)
        ws.cell(row=next_row, column=4, value=month_num)
        ws.cell(row=next_row, column=5, value=voucher_type)
        ws.cell(row=next_row, column=6, value=cost_voucher_no)
        ws.cell(row=next_row, column=7, value=1)
        ws.cell(row=next_row, column=8, value=summary)
        ws.cell(row=next_row, column=9, value=account_cost)
        ws.cell(row=next_row, column=10, value=dept_code)
        ws.cell(row=next_row, column=13, value=cost_notax)   # 借方
        ws.cell(row=next_row, column=16, value=cost_notax)   # 贷方
        ws.cell(row=next_row, column=33, value=f"{table_no}-{brand}")
        ws.cell(row=next_row, column=34, value="客户")
        ws.cell(row=next_row, column=35, value=customer_id)
        ws.cell(row=next_row, column=37, value="项目")
        ws.cell(row=next_row, column=38, value=project_code)
        next_row += 1

    # 第四组：应付供应商汇总分录
    summary = f"{month_num}月租户水费成本结转"
    ws.cell(row=next_row, column=1, value=company_code)
    ws.cell(row=next_row, column=2, value=dt_date)
    ws.cell(row=next_row, column=3, value=dt_date)
    ws.cell(row=next_row, column=4, value=month_num)
    ws.cell(row=next_row, column=5, value=voucher_type)
    ws.cell(row=next_row, column=6, value=cost_voucher_no)
    ws.cell(row=next_row, column=7, value=2)
    ws.cell(row=next_row, column=8, value=summary)
    ws.cell(row=next_row, column=9, value=account_payable)
    ws.cell(row=next_row, column=10, value=dept_code)
    ws.cell(row=next_row, column=13, value=total_cost_notax)   # 借方
    ws.cell(row=next_row, column=17, value=total_cost_notax)   # 贷方
    ws.cell(row=next_row, column=33, value=summary)
    ws.cell(row=next_row, column=34, value="供应商")
    ws.cell(row=next_row, column=35, value=supplier_code)
    ws.cell(row=next_row, column=37, value="合同号")
    ws.cell(row=next_row, column=38, value=contract_code)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
def append_meter_readings_from_downloaded_file_water(excel_main_path, downloaded_xls_path):
    """
    读取下载的能耗表底.xls中的“水”sheet，提取A列(表号)和D列(表底)，
    追加到能耗_水费数据.xlsx的“表底”工作表中，
    B列填充当前日期（Excel日期格式），显示格式为 yyyy/m/d。
    """
    if not os.path.exists(downloaded_xls_path):
        raise FileNotFoundError(f"未找到下载的文件：{downloaded_xls_path}")

    # 读取水 sheet（header=None，避免表头干扰）
    df_download = pd.read_excel(downloaded_xls_path, sheet_name='水', header=None, dtype=str)
    if df_download.shape[1] < 4:
        raise ValueError("下载的表底文件（水 sheet）列数不足4列，无法提取A、D列")

    meter_no = df_download.iloc[:, 0].astype(str).str.strip()
    meter_value = df_download.iloc[:, 3].astype(str).str.replace(',', '', regex=False).str.strip()
    meter_value = pd.to_numeric(meter_value, errors='coerce')

    valid = meter_no.notna() & (meter_no != '') & (meter_value.notna())
    meter_no = meter_no[valid]
    meter_value = meter_value[valid]

    if meter_no.empty:
        raise ValueError("下载的表底文件（水 sheet）中没有有效的表号或表底数据")

    today_date = date.today()
    new_rows = [[no, today_date, val] for no, val in zip(meter_no, meter_value)]

    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import numbers

    if not os.path.exists(excel_main_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "表底"
    else:
        wb = load_workbook(excel_main_path)

    if "表底" not in wb.sheetnames:
        ws = wb.create_sheet("表底")
        ws.cell(row=1, column=1, value="表号")
        ws.cell(row=1, column=2, value="日期")
        ws.cell(row=1, column=3, value="表底")
    else:
        ws = wb["表底"]

    max_row = ws.max_row
    start_row = max_row + 1 if max_row >= 1 else 2

    for i, row_data in enumerate(new_rows, start=start_row):
        ws.cell(row=i, column=1, value=row_data[0])
        date_cell = ws.cell(row=i, column=2, value=row_data[1])
        date_cell.number_format = 'yyyy/m/d'
        ws.cell(row=i, column=3, value=row_data[2])

    wb.save(excel_main_path)
    wb.close()
    return len(new_rows)
def apply_custom_css():
    st.markdown("""
    <style>
        :root {
            --primary-color: #1a365d;
            --primary-light: #2c5282;
            --primary-dark: #0d1f3c;
            --accent-color: #3182ce;
            --success-color: #38a169;
            --warning-color: #d69e2e;
            --danger-color: #e53e3e;
            --bg-main: #f7fafc;
            --bg-card: #ffffff;
            --bg-sidebar: #1a365d;
            --text-primary: #1a202c;
            --text-secondary: #4a5568;
            --text-light: #718096;
            --border-color: #e2e8f0;
            --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
            --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
            --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
            --radius-sm: 4px;
            --radius-md: 8px;
            --radius-lg: 12px;
        }

        /* 隐藏默认Streamlit顶部空白和菜单 */
        .stApp > header {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            padding: 0.75rem 1.5rem;
            margin: 0;
            border-bottom: 3px solid var(--accent-color);
        }

        .stApp {
            background: var(--bg-main);
        }

        .main .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
            background: var(--bg-main);
        }

        /* 确保侧边栏可见 */
        section[data-testid="stSidebar"],
        [data-testid="stSidebar"],
        div[data-testid="stSidebar"],
        .stSidebar {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        }

        /* 顶部企业标识栏 */
        .enterprise-header {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            color: white;
            padding: 1rem 1.5rem;
            border-radius: var(--radius-md);
            margin-bottom: 1.5rem;
            box-shadow: var(--shadow-md);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .enterprise-header h1 {
            font-size: 1.5rem;
            font-weight: 700;
            margin: 0;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        .enterprise-header .subtitle {
            font-size: 0.875rem;
            opacity: 0.9;
            font-weight: 400;
        }

        .enterprise-header .header-right {
            text-align: right;
            font-size: 0.8rem;
            opacity: 0.85;
        }

        /* 侧边栏整体样式 */
        section[data-testid="stSidebar"],
        [data-testid="stSidebar"],
        div[data-testid="stSidebar"] {
            background: linear-gradient(180deg, var(--bg-sidebar) 0%, #0f2440 100%) !important;
            border-right: 1px solid rgba(255, 255, 255, 0.1);
            padding-top: 1rem;
        }

        section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        div[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
            color: white;
        }

        /* 侧边栏导航标题 */
        .sidebar-nav-title {
            color: rgba(255, 255, 255, 0.6);
            font-size: 0.7rem;
            text-transform: uppercase;
            letter-spacing: 0.1em;
            padding: 0.75rem 1rem 0.5rem;
            font-weight: 600;
        }

        /* 侧边栏导航按钮 */
        section[data-testid="stSidebar"] .stButton > button,
        [data-testid="stSidebar"] .stButton > button,
        div[data-testid="stSidebar"] .stButton > button {
            width: 100%;
            background: transparent;
            border: 1px solid rgba(255, 255, 255, 0.15);
            color: rgba(255, 255, 255, 0.85);
            padding: 0.75rem 1rem;
            border-radius: var(--radius-md);
            margin: 0.25rem 0;
            font-size: 0.9rem;
            text-align: left;
            transition: all 0.25s ease;
            display: flex;
            align-items: center;
            gap: 0.75rem;
        }

        section[data-testid="stSidebar"] .stButton > button:hover,
        [data-testid="stSidebar"] .stButton > button:hover,
        div[data-testid="stSidebar"] .stButton > button:hover {
            background: rgba(255, 255, 255, 0.1);
            border-color: rgba(255, 255, 255, 0.3);
            color: white;
            transform: translateX(4px);
        }

        section[data-testid="stSidebar"] .stButton > button[disabled],
        [data-testid="stSidebar"] .stButton > button[disabled],
        div[data-testid="stSidebar"] .stButton > button[disabled] {
            background: linear-gradient(135deg, var(--accent-color) 0%, #2b6cb0 100%);
            border-color: var(--accent-color);
            color: white;
            font-weight: 600;
            cursor: default;
            box-shadow: 0 2px 8px rgba(49, 130, 206, 0.4);
        }

        section[data-testid="stSidebar"] .stButton > button[disabled]:hover,
        [data-testid="stSidebar"] .stButton > button[disabled]:hover,
        div[data-testid="stSidebar"] .stButton > button[disabled]:hover {
            transform: none;
        }

        /* Tab 样式优化 */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
            background: var(--bg-card);
            padding: 0.5rem;
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--border-color);
        }

        .stTabs [data-baseweb="tab"] {
            font-weight: 600;
            color: var(--text-secondary);
            border-radius: var(--radius-md);
            padding: 0.6rem 1.25rem;
            transition: all 0.25s ease;
            font-size: 0.9rem;
        }

        .stTabs [data-baseweb="tab"]:hover {
            background: rgba(49, 130, 206, 0.08);
            color: var(--accent-color);
        }

        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            color: white;
            box-shadow: var(--shadow-md);
        }

        /* 卡片容器样式 */
        .card {
            background: var(--bg-card);
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-md);
            border: 1px solid var(--border-color);
            padding: 1.25rem;
            margin-bottom: 1rem;
            transition: all 0.25s ease;
        }

        .card:hover {
            box-shadow: var(--shadow-lg);
        }

        .card-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 1rem;
            padding-bottom: 0.75rem;
            border-bottom: 1px solid var(--border-color);
        }

        .card-title {
            font-size: 1rem;
            font-weight: 700;
            color: var(--text-primary);
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        /* 指标卡片样式 */
        .metric-card {
            background: var(--bg-card);
            border-radius: var(--radius-lg);
            padding: 1.25rem;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--border-color);
            text-align: center;
            transition: all 0.25s ease;
            position: relative;
            overflow: hidden;
        }

        .metric-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, var(--accent-color), var(--primary-light));
        }

        .metric-card:hover {
            transform: translateY(-2px);
            box-shadow: var(--shadow-md);
        }

        .metric-value {
            font-size: 1.75rem;
            font-weight: 700;
            color: var(--primary-color);
            margin-bottom: 0.25rem;
        }

        .metric-label {
            font-size: 0.8rem;
            color: var(--text-light);
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .metric-change {
            font-size: 0.75rem;
            margin-top: 0.5rem;
            padding: 0.2rem 0.5rem;
            border-radius: 9999px;
            display: inline-block;
        }

        .metric-change.positive {
            background: rgba(56, 161, 105, 0.1);
            color: var(--success-color);
        }

        .metric-change.negative {
            background: rgba(229, 62, 62, 0.1);
            color: var(--danger-color);
        }

        /* 按钮样式 */
        .stButton > button {
            border-radius: var(--radius-md);
            font-weight: 600;
            transition: all 0.2s ease;
            padding: 0.5rem 1rem;
            border: none;
            box-shadow: var(--shadow-sm);
        }

        .stButton > button:hover {
            transform: translateY(-1px);
            box-shadow: var(--shadow-md);
        }

        .stButton > button:active {
            transform: translateY(0);
        }

        .stButton > button[kind="primary"],
        .stButton > button[data-qa="primary-button"] {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            color: white;
            border: none;
        }

        .stButton > button[kind="primary"]:hover,
        .stButton > button[data-qa="primary-button"]:hover {
            background: linear-gradient(135deg, var(--primary-light) 0%, var(--accent-color) 100%);
        }

        .stButton > button[kind="secondary"],
        .stButton > button:not([kind]) {
            background: var(--bg-card);
            color: var(--text-primary);
            border: 1px solid var(--border-color);
        }

        .stButton > button[kind="secondary"]:hover,
        .stButton > button:not([kind]):hover {
            background: var(--bg-main);
            border-color: var(--accent-color);
            color: var(--accent-color);
        }

        /* 成功按钮 */
        .success-btn > button {
            background: linear-gradient(135deg, var(--success-color) 0%, #2f855a 100%) !important;
            color: white !important;
        }

        /* 下载按钮样式 */
        .stDownloadButton > button {
            background: linear-gradient(135deg, var(--success-color) 0%, #2f855a 100%);
            color: white;
            border: none;
            border-radius: var(--radius-md);
            font-weight: 600;
        }

        .stDownloadButton > button:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(56, 161, 105, 0.4);
        }

        /* 表单输入框样式 */
        .stTextInput > div > div > input,
        .stNumberInput > div > div > input,
        .stDateInput > div > div > input,
        .stSelectbox > div > div > select,
        .stTextArea > div > div > textarea {
            border-radius: var(--radius-md);
            border: 1px solid var(--border-color);
            padding: 0.6rem 0.75rem;
            transition: all 0.2s ease;
            background: var(--bg-card);
        }

        .stTextInput > div > div > input:focus,
        .stNumberInput > div > div > input:focus,
        .stDateInput > div > div > input:focus,
        .stSelectbox > div > div > select:focus,
        .stTextArea > div > div > textarea:focus {
            border-color: var(--accent-color);
            box-shadow: 0 0 0 3px rgba(49, 130, 206, 0.15);
            outline: none;
        }

        /* 数据表格样式 */
        .dataframe {
            border-radius: var(--radius-md);
            overflow: hidden;
            box-shadow: var(--shadow-sm);
            border: 1px solid var(--border-color);
            font-size: 0.875rem;
        }

        .dataframe thead {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            color: white;
        }

        .dataframe thead th {
            font-weight: 600;
            padding: 0.75rem;
            text-align: left;
        }

        .dataframe tbody tr {
            transition: background 0.15s ease;
        }

        .dataframe tbody tr:nth-child(even) {
            background: rgba(49, 130, 206, 0.03);
        }

        .dataframe tbody tr:hover {
            background: rgba(49, 130, 206, 0.08);
        }

        .dataframe td {
            padding: 0.6rem 0.75rem;
        }

        /* 分隔线样式 */
        hr {
            border: none;
            height: 1px;
            background: var(--border-color);
            margin: 1.5rem 0;
        }

        /* 消息提示框样式 */
        .stAlert {
            border-radius: var(--radius-md);
            border-left-width: 4px;
            box-shadow: var(--shadow-sm);
        }

        .stAlert[data-baseweb="notification"] {
            border-radius: var(--radius-md);
        }

        /* 进度条样式 */
        .stProgress > div > div > div {
            background: linear-gradient(90deg, var(--accent-color), var(--primary-light));
            border-radius: 9999px;
        }

        /* 展开器样式 */
        .streamlit-expanderHeader {
            background: var(--bg-card);
            border-radius: var(--radius-md);
            border: 1px solid var(--border-color);
            font-weight: 600;
        }

        .streamlit-expanderHeader:hover {
            background: var(--bg-main);
        }

        .streamlit-expanderContent {
            background: var(--bg-card);
            border: 1px solid var(--border-color);
            border-top: none;
            border-radius: 0 0 var(--radius-md) var(--radius-md);
        }

        /* 滑块样式 */
        .stSlider > div > div > div > div {
            background: linear-gradient(90deg, var(--accent-color), var(--primary-light));
        }

        /* 多选框样式 */
        .stMultiSelect > div > div {
            border-radius: var(--radius-md);
        }

        /* 单选按钮样式 */
        .stRadio > div {
            background: var(--bg-card);
            padding: 0.5rem;
            border-radius: var(--radius-md);
            border: 1px solid var(--border-color);
        }

        .stRadio label {
            padding: 0.4rem 0.75rem;
            border-radius: var(--radius-sm);
            transition: all 0.2s ease;
        }

        /* 分栏布局间距优化 */
        [data-testid="column"] {
            padding: 0.25rem;
        }

        /* 空状态样式 */
        .empty-state {
            text-align: center;
            padding: 3rem 2rem;
            color: var(--text-light);
        }

        .empty-state-icon {
            font-size: 3rem;
            margin-bottom: 1rem;
            opacity: 0.5;
        }

        /* 状态标签 */
        .status-badge {
            display: inline-flex;
            align-items: center;
            padding: 0.25rem 0.75rem;
            border-radius: 9999px;
            font-size: 0.75rem;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.03em;
        }

        .status-badge.success {
            background: rgba(56, 161, 105, 0.1);
            color: var(--success-color);
        }

        .status-badge.warning {
            background: rgba(214, 158, 46, 0.1);
            color: var(--warning-color);
        }

        .status-badge.danger {
            background: rgba(229, 62, 62, 0.1);
            color: var(--danger-color);
        }

        .status-badge.info {
            background: rgba(49, 130, 206, 0.1);
            color: var(--accent-color);
        }

        /* 表格工具栏 */
        .table-toolbar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
            padding: 0.75rem 1rem;
            background: var(--bg-card);
            border-radius: var(--radius-md);
            border: 1px solid var(--border-color);
        }

        .table-info {
            font-size: 0.85rem;
            color: var(--text-secondary);
        }

        /* 加载动画 */
        .loading-spinner {
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 2rem;
        }

        /* 步骤指示器 */
        .step-indicator {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 0.5rem;
            margin: 1.5rem 0;
        }

        .step {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .step-number {
            width: 28px;
            height: 28px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.8rem;
            font-weight: 600;
            background: var(--border-color);
            color: var(--text-light);
        }

        .step.active .step-number {
            background: linear-gradient(135deg, var(--accent-color), var(--primary-light));
            color: white;
        }

        .step.completed .step-number {
            background: var(--success-color);
            color: white;
        }

        .step-label {
            font-size: 0.85rem;
            color: var(--text-light);
        }

        .step.active .step-label {
            color: var(--text-primary);
            font-weight: 600;
        }

        /* 弹窗/模态框背景 */
        .modal-overlay {
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 1000;
        }

        .modal-content {
            background: var(--bg-card);
            border-radius: var(--radius-lg);
            padding: 1.5rem;
            max-width: 500px;
            width: 90%;
            box-shadow: var(--shadow-lg);
        }

        /* 工具提示 */
        .tooltip {
            position: relative;
            display: inline-flex;
        }

        .tooltip-text {
            visibility: hidden;
            background: var(--primary-color);
            color: white;
            padding: 0.4rem 0.75rem;
            border-radius: var(--radius-sm);
            font-size: 0.75rem;
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            white-space: nowrap;
            z-index: 100;
        }

        .tooltip:hover .tooltip-text {
            visibility: visible;
        }

        /* 面包屑导航 */
        .breadcrumb {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            font-size: 0.85rem;
            color: var(--text-light);
            margin-bottom: 1rem;
        }

        .breadcrumb-item {
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .breadcrumb-item::after {
            content: '>';
            color: var(--border-color);
        }

        .breadcrumb-item:last-child::after {
            content: '';
        }

        .breadcrumb-item.active {
            color: var(--text-primary);
            font-weight: 600;
        }

        /* 页脚样式 */
        .footer {
            text-align: center;
            margin-top: 3rem;
            padding-top: 1.5rem;
            border-top: 1px solid var(--border-color);
            font-size: 0.8rem;
            color: var(--text-light);
        }

        /* 隐藏默认Streamlit元素 */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stToolbar"] {visibility: hidden;}

        /* 滚动条样式 */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }

        ::-webkit-scrollbar-track {
            background: var(--bg-main);
            border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb {
            background: var(--border-color);
            border-radius: 4px;
        }

        ::-webkit-scrollbar-thumb:hover {
            background: var(--text-light);
        }

        /* 选中文字样式 */
        ::selection {
            background: rgba(49, 130, 206, 0.3);
        }

        /* 表单分组 */
        .form-group {
            background: var(--bg-card);
            border-radius: var(--radius-lg);
            padding: 1.25rem;
            margin-bottom: 1rem;
            border: 1px solid var(--border-color);
        }

        .form-group-title {
            font-size: 0.9rem;
            font-weight: 700;
            color: var(--text-primary);
            margin-bottom: 1rem;
            padding-bottom: 0.5rem;
            border-bottom: 2px solid var(--accent-color);
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        /* 操作按钮组 */
        .button-group {
            display: flex;
            gap: 0.75rem;
            margin: 1rem 0;
        }

        /* 搜索框样式 */
        .search-box {
            position: relative;
        }

        .search-box input {
            padding-left: 2.5rem;
        }

        .search-box::before {
            content: '🔍';
            position: absolute;
            left: 0.75rem;
            top: 50%;
            transform: translateY(-50%);
            font-size: 0.9rem;
            opacity: 0.5;
        }

        /* 统计概览卡片 */
        .stats-overview {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 1rem;
            margin-bottom: 1.5rem;
        }

        /* 时间选择器特殊样式 */
        input[type="date"] {
            border-radius: var(--radius-md);
        }

        /* 响应式调整 */
        @media (max-width: 768px) {
            .enterprise-header {
                flex-direction: column;
                text-align: center;
                gap: 0.5rem;
            }

            .enterprise-header .header-right {
                text-align: center;
            }

            .stats-overview {
                grid-template-columns: 1fr;
            }
        }
    </style>
    """, unsafe_allow_html=True)

# ============== 主界面 ==============
def main():
    st.set_page_config(
        page_title="财务共享服务中心",
        page_icon="💹",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    apply_custom_css()

    # ---------- 企业级顶部标题栏 ----------
    st.markdown("""
    <div class="enterprise-header">
        <div>
            <h1>🏢 财务共享服务中心</h1>
        </div>
        <div class="header-right">
            <div>唐山中骏商业管理有限公司</div>
            <div style="font-size: 0.75rem; margin-top: 0.25rem; opacity: 0.8;">财务ERP系统 v2.0</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ---------- 导航状态初始化 ----------
    if 'current_page' not in st.session_state:
        st.session_state.current_page = "能耗系统-电费"

    # ---------- 侧边栏：专业导航 ----------
    st.sidebar.markdown('<div class="sidebar-nav-title">📋 功能菜单</div>', unsafe_allow_html=True)

    # 所有页面配置（以后增加功能只需在这里加一行）
    page_options = {
        "能耗系统-电费": "⚡ 能耗系统-电费",
        "能耗系统-水费": "💧 能耗系统-水费",
        "月末结转": "📆 月末结转-电费",
        "数据看板": "📊 数据看板",
        "数据更新": "🔄 数据更新",
    }

    for page_key, page_label in page_options.items():
        is_current = (st.session_state.current_page == page_key)
        if is_current:
            # 当前页按钮：显示但不执行任何操作（纯指示作用）
            st.sidebar.button(
                page_label,
                key=f"nav_{page_key}",
                use_container_width=True,
                disabled=True  # 禁用点击，纯粹显示
            )
        else:
            if st.sidebar.button(
                    page_label,
                    key=f"nav_{page_key}",
                    use_container_width=True
            ):
                st.session_state.current_page = page_key
                st.rerun()

    # 从 session_state 获取当前页面
    current_page = st.session_state.current_page

    st.sidebar.markdown("---")
    st.sidebar.markdown("""
    <div style="text-align: center; padding: 1rem 0; color: rgba(255,255,255,0.5); font-size: 0.7rem;">
        <div style="margin-bottom: 0.25rem;">财务ERP系统 v2.0</div>
        <div style="opacity: 0.7;">© 2026 唐山中骏商业管理有限公司</div>
    </div>
    """, unsafe_allow_html=True)

    st.sidebar.markdown(
        """
        <style>
        section[data-testid="stSidebar"] .stButton > button[disabled],
        [data-testid="stSidebar"] .stButton > button[disabled],
        div[data-testid="stSidebar"] .stButton > button[disabled] {
            background: linear-gradient(135deg, #3182ce 0%, #2b6cb0 100%) !important;
            border-color: #3182ce !important;
            color: white !important;
            font-weight: bold;
            box-shadow: 0 2px 8px rgba(49, 130, 206, 0.4);
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    df_existing = load_excel(excel_path)

    # ========== 能耗系统-电费 ==========
    if current_page == "能耗系统-电费":
        tab1, tab2, tab3, tab_alert = st.tabs(
            ["📋清表录入", "⚡启用明细", "🧾模板转换", "⚠️余额预警"],
            key="energy_tabs"  # 任意唯一字符串
        )

        # ========== 标签页1：清表录入 ==========
        with tab1:
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                table_no = st.text_input("表号", placeholder="来自能耗系统", key="table_no")
            with c2:
                brand = st.text_input("品牌", placeholder="百购系统规范名称", key="brand")
            with c3:
                customer_id = st.text_input("客户编号", placeholder="金蝶系统规范编号", key="customer_id")
            with c4:
                apply_date = st.date_input("申请日期", value=date.today(), key="apply_date")
            with c5:
                meter_bottom = st.number_input("表底", min_value=0.0, step=0.1, format="%.2f", key="meter_bottom")

            # ---------- 清表录入区域的完整替换 ----------
            col_type, col_id, col_meter_type, col_contract = st.columns([1, 1, 1, 1])
            with col_type:
                clear_type = st.radio("清表类型", ["进场清表", "撤场清表"], horizontal=True, key="clear_type")

            # 自动生成自建ID（进场清表且表号、品牌、申请日期均存在时）
            if clear_type == "进场清表" and table_no and brand and apply_date:
                date_str = f"{apply_date.year}/{apply_date.month}/{apply_date.day}"
                auto_id = f"{table_no}{brand}{date_str}"
                st.session_state.id_input = auto_id
            else:
                if clear_type == "进场清表":
                    st.session_state.id_input = ""

            with col_id:
                st.text_input("自建ID（可修改）", key="id_input", placeholder="自动生成或手动填写")

            with col_meter_type:
                meter_type = st.selectbox(
                    "电表类型",
                    options=["预付费表", "非预付费表"],
                    key="meter_type"
                )

            with col_contract:
                contract_no = st.text_input("合同号", key="contract_no", placeholder="可选，可手动填写")

            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                submitted = st.button("提交数据", type="primary", use_container_width=True)
            with col_btn2:
                st.button("清空", on_click=clear_form, use_container_width=True)

            if submitted:
                if not table_no or not brand or not customer_id:
                    st.error("表号、品牌和客户编号不能为空！")
                else:
                    new_row = {
                        "表号": table_no,
                        "品牌": brand,
                        "客户编号": customer_id,
                        "申请日期": str(apply_date),
                        "表底": meter_bottom,
                        "清表类型": clear_type,
                        "自建ID": st.session_state.id_input,
                        "电表类型": st.session_state.meter_type,
                        "合同号": st.session_state.contract_no,
                    }
                    try:
                        new_df = pd.DataFrame([new_row])
                        if os.path.exists(excel_path):
                            df_cur = pd.read_excel(excel_path, sheet_name=0, dtype=str)
                            df_cur = process_date_column(df_cur)
                        else:
                            df_cur = pd.DataFrame(
                                columns=["表号", "品牌", "客户编号", "申请日期", "表底", "清表类型", "自建ID",
                                         "电表类型", "合同号"]
                            )
                        df_cur = pd.concat([df_cur, new_df], ignore_index=True)
                        # 确保列顺序
                        column_order = ["表号", "品牌", "客户编号", "申请日期", "表底", "清表类型", "自建ID",
                                        "电表类型", "合同号"]
                        df_cur = df_cur[column_order]

                        # 保存至Excel（此处保留原保存代码不变，用openpyxl写入）
                        from openpyxl import load_workbook
                        if os.path.exists(excel_path):
                            wb = load_workbook(excel_path)
                        else:
                            wb = Workbook()
                        if 'Sheet1' in wb.sheetnames:
                            ws = wb['Sheet1']
                            for row in ws.iter_rows():
                                for cell in row:
                                    cell.value = None
                        else:
                            ws = wb.active
                            ws.title = 'Sheet1'
                        for col_idx, col_name in enumerate(df_cur.columns, 1):
                            ws.cell(row=1, column=col_idx, value=col_name)
                        for r_idx, row in df_cur.iterrows():
                            for c_idx, value in enumerate(row):
                                ws.cell(row=r_idx + 2, column=c_idx + 1, value=value)
                        wb.save(excel_path)
                        wb.close()
                        load_excel.clear()
                        append_single_meter_reading(excel_path, table_no, apply_date, meter_bottom)
                        st.success("录入成功")
                        st.rerun()
                    except Exception as e:
                        st.error(f"保存失败：{e}")

            # 在清表录入 tab 内，筛选生成 filtered_df 之后
            st.markdown("---")
            st.subheader("📊 清表记录")

            # 应用筛选条件（原有代码）
            filtered_df = df_existing.copy()
            if table_no:
                filtered_df = filtered_df[filtered_df["表号"].str.contains(table_no, na=False)]
            if brand:
                filtered_df = filtered_df[filtered_df["品牌"].str.contains(brand, na=False)]
            if customer_id:
                filtered_df = filtered_df[filtered_df["客户编号"].str.contains(customer_id, na=False)]

            # ---------- 新增：行选择自动填充 ----------
            if clear_type == "撤场清表":
                # 显示可选择表格
                st.data_editor(
                    filtered_df,
                    use_container_width=True,
                    disabled=True,
                    key="clear_table_selector",
                    hide_index=True,
                    num_rows="fixed",
                )
                # 处理选中行
                selection = st.session_state.get("clear_table_selector", {}).get("selection", {})
                if selection and "rows" in selection and selection["rows"]:
                    selected_idx = selection["rows"][0]
                    if "last_filled_idx" not in st.session_state:
                        st.session_state.last_filled_idx = -1
                    if selected_idx != st.session_state.last_filled_idx:
                        st.session_state.last_filled_idx = selected_idx
                        if selected_idx < len(filtered_df):
                            row = filtered_df.iloc[selected_idx]
                            st.session_state.brand = str(row.get("品牌", ""))
                            st.session_state.customer_id = str(row.get("客户编号", ""))
                            st.session_state.id_input = str(row.get("自建ID", ""))
                            st.session_state.meter_type = str(row.get("电表类型", "预付费表"))
                            st.session_state.contract_no = str(row.get("合同号", ""))
                            st.session_state.table_no = str(row.get("表号", ""))
                            try:
                                st.session_state.meter_bottom = float(row.get("表底", 0.0))
                            except:
                                st.session_state.meter_bottom = 0.0
                            app_date = row.get("申请日期", "")
                            if app_date:
                                try:
                                    st.session_state.apply_date = pd.to_datetime(app_date).date()
                                except:
                                    pass
                            st.rerun()
            else:
                st.dataframe(filtered_df, use_container_width=True)

            st.caption(f"共 {len(filtered_df)} 条记录（总记录数：{len(df_existing)}）")

        # ========== 标签页2：启用明细 ==========
        with tab2:


            col_date, col_status, col_no, col_brand = st.columns([2, 2, 1.5, 1.5])
            with col_date:
                settlement_date = st.date_input("结账日期", value=date.today(), key="settlement_date")
            with col_status:
                status_placeholder = st.empty()
            with col_no:
                filter_table_no = st.text_input("表号", placeholder="模糊查询", key="filter_table_no")
            with col_brand:
                filter_brand = st.text_input("品牌", placeholder="模糊查询", key="filter_brand")

            if not df_existing.empty:
                from datetime import datetime

                df_in = df_existing[df_existing["清表类型"] == "进场清表"].copy()
                df_out = df_existing[df_existing["清表类型"] == "撤场清表"].copy()

                df_in["申请日期_dt"] = pd.to_datetime(df_in["申请日期"], errors="coerce")
                df_out["申请日期_dt"] = pd.to_datetime(df_out["申请日期"], errors="coerce")

                df_in_max = df_in.loc[df_in.groupby("自建ID")["申请日期_dt"].idxmax()]
                df_out_max = df_out.loc[df_out.groupby("自建ID")["申请日期_dt"].idxmax()]

                df_merged = pd.merge(
                    df_in_max[["自建ID", "表号", "品牌", "客户编号", "合同号", "申请日期_dt", "表底"]],
                    df_out_max[["自建ID", "申请日期_dt", "表底"]],
                    on="自建ID", how="outer", suffixes=("_in", "_out")
                )
                df_merged.rename(columns={
                    "申请日期_dt_in": "进场日期", "表底_in": "进场表底",
                    "申请日期_dt_out": "撤场日期", "表底_out": "撤场表底"
                }, inplace=True)

                settlement = pd.Timestamp(settlement_date)
                month_start = settlement.replace(day=1)

                def calc_status(row):
                    in_date = row["进场日期"]
                    out_date = row["撤场日期"]
                    if pd.isna(in_date):
                        return "无进场记录"
                    if in_date > settlement:
                        return "未启用"
                    if (in_date >= month_start) and (pd.notna(out_date) and out_date <= settlement):
                        return "本月调整"
                    if (in_date < month_start) and (pd.notna(out_date) and month_start <= out_date <= settlement):
                        return "本月终止"
                    if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start):
                        return "已终止"
                    if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement):
                        return "本月启用"
                    if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement):
                        return "启用中"
                    return "异常状态"

                df_merged["电表状态"] = df_merged.apply(calc_status, axis=1)
                df_result = df_merged[
                    ["自建ID", "表号", "品牌", "客户编号", "合同号", "进场日期", "进场表底", "撤场日期", "撤场表底", "电表状态"]
                ].copy()
                df_result["进场日期"] = df_result["进场日期"].dt.strftime("%Y-%m-%d")
                df_result["撤场日期"] = df_result["撤场日期"].dt.strftime("%Y-%m-%d")
                df_result = df_result.fillna("")

                df_result_full = df_result.copy()
                st.session_state.df_result_full = df_result_full.copy()

                enabled_df = df_result_full[df_result_full["电表状态"].isin(["启用中", "本月启用"])]
                st.session_state.enabled_df = enabled_df.copy()

                status_options = sorted(df_result["电表状态"].unique().tolist())
                default_statuses = [s for s in ["启用中", "本月启用"] if s in status_options]
                with status_placeholder:
                    selected_statuses = st.multiselect(
                        "电表状态（可多选）",
                        options=status_options,
                        default=default_statuses,
                        key="status_filter"
                    )

                if selected_statuses:
                    df_result = df_result[df_result["电表状态"].isin(selected_statuses)]
                if filter_table_no:
                    df_result = df_result[df_result["表号"].str.contains(filter_table_no, na=False)]
                if filter_brand:
                    df_result = df_result[df_result["品牌"].str.contains(filter_brand, na=False)]

                st.session_state.detail_df = df_result.copy()
                st.dataframe(df_result, use_container_width=True)
                st.caption(f"共 {len(df_result)} 条（结账日期：{settlement_date}）")
            else:
                st.info("暂无清表数据，请先在“清表录入”中添加记录。")
                empty_cols = ["自建ID", "表号", "品牌", "客户编号", "合同号", "进场日期", "进场表底", "撤场日期", "撤场表底", "电表状态"]
                st.session_state.df_result_full = pd.DataFrame(columns=empty_cols)
                st.session_state.enabled_df = pd.DataFrame(columns=["表号", "客户编号", "合同号"])

        # ========== 标签页3：模板转换 ==========
        with tab3:


            col_upload, col_btn, col_download = st.columns([4, 1.2, 2.5])
            with col_upload:
                uploaded_file = st.file_uploader(
                    "上传充值结算记录（.xls 或 .xlsx）",
                    type=["xls", "xlsx"],
                    key="template_upload"
                )
            with col_btn:
                st.write("")
                if st.button("📥 下载百购模板"):
                    if ('converted_df' not in st.session_state or
                            st.session_state.converted_df.empty):
                        st.error("请先上传文件并完成转换，再下载模板。")
                    else:
                        try:
                            src = st.session_state.converted_df
                            df4 = pd.DataFrame()
                            df4['金额'] = src['充值金额（元）']
                            df4['金蝶客商编码'] = src['客户编号']
                            df4['金蝶客商名称'] = src['客户名称']
                            df4['费用起日'] = src['订单时间']
                            df4['费用止日'] = src['订单时间']
                            df4['应收日'] = (pd.to_datetime(src['订单时间'], format='%Y%m%d') + pd.Timedelta(days=1)).dt.strftime('%Y%m%d')
                            df4['费用项名称'] = '1104 电费'
                            df4['主体公司名称'] = '唐山中骏商业管理有限公司'
                            df4['合同号'] = ''
                            df4['备注'] = ''
                            df4 = df4[['费用项名称', '合同号', '金额', '费用起日', '费用止日', '应收日',
                                       '主体公司名称', '金蝶客商编码', '金蝶客商名称', '备注']]

                            if not os.path.exists(bai_gou_template_path):
                                st.error(f"模板文件不存在：{bai_gou_template_path}")
                            else:
                                import xlrd
                                from xlutils.copy import copy as xl_copy
                                from io import BytesIO
                                from xlwt import XFStyle

                                rb = xlrd.open_workbook(bai_gou_template_path, formatting_info=True)
                                wb = xl_copy(rb)
                                ws = wb.get_sheet(0)

                                money_style = XFStyle()
                                money_style.num_format_str = '#,##0.00'

                                df_export = df4.copy()
                                df_export['金额'] = (
                                    df_export['金额'].astype(str)
                                    .str.replace(',', '', regex=False)
                                    .str.replace(' ', '', regex=False)
                                )
                                df_export['金额'] = pd.to_numeric(df_export['金额'], errors='coerce').fillna(0).astype(float)

                                for i, row in df_export.iterrows():
                                    for j, value in enumerate(row):
                                        if j == 2:
                                            ws.write(i + 2, j, value, money_style)
                                        else:
                                            ws.write(i + 2, j, str(value) if value is not None else '')

                                output = BytesIO()
                                wb.save(output)
                                output.seek(0)
                                st.session_state.download_data = output
                                st.session_state.download_ready = True
                        except ImportError:
                            st.error("缺少 xlutils 库，请安装后重试")
                        except Exception as e:
                            st.error(f"生成失败：{e}")

                if st.button("📥 下载开票模板"):
                    if ('converted_df' not in st.session_state or
                            st.session_state.converted_df.empty):
                        st.error("请先上传文件并完成转换，再下载开票模板。")
                    else:
                        try:
                            from openpyxl import load_workbook
                            from io import BytesIO

                            src = st.session_state.converted_df
                            amount_clean = (src['充值金额（元）'].astype(str)
                                            .str.replace(',', '', regex=False)
                                            .str.replace(' ', '', regex=False))
                            amount = pd.to_numeric(amount_clean, errors='coerce').fillna(0.0)
                            order_no = src['单号'].astype(str).str[-20:]

                            if not os.path.exists(kai_piao_template_path):
                                st.error(f"模板文件不存在：{kai_piao_template_path}")
                            else:
                                wb = load_workbook(kai_piao_template_path)
                                ws1 = wb['1-发票基本信息']
                                ws2 = wb['2-发票明细信息']

                                for i in range(len(src)):
                                    row_excel = i + 4
                                    ws1.cell(row=row_excel, column=1, value=order_no.iloc[i])
                                    ws1.cell(row=row_excel, column=2,
                                             value=str(src['发票票种'].iloc[i]) if pd.notna(src['发票票种'].iloc[i]) else '')
                                    ws1.cell(row=row_excel, column=4, value='是')
                                    ws1.cell(row=row_excel, column=6,
                                             value=str(src['购买方名称'].iloc[i]) if pd.notna(src['购买方名称'].iloc[i]) else '')
                                    ws1.cell(row=row_excel, column=7,
                                             value=str(src['购方识别号'].iloc[i]) if pd.notna(src['购方识别号'].iloc[i]) else '')
                                    ws1.cell(row=row_excel, column=31,
                                             value=str(src['邮箱'].iloc[i]) if pd.notna(src['邮箱'].iloc[i]) else '')
                                    ws2.cell(row=row_excel, column=1, value=order_no.iloc[i])
                                    ws2.cell(row=row_excel, column=2, value='电费')
                                    ws2.cell(row=row_excel, column=3, value='1100101020200000000')
                                    ws2.cell(row=row_excel, column=5, value='度')
                                    money = amount.iloc[i]
                                    ws2.cell(row=row_excel, column=6, value=round(money / 0.75, 10))
                                    ws2.cell(row=row_excel, column=8, value=money)
                                    ws2.cell(row=row_excel, column=9, value=0.13)

                                output = BytesIO()
                                wb.save(output)
                                output.seek(0)
                                st.session_state.kaipiao_download_data = output
                                st.session_state.kaipiao_download_ready = True
                                st.success("开票模板生成成功！")
                        except ImportError:
                            st.error("缺少 openpyxl 库，请安装后重试")
                        except Exception as e:
                            st.error(f"生成失败：{e}")

            with col_download:
                if st.session_state.get('download_ready'):
                    st.download_button(
                        label="💾 下载应收帐单导入.xls",
                        data=st.session_state.download_data,
                        file_name="应收帐单导入.xls",
                        mime="application/vnd.ms-excel"
                    )
                    st.caption("文件已生成，点击上方按钮即可下载。")
                if st.session_state.get('kaipiao_download_ready'):
                    st.download_button(
                        label="💾 下载批量开票模板.xlsx",
                        data=st.session_state.kaipiao_download_data,
                        file_name="批量开票-导入开票模板.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.caption("开票模板已生成，点击上方按钮即可下载。")

            if uploaded_file is not None:
                st.session_state.download_ready = False
                st.session_state.kaipiao_download_ready = False

                try:
                    df = pd.read_excel(uploaded_file, dtype={'表号': str})
                    df = df.drop(index=0).reset_index(drop=True)
                    df['订单时间'] = pd.to_datetime(df['订单时间']).dt.strftime('%Y%m%d')
                    df = df[['订单时间', '单号', '租户名称', '表号', '充值金额（元）', '付费途径']]

                    if 'enabled_df' not in st.session_state or st.session_state.enabled_df.empty:
                        st.error("请先在“启用明细”页面生成数据，再执行模板转换。")
                        st.stop()
                    df2 = st.session_state.enabled_df[['表号', '客户编号', '合同号']].drop_duplicates()

                    if not os.path.exists(customer_info_path):
                        st.error(f"缺少文件：{customer_info_path}")
                        st.stop()
                    df3 = pd.read_excel(customer_info_path)[['编码', '名称']]

                    for p in [zhengpu_contract_path, duojing_contract_path]:
                        if not os.path.exists(p):
                            st.error(f"缺少文件：{p}")
                            st.stop()

                    正铺 = pd.read_excel(zhengpu_contract_path, skiprows=1)
                    正铺['铺位类型'] = '正铺'
                    正铺 = 正铺[['铺位类型', '状态', '合同号', '供应商', '品牌', '业态', '细分品类',
                                 '合同\n开始日期', '合同\n结束日期', '终止日期']]
                    正铺.columns = ['铺位类型', '状态', '合同号', '供应商', '品牌', '业态', '细分品类',
                                    '合同开始日期', '合同结束日期', '终止日期']
                    for col in ['合同开始日期', '合同结束日期', '终止日期']:
                        正铺[col] = pd.to_datetime(正铺[col], errors='coerce').dt.strftime('%Y%m%d')
                    正铺['终止日期'] = 正铺['终止日期'].fillna(正铺['合同结束日期'])

                    多经 = pd.read_excel(duojing_contract_path)
                    多经['铺位类型'] = '多经'
                    多经 = 多经[['铺位类型', '合同状态', '合同编号', '供应商', '签约品牌', '点位类型', '点位类型',
                                 '合同开始日期', '合同结束日期', '终止日期']]
                    多经.columns = ['铺位类型', '状态', '合同号', '供应商', '品牌', '业态', '细分品类',
                                    '合同开始日期', '合同结束日期', '终止日期']
                    for col in ['合同开始日期', '合同结束日期', '终止日期']:
                        多经[col] = pd.to_datetime(多经[col], errors='coerce').dt.strftime('%Y%m%d')
                    多经['终止日期'] = 多经['终止日期'].fillna(多经['合同结束日期'])

                    合同台账 = pd.concat([正铺, 多经], ignore_index=True)

                    if not os.path.exists(invoice_info_path):
                        st.error(f"缺少文件：{invoice_info_path}")
                        st.stop()
                    开票 = pd.read_excel(invoice_info_path)
                    开票 = 开票[['供应商名称', '购买方名称', '购方识别号', '发票票种', '邮箱']]
                    开票.columns = ['供应商', '购买方名称', '购方识别号', '发票票种', '邮箱']

                    df['表号'] = df['表号'].astype(str)
                    df2['表号'] = df2['表号'].astype(str)
                    df = df.merge(df2[['表号', '客户编号', '合同号']], on='表号', how='left')
                    df = df.merge(df3, left_on='客户编号', right_on='编码', how='left')
                    df.drop('编码', axis=1, inplace=True)
                    df.rename(columns={'名称': '客户名称'}, inplace=True)

                    df['合同号'] = df['合同号'].astype(str).str.strip()
                    合同台账['合同号'] = 合同台账['合同号'].astype(str).str.strip()
                    df = df.merge(合同台账[['合同号', '供应商']], on='合同号', how='left')

                    df['供应商'] = df['供应商'].astype(str).str.strip()
                    开票['供应商'] = 开票['供应商'].astype(str).str.strip()
                    df = df.merge(开票, on='供应商', how='left')

                    df = df[['订单时间', '单号', '租户名称', '表号', '充值金额（元）', '付费途径',
                             '客户编号', '客户名称', '购买方名称', '购方识别号', '发票票种', '邮箱']]

                    st.success("转换成功！下方可编辑表格：")
                    st.session_state.converted_df = st.data_editor(
                        df, num_rows="dynamic", use_container_width=True, key="converted_editor"
                    )
                except Exception as e:
                    st.error(f"转换出错：{e}")

        # ========== 标签页4：余额预警 ==========
        with tab_alert:

            if 'df_result_full' not in st.session_state or st.session_state.df_result_full.empty:
                st.warning("请先在“启用明细”中生成数据。")
            else:
                if not os.path.exists(tenant_overview_path):
                    st.error(f"缺少文件：{tenant_overview_path}")
                else:
                    bal = pd.read_excel(tenant_overview_path)
                    bal = bal[['租户ID', '租户名称', '账户余额(元)', '剩余天数']]
                    bal.columns = ['表号', '租户名称', '剩余金额', '剩余天数']
                    bal['剩余天数'] = (bal['剩余天数'].astype(str).str.split('~').str[0]
                                       .pipe(pd.to_numeric, errors='coerce').fillna(0).astype(int))
                    bal['剩余金额'] = (bal['剩余金额'].astype(str)
                                       .str.replace(',', '', regex=False)
                                       .str.replace(' ', '', regex=False)
                                       .pipe(pd.to_numeric, errors='coerce').fillna(0.0))

                    df_full = st.session_state.df_result_full.copy()
                    df_full = df_full[df_full["电表状态"].isin(["启用中", "本月启用"])]

                    bal['表号'] = bal['表号'].astype(str)
                    df_full['表号'] = df_full['表号'].astype(str)

                    merged = df_full[['表号', '品牌', '进场日期']].drop_duplicates(subset='表号')
                    bal = bal.merge(merged, on='表号', how='inner')
                    bal = bal[['表号', '品牌', '剩余金额', '剩余天数', '租户名称']]
                    bal = bal[bal['品牌'] != "McDonald's 麦当劳"]

                    col_amt, col_day, col_logic, col_update, col_time = st.columns([1, 1, 1, 1, 1.2])
                    with col_amt:
                        amt_limit = st.number_input("剩余金额 ≤", min_value=0.0, value=100.0, step=0.01, format="%.2f",
                                                    key="alert_amt")
                    with col_day:
                        day_limit = st.number_input("剩余天数 ≤", min_value=0, value=3, step=1, format="%d",
                                                    key="alert_day")
                    with col_logic:
                        logic = st.selectbox("逻辑", ["且", "或"], key="alert_logic")
                    with col_update:
                        st.write("")
                        st.write("")
                        update_clicked = st.button("🔄 更新数据", use_container_width=True)
                    with col_time:
                        # 直接读取租户概览文件的修改时间
                        if os.path.exists(tenant_overview_path):
                            mtime = os.path.getmtime(tenant_overview_path)
                            last_time_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_time_str = "文件不存在"
                        st.caption(f"📅 更新时间：{last_time_str}")

                    if update_clicked:
                        with st.spinner("正在更新数据，请稍候..."):
                            success, msg = download_tenant_overview(tenant_overview_path)
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)

                    filtered = bal.copy()
                    cond_amt = filtered['剩余金额'] <= amt_limit
                    cond_day = filtered['剩余天数'] <= day_limit
                    final_cond = cond_amt & cond_day if logic == "且" else cond_amt | cond_day
                    filtered = filtered[final_cond]

                    st.markdown("---")
                    st.dataframe(filtered, use_container_width=True, hide_index=True)
                    st.caption(f"共 {len(filtered)} 条（总计 {len(bal)} 条）")

    # ========== 月末结转 ==========
    elif current_page == "月末结转":
        tab4, tab_balance, tab5, tab_contract, tab_price_adjust = st.tabs(
            ["💰电费结转", "💵余额核对", "⚙️参数表", "📄合同检查", "📈单价调整"],
            key="carryover_tabs"
        )

        with tab4:

            col1, col2, col3, col4 = st.columns([1, 1, 1, 1.5])
            with col1:
                carryover_date = st.date_input(
                    "结账日期",
                    value=date.today() - timedelta(days=1),   # 修改为前一天
                    key="carryover_date"
                )
            with col2:
                st.write("")
                st.write("")
                generate = st.button("🔍 生成明细", type="primary", use_container_width=True)
            with col3:
                st.write("")
                st.write("")
                has_data = not st.session_state.get("carryover_df", pd.DataFrame()).empty
                archive_clicked = st.button(
                    "💾 存档",
                    type="secondary",
                    use_container_width=True,
                    disabled=not has_data
                )
            with col4:
                st.write("")
                st.write("")
                # 下载凭证模板按钮
                voucher_template_path = os.path.join(BASE_DIR, "0-模板", "凭证模板.xlsx")
                download_voucher_clicked = st.button(
                    "📎 下载凭证模板",
                    type="secondary",
                    use_container_width=True,
                    disabled=not has_data
                )

            if 'carryover_df' not in st.session_state:
                st.session_state.carryover_df = pd.DataFrame()

            if archive_clicked:
                try:
                    period_str = carryover_date.strftime("%Y-%m")
                    df_to_save = st.session_state.carryover_df.copy()
                    df_to_save["期间"] = period_str

                    if os.path.exists(archive_path):
                        existing_df = pd.read_excel(archive_path, dtype=str)
                        if "期间" in existing_df.columns:
                            existing_df = existing_df[existing_df["期间"] != period_str]
                        combined_df = pd.concat([existing_df, df_to_save], ignore_index=True)
                    else:
                        combined_df = df_to_save

                    os.makedirs(os.path.dirname(archive_path), exist_ok=True)
                    combined_df.to_excel(archive_path, index=False, engine='openpyxl')
                    st.success(f"✅ 存档成功！")
                except Exception as e:
                    st.error(f"存档失败：{e}")
            if download_voucher_clicked:
                if not os.path.exists(voucher_template_path):
                    st.error(f"凭证模板文件不存在：{voucher_template_path}")
                else:
                    current_carryover_df = st.session_state.get("carryover_df", pd.DataFrame())
                    if current_carryover_df.empty:
                        st.warning("暂无结转明细，请先点击「生成明细」。")
                    else:
                        # 获取当前的结账日期（从 session_state 中存储，或从界面控件获取）
                        # 注意：carryover_date 已经在上面定义，但它在 generate 按钮外部也可见
                        # 可以直接使用 carryover_date（因为它在同一个 with tab4 块内）
                        try:
                            voucher_io = generate_voucher_from_carryover(
                                current_carryover_df,
                                carryover_date,  # 当前选中的结账日期
                                voucher_template_path
                            )
                            if voucher_io:
                                st.download_button(
                                    label="📥 点击下载凭证模板",
                                    data=voucher_io,
                                    file_name=f"凭证模板_{carryover_date.strftime('%Y%m%d')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_voucher_btn"
                                )
                            else:
                                st.error("生成凭证模板失败，请检查数据。")
                        except Exception as e:
                            st.error(f"生成凭证模板出错：{e}")

            if generate:
                if not os.path.exists(excel_path):
                    st.error(f"主数据文件不存在：{excel_path}")
                    st.stop()
                if df_existing.empty:
                    st.warning("暂无清表录入数据，无法计算结转。")
                    st.stop()

                try:
                    sheet_meter = "表底"
                    df_meter = pd.read_excel(excel_path, sheet_name=sheet_meter, dtype=str)
                    df_meter.columns = [c.strip() for c in df_meter.columns]
                    col_meter_no, col_meter_date, col_meter_value = df_meter.columns[:3]
                    df_meter[col_meter_date] = pd.to_datetime(df_meter[col_meter_date], errors='coerce')
                    df_meter[col_meter_value] = pd.to_numeric(
                        df_meter[col_meter_value].str.replace(',', '', regex=False).str.strip(),
                        errors='coerce'
                    ).fillna(0.0)
                    df_meter[col_meter_no] = df_meter[col_meter_no].astype(str)

                    from openpyxl import load_workbook
                    wb = load_workbook(excel_path, data_only=True)
                    if '参数表' not in wb.sheetnames:
                        st.error("缺少“参数表”工作表。")
                        st.stop()
                    ws = wb['参数表']
                    normal_income = float(ws['B1'].value) if ws['B1'].value is not None else 0.0
                    normal_cost = float(ws['B2'].value) if ws['B2'].value is not None else 0.0
                    wb.close()

                    try:
                        special_income_df = pd.read_excel(
                            excel_path, sheet_name='参数表', usecols='D:E', header=0, dtype=str
                        )
                        if special_income_df.empty:
                            special_income_df = pd.DataFrame(columns=['品牌', '收入单价'])
                        else:
                            special_income_df.columns = [special_income_df.columns[0], special_income_df.columns[1]]
                            special_income_df[special_income_df.columns[1]] = pd.to_numeric(
                                special_income_df[special_income_df.columns[1]].str.replace(',', '').str.strip(),
                                errors='coerce'
                            ).fillna(0.0)
                            special_income_df = special_income_df.dropna(subset=[special_income_df.columns[0]])
                            special_income_df[special_income_df.columns[0]] = special_income_df[
                                special_income_df.columns[0]].astype(str).str.strip()
                    except Exception:
                        special_income_df = pd.DataFrame(columns=['品牌', '收入单价'])

                    try:
                        special_cost_df = pd.read_excel(
                            excel_path, sheet_name='参数表', usecols='G:H', header=0, dtype=str
                        )
                        if special_cost_df.empty:
                            special_cost_df = pd.DataFrame(columns=['品牌', '成本单价'])
                        else:
                            special_cost_df.columns = [special_cost_df.columns[0], special_cost_df.columns[1]]
                            special_cost_df[special_cost_df.columns[1]] = pd.to_numeric(
                                special_cost_df[special_cost_df.columns[1]].str.replace(',', '').str.strip(),
                                errors='coerce'
                            ).fillna(0.0)
                            special_cost_df = special_cost_df.dropna(subset=[special_cost_df.columns[0]])
                            special_cost_df[special_cost_df.columns[0]] = special_cost_df[
                                special_cost_df.columns[0]].astype(str).str.strip()
                    except Exception:
                        special_cost_df = pd.DataFrame(columns=['品牌', '成本单价'])

                    df_in = df_existing[df_existing["清表类型"] == "进场清表"].copy()
                    df_out = df_existing[df_existing["清表类型"] == "撤场清表"].copy()
                    df_in["申请日期_dt"] = pd.to_datetime(df_in["申请日期"], errors="coerce")
                    df_out["申请日期_dt"] = pd.to_datetime(df_out["申请日期"], errors="coerce")

                    df_in_max = df_in.loc[df_in.groupby("自建ID")["申请日期_dt"].idxmax()]
                    df_out_max = df_out.loc[df_out.groupby("自建ID")["申请日期_dt"].idxmax()]

                    df_merged = pd.merge(
                        df_in_max[["自建ID", "表号", "品牌", "客户编号", "申请日期_dt", "表底"]],
                        df_out_max[["自建ID", "申请日期_dt", "表底"]],
                        on="自建ID", how="outer", suffixes=("_in", "_out")
                    )
                    df_merged.rename(columns={
                        "申请日期_dt_in": "进场日期", "表底_in": "进场表底",
                        "申请日期_dt_out": "撤场日期", "表底_out": "撤场表底"
                    }, inplace=True)

                    settlement = pd.Timestamp(carryover_date)
                    month_start = settlement.replace(day=1)

                    def calc_status(row):
                        in_date = row["进场日期"]
                        out_date = row["撤场日期"]
                        if pd.isna(in_date):
                            return "无进场记录"
                        if in_date > settlement:
                            return "未启用"
                        if (in_date >= month_start) and (pd.notna(out_date) and out_date <= settlement):
                            return "本月调整"
                        if (in_date < month_start) and (pd.notna(out_date) and month_start <= out_date <= settlement):
                            return "本月终止"
                        if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start):
                            return "已终止"
                        if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement):
                            return "本月启用"
                        if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement):
                            return "启用中"
                        return "异常状态"

                    df_merged["电表状态"] = df_merged.apply(calc_status, axis=1)
                    carryover_statuses = ["启用中", "本月启用", "本月终止", "本月调整"]
                    df_filtered = df_merged[df_merged["电表状态"].isin(carryover_statuses)].copy()

                    if df_filtered.empty:
                        st.info("当前结账日期下没有需要结转的电表。")
                        st.session_state.carryover_df = pd.DataFrame()
                        st.stop()

                    def get_start_date(row):
                        return max(row["进场日期"], month_start) if pd.notna(row["进场日期"]) else month_start

                    def get_end_date(row):
                        if row["电表状态"] in ("本月终止", "本月调整") and pd.notna(row["撤场日期"]):
                            return row["撤场日期"]
                        else:
                            return settlement + pd.Timedelta(days=1)

                    df_filtered["开始日期"] = df_filtered.apply(get_start_date, axis=1)
                    df_filtered["结束日期"] = df_filtered.apply(get_end_date, axis=1)

                    def get_meter_value(table_no, target_date):
                        if pd.isna(target_date):
                            return 0.0
                        match = df_meter[
                            (df_meter[col_meter_no] == str(table_no)) &
                            (df_meter[col_meter_date] == target_date)
                            ]
                        if not match.empty:
                            return match[col_meter_value].iloc[0]
                        return 0.0

                    start_values = [get_meter_value(r["表号"], r["开始日期"]) for _, r in df_filtered.iterrows()]
                    end_values = [get_meter_value(r["表号"], r["结束日期"]) for _, r in df_filtered.iterrows()]
                    df_filtered["开始表底"] = start_values
                    df_filtered["结束表底"] = end_values
                    df_filtered["用电量"] = df_filtered["结束表底"] - df_filtered["开始表底"]

                    inc_brand_col = special_income_df.columns[0] if not special_income_df.empty else None
                    inc_price_col = special_income_df.columns[1] if not special_income_df.empty else None
                    cost_brand_col = special_cost_df.columns[0] if not special_cost_df.empty else None
                    cost_price_col = special_cost_df.columns[1] if not special_cost_df.empty else None

                    def get_income_price(brand):
                        if inc_brand_col:
                            row = special_income_df[special_income_df[inc_brand_col] == str(brand).strip()]
                            if not row.empty:
                                return row[inc_price_col].iloc[0]
                        return normal_income

                    def get_cost_price(brand):
                        if cost_brand_col:
                            row = special_cost_df[special_cost_df[cost_brand_col] == str(brand).strip()]
                            if not row.empty:
                                return row[cost_price_col].iloc[0]
                        return normal_cost

                    df_filtered["收入单价"] = df_filtered["品牌"].apply(get_income_price)
                    df_filtered["成本单价"] = df_filtered["品牌"].apply(get_cost_price)

                    df_filtered["收入_含税"] = df_filtered["用电量"] * df_filtered["收入单价"]
                    df_filtered["收入_不含税"] = df_filtered["收入_含税"] / 1.13
                    df_filtered["成本_含税"] = df_filtered["用电量"] * df_filtered["成本单价"]
                    df_filtered["成本_不含税"] = df_filtered["成本_含税"] / 1.13

                    result = df_filtered[[
                        "电表状态", "表号", "品牌", "客户编号",
                        "开始表底", "结束表底", "用电量",
                        "收入单价", "收入_含税", "收入_不含税",
                        "成本单价", "成本_含税", "成本_不含税",
                        "开始日期", "结束日期"
                    ]].copy()

                    for col in ["开始表底", "结束表底", "用电量",
                                "收入_含税", "收入_不含税", "成本_含税", "成本_不含税"]:
                        result[col] = result[col].round(2)

                    result["开始日期"] = result["开始日期"].dt.strftime("%Y-%m-%d")
                    result["结束日期"] = result["结束日期"].dt.strftime("%Y-%m-%d")

                    st.session_state.carryover_df = result

                    st.success("✅ 结转明细生成成功！")
                    st.dataframe(result, use_container_width=True, hide_index=True)

                    total_blocks = len(result)
                    total_kwh = result["用电量"].sum()
                    total_inc_tax = result["收入_含税"].sum()
                    total_inc_notax = result["收入_不含税"].sum()
                    total_cost_tax = result["成本_含税"].sum()
                    total_cost_notax = result["成本_不含税"].sum()

                    st.info(
                        f"📊 本次结转电表 **{total_blocks}** 块，"
                        f"合计电量 **{total_kwh:.2f}** 度，"
                        f"含税收入 **{total_inc_tax:.2f}** 元，"
                        f"不含税收入 **{total_inc_notax:.2f}** 元，"
                        f"含税成本 **{total_cost_tax:.2f}** 元，"
                        f"不含税成本 **{total_cost_notax:.2f}** 元。"
                    )

                    from io import BytesIO
                    output = BytesIO()
                    # 添加汇总行后再导出
                    export_df = add_total_row_to_export(result)
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        export_df.to_excel(writer, index=False, sheet_name='结转明细')
                    output.seek(0)
                    st.download_button(
                        label="📥 下载结转明细",
                        data=output,
                        file_name=f"电费结转_{carryover_date.strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_carryover"
                    )

                except Exception as e:
                    st.error(f"生成结转明细时出错：{e}")
                    st.session_state.carryover_df = pd.DataFrame()

            elif not st.session_state.carryover_df.empty:
                result = st.session_state.carryover_df
                st.success("✅ 结转明细已生成（上次结果）")
                st.dataframe(result, use_container_width=True, hide_index=True)

                from io import BytesIO
                output = BytesIO()
                export_df = add_total_row_to_export(result)
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    export_df.to_excel(writer, index=False, sheet_name='结转明细')
                output.seek(0)
                st.download_button(
                    label="📥 下载结转明细",
                    data=output,
                    file_name=f"电费结转_{carryover_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_carryover_refresh"
                )
            else:
                st.info("请选择结账日期并点击“生成明细”按钮。")

        # ========== 余额核对 ==========
        with tab_balance:
            st.subheader("💰 余额核对")

            c1, c2, c3 = st.columns([1, 1, 1.5])
            with c1:
                period_str = st.text_input(
                    "核对期间",
                    value=date.today().strftime("%Y-%m"),
                    placeholder="格式：2026-04",
                    key="balance_period"
                )
            with c2:
                st.write("")
                st.write("")
                generate_balance = st.button("🔍 生成核对表", type="primary", use_container_width=True)
            with c3:
                st.write("")
                st.write("")
                download_report = st.button("📄 下载检查报告", type="secondary", use_container_width=True)

            def generate_inspection_report(period_label, balance_df):
                """生成Word检查报告，balance_df 参数保留但不再使用（仅为兼容）"""
                try:
                    from docx import Document
                    from docx.shared import Pt, Cm
                    from docx.enum.text import WD_ALIGN_PARAGRAPH
                    from docx.enum.section import WD_ORIENTATION
                    from docx.oxml.ns import qn
                except ImportError:
                    st.error("请先安装 python-docx 库：pip install python-docx")
                    return None

                # ---------- 创建文档，横向，黑体 ----------
                doc = Document()
                section = doc.sections[0]
                section.orientation = WD_ORIENTATION.LANDSCAPE
                section.page_width = Cm(29.7)
                section.page_height = Cm(21.0)
                doc.styles['Normal'].font.name = '黑体'
                doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
                for style in doc.styles:
                    if style.type == 1:  # 段落样式
                        style.font.name = '黑体'
                        style._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

                # ---------- 解析截止日期 ----------
                try:
                    year, month = map(int, period_label.split('-'))
                    from datetime import datetime, timedelta
                    if month == 12:
                        last_day = datetime(year, 12, 31)
                    else:
                        last_day = datetime(year, month + 1, 1) - timedelta(days=1)
                    deadline = last_day.strftime("%Y/%m/%d")
                except:
                    deadline = period_label

                # ---------- 1. 合同统计（正铺+多经，排除广告位）----------
                try:
                    if os.path.exists(zhengpu_contract_path):
                        zhengpu = pd.read_excel(zhengpu_contract_path, skiprows=1)
                        zhengpu = zhengpu[["状态", "合同号"]].copy()
                        zhengpu["状态"] = zhengpu["状态"].astype(str).str.strip()
                        zhengpu = zhengpu[zhengpu["状态"].isin(["生效中", "待生效"])]
                        zhengpu_count = zhengpu.shape[0]
                    else:
                        zhengpu_count = 0

                    duojing_count = 0
                    if os.path.exists(duojing_contract_path):
                        duojing = pd.read_excel(duojing_contract_path)
                        if "点位类型" in duojing.columns:
                            duojing = duojing[duojing["点位类型"] != "广告点位"]
                        if "合同状态" in duojing.columns:
                            duojing = duojing[duojing["合同状态"].isin(["生效中", "待生效"])]
                        duojing_count = duojing.shape[0]
                except:
                    zhengpu_count = 0
                    duojing_count = 0
                total_contracts = zhengpu_count + duojing_count

                # ---------- 2. 电表状态及合同异常明细（基于截止日期）----------
                if df_existing.empty:
                    st.warning("清表数据为空，无法生成完整报告。")
                    return None

                df_in = df_existing[df_existing["清表类型"] == "进场清表"].copy()
                df_out = df_existing[df_existing["清表类型"] == "撤场清表"].copy()
                df_in["申请日期_dt"] = pd.to_datetime(df_in["申请日期"], errors="coerce")
                df_out["申请日期_dt"] = pd.to_datetime(df_out["申请日期"], errors="coerce")
                df_in_max = df_in.loc[df_in.groupby("自建ID")["申请日期_dt"].idxmax()]
                df_out_max = df_out.loc[df_out.groupby("自建ID")["申请日期_dt"].idxmax()]

                df_merged = pd.merge(
                    df_in_max[["自建ID", "表号", "品牌", "客户编号", "合同号", "申请日期_dt", "表底"]],
                    df_out_max[["自建ID", "申请日期_dt", "表底"]],
                    on="自建ID", how="outer", suffixes=("_in", "_out")
                )
                df_merged.rename(columns={
                    "申请日期_dt_in": "进场日期", "表底_in": "进场表底",
                    "申请日期_dt_out": "撤场日期", "表底_out": "撤场表底"
                }, inplace=True)

                settlement = pd.Timestamp(last_day)
                month_start = settlement.replace(day=1)

                def calc_status(row):
                    in_date = row["进场日期"]
                    out_date = row["撤场日期"]
                    if pd.isna(in_date): return "无进场记录"
                    if in_date > settlement: return "未启用"
                    if (in_date >= month_start) and (pd.notna(out_date) and out_date <= settlement): return "本月调整"
                    if (in_date < month_start) and (
                            pd.notna(out_date) and month_start <= out_date <= settlement): return "本月终止"
                    if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start): return "已终止"
                    if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement): return "本月启用"
                    if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement): return "启用中"
                    return "异常状态"

                df_merged["电表状态"] = df_merged.apply(calc_status, axis=1)
                active_meters = df_merged[df_merged["电表状态"].isin(["启用中", "本月启用"])].copy()
                active_meters = active_meters[["表号", "品牌", "客户编号", "合同号"]].drop_duplicates()

                # 加载合同台账
                try:
                    if os.path.exists(zhengpu_contract_path):
                        zhengpu_full = pd.read_excel(zhengpu_contract_path, skiprows=1)
                        zhengpu_full = zhengpu_full[["状态", "合同号", "品牌"]].copy()
                        zhengpu_full["状态"] = zhengpu_full["状态"].astype(str).str.strip()
                    else:
                        zhengpu_full = pd.DataFrame(columns=["状态", "合同号", "品牌"])
                    if os.path.exists(duojing_contract_path):
                        duojing_full = pd.read_excel(duojing_contract_path)
                        duojing_full = duojing_full[["合同状态", "合同编号", "签约品牌", "点位类型"]].copy()
                        duojing_full.rename(columns={"合同状态": "状态", "合同编号": "合同号", "签约品牌": "品牌"},
                                            inplace=True)
                        if "点位类型" in duojing_full.columns:
                            duojing_full = duojing_full[duojing_full["点位类型"] != "广告点位"]
                        duojing_full["状态"] = duojing_full["状态"].astype(str).str.strip()
                        duojing_full = duojing_full[["状态", "合同号", "品牌"]]
                    else:
                        duojing_full = pd.DataFrame(columns=["状态", "合同号", "品牌"])
                    contracts = pd.concat([zhengpu_full, duojing_full], ignore_index=True)
                    contracts["合同号"] = contracts["合同号"].astype(str).str.strip()
                    contracts["品牌"] = contracts["品牌"].astype(str).str.strip()
                    valid_status = ["生效中", "待生效"]
                    active_contracts = contracts[contracts["状态"].isin(valid_status)].copy()
                    active_contracts = active_contracts[["合同号", "品牌"]].drop_duplicates()
                except Exception as e:
                    st.error(f"合同台账读取失败：{e}")
                    return None

                # 有电表合同异常
                meter_with_contract = active_meters.merge(
                    contracts[["合同号", "状态"]],
                    on="合同号",
                    how="left"
                )
                cond_no_contract = meter_with_contract["合同号"].isna() | (meter_with_contract["合同号"] == "")
                cond_status_invalid = ~meter_with_contract["状态"].isin(valid_status)
                abnormal_meters = meter_with_contract[cond_no_contract | cond_status_invalid]
                abnormal_meters = abnormal_meters[["表号", "品牌", "客户编号", "合同号"]].drop_duplicates()

                # 有合同无电表
                active_contract_nos = set(active_contracts["合同号"].tolist())
                meter_contract_nos = set(active_meters["合同号"].dropna().astype(str).tolist())
                missing_contract_nos = active_contract_nos - meter_contract_nos
                missing_contracts = active_contracts[active_contracts["合同号"].isin(missing_contract_nos)]
                missing_contracts = missing_contracts[["合同号", "品牌"]].drop_duplicates()

                # 加载备注
                def load_remarks(sheet_name, key_col):
                    if not os.path.exists(excel_path):
                        return {}
                    try:
                        with pd.ExcelFile(excel_path) as xl:
                            if sheet_name not in xl.sheet_names:
                                return {}
                            df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)
                            if key_col not in df.columns or "备注" not in df.columns:
                                return {}
                            df[key_col] = df[key_col].astype(str)
                            df["备注"] = df["备注"].fillna("")
                            return dict(zip(df[key_col], df["备注"]))
                    except:
                        return {}

                abnormal_remarks = load_remarks("有电表合同异常", "表号")
                abnormal_meters["备注"] = abnormal_meters["表号"].astype(str).map(abnormal_remarks).fillna("")
                missing_remarks = load_remarks("有合同无电表", "合同号")
                missing_contracts["备注"] = missing_contracts["合同号"].astype(str).map(missing_remarks).fillna("")

                # ---------- 3. 生成 Word 报告 ----------
                title = doc.add_heading(f'{period_label}电费能耗系统检查报告', 0)
                title.alignment = WD_ALIGN_PARAGRAPH.CENTER

                p = doc.add_paragraph(f"截止{deadline}，")
                p.add_run(
                    f"共用合同{total_contracts}份（不含广告位合同），其中正铺合同{zhengpu_count}份，多经合同{duojing_count}份；")

                # 有合同无电表明细（含备注）
                doc.add_heading('有合同无电表明细如下：', level=1)
                if missing_contracts.empty:
                    doc.add_paragraph('无')
                else:
                    table = doc.add_table(rows=1, cols=3)
                    table.style = 'Light Grid Accent 1'
                    hdr = table.rows[0].cells
                    hdr[0].text = '合同号'
                    hdr[1].text = '品牌'
                    hdr[2].text = '备注'
                    for _, row in missing_contracts.iterrows():
                        cells = table.add_row().cells
                        cells[0].text = str(row['合同号'])
                        cells[1].text = str(row['品牌'])
                        cells[2].text = str(row.get('备注', ''))

                # 有电表合同异常明细（含备注）
                doc.add_heading('有电表合同异常明细如下：', level=1)
                if abnormal_meters.empty:
                    doc.add_paragraph('无')
                else:
                    table = doc.add_table(rows=1, cols=5)
                    table.style = 'Light Grid Accent 1'
                    hdr = table.rows[0].cells
                    hdr[0].text = '表号'
                    hdr[1].text = '品牌'
                    hdr[2].text = '客户编号'
                    hdr[3].text = '合同号'
                    hdr[4].text = '备注'
                    for _, row in abnormal_meters.iterrows():
                        cells = table.add_row().cells
                        cells[0].text = str(row['表号'])
                        cells[1].text = str(row['品牌'])
                        cells[2].text = str(row['客户编号'])
                        cells[3].text = str(row['合同号'])
                        cells[4].text = str(row.get('备注', ''))

                # 仅保留文字“金蝶余额与能耗系统余额差异对比表：”，不添加表格
                doc.add_heading('金蝶余额与能耗系统余额差异对比表：', level=1)

                # 保存文档
                from io import BytesIO
                doc_bytes = BytesIO()
                doc.save(doc_bytes)
                doc_bytes.seek(0)
                return doc_bytes
            # 处理下载报告
            if download_report:
                if not period_str:
                    st.error("请填写核对期间")
                else:
                    if "balance_edited_df" not in st.session_state or st.session_state.balance_edited_df.empty:
                        st.warning("请先生成核对表（点击「生成核对表」按钮）")
                    else:
                        final_balance_df = st.session_state.balance_edited_df.copy()
                        doc_io = generate_inspection_report(period_str, final_balance_df)
                        if doc_io:
                            st.download_button(
                                label="📥 点击保存报告",
                                data=doc_io,
                                file_name=f"电费能耗检查报告_{period_str}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="download_report_btn"
                            )

            if generate_balance:
                import re
                if not re.match(r'^\d{4}-\d{2}$', period_str):
                    st.error("核对期间格式错误，请输入如 2026-04 的格式。")
                    st.stop()

                def get_active_meters_for_period(period_label, df_full=None):
                    if df_full is None or df_full.empty:
                        return pd.DataFrame(columns=['表号', '客户编号'])
                    from datetime import datetime
                    try:
                        year, month = map(int, period_label.split('-'))
                        settlement = pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)
                    except:
                        return pd.DataFrame(columns=['表号', '客户编号'])

                    df = df_full.copy()
                    df['进场日期_dt'] = pd.to_datetime(df['进场日期'], errors='coerce')
                    df['撤场日期_dt'] = pd.to_datetime(df['撤场日期'], errors='coerce')
                    month_start = settlement.replace(day=1)

                    def calc_status(row):
                        in_date = row['进场日期_dt']
                        out_date = row['撤场日期_dt']
                        if pd.isna(in_date):
                            return "无进场记录"
                        if in_date > settlement:
                            return "未启用"
                        if (in_date >= month_start) and (pd.notna(out_date) and out_date <= settlement):
                            return "本月调整"
                        if (in_date < month_start) and (pd.notna(out_date) and month_start <= out_date <= settlement):
                            return "本月终止"
                        if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start):
                            return "已终止"
                        if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement):
                            return "本月启用"
                        if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement):
                            return "启用中"
                        return "异常状态"

                    df['电表状态_核对'] = df.apply(calc_status, axis=1)
                    active = df[df['电表状态_核对'].isin(['启用中', '本月启用'])][
                        ['表号', '客户编号']].drop_duplicates()
                    return active

                try:
                    if not os.path.exists(customer_info_path):
                        st.error(f"客户资料文件不存在：{customer_info_path}")
                        st.stop()
                    df_cust = pd.read_excel(customer_info_path, dtype=str)[['编码', '名称']]
                    df_cust.columns = ['客户编号', '客户名称']

                    if not os.path.exists(balance_path):
                        st.error(f"辅助核算项目余额表不存在：{balance_path}")
                        st.stop()
                    df_bal = pd.read_excel(balance_path, header=0)
                    cust_bal_col = df_bal.columns[0]
                    end_bal_col = df_bal.columns[-1]
                    balance_map = dict(zip(df_bal[cust_bal_col].astype(str),
                                           pd.to_numeric(df_bal[end_bal_col], errors='coerce').fillna(0)))

                    carry_map = {}
                    archive_usable = pd.DataFrame()
                    if os.path.exists(archive_path):
                        df_archive = pd.read_excel(archive_path, dtype=str)
                        if '期间' in df_archive.columns and '收入_不含税' in df_archive.columns and '客户编号' in df_archive.columns:
                            df_period = df_archive[df_archive['期间'] == period_str].copy()
                            df_period['收入_不含税_num'] = pd.to_numeric(df_period['收入_不含税'],
                                                                         errors='coerce').fillna(0)
                            carry_map = df_period.groupby('客户编号')['收入_不含税_num'].sum().to_dict()
                            archive_usable = df_period
                        else:
                            st.warning("结转留存文件缺少必要列，将尝试从启用明细获取电表映射。")

                    if not os.path.exists(energy_balance_path):
                        st.error(f"能耗余额文件不存在：{energy_balance_path}")
                        st.stop()
                    df_energy = pd.read_excel(energy_balance_path, header=None, skiprows=2)
                    if df_energy.shape[1] < 4:
                        st.error("能耗余额文件列数不足（需要至少4列）。")
                        st.stop()
                    df_energy = df_energy[[0, 3]]
                    df_energy.columns = ['表号', '能耗余额']
                    df_energy['表号'] = df_energy['表号'].astype(str)
                    df_energy['能耗余额'] = (df_energy['能耗余额'].astype(str)
                                             .str.replace(',', '', regex=False)
                                             .str.replace(' ', '', regex=False))
                    df_energy['能耗余额'] = pd.to_numeric(df_energy['能耗余额'], errors='coerce').fillna(0)

                    active_archive = pd.DataFrame(columns=['表号', '客户编号'])
                    if not archive_usable.empty and '电表状态' in archive_usable.columns and '表号' in archive_usable.columns:
                        active_archive = archive_usable[archive_usable['电表状态'].isin(['启用中', '本月启用'])][
                            ['表号', '客户编号']].drop_duplicates()
                    else:
                        if 'df_result_full' in st.session_state and not st.session_state.df_result_full.empty:
                            active_archive = get_active_meters_for_period(period_str, st.session_state.df_result_full)
                            if not active_archive.empty:
                                st.info("已根据启用明细数据生成本期电表映射。如需使用结转留存数据，请先在电费结转页面存档。")
                        else:
                            st.warning("既无结转留存数据，也无启用明细数据，无法确定启用中电表，未到账金额将为0。")

                    if not active_archive.empty:
                        merged_energy = active_archive.merge(df_energy, on='表号', how='left')
                        merged_energy['能耗余额'] = merged_energy['能耗余额'].fillna(0)
                        energy_map = merged_energy.groupby('客户编号')['能耗余额'].sum().to_dict()
                    else:
                        energy_map = {}

                    finance_map = {}
                    try:
                        df_unsettled = pd.read_excel(excel_path, sheet_name='未到账明细', dtype=str)
                        if '表号' in df_unsettled.columns and '金额' in df_unsettled.columns:
                            df_unsettled['表号'] = df_unsettled['表号'].astype(str).str.strip()
                            df_unsettled['金额'] = (
                                df_unsettled['金额']
                                .astype(str)
                                .str.replace(',', '', regex=False)
                                .str.replace(' ', '', regex=False)
                            )
                            df_unsettled['金额'] = pd.to_numeric(df_unsettled['金额'], errors='coerce').fillna(0)
                            if not active_archive.empty:
                                merged = active_archive.merge(df_unsettled, on='表号', how='inner')
                                finance_map = merged.groupby('客户编号')['金额'].sum().to_dict()
                        else:
                            st.warning("“未到账明细”工作表中缺少“表号”或“金额”列，未到账金额置0。")
                    except ValueError:
                        st.warning("主数据文件中没有“未到账明细”工作表，未到账金额置0。")
                    except Exception as e:
                        st.warning(f"读取未到账明细失败：{e}，未到账金额置0。")

                    diff_map = {}
                    try:
                        wb_data = pd.ExcelFile(excel_path)
                        if '差异' in wb_data.sheet_names:
                            df_diff = pd.read_excel(excel_path, sheet_name='差异', header=0)
                            if len(df_diff.columns) >= 2:
                                diff_col0 = df_diff.columns[0]
                                diff_col1 = df_diff.columns[1]
                                diff_map = dict(zip(df_diff[diff_col0].astype(str),
                                                    pd.to_numeric(df_diff[diff_col1], errors='coerce').fillna(0)))
                    except Exception as e_diff:
                        st.warning(f"读取前期差异表失败：{e_diff}")

                    result_df = df_cust.copy()
                    result_df['金蝶期末余额'] = result_df['客户编号'].map(balance_map).fillna(0)
                    result_df['本月结转_不含税'] = result_df['客户编号'].map(carry_map).fillna(0)
                    result_df['金蝶测试金额_不含税'] = result_df['金蝶期末余额'] - result_df['本月结转_不含税']
                    result_df['金蝶测试金额_含税'] = round(result_df['金蝶测试金额_不含税'] * 1.13, 2)
                    result_df['能耗系统余额_汇总'] = result_df['客户编号'].map(energy_map).fillna(0)
                    result_df['未到账金额_含税'] = result_df['客户编号'].map(finance_map).fillna(0)
                    result_df['前期差异'] = result_df['客户编号'].map(diff_map).fillna(0)
                    result_df['新增差异'] = result_df['金蝶测试金额_含税'] - result_df['能耗系统余额_汇总'] + result_df[
                        '未到账金额_含税'] - result_df['前期差异']
                    result_df['新增差异'] = result_df['新增差异'].round(2)

                    df_sheet1 = pd.read_excel(excel_path, sheet_name=0, dtype=str)
                    meter_type_col = '电表类型' if '电表类型' in df_sheet1.columns else None

                    def get_reason(cust_id):
                        cust_meters = active_archive[active_archive['客户编号'] == cust_id][
                            '表号'].tolist() if not active_archive.empty else []
                        if not cust_meters:
                            return '已终止'
                        if meter_type_col is not None:
                            for meter in cust_meters:
                                meter_rows = df_sheet1[df_sheet1['表号'] == str(meter)]
                                if not meter_rows.empty:
                                    val = meter_rows[meter_type_col].iloc[0]
                                    if str(val).strip() != '预付费表':
                                        return '非预付费表'
                        return ''

                    result_df['差异原因'] = result_df['客户编号'].apply(get_reason)

                    result_df['abs_新增差异'] = result_df['新增差异'].abs()
                    result_df = result_df.sort_values('abs_新增差异', ascending=False).reset_index(drop=True)
                    result_df.drop(columns='abs_新增差异', inplace=True)

                    check_zero_cols = [
                        '金蝶期末余额', '本月结转_不含税', '金蝶测试金额_不含税',
                        '金蝶测试金额_含税', '能耗系统余额_汇总', '未到账金额_含税'
                    ]
                    mask = result_df[check_zero_cols].ne(0).any(axis=1)
                    result_df = result_df[mask].copy()

                    if result_df.empty:
                        st.warning("核对期内所有客户指定数据均为0，无有效核对记录。")
                        st.stop()

                    numeric_cols = [
                        '金蝶期末余额', '本月结转_不含税', '金蝶测试金额_不含税',
                        '金蝶测试金额_含税', '能耗系统余额_汇总', '未到账金额_含税',
                        '前期差异', '新增差异'
                    ]
                    total_row = {col: '' for col in result_df.columns}
                    total_row['客户编号'] = '合计'
                    total_row['客户名称'] = ''
                    for col in numeric_cols:
                        total_row[col] = result_df[col].sum()
                    total_row['差异原因'] = ''

                    total_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)

                    st.success("✅ 核对表生成成功！")
                    st.session_state.balance_edited_df = st.data_editor(
                        total_df,
                        num_rows="dynamic",
                        use_container_width=True,
                        key="balance_editor"
                    )

                    from io import BytesIO
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='余额核对')
                    output.seek(0)
                    st.download_button(
                        label="📥 下载核对表",
                        data=output,
                        file_name=f"余额核对_{period_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_balance"
                    )

                except Exception as e:
                    st.error(f"生成核对表失败：{e}")

        # ========== 参数表 ==========
        with tab5:

            from datetime import datetime

            if not os.path.exists(excel_path):
                st.error(f"未找到主数据文件：{excel_path}")
                st.stop()

            try:
                from openpyxl import load_workbook

                wb = load_workbook(excel_path)
                if '参数表' not in wb.sheetnames:
                    st.error("主数据文件中缺少“参数表”工作表，请检查模板。")
                    st.stop()
                ws = wb['参数表']

                normal_income = ws['B1'].value
                normal_income = float(normal_income) if normal_income is not None else 0.0

                normal_cost = ws['B2'].value
                normal_cost = float(normal_cost) if normal_cost is not None else 0.0

                try:
                    special_income_df = pd.read_excel(
                        excel_path, sheet_name='参数表',
                        usecols='D:E', header=0, dtype=str
                    )
                    if special_income_df.empty:
                        special_income_df = pd.DataFrame(columns=['商户', '收入单价'])
                    else:
                        cols = list(special_income_df.columns)
                        if len(cols) < 2:
                            special_income_df = pd.DataFrame(columns=['商户', '收入单价'])
                        else:
                            special_income_df.columns = [cols[0], cols[1]]
                        special_income_df[cols[1]] = pd.to_numeric(
                            special_income_df[cols[1]].str.replace(',', '', regex=False).str.strip(),
                            errors='coerce'
                        ).fillna(0.0)
                except Exception:
                    special_income_df = pd.DataFrame(columns=['商户', '收入单价'])

                try:
                    special_cost_df = pd.read_excel(
                        excel_path, sheet_name='参数表',
                        usecols='G:H', header=0, dtype=str
                    )
                    if special_cost_df.empty:
                        special_cost_df = pd.DataFrame(columns=['商户', '成本单价'])
                    else:
                        cols = list(special_cost_df.columns)
                        if len(cols) < 2:
                            special_cost_df = pd.DataFrame(columns=['商户', '成本单价'])
                        else:
                            special_cost_df.columns = [cols[0], cols[1]]
                        special_cost_df[cols[1]] = pd.to_numeric(
                            special_cost_df[cols[1]].str.replace(',', '', regex=False).str.strip(),
                            errors='coerce'
                        ).fillna(0.0)
                except Exception:
                    special_cost_df = pd.DataFrame(columns=['商户', '成本单价'])

                wb.close()

                # 重新布局：提交按钮旁边增加三个下载按钮
                col_income, col_cost, col_submit, col_dl_table, col_dl_balance, col_dl_order = st.columns([1, 1, 0.6, 0.8, 0.8, 0.8])

                with col_income:
                    new_normal_income = st.number_input(
                        "普通收入单价",
                        value=normal_income,
                        format="%.8f",
                        step=0.00000001,
                        key="param_normal_income"
                    )
                with col_cost:
                    new_normal_cost = st.number_input(
                        "普通成本单价",
                        value=normal_cost,
                        format="%.8f",
                        step=0.00000001,
                        key="param_normal_cost"
                    )
                with col_submit:
                    st.write("")
                    st.write("")
                    submitted = st.button("提交参数", type="primary", use_container_width=True)

                # ---------- 下载表底按钮 ----------
                with col_dl_table:
                    st.write("")
                    st.write("")
                    if st.button("下载表底", use_container_width=True, key="dl_meter_btn"):
                        with st.spinner("正在下载表底数据，请稍候..."):
                            success, msg = run_download_script("能耗_表底.py")
                            if success:
                                st.success("表底数据下载成功")
                                downloaded_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗表底.xls")
                                try:
                                    row_count = append_meter_readings_from_downloaded_file(excel_path, downloaded_file)
                                    st.success(f"表底追加成功")
                                except Exception as e:
                                    st.error(f"追加表底数据失败：{e}")
                            else:
                                st.error(f"下载表底失败：{msg}")
                    # 显示下载文件的最后修改时间
                    target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗表底.xls")
                    if os.path.exists(target_file):
                        mtime = os.path.getmtime(target_file)
                        date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                        st.caption(f"文件更新时间：{date_str}")
                    else:
                        st.caption("文件尚未下载")
                # ---------- 下载余额按钮 ----------
                with col_dl_balance:
                    st.write("")
                    st.write("")
                    if st.button("下载余额", use_container_width=True, key="dl_balance_btn"):
                        with st.spinner("正在下载余额数据，请稍候..."):
                            success, msg = run_download_script("能耗_余额.py")
                            if success:
                                st.success("余额数据下载成功")
                            else:
                                st.error(f"下载余额失败：{msg}")
                    target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗余额.xls")
                    if os.path.exists(target_file):
                        mtime = os.path.getmtime(target_file)
                        date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                        st.caption(f"文件更新时间：{date_str}")
                    else:
                        st.caption("文件尚未下载")

                # ---------- 下载财务订单按钮 ----------
                with col_dl_order:
                    st.write("")
                    st.write("")
                    if st.button("下载财务订单", use_container_width=True, key="dl_order_btn"):
                        with st.spinner("正在下载财务订单，请稍候..."):
                            success, msg = run_download_script("能耗_财务订单.py")
                            if success:
                                st.success("财务订单下载成功")
                            else:
                                st.error(f"下载财务订单失败：{msg}")
                    target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "财务订单.xls")
                    if os.path.exists(target_file):
                        mtime = os.path.getmtime(target_file)
                        date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                        st.caption(f"文件更新时间：{date_str}")
                    else:
                        st.caption("文件尚未下载")

                st.markdown("---")

                col_left, col_right = st.columns(2)
                with col_left:
                    st.caption("特殊收入单价")
                    edited_special_income = st.data_editor(
                        special_income_df,
                        num_rows="dynamic",
                        use_container_width=True,
                        key="editable_income",
                        column_config={
                            special_income_df.columns[0]: st.column_config.TextColumn(special_income_df.columns[0]),
                            special_income_df.columns[1]: st.column_config.NumberColumn(
                                special_income_df.columns[1],
                                format="%.8f",
                            )
                        }
                    )
                with col_right:
                    st.caption("特殊成本单价")
                    edited_special_cost = st.data_editor(
                        special_cost_df,
                        num_rows="dynamic",
                        use_container_width=True,
                        key="editable_cost",
                        column_config={
                            special_cost_df.columns[0]: st.column_config.TextColumn(special_cost_df.columns[0]),
                            special_cost_df.columns[1]: st.column_config.NumberColumn(
                                special_cost_df.columns[1],
                                format="%.8f",
                            )
                        }
                    )

                if submitted:
                    try:
                        wb_write = load_workbook(excel_path)
                        ws_write = wb_write['参数表']

                        ws_write['B1'] = new_normal_income
                        ws_write['B2'] = new_normal_cost

                        max_row = ws_write.max_row
                        for row in range(1, max_row + 1):
                            for col in [4, 5]:
                                ws_write.cell(row=row, column=col).value = None

                        income_cols = list(edited_special_income.columns)
                        ws_write.cell(row=1, column=4, value=income_cols[0])
                        ws_write.cell(row=1, column=5, value=income_cols[1])
                        for r_idx, row in edited_special_income.iterrows():
                            ws_write.cell(row=r_idx + 2, column=4, value=row[income_cols[0]])
                            ws_write.cell(row=r_idx + 2, column=5, value=float(row[income_cols[1]]))

                        for row in range(1, max_row + 1):
                            for col in [7, 8]:
                                ws_write.cell(row=row, column=col).value = None

                        cost_cols = list(edited_special_cost.columns)
                        ws_write.cell(row=1, column=7, value=cost_cols[0])
                        ws_write.cell(row=1, column=8, value=cost_cols[1])
                        for r_idx, row in edited_special_cost.iterrows():
                            ws_write.cell(row=r_idx + 2, column=7, value=row[cost_cols[0]])
                            ws_write.cell(row=r_idx + 2, column=8, value=float(row[cost_cols[1]]))

                        wb_write.save(excel_path)
                        wb_write.close()
                        st.success("参数保存成功！正在刷新页面...")
                        st.rerun()

                    except Exception as e:
                        st.error(f"保存失败：{e}")

                st.markdown("---")
                st.caption("📋 财务订单（仅显示部分列）")

                if os.path.exists(cwj_order_path):
                    try:
                        cols_to_keep = [0, 2, 4, 5, 6, 7, 9]
                        df_cwj = pd.read_excel(cwj_order_path, header=0, usecols=cols_to_keep)
                        edited_cwj = st.data_editor(
                            df_cwj,
                            num_rows="dynamic",
                            use_container_width=True,
                            key="cwj_editor_ui"
                        )
                        st.session_state.cwj_edited_df = edited_cwj

                        st.markdown("---")
                        save_unsettled = st.button("💾 保存未到账", type="secondary", use_container_width=True)
                        if save_unsettled:
                            try:
                                df_to_save = st.session_state.get('cwj_edited_df')
                                if df_to_save is None or df_to_save.empty:
                                    st.warning("当前没有编辑后的财务订单数据，请先编辑。")
                                else:
                                    meter_col = df_to_save.columns[2]
                                    amount_col = df_to_save.columns[5]

                                    df_write = df_to_save[[meter_col, amount_col]].copy()
                                    df_write.columns = ['表号', '金额']
                                    df_write = df_write.reset_index(drop=True)

                                    amount_series = (
                                        df_write['金额']
                                        .astype(str)
                                        .str.replace(',', '', regex=False)
                                        .str.replace(' ', '', regex=False)
                                    )
                                    amount_num = pd.to_numeric(amount_series, errors='coerce').fillna(0)
                                    total_amount = amount_num.sum()

                                    wb_target = load_workbook(excel_path)
                                    sheet_name = '未到账明细'

                                    if sheet_name in wb_target.sheetnames:
                                        del wb_target[sheet_name]

                                    ws_target = wb_target.create_sheet(sheet_name)
                                    ws_target.append(['表号', '金额'])

                                    for i in range(len(df_write)):
                                        ws_target.append([str(df_write.iloc[i]['表号']), amount_num.iloc[i]])

                                    wb_target.save(excel_path)
                                    wb_target.close()
                                    st.success(f"保存成功，未到账金额 {total_amount:,.2f} 元")

                            except Exception as e:
                                st.error(f"保存未到账明细失败：{e}")

                    except Exception as e:
                        st.error(f"读取财务订单文件失败：{e}")
                else:
                    st.warning(f"财务订单文件不存在：{cwj_order_path}")

            except ImportError:
                st.error("缺少 openpyxl 库，请先安装：pip install openpyxl")
            except Exception as e:
                st.error(f"读取参数表时发生错误：{e}")

        # ========== 合同检查 ==========
        with tab_contract:
            st.subheader("📄 合同检查")

            # 三列布局：结账日期 | 刷新按钮 | 合同替换功能（包含保存按钮）
            col_date, col_btn, col_replace = st.columns([2, 1, 2.5])
            with col_date:
                check_date = st.date_input(
                    "结账日期",
                    value=date.today(),
                    key="contract_check_date"  # 唯一 key
                )
            with col_btn:
                st.write("")
                st.write("")
                refresh_clicked = st.button("🔄 刷新数据", type="primary", use_container_width=True,
                                            key="refresh_contract_btn")

            with col_replace:
                st.write("**合同号替换**")
                col_old, col_new, col_submit, col_save = st.columns([2, 2, 1, 1.2])
                with col_old:
                    old_contract = st.text_input(
                        "原合同号", placeholder="输入现有合同号",
                        key="contract_old_contract"  # 唯一 key
                    )
                with col_new:
                    new_contract = st.text_input(
                        "新合同号", placeholder="输入新合同号",
                        key="contract_new_contract"  # 唯一 key
                    )
                with col_submit:
                    st.write("")
                    st.write("")
                    replace_clicked = st.button("替换", use_container_width=True, key="replace_contract_btn")
                with col_save:
                    st.write("")
                    st.write("")
                    save_remarks_clicked = st.button("保存备注", type="primary", use_container_width=True,
                                                     key="save_remarks_btn")

            # ---------- 辅助函数：加载指定 sheet 的备注 ----------
            def load_remarks_from_excel(sheet_name, key_col, remark_col="备注"):
                """从 Excel 读取备注，返回 {key: remark} 字典"""
                if not os.path.exists(excel_path):
                    return {}
                try:
                    with pd.ExcelFile(excel_path) as xl:
                        if sheet_name not in xl.sheet_names:
                            return {}
                        df = pd.read_excel(xl, sheet_name=sheet_name, dtype=str)
                        if key_col not in df.columns or remark_col not in df.columns:
                            return {}
                        df[key_col] = df[key_col].astype(str)
                        df[remark_col] = df[remark_col].fillna("")
                        return dict(zip(df[key_col], df[remark_col]))
                except Exception:
                    return {}

            def save_remarks_to_excel(dataframe, sheet_name, key_col, remark_col="备注"):
                """保存 dataframe 到指定 sheet，只保留主键和备注列（覆盖整个 sheet）"""
                if dataframe.empty:
                    try:
                        from openpyxl import load_workbook
                        wb = load_workbook(excel_path)
                        if sheet_name in wb.sheetnames:
                            del wb[sheet_name]
                            wb.save(excel_path)
                        return
                    except:
                        pass
                    return
                save_df = dataframe[[key_col, remark_col]].copy()
                save_df = save_df.drop_duplicates(subset=[key_col])
                from openpyxl import load_workbook, Workbook
                if os.path.exists(excel_path):
                    wb = load_workbook(excel_path)
                else:
                    wb = Workbook()
                if sheet_name in wb.sheetnames:
                    del wb[sheet_name]
                ws = wb.create_sheet(sheet_name)
                ws.cell(row=1, column=1, value=key_col)
                ws.cell(row=1, column=2, value=remark_col)
                for i, (key, remark) in enumerate(save_df.itertuples(index=False), start=2):
                    ws.cell(row=i, column=1, value=str(key))
                    ws.cell(row=i, column=2, value=str(remark))
                wb.save(excel_path)

            # ---------- 处理合同号替换 ----------
            if replace_clicked:
                if not old_contract or not new_contract:
                    st.error("原合同号和新合同号均不能为空！")
                else:
                    try:
                        if not os.path.exists(excel_path):
                            st.error(f"主数据文件不存在：{excel_path}")
                        else:
                            df_cur = pd.read_excel(excel_path, sheet_name=0, dtype=str)
                            df_cur = process_date_column(df_cur)
                            if old_contract not in df_cur["合同号"].values:
                                st.warning(f"未找到合同号为「{old_contract}」的记录，无需替换。")
                            else:
                                mask = df_cur["合同号"] == old_contract
                                count = mask.sum()
                                df_cur.loc[mask, "合同号"] = new_contract
                                from openpyxl import load_workbook
                                wb = load_workbook(excel_path)
                                ws = wb['Sheet1']
                                # 清空数据行（保留表头）
                                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                                    for cell in row:
                                        cell.value = None
                                for col_idx, col_name in enumerate(df_cur.columns, 1):
                                    ws.cell(row=1, column=col_idx, value=col_name)
                                for r_idx, row in df_cur.iterrows():
                                    for c_idx, value in enumerate(row):
                                        ws.cell(row=r_idx + 2, column=c_idx + 1, value=value)
                                wb.save(excel_path)
                                wb.close()
                                load_excel.clear()
                                st.success(f"已将 {count} 条记录的合同号从「{old_contract}」替换为「{new_contract}」")
                                st.rerun()
                    except Exception as e:
                        st.error(f"替换失败：{e}")

            # ---------- 数据刷新逻辑 ----------
            if refresh_clicked or "contract_check_result" not in st.session_state:
                with st.spinner("正在分析合同与电表关系..."):
                    try:
                        if df_existing.empty:
                            st.warning("暂无清表录入数据，无法分析电表状态。")
                            st.stop()
                        df_in = df_existing[df_existing["清表类型"] == "进场清表"].copy()
                        df_out = df_existing[df_existing["清表类型"] == "撤场清表"].copy()
                        df_in["申请日期_dt"] = pd.to_datetime(df_in["申请日期"], errors="coerce")
                        df_out["申请日期_dt"] = pd.to_datetime(df_out["申请日期"], errors="coerce")
                        df_in_max = df_in.loc[df_in.groupby("自建ID")["申请日期_dt"].idxmax()]
                        df_out_max = df_out.loc[df_out.groupby("自建ID")["申请日期_dt"].idxmax()]
                        df_merged = pd.merge(
                            df_in_max[["自建ID", "表号", "品牌", "客户编号", "合同号", "申请日期_dt", "表底"]],
                            df_out_max[["自建ID", "申请日期_dt", "表底"]],
                            on="自建ID", how="outer", suffixes=("_in", "_out")
                        )
                        df_merged.rename(columns={
                            "申请日期_dt_in": "进场日期", "表底_in": "进场表底",
                            "申请日期_dt_out": "撤场日期", "表底_out": "撤场表底"
                        }, inplace=True)

                        settlement = pd.Timestamp(check_date)
                        month_start = settlement.replace(day=1)

                        def calc_status(row):
                            in_date = row["进场日期"]
                            out_date = row["撤场日期"]
                            if pd.isna(in_date): return "无进场记录"
                            if in_date > settlement: return "未启用"
                            if (in_date >= month_start) and (
                                    pd.notna(out_date) and out_date <= settlement): return "本月调整"
                            if (in_date < month_start) and (
                                    pd.notna(out_date) and month_start <= out_date <= settlement): return "本月终止"
                            if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start): return "已终止"
                            if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement): return "本月启用"
                            if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement): return "启用中"
                            return "异常状态"

                        df_merged["电表状态"] = df_merged.apply(calc_status, axis=1)
                        active_meters = df_merged[df_merged["电表状态"].isin(["启用中", "本月启用"])].copy()
                        active_meters = active_meters[["表号", "品牌", "客户编号", "合同号"]].drop_duplicates()

                        # 合同台账（正铺 + 多经，多经排除广告点位）
                        if not os.path.exists(zhengpu_contract_path) or not os.path.exists(duojing_contract_path):
                            st.error("缺少合同台账文件，请检查路径配置。")
                            st.stop()

                        zhengpu = pd.read_excel(zhengpu_contract_path, skiprows=1)
                        zhengpu = zhengpu[["状态", "合同号", "品牌"]].copy()
                        zhengpu["状态"] = zhengpu["状态"].astype(str).str.strip()

                        duojing = pd.read_excel(duojing_contract_path)
                        duojing = duojing[["合同状态", "合同编号", "签约品牌", "点位类型"]].copy()
                        duojing.rename(columns={"合同状态": "状态", "合同编号": "合同号", "签约品牌": "品牌"},
                                       inplace=True)
                        duojing = duojing[duojing["点位类型"] != "广告点位"]
                        duojing["状态"] = duojing["状态"].astype(str).str.strip()
                        duojing = duojing[["状态", "合同号", "品牌"]]

                        contracts = pd.concat([zhengpu, duojing], ignore_index=True)
                        contracts["合同号"] = contracts["合同号"].astype(str).str.strip()
                        contracts["品牌"] = contracts["品牌"].astype(str).str.strip()
                        valid_status = ["生效中", "待生效"]
                        active_contracts = contracts[contracts["状态"].isin(valid_status)].copy()
                        active_contracts = active_contracts[["合同号", "品牌"]].drop_duplicates()

                        # 有电表合同异常
                        meter_with_contract = active_meters.merge(
                            contracts[["合同号", "状态"]],
                            on="合同号",
                            how="left"
                        )
                        cond_no_contract = meter_with_contract["合同号"].isna() | (meter_with_contract["合同号"] == "")
                        cond_status_invalid = ~meter_with_contract["状态"].isin(valid_status)
                        abnormal_meters = meter_with_contract[cond_no_contract | cond_status_invalid]
                        abnormal_meters = abnormal_meters[["表号", "品牌", "客户编号", "合同号"]].drop_duplicates()
                        # 加载备注
                        abnormal_remarks = load_remarks_from_excel("有电表合同异常", "表号", "备注")
                        abnormal_meters["备注"] = abnormal_meters["表号"].astype(str).map(abnormal_remarks).fillna("")

                        # 有合同无电表
                        active_contract_nos = set(active_contracts["合同号"].tolist())
                        meter_contract_nos = set(active_meters["合同号"].dropna().astype(str).tolist())
                        missing_contract_nos = active_contract_nos - meter_contract_nos
                        missing_contracts = active_contracts[active_contracts["合同号"].isin(missing_contract_nos)]
                        missing_contracts = missing_contracts[["合同号", "品牌"]].drop_duplicates()
                        missing_remarks = load_remarks_from_excel("有合同无电表", "合同号", "备注")
                        missing_contracts["备注"] = missing_contracts["合同号"].astype(str).map(missing_remarks).fillna(
                            "")

                        st.session_state.contract_check_result = {
                            "abnormal_meters": abnormal_meters,
                            "missing_contracts": missing_contracts
                        }
                        st.success("分析完成！")
                    except Exception as e:
                        st.error(f"分析过程中出错：{e}")
                        st.session_state.contract_check_result = None

            # ---------- 显示可编辑表格 ----------
            result = st.session_state.get("contract_check_result")
            if result:
                abnormal_df = result["abnormal_meters"]
                missing_df = result["missing_contracts"]

                col_left, col_right = st.columns(2)
                with col_left:
                    st.markdown("#### ⚠️ 有电表合同异常")
                    if abnormal_df.empty:
                        st.info("暂无异常电表（所有启用电表的合同均为生效中/待生效）")
                    else:
                        edited_abnormal = st.data_editor(
                            abnormal_df,
                            use_container_width=True,
                            hide_index=True,
                            key="abnormal_editor",
                            column_config={
                                "表号": st.column_config.TextColumn("表号", disabled=True),
                                "品牌": st.column_config.TextColumn("品牌", disabled=True),
                                "客户编号": st.column_config.TextColumn("客户编号", disabled=True),
                                "合同号": st.column_config.TextColumn("合同号", disabled=True),
                                "备注": st.column_config.TextColumn("备注", required=False),
                            }
                        )
                        st.session_state.edited_abnormal = edited_abnormal
                        st.caption(f"共 {len(abnormal_df)} 条")
                with col_right:
                    st.markdown("#### 📄 有合同无电表")
                    if missing_df.empty:
                        st.info("暂无缺失电表的合同（所有生效中/待生效合同均已关联电表）")
                    else:
                        edited_missing = st.data_editor(
                            missing_df,
                            use_container_width=True,
                            hide_index=True,
                            key="missing_editor",
                            column_config={
                                "合同号": st.column_config.TextColumn("合同号", disabled=True),
                                "品牌": st.column_config.TextColumn("品牌", disabled=True),
                                "备注": st.column_config.TextColumn("备注", required=False),
                            }
                        )
                        st.session_state.edited_missing = edited_missing
                        st.caption(f"共 {len(missing_df)} 条")
            else:
                st.info("请点击「刷新数据」按钮生成合同检查报告。")

            # ---------- 保存备注 ----------
            if save_remarks_clicked:
                if "edited_abnormal" in st.session_state and "edited_missing" in st.session_state:
                    edited_ab = st.session_state.edited_abnormal
                    if not edited_ab.empty:
                        save_remarks_to_excel(edited_ab, "有电表合同异常", "表号", "备注")
                    else:
                        save_remarks_to_excel(pd.DataFrame(), "有电表合同异常", "表号", "备注")

                    edited_mis = st.session_state.edited_missing
                    if not edited_mis.empty:
                        save_remarks_to_excel(edited_mis, "有合同无电表", "合同号", "备注")
                    else:
                        save_remarks_to_excel(pd.DataFrame(), "有合同无电表", "合同号", "备注")

                    st.success("备注已保存！")
                    # 刷新当前显示的备注
                    if result:
                        abnormal_remarks = load_remarks_from_excel("有电表合同异常", "表号", "备注")
                        missing_remarks = load_remarks_from_excel("有合同无电表", "合同号", "备注")
                        result["abnormal_meters"]["备注"] = result["abnormal_meters"]["表号"].astype(str).map(
                            abnormal_remarks).fillna("")
                        result["missing_contracts"]["备注"] = result["missing_contracts"]["合同号"].astype(str).map(
                            missing_remarks).fillna("")
                        st.session_state.contract_check_result = result
                        st.rerun()
                else:
                    st.warning("请先刷新数据生成表格。")
           # ========== 单价调整 ==========
        with tab_price_adjust:
                from io import BytesIO
                st.subheader("📈 单价调整")

                col_period, col_price, col_btn1, col_btn2 = st.columns([1, 1, 1, 1])
                with col_period:
                    adjust_period = st.text_input(
                        "调整期间",
                        value=date.today().strftime("%Y-%m"),
                        placeholder="格式：2026-04",
                        key="adjust_period"
                    )
                with col_price:
                    adjust_unit_price = st.number_input(
                        "调整单价（元/度）",
                        min_value=0.0,
                        step=0.00000001,
                        format="%.10f",
                        value=0.75,
                        key="adjust_unit_price"
                    )

                # 初始化变量
                detail_file_bytes = None
                voucher_file_bytes = None
                display_with_total = pd.DataFrame()

                if os.path.exists(archive_path):
                    try:
                        df_archive = pd.read_excel(archive_path, dtype=str)
                        if '期间' not in df_archive.columns:
                            st.error("结转留存文件中缺少“期间”列")
                        else:
                            df_period = df_archive[df_archive['期间'] == adjust_period].copy()
                            if not df_period.empty:
                                required_cols = [
                                    '电表状态', '表号', '品牌', '客户编号',
                                    '开始表底', '结束表底', '用电量',
                                    '成本单价', '成本_含税', '成本_不含税'
                                ]
                                missing = [c for c in required_cols if c not in df_period.columns]
                                if missing:
                                    st.error(f"缺少列：{missing}")
                                else:
                                    df_display = df_period[required_cols].copy()
                                    df_display.rename(columns={
                                        '成本单价': '原成本单价',
                                        '成本_含税': '原成本_含税',
                                        '成本_不含税': '原成本_不含税'
                                    }, inplace=True)

                                    numeric_cols = ['用电量', '原成本_含税', '原成本_不含税']
                                    for col in numeric_cols:
                                        df_display[col] = pd.to_numeric(df_display[col], errors='coerce').fillna(0)

                                    df_display['新成本单价'] = adjust_unit_price
                                    df_display['新成本_含税'] = (df_display['用电量'] * df_display['新成本单价']).round(
                                        2)
                                    df_display['新成本_不含税'] = (df_display['新成本_含税'] / 1.13).round(2)

                                    # 带汇总行的显示表格
                                    display_with_total = df_display.copy()
                                    total_row = {}
                                    for col in display_with_total.columns:
                                        if col in numeric_cols + ['新成本_含税', '新成本_不含税']:
                                            total_row[col] = display_with_total[col].sum()
                                        else:
                                            total_row[col] = ''
                                    total_row['表号'] = '合计'
                                    display_with_total = pd.concat([display_with_total, pd.DataFrame([total_row])],
                                                                   ignore_index=True)

                                    # 生成明细下载文件
                                    detail_io = BytesIO()
                                    with pd.ExcelWriter(detail_io, engine='openpyxl') as writer:
                                        display_with_total.to_excel(writer, index=False, sheet_name='单价调整')
                                    detail_io.seek(0)
                                    detail_file_bytes = detail_io

                                    # 生成凭证下载文件
                                    template_voucher_path = os.path.join(BASE_DIR, "0-模板", "凭证模板.xlsx")
                                    if os.path.exists(template_voucher_path):
                                        voucher_io = generate_adjust_voucher(
                                            df_display,
                                            adjust_period,
                                            template_voucher_path
                                        )
                                        voucher_file_bytes = voucher_io
                                    else:
                                        st.warning(f"凭证模板文件不存在：{template_voucher_path}")
                    except Exception as e:
                        st.error(f"处理数据出错：{e}")

                # 下载按钮（放在同一行）
                with col_btn1:
                    if detail_file_bytes is not None:
                        st.download_button(
                            label="📥 下载调整后明细",
                            data=detail_file_bytes,
                            file_name=f"单价调整_{adjust_period}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_download_detail"
                        )
                    else:
                        st.button("📥 下载调整后明细", disabled=True, use_container_width=True)

                with col_btn2:
                    if voucher_file_bytes is not None:
                        st.download_button(
                            label="📎 下载凭证模板（调整后）",
                            data=voucher_file_bytes,
                            file_name=f"电费调整凭证_{adjust_period}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="btn_download_voucher"
                        )
                    else:
                        st.button("📎 下载凭证模板（调整后）", disabled=True, use_container_width=True)

                # 显示表格
                if not display_with_total.empty:
                    st.markdown("---")
                    st.dataframe(display_with_total, use_container_width=True, hide_index=True)
                elif adjust_period:
                    st.info("没有找到对应期间的数据或数据缺少必要列。")
    elif current_page == "能耗系统-水费":
        # 水费清表录入标签页
        tab1, tab2, tab3, tab_alert, tab_balance = st.tabs(
            ["📋清表录入", "⚡启用明细", "⚙️参数表", "⚠️水费结转", "💧水费余额"],
            key="water_tabs"
        )

        with tab1:
            # ---------- 清表录入表单 ----------
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                water_table_no = st.text_input("表号", placeholder="能耗系统表号", key="water_table_no")
            with c2:
                water_brand = st.text_input("品牌", placeholder="规范名称", key="water_brand")
            with c3:
                water_customer_id = st.text_input("客户编号", placeholder="金蝶系统编号", key="water_customer_id")
            with c4:
                water_apply_date = st.date_input("申请日期", value=date.today(), key="water_apply_date")
            with c5:
                water_meter_bottom = st.number_input("表底", min_value=0.0, step=0.1, format="%.2f",
                                                     key="water_meter_bottom")

            col_type, col_id, col_contract = st.columns([1, 1, 1])
            with col_type:
                water_clear_type = st.radio("清表类型", ["进场清表", "撤场清表"], horizontal=True,
                                            key="water_clear_type")

            with col_id:
                # 自动生成自建ID（进场清表且表号、品牌、申请日期均存在时）
                if water_clear_type == "进场清表" and water_table_no and water_brand and water_apply_date:
                    date_str = f"{water_apply_date.year}/{water_apply_date.month}/{water_apply_date.day}"
                    auto_id = f"{water_table_no}{water_brand}{date_str}"
                    st.session_state.water_id_input = auto_id
                else:
                    if water_clear_type == "进场清表":
                        st.session_state.water_id_input = ""
                st.text_input("自建ID（可修改）", key="water_id_input", placeholder="自动生成或手动填写")

            with col_contract:
                water_contract_no = st.text_input("合同号", key="water_contract_no", placeholder="可选，可手动填写")

            col_btn1, col_btn2 = st.columns([1, 1])
            with col_btn1:
                water_submitted = st.button("提交数据", type="primary", use_container_width=True, key="water_submit")
            with col_btn2:
                # 清空按钮（定义清空回调函数）
                def clear_water_form():
                    for key in ["water_table_no", "water_brand", "water_customer_id", "water_apply_date",
                                "water_meter_bottom", "water_clear_type", "water_id_input", "water_contract_no"]:
                        if key in st.session_state:
                            if key == "water_apply_date":
                                st.session_state[key] = date.today()
                            elif key == "water_meter_bottom":
                                st.session_state[key] = 0.0
                            else:
                                st.session_state[key] = ""

                st.button("清空", on_click=clear_water_form, use_container_width=True, key="water_clear")

            # ---------- 提交处理 ----------
            if water_submitted:
                if not water_table_no or not water_brand or not water_customer_id:
                    st.error("表号、品牌和客户编号不能为空！")
                else:
                    new_row = {
                        "表号": water_table_no,
                        "品牌": water_brand,
                        "客户编号": water_customer_id,
                        "申请日期": str(water_apply_date),
                        "表底": water_meter_bottom,
                        "清表类型": water_clear_type,
                        "自建ID": st.session_state.get("water_id_input", ""),
                        "合同号": water_contract_no,
                    }
                    try:
                        # 读取现有数据
                        if os.path.exists(water_excel_path):
                            df_cur = pd.read_excel(water_excel_path, sheet_name=0, dtype=str)
                            df_cur = process_date_column(df_cur)
                        else:
                            df_cur = pd.DataFrame(
                                columns=["表号", "品牌", "客户编号", "申请日期", "表底", "清表类型", "自建ID",
                                         "合同号"])
                        new_df = pd.DataFrame([new_row])
                        df_cur = pd.concat([df_cur, new_df], ignore_index=True)
                        # 保证列顺序
                        column_order = ["表号", "品牌", "客户编号", "申请日期", "表底", "清表类型", "自建ID", "合同号"]
                        df_cur = df_cur[column_order]

                        # 保存到Excel
                        from openpyxl import load_workbook, Workbook
                        if os.path.exists(water_excel_path):
                            wb = load_workbook(water_excel_path)
                            if 'Sheet1' in wb.sheetnames:
                                ws = wb['Sheet1']
                            else:
                                ws = wb.active
                                ws.title = 'Sheet1'
                            # 清空数据区域（保留表头）
                            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                                for cell in row:
                                    cell.value = None
                        else:
                            wb = Workbook()
                            ws = wb.active
                            ws.title = 'Sheet1'
                        # 写入表头
                        for col_idx, col_name in enumerate(df_cur.columns, 1):
                            ws.cell(row=1, column=col_idx, value=col_name)
                        for r_idx, row in df_cur.iterrows():
                            for c_idx, value in enumerate(row):
                                ws.cell(row=r_idx + 2, column=c_idx + 1, value=value)
                        wb.save(water_excel_path)
                        wb.close()

                        # 追加表底记录到“表底”工作表
                        append_single_meter_reading(water_excel_path, water_table_no, water_apply_date,
                                                    water_meter_bottom)
                        st.success("录入成功")
                        st.rerun()
                    except Exception as e:
                        st.error(f"保存失败：{e}")

            # ---------- 展示已有清表记录 ----------
            st.markdown("---")
            st.subheader("📊 清表记录")
            # 加载现有水费数据
            df_water_existing = load_excel(water_excel_path)  # 复用 load_excel 函数
            # 筛选条件（根据当前表单输入）
            filtered_water = df_water_existing.copy()
            if water_table_no:
                filtered_water = filtered_water[filtered_water["表号"].str.contains(water_table_no, na=False)]
            if water_brand:
                filtered_water = filtered_water[filtered_water["品牌"].str.contains(water_brand, na=False)]
            if water_customer_id:
                filtered_water = filtered_water[filtered_water["客户编号"].str.contains(water_customer_id, na=False)]

            # 撤场清表时支持行点击填充
            if water_clear_type == "撤场清表":
                st.data_editor(
                    filtered_water,
                    use_container_width=True,
                    disabled=True,
                    key="water_clear_table_selector",
                    hide_index=True,
                    num_rows="fixed",
                )
                selection = st.session_state.get("water_clear_table_selector", {}).get("selection", {})
                if selection and "rows" in selection and selection["rows"]:
                    selected_idx = selection["rows"][0]
                    if "water_last_filled_idx" not in st.session_state:
                        st.session_state.water_last_filled_idx = -1
                    if selected_idx != st.session_state.water_last_filled_idx:
                        st.session_state.water_last_filled_idx = selected_idx
                        if selected_idx < len(filtered_water):
                            row = filtered_water.iloc[selected_idx]
                            st.session_state.water_brand = str(row.get("品牌", ""))
                            st.session_state.water_customer_id = str(row.get("客户编号", ""))
                            st.session_state.water_id_input = str(row.get("自建ID", ""))
                            st.session_state.water_contract_no = str(row.get("合同号", ""))
                            st.session_state.water_table_no = str(row.get("表号", ""))
                            try:
                                st.session_state.water_meter_bottom = float(row.get("表底", 0.0))
                            except:
                                st.session_state.water_meter_bottom = 0.0
                            app_date = row.get("申请日期", "")
                            if app_date:
                                try:
                                    st.session_state.water_apply_date = pd.to_datetime(app_date).date()
                                except:
                                    pass
                            st.rerun()
            else:
                st.dataframe(filtered_water, use_container_width=True)

            st.caption(f"共 {len(filtered_water)} 条记录（总记录数：{len(df_water_existing)}）")

        # 其余标签页可以显示“建设中”或留空
        with tab2:
            col_date, col_status, col_no, col_brand = st.columns([2, 2, 1.5, 1.5])
            with col_date:
                water_settlement_date = st.date_input("结账日期", value=date.today(), key="water_settlement_date")
            with col_status:
                water_status_placeholder = st.empty()
            with col_no:
                water_filter_table_no = st.text_input("表号", placeholder="模糊查询", key="water_filter_table_no")
            with col_brand:
                water_filter_brand = st.text_input("品牌", placeholder="模糊查询", key="water_filter_brand")

            # 加载水费清表数据
            df_water_existing = load_excel(water_excel_path)

            if not df_water_existing.empty:
                from datetime import datetime

                df_in = df_water_existing[df_water_existing["清表类型"] == "进场清表"].copy()
                df_out = df_water_existing[df_water_existing["清表类型"] == "撤场清表"].copy()

                df_in["申请日期_dt"] = pd.to_datetime(df_in["申请日期"], errors="coerce")
                df_out["申请日期_dt"] = pd.to_datetime(df_out["申请日期"], errors="coerce")

                df_in_max = df_in.loc[df_in.groupby("自建ID")["申请日期_dt"].idxmax()]
                df_out_max = df_out.loc[df_out.groupby("自建ID")["申请日期_dt"].idxmax()]

                df_merged = pd.merge(
                    df_in_max[["自建ID", "表号", "品牌", "客户编号", "申请日期_dt", "表底"]],
                    df_out_max[["自建ID", "申请日期_dt", "表底"]],
                    on="自建ID", how="outer", suffixes=("_in", "_out")
                )
                df_merged.rename(columns={
                    "申请日期_dt_in": "进场日期", "表底_in": "进场表底",
                    "申请日期_dt_out": "撤场日期", "表底_out": "撤场表底"
                }, inplace=True)

                settlement = pd.Timestamp(water_settlement_date)
                month_start = settlement.replace(day=1)

                def calc_status(row):
                    in_date = row["进场日期"]
                    out_date = row["撤场日期"]
                    if pd.isna(in_date):
                        return "无进场记录"
                    if in_date > settlement:
                        return "未启用"
                    if (in_date >= month_start) and (pd.notna(out_date) and out_date <= settlement):
                        return "本月调整"
                    if (in_date < month_start) and (pd.notna(out_date) and month_start <= out_date <= settlement):
                        return "本月终止"
                    if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start):
                        return "已终止"
                    if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement):
                        return "本月启用"
                    if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement):
                        return "启用中"
                    return "异常状态"

                df_merged["电表状态"] = df_merged.apply(calc_status, axis=1)
                df_result = df_merged[
                    ["自建ID", "表号", "品牌", "客户编号", "合同号", "进场日期", "进场表底", "撤场日期", "撤场表底", "电表状态"]
                ].copy()
                df_result["进场日期"] = df_result["进场日期"].dt.strftime("%Y-%m-%d")
                df_result["撤场日期"] = df_result["撤场日期"].dt.strftime("%Y-%m-%d")
                df_result = df_result.fillna("")

                # 保存完整结果到 session_state（水费专用）
                st.session_state.water_df_result_full = df_result.copy()

                enabled_df = df_result[df_result["电表状态"].isin(["启用中", "本月启用"])]
                st.session_state.water_enabled_df = enabled_df.copy()

                status_options = sorted(df_result["电表状态"].unique().tolist())
                default_statuses = [s for s in ["启用中", "本月启用"] if s in status_options]
                with water_status_placeholder:
                    selected_statuses = st.multiselect(
                        "水表状态（可多选）",
                        options=status_options,
                        default=default_statuses,
                        key="water_status_filter"
                    )

                if selected_statuses:
                    df_result = df_result[df_result["电表状态"].isin(selected_statuses)]
                if water_filter_table_no:
                    df_result = df_result[df_result["表号"].str.contains(water_filter_table_no, na=False)]
                if water_filter_brand:
                    df_result = df_result[df_result["品牌"].str.contains(water_filter_brand, na=False)]

                st.session_state.water_detail_df = df_result.copy()
                st.dataframe(df_result, use_container_width=True)
                st.caption(f"共 {len(df_result)} 条（结账日期：{water_settlement_date}）")
            else:
                st.info("暂无清表数据，请先在“清表录入”中添加记录。")
                empty_cols = ["自建ID", "表号", "品牌", "客户编号", "进场日期", "进场表底", "撤场日期", "撤场表底",
                              "电表状态"]
                st.session_state.water_df_result_full = pd.DataFrame(columns=empty_cols)
                st.session_state.water_enabled_df = pd.DataFrame(columns=["表号", "客户编号"])
        with tab3:
            from datetime import datetime
            
            # 确保水费文件存在，并读取或创建参数表
            if not os.path.exists(water_excel_path):
                st.error(f"水费数据文件不存在：{water_excel_path}")
                st.stop()

            try:
                from openpyxl import load_workbook, Workbook
                wb = load_workbook(water_excel_path)
                if '参数表' not in wb.sheetnames:
                    # 创建参数表并写入默认结构
                    ws_param = wb.create_sheet('参数表')
                    ws_param['A1'] = '普通收入单价'
                    ws_param['B1'] = 0.0
                    ws_param['A2'] = '普通成本单价'
                    ws_param['B2'] = 0.0
                    ws_param['D1'] = '品牌'
                    ws_param['E1'] = '收入单价'
                    ws_param['G1'] = '品牌'
                    ws_param['H1'] = '成本单价'
                    wb.save(water_excel_path)
                wb.close()

                # 读取当前参数
                wb = load_workbook(water_excel_path)
                ws = wb['参数表']
                normal_income = ws['B1'].value if ws['B1'].value is not None else 0.0
                normal_cost = ws['B2'].value if ws['B2'].value is not None else 0.0
                wb.close()

                # 读取特殊单价表
                try:
                    special_income_df = pd.read_excel(water_excel_path, sheet_name='参数表', usecols='D:E', header=0,
                                                      dtype=str)
                    if special_income_df.empty:
                        special_income_df = pd.DataFrame(columns=['品牌', '收入单价'])
                    else:
                        special_income_df.columns = ['品牌', '收入单价']
                        special_income_df['收入单价'] = pd.to_numeric(special_income_df['收入单价'],
                                                                      errors='coerce').fillna(0.0)
                except Exception:
                    special_income_df = pd.DataFrame(columns=['品牌', '收入单价'])

                try:
                    special_cost_df = pd.read_excel(water_excel_path, sheet_name='参数表', usecols='G:H', header=0,
                                                    dtype=str)
                    if special_cost_df.empty:
                        special_cost_df = pd.DataFrame(columns=['品牌', '成本单价'])
                    else:
                        special_cost_df.columns = ['品牌', '成本单价']
                        special_cost_df['成本单价'] = pd.to_numeric(special_cost_df['成本单价'],
                                                                    errors='coerce').fillna(0.0)
                except Exception:
                    special_cost_df = pd.DataFrame(columns=['品牌', '成本单价'])

                # ---------- 界面布局 ----------
                col_income, col_cost, col_submit, col_dl_table = st.columns([1, 1, 0.6, 1])
                with col_income:
                    new_normal_income = st.number_input(
                        "普通收入单价（元/吨）",
                        value=float(normal_income),
                        format="%.8f",
                        step=0.00000001,
                        key="water_normal_income"
                    )
                with col_cost:
                    new_normal_cost = st.number_input(
                        "普通成本单价（元/吨）",
                        value=float(normal_cost),
                        format="%.8f",
                        step=0.00000001,
                        key="water_normal_cost"
                    )
                with col_submit:
                    st.write("")
                    st.write("")
                    submitted = st.button("提交参数", type="primary", use_container_width=True,
                                          key="water_param_submit")

                # ---------- 下载表底按钮 ----------
                with col_dl_table:
                    st.write("")
                    st.write("")
                    if st.button("下载表底", use_container_width=True, key="water_dl_meter_btn"):
                        with st.spinner("正在下载水表底数据，请稍候..."):
                            success, msg = run_download_script("能耗_表底.py")
                            if success:
                                st.success("表底数据下载成功")
                                downloaded_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗表底.xls")
                                try:
                                    row_count = append_meter_readings_from_downloaded_file_water(water_excel_path,
                                                                                                 downloaded_file)
                                    st.success(f"成功追加 {row_count} 条水表底记录到能耗_水费数据.xlsx")
                                except Exception as e:
                                    st.error(f"追加表底数据失败：{e}")
                            else:
                                st.error(f"下载表底失败：{msg}")
                    # 显示下载文件的最后修改时间
                    target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗表底.xls")
                    if os.path.exists(target_file):
                        mtime = os.path.getmtime(target_file)
                        date_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                        st.caption(f"文件更新时间：{date_str}")
                    else:
                        st.caption("文件尚未下载")

                st.markdown("---")

                col_left, col_right = st.columns(2)
                with col_left:
                    st.caption("特殊收入单价（按品牌）")
                    edited_special_income = st.data_editor(
                        special_income_df,
                        num_rows="dynamic",
                        use_container_width=True,
                        key="water_editable_income",
                        column_config={
                            "品牌": st.column_config.TextColumn("品牌"),
                            "收入单价": st.column_config.NumberColumn("收入单价", format="%.8f"),
                        }
                    )
                with col_right:
                    st.caption("特殊成本单价（按品牌）")
                    edited_special_cost = st.data_editor(
                        special_cost_df,
                        num_rows="dynamic",
                        use_container_width=True,
                        key="water_editable_cost",
                        column_config={
                            "品牌": st.column_config.TextColumn("品牌"),
                            "成本单价": st.column_config.NumberColumn("成本单价", format="%.8f"),
                        }
                    )

                if submitted:
                    try:
                        wb_write = load_workbook(water_excel_path)
                        ws_write = wb_write['参数表']

                        # 保存普通单价
                        ws_write['B1'] = new_normal_income
                        ws_write['B2'] = new_normal_cost

                        # 清空原有特殊单价区域
                        max_row = ws_write.max_row
                        for row in range(1, max_row + 1):
                            for col in [4, 5]:
                                ws_write.cell(row=row, column=col).value = None

                        # 写入特殊收入单价
                        ws_write.cell(row=1, column=4, value="品牌")
                        ws_write.cell(row=1, column=5, value="收入单价")
                        for r_idx, row in edited_special_income.iterrows():
                            ws_write.cell(row=r_idx + 2, column=4, value=row['品牌'])
                            ws_write.cell(row=r_idx + 2, column=5, value=float(row['收入单价']))

                        # 清空特殊成本单价区域
                        for row in range(1, max_row + 1):
                            for col in [7, 8]:
                                ws_write.cell(row=row, column=col).value = None

                        ws_write.cell(row=1, column=7, value="品牌")
                        ws_write.cell(row=1, column=8, value="成本单价")
                        for r_idx, row in edited_special_cost.iterrows():
                            ws_write.cell(row=r_idx + 2, column=7, value=row['品牌'])
                            ws_write.cell(row=r_idx + 2, column=8, value=float(row['成本单价']))

                        wb_write.save(water_excel_path)
                        wb_write.close()
                        st.success("水费参数保存成功！正在刷新页面...")
                        st.rerun()
                    except Exception as e:
                        st.error(f"保存失败：{e}")

            except Exception as e:
                st.error(f"读取参数表失败：{e}")
        with tab_alert:
            st.subheader("💧 水费结转")

            col1, col2, col3, col4 = st.columns([1, 1, 1, 1.5])
            with col1:
                water_carryover_date = st.date_input(
                    "结账日期",
                    value=date.today() - timedelta(days=1),   # 修改为前一天
                    key="water_carryover_date"
                )
            with col2:
                st.write("")
                st.write("")
                water_generate = st.button("🔍 生成明细", type="primary", use_container_width=True, key="water_gen_btn")
            with col3:
                st.write("")
                st.write("")
                water_has_data = not st.session_state.get("water_carryover_df", pd.DataFrame()).empty
                water_archive_clicked = st.button(
                    "💾 存档",
                    type="secondary",
                    use_container_width=True,
                    disabled=not water_has_data,
                    key="water_archive_btn"
                )
            with col4:
                st.write("")
                st.write("")
                water_voucher_template_path = os.path.join(BASE_DIR, "0-模板", "凭证模板.xlsx")
                water_download_voucher_clicked = st.button(
                    "📎 下载凭证模板",
                    type="secondary",
                    use_container_width=True,
                    disabled=not water_has_data,
                    key="water_voucher_btn"
                )

            if 'water_carryover_df' not in st.session_state:
                st.session_state.water_carryover_df = pd.DataFrame()

            # ---------- 存档逻辑 ----------
            if water_archive_clicked:
                try:
                    period_str = water_carryover_date.strftime("%Y-%m")
                    df_to_save = st.session_state.water_carryover_df.copy()
                    df_to_save["期间"] = period_str

                    # 读取现有结转留存工作表
                    if os.path.exists(water_excel_path):
                        with pd.ExcelFile(water_excel_path) as xl:
                            if "结转留存" in xl.sheet_names:
                                existing_df = pd.read_excel(water_excel_path, sheet_name="结转留存", dtype=str)
                                if "期间" in existing_df.columns:
                                    existing_df = existing_df[existing_df["期间"] != period_str]
                            else:
                                existing_df = pd.DataFrame()
                    else:
                        existing_df = pd.DataFrame()

                    combined_df = pd.concat([existing_df, df_to_save], ignore_index=True)

                    # 保存到Excel的“结转留存”工作表
                    from openpyxl import load_workbook, Workbook
                    if os.path.exists(water_excel_path):
                        wb = load_workbook(water_excel_path)
                        if "结转留存" in wb.sheetnames:
                            del wb["结转留存"]
                        ws = wb.create_sheet("结转留存")
                    else:
                        wb = Workbook()
                        ws = wb.active
                        ws.title = "结转留存"
                    # 写入表头
                    for col_idx, col_name in enumerate(combined_df.columns, 1):
                        ws.cell(row=1, column=col_idx, value=col_name)
                    for r_idx, row in combined_df.iterrows():
                        for c_idx, value in enumerate(row):
                            ws.cell(row=r_idx + 2, column=c_idx + 1, value=value)
                    wb.save(water_excel_path)
                    wb.close()
                    st.success(f"✅ 水费结转存档成功！期间：{period_str}")
                except Exception as e:
                    st.error(f"存档失败：{e}")

            # ---------- 凭证模板下载 ----------
            if water_download_voucher_clicked:
                if not os.path.exists(water_voucher_template_path):
                    st.error(f"凭证模板文件不存在：{water_voucher_template_path}")
                else:
                    current_carryover_df = st.session_state.get("water_carryover_df", pd.DataFrame())
                    if current_carryover_df.empty:
                        st.warning("暂无结转明细，请先点击「生成明细」。")
                    else:
                        try:
                            voucher_io = generate_water_voucher_from_carryover(
                                current_carryover_df,
                                water_carryover_date,
                                water_voucher_template_path
                            )
                            if voucher_io:
                                st.download_button(
                                    label="📥 点击下载水费凭证模板",
                                    data=voucher_io,
                                    file_name=f"水费凭证模板_{water_carryover_date.strftime('%Y%m%d')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_water_voucher_btn"
                                )
                            else:
                                st.error("生成凭证模板失败，请检查数据。")
                        except Exception as e:
                            st.error(f"生成凭证模板出错：{e}")

            # ---------- 生成结转明细 ----------
            if water_generate:
                if not os.path.exists(water_excel_path):
                    st.error(f"水费主数据文件不存在：{water_excel_path}")
                    st.stop()
                df_water_existing = load_excel(water_excel_path)
                if df_water_existing.empty:
                    st.warning("暂无清表录入数据，无法计算结转。")
                    st.stop()

                try:
                    # 读取表底数据
                    sheet_meter = "表底"
                    df_meter = pd.read_excel(water_excel_path, sheet_name=sheet_meter, dtype=str)
                    df_meter.columns = [c.strip() for c in df_meter.columns]
                    col_meter_no, col_meter_date, col_meter_value = df_meter.columns[:3]
                    df_meter[col_meter_date] = pd.to_datetime(df_meter[col_meter_date], errors='coerce')
                    df_meter[col_meter_value] = pd.to_numeric(
                        df_meter[col_meter_value].str.replace(',', '', regex=False).str.strip(),
                        errors='coerce'
                    ).fillna(0.0)
                    df_meter[col_meter_no] = df_meter[col_meter_no].astype(str)

                    # 读取参数表
                    from openpyxl import load_workbook
                    wb = load_workbook(water_excel_path, data_only=True)
                    if '参数表' not in wb.sheetnames:
                        st.error("水费文件缺少“参数表”工作表。")
                        st.stop()
                    ws = wb['参数表']
                    normal_income = float(ws['B1'].value) if ws['B1'].value is not None else 0.0
                    normal_cost = float(ws['B2'].value) if ws['B2'].value is not None else 0.0
                    wb.close()

                    # 特殊收入单价
                    try:
                        special_income_df = pd.read_excel(
                            water_excel_path, sheet_name='参数表', usecols='D:E', header=0, dtype=str
                        )
                        if special_income_df.empty:
                            special_income_df = pd.DataFrame(columns=['品牌', '收入单价'])
                        else:
                            special_income_df.columns = ['品牌', '收入单价']
                            special_income_df['收入单价'] = pd.to_numeric(
                                special_income_df['收入单价'].str.replace(',', '').str.strip(),
                                errors='coerce'
                            ).fillna(0.0)
                            special_income_df = special_income_df.dropna(subset=['品牌'])
                            special_income_df['品牌'] = special_income_df['品牌'].astype(str).str.strip()
                    except Exception:
                        special_income_df = pd.DataFrame(columns=['品牌', '收入单价'])

                    # 特殊成本单价
                    try:
                        special_cost_df = pd.read_excel(
                            water_excel_path, sheet_name='参数表', usecols='G:H', header=0, dtype=str
                        )
                        if special_cost_df.empty:
                            special_cost_df = pd.DataFrame(columns=['品牌', '成本单价'])
                        else:
                            special_cost_df.columns = ['品牌', '成本单价']
                            special_cost_df['成本单价'] = pd.to_numeric(
                                special_cost_df['成本单价'].str.replace(',', '').str.strip(),
                                errors='coerce'
                            ).fillna(0.0)
                            special_cost_df = special_cost_df.dropna(subset=['品牌'])
                            special_cost_df['品牌'] = special_cost_df['品牌'].astype(str).str.strip()
                    except Exception:
                        special_cost_df = pd.DataFrame(columns=['品牌', '成本单价'])

                    # 处理进场/撤场记录
                    df_in = df_water_existing[df_water_existing["清表类型"] == "进场清表"].copy()
                    df_out = df_water_existing[df_water_existing["清表类型"] == "撤场清表"].copy()
                    df_in["申请日期_dt"] = pd.to_datetime(df_in["申请日期"], errors="coerce")
                    df_out["申请日期_dt"] = pd.to_datetime(df_out["申请日期"], errors="coerce")

                    df_in_max = df_in.loc[df_in.groupby("自建ID")["申请日期_dt"].idxmax()]
                    df_out_max = df_out.loc[df_out.groupby("自建ID")["申请日期_dt"].idxmax()]

                    df_merged = pd.merge(
                        df_in_max[["自建ID", "表号", "品牌", "客户编号", "申请日期_dt", "表底"]],
                        df_out_max[["自建ID", "申请日期_dt", "表底"]],
                        on="自建ID", how="outer", suffixes=("_in", "_out")
                    )
                    df_merged.rename(columns={
                        "申请日期_dt_in": "进场日期", "表底_in": "进场表底",
                        "申请日期_dt_out": "撤场日期", "表底_out": "撤场表底"
                    }, inplace=True)

                    settlement = pd.Timestamp(water_carryover_date)
                    month_start = settlement.replace(day=1)

                    def calc_status(row):
                        in_date = row["进场日期"]
                        out_date = row["撤场日期"]
                        if pd.isna(in_date):
                            return "无进场记录"
                        if in_date > settlement:
                            return "未启用"
                        if (in_date >= month_start) and (pd.notna(out_date) and out_date <= settlement):
                            return "本月调整"
                        if (in_date < month_start) and (pd.notna(out_date) and month_start <= out_date <= settlement):
                            return "本月终止"
                        if (in_date < month_start) and (pd.notna(out_date) and out_date < month_start):
                            return "已终止"
                        if (in_date >= month_start) and (pd.isna(out_date) or out_date > settlement):
                            return "本月启用"
                        if (in_date < month_start) and (pd.isna(out_date) or out_date > settlement):
                            return "启用中"
                        return "异常状态"

                    df_merged["水表状态"] = df_merged.apply(calc_status, axis=1)
                    carryover_statuses = ["启用中", "本月启用", "本月终止", "本月调整"]
                    df_filtered = df_merged[df_merged["水表状态"].isin(carryover_statuses)].copy()

                    if df_filtered.empty:
                        st.info("当前结账日期下没有需要结转的水表。")
                        st.session_state.water_carryover_df = pd.DataFrame()
                        st.stop()

                    def get_start_date(row):
                        return max(row["进场日期"], month_start) if pd.notna(row["进场日期"]) else month_start

                    def get_end_date(row):
                        if row["水表状态"] in ("本月终止", "本月调整") and pd.notna(row["撤场日期"]):
                            return row["撤场日期"]
                        else:
                            return settlement + pd.Timedelta(days=1)

                    df_filtered["开始日期"] = df_filtered.apply(get_start_date, axis=1)
                    df_filtered["结束日期"] = df_filtered.apply(get_end_date, axis=1)

                    def get_meter_value(table_no, target_date):
                        if pd.isna(target_date):
                            return 0.0
                        match = df_meter[
                            (df_meter[col_meter_no] == str(table_no)) &
                            (df_meter[col_meter_date] == target_date)
                            ]
                        if not match.empty:
                            return match[col_meter_value].iloc[0]
                        return 0.0

                    start_values = [get_meter_value(r["表号"], r["开始日期"]) for _, r in df_filtered.iterrows()]
                    end_values = [get_meter_value(r["表号"], r["结束日期"]) for _, r in df_filtered.iterrows()]
                    df_filtered["开始表底"] = start_values
                    df_filtered["结束表底"] = end_values
                    df_filtered["用水量"] = df_filtered["结束表底"] - df_filtered["开始表底"]

                    def get_income_price(brand):
                        row = special_income_df[special_income_df['品牌'] == str(brand).strip()]
                        if not row.empty:
                            return row['收入单价'].iloc[0]
                        return normal_income

                    def get_cost_price(brand):
                        row = special_cost_df[special_cost_df['品牌'] == str(brand).strip()]
                        if not row.empty:
                            return row['成本单价'].iloc[0]
                        return normal_cost

                    df_filtered["收入单价"] = df_filtered["品牌"].apply(get_income_price)
                    df_filtered["成本单价"] = df_filtered["品牌"].apply(get_cost_price)

                    # 水费税率 9%
                    df_filtered["收入_含税"] = df_filtered["用水量"] * df_filtered["收入单价"]
                    df_filtered["收入_不含税"] = df_filtered["收入_含税"] / 1.09
                    df_filtered["成本_含税"] = df_filtered["用水量"] * df_filtered["成本单价"]
                    df_filtered["成本_不含税"] = df_filtered["成本_含税"] / 1.09

                    result = df_filtered[[
                        "水表状态", "表号", "品牌", "客户编号",
                        "开始表底", "结束表底", "用水量",
                        "收入单价", "收入_含税", "收入_不含税",
                        "成本单价", "成本_含税", "成本_不含税",
                        "开始日期", "结束日期"
                    ]].copy()

                    for col in ["开始表底", "结束表底", "用水量",
                                "收入_含税", "收入_不含税", "成本_含税", "成本_不含税"]:
                        result[col] = result[col].round(2)

                    result["开始日期"] = result["开始日期"].dt.strftime("%Y-%m-%d")
                    result["结束日期"] = result["结束日期"].dt.strftime("%Y-%m-%d")

                    st.session_state.water_carryover_df = result

                    st.success("✅ 水费结转明细生成成功！")
                    st.dataframe(result, use_container_width=True, hide_index=True)

                    total_blocks = len(result)
                    total_water = result["用水量"].sum()
                    total_inc_tax = result["收入_含税"].sum()
                    total_inc_notax = result["收入_不含税"].sum()
                    total_cost_tax = result["成本_含税"].sum()
                    total_cost_notax = result["成本_不含税"].sum()

                    st.info(
                        f"📊 本次结转水表 **{total_blocks}** 块，"
                        f"合计用水量 **{total_water:.2f}** 吨，"
                        f"含税收入 **{total_inc_tax:.2f}** 元，"
                        f"不含税收入 **{total_inc_notax:.2f}** 元，"
                        f"含税成本 **{total_cost_tax:.2f}** 元，"
                        f"不含税成本 **{total_cost_notax:.2f}** 元。"
                    )

                    from io import BytesIO
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result.to_excel(writer, index=False, sheet_name='水费结转明细')
                    output.seek(0)
                    st.download_button(
                        label="📥 下载水费结转明细",
                        data=output,
                        file_name=f"水费结转_{water_carryover_date.strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_water_carryover"
                    )

                except Exception as e:
                    st.error(f"生成水费结转明细时出错：{e}")
                    st.session_state.water_carryover_df = pd.DataFrame()

            elif not st.session_state.water_carryover_df.empty:
                result = st.session_state.water_carryover_df
                st.success("✅ 水费结转明细已生成（上次结果）")
                st.dataframe(result, use_container_width=True, hide_index=True)

                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result.to_excel(writer, index=False, sheet_name='水费结转明细')
                output.seek(0)
                st.download_button(
                    label="📥 下载水费结转明细",
                    data=output,
                    file_name=f"水费结转_{water_carryover_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_water_carryover_refresh"
                )
            else:
                st.info("请选择结账日期并点击“生成明细”按钮。")
        with tab_balance:
            from openpyxl import load_workbook, Workbook
            from datetime import datetime
            from io import BytesIO

            st.subheader("💰 水费余额 - 收款明细导入及余额查询")

            # 两列布局：左侧上传文件，右侧期间选择和生成按钮
            col_upload, col_period = st.columns([1, 1])
            with col_upload:
                uploaded_file = st.file_uploader(
                    "上传本月收款明细（.xlsx）",
                    type=["xlsx"],
                    key="water_balance_upload"
                )
                if uploaded_file is not None:
                    # 处理上传文件的逻辑（保留原有解析和写入功能，但不显示表格）
                    try:
                        df_raw = pd.read_excel(uploaded_file, dtype=str)
                        required_cols = ['业务日期', '核算项目', '贷方']
                        missing = [c for c in required_cols if c not in df_raw.columns]
                        if missing:
                            st.error(f"上传文件缺少必要列：{missing}，请检查模板。")
                        else:
                            df_raw = df_raw[
                                ~df_raw['凭证字号'].astype(str).str.contains('期初余额|本期合计|本年累计', na=False)]
                            df_raw = df_raw[df_raw['贷方'].notna() & (df_raw['贷方'] != '')]

                            def parse_customer(info):
                                if pd.isna(info):
                                    return None, None
                                s = str(info).strip()
                                if not s.startswith('客户:'):
                                    return None, None
                                rest = s[3:].strip()
                                parts = rest.split(' ', 1)
                                if len(parts) == 2:
                                    return parts[0].strip(), parts[1].strip()
                                return rest, ''

                            df_raw[['客户编号', '客户名称']] = df_raw['核算项目'].apply(
                                lambda x: pd.Series(parse_customer(x))
                            )
                            df_raw = df_raw[df_raw['客户编号'].notna()]
                            if df_raw.empty:
                                st.warning("未识别到任何有效的收款记录，请检查核算项目格式。")
                            else:
                                df_raw['收款金额'] = (
                                    df_raw['贷方']
                                    .astype(str)
                                    .str.replace(',', '', regex=False)
                                    .str.replace(' ', '', regex=False)
                                )
                                df_raw['收款金额'] = pd.to_numeric(df_raw['收款金额'], errors='coerce').fillna(0)
                                df_raw['期间_dt'] = pd.to_datetime(df_raw['业务日期'], errors='coerce')
                                df_raw['期间'] = df_raw['期间_dt'].dt.date
                                df_raw = df_raw[df_raw['期间'].notna()]
                                df_balance = df_raw[['客户编号', '客户名称', '期间', '收款金额']].drop_duplicates()
                                df_balance['收款金额'] = pd.to_numeric(df_balance['收款金额'], errors='coerce').fillna(
                                    0)

                                if not df_balance.empty:
                                    # 确定本次导入涉及的所有月份
                                    df_balance['月份'] = pd.to_datetime(df_balance['期间']).dt.strftime('%Y-%m')
                                    import_months = df_balance['月份'].unique().tolist()

                                    # 读取现有收款记录
                                    existing_df = pd.DataFrame()
                                    if os.path.exists(water_excel_path):
                                        try:
                                            with pd.ExcelFile(water_excel_path) as xl:
                                                if "收款记录" in xl.sheet_names:
                                                    existing_df = pd.read_excel(water_excel_path, sheet_name="收款记录")
                                                    if '期间' in existing_df.columns:
                                                        existing_df['期间'] = pd.to_datetime(existing_df['期间'],
                                                                                             errors='coerce').dt.date
                                        except Exception as e:
                                            st.warning(f"读取现有收款记录失败：{e}")

                                    # 保留不在本次导入月份中的记录
                                    if not existing_df.empty:
                                        existing_df['月份'] = pd.to_datetime(existing_df['期间']).dt.strftime('%Y-%m')
                                        keep_df = existing_df[~existing_df['月份'].isin(import_months)].copy()
                                        keep_df.drop(columns=['月份'], inplace=True)
                                    else:
                                        keep_df = pd.DataFrame()

                                    final_df = pd.concat([keep_df, df_balance], ignore_index=True)
                                    final_df = final_df[final_df['期间'].notna()]
                                    final_df['收款金额'] = pd.to_numeric(final_df['收款金额'], errors='coerce').fillna(
                                        0)

                                    # 写入工作簿
                                    if os.path.exists(water_excel_path):
                                        wb = load_workbook(water_excel_path)
                                        if "收款记录" in wb.sheetnames:
                                            del wb["收款记录"]
                                        ws = wb.create_sheet("收款记录")
                                    else:
                                        wb = Workbook()
                                        ws = wb.active
                                        ws.title = "收款记录"

                                    ws.cell(row=1, column=1, value="客户编号")
                                    ws.cell(row=1, column=2, value="客户名称")
                                    ws.cell(row=1, column=3, value="期间")
                                    ws.cell(row=1, column=4, value="收款金额")
                                    col_letter = ws.cell(row=1, column=3).column_letter
                                    ws.column_dimensions[col_letter].number_format = 'yyyy-mm-dd'

                                    for i, row in final_df.iterrows():
                                        ws.cell(row=i + 2, column=1, value=str(row['客户编号']))
                                        ws.cell(row=i + 2, column=2, value=str(row['客户名称']))
                                        ws.cell(row=i + 2, column=3, value=row['期间'])
                                        ws.cell(row=i + 2, column=4, value=float(row['收款金额']))

                                    wb.save(water_excel_path)
                                    wb.close()
                                    st.success(
                                        f"收款记录已更新：删除了月份 {', '.join(import_months)} 的旧数据，新增 {len(df_balance)} 条记录。")
                                else:
                                    st.warning("解析后无有效数据，未写入。")
                    except Exception as e:
                        st.error(f"处理上传文件失败：{e}")

            with col_period:
                st.write("")
                st.write("")
                period_input = st.text_input(
                    "期间 (YYYY-MM)",
                    value=datetime.now().strftime("%Y-%m"),
                    key="balance_period_input",
                    placeholder="例如：2026-05"
                )
                generate_btn = st.button("生成余额表", type="primary", use_container_width=True)

            # 显示余额表
            if generate_btn:
                import re
                if not re.match(r'^\d{4}-\d{2}$', period_input):
                    st.error("期间格式错误，请输入 YYYY-MM 格式（如 2026-05）")
                else:
                    try:
                        year, month = map(int, period_input.split('-'))
                        current_period_start = datetime(year, month, 1).date()
                        if month == 12:
                            next_year, next_month = year + 1, 1
                        else:
                            next_year, next_month = year, month + 1
                        next_month_start = datetime(next_year, next_month, 1).date()

                        # 1. 加载收款记录（金额已是不含税）
                        payment_df = pd.DataFrame()
                        if os.path.exists(water_excel_path):
                            try:
                                with pd.ExcelFile(water_excel_path) as xl:
                                    if "收款记录" in xl.sheet_names:
                                        payment_df = pd.read_excel(water_excel_path, sheet_name="收款记录")
                                        if '期间' in payment_df.columns:
                                            payment_df['期间'] = pd.to_datetime(payment_df['期间'],
                                                                                errors='coerce').dt.date
                                        if '收款金额' in payment_df.columns:
                                            payment_df['收款金额'] = pd.to_numeric(payment_df['收款金额'],
                                                                                   errors='coerce').fillna(0)
                            except Exception as e:
                                st.warning(f"读取收款记录失败：{e}")

                        # 2. 加载结转留存（成本_不含税）
                        carryover_df = pd.DataFrame()
                        if os.path.exists(water_excel_path):
                            try:
                                with pd.ExcelFile(water_excel_path) as xl:
                                    if "结转留存" in xl.sheet_names:
                                        carryover_df = pd.read_excel(water_excel_path, sheet_name="结转留存")
                                        if '期间' not in carryover_df.columns:
                                            st.error("结转留存工作表中缺少“期间”列，无法计算。")
                                            st.stop()
                                        if '成本_不含税' in carryover_df.columns:
                                            carryover_df['成本_不含税'] = pd.to_numeric(carryover_df['成本_不含税'],
                                                                                        errors='coerce').fillna(0)
                            except Exception as e:
                                st.warning(f"读取结转留存失败：{e}")

                        # 3. 加载客户资料
                        cust_df = pd.DataFrame()
                        if os.path.exists(customer_info_path):
                            cust_df = pd.read_excel(customer_info_path, dtype=str)[['编码', '名称']]
                            cust_df.columns = ['客户编号', '客户名称']
                        else:
                            st.warning(f"客户资料文件不存在：{customer_info_path}，客户名称将显示为客户编号")

                        all_customers = set()
                        if not payment_df.empty:
                            all_customers.update(payment_df['客户编号'].astype(str))
                        if not carryover_df.empty:
                            all_customers.update(carryover_df['客户编号'].astype(str))
                        all_customers = sorted(all_customers)

                        if not all_customers:
                            st.info("暂无任何收款或结转数据，无法生成余额表。")
                        else:
                            result_rows = []
                            for cust_id in all_customers:
                                # 本月收款（不含税，期间属于当前月份）
                                current_payment_notax = 0.0
                                if not payment_df.empty:
                                    cust_payments = payment_df[payment_df['客户编号'].astype(str) == cust_id]
                                    current_payments = cust_payments[
                                        (cust_payments['期间'] >= current_period_start) & (
                                                    cust_payments['期间'] < next_month_start)
                                        ]
                                    current_payment_notax = current_payments['收款金额'].sum()

                                # 上月累计收款（不含税，期间 < current_period_start）
                                prev_payment_notax = 0.0
                                if not payment_df.empty:
                                    cust_payments = payment_df[payment_df['客户编号'].astype(str) == cust_id]
                                    prev_payments = cust_payments[cust_payments['期间'] < current_period_start]
                                    prev_payment_notax = prev_payments['收款金额'].sum()

                                # 本月结转（成本_不含税，期间 == period_input）
                                current_carry_notax = 0.0
                                if not carryover_df.empty:
                                    cust_carry = carryover_df[carryover_df['客户编号'].astype(str) == cust_id]
                                    for _, row in cust_carry.iterrows():
                                        if str(row['期间']).strip() == period_input:
                                            current_carry_notax += row['成本_不含税']

                                # 上月累计结转（成本_不含税，期间 < period_input）
                                prev_carry_notax = 0.0
                                if not carryover_df.empty:
                                    cust_carry = carryover_df[carryover_df['客户编号'].astype(str) == cust_id]
                                    for _, row in cust_carry.iterrows():
                                        period_str = str(row['期间']).strip()
                                        if period_str < period_input:
                                            prev_carry_notax += row['成本_不含税']

                                last_month_balance_notax = prev_payment_notax - prev_carry_notax
                                current_month_balance_notax = last_month_balance_notax + current_payment_notax - current_carry_notax
                                current_month_balance_tax = current_month_balance_notax * 1.09

                                cust_name = cust_id
                                if not cust_df.empty:
                                    name_match = cust_df[cust_df['客户编号'].astype(str) == cust_id]
                                    if not name_match.empty:
                                        cust_name = name_match.iloc[0]['客户名称']

                                result_rows.append({
                                    '客户编号': cust_id,
                                    '客户名称': cust_name,
                                    '上月余额_不含税': round(last_month_balance_notax, 2),
                                    '本月收款_不含税': round(current_payment_notax, 2),
                                    '本月结转_不含税': round(current_carry_notax, 2),
                                    '本月余额_不含税': round(current_month_balance_notax, 2),
                                    '本月余额_含税': round(current_month_balance_tax, 2)
                                })

                            result_df = pd.DataFrame(result_rows)
                            result_df = result_df.sort_values('本月余额_含税', ascending=False)

                            if not result_df.empty:
                                total_row = {
                                    '客户编号': '合计',
                                    '客户名称': '',
                                    '上月余额_不含税': result_df['上月余额_不含税'].sum(),
                                    '本月收款_不含税': result_df['本月收款_不含税'].sum(),
                                    '本月结转_不含税': result_df['本月结转_不含税'].sum(),
                                    '本月余额_不含税': result_df['本月余额_不含税'].sum(),
                                    '本月余额_含税': result_df['本月余额_含税'].sum()
                                }
                                result_df_with_total = pd.concat([result_df, pd.DataFrame([total_row])],
                                                                 ignore_index=True)
                            else:
                                result_df_with_total = result_df

                            st.dataframe(result_df_with_total, use_container_width=True, hide_index=True)
                            st.caption(f"共 {len(result_df)} 个客户，数据期间截止：{period_input}")

                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                result_df_with_total.to_excel(writer, index=False, sheet_name='水费余额表')
                            output.seek(0)
                            st.download_button(
                                label="📥 下载余额表 (Excel)",
                                data=output,
                                file_name=f"水费余额表_{period_input}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_balance_table"
                            )

                    except Exception as e:
                        st.error(f"生成余额表失败：{e}")
    # ========== 数据看板 ==========
    elif current_page == "数据看板":
        tab_point, tab_other = st.tabs(["📈 积分数据", "📊 其他看板（预留）"])

        with tab_point:

            db_file = os.path.join(BASE_DIR, "4-报表中心", "积分列表.db")
            points_file = db_file.replace('.db', '.xlsx')
            
            # 删除不再需要的Excel文件（已迁移到数据库）
            if os.path.exists(points_file):
                try:
                    os.remove(points_file)
                except:
                    pass
            
            # 获取文件修改时间作为缓存键，实现实时更新
            db_mtime = get_file_mod_time(db_file)
            
            df_points = load_points_data(points_file, db_mtime, 0)
            if df_points is None:
                st.error(f"❌ 积分数据文件不存在：{db_file}\n\n请确认数据库文件位于「4-报表中心」文件夹下。")
            else:
                try:
                    # 根据积分方式字段判断获取与消耗
                    # 消耗积分：积分方式 == "积分消耗"
                    # 获取积分：积分方式 != "积分消耗"（所有其他取值）
                    df_consume = df_points[df_points['积分方式'] == '积分消耗'].copy()
                    df_gain = df_points[df_points['积分方式'] != '积分消耗'].copy()

                    today = date.today()
                    yesterday = today - timedelta(days=1)
                    current_year = today.year
                    current_month = today.month
                    current_week_start = today - timedelta(days=today.weekday())

                    mask_yesterday = (df_gain['积分时间'].dt.date == yesterday)
                    mask_month = (df_gain['积分时间'].dt.year == current_year) & (df_gain['积分时间'].dt.month == current_month)
                    mask_week = (df_gain['积分时间'].dt.date >= current_week_start)

                    yesterday_points = df_gain.loc[mask_yesterday, '积分数'].sum()
                    month_points = df_gain.loc[mask_month, '积分数'].sum()
                    week_points = df_gain.loc[mask_week, '积分数'].sum()

                    month_consume_df = df_consume[
                        (df_consume['积分时间'].dt.year == current_year) & 
                        (df_consume['积分时间'].dt.month == current_month)
                    ]
                    month_consume = month_consume_df['积分数'].abs().sum() if not month_consume_df.empty else 0
                    
                    # 昨日消耗积分
                    mask_yesterday_consume = (df_consume['积分时间'].dt.date == yesterday)
                    yesterday_consume = df_consume.loc[mask_yesterday_consume, '积分数'].abs().sum()

                    total_members = df_points['会员卡号'].nunique()
                    # 精确筛选本月（包含年份判断），避免跨年度数据干扰
                    active_members_df = df_gain[
                        (df_gain['积分时间'].dt.year == current_year) & 
                        (df_gain['积分时间'].dt.month == current_month)
                    ]
                    active_members = active_members_df['会员卡号'].nunique()

                    # 数据更新时间显示（显示数据库中积分时间的最大日期）
                    max_points_date = df_points['积分时间'].max()
                    if pd.notna(max_points_date):
                        update_time_str = max_points_date.strftime("%Y-%m-%d")
                        st.markdown(f"""
                        <div style="
                            display: flex;
                            justify-content: space-between;
                            align-items: center;
                            padding: 0.5rem 1rem;
                            background: #f8fafc;
                            border-radius: 8px;
                            margin-bottom: 1rem;
                        ">
                            <span style="color: #94a3b8; font-size: 0.85rem;">
                                🔄 数据更新时间: {update_time_str}
                            </span>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("### 📊 核心运营指标")
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.markdown(f"""
                        <div style="
                            background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
                            padding: 1.25rem;
                            border-radius: 10px;
                            color: white;
                            box-shadow: 0 2px 8px rgba(72,187,120,0.3);
                        ">
                            <div style="font-size: 0.8rem; opacity: 0.9; margin-bottom: 0.5rem;">📅 昨日新增积分</div>
                            <div style="font-size: 1.8rem; font-weight: bold;">{yesterday_points:,.0f}</div>
                            <div style="font-size: 0.75rem; opacity: 0.8; margin-top: 0.5rem;">分</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col2:
                        st.markdown(f"""
                        <div style="
                            background: linear-gradient(135deg, #f56565 0%, #e53e3e 100%);
                            padding: 1.25rem;
                            border-radius: 10px;
                            color: white;
                            box-shadow: 0 2px 8px rgba(245,101,101,0.3);
                        ">
                            <div style="font-size: 0.8rem; opacity: 0.9; margin-bottom: 0.5rem;">💸 昨日消耗积分</div>
                            <div style="font-size: 1.8rem; font-weight: bold;">{yesterday_consume:,.0f}</div>
                            <div style="font-size: 0.75rem; opacity: 0.8; margin-top: 0.5rem;">分</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col3:
                        st.markdown(f"""
                        <div style="
                            background: linear-gradient(135deg, #4299e1 0%, #3182ce 100%);
                            padding: 1.25rem;
                            border-radius: 10px;
                            color: white;
                            box-shadow: 0 2px 8px rgba(66,153,225,0.3);
                        ">
                            <div style="font-size: 0.8rem; opacity: 0.9; margin-bottom: 0.5rem;">🗓️ 本月累计积分</div>
                            <div style="font-size: 1.8rem; font-weight: bold;">{month_points:,.0f}</div>
                            <div style="font-size: 0.75rem; opacity: 0.8; margin-top: 0.5rem;">分</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col4:
                        st.markdown(f"""
                        <div style="
                            background: linear-gradient(135deg, #ed8936 0%, #dd6b20 100%);
                            padding: 1.25rem;
                            border-radius: 10px;
                            color: white;
                            box-shadow: 0 2px 8px rgba(237,137,54,0.3);
                        ">
                            <div style="font-size: 0.8rem; opacity: 0.9; margin-bottom: 0.5rem;">📤 本月消耗积分</div>
                            <div style="font-size: 1.8rem; font-weight: bold;">{month_consume:,.0f}</div>
                            <div style="font-size: 0.75rem; opacity: 0.8; margin-top: 0.5rem;">分</div>
                        </div>
                        """, unsafe_allow_html=True)
                    with col5:
                        st.markdown(f"""
                        <div style="
                            background: linear-gradient(135deg, #9f7aea 0%, #805ad5 100%);
                            padding: 1.25rem;
                            border-radius: 10px;
                            color: white;
                            box-shadow: 0 2px 8px rgba(159,122,234,0.3);
                        ">
                            <div style="font-size: 0.8rem; opacity: 0.9; margin-bottom: 0.5rem;">👥 活跃会员数</div>
                            <div style="font-size: 1.8rem; font-weight: bold;">{active_members:,}</div>
                            <div style="font-size: 0.75rem; opacity: 0.8; margin-top: 0.5rem;">人 / 总会员 {total_members:,}</div>
                        </div>
                        """, unsafe_allow_html=True)

                    st.markdown("---")

                    col_filter1, col_filter2, col_filter3 = st.columns([1, 1, 1])
                    with col_filter1:
                        date_range = st.date_input(
                            "📅 选择日期范围",
                            value=(today - timedelta(days=30), today),
                            key="points_date_range"
                        )
                    with col_filter2:
                        all_businesses = sorted(df_points['业态'].dropna().unique().tolist())
                        selected_businesses = st.multiselect(
                            "🏢 业态筛选",
                            options=all_businesses,
                            default=all_businesses,
                            key="points_business"
                        )
                    with col_filter3:
                        member_levels = ['全部'] + sorted(df_points['等级'].dropna().unique().tolist())
                        selected_level = st.selectbox("⭐ 会员等级", member_levels, key="points_level")

                    if len(date_range) == 2:
                        start_date, end_date = date_range
                        mask_date = (df_gain['积分时间'].dt.date >= start_date) & (df_gain['积分时间'].dt.date <= end_date)
                        df_filtered = df_gain[mask_date].copy()
                    else:
                        df_filtered = df_gain.copy()

                    if selected_businesses:
                        df_filtered = df_filtered[df_filtered['业态'].isin(selected_businesses)]
                    if selected_level != '全部':
                        df_filtered = df_filtered[df_filtered['等级'] == selected_level]

                    st.markdown("### 📈 数据可视化分析")

                    tab_chart1, tab_chart2, tab_chart3, tab_chart4, tab_chart5, tab_chart6 = st.tabs([
                        "📊 积分趋势", 
                        "🥧 类型分布", 
                        "👥 会员等级",
                        "⚠️ 积分预警",
                        "💸 积分消耗",
                        "📋 详细数据查询"
                    ])

                    with tab_chart1:
                        if not df_filtered.empty:
                            df_filtered_copy = df_filtered.copy()
                            df_filtered_copy['日期'] = df_filtered_copy['积分时间'].dt.date
                            daily_points = df_filtered_copy.groupby('日期')['积分数'].sum().reset_index()
                            daily_points = daily_points.sort_values('日期')

                            import plotly.express as px
                            
                            # 每日积分获取趋势（折线图）和业态积分贡献 TOP 10（横向柱状图）放在同一行
                            col_main1, col_main2 = st.columns(2)
                            
                            with col_main1:
                                fig_line = px.line(
                                    daily_points,
                                    x='日期',
                                    y='积分数',
                                    title=f'每日积分获取趋势（{start_date} 至 {end_date}）',
                                    labels={'积分数': '积分数量', '日期': '日期'},
                                    markers=True
                                )
                                fig_line.update_layout(
                                    plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.3,
                                    xaxis_title_font_size=12,
                                    yaxis_title_font_size=12,
                                    xaxis=dict(tickformat='%Y-%m-%d')
                                )
                                fig_line.update_traces(line=dict(color='#3182ce', width=2), marker=dict(size=6, color='#2c5282'))
                                st.plotly_chart(fig_line, use_container_width=True)

                            with col_main2:
                                if '业态' in df_filtered.columns:
                                    biz_dist = df_filtered[df_filtered['业态'] != '--'].groupby('业态')['积分数'].sum().reset_index()
                                    biz_dist = biz_dist.sort_values('积分数', ascending=False)
                                    fig_biz = px.bar(
                                        biz_dist.head(10),
                                        y='业态',
                                        x='积分数',
                                        orientation='h',
                                        title='业态积分贡献',
                                        text='积分数',
                                        color='积分数',
                                        color_continuous_scale='blues'
                                    )
                                    fig_biz.update_layout(
                                        plot_bgcolor='white',
                                        paper_bgcolor='white',
                                        title_font_size=14,
                                        title_x=0.5,
                                        showlegend=False
                                    )
                                    fig_biz.update_traces(textposition='outside')
                                    st.plotly_chart(fig_biz, use_container_width=True)

                            col_chart1, col_chart2 = st.columns(2)
                            with col_chart1:
                                if '消费金额(元)' in df_filtered_copy.columns:
                                    daily_amount = df_filtered_copy.groupby('日期')['消费金额(元)'].sum().reset_index()
                                    daily_amount = daily_amount.sort_values('日期')
                                    fig_bar = px.bar(
                                        daily_amount,
                                        x='日期',
                                        y='消费金额(元)',
                                        title='每日消费金额趋势',
                                        labels={'消费金额(元)': '金额（元）', '日期': '日期'}
                                    )
                                    fig_bar.update_layout(
                                        plot_bgcolor='white',
                                        paper_bgcolor='white',
                                        title_font_size=14,
                                        title_x=0.5,
                                        xaxis=dict(tickformat='%Y-%m-%d')
                                    )
                                    fig_bar.update_traces(marker_color='#48bb78')
                                    st.plotly_chart(fig_bar, use_container_width=True)

                            with col_chart2:
                                hourly_dist = df_filtered_copy.copy()
                                hourly_dist['小时'] = hourly_dist['积分时间'].dt.hour
                                hourly_points = hourly_dist.groupby('小时')['积分数'].sum().reset_index()
                                fig_hourly = px.bar(
                                    hourly_points,
                                    x='小时',
                                    y='积分数',
                                    title='积分获取时段分布',
                                    labels={'积分数': '积分数量', '小时': '小时'},
                                    color='积分数',
                                    color_continuous_scale='blues'
                                )
                                fig_hourly.update_layout(
                                    plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.5,
                                    xaxis=dict(tickmode='array', tickvals=list(range(0, 24)))
                                )
                                st.plotly_chart(fig_hourly, use_container_width=True)
                            
                            # 租户积分排行榜 TOP 20
                            st.markdown("#### 🏪 租户积分排行榜 TOP 20")
                            if '租户' in df_filtered.columns:
                                tenant_points = df_filtered[df_filtered['租户'] != '--'].groupby('租户').agg({
                                    '积分数': 'sum',
                                    '会员卡号': 'nunique',
                                    '消费金额(元)': 'sum'
                                }).reset_index()
                                tenant_points.columns = ['租户', '总积分', '会员数', '总消费金额']
                                tenant_points = tenant_points.sort_values('总积分', ascending=False).head(20)
                                # 保留整数并添加千分符
                                tenant_points['总积分'] = tenant_points['总积分'].round(0).astype(int).apply(lambda x: f'{x:,}')
                                tenant_points['会员数'] = tenant_points['会员数'].astype(int).apply(lambda x: f'{x:,}')
                                tenant_points['总消费金额'] = tenant_points['总消费金额'].round(0).astype(int).apply(lambda x: f'{x:,}')
                                st.dataframe(tenant_points, use_container_width=True, hide_index=True)
                        else:
                            st.info("当前筛选条件下没有数据")

                    with tab_chart2:
                        if not df_filtered.empty and '积分类型' in df_filtered.columns:
                            type_dist = df_filtered.groupby('积分类型')['积分数'].sum().reset_index()
                            type_dist = type_dist.sort_values('积分数', ascending=False)

                            col_pie1, col_pie2 = st.columns(2)
                            with col_pie1:
                                fig_pie = px.pie(
                                    type_dist,
                                    values='积分数',
                                    names='积分类型',
                                    title='积分类型分布占比',
                                    hole=0.4,
                                    color_discrete_sequence=px.colors.qualitative.Set2
                                )
                                fig_pie.update_layout(
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.5,
                                    legend=dict(orientation="h", yanchor="bottom", y=-0.2)
                                )
                                st.plotly_chart(fig_pie, use_container_width=True)

                            with col_pie2:
                                fig_bar_type = px.bar(
                                    type_dist,
                                    x='积分类型',
                                    y='积分数',
                                    title='各类型积分获取量排名',
                                    text='积分数',
                                    color='积分数',
                                    color_continuous_scale='viridis'
                                )
                                fig_bar_type.update_layout(
                                    plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.5,
                                    showlegend=False
                                )
                                fig_bar_type.update_traces(textposition='outside')
                                st.plotly_chart(fig_bar_type, use_container_width=True)

                            st.markdown("#### 📋 积分类型详细统计")
                            # 使用完整数据进行统计（受日期筛选影响，不受业态和等级筛选影响）
                            # 先根据日期范围筛选数据
                            if len(date_range) == 2:
                                start_date, end_date = date_range
                                mask_date = (df_points['积分时间'].dt.date >= start_date) & (df_points['积分时间'].dt.date <= end_date)
                                df_points_filtered = df_points[mask_date].copy()
                            else:
                                df_points_filtered = df_points.copy()
                            
                            # 增加"类型"列：严格按照积分方式判断
                            # 获取积分：积分方式 != "积分消耗"
                            # 消耗积分：积分方式 == "积分消耗"
                            df_points_filtered['类型'] = df_points_filtered['积分方式'].apply(
                                lambda x: '消耗' if x == '积分消耗' else '获取'
                            )
                            
                            # 按类型和积分类型分组统计
                            type_detail = df_points_filtered.groupby(['类型', '积分类型']).agg({
                                '积分数': ['sum', 'count', 'mean']
                            }).reset_index()
                            type_detail.columns = ['类型', '积分类型', '总积分', '次数', '平均积分']
                            
                            # 保留整数
                            type_detail['总积分'] = type_detail['总积分'].round(0).astype(int)
                            type_detail['次数'] = type_detail['次数'].astype(int)
                            type_detail['平均积分'] = type_detail['平均积分'].round(0).astype(int)
                            
                            # 添加千分符格式化
                            type_detail['总积分'] = type_detail['总积分'].apply(
                                lambda x: f'{x:,}' if isinstance(x, (int, float)) else x
                            )
                            type_detail['次数'] = type_detail['次数'].apply(
                                lambda x: f'{x:,}' if isinstance(x, (int, float)) else x
                            )
                            type_detail['平均积分'] = type_detail['平均积分'].apply(
                                lambda x: f'{x:,}' if isinstance(x, (int, float)) else x
                            )
                            
                            # 排序：先按类型（获取→消耗），再按总积分的绝对值降序
                            type_detail['类型排序'] = type_detail['类型'].map({'获取': 0, '消耗': 1})
                            # 先转换回数值用于排序
                            type_detail['总积分数值'] = type_detail['总积分'].str.replace(',', '').astype(int)
                            type_detail = type_detail.sort_values(['类型排序', '总积分数值'], 
                                                               key=lambda x: x.abs() if x.name == '总积分数值' else x,
                                                               ascending=[True, False])
                            type_detail = type_detail.drop(columns=['类型排序', '总积分数值'])
                            
                            st.dataframe(type_detail, use_container_width=True, hide_index=True)
                        else:
                            st.info("暂无积分类型数据")

                    # ========== 会员等级分析 ==========
                    with tab_chart3:
                        if not df_filtered.empty and '等级' in df_filtered.columns:
                            col_level1, col_level2 = st.columns(2)
                            
                            with col_level1:
                                # 各等级会员数量分布
                                level_member_count = df_filtered.groupby('等级')['会员卡号'].nunique().reset_index()
                                level_member_count.columns = ['会员等级', '会员数量']
                                level_member_count = level_member_count.sort_values('会员数量', ascending=False)
                                
                                fig_level_count = px.bar(
                                    level_member_count,
                                    x='会员等级',
                                    y='会员数量',
                                    title='各等级会员数量分布',
                                    text='会员数量',
                                    color='会员数量',
                                    color_continuous_scale='blues'
                                )
                                fig_level_count.update_layout(
                                    plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.5,
                                    showlegend=False
                                )
                                st.plotly_chart(fig_level_count, use_container_width=True)
                            
                            with col_level2:
                                # 各等级积分贡献
                                level_points = df_filtered.groupby('等级')['积分数'].sum().reset_index()
                                level_points.columns = ['会员等级', '总积分']
                                level_points = level_points.sort_values('总积分', ascending=False)
                                
                                fig_level_points = px.pie(
                                    level_points,
                                    values='总积分',
                                    names='会员等级',
                                    title='各等级积分贡献占比',
                                    hole=0.4,
                                    color_discrete_sequence=px.colors.qualitative.D3
                                )
                                fig_level_points.update_layout(
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.5,
                                    legend=dict(orientation="h", yanchor="bottom", y=-0.2)
                                )
                                st.plotly_chart(fig_level_points, use_container_width=True)
                            
                            # 会员等级详细统计
                            st.markdown("#### 📋 会员等级详细统计")
                            level_detail = df_filtered.groupby('等级').agg({
                                '会员卡号': 'nunique',
                                '积分数': ['sum', 'count', 'mean', 'max', 'min'],
                                '消费金额(元)': 'sum'
                            }).reset_index()
                            level_detail.columns = ['会员等级', '会员数', '总积分', '交易次数', '平均积分', '最大积分', '最小积分', '总消费金额']
                            # 保留整数并添加千分符
                            level_detail['会员数'] = level_detail['会员数'].astype(int).apply(lambda x: f'{x:,}')
                            level_detail['总积分'] = level_detail['总积分'].round(0).astype(int).apply(lambda x: f'{x:,}')
                            level_detail['交易次数'] = level_detail['交易次数'].astype(int).apply(lambda x: f'{x:,}')
                            level_detail['平均积分'] = level_detail['平均积分'].round(0).astype(int).apply(lambda x: f'{x:,}')
                            level_detail['最大积分'] = level_detail['最大积分'].round(0).astype(int).apply(lambda x: f'{x:,}')
                            level_detail['最小积分'] = level_detail['最小积分'].round(0).astype(int).apply(lambda x: f'{x:,}')
                            level_detail['总消费金额'] = level_detail['总消费金额'].round(0).astype(int).apply(lambda x: f'{x:,}')
                            level_detail['会员贡献度(金额/会员数)'] = (level_detail['总消费金额'].str.replace(',', '').astype(int) / level_detail['会员数'].str.replace(',', '').astype(int)).round(0).astype(int).apply(lambda x: f'{x:,}')
                            level_detail = level_detail.sort_values('总积分', key=lambda x: x.str.replace(',', '').astype(int), ascending=False)
                            st.dataframe(level_detail, use_container_width=True, hide_index=True)
                            
                        else:
                            st.info("当前筛选条件下没有会员等级数据")

                    # ========== 积分预警分析 ==========
                    with tab_chart4:
                        if not df_filtered.empty:
                            st.markdown("#### 🎯 预警参数设置")
                            col_warning1, col_warning2, col_warning3 = st.columns(3)
                            with col_warning1:
                                warning_multiplier = st.slider(
                                    "业态平均倍数阈值",
                                    min_value=3.0,
                                    max_value=10.0,
                                    value=5.0,
                                    step=0.5,
                                    help="超过业态平均积分的倍数将被标记为异常"
                                )
                            with col_warning2:
                                warning_top_n = st.slider(
                                    "显示异常记录数",
                                    min_value=10,
                                    max_value=100,
                                    value=30,
                                    step=10
                                )
                            with col_warning3:
                                # 默认选择'零售'业态
                                biz_options = df_filtered[df_filtered['业态'] != '--']['业态'].unique().tolist() if '业态' in df_filtered.columns else []
                                biz_default = ['零售'] if '零售' in biz_options else []
                                biz_warning_filter = st.multiselect(
                                    "🏢 筛选业态",
                                    options=biz_options,
                                    default=biz_default,
                                    key="biz_warning_filter",
                                    help="选择要预警的业态，为空则全部"
                                )
                            
                            st.markdown("---")
                            
                            # 1. 业态平均倍数预警
                            st.markdown("#### ⚠️ 业态平均倍数异常预警")
                            if '业态' in df_filtered.columns:
                                # 根据筛选条件过滤数据
                                df_biz_filtered = df_filtered.copy()
                                if biz_warning_filter:
                                    df_biz_filtered = df_biz_filtered[df_biz_filtered['业态'].isin(biz_warning_filter)]
                                
                                # 计算各业态平均积分
                                biz_avg = df_biz_filtered[df_biz_filtered['业态'] != '--'].groupby('业态')['积分数'].mean().reset_index()
                                biz_avg.columns = ['业态', '业态平均积分']
                                
                                # 合并业态平均积分到主数据
                                df_warning = df_biz_filtered.merge(biz_avg, on='业态', how='left')
                                df_warning['业态平均积分'] = df_warning['业态平均积分'].fillna(biz_avg['业态平均积分'].mean() if not biz_avg.empty else 0)
                                
                                # 计算倍数
                                df_warning['超出倍数'] = df_warning['积分数'] / df_warning['业态平均积分']
                                df_warning['超出倍数'] = df_warning['超出倍数'].round(2)
                                
                                # 筛选异常记录
                                abnormal_biz = df_warning[df_warning['超出倍数'] >= warning_multiplier].copy()
                                abnormal_biz = abnormal_biz.sort_values('超出倍数', ascending=False).head(warning_top_n)
                                
                                if not abnormal_biz.empty:
                                    abnormal_biz_display = abnormal_biz[['会员卡号', '手机号', '等级', '业态', '租户', '积分数', '业态平均积分', '超出倍数', '积分类型', '积分时间']].copy()
                                    abnormal_biz_display['积分时间'] = abnormal_biz_display['积分时间'].dt.strftime('%Y-%m-%d %H:%M')
                                    abnormal_biz_display['积分数'] = abnormal_biz_display['积分数'].round(0).astype(int)
                                    abnormal_biz_display['业态平均积分'] = abnormal_biz_display['业态平均积分'].round(2)
                                    
                                    st.warning(f"发现 {len(abnormal_biz)} 条异常记录（积分数超过业态平均 {warning_multiplier} 倍）")
                                    st.dataframe(abnormal_biz_display, use_container_width=True, hide_index=True)
                                else:
                                    st.success(f"✅ 未发现超过业态平均 {warning_multiplier} 倍的异常积分记录")
                            
                            st.markdown("---")
                            
                            # 2. 单个会员单次高额积分预警（受业态筛选控制）
                            st.markdown("#### 💰 单次高额积分预警")
                            # 使用筛选后的业态数据
                            high_threshold = df_biz_filtered['积分数'].quantile(0.95) if not df_biz_filtered.empty else 0
                            high_points = df_biz_filtered[df_biz_filtered['积分数'] >= high_threshold].copy()
                            high_points = high_points.sort_values('积分数', ascending=False).head(warning_top_n)  # 按积分数降序
                            
                            if not high_points.empty:
                                high_points_display = high_points[['会员卡号', '手机号', '等级', '业态', '租户', '积分数', '消费金额(元)', '积分类型', '积分时间']].copy()
                                high_points_display['积分时间'] = high_points_display['积分时间'].dt.strftime('%Y-%m-%d %H:%M')
                                high_points_display['积分数'] = high_points_display['积分数'].round(0).astype(int)
                                high_points_display['消费金额(元)'] = high_points_display['消费金额(元)'].round(2)
                                
                                st.warning(f"发现 {len(high_points)} 条高额积分记录（超过95%分位数 {int(high_threshold)} 分）")
                                st.dataframe(high_points_display, use_container_width=True, hide_index=True)
                            else:
                                st.success("✅ 未发现高额积分记录")
                            
                            st.markdown("---")
                            
                            # 3. 会员消费次数异常预警（当前期间内消费次数超过业态平均倍数）
                            # 倍数阈值直接使用业态平均倍数阈值
                            st.markdown(f"#### 📊 会员消费次数异常预警 (阈值: ≥{warning_multiplier}倍)")
                            # 统计每个会员的消费次数和积分数（按会员卡号和业态分组）
                            member_stats = df_biz_filtered.groupby(['会员卡号', '业态']).agg(
                                消费次数=('会员卡号', 'count'),
                                积分数=('积分数', 'sum')
                            ).reset_index()
                            # 生成备注（租户+次数，按次数降序）
                            tenant_count = df_biz_filtered.groupby(['会员卡号', '业态', '租户']).size().reset_index(name='次数')
                            tenant_count = tenant_count.sort_values('次数', ascending=False)
                            tenant_summary = tenant_count.groupby(['会员卡号', '业态']).apply(
                                lambda x: ','.join([f"{row['租户']}{row['次数']}次" for _, row in x.iterrows()])
                            ).reset_index(name='备注')
                            member_stats = member_stats.merge(tenant_summary, on=['会员卡号', '业态'], how='left')
                            # 统计每个业态的会员数和总消费次数
                            biz_member_stats = df_biz_filtered.groupby('业态').agg(
                                业态总消费次数=('会员卡号', 'count'),
                                业态会员数=('会员卡号', 'nunique')
                            ).reset_index()
                            # 计算业态单人平均消费次数
                            biz_member_stats['平均次数'] = biz_member_stats['业态总消费次数'] / biz_member_stats['业态会员数']
                            biz_member_stats['平均次数'] = biz_member_stats['平均次数'].round(2)
                            # 合并到会员消费次数数据
                            member_stats = member_stats.merge(biz_member_stats[['业态', '平均次数']], on='业态', how='left')
                            # 计算超出倍数
                            member_stats['超出倍数'] = member_stats['消费次数'] / member_stats['平均次数']
                            member_stats['超出倍数'] = member_stats['超出倍数'].round(2)
                            # 筛选超出业态平均倍数阈值的异常记录
                            abnormal_counts = member_stats[member_stats['超出倍数'] >= warning_multiplier].copy()
                            abnormal_counts = abnormal_counts.sort_values('超出倍数', ascending=False).head(warning_top_n)
                            
                            if not abnormal_counts.empty:
                                # 获取会员的等级和手机号信息
                                member_info = df_biz_filtered[['会员卡号', '等级', '手机号']].drop_duplicates('会员卡号')
                                abnormal_counts_display = abnormal_counts.merge(member_info, on='会员卡号', how='left')
                                abnormal_counts_display = abnormal_counts_display[['会员卡号', '手机号', '等级', '业态', '消费次数', '平均次数', '超出倍数', '积分数', '备注']].copy()
                                abnormal_counts_display['积分数'] = abnormal_counts_display['积分数'].round(0).astype(int)
                                
                                st.warning(f"发现 {len(abnormal_counts)} 条异常记录（消费次数超出业态平均 {warning_multiplier} 倍）")
                                st.dataframe(abnormal_counts_display, use_container_width=True, hide_index=True)
                            else:
                                st.success("✅ 未发现消费次数异常的会员记录")
                            
                            st.markdown("---")
                            
                            # 4. 同一会员短时间多次积分预警（使用消费时间字段）
                            st.markdown("#### ⚡ 高频积分预警")
                            # 使用筛选后的业态数据
                            df_freq_filtered = df_biz_filtered.copy()
                            
                            # 确定时间字段并确保其为datetime类型
                            if '消费时间' in df_freq_filtered.columns:
                                time_field = '消费时间'
                            else:
                                time_field = '积分时间'
                            
                            # 确保时间字段为datetime类型
                            if df_freq_filtered[time_field].dtype != 'datetime64[ns]':
                                df_freq_filtered[time_field] = pd.to_datetime(df_freq_filtered[time_field], errors='coerce')
                            
                            # 按会员卡号和时间字段排序
                            df_freq_filtered = df_freq_filtered.sort_values(['会员卡号', time_field])
                            
                            # 计算相邻积分的时间间隔（分钟）
                            df_freq_filtered['时间间隔(分钟)'] = df_freq_filtered.groupby('会员卡号')[time_field].diff().dt.total_seconds() / 60
                            df_freq_filtered['时间间隔(分钟)'] = df_freq_filtered['时间间隔(分钟)'].round(2)
                            
                            # 获取上笔消费信息
                            df_freq_filtered['上笔消费租户'] = df_freq_filtered.groupby('会员卡号')['租户'].shift(1)
                            # 获取上笔积分数
                            df_freq_filtered['上笔积分数'] = df_freq_filtered.groupby('会员卡号')['积分数'].shift(1)
                            
                            # 筛选短时间内多次积分的记录（间隔小于5分钟且积分数大于100）
                            frequent_points = df_freq_filtered[(df_freq_filtered['时间间隔(分钟)'] < 5) & (df_freq_filtered['积分数'] > 100)].copy()
                            frequent_points = frequent_points.sort_values('时间间隔(分钟)').head(warning_top_n)
                            
                            if not frequent_points.empty:
                                frequent_points_display = frequent_points[['会员卡号', '等级', '业态', '租户', '积分数', '时间间隔(分钟)', '积分类型', '上笔消费租户', '上笔积分数', time_field]].copy()
                                frequent_points_display[time_field] = frequent_points_display[time_field].dt.strftime('%Y-%m-%d %H:%M')
                                frequent_points_display['积分数'] = frequent_points_display['积分数'].round(0).astype(int)
                                if '上笔积分数' in frequent_points_display.columns:
                                    frequent_points_display['上笔积分数'] = frequent_points_display['上笔积分数'].round(0).astype(int)
                                
                                st.warning(f"发现 {len(frequent_points)} 条高频积分记录（间隔小于5分钟且积分数大于100分）")
                                st.dataframe(frequent_points_display, use_container_width=True, hide_index=True)
                            else:
                                st.success("✅ 未发现高频积分记录")
                            
                            st.markdown("---")
                        else:
                            st.info("当前筛选条件下没有数据")

                    st.markdown("---")

                    with tab_chart5:
                        if df_consume.empty:
                            st.info("当前数据中没有积分消耗记录（积分方式 != '积分消耗'）")
                        else:
                            # 应用筛选条件到消耗数据
                            df_consume_filtered = df_consume.copy()
                            
                            # 日期筛选
                            if len(date_range) == 2:
                                mask_date_consume = (df_consume_filtered['积分时间'].dt.date >= start_date) & (df_consume_filtered['积分时间'].dt.date <= end_date)
                                df_consume_filtered = df_consume_filtered[mask_date_consume]
                            
                            # 业态筛选
                            if selected_businesses:
                                df_consume_filtered = df_consume_filtered[df_consume_filtered['业态'].isin(selected_businesses)]
                            
                            # 会员等级筛选
                            if selected_level != '全部':
                                df_consume_filtered = df_consume_filtered[df_consume_filtered['等级'] == selected_level]
                            
                            if df_consume_filtered.empty:
                                st.info("当前筛选条件下没有积分消耗记录")
                            else:
                                df_consume_copy = df_consume_filtered.copy()
                                df_consume_copy['消耗积分'] = df_consume_copy['积分数'].abs()
                                df_consume_copy['日期'] = df_consume_copy['积分时间'].dt.date
                                
                                # 获取与消耗对比（放到每日消耗趋势上方）
                                st.markdown("#### 📊 获取与消耗对比")
                                # 获取与消耗对比只受日期筛选影响，不受理态和会员等级筛选影响
                                df_gain_filtered = df_gain.copy()
                                if len(date_range) == 2:
                                    mask_date_gain = (df_gain_filtered['积分时间'].dt.date >= start_date) & (df_gain_filtered['积分时间'].dt.date <= end_date)
                                    df_gain_filtered = df_gain_filtered[mask_date_gain]
                                
                                df_consume_summary = df_consume.copy()
                                if len(date_range) == 2:
                                    mask_date_consume = (df_consume_summary['积分时间'].dt.date >= start_date) & (df_consume_summary['积分时间'].dt.date <= end_date)
                                    df_consume_summary = df_consume_summary[mask_date_consume]
                                
                                total_gain = df_gain_filtered['积分数'].sum()
                                total_consume = df_consume_summary['积分数'].abs().sum()
                                
                                col_ratio1, col_ratio2, col_ratio3 = st.columns(3)
                                
                                with col_ratio1:
                                    st.metric(
                                        label="💰 总获取积分",
                                        value=f"{total_gain:,.0f}",
                                        delta=f"{total_gain - total_consume:,.0f} 净增"
                                    )
                                
                                with col_ratio2:
                                    st.metric(
                                        label="💸 总消耗积分",
                                        value=f"{total_consume:,.0f}",
                                        delta=f"{total_consume / total_gain * 100:.1f}%" if total_gain > 0 else "0%",
                                        delta_color="inverse"
                                    )
                                
                                with col_ratio3:
                                    ratio = (total_consume / total_gain * 100) if total_gain > 0 else 0
                                    st.metric(
                                        label="📈 消耗率",
                                        value=f"{ratio:.1f}%",
                                        delta=f"{'健康' if ratio < 80 else '偏高'}"
                                    )
                                
                                # 每日消耗趋势 - 占满整行
                                st.markdown("#### 📅 每日消耗趋势")
                                daily_consume = df_consume_copy.groupby('日期')['消耗积分'].sum().reset_index()
                                daily_consume = daily_consume.sort_values('日期')
                                
                                fig_consume_line = px.line(
                                    daily_consume,
                                    x='日期',
                                    y='消耗积分',
                                    title='每日积分消耗趋势',
                                    labels={'消耗积分': '消耗积分数量', '日期': '日期'},
                                    markers=True
                                )
                                fig_consume_line.update_layout(
                                    plot_bgcolor='white',
                                    paper_bgcolor='white',
                                    title_font_size=14,
                                    title_x=0.5,
                                    xaxis=dict(tickformat='%Y-%m-%d')
                                )
                                fig_consume_line.update_traces(line=dict(color='#e53e3e', width=2), marker=dict(size=6, color='#c53030'))
                                st.plotly_chart(fig_consume_line, use_container_width=True)
                                
                                # 第二行：消耗类型分布（左）和消耗时段分布（右）
                                col_consume_left, col_consume_right = st.columns(2)
                                
                                with col_consume_left:
                                    if '积分类型' in df_consume_copy.columns:
                                        st.markdown("#### 🥧 消耗类型分布")
                                        type_consume = df_consume_copy.groupby('积分类型')['消耗积分'].sum().reset_index()
                                        type_consume = type_consume.sort_values('消耗积分', ascending=False)
                                        
                                        fig_consume_pie = px.pie(
                                            type_consume,
                                            values='消耗积分',
                                            names='积分类型',
                                            title='积分消耗类型占比',
                                            hole=0.4,
                                            color_discrete_sequence=px.colors.qualitative.Set2
                                        )
                                        fig_consume_pie.update_layout(
                                            paper_bgcolor='white',
                                            title_font_size=14,
                                            title_x=0.5
                                        )
                                        st.plotly_chart(fig_consume_pie, use_container_width=True)
                                
                                with col_consume_right:
                                    st.markdown("#### ⏰ 消耗时段分布")
                                    df_consume_copy['小时'] = df_consume_copy['积分时间'].dt.hour
                                    hourly_consume = df_consume_copy.groupby('小时')['消耗积分'].sum().reset_index()
                                    
                                    fig_hourly_consume = px.bar(
                                        hourly_consume,
                                        x='小时',
                                        y='消耗积分',
                                        title='积分消耗时段分布',
                                        labels={'消耗积分': '消耗积分数量', '小时': '小时'},
                                        color='消耗积分',
                                        color_continuous_scale='reds'
                                    )
                                    fig_hourly_consume.update_layout(
                                        plot_bgcolor='white',
                                        paper_bgcolor='white',
                                        title_font_size=14,
                                        title_x=0.5,
                                        xaxis=dict(tickmode='array', tickvals=list(range(0, 24)))
                                    )
                                    st.plotly_chart(fig_hourly_consume, use_container_width=True)
                                
                                st.markdown("#### 👑 积分消耗TOP会员")
                                # TOP会员数据只受日期筛选影响
                                # 筛选获取数据
                                df_gain_top = df_gain.copy()
                                if len(date_range) == 2:
                                    mask_date = (df_gain_top['积分时间'].dt.date >= start_date) & (df_gain_top['积分时间'].dt.date <= end_date)
                                    df_gain_top = df_gain_top[mask_date]
                                
                                # 筛选消耗数据
                                df_consume_top = df_consume.copy()
                                if len(date_range) == 2:
                                    mask_date = (df_consume_top['积分时间'].dt.date >= start_date) & (df_consume_top['积分时间'].dt.date <= end_date)
                                    df_consume_top = df_consume_top[mask_date]
                                
                                # 统计获取积分
                                gain_stats = df_gain_top.groupby('会员卡号').agg(
                                    获取积分数=('积分数', 'sum'),
                                    获取积分次数=('积分数', 'count'),
                                    消费金额=('消费金额(元)', 'sum')
                                ).reset_index()
                                
                                # 统计消耗积分
                                consume_stats = df_consume_top.groupby('会员卡号').agg(
                                    消耗积分数=('积分数', lambda x: x.abs().sum()),
                                    消耗积分次数=('积分数', 'count')
                                ).reset_index()
                                
                                # 获取手机号信息
                                phone_info = df_points.drop_duplicates('会员卡号')[['会员卡号', '手机号']]
                                
                                # 合并数据
                                member_stats = gain_stats.merge(consume_stats, on='会员卡号', how='outer').fillna(0)
                                member_stats = member_stats.merge(phone_info, on='会员卡号', how='left')
                                
                                # 按消耗积分数排序取TOP10
                                member_stats = member_stats.sort_values('消耗积分数', ascending=False).head(10)
                                
                                # 格式化数字
                                member_stats['获取积分数'] = member_stats['获取积分数'].round(0).astype(int)
                                member_stats['消耗积分数'] = member_stats['消耗积分数'].round(0).astype(int)
                                member_stats['消费金额'] = member_stats['消费金额'].round(2)
                                
                                if not member_stats.empty:
                                    st.dataframe(
                                        member_stats[['会员卡号', '手机号', '获取积分数', '消费金额', '获取积分次数', '消耗积分数', '消耗积分次数']],
                                        use_container_width=True,
                                        hide_index=True
                                    )
                                else:
                                    st.info("当前日期范围内没有积分消耗记录")
                                
                    # ========== 详细数据查询 ==========
                    with tab_chart6:
                        col_detail1, col_detail2, col_detail3, col_detail4, col_detail5, col_detail6, col_detail7, col_detail8 = st.columns([2, 2, 2, 2, 2, 2, 2, 1])
                        with col_detail1:
                            search_tenant = st.text_input("🔍 搜索租户", placeholder="输入租户名称", key="search_tenant")
                        with col_detail2:
                            search_member = st.text_input("🆔 会员卡号", placeholder="输入会员卡号", key="search_member")
                        with col_detail3:
                            point_type_filter = st.multiselect(
                                "🎯 积分类型",
                                options=sorted(df_points['积分类型'].dropna().unique().tolist()) if '积分类型' in df_points.columns else [],
                                default=[],
                                key="point_type_filter"
                            )
                        with col_detail4:
                            business_filter = st.multiselect(
                                "🏢 业态",
                                options=sorted(df_points['业态'].dropna().unique().tolist()) if '业态' in df_points.columns else [],
                                default=[],
                                key="business_filter"
                            )
                        with col_detail5:
                            level_filter = st.multiselect(
                                "👥 会员等级",
                                options=sorted(df_points['等级'].dropna().unique().tolist()) if '等级' in df_points.columns else [],
                                default=[],
                                key="level_filter"
                            )
                        with col_detail6:
                            method_filter = st.multiselect(
                                "🔧 积分方式",
                                options=sorted(df_points['积分方式'].dropna().unique().tolist()) if '积分方式' in df_points.columns else [],
                                default=[],
                                key="method_filter"
                            )
                        with col_detail7:
                            search_phone = st.text_input("📱 手机号", placeholder="输入手机号", key="search_phone")
                        with col_detail8:
                            export_format = st.selectbox("📥 导出格式", ["Excel", "CSV"], key="export_format")

                        # 详细数据查询基于原始数据，只受日期范围筛选影响
                        df_detail = df_points.copy()
                        
                        # 应用日期范围筛选
                        if len(date_range) == 2:
                            start_date, end_date = date_range
                            mask_date = (df_detail['积分时间'].dt.date >= start_date) & (df_detail['积分时间'].dt.date <= end_date)
                            df_detail = df_detail[mask_date].copy()
                        
                        # 应用详细数据查询内的筛选条件
                        if search_tenant:
                            df_detail = df_detail[df_detail['租户'].str.contains(search_tenant, na=False)]
                        if search_member:
                            df_detail = df_detail[df_detail['会员卡号'].astype(str).str.contains(search_member, na=False)]
                        if point_type_filter:
                            df_detail = df_detail[df_detail['积分类型'].isin(point_type_filter)]
                        if business_filter:
                            df_detail = df_detail[df_detail['业态'].isin(business_filter)]
                        if level_filter:
                            df_detail = df_detail[df_detail['等级'].isin(level_filter)]
                        if method_filter:
                            df_detail = df_detail[df_detail['积分方式'].isin(method_filter)]
                        if search_phone:
                            if '手机号' in df_detail.columns:
                                df_detail = df_detail[df_detail['手机号'].astype(str).str.contains(search_phone, na=False)]
                            elif '电话' in df_detail.columns:
                                df_detail = df_detail[df_detail['电话'].astype(str).str.contains(search_phone, na=False)]

                        df_detail_display = df_detail.copy()
                        df_detail_display['积分时间'] = df_detail_display['积分时间'].dt.strftime('%Y-%m-%d %H:%M')
                        
                        # 处理消费时间字段，将Excel日期序列号转换为统一格式
                        if '消费时间' in df_detail_display.columns:
                            def convert_datetime(x):
                                if pd.isna(x) or x == '' or x is None:
                                    return ''
                                # 数字类型按Excel日期序列号处理
                                if isinstance(x, (int, float)):
                                    # 检查是否在合理的Excel日期范围内（约1900-2100）
                                    if 1 <= x <= 73000:
                                        try:
                                            return pd.to_datetime(x, unit='D', origin='1899-12-30').strftime('%Y-%m-%d %H:%M')
                                        except:
                                            return str(x)
                                    else:
                                        return str(x)
                                # 字符串类型尝试解析
                                try:
                                    return pd.to_datetime(x).strftime('%Y-%m-%d %H:%M')
                                except:
                                    return str(x)
                            df_detail_display['消费时间'] = df_detail_display['消费时间'].apply(convert_datetime)
                        
                        df_display = df_detail_display.sort_values('积分时间', ascending=False).head(100)

                        st.dataframe(
                            df_display,
                            use_container_width=True,
                            hide_index=True
                        )
                        st.caption(f"显示最近 100 条记录（共 {len(df_detail)} 条匹配记录）")

                        col_export1, col_export2 = st.columns([1, 3])
                        with col_export1:
                            if not df_detail.empty:
                                if export_format == "Excel":
                                    from io import BytesIO
                                    output = BytesIO()
                                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                        df_detail.to_excel(writer, index=False, sheet_name='积分明细')
                                    output.seek(0)
                                    st.download_button(
                                        label="💾 下载数据",
                                        data=output,
                                        file_name=f"积分明细_{today}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )
                                else:
                                    csv_data = df_detail.to_csv(index=False).encode('utf-8-sig')
                                    st.download_button(
                                        label="💾 下载 CSV",
                                        data=csv_data,
                                        file_name=f"积分明细_{today}.csv",
                                        mime="text/csv",
                                        use_container_width=True
                                    )

                except Exception as e:
                    st.error(f"❌ 读取或处理积分数据时出错：{e}")

        with tab_other:
            st.info("🚧 其他数据看板正在建设中，敬请期待...")

    # ========== 数据更新 ==========
    elif current_page == "数据更新":
        import time
        st.markdown("### 🔄 数据更新")
        st.markdown("---")

        def get_formatted_file_mod_time(file_path):
            import time
            if os.path.exists(file_path):
                mod_time = os.path.getmtime(file_path)
                return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mod_time))
            return "文件不存在"

        st.markdown("""
        <style>
        .section-container {
            display: flex;
            align-items: stretch;
            margin-bottom: 0;
        }
        .section-title {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 1.5rem;
            border-radius: 10px 0 0 10px;
            min-width: 180px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .section-title-energy {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }
        .section-title-baigou {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            border-radius: 10px;
        }
        .section-title-report {
            background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
            border-radius: 10px;
        }
        .section-content {
            flex: 1;
            background: #f8f9fa;
            padding: 1rem 1.5rem;
            border-radius: 0 10px 10px 0;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .section-hr {
            border: none;
            border-top: 2px solid #e2e8f0;
            margin: 1.5rem 0;
        }
        </style>
        """, unsafe_allow_html=True)

        col_title1, col_content1 = st.columns([1, 4])
        with col_title1:
            st.markdown('<div class="section-title"><h4>🏢 金蝶数据更新</h4></div>', unsafe_allow_html=True)
        with col_content1:
            st.markdown("""
            <div style="
                background: #f8f9fa;
                padding: 2rem;
                border-radius: 0 10px 10px 0;
                border-left: 4px solid #667eea;
            ">
                <p style="margin: 0; color: #666; font-size: 1.1rem;">📂 请手动将金蝶数据文件更新至文件夹中</p>
                <p style="margin: 0.5rem 0 0 0; color: #999; font-size: 0.9rem;">提示：金蝶数据需要从金蝶ERP系统导出后手动放置</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown('<hr class="section-hr">', unsafe_allow_html=True)

        col_title2, col_content2 = st.columns([1, 4])
        with col_title2:
            st.markdown('<div class="section-title section-title-energy"><h4>⚡ 能耗系统更新</h4></div>', unsafe_allow_html=True)
        with col_content2:
            btn_cwj, btn_meter, btn_balance, btn_tenant = st.columns(4)
            with btn_cwj:
                st.markdown("##### 📋 财务订单")
                if st.button("📥 更新财务订单", use_container_width=True, key="btn_cwj_order"):
                    with st.spinner("正在下载财务订单，请稍候..."):
                        success, msg = run_download_script("能耗_财务订单.py")
                        if success:
                            st.success("财务订单下载成功")
                        else:
                            st.error(f"下载财务订单失败：{msg}")
                target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "财务订单.xls")
                if os.path.exists(target_file):
                    mtime = os.path.getmtime(target_file)
                    date_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
                    st.caption(f"文件更新时间：{date_str}")
                else:
                    st.caption("文件尚未下载")

            with btn_meter:
                st.markdown("##### 📊 能耗表底")
                if st.button("📥 更新能耗表底", use_container_width=True, key="btn_energy_meter"):
                    with st.spinner("正在下载表底数据，请稍候..."):
                        success, msg = run_download_script("能耗_表底.py")
                        if success:
                            st.success("表底数据下载成功")
                            downloaded_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗表底.xls")
                            try:
                                row_count = append_meter_readings_from_downloaded_file(excel_path, downloaded_file)
                                st.success(f"表底追加成功")
                            except Exception as e:
                                st.error(f"追加表底数据失败：{e}")
                        else:
                            st.error(f"下载表底失败：{msg}")
                target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗表底.xls")
                if os.path.exists(target_file):
                    mtime = os.path.getmtime(target_file)
                    date_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
                    st.caption(f"文件更新时间：{date_str}")
                else:
                    st.caption("文件尚未下载")

            with btn_balance:
                st.markdown("##### 💰 能耗余额")
                if st.button("📥 更新能耗余额", use_container_width=True, key="btn_energy_balance"):
                    with st.spinner("正在下载余额数据，请稍候..."):
                        success, msg = run_download_script("能耗_余额.py")
                        if success:
                            st.success("余额数据下载成功")
                        else:
                            st.error(f"下载余额失败：{msg}")
                target_file = os.path.join(BASE_DIR, "2-能耗系统数据", "能耗余额.xls")
                if os.path.exists(target_file):
                    mtime = os.path.getmtime(target_file)
                    date_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
                    st.caption(f"文件更新时间：{date_str}")
                else:
                    st.caption("文件尚未下载")

            with btn_tenant:
                st.markdown("##### 🏢 租户概览")
                if st.button("🔄 更新租户概览", use_container_width=True, key="btn_tenant_overview"):
                    with st.spinner("正在更新数据，请稍候..."):
                        success, msg = download_tenant_overview(tenant_overview_path)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
                st.caption(f"最后更新: {get_formatted_file_mod_time(tenant_overview_path)}")

        st.markdown('<hr class="section-hr">', unsafe_allow_html=True)

        col_title3, col_content3 = st.columns([1, 4])
        with col_title3:
            st.markdown('<div class="section-title section-title-baigou"><h4>🛒 百购系统更新</h4></div>', unsafe_allow_html=True)
        with col_content3:
            btn_zhengpu, btn_duojing = st.columns(2)
            with btn_zhengpu:
                st.markdown("##### 📄 正铺合同台账")
                if st.button("📥 更新正铺合同", use_container_width=True, key="btn_zhengpu_contract"):
                    with st.spinner("正在更新正铺合同..."):
                        try:
                            script_path = os.path.join(os.path.dirname(__file__), "download_zhengpu_contract.py")
                            result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, timeout=120)
                            if result.returncode == 0:
                                st.success("✅ 正铺合同更新成功！")
                            else:
                                st.error(f"❌ 更新失败: {result.stderr}")
                        except Exception as e:
                            st.error(f"❌ 更新出错: {e}")
                st.caption(f"最后更新: {get_formatted_file_mod_time(zhengpu_contract_path)}")

            with btn_duojing:
                st.markdown("##### 📄 多经合同台账")
                if st.button("📥 更新多经合同", use_container_width=True, key="btn_duojing_contract"):
                    with st.spinner("正在更新多经合同..."):
                        try:
                            script_path = os.path.join(os.path.dirname(__file__), "download_duojing_contract.py")
                            result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, timeout=120)
                            if result.returncode == 0:
                                st.success("✅ 多经合同更新成功！")
                            else:
                                st.error(f"❌ 更新失败: {result.stderr}")
                        except Exception as e:
                            st.error(f"❌ 更新出错: {e}")
                st.caption(f"最后更新: {get_formatted_file_mod_time(duojing_contract_path)}")

        st.markdown('<hr class="section-hr">', unsafe_allow_html=True)

        col_title4, col_content4 = st.columns([1, 4])
        with col_title4:
            st.markdown('<div class="section-title section-title-report"><h4>📈 报表中心更新</h4></div>', unsafe_allow_html=True)
        with col_content4:
            btn_points, empty1, empty2, empty3 = st.columns(4)
            with btn_points:
                st.markdown("##### 🎫 积分列表")
                if st.button("📥 更新积分列表", use_container_width=True, key="btn_points_list"):
                    with st.spinner("正在更新积分列表..."):
                        success, msg = update_points_list(BASE_DIR)
                        if success:
                            st.success(f"✅ {msg}")
                            st.rerun()
                        else:
                            st.error(f"❌ {msg}")
                points_db_file = os.path.join(BASE_DIR, "4-报表中心", "积分列表.db")
                st.caption(f"最后更新: {get_formatted_file_mod_time(points_db_file)}")

    # ---------- 专业级页脚 ----------
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; padding: 1.5rem 0; color: #718096; font-size: 0.8rem;">
        <div style="margin-bottom: 0.5rem;">
            <strong style="color: #1a365d;">唐山中骏商业管理有限公司</strong> · 财务共享服务中心
        </div>
        <div style="opacity: 0.7;">
            财务ERP系统 v2.0  
        </div>
    </div>
    """, unsafe_allow_html=True)

if __name__ == '__main__':
    main()