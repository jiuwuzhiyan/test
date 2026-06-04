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

# ---------- 全局文件基础目录 ----------
_LOCAL_DIR = os.path.join(os.path.dirname(__file__), "财务共享资料")
if os.path.exists(_LOCAL_DIR):
    BASE_DIR = _LOCAL_DIR
else:
    BASE_DIR = _LOCAL_DIR

# ---------- 数据文件路径 ----------
excel_path = os.path.join(BASE_DIR, "9-数据库", "能耗_电费数据.xlsx")
tenant_overview_path = os.path.join(BASE_DIR, "2-能耗系统数据", "租户概览.xls")

# ---------- 工具函数 ----------
def download_tenant_overview(target_path):
    """通过独立脚本自动下载租户概览文件"""
    script_path = os.path.join(os.path.dirname(__file__), "download_overview.py")
    try:
        result = subprocess.run(
            [sys.executable, script_path, target_path],
            capture_output=True,
            text=True,
            timeout=60
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

def get_file_mod_time(file_path):
    """获取文件修改时间戳"""
    if os.path.exists(file_path):
        return os.path.getmtime(file_path)
    return 0

def apply_custom_css():
    st.markdown("""
    <style>
        :root {
            --primary-color: #1a365d;
            --primary-light: #2c5282;
            --accent-color: #3182ce;
            --success-color: #38a169;
            --warning-color: #d69e2e;
            --danger-color: #e53e3e;
            --bg-main: #f7fafc;
            --bg-card: #ffffff;
            --text-primary: #1a202c;
            --text-secondary: #4a5568;
            --text-light: #718096;
            --border-color: #e2e8f0;
        }

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

        .enterprise-header {
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            color: white;
            padding: 1rem 1.5rem;
            border-radius: 8px;
            margin-bottom: 1.5rem;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
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

        .stTabs [data-baseweb="tab-list"] {
            gap: 0.5rem;
            background: var(--bg-card);
            padding: 0.5rem;
            border-radius: 12px;
            box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
            border: 1px solid var(--border-color);
        }

        .stTabs [data-baseweb="tab"] {
            font-weight: 600;
            color: var(--text-secondary);
            border-radius: 8px;
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
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }

        .dataframe {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
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

        .dataframe tbody tr:nth-child(even) {
            background: rgba(49, 130, 206, 0.03);
        }

        .dataframe tbody tr:hover {
            background: rgba(49, 130, 206, 0.08);
        }

        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        [data-testid="stToolbar"] {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="能耗系统-电费（启用明细 & 余额预警）",
        page_icon="⚡",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    apply_custom_css()

    # ---------- 顶部标题栏 ----------
    st.markdown("""
    <div class="enterprise-header">
        <div>
            <h1>⚡ 能耗系统-电费</h1>
        </div>
        <div style="text-align: right; font-size: 0.85rem; opacity: 0.9;">
            <div>唐山中骏商业管理有限公司</div>
            <div style="font-size: 0.75rem; margin-top: 0.25rem; opacity: 0.8;">启用明细 & 余额预警</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_existing = load_excel(excel_path)

    tab2, tab_alert = st.tabs(
        ["⚡启用明细", "⚠️余额预警"],
        key="energy_tabs_simple"
    )

    # ========== 标签页2：启用明细 ==========
    with tab2:
        col_date, col_status, col_no, col_brand = st.columns([2, 2, 1.5, 1.5])
        with col_date:
            settlement_date = st.date_input("结账日期", value=date.today(), key="settlement_date_v2")
        with col_status:
            status_placeholder = st.empty()
        with col_no:
            filter_table_no = st.text_input("表号", placeholder="模糊查询", key="filter_table_no_v2")
        with col_brand:
            filter_brand = st.text_input("品牌", placeholder="模糊查询", key="filter_brand_v2")

        if not df_existing.empty:
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
                    key="status_filter_v2"
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
            st.info("暂无清表数据，请先在主系统中添加清表记录。")
            empty_cols = ["自建ID", "表号", "品牌", "客户编号", "合同号", "进场日期", "进场表底", "撤场日期", "撤场表底", "电表状态"]
            st.session_state.df_result_full = pd.DataFrame(columns=empty_cols)
            st.session_state.enabled_df = pd.DataFrame(columns=["表号", "客户编号", "合同号"])

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
                                                key="alert_amt_v2")
                with col_day:
                    day_limit = st.number_input("剩余天数 ≤", min_value=0, value=3, step=1, format="%d",
                                                key="alert_day_v2")
                with col_logic:
                    logic = st.selectbox("逻辑", ["且", "或"], key="alert_logic_v2")
                with col_update:
                    st.write("")
                    st.write("")
                    update_clicked = st.button("🔄 更新数据", use_container_width=True)
                with col_time:
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


if __name__ == "__main__":
    main()
