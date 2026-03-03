import re
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="按签订单位-年度-客商统计", layout="wide")
st.title("📊 按签订单位分组：年度客商Top统计（树状明细 + 可导出Excel）")

# =========================
# Helpers
# =========================
def normalize_colname(x: str) -> str:
    return str(x).replace("\u3000", " ").replace("\t", " ").strip()

def extract_year_from_filename(filename: str) -> str | None:
    m = re.search(r"(20\d{2})", filename)
    return m.group(1) if m else None

def read_csv_safely(uploaded_file) -> pd.DataFrame:
    encodings = ["utf-8", "utf-8-sig", "gbk"]
    last_err = None
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc, engine="python")
        except Exception as e:
            last_err = e
    raise last_err

def to_number_series(s: pd.Series) -> pd.Series:
    return pd.to_numeric(
        s.astype(str)
         .str.replace(",", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("\t", "", regex=False),
        errors="coerce"
    )

def top_n_by_unit(unit_name: str) -> int:
    special = ["智能所", "研究中心"]
    return 5 if any(k in str(unit_name) for k in special) else 10

def safe_sheet_name(name: str) -> str:
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name))
    return name[:31] if len(name) > 31 else name

def choose_excel_engine() -> str | None:
    """
    自动选择可用的 ExcelWriter 引擎：优先 xlsxwriter，其次 openpyxl
    """
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa: F401
            return "openpyxl"
        except Exception:
            return None

def render_tree(df_uy: pd.DataFrame, metric_col: str, metric_label: str, top_n: int):
    """
    在某一个（签订单位 + 年份）范围内：
    - 先输出客商 TopN 汇总表（按 metric）
    - 再以树状折叠展示：客商 -> 签订日期 -> 合同名称 -> 合同编号(明细)
    """
    # TopN 汇总
    g = (df_uy.groupby("客商名称", as_index=False)[metric_col]
           .sum(min_count=1)
           .dropna(subset=[metric_col]))

    if g.empty:
        st.info(f"该范围内 {metric_label} 无可统计数据。")
        return pd.DataFrame(columns=["客商名称", f"{metric_label}(万元)"])

    g[f"{metric_label}(万元)"] = g[metric_col] / 10000
    g = g.sort_values(metric_col, ascending=False).head(top_n).reset_index(drop=True)

    show_df = g[["客商名称", f"{metric_label}(万元)"]].copy()
    show_df[f"{metric_label}(万元)"] = show_df[f"{metric_label}(万元)"].round(2)

    st.dataframe(show_df, use_container_width=True)

    # 明细树
    top_customers = g["客商名称"].tolist()
    for cust in top_customers:
        cust_total_w = float(g.loc[g["客商名称"] == cust, f"{metric_label}(万元)"].iloc[0])

        with st.expander(f"▶ {cust}  |  {metric_label}合计：{cust_total_w:.2f} 万元", expanded=False):
            d0 = df_uy[df_uy["客商名称"] == cust].copy()

            # 日期展示 + 排序
            if "签订日期" in d0.columns:
                d0["_dt"] = pd.to_datetime(d0["签订日期"], errors="coerce")
                d0 = d0.sort_values("_dt")
                d0["日期显示"] = d0["_dt"].dt.strftime("%Y/%m/%d").fillna("未知日期")
            else:
                d0["日期显示"] = "未知日期"

            # 日期层
            for date_key, d_date in d0.groupby("日期显示", dropna=False):
                date_sum_w = (d_date[metric_col].sum(skipna=True) / 10000) if metric_col in d_date else 0.0

                with st.expander(f"  📅 {date_key}  |  小计：{date_sum_w:.2f} 万元", expanded=False):
                    # 合同名称层
                    for cname, d_cname in d_date.groupby("合同名称", dropna=False):
                        cname_show = str(cname) if pd.notna(cname) and str(cname).strip() else "（空合同名称）"
                        cname_sum_w = d_cname[metric_col].sum(skipna=True) / 10000

                        with st.expander(f"    📄 {cname_show}  |  小计：{cname_sum_w:.2f} 万元", expanded=False):
                            cols_show = ["合同编号", "合同名称", "合同金额", "结算金额", "完成量金额"]
                            cols_show = [c for c in cols_show if c in d_cname.columns]

                            dd = d_cname[cols_show].copy()

                            # 转万元显示
                            for c in ["合同金额", "结算金额", "完成量金额"]:
                                if c in dd.columns:
                                    dd[c] = (dd[c].fillna(0) / 10000).round(2)

                            st.dataframe(dd, use_container_width=True)

    return show_df


# =========================
# Upload
# =========================
uploaded_files = st.file_uploader(
    "上传 2022–2025 年的收入 CSV 文件（UTF-8/GBK 均可；建议文件名含年份如：xxx_2023.csv）",
    type=["csv"],
    accept_multiple_files=True
)

if not uploaded_files:
    st.stop()

# =========================
# Load & unify
# =========================
all_rows = []
errors = []

for f in uploaded_files:
    fname = f.name
    year_from_name = extract_year_from_filename(fname)

    try:
        df = read_csv_safely(f)
        df = df.rename(columns=normalize_colname)

        required = ["签订单位", "客商名称", "合同编号", "合同名称", "合同金额", "结算金额", "完成量金额"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            errors.append(f"❌ {fname} 缺少列：{', '.join(missing)}")
            continue

        keep_cols = required + (["签订日期"] if "签订日期" in df.columns else [])
        df = df[keep_cols].copy()

        # 年份字段：优先文件名；否则用签订日期解析
        if year_from_name:
            df["年份"] = year_from_name
        else:
            if "签订日期" in df.columns:
                dt = pd.to_datetime(df["签订日期"], errors="coerce")
                df["年份"] = dt.dt.year.astype("Int64").astype(str)
                df.loc[df["年份"].isin(["<NA>", "NaT", "nan"]), "年份"] = "未知"
            else:
                df["年份"] = "未知"

        # 数值清洗
        for col in ["合同金额", "结算金额", "完成量金额"]:
            df[col] = to_number_series(df[col])

        # 关键维度清洗
        df["签订单位"] = df["签订单位"].astype(str).str.strip()
        df["客商名称"] = df["客商名称"].astype(str).str.strip()
        df["合同编号"] = df["合同编号"].astype(str).str.strip()
        df["合同名称"] = df["合同名称"].astype(str).str.strip()

        # 去掉完全无效的客商
        df = df[df["客商名称"].notna() & (df["客商名称"] != "")]
        all_rows.append(df)

    except Exception as e:
        errors.append(f"❌ {fname} 读取/处理失败：{e}")

if errors:
    for msg in errors:
        st.error(msg)

if not all_rows:
    st.warning("没有可用数据（所有文件都缺列或读取失败）。")
    st.stop()

data_all = pd.concat(all_rows, ignore_index=True)

# =========================
# UI Filters
# =========================
units = sorted(data_all["签订单位"].dropna().unique().tolist())
years = sorted(data_all["年份"].dropna().unique().tolist())

with st.sidebar:
    st.header("筛选 / Filters")
    selected_units = st.multiselect("签订单位（可多选）", units, default=units)
    selected_years = st.multiselect("年份（可多选）", years, default=years)

df0 = data_all[data_all["签订单位"].isin(selected_units) & data_all["年份"].isin(selected_years)].copy()

if df0.empty:
    st.info("筛选后无数据。")
    st.stop()

# =========================
# Core: Unit -> Year -> Tree Tables
# =========================
# 导出缓存：每个单位-年份-指标一张sheet（Top榜）
export_sheets = {}

for unit in sorted(df0["签订单位"].unique().tolist()):
    n_top = top_n_by_unit(unit)

    with st.expander(f"🏢 签订单位：{unit}（Top{n_top} | 智能所/研究中心=5，其它=10）", expanded=False):
        df_unit = df0[df0["签订单位"] == unit].copy()

        for year in sorted(df_unit["年份"].unique().tolist()):
            with st.expander(f"📅 年份：{year}", expanded=False):
                df_uy = df_unit[df_unit["年份"] == year].copy()

                st.markdown("### ① 按合同金额 Top客商（树状明细）")
                top_table_1 = render_tree(df_uy, metric_col="合同金额", metric_label="合同金额", top_n=n_top)
                export_sheets[f"{unit}_{year}_合同金额"] = top_table_1

                st.markdown("### ② 按结算金额 Top客商（树状明细）")
                top_table_2 = render_tree(df_uy, metric_col="结算金额", metric_label="结算金额", top_n=n_top)
                export_sheets[f"{unit}_{year}_结算金额"] = top_table_2

                st.markdown("### ③ 按完成量金额 Top客商（树状明细）")
                top_table_3 = render_tree(df_uy, metric_col="完成量金额", metric_label="完成量金额", top_n=n_top)
                export_sheets[f"{unit}_{year}_完成量金额"] = top_table_3

st.divider()

# =========================
# Export Excel (engine fallback)
# =========================
if export_sheets:
    engine = choose_excel_engine()
    if not engine:
        st.error("当前环境缺少 Excel 导出依赖：xlsxwriter 或 openpyxl。请安装其中任意一个。")
        st.stop()

    output = BytesIO()
    with pd.ExcelWriter(output, engine=engine) as writer:
        for k, v in export_sheets.items():
            if v is None or v.empty:
                continue
            sheet = safe_sheet_name(k)
            v.to_excel(writer, sheet_name=sheet, index=False)

    st.download_button(
        label=f"📥 下载统计结果Excel（按单位_年份_指标分Sheet | engine={engine}）",
        data=output.getvalue(),
        file_name="Unit_Year_Customer_Top_Stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )