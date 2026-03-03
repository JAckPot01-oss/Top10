import re
import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="按签订单位-年度-客商统计", layout="wide")
st.title("📊 按签订单位分组：年度客商Top统计（含合同明细折叠）")

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
    # 去逗号、空格、Tab，强制转数值，异常->NaN
    return pd.to_numeric(
        s.astype(str)
         .str.replace(",", "", regex=False)
         .str.replace(" ", "", regex=False)
         .str.replace("\t", "", regex=False),
        errors="coerce"
    )

def top_n_by_unit(unit_name: str) -> int:
    # 智能所、研究中心 取前5，其它前10
    special = ["智能所", "研究中心"]
    return 5 if any(k in str(unit_name) for k in special) else 10

def safe_sheet_name(name: str) -> str:
    # Excel sheet最长31，且不能含某些特殊字符
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", str(name))
    return name[:31] if len(name) > 31 else name

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

        # 必要列
        required = ["签订单位", "客商名称", "合同编号", "合同名称", "合同金额", "结算金额", "完成量金额"]
        missing = [c for c in required if c not in df.columns]

        if missing:
            errors.append(f"❌ {fname} 缺少列：{', '.join(missing)}")
            continue

        df = df[required + (["签订日期"] if "签订日期" in df.columns else [])].copy()

        # 年份字段：优先用文件名年份；否则用签订日期解析
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
# UI Filters (optional)
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
# Core: group by Unit -> Year -> Customer
# =========================
metrics = [
    ("合同金额", "合同金额"),
    ("结算金额", "结算金额"),
    ("完成量金额", "完成量金额"),
]

# 导出缓存：每个单位-年份-指标都出一张sheet
export_sheets = {}

for unit in sorted(df0["签订单位"].unique().tolist()):
    n_top = top_n_by_unit(unit)

    with st.expander(f"🏢 签订单位：{unit}（Top{n_top} | 智能所/研究中心=5，其它=10）", expanded=False):
        df_unit = df0[df0["签订单位"] == unit].copy()

        for year in sorted(df_unit["年份"].unique().tolist()):
            with st.expander(f"📅 年份：{year}", expanded=False):
                df_uy = df_unit[df_unit["年份"] == year].copy()

                # 明细数据（用于合同展开）
                # 这里不去重：同一合同在数据里如果重复行，你可以按需去重
                detail_cols = ["客商名称", "合同编号", "合同名称", "合同金额", "结算金额", "完成量金额"]
                df_detail = df_uy[detail_cols].copy()

                # 对每个指标分别做Top表
                for metric_label, metric_col in metrics:
                    st.subheader(f"📌 按 {metric_label} 排名（Top{n_top} 客商）")

                    # 客商汇总
                    g = (
                        df_uy.groupby("客商名称", as_index=False)[metric_col]
                        .sum(min_count=1)
                        .rename(columns={metric_col: f"{metric_label}(元)"})
                    )
                    g = g.dropna(subset=[f"{metric_label}(元)"])
                    g[f"{metric_label}(万元)"] = g[f"{metric_label}(元)"] / 10000
                    g = g.sort_values(by=f"{metric_label}(元)", ascending=False).head(n_top).reset_index(drop=True)

                    # 展示Top表
                    show_df = g[["客商名称", f"{metric_label}(万元)"]].copy()
                    show_df[f"{metric_label}(万元)"] = show_df[f"{metric_label}(万元)"].round(2)
                    st.dataframe(show_df, use_container_width=True)

                    # 合同明细折叠：每个Top客商一个 expander
                    top_customers = g["客商名称"].tolist()
                    for cust in top_customers:
                        with st.expander(f"🔎 客商：{cust} | 查看合同明细（合同名称 + 合同编号）", expanded=False):
                            d = df_detail[df_detail["客商名称"] == cust].copy()

                            # 只展示关键信息 + 三指标（便于核对）
                            d["合同金额"] = d["合同金额"].fillna(0)
                            d["结算金额"] = d["结算金额"].fillna(0)
                            d["完成量金额"] = d["完成量金额"].fillna(0)

                            d_show = d[["合同编号", "合同名称", "合同金额", "结算金额", "完成量金额"]].copy()
                            # 转万元显示更友好
                            for c in ["合同金额", "结算金额", "完成量金额"]:
                                d_show[c] = (d_show[c] / 10000).round(2)

                            st.dataframe(d_show, use_container_width=True)

                    # ============ Export sheet ============
                    sheet_key = f"{unit}_{year}_{metric_label}"
                    export_sheets[sheet_key] = show_df.copy()

st.divider()

# =========================
# Export Excel (all sheets)
# =========================
if export_sheets:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for k, v in export_sheets.items():
            writer_sheet = safe_sheet_name(k)
            v.to_excel(writer, sheet_name=writer_sheet, index=False)

    st.download_button(
        label="📥 下载统计结果Excel（按单位_年份_指标分Sheet）",
        data=output.getvalue(),
        file_name="Unit_Year_Customer_Top_Stats.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )