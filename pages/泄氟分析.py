# ==== 本页隐藏侧栏（更强兼容版）+ 放大字体 ==== 
import streamlit as st
import streamlit.components.v1 as components

# 只保留一次 set_page_config，放最前（若上层已设置则忽略）
try:
    st.set_page_config(page_title="泄漏分析（1.1.3.1）", layout="wide", initial_sidebar_state="collapsed")
except Exception:
    pass

# —— CSS & JS（隐藏侧栏）
st.markdown("""
<style>
aside[aria-label="sidebar"],aside[aria-label="Sidebar"],aside[class*="sidebar"],
[data-testid="stSidebar"],[data-testid^="stSidebar"],[data-testid*="Sidebar"],
[data-testid="stSidebarNav"],[data-testid="stSidebarContent"],[data-testid="stSidebarCollapsedControl"],
[data-testid*="Collapse"],[data-testid*="collapsed"],nav[aria-label="Sidebar navigation"]{
  display:none!important;width:0!important;min-width:0!important;max-width:0!important;
  visibility:hidden!important;pointer-events:none!important;opacity:0!important;
}
div.block-container{padding-left:1rem!important;padding-right:1rem!important;margin-left:0!important;}
section.main,div.main,[data-testid="stAppViewContainer"]{margin-left:0!important;}
</style>
""", unsafe_allow_html=True)

components.html("""
<script>
(function(){
  function nuke(){
    const sels=[
      'aside[aria-label="sidebar"]','aside[aria-label="Sidebar"]','aside[class*="sidebar"]',
      '[data-testid="stSidebar"]','[data-testid^="stSidebar"]','[data-testid*="Sidebar"]',
      '[data-testid="stSidebarNav"]','[data-testid="stSidebarContent"]','[data-testid="stSidebarCollapsedControl"]',
      '[data-testid*="Collapse"]','[data-testid*="collapsed"]','nav[aria-label="Sidebar navigation"]'
    ];
    for(const s of sels){
      document.querySelectorAll(s).forEach(el=>{
        el.style.display='none';el.style.width='0';el.style.minWidth='0';el.style.maxWidth='0';
        el.style.visibility='hidden';el.style.pointerEvents='none';el.style.opacity='0';
      });
    }
    const bc=document.querySelector('div.block-container');
    if(bc){bc.style.paddingLeft='1rem';bc.style.paddingRight='1rem';}
  }
  nuke();
  new MutationObserver(nuke).observe(document.body,{subtree:true,childList:true,attributes:true});
  setInterval(nuke,500);
})();
</script>
""", height=0)

# ================= 基础依赖 =================
import pandas as pd
import numpy as np
import re, json
from pathlib import Path
from urllib.parse import unquote

# ================= 页面抬头 =================
try:
    st.set_page_config(page_title="泄漏配件分析（1.1.3.1）", layout="wide", initial_sidebar_state="expanded")
except Exception:
    pass
st.title("泄漏配件分析 1.1.3.1")

# ================= 持久化路径 =================
STORE_PARQUET     = Path("data_store.parquet")
STORE_CSV         = Path("data_store.csv")
MODEL_FLEET_JSON  = Path("model_fleet_counts.json")   # 型号系列在保量（可选）
SPEC_FLEET_JSON   = Path("spec_fleet_counts.json")    # ✅ 具体型号在保量（通用）
RAW_EXCEL_STORE   = Path("raw_excel_store.parquet")   # “冰箱数据”页的原始 Excel 库

def _load_raw_excel_store() -> pd.DataFrame:
    if RAW_EXCEL_STORE.exists():
        try:
            return pd.read_parquet(RAW_EXCEL_STORE)
        except Exception:
            try:
                return pd.read_csv(RAW_EXCEL_STORE.with_suffix(".csv"), dtype=str)
            except Exception:
                pass
    return pd.DataFrame()

# ============== 解析工具 ==============
PROD_DATE_CANDS = ["生产日期","生产日","productiondate","proddate","制造日期","出厂日期"]

# === 新增：统一格式化工具 ===
def normalize_prod_date_series_to_text(raw: pd.Series) -> pd.Series:
    s = raw.astype(str).str.strip()
    out = pd.Series("", index=s.index, dtype="object")

    m8 = s.str.fullmatch(r"\d{8}", na=False)
    if m8.any():
        dt8 = pd.to_datetime(s[m8], format="%Y%m%d", errors="coerce")
        out.loc[m8] = dt8.dt.strftime("%Y-%m-%d")

    m6 = s.str.fullmatch(r"\d{6}", na=False)
    if m6.any():
        s6 = s[m6]
        is_yyyymm = s6.str.match(r"(19|20)\d{2}(0[1-9]|1[0-2])$")
        idx_y4 = s6.index[is_yyyymm.fillna(False)]
        if len(idx_y4) > 0:
            dt_a = pd.to_datetime(s6.loc[idx_y4] + "01", format="%Y%m%d", errors="coerce")
            out.loc[idx_y4] = dt_a.dt.strftime("%Y-%m-%d")
        idx_y2 = s6.index[~is_yyyymm.fillna(False)]
        if len(idx_y2) > 0:
            y = "20" + s6.loc[idx_y2].str.slice(0, 2)
            m = s6.loc[idx_y2].str.slice(2, 4)
            d = s6.loc[idx_y2].str.slice(4, 6)
            dt_b = pd.to_datetime(y + m + d, format="%Y%m%d", errors="coerce")
            out.loc[idx_y2] = dt_b.dt.strftime("%Y-%m-%d")

    rest = (out == "")
    if rest.any():
        dt = pd.to_datetime(s[rest], errors="coerce", infer_datetime_format=True)
        out.loc[rest] = dt.dt.strftime("%Y-%m-%d")

    failed = out.isna() | (out == "NaT") | (out == "")
    if failed.any():
        out.loc[failed] = s.loc[failed]
    return out

def _parse_prod_year_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    y = pd.Series(np.nan, index=s.index, dtype="float")

    m8 = s.str.fullmatch(r"\d{8}", na=False)
    if m8.any():
        y.loc[m8] = s.loc[m8].str.slice(0, 4).astype(int)

    m6 = s.str.fullmatch(r"\d{6}", na=False)
    if m6.any():
        s6 = s.loc[m6]
        is_yyyymm = s6.str.match(r"(19|20)\d{2}(0[1-9]|1[0-2])$")
        if is_yyyymm.any():
            y.loc[s6.index[is_yyyymm]] = s6[is_yyyymm].str.slice(0, 4).astype(int)
        idx_y2 = s6.index[~is_yyyymm.fillna(False)]
        if len(idx_y2) > 0:
            y.loc[idx_y2] = 2000 + s6.loc[idx_y2].str.slice(0, 2).astype(int)

    rest_idx = y.index[y.isna()]
    if len(rest_idx) > 0:
        dt2 = pd.to_datetime(s.loc[rest_idx], errors="coerce", infer_datetime_format=True)
        ok_idx = dt2.index[dt2.notna()]
        if len(ok_idx) > 0:
            y.loc[ok_idx] = dt2.loc[ok_idx].dt.year.astype(int)
    return y

def _parse_prod_ym_series(s: pd.Series):
    s = s.astype(str).str.strip()
    year = pd.Series(np.nan, index=s.index, dtype="float")
    month = pd.Series(np.nan, index=s.index, dtype="float")

    m8 = s.str.fullmatch(r"\d{8}", na=False)
    if m8.any():
        year.loc[m8]  = s.loc[m8].str.slice(0, 4).astype(int)
        month.loc[m8] = s.loc[m8].str.slice(4, 6).astype(int)

    m6 = s.str.fullmatch(r"\d{6}", na=False)
    if m6.any():
        s6 = s.loc[m6]
        is_yyyymm = s6.str.match(r"(19|20)\d{2}(0[1-9]|1[0-2])$")
        if is_yyyymm.any():
            year.loc[s6.index[is_yyyymm]]  = s6[is_yyyymm].str.slice(0, 4).astype(int)
            month.loc[s6.index[is_yyyymm]] = s6[is_yyyymm].str.slice(4, 6).astype(int)
        idx_y2 = s6.index[~is_yyyymm.fillna(False)]
        if len(idx_y2) > 0:
            year.loc[idx_y2]  = 2000 + s6.loc[idx_y2].str.slice(0, 2).astype(int)
            month.loc[idx_y2] = s6.loc[idx_y2].str.slice(2, 4).astype(int)
    need_idx = year.index[year.isna() | month.isna()]
    if len(need_idx) > 0:
        dt2 = pd.to_datetime(s.loc[need_idx], errors="coerce", infer_datetime_format=True)
        ok_idx = dt2.index[dt2.notna()]
        if len(ok_idx) > 0:
            year.loc[ok_idx]  = dt2.loc[ok_idx].dt.year.astype(int)
            month.loc[ok_idx] = dt2.loc[ok_idx].dt.month.astype(int)
    return year, month

def parse_any_prod_date_series(raw: pd.Series) -> pd.Series:
    s = raw.astype(str).str.strip()
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")

    m8 = s.str.fullmatch(r"\d{8}", na=False)
    if m8.any():
        out.loc[m8] = pd.to_datetime(s[m8], format="%Y%m%d", errors="coerce")

    m6 = s.str.fullmatch(r"\d{6}", na=False)
    if m6.any():
        s6 = s[m6]
        is_yyyymm = s6.str.slice(0, 4).astype(int) >= 1900
        idx_a = s6.index[is_yyyymm]
        if len(idx_a) > 0:
            out.loc[idx_a] = pd.to_datetime(s6.loc[idx_a] + "01", format="%Y%m%d", errors="coerce")
        idx_b = s6.index[~is_yyyymm]
        if len(idx_b) > 0:
            y = "20" + s6.loc[idx_b].str.slice(0, 2)
            m = s6.loc[idx_b].str.slice(2, 4)
            d = s6.loc[idx_b].str.slice(4, 6)
            out.loc[idx_b] = pd.to_datetime(y + m + d, format="%Y%m%d", errors="coerce")

    rest = out.isna()
    if rest.any():
        out.loc[rest] = pd.to_datetime(s[rest], errors="coerce", infer_datetime_format=True)
    return out

# === 保留：具体型号在保量映射 ===
def normalize_model_key(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u3000"," ")
    s = re.sub(r"\s+"," ",s).strip()
    return s.upper()

def _load_spec_fleet_any():
    if not SPEC_FLEET_JSON.exists(): return [],{}
    try:
        data = json.loads(SPEC_FLEET_JSON.read_text(encoding="utf-8"))
    except Exception:
        return [],{}
    if isinstance(data, dict) and "triples" in data:
        triples = data["triples"]
    elif isinstance(data, dict):
        triples = [{"a":k,"b":"", "fleet":v} for k,v in data.items()]
    elif isinstance(data, list):
        triples = data
    else:
        triples = []
    mp={}
    for r in triples:
        a = normalize_model_key(r.get("a",""))
        b = normalize_model_key(r.get("b",""))
        f = int(pd.to_numeric(r.get("fleet",0), errors="coerce") or 0)
        if a: mp[a]=f
        if b: mp[b]=f
    return triples, mp

if "spec_fleet_triples" not in st.session_state or "spec_fleet_map" not in st.session_state:
    triples, mp = _load_spec_fleet_any()
    st.session_state.spec_fleet_triples = triples
    st.session_state.spec_fleet_map = mp

def lookup_spec_fleet(model_name: str) -> int:
    return int(st.session_state.spec_fleet_map.get(normalize_model_key(model_name), 0))

# ================= 工具函数（清洗/标准化） =================
def _normalize_df_date_column_inplace(df: pd.DataFrame, src_col: str, out_col: str = None):
    if df is None or df.empty or (src_col not in df.columns):
        return
    out_col = out_col or src_col
    df[out_col] = normalize_prod_date_series_to_text(df[src_col])

def ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Date","Category1","Category2","TAG1","TAG2","TAG3","TAG4",
            "CostUSD","PaidQty","IsWarranty","Completed","Year","Month","Quarter","_SrcYM"
        ])

    mapping={}
    prod_src_col=None

    for c in df.columns:
        lc=str(c).strip().lower()
        if (lc in PROD_DATE_CANDS) and (prod_src_col is None):
            prod_src_col=c
        if lc in ["日期","date","created_at","发生日期"]:
            mapping[c]="Date"
        elif lc in ["分类1","category1","大类"]:
            mapping[c]="Category1"
        elif lc in ["分类2","category2","小类"]:
            mapping[c]="Category2"
        elif lc in ["tag1","tag 1","tag-1"]:
            mapping[c]="TAG1"
        elif lc in ["tag2","tag 2","tag-2"]:
            mapping[c]="TAG2"
        elif lc in ["tag3","tag 3","tag-3"]:
            mapping[c]="TAG3"
        elif lc in ["tag4","tag 4","tag-4"]:
            mapping[c]="TAG4"
        elif lc in ["cost","costusd","amount","费用","金额","fee","cost(usd)","cost usd"]:
            mapping[c]="CostUSD"
        elif lc in ["paidqty","paid_qty","已付款数量"]:
            mapping[c]="PaidQty"
        elif lc in ["iswarranty","warranty","是否承保"]:
            mapping[c]="IsWarranty"
        elif lc in ["completed","是否完成","维修状态","状态"]:
            mapping[c]="Completed"

    if mapping:
        df=df.rename(columns=mapping)

    if prod_src_col:
        df["原始生产日期"] = df[prod_src_col]
        df["_NormProdDate"] = normalize_prod_date_series_to_text(df["原始生产日期"]) 

    for col,default in [
        ("Category1",""),("Category2",""),("TAG1",""),("TAG2",""),("TAG3",""),("TAG4",""),
        ("CostUSD",0.0),("PaidQty",0),("IsWarranty",False),("Completed",False)
    ]:
        if col not in df.columns:
            df[col]=default

    df["CostUSD"]=pd.to_numeric(df["CostUSD"], errors="coerce").fillna(0.0)
    df["PaidQty"]=pd.to_numeric(df["PaidQty"], errors="coerce").fillna(0).astype(int)

    if prod_src_col:
        if "Year" not in df.columns or df["Year"].isna().all():
            yy,_ = _parse_prod_ym_series(df["_NormProdDate"])
            df["Year"]=yy
        if "Month" not in df.columns or df["Month"].isna().all():
            _,mm = _parse_prod_ym_series(df["_NormProdDate"])
            df["Month"]=mm

    if "Year" in df.columns:  df["Year"]=pd.to_numeric(df["Year"], errors="coerce")
    if "Month" in df.columns: df["Month"]=pd.to_numeric(df["Month"], errors="coerce")

    return df

def _clean_text(s: str) -> str:
    s = ("" if pd.isna(s) else str(s)).lower().replace("\u3000"," ")
    s = re.sub(r"[\s\-_\/]+"," ", s)
    return s.strip()

def get_param(name: str, default=None, cast=int):
    try:
        v = st.query_params.get(name, default)
    except Exception:
        qp = st.experimental_get_query_params()
        v = (qp.get(name,[default]) or [default])[0]
    if v is None: return default
    if cast is None: return v
    try: return cast(v)
    except Exception: return default

def get_param_str(name: str, default=None):
    try:
        v = st.query_params.get(name, default)
    except Exception:
        qp = st.experimental_get_query_params()
        v = (qp.get(name,[default]) or [default])[0]
    if isinstance(v,str): return unquote(v)
    return default

def load_store_df_from_disk()->pd.DataFrame:
    if STORE_PARQUET.exists():
        try: return pd.read_parquet(STORE_PARQUET)
        except Exception: pass
    if STORE_CSV.exists():
        try: return pd.read_csv(STORE_CSV)
        except Exception: pass
    return pd.DataFrame()

def get_store_df()->pd.DataFrame:
    if "store_df" in st.session_state and isinstance(st.session_state["store_df"], pd.DataFrame):
        df = st.session_state["store_df"].copy()
    else:
        df = load_store_df_from_disk()
    df = ensure_cols(df)
    if "原始生产日期" in df.columns:
        df["原始生产日期"] = normalize_prod_date_series_to_text(df["原始生产日期"])
        df["_NormProdDate"] = df["原始生产日期"]
    return df

def money(v)->str:
    try: return f"${float(v):,.0f}"
    except Exception: return "$0"

# ================= 参数与数据读取 =================
code  = get_param_str("code", default="1.1.3.1")
year  = get_param("year", default=None, cast=int)
month = get_param("month", default=None, cast=int)
months_param = get_param_str("months", default=None)
q_param      = get_param_str("q", default=None)

df_all = get_store_df()
if df_all.empty:
    st.warning("未获取到任何数据：请先在“冰箱数据”页面上传数据；或确认 data_store.parquet / data_store.csv 是否存在。")
    st.stop()

yy = pd.to_numeric(df_all.get("Year"), errors="coerce")
mm = pd.to_numeric(df_all.get("Month"), errors="coerce")

if year is None:
    if yy.notna().any():
        year = int(yy.dropna().max())
    else:
        year = 2025

sel_months = None
if months_param:
    try:
        sel_months = [int(x) for x in re.split(r"[,\s]+", months_param) if x]
    except Exception:
        sel_months = None
if (sel_months is None) and (month is not None):
    sel_months = [int(month)]
if (sel_months is None) and q_param:
    _q_map = {"Q1":[1,2,3], "Q2":[4,5,6], "Q3":[7,8,9], "Q4":[10,11,12]}
    sel_months = _q_map.get(str(q_param).upper(), None)
if sel_months is None:
    mm_this_year = pd.to_numeric(df_all.loc[yy == year, "Month"], errors="coerce")
    if mm_this_year.notna().any():
        sel_months = [int(mm_this_year.dropna().max())]
    else:
        sel_months = [1]

month_for_fleet = int(sorted(sel_months)[-1])

mask_year   = (yy == int(year))
mask_months = mm.isin([int(m) for m in sel_months])
scope_df    = df_all[mask_year & mask_months].copy()

_set = set(sel_months)
if _set == {1,2,3}:
    scope_text = f"{year} Q1"
elif _set == {4,5,6}:
    scope_text = f"{year} Q2"
elif _set == {7,8,9}:
    scope_text = f"{year} Q3"
elif _set == {10,11,12}:
    scope_text = f"{year} Q4"
elif len(sel_months) == 1:
    scope_text = f"{year}-{int(sel_months[0]):02d}"
else:
    scope_text = f"{year} 月份：{','.join(str(m) for m in sel_months)}"

if "_NormProdDate" not in scope_df.columns:
    cand = next((c for c in [
        "原始生产日期","生产日期","生产日","productiondate","proddate","制造日期","出厂日期","Date"
    ] if c in scope_df.columns), None)
    if cand:
        scope_df["_NormProdDate"] = normalize_prod_date_series_to_text(scope_df[cand])

if "_NormProdDate" in scope_df.columns:
    py = _parse_prod_year_series(scope_df["_NormProdDate"])
    if py.notna().any():
        scope_df = scope_df.loc[py >= 2024].copy()
        st.caption("（统计：生产年 ≥ 2024）")
    else:
        st.warning("未能从【原始生产日期】解析出年份（可能日期格式异常或列为空），暂不进行 ≥2024 过滤。")
else:
    st.warning("未找到【原始生产日期/生产日期/Date】等可用于解析的日期列，暂不进行 ≥2024 过滤。")

if scope_df.empty:
    st.info(f"{scope_text} 没有可统计数据。")
    st.stop()

# === 只在“缺陷码列”里匹配 ===
def _detect_code_col(df: pd.DataFrame, code_root: str) -> str | None:
    if df is None or df.empty: return None
    pat = re.compile(rf'(?<!\d){re.escape(code_root)}(?:\.\d+)?(?!\d)', re.I)
    best_col, best_hit, best_ratio = None, 0, 0.0
    for c in df.columns:
        try: s = df[c].astype(str)
        except Exception: continue
        nn = (s != "").sum()
        if nn == 0: continue
        hit = s.str.contains(pat, na=False).sum()
        if hit == 0: continue
        ratio = hit / float(nn)
        if (hit > best_hit) or (hit == best_hit and ratio > best_ratio):
            best_col, best_hit, best_ratio = c, hit, ratio
    if best_col is None:
        candidates = ["缺陷码", "缺陷编码", "DefectCode", "FaultCode", "IssueCode", "问题代码", "问题编码"]
        for c in candidates:
            if c in df.columns:
                return c
    return best_col

code_root = (code or "1.1.3.1").strip()
code_col  = _detect_code_col(scope_df, code_root)
if code_col is None:
    st.warning("未找到明显的“缺陷码”列。请检查原始数据中是否有明确的缺陷码字段。")
    st.stop()

pat_code = re.compile(rf'(?<!\d){re.escape(code_root)}(?:\.\d+)?(?!\d)', re.I)
try:
    col_as_str = scope_df[code_col].astype(str)
except Exception:
    st.warning(f"列 {code_col} 无法转换为字符串用于匹配，请检查数据。")
    st.stop()

mask_code = col_as_str.str.contains(pat_code, na=False)
df_code   = scope_df[mask_code].copy()
st.caption(f"（缺陷码列自动识别为：{code_col}，命中 {int(mask_code.sum())} 条）")

# 关联原始库补入原始生产日期
raw_hist = _load_raw_excel_store()
if not raw_hist.empty:
    DEDUP_KEYS = [
        "工单号","索赔号","CaseNo","TicketNo",
        "S/N","Serial No","SN","Serial","SerialNo",
        "序列号","序列号SN"
    ]
    join_keys = [k for k in DEDUP_KEYS if (k in df_code.columns) and (k in raw_hist.columns)]
    RAW_PROD_CANDS = ["生产日期","生产日","ProductionDate","ProdDate","制造日期","出厂日期"]
    raw_prod_col = next((c for c in RAW_PROD_CANDS if c in raw_hist.columns), None)

    if join_keys and raw_prod_col:
        patch = raw_hist.loc[:, join_keys + [raw_prod_col]].drop_duplicates()
        patch = patch.rename(columns={raw_prod_col: "原始生产日期"})
        df_code = df_code.merge(patch, on=join_keys, how="left", suffixes=("", "_rawdup"))
        df_code["原始生产日期"] = normalize_prod_date_series_to_text(df_code["原始生产日期"])
        norm_text = normalize_prod_date_series_to_text(df_code["原始生产日期"])
        if "_NormProdDate" in df_code.columns:
            df_code["_NormProdDate"] = np.where(df_code["原始生产日期"].notna(), norm_text, df_code["_NormProdDate"])
        else:
            df_code["_NormProdDate"] = norm_text
        dt = parse_any_prod_date_series(df_code["原始生产日期"].astype(str))
        if "Year" not in df_code.columns:  df_code["Year"]  = np.nan
        if "Month" not in df_code.columns: df_code["Month"] = np.nan
        ok = dt.notna()
        df_code.loc[ok, "Year"]  = dt.loc[ok].dt.year
        df_code.loc[ok, "Month"] = dt.loc[ok].dt.month

st.caption(f"统计口径：{scope_text} | 缺陷码：{code}")
if df_code.empty:
    st.warning("该口径下未命中任何 1.1.3.1（泄漏）相关记录。")
    st.stop()

# === 以“原始明细”的合并逻辑构造统计底表 df_base（仅补原始生产日期） ===
df_base = df_code.copy()
raw_hist = _load_raw_excel_store()
if not raw_hist.empty:
    DEDUP_KEYS = ["工单号","索赔号","CaseNo","TicketNo","SN","Serial","SerialNo","S/N","Serial No","序列号","序列号SN"]
    join_keys = [k for k in DEDUP_KEYS if (k in df_base.columns) and (k in raw_hist.columns)]
    RAW_PROD_CANDS = ["生产日期","生产日","ProductionDate","ProdDate","制造日期","出厂日期"]
    def _pick_first(df, cands):
        for c in cands:
            if c in df.columns:
                return c
        return None
    raw_prod_col = _pick_first(raw_hist, RAW_PROD_CANDS)
    if join_keys and raw_prod_col:
        take_cols = [c for c in join_keys if c in raw_hist.columns]
        if raw_prod_col not in take_cols:
            take_cols.append(raw_prod_col)
        patch = raw_hist.loc[:, take_cols].drop_duplicates().copy()
        if raw_prod_col in join_keys:
            patch["原始生产日期"] = normalize_prod_date_series_to_text(patch[raw_prod_col])
        else:
            patch = patch.rename(columns={raw_prod_col: "原始生产日期"})
            patch["原始生产日期"] = normalize_prod_date_series_to_text(patch["原始生产日期"])
        df_base = df_base.merge(patch, on=join_keys, how="left", suffixes=("", "_rawdup"))

if "原始生产日期" in df_base.columns:
    df_base["_NormProdDate"] = normalize_prod_date_series_to_text(df_base["原始生产日期"]).where(
        df_base["原始生产日期"].notna(), df_base.get("_NormProdDate")
    )

if "_NormProdDate" not in df_code.columns and "原始生产日期" in df_code.columns:
    df_code["_NormProdDate"] = normalize_prod_date_series_to_text(df_code["原始生产日期"])    
prod_year = _parse_prod_year_series(df_code["_NormProdDate"]) if "_NormProdDate" in df_code.columns else pd.Series(np.nan, index=df_code.index)
df_code = df_code.loc[prod_year >= 2024].copy()
if df_code.empty:
    st.warning("按要求过滤后无数据；请检查生产日期列或更换统计口径。")
    st.stop()

# —— 可选去重：同一工单/索赔只保留一条
DEDUP_KEYS = ["工单号","索赔号","CaseNo","TicketNo","SN","Serial","SerialNo"]
dedup_key = next((c for c in DEDUP_KEYS if c in df_code.columns), None)
if dedup_key:
    before = len(df_code)
    df_code = df_code.drop_duplicates(subset=[dedup_key], keep="first")
    st.caption(f"（已按 {dedup_key} 去重：{before} → {len(df_code)}）")

# ================= 顶部抬头条 =================
def render_topbar(_code: str, _scope_text: str, df_scope: pd.DataFrame):
    paid = pd.to_numeric(df_scope.get("CostUSD", 0), errors="coerce").fillna(0)
    qty  = len(df_scope)
    cost = float(paid[paid>0].sum())
    avg  = float(paid[paid>0].mean()) if (paid>0).any() else 0.0
    col1,col2,col3,col4 = st.columns([1.2,1,1,1])
    with col1:
        st.markdown(f"""
        <div style="padding:12px 16px;border-radius:14px;background:#0f172a;color:#fff;">
          <div style="font-size:13px;opacity:.7;">缺陷码 / 统计口径</div>
          <div style="font-size:20px;font-weight:700;line-height:1.2;margin-top:2px;">{_code}</div>
          <div style="font-size:13px;margin-top:2px;">{_scope_text}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2: st.metric("记录数", f"{qty}")
    with col3: st.metric("已付款总费用", money(cost))
    with col4: st.metric("单均已付款费用", money(avg))

render_topbar(code, scope_text, df_code)
st.markdown("---")

# ================= 1.1.3.1.x（泄漏）分项 =================
def build_1131_subtable(df_in: pd.DataFrame) -> pd.DataFrame:
    TEXT_HINT_CANDS = [
        "Category1","Category2","TAG1","TAG2","TAG3","TAG4",
        "故障描述","问题现象","症状","维修内容","处理措施","备注","描述",
        "Notes","Note","Detail","Details","Description","Comments",
        "原始EXCEL表映射","原始EXCEL映射","原始Excel映射","原始映射","缺陷码(原始)","缺陷码原始",
        "Part","PartName","配件","配件名称","工单内容"
    ]
    text_cols = [c for c in [code_col] + TEXT_HINT_CANDS if c in df_in.columns]
    if not text_cols:
        text_cols = list(df_in.columns)

    text = (
        df_in[text_cols]
        .astype(str)
        .apply(lambda s: s.str.replace("\u3000", " ", regex=False))
        .apply(lambda s: s.str.replace(r"\s*[\.\·]\s*", ".", regex=True))
        .apply(lambda s: s.str.replace(r"[\s\-_\/]+", " ", regex=True))
    )
    text = text.apply(lambda row: " ".join(row.values), axis=1).str.lower().str.strip()

    def _pat_exact(code: str):
        return re.compile(rf"(?<!\d){re.escape(code)}(?!\d)", re.I)

    code_map = {
        "1.1.3.1.15": ("1.1.3.1.15 Evap Coil","蒸发器"),
        "1.1.3.1.16": ("1.1.3.1.16 Condenser Coil","冷凝器"),
        "1.1.3.1.13": ("1.1.3.1.13 L1 leaking points on Suction tube B","蒸发器出口处"),
        "1.1.3.1.17": ("1.1.3.1.17 Filter Drier","干燥过滤器"),
        "1.1.3.1.2":  ("1.1.3.1.2 Leaking freon in the wall","内部泄漏"),
        "1.1.3.1.3":  ("1.1.3.1.3 Found leaking points somewhere and fixed it","焊点泄漏"),
        "1.1.3.1.1":  ("1.1.3.1.1 Leak not found/Vacuum and Recharge","未找到漏点"),
        "1.1.3.1.14": ("1.1.3.1.14 L2/L3 leaking points on Suction tube A","回气管焊接处"),
        "1.1.3.1.12": ("1.1.3.1.12 H19 leaking points on T Defrost tube B","蒸发器化霜加热管焊点"),
        "1.1.3.1.5":  ("1.1.3.1.5 H4/H5 leaking points on Condenser","冷凝器焊点泄漏"),
        "1.1.3.1.7":  ("1.1.3.1.7 H8/H9/H10 leaking points on Capillary tube A","外部毛细管焊点泄漏"),
        "1.1.3.1.4":  ("1.1.3.1.4 H2/H3 leaking points on Discharge Tube Loop","散热管焊点泄漏"),
        "1.1.3.1.9":  ("1.1.3.1.9 H12/H13 leaking points on Solenoid Valve Normally Open","常通电池阀泄漏"),
        "1.1.3.1.11": ("1.1.3.1.11 H16/H17/H18 leaking points on Solenoid Valve Normally close","常闭电池阀泄漏"),
        "1.1.3.1.8":  ("1.1.3.1.8 H11 leaking points on Capillary tube B","蒸发器毛细管泄漏"),
        "1.1.3.1.6":  ("1.1.3.1.6 H6/H7 leaking points on Filter drye","干燥过滤器附近焊点"),
    }

    order_codes = sorted(code_map.keys(), key=len, reverse=True)
    unassigned = pd.Series(True, index=df_in.index)
    cost = pd.to_numeric(df_in.get("CostUSD", 0), errors="coerce").fillna(0.0)

    rows = []
    for code in order_codes:
        en, zh = code_map[code]
        m = unassigned & text.str.contains(_pat_exact(code), na=False)
        rows.append({
            "英文名称": en,
            "质量问题": zh,
            "数量": int(m.sum()),
            "费用": float(cost[m].sum())
        })
        unassigned &= (~m)

    df_out = pd.DataFrame(rows)
    total_row = {
        "英文名称":"总计",
        "质量问题":"",
        "数量": int(pd.to_numeric(df_out["数量"], errors="coerce").fillna(0).sum()),
        "费用": float(pd.to_numeric(df_out["费用"], errors="coerce").fillna(0).sum()),
    }
    df_out = pd.concat([df_out, pd.DataFrame([total_row])], ignore_index=True)

    st.caption("（分项匹配使用列：%s）" % ", ".join(text_cols))
    return df_out

def _keep_positive_rows(df: pd.DataFrame, keep_total=True) -> pd.DataFrame:
    if df is None or df.empty: return df
    out = df.copy()
    if "数量" in out.columns:
        out["数量"] = pd.to_numeric(out["数量"], errors="coerce").fillna(0).astype(int)
        if keep_total:
            mask = (out["数量"] > 0) | (out["英文名称"].astype(str) == "总计")
        else:
            mask = (out["数量"] > 0)
        out = out[mask]
    return out

df_1131 = build_1131_subtable(df_code)
df_1131 = _keep_positive_rows(df_1131, keep_total=True)
st.markdown("### 1.1.3.1.x 分项（泄漏）")
if not df_1131.empty:
    show = df_1131.copy(); show["费用"] = show["费用"].map(lambda v: f"${float(v):,.0f}")
    st.dataframe(show, use_container_width=True)
    st.download_button("下载 CSV（1.1.3.1 分项）",
        data=df_1131.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"leaking_1131_breakdown_{scope_text.replace(' ','_')}.csv",
        mime="text/csv", use_container_width=True)
else:
    st.info("暂无 1.1.3.1 相关记录。")

# ================= 统计表（按系列 / 按具体型号） =================
USE_MONTH_CAND = [
    "已使用保修时长","已使用保修月数","保修已使用时长","已使用月份","保修使用月数",
    "使用月数","使用月","ServiceMonths","months_in_use","use_months"
]
SERIES_PRI_CAND = ["型号归类","型号系列","series","model_series"]
SERIES_SEC_CAND = ["机型","产品型号","MODEL","Model"]
SPEC_MODEL_CAND = ["型号","具体型号","机型明细","MODEL","Model","Part Model","PartModel","Part"]

MODEL_CATEGORY_MAP = {
    "MBF":"冰箱","MCF":"陈列柜","MSF":"沙拉台","MGF":"工作台","MPF":"玻璃台","CHEF BASE":"抽屉工作台",
    "MBC":"滑盖冷藏柜","MBB":"卧式吧台","SBB":"移门吧台","MKC":"啤酒柜","MBGF":"滑盖冷冻柜",
    "AMC":"牛奶柜","RDCS":"方形展示柜","CRDC":"弧形展示柜","CRDS":"台面方形展示柜","ATHOM":"饮料柜",
    "AOM":"立式饮料柜","MSCT":"配料台",
}
MODEL_ORDER = ["AOM","CHEF BASE","MBC","MBF","MBGF","MCF","MGF","MPF","MSF"]
ALIASES = {"CHEFBASE":"CHEF BASE","CHEF-BASE":"CHEF BASE","CHEF_BASE":"CHEF BASE","CHEF BASE":"CHEF BASE"}

def resolve_model_key(raw: str) -> str:
    s=(raw or "").strip().replace("\u3000", " ")
    s_norm=re.sub(r"[\s\-_\/]+", " ", s).strip()
    key_up=s_norm.upper().replace(" ","")
    for k,v in ALIASES.items():
        if key_up==k.upper().replace(" ",""): return v
    toks=s_norm.split()
    if len(toks)>1: return " ".join(t.upper() for t in toks)
    return s_norm.upper()

def _pick_col(df: pd.DataFrame, cands: list):
    for c in cands:
        if c in df.columns:
            return c
    norm = {re.sub(r"[\s_]+", "", str(c).lower()): c for c in df.columns}
    for want in cands:
        k = re.sub(r"[\s_]+", "", str(want).lower())
        if k in norm:
            return norm[k]
    return None

def load_model_fleet():
    if MODEL_FLEET_JSON.exists():
        try: return json.loads(MODEL_FLEET_JSON.read_text(encoding="utf-8"))
        except Exception: pass
    return {}

def get_fleet_month(fleet: dict, year: int=None, month: int=None)->dict:
    if not fleet: return {k:0 for k in MODEL_CATEGORY_MAP.keys()}
    if year is None or month is None:
        try:
            yy=max(int(y) for y in fleet.keys())
            mm=max(int(m) for m in fleet[str(yy)].keys())
            year,month=yy,mm
        except Exception:
            return {k:0 for k in MODEL_CATEGORY_MAP.keys()}
    y,m=str(int(year)),str(int(month))
    return {k:int(fleet.get(y,{}).get(m,{}).get(k,0)) for k in MODEL_CATEGORY_MAP.keys()}

def _bucket_series(months: pd.Series):
    m=pd.to_numeric(months, errors="coerce").where(lambda x: x>0)
    return pd.cut(m, bins=[0,4,8,np.inf], labels=["短期（1-4月）","中期（5-8月）","长期（9-24月）"],
                  include_lowest=True, right=True)

def build_tables(df_subset: pd.DataFrame, *, year=None, month=None):
    months_col = _pick_col(df_subset, USE_MONTH_CAND)
    series_col = _pick_col(df_subset, SERIES_PRI_CAND) or _pick_col(df_subset, SERIES_SEC_CAND)
    spec_col   = _pick_col(df_subset, SPEC_MODEL_CAND)

    if months_col is None:
        return pd.DataFrame(), pd.DataFrame()

    if series_col is None and spec_col is not None:
        tmp = df_subset[spec_col].astype(str).str.upper().str.extract(r"^([A-Z]{2,4})")
        df_subset = df_subset.copy()
        df_subset["__SERIES__"] = tmp[0]
        series_col = "__SERIES__"

    m = pd.to_numeric(df_subset[months_col], errors="coerce")
    valid_mask = m.notna() & (m>0)
    df_valid = df_subset[valid_mask].copy()
    if df_valid.empty:
        return pd.DataFrame(), pd.DataFrame()
    df_valid["__bucket__"] = _bucket_series(df_valid[months_col])

    # 表1：按系列
    tbl1=pd.DataFrame()
    if series_col is not None:
        g = pd.pivot_table(df_valid, index=series_col, columns="__bucket__", values="_NormProdDate",
                           aggfunc="count", fill_value=0)
        for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）"]:
            if c not in g.columns: g[c]=0
        g = g[["短期（1-4月）","中期（5-8月）","长期（9-24月）"]]
        g["总计"]=g.sum(axis=1)
        g.reset_index(inplace=True)
        g.rename(columns={series_col:"型号"}, inplace=True)

        fleet=load_model_fleet()
        month_fleet=get_fleet_month(fleet, year, month)
        def _fleet_lookup(s):
            key=resolve_model_key(s)
            return int(month_fleet.get(key,0))
        g["在保量"]=g["型号"].map(_fleet_lookup).fillna(0).astype(int)
        g["问题比例(%)"]=np.where(g["在保量"]>0, g["总计"]/g["在保量"]*100.0, np.nan)

        order_map={k:i for i,k in enumerate(MODEL_ORDER)}
        g["__ord__"]=g["型号"].map(order_map).fillna(9999)
        g=g.sort_values(["__ord__","总计"], ascending=[True,False]).drop(columns="__ord__")
        tbl1=g

    # 表2：按具体型号
    tbl2=pd.DataFrame()
    if spec_col is not None:
        g2 = pd.pivot_table(df_valid, index=spec_col, columns="__bucket__", values="_NormProdDate",
                            aggfunc="count", fill_value=0)
        for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）"]:
            if c not in g2.columns: g2[c]=0
        g2 = g2[["短期（1-4月）","中期（5-8月）","长期（9-24月）"]]
        g2["总计"]=g2.sum(axis=1)
        g2.reset_index(inplace=True)
        g2.rename(columns={spec_col:"具体型号"}, inplace=True)

        g2["在保量"]=g2["具体型号"].map(lookup_spec_fleet).fillna(0).astype(int)
        g2["比例(%)"]=np.where(g2["在保量"]>0, g2["总计"]/g2["在保量"]*100.0, np.nan)
        g2=g2.sort_values("总计", ascending=False)
        tbl2=g2

    return tbl1, tbl2

# ========= 新增：构建“明细索引（三列）”的通用函数 =========
SERIAL_CANDS_MAIN = ["序列号","S/N","SN","Serial No","SerialNo","Serial","序列号SN"]
SPEC_MODEL_CAND_ALL = ["产品型号","具体型号","型号","机型明细","MODEL","Model","Part Model","PartModel","Part"]

def build_companion_index(df_src: pd.DataFrame) -> pd.DataFrame:
    if df_src is None or df_src.empty:
        return pd.DataFrame(columns=["产品型号","序列号","原始生产日期"])
    # 产品型号
    model_col = next((c for c in SPEC_MODEL_CAND_ALL if c in df_src.columns), None)
    if model_col is None:
        # 退化为系列或其它可用列
        model_col = next((c for c in ["具体型号","型号","型号系列","series","model_series"] if c in df_src.columns), None)
    # 序列号
    serial_col = next((c for c in SERIAL_CANDS_MAIN if c in df_src.columns), None)
    # 原始生产日期
    prod_col = "原始生产日期" if "原始生产日期" in df_src.columns else ("_NormProdDate" if "_NormProdDate" in df_src.columns else None)

    out = pd.DataFrame({
        "产品型号": (df_src[model_col].astype(str) if model_col else ""),
        "序列号":   (df_src[serial_col].astype(str) if serial_col else ""),
        "原始生产日期": (normalize_prod_date_series_to_text(df_src[prod_col]) if prod_col else "")
    })
    # 过滤空白 & 去重
    def _nonempty_row(r):
        return any([str(r["产品型号"]).strip(), str(r["序列号"]).strip(), str(r["原始生产日期"]).strip()])
    out = out[out.apply(_nonempty_row, axis=1)]
    out = out.drop_duplicates().reset_index(drop=True)
    return out

# ========= 新增：并排渲染（左原表 / 右明细索引） =========
def render_table_pair(title: str, df_stats: pd.DataFrame, order_cols, df_for_index: pd.DataFrame):
    col_left, col_right = st.columns([2, 1], vertical_alignment="top")
    with col_left:
        cols=[c for c in order_cols if c in df_stats.columns]
        view=df_stats.loc[:, cols].copy() if cols else df_stats.copy()
        for pct_col in ["问题比例(%)","比例(%)"]:
            if pct_col in view.columns:
                view[pct_col] = view[pct_col].map(lambda x: "" if pd.isna(x) else f"{float(x):.2f}%")
        st.subheader(title)
        st.dataframe(view, use_container_width=True)
    with col_right:
        st.subheader("明细索引")
        idx_df = build_companion_index(df_for_index)
        if idx_df.empty:
            st.info("无可展示的 产品型号/序列号/原始生产日期。")
        else:
            st.dataframe(idx_df[["产品型号","序列号","原始生产日期"]], use_container_width=True, height=min(400, 42 + 28*min(len(idx_df), 10)))

# —— 生成统计表并渲染并排视图
t1, t2 = build_tables(df_code, year=year, month=month_for_fleet)
base1 = pd.DataFrame(columns=["型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])
base2 = pd.DataFrame(columns=["具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","比例(%)"])

render_table_pair(
    "泄漏（按系列）",
    (t1 if not t1.empty else base1),
    order_cols=("型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"),
    df_for_index=df_code
)
render_table_pair(
    "泄漏（按具体型号）",
    (t2 if not t2.empty else base2),
    order_cols=("具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","比例(%)"),
    df_for_index=df_code
)

# ===== 分项深入分析（仅数量>0，默认只展开第一项）=====
st.markdown("## 分项深入分析（仅数量>0）")

def _clean_join(df_):
    j = df_.astype(str).agg(" ".join, axis=1).str.lower().str.replace("\u3000", " ", regex=False)
    return j.str.replace(r"[\s\-_\/]+", " ", regex=True).str.strip()

def _subset_1131_item(df_src: pd.DataFrame, en_name: str) -> pd.DataFrame:
    """
    从“英文名称”中精确抽取完整子码（如 1.1.3.1.15），
    再在原始文本里做精准匹配，避免把 15 误当成 1 的前缀。
    """
    if df_src is None or df_src.empty:
        return df_src.iloc[0:0].copy()

    # 1) 从标题里抽取完整 code：优先正则（行首），退化到第一个 token
    code_pat = re.compile(r"^1\.1\.3\.1\.(?:1[0-7]|[1-9])\b", re.I)
    m = code_pat.search(en_name or "")
    code_full = m.group(0) if m else (str(en_name).strip().split()[0] if en_name else "")

    if not code_full:
        return df_src.iloc[0:0].copy()

    # 2) 拼接原始文本做匹配（保持你原来的清洗）
    def _clean_join(df_):
        j = df_.astype(str).agg(" ".join, axis=1).str.lower().str.replace("\u3000", " ", regex=False)
        return j.str.replace(r"[\s\-_\/]+", " ", regex=True).str.strip()

    j = _clean_join(df_src)

    # 3) 用“精确 code”做边界匹配，确保不误伤 1.1.3.1.1 vs 1.1.3.1.15
    pat = re.compile(rf"(?<![0-9a-z]){re.escape(code_full)}(?![0-9a-z])", re.I)
    return df_src[j.str.contains(pat, na=False)].copy()


rows_need = (
    df_1131[
        df_1131["英文名称"].astype(str).str.match(r"^1\.1\.3\.1\.(?:[1-9]|1[0-7])\b", na=False) &
        (pd.to_numeric(df_1131["数量"], errors="coerce").fillna(0) > 0)
    ]
    .sort_values("数量", ascending=False)
    .copy()
)

if rows_need.empty:
    st.info("当前 1.1.3.1.x 没有数量>0 的条目可深入分析。")
else:
    for idx, r in rows_need.iterrows():
        en_name = str(r["英文名称"]); zh_name = str(r.get("质量问题",""))
        sub_df = _subset_1131_item(df_code, en_name)
        with st.expander(f"【{en_name}】{zh_name}｜记录 {int(r['数量'])}，费用 {money(float(r.get('费用',0)))}",
                         expanded=(idx==rows_need.index[0])):
            a1, a2 = build_tables(sub_df, year=year, month=month_for_fleet)
            _b1 = pd.DataFrame(columns=["型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])
            _b2 = pd.DataFrame(columns=["具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","比例(%)"])
            # —— 每个分项里的两张统计表也加“齐平的明细索引”
            render_table_pair("按系列", (a1 if not a1.empty else _b1),
                order_cols=("型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"),
                df_for_index=sub_df)
            render_table_pair("按具体型号", (a2 if not a2.empty else _b2),
                order_cols=("具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","比例(%)"),
                df_for_index=sub_df)

# ================= 原始明细（命中 1.1.3.1 的记录） =================
with st.expander("原始明细（命中 1.1.3.1 的记录）", expanded=False):
    base_cols = [
        "Category1","Category2","TAG1","TAG2","TAG3","TAG4",
        "CostUSD","PaidQty","IsWarranty","Completed"
    ]
    base_cols = [c for c in base_cols if c in df_base.columns]

    RAW_PROD_CANDS = ["生产日期","生产日","ProductionDate","ProdDate","制造日期","出厂日期"]
    def _pick_first(df, cands):
        for c in cands:
            if c in df.columns:
                return c
        return None

    df_view = df_base.copy()

    raw_hist = _load_raw_excel_store()
    if not raw_hist.empty:
        DEDUP_KEYS = ["工单号","索赔号","CaseNo","TicketNo","SN","Serial","SerialNo","S/N","Serial No","序列号","序列号SN"]
        join_keys = [k for k in DEDUP_KEYS if (k in df_view.columns) and (k in raw_hist.columns)]
        raw_prod_col = _pick_first(raw_hist, RAW_PROD_CANDS)
        if join_keys and raw_prod_col:
            take_cols = join_keys[:]
            if raw_prod_col not in take_cols:
                take_cols.append(raw_prod_col)
            patch = raw_hist.loc[:, take_cols].drop_duplicates()
            patch = patch.rename(columns={raw_prod_col: "原始生产日期"})
            df_view = df_view.merge(patch, on=join_keys, how="left", suffixes=("", "_rawdup"))
            st.caption("（已从📎 原始 Excel 表映射按 %s 关联补入：原始生产日期）" % ", ".join(join_keys))
    else:
        st.caption("（未找到📎 原始 Excel 表映射，显示现有列）")

    if "原始生产日期" in df_view.columns:
        df_view["原始生产日期"] = normalize_prod_date_series_to_text(df_view["原始生产日期"])

    serial_src = next((c for c in ["序列号","S/N","SN","Serial No","SerialNo","Serial","序列号SN"] if c in df_view.columns), None)
    df_view["序列号"] = df_view[serial_src].astype(str) if serial_src else ""

    show_cols = []
    if "原始生产日期" in df_view.columns: show_cols.append("原始生产日期")
    show_cols.append("序列号")
    show_cols.extend([c for c in base_cols if c not in show_cols])

    st.caption("提示：已去除“原始序列号”列；仅展示统一【序列号】。")
    st.dataframe(df_view.loc[:, show_cols], use_container_width=True)

    st.download_button(
        "下载 CSV（原始明细）",
        data=df_view.loc[:, show_cols].to_csv(index=False, encoding="utf-8-sig"),
        file_name="leaking_1131_raw_details.csv",
        mime="text/csv",
        use_container_width=True
    )
