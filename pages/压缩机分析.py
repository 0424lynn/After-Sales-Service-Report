# ==== 本页隐藏侧栏（更强兼容版）+ 放大字体 ====
import streamlit as st
import streamlit.components.v1 as components

# 只保留一次 set_page_config，放最前（如果被上层调用已设置则忽略异常）
try:
    st.set_page_config(page_title="压缩机分析（1.1.3.2）", layout="wide", initial_sidebar_state="collapsed")
except Exception:
    pass

# —— CSS：强制隐藏各种 Sidebar 容器与折叠按钮（覆盖不同版本的 data-testid）
st.markdown("""
<style>
aside[aria-label="sidebar"],
aside[aria-label="Sidebar"],
aside[class*="sidebar"],
[data-testid="stSidebar"],
[data-testid^="stSidebar"],
[data-testid*="Sidebar"],
[data-testid="stSidebarNav"],
[data-testid="stSidebarContent"],
[data-testid="stSidebarCollapsedControl"],
[data-testid*="Collapse"],
[data-testid*="collapsed"],
nav[aria-label="Sidebar navigation"]{
  display: none !important;
  width: 0 !important; min-width: 0 !important; max-width: 0 !important;
  visibility: hidden !important; pointer-events: none !important; opacity: 0 !important;
}
div.block-container{ padding-left: 1rem !important; padding-right: 1rem !important; margin-left: 0 !important; }
section.main, div.main, [data-testid="stAppViewContainer"]{ margin-left: 0 !important; }
</style>
""", unsafe_allow_html=True)

# —— JS：若用户点击“展开侧栏”，也立刻再隐藏
components.html("""
<script>
(function(){
  function nuke(){
    const sels = [
      'aside[aria-label="sidebar"]','aside[aria-label="Sidebar"]','aside[class*="sidebar"]',
      '[data-testid="stSidebar"]','[data-testid^="stSidebar"]','[data-testid*="Sidebar"]',
      '[data-testid="stSidebarNav"]','[data-testid="stSidebarContent"]','[data-testid="stSidebarCollapsedControl"]',
      '[data-testid*="Collapse"]','[data-testid*="collapsed"]','nav[aria-label="Sidebar navigation"]'
    ];
    for (const s of sels){
      document.querySelectorAll(s).forEach(el=>{
        el.style.display='none';
        el.style.width='0'; el.style.minWidth='0'; el.style.maxWidth='0';
        el.style.visibility='hidden'; el.style.pointerEvents='none'; el.style.opacity='0';
      });
    }
    const bc = document.querySelector('div.block-container');
    if (bc){ bc.style.paddingLeft='1rem'; bc.style.paddingRight='1rem'; }
  }
  nuke();
  new MutationObserver(nuke).observe(document.body, {subtree:true, childList:true, attributes:true});
  setInterval(nuke, 500);
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
st.title("压缩机配件分析 1.1.3.2")

# ================= 持久化路径 =================
STORE_PARQUET     = Path("data_store.parquet")
STORE_CSV         = Path("data_store.csv")
MODEL_FLEET_JSON  = Path("model_fleet_counts.json")   # 型号系列在保量（可选）
SPEC_FLEET_JSON   = Path("spec_fleet_counts.json")    # ✅ 具体型号在保量（通用）
CTRL_2COL_JSON    = Path("ctrl_model_map_2col.json")  # ✅ 两列映射，复用以便补充“物料编码”等（可空）
# —— 原始 Excel 库（用于补“原始生产日期/序列号”）
RAW_EXCEL_STORE = Path("raw_excel_store.parquet")

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


# === 统一兼容版：具体型号在保量 ===
def normalize_model_key(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.upper()

def _load_spec_fleet_any():
    if not SPEC_FLEET_JSON.exists():
        return [], {}
    try:
        data = json.loads(SPEC_FLEET_JSON.read_text(encoding="utf-8"))
    except Exception:
        return [], {}
    if isinstance(data, dict) and "triples" in data:
        triples = data["triples"]
    elif isinstance(data, dict):
        triples = [{"a": k, "b": "", "fleet": v} for k, v in data.items()]
    elif isinstance(data, list):
        triples = data
    else:
        triples = []
    mp = {}
    for r in triples:
        a = normalize_model_key(r.get("a", ""))
        b = normalize_model_key(r.get("b", ""))
        f = int(pd.to_numeric(r.get("fleet", 0), errors="coerce") or 0)
        if a: mp[a] = f
        if b: mp[b] = f
    return triples, mp

if "spec_fleet_triples" not in st.session_state or "spec_fleet_map" not in st.session_state:
    triples, mp = _load_spec_fleet_any()
    st.session_state.spec_fleet_triples = triples
    st.session_state.spec_fleet_map = mp

def lookup_spec_fleet(model_name: str) -> int:
    return int(st.session_state.spec_fleet_map.get(normalize_model_key(model_name), 0))

# ================== 两列粘贴映射（复用） ==================
def _norm_txt2(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _norm_key2(s: str) -> str:
    return _norm_txt2(s).upper()

def _parse_left_cell(cell: str):
    s = _norm_txt2(cell)
    m = re.match(r"^(?P<ctrl>.+?)\s+(?P<vendor>.+?)\/\s*$", s)
    if m:
        return _norm_txt2(m.group("ctrl")), _norm_txt2(m.group("vendor"))
    m2 = re.match(r"^(?P<ctrl>.+?)\s+(?P<vendor>[^\/]+)$", s)
    if m2:
        return _norm_txt2(m2.group("ctrl")), _norm_txt2(m2.group("vendor"))
    return s, ""

def _load_ctrl2col_rows() -> list[dict]:
    if CTRL_2COL_JSON.exists():
        try:
            data = json.loads(CTRL_2COL_JSON.read_text(encoding="utf-8"))
            if isinstance(data, list): return data
            if isinstance(data, dict) and "rows" in data: return data["rows"]
        except Exception:
            pass
    return []

def _save_ctrl2col_rows(rows: list[dict]) -> tuple[bool, str]:
    try:
        CTRL_2COL_JSON.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        return True, ""
    except Exception as e:
        return False, str(e)

if "ctrl2col_rows" not in st.session_state:
    st.session_state.ctrl2col_rows = _load_ctrl2col_rows()

def _build_ctrl_maps_2col(rows: list[dict]):
    ctrl2code, ctrl2vendor = {}, {}
    for r in rows:
        ctrl = _norm_key2(r.get("ctrl", "")) or _norm_key2(r.get("fan", ""))  # 兼容历史字段名
        if not ctrl: continue
        ctrl2code[ctrl] = _norm_txt2(r.get("code", ""))
        ctrl2vendor[ctrl] = _norm_txt2(r.get("vendor", ""))
    return ctrl2code, ctrl2vendor

CTRL2CODE, CTRL2VENDOR = _build_ctrl_maps_2col(st.session_state.ctrl2col_rows)

def ctrl_map_lookup(ctrl: str = "", *, product: str | None = None):
    rows = st.session_state.get("ctrl2col_rows", []) or []
    ctrl_index, product_index = {}, {}
    for r in rows:
        ctrl_u = _norm_key2(r.get("ctrl", "") or r.get("fan", ""))
        prod_u = normalize_model_key(r.get("product", ""))
        if ctrl_u:   ctrl_index.setdefault(ctrl_u, []).append(r)
        if prod_u:   product_index.setdefault(prod_u, []).append(r)

    def _choose_row(rr):
        for x in rr:
            if _norm_txt2(x.get("code", "")): return x
        return rr[0] if rr else None

    def _display(row):
        c = _norm_txt2(row.get("ctrl", "") or row.get("fan", ""))
        v = _norm_txt2(row.get("vendor", ""))
        return f"{c}  {v}/" if c and v else c

    if ctrl:
        ckey = _norm_key2(ctrl)
        if ckey in ctrl_index:
            hit = _choose_row(ctrl_index[ckey])
            if hit: return _display(hit), _norm_txt2(hit.get("code", ""))

    if product:
        pkey = normalize_model_key(product)
        if pkey in product_index:
            hit = _choose_row(product_index[pkey])
            if hit: return _display(hit), _norm_txt2(hit.get("code", ""))

    return _norm_txt2(ctrl), ""

with st.expander("（可选）部件/物料映射（两列粘贴；左=“型号  供应商/”，右=“物料编码”）", expanded=False):
    st.caption("此表在压缩机页用于补齐“物料编码”等信息；可留空。")
    pasted = st.text_area(
        "粘贴两列文本（每行一条；两列之间用Tab或≥2空格；左列末尾保留“/”更稳）：",
        height=140,
        placeholder="示例：\nCOMP-XYZ  EMBRACO/    R03003000011\nCOMP-ABC  SECOP/     R03003000022"
    )
    if st.button("解析并追加到下方表格", use_container_width=True):
        add_rows = []
        for line in pasted.splitlines():
            line = line.strip()
            if not line: continue
            parts = re.split(r"[\t]{1,}|\s{2,}", line)
            if len(parts) < 2:
                parts = re.split(r"\s+", line)
                if len(parts) >= 2:
                    left = " ".join(parts[:-1]); right = parts[-1]
                else:
                    continue
            else:
                left, right = parts[0], parts[1]
            ctrl, vendor = _parse_left_cell(left)
            code = _norm_txt2(right)
            add_rows.append({"ctrl": ctrl, "vendor": vendor, "code": code, "product": ""})

        if add_rows:
            cols = ["ctrl","vendor","code","product"]
            cur = pd.DataFrame(st.session_state.ctrl2col_rows) if st.session_state.ctrl2col_rows else pd.DataFrame(columns=cols)
            cur = cur.reindex(columns=cols)
            if not cur.empty: cur["K"] = cur["ctrl"].map(_norm_key2)
            for r in add_rows:
                k = _norm_key2(r["ctrl"])
                if not cur.empty and (cur["K"] == k).any():
                    idx = cur.index[cur["K"] == k][0]
                    cur.loc[idx, ["vendor","code"]] = [r["vendor"], r["code"]]
                else:
                    cur = pd.concat([cur, pd.DataFrame([r])], ignore_index=True)
            if "K" in cur.columns: cur = cur.drop(columns=["K"])
            st.session_state.ctrl2col_rows = cur.fillna("").to_dict(orient="records")
            CTRL2CODE, CTRL2VENDOR = _build_ctrl_maps_2col(st.session_state.ctrl2col_rows)
            st.success(f"已追加/覆盖 {len(add_rows)} 条。")

    _rows = st.session_state.ctrl2col_rows or []
    view2 = (
        pd.DataFrame(_rows)
        .reindex(columns=["ctrl","vendor","product","code"])
        .rename(columns={"ctrl":"型号","vendor":"供应商","product":"产品型号（可空）","code":"物料编码"})
    )
    view2 = view2 if not view2.empty else pd.DataFrame(columns=["型号","供应商","产品型号（可空）","物料编码"])
    edited2 = st.data_editor(view2, num_rows="dynamic", use_container_width=True, key="comp_map_editor")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("保存到本地（ctrl_model_map_2col.json）", use_container_width=True):
            rows = []
            for _, r in edited2.iterrows():
                rows.append({"ctrl": _norm_txt2(r.get("型号","")),
                             "vendor": _norm_txt2(r.get("供应商","")),
                             "code": _norm_txt2(r.get("物料编码","")),
                             "product": _norm_txt2(r.get("产品型号（可空）",""))})
            ok, msg = _save_ctrl2col_rows(rows)
            if ok:
                st.session_state.ctrl2col_rows = rows
                CTRL2CODE, CTRL2VENDOR = _build_ctrl_maps_2col(rows)
                st.success(f"已保存 {len(rows)} 条。")
            else:
                st.error(f"保存失败：{msg}")
    with c2:
        if not edited2.empty:
            st.download_button("导出当前表为 CSV",
                               data=edited2.to_csv(index=False, encoding="utf-8-sig"),
                               file_name="comp_model_map_2col.csv",
                               mime="text/csv", use_container_width=True)
    with c3:
        if st.button("清空（仅内存，不落盘）", use_container_width=True):
            st.session_state.ctrl2col_rows = []
            CTRL2CODE, CTRL2VENDOR = _build_ctrl_maps_2col([])
            st.rerun()

# ================= 工具函数（轻量清洗） =================
def ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Date","Category1","Category2","TAG1","TAG2","TAG3","TAG4",
            "CostUSD","PaidQty","IsWarranty","Completed","Year","Month","Quarter","_SrcYM"
        ])
    mapping = {}
    for c in df.columns:
        lc = str(c).strip().lower()
        if lc in ["日期","date","created_at","发生日期","生产日期"]: mapping[c] = "Date"
        elif lc in ["分类1","category1","大类"]: mapping[c] = "Category1"
        elif lc in ["分类2","category2","小类"]: mapping[c] = "Category2"
        elif lc in ["tag1","tag 1","tag-1"]: mapping[c] = "TAG1"
        elif lc in ["tag2","tag 2","tag-2"]: mapping[c] = "TAG2"
        elif lc in ["tag3","tag 3","tag-3"]: mapping[c] = "TAG3"
        elif lc in ["tag4","tag 4","tag-4"]: mapping[c] = "TAG4"
        elif lc in ["cost","costusd","amount","费用","金额","fee","cost(usd)","cost usd"]: mapping[c] = "CostUSD"
        elif lc in ["paidqty","paid_qty","已付款数量"]: mapping[c] = "PaidQty"
        elif lc in ["iswarranty","warranty","是否承保"]: mapping[c] = "IsWarranty"
        elif lc in ["completed","是否完成","维修状态","状态"]: mapping[c] = "Completed"
    if mapping: df = df.rename(columns=mapping)
    for col, default in [("Category1",""),("Category2",""),("TAG1",""),("TAG2",""),("TAG3",""),("TAG4",""),
                         ("CostUSD",0.0),("PaidQty",0),("IsWarranty",False),("Completed",False)]:
        if col not in df.columns: df[col] = default
    df["CostUSD"] = pd.to_numeric(df["CostUSD"], errors="coerce").fillna(0.0)
    df["PaidQty"] = pd.to_numeric(df["PaidQty"], errors="coerce").fillna(0).astype(int)
    if "Year" in df.columns:  df["Year"]  = pd.to_numeric(df["Year"], errors="coerce")
    if "Month" in df.columns: df["Month"] = pd.to_numeric(df["Month"], errors="coerce")
    return df

def _clean_text(s: str) -> str:
    s = ("" if pd.isna(s) else str(s)).lower().replace("\u3000", " ")
    s = re.sub(r"[\s\-_\/]+", " ", s)
    return s.strip()

def get_param(name: str, default=None, cast=int):
    try:
        v = st.query_params.get(name, default)
    except Exception:
        qp = st.experimental_get_query_params()
        v = (qp.get(name, [default]) or [default])[0]
    if v is None: return default
    if cast is None: return v
    try: return cast(v)
    except Exception: return default

def get_param_str(name: str, default=None):
    try:
        v = st.query_params.get(name, default)
    except Exception:
        qp = st.experimental_get_query_params()
        v = (qp.get(name, [default]) or [default])[0]
    if isinstance(v, str): return unquote(v)
    return default

def load_store_df_from_disk() -> pd.DataFrame:
    if STORE_PARQUET.exists():
        try: return pd.read_parquet(STORE_PARQUET)
        except Exception: pass
    if STORE_CSV.exists():
        try: return pd.read_csv(STORE_CSV)
        except Exception: pass
    return pd.DataFrame()

def get_store_df() -> pd.DataFrame:
    if "store_df" in st.session_state and isinstance(st.session_state["store_df"], pd.DataFrame):
        return ensure_cols(st.session_state["store_df"].copy())
    return ensure_cols(load_store_df_from_disk())

def money(v) -> str:
    try: return f"${float(v):,.0f}"
    except Exception: return "$0"

# ================= 参数与数据读取 =================
code  = get_param_str("code", default="1.1.3.2")   # ← 默认 1.1.3.2（压缩机）
year  = get_param("year", default=None, cast=int)
month = get_param("month", default=None, cast=int)

# 新增：支持多月/季度
months_param = get_param_str("months", default=None)   # 形如 "1,2,3"
q_param      = get_param_str("q", default=None)        # 形如 "Q1"

df_all = get_store_df()
if df_all.empty:
    st.warning("未获取到任何数据：请先在“冰箱数据”页面上传数据；或确认 data_store.parquet / data_store.csv 是否存在。")
    st.stop()

yy = pd.to_numeric(df_all.get("Year"),  errors="coerce")
mm = pd.to_numeric(df_all.get("Month"), errors="coerce")

# —— 年份：优先取参数；否则用数据中最大年
if year is None:
    if yy.notna().any():
        year = int(yy.dropna().max())
    else:
        year = 2025  # 兜底

# —— 解析月份集合（优先级：months > month > q > 当年最新月）
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

# 期间最后一月（在保量口径用）
month_for_fleet = int(sorted(sel_months)[-1])

# —— 过滤数据
mask_year   = (yy == int(year))
mask_months = mm.isin([int(m) for m in sel_months])
scope_df    = df_all[mask_year & mask_months].copy()

# —— 口径标题
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

if scope_df.empty:
    st.info(f"{scope_text} 没有可统计数据。")
    st.stop()

# —— 根据缺陷码取子集
joined_all = scope_df.astype(str).agg(" ".join, axis=1).map(_clean_text)
pat_code   = re.compile(rf"(?<![0-9a-z]){re.escape(code)}(?![0-9a-z])", re.I)
mask_code  = joined_all.str.contains(pat_code)
df_code    = scope_df[mask_code].copy()

st.caption(f"统计口径：{scope_text} | 缺陷码：{code}")
if df_code.empty:
    st.warning("该口径下未命中任何 1.1.3.2 相关记录。")
    st.stop()


# ================= 顶部抬头条 =================
def render_topbar(_code: str, _scope_text: str, df_scope: pd.DataFrame):
    paid = pd.to_numeric(df_scope.get("CostUSD", 0), errors="coerce").fillna(0)
    qty  = len(df_scope)
    cost = float(paid[paid > 0].sum())
    avg  = float(paid[paid > 0].mean()) if (paid > 0).any() else 0.0
    col1, col2, col3, col4 = st.columns([1.2, 1, 1, 1])
    with col1:
        st.markdown(
            f"""
            <div style="padding:12px 16px;border-radius:14px;background:#0f172a;color:#fff;">
              <div style="font-size:13px;opacity:.7;">缺陷码 / 统计口径</div>
              <div style="font-size:20px;font-weight:700;line-height:1.2;margin-top:2px;">{_code}</div>
              <div style="font-size:13px;margin-top:2px;">{_scope_text}</div>
            </div>
            """, unsafe_allow_html=True
        )
    with col2: st.metric("记录数", f"{qty}")
    with col3: st.metric("已付款总费用", money(cost))
    with col4: st.metric("单均已付款费用", money(avg))

render_topbar(code, scope_text, df_code)
st.markdown("---")

# ================= 1.1.3.2.x（压缩机）分项 =================
def build_1132_subtable(df_in: pd.DataFrame) -> pd.DataFrame:
    join = df_in.astype(str).agg(" ".join, axis=1).str.lower().str.replace("\u3000", " ", regex=False)
    join = join.str.replace(r"[\s\-_\/]+", " ", regex=True).str.strip()
    def _pat(s): 
        return re.compile(rf"(?<![0-9a-z]){re.escape(s)}(?![0-9a-z])", re.I)

    # 仅一个明确子类：1.1.3.2.1（压缩机组件）
    pats = {
        "1.1.3.2.1 Start Components/ 3 in 1": ("压缩机组件", _pat("1.1.3.2.1")),
    }

    mask_any_sub = np.zeros(len(df_in), dtype=bool)
    rows = []

    # 子类行
    for en_name, (cn_name, pat) in pats.items():
        m = join.str.contains(pat)
        mask_any_sub |= m.values
        qty = int(m.sum())
        fee = float(pd.to_numeric(df_in.loc[m, "CostUSD"], errors="coerce").fillna(0).sum())
        rows.append({"英文名称": en_name, "质量问题": cn_name, "数量": qty, "费用": fee})

    # 顶层 1.1.3.2（不含 .2.1）
    p_top = _pat("1.1.3.2")
    m_top_only = join.str.contains(p_top) & (~mask_any_sub)
    qty_top = int(m_top_only.sum())
    fee_top = float(pd.to_numeric(df_in.loc[m_top_only, "CostUSD"], errors="coerce").fillna(0).sum())
    rows.insert(0, {"英文名称": "1.1.3.2 Compressor", "质量问题": "压缩机", "数量": qty_top, "费用": fee_top})

    # 总计
    df_out = pd.DataFrame(rows)
    total_row = {"英文名称": "总计", "质量问题": "", "数量": int(df_out["数量"].sum()), "费用": float(df_out["费用"].sum())}
    df_out = pd.concat([df_out, pd.DataFrame([total_row])], ignore_index=True)
    return df_out

# —— 渲染子项表
df_1132 = build_1132_subtable(df_code)
st.markdown("### 1.1.3.2.x 分项（压缩机）")
if not df_1132.empty:
    df_view = df_1132.copy()
    df_view["费用"] = df_view["费用"].map(lambda v: f"${float(v):,.0f}")
    st.dataframe(df_view, use_container_width=True)
    st.download_button(
        "下载 CSV（1.1.3.2 分项）",
        data=df_1132.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"compressor_1132_breakdown_{scope_text.replace(' ','_')}.csv",
        mime="text/csv",
        use_container_width=True
    )
else:
    st.info("暂无 1.1.3.2 相关记录。")

# ================= 三张统计表（复用你现有 build_tables） =================
USE_MONTH_CAND = [
    "已使用保修时长","已使用保修月数","保修已使用时长","已使用月份","保修使用月数",
    "使用月数","使用月","ServiceMonths","months_in_use","use_months"
]
SERIES_PRI_CAND = ["型号归类","型号系列","series","model_series"]
SERIES_SEC_CAND = ["机型","产品型号","MODEL","Model"]
SPEC_MODEL_CAND = ["型号","具体型号","机型明细","MODEL","Model","Part Model","PartModel","Part"]
CTRL_KIND_CAND  = ["压缩机型号","控制器型号","Controller","ControllerModel","Thermo","Thermostat"]   # 对压缩机可为空
PART_CODE_CAND  = ["R码","物料编码","物料号","PartCode","Part No"]

MODEL_CATEGORY_MAP = {
    "MBF":"冰箱","MCF":"陈列柜","MSF":"沙拉台","MGF":"工作台","MPF":"玻璃台","CHEF BASE":"抽屉工作台",
    "MBC":"滑盖冷藏柜","MBB":"卧式吧台","SBB":"移门吧台","MKC":"啤酒柜","MBGF":"滑盖冷冻柜",
    "AMC":"牛奶柜","RDCS":"方形展示柜","CRDC":"弧形展示柜","CRDS":"台面方形展示柜","ATHOM":"饮料柜",
    "AOM":"立式饮料柜","MSCT":"配料台",
}
MODEL_ORDER = ["AOM","CHEF BASE","MBC","MBF","MBGF","MCF","MGF","MPF","MSF"]
ALIASES = {"CHEFBASE":"CHEF BASE","CHEF-BASE":"CHEF BASE","CHEF_BASE":"CHEF BASE","CHEF BASE":"CHEF BASE"}

def resolve_model_key(raw: str) -> str:
    s = (raw or "").strip().replace("\u3000"," ")
    s_norm = re.sub(r"[\s\-_\/]+", " ", s).strip()
    key_up = s_norm.upper().replace(" ", "")
    for k,v in ALIASES.items():
        if key_up == k.upper().replace(" ",""): return v
    toks = s_norm.split()
    if len(toks)>1: return " ".join(t.upper() for t in toks)
    return s_norm.upper()

def _pick_col(df: pd.DataFrame, cands: list):
    for c in cands:
        if c in df.columns: return c
    norm = {re.sub(r"[\s_]+","", str(c).lower()): c for c in df.columns}
    for want in cands:
        k = re.sub(r"[\s_]+","", str(want).lower())
        if k in norm: return norm[k]
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
            yy = max(int(y) for y in fleet.keys())
            mm = max(int(m) for m in fleet[str(yy)].keys())
            year, month = yy, mm
        except Exception:
            return {k:0 for k in MODEL_CATEGORY_MAP.keys()}
    y, m = str(int(year)), str(int(month))
    return {k:int(fleet.get(y,{}).get(m,{}).get(k,0)) for k in MODEL_CATEGORY_MAP.keys()}

def _bucket_series(months: pd.Series):
    m = pd.to_numeric(months, errors="coerce").where(lambda x: x > 0)
    return pd.cut(m, bins=[0, 4, 8, np.inf],
                  labels=["短期（1-4月）","中期（5-8月）","长期（9-24月）"],
                  include_lowest=True, right=True)

def build_tables(df_subset: pd.DataFrame, *, year=None, month=None):
    months_col = _pick_col(df_subset, USE_MONTH_CAND)
    series_col = _pick_col(df_subset, SERIES_PRI_CAND) or _pick_col(df_subset, SERIES_SEC_CAND)
    spec_col   = _pick_col(df_subset, SPEC_MODEL_CAND)
    ctype_col  = _pick_col(df_subset, CTRL_KIND_CAND)  # 压缩机页通常没有，不影响

    if months_col is None:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    if series_col is None and spec_col is not None:
        tmp = df_subset[spec_col].astype(str).str.upper().str.extract(r"^([A-Z]{2,4})")
        df_subset = df_subset.copy()
        df_subset["__SERIES__"] = tmp[0]
        series_col = "__SERIES__"

    m = pd.to_numeric(df_subset[months_col], errors="coerce")
    valid_mask = m.notna() & (m > 0)
    df_valid = df_subset[valid_mask].copy()
    if df_valid.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_valid["__bucket__"] = _bucket_series(df_valid[months_col])

        # 表1：按系列
    tbl1 = pd.DataFrame()
    if series_col is not None:
        g = pd.pivot_table(
            df_valid, index=series_col, columns="__bucket__", values="Date",
            aggfunc="count", fill_value=0
        )
        for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）"]:
            if c not in g.columns: g[c] = 0
        g = g[["短期（1-4月）","中期（5-8月）","长期（9-24月）"]]
        g["总计"] = g.sum(axis=1)
        g.reset_index(inplace=True)
        g.rename(columns={series_col: "型号"}, inplace=True)

        # 在保量与问题比例
        fleet = load_model_fleet()
        month_fleet = get_fleet_month(fleet, year, month)

        def _fleet_lookup(s):
            key = resolve_model_key(s)
            return int(month_fleet.get(key, 0))

        g["在保量"] = g["型号"].map(_fleet_lookup).fillna(0).astype(int)
        g["问题比例(%)"] = np.where(g["在保量"] > 0, g["总计"] / g["在保量"] * 100.0, np.nan)

        # 固定系列排序（可选）
        order_map = {k:i for i,k in enumerate(MODEL_ORDER)}
        g["__ord__"] = g["型号"].map(order_map).fillna(9999)
        g = g.sort_values(["__ord__","总计"], ascending=[True, False]).drop(columns="__ord__")

        tbl1 = g
    

    # 表2：按具体型号
    tbl2 = pd.DataFrame()
    if spec_col is not None:
        g2 = pd.pivot_table(df_valid, index=spec_col, columns="__bucket__", values="Date",
                            aggfunc="count", fill_value=0)
        for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）"]:
            if c not in g2.columns: g2[c] = 0
        g2 = g2[["短期（1-4月）","中期（5-8月）","长期（9-24月）"]]
        g2["总计"] = g2.sum(axis=1)
        g2.reset_index(inplace=True)
        g2.rename(columns={spec_col: "具体型号"}, inplace=True)

        # “压缩机型号/物料编码”字段：在压缩机页等价于“部件名称/物料编码”，从映射或众数补齐（若无则空）
        def _mode(series):
            s = series.dropna().astype(str)
            if s.empty: return ""
            m0 = s.mode()
            return m0.iloc[0] if not m0.empty else s.iloc[0]
        if ctype_col:
            base_ctrl = (
                df_subset.groupby(spec_col)[ctype_col].apply(_mode)
                .reindex(g2["具体型号"]).fillna("").values
            )
        else:
            base_ctrl = [""] * len(g2)

        out_ctrl_display, out_code = [], []
        for i, row in g2.iterrows():
            spec_model = _norm_txt2(row["具体型号"])
            ctrl_guess = _norm_txt2(base_ctrl[i]) if isinstance(base_ctrl, (list, np.ndarray)) else _norm_txt2(base_ctrl)
            display_ctrl, code_from_map = ctrl_map_lookup(ctrl=ctrl_guess, product=spec_model)
            out_ctrl_display.append(display_ctrl if display_ctrl else ctrl_guess)
            out_code.append(code_from_map)

        g2["压缩机型号"] = out_ctrl_display    # 名称沿用字段，含义=“部件/型号”
        g2["物料编码"] = out_code

        g2["在保量"] = g2["具体型号"].map(lookup_spec_fleet).fillna(0).astype(int)
        g2["比例(%)"] = np.where(g2["在保量"]>0, g2["总计"]/g2["在保量"]*100.0, np.nan)
        g2 = g2.sort_values("总计", ascending=False)
        for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量"]:
            if c not in g2.columns: g2[c] = 0
        for c in ["压缩机型号","物料编码"]:
            if c not in g2.columns: g2[c] = ""
        tbl2 = g2

    # 表3：由“按具体型号”汇总 -> “压缩机型号”(=部件/型号)
    tbl3 = pd.DataFrame()
    if spec_col is not None and not tbl2.empty:
        tmp = tbl2.copy()
        tmp["压缩机型号"] = tmp["压缩机型号"].replace("", np.nan).fillna("(未识别)")
        sum_cols = ["短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量"]
        for c in sum_cols:
            if c not in tmp.columns: tmp[c] = 0
        g3 = tmp.groupby("压缩机型号", as_index=False)[sum_cols].sum()
        g3["问题比例(%)"] = np.where(g3["在保量"] > 0, g3["总计"] / g3["在保量"] * 100.0, np.nan)
        g3 = g3.sort_values("总计", ascending=False)
        tbl3 = g3

    # 不需要“按系列”卡片，返回空表占位
    return tbl1, tbl2, tbl3

# —— 渲染（仅“按具体型号 / 按压缩机型号”两张）
t1, t2, t3 = build_tables(df_code, year=year, month=month_for_fleet)
base1 = pd.DataFrame(columns=["型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])
base2 = pd.DataFrame(columns=["具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","压缩机型号","物料编码","总计","在保量","比例(%)"])
base3 = pd.DataFrame(columns=["压缩机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])

def _fmt_pct(s):
    return "" if pd.isna(s) else f"{float(s):.2f}%"

def render_sortable_table(title: str, df: pd.DataFrame, *, order_cols):
    cols = [c for c in order_cols if c in df.columns]
    view = df.loc[:, cols].copy() if cols else df.copy()
    for pct_col in ["问题比例(%)", "比例(%)"]:
        if pct_col in view.columns:
            view[pct_col] = view[pct_col].map(_fmt_pct)
    st.subheader(title)
    st.caption("提示：点击列头可排序；再次点击切换升/降序；按住 Shift 支持多列排序。")
    st.dataframe(view, use_container_width=True)
render_sortable_table("压缩机（按系列）", (t1 if not t1.empty else base1),
                      order_cols=("型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))    

render_sortable_table("压缩机（按具体型号）", (t2 if not t2.empty else base2),
                      order_cols=("具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","压缩机型号","物料编码","总计","在保量","比例(%)"))

render_sortable_table("压缩机（按部件/型号）", (t3 if not t3.empty else base3),
                      order_cols=("压缩机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))

# ===== 分项深入分析（仅数量>0，含 1.1.3.2.1；默认折叠）=====
st.markdown("## 分项深入分析（仅数量>0，含 1.1.3.2.1）")

def _clean_join(df_):
    j = df_.astype(str).agg(" ".join, axis=1).str.lower().str.replace("\u3000", " ", regex=False)
    return j.str.replace(r"[\s\-_\/]+", " ", regex=True).str.strip()

def _subset_1132_item(df_src: pd.DataFrame, en_name: str) -> pd.DataFrame:
    if df_src is None or df_src.empty: return df_src.iloc[0:0].copy()
    j = _clean_join(df_src)
    p_sub = re.compile(r"(?<![0-9a-z])1\.1\.3\.2\.1(?![0-9a-z])", re.I)
    if re.search(r"^1\.1\.3\.2\.1\b", str(en_name or "")):
        return df_src[j.str.contains(p_sub)].copy()
    return df_src.iloc[0:0].copy()

rows_need = (
    df_1132[
        df_1132["英文名称"].astype(str).str.match(r"^1\.1\.3\.2\.1\b", na=False) &
        (pd.to_numeric(df_1132["数量"], errors="coerce").fillna(0) > 0)
    ]
    .sort_values("数量", ascending=False)
    .copy()
)

if rows_need.empty:
    st.info("当前 1.1.3.2.1 没有数量>0 的条目可深入分析。")
else:
    for _, r in rows_need.iterrows():
        en_name = str(r["英文名称"]); zh_name = str(r.get("质量问题",""))
        sub_df = _subset_1132_item(df_code, en_name)
        with st.expander(f"【{en_name}】{zh_name}｜记录 {int(r['数量'])}，费用 {money(float(r.get('费用', 0)))}",
                         expanded=False):  # ← 默认折叠
            a1, a2, a3 = build_tables(sub_df, year=year, month=month_for_fleet)
            _base2 = pd.DataFrame(columns=["具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","压缩机型号","物料编码","总计","在保量","比例(%)"])
            _base3 = pd.DataFrame(columns=["压缩机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])
            render_sortable_table("按具体型号", (a2 if not a2.empty else _base2),
                                  order_cols=("具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","压缩机型号","物料编码","总计","在保量","比例(%)"))
            render_sortable_table("按部件/型号", (a3 if not a3.empty else _base3),
                                  order_cols=("压缩机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))

# ================= 原始明细（命中 1.1.3.2 的记录） =================
with st.expander("原始明细（命中 1.1.3.2 的记录）", expanded=False):
    # —— 统一把 231221 / 202405 / 2024-5-6 … 转成 YYYY-MM-DD
    def _norm_date_text(raw: pd.Series) -> pd.Series:
        s = raw.astype(str).str.strip()
        out = pd.Series("", index=s.index, dtype="object")

        # 8位：YYYYMMDD
        m8 = s.str.fullmatch(r"\d{8}", na=False)
        if m8.any():
            dt8 = pd.to_datetime(s[m8], format="%Y%m%d", errors="coerce")
            out.loc[m8] = dt8.dt.strftime("%Y-%m-%d")

        # 6位：优先 YYMMDD（如 231221）；若是 YYYYMM（如 202405）则补 01 日
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

        # 其余交给 to_datetime
        rest = (out == "")
        if rest.any():
            dt = pd.to_datetime(s[rest], errors="coerce", infer_datetime_format=True)
            out.loc[rest] = dt.dt.strftime("%Y-%m-%d")

        # 失败保留原值
        failed = out.isna() | (out == "NaT") | (out == "")
        if failed.any():
            out.loc[failed] = s.loc[failed]
        return out

    # 基表
    df_view = df_code.copy()

    # —— 与“原始 Excel 表映射”关联，补【原始生产日期】【原始序列号】
    RAW_PROD_CANDS   = ["生产日期","生产日","ProductionDate","ProdDate","制造日期","出厂日期"]
    RAW_SERIAL_CANDS = ["序列号","S/N","SN","Serial No","SerialNo","Serial","序列号SN"]
    DEDUP_KEYS = [
        "工单号","索赔号","CaseNo","TicketNo",
        "S/N","Serial No","SN","Serial","SerialNo",
        "序列号","序列号SN"
    ]

    def _pick_first(df, cands):
        for c in cands:
            if c in df.columns:
                return c
        return None

    raw_hist = _load_raw_excel_store()
    if not raw_hist.empty:
        join_keys = [k for k in DEDUP_KEYS if (k in df_view.columns) and (k in raw_hist.columns)]
        raw_prod_col   = _pick_first(raw_hist, RAW_PROD_CANDS)
        raw_serial_col = _pick_first(raw_hist, RAW_SERIAL_CANDS)

        if join_keys and (raw_prod_col or raw_serial_col):
            take_cols = [k for k in join_keys]
            if raw_prod_col   and raw_prod_col   not in take_cols: take_cols.append(raw_prod_col)
            if raw_serial_col and raw_serial_col not in take_cols: take_cols.append(raw_serial_col)

            patch = raw_hist.loc[:, take_cols].drop_duplicates().copy()

            # 不改 join key 的列名；仅复制出“原始*”列
            if raw_prod_col:
                if raw_prod_col in join_keys:
                    patch["原始生产日期"] = _norm_date_text(patch[raw_prod_col])
                else:
                    patch = patch.rename(columns={raw_prod_col: "原始生产日期"})
                    patch["原始生产日期"] = _norm_date_text(patch["原始生产日期"])
            if raw_serial_col:
                if raw_serial_col in join_keys:
                    patch["原始序列号"] = patch[raw_serial_col]
                else:
                    patch = patch.rename(columns={raw_serial_col: "原始序列号"})

            df_view = df_view.merge(patch, on=join_keys, how="left", suffixes=("", "_rawdup"))
            st.caption("（已从📎 原始 Excel 表映射按 %s 关联补入：%s）" % (
                ", ".join(join_keys),
                "、".join([c for c in ["原始生产日期","原始序列号"] if c in df_view.columns]) or "无"
            ))
    else:
        st.caption("（未找到📎 原始 Excel 表映射，显示现有列）")

    # —— 若没从原始库取到，也尝试用当前表中的候选列生成“原始生产日期”
    if "原始生产日期" not in df_view.columns:
        date_src = next((c for c in ["原始生产日期","生产日期","生产日","productiondate","proddate","制造日期","出厂日期","Date"] if c in df_view.columns), None)
        if date_src:
            df_view["原始生产日期"] = _norm_date_text(df_view[date_src])
        else:
            df_view["原始生产日期"] = ""

    # —— 新增统一“序列号”列（优先用“原始序列号”，否则回退到常见列）
    SERIAL_CANDS_SHOW = ["原始序列号","序列号","S/N","SN","Serial No","SerialNo","Serial","序列号SN"]
    serial_src = next((c for c in SERIAL_CANDS_SHOW if c in df_view.columns), None)
    df_view["序列号"] = df_view[serial_src].astype(str) if serial_src else ""

    # —— 展示列顺序
    base_cols = ["Category1","Category2","TAG1","TAG2","TAG3","TAG4","CostUSD","PaidQty","IsWarranty","Completed"]
    base_cols = [c for c in base_cols if c in df_view.columns]
    show_cols = ["原始生产日期","序列号"] + base_cols

    st.dataframe(df_view.loc[:, show_cols], use_container_width=True)

    st.download_button(
        "下载 CSV（原始明细）",
        data=df_view.loc[:, show_cols].to_csv(index=False, encoding="utf-8-sig"),
        file_name="compressor_1132_raw_details.csv",
        mime="text/csv",
        use_container_width=True
    )
