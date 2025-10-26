# ==== 本页隐藏侧栏 + 放大字体（完整段） ====
import streamlit as st
import streamlit.components.v1 as components

# 仅保留一次 set_page_config
try:
    st.set_page_config(page_title="风机配件分析（1.1.3.5）", layout="wide", initial_sidebar_state="collapsed")
except Exception:
    pass

# —— 样式：标题、表头、单元格、编辑器、table 兜底、行距等
st.markdown("""
<style>
/* ===== 标题更大更醒目 ===== */
h1, .stMarkdown h1 { font-size: 2.2rem !important; font-weight: 800 !important; line-height: 1.25; }
h2, .stMarkdown h2 { font-size: 1.8rem !important; font-weight: 800 !important; line-height: 1.30; }
h3, .stMarkdown h3 { font-size: 1.4rem !important; font-weight: 800 !important; line-height: 1.35; }

/* ====== st.dataframe（AG Grid）——表头更大 + 加粗 ====== */
/* 兼容不同主题类名：ag-theme-*, st-ag-grid */
div[data-testid="stDataFrame"] .ag-root-wrapper,
div[data-testid="stDataFrame"] .ag-root,
div[data-testid="stDataFrame"] .ag-theme-alpine,
div[data-testid="stDataFrame"] .ag-theme-balham,
div[data-testid="stDataFrame"] .st-ag-grid {
  --ag-font-size: 16px;        /* AG Grid 的全局字体变量（内容区） */
}

/* 表头 */
div[data-testid="stDataFrame"] .ag-header,
div[data-testid="stDataFrame"] .ag-header-cell,
div[data-testid="stDataFrame"] .ag-header-cell-text,
div[data-testid="stDataFrame"] .ag-header-group-cell,
div[data-testid="stDataFrame"] .ag-header-group-text {
  font-size: 18px !important;   /* 表头字号 */
  font-weight: 800 !important;  /* 表头加粗 */
  line-height: 1.2 !important;
}

/* 内容区（仅放大不加粗） */
div[data-testid="stDataFrame"] .ag-center-cols-container .ag-cell,
div[data-testid="stDataFrame"] .ag-center-cols-container .ag-cell-value,
div[data-testid="stDataFrame"] .ag-pinned-left-cols-container .ag-cell,
div[data-testid="stDataFrame"] .ag-pinned-right-cols-container .ag-cell {
  font-size: 16px !important;   /* 内容字号 */
  font-weight: 400 !important;  /* 内容不加粗 */
  line-height: 1.5 !important;  /* 放大后更易读 */
}

/* AG Grid 行号/索引区也跟随放大 */
div[data-testid="stDataFrame"] .ag-pinned-left-cols-container .ag-cell {
  font-size: 16px !important;
  font-weight: 400 !important;
}

/* ====== st.data_editor —— 放大但不加粗 ====== */
div[data-testid="stDataEditor"] {
  font-size: 16px !important;
  font-weight: 400 !important;
}
div[data-testid="stDataEditor"] .st-de-header {
  font-size: 18px !important;     /* 编辑器的列头 */
  font-weight: 800 !important;
}
div[data-testid="stDataEditor"] .st-de-cell {
  font-size: 16px !important;     /* 编辑器单元格 */
  font-weight: 400 !important;
  line-height: 1.5 !important;
}

/* ====== st.table 兜底：有些场景不是 AG Grid ====== */
div[data-testid="stTable"] table thead tr th {
  font-size: 18px !important;
  font-weight: 800 !important;
}
div[data-testid="stTable"] table tbody tr td {
  font-size: 16px !important;
  font-weight: 400 !important;
  line-height: 1.5 !important;
}

/* ====== caption、metric、小字等微调（不加粗） ====== */
.block-container .stCaption, .stMarkdown small, .stMarkdown .caption {
  font-size: 1rem !important;
  font-weight: 400 !important;
}
[data-testid="stMetricValue"] { font-size: 1.4rem !important; }
[data-testid="stMetricLabel"] { font-size: 0.95rem !important; }

/* ====== 让 DataFrame 宽一点，文本不被挤压 ====== */
div[data-testid="stDataFrame"] .ag-root-wrapper { font-size: 16px !important; }

/* ======（可选）行高略增，便于阅读 ====== */
div[data-testid="stDataFrame"] .ag-row,
div[data-testid="stDataEditor"] .st-de-row {
  line-height: 1.5 !important;
}
</style>
""", unsafe_allow_html=True)

# —— 仅本页隐藏侧栏（DOM 变动时也再次隐藏）
components.html("""
<script>
(function(){
  function hideSidebar(){
    const sels=[
      'aside[data-testid="stSidebar"]',
      'section[data-testid="stSidebar"]',
      'div[data-testid="stSidebar"]',
      'nav[aria-label="Sidebar navigation"]',
      '[data-testid="stSidebarNav"]',
      '[data-testid="stSidebarContent"]',
      '[data-testid="collapsedControl"]'
    ];
    for(const s of sels){
      document.querySelectorAll(s).forEach(el=>{
        el.style.display='none';
        el.style.width='0'; el.style.minWidth='0'; el.style.maxWidth='0';
        el.style.visibility='hidden'; el.style.pointerEvents='none'; el.style.opacity='0';
      });
    }
    const bc=document.querySelector('div.block-container');
    if(bc){ bc.style.paddingLeft='1rem'; bc.style.paddingRight='1rem'; }
  }
  hideSidebar();
  new MutationObserver(hideSidebar).observe(document.body,{subtree:true,childList:true,attributes:true});
})();
</script>
""", height=0)
# ==== 本段结束 ====

import streamlit as st
try:
    st.set_page_config(layout="wide", initial_sidebar_state="collapsed")
except Exception:
    pass

st.markdown("""
<style>
aside[data-testid="stSidebar"],
section[data-testid="stSidebar"],
div[data-testid="stSidebar"],
nav[aria-label="Sidebar navigation"],
[data-testid="stSidebarNav"],
[data-testid="stSidebarContent"],
[data-testid="collapsedControl"]{
  display:none !important; width:0 !important; min-width:0 !important; max-width:0 !important;
  visibility:hidden !important; pointer-events:none !important; opacity:0 !important;
}
div.block-container{ padding-left:1rem !important; padding-right:1rem !important; }
</style>
""", unsafe_allow_html=True)
import streamlit.components.v1 as components
components.html("""
<script>
(function(){
  function kill(){
    const sels=[
      'aside[data-testid="stSidebar"]',
      'section[data-testid="stSidebar"]',
      'div[data-testid="stSidebar"]',
      'nav[aria-label="Sidebar navigation"]',
      '[data-testid="stSidebarNav"]',
      '[data-testid="stSidebarContent"]',
      '[data-testid="collapsedControl"]'
    ];
    for(const s of sels){
      document.querySelectorAll(s).forEach(el=>{
        el.style.display='none';
        el.style.width='0'; el.style.minWidth='0'; el.style.maxWidth='0';
        el.style.visibility='hidden'; el.style.pointerEvents='none'; el.style.opacity='0';
      });
    }
    const bc=document.querySelector('div.block-container');
    if(bc){ bc.style.paddingLeft='1rem'; bc.style.paddingRight='1rem'; }
  }
  kill();
  new MutationObserver(kill).observe(document.body,{subtree:true,childList:true,attributes:true});
})();
</script>
""", height=0)
# ==== 隐藏侧栏结束 ====

# ================= 基础依赖 =================
import pandas as pd
import numpy as np
import re
import json
from pathlib import Path
from urllib.parse import unquote

# ================= 页面抬头 =================
st.title("风机配件分析 1.1.3.5")

# ================= 持久化路径 =================
STORE_PARQUET     = Path("data_store.parquet")
STORE_CSV         = Path("data_store.csv")
MODEL_FLEET_JSON  = Path("model_fleet_counts.json")
SPEC_FLEET_JSON   = Path("spec_fleet_counts.json")
FAN_2COL_JSON     = Path("fan_model_map_2col.json")
# ★ 新增：原始 Excel 库（与泄漏页保持一致）
RAW_EXCEL_STORE   = Path("raw_excel_store.parquet")

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


# ================= 具体型号在保量 · 通用字典 =================
def normalize_model_key(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.upper()

def load_spec_fleet() -> dict:
    if SPEC_FLEET_JSON.exists():
        try:
            return json.loads(SPEC_FLEET_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def save_spec_fleet(d: dict):
    try:
        SPEC_FLEET_JSON.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        return True, ""
    except Exception as e:
        return False, str(e)

if "spec_fleet" not in st.session_state:
    st.session_state.spec_fleet = load_spec_fleet()

def lookup_spec_fleet(model_name: str) -> int:
    key = normalize_model_key(model_name)
    try:
        return int(st.session_state.spec_fleet.get(key, 0))
    except Exception:
        return 0

# ================== 风机型号两列粘贴 · 独立映射（仅风机分析使用） ==================
def _norm_txt2(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u3000", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _norm_key2(s: str) -> str:
    return _norm_txt2(s).upper()

def _parse_left_cell(cell: str) -> tuple[str, str]:
    """
    解析左列：`风机型号  供应商/`
    返回：(风机型号, 供应商)
    """
    s = _norm_txt2(cell)
    m = re.match(r"^(?P<fan>.+?)\s+(?P<vendor>.+?)\/\s*$", s)
    if m:
        return _norm_txt2(m.group("fan")), _norm_txt2(m.group("vendor"))
    m2 = re.match(r"^(?P<fan>.+?)\s+(?P<vendor>[^\/]+)$", s)
    if m2:
        return _norm_txt2(m2.group("fan")), _norm_txt2(m2.group("vendor"))
    return s, ""

def _load_fan2col_rows() -> list[dict]:
    if FAN_2COL_JSON.exists():
        try:
            data = json.loads(FAN_2COL_JSON.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
            if isinstance(data, dict) and "rows" in data:
                return data["rows"]
        except Exception:
            pass
    return []

def _save_fan2col_rows(rows: list[dict]) -> tuple[bool, str]:
    try:
        FAN_2COL_JSON.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
        return True, ""
    except Exception as e:
        return False, str(e)

# —— 会话态与查找函数
if "fan2col_rows" not in st.session_state:
    st.session_state.fan2col_rows = _load_fan2col_rows()

def _build_fan_maps_2col(rows: list[dict]):
    fan2code, fan2vendor = {}, {}
    for r in rows:
        fan = _norm_key2(r.get("fan", ""))
        if not fan:
            continue
        fan2code[fan] = _norm_txt2(r.get("code", ""))
        fan2vendor[fan] = _norm_txt2(r.get("vendor", ""))
    return fan2code, fan2vendor

FAN2CODE, FAN2VENDOR = _build_fan_maps_2col(st.session_state.fan2col_rows)

def fan2code(fan: str) -> str:
    return FAN2CODE.get(_norm_key2(fan), "")

def fan2vendor(fan: str) -> str:
    return FAN2VENDOR.get(_norm_key2(fan), "")
# === 基于映射的带类别/产品查找工具（新增） ===
def _build_index_from_session_rows():
    rows = st.session_state.get("fan2col_rows", []) or []
    fan_index = {}       # { FAN_UPPER: [row, row, ...] }
    product_index = {}   # { PRODUCT_UPPER: [row, row, ...] }
    for r in rows:
        fan_u = _norm_key2(r.get("fan", ""))
        prod_u = normalize_model_key(r.get("product", ""))
        fan_index.setdefault(fan_u, []).append(r)
        if prod_u:
            product_index.setdefault(prod_u, []).append(r)
    return fan_index, product_index

def _choose_row(rows, want_type: str | None):
    """
    从一组映射行里，优先返回 'type' 与 want_type 匹配的那条；
    若找不到匹配类型，则回退到第一条非空 code 的行，再回退任意一条。
    """
    want = (_norm_txt2(want_type) if want_type else "")
    # 1) 精确按类型匹配（例如 '蒸发风机'、'冷凝风机'）
    if want:
        for r in rows:
            if _norm_txt2(r.get("type", "")) == want:
                return r
    # 2) 回退：有 code 的任意行
    for r in rows:
        if _norm_txt2(r.get("code", "")):
            return r
    # 3) 再回退：第一条
    return rows[0] if rows else None

def fan_map_lookup(fan: str = "", *, want_type: str | None = None, product: str | None = None):
    """
    从 fan_model_map_2col.json 的会话态里查：
    - 先用 fan 命中（优先按 want_type 类型），
    - 若 fan 为空或未命中，则用 product 命中（同样优先匹配类型）。
    返回 (display_fan_with_vendor, code)：
        display_fan_with_vendor 形如: 'CYR400P  NIDEC/' 或原样 fan（当无 vendor）
        code 为物料编码（找不到则空字符串）
    """
    fan_index, product_index = _build_index_from_session_rows()

    def _display_fan_vendor(row):
        f = _norm_txt2(row.get("fan", ""))
        v = _norm_txt2(row.get("vendor", ""))
        return f"{f}  {v}/" if f and v else f

    # 1) 先用 fan 命中
    if fan:
        fkey = _norm_key2(fan)
        if fkey in fan_index:
            hit = _choose_row(fan_index[fkey], want_type)
            if hit:
                return _display_fan_vendor(hit), _norm_txt2(hit.get("code", ""))

    # 2) 回退：用 product 命中
    if product:
        pkey = normalize_model_key(product)
        if pkey in product_index:
            hit = _choose_row(product_index[pkey], want_type)
            if hit:
                return _display_fan_vendor(hit), _norm_txt2(hit.get("code", ""))

    # 3) 都没命中：返回输入 fan（不带 vendor），code 为空
    return _norm_txt2(fan), ""


with st.expander("风机型号映射（两列粘贴；左=“风机型号  供应商/”，右=“物料编码”）", expanded=False):
        st.caption("此表仅用于风机分析，不影响其他配件。可直接从两列表格复制粘贴。")
        pasted = st.text_area(
            "粘贴两列文本（每行一条；两列之间用Tab或≥2空格；左列末尾保留“/”更稳）：",
            height=140,
            placeholder="示例：\nCYR400P  NIDEC/    R03001000099\nAC罩机风机  赛微/    R03001000088"
        )

        # === 1) 解析并写入会话态（保持你原有逻辑，并补充 product/type 字段） ===
        if st.button("解析并追加到下方表格", use_container_width=True):
            add_rows = []
            for line in pasted.splitlines():
                line = line.strip()
                if not line:
                    continue
                parts = re.split(r"[\t]{1,}|\s{2,}", line)
                if len(parts) < 2:
                    parts = re.split(r"\s+", line)
                    if len(parts) >= 2:
                        left = " ".join(parts[:-1]); right = parts[-1]
                    else:
                        continue
                else:
                    left, right = parts[0], parts[1]

                fan, vendor = _parse_left_cell(left)
                code = _norm_txt2(right)
                # 新增字段：type/product 默认留空，供下方表格编辑
                add_rows.append({"type": "", "product": "", "fan": fan, "vendor": vendor, "code": code})

            if add_rows:
                cols = ["type","product","fan","vendor","code"]
                cur = pd.DataFrame(st.session_state.fan2col_rows) if st.session_state.fan2col_rows else pd.DataFrame(columns=cols)
                cur = cur.reindex(columns=cols)
                if not cur.empty:
                    cur["K"] = cur["fan"].map(_norm_key2)

                for r in add_rows:
                    k = _norm_key2(r["fan"])
                    if not cur.empty and (cur["K"] == k).any():
                        idx = cur.index[cur["K"] == k][0]
                        # 覆盖供应商与物料编码；保留已存在的 type/product
                        cur.loc[idx, ["vendor","code"]] = [r["vendor"], r["code"]]
                    else:
                        cur = pd.concat([cur, pd.DataFrame([r])], ignore_index=True)

                if "K" in cur.columns:
                    cur = cur.drop(columns=["K"])
                st.session_state.fan2col_rows = cur.fillna("").to_dict(orient="records")
                # 更新查找表
                FAN2CODE, FAN2VENDOR = _build_fan_maps_2col(st.session_state.fan2col_rows)
                st.success(f"已追加/覆盖 {len(add_rows)} 条。")

        # === 2) 三列显示 + 可编辑表格（含“风机类别”） ===
        def _fan_vendor_join(fan: str, vendor: str) -> str:
            fan = _norm_txt2(fan); vendor = _norm_txt2(vendor)
            return f"{fan}  {vendor}/" if vendor else fan

        # 兼容老数据：没有 type/product 字段时补空
        _rows = st.session_state.fan2col_rows or []
        for r in _rows:
            if "type" not in r: r["type"] = ""
            if "product" not in r: r["product"] = ""

        view2 = (
            pd.DataFrame(_rows)
            .reindex(columns=["type","product","fan","vendor","code"])
            .rename(columns={
                "type":"风机类别",
                "product":"型号（产品型号）",
                "fan":"_fan",
                "vendor":"_vendor",
                "code":"物料编码"
            })
        )
        if not view2.empty:
            view2.insert(2, "风机型号+供应商", view2.apply(lambda r: _fan_vendor_join(r["_fan"], r["_vendor"]), axis=1))
            view2 = view2[["风机类别","型号（产品型号）","风机型号+供应商","物料编码"]].sort_values(
                ["风机类别","型号（产品型号）","风机型号+供应商"], ignore_index=True
            )
        else:
            view2 = pd.DataFrame(columns=["风机类别","型号（产品型号）","风机型号+供应商","物料编码"])

        edited2 = st.data_editor(
            view2,
            num_rows="dynamic",
            use_container_width=True,
            key="fan2col_editor",
            column_config={
                "风机类别": st.column_config.SelectboxColumn(
                    options=["", "蒸发风机", "冷凝风机"],
                    help="请选择风机类别；也可留空稍后再区分"
                ),
                "型号（产品型号）": st.column_config.TextColumn(help="整机产品型号，例如 AOM-48 等，可留空"),
                "风机型号+供应商": st.column_config.TextColumn(help="示例：CYR400P  NIDEC/ 或 AC罩机风机  赛微/"),
                "物料编码": st.column_config.TextColumn(help="示例：R03001000099"),
            }
        )

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("保存到本地（fan_model_map_2col.json）", use_container_width=True):
                rows = []
                for _, r in edited2.iterrows():
                    t = _norm_txt2(r.get("风机类别",""))
                    product = _norm_txt2(r.get("型号（产品型号）",""))
                    fv = _norm_txt2(r.get("风机型号+供应商",""))
                    code = _norm_txt2(r.get("物料编码",""))
                    if not fv and not code and not product and not t:
                        continue
                    fan, vendor = _parse_left_cell(fv or "")
                    rows.append({"type": t, "product": product, "fan": fan, "vendor": vendor, "code": code})

                ok, msg = _save_fan2col_rows(rows)
                if ok:
                    st.session_state.fan2col_rows = rows
                    FAN2CODE, FAN2VENDOR = _build_fan_maps_2col(rows)
                    st.success(f"已保存 {len(rows)} 条。")
                else:
                    st.error(f"保存失败：{msg}")

        with c2:
            if not edited2.empty:
                st.download_button(
                    "导出当前表为 CSV",
                    data=edited2.to_csv(index=False, encoding="utf-8-sig"),
                    file_name="fan_model_map_2col.csv",
                    mime="text/csv",
                    use_container_width=True
                )

        with c3:
            if st.button("清空（仅内存，不落盘）", use_container_width=True):
                st.session_state.fan2col_rows = []
                FAN2CODE, FAN2VENDOR = _build_fan_maps_2col([])
                st.rerun()
# ================= 两列粘贴映射 · 结束 =================
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
    if mapping:
        df = df.rename(columns=mapping)
    for col, default in [
        ("Category1",""),("Category2",""),("TAG1",""),("TAG2",""),("TAG3",""),("TAG4",""),
        ("CostUSD",0.0),("PaidQty",0),("IsWarranty",False),("Completed",False)
    ]:
        if col not in df.columns:
            df[col] = default
    df["CostUSD"] = pd.to_numeric(df["CostUSD"], errors="coerce").fillna(0.0)
    df["PaidQty"] = pd.to_numeric(df["PaidQty"], errors="coerce").fillna(0).astype(int)
    if "Year" in df.columns:
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce")
    if "Month" in df.columns:
        df["Month"] = pd.to_numeric(df["Month"], errors="coerce")
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
    if v is None:
        return default
    if cast is None:
        return v
    try:
        return cast(v)
    except Exception:
        return default

def get_param_str(name: str, default=None):
    try:
        v = st.query_params.get(name, default)
    except Exception:
        qp = st.experimental_get_query_params()
        v = (qp.get(name, [default]) or [default])[0]
    if isinstance(v, str):
        return unquote(v)
    return default

def load_store_df_from_disk() -> pd.DataFrame:
    if STORE_PARQUET.exists():
        try:
            return pd.read_parquet(STORE_PARQUET)
        except Exception:
            pass
    if STORE_CSV.exists():
        try:
            return pd.read_csv(STORE_CSV)
        except Exception:
            pass
    return pd.DataFrame()

def get_store_df() -> pd.DataFrame:
    if "store_df" in st.session_state and isinstance(st.session_state["store_df"], pd.DataFrame):
        return ensure_cols(st.session_state["store_df"].copy())
    return ensure_cols(load_store_df_from_disk())

def money(v) -> str:
    try:
        return f"${float(v):,.0f}"
    except Exception:
        return "$0"
# === 百分比统一显示为 0.00% ===
def _fmt_pct(v, digits=2):
    try:
        if v is None or (isinstance(v, float) and np.isnan(v)): 
            return ""
        return f"{float(v):.{digits}f}%"
    except Exception:
        return ""

def format_percent_cols(df: pd.DataFrame, cols: list[str], digits: int = 2) -> pd.DataFrame:
    if df is None or df.empty: 
        return df
    df2 = df.copy()
    for c in cols:
        if c in df2.columns:
            df2[c] = df2[c].map(lambda x: _fmt_pct(x, digits))
    return df2

# ================= 参数与数据读取 =================
code  = get_param_str("code", default="1.1.3.5")
year  = get_param("year", default=None, cast=int)
month = get_param("month", default=None, cast=int)

# 新增：支持多月/季度
months_param = get_param_str("months", default=None)   # 形如 "1,2,3"
q_param      = get_param_str("q", default=None)        # 形如 "Q1"

df_all = get_store_df()
if df_all.empty:
    st.warning("未获取到任何数据：请先在“冰箱数据”页面上传数据；或确认 data_store.parquet / data_store.csv 是否存在。")
    st.stop()

yy = pd.to_numeric(df_all.get("Year"), errors="coerce")
mm = pd.to_numeric(df_all.get("Month"), errors="coerce")

# —— 年份：优先参数；否则取数据中最大年
if year is None:
    if yy.notna().any():
        year = int(yy.dropna().max())
    else:
        year = 2025

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
if set(sel_months) == {1,2,3}:
    scope_text = f"{year} Q1"
elif set(sel_months) == {4,5,6}:
    scope_text = f"{year} Q2"
elif set(sel_months) == {7,8,9}:
    scope_text = f"{year} Q3"
elif set(sel_months) == {10,11,12}:
    scope_text = f"{year} Q4"
elif len(sel_months) == 1:
    scope_text = f"{year}-{int(sel_months[0]):02d}"
else:
    scope_text = f"{year} 月份：{','.join(str(m) for m in sel_months)}"

if scope_df.empty:
    st.info(f"{scope_text} 没有可统计数据。")
    st.stop()


joined_all = scope_df.astype(str).agg(" ".join, axis=1).map(_clean_text)
pat_code = re.compile(rf"(?<![0-9a-z]){re.escape(code)}(?![0-9a-z])", re.I)
mask_code = joined_all.str.contains(pat_code)
df_code   = scope_df[mask_code].copy()

st.caption(f"统计口径：{scope_text} | 缺陷码：{code}")
if df_code.empty:
    st.warning("该口径下未命中任何 1.1.3.5 相关记录。")
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

# ================= 1.1.3.5.x（蒸发/冷凝）分项 =================
def build_1135_subtable(df_in: pd.DataFrame):
    join = df_in.astype(str).agg(" ".join, axis=1).map(_clean_text)
    pats = {
        "1.1.3.5.1 Evap Fan Motor": re.compile(r"(?<![0-9a-z])1\.1\.3\.5\.1(?![0-9a-z])", re.I),
        "1.1.3.5.2 Condenser Fan Motor": re.compile(r"(?<![0-9a-z])1\.1\.3\.5\.2(?![0-9a-z])", re.I),
    }
    rows = []
    for name, pat in pats.items():
        m = join.str.contains(pat)
        qty = int(m.sum())
        fee = float(pd.to_numeric(df_in.loc[m, "CostUSD"], errors="coerce").fillna(0).sum())
        if qty > 0 or fee > 0:
            cn = "蒸发风机" if name.endswith("5.1 Evap Fan Motor") else "冷凝风机"
            rows.append({"英文名称": name, "质量问题": cn, "数量": qty, "费用(USD)": fee})
    if not rows:
        return pd.DataFrame()
    df_out = pd.DataFrame(rows)
    total_row = {"英文名称":"总计","质量问题":"","数量":int(df_out["数量"].sum()),"费用(USD)":float(df_out["费用(USD)"].sum())}
    df_out = pd.concat([df_out, pd.DataFrame([total_row])], ignore_index=True)
    return df_out

df_1135 = build_1135_subtable(df_code)
st.markdown("### 1.1.3.5.x 分项（蒸发 / 冷凝）")
if not df_1135.empty:
    st.dataframe(df_1135, use_container_width=True)
    st.download_button(
        "下载 CSV（1.1.3.5 分项）",
        data=df_1135.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"fan_1135_breakdown_{scope_text.replace(' ','_')}.csv",
        mime="text/csv",
        use_container_width=True
    )
else:
    st.info("暂无有效数据可显示。")

# ========== 三元结构：具体型号在保量 · 管理（3列：型号A | 型号B | 在保量） ==========
SPEC_FLEET_JSON = Path("spec_fleet_counts.json")  # 仍沿用这个文件名

def _load_spec_fleet_triples() -> list[dict]:
    if SPEC_FLEET_JSON.exists():
        try:
            data = json.loads(SPEC_FLEET_JSON.read_text(encoding="utf-8"))
            if isinstance(data, dict) and "triples" not in data:
                triples = [{"a": k, "b": "", "fleet": v} for k, v in data.items()]
                return triples
            if isinstance(data, dict) and "triples" in data:
                return data["triples"]
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []

def _save_spec_fleet_triples(triples: list[dict]) -> tuple[bool, str]:
    try:
        payload = {"triples": triples}
        SPEC_FLEET_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return True, ""
    except Exception as e:
        return False, str(e)

def _rebuild_spec_fleet_map(triples: list[dict]) -> dict:
    mp = {}
    for r in triples:
        a = normalize_model_key(r.get("a", ""))
        b = normalize_model_key(r.get("b", ""))
        fleet = int(pd.to_numeric(r.get("fleet", 0), errors="coerce") or 0)
        if a: mp[a] = fleet
        if b: mp[b] = fleet
    return mp

# —— 会话态：triples + 扁平映射
if "spec_fleet_triples" not in st.session_state:
    st.session_state.spec_fleet_triples = _load_spec_fleet_triples()
st.session_state.spec_fleet_map = _rebuild_spec_fleet_map(st.session_state.spec_fleet_triples)

def lookup_spec_fleet(model_name: str) -> int:
    key = normalize_model_key(model_name)
    try:
        return int(st.session_state.spec_fleet_map.get(key, 0))
    except Exception:
        return 0

with st.expander("具体型号在保量 · 管理（三列表，通用）", expanded=False):
    st.caption("使用三列表结构（型号A、型号B、在保量）。保存后会写入 spec_fleet_counts.json，且自动生成扁平映射供统计使用。")

    up = st.file_uploader("上传 CSV/Excel（列名示例：型号A, 型号B, 在保量 或 具体型号, 在保量）", type=["csv","xlsx","xls"])

    def _norm(s: str) -> str:
        s = str(s).replace("\u3000", " ").strip().lower()
        return re.sub(r"\s+", "", s)

    if up is not None:
        try:
            if up.name.lower().endswith(".csv"):
                tmp = pd.read_csv(up)
            else:
                tmp = pd.read_excel(up)

            cand_a = {"型号a","型号","具体型号","model","机型","机型明细","part","partmodel","part model","model a","型号1"}
            cand_b = {"型号b","对应型号","映射型号","标准型号","替代型号","别名","model b","型号2"}
            cand_f = {"在保量","在保","保有量","装机量","保有量(台)","在保量(台)","installed","fleet","qty"}

            def _pick(cands):
                for c in tmp.columns:
                    if _norm(c) in {_norm(x) for x in cands}:
                        return c
                return None

            col_a = _pick(cand_a)
            col_b = _pick(cand_b)
            col_f = _pick(cand_f)

            missing = []
            if not col_a: missing.append("型号A/具体型号")
            if not col_f: missing.append("在保量")
            if missing:
                st.error("缺少必须列：" + "、".join(missing) + f"。当前列：{list(tmp.columns)}")
            else:
                if col_b:
                    tmp = tmp[[col_a, col_b, col_f]].copy()
                    tmp.columns = ["型号A", "型号B", "在保量"]
                else:
                    tmp = tmp[[col_a, col_f]].copy()
                    tmp.columns = ["型号A", "在保量"]
                    tmp["型号B"] = ""

                tmp["型号A"] = tmp["型号A"].astype(str).map(normalize_model_key)
                tmp["型号B"] = tmp["型号B"].astype(str).map(normalize_model_key)
                tmp["在保量"] = pd.to_numeric(tmp["在保量"], errors="coerce").fillna(0).astype(int)

                merged = 0
                cur = pd.DataFrame(st.session_state.spec_fleet_triples) if st.session_state.spec_fleet_triples else pd.DataFrame(columns=["a","b","fleet"])
                if not cur.empty:
                    cur["a"] = cur["a"].astype(str).map(normalize_model_key)
                    cur["b"] = cur["b"].astype(str).map(normalize_model_key)
                    cur["fleet"] = pd.to_numeric(cur["fleet"], errors="coerce").fillna(0).astype(int)
                for _, r in tmp.iterrows():
                    a, b, f = r["型号A"], r["型号B"], int(r["在保量"])
                    hit = (not cur.empty) and (
                        ((cur["a"]==a) & (cur["b"]==b)).any() or
                        ((cur["a"]==a) & (cur["b"]=="") & (b=="")).any()
                    )
                    if hit:
                        idx = cur[((cur["a"]==a) & (cur["b"]==b)) | ((cur["a"]==a) & (cur["b"]=="") & (b==""))].index
                        cur.loc[idx, "fleet"] = f
                    else:
                        cur = pd.concat([cur, pd.DataFrame([{"a":a, "b":b, "fleet":f}])], ignore_index=True)
                    merged += 1

                st.session_state.spec_fleet_triples = cur[["a","b","fleet"]].to_dict(orient="records")
                st.session_state.spec_fleet_map = _rebuild_spec_fleet_map(st.session_state.spec_fleet_triples)
                st.success(f"已合并 {merged} 条记录。")

        except Exception as e:
            st.error(f"解析失败：{e}")
        
    view_df = (
        pd.DataFrame(st.session_state.spec_fleet_triples)
          .rename(columns={"a":"型号A","b":"型号B","fleet":"在保量"})
          .sort_values(["型号A","型号B"], ignore_index=True)
        if st.session_state.spec_fleet_triples else
        pd.DataFrame(columns=["型号A","型号B","在保量"])
    )

    edited = st.data_editor(
        view_df,
        num_rows="dynamic",
        use_container_width=True,
        key="spec_fleet_editor_triples",
        column_config={
            "型号A": st.column_config.TextColumn(help="主型号；保存时自动规范化（大写、去多余空格）"),
            "型号B": st.column_config.TextColumn(help="别名/对应型号，可留空"),
            "在保量": st.column_config.NumberColumn(min_value=0, step=1),
        }
    )

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        if st.button("保存到本地（spec_fleet_counts.json）", use_container_width=True):
            triples = []
            for _, r in edited.iterrows():
                a = normalize_model_key(r.get("型号A",""))
                b = normalize_model_key(r.get("型号B",""))
                f = int(pd.to_numeric(r.get("在保量", 0), errors="coerce") or 0)
                if not a and not b:
                    continue
                triples.append({"a": a, "b": b, "fleet": f})

            ok, msg = _save_spec_fleet_triples(triples)
            if ok:
                st.session_state.spec_fleet_triples = triples
                st.session_state.spec_fleet_map = _rebuild_spec_fleet_map(triples)
                st.success(f"已保存 {len(triples)} 条（并已刷新用于统计的映射）。")
            else:
                st.error(f"保存失败：{msg}")

    with col_s2:
        if st.button("清空当前编辑（仅内存，不落盘）", use_container_width=True):
            st.session_state.spec_fleet_triples = []
            st.session_state.spec_fleet_map = {}
            st.rerun()
# ==================== 三元结构管理 结束 ====================

# ======================= 三张统计表（全部可排序） =======================

# —— 字段名候选
USE_MONTH_CAND = [
    "已使用保修时长","已使用保修月数","保修已使用时长","已使用月份","保修使用月数",
    "使用月数","使用月","ServiceMonths","months_in_use","use_months"
]
SERIES_PRI_CAND = ["型号归类","型号系列","series","model_series"]
SERIES_SEC_CAND = ["机型","产品型号","MODEL","Model"]
SPEC_MODEL_CAND = ["型号","具体型号","机型明细","MODEL","Model","Part Model","PartModel","Part"]
FAN_KIND_CAND   = ["风机型号","风机型号","FanType","风机类型"]
PART_CODE_CAND  = ["R码","物料编码","物料号","PartCode","Part No"]

# —— 机型归一
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
        if key_up == k.upper().replace(" ",""):
            return v
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
        try:
            return json.loads(MODEL_FLEET_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def get_fleet_month(fleet: dict, year: int=None, month: int=None)->dict:
    if not fleet:
        return {k:0 for k in MODEL_CATEGORY_MAP.keys()}
    if year is None or month is None:
        try:
            yy = max(int(y) for y in fleet.keys())
            mm = max(int(m) for m in fleet[str(yy)].keys())
            year, month = yy, mm
        except Exception:
            return {k:0 for k in MODEL_CATEGORY_MAP.keys()}
    y, m = str(int(year)), str(int(month))
    return {k:int(fleet.get(y,{}).get(m,{}).get(k,0)) for k in MODEL_CATEGORY_MAP.keys()}

# —— 月份分段（仅 >0 计入；空/0 不入桶）
def _bucket_series(months: pd.Series):
    m = pd.to_numeric(months, errors="coerce").where(lambda x: x > 0)
    return pd.cut(m, bins=[0, 4, 8, np.inf],
                  labels=["短期（1-4月）","中期（5-8月）","长期（9-24月）"],
                  include_lowest=True, right=True)

# —— 三表构造
def build_tables(df_subset: pd.DataFrame, *, year=None, month=None, fan_type: str | None = None):
        months_col = _pick_col(df_subset, USE_MONTH_CAND)
        series_col = _pick_col(df_subset, SERIES_PRI_CAND) or _pick_col(df_subset, SERIES_SEC_CAND)
        spec_col   = _pick_col(df_subset, SPEC_MODEL_CAND)
        ftype_col  = _pick_col(df_subset, FAN_KIND_CAND)
        pcode_col  = _pick_col(df_subset, PART_CODE_CAND)

        if months_col is None:
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        if series_col is None and spec_col is not None:
            tmp = df_subset[spec_col].astype(str).str.upper().str.extract(r"^([A-Z]{2,4})")
            df_subset = df_subset.copy()
            df_subset["__SERIES__"] = tmp[0].map(lambda x: ALIASES.get((x or "").replace(" ",""), x or ""))
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
            g = pd.pivot_table(df_valid, index=series_col, columns="__bucket__", values="Date",
                               aggfunc="count", fill_value=0)
            for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）"]:
                if c not in g.columns: g[c] = 0
            g = g[["短期（1-4月）","中期（5-8月）","长期（9-24月）"]]
            g["总计"] = g.sum(axis=1)
            g.reset_index(inplace=True)
            g.rename(columns={series_col:"型号"}, inplace=True)

            fleet = load_model_fleet()
            month_fleet = get_fleet_month(fleet, year, month)
            def _fleet_lookup(s):
                key = resolve_model_key(s)
                return int(month_fleet.get(key, 0))
            g["在保量"] = g["型号"].map(_fleet_lookup).fillna(0).astype(int)
            g["问题比例(%)"] = np.where(g["在保量"]>0, g["总计"]/g["在保量"]*100.0, np.nan)

            order_map = {k:i for i,k in enumerate(MODEL_ORDER)}
            g["__ord__"] = g["型号"].map(order_map).fillna(9999)
            g = g.sort_values(["__ord__","总计"], ascending=[True,False]).drop(columns="__ord__")
            tbl1 = g

        # 表2：按具体型号（确保“长期/风机型号/型号(物料编码)”列，且优先使用映射覆盖）
        tbl2 = pd.DataFrame()
        if spec_col is not None:
            g2 = pd.pivot_table(
                df_valid, index=spec_col, columns="__bucket__", values="Date",
                aggfunc="count", fill_value=0
            )
            for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）"]:
                if c not in g2.columns: g2[c] = 0
            g2 = g2[["短期（1-4月）","中期（5-8月）","长期（9-24月）"]]

            g2["总计"] = g2.sum(axis=1)
            g2.reset_index(inplace=True)
            g2.rename(columns={spec_col: "具体型号"}, inplace=True)

            # 源数据的“风机型号”众数（若有）
            def _mode(series):
                s = series.dropna().astype(str)
                if s.empty: return ""
                m0 = s.mode()
                return m0.iloc[0] if not m0.empty else s.iloc[0]
            if ftype_col:
                base_fan = (
                    df_subset.groupby(spec_col)[ftype_col].apply(_mode)
                    .reindex(g2["具体型号"]).fillna("").values
                )
            else:
                base_fan = [""] * len(g2)

             # 逐行用“映射”覆盖/补齐：先按风机型号匹配；若风机型号为空再按产品型号匹配；两者都按 fan_type 优先
            out_fan_display = []
            out_code        = []
            for i, row in g2.iterrows():
                spec_model = _norm_txt2(row["具体型号"])
                fan_guess  = _norm_txt2(base_fan[i]) if isinstance(base_fan, (list, np.ndarray)) else _norm_txt2(base_fan)
                display_fan, code_from_map = fan_map_lookup(
                    fan=fan_guess,
                    want_type=fan_type,
                    product=spec_model
                )
                # 显示：风机型号一律显示“型号  供应商/”（若映射无 vendor 则只显示 fan）
                out_fan_display.append(display_fan if display_fan else fan_guess)

                # 重要：这里的“型号”=“风机物料编码”，只取映射。没有就留空，避免误用源数据里的“型号”(具体机型)
                out_code.append(code_from_map)

            g2["风机型号"] = out_fan_display
            g2["型号"]     = out_code

            # 物料编码（优先用映射）


            # 在保量与比例
            g2["在保量"] = g2["具体型号"].map(lookup_spec_fleet).fillna(0).astype(int)
            g2["比例(%)"] = np.where(g2["在保量"]>0, g2["总计"]/g2["在保量"]*100.0, np.nan)

            g2 = g2.sort_values("总计", ascending=False)

            # 防御性补列（数值/文本）
            for c in ["短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量"]:
                if c not in g2.columns: g2[c] = 0
            for c in ["风机型号","型号"]:
                if c not in g2.columns: g2[c] = ""

            tbl2 = g2

                      # ===== 表3：由“按具体型号”汇总 -> 温控型号（即风机型号） =====
        tbl3 = pd.DataFrame()
        if spec_col is not None:
            # 这里使用上面刚算好的 g2（按具体型号），其中的「风机型号」就是我们要显示的温控型号
            if not g2.empty:
                tmp = g2.copy()
                # 用“(未识别)”占位，避免空字符串分组丢失
                tmp["风机型号"] = tmp["风机型号"].replace("", np.nan).fillna("(未识别)")

                # 分组汇总（短/中/长期、总计、在保量）
                sum_cols = ["短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量"]
                for c in sum_cols:
                    if c not in tmp.columns:
                        tmp[c] = 0
                g3 = tmp.groupby("风机型号", as_index=False)[sum_cols].sum()

                # 问题比例 = 总计 / 在保量 * 100（无在保量则 NaN）
                g3["问题比例(%)"] = np.where(g3["在保量"] > 0, g3["总计"] / g3["在保量"] * 100.0, np.nan)

                # 排序
                g3 = g3.sort_values("总计", ascending=False)

                tbl3 = g3

        return tbl1, tbl2, tbl3


# —— 子集：蒸发 / 冷凝
def _subset_by_code(df_src: pd.DataFrame, subcode: str):
    j = df_src.astype(str).agg(" ".join, axis=1).map(_clean_text)
    pat = re.compile(rf"(?<![0-9a-z]){re.escape(subcode)}(?![0-9a-z])", re.I)
    return df_src[j.str.contains(pat)]

df_evap = _subset_by_code(df_code, "1.1.3.5.1")
df_cond = _subset_by_code(df_code, "1.1.3.5.2")

# —— 通用渲染：可排序表格
def render_sortable_table(title: str, df: pd.DataFrame, *, order_cols):
    cols = [c for c in order_cols if c in df.columns]
    view = df.loc[:, cols].copy() if cols else df.copy()

    # —— 百分比列内联格式化为 "xx.xx%"
    for pct_col in ["问题比例(%)", "比例(%)"]:
        if pct_col in view.columns:
            view[pct_col] = view[pct_col].map(lambda x: "" if pd.isna(x) else f"{float(x):.2f}%")

    st.subheader(title)
    st.caption("提示：点击列头可排序；再次点击切换升/降序；按住 Shift 支持多列排序。")
    st.dataframe(view, use_container_width=True)


def render_sortable_table_collapsed(title: str, df: pd.DataFrame, *, order_cols, expanded=False):
    cols = [c for c in order_cols if c in df.columns]
    view = df.loc[:, cols].copy() if cols else df.copy()

    # —— 百分比列内联格式化为 "xx.xx%"
    for pct_col in ["问题比例(%)", "比例(%)"]:
        if pct_col in view.columns:
            view[pct_col] = view[pct_col].map(lambda x: "" if pd.isna(x) else f"{float(x):.2f}%")

    with st.expander(title, expanded=expanded):
        st.caption("提示：点击列头可排序；再次点击切换升/降序；按住 Shift 支持多列排序。")
        st.dataframe(view, use_container_width=True)


# ===== 蒸发风机（1.1.3.5.1）
t1, t2, t3 = build_tables(df_evap, year=year, month=month_for_fleet, fan_type="蒸发风机")
base1 = pd.DataFrame(columns=["型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])
base2 = pd.DataFrame(columns=["具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","风机型号","型号","总计","在保量","比例(%)"])
base3 = pd.DataFrame(columns=["风机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"])

render_sortable_table("蒸发风机（按系列）", (t1 if not t1.empty else base1),
                      order_cols=("型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))

render_sortable_table_collapsed("蒸发风机（按具体型号）", (t2 if not t2.empty else base2),
                                order_cols=("具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","风机型号","型号","总计","在保量","比例(%)"),
                                expanded=True)

render_sortable_table("蒸发风机（按厂家型号）", (t3 if not t3.empty else base3),
                      order_cols=("风机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))

# ===== 冷凝风机（1.1.3.5.2）
c1, c2, c3 = build_tables(df_cond, year=year, month=month_for_fleet, fan_type="冷凝风机")

render_sortable_table("冷凝风机（按系列）", (c1 if not c1.empty else base1),
                      order_cols=("型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))

render_sortable_table_collapsed("冷凝风机（按具体型号）", (c2 if not c2.empty else base2),
                                order_cols=("具体型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","风机型号","型号","总计","在保量","比例(%)"),
                                expanded=True)

render_sortable_table("冷凝风机（按厂家型号）", (c3 if not c3.empty else base3),
                      order_cols=("风机型号","短期（1-4月）","中期（5-8月）","长期（9-24月）","总计","在保量","问题比例(%)"))

# ================= 原始明细（命中 1.1.3.5 的记录） =================
with st.expander("原始明细（命中 1.1.3.5 的记录）", expanded=False):
    # —— 统一把“231221 / 202405 / 2024-5-6 …”转成 YYYY-MM-DD
    def _norm_date_text(raw: pd.Series) -> pd.Series:
        s = raw.astype(str).str.strip()
        out = pd.Series("", index=s.index, dtype="object")

        # 8位：YYYYMMDD
        m8 = s.str.fullmatch(r"\d{8}", na=False)
        if m8.any():
            dt8 = pd.to_datetime(s[m8], format="%Y%m%d", errors="coerce")
            out.loc[m8] = dt8.dt.strftime("%Y-%m-%d")

        # 6位：优先 YYMMDD（如 231221）；仅 (19|20)YYYYMM（如 202405）按 YYYYMM→补01日
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

    # 以当前命中记录为基底（如需，可先 df_base = df_code.copy() 再用 df_base）
    df_view = df_code.copy()

    # —— 与“原始 Excel 库”关联，补出【原始生产日期】【原始序列号】（逻辑与泄漏页一致）
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
            # 只取右表存在的键
            take_cols = [k for k in join_keys]
            if raw_prod_col   and raw_prod_col   not in take_cols: take_cols.append(raw_prod_col)
            if raw_serial_col and raw_serial_col not in take_cols: take_cols.append(raw_serial_col)
            patch = raw_hist.loc[:, take_cols].drop_duplicates().copy()

            # ★ 不要改 join key 的列名；仅复制出“原始*”列
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
            st.caption("（原始库可用，但未找到共同主键或原始“生产日期/序列号”列，显示现有列）")
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
        file_name="fan_1135_raw_details.csv",
        mime="text/csv",
        use_container_width=True
    )
