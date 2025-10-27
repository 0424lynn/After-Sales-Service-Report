# -*- coding: utf-8 -*-
# === 制冰机数据（空数据版，模板与冰箱数据一致；存储文件独立 *_ice，确保起始即为空） ===

import streamlit as st
from auth import require_login, current_user, is_admin, logout_button

# === 临时免登录（调试用）==========================================
import os, re, sys, json, hashlib, pathlib
from pathlib import Path
from urllib.parse import quote, unquote
from typing import List

import pandas as pd
import numpy as np
import streamlit.components.v1 as _components
from urllib.parse import quote as _urlquote

def _get_debug_flag_from_url() -> bool:
    try:
        qp = st.query_params
        val = qp.get("debug", "0")
        if isinstance(val, list):
            val = (val[0] if val else "0")
    except Exception:
        try:
            qp = st.experimental_get_query_params()
            val = (qp.get("debug", ["0"]) or ["0"])[0]
        except Exception:
            val = "0"
    return str(val).strip() in {"1", "true", "yes", "on"}

DEBUG_NO_LOGIN = (
    os.getenv("DEBUG_NO_LOGIN", "0").lower() in {"1", "true", "yes", "on"}
    or _get_debug_flag_from_url()
)

if DEBUG_NO_LOGIN:
    def _debug_user():
        return {"username": "debug", "name": "调试模式", "role": "admin"}

    require_login = (lambda *a, **k: None)
    current_user  = _debug_user
    is_admin      = (lambda: True)
    logout_button = (lambda *a, **k: None)
    try:
        st.toast("🔓 已启用『调试免登录』，上线前请关闭。")
    except Exception:
        pass
# ================================================================

# === BEGIN: 彻底从侧栏移除“分析/配件分析/ICE配件页”（入口页口径） ===
def _get_root_script_path():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        ctx = get_script_run_ctx()
        if ctx:
            for attr in ("main_script_path", "script_path"):
                p = getattr(ctx, attr, None)
                if p:
                    return p
    except Exception:
        pass
    return os.path.abspath(sys.argv[0])

def remove_analysis_pages_from_sidebar(*, extra_basenames=(), name_regex=r".*分析$"):
    """必须用入口脚本的 get_pages(root) 才能从源头移除侧栏项"""
    try:
        from streamlit.source_util import get_pages
    except Exception:
        return
    root_script = _get_root_script_path()
    pages = get_pages(root_script) or {}
    if not pages:
        return

    rx = re.compile(name_regex, flags=re.I) if name_regex else None
    need_hide = {
        # —— 旧页面（冰箱分析）
        "配件分析.py", "peijianfenxi.py",
        "风机分析.py", "温控器分析.py", "压缩机分析.py",
        "外部漏水分析.py", "泄氟分析.py",

        # —— ICE 专用页面（中文文件名）
        "ice温控板.py", "ice压缩机.py", "ice调式厚冰.py",
        "ice化霜探头.py", "ice水泵.py", "ice水位探头.py", "ice蒸发器损坏.py",

        # —— ICE 专用页面（英文或驼峰）
        "ice_controlboard.py", "ice_compressor.py", "ice_icethicknessdebug.py",
        "ice_harvestprobe.py", "ice_waterpump.py", "ice_waterlevelprobe.py", "ice_evaporator.py",
        "ICE_ControlBoard.py", "ICE_Compressor.py", "ICE_IceThicknessDebug.py",
        "ICE_HarvestProbe.py", "ICE_WaterPump.py", "ICE_WaterLevelProbe.py", "ICE_Evaporator.py",
        "ICE温控板.py", "ICE压缩机.py", "ICE调式厚冰.py",
        "ICE化霜探头.py", "ICE水泵.py", "ICE水位探头.py", "ICE蒸发器损坏.py",
    }
    need_hide |= {str(x).strip().lower() for x in extra_basenames if str(x).strip()}

    for k, d in list(pages.items()):
        page_name = str(d.get("page_name", "")).strip()
        base = os.path.basename(str(d.get("script_path", "")).replace("\\", "/")).lower()
        if base in need_hide or (rx and rx.fullmatch(page_name)):
            pages.pop(k, None)

# —— 立刻调用（必须在任何 set_page_config 之前）
remove_analysis_pages_from_sidebar(
    extra_basenames=(
        "配件分析.py", "peijianfenxi.py",
        "风机分析.py", "温控器分析.py", "压缩机分析.py",
        "外部漏水分析.py", "泄氟分析.py",
        "ICE温控板.py", "ICE压缩机.py", "ICE调式厚冰.py",
        "ICE化霜探头.py", "ICE水泵.py", "ICE水位探头.py", "ICE蒸发器损坏.py",
        "ICE_ControlBoard.py", "ICE_Compressor.py", "ICE_IceThicknessDebug.py",
        "ICE_HarvestProbe.py", "ICE_WaterPump.py", "ICE_WaterLevelProbe.py", "ICE_Evaporator.py",
    ),
    name_regex=r".*分析$",
)
# === END ===

# ============== 页面配置（改名为“制冰机数据”） ==============
st.set_page_config(page_title="制冰机数据", layout="wide", initial_sidebar_state="expanded")
st.title("制冰机数据")
require_login()

# ============== 跳转辅助（入口页口径，修复 __file__ 问题） ==============
def _find_page_url(target_tail: str):
    """用入口脚本页表查目标文件对应的 URL；找不到就返回 None"""
    try:
        from streamlit.source_util import get_pages
    except Exception:
        return None
    try:
        pages = get_pages(_get_root_script_path()) or {}
        target_tail_norm = re.sub(r"[\\/]+", "/", str(target_tail)).lower()
        for _, data in pages.items():
            sp = data.get("script_path", "")
            sp_norm = re.sub(r"[\\/]+", "/", str(sp)).lower()
            if sp_norm.endswith(target_tail_norm) or os.path.basename(sp_norm) == os.path.basename(target_tail_norm):
                url = data.get("url_pathname")
                if url:
                    return url
                page_name = data.get("page_name")
                if page_name:
                    return "/" + str(page_name)
        return None
    except Exception:
        return None

# —— 质量缺陷编码 → 子页文件名（同时兼容中英文文件名）
_FORCE_CODE2PAGE = {
    "2.1.4.7":  "ICE温控板.py",
    "2.1.4.5":  "ICE压缩机.py",
    "2.1.5.1":  "ICE调式厚冰.py",
    "2.1.4.8":  "ICE化霜探头.py",
    "2.1.4.4":  "ICE水泵.py",
    "2.1.4.9":  "ICE水位探头.py",
    "2.1.4.17": "ICE蒸发器损坏.py",
    # 可选英文化（若你的文件名改成英文）
    # "2.1.4.7":  "ICE_ControlBoard.py",
    # ...
}

def _guess_url_from_filename(target: str) -> str:
    base = os.path.splitext(os.path.basename(str(target)))[0]
    base = re.sub(r"^\d+[_\-\s]*", "", base)
    return f"/{base}"

def page_url_by_code(caller_file: str, code: str) -> str:
    key = str(code or "").strip()
    target = _FORCE_CODE2PAGE.get(key)
    if not target:
        return None
    url = _find_page_url(target)
    if url:
        return url
    return _guess_url_from_filename(target)

def hide_analysis_pages_in_sidebar(caller_file: str, *, debug=False):
    ALLOW_NAMES = {"制冰机数据"}  # 允许“制冰机数据”在侧栏显示
    HIDE_NAME_PATTERNS = [
        r"分析$", r"分析页面", r"配件分析",
        r"风机分析", r"温控器分析", r"压缩机分析",
        r"外部漏水分析", r"泄氟分析",
        r"analysis$"
    ]
    HIDE_URL_TAILS = [
        # —— 旧分析页
        "/配件分析", "/风机分析", "/温控器分析", "/压缩机分析", "/外部漏水分析", "/泄氟分析",
        "/peijianfenxi", "/fengjifenxi", "/wenkongqifenxi", "/yasuoji", "/waibuloushui", "/xiefen",
        # —— ICE 专用页面（中文/英文）
        "/ICE温控板", "/ICE压缩机", "/ICE调式厚冰",
        "/ICE化霜探头", "/ICE水泵", "/ICE水位探头", "/ICE蒸发器损坏",
        "/ICE_ControlBoard", "/ICE_Compressor", "/ICE_IceThicknessDebug",
        "/ICE_HarvestProbe", "/ICE_WaterPump", "/ICE_WaterLevelProbe", "/ICE_Evaporator",
    ]
    enc_tails = []
    for t in HIDE_URL_TAILS:
        enc = "/".join(_urlquote(p, safe="") for p in t.split("/"))
        enc_tails.extend({t, enc, t+"/", enc+"/"})
    HREFS = list(sorted(set(enc_tails)))
    payload = {"allow_names": sorted(ALLOW_NAMES), "name_regexes": HIDE_NAME_PATTERNS, "href_tails": HREFS, "debug": bool(debug)}
    js = f"""
    <script>
    (function() {{
      const CONF = { json.dumps(payload, ensure_ascii=False) };
      const allow = new Set(CONF.allow_names || []);
      const nameRes = (CONF.name_regexes || []).map(p => new RegExp(p));
      const hrefTails = new Set(CONF.href_tails || []);
      function normPath(href) {{
        try {{ return new URL(href, window.location.origin).pathname || ""; }}
        catch {{ return href || ""; }}
      }}
      function shouldHideByText(txt) {{
        if (!txt) return false;
        const t = txt.trim();
        if (allow.has(t)) return false;
        return nameRes.some(r => r.test(t));
      }}
      function shouldHideByHref(href) {{
        if (!href) return false;
        const p = normPath(href);
        for (const tail of hrefTails) {{
          if (!tail) continue;
          if (p.endsWith(tail) || p.includes(tail + "/") || p.includes(tail + "?")) return true;
        }}
        return false;
      }}
      function kill(el) {{ if (!el) return; el.style.setProperty('display','none','important'); el.setAttribute('data-hidden-by','bf-hide'); }}
      function scan(root) {{
        const sidebar = document.querySelector('aside[data-testid="stSidebar"]')
                      || document.querySelector('section[data-testid="stSidebar"]')
                      || document.querySelector('[data-testid="stSidebar"]')
                      || document.querySelector('nav[aria-label="Sidebar navigation"]')
                      || (root || document);
        if (!sidebar) return;
        sidebar.querySelectorAll('a[href]').forEach(a => {{
          const href = a.getAttribute('href') || '';
          const txt  = (a.textContent || '').trim();
          if (shouldHideByText(txt) || shouldHideByHref(href)) {{
            const li = a.closest('li') || a.parentElement; kill(li || a);
          }}
        }});
        sidebar.querySelectorAll('*:not(a)').forEach(el => {{
          const txt = (el.textContent || '').trim();
          if (shouldHideByText(txt)) {{
            const li = el.closest('li') || el.closest('[data-testid="stSidebarNav"]') || el.parentElement;
            kill(li || el);
          }}
        }});
      }}
      let tick = 0;
      const timer = setInterval(() => {{ scan(document); tick++; if (tick > 60) clearInterval(timer); }}, 300);
      scan(document);
      const mo = new MutationObserver(() => scan(document));
      mo.observe(document.body, {{ childList: true, subtree: true }});
      if (CONF.debug) console.log('[bf-hide:ready]', CONF);
    }})();
    </script>
    """
    _components.html(js, height=0)

hide_analysis_pages_in_sidebar(__file__, debug=False)

# ============== 常量/配置（制冰机） ==============
# —— 质量明细固定清单（按你给的表逐条对应）
FIXED_DEFECTS = [
    {"code": "2.1.4.7",  "en": "Control bord",                           "cn": "温控板"},
    {"code": "2.1.5.1",  "en": "Debug",                                   "cn": "调试冰厚"},
    {"code": "2.1.8.1",  "en": "Door Hinge Broken",                       "cn": "物理-门轴断裂"},
    {"code": "2.1.4.4",  "en": "Water Pump",                              "cn": "水泵"},
    {"code": "2.1.4.1",  "en": "Leaking Refrigerant",                     "cn": "制冷剂泄漏"},
    {"code": "2.1.4.9",  "en": "Water probe",                             "cn": "水位探头"},
    {"code": "2.1.4.2",  "en": "Inlet Water Valve",                       "cn": "进水阀"},
    {"code": "2.1.4.8",  "en": "Harvest Probe",                           "cn": "化霜探头"},
    {"code": "2.1.8.8",  "en": "Water Curtain Hinge Broken",              "cn": "水帘板"},
    {"code": "2.1.4.5",  "en": "Compressor",                              "cn": "压缩机"},
    {"code": "2.1.4.17", "en": "evaporator",                              "cn": "物理-蒸发器损坏"},
    {"code": "2.1.9.1",  "en": "Leaking water form bin bottom",           "cn": "水箱底部漏水"},
    {"code": "2.1.9.4",  "en": "The drain is clogged",                    "cn": "排水管堵住"},
    {"code": "2.1.8.5",  "en": "Water Trough",                            "cn": "物理-水槽损坏"},
    {"code": "2.1.6.1",  "en": "Fan Motor working with noise",            "cn": "风机噪音"},
    {"code": "2.1.4.18", "en": "Drain Valve",                             "cn": "排水阀"},
    {"code": "2.1.4.11", "en": "Loose wiring harness connection",         "cn": "不制冷-线束松动"},
    {"code": "2.1.4.16", "en": "Water Curtain stuck by ice",              "cn": "水帘板"},
    {"code": "2.1.9.3",  "en": "Water leakage at the hose connection",    "cn": "接水管连接处漏水"},
    {"code": "2.1.8.3",  "en": "Ice Harvest Probe",                       "cn": "物理损坏-探头"},
    {"code": "2.1.9.2",  "en": "Leaking water in the front",              "cn": "前面漏水"},
    {"code": "2.1.8.6",  "en": "Thumb Screw falls off",                   "cn": "水槽螺丝脱落"},
    {"code": "2.1.4.18", "en": "Solenoid valve",                          "cn": "电磁阀"},
    {"code": "2.1.4.17", "en": "Water Curtain sensor",                    "cn": "门磁开关"},
    {"code": "2.1.6.3",  "en": "Loose screws/hit something",              "cn": "螺丝松动/打到异物"},
    {"code": "2.1.4.15", "en": "condenser fan",                           "cn": "冷凝风机"},
    {"code": "2.1.6.2",  "en": "Water pump working with noise",           "cn": "玻璃门玻璃之间凝露/起雾"},
]

# 注意：上面移除了你原先重复的 2.1.4.17（避免与蒸发器重复）

# —— 型号系列维度（按在保数量口径）
ICE_MODEL_CATEGORY_MAP = {
    "YR140":  "制冰机-140",
    "YR280":  "制冰机-280",
    "YR450":  "制冰机-450",
    "YR800":  "制冰机-800",
    "CYR400": "冰桶-400",
    "CYR700": "冰桶-700",
    "HD350":  "制冰机-350",
}
ICE_MODEL_ORDER = ["合计", "YR140","YR280","YR450","YR800","CYR400","CYR700","HD350"]
ALIASES = {

    "Ice Machine-140": "YR140",
    "Ice Machine-280": "YR280",
    "Ice Machine-450": "YR450",
    "Ice Machine-800": "YR800",
    "Ice Machine-350": "HD350",
    "Ice machine-CYR400P": "CYR400",
    "Ice machine-CYR700P": "CYR700",
}

YEARS_FIXED = list(range(2025, 2031))

# —— 持久化文件名全部独立 *_ice
STORE_PARQUET = Path("data_store_ice.parquet")
STORE_CSV     = Path("data_store_ice.csv")
TARGETS_JSON  = Path("targets_config_ice.json")
MODEL_FLEET_JSON  = Path("model_fleet_counts_ice.json")
RAW_EXCEL_STORE = Path("raw_excel_store_ice.parquet")

# ============== 数据/文件存取 ==============
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

def _save_raw_excel_store(df: pd.DataFrame):
    try:
        df.to_parquet(RAW_EXCEL_STORE, index=False)
    except Exception:
        df.to_csv(RAW_EXCEL_STORE.with_suffix(".csv"), index=False, encoding="utf-8-sig")

def _clear_raw_excel_history():
    """清空原始 Excel 上传历史（内存 + 磁盘文件）"""
    st.session_state["excel_raw_history"] = pd.DataFrame()
    try: RAW_EXCEL_STORE.unlink(missing_ok=True)
    except Exception: pass
    try: RAW_EXCEL_STORE.with_suffix(".csv").unlink(missing_ok=True)
    except Exception: pass

def load_store_df()->pd.DataFrame:
    if STORE_PARQUET.exists():
        try: return pd.read_parquet(STORE_PARQUET)
        except Exception: pass
    if STORE_CSV.exists():
        try: return pd.read_csv(STORE_CSV)
        except Exception: pass
    return pd.DataFrame()

def save_store_df(df: pd.DataFrame):
    try: df.to_parquet(STORE_PARQUET, index=False)
    except Exception: df.to_csv(STORE_CSV, index=False, encoding="utf-8-sig")

def load_targets()->dict:
    if TARGETS_JSON.exists():
        try: return json.loads(TARGETS_JSON.read_text(encoding="utf-8"))
        except Exception: pass
    return {"target_year_rate":5.0,"q1_t":0.0,"q2_t":0.0,"q3_t":0.0,"q4_t":0.0,"year_t":5.0}

def save_targets(cfg: dict):
    try: TARGETS_JSON.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception: pass

def load_model_fleet()->dict:
    if MODEL_FLEET_JSON.exists():
        try: return json.loads(MODEL_FLEET_JSON.read_text(encoding="utf-8"))
        except Exception: pass
    return {}

def save_model_fleet(d: dict):
    try: MODEL_FLEET_JSON.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception: pass

def ensure_fleet_month(fleet: dict, year: int, month: int)->dict:
    y, m = str(year), str(month)
    fleet.setdefault(y, {})
    fleet[y].setdefault(m, {k:0 for k in ICE_MODEL_ORDER if k!="合计"})
    return fleet

def get_fleet_month(fleet: dict, year: int, month: int)->dict:
    y, m = str(year), str(month)
    return {k:int(fleet.get(y,{}).get(m,{}).get(k,0)) for k in ICE_MODEL_ORDER if k!="合计"}

def set_fleet_month(fleet: dict, year: int, month: int, updates: dict)->dict:
    fleet = ensure_fleet_month(fleet, year, month)
    y, m = str(year), str(month)
    for mdl, v in updates.items():
        if mdl in ICE_MODEL_CATEGORY_MAP:
            fleet[y][m][mdl] = int(max(0, v))
    return fleet

# ============== 字段/解析辅助 ==============
CATEGORY_SCHEMA = {
    "质量问题": ["Atosa"],
    "非质量问题": ["合计","安装/客户取消/联系不到客户","没安装净水器/使用错误","维修工未发现问题","客户付钱","脏堵/排水管堵"],
    "出保": ["只寄配件"],
    "待定": ["未收到账单/其他人付款"]
}
KEYWORD_RULES = [
    {"patterns":[r"2\.1\.a"],"cat1":"质量问题","cat2":"Atosa","paidqty":1,"iswarranty":True},
    {"patterns":[r"\batosa\b"],"cat1":"质量问题","cat2":"Atosa","paidqty":1,"iswarranty":True},
    {"patterns":[r"(?<!\d)2\.2(?!\d)"],"cat1":"非质量问题","cat2":"安装/客户取消/联系不到客户","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.3(?!\d)"],"cat1":"非质量问题","cat2":"安装/客户取消/联系不到客户","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.2(?!\d)"],"cat1":"非质量问题","cat2":"安装/客户取消/联系不到客户","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.5(?!\d)"],"cat1":"非质量问题","cat2":"没安装净水器/使用错误","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.8(?!\d)"],"cat1":"非质量问题","cat2":"没安装净水器/使用错误","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.11(?!\d)"],"cat1":"非质量问题","cat2":"没安装净水器/使用错误","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.7(?!\d)"],"cat1":"非质量问题","cat2":"维修工未发现问题","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.9(?!\d)"],"cat1":"非质量问题","cat2":"客户付钱","paidqty":0,"iswarranty":False},
    # 允许英文描述命中仓库维修类（移除你原先那条前缀“1”的误写）
    {"patterns":[r"repair(?:ing)?\s+in\s+(?:the\s+)?warehouse(?:s)?\.?"],"cat1":"非质量问题","cat2":"脏堵/排水管堵","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.4(?!\d)"],"cat1":"非质量问题","cat2":"脏堵/排水管堵","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.4\.6(?!\d)"],"cat1":"非质量问题","cat2":"脏堵/排水管堵","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.3(?!\d)"],"cat1":"出保","cat2":"只寄配件","paidqty":0,"iswarranty":False},
    {"patterns":[r"(?<!\d)2\.5(?!\d)"],"cat1":"待定","cat2":"未收到账单/其他人付款","paidqty":0,"iswarranty":False},
]
TEXT_COLUMNS_FOR_MATCH = [
    "TAG1","TAG2","TAG3","TAG4","Category1","Category2",
    "问题描述","描述","原因","备注","状态","标题","Summary","Notes",
    "编码","Code","DefectCode","问题","Problem"
]

def fmt_money(x):
    try: return f"${x:,.0f}"
    except Exception: return "$0"

def _clean_text(s: str)->str:
    s = ("" if pd.isna(s) else str(s)).lower().replace("\u3000", " ")
    s = re.sub(r"[\s\-_\/]+", " ", s)
    return s.strip()

_MMYYYY_ANY = re.compile(r'(?:^|\D)(?P<mm>0[1-9]|1[0-2])(?P<yyyy>20\d{2})(?:\D|$)')

def _extract_year_month(df: pd.DataFrame, date_col_name: str=None):
    if df.empty: return
    if date_col_name is None or date_col_name not in df.columns:
        date_col_name = df.columns[0]
    s = df[date_col_name].astype(str)
    m = s.str.extract(_MMYYYY_ANY)
    mm = pd.to_numeric(m["mm"], errors="coerce").astype("Int64")
    yy = pd.to_numeric(m["yyyy"], errors="coerce").astype("Int64")
    mm = mm.where((mm>=1)&(mm<=12), pd.NA)
    df["_SrcYM"] = (m["mm"].fillna("") + m["yyyy"].fillna("")).where(m["mm"].notna(), "")
    df["Year"] = yy
    df["Month"]= mm
    df["Date"] = pd.to_datetime(dict(year=yy.astype("float"), month=mm.astype("float"), day=1), errors="coerce")
    q = pd.Series(pd.NA, index=df.index, dtype="Int64")
    ok = df["Month"].notna()
    q.loc[ok] = ((df.loc[ok,"Month"].astype(int)-1)//3+1).astype("Int64")
    df["Quarter"]= q

def ensure_cols(df: pd.DataFrame, date_col_name: str=None)->pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Date","Category1","Category2","TAG1","TAG2","TAG3","TAG4","CostUSD","PaidQty","IsWarranty","Completed","Year","Month","Quarter","_SrcYM"])
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
        elif lc in ["cost","costusd","amount","amount_usd","费用","金额","fee","cost(usd)","cost usd"]: mapping[c] = "CostUSD"
        elif lc in ["已付款数量","paidqty","paid_qty"]: mapping[c] = "PaidQty"
        elif lc in ["是否承保","iswarranty","warranty"]: mapping[c] = "IsWarranty"
        elif lc in ["是否完成","completed","iscompleted","维修状态","状态"]: mapping[c] = "Completed"
    if mapping: df = df.rename(columns=mapping)
    if df.columns.duplicated().any():
       df = df.loc[:, ~df.columns.duplicated(keep="first")]
    for i in range(1,4+1):
        std = f"TAG{i}"
        if std not in df.columns:
            cand = [c for c in df.columns if re.fullmatch(rf"\s*tag[\s_-]*{i}\s*", str(c), flags=re.I)]
            if cand: df.rename(columns={cand[0]: std}, inplace=True)
    for col, default in [("Category1",""),("Category2",""),("TAG1",""),("TAG2",""),("TAG3",""),("TAG4",""),("CostUSD",0.0),("PaidQty",0),("IsWarranty",False),("Completed",False)]:
        if col not in df.columns: df[col] = default
    df["CostUSD"] = pd.to_numeric(df["CostUSD"], errors="coerce").fillna(0.0)
    df["PaidQty"] = pd.to_numeric(df["PaidQty"], errors="coerce").fillna(0).astype(int)
    if df["Completed"].dtype == object:
        df["Completed"] = df["Completed"].astype(str).str.contains(r"complete|已完成|1", case=False, na=False)
    else:
        df["Completed"] = df["Completed"].astype(bool)
    df["IsWarranty"] = df["IsWarranty"].astype(bool)
    need_parse = ("Year" not in df.columns) or ("Month" not in df.columns) or (pd.isna(pd.to_numeric(df.get("Year"), errors="coerce")).all()) or (pd.isna(pd.to_numeric(df.get("Month"), errors="coerce")).all())
    if need_parse: _extract_year_month(df, date_col_name)
    df = apply_keyword_overrides(df)
    return df

def apply_keyword_overrides(df: pd.DataFrame)->pd.DataFrame:
    if df.empty: return df
    joined_all = df.astype(str).agg(" ".join, axis=1).map(_clean_text)
    assigned = pd.Series(False, index=df.index)
    for rule in KEYWORD_RULES:
        pats = [re.compile(p, re.I) for p in rule["patterns"]]
        mask = ~assigned
        for p in pats: mask &= joined_all.str.contains(p)
        if mask.any():
            df.loc[mask,"Category1"] = rule["cat1"]
            df.loc[mask,"Category2"] = rule["cat2"]
            if "paidqty" in rule: df.loc[mask,"PaidQty"] = int(rule["paidqty"])
            if "iswarranty" in rule: df.loc[mask,"IsWarranty"] = bool(rule["iswarranty"])
            assigned |= mask
    return df

def compute_category_stats(df_year: pd.DataFrame):
    stats = {}
    if not df_year.empty:
        tmp = df_year.copy()
        tmp["_paid_flag"] = (pd.to_numeric(tmp["CostUSD"], errors="coerce").fillna(0)>0).astype(int)
        g = tmp.groupby(["Category1","Category2"], dropna=False).agg(数量=("CostUSD","size"),费用=("CostUSD","sum"),已付款数量=("_paid_flag","sum"))
        g["平均费用"] = g.apply(lambda r: (r["费用"]/r["已付款数量"]) if r["已付款数量"]>0 else 0.0, axis=1)
        for (c1,c2), r in g.iterrows():
            stats[(str(c1),str(c2))] = (int(r["数量"]), float(r["费用"]), int(r["已付款数量"]), float(r["平均费用"]))
    _nonq = ["安装/客户取消/联系不到客户","没安装净水器/使用错误","维修工未发现问题","脏堵/排水管堵"]
    _big = "非质量问题"; _q=_f=_p=0
    for s in _nonq:
        q,f,p,_ = stats.get((_big,s),(0,0.0,0,0.0))
        _q+=q; _f+=f; _p+=p
    _avg = (_f/_p) if _p>0 else 0.0
    stats[(_big,"合计")] = (_q,_f,_p,_avg)
    for big, subs in CATEGORY_SCHEMA.items():
        for sub in subs:
            stats.setdefault((big,sub),(0,0.0,0,0.0))
    return stats

def render_category_table_html(stats: dict)->str:
    if not isinstance(stats, dict):
        stats = {}
    for big, subs in CATEGORY_SCHEMA.items():
        for sub in subs:
            stats.setdefault((big, sub), (0, 0.0, 0, 0.0))
    BG, HEAD_BG = "#0b1534", "#0a1a55"
    ROW_ALT, TEXT, SUBTEXT = "rgba(255,255,255,.03)", "#e8edf7", "#c8d0e8"
    BORDER, BORDER_SOFT, HILITE_BG = "rgba(255,255,255,.14)", "rgba(255,255,255,.10)", "rgba(255,255,255,.05)"
    th = """
    <thead><tr>
      <th style="text-align:left;">分类1</th>
      <th style="text-align:left;">分类2</th>
      <th style="text-align:right;">数量</th>
      <th style="text-align:right;">费用</th>
      <th style="text-align:right;">已付款数量</th>
      <th style="text-align:right;">平均费用</th>
    </tr></thead>"""
    rows_html = []
    for big, subs in CATEGORY_SCHEMA.items():
        rowspan = len(subs); first = True
        for sub in subs:
            qty, fee, paid, avg = stats.get((big, sub), (0, 0.0, 0, 0.0))
            tr = "<tr>"
            if first:
                tr += f'<td rowspan="{rowspan}" class="td-big"><b>{big}</b></td>'
                first = False
            if (big == "非质量问题" and sub == "合计"):
                tr += (
                    f'<td class="td-strong">{sub}</td>'
                    f'<td class="td-num td-strong">{qty}</td>'
                    f'<td class="td-num td-strong">{fmt_money(fee)}</td>'
                    f'<td class="td-num td-strong">{paid}</td>'
                    f'<td class="td-num td-strong">{fmt_money(avg)}</td>'
                )
            else:
                tr += (
                    f"<td>{sub}</td>"
                    f'<td class="td-num">{qty}</td>'
                    f'<td class="td-num">{fmt_money(fee)}</td>'
                    f'<td class="td-num">{paid}</td>'
                    f'<td class="td-num">{fmt_money(avg)}</td>'
                )
            tr += "</tr>"
            rows_html.append(tr)
    styles = f"""
    <style>
      .tbl-wrap {{ background:{BG}; border:1px solid {BORDER}; border-radius:14px; overflow:hidden; box-shadow: 0 4px 14px rgba(0,0,0,.3); }}
      table.tbl {{ width:100%; border-collapse:separate; border-spacing:0; font-size:14px; color:{TEXT}; }}
      .tbl thead th {{ background:{HEAD_BG}; color:{TEXT}; padding:10px 12px; border-bottom:1px solid {BORDER}; font-weight:700; white-space:nowrap; }}
      .tbl tbody td {{ padding:9px 12px; border-bottom:1px solid {BORDER_SOFT}; color:{SUBTEXT}; background:transparent; }}
      .tbl tbody tr:nth-child(even) td {{ background:{ROW_ALT}; }}
      .tbl .td-num {{ text-align:right; color:{TEXT}; }}
      .tbl .td-big {{ vertical-align:top; border-right:1px solid {BORDER}; color:{TEXT}; background:rgba(255,255,255,.02); min-width:120px; }}
      .tbl .td-strong {{ color:{TEXT}; background:{HILITE_BG}; font-weight:700; }}
    </style>
    """
    html = f"""{styles}<div class="tbl-wrap"><table class="tbl">{th}<tbody>{''.join(rows_html)}</tbody></table></div>"""
    return html

def build_monthly_numeric(df_year: pd.DataFrame, fleet_all: dict, current_year: int)->pd.DataFrame:
    months = range(1, 12+1)
    joined_year = (pd.Series(dtype=str) if df_year.empty else
                   df_year.astype(str).agg(" ".join, axis=1).map(_clean_text))

    rows = {
        "月度维修量": [],
        "月度质量维修量": [],
        "平均维修单费用（美元）": [],
        "月度预估保内维修费用（美元）": [],
        "月度保内数量": [],
        "折算保内维修率%": [],
        "月度费用（分类口径）": [],
        "月度已付款数量（分类口径）": [],
    }

    pat_21a  = re.compile(r"(?<![0-9a-z])2\.1\.a(?![0-9a-z])", re.I)
    pat_213x = re.compile(r"(?<![0-9a-z])2\.1\.3\.\d+(?![0-9a-z])", re.I)

    for m in months:
        if df_year.empty:
            total = 0
            quality_cnt = 0
            fees4 = 0.0
            paid4 = 0
        else:
            dfm = df_year[df_year["Month"] == m]
            total = len(dfm)
            if total > 0:
                cat_quality = dfm["Category1"].astype(str).eq("质量问题")
                text_block  = joined_year.loc[dfm.index]
                text_hits   = text_block.str.contains(pat_21a, na=False) | text_block.str.contains(pat_213x, na=False)
                quality_cnt = int((cat_quality | text_hits).sum())
            else:
                quality_cnt = 0

            stats_m = compute_category_stats(dfm)
            def _pick(c1, c2):
                return stats_m.get((c1, c2), (0, 0.0, 0, 0.0))
            _, fee_atosa,      paid_atosa,      _ = _pick("质量问题", "Atosa")
            _, fee_nonq_sum,   paid_nonq_sum,   _ = _pick("非质量问题", "合计")
            _, fee_out,        paid_out,        _ = _pick("出保", "只寄配件")
            _, fee_pending,    paid_pending,    _ = _pick("待定", "未收到账单/其他人付款")

            fees4 = float(fee_atosa + fee_nonq_sum + fee_out + fee_pending)
            paid4 = int(paid_atosa + paid_nonq_sum + paid_out + paid_pending)

        avg_cost = (fees4 / paid4) if paid4 > 0 else 0.0
        month_fleet = get_fleet_month(fleet_all, int(current_year), int(m))
        warranty_cnt = int(sum(month_fleet.values()))
        est_cost = avg_cost * quality_cnt
        rate_pct = (quality_cnt / warranty_cnt * 100.0) if warranty_cnt > 0 else 0.0

        rows["月度维修量"].append(total)
        rows["月度质量维修量"].append(quality_cnt)
        rows["平均维修单费用（美元）"].append(avg_cost)
        rows["月度预估保内维修费用（美元）"].append(est_cost)
        rows["月度保内数量"].append(warranty_cnt)
        rows["折算保内维修率%"].append(rate_pct)
        rows["月度费用（分类口径）"].append(fees4)
        rows["月度已付款数量（分类口径）"].append(paid4)

    pv = pd.DataFrame(rows).T
    pv = pv.reindex([
        "月度维修量",
        "月度质量维修量",
        "平均维修单费用（美元）",
        "月度预估保内维修费用（美元）",
        "月度保内数量",
        "折算保内维修率%",
        "月度费用（分类口径）",
        "月度已付款数量（分类口径）",
    ])
    pv.columns = [f"{m}月" for m in months]
    for c in pv.columns:
        pv[c] = pd.to_numeric(pv[c], errors="coerce").fillna(0)
    return pv

def build_monthly_display(pv_numeric: pd.DataFrame)->pd.DataFrame:
    pv = pv_numeric.copy()
    for r in ["平均维修单费用（美元）","月度预估保内维修费用（美元）"]:
        pv.loc[r] = pv.loc[r].map(fmt_money)
    pv.loc["折算保内维修率%"] = pv.loc["折算保内维修率%"].map(lambda x: f"{x:.2f}%")
    for r in ["月度维修量","月度质量维修量","月度保内数量"]:
        pv.loc[r] = pv.loc[r].astype(int)
    pv = pv.drop(index=["月度费用（分类口径）","月度已付款数量（分类口径）"], errors="ignore")
    return pv

def build_quarter_from_monthly(pv_numeric: pd.DataFrame, *, fleet_mode: str = "avg")->pd.DataFrame:
    q_cols = {
        "Q1": ["1月","2月","3月"],
        "Q2": ["4月","5月","6月"],
        "Q3": ["7月","8月","9月"],
        "Q4": ["10月","11月","12月"],
    }
    q_num = pd.DataFrame(index=pv_numeric.index, columns=list(q_cols.keys())+["全年"], dtype=float)
    for q, cols in q_cols.items():
        q_num[q] = pv_numeric[cols].sum(axis=1, numeric_only=True)
    q_num["全年"] = pv_numeric.sum(axis=1, numeric_only=True)

    fee_row  = "月度费用（分类口径）"
    paid_row = "月度已付款数量（分类口径）"
    avg_row  = "平均维修单费用（美元）"
    if fee_row in pv_numeric.index and paid_row in pv_numeric.index and avg_row in pv_numeric.index:
        for q, cols in q_cols.items():
            fees_q  = float(pv_numeric.loc[fee_row,  cols].sum())
            paid_q  = float(pv_numeric.loc[paid_row, cols].sum())
            q_num.loc[avg_row, q] = (fees_q / paid_q) if paid_q > 0 else 0.0
        fees_y = float(pv_numeric.loc[fee_row].sum())
        paid_y = float(pv_numeric.loc[paid_row].sum())
        q_num.loc[avg_row, "全年"] = (fees_y / paid_y) if paid_y > 0 else 0.0

    rename = {
        "月度维修量": "季度维修量",
        "月度质量维修量": "季度质量维修量",
        "平均维修单费用（美元）": "平均维修单费用（美元）",
        "月度预估保内维修费用（美元）": "季度预估保内维修费用（美元）",
        "月度保内数量": "季度保内数量",
        "折算保内维修率%": "折算保内维修率%",
        fee_row: fee_row,
        paid_row: paid_row,
    }
    q_num.index = [rename.get(i, i) for i in q_num.index]

    if "季度保内数量" in q_num.index and "月度保内数量" in pv_numeric.index:
        if fleet_mode.lower() == "avg":
            for q, cols in q_cols.items():
                q_num.loc["季度保内数量", q] = float(pv_numeric.loc["月度保内数量", cols].mean())
            q_num.loc["季度保内数量", "全年"] = float(pv_numeric.loc["月度保内数量"].mean())

    q_show = q_num.copy()
    for r in ["平均维修单费用（美元）", "季度预估保内维修费用（美元）"]:
        if r in q_show.index:
            q_show.loc[r] = q_show.loc[r].map(lambda v: f"${v:,.0f}")
    if "折算保内维修率%" in q_show.index:
        q_show.loc["折算保内维修率%"] = q_show.loc["折算保内维修率%"].map(lambda v: f"{float(v):.2f}%")
    for r in ["季度维修量", "季度质量维修量", "季度保内数量"]:
        if r in q_show.index:
            q_show.loc[r] = q_show.loc[r].astype(int)

    q_show = q_show.drop(index=[rename[fee_row], rename[paid_row]], errors="ignore")
    return q_show

# ============== 侧栏（上传/清空仅管理员） ==============
with st.sidebar:
    logout_button()
    u = current_user() or {"username":"未登录","name":"访客","role":"viewer"}
    st.markdown(f"**当前用户**：{u.get('name','')}（{u.get('username','')} / {u.get('role','')}）")

    persisted_file = ensure_cols(load_store_df())
    st.session_state["store_df"] = ensure_cols(persisted_file.copy())

    if is_admin():
        st.markdown("---")
        st.subheader("🔐 管理员区")

        up = st.file_uploader("上传数据（CSV/XLSX）", type=["csv","xlsx","xls"])

        def _read_any(up_file):
            if up_file is None: return pd.DataFrame()
            name = up_file.name.lower()
            try:
                if name.endswith((".xlsx",".xls")):
                    df0 = pd.read_excel(up_file, sheet_name=0, header=0, dtype=str)
                else:
                    df0 = pd.read_csv(up_file, dtype=str, encoding="utf-8-sig")
            except UnicodeDecodeError:
                df0 = pd.read_csv(up_file, dtype=str, encoding="gbk", errors="ignore")
            except Exception as e:
                st.error(f"读取失败：{e}"); return pd.DataFrame()
            return df0

        df_raw = _read_any(up)

        raw_store = _load_raw_excel_store()
        if "excel_raw_history" not in st.session_state:
            st.session_state["excel_raw_history"] = raw_store.copy()
        if not df_raw.empty:
            df_raw_add = df_raw.copy()
            df_raw_add["__SourceFile"] = getattr(up, "name", "") if up is not None else ""
            df_raw_add["__ImportedAt"] = pd.Timestamp.now(tz=None)
            st.session_state["excel_raw_history"] = pd.concat(
                [st.session_state["excel_raw_history"], df_raw_add], ignore_index=True
            )
            _save_raw_excel_store(st.session_state["excel_raw_history"])

        date_col_name = (list(df_raw.columns)[0] if not df_raw.empty else None)
        st.session_state["date_col_name"] = date_col_name
        df = ensure_cols(df_raw, date_col_name=date_col_name)

        def _bytes_md5(b: bytes)->str:
            h=hashlib.md5(); h.update(b); return h.hexdigest()
        upload_md5=None
        if up is not None:
            try: upload_md5=_bytes_md5(up.getvalue())
            except Exception: upload_md5=None
        is_new_upload=False
        if upload_md5:
            last_md5=st.session_state.get("_last_upload_md5_ice")
            if last_md5!=upload_md5:
                is_new_upload=True
                st.session_state["_last_upload_md5_ice"]=upload_md5

        if not df.empty:
            df = df.copy()
            df["SourceFile"] = getattr(up, "name", "") if up is not None else ""
            df["ImportedAt"] = pd.Timestamp.now(tz=None)

        merge_mode = "仅追加（不去重）"
        st.caption("合并策略：仅追加（不去重）")

        if is_new_upload and not df.empty:
            merged = ensure_cols(pd.concat([persisted_file, df], ignore_index=True))
            save_store_df(merged)
            persisted = merged
        else:
            persisted = persisted_file

        st.session_state["store_df"] = ensure_cols(persisted.copy())
        st.caption(f"已存数据量：{len(st.session_state['store_df'])} 条（合并策略：{merge_mode}）")

        if not st.session_state["store_df"].empty:
            tmp=st.session_state["store_df"].copy()
            tmp["Year"]=pd.to_numeric(tmp["Year"], errors="coerce")
            tmp["Month"]=pd.to_numeric(tmp["Month"], errors="coerce")
            ym_count=(tmp.dropna(subset=["Year","Month"]).astype({"Year":int,"Month":int})
                      .groupby(["Year","Month"]).size().reset_index(name="数量").sort_values(["Year","Month"]))
            with st.expander("📊 Year/Month 数量分布（默认收起）", expanded=False):
                st.dataframe(ym_count, use_container_width=True, height=200)

        st.markdown("---")
        with st.expander("📎 原始 Excel 历史维护", expanded=False):
            st.caption("这里只影响『原始 Excel 表映射』用到的历史，不影响已入库的分析数据。")
            if st.button("清空原始 Excel 上传历史", type="primary", use_container_width=True):
                _clear_raw_excel_history()
                st.success("已清空『原始 Excel 上传历史』（内存 + 磁盘）。")
                st.rerun()

        st.subheader("🧹 数据清空（管理员）")
        clear_mode = st.radio("选择清空范围", ["清空所选 年份+月份","清空所选 年份（整年）","清空全部数据","仅清空年度目标配置"], index=0)
        confirm = st.checkbox("我已确认要执行清空操作")
        if st.button("执行清空", type="primary", use_container_width=True, disabled=not confirm):
            df_all = st.session_state.get("store_df", pd.DataFrame()).copy()
            if clear_mode=="仅清空年度目标配置":
                save_targets({"target_year_rate":5.0,"q1_t":0.0,"q2_t":0.0,"q3_t":0.0,"q4_t":0.0,"year_t":5.0})
                for k,v in {"target_year_rate":5.0,"q1_t":0.0,"q2_t":0.0,"q3_t":0.0,"q4_t":0.0,"year_t":5.0}.items():
                    st.session_state[k]=v
                st.success("已清空年度目标配置。")
            else:
                if clear_mode=="清空全部数据":
                    df_all = pd.DataFrame(columns=ensure_cols(pd.DataFrame()).columns)
                    _clear_raw_excel_history()
                elif clear_mode=="清空所选 年份（整年）":
                    y_ser = pd.to_numeric(df_all.get("Year", pd.Series(dtype="Int64")), errors="coerce")
                    df_all = df_all[~(y_ser==int(st.session_state.get("year", 2025)))].reset_index(drop=True)
                elif clear_mode=="清空所选 年份+月份":
                    y_ser = pd.to_numeric(df_all.get("Year"), errors="coerce")
                    m_ser = pd.to_numeric(df_all.get("Month"), errors="coerce")
                    months_to_clear = st.session_state.get("selected_period_months", [1])
                    df_all = df_all[~((y_ser==int(st.session_state.get("year", 2025))) & (m_ser.isin(months_to_clear)))].reset_index(drop=True)
                save_store_df(ensure_cols(df_all))
                st.session_state["store_df"]=ensure_cols(df_all.copy())
                st.success("已完成清空并保存。")

    # ===== 所有人可见：年份 / 月份(季度)选择 & 年度目标配置 =====
    df_all_ss = st.session_state.get("store_df", pd.DataFrame())
    years=YEARS_FIXED[:]
    if not df_all_ss.empty:
        _yy=pd.to_numeric(df_all_ss.get("Year", pd.Series(dtype="Int64")), errors="coerce")
        present_years=sorted(set(int(y) for y in _yy.dropna().astype(int).tolist() if y in years))
    else:
        present_years=[]
    default_year=(max(present_years) if present_years else 2025)
    default_year_idx=years.index(default_year)

    if not df_all_ss.empty:
        _yy2=pd.to_numeric(df_all_ss.get("Year", pd.Series(dtype="Int64")), errors="coerce")
        _mm2=pd.to_numeric(df_all_ss.get("Month", pd.Series(dtype="Int64")), errors="coerce")
        months_in_default=sorted(_mm2[(_yy2==default_year)].dropna().astype(int).unique().tolist())
    else:
        months_in_default=[]
    default_month=(max(months_in_default) if months_in_default else 1)

    year = st.selectbox("选择年份", years, index=default_year_idx, key="year")

    PERIOD_OPTIONS = (
        [{"type": "M", "months": [m], "label": f"{m:02d}月"} for m in range(1, 13)]
        + [
            {"type": "Q", "months": [1, 2, 3],   "label": "一季度 (01–03)"},
            {"type": "Q", "months": [4, 5, 6],   "label": "二季度 (04–06)"},
            {"type": "Q", "months": [7, 8, 9],   "label": "三季度 (07–09)"},
            {"type": "Q", "months": [10, 11, 12],"label": "四季度 (10–12)"},
        ]
    )
    _default_label = f"{int(default_month):02d}月"
    _default_idx = next((i for i, o in enumerate(PERIOD_OPTIONS) if o["label"] == _default_label), 0)

    period_sel = st.selectbox(
        "选择月份 / 季度",
        PERIOD_OPTIONS,
        index=_default_idx,
        format_func=lambda o: o["label"],
    )
    st.session_state["selected_period_months"] = period_sel["months"]
    st.session_state["selected_period_label"]  = period_sel["label"]
    st.session_state["selected_period_is_quarter"] = (period_sel["type"] == "Q")

# 统计口径：只看 2025~2030
df_all = ensure_cols(st.session_state.get("store_df", pd.DataFrame())).copy()
st.session_state["store_df_ready"] = True
st.session_state["store_df"] = ensure_cols(df_all.copy())

if df_all.empty:
    st.info("当前没有任何可筛选的数据。请先由管理员上传（制冰机数据存储为独立 *_ice 文件，默认空）。")
else:
    _year_series  = pd.to_numeric(df_all.get("Year",  pd.Series(dtype="Int64")), errors="coerce")
    mask_valid_year = _year_series.isin(YEARS_FIXED)
    df_all = df_all[mask_valid_year].reset_index(drop=True)

_year_series  = pd.to_numeric(df_all.get("Year",  pd.Series(dtype="Int64")), errors="coerce")
_month_series = pd.to_numeric(df_all.get("Month", pd.Series(dtype="Int64")), errors="coerce")

if not df_all.empty:
    with st.expander("🧪 解析自检（默认收起）", expanded=False):
        valid = _year_series.notna() & _month_series.notna()
        st.write("有效可筛选的行数：", int(valid.sum()))
        st.write("当前选择（年份众数）：", f"{int(_year_series.mode().iloc[0]) if valid.any() else ''}")
        ym_dist = (pd.DataFrame({"Year": _year_series, "Month": _month_series})
                   .loc[(_year_series.notna()) & (_month_series.notna())]
                   .groupby(["Year","Month"]).size().reset_index(name="数量").sort_values(["Year","Month"]))
        st.dataframe(ym_dist, use_container_width=True, height=180)

cur_year  = int(st.session_state.get("year", 2025))
_sel_months: List[int] = st.session_state.get("selected_period_months", [1])
_sel_label  = st.session_state.get("selected_period_label", f"{_sel_months[0]:02d}月")
_is_quarter = bool(st.session_state.get("selected_period_is_quarter", False))

m_ser = pd.to_numeric(df_all.get("Month", pd.Series(dtype="Int64")), errors="coerce")
y_ser = pd.to_numeric(df_all.get("Year",  pd.Series(dtype="Int64")), errors="coerce")
df_period = df_all[(y_ser == cur_year) & (m_ser.isin(_sel_months))].copy()

edit_month = int(sorted(_sel_months)[-1])

st.markdown(f"### 分类统计（按所选期间：{cur_year}年 {_sel_label}）")
if df_period.empty:
    st.warning("当前所选期间没有匹配到任何数据。")
else:
    st.markdown(render_category_table_html(compute_category_stats(df_period)), unsafe_allow_html=True)

df_year_only = df_all[y_ser == cur_year]
st.markdown(f"### {cur_year}年度目标维修率为{int(load_targets()['target_year_rate']):d}%")
fleet = load_model_fleet()

pv_num  = build_monthly_numeric(df_year_only, fleet, int(cur_year))
pv_show = build_monthly_display(pv_num)
st.dataframe(pv_show, use_container_width=True)

st.markdown("### 季度统计数据（整年）")
q_show = build_quarter_from_monthly(pv_num, fleet_mode="avg")
st.dataframe(q_show, use_container_width=True)

# ============== 型号解析（与冰箱风格一致，兼容 Ice Machine / CYRxxxP / 纯数字） ==============
def resolve_model_key(raw: str) -> str:
    s = (raw or "").strip().replace("\u3000", " ")
    # 统一分隔符
    s = re.sub(r"[\s\-_\/]+", " ", s, flags=re.I).strip()

    # 先尝试按别名直接匹配（大小写/空格不敏感）
    key_up = s.upper().replace(" ", "")
    for k, v in ALIASES.items():
        if key_up == str(k).upper().replace(" ", ""):
            return v

    # 去掉前缀：REFRIGERATION- / ICE MACHINE-
    s = re.sub(r"^\s*(REFRIGERATION|ICE\s*MACHINE)\s*[-_:：/\\]*", "", s, flags=re.I).strip()
    s0 = re.sub(r"[\s\-_\/]+", "", s, flags=re.I).upper()

    # CYR400P / CYR700P -> 去尾缀 P
    s0 = re.sub(r"^CYR(400|700)P$", r"CYR\1", s0)

    # 纯数字：140/280/450/800/350
    if re.fullmatch(r"\d{3}", s0):
        if s0 == "350":
            return "HD350"
        elif s0 in {"140","280","450","800"}:
            return f"YR{s0}"

    # CYR 桶（容忍后缀）
    m = re.match(r"^CYR(400|700)", s0)
    if m:
        return f"CYR{m.group(1)}"

    # 兜底：已知前缀
    if s0.startswith("YR140"): return "YR140"
    if s0.startswith("YR280"): return "YR280"
    if s0.startswith("YR450"): return "YR450"
    if s0.startswith("YR800"): return "YR800"
    if s0.startswith("HD350") or s0.startswith("ICEBUCKET350"): return "HD350"

    # 其他未识别：返回标准化字符串（不会入库，因为不在 MAP 里）
    return s0
def _parse_bulk_paste(txt: str) -> dict:
    """
    一行一条，支持：
      - YR140 1000
      - Ice Machine-280,  900
      - CYR400P 200台
      - 350  1200        （→ HD350）
      - YR450\t800
    逻辑：取行尾数字（容许千分位/小数/单位），前半段做型号归一。
    """
    updates = {}
    if not txt or not txt.strip():
        return updates

    tail_num = re.compile(r"(?P<num>\d[\d,\.]*)\s*$")  # 抓行尾数字

    for raw_line in txt.strip().splitlines():
        line = str(raw_line).strip()
        if not line:
            continue
        # 跳过表头类行
        if re.search(r"(型号|series|在保|数量|合计)", line, re.I):
            continue

        m = tail_num.search(line)
        if not m:
            continue

        num_str = m.group("num")
        model_part = line[:m.start()].strip()

        mdl = resolve_model_key(model_part)
        if mdl not in ICE_MODEL_CATEGORY_MAP:
            continue

        try:
            val = int(float(num_str.replace(",", "")))
        except Exception:
            continue

        updates[mdl] = updates.get(mdl, 0) + max(0, val)

    return updates



st.markdown("### 型号系列维度（按在保数量口径）")
with st.expander("📝 录入：各型号在保数量（按月）", expanded=False):
    st.caption(f"当前作用年月：{cur_year}-{int(edit_month):02d}（若选择季度，这里默认使用该期间最后一个月）")
    tab1, tab2 = st.tabs(["批量粘贴（推荐）","逐个输入"])
    with tab1:
        sample="型号系列\t当月在保数量\nYR140\t1000\nYR280\t900\n..."
        txt=st.text_area("从 Excel 复制两列（型号系列 + 当月在保数量）后粘贴到此处：", height=180, placeholder=sample)
        col_l, col_r = st.columns([1,1])
        with col_l:
            if st.button("解析并写入当月在保数量", type="primary", use_container_width=True, key="btn_parse_fleet_month_ice"):
                upd=_parse_bulk_paste(txt)
                if not upd:
                    st.warning("没有解析到有效数据，请检查是否为两列（型号系列 + 当月在保数量），且型号在 ICE_MODEL_CATEGORY_MAP 内。")
                else:
                    fleet=set_fleet_month(fleet, int(cur_year), int(edit_month), upd)
                    save_model_fleet(fleet)
                    st.success(f"已写入 {len(upd)} 个型号的当月在保数量。")
                    st.rerun()
        with col_r:
            month_fleet=get_fleet_month(fleet, int(cur_year), int(edit_month))
            preview=pd.DataFrame({
                "型号系列":[m for m in ICE_MODEL_ORDER if m!="合计"],
                f"当月在保数量（{cur_year}-{int(edit_month):02d}）":[int(month_fleet.get(m,0)) for m in ICE_MODEL_ORDER if m!="合计"],
                "类别":[ICE_MODEL_CATEGORY_MAP.get(m,"") for m in ICE_MODEL_ORDER if m!="合计"],
            })
            st.dataframe(preview, use_container_width=True, height=240)
    with tab2:
        month_fleet=get_fleet_month(fleet, int(cur_year), int(edit_month))
        cols=st.columns(3); changed=False
        for i, mdl in enumerate([m for m in ICE_MODEL_ORDER if m!="合计"]):
            with cols[i%3]:
                cur=int(month_fleet.get(mdl,0))
                val=st.number_input(mdl, min_value=0, max_value=10_000_000, value=cur, step=1, key=f"fleet_ice_{cur_year}_{edit_month}_{mdl}")
                if val!=cur:
                    month_fleet[mdl]=int(val); changed=True
        if changed:
            fleet=set_fleet_month(fleet, int(cur_year), int(edit_month), month_fleet)
            save_model_fleet(fleet)
            st.success("已保存当月在保数量。")

def _find_model_col(df0: pd.DataFrame):
    pri_candidates = ["型号归类","型号系列","series","model_series"]
    sec_candidates = ["机型","具体型号","产品型号","Part Model","PartModel","Part","MODEL","model","型号"]
    for c in df0.columns:
        if str(c).strip().lower() in [n.lower() for n in pri_candidates]:
            return c
    for c in df0.columns:
        if str(c).strip().lower() in [n.lower() for n in sec_candidates]:
            return c
    return None

def build_model_table(df_all: pd.DataFrame, model_col: str, fleet_all: dict,
                      cur_year: int, sel_months: List[int], period_label: str,
                      use_avg_fleet: bool = False) -> pd.DataFrame:
    rows=[]; pat_21a = re.compile(r"(?<![0-9a-z])2\.1\.a(?![0-9a-z])", re.I)

    def _sum_fleet_months(fleet: dict, year: int, months: List[int]) -> dict:
        acc={k:0 for k in ICE_MODEL_ORDER if k!="合计"}
        for m in months:
            mf = get_fleet_month(fleet, year, int(m))
            for mdl, v in mf.items():
                acc[mdl] = acc.get(mdl, 0) + int(v or 0)
        return acc

    if df_all.empty or (model_col is None):
        df_cur = pd.DataFrame(columns=list(df_all.columns)+["_model_norm","_is_21a","_paid"])
    else:
        y = pd.to_numeric(df_all.get("Year"), errors="coerce")
        m = pd.to_numeric(df_all.get("Month"), errors="coerce")
        df_cur = df_all[(y==int(cur_year)) & (m.isin(sel_months))].copy()

    if not df_cur.empty and model_col in df_cur.columns:
        known = list(ICE_MODEL_CATEGORY_MAP.keys())
        known_sorted = sorted(known, key=lambda k: len(k.replace(" ","")), reverse=True)
        known_nospace = {k.replace(" ",""):k for k in known}
        def normalize_series(x:str)->str:
            s=(x or "")
            s_up = re.sub(r"^\s*REFRIGERATION\s*[-_:：/\\]\s*", "", s.upper())
            s_up = re.sub(r"[\s\-_\/]+"," ", s_up).strip()
            s0   = re.sub(r"[\s\-_\/]+","", s_up)
            if s0 in known_nospace: return known_nospace[s0]
            for k in known_sorted:
                if s0.startswith(k.replace(" ","")): return k
            m0 = re.match(r"^([A-Z]{2,4})(\d.*)?$", s0)
            if m0 and m0.group(1) in known_nospace: return known_nospace[m0.group(1)]
            return ""
        df_cur["_model_norm"] = df_cur[model_col].astype(str).map(normalize_series)
        joined = df_cur.astype(str).agg(" ".join, axis=1).map(_clean_text)
        df_cur["_is_21a"] = joined.str.contains(pat_21a, regex=True)
        df_cur["_paid"]   = pd.to_numeric(df_cur.get("CostUSD", 0), errors="coerce").fillna(0) > 0
    else:
        df_cur = pd.DataFrame(columns=list(df_all.columns)+["_model_norm","_is_21a","_paid"])

    period_fleet_sum = _sum_fleet_months(fleet_all, cur_year, sel_months)
    n_months = max(1, len(sel_months))
    col_title = (f"在保数量（{cur_year}年 {period_label} 平均）" if use_avg_fleet
                 else f"在保数量（{cur_year}年 {period_label} 合计）")

    total_fleet_sum = total_qp = 0
    total_cost  = 0.0

    for mdl in ICE_MODEL_ORDER:
        if mdl == "合计": continue
        fleet_sum = int(period_fleet_sum.get(mdl, 0))
        fleet_for_rate = int(round(fleet_sum / n_months)) if use_avg_fleet else fleet_sum
        total_fleet_sum += (fleet_sum if not use_avg_fleet else fleet_for_rate)

        if not df_cur.empty and "_model_norm" in df_cur.columns:
            mask_model = (df_cur["_model_norm"] == mdl)
            qual_cnt   = int((mask_model & df_cur["_is_21a"]).sum())
            month_cost = float(pd.to_numeric(df_cur.loc[mask_model & df_cur["_is_21a"] & df_cur["_paid"], "CostUSD"], errors="coerce").fillna(0).sum())
        else:
            qual_cnt = 0; month_cost = 0.0

        total_qp   += qual_cnt
        total_cost += month_cost
        rate = (qual_cnt / fleet_for_rate * 100.0) if fleet_for_rate > 0 else 0.0

        rows.append({
            "型号系列": mdl,
            "类别": ICE_MODEL_CATEGORY_MAP.get(mdl, ""),
            col_title: fleet_for_rate,
            "质量问题（期间合计）": qual_cnt,
            "期间维修率": f"{rate:.3f}%",
            "期间维修费用合计": fmt_money(month_cost),
        })

    total_rate = (total_qp / total_fleet_sum * 100.0) if total_fleet_sum > 0 else 0.0
    rows.insert(0, {
        "型号系列": "合计", "类别": "",
        col_title: total_fleet_sum,
        "质量问题（期间合计）": total_qp,
        "期间维修率": f"{total_rate:.3f}%",
        "期间维修费用合计": fmt_money(total_cost),
    })
    return pd.DataFrame(rows)

model_col = _find_model_col(df_all) if not df_all.empty else None
model_table = build_model_table(
    df_all, model_col, fleet, int(cur_year),
    sel_months=_sel_months,
    period_label=_sel_label,
    use_avg_fleet=_is_quarter
)
st.dataframe(model_table, use_container_width=True)

# ============== 固定清单聚合 + 链接到配件分析 ==============
def build_fixed_issue_table(df_scope: pd.DataFrame, mapping: list, *, drop_dupes: bool=False)->pd.DataFrame:
    if df_scope.empty:
        return pd.DataFrame(columns=["项目（编码+英文）","问题中文名","数量","费用"])

    # 直接按归一分类：质量问题/Atosa
    mask_quality = (df_scope["Category1"].astype(str)=="质量问题") & (df_scope["Category2"].astype(str)=="Atosa")
    if not mask_quality.any():
        return pd.DataFrame(columns=["项目（编码+英文）","问题中文名","数量","费用"])

    df_21a = df_scope.loc[mask_quality].copy()

    text_cols = ["Category1","Category2","TAG1","TAG2","TAG3","TAG4","问题描述","描述","原因","备注","状态","标题","Summary","Notes"]
    has_cols = [c for c in text_cols if c in df_21a.columns] or df_21a.columns.tolist()
    joined = df_21a[has_cols].astype(str).agg(" ".join, axis=1).str.lower()

    if drop_dupes:
        dedupe_keys = [c for c in ["Date","Category1","Category2","TAG1","TAG2","TAG3","TAG4","Summary","Notes"] if c in df_21a.columns]
        if dedupe_keys:
            df_21a["_dedupe_key"] = df_21a[dedupe_keys].astype(str).agg("|".join, axis=1)
            df_21a = df_21a.drop_duplicates(subset=["_dedupe_key"], keep="first")
            joined = df_21a[has_cols].astype(str).agg(" ".join, axis=1).str.lower()

    cost = pd.to_numeric(df_21a.get("CostUSD", 0), errors="coerce").fillna(0.0)
    paid_mask = (cost > 0)

    rows = []
    for item in mapping:
        code = item["code"].strip()
        en   = item["en"].strip()
        cn   = item["cn"].strip()
        code_pat = re.compile(rf"(?<![0-9a-z]){re.escape(code)}(?![0-9a-z])", re.I)
        m = joined.str.contains(code_pat)
        qty = int(m.sum())
        fee_val = float(cost[m & paid_mask].sum())
        rows.append({
            "项目（编码+英文）": f"{code} {en}",
            "问题中文名": cn,
            "数量": qty,
            "费用": f"${fee_val:,.0f}",
            "_fee_val": fee_val,
            "_code": code
        })

    out = (pd.DataFrame(rows)
           .sort_values(["数量","_fee_val"], ascending=[False, False])
           .reset_index(drop=True))
    return out

st.markdown("### 质量问题明细（固定清单，自动汇总）")
scope = st.radio("统计口径", ["按所选期间（月份/季度）","按全年"], index=0, horizontal=True)
df_year_only = df_all[_year_series==cur_year]
if scope == "按所选期间（月份/季度）":
    df_scope = df_period
    title_suffix = f"{cur_year}年 {_sel_label}"
    _dl_tag = _sel_label.replace(" ", "").replace("(", "").replace(")", "")
    download_name = f"固定质量问题清单_制冰机_{cur_year}_{_dl_tag}.csv"
else:
    df_scope = df_year_only
    title_suffix = f"{cur_year}全年"
    download_name = f"固定质量问题清单_制冰机_{cur_year}.csv"

st.session_state['store_df'] = ensure_cols(df_all.copy())

fixed_table = build_fixed_issue_table(df_scope, FIXED_DEFECTS)
if fixed_table.empty:
    st.info(f"{title_suffix} 无数据可统计。")
else:
    _months_param = ",".join(str(m) for m in _sel_months)
    _q_name = None
    if set(_sel_months) == {1,2,3}:   _q_name = "Q1"
    elif set(_sel_months) == {4,5,6}: _q_name = "Q2"
    elif set(_sel_months) == {7,8,9}: _q_name = "Q3"
    elif set(_sel_months) == {10,11,12}: _q_name = "Q4"

    def _build_link(code: str) -> str:
        base = (page_url_by_code(__file__, str(code)) or "#")
        if _q_name:
            return f"{base}?code={quote(str(code))}&year={int(cur_year)}&months={_months_param}&q={_q_name}"
        else:
            return f"{base}?code={quote(str(code))}&year={int(cur_year)}&months={_months_param}"

    fixed_table["配件分析"] = fixed_table["_code"].apply(_build_link)

    st.caption(f"统计口径：{title_suffix}")
    st.data_editor(
        fixed_table[["项目（编码+英文）","问题中文名","数量","费用","配件分析"]],
        use_container_width=True, height=520, disabled=True,
        column_config={"配件分析": st.column_config.LinkColumn("配件分析", display_text="打开配件分析")}
    )
    st.download_button(
        "下载 CSV",
        data=fixed_table.drop(columns=["_fee_val","_code"], errors="ignore").to_csv(index=False, encoding="utf-8-sig"),
        file_name=download_name,
        mime="text/csv",
        use_container_width=True
    )

# ============== 原始 Excel 表映射（制冰机专用历史） ==============
def _filter_excel_by_mm_yyyy(df_hist: pd.DataFrame, y: int, m: int) -> pd.DataFrame:
    if df_hist is None or df_hist.empty:
        return pd.DataFrame(columns=(df_hist.columns if df_hist is not None else []))

    pat = re.compile(r'(?:^|\D)(?P<mm>0[1-9]|1[0-2])(?P<yyyy>20\d{2})(?:\D|$)')
    hits = []
    for col in df_hist.columns:
        if col in {"__SourceFile", "__ImportedAt"}:
            continue
        ext = df_hist[col].astype(str).str.extract(pat)
        if ext.empty or (ext.isna().all().all()):
            continue
        yy = pd.to_numeric(ext["yyyy"], errors="coerce")
        mm = pd.to_numeric(ext["mm"], errors="coerce")
        mask = (yy == int(y)) & (mm == int(m))
        if mask.any():
            hits.append(df_hist.loc[mask, df_hist.columns])

    if not hits:
        return pd.DataFrame(columns=df_hist.columns)

    out = pd.concat(hits, ignore_index=True)
    key_cols = [c for c in out.columns if not str(c).startswith("__")]
    if key_cols:
        sig = out[key_cols].astype(str).agg("|".join, axis=1).str.lower()
        out = out.loc[~sig.duplicated(keep="first")].reset_index(drop=True)
    return out

if is_admin():
    with st.expander("📎 原始 Excel 表映射（原样显示“生产日期”，无 Year/Month）", expanded=False):
        _hist = st.session_state.get("excel_raw_history", pd.DataFrame())
        if _hist is None or _hist.empty:
            st.info("暂无原始上传历史。请先上传 Excel（制冰机专用通道，文件独立保存）。")
        else:
            _excel_filtered = _filter_excel_by_mm_yyyy(_hist, int(cur_year), int(edit_month))
            st.caption(f"显示口径：{cur_year}-{edit_month:02d}（自动从所有列中识别 MMYYYY）")

            if _excel_filtered.empty:
                st.info("该月在原始历史中没有匹配到数据。请更换年月或检查数据中的日期文本是否包含 MMYYYY。")
            else:
                cols = [c for c in _excel_filtered.columns if c not in {"__SourceFile","__ImportedAt"}]
                cols += [c for c in ["__SourceFile","__ImportedAt"] if c in _excel_filtered.columns]
                _view = _excel_filtered[cols]

                def _fmt_prod_date_cell(v):
                    s = "" if pd.isna(v) else str(v).strip()
                    m6 = re.fullmatch(r"(\d{2})(\d{2})(\d{2})", s)
                    if m6:
                        yy, mm, dd = m6.groups()
                        return f"{2000 + int(yy)}/{mm}/{dd}"
                    m8 = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s)
                    if m8:
                        y0, mm, dd = m8.groups()
                        return f"{int(y0)}/{mm}/{dd}"
                    try:
                        dt = pd.to_datetime(s, errors="coerce")
                        if pd.isna(dt):
                            return v
                        return f"{dt.year}/{dt.month}/{dt.day}"
                    except Exception:
                        return v

                prod_cols = [c for c in _view.columns if re.search(r"生产\s*日期", str(c), re.I)]
                _view_display = _view.copy()
                for c in prod_cols:
                    _view_display[c] = _view_display[c].map(_fmt_prod_date_cell)

                _n = int(min(len(_view_display), 16))
                _h = int(46 + 28 * max(_n, 1))
                st.dataframe(_view_display, use_container_width=True, height=min(_h, 600))

# ============== 调试区域（*_ice 状态） ==============
with st.expander("调试：本月关键字命中统计（当选择季度时，这里默认用该期间最后一个月）", expanded=False):
    df_month_real = df_all[(_year_series==cur_year) & (_month_series==edit_month)]
    if not df_month_real.empty:
        joined=df_month_real.astype(str).agg(" ".join, axis=1).map(_clean_text)
        cnt_21a = int(joined.str.contains(re.compile(r"2\.1\.a", re.I)).sum())
        cnt_atosa=int(joined.str.contains(re.compile(r"\batosa\b", re.I)).sum())
        real_cat=df_month_real[(df_month_real["Category1"]=="质量问题") & (df_month_real["Category2"]=="Atosa")]
        st.write({"2.1.a 命中数(文本)":cnt_21a,"atosa 命中数(文本)":cnt_atosa,"最终分类到『质量问题/Atosa』":len(real_cat)})
    else:
        st.write("本月无数据")

with st.expander("调试：来源与口径", expanded=False):
    if not df_all.empty:
        st.write("总记录数（已按 2025–2030 过滤后）：", len(df_all))
        if "_SrcYM" in df_all.columns:
            st.write("_SrcYM 分布（Top 10）："); st.write(df_all["_SrcYM"].value_counts().head(10))
        st.write("当月记录数：", len(df_all[(_year_series==cur_year)&(_month_series==edit_month)]), " | 全年记录数：", len(df_year_only))
        st.write("最近一次上传文件 MD5（制冰机通道）：", st.session_state.get("_last_upload_md5_ice","无"))
    else:
        st.write("暂无数据")

with st.expander("🔎 跳转自检", expanded=False):
    st.write("映射表：", {k: v for k, v in _FORCE_CODE2PAGE.items()})
    for code, target in _FORCE_CODE2PAGE.items():
        via_pages = _find_page_url(target)
        guessed   = _guess_url_from_filename(target)
        final     = page_url_by_code(__file__, code)
        st.write(f"{code} → {target} → get_pages: {via_pages} | guessed: {guessed} | final: {final}")
