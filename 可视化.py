# ==== 放在 App.py 顶部（所有 st.* 之前） ====
import os, re, sys
import streamlit as st

# 1) 自检：确保从 App.py 启动（取消提示，不再弹 warning）
try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx
    _main = getattr(get_script_run_ctx(), "main_script_path", "") or os.path.abspath(sys.argv[0])
except Exception:
    _main = os.path.abspath(sys.argv[0])

_main_norm = os.path.basename(_main).lower().strip()
# 仅检测，不提示
# if _main_norm not in {"app.py", "./app.py"}:
#     st.warning("⚠️ 检测到主脚本路径包含特殊字符或大小写差异，但程序已自动继续执行。")
# else:
#     pass

# 2) 构建并移除需要隐藏的 pages
def _remove_analysis_pages():
    try:
        from streamlit.source_util import get_pages
    except Exception:
        return

    pages = get_pages(_main) or {}
    if not pages:
        return

    # 只要写成 pages/xxx.py 就稳
    need_hide = {
        "pages/配件分析.py",
        "pages/风机分析.py",
        "pages/温控器分析.py",
        "pages/压缩机分析.py",
        "pages/外部漏水分析.py",
        "pages/泄氟分析.py",
        "pages/蒸发器线圈结冰分析.py",
    }
    need_hide = {p.lower() for p in need_hide}
    name_rx = re.compile(r".*分析$")   # 页面名以“分析”结尾的一律干掉

    for k, d in list(pages.items()):
        sp = str(d.get("script_path", "")).replace("\\", "/").lower()
        name = str(d.get("page_name", "")).strip()
        if any(sp.endswith(t) for t in need_hide) or name_rx.fullmatch(name):
            pages.pop(k, None)

_remove_analysis_pages()

# 3) DOM 兜底（防止后续重绘把它们又画出来）
import streamlit.components.v1 as components
components.html("""
<script>
(function(){
  const badText = [/分析$/,'配件分析','风机分析','温控器分析','压缩机分析','外部漏水分析','泄氟分析','蒸发器线圈结冰分析'];
  function hideNav(){
    const nav = parent.document.querySelector('[data-testid="stSidebarNav"]') 
            || parent.document.querySelector('nav[aria-label="Sidebar navigation"]');
    if(!nav) return;
    nav.querySelectorAll('a,button,li').forEach(el=>{
      const t=(el.textContent||'').trim();
      if(badText.some(b => b instanceof RegExp ? b.test(t) : t.includes(b))){
        const li = el.closest('li') || el;
        li.style.display='none';
        li.setAttribute('data-hidden-by','hide-analysis');
      }
    });
  }
  hideNav();
  let ticks = 0;
  const mo = new MutationObserver(()=>{ hideNav(); if(++ticks>200) mo.disconnect(); });
  mo.observe(parent.document.body, {subtree:true, childList:true});
})();
</script>
""", height=0)

# ==== 顶部隐藏段落结束，下面才写 st.set_page_config / 页面逻辑 ====

# =========================
# 2025 目标维修率 — 可视化（置顶）
# =========================
import pandas as pd
import numpy as np
import re, json
from pathlib import Path

# 让可视化页更“顶”，并宽屏显示
try:
    st.set_page_config(page_title="可视化", layout="wide", initial_sidebar_state="collapsed")
except Exception:
    pass

st.markdown("## 📈 2025 目标维修率（冰箱数据）")

# ---- 读本地持久化（与“冰箱数据”页同口径）
STORE_PARQUET = Path("data_store.parquet")
STORE_CSV     = Path("data_store.csv")
TARGETS_JSON  = Path("targets_config.json")
MODEL_FLEET_JSON = Path("model_fleet_counts.json")

def _read_store_df():
    if "store_df" in st.session_state and not st.session_state["store_df"].empty:
        return st.session_state["store_df"].copy()
    # 兜底：直接读持久化文件
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

def _load_targets():
    if TARGETS_JSON.exists():
        try:
            return json.loads(TARGETS_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"target_year_rate": 5.0, "q1_t":0.0,"q2_t":0.0,"q3_t":0.0,"q4_t":0.0,"year_t":5.0}

def _load_model_fleet():
    if MODEL_FLEET_JSON.exists():
        try:
            return json.loads(MODEL_FLEET_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}

def _clean_text(s: str)->str:
    import re
    return re.sub(r"[\s\-_\/]+", " ", ("" if pd.isna(s) else str(s)).lower()).strip()

# 与冰箱数据页一致：识别 1.1.a（质量问题归口）
_PAT_11A = re.compile(r"(?<![0-9a-z])1\.1\.a(?![0-9a-z])", re.I)

def _ensure_cols(df: pd.DataFrame) -> pd.DataFrame:
    """最小映射：保证有 Year/Month/CostUSD 和文本拼接口径（与冰箱数据页兼容）"""
    if df.empty:
        return pd.DataFrame(columns=["Year","Month","CostUSD","Category1","Category2","TAG1","TAG2","TAG3","TAG4","Summary","Notes"])
    # 统一常见列名
    mapping = {}
    for c in df.columns:
        lc = str(c).strip().lower()
        if lc in ["cost","costusd","amount","金额","费用","fee","cost(usd)","cost usd"]:
            mapping[c] = "CostUSD"
        elif lc in ["year","年份"]:
            mapping[c] = "Year"
        elif lc in ["month","月份"]:
            mapping[c] = "Month"
    if mapping:
        df = df.rename(columns=mapping)
    if "CostUSD" not in df.columns:
        df["CostUSD"] = 0.0
    df["CostUSD"] = pd.to_numeric(df["CostUSD"], errors="coerce").fillna(0.0)

    # 若没有 Year/Month，但第一列里有 MMYYYY 的串，这里做一个轻解析
    if ("Year" not in df.columns) or ("Month" not in df.columns) or pd.to_numeric(df.get("Year"), errors="coerce").isna().all() or pd.to_numeric(df.get("Month"), errors="coerce").isna().all():
        _MMYYYY_ANY = re.compile(r'(?:^|\\D)(?P<mm>0[1-9]|1[0-2])(?P<yyyy>20\\d{2})(?:\\D|$)')
        s = df[df.columns[0]].astype(str)
        m = s.str.extract(_MMYYYY_ANY)
        mm = pd.to_numeric(m["mm"], errors="coerce")
        yy = pd.to_numeric(m["yyyy"], errors="coerce")
        df["Year"] = yy
        df["Month"] = mm
    return df

def _get_warranty_fleet_total(fleet: dict, year: int, month: int) -> int:
    y, m = str(year), str(month)
    try:
        mon = fleet.get(y, {}).get(m, {})
        # 该 JSON 结构为 {年: {月: {型号系列: 数量}}}
        return int(sum(int(v) for v in mon.values()))
    except Exception:
        return 0

# ---- 取数据
df_all = _ensure_cols(_read_store_df())
targets = _load_targets()
fleet   = _load_model_fleet()

if df_all.empty:
    st.info("暂无数据：请先在『冰箱数据』页上传并保存，再回到此页查看可视化。")
else:
    # 年份选择（默认 2025；同时支持读取冰箱页侧栏的 year）
    cand_years = sorted([int(y) for y in pd.to_numeric(df_all["Year"], errors="coerce").dropna().unique() if 2025 <= int(y) <= 2030])
    default_year = 2025 if 2025 in cand_years else (cand_years[0] if cand_years else 2025)
    cur_year = int(st.session_state.get("year", default_year))
    cur_year = cur_year if cur_year in cand_years else default_year

    st.caption(f"统计口径：{cur_year} 年（与『冰箱数据』页保持一致）")

    # 当年数据
    yy = pd.to_numeric(df_all["Year"], errors="coerce")
    mm = pd.to_numeric(df_all["Month"], errors="coerce")
    df_year = df_all[(yy == cur_year) & (mm.between(1, 12))].copy()

    # 拼接文本供 1.1.a 匹配
    text_cols = [c for c in ["Category1","Category2","TAG1","TAG2","TAG3","TAG4","Summary","Notes"] if c in df_year.columns]
    if not text_cols:
        text_cols = df_year.columns.tolist()
    joined_all = (df_year[text_cols].astype(str).agg(" ".join, axis=1).map(_clean_text)) if not df_year.empty else pd.Series([], dtype=str)

    # 逐月计算
    months = list(range(1, 12 + 1))
    data = []
    for m in months:
        dfm = df_year[pd.to_numeric(df_year["Month"], errors="coerce") == m]
        # 平均维修单费用（只统计有费用>0的工单）
        costs = pd.to_numeric(dfm["CostUSD"], errors="coerce").fillna(0.0)
        avg_cost = float(costs[costs > 0].mean()) if not dfm.empty and (costs > 0).any() else 0.0
        # 质量问题数量（按 1.1.a 关键字命中，复用你的口径）
        if not dfm.empty:
            idx = dfm.index
            qual_cnt = int(joined_all.loc[idx].str.contains(_PAT_11A).sum())
        else:
            qual_cnt = 0
        # 保内数量（来自在保量 JSON 的当月合计）
        warranty_cnt = _get_warranty_fleet_total(fleet, cur_year, m)
        # 折算保内维修率 %
        rate_pct = (qual_cnt / warranty_cnt * 100.0) if warranty_cnt > 0 else 0.0

        data.append({
            "月份": m,
            "月度保内数量": int(warranty_cnt),
            "平均维修单费用（美元）": float(avg_cost),
            "折算保内维修率%": float(rate_pct),
        })

    df_vis = pd.DataFrame(data)

    # ---- 同一行 3 张图：数量（柱）、维修率（含目标线）、平均费用（线）
    col1, col2, col3 = st.columns(3, gap="large")

    with col1:
        st.markdown("#### 🧮 月度保内数量")
        st.bar_chart(df_vis.set_index("月份")[["月度保内数量"]])

    with col2:
        st.markdown("#### 🎯 折算保内维修率 %（含年度目标线）")
        target_pct = float(targets.get("year_t", targets.get("target_year_rate", 5.0)))
        df_rate = df_vis.set_index("月份")[["折算保内维修率%"]].copy()
        df_rate["年度目标%"] = target_pct  # 常数线
        st.line_chart(df_rate)

    with col3:
        st.markdown("#### 💵 平均维修单费用（美元）")
        st.line_chart(df_vis.set_index("月份")[["平均维修单费用（美元）"]])

    # 可选：展开明细数据用于校核（放在图下方）
    with st.expander("查看明细数据（用于校核）", expanded=False):
        st.dataframe(df_vis, use_container_width=True)
    # ==== 质量问题明细：TOP6 数量对比（全年） ====
st.markdown("### 🧩 质量问题 TOP6（按数量，全年）")

# 常见固定缺陷清单（可按需扩充/调整名称）
_DEFECTS = [
    {"code":"1.1.3.5",  "en":"Fan Motor",                       "cn":"风机"},
    {"code":"1.1.3.3",  "en":"Temp Controller",                 "cn":"温控器"},
    {"code":"1.1.3.1",  "en":"Leaking Refrigerant",             "cn":"泄氟"},
    {"code":"1.1.3.2",  "en":"Compressor",                      "cn":"压缩机"},
    {"code":"1.1.5.2",  "en":"Outside Leak",                    "cn":"外部漏水"},
    {"code":"1.1.5.1",  "en":"Inside Leak",                     "cn":"内部漏水"},
    {"code":"1.1.3.4",  "en":"Evap Coil/Drain Ice Up",          "cn":"蒸发器/排水结冰"},
    {"code":"1.1.3.9",  "en":"Cap Tube Restriction",            "cn":"毛细管堵塞"},
    {"code":"1.1.11.2", "en":"Gasket",                          "cn":"门封条"},
    {"code":"1.1.7.1",  "en":"Hinge/Spring Bar",                "cn":"门铰链/门弹簧"},
    {"code":"1.1.4.3",  "en":"LED Strip/LED Light",             "cn":"LED 灯"},
    {"code":"1.1.3.10", "en":"Fan blade blocked by foreign obj","cn":"风机叶片被异物卡住"},
]

# 选全年数据（与上文 df_year 一致）
df_scope = df_year.copy()

def _count_defects(df_scope: pd.DataFrame, defects: list, joined_all: pd.Series) -> pd.DataFrame:
    import re
    if df_scope.empty:
        return pd.DataFrame(columns=["code","问题","数量"])
    rows = []
    idx = df_scope.index
    joined = joined_all.loc[idx]
    for d in defects:
        code = d["code"].strip()
        pat  = re.compile(rf"(?<![0-9a-z]){re.escape(code)}(?![0-9a-z])", re.I)
        cnt  = int(joined.str.contains(pat).sum())
        if cnt > 0:
            label = f"{code} {d.get('cn') or d.get('en')}"
            rows.append({"code": code, "问题": label, "数量": cnt})
    if not rows:
        return pd.DataFrame(columns=["code","问题","数量"])
    out = pd.DataFrame(rows).sort_values("数量", ascending=False).head(6).reset_index(drop=True)
    return out

df_top6 = _count_defects(df_scope, _DEFECTS, joined_all)

if df_top6.empty:
    st.info(f"{cur_year} 年暂无可统计的缺陷编码数据。")
else:
    # 画柱状图（按数量降序）
    chart_df = df_top6.set_index("问题")[["数量"]]
    st.bar_chart(chart_df)

    # 小结
    st.caption(f"注：年度目标读取自 `targets_config.json`（year_t={float(targets.get('year_t', targets.get('target_year_rate', 5.0))):.2f}%）。如需调整，请在『冰箱数据』页侧栏进行配置。")
# ==== 各主要质量问题（月度趋势：复选框显示/隐藏） ====
st.markdown("### 📊 质量问题月度趋势（TOP6，可勾选显示/隐藏）")

if df_top6.empty:
    st.info(f"{cur_year} 年暂无可视化数据。")
else:
    import re

    top_codes = df_top6["code"].tolist()[:6]
    code2name = {r["code"]: r["问题"] for _, r in df_top6.iterrows()}
    months = range(1, 13)

    # 预先统计 1-12 月数量
    trend_rows = []
    for code in top_codes:
        pat = re.compile(rf"(?<![0-9a-z]){re.escape(code)}(?![0-9a-z])", re.I)
        for m in months:
            dfm = df_year[pd.to_numeric(df_year["Month"], errors="coerce") == m]
            if dfm.empty:
                cnt = 0
            else:
                idx = dfm.index
                joined_m = joined_all.loc[idx]
                cnt = int(joined_m.str.contains(pat).sum())
            trend_rows.append({"月份": m, "问题": code2name[code], "数量": cnt})

    df_trend = pd.DataFrame(trend_rows)
    df_pivot = df_trend.pivot(index="月份", columns="问题", values="数量").fillna(0).astype(int)

    # 勾选要显示的曲线（默认全选）
    st.caption("勾选要显示的问题：")
    cols = st.columns(3)
    selected = []
    names = list(df_pivot.columns)
    for i, name in enumerate(names):
        with cols[i % 3]:
            if st.checkbox(name, value=True, key=f"show_{i}"):
                selected.append(name)

    if not selected:
        st.info("未选择任何问题，请至少勾选一个。")
    else:
        st.line_chart(df_pivot[selected], use_container_width=True)

