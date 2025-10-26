# ==== 顶部追加：新会话（刷新/直接打开子页）时，回到“可视化” ====
import streamlit as st

# 说明：
# - Streamlit 每次“硬刷新/新打开标签页”都会创建新的会话（session），此时 session_state 为空。
# - 我们用一个标记 _first_load_done 来识别“本会话的第一次加载”：
#   1) 如果用户刷新或直接打开子页（新的会话），标记不存在 -> 立即跳去“可视化”，并把标记设为 True。
#   2) 如果是站内点击导航（仍在同一会话），标记已经存在 -> 不跳转，正常停留在当前子页。

if "_first_load_done" not in st.session_state:
    st.session_state["_first_load_done"] = True
    # 尝试用脚本名跳；若路径不同再兜底用页面名
    try:
        st.switch_page("可视化.py")
    except Exception:
        try:
            st.switch_page("pages/可视化.py")
        except Exception:
            # 最后兜底：用一个超轻的前端跳转，避免路径差异
            import streamlit.components.v1 as components
            components.html("<script>window.parent.location.pathname = decodeURIComponent(window.parent.location.pathname).replace(/[^/]+$/, '可视化');</script>", height=0)
# ==== 顶部追加结束 ====

