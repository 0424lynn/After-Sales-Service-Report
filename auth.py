# -*- coding: utf-8 -*-
import json
import time
import hmac
import hashlib
import secrets
from pathlib import Path
from typing import Dict, Any, Optional

import streamlit as st

# ===== 应用信息（登录页抬头） =====
APP_NAME = "Atosa-美国售后报告"
APP_SUBTITLE = "售后报告 · 数据看板 · 质量洞察"

# ===== 存储与会话 =====
USERS_JSON = Path("users.json")
SESSION_TTL_SECONDS = 8 * 3600  # 登录有效期 8 小时

# ===== 密码哈希（PBKDF2-SHA256） =====
def _pbkdf2_hash(password: str, salt: str, rounds: int = 200_000) -> str:
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), rounds, dklen=32)
    return dk.hex()

def hash_password(password: str) -> tuple[str, str]:
    salt = secrets.token_hex(16)
    return salt, _pbkdf2_hash(password, salt)

def verify_password(stored_salt: str, stored_hash: str, plain: str) -> bool:
    test_hash = _pbkdf2_hash(plain or "", stored_salt or "")
    return hmac.compare_digest(stored_hash or "", test_hash)

# ===== 用户存取 =====
def load_users() -> Dict[str, Any]:
    if USERS_JSON.exists():
        try:
            return json.loads(USERS_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass

    # 首次运行：写入默认用户（请尽快修改）
    admin_salt, admin_hash = hash_password("admin123")
    analyst_salt, analyst_hash = hash_password("analyst123")
    default = {
        "users": [
            {"username": "admin",   "name": "管理员", "role": "admin",   "salt": admin_salt,   "password_hash": admin_hash},
            {"username": "analyst", "name": "分析员", "role": "analyst", "salt": analyst_salt, "password_hash": analyst_hash},
        ]
    }
    try:
        USERS_JSON.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print("WARN: failed to write users.json:", e)
    return default

def save_users(data: Dict[str, Any]) -> None:
    try:
        USERS_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as e:
        print("WARN: failed to write users.json:", e)

def get_user(username: str) -> Optional[Dict[str, Any]]:
    data = load_users()
    for u in data.get("users", []):
        if u.get("username") == username:
            return u
    return None

# ===== 登录态便捷函数 =====
def current_user() -> Optional[Dict[str, Any]]:
    """返回 {'username','name','role','login_at'} 或 None"""
    return st.session_state.get("user")

def is_admin() -> bool:
    u = current_user() or {}
    return str(u.get("role", "")).strip().lower() in {"admin", "administrator", "root"}

def require_role(*roles: str) -> bool:
    u = current_user() or {}
    return u.get("role") in set(roles)

# ===== 样式（隐藏侧栏 + 登录卡片） =====
def hide_sidebar_css():
    st.markdown(
        """
        <style>
        aside[aria-label="sidebar"], [data-testid="stSidebar"] { display:none !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

def _inject_login_css():
    st.markdown(
        """
        <style>
        /* 背景与文字主色 */
        html, body, [data-testid="stAppViewContainer"] {
            background: linear-gradient(135deg, #0a192f 0%, #1e3a8a 55%, #0a192f 100%) !important;
            color: #e2e8f0 !important;
        }
        /* 主容器：限制宽度 + 垂直居中 */
        .block-container {
            max-width: 480px !important;
            padding-top: 10vh !important;
            padding-bottom: 8vh !important;
        }
        /* 标题与副标题（更大） */
        h1.atosa-title {
            text-align: center;
            font-size: 28px;           /* ← 加大 */
            line-height: 1.25;
            color: #f8fafc;
            margin-bottom: 8px;
            font-weight: 800;          /* ← 更粗 */
            letter-spacing: .4px;
        }
        p.atosa-subtitle {
            text-align: center;
            color: #cbd5e1;
            font-size: 13.5px;
            margin-top: 0;
            margin-bottom: 16px;
        }

        /* 卡片外层：给一个唯一作用域，避免主题样式覆盖 */
        #atosa-card {
            background: rgba(255,255,255,0.08);
            border: 1px solid rgba(255,255,255,0.15);
            border-radius: 16px;
            padding: 18px 16px 16px;
            box-shadow: 0 10px 28px rgba(0,0,0,.45);
            backdrop-filter: blur(8px);
        }

        /* 仅在卡片作用域内隐藏标签，让输入框与卡片齐平 */
        #atosa-card .stTextInput > label,
        #atosa-card .stPassword   > label {
            display: none !important;
        }

        /* 输入框：100% 宽、与卡片齐平 */
        #atosa-card .stTextInput input,
        #atosa-card .stPassword   input {
            width: 100% !important;
            background: rgba(15,23,42,.92) !important;
            color: #e5e7eb !important;
            border: 1px solid rgba(255,255,255,.18) !important;
            border-radius: 10px !important;
            height: 40px !important;
            padding: 0 12px !important;
        }

        /* 按钮：100% 宽，与输入框对齐 */
        #atosa-card .stButton > button {
            width: 100% !important;
            height: 40px !important;
            border-radius: 10px !important;
            font-weight: 700 !important;
            background: linear-gradient(90deg,#ef4444,#f97316) !important;
            color: #fff !important;
            border: none !important;
        }
        #atosa-card .stButton > button:hover {
            background: linear-gradient(90deg,#f97316,#ea580c) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# ===== 登录/退出/权限 =====
def _session_valid() -> bool:
    u = st.session_state.get("user")
    if not u or not st.session_state.get("auth_ok"):
        return False
    if int(time.time()) - int(u.get("login_at", 0)) > SESSION_TTL_SECONDS:
        return False
    return True

def login_form():
    """渲染登录卡片：居中、小尺寸、输入框与卡片齐平"""
    hide_sidebar_css()
    _inject_login_css()

    # 标题
    st.markdown(f"<h1 class='atosa-title'>{APP_NAME}</h1>", unsafe_allow_html=True)
    st.markdown(f"<p class='atosa-subtitle'>{APP_SUBTITLE}</p>", unsafe_allow_html=True)

    # 表单（用占位符 + 折叠标签，让输入框与卡片齐平）
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("用户名", placeholder="用户名", label_visibility="collapsed", key="login_user_input")
        password = st.text_input("密码", placeholder="密码", type="password", label_visibility="collapsed", key="login_pwd_input")
        show_pw = st.checkbox("显示密码", value=False, help="勾选后可查看明文密码")
        if show_pw and st.session_state.get("login_pwd_input"):
            # 切换为明文展示（只在界面层显示，不影响验证）
            st.text_input("明文密码", value=st.session_state["login_pwd_input"], label_visibility="collapsed", disabled=True)
        submitted = st.form_submit_button("登录", type="primary")

    if submitted:
        user = get_user((username or "").strip())
        ok = bool(user) and verify_password(user["salt"], user["password_hash"], password or "")
        if not ok:
            st.error("用户名或密码错误")
        else:
            # 建立会话
            st.session_state["auth_ok"] = True
            st.session_state["user"] = {
                "username": user["username"],
                "name": user.get("name") or user["username"],
                "role": user.get("role", "viewer"),
                "login_at": int(time.time()),
            }
            st.success("登录成功")
            # 优先跳转到“可视化”页；无法定位则刷新
            try:
                st.switch_page("可视化.py")
            except Exception:
                try:
                    st.switch_page("pages/可视化.py")
                except Exception:
                    st.rerun()
        st.stop()

    # 首次使用提醒
    if not USERS_JSON.exists():
        st.info("首次使用：默认账户 **admin/admin123**、**analyst/analyst123**，请登录后尽快在 users.json 中修改密码。")

def logout_button(label: str = "退出登录"):
    """放在侧栏或页面任何位置都可用"""
    u = st.session_state.get("user")
    if u:
        with st.sidebar:
            st.write(f"👤 {u.get('name','')}  ·  角色：{u.get('role','')}")
            if st.button(label, use_container_width=True):
                for k in ("auth_ok", "user"):
                    st.session_state.pop(k, None)
                st.success("已退出")
                st.rerun()

def require_login():
    """在每个页面顶部调用；未登录则展示登录卡片并停止渲染。"""
    if _session_valid():
        return
    # 会话过期或未登录：清理并拉起登录
    st.session_state.pop("auth_ok", None)
    st.session_state.pop("user", None)
    login_form()
    st.stop()
