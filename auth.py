# auth.py
# -*- coding: utf-8 -*-
import json
import time
import hmac
import hashlib
import secrets
from pathlib import Path
from typing import Dict, Any, Optional

import streamlit as st

# ========= 应用信息（登录页抬头） =========
APP_NAME = "Atosa-美国售后报告分析"
APP_SUBTITLE = "售后报告 · 数据看板 · 质量洞察"

# ========= 存储与会话 =========
USERS_JSON = Path("users.json")
SESSION_TTL_SECONDS = 8 * 3600  # 登录有效期 8 小时

# ========= 密码哈希（PBKDF2-SHA256） =========
def _pbkdf2_hash(password: str, salt: str, rounds: int = 200_000) -> str:
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), rounds, dklen=32)
    return dk.hex()

def hash_password(password: str) -> tuple[str, str]:
    salt = secrets.token_hex(16)
    return salt, _pbkdf2_hash(password, salt)

def verify_password(stored_salt: str, stored_hash: str, plain: str) -> bool:
    test_hash = _pbkdf2_hash(plain or "", stored_salt or "")
    return hmac.compare_digest(stored_hash or "", test_hash)

# ========= 用户存取 =========
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

# ========= 登录态便捷函数 =========
def current_user() -> Optional[Dict[str, Any]]:
    """返回 {'username','name','role','login_at'} 或 None"""
    return st.session_state.get("user")

def is_admin() -> bool:
    u = current_user() or {}
    return str(u.get("role", "")).strip().lower() in {"admin", "administrator", "root"}

def require_role(*roles: str) -> bool:
    u = current_user() or {}
    return u.get("role") in set(roles)

# ========= 样式（隐藏侧栏 + 登录卡片） =========
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
            max-width: 420px !important;
            padding-top: 12vh !important;
            padding-bottom: 8vh !important;
        }
        /* 表单做成小卡片 */
        form[data-testid="stForm"] {
            background: rgba(255,255,255,0.08);
            border: 1px solid rgba(255,255,255,0.15);
            border-radius: 16px;
            padding: 22px 18px 18px;
            box-shadow: 0 10px 28px rgba(0,0,0,.45);
            backdrop-filter: blur(8px);
        }
        /* 标题与副标题 */
        h1.atosa-title {
            text-align: center;
            font-size: 20px;
            color: #f8fafc;
            margin-bottom: 6px;
            font-weight: 700;
            letter-spacing: .2px;
        }
        p.atosa-subtitle {
            text-align: center;
            color: #94a3b8;
            font-size: 12.5px;
            margin-top: 0;
            margin-bottom: 18px;
        }
        /* 输入框外观 */
        .stTextInput input, .stPassword input {
            background: rgba(15,23,42,.85) !important;
            color: #e5e7eb !important;
            border: 1px solid rgba(255,255,255,.18) !important;
            border-radius: 10px !important;
            height: 38px !important;
        }
        /* 提交按钮 */
        .stButton > button {
            width: 100% !important;
            height: 38px !important;
            border-radius: 10px !important;
            font-weight: 600 !important;
            background: linear-gradient(90deg,#3b82f6,#2563eb) !important;
            color: #fff !important;
            border: none !important;
        }
        .stButton > button:hover {
            background: linear-gradient(90deg,#2563eb,#1d4ed8) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# ========= 登录/退出/权限 =========
def _session_valid() -> bool:
    u = st.session_state.get("user")
    if not u or not st.session_state.get("auth_ok"):
        return False
    if int(time.time()) - int(u.get("login_at", 0)) > SESSION_TTL_SECONDS:
        return False
    return True

def login_form():
    """渲染登录卡片：居中、小尺寸、带显示密码开关"""
    hide_sidebar_css()
    _inject_login_css()

    # 标题
    st.markdown(f"<h1 class='atosa-title'>{APP_NAME}</h1>", unsafe_allow_html=True)
    st.markdown(f"<p class='atosa-subtitle'>{APP_SUBTITLE}</p>", unsafe_allow_html=True)

    # 表单（不再用自定义外层 div 包裹，避免控件错位）
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("用户名", key="login_user_input")
        col1, col2 = st.columns([1, 1])
        with col1:
            show_pw = st.checkbox("显示密码", value=False, key="login_show_pw")
        with col2:
            st.write("")  # 对齐
        password = st.text_input("密码", type=("default" if show_pw else "password"), key="login_pwd_input")
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
