# auth.py
# -*- coding: utf-8 -*-
import json, hashlib, os
from pathlib import Path
import streamlit as st

USERS_JSON = Path("users.json")

# ---------- 基础：加盐哈希 ----------
def hash_password(plain: str):
    if not isinstance(plain, str):
        plain = str(plain or "")
    salt = os.urandom(16).hex()
    pw_hash = hashlib.sha256((salt + plain).encode("utf-8")).hexdigest()
    return salt, pw_hash

def verify_password(plain: str, salt: str, pw_hash: str) -> bool:
    if not salt or not pw_hash:
        return False
    try:
        calc = hashlib.sha256((str(salt) + str(plain)).encode("utf-8")).hexdigest()
        return calc == str(pw_hash)
    except Exception:
        return False

# ---------- 用户存取 ----------
def _init_default_users() -> dict:
    # 默认内置 admin/admin123、analyst/analyst123
    salt1, hash1 = hash_password("admin123")
    salt2, hash2 = hash_password("analyst123")
    default = {
        "users": [
            {"username": "admin",   "name": "管理员", "role": "admin",   "salt": salt1, "password_hash": hash1},
            {"username": "analyst", "name": "分析员", "role": "analyst", "salt": salt2, "password_hash": hash2},
        ]
    }
    try:
        USERS_JSON.write_text(json.dumps(default, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return default

def load_users() -> dict:
    if USERS_JSON.exists():
        try:
            return json.loads(USERS_JSON.read_text(encoding="utf-8"))
        except Exception:
            pass
    return _init_default_users()

def save_users(data: dict):
    try:
        USERS_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

# ---------- 会话 & 权限 ----------
def current_user():
    # 登录成功后：{'username':..., 'name':..., 'role':...}
    return st.session_state.get("_login_user")

def is_admin() -> bool:
    u = current_user() or {}
    role = str(u.get("role", "")).strip().lower()
    return role in {"admin", "administrator", "root"}

def login_status_badge():
    u = current_user()
    if u:
        st.caption(f"已登录：**{u.get('name') or u.get('username')}**（{u.get('role')}）")
    else:
        st.caption("未登录")

# ---------- UI：隐藏侧栏 ----------
def hide_sidebar_css():
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"] {display:none !important;}
        </style>
        """,
        unsafe_allow_html=True,
    )

# ---------- UI：登录表单 ----------
def login_form():
    st.header("登录")
    users = load_users()
    usernames = [u["username"] for u in users.get("users", [])]
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("用户名", value="")
        password = st.text_input("密码", type="password", value="")
        submitted = st.form_submit_button("登录", type="primary", use_container_width=True)

    if submitted:
        if not username or not password:
            st.error("请输入用户名和密码")
            return

        user = next((u for u in users.get("users", []) if u.get("username") == username), None)
        if user and verify_password(password, user.get("salt",""), user.get("password_hash","")):
            # ✅ 登录成功：写入会话
            st.session_state["_login_user"] = {
                "username": user.get("username"),
                "name": user.get("name") or user.get("username"),
                "role": user.get("role") or "analyst",
            }
            st.success("登录成功")
            st.rerun()
        else:
            st.error("用户名或密码错误")

    # 首次使用强提醒
    if not USERS_JSON.exists():
        st.info("首次使用：默认账户 **admin/admin123**，请登录后尽快在 users.json 中修改密码。")

# ---------- UI：退出 ----------
def logout_button():
    if st.button("退出登录", use_container_width=True):
        if "_login_user" in st.session_state:
            del st.session_state["_login_user"]
        st.success("已退出")
        st.rerun()
