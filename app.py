# -*- coding: utf-8 -*-
"""价值观考核 · 在线填写（独立模块）

流程：
  1. 员工打开入口页，填手机号 + 姓名 + 选部门，直接开始自评（不用 HR 逐个建单）
  2. HR 在后台维护「部门树」，并给每个部门指定主管（子部门没指定就继承上级）
  3. 主管打开自己的链接，看到「待我评价」名单，点一个人接着填他评
  4. HR 一键导出：每人一页同版式 Excel + 一张汇总表

启动：python app.py   （默认 http://127.0.0.1:8780/admin）
"""
import csv
import hashlib
import hmac
import io
import json
import os
import re
import secrets
import sqlite3
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any

ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

import uvicorn
from fastapi import Depends, FastAPI, HTTPException, Request, Response
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from content import VALUES
from export_xlsx import build_workbook

BASE = ROOT_DIR
DB_PATH = Path(os.environ.get("REVIEW_DB") or (BASE / "data" / "review.db"))
SECRET = os.environ.get("REVIEW_SECRET", "dasha-yu-review-2026")
ADMIN_PASSWORD = os.environ.get("REVIEW_ADMIN_PASSWORD", "admin123")
PORT = int(os.environ.get("REVIEW_PORT", "8780"))

DB_PATH.parent.mkdir(parents=True, exist_ok=True)
app = FastAPI(title="价值观考核 · 在线填写", docs_url=None, redoc_url=None)


# ---------------- 数据库 ----------------
def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


SCHEMA = """
CREATE TABLE IF NOT EXISTS settings (k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    parent_id INTEGER,
    manager_id INTEGER,
    sort INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS managers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    phone TEXT DEFAULT '',
    token TEXT NOT NULL UNIQUE
);
CREATE TABLE IF NOT EXISTS reviews (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    period TEXT NOT NULL,
    emp TEXT NOT NULL,
    phone TEXT DEFAULT '',
    dept_id INTEGER,
    mgr_id INTEGER,
    emp_token TEXT NOT NULL UNIQUE,
    mgr_token TEXT NOT NULL UNIQUE,
    status TEXT NOT NULL DEFAULT 'draft',
    answers TEXT NOT NULL DEFAULT '[]',
    created_at TEXT NOT NULL,
    emp_at TEXT,
    mgr_at TEXT
);
"""


def init_db() -> None:
    conn = db()
    cols = [r[1] for r in conn.execute("PRAGMA table_info(reviews)").fetchall()]
    if cols and "dept_id" not in cols:
        # 旧版（按人建单）数据库，备份后重建
        conn.close()
        bak = DB_PATH.with_name("review.db.bak")
        DB_PATH.replace(bak)
        print(f"[迁移] 旧数据库已备份到 {bak.name}，按新结构重建")
        conn = db()
    with conn:
        conn.executescript(SCHEMA)
        conn.execute("INSERT OR IGNORE INTO settings(k, v) VALUES('period', ?)",
                     (f"{date.today().year}年{date.today().month}月",))
        conn.execute("INSERT OR IGNORE INTO settings(k, v) VALUES('entry_open', '1')")
        n = conn.execute("SELECT COUNT(*) c FROM departments").fetchone()["c"]
        if n == 0:
            seed_demo(conn)
    conn.close()


def seed_demo(conn: sqlite3.Connection) -> None:
    def mgr(name, phone):
        cur = conn.execute("INSERT INTO managers(name, phone, token) VALUES(?,?,?)",
                           (name, phone, secrets.token_urlsafe(12)))
        return cur.lastrowid

    m_prod = mgr("周厂长", "13900000001")
    m_asm = mgr("陈主管", "13900000002")
    m_qc = mgr("赵主管", "13900000003")

    def dept(name, parent, manager=None):
        cur = conn.execute("INSERT INTO departments(name, parent_id, manager_id) VALUES(?,?,?)",
                           (name, parent, manager))
        return cur.lastrowid

    prod = dept("生产中心", None, m_prod)
    dept("装配车间", prod, m_asm)
    dept("机加工车间", prod, None)      # 未指定，继承「生产中心」的周厂长
    qc = dept("质检部", None, m_qc)
    dept("来料检验组", qc, None)
    dept("行政部", None, None)


init_db()


# ---------------- 通用 ----------------
def sign(value: str) -> str:
    return hmac.new(SECRET.encode(), value.encode(), hashlib.sha256).hexdigest()


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def get_setting(k: str, default: str = "") -> str:
    with db() as conn:
        r = conn.execute("SELECT v FROM settings WHERE k=?", (k,)).fetchone()
    return r["v"] if r else default


def set_setting(k: str, v: str) -> None:
    with db() as conn:
        conn.execute("INSERT INTO settings(k, v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v", (k, v))


def require_admin(request: Request) -> None:
    cookie = request.cookies.get("rf_auth", "")
    if not cookie or not hmac.compare_digest(cookie, sign("admin")):
        raise HTTPException(401, "未登录")


def dept_paths(conn: sqlite3.Connection) -> list[dict]:
    rows = [dict(r) for r in conn.execute("SELECT * FROM departments ORDER BY sort, id").fetchall()]
    by_id = {r["id"]: r for r in rows}
    for r in rows:
        names, cur, guard = [], r, 0
        while cur and guard < 10:
            names.append(cur["name"])
            cur = by_id.get(cur["parent_id"])
            guard += 1
        r["path"] = " / ".join(reversed(names))
        r["level"] = len(names)
    return rows


def resolve_manager(conn: sqlite3.Connection, dept_id: int | None) -> int | None:
    """部门自己没主管就往上找，找到最近的一级。"""
    cur, guard = dept_id, 0
    while cur and guard < 10:
        row = conn.execute("SELECT manager_id, parent_id FROM departments WHERE id=?", (cur,)).fetchone()
        if not row:
            return None
        if row["manager_id"]:
            return row["manager_id"]
        cur = row["parent_id"]
        guard += 1
    return None


def fetch_review(rid: int) -> sqlite3.Row:
    with db() as conn:
        r = conn.execute("SELECT * FROM reviews WHERE id=?", (rid,)).fetchone()
    if not r:
        raise HTTPException(404, "记录不存在")
    return r


def review_detail(r: sqlite3.Row) -> dict:
    with db() as conn:
        dept = conn.execute("SELECT name FROM departments WHERE id=?", (r["dept_id"],)).fetchone()
        mgr = conn.execute("SELECT name FROM managers WHERE id=?", (r["mgr_id"],)).fetchone()
    try:
        answers = json.loads(r["answers"] or "[]")
    except json.JSONDecodeError:
        answers = []
    return {
        "id": r["id"], "period": r["period"], "emp": r["emp"], "phone": r["phone"],
        "dept": dept["name"] if dept else "", "mgr": mgr["name"] if mgr else "",
        "status": r["status"],
        "emp_token": r["emp_token"], "mgr_token": r["mgr_token"],
        "created_at": r["created_at"], "emp_at": r["emp_at"], "mgr_at": r["mgr_at"],
        "answers": answers,
    }


def validate_answers(answers: Any) -> list[dict]:
    if not isinstance(answers, list) or len(answers) != len(VALUES):
        raise HTTPException(400, "答案数量不对")
    out = []
    for i, a in enumerate(answers):
        a = a if isinstance(a, dict) else {}
        item = {
            "self": a.get("self"),
            "self_note": (a.get("self_note") or "").strip()[:300],
            "mgr": a.get("mgr"),
            "mgr_note": (a.get("mgr_note") or "").strip()[:300],
            "final": a.get("final"),
        }
        for key in ("self", "mgr", "final"):
            v = item[key]
            if v is not None and (not isinstance(v, int) or not 1 <= v <= 5):
                raise HTTPException(400, f"第 {i + 1} 项分数不合法")
        out.append(item)
    return out


def person_for_export(r: sqlite3.Row) -> dict:
    d = review_detail(r)
    return {
        "emp": d["emp"], "dept": d["dept"], "mgr": d["mgr"] or "（未指定）", "period": d["period"],
        "status": d["status"], "date": (d["mgr_at"] or d["emp_at"] or d["created_at"] or "")[:10],
        "values": d["answers"], "emp_at": d["emp_at"] or "", "mgr_at": d["mgr_at"] or "",
    }


# ---------------- 页面 ----------------
app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")


@app.get("/")
def join_page():
    return FileResponse(BASE / "static" / "join.html")


@app.get("/admin")
def admin_page():
    return FileResponse(BASE / "static" / "admin.html")


@app.get("/f/{token}")
def form_page(token: str):
    return FileResponse(BASE / "static" / "form.html")


@app.get("/m/{token}")
def manager_page(token: str):
    return FileResponse(BASE / "static" / "manager.html")


# ---------------- 登录 ----------------
@app.post("/api/login")
def login(payload: dict, response: Response):
    if not hmac.compare_digest(str(payload.get("password", "")), ADMIN_PASSWORD):
        raise HTTPException(401, "密码不对")
    response.set_cookie("rf_auth", sign("admin"), httponly=True, samesite="lax", max_age=7 * 86400)
    return {"ok": True}


@app.post("/api/logout")
def logout(response: Response):
    response.delete_cookie("rf_auth")
    return {"ok": True}


@app.get("/api/me")
def me(request: Request):
    cookie = request.cookies.get("rf_auth", "")
    return {"authed": bool(cookie) and hmac.compare_digest(cookie, sign("admin"))}


# ---------------- 员工入口 ----------------
@app.get("/api/join/options")
def join_options():
    with db() as conn:
        depts = dept_paths(conn)
    return {
        "period": get_setting("period"),
        "open": get_setting("entry_open", "1") == "1",
        "departments": [{"id": d["id"], "name": d["name"], "path": d["path"], "level": d["level"],
                         "parent_id": d["parent_id"]} for d in depts],
    }


@app.post("/api/join")
def join(payload: dict):
    if get_setting("entry_open", "1") != "1":
        raise HTTPException(403, "本期填写已关闭，请联系 HR")
    phone = str(payload.get("phone", "")).strip()
    name = str(payload.get("name", "")).strip()
    dept_id = payload.get("dept_id")
    if not re.fullmatch(r"1\d{10}", phone):
        raise HTTPException(400, "手机号要填 11 位")
    if not (2 <= len(name) <= 20):
        raise HTTPException(400, "姓名长度不对")
    if not dept_id:
        raise HTTPException(400, "请选择部门")
    period = get_setting("period")
    with db() as conn:
        dept = conn.execute("SELECT * FROM departments WHERE id=?", (dept_id,)).fetchone()
        if not dept:
            raise HTTPException(400, "部门不存在")
        old = conn.execute("SELECT * FROM reviews WHERE phone=? AND period=?", (phone, period)).fetchone()
        if old:
            if old["emp"] != name:
                raise HTTPException(409, "这个手机号本期已用其他姓名登记过，请联系 HR")
            # 部门可能变了，按最新选择更新（未提交时）
            if old["status"] == "draft":
                conn.execute("UPDATE reviews SET dept_id=?, mgr_id=? WHERE id=?",
                             (dept_id, resolve_manager(conn, dept_id), old["id"]))
            return {"token": old["emp_token"], "resume": True, "status": old["status"]}
        mgr_id = resolve_manager(conn, dept_id)
        emp_token = secrets.token_urlsafe(12)
        mgr_token = secrets.token_urlsafe(12)
        conn.execute(
            "INSERT INTO reviews(period,emp,phone,dept_id,mgr_id,emp_token,mgr_token,status,answers,created_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?)",
            (period, name, phone, dept_id, mgr_id, emp_token, mgr_token, "draft",
             json.dumps([{} for _ in VALUES], ensure_ascii=False), now()),
        )
        return {"token": emp_token, "resume": False}


# ---------------- 填表（员工 / 主管共用） ----------------
def _review_by_token(token: str) -> tuple[sqlite3.Row, str]:
    with db() as conn:
        r = conn.execute("SELECT * FROM reviews WHERE emp_token=? OR mgr_token=?", (token, token)).fetchone()
    if not r:
        raise HTTPException(404, "链接无效或已失效")
    return r, ("emp" if r["emp_token"] == token else "mgr")


@app.get("/api/form/{token}")
def get_form(token: str):
    r, role = _review_by_token(token)
    d = review_detail(r)
    if role == "mgr" and d["status"] == "draft":
        raise HTTPException(403, "员工还没提交自评，暂时不能填写")
    answers = d["answers"]
    if len(answers) != len(VALUES):
        answers = [{} for _ in VALUES]
    return {
        "role": role, "period": d["period"], "emp": d["emp"], "dept": d["dept"], "mgr": d["mgr"],
        "status": d["status"], "answers": answers, "emp_at": d["emp_at"], "mgr_at": d["mgr_at"],
        "values": VALUES,
    }


@app.post("/api/form/{token}")
def post_form(token: str, payload: dict):
    r, role = _review_by_token(token)
    d = review_detail(r)
    answers = validate_answers(payload.get("answers"))
    submit = bool(payload.get("submit"))

    if role == "emp":
        if d["status"] != "draft":
            raise HTTPException(403, "自评已提交，如需修改请联系主管或 HR")
        if submit:
            for i, a in enumerate(answers):
                if not a["self"]:
                    raise HTTPException(400, f"第 {i + 1} 项「{VALUES[i]['name'].split('：')[0]}」还没选级别")
                if a["self"] >= 4 and len(a["self_note"]) < 8:
                    raise HTTPException(400, f"第 {i + 1} 项选了 {a['self']} 分，例证至少写 8 个字")
        with db() as conn:
            mgr_id = resolve_manager(conn, r["dept_id"])
            conn.execute(
                "UPDATE reviews SET answers=?, status=?, emp_at=?, mgr_id=? WHERE id=?",
                (json.dumps(answers, ensure_ascii=False), "submitted" if submit else "draft",
                 now() if submit else r["emp_at"], mgr_id, r["id"]),
            )
        return {"ok": True, "status": "submitted" if submit else "draft"}

    if d["status"] == "draft":
        raise HTTPException(403, "员工还没提交自评")
    if d["status"] == "done":
        raise HTTPException(403, "主管评已提交，如需修改请联系 HR")
    if submit:
        for i, a in enumerate(answers):
            if not a["mgr"]:
                raise HTTPException(400, f"第 {i + 1} 项「{VALUES[i]['name'].split('：')[0]}」还没选级别")
            if a["mgr"] >= 4 and len(a["mgr_note"]) < 8:
                raise HTTPException(400, f"第 {i + 1} 项选了 {a['mgr']} 分，例证至少写 8 个字")
            a["final"] = a.get("final") or a["mgr"]
    with db() as conn:
        conn.execute(
            "UPDATE reviews SET answers=?, status=?, mgr_at=? WHERE id=?",
            (json.dumps(answers, ensure_ascii=False), "done" if submit else d["status"],
             now() if submit else d["mgr_at"], r["id"]),
        )
    return {"ok": True, "status": "done" if submit else d["status"]}


# ---------------- 主管端 ----------------
@app.get("/api/manager/{token}")
def manager_home(token: str):
    with db() as conn:
        m = conn.execute("SELECT * FROM managers WHERE token=?", (token,)).fetchone()
        if not m:
            raise HTTPException(404, "链接无效或主管已调整，请联系 HR")
        all_rows = conn.execute("SELECT * FROM reviews ORDER BY id DESC").fetchall()
        rows = [r for r in all_rows if r["dept_id"] and _dept_under(conn, r["dept_id"], m["id"])]
        # 顺手把归属补正（部门主管后改过的情况）
        for r in rows:
            if r["mgr_id"] != m["id"]:
                conn.execute("UPDATE reviews SET mgr_id=? WHERE id=?", (m["id"], r["id"]))
        rows = [dict(r) for r in rows]
    order = {"submitted": 0, "done": 1, "draft": 2}
    rows.sort(key=lambda r: (order.get(r["status"], 3), r["id"]))
    items = []
    for r in rows:
        d = review_detail(r)
        items.append({
            "id": d["id"], "emp": d["emp"], "dept": d["dept"], "period": d["period"],
            "status": d["status"], "emp_at": d["emp_at"] or "", "mgr_at": d["mgr_at"] or "",
            "link": f"/f/{d['mgr_token']}",
        })
    return {
        "manager": {"name": m["name"], "phone": m["phone"]},
        "todo": [i for i in items if i["status"] == "submitted"],
        "done": [i for i in items if i["status"] == "done"],
        "pending_self": sum(1 for i in items if i["status"] == "draft"),
    }


def _dept_under(conn: sqlite3.Connection, dept_id: int, manager_id: int) -> bool:
    """该部门（或它的某个上级）是否由这个主管负责。"""
    cur, guard = dept_id, 0
    while cur and guard < 10:
        row = conn.execute("SELECT manager_id, parent_id FROM departments WHERE id=?", (cur,)).fetchone()
        if not row:
            return False
        if row["manager_id"] == manager_id:
            return True
        if row["manager_id"]:
            return False          # 上一级已经被别人负责，不再往上继承
        cur = row["parent_id"]
        guard += 1
    return False


# ---------------- 后台：组织 ----------------
@app.get("/api/org")
def get_org(request: Request, _: None = Depends(require_admin)):
    base = str(request.base_url).rstrip("/")
    with db() as conn:
        depts = dept_paths(conn)
        managers = [dict(m) for m in conn.execute("SELECT * FROM managers ORDER BY id").fetchall()]
        counts = {r["dept_id"]: r["c"] for r in conn.execute(
            "SELECT dept_id, COUNT(*) c FROM reviews GROUP BY dept_id").fetchall()}
    for d in depts:
        d["manager_name"] = next((m["name"] for m in managers if m["id"] == d["manager_id"]), "")
        d["staff"] = counts.get(d["id"], 0)
    return {
        "period": get_setting("period"),
        "entry_open": get_setting("entry_open", "1") == "1",
        "departments": depts,
        "managers": [{"id": m["id"], "name": m["name"], "phone": m["phone"],
                      "link": f"{base}/m/{m['token']}"} for m in managers],
    }


@app.post("/api/settings")
def save_settings(payload: dict, _: None = Depends(require_admin)):
    if "period" in payload:
        p = str(payload["period"]).strip()
        if not p:
            raise HTTPException(400, "周期不能为空")
        set_setting("period", p)
    if "entry_open" in payload:
        set_setting("entry_open", "1" if payload["entry_open"] else "0")
    return {"ok": True}


@app.post("/api/departments")
def add_department(payload: dict, _: None = Depends(require_admin)):
    name = str(payload.get("name", "")).strip()
    if not name:
        raise HTTPException(400, "部门名不能为空")
    parent = payload.get("parent_id") or None
    with db() as conn:
        if parent and not conn.execute("SELECT 1 FROM departments WHERE id=?", (parent,)).fetchone():
            raise HTTPException(400, "上级部门不存在")
        same = conn.execute("SELECT 1 FROM departments WHERE name=? AND IFNULL(parent_id,0)=IFNULL(?,0)",
                            (name, parent)).fetchone()
        if same:
            raise HTTPException(400, "同级已经有这个部门了")
        cur = conn.execute("INSERT INTO departments(name, parent_id, manager_id) VALUES(?,?,NULL)", (name, parent))
        return {"id": cur.lastrowid}


@app.post("/api/departments/{did}")
def update_department(did: int, payload: dict, _: None = Depends(require_admin)):
    with db() as conn:
        if not conn.execute("SELECT 1 FROM departments WHERE id=?", (did,)).fetchone():
            raise HTTPException(404, "部门不存在")
        if "name" in payload:
            name = str(payload["name"]).strip()
            if not name:
                raise HTTPException(400, "部门名不能为空")
            conn.execute("UPDATE departments SET name=? WHERE id=?", (name, did))
        if "manager_id" in payload:
            mid = payload["manager_id"] or None
            if mid and not conn.execute("SELECT 1 FROM managers WHERE id=?", (mid,)).fetchone():
                raise HTTPException(400, "主管不存在")
            conn.execute("UPDATE departments SET manager_id=? WHERE id=?", (mid, did))
    return {"ok": True}


@app.post("/api/departments/{did}/delete")
def delete_department(did: int, _: None = Depends(require_admin)):
    with db() as conn:
        if conn.execute("SELECT 1 FROM departments WHERE parent_id=?", (did,)).fetchone():
            raise HTTPException(409, "还有子部门，先删子部门")
        if conn.execute("SELECT 1 FROM reviews WHERE dept_id=?", (did,)).fetchone():
            raise HTTPException(409, "这个部门下还有考核记录，不能删")
        conn.execute("DELETE FROM departments WHERE id=?", (did,))
    return {"ok": True}


@app.post("/api/departments/bulk")
def bulk_departments(payload: dict, _: None = Depends(require_admin)):
    text = str(payload.get("text", ""))
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        raise HTTPException(400, "没有内容")
    if len(lines) > 300:
        raise HTTPException(400, "一次最多 300 行")
    created = 0
    with db() as conn:
        cache: dict[tuple, int] = {}
        for ln in lines:
            parts = [p.strip() for p in re.split(r"[/／>、,，\t]+", ln) if p.strip()]
            parent = None
            key = ()
            for name in parts:
                key = key + (name,)
                if key in cache:
                    parent = cache[key]
                    continue
                row = conn.execute(
                    "SELECT id FROM departments WHERE name=? AND IFNULL(parent_id,0)=IFNULL(?,0)",
                    (name, parent)).fetchone()
                if row:
                    parent = row["id"]
                else:
                    parent = conn.execute(
                        "INSERT INTO departments(name, parent_id, manager_id) VALUES(?,?,NULL)",
                        (name, parent)).lastrowid
                    created += 1
                cache[key] = parent
    return {"created": created}


@app.post("/api/managers")
def add_manager(payload: dict, request: Request, _: None = Depends(require_admin)):
    name = str(payload.get("name", "")).strip()
    phone = str(payload.get("phone", "")).strip()
    if not name:
        raise HTTPException(400, "主管姓名不能为空")
    with db() as conn:
        cur = conn.execute("INSERT INTO managers(name, phone, token) VALUES(?,?,?)",
                           (name, phone, secrets.token_urlsafe(12)))
        mid = cur.lastrowid
        token = conn.execute("SELECT token FROM managers WHERE id=?", (mid,)).fetchone()["token"]
    base = str(request.base_url).rstrip("/")
    return {"id": mid, "link": f"{base}/m/{token}"}


@app.post("/api/managers/{mid}/delete")
def delete_manager(mid: int, _: None = Depends(require_admin)):
    with db() as conn:
        used = conn.execute("SELECT COUNT(*) c FROM reviews WHERE mgr_id=?", (mid,)).fetchone()["c"]
        if used:
            raise HTTPException(409, "还有考核记录挂在这个主管名下，先改部门主管")
        conn.execute("DELETE FROM managers WHERE id=?", (mid,))
    return {"ok": True}


# ---------------- 后台：记录 ----------------
@app.get("/api/reviews")
def list_reviews(request: Request, _: None = Depends(require_admin)):
    with db() as conn:
        rows = conn.execute("SELECT * FROM reviews ORDER BY id DESC").fetchall()
    base = str(request.base_url).rstrip("/")
    items = []
    for r in rows:
        d = review_detail(r)
        items.append({
            "id": d["id"], "emp": d["emp"], "phone": d["phone"], "dept": d["dept"],
            "dept_id": r["dept_id"], "mgr": d["mgr"],
            "period": d["period"], "status": d["status"], "created_at": d["created_at"],
            "emp_at": d["emp_at"], "mgr_at": d["mgr_at"],
            "emp_link": f"{base}/f/{d['emp_token']}", "mgr_link": f"{base}/f/{d['mgr_token']}",
        })
    return {"items": items}


@app.post("/api/reviews/{rid}/dept")
def change_dept(rid: int, payload: dict, _: None = Depends(require_admin)):
    """员工选错部门时，HR 改一下，主管归属跟着重算。"""
    dept_id = payload.get("dept_id")
    with db() as conn:
        if not dept_id or not conn.execute("SELECT 1 FROM departments WHERE id=?", (dept_id,)).fetchone():
            raise HTTPException(400, "部门不存在")
        conn.execute("UPDATE reviews SET dept_id=?, mgr_id=? WHERE id=?",
                     (dept_id, resolve_manager(conn, dept_id), rid))
    return {"ok": True}


@app.post("/api/reviews/{rid}/reset")
def reset_review(rid: int, _: None = Depends(require_admin)):
    fetch_review(rid)
    with db() as conn:
        conn.execute("UPDATE reviews SET status='draft', emp_at=NULL, mgr_at=NULL WHERE id=?", (rid,))
    return {"ok": True}


@app.post("/api/reviews/{rid}/delete")
def delete_review(rid: int, _: None = Depends(require_admin)):
    with db() as conn:
        conn.execute("DELETE FROM reviews WHERE id=?", (rid,))
    return {"ok": True}


@app.get("/api/reviews/{rid}/export")
def export_one(rid: int, _: None = Depends(require_admin)):
    p = person_for_export(fetch_review(rid))
    buf = build_workbook([p])
    name = f"{p['emp']}-价值观考核-{p['period']}.xlsx"
    return StreamingResponse(
        buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{_quote(name)}"},
    )


@app.get("/api/export")
def export_all(_: None = Depends(require_admin)):
    with db() as conn:
        rows = conn.execute("SELECT * FROM reviews ORDER BY id").fetchall()
    persons = [person_for_export(r) for r in rows]
    if not persons:
        raise HTTPException(400, "还没有考核记录")
    buf = build_workbook(persons)
    name = f"价值观考核汇总-{date.today().isoformat()}.xlsx"
    return StreamingResponse(
        buf, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{_quote(name)}"},
    )


def _quote(s: str) -> str:
    from urllib.parse import quote
    return quote(s)


@app.get("/api/links.csv")
def export_links(request: Request, _: None = Depends(require_admin)):
    base = str(request.base_url).rstrip("/")
    with db() as conn:
        rows = conn.execute("SELECT * FROM reviews ORDER BY id").fetchall()
        managers = conn.execute("SELECT * FROM managers ORDER BY id").fetchall()
    st = {"draft": "待自评", "submitted": "待主管评", "done": "已完成"}
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["类型", "姓名", "部门", "主管", "周期", "状态", "链接"])
    for r in rows:
        d = review_detail(r)
        w.writerow(["员工", d["emp"], d["dept"], d["mgr"], d["period"], st.get(d["status"], ""),
                    f"{base}/f/{d['emp_token']}"])
        w.writerow(["主管", d["mgr"], d["dept"], "", d["period"], "",
                    f"{base}/f/{d['mgr_token']}"])
    for m in managers:
        w.writerow(["主管", m["name"], "", "", "", "", f"{base}/m/{m['token']}"])
    data = "\ufeff" + buf.getvalue()
    name = f"考核链接-{date.today().isoformat()}.csv"
    return Response(
        content=data.encode("utf-8"), media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{_quote(name)}"},
    )


@app.exception_handler(HTTPException)
def http_err(request: Request, exc: HTTPException):
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)


if __name__ == "__main__":
    print(f"价值观考核 · 在线填写  →  入口 http://127.0.0.1:{PORT}/   后台 http://127.0.0.1:{PORT}/admin")
    uvicorn.run(app, host="0.0.0.0", port=PORT, log_level="info")
