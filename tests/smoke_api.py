# -*- coding: utf-8 -*-
"""接口冒烟测试：起一个临时数据库跑完整流程，跑完自动清理。

用法：python tests/smoke_api.py
不需要额外依赖（标准库 + 项目依赖 openpyxl 可选）。
"""
import http.cookiejar
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = int(os.environ.get("SMOKE_PORT", "8799"))
BASE = f"http://127.0.0.1:{PORT}"

results: list[tuple[bool, str, str]] = []


def check(name: str, ok: bool, extra: str = "") -> None:
    results.append((bool(ok), name, extra))
    print(("  ✔ " if ok else "  ✘ ") + name + (f"  ({extra})" if extra and not ok else ""))


class Client:
    def __init__(self) -> None:
        self.jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar))

    def call(self, path: str, data=None, raw: bool = False):
        req = urllib.request.Request(
            BASE + path,
            data=json.dumps(data, ensure_ascii=False).encode() if data is not None else None,
            headers={"Content-Type": "application/json"},
        )
        try:
            with self.opener.open(req, timeout=10) as r:
                body = r.read()
                return r.status, (body if raw else json.loads(body))
        except urllib.error.HTTPError as e:
            try:
                return e.code, json.loads(e.read() or b"{}")
            except json.JSONDecodeError:
                return e.code, {}


def wait_ready(timeout: float = 25.0) -> bool:
    client = Client()
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            st, _ = client.call("/api/join/options")
            if st == 200:
                return True
        except Exception:
            pass
        time.sleep(0.4)
    return False


def main() -> int:
    tmpdir = tempfile.mkdtemp(prefix="values-review-smoke-")
    db_path = Path(tmpdir) / "smoke.db"
    env = dict(os.environ)
    env.update({
        "REVIEW_DB": str(db_path),
        "REVIEW_PORT": str(PORT),
        "REVIEW_ADMIN_PASSWORD": "smoke-pass",
        "PYTHONIOENCODING": "utf-8",
    })

    print(f"启动临时服务（数据库 {db_path}）…")
    proc = subprocess.Popen([sys.executable, "app.py"], cwd=str(ROOT), env=env,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        if not wait_ready():
            print("服务没起来，检查端口占用或依赖是否安装")
            return 1

        c = Client()
        ok = c.call("/api/join/options")
        depts = {d["name"]: d["id"] for d in ok[1]["departments"]}
        check("读入口配置（含示例部门）", ok[0] == 200 and len(depts) >= 5, str(ok[1])[:80])

        # 登录
        check("错误密码被拒", c.call("/api/login", {"password": "x"})[0] == 401)
        check("正确密码登录", c.call("/api/login", {"password": "smoke-pass"})[0] == 200)
        check("未登录访问后台接口被拒", Client().call("/api/reviews")[0] == 401)

        # 员工登记
        st, j = c.call("/api/join", {"phone": "13800138000", "name": "测试员工", "dept_id": depts["装配车间"]})
        check("员工登记", st == 200 and j.get("token"), str(j))
        emp_tok = j["token"]
        st, j2 = c.call("/api/join", {"phone": "13800138000", "name": "测试员工", "dept_id": depts["装配车间"]})
        check("同手机号重复进入 → 续填", st == 200 and j2.get("resume") is True)
        st, j3 = c.call("/api/join", {"phone": "13800138000", "name": "别人", "dept_id": depts["装配车间"]})
        check("同手机号换姓名 → 拒绝", st == 409)
        st, _ = c.call("/api/join", {"phone": "123", "name": "测试", "dept_id": depts["装配车间"]})
        check("手机号格式校验", st == 400)

        # 自评
        st, form = c.call(f"/api/form/{emp_tok}")
        check("员工取表单（部门/主管已解析）", st == 200 and form["dept"] == "装配车间" and form["mgr"] == "陈主管",
              f"{form.get('dept')}/{form.get('mgr')}")
        answers = [{"self": n, "self_note": (f"11月案例{i}：现场处理客户问题，客户确认满意。" if n >= 4 else "")}
                   for i, n in enumerate([4, 3, 5, 2, 4, 3])]
        bad = [dict(a) for a in answers]
        bad[3]["self"] = None
        check("缺项提交 → 400", c.call(f"/api/form/{emp_tok}", {"answers": bad, "submit": True})[0] == 400)
        bad2 = [dict(a) for a in answers]
        bad2[0]["self_note"] = "太短"
        check("4 分例证过短 → 400", c.call(f"/api/form/{emp_tok}", {"answers": bad2, "submit": True})[0] == 400)
        st, _ = c.call(f"/api/form/{emp_tok}", {"answers": answers, "submit": True})
        check("自评提交成功", st == 200)
        check("重复提交自评 → 403", c.call(f"/api/form/{emp_tok}", {"answers": answers, "submit": True})[0] == 403)

        # 主管
        st, org = c.call("/api/org")
        mgr = {m["name"]: m for m in org["managers"]}
        chen_tok = mgr["陈主管"]["link"].rsplit("/", 1)[-1]
        zhou_tok = mgr["周厂长"]["link"].rsplit("/", 1)[-1]
        st, q = c.call(f"/api/manager/{chen_tok}")
        check("主管队列出现待评价", st == 200 and any(x["emp"] == "测试员工" for x in q["todo"]),
              json.dumps([x["emp"] for x in q["todo"]], ensure_ascii=False))
        st, q2 = c.call(f"/api/manager/{zhou_tok}")
        check("其他主管看不到这个员工", all(x["emp"] != "测试员工" for x in q2["todo"]))

        rec = [x for x in c.call("/api/reviews")[1]["items"] if x["emp"] == "测试员工"][0]
        mgr_tok = rec["mgr_link"].rsplit("/", 1)[-1]
        st, mf = c.call(f"/api/form/{mgr_tok}")
        check("主管能看到员工自评", st == 200 and mf["answers"][0]["self"] == 4)
        mgr_answers = [{"self": a["self"], "self_note": a["self_note"], "mgr": a["self"],
                        "mgr_note": ("主管核实：现场跟踪到位。" if a["self"] >= 4 else ""),
                        "final": a["self"]} for a in mf["answers"]]
        st, _ = c.call(f"/api/form/{mgr_tok}", {"answers": mgr_answers, "submit": True})
        check("主管评价提交", st == 200)
        check("重复提交主管评 → 403", c.call(f"/api/form/{mgr_tok}", {"answers": mgr_answers, "submit": True})[0] == 403)
        st, q3 = c.call(f"/api/manager/{chen_tok}")
        check("队列变为已完成", not any(x["emp"] == "测试员工" for x in q3["todo"]))

        # 继承
        st, j4 = c.call("/api/join", {"phone": "13800138001", "name": "继承测试", "dept_id": depts["机加工车间"]})
        rec2 = [x for x in c.call("/api/reviews")[1]["items"] if x["emp"] == "继承测试"][0]
        check("子部门继承上级主管", rec2["mgr"] == "周厂长", rec2["mgr"])

        # 改部门
        st, _ = c.call(f"/api/reviews/{rec2['id']}/dept", {"dept_id": depts["质检部"]})
        rec3 = [x for x in c.call("/api/reviews")[1]["items"] if x["emp"] == "继承测试"][0]
        check("改部门后主管重算", rec3["mgr"] == "赵主管", rec3["mgr"])

        # 导出
        st, blob = c.call(f"/api/reviews/{rec['id']}/export", raw=True)
        check("单人导出返回 xlsx", st == 200 and blob[:2] == b"PK", f"{st}/{len(blob)}")
        import io
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            names = z.namelist()
        check("xlsx 结构完整", "xl/workbook.xml" in names)
        st, blob2 = c.call("/api/export", raw=True)
        check("批量导出返回 xlsx", st == 200 and blob2[:2] == b"PK")

        # 关入口
        c.call("/api/settings", {"entry_open": False})
        check("关闭入口后登记被拒", c.call("/api/join", {"phone": "13800138002", "name": "关门", "dept_id": depts["行政部"]})[0] == 403)
        c.call("/api/settings", {"entry_open": True})
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=8)
        except subprocess.TimeoutExpired:
            proc.kill()
        for f in Path(tmpdir).glob("*"):
            try:
                f.unlink()
            except OSError:
                pass
        try:
            Path(tmpdir).rmdir()
        except OSError:
            pass

    passed = sum(1 for ok, _, _ in results if ok)
    failed = [(n, e) for ok, n, e in results if not ok]
    print(f"\n{passed}/{len(results)} 项通过")
    if failed:
        print("失败项：")
        for n, e in failed:
            print("  -", n, e)
        return 1
    print("全部通过 ✅")
    return 0


if __name__ == "__main__":
    sys.exit(main())
