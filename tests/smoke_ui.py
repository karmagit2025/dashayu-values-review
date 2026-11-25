# -*- coding: utf-8 -*-
"""浏览器走查（可选）：起临时服务 → 员工登记 → 自评 → 主管评价 → 截图。

依赖：pip install playwright && playwright install chromium
用法：python tests/smoke_ui.py
"""
import os
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SHOTS = ROOT / "shots"
PORT = int(os.environ.get("SMOKE_UI_PORT", "8798"))
BASE = f"http://127.0.0.1:{PORT}"
LEVELS = [4, 3, 5, 2, 4, 3]
NOTES = ["11月客户设备故障，2小时内到场处理，客户书面表扬。", "",
         "主动推动夹具改造，换型从40分钟降到18分钟。", "",
         "连续3个月提前完成排产，主动帮同事顶班。", ""]


def wait_ready(timeout=25.0):
    end = time.time() + timeout
    while time.time() < end:
        try:
            urllib.request.urlopen(BASE + "/api/join/options", timeout=3)
            return True
        except Exception:
            time.sleep(0.4)
    return False


def main() -> int:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("没装 playwright：pip install playwright && playwright install chromium")
        return 2

    tmpdir = tempfile.mkdtemp(prefix="values-review-ui-")
    env = dict(os.environ, REVIEW_DB=str(Path(tmpdir) / "ui.db"), REVIEW_PORT=str(PORT),
               PYTHONIOENCODING="utf-8")
    proc = subprocess.Popen([sys.executable, "app.py"], cwd=str(ROOT), env=env,
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    SHOTS.mkdir(exist_ok=True)
    try:
        if not wait_ready():
            print("服务没起来")
            return 1
        with sync_playwright() as p:
            try:
                b = p.chromium.launch()
            except Exception:
                # 没下载 playwright 自带浏览器时，退回本机 Edge
                b = p.chromium.launch(channel="msedge")
            m = b.new_context(viewport={"width": 390, "height": 844}, device_scale_factor=2,
                              is_mobile=True)
            page = m.new_page()
            page.goto(BASE + "/")
            page.wait_for_selector("#start", timeout=8000)
            page.screenshot(path=str(SHOTS / "ui-1-entry.png"))

            opt = page.evaluate("async () => await (await fetch('/api/join/options')).json()")
            dept_id = [d for d in opt["departments"] if d["name"] == "装配车间"][0]["id"]
            page.fill("#phone", "13800000099")
            page.fill("#name", "走查员工")
            page.select_option("#dept", str(dept_id))
            page.click("#start")
            page.wait_for_selector(".row", timeout=8000)
            page.screenshot(path=str(SHOTS / "ui-2-self.png"))

            for i, lv in enumerate(LEVELS):
                card = page.locator(f'[data-card="{i}"]')
                card.locator(f'.row[data-lv="{lv}"]').click()
                page.wait_for_timeout(80)
                if NOTES[i]:
                    card.locator("textarea").fill(NOTES[i])
            page.click("#submitBtn")
            page.wait_for_selector("#mask.show", timeout=5000)
            page.click("#dlgOk")
            page.wait_for_selector(".done", timeout=8000)
            print("✔ 员工自评提交")

            # 登录后台拿主管链接
            page.goto(BASE + "/admin")
            page.fill("#pwd", "admin123")
            page.click("#loginBtn")
            page.wait_for_selector(".tabs", timeout=8000)
            org = page.evaluate("async () => await (await fetch('/api/org')).json()")
            chen = [x for x in org["managers"] if x["name"] == "陈主管"][0]

            q = m.new_page()
            q.goto(chen["link"])
            q.wait_for_selector(".person", timeout=8000)
            q.screenshot(path=str(SHOTS / "ui-3-queue.png"))
            print("✔ 主管队列出现待评价")

            q.locator(".person a.btn").first.click()
            q.wait_for_selector(".selfbox", timeout=8000)
            for i, lv in enumerate(LEVELS):
                card = q.locator(f'[data-card="{i}"]')
                card.locator(f'.row[data-lv="{lv}"]').click()
                q.wait_for_timeout(70)
                if NOTES[i]:
                    card.locator("textarea").fill("已核实：" + NOTES[i])
            q.click("#submitBtn")
            q.wait_for_selector("#mask.show", timeout=5000)
            q.click("#dlgOk")
            q.wait_for_selector(".done", timeout=8000)
            q.screenshot(path=str(SHOTS / "ui-4-done.png"))
            print("✔ 主管评价提交")

            q.goto(chen["link"])
            q.wait_for_selector(".stats", timeout=8000)
            page.wait_for_timeout(300)
            print("✔ 队列已清空，流程通过")
            b.close()
        return 0
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


if __name__ == "__main__":
    sys.exit(main())
