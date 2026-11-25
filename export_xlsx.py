# -*- coding: utf-8 -*-
"""把一页考核记录导出成与纸质表同版的 Excel。"""
import io
import math
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties

from content import VALUES, CHINESE_NUMS, LEVEL_MARKS

# ---------- 版式常量（与纸质排版版一致） ----------
PAGE_W = 210
MARGIN_LR = 8
CONTENT_W = PAGE_W - 2 * MARGIN_LR
COL = {"A": 8, "B": 79, "C": 79, "D": 9, "E": 10, "F": 9}

BODY_PT = 8
PAD_MM = 0.95
ROW_MIN = 4.3

NAVY = "1F3864"
thin = Side(style="thin", color="A6A6A6")
med = Side(style="medium", color=NAVY)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

F_TITLE = Font(name="黑体", size=15, bold=True, color="1F1F1F")
F_INFO = Font(name="宋体", size=9, color="333333")
F_NOTE = Font(name="宋体", size=7.5, color="333333")
F_HEAD = Font(name="黑体", size=8, bold=True, color="1F1F1F")
F_BAND = Font(name="黑体", size=9, bold=True, color="FFFFFF")
F_IDX = Font(name="宋体", size=7.5, color="595959")
F_TEXT = Font(name="宋体", size=BODY_PT, color="000000")
F_EVID = Font(name="宋体", size=7.5, color="000000")
F_TICK = Font(name="宋体", size=11, bold=True, color=NAVY)

FILL_BAND = PatternFill("solid", fgColor=NAVY)
FILL_HEAD = PatternFill("solid", fgColor="DCE3F0")
FILL_NOTE = PatternFill("solid", fgColor="F4F4F4")
FILL_EVID = PatternFill("solid", fgColor="FFFBEF")
FILL_IDX = PatternFill("solid", fgColor="F7F7F7")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")


def mm2w(mm):
    return (mm * 96 / 25.4 - 5) / 7


def mm2h(mm):
    return mm * 72 / 25.4


def _lines(text, width_mm, pt):
    usable = width_mm * 96 / 25.4 - 4
    cap = max(1, int(usable / (pt * 96 / 72)))
    n = 0
    for seg in str(text).split("\n"):
        n += max(1, math.ceil(len(seg) / cap))
    return n


def _style_range(ws, rng, font=None, fill=None, align=None, border=None):
    for rr in ws[rng]:
        for c in rr:
            if font:
                c.font = font
            if fill:
                c.fill = fill
            if align:
                c.alignment = align
            if border:
                c.border = border


def _setup_print(ws, last_row):
    ws.print_area = f"A1:F{last_row}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = "9"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = MARGIN_LR / 25.4
    ws.page_margins.right = MARGIN_LR / 25.4
    ws.page_margins.top = 9 / 25.4
    ws.page_margins.bottom = 8 / 25.4
    ws.page_margins.header = 0
    ws.page_margins.footer = 0
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False


def build_person_sheet(wb, person, first=False):
    title = f"{person.get('emp', '')}-{person.get('period', '')}"
    title = "".join(ch for ch in title if ch not in '[]:*?/\\')[:28] or "考核表"
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    for k, w in COL.items():
        ws.column_dimensions[k].width = mm2w(w)

    row = 1

    def set_h(r, mm):
        ws.row_dimensions[r].height = mm2h(mm)

    # 标题
    set_h(row, 8.0)
    ws.merge_cells(f"A{row}:F{row}")
    c = ws[f"A{row}"]
    c.value = "泉州市大鲨鱼机械科技有限公司　价值观考核表"
    c.font = F_TITLE
    c.alignment = Alignment(horizontal="center", vertical="center")
    row += 1

    # 信息行
    set_h(row, 6.5)
    ws.merge_cells(f"A{row}:F{row}")
    c = ws[f"A{row}"]
    c.value = (f"姓　名：{person.get('emp', '')}　　所在部门：{person.get('dept', '')}"
               f"　　他评人员：{person.get('mgr', '')}　　填写周期：{person.get('period', '')}")
    c.font = F_INFO
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border = Border(bottom=thin)
    row += 1

    # 填写须知
    note = ("1. 价值观是大鲨鱼未来成长重要的航标，请认真自评与他评，真实、客观、公正。\n"
            "2. 考核标准：采用通关制，勾选对应分值；前面小分值未达到，后面的考核分数不计分值。评 4-5 分须填写例证，用真实案例佐证。")
    set_h(row, 7.6)
    ws.merge_cells(f"A{row}:F{row}")
    _style_range(ws, f"A{row}:F{row}", fill=FILL_NOTE)
    c = ws[f"A{row}"]
    c.value = note
    c.font = F_NOTE
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    row += 1

    # 表头
    set_h(row, 6.5)
    _style_range(ws, f"A{row}:F{row}", font=F_HEAD, fill=FILL_HEAD, align=CENTER, border=BOX)
    ws.merge_cells(f"B{row}:C{row}")
    for col, v in {"A": "序号", "B": "评估说明（勾选所达级别）", "D": "自评", "E": "主管评", "F": "综合"}.items():
        ws[f"{col}{row}"].value = v
    row += 1

    answers = person.get("values") or [{} for _ in VALUES]
    n = 1
    for vi, val in enumerate(VALUES):
        ans = answers[vi] if vi < len(answers) else {}
        set_h(row, 5.4)
        _style_range(ws, f"A{row}:F{row}", font=F_BAND, fill=FILL_BAND,
                     align=Alignment(horizontal="left", vertical="center", indent=1))
        ws.merge_cells(f"A{row}:F{row}")
        ws[f"A{row}"].value = f"{CHINESE_NUMS[vi]}、{val['name']}"
        ws[f"A{row}"].border = Border(top=med, bottom=thin, left=med, right=med)
        row += 1

        first_item = row
        for li, text in enumerate(val["items"]):
            set_h(row, 4.48)
            _style_range(ws, f"A{row}:F{row}", border=BOX)
            ws[f"A{row}"].value = n
            ws[f"A{row}"].font = F_IDX
            ws[f"A{row}"].fill = FILL_IDX
            ws[f"A{row}"].alignment = CENTER
            ws.merge_cells(f"B{row}:C{row}")
            ws[f"B{row}"].value = LEVEL_MARKS[li] + text
            ws[f"B{row}"].font = F_TEXT
            ws[f"B{row}"].alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=True, indent=1)
            # 打勾
            lvl = li + 1
            if ans.get("self") == lvl:
                ws[f"D{row}"].value = "√"
                ws[f"D{row}"].font = F_TICK
            if ans.get("mgr") == lvl:
                ws[f"E{row}"].value = "√"
                ws[f"E{row}"].font = F_TICK
            for col in ("D", "E", "F"):
                ws[f"{col}{row}"].alignment = CENTER
            n += 1
            row += 1

        ws.merge_cells(f"F{first_item}:F{row - 1}")
        final = ans.get("final")
        if final:
            ws[f"F{first_item}"].value = final
            ws[f"F{first_item}"].font = Font(name="宋体", size=11, bold=True, color=NAVY)
        ws[f"F{first_item}"].alignment = CENTER

        # 例证
        self_note = (ans.get("self_note") or "").strip()
        mgr_note = (ans.get("mgr_note") or "").strip()
        lines = max(_lines(self_note, COL["B"] - 4, 7.5) if self_note else 1,
                    _lines(mgr_note, COL["C"] + COL["D"] + COL["E"] + COL["F"] - 4, 7.5) if mgr_note else 1)
        h = max(12.8, lines * 3.35 + 1.6)
        set_h(row, h)
        _style_range(ws, f"A{row}:F{row}", border=BOX, fill=FILL_EVID)
        ws[f"A{row}"].value = "例证"
        ws[f"A{row}"].font = Font(name="宋体", size=7.5, color="595959")
        ws[f"A{row}"].alignment = CENTER
        ws[f"A{row}"].fill = PatternFill("solid", fgColor="F0EEE6")
        ws[f"B{row}"].value = self_note or None
        ws[f"B{row}"].font = F_EVID
        ws[f"B{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        ws.merge_cells(f"C{row}:F{row}")
        ws[f"C{row}"].value = mgr_note or None
        ws[f"C{row}"].font = F_EVID
        ws[f"C{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
        row += 1

    # 签字行
    set_h(row, 6.5)
    ws.merge_cells(f"A{row}:F{row}")
    c = ws[f"A{row}"]
    c.value = (f"员工签字：＿＿＿＿＿＿＿＿　　主管签字：＿＿＿＿＿＿＿＿　　"
               f"日期：{person.get('date', '')}")
    c.font = F_INFO
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    c.border = Border(top=thin)

    _setup_print(ws, row)
    return ws


def build_workbook(persons):
    wb = Workbook()
    for i, p in enumerate(persons):
        build_person_sheet(wb, p, first=(i == 0))
    if len(persons) > 1:
        ws = wb.create_sheet("汇总", 0)
        _build_summary(ws, persons)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _build_summary(ws, persons):
    ws.sheet_view.showGridLines = False
    heads = ["姓名", "部门", "主管", "周期", "状态"] + [v["name"].split("：")[0] for v in VALUES] + ["平均", "自评提交", "主管完成"]
    widths = [10, 14, 10, 12, 10] + [9] * len(VALUES) + [8, 16, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(heads))
    c = ws.cell(1, 1, "价值观考核汇总")
    c.font = Font(name="黑体", size=13, bold=True, color="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for j, h in enumerate(heads, start=1):
        c = ws.cell(2, j, h)
        c.font = Font(name="黑体", size=9, bold=True)
        c.fill = PatternFill("solid", fgColor="DCE3F0")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[2].height = 30

    status_label = {"draft": "待自评", "submitted": "待主管评", "done": "已完成"}
    for i, p in enumerate(persons, start=3):
        vals = [v.get("final") for v in (p.get("values") or [])]
        nums = [v for v in vals if isinstance(v, (int, float)) and v > 0]
        avg = round(sum(nums) / len(nums), 2) if nums else ""
        row_data = [p.get("emp", ""), p.get("dept", ""), p.get("mgr", ""), p.get("period", ""),
                    status_label.get(p.get("status", ""), "")]
        row_data += [v if isinstance(v, int) else "" for v in vals[:len(VALUES)]]
        row_data += [avg, p.get("emp_at", ""), p.get("mgr_at", "")]
        for j, v in enumerate(row_data, start=1):
            c = ws.cell(i, j, v)
            c.font = Font(name="宋体", size=9)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BOX
    ws.freeze_panes = "A3"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = "9"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = 0.3
    ws.page_margins.right = 0.3
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
