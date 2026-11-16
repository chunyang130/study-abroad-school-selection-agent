#!/usr/bin/env python3
"""依 SKILL.md 的格式規則產生選校表（初選表／終選校表）xlsx，並在存檔前自動檢查。

用法:
  python3 build_school_list.py --input examples/sample_student.json --out output/sample_初選表.xlsx

輸入 JSON：
  student      學生背景（多行文字）
  direction    科系方向要求
  sheet        "初選表" 或 "終選校表"
  country      "美國" | "英國" | "澳洲" | "加拿大"（決定應有校數）
  rows         每所學校一筆，鍵名用欄位短名；值可以是 [文字, 連結] 代表超連結
  sources      資料來源工作表的列：[資料類別, 來源, 版本, 查證日期]
"""
import argparse, json, os, sys
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font

TIER_FILL = {"夢幻": "F4CCCC", "合適": "FFF2CC", "低風險": "D9EAD3"}
HEAD_FILL = "C9DAF8"
INIT_COLS = ["確定申請", "", "學校", "科系(&LINK)", "綜合排名\nQS RANKING", "科系排名", "學校位置", "先決背景",
             "在校成績門檻", "TOEFL", "STEM", "GRE ", "課綱 ", "成績單要求", "WES", "CV", "SOP ", "LoR", "Essay", "時長", "學分",
             "Deadline ", "學費", "其他(ex.科系特色)"]
FINAL_COLS = ["確定申請", "", "學校", "科系(&LINK)", "美國綜合排名", "科系排名", "學校位置", "先決背景", "在校成績門檻",
              "TOEFL \n總分(單項要求)", "STEM", "GRE ", "Lab / Facilities", "課綱 ", "成績單要求", "CV", "SOP ", "LoR", "Essay", "學分",
              "Deadline ", "學費\n(會漲價)", "其他", "", "備註/選擇", "", "more schools"]
WIDTH = {"A": 11, "B": 9, "C": 23, "D": 28, "E": 13, "F": 13, "G": 16, "H": 15, "I": 15, "J": 16, "M": 15, "N": 15, "O": 15,
         "U": 20, "V": 23, "W": 31, "Y": 28, "Z": 28, "AA": 20}

# 校數與預設分層（SKILL.md「分層比例」）
QUOTA = {
    ("初選表", None): (6, {"夢幻": 2, "合適": 2, "低風險": 2}),
    ("終選校表", "英國"): (10, {"夢幻": 3, "合適": 4, "低風險": 3}),
    ("終選校表", "澳洲"): (10, {"夢幻": 3, "合適": 4, "低風險": 3}),
    ("終選校表", "美國"): (20, {"夢幻": 6, "合適": 7, "低風險": 7}),
    ("終選校表", "加拿大"): (20, {"夢幻": 6, "合適": 7, "低風險": 7}),
}
# 每所學校一定要有值的欄位；查不到時要明確寫「官網未列」，不能空白
REQUIRED = ["學校", "科系", "在校成績門檻", "TOEFL", "Deadline", "學費"]
# 終選校表另外要寫落點依據（初選表沿用公司範例格式，沒有這一欄）
REQUIRED_FINAL = ["備註/選擇"]


def key(h):
    """欄位短名，例如 '科系(&LINK)' -> '科系'，TOEFL 欄 -> 'TOEFL'，'學費(會漲價)' -> '學費'。"""
    return h.split("\n")[0].split("(")[0].strip()


def text_of(v):
    return v[0] if isinstance(v, (list, tuple)) else v


def check(rows, sheet, country):
    """交付前檢查（SKILL.md「填寫規則」「資料查證規則」）。回傳問題清單，空清單代表通過。"""
    problems = []
    n, quota = QUOTA.get((sheet, None if sheet == "初選表" else country), (None, None))
    if n is not None:
        if len(rows) != n:
            problems.append(f"{sheet}應為 {n} 所，目前 {len(rows)} 所")
        got = {t: sum(r.get("分層") == t for r in rows) for t in TIER_FILL}
        if got != quota:
            problems.append(f"分層 {got} 與預設 {quota} 不同（若是依學生要求調整，請在回覆中說明）")
    if sheet == "終選校表" and sum(r.get("分層") == "低風險" for r in rows) < 2:
        problems.append("終選校表至少要保留 2 所低風險")
    for r in rows:
        name = text_of(r.get("學校", "?"))
        if r.get("分層") not in TIER_FILL:
            problems.append(f"{name}：分層只能是 夢幻／合適／低風險")
        for f in REQUIRED + (REQUIRED_FINAL if sheet == "終選校表" else []):
            if not text_of(r.get(f)):
                problems.append(f"{name}：「{f}」空白，查不到請寫「官網未列」")
        for f, v in r.items():
            t = str(text_of(v) or "")
            if t.startswith("="):
                problems.append(f"{name}：「{f}」以 = 開頭會被 Excel 當成公式，請改寫（例：40=）")
        if not isinstance(r.get("科系"), (list, tuple)):
            problems.append(f"{name}：科系要附學程官網連結")
    return problems


def build_sheet(ws, cols, rows, background, direction, freeze=None):
    for i, h in enumerate(cols, 1):
        c = ws.cell(1, i, h)
        c.fill = PatternFill("solid", fgColor=HEAD_FILL)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    order = {"夢幻": 0, "合適": 1, "低風險": 2}
    rows = sorted(rows, key=lambda r: order.get(r.get("分層"), 9))
    for r_i, row in enumerate(rows, 2):
        ws.cell(r_i, 1, bool(row.get("確定申請", False)))
        t = ws.cell(r_i, 2, row["分層"])
        t.fill = PatternFill("solid", fgColor=TIER_FILL[row["分層"]])
        for c_i, h in enumerate(cols[2:], 3):
            if not h:
                continue
            v = row.get(key(h))
            if v is None:
                continue
            cell = ws.cell(r_i, c_i)
            if isinstance(v, (list, tuple)):
                cell.value, cell.hyperlink = v[0], v[1]
                cell.font = Font(color="1155CC", underline="single")
            else:
                cell.value = v
        for c_i in range(1, len(cols) + 1):
            ws.cell(r_i, c_i).alignment = Alignment(wrap_text=True, vertical="top")
    start = len(rows) + 4
    for off, (label, text) in enumerate([("學生背景", background), ("科系方向要求", direction)]):
        r = start + off * 2
        ws.cell(r, 2, label).font = Font(bold=True)
        ws.cell(r + 1, 2, text).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=3)
        ws.row_dimensions[r + 1].height = 15 * (text.count("\n") + 1)
    for col, w in WIDTH.items():
        ws.column_dimensions[col].width = w
    if freeze:
        ws.freeze_panes = freeze


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    d = json.load(open(a.input, encoding="utf-8"))
    sheet, country = d.get("sheet", "初選表"), d.get("country", "美國")

    problems = check(d["rows"], sheet, country)
    if problems:
        print("檢查未通過，未產出檔案：")
        for p in problems:
            print(" -", p)
        sys.exit(1)

    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    if sheet == "初選表":
        build_sheet(ws, INIT_COLS, d["rows"], d["student"], d["direction"])
    else:
        build_sheet(ws, FINAL_COLS, d["rows"], d["student"], d["direction"], freeze="E2")
    src = wb.create_sheet("資料來源")
    src.append(["資料類別", "來源（僅限首要來源）", "版本", "查證日期"])
    for c in src[1]:
        c.fill, c.font = PatternFill("solid", fgColor=HEAD_FILL), Font(bold=True)
    for row in d.get("sources", []):
        src.append(row)
    for col, w in zip("ABCD", [22, 50, 22, 12]):
        src.column_dimensions[col].width = w
    os.makedirs(os.path.dirname(a.out) or ".", exist_ok=True)
    wb.save(a.out)
    tiers = {t: sum(r["分層"] == t for r in d["rows"]) for t in TIER_FILL}
    print(f"檢查通過，已產出 {a.out}（{sheet}，{len(d['rows'])} 所：{tiers}）")


if __name__ == "__main__":
    main()
