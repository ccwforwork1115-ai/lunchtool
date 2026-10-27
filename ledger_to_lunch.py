# -*- coding: utf-8 -*-
"""
明細帳(PDF) → 學校午餐費明細分類帳(Excel) 自動轉入

用法：
  python ledger_to_lunch.py 午餐結算明細表格式.xlsx L30076.pdf L20022.pdf -o 午餐結算明細表_已填.xlsx

規則重點：
  1. 一個會計子目(L碼)的 PDF = 活頁簿中的一個「基金區塊」(免費學生午餐補助款…)
  2. 同月同日同憑單號的支出，合併成一列，金額依「科目」分散到 主食/副食/食油… 各欄
  3. 科目優先取摘要中「-科目-」格式，其次用關鍵字推定(會列入報告請你確認)
  4. 本學年度開始(8/1)以前的收支，轉為 08分類帳「上月份結餘額轉」
  5. 午餐退費 → 收入欄填負數，並歸入「午餐退費收入減帳」
"""
import argparse, csv, os, re, shutil, subprocess, sys, tempfile
from collections import OrderedDict
import openpyxl
from copy import copy
from parse_ledger import parse_ledger_pdf

# ───────────── 設定區（可依實際情形修改） ─────────────
SECTION_TITLES = ["免費學生午餐補助款", "教職員及其他相關午餐收費", "精進學校菜單方案補助款",
                  "午餐維運經費", "偏鄉學校午餐補助款", "小型偏遠學校午餐補助費", "三章一Q獎勵金"]

# 會計子目(L碼) → 活頁簿區塊名稱。名稱與區塊相同的子目會自動對應，不用列在這裡
FUND_MAP = {
    "L30076": "免費學生午餐補助款",
    "L20022": "教職員及其他相關午餐收費",
    "L30077": "精進學校菜單方案補助款",
    # "Lxxxxx": "午餐維運經費",
}

CATEGORY_COLS = OrderedDict([("主食", "G"), ("副食", "H"), ("食油", "I"), ("調味品", "J"),
                             ("人事費", "K"), ("燃料費(水電)", "L"), ("維護設備費", "M"), ("雜支", "N")])
CATEGORY_ALIAS = {"設備維護費": "維護設備費", "維護設備": "維護設備費", "燃料費": "燃料費(水電)",
                  "水電": "燃料費(水電)", "水電費": "燃料費(水電)", "燃料費（水電）": "燃料費(水電)"}

# 摘要沒有「-科目-」時，用關鍵字推定（順序即優先順序）
KEYWORD_RULES = [
    (("勞保", "健保", "薪資", "工資", "年終", "勞退", "二代健保", "體檢"), "人事費"),
    (("電費", "水費", "瓦斯費", "燃料", "水電"), "燃料費(水電)"),
    (("維修", "修繕", "維護", "設備"), "維護設備費"),
    (("米糧", "白米", "麵", "主食"), "主食"),
    (("沙拉油", "食油", "橄欖油"), "食油"),
    (("調味", "醬油", "鹽", "糖"), "調味品"),
    (("肉", "蔬", "菜", "魚", "蛋", "豆", "水果", "副食"), "副食"),
    (("手套", "洗潔", "雜支", "出席費", "清潔"), "雜支"),
]
DEFAULT_CATEGORY = "雜支"           # 都比對不到時
OTHER_INCOME_KEYWORDS = ("利息", "租金", "回收", "賠償", "其他收入", "捐")
REFUND_KEYWORDS = ("退", "退費")    # 借方且摘要含「退」「午餐費」→ 午餐退費收入減帳
ENTRY_ROW_COUNT = 6                 # 每個區塊每月可填的列數
PRE_YEAR_AS_OPENING = True          # 學年開始前收入 → 08月「上月份結餘額轉」；False 則報錯提醒
# ──────────────────────────────────────────────────────


def norm(s):
    return re.sub(r"[\s\n]", "", str(s or ""))


def ensure_xlsx(path, workdir):
    if path.lower().endswith(".xlsx"):
        return path
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("範本是 .xls，請先用 Excel 另存為 .xlsx 再執行（或安裝 LibreOffice）。")
    subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", workdir, path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + ".xlsx")


def find_sections(ws):
    """回傳 {區塊名稱: dict(title, opening, first, last, sub, cum, summary)}，由 E 欄標題掃描得出"""
    secs, summary_rows = {}, {}
    for r in range(1, ws.max_row + 1):
        v = norm(ws.cell(r, 5).value)
        if v in SECTION_TITLES and v not in secs and "結餘額轉" in norm(ws.cell(r + 2, 5).value):
            secs[v] = dict(title=r, opening=r + 2, first=r + 3)
        elif v in SECTION_TITLES and v not in summary_rows and v in secs:
            summary_rows[v] = r            # 下方「收入分類」表
    for name, s in secs.items():
        r = s["first"]
        while norm(ws.cell(r, 5).value) != "以下空白":
            r += 1
        s["last"] = r - 1
        s["sub"], s["cum"] = r + 1, r + 2
        s["summary"] = summary_rows.get(name)
    return secs


def pick_category(desc):
    m = re.search(r"-([^-]+)-", desc)
    if m:
        c = CATEGORY_ALIAS.get(m.group(1), m.group(1))
        if c in CATEGORY_COLS:
            return c, "明確(摘要)", desc[m.end():]
    for kws, cat in KEYWORD_RULES:
        if any(k in desc for k in kws):
            return cat, "關鍵字推定，請確認", desc
    return DEFAULT_CATEGORY, "無法判斷，暫列雜支，請確認", desc


def classify(e):
    """回傳 (類型, 科目, 備註, 明細摘要)；類型: income/other_income/refund/expense"""
    d = e["摘要"]
    if e["貸方"]:
        kind = "other_income" if any(k in d for k in OTHER_INCOME_KEYWORDS) else "income"
        return kind, None, "", d
    if "午餐費" in d and any(k in d for k in REFUND_KEYWORDS) and "-" not in d:
        return "refund", None, "", d
    cat, how, detail = pick_category(d)
    return "expense", cat, how, detail


def build_rows(entries, year, sy):
    """依學年度切分：回傳 (opening, {月: [列]})，列=dict(月,日,字,號,摘要,F,cats,type,src)"""
    opening, pre_income_found = 0, False
    months, report = OrderedDict(), []
    for e in entries:
        m = e["月"]
        in_sy = (year == sy and m >= 8) or (year == sy + 1 and m <= 7)
        pre = (year == sy and m < 8)
        if pre:
            opening = e["餘額"]
            pre_income_found = pre_income_found or bool(e["貸方"])
            report.append((e, "學年開始前 → 併入 08月「上月份結餘額轉」", None))
            continue
        if not in_sy:
            report.append((e, "不屬本學年度，略過", None))
            continue
        kind, cat, how, detail = classify(e)
        months.setdefault(m, []).append((e, kind, cat, how, detail))
    if pre_income_found and not PRE_YEAR_AS_OPENING:
        sys.exit("有學年開始前的收入，請決定記入方式。")
    return opening, months, report


def group_month(items):
    """同月日憑單的支出合併成一列；收入/退費各自一列"""
    rows, exp = [], OrderedDict()
    for e, kind, cat, how, detail in items:
        if kind == "expense":
            key = (e["月"], e["日"], e["憑單"])
            g = exp.setdefault(key, dict(月=e["月"], 日=e["日"], 憑單=e["憑單"], cats=OrderedDict(),
                                         type="expense", F=None, details=[], src=[]))
            g["cats"][cat] = g["cats"].get(cat, 0) + e["借方"]
            g["details"].append(detail)
            g["src"].append((e, cat, how))
            if key not in [r["_k"] for r in rows if "_k" in r]:
                g["_k"] = key
                rows.append(g)
        else:
            amt = e["貸方"] if kind != "refund" else -e["借方"]
            rows.append(dict(月=e["月"], 日=e["日"], 憑單=e["憑單"], cats={}, type=kind, F=amt,
                             summary=e["摘要"], src=[(e, "", "")]))
    for g in rows:
        if g["type"] == "expense":
            cats = [c for c in CATEGORY_COLS if c in g["cats"]]
            g["summary"] = (g["details"][0] if len(g["details"]) == 1 and g["details"][0].startswith("付")
                            else "付" + "、".join(cats) + "等費用")
    rows.sort(key=lambda r: (r["日"], r["憑單"], r["type"] != "income"))
    return rows


def write_section(ws, sec, rows, report_rows, label):
    cap = sec["last"] - sec["first"] + 1
    if len(rows) > cap:
        sys.exit(f"【{label}】{ws.title} 需要 {len(rows)} 列，但範本每區塊只有 {cap} 列。"
                 f"請在該區塊插入列，或把同日憑單合併。")
    for r in range(sec["first"], sec["last"] + 1):                # 清除舊資料（保留 O、P 公式）
        for col in "ABCDEFGHIJKLMN":
            ws[f"{col}{r}"].value = None
    income_cells = {"income": [], "other_income": [], "refund": []}
    for i, g in enumerate(rows):
        r = sec["first"] + i
        ws[f"A{r}"], ws[f"B{r}"] = g["月"], g["日"]
        ws[f"C{r}"], ws[f"D{r}"] = g["憑單"][0], int(g["憑單"][1:])
        ws[f"E{r}"] = g["summary"]
        if g["F"] is not None:
            ws[f"F{r}"] = g["F"]
            income_cells[g["type"]].append(f"F{r}")
        for c, amt in g["cats"].items():
            ws[f"{CATEGORY_COLS[c]}{r}"] = amt
        for e, c, how in g["src"]:
            report_rows.append((e, f"{ws.title} 第{r}列" + (f" → {c}" if c else " → 收入欄"), how))
    if sec["summary"]:
        sr = sec["summary"]
        for col, key in (("G", "income"), ("H", "other_income"), ("I", "refund")):
            cells = income_cells[key]
            ws[f"{col}{sr}"] = ("=" + "+".join(cells)) if cells else None


def add_opening_note(wb, fund, pre_entries, opening):
    """學年開始前收入轉為期初時，註記在 08結算 備註欄"""
    cell = wb["08結算"]["B16"]
    inc = [e for e in pre_entries if e["貸方"]]
    if not opening or not inc:
        return
    src = "、".join(f"{e['月']}月{e['日']}日({e['憑單']})收入{e['貸方']:,}元" for e in inc)
    line = f"三、上學年度結餘額轉含【{fund}】{opening:,}元，係學年度開始前{src}，轉入本學年度。"
    v = str(cell.value or "")
    if line not in v:
        cell.value = v.rstrip("\n") + "\n" + line


def clean_stray_refs(ws, secs):
    """範本「收入分類」表(G~N)若有指向其他區塊的殘留公式（例如維運經費列的 =F21），會重複計入，予以清除"""
    cleared = []
    for name, sec in secs.items():
        sr = sec["summary"]
        if not sr:
            continue
        for col in "GHIJKLMN":
            v = ws[f"{col}{sr}"].value
            if isinstance(v, str) and v.startswith("="):
                rows = [int(x) for x in re.findall(r"[A-Z]+(\d+)", v)]
                if any(not (sec["title"] <= r <= sec["cum"]) for r in rows):
                    ws[f"{col}{sr}"].value = None
                    cleared.append(f"{ws.title}!{col}{sr}（原 {v}）")
    return cleared


def recalc_and_check(out_path, expectations):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("（未安裝 LibreOffice，略過自動驗算；請用 Excel 開啟後核對各區塊「截至本月份累計數」餘額）")
        return True
    tmp = tempfile.mkdtemp()
    subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", tmp, out_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    wb = openpyxl.load_workbook(os.path.join(tmp, os.path.basename(out_path)), data_only=True)
    ok = True
    print("\n【驗算】活頁簿月底餘額 vs 會計系統明細帳餘額")
    for (sheet, cum_row, label, month), want in expectations.items():
        got = wb[sheet][f"P{cum_row}"].value
        flag = "OK" if got == want else "不符!!"
        ok &= got == want
        print(f"  {label} {month}月  活頁簿 {got:,.0f}  明細帳 {want:,.0f}  {flag}")
    return ok


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("template")
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("-o", "--out", default="午餐結算明細表_已填.xlsx")
    ap.add_argument("--school-year", type=int, help="學年度，如 115（預設讀取「基本資料」）")
    a = ap.parse_args()

    work = tempfile.mkdtemp()
    wb = openpyxl.load_workbook(ensure_xlsx(a.template, work))
    sy = a.school_year or int(re.search(r"(\d+)學年度", str(wb["基本資料"]["A2"].value)).group(1))
    report, expectations, warnings, stray = [], {}, [], []

    for pdf in a.pdfs:
        hdr, entries = parse_ledger_pdf(pdf)
        code, name = hdr.get("code"), hdr.get("name")
        fund = FUND_MAP.get(code) or next((t for t in SECTION_TITLES if t == name), None)
        if not fund:
            sys.exit(f"{pdf}：子目 {code}（{name}）尚未對應到活頁簿區塊，請在 FUND_MAP 補上。")
        opening, months, rep = build_rows(entries, hdr["year"], sy)
        report += rep
        order = [8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6, 7]
        # 期初：寫入 08分類帳 的「上月份結餘額轉」
        if hdr["year"] == sy:                       # 116年度的 PDF 不動期初
            sec8 = find_sections(wb["08分類帳"])[fund]
            wb["08分類帳"][f"F{sec8['opening']}"] = opening
            add_opening_note(wb, fund, [e for e in entries if e["月"] < 8], opening)
        for m in months:
            sheet = f"{m:02d}分類帳"
            ws = wb[sheet]
            sec = find_sections(ws)[fund]
            rows = group_month(months[m])
            rr = []
            write_section(ws, sec, rows, rr, fund)
            stray += clean_stray_refs(ws, find_sections(ws))
            report += rr
            last = months[m][-1][0]["餘額"]
            expectations[(sheet, sec["cum"], fund, m)] = last
            for e, kind, cat, how, d in months[m]:
                if kind == "expense" and how != "明確(摘要)":
                    warnings.append(f"{fund} {e['月']}/{e['日']} {e['憑單']} 「{e['摘要']}」→ {cat}（{how}）")
        print(f"✔ {os.path.basename(pdf)}：{code} {name} → 區塊【{fund}】，期初 {opening:,}，"
              f"已轉入 {sum(len(v) for v in months.values())} 筆分錄")

    # Remove empty legacy function names carried over from the template.
    for name, definition in list(wb.defined_names.items()):
        value = str(definition.attr_text or '').strip()
        if not value or (name == '_xlfn_SINGLE' and value.lower() == 'none'):
            del wb.defined_names[name]
    # A4 landscape: settlements fit one page; ledgers continue vertically.
    for ws in wb.worksheets:
        settlement = bool(re.fullmatch(r'\d{2}結算', ws.title))
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1 if settlement else 0
        ws.page_setup.scale = None
        ws.page_setup.paperHeight = None
        ws.page_setup.paperWidth = None
        ws.page_setup.usePrinterDefaults = False
        for key, value in dict(left=0.25, right=0.25, top=0.3, bottom=0.3,
                               header=0.15, footer=0.15).items():
            setattr(ws.page_margins, key, value)
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None and cell.data_type in ('n', 'f'):
                    alignment = copy(cell.alignment)
                    alignment.shrinkToFit = True
                    alignment.wrapText = False
                    cell.alignment = alignment
        ws.row_breaks.brk = []
        ws.col_breaks.brk = []
    wb.save(a.out)
    csv_path = os.path.splitext(a.out)[0] + "_轉換對照.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["月", "日", "憑單", "摘要", "借方", "貸方", "去向", "科目判定"])
        for e, dest, how in report:
            w.writerow([e["月"], e["日"], e["憑單"], e["摘要"], e["借方"], e["貸方"], dest, how or ""])
    ok = recalc_and_check(a.out, expectations)
    if stray:
        print("\n【已清除範本殘留公式】" + "、".join(stray))
    if warnings:
        print("\n【請確認的科目】")
        for x in warnings:
            print("  -", x)
    print(f"\n已輸出：{a.out}\n對照表：{csv_path}")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
