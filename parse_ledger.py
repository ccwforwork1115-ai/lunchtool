import re, pdfplumber

SUMMARY_LABELS = ("本月合計", "截至本月底止累計數")

def _num(s):
    s = (s or "").replace(",", "").strip()
    return int(s) if s else 0

def parse_ledger_pdf(path):
    """回傳 (header, entries)。entries: dict(月,日,憑單,摘要,借方,貸方,餘額)"""
    entries, header = [], {}
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            m = re.search(r"(L\d{5})", text)
            if m and "code" not in header:
                header["code"] = m.group(1)
            m = re.search(r"中華民國(\d+)年度", text)
            if m and "year" not in header:
                header["year"] = int(m.group(1))
            m = re.search(r"L\d{5}\s*(\S+)", text)
            if m and "name" not in header:
                header["name"] = m.group(1)
            for tb in page.extract_tables():
                for row in tb:
                    row = [(c or "") for c in row]
                    if len(row) < 8 or row[0] in ("115年", "月"):
                        continue
                    # 併列情形：第一行是「截至本月底止累計數」，第二行才是真正的分錄
                    first = row[3].split("\n")[0].strip()
                    if first in SUMMARY_LABELS and "\n" in row[7]:
                        fixed = []
                        for i, c in enumerate(row):
                            lines = c.split("\n")
                            if i == 3:
                                fixed.append("".join(lines[1:]))   # 摘要自折行合併
                            elif len(lines) > 1:
                                fixed.append(lines[-1])             # 多行欄位：取最後一行
                            else:
                                fixed.append("" if i >= 4 else c)   # 單行數字欄屬於累計數
                        row = fixed
                    _add(entries, row)
    return header, entries

def _add(entries, row):
    mo, dy, no, desc, dr, cr, dc, bal = row[:8]
    desc = desc.replace("\n", "").strip()
    if desc in SUMMARY_LABELS or not desc:
        return
    if not (mo.strip() and no.strip()):
        # 前一行的月日憑單延用（合併列拆分後可能為空）
        if entries and not mo.strip():
            mo, dy, no = entries[-1]["月"], entries[-1]["日"], entries[-1]["憑單"]
        else:
            return
    entries.append(dict(月=int(mo), 日=int(dy), 憑單=no.strip(), 摘要=desc,
                        借方=_num(dr), 貸方=_num(cr),
                        餘額=(-1 if dc.strip().startswith("借") else 1) * _num(bal)))

if __name__ == "__main__":
    import sys, json
    for f in sys.argv[1:]:
        h, e = parse_ledger_pdf(f)
        print(f, h, len(e))
        for x in e: print(x)
