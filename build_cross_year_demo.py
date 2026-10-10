"""Generate explicitly synthetic PDF fixtures from last year's combined lunch ledger."""
import sys, json, re
from pathlib import Path
sys.path.insert(0, 'tmp/python_deps')
import xlrd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pdfplumber

OUT = Path('跨年度測試'); OUT.mkdir(exist_ok=True)
SOURCE = r'C:\Users\asd11\Desktop\114學年度午餐會計帳簿及結算表-更新.xls'
b = xlrd.open_workbook(SOURCE)
cats = ['主食','副食','食油','調味品','人事費','燃料費(水電)','維護設備費','雜支']
balance = int(b.sheet_by_name('114-09分類帳').cell_value(3,5))
opening = balance
years = {115: [],116: []}
# Artificial July opening lets the converter establish the August opening balance.
years[115].append([7,31,'H90000','測試假設：學年度開始前結餘',0,balance,balance])
expected = {}
for name, year, month in [('114-09分類帳',115,9),('114-10分類帳',115,10),('114-11分類帳',115,11),('114-12分類帳',115,12),('115-1分類帳',116,1)]:
    sheet = b.sheet_by_name(name)
    if year == 116:
        years[year].append([1,1,'*J00000','(上期結轉)',0,balance,balance])
    totals = dict(income=0,other=0,refund=0,expense=0,count=0,opening=balance)
    for i in range(4,sheet.nrows):
        row = sheet.row_values(i)
        if not isinstance(row[0],(float,int)) or str(row[2]).strip() not in ['收','支']: continue
        amount = int(row[5] or 0)
        expenses = [int(x or 0) for x in row[6:14]]
        debit = sum(expenses) + max(0,-amount); credit = max(0,amount)
        desc = re.sub(r'11[345]',lambda m:str(int(m[0])+1),str(row[4]))
        desc = re.sub(r'[\u4e00-\u9fff]{2,4}代支','代支',desc)
        if sum(expenses): desc = '-'+cats[next(j for j,v in enumerate(expenses) if v)]+'-'+desc
        # Refund descriptions contain class hyphens; normalize for the existing refund rule.
        elif debit: desc = '午餐費退費（測試示意）'
        balance += credit-debit
        voucher = ('H' if credit else 'I')+str(int(row[3])).zfill(5)
        years[year].append([month,int(row[1]),voucher,desc,debit,credit,balance])
        totals['count'] += 1
        if sum(expenses): totals['expense'] += sum(expenses)
        elif debit: totals['refund'] += debit
        elif '利息' in desc: totals['other'] += credit
        else: totals['income'] += credit
    totals['balance'] = balance
    expected[str(month)] = totals

pdfmetrics.registerFont(TTFont('MSung-Light',r'C:\Windows\Fonts\kaiu.ttf'))
paths=[]
for year, rows in years.items():
    path=OUT/f'測試示意_{year}年_免費學生午餐補助款.pdf'
    c=canvas.Canvas(str(path),pagesize=(842,595))
    c.setTitle(f'測試示意 {year}年度；非正式帳冊')
    for offset in range(0,len(rows),16):
        c.setFont('MSung-Light',13)
        c.drawCentredString(421,561,'嘉義縣民雄鄉興中國民小學')
        c.drawCentredString(421,539,f'中華民國{year}年度')
        c.drawCentredString(421,517,'2123 應付代收款明細分類帳')
        c.setFont('MSung-Light',10);c.drawString(60,494,'L30076 免費學生午餐補助款')
        c.setFillColorRGB(.7,0,0); c.drawString(60,476,'測試示意：依舊午餐帳簿平移日期、合併為測試子目；非正式會計紀錄')
        c.setFillColorRGB(0,0,0)
        edges=[55,80,105,153,560,620,685,730,820]
        for x in edges: c.line(x,90,x,460)
        c.line(55,460,820,460)
        for x,t in zip([60,85,108,160,565,625,690,740],['月','日','傳票號數','摘要','借方','貸方','借或貸','餘額']): c.drawString(x,445,t)
        c.line(55,435,820,435)
        for n,row in enumerate(rows[offset:offset+16]):
            m,d,v,desc,dr,cr,bal=row; y=419-n*20
            c.setFont('MSung-Light',9)
            for x,t in [(60,f'{m:02}'),(85,f'{d:02}'),(108,v),(160,desc[:43]),(565,f'{dr:,}' if dr else ''),(625,f'{cr:,}' if cr else ''),(690,'借' if bal<0 else '貸'),(740,f'{abs(bal):,}')]: c.drawString(x,y,t)
            c.line(55,y-6,820,y-6)
        c.drawString(60,65,'測試用途；金額依逐筆收支重算。傳票號碼、日期、子目歸屬為模擬。')
        c.showPage()
    c.save(); paths.append(str(path))
    with pdfplumber.open(path) as p: p.pages[0].to_image(resolution=110).save(OUT/f'{year}預覽.png')
(OUT/'expected.json').write_text(json.dumps({'source':SOURCE,'opening':opening,'months':expected,'pdfs':paths},ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(expected,ensure_ascii=False,indent=2))
