import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

OUT = "/home/user/-/hualalai/Hualalai_収益管理表_2026-11.xlsx"
FONT = "Yu Gothic"
BLUE = "0000FF"; BLACK = "000000"; GREEN = "008000"
YEL = PatternFill("solid", fgColor="FFFF00")
HEAD = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E1F2")
GREY = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
NUM = '#,##0;-#,##0;"-"'
PCT = '0.0%;-0.0%;"-"'
DATEF = 'm/d'

def f(color=BLACK, bold=False, size=10, white=False):
    return Font(name=FONT, size=size, bold=bold, color="FFFFFF" if white else color)

def cell(ws, ref, value=None, color=BLACK, bold=False, fmt=None, fill=None, align=None, size=10, border=True, white=False):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = f(color, bold, size, white)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if align: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    else: c.alignment = Alignment(vertical="center", wrap_text=True)
    if border: c.border = BORDER
    return c

wb = Workbook()

# ---------------- 設定 ----------------
st = wb.active; st.title = "設定"
st.column_dimensions["A"].width = 30; st.column_dimensions["B"].width = 18; st.column_dimensions["C"].width = 70
cell(st, "A1", "Hualalai 収益管理表 ― 設定", bold=True, size=14, border=False)
cell(st, "A2", "黄色セル（青字）が入力欄です。黒字は数式のため編集不要です。", border=False)
rows = [
    ("店舗名", "1階カフェ Hualalai", None, "日報の店舗名"),
    ("対象月（月初日）", dt.date(2026, 11, 1), DATEF.replace('m/d', 'yyyy/m/d'), "月を変える場合はこの日付を変更（日次入力シートの日付が連動）"),
    ("月間目標売上（税抜・円）", 2000000, NUM, "11/5日報「目標売上：2,000,000円」"),
    ("当月の営業予定日数", 30, NUM, "仮置き（要確認）。月次サマリーの「残営業日数」「必要日商」に使用"),
    ("前日までの累計売上（繰越・円）", 211383, NUM, "11/5日報「累計売上 267,207円」−「当日売上 55,824円」から逆算。11/4以前の日別内訳が分かれば、日次入力に転記しここを0にする"),
    ("前日までの稼働日数（繰越・日）", 2, NUM, "11/5日報「稼働日数 3日」−1日。上と同様に内訳入力後は0にする"),
    ("前日までの累計来客数（繰越・名）", 0, NUM, "11/4以前の来客数は日報に記載なし。分かれば入力"),
]
r = 4
cell(st, "A3", "項目", bold=True, fill=SUB); cell(st, "B3", "値", bold=True, fill=SUB); cell(st, "C3", "備考・出典", bold=True, fill=SUB)
for name, val, fmt, note in rows:
    cell(st, f"A{r}", name, bold=True, fill=GREY)
    cell(st, f"B{r}", val, color=BLUE, fmt=fmt, fill=YEL, align="right")
    cell(st, f"C{r}", note)
    r += 1
# named positions
S_MONTH = "設定!$B$5"; S_TARGET = "設定!$B$6"; S_DAYS = "設定!$B$7"; S_CARRY_SALES = "設定!$B$8"; S_CARRY_DAYS = "設定!$B$9"; S_CARRY_CUST = "設定!$B$10"

r += 1
cell(st, f"A{r}", "時間帯区分", bold=True, fill=SUB); cell(st, f"B{r}", "時間", bold=True, fill=SUB); cell(st, f"C{r}", "部門区分", bold=True, fill=SUB); r += 1
for a, b, c in [("モーニング", "7:00-11:00", "社販（社員販売）"), ("ランチ", "11:00-14:00", "一般"), ("カフェ", "14:00-17:00", "TO（テイクアウト）"), ("ディナー", "17:00-21:00", "物販・その他（社販gugut等）")]:
    cell(st, f"A{r}", a); cell(st, f"B{r}", b); cell(st, f"C{r}", c); r += 1
r += 1
cell(st, f"A{r}", "日報からの転記手順", bold=True, size=11, border=False); r += 1
for t in [
    "1. 「日次入力」シートで該当日の行を探す（日付は設定の対象月から自動生成）。",
    "2. 日報【売上実績】の当日売上・来客数を入力。客単価・累計・達成率は自動計算。",
    "3. 【時間帯別】【部門別】【物販・その他】の金額・人数・数量を入力。",
    "4. 「時間帯差異」「部門差異」が 0 以外なら転記ミスか日報側の不一致（11/5は部門合計が当日売上より232円少ない）。",
    "5. 「月次サマリー」で目標達成率・必要日商・時間帯/部門構成を確認。",
]:
    cell(st, f"A{r}", t, border=False); st.merge_cells(f"A{r}:C{r}"); r += 1

# ---------------- 日次入力 ----------------
ws = wb.create_sheet("日次入力")
HR = 4  # header row
FIRST = 5; NDAYS = 31; LAST = FIRST + NDAYS - 1
cols = [
    ("日付", 8), ("曜日", 5), ("営業", 5), ("当日売上\n(円)", 11), ("来客数\n(名)", 8), ("客単価\n(円)", 9),
    ("ﾓｰﾆﾝｸﾞ\n売上", 10), ("ﾓｰﾆﾝｸﾞ\n客数", 7), ("ランチ\n売上", 10), ("ランチ\n客数", 7), ("カフェ\n売上", 10), ("カフェ\n客数", 7), ("ディナー\n売上", 10), ("ディナー\n客数", 7),
    ("時間帯\n合計", 10), ("時間帯\n差異", 8),
    ("社販\n売上", 10), ("社販\n数量", 7), ("一般\n売上", 10), ("一般\n数量", 7), ("TO\n売上", 10), ("TO\n数量", 7), ("物販・他\n売上", 10), ("物販・他\n件数", 7),
    ("部門\n合計", 10), ("部門\n差異", 8),
    ("累計売上\n(円)", 12), ("累計\n稼働日", 7), ("平均日商\n(円)", 11), ("目標\n達成率", 8), ("メモ", 60),
]
cell(ws, "A1", "Hualalai 日次収益入力表（金額は全て税抜）", bold=True, size=14, border=False)
cell(ws, "A2", "黄色セル（青字）＝日報から転記する入力欄。黒字＝自動計算。「営業」は売上がある日に自動で1。", border=False)
# group headers row 3
groups = [("A", "F", "売上実績"), ("G", "P", "時間帯別"), ("Q", "Z", "部門別"), ("AA", "AD", "累計・目標"), ("AE", "AE", "")]
for a, b, t in groups:
    cell(ws, f"{a}3", t, bold=True, fill=HEAD, white=True, align="center")
    if a != b: ws.merge_cells(f"{a}3:{b}3")
    for ci in range(ws[f"{a}3"].column, ws[f"{b}3"].column + 1):
        ws.cell(row=3, column=ci).fill = HEAD; ws.cell(row=3, column=ci).border = BORDER
for i, (h, w) in enumerate(cols, start=1):
    L = get_column_letter(i)
    ws.column_dimensions[L].width = w
    cell(ws, f"{L}{HR}", h, bold=True, fill=SUB, align="center")
ws.row_dimensions[HR].height = 32
INPUT_COLS = ["D", "E", "G", "H", "I", "J", "K", "L", "M", "N", "Q", "R", "S", "T", "U", "V", "W", "X", "AE"]

for k in range(NDAYS):
    r = FIRST + k
    # date: month start + k, blank if beyond month end
    if k == 0:
        cell(ws, f"A{r}", f"={S_MONTH}", color=GREEN, fmt=DATEF, align="center")
    else:
        cell(ws, f"A{r}", f'=IF(A{r-1}="","",IF(MONTH(A{r-1}+1)=MONTH({S_MONTH}),A{r-1}+1,""))', color=GREEN, fmt=DATEF, align="center")
    cell(ws, f"B{r}", f'=IF(A{r}="","",CHOOSE(WEEKDAY(A{r}),"日","月","火","水","木","金","土"))', align="center")
    cell(ws, f"C{r}", f'=IF(N(D{r})>0,1,0)', fmt=NUM, align="center")
    for L in INPUT_COLS:
        cell(ws, f"{L}{r}", None, color=BLUE, fmt=None if L == "AE" else NUM, fill=YEL, align=None if L == "AE" else "right")
    cell(ws, f"F{r}", f'=IF(N(E{r})>0,D{r}/E{r},0)', fmt=NUM)
    cell(ws, f"O{r}", f'=SUM(G{r},I{r},K{r},M{r})', fmt=NUM)
    cell(ws, f"P{r}", f'=IF(N(D{r})=0,0,D{r}-O{r})', fmt=NUM)
    cell(ws, f"Y{r}", f'=SUM(Q{r},S{r},U{r},W{r})', fmt=NUM)
    cell(ws, f"Z{r}", f'=IF(N(D{r})=0,0,D{r}-Y{r})', fmt=NUM)
    cell(ws, f"AA{r}", f'=IF(A{r}="","",{S_CARRY_SALES}+SUM($D${FIRST}:D{r}))', fmt=NUM)
    cell(ws, f"AB{r}", f'=IF(A{r}="","",{S_CARRY_DAYS}+SUM($C${FIRST}:C{r}))', fmt=NUM)
    cell(ws, f"AC{r}", f'=IF(A{r}="","",IF(AB{r}>0,AA{r}/AB{r},0))', fmt=NUM)
    cell(ws, f"AD{r}", f'=IF(A{r}="","",IF({S_TARGET}>0,AA{r}/{S_TARGET},0))', fmt=PCT)
# totals row
TR = LAST + 1
cell(ws, f"A{TR}", "月合計", bold=True, fill=SUB, align="center"); ws.merge_cells(f"A{TR}:B{TR}"); ws[f"B{TR}"].fill = SUB; ws[f"B{TR}"].border = BORDER
for L in ["C", "D", "E", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]:
    cell(ws, f"{L}{TR}", f'=SUM({L}{FIRST}:{L}{LAST})', bold=True, fmt=NUM, fill=SUB)
cell(ws, f"F{TR}", f'=IF(E{TR}>0,D{TR}/E{TR},0)', bold=True, fmt=NUM, fill=SUB)
for L in ["AA", "AB", "AC", "AD", "AE"]:
    cell(ws, f"{L}{TR}", None, fill=SUB)
cell(ws, f"AE{TR}", "※月合計は入力済みの日のみ（繰越分は含まない）", fill=SUB)
ws.freeze_panes = "D5"

# 11/5 data (row index: 11/5 -> k=4 -> row 9)
r = FIRST + 4
data = {"D": 55824, "E": 74, "G": 21821, "H": 28, "I": 21365, "J": 24, "K": 12638, "L": 22, "M": 0, "N": 0,
        "Q": 6629, "R": 20, "S": 41547, "T": 57, "U": 1416, "V": 2, "W": 6000, "X": 2}
for L, v in data.items():
    ws[f"{L}{r}"].value = v
ws[f"AE{r}"].value = ("どの時間帯にも満遍なく来店。フードはドッグ・カレー合わせて12件、スイーツ6件。"
                      "社販gugut 6,000円/2件。※部門合計49,592円+物販6,000円=55,592円で当日売上より232円少（日報どおり転記）")
ws[f"S{r}"].comment = Comment("日報の表記は「41,5447円」。単価729円×57点から41,547円と判断。", "Claude")
ws[f"D{r}"].comment = Comment("出典：11/5分 Hualalai 業務日報・売上報告（榎本、11/6）", "Claude")

# ---------------- 月次サマリー ----------------
sm = wb.create_sheet("月次サマリー")
for L, w in zip("ABCDEFG", [22, 14, 12, 12, 12, 12, 12]):
    sm.column_dimensions[L].width = w
cell(sm, "A1", "Hualalai 月次収益サマリー", bold=True, size=14, border=False)
cell(sm, "A2", f'="対象月："&TEXT({S_MONTH},"yyyy年m月")&"　（金額は全て税抜・円）"', border=False)
D = "日次入力"
kpi = [
    ("月間目標売上", f"={S_TARGET}", NUM, GREEN),
    ("累計売上", f"={S_CARRY_SALES}+{D}!D{TR}", NUM, BLACK),
    ("目標達成率", f"=IF(B5>0,B6/B5,0)", PCT, BLACK),
    ("目標までの残額", f"=MAX(B5-B6,0)", NUM, BLACK),
    ("稼働日数", f"={S_CARRY_DAYS}+{D}!C{TR}", NUM, BLACK),
    ("平均日商（累計売上÷稼働日数）", f"=IF(B9>0,B6/B9,0)", NUM, BLACK),
    ("残営業日数（予定−稼働）", f"=MAX({S_DAYS}-B9,0)", NUM, BLACK),
    ("必要日商（残額÷残営業日数）", f"=IF(B11>0,B8/B11,0)", NUM, BLACK),
    ("累計来客数", f"={S_CARRY_CUST}+{D}!E{TR}", NUM, BLACK),
    ("平均客単価（入力済み日：売上÷来客数）", f"=IF({D}!E{TR}>0,{D}!D{TR}/{D}!E{TR},0)", NUM, BLACK),
    ("目標ペース（目標÷営業予定日数）", f"=IF({S_DAYS}>0,B5/{S_DAYS},0)", NUM, BLACK),
    ("日商のペース差（平均日商−目標ペース）", f"=B10-B15", NUM, BLACK),
]
cell(sm, "A4", "KPI", bold=True, fill=SUB); cell(sm, "B4", "値", bold=True, fill=SUB)
for i, (n, fm, fmt, col) in enumerate(kpi):
    r = 5 + i
    cell(sm, f"A{r}", n, fill=GREY); cell(sm, f"B{r}", fm, color=col, fmt=fmt, align="right")
sm["B14"].comment = Comment("繰越来客数は日報に記載がないため、日次入力に転記済みの日のみで算出。", "Claude")

# 時間帯別
r0 = 19
cell(sm, f"A{r0}", "時間帯別（入力済み日の合計）", bold=True, size=11, border=False)
hdr = ["時間帯", "売上", "客数", "売上構成比", "客単価"]
for i, h in enumerate(hdr):
    cell(sm, f"{get_column_letter(i+1)}{r0+1}", h, bold=True, fill=SUB, align="center")
bands = [("モーニング 7-11", "G", "H"), ("ランチ 11-14", "I", "J"), ("カフェ 14-17", "K", "L"), ("ディナー 17-21", "M", "N")]
for i, (n, sc, cc) in enumerate(bands):
    r = r0 + 2 + i
    cell(sm, f"A{r}", n, fill=GREY)
    cell(sm, f"B{r}", f"={D}!{sc}{TR}", color=GREEN, fmt=NUM)
    cell(sm, f"C{r}", f"={D}!{cc}{TR}", color=GREEN, fmt=NUM)
    cell(sm, f"D{r}", f"=IF($B${r0+6}>0,B{r}/$B${r0+6},0)", fmt=PCT)
    cell(sm, f"E{r}", f"=IF(C{r}>0,B{r}/C{r},0)", fmt=NUM)
r = r0 + 6
cell(sm, f"A{r}", "合計", bold=True, fill=SUB)
cell(sm, f"B{r}", f"=SUM(B{r0+2}:B{r0+5})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"C{r}", f"=SUM(C{r0+2}:C{r0+5})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"D{r}", f"=IF(B{r}>0,SUM(D{r0+2}:D{r0+5}),0)", bold=True, fmt=PCT, fill=SUB)
cell(sm, f"E{r}", f"=IF(C{r}>0,B{r}/C{r},0)", bold=True, fmt=NUM, fill=SUB)

# 部門別
r1 = r0 + 8
cell(sm, f"A{r1}", "部門別（入力済み日の合計）", bold=True, size=11, border=False)
hdr = ["部門", "売上", "数量・件数", "売上構成比", "単価"]
for i, h in enumerate(hdr):
    cell(sm, f"{get_column_letter(i+1)}{r1+1}", h, bold=True, fill=SUB, align="center")
deps = [("社販", "Q", "R"), ("一般", "S", "T"), ("TO（テイクアウト）", "U", "V"), ("物販・その他", "W", "X")]
for i, (n, sc, cc) in enumerate(deps):
    r = r1 + 2 + i
    cell(sm, f"A{r}", n, fill=GREY)
    cell(sm, f"B{r}", f"={D}!{sc}{TR}", color=GREEN, fmt=NUM)
    cell(sm, f"C{r}", f"={D}!{cc}{TR}", color=GREEN, fmt=NUM)
    cell(sm, f"D{r}", f"=IF($B${r1+6}>0,B{r}/$B${r1+6},0)", fmt=PCT)
    cell(sm, f"E{r}", f"=IF(C{r}>0,B{r}/C{r},0)", fmt=NUM)
r = r1 + 6
cell(sm, f"A{r}", "合計", bold=True, fill=SUB)
cell(sm, f"B{r}", f"=SUM(B{r1+2}:B{r1+5})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"C{r}", f"=SUM(C{r1+2}:C{r1+5})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"D{r}", f"=IF(B{r}>0,SUM(D{r1+2}:D{r1+5}),0)", bold=True, fmt=PCT, fill=SUB)
cell(sm, f"E{r}", f"=IF(C{r}>0,B{r}/C{r},0)", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"A{r+1}", f'="当日売上との差異（部門合計−売上）："&TEXT({D}!Z{TR}*-1,"#,##0")&"円"', border=False)

# 曜日別
r2 = r1 + 9
cell(sm, f"A{r2}", "曜日別（入力済み日）", bold=True, size=11, border=False)
hdr = ["曜日", "売上合計", "稼働日数", "平均日商", "来客数", "客単価"]
for i, h in enumerate(hdr):
    cell(sm, f"{get_column_letter(i+1)}{r2+1}", h, bold=True, fill=SUB, align="center")
for i, w in enumerate(["月", "火", "水", "木", "金", "土", "日"]):
    r = r2 + 2 + i
    rngB = f"{D}!$B${FIRST}:$B${LAST}"
    cell(sm, f"A{r}", w, fill=GREY, align="center")
    cell(sm, f"B{r}", f'=SUMIF({rngB},A{r},{D}!$D${FIRST}:$D${LAST})', fmt=NUM)
    cell(sm, f"C{r}", f'=SUMIF({rngB},A{r},{D}!$C${FIRST}:$C${LAST})', fmt=NUM)
    cell(sm, f"D{r}", f'=IF(C{r}>0,B{r}/C{r},0)', fmt=NUM)
    cell(sm, f"E{r}", f'=SUMIF({rngB},A{r},{D}!$E${FIRST}:$E${LAST})', fmt=NUM)
    cell(sm, f"F{r}", f'=IF(E{r}>0,B{r}/E{r},0)', fmt=NUM)
r = r2 + 9
cell(sm, f"A{r}", "合計", bold=True, fill=SUB)
cell(sm, f"B{r}", f"=SUM(B{r2+2}:B{r2+8})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"C{r}", f"=SUM(C{r2+2}:C{r2+8})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"D{r}", f"=IF(C{r}>0,B{r}/C{r},0)", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"E{r}", f"=SUM(E{r2+2}:E{r2+8})", bold=True, fmt=NUM, fill=SUB)
cell(sm, f"F{r}", f"=IF(E{r}>0,B{r}/E{r},0)", bold=True, fmt=NUM, fill=SUB)

cell(sm, f"A{r+2}", "凡例：青字＝入力値／緑字＝他シート参照／黒字＝数式。入力は「設定」「日次入力」の黄色セルのみ。", border=False)
sm.freeze_panes = "A4"

wb.move_sheet("月次サマリー", offset=-2)  # summary first
wb.save(OUT)
print("saved", OUT)
