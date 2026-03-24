# ══════════════════════════════════════════════
# _v42_part2b.py: 月次CF（キャッシュフロー）シート — 36ヶ月
# Part1 + Part2a の続き。wb 変数引き継ぎ前提。
# ══════════════════════════════════════════════

LAST_COL = gl(37)  # AK列 = Month36

# ──── SHEET 5: 月次CF（36ヶ月）────
ws_cf = wb.create_sheet("月次CF（36ヶ月）")
ws_cf.column_dimensions["A"].width = 28
for m in range(1, 37):
    ws_cf.column_dimensions[gl(m + 1)].width = 9

title_row(ws_cf, 1, "UNLID 月次キャッシュフロー v4.2  ─  単位：万円  |  36ヶ月展開", 37)
sc(ws_cf, 2, 1,
   "★ 月次値 = 年次値÷12（平均配分）。設備投資は各年度初月に一括計上。財務調達は固定月（M13/M25）に計上。",
   italic=True, fc=C_NOTE, size=9)
ws_cf.merge_cells(f"A2:{LAST_COL}2")
ws_cf.freeze_panes = "B5"  # A列ラベル固定 + 上4行固定

# Row 3: Year帯ラベル（12ヶ月ずつmerge）
sc(ws_cf, 3, 1, "", fc=C_BHDR)
for start, end, lbl, fc_yr in [
    (1,  12, "Year 1  （M1〜M12）",  C_NAVY),
    (13, 24, "Year 2  （M13〜M24）", C_BLUE),
    (25, 36, "Year 3  （M25〜M36）", C_NAVY),
]:
    sc(ws_cf, 3, start + 1, lbl, bold=True, fc=fc_yr, color="FFFFFF", align="center")
    ws_cf.merge_cells(f"{gl(start+1)}3:{gl(end+1)}3")

# Row 4: 月番号ラベル
sc(ws_cf, 4, 1, "項目 / 月", bold=True, fc=C_BHDR, color="FFFFFF")
for m in range(1, 37):
    bg = C_NAVY if (m - 1) // 12 % 2 == 0 else C_BLUE
    sc(ws_cf, 4, m + 1, f"M{m}", bold=True, fc=bg, color="FFFFFF", align="center")

def yr_col(m): return ["B","C","D"][(m - 1) // 12]  # 年次列 → 前提条件シート列

# ── P&Lサマリーセクションヘッダー ──
for row, txt in [
    (5,  "▌売上高"),
    (10, "▌売上原価・粗利"),
    (13, "▌SG&A・営業利益"),
]:
    sc(ws_cf, row, 1, txt, bold=True, fc=C_NAVY, color="FFFFFF")
    ws_cf.merge_cells(f"A{row}:{LAST_COL}{row}")

# P&Lサマリー月次行: (row, label, 前提条件行番号, bold, fc)
PL_CF = [
    (6,  "  B2B受託収益",         8,  False, None),
    (7,  "  人材紹介フィー",       11, False, None),
    (8,  "  プロジェクト収益",     14, False, None),
    (9,  "売上合計",               15, True,  C_TOT),
    (11, "  売上原価合計",         25, False, C_SUB),
    (12, "売上総利益",             26, True,  C_GPFT),
    (14, "  SG&A合計",            39, False, C_SUB),
    (15, "営業利益",               40, True,  C_TOT),
    (16, "当月純利益（税引後）",    67, True,  C_GPFT),
    (18, "当月償却費（参考・非現金）", 54, False, C_NOTE),
]

for row, lbl, mae_row, bold, fc in PL_CF:
    sc(ws_cf, row, 1, lbl, bold=bold, fc=fc)
    for m in range(1, 37):
        yc = yr_col(m)
        sc(ws_cf, row, m + 1,
           f=f"=ROUND(前提条件!{yc}{mae_row}/12,0)",
           bold=bold, fc=fc, nf=NF_MAN, align="center")

# ── 月次CFセクション ──
sc(ws_cf, 19, 1, "▌月次キャッシュフロー", bold=True, fc=C_NAVY, color="FFFFFF")
ws_cf.merge_cells(f"A19:{LAST_COL}19")

# ── Row 20: 期首現金残高（BSシートY0起点 → 前月期末を連鎖）──
sc(ws_cf, 20, 1, "期首現金残高")
for m in range(1, 37):
    col = m + 1
    if m == 1:
        f_str = "=BS（貸借対照表）!B6"     # Y0期首現金 = 資本金+借入金
    else:
        f_str = f"={gl(m)}26"               # 前月Row26（期末）を参照
    sc(ws_cf, 20, col, f=f_str, nf=NF_MAN, align="center")

# ── Row 21: 当月純利益 ──
sc(ws_cf, 21, 1, "当月純利益")
for m in range(1, 37):
    yc = yr_col(m)
    sc(ws_cf, 21, m + 1, f=f"=ROUND(前提条件!{yc}67/12,0)", nf=NF_MAN, align="center")

# ── Row 22: 当月償却費（非現金加算） ──
sc(ws_cf, 22, 1, "当月償却費（非現金加算）")
for m in range(1, 37):
    yc = yr_col(m)
    sc(ws_cf, 22, m + 1, f=f"=ROUND(前提条件!{yc}54/12,0)", nf=NF_MAN, align="center")

# ── Row 23: 設備投資（各年度初月 M1/M13/M25 に一括計上）──
sc(ws_cf, 23, 1, "設備投資（－）")
for m in range(1, 37):
    col = m + 1
    if m == 1:
        sc(ws_cf, 23, col, f="=前提条件!B52", fc=C_LOSS, nf=NF_MAN, align="center")
    elif m == 13:
        sc(ws_cf, 23, col, f="=前提条件!C52", fc=C_LOSS, nf=NF_MAN, align="center")
    elif m == 25:
        sc(ws_cf, 23, col, f="=前提条件!D52", fc=C_LOSS, nf=NF_MAN, align="center")
    else:
        sc(ws_cf, 23, col, v=0, nf=NF_MAN, align="center")

# ── Row 24: 運転資本変動（ΔAR − ΔAP、正 = 現金流出）──
# Y1: AR・APともゼロ起点から年次値まで均等増
# Y2/Y3: 年次差分を均等配分
sc(ws_cf, 24, 1, "運転資本変動（ΔAR−ΔAP）")
for m in range(1, 37):
    col = m + 1
    yr = (m - 1) // 12
    if yr == 0:
        f_str = (
            "=ROUND(前提条件!B47/12,0)"
            "-ROUND(前提条件!B48/12,0)"
        )
    elif yr == 1:
        f_str = (
            "=ROUND((前提条件!C47-前提条件!B47)/12,0)"
            "-ROUND((前提条件!C48-前提条件!B48)/12,0)"
        )
    else:
        f_str = (
            "=ROUND((前提条件!D47-前提条件!C47)/12,0)"
            "-ROUND((前提条件!D48-前提条件!C48)/12,0)"
        )
    sc(ws_cf, 24, col, f=f_str, nf=NF_MAN, align="center")

# ── Row 25: 財務調達（シード M13 / シリーズA M25）──
sc(ws_cf, 25, 1, "財務調達（調達・借入金）")
for m in range(1, 37):
    col = m + 1
    if m == 13:
        sc(ws_cf, 25, col, f="=前提条件!C60", fc=C_GPFT, nf=NF_MAN, align="center")
    elif m == 25:
        sc(ws_cf, 25, col, f="=前提条件!D61", fc=C_GPFT, nf=NF_MAN, align="center")
    else:
        sc(ws_cf, 25, col, v=0, nf=NF_MAN, align="center")

# ── Row 26: 期末現金残高 ──
# = 期首 + 純利益 + 償却費 − 設備投資 − 運転資本変動 + 財務調達
sc(ws_cf, 26, 1, "期末現金残高", bold=True, fc=C_TOT)
for m in range(1, 37):
    cl = gl(m + 1)
    sc(ws_cf, 26, m + 1,
       f=f"={cl}20+{cl}21+{cl}22-{cl}23-{cl}24+{cl}25",
       bold=True, fc=C_TOT, nf=NF_MAN, align="center")

bdr(ws_cf, 4, 26, 1, 37)
