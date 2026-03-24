# ══════════════════════════════════════════════
# _v42_part2a.py: PL（損益計算書）+ BS（貸借対照表）
# Part1(_v42_part1.py)の続き。wb変数を引き継ぐ前提。
# ══════════════════════════════════════════════

# ──── SHEET 3: PL（損益計算書）────
ws_pl = wb.create_sheet("PL（損益計算書）")
ws_pl.column_dimensions["A"].width = 32
for c in ["B","C","D","E"]:
    ws_pl.column_dimensions[c].width = 15

title_row(ws_pl, 1, "UNLID 損益計算書（P/L）v4.2  ─  単位：万円  |  前提条件シート連動", 5)
sc(ws_pl, 2, 1,
   "★ 人件費・採用費は人員計画シートから自動参照。全数値は前提条件シートのドライバーで算出。",
   italic=True, fc=C_NOTE, size=9)
mrg(ws_pl, 2, 1, 5)
col_hdrs(ws_pl, 4, ["項目", "Year1", "Year2", "Year3", "3年累計"], fc_def=C_BHDR)

def p(r, y):
    """前提条件シートのセル参照（r=行番号, y=1/2/3→B/C/D列）"""
    return f"前提条件!{['B','C','D'][y-1]}{r}"

# セクションヘッダー（merge行）
for row, txt in [
    (5,  "▌売上高"),
    (10, "▌売上原価（COGS）"),
    (17, "▌SG&A（販売費及び一般管理費）"),
    (33, "▌参考指標"),
]:
    sc(ws_pl, row, 1, txt, bold=True, fc=C_NAVY, color="FFFFFF")
    mrg(ws_pl, row, 1, 5)

# (行, ラベル, Year1式, Year2式, Year3式, bold, fc)
PL_DATA = [
    # 売上高
    (6,  "  B2B受託収益",         f"={p(8,1)}",  f"={p(8,2)}",  f"={p(8,3)}",  False, None),
    (7,  "  人材紹介フィー",       f"={p(11,1)}", f"={p(11,2)}", f"={p(11,3)}", False, None),
    (8,  "  プロジェクト収益",     f"={p(14,1)}", f"={p(14,2)}", f"={p(14,3)}", False, None),
    (9,  "売上合計",               f"={p(15,1)}", f"={p(15,2)}", f"={p(15,3)}", True,  C_TOT),
    # COGS
    (11, "  AIサーバー・API費",    f"={p(20,1)}+{p(21,1)}", f"={p(20,2)}+{p(21,2)}", f"={p(20,3)}+{p(21,3)}", False, None),
    (12, "  副業ワーカー報酬",     f"={p(23,1)}", f"={p(23,2)}", f"={p(23,3)}", False, None),
    (13, "  外注エンジニア費",     f"={p(24,1)}", f"={p(24,2)}", f"={p(24,3)}", False, None),
    (14, "売上原価合計",           f"={p(25,1)}", f"={p(25,2)}", f"={p(25,3)}", True,  C_SUB),
    (15, "売上総利益",             f"={p(26,1)}", f"={p(26,2)}", f"={p(26,3)}", True,  C_GPFT),
    (16, "  粗利率",               f"={p(27,1)}", f"={p(27,2)}", f"={p(27,3)}", False, None),
    # SG&A
    (18, "  人件費合計",           "=人員計画!D32", "=人員計画!E32", "=人員計画!F32", False, None),
    (19, "  採用費合計",           "=人員計画!D42", "=人員計画!E42", "=人員計画!F42", False, None),
    (20, "  B2Bマーケ・CAC費",    f"={p(33,1)}", f"={p(33,2)}", f"={p(33,3)}", False, None),
    (21, "  人材紹介広告費",       f"={p(34,1)}", f"={p(34,2)}", f"={p(34,3)}", False, None),
    (22, "  地代家賃・オフィス",   f"={p(35,1)}", f"={p(35,2)}", f"={p(35,3)}", False, None),
    (23, "  SaaS・システム費",     f"={p(36,1)}", f"={p(36,2)}", f"={p(36,3)}", False, None),
    (24, "  法務・許認可費",       f"={p(37,1)}", f"={p(37,2)}", f"={p(37,3)}", False, None),
    (25, "  その他固定費",         f"={p(38,1)}", f"={p(38,2)}", f"={p(38,3)}", False, None),
    (26, "SG&A合計",              f"={p(39,1)}", f"={p(39,2)}", f"={p(39,3)}", True,  C_SUB),
    # 利益
    (27, "営業利益",               f"={p(40,1)}", f"={p(40,2)}", f"={p(40,3)}", True,  C_TOT),
    (28, "  営業利益率",           f"={p(41,1)}", f"={p(41,2)}", f"={p(41,3)}", False, None),
    (29, "  支払利息",             f"={p(64,1)}", f"={p(64,2)}", f"={p(64,3)}", False, C_LOSS),
    (30, "税引前利益",             f"={p(65,1)}", f"={p(65,2)}", f"={p(65,3)}", True,  C_TOT),
    (31, "  法人税等（30%）",      f"={p(66,1)}", f"={p(66,2)}", f"={p(66,3)}", False, None),
    (32, "当期純利益",             f"={p(67,1)}", f"={p(67,2)}", f"={p(67,3)}", True,  C_GPFT),
    # 参考
    (34, "  減価償却費（参考）",   f"={p(54,1)}", f"={p(54,2)}", f"={p(54,3)}", False, C_NOTE),
    (35, "  EBITDA（参考）",       f"={p(40,1)}+{p(54,1)}", f"={p(40,2)}+{p(54,2)}", f"={p(40,3)}+{p(54,3)}", True, C_GPFT),
]

for row, lbl, f1, f2, f3, bold, fc in PL_DATA:
    sc(ws_pl, row, 1, lbl, bold=bold, fc=fc)
    nf = NF_PCT if "率" in lbl else NF_MAN
    for c, fml in [(2, f1), (3, f2), (4, f3)]:
        sc(ws_pl, row, c, f=fml, bold=bold, fc=fc, nf=nf, align="center")
    if "率" not in lbl:
        sc(ws_pl, row, 5, f=f"=B{row}+C{row}+D{row}", bold=bold, fc=fc, nf=NF_MAN, align="center")

bdr(ws_pl, 4, 35, 1, 5)

# ──── SHEET 4: BS（貸借対照表）────
ws_bs = wb.create_sheet("BS（貸借対照表）")
ws_bs.column_dimensions["A"].width = 30
for c in ["B", "C", "D", "E"]:
    ws_bs.column_dimensions[c].width = 16

title_row(ws_bs, 1, "UNLID 貸借対照表（B/S）v4.2  ─  単位：万円  |  期末残高ベース", 5)
sc(ws_bs, 2, 1,
   "注：現金残高は「営業CF + 投資CF + 財務CF」の累積。貸借差額セル（Row28）で一致確認。",
   italic=True, fc=C_NOTE, size=9)
mrg(ws_bs, 2, 1, 5)
col_hdrs(ws_bs, 4, ["項目", "Y0（期首）", "Y1末", "Y2末", "Y3末"], fc_def=C_BHDR)

Q = "前提条件!"  # 前提条件シート参照プレフィックス

# セクションヘッダー
for row, txt in [
    (5,  "▌流動資産"),
    (9,  "▌固定資産"),
    (14, "▌流動負債"),
    (18, "▌固定負債"),
    (23, "▌純資産"),
]:
    sc(ws_bs, row, 1, txt, bold=True, fc=C_NAVY, color="FFFFFF")
    mrg(ws_bs, row, 1, 5)

# 現金・預金（累積CF計算）
# Y0 = 資本金払込 + 借入金
# Y1 = Y0 + 純利益 + 償却費 - ソフト投資 - 売掛金 + 買掛金
# Y2 = Y1 + 上記変動分 + シード調達
# Y3 = Y2 + 上記変動分 + シリーズA
sc(ws_bs, 6, 1, "現金・預金", bold=True)
sc(ws_bs, 6, 2, f"={Q}B59+{Q}B62",
   bold=True, nf=NF_MAN, align="center")
sc(ws_bs, 6, 3,
   f"=B6+{Q}B67+{Q}B54-{Q}B52-{Q}B47+{Q}B48",
   bold=True, nf=NF_MAN, align="center")
sc(ws_bs, 6, 4,
   f"=C6+{Q}C67+{Q}C54-{Q}C52-({Q}C47-{Q}B47)+({Q}C48-{Q}B48)+{Q}C60",
   bold=True, nf=NF_MAN, align="center")
sc(ws_bs, 6, 5,
   f"=D6+{Q}D67+{Q}D54-{Q}D52-({Q}D47-{Q}C47)+({Q}D48-{Q}C48)+{Q}D61",
   bold=True, nf=NF_MAN, align="center")

# 売掛金（前提条件シート参照）
sc(ws_bs, 7, 1, "売掛金")
sc(ws_bs, 7, 2, v=0, nf=NF_MAN, align="center")
for c, col in [(3, "B"), (4, "C"), (5, "D")]:
    sc(ws_bs, 7, c, f=f"={Q}{col}47", nf=NF_MAN, align="center")

# 流動資産合計
sc(ws_bs, 8, 1, "流動資産合計", bold=True, fc=C_SUB)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 8, c, f=f"={gl(c)}6+{gl(c)}7", bold=True, fc=C_SUB, nf=NF_MAN, align="center")

# ソフトウェア（純額）
sc(ws_bs, 10, 1, "ソフトウェア（純額）")
sc(ws_bs, 10, 2, v=0, nf=NF_MAN, align="center")
for c, col in [(3, "B"), (4, "C"), (5, "D")]:
    sc(ws_bs, 10, c, f=f"={Q}{col}55", nf=NF_MAN, align="center")

# 固定資産合計
sc(ws_bs, 11, 1, "固定資産合計", bold=True, fc=C_SUB)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 11, c, f=f"={gl(c)}10", bold=True, fc=C_SUB, nf=NF_MAN, align="center")

# 資産合計
sc(ws_bs, 12, 1, "資産合計", bold=True, fc=C_TOT)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 12, c, f=f"={gl(c)}8+{gl(c)}11", bold=True, fc=C_TOT, nf=NF_MAN, align="center")
bdr(ws_bs, 4, 12, 1, 5)

# 買掛金
sc(ws_bs, 15, 1, "買掛金")
sc(ws_bs, 15, 2, v=0, nf=NF_MAN, align="center")
for c, col in [(3, "B"), (4, "C"), (5, "D")]:
    sc(ws_bs, 15, c, f=f"={Q}{col}48", nf=NF_MAN, align="center")

# 未払費用（人件費1ヶ月相当 → 人員計画シート参照）
sc(ws_bs, 16, 1, "未払費用（人件費1ヶ月相当）")
sc(ws_bs, 16, 2, v=0, nf=NF_MAN, align="center")
for c, col in [(3, "D"), (4, "E"), (5, "F")]:
    sc(ws_bs, 16, c, f=f"=ROUND(人員計画!{col}32/12,0)", nf=NF_MAN, align="center")

# 流動負債合計
sc(ws_bs, 17, 1, "流動負債合計", bold=True, fc=C_SUB)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 17, c, f=f"={gl(c)}15+{gl(c)}16", bold=True, fc=C_SUB, nf=NF_MAN, align="center")

# 借入金（返済なし想定）
sc(ws_bs, 19, 1, "借入金（期中返済なし想定）")
for c in [2, 3, 4, 5]:
    sc(ws_bs, 19, c, f=f"={Q}B62", nf=NF_MAN, align="center")

# 固定負債合計
sc(ws_bs, 20, 1, "固定負債合計", bold=True, fc=C_SUB)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 20, c, f=f"={gl(c)}19", bold=True, fc=C_SUB, nf=NF_MAN, align="center")

# 負債合計
sc(ws_bs, 21, 1, "負債合計", bold=True, fc=C_TOT)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 21, c, f=f"={gl(c)}17+{gl(c)}20", bold=True, fc=C_TOT, nf=NF_MAN, align="center")
bdr(ws_bs, 14, 21, 1, 5)

# 払込資本（資本金＋調達額累計）
sc(ws_bs, 24, 1, "払込資本（資本金＋調達額累計）")
sc(ws_bs, 24, 2, f"={Q}B59", nf=NF_MAN, align="center")          # Y0: 資本金のみ
sc(ws_bs, 24, 3, f"={Q}B59", nf=NF_MAN, align="center")          # Y1: 新規調達なし
sc(ws_bs, 24, 4, f"={Q}B59+{Q}C60", nf=NF_MAN, align="center")  # Y2: +シード3,000万
sc(ws_bs, 24, 5, f"={Q}B59+{Q}C60+{Q}D61", nf=NF_MAN, align="center")  # Y3: +シリーズA10,000万

# 利益剰余金（累積純利益）
sc(ws_bs, 25, 1, "利益剰余金（累積純利益）")
sc(ws_bs, 25, 2, v=0, nf=NF_MAN, align="center")
sc(ws_bs, 25, 3, f"={Q}B67", nf=NF_MAN, align="center")
sc(ws_bs, 25, 4, f"={Q}B67+{Q}C67", nf=NF_MAN, align="center")
sc(ws_bs, 25, 5, f"={Q}B67+{Q}C67+{Q}D67", nf=NF_MAN, align="center")

# 純資産合計
sc(ws_bs, 26, 1, "純資産合計", bold=True, fc=C_GPFT)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 26, c, f=f"={gl(c)}24+{gl(c)}25", bold=True, fc=C_GPFT, nf=NF_MAN, align="center")

# 検証：負債・純資産合計（資産合計と一致すれば貸借OK）
sc(ws_bs, 27, 1, "負債・純資産合計（検証）", bold=True, fc=C_TOT)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 27, c, f=f"={gl(c)}21+{gl(c)}26", bold=True, fc=C_TOT, nf=NF_MAN, align="center")

sc(ws_bs, 28, 1, "貸借差額（0 = 正常・要確認）", bold=True)
for c in [2, 3, 4, 5]:
    sc(ws_bs, 28, c, f=f"={gl(c)}12-{gl(c)}27", bold=True,
       fc=C_GPFT, nf=NF_MAN, align="center")

bdr(ws_bs, 23, 28, 1, 5)
