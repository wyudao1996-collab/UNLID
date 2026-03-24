# ══════════════════════════════════════════════
# _v42_part2c.py: 感度分析シート + wb.save()
# Part1 + 2a + 2b の続き。wb 変数引き継ぎ前提。
# ══════════════════════════════════════════════

# ──── SHEET 6: 感度分析 ────
ws_sa = wb.create_sheet("感度分析")
ws_sa.column_dimensions["A"].width = 32
for c in ["B","C","D","E","F","G","H","I"]:
    ws_sa.column_dimensions[c].width = 15

title_row(ws_sa, 1, "UNLID 感度分析 v4.2  ─  ユニットエコノミクス × シナリオ比較", 9)
sc(ws_sa, 2, 1,
   "★ マトリックスは静的計算。シナリオサマリーは前提条件シート基準値に倍率を乗算。",
   italic=True, fc=C_NOTE, size=9)
mrg(ws_sa, 2, 1, 9)

MAE = "前提条件!"  # 前提条件シート参照プレフィックス

# ────────────────────────────────────────────
# TABLE 1: 月商マトリックス（FDE人数 × 担当社数）
# ────────────────────────────────────────────
sc(ws_sa, 4, 1, "▌月商マトリックス（万円/月）  ─  単価：20万円/社 固定", bold=True, fc=C_NAVY, color="FFFFFF")
mrg(ws_sa, 4, 1, 9)

sc(ws_sa, 5, 1, "FDE人数 ╲ 担当社数", bold=True, fc=C_BHDR, color="FFFFFF")
CLIENTS_LIST = [4, 5, 6, 7, 8]
for ci, clients in enumerate(CLIENTS_LIST):
    lbl = f"{clients}社 ★BEP" if clients == 6 else f"{clients}社"
    bg  = C_BLUE if clients == 6 else C_BHDR
    sc(ws_sa, 5, ci + 2, lbl, bold=True, fc=bg, color="FFFFFF", align="center")

for fi, fde in enumerate([1, 2, 3, 4, 5]):
    row = 6 + fi
    sc(ws_sa, row, 1, f"FDE {fde}名", bold=True)
    for ci, clients in enumerate(CLIENTS_LIST):
        col  = ci + 2
        rev  = fde * clients * 20          # 月商（万円）
        is_bep = (fde == 1 and clients == 6)
        bg = C_GPFT if is_bep else (C_LBLUE if clients == 6 else None)
        sc(ws_sa, row, col, v=rev, bold=(clients == 6), fc=bg, nf=NF_MAN, align="center")
bdr(ws_sa, 5, 10, 1, 6)

# ────────────────────────────────────────────
# TABLE 2: 粗利マトリックス（ワーカー費15% + FDE給与40万/名を控除）
# ────────────────────────────────────────────
sc(ws_sa, 12, 1, "▌粗利マトリックス（万円/月）  ─  ワーカー費15%控除 + FDE給与40万/名控除", bold=True, fc=C_NAVY, color="FFFFFF")
mrg(ws_sa, 12, 1, 9)

sc(ws_sa, 13, 1, "FDE人数 ╲ 担当社数", bold=True, fc=C_BHDR, color="FFFFFF")
for ci, clients in enumerate(CLIENTS_LIST):
    lbl = f"{clients}社 ★BEP" if clients == 6 else f"{clients}社"
    bg  = C_BLUE if clients == 6 else C_BHDR
    sc(ws_sa, 13, ci + 2, lbl, bold=True, fc=bg, color="FFFFFF", align="center")

for fi, fde in enumerate([1, 2, 3, 4, 5]):
    row = 14 + fi
    sc(ws_sa, row, 1, f"FDE {fde}名", bold=True)
    for ci, clients in enumerate(CLIENTS_LIST):
        col   = ci + 2
        gross = round(fde * clients * 20 * 0.85 - fde * 40)  # 粗利（万円/月）
        is_bep = (fde == 1 and clients == 6)
        bg = C_GPFT if is_bep else (C_LBLUE if clients == 6 else (C_LOSS if gross < 0 else None))
        sc(ws_sa, row, col, v=gross, bold=(clients == 6), fc=bg, nf=NF_MAN, align="center")
bdr(ws_sa, 13, 18, 1, 6)

# ────────────────────────────────────────────
# TABLE 3: B2B単価感度（Year1）
# ────────────────────────────────────────────
sc(ws_sa, 20, 1, "▌B2B単価感度（Year1）  ─  FDE×社数構成不変、月額単価のみ変動", bold=True, fc=C_NAVY, color="FFFFFF")
mrg(ws_sa, 20, 1, 9)

col_hdrs(ws_sa, 21,
    ["月額単価", "Y1 売上（万円）", "Y1 粗利（万円）", "Y1 営業利益（万円）", "Base比変化率"],
    fc_def=C_BHDR)

for pi, price in enumerate([15, 17, 20, 23, 25]):
    row   = 22 + pi
    ratio = price / 20                     # Base比（20万基準）
    is_base = (price == 20)
    bg = C_LBLUE if is_base else (C_GPFT if price > 20 else C_LOSS)
    lbl = f"{price}万円/社{'  ★Base' if is_base else ''}"
    sc(ws_sa, row, 1, lbl, bold=is_base, fc=bg)
    sc(ws_sa, row, 2, f"=ROUND({MAE}B15*{ratio},0)", bold=is_base, fc=bg, nf=NF_MAN, align="center")
    sc(ws_sa, row, 3, f"=ROUND({MAE}B26*{ratio},0)", bold=is_base, fc=bg, nf=NF_MAN, align="center")
    sc(ws_sa, row, 4, f"=ROUND({MAE}B40*{ratio},0)", bold=is_base, fc=bg, nf=NF_MAN, align="center")
    sc(ws_sa, row, 5, v=round(ratio - 1, 3),        bold=is_base, fc=bg, nf=NF_PCT, align="center")
bdr(ws_sa, 21, 26, 1, 5)

# ────────────────────────────────────────────
# TABLE 4: 3シナリオサマリー（Bear / Base / Bull）
# ────────────────────────────────────────────
sc(ws_sa, 28, 1, "▌3シナリオサマリー（前提条件基準値 × 倍率）", bold=True, fc=C_NAVY, color="FFFFFF")
mrg(ws_sa, 28, 1, 9)

col_hdrs(ws_sa, 29,
    ["シナリオ", "Y1 売上", "Y1 粗利", "Y1 営業利益", "Y2 売上", "Y2 粗利", "Y2 営業利益", "Y3 売上", "Y3 営業利益"],
    fc_def=C_BHDR)

for si, (lbl, mult, fc) in enumerate([
    ("Bear（悲観）  ×0.70", 0.70, C_LOSS),
    ("Base（中央）  ×1.00", 1.00, C_LBLUE),
    ("Bull（楽観）  ×1.30", 1.30, C_GPFT),
]):
    row = 30 + si
    sc(ws_sa, row, 1, lbl, bold=(mult == 1.0), fc=fc)
    for c, fml in enumerate([
        f"=ROUND({MAE}B15*{mult},0)", f"=ROUND({MAE}B26*{mult},0)", f"=ROUND({MAE}B40*{mult},0)",
        f"=ROUND({MAE}C15*{mult},0)", f"=ROUND({MAE}C26*{mult},0)", f"=ROUND({MAE}C40*{mult},0)",
        f"=ROUND({MAE}D15*{mult},0)",                                f"=ROUND({MAE}D40*{mult},0)",
    ], start=2):
        sc(ws_sa, row, c, f=fml, bold=(mult == 1.0), fc=fc, nf=NF_MAN, align="center")
bdr(ws_sa, 29, 32, 1, 9)

# ════════════════════════════════════════════
# 保存
# ════════════════════════════════════════════
OUTPUT_FILE = "UNLID_財務モデル_v4.2.xlsx"
wb.save(OUTPUT_FILE)
print(f"✅  生成完了 → {OUTPUT_FILE}")
print(f"    シート一覧: {[s.title for s in wb.worksheets]}")
