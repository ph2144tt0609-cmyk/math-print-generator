# -*- coding: utf-8 -*-
"""小学2年生 算数100問プリント PDF生成"""
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# ===== フォント登録 =====
pdfmetrics.registerFont(TTFont("NotoJP", r"C:\Windows\Fonts\NotoSansJP-VF.ttf"))

# ===== スタイル =====
title_style = ParagraphStyle(
    "title", fontName="NotoJP", fontSize=18, leading=24,
    alignment=TA_CENTER, textColor=colors.HexColor("#1F4E79"),
)
subtitle_style = ParagraphStyle(
    "subtitle", fontName="NotoJP", fontSize=10, leading=14,
    alignment=TA_CENTER, textColor=colors.HexColor("#555555"),
)
section_style = ParagraphStyle(
    "section", fontName="NotoJP", fontSize=13, leading=18,
    alignment=TA_LEFT, textColor=colors.white,
    backColor=colors.HexColor("#1F4E79"),
    leftIndent=4, rightIndent=4,
    spaceBefore=4, spaceAfter=4,
    borderPadding=(3, 4, 3, 4),
)
q_calc_style = ParagraphStyle(  # 計算問題用（大きめ）
    "qcalc", fontName="NotoJP", fontSize=14, leading=22,
    alignment=TA_LEFT,
)
q_word_style = ParagraphStyle(  # 文章題用
    "qword", fontName="NotoJP", fontSize=12, leading=20,
    alignment=TA_LEFT,
)
ans_style = ParagraphStyle(  # 答え欄ラベル
    "ans", fontName="NotoJP", fontSize=11, leading=16,
    alignment=TA_LEFT, textColor=colors.HexColor("#888888"),
)
name_style = ParagraphStyle(
    "name", fontName="NotoJP", fontSize=11, leading=16,
    alignment=TA_LEFT,
)

# ===== 問題データ =====
sec1 = [  # たし算・ひき算 1-25
    "27＋15＝", "48＋32＝", "63－28＝", "91－45＝", "36＋27＝",
    "74－19＝", "58＋14＝", "80－36＝", "29＋49＝", "100－57＝",
    "18＋26＝", "35＋47＝", "92－37＝", "65－18＝", "44＋39＝",
    "71－24＝", "53＋28＝", "88－49＝", "67＋16＝", "90－34＝",
    "25＋58＝", "73－46＝", "39＋42＝", "81－29＝", "56＋35＝",
]
sec2 = [  # 筆算 26-40
    "48＋37", "56＋29", "67＋28", "84＋19", "73＋38",
    "92－47", "81－35", "74－28", "63－39", "100－48",
    "58＋27－16", "42＋39－25", "90－27＋18", "63－19＋25", "55＋28－34",
]
sec3 = [  # かけ算準備 41-50（文章短め）
    "5が4こあります。ぜんぶでいくつ？",
    "3こ入りのみかんのふくろが6つあります。みかんは何こ？",
    "2まいずつクッキーをならべます。8人分では何まい？",
    "10円玉が7まいあります。何円？",
    "4人のチームが3つあります。みんなで何人？",
    "6このイスに1人ずつすわります。何人？",
    "5本入りのえんぴつを2セット買いました。何本？",
    "3本の花が7列あります。花は何本？",
    "2こずつボールがあります。9組では何こ？",
    "4だんのたなに3さつずつ本があります。何さつ？",
]
sec4 = [  # わくわく文章題 51-70
    "ミナトくんはシールを28まい持っています。おじいちゃんから15まいもらいました。何まいになりましたか。",
    "公園に45人いました。13人帰りました。何人いますか。",
    "りんごが18こあります。7こ食べました。のこりは？",
    "電車に34人乗っています。次の駅で25人乗りました。何人になりましたか。",
    "魚を12ひきつりました。お父さんが8ひきつりました。あわせて何ひき？",
    "どんぐりを39こ集めました。10こ友だちにあげました。のこりは？",
    "野球カードを25まい持っています。18まい買いました。何まい？",
    "アイスを40本用意しました。27本売れました。のこりは？",
    "ひまわりが23本あります。さらに19本植えました。何本？",
    "金魚が17ひきいました。5ひき増えました。何ひき？",
    "かき氷を32こ作りました。14こ売れました。のこりは？",
    "サッカーボールが48こあります。9こ使っています。使っていないボールは？",
    "クラスで35人います。2人休みです。何人来ていますか。",
    "本を60ページ読みました。あと25ページあります。本は全部で何ページ？",
    "おこづかい100円のうち36円使いました。のこりはいくら？",
    "チョコレートを29こ作りました。11こ食べました。のこりは？",
    "カブトムシを13ひき見つけました。友だちが9ひき見つけました。あわせて？",
    "うさぎが22わいます。7わ増えました。何わ？",
    "公園のベンチに15人座っています。あと8人来ました。何人？",
    "パンを50こ焼きました。24こ売れました。のこりは？",
]
sec5 = [  # お金 71-80
    "50円＋20円＝",
    "100円－30円＝",
    "10円玉が8まいあります。何円？",
    "50円玉2まいで何円？",
    "120円のおかしを買いました。200円出しました。おつりは？",
    "70円のジュースと30円のガムを買いました。いくら？",
    "300円持っています。120円使いました。のこりは？",
    "50円玉3まいと10円玉2まいで何円？",
    "90円のアイスを2こ買いました。いくら？",
    "250円の本を買うのに300円出しました。おつりは？",
]
sec6 = [  # 時こく・カレンダー 81-90
    "3時の1時間後は？",
    "7時の2時間後は？",
    "10時の3時間前は？",
    "8時30分の30分後は？",
    "1時15分の45分後は？",
    "4時から6時まで何時間？",
    "9時から11時まで何時間？",
    "月曜日の次の日は？",
    "金曜日の2日後は？",
    "1週間は何日？",
]
sec7 = [  # なぞとき 91-100
    "ねこが4ひきいます。足は全部で何本？",
    "いぬが3びきいます。足は全部で何本？",
    "三輪車が5台あります。タイヤは全部で何こ？",
    "クモが2ひきいます。足は全部で何本？",
    "たこ焼きが24こあります。6人で同じ数ずつ食べると1人何こ？",
    "ピザを8つに切りました。4人で同じ数ずつ食べると何切れ？",
    "□＋15＝40　　□はいくつ？",
    "60－□＝25　　□はいくつ？",
    "ある数に20を足すと55になりました。ある数は？",
    "ミナトくんは宝箱を見つけました。中に金貨が32まいあります。ドラゴンをたおしたら18まい増えました。さらに10まい使いました。今何まいありますか？",
]

# ===== 描画ヘルパ =====
PAGE_W, PAGE_H = A4
LEFT = RIGHT = 14 * mm
TOP = 14 * mm
BOTTOM = 14 * mm
USABLE_W = PAGE_W - LEFT - RIGHT

def section_header(no_range, title):
    return Paragraph(f"【{no_range}】{title}", section_style)

def calc_cell(idx, q):
    """計算問題（短い）— 番号＋式＋答え欄の下線"""
    # 番号は丸囲い風に番号＋式、その下にうすい下線
    p = Paragraph(f"<b>{idx}.</b>　{q}　<font color='#BBBBBB'>＿＿＿＿</font>", q_calc_style)
    return p

def hissan_cell(idx, q):
    """筆算用 — 数字を大きくして書き込みスペースを下に確保"""
    # 1つのTableに式の行＋書き込み用空白行を入れる（colWidthsは外側カラムに収まる値）
    t = Table(
        [[Paragraph(f"<b>{idx}.</b>　<font size='15'>{q}</font>", q_calc_style)],
         [""]],
        colWidths=[48*mm],
        rowHeights=[8*mm, 24*mm],
    )
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOX", (0, 1), (-1, 1), 0.5, colors.HexColor("#CCCCCC")),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FAFAFA")),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t

def word_cell(idx, q, height_mm=18):
    """文章題 — 問題文の下に書き込み欄（単一Paragraphで返す）"""
    txt = (
        f"<b>{idx}.</b>　{q}<br/>"
        f"<br/>"
        f"<font color='#888888'>こたえ：</font>"
        f"<font color='#BBBBBB'>＿＿＿＿＿＿＿＿＿＿＿＿</font>"
    )
    return Paragraph(txt, q_word_style)

def short_word_cell(idx, q):
    """短い文章題（お金・時こく・かけ算準備）"""
    txt = (
        f"<b>{idx}.</b>　{q}<br/>"
        f"<font color='#888888'>こたえ：</font>"
        f"<font color='#BBBBBB'>＿＿＿＿＿＿＿＿＿＿</font>"
    )
    return Paragraph(txt, q_word_style)

def two_col_table(cells):
    """2列で並べる。cellsは奇数個でもOK"""
    rows = []
    col_w = (USABLE_W - 6*mm) / 2
    for i in range(0, len(cells), 2):
        left = cells[i]
        right = cells[i+1] if i+1 < len(cells) else ""
        rows.append([left, right])
    t = Table(rows, colWidths=[col_w, col_w])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t

def three_col_table(cells):
    """3列（筆算用）"""
    rows = []
    col_w = (USABLE_W - 4*mm) / 3
    for i in range(0, len(cells), 3):
        row = [cells[j] if j < len(cells) else "" for j in (i, i+1, i+2)]
        rows.append(row)
    t = Table(rows, colWidths=[col_w]*3)
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return t

# ===== ヘッダー・フッター =====
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont("NotoJP", 9)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawRightString(PAGE_W - RIGHT, 8*mm, f"- {doc.page} -")
    canvas.drawString(LEFT, 8*mm, "小学2年生 算数100問")
    canvas.restoreState()

# ===== ドキュメント組み立て =====
OUT = r"H:\マイドライブ\20_claude code\算数100問プリント.pdf"
doc = SimpleDocTemplate(
    OUT, pagesize=A4,
    leftMargin=LEFT, rightMargin=RIGHT,
    topMargin=TOP, bottomMargin=BOTTOM,
    title="小学2年生 算数100問プリント",
)

story = []

# 表紙的なタイトル
story.append(Paragraph("小学2年生　算数100問チャレンジ", title_style))
story.append(Spacer(1, 2*mm))
story.append(Paragraph("計算力・考える力・文章題・図形感覚・お金・時こく", subtitle_style))
story.append(Spacer(1, 4*mm))

# なまえ欄
name_table = Table(
    [[Paragraph("なまえ：", name_style),
      "",
      Paragraph("日づけ：", name_style),
      ""]],
    colWidths=[18*mm, 70*mm, 18*mm, 50*mm],
)
name_table.setStyle(TableStyle([
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("LINEBELOW", (1, 0), (1, 0), 0.7, colors.black),
    ("LINEBELOW", (3, 0), (3, 0), 0.7, colors.black),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
]))
story.append(name_table)
story.append(Spacer(1, 5*mm))

# セクション1：たし算・ひき算
story.append(section_header("1〜25", "たし算・ひき算"))
story.append(Spacer(1, 2*mm))
cells = [calc_cell(i+1, q) for i, q in enumerate(sec1)]
story.append(two_col_table(cells))
story.append(PageBreak())

# セクション2：筆算
story.append(section_header("26〜40", "筆算チャレンジ"))
story.append(Spacer(1, 2*mm))
cells = [hissan_cell(26+i, q) for i, q in enumerate(sec2)]
story.append(three_col_table(cells))
story.append(PageBreak())

# セクション3：かけ算準備
story.append(section_header("41〜50", "かけ算じゅんび"))
story.append(Spacer(1, 2*mm))
cells = [short_word_cell(41+i, q) for i, q in enumerate(sec3)]
story.append(two_col_table(cells))
story.append(Spacer(1, 4*mm))

# セクション5（お金）も同じページに入るなら入れたいが、まずは順番通り
# セクション4：わくわく文章題（20問、長め）
story.append(section_header("51〜70", "わくわく文章題"))
story.append(Spacer(1, 2*mm))
cells = [word_cell(51+i, q) for i, q in enumerate(sec4)]
story.append(two_col_table(cells))
story.append(PageBreak())

# セクション5：お金
story.append(section_header("71〜80", "お金マスター"))
story.append(Spacer(1, 2*mm))
cells = [short_word_cell(71+i, q) for i, q in enumerate(sec5)]
story.append(two_col_table(cells))
story.append(Spacer(1, 4*mm))

# セクション6：時こく
story.append(section_header("81〜90", "時こくとカレンダー"))
story.append(Spacer(1, 2*mm))
cells = [short_word_cell(81+i, q) for i, q in enumerate(sec6)]
story.append(two_col_table(cells))
story.append(PageBreak())

# セクション7：なぞとき
story.append(section_header("91〜100", "なぞとき算数"))
story.append(Spacer(1, 2*mm))
cells = [word_cell(91+i, q) for i, q in enumerate(sec7)]
story.append(two_col_table(cells))

# ボーナス
story.append(Spacer(1, 5*mm))
story.append(section_header("ボーナス", "算数好きになるチャレンジ"))
story.append(Spacer(1, 2*mm))
bonus = [
    "★ 100円で買えるお菓子を3つ考えよう",
    "★ 家の中にある四角を10こ探そう",
    "★ スーパーで1000円以内のお買い物ゲームをしよう",
    "★ 野球の点数を使って算数の問題を自分で作ろう",
    "★ 今日1日で見つけた数字を20こ書き出そう",
]
for b in bonus:
    story.append(Paragraph(b, q_word_style))
    story.append(Spacer(1, 2))

doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f"OK: {OUT}")
