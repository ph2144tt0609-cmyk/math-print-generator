# -*- coding: utf-8 -*-
"""資産管理台帳 分析レポート PDF 生成（A4・6ページ・日本語）

P1: ダッシュボード（KPI／金融資産内訳／BS／不動産×ローン／前月からの資産の変化）
P2: 投資の見える化（NISA生涯枠の進捗／企業型DCの積み上げ・年7%／金融資産の歩み）
P3: 不動産投資と老後の設計（住宅＋投資ローン／サブリース／老後収入の柱・夫婦の私的年金）
P4: 強み・弱点・アクション／提出資料チェックリスト
P5: 保険・保障の総点検（保障マップ／加入保険一覧／名義変更スキーム／保険の弱点・残る穴）
P6: 世帯年収（実収入）の推移（旅費_世帯年収まとめ.xlsx を参照し年次推移を自動表示）
"""
import os
import csv as _csv
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF

pdfmetrics.registerFont(TTFont("JP",  r"C:\Windows\Fonts\meiryo.ttc",  subfontIndex=0))
pdfmetrics.registerFont(TTFont("JPB", r"C:\Windows\Fonts\meiryob.ttc", subfontIndex=0))

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "資産管理台帳_分析レポート_20260708_v30.pdf")

W, H = A4                      # 595.27 x 841.89
ML, MR = 36, 36                # 左右マージン
CW = W - ML - MR               # コンテンツ幅 523

NAVY   = HexColor("#1B2A4A")
TEAL   = HexColor("#2A9D8F")
BLUE   = HexColor("#4A6FA5")
GOLD   = HexColor("#E9C46A")
DARK   = HexColor("#264653")
ROSE   = HexColor("#B5838D")
GRAY   = HexColor("#9AA0A6")
LGRAY  = HexColor("#F2F4F7")
MGRAY  = HexColor("#DDE2E8")
RED    = HexColor("#C0392B")
REDBG  = HexColor("#FBEAE8")
AMBER  = HexColor("#B07D2B")
AMBERBG = HexColor("#FBF3E2")
PURPLE = HexColor("#7B6CA8")
TXT    = HexColor("#202124")
SUB    = HexColor("#5F6368")
WHITE  = HexColor("#FFFFFF")
TEALBG = HexColor("#E8F4F2")

# ============================================================ データ
# ---- 金融資産（台帳・2026-06-12）----
TOTAL_ASSET   = 151_966_716
FIN_TOTAL     = 41_166_716
RE_TOTAL      = 110_800_000

# ---- NISA実績（亨・SBI証券 2026/06/12時点の画面より）----
NISA_LIFE_CAP    = 18_000_000   # 生涯投資枠
NISA_GROWTH_CAP  = 12_000_000   # うち成長投資枠の上限
NISA_T_USED      = 3_692_848    # 生涯投資枠 利用額（簿価）
NISA_T_GROWTH    = 3_152_865    # うち成長投資枠
NISA_T_TSUMI     = NISA_T_USED - NISA_T_GROWTH   # うちつみたて投資枠 539,983
NISA_T_VALUE     = 4_545_161    # 評価残高
NISA_T_PL        = 769_733      # 評価損益 +20.38%
Y26_GROWTH_USED  = 1_973_606    # 2026年 成長投資枠 利用済み（/240万）
Y26_TSUMI_USED   = 40_000       # 2026年 つみたて投資枠 利用済み（/120万）

# ---- 美香さんのNISA（SBI証券 2026/06/13時点の画面より）----
NISA_M_USED      = 539_985      # 生涯投資枠 利用額（全額つみたて枠）
NISA_M_GROWTH    = 0            # 成長投資枠は未使用
NISA_M_TSUMI     = 539_985      # つみたて投資枠
NISA_M_VALUE     = 817_293      # 評価残高
NISA_M_PL        = 277_286      # 評価損益 +51.34%
NISA_COUPLE_USED = NISA_T_USED + NISA_M_USED   # 夫婦合計 4,232,833

# ---- 成長投資枠を「優先して埋める」方針（ユーザー指示・2026-06-16）----
# 商品の幅が広く枠が限られる成長投資枠（1人1,200万円）を先に満たし、残り600万円を
# つみたて枠で埋める。年間上限は成長240万円／つみたて120万円。
NISA_GROWTH_ANNUAL  = 2_400_000                       # 成長投資枠の年間上限（1人）
NISA_T_GROWTH_LEFT  = NISA_GROWTH_CAP - NISA_T_GROWTH  # 亨 成長枠の残り 8,847,135
NISA_M_GROWTH_LEFT  = NISA_GROWTH_CAP - NISA_M_GROWTH  # 美香 成長枠の残り 12,000,000
NISA_COUPLE_GROWTH      = NISA_T_GROWTH + NISA_M_GROWTH            # 夫婦 成長枠 利用 3,152,865
NISA_COUPLE_GROWTH_LEFT = NISA_T_GROWTH_LEFT + NISA_M_GROWTH_LEFT  # 夫婦 成長枠 残り 20,847,135

# ---- 企業型DC（SBIベネフィット・2026/06/12時点・夫婦合算）----
DC_NOW        = 6_520_442 + 5_353_109     # 11,873,551 現在残高（亨＋美香）
DC_PRIN_NOW   = 5_780_576 + 4_747_670     # 10,528,246 拠出元本累計（現在）
DC_GAIN_NOW   = DC_NOW - DC_PRIN_NOW       # 1,345,305 含み益（+12.8%）
DC_ANNUAL     = 55_000 * 2 * 12            # 1,320,000 夫婦の年間拠出（各月55,000円）
DC_RATE       = 0.07                       # 想定運用利回り（年7%・保守的に設定）
DC_YEARS      = 25                         # 亨65歳まで

# ---- 不動産投資ローン（オリックス銀行・田中亨名義・2026/06/13照会）----
# 当初借入額＝台帳の評価額が一致することで物件を同定（台帳の負債欄は
# 西台/鵜の木のローン残高が取り違えられていたため、ここで訂正している）
LOAN_UNOKI_GEN  = 25_500_000    # 鵜の木（貸出000001）当初借入
LOAN_UNOKI_ZAN  = 24_266_159    # 鵜の木 残高
LOAN_UNOKI_PMT  = 85_787        # 鵜の木 毎月返済
LOAN_NISHI_GEN  = 22_500_000    # 西台（貸出000002）当初借入
LOAN_NISHI_ZAN  = 21_411_324    # 西台 残高
LOAN_NISHI_PMT  = 75_694        # 西台 毎月返済
RE_LOAN_TOTAL   = LOAN_UNOKI_ZAN + LOAN_NISHI_ZAN     # 45,677,483
RE_PMT_TOTAL    = LOAN_UNOKI_PMT + LOAN_NISHI_PMT     # 161,481 /月

# ---- 住宅ローン（人宿町マンション・静岡銀行 本店営業部・2026/06/25 利率変更通知で更新）----
# 「お借入れ明細表（変動金利方式）」2026-06-25作成分（受領資料_20260702）で更新。
# 旧基準：2025/12照会5枚＝残高57,090,963（2026年1月・MF登録値と一致）。
HOME_LOAN       = 56_359_539    # 残高（2026年6月25日通知書時点・お取扱番号00002-0962137）
HOME_LOAN_GEN   = 62_800_000    # 当初借入額（2022年6月24日）
HOME_LOAN_PMT   = 160_944       # 毎月返済（毎月26日・利率変更後も据置）
HOME_LOAN_RATE  = 1.075         # 適用金利（変動・2026/07〜2026/12）
HOME_LOAN_END   = "2057-06"     # 最終回ご返済日 2057年6月（35年返済）
# 金利推移：0.425%（2022借入〜2024）→0.575%（2025/01〜）→0.825%（2025/07〜）→1.075%（2026/07〜・+0.25%）

CARD_DEBT       = 763_270       # クレジットカード未決済（MF負債・2026-06-13）
DEBT_TOTAL      = HOME_LOAN + RE_LOAN_TOTAL + CARD_DEBT   # 102,800,292（住宅ローンのみ2026-06-25通知値。MF負債103,531,716は2026年1月基準＝差はローン返済進捗）
NET_WORTH       = TOTAL_ASSET - DEBT_TOTAL               # 49,166,424

# ---- サブリース保証賃料（契約書で確定・2026/06/13）----
RENT_UNOKI_M    = 86_000        # 鵜の木：特定賃貸借契約 86,000円/月（管理費込・2025/10〜2030/09）
RENT_NISHI_M    = 76_050        # 西台：区分賃貸借契約 76,050円/月（2023/08〜2026/09・更新期限近い）
RENT_Y          = (RENT_UNOKI_M + RENT_NISHI_M) * 12     # 年間保証家賃 1,944,600

# ---- 亨さんの保険（ニッセイ・ウェルス生命／募集代理店=静岡銀行・2025/10照会）----
# 商品「つみたてねんきん2（外貨建）」予定利率連動型 米ドル建個人年金・証券FR007360
TORU_ANN_PREM_M  = 20_000       # 月払保険料（クレジットカード払・約200米ドル）
TORU_ANN_PAID    = 1_810_000    # 払込保険料 円換算累計（2025年10月時点）
TORU_ANN_FUND_USD = 6_360.36    # 積立金額（米ドル）
TORU_ANN_CV_USD   = 2_625.50    # 解約払戻金（米ドル・市場価格調整-3,595.51＋解約控除-139.35後）
TORU_ANN_START   = 2051         # 年金開始 2051年9月（65歳）
TORU_ANN_YRS     = 10           # 確定年金10年
TORU_ANN_Y       = 870_000      # 設計書試算の年金額（年・1ドル110円前提／現状の円安なら円換算は増）
TORU_ANN_MF      = 1_706_779    # MF「保険」列＝台帳「変額年金等」＝この年金の評価額

# ---- 美香さんの私的年金（ソニー生命・変額個人年金 2本・2026/06/12）----
MIKA_PREMIUM_M   = 10_122 * 2   # 月額保険料合計 20,244
MIKA_ANNUITY_Y   = 600_000 * 2  # 年金額合計 120万円/年
MIKA_ANNUITY_YRS = 15           # 15年確定
MIKA_DEATH       = 917_033 * 2  # 死亡保障合計 1,834,066

# ---- 医療保険（メットライフ生命「マイ フレキシィ」終身医療保障保険・無解約返戻金型・入院日数連動60日型）----
# 契約者=株式会社しずく（法人）／被保険者=本人。10年払込完了後に名義を法人→個人へ変更予定。
# 掛け捨て（無解約返戻金型）＝資産価値ほぼゼロ・法人負担＝家計支出でもないため、BS・純資産・資産推移には計上しない。
# 美香：設計書2026-06-10で確定（38歳・入院日額1.3万＋手術総合特約Ⅱ／先進医療特約／健康サポート特則）。
# 亨：設計書2025-09-19（基準2025-10-01・39歳）で確定。同型・60日型・入院日額1.4万＋先進医療特約／健康サポート特則3年型。
MED_MIKA_PREM_M  = 24_883       # 美香 月払保険料
MED_MIKA_PREM_Y  = 288_483      # 美香 年払（参考）
MED_MIKA_TOTAL   = 2_985_960    # 美香 10年払込総額
MED_MIKA_DAILY   = 13_000       # 美香 入院日額
MED_TORU_PREM_M  = 24_488       # 亨 月払保険料（設計書で確定）
MED_TORU_PREM_Y  = 283_999      # 亨 年払（参考）
MED_TORU_TOTAL   = 2_938_560    # 亨 10年払込総額（設計書P8・払込保険料累計）
MED_TORU_DAILY   = 14_000       # 亨 入院日額（美香1.3万より高い1.4万）
MED_PAY_YEARS    = 10           # 払込期間（完了後に名義変更・保障は終身）

# ---- 金融資産推移（マネーフォワード月次CSV）----
# CSVは 00_マスタデータ置き場\32_家計_資産データ\資産推移月次\ に集約（投入口の一本化・2026-06-25）。
# フォルダ内の最新（ファイル名末尾の日付順）を自動で読むので、毎月ファイル名を直す必要はない。
import glob as _glob
CSV_DIR = os.path.normpath(os.path.join(BASE, "..", "..", "00_マスタデータ置き場", "32_家計_資産データ", "資産推移月次"))
_csv_files = sorted(_glob.glob(os.path.join(CSV_DIR, "資産推移月次_*.csv")))
if not _csv_files:
    raise FileNotFoundError(f"資産推移CSVが見つかりません: {CSV_DIR}")
CSV_PATH = _csv_files[-1]   # ファイル名末尾の日付が最新のものを採用
# 金融資産カテゴリ（＝P1円グラフの6分類）。P1「前月からの資産の変化」の内訳に使う。
# (表示名, CSV列名, 発散バーの色)。列名はCSVヘッダと完全一致必須（cp932）。
CAT_DEFS = [
    ("現金・預金",     "預金・現金・暗号資産（円）", BLUE),
    ("高配当株",       "株式(現物)（円）",          TEAL),
    ("投資信託",       "投資信託（円）",            GOLD),
    ("企業型DC",       "年金（円）",               DARK),
    ("外貨建年金(亨)",  "保険（円）",               ROSE),
    ("ポイント",       "ポイント（円）",            GRAY),
]
trend = []      # (date_str, 金融資産 = 合計 - 不動産)
tot_raw = []    # (date_str, 総資産 = 合計)
cat_rows = {lbl: [] for lbl, _c, _col in CAT_DEFS}   # カテゴリ別 (date, 値)
with open(CSV_PATH, encoding="cp932") as f:
    for row in _csv.DictReader(f):
        total = int(row["合計（円）"])
        re_v  = int(row["不動産（円）"])
        trend.append((row["日付"], total - re_v))
        tot_raw.append((row["日付"], total))
        for lbl, _c, _col in CAT_DEFS:
            cat_rows[lbl].append((row["日付"], int(row[_c])))
trend.sort()
tot_raw.sort()
monthly = {}
month_last_date = {}          # 月 → その月の最終データ日（"YYYY/MM/DD"）
for dt, v in trend:
    monthly[dt[:7]] = v
    month_last_date[dt[:7]] = dt
monthly_tot = {}
for dt, v in tot_raw:
    monthly_tot[dt[:7]] = v

# ---- 月末基準：当月が未完了（最終データ日≠暦上の月末）なら最新月を落とす ----
# 例）7/1しか無い7月は外し、直近の完了月末（6/30）を最新にする。毎月自動判定。
import calendar as _cal
if monthly:
    _lm = max(monthly)                          # 最新月キー "YYYY/MM"
    _ly, _lmo = int(_lm[:4]), int(_lm[5:7])
    _lastday = _cal.monthrange(_ly, _lmo)[1]    # その月の暦上の末日
    _havday = int(month_last_date[_lm][8:10])   # 実際の最終データ日
    if _havday != _lastday and len(monthly) >= 2:
        del monthly[_lm]
        monthly_tot.pop(_lm, None)
trend = sorted(monthly.items())

# ---- 前月比（最新月 vs 前月・MF月次 / 月末残高ベース）----
_ms = sorted(monthly)
if len(_ms) >= 2:
    CUR_M, PREV_M = _ms[-1], _ms[-2]
    FIN_DIFF = monthly[CUR_M] - monthly[PREV_M]
    TOT_DIFF = monthly_tot[CUR_M] - monthly_tot[PREV_M]
else:
    CUR_M = PREV_M = None
    FIN_DIFF = TOT_DIFF = 0

# ---- 前月比（カテゴリ別・最新月 vs 前月）----
# 各カテゴリの月次は「その月の最終スナップショット」を採用（金融資産／総資産と同じ流儀）。
cat_diff = []   # (label, prev, cur, diff, color)  ※描画側で寄与額の大きい順にソート
if CUR_M and PREV_M:
    for lbl, _c, col in CAT_DEFS:
        mm = {}
        for dt, v in sorted(cat_rows[lbl]):
            mm[dt[:7]] = v
        cur = mm.get(CUR_M, 0)
        prev = mm.get(PREV_M, 0)
        cat_diff.append((lbl, prev, cur, cur - prev, col))

# ---- 世帯実年収（旅費_世帯年収まとめ.xlsx を参照・2026-07-01 リンク）----
# 出張旅費の自動転記先ブックから、給与＋非課税の出張日当＋DPC＋配偶者収入＋企業型DCを
# 年ごとに読み、世帯の「実年収」を構成する。旅費行(=B16等)とSUM行は数式でopenpyxlの
# キャッシュがNoneのことがあるため、旅費は月セルから再計算し、他は定数セルを直読みする。
import openpyxl as _oxl
INCOME_XLSX = os.path.normpath(os.path.join(
    BASE, "..", "..", "..", "02_しずく", "05_出張旅費", "旅費_世帯年収まとめ.xlsx"))
INCOME_OK = os.path.exists(INCOME_XLSX)
income_years = []   # [dict(year, toru, nittou, dpc, mika, dc, total, lastm)]
_iw = None
if INCOME_OK:
    try:                                        # xlsxが壊れていてもレポート全体は止めない（P6は注記表示へ退避）
        _wb = _oxl.load_workbook(INCOME_XLSX, data_only=True)
        _iw = _wb["Sheet1"] if "Sheet1" in _wb.sheetnames else _wb.worksheets[0]
    except Exception as _ie:
        INCOME_OK = False
        print("P6 世帯年収xlsxの読込に失敗（レポートは継続）:", type(_ie).__name__, _ie)
if INCOME_OK and _iw is not None:
    _YCOL = {2025: 2, 2026: 3, 2027: 4, 2028: 5, 2029: 6, 2030: 7}

    def _num(r, cc):
        v = _iw.cell(r, cc).value
        return float(v) if isinstance(v, (int, float)) else 0.0

    def _travel(first, cc):
        return sum(_num(first + m, cc) for m in range(12))

    def _last_travel_month(cc):
        last = 0
        for m in range(12):
            v = _iw.cell(4 + m, cc).value
            if isinstance(v, (int, float)) and v > 0:
                last = m + 1
        return last

    for _y, _cc in _YCOL.items():
        toru = _num(35, _cc); dpc = _num(38, _cc); mika = _num(39, _cc)
        dc = _num(40, _cc) + _num(41, _cc)
        nittou = _travel(4, _cc) + _travel(19, _cc)   # 亨＋美香の出張日当（非課税）
        if toru + dpc + mika <= 0:      # 給与系が無い年（DCのみ・未確定）は推移に載せない
            continue
        income_years.append(dict(
            year=_y, toru=toru, nittou=nittou, dpc=dpc, mika=mika, dc=dc,
            total=toru + nittou + dpc + mika + dc, lastm=_last_travel_month(_cc)))
    income_years.sort(key=lambda d: d["year"])


# ============================================================ 恩株データ（高配当株の配当による元本回収）
# 日本株＝SBI保有証券CSV＋配当履歴CSV、SPYD＝外国株式の保有画面（取得額）から集計。
# 配当は税引後・生涯累計。恩株達成度＝生涯受取配当÷取得額。更新時はこの定数を差し替える。
ONK_COST = 7_186_538            # 取得額(簿価) 夫婦日本株＋SPYD
ONK_DIV = 598_905              # 生涯 受取配当(税引後・保有銘柄)
ONK_PCT = ONK_DIV / ONK_COST * 100   # 恩株達成度（約8.3%）
ONK_REM_Y = 27                # 現配当維持での残り年数(概算)
ONK_TTM = 246_904             # 年間配当ペース（直近1年・税引後）
ONK_LIFE_TOTAL = 763_777       # 受取配当 生涯累計(国内＋米国・売却分含む)
ONK_BY_YEAR = [(2022, 12_956), (2023, 134_392), (2024, 251_449), (2025, 242_525), (2026, 122_456)]
# 銘柄別 恩株（上位12）：(銘柄, 所有者, 口座, 取得単価, 株数, 取得額, 配当累計, 恩株%)
ONK_RANK = [
    ("商船三井", "亨", "特定", 3294, 50, 164_700, 31_081, 18.9),
    ("郵船", "亨", "特定", 3141, 50, 157_050, 27_695, 17.6),
    ("SPYD(米高配当ETF)", "亨", "特+N", 5359, 312, 1_672_068, 210_932, 12.6),
    ("第一ライフグループ", "亨", "特定", 719, 400, 287_600, 35_494, 12.3),
    ("武田薬", "亨", "NISA", 4132, 100, 413_200, 48_590, 11.8),
    ("アステラス薬", "亨", "NISA", 1605, 100, 160_500, 15_899, 9.9),
    ("ＮＴＴ", "亨", "NISA", 151, 500, 75_500, 7_283, 9.6),
    ("東京海上", "亨", "特定", 3938, 100, 393_800, 37_944, 9.6),
    ("丸紅", "亨", "NISA", 2613, 100, 261_300, 25_129, 9.6),
    ("伯東", "亨", "NISA", 5080, 100, 508_000, 47_674, 9.4),
    ("三菱商事", "亨", "特定", 2700, 100, 270_000, 18_412, 6.8),
    ("伊藤忠", "亨", "特定", 1139, 500, 569_500, 36_499, 6.4),
]
ONK_MORE = "ほか9銘柄：ヤマハ発5.8%・住友商3.5%・三菱UFJ/三菱重工1.5%・倉元/DCM/ファンペップは無配または新規"
ONK_ACCT = [  # (口座, 取得額, 恩株%, 直近1年配当)
    ("特定口座(国内)", 2_364_570, 8.3, 80_136),
    ("NISA口座(国内)", 3_149_900, 6.1, 82_300),
    ("SPYD(米国ETF)", 1_672_068, 12.6, 84_468),
]
ONK_NEAR = "郵船 残り約14年 ／ 第一ライフ 約14.5年 ／ 商船三井 約16.8年"

c = canvas.Canvas(OUT, pagesize=A4)
c.setTitle("資産管理台帳 分析レポート 2026-07-08 v30（v29の住宅ローン利率更新等の家計改善を土台に、P3へ高配当株の恩株ページを再統合＝配当による元本回収を全保有銘柄・口座別・ハイライトで可視化。夫婦合算＋SPYD。7ページ構成。恩株はP3に恒久化＝再生成時も落とさない）")


# ============================================================ 共通関数
def Y(t):
    return H - t


def text(x, t, s, size=9, font="JP", color=TXT, anchor="l"):
    c.setFont(font, size)
    c.setFillColor(color)
    if anchor == "l":
        c.drawString(x, Y(t), s)
    elif anchor == "r":
        c.drawRightString(x, Y(t), s)
    else:
        c.drawCentredString(x, Y(t), s)


def wrap(s, font, size, maxw):
    lines, cur = [], ""
    for ch in s:
        if pdfmetrics.stringWidth(cur + ch, font, size) > maxw:
            lines.append(cur)
            cur = ch
        else:
            cur += ch
    if cur:
        lines.append(cur)
    return lines


def box(x, t, w, h, fill=LGRAY, stroke=None, r=4):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
    c.roundRect(x, Y(t + h), w, h, r, fill=1, stroke=1 if stroke else 0)


def oku_man(n):
    """円 → 「1.48億」or「4,159万」表記"""
    if n >= 100_000_000:
        return f"{n / 100_000_000:.2f}億"
    return f"{int(n // 10_000):,}万"


# ============================================================ Page 1 ダッシュボード
c.setFillColor(NAVY)
c.rect(0, Y(64), W, 64, fill=1, stroke=0)
text(ML, 34, "資産管理 分析レポート", 19, "JPB", WHITE)
text(ML, 52, "田中亨・美香 世帯（本人40歳／妻39歳／湊人さん8歳）", 8.5, "JP", HexColor("#C8D2E4"))
text(W - MR, 34, "台帳基準日 2026-06-12", 9, "JP", HexColor("#C8D2E4"), "r")
text(W - MR, 48, "分析日 2026-07-07", 9, "JP", HexColor("#C8D2E4"), "r")

# ---- KPIカード
kpis = [
    ("総資産",  "1億5,196万円", "金融4,117万＋不動産1億1,080万", DARK),
    ("総負債",  "1億280万円",   "ローン計1億204万＋カード76万",  RED),
    ("純資産",  "4,917万円",    "自己資本比率 32.4%",           TEAL),
]
kw = (CW - 20) / 3
for i, (label, big, small, col) in enumerate(kpis):
    x = ML + i * (kw + 10)
    box(x, 78, kw, 62, LGRAY)
    c.setFillColor(col)
    c.rect(x, Y(140), 3, 62, fill=1, stroke=0)
    text(x + 12, 95, label, 9, "JPB", SUB)
    text(x + 12, 116, big, 15, "JPB", col)
    text(x + 12, 132, small, 7.5, "JP", SUB)

# ---- 左：金融資産の内訳（円グラフ＋凡例）
text(ML, 165, "■ 金融資産の内訳（4,117万円）", 10.5, "JPB", NAVY)
fin = [
    ("現金・預金",  6684804, BLUE),
    ("高配当株",    9361519, TEAL),
    ("投資信託",   10960733, GOLD),
    ("企業型DC",   12285029, DARK),
    ("外貨建年金(亨)", 1706779, ROSE),
    ("ポイント",     167852, GRAY),
]
ftotal = sum(v for _, v, _ in fin)

d = Drawing(120, 120)
pie = Pie()
pie.x = pie.y = 5
pie.width = pie.height = 110
pie.data = [v for _, v, _ in fin]
pie.slices.strokeColor = WHITE
pie.slices.strokeWidth = 1
for i, (_, _, col) in enumerate(fin):
    pie.slices[i].fillColor = col
pie.labels = None
d.add(pie)
renderPDF.draw(d, c, ML, Y(300))

lx = ML + 132
ly = 192
for name, v, col in fin:
    c.setFillColor(col)
    c.rect(lx, Y(ly + 2), 8, 8, fill=1, stroke=0)
    text(lx + 13, ly + 6, name, 8, "JP", TXT)
    text(lx + 118, ly + 6, f"{v:,}円", 8, "JP", TXT, "r")
    text(lx + 148, ly + 6, f"{v/ftotal*100:.1f}%", 8, "JP", SUB, "r")
    ly += 17

text(ML, 322, "リスク資産（株式・投信・DC・外貨建年金）が金融資産の83.3%", 7.5, "JP", SUB)

# ---- 右：バランスシート構成（積み上げ棒）
bx = ML + 300
text(bx, 165, "■ バランスシート構成", 10.5, "JPB", NAVY)
scale = 132 / TOTAL_ASSET
bar_w = 58
base_t = 312

fin_h = FIN_TOTAL * scale
re_h  = RE_TOTAL * scale
c.setFillColor(BLUE)
c.rect(bx + 10, Y(base_t), bar_w, re_h, fill=1, stroke=0)
c.setFillColor(TEAL)
c.rect(bx + 10, Y(base_t - re_h), bar_w, fin_h, fill=1, stroke=0)
text(bx + 10 + bar_w / 2, base_t - re_h / 2 + 3, "不動産", 8, "JPB", WHITE, "c")
text(bx + 10 + bar_w / 2, base_t - re_h / 2 + 13, "1億1,080万", 7, "JP", WHITE, "c")
text(bx + 10 + bar_w / 2, base_t - re_h - fin_h / 2 + 3, "金融資産", 8, "JPB", WHITE, "c")
text(bx + 10 + bar_w / 2, base_t - re_h - fin_h / 2 + 13, "4,117万", 7, "JP", WHITE, "c")
text(bx + 10 + bar_w / 2, base_t + 12, "資産", 8.5, "JPB", SUB, "c")

debt_h = DEBT_TOTAL * scale
nw_h   = NET_WORTH * scale
bx2 = bx + 10 + bar_w + 34
c.setFillColor(HexColor("#D08770"))
c.rect(bx2, Y(base_t), bar_w, debt_h, fill=1, stroke=0)
c.setFillColor(HexColor("#A3BE8C"))
c.rect(bx2, Y(base_t - debt_h), bar_w, nw_h, fill=1, stroke=0)
text(bx2 + bar_w / 2, base_t - debt_h / 2 + 3, "負債", 8, "JPB", WHITE, "c")
text(bx2 + bar_w / 2, base_t - debt_h / 2 + 13, "1億280万", 7, "JP", WHITE, "c")
text(bx2 + bar_w / 2, base_t - debt_h - nw_h / 2 + 3, "純資産", 8, "JPB", WHITE, "c")
text(bx2 + bar_w / 2, base_t - debt_h - nw_h / 2 + 13, "4,917万", 7, "JP", WHITE, "c")
text(bx2 + bar_w / 2, base_t + 12, "負債＋純資産", 8.5, "JPB", SUB, "c")

text(bx, 340, "総資産の72.9%が不動産。負債比率67.6%。", 7.5, "JP", SUB)

# ---- 不動産×ローン明細表
text(ML, 358, "■ 不動産とローンの対応（含み損益＝評価額−ローン残高）", 10.5, "JPB", NAVY)
cols = [ML, ML + 92, ML + 134, ML + 222, ML + 310, ML + 398, ML + 446, W - MR]
headers = ["物件", "用途", "評価額", "ローン残高", "含み損益", "LTV", "家賃収入"]
rt = 368
c.setFillColor(NAVY)
c.rect(ML, Y(rt + 15), CW, 15, fill=1, stroke=0)
for i, htxt in enumerate(headers):
    if i < 2:
        text(cols[i] + 4, rt + 11, htxt, 8, "JPB", WHITE)
    else:
        text(cols[i + 1] - 4, rt + 11, htxt, 8, "JPB", WHITE, "r")

rows = [
    ("人宿町マンション", "自宅",  "62,800,000", "56,359,539", "+6,440,461", "89.7%", "－",         False),
    ("鵜の木",          "投資用", "25,500,000", "24,266,159", "+1,233,841", "95.2%", "8.6万円/月", False),
    ("西台",            "投資用", "22,500,000", "21,411,324", "+1,088,676", "95.2%", "7.6万円/月", False),
    ("合計",            "",      "110,800,000", "102,037,022", "+8,762,978", "92.1%", "16.2万円/月", False),
]
rt += 15
for name, use, val, loan, eq, ltv, rent, warn in rows:
    is_total = name == "合計"
    c.setFillColor(REDBG if warn else (LGRAY if is_total else WHITE))
    c.rect(ML, Y(rt + 16), CW, 16, fill=1, stroke=0)
    fnt = "JPB" if (is_total or warn) else "JP"
    fcol = RED if warn else TXT
    text(cols[0] + 4, rt + 12, name, 8, fnt, fcol)
    text(cols[1] + 4, rt + 12, use, 8, "JP", fcol)
    for j, v in enumerate([val, loan, eq, ltv, rent]):
        col2 = RED if v.startswith("▲") else (RED if warn else TXT)
        text(cols[j + 3] - 4, rt + 12, v, 8, fnt, col2, "r")
    rt += 16
c.setStrokeColor(MGRAY)
c.setLineWidth(0.5)
c.line(ML, Y(rt), ML + CW, Y(rt))
text(ML, rt + 11,
     "※住宅ローンは静岡銀行の利率変更通知で更新（残高5,636万・基準日2026年6月25日）、投資用2戸はオリックス照会値。カード76万はMF（2026-06-13）。",
     7, "JP", SUB)

# ---- 指標チップ
chips = [
    ("不動産比率", "72.9%", "総資産に占める割合"),
    ("負債比率", "67.6%", "総負債÷総資産"),
    ("生活防衛資金", "668万円", "現金・預金"),
    ("投資用表面利回り", "4.1%", "年間家賃195万÷評価4,800万"),
]
ct = rt + 24
cw2 = (CW - 30) / 4
for i, (label, val, note) in enumerate(chips):
    x = ML + i * (cw2 + 10)
    box(x, ct, cw2, 46, LGRAY)
    text(x + 8, ct + 14, label, 7.5, "JP", SUB)
    text(x + 8, ct + 30, val, 13, "JPB", NAVY)
    text(x + 8, ct + 41, note, 6.3, "JP", SUB)

# ---- 総合評価ボックス
st = ct + 60
box(ML, st, CW, 90, TEALBG)
c.setFillColor(TEAL)
c.rect(ML, Y(st + 90), 3, 90, fill=1, stroke=0)
text(ML + 12, st + 16, "総合評価", 9.5, "JPB", DARK)
summary = ("40歳で純資産約4,917万円・年間積立力は同年代上位の優良世帯。住宅ローンは静岡銀行の利率変更通知（2026-06-25）で残高を5,636万に更新。"
           "投資用2戸はいずれも含み益プラス（合計＋876万円）。亨さんの保険も外貨建個人年金（ニッセイ・ウェルス／静岡銀行窓販）と判明。"
           "ただし不動産レバレッジは依然高く、住宅0.425→1.075%・投資用2.25→2.85%（予定）と3本すべての変動金利が上昇しており、金利上昇が最大のリスク。"
           "優先課題は ①変動金利上昇への備え（投資用は+1%で年46万円の利息増） ②将来予測利回りの現実化 ③教育資金の独立確保 の3点。")
yy = st + 31
for ln in wrap(summary, "JP", 8.5, CW - 24):
    text(ML + 12, yy, ln, 8.5, "JP", TXT)
    yy += 13

# ---- 前月からの資産の変化（カテゴリ別・MF月次スナップショット）
# 総合評価ボックス（st+90）の下の余白に配置。金融資産の6分類の増減を発散バーで見える化。
if CUR_M and PREV_M and cat_diff:
    ch_t = st + 90 + 20
    _pm, _cm = int(PREV_M[5:7]), int(CUR_M[5:7])
    text(ML, ch_t, f"■ 前月からの資産の変化（MF月次・{_pm}月末 → {_cm}月末）", 10.5, "JPB", NAVY)
    _tsgn = "+" if TOT_DIFF >= 0 else "−"
    text(W - MR, ch_t, f"総資産 {_tsgn}{abs(TOT_DIFF) // 10000:,}万円",
         11, "JPB", TEAL if TOT_DIFF >= 0 else RED, "r")
    axis_x = ML + 312           # 発散バーの中心（増=右/減=左）
    half_w = 110
    maxabs = max((abs(d) for _l, _p, _cc, d, _col in cat_diff), default=1) or 1
    ry = ch_t + 18
    for lbl, prev, cur, diff, col in sorted(cat_diff, key=lambda r: -abs(r[3])):
        man = abs(diff) // 10000                    # 増減の万単位（絶対値・切り捨て）
        sgn = 1 if diff >= 0 else -1
        prev_m = prev // 10000
        cur_m = prev_m + sgn * man                  # prev→cur を増減表示と必ず一致させる
        text(ML + 4, ry + 10, lbl, 8, "JP", TXT)
        text(ML + 108, ry + 10, f"{prev_m:,}万→{cur_m:,}万", 7, "JP", SUB)
        c.setStrokeColor(MGRAY)
        c.setLineWidth(0.5)
        c.line(axis_x, Y(ry + 1), axis_x, Y(ry + 13))
        bw = half_w * abs(diff) / maxabs
        if diff >= 0:
            c.setFillColor(TEAL)
            c.rect(axis_x, Y(ry + 12), bw, 9, fill=1, stroke=0)
        else:
            c.setFillColor(RED)
            c.rect(axis_x - bw, Y(ry + 12), bw, 9, fill=1, stroke=0)
        _dsgn = "+" if diff >= 0 else "−"
        text(W - MR, ry + 10, f"{_dsgn}{abs(diff) // 10000:,}万円",
             8, "JPB", TEAL if diff >= 0 else RED, "r")
        ry += 15
    text(ML, ry + 10,
         "※MFの各月末残高で前月末と比較（当月途中の日次は月末確定後に反映）。不動産は評価額固定。金額は万円未満切り捨て。",
         6.5, "JP", SUB)

text(W / 2, 826, "－ 1 / 7 －", 8, "JP", GRAY, "c")
c.showPage()

# ============================================================ Page 2 投資の見える化
c.setFillColor(NAVY)
c.rect(0, Y(44), W, 44, fill=1, stroke=0)
text(ML, 28, "投資の見える化（NISA・企業型DC）", 14, "JPB", WHITE)
text(W - MR, 28, "1,800万円の枠と、老後資産の積み上がり", 8, "JP", HexColor("#C8D2E4"), "r")

# ---- NISA 生涯投資枠の進捗
nt = 64
text(ML, nt, "■ NISA 生涯投資枠の進捗（夫婦で最大3,600万円の非課税枠）", 10.5, "JPB", NAVY)

text(ML, nt + 15, "田中亨さん（SBI証券・2026/06/12時点）", 8, "JPB", TXT)
text(W - MR, nt + 15, f"{NISA_T_USED:,}円 / 1,800万円（{NISA_T_USED/NISA_LIFE_CAP*100:.1f}%）",
     8.5, "JPB", NAVY, "r")
bar_t, bar_h = nt + 20, 14
w_g = CW * NISA_T_GROWTH / NISA_LIFE_CAP
w_t = CW * NISA_T_TSUMI / NISA_LIFE_CAP
c.setFillColor(MGRAY)
c.rect(ML, Y(bar_t + bar_h), CW, bar_h, fill=1, stroke=0)
c.setFillColor(BLUE)
c.rect(ML, Y(bar_t + bar_h), w_g, bar_h, fill=1, stroke=0)
c.setFillColor(TEAL)
c.rect(ML + w_g, Y(bar_t + bar_h), w_t, bar_h, fill=1, stroke=0)
for m in range(0, 19, 3):
    x = ML + CW * m / 18
    c.setStrokeColor(GRAY)
    c.setLineWidth(0.4)
    c.line(x, Y(bar_t + bar_h), x, Y(bar_t + bar_h + 3))
    text(x, bar_t + bar_h + 11, f"{m*100:,}万" if m else "0", 6, "JP", GRAY, "c")
c.setFillColor(BLUE)
c.rect(ML, Y(nt + 52), 7, 7, fill=1, stroke=0)
text(ML + 11, nt + 58, f"成長投資枠 {NISA_T_GROWTH:,}円（上限1,200万円の{NISA_T_GROWTH/NISA_GROWTH_CAP*100:.1f}%）", 7.5, "JP", TXT)
c.setFillColor(TEAL)
c.rect(ML + 230, Y(nt + 52), 7, 7, fill=1, stroke=0)
text(ML + 241, nt + 58, f"つみたて投資枠 {NISA_T_TSUMI:,}円", 7.5, "JP", TXT)
text(W - MR, nt + 58, f"評価額 {NISA_T_VALUE:,}円（損益 +{NISA_T_PL:,}円／+20.4%）", 7.5, "JP", TEAL, "r")

# 美香さんのバー（SBI証券・2026/06/13）
text(ML, nt + 73, "田中美香さん（SBI証券・2026/06/13時点）", 8, "JPB", TXT)
text(W - MR, nt + 73, f"{NISA_M_USED:,}円 / 1,800万円（{NISA_M_USED/NISA_LIFE_CAP*100:.1f}%）",
     8.5, "JPB", NAVY, "r")
bar2_t = nt + 78
wm_t = CW * NISA_M_TSUMI / NISA_LIFE_CAP
c.setFillColor(MGRAY)
c.rect(ML, Y(bar2_t + bar_h), CW, bar_h, fill=1, stroke=0)
c.setFillColor(TEAL)
c.rect(ML, Y(bar2_t + bar_h), wm_t, bar_h, fill=1, stroke=0)
for m in range(0, 19, 3):
    x = ML + CW * m / 18
    c.setStrokeColor(GRAY)
    c.setLineWidth(0.4)
    c.line(x, Y(bar2_t + bar_h), x, Y(bar2_t + bar_h + 3))
c.setFillColor(TEAL)
c.rect(ML, Y(nt + 110), 7, 7, fill=1, stroke=0)
text(ML + 11, nt + 116, f"つみたて投資枠 {NISA_M_TSUMI:,}円（成長枠は未使用・1,200万円が空き）", 7.5, "JP", TXT)
text(W - MR, nt + 116, f"評価額 {NISA_M_VALUE:,}円（損益 +{NISA_M_PL:,}円／+51.3%）", 7.5, "JP", TEAL, "r")
text(ML, nt + 132, f"夫婦合計：{NISA_COUPLE_USED:,}円 / 3,600万円（{NISA_COUPLE_USED/36_000_000*100:.1f}%）",
     8.5, "JPB", DARK)

# ---- 成長投資枠を優先して埋める（ユーザー方針・2026-06-16）----
gt0 = nt + 152
text(ML, gt0, "■ 成長投資枠（1人1,200万円）を優先して埋める ― 夫婦で最大2,400万円", 10.5, "JPB", NAVY)
text(ML, gt0 + 14,
     "方針：商品の幅が広く枠が限られる成長枠（1,200万円）を先に満たし、残り600万円をつみたて枠で。年間上限は成長240万円／つみたて120万円。",
     7.2, "JP", SUB)


def growth_bar(label, used, ytop, accent):
    """成長枠（1,200万円）の進捗バー。240万円（＝年間上限1年分）ごとに目盛りを刻む。"""
    text(ML, ytop, label, 8, "JPB", TXT)
    text(W - MR, ytop, f"{used:,}円 / 1,200万円（{used/NISA_GROWTH_CAP*100:.1f}%）", 8.5, "JPB", NAVY, "r")
    by, bh = ytop + 5, 13
    c.setFillColor(MGRAY)
    c.rect(ML, Y(by + bh), CW, bh, fill=1, stroke=0)
    c.setFillColor(accent)
    c.rect(ML, Y(by + bh), CW * used / NISA_GROWTH_CAP, bh, fill=1, stroke=0)
    for k in range(1, 5):                       # 240万円ごと＝1年分の上限ペース
        x = ML + CW * k / 5
        c.setStrokeColor(WHITE)
        c.setLineWidth(0.8)
        c.line(x, Y(by + bh), x, Y(by))
    return by + bh


gy = growth_bar("田中亨さん", NISA_T_GROWTH, gt0 + 34, BLUE)
text(ML, gy + 11,
     f"2026年は上限240万円のうち約197万円を消化＝すでに成長枠を優先中。生涯の成長枠残りは{NISA_T_GROWTH_LEFT:,}円。",
     7.2, "JP", SUB)
gy2 = growth_bar("田中美香さん", NISA_M_GROWTH, gt0 + 70, ROSE)
text(ML, gy2 + 11,
     "全額つみたて枠（月1万円）のみ。成長枠1,200万円が丸ごと空き＝夫婦で最優先に埋めるべき枠はここ。",
     7.2, "JP", RED)

gbx = gt0 + 92
box(ML, gbx, CW, 40, TEALBG)
c.setFillColor(TEAL)
c.rect(ML, Y(gbx + 40), 3, 40, fill=1, stroke=0)
text(ML + 12, gbx + 15, f"残りの成長枠は夫婦合計 {NISA_COUPLE_GROWTH_LEFT:,}円（亨 885万＋美香 1,200万）。", 8, "JPB", DARK)
text(ML + 12, gbx + 30,
     "美香さんの1,200万円を埋めるペース＝年240万（上限）で5年／月10万円で10年／月5万円で20年。積立を成長枠へ振り向けると非課税メリット最大。",
     7.4, "JP", DARK)

# ---- 企業型DCの積み上げ（年7%）
dt = nt + 300
text(ML, dt, "■ 企業型DCの積み上げ（夫婦合算・全額S&P500・想定年7%）", 10.5, "JPB", NAVY)
text(W - MR, dt, "現在 1,187万円（含み益+12.8%）／掛金 月11万円", 7.5, "JP", SUB, "r")

# DC将来シリーズを計算（5年刻み）
dc_series = []   # (年, 年齢, 元本累計, 運用益, 残高)
for t in range(0, DC_YEARS + 1, 5):
    bal = DC_NOW * (1 + DC_RATE) ** t + DC_ANNUAL * (((1 + DC_RATE) ** t - 1) / DC_RATE)
    prin = DC_PRIN_NOW + DC_ANNUAL * t
    dc_series.append((2026 + t, 40 + t, prin, bal - prin, bal))

gtop = dt + 14
gh = 116
gymax = 160_000_000
gx0 = ML + 40
gx1 = W - MR - 6
nb = len(dc_series)
bw = 30
# y軸グリッド
for gv in range(0, gymax + 1, 40_000_000):
    yline = gtop + gh - gh * gv / gymax
    c.setStrokeColor(MGRAY)
    c.setLineWidth(0.4)
    c.line(gx0, Y(yline), gx1, Y(yline))
    if gv == 0:
        lbl = "0"
    elif gv < 100_000_000:
        lbl = f"{gv // 10_000_000 * 1000:,}万"
    else:
        lbl = f"{gv / 100_000_000:.1f}億"
    text(gx0 - 4, yline + 2, lbl, 6, "JP", GRAY, "r")
# 積み上げ棒
slot = (gx1 - gx0) / nb
for i, (yr, age, prin, gain, bal) in enumerate(dc_series):
    cxb = gx0 + slot * (i + 0.5)
    ph = gh * prin / gymax
    gnh = gh * gain / gymax
    c.setFillColor(DARK)                       # 拠出元本（累計）
    c.rect(cxb - bw / 2, Y(gtop + gh), bw, ph, fill=1, stroke=0)
    c.setFillColor(TEAL)                       # 運用益（複利）
    c.rect(cxb - bw / 2, Y(gtop + gh - ph), bw, gnh, fill=1, stroke=0)
    text(cxb, gtop + gh - ph - gnh - 3, oku_man(bal), 7, "JPB", NAVY, "c")
    text(cxb, gtop + gh + 9, f"{yr}", 6.3, "JP", SUB, "c")
    text(cxb, gtop + gh + 17, f"{age}歳", 6, "JP", GRAY, "c")
# 凡例
lg = gtop + gh + 26
c.setFillColor(DARK)
c.rect(gx0, Y(lg), 7, 7, fill=1, stroke=0)
text(gx0 + 11, lg + 6, "拠出元本（累計）", 7, "JP", TXT)
c.setFillColor(TEAL)
c.rect(gx0 + 110, Y(lg), 7, 7, fill=1, stroke=0)
text(gx0 + 121, lg + 6, "運用益（複利）", 7, "JP", TXT)

# DCポイントボックス
dpt = lg + 14
box(ML, dpt, CW, 36, TEALBG)
c.setFillColor(TEAL)
c.rect(ML, Y(dpt + 36), 3, 36, fill=1, stroke=0)
dcp = "現在1,187万円＋月11万円（年132万円）を想定年7%で25年運用 → 65歳で約1.48億円（拠出元本4,353万＋運用益1.04億）。"
dcp2 = "台帳の将来予測「1.7億円」は年8%前提。本レポートは保守的に年7%で見ており、年5%なら約1.0億円。"
text(ML + 12, dpt + 15, dcp, 7.6, "JP", DARK)
text(ML + 12, dpt + 28, dcp2, 7.6, "JP", SUB)

# ---- 金融資産の歩み（不動産除く・マネーフォワード月次）※P4から移設・拡大（余白平準化・2026-07-06 v25）
fw_t = dpt + 36 + 24
text(ML, fw_t, "■ 金融資産の歩み（不動産除く）", 10.5, "JPB", NAVY)
fw_t += 10
fw_ch = 90
fw_ymax = 40_000_000
fw_n = len(trend)


def fw_cx(i):
    return ML + CW * i / (fw_n - 1)


def fw_cy(v):
    return fw_t + fw_ch - fw_ch * v / fw_ymax


for gv in range(0, fw_ymax + 1, 10_000_000):
    yline = fw_t + fw_ch - fw_ch * gv / fw_ymax
    c.setStrokeColor(MGRAY)
    c.setLineWidth(0.4)
    c.line(ML, Y(yline), ML + CW, Y(yline))
    if gv:
        text(ML - 2, yline + 2, f"{gv // 10_000_000},000万", 6, "JP", GRAY, "r")
for i, (ym, _) in enumerate(trend):
    if ym.endswith("/01"):
        text(fw_cx(i), fw_t + fw_ch + 9, ym[:4], 6, "JP", GRAY, "c")
fw_p = c.beginPath()
fw_p.moveTo(fw_cx(0), Y(fw_t + fw_ch))
for i, (_, v) in enumerate(trend):
    fw_p.lineTo(fw_cx(i), Y(fw_cy(v)))
fw_p.lineTo(fw_cx(fw_n - 1), Y(fw_t + fw_ch))
fw_p.close()
c.setFillColor(TEALBG)
c.drawPath(fw_p, fill=1, stroke=0)
fw_p2 = c.beginPath()
fw_p2.moveTo(fw_cx(0), Y(fw_cy(trend[0][1])))
for i, (_, v) in enumerate(trend):
    fw_p2.lineTo(fw_cx(i), Y(fw_cy(v)))
c.setStrokeColor(TEAL)
c.setLineWidth(1.3)
c.drawPath(fw_p2, fill=0, stroke=1)
c.setFillColor(TEAL)
c.circle(fw_cx(fw_n - 1), Y(fw_cy(trend[-1][1])), 2.2, fill=1, stroke=0)
text(W - MR, fw_cy(trend[-1][1]) - 5, f"{trend[-1][1]:,}円", 7.5, "JPB", DARK, "r")
# ※x軸の年ラベル（fw_t+fw_ch+9）と注記の間隔を15pt確保し、旧版で起きていた重なりを解消
text(ML, fw_t + fw_ch + 24, "※2026年1月の急増は口座連携の追加（株式・DC口座等）によるもの。実体の増加と区別して読むこと。", 6.5, "JP", SUB)

text(W / 2, 826, "－ 2 / 7 －", 8, "JP", GRAY, "c")
c.showPage()

# ============================================================ Page 3 投資の見える化②（高配当株の恩株）
c.setFillColor(NAVY)
c.rect(0, Y(44), W, 44, fill=1, stroke=0)
text(ML, 28, "投資の見える化 ②  高配当株の恩株", 14, "JPB", WHITE)
text(W - MR, 28, "配当だけで取得元本をどれだけ取り戻したか", 8, "JP", HexColor("#C8D2E4"), "r")

_man = lambda v: f"{int(round(v / 10000)):,}万"
y0 = 64
text(ML, y0, "■ 高配当株の恩株（配当による元本回収・夫婦合算＋SPYD）", 10.5, "JPB", NAVY)
text(ML, y0 + 13, "日本株 夫婦20銘柄＋SPYD（米国高配当ETF）／配当は税引後。恩株達成度＝生涯受取配当÷取得額。100%で配当だけで元本回収＝恩株。", 7.2, "JP", SUB)
onk_kpis = [
    ("取得額（簿価）", _man(ONK_COST) + "円", "日本株＋SPYD", DARK),
    ("生涯 受取配当(税引後)", f"{ONK_DIV:,}円", "保有銘柄から累計", TEAL),
    ("恩株 達成度", f"{ONK_PCT:.1f}%", "配当÷取得額", GOLD),
    ("年間配当ペース", f"{ONK_TTM:,}円", "直近1年・税引後", BLUE),
]
kw = (CW - 30) / 4
kt = y0 + 20
for i, (label, big, small, col) in enumerate(onk_kpis):
    x = ML + i * (kw + 10)
    box(x, kt, kw, 50, LGRAY)
    c.setFillColor(col)
    c.rect(x, Y(kt + 50), 3, 50, fill=1, stroke=0)
    text(x + 9, kt + 13, label, 7, "JP", SUB)
    text(x + 9, kt + 31, big, 12.5, "JPB", col)
    text(x + 9, kt + 44, small, 6.3, "JP", SUB)

rt = kt + 62
text(ML, rt, "■ 銘柄別 恩株達成度ランキング（上位12・夫=亨／妻=美香）", 10, "JPB", NAVY)
rt += 4
rmax = 100.0
barx = ML + 224
barw = CW - 224 - 50
rh = 19
text(barx, rt + 5, "0%", 6, "JP", GRAY)
text(barx + barw, rt + 5, "100%（恩株達成）", 6, "JP", GRAY, "r")
rt += 8
for i, (nm, owner, acct, avg, sh, cost, life, prog) in enumerate(ONK_RANK):
    yy = rt + i * rh
    text(ML, yy + 8, nm[:13], 8, "JP", TXT)
    text(barx - 6, yy + 8, f"{owner}·{acct}", 6, "JP", GRAY, "r")
    text(ML, yy + 16.5, f"取得 {avg:,}円/株×{sh:,}株={cost:,}円 ／ 配当累計 {life:,}円", 6, "JP", SUB)
    by = yy + 4
    c.setFillColor(MGRAY)
    c.rect(barx, Y(by + 10), barw, 10, fill=1, stroke=0)
    for k in range(1, 4):
        gx = barx + barw * k / 4
        c.setStrokeColor(WHITE)
        c.setLineWidth(0.7)
        c.line(gx, Y(by + 10), gx, Y(by))
    col = GOLD if prog >= 15 else TEAL
    c.setFillColor(col)
    c.rect(barx, Y(by + 10), barw * min(prog, rmax) / rmax, 10, fill=1, stroke=0)
    text(barx + barw + 4, yy + 9, f"{prog:.1f}%", 7.5, "JPB", col, "r")
mt = rt + len(ONK_RANK) * rh + 9
text(ML, mt, ONK_MORE, 6.5, "JP", SUB)

# 受取配当の推移（年別）
gt = mt + 16
text(ML, gt, "■ 受取配当(税引後)の推移 ― 増配で恩株が加速", 10, "JPB", NAVY)
text(W - MR, gt, f"生涯累計 {ONK_LIFE_TOTAL:,}円（国内＋米国・税引後）", 6.5, "JP", SUB, "r")
gt += 8
gymax = max(v for _, v in ONK_BY_YEAR)
slot = CW / len(ONK_BY_YEAR)
bw = 30
gh = 52
for i, (yk, v) in enumerate(ONK_BY_YEAR):
    cxb = ML + slot * (i + 0.5)
    hh = gh * v / gymax
    c.setFillColor(TEAL)
    c.rect(cxb - bw / 2, Y(gt + gh), bw, hh, fill=1, stroke=0)
    text(cxb, gt + gh - hh - 3, _man(v), 7, "JPB", NAVY, "c")
    text(cxb, gt + gh + 10, str(yk), 6.5, "JP", SUB, "c")

# 口座別サマリー（3カラム）
at = gt + gh + 24
text(ML, at, "■ 口座別の恩株", 10, "JPB", NAVY)
at += 6
aw = (CW - 20) / 3
for i, (lab, cost, pct, ttm) in enumerate(ONK_ACCT):
    x = ML + i * (aw + 10)
    box(x, at, aw, 50, LGRAY)
    c.setFillColor(TEAL if i == 0 else (BLUE if i == 1 else PURPLE))
    c.rect(x, Y(at + 50), 3, 50, fill=1, stroke=0)
    text(x + 9, at + 13, lab, 7.5, "JPB", DARK)
    text(x + 9, at + 30, f"恩株 {pct:.1f}%", 12, "JPB", NAVY)
    text(x + 9, at + 43, f"取得 {cost:,}円・年配当 {ttm:,}円", 6.3, "JP", SUB)

# ハイライト
ht0 = at + 62
box(ML, ht0, CW, 40, TEALBG)
c.setFillColor(TEAL)
c.rect(ML, Y(ht0 + 40), 3, 40, fill=1, stroke=0)
text(ML + 12, ht0 + 15, f"ハイライト：恩株達成が近いのは {ONK_NEAR}。年間 約{ONK_TTM:,}円(税引後)のペースで回収中。", 7.6, "JPB", DARK)
text(ML + 12, ht0 + 30, f"高配当株の取得元本 {_man(ONK_COST)}円のうち {ONK_PCT:.1f}%（{ONK_DIV:,}円）を配当だけで回収済み。増配が続けば残り約{ONK_REM_Y}年で全額回収＝恩株化。", 7.6, "JP", TXT)

# 注記
nt = ht0 + 50
box(ML, nt, CW, 40, AMBERBG)
c.setFillColor(AMBER)
c.rect(ML, Y(nt + 40), 3, 40, fill=1, stroke=0)
text(ML + 12, nt + 14, "※対象：日本株（夫婦合算）＋SPYD（米国高配当ETF・312株）。台帳P1「高配当株 9,361,519円」＝日本株＋SPYDの範囲とほぼ整合。", 6.8, "JP", TXT)
text(ML + 12, nt + 26, "　SPYD：取得1,672,068円／生涯配当210,933円（税引後）＝恩株12.6%。SPCX・ボーイングは無配の成長株のため恩株には含めません。", 6.8, "JPB", AMBER)

text(W / 2, 826, "－ 3 / 7 －", 8, "JP", GRAY, "c")
c.showPage()

# ============================================================ Page 3 不動産投資と老後の設計
c.setFillColor(NAVY)
c.rect(0, Y(44), W, 44, fill=1, stroke=0)
text(ML, 28, "不動産投資と老後の設計", 14, "JPB", WHITE)
text(W - MR, 28, "投資用ローン・サブリース・老後収入の柱", 8, "JP", HexColor("#C8D2E4"), "r")

# ---- ローンの状況（住宅＋投資用・すべて変動金利）
lt = 66
text(ML, lt, "■ ローンの状況（住宅＋投資用2戸・すべて変動金利／基準2026-06）", 10.5, "JPB", NAVY)
lcols = [ML, ML + 132, ML + 222, ML + 304, ML + 392, W - MR]
lhead = ["ローン（借入先）", "ローン残高", "毎月返済", "適用金利", "完済予定"]
lt += 10
c.setFillColor(NAVY)
c.rect(ML, Y(lt + 15), CW, 15, fill=1, stroke=0)
text(lcols[0] + 4, lt + 11, lhead[0], 8, "JPB", WHITE)
for i in range(1, 5):
    text(lcols[i + 1] - 4, lt + 11, lhead[i], 8, "JPB", WHITE, "r")
lrows = [
    ("人宿町 自宅（静岡銀行）",   f"{HOME_LOAN:,}円",     f"{HOME_LOAN_PMT:,}円", "変動 1.075%", HOME_LOAN_END, False),
    ("鵜の木 投資用（ｵﾘｯｸｽ000001）", f"{LOAN_UNOKI_ZAN:,}円", f"{LOAN_UNOKI_PMT:,}円", "変動 2.600%", "2058-09", False),
    ("西台 投資用（ｵﾘｯｸｽ000002）",   f"{LOAN_NISHI_ZAN:,}円", f"{LOAN_NISHI_PMT:,}円", "変動 2.600%", "2058-09", False),
    ("合計（住宅＋投資用）",       f"{HOME_LOAN + RE_LOAN_TOTAL:,}円", f"{HOME_LOAN_PMT + RE_PMT_TOTAL:,}円", "全て変動", "—", True),
]
lt += 15
for name, zan, pmt, rate, end, is_total in lrows:
    c.setFillColor(LGRAY if is_total else WHITE)
    c.rect(ML, Y(lt + 16), CW, 16, fill=1, stroke=0)
    fnt = "JPB" if is_total else "JP"
    text(lcols[0] + 4, lt + 12, name, 8, fnt, TXT)
    for i, v in enumerate([zan, pmt, rate, end]):
        text(lcols[i + 2] - 4, lt + 12, v, 8, fnt, TXT, "r")
    lt += 16

# 金利推移＋キャッシュフロー注記（左右2カラム）
lt += 10
colw = (CW - 12) / 2
box(ML, lt, colw, 56, AMBERBG)
text(ML + 10, lt + 14, "3本とも変動金利で上昇中", 8.5, "JPB", AMBER)
text(ML + 10, lt + 29, "住宅 0.425→0.575→0.825→1.075% ／ 投資用 2.25→2.50→2.60%", 7.0, "JP", TXT)
text(ML + 10, lt + 41, "→ 投資用は次回2.85%予定。投資用+1%で年約46万円増。", 7.0, "JPB", RED)
bx3 = ML + colw + 12
box(bx3, lt, colw, 56, LGRAY)
text(bx3 + 10, lt + 14, "サブリースで空室・滞納リスクなし", 8.5, "JPB", TEAL)
text(bx3 + 10, lt + 29, f"保証家賃 {RENT_Y:,}円/年（鵜の木86,000・西台76,050/月）", 7.0, "JP", TXT)
text(bx3 + 10, lt + 41, "西台は2026年9月が更新期限→保証賃料の改定（減額）に注意。", 7.0, "JP", SUB)

# ---- 老後の収入設計
ot = lt + 80
text(ML, ot, "■ 老後の収入の柱（65〜85歳の想定・現在価値）", 10.5, "JPB", NAVY)
ot += 14
pillars = [
    ("公的年金", 250, "夫婦合算（想定）", BLUE),
    ("不動産家賃", 195, "投資用2戸・サブリース保証（鵜86,000+西76,050）", TEAL),
    ("高配当株 配当", 280, "65歳・約7,000万円想定", GOLD),
    ("美香さん私的年金", 120, "ソニー変額年金・70〜85歳（確定15年）", PURPLE),
    ("亨さん私的年金", 87, "ニッセイ外貨建・65〜75歳（確定10年・試算）", ROSE),
]
pmax = 300
pbar_x = ML + 110
pbar_w = CW - 110 - 60
for i, (label, amt, note, col) in enumerate(pillars):
    y = ot + i * 26
    text(ML, y + 11, label, 8, "JPB", TXT)
    c.setFillColor(MGRAY)
    c.rect(pbar_x, Y(y + 14), pbar_w, 12, fill=1, stroke=0)
    c.setFillColor(col)
    c.rect(pbar_x, Y(y + 14), pbar_w * amt / pmax, 12, fill=1, stroke=0)
    text(pbar_x + pbar_w * amt / pmax + 4, y + 11, f"{amt}万円/年", 7.5, "JPB", col)
    text(ML, y + 20, note, 6.5, "JP", SUB)

# DC取り崩しの補足（老後の柱と別枠）＋夫婦の私的年金ボックス
mt = ot + 5 * 26 + 6
box(ML, mt, CW, 30, LGRAY)
text(ML + 12, mt + 13, "＋ 企業型DC：65歳で約1.48億円（年7%・P2）。一時金または年金で受け取り、上記の柱に上乗せされる老後資産の主軸。",
     7.6, "JP", DARK)
text(ML + 12, mt + 25, "（公的年金・家賃・配当・私的年金の年間フローに加え、DCは取り崩し原資として別建てで確保できている）",
     7, "JP", SUB)

mt2 = mt + 38
box(ML, mt2, CW, 64, TEALBG)
c.setFillColor(PURPLE)
c.rect(ML, Y(mt2 + 64), 3, 64, fill=1, stroke=0)
text(ML + 12, mt2 + 15, "夫婦の私的年金（個人年金保険・NISA/DCとは別建ての老後の柱）", 8.5, "JPB", PURPLE)
text(ML + 12, mt2 + 30,
     f"亨：ニッセイ・ウェルス生命 外貨建個人年金（静岡銀行窓販）月{TORU_ANN_PREM_M:,}円・累計{TORU_ANN_PAID:,}円 → 65歳から確定年金10年（試算 年約87万円）。",
     7.2, "JP", TXT)
text(ML + 12, mt2 + 43,
     f"美香：ソニー生命 変額個人年金2本 月{MIKA_PREMIUM_M:,}円 → 70歳から15年 年{MIKA_ANNUITY_Y:,}円（合計1,800万円目安）。死亡保障{MIKA_DEATH:,}円。",
     7.2, "JP", TXT)
text(ML + 12, mt2 + 56,
     "※亨の外貨建は米ドル建で為替リスクあり＋途中解約は市場価格調整で大幅減（解約返戻2,625ドル/積立6,360ドル）。満期保有が前提。",
     7.2, "JP", SUB)

# ---- A-1 変動金利の推移ステップチャート（3本すべて変動金利で上昇中）
rl_sect = mt2 + 64 + 18
text(ML, rl_sect, "■ 変動金利の推移（3本すべて変動金利・実際に上昇中）", 9.5, "JPB", NAVY)
rl_ch = 72           # チャート高さ（px）
rl_top = rl_sect + 16
rl_lx = ML + 40      # y軸ラベル幅を確保
rl_rx = W - MR - 10
rl_w = rl_rx - rl_lx
rl_ymin, rl_ymax = 0.0, 3.5
rl_xlabels = ["2022\n借入", "2024", "2025/1", "2025/7", "2026/7\n新利率"]
rl_n = len(rl_xlabels)
rl_xs = [rl_lx + rl_w * i / (rl_n - 1) for i in range(rl_n)]


def rl_y(r):
    """金利 r (%) → reportlab Y 座標（パス・サークル用）"""
    return Y(rl_top + rl_ch * (1 - (r - rl_ymin) / (rl_ymax - rl_ymin)))


def rl_tc(r):
    """金利 r (%) → t 座標（text() 第2引数用）"""
    return rl_top + rl_ch * (1 - (r - rl_ymin) / (rl_ymax - rl_ymin))


# y軸グリッドとラベル
for gv in [0.0, 1.0, 2.0, 3.0]:
    gy = rl_top + rl_ch * (1 - (gv - rl_ymin) / (rl_ymax - rl_ymin))
    c.setStrokeColor(MGRAY)
    c.setLineWidth(0.4)
    c.line(rl_lx, Y(gy), rl_rx, Y(gy))
    text(rl_lx - 4, gy + 2, f"{gv:.0f}%", 6, "JP", GRAY, "r")

# x軸ラベル（2行対応）
for i, lbl in enumerate(rl_xlabels):
    parts = lbl.split("\n")
    for j, part in enumerate(parts):
        text(rl_xs[i], rl_top + rl_ch + 8 + j * 9, part, 6, "JP", GRAY, "c")

# 住宅ローン（BLUE）: 0.425% → 0.575% → 0.825% → 1.075%（2026/07〜）
h_rates = [0.425, 0.425, 0.575, 0.825, 1.075]
ph = c.beginPath()
ph.moveTo(rl_xs[0], rl_y(h_rates[0]))
for i in range(1, rl_n):
    ph.lineTo(rl_xs[i], rl_y(h_rates[i - 1]))
    ph.lineTo(rl_xs[i], rl_y(h_rates[i]))
c.setStrokeColor(BLUE)
c.setLineWidth(1.8)
c.drawPath(ph, fill=0, stroke=1)

# 投資用ローン（RED）: 2.25% → 2.50% → 2.60%
v_rates = [2.25, 2.50, 2.50, 2.60, 2.60]
pv = c.beginPath()
pv.moveTo(rl_xs[0], rl_y(v_rates[0]))
for i in range(1, rl_n):
    pv.lineTo(rl_xs[i], rl_y(v_rates[i - 1]))
    pv.lineTo(rl_xs[i], rl_y(v_rates[i]))
c.setStrokeColor(RED)
c.setLineWidth(1.8)
c.drawPath(pv, fill=0, stroke=1)

# 最終値にドット＋値ラベル
c.setFillColor(BLUE)
c.circle(rl_xs[-1], rl_y(h_rates[-1]), 2.5, fill=1, stroke=0)
text(rl_xs[-1] - 4, rl_tc(h_rates[-1]) - 6, "1.075%", 6.5, "JPB", BLUE, "r")
c.setFillColor(RED)
c.circle(rl_xs[-1], rl_y(v_rates[-1]), 2.5, fill=1, stroke=0)
text(rl_xs[-1] - 4, rl_tc(v_rates[-1]) - 6, "2.60%", 6.5, "JPB", RED, "r")

# 凡例
rl_lg = rl_top + rl_ch + 28
c.setStrokeColor(BLUE)
c.setLineWidth(1.5)
c.line(rl_lx, Y(rl_lg), rl_lx + 22, Y(rl_lg))
text(rl_lx + 26, rl_lg + 4, "住宅ローン（静岡銀行・0.425%→1.075%）", 7, "JP", TXT)
c.setStrokeColor(RED)
c.line(rl_lx + 190, Y(rl_lg), rl_lx + 212, Y(rl_lg))
text(rl_lx + 216, rl_lg + 4, "投資用ローン（オリックス2戸・2.25%→2.60%）", 7, "JP", TXT)
text(W - MR, rl_lg + 16, "※投資用は次回金利更改で2.85%へ上昇予定（+1%で年約46万円増）", 7, "JPB", RED, "r")

text(W / 2, 826, "－ 4 / 7 －", 8, "JP", GRAY, "c")
c.showPage()

# ============================================================ Page 4 強み・弱点・アクション
c.setFillColor(NAVY)
c.rect(0, Y(44), W, 44, fill=1, stroke=0)
text(ML, 28, "分析サマリー：強み・弱点・アクション", 14, "JPB", WHITE)
text(W - MR, 28, "田中家 資産管理レポート 2026-07-07", 8, "JP", HexColor("#C8D2E4"), "r")

# ---- 強み
text(ML, 66, "■ 強み", 10.5, "JPB", TEAL)
strengths = [
    "40歳で純資産約4,917万円。夫婦の企業型DCはS&P500へ月11万円拠出（現在1,187万円）で資産形成エンジンが強力。",
    "住宅ローン（静岡銀行・残高5,636万）と亨の外貨建個人年金が確定。投資用2戸も含み益プラス（合計＋876万円）。",
    "夫婦のNISAは合計424万円・含み益+105万円（亨+20%／美香+51%）。生涯枠3,600万円の11.8%を消化し運用好調。",
    "金融資産は10年で約253万円→約3,944万円（約15.6倍）。入金力と継続力が数字で実証されている。",
    "老後の柱は公的年金250＋家賃195＋配当280＋私的年金（美香120・亨87）万円と多層で、DC約1.48億も別途控える。",
]
st = 76
box(ML, st, CW, 16 * len(strengths) + 12, TEALBG)
yy = st + 17
for s in strengths:
    text(ML + 12, yy, "・" + s, 8.5, "JP", TXT)
    yy += 16

# ---- 弱点・リスク表
wt = st + 16 * len(strengths) + 28
text(ML, wt, "■ 弱点・リスク（優先度順）", 10.5, "JPB", RED)
wt += 10

wcol_p, wcol_t, wcol_d = 34, 124, CW - 34 - 124
weaknesses = [
    ("高", "変動金利の上昇が現実化（3本とも変動）",
     "投資用2戸は2.25→2.60%（次回2.85%予定）、住宅ローンも0.425→1.075%へ実際に上昇（2026-07適用）。3本すべて変動金利で、投資用は残高4,568万に+1%で年約46万円増。家賃はサブリースで安定だが返済増は手残りを直撃する。固定化や繰上返済を検討。"),
    ("中", "DC予測は年8%だと強気（本試算は年7%）",
     "台帳の「65歳で1.7億」は年8%前提。本レポートは保守的に年7%で見て約1.48億（P2）。年5%なら約1.0億まで下がる。S&P500一本のため、下振れ局面に備えて取り崩し開始前の数年はリスクを落とす出口戦略も検討を。"),
    ("中", "投資用2戸の含み益が薄い",
     "LTVは2戸とも約95%、含み益は西台+109万・鵜の木+123万のみ。評価額は取得額ベースで、中古区分は実勢が下振れしやすい。売却を考えるなら査定を取り、出口時期と想定価格を点検。"),
    ("中", "米国株・ドルへの集中",
     "DC（S&P500）・投信（オルカン/S&P500/VTI）・SPYDと、リスク資産の大半が米国株。円高と米株調整が同時に来ると金融資産の大部分が直撃される。"),
    ("中", "SPYDをNISAで保有",
     "NISAでは外国税額控除が使えず、米国源泉徴収10%は取り戻せない。NISA枠は日本高配当株を優先し、米国ETFは特定口座の方が税効率が良い。"),
    ("中", "成長枠の消化ペースが緩い",
     "方針は成長枠（夫婦2,400万円）を優先だが、消化は夫婦で315万＝13%のみ。亨さんは年240万に近いが美香さんは0。美香さんが月10万円なら10年、年240万（上限）なら5年で1,200万円を満了。つみたて枠の増額は成長枠を埋めた後でよい。"),
    ("中", "教育資金の指定席がない",
     "湊人さんは約10年後に大学進学（私大4年で500〜800万円規模）。生活防衛資金とも投資とも別に、教育費の目標額と置き場所を決めておく。"),
    ("中", "美香さんのNISAが手薄",
     "成長投資枠1,200万円が丸ごと未使用で、つみたて月1万円のみ＝生涯枠の消化は2.9%。含み益率+51%と運用は好調。方針どおり成長枠を優先するなら、夫婦で最優先に入金すべきはここで、非課税の恩恵が最も大きい。"),
    ("低", "亨の外貨建個人年金は為替・解約に注意",
     "亨のニッセイ・ウェルス生命 外貨建個人年金（静岡銀行窓販・残高はMF保険列170万）は米ドル建で円高だと円受取が目減り。金利上昇で市場価格調整が大きく、今解約すると返戻は積立の約4割（2,625ドル/6,360ドル）。満期保有が前提で、個人年金保険料控除は活用を。住宅ローンの基準日も2026年6月に更新。"),
]
pcol = {"高": RED, "中": AMBER, "低": GRAY}

c.setFillColor(NAVY)
c.rect(ML, Y(wt + 14), CW, 14, fill=1, stroke=0)
text(ML + 6, wt + 10.5, "優先", 7.5, "JPB", WHITE)
text(ML + wcol_p + 6, wt + 10.5, "弱点", 7.5, "JPB", WHITE)
text(ML + wcol_p + wcol_t + 6, wt + 10.5, "内容と推奨アクション", 7.5, "JPB", WHITE)
wt += 14

for i, (pri, title, desc) in enumerate(weaknesses):
    tl = wrap(title, "JPB", 8, wcol_t - 10)
    dl = wrap(desc, "JP", 7.5, wcol_d - 12)
    rh = max(len(tl) * 10.5, len(dl) * 10) + 8
    if i % 2 == 0:
        c.setFillColor(LGRAY)
        c.rect(ML, Y(wt + rh), CW, rh, fill=1, stroke=0)
    c.setFillColor(pcol[pri])
    c.roundRect(ML + 5, Y(wt + 16), 22, 12, 3, fill=1, stroke=0)
    text(ML + 16, wt + 13, pri, 7.5, "JPB", WHITE, "c")
    yy = wt + 13
    for ln in tl:
        text(ML + wcol_p + 6, yy, ln, 8, "JPB", TXT)
        yy += 10.5
    yy = wt + 12.5
    for ln in dl:
        text(ML + wcol_p + wcol_t + 6, yy, ln, 7.5, "JP", TXT)
        yy += 10
    wt += rh

c.setStrokeColor(MGRAY)
c.line(ML, Y(wt), ML + CW, Y(wt))

# ---- アクションリスト
wt += 16
text(ML, wt, "■ 今後のアクション", 10.5, "JPB", NAVY)
wt += 6
actions = [
    ("今月", "住宅1.075%・投資用2.6%と3本とも変動金利が上昇中。固定化／繰上返済の効果を試算する"),
    ("今月", "美香さんのNISA成長枠（1,200万円・未使用）への入金を最優先で開始する（つみたて増額より先に）"),
    ("今月", "住宅ローン（静岡銀行）の団信の保障内容を確認する（残高・金利・残期間・基準日は確定済み）"),
    ("年内", "亨の外貨建個人年金は満期保有を前提に出口（為替）を計画。個人年金保険料控除を年末調整で適用する"),
    ("年内", "亨さんは成長枠を年240万円ペースで満了（残り885万円）。つみたて増額は成長枠を埋めた後に回す"),
    ("年内", "NISA内の米国高配当ETF（SPYD）→日本高配当株または投信への入替を検討"),
    ("年内", "教育費の目標額（例：2036年までに600万円）を設定し、資金の置き場所を決める"),
    ("年内", "DC・将来予測を年7%を基準に（下振れ年5%も）見直し、65歳の出口戦略を具体化する"),
]
for due, act in actions:
    box(ML, wt, 30, 11, LGRAY, r=2)
    text(ML + 15, wt + 8.5, due, 7, "JPB", SUB, "c")
    text(ML + 38, wt + 8.5, "□ " + act, 8.5, "JP", TXT)
    wt += 15

# ---- 定期アップロード資料チェックリスト
# 「金融資産の歩み」チャートはP2（■企業型DCの積み上げ の下の余白）へ移設・拡大済み（余白平準化・2026-07-06 v25）
wt += 22
text(ML, wt, "■ 定期アップロード資料チェックリスト（毎回このフォルダに入れる資料）", 10, "JPB", NAVY)
checklist = [
    ("資産推移 月次CSV（MF）",              "毎月",            False),
    ("NISA状況 亨・美香（SBI証券）",         "四半期",           False),
    ("DC資産状況 亨・美香（SBIベネフィット）", "四半期",           False),
    ("負債一覧（MF）",                     "四半期",           False),
    ("資産管理台帳／基礎データ",              "更新都度",          False),
    ("不動産ローン契約内容 2戸（オリックス）",   "半年・金利更改時",   False),
    ("サブリース契約書 2戸（鵜の木・西台）",     "年1回",           False),
    ("変額年金 美香2本（ソニー生命）",         "年1回",           False),
    ("住宅ローン明細表（人宿町・静岡銀行）",     "年1回",           False),
    ("亨 外貨建年金（ﾆｯｾｲ･ｳｪﾙｽ/静岡銀行）",    "年1回",           False),
]
clw = (CW - 12) / 2
fcol = {"毎月": BLUE, "四半期": TEAL, "半年・金利更改時": DARK, "更新都度": SUB, "年1回": GRAY}
for i, (name, freq, pending) in enumerate(checklist):
    col = 0 if i < 5 else 1
    row = i if i < 5 else i - 5
    x = ML + col * (clw + 12)
    y = wt + 15 + row * 12
    c.setStrokeColor(GRAY)
    c.setLineWidth(0.6)
    c.rect(x, Y(y + 6.5), 8, 8, fill=0, stroke=1)
    text(x + 13, y + 7, name, 7.3, "JP", AMBER if pending else TXT)
    text(x + clw, y + 7, freq, 6.5, "JP", AMBER if pending else fcol.get(freq, SUB), "r")
wt += 15 + 5 * 12

text(ML, 813, "※本レポートは台帳・各社照会の数値に基づく概算・参考情報であり、特定の金融商品の売買を推奨するものではありません。", 6.5, "JP", GRAY)
text(W / 2, 826, "－ 5 / 7 －", 8, "JP", GRAY, "c")
c.showPage()

# ============================================================ Page 5 保険・保障の総点検
c.setFillColor(NAVY)
c.rect(0, Y(44), W, 44, fill=1, stroke=0)
text(ML, 28, "保険・保障の総点検", 14, "JPB", WHITE)
text(W - MR, 28, "医療・死亡・就労不能・介護・老後をまとめて点検", 8, "JP", HexColor("#C8D2E4"), "r")

mark_col = {"○": TEAL, "△": AMBER, "×": RED}

# ---- A) 世帯の保障マップ
text(ML, 64, "■ 世帯の保障マップ（誰が・どのリスクに・何で備えているか）", 10.5, "JPB", NAVY)
mc = [ML, ML + 104, ML + 104 + (CW - 104) // 2]      # リスク / 亨 / 美香
ah = 78
c.setFillColor(NAVY)
c.rect(ML, Y(ah + 15), CW, 15, fill=1, stroke=0)
text(mc[0] + 6, ah + 11, "リスク", 8, "JPB", WHITE)
text(mc[1] + 6, ah + 11, "亨さん（本人40歳）", 8, "JPB", WHITE)
text(mc[2] + 6, ah + 11, "美香さん（39歳）", 8, "JPB", WHITE)
ah += 15
maprows = [
    ("医療（入院・手術）", "○", "メットライフ 日額1.4万",  "○", "メットライフ 日額1.3万"),
    ("先進医療",          "○", "特約（技術料を実費補償）",   "○", "特約（技術料を実費補償）"),
    ("死亡（遺族保障）",   "○", "団信で自宅＋投資用ローン完済",  "△", "死亡保障 約183万のみ＝薄い"),
    ("就労不能・働けない", "×", "傷病手当金1.5年の先が穴",     "△", "扶養内＝影響は小さい"),
    ("介護",              "△", "医療保険に介護特約は未付加",  "△", "医療保険に介護特約は未付加"),
    ("老後の私的年金",     "○", "ニッセイ外貨建 65歳〜10年",  "○", "ソニー変額 70歳〜15年"),
]
for i, (risk, m1, t1, m2, t2) in enumerate(maprows):
    rh = 22
    c.setFillColor(LGRAY if i % 2 == 0 else WHITE)
    c.rect(ML, Y(ah + rh), CW, rh, fill=1, stroke=0)
    text(mc[0] + 6, ah + 14, risk, 7.6, "JPB", TXT)
    text(mc[1] + 6, ah + 14, m1, 9.5, "JPB", mark_col[m1])
    text(mc[1] + 20, ah + 14, t1, 7.2, "JP", TXT)
    text(mc[2] + 6, ah + 14, m2, 9.5, "JPB", mark_col[m2])
    text(mc[2] + 20, ah + 14, t2, 7.2, "JP", TXT)
    ah += rh
text(ML, ah + 13, "○=備えあり ／ △=一部・要確認 ／ ×=未対応。医療は今回のメットライフ2契約で夫婦とも空白が埋まった。", 7, "JP", SUB)

# ---- B) 加入している保険の一覧
bt = ah + 32
text(ML, bt, "■ 加入している保険の一覧（保障の中身と保険料）", 10.5, "JPB", NAVY)
bc = [ML, ML + 150, ML + 220, ML + 286, ML + 374, W - MR]
bh = bt + 10
c.setFillColor(NAVY)
c.rect(ML, Y(bh + 15), CW, 15, fill=1, stroke=0)
for j, htxt in enumerate(["商品・引受先", "種類", "契約者", "保険料", "受取・備考"]):
    text(bc[j] + 5, bh + 11, htxt, 7.6, "JPB", WHITE)
bh += 15
inslist = [
    ("メットライフ マイフレキシィ（亨）",  "医療・終身", "法人※", "月24,488円",     "入院日額1.4万＋手術・先進医療"),
    ("メットライフ マイフレキシィ（美香）", "医療・終身", "法人※", "月24,883円",     "入院日額1.3万＋手術・先進医療"),
    ("ソニー生命 変額個人年金2本（美香）",  "個人年金",   "個人",   "月20,244円",     "70歳〜15年・年120万／死亡183万"),
    ("ニッセイ・ウェルス 外貨建年金（亨）",  "個人年金",   "個人",   "月20,000円",     "65歳〜10年・年約87万(米ドル建)"),
    ("住宅ローン 団体信用生命（亨・連帯）",  "死亡保障",   "—",     "金利に内包",     "死亡時に住宅ローン残債を完済"),
]
for i, (prod, kind, holder, prem, note) in enumerate(inslist):
    rh = 26
    c.setFillColor(LGRAY if i % 2 == 0 else WHITE)
    c.rect(ML, Y(bh + rh), CW, rh, fill=1, stroke=0)
    text(bc[0] + 5, bh + 15, prod, 7.2, "JP", TXT)
    text(bc[1] + 5, bh + 15, kind, 7.2, "JP", TXT)
    is_corp = holder == "法人※"
    text(bc[2] + 5, bh + 15, holder, 7.2, "JPB" if is_corp else "JP", AMBER if is_corp else TXT)
    text(bc[3] + 5, bh + 15, prem, 7.2, "JP", TXT)
    text(bc[4] + 5, bh + 15, note, 7.0, "JP", TXT)
    bh += rh
text(ML, bh + 13,
     "※印＝株式会社しずく（法人）契約。10年の払込完了後に名義を法人→個人へ変更予定（下記）。夫婦とも設計書で保険料・保障を確定済み。",
     7, "JP", SUB)

# ---- C) 名義変更スキームの注記
ct = bh + 28
box(ML, ct, CW, 56, AMBERBG)
c.setFillColor(AMBER)
c.rect(ML, Y(ct + 56), 3, 56, fill=1, stroke=0)
text(ML + 12, ct + 16, "法人契約の医療保険2本は「10年払込 → 名義を法人から個人へ変更」する設計", 8.5, "JPB", AMBER)
text(ML + 12, ct + 31,
     "払込期間中は会社が保険料を負担し、完了後は本人の終身医療保障として個人に残る。掛け捨て型で資産価値ゼロ＝家計のBS・資産推移には計上しない。",
     7.2, "JP", TXT)
text(ML + 12, ct + 46,
     "※名義変更時の評価額・課税（給与/賞与扱い等）や損金処理は税制改正の影響あり。実行前に顧問税理士へ必ず確認。",
     7.2, "JPB", RED)

# ---- D) 保険からみた弱点・残る穴
dt = ct + 56 + 18
text(ML, dt, "■ 保険からみた弱点・残る穴（投資・資産の弱点はP4を参照）", 10.5, "JPB", RED)
dt += 10
pill_col = {"解消": TEAL, "高": RED, "中": AMBER, "低": GRAY}
insweak = [
    ("解消", "医療の空白＝メットライフ2契約で手当て済み。亨さんの死亡も団信（自宅＋投資用）で全ローン完済＝手厚い"),
    ("高",  "亨さん（主稼得）の長期就労不能が最大の穴。傷病手当金1.5年の先を就業不能保険（月20〜30万・65歳まで）で"),
    ("中",  "美香さん（扶養内）の死亡保障が約183万と薄い。家事・育児の代替費用見合いで500〜1,000万を収入保障で検討"),
    ("低",  "介護保障は夫婦とも医療保険に特約未付加。公的介護保険＋貯蓄で対応し必要なら介護特約を検討"),
    ("低",  "法人契約2本の損金処理・名義変更時の課税を顧問税理士に確認"),
]
for pri, desc in insweak:
    c.setFillColor(pill_col[pri])
    c.roundRect(ML, Y(dt + 13), 26, 11, 3, fill=1, stroke=0)
    text(ML + 13, dt + 11, pri, 7, "JPB", WHITE, "c")
    text(ML + 34, dt + 11, desc, 8, "JP", TXT)
    dt += 17

# ---- E) 保険まわりのアクション
et = dt + 14
text(ML, et, "■ 保険まわりのアクション", 10.5, "JPB", NAVY)
et += 7
insact = [
    ("今月", "亨さん（主稼得）の就業不能保険の見積りを依頼（月20〜30万・65歳まで・免責は傷病手当金1.5年に接続）"),
    ("年内", "亨さん（主稼得）に就業不能保険を検討（月20〜30万・65歳まで・免責を傷病手当金1.5年に接続）"),
    ("年内", "美香さんの死亡保障を収入保障で500〜1,000万円上乗せ検討＋法人契約の税務を税理士確認"),
    ("随時", "保険は年1回見直し。設計書・証券をこのフォルダに保管（亨／美香のメットライフ・ソニー・ニッセイ）"),
]
for due, act in insact:
    box(ML, et, 30, 11, LGRAY, r=2)
    text(ML + 15, et + 8.5, due, 7, "JPB", SUB, "c")
    text(ML + 38, et + 8.5, "□ " + act, 8.5, "JP", TXT)
    et += 16

# ---- F) 保障の総括
ft = et + 14
box(ML, ft, CW, 84, TEALBG)
c.setFillColor(TEAL)
c.rect(ML, Y(ft + 84), 3, 84, fill=1, stroke=0)
text(ML + 12, ft + 16, "保障の総括", 9.5, "JPB", DARK)
summ5 = ("これまで資産形成（NISA・DC・不動産）に偏っていた本レポートに、今回はじめて保険・保障の軸を加えた。"
         "医療は夫婦ともメットライフ（法人契約・10年後に個人へ名義変更）で空白が埋まり、亨さんの死亡も"
         "団信で自宅＋投資用の全ローンが完済される＝遺族保障は手厚く、追加の死亡保険は不要。"
         "最優先の穴は亨さん（主稼得）の長期就労不能で、傷病手当金1.5年の先を就業不能保険（月20〜30万・65歳まで）で。"
         "美香さんの遺族保障の薄さは家事育児コスト見合いで小さめに補えばよい。")
yy = ft + 31
for ln in wrap(summ5, "JP", 8, CW - 24):
    text(ML + 12, yy, ln, 8, "JP", TXT)
    yy += 13

text(ML, 813, "※医療保険2本は掛け捨て（無解約返戻金型）・法人負担のため、P1〜P4の資産・負債・純資産の数値には影響しません。", 6.5, "JP", GRAY)
text(W / 2, 826, "－ 6 / 7 －", 8, "JP", GRAY, "c")

c.showPage()

# ============================================================ Page 6 世帯年収（実収入）の推移
c.setFillColor(NAVY)
c.rect(0, Y(44), W, 44, fill=1, stroke=0)
text(ML, 28, "世帯年収（実収入）の推移", 14, "JPB", WHITE)
text(W - MR, 28, "給与＋非課税の出張日当＋DPC＋配偶者収入＋企業型DC", 8, "JP", HexColor("#C8D2E4"), "r")


def _man(v):
    s = f"{v:,.1f}"
    return s[:-2] if s.endswith(".0") else s


if not income_years:
    text(ML, 90, "※ 旅費_世帯年収まとめ.xlsx を参照できないか、年収データが未入力のため表示できません。",
         9.5, "JP", RED)
    text(ML, 108, f"参照先：{INCOME_XLSX}", 7, "JP", SUB)
else:
    it0 = 66
    text(ML, it0, "■ 世帯の実年収 ― 額面給与だけでなく、非課税の出張日当・DPC・企業型DCまで含めた実収入ベース",
         10.0, "JPB", NAVY)
    text(ML, it0 + 14,
         "出典：02_しずく＼05_出張旅費＼旅費_世帯年収まとめ.xlsx を自動参照。出張日当は毎月の精算ファイルから自動転記され、この推移も自動で更新されます。",
         7.0, "JP", SUB)

    SEGS = [("給与（亨）", "toru", NAVY),
            ("出張日当（非課税）", "nittou", GOLD),
            ("DPC", "dpc", TEAL),
            ("給与（美香）", "mika", ROSE),
            ("企業型DC（夫婦）", "dc", PURPLE)]

    # ---- 積み上げ棒グラフ（年ごと）----
    gtop, gh = it0 + 34, 150
    gx0, gx1 = ML + 44, ML + 300
    ymax = 400
    _mx = max(d["total"] for d in income_years)
    while ymax < _mx:
        ymax += 400
    for gv in range(0, ymax + 1, 400):
        yl = gtop + gh - gh * gv / ymax
        c.setStrokeColor(MGRAY)
        c.setLineWidth(0.4)
        c.line(gx0, Y(yl), gx1, Y(yl))
        text(gx0 - 4, yl + 2, ("0" if gv == 0 else f"{gv:,}万"), 6, "JP", GRAY, "r")
    nb = len(income_years)
    slot = (gx1 - gx0) / nb
    bw = min(50, slot * 0.5)
    for i, d in enumerate(income_years):
        cxb = gx0 + slot * (i + 0.5)
        acc = 0.0
        for (lbl, key, col) in SEGS:
            val = d[key]
            if val <= 0:
                continue
            c.setFillColor(col)
            c.rect(cxb - bw / 2, Y(gtop + gh - gh * acc / ymax),
                   bw, gh * val / ymax, fill=1, stroke=0)
            acc += val
        topy = gtop + gh - gh * d["total"] / ymax
        text(cxb, topy - 4, f"{_man(d['total'])}万", 8, "JPB", NAVY, "c")
        text(cxb, gtop + gh + 11, f"{d['year']}年", 8, "JPB", TXT, "c")
        if 0 < d["lastm"] < 12:
            text(cxb, gtop + gh + 20, f"{d['lastm']}月まで・進行中", 5.8, "JP", AMBER, "c")

    # ---- 凡例（右・最新年の内訳を併記）----
    lx, ly = gx1 + 24, gtop + 10
    latest = income_years[-1]
    for (lbl, key, col) in SEGS:
        c.setFillColor(col)
        c.rect(lx, Y(ly), 8, 8, fill=1, stroke=0)
        text(lx + 12, ly + 7, lbl, 6.8, "JP", TXT)
        text(W - MR, ly + 7, f"{_man(latest[key])}万", 6.8, "JP", SUB, "r")
        ly += 15
    c.setStrokeColor(MGRAY)
    c.setLineWidth(0.5)
    c.line(lx, Y(ly + 1), W - MR, Y(ly + 1))
    ly += 13
    text(lx, ly, f"{latest['year']}年 実年収", 7.2, "JPB", NAVY)
    text(W - MR, ly, f"{_man(latest['total'])}万", 8.5, "JPB", NAVY, "r")

    # ---- 内訳テーブル ----
    tt = gtop + gh + 44
    text(ML, tt, "■ 内訳（万円）", 10.0, "JPB", NAVY)
    colx = [ML + 6, ML + 128, ML + 204, ML + 272, ML + 352, ML + 440, W - MR - 6]
    heads = ["年", "給与(亨)", "日当(非課税)", "DPC", "給与(美香)", "企業型DC", "実年収"]
    hy = tt + 18
    box(ML, hy - 12, CW, 16, NAVY)
    for x, h in zip(colx, heads):
        text(x, hy, h, 7.5, "JPB", WHITE, "l" if h == "年" else "r")
    ry = hy + 18
    for d in income_years:
        vals = [f"{d['year']}", _man(d["toru"]), _man(d["nittou"]), _man(d["dpc"]),
                _man(d["mika"]), _man(d["dc"]), _man(d["total"])]
        for j, (x, v) in enumerate(zip(colx, vals)):
            text(x, ry, v, 8, "JPB" if j in (0, 6) else "JP",
                 NAVY if j == 6 else TXT, "l" if j == 0 else "r")
        c.setStrokeColor(MGRAY)
        c.setLineWidth(0.3)
        c.line(ML, Y(ry + 5), W - MR, Y(ry + 5))
        ry += 18

    # ---- 前年比 ----
    if len(income_years) >= 2:
        a, b = income_years[-2], income_years[-1]
        diff = b["total"] - a["total"]
        sign = "＋" if diff >= 0 else "－"
        text(ML, ry + 8,
             f"前年比：{a['year']}年 {_man(a['total'])}万 → {b['year']}年 {_man(b['total'])}万"
             f"（{sign}{_man(abs(diff))}万）"
             + ("　※最新年は進行中のため今後増加します" if 0 < b["lastm"] < 12 else ""),
             8, "JPB", DARK)

    # ---- 注記ボックス ----
    nbx = ry + 22
    box(ML, nbx, CW, 76, TEALBG)
    c.setFillColor(TEAL)
    c.rect(ML, Y(nbx + 76), 3, 76, fill=1, stroke=0)
    text(ML + 12, nbx + 15, "この『実年収』の考え方", 9.5, "JPB", DARK)
    notes = [
        "・額面給与に、非課税の出張日当・DPC報酬・配偶者（美香）の給与・企業型DC（将来資産）まで合算した世帯の実収入ベース。",
        "・出張日当は所得税・住民税・社会保険料が非課税のため、同じ手取りでも額面給与より効率がよい“隠れ収入”。",
        "・2024年分の出張日当47.5万円は2025年に一括計上した経緯があり、旅費まとめの備考に記録済み（本表は2025年以降が対象）。",
        "・数値は毎月の出張精算→旅費_世帯年収まとめ.xlsx（自動転記）を参照。家計簿を更新するたびに、この推移も最新化されます。",
    ]
    ny = nbx + 30
    for n in notes:
        for ln in wrap(n, "JP", 7.4, CW - 24):
            text(ML + 12, ny, ln, 7.4, "JP", TXT)
            ny += 11

text(W / 2, 826, "－ 7 / 7 －", 8, "JP", GRAY, "c")

c.save()
print("OK:", OUT)


# ============================================================ 生成したPDFを自分のGmailへ自動送信
# レポートを生成するたびに、最新PDFを自分のGmailへ添付メールで送る。
# スマホはGmailの通知をタップするだけで最新版が開ける（Driveの階層を潜らなくてよい）。
# アプリパスワードはコードに直書きせず、環境変数 KAKEI_MAIL_PASS から読む。
# 未設定ならスキップ＝パスワード未登録でもレポート生成自体は必ず成功させる。
def send_report_mail(pdf_path):
    import smtplib
    import ssl
    from email.message import EmailMessage
    from datetime import datetime

    ADDR = "ph2144.tt.0609@gmail.com"           # 送信元・宛先とも本人（自分→自分）
    LOG_MD = os.path.join(BASE, "メール自動送信_仕組みと送信ログ.md")

    def _log(result):
        # 実行のたびに送信ログMDの表末尾へ1行追記（成功/スキップ/失敗）。失敗しても本処理は止めない。
        try:
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            with open(LOG_MD, "a", encoding="utf-8") as lf:
                lf.write(f"| {stamp} | {os.path.basename(pdf_path)} | {ADDR} | {result} |\n")
        except Exception as _e:
            print("（送信ログ追記に失敗:", _e, "）")

    app_pass = os.environ.get("KAKEI_MAIL_PASS")
    if not app_pass:
        print("メール送信スキップ：環境変数 KAKEI_MAIL_PASS が未設定です（PDFは生成済み）")
        _log("スキップ（KAKEI_MAIL_PASS未設定）")
        return
    app_pass = app_pass.replace(" ", "")        # アプリパスワードの空白は除去して結合

    msg = EmailMessage()
    msg["Subject"] = "【家計レポート】最新版 " + datetime.now().strftime("%Y%m%d")
    msg["From"] = ADDR
    msg["To"] = ADDR
    msg.set_content(
        "田中家の資産管理 分析レポート（最新版）を添付しました。\n"
        "\n"
        "■ 毎月の更新のしかた（迷ったらこちら）\n"
        "https://ph2144tt0609-cmyk.github.io/apps-portal/kakei-update.html\n"
        "↑ 何のデータを・どのフォルダに置けばいいか、ステップごとに説明しています。\n"
        "月に1回、ここを見れば更新のしかたを思い出せます。\n"
        "\n"
        "■ この仕組みについて\n"
        "・PCで家計レポートを作成・更新すると、最新のPDFがこのメールで自動的に届きます。\n"
        "・スマホでは、このメールの添付PDFを開くだけで常に最新版が見られます"
        "（Googleドライブの中を探す必要はありません）。\n"
        "・レポートの数値（資産・NISA・企業型DC・ローン・保険など）は手入力で更新しています。"
        "新しい資料が届いたら数値を反映し、再びこのメールで届きます。\n"
        "\n"
        "■ 安全性\n"
        "・このメールは ph2144.tt.0609@gmail.com から自分自身宛に自動送信しています。"
        "Googleの中で完結し、第三者には届きません。\n"
        "\n"
        "（PCの create_report.py が自動送信しています）"
    )
    with open(pdf_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf",
                           filename=os.path.basename(pdf_path))
    try:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
            s.login(ADDR, app_pass)
            s.send_message(msg)
    except Exception as e:
        print(f"メール送信失敗: {type(e).__name__}: {e}（PDFは生成済み）")
        _log(f"失敗（{type(e).__name__}）")
        return
    print(f"メール送信完了 → {ADDR}（{os.path.basename(pdf_path)}）")
    _log("成功")


send_report_mail(OUT)
